; ==============================================================================
; CMDT - Run as TrustedInstaller
; Command-Line / File-Run Dispatch Module
;
; Author: Marek Wesolowski (wesmar)
; Ported to: Native ARM64 (armasm64) for Windows on ARM
;
; Purpose: Hosts the two execution-style modes that actually run user
;          commands as TrustedInstaller — the explicit '-cli' flow
;          and the context-menu 'cmdt.exe ""' flow. Both end up calling
;          RunAsTrustedInstaller; this module handles command-line parsing,
;          optional .lnk resolution, and the internal -outfile relay protocol
;          that the non-admin path uses to capture output.
;
; Exported labels (EXPORT):
;   mode_cli_found  - Jumped to from mainCRTStartup when argv[1] == "-cli".
;                     Validates -new placement and the argv count.
;   mode_file_run   - Jumped to when argv[1] is not a recognized switch and
;                     not a help token, e.g. a file path passed by Explorer's
;                     context menu integration.
;
; ARM64 Port Notes:
;   - These labels share mainCRTStartup's stack frame (x29 = frame pointer).
;   - argv is passed in x20 (preserved from mainCRTStartup).
;   - argc is stored at [x29, #-64] (set by mainCRTStartup).
;   - SECURITY_ATTRIBUTES block is at [x29, #-104].
;   - All exit paths terminate via ExitProcess; control never returns.
;
; NOTE on repeated .lnk-detection logic: the "scan for space/quote, check
; last 4 chars against .lnk" pattern appears three times in this file.
; This is intentional duplication for clarity and register-state isolation,
; consistent with the original x64 design decision.
; ==============================================================================

; ==============================================================================
; EXTERNAL REFERENCES
; ==============================================================================

; Cross-module strings (defined in main.asm, EXPORT there)
    IMPORT str_runas
    IMPORT str_newSwitch
    IMPORT str_extLnk_m
    IMPORT str_space
    IMPORT str_outfileFlag
    IMPORT str_errfileFlag

; Cross-module jump target inside mainCRTStartup
    IMPORT mode_gui

; Win32 APIs
    IMPORT GetCommandLineW
    IMPORT LocalFree
    IMPORT CreateFileW
    IMPORT CloseHandle
    IMPORT ExitProcess

; Other modules
    IMPORT RunAsTrustedInstaller
    IMPORT ResolveLnkPath
    IMPORT skip_spaces
    IMPORT wcscpy_p
    IMPORT wcscat_p
    IMPORT wcscmp_ci
    IMPORT wcscmp_token
    IMPORT wcslen_p

; Global buffers (defined in main.asm BSS)
    IMPORT g_filePath
    IMPORT g_argsBuf
    IMPORT g_tempBuf
    IMPORT g_cmdBuf
    IMPORT g_relayPath
    IMPORT g_relayErrPath
    IMPORT g_relayHandle
    IMPORT g_relayErrHandle
    IMPORT g_useNewConsole
    IMPORT g_childExitCode

; ==============================================================================
; CODE SECTION
; ==============================================================================
    AREA |.text|, CODE, READONLY, ALIGN=4

    EXPORT mode_cli_found
    EXPORT mode_file_run

; ==============================================================================
; CONSTANTS FOR THIS MODULE
; ==============================================================================
GENERIC_WRITE       EQU 0x40000000
FILE_SHARE_READ     EQU 0x00000001
CREATE_ALWAYS       EQU 2
FILE_ATTRIBUTE_NORMAL EQU 0x00000080
INVALID_HANDLE_VALUE EQU 0xFFFFFFFFFFFFFFFF

; ==============================================================================
; mode_file_run - Right-click "Run as TrustedInstaller" file handler
;
; Triggered when argv[1] is not a recognized switch — typically a quoted
; file path delivered by the Explorer context menu. We resolve .lnk
; shortcuts ourselves so the target gets the TrustedInstaller token
; rather than launching cmd.exe with a shortcut argument.
;
; Entry state (from mainCRTStartup):
;   x20 = argv pointer (to be freed)
;   x29 = frame pointer of mainCRTStartup
; ==============================================================================
mode_file_run
    ; Free argv array allocated by CommandLineToArgvW
    MOV x0, x20                     ; x0 = argv (handle to free)
    BL LocalFree

    ; Get the raw command line to extract the file path
    BL GetCommandLineW
    MOV x19, x0                     ; x19 = command line pointer
    MOV x21, XZR                    ; x21 = quote state flag (0 = outside quotes)

    ; Skip past the executable path (which may be quoted)
skip_exe_for_file
    LDRH w0, [x19]
    CBZ w0, file_run_to_gui         ; No file argument found, show GUI
    CMP w0, #'"'
    B.NE fcheck_space
    EOR x21, x21, #1               ; Toggle quote state

fcheck_space
    CMP w0, #' '
    B.NE fnext_char
    CBNZ x21, fnext_char            ; Inside quotes: space is part of path

    ; Space ends executable path — skip whitespace after exe
    ADD x19, x19, #2
    MOV x0, x19
    BL skip_spaces
    MOV x19, x0                     ; x19 now points to file path
    B run_file_direct

fnext_char
    ADD x19, x19, #2
    B skip_exe_for_file

file_run_to_gui
    B mode_gui                      ; No arguments: fall back to GUI mode

; ==============================================================================
; run_file_direct - Process the file path argument
; ==============================================================================
run_file_direct
    ; x19 points to the file path argument (may be quoted from context menu).
    ; Check for .lnk shortcut and resolve before execution.

    ; Copy path to g_filePath, stripping leading quote
    MOV x1, x19                     ; x1 = source (rdi equivalent)
    LDRH w2, [x1]
    CMP w2, #'"'
    B.NE frf_no_leading_quote
    ADD x1, x1, #2                  ; Skip leading quote

frf_no_leading_quote
    ADRP x0, g_filePath
    ADD x0, x0, g_filePath
    BL wcscpy_p                     ; Copy to g_filePath

    ; Remove trailing quote if present
    ADRP x0, g_filePath
    ADD x0, x0, g_filePath
    BL wcslen_p
    MOV x22, x0                     ; x22 = string length

    CBZ x22, run_file_exec          ; Empty path, just execute

    ; Check if last character is a quote
    ADRP x1, g_filePath
    ADD x1, x1, g_filePath
    ADD x1, x1, x22, LSL #1        ; Point to end
    SUB x1, x1, #2                  ; Point to last char
    LDRH w2, [x1]
    CMP w2, #'"'
    B.NE frf_no_trailing_quote
    MOV w2, #0
    STRH w2, [x1]                   ; Null out trailing quote
    SUB x22, x22, #1                ; Adjust length

frf_no_trailing_quote
    ; Need at least 4 characters for .lnk extension
    CMP x22, #4
    B.LT run_file_exec

    ; Check if last 4 characters match ".lnk"
    ADRP x0, g_filePath
    ADD x0, x0, g_filePath
    ADD x0, x0, x22, LSL #1        ; Point to end
    SUB x0, x0, #8                  ; Point to start of ".lnk" (4 chars * 2 bytes)
    ADRP x1, str_extLnk_m
    ADD x1, x1, str_extLnk_m
    BL wcscmp_ci
    CBZ x0, run_file_exec           ; x0==0 (mismatch): not .lnk, execute directly

    ; --- .lnk detected: resolve shortcut ---

    ; Clear g_tempBuf (260 QWORDs = 2080 bytes) for shortcut arguments
    ADRP x10, g_tempBuf
    ADD x10, x10, g_tempBuf
    MOV x11, XZR
    MOV x12, #260
frf_clear_loop
    STR x11, [x10], #8
    SUBS x12, x12, #1
    B.NE frf_clear_loop

    ; Resolve .lnk target path and embedded arguments
    ; ResolveLnkPath(g_filePath, g_cmdBuf, g_tempBuf)
    ADRP x0, g_filePath
    ADD x0, x0, g_filePath
    ADRP x1, g_cmdBuf
    ADD x1, x1, g_cmdBuf
    ADRP x2, g_tempBuf
    ADD x2, x2, g_tempBuf
    BL ResolveLnkPath
    CBZ x0, run_file_exec           ; Resolution failed, try direct

    ; Build command line from resolved target and arguments
    ADRP x0, g_cmdBuf
    ADD x0, x0, g_cmdBuf
    BL wcslen_p
    CBZ x0, run_file_lnk_args       ; No target path

    ; Check for embedded arguments in .lnk
    ADRP x0, g_tempBuf
    ADD x0, x0, g_tempBuf
    BL wcslen_p
    CBZ x0, run_file_lnk_cmd        ; No embedded args, run target only

    ; Append space + embedded arguments to target
    ADRP x0, g_cmdBuf
    ADD x0, x0, g_cmdBuf
    ADRP x1, str_space
    ADD x1, x1, str_space
    BL wcscat_p

    ADRP x0, g_cmdBuf
    ADD x0, x0, g_cmdBuf
    ADRP x1, g_tempBuf
    ADD x1, x1, g_tempBuf
    BL wcscat_p
    B run_file_lnk_cmd

run_file_lnk_args
    ; No target path, use embedded arguments only
    ADRP x0, g_cmdBuf
    ADD x0, x0, g_cmdBuf
    ADRP x1, g_tempBuf
    ADD x1, x1, g_tempBuf
    BL wcscpy_p

run_file_lnk_cmd
    ; Execute resolved command as TrustedInstaller
    ADRP x0, g_cmdBuf
    ADD x0, x0, g_cmdBuf
    MOV x1, XZR                     ; useNewConsole = 0
    BL RunAsTrustedInstaller
    MOV x0, XZR
    BL ExitProcess

run_file_exec
    ; Not a .lnk file or resolution failed, execute as-is
    MOV x0, x19                     ; Original path pointer
    MOV x1, XZR                     ; useNewConsole = 0
    BL RunAsTrustedInstaller
    MOV x0, XZR
    BL ExitProcess

; ==============================================================================
; mode_cli_found - Dispatcher target for 'cmdt -cli <command>'
;
; Validates argc, recognizes the optional -new flag, then jumps to
; mode_cli_setup which parses the actual command. Frees the argv array
; before transferring control so the rest of the flow can rely on a
; raw GetCommandLineW pointer.
;
; Entry state:
;   x20 = argv pointer
;   [x29, #-64] = argc (32-bit value stored by mainCRTStartup)
; ==============================================================================
mode_cli_found
    ; CLI mode detected - check for minimum arguments (exe, -cli, command)
    LDR w0, [x29, #-64]            ; w0 = argc
    CMP w0, #3
    B.LT cli_no_cmd_free            ; Error: no command specified

    ; Check if argv[2] is "-new" (new console flag)
    LDR x1, [x20, #16]             ; x1 = argv[2]
    ADRP x2, str_newSwitch
    ADD x2, x2, str_newSwitch
    MOV x0, x1
    BL wcscmp_ci
    CBZ x0, cli_no_new_flag         ; Not "-new" (x0==0 mismatch)

    ; "-new" flag found: need at least 4 args (exe, -cli, -new, command)
    LDR w0, [x29, #-64]
    CMP w0, #4
    B.LT cli_no_cmd_free            ; Error: no command after -new
    ADRP x1, g_useNewConsole
    ADD x1, x1, g_useNewConsole
    MOV w0, #1
    STR w0, [x1]                    ; g_useNewConsole = 1
    B cli_free_and_setup

cli_no_new_flag
    ADRP x1, g_useNewConsole
    ADD x1, x1, g_useNewConsole
    MOV w0, #0
    STR w0, [x1]                    ; g_useNewConsole = 0

cli_free_and_setup
    ; Free argv array allocated by CommandLineToArgvW
    MOV x0, x20
    BL LocalFree
    B mode_cli_setup

cli_no_cmd_free
    ; Error: insufficient arguments for CLI mode
    MOV x0, x20
    BL LocalFree
    MOV x0, #1                      ; Exit code 1 (error)
    BL ExitProcess

; ==============================================================================
; mode_cli_setup - Parse the raw cmdline and dispatch to .lnk-aware runner
;
; Walks the GetCommandLineW string past the exe path, past "-cli", optionally
; honours the internal "-outfile " flag (relay protocol from the non-admin
; parent), optionally skips "-new", then locates the actual user command.
; If the command looks like a path-to-.lnk it is resolved into the real
; target before RunAsTrustedInstaller is invoked.
; ==============================================================================
mode_cli_setup
    BL GetCommandLineW
    MOV x19, x0                     ; x19 = command line pointer
    MOV x21, XZR                    ; x21 = quote state flag
    MOV x23, XZR                    ; x23 = misc scratch

    ; --- Skip past the executable path (which may be quoted) ---
skip_exe_loop
    LDRH w0, [x19]
    CBZ w0, cli_failed_setup        ; Unexpected end of command line
    CMP w0, #'"'
    B.NE ses_check_space
    EOR x21, x21, #1               ; Toggle quote state

ses_check_space
    CMP w0, #' '
    B.NE ses_next_char
    CBNZ x21, ses_next_char         ; Inside quotes: space is part of path
    ADD x19, x19, #2               ; Space ends exe path
    B skip_switch_init

ses_next_char
    ADD x19, x19, #2
    B skip_exe_loop

skip_switch_init
    ; Skip leading spaces before CLI switch
    MOV x0, x19
    BL skip_spaces
    MOV x19, x0
    MOV x21, XZR                    ; Reset quote state

    ; --- Skip the CLI switch argument ("-cli") ---
skip_switch_loop
    LDRH w0, [x19]
    CBZ w0, cli_failed_setup
    CMP w0, #' '
    B.NE ssl_next
    ADD x19, x19, #2
    B after_switch

ssl_next
    ADD x19, x19, #2
    B skip_switch_loop

after_switch
    ; Skip spaces after CLI switch
    MOV x0, x19
    BL skip_spaces
    MOV x19, x0

    ; ===== Detect internal -outfile token (relay mode from non-admin parent) =====
    ; If present, the next token is "-outfile" followed by a (possibly
    ; quoted) temp file path. Open the file for inheritable write access
    ; and store the handle in g_relayHandle so RunAsTrustedInstaller will
    ; redirect the spawned process's stdout/stderr to it.
    MOV x0, x19
    ADRP x1, str_outfileFlag
    ADD x1, x1, str_outfileFlag
    BL wcscmp_token
    CBZ x0, after_outfile           ; Not -outfile (x0==0 mismatch), skip relay setup

    ; Advance past "-outfile" (8 wchars = 16 bytes) and following spaces
    ADD x19, x19, #16
    MOV x0, x19
    BL skip_spaces
    MOV x19, x0

    ; --- Extract path token into g_relayPath ---
    MOV x24, x19                    ; x24 = scan pointer (rdi equivalent)
    MOV x25, XZR                    ; x25 = 1 if path was quoted

    LDRH w0, [x24]
    CMP w0, #'"'
    B.NE outfile_scan_unquoted
    ADD x24, x24, #2               ; Skip opening quote
    MOV x19, x24                    ; x19 = first char of path
    MOV x25, #1                     ; Marked as quoted

outfile_scan_quoted
    LDRH w0, [x24]
    CBZ w0, outfile_copy
    CMP w0, #'"'
    B.EQ outfile_copy
    ADD x24, x24, #2
    B outfile_scan_quoted

outfile_scan_unquoted
    LDRH w0, [x24]
    CBZ w0, outfile_copy
    CMP w0, #' '
    B.EQ outfile_copy
    ADD x24, x24, #2
    B outfile_scan_unquoted

outfile_copy
    ; Copy [x19..x24) into g_relayPath, null-terminate
    ADRP x10, g_relayPath
    ADD x10, x10, g_relayPath
    MOV x11, x19                    ; Source start
    MOV w12, #0                     ; Character counter

outfile_copy_loop
    CMP x11, x24
    B.EQ outfile_copy_done
    CMP w12, #259                   ; MAX_PATH - 1
    B.HS cli_failed_setup
    LDRH w0, [x11]
    STRH w0, [x10]
    ADD x11, x11, #2
    ADD x10, x10, #2
    ADD w12, w12, #1
    B outfile_copy_loop

outfile_copy_done
    MOV w0, #0
    STRH w0, [x10]                  ; Null terminate

    ; Advance x19 past the path token
    MOV x19, x24
    CBZ x25, outfile_advance_spaces
    LDRH w0, [x19]
    CMP w0, #'"'
    B.NE outfile_advance_spaces
    ADD x19, x19, #2               ; Skip closing quote

outfile_advance_spaces
    MOV x0, x19
    BL skip_spaces
    MOV x19, x0

    ; --- Set up SECURITY_ATTRIBUTES at [x29, #-104] ---
    ; Layout (24 bytes):
    ;   +0:  DWORD nLength = 24
    ;   +4:  DWORD padding = 0
    ;   +8:  QWORD lpSecurityDescriptor = NULL
    ;   +16: DWORD bInheritHandle = TRUE
    ;   +20: DWORD padding = 0
    MOV w0, #24
    STR w0, [x29, #-104]           ; nLength
    MOV w0, #0
    STR w0, [x29, #-100]           ; padding
    STR x0, [x29, #-96]            ; lpSecurityDescriptor = NULL
    MOV w0, #1
    STR w0, [x29, #-88]            ; bInheritHandle = TRUE
    MOV w0, #0
    STR w0, [x29, #-84]            ; padding

    ; --- CreateFileW(g_relayPath, GENERIC_WRITE, FILE_SHARE_READ, &sa,
    ;                 CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL) ---
    ADRP x0, g_relayPath
    ADD x0, x0, g_relayPath   ; lpFileName
    MOV w1, #0x40000000             ; GENERIC_WRITE
    MOV w2, #1                      ; FILE_SHARE_READ
    ADD x3, x29, #-104             ; &SECURITY_ATTRIBUTES
    MOV w4, #2                      ; CREATE_ALWAYS
    MOV w5, #0x80                   ; FILE_ATTRIBUTE_NORMAL
    MOV x6, XZR                     ; hTemplateFile = NULL
    ; Stack args for CreateFileW (7 params, 4 in regs + 3 on stack)
    STP x4, x5, [sp, #-16]!       ; Push dwCreationDisposition, dwFlagsAndAttributes
    STR x6, [sp, #-16]!            ; Push hTemplateFile
    SUB sp, sp, #16                ; Shadow space alignment
    BL CreateFileW
    ADD sp, sp, #48                 ; Restore stack (16+16+16)

    ; Check for INVALID_HANDLE_VALUE (-1)
    ADRP x1, invalid_handle_const
    ADD x1, x1, invalid_handle_const
    LDR x2, [x1]
    CMP x0, x2
    B.EQ after_outfile              ; CreateFile failed -> ignore relay

    ADRP x1, g_relayHandle
    ADD x1, x1, g_relayHandle
    STR x0, [x1]                    ; g_relayHandle = handle

    ; --- Detect -errfile token (stderr relay) ---
    MOV x0, x19
    ADRP x1, str_errfileFlag
    ADD x1, x1, str_errfileFlag
    BL wcscmp_token
    CBZ x0, after_outfile           ; Not -errfile (x0==0 mismatch)

    ADD x19, x19, #16              ; Skip "-errfile"
    MOV x0, x19
    BL skip_spaces
    MOV x19, x0

    ; --- Extract errfile path into g_relayErrPath ---
    MOV x24, x19
    MOV x25, XZR
    LDRH w0, [x24]
    CMP w0, #'"'
    B.NE errfile_scan_unquoted
    ADD x24, x24, #2
    MOV x19, x24
    MOV x25, #1

errfile_scan_quoted
    LDRH w0, [x24]
    CBZ w0, errfile_copy
    CMP w0, #'"'
    B.EQ errfile_copy
    ADD x24, x24, #2
    B errfile_scan_quoted

errfile_scan_unquoted
    LDRH w0, [x24]
    CBZ w0, errfile_copy
    CMP w0, #' '
    B.EQ errfile_copy
    ADD x24, x24, #2
    B errfile_scan_unquoted

errfile_copy
    ADRP x10, g_relayErrPath
    ADD x10, x10, g_relayErrPath
    MOV x11, x19
    MOV w12, #0

errfile_copy_loop
    CMP x11, x24
    B.EQ errfile_copy_done
    CMP w12, #259
    B.HS cli_failed_setup
    LDRH w0, [x11]
    STRH w0, [x10]
    ADD x11, x11, #2
    ADD x10, x10, #2
    ADD w12, w12, #1
    B errfile_copy_loop

errfile_copy_done
    MOV w0, #0
    STRH w0, [x10]

    MOV x19, x24
    CBZ x25, errfile_advance
    LDRH w0, [x19]
    CMP w0, #'"'
    B.NE errfile_advance
    ADD x19, x19, #2

errfile_advance
    MOV x0, x19
    BL skip_spaces
    MOV x19, x0

    ; --- CreateFileW for stderr relay ---
    ADRP x0, g_relayErrPath
    ADD x0, x0, g_relayErrPath
    MOV w1, #0x40000000             ; GENERIC_WRITE
    MOV w2, #1                      ; FILE_SHARE_READ
    ADD x3, x29, #-104             ; &SECURITY_ATTRIBUTES (reuse)
    MOV w4, #2                      ; CREATE_ALWAYS
    MOV w5, #0x80                   ; FILE_ATTRIBUTE_NORMAL
    MOV x6, XZR
    STP x4, x5, [sp, #-16]!
    STR x6, [sp, #-16]!
    SUB sp, sp, #16
    BL CreateFileW
    ADD sp, sp, #48

    ADRP x1, invalid_handle_const
    ADD x1, x1, invalid_handle_const
    LDR x2, [x1]
    CMP x0, x2
    B.EQ cli_failed_setup           ; Failed to create stderr relay

    ADRP x1, g_relayErrHandle
    ADD x1, x1, g_relayErrHandle
    STR x0, [x1]

after_outfile
    ; --- Check if we need to skip the "-new" token ---
    ADRP x0, g_useNewConsole
    ADD x0, x0, g_useNewConsole
    LDR w0, [x0]
    CBZ w0, run_command             ; No "-new" flag: proceed to command

skip_new_token
    ; Skip the "-new" token characters until space or null
    LDRH w0, [x19]
    CBZ w0, cli_failed_setup
    CMP w0, #' '
    B.NE snt_next
    ADD x19, x19, #2               ; Skip space after "-new"
    B run_command

snt_next
    ADD x19, x19, #2
    B skip_new_token

; ==============================================================================
; run_command - Locate and dispatch the actual user command
; ==============================================================================
run_command
    ; Skip spaces before the actual command
    MOV x0, x19
    BL skip_spaces
    MOV x19, x0                     ; x19 now points to the command

    ; Check if the command is a .lnk file (shortcut)
    MOV x0, x19
    BL wcslen_p
    MOV x26, x0                     ; x26 = command length
    CMP x26, #4                     ; Minimum length for ".lnk"
    B.LT run_no_lnk

    ; --- Scan for space or quote to find end of path ---
    MOV x24, x19                    ; x24 = scan pointer
    MOV x27, XZR                    ; x27 = pointer to first space (if found)
    MOV x28, XZR                    ; x28 = quote state

find_space_or_quote
    LDRH w0, [x24]
    CBZ w0, check_lnk_ext           ; End of string
    CMP w0, #'"'
    B.NE fsq_check_space
    EOR x28, x28, #1               ; Toggle quote state

fsq_check_space
    CMP w0, #' '
    B.NE fsq_next
    CBNZ x28, fsq_next              ; Inside quotes: ignore space
    MOV x27, x24                    ; Mark first space position
    B check_lnk_ext                 ; Stop at first space

fsq_next
    ADD x24, x24, #2
    B find_space_or_quote

check_lnk_ext
    ; Check if path (before first space) ends with .lnk
    CBNZ x27, check_bounded_path    ; Space found: bounded path

    ; --- No space found: check entire string ---
check_whole_path
    MOV x10, x19                    ; x10 = start
    MOV x11, x26                    ; x11 = length in chars

    ; Strip leading quote
    LDRH w0, [x10]
    CMP w0, #'"'
    B.NE cwp_no_leading
    ADD x10, x10, #2
    SUB x11, x11, #1

cwp_no_leading
    CMP x11, #1
    B.LT run_no_lnk

    ; Strip trailing quote
    ADD x12, x10, x11, LSL #1      ; End pointer
    SUB x12, x12, #2               ; Last char
    LDRH w0, [x12]
    CMP w0, #'"'
    B.NE cwp_no_trailing
    SUB x11, x11, #1

cwp_no_trailing
    CMP x11, #4
    B.LT run_no_lnk

    ; Check if last 4 characters are ".lnk"
    ADD x0, x10, x11, LSL #1       ; End
    SUB x0, x0, #8                  ; Start of ".lnk"
    ADRP x1, str_extLnk_m
    ADD x1, x1, str_extLnk_m
    BL wcscmp_ci
    CBZ x0, run_no_lnk             ; Not a .lnk file (x0==0 mismatch)

    ; --- Resolve the whole-path .lnk ---
    ; Clear g_tempBuf
    ADRP x10, g_tempBuf
    ADD x10, x10, g_tempBuf
    MOV x11, XZR
    MOV x12, #260
cwp_clear
    STR x11, [x10], #8
    SUBS x12, x12, #1
    B.NE cwp_clear

    ; Copy clean path to g_filePath
    ADRP x13, g_filePath
    ADD x13, x13, g_filePath
    MOV x14, x10                    ; Restore start (was modified by clear loop)
    ; Re-derive start: we stored it in x10 before the clear. Need to re-compute.
    ; Actually, x10 was the start of path. Let's use a saved copy.
    ; We'll re-derive from x19.
    MOV x14, x19                    ; Original command start
    LDRH w0, [x14]
    CMP w0, #'"'
    B.NE cwp_copy_start
    ADD x14, x14, #2
cwp_copy_start
    MOV x12, x26                    ; Length
    ; Recompute clean length (strip quotes)
    ; For simplicity, re-scan
    MOV x15, x14
    MOV w16, #0
cwp_copy_loop
    CMP w16, w12
    B.HS cwp_copy_done
    LDRH w0, [x15]
    CBZ w0, cwp_copy_done
    CMP w0, #'"'
    B.EQ cwp_copy_done
    STRH w0, [x13]
    ADD x15, x15, #2
    ADD x13, x13, #2
    ADD w16, w16, #1
    B cwp_copy_loop

cwp_copy_done
    MOV w0, #0
    STRH w0, [x13]                  ; Null terminate

    ; Resolve the .lnk file
    ADRP x0, g_filePath
    ADD x0, x0, g_filePath
    ADRP x1, g_cmdBuf
    ADD x1, x1, g_cmdBuf
    ADRP x2, g_tempBuf
    ADD x2, x2, g_tempBuf
    BL ResolveLnkPath
    CBZ x0, run_no_lnk             ; Resolution failed

    ; Build command from resolved target and arguments
    ADRP x0, g_cmdBuf
    ADD x0, x0, g_cmdBuf
    BL wcslen_p
    CBZ x0, use_lnk_args_only

    ; Append space after target path
    ADRP x1, g_cmdBuf
    ADD x1, x1, g_cmdBuf
    ADD x1, x1, x0, LSL #1         ; End of target
    MOV w2, #' '
    STRH w2, [x1]
    ADD x1, x1, #2
    MOV w2, #0
    STRH w2, [x1]

use_lnk_args_only
    ; Append .lnk arguments
    ADRP x0, g_cmdBuf
    ADD x0, x0, g_cmdBuf
    ADRP x1, g_tempBuf
    ADD x1, x1, g_tempBuf
    BL wcscat_p
    B run_resolved

    ; --- Bounded path (space found) ---
check_bounded_path
    ; x27 = pointer to first space, x19 = start of command
    ; Compute clean bounds (strip surrounding quotes from path)
    MOV x10, x19                    ; x10 = start of path
    MOV x11, x27                    ; x11 = end of path (space position)

    ; Strip leading quote
    LDRH w0, [x10]
    CMP w0, #'"'
    B.NE cbp_no_leading
    ADD x10, x10, #2

cbp_no_leading
    ; Strip trailing quote (char before space)
    SUB x12, x11, #2               ; Point to char before space
    LDRH w0, [x12]
    CMP w0, #'"'
    B.NE cbp_no_trailing
    SUB x11, x11, #2

cbp_no_trailing
    ; Compute path length in characters
    SUB x13, x11, x10
    LSR x13, x13, #1               ; Divide by 2 (bytes -> chars)
    CMP x13, #4
    B.LT run_no_lnk                 ; Too short for .lnk

    ; Check if last 4 characters are ".lnk"
    ADD x0, x10, x13, LSL #1       ; End of path
    SUB x0, x0, #8                  ; Start of ".lnk"
    ADRP x1, str_extLnk_m
    ADD x1, x1, str_extLnk_m
    BL wcscmp_ci
    CBZ x0, run_no_lnk             ; Not a .lnk file (x0==0 mismatch)

    ; --- Save context and resolve ---
    ; Copy clean path (without quotes) to g_filePath
    ADRP x14, g_filePath
    ADD x14, x14, g_filePath
    MOV x15, x13                    ; Character count

cbp_copy_loop
    CBZ x15, cbp_copy_done
    LDRH w0, [x10]
    STRH w0, [x14]
    ADD x10, x10, #2
    ADD x14, x14, #2
    SUB x15, x15, #1
    B cbp_copy_loop

cbp_copy_done
    MOV w0, #0
    STRH w0, [x14]                  ; Null terminate

    ; Extract arguments after the path
    ADD x0, x27, #2                ; Skip the space
    BL skip_spaces
    MOV x19, x0                     ; x19 = args start

    ; Copy arguments to g_argsBuf
    ADRP x0, g_argsBuf
    ADD x0, x0, g_argsBuf
    MOV x1, x19
    BL wcscpy_p

    ; Clear g_tempBuf (for shortcut arguments from .lnk)
    ADRP x10, g_tempBuf
    ADD x10, x10, g_tempBuf
    MOV x11, XZR
    MOV x12, #260
cbp_clear
    STR x11, [x10], #8
    SUBS x12, x12, #1
    B.NE cbp_clear

    ; Resolve the .lnk file to get target path and built-in arguments
    ADRP x0, g_filePath
    ADD x0, x0, g_filePath
    ADRP x1, g_cmdBuf
    ADD x1, x1, g_cmdBuf
    ADRP x2, g_tempBuf
    ADD x2, x2, g_tempBuf
    BL ResolveLnkPath
    CBZ x0, run_no_lnk             ; Resolution failed

    ; Build final command: target + .lnk args + user args
    ADRP x0, g_cmdBuf
    ADD x0, x0, g_cmdBuf
    BL wcslen_p
    CBZ x0, use_args_only

    ; Append space after target path
    ADRP x1, g_cmdBuf
    ADD x1, x1, g_cmdBuf
    ADD x1, x1, x0, LSL #1
    MOV w2, #' '
    STRH w2, [x1]
    ADD x1, x1, #2
    MOV w2, #0
    STRH w2, [x1]

use_args_only
    ; Append .lnk arguments
    ADRP x0, g_cmdBuf
    ADD x0, x0, g_cmdBuf
    ADRP x1, g_tempBuf
    ADD x1, x1, g_tempBuf
    BL wcscat_p

    ; Check if user provided additional arguments
    ADRP x0, g_argsBuf
    ADD x0, x0, g_argsBuf
    BL wcslen_p
    CBZ x0, run_resolved            ; No user args

    ; Append space and user arguments
    ADRP x0, g_cmdBuf
    ADD x0, x0, g_cmdBuf
    ADRP x1, str_space
    ADD x1, x1, str_space
    BL wcscat_p

    ADRP x0, g_cmdBuf
    ADD x0, x0, g_cmdBuf
    ADRP x1, g_argsBuf
    ADD x1, x1, g_argsBuf
    BL wcscat_p
    B run_resolved

; ==============================================================================
; run_resolved - Execute the resolved command as TrustedInstaller
; ==============================================================================
run_resolved
    ADRP x1, g_useNewConsole
    ADD x1, x1, g_useNewConsole
    LDR w1, [x1]                    ; x1 = useNewConsole flag
    ADRP x0, g_cmdBuf
    ADD x0, x0, g_cmdBuf
    BL RunAsTrustedInstaller
    MOV x20, x0                     ; Preserve result in x20
    B run_check_result

; ==============================================================================
; run_no_lnk - Execute the original command (not a .lnk file)
; ==============================================================================
run_no_lnk
    ADRP x1, g_useNewConsole
    ADD x1, x1, g_useNewConsole
    LDR w1, [x1]
    MOV x0, x19                     ; Original command pointer
    BL RunAsTrustedInstaller
    MOV x20, x0                     ; Preserve result

run_check_result
    ; Close the relay file handle (if any) so the parent reads a flushed,
    ; complete file. OS would close on exit anyway, but doing it here is
    ; explicit and lets the parent observe the file immediately.
    ADRP x0, g_relayHandle
    ADD x0, x0, g_relayHandle
    LDR x0, [x0]
    CBZ x0, close_err_handle
    BL CloseHandle
    ADRP x0, g_relayHandle
    ADD x0, x0, g_relayHandle
    STR XZR, [x0]                  ; g_relayHandle = 0

close_err_handle
    ADRP x0, g_relayErrHandle
    ADD x0, x0, g_relayErrHandle
    LDR x0, [x0]
    CBZ x0, check_exec_result
    BL CloseHandle
    ADRP x0, g_relayErrHandle
    ADD x0, x0, g_relayErrHandle
    STR XZR, [x0]                  ; g_relayErrHandle = 0

check_exec_result
    ; Check execution result
    CBZ x20, cli_failed             ; RunAsTrustedInstaller returned FALSE

    ; Propagate the command's exit status to cmd.exe/PowerShell
    ADRP x0, g_childExitCode
    ADD x0, x0, g_childExitCode
    LDR w0, [x0]
    BL ExitProcess

; ==============================================================================
; Error paths
; ==============================================================================
cli_failed_setup
cli_failed
    ; Close relay handles on failure path too
    ADRP x0, g_relayHandle
    ADD x0, x0, g_relayHandle
    LDR x0, [x0]
    CBZ x0, fail_close_err
    BL CloseHandle
    ADRP x0, g_relayHandle
    ADD x0, x0, g_relayHandle
    STR XZR, [x0]

fail_close_err
    ADRP x0, g_relayErrHandle
    ADD x0, x0, g_relayErrHandle
    LDR x0, [x0]
    CBZ x0, fail_exit
    BL CloseHandle
    ADRP x0, g_relayErrHandle
    ADD x0, x0, g_relayErrHandle
    STR XZR, [x0]

fail_exit
    ; Failure: exit with code 1
    MOV x0, #1
    BL ExitProcess

; ==============================================================================
; DATA SECTION (module-local constants)
; ==============================================================================
    AREA |.const|, DATA, READONLY, ALIGN=3

invalid_handle_const
    DCQ 0xFFFFFFFFFFFFFFFF          ; INVALID_HANDLE_VALUE

    END