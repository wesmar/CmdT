; ==============================================================================
; CMDT - Run as TrustedInstaller
; Process Creation Module
;
; Author: Marek Wesolowski (wesmar)
; Ported to: Native ARM64 (armasm64) for Windows on ARM
;
; Purpose: Handles process creation with TrustedInstaller privileges using
;          CreateProcessWithTokenW API. Supports three modes:
;          Mode 1 - Inherit standard handles from parent (GUI/CLI)
;          Mode 2 - Create new console window (-new flag)
;          Mode 3 - Relay mode (stdout/stderr to temp file)
;
; Exported procedures:
;   PrepareBatchCommand(cmdLine) - Detects .cmd/.bat and wraps in cmd.exe
;   RunAsTrustedInstaller(cmdLine, useNewConsole) - Creates TI process
;
; ARM64 Port Notes:
;   - x64 shadow space (32 bytes before each CALL) eliminated; ARM64 ABI
;     passes first 8 args in x0-x7 with no shadow space requirement.
;   - CreateProcessWithTokenW: 9 params → x0-x7 + 1 on stack.
;   - DuplicateHandle: 7 params → all in x0-x6.
;   - CreateFileW: 7 params → all in x0-x6.
;   - Stack frame: 208 bytes for STARTUPINFOW + PROCESS_INFORMATION +
;     environment pointer + flags + SECURITY_ATTRIBUTES + relay handle.
; ==============================================================================

; ==============================================================================
; EXTERNAL REFERENCES
; ==============================================================================

    IMPORT GetTIToken
    IMPORT CreateEnvironmentBlock
    IMPORT DestroyEnvironmentBlock
    IMPORT CreateProcessWithTokenW
    IMPORT CloseHandle
    IMPORT GetSystemDirectoryW
    IMPORT GetCurrentDirectoryW
    IMPORT GetStdHandle
    IMPORT GetCurrentProcess
    IMPORT DuplicateHandle
    IMPORT WaitForSingleObject
    IMPORT GetExitCodeProcess
    IMPORT CreateFileW
    IMPORT GetLastError
    IMPORT wcslen_p

; Global variables (defined in main.asm)
    IMPORT g_relayHandle
    IMPORT g_relayErrHandle
    IMPORT g_childExitCode

; ==============================================================================
; WINDOWS API CONSTANTS
; ==============================================================================

STARTUPINFOW_SIZE       EQU 104
STARTF_USESTDHANDLES    EQU 0x00000100
STARTF_USESHOWWINDOW    EQU 0x00000001
SW_SHOWNORMAL           EQU 1
CREATE_UNICODE_ENVIRONMENT EQU 0x00000400
CREATE_NEW_CONSOLE      EQU 0x00000010
CREATE_NO_WINDOW        EQU 0x08000000
LOGON_WITH_PROFILE      EQU 1
INFINITE                EQU 0xFFFFFFFF
STD_INPUT_HANDLE        EQU -10
STD_OUTPUT_HANDLE       EQU -11
STD_ERROR_HANDLE        EQU -12
DUPLICATE_SAME_ACCESS   EQU 2
OPEN_EXISTING           EQU 3
FILE_ATTRIBUTE_NORMAL   EQU 0x00000080
GENERIC_READ            EQU 0x80000000
FILE_SHARE_READ         EQU 0x00000001
FILE_SHARE_WRITE        EQU 0x00000002
INVALID_HANDLE_VALUE    EQU 0xFFFFFFFFFFFFFFFF

; ==============================================================================
; CONSTANT STRING DATA (READ-ONLY)
; ==============================================================================
    AREA |.const|, DATA, READONLY, ALIGN=3

str_nulDevice       DCW 'N','U','L',0
str_batchPrefix     DCW 'c','m','d','.','e','x','e',' ','/','d',' ','/','s',' ','/','c',' ',' ',0x22,0

; ==============================================================================
; UNINITIALIZED DATA SECTION (BSS)
; ==============================================================================
    AREA |.bss|, DATA, READWRITE, NOINIT, ALIGN=3

; Child working directory and private wrapper for .cmd/.bat command lines.
sysDirBuf           SPACE 520
batchCmdBuf         SPACE 65536

; ==============================================================================
; CODE SECTION
; ==============================================================================
    AREA |.text|, CODE, READONLY, ALIGN=4

    EXPORT PrepareBatchCommand
    EXPORT RunAsTrustedInstaller

; ==============================================================================
; PrepareBatchCommand - Detect .cmd/.bat and Wrap in cmd.exe
;
; Purpose: CreateProcess does not invoke the command interpreter for batch
;          files. This function detects a .cmd/.bat first token and returns
;          a private command line wrapped as:
;            cmd.exe /d /s /c "<original command line>"
;          Non-batch commands are returned unchanged.
;
; Parameters:
;   x0 = Pointer to command line string
;
; Returns:
;   x0 = Pointer to command line (original or wrapped in batchCmdBuf)
;
; Register allocation:
;   x19 = command line pointer (preserved across calls)
;   x20 = scan pointer
;   x21 = token length (bytes)
;   x22 = temp pointer for extension check
;   x23 = destination pointer (batchCmdBuf)
;   x24 = source pointer / prefix pointer
;   w25 = capacity counter
;
; Algorithm:
;   1. Strip outer quotes if present (before extension scan)
;   2. Scan first token (until space, tab, quote, or null)
;   3. Check if last 4 chars of token match ".cmd" or ".bat" (case-insensitive)
;   4. If match, build wrapped command in batchCmdBuf
;   5. Return pointer to original or wrapped command
; ==============================================================================
PrepareBatchCommand PROC
    STP x29, x30, [sp, #-16]!
    MOV x29, sp
    STP x19, x20, [sp, #-16]!
    STP x21, x22, [sp, #-16]!
    STP x23, x24, [sp, #-16]!
    STP x25, x26, [sp, #-16]!

    MOV x19, x0                 ; x19 = command line
    MOV x0, x19                 ; Default return: original command line

    ; --- Strip outer quotes if present ---
    ; This must happen BEFORE the .cmd/.bat extension scan: that scan
    ; measures the token length and reads backward from its end to compare
    ; against ".cmd"/".bat". A trailing quote would shift every one of
    ; those offsets by one WCHAR and make the extension check miss a
    ; legitimately quoted batch file.
    LDRH w1, [x19]
    CMP w1, #0x22               ; '"'
    B.NE pbc_no_strip

    ; Get length of command line
    MOV x0, x19
    BL wcslen_p
    CBZ x0, pbc_no_strip        ; Empty string

    ; Check if last char is also a quote
    ; x2 = x19 + len*2 - 2 (pointer to last char)
    ADD x2, x19, x0, LSL #1
    SUB x2, x2, #2
    LDRH w1, [x2]
    CMP w1, #0x22               ; '"'
    B.NE pbc_no_strip

    ; Strip quotes: skip opening, null out closing
    ADD x19, x19, #2            ; Skip opening quote
    MOV w1, #0
    STRH w1, [x2]               ; Null out closing quote
    MOV x0, x19                 ; Update default return

pbc_no_strip
    ; --- Scan first token ---
    MOV x20, x19                ; x20 = scan pointer

    ; Check if first char is a quote (quoted path)
    LDRH w1, [x20]
    CMP w1, #'"'
    B.NE pbc_scan_unquoted

    ; Quoted scan: look for closing quote
    ADD x20, x20, #2            ; Skip opening quote

pbc_scan_quoted
    LDRH w1, [x20]
    CBZ w1, pbc_return          ; End of string (malformed, but handle gracefully)
    CMP w1, #'"'
    B.EQ pbc_token_end
    ADD x20, x20, #2
    B pbc_scan_quoted

pbc_scan_unquoted
    ; Unquoted scan: look for space, tab, or null
    LDRH w1, [x20]
    CBZ w1, pbc_token_end
    CMP w1, #' '
    B.EQ pbc_token_end
    CMP w1, #9                  ; Tab
    B.EQ pbc_token_end
    ADD x20, x20, #2
    B pbc_scan_unquoted

pbc_token_end
    ; x20 = one past last char of token
    ; Compute token length in bytes: x21 = x20 - x19
    SUB x21, x20, x19

    ; If first char was a quote, exclude it from token length
    LDRH w1, [x19]
    CMP w1, #'"'
    B.NE pbc_check_ext
    SUB x21, x21, #2            ; Exclude opening quote from length

pbc_check_ext
    ; Check if token is at least 4 chars (8 bytes) for ".cmd"/".bat"
    CMP x21, #8
    B.LT pbc_return

    ; Extension check reads the last 4 WCHARs of the token backward from
    ; its end (x20 = one-past-last character):
    ;   [x20-8] is 4 chars back (should be '.')
    ;   [x20-6] is 3 chars back (should be 'c' or 'b')
    ;   [x20-4] is 2 chars back (should be 'm'/'d' or 'a')
    ;   [x20-2] is 1 char back  (should be 'd' or 't')
    ; Each offset is -2 bytes because characters are UTF-16 (2 bytes each).

    ; Check '.' at [x20-8]
    SUB x22, x20, #8
    LDRH w1, [x22]
    CMP w1, #'.'
    B.NE pbc_return

    ; Check 'c' or 'b' at [x20-6] (case-insensitive)
    SUB x22, x20, #6
    LDRH w1, [x22]
    ORR w1, w1, #0x20           ; Convert to lowercase
    CMP w1, #'c'
    B.EQ pbc_check_cmd
    CMP w1, #'b'
    B.NE pbc_return

    ; Check 'a' at [x20-4]
    SUB x22, x20, #4
    LDRH w1, [x22]
    ORR w1, w1, #0x20
    CMP w1, #'a'
    B.NE pbc_return

    ; Check 't' at [x20-2]
    SUB x22, x20, #2
    LDRH w1, [x22]
    ORR w1, w1, #0x20
    CMP w1, #'t'
    B.NE pbc_return
    B pbc_wrap                  ; It's ".bat"

pbc_check_cmd
    ; Check 'm' at [x20-4]
    SUB x22, x20, #4
    LDRH w1, [x22]
    ORR w1, w1, #0x20
    CMP w1, #'m'
    B.NE pbc_return

    ; Check 'd' at [x20-2]
    SUB x22, x20, #2
    LDRH w1, [x22]
    ORR w1, w1, #0x20
    CMP w1, #'d'
    B.NE pbc_return
    ; It's ".cmd" - fall through to pbc_wrap

pbc_wrap
    ; Build wrapped command: cmd.exe /d /s /c "<original>"
    ADRP x23, batchCmdBuf
    ADD x23, x23, batchCmdBuf
    ADRP x24, str_batchPrefix
    ADD x24, x24, str_batchPrefix

pbc_copy_prefix
    LDRH w1, [x24]
    STRH w1, [x23]
    ADD x24, x24, #2
    ADD x23, x23, #2
    CBZ w1, pbc_prefix_done
    B pbc_copy_prefix

pbc_prefix_done
    SUB x23, x23, #2            ; Overwrite the prefix terminator

    ; Copy original command line into batchCmdBuf after prefix
    MOV x24, x19                ; x24 = source (original command)
    MOV w25, #32748             ; Capacity: 32768 - 18 (prefix) - 1 (closing quote) - 1 (NUL)

pbc_copy_command
    LDRH w1, [x24]
    CBZ w1, pbc_close           ; End of source
    CBZ w25, pbc_return         ; Buffer full (truncation, no overflow)
    STRH w1, [x23]
    ADD x24, x24, #2
    ADD x23, x23, #2
    SUB w25, w25, #1
    B pbc_copy_command

pbc_close
    ; Append closing quote and null terminator
    MOV w1, #'"'
    STRH w1, [x23]
    ADD x23, x23, #2
    MOV w1, #0
    STRH w1, [x23]

    ; Return pointer to wrapped command
    ADRP x0, batchCmdBuf
    ADD x0, x0, batchCmdBuf
    B pbc_return

pbc_return
    ; x0 contains either original command line or wrapped batchCmdBuf
    LDP x25, x26, [sp], #16
    LDP x23, x24, [sp], #16
    LDP x21, x22, [sp], #16
    LDP x19, x20, [sp], #16
    LDP x29, x30, [sp], #16
    RET
PrepareBatchCommand ENDP

; ==============================================================================
; RunAsTrustedInstaller - Execute Command with TrustedInstaller Privileges
;
; Purpose: Creates a new process running as TrustedInstaller using the token
;          obtained from GetTIToken. Supports three modes of operation:
;          Mode 1 - Inherit standard handles from parent (default GUI/CLI)
;          Mode 2 - Create new console window (-new flag)
;          Mode 3 - Relay mode (stdout/stderr redirected to temp file)
;
; Parameters:
;   x0 = Pointer to command line string
;   w1 = useNewConsole flag (0 = inherit/create hidden, 1 = new console)
;
; Returns:
;   x0 = 1 on success, 0 on failure
;
; Stack frame: 208 bytes
;   [sp+0..103]    STARTUPINFOW (104 bytes)
;                  cb at +0, dwFlags at +60, wShowWindow at +64,
;                  hStdInput at +80, hStdOutput at +88, hStdError at +96
;   [sp+104..127]  PROCESS_INFORMATION (24 bytes)
;                  hProcess at +104, hThread at +112, dwProcessId at +120,
;                  dwThreadId at +124
;   [sp+128..135]  lpEnvironment pointer (8 bytes)
;   [sp+136..139]  stdin duplicate-succeeded flag (DWORD, 0/1)
;   [sp+140..143]  stdout duplicate-succeeded flag (DWORD, 0/1)
;   [sp+144..147]  stderr duplicate-succeeded flag (DWORD, 0/1)
;   [sp+148..151]  child exit-code scratch (DWORD)
;   [sp+152..175]  SECURITY_ATTRIBUTES for relay-mode NUL-device open (24 bytes)
;   [sp+176..183]  relay-mode NUL-device read handle (8 bytes)
;   [sp+184..207]  padding for 16-byte alignment
;
; Register allocation:
;   x19 = useNewConsole flag
;   x20 = command line (from PrepareBatchCommand)
;   x21 = TrustedInstaller token handle
;   x22 = current process pseudo-handle
;   w23 = CreateProcess result (TRUE/FALSE)
; ==============================================================================
RunAsTrustedInstaller PROC
    ; Prologue: save FP/LR and callee-saved registers
    STP x29, x30, [sp, #-16]!
    MOV x29, sp
    STP x19, x20, [sp, #-16]!
    STP x21, x22, [sp, #-16]!
    STP x23, x24, [sp, #-16]!
    SUB sp, sp, #208

    MOV w19, w1                 ; w19 = useNewConsole flag
    MOV x0, x0                  ; x0 = command line (already in x0)

    ; Prepare command line (wrap .cmd/.bat if needed)
    BL PrepareBatchCommand
    MOV x20, x0                 ; x20 = executable or wrapped batch command

    ; Obtain TrustedInstaller token
    BL GetTIToken
    CBZ x0, rp_no_token         ; Failed to get token
    MOV x21, x0                 ; x21 = TrustedInstaller token handle

    ; --- Initialize STARTUPINFOW structure to zero ---
    ; 104 bytes = 13 qwords
    ADD x0, sp, #0              ; x0 = &STARTUPINFOW
    MOV x1, XZR
    MOV w2, #13
rp_zero_si
    STR x1, [x0], #8
    SUBS w2, w2, #1
    B.NE rp_zero_si

    ; Set structure size (cb field)
    MOV w0, #STARTUPINFOW_SIZE
    STR w0, [sp, #0]            ; cb at [sp+0]

    ; Initialize flags and handles
    MOV w0, #0
    STR w0, [sp, #136]          ; stdin duplicate flag = 0
    STR w0, [sp, #140]          ; stdout duplicate flag = 0
    STR w0, [sp, #144]          ; stderr duplicate flag = 0
    STR XZR, [sp, #176]         ; relay NUL-input handle = 0
    MOV w0, #1
    ADRP x1, g_childExitCode
    ADD x1, x1, g_childExitCode
    STR w0, [x1]                ; g_childExitCode = 1 (failure-safe default)

    ; --- Check for output-relay mode ---
    ADRP x0, g_relayHandle
    ADD x0, x0, g_relayHandle
    LDR x0, [x0]
    CBNZ x0, rp_relay_mode      ; Relay mode: redirect stdout/stderr

    ; --- Check console mode ---
    CBNZ w19, rp_new_console    ; -new flag: create new console window

    ; ===== Mode 1: Inherit standard handles from parent process =====
    ; cmd.exe can give this GUI-subsystem process usable redirected std
    ; handles that are not themselves inheritable by the TI child.
    ; Duplicate them as inheritable handles before placing in STARTUPINFO.

    ; Get current process pseudo-handle
    BL GetCurrentProcess
    MOV x22, x0                 ; x22 = current process handle

    ; --- Duplicate STDIN ---
    MOVN x0, #9                  ; STD_INPUT_HANDLE = -10
    BL GetStdHandle
    STR x0, [sp, #80]           ; hStdInput at [sp+80]
    CBZ x0, rp_dup_stdout       ; NULL handle
    MOV x1, #0xFFFFFFFFFFFFFFFF
    CMP x0, x1
    B.EQ rp_dup_stdout          ; INVALID_HANDLE_VALUE

    ; DuplicateHandle(currentProc, hStdInput, currentProc, &hStdInput,
    ;                 0, TRUE, DUPLICATE_SAME_ACCESS)
    MOV x0, x22                 ; hSourceProcessHandle
    LDR x1, [sp, #80]           ; hSourceHandle
    MOV x2, x22                 ; hTargetProcessHandle
    ADD x3, sp, #80             ; lpTargetHandle (overwrite in place)
    MOV w4, #0                  ; dwDesiredAccess (ignored)
    MOV w5, #1                  ; bInheritHandle = TRUE
    MOV w6, #2                  ; DUPLICATE_SAME_ACCESS
    BL DuplicateHandle
    CBZ w0, rp_dup_stdout       ; Duplication failed
    MOV w0, #1
    STR w0, [sp, #136]          ; stdin duplicate flag = 1

rp_dup_stdout
    ; --- Duplicate STDOUT ---
    MOVN x0, #10                 ; STD_OUTPUT_HANDLE = -11
    BL GetStdHandle
    STR x0, [sp, #88]           ; hStdOutput at [sp+88]
    CBZ x0, rp_dup_stderr
    MOV x1, #0xFFFFFFFFFFFFFFFF
    CMP x0, x1
    B.EQ rp_dup_stderr

    MOV x0, x22
    LDR x1, [sp, #88]
    MOV x2, x22
    ADD x3, sp, #88
    MOV w4, #0
    MOV w5, #1
    MOV w6, #2
    BL DuplicateHandle
    CBZ w0, rp_dup_stderr
    MOV w0, #1
    STR w0, [sp, #140]          ; stdout duplicate flag = 1

rp_dup_stderr
    ; --- Duplicate STDERR ---
    MOVN x0, #11                 ; STD_ERROR_HANDLE = -12
    BL GetStdHandle
    STR x0, [sp, #96]           ; hStdError at [sp+96]
    CBZ x0, rp_stdio_ready
    MOV x1, #0xFFFFFFFFFFFFFFFF
    CMP x0, x1
    B.EQ rp_stdio_ready

    MOV x0, x22
    LDR x1, [sp, #96]
    MOV x2, x22
    ADD x3, sp, #96
    MOV w4, #0
    MOV w5, #1
    MOV w6, #2
    BL DuplicateHandle
    CBZ w0, rp_stdio_ready
    MOV w0, #1
    STR w0, [sp, #144]          ; stderr duplicate flag = 1

rp_stdio_ready
    ; Set STARTF_USESTDHANDLES flag
    MOV w0, #STARTF_USESTDHANDLES
    STR w0, [sp, #60]           ; dwFlags at [sp+60]
    B rp_setup_env

rp_relay_mode
    ; ===== Mode 3: Redirect stdout/stderr to temp file, stdin from NUL =====
    ; STARTF_USESTDHANDLES requires valid handles; NULL is not a substitute
    ; for the Windows NUL device. Open NUL for inheritable read access.

    ; Set up SECURITY_ATTRIBUTES at [sp+152] (24 bytes)
    MOV w0, #24
    STR w0, [sp, #152]          ; nLength
    MOV w0, #0
    STR w0, [sp, #156]          ; padding
    STR XZR, [sp, #160]         ; lpSecurityDescriptor = NULL
    MOV w0, #1
    STR w0, [sp, #168]          ; bInheritHandle = TRUE
    MOV w0, #0
    STR w0, [sp, #172]          ; padding

    ; CreateFileW("NUL", GENERIC_READ, FILE_SHARE_READ|FILE_SHARE_WRITE,
    ;             &sa, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, NULL)
    ADRP x0, str_nulDevice
    ADD x0, x0, str_nulDevice
    MOV w1, #0x80000000         ; GENERIC_READ
    MOV w2, #3                  ; FILE_SHARE_READ | FILE_SHARE_WRITE
    ADD x3, sp, #152            ; &SECURITY_ATTRIBUTES
    MOV w4, #3                  ; OPEN_EXISTING
    MOV w5, #0x80               ; FILE_ATTRIBUTE_NORMAL
    MOV x6, XZR                 ; hTemplateFile = NULL
    BL CreateFileW

    MOV x1, #0xFFFFFFFFFFFFFFFF
    CMP x0, x1
    B.EQ rp_fail                ; Failed to open NUL

    STR x0, [sp, #176]          ; Save NUL handle
    STR x0, [sp, #80]           ; hStdInput = NUL handle

    ; Set hStdOutput to relay handle
    ADRP x0, g_relayHandle
    ADD x0, x0, g_relayHandle
    LDR x0, [x0]
    STR x0, [sp, #88]           ; hStdOutput

    ; Set hStdError to relay error handle (or fallback to stdout relay)
    ADRP x0, g_relayErrHandle
    ADD x0, x0, g_relayErrHandle
    LDR x0, [x0]
    CBNZ x0, rp_set_stderr
    ADRP x0, g_relayHandle
    ADD x0, x0, g_relayHandle
    LDR x0, [x0]                ; Backward-compatible fallback

rp_set_stderr
    STR x0, [sp, #96]           ; hStdError

    ; Set STARTF_USESTDHANDLES flag
    MOV w0, #STARTF_USESTDHANDLES
    STR w0, [sp, #60]
    B rp_setup_env

rp_new_console
    ; ===== Mode 2: Create new console window =====
    MOV w0, #STARTF_USESHOWWINDOW
    STR w0, [sp, #60]           ; dwFlags
    MOV w0, #SW_SHOWNORMAL
    STRH w0, [sp, #64]          ; wShowWindow

rp_setup_env
    ; --- Initialize PROCESS_INFORMATION structure to zero ---
    ADD x0, sp, #104
    MOV x1, XZR
    STR x1, [x0]                ; hProcess
    STR x1, [x0, #8]            ; hThread
    STR x1, [x0, #16]           ; dwProcessId + dwThreadId

    ; Initialize lpEnvironment pointer to NULL
    STR XZR, [sp, #128]

    ; Create environment block for the TrustedInstaller token
    ; CreateEnvironmentBlock(&lpEnvironment, hToken, FALSE)
    ADD x0, sp, #128            ; &lpEnvironment
    MOV x1, x21                 ; hToken
    MOV w2, #0                  ; bInherit = FALSE
    BL CreateEnvironmentBlock
    ; Continue even if this fails (environment will be NULL)

    ; --- Preserve the caller's working directory ---
    ; Required for relative commands such as `cmdt -cli maintenance.cmd`;
    ; fall back to System32 only if the current directory cannot be
    ; represented in this MAX_PATH buffer.
    MOV w0, #260
    ADRP x1, sysDirBuf
    ADD x1, x1, sysDirBuf
    BL GetCurrentDirectoryW
    CBZ w0, rp_workdir_fallback
    CMP w0, #260
    B.LT rp_workdir_ready

rp_workdir_fallback
    ADRP x0, sysDirBuf
    ADD x0, x0, sysDirBuf
    MOV w1, #260
    BL GetSystemDirectoryW

rp_workdir_ready
    ; --- Prepare parameters for CreateProcessWithTokenW ---
    ; 9 parameters: x0-x7 in registers, 1 on stack

    ; Determine creation flags based on mode
    ADRP x0, g_relayHandle
    ADD x0, x0, g_relayHandle
    LDR x0, [x0]
    CBNZ x0, rp_flags_relay     ; Relay mode

    CBNZ w19, rp_flags_new      ; New console mode

    ; Mode 1: Inherit handles (no special flags)
    MOV w0, #CREATE_UNICODE_ENVIRONMENT
    B rp_flags_done

rp_flags_new
    ; Mode 2: New console window
    MOV w0, #(CREATE_NEW_CONSOLE | CREATE_UNICODE_ENVIRONMENT)
    B rp_flags_done

rp_flags_relay
    ; Mode 3: Relay (no window). 0x08000400 spans two halfwords, so build it
    ; with MOVZ/MOVK rather than a single (illegal) MOV immediate.
    MOVZ w0, #0x0400            ; CREATE_NO_WINDOW | CREATE_UNICODE_ENVIRONMENT
    MOVK w0, #0x0800, LSL #16

rp_flags_done
    ; Set up register parameters for CreateProcessWithTokenW.
    ; The creation flags are still in w0 from rp_flags_* -- capture them into
    ; x4 (dwCreationFlags) BEFORE reusing x0 for hToken, otherwise the child
    ; gets hToken's low bits as its creation flags and the call fails.
    MOV w4, w0                  ; x4 = dwCreationFlags (must come first)
    MOV x0, x21                 ; x0 = hToken
    MOV w1, #LOGON_WITH_PROFILE ; x1 = dwLogonFlags
    MOV x2, XZR                 ; x2 = lpApplicationName = NULL
    MOV x3, x20                 ; x3 = lpCommandLine
    LDR x5, [sp, #128]          ; x5 = lpEnvironment
    ADRP x6, sysDirBuf
    ADD x6, x6, sysDirBuf ; x6 = lpCurrentDirectory
    ADD x7, sp, #0              ; x7 = lpStartupInfo

    ; Push lpProcessInformation on stack (9th parameter)
    SUB sp, sp, #16             ; 8 bytes for arg + 8 bytes padding
    ADD x8, sp, #16             ; x8 = original sp (before sub)
    ADD x8, x8, #104            ; x8 = &PROCESS_INFORMATION
    STR x8, [sp, #0]            ; lpProcessInformation

    BL CreateProcessWithTokenW
    ADD sp, sp, #16             ; Restore stack

    MOV w23, w0                 ; w23 = result (TRUE/FALSE)

    ; --- Destroy environment block ---
    LDR x0, [sp, #128]
    CBZ x0, rp_skip_destroy_env
    BL DestroyEnvironmentBlock

rp_skip_destroy_env
    ; --- Check if process creation succeeded ---
    CBZ w23, rp_fail

    ; --- Determine if we should wait for the child process ---
    ; Wait if we are in relay mode OR if we are in inherit mode (CLI)
    ADRP x0, g_relayHandle
    ADD x0, x0, g_relayHandle
    LDR x0, [x0]
    CBNZ x0, rp_do_wait         ; Always wait in relay mode

    CBNZ w19, rp_skip_wait      ; Don't wait in new-console/GUI mode

rp_do_wait
    ; Wait for child process to exit
    LDR x0, [sp, #104]          ; hProcess
    CBZ x0, rp_skip_wait
    MOV w1, #INFINITE
    BL WaitForSingleObject

    ; Preserve the actual program result independently of this routine's
    ; BOOL return value. Done before either process handle is closed.
    LDR x0, [sp, #104]          ; hProcess
    ADD x1, sp, #148            ; &exitCode scratch
    BL GetExitCodeProcess
    CBZ w0, rp_skip_wait        ; GetExitCodeProcess failed

    LDR w0, [sp, #148]          ; Get exit code
    ADRP x1, g_childExitCode
    ADD x1, x1, g_childExitCode
    STR w0, [x1]                ; g_childExitCode = child's exit code

rp_skip_wait
    ; --- Close inheritable duplicates made only for the child process ---
    ; Do this after the wait in CLI mode so redirected output remains open
    ; until the child has finished writing.

    ; Close stdin duplicate if we made one
    LDR w0, [sp, #136]
    CBZ w0, rp_close_dup_stdout
    LDR x0, [sp, #80]
    BL CloseHandle

rp_close_dup_stdout
    ; Close stdout duplicate if we made one
    LDR w0, [sp, #140]
    CBZ w0, rp_close_dup_stderr
    LDR x0, [sp, #88]
    BL CloseHandle

rp_close_dup_stderr
    ; Close stderr duplicate if we made one
    LDR w0, [sp, #144]
    CBZ w0, rp_close_nul
    LDR x0, [sp, #96]
    BL CloseHandle

rp_close_nul
    ; Close relay NUL-device handle if we opened one
    LDR x0, [sp, #176]
    CBZ x0, rp_close_pi
    BL CloseHandle
    STR XZR, [sp, #176]         ; Clear handle

rp_close_pi
    ; Close process and thread handles
    LDR x0, [sp, #104]          ; hProcess
    CBZ x0, rp_skip_hp
    BL CloseHandle

rp_skip_hp
    LDR x0, [sp, #112]          ; hThread
    CBZ x0, rp_skip_ht
    BL CloseHandle

rp_skip_ht
    MOV w0, #1                  ; Success
    B rp_done

rp_fail
    ; Get last error and store as exit code
    BL GetLastError
    ADRP x1, g_childExitCode
    ADD x1, x1, g_childExitCode
    STR w0, [x1]

    ; Close relay NUL handle if open
    LDR x0, [sp, #176]
    CBZ x0, rp_no_token
    BL CloseHandle

rp_no_token
    MOV w0, WZR                 ; Failure

rp_done
    ADD sp, sp, #208
    LDP x23, x24, [sp], #16
    LDP x21, x22, [sp], #16
    LDP x19, x20, [sp], #16
    LDP x29, x30, [sp], #16
    RET
RunAsTrustedInstaller ENDP

    END