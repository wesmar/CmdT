; ==============================================================================
; CMDT - Run as TrustedInstaller
; Non-Admin Output-Relay Module
;
; Author: Marek Wesolowski (wesmar)
; Ported to: Native ARM64 (armasm64) for Windows on ARM
;
; Purpose: Implements the relay path that lets `cmdt -cli <command>` run
;          from a non-admin shell still return its child's stdout/stderr to
;          the caller's redirect target or console. The trick: spawn an
;          elevated copy of cmdt with internal `-outfile` and `-errfile`
;          paths telling it to redirect spawned-process output into a temp
;          file, wait for that elevated process to exit, then stream the
;          temp file contents back to this (non-admin) process's
;          STD_OUTPUT — which cmd.exe wired up before launching us, so
;          `>file` / `|pipe` / `>>file` all work transparently.
;
; Exported routines:
;   NonAdminRelayLaunch(argc, argv, cmdline) - Attempt the relay path.
;     Returns 0 if it declined (e.g. user passed -new, which conflicts with
;     output capture; or temp-file setup failed). Never returns if the relay
;     actually ran — every success/failure past ShellExecuteEx ends in
;     ExitProcess.
;   AdminRelayLaunch(argc, argv, cmdline) - In-process relay for elevated
;     callers. Avoids the double-UAC-prompt regression by calling
;     RunAsTrustedInstaller directly with a temp-file relay handle.
;
; ARM64 Port Notes:
;   - Register allocation: x19=cmdline, x20=argv, x21=file handle,
;     x22=std handle, x23=table/result, x24=scan ptr, x25=quote flag,
;     w26=argc.
;   - Stack layout: 224 bytes total. SHELLEXECUTEINFOW (112) at [sp+32],
;     bytesRead at [sp+144], bytesWritten at [sp+152], exit status at
;     [sp+160]. All BL sites keep 16-byte alignment.
;   - All Win32 API calls follow ARM64 Windows ABI (x0-x3 args, shadow
;     space provided by callee, extra args on stack).
; ==============================================================================

; ==============================================================================
; EXTERNAL REFERENCES
; ==============================================================================

; Cross-module strings owned by main.asm
    IMPORT str_runas
    IMPORT str_newSwitch
    IMPORT str_outfileFlag
    IMPORT str_errfileFlag

; Win32 APIs
    IMPORT GetModuleFileNameW
    IMPORT GetTempPathW
    IMPORT GetTempFileNameW
    IMPORT ShellExecuteExW
    IMPORT WaitForSingleObject
    IMPORT GetExitCodeProcess
    IMPORT CloseHandle
    IMPORT CreateFileW
    IMPORT ReadFile
    IMPORT WriteFile
    IMPORT DeleteFileW
    IMPORT GetStdHandle
    IMPORT ExitProcess

; In-project helpers
    IMPORT wcscpy_p
    IMPORT wcscat_p
    IMPORT wcscmp_ci
    IMPORT skip_spaces
    IMPORT RunAsTrustedInstaller

; Global buffers (defined in main.asm)
    IMPORT g_exePath
    IMPORT g_tempDirBuf
    IMPORT g_relayPath
    IMPORT g_relayErrPath
    IMPORT g_relayArgs
    IMPORT g_relayReadBuf
    IMPORT g_cmdBuf
    IMPORT g_relayHandle
    IMPORT g_relayErrHandle
    IMPORT g_childExitCode

; ==============================================================================
; WINDOWS API CONSTANTS
; ==============================================================================

SEE_MASK_NOCLOSEPROCESS EQU 0x00000040
SW_HIDE                 EQU 0
GENERIC_READ            EQU 0x80000000
GENERIC_WRITE           EQU 0x40000000
FILE_SHARE_READ         EQU 0x00000001
CREATE_ALWAYS           EQU 2
OPEN_EXISTING           EQU 3
FILE_ATTRIBUTE_NORMAL   EQU 0x00000080
INFINITE                EQU 0xFFFFFFFF
STD_OUTPUT_HANDLE       EQU -11
STD_ERROR_HANDLE        EQU -12
INVALID_HANDLE_VALUE    EQU 0xFFFFFFFFFFFFFFFF

; ==============================================================================
; CONSTANT STRING DATA (READ-ONLY)
; ==============================================================================
    AREA |.const|, DATA, READONLY, ALIGN=3

; 3-char prefix used by GetTempFileNameW to build the temp file name.
; Three chars max per the API contract; the API itself appends a hex
; sequence + .TMP.
str_cmdtPrefix      DCW 'C','M','D',0

; Fixed prefix injected into the elevated child's argument string by the
; non-admin parent. The full string built becomes:
;    "-cli -outfile \"" <stdout-path> "\" -errfile \"" <stderr-path> "\" " <command>
str_relayPrefix     DCW '-','c','l','i',' ','-','o','u','t','f','i','l','e',' ',' ',0x22,0
str_relayMid        DCW ' ',0x22,' ','-','e','r','r','f','i','l','e',' ',' ',0x22,0
str_relayTail       DCW ' ',0x22,' ',0

; Interactive shells that cannot tolerate the relay path (CREATE_NO_WINDOW +
; redirected stdout to a temp file). When the user runs `cmdt -cli <shell>`
; with no further arguments we decline relay so the plain UAC self-elevate
; fallback gives them a real, attachable console window instead.
str_shell_cmd       DCW 'c','m','d',0
str_shell_cmd_exe   DCW 'c','m','d','.','e','x','e',0
str_shell_ps        DCW 'p','o','w','e','r','s','h','e','l','l',0
str_shell_ps_exe    DCW 'p','o','w','e','r','s','h','e','l','l','.','e','x','e',0
str_shell_pwsh      DCW 'p','w','s','h',0
str_shell_pwsh_exe  DCW 'p','w','s','h','.','e','x','e',0

; Null-terminated pointer table walked by the shell-guard loop in
; NonAdminRelayLaunch / AdminRelayLaunch. Add new shell names by inserting
; another pointer before the terminating zero.
shell_names_table
    DCQ str_shell_cmd, str_shell_cmd_exe
    DCQ str_shell_ps, str_shell_ps_exe
    DCQ str_shell_pwsh, str_shell_pwsh_exe
    DCQ 0

; ==============================================================================
; CODE SECTION
; ==============================================================================
    AREA |.text|, CODE, READONLY, ALIGN=4

    EXPORT NonAdminRelayLaunch
    EXPORT AdminRelayLaunch

; ==============================================================================
; NonAdminRelayLaunch - Run the non-admin -cli output-relay flow
;
; Parameters (ARM64 Windows ABI):
;   x0 (w0) = argc
;   x1      = argv pointer (LocalFree-owned, but we never free it; caller drops it)
;   x2      = raw command-line string from GetCommandLineW
;
; Returns:
;   x0 = 0 if relay declined (-new flag present, or temp setup failed).
;        Caller should fall back to plain UAC self-elevate.
;   Never returns once the elevated child has been spawned — every exit path
;   from that point on goes through ExitProcess.
;
; Register allocation (callee-saved):
;   x19 = raw cmdline
;   x20 = argv
;   x21 = relay-file read handle
;   x22 = std output/error handle
;   x23 = shell table pointer / misc
;   x24 = cmdline scan pointer
;   x25 = quote-state flag
;   w26 = argc stash
;
; Stack frame layout (post-prologue, all offsets sp-relative):
;   [sp+0..31]     shadow space for callee parameters
;   [sp+32..143]   SHELLEXECUTEINFOW (112 bytes)
;   [sp+144..151]  bytesRead temporary (ReadFile output)
;   [sp+152..159]  bytesWritten temporary (WriteFile output)
;   [sp+160..167]  relay/elevated-child exit status
;   [sp+168..223]  scratch / alignment padding
;
; Sizing: 5 callee-saved pair pushes + sub sp, #224 keeps every BL site
;         16-byte aligned (224 mod 16 = 0).
; ==============================================================================
NonAdminRelayLaunch PROC
    ; Prologue: save FP/LR and callee-saved registers
    STP x29, x30, [sp, #-16]!
    MOV x29, sp
    SUB sp, sp, #224
    STP x19, x20, [sp, #16]
    STP x21, x22, [sp, #32]
    STP x23, x24, [sp, #48]
    STP x25, x26, [sp, #64]
    STP x27, x28, [sp, #80]

    MOV x19, x2                     ; x19 = raw cmdline
    MOV x20, x1                     ; x20 = argv
    MOV w26, w0                     ; w26 = argc
    MOV w0, #1
    STR w0, [sp, #160]             ; [sp+160] = relay/elevated-child exit status

    ; If user requested -new, the spawned command must run in a visible new
    ; console window — that conflicts with output capture (which requires
    ; CREATE_NO_WINDOW). Decline so the caller falls back to plain UAC.
    CMP w26, #3
    B.LT narl_setup

    ; Check argv[2] for "-new"
    LDR x0, [x20, #16]             ; argv[2]
    ADRP x1, str_newSwitch
    ADD x1, x1, str_newSwitch
    BL wcscmp_ci
    CBNZ x0, narl_decline           ; Not "-new" → check outfile

    ; The elevated relay child is launched as:
    ;   cmdt -cli -outfile "<temp>" <command>
    ; It must execute the command and write to the temp file, not start a
    ; second relay cycle.
    LDR x0, [x20, #16]             ; argv[2]
    ADRP x1, str_outfileFlag
    ADD x1, x1, str_outfileFlag
    BL wcscmp_ci
    CBNZ x0, narl_decline           ; Not "-outfile" → continue

    ; Interactive-shell guard. The relay path uses CREATE_NO_WINDOW plus a
    ; temp-file capture of stdout — fundamentally incompatible with a shell
    ; that expects an attached console. When argv looks like exactly
    ; `cmdt -cli <shell>` (argc == 3, no extra tokens), bail out.
    CMP w26, #3
    B.NE narl_setup

    ADRP x23, shell_names_table
    ADD x23, x23, shell_names_table

narl_shell_loop
    LDR x1, [x23]                   ; x1 = shell name pointer
    CBZ x1, narl_setup              ; List exhausted, not an interactive shell
    LDR x0, [x20, #16]             ; argv[2]
    BL wcscmp_ci
    CBNZ x0, narl_decline           ; Match: interactive shell → fall back
    ADD x23, x23, #8
    B narl_shell_loop

narl_setup
    ; Get exe path for ShellExecuteExW.lpFile
    ; GetModuleFileNameW(hModule=NULL, lpFilename=g_exePath, nSize=260)
    MOV x0, XZR
    ADRP x1, g_exePath
    ADD x1, x1, g_exePath
    MOV w2, #260
    BL GetModuleFileNameW

    ; Get system temp directory
    MOV w0, #260
    ADRP x1, g_tempDirBuf
    ADD x1, x1, g_tempDirBuf
    BL GetTempPathW
    CBZ w0, narl_decline

    ; Create unique temp file name for stdout
    ADRP x0, g_tempDirBuf
    ADD x0, x0, g_tempDirBuf
    ADRP x1, str_cmdtPrefix
    ADD x1, x1, str_cmdtPrefix
    MOV w2, #0                      ; uUnique = 0 (use system time)
    ADRP x3, g_relayPath
    ADD x3, x3, g_relayPath
    BL GetTempFileNameW
    CBZ w0, narl_decline

    ; Create unique temp file name for stderr
    ADRP x0, g_tempDirBuf
    ADD x0, x0, g_tempDirBuf
    ADRP x1, str_cmdtPrefix
    ADD x1, x1, str_cmdtPrefix
    MOV w2, #0
    ADRP x3, g_relayErrPath
    ADD x3, x3, g_relayErrPath
    BL GetTempFileNameW
    CBZ w0, narl_delete_stdout_decline

    ; Build the modified argument string in g_relayArgs:
    ;   "-cli -outfile \"" + g_relayPath + "\" " + REST
    ADRP x0, g_relayArgs
    ADD x0, x0, g_relayArgs
    MOV w1, #0
    STRH w1, [x0]                   ; g_relayArgs[0] = 0

    ; wcscpy_p(g_relayArgs, str_relayPrefix)
    ADRP x0, g_relayArgs
    ADD x0, x0, g_relayArgs
    ADRP x1, str_relayPrefix
    ADD x1, x1, str_relayPrefix
    BL wcscpy_p

    ; wcscat_p(g_relayArgs, g_relayPath)
    ADRP x0, g_relayArgs
    ADD x0, x0, g_relayArgs
    ADRP x1, g_relayPath
    ADD x1, x1, g_relayPath
    BL wcscat_p

    ; wcscat_p(g_relayArgs, str_relayMid)
    ADRP x0, g_relayArgs
    ADD x0, x0, g_relayArgs
    ADRP x1, str_relayMid
    ADD x1, x1, str_relayMid
    BL wcscat_p

    ; wcscat_p(g_relayArgs, g_relayErrPath)
    ADRP x0, g_relayArgs
    ADD x0, x0, g_relayArgs
    ADRP x1, g_relayErrPath
    ADD x1, x1, g_relayErrPath
    BL wcscat_p

    ; wcscat_p(g_relayArgs, str_relayTail)
    ADRP x0, g_relayArgs
    ADD x0, x0, g_relayArgs
    ADRP x1, str_relayTail
    ADD x1, x1, str_relayTail
    BL wcscat_p

    ; Locate REST by walking the raw cmdline: skip exe path, then skip
    ; the "-cli" token, leaving x24 at the start of REST (or '\0').
    MOV x24, x19                    ; x24 = scan pointer
    MOV x25, XZR                    ; x25 = quote-state flag

narl_skip_exe
    LDRH w0, [x24]
    CBZ w0, narl_append_rest        ; End of string
    CMP w0, #'"'
    B.NE narl_check_space
    EOR x25, x25, #1               ; Toggle quote state

narl_check_space
    CMP w0, #' '
    B.NE narl_next_char
    CBNZ x25, narl_next_char        ; Inside quotes: space is part of path
    ADD x24, x24, #2
    MOV x0, x24
    BL skip_spaces
    MOV x24, x0
    B narl_skip_cli

narl_next_char
    ADD x24, x24, #2
    B narl_skip_exe

narl_skip_cli
    ; x24 points at "-cli". Walk to next space (or '\0').
    LDRH w0, [x24]
    CBZ w0, narl_append_rest
    CMP w0, #' '
    B.NE narl_cli_next
    ADD x24, x24, #2
    MOV x0, x24
    BL skip_spaces
    MOV x24, x0
    B narl_append_rest

narl_cli_next
    ADD x24, x24, #2
    B narl_skip_cli

narl_append_rest
    ; Append REST to g_relayArgs
    ADRP x0, g_relayArgs
    ADD x0, x0, g_relayArgs
    MOV x1, x24
    BL wcscat_p

    ; Zero SHELLEXECUTEINFOW at [sp+32] (112 bytes = 14 qwords)
    ADD x0, sp, #32
    MOV x1, XZR
    MOV w2, #14
narl_zero_sei
    STR x1, [x0], #8
    SUBS w2, w2, #1
    B.NE narl_zero_sei

    ; Fill SHELLEXECUTEINFOW. fMask = SEE_MASK_NOCLOSEPROCESS so we get
    ; back a process handle to wait on.
    MOV w0, #112
    STR w0, [sp, #32]              ; cbSize
    MOV w0, #SEE_MASK_NOCLOSEPROCESS
    STR w0, [sp, #36]              ; fMask

    ADRP x0, str_runas
    ADD x0, x0, str_runas
    STR x0, [sp, #48]              ; lpVerb (offset 16)

    ADRP x0, g_exePath
    ADD x0, x0, g_exePath
    STR x0, [sp, #56]              ; lpFile (offset 24)

    ADRP x0, g_relayArgs
    ADD x0, x0, g_relayArgs
    STR x0, [sp, #64]              ; lpParameters (offset 32)

    MOV w0, #SW_HIDE
    STR w0, [sp, #80]              ; nShow (offset 48) - elevated child
                                    ; has no console window of its own

    ; ShellExecuteExW(&sei)
    ADD x0, sp, #32
    BL ShellExecuteExW
    CBZ w0, narl_delete_only        ; UAC denied / cancelled

    ; Wait for elevated child to finish writing temp file.
    ; hProcess is at offset 104 in SHELLEXECUTEINFOW (last field).
    LDR x0, [sp, #136]             ; sp+32+104 = sp+136
    CBZ x0, narl_open_file
    MOV w1, #INFINITE
    BL WaitForSingleObject

    ; GetExitCodeProcess(hProcess, &exitCode)
    LDR x0, [sp, #136]
    ADD x1, sp, #160               ; &exitCode at [sp+160]
    BL GetExitCodeProcess           ; failure leaves conservative status 1

    ; CloseHandle(hProcess)
    LDR x0, [sp, #136]
    BL CloseHandle

narl_open_file
    ; CreateFileW(g_relayPath, GENERIC_READ, FILE_SHARE_READ, NULL,
    ;             OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, NULL)
    ; Stack args: dwCreationDisposition, dwFlagsAndAttributes, hTemplateFile
    SUB sp, sp, #48                 ; 32 shadow + 16 extra args
    MOV w0, #OPEN_EXISTING
    STR w0, [sp, #32]
    MOV w0, #FILE_ATTRIBUTE_NORMAL
    STR w0, [sp, #40]
    STR XZR, [sp, #48]             ; hTemplateFile = NULL

    ADRP x0, g_relayPath
    ADD x0, x0, g_relayPath
    MOV w1, #0x80000000             ; GENERIC_READ
    MOV w2, #1                      ; FILE_SHARE_READ
    MOV x3, XZR                     ; lpSecurityAttributes = NULL
    ; ARM64: args 5-7 go in x4-x6 (NOT the stack). Load the values that were
    ; staged on the stack above into their register homes before the call.
    LDR w4, [sp, #32]              ; dwCreationDisposition
    LDR w5, [sp, #40]              ; dwFlagsAndAttributes
    LDR x6, [sp, #48]             ; hTemplateFile
    BL CreateFileW
    ADD sp, sp, #48

    MOV x23, #0xFFFFFFFFFFFFFFFF
    CMP x0, x23
    B.EQ narl_delete_only           ; INVALID_HANDLE_VALUE

    MOV x21, x0                     ; x21 = relay-file read handle

    ; Get our STD_OUTPUT_HANDLE
    MOVN x0, #10                     ; STD_OUTPUT_HANDLE = -11
    BL GetStdHandle
    MOV x22, x0                     ; x22 = std output handle

narl_copy_loop
    ; ReadFile(x21, g_relayReadBuf, 4096, &[sp+144], NULL)
    MOV x0, x21
    ADRP x1, g_relayReadBuf
    ADD x1, x1, g_relayReadBuf
    MOV w2, #4096
    ADD x3, sp, #144               ; &bytesRead
    SUB sp, sp, #32                 ; shadow space for lpOverlapped
    STR XZR, [sp, #32]             ; lpOverlapped = NULL
    MOV x4, XZR                     ; ARM64: lpOverlapped is x4, not a stack arg
    BL ReadFile
    ADD sp, sp, #32
    CBZ w0, narl_close_file         ; ReadFile failed

    LDR w0, [sp, #144]             ; bytesRead
    CBZ w0, narl_close_file         ; EOF

    ; WriteFile(x22, g_relayReadBuf, bytesRead, &[sp+152], NULL)
    MOV x0, x22
    ADRP x1, g_relayReadBuf
    ADD x1, x1, g_relayReadBuf
    LDR w2, [sp, #144]             ; bytesRead (x0 was just reused as the handle)
    ADD x3, sp, #152               ; &bytesWritten
    SUB sp, sp, #32
    STR XZR, [sp, #32]             ; lpOverlapped = NULL
    MOV x4, XZR                     ; ARM64: lpOverlapped is x4, not a stack arg
    BL WriteFile
    ADD sp, sp, #32
    B narl_copy_loop

narl_close_file
    MOV x0, x21
    BL CloseHandle

    ; Replay stderr through this process's original STDERR handle
    SUB sp, sp, #48
    MOV w0, #OPEN_EXISTING
    STR w0, [sp, #32]
    MOV w0, #FILE_ATTRIBUTE_NORMAL
    STR w0, [sp, #40]
    STR XZR, [sp, #48]

    ADRP x0, g_relayErrPath
    ADD x0, x0, g_relayErrPath
    MOV w1, #0x80000000             ; GENERIC_READ
    MOV w2, #1                      ; FILE_SHARE_READ
    MOV x3, XZR
    ; ARM64: args 5-7 go in x4-x6 (NOT the stack). Load the values that were
    ; staged on the stack above into their register homes before the call.
    LDR w4, [sp, #32]              ; dwCreationDisposition
    LDR w5, [sp, #40]              ; dwFlagsAndAttributes
    LDR x6, [sp, #48]             ; hTemplateFile
    BL CreateFileW
    ADD sp, sp, #48

    MOV x23, #0xFFFFFFFFFFFFFFFF
    CMP x0, x23
    B.EQ narl_delete_only

    MOV x21, x0                     ; x21 = stderr relay handle

    ; Get STDERR handle
    MOVN x0, #11                     ; STD_ERROR_HANDLE = -12
    BL GetStdHandle
    MOV x22, x0

narl_err_copy_loop
    MOV x0, x21
    ADRP x1, g_relayReadBuf
    ADD x1, x1, g_relayReadBuf
    MOV w2, #4096
    ADD x3, sp, #144
    SUB sp, sp, #32
    STR XZR, [sp, #32]
    MOV x4, XZR                     ; ARM64: lpOverlapped is x4, not a stack arg
    BL ReadFile
    ADD sp, sp, #32
    CBZ w0, narl_err_close

    LDR w0, [sp, #144]
    CBZ w0, narl_err_close

    MOV x0, x22
    ADRP x1, g_relayReadBuf
    ADD x1, x1, g_relayReadBuf
    LDR w2, [sp, #144]             ; bytesRead (x0 was just reused as the handle)
    ADD x3, sp, #152
    SUB sp, sp, #32
    STR XZR, [sp, #32]
    MOV x4, XZR                     ; ARM64: lpOverlapped is x4, not a stack arg
    BL WriteFile
    ADD sp, sp, #32
    B narl_err_copy_loop

narl_err_close
    MOV x0, x21
    BL CloseHandle

narl_delete_only
    ; Delete the temp files (best effort) and exit the process
    ADRP x0, g_relayPath
    ADD x0, x0, g_relayPath
    BL DeleteFileW

    ADRP x0, g_relayErrPath
    ADD x0, x0, g_relayErrPath
    BL DeleteFileW

    ; Exit with relay status
    LDR w0, [sp, #160]
    BL ExitProcess

narl_decline
    ; Bail out without touching anything else. Caller falls back to plain UAC.
    MOV x0, XZR                     ; return 0
    B narl_exit

narl_delete_stdout_decline
    ADRP x0, g_relayPath
    ADD x0, x0, g_relayPath
    BL DeleteFileW
    B narl_decline

narl_exit
    ; Epilogue
    LDP x27, x28, [sp, #80]
    LDP x25, x26, [sp, #64]
    LDP x23, x24, [sp, #48]
    LDP x21, x22, [sp, #32]
    LDP x19, x20, [sp, #16]
    ADD sp, sp, #224
    LDP x29, x30, [sp], #16
    RET
NonAdminRelayLaunch ENDP

; ==============================================================================
; AdminRelayLaunch - Output-relay for -cli when the caller is ALREADY elevated
;
; Same problem as NonAdminRelayLaunch, but the caller here is already
; Administrator, so re-elevating through ShellExecuteExW("runas") would be
; wrong. Instead this captures output entirely in-process: open a local temp
; file, set g_relayHandle, call RunAsTrustedInstaller directly (token
; duplication only, no shell elevation, no second prompt), then stream the
; temp file to our own STD_OUTPUT_HANDLE.
;
; Parameters:
;   x0 (w0) = argc
;   x1      = argv pointer
;   x2      = raw command-line string from GetCommandLineW
;
; Returns:
;   x0 = 0 if declined (-new flag present, or an interactive shell target).
;        Caller should fall through to normal admin dispatch.
;   Never returns once the local run has started.
; ==============================================================================
AdminRelayLaunch PROC
    STP x29, x30, [sp, #-16]!
    MOV x29, sp
    SUB sp, sp, #224
    STP x19, x20, [sp, #16]
    STP x21, x22, [sp, #32]
    STP x23, x24, [sp, #48]
    STP x25, x26, [sp, #64]
    STP x27, x28, [sp, #80]

    MOV x19, x2                     ; x19 = raw cmdline
    MOV x20, x1                     ; x20 = argv
    MOV w26, w0                     ; w26 = argc

    ; -new conflicts with output capture (needs CREATE_NO_WINDOW) -- decline
    CMP w26, #3
    B.LT arl_setup

    LDR x0, [x20, #16]             ; argv[2]
    ADRP x1, str_newSwitch
    ADD x1, x1, str_newSwitch
    BL wcscmp_ci
    CBNZ x0, arl_decline

    ; If the non-admin relay's elevated child lands here, argv[2] is the
    ; internal "-outfile" token, not a real command. Decline so it falls
    ; through to the normal dispatch.
    LDR x0, [x20, #16]
    ADRP x1, str_outfileFlag
    ADD x1, x1, str_outfileFlag
    BL wcscmp_ci
    CBNZ x0, arl_decline

    ; Interactive-shell guard
    CMP w26, #3
    B.NE arl_setup

    ADRP x23, shell_names_table
    ADD x23, x23, shell_names_table

arl_shell_loop
    LDR x1, [x23]
    CBZ x1, arl_setup
    LDR x0, [x20, #16]
    BL wcscmp_ci
    CBNZ x0, arl_decline
    ADD x23, x23, #8
    B arl_shell_loop

arl_setup
    ; Temp file to capture the child's stdout/stderr
    MOV w0, #260
    ADRP x1, g_tempDirBuf
    ADD x1, x1, g_tempDirBuf
    BL GetTempPathW
    CBZ w0, arl_decline

    ADRP x0, g_tempDirBuf
    ADD x0, x0, g_tempDirBuf
    ADRP x1, str_cmdtPrefix
    ADD x1, x1, str_cmdtPrefix
    MOV w2, #0
    ADRP x3, g_relayPath
    ADD x3, x3, g_relayPath
    BL GetTempFileNameW
    CBZ w0, arl_decline

    ADRP x0, g_tempDirBuf
    ADD x0, x0, g_tempDirBuf
    ADRP x1, str_cmdtPrefix
    ADD x1, x1, str_cmdtPrefix
    MOV w2, #0
    ADRP x3, g_relayErrPath
    ADD x3, x3, g_relayErrPath
    BL GetTempFileNameW
    CBZ w0, arl_delete_stdout_decline

    ; Open stdout relay file for inheritable write access
    ; SECURITY_ATTRIBUTES at [sp+96] (24 bytes)
    MOV w0, #24
    STR w0, [sp, #96]              ; nLength
    MOV w0, #0
    STR w0, [sp, #100]             ; padding
    STR XZR, [sp, #104]            ; lpSecurityDescriptor = NULL
    MOV w0, #1
    STR w0, [sp, #112]             ; bInheritHandle = TRUE
    MOV w0, #0
    STR w0, [sp, #116]             ; padding

    ; CreateFileW(g_relayPath, GENERIC_WRITE, FILE_SHARE_READ, &sa,
    ;             CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL)
    SUB sp, sp, #48
    MOV w0, #CREATE_ALWAYS
    STR w0, [sp, #32]
    MOV w0, #FILE_ATTRIBUTE_NORMAL
    STR w0, [sp, #40]
    STR XZR, [sp, #48]

    ADRP x0, g_relayPath
    ADD x0, x0, g_relayPath
    MOV w1, #0x40000000             ; GENERIC_WRITE
    MOV w2, #1                      ; FILE_SHARE_READ
    ADD x3, sp, #96                ; &SECURITY_ATTRIBUTES (before sub)
    ADD x3, x3, #48                ; Adjust for sub sp, #48
    ; ARM64: args 5-7 go in x4-x6 (NOT the stack). Load the values that were
    ; staged on the stack above into their register homes before the call.
    LDR w4, [sp, #32]              ; dwCreationDisposition
    LDR w5, [sp, #40]              ; dwFlagsAndAttributes
    LDR x6, [sp, #48]             ; hTemplateFile
    BL CreateFileW
    ADD sp, sp, #48

    MOV x23, #0xFFFFFFFFFFFFFFFF
    CMP x0, x23
    B.EQ arl_decline

    ADRP x1, g_relayHandle
    ADD x1, x1, g_relayHandle
    STR x0, [x1]                    ; g_relayHandle = handle

    ; CreateFileW for stderr relay
    SUB sp, sp, #48
    MOV w0, #CREATE_ALWAYS
    STR w0, [sp, #32]
    MOV w0, #FILE_ATTRIBUTE_NORMAL
    STR w0, [sp, #40]
    STR XZR, [sp, #48]

    ADRP x0, g_relayErrPath
    ADD x0, x0, g_relayErrPath
    MOV w1, #0x40000000
    MOV w2, #1
    ADD x3, sp, #96
    ADD x3, x3, #48
    ; ARM64: args 5-7 go in x4-x6 (NOT the stack). Load the values that were
    ; staged on the stack above into their register homes before the call.
    LDR w4, [sp, #32]              ; dwCreationDisposition
    LDR w5, [sp, #40]              ; dwFlagsAndAttributes
    LDR x6, [sp, #48]             ; hTemplateFile
    BL CreateFileW
    ADD sp, sp, #48

    MOV x23, #0xFFFFFFFFFFFFFFFF
    CMP x0, x23
    B.EQ arl_close_stdout_decline

    ADRP x1, g_relayErrHandle
    ADD x1, x1, g_relayErrHandle
    STR x0, [x1]

    ; Locate REST (everything after "-cli") by walking the raw cmdline
    MOV x24, x19
    MOV x25, XZR

arl_skip_exe
    LDRH w0, [x24]
    CBZ w0, arl_run
    CMP w0, #'"'
    B.NE arl_check_space
    EOR x25, x25, #1

arl_check_space
    CMP w0, #' '
    B.NE arl_next_char
    CBNZ x25, arl_next_char
    ADD x24, x24, #2
    MOV x0, x24
    BL skip_spaces
    MOV x24, x0
    B arl_skip_cli

arl_next_char
    ADD x24, x24, #2
    B arl_skip_exe

arl_skip_cli
    LDRH w0, [x24]
    CBZ w0, arl_run
    CMP w0, #' '
    B.NE arl_cli_next
    ADD x24, x24, #2
    MOV x0, x24
    BL skip_spaces
    MOV x24, x0
    B arl_run

arl_cli_next
    ADD x24, x24, #2
    B arl_skip_cli

arl_run
    ; Copy REST into g_cmdBuf
    ADRP x0, g_cmdBuf
    ADD x0, x0, g_cmdBuf
    MOV w1, #0
    STRH w1, [x0]                   ; g_cmdBuf[0] = 0

    ADRP x0, g_cmdBuf
    ADD x0, x0, g_cmdBuf
    MOV x1, x24
    BL wcscpy_p

    ; RunAsTrustedInstaller sees g_relayHandle != 0 and runs relay Mode 3:
    ; CREATE_NO_WINDOW, stdout/stderr -> our temp file, waits for the child.
    ADRP x0, g_cmdBuf
    ADD x0, x0, g_cmdBuf
    MOV x1, XZR                     ; useNewConsole = 0
    BL RunAsTrustedInstaller
    MOV w23, w0                     ; w23 = BOOL result

    ; Close the write handle so the file is flushed before we read it back
    ADRP x0, g_relayHandle
    ADD x0, x0, g_relayHandle
    LDR x0, [x0]
    BL CloseHandle
    ADRP x0, g_relayHandle
    ADD x0, x0, g_relayHandle
    STR XZR, [x0]

    ADRP x0, g_relayErrHandle
    ADD x0, x0, g_relayErrHandle
    LDR x0, [x0]
    BL CloseHandle
    ADRP x0, g_relayErrHandle
    ADD x0, x0, g_relayErrHandle
    STR XZR, [x0]

    ; Re-open for read and stream to our own STD_OUTPUT_HANDLE
    SUB sp, sp, #48
    MOV w0, #OPEN_EXISTING
    STR w0, [sp, #32]
    MOV w0, #FILE_ATTRIBUTE_NORMAL
    STR w0, [sp, #40]
    STR XZR, [sp, #48]

    ADRP x0, g_relayPath
    ADD x0, x0, g_relayPath
    MOV w1, #0x80000000
    MOV w2, #1
    MOV x3, XZR
    ; ARM64: args 5-7 go in x4-x6 (NOT the stack). Load the values that were
    ; staged on the stack above into their register homes before the call.
    LDR w4, [sp, #32]              ; dwCreationDisposition
    LDR w5, [sp, #40]              ; dwFlagsAndAttributes
    LDR x6, [sp, #48]             ; hTemplateFile
    BL CreateFileW
    ADD sp, sp, #48

    MOV x24, #0xFFFFFFFFFFFFFFFF
    CMP x0, x24
    B.EQ arl_delete_only

    MOV x21, x0                     ; x21 = relay-file read handle

    MOVN x0, #10                     ; STD_OUTPUT_HANDLE
    BL GetStdHandle
    MOV x22, x0

arl_copy_loop
    MOV x0, x21
    ADRP x1, g_relayReadBuf
    ADD x1, x1, g_relayReadBuf
    MOV w2, #4096
    ADD x3, sp, #144
    SUB sp, sp, #32
    STR XZR, [sp, #32]
    MOV x4, XZR                     ; ARM64: lpOverlapped is x4, not a stack arg
    BL ReadFile
    ADD sp, sp, #32
    CBZ w0, arl_close_file

    LDR w0, [sp, #144]
    CBZ w0, arl_close_file

    MOV x0, x22
    ADRP x1, g_relayReadBuf
    ADD x1, x1, g_relayReadBuf
    LDR w2, [sp, #144]             ; bytesRead (x0 was just reused as the handle)
    ADD x3, sp, #152
    SUB sp, sp, #32
    STR XZR, [sp, #32]
    MOV x4, XZR                     ; ARM64: lpOverlapped is x4, not a stack arg
    BL WriteFile
    ADD sp, sp, #32
    B arl_copy_loop

arl_close_file
    MOV x0, x21
    BL CloseHandle

    ; Stream stderr relay
    SUB sp, sp, #48
    MOV w0, #OPEN_EXISTING
    STR w0, [sp, #32]
    MOV w0, #FILE_ATTRIBUTE_NORMAL
    STR w0, [sp, #40]
    STR XZR, [sp, #48]

    ADRP x0, g_relayErrPath
    ADD x0, x0, g_relayErrPath
    MOV w1, #0x80000000
    MOV w2, #1
    MOV x3, XZR
    ; ARM64: args 5-7 go in x4-x6 (NOT the stack). Load the values that were
    ; staged on the stack above into their register homes before the call.
    LDR w4, [sp, #32]              ; dwCreationDisposition
    LDR w5, [sp, #40]              ; dwFlagsAndAttributes
    LDR x6, [sp, #48]             ; hTemplateFile
    BL CreateFileW
    ADD sp, sp, #48

    MOV x24, #0xFFFFFFFFFFFFFFFF
    CMP x0, x24
    B.EQ arl_delete_only

    MOV x21, x0

    MOVN x0, #11                     ; STD_ERROR_HANDLE
    BL GetStdHandle
    MOV x22, x0

arl_err_copy_loop
    MOV x0, x21
    ADRP x1, g_relayReadBuf
    ADD x1, x1, g_relayReadBuf
    MOV w2, #4096
    ADD x3, sp, #144
    SUB sp, sp, #32
    STR XZR, [sp, #32]
    MOV x4, XZR                     ; ARM64: lpOverlapped is x4, not a stack arg
    BL ReadFile
    ADD sp, sp, #32
    CBZ w0, arl_err_close

    LDR w0, [sp, #144]
    CBZ w0, arl_err_close

    MOV x0, x22
    ADRP x1, g_relayReadBuf
    ADD x1, x1, g_relayReadBuf
    LDR w2, [sp, #144]             ; bytesRead (x0 was just reused as the handle)
    ADD x3, sp, #152
    SUB sp, sp, #32
    STR XZR, [sp, #32]
    MOV x4, XZR                     ; ARM64: lpOverlapped is x4, not a stack arg
    BL WriteFile
    ADD sp, sp, #32
    B arl_err_copy_loop

arl_err_close
    MOV x0, x21
    BL CloseHandle

arl_delete_only
    ADRP x0, g_relayPath
    ADD x0, x0, g_relayPath
    BL DeleteFileW

    ADRP x0, g_relayErrPath
    ADD x0, x0, g_relayErrPath
    BL DeleteFileW

    ; Exit with appropriate code
    CBZ w23, arl_exit_fail
    ADRP x0, g_childExitCode
    ADD x0, x0, g_childExitCode
    LDR w0, [x0]
    B arl_exit_code

arl_exit_fail
    MOV w0, #1

arl_exit_code
    BL ExitProcess

arl_decline
    MOV x0, XZR                     ; return 0
    B arl_exit

arl_close_stdout_decline
    ADRP x0, g_relayHandle
    ADD x0, x0, g_relayHandle
    LDR x0, [x0]
    BL CloseHandle
    ADRP x0, g_relayHandle
    ADD x0, x0, g_relayHandle
    STR XZR, [x0]

arl_delete_stdout_decline
    ADRP x0, g_relayPath
    ADD x0, x0, g_relayPath
    BL DeleteFileW
    B arl_decline

arl_exit
    LDP x27, x28, [sp, #80]
    LDP x25, x26, [sp, #64]
    LDP x23, x24, [sp, #48]
    LDP x21, x22, [sp, #32]
    LDP x19, x20, [sp, #16]
    ADD sp, sp, #224
    LDP x29, x30, [sp], #16
    RET
AdminRelayLaunch ENDP

    END