; ==============================================================================
; CMDT - Run as TrustedInstaller
; Main Entry Point and Command-Line Processing Module
; 
; Author: Marek Wesołowski (wesmar)
; Purpose: Provides the entry point and command-line argument parsing for the
;          TrustedInstaller privilege elevation utility.
; ==============================================================================

option casemap:none

include consts.inc

; External function declarations
EXTRN CreateMainWindow:PROC
EXTRN RunAsTrustedInstaller:PROC
EXTRN ResolveLnkPath:PROC
EXTRN GetModuleHandleW:PROC
EXTRN GetMessageW:PROC
EXTRN TranslateMessage:PROC
EXTRN DispatchMessageW:PROC
EXTRN IsDialogMessageW:PROC
EXTRN ExitProcess:PROC
EXTRN GetCommandLineW:PROC
EXTRN CommandLineToArgvW:PROC
EXTRN LocalFree:PROC
EXTRN SetFocus:PROC
EXTRN IsUserAnAdmin:PROC
EXTRN ShellExecuteExW:PROC
EXTRN GetModuleFileNameW:PROC
EXTRN AttachConsole:PROC
EXTRN GetStdHandle:PROC
EXTRN GetFileType:PROC
EXTRN WaitForSingleObject:PROC
EXTRN CloseHandle:PROC
EXTRN CreateFileW:PROC
EXTRN ReadFile:PROC
EXTRN WriteFile:PROC
EXTRN DeleteFileW:PROC
EXTRN GetTempPathW:PROC
EXTRN GetTempFileNameW:PROC

; String helpers (defined in strutil.asm)
EXTRN DecryptWideStr:PROC
EXTRN wcscpy_p:PROC
EXTRN wcscat_p:PROC
EXTRN wcscmp_ci:PROC
EXTRN wcscmp_token:PROC
EXTRN skip_spaces:PROC
EXTRN wcslen_p:PROC

; Help / usage display (defined in help.asm)
EXTRN HelpCheckAndExit:PROC
EXTRN ShowUsage:PROC

; Installation / hook management (defined in install.asm)
EXTRN InstallContextMenu:PROC
EXTRN UninstallContextMenu:PROC
EXTRN InstallShift:PROC
EXTRN UninstallShift:PROC

; Non-admin output relay (defined in relay.asm)
EXTRN NonAdminRelayLaunch:PROC

; CLI / file-run dispatch (defined in cli.asm)
EXTRN mode_cli_found:PROC
EXTRN mode_file_run:PROC

; mode_gui is an in-proc label inside mainCRTStartup; cli.asm jumps into it
; when no file argument is supplied. We declare it using the `::` (double
; colon) syntax at the label definition to make it externally visible.

; ==============================================================================
; CONSTANT STRING DATA
; ==============================================================================
.const

; Strings shared with relay.asm and cli.asm (made PUBLIC so cross-module
; references resolve at link time). Help-switch and install-switch strings
; are dispatcher-only and stay private to this module.
PUBLIC str_runas, str_newSwitch, str_extLnk_m, str_space, str_outfileFlag

; Command-line switch for CLI mode
str_cliSwitch1  dw '-','c','l','i',0           ; CLI mode switch

; Internal flag passed by non-admin parent to elevated child to redirect
; child's stdout/stderr to a temporary file for relaying back to the
; original shell's stdout (handles UAC handle-inheritance limitation).
str_outfileFlag dw '-','o','u','t','f','i','l','e',0

; Switch to request new console window
str_newSwitch   dw '-','n','e','w',0

; File extension for Windows shortcuts
str_extLnk_m    dw '.','l','n','k',0

; Space character for string concatenation
str_space       dw ' ',0

; UAC self-elevation verb
str_runas           dw 'r','u','n','a','s',0

; Context menu registration switches
str_installSwitch   dw '-','i','n','s','t','a','l','l',0
str_uninstallSwitch dw '-','u','n','i','n','s','t','a','l','l',0

; Sticky Keys (sethc.exe) IFEO switches
str_shiftSwitch     dw '-','s','h','i','f','t',0
str_unshiftSwitch   dw '-','u','n','s','h','i','f','t',0


; ==============================================================================
; PRIVILEGE NAME STRINGS
; These strings are used to construct full privilege names by combining with
; the "Se" prefix and "Privilege" suffix (e.g., "SeDebugPrivilege")
; ==============================================================================
privStr_0  dw 'A','s','s','i','g','n','P','r','i','m','a','r','y','T','o','k','e','n',0
privStr_1  dw 'B','a','c','k','u','p',0
privStr_2  dw 'R','e','s','t','o','r','e',0
privStr_3  dw 'D','e','b','u','g',0
privStr_4  dw 'I','m','p','e','r','s','o','n','a','t','e',0
privStr_5  dw 'T','a','k','e','O','w','n','e','r','s','h','i','p',0
privStr_6  dw 'L','o','a','d','D','r','i','v','e','r',0
privStr_7  dw 'S','y','s','t','e','m','E','n','v','i','r','o','n','m','e','n','t',0
privStr_8  dw 'M','a','n','a','g','e','V','o','l','u','m','e',0
privStr_9  dw 'S','e','c','u','r','i','t','y',0
privStr_10 dw 'S','h','u','t','d','o','w','n',0
privStr_11 dw 'S','y','s','t','e','m','t','i','m','e',0
privStr_12 dw 'T','c','b',0
privStr_13 dw 'I','n','c','r','e','a','s','e','Q','u','o','t','a',0
privStr_14 dw 'A','u','d','i','t',0
privStr_15 dw 'C','h','a','n','g','e','N','o','t','i','f','y',0
privStr_16 dw 'U','n','d','o','c','k',0
privStr_17 dw 'C','r','e','a','t','e','T','o','k','e','n',0
privStr_18 dw 'L','o','c','k','M','e','m','o','r','y',0
privStr_19 dw 'C','r','e','a','t','e','P','a','g','e','f','i','l','e',0
privStr_20 dw 'C','r','e','a','t','e','P','e','r','m','a','n','e','n','t',0
privStr_21 dw 'S','y','s','t','e','m','P','r','o','f','i','l','e',0
privStr_22 dw 'P','r','o','f','i','l','e','S','i','n','g','l','e','P','r','o','c','e','s','s',0
privStr_23 dw 'C','r','e','a','t','e','G','l','o','b','a','l',0
privStr_24 dw 'T','i','m','e','Z','o','n','e',0
privStr_25 dw 'C','r','e','a','t','e','S','y','m','b','o','l','i','c','L','i','n','k',0
privStr_26 dw 'I','n','c','r','e','a','s','e','B','a','s','e','P','r','i','o','r','i','t','y',0
privStr_27 dw 'R','e','m','o','t','e','S','h','u','t','d','o','w','n',0
privStr_28 dw 'I','n','c','r','e','a','s','e','W','o','r','k','i','n','g','S','e','t',0
privStr_29 dw 'R','e','l','a','b','e','l',0
privStr_30 dw 'D','e','l','e','g','a','t','e','S','e','s','s','i','o','n','U','s','e','r','I','m','p','e','r','s','o','n','a','t','e',0
privStr_31 dw 'T','r','u','s','t','e','d','C','r','e','d','M','a','n','A','c','c','e','s','s',0
privStr_32 dw 'E','n','a','b','l','e','D','e','l','e','g','a','t','i','o','n',0
privStr_33 dw 'S','y','n','c','A','g','e','n','t',0

; Prefix and suffix for privilege name construction
privPrefix dw 'S','e',0
privSuffix dw 'P','r','i','v','i','l','e','g','e',0

; ==============================================================================
; INITIALIZED DATA SECTION
; ==============================================================================
.data
    align 8

; Privilege table: Array of pointers to privilege name strings
; Used by token manipulation code to enable all available privileges
PUBLIC g_privTable
g_privTable dq offset privStr_0,offset privStr_1,offset privStr_2,offset privStr_3,offset privStr_4,offset privStr_5
            dq offset privStr_6,offset privStr_7,offset privStr_8,offset privStr_9,offset privStr_10,offset privStr_11
            dq offset privStr_12,offset privStr_13,offset privStr_14,offset privStr_15,offset privStr_16,offset privStr_17
            dq offset privStr_18,offset privStr_19,offset privStr_20,offset privStr_21,offset privStr_22,offset privStr_23
            dq offset privStr_24,offset privStr_25,offset privStr_26,offset privStr_27,offset privStr_28,offset privStr_29
            dq offset privStr_30,offset privStr_31,offset privStr_32,offset privStr_33

; Global variables exported to other modules
PUBLIC g_cachedToken, g_tokenTime, g_hwndMain, g_hwndEdit, g_hwndBtn, g_hwndStatus, g_hConsoleOut, g_hInstance
PUBLIC g_useNewConsole
PUBLIC privPrefix, privSuffix

; Cached TrustedInstaller token handle (for performance optimization)
g_cachedToken   dq 0

; Timestamp of when the cached token was obtained (in milliseconds)
g_tokenTime     dd 0
                dd 0            ; Padding for alignment

; Window handles for GUI controls
g_hwndMain      dq 0            ; Main window handle
g_hwndEdit      dq 0            ; ComboBox edit control handle
g_hwndBtn       dq 0            ; Run button handle
g_hwndStatus    dq 0            ; Status label handle

; Console output handle (for CLI mode)
g_hConsoleOut   dq 0

; Flag indicating whether to create a new console window
g_useNewConsole dd 0
                dd 0            ; Padding for alignment

; Application instance handle
g_hInstance     dq 0

; Handle to output-relay file (set in elevated child when -outfile flag is
; present on the command line). Zero means no relay.
PUBLIC g_relayHandle
g_relayHandle   dq 0

; ==============================================================================
; UNINITIALIZED DATA SECTION
; ==============================================================================
.data?

; Export buffer declarations for use by other modules
PUBLIC g_cmdBuf, g_statusBuf, g_filePath, g_argsBuf, g_tempBuf, g_exePath, g_decryptBuf
PUBLIC g_relayPath, g_tempDirBuf, g_relayArgs, g_relayReadBuf

; Buffer for command line text (520 WCHARs = 1040 bytes)
g_cmdBuf        dw 520 dup(?)

; Buffer for status text (520 WCHARs)
g_statusBuf     dw 520 dup(?)

; Buffer for file paths (520 WCHARs)
g_filePath      dw 520 dup(?)

; Buffer for command-line arguments (520 WCHARs)
g_argsBuf       dw 520 dup(?)

; Temporary buffer for various operations (1040 WCHARs = 2080 bytes)
g_tempBuf       dw 1040 dup(?)

; Buffer for exe path (UAC elevation and context menu registration)
g_exePath       dw 260 dup(?)

; Buffer for decrypted strings (reusable, 520 WCHARs)
g_decryptBuf    dw 520 dup(?)

; Temp directory path returned by GetTempPathW (MAX_PATH WCHARs)
g_tempDirBuf    dw 260 dup(?)

; Relay temp file full path produced by GetTempFileNameW (MAX_PATH WCHARs)
g_relayPath     dw 260 dup(?)

; Buffer used to build the modified arguments string passed to the elevated
; child via ShellExecuteExW (`-cli -outfile <path> <rest>`)
g_relayArgs     dw 1040 dup(?)

; Scratch buffer used to stream temp-file bytes back to original stdout
g_relayReadBuf  db 4096 dup(?)

; ==============================================================================
; CODE SECTION
; ==============================================================================
.code


; ==============================================================================
; mainCRTStartup - Application Entry Point
;
; Purpose: Parses command-line arguments and routes to GUI, CLI, install,
;          or self-elevation paths. Also handles the output-relay path that
;          delivers child stdout/stderr back to a non-admin shell across
;          the UAC handle-inheritance boundary.
;
; Command-line syntax (user-facing):
;   - GUI mode:               <exe> (no arguments)
;   - CLI mode:               <exe> -cli <command>
;   - CLI new console:        <exe> -cli -new <command>
;   - Show usage:             <exe> {-help|-h|--help|-?|/?|/h|/help}
;   - File run (right click): <exe> "C:\path\to\file.exe"
;   - Context menu install:   <exe> -install | -uninstall
;   - Sticky-keys hook:       <exe> -shift   | -unshift
;
; Internal switch (set automatically by the non-admin parent when relaying):
;   - <exe> -cli -outfile "<path>" <command>
;     The elevated child opens <path> with inheritable write access and
;     uses it as the spawned process's stdout/stderr.
;
; Control flow:
;   1. Parse argv early so help and relay decisions can run before UAC.
;   2. AttachConsole(ATTACH_PARENT_PROCESS) — best-effort, lets us reach
;      the parent shell's std handles or redirect targets.
;   3. IsUserAnAdmin:
;        - Yes: admin_dispatch → uac_already_admin → dispatch switches.
;        - No, argv[1] == "-cli" (and not -new): nonadmin_relay.
;        - No, otherwise: nonadmin_plain (UAC self-elevate, no relay).
;
; Returns: Does not return (every leaf path calls ExitProcess).
;
; Stack frame: 312 bytes local variables, including:
;   - [rbp-64]   argc
;   - [rbp-72]   ReadFile/WriteFile byte-count temporaries (relay path)
;   - [rbp-80]   …or INPUT_RECORD field (show_usage path)
;   - [rbp-96]   INPUT_RECORD base (show_usage path)
;   - [rbp-104]  SECURITY_ATTRIBUTES (mode_cli_setup -outfile parsing)
;   - [rbp-312]  SHELLEXECUTEINFOW (UAC self-elevate / relay launch)
; Each code path uses a disjoint subset, so the overlaps above are safe.
; ==============================================================================
mainCRTStartup proc frame
    push rbp
    .pushreg rbp
    mov rbp, rsp
    .setframe rbp, 0
    push rbx
    .pushreg rbx
    push rsi
    .pushreg rsi
    push rdi
    .pushreg rdi
    push r12
    .pushreg r12
    push r13
    .pushreg r13
    push r14
    .pushreg r14
    push r15
    .pushreg r15
    sub rsp, 312
    .allocstack 312
    .endprolog

    cld                         ; Clear direction flag for string operations

    ; ===== Attach to parent shell's console (best effort) =====
    ; Redirection and pipes must keep their inherited handles intact. A real
    ; console handle is different: the GUI-subsystem parent may still not be
    ; attached to that console, and then a child console process can flash in
    ; its own window instead of writing inline. Attach for CHAR/invalid stdout,
    ; but leave FILE/PIPE stdout alone.
    mov ecx, STD_OUTPUT_HANDLE
    sub rsp, 32
    call GetStdHandle
    add rsp, 32
    test rax, rax
    jz early_do_attach
    cmp rax, -1
    je early_do_attach
    mov rcx, rax
    sub rsp, 32
    call GetFileType
    add rsp, 32
    cmp eax, 2                  ; FILE_TYPE_CHAR = console/character device
    je early_do_attach
    jmp early_attach_done

early_do_attach:
    mov ecx, ATTACH_PARENT_PROCESS
    sub rsp, 32
    call AttachConsole
    add rsp, 32

early_attach_done:

    ; ===== Early help-switch check (runs BEFORE UAC self-elevation) =====
    ; Reason: UAC self-elevation spawns a new process detached from the
    ; original shell, so stdout redirect (`cmdt -help >out.txt`) and
    ; AttachConsole(ATTACH_PARENT_PROCESS) cannot reach the original cmd.
    ; By printing usage from the non-elevated process we stay attached to
    ; the launching console.
    sub rsp, 32
    call GetCommandLineW
    add rsp, 32
    mov r12, rax

    lea rdx, [rbp-64]
    mov rcx, r12
    sub rsp, 32
    call CommandLineToArgvW
    add rsp, 32
    mov r13, rax
    test rax, rax
    jz early_help_skip

    ; Delegate help-switch detection to help.asm. If a help variant is
    ; present, HelpCheckAndExit frees argv via LocalFree and never returns
    ; (it tail-calls ShowUsage -> ExitProcess). Otherwise it returns and
    ; we continue with normal dispatch.
    mov ecx, dword ptr [rbp-64] ; argc
    mov rdx, r13                ; argv
    sub rsp, 32
    call HelpCheckAndExit
    add rsp, 32

early_help_skip:
    ; ===== Admin check =====
    sub rsp, 32
    call IsUserAnAdmin
    add rsp, 32
    test eax, eax
    jnz admin_dispatch

    ; ===== Non-admin: decide between relay and plain UAC self-elevate =====
    ; Relay is required when the user runs `cmdt -cli ...` from a non-admin
    ; shell because UAC starts the elevated child in a brand new process
    ; tree with no handle inheritance. To deliver the child's output back
    ; to the original shell (or its redirect), we:
    ;   1. Build a temp file path.
    ;   2. Insert "-outfile <path>" after "-cli" in the args.
    ;   3. Spawn elevated child, wait for it.
    ;   4. Stream the temp file to our own stdout (which is inherited from
    ;      cmd, so redirects work transparently).
    ; For other switches (-install, -shift, file paths, etc.) the user
    ; doesn't expect captured output, so we keep the original fast path.
    test r13, r13
    jz nonadmin_plain
    mov eax, dword ptr [rbp-64]
    cmp eax, 2
    jl nonadmin_plain

    mov r14, [r13+8]            ; argv[1]
    lea rdx, str_cliSwitch1
    mov rcx, r14
    call wcscmp_ci
    test rax, rax
    jz nonadmin_plain           ; argv[1] != "-cli" → plain UAC

    ; argv[1] == "-cli" — try the relay path. It returns 0 if it declines
    ; (e.g. -new flag clashes with output capture); otherwise it never
    ; returns (every success/failure past that point calls ExitProcess).
    ; On a 0 return we fall through into nonadmin_plain below.
    mov ecx, dword ptr [rbp-64] ; argc
    mov rdx, r13                ; argv
    mov r8, r12                 ; cmdline
    sub rsp, 32
    call NonAdminRelayLaunch
    add rsp, 32

nonadmin_plain:
    ; Free early-parse argv before plain UAC self-elevate.
    test r13, r13
    jz nonadmin_plain_setup
    mov rcx, r13
    sub rsp, 32
    call LocalFree
    add rsp, 32
    xor r13, r13

nonadmin_plain_setup:
    ; ===== Plain UAC self-elevate (original behavior, no relay) =====
    lea rdx, g_exePath
    mov r8d, 260
    xor ecx, ecx
    sub rsp, 32
    call GetModuleFileNameW
    add rsp, 32

    ; Extract arguments from command line (skip exe path)
    sub rsp, 32
    call GetCommandLineW
    add rsp, 32
    mov rsi, rax
    xor edi, edi
uac_skip_exe:
    mov ax, word ptr [rsi]
    test ax, ax
    jz uac_no_args
    cmp ax, '"'
    jne @F
    xor edi, 1
@@:
    cmp ax, ' '
    jne @F
    test edi, edi
    jnz @F
    mov rcx, rsi
    call skip_spaces
    mov r15, rax
    jmp uac_launch
@@:
    add rsi, 2
    jmp uac_skip_exe

uac_no_args:
    mov r15, rsi                ; Points to null terminator (empty args)

uac_launch:
    ; Zero SHELLEXECUTEINFOW at [rbp-312] (112 bytes)
    lea rdi, [rbp-312]
    xor rax, rax
    mov rcx, 14
uac_zero:
    mov qword ptr [rdi], rax
    add rdi, 8
    dec rcx
    jnz uac_zero

    ; Fill SHELLEXECUTEINFOW fields
    mov dword ptr [rbp-312], 112            ; cbSize
    lea rax, str_runas
    mov qword ptr [rbp-312+16], rax         ; lpVerb = "runas"
    lea rax, g_exePath
    mov qword ptr [rbp-312+24], rax         ; lpFile = exe path
    mov qword ptr [rbp-312+32], r15         ; lpParameters = arguments
    mov dword ptr [rbp-312+48], SW_SHOWNORMAL ; nShow

    lea rcx, [rbp-312]
    sub rsp, 32
    call ShellExecuteExW
    add rsp, 32

    ; Exit - elevated instance takes over, or user cancelled UAC
    xor ecx, ecx
    sub rsp, 32
    call ExitProcess
    add rsp, 32

admin_dispatch:
    ; Free early-parse argv (uac_already_admin will re-parse, matching
    ; original behavior). This keeps later dispatch self-contained.
    test r13, r13
    jz uac_already_admin
    mov rcx, r13
    sub rsp, 32
    call LocalFree
    add rsp, 32
    xor r13, r13
    jmp uac_already_admin


uac_already_admin:
    ; Get the full command line string
    sub rsp, 32
    call GetCommandLineW
    add rsp, 32
    mov r12, rax                ; R12 = command line string

    ; Parse command line into argv array
    lea rdx, [rbp-64]           ; RDX = pointer to argc output
    mov rcx, r12                ; RCX = command line string
    sub rsp, 32
    call CommandLineToArgvW
    add rsp, 32
    mov r13, rax                ; R13 = argv array

    ; Check if at least 2 arguments (exe + switch)
    mov eax, dword ptr [rbp-64] ; EAX = argc
    cmp eax, 2
    jl mode_gui_free            ; Less than 2 args: GUI mode

    ; Get second argument (first after exe path)
    mov rax, [r13+8]            ; RAX = argv[1]
    mov r14, rax                ; R14 = argv[1]

    ; Check if argv[1] matches "-cli"
    lea rdx, str_cliSwitch1
    mov rcx, r14
    call wcscmp_ci
    test rax, rax
    jnz mode_cli_found

    ; Check if argv[1] matches "-install"
    lea rdx, str_installSwitch
    mov rcx, r14
    call wcscmp_ci
    test rax, rax
    jnz mode_install_found

    ; Check if argv[1] matches "-uninstall"
    lea rdx, str_uninstallSwitch
    mov rcx, r14
    call wcscmp_ci
    test rax, rax
    jnz mode_uninstall_found

    ; Check if argv[1] matches "-shift"
    lea rdx, str_shiftSwitch
    mov rcx, r14
    call wcscmp_ci
    test rax, rax
    jnz mode_shift_found

    ; Check if argv[1] matches "-unshift"
    lea rdx, str_unshiftSwitch
    mov rcx, r14
    call wcscmp_ci
    test rax, rax
    jnz mode_unshift_found

    ; No recognized switch: an unknown switch starting with '-' (or '/')
    ; shows usage; anything else is treated as a file path for "Run as TI".
    cmp word ptr [r14], '-'
    je dispatch_show_usage
    cmp word ptr [r14], '/'
    je dispatch_show_usage
    jmp mode_file_run

dispatch_show_usage:
    mov rcx, r13                ; argv to free
    sub rsp, 32
    call ShowUsage              ; never returns

    add rsp, 32


mode_install_found:
    ; Free argv and register context menu
    mov rcx, r13
    sub rsp, 32
    call LocalFree
    add rsp, 32
    call InstallContextMenu
    xor ecx, ecx
    sub rsp, 32
    call ExitProcess
    add rsp, 32

mode_uninstall_found:
    ; Free argv and remove context menu
    mov rcx, r13
    sub rsp, 32
    call LocalFree
    add rsp, 32
    call UninstallContextMenu
    xor ecx, ecx
    sub rsp, 32
    call ExitProcess
    add rsp, 32

mode_shift_found:
    ; Free argv and install sethc.exe IFEO hook
    mov rcx, r13
    sub rsp, 32
    call LocalFree
    add rsp, 32
    call InstallShift
    xor ecx, ecx
    sub rsp, 32
    call ExitProcess
    add rsp, 32

mode_unshift_found:
    ; Free argv and remove sethc.exe IFEO hook
    mov rcx, r13
    sub rsp, 32
    call LocalFree
    add rsp, 32
    call UninstallShift
    xor ecx, ecx
    sub rsp, 32
    call ExitProcess
    add rsp, 32

mode_gui_free:
    ; Free argv array and proceed to GUI mode
    mov rcx, r13
    sub rsp, 32
    call LocalFree
    add rsp, 32
    jmp mode_gui                ; transfer into the GUI proc (which never
                                ; returns) — control never falls through

    ; Unreachable cleanup epilog, kept for unwind-table correctness.
    add rsp, 312
    pop r15
    pop r14
    pop r13
    pop r12
    pop rdi
    pop rsi
    pop rbx
    pop rbp
    ret
mainCRTStartup endp

; ==============================================================================
; mode_gui - GUI Mode entry point
;
; Purpose: Creates the main application window and runs the message loop
;          until WM_QUIT (or ESC). Reached either from main's dispatch when
;          no recognized switch was supplied, or from cli.asm's
;          skip_exe_for_file when the file-run path had no file argument.
;
; Parameters: None
;
; Returns: Never (every exit path goes through ExitProcess).
;
; Stack frame: 208 bytes (rbp-relative, mimicking the original frame so
;              [rbp-200] still points to a 48-byte MSG buffer).
; ==============================================================================
mode_gui proc frame
    push rbp
    .pushreg rbp
    mov rbp, rsp
    .setframe rbp, 0
    sub rsp, 216
    .allocstack 216
    .endprolog

    ; Get application instance handle.
    xor ecx, ecx
    sub rsp, 32
    call GetModuleHandleW
    add rsp, 32
    mov g_hInstance, rax

    ; Create the main application window.
    mov rcx, rax
    sub rsp, 32
    call CreateMainWindow
    add rsp, 32
    test rax, rax
    jz gui_exit                 ; Window creation failed

gui_msg_loop:
    ; Main message loop. MSG structure lives at [rbp-200..rbp-152].
    lea rcx, [rbp-200]
    xor edx, edx
    xor r8d, r8d
    xor r9d, r9d
    sub rsp, 32
    call GetMessageW
    add rsp, 32
    test eax, eax
    jz gui_exit                 ; WM_QUIT received

    ; Check for ESC key to exit application.
    mov eax, dword ptr [rbp-200+8]
    cmp eax, WM_KEYDOWN
    jne gui_msg_not_esc
    mov rax, qword ptr [rbp-200+16]
    cmp eax, VK_ESCAPE
    je gui_exit

gui_msg_not_esc:
    ; Process dialog messages (tab navigation, accelerators, etc.).
    lea rdx, [rbp-200]
    mov rcx, g_hwndMain
    sub rsp, 32
    call IsDialogMessageW
    add rsp, 32
    test eax, eax
    jnz gui_msg_loop            ; Message was consumed by IsDialogMessage

    ; Translate and dispatch regular window messages.
    lea rcx, [rbp-200]
    sub rsp, 32
    call TranslateMessage
    add rsp, 32

    lea rcx, [rbp-200]
    sub rsp, 32
    call DispatchMessageW
    add rsp, 32
    jmp gui_msg_loop

gui_exit:
    xor ecx, ecx
    sub rsp, 32
    call ExitProcess
    add rsp, 32

    ; Unreachable but present for the unwind-table tooling.
    add rsp, 216
    pop rbp
    ret
mode_gui endp

; ==============================================================================
; GetExeFileName - Extract filename from g_exePath
;
; Purpose: Scans g_exePath for the last backslash and returns a pointer
;          to the character after it (the filename portion).
;          Must call GetModuleFileNameW into g_exePath first.
;
; Parameters: None
;
; Returns: RAX = pointer to filename within g_exePath
;
; Modifies: RAX, RCX
; ==============================================================================

end
