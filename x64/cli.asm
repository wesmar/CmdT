; ==============================================================================
; CMDT - Run as TrustedInstaller
; Command-Line / File-Run Dispatch Module
;
; Author: Marek Wesolowski (wesmar)
; Purpose: Hosts the two execution-style modes that actually run user
;          commands as TrustedInstaller — the explicit `-cli <command>` flow
;          and the context-menu `cmdt.exe "<file>"` flow. Both end up calling
;          RunAsTrustedInstaller; this module is what handles the command-
;          line parsing, optional .lnk resolution, and the internal -outfile
;          relay protocol that the non-admin path uses to capture output.
;
; Exported labels (PUBLIC via :: syntax):
;   mode_cli_found  - Jumped to from mainCRTStartup when argv[1] == "-cli".
;                     Validates -new placement and the argv count.
;   mode_file_run   - Jumped to when argv[1] is not a recognized switch and
;                     not a help token, e.g. a file path passed by Explorer's
;                     context menu integration.
;
; Both entry points share mainCRTStartup's stack frame — they are *jumped*
; to, never called, and rely on the existing rbp-relative locals
; ([rbp-64] argc, [rbp-104] SECURITY_ATTRIBUTES for the -outfile open).
; Every exit path ends in ExitProcess; control never returns to the caller.
; ==============================================================================

option casemap:none

include consts.inc
include globals.inc

; --- Cross-module strings (defined in main.asm, PUBLIC there) ---
EXTRN str_runas:WORD
EXTRN str_newSwitch:WORD
EXTRN str_extLnk_m:WORD
EXTRN str_space:WORD
EXTRN str_outfileFlag:WORD

; --- Cross-module jump target inside mainCRTStartup ---
EXTRN mode_gui:PROC

; --- Win32 APIs ---
EXTRN GetCommandLineW:PROC
EXTRN LocalFree:PROC
EXTRN CreateFileW:PROC
EXTRN CloseHandle:PROC
EXTRN ExitProcess:PROC

; --- Other modules ---
EXTRN RunAsTrustedInstaller:PROC
EXTRN ResolveLnkPath:PROC
EXTRN skip_spaces:PROC
EXTRN wcscpy_p:PROC
EXTRN wcscat_p:PROC
EXTRN wcscmp_ci:PROC
EXTRN wcscmp_token:PROC
EXTRN wcslen_p:PROC

; ==============================================================================
; CODE SECTION
;
; The labels below live in a single container `proc` purely so that ML64 is
; happy to host them in a .code section; nothing ever CALLs cli_module. Each
; entry point uses `::` so the linker can resolve cross-module jumps from
; mainCRTStartup's dispatch table.
; ==============================================================================
.code

PUBLIC mode_cli_found
PUBLIC mode_file_run

; ------------------------------------------------------------------------------
; mode_file_run - Right-click "Run as TrustedInstaller" file handler
;
; Triggered when argv[1] is not a recognized switch — typically a quoted
; file path delivered by the Explorer context menu we register. We resolve
; .lnk shortcuts ourselves so the target gets the TrustedInstaller token
; rather than launching cmd.exe with a shortcut argument.
; ------------------------------------------------------------------------------
mode_file_run:
    ; Free argv array allocated by CommandLineToArgvW
    mov rcx, r13
    sub rsp, 32
    call LocalFree
    add rsp, 32

    ; Get the raw command line to extract the file path
    sub rsp, 32
    call GetCommandLineW
    add rsp, 32
    mov rsi, rax                ; RSI = command line pointer
    xor edi, edi                ; EDI = quote state flag

    ; Skip past the executable path (which may be quoted)
skip_exe_for_file:
    mov ax, word ptr [rsi]
    test ax, ax
    jz mode_gui                 ; No file argument found, show GUI
    cmp ax, '"'
    jne @F
    xor edi, 1                  ; Toggle quote state
@@:
    cmp ax, ' '
    jne @F
    test edi, edi               ; Inside quotes?
    jnz @F
    add rsi, 2                  ; Space ends executable path
    mov rcx, rsi
    call skip_spaces            ; Skip whitespace after exe
    mov rsi, rax
    jmp run_file_direct         ; RSI now points to file path
@@:
    add rsi, 2
    jmp skip_exe_for_file

run_file_direct:
    ; RSI points to the file path argument (may be quoted from context menu).
    ; Check for .lnk shortcut and resolve before execution.

    ; Copy path to g_filePath, stripping leading quote
    mov rdx, rsi
    cmp word ptr [rdx], '"'
    jne @F
    add rdx, 2                  ; Skip leading quote
@@:
    lea rcx, g_filePath
    call wcscpy_p

    ; Remove trailing quote if present
    lea rcx, g_filePath
    call wcslen_p
    test rax, rax
    jz run_file_exec
    lea rcx, g_filePath
    cmp word ptr [rcx + rax*2 - 2], '"'
    jne @F
    mov word ptr [rcx + rax*2 - 2], 0
    dec rax
@@:
    ; Need at least 4 characters for .lnk extension
    cmp rax, 4
    jl run_file_exec

    ; Check if last 4 characters match ".lnk"
    lea rcx, g_filePath
    lea rcx, [rcx + rax*2 - 8]
    lea rdx, str_extLnk_m
    call wcscmp_ci
    test rax, rax
    jz run_file_exec            ; Not .lnk, execute directly

    ; Clear temporary buffer for shortcut arguments
    lea rdi, g_tempBuf
    xor rax, rax
    mov rcx, 260
@@:
    mov qword ptr [rdi], rax
    add rdi, 8
    dec rcx
    jnz @B

    ; Resolve .lnk target path and embedded arguments
    lea r8, g_tempBuf
    lea rdx, g_cmdBuf
    lea rcx, g_filePath
    sub rsp, 32
    call ResolveLnkPath
    add rsp, 32
    test rax, rax
    jz run_file_exec            ; Resolution failed, try direct

    ; Build command line from resolved target and arguments
    lea rcx, g_cmdBuf
    call wcslen_p
    test rax, rax
    jz run_file_lnk_args

    lea rcx, g_tempBuf
    call wcslen_p
    test rax, rax
    jz run_file_lnk_cmd         ; No embedded args, run target only

    lea rcx, g_cmdBuf
    lea rdx, str_space
    call wcscat_p
    lea rcx, g_cmdBuf
    lea rdx, g_tempBuf
    call wcscat_p
    jmp run_file_lnk_cmd

run_file_lnk_args:
    ; No target path, use embedded arguments only
    lea rcx, g_cmdBuf
    lea rdx, g_tempBuf
    call wcscpy_p

run_file_lnk_cmd:
    lea rcx, g_cmdBuf
    xor edx, edx
    sub rsp, 32
    call RunAsTrustedInstaller
    add rsp, 32
    xor ecx, ecx
    sub rsp, 32
    call ExitProcess
    add rsp, 32

run_file_exec:
    ; Not a .lnk file or resolution failed, execute as-is
    mov rcx, rsi
    xor edx, edx
    sub rsp, 32
    call RunAsTrustedInstaller
    add rsp, 32
    xor ecx, ecx
    sub rsp, 32
    call ExitProcess
    add rsp, 32

; ------------------------------------------------------------------------------
; mode_cli_found - Dispatcher target for `cmdt -cli <args>`
;
; Validates argc, recognizes the optional -new flag, then jumps to
; mode_cli_setup which parses the actual command. Frees the argv array
; before transferring control so the rest of the flow can rely on a
; raw GetCommandLineW pointer.
; ------------------------------------------------------------------------------
mode_cli_found:
    ; CLI mode detected - check for minimum arguments
    mov eax, dword ptr [rbp-64] ; EAX = argc
    cmp eax, 3
    jl cli_no_cmd_free          ; Error: no command specified

    ; Check if argv[2] is "-new" (new console flag)
    mov rax, [r13+16]           ; RAX = argv[2]
    lea rdx, str_newSwitch
    mov rcx, rax
    call wcscmp_ci
    test rax, rax
    jz cli_no_new_flag          ; Not "-new"

    ; "-new" flag found: need at least 4 args (exe, switch, -new, command)
    mov eax, dword ptr [rbp-64]
    cmp eax, 4
    jl cli_no_cmd_free          ; Error: no command after -new
    mov dword ptr g_useNewConsole, 1
    jmp cli_free_and_setup

cli_no_new_flag:
    mov dword ptr g_useNewConsole, 0

cli_free_and_setup:
    ; Free argv array allocated by CommandLineToArgvW
    mov rcx, r13
    sub rsp, 32
    call LocalFree
    add rsp, 32
    jmp mode_cli_setup

cli_no_cmd_free:
    ; Error: insufficient arguments for CLI mode
    mov rcx, r13
    sub rsp, 32
    call LocalFree
    add rsp, 32
    mov ecx, 1                  ; Exit code 1 (error)
    sub rsp, 32
    call ExitProcess
    add rsp, 32

; ------------------------------------------------------------------------------
; mode_cli_setup - Parse the raw cmdline and dispatch to .lnk-aware runner
;
; Walks the GetCommandLineW string past the exe path, past "-cli", optionally
; honours the internal "-outfile <path>" flag (relay protocol from the
; non-admin parent), optionally skips "-new", then locates the actual user
; command. If the command looks like a path-to-.lnk it is resolved into the
; real target before RunAsTrustedInstaller is invoked.
; ------------------------------------------------------------------------------
mode_cli_setup:
    sub rsp, 32
    call GetCommandLineW
    add rsp, 32
    mov rsi, rax                ; RSI = command line pointer
    xor r8, r8
    xor rdi, rdi                ; RDI = quote state flag

skip_exe_loop:
    ; Skip past the executable path (which may be quoted)
    mov ax, word ptr [rsi]
    test ax, ax
    jz cli_failed_setup         ; Unexpected end of command line
    cmp ax, '"'
    jne @F
    xor edi, 1                  ; Toggle quote state
@@:
    cmp ax, ' '
    jne @F
    test edi, edi               ; Are we inside quotes?
    jnz @F                      ; Yes: space is part of path
    add rsi, 2                  ; No: space ends exe path
    jmp skip_switch_init
@@:
    add rsi, 2                  ; Continue scanning
    jmp skip_exe_loop

skip_switch_init:
    ; Skip leading spaces before CLI switch
    mov rcx, rsi
    call skip_spaces
    mov rsi, rax

    xor edi, edi
skip_switch_loop:
    ; Skip the CLI switch argument
    mov ax, word ptr [rsi]
    test ax, ax
    jz cli_failed_setup
    cmp ax, ' '
    jne @F
    add rsi, 2
    jmp after_switch
@@:
    add rsi, 2
    jmp skip_switch_loop

after_switch:
    ; Skip spaces after CLI switch
    mov rcx, rsi
    call skip_spaces
    mov rsi, rax

    ; ===== Detect internal -outfile token (relay mode from non-admin parent) =====
    ; If present, the next token is "-outfile" followed by a (possibly
    ; quoted) temp file path. Open the file for inheritable write access
    ; and store the handle in g_relayHandle so RunAsTrustedInstaller will
    ; redirect the spawned process's stdout/stderr to it.
    mov rcx, rsi
    lea rdx, str_outfileFlag
    call wcscmp_token
    test rax, rax
    jz after_outfile

    ; Advance past "-outfile" (8 wchars = 16 bytes) and following spaces.
    add rsi, 16
    mov rcx, rsi
    call skip_spaces
    mov rsi, rax

    ; Extract path token into g_relayPath. Path may be quoted.
    mov rdi, rsi
    xor r8, r8                  ; r8 = 1 if path was quoted
    cmp word ptr [rdi], '"'
    jne outfile_scan_unquoted
    add rdi, 2                  ; skip opening quote
    mov rsi, rdi                ; rsi now at first char of path
    mov r8, 1
outfile_scan_quoted:
    mov ax, word ptr [rdi]
    test ax, ax
    jz outfile_copy
    cmp ax, '"'
    je outfile_copy
    add rdi, 2
    jmp outfile_scan_quoted

outfile_scan_unquoted:
    mov ax, word ptr [rdi]
    test ax, ax
    jz outfile_copy
    cmp ax, ' '
    je outfile_copy
    add rdi, 2
    jmp outfile_scan_unquoted

outfile_copy:
    ; Copy [rsi..rdi) into g_relayPath, null-terminate.
    lea r10, g_relayPath
    mov r11, rsi
outfile_copy_loop:
    cmp r11, rdi
    je outfile_copy_done
    mov ax, word ptr [r11]
    mov word ptr [r10], ax
    add r11, 2
    add r10, 2
    jmp outfile_copy_loop
outfile_copy_done:
    mov word ptr [r10], 0

    ; Advance rsi past the path token: skip optional closing quote and spaces.
    mov rsi, rdi
    test r8, r8
    jz outfile_advance_spaces
    cmp word ptr [rsi], '"'
    jne outfile_advance_spaces
    add rsi, 2                  ; skip closing quote
outfile_advance_spaces:
    mov rcx, rsi
    call skip_spaces
    mov rsi, rax

    ; Set up SECURITY_ATTRIBUTES at [rbp-104] with bInheritHandle = TRUE.
    ; Layout (24 bytes, x64 aligned):
    ;   +0  DWORD nLength
    ;   +4  DWORD (padding)
    ;   +8  QWORD lpSecurityDescriptor
    ;   +16 DWORD bInheritHandle
    ;   +20 DWORD (padding)
    mov dword ptr [rbp-104], 24
    mov dword ptr [rbp-100], 0
    mov qword ptr [rbp-96], 0
    mov dword ptr [rbp-88], 1
    mov dword ptr [rbp-84], 0

    ; CreateFileW(g_relayPath, GENERIC_WRITE, FILE_SHARE_READ, &sa,
    ;             CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL)
    sub rsp, 64
    mov dword ptr [rsp+32], CREATE_ALWAYS
    mov dword ptr [rsp+40], FILE_ATTRIBUTE_NORMAL
    mov qword ptr [rsp+48], 0
    lea r9, [rbp-104]
    mov r8d, FILE_SHARE_READ
    mov edx, GENERIC_WRITE
    lea rcx, g_relayPath
    call CreateFileW
    add rsp, 64
    cmp rax, -1
    je after_outfile            ; CreateFile failed → ignore relay
    mov qword ptr g_relayHandle, rax

after_outfile:
    ; Check if we need to skip the "-new" token
    cmp dword ptr g_useNewConsole, 0
    je run_command              ; No "-new" flag: proceed to command

skip_new_token:
    ; Skip the "-new" token
    mov ax, word ptr [rsi]
    test ax, ax
    jz cli_failed_setup
    cmp ax, ' '
    jne @F
    add rsi, 2                  ; Skip space after "-new"
    jmp run_command
@@:
    add rsi, 2
    jmp skip_new_token

run_command:
    ; Skip spaces before the actual command
    mov rcx, rsi
    call skip_spaces
    mov rsi, rax                ; RSI now points to the command

    ; Check if the command is a .lnk file (shortcut)
    mov rcx, rsi
    call wcslen_p
    mov r8, rax                 ; R8 = command length
    cmp r8, 4                   ; Minimum length for ".lnk"
    jl run_no_lnk

    ; Scan for space or quote to find end of path
    mov rdi, rsi
    xor rbx, rbx                ; RBX = pointer to first space (if found)
    xor r9, r9                  ; R9 = quote state
find_space_or_quote:
    mov ax, word ptr [rdi]
    test ax, ax
    jz check_lnk_ext            ; End of string
    cmp ax, '"'
    jne @F
    xor r9d, 1                  ; Toggle quote state
@@:
    cmp ax, ' '
    jne @F
    test r9d, r9d               ; Inside quotes?
    jnz @F                      ; Yes: ignore space
    mov rbx, rdi                ; No: mark first space position
    jmp check_lnk_ext           ; Stop at first space
@@:
    add rdi, 2
    jmp find_space_or_quote

check_lnk_ext:
    ; Check if path (before first space) ends with .lnk
    test rbx, rbx
    jz check_whole_path         ; No space found: check entire string

    ; Compute clean bounds (strip surrounding quotes from path)
    mov r10, rsi                ; R10 = start of path
    mov r11, rbx                ; R11 = end of path (space position)
    cmp word ptr [r10], '"'
    jne @F
    add r10, 2                  ; Skip opening quote
@@:
    cmp word ptr [r11-2], '"'
    jne @F
    sub r11, 2                  ; Skip closing quote
@@:
    mov rcx, r11
    sub rcx, r10
    shr rcx, 1                  ; RCX = path length in characters
    cmp rcx, 4
    jl run_no_lnk               ; Too short for .lnk

    ; Check if last 4 characters are ".lnk"
    lea rdi, [r10 + rcx*2 - 8]
    lea rdx, str_extLnk_m
    mov rcx, rdi
    call wcscmp_ci
    test rax, rax
    jz run_no_lnk               ; Not a .lnk file

    ; Save original RSI and RBX
    push rsi
    push rbx

    ; Copy clean path (without quotes) to g_filePath
    mov rcx, r11
    sub rcx, r10
    shr rcx, 1                  ; RCX = number of characters
    mov r8, rcx
    mov rsi, r10
    lea rdi, g_filePath
@@:
    test r8, r8
    jz @F
    mov ax, word ptr [rsi]
    mov word ptr [rdi], ax
    add rsi, 2
    add rdi, 2
    dec r8
    jmp @B
@@:
    mov word ptr [rdi], 0       ; Null terminate

    ; Extract arguments after the path
    pop rbx
    add rbx, 2                  ; Skip the space
    mov rcx, rbx
    call skip_spaces
    mov rsi, rax
    mov rdx, rsi
    lea rcx, g_argsBuf
    xchg rcx, rdx
    call wcscpy_p               ; Copy arguments to g_argsBuf

    ; Clear g_tempBuf (for shortcut arguments from .lnk)
    lea rdi, g_tempBuf
    xor rax, rax
    mov rcx, 260
@@:
    test rcx, rcx
    jz @F
    mov qword ptr [rdi], rax
    add rdi, 8
    dec rcx
    jmp @B
@@:

    ; Resolve the .lnk file to get target path and built-in arguments
    lea r8, g_tempBuf
    lea rdx, g_cmdBuf
    lea rcx, g_filePath
    sub rsp, 32
    call ResolveLnkPath
    add rsp, 32
    test rax, rax
    pop rsi
    jz run_no_lnk

    ; Build final command: target + .lnk args + user args
    lea rcx, g_cmdBuf
    call wcslen_p
    test rax, rax
    jz use_args_only

    ; Append space after target path
    lea rdi, g_cmdBuf
    lea rdi, [rdi + rax*2]
    mov word ptr [rdi], ' '
    add rdi, 2
    mov word ptr [rdi], 0

use_args_only:
    ; Append .lnk arguments
    lea rdx, g_tempBuf
    lea rcx, g_cmdBuf
    call wcscat_p

    ; Check if user provided additional arguments
    lea rcx, g_argsBuf
    call wcslen_p
    test rax, rax
    jz run_resolved

    ; Append space and user arguments
    lea rdx, str_space
    lea rcx, g_cmdBuf
    call wcscat_p
    lea rdx, g_argsBuf
    lea rcx, g_cmdBuf
    call wcscat_p
    jmp run_resolved

check_whole_path:
    ; No space found: entire command might be a .lnk file
    mov r10, rsi                ; R10 = start
    mov r11, r8                 ; R11 = length
    cmp word ptr [r10], '"'
    jne @F
    add r10, 2                  ; Skip opening quote
    dec r11
@@:
    cmp r11, 1
    jl run_no_lnk
    cmp word ptr [r10 + r11*2 - 2], '"'
    jne @F
    dec r11                     ; Exclude closing quote
@@:
    cmp r11, 4
    jl run_no_lnk

    ; Check if last 4 characters are ".lnk"
    lea rdi, [r10 + r11*2 - 8]
    lea rdx, str_extLnk_m
    mov rcx, rdi
    call wcscmp_ci
    test rax, rax
    jz run_no_lnk

    ; Clear g_tempBuf
    lea rdi, g_tempBuf
    xor rax, rax
    mov rcx, 260
@@:
    test rcx, rcx
    jz @F
    mov qword ptr [rdi], rax
    add rdi, 8
    dec rcx
    jmp @B
@@:

    ; Copy clean path to g_filePath
    mov rcx, r11
    lea rdi, g_filePath
    mov rdx, r10
@@:
    test rcx, rcx
    jz @F
    mov ax, word ptr [rdx]
    mov word ptr [rdi], ax
    add rdx, 2
    add rdi, 2
    dec rcx
    jmp @B
@@:
    mov word ptr [rdi], 0

    ; Resolve the .lnk file
    lea r8, g_tempBuf
    lea rdx, g_cmdBuf
    lea rcx, g_filePath
    sub rsp, 32
    call ResolveLnkPath
    add rsp, 32
    test rax, rax
    jz run_no_lnk

    ; Build command from resolved target and arguments
    lea rcx, g_cmdBuf
    call wcslen_p
    test rax, rax
    jz use_lnk_args_only

    ; Append space after target path
    lea rdi, g_cmdBuf
    lea rdi, [rdi + rax*2]
    mov word ptr [rdi], ' '
    add rdi, 2
    mov word ptr [rdi], 0

use_lnk_args_only:
    ; Append .lnk arguments
    lea rdx, g_tempBuf
    lea rcx, g_cmdBuf
    call wcscat_p
    jmp run_resolved

run_resolved:
    ; Execute the resolved command as TrustedInstaller
    mov edx, dword ptr g_useNewConsole
    lea rcx, g_cmdBuf
    sub rsp, 32
    call RunAsTrustedInstaller
    add rsp, 32
    jmp run_check_result

run_no_lnk:
    ; Execute the original command (not a .lnk file)
    mov edx, dword ptr g_useNewConsole
    mov rcx, rsi
    sub rsp, 32
    call RunAsTrustedInstaller
    add rsp, 32

run_check_result:
    ; Close the relay file handle (if any) so the parent reads a flushed,
    ; complete file. OS would close on exit anyway, but doing it here is
    ; explicit and lets the parent observe the file immediately.
    mov rcx, qword ptr g_relayHandle
    test rcx, rcx
    jz @F
    sub rsp, 32
    call CloseHandle
    add rsp, 32
    mov qword ptr g_relayHandle, 0
@@:
    ; Check execution result
    test rax, rax
    jz cli_failed

    ; Success: exit with code 0
    xor ecx, ecx
    sub rsp, 32
    call ExitProcess
    add rsp, 32

cli_failed_setup:
cli_failed:
    ; Close relay handle on failure path too
    mov rcx, qword ptr g_relayHandle
    test rcx, rcx
    jz @F
    sub rsp, 32
    call CloseHandle
    add rsp, 32
    mov qword ptr g_relayHandle, 0
@@:
    ; Failure: exit with code 1
    mov ecx, 1
    sub rsp, 32
    call ExitProcess
    add rsp, 32

end
