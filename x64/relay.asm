; ==============================================================================
; CMDT - Run as TrustedInstaller
; Non-Admin Output-Relay Module
;
; Author: Marek Wesolowski (wesmar)
; Purpose: Implements the relay path that lets `cmdt -cli <command>` run
;          from a non-admin shell still return its child's stdout/stderr to
;          the caller's redirect target or console. The trick: spawn an
;          elevated copy of cmdt with an internal `-outfile <path>` flag
;          telling it to redirect spawned-process output into a temp file,
;          wait for that elevated process to exit, then stream the temp file
;          contents back to *this* (non-admin) process's STD_OUTPUT — which
;          cmd.exe wired up before launching us, so `>file` / `|pipe` /
;          `>>file` all work transparently.
;
; Exported routine:
;   NonAdminRelayLaunch - Attempt the relay path. Returns 0 if it declined
;                         (e.g. user passed -new, which conflicts with output
;                         capture; or temp-file setup failed). Never returns
;                         if the relay actually ran — every success/failure
;                         past ShellExecuteEx ends in ExitProcess.
; ==============================================================================

option casemap:none

include consts.inc
include globals.inc

; --- Cross-module strings owned by main.asm ---
EXTRN str_runas:WORD
EXTRN str_newSwitch:WORD

; --- Win32 APIs ---
EXTRN GetModuleFileNameW:PROC
EXTRN GetTempPathW:PROC
EXTRN GetTempFileNameW:PROC
EXTRN ShellExecuteExW:PROC
EXTRN WaitForSingleObject:PROC
EXTRN CloseHandle:PROC
EXTRN CreateFileW:PROC
EXTRN ReadFile:PROC
EXTRN WriteFile:PROC
EXTRN DeleteFileW:PROC
EXTRN GetStdHandle:PROC
EXTRN ExitProcess:PROC

; --- In-project helpers ---
EXTRN wcscpy_p:PROC
EXTRN wcscat_p:PROC
EXTRN wcscmp_ci:PROC
EXTRN skip_spaces:PROC

EXTRN NudgeConsolePrompt:PROC

; ==============================================================================
; CONSTANT STRING DATA - private to this module
; ==============================================================================
.const

; 3-char prefix used by GetTempFileNameW to build the temp file name. Three
; chars max per the API contract; the API itself appends a hex sequence + .TMP.
str_cmdtPrefix  dw 'C','M','D',0

; Fixed prefix injected into the elevated child's argument string by the
; non-admin parent. The full string built becomes:
;   "-cli -outfile \"<relay-path>\" <rest of original args after -cli>"
str_relayPrefix dw '-','c','l','i',' ','-','o','u','t','f','i','l','e',' ','"',0
str_relayMid    dw '"',' ',0

; Interactive shells that cannot tolerate the relay path (CREATE_NO_WINDOW +
; redirected stdout to a temp file). When the user runs `cmdt -cli <shell>`
; with no further arguments we decline relay so the plain UAC self-elevate
; fallback gives them a real, attachable console window instead.
str_shell_cmd      dw 'c','m','d',0
str_shell_cmd_exe  dw 'c','m','d','.','e','x','e',0
str_shell_ps       dw 'p','o','w','e','r','s','h','e','l','l',0
str_shell_ps_exe   dw 'p','o','w','e','r','s','h','e','l','l','.','e','x','e',0
str_shell_pwsh     dw 'p','w','s','h',0
str_shell_pwsh_exe dw 'p','w','s','h','.','e','x','e',0

; Null-terminated pointer table walked by the shell-guard loop in
; NonAdminRelayLaunch. Add new shell names by inserting another pointer
; before the terminating zero.
shell_names_table  dq str_shell_cmd, str_shell_cmd_exe, \
                      str_shell_ps,  str_shell_ps_exe,  \
                      str_shell_pwsh, str_shell_pwsh_exe, 0

; ==============================================================================
; CODE SECTION
; ==============================================================================
.code

; ==============================================================================
; NonAdminRelayLaunch - Run the non-admin -cli output-relay flow
;
; Parameters:
;   ECX = argc
;   RDX = argv pointer (LocalFree-owned, but we never free it; caller drops it)
;   R8  = raw command-line string from GetCommandLineW
;
; Returns:
;   RAX = 0 if relay declined (-new flag present, or temp setup failed).
;         Caller should fall back to plain UAC self-elevate.
;   Never returns once the elevated child has been spawned — every exit path
;   from that point on goes through ExitProcess.
;
; Stack frame layout (post-prolog, all offsets rsp-relative):
;   [rsp+0..31]    shadow space for callee parameters
;   [rsp+40..151]  SHELLEXECUTEINFOW (112 bytes)
;   [rsp+152..159] bytesRead temporary (ReadFile output)
;   [rsp+160..167] bytesWritten temporary (WriteFile output)
;   [rsp+168..223] scratch / alignment padding
;
; Sizing: 7 callee-saved pushes + sub rsp,224 keeps every CALL site
;         16-byte aligned (224 mod 16 = 0).
; ==============================================================================
NonAdminRelayLaunch proc frame
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
    sub rsp, 224
    .allocstack 224
    .endprolog

    mov r12, r8                 ; r12 = raw cmdline
    mov r13, rdx                ; r13 = argv

    ; If user requested -new, the spawned command must run in a visible new
    ; console window — that conflicts with output capture (which requires
    ; CREATE_NO_WINDOW). Decline so the caller falls back to plain UAC.
    cmp ecx, 3
    jl narl_setup

    ; Stash argc in ebx (preserved across wcscmp_ci calls; rbx is rewritten
    ; later at narl_open_file with the relay-file read handle, after the
    ; shell-guard loop below is finished with it).
    mov ebx, ecx

    mov rcx, [r13+16]                   ; argv[2]
    lea rdx, str_newSwitch
    call wcscmp_ci
    test rax, rax
    jnz narl_decline

    ; Interactive-shell guard. The relay path uses CREATE_NO_WINDOW plus a
    ; temp-file capture of stdout — fundamentally incompatible with a shell
    ; that expects an attached console (no stdin, no prompt redraw, no
    ; output to the user). When argv looks like exactly `cmdt -cli <shell>`
    ; (argc == 3, no extra tokens like `/c` or `-Command`), bail out and let
    ; the caller's plain UAC self-elevate path spawn a real new console.
    ; Anything with extra arguments (e.g. `cmd /c dir`) keeps the relay so
    ; its output is still streamed back to the caller.
    cmp ebx, 3
    jne narl_setup

    lea r14, shell_names_table
narl_shell_loop:
    mov rdx, qword ptr [r14]
    test rdx, rdx
    jz narl_setup                       ; List exhausted, not an interactive shell
    mov rcx, [r13+16]                   ; argv[2]
    call wcscmp_ci
    test rax, rax
    jnz narl_decline                    ; Match: interactive shell → fall back to plain UAC
    add r14, 8
    jmp narl_shell_loop

narl_setup:
    ; Get exe path for ShellExecuteExW.lpFile.
    lea rdx, g_exePath
    mov r8d, 260
    xor ecx, ecx
    call GetModuleFileNameW

    ; Get system temp directory.
    lea rdx, g_tempDirBuf
    mov ecx, 260
    call GetTempPathW
    test eax, eax
    jz narl_decline

    ; Create unique temp file name.
    lea rcx, g_tempDirBuf
    lea rdx, str_cmdtPrefix
    xor r8d, r8d                ; uUnique = 0 (use system time)
    lea r9, g_relayPath
    call GetTempFileNameW
    test eax, eax
    jz narl_decline

    ; Build the modified argument string in g_relayArgs:
    ;   "-cli -outfile \"" + g_relayPath + "\" " + REST
    ; where REST is everything in the original cmdline after the "-cli"
    ; token (preserving original quoting/spacing).
    mov word ptr g_relayArgs, 0
    lea rcx, g_relayArgs
    lea rdx, str_relayPrefix
    call wcscpy_p

    lea rcx, g_relayArgs
    lea rdx, g_relayPath
    call wcscat_p

    lea rcx, g_relayArgs
    lea rdx, str_relayMid
    call wcscat_p

    ; Locate REST by walking the raw cmdline: skip exe path, then skip
    ; the "-cli" token, leaving rsi at the start of REST (or '\0').
    mov rsi, r12
    xor edi, edi                ; quote-state flag
narl_skip_exe:
    mov ax, word ptr [rsi]
    test ax, ax
    jz narl_append_rest
    cmp ax, '"'
    jne @F
    xor edi, 1
@@:
    cmp ax, ' '
    jne @F
    test edi, edi
    jnz @F
    add rsi, 2
    mov rcx, rsi
    call skip_spaces
    mov rsi, rax
    jmp narl_skip_cli
@@:
    add rsi, 2
    jmp narl_skip_exe

narl_skip_cli:
    ; rsi points at "-cli". Walk to next space (or '\0').
    mov ax, word ptr [rsi]
    test ax, ax
    jz narl_append_rest
    cmp ax, ' '
    jne @F
    add rsi, 2
    mov rcx, rsi
    call skip_spaces
    mov rsi, rax
    jmp narl_append_rest
@@:
    add rsi, 2
    jmp narl_skip_cli

narl_append_rest:
    lea rcx, g_relayArgs
    mov rdx, rsi
    call wcscat_p

    ; Zero SHELLEXECUTEINFOW at [rsp+40] (112 bytes = 14 qwords).
    lea rdi, [rsp+40]
    xor rax, rax
    mov rcx, 14
@@:
    mov qword ptr [rdi], rax
    add rdi, 8
    dec rcx
    jnz @B

    ; Fill SHELLEXECUTEINFOW. fMask = SEE_MASK_NOCLOSEPROCESS so we get
    ; back a process handle to wait on.
    mov dword ptr [rsp+40], 112                         ; cbSize
    mov dword ptr [rsp+40+4], SEE_MASK_NOCLOSEPROCESS   ; fMask
    lea rax, str_runas
    mov qword ptr [rsp+40+16], rax                      ; lpVerb
    lea rax, g_exePath
    mov qword ptr [rsp+40+24], rax                      ; lpFile
    lea rax, g_relayArgs
    mov qword ptr [rsp+40+32], rax                      ; lpParameters
    mov dword ptr [rsp+40+48], SW_HIDE                  ; nShow (the elevated
                                                        ; child has no console
                                                        ; window of its own)

    lea rcx, [rsp+40]
    call ShellExecuteExW
    test eax, eax
    jz narl_delete_only         ; UAC denied / cancelled

    ; Wait for elevated child to finish writing temp file. hProcess is at
    ; offset 104 in SHELLEXECUTEINFOW on x64 (last field of the struct).
    mov rcx, qword ptr [rsp+40+104]
    test rcx, rcx
    jz narl_open_file
    mov edx, INFINITE
    call WaitForSingleObject

    mov rcx, qword ptr [rsp+40+104]
    call CloseHandle

narl_open_file:
    ; CreateFileW(g_relayPath, GENERIC_READ, FILE_SHARE_READ, NULL,
    ;             OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, NULL).
    ;
    ; Stack layout for the call: [rsp+32] dwCreationDisposition,
    ; [rsp+40] dwFlagsAndAttributes, [rsp+48] hTemplateFile.
    sub rsp, 32
    mov dword ptr [rsp+32], OPEN_EXISTING
    mov dword ptr [rsp+40], FILE_ATTRIBUTE_NORMAL
    mov qword ptr [rsp+48], 0
    xor r9, r9
    mov r8d, FILE_SHARE_READ
    mov edx, GENERIC_READ
    lea rcx, g_relayPath
    call CreateFileW
    add rsp, 32
    cmp rax, -1
    je narl_delete_only
    mov rbx, rax                ; rbx = relay-file read handle

    ; Get our STD_OUTPUT_HANDLE (the parent shell wired this up, either as
    ; its console or as a redirected file/pipe).
    mov ecx, STD_OUTPUT_HANDLE
    call GetStdHandle
    mov r15, rax

narl_copy_loop:
    ; ReadFile(rbx, g_relayReadBuf, 4096, &[rsp+152], NULL)
    lea r9, [rsp+152]           ; lpNumberOfBytesRead
    mov r8d, 4096
    lea rdx, g_relayReadBuf
    mov rcx, rbx
    mov qword ptr [rsp+32], 0   ; lpOverlapped (in our existing shadow space)
    call ReadFile
    test eax, eax
    jz narl_close_file
    mov eax, dword ptr [rsp+152]
    test eax, eax
    jz narl_close_file          ; EOF

    ; WriteFile(r15, g_relayReadBuf, bytesRead, &[rsp+160], NULL)
    lea r9, [rsp+160]
    mov r8d, eax
    lea rdx, g_relayReadBuf
    mov rcx, r15
    mov qword ptr [rsp+32], 0
    call WriteFile
    jmp narl_copy_loop

narl_close_file:
    mov rcx, rbx
    call CloseHandle

narl_delete_only:
    call NudgeConsolePrompt

    ; Delete the temp file (best effort) and exit the process. Once we've
    ; spawned and waited on the elevated child we don't return to the
    ; caller — there's no useful fallback left at this point.
    lea rcx, g_relayPath
    call DeleteFileW

    xor ecx, ecx
    call ExitProcess

narl_decline:
    ; Bail out without touching anything else. Caller falls back to the
    ; plain UAC self-elevate path.
    xor eax, eax
    add rsp, 224
    pop r15
    pop r14
    pop r13
    pop r12
    pop rdi
    pop rsi
    pop rbx
    ret
NonAdminRelayLaunch endp

end
