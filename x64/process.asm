; ==============================================================================
; CMDT - Run as TrustedInstaller
; Process Creation Module
; 
; Author: Marek Wesołowski (wesmar)
; Purpose: Handles process creation with TrustedInstaller privileges using
;          CreateProcessWithTokenW API.
; ==============================================================================

option casemap:none

include consts.inc
include globals.inc

; External function declarations
EXTRN GetTIToken:PROC
EXTRN CreateEnvironmentBlock:PROC
EXTRN DestroyEnvironmentBlock:PROC
EXTRN CreateProcessWithTokenW:PROC
EXTRN CloseHandle:PROC
EXTRN GetSystemDirectoryW:PROC
EXTRN GetStdHandle:PROC
EXTRN GetCurrentProcess:PROC
EXTRN DuplicateHandle:PROC
EXTRN WaitForSingleObject:PROC

; ==============================================================================
; UNINITIALIZED DATA SECTION
; ==============================================================================
.data?
; Buffer for system directory path (MAX_PATH = 260 WCHARs)
sysDirBuf       dw 260 dup(?)

; ==============================================================================
; CODE SECTION
; ==============================================================================
.code

; ==============================================================================
; RunAsTrustedInstaller - Execute Command with Elevated Privileges
; ==============================================================================
RunAsTrustedInstaller proc frame
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
    sub rsp, 264
    .allocstack 264
    .endprolog

    mov r12, rcx                ; R12 = command line string
    mov ebx, edx                ; EBX = useNewConsole flag

    ; Obtain TrustedInstaller token
    call GetTIToken
    test rax, rax
    jz rp_no_token
    mov r13, rax                ; R13 = TrustedInstaller token handle

    ; Initialize STARTUPINFOW structure to zero
    lea rdi, [rsp+40]
    xor rax, rax
    mov rcx, 13
@@:
    test rcx, rcx
    jz @F
    mov qword ptr [rdi], rax
    add rdi, 8
    dec rcx
    jmp @B
@@:

    ; Set structure size (cb field)
    mov dword ptr [rsp+40], STARTUPINFOW_SIZE
    mov dword ptr [rsp+192], 0      ; stdin duplicate flag
    mov dword ptr [rsp+196], 0      ; stdout duplicate flag
    mov dword ptr [rsp+200], 0      ; stderr duplicate flag

    ; Check for output-relay mode
    cmp qword ptr g_relayHandle, 0
    jne rp_relay_mode

    ; Check console mode
    test ebx, ebx
    jnz rp_new_console

    ; --- Mode 1: Inherit standard handles from parent process ---
    ; cmd.exe can give this GUI-subsystem process usable redirected std
    ; handles that are not themselves inheritable by the TI child. Duplicate
    ; them as inheritable handles before placing them in STARTUPINFO.
    sub rsp, 32
    call GetCurrentProcess
    add rsp, 32
    mov rsi, rax                    ; current process pseudo-handle

    mov ecx, STD_INPUT_HANDLE
    sub rsp, 32
    call GetStdHandle
    add rsp, 32
    mov qword ptr [rsp+40+80], rax  ; hStdInput
    test rax, rax
    jz rp_dup_stdout
    cmp rax, -1
    je rp_dup_stdout
    sub rsp, 64
    mov qword ptr [rsp+32], 0       ; dwDesiredAccess (ignored)
    mov dword ptr [rsp+40], 1       ; bInheritHandle = TRUE
    mov dword ptr [rsp+48], 2       ; DUPLICATE_SAME_ACCESS
    lea r9, [rsp+64+40+80]
    mov r8, rsi
    mov rdx, qword ptr [rsp+64+40+80]
    mov rcx, rsi
    call DuplicateHandle
    add rsp, 64
    test eax, eax
    jz rp_dup_stdout
    mov dword ptr [rsp+192], 1

rp_dup_stdout:
    mov ecx, STD_OUTPUT_HANDLE
    sub rsp, 32
    call GetStdHandle
    add rsp, 32
    mov qword ptr [rsp+40+88], rax  ; hStdOutput
    test rax, rax
    jz rp_dup_stderr
    cmp rax, -1
    je rp_dup_stderr
    sub rsp, 64
    mov qword ptr [rsp+32], 0
    mov dword ptr [rsp+40], 1
    mov dword ptr [rsp+48], 2
    lea r9, [rsp+64+40+88]
    mov r8, rsi
    mov rdx, qword ptr [rsp+64+40+88]
    mov rcx, rsi
    call DuplicateHandle
    add rsp, 64
    test eax, eax
    jz rp_dup_stderr
    mov dword ptr [rsp+196], 1

rp_dup_stderr:
    mov ecx, STD_ERROR_HANDLE
    sub rsp, 32
    call GetStdHandle
    add rsp, 32
    mov qword ptr [rsp+40+96], rax  ; hStdError
    test rax, rax
    jz rp_stdio_ready
    cmp rax, -1
    je rp_stdio_ready
    sub rsp, 64
    mov qword ptr [rsp+32], 0
    mov dword ptr [rsp+40], 1
    mov dword ptr [rsp+48], 2
    lea r9, [rsp+64+40+96]
    mov r8, rsi
    mov rdx, qword ptr [rsp+64+40+96]
    mov rcx, rsi
    call DuplicateHandle
    add rsp, 64
    test eax, eax
    jz rp_stdio_ready
    mov dword ptr [rsp+200], 1

rp_stdio_ready:
    mov dword ptr [rsp+40+60], STARTF_USESTDHANDLES
    jmp rp_setup_env

rp_relay_mode:
    ; --- Mode 3: Redirect child stdout/stderr to relay file ---
    mov qword ptr [rsp+40+80], 0
    mov rax, qword ptr g_relayHandle
    mov qword ptr [rsp+40+88], rax
    mov qword ptr [rsp+40+96], rax
    mov dword ptr [rsp+40+60], STARTF_USESTDHANDLES
    jmp rp_setup_env

rp_new_console:
    ; --- Mode 2: Create new console window ---
    mov dword ptr [rsp+40+60], STARTF_USESHOWWINDOW
    mov word ptr [rsp+40+64], SW_SHOWNORMAL

rp_setup_env:
    ; Initialize PROCESS_INFORMATION structure to zero
    lea rdi, [rsp+152]
    xor rax, rax
    mov qword ptr [rdi], rax
    mov qword ptr [rdi+8], rax
    mov qword ptr [rdi+16], rax

    ; Initialize lpEnvironment pointer to NULL
    mov qword ptr [rsp+184], 0

    ; Create environment block for the TrustedInstaller token
    xor r8d, r8d
    mov rdx, r13
    lea rcx, [rsp+184]
    sub rsp, 32
    call CreateEnvironmentBlock
    add rsp, 32

    ; Get Windows System directory path
    lea rcx, sysDirBuf
    mov edx, 260
    sub rsp, 32
    call GetSystemDirectoryW
    add rsp, 32

    ; Prepare stack parameters for CreateProcessWithTokenW
    ; Function requires 10 parameters (4 in registers, 6 on stack)
    sub rsp, 80                     ; Space for 6 stack parameters + 32 shadow + padding

    ; Set up register parameters
    mov rcx, r13                    ; RCX = hToken
    mov edx, 1                      ; EDX = LOGON_WITH_PROFILE
    xor r8, r8                      ; R8 = lpApplicationName = NULL
    mov r9, r12                     ; R9 = lpCommandLine

    ; Creation flags
    cmp qword ptr g_relayHandle, 0
    jne rp_flags_relay
    test ebx, ebx
    jnz rp_flags_new
    mov eax, CREATE_UNICODE_ENVIRONMENT
    jmp rp_flags_done
rp_flags_new:
    mov eax, CREATE_NEW_CONSOLE or CREATE_UNICODE_ENVIRONMENT
    jmp rp_flags_done
rp_flags_relay:
    mov eax, CREATE_NO_WINDOW or CREATE_UNICODE_ENVIRONMENT
rp_flags_done:
    mov [rsp+32], rax               ; dwCreationFlags

    mov rax, [rsp+80+184]
    mov [rsp+40], rax               ; lpEnvironment

    lea rax, sysDirBuf
    mov [rsp+48], rax               ; lpCurrentDirectory

    lea rax, [rsp+80+40]
    mov [rsp+56], rax               ; lpStartupInfo

    lea rax, [rsp+80+152]
    mov [rsp+64], rax               ; lpProcessInformation

    call CreateProcessWithTokenW
    add rsp, 80

    mov r14d, eax                   ; R14 = result (TRUE/FALSE)

    ; Destroy environment block
    mov rcx, [rsp+184]
    test rcx, rcx
    jz rp_skip_destroy_env
    sub rsp, 32
    call DestroyEnvironmentBlock
    add rsp, 32
rp_skip_destroy_env:

    ; Check if process creation succeeded
    test r14d, r14d
    jz rp_fail

    ; Determine if we should wait for the child process
    ; Wait if we are in relay mode OR if we are in inherit mode (CLI)
    cmp qword ptr g_relayHandle, 0
    jne rp_do_wait                  ; Always wait in relay mode
    test ebx, ebx
    jnz rp_skip_wait                ; Don't wait in new-console/GUI mode

rp_do_wait:
    mov rax, [rsp+152]              ; hProcess
    test rax, rax
    jz rp_skip_wait
    mov rcx, rax
    mov edx, INFINITE
    sub rsp, 32
    call WaitForSingleObject
    add rsp, 32

rp_skip_wait:

    ; Close inheritable duplicates made only for the child process. Do this
    ; after the wait in CLI mode so redirected output remains open until the
    ; child has finished writing.
    cmp dword ptr [rsp+192], 0
    je rp_close_dup_stdout
    mov rcx, qword ptr [rsp+40+80]
    sub rsp, 32
    call CloseHandle
    add rsp, 32
rp_close_dup_stdout:
    cmp dword ptr [rsp+196], 0
    je rp_close_dup_stderr
    mov rcx, qword ptr [rsp+40+88]
    sub rsp, 32
    call CloseHandle
    add rsp, 32
rp_close_dup_stderr:
    cmp dword ptr [rsp+200], 0
    je rp_close_pi
    mov rcx, qword ptr [rsp+40+96]
    sub rsp, 32
    call CloseHandle
    add rsp, 32

rp_close_pi:
    ; Close process and thread handles
    mov rax, [rsp+152]
    test rax, rax
    jz rp_skip_hp
    mov rcx, rax
    sub rsp, 32
    call CloseHandle
    add rsp, 32
rp_skip_hp:
    mov rax, [rsp+152+8]
    test rax, rax
    jz rp_skip_ht
    mov rcx, rax
    sub rsp, 32
    call CloseHandle
    add rsp, 32
rp_skip_ht:
    mov eax, 1                      ; Success
    jmp rp_done

rp_fail:
rp_no_token:
    xor eax, eax                    ; Failure

rp_done:
    add rsp, 264
    pop r14
    pop r13
    pop r12
    pop rdi
    pop rsi
    pop rbx
    ret
RunAsTrustedInstaller endp

end
