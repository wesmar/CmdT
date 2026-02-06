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
;
; Purpose: Launches a process with TrustedInstaller privileges. Obtains a
;          TrustedInstaller token and creates a new process using that token.
;          Supports both console inheritance and new console creation.
;
; Parameters:
;   RCX = Pointer to command line string (wide char)
;   EDX = useNewConsole flag (0 = inherit handles, 1 = create new console)
;
; Returns:
;   RAX = 1 on success, 0 on failure
;
; Stack frame: 264 bytes for local variables including:
;   - STARTUPINFOW structure (104 bytes)
;   - PROCESS_INFORMATION structure (24 bytes)
;   - Environment block pointer
;   - System directory buffer pointer
;
; Notes:
;   - Gets TrustedInstaller token from GetTIToken()
;   - Creates environment block for the token
;   - Sets working directory to Windows System directory
;   - Handles both console modes (inherit or new)
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
    jz rp_no_token              ; Failed to get token
    mov r13, rax                ; R13 = TrustedInstaller token handle

    ; Initialize STARTUPINFOW structure to zero
    ; Structure is at [rsp+40], size = 104 bytes = 13 QWORDs
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

    ; Check console mode
    test ebx, ebx
    jnz rp_new_console          ; Create new console

    ; --- Mode 1: Inherit standard handles from parent process ---
    
    ; Get standard input handle
    mov ecx, STD_INPUT_HANDLE
    sub rsp, 32
    call GetStdHandle
    add rsp, 32
    mov qword ptr [rsp+40+80], rax  ; hStdInput

    ; Get standard output handle
    mov ecx, STD_OUTPUT_HANDLE
    sub rsp, 32
    call GetStdHandle
    add rsp, 32
    mov qword ptr [rsp+40+88], rax  ; hStdOutput

    ; Get standard error handle
    mov ecx, STD_ERROR_HANDLE
    sub rsp, 32
    call GetStdHandle
    add rsp, 32
    mov qword ptr [rsp+40+96], rax  ; hStdError

    ; Set dwFlags to use standard handles
    mov dword ptr [rsp+40+60], STARTF_USESTDHANDLES
    jmp rp_setup_env

rp_new_console:
    ; --- Mode 2: Create new console window ---
    
    ; Set dwFlags to use show window setting
    mov dword ptr [rsp+40+60], STARTF_USESHOWWINDOW
    
    ; Set wShowWindow to SW_SHOWNORMAL
    mov word ptr [rsp+40+64], SW_SHOWNORMAL

rp_setup_env:
    ; Initialize PROCESS_INFORMATION structure to zero
    ; Structure is at [rsp+152], size = 24 bytes
    lea rdi, [rsp+152]
    xor rax, rax
    mov qword ptr [rdi], rax        ; hProcess
    mov qword ptr [rdi+8], rax      ; hThread
    mov qword ptr [rdi+16], rax     ; dwProcessId and dwThreadId

    ; Initialize lpEnvironment pointer to NULL
    ; Will be filled by CreateEnvironmentBlock if successful
    mov qword ptr [rsp+184], 0

    ; Create environment block for the TrustedInstaller token
    ; BOOL CreateEnvironmentBlock(
    ;   [out] LPVOID  *lpEnvironment,    -> [rsp+184]
    ;   [in]  HANDLE  hToken,            -> r13
    ;   [in]  BOOL    bInherit           -> FALSE (0)
    ; )
    xor r8d, r8d                    ; R8 = bInherit = FALSE
    mov rdx, r13                    ; RDX = token handle
    lea rcx, [rsp+184]              ; RCX = &lpEnvironment
    sub rsp, 32
    call CreateEnvironmentBlock
    add rsp, 32
    ; Note: Continue even if this fails (lpEnvironment remains NULL)

    ; Get Windows System directory path
    ; UINT GetSystemDirectoryW(
    ;   [out] LPWSTR lpBuffer,          -> sysDirBuf
    ;   [in]  UINT   uSize              -> 260
    ; )
    lea rcx, sysDirBuf
    mov edx, 260
    sub rsp, 32
    call GetSystemDirectoryW
    add rsp, 32

    ; Prepare stack parameters for CreateProcessWithTokenW
    ; Function requires 10 parameters (4 in registers, 6 on stack)
    sub rsp, 80                     ; Reserve space for 6 stack parameters + shadow

    ; Set up register parameters
    mov rcx, r13                    ; RCX = hToken
    mov edx, 1                      ; EDX = dwLogonFlags = LOGON_WITH_PROFILE
    xor r8, r8                      ; R8 = lpApplicationName = NULL
    mov r9, r12                     ; R9 = lpCommandLine

    ; Set up stack parameters (in order):
    ; [rsp+32] = dwCreationFlags
    ; [rsp+40] = lpEnvironment
    ; [rsp+48] = lpCurrentDirectory
    ; [rsp+56] = lpStartupInfo
    ; [rsp+64] = lpProcessInformation

    ; Determine creation flags based on console mode
    test ebx, ebx
    jnz rp_flags_new
    mov eax, CREATE_UNICODE_ENVIRONMENT         ; Inherit console
    jmp rp_flags_done
rp_flags_new:
    mov eax, CREATE_NEW_CONSOLE or CREATE_UNICODE_ENVIRONMENT
rp_flags_done:
    mov [rsp+32], rax               ; dwCreationFlags

    mov rax, [rsp+80+184]           ; Get lpEnvironment
    mov [rsp+40], rax               ; lpEnvironment

    lea rax, sysDirBuf
    mov [rsp+48], rax               ; lpCurrentDirectory = System directory

    lea rax, [rsp+80+40]
    mov [rsp+56], rax               ; lpStartupInfo

    lea rax, [rsp+80+152]
    mov [rsp+64], rax               ; lpProcessInformation

    ; BOOL CreateProcessWithTokenW(
    ;   [in]            HANDLE                hToken,
    ;   [in]            DWORD                 dwLogonFlags,
    ;   [in, optional]  LPCWSTR               lpApplicationName,
    ;   [in, out]       LPWSTR                lpCommandLine,
    ;   [in]            DWORD                 dwCreationFlags,
    ;   [in, optional]  LPVOID                lpEnvironment,
    ;   [in, optional]  LPCWSTR               lpCurrentDirectory,
    ;   [in]            LPSTARTUPINFOW        lpStartupInfo,
    ;   [out]           LPPROCESS_INFORMATION lpProcessInformation
    ; )
    call CreateProcessWithTokenW
    add rsp, 80

    mov r14d, eax                   ; R14 = result (TRUE/FALSE)

    ; Destroy environment block if it was created
    mov rcx, [rsp+184]
    test rcx, rcx
    jz rp_skip_destroy_env          ; NULL: wasn't created
    sub rsp, 32
    call DestroyEnvironmentBlock
    add rsp, 32
rp_skip_destroy_env:

    ; Check if process creation succeeded
    test r14d, r14d
    jz rp_fail                      ; Failed

    ; Close process handle (we don't need it)
    mov rax, [rsp+152]              ; hProcess
    test rax, rax
    jz rp_skip_hp
    mov rcx, rax
    sub rsp, 32
    call CloseHandle
    add rsp, 32

rp_skip_hp:
    ; Close thread handle (we don't need it)
    mov rax, [rsp+152+8]            ; hThread
    test rax, rax
    jz rp_skip_ht
    mov rcx, rax
    sub rsp, 32
    call CloseHandle
    add rsp, 32

rp_skip_ht:
    ; Return success
    mov eax, 1
    jmp rp_done

rp_fail:
rp_no_token:
    ; Return failure
    xor eax, eax

rp_done:
    ; Cleanup and return
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
