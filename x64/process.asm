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
EXTRN GetCurrentDirectoryW:PROC
EXTRN GetStdHandle:PROC
EXTRN GetCurrentProcess:PROC
EXTRN DuplicateHandle:PROC
EXTRN WaitForSingleObject:PROC
EXTRN GetExitCodeProcess:PROC
EXTRN CreateFileW:PROC
EXTRN GetLastError:PROC
EXTRN wcslen_p:PROC

.const
str_nulDevice dw 'N','U','L',0
str_batchPrefix dw 'c','m','d','.','e','x','e',' ','/','d',' ','/','s',' ','/','c',' ','"',0

; ==============================================================================
; UNINITIALIZED DATA SECTION
; ==============================================================================
.data?
; Child working directory and private wrapper for .cmd/.bat command lines.
sysDirBuf       dw 260 dup(?)
batchCmdBuf     dw 32768 dup(?)

; ==============================================================================
; CODE SECTION
; ==============================================================================
.code

; CreateProcess does not invoke the command interpreter for batch files.
; Detect a .cmd/.bat first token and return a private command line wrapped as:
;   cmd.exe /d /s /c "<original command line>"
; Non-batch commands are returned unchanged in RAX.
PrepareBatchCommand proc
    ; Strip outer quotes if present (e.g. "whoami /groups" or "cmd /c ...").
    ; This must happen *before* the .cmd/.bat extension scan below: that scan
    ; measures the token length and reads backward from its end ([r8-8]..
    ; [r8-2]) to compare against ".cmd"/".bat". A trailing quote would shift
    ; every one of those offsets by one WCHAR and make the extension check
    ; miss a legitimately quoted batch file. Stripping first keeps the scan's
    ; offsets aligned to the real end of the token in both call paths (direct
    ; RunAsTrustedInstaller and the relay's re-spawned command).
    cmp word ptr [rcx], 0022h
    jne pbc_no_strip
    
    sub rsp, 40
    mov [rsp+32], rcx           ; save rcx
    call wcslen_p
    mov rcx, [rsp+32]           ; restore rcx
    add rsp, 40
    
    test rax, rax
    jz pbc_no_strip
    lea r8, [rcx + rax*2 - 2]
    cmp word ptr [r8], 0022h
    jne pbc_no_strip
    add rcx, 2                  ; skip opening quote
    mov word ptr [r8], 0        ; strip closing quote
pbc_no_strip:
    mov rax, rcx
    mov r8, rcx
    cmp word ptr [r8], '"'
    jne pbc_scan_unquoted
    add r8, 2
pbc_scan_quoted:
    movzx edx, word ptr [r8]
    test edx, edx
    jz pbc_return
    cmp edx, '"'
    je pbc_token_end
    add r8, 2
    jmp pbc_scan_quoted
pbc_scan_unquoted:
    movzx edx, word ptr [r8]
    test edx, edx
    jz pbc_token_end
    cmp edx, ' '
    je pbc_token_end
    cmp edx, 9
    je pbc_token_end
    add r8, 2
    jmp pbc_scan_unquoted

pbc_token_end:
    mov r9, r8
    sub r9, rcx
    cmp word ptr [rcx], '"'
    jne @F
    sub r9, 2                  ; exclude the opening quote from token length
@@:
    ; Extension check reads the last 4 WCHARs of the token *backward from its
    ; end* (r8 = one-past-last character): [r8-8] is 4 chars back, [r8-6] is
    ; 3 back, and so on, each offset -2 bytes because characters here are
    ; UTF-16 (2 bytes each), not the 1-byte ASCII a C string would use.
    ; This only looks at ".cmd" / ".bat"; anything else (.exe, no extension,
    ; a bare command like "dir") falls through to pbc_return unmodified.
    cmp r9, 8                  ; four UTF-16 characters
    jb pbc_return
    cmp word ptr [r8-8], '.'
    jne pbc_return
    movzx edx, word ptr [r8-6]
    or edx, 20h
    cmp edx, 'c'
    je pbc_check_cmd
    cmp edx, 'b'
    jne pbc_return
    movzx edx, word ptr [r8-4]
    or edx, 20h
    cmp edx, 'a'
    jne pbc_return
    movzx edx, word ptr [r8-2]
    or edx, 20h
    cmp edx, 't'
    jne pbc_return
    jmp pbc_wrap
pbc_check_cmd:
    movzx edx, word ptr [r8-4]
    or edx, 20h
    cmp edx, 'm'
    jne pbc_return
    movzx edx, word ptr [r8-2]
    or edx, 20h
    cmp edx, 'd'
    jne pbc_return

pbc_wrap:
    lea r10, batchCmdBuf
    lea r11, str_batchPrefix
pbc_copy_prefix:
    mov dx, word ptr [r11]
    mov word ptr [r10], dx
    add r11, 2
    add r10, 2
    test dx, dx
    jnz pbc_copy_prefix
    sub r10, 2                 ; overwrite the prefix terminator
    mov r11, rcx
    ; Cap = sizeof(batchCmdBuf) [32768 WCHARs] minus room already spent:
    ;   - str_batchPrefix content just copied in above (18 WCHARs)
    ;   - the closing quote written at pbc_close (1 WCHAR)
    ;   - the final NUL terminator (1 WCHAR)
    ; 32768 - 18 - 1 - 1 = 32748. If the source command line is longer than
    ; this, it is silently truncated (loop below stops when ecx hits 0);
    ; there is no overflow, just truncation of the copied command text.
    mov ecx, 32748             ; capacity minus prefix, closing quote and NUL
pbc_copy_command:
    mov dx, word ptr [r11]
    test dx, dx
    jz pbc_close
    test ecx, ecx
    jz pbc_return
    mov word ptr [r10], dx
    add r11, 2
    add r10, 2
    dec ecx
    jmp pbc_copy_command
pbc_close:
    mov word ptr [r10], '"'
    mov word ptr [r10+2], 0
    lea rax, batchCmdBuf
pbc_return:
    ret
PrepareBatchCommand endp

; ==============================================================================
; RunAsTrustedInstaller - Execute Command with Elevated Privileges
;
; Stack frame layout (post-prolog, all offsets rsp-relative, frame size 264):
;   [rsp+40..143]   STARTUPINFOW (104 bytes, zeroed as 13 qwords: cb field at
;                    +40, dwFlags at +40+60, wShowWindow at +40+64,
;                    hStdInput/Output/Error at +40+80/+40+88/+40+96)
;   [rsp+152..175]  PROCESS_INFORMATION (hProcess, hThread, dwProcessId,
;                    dwThreadId - 3 qwords zeroed, only first two are handles)
;   [rsp+184]       lpEnvironment (QWORD, filled by CreateEnvironmentBlock)
;   [rsp+192]       stdin duplicate-succeeded flag (DWORD, 0/1)
;   [rsp+196]       stdout duplicate-succeeded flag (DWORD, 0/1)
;   [rsp+200]       stderr duplicate-succeeded flag (DWORD, 0/1)
;   [rsp+204]       child exit-code scratch (DWORD, for GetExitCodeProcess)
;   [rsp+208..231]  SECURITY_ATTRIBUTES for the relay-mode NUL-device open
;                    (24 bytes: nLength, padding, lpSecurityDescriptor,
;                    bInheritHandle, padding)
;   [rsp+232]       relay-mode NUL-device read handle (QWORD), closed on
;                    every exit path once no longer needed
; Everything above is addressed at a further +80 (e.g. [rsp+80+40]) inside
; the CreateProcessWithTokenW call site, because that call temporarily does
; `sub rsp, 80` for its own stack parameters/shadow space first.
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

    mov ebx, edx                ; EBX = useNewConsole flag (save before it is overwritten by PrepareBatchCommand)
    call PrepareBatchCommand
    mov r12, rax                ; R12 = executable or wrapped batch command

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
    mov qword ptr [rsp+232], 0      ; relay NUL-input handle
    mov dword ptr g_childExitCode, 1 ; failure-safe default

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
    ; --- Mode 3: Redirect stdout/stderr independently while giving the child a
    ; real EOF-producing stdin. STARTF_USESTDHANDLES requires valid handles;
    ; NULL is not a substitute for the Windows NUL device.
    mov dword ptr [rsp+208], 24
    mov dword ptr [rsp+212], 0
    mov qword ptr [rsp+216], 0
    mov dword ptr [rsp+224], 1
    mov dword ptr [rsp+228], 0
    sub rsp, 64
    mov dword ptr [rsp+32], OPEN_EXISTING
    mov dword ptr [rsp+40], FILE_ATTRIBUTE_NORMAL
    mov qword ptr [rsp+48], 0
    lea r9, [rsp+64+208]
    mov r8d, FILE_SHARE_READ or FILE_SHARE_WRITE
    mov edx, GENERIC_READ
    lea rcx, str_nulDevice
    call CreateFileW
    add rsp, 64
    cmp rax, -1
    je rp_fail
    mov qword ptr [rsp+232], rax
    mov qword ptr [rsp+40+80], rax
    mov rax, qword ptr g_relayHandle
    mov qword ptr [rsp+40+88], rax
    mov rax, qword ptr g_relayErrHandle
    test rax, rax
    jnz @F
    mov rax, qword ptr g_relayHandle        ; backward-compatible fallback
@@:
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

    ; Preserve the caller's working directory. This is required for relative
    ; commands such as `cmdt -cli maintenance.cmd`; fall back to System32 only
    ; if the current directory cannot be represented in this MAX_PATH buffer.
    mov ecx, 260
    lea rdx, sysDirBuf
    sub rsp, 32
    call GetCurrentDirectoryW
    add rsp, 32
    test eax, eax
    jz rp_workdir_fallback
    cmp eax, 260
    jb rp_workdir_ready
rp_workdir_fallback:
    lea rcx, sysDirBuf
    mov edx, 260
    sub rsp, 32
    call GetSystemDirectoryW
    add rsp, 32
rp_workdir_ready:

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

    ; Preserve the actual program result independently of this routine's BOOL
    ; return value. This is deliberately done before either process handle is
    ; closed and before any cleanup API can overwrite RAX.
    mov rcx, qword ptr [rsp+152]
    lea rdx, [rsp+204]
    sub rsp, 32
    call GetExitCodeProcess
    add rsp, 32
    test eax, eax
    jz rp_skip_wait
    mov eax, dword ptr [rsp+204]
    mov dword ptr g_childExitCode, eax

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
    mov rcx, qword ptr [rsp+232]
    test rcx, rcx
    jz @F
    sub rsp, 32
    call CloseHandle
    add rsp, 32
    mov qword ptr [rsp+232], 0
@@:
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
    sub rsp, 32
    call GetLastError
    add rsp, 32
    mov dword ptr g_childExitCode, eax

    mov rcx, qword ptr [rsp+232]
    test rcx, rcx
    jz @F
    sub rsp, 32
    call CloseHandle
    add rsp, 32
@@:
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
