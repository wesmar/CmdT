; ==============================================================================
; CMDT - Run as TrustedInstaller
; Process Creation Module
; 
; Author: Marek Wesołowski (wesmar)
; Purpose: Implements process creation with TrustedInstaller privileges.
;          Handles environment block creation and process spawning with proper
;          console and handle inheritance based on operation mode.
;
; Features:
;          - Creates processes using TrustedInstaller security token
;          - Environment block management with user profile
;          - Dual console mode: inherit parent or create new console
;          - Standard handle inheritance for CLI mode
;          - Proper resource cleanup (handles, environment blocks)
; ==============================================================================

.586                            ; Target 80586 instruction set
.model flat, stdcall            ; 32-bit flat memory model, stdcall convention
option casemap:none             ; Case-sensitive symbol names

include consts.inc              ; Windows API constants
include globals.inc             ; Global variable declarations

; ==============================================================================
; EXTERNAL FUNCTION PROTOTYPES
; ==============================================================================

; Token management
GetTIToken                  PROTO   ; Acquires TrustedInstaller token

; Environment and process creation APIs
CreateEnvironmentBlock      PROTO :DWORD,:DWORD,:DWORD
DestroyEnvironmentBlock     PROTO :DWORD
CreateProcessWithTokenW     PROTO :DWORD,:DWORD,:DWORD,:DWORD,:DWORD,:DWORD,:DWORD,:DWORD,:DWORD
CloseHandle                 PROTO :DWORD
GetSystemDirectoryW         PROTO :DWORD,:DWORD
GetCurrentDirectoryW        PROTO :DWORD,:DWORD
GetStdHandle                PROTO :DWORD
GetCurrentProcess           PROTO
DuplicateHandle             PROTO :DWORD,:DWORD,:DWORD,:DWORD,:DWORD,:DWORD,:DWORD
WaitForSingleObject         PROTO :DWORD,:DWORD
GetExitCodeProcess          PROTO :DWORD,:DWORD
CreateFileW                 PROTO :DWORD,:DWORD,:DWORD,:DWORD,:DWORD,:DWORD,:DWORD
CommandLineToArgvW          PROTO :DWORD,:DWORD
LocalFree                   PROTO :DWORD
wcslen_p                   PROTO :DWORD
wcscpy_p                   PROTO :DWORD,:DWORD
wcscat_p                   PROTO :DWORD,:DWORD

.const
str_nulDevice dw 'N','U','L',0
str_batchPrefix dw 'c','m','d','.','e','x','e',' ','/','d',' ','/','s',' ','/','c',' ','"',0
str_batchSuffix dw '"',0

; ==============================================================================
; UNINITIALIZED DATA SECTION
; ==============================================================================
.data?
sysDirBuf       dw 260 dup(?)   ; Child working directory
batchCmdBuf     dw 32768 dup(?) ; Private cmd.exe wrapper for .cmd/.bat

; ==============================================================================
; CODE SECTION
; ==============================================================================
.code

PrepareBatchCommand proc uses esi edi lpCmd:DWORD
    LOCAL argc:DWORD
    LOCAL result:DWORD

    ; Strip outer quotes if present
    mov eax, lpCmd
    cmp word ptr [eax], 0022h
    jne pbc_no_strip
    invoke wcslen_p, eax
    test eax, eax
    jz pbc_no_strip
    mov ecx, lpCmd
    lea edx, [ecx + eax*2 - 2]
    cmp word ptr [edx], 0022h
    jne pbc_no_strip
    add ecx, 2                  ; skip opening quote
    mov lpCmd, ecx              ; update parameter in-place
    mov word ptr [edx], 0       ; strip closing quote
pbc_no_strip:

    mov eax, lpCmd
    mov result, eax
    invoke CommandLineToArgvW, lpCmd, addr argc
    test eax, eax
    jz pbc_return
    mov esi, eax                         ; LocalFree-able argv
    mov edi, [esi]                       ; argv[0]
    invoke wcslen_p, edi
    cmp eax, 4
    jb pbc_free
    lea edi, [edi+eax*2]                 ; end of first token
    cmp word ptr [edi-8], '.'
    jne pbc_free
    movzx eax, word ptr [edi-6]
    or eax, 20h
    cmp eax, 'c'
    je pbc_cmd
    cmp eax, 'b'
    jne pbc_free
    movzx eax, word ptr [edi-4]
    or eax, 20h
    cmp eax, 'a'
    jne pbc_free
    movzx eax, word ptr [edi-2]
    or eax, 20h
    cmp eax, 't'
    jne pbc_free
    jmp pbc_wrap
pbc_cmd:
    movzx eax, word ptr [edi-4]
    or eax, 20h
    cmp eax, 'm'
    jne pbc_free
    movzx eax, word ptr [edi-2]
    or eax, 20h
    cmp eax, 'd'
    jne pbc_free
pbc_wrap:
    invoke wcslen_p, lpCmd
    cmp eax, 32748
    ja pbc_free
    invoke wcscpy_p, offset batchCmdBuf, offset str_batchPrefix
    invoke wcscat_p, offset batchCmdBuf, lpCmd
    invoke wcscat_p, offset batchCmdBuf, offset str_batchSuffix
    mov result, offset batchCmdBuf
pbc_free:
    invoke LocalFree, esi
pbc_return:
    mov eax, result
    ret
PrepareBatchCommand endp

; ==============================================================================
; RunAsTrustedInstaller - Execute Command with TrustedInstaller Privileges
;
; Purpose: Creates a new process running with TrustedInstaller security context.
;          This is the core function that enables privilege escalation for
;          administrative tasks. Supports both console inheritance (CLI mode)
;          and new console creation (GUI mode).
;
; Parameters:
;   cmdLine        - Wide character command line string to execute
;   useNewConsole  - Flag: 0 = inherit console, 1 = create new console
;
; Returns:
;   EAX = 1 on success, 0 on failure
;
; Process flow:
;   1. Acquire TrustedInstaller token via GetTIToken
;   2. Initialize STARTUPINFO structure
;   3. Configure console mode (inherit or new)
;   4. Create environment block for token
;   5. Preserve caller's current directory (System32 only as fallback)
;   6. Create process with CreateProcessWithTokenW
;   7. Clean up handles and environment block
;   8. Return success/failure status
;
; Console modes:
;   Inherit mode (CLI): Redirects stdin/stdout/stderr to parent process
;   New console mode (GUI): Creates separate console window
;
; Registers used: EBX, ESI, EDI (preserved)
; ==============================================================================
RunAsTrustedInstaller proc uses ebx esi edi cmdLine:DWORD, useNewConsole:DWORD
    LOCAL hToken:DWORD                      ; TrustedInstaller token handle
    LOCAL hEnv:DWORD                        ; Environment block handle
    LOCAL startupInfo[17]:DWORD             ; STARTUPINFO structure (68 bytes)
    LOCAL procInfo[4]:DWORD                 ; PROCESS_INFORMATION structure (16 bytes)
    LOCAL dwCreationFlags:DWORD             ; Process creation flags
    LOCAL dupIn:DWORD                       ; TRUE if stdin was duplicated
    LOCAL dupOut:DWORD                      ; TRUE if stdout was duplicated
    LOCAL dupErr:DWORD                      ; TRUE if stderr was duplicated
    LOCAL childExit:DWORD                   ; result returned by waited child
    LOCAL hNullInput:DWORD                  ; inheritable NUL used by relay
    LOCAL nullSA[3]:DWORD                   ; SECURITY_ATTRIBUTES (x86)

    invoke PrepareBatchCommand, cmdLine
    mov cmdLine, eax

    ; Acquire TrustedInstaller security token
    invoke GetTIToken
    test eax, eax
    jz rp_no_token                          ; Token acquisition failed
    mov hToken, eax

    ; Initialize STARTUPINFO structure to zero
    lea edi, startupInfo
    xor eax, eax
    mov ecx, 17                             ; 17 DWORDs = 68 bytes
    rep stosd                               ; Zero fill structure
    mov dword ptr [startupInfo], STARTUPINFOW_SIZE  ; Set cb (structure size)
    mov dupIn, 0
    mov dupOut, 0
    mov dupErr, 0
    mov g_childExitCode, 1                   ; failure-safe default
    mov hNullInput, 0

    ; Relay mode: elevated child redirects stdout/stderr to separate temp files.
    cmp g_relayHandle, 0
    jne rp_relay_mode

    ; Check console mode flag
    cmp useNewConsole, 0
    jne rp_new_console                      ; Jump if new console requested

    ; --- Console inheritance mode (CLI) ---
    ; This mode is used when running from command line to see output
    ; Standard handles are inherited from parent process
    invoke GetCurrentProcess
    mov esi, eax
    
    ; Get standard input handle
    invoke GetStdHandle, STD_INPUT_HANDLE
    mov dword ptr [startupInfo+56], eax     ; hStdInput field
    test eax, eax
    jz rp_dup_stdout
    cmp eax, -1
    je rp_dup_stdout
    lea edx, [startupInfo+56]
    invoke DuplicateHandle, esi, eax, esi, edx, 0, 1, DUPLICATE_SAME_ACCESS
    test eax, eax
    jz rp_dup_stdout
    mov dupIn, 1

rp_dup_stdout:
    ; Get standard output handle
    invoke GetStdHandle, STD_OUTPUT_HANDLE
    mov dword ptr [startupInfo+60], eax     ; hStdOutput field
    test eax, eax
    jz rp_dup_stderr
    cmp eax, -1
    je rp_dup_stderr
    lea edx, [startupInfo+60]
    invoke DuplicateHandle, esi, eax, esi, edx, 0, 1, DUPLICATE_SAME_ACCESS
    test eax, eax
    jz rp_dup_stderr
    mov dupOut, 1

rp_dup_stderr:
    ; Get standard error handle
    invoke GetStdHandle, STD_ERROR_HANDLE
    mov dword ptr [startupInfo+64], eax     ; hStdError field
    test eax, eax
    jz rp_stdio_ready
    cmp eax, -1
    je rp_stdio_ready
    lea edx, [startupInfo+64]
    invoke DuplicateHandle, esi, eax, esi, edx, 0, 1, DUPLICATE_SAME_ACCESS
    test eax, eax
    jz rp_stdio_ready
    mov dupErr, 1

rp_stdio_ready:
    ; Set dwFlags to use standard handles
    mov dword ptr [startupInfo+44], STARTF_USESTDHANDLES
    
    ; Set creation flags for Unicode environment only
    mov dwCreationFlags, CREATE_UNICODE_ENVIRONMENT
    jmp rp_setup_env

rp_relay_mode:
    ; STARTF_USESTDHANDLES requires three valid handles. A real NUL handle
    ; supplies deterministic EOF without making relay commands interactive.
    mov nullSA[0], 12
    mov nullSA[4], 0
    mov nullSA[8], 1
    invoke CreateFileW, offset str_nulDevice, GENERIC_READ, FILE_SHARE_READ, addr nullSA, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0
    cmp eax, -1
    je rp_fail
    mov hNullInput, eax
    mov dword ptr [startupInfo+56], eax
    mov eax, g_relayHandle
    mov dword ptr [startupInfo+60], eax      ; hStdOutput = relay file
    mov eax, g_relayErrHandle
    test eax, eax
    jnz @F
    mov eax, g_relayHandle                   ; backward-compatible fallback
@@:
    mov dword ptr [startupInfo+64], eax      ; hStdError = separate relay file
    mov dword ptr [startupInfo+44], STARTF_USESTDHANDLES
    mov dwCreationFlags, CREATE_NO_WINDOW or CREATE_UNICODE_ENVIRONMENT
    jmp rp_setup_env

rp_new_console:
    ; --- New console mode (GUI) ---
    ; This mode creates a separate console window for the process
    
    ; Set dwFlags to use wShowWindow field
    mov dword ptr [startupInfo+44], STARTF_USESHOWWINDOW
    
    ; Set wShowWindow to normal (show window)
    mov word ptr [startupInfo+48], SW_SHOWNORMAL
    
    ; Set creation flags for new console + Unicode environment
    mov dwCreationFlags, CREATE_NEW_CONSOLE or CREATE_UNICODE_ENVIRONMENT

rp_setup_env:
    ; Initialize PROCESS_INFORMATION structure
    lea edi, procInfo
    xor eax, eax
    mov ecx, 4                              ; 4 DWORDs = 16 bytes
    rep stosd                               ; Zero fill
    
    ; Create environment block from token (inherits user profile settings)
    mov hEnv, 0                             ; Initialize to NULL
    invoke CreateEnvironmentBlock, addr hEnv, hToken, 0
    
    ; Preserve the caller's working directory so relative .cmd/.bat paths work.
    ; Fall back to System32 if it does not fit in the MAX_PATH buffer.
    invoke GetCurrentDirectoryW, 260, offset sysDirBuf
    test eax, eax
    jz rp_workdir_fallback
    cmp eax, 260
    jb rp_workdir_ready
rp_workdir_fallback:
    invoke GetSystemDirectoryW, offset sysDirBuf, 260
rp_workdir_ready:
    
    ; Create process with TrustedInstaller token
    ; Parameters:
    ;   hToken              - TrustedInstaller token
    ;   dwLogonFlags        - 1 (LOGON_WITH_PROFILE)
    ;   lpApplicationName   - NULL (use command line)
    ;   lpCommandLine       - Command to execute
    ;   dwCreationFlags     - Console + Unicode environment flags
    ;   lpEnvironment       - Environment block
    ;   lpCurrentDirectory  - Caller's current directory
    ;   lpStartupInfo       - Startup configuration
    ;   lpProcessInformation- Receives process/thread handles
    invoke CreateProcessWithTokenW, hToken, 1, 0, cmdLine, dwCreationFlags, hEnv, offset sysDirBuf, addr startupInfo, addr procInfo
    mov ebx, eax                            ; Save result
    
    ; Destroy environment block if it was created
    mov eax, hEnv
    test eax, eax
    jz @F                                   ; NULL, skip destruction
    invoke DestroyEnvironmentBlock, hEnv
@@:
    mov eax, ebx                            ; Restore CreateProcessWithTokenW result
    test eax, eax
    jz rp_fail                              ; Process creation failed

    ; Wait for CLI/relay children so redirected output is complete on return.
    cmp g_relayHandle, 0
    jne rp_wait_child
    cmp useNewConsole, 0
    jne rp_skip_wait

rp_wait_child:
    mov eax, [procInfo]                     ; hProcess
    test eax, eax
    jz rp_skip_wait
    invoke WaitForSingleObject, eax, 0FFFFFFFFh
    invoke GetExitCodeProcess, dword ptr [procInfo], addr childExit
    test eax, eax
    jz rp_skip_wait
    mov eax, childExit
    mov g_childExitCode, eax

rp_skip_wait:

    ; Close inheritable duplicates made only for child startup.
    cmp dupIn, 0
    je rp_close_dup_out
    invoke CloseHandle, dword ptr [startupInfo+56]
rp_close_dup_out:
    cmp dupOut, 0
    je rp_close_dup_err
    invoke CloseHandle, dword ptr [startupInfo+60]
rp_close_dup_err:
    cmp dupErr, 0
    je rp_close_proc_handles
    invoke CloseHandle, dword ptr [startupInfo+64]
    
rp_close_proc_handles:
    cmp hNullInput, 0
    je @F
    invoke CloseHandle, hNullInput
    mov hNullInput, 0
@@:
    ; Close process handle (we don't need to wait for it)
    mov eax, [procInfo]                     ; hProcess
    test eax, eax
    jz rp_skip_hp                           ; NULL handle, skip
    invoke CloseHandle, eax
rp_skip_hp:
    
    ; Close thread handle
    mov eax, [procInfo+4]                   ; hThread
    test eax, eax
    jz rp_skip_ht                           ; NULL handle, skip
    invoke CloseHandle, eax
rp_skip_ht:
    
    mov eax, 1                              ; Return success
    ret

rp_fail:
    cmp hNullInput, 0
    je @F
    invoke CloseHandle, hNullInput
    mov hNullInput, 0
@@:
    cmp dupIn, 0
    je rp_fail_dup_out
    invoke CloseHandle, dword ptr [startupInfo+56]
rp_fail_dup_out:
    cmp dupOut, 0
    je rp_fail_dup_err
    invoke CloseHandle, dword ptr [startupInfo+60]
rp_fail_dup_err:
    cmp dupErr, 0
    je rp_no_token
    invoke CloseHandle, dword ptr [startupInfo+64]
rp_no_token:
    xor eax, eax                            ; Return failure
    ret
RunAsTrustedInstaller endp

end
