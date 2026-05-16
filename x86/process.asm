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
GetStdHandle                PROTO :DWORD
GetCurrentProcess           PROTO
DuplicateHandle             PROTO :DWORD,:DWORD,:DWORD,:DWORD,:DWORD,:DWORD,:DWORD
WaitForSingleObject         PROTO :DWORD,:DWORD

; ==============================================================================
; UNINITIALIZED DATA SECTION
; ==============================================================================
.data?
sysDirBuf       dw 260 dup(?)   ; Buffer for system directory path

; ==============================================================================
; CODE SECTION
; ==============================================================================
.code

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
;   5. Get system directory for working directory
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

    ; Relay mode: elevated child redirects spawned process output to a temp file.
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
    mov dword ptr [startupInfo+56], 0        ; hStdInput = NULL
    mov eax, g_relayHandle
    mov dword ptr [startupInfo+60], eax      ; hStdOutput = relay file
    mov dword ptr [startupInfo+64], eax      ; hStdError = relay file
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
    
    ; Get Windows system directory for working directory
    invoke GetSystemDirectoryW, offset sysDirBuf, 260
    
    ; Create process with TrustedInstaller token
    ; Parameters:
    ;   hToken              - TrustedInstaller token
    ;   dwLogonFlags        - 1 (LOGON_WITH_PROFILE)
    ;   lpApplicationName   - NULL (use command line)
    ;   lpCommandLine       - Command to execute
    ;   dwCreationFlags     - Console + Unicode environment flags
    ;   lpEnvironment       - Environment block
    ;   lpCurrentDirectory  - System directory
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
