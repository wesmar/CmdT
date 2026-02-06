; ==============================================================================
; CMDT - Run as TrustedInstaller
; Security Token Management Module
; 
; Author: Marek Wesołowski (wesmar)
; Purpose: Implements acquisition and management of TrustedInstaller security
;          tokens. Handles privilege escalation, system impersonation, and
;          service management required to obtain elevated privileges.
;
; Features:
;          - TrustedInstaller token acquisition with 30-second caching
;          - System-level privilege enablement (Debug, Impersonate, etc.)
;          - Process impersonation via winlogon.exe
;          - TrustedInstaller service management (start/query)
;          - Token duplication with maximum privileges
;          - Complete privilege enablement (all 34 Windows privileges)
; ==============================================================================

.586                            ; Target 80586 instruction set
.model flat, stdcall            ; 32-bit flat memory model, stdcall convention
option casemap:none             ; Case-sensitive symbol names

include consts.inc              ; Windows API constants
include globals.inc             ; Global variable declarations

; ==============================================================================
; EXTERNAL FUNCTION PROTOTYPES - Windows API
; ==============================================================================

; Process and token management
GetCurrentProcess           PROTO
OpenProcessToken            PROTO :DWORD,:DWORD,:DWORD
LookupPrivilegeValueW       PROTO :DWORD,:DWORD,:DWORD
AdjustTokenPrivileges       PROTO :DWORD,:DWORD,:DWORD,:DWORD,:DWORD,:DWORD
GetLastError                PROTO
OpenProcess                 PROTO :DWORD,:DWORD,:DWORD
DuplicateTokenEx            PROTO :DWORD,:DWORD,:DWORD,:DWORD,:DWORD,:DWORD
ImpersonateLoggedOnUser     PROTO :DWORD
RevertToSelf                PROTO
CloseHandle                 PROTO :DWORD
GetTickCount                PROTO
Sleep                       PROTO :DWORD

; Service control manager
OpenSCManagerW              PROTO :DWORD,:DWORD,:DWORD
OpenServiceW                PROTO :DWORD,:DWORD,:DWORD
QueryServiceStatusEx        PROTO :DWORD,:DWORD,:DWORD,:DWORD,:DWORD
StartServiceW               PROTO :DWORD,:DWORD,:DWORD
CloseServiceHandle          PROTO :DWORD

; Process enumeration
CreateToolhelp32Snapshot    PROTO :DWORD,:DWORD
Process32FirstW             PROTO :DWORD,:DWORD
Process32NextW              PROTO :DWORD,:DWORD

; External privilege name components
EXTRN privPrefix:WORD                       ; "Se" prefix
EXTRN privSuffix:WORD                       ; "Privilege" suffix

; ==============================================================================
; CONSTANT STRING DATA
; ==============================================================================
.const

; Target process for system impersonation
str_winlogon    dw 'w','i','n','l','o','g','o','n','.','e','x','e',0

; TrustedInstaller service name
str_tiSvcName   dw 'T','r','u','s','t','e','d','I','n','s','t','a','l','l','e','r',0

; ==============================================================================
; CODE SECTION
; ==============================================================================
.code

; ==============================================================================
; wcscpy_t - Wide Character String Copy (Token Module Version)
;
; Purpose: Local implementation of wide string copy for internal use.
;          Copies null-terminated wide string from source to destination.
;
; Parameters:
;   dest - Destination buffer pointer
;   src  - Source string pointer
;
; Returns: None (modifies destination)
;
; Registers modified: EAX, ESI, EDI
; ==============================================================================
wcscpy_t proc dest:DWORD, src:DWORD
    push esi
    push edi
    mov edi, dest
    mov esi, src
@@:
    mov ax, word ptr [esi]                  ; Copy wide character
    mov word ptr [edi], ax
    test ax, ax                             ; Check for null terminator
    jz @F
    add esi, 2                              ; Next character (2 bytes)
    add edi, 2
    jmp @B
@@:
    pop edi
    pop esi
    ret
wcscpy_t endp

; ==============================================================================
; wcscat_t - Wide Character String Concatenate (Token Module Version)
;
; Purpose: Local implementation of wide string concatenation.
;          Appends source string to end of destination string.
;
; Parameters:
;   dest - Destination buffer (must contain null-terminated string)
;   src  - Source string to append
;
; Returns: None (modifies destination)
;
; Registers modified: EAX, ESI, EDI
; ==============================================================================
wcscat_t proc dest:DWORD, src:DWORD
    push esi
    push edi
    mov edi, dest
@@:
    cmp word ptr [edi], 0                   ; Find end of destination
    je cat_found
    add edi, 2
    jmp @B
cat_found:
    mov esi, src
cat_loop:
    mov ax, word ptr [esi]                  ; Copy source to end
    mov word ptr [edi], ax
    test ax, ax                             ; Check for null terminator
    jz cat_done
    add esi, 2
    add edi, 2
    jmp cat_loop
cat_done:
    pop edi
    pop esi
    ret
wcscat_t endp

; ==============================================================================
; BuildPrivilegeName - Construct Full Windows Privilege Name
;
; Purpose: Builds complete privilege name by combining prefix "Se", privilege
;          core name, and suffix "Privilege". For example, "Debug" becomes
;          "SeDebugPrivilege".
;
; Parameters:
;   privPart - Pointer to core privilege name (e.g., "Debug")
;   outBuf   - Output buffer for complete name (min 260 wide chars)
;
; Returns:
;   EAX = Pointer to output buffer
;
; Example:
;   Input:  "Debug"
;   Output: "SeDebugPrivilege"
;
; Registers modified: EAX, EBX, ESI, EDI
; ==============================================================================
BuildPrivilegeName proc uses ebx esi edi privPart:DWORD, outBuf:DWORD
    invoke wcscpy_t, outBuf, offset privPrefix  ; Copy "Se"
    invoke wcscat_t, outBuf, privPart           ; Append core name
    invoke wcscat_t, outBuf, offset privSuffix  ; Append "Privilege"
    mov eax, outBuf                             ; Return buffer pointer
    ret
BuildPrivilegeName endp

; ==============================================================================
; EnablePrivilege - Enable Single Windows Privilege
;
; Purpose: Enables a specific privilege in the current process token.
;          Used to acquire necessary privileges for system-level operations.
;
; Parameters:
;   index - Privilege index in g_privTable (0-33)
;
; Returns:
;   EAX = 1 on success, 0 on failure
;
; Process flow:
;   1. Validate index (must be < 34)
;   2. Get privilege name from g_privTable
;   3. Build full privilege name (Se...Privilege)
;   4. Open current process token
;   5. Lookup privilege LUID
;   6. Adjust token privileges to enable it
;   7. Verify success via GetLastError
;   8. Clean up and return status
;
; Common privileges enabled:
;   Index 3:  SeDebugPrivilege (access any process)
;   Index 4:  SeImpersonatePrivilege (impersonate security contexts)
;
; Registers modified: EAX, EBX, ESI, EDI
; ==============================================================================
EnablePrivilege proc uses ebx esi edi index:DWORD
    LOCAL hToken:DWORD                      ; Process token handle
    LOCAL privPart:DWORD                    ; Pointer to privilege core name
    LOCAL luid[2]:DWORD                     ; LUID structure (8 bytes)
    LOCAL tp[4]:DWORD                       ; TOKEN_PRIVILEGES structure (16 bytes)
    LOCAL fullPrivName[260]:WORD            ; Buffer for full privilege name

    ; Validate privilege index
    cmp index, 34
    jae ep_fail                             ; Index out of range

    ; Get pointer to privilege name from table
    mov eax, index
    shl eax, 2                              ; Multiply by 4 (DWORD size)
    add eax, offset g_privTable             ; Add table base address
    mov ebx, dword ptr [eax]                ; Get privilege name pointer
    mov privPart, ebx

    ; Build complete privilege name
    invoke BuildPrivilegeName, privPart, addr fullPrivName

    ; Open current process token with query and adjust privileges access
    invoke GetCurrentProcess
    mov ebx, eax
    invoke OpenProcessToken, ebx, TOKEN_QUERY_ADJUST, addr hToken
    test eax, eax
    jz ep_fail                              ; Token open failed

    ; Lookup privilege LUID (Locally Unique Identifier)
    invoke LookupPrivilegeValueW, 0, addr fullPrivName, addr luid
    test eax, eax
    jz ep_close_fail                        ; Privilege lookup failed

    ; Build TOKEN_PRIVILEGES structure
    mov dword ptr [tp], 1                   ; PrivilegeCount = 1
    mov eax, [luid]                         ; Copy LUID low part
    mov [tp+4], eax
    mov eax, [luid+4]                       ; Copy LUID high part
    mov [tp+8], eax
    mov dword ptr [tp+12], SE_PRIVILEGE_ENABLED ; Attributes = enabled

    ; Adjust token privileges
    invoke AdjustTokenPrivileges, hToken, 0, addr tp, 16, 0, 0
    test eax, eax
    jz ep_close_fail                        ; API call failed

    ; Check if privilege was actually enabled (GetLastError must return 0)
    invoke GetLastError
    test eax, eax
    jnz ep_close_fail                       ; Privilege not enabled

    ; Success - close token and return
    invoke CloseHandle, hToken
    mov eax, 1
    ret

ep_close_fail:
    invoke CloseHandle, hToken
ep_fail:
    xor eax, eax                            ; Return failure
    ret
EnablePrivilege endp

; ==============================================================================
; GetProcessIdByName - Find Process ID by Executable Name
;
; Purpose: Searches running processes to find one matching the given name.
;          Uses Toolhelp32 snapshot to enumerate all processes.
;
; Parameters:
;   procName - Pointer to process executable name (e.g., "winlogon.exe")
;
; Returns:
;   EAX = Process ID if found, 0 if not found or error
;
; Process flow:
;   1. Create process snapshot
;   2. Get first process entry
;   3. Compare process name (case-sensitive)
;   4. If match: return PID
;   5. If no match: get next process
;   6. Repeat until found or end of list
;   7. Clean up snapshot handle
;
; Used by: ImpersonateSystem (to find winlogon.exe)
;
; Registers modified: EAX, EBX, ESI, EDI
; ==============================================================================
GetProcessIdByName proc uses ebx esi edi procName:DWORD
    LOCAL hSnap:DWORD                       ; Snapshot handle
    LOCAL pe32[140]:DWORD                   ; PROCESSENTRY32W structure (560 bytes)
    LOCAL foundPid:DWORD                    ; Found process ID

    ; Create snapshot of all processes
    invoke CreateToolhelp32Snapshot, TH32CS_SNAPPROCESS, 0
    cmp eax, -1
    je gp_fail                              ; Snapshot creation failed
    mov hSnap, eax

    ; Initialize PROCESSENTRY32W structure
    mov dword ptr [pe32], PROCESSENTRY32W_SIZE

    ; Get first process
    invoke Process32FirstW, hSnap, addr pe32
    test eax, eax
    jz gp_close_fail                        ; No processes found

gp_loop:
    ; Compare process name (offset 36 in PROCESSENTRY32W)
    lea esi, [pe32+36]                      ; szExeFile field
    mov edi, procName
gp_cmp:
    mov ax, word ptr [esi]                  ; Compare wide characters
    mov dx, word ptr [edi]
    cmp ax, dx
    jne gp_next                             ; Names differ
    test ax, ax
    jz gp_match                             ; End of string, match found
    add esi, 2
    add edi, 2
    jmp gp_cmp

gp_next:
    ; Get next process
    invoke Process32NextW, hSnap, addr pe32
    test eax, eax
    jnz gp_loop                             ; More processes to check
    jmp gp_close_fail                       ; No more processes, not found

gp_match:
    ; Process found - get PID (offset 8 in structure)
    mov eax, [pe32+8]                       ; th32ProcessID field
    mov foundPid, eax
    invoke CloseHandle, hSnap
    mov eax, foundPid
    ret

gp_close_fail:
    invoke CloseHandle, hSnap
gp_fail:
    xor eax, eax                            ; Return 0 (not found)
    ret
GetProcessIdByName endp

; ==============================================================================
; ImpersonateSystem - Impersonate SYSTEM Security Context
;
; Purpose: Acquires SYSTEM-level privileges by impersonating the winlogon.exe
;          process. This is necessary to access TrustedInstaller service and
;          perform system-level operations.
;
; Parameters: None
;
; Returns:
;   EAX = 1 on success, 0 on failure
;
; Process flow:
;   1. Enable SeDebugPrivilege (required to access winlogon.exe)
;   2. Find winlogon.exe process ID
;   3. Open winlogon.exe process
;   4. Open process token
;   5. Duplicate token as impersonation token
;   6. Impersonate using duplicated token
;   7. Clean up handles
;   8. Return success (impersonation active until RevertToSelf)
;
; Security note: Impersonation remains active until RevertToSelf is called.
;                Caller must ensure proper cleanup.
;
; Registers modified: EAX, EBX, ESI, EDI
; ==============================================================================
ImpersonateSystem proc uses ebx esi edi
    LOCAL hProcess:DWORD                    ; Winlogon process handle
    LOCAL hToken:DWORD                      ; Winlogon token handle
    LOCAL hDupToken:DWORD                   ; Duplicated token handle

    ; Enable debug privilege (required to access system processes)
    invoke EnablePrivilege, 3               ; Index 3 = SeDebugPrivilege
    
    ; Find winlogon.exe process
    invoke GetProcessIdByName, offset str_winlogon
    test eax, eax
    jz is_fail                              ; Winlogon not found

    ; Open winlogon.exe process
    invoke OpenProcess, PROCESS_QUERY_DUP, 0, eax
    test eax, eax
    jz is_fail                              ; Process open failed
    mov hProcess, eax

    ; Open process token
    invoke OpenProcessToken, hProcess, TOKEN_DUP_QUERY, addr hToken
    test eax, eax
    jz is_close_proc                        ; Token open failed

    ; Duplicate token as impersonation token
    invoke DuplicateTokenEx, hToken, MAXIMUM_ALLOWED, 0, SECURITY_IMPERSONATION_LVL, TOKEN_TYPE_IMPERSONATION, addr hDupToken
    test eax, eax
    jz is_close_sys                         ; Duplication failed

    ; Impersonate using duplicated token
    invoke ImpersonateLoggedOnUser, hDupToken
    test eax, eax
    jz is_close_dup                         ; Impersonation failed

    ; Success - clean up handles and return
    invoke CloseHandle, hDupToken
    invoke CloseHandle, hToken
    invoke CloseHandle, hProcess
    mov eax, 1                              ; Success
    ret

is_close_dup:
    invoke CloseHandle, hDupToken
is_close_sys:
    invoke CloseHandle, hToken
is_close_proc:
    invoke CloseHandle, hProcess
is_fail:
    xor eax, eax                            ; Failure
    ret
ImpersonateSystem endp

; ==============================================================================
; StartTIService - Start TrustedInstaller Service and Get Process ID
;
; Purpose: Ensures TrustedInstaller service is running and returns its PID.
;          If service is stopped, starts it and waits up to ~2 seconds for
;          startup completion.
;
; Parameters: None
;
; Returns:
;   EAX = TrustedInstaller process ID on success, 0 on failure
;
; Process flow:
;   1. Open Service Control Manager
;   2. Open TrustedInstaller service
;   3. Query service status
;   4. If running: return PID
;   5. If stopped: start service
;   6. Retry loop: check status every 200ms (up to 10 times = 2 seconds)
;   7. Return PID when service is running
;
; Service states:
;   SERVICE_RUNNING (4) - Service is active
;   SERVICE_STOPPED (1) - Service is not running
;
; Registers modified: EAX, EBX, ESI, EDI
; ==============================================================================
StartTIService proc uses ebx esi edi
    LOCAL hSCM:DWORD                        ; Service Control Manager handle
    LOCAL hService:DWORD                    ; Service handle
    LOCAL ssp[9]:DWORD                      ; SERVICE_STATUS_PROCESS (36 bytes)
    LOCAL bytesNeeded:DWORD                 ; Bytes needed for query
    LOCAL tiPid:DWORD                       ; TrustedInstaller PID
    LOCAL retryCount:DWORD                  ; Retry counter

    ; Open Service Control Manager
    invoke OpenSCManagerW, 0, 0, SC_MANAGER_CONNECT
    test eax, eax
    jz ss_fail                              ; SCM open failed
    mov hSCM, eax

    ; Open TrustedInstaller service
    invoke OpenServiceW, hSCM, offset str_tiSvcName, SERVICE_QS
    test eax, eax
    jz ss_close_scm                         ; Service open failed
    mov hService, eax

    ; Query service status
    invoke QueryServiceStatusEx, hService, SC_STATUS_PROCESS_INFO, addr ssp, SERVICE_STATUS_PROCESS_SIZE, addr bytesNeeded
    test eax, eax
    jz ss_close_svc                         ; Query failed

    ; Check service state
    cmp dword ptr [ssp+4], SERVICE_RUNNING  ; dwCurrentState at offset 4
    je ss_running                           ; Already running
    cmp dword ptr [ssp+4], SERVICE_STOPPED
    jne ss_close_svc                        ; Unexpected state

    ; Service is stopped - start it
    invoke StartServiceW, hService, 0, 0

    ; Retry loop: wait for service to start (up to 2 seconds)
    mov retryCount, 10                      ; 10 retries × 200ms = 2000ms
ss_retry:
    invoke Sleep, 200                       ; Wait 200 milliseconds
    invoke QueryServiceStatusEx, hService, SC_STATUS_PROCESS_INFO, addr ssp, SERVICE_STATUS_PROCESS_SIZE, addr bytesNeeded
    test eax, eax
    jz ss_close_svc                         ; Query failed
    cmp dword ptr [ssp+4], SERVICE_RUNNING
    je ss_running                           ; Service started
    dec retryCount
    jnz ss_retry                            ; Try again
    jmp ss_close_svc                        ; Timeout - service didn't start

ss_running:
    ; Service is running - get process ID (offset 28 in structure)
    mov eax, [ssp+28]                       ; dwProcessId field
    mov tiPid, eax
    invoke CloseServiceHandle, hService
    invoke CloseServiceHandle, hSCM
    mov eax, tiPid
    ret

ss_close_svc:
    invoke CloseServiceHandle, hService
ss_close_scm:
    invoke CloseServiceHandle, hSCM
ss_fail:
    xor eax, eax                            ; Return 0 (failure)
    ret
StartTIService endp

; ==============================================================================
; GetTIToken - Acquire TrustedInstaller Security Token
;
; Purpose: Main function for obtaining a TrustedInstaller security token with
;          all privileges enabled. Implements 30-second token caching to avoid
;          repeated privilege escalation overhead.
;
; Parameters: None
;
; Returns:
;   EAX = TrustedInstaller token handle on success, 0 on failure
;
; Process flow:
;   1. Check cached token age (< 30 seconds = valid)
;   2. If valid cache: return cached token
;   3. If expired/missing:
;      a. Close old cached token
;      b. Enable debug and impersonate privileges
;      c. Impersonate SYSTEM security context
;      d. Start TrustedInstaller service
;      e. Open TrustedInstaller process
;      f. Open process token
;      g. Duplicate token with maximum privileges
;      h. Enable all 34 Windows privileges in token
;      i. Revert impersonation
;      j. Cache token with current timestamp
;      k. Return token handle
;
; Token caching:
;   - Tokens expire after 30 seconds (30000 milliseconds)
;   - Reduces overhead of repeated privilege escalation
;   - Cached token stored in g_cachedToken
;   - Timestamp stored in g_tokenTime
;
; Privileges enabled (all 34):
;   Includes Debug, Impersonate, TakeOwnership, Backup, Restore, etc.
;   See g_privTable in main.asm for complete list
;
; Security note: Caller is responsible for closing returned token handle
;
; Registers modified: EAX, EBX, ESI, EDI
; ==============================================================================
GetTIToken proc uses ebx esi edi
    LOCAL tiPid:DWORD                       ; TrustedInstaller process ID
    LOCAL hProcess:DWORD                    ; TrustedInstaller process handle
    LOCAL hToken:DWORD                      ; TrustedInstaller token handle
    LOCAL hDupToken:DWORD                   ; Duplicated token handle
    LOCAL privIndex:DWORD                   ; Privilege loop index
    LOCAL luid[2]:DWORD                     ; Privilege LUID
    LOCAL tp[4]:DWORD                       ; TOKEN_PRIVILEGES structure
    LOCAL currentTime:DWORD                 ; Current tick count
    LOCAL privPart:DWORD                    ; Privilege name pointer
    LOCAL fullPrivName[260]:WORD            ; Full privilege name buffer

    ; Get current time for cache expiration check
    invoke GetTickCount
    mov currentTime, eax
    
    ; Check if cached token is still valid (< 30 seconds old)
    mov ecx, g_tokenTime
    sub eax, ecx                            ; Time elapsed = current - cached
    cmp eax, 30000                          ; 30,000 ms = 30 seconds
    ja gt_expired                           ; Cache expired
    
    ; Cache is valid - return cached token
    mov eax, g_cachedToken
    test eax, eax
    jz gt_expired                           ; No cached token
    mov eax, g_cachedToken
    ret

gt_expired:
    ; Cache expired or doesn't exist - acquire new token
    
    ; Close old cached token if it exists
    mov eax, g_cachedToken
    test eax, eax
    jz gt_no_old                            ; No old token
    invoke CloseHandle, eax
    mov g_cachedToken, 0                    ; Clear cache
gt_no_old:

    ; Enable required privileges for token acquisition
    invoke EnablePrivilege, 3               ; SeDebugPrivilege
    invoke EnablePrivilege, 4               ; SeImpersonatePrivilege
    
    ; Impersonate SYSTEM to access TrustedInstaller
    invoke ImpersonateSystem
    test eax, eax
    jz gt_fail                              ; Impersonation failed
    
    ; Start TrustedInstaller service and get PID
    invoke StartTIService
    test eax, eax
    jz gt_revert                            ; Service start failed
    mov tiPid, eax
    
    ; Open TrustedInstaller process
    invoke OpenProcess, PROCESS_QUERY_INFORMATION, 0, tiPid
    test eax, eax
    jz gt_revert                            ; Process open failed
    mov hProcess, eax
    
    ; Open TrustedInstaller process token
    invoke OpenProcessToken, hProcess, TOKEN_DUP_QUERY_ADJ, addr hToken
    test eax, eax
    jz gt_close_proc                        ; Token open failed
    
    ; Duplicate token as primary token with maximum privileges
    invoke DuplicateTokenEx, hToken, MAXIMUM_ALLOWED, 0, SECURITY_IMPERSONATION_LVL, 1, addr hDupToken
    test eax, eax
    jz gt_close_titoken                     ; Duplication failed
    
    ; Enable all privileges in duplicated token
    ; Loop through all 34 privileges in g_privTable
    mov privIndex, 0
gt_priv_loop:
    cmp privIndex, 34
    jge gt_priv_done                        ; All privileges processed
    
    ; Get privilege name from table
    mov eax, privIndex
    shl eax, 2                              ; Index × 4
    add eax, offset g_privTable
    mov ebx, dword ptr [eax]
    mov privPart, ebx

    ; Build full privilege name
    invoke BuildPrivilegeName, privPart, addr fullPrivName
    
    ; Lookup privilege LUID
    invoke LookupPrivilegeValueW, 0, addr fullPrivName, addr luid
    test eax, eax
    jz gt_priv_next                         ; Privilege not found, skip
    
    ; Build TOKEN_PRIVILEGES structure
    mov dword ptr [tp], 1                   ; PrivilegeCount = 1
    mov eax, [luid]
    mov [tp+4], eax
    mov eax, [luid+4]
    mov [tp+8], eax
    mov dword ptr [tp+12], SE_PRIVILEGE_ENABLED
    
    ; Enable this privilege in token
    invoke AdjustTokenPrivileges, hDupToken, 0, addr tp, 16, 0, 0
    
gt_priv_next:
    inc privIndex
    jmp gt_priv_loop

gt_priv_done:
    ; All privileges enabled - clean up and cache token
    invoke RevertToSelf                     ; End impersonation
    invoke CloseHandle, hToken
    invoke CloseHandle, hProcess
    
    ; Cache token and timestamp
    mov eax, hDupToken
    mov g_cachedToken, eax
    invoke GetTickCount
    mov g_tokenTime, eax
    
    ; Return duplicated token handle
    mov eax, g_cachedToken
    ret

gt_close_titoken:
    invoke CloseHandle, hToken
gt_close_proc:
    invoke CloseHandle, hProcess
gt_revert:
    invoke RevertToSelf
gt_fail:
    xor eax, eax                            ; Return 0 (failure)
    ret
GetTIToken endp

end
