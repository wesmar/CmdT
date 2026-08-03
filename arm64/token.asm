; ==============================================================================
; CMDT - Run as TrustedInstaller
; Token Manipulation and Privilege Management Module
;
; Author: Marek Wesolowski (wesmar)
; Ported to: Native ARM64 (armasm64) for Windows on ARM
;
; Purpose: Handles Windows security token operations including privilege
;          elevation, SYSTEM impersonation via winlogon.exe token
;          duplication, TrustedInstaller service management, and
;          TrustedInstaller token acquisition with 30-second caching.
;
; Exported procedures:
;   BuildPrivilegeName(baseName, outBuf) - Constructs "Se<name>Privilege"
;   EnablePrivilege(index)               - Enables privilege in current token
;   GetProcessIdByName(name)             - Finds PID via toolhelp32 snapshot
;   ImpersonateSystem()                  - Impersonates SYSTEM via winlogon
;   StartTIService()                     - Starts TrustedInstaller service
;   GetTIToken()                         - Obtains cached TI token
;
; ARM64 Port Notes:
;   - x64 shadow space (32 bytes before each CALL) eliminated; ARM64 ABI
;     passes first 8 args in x0-x7 with no shadow space requirement.
;   - x64 [rsp+N] stack-relative addressing replaced with [sp+#N] using
;     explicit frame layout documented per function.
;   - Callee-saved registers: x19-x28 (vs rbx,rsi,rdi,r12-r15 on x64).
;   - All branch conditions faithfully mirror the x64 originals.
; ==============================================================================

; ==============================================================================
; EXTERNAL REFERENCES
; ==============================================================================

; Windows API
    IMPORT GetCurrentProcess
    IMPORT OpenProcessToken
    IMPORT LookupPrivilegeValueW
    IMPORT AdjustTokenPrivileges
    IMPORT GetLastError
    IMPORT OpenProcess
    IMPORT DuplicateTokenEx
    IMPORT ImpersonateLoggedOnUser
    IMPORT RevertToSelf
    IMPORT CloseHandle
    IMPORT GetTickCount
    IMPORT Sleep
    IMPORT OpenSCManagerW
    IMPORT OpenServiceW
    IMPORT QueryServiceStatusEx
    IMPORT StartServiceW
    IMPORT CloseServiceHandle
    IMPORT CreateToolhelp32Snapshot
    IMPORT Process32FirstW
    IMPORT Process32NextW

; In-project helpers
    IMPORT wcscpy_p
    IMPORT wcscat_p
    IMPORT wcscmp_ci
    IMPORT DecryptWideStr

; External data
    IMPORT privPrefix
    IMPORT privSuffix
    IMPORT g_privTable
    IMPORT g_decryptBuf
    IMPORT g_cachedToken
    IMPORT g_tokenTime

; ==============================================================================
; WINDOWS API CONSTANTS
; ==============================================================================

TOKEN_QUERY_ADJUST        EQU 0x0028   ; TOKEN_QUERY | TOKEN_ADJUST_PRIVILEGES
TOKEN_DUP_QUERY           EQU 0x000A   ; TOKEN_DUPLICATE | TOKEN_QUERY
TOKEN_DUP_QUERY_ADJ       EQU 0x002A   ; TOKEN_DUPLICATE | TOKEN_QUERY | TOKEN_ADJUST_PRIVILEGES
TOKEN_TYPE_IMPERSONATION  EQU 2
TOKEN_TYPE_PRIMARY        EQU 1
SECURITY_IMPERSONATION_LVL EQU 2
MAXIMUM_ALLOWED           EQU 0x02000000
PROCESS_QUERY_DUP         EQU 0x0440   ; PROCESS_QUERY_INFORMATION | PROCESS_DUP_HANDLE
PROCESS_QUERY_INFORMATION EQU 0x0400
SE_PRIVILEGE_ENABLED      EQU 0x00000002
TH32CS_SNAPPROCESS        EQU 0x00000002
PROCESSENTRY32W_SIZE      EQU 568
INVALID_HANDLE_VALUE      EQU 0xFFFFFFFFFFFFFFFF
SC_MANAGER_CONNECT        EQU 0x0001
SERVICE_QS                EQU 0x0014   ; SERVICE_QUERY_STATUS | SERVICE_START
SC_STATUS_PROCESS_INFO    EQU 0
SERVICE_STATUS_PROCESS_SIZE EQU 36
SERVICE_RUNNING           EQU 0x00000004
SERVICE_STOPPED           EQU 0x00000001

; ==============================================================================
; CONSTANT STRING DATA - OBFUSCATED (READ-ONLY)
; ==============================================================================
    AREA |.const|, DATA, READONLY, ALIGN=3

; Process name to impersonate (encrypted: "winlogon.exe")
str_winlogon_enc
    DCB 0xDD,0xAA,0xC3,0xAA,0xC4,0xAA,0xC6,0xAA,0xC5,0xAA,0xCD,0xAA,0xC5,0xAA,0xC4,0xAA
    DCB 0x84,0xAA,0xCF,0xAA,0xD2,0xAA,0xCF,0xAA,0xAA,0xAA,0

; TrustedInstaller service name (encrypted)
str_tiSvcName_enc
    DCB 0xFE,0xAA,0xD8,0xAA,0xDF,0xAA,0xD9,0xAA,0xDE,0xAA,0xCF,0xAA,0xCE,0xAA,0xE3,0xAA
    DCB 0xC4,0xAA,0xD9,0xAA,0xDE,0xAA,0xCB,0xAA,0xC6,0xAA,0xC6,0xAA,0xCF,0xAA,0xD8,0xAA
    DCB 0xAA,0xAA,0

; ==============================================================================
; CODE SECTION
; ==============================================================================
    AREA |.text|, CODE, READONLY, ALIGN=4

    EXPORT BuildPrivilegeName
    EXPORT EnablePrivilege
    EXPORT GetProcessIdByName
    EXPORT ImpersonateSystem
    EXPORT StartTIService
    EXPORT GetTIToken

; ==============================================================================
; BuildPrivilegeName - Construct Full Privilege Name String
;
; Purpose: Builds a complete Windows privilege name by combining:
;          "Se" + privilege_base_name + "Privilege"
;          Example: "Se" + "Debug" + "Privilege" = "SeDebugPrivilege"
;
; Parameters:
;   x0 = Pointer to privilege base name (e.g., "Debug")
;   x1 = Pointer to output buffer
;
; Returns:
;   x0 = Pointer to output buffer (same as x1 input)
; ==============================================================================
BuildPrivilegeName PROC
    STP x29, x30, [sp, #-16]!
    MOV x29, sp
    STP x19, x20, [sp, #-16]!

    MOV x19, x0                 ; x19 = privilege base name
    MOV x20, x1                 ; x20 = output buffer

    ; Copy "Se" prefix
    ADRP x1, privPrefix
    ADD x1, x1, privPrefix
    MOV x0, x20
    BL wcscpy_p

    ; Append privilege base name
    MOV x0, x20
    MOV x1, x19
    BL wcscat_p

    ; Append "Privilege" suffix
    ADRP x1, privSuffix
    ADD x1, x1, privSuffix
    MOV x0, x20
    BL wcscat_p

    MOV x0, x20                 ; Return output buffer pointer

    LDP x19, x20, [sp], #16
    LDP x29, x30, [sp], #16
    RET
BuildPrivilegeName ENDP

; ==============================================================================
; EnablePrivilege - Enable a Specific Privilege in Current Process Token
;
; Purpose: Enables a Windows privilege in the current process's access token.
;
; Parameters:
;   w0 = Privilege index (0-33, corresponding to g_privTable array)
;
; Returns:
;   x0 = 1 on success, 0 on failure
;
; Stack frame: 160 bytes
;   [sp+0]   = token handle (HANDLE)
;   [sp+8]   = LUID output (8 bytes)
;   [sp+16]  = TOKEN_PRIVILEGES (16 bytes)
;              PrivilegeCount at +16, Luid at +20, Attributes at +28
;   [sp+32]  = privilege name buffer (128 bytes)
;
; Privilege indices:
;   3 = SeDebugPrivilege (required for opening winlogon.exe)
;   4 = SeImpersonatePrivilege (required for impersonation)
; ==============================================================================
EnablePrivilege PROC
    STP x29, x30, [sp, #-16]!
    MOV x29, sp
    STP x19, x20, [sp, #-16]!
    STP x21, x22, [sp, #-16]!
    SUB sp, sp, #160

    MOV w19, w0                 ; w19 = privilege index

    ; Validate privilege index (must be 0-33)
    CMP w19, #34
    B.HS ep_fail                ; Out of range

    ; Get pointer to privilege base name from g_privTable
    ; g_privTable is an array of 64-bit pointers
    ADRP x0, g_privTable
    ADD x0, x0, g_privTable
    LDR x20, [x0, x19, LSL #3] ; x20 = pointer to privilege base name

    ; Build full privilege name (e.g., "SeDebugPrivilege")
    ADD x1, sp, #32             ; x1 = output buffer on stack
    MOV x0, x20                 ; x0 = privilege base name
    BL BuildPrivilegeName

    ; Get current process pseudo-handle
    BL GetCurrentProcess
    MOV x21, x0                 ; x21 = current process handle

    ; Open process token with query and adjust privileges access
    ; OpenProcessToken(ProcessHandle, DesiredAccess, &TokenHandle)
    ADD x2, sp, #0              ; x2 = &token handle
    MOV w1, #TOKEN_QUERY_ADJUST ; x1 = desired access
    MOV x0, x21                 ; x0 = process handle
    BL OpenProcessToken
    CBZ w0, ep_fail             ; Failed to open token

    ; Look up the LUID for the privilege
    ; LookupPrivilegeValueW(NULL, privilegeName, &luid)
    ADD x2, sp, #8              ; x2 = &LUID output
    ADD x1, sp, #32             ; x1 = privilege name
    MOV x0, XZR                 ; x0 = NULL (local system)
    BL LookupPrivilegeValueW
    CBZ w0, ep_close_fail       ; Failed to lookup privilege

    ; Build TOKEN_PRIVILEGES structure
    ; PrivilegeCount = 1
    MOV w0, #1
    STR w0, [sp, #16]           ; [sp+16] = PrivilegeCount

    ; Copy LUID into TOKEN_PRIVILEGES.Luid
    LDR x0, [sp, #8]           ; Get LUID (8 bytes)
    STR x0, [sp, #20]          ; [sp+20] = Luid

    ; Set Attributes = SE_PRIVILEGE_ENABLED
    MOV w0, #SE_PRIVILEGE_ENABLED
    STR w0, [sp, #28]          ; [sp+28] = Attributes

    ; Adjust token privileges
    ; AdjustTokenPrivileges(TokenHandle, FALSE, NewState, 16, NULL, NULL)
    LDR x0, [sp, #0]           ; x0 = TokenHandle
    MOV w1, #0                  ; x1 = DisableAllPrivileges = FALSE
    ADD x2, sp, #16             ; x2 = NewState (&TOKEN_PRIVILEGES)
    MOV w3, #16                 ; x3 = BufferLength
    MOV x4, XZR                 ; x4 = PreviousState = NULL
    MOV x5, XZR                 ; x5 = ReturnLength = NULL
    BL AdjustTokenPrivileges
    CBZ w0, ep_close_fail       ; API failed

    ; Check for errors even if API returned TRUE
    ; AdjustTokenPrivileges returns TRUE even if some privileges weren't set
    BL GetLastError
    CBNZ w0, ep_close_fail     ; Error occurred

    ; Success - close token handle
    LDR x0, [sp, #0]
    BL CloseHandle
    MOV w0, #1                  ; Return success
    B ep_done

ep_close_fail
    ; Failure - close token handle before returning
    LDR x0, [sp, #0]
    BL CloseHandle

ep_fail
    MOV w0, WZR                 ; Return failure

ep_done
    ADD sp, sp, #160
    LDP x21, x22, [sp], #16
    LDP x19, x20, [sp], #16
    LDP x29, x30, [sp], #16
    RET
EnablePrivilege ENDP

; ==============================================================================
; GetProcessIdByName - Find Process ID by Executable Name
;
; Purpose: Searches for a running process by its executable name and returns
;          its process ID. Uses toolhelp32 API to enumerate processes.
;
; Parameters:
;   x0 = Pointer to process name string (e.g., "winlogon.exe")
;
; Returns:
;   x0 = Process ID if found, 0 if not found
;
; Stack frame: 592 bytes
;   [sp+0]   = saved PID (temporary)
;   [sp+8]   = PROCESSENTRY32W (568 bytes)
;              dwSize at +8, th32ProcessID at +16, szExeFile at +52
;
; Note: szExeFile is at offset 44 within PROCESSENTRY32W on x64/ARM64.
; ==============================================================================
GetProcessIdByName PROC
    STP x29, x30, [sp, #-16]!
    MOV x29, sp
    STP x19, x20, [sp, #-16]!
    SUB sp, sp, #592

    MOV x19, x0                 ; x19 = target process name

    ; Create a snapshot of all processes
    ; CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
    MOV w1, #0                  ; x1 = th32ProcessID = 0 (all processes)
    MOV w0, #TH32CS_SNAPPROCESS ; x0 = dwFlags
    BL CreateToolhelp32Snapshot

    ; Check for INVALID_HANDLE_VALUE (-1)
    MOV x20, #0xFFFFFFFFFFFFFFFF
    CMP x0, x20
    B.EQ gp_fail                ; Snapshot creation failed
    MOV x20, x0                 ; x20 = snapshot handle

    ; Initialize PROCESSENTRY32W structure size
    MOV w0, #PROCESSENTRY32W_SIZE
    STR w0, [sp, #8]            ; dwSize at [sp+8]

    ; Get first process entry
    ; Process32FirstW(hSnapshot, &ppe)
    ADD x1, sp, #8              ; x1 = &PROCESSENTRY32W
    MOV x0, x20                 ; x0 = snapshot handle
    BL Process32FirstW
    CBZ w0, gp_close_fail       ; No processes found

gp_loop
    ; Compare process name
    ; szExeFile is at offset 44 within PROCESSENTRY32W
    ; PROCESSENTRY32W starts at [sp+8], so szExeFile is at [sp+8+44] = [sp+52]
    ADD x0, sp, #52             ; x0 = current process name (szExeFile)
    MOV x1, x19                 ; x1 = target process name
    BL wcscmp_ci
    CBNZ w0, gp_match           ; Match found (non-zero = match)

gp_next
    ; Get next process entry
    ; Process32NextW(hSnapshot, &ppe)
    ADD x1, sp, #8              ; x1 = &PROCESSENTRY32W
    MOV x0, x20                 ; x0 = snapshot handle
    BL Process32NextW
    CBNZ w0, gp_loop            ; More processes to check
    B gp_close_fail             ; No more processes, not found

gp_match
    ; Process found - extract and return process ID
    ; th32ProcessID is at offset 8 within PROCESSENTRY32W
    ; PROCESSENTRY32W starts at [sp+8], so th32ProcessID is at [sp+8+8] = [sp+16]
    LDR w0, [sp, #16]          ; Get PID
    STR w0, [sp, #0]           ; Save PID temporarily

    ; Close snapshot handle
    MOV x0, x20
    BL CloseHandle

    LDR w0, [sp, #0]           ; Return PID
    B gp_done

gp_close_fail
    ; Close snapshot handle before failing
    MOV x0, x20
    BL CloseHandle

gp_fail
    MOV w0, WZR                 ; Return 0 (not found)

gp_done
    ADD sp, sp, #592
    LDP x19, x20, [sp], #16
    LDP x29, x30, [sp], #16
    RET
GetProcessIdByName ENDP

; ==============================================================================
; ImpersonateSystem - Impersonate SYSTEM Account via winlogon.exe
;
; Purpose: Impersonates the SYSTEM account by duplicating the token from
;          winlogon.exe process. This is a critical step to gain sufficient
;          privileges to access the TrustedInstaller service.
;
; Process:
;   1. Enable SeDebugPrivilege to open system processes
;   2. Find winlogon.exe process ID
;   3. Open winlogon.exe process
;   4. Open process token
;   5. Duplicate token as impersonation token
;   6. Impersonate using duplicated token
;
; Parameters: None
;
; Returns:
;   x0 = 1 on success, 0 on failure
;
; Stack frame: 96 bytes
;   [sp+0]  = original token handle (from OpenProcessToken)
;   [sp+8]  = duplicated token handle (from DuplicateTokenEx)
;   [sp+16] = process handle (from OpenProcess)
;
; Register allocation:
;   x19 = winlogon process handle
;   x20 = original token handle
;   x21 = duplicated token handle
; ==============================================================================
ImpersonateSystem PROC
    STP x29, x30, [sp, #-16]!
    MOV x29, sp
    STP x19, x20, [sp, #-16]!
    STP x21, x22, [sp, #-16]!
    SUB sp, sp, #96

    ; Enable SeDebugPrivilege (required to open winlogon.exe)
    MOV w0, #3                  ; Privilege index 3 = SeDebugPrivilege
    BL EnablePrivilege
    ; Continue even if this fails (might already be enabled)

    ; Decrypt winlogon.exe process name
    ADRP x0, str_winlogon_enc
    ADD x0, x0, str_winlogon_enc
    ADRP x1, g_decryptBuf
    ADD x1, x1, g_decryptBuf
    BL DecryptWideStr

    ; Find winlogon.exe process
    ADRP x0, g_decryptBuf
    ADD x0, x0, g_decryptBuf
    BL GetProcessIdByName
    CBZ w0, is_fail             ; winlogon.exe not found
    MOV w22, w0                 ; w22 = winlogon process ID

    ; Open winlogon.exe process
    ; OpenProcess(PROCESS_QUERY_DUP, FALSE, processId)
    MOV w2, w22                 ; x2 = process ID
    MOV w1, #0                  ; x1 = don't inherit handle
    MOV w0, #PROCESS_QUERY_DUP  ; x0 = desired access
    BL OpenProcess
    CBZ x0, is_fail             ; Failed to open process
    MOV x19, x0                 ; x19 = process handle

    ; Open process token
    ; OpenProcessToken(processHandle, TOKEN_DUP_QUERY, &tokenHandle)
    ADD x2, sp, #0              ; x2 = &token handle [sp+0]
    MOV w1, #TOKEN_DUP_QUERY    ; x1 = desired access
    MOV x0, x19                 ; x0 = process handle
    BL OpenProcessToken
    CBZ w0, is_close_proc       ; Failed to open token
    LDR x20, [sp, #0]          ; x20 = original token handle

    ; Duplicate token as impersonation token
    ; DuplicateTokenEx(existingToken, MAXIMUM_ALLOWED, NULL,
    ;                  SecurityImpersonation, TokenImpersonation, &newToken)
    ADD x5, sp, #8              ; x5 = &newToken [sp+8]
    MOV w4, #TOKEN_TYPE_IMPERSONATION ; x4 = TokenType
    MOV w3, #SECURITY_IMPERSONATION_LVL ; x3 = ImpersonationLevel
    MOV x2, XZR                 ; x2 = lpTokenAttributes = NULL
    MOV w1, #MAXIMUM_ALLOWED    ; x1 = dwDesiredAccess
    MOV x0, x20                 ; x0 = ExistingTokenHandle
    BL DuplicateTokenEx
    CBZ w0, is_close_sys        ; Duplication failed
    LDR x21, [sp, #8]          ; x21 = duplicated token handle

    ; Impersonate using the duplicated token
    ; ImpersonateLoggedOnUser(duplicatedToken)
    MOV x0, x21
    BL ImpersonateLoggedOnUser
    CBZ w0, is_close_dup        ; Impersonation failed

    ; Success - clean up handles and return
    MOV x0, x21                 ; Close duplicated token
    BL CloseHandle
    MOV x0, x20                 ; Close original token
    BL CloseHandle
    MOV x0, x19                 ; Close process handle
    BL CloseHandle
    MOV w0, #1                  ; Return success
    B is_done

is_close_dup
    ; Cleanup: close duplicated token
    MOV x0, x21
    BL CloseHandle

is_close_sys
    ; Cleanup: close original token
    MOV x0, x20
    BL CloseHandle

is_close_proc
    ; Cleanup: close process handle
    MOV x0, x19
    BL CloseHandle

is_fail
    MOV w0, WZR                 ; Return failure

is_done
    ADD sp, sp, #96
    LDP x21, x22, [sp], #16
    LDP x19, x20, [sp], #16
    LDP x29, x30, [sp], #16
    RET
ImpersonateSystem ENDP

; ==============================================================================
; StartTIService - Start the TrustedInstaller Service
;
; Purpose: Ensures the TrustedInstaller service is running. Opens the service,
;          checks its status, and starts it if necessary. Waits for the service
;          to reach running state before returning.
;
; Process:
;   1. Open Service Control Manager
;   2. Open TrustedInstaller service
;   3. Query service status
;   4. If stopped, start the service
;   5. Wait up to 2 seconds for service to start (10 retries x 200ms)
;   6. Return process ID of running service
;
; Parameters: None
;
; Returns:
;   x0 = Process ID of TrustedInstaller service if running, 0 on failure
;
; Stack frame: 144 bytes
;   [sp+0]  = SERVICE_STATUS_PROCESS structure (36 bytes)
;   [sp+40] = bytes needed (DWORD, for QueryServiceStatusEx)
;
; Register allocation:
;   x19 = SCM handle
;   x20 = service handle
;   w21 = retry counter
;   w22 = service process ID
; ==============================================================================
StartTIService PROC
    STP x29, x30, [sp, #-16]!
    MOV x29, sp
    STP x19, x20, [sp, #-16]!
    STP x21, x22, [sp, #-16]!
    SUB sp, sp, #144

    ; Open Service Control Manager
    ; OpenSCManagerW(NULL, NULL, SC_MANAGER_CONNECT)
    MOV w2, #SC_MANAGER_CONNECT ; x2 = desired access
    MOV x1, XZR                 ; x1 = database name (NULL)
    MOV x0, XZR                 ; x0 = machine name (NULL)
    BL OpenSCManagerW
    CBZ x0, ss_fail             ; Failed to open SCM
    MOV x19, x0                 ; x19 = SCM handle

    ; Decrypt TrustedInstaller service name
    ADRP x0, str_tiSvcName_enc
    ADD x0, x0, str_tiSvcName_enc
    ADRP x1, g_decryptBuf
    ADD x1, x1, g_decryptBuf
    BL DecryptWideStr

    ; Open TrustedInstaller service
    ; OpenServiceW(hSCManager, serviceName, SERVICE_QS)
    MOV w2, #SERVICE_QS         ; x2 = query status + start access
    ADRP x1, g_decryptBuf
    ADD x1, x1, g_decryptBuf
    MOV x0, x19                 ; x0 = SCM handle
    BL OpenServiceW
    CBZ x0, ss_close_scm        ; Failed to open service
    MOV x20, x0                 ; x20 = service handle

    ; Query service status
    ; QueryServiceStatusEx(hService, SC_STATUS_PROCESS_INFO, &ssp, size, &needed)
    ADD x4, sp, #40             ; x4 = &bytesNeeded [sp+40]
    MOV w3, #SERVICE_STATUS_PROCESS_SIZE ; x3 = cbBufSize
    ADD x2, sp, #0              ; x2 = &SERVICE_STATUS_PROCESS [sp+0]
    MOV w1, #SC_STATUS_PROCESS_INFO ; x1 = InfoLevel
    MOV x0, x20                 ; x0 = hService
    BL QueryServiceStatusEx
    CBZ w0, ss_close_svc        ; Query failed

    ; Check current service state (dwCurrentState at offset 4 in SSP)
    LDR w0, [sp, #4]           ; dwCurrentState
    CMP w0, #SERVICE_RUNNING
    B.EQ ss_running             ; Already running
    CMP w0, #SERVICE_STOPPED
    B.NE ss_close_svc           ; Service in unexpected state

    ; Service is stopped - start it
    ; StartServiceW(hService, 0, NULL)
    MOV x2, XZR                 ; x2 = lpServiceArgVectors = NULL
    MOV w1, #0                  ; x1 = dwNumServiceArgs = 0
    MOV x0, x20                 ; x0 = hService
    BL StartServiceW
    ; Continue even if start fails (might already be starting)

    ; Retry loop: Wait for service to reach running state
    ; Maximum 10 attempts with 200ms delay = ~2 seconds total
    MOV w21, #10                ; w21 = retry counter

ss_retry
    ; Sleep for 200 milliseconds
    MOV w0, #200
    BL Sleep

    ; Query service status again
    ADD x4, sp, #40
    MOV w3, #SERVICE_STATUS_PROCESS_SIZE
    ADD x2, sp, #0
    MOV w1, #SC_STATUS_PROCESS_INFO
    MOV x0, x20
    BL QueryServiceStatusEx
    CBZ w0, ss_close_svc        ; Query failed

    ; Check if service is now running
    LDR w0, [sp, #4]           ; dwCurrentState
    CMP w0, #SERVICE_RUNNING
    B.EQ ss_running             ; Service started successfully

    ; Decrement retry counter and try again if not exhausted
    SUBS w21, w21, #1
    B.NE ss_retry               ; More retries remaining
    B ss_close_svc              ; Timeout - service didn't start

ss_running
    ; Service is running - extract process ID
    ; dwProcessId is at offset 28 in SERVICE_STATUS_PROCESS
    LDR w22, [sp, #28]         ; w22 = service process ID

    ; Close service handle
    MOV x0, x20
    BL CloseServiceHandle

    ; Close SCM handle
    MOV x0, x19
    BL CloseServiceHandle

    MOV w0, w22                 ; Return process ID
    B ss_done

ss_close_svc
    ; Cleanup: close service handle
    MOV x0, x20
    BL CloseServiceHandle

ss_close_scm
    ; Cleanup: close SCM handle
    MOV x0, x19
    BL CloseServiceHandle

ss_fail
    MOV w0, WZR                 ; Return 0 (failure)

ss_done
    ADD sp, sp, #144
    LDP x21, x22, [sp], #16
    LDP x19, x20, [sp], #16
    LDP x29, x30, [sp], #16
    RET
StartTIService ENDP

; ==============================================================================
; GetTIToken - Obtain TrustedInstaller Access Token
;
; Purpose: Core function that obtains a TrustedInstaller token with all
;          privileges enabled. Implements token caching for performance.
;
; Process overview:
;   1. Check cache (token valid for 30 seconds)
;   2. If expired or missing:
;      a. Enable SeDebugPrivilege and SeImpersonatePrivilege
;      b. Impersonate SYSTEM account
;      c. Start TrustedInstaller service
;      d. Open TrustedInstaller process
;      e. Duplicate its token as primary
;      f. Enable all 34 privileges on the duplicated token
;      g. Revert impersonation
;      h. Cache the token
;
; Parameters: None
;
; Returns:
;   x0 = TrustedInstaller token handle on success, 0 on failure
;
; Stack frame: 192 bytes
;   [sp+0]   = original TI token handle (from OpenProcessToken)
;   [sp+8]   = LUID output (8 bytes)
;   [sp+16]  = TOKEN_PRIVILEGES (16 bytes)
;   [sp+32]  = privilege name buffer (128 bytes)
;   [sp+160] = privilege loop counter (4 bytes)
;
; Register allocation:
;   x19 = current tick count
;   x20 = privilege loop index
;   x21 = TI process handle
;   x22 = original TI token handle
;   x23 = duplicated TI token handle
;   w24 = TI process ID
;   x25 = privilege base name pointer
;
; Cache behavior:
;   - Cached token is valid for 30 seconds (30000 milliseconds)
;   - Old token is closed when cache expires
;   - Cache timestamp stored in g_tokenTime
;   - Cached handle stored in g_cachedToken
; ==============================================================================
GetTIToken PROC
    STP x29, x30, [sp, #-16]!
    MOV x29, sp
    STP x19, x20, [sp, #-16]!
    STP x21, x22, [sp, #-16]!
    STP x23, x24, [sp, #-16]!
    STP x25, x26, [sp, #-16]!
    SUB sp, sp, #192

    ; Get current time in milliseconds
    BL GetTickCount
    MOV w19, w0                 ; w19 = current time

    ; Check if cached token is still valid (< 30 seconds old)
    ADRP x0, g_tokenTime
    ADD x0, x0, g_tokenTime
    LDR w1, [x0]               ; w1 = cached token timestamp
    SUB w0, w19, w1            ; w0 = time difference
    MOV w1, #30000             ; 30 s expiry (too large for a CMP immediate)
    CMP w0, w1                 ; Compare with 30 seconds
    B.HI gt_expired             ; Token expired

    ; Check if we have a cached token
    ADRP x0, g_cachedToken
    ADD x0, x0, g_cachedToken
    LDR x1, [x0]
    CBZ x1, gt_expired          ; No cached token

    ; Return cached token
    MOV x0, x1
    B gt_done

gt_expired
    ; Token expired or doesn't exist - close old token if present
    ADRP x0, g_cachedToken
    ADD x0, x0, g_cachedToken
    LDR x1, [x0]
    CBZ x1, gt_no_old           ; No old token to close
    MOV x0, x1
    BL CloseHandle
    ADRP x0, g_cachedToken
    ADD x0, x0, g_cachedToken
    STR XZR, [x0]              ; Clear cached token

gt_no_old
    ; Enable required privileges for the operation
    MOV w0, #3                  ; SeDebugPrivilege
    BL EnablePrivilege
    MOV w0, #4                  ; SeImpersonatePrivilege
    BL EnablePrivilege

    ; Impersonate SYSTEM to access TrustedInstaller service
    BL ImpersonateSystem
    CBZ w0, gt_fail             ; Impersonation failed

    ; Start TrustedInstaller service and get its process ID
    BL StartTIService
    CBZ w0, gt_revert           ; Service start failed
    MOV w24, w0                 ; w24 = TrustedInstaller process ID

    ; Open TrustedInstaller process
    ; OpenProcess(PROCESS_QUERY_INFORMATION, FALSE, processId)
    MOV w2, w24                 ; x2 = process ID
    MOV w1, #0                  ; x1 = don't inherit
    MOV w0, #PROCESS_QUERY_INFORMATION ; x0 = desired access
    BL OpenProcess
    CBZ x0, gt_revert           ; Failed to open process
    MOV x21, x0                 ; x21 = TI process handle

    ; Open TrustedInstaller process token
    ; OpenProcessToken(processHandle, TOKEN_DUP_QUERY_ADJ, &tokenHandle)
    ADD x2, sp, #0              ; x2 = &token handle [sp+0]
    MOV w1, #TOKEN_DUP_QUERY_ADJ ; x1 = Query + Duplicate + Adjust
    MOV x0, x21                 ; x0 = process handle
    BL OpenProcessToken
    CBZ w0, gt_close_proc       ; Failed to open token
    LDR x22, [sp, #0]          ; x22 = original TI token handle

    ; Duplicate the TrustedInstaller token as a primary token
    ; DuplicateTokenEx(existingToken, MAXIMUM_ALLOWED, NULL,
    ;                  SecurityImpersonation, TokenPrimary, &newToken)
    ADD x5, sp, #8              ; x5 = &newToken [sp+8]
    MOV w4, #TOKEN_TYPE_PRIMARY ; x4 = TokenType = TokenPrimary
    MOV w3, #SECURITY_IMPERSONATION_LVL ; x3 = ImpersonationLevel
    MOV x2, XZR                 ; x2 = lpTokenAttributes = NULL
    MOV w1, #MAXIMUM_ALLOWED    ; x1 = dwDesiredAccess
    MOV x0, x22                 ; x0 = ExistingTokenHandle
    BL DuplicateTokenEx
    CBZ w0, gt_close_titoken    ; Duplication failed
    LDR x23, [sp, #8]          ; x23 = duplicated TI token handle

    ; Initialize privilege loop counter
    MOV w20, #0                 ; w20 = privilege index

gt_priv_loop
    ; Check if all privileges have been processed (0-33 = 34 total)
    CMP w20, #34
    B.GE gt_priv_done           ; All privileges processed

    ; Get pointer to current privilege base name from g_privTable
    ADRP x0, g_privTable
    ADD x0, x0, g_privTable
    LDR x25, [x0, x20, LSL #3] ; x25 = privilege base name

    ; Build full privilege name
    ADD x1, sp, #32             ; x1 = output buffer
    MOV x0, x25                 ; x0 = base name
    BL BuildPrivilegeName

    ; Look up privilege LUID
    ; LookupPrivilegeValueW(NULL, privilegeName, &luid)
    ADD x2, sp, #8              ; x2 = &LUID output [sp+8]
    ADD x1, sp, #32             ; x1 = privilege name
    MOV x0, XZR                 ; x0 = NULL (local system)
    BL LookupPrivilegeValueW
    CBZ w0, gt_priv_next        ; Privilege not found, skip

    ; Build TOKEN_PRIVILEGES structure for this privilege
    MOV w0, #1
    STR w0, [sp, #16]           ; PrivilegeCount = 1
    LDR x0, [sp, #8]           ; Get LUID
    STR x0, [sp, #20]          ; Set Luid
    MOV w0, #SE_PRIVILEGE_ENABLED
    STR w0, [sp, #28]          ; Attributes = SE_PRIVILEGE_ENABLED

    ; Adjust token privileges to enable this privilege
    ; AdjustTokenPrivileges(dupToken, FALSE, NewState, 16, NULL, NULL)
    MOV x0, x23                 ; x0 = duplicated token handle
    MOV w1, #0                  ; x1 = DisableAllPrivileges = FALSE
    ADD x2, sp, #16             ; x2 = NewState
    MOV w3, #16                 ; x3 = BufferLength
    MOV x4, XZR                 ; x4 = PreviousState = NULL
    MOV x5, XZR                 ; x5 = ReturnLength = NULL
    BL AdjustTokenPrivileges
    ; Continue even if this fails - some privileges might not be available

gt_priv_next
    ; Move to next privilege
    ADD w20, w20, #1
    B gt_priv_loop

gt_priv_done
    ; All privileges processed - revert impersonation
    BL RevertToSelf

    ; Close TrustedInstaller process token (we have our duplicate)
    MOV x0, x22
    BL CloseHandle

    ; Close TrustedInstaller process handle
    MOV x0, x21
    BL CloseHandle

    ; Cache the new token
    ADRP x0, g_cachedToken
    ADD x0, x0, g_cachedToken
    STR x23, [x0]              ; Store duplicated token in cache

    ; Update cache timestamp
    BL GetTickCount
    ADRP x1, g_tokenTime
    ADD x1, x1, g_tokenTime
    STR w0, [x1]

    ; Return cached token
    MOV x0, x23
    B gt_done

gt_close_titoken
    ; Cleanup: close TrustedInstaller token
    MOV x0, x22
    BL CloseHandle

gt_close_proc
    ; Cleanup: close TrustedInstaller process
    MOV x0, x21
    BL CloseHandle

gt_revert
    ; Cleanup: revert impersonation
    BL RevertToSelf

gt_fail
    MOV x0, XZR                 ; Return NULL (failure)

gt_done
    ADD sp, sp, #192
    LDP x25, x26, [sp], #16
    LDP x23, x24, [sp], #16
    LDP x21, x22, [sp], #16
    LDP x19, x20, [sp], #16
    LDP x29, x30, [sp], #16
    RET
GetTIToken ENDP

    END