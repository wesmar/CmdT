; ==============================================================================
; CMDT - Run as TrustedInstaller
; Main Entry Point and Command-Line Interface Module
; 
; Author: Marek Wesołowski (wesmar)
; Purpose: Implements the main entry point for the application with support for
;          both GUI and CLI modes. Handles command-line argument parsing,
;          .lnk shortcut resolution, and process creation with TrustedInstaller
;          privileges.
;
; Features:
;          - Dual mode operation: GUI (default) or CLI (with -cli flag)
;          - Command-line argument parsing with multiple switch formats
;          - Optional new console window creation (-new flag)
;          - Windows shortcut (.lnk) file resolution
;          - Automatic command and argument extraction from shortcuts
;          - Integration with TrustedInstaller token acquisition
; ==============================================================================

.586                            ; Target 80586 instruction set
.model flat, stdcall            ; 32-bit flat memory model, stdcall convention
option casemap:none             ; Case-sensitive symbol names

include consts.inc              ; Windows API constants and structures

; ==============================================================================
; EXTERNAL FUNCTION PROTOTYPES
; ==============================================================================

; Window and GUI functions
CreateMainWindow        PROTO :DWORD

; Process and token management
RunAsTrustedInstaller   PROTO :DWORD,:DWORD

; Shortcut (.lnk) file resolution
ResolveLnkPath          PROTO :DWORD,:DWORD,:DWORD

; Wide string manipulation utilities
wcscpy_p                PROTO :DWORD,:DWORD
wcscat_p                PROTO :DWORD,:DWORD

; Windows API functions
GetModuleHandleW        PROTO :DWORD
GetMessageW             PROTO :DWORD,:DWORD,:DWORD,:DWORD
TranslateMessage        PROTO :DWORD
DispatchMessageW        PROTO :DWORD
IsDialogMessageW        PROTO :DWORD,:DWORD
ExitProcess             PROTO :DWORD
GetCommandLineW         PROTO
CommandLineToArgvW      PROTO :DWORD,:DWORD
LocalFree               PROTO :DWORD
SetFocus                PROTO :DWORD
WideCharToMultiByte     PROTO :DWORD,:DWORD,:DWORD,:DWORD,:DWORD,:DWORD,:DWORD,:DWORD

; ==============================================================================
; CONSTANT STRING DATA
; ==============================================================================
.const

; Registry key for application settings
str_regKey      dw 'S','o','f','t','w','a','r','e','\','c','m','d','t',0

; Command-line switch variations for CLI mode
str_cliSwitch1  dw '-','c','l','i',0      ; Standard Unix-style switch
str_cliSwitch2  dw '-','-','c','l','i',0  ; GNU-style long option
str_cliSwitch3  dw 'c','l','i',0          ; Bare switch without hyphen

; Switch for new console window creation
str_newSwitch   dw '-','n','e','w',0

; File extension strings
str_extLnk_m    dw '.','l','n','k',0      ; Windows shortcut extension
str_space       dw ' ',0                   ; Space character for string building

; ==============================================================================
; PRIVILEGE STRING DEFINITIONS
; Windows privileges without "Se" prefix and "Privilege" suffix
; These are combined with privPrefix and privSuffix to form complete names
; ==============================================================================

; Privilege name parts (core names only)
privStr_0  dw 'A','s','s','i','g','n','P','r','i','m','a','r','y','T','o','k','e','n',0      ; SeAssignPrimaryTokenPrivilege
privStr_1  dw 'B','a','c','k','u','p',0                                                        ; SeBackupPrivilege
privStr_2  dw 'R','e','s','t','o','r','e',0                                                    ; SeRestorePrivilege
privStr_3  dw 'D','e','b','u','g',0                                                            ; SeDebugPrivilege
privStr_4  dw 'I','m','p','e','r','s','o','n','a','t','e',0                                    ; SeImpersonatePrivilege
privStr_5  dw 'T','a','k','e','O','w','n','e','r','s','h','i','p',0                            ; SeTakeOwnershipPrivilege
privStr_6  dw 'L','o','a','d','D','r','i','v','e','r',0                                        ; SeLoadDriverPrivilege
privStr_7  dw 'S','y','s','t','e','m','E','n','v','i','r','o','n','m','e','n','t',0            ; SeSystemEnvironmentPrivilege
privStr_8  dw 'M','a','n','a','g','e','V','o','l','u','m','e',0                                ; SeManageVolumePrivilege
privStr_9  dw 'S','e','c','u','r','i','t','y',0                                                ; SeSecurityPrivilege
privStr_10 dw 'S','h','u','t','d','o','w','n',0                                                ; SeShutdownPrivilege
privStr_11 dw 'S','y','s','t','e','m','t','i','m','e',0                                        ; SeSystemtimePrivilege
privStr_12 dw 'T','c','b',0                                                                    ; SeTcbPrivilege
privStr_13 dw 'I','n','c','r','e','a','s','e','Q','u','o','t','a',0                            ; SeIncreaseQuotaPrivilege
privStr_14 dw 'A','u','d','i','t',0                                                            ; SeAuditPrivilege
privStr_15 dw 'C','h','a','n','g','e','N','o','t','i','f','y',0                                ; SeChangeNotifyPrivilege
privStr_16 dw 'U','n','d','o','c','k',0                                                        ; SeUndockPrivilege
privStr_17 dw 'C','r','e','a','t','e','T','o','k','e','n',0                                    ; SeCreateTokenPrivilege
privStr_18 dw 'L','o','c','k','M','e','m','o','r','y',0                                        ; SeLockMemoryPrivilege
privStr_19 dw 'C','r','e','a','t','e','P','a','g','e','f','i','l','e',0                        ; SeCreatePagefilePrivilege
privStr_20 dw 'C','r','e','a','t','e','P','e','r','m','a','n','e','n','t',0                    ; SeCreatePermanentPrivilege
privStr_21 dw 'S','y','s','t','e','m','P','r','o','f','i','l','e',0                            ; SeSystemProfilePrivilege
privStr_22 dw 'P','r','o','f','i','l','e','S','i','n','g','l','e','P','r','o','c','e','s','s',0 ; SeProfileSingleProcessPrivilege
privStr_23 dw 'C','r','e','a','t','e','G','l','o','b','a','l',0                                ; SeCreateGlobalPrivilege
privStr_24 dw 'T','i','m','e','Z','o','n','e',0                                                ; SeTimeZonePrivilege
privStr_25 dw 'C','r','e','a','t','e','S','y','m','b','o','l','i','c','L','i','n','k',0        ; SeCreateSymbolicLinkPrivilege
privStr_26 dw 'I','n','c','r','e','a','s','e','B','a','s','e','P','r','i','o','r','i','t','y',0 ; SeIncreaseBasePriorityPrivilege
privStr_27 dw 'R','e','m','o','t','e','S','h','u','t','d','o','w','n',0                        ; SeRemoteShutdownPrivilege
privStr_28 dw 'I','n','c','r','e','a','s','e','W','o','r','k','i','n','g','S','e','t',0        ; SeIncreaseWorkingSetPrivilege
privStr_29 dw 'R','e','l','a','b','e','l',0                                                    ; SeRelabelPrivilege
privStr_30 dw 'D','e','l','e','g','a','t','e','S','e','s','s','i','o','n','U','s','e','r','I','m','p','e','r','s','o','n','a','t','e',0 ; SeDelegateSessionUserImpersonatePrivilege
privStr_31 dw 'T','r','u','s','t','e','d','C','r','e','d','M','a','n','A','c','c','e','s','s',0 ; SeTrustedCredManAccessPrivilege
privStr_32 dw 'E','n','a','b','l','e','D','e','l','e','g','a','t','i','o','n',0                ; SeEnableDelegationPrivilege
privStr_33 dw 'S','y','n','c','A','g','e','n','t',0                                            ; SeSyncAgentPrivilege

; Privilege name prefix and suffix for building complete privilege names
privPrefix dw 'S','e',0                    ; Standard Windows privilege prefix
privSuffix dw 'P','r','i','v','i','l','e','g','e',0 ; Standard Windows privilege suffix

; ==============================================================================
; INITIALIZED DATA SECTION
; ==============================================================================
.data
    align 4                                 ; Align to 4-byte boundary for performance

; Privilege table: array of pointers to privilege name strings
; Used for iterating through all privileges when enabling them
PUBLIC g_privTable
g_privTable dd offset privStr_0,offset privStr_1,offset privStr_2,offset privStr_3,offset privStr_4,offset privStr_5
            dd offset privStr_6,offset privStr_7,offset privStr_8,offset privStr_9,offset privStr_10,offset privStr_11
            dd offset privStr_12,offset privStr_13,offset privStr_14,offset privStr_15,offset privStr_16,offset privStr_17
            dd offset privStr_18,offset privStr_19,offset privStr_20,offset privStr_21,offset privStr_22,offset privStr_23
            dd offset privStr_24,offset privStr_25,offset privStr_26,offset privStr_27,offset privStr_28,offset privStr_29
            dd offset privStr_30,offset privStr_31,offset privStr_32,offset privStr_33

; Global variables exported for use in other modules
PUBLIC g_cachedToken, g_tokenTime, g_hwndMain, g_hwndEdit, g_hwndBtn, g_hwndStatus, g_hConsoleOut, g_hInstance
PUBLIC privPrefix, privSuffix

g_cachedToken   dd 0                        ; Cached TrustedInstaller token handle
g_tokenTime     dd 0                        ; Timestamp of cached token (for expiration)
g_hwndMain      dd 0                        ; Main window handle
g_hwndEdit      dd 0                        ; Edit/ComboBox control handle
g_hwndBtn       dd 0                        ; Run button handle
g_hwndStatus    dd 0                        ; Status label handle
g_hConsoleOut   dd 0                        ; Console output handle (CLI mode)
g_useNewConsole dd 0                        ; Flag: create new console window
g_hInstance     dd 0                        ; Application instance handle

; ==============================================================================
; UNINITIALIZED DATA SECTION
; Large buffers for strings and temporary data
; ==============================================================================
.data?
PUBLIC g_cmdBuf, g_statusBuf, g_filePath, g_argsBuf, g_tempBuf

g_cmdBuf        dw 520 dup(?)               ; Command line buffer (1040 bytes)
g_statusBuf     dw 520 dup(?)               ; Status message buffer (1040 bytes)
g_filePath      dw 520 dup(?)               ; File path buffer (1040 bytes)
g_argsBuf       dw 520 dup(?)               ; Arguments buffer (1040 bytes)
g_tempBuf       dw 1040 dup(?)              ; Temporary work buffer (2080 bytes)

; ==============================================================================
; CODE SECTION
; ==============================================================================
.code

; ==============================================================================
; wcscpy_p - Wide Character String Copy
;
; Purpose: Copies a null-terminated wide character string from source to
;          destination, including the null terminator.
;
; Parameters:
;   dest - Destination buffer pointer
;   src  - Source string pointer
;
; Returns: None (modifies destination buffer)
;
; Registers modified: EAX, ESI, EDI
; ==============================================================================
wcscpy_p proc dest:DWORD, src:DWORD
    push esi                                ; Preserve source index register
    push edi                                ; Preserve destination index register
    mov edi, dest                           ; EDI = destination pointer
    mov esi, src                            ; ESI = source pointer
@@:
    mov ax, word ptr [esi]                  ; Read wide character (16-bit)
    mov word ptr [edi], ax                  ; Write to destination
    test ax, ax                             ; Check for null terminator
    jz @F                                   ; If zero, we're done
    add esi, 2                              ; Move to next wide char (2 bytes)
    add edi, 2                              ; Move destination pointer
    jmp @B                                  ; Continue loop
@@:
    pop edi                                 ; Restore registers
    pop esi
    ret
wcscpy_p endp

; ==============================================================================
; wcscat_p - Wide Character String Concatenate
;
; Purpose: Appends source string to the end of destination string. Finds the
;          null terminator in destination and copies source starting there.
;
; Parameters:
;   dest - Destination buffer pointer (must contain null-terminated string)
;   src  - Source string pointer to append
;
; Returns: None (modifies destination buffer)
;
; Registers modified: EAX, ESI, EDI
; ==============================================================================
wcscat_p proc dest:DWORD, src:DWORD
    push esi                                ; Preserve registers
    push edi
    mov edi, dest                           ; EDI = destination pointer
@@:
    cmp word ptr [edi], 0                   ; Find end of destination string
    je @F
    add edi, 2                              ; Move to next character
    jmp @B
@@:
    mov esi, src                            ; ESI = source pointer
@@:
    mov ax, word ptr [esi]                  ; Copy characters from source
    mov word ptr [edi], ax
    test ax, ax                             ; Check for null terminator
    jz @F
    add esi, 2
    add edi, 2
    jmp @B
@@:
    pop edi                                 ; Restore registers
    pop esi
    ret
wcscat_p endp

; ==============================================================================
; wcscmp_ci - Wide Character String Compare (Case-Insensitive)
;
; Purpose: Compares two wide character strings ignoring case differences.
;          Converts A-Z to a-z before comparison.
;
; Parameters:
;   str1 - First string pointer
;   str2 - Second string pointer
;
; Returns:
;   EAX = 1 if strings match (case-insensitive), 0 if different
;
; Registers modified: EAX, EDX, ESI, EDI
; ==============================================================================
wcscmp_ci proc str1:DWORD, str2:DWORD
    push esi                                ; Preserve registers
    push edi
    mov esi, str1                           ; ESI = first string
    mov edi, str2                           ; EDI = second string
wci_loop:
    mov ax, word ptr [esi]                  ; Read characters from both strings
    mov dx, word ptr [edi]
    ; Convert first character to lowercase if uppercase (A-Z → a-z)
    cmp ax, 'A'
    jb wci_skip1                            ; Below 'A', no conversion needed
    cmp ax, 'Z'
    ja wci_skip1                            ; Above 'Z', no conversion needed
    add ax, 32                              ; Convert to lowercase (ASCII offset)
wci_skip1:
    ; Convert second character to lowercase if uppercase
    cmp dx, 'A'
    jb wci_skip2
    cmp dx, 'Z'
    ja wci_skip2
    add dx, 32                              ; Convert to lowercase
wci_skip2:
    cmp ax, dx                              ; Compare normalized characters
    jne not_eq                              ; Different characters
    test ax, ax                             ; Check if both are null terminators
    jz equal                                ; End of strings, they match
    add esi, 2                              ; Move to next characters
    add edi, 2
    jmp wci_loop
equal:
    pop edi                                 ; Restore registers
    pop esi
    mov eax, 1                              ; Return 1 (strings match)
    ret
not_eq:
    pop edi
    pop esi
    xor eax, eax                            ; Return 0 (strings differ)
    ret
wcscmp_ci endp

; ==============================================================================
; strcmp_ci - ANSI String Compare (Case-Insensitive)
;
; Purpose: Compares two ANSI (single-byte) strings ignoring case differences.
;          Similar to wcscmp_ci but for 8-bit characters.
;
; Parameters:
;   str1 - First ANSI string pointer
;   str2 - Second ANSI string pointer
;
; Returns:
;   EAX = 1 if strings match (case-insensitive), 0 if different
;
; Registers modified: EAX, EDX, ESI, EDI
; ==============================================================================
strcmp_ci proc str1:DWORD, str2:DWORD
    push esi                                ; Preserve registers
    push edi
    mov esi, str1                           ; ESI = first string
    mov edi, str2                           ; EDI = second string
sci_loop:
    mov al, byte ptr [esi]                  ; Read single bytes (ANSI)
    mov dl, byte ptr [edi]
    ; Convert first character to lowercase if uppercase
    cmp al, 'A'
    jb sci_skip1
    cmp al, 'Z'
    ja sci_skip1
    add al, 32                              ; A-Z → a-z conversion
sci_skip1:
    ; Convert second character to lowercase if uppercase
    cmp dl, 'A'
    jb sci_skip2
    cmp dl, 'Z'
    ja sci_skip2
    add dl, 32
sci_skip2:
    cmp al, dl                              ; Compare normalized characters
    jne not_eq_s                            ; Different
    test al, al                             ; Check for null terminator
    jz equal_s                              ; Both strings ended, match
    inc esi                                 ; Move to next byte
    inc edi
    jmp sci_loop
equal_s:
    pop edi
    pop esi
    mov eax, 1                              ; Return 1 (match)
    ret
not_eq_s:
    pop edi
    pop esi
    xor eax, eax                            ; Return 0 (differ)
    ret
strcmp_ci endp

; ==============================================================================
; skip_spaces - Skip Leading Whitespace
;
; Purpose: Advances pointer past any leading space characters (U+0020) in a
;          wide character string.
;
; Parameters:
;   lpStr - Pointer to wide character string
;
; Returns:
;   EAX = Pointer to first non-space character
;
; Registers modified: EAX
; ==============================================================================
skip_spaces proc lpStr:DWORD
    mov eax, lpStr                          ; EAX = current position
@@:
    cmp word ptr [eax], ' '                 ; Check if current char is space
    jne @F                                  ; Not a space, we're done
    add eax, 2                              ; Skip space (2 bytes for wide char)
    jmp @B                                  ; Continue checking
@@:
    ret                                     ; Return pointer to non-space char
skip_spaces endp

; ==============================================================================
; wcslen_p - Wide Character String Length
;
; Purpose: Calculates the length of a null-terminated wide character string
;          (number of characters, not including null terminator).
;
; Parameters:
;   lpStr - Pointer to wide character string
;
; Returns:
;   EAX = Number of characters (excluding null terminator)
;
; Registers modified: EAX, ECX
; ==============================================================================
wcslen_p proc lpStr:DWORD
    mov eax, lpStr                          ; EAX = string pointer
    xor ecx, ecx                            ; ECX = character counter
@@:
    cmp word ptr [eax + ecx*2], 0           ; Check for null terminator
    je @F                                   ; Found null, exit loop
    inc ecx                                 ; Increment count
    jmp @B                                  ; Continue
@@:
    mov eax, ecx                            ; Return count in EAX
    ret
wcslen_p endp

; ==============================================================================
; wcstombs_p - Wide Character to Multibyte String Conversion
;
; Purpose: Converts a wide character (UTF-16) string to a multibyte (ANSI/UTF-8)
;          string using Windows API.
;
; Parameters:
;   wstr   - Source wide character string pointer
;   mbstr  - Destination multibyte buffer pointer
;   maxlen - Maximum length of destination buffer
;
; Returns:
;   EAX = Number of bytes written (from WideCharToMultiByte)
;
; Registers modified: EAX (API return value)
; ==============================================================================
wcstombs_p proc wstr:DWORD, mbstr:DWORD, maxlen:DWORD
    ; Call WideCharToMultiByte with code page 0 (CP_ACP), no flags
    invoke WideCharToMultiByte, 0, 0, wstr, -1, mbstr, maxlen, 0, 0
    ret
wcstombs_p endp

; ==============================================================================
; start - Main Entry Point
;
; Purpose: Application entry point. Parses command-line arguments to determine
;          operating mode (GUI vs CLI), processes .lnk shortcuts if necessary,
;          and either launches the GUI or executes commands in CLI mode.
;
; Command-line usage:
;   GUI mode (default):
;     cmdt.exe
;
;   CLI mode:
;     cmdt.exe -cli <command>           - Run command, inherit console
;     cmdt.exe -cli -new <command>      - Run command in new console window
;     cmdt.exe --cli <command>          - Alternative CLI switch format
;     cmdt.exe cli <command>            - Short form CLI switch
;
; Process flow:
;   1. Parse command line using CommandLineToArgvW
;   2. Check for CLI switches (-cli, --cli, cli)
;   3. Check for -new flag (new console window)
;   4. Extract command portion from arguments
;   5. Check if command is a .lnk file
;   6. If .lnk: resolve to target executable and arguments
;   7. Execute via RunAsTrustedInstaller or launch GUI
;   8. Clean up and exit
;
; Local variables:
;   pArgv   - Pointer to argument vector array
;   argc    - Argument count
;   msg     - Windows message structure (for GUI mode)
;   argv1   - Pointer to first argument (for switch checking)
;
; Returns: Does not return (calls ExitProcess)
; ==============================================================================
start proc
    LOCAL pArgv:DWORD                       ; Pointer to argv array
    LOCAL argc:DWORD                        ; Argument count
    LOCAL msg:MSG                           ; Message structure for message loop
    LOCAL argv1:DWORD                       ; First argument pointer
    
    cld                                     ; Clear direction flag (forward string ops)
    
    ; Get command line and parse into arguments
    invoke GetCommandLineW                  ; Returns pointer to command line
    mov ecx, eax
    invoke CommandLineToArgvW, ecx, addr argc ; Parse into argc/argv
    mov pArgv, eax
    
    ; Check argument count: need at least 2 for CLI mode (exe + switch)
    cmp argc, 2
    jl mode_gui_free                        ; Less than 2 args → GUI mode
    
    ; Retrieve argv[1] (first argument after executable name)
    mov esi, pArgv
    mov eax, [esi+4]                        ; argv[1] pointer
    mov argv1, eax
    
    ; Check if argv[1] == "-cli" (standard switch)
    push offset str_cliSwitch1
    push argv1
    call wcscmp_ci                          ; Case-insensitive comparison
    test eax, eax
    jnz mode_cli_found                      ; Non-zero = match found

    ; Check if argv[1] == "--cli" (GNU-style long option)
    push offset str_cliSwitch2
    push argv1
    call wcscmp_ci
    test eax, eax
    jnz mode_cli_found                      ; Match found

    ; Check if argv[1] == "cli" (bare switch without hyphens)
    push offset str_cliSwitch3
    push argv1
    call wcscmp_ci
    test eax, eax
    jnz mode_cli_found                      ; Match found

    jmp mode_gui_free                       ; No CLI switch found → GUI mode

mode_cli_found:
    ; CLI mode detected, need at least 3 args (exe, switch, command)
    cmp argc, 3
    jl cli_no_cmd_free                      ; Not enough args for command

    ; Check if argv[2] is "-new" (new console window flag)
    mov esi, pArgv
    mov eax, [esi+8]                        ; argv[2] pointer
    invoke wcscmp_ci, eax, offset str_newSwitch
    test eax, eax
    jz cli_no_new_flag                      ; Zero = no match, not -new

    ; -new flag found: need argc >= 4 (exe, switch, -new, command)
    cmp argc, 4
    jl cli_no_cmd_free                      ; Not enough args
    mov g_useNewConsole, 1                  ; Set new console flag
    jmp cli_free_and_setup

cli_no_new_flag:
    mov g_useNewConsole, 0                  ; No new console, inherit current

cli_free_and_setup:
    invoke LocalFree, pArgv                 ; Free argv array
    jmp mode_cli_setup                      ; Continue to CLI processing
    
mode_gui_free:
    invoke LocalFree, pArgv                 ; Free argv and go to GUI mode
    jmp mode_gui

cli_no_cmd_free:
    invoke LocalFree, pArgv                 ; Insufficient args for CLI
    invoke ExitProcess, 1                   ; Exit with error code

mode_cli_setup:
    ; Re-parse command line manually to extract the command portion
    ; This is necessary because we need the exact command string as entered,
    ; not split by CommandLineToArgvW
    
    invoke GetCommandLineW
    mov esi, eax                            ; ESI = raw command line pointer
    xor ecx, ecx                            ; ECX = unused
    mov edi, 0                              ; EDI = quote flag (inside quotes?)
    
    ; Skip past executable name (may be quoted)
skip_exe_loop:
    mov ax, word ptr [esi]                  ; Read current character
    test ax, ax                             ; Check for end of string
    jz cli_failed_setup
    cmp ax, '"'                             ; Toggle quote flag on quote char
    jne @F
    xor edi, 1                              ; Flip quote state
@@:
    cmp ax, ' '                             ; Check for space
    jne @F
    test edi, edi                           ; Are we inside quotes?
    jnz @F                                  ; Yes, keep searching
    add esi, 2                              ; Space found outside quotes
    jmp skip_switch_init                    ; Done skipping executable
@@:
    add esi, 2                              ; Move to next character
    jmp skip_exe_loop

skip_switch_init:
    invoke skip_spaces, esi                 ; Skip whitespace after executable
    mov esi, eax
    
    ; Skip the CLI switch (we already validated it exists)
    mov edi, 0                              ; Reset quote flag
skip_switch_loop:
    mov ax, word ptr [esi]                  ; Read character
    test ax, ax
    jz cli_failed_setup                     ; Unexpected end of string
    cmp ax, ' '                             ; Find space after switch
    jne @F
    add esi, 2
    jmp after_switch                        ; Switch done
@@:
    add esi, 2                              ; Continue through switch
    jmp skip_switch_loop

after_switch:
    invoke skip_spaces, esi                 ; Skip whitespace after switch
    mov esi, eax

    ; If -new flag is present, skip it
    cmp g_useNewConsole, 0
    je run_command                          ; No -new flag, ESI points to command

skip_new_token:
    mov ax, word ptr [esi]                  ; Skip past -new token
    test ax, ax
    jz cli_failed_setup
    cmp ax, ' '                             ; Find space after -new
    jne @F
    add esi, 2
    jmp run_command                         ; Done skipping -new
@@:
    add esi, 2
    jmp skip_new_token

run_command:
    invoke skip_spaces, esi                 ; Skip any remaining whitespace
    mov esi, eax                            ; ESI now points to actual command

    ; Check if command is a .lnk file (needs at least 4 chars for ".lnk")
    invoke wcslen_p, esi
    mov ecx, eax                            ; ECX = command string length
    cmp ecx, 4
    jl run_no_lnk                           ; Too short to be .lnk
    
    ; Find the first space or end of string to get just the executable path
    ; This separates the .lnk path from any additional arguments
    mov edi, esi                            ; EDI = search pointer
    xor ebx, ebx                            ; EBX = space position (0 = not found)
    xor edx, edx                            ; EDX = quote flag
find_space_or_quote:
    mov ax, word ptr [edi]                  ; Read character
    test ax, ax                             ; End of string?
    jz check_lnk_ext
    cmp ax, '"'                             ; Quote character?
    jne @F
    xor edx, 1                              ; Toggle quote state
@@:
    cmp ax, ' '                             ; Space character?
    jne @F
    test edx, edx                           ; Inside quotes?
    jnz @F                                  ; Yes, ignore this space
    mov ebx, edi                            ; Save space position
@@:
    add edi, 2
    jmp find_space_or_quote

check_lnk_ext:
    ; Check if we found a space (command has additional arguments)
    test ebx, ebx
    jz check_whole_path                     ; No space, check entire string

    ; Space found: extract path portion and check for .lnk extension
    ; Clean up quote characters around the path
    mov ecx, ebx                            ; ECX = end position
    cmp word ptr [ecx-2], '"'               ; Trailing quote before space?
    jne @F
    sub ecx, 2                              ; Exclude trailing quote
@@:
    mov edx, esi                            ; EDX = start position
    cmp word ptr [edx], '"'                 ; Leading quote?
    jne @F
    add edx, 2                              ; Exclude leading quote
@@:
    sub ecx, edx                            ; ECX = path length in bytes
    shr ecx, 1                              ; Convert to character count
    cmp ecx, 4                              ; Long enough for .lnk?
    jl run_no_lnk

    ; Check last 4 characters for ".lnk" extension
    lea edi, [edx + ecx*2 - 8]              ; Point to last 4 chars
    invoke wcscmp_ci, edi, offset str_extLnk_m
    test eax, eax
    jz run_no_lnk                           ; Not .lnk, run as-is

    ; .lnk file with arguments: split into path and args
    push esi                                ; Save original command pointer
    push ebx                                ; Save space position

    ; Recompute clean start/length for path copy (EDX was clobbered)
    mov edx, esi
    cmp word ptr [edx], '"'
    jne @F
    add edx, 2                              ; Skip leading quote
@@:
    mov ecx, ebx                            ; End position
    cmp word ptr [ecx-2], '"'
    jne @F
    sub ecx, 2                              ; Skip trailing quote
@@:
    sub ecx, edx                            ; Length in bytes
    shr ecx, 1                              ; Length in characters
    
    ; Copy path to g_filePath
    mov esi, edx
    mov edi, offset g_filePath
    rep movsw                               ; Copy ECX wide characters
    mov word ptr [edi], 0                   ; Null terminate
    
    ; Extract arguments portion (after the space)
    pop ebx                                 ; Restore space position
    add ebx, 2                              ; Skip past space
    invoke skip_spaces, ebx
    mov esi, eax
    invoke wcscpy_p, offset g_argsBuf, esi  ; Copy args to buffer
    
    ; Zero out temporary buffer for shortcut arguments
    push edi
    mov edi, offset g_tempBuf
    xor eax, eax
    mov ecx, 260
    rep stosd                               ; Clear 260 DWORDs (1040 bytes)
    pop edi
    
    ; Resolve .lnk to get target path and embedded arguments
    invoke ResolveLnkPath, offset g_filePath, offset g_cmdBuf, offset g_tempBuf
    test eax, eax
    pop esi                                 ; Restore original command pointer
    jz run_no_lnk                           ; Resolution failed, use original
    
    ; Check if target path is empty
    invoke wcslen_p, offset g_cmdBuf
    test eax, eax
    jz use_args_only                        ; No target, just use args
    
    ; Append space after target path
    mov edi, offset g_cmdBuf
    lea edi, [edi + eax*2]                  ; Point to end of target path
    mov word ptr [edi], ' '                 ; Add space
    add edi, 2
    mov word ptr [edi], 0                   ; Null terminate
    
use_args_only:
    ; Append embedded shortcut arguments
    invoke wcscat_p, offset g_cmdBuf, offset g_tempBuf
    
    ; Check if we have additional command-line arguments
    invoke wcslen_p, offset g_argsBuf
    test eax, eax
    jz run_resolved                         ; No additional args
    
    ; Append space and additional arguments
    invoke wcscat_p, offset g_cmdBuf, offset str_space
    invoke wcscat_p, offset g_cmdBuf, offset g_argsBuf
    jmp run_resolved

check_whole_path:
    ; No space found - check entire command string for .lnk extension
    ; ECX still contains length from wcslen_p earlier
    mov edx, esi                            ; EDX = start of command
    mov edi, ecx                            ; EDI = length
    
    ; Strip surrounding quotes
    cmp word ptr [edx], '"'
    jne @F
    add edx, 2                              ; Skip leading quote
    dec edi                                 ; Reduce length
@@:
    cmp edi, 1
    jl run_no_lnk                           ; Too short after quote removal
    cmp word ptr [edx + edi*2 - 2], '"'
    jne @F
    dec edi                                 ; Skip trailing quote
@@:
    cmp edi, 4                              ; Long enough for .lnk?
    jl run_no_lnk

    ; Check last 4 characters for .lnk extension
    lea eax, [edx + edi*2 - 8]
    invoke wcscmp_ci, eax, offset str_extLnk_m
    test eax, eax
    jz run_no_lnk                           ; Not .lnk

    ; Zero temporary buffer for shortcut arguments
    push edi                                ; Preserve length
    mov edi, offset g_tempBuf
    xor eax, eax
    mov ecx, 260
    rep stosd                               ; Clear buffer
    pop edi                                 ; Restore length

    ; Copy path to g_filePath (recompute clean start after EDX clobber)
    mov edx, esi
    cmp word ptr [edx], '"'
    jne @F
    add edx, 2                              ; Skip leading quote
@@:
    push esi                                ; Save original pointer
    mov esi, edx
    mov ecx, edi                            ; ECX = length to copy
    mov edi, offset g_filePath
    rep movsw                               ; Copy path
    mov word ptr [edi], 0                   ; Null terminate
    pop esi                                 ; Restore original pointer

    ; Resolve .lnk shortcut
    invoke ResolveLnkPath, offset g_filePath, offset g_cmdBuf, offset g_tempBuf
    test eax, eax
    jz run_no_lnk                           ; Resolution failed
    
    ; Check if target path is empty
    invoke wcslen_p, offset g_cmdBuf
    test eax, eax
    jz use_lnk_args_only                    ; No target, just use args
    
    ; Append space after target path
    mov edi, offset g_cmdBuf
    lea edi, [edi + eax*2]
    mov word ptr [edi], ' '
    add edi, 2
    mov word ptr [edi], 0
    
use_lnk_args_only:
    ; Append embedded shortcut arguments
    invoke wcscat_p, offset g_cmdBuf, offset g_tempBuf
    jmp run_resolved
    
run_resolved:
    ; Execute resolved .lnk target with all arguments
    invoke RunAsTrustedInstaller, offset g_cmdBuf, g_useNewConsole
    jmp run_check_result
    
run_no_lnk:
    ; Not a .lnk file - execute command directly as entered
    invoke RunAsTrustedInstaller, esi, g_useNewConsole
    
run_check_result:
    test eax, eax
    jz cli_failed                           ; Execution failed

    invoke ExitProcess, 0                   ; Success exit

cli_failed_setup:
cli_failed:
    invoke ExitProcess, 1                   ; Error exit
    
mode_gui:
    ; GUI mode: create window and enter message loop
    invoke GetModuleHandleW, 0              ; Get application instance
    mov g_hInstance, eax
    invoke CreateMainWindow, eax            ; Create main window
    test eax, eax
    jz exit_app                             ; Window creation failed
    
msg_loop:
    ; Standard Windows message loop
    invoke GetMessageW, addr msg, 0, 0, 0
    test eax, eax
    jz exit_app                             ; WM_QUIT received
    
    ; Check for ESC key to exit application
    cmp [msg.message], WM_KEYDOWN
    jne msg_not_esc
    cmp [msg.wParam], VK_ESCAPE
    je exit_app                             ; ESC pressed, exit
    
msg_not_esc:
    ; Process dialog messages (for tab key navigation, etc.)
    invoke IsDialogMessageW, g_hwndMain, addr msg
    test eax, eax
    jnz msg_loop                            ; Message was processed
    
    ; Standard message translation and dispatch
    invoke TranslateMessage, addr msg
    invoke DispatchMessageW, addr msg
    jmp msg_loop
    
exit_app:
    invoke ExitProcess, 0                   ; Normal exit
    ret
start endp

end start
