; ==============================================================================
; CMDT - Run as TrustedInstaller
; Main Entry Point and Command-Line Processing Module
; 
; Author: Marek Wesołowski (wesmar)
; Purpose: Provides the entry point and command-line argument parsing for the
;          TrustedInstaller privilege elevation utility.
; ==============================================================================

option casemap:none

include consts.inc

; External function declarations
EXTRN CreateMainWindow:PROC
EXTRN RunAsTrustedInstaller:PROC
EXTRN ResolveLnkPath:PROC
EXTRN GetModuleHandleW:PROC
EXTRN GetMessageW:PROC
EXTRN TranslateMessage:PROC
EXTRN DispatchMessageW:PROC
EXTRN IsDialogMessageW:PROC
EXTRN ExitProcess:PROC
EXTRN GetCommandLineW:PROC
EXTRN CommandLineToArgvW:PROC
EXTRN LocalFree:PROC
EXTRN SetFocus:PROC
EXTRN IsUserAnAdmin:PROC
EXTRN ShellExecuteExW:PROC
EXTRN GetModuleFileNameW:PROC
EXTRN RegCreateKeyExW:PROC
EXTRN RegSetValueExW:PROC
EXTRN RegDeleteKeyW:PROC
EXTRN RegDeleteValueW:PROC
EXTRN RegOpenKeyExW:PROC
EXTRN RegCloseKey:PROC
EXTRN AttachConsole:PROC
EXTRN GetStdHandle:PROC
EXTRN WriteConsoleW:PROC
EXTRN WriteConsoleInputW:PROC
EXTRN WaitForSingleObject:PROC
EXTRN CloseHandle:PROC
EXTRN CoInitializeEx:PROC
EXTRN CoInitializeSecurity:PROC
EXTRN CoCreateInstance:PROC
EXTRN CoSetProxyBlanket:PROC
EXTRN SysAllocString:PROC
EXTRN SysFreeString:PROC
EXTRN VariantInit:PROC
EXTRN VariantClear:PROC
EXTRN SafeArrayCreate:PROC
EXTRN SafeArrayPutElement:PROC
EXTRN CoUninitialize:PROC

; ==============================================================================
; CONSTANT STRING DATA
; ==============================================================================
.const
; Registry key for storing application settings and MRU list
str_regKey      dw 'S','o','f','t','w','a','r','e','\','c','m','d','t',0

; Command-line switch for CLI mode
str_cliSwitch1  dw '-','c','l','i',0           ; CLI mode switch

; Command-line switch for help/usage
str_helpSwitch  dw '-','h','e','l','p',0        ; Display usage and exit

; Switch to request new console window
str_newSwitch   dw '-','n','e','w',0

; File extension for Windows shortcuts
str_extLnk_m    dw '.','l','n','k',0

; Space character for string concatenation
str_space       dw ' ',0

; UAC self-elevation verb
str_runas           dw 'r','u','n','a','s',0

; Context menu registration switches
str_installSwitch   dw '-','i','n','s','t','a','l','l',0
str_uninstallSwitch dw '-','u','n','i','n','s','t','a','l','l',0

; Sticky Keys (sethc.exe) IFEO switches
str_shiftSwitch     dw '-','s','h','i','f','t',0
str_unshiftSwitch   dw '-','u','n','s','h','i','f','t',0

; Registry paths for Explorer context menu - Directory entries
str_ctxKeyBg        dw 'D','i','r','e','c','t','o','r','y','\','B','a','c','k','g','r','o','u','n','d','\','s','h','e','l','l','\','C','M','D','T',0
str_ctxKeyCmdBg     dw 'D','i','r','e','c','t','o','r','y','\','B','a','c','k','g','r','o','u','n','d','\','s','h','e','l','l','\','C','M','D','T','\','c','o','m','m','a','n','d',0
str_ctxKeyDir       dw 'D','i','r','e','c','t','o','r','y','\','s','h','e','l','l','\','C','M','D','T',0
str_ctxKeyCmdDir    dw 'D','i','r','e','c','t','o','r','y','\','s','h','e','l','l','\','C','M','D','T','\','c','o','m','m','a','n','d',0

; Registry paths for Explorer context menu - Executable file entries
str_ctxKeyExe       dw 'e','x','e','f','i','l','e','\','s','h','e','l','l','\','C','M','D','T',0
str_ctxKeyCmdExe    dw 'e','x','e','f','i','l','e','\','s','h','e','l','l','\','C','M','D','T','\','c','o','m','m','a','n','d',0

; Registry paths for Explorer context menu - Shortcut file entries
str_ctxKeyLnk       dw 'l','n','k','f','i','l','e','\','s','h','e','l','l','\','C','M','D','T',0
str_ctxKeyCmdLnk    dw 'l','n','k','f','i','l','e','\','s','h','e','l','l','\','C','M','D','T','\','c','o','m','m','a','n','d',0

; Context menu display text - Directory context menus
str_ctxTextDir      dw 'O','p','e','n',' ','C','M','D',' ','a','s',' ','T','r','u','s','t','e','d','I','n','s','t','a','l','l','e','r',0

; Context menu display text - File context menus (executables and shortcuts)
str_ctxTextFile     dw 'R','u','n',' ','a','s',' ','T','r','u','s','t','e','d','I','n','s','t','a','l','l','e','r',0

; Registry value names and icon paths
str_iconVal         dw 'I','c','o','n',0
str_iconPath        dw 's','h','e','l','l','3','2','.','d','l','l',',','1','0','4',0

; Command template components
str_cmdQuote        dw '"',0
str_cmdSuffixDir    dw '"',' ','-','c','l','i',' ','-','n','e','w',' ','c','m','d','.','e','x','e',' ','/','k',' ','c','d',' ','/','d',' ','"','%','V','"',0
str_cmdSuffixFile   dw '"',' ','"','%','1','"',0

; ==============================================================================
; OBFUSCATED STRINGS - XOR encrypted with key 0x0aah
; These strings are decrypted at runtime to avoid static string detection
; ==============================================================================

; IFEO registry path and values for sethc.exe hook (encrypted)
str_ifeoKey_enc     db 0f9h,0aah,0e5h,0aah,0ech,0aah,0feh,0aah,0fdh,0aah,0ebh,0aah,0f8h,0aah,0efh,0aah
                    db 0f6h,0aah,0e7h,0aah,0c3h,0aah,0c9h,0aah,0d8h,0aah,0c5h,0aah,0d9h,0aah,0c5h,0aah
                    db 0cch,0aah,0deh,0aah,0f6h,0aah,0fdh,0aah,0c3h,0aah,0c4h,0aah,0ceh,0aah,0c5h,0aah
                    db 0ddh,0aah,0d9h,0aah,08ah,0aah,0e4h,0aah,0feh,0aah,0f6h,0aah,0e9h,0aah,0dfh,0aah
                    db 0d8h,0aah,0d8h,0aah,0cfh,0aah,0c4h,0aah,0deh,0aah,0fch,0aah,0cfh,0aah,0d8h,0aah
                    db 0d9h,0aah,0c3h,0aah,0c5h,0aah,0c4h,0aah,0f6h,0aah,0e3h,0aah,0c7h,0aah,0cbh,0aah
                    db 0cdh,0aah,0cfh,0aah,08ah,0aah,0ech,0aah,0c3h,0aah,0c6h,0aah,0cfh,0aah,08ah,0aah
                    db 0efh,0aah,0d2h,0aah,0cfh,0aah,0c9h,0aah,0dfh,0aah,0deh,0aah,0c3h,0aah,0c5h,0aah
                    db 0c4h,0aah,08ah,0aah,0e5h,0aah,0dah,0aah,0deh,0aah,0c3h,0aah,0c5h,0aah,0c4h,0aah
                    db 0d9h,0aah,0f6h,0aah,0d9h,0aah,0cfh,0aah,0deh,0aah,0c2h,0aah,0c9h,0aah,084h,0aah
                    db 0cfh,0aah,0d2h,0aah,0cfh,0aah,0aah,0aah

str_debuggerVal_enc db 0eeh,0aah,0cfh,0aah,0c8h,0aah,0dfh,0aah,0cdh,0aah,0cdh,0aah,0cfh,0aah,0d8h,0aah
                    db 0aah,0aah

str_shiftSuffix     dw ' ','-','c','l','i',' ','-','n','e','w',' ','c','m','d','.','e','x','e',0

; ==============================================================================
; WMI COM GUIDs
; ==============================================================================
CLSID_WbemLocator db 011h,0F8h,090h,045h,03Ah,01Dh,0D0h,011h,089h,01Fh,000h,0AAh,000h,04Bh,02Eh,024h
IID_IWbemLocator  db 087h,0A6h,012h,0DCh,07Fh,073h,0CFh,011h,088h,04Dh,000h,0AAh,000h,04Bh,02Eh,024h

; ==============================================================================
; WMI AND DEFENDER STRINGS - XOR encrypted with key 0x0aah
; ==============================================================================
str_wmi_namespace db 0f8h,0aah,0e5h,0aah,0e5h,0aah,0feh,0aah,0f6h,0aah,0e7h,0aah,0c3h,0aah,0c9h,0aah
                  db 0d8h,0aah,0c5h,0aah,0d9h,0aah,0c5h,0aah,0cch,0aah,0deh,0aah,0f6h,0aah,0fdh,0aah
                  db 0c3h,0aah,0c4h,0aah,0ceh,0aah,0c5h,0aah,0ddh,0aah,0d9h,0aah,0f6h,0aah,0eeh,0aah
                  db 0cfh,0aah,0cch,0aah,0cfh,0aah,0c4h,0aah,0ceh,0aah,0cfh,0aah,0d8h,0aah,0aah,0aah

str_wmi_class     db 0e7h,0aah,0f9h,0aah,0ech,0aah,0feh,0aah,0f5h,0aah,0e7h,0aah,0dah,0aah,0fah,0aah
                  db 0d8h,0aah,0cfh,0aah,0cch,0aah,0cfh,0aah,0d8h,0aah,0cfh,0aah,0c4h,0aah,0c9h,0aah
                  db 0cfh,0aah,0aah,0aah

str_wmi_add       db 0ebh,0aah,0ceh,0aah,0ceh,0aah,0aah,0aah

str_wmi_rem       db 0f8h,0aah,0cfh,0aah,0c7h,0aah,0c5h,0aah,0dch,0aah,0cfh,0aah,0aah,0aah

str_wmi_prop      db 0efh,0aah,0d2h,0aah,0c9h,0aah,0c6h,0aah,0dfh,0aah,0d9h,0aah,0c3h,0aah,0c5h,0aah
                  db 0c4h,0aah,0fah,0aah,0d8h,0aah,0c5h,0aah,0c9h,0aah,0cfh,0aah,0d9h,0aah,0d9h,0aah
                  db 0aah,0aah

str_cmd_exe       db 0c9h,0aah,0c7h,0aah,0ceh,0aah,084h,0aah,0cfh,0aah,0d2h,0aah,0cfh,0aah,0aah,0aah

; Usage help text displayed when an unknown switch is given
str_usage       dw 13,10
                dw 'U','s','a','g','e',':',' ','c','m','d','t','.','e','x','e',' ','[','o','p','t','i','o','n',']',13,10
                dw 13,10
                dw ' ',' ','-','c','l','i',' ','<','c','m','d','>'
                dw ' ',' ',' ',' ',' ',' ',' ',' ',' ',' '
                dw 'R','u','n',' ','c','o','m','m','a','n','d',13,10
                dw ' ',' ','-','c','l','i',' ','-','n','e','w',' ','<','c','m','d','>'
                dw ' ',' ',' ',' ',' '
                dw 'R','u','n',' ','i','n',' ','n','e','w',' ','c','o','n','s','o','l','e',13,10
                dw ' ',' ','-','i','n','s','t','a','l','l'
                dw ' ',' ',' ',' ',' ',' ',' ',' ',' ',' ',' ',' '
                dw 'A','d','d',' ','c','o','n','t','e','x','t',' ','m','e','n','u',13,10
                dw ' ',' ','-','u','n','i','n','s','t','a','l','l'
                dw ' ',' ',' ',' ',' ',' ',' ',' ',' ',' '
                dw 'R','e','m','o','v','e',' ','c','o','n','t','e','x','t',' ','m','e','n','u',13,10
                dw ' ',' ','-','s','h','i','f','t'
                dw ' ',' ',' ',' ',' ',' ',' ',' ',' ',' ',' ',' ',' ',' '
                dw 'H','o','o','k',' ','s','e','t','h','c','.','e','x','e',13,10
                dw ' ',' ','-','u','n','s','h','i','f','t'
                dw ' ',' ',' ',' ',' ',' ',' ',' ',' ',' ',' ',' '
                dw 'U','n','h','o','o','k',' ','s','e','t','h','c','.','e','x','e',13,10
                dw ' ',' ','-','h','e','l','p'
                dw ' ',' ',' ',' ',' ',' ',' ',' ',' ',' ',' ',' ',' ',' '
                dw 'S','h','o','w',' ','t','h','i','s',' ','h','e','l','p',13,10
                dw 13,10
                dw ' ',' ','N','o',' ','a','r','g','s',' ','t','o',' ','s','t','a','r','t',' ','G','U','I','.',13,10
                dw 0

; ==============================================================================
; PRIVILEGE NAME STRINGS
; These strings are used to construct full privilege names by combining with
; the "Se" prefix and "Privilege" suffix (e.g., "SeDebugPrivilege")
; ==============================================================================
privStr_0  dw 'A','s','s','i','g','n','P','r','i','m','a','r','y','T','o','k','e','n',0
privStr_1  dw 'B','a','c','k','u','p',0
privStr_2  dw 'R','e','s','t','o','r','e',0
privStr_3  dw 'D','e','b','u','g',0
privStr_4  dw 'I','m','p','e','r','s','o','n','a','t','e',0
privStr_5  dw 'T','a','k','e','O','w','n','e','r','s','h','i','p',0
privStr_6  dw 'L','o','a','d','D','r','i','v','e','r',0
privStr_7  dw 'S','y','s','t','e','m','E','n','v','i','r','o','n','m','e','n','t',0
privStr_8  dw 'M','a','n','a','g','e','V','o','l','u','m','e',0
privStr_9  dw 'S','e','c','u','r','i','t','y',0
privStr_10 dw 'S','h','u','t','d','o','w','n',0
privStr_11 dw 'S','y','s','t','e','m','t','i','m','e',0
privStr_12 dw 'T','c','b',0
privStr_13 dw 'I','n','c','r','e','a','s','e','Q','u','o','t','a',0
privStr_14 dw 'A','u','d','i','t',0
privStr_15 dw 'C','h','a','n','g','e','N','o','t','i','f','y',0
privStr_16 dw 'U','n','d','o','c','k',0
privStr_17 dw 'C','r','e','a','t','e','T','o','k','e','n',0
privStr_18 dw 'L','o','c','k','M','e','m','o','r','y',0
privStr_19 dw 'C','r','e','a','t','e','P','a','g','e','f','i','l','e',0
privStr_20 dw 'C','r','e','a','t','e','P','e','r','m','a','n','e','n','t',0
privStr_21 dw 'S','y','s','t','e','m','P','r','o','f','i','l','e',0
privStr_22 dw 'P','r','o','f','i','l','e','S','i','n','g','l','e','P','r','o','c','e','s','s',0
privStr_23 dw 'C','r','e','a','t','e','G','l','o','b','a','l',0
privStr_24 dw 'T','i','m','e','Z','o','n','e',0
privStr_25 dw 'C','r','e','a','t','e','S','y','m','b','o','l','i','c','L','i','n','k',0
privStr_26 dw 'I','n','c','r','e','a','s','e','B','a','s','e','P','r','i','o','r','i','t','y',0
privStr_27 dw 'R','e','m','o','t','e','S','h','u','t','d','o','w','n',0
privStr_28 dw 'I','n','c','r','e','a','s','e','W','o','r','k','i','n','g','S','e','t',0
privStr_29 dw 'R','e','l','a','b','e','l',0
privStr_30 dw 'D','e','l','e','g','a','t','e','S','e','s','s','i','o','n','U','s','e','r','I','m','p','e','r','s','o','n','a','t','e',0
privStr_31 dw 'T','r','u','s','t','e','d','C','r','e','d','M','a','n','A','c','c','e','s','s',0
privStr_32 dw 'E','n','a','b','l','e','D','e','l','e','g','a','t','i','o','n',0
privStr_33 dw 'S','y','n','c','A','g','e','n','t',0

; Prefix and suffix for privilege name construction
privPrefix dw 'S','e',0
privSuffix dw 'P','r','i','v','i','l','e','g','e',0

; ==============================================================================
; INITIALIZED DATA SECTION
; ==============================================================================
.data
    align 8

; Privilege table: Array of pointers to privilege name strings
; Used by token manipulation code to enable all available privileges
PUBLIC g_privTable
g_privTable dq offset privStr_0,offset privStr_1,offset privStr_2,offset privStr_3,offset privStr_4,offset privStr_5
            dq offset privStr_6,offset privStr_7,offset privStr_8,offset privStr_9,offset privStr_10,offset privStr_11
            dq offset privStr_12,offset privStr_13,offset privStr_14,offset privStr_15,offset privStr_16,offset privStr_17
            dq offset privStr_18,offset privStr_19,offset privStr_20,offset privStr_21,offset privStr_22,offset privStr_23
            dq offset privStr_24,offset privStr_25,offset privStr_26,offset privStr_27,offset privStr_28,offset privStr_29
            dq offset privStr_30,offset privStr_31,offset privStr_32,offset privStr_33

; Global variables exported to other modules
PUBLIC g_cachedToken, g_tokenTime, g_hwndMain, g_hwndEdit, g_hwndBtn, g_hwndStatus, g_hConsoleOut, g_hInstance
PUBLIC privPrefix, privSuffix

; Cached TrustedInstaller token handle (for performance optimization)
g_cachedToken   dq 0

; Timestamp of when the cached token was obtained (in milliseconds)
g_tokenTime     dd 0
                dd 0            ; Padding for alignment

; Window handles for GUI controls
g_hwndMain      dq 0            ; Main window handle
g_hwndEdit      dq 0            ; ComboBox edit control handle
g_hwndBtn       dq 0            ; Run button handle
g_hwndStatus    dq 0            ; Status label handle

; Console output handle (for CLI mode)
g_hConsoleOut   dq 0

; Flag indicating whether to create a new console window
g_useNewConsole dd 0
                dd 0            ; Padding for alignment

; Application instance handle
g_hInstance     dq 0

; ==============================================================================
; UNINITIALIZED DATA SECTION
; ==============================================================================
.data?

; Export buffer declarations for use by other modules
PUBLIC g_cmdBuf, g_statusBuf, g_filePath, g_argsBuf, g_tempBuf, g_exePath, g_decryptBuf

; Buffer for command line text (520 WCHARs = 1040 bytes)
g_cmdBuf        dw 520 dup(?)

; Buffer for status text (520 WCHARs)
g_statusBuf     dw 520 dup(?)

; Buffer for file paths (520 WCHARs)
g_filePath      dw 520 dup(?)

; Buffer for command-line arguments (520 WCHARs)
g_argsBuf       dw 520 dup(?)

; Temporary buffer for various operations (1040 WCHARs = 2080 bytes)
g_tempBuf       dw 1040 dup(?)

; Buffer for exe path (UAC elevation and context menu registration)
g_exePath       dw 260 dup(?)

; Buffer for decrypted strings (reusable, 520 WCHARs)
g_decryptBuf    dw 520 dup(?)

; ==============================================================================
; CODE SECTION
; ==============================================================================
.code

; ==============================================================================
; DecryptWideStr - XOR Decrypt Wide String In-Place
;
; Purpose: Decrypts a XOR-encrypted wide character string into destination buffer
;          Uses simple XOR with single-byte key applied to each byte
;
; Parameters:
;   RCX = Pointer to encrypted source string
;   RDX = Pointer to destination buffer
;
; Returns:
;   RAX = Pointer to destination buffer (same as RDX input)
;
; Modifies: RAX, RSI, RDI
;
; Notes:
;   - XOR key is hardcoded as 0x0aah
;   - Decryption stops at null terminator (0x0000)
;   - Each byte of the wide string is XORed independently
; ==============================================================================
DecryptWideStr proc
    push rsi
    push rdi
    
    mov rsi, rcx                ; RSI = source (encrypted)
    mov rdi, rdx                ; RDI = destination
    
dws_loop:
    ; Decrypt first byte of wide char
    mov al, byte ptr [rsi]
    xor al, 0aah                ; XOR with key
    mov byte ptr [rdi], al
    
    ; Decrypt second byte of wide char
    mov al, byte ptr [rsi+1]
    xor al, 0aah                ; XOR with key
    mov byte ptr [rdi+1], al
    
    ; Check if we hit null terminator
    cmp word ptr [rdi], 0
    je dws_done
    
    ; Move to next character
    add rsi, 2
    add rdi, 2
    jmp dws_loop
    
dws_done:
    mov rax, rdx                ; Return destination pointer
    pop rdi
    pop rsi
    ret
DecryptWideStr endp

; ==============================================================================
; wcscpy_p - Wide Character String Copy (Private Implementation)
;
; Purpose: Copies a null-terminated wide character string from source to dest
;
; Parameters:
;   RCX = Destination buffer pointer
;   RDX = Source string pointer
;
; Returns: None
;
; Modifies: RAX, RDI, RSI (saved/restored), destination buffer
; ==============================================================================
wcscpy_p proc
    push rsi
    push rdi
    mov rdi, rcx                ; RDI = destination pointer
    mov rsi, rdx                ; RSI = source pointer
@@:
    mov ax, word ptr [rsi]      ; Read wide character from source
    mov word ptr [rdi], ax      ; Write to destination
    test ax, ax                 ; Check for null terminator
    jz @F                       ; Exit if null terminator found
    add rsi, 2                  ; Advance source pointer
    add rdi, 2                  ; Advance destination pointer
    jmp @B                      ; Continue copying
@@:
    pop rdi
    pop rsi
    ret
wcscpy_p endp

; ==============================================================================
; wcscat_p - Wide Character String Concatenate (Private Implementation)
;
; Purpose: Appends source string to the end of destination string
;
; Parameters:
;   RCX = Destination buffer pointer
;   RDX = Source string pointer
;
; Returns: None
;
; Modifies: RAX, RDI, RSI (saved/restored), destination buffer
; ==============================================================================
wcscat_p proc
    push rsi
    push rdi
    mov rdi, rcx                ; RDI = destination pointer
@@:
    cmp word ptr [rdi], 0       ; Check for null terminator
    je @F                       ; Found end of destination string
    add rdi, 2                  ; Move to next character
    jmp @B                      ; Continue searching
@@:
    mov rsi, rdx                ; RSI = source pointer
@@:
    mov ax, word ptr [rsi]      ; Read wide character from source
    mov word ptr [rdi], ax      ; Write to destination
    test ax, ax                 ; Check for null terminator
    jz @F                       ; Exit if null terminator found
    add rsi, 2                  ; Advance source pointer
    add rdi, 2                  ; Advance destination pointer
    jmp @B                      ; Continue copying
@@:
    pop rdi
    pop rsi
    ret
wcscat_p endp

; ==============================================================================
; wcscmp_ci - Wide Character String Compare (Case-Insensitive)
;
; Purpose: Compares two wide character strings ignoring case differences
;
; Parameters:
;   RCX = First string pointer
;   RDX = Second string pointer
;
; Returns:
;   RAX = 1 if strings are equal (case-insensitive), 0 otherwise
;
; Modifies: RAX, RDX, RSI, RDI (saved/restored)
; ==============================================================================
wcscmp_ci proc
    push rsi
    push rdi
    mov rsi, rcx                ; RSI = first string
    mov rdi, rdx                ; RDI = second string
wci_loop:
    mov ax, word ptr [rsi]      ; Load character from first string
    mov dx, word ptr [rdi]      ; Load character from second string
    
    ; Convert first character to lowercase if uppercase
    cmp ax, 'A'
    jb wci_skip1
    cmp ax, 'Z'
    ja wci_skip1
    add ax, 32                  ; Convert A-Z to a-z
wci_skip1:
    
    ; Convert second character to lowercase if uppercase
    cmp dx, 'A'
    jb wci_skip2
    cmp dx, 'Z'
    ja wci_skip2
    add dx, 32                  ; Convert A-Z to a-z
wci_skip2:
    
    cmp ax, dx                  ; Compare normalized characters
    jne not_eq                  ; Characters differ
    test ax, ax                 ; Check if end of strings
    jz equal                    ; Both null terminators reached
    add rsi, 2                  ; Advance first string pointer
    add rdi, 2                  ; Advance second string pointer
    jmp wci_loop                ; Continue comparison
equal:
    pop rdi
    pop rsi
    mov rax, 1                  ; Return 1 (strings equal)
    ret
not_eq:
    pop rdi
    pop rsi
    xor rax, rax                ; Return 0 (strings differ)
    ret
wcscmp_ci endp

; ==============================================================================
; skip_spaces - Skip Leading Whitespace in Wide String
;
; Purpose: Advances a string pointer past any leading space characters
;
; Parameters:
;   RCX = String pointer
;
; Returns:
;   RAX = Pointer to first non-space character
;
; Modifies: RAX
; ==============================================================================
skip_spaces proc
    mov rax, rcx                ; RAX = input pointer
@@:
    cmp word ptr [rax], ' '     ; Check if current character is space
    jne @F                      ; Exit if non-space found
    add rax, 2                  ; Skip this space
    jmp @B                      ; Continue checking
@@:
    ret
skip_spaces endp

; ==============================================================================
; wcslen_p - Wide Character String Length (Private Implementation)
;
; Purpose: Calculates the length of a null-terminated wide character string
;
; Parameters:
;   RCX = String pointer
;
; Returns:
;   RAX = Number of wide characters (excluding null terminator)
;
; Modifies: RAX
; ==============================================================================
wcslen_p proc
    mov rax, rcx                ; RAX = string pointer
@@:
    cmp word ptr [rax], 0       ; Check for null terminator
    jz @F                       ; Exit if found
    add rax, 2                  ; Advance pointer
    jmp @B                      ; Continue counting
@@:
    sub rax, rcx                ; Calculate byte difference
    sar rax, 1                  ; Divide by 2 to get character count
    ret
wcslen_p endp

; ==============================================================================
; mainCRTStartup - Application Entry Point
;
; Purpose: Main entry point for the application. Parses command-line arguments
;          and either runs in CLI mode or launches the GUI window.
;
; Command-line syntax:
;   - GUI mode: <exe> (no arguments)
;   - CLI mode: <exe> -cli <command>
;   - CLI mode with new console: <exe> -cli -new <command>
;
; Returns: Does not return (calls ExitProcess)
;
; Stack frame: 312 bytes local variables
; ==============================================================================
mainCRTStartup proc frame
    push rbp
    .pushreg rbp
    mov rbp, rsp
    .setframe rbp, 0
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
    sub rsp, 312
    .allocstack 312
    .endprolog

    cld                         ; Clear direction flag for string operations

    ; Check if running as administrator
    sub rsp, 32
    call IsUserAnAdmin
    add rsp, 32
    test eax, eax
    jnz uac_already_admin

    ; Not admin - relaunch with UAC elevation prompt
    lea rdx, g_exePath
    mov r8d, 260
    xor ecx, ecx
    sub rsp, 32
    call GetModuleFileNameW
    add rsp, 32

    ; Extract arguments from command line (skip exe path)
    sub rsp, 32
    call GetCommandLineW
    add rsp, 32
    mov rsi, rax
    xor edi, edi
uac_skip_exe:
    mov ax, word ptr [rsi]
    test ax, ax
    jz uac_no_args
    cmp ax, '"'
    jne @F
    xor edi, 1
@@:
    cmp ax, ' '
    jne @F
    test edi, edi
    jnz @F
    mov rcx, rsi
    call skip_spaces
    mov r15, rax
    jmp uac_launch
@@:
    add rsi, 2
    jmp uac_skip_exe

uac_no_args:
    mov r15, rsi                ; Points to null terminator (empty args)

uac_launch:
    ; Zero SHELLEXECUTEINFOW at [rbp-312] (112 bytes)
    lea rdi, [rbp-312]
    xor rax, rax
    mov rcx, 14
uac_zero:
    mov qword ptr [rdi], rax
    add rdi, 8
    dec rcx
    jnz uac_zero

    ; Fill SHELLEXECUTEINFOW fields
    mov dword ptr [rbp-312], 112            ; cbSize
    lea rax, str_runas
    mov qword ptr [rbp-312+16], rax         ; lpVerb = "runas"
    lea rax, g_exePath
    mov qword ptr [rbp-312+24], rax         ; lpFile = exe path
    mov qword ptr [rbp-312+32], r15         ; lpParameters = arguments
    mov dword ptr [rbp-312+48], SW_SHOWNORMAL ; nShow

    lea rcx, [rbp-312]
    sub rsp, 32
    call ShellExecuteExW
    add rsp, 32

    ; Exit - elevated instance takes over, or user cancelled UAC
    xor ecx, ecx
    sub rsp, 32
    call ExitProcess
    add rsp, 32

uac_already_admin:
    ; Get the full command line string
    sub rsp, 32
    call GetCommandLineW
    add rsp, 32
    mov r12, rax                ; R12 = command line string

    ; Parse command line into argv array
    lea rdx, [rbp-64]           ; RDX = pointer to argc output
    mov rcx, r12                ; RCX = command line string
    sub rsp, 32
    call CommandLineToArgvW
    add rsp, 32
    mov r13, rax                ; R13 = argv array

    ; Check if at least 2 arguments (exe + switch)
    mov eax, dword ptr [rbp-64] ; EAX = argc
    cmp eax, 2
    jl mode_gui_free            ; Less than 2 args: GUI mode

    ; Get second argument (first after exe path)
    mov rax, [r13+8]            ; RAX = argv[1]
    mov r14, rax                ; R14 = argv[1]

    ; Check if argv[1] matches "-cli"
    lea rdx, str_cliSwitch1
    mov rcx, r14
    call wcscmp_ci
    test rax, rax
    jnz mode_cli_found

    ; Check if argv[1] matches "-install"
    lea rdx, str_installSwitch
    mov rcx, r14
    call wcscmp_ci
    test rax, rax
    jnz mode_install_found

    ; Check if argv[1] matches "-uninstall"
    lea rdx, str_uninstallSwitch
    mov rcx, r14
    call wcscmp_ci
    test rax, rax
    jnz mode_uninstall_found

    ; Check if argv[1] matches "-shift"
    lea rdx, str_shiftSwitch
    mov rcx, r14
    call wcscmp_ci
    test rax, rax
    jnz mode_shift_found

    ; Check if argv[1] matches "-unshift"
    lea rdx, str_unshiftSwitch
    mov rcx, r14
    call wcscmp_ci
    test rax, rax
    jnz mode_unshift_found

    ; Check if argv[1] matches "-help"
    lea rdx, str_helpSwitch
    mov rcx, r14
    call wcscmp_ci
    test rax, rax
    jnz show_usage

    ; No recognized switch: check if argv[1] starts with '-'
    cmp word ptr [r14], '-'
    je show_usage               ; Unknown switch, display available options
    jmp mode_file_run           ; Not a switch, treat as file path

mode_file_run:
    ; ===== Direct File Execution Mode =====
    ; This mode is triggered when the user right-clicks on a file (.exe or .lnk)
    ; and selects "Run as TrustedInstaller" from the context menu.
    ; 
    ; The context menu passes: cmdt.exe "%1"
    ; where %1 is the full path to the clicked file.
    ;
    ; Since argv[1] is not a recognized switch (-cli, -install, etc.),
    ; we treat it as a file path and execute it directly with TrustedInstaller
    ; privileges without showing the GUI.
    ;
    ; Process flow:
    ;   1. Free the argv array (no longer needed)
    ;   2. Get the raw command line
    ;   3. Skip past the executable path to find the file argument
    ;   4. Execute the file with RunAsTrustedInstaller
    ;   5. Exit immediately (no GUI)
    
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
    mov ax, word ptr [rsi]      ; Read current character
    test ax, ax                 ; Check for end of string
    jz mode_gui                 ; No file argument found, show GUI
    cmp ax, '"'                 ; Quote character?
    jne @F
    xor edi, 1                  ; Toggle quote state
@@:
    cmp ax, ' '                 ; Space character?
    jne @F
    test edi, edi               ; Inside quotes?
    jnz @F                      ; Yes, continue scanning
    add rsi, 2                  ; No, space ends executable path
    mov rcx, rsi
    call skip_spaces            ; Skip whitespace after exe
    mov rsi, rax
    jmp run_file_direct         ; RSI now points to file path
@@:
    add rsi, 2                  ; Move to next character
    jmp skip_exe_for_file

run_file_direct:
    ; RSI points to the file path argument (may be quoted from context menu)
    ; Check for .lnk shortcut and resolve before execution

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
    mov dword ptr g_useNewConsole, 1 ; Set new console flag
    jmp cli_free_and_setup

cli_no_new_flag:
    mov dword ptr g_useNewConsole, 0 ; Clear new console flag

cli_free_and_setup:
    ; Free argv array allocated by CommandLineToArgvW
    mov rcx, r13
    sub rsp, 32
    call LocalFree
    add rsp, 32
    jmp mode_cli_setup

mode_install_found:
    ; Free argv and register context menu
    mov rcx, r13
    sub rsp, 32
    call LocalFree
    add rsp, 32
    call InstallContextMenu
    xor ecx, ecx
    sub rsp, 32
    call ExitProcess
    add rsp, 32

mode_uninstall_found:
    ; Free argv and remove context menu
    mov rcx, r13
    sub rsp, 32
    call LocalFree
    add rsp, 32
    call UninstallContextMenu
    xor ecx, ecx
    sub rsp, 32
    call ExitProcess
    add rsp, 32

mode_shift_found:
    ; Free argv and install sethc.exe IFEO hook
    mov rcx, r13
    sub rsp, 32
    call LocalFree
    add rsp, 32
    call InstallShift
    xor ecx, ecx
    sub rsp, 32
    call ExitProcess
    add rsp, 32

mode_unshift_found:
    ; Free argv and remove sethc.exe IFEO hook
    mov rcx, r13
    sub rsp, 32
    call LocalFree
    add rsp, 32
    call UninstallShift
    xor ecx, ecx
    sub rsp, 32
    call ExitProcess
    add rsp, 32

show_usage:
    ; Unknown switch detected, display available options and exit
    mov rcx, r13
    sub rsp, 32
    call LocalFree
    add rsp, 32

    ; Attach to parent console (ATTACH_PARENT_PROCESS)
    mov ecx, ATTACH_PARENT_PROCESS
    sub rsp, 32
    call AttachConsole
    add rsp, 32
    test eax, eax
    jz show_usage_exit

    ; Get stdout handle (STD_OUTPUT_HANDLE)
    mov ecx, STD_OUTPUT_HANDLE
    sub rsp, 32
    call GetStdHandle
    add rsp, 32
    mov rbx, rax

    ; Calculate usage text length and write to console
    lea rcx, str_usage
    call wcslen_p
    mov r12, rax

    sub rsp, 48
    mov qword ptr [rsp+32], 0  ; lpReserved
    lea r9, [rbp-64]           ; lpNumberOfCharsWritten
    mov r8, r12                ; nNumberOfCharsToWrite
    lea rdx, str_usage         ; lpBuffer
    mov rcx, rbx               ; hConsoleOutput
    call WriteConsoleW
    add rsp, 48

    ; Inject Enter key into console input so cmd.exe redraws prompt immediately
    ; (AttachConsole causes cmd.exe to show its prompt before our output appears,
    ;  so after we exit the cursor blinks waiting for input - sending Enter fixes it)
    mov ecx, STD_INPUT_HANDLE
    sub rsp, 32
    call GetStdHandle
    add rsp, 32
    test rax, rax
    jz show_usage_exit
    mov rdi, rax                ; RDI = stdin handle

    ; Build INPUT_RECORD at [rbp-96] (20 bytes, safe in local frame)
    mov word ptr [rbp-96], 1    ; EventType = KEY_EVENT
    mov word ptr [rbp-94], 0    ; padding
    mov dword ptr [rbp-92], 1   ; bKeyDown = TRUE
    mov word ptr [rbp-88], 1    ; wRepeatCount = 1
    mov word ptr [rbp-86], 0Dh  ; wVirtualKeyCode = VK_RETURN
    mov word ptr [rbp-84], 1Ch  ; wVirtualScanCode
    mov word ptr [rbp-82], 0Dh  ; uChar.UnicodeChar = '\r'
    mov dword ptr [rbp-80], 0   ; dwControlKeyState = 0

    sub rsp, 32
    lea r9, [rbp-64]            ; lpNumberOfEventsWritten
    mov r8d, 1                  ; nLength = 1
    lea rdx, [rbp-96]           ; lpBuffer = &inputRecord
    mov rcx, rdi                ; hConsoleInput
    call WriteConsoleInputW
    add rsp, 32

show_usage_exit:
    mov ecx, 1
    sub rsp, 32
    call ExitProcess
    add rsp, 32

mode_gui_free:
    ; Free argv array and proceed to GUI mode
    mov rcx, r13
    sub rsp, 32
    call LocalFree
    add rsp, 32
    jmp mode_gui

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

mode_cli_setup:
    ; Parse the raw command line to extract the actual command
    ; Need to skip: exe path, CLI switch, and optionally "-new"
    
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
    jl run_no_lnk               ; Too short to be .lnk

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
    jmp check_lnk_ext            ; Stop at first space
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
    lea rdi, [r10 + rcx*2 - 8]  ; Point to last 4 chars
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
    lea r8, g_tempBuf           ; R8 = buffer for .lnk arguments
    lea rdx, g_cmdBuf           ; RDX = buffer for target path
    lea rcx, g_filePath         ; RCX = .lnk file path
    sub rsp, 32
    call ResolveLnkPath
    add rsp, 32
    test rax, rax
    pop rsi
    jz run_no_lnk               ; Resolution failed

    ; Build final command: target + .lnk args + user args
    lea rcx, g_cmdBuf
    call wcslen_p
    test rax, rax
    jz use_args_only            ; No target path

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
    jz run_resolved             ; No user arguments

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
    ; Compute clean bounds (strip surrounding quotes)
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
    jl run_no_lnk               ; Too short for .lnk

    ; Check if last 4 characters are ".lnk"
    lea rdi, [r10 + r11*2 - 8]
    lea rdx, str_extLnk_m
    mov rcx, rdi
    call wcscmp_ci
    test rax, rax
    jz run_no_lnk               ; Not a .lnk file

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
    mov rcx, r11                ; RCX = length
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
    mov word ptr [rdi], 0       ; Null terminate

    ; Resolve the .lnk file
    lea r8, g_tempBuf
    lea rdx, g_cmdBuf
    lea rcx, g_filePath
    sub rsp, 32
    call ResolveLnkPath
    add rsp, 32
    test rax, rax
    jz run_no_lnk               ; Resolution failed

    ; Build command from resolved target and arguments
    lea rcx, g_cmdBuf
    call wcslen_p
    test rax, rax
    jz use_lnk_args_only        ; No target path

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
    ; Check execution result
    test rax, rax
    jz cli_failed               ; Execution failed

    ; Success: exit with code 0
    xor ecx, ecx
    sub rsp, 32
    call ExitProcess
    add rsp, 32

cli_failed_setup:
cli_failed:
    ; Failure: exit with code 1
    mov ecx, 1
    sub rsp, 32
    call ExitProcess
    add rsp, 32

mode_gui:
    ; GUI Mode: Create and show main window
    
    ; Get application instance handle
    xor ecx, ecx
    sub rsp, 32
    call GetModuleHandleW
    add rsp, 32
    mov g_hInstance, rax

    ; Create the main application window
    mov rcx, rax
    sub rsp, 32
    call CreateMainWindow
    add rsp, 32
    test rax, rax
    jz exit_app                 ; Window creation failed

msg_loop:
    ; Main message loop
    lea rcx, [rbp-200]          ; RCX = &msg structure
    xor edx, edx                ; EDX = hWnd (NULL = all windows)
    xor r8d, r8d                ; R8D = wMsgFilterMin (0 = all messages)
    xor r9d, r9d                ; R9D = wMsgFilterMax (0 = all messages)
    sub rsp, 32
    call GetMessageW
    add rsp, 32
    test eax, eax
    jz exit_app                 ; WM_QUIT received

    ; Check for ESC key to exit application
    mov eax, dword ptr [rbp-200+8]  ; EAX = message
    cmp eax, WM_KEYDOWN
    jne msg_not_esc
    mov rax, qword ptr [rbp-200+16] ; RAX = wParam
    cmp eax, VK_ESCAPE
    je exit_app                 ; ESC pressed: exit

msg_not_esc:
    ; Process dialog messages (for tab navigation, etc.)
    lea rdx, [rbp-200]
    mov rcx, g_hwndMain
    sub rsp, 32
    call IsDialogMessageW
    add rsp, 32
    test eax, eax
    jnz msg_loop                ; Message was processed

    ; Translate and dispatch regular window messages
    lea rcx, [rbp-200]
    sub rsp, 32
    call TranslateMessage
    add rsp, 32

    lea rcx, [rbp-200]
    sub rsp, 32
    call DispatchMessageW
    add rsp, 32
    jmp msg_loop

exit_app:
    ; Exit the application
    xor ecx, ecx
    sub rsp, 32
    call ExitProcess
    add rsp, 32

    ; Cleanup and return (unreachable, but proper epilog)
    add rsp, 312
    pop r15
    pop r14
    pop r13
    pop r12
    pop rdi
    pop rsi
    pop rbx
    pop rbp
    ret
mainCRTStartup endp

; ==============================================================================
; GetExeFileName - Extract filename from g_exePath
;
; Purpose: Scans g_exePath for the last backslash and returns a pointer
;          to the character after it (the filename portion).
;          Must call GetModuleFileNameW into g_exePath first.
;
; Parameters: None
;
; Returns: RAX = pointer to filename within g_exePath
;
; Modifies: RAX, RCX
; ==============================================================================
GetExeFileName proc
    lea rax, g_exePath
    mov rcx, rax
@@:
    cmp word ptr [rcx], 0
    je @F
    cmp word ptr [rcx], '\'
    jne gef_next
    lea rax, [rcx+2]
gef_next:
    add rcx, 2
    jmp @B
@@:
    ret
GetExeFileName endp

; ==============================================================================
; InstallContextMenu - Register Explorer context menu entries
;
; Purpose: Creates registry keys under HKEY_CLASSES_ROOT for context menu
;          entries that allow running executables and opening directories
;          with TrustedInstaller privileges.
;
; Registry locations created:
;   - Directory\Background\shell\CMDT (background right-click in folders)
;   - Directory\shell\CMDT (right-click on folder icons)
;   - exefile\shell\CMDT (right-click on .exe files)
;   - lnkfile\shell\CMDT (right-click on .lnk shortcut files)
;
; Each entry includes:
;   - Default value: Menu text ("Open CMD as TrustedInstaller" or "Run as TrustedInstaller")
;   - Icon value: Path to shell32.dll icon #104 (UAC shield icon)
;   - command subkey: Command line to execute when menu item is selected
;
; Commands generated:
;   - For directories: "<exepath>" -cli -new cmd.exe /k cd /d "%V"
;   - For files: "<exepath>" "%1"
;
; Parameters: None (uses global g_exePath buffer)
;
; Returns: None (ignores errors to allow partial installation)
;
; Stack frame: 104 bytes for registry handles and stack parameters
; ==============================================================================
InstallContextMenu proc
    push rbx
    push r12
    push r13
    push r14
    sub rsp, 104                ; 32 shadow + 40 stack params + 8 hKey + 8 dwDisp + alignment
    ; [rsp+72] = hKey, [rsp+80] = dwDisp

    ; Get our exe path for command strings
    lea rdx, g_exePath
    mov r8d, 260
    xor ecx, ecx
    call GetModuleFileNameW

    ; Build directory command string: "<exepath>" -cli -new cmd.exe /k cd /d "%V"
    lea rcx, g_tempBuf
    lea rdx, str_cmdQuote
    call wcscpy_p
    lea rcx, g_tempBuf
    lea rdx, g_exePath
    call wcscat_p
    lea rcx, g_tempBuf
    lea rdx, str_cmdSuffixDir
    call wcscat_p

    ; Calculate string byte sizes (characters * 2 + null terminator * 2)
    lea rcx, g_tempBuf
    call wcslen_p
    lea rbx, [rax+1]
    shl rbx, 1                  ; RBX = directory command string byte size

    lea rcx, str_ctxTextDir
    call wcslen_p
    lea r12, [rax+1]
    shl r12, 1                  ; R12 = directory menu text byte size

    lea rcx, str_iconPath
    call wcslen_p
    lea r13, [rax+1]
    shl r13, 1                  ; R13 = icon path byte size (shell32.dll,104)

    ; --- Directory\Background\shell\CMDT (parent key) ---
    lea rax, [rsp+80]
    mov qword ptr [rsp+64], rax
    lea rax, [rsp+72]
    mov qword ptr [rsp+56], rax
    mov qword ptr [rsp+48], 0
    mov qword ptr [rsp+40], KEY_WRITE
    mov qword ptr [rsp+32], 0
    xor r9, r9
    xor r8d, r8d
    lea rdx, str_ctxKeyBg
    mov ecx, HKEY_CLASSES_ROOT
    call RegCreateKeyExW
    test eax, eax
    jnz ctx_install_done

    ; Set default value = "Open CMD as TrustedInstaller"
    mov qword ptr [rsp+40], r12
    lea rax, str_ctxTextDir
    mov qword ptr [rsp+32], rax
    mov r9d, REG_SZ
    xor r8d, r8d
    xor edx, edx
    mov rcx, [rsp+72]
    call RegSetValueExW

    ; Set Icon = shell32.dll,104
    mov qword ptr [rsp+40], r13
    lea rax, str_iconPath
    mov qword ptr [rsp+32], rax
    mov r9d, REG_SZ
    xor r8d, r8d
    lea rdx, str_iconVal
    mov rcx, [rsp+72]
    call RegSetValueExW

    mov rcx, [rsp+72]
    call RegCloseKey

    ; --- Directory\Background\shell\CMDT\command (command subkey) ---
    lea rax, [rsp+80]
    mov qword ptr [rsp+64], rax
    lea rax, [rsp+72]
    mov qword ptr [rsp+56], rax
    mov qword ptr [rsp+48], 0
    mov qword ptr [rsp+40], KEY_WRITE
    mov qword ptr [rsp+32], 0
    xor r9, r9
    xor r8d, r8d
    lea rdx, str_ctxKeyCmdBg
    mov ecx, HKEY_CLASSES_ROOT
    call RegCreateKeyExW
    test eax, eax
    jnz ctx_install_done

    mov qword ptr [rsp+40], rbx
    lea rax, g_tempBuf
    mov qword ptr [rsp+32], rax
    mov r9d, REG_SZ
    xor r8d, r8d
    xor edx, edx
    mov rcx, [rsp+72]
    call RegSetValueExW

    mov rcx, [rsp+72]
    call RegCloseKey

    ; --- Directory\shell\CMDT (parent key) ---
    lea rax, [rsp+80]
    mov qword ptr [rsp+64], rax
    lea rax, [rsp+72]
    mov qword ptr [rsp+56], rax
    mov qword ptr [rsp+48], 0
    mov qword ptr [rsp+40], KEY_WRITE
    mov qword ptr [rsp+32], 0
    xor r9, r9
    xor r8d, r8d
    lea rdx, str_ctxKeyDir
    mov ecx, HKEY_CLASSES_ROOT
    call RegCreateKeyExW
    test eax, eax
    jnz ctx_install_done

    mov qword ptr [rsp+40], r12
    lea rax, str_ctxTextDir
    mov qword ptr [rsp+32], rax
    mov r9d, REG_SZ
    xor r8d, r8d
    xor edx, edx
    mov rcx, [rsp+72]
    call RegSetValueExW

    mov qword ptr [rsp+40], r13
    lea rax, str_iconPath
    mov qword ptr [rsp+32], rax
    mov r9d, REG_SZ
    xor r8d, r8d
    lea rdx, str_iconVal
    mov rcx, [rsp+72]
    call RegSetValueExW

    mov rcx, [rsp+72]
    call RegCloseKey

    ; --- Directory\shell\CMDT\command (command subkey) ---
    lea rax, [rsp+80]
    mov qword ptr [rsp+64], rax
    lea rax, [rsp+72]
    mov qword ptr [rsp+56], rax
    mov qword ptr [rsp+48], 0
    mov qword ptr [rsp+40], KEY_WRITE
    mov qword ptr [rsp+32], 0
    xor r9, r9
    xor r8d, r8d
    lea rdx, str_ctxKeyCmdDir
    mov ecx, HKEY_CLASSES_ROOT
    call RegCreateKeyExW
    test eax, eax
    jnz ctx_install_done

    mov qword ptr [rsp+40], rbx
    lea rax, g_tempBuf
    mov qword ptr [rsp+32], rax
    mov r9d, REG_SZ
    xor r8d, r8d
    xor edx, edx
    mov rcx, [rsp+72]
    call RegSetValueExW

    mov rcx, [rsp+72]
    call RegCloseKey

    ; Build file command string: "<exepath>" "%1"
    lea rcx, g_tempBuf
    lea rdx, str_cmdQuote
    call wcscpy_p
    lea rcx, g_tempBuf
    lea rdx, g_exePath
    call wcscat_p
    lea rcx, g_tempBuf
    lea rdx, str_cmdSuffixFile
    call wcscat_p

    ; Calculate file command string byte size
    lea rcx, g_tempBuf
    call wcslen_p
    lea r14, [rax+1]
    shl r14, 1                  ; R14 = file command string byte size

    ; Calculate file menu text byte size
    lea rcx, str_ctxTextFile
    call wcslen_p
    lea rbx, [rax+1]
    shl rbx, 1                  ; RBX = file menu text byte size

    ; --- exefile\shell\CMDT (parent key) ---
    lea rax, [rsp+80]
    mov qword ptr [rsp+64], rax
    lea rax, [rsp+72]
    mov qword ptr [rsp+56], rax
    mov qword ptr [rsp+48], 0
    mov qword ptr [rsp+40], KEY_WRITE
    mov qword ptr [rsp+32], 0
    xor r9, r9
    xor r8d, r8d
    lea rdx, str_ctxKeyExe
    mov ecx, HKEY_CLASSES_ROOT
    call RegCreateKeyExW
    test eax, eax
    jnz ctx_install_done

    ; Set default value = "Run as TrustedInstaller"
    mov qword ptr [rsp+40], rbx
    lea rax, str_ctxTextFile
    mov qword ptr [rsp+32], rax
    mov r9d, REG_SZ
    xor r8d, r8d
    xor edx, edx
    mov rcx, [rsp+72]
    call RegSetValueExW

    ; Set Icon = shell32.dll,104
    mov qword ptr [rsp+40], r13
    lea rax, str_iconPath
    mov qword ptr [rsp+32], rax
    mov r9d, REG_SZ
    xor r8d, r8d
    lea rdx, str_iconVal
    mov rcx, [rsp+72]
    call RegSetValueExW

    mov rcx, [rsp+72]
    call RegCloseKey

    ; --- exefile\shell\CMDT\command (command subkey) ---
    lea rax, [rsp+80]
    mov qword ptr [rsp+64], rax
    lea rax, [rsp+72]
    mov qword ptr [rsp+56], rax
    mov qword ptr [rsp+48], 0
    mov qword ptr [rsp+40], KEY_WRITE
    mov qword ptr [rsp+32], 0
    xor r9, r9
    xor r8d, r8d
    lea rdx, str_ctxKeyCmdExe
    mov ecx, HKEY_CLASSES_ROOT
    call RegCreateKeyExW
    test eax, eax
    jnz ctx_install_done

    mov qword ptr [rsp+40], r14
    lea rax, g_tempBuf
    mov qword ptr [rsp+32], rax
    mov r9d, REG_SZ
    xor r8d, r8d
    xor edx, edx
    mov rcx, [rsp+72]
    call RegSetValueExW

    mov rcx, [rsp+72]
    call RegCloseKey

    ; --- lnkfile\shell\CMDT (parent key) ---
    lea rax, [rsp+80]
    mov qword ptr [rsp+64], rax
    lea rax, [rsp+72]
    mov qword ptr [rsp+56], rax
    mov qword ptr [rsp+48], 0
    mov qword ptr [rsp+40], KEY_WRITE
    mov qword ptr [rsp+32], 0
    xor r9, r9
    xor r8d, r8d
    lea rdx, str_ctxKeyLnk
    mov ecx, HKEY_CLASSES_ROOT
    call RegCreateKeyExW
    test eax, eax
    jnz ctx_install_done

    ; Set default value = "Run as TrustedInstaller"
    mov qword ptr [rsp+40], rbx
    lea rax, str_ctxTextFile
    mov qword ptr [rsp+32], rax
    mov r9d, REG_SZ
    xor r8d, r8d
    xor edx, edx
    mov rcx, [rsp+72]
    call RegSetValueExW

    ; Set Icon = shell32.dll,104
    mov qword ptr [rsp+40], r13
    lea rax, str_iconPath
    mov qword ptr [rsp+32], rax
    mov r9d, REG_SZ
    xor r8d, r8d
    lea rdx, str_iconVal
    mov rcx, [rsp+72]
    call RegSetValueExW

    mov rcx, [rsp+72]
    call RegCloseKey

    ; --- lnkfile\shell\CMDT\command (command subkey) ---
    lea rax, [rsp+80]
    mov qword ptr [rsp+64], rax
    lea rax, [rsp+72]
    mov qword ptr [rsp+56], rax
    mov qword ptr [rsp+48], 0
    mov qword ptr [rsp+40], KEY_WRITE
    mov qword ptr [rsp+32], 0
    xor r9, r9
    xor r8d, r8d
    lea rdx, str_ctxKeyCmdLnk
    mov ecx, HKEY_CLASSES_ROOT
    call RegCreateKeyExW
    test eax, eax
    jnz ctx_install_done

    mov qword ptr [rsp+40], r14
    lea rax, g_tempBuf
    mov qword ptr [rsp+32], rax
    mov r9d, REG_SZ
    xor r8d, r8d
    xor edx, edx
    mov rcx, [rsp+72]
    call RegSetValueExW

    mov rcx, [rsp+72]
    call RegCloseKey

ctx_install_done:
    add rsp, 104
    pop r14
    pop r13
    pop r12
    pop rbx
    ret
InstallContextMenu endp

; ==============================================================================
; UninstallContextMenu - Remove Explorer context menu entries
;
; Purpose: Deletes all CMDT registry keys from HKEY_CLASSES_ROOT that were
;          created by InstallContextMenu. Removes context menu entries for
;          directories, executable files, and shortcut files.
;
; Registry locations deleted (in order, children first):
;   - Directory\Background\shell\CMDT\command
;   - Directory\Background\shell\CMDT
;   - Directory\shell\CMDT\command
;   - Directory\shell\CMDT
;   - exefile\shell\CMDT\command
;   - exefile\shell\CMDT
;   - lnkfile\shell\CMDT\command
;   - lnkfile\shell\CMDT
;
; Parameters: None
;
; Returns: None (ignores deletion errors)
;
; Stack frame: 40 bytes (32 shadow space + 8 alignment)
;
; Note: Keys must be deleted in order from leaf nodes to parent nodes,
;       as Windows registry does not allow deletion of keys with subkeys.
; ==============================================================================
UninstallContextMenu proc
    sub rsp, 40                 ; 32 shadow + 8 alignment

    ; Delete Directory\Background entries (leaf keys first)
    lea rdx, str_ctxKeyCmdBg
    mov ecx, HKEY_CLASSES_ROOT
    call RegDeleteKeyW

    lea rdx, str_ctxKeyBg
    mov ecx, HKEY_CLASSES_ROOT
    call RegDeleteKeyW

    ; Delete Directory entries (leaf keys first)
    lea rdx, str_ctxKeyCmdDir
    mov ecx, HKEY_CLASSES_ROOT
    call RegDeleteKeyW

    lea rdx, str_ctxKeyDir
    mov ecx, HKEY_CLASSES_ROOT
    call RegDeleteKeyW

    ; Delete exefile entries (leaf keys first)
    lea rdx, str_ctxKeyCmdExe
    mov ecx, HKEY_CLASSES_ROOT
    call RegDeleteKeyW

    lea rdx, str_ctxKeyExe
    mov ecx, HKEY_CLASSES_ROOT
    call RegDeleteKeyW

    ; Delete lnkfile entries (leaf keys first)
    lea rdx, str_ctxKeyCmdLnk
    mov ecx, HKEY_CLASSES_ROOT
    call RegDeleteKeyW

    lea rdx, str_ctxKeyLnk
    mov ecx, HKEY_CLASSES_ROOT
    call RegDeleteKeyW

    add rsp, 40
    ret
UninstallContextMenu endp

; ==============================================================================
; InstallShift - Set sethc.exe IFEO debugger hook
;
; Purpose: Creates an Image File Execution Options registry entry for sethc.exe
;          that redirects execution to CMDT. When Sticky Keys is triggered
;          (5x Shift at login screen), CMDT launches cmd.exe as TrustedInstaller
;          instead of sethc.exe. Also adds Defender process exclusions for both
;          the exe itself and cmd.exe.
;
; Registry location:
;   HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\
;     Image File Execution Options\sethc.exe
;   Value: Debugger = <exename> -cli -new cmd.exe
;
; Parameters: None (uses global g_exePath and g_cmdBuf buffers)
;
; Returns: None
;
; Stack frame: 104 bytes
; ==============================================================================
InstallShift proc
    push rbx
    push r12
    sub rsp, 104
    ; [rsp+72] = hKey, [rsp+80] = dwDisp

    ; Get our exe path
    lea rdx, g_exePath
    mov r8d, 260
    xor ecx, ecx
    call GetModuleFileNameW

    ; Get just the filename (e.g. "cmdt.exe")
    call GetExeFileName
    mov r12, rax

    ; Decrypt 'Add' method name for WMI
    lea rcx, str_wmi_add
    lea rdx, g_tempBuf
    call DecryptWideStr

    ; Add CMDT executable to exclusions
    mov rcx, r12
    lea rdx, g_tempBuf
    call ManageDefenderExclusion

    ; Add cmd.exe to exclusions
    lea rcx, str_cmd_exe
    lea rdx, g_decryptBuf
    call DecryptWideStr

    lea rcx, g_decryptBuf
    lea rdx, g_tempBuf          ; method name is still in g_tempBuf
    call ManageDefenderExclusion

    ; Build IFEO debugger value in g_tempBuf: <full_path> -cli -new cmd.exe
    lea rcx, g_tempBuf
    lea rdx, g_exePath
    call wcscpy_p
    lea rcx, g_tempBuf
    lea rdx, str_shiftSuffix
    call wcscat_p

    ; Calculate byte size (chars+1) * 2
    lea rcx, g_tempBuf
    call wcslen_p
    lea rbx, [rax+1]
    shl rbx, 1

    ; Decrypt IFEO registry key path
    lea rcx, str_ifeoKey_enc
    lea rdx, g_decryptBuf
    call DecryptWideStr

    ; Create/open IFEO sethc.exe key under HKLM
    lea rax, [rsp+80]
    mov qword ptr [rsp+64], rax     ; lpdwDisposition
    lea rax, [rsp+72]
    mov qword ptr [rsp+56], rax     ; phkResult
    mov qword ptr [rsp+48], 0       ; lpSecurityAttributes
    mov qword ptr [rsp+40], KEY_WRITE
    mov qword ptr [rsp+32], 0       ; dwOptions
    xor r9, r9                      ; lpClass
    xor r8d, r8d                    ; Reserved
    lea rdx, g_decryptBuf
    mov ecx, HKEY_LOCAL_MACHINE
    call RegCreateKeyExW
    test eax, eax
    jnz shift_install_done

    ; Set Debugger = command string
    mov qword ptr [rsp+40], rbx     ; cbData
    lea rax, g_tempBuf
    mov qword ptr [rsp+32], rax     ; lpData
    mov r9d, REG_SZ                 ; dwType
    xor r8d, r8d                    ; Reserved
    
    ; Decrypt Debugger value name
    lea rcx, str_debuggerVal_enc
    lea rdx, g_filePath
    call DecryptWideStr
    
    lea rdx, g_filePath             ; lpValueName = "Debugger"
    mov rcx, [rsp+72]              ; hKey
    call RegSetValueExW

    mov rcx, [rsp+72]
    call RegCloseKey

shift_install_done:
    add rsp, 104
    pop r12
    pop rbx
    ret
InstallShift endp

; ==============================================================================
; UninstallShift - Remove sethc.exe IFEO Debugger value
;
; Purpose: Deletes only the "Debugger" value from the Image File Execution
;          Options\sethc.exe registry key, restoring normal Sticky Keys
;          behavior. Also removes Defender process exclusions for both the
;          exe itself and cmd.exe.
;
; Parameters: None
;
; Returns: None
;
; Stack frame: 48 bytes
; ==============================================================================
UninstallShift proc
    push r12
    sub rsp, 48
    ; [rsp+40] = hKey

    ; Decrypt IFEO registry key path
    lea rcx, str_ifeoKey_enc
    lea rdx, g_decryptBuf
    call DecryptWideStr

    ; Open IFEO sethc.exe key
    lea rax, [rsp+40]
    mov qword ptr [rsp+32], rax     ; phkResult = &hKey
    mov r9d, KEY_WRITE              ; samDesired
    xor r8d, r8d                    ; ulOptions
    lea rdx, g_decryptBuf           ; lpSubKey
    mov ecx, HKEY_LOCAL_MACHINE     ; hKey
    call RegOpenKeyExW
    test eax, eax
    jnz unshift_ps                  ; Key doesn't exist, skip to PS cleanup

    ; Decrypt Debugger value name
    lea rcx, str_debuggerVal_enc
    lea rdx, g_filePath
    call DecryptWideStr

    ; Delete only the Debugger value
    lea rdx, g_filePath
    mov rcx, [rsp+40]
    call RegDeleteValueW

    ; Close key
    mov rcx, [rsp+40]
    call RegCloseKey

unshift_ps:
    ; Get our exe path for dynamic filename
    lea rdx, g_exePath
    mov r8d, 260
    xor ecx, ecx
    call GetModuleFileNameW

    ; Get just the filename
    call GetExeFileName
    mov r12, rax

    ; Decrypt 'Remove' method name for WMI
    lea rcx, str_wmi_rem
    lea rdx, g_tempBuf
    call DecryptWideStr

    ; Remove CMDT executable from exclusions
    mov rcx, r12
    lea rdx, g_tempBuf
    call ManageDefenderExclusion

    ; Remove cmd.exe from exclusions
    lea rcx, str_cmd_exe
    lea rdx, g_decryptBuf
    call DecryptWideStr

    lea rcx, g_decryptBuf
    lea rdx, g_tempBuf          ; method name is still in g_tempBuf
    call ManageDefenderExclusion

    add rsp, 48
    pop r12
    ret
UninstallShift endp

; ==============================================================================
; ManageDefenderExclusion - Adds or removes a process exclusion in MS Defender
;
; Purpose: Replaces RunPsCommand by directly interacting with the WMI interface.
;
; Parameters:
;   RCX = Pointer to the executable name (wide string) to exclude
;   RDX = Pointer to the method name (wide string, "Add" or "Remove")
; ==============================================================================
ManageDefenderExclusion proc frame
    push rbp
    .pushreg rbp
    mov rbp, rsp
    .setframe rbp, 0
    push rbx
    .pushreg rbx
    push r12
    .pushreg r12
    push r13
    .pushreg r13
    push r14
    .pushreg r14
    push r15
    .pushreg r15
    sub rsp, 248                ; 248 maintains 16-byte alignment
    .allocstack 248
    .endprolog

    mov r12, rcx                ; r12 = pszProcessName
    mov r13, rdx                ; r13 = pszMethodName

    ; Zero initialize COM pointers
    mov qword ptr [rbp-40], 0   ; pLoc
    mov qword ptr [rbp-48], 0   ; pSvc
    mov qword ptr [rbp-56], 0   ; pClass
    mov qword ptr [rbp-64], 0   ; pInParamsDef
    mov qword ptr [rbp-72], 0   ; pClassInstance

    ; CoInitializeEx(NULL, COINIT_MULTITHREADED)
    xor ecx, ecx
    xor edx, edx
    call CoInitializeEx

    ; CoInitializeSecurity
    xor ecx, ecx
    mov edx, -1
    xor r8d, r8d
    xor r9d, r9d
    mov qword ptr [rsp+32], 0
    mov qword ptr [rsp+40], 3   ; RPC_C_IMP_LEVEL_IMPERSONATE
    mov qword ptr [rsp+48], 0
    mov dword ptr [rsp+56], 0
    mov qword ptr [rsp+64], 0
    call CoInitializeSecurity
    cmp eax, 80010119h          ; RPC_E_TOO_LATE
    je init_sec_ok
    test eax, eax
    jl wmi_cleanup
init_sec_ok:
    ; COM security already initialized, proceed with WMI connection

    ; CoCreateInstance(CLSID_WbemLocator)
    lea rcx, CLSID_WbemLocator
    xor edx, edx
    mov r8d, 1                  ; CLSCTX_INPROC_SERVER
    lea r9, IID_IWbemLocator
    lea rax, [rbp-40]
    mov qword ptr [rsp+32], rax
    call CoCreateInstance
    test eax, eax
    jl wmi_cleanup              ; FAILED(hr)

    ; pLoc->ConnectServer("ROOT\Microsoft\Windows\Defender")
    lea rcx, str_wmi_namespace
    lea rdx, g_decryptBuf
    call DecryptWideStr
    lea rcx, g_decryptBuf
    call SysAllocString
    mov rbx, rax                ; rbx = bstrNamespace

    mov rcx, [rbp-40]           ; pLoc
    mov rdx, rbx                ; strNetworkResource
    xor r8d, r8d                ; strUser
    xor r9d, r9d                ; strPassword
    mov qword ptr [rsp+32], 0   ; strLocale
    mov qword ptr [rsp+40], 0   ; lSecurityFlags
    mov qword ptr [rsp+48], 0   ; strAuthority
    mov qword ptr [rsp+56], 0   ; pCtx
    lea rax, [rbp-48]
    mov qword ptr [rsp+64], rax ; ppNamespace
    mov rax, [rcx]              ; pLoc vtable
    call qword ptr [rax+24]     ; ConnectServer (index 3)

    mov rcx, rbx
    mov r14d, eax               ; save hr
    call SysFreeString
    test r14d, r14d
    jl wmi_cleanup

    ; CoSetProxyBlanket
    mov rcx, [rbp-48]           ; pSvc
    mov edx, 10                 ; RPC_C_AUTHN_WINNT
    xor r8d, r8d                ; RPC_C_AUTHZ_NONE
    xor r9d, r9d                ; pServerPrincName
    mov qword ptr [rsp+32], 3   ; RPC_C_AUTHN_LEVEL_CALL
    mov qword ptr [rsp+40], 3   ; RPC_C_IMP_LEVEL_IMPERSONATE
    mov qword ptr [rsp+48], 0   ; pAuthInfo
    mov dword ptr [rsp+56], 0   ; EOAC_NONE
    call CoSetProxyBlanket

    ; GetObject("MSFT_MpPreference")
    lea rcx, str_wmi_class
    lea rdx, g_decryptBuf
    call DecryptWideStr
    lea rcx, g_decryptBuf
    call SysAllocString
    mov rbx, rax

    mov rcx, [rbp-48]           ; pSvc
    mov rdx, rbx                ; strObjectPath
    xor r8d, r8d                ; lFlags
    xor r9d, r9d                ; pCtx
    lea rax, [rbp-56]
    mov qword ptr [rsp+32], rax ; ppObject
    mov qword ptr [rsp+40], 0   ; ppCallResult
    mov rax, [rcx]
    call qword ptr [rax+48]     ; GetObject (index 6)

    mov rcx, rbx
    mov r14d, eax
    call SysFreeString
    test r14d, r14d
    jl wmi_cleanup

    ; GetMethod(pszMethodName)
    mov rcx, r13                ; method name is already unencrypted
    call SysAllocString
    mov rbx, rax

    mov rcx, [rbp-56]           ; pClass
    mov rdx, rbx                ; strName
    xor r8d, r8d                ; lFlags
    lea r9, [rbp-64]            ; ppInSignature
    mov qword ptr [rsp+32], 0   ; ppOutSignature
    mov rax, [rcx]
    call qword ptr [rax+152]    ; GetMethod (index 19)

    mov rcx, rbx
    mov r14d, eax
    call SysFreeString
    test r14d, r14d
    jl wmi_cleanup

    ; SpawnInstance
    mov rcx, [rbp-64]           ; pInParamsDef
    xor edx, edx                ; lFlags
    lea r8, [rbp-72]            ; ppNewInstance
    mov rax, [rcx]
    call qword ptr [rax+120]    ; SpawnInstance (index 15)
    test eax, eax
    jl wmi_cleanup

    ; Prepare SAFEARRAY for Variant
    lea rcx, [rbp-104]          ; VARIANT var
    call VariantInit

    ; Create SAFEARRAY
    mov dword ptr [rbp-120], 1  ; bounds.cElements = 1
    mov dword ptr [rbp-116], 0  ; bounds.lLbound = 0
    mov ecx, 8                  ; VT_BSTR
    mov edx, 1                  ; cDims
    lea r8, [rbp-120]           ; rgsabound
    call SafeArrayCreate
    mov r15, rax                ; r15 = parray

    ; Put element in SAFEARRAY
    mov dword ptr [rbp-128], 0  ; ix[0] = 0
    mov rcx, r12                ; pszProcessName
    call SysAllocString
    mov rbx, rax                ; BSTR val

    mov rcx, r15                ; parray
    lea rdx, [rbp-128]          ; rgIndices
    mov r8, rbx                 ; pv
    call SafeArrayPutElement

    mov rcx, rbx
    call SysFreeString

    ; Set up VARIANT
    mov word ptr [rbp-104], 2008h ; VT_ARRAY | VT_BSTR
    mov qword ptr [rbp-96], r15   ; var.parray = parray

    ; Put property
    lea rcx, str_wmi_prop
    lea rdx, g_decryptBuf
    call DecryptWideStr
    lea rcx, g_decryptBuf
    call SysAllocString
    mov rbx, rax

    mov rcx, [rbp-72]           ; pClassInstance
    mov rdx, rbx                ; strName
    xor r8d, r8d                ; lFlags
    lea r9, [rbp-104]           ; pVal = &var
    mov qword ptr [rsp+32], 0   ; Type
    mov rax, [rcx]
    call qword ptr [rax+40]     ; Put (index 5)

    mov rcx, rbx
    call SysFreeString

    lea rcx, [rbp-104]
    call VariantClear

    ; ExecMethod
    lea rcx, str_wmi_class
    lea rdx, g_decryptBuf
    call DecryptWideStr
    lea rcx, g_decryptBuf
    call SysAllocString
    mov r14, rax                ; strObjectPath

    mov rcx, r13
    call SysAllocString
    mov rbx, rax                ; strMethodName

    mov rcx, [rbp-48]           ; pSvc
    mov rdx, r14                ; strObjectPath
    mov r8, rbx                 ; strMethodName
    xor r9d, r9d                ; lFlags
    mov qword ptr [rsp+32], 0   ; pCtx
    mov rax, [rbp-72]
    mov qword ptr [rsp+40], rax ; pInParams
    mov qword ptr [rsp+48], 0   ; ppOutParams
    mov qword ptr [rsp+56], 0   ; ppCallResult
    mov rax, [rcx]
    call qword ptr [rax+192]    ; ExecMethod (index 24)

    mov rcx, r14
    call SysFreeString
    mov rcx, rbx
    call SysFreeString

wmi_cleanup:
    mov rcx, [rbp-72]
    test rcx, rcx
    jz @F
    mov rax, [rcx]
    call qword ptr [rax+16]     ; Release (index 2)
@@:
    mov rcx, [rbp-64]
    test rcx, rcx
    jz @F
    mov rax, [rcx]
    call qword ptr [rax+16]
@@:
    mov rcx, [rbp-56]
    test rcx, rcx
    jz @F
    mov rax, [rcx]
    call qword ptr [rax+16]
@@:
    mov rcx, [rbp-48]
    test rcx, rcx
    jz @F
    mov rax, [rcx]
    call qword ptr [rax+16]
@@:
    mov rcx, [rbp-40]
    test rcx, rcx
    jz @F
    mov rax, [rcx]
    call qword ptr [rax+16]
@@:
    call CoUninitialize

    add rsp, 248
    pop r15
    pop r14
    pop r13
    pop r12
    pop rbx
    pop rbp
    ret
ManageDefenderExclusion endp

end
