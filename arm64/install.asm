; ==============================================================================
; CMDT - Run as TrustedInstaller
; Installation / Hook Management Module
;
; Author: Marek Wesolowski (wesmar)
; Ported to: Native ARM64 (armasm64) for Windows on ARM
;
; Purpose: Owns every persistent change CMDT can make to the host system:
;          Explorer context-menu entries, the sethc.exe Sticky-Keys IFEO hook
;          used for login-screen rescue access, and the matching Defender
;          process exclusions installed alongside the hook.
;
; Exported routines:
;   InstallContextMenu      - Register HKCR shell entries for "Run as TI"
;   UninstallContextMenu    - Remove the above entries
;   InstallShift            - Install sethc.exe IFEO Debugger hook + exclusions
;   UninstallShift          - Remove the IFEO Debugger value + exclusions
;   ManageDefenderExclusion - Add/Remove a process from Defender exclusions
;                             via the MSFT_MpPreference WMI class
;   GetExeFileName          - Return pointer to leaf filename of g_exePath
;
; ARM64 Port Notes:
;   - x64 shadow space eliminated; ARM64 ABI passes first 8 args in x0-x7.
;   - COM vtable calls use LDR + BLR pattern (no indirect CALL syntax).
;   - Callee-saved registers: x19-x26 (vs rbx,rsi,rdi,r12-r15 on x64).
;   - Encrypted strings preserved byte-for-byte (XOR key 0xAA unchanged).
; ==============================================================================

; ==============================================================================
; EXTERNAL REFERENCES
; ==============================================================================

; Win32 / COM / WMI dependencies
    IMPORT GetModuleFileNameW
    IMPORT RegCreateKeyExW
    IMPORT RegSetValueExW
    IMPORT RegDeleteKeyW
    IMPORT RegDeleteValueW
    IMPORT RegOpenKeyExW
    IMPORT RegCloseKey
    IMPORT CoInitializeEx
    IMPORT CoInitializeSecurity
    IMPORT CoCreateInstance
    IMPORT CoSetProxyBlanket
    IMPORT CoUninitialize
    IMPORT SysAllocString
    IMPORT SysFreeString
    IMPORT VariantInit
    IMPORT VariantClear
    IMPORT SafeArrayCreate
    IMPORT SafeArrayPutElement

; In-project helpers
    IMPORT DecryptWideStr
    IMPORT wcscpy_p
    IMPORT wcscat_p
    IMPORT wcslen_p

; Global buffers (defined in main.asm)
    IMPORT g_exePath
    IMPORT g_tempBuf
    IMPORT g_decryptBuf
    IMPORT g_filePath

; ==============================================================================
; WINDOWS API CONSTANTS
; ==============================================================================

HKEY_CLASSES_ROOT   EQU 0x80000000
HKEY_LOCAL_MACHINE  EQU 0x80000002
KEY_WRITE           EQU 0x00020006
REG_SZ              EQU 1
COINIT_MULTITHREADED EQU 0
CLSCTX_INPROC_SERVER EQU 1
RPC_E_TOO_LATE      EQU 0x80010119
RPC_C_AUTHN_WINNT   EQU 10
RPC_C_AUTHZ_NONE    EQU 0
RPC_C_AUTHN_LEVEL_CALL EQU 3
RPC_C_AUTHN_LEVEL_DEFAULT EQU 0
RPC_C_IMP_LEVEL_IMPERSONATE EQU 3
EOAC_NONE           EQU 0
VT_BSTR             EQU 8
VT_ARRAY            EQU 0x2000
VT_ARRAY_BSTR       EQU 0x2008

; COM vtable method offsets (index * 8 bytes per pointer)
VT_RELEASE          EQU 16      ; IUnknown::Release (index 2)
VT_CONNECTSERVER    EQU 24      ; IWbemLocator::ConnectServer (index 3)
VT_PUT              EQU 40      ; IWbemClassObject::Put (index 5)
VT_GETOBJECT        EQU 48      ; IWbemServices::GetObject (index 6)
VT_SPAWNINSTANCE    EQU 120     ; IWbemClassObject::SpawnInstance (index 15)
VT_GETMETHOD        EQU 152     ; IWbemClassObject::GetMethod (index 19)
VT_EXECMETHOD       EQU 192     ; IWbemServices::ExecMethod (index 24)

; ==============================================================================
; CONSTANT STRING DATA (READ-ONLY)
; ==============================================================================
    AREA |.const|, DATA, READONLY, ALIGN=3

; --- Explorer context-menu registry paths (HKCR-relative) ---
str_ctxKeyBg
    DCW 'D','i','r','e','c','t','o','r','y',0x5C,'B','a','c','k','g','r','o','u','n','d'
    DCW 0x5C,'s','h','e','l','l',0x5C,'C','M','D','T',0
str_ctxKeyCmdBg
    DCW 'D','i','r','e','c','t','o','r','y',0x5C,'B','a','c','k','g','r','o','u','n','d'
    DCW 0x5C,'s','h','e','l','l',0x5C,'C','M','D','T',0x5C,'c','o','m','m','a','n','d',0
str_ctxKeyDir
    DCW 'D','i','r','e','c','t','o','r','y',0x5C,'s','h','e','l','l',0x5C,'C','M','D','T',0
str_ctxKeyCmdDir
    DCW 'D','i','r','e','c','t','o','r','y',0x5C,'s','h','e','l','l',0x5C,'C','M','D','T'
    DCW 0x5C,'c','o','m','m','a','n','d',0
str_ctxKeyExe
    DCW 'e','x','e','f','i','l','e',0x5C,'s','h','e','l','l',0x5C,'C','M','D','T',0
str_ctxKeyCmdExe
    DCW 'e','x','e','f','i','l','e',0x5C,'s','h','e','l','l',0x5C,'C','M','D','T'
    DCW 0x5C,'c','o','m','m','a','n','d',0
str_ctxKeyLnk
    DCW 'l','n','k','f','i','l','e',0x5C,'s','h','e','l','l',0x5C,'C','M','D','T',0
str_ctxKeyCmdLnk
    DCW 'l','n','k','f','i','l','e',0x5C,'s','h','e','l','l',0x5C,'C','M','D','T'
    DCW 0x5C,'c','o','m','m','a','n','d',0

; --- Menu labels shown in Explorer ---
str_ctxTextDir
    DCW 'O','p','e','n',' ','C','M','D',' ','a','s',' ','T','r','u','s','t','e','d'
    DCW 'I','n','s','t','a','l','l','e','r',0
str_ctxTextFile
    DCW 'R','u','n',' ','a','s',' ','T','r','u','s','t','e','d','I','n','s','t','a','l'
    DCW 'l','e','r',0

; --- Icon (UAC shield from shell32.dll resource #104) ---
str_iconVal         DCW 'I','c','o','n',0
str_iconPath        DCW 's','h','e','l','l','3','2','.','d','l','l',',','1','0','4',0

; --- Command-line templates for the registered shell entries ---
str_cmdQuote        DCW '"',0
str_cmdSuffixDir
    DCW '"',' ','-','c','l','i',' ','-','n','e','w',' ','c','m','d','.','e','x','e'
    DCW ' ','/','k',' ','c','d',' ','/','d',' ','"','%','V','"',0
str_cmdSuffixFile   DCW '"',' ','"','%','1','"',0

; --- Suffix appended to exe path for the sethc.exe IFEO Debugger value ---
str_shiftSuffix     DCW ' ','-','c','l','i',' ','-','n','e','w',' ','c','m','d','.','e','x','e',0

; --- XOR-obfuscated strings (key = 0xAA) ---
; The IFEO and WMI paths are kept encrypted in the binary so that simple
; `strings` scans don't reveal the hooked location or the Defender-exclusion
; mechanism. They are decrypted on demand via DecryptWideStr.

str_ifeoKey_enc
    DCB 0xF9,0xAA,0xE5,0xAA,0xEC,0xAA,0xFE,0xAA,0xFD,0xAA,0xEB,0xAA,0xF8,0xAA,0xEF,0xAA
    DCB 0xF6,0xAA,0xE7,0xAA,0xC3,0xAA,0xC9,0xAA,0xD8,0xAA,0xC5,0xAA,0xD9,0xAA,0xC5,0xAA
    DCB 0xCC,0xAA,0xDE,0xAA,0xF6,0xAA,0xFD,0xAA,0xC3,0xAA,0xC4,0xAA,0xCE,0xAA,0xC5,0xAA
    DCB 0xDD,0xAA,0xD9,0xAA,0x8A,0xAA,0xE4,0xAA,0xFE,0xAA,0xF6,0xAA,0xE9,0xAA,0xDF,0xAA
    DCB 0xD8,0xAA,0xD8,0xAA,0xCF,0xAA,0xC4,0xAA,0xDE,0xAA,0xFC,0xAA,0xCF,0xAA,0xD8,0xAA
    DCB 0xD9,0xAA,0xC3,0xAA,0xC5,0xAA,0xC4,0xAA,0xF6,0xAA,0xE3,0xAA,0xC7,0xAA,0xCB,0xAA
    DCB 0xCD,0xAA,0xCF,0xAA,0x8A,0xAA,0xEC,0xAA,0xC3,0xAA,0xC6,0xAA,0xCF,0xAA,0x8A,0xAA
    DCB 0xEF,0xAA,0xD2,0xAA,0xCF,0xAA,0xC9,0xAA,0xDF,0xAA,0xDE,0xAA,0xC3,0xAA,0xC5,0xAA
    DCB 0xC4,0xAA,0x8A,0xAA,0xE5,0xAA,0xDA,0xAA,0xDE,0xAA,0xC3,0xAA,0xC5,0xAA,0xC4,0xAA
    DCB 0xD9,0xAA,0xF6,0xAA,0xD9,0xAA,0xCF,0xAA,0xDE,0xAA,0xC2,0xAA,0xC9,0xAA,0x84,0xAA
    DCB 0xCF,0xAA,0xD2,0xAA,0xCF,0xAA,0xAA,0xAA

str_debuggerVal_enc
    DCB 0xEE,0xAA,0xCF,0xAA,0xC8,0xAA,0xDF,0xAA,0xCD,0xAA,0xCD,0xAA,0xCF,0xAA,0xD8,0xAA
    DCB 0xAA,0xAA

; --- WMI plumbing ---
CLSID_WbemLocator
    DCB 0x11,0xF8,0x90,0x45,0x3A,0x1D,0xD0,0x11,0x89,0x1F,0x00,0xAA,0x00,0x4B,0x2E,0x24
IID_IWbemLocator
    DCB 0x87,0xA6,0x12,0xDC,0x7F,0x73,0xCF,0x11,0x88,0x4D,0x00,0xAA,0x00,0x4B,0x2E,0x24

str_wmi_namespace
    DCB 0xF8,0xAA,0xE5,0xAA,0xE5,0xAA,0xFE,0xAA,0xF6,0xAA,0xE7,0xAA,0xC3,0xAA,0xC9,0xAA
    DCB 0xD8,0xAA,0xC5,0xAA,0xD9,0xAA,0xC5,0xAA,0xCC,0xAA,0xDE,0xAA,0xF6,0xAA,0xFD,0xAA
    DCB 0xC3,0xAA,0xC4,0xAA,0xCE,0xAA,0xC5,0xAA,0xDD,0xAA,0xD9,0xAA,0xF6,0xAA,0xEE,0xAA
    DCB 0xCF,0xAA,0xCC,0xAA,0xCF,0xAA,0xC4,0xAA,0xCE,0xAA,0xCF,0xAA,0xD8,0xAA,0xAA,0xAA

str_wmi_class
    DCB 0xE7,0xAA,0xF9,0xAA,0xEC,0xAA,0xFE,0xAA,0xF5,0xAA,0xE7,0xAA,0xDA,0xAA,0xFA,0xAA
    DCB 0xD8,0xAA,0xCF,0xAA,0xCC,0xAA,0xCF,0xAA,0xD8,0xAA,0xCF,0xAA,0xC4,0xAA,0xC9,0xAA
    DCB 0xCF,0xAA,0xAA,0xAA

str_wmi_add         DCB 0xEB,0xAA,0xCE,0xAA,0xCE,0xAA,0xAA,0xAA
str_wmi_rem         DCB 0xF8,0xAA,0xCF,0xAA,0xC7,0xAA,0xC5,0xAA,0xDC,0xAA,0xCF,0xAA,0xAA,0xAA

str_wmi_prop
    DCB 0xEF,0xAA,0xD2,0xAA,0xC9,0xAA,0xC6,0xAA,0xDF,0xAA,0xD9,0xAA,0xC3,0xAA,0xC5,0xAA
    DCB 0xC4,0xAA,0xFA,0xAA,0xD8,0xAA,0xC5,0xAA,0xC9,0xAA,0xCF,0xAA,0xD9,0xAA,0xD9,0xAA
    DCB 0xAA,0xAA

; "ExclusionPath" -- used to exclude the portable CMDT image itself.
str_wmi_path
    DCB 0xEF,0xAA,0xD2,0xAA,0xC9,0xAA,0xC6,0xAA,0xDF,0xAA,0xD9,0xAA,0xC3,0xAA
    DCB 0xC5,0xAA,0xC4,0xAA,0xFA,0xAA,0xCB,0xAA,0xDE,0xAA,0xC2,0xAA,0xAA,0xAA

str_cmd_exe         DCB 0xC9,0xAA,0xC7,0xAA,0xCE,0xAA,0x84,0xAA,0xCF,0xAA,0xD2,0xAA,0xCF,0xAA,0xAA,0xAA

; ==============================================================================
; CODE SECTION
; ==============================================================================
    AREA |.text|, CODE, READONLY, ALIGN=4

    EXPORT GetExeFileName
    EXPORT InstallContextMenu
    EXPORT UninstallContextMenu
    EXPORT InstallShift
    EXPORT UninstallShift
    EXPORT ManageDefenderExclusion

; ==============================================================================
; GetExeFileName - Locate the leaf filename component of g_exePath
;
; Purpose: Walks g_exePath looking for the last backslash and returns a
;          pointer to the character following it (the filename without
;          directory). Used by the Sticky-Keys hook so the IFEO value can
;          reference the executable by its short name even if cmdt is
;          installed under a non-default directory.
;
; Parameters: None (reads g_exePath)
;
; Returns:
;   x0 = Pointer into g_exePath at the start of the filename component.
;        Points at g_exePath itself if no backslash is found.
;
; Modifies: x0, x1, w2
; ==============================================================================
GetExeFileName PROC
    ADRP x0, g_exePath
    ADD x0, x0, g_exePath
    MOV x1, x0                  ; x1 = scan pointer

gef_loop
    LDRH w2, [x1]
    CBZ w2, gef_done            ; End of string
    CMP w2, #0x5C
    B.NE gef_next
    ADD x0, x1, #2             ; x0 = pointer after backslash

gef_next
    ADD x1, x1, #2
    B gef_loop

gef_done
    RET
GetExeFileName ENDP

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
;   - Default value: Menu text ("Open CMD as TrustedInstaller" or
;     "Run as TrustedInstaller")
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
; Stack frame: 16 bytes
;   [sp+0] = hKey (HANDLE)
;   [sp+8] = dwDisposition (DWORD)
;
; Register allocation:
;   x19 = directory command string byte size
;   x20 = directory menu text byte size
;   x21 = icon path byte size (shell32.dll,104)
;   x22 = file command string byte size
;   x23 = file menu text byte size
; ==============================================================================
InstallContextMenu PROC
    STP x29, x30, [sp, #-16]!
    MOV x29, sp
    STP x19, x20, [sp, #-16]!
    STP x21, x22, [sp, #-16]!
    STP x23, x24, [sp, #-16]!
    SUB sp, sp, #16

    ; Get our exe path for command strings
    ADRP x1, g_exePath
    ADD x1, x1, g_exePath
    MOV w2, #260
    MOV x0, XZR
    BL GetModuleFileNameW

    ; Build directory command string: "<exepath>" -cli -new cmd.exe /k cd /d "%V"
    ADRP x0, g_tempBuf
    ADD x0, x0, g_tempBuf
    ADRP x1, str_cmdQuote
    ADD x1, x1, str_cmdQuote
    BL wcscpy_p

    ADRP x0, g_tempBuf
    ADD x0, x0, g_tempBuf
    ADRP x1, g_exePath
    ADD x1, x1, g_exePath
    BL wcscat_p

    ADRP x0, g_tempBuf
    ADD x0, x0, g_tempBuf
    ADRP x1, str_cmdSuffixDir
    ADD x1, x1, str_cmdSuffixDir
    BL wcscat_p

    ; Calculate string byte sizes (characters * 2 + null terminator * 2)
    ADRP x0, g_tempBuf
    ADD x0, x0, g_tempBuf
    BL wcslen_p
    ADD x19, x0, #1
    LSL x19, x19, #1            ; x19 = directory command string byte size

    ADRP x0, str_ctxTextDir
    ADD x0, x0, str_ctxTextDir
    BL wcslen_p
    ADD x20, x0, #1
    LSL x20, x20, #1            ; x20 = directory menu text byte size

    ADRP x0, str_iconPath
    ADD x0, x0, str_iconPath
    BL wcslen_p
    ADD x21, x0, #1
    LSL x21, x21, #1            ; x21 = icon path byte size

    ; --- Directory\Background\shell\CMDT (parent key) ---
    SUB sp, sp, #16
    ADD x8, sp, #24             ; &dwDisposition (adjusted for sub)
    STR x8, [sp, #0]           ; lpdwDisposition
    ADD x7, sp, #16             ; &hKey (adjusted for sub)
    MOV x6, XZR                 ; lpSecurityAttributes = NULL
    MOVZ w5, #0x0006          ; samDesired  ; KEY_WRITE = 0x00020006
    MOVK w5, #0x0002, LSL #16
    MOV w4, #0                  ; dwOptions
    MOV x3, XZR                 ; lpClass = NULL
    MOV w2, #0                  ; Reserved
    ADRP x1, str_ctxKeyBg
    ADD x1, x1, str_ctxKeyBg
    MOVZ w0, #0x8000, LSL #16  ; HKEY_CLASSES_ROOT
    BL RegCreateKeyExW
    ADD sp, sp, #16
    CBNZ w0, icm_done           ; Error: skip rest

    ; Set default value = "Open CMD as TrustedInstaller"
    LDR x0, [sp, #0]           ; hKey
    MOV x1, XZR                 ; lpValueName = NULL (default)
    MOV w2, #0                  ; Reserved
    MOV w3, #REG_SZ             ; dwType
    ADRP x4, str_ctxTextDir
    ADD x4, x4, str_ctxTextDir
    MOV x5, x20                 ; cbData
    BL RegSetValueExW

    ; Set Icon = shell32.dll,104
    LDR x0, [sp, #0]
    ADRP x1, str_iconVal
    ADD x1, x1, str_iconVal
    MOV w2, #0
    MOV w3, #REG_SZ
    ADRP x4, str_iconPath
    ADD x4, x4, str_iconPath
    MOV x5, x21
    BL RegSetValueExW

    LDR x0, [sp, #0]
    BL RegCloseKey

    ; --- Directory\Background\shell\CMDT\command (command subkey) ---
    SUB sp, sp, #16
    ADD x8, sp, #24
    STR x8, [sp, #0]
    ADD x7, sp, #16
    MOV x6, XZR
    MOVZ w5, #0x0006  ; KEY_WRITE = 0x00020006
    MOVK w5, #0x0002, LSL #16
    MOV w4, #0
    MOV x3, XZR
    MOV w2, #0
    ADRP x1, str_ctxKeyCmdBg
    ADD x1, x1, str_ctxKeyCmdBg
    MOVZ w0, #0x8000, LSL #16
    BL RegCreateKeyExW
    ADD sp, sp, #16
    CBNZ w0, icm_done

    LDR x0, [sp, #0]
    MOV x1, XZR
    MOV w2, #0
    MOV w3, #REG_SZ
    ADRP x4, g_tempBuf
    ADD x4, x4, g_tempBuf
    MOV x5, x19
    BL RegSetValueExW

    LDR x0, [sp, #0]
    BL RegCloseKey

    ; --- Directory\shell\CMDT (parent key) ---
    SUB sp, sp, #16
    ADD x8, sp, #24
    STR x8, [sp, #0]
    ADD x7, sp, #16
    MOV x6, XZR
    MOVZ w5, #0x0006  ; KEY_WRITE = 0x00020006
    MOVK w5, #0x0002, LSL #16
    MOV w4, #0
    MOV x3, XZR
    MOV w2, #0
    ADRP x1, str_ctxKeyDir
    ADD x1, x1, str_ctxKeyDir
    MOVZ w0, #0x8000, LSL #16
    BL RegCreateKeyExW
    ADD sp, sp, #16
    CBNZ w0, icm_done

    LDR x0, [sp, #0]
    MOV x1, XZR
    MOV w2, #0
    MOV w3, #REG_SZ
    ADRP x4, str_ctxTextDir
    ADD x4, x4, str_ctxTextDir
    MOV x5, x20
    BL RegSetValueExW

    LDR x0, [sp, #0]
    ADRP x1, str_iconVal
    ADD x1, x1, str_iconVal
    MOV w2, #0
    MOV w3, #REG_SZ
    ADRP x4, str_iconPath
    ADD x4, x4, str_iconPath
    MOV x5, x21
    BL RegSetValueExW

    LDR x0, [sp, #0]
    BL RegCloseKey

    ; --- Directory\shell\CMDT\command (command subkey) ---
    SUB sp, sp, #16
    ADD x8, sp, #24
    STR x8, [sp, #0]
    ADD x7, sp, #16
    MOV x6, XZR
    MOVZ w5, #0x0006  ; KEY_WRITE = 0x00020006
    MOVK w5, #0x0002, LSL #16
    MOV w4, #0
    MOV x3, XZR
    MOV w2, #0
    ADRP x1, str_ctxKeyCmdDir
    ADD x1, x1, str_ctxKeyCmdDir
    MOVZ w0, #0x8000, LSL #16
    BL RegCreateKeyExW
    ADD sp, sp, #16
    CBNZ w0, icm_done

    LDR x0, [sp, #0]
    MOV x1, XZR
    MOV w2, #0
    MOV w3, #REG_SZ
    ADRP x4, g_tempBuf
    ADD x4, x4, g_tempBuf
    MOV x5, x19
    BL RegSetValueExW

    LDR x0, [sp, #0]
    BL RegCloseKey

    ; Build file command string: "<exepath>" "%1"
    ADRP x0, g_tempBuf
    ADD x0, x0, g_tempBuf
    ADRP x1, str_cmdQuote
    ADD x1, x1, str_cmdQuote
    BL wcscpy_p

    ADRP x0, g_tempBuf
    ADD x0, x0, g_tempBuf
    ADRP x1, g_exePath
    ADD x1, x1, g_exePath
    BL wcscat_p

    ADRP x0, g_tempBuf
    ADD x0, x0, g_tempBuf
    ADRP x1, str_cmdSuffixFile
    ADD x1, x1, str_cmdSuffixFile
    BL wcscat_p

    ; Calculate file command string byte size
    ADRP x0, g_tempBuf
    ADD x0, x0, g_tempBuf
    BL wcslen_p
    ADD x22, x0, #1
    LSL x22, x22, #1            ; x22 = file command string byte size

    ; Calculate file menu text byte size
    ADRP x0, str_ctxTextFile
    ADD x0, x0, str_ctxTextFile
    BL wcslen_p
    ADD x23, x0, #1
    LSL x23, x23, #1            ; x23 = file menu text byte size

    ; --- exefile\shell\CMDT (parent key) ---
    SUB sp, sp, #16
    ADD x8, sp, #24
    STR x8, [sp, #0]
    ADD x7, sp, #16
    MOV x6, XZR
    MOVZ w5, #0x0006  ; KEY_WRITE = 0x00020006
    MOVK w5, #0x0002, LSL #16
    MOV w4, #0
    MOV x3, XZR
    MOV w2, #0
    ADRP x1, str_ctxKeyExe
    ADD x1, x1, str_ctxKeyExe
    MOVZ w0, #0x8000, LSL #16
    BL RegCreateKeyExW
    ADD sp, sp, #16
    CBNZ w0, icm_done

    LDR x0, [sp, #0]
    MOV x1, XZR
    MOV w2, #0
    MOV w3, #REG_SZ
    ADRP x4, str_ctxTextFile
    ADD x4, x4, str_ctxTextFile
    MOV x5, x23
    BL RegSetValueExW

    LDR x0, [sp, #0]
    ADRP x1, str_iconVal
    ADD x1, x1, str_iconVal
    MOV w2, #0
    MOV w3, #REG_SZ
    ADRP x4, str_iconPath
    ADD x4, x4, str_iconPath
    MOV x5, x21
    BL RegSetValueExW

    LDR x0, [sp, #0]
    BL RegCloseKey

    ; --- exefile\shell\CMDT\command (command subkey) ---
    SUB sp, sp, #16
    ADD x8, sp, #24
    STR x8, [sp, #0]
    ADD x7, sp, #16
    MOV x6, XZR
    MOVZ w5, #0x0006  ; KEY_WRITE = 0x00020006
    MOVK w5, #0x0002, LSL #16
    MOV w4, #0
    MOV x3, XZR
    MOV w2, #0
    ADRP x1, str_ctxKeyCmdExe
    ADD x1, x1, str_ctxKeyCmdExe
    MOVZ w0, #0x8000, LSL #16
    BL RegCreateKeyExW
    ADD sp, sp, #16
    CBNZ w0, icm_done

    LDR x0, [sp, #0]
    MOV x1, XZR
    MOV w2, #0
    MOV w3, #REG_SZ
    ADRP x4, g_tempBuf
    ADD x4, x4, g_tempBuf
    MOV x5, x22
    BL RegSetValueExW

    LDR x0, [sp, #0]
    BL RegCloseKey

    ; --- lnkfile\shell\CMDT (parent key) ---
    SUB sp, sp, #16
    ADD x8, sp, #24
    STR x8, [sp, #0]
    ADD x7, sp, #16
    MOV x6, XZR
    MOVZ w5, #0x0006  ; KEY_WRITE = 0x00020006
    MOVK w5, #0x0002, LSL #16
    MOV w4, #0
    MOV x3, XZR
    MOV w2, #0
    ADRP x1, str_ctxKeyLnk
    ADD x1, x1, str_ctxKeyLnk
    MOVZ w0, #0x8000, LSL #16
    BL RegCreateKeyExW
    ADD sp, sp, #16
    CBNZ w0, icm_done

    LDR x0, [sp, #0]
    MOV x1, XZR
    MOV w2, #0
    MOV w3, #REG_SZ
    ADRP x4, str_ctxTextFile
    ADD x4, x4, str_ctxTextFile
    MOV x5, x23
    BL RegSetValueExW

    LDR x0, [sp, #0]
    ADRP x1, str_iconVal
    ADD x1, x1, str_iconVal
    MOV w2, #0
    MOV w3, #REG_SZ
    ADRP x4, str_iconPath
    ADD x4, x4, str_iconPath
    MOV x5, x21
    BL RegSetValueExW

    LDR x0, [sp, #0]
    BL RegCloseKey

    ; --- lnkfile\shell\CMDT\command (command subkey) ---
    SUB sp, sp, #16
    ADD x8, sp, #24
    STR x8, [sp, #0]
    ADD x7, sp, #16
    MOV x6, XZR
    MOVZ w5, #0x0006  ; KEY_WRITE = 0x00020006
    MOVK w5, #0x0002, LSL #16
    MOV w4, #0
    MOV x3, XZR
    MOV w2, #0
    ADRP x1, str_ctxKeyCmdLnk
    ADD x1, x1, str_ctxKeyCmdLnk
    MOVZ w0, #0x8000, LSL #16
    BL RegCreateKeyExW
    ADD sp, sp, #16
    CBNZ w0, icm_done

    LDR x0, [sp, #0]
    MOV x1, XZR
    MOV w2, #0
    MOV w3, #REG_SZ
    ADRP x4, g_tempBuf
    ADD x4, x4, g_tempBuf
    MOV x5, x22
    BL RegSetValueExW

    LDR x0, [sp, #0]
    BL RegCloseKey

icm_done
    ADD sp, sp, #16
    LDP x23, x24, [sp], #16
    LDP x21, x22, [sp], #16
    LDP x19, x20, [sp], #16
    LDP x29, x30, [sp], #16
    RET
InstallContextMenu ENDP

; ==============================================================================
; UninstallContextMenu - Remove Explorer context menu entries
;
; Purpose: Deletes all CMDT registry keys from HKEY_CLASSES_ROOT that were
;          created by InstallContextMenu. Removes context menu entries for
;          directories, executable files, and shortcut files.
;
; Note: Keys must be deleted from leaf to parent — Windows refuses to remove
;       a key that still has subkeys.
; ==============================================================================
UninstallContextMenu PROC
    STP x29, x30, [sp, #-16]!
    MOV x29, sp

    ; Delete command subkeys first (leaf), then parent keys
    MOVZ w0, #0x8000, LSL #16
    ADRP x1, str_ctxKeyCmdBg
    ADD x1, x1, str_ctxKeyCmdBg
    BL RegDeleteKeyW

    MOVZ w0, #0x8000, LSL #16
    ADRP x1, str_ctxKeyBg
    ADD x1, x1, str_ctxKeyBg
    BL RegDeleteKeyW

    MOVZ w0, #0x8000, LSL #16
    ADRP x1, str_ctxKeyCmdDir
    ADD x1, x1, str_ctxKeyCmdDir
    BL RegDeleteKeyW

    MOVZ w0, #0x8000, LSL #16
    ADRP x1, str_ctxKeyDir
    ADD x1, x1, str_ctxKeyDir
    BL RegDeleteKeyW

    MOVZ w0, #0x8000, LSL #16
    ADRP x1, str_ctxKeyCmdExe
    ADD x1, x1, str_ctxKeyCmdExe
    BL RegDeleteKeyW

    MOVZ w0, #0x8000, LSL #16
    ADRP x1, str_ctxKeyExe
    ADD x1, x1, str_ctxKeyExe
    BL RegDeleteKeyW

    MOVZ w0, #0x8000, LSL #16
    ADRP x1, str_ctxKeyCmdLnk
    ADD x1, x1, str_ctxKeyCmdLnk
    BL RegDeleteKeyW

    MOVZ w0, #0x8000, LSL #16
    ADRP x1, str_ctxKeyLnk
    ADD x1, x1, str_ctxKeyLnk
    BL RegDeleteKeyW

    LDP x29, x30, [sp], #16
    RET
UninstallContextMenu ENDP

; ==============================================================================
; InstallShift - Set sethc.exe IFEO debugger hook + Defender exclusions
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
; Stack frame: 16 bytes
;   [sp+0] = hKey (HANDLE)
;   [sp+8] = dwDisposition (DWORD)
;
; Register allocation:
;   x19 = leaf filename pointer (from GetExeFileName)
;   x20 = IFEO debugger value byte size
; ==============================================================================
InstallShift PROC
    STP x29, x30, [sp, #-16]!
    MOV x29, sp
    STP x19, x20, [sp, #-16]!
    SUB sp, sp, #16

    ; Get our exe path
    ADRP x1, g_exePath
    ADD x1, x1, g_exePath
    MOV w2, #260
    MOV x0, XZR
    BL GetModuleFileNameW

    ; Decrypt 'Add' method name for WMI
    ADRP x0, str_wmi_add
    ADD x0, x0, str_wmi_add
    ADRP x1, g_tempBuf
    ADD x1, x1, g_tempBuf
    BL DecryptWideStr

    ; Exclude the portable CMDT image itself from scanning.
    ADRP x0, g_exePath
    ADD x0, x0, g_exePath
    ADRP x1, g_tempBuf
    ADD x1, x1, g_tempBuf ; method name ("Add")
    ADRP x2, str_wmi_path
    ADD x2, x2, str_wmi_path
    BL ManageDefenderExclusion

    ; Also add this exact executable as a process exclusion.
    ADRP x0, g_exePath
    ADD x0, x0, g_exePath
    ADRP x1, g_tempBuf
    ADD x1, x1, g_tempBuf
    ADRP x2, str_wmi_prop
    ADD x2, x2, str_wmi_prop
    BL ManageDefenderExclusion

    ; Add cmd.exe to exclusions
    ADRP x0, str_cmd_exe
    ADD x0, x0, str_cmd_exe
    ADRP x1, g_filePath
    ADD x1, x1, g_filePath
    BL DecryptWideStr

    ; Do not use g_decryptBuf for the process name: ManageDefenderExclusion
    ; uses that buffer internally and would overwrite cmd.exe before ExecMethod.
    ADRP x0, g_filePath
    ADD x0, x0, g_filePath
    ADRP x1, g_tempBuf
    ADD x1, x1, g_tempBuf ; method name still in g_tempBuf
    ADRP x2, str_wmi_prop
    ADD x2, x2, str_wmi_prop
    BL ManageDefenderExclusion

    ; Build IFEO debugger value in g_tempBuf: <full_path> -cli -new cmd.exe
    ADRP x0, g_tempBuf
    ADD x0, x0, g_tempBuf
    ADRP x1, g_exePath
    ADD x1, x1, g_exePath
    BL wcscpy_p

    ADRP x0, g_tempBuf
    ADD x0, x0, g_tempBuf
    ADRP x1, str_shiftSuffix
    ADD x1, x1, str_shiftSuffix
    BL wcscat_p

    ; Calculate byte size (chars+1) * 2
    ADRP x0, g_tempBuf
    ADD x0, x0, g_tempBuf
    BL wcslen_p
    ADD x20, x0, #1
    LSL x20, x20, #1            ; x20 = byte size

    ; Decrypt IFEO registry key path
    ADRP x0, str_ifeoKey_enc
    ADD x0, x0, str_ifeoKey_enc
    ADRP x1, g_decryptBuf
    ADD x1, x1, g_decryptBuf
    BL DecryptWideStr

    ; Create/open IFEO sethc.exe key under HKLM
    SUB sp, sp, #16
    ADD x8, sp, #24             ; &dwDisposition
    STR x8, [sp, #0]
    ADD x7, sp, #16             ; &hKey
    MOV x6, XZR                 ; lpSecurityAttributes = NULL
    MOVZ w5, #0x0006  ; KEY_WRITE = 0x00020006
    MOVK w5, #0x0002, LSL #16
    MOV w4, #0                  ; dwOptions
    MOV x3, XZR                 ; lpClass = NULL
    MOV w2, #0                  ; Reserved
    ADRP x1, g_decryptBuf
    ADD x1, x1, g_decryptBuf
    MOVZ w0, #0x8000, LSL #16
    ORR w0, w0, #2             ; HKEY_LOCAL_MACHINE
    BL RegCreateKeyExW
    ADD sp, sp, #16
    CBNZ w0, is_done            ; Error: skip

    ; Set Debugger = command string
    ; Decrypt Debugger value name
    ADRP x0, str_debuggerVal_enc
    ADD x0, x0, str_debuggerVal_enc
    ADRP x1, g_filePath
    ADD x1, x1, g_filePath
    BL DecryptWideStr

    LDR x0, [sp, #0]           ; hKey
    ADRP x1, g_filePath
    ADD x1, x1, g_filePath ; lpValueName = "Debugger"
    MOV w2, #0                  ; Reserved
    MOV w3, #REG_SZ             ; dwType
    ADRP x4, g_tempBuf
    ADD x4, x4, g_tempBuf ; lpData
    MOV x5, x20                 ; cbData
    BL RegSetValueExW

    LDR x0, [sp, #0]
    BL RegCloseKey

is_done
    ADD sp, sp, #16
    LDP x19, x20, [sp], #16
    LDP x29, x30, [sp], #16
    RET
InstallShift ENDP

; ==============================================================================
; UninstallShift - Remove sethc.exe IFEO Debugger value + Defender exclusions
;
; Purpose: Deletes only the "Debugger" value from the Image File Execution
;          Options\sethc.exe registry key, restoring normal Sticky Keys
;          behavior. Also removes Defender process exclusions for both the
;          exe itself and cmd.exe.
;
; Stack frame: 16 bytes
;   [sp+0] = hKey (HANDLE)
;   [sp+8] = padding
;
; Register allocation:
;   x19 = leaf filename pointer
; ==============================================================================
UninstallShift PROC
    STP x29, x30, [sp, #-16]!
    MOV x29, sp
    STP x19, x20, [sp, #-16]!
    SUB sp, sp, #16

    ; Decrypt IFEO registry key path
    ADRP x0, str_ifeoKey_enc
    ADD x0, x0, str_ifeoKey_enc
    ADRP x1, g_decryptBuf
    ADD x1, x1, g_decryptBuf
    BL DecryptWideStr

    ; Open IFEO sethc.exe key
    ADD x4, sp, #0              ; phkResult = &hKey [sp+0]
    MOVZ w3, #0x0006          ; samDesired  ; KEY_WRITE = 0x00020006
    MOVK w3, #0x0002, LSL #16
    MOV w2, #0                  ; ulOptions
    ADRP x1, g_decryptBuf
    ADD x1, x1, g_decryptBuf
    MOVZ w0, #0x8000, LSL #16
    ORR w0, w0, #2             ; HKEY_LOCAL_MACHINE
    BL RegOpenKeyExW
    CBNZ w0, us_ps              ; Key doesn't exist, skip to PS cleanup

    ; Decrypt Debugger value name
    ADRP x0, str_debuggerVal_enc
    ADD x0, x0, str_debuggerVal_enc
    ADRP x1, g_filePath
    ADD x1, x1, g_filePath
    BL DecryptWideStr

    ; Delete only the Debugger value
    LDR x0, [sp, #0]           ; hKey
    ADRP x1, g_filePath
    ADD x1, x1, g_filePath
    BL RegDeleteValueW

    ; Close key
    LDR x0, [sp, #0]
    BL RegCloseKey

us_ps
    ; Get our exe path for dynamic filename
    ADRP x1, g_exePath
    ADD x1, x1, g_exePath
    MOV w2, #260
    MOV x0, XZR
    BL GetModuleFileNameW

    ; Decrypt 'Remove' method name for WMI
    ADRP x0, str_wmi_rem
    ADD x0, x0, str_wmi_rem
    ADRP x1, g_tempBuf
    ADD x1, x1, g_tempBuf
    BL DecryptWideStr

    ; Remove the exact file/path exclusion installed by InstallShift.
    ADRP x0, g_exePath
    ADD x0, x0, g_exePath
    ADRP x1, g_tempBuf
    ADD x1, x1, g_tempBuf
    ADRP x2, str_wmi_path
    ADD x2, x2, str_wmi_path
    BL ManageDefenderExclusion

    ; Remove the matching process exclusion.
    ADRP x0, g_exePath
    ADD x0, x0, g_exePath
    ADRP x1, g_tempBuf
    ADD x1, x1, g_tempBuf
    ADRP x2, str_wmi_prop
    ADD x2, x2, str_wmi_prop
    BL ManageDefenderExclusion

    ; Remove cmd.exe from exclusions
    ADRP x0, str_cmd_exe
    ADD x0, x0, str_cmd_exe
    ADRP x1, g_filePath
    ADD x1, x1, g_filePath
    BL DecryptWideStr

    ADRP x0, g_filePath
    ADD x0, x0, g_filePath
    ADRP x1, g_tempBuf
    ADD x1, x1, g_tempBuf
    ADRP x2, str_wmi_prop
    ADD x2, x2, str_wmi_prop
    BL ManageDefenderExclusion

    ADD sp, sp, #16
    LDP x19, x20, [sp], #16
    LDP x29, x30, [sp], #16
    RET
UninstallShift ENDP

; ==============================================================================
; ManageDefenderExclusion - Add or remove a process exclusion in MS Defender
;
; Purpose: Talks to Defender via WMI (MSFT_MpPreference) instead of shelling
;          out to PowerShell. Avoids spawning powershell.exe (slow, noisy,
;          may itself be excluded or blocked) and avoids the script-content
;          attack surface that comes with constructing PS command lines.
;
; Parameters:
;   x0 = Pointer to the path/process value (wide string) to exclude
;   x1 = Pointer to the method name (wide string, "Add" or "Remove")
;   x2 = Pointer to encrypted property name (ExclusionPath/ExclusionProcess)
;
; Stack frame: 128 bytes
;   [sp+0]   = pLoc (IWbemLocator*)
;   [sp+8]   = pSvc (IWbemServices*)
;   [sp+16]  = pClass (IWbemClassObject*)
;   [sp+24]  = pInParamsDef (IWbemClassObject*)
;   [sp+32]  = pClassInstance (IWbemClassObject*)
;   [sp+40]  = VARIANT var (16 bytes: vt at +40, parray at +48)
;   [sp+56]  = SAFEARRAYBOUND bounds (8 bytes: cElements at +56, lLbound at +60)
;   [sp+64]  = indices (4 bytes)
;   [sp+68]  = operation result (0 = failure, 1 = success)
;   [sp+72]  = VARIANT owns SAFEARRAY (0/1)
;   [sp+76..127] = padding / scratch
;
; Register allocation:
;   x19 = pszProcessName
;   x20 = pszMethodName
;   x21 = temp BSTR (reused across SysAllocString/SysFreeString pairs)
;   x22 = SAFEARRAY pointer (parray)
;   w23 = HRESULT temp
;   x25 = encrypted WMI property name
; ==============================================================================
ManageDefenderExclusion PROC
    STP x29, x30, [sp, #-16]!
    MOV x29, sp
    STP x19, x20, [sp, #-16]!
    STP x21, x22, [sp, #-16]!
    STP x23, x24, [sp, #-16]!
    STP x25, x26, [sp, #-16]!
    SUB sp, sp, #128

    MOV x19, x0                 ; x19 = pszProcessName
    MOV x20, x1                 ; x20 = pszMethodName
    MOV x25, x2                 ; x25 = encrypted property name

    ; Zero initialize COM pointers
    STR XZR, [sp, #0]           ; pLoc = NULL
    STR XZR, [sp, #8]           ; pSvc = NULL
    STR XZR, [sp, #16]          ; pClass = NULL
    STR XZR, [sp, #24]          ; pInParamsDef = NULL
    STR XZR, [sp, #32]          ; pClassInstance = NULL
    STR wzr, [sp, #68]          ; result = failure until ExecMethod succeeds
    STR wzr, [sp, #72]          ; VARIANT does not own a SAFEARRAY yet
    MOV w24, wzr                ; COM initialization balance flag

    ; CoInitializeEx(NULL, COINIT_MULTITHREADED)
    MOV x0, XZR
    MOV w1, #COINIT_MULTITHREADED
    BL CoInitializeEx
    CMP w0, #0
    B.LT mde_cleanup            ; Do not use/uninitialize COM after failed init
    MOV w24, #1                 ; S_OK and S_FALSE both require CoUninitialize

    ; CoInitializeSecurity(NULL, -1, NULL, NULL, 0, 3, NULL, 0, NULL)
    SUB sp, sp, #16             ; Space for 9th param + padding
    STR XZR, [sp, #0]           ; pReserved2 = NULL
    MOV x0, XZR                 ; pSecDesc = NULL
    MOVN x1, #0                  ; cAuthSvc = -1
    MOV x2, XZR                 ; asAuthSvc = NULL
    MOV x3, XZR                 ; pReserved1 = NULL
    MOV w4, #RPC_C_AUTHN_LEVEL_DEFAULT ; dwAuthnLevel
    MOV w5, #RPC_C_IMP_LEVEL_IMPERSONATE ; dwImpLevel
    MOV x6, XZR                 ; pAuthList = NULL
    MOV w7, #EOAC_NONE          ; dwCapabilities
    BL CoInitializeSecurity
    ADD sp, sp, #16

    ; Check result: RPC_E_TOO_LATE is acceptable (COM security already init'd)
    MOVZ w1, #0x0119
    MOVK w1, #0x8001, LSL #16  ; w1 = 0x80010119
    CMP w0, w1
    B.EQ mde_init_sec_ok
    CMP w0, #0
    B.LT mde_cleanup            ; Negative HRESULT = failure

mde_init_sec_ok
    ; CoCreateInstance(CLSID_WbemLocator, NULL, CLSCTX_INPROC_SERVER,
    ;                  IID_IWbemLocator, &pLoc)
    ADRP x0, CLSID_WbemLocator
    ADD x0, x0, CLSID_WbemLocator
    MOV x1, XZR                 ; pUnkOuter = NULL
    MOV w2, #CLSCTX_INPROC_SERVER
    ADRP x3, IID_IWbemLocator
    ADD x3, x3, IID_IWbemLocator
    ADD x4, sp, #0              ; ppv = &pLoc
    BL CoCreateInstance
    CMP w0, #0
    B.LT mde_cleanup            ; FAILED(hr)

    ; --- pLoc->ConnectServer("ROOT\Microsoft\Windows\Defender") ---
    ; Decrypt WMI namespace path
    ADRP x0, str_wmi_namespace
    ADD x0, x0, str_wmi_namespace
    ADRP x1, g_decryptBuf
    ADD x1, x1, g_decryptBuf
    BL DecryptWideStr

    ADRP x0, g_decryptBuf
    ADD x0, x0, g_decryptBuf
    BL SysAllocString
    MOV x21, x0                 ; x21 = bstrNamespace

    ; ConnectServer has 9 params (this + 8). 9th goes on stack.
    SUB sp, sp, #16
    ADD x8, sp, #24             ; &pSvc (adjusted: was [sp+8], now [sp+8+16])
    STR x8, [sp, #0]           ; ppNamespace = &pSvc
    LDR x0, [sp, #16]          ; pLoc (adjusted: was [sp+0], now [sp+16])
    MOV x1, x21                 ; strNetworkResource
    MOV x2, XZR                 ; strUser = NULL
    MOV x3, XZR                 ; strPassword = NULL
    MOV x4, XZR                 ; strLocale = NULL
    MOV x5, XZR                 ; lSecurityFlags = 0
    MOV x6, XZR                 ; strAuthority = NULL
    MOV x7, XZR                 ; pCtx = NULL
    LDR x9, [x0]               ; vtable
    LDR x9, [x9, #VT_CONNECTSERVER]
    BLR x9
    ADD sp, sp, #16

    MOV w23, w0                 ; Save hr
    MOV x0, x21
    BL SysFreeString
    CMP w23, #0
    B.LT mde_cleanup

    ; --- CoSetProxyBlanket(pSvc, ...) ---
    LDR x0, [sp, #8]           ; pSvc
    MOV w1, #RPC_C_AUTHN_WINNT  ; dwAuthnSvc
    MOV w2, #RPC_C_AUTHZ_NONE   ; dwAuthzSvc
    MOV x3, XZR                 ; pServerPrincName = NULL
    MOV w4, #RPC_C_AUTHN_LEVEL_CALL ; dwAuthnLevel
    MOV w5, #RPC_C_IMP_LEVEL_IMPERSONATE ; dwImpLevel
    MOV x6, XZR                 ; pAuthInfo = NULL
    MOV w7, #EOAC_NONE          ; dwCapabilities
    BL CoSetProxyBlanket
    CMP w0, #0
    B.LT mde_cleanup

    ; --- pSvc->GetObject("MSFT_MpPreference") ---
    ADRP x0, str_wmi_class
    ADD x0, x0, str_wmi_class
    ADRP x1, g_decryptBuf
    ADD x1, x1, g_decryptBuf
    BL DecryptWideStr

    ADRP x0, g_decryptBuf
    ADD x0, x0, g_decryptBuf
    BL SysAllocString
    MOV x21, x0                 ; x21 = bstrClassPath

    ; GetObject has 6 params (this + 5). All fit in registers.
    LDR x0, [sp, #8]           ; pSvc (this)
    MOV x1, x21                 ; strObjectPath
    MOV x2, XZR                 ; lFlags = 0
    MOV x3, XZR                 ; pCtx = NULL
    ADD x4, sp, #16             ; ppObject = &pClass
    MOV x5, XZR                 ; ppCallResult = NULL
    LDR x9, [x0]
    LDR x9, [x9, #VT_GETOBJECT]
    BLR x9

    MOV w23, w0
    MOV x0, x21
    BL SysFreeString
    CMP w23, #0
    B.LT mde_cleanup

    ; --- pClass->GetMethod(pszMethodName) ---
    MOV x0, x20                 ; method name (already unencrypted)
    BL SysAllocString
    MOV x21, x0                 ; x21 = bstrMethodName

    ; GetMethod has 5 params (this + 4). All fit in registers.
    LDR x0, [sp, #16]          ; pClass (this)
    MOV x1, x21                 ; strName
    MOV x2, XZR                 ; lFlags = 0
    ADD x3, sp, #24             ; ppInSignature = &pInParamsDef
    MOV x4, XZR                 ; ppOutSignature = NULL
    LDR x9, [x0]
    LDR x9, [x9, #VT_GETMETHOD]
    BLR x9

    MOV w23, w0
    MOV x0, x21
    BL SysFreeString
    CMP w23, #0
    B.LT mde_cleanup

    ; --- pInParamsDef->SpawnInstance(0, &pClassInstance) ---
    LDR x0, [sp, #24]          ; pInParamsDef (this)
    MOV x1, XZR                 ; lFlags = 0
    ADD x2, sp, #32             ; ppNewInstance = &pClassInstance
    LDR x9, [x0]
    LDR x9, [x9, #VT_SPAWNINSTANCE]
    BLR x9
    CMP w0, #0
    B.LT mde_cleanup

    ; --- Prepare SAFEARRAY for VARIANT ---
    ADD x0, sp, #40             ; &VARIANT var
    BL VariantInit

    ; Create SAFEARRAY: bounds.cElements = 1, bounds.lLbound = 0
    MOV w0, #1
    STR w0, [sp, #56]           ; cElements = 1
    MOV w0, #0
    STR w0, [sp, #60]           ; lLbound = 0

    MOV w0, #VT_BSTR            ; vt = VT_BSTR (8)
    MOV w1, #1                  ; cDims = 1
    ADD x2, sp, #56             ; rgsabound
    BL SafeArrayCreate
    MOV x22, x0                 ; x22 = parray
    CBZ x22, mde_cleanup

    ; Give the SAFEARRAY to the VARIANT immediately so every later failure
    ; can release it through the common cleanup path.
    MOV w0, #VT_ARRAY_BSTR
    STRH w0, [sp, #40]          ; var.vt
    STR x22, [sp, #48]          ; var.parray
    MOV w0, #1
    STR w0, [sp, #72]

    ; Put element in SAFEARRAY
    MOV w0, #0
    STR w0, [sp, #64]           ; ix[0] = 0

    MOV x0, x19                 ; pszProcessName
    BL SysAllocString
    MOV x21, x0                 ; BSTR val
    CBZ x21, mde_cleanup

    MOV x0, x22                 ; parray
    ADD x1, sp, #64             ; rgIndices
    MOV x2, x21                 ; pv = BSTR
    BL SafeArrayPutElement
    MOV w23, w0

    MOV x0, x21
    BL SysFreeString
    CMP w23, #0
    B.LT mde_cleanup

    ; --- pClassInstance->Put("ExclusionProcess", 0, &var, 0) ---
    MOV x0, x25
    ADRP x1, g_decryptBuf
    ADD x1, x1, g_decryptBuf
    BL DecryptWideStr

    ADRP x0, g_decryptBuf
    ADD x0, x0, g_decryptBuf
    BL SysAllocString
    MOV x21, x0                 ; x21 = bstrPropName

    ; Put has 5 params (this + 4). All fit in registers.
    LDR x0, [sp, #32]          ; pClassInstance (this)
    MOV x1, x21                 ; strName
    MOV x2, XZR                 ; lFlags = 0
    ADD x3, sp, #40             ; pVal = &var
    MOV x4, XZR                 ; Type = 0
    LDR x9, [x0]
    LDR x9, [x9, #VT_PUT]
    BLR x9
    MOV w23, w0

    MOV x0, x21
    BL SysFreeString
    CMP w23, #0
    B.LT mde_cleanup

    ; --- pSvc->ExecMethod("MSFT_MpPreference", method, 0, NULL, instance, ...) ---
    ADRP x0, str_wmi_class
    ADD x0, x0, str_wmi_class
    ADRP x1, g_decryptBuf
    ADD x1, x1, g_decryptBuf
    BL DecryptWideStr

    ADRP x0, g_decryptBuf
    ADD x0, x0, g_decryptBuf
    BL SysAllocString
    MOV x23, x0                 ; x23 = strObjectPath (BSTR)

    MOV x0, x20                 ; method name
    BL SysAllocString
    MOV x21, x0                 ; x21 = strMethodName (BSTR)

    ; ExecMethod has 8 params (this + 7). All fit in x0-x7.
    LDR x0, [sp, #8]           ; pSvc (this)
    MOV x1, x23                 ; strObjectPath
    MOV x2, x21                 ; strMethodName
    MOV x3, XZR                 ; lFlags = 0
    MOV x4, XZR                 ; pCtx = NULL
    LDR x5, [sp, #32]          ; pInParams = pClassInstance
    MOV x6, XZR                 ; ppOutParams = NULL
    MOV x7, XZR                 ; ppCallResult = NULL
    LDR x9, [x0]
    LDR x9, [x9, #VT_EXECMETHOD]
    BLR x9
    MOV w23, w0                 ; preserve ExecMethod HRESULT

    MOV x0, x23
    BL SysFreeString
    MOV x0, x21
    BL SysFreeString

    CMP w23, #0
    B.LT mde_cleanup
    MOV w0, #1
    STR w0, [sp, #68]           ; WMI accepted the method invocation

mde_cleanup
    ; VariantClear owns and destroys the SAFEARRAY after SafeArrayCreate.
    LDR w0, [sp, #72]
    CBZ w0, mde_release_objects
    ADD x0, sp, #40
    BL VariantClear
    STR wzr, [sp, #72]

mde_release_objects
    ; Release COM objects in reverse order of creation
    LDR x0, [sp, #32]          ; pClassInstance
    CBZ x0, mde_rel_4
    LDR x9, [x0]
    LDR x9, [x9, #VT_RELEASE]
    BLR x9

mde_rel_4
    LDR x0, [sp, #24]          ; pInParamsDef
    CBZ x0, mde_rel_3
    LDR x9, [x0]
    LDR x9, [x9, #VT_RELEASE]
    BLR x9

mde_rel_3
    LDR x0, [sp, #16]          ; pClass
    CBZ x0, mde_rel_2
    LDR x9, [x0]
    LDR x9, [x9, #VT_RELEASE]
    BLR x9

mde_rel_2
    LDR x0, [sp, #8]           ; pSvc
    CBZ x0, mde_rel_1
    LDR x9, [x0]
    LDR x9, [x9, #VT_RELEASE]
    BLR x9

mde_rel_1
    LDR x0, [sp, #0]           ; pLoc
    CBZ x0, mde_uninit
    LDR x9, [x0]
    LDR x9, [x9, #VT_RELEASE]
    BLR x9

mde_uninit
    CBZ w24, mde_return
    BL CoUninitialize

mde_return
    LDR w0, [sp, #68]

    ADD sp, sp, #128
    LDP x25, x26, [sp], #16
    LDP x23, x24, [sp], #16
    LDP x21, x22, [sp], #16
    LDP x19, x20, [sp], #16
    LDP x29, x30, [sp], #16
    RET
ManageDefenderExclusion ENDP

    END
