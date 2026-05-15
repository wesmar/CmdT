; ==============================================================================
; CMDT - Run as TrustedInstaller
; Installation / Hook Management Module
;
; Author: Marek Wesolowski (wesmar)
; Purpose: Owns every persistent change CMDT can make to the host system:
;          Explorer context-menu entries, the sethc.exe Sticky-Keys IFEO hook
;          used for login-screen rescue access, and the matching Defender
;          process exclusions installed alongside the hook.
;
; Exported routines:
;   InstallContextMenu      - Register HKCR\* shell entries for "Run as TI"
;   UninstallContextMenu    - Remove the above entries
;   InstallShift            - Install sethc.exe IFEO Debugger hook + exclusions
;   UninstallShift          - Remove the IFEO Debugger value + exclusions
;   ManageDefenderExclusion - Add/Remove a process from Defender exclusions
;                             via the MSFT_MpPreference WMI class
;   GetExeFileName          - Return pointer to leaf filename of g_exePath
; ==============================================================================

option casemap:none

include consts.inc
include globals.inc

; --- Win32 / COM / WMI dependencies ---
EXTRN GetModuleFileNameW:PROC
EXTRN RegCreateKeyExW:PROC
EXTRN RegSetValueExW:PROC
EXTRN RegDeleteKeyW:PROC
EXTRN RegDeleteValueW:PROC
EXTRN RegOpenKeyExW:PROC
EXTRN RegCloseKey:PROC
EXTRN CoInitializeEx:PROC
EXTRN CoInitializeSecurity:PROC
EXTRN CoCreateInstance:PROC
EXTRN CoSetProxyBlanket:PROC
EXTRN CoUninitialize:PROC
EXTRN SysAllocString:PROC
EXTRN SysFreeString:PROC
EXTRN VariantInit:PROC
EXTRN VariantClear:PROC
EXTRN SafeArrayCreate:PROC
EXTRN SafeArrayPutElement:PROC

; --- In-project helpers ---
EXTRN DecryptWideStr:PROC
EXTRN wcscpy_p:PROC
EXTRN wcscat_p:PROC
EXTRN wcslen_p:PROC

; ==============================================================================
; CONSTANT STRING DATA
; ==============================================================================
.const

; --- Explorer context-menu registry paths (HKCR-relative) ---
str_ctxKeyBg        dw 'D','i','r','e','c','t','o','r','y','\','B','a','c','k','g','r','o','u','n','d','\','s','h','e','l','l','\','C','M','D','T',0
str_ctxKeyCmdBg     dw 'D','i','r','e','c','t','o','r','y','\','B','a','c','k','g','r','o','u','n','d','\','s','h','e','l','l','\','C','M','D','T','\','c','o','m','m','a','n','d',0
str_ctxKeyDir       dw 'D','i','r','e','c','t','o','r','y','\','s','h','e','l','l','\','C','M','D','T',0
str_ctxKeyCmdDir    dw 'D','i','r','e','c','t','o','r','y','\','s','h','e','l','l','\','C','M','D','T','\','c','o','m','m','a','n','d',0
str_ctxKeyExe       dw 'e','x','e','f','i','l','e','\','s','h','e','l','l','\','C','M','D','T',0
str_ctxKeyCmdExe    dw 'e','x','e','f','i','l','e','\','s','h','e','l','l','\','C','M','D','T','\','c','o','m','m','a','n','d',0
str_ctxKeyLnk       dw 'l','n','k','f','i','l','e','\','s','h','e','l','l','\','C','M','D','T',0
str_ctxKeyCmdLnk    dw 'l','n','k','f','i','l','e','\','s','h','e','l','l','\','C','M','D','T','\','c','o','m','m','a','n','d',0

; --- Menu labels shown in Explorer ---
str_ctxTextDir      dw 'O','p','e','n',' ','C','M','D',' ','a','s',' ','T','r','u','s','t','e','d','I','n','s','t','a','l','l','e','r',0
str_ctxTextFile     dw 'R','u','n',' ','a','s',' ','T','r','u','s','t','e','d','I','n','s','t','a','l','l','e','r',0

; --- Icon (UAC shield from shell32.dll resource #104) ---
str_iconVal         dw 'I','c','o','n',0
str_iconPath        dw 's','h','e','l','l','3','2','.','d','l','l',',','1','0','4',0

; --- Command-line templates for the registered shell entries ---
str_cmdQuote        dw '"',0
str_cmdSuffixDir    dw '"',' ','-','c','l','i',' ','-','n','e','w',' ','c','m','d','.','e','x','e',' ','/','k',' ','c','d',' ','/','d',' ','"','%','V','"',0
str_cmdSuffixFile   dw '"',' ','"','%','1','"',0

; --- Suffix appended to exe path for the sethc.exe IFEO Debugger value ---
str_shiftSuffix     dw ' ','-','c','l','i',' ','-','n','e','w',' ','c','m','d','.','e','x','e',0

; --- XOR-obfuscated strings (key = 0xAA). The IFEO and WMI paths are kept
;     encrypted in the binary so that simple `strings` scans don't reveal
;     the hooked location or the Defender-exclusion mechanism. They are
;     decrypted on demand via DecryptWideStr. ---
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

; --- WMI plumbing ---
CLSID_WbemLocator db 011h,0F8h,090h,045h,03Ah,01Dh,0D0h,011h,089h,01Fh,000h,0AAh,000h,04Bh,02Eh,024h
IID_IWbemLocator  db 087h,0A6h,012h,0DCh,07Fh,073h,0CFh,011h,088h,04Dh,000h,0AAh,000h,04Bh,02Eh,024h

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

; ==============================================================================
; CODE SECTION
; ==============================================================================
.code

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
;   RAX = Pointer into g_exePath at the start of the filename component.
;         Points at g_exePath itself if no backslash is found.
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
; Note: Keys must be deleted from leaf to parent — Windows refuses to remove
;       a key that still has subkeys.
; ==============================================================================
UninstallContextMenu proc
    sub rsp, 40                 ; 32 shadow + 8 alignment

    lea rdx, str_ctxKeyCmdBg
    mov ecx, HKEY_CLASSES_ROOT
    call RegDeleteKeyW

    lea rdx, str_ctxKeyBg
    mov ecx, HKEY_CLASSES_ROOT
    call RegDeleteKeyW

    lea rdx, str_ctxKeyCmdDir
    mov ecx, HKEY_CLASSES_ROOT
    call RegDeleteKeyW

    lea rdx, str_ctxKeyDir
    mov ecx, HKEY_CLASSES_ROOT
    call RegDeleteKeyW

    lea rdx, str_ctxKeyCmdExe
    mov ecx, HKEY_CLASSES_ROOT
    call RegDeleteKeyW

    lea rdx, str_ctxKeyExe
    mov ecx, HKEY_CLASSES_ROOT
    call RegDeleteKeyW

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
; ==============================================================================
InstallShift proc
    push rbx
    push r12
    sub rsp, 104

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
    mov rcx, [rsp+72]               ; hKey
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
; UninstallShift - Remove sethc.exe IFEO Debugger value + Defender exclusions
;
; Purpose: Deletes only the "Debugger" value from the Image File Execution
;          Options\sethc.exe registry key, restoring normal Sticky Keys
;          behavior. Also removes Defender process exclusions for both the
;          exe itself and cmd.exe.
; ==============================================================================
UninstallShift proc
    push r12
    sub rsp, 48

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
; ManageDefenderExclusion - Add or remove a process exclusion in MS Defender
;
; Purpose: Talks to Defender via WMI (MSFT_MpPreference) instead of shelling
;          out to PowerShell. Avoids spawning powershell.exe (slow, noisy,
;          may itself be excluded or blocked) and avoids the script-content
;          attack surface that comes with constructing PS command lines.
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
