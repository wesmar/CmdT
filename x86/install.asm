; ==============================================================================
; CMDT - Run as TrustedInstaller (x86)
; Installation / Hook Management Module
;
; Author: Marek Wesolowski (wesmar)
; Purpose: Owns every persistent change CMDT can make to the host system —
;          Explorer context-menu entries, the sethc.exe Sticky-Keys IFEO
;          debugger hook used for login-screen rescue access, and the
;          matching Defender process exclusions wrapped around that hook.
;
; Exported routines (stdcall):
;   InstallContextMenu      - Register HKCR\* shell entries for "Run as TI"
;   UninstallContextMenu    - Remove the above entries
;   InstallShift            - Install sethc.exe IFEO Debugger + exclusions
;   UninstallShift          - Remove the IFEO Debugger value + exclusions
;
; Module-private helpers:
;   GetExeFileName  - Locate the filename portion of g_exePath
;   RunPsCommand    - Spawn powershell.exe synchronously with given args,
;                     used here to drive Add/Remove-MpPreference
; ==============================================================================

.586
.model flat, stdcall
option casemap:none

include consts.inc
include globals.inc

; --- Win32 / Reg / Shell dependencies ---
GetModuleFileNameW      PROTO :DWORD,:DWORD,:DWORD
ShellExecuteExW         PROTO :DWORD
WaitForSingleObject     PROTO :DWORD,:DWORD
CloseHandle             PROTO :DWORD
RegCreateKeyExW         PROTO :DWORD,:DWORD,:DWORD,:DWORD,:DWORD,:DWORD,:DWORD,:DWORD,:DWORD
RegSetValueExW          PROTO :DWORD,:DWORD,:DWORD,:DWORD,:DWORD,:DWORD
RegDeleteKeyW           PROTO :DWORD,:DWORD
RegOpenKeyExW           PROTO :DWORD,:DWORD,:DWORD,:DWORD,:DWORD
RegDeleteValueW         PROTO :DWORD,:DWORD
RegCloseKey             PROTO :DWORD

; --- In-project helpers from strutil.asm ---
wcscpy_p                PROTO :DWORD,:DWORD
wcscat_p                PROTO :DWORD,:DWORD
wcslen_p                PROTO :DWORD

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

; --- IFEO registry path / values + suffix appended for the Debugger value ---
str_ifeoKey         dw 'S','O','F','T','W','A','R','E','\'
                    dw 'M','i','c','r','o','s','o','f','t','\'
                    dw 'W','i','n','d','o','w','s',' ','N','T','\'
                    dw 'C','u','r','r','e','n','t','V','e','r','s','i','o','n','\'
                    dw 'I','m','a','g','e',' ','F','i','l','e',' '
                    dw 'E','x','e','c','u','t','i','o','n',' '
                    dw 'O','p','t','i','o','n','s','\'
                    dw 's','e','t','h','c','.','e','x','e',0
str_debuggerVal     dw 'D','e','b','u','g','g','e','r',0
str_shiftSuffix     dw ' ','-','c','l','i',' ','-','n','e','w',' ','c','m','d','.','e','x','e',0

; --- PowerShell wrapper for Defender exclusions ---
str_powershell      dw 'p','o','w','e','r','s','h','e','l','l','.','e','x','e',0
str_psAddPfx        dw '-','N','o','P','r','o','f','i','l','e',' ','-','c',' '
                    dw '"','A','d','d','-','M','p','P','r','e','f','e','r','e','n','c','e',' '
                    dw '-','E','x','c','l','u','s','i','o','n','P','r','o','c','e','s','s',' ',0
str_psRemPfx        dw '-','N','o','P','r','o','f','i','l','e',' ','-','c',' '
                    dw '"','R','e','m','o','v','e','-','M','p','P','r','e','f','e','r','e','n','c','e',' '
                    dw '-','E','x','c','l','u','s','i','o','n','P','r','o','c','e','s','s',' ',0
str_psSuffix        dw ',','c','m','d','.','e','x','e',' ','-','F','o','r','c','e','"',0

; ==============================================================================
; CODE SECTION
; ==============================================================================
.code

; ==============================================================================
; GetExeFileName - Return pointer to filename portion of g_exePath
;
; Walks g_exePath looking for the last backslash and returns a pointer to
; the character following it. The IFEO Debugger value references the
; executable by its short name, so we feed cmd.exe's name (or whatever
; cmdt was renamed to) rather than its full path into the Defender
; exclusion call.
; ==============================================================================
GetExeFileName proc
    mov eax, offset g_exePath
    mov ecx, eax
@@:
    cmp word ptr [ecx], 0
    je @F
    cmp word ptr [ecx], '\'
    jne gef_next
    lea eax, [ecx+2]
gef_next:
    add ecx, 2
    jmp @B
@@:
    ret
GetExeFileName endp

; ==============================================================================
; RunPsCommand - Run a PowerShell one-liner synchronously
;
; Spawns powershell.exe with SW_HIDE via ShellExecuteExW, waits for it to
; finish, and closes the process handle. Used to drive Add-MpPreference /
; Remove-MpPreference when the Sticky-Keys IFEO hook is installed or
; removed; PowerShell is the only documented user-mode way to manage
; Defender exclusions without shipping our own COM/WMI client.
; ==============================================================================
RunPsCommand proc uses edi lpParams:DWORD
    LOCAL sei[60]:BYTE
    LOCAL hProc:DWORD

    ; Zero SHELLEXECUTEINFOW (60 bytes on x86)
    lea edi, sei
    xor eax, eax
    mov ecx, 15
    rep stosd

    lea edi, sei
    mov dword ptr [edi], 60                         ; cbSize
    mov dword ptr [edi+4], 00000040h                ; fMask = SEE_MASK_NOCLOSEPROCESS
    mov dword ptr [edi+16], offset str_powershell   ; lpFile
    mov eax, lpParams
    mov dword ptr [edi+20], eax                     ; lpParameters
    mov dword ptr [edi+28], SW_HIDE                 ; nShow

    invoke ShellExecuteExW, edi
    test eax, eax
    jz ps_done

    lea edi, sei
    mov eax, dword ptr [edi+56]                     ; hProcess
    mov hProc, eax
    test eax, eax
    jz ps_done

    invoke WaitForSingleObject, hProc, 0FFFFFFFFh
    invoke CloseHandle, hProc

ps_done:
    ret
RunPsCommand endp

; ==============================================================================
; InstallContextMenu - Register Explorer context-menu entries under HKCR
;
; Adds the four CMDT entries (Directory\Background, Directory, exefile,
; lnkfile) plus their command subkeys and shared shell32.dll,104 icon. Any
; partial failure short-circuits the rest — running InstallContextMenu twice
; on top of itself is harmless because all writes are idempotent.
; ==============================================================================
InstallContextMenu proc uses ebx esi edi
    LOCAL hKey:DWORD
    LOCAL dwDisp:DWORD

    invoke GetModuleFileNameW, 0, offset g_exePath, 260

    ; Build directory command string: "<exepath>" -cli -new cmd.exe /k cd /d "%V"
    invoke wcscpy_p, offset g_tempBuf, offset str_cmdQuote
    invoke wcscat_p, offset g_tempBuf, offset g_exePath
    invoke wcscat_p, offset g_tempBuf, offset str_cmdSuffixDir

    invoke wcslen_p, offset g_tempBuf
    inc eax
    shl eax, 1
    mov ebx, eax                ; ebx = directory command string byte size

    invoke wcslen_p, offset str_ctxTextDir
    inc eax
    shl eax, 1
    push eax                    ; menu-text byte size kept on stack so all
                                ; four "set menu label" calls can reuse it

    invoke wcslen_p, offset str_iconPath
    inc eax
    shl eax, 1
    mov esi, eax                ; esi = icon path byte size

    ; --- Directory\Background\shell\CMDT ---
    invoke RegCreateKeyExW, HKEY_CLASSES_ROOT, offset str_ctxKeyBg, 0, 0, 0, KEY_WRITE, 0, addr hKey, addr dwDisp
    test eax, eax
    jnz ctx_install_done_pop

    pop edi
    push edi
    invoke RegSetValueExW, hKey, 0, 0, REG_SZ, offset str_ctxTextDir, edi
    invoke RegSetValueExW, hKey, offset str_iconVal, 0, REG_SZ, offset str_iconPath, esi
    invoke RegCloseKey, hKey

    invoke RegCreateKeyExW, HKEY_CLASSES_ROOT, offset str_ctxKeyCmdBg, 0, 0, 0, KEY_WRITE, 0, addr hKey, addr dwDisp
    test eax, eax
    jnz ctx_install_done_pop
    invoke RegSetValueExW, hKey, 0, 0, REG_SZ, offset g_tempBuf, ebx
    invoke RegCloseKey, hKey

    ; --- Directory\shell\CMDT ---
    invoke RegCreateKeyExW, HKEY_CLASSES_ROOT, offset str_ctxKeyDir, 0, 0, 0, KEY_WRITE, 0, addr hKey, addr dwDisp
    test eax, eax
    jnz ctx_install_done_pop

    pop edi
    push edi
    invoke RegSetValueExW, hKey, 0, 0, REG_SZ, offset str_ctxTextDir, edi
    invoke RegSetValueExW, hKey, offset str_iconVal, 0, REG_SZ, offset str_iconPath, esi
    invoke RegCloseKey, hKey

    invoke RegCreateKeyExW, HKEY_CLASSES_ROOT, offset str_ctxKeyCmdDir, 0, 0, 0, KEY_WRITE, 0, addr hKey, addr dwDisp
    test eax, eax
    jnz ctx_install_done_pop
    invoke RegSetValueExW, hKey, 0, 0, REG_SZ, offset g_tempBuf, ebx
    invoke RegCloseKey, hKey

    ; Build file command string: "<exepath>" "%1"
    invoke wcscpy_p, offset g_tempBuf, offset str_cmdQuote
    invoke wcscat_p, offset g_tempBuf, offset g_exePath
    invoke wcscat_p, offset g_tempBuf, offset str_cmdSuffixFile

    invoke wcslen_p, offset g_tempBuf
    inc eax
    shl eax, 1
    mov ebx, eax                ; ebx = file command string byte size

    invoke wcslen_p, offset str_ctxTextFile
    inc eax
    shl eax, 1
    mov edi, eax                ; edi = file menu text byte size

    ; --- exefile\shell\CMDT ---
    invoke RegCreateKeyExW, HKEY_CLASSES_ROOT, offset str_ctxKeyExe, 0, 0, 0, KEY_WRITE, 0, addr hKey, addr dwDisp
    test eax, eax
    jnz ctx_install_done_pop
    invoke RegSetValueExW, hKey, 0, 0, REG_SZ, offset str_ctxTextFile, edi
    invoke RegSetValueExW, hKey, offset str_iconVal, 0, REG_SZ, offset str_iconPath, esi
    invoke RegCloseKey, hKey

    invoke RegCreateKeyExW, HKEY_CLASSES_ROOT, offset str_ctxKeyCmdExe, 0, 0, 0, KEY_WRITE, 0, addr hKey, addr dwDisp
    test eax, eax
    jnz ctx_install_done_pop
    invoke RegSetValueExW, hKey, 0, 0, REG_SZ, offset g_tempBuf, ebx
    invoke RegCloseKey, hKey

    ; --- lnkfile\shell\CMDT ---
    invoke RegCreateKeyExW, HKEY_CLASSES_ROOT, offset str_ctxKeyLnk, 0, 0, 0, KEY_WRITE, 0, addr hKey, addr dwDisp
    test eax, eax
    jnz ctx_install_done_pop
    invoke RegSetValueExW, hKey, 0, 0, REG_SZ, offset str_ctxTextFile, edi
    invoke RegSetValueExW, hKey, offset str_iconVal, 0, REG_SZ, offset str_iconPath, esi
    invoke RegCloseKey, hKey

    invoke RegCreateKeyExW, HKEY_CLASSES_ROOT, offset str_ctxKeyCmdLnk, 0, 0, 0, KEY_WRITE, 0, addr hKey, addr dwDisp
    test eax, eax
    jnz ctx_install_done_pop
    invoke RegSetValueExW, hKey, 0, 0, REG_SZ, offset g_tempBuf, ebx
    invoke RegCloseKey, hKey

ctx_install_done_pop:
    pop eax                     ; discard saved directory menu-text byte size
    ret
InstallContextMenu endp

; ==============================================================================
; UninstallContextMenu - Remove Explorer context-menu entries
;
; Keys are deleted leaf-first because the Windows registry refuses to remove
; a key that still has subkeys.
; ==============================================================================
UninstallContextMenu proc
    invoke RegDeleteKeyW, HKEY_CLASSES_ROOT, offset str_ctxKeyCmdBg
    invoke RegDeleteKeyW, HKEY_CLASSES_ROOT, offset str_ctxKeyBg

    invoke RegDeleteKeyW, HKEY_CLASSES_ROOT, offset str_ctxKeyCmdDir
    invoke RegDeleteKeyW, HKEY_CLASSES_ROOT, offset str_ctxKeyDir

    invoke RegDeleteKeyW, HKEY_CLASSES_ROOT, offset str_ctxKeyCmdExe
    invoke RegDeleteKeyW, HKEY_CLASSES_ROOT, offset str_ctxKeyExe

    invoke RegDeleteKeyW, HKEY_CLASSES_ROOT, offset str_ctxKeyCmdLnk
    invoke RegDeleteKeyW, HKEY_CLASSES_ROOT, offset str_ctxKeyLnk
    ret
UninstallContextMenu endp

; ==============================================================================
; InstallShift - Install sethc.exe IFEO Debugger hook + Defender exclusions
;
; Writes HKLM\...\Image File Execution Options\sethc.exe!Debugger = "<our
; full path> -cli -new cmd.exe", which causes the Sticky-Keys handler to be
; replaced by cmdt at login screen. Without the Defender exclusions added
; here, AV would flag the IFEO write itself; we exclude the cmdt filename
; and cmd.exe via Add-MpPreference.
; ==============================================================================
InstallShift proc uses ebx esi edi
    LOCAL hKey:DWORD
    LOCAL dwDisp:DWORD

    invoke GetModuleFileNameW, 0, offset g_exePath, 260

    call GetExeFileName
    mov esi, eax                ; esi = pointer to exe basename inside g_exePath

    ; Build PS "add exclusion" command and run it.
    invoke wcscpy_p, offset g_cmdBuf, offset str_psAddPfx
    invoke wcscat_p, offset g_cmdBuf, esi
    invoke wcscat_p, offset g_cmdBuf, offset str_psSuffix
    invoke RunPsCommand, offset g_cmdBuf

    ; Build IFEO debugger value: <full_path> -cli -new cmd.exe
    invoke wcscpy_p, offset g_tempBuf, offset g_exePath
    invoke wcscat_p, offset g_tempBuf, offset str_shiftSuffix

    invoke wcslen_p, offset g_tempBuf
    inc eax
    shl eax, 1
    mov ebx, eax

    invoke RegCreateKeyExW, HKEY_LOCAL_MACHINE, offset str_ifeoKey, 0, 0, 0, KEY_WRITE, 0, addr hKey, addr dwDisp
    test eax, eax
    jnz shift_install_done

    invoke RegSetValueExW, hKey, offset str_debuggerVal, 0, REG_SZ, offset g_tempBuf, ebx
    invoke RegCloseKey, hKey

shift_install_done:
    ret
InstallShift endp

; ==============================================================================
; UninstallShift - Remove sethc.exe IFEO Debugger value + Defender exclusions
;
; Deletes only the Debugger value (not the whole IFEO key — other tools may
; rely on it for unrelated settings) then runs Remove-MpPreference to undo
; the exclusions we added in InstallShift.
; ==============================================================================
UninstallShift proc uses esi
    LOCAL hKey:DWORD

    invoke RegOpenKeyExW, HKEY_LOCAL_MACHINE, offset str_ifeoKey, 0, KEY_WRITE, addr hKey
    test eax, eax
    jnz unshift_ps

    invoke RegDeleteValueW, hKey, offset str_debuggerVal
    invoke RegCloseKey, hKey

unshift_ps:
    invoke GetModuleFileNameW, 0, offset g_exePath, 260

    call GetExeFileName
    mov esi, eax

    invoke wcscpy_p, offset g_cmdBuf, offset str_psRemPfx
    invoke wcscat_p, offset g_cmdBuf, esi
    invoke wcscat_p, offset g_cmdBuf, offset str_psSuffix
    invoke RunPsCommand, offset g_cmdBuf
    ret
UninstallShift endp

end
