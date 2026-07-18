; ==============================================================================
; CMDT - Run as TrustedInstaller
; GUI Window and User Interface Module
; 
; Author: Marek Wesołowski (wesmar)
; Purpose: Implements the complete graphical user interface including:
;          - Main window creation and window class registration
;          - Window procedure for message handling (WndProc)
;          - File browsing dialog and drag-and-drop file support
;          - Registry MRU (Most Recently Used) command list management
;          - Windows .lnk (shortcut) file resolution using COM interfaces
;          - Menu system (File menu with Browse, Enable History, About, Exit)
;          - Dynamic window resizing and control repositioning
;          - Integration with RunAsTrustedInstaller for command execution
; ==============================================================================

option casemap:none

include consts.inc
include globals.inc

; ==============================================================================
; EXTERNAL FUNCTION DECLARATIONS
; ==============================================================================

; Application-specific functions
; Application-specific functions
EXTRN RunAsTrustedInstaller:PROC

; Windows User32 API - Window management
EXTRN RegisterClassW:PROC

; Registry API functions for MRU list
; Registry API functions for MRU list
EXTRN RegCreateKeyExW:PROC
EXTRN RegOpenKeyExW:PROC
EXTRN RegSetValueExW:PROC
EXTRN RegEnumValueW:PROC
EXTRN RegQueryValueExW:PROC
EXTRN RegDeleteValueW:PROC
EXTRN RegDeleteTreeW:PROC
EXTRN RegCloseKey:PROC

; Window creation and management functions
EXTRN CreateWindowExW:PROC
EXTRN DefWindowProcW:PROC
EXTRN PostQuitMessage:PROC
EXTRN MessageBoxW:PROC
EXTRN GetWindowTextW:PROC
EXTRN SetWindowTextW:PROC
EXTRN GetClientRect:PROC
EXTRN MoveWindow:PROC
EXTRN ShowWindow:PROC
EXTRN UpdateWindow:PROC
EXTRN LoadIconW:PROC
EXTRN ExtractIconExW:PROC
EXTRN LoadCursorW:PROC

; Menu functions
EXTRN CreateMenu:PROC
EXTRN CreatePopupMenu:PROC
EXTRN AppendMenuW:PROC
EXTRN ModifyMenuW:PROC
EXTRN GetMenu:PROC
EXTRN SetMenuInfo:PROC
EXTRN SetMenu:PROC
EXTRN CheckMenuItem:PROC

; File dialog and file operations
EXTRN GetOpenFileNameW:PROC
EXTRN lstrcpyW:PROC

; Control and focus management
EXTRN SetFocus:PROC
EXTRN GetFocus:PROC
EXTRN SendMessageW:PROC
EXTRN GetDlgItem:PROC

; Drag and drop support
EXTRN DragAcceptFiles:PROC
EXTRN DragQueryFileW:PROC
EXTRN DragFinish:PROC
EXTRN ChangeWindowMessageFilterEx:PROC

; COM functions for .lnk resolution
EXTRN CoInitialize:PROC
EXTRN CoUninitialize:PROC
EXTRN CoCreateInstance:PROC

; Desktop Window Manager (DWM) functions for Windows 11 visual effects
EXTRN DwmSetWindowAttribute:PROC
EXTRN SetWindowTheme:PROC
EXTRN CreateSolidBrush:PROC
EXTRN GetStockObject:PROC
EXTRN DeleteObject:PROC
EXTRN SetTextColor:PROC
EXTRN SetBkColor:PROC
EXTRN SetBkMode:PROC
EXTRN FillRect:PROC
EXTRN DrawTextW:PROC
EXTRN InvalidateRect:PROC
EXTRN TrayHandleWindowMessage:PROC
EXTRN TrayShutdown:PROC
EXTRN DrawMenuBar:PROC

; Per-monitor DPI queries (Windows 10 1607+). The manifest declares
; PerMonitorV2 awareness, so Windows hands this process real physical
; pixels with no automatic bitmap stretching -- every control coordinate
; below is authored for 96 DPI (100%) and must be scaled by hand or the
; window renders postage-stamp small on any scaled display.
;
; These three are NOT statically imported (no EXTRN :PROC here) -- they
; don't exist in user32.dll before Windows 10 1607, and a static import
; table entry that the loader can't resolve makes the whole process fail
; to start on Windows 7/8/8.1 ("The procedure entry point ... could not be
; located"), a real crash report from a Windows 7 user. Resolved instead
; at runtime via GetModuleHandleA+GetProcAddress in InitDpiApis (called
; once from CreateMainWindow); every call site checks the resulting
; function pointer and falls back to fixed 96-DPI behavior if it's NULL.
EXTRN GetModuleHandleA:PROC
EXTRN GetProcAddress:PROC

; ==============================================================================
; CONSTANT STRING DATA
; ==============================================================================
.const

; ANSI (not wide-char) names for the runtime-resolved DPI APIs --
; GetModuleHandleA/GetProcAddress both take LPCSTR, unlike the rest of
; this file's Unicode string constants.
str_user32A          db 'user32.dll',0
str_fnGetDpiForSystem       db 'GetDpiForSystem',0
str_fnGetDpiForWindow       db 'GetDpiForWindow',0
str_fnAdjustWindowRectExForDpi db 'AdjustWindowRectExForDpi',0
str_ntdllA           db 'ntdll.dll',0
str_uxthemeA         db 'uxtheme.dll',0
str_fnRtlGetVersion  db 'RtlGetVersion',0

; Window class names for controls
str_ComboBox    dw 'C','o','m','b','o','B','o','x',0
str_Button      dw 'B','u','t','t','o','n',0
str_Static      dw 'S','t','a','t','i','c',0
str_DarkExplorer dw 'D','a','r','k','M','o','d','e','_','E','x','p','l','o','r','e','r',0
str_Explorer     dw 'E','x','p','l','o','r','e','r',0
str_MenuCheck    dw 2713h,0                  ; Unicode check mark

; Application window class and title
str_ClassName   dw 'T','I','R','u','n','n','e','r','C','l','a','s','s',0
str_Title       dw 'R','u','n',' ','a','s',' ','T','r','u','s','t','e','d','I','n','s','t','a','l','l','e','r',0

; Button labels
str_BtnRun      dw 'R','u','n',0
str_BtnBrowse   dw '&','B','r','o','w','s','e','.','.','.',0

; Status bar messages
str_StatusInit  dw 'R','e','a','d','y','.',' ','E','n','t','e','r',' ','c','o','m','m','a','n','d','.',0
str_StatusRunning dw 'L','a','u','n','c','h','i','n','g','.','.','.',0
str_StatusOK    dw 'P','r','o','c','e','s','s',' ','O','K',0
PUBLIC str_StatusFail, str_ErrEmpty, str_TitleErr
str_StatusFail  dw 'F','a','i','l','e','d',0

; Error messages
str_ErrEmpty    dw 'E','n','t','e','r',' ','c','o','m','m','a','n','d',0
str_TitleErr    dw 'E','r','r','o','r',0

; Menu item text
str_MenuFile    dw '&','F','i','l','e',0
str_MenuBrowse  dw '&','O','p','e','n',' ','w','i','t','h',' ','T','r','u','s','t','e','d','I','n','s','t','a','l','l','e','r',0
str_MenuHistory dw 'E','n','a','b','l','e',' ','H','i','s','t','o','r','y',0
str_MenuExit    dw 'E','&','x','i','t',0
str_MenuAbout   dw '&','A','b','o','u','t',0

; About dialog text (includes author information and all CLI commands)
str_AboutTitle  dw 'A','b','o','u','t',0
str_AboutText   dw 'C','M','D','T',' ',' ','v','1','.','0','.','0','.','7',10,10
                dw 'A','u','t','h','o','r',':',' ',' ','M','a','r','e','k',' ','W','e','s','o',0142h,'o','w','s','k','i',10
                dw 'h','t','t','p','s',':','/','/','k','v','c','.','p','l',10
                dw 'm','a','r','e','k','@','k','v','c','.','p','l',10,10
                dw 'C','L','I',':',10
                dw 'c','m','d','t','.','e','x','e',' ','-','c','l','i',' ','<','c','o','m','m','a','n','d','>',10
                dw 'c','m','d','t','.','e','x','e',' ','-','c','l','i',' ','-','n','e','w',' ','<','c','o','m','m','a','n','d','>',10
                dw 'c','m','d','t','.','e','x','e',' ','-','i','n','s','t','a','l','l',10
                dw 'c','m','d','t','.','e','x','e',' ','-','u','n','i','n','s','t','a','l','l',10
                dw 'c','m','d','t','.','e','x','e',' ','-','s','h','i','f','t',10
                dw 'c','m','d','t','.','e','x','e',' ','-','u','n','s','h','i','f','t',10
                dw 'c','m','d','t','.','e','x','e',' ','-','h','i','s','t','o','r','y','-','c','l','e','a','r',0

; File dialog filter string (double-null terminated)
str_Filter      dw 'E','x','e','c','u','t','a','b','l','e','s',0,'*','.','e','x','e',';','*','.','l','n','k',0,'A','l','l',' ','F','i','l','e','s',0,'*','.','*',0,0

; Default paths and filenames
str_DefPath     dw 'C',':','\',0
str_Shell32     dw 's','h','e','l','l','3','2','.','d','l','l',0

; Registry key for storing MRU list (PUBLIC: also used by main.asm's
; CLI -history-clear switch)
PUBLIC str_regKey
str_regKey      dw 'S','o','f','t','w','a','r','e','\','c','m','d','t',0
; Registry key for app theme preference
str_regThemeKey dw 'S','o','f','t','w','a','r','e','\'
                dw 'M','i','c','r','o','s','o','f','t','\'
                dw 'W','i','n','d','o','w','s','\'
                dw 'C','u','r','r','e','n','t','V','e','r','s','i','o','n','\'
                dw 'T','h','e','m','e','s','\'
                dw 'P','e','r','s','o','n','a','l','i','z','e',0
str_regAppsUseLightTheme dw 'A','p','p','s','U','s','e','L','i','g','h','t','T','h','e','m','e',0

; File extensions
str_extLnk      dw '.','l','n','k',0
str_extExe      dw '.','e','x','e',0
; ==============================================================================
; COM INTERFACE IDENTIFIERS (GUIDs)
; Used for resolving Windows shortcut (.lnk) files
; ==============================================================================

; CLSID_ShellLink: {00021401-0000-0000-C000-000000000046}
; Used to create IShellLink COM object
CLSID_ShellLink dd 00021401h
                dw 0000h, 0000h
                db 0C0h, 00h, 00h, 00h, 00h, 00h, 00h, 46h

; IID_IShellLinkW: {000214F9-0000-0000-C000-000000000046}
; Interface for manipulating Shell Link objects
IID_IShellLinkW dd 000214F9h
                dw 0000h, 0000h
                db 0C0h, 00h, 00h, 00h, 00h, 00h, 00h, 46h

; IID_IPersistFile: {0000010B-0000-0000-C000-000000000046}
; Interface for loading .lnk files
IID_IPersistFile dd 0000010bh
                 dw 0000h, 0000h
                 db 0C0h, 00h, 00h, 00h, 00h, 00h, 00h, 46h

; ==============================================================================
; RUNTIME-RESOLVED DPI API POINTERS
; ==============================================================================
.data?

; Theme resources are private to the GUI module. Colors follow the Windows 11
; Notepad dark palette while remaining plain GDI objects on older systems.
g_darkMode       dd ?
g_nativeDarkSupported dd ?
g_themeApplying  dd ?          ; guards synchronous WM_THEMECHANGED re-entry
g_brushBg        dq ?          ; #202020 main client background
g_brushSurface   dq ?          ; #2B2B2B menu/status surface
g_brushHover     dq ?          ; #383838 selected/hover surface
g_brushInput     dq ?          ; #242424 edit/list background
    align 8
g_pGetDpiForSystem           dq ?
g_pGetDpiForWindow           dq ?
g_pAdjustWindowRectExForDpi  dq ?
g_dpiApisInit                dd ?

; ==============================================================================
; CODE SECTION
; ==============================================================================
.code

; Implementation is split by responsibility while remaining one MASM
; translation unit. This preserves internal labels and exact ABI behavior.
include window_theme.inc
include window_lifecycle.inc
include window_commands.inc
include window_dispatch.inc
include window_events.inc
include window_mru.inc

end
