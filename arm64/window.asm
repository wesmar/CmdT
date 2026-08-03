; ==============================================================================
; CMDT - Run as TrustedInstaller
; GUI Window and User Interface Module
;
; Author: Marek Wesolowski (wesmar)
; Ported to: Native ARM64 (armasm64) for Windows on ARM
;
; Purpose: Implements the complete graphical user interface including:
;          - Main window creation and window class registration
;          - Window procedure for message handling (WndProc)
;          - File browsing dialog and drag-and-drop file support
;          - Registry MRU (Most Recently Used) command list management
;          - Windows .lnk (shortcut) file resolution using COM interfaces
;          - Menu system (File menu with Browse, Enable History, About, Exit)
;          - Dynamic window resizing and control repositioning
;          - Integration with RunAsTrustedInstaller for command execution
;
; ARM64 Port Notes:
;   - This module is a translation unit wrapper; actual implementation is
;     split across six .inc files (theme, lifecycle, commands, dispatch,
;     events, mru) included at the end of the code section.
;   - All EXTRN declarations converted to ARM64 IMPORT.
;   - Runtime-resolved DPI APIs (GetDpiForSystem, GetDpiForWindow,
;     AdjustWindowRectExForDpi) remain dynamically loaded via
;     GetModuleHandleA+GetProcAddress to preserve Windows 7/8 compatibility.
;   - COM GUIDs preserved byte-for-byte (CLSID_ShellLink, IID_IShellLinkW,
;     IID_IPersistFile).
;   - Theme resources (g_darkMode, brushes) stored in BSS with explicit
;     4-byte padding before 8-byte-aligned brush handles.
; ==============================================================================

; ==============================================================================
; EXTERNAL REFERENCES
; ==============================================================================

; Application-specific functions
    IMPORT RunAsTrustedInstaller

; Windows User32 API - Window management
    IMPORT RegisterClassW

; Registry API functions for MRU list
    IMPORT RegCreateKeyExW
    IMPORT RegOpenKeyExW
    IMPORT RegSetValueExW
    IMPORT RegEnumValueW
    IMPORT RegQueryValueExW
    IMPORT RegDeleteValueW
    IMPORT RegDeleteTreeW
    IMPORT RegCloseKey

; Window creation and management functions
    IMPORT CreateWindowExW
    IMPORT DefWindowProcW
    IMPORT PostQuitMessage
    IMPORT MessageBoxW
    IMPORT GetWindowTextW
    IMPORT SetWindowTextW
    IMPORT GetClientRect
    IMPORT MoveWindow
    IMPORT ShowWindow
    IMPORT UpdateWindow
    IMPORT LoadIconW
    IMPORT ExtractIconExW
    IMPORT LoadCursorW

; Menu functions
    IMPORT CreateMenu
    IMPORT CreatePopupMenu
    IMPORT AppendMenuW
    IMPORT ModifyMenuW
    IMPORT GetMenu
    IMPORT SetMenuInfo
    IMPORT SetMenu
    IMPORT CheckMenuItem

; File dialog and file operations
    IMPORT GetOpenFileNameW
    IMPORT lstrcpyW

; Control and focus management
    IMPORT SetFocus
    IMPORT GetFocus
    IMPORT SendMessageW
    IMPORT GetDlgItem

; Drag and drop support
    IMPORT DragAcceptFiles
    IMPORT DragQueryFileW
    IMPORT DragFinish
    IMPORT ChangeWindowMessageFilterEx

; COM functions for .lnk resolution
    IMPORT CoInitialize
    IMPORT CoUninitialize
    IMPORT CoCreateInstance

; Desktop Window Manager (DWM) functions for Windows 11 visual effects
    IMPORT DwmSetWindowAttribute
    IMPORT SetWindowTheme
    IMPORT CreateSolidBrush
    IMPORT GetStockObject
    IMPORT DeleteObject
    IMPORT SetTextColor
    IMPORT SetBkColor
    IMPORT SetBkMode
    IMPORT FillRect
    IMPORT DrawTextW
    IMPORT InvalidateRect
    IMPORT TrayHandleWindowMessage
    IMPORT TrayShutdown
    IMPORT DrawMenuBar

; Per-monitor DPI queries (Windows 10 1607+). The manifest declares
; PerMonitorV2 awareness, so Windows hands this process real physical
; pixels with no automatic bitmap stretching -- every control coordinate
; below is authored for 96 DPI (100%) and must be scaled by hand or the
; window renders postage-stamp small on any scaled display.
;
; These three are NOT statically imported (no IMPORT here) -- they
; don't exist in user32.dll before Windows 10 1607, and a static import
; table entry that the loader can't resolve makes the whole process fail
; to start on Windows 7/8/8.1 ("The procedure entry point ... could not be
; located"), a real crash report from a Windows 7 user. Resolved instead
; at runtime via GetModuleHandleA+GetProcAddress in InitDpiApis (called
; once from CreateMainWindow); every call site checks the resulting
; function pointer and falls back to fixed 96-DPI behavior if it's NULL.
    IMPORT GetModuleHandleA
    IMPORT GetProcAddress

; ==============================================================================
; CONSTANT STRING DATA (READ-ONLY)
; ==============================================================================
    AREA |.const|, DATA, READONLY, ALIGN=3

; Exported symbols used by other modules
    EXPORT str_StatusFail
    EXPORT str_ErrEmpty
    EXPORT str_TitleErr
    EXPORT str_regKey

; ANSI (not wide-char) names for the runtime-resolved DPI APIs --
; GetModuleHandleA/GetProcAddress both take LPCSTR, unlike the rest of
; this file's Unicode string constants.
str_user32A                 DCB "user32.dll", 0
str_fnGetDpiForSystem       DCB "GetDpiForSystem", 0
str_fnGetDpiForWindow       DCB "GetDpiForWindow", 0
str_fnAdjustWindowRectExForDpi DCB "AdjustWindowRectExForDpi", 0
str_ntdllA                  DCB "ntdll.dll", 0
str_uxthemeA                DCB "uxtheme.dll", 0
str_fnRtlGetVersion         DCB "RtlGetVersion", 0

; Window class names for controls
str_ComboBox    DCW 'C','o','m','b','o','B','o','x', 0
str_Button      DCW 'B','u','t','t','o','n', 0
str_Static      DCW 'S','t','a','t','i','c', 0
str_DarkExplorer DCW 'D','a','r','k','M','o','d','e','_','E','x','p','l','o','r','e','r', 0
str_Explorer    DCW 'E','x','p','l','o','r','e','r', 0
str_MenuCheck   DCW 0x2713, 0               ; Unicode check mark ✓

; Application window class and title
str_ClassName   DCW 'T','I','R','u','n','n','e','r','C','l','a','s','s', 0
str_Title       DCW 'R','u','n',' ','a','s',' ','T','r','u','s','t','e','d'
                DCW 'I','n','s','t','a','l','l','e','r', 0

; Button labels
str_BtnRun      DCW 'R','u','n', 0
str_BtnBrowse   DCW '&','B','r','o','w','s','e','.','.','.', 0

; Status bar messages
str_StatusInit  DCW 'R','e','a','d','y','.',' ','E','n','t','e','r',' '
                DCW 'c','o','m','m','a','n','d','.', 0
str_StatusRunning DCW 'L','a','u','n','c','h','i','n','g','.','.','.', 0
str_StatusOK    DCW 'P','r','o','c','e','s','s',' ','O','K', 0
str_StatusFail  DCW 'F','a','i','l','e','d', 0

; Error messages
str_ErrEmpty    DCW 'E','n','t','e','r',' ','c','o','m','m','a','n','d', 0
str_TitleErr    DCW 'E','r','r','o','r', 0

; Menu item text
str_MenuFile    DCW '&','F','i','l','e', 0
str_MenuBrowse  DCW '&','O','p','e','n',' ','w','i','t','h',' '
                DCW 'T','r','u','s','t','e','d','I','n','s','t','a','l','l','e','r', 0
str_MenuHistory DCW 'E','n','a','b','l','e',' ','H','i','s','t','o','r','y', 0
str_MenuExit    DCW 'E','&','x','i','t', 0
str_MenuAbout   DCW '&','A','b','o','u','t', 0

; About dialog text (includes author information and all CLI commands)
str_AboutTitle  DCW 'A','b','o','u','t', 0
str_AboutText   DCW 'C','M','D','T',' ',' ','v','1','.','0','.','0','.','8', 10, 10
                DCW 'A','u','t','h','o','r',':',' ',' ','M','a','r','e','k',' '
                DCW 'W','e','s','o', 0x0142, 'o','w','s','k','i', 10
                DCW 'h','t','t','p','s',':','/','/','k','v','c','.','p','l', 10
                DCW 'm','a','r','e','k','@','k','v','c','.','p','l', 10, 10
                DCW 'C','L','I',':', 10
                DCW 'c','m','d','t','.','e','x','e',' ','-','c','l','i',' '
                DCW '<','c','o','m','m','a','n','d','>', 10
                DCW 'c','m','d','t','.','e','x','e',' ','-','c','l','i',' '
                DCW '-','n','e','w',' ','<','c','o','m','m','a','n','d','>', 10
                DCW 'c','m','d','t','.','e','x','e',' ','-','i','n','s','t','a','l','l', 10
                DCW 'c','m','d','t','.','e','x','e',' ','-','u','n','i','n','s','t','a','l','l', 10
                DCW 'c','m','d','t','.','e','x','e',' ','-','s','h','i','f','t', 10
                DCW 'c','m','d','t','.','e','x','e',' ','-','u','n','s','h','i','f','t', 10
                DCW 'c','m','d','t','.','e','x','e',' ','-','h','i','s','t','o','r','y'
                DCW '-','c','l','e','a','r', 0

; File dialog filter string (double-null terminated)
str_Filter      DCW 'E','x','e','c','u','t','a','b','l','e','s', 0
                DCW '*','.','e','x','e',';','*','.','l','n','k', 0
                DCW 'A','l','l',' ','F','i','l','e','s', 0
                DCW '*','.','*', 0, 0

; Default paths and filenames
str_DefPath     DCW 'C',':',0x5C, 0
str_Shell32     DCW 's','h','e','l','l','3','2','.','d','l','l', 0

; Registry key for storing MRU list (EXPORT: also used by main.asm's
; CLI -history-clear switch)
str_regKey      DCW 'S','o','f','t','w','a','r','e',0x5C,'c','m','d','t', 0

; Registry key for app theme preference
str_regThemeKey DCW 'S','o','f','t','w','a','r','e',0x5C
                DCW 'M','i','c','r','o','s','o','f','t',0x5C
                DCW 'W','i','n','d','o','w','s',0x5C
                DCW 'C','u','r','r','e','n','t','V','e','r','s','i','o','n',0x5C
                DCW 'T','h','e','m','e','s',0x5C
                DCW 'P','e','r','s','o','n','a','l','i','z','e', 0
str_regAppsUseLightTheme DCW 'A','p','p','s','U','s','e','L','i','g','h','t'
                DCW 'T','h','e','m','e', 0

; File extensions
str_extLnk      DCW '.','l','n','k', 0
str_extExe      DCW '.','e','x','e', 0

; ==============================================================================
; COM INTERFACE IDENTIFIERS (GUIDs)
; Used for resolving Windows shortcut (.lnk) files
; ==============================================================================

; GUIDs are emitted as raw contiguous bytes (Data1/Data2/Data3 little-endian,
; Data4 big-endian) using DCB only. Using DCD/DCW here is a trap: armasm
; re-aligns the DCD to a 4-byte boundary and, when the label is not already
; 4-aligned, inserts padding BETWEEN the label and the data -- so the pointer
; we pass to COM points at padding + a shifted GUID and CoCreateInstance
; fails with REGDB_E_CLASSNOTREG. ALIGN 8 keeps each GUID naturally aligned
; and DCB guarantees the label sits exactly on the first byte.

; CLSID_ShellLink: {00021401-0000-0000-C000-000000000046}
    ALIGN 8
CLSID_ShellLink
    DCB 0x01, 0x14, 0x02, 0x00, 0x00, 0x00, 0x00, 0x00
    DCB 0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46

; IID_IShellLinkW: {000214F9-0000-0000-C000-000000000046}
    ALIGN 8
IID_IShellLinkW
    DCB 0xF9, 0x14, 0x02, 0x00, 0x00, 0x00, 0x00, 0x00
    DCB 0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46

; IID_IPersistFile: {0000010B-0000-0000-C000-000000000046}
    ALIGN 8
IID_IPersistFile
    DCB 0x0B, 0x01, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00
    DCB 0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46

; ==============================================================================
; UNINITIALIZED DATA SECTION (BSS)
; ==============================================================================
    AREA |.bss|, DATA, READWRITE, NOINIT, ALIGN=3

; Theme resources are private to the GUI module. Colors follow the Windows 11
; Notepad dark palette while remaining plain GDI objects on older systems.
g_darkMode              DCD 0
g_nativeDarkSupported   DCD 0
g_themeApplying         DCD 0       ; guards synchronous WM_THEMECHANGED re-entry
                        SPACE 4     ; padding for 8-byte alignment of brushes
g_brushBg               DCQ 0       ; #202020 main client background
g_brushSurface          DCQ 0       ; #2B2B2B menu/status surface
g_brushHover            DCQ 0       ; #383838 selected/hover surface
g_brushInput            DCQ 0       ; #242424 edit/list background

; Runtime-resolved DPI API function pointers. NULL on Windows < 10 1607;
; every call site checks for NULL and falls back to fixed 96-DPI behavior.
g_pGetDpiForSystem              DCQ 0
g_pGetDpiForWindow              DCQ 0
g_pAdjustWindowRectExForDpi     DCQ 0
g_dpiApisInit                   DCD 0

; ==============================================================================
; CODE SECTION
; ==============================================================================
    AREA |.text|, CODE, READONLY, ALIGN=4

; Implementation is split by responsibility while remaining one ARM64
; translation unit. This preserves internal labels and exact ABI behavior.
; Each .inc file is included here and contributes code to this .text section.



    INCLUDE window_theme.inc
    INCLUDE window_lifecycle.inc
    INCLUDE window_commands.inc
    INCLUDE window_dispatch.inc
    INCLUDE window_events.inc
    INCLUDE window_mru.inc

    END
