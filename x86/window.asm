; ==============================================================================
; CMDT - Run as TrustedInstaller
; GUI Window and User Interface Module (32-bit version)
; 
; Author: Marek Wesołowski (wesmar)
; Purpose: Implements the complete graphical user interface for the 32-bit
;          version of the application. Includes window creation, message
;          handling, file browsing, drag-and-drop support, registry-based
;          MRU (Most Recently Used) list management, and Windows shortcut
;          (.lnk) file resolution using COM interfaces.
;
; Features:
;          - Main window with ComboBox, buttons, and status label
;          - File browsing dialog with .lnk and .exe filtering
;          - Drag-and-drop file support with UIPI bypass
;          - MRU command list (5 most recent commands)
;          - Registry persistence of MRU list
;          - Windows shortcut (.lnk) resolution via COM
;          - Menu system (File menu with Browse, Enable History, About, Exit)
;          - Dynamic window resizing and control repositioning
;          - ESC key to exit application
; ==============================================================================

.586                            ; Target 80586 instruction set
.model flat, stdcall            ; 32-bit flat memory model, stdcall convention
option casemap:none             ; Case-sensitive symbol names

include consts.inc              ; Windows API constants
include globals.inc             ; Global variable declarations

; ==============================================================================
; EXTERNAL FUNCTION PROTOTYPES
; ==============================================================================

; Application-specific functions
RunAsTrustedInstaller   PROTO :DWORD,:DWORD
ResolveLnkPath          PROTO :DWORD,:DWORD,:DWORD
FixRegeditPath          PROTO :DWORD
ApplyWindowTheme        PROTO :DWORD
ApplyDarkMenuBar        PROTO :DWORD

; Windows User32 API - Window management
RegisterClassW          PROTO :DWORD
CreateWindowExW         PROTO :DWORD,:DWORD,:DWORD,:DWORD,:DWORD,:DWORD,:DWORD,:DWORD,:DWORD,:DWORD,:DWORD,:DWORD
DefWindowProcW          PROTO :DWORD,:DWORD,:DWORD,:DWORD
PostQuitMessage         PROTO :DWORD
MessageBoxW             PROTO :DWORD,:DWORD,:DWORD,:DWORD
GetWindowTextW          PROTO :DWORD,:DWORD,:DWORD
SetWindowTextW          PROTO :DWORD,:DWORD
GetClientRect           PROTO :DWORD,:DWORD
MoveWindow              PROTO :DWORD,:DWORD,:DWORD,:DWORD,:DWORD,:DWORD
ShowWindow              PROTO :DWORD,:DWORD
UpdateWindow            PROTO :DWORD
LoadIconW               PROTO :DWORD,:DWORD
ExtractIconExW          PROTO :DWORD,:DWORD,:DWORD,:DWORD,:DWORD
LoadCursorW             PROTO :DWORD,:DWORD
GetStockObject          PROTO :DWORD
DwmSetWindowAttribute   PROTO :DWORD,:DWORD,:DWORD,:DWORD
SetWindowTheme          PROTO :DWORD,:DWORD,:DWORD
CreateSolidBrush        PROTO :DWORD
DeleteObject            PROTO :DWORD
SetTextColor            PROTO :DWORD,:DWORD
SetBkColor              PROTO :DWORD,:DWORD
SetBkMode               PROTO :DWORD,:DWORD
FillRect                PROTO :DWORD,:DWORD,:DWORD
InvalidateRect          PROTO :DWORD,:DWORD,:DWORD
DrawMenuBar             PROTO :DWORD
DrawTextW               PROTO :DWORD,:DWORD,:DWORD,:DWORD,:DWORD

; Per-monitor DPI queries (Windows 10 1607+). The manifest declares
; PerMonitorV2 awareness, so Windows hands this process real physical
; pixels with no automatic bitmap stretching -- every control coordinate
; below is authored for 96 DPI (100%) and must be scaled by hand or the
; window renders postage-stamp small on any scaled display.
;
; These are NOT declared as PROTO-and-invoke externals -- they don't exist
; in user32.dll before Windows 10 1607, and a static import table entry
; the loader can't resolve makes the whole process fail to start on
; Windows 7/8/8.1 ("The procedure entry point ... could not be located"),
; a real crash report from a Windows 7 user. Resolved instead at runtime
; via GetModuleHandleA+GetProcAddress in InitDpiApis (called once from
; CreateMainWindow); every call site checks the resulting pointer and
; falls back to fixed 96-DPI behavior if it's NULL.
GetModuleHandleA        PROTO :DWORD
GetProcAddress          PROTO :DWORD,:DWORD

; Menu functions
CreateMenu              PROTO
CreatePopupMenu         PROTO
AppendMenuW             PROTO :DWORD,:DWORD,:DWORD,:DWORD
SetMenu                 PROTO :DWORD,:DWORD
CheckMenuItem           PROTO :DWORD,:DWORD,:DWORD
GetMenu                 PROTO :DWORD
SetMenuInfo             PROTO :DWORD,:DWORD
ModifyMenuW             PROTO :DWORD,:DWORD,:DWORD,:DWORD,:DWORD

; File dialog and file operations
GetOpenFileNameW        PROTO :DWORD
lstrcpyW                PROTO :DWORD,:DWORD

; Control and focus management
SetFocus                PROTO :DWORD
GetFocus                PROTO
SendMessageW            PROTO :DWORD,:DWORD,:DWORD,:DWORD
GetDlgItem              PROTO :DWORD,:DWORD

; Drag and drop support
DragAcceptFiles         PROTO :DWORD,:DWORD
DragQueryFileW          PROTO :DWORD,:DWORD,:DWORD,:DWORD
DragFinish              PROTO :DWORD
ChangeWindowMessageFilterEx PROTO :DWORD,:DWORD,:DWORD,:DWORD

; COM functions for .lnk resolution
CoInitialize            PROTO :DWORD
CoUninitialize          PROTO
CoCreateInstance        PROTO :DWORD,:DWORD,:DWORD,:DWORD,:DWORD

; Registry API functions for MRU list
RegCreateKeyExW         PROTO :DWORD,:DWORD,:DWORD,:DWORD,:DWORD,:DWORD,:DWORD,:DWORD,:DWORD
RegOpenKeyExW           PROTO :DWORD,:DWORD,:DWORD,:DWORD,:DWORD
RegSetValueExW          PROTO :DWORD,:DWORD,:DWORD,:DWORD,:DWORD,:DWORD
RegEnumValueW           PROTO :DWORD,:DWORD,:DWORD,:DWORD,:DWORD,:DWORD,:DWORD,:DWORD
RegQueryValueExW        PROTO :DWORD,:DWORD,:DWORD,:DWORD,:DWORD,:DWORD
RegDeleteValueW         PROTO :DWORD,:DWORD
RegDeleteTreeW          PROTO :DWORD,:DWORD
RegCloseKey             PROTO :DWORD

; String utility prototypes (local implementations)
wcslen_w                PROTO :DWORD
wcscmp_ci_w             PROTO :DWORD,:DWORD
wcscat_w                PROTO :DWORD,:DWORD

; ==============================================================================
; CONSTANT STRING DATA
; ==============================================================================
.const

; ANSI (not wide-char) names for the runtime-resolved DPI APIs --
; GetModuleHandleA/GetProcAddress both take LPCSTR, unlike the rest of
; this file's Unicode string constants.
str_user32A                    db 'user32.dll',0
str_ntdllA                     db 'ntdll.dll',0
str_uxthemeA                   db 'uxtheme.dll',0
str_fnRtlGetVersion            db 'RtlGetVersion',0
str_fnGetDpiForSystem          db 'GetDpiForSystem',0
str_fnGetDpiForWindow          db 'GetDpiForWindow',0
str_fnAdjustWindowRectExForDpi db 'AdjustWindowRectExForDpi',0

; Window class names for controls
str_ComboBox    dw 'C','o','m','b','o','B','o','x',0
str_Button      dw 'B','u','t','t','o','n',0
str_Static      dw 'S','t','a','t','i','c',0

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

; About dialog text (includes author information)
str_AboutTitle  dw 'A','b','o','u','t',0
str_AboutText   dw 'C','M','D','T',' ',' ','v','1','.','0','.','0','.','0',10,10
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
str_regThemeKey dw 'S','o','f','t','w','a','r','e','\','M','i','c','r','o','s','o','f','t','\'
                dw 'W','i','n','d','o','w','s','\','C','u','r','r','e','n','t','V','e','r','s','i','o','n','\'
                dw 'T','h','e','m','e','s','\','P','e','r','s','o','n','a','l','i','z','e',0
str_regAppsUseLightTheme dw 'A','p','p','s','U','s','e','L','i','g','h','t','T','h','e','m','e',0
str_DarkExplorer dw 'D','a','r','k','M','o','d','e','_','E','x','p','l','o','r','e','r',0
str_Explorer dw 'E','x','p','l','o','r','e','r',0

; File extensions
str_extLnk      dw '.','l','n','k',0
str_extExe      dw '.','e','x','e',0

; ==============================================================================
; COM INTERFACE IDENTIFIERS (GUIDs)
; Used for resolving Windows shortcut (.lnk) files
; ==============================================================================

; CLSID_ShellLink: {00021401-0000-0000-C000-000000000046}
; Class ID for creating IShellLink COM object
CLSID_ShellLink dd 00021401h
                dw 0000h, 0000h
                db 0C0h, 00h, 00h, 00h, 00h, 00h, 00h, 46h

; IID_IShellLinkW: {000214F9-0000-0000-C000-000000000046}
; Interface ID for Shell Link manipulation (wide char version)
IID_IShellLinkW dd 000214F9h
                dw 0000h, 0000h
                db 0C0h, 00h, 00h, 00h, 00h, 00h, 00h, 46h

; IID_IPersistFile: {0000010B-0000-0000-C000-000000000046}
; Interface ID for loading persistent files (.lnk files)
IID_IPersistFile dd 0000010bh
                 dw 0000h, 0000h
                 db 0C0h, 00h, 00h, 00h, 00h, 00h, 00h, 46h

; ==============================================================================
; RUNTIME-RESOLVED DPI API POINTERS
; ==============================================================================
.data?
g_pGetDpiForSystem           dd ?
g_pGetDpiForWindow           dd ?
g_pAdjustWindowRectExForDpi  dd ?
g_dpiApisInit                dd ?
g_darkMode                   dd ?
g_nativeDarkSupported        dd ?
g_themeApplying              dd ?          ; guards synchronous theme re-entry
g_brushBg                    dd ?
g_brushSurface               dd ?
g_brushHover                 dd ?
g_brushInput                 dd ?

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
