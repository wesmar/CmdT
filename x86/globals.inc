; ==============================================================================
; CMDT - Run as TrustedInstaller
; Global Variables and External Declarations
; 
; Author: Marek Wesołowski (wesmar)
; Purpose: Declares external global variables that are defined in main.asm
;          and shared across all modules. These variables provide shared state
;          for the application including window handles, buffers, and security
;          token caching.
;
; Usage: Include this file in any module that needs access to global state.
;        The actual storage for these variables is allocated in main.asm.
; ==============================================================================

; ==============================================================================
; PRIVILEGE TABLE
; Array of pointers to privilege name strings (34 elements)
; Defined in main.asm, used by token.asm for privilege enumeration
; ==============================================================================
EXTRN g_privTable:DWORD         ; Privilege name pointer array (34 entries)

; ==============================================================================
; TOKEN CACHING VARIABLES
; Used for caching TrustedInstaller token to avoid repeated privilege
; escalation overhead. Token is valid for 30 seconds from acquisition.
; ==============================================================================
EXTRN g_cachedToken:DWORD       ; Cached TrustedInstaller token handle
EXTRN g_tokenTime:DWORD         ; Timestamp of cached token (GetTickCount)

; ==============================================================================
; WINDOW HANDLES
; Global handles to main window and child controls
; ==============================================================================
EXTRN g_hwndMain:DWORD          ; Main application window handle
EXTRN g_hwndEdit:DWORD          ; ComboBox/Edit control handle (command input)
EXTRN g_hwndBtn:DWORD           ; Run button handle
EXTRN g_hwndStatus:DWORD        ; Status label handle (shows operation status)

; ==============================================================================
; CONSOLE AND INSTANCE HANDLES
; ==============================================================================
EXTRN g_hConsoleOut:DWORD       ; Console output handle (for CLI mode)
EXTRN g_hInstance:DWORD         ; Application instance handle (HINSTANCE)

; ==============================================================================
; GLOBAL BUFFERS
; Large string buffers for command processing and file operations
; All buffers are wide character (Unicode) arrays
; ==============================================================================
EXTRN g_cmdBuf:WORD             ; Command line buffer (520 wide chars = 1040 bytes)
EXTRN g_statusBuf:WORD          ; Status message buffer (520 wide chars)
EXTRN g_filePath:WORD           ; File path buffer (520 wide chars)
EXTRN g_argsBuf:WORD            ; Arguments buffer (520 wide chars)
EXTRN g_tempBuf:WORD            ; Temporary work buffer (1040 wide chars = 2080 bytes)
