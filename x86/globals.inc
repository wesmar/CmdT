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
EXTRN g_relayHandle:DWORD       ; Relay output file handle for elevated -outfile mode

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

EXTRN g_exePath:WORD            ; Exe path (UAC self-elevate + context menu)
EXTRN g_useNewConsole:DWORD     ; New console window flag for spawned command
EXTRN g_tempDirBuf:WORD         ; Temp dir buffer for relay setup
EXTRN g_relayPath:WORD          ; Temp file path produced by relay parent
EXTRN g_relayArgs:WORD          ; Modified args string for elevated relay child
EXTRN g_relayReadBuf:BYTE       ; Streaming buffer for relay-back to stdout

; Promoted from LOCALs of `start` so cli.asm's dispatch labels can reach them.
EXTRN g_argv:DWORD              ; Argument vector (LocalFree-able)
EXTRN g_argc:DWORD              ; Argument count
EXTRN g_argv1:DWORD             ; Cached argv[1] pointer
EXTRN g_sa:DWORD                ; SECURITY_ATTRIBUTES for the relay file
