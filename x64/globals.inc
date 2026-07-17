; ==============================================================================
; CMDT - Run as TrustedInstaller
; Global Variable Declarations
; 
; Author: Marek Wesołowski (wesmar)
; Purpose: External declarations for global variables defined in main.asm
;          These variables are shared across all modules of the application.
; ==============================================================================

; ==============================================================================
; PRIVILEGE MANAGEMENT
; ==============================================================================

    ; Table of pointers to privilege name strings (34 privileges)
    ; Used for building privilege names like "SeDebugPrivilege"
    EXTRN   g_privTable:QWORD

; ==============================================================================
; TOKEN CACHING
; ==============================================================================

    ; Cached TrustedInstaller token handle
    ; Valid for 30 seconds to avoid repeated token acquisition
    EXTRN   g_cachedToken:QWORD
    
    ; Timestamp (in milliseconds) when cached token was obtained
    ; Used to determine token expiration
    EXTRN   g_tokenTime:DWORD

; ==============================================================================
; WINDOW HANDLES
; ==============================================================================

    ; Main application window handle
    EXTRN   g_hwndMain:QWORD
    
    ; ComboBox edit control handle (for command input)
    EXTRN   g_hwndEdit:QWORD
    
    ; Run button handle
    EXTRN   g_hwndBtn:QWORD
    
    ; Status label handle (displays operation status)
    EXTRN   g_hwndStatus:QWORD

; ==============================================================================
; APPLICATION HANDLES
; ==============================================================================

    ; Console output handle (for CLI mode)
    EXTRN   g_hConsoleOut:QWORD
    
    ; Application instance handle
    EXTRN   g_hInstance:QWORD

    ; Output-relay file handle. Non-zero only in the elevated child when
    ; the parent (non-admin) launched it with the -outfile flag, telling
    ; the child to redirect spawned process stdout/stderr to a temp file.
    EXTRN   g_relayHandle:QWORD
    EXTRN   g_relayErrHandle:QWORD

    ; Exit code reported by the last synchronously waited child. The process
    ; routine still returns BOOL, so exit code 0 remains distinct from a
    ; failure to create the process.
    EXTRN   g_childExitCode:DWORD

    ; Boolean flag: should the spawned command run in a brand new console
    ; window (CREATE_NEW_CONSOLE) instead of inheriting the parent's?
    EXTRN   g_useNewConsole:DWORD

    ; Boolean flag: is command history (MRU registry read/write) enabled?
    ; Reflects whether HKCU\Software\cmdt exists; toggled from the File menu.
    EXTRN   g_historyEnabled:DWORD

    ; File popup submenu handle (needed to update the History checkbox mark).
    EXTRN   g_hMenuFile:QWORD

; ==============================================================================
; WORKING BUFFERS
; ==============================================================================

    ; Command line buffer (full Win32 32768-WCHAR capacity)
    ; Stores the command to be executed
    EXTRN   g_cmdBuf:WORD
    
    ; Status message buffer (520 wide characters)
    ; For constructing status messages
    EXTRN   g_statusBuf:WORD
    
    ; File path buffer (520 wide characters)
    ; Stores selected file paths from browse dialog or drag-drop
    EXTRN   g_filePath:WORD
    
    ; Arguments buffer (520 wide characters)
    ; Stores command-line arguments extracted from .lnk files
    EXTRN   g_argsBuf:WORD
    
    ; Temporary buffer (1040 wide characters = 2080 bytes)
    ; General-purpose buffer for various string operations
    EXTRN   g_tempBuf:WORD

    ; Buffer for exe path (UAC elevation and context menu registration)
    EXTRN   g_exePath:WORD

    ; Buffer for decrypted strings (reusable, 520 WCHARs)
    EXTRN   g_decryptBuf:WORD

    ; Temp directory path returned by GetTempPathW (relay path setup).
    EXTRN   g_tempDirBuf:WORD

    ; Temp file path produced by GetTempFileNameW (relay payload location).
    EXTRN   g_relayPath:WORD
    EXTRN   g_relayErrPath:WORD

    ; Modified argument string built by relay parent for the elevated child.
    EXTRN   g_relayArgs:WORD

    ; 4 KB scratch buffer used to stream relay-temp bytes back to stdout.
    EXTRN   g_relayReadBuf:BYTE
