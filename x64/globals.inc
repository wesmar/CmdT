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

; ==============================================================================
; WORKING BUFFERS
; ==============================================================================

    ; Command line buffer (520 wide characters = 1040 bytes)
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
