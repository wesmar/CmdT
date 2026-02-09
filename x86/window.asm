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
;          - Menu system (File menu with Browse, About, Exit)
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

; Menu functions
CreateMenu              PROTO
CreatePopupMenu         PROTO
AppendMenuW             PROTO :DWORD,:DWORD,:DWORD,:DWORD
SetMenu                 PROTO :DWORD,:DWORD

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
RegCloseKey             PROTO :DWORD

; String utility prototypes (local implementations)
wcslen_w                PROTO :DWORD
wcscmp_ci_w             PROTO :DWORD,:DWORD
wcscat_w                PROTO :DWORD,:DWORD

; ==============================================================================
; CONSTANT STRING DATA
; ==============================================================================
.const

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
                dw 'c','m','d','t','.','e','x','e',' ','-','u','n','s','h','i','f','t',0

; File dialog filter string (double-null terminated)
str_Filter      dw 'E','x','e','c','u','t','a','b','l','e','s',0,'*','.','e','x','e',';','*','.','l','n','k',0,'A','l','l',' ','F','i','l','e','s',0,'*','.','*',0,0

; Default paths and filenames
str_DefPath     dw 'C',':','\',0
str_Shell32     dw 's','h','e','l','l','3','2','.','d','l','l',0

; Registry key for storing MRU list
str_regKey      dw 'S','o','f','t','w','a','r','e','\','c','m','d','t',0

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
; CODE SECTION
; ==============================================================================
.code

; ==============================================================================
; RunCommand - Execute Command as TrustedInstaller
;
; Purpose: Retrieves command text from the ComboBox, validates it, and executes
;          it using TrustedInstaller privileges in a new console window. Updates
;          status display and saves command to MRU list on success.
;
; Parameters:
;   hWnd - Parent window handle (for message boxes)
;
; Returns:
;   EAX = 1 on success, 0 on failure or empty command
;
; Process flow:
;   1. Get command text from ComboBox (g_hwndEdit)
;   2. Validate command is not empty
;   3. Update status to "Launching..."
;   4. Call RunAsTrustedInstaller with new console flag (1)
;   5. On success: save to MRU list and show "Process OK"
;   6. On failure: show "Failed" status
;   7. On empty: show error message box
;
; Registers modified: EAX
; ==============================================================================
RunCommand proc hWnd:DWORD
    ; Get text from ComboBox edit control
    invoke GetWindowTextW, g_hwndEdit, offset g_cmdBuf, 520
    test eax, eax
    jz rc_empty                             ; No text entered

    ; Update status to show launching
    invoke SetWindowTextW, g_hwndStatus, offset str_StatusRunning

    ; Execute command with TrustedInstaller privileges
    ; Parameter 1: new console window
    invoke FixRegeditPath, offset g_cmdBuf
    invoke RunAsTrustedInstaller, eax, 1
    test eax, eax
    jz rc_fail                              ; Execution failed

    ; Success - save to MRU list
    call SaveMRU

    ; Update status to success
    invoke SetWindowTextW, g_hwndStatus, offset str_StatusOK
    mov eax, 1
    ret

rc_fail:
    ; Execution failed - update status
    invoke SetWindowTextW, g_hwndStatus, offset str_StatusFail
    xor eax, eax
    ret

rc_empty:
    ; No command entered - show error message
    invoke MessageBoxW, hWnd, offset str_ErrEmpty, offset str_TitleErr, MB_ICONERROR
    xor eax, eax
    ret
RunCommand endp

; ==============================================================================
; WndProc - Main Window Procedure
;
; Purpose: Processes all Windows messages for the main application window.
;          This is the central message handler that responds to user
;          interactions and system events.
;
; Messages handled:
;   WM_CREATE     - Initialize child controls and menu system
;   WM_COMMAND    - Process menu selections and button clicks
;   WM_CHAR       - Handle Enter key to execute command
;   WM_SIZE       - Resize and reposition controls dynamically
;   WM_DESTROY    - Clean up and quit application
;   WM_DROPFILES  - Handle drag-and-drop file operations
;   WM_KEYDOWN    - Handle ESC key to exit
;
; Parameters:
;   hWnd   - Window handle
;   uMsg   - Message identifier
;   wParam - First message parameter (varies by message)
;   lParam - Second message parameter (varies by message)
;
; Returns:
;   EAX = Message-specific return value (0 if processed, DefWindowProc if not)
;
; Registers used: EBX, ESI, EDI (preserved via uses clause)
; ==============================================================================
WndProc proc uses ebx esi edi hWnd:DWORD, uMsg:DWORD, wParam:DWORD, lParam:DWORD
    LOCAL rect[4]:DWORD                     ; Client rectangle (RECT structure)
    LOCAL ofn[22]:DWORD                     ; OPENFILENAMEW structure
    LOCAL hMenu:DWORD                       ; Main menu handle
    LOCAL hFileMenu:DWORD                   ; File submenu handle
    LOCAL notifyCode:DWORD                  ; Notification code from WM_COMMAND
    LOCAL hBtnBrowse:DWORD                  ; Browse button handle
    
    ; Dispatch message to appropriate handler
    mov eax, uMsg
    cmp eax, WM_CREATE
    je wp_create
    cmp eax, WM_COMMAND
    je wp_command
    cmp eax, WM_CHAR
    je wp_char
    cmp eax, WM_SIZE
    je wp_size
    cmp eax, WM_DESTROY
    je wp_destroy
    cmp eax, WM_DROPFILES
    je wp_dropfiles
    cmp eax, WM_KEYDOWN
    je wp_keydown
    jmp wp_defproc                          ; Unhandled message

wp_create:
    ; ===== WM_CREATE: Initialize window contents =====
    ; Extract instance handle from CREATESTRUCT
    mov eax, lParam
    mov eax, [eax+4]                        ; hInstance at offset 4
    
    ; Create main menu
    invoke CreateMenu
    mov hMenu, eax
    
    ; Create File submenu
    invoke CreatePopupMenu
    mov hFileMenu, eax
    
    ; Add menu items to File menu
    invoke AppendMenuW, hFileMenu, MF_STRING, ID_FILE_BROWSE, offset str_MenuBrowse
    invoke AppendMenuW, hFileMenu, MF_SEPARATOR, 0, 0
    invoke AppendMenuW, hFileMenu, MF_STRING, ID_FILE_ABOUT, offset str_MenuAbout
    invoke AppendMenuW, hFileMenu, MF_SEPARATOR, 0, 0
    invoke AppendMenuW, hFileMenu, MF_STRING, ID_FILE_EXIT, offset str_MenuExit
    
    ; Add File submenu to main menu
    invoke AppendMenuW, hMenu, MF_POPUP, hFileMenu, offset str_MenuFile
    
    ; Attach menu to window
    invoke SetMenu, hWnd, hMenu
    
    ; Create ComboBox control (dropdown with edit field and MRU list)
    ; Style: WS_CHILD | WS_VISIBLE | WS_TABSTOP | CBS_DROPDOWN | CBS_AUTOHSCROLL
    ; Initial position will be adjusted by WM_SIZE
    mov eax, lParam
    mov eax, [eax+4]                        ; hInstance
    invoke CreateWindowExW, WS_EX_CLIENTEDGE, offset str_ComboBox, 0, STY_COMBO, 10, 10, 460, 200, hWnd, IDC_COMBO, eax, 0
    test eax, eax
    jz wp_create_fail
    mov g_hwndEdit, eax
    
    ; Load MRU list from registry into ComboBox
    call LoadMRU
    
    ; Create Browse button
    mov eax, lParam
    mov eax, [eax+4]
    invoke CreateWindowExW, 0, offset str_Button, offset str_BtnBrowse, STY_BUTTON, 10, 45, 80, 25, hWnd, IDC_BTN_BROWSE, eax, 0
    test eax, eax
    jz wp_create_fail

    ; Create Run button (default push button)
    mov eax, lParam
    mov eax, [eax+4]
    invoke CreateWindowExW, 0, offset str_Button, offset str_BtnRun, STY_BUTTON_DEF, 400, 45, 80, 25, hWnd, IDC_BTN_RUN, eax, 0
    test eax, eax
    jz wp_create_fail
    mov g_hwndBtn, eax
    
    ; Create status label (static text control)
    mov eax, lParam
    mov eax, [eax+4]
    invoke CreateWindowExW, 0, offset str_Static, offset str_StatusInit, STY_STATIC, 100, 48, 200, 20, hWnd, IDC_STATIC_STATUS, eax, 0
    test eax, eax
    jz wp_create_fail
    mov g_hwndStatus, eax
    
    ; Set initial focus to ComboBox
    invoke SetFocus, g_hwndEdit
    
    ; Enable drag-and-drop file acceptance
    invoke DragAcceptFiles, hWnd, 1
    
    ; Bypass UIPI (User Interface Privilege Isolation) to allow
    ; drag-and-drop from non-elevated Explorer windows
    invoke ChangeWindowMessageFilterEx, hWnd, WM_DROPFILES, MSGFLT_ALLOW, 0
    invoke ChangeWindowMessageFilterEx, hWnd, WM_COPYGLOBALDATA, MSGFLT_ALLOW, 0

    xor eax, eax                            ; Return 0 (success)
    ret

wp_create_fail:
    mov eax, -1                             ; Return -1 (destroy window)
    ret
    
wp_command:
    ; ===== WM_COMMAND: Handle menu and control notifications =====
    ; Extract notification code and control ID
    mov eax, wParam
    shr eax, 16
    mov notifyCode, eax                     ; High word = notification code
    mov eax, wParam
    and eax, 0FFFFh
    ; EAX now contains control/menu ID
    
    ; Dispatch by control/menu ID
    cmp eax, IDC_EDIT
    je wp_check_edit
    cmp eax, IDC_BTN_RUN
    je wp_cmd_run
    cmp eax, 1                              ; IDOK (Enter key default action)
    je wp_cmd_run
    cmp eax, IDC_BTN_BROWSE
    je wp_cmd_browse
    cmp eax, ID_FILE_BROWSE
    je wp_cmd_browse
    cmp eax, ID_FILE_EXIT
    je wp_cmd_exit
    cmp eax, ID_FILE_ABOUT
    je wp_cmd_about
    jmp wp_defproc

wp_check_edit:
    ; Check edit control notification
    cmp notifyCode, EN_SETFOCUS
    je wp_edit_focus
    jmp wp_defproc
    
wp_edit_focus:
    ; Edit control got focus - make Run button default
    invoke SendMessageW, g_hwndBtn, BM_SETSTYLE, BS_DEFPUSHBUTTON, 1
    xor eax, eax
    ret

wp_char:
    ; ===== WM_CHAR: Handle character input =====
    mov eax, wParam
    cmp eax, 13                             ; Check for Enter key (CR)
    je wp_char_enter
    jmp wp_defproc
    
wp_char_enter:
    ; Enter key pressed - execute if ComboBox has focus
    invoke GetFocus
    cmp eax, g_hwndEdit
    jne wp_defproc                          ; Wrong control has focus
    invoke RunCommand, hWnd
    xor eax, eax
    ret
    
wp_cmd_run:
    ; Run button clicked or Enter pressed
    invoke RunCommand, hWnd
    ret
    
wp_cmd_browse:
    ; ===== Browse button/menu: Open file dialog =====
    ; Initialize OPENFILENAMEW structure
    lea edi, ofn
    xor eax, eax
    mov ecx, 22                             ; Clear structure
    rep stosd
    
    ; Set structure fields
    mov dword ptr [ofn], 88                 ; lStructSize
    mov eax, hWnd
    mov dword ptr [ofn+4], eax              ; hwndOwner
    lea eax, str_Filter
    mov dword ptr [ofn+12], eax             ; lpstrFilter
    lea eax, g_filePath
    mov dword ptr [ofn+28], eax             ; lpstrFile
    mov dword ptr [ofn+32], 520             ; nMaxFile
    lea eax, str_DefPath
    mov dword ptr [ofn+44], eax             ; lpstrInitialDir
    mov dword ptr [ofn+52], OFN_FILEMUSTEXIST or OFN_PATHMUSTEXIST ; Flags
    
    ; Clear file path buffer
    mov g_filePath, 0
    
    ; Display file open dialog
    invoke GetOpenFileNameW, addr ofn
    test eax, eax
    jz wp_browse_cancel                     ; User cancelled

    ; Check if selected file is a .lnk shortcut
    invoke wcslen_w, offset g_filePath
    mov ecx, eax
    cmp ecx, 4
    jl wp_browse_set                        ; Too short for .lnk
    
    ; Point to last 4 characters (.xxx)
    lea esi, g_filePath
    lea esi, [esi + ecx*2 - 8]
    invoke wcscmp_ci_w, esi, offset str_extLnk
    test eax, eax
    jnz wp_browse_resolve_lnk               ; Is .lnk file
    jmp wp_browse_set                       ; Not .lnk

wp_browse_resolve_lnk:
    ; Resolve .lnk shortcut to target executable and arguments
    invoke ResolveLnkPath, offset g_filePath, offset g_cmdBuf, offset g_tempBuf
    test eax, eax
    jz wp_browse_set                        ; Resolution failed, use path as-is
    
    ; Check if target path is empty
    invoke wcslen_w, offset g_cmdBuf
    test eax, eax
    jz wp_browse_args_only                  ; No target, just use args
    
    ; Append space after target path
    mov esi, offset g_cmdBuf
    lea esi, [esi + eax*2]
    mov word ptr [esi], ' '
    add esi, 2
    mov word ptr [esi], 0
    
wp_browse_args_only:
    ; Concatenate embedded shortcut arguments
    invoke wcscat_w, offset g_cmdBuf, offset g_tempBuf
    invoke SetWindowTextW, g_hwndEdit, offset g_cmdBuf
    jmp wp_browse_cancel

wp_browse_set:
    ; Set file path directly in ComboBox
    invoke SetWindowTextW, g_hwndEdit, offset g_filePath

wp_browse_cancel:
    ; Restore focus to ComboBox
    invoke SetFocus, g_hwndEdit
    invoke SendMessageW, g_hwndBtn, BM_SETSTYLE, BS_DEFPUSHBUTTON, 1
    xor eax, eax
    ret
    
wp_cmd_about:
    ; Display About dialog
    invoke MessageBoxW, hWnd, offset str_AboutText, offset str_AboutTitle, MB_ICONINFORMATION
    xor eax, eax
    ret

wp_cmd_exit:
    ; Exit menu item - quit application
    invoke PostQuitMessage, 0
    xor eax, eax
    ret
    
wp_size:
    ; ===== WM_SIZE: Resize controls to fit window =====
    invoke GetClientRect, hWnd, addr rect
    mov esi, [rect+8]                       ; Width
    mov edi, [rect+12]                      ; Height
    
    ; Row 1: ComboBox (spans full width with margins)
    ; Position: (10, 10), Size: (Width-20, 25)
    mov ebx, esi
    sub ebx, 20
    invoke MoveWindow, g_hwndEdit, 10, 10, ebx, 25, 1
    
    ; Row 2: Browse button, Status label, Run button
    ; All aligned at y=45
    
    ; Browse Button: Left-aligned
    ; Position: (10, 45), Size: (80, 25)
    invoke GetDlgItem, hWnd, IDC_BTN_BROWSE
    mov hBtnBrowse, eax
    invoke MoveWindow, hBtnBrowse, 10, 45, 80, 25, 1
    
    ; Run Button: Right-aligned
    ; Position: (Width-90, 45), Size: (80, 25)
    mov ebx, esi
    sub ebx, 90
    push ebx                                ; Save X position
    invoke MoveWindow, g_hwndBtn, ebx, 45, 80, 25, 1
    pop ebx                                 ; Restore X position
    
    ; Status Label: Centered between Browse and Run buttons
    ; Position: (100, 48), Width: RunBtnX - 110
    sub ebx, 110                            ; Calculate width
    invoke MoveWindow, g_hwndStatus, 100, 48, ebx, 20, 1
    
    xor eax, eax
    ret
    
wp_dropfiles:
    ; ===== WM_DROPFILES: Handle file drag-and-drop =====
    ; wParam contains HDROP handle
    invoke DragQueryFileW, wParam, 0, offset g_filePath, 520
    invoke DragFinish, wParam
    
    ; Check if dropped file is a .lnk shortcut
    invoke wcslen_w, offset g_filePath
    mov ecx, eax
    cmp ecx, 4
    jl wp_drop_set                          ; Too short
    
    ; Check last 4 characters for .lnk
    lea esi, g_filePath
    lea esi, [esi + ecx*2 - 8]
    invoke wcscmp_ci_w, esi, offset str_extLnk
    test eax, eax
    jnz wp_drop_resolve_lnk                 ; Is .lnk
    jmp wp_drop_set

wp_drop_resolve_lnk:
    ; Resolve .lnk shortcut
    invoke ResolveLnkPath, offset g_filePath, offset g_cmdBuf, offset g_tempBuf
    test eax, eax
    jz wp_drop_set                          ; Resolution failed
    
    ; Check if target path is empty
    invoke wcslen_w, offset g_cmdBuf
    test eax, eax
    jz wp_drop_args_only
    
    ; Append space after target
    mov esi, offset g_cmdBuf
    lea esi, [esi + eax*2]
    mov word ptr [esi], ' '
    add esi, 2
    mov word ptr [esi], 0
    
wp_drop_args_only:
    ; Concatenate arguments
    invoke wcscat_w, offset g_cmdBuf, offset g_tempBuf
    invoke SetWindowTextW, g_hwndEdit, offset g_cmdBuf
    jmp wp_drop_run

wp_drop_set:
    ; Set file path directly
    invoke SetWindowTextW, g_hwndEdit, offset g_filePath

wp_drop_run:
    ; Auto-execute dropped file
    invoke RunCommand, hWnd
    xor eax, eax
    ret

wp_keydown:
    ; ===== WM_KEYDOWN: Handle special keys =====
    cmp wParam, VK_ESCAPE
    jne wp_defproc
    ; ESC key pressed - exit application
    invoke PostQuitMessage, 0
    xor eax, eax
    ret

wp_destroy:
    ; ===== WM_DESTROY: Window is being destroyed =====
    invoke PostQuitMessage, 0
    xor eax, eax
    ret

wp_defproc:
    ; Unhandled message - pass to default window procedure
    invoke DefWindowProcW, hWnd, uMsg, wParam, lParam
    ret
WndProc endp

; ==============================================================================
; LoadMRU - Load Most Recently Used Commands from Registry
;
; Purpose: Reads up to 5 most recently used commands from the Windows registry
;          and populates the ComboBox dropdown list. Commands are stored as
;          registry values "0" through "4" under HKCU\Software\cmdt.
;
; Parameters: None (uses global variables)
;
; Returns: None
;
; Process flow:
;   1. Open registry key (HKEY_CURRENT_USER\Software\cmdt)
;   2. Loop through values "0" through "4" (0 = most recent)
;   3. For each value, read command string
;   4. Add command to ComboBox using CB_ADDSTRING
;   5. Set ComboBox selection to first item (most recent)
;   6. Close registry key
;
; Registry structure:
;   Key: HKEY_CURRENT_USER\Software\cmdt
;   Values: "0" = newest, "4" = oldest (REG_SZ type)
;
; Registers modified: EAX, EBX, ESI, EDI
; ==============================================================================
LoadMRU proc uses ebx esi edi
    LOCAL hKey:DWORD                        ; Registry key handle
    LOCAL valName[4]:WORD                   ; Value name buffer ("0" to "4")
    LOCAL valDataLen:DWORD                  ; Data length for RegQueryValueExW
    LOCAL idx:DWORD                         ; Loop index

    ; Open registry key for reading
    invoke RegOpenKeyExW, HKEY_CURRENT_USER, offset str_regKey, 0, KEY_READ, addr hKey
    test eax, eax
    jnz lm_done                             ; Key doesn't exist, exit

    ; Loop through MRU values (0-4)
    mov idx, 0
lm_loop:
    cmp idx, MRU_MAX_ITEMS
    jge lm_setsel                           ; All items loaded

    ; Prepare value name ("0", "1", etc.)
    mov ax, '0'
    add ax, word ptr [idx]
    mov word ptr [valName], ax
    mov word ptr [valName+2], 0

    ; Read value from registry
    mov valDataLen, 1040                    ; Buffer size in bytes
    invoke RegQueryValueExW, hKey, addr valName, 0, 0, offset g_tempBuf, addr valDataLen
    test eax, eax
    jnz lm_setsel                           ; Value doesn't exist, stop

    ; Add command to ComboBox
    invoke SendMessageW, g_hwndEdit, CB_ADDSTRING, 0, offset g_tempBuf

    inc idx
    jmp lm_loop

lm_setsel:
    ; Select first item in ComboBox (most recent command)
    invoke SendMessageW, g_hwndEdit, CB_SETCURSEL, 0, 0
    invoke RegCloseKey, hKey

lm_done:
    ret
LoadMRU endp

; ==============================================================================
; SaveMRU - Save Command to Most Recently Used List
;
; Purpose: Saves the current command to the MRU registry list. Shifts existing
;          entries down (0→1, 1→2, etc.) and inserts new command at position 0.
;          Maintains maximum of 5 entries by deleting oldest when full.
;
; Parameters: None (uses global variables)
;
; Returns: None
;
; Process flow:
;   1. Get command text from ComboBox
;   2. Open/create registry key
;   3. Delete value "4" (oldest entry, if exists)
;   4. Shift values down: "3"→"4", "2"→"3", "1"→"2", "0"→"1"
;   5. Write new command as value "0"
;   6. Insert command at position 0 in ComboBox
;   7. Limit ComboBox to 5 items (delete oldest if needed)
;   8. Select first item
;   9. Close registry key
;
; Registry structure:
;   Key: HKEY_CURRENT_USER\Software\cmdt
;   Values: "0" = newest, "4" = oldest (REG_SZ type)
;
; Registers modified: EAX, EBX, ESI, EDI
; ==============================================================================
SaveMRU proc uses ebx esi edi
    LOCAL hKey:DWORD                        ; Registry key handle
    LOCAL disp:DWORD                        ; Disposition (key created/opened)
    LOCAL valName[4]:WORD                   ; Value name buffer
    LOCAL textLen:DWORD                     ; Command text length in bytes
    LOCAL idx:DWORD                         ; Loop index
    LOCAL dataLen:DWORD                     ; Data length for read operations

    ; Get command text from ComboBox
    invoke GetWindowTextW, g_hwndEdit, offset g_cmdBuf, 520
    test eax, eax
    jz sm_done                              ; Empty text, don't save
    
    ; Calculate text length in bytes (including null terminator)
    shl eax, 1                              ; Length in wide chars → bytes
    add eax, 2                              ; Add null terminator
    mov textLen, eax

    ; Open or create registry key
    invoke RegCreateKeyExW, HKEY_CURRENT_USER, offset str_regKey, 0, 0, 0, KEY_WRITE or KEY_READ, 0, addr hKey, addr disp
    test eax, eax
    jnz sm_done                             ; Key creation failed

    ; Delete oldest value "4" to make room (ignore if doesn't exist)
    mov word ptr [valName], '4'
    mov word ptr [valName+2], 0
    invoke RegDeleteValueW, hKey, addr valName

    ; Shift existing values down: 3→4, 2→3, 1→2, 0→1
    mov idx, 3
sm_shift_loop:
    cmp idx, 0
    jl sm_write_new                         ; All shifted, write new

    ; Read value at index
    mov ax, '0'
    add ax, word ptr [idx]
    mov word ptr [valName], ax
    mov word ptr [valName+2], 0

    mov dataLen, 1040
    invoke RegQueryValueExW, hKey, addr valName, 0, 0, offset g_tempBuf, addr dataLen
    test eax, eax
    jnz sm_shift_next                       ; Value doesn't exist, skip

    ; Write to index+1
    mov ax, '1'
    add ax, word ptr [idx]
    mov word ptr [valName], ax
    mov word ptr [valName+2], 0

    invoke RegSetValueExW, hKey, addr valName, 0, REG_SZ, offset g_tempBuf, dataLen

sm_shift_next:
    dec idx
    jmp sm_shift_loop

sm_write_new:
    ; Write new command as value "0"
    mov word ptr [valName], '0'
    mov word ptr [valName+2], 0
    invoke RegSetValueExW, hKey, addr valName, 0, REG_SZ, offset g_cmdBuf, textLen

    ; Add to ComboBox at position 0
    invoke SendMessageW, g_hwndEdit, CB_INSERTSTRING, 0, offset g_cmdBuf

    ; Limit ComboBox to 5 items
    invoke SendMessageW, g_hwndEdit, CB_GETCOUNT, 0, 0
    cmp eax, MRU_MAX_ITEMS
    jle sm_setsel
    invoke SendMessageW, g_hwndEdit, CB_DELETESTRING, MRU_MAX_ITEMS, 0

sm_setsel:
    ; Select first item
    invoke SendMessageW, g_hwndEdit, CB_SETCURSEL, 0, 0
    invoke RegCloseKey, hKey

sm_done:
    ret
SaveMRU endp

; ==============================================================================
; ResolveLnkPath - Resolve Windows Shortcut (.lnk) File
;
; Purpose: Uses COM interfaces to resolve a .lnk shortcut file to its target
;          executable path and command-line arguments embedded in the shortcut.
;
; Parameters:
;   lpLnkPath  - Pointer to .lnk file path (wide string)
;   lpOutPath  - Pointer to output buffer for target path (min 520 wide chars)
;   lpArgsPath - Pointer to output buffer for arguments (min 520 wide chars)
;
; Returns:
;   EAX = 1 on success, 0 on failure
;
; Output buffers:
;   lpOutPath  - Receives target executable path (e.g., "C:\Windows\notepad.exe")
;   lpArgsPath - Receives command-line arguments embedded in shortcut
;
; COM interfaces used:
;   IShellLinkW  - Shell Link manipulation interface
;   IPersistFile - Persistent file loading interface
;
; Process flow:
;   1. Initialize COM (CoInitialize)
;   2. Create IShellLink object (CoCreateInstance)
;   3. Query for IPersistFile interface
;   4. Load .lnk file (IPersistFile::Load)
;   5. Get target path (IShellLinkW::GetPath)
;   6. Zero argument buffer
;   7. Get arguments (IShellLinkW::GetArguments)
;   8. Release IPersistFile interface
;   9. Release IShellLink interface
;   10. Uninitialize COM
;   11. Return success/failure based on GetPath result
;
; Error handling: All COM interface pointers are properly released on error
;
; Registers modified: EAX, EBX, ESI, EDI
; ==============================================================================
ResolveLnkPath proc uses ebx esi edi lpLnkPath:DWORD, lpOutPath:DWORD, lpArgsPath:DWORD
    LOCAL pShellLink:DWORD                  ; IShellLink interface pointer
    LOCAL pPersistFile:DWORD                ; IPersistFile interface pointer
    LOCAL hr:DWORD                          ; HRESULT from GetPath

    ; Initialize COM library
    invoke CoInitialize, 0

    ; Create ShellLink COM object
    invoke CoCreateInstance, offset CLSID_ShellLink, 0, CLSCTX_INPROC_SERVER, offset IID_IShellLinkW, addr pShellLink
    test eax, eax
    jnz rlp_fail_uninit                     ; COM object creation failed

    ; Query for IPersistFile interface
    mov eax, pShellLink
    mov eax, [eax]                          ; Get vtable pointer
    mov ebx, pShellLink
    lea ecx, pPersistFile
    push ecx                                ; ppvObject
    push offset IID_IPersistFile            ; riid
    push ebx                                ; this
    call dword ptr [eax]                    ; QueryInterface at vtable[0]
    test eax, eax
    jnz rlp_release_link                    ; QI failed

    ; Load .lnk file using IPersistFile::Load
    mov eax, pPersistFile
    mov eax, [eax]                          ; Get vtable
    mov ebx, pPersistFile
    push 0                                  ; grfMode = 0
    push lpLnkPath                          ; pszFileName
    push ebx                                ; this
    call dword ptr [eax+20]                 ; Load at vtable[5] = offset 20
    test eax, eax
    jnz rlp_release_both                    ; Load failed

    ; Get target path using IShellLinkW::GetPath
    mov eax, pShellLink
    mov eax, [eax]                          ; Get vtable
    mov ebx, pShellLink
    push 0                                  ; fFlags = 0
    push 0                                  ; pfd = NULL
    push 260                                ; cch = MAX_PATH
    push lpOutPath                          ; pszFile
    push ebx                                ; this
    call dword ptr [eax+12]                 ; GetPath at vtable[3] = offset 12
    mov hr, eax                             ; Save result

    ; Zero argument buffer before getting arguments
    push edi
    mov edi, lpArgsPath
    xor eax, eax
    mov ecx, 260
    rep stosd                               ; Clear 260 DWORDs (1040 bytes)
    pop edi

    ; Get arguments using IShellLinkW::GetArguments
    mov eax, pShellLink
    mov eax, [eax]                          ; Get vtable
    mov ebx, pShellLink
    push 520                                ; cch (buffer size)
    push lpArgsPath                         ; pszArgs
    push ebx                                ; this
    call dword ptr [eax+40]                 ; GetArguments at vtable[10] = offset 40

    ; Release IPersistFile interface
    mov eax, pPersistFile
    mov eax, [eax]
    mov ebx, pPersistFile
    push ebx
    call dword ptr [eax+8]                  ; Release at vtable[2] = offset 8

    ; Release IShellLink interface
    mov eax, pShellLink
    mov eax, [eax]
    mov ebx, pShellLink
    push ebx
    call dword ptr [eax+8]                  ; Release at vtable[2] = offset 8

    ; Uninitialize COM
    invoke CoUninitialize

    ; Return success/failure based on GetPath result
    mov eax, hr
    test eax, eax
    jnz rlp_fail
    mov eax, 1                              ; Success
    ret

rlp_release_both:
    ; Release both interfaces on error
    mov eax, pPersistFile
    mov eax, [eax]
    mov ebx, pPersistFile
    push ebx
    call dword ptr [eax+8]

rlp_release_link:
    ; Release ShellLink interface
    mov eax, pShellLink
    mov eax, [eax]
    mov ebx, pShellLink
    push ebx
    call dword ptr [eax+8]

rlp_fail_uninit:
    ; Uninitialize COM on error
    invoke CoUninitialize

rlp_fail:
    xor eax, eax                            ; Return failure
    ret
ResolveLnkPath endp

; ==============================================================================
; wcslen_w - Wide Character String Length
;
; Purpose: Calculates the length of a null-terminated wide character string.
;
; Parameters:
;   lpStr - Pointer to wide character string
;
; Returns:
;   EAX = Number of characters (excluding null terminator)
;
; Registers modified: EAX, ECX
; ==============================================================================
wcslen_w proc lpStr:DWORD
    mov eax, lpStr
    xor ecx, ecx                            ; Counter
@@:
    cmp word ptr [eax + ecx*2], 0           ; Check for null terminator
    je @F
    inc ecx
    jmp @B
@@:
    mov eax, ecx                            ; Return count
    ret
wcslen_w endp

; ==============================================================================
; wcscmp_ci_w - Wide Character String Compare (Case-Insensitive)
;
; Purpose: Compares two wide character strings ignoring case differences.
;
; Parameters:
;   str1 - First string pointer
;   str2 - Second string pointer
;
; Returns:
;   EAX = 1 if strings match (case-insensitive), 0 otherwise
;
; Registers modified: EAX, EDX, ESI, EDI
; ==============================================================================
wcscmp_ci_w proc str1:DWORD, str2:DWORD
    push esi
    push edi
    mov esi, str1
    mov edi, str2
wcs_loop:
    mov ax, word ptr [esi]
    mov dx, word ptr [edi]
    ; Lowercase first character
    cmp ax, 'A'
    jb wcs_skip_lower1
    cmp ax, 'Z'
    ja wcs_skip_lower1
    add ax, 32                              ; A-Z → a-z
wcs_skip_lower1:
    ; Lowercase second character
    cmp dx, 'A'
    jb wcs_skip_lower2
    cmp dx, 'Z'
    ja wcs_skip_lower2
    add dx, 32
wcs_skip_lower2:
    cmp ax, dx                              ; Compare
    jne wcs_not_eq
    test ax, ax
    jz wcs_equal                            ; End of strings
    add esi, 2
    add edi, 2
    jmp wcs_loop
wcs_equal:
    pop edi
    pop esi
    mov eax, 1                              ; Match
    ret
wcs_not_eq:
    pop edi
    pop esi
    xor eax, eax                            ; No match
    ret
wcscmp_ci_w endp

; ==============================================================================
; wcscat_w - Wide Character String Concatenate
;
; Purpose: Appends source string to destination string.
;
; Parameters:
;   dest - Destination buffer
;   src  - Source string
;
; Returns: None
;
; Registers modified: EAX, ESI, EDI
; ==============================================================================
wcscat_w proc dest:DWORD, src:DWORD
    push esi
    push edi
    mov edi, dest
@@:
    cmp word ptr [edi], 0                   ; Find end
    je @F
    add edi, 2
    jmp @B
@@:
    mov esi, src
@@:
    mov ax, word ptr [esi]                  ; Copy
    mov word ptr [edi], ax
    test ax, ax
    jz @F
    add esi, 2
    add edi, 2
    jmp @B
@@:
    pop edi
    pop esi
    ret
wcscat_w endp

; ==============================================================================
; CreateMainWindow - Create and Initialize Main Application Window
;
; Purpose: Registers the window class and creates the main application window.
;          Loads application icon from shell32.dll.
;
; Parameters:
;   hInstance - Application instance handle
;
; Returns:
;   EAX = Main window handle on success, 0 on failure
;
; Window properties:
;   Class name: "TIRunnerClass"
;   Title: "Run as TrustedInstaller"
;   Size: 420x150 pixels
;   Position: (100, 100)
;   Style: Overlapped window with system menu
;   Icon: Extracted from shell32.dll (index 104 - shield icon)
;
; Registers modified: EAX, EBX, ESI, EDI
; ==============================================================================
CreateMainWindow proc uses ebx esi edi hInstance:DWORD
    LOCAL wc[10]:DWORD                      ; WNDCLASSW structure
    LOCAL hIconLarge:DWORD                  ; Large icon handle
    LOCAL hIconSmall:DWORD                  ; Small icon handle

    ; Initialize WNDCLASSW structure
    lea edi, wc
    xor eax, eax
    mov ecx, 10                             ; 10 DWORDs = 40 bytes
    rep stosd
    
    ; Set window class fields
    mov dword ptr [wc], CS_HREDRAW or CS_VREDRAW ; style
    lea eax, WndProc
    mov dword ptr [wc+4], eax               ; lpfnWndProc
    mov eax, hInstance
    mov dword ptr [wc+16], eax              ; hInstance
    
    ; Extract icon from shell32.dll (index 104 = shield icon)
    invoke ExtractIconExW, offset str_Shell32, 104, addr hIconLarge, addr hIconSmall, 1
    mov eax, hIconLarge
    mov dword ptr [wc+20], eax              ; hIcon
    
    ; Load standard arrow cursor
    invoke LoadCursorW, 0, IDC_ARROW_ATOM
    mov dword ptr [wc+24], eax              ; hCursor
    
    ; Use white background brush
    invoke GetStockObject, WHITE_BRUSH
    mov dword ptr [wc+28], eax              ; hbrBackground
    
    ; Set class name
    lea eax, str_ClassName
    mov dword ptr [wc+36], eax              ; lpszClassName
    
    ; Register window class
    invoke RegisterClassW, addr wc
    test eax, eax
    jz cmw_fail
    
    ; Create main window
    ; Style: WS_OVERLAPPEDWINDOW | WS_VISIBLE
    invoke CreateWindowExW, 0, offset str_ClassName, offset str_Title, STY_MAINWIN, 100, 100, 420, 150, 0, 0, hInstance, 0
    test eax, eax
    jz cmw_fail
    mov g_hwndMain, eax
    
    ; Show and update window
    invoke ShowWindow, eax, SW_SHOWNORMAL
    invoke UpdateWindow, g_hwndMain
    mov eax, g_hwndMain
    ret

cmw_fail:
    xor eax, eax                            ; Return NULL on failure
    ret
CreateMainWindow endp

end
