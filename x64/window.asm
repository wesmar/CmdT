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
;          - Menu system (File menu with Browse, About, Exit)
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
EXTRN SetMenu:PROC

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

; About dialog text (includes author information and all CLI commands)
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
; CODE SECTION
; ==============================================================================
.code

; ==============================================================================
; wcslen_w - Wide Character String Length
;
; Purpose: Calculates the length of a null-terminated wide character string
;
; Parameters:
;   RCX = Pointer to wide character string
;
; Returns:
;   RAX = Number of characters (excluding null terminator)
; ==============================================================================
wcslen_w proc
    mov rax, rcx                ; RAX = string pointer
    xor r8, r8                  ; R8 = character counter
@@:
    cmp word ptr [rax + r8*2], 0 ; Check for null terminator
    je @F
    inc r8                      ; Increment counter
    jmp @B
@@:
    mov rax, r8                 ; Return count
    ret
wcslen_w endp

; ==============================================================================
; wcscmp_ci_w - Wide Character String Compare (Case-Insensitive)
;
; Purpose: Compares two wide character strings ignoring case
;
; Parameters:
;   RCX = First string pointer
;   RDX = Second string pointer
;
; Returns:
;   RAX = 1 if strings are equal (case-insensitive), 0 otherwise
; ==============================================================================
wcscmp_ci_w proc
    push rsi
    push rdi
    mov rsi, rcx                ; RSI = first string
    mov rdi, rdx                ; RDI = second string
wcs_loop:
    mov ax, word ptr [rsi]      ; Read character from first string
    mov dx, word ptr [rdi]      ; Read character from second string
    
    ; Convert first character to lowercase if uppercase (A-Z)
    cmp ax, 'A'
    jb wcs_skip_lower1
    cmp ax, 'Z'
    ja wcs_skip_lower1
    add ax, 32                  ; Convert to lowercase
wcs_skip_lower1:
    
    ; Convert second character to lowercase if uppercase (A-Z)
    cmp dx, 'A'
    jb wcs_skip_lower2
    cmp dx, 'Z'
    ja wcs_skip_lower2
    add dx, 32                  ; Convert to lowercase
wcs_skip_lower2:
    
    cmp ax, dx                  ; Compare characters
    jne wcs_not_eq              ; Different
    test ax, ax                 ; End of strings?
    jz wcs_equal                ; Both null terminators - equal
    add rsi, 2                  ; Next character
    add rdi, 2
    jmp wcs_loop
wcs_equal:
    pop rdi
    pop rsi
    mov rax, 1                  ; Return 1 (equal)
    ret
wcs_not_eq:
    pop rdi
    pop rsi
    xor rax, rax                ; Return 0 (not equal)
    ret
wcscmp_ci_w endp

; ==============================================================================
; wcscat_w - Wide Character String Concatenate
;
; Purpose: Appends source string to destination string
;
; Parameters:
;   RCX = Destination buffer pointer
;   RDX = Source string pointer
;
; Returns: None (modifies destination buffer)
; ==============================================================================
wcscat_w proc
    push rsi
    push rdi
    mov rdi, rcx                ; RDI = destination
@@:
    cmp word ptr [rdi], 0       ; Find end of destination
    je @F
    add rdi, 2
    jmp @B
@@:
    mov rsi, rdx                ; RSI = source
@@:
    mov ax, word ptr [rsi]      ; Copy character
    mov word ptr [rdi], ax
    test ax, ax                 ; End of source?
    jz @F
    add rsi, 2
    add rdi, 2
    jmp @B
@@:
    pop rdi
    pop rsi
    ret
wcscat_w endp

; ==============================================================================
; RunCommand - Execute Command as TrustedInstaller
;
; Purpose: Retrieves command text from the ComboBox, validates it, and executes
;          it using TrustedInstaller privileges. Updates status display and
;          saves to MRU list on success.
;
; Parameters:
;   RCX = Parent window handle (for message boxes)
;
; Returns:
;   RAX = 1 on success, 0 on failure or empty command
;
; Process:
;   1. Get command text from ComboBox (g_hwndEdit)
;   2. Validate command is not empty
;   3. Update status to "Launching..."
;   4. Call RunAsTrustedInstaller with new console flag (1)
;   5. On success: save to MRU list and show "Process OK"
;   6. On failure: show "Failed" status
;   7. On empty: show error message box
; ==============================================================================
RunCommand proc frame
    push rbx
    .pushreg rbx
    push rsi
    .pushreg rsi
    push rdi
    .pushreg rdi
    sub rsp, 48
    .allocstack 48
    .endprolog

    mov rbx, rcx                ; Save window handle

    ; Get text from ComboBox edit control
    mov r8d, 520                ; Buffer size
    lea rdx, g_cmdBuf           ; Buffer pointer
    mov rcx, g_hwndEdit         ; ComboBox handle
    sub rsp, 32
    call GetWindowTextW
    add rsp, 32
    test eax, eax
    jz rc_empty                 ; No text entered

    ; Update status to show launching
    lea rdx, str_StatusRunning
    mov rcx, g_hwndStatus
    sub rsp, 32
    call SetWindowTextW
    add rsp, 32

    ; Execute command with TrustedInstaller privileges
    mov edx, 1                  ; new console window flag
    lea rcx, g_cmdBuf           ; Command to execute
    sub rsp, 32
    call RunAsTrustedInstaller
    add rsp, 32
    test eax, eax
    jz rc_fail                  ; Execution failed

    ; Success - save to MRU list
    sub rsp, 32
    call SaveMRU
    add rsp, 32

    ; Update status to success
    lea rdx, str_StatusOK
    mov rcx, g_hwndStatus
    sub rsp, 32
    call SetWindowTextW
    add rsp, 32

    mov eax, 1                  ; Return success
    jmp rc_done

rc_fail:
    ; Execution failed - update status
    lea rdx, str_StatusFail
    mov rcx, g_hwndStatus
    sub rsp, 32
    call SetWindowTextW
    add rsp, 32
    xor eax, eax                ; Return failure
    jmp rc_done

rc_empty:
    ; No command entered - show error message
    mov r9d, MB_ICONERROR
    lea r8, str_TitleErr
    lea rdx, str_ErrEmpty
    mov rcx, rbx
    sub rsp, 32
    call MessageBoxW
    add rsp, 32
    xor eax, eax                ; Return failure

rc_done:
    add rsp, 48
    pop rdi
    pop rsi
    pop rbx
    ret
RunCommand endp


; ==============================================================================
; WndProc - Main Window Procedure
;
; Purpose: Processes all Windows messages for the main application window.
;          Handles user interactions, control creation, resizing, and events.
;
; Messages handled:
;   WM_CREATE     - Create child controls (ComboBox, buttons, status label)
;   WM_COMMAND    - Process menu selections and button clicks
;   WM_CHAR       - Handle Enter key to execute command
;   WM_SIZE       - Resize and reposition controls dynamically
;   WM_SETTINGCHANGE - React to system theme changes
;   WM_THEMECHANGED  - React to theme changes
;   WM_DESTROY    - Clean up and quit application
;   WM_DROPFILES  - Handle drag-and-drop file operations
;   WM_KEYDOWN    - Handle ESC key to exit
;
; Parameters:
;   RCX = hWnd (window handle)
;   EDX = uMsg (message identifier)
;   R8  = wParam (first message parameter)
;   R9  = lParam (second message parameter)
;
; Returns:
;   RAX = Message-specific return value (0 if processed, DefWindowProc if not)
;
; Stack frame: 312 bytes for local variables and message parameters
; ==============================================================================
WndProc proc frame
    push rbp
    .pushreg rbp
    mov rbp, rsp
    .setframe rbp, 0
    push rbx
    .pushreg rbx
    push rsi
    .pushreg rsi
    push rdi
    .pushreg rdi
    push r12
    .pushreg r12
    push r13
    .pushreg r13
    push r14
    .pushreg r14
    push r15
    .pushreg r15
    sub rsp, 312
    .allocstack 312
    .endprolog

    ; Save message parameters
    mov [rbp-72], rcx           ; hWnd
    mov [rbp-80], edx           ; uMsg
    mov [rbp-88], r8            ; wParam
    mov [rbp-96], r9            ; lParam

    ; Dispatch message to appropriate handler
    mov eax, edx
    cmp eax, WM_CREATE
    je wp_create
    cmp eax, WM_COMMAND
    je wp_command
    cmp eax, WM_CHAR
    je wp_char
    cmp eax, WM_SIZE
    je wp_size
    cmp eax, WM_SETTINGCHANGE
    je wp_themechange
    cmp eax, WM_THEMECHANGED
    je wp_themechange
    cmp eax, WM_DESTROY
    je wp_destroy
    cmp eax, WM_DROPFILES
    je wp_dropfiles
    cmp eax, WM_KEYDOWN
    je wp_keydown
    jmp wp_defproc

wp_create:
    ; Extract hInstance from CREATESTRUCT
    mov rax, r9                 ; lParam = CREATESTRUCT pointer
    mov rax, [rax+8]            ; lpCreateParams->hInstance
    mov r15, rax                ; Save hInstance

    ; Create main menu
    sub rsp, 32
    call CreateMenu
    add rsp, 32
    mov r12, rax                ; Save menu handle

    ; Create File submenu
    sub rsp, 32
    call CreatePopupMenu
    add rsp, 32
    mov r13, rax                ; Save popup menu handle

    ; Add "Browse" menu item
    lea r9, str_MenuBrowse
    mov r8d, ID_FILE_BROWSE
    mov edx, MF_STRING
    mov rcx, r13
    sub rsp, 32
    call AppendMenuW
    add rsp, 32

    ; Add separator
    xor r9d, r9d
    xor r8d, r8d
    mov edx, MF_SEPARATOR
    mov rcx, r13
    sub rsp, 32
    call AppendMenuW
    add rsp, 32

    ; Add "About" menu item
    lea r9, str_MenuAbout
    mov r8d, ID_FILE_ABOUT
    mov edx, MF_STRING
    mov rcx, r13
    sub rsp, 32
    call AppendMenuW
    add rsp, 32

    ; Add separator
    xor r9d, r9d
    xor r8d, r8d
    mov edx, MF_SEPARATOR
    mov rcx, r13
    sub rsp, 32
    call AppendMenuW
    add rsp, 32

    ; Add "Exit" menu item
    lea r9, str_MenuExit
    mov r8d, ID_FILE_EXIT
    mov edx, MF_STRING
    mov rcx, r13
    sub rsp, 32
    call AppendMenuW
    add rsp, 32

    ; Add File submenu to main menu
    lea r9, str_MenuFile
    mov r8, r13                 ; Popup menu handle
    mov edx, MF_POPUP
    mov rcx, r12                ; Main menu handle
    sub rsp, 32
    call AppendMenuW
    add rsp, 32

    ; Attach menu to window
    mov rdx, r12
    mov rcx, [rbp-72]           ; hWnd
    sub rsp, 32
    call SetMenu
    add rsp, 32

    ; Create ComboBox control
    sub rsp, 96
    xor rax, rax
    mov [rsp+88], rax           ; lpParam
    mov [rsp+80], r15           ; hInstance
    mov rax, IDC_COMBO
    mov [rsp+72], rax           ; hMenu (control ID)
    mov rax, [rbp-72]
    mov [rsp+64], rax           ; hWndParent
    mov dword ptr [rsp+56], 200 ; nHeight
    mov dword ptr [rsp+48], 460 ; nWidth
    mov dword ptr [rsp+40], 10  ; Y
    mov dword ptr [rsp+32], 10  ; X
    mov r9d, STY_COMBO          ; dwStyle
    xor r8d, r8d                ; lpWindowName
    lea rdx, str_ComboBox       ; lpClassName
    mov ecx, WS_EX_CLIENTEDGE   ; dwExStyle
    call CreateWindowExW
    add rsp, 96
    test rax, rax
    jz wp_create_fail
    mov g_hwndEdit, rax         ; Save ComboBox handle

    ; Load MRU list into ComboBox
    call LoadMRU

    ; Get hInstance again
    mov rax, [rbp-96]
    mov rax, [rax+8]
    mov r15, rax

    ; Create Browse button
    sub rsp, 96
    xor rax, rax
    mov [rsp+88], rax           ; lpParam
    mov [rsp+80], r15           ; hInstance
    mov rax, IDC_BTN_BROWSE
    mov [rsp+72], rax           ; hMenu (control ID)
    mov rax, [rbp-72]
    mov [rsp+64], rax           ; hWndParent
    mov dword ptr [rsp+56], 25  ; nHeight
    mov dword ptr [rsp+48], 80  ; nWidth
    mov dword ptr [rsp+40], 45  ; Y
    mov dword ptr [rsp+32], 10  ; X
    mov r9d, STY_BUTTON         ; dwStyle
    lea r8, str_BtnBrowse       ; lpWindowName
    lea rdx, str_Button         ; lpClassName
    xor ecx, ecx                ; dwExStyle
    call CreateWindowExW
    add rsp, 96
    test rax, rax
    jz wp_create_fail

    ; Get hInstance again
    mov rax, [rbp-96]
    mov rax, [rax+8]
    mov r15, rax

    ; Create Run button (default button)
    sub rsp, 96
    xor rax, rax
    mov [rsp+88], rax           ; lpParam
    mov [rsp+80], r15           ; hInstance
    mov rax, IDC_BTN_RUN
    mov [rsp+72], rax           ; hMenu (control ID)
    mov rax, [rbp-72]
    mov [rsp+64], rax           ; hWndParent
    mov dword ptr [rsp+56], 25  ; nHeight
    mov dword ptr [rsp+48], 80  ; nWidth
    mov dword ptr [rsp+40], 45  ; Y
    mov dword ptr [rsp+32], 400 ; X
    mov r9d, STY_BUTTON_DEF     ; dwStyle (default button)
    lea r8, str_BtnRun          ; lpWindowName
    lea rdx, str_Button         ; lpClassName
    xor ecx, ecx                ; dwExStyle
    call CreateWindowExW
    add rsp, 96
    test rax, rax
    jz wp_create_fail
    mov g_hwndBtn, rax          ; Save Run button handle

    ; Get hInstance again
    mov rax, [rbp-96]
    mov rax, [rax+8]
    mov r15, rax

    ; Create status label
    sub rsp, 96
    xor rax, rax
    mov [rsp+88], rax           ; lpParam
    mov [rsp+80], r15           ; hInstance
    mov rax, IDC_STATIC_STATUS
    mov [rsp+72], rax           ; hMenu (control ID)
    mov rax, [rbp-72]
    mov [rsp+64], rax           ; hWndParent
    mov dword ptr [rsp+56], 20  ; nHeight
    mov dword ptr [rsp+48], 200 ; nWidth
    mov dword ptr [rsp+40], 48  ; Y
    mov dword ptr [rsp+32], 100 ; X
    mov r9d, STY_STATIC         ; dwStyle
    lea r8, str_StatusInit      ; lpWindowName (initial text)
    lea rdx, str_Static         ; lpClassName
    xor ecx, ecx                ; dwExStyle
    call CreateWindowExW
    add rsp, 96
    test rax, rax
    jz wp_create_fail
    mov g_hwndStatus, rax       ; Save status label handle

    ; Set focus to ComboBox
    mov rcx, g_hwndEdit
    sub rsp, 32
    call SetFocus
    add rsp, 32

    ; Enable drag-and-drop
    mov edx, 1                  ; Accept files
    mov rcx, [rbp-72]
    sub rsp, 32
    call DragAcceptFiles
    add rsp, 32

    ; Allow WM_DROPFILES through UIPI
    xor r9d, r9d
    mov r8d, MSGFLT_ALLOW
    mov edx, WM_DROPFILES
    mov rcx, [rbp-72]
    sub rsp, 32
    call ChangeWindowMessageFilterEx
    add rsp, 32

    ; Allow WM_COPYGLOBALDATA through UIPI
    xor r9d, r9d
    mov r8d, MSGFLT_ALLOW
    mov edx, WM_COPYGLOBALDATA
    mov rcx, [rbp-72]
    sub rsp, 32
    call ChangeWindowMessageFilterEx
    add rsp, 32

    xor eax, eax                ; Return 0 (success)
    jmp wp_done

wp_create_fail:
    mov rax, -1                 ; Return -1 (failure)
    jmp wp_done

wp_command:
    ; Extract notification code and control ID
    mov rax, [rbp-88]           ; wParam
    shr rax, 16                 ; High word = notification code
    mov r14d, eax
    mov rax, [rbp-88]
    and eax, 0FFFFh             ; Low word = control ID
    mov r15d, eax

    ; Check which control/menu sent the message
    cmp r15d, IDC_EDIT
    je wp_check_edit
    cmp r15d, IDC_BTN_RUN
    je wp_cmd_run
    cmp r15d, 1                 ; IDOK
    je wp_cmd_run
    cmp r15d, IDC_BTN_BROWSE
    je wp_cmd_browse
    cmp r15d, ID_FILE_BROWSE
    je wp_cmd_browse
    cmp r15d, ID_FILE_EXIT
    je wp_cmd_exit
    cmp r15d, ID_FILE_ABOUT
    je wp_cmd_about
    jmp wp_defproc

wp_check_edit:
    ; Check edit control notification
    cmp r14d, EN_SETFOCUS
    je wp_edit_focus
    jmp wp_defproc

wp_edit_focus:
    ; Edit control got focus - make Run button default
    mov r9d, 1                  ; Redraw
    mov r8d, BS_DEFPUSHBUTTON
    mov edx, BM_SETSTYLE
    mov rcx, g_hwndBtn
    sub rsp, 32
    call SendMessageW
    add rsp, 32
    xor eax, eax
    jmp wp_done

wp_char:
    ; Check for Enter key
    mov rax, [rbp-88]
    cmp eax, 13                 ; VK_RETURN
    je wp_char_enter
    jmp wp_defproc

wp_char_enter:
    ; Enter pressed - check if ComboBox has focus
    sub rsp, 32
    call GetFocus
    add rsp, 32
    cmp rax, g_hwndEdit
    jne wp_defproc
    ; Execute command
    mov rcx, [rbp-72]
    call RunCommand
    xor eax, eax
    jmp wp_done

wp_cmd_run:
    ; Run button clicked or menu selected
    mov rcx, [rbp-72]
    call RunCommand
    jmp wp_done

wp_cmd_browse:
    ; Browse button clicked - show file open dialog
    ; Initialize OPENFILENAME structure
    lea rdi, [rsp+40]
    xor rax, rax
    mov rcx, 19                 ; 19 QWORDs to zero
@@:
    test rcx, rcx
    jz @F
    mov qword ptr [rdi], rax
    add rdi, 8
    dec rcx
    jmp @B
@@:

    ; Fill in OPENFILENAME fields
    mov dword ptr [rsp+40], 152         ; lStructSize
    mov rax, [rbp-72]
    mov [rsp+40+8], rax                 ; hwndOwner
    lea rax, str_Filter
    mov [rsp+40+24], rax                ; lpstrFilter
    lea rax, g_filePath
    mov [rsp+40+48], rax                ; lpstrFile
    mov dword ptr [rsp+40+56], 520      ; nMaxFile
    lea rax, str_DefPath
    mov [rsp+40+80], rax                ; lpstrInitialDir
    mov dword ptr [rsp+40+96], OFN_FILEMUSTEXIST or OFN_PATHMUSTEXIST ; Flags

    ; Clear file path buffer
    mov word ptr g_filePath, 0

    ; Show dialog
    lea rcx, [rsp+40]
    sub rsp, 32
    call GetOpenFileNameW
    add rsp, 32
    test eax, eax
    jz wp_browse_cancel         ; User cancelled

    ; Check if file is a .lnk shortcut
    lea rcx, g_filePath
    call wcslen_w
    mov r12, rax                ; Save length
    cmp r12, 4
    jl wp_browse_set            ; Too short for extension

    ; Compare last 4 characters with ".lnk"
    lea rsi, g_filePath
    lea rsi, [rsi + r12*2 - 8]  ; Point to last 4 chars
    lea rdx, str_extLnk
    mov rcx, rsi
    call wcscmp_ci_w
    test rax, rax
    jnz wp_browse_resolve_lnk   ; Is .lnk file
    jmp wp_browse_set

wp_browse_resolve_lnk:
    ; Resolve .lnk file to target
    lea r8, g_tempBuf           ; Arguments buffer
    lea rdx, g_cmdBuf           ; Target path buffer
    lea rcx, g_filePath         ; .lnk file path
    call ResolveLnkPath
    test rax, rax
    jz wp_browse_set            ; Resolution failed

    ; Check if target path is empty
    lea rcx, g_cmdBuf
    call wcslen_w
    test rax, rax
    jz wp_browse_args_only

    ; Append space after target path
    lea rsi, g_cmdBuf
    lea rsi, [rsi + rax*2]
    mov word ptr [rsi], ' '
    add rsi, 2
    mov word ptr [rsi], 0

wp_browse_args_only:
    ; Concatenate arguments
    lea rdx, g_tempBuf
    lea rcx, g_cmdBuf
    call wcscat_w

    ; Set ComboBox text to resolved command
    lea rdx, g_cmdBuf
    mov rcx, g_hwndEdit
    sub rsp, 32
    call SetWindowTextW
    add rsp, 32
    jmp wp_browse_cancel

wp_browse_set:
    ; Set ComboBox text to file path
    lea rdx, g_filePath
    mov rcx, g_hwndEdit
    sub rsp, 32
    call SetWindowTextW
    add rsp, 32

wp_browse_cancel:
    ; Return focus to ComboBox
    mov rcx, g_hwndEdit
    sub rsp, 32
    call SetFocus
    add rsp, 32

    ; Make Run button default
    mov r9d, 1
    mov r8d, BS_DEFPUSHBUTTON
    mov edx, BM_SETSTYLE
    mov rcx, g_hwndBtn
    sub rsp, 32
    call SendMessageW
    add rsp, 32

    xor eax, eax
    jmp wp_done

wp_cmd_about:
    ; Show About message box
    mov r9d, MB_ICONINFORMATION
    lea r8, str_AboutTitle
    lea rdx, str_AboutText
    mov rcx, [rbp-72]
    sub rsp, 32
    call MessageBoxW
    add rsp, 32
    xor eax, eax
    jmp wp_done

wp_cmd_exit:
    ; Exit application
    xor ecx, ecx
    sub rsp, 32
    call PostQuitMessage
    add rsp, 32
    xor eax, eax
    jmp wp_done

wp_size:
    ; Get client area dimensions
    lea rdx, [rsp+40]           ; RECT buffer
    mov rcx, [rbp-72]
    sub rsp, 32
    call GetClientRect
    add rsp, 32

    ; Extract width and height
    mov esi, dword ptr [rsp+40+8]   ; Width
    mov edi, dword ptr [rsp+40+12]  ; Height

    ; Calculate ComboBox width (client width - 20)
    mov r12d, esi
    sub r12d, 20

    ; Resize ComboBox
    sub rsp, 48
    mov dword ptr [rsp+40], 1   ; bRepaint
    mov qword ptr [rsp+32], 25  ; Height
    mov r9d, r12d               ; Width
    mov r8d, 10                 ; Y
    mov edx, 10                 ; X
    mov rcx, g_hwndEdit
    call MoveWindow
    add rsp, 48

    ; Get Browse button handle
    mov edx, IDC_BTN_BROWSE
    mov rcx, [rbp-72]
    sub rsp, 32
    call GetDlgItem
    add rsp, 32
    mov r14, rax

    ; Reposition Browse button
    sub rsp, 48
    mov dword ptr [rsp+40], 1   ; bRepaint
    mov qword ptr [rsp+32], 25  ; Height
    mov r9d, 80                 ; Width
    mov r8d, 45                 ; Y
    mov edx, 10                 ; X
    mov rcx, r14
    call MoveWindow
    add rsp, 48

    ; Calculate Run button X position
    mov r13d, r12d
    sub r13d, 70

    ; Reposition Run button
    sub rsp, 48
    mov dword ptr [rsp+40], 1   ; bRepaint
    mov qword ptr [rsp+32], 25  ; Height
    mov r9d, 80                 ; Width
    mov r8d, 45                 ; Y
    mov edx, r13d               ; X
    mov rcx, g_hwndBtn
    call MoveWindow
    add rsp, 48

    ; Reposition status label
    sub rsp, 48
    mov dword ptr [rsp+40], 1   ; bRepaint
    mov qword ptr [rsp+32], 20  ; Height
    mov r9d, 200                ; Width (fixed to not overlap Run)
    mov r8d, 48                 ; Y (aligned with buttons)
    mov edx, 100                ; X (after Browse button)
    mov rcx, g_hwndStatus
    call MoveWindow
    add rsp, 48

    xor eax, eax
    jmp wp_done

wp_dropfiles:
    ; Get dropped file path
    mov r9d, 520                ; Buffer size
    lea r8, g_filePath          ; Buffer
    xor edx, edx                ; File index (first file)
    mov rcx, [rbp-88]           ; HDROP handle
    sub rsp, 32
    call DragQueryFileW
    add rsp, 32

    ; Finish drag operation
    mov rcx, [rbp-88]
    sub rsp, 32
    call DragFinish
    add rsp, 32

    ; Check if file is a .lnk shortcut
    lea rcx, g_filePath
    call wcslen_w
    mov r12, rax
    cmp r12, 4
    jl wp_drop_set              ; Too short for extension

    ; Compare extension
    lea rsi, g_filePath
    lea rsi, [rsi + r12*2 - 8]
    lea rdx, str_extLnk
    mov rcx, rsi
    call wcscmp_ci_w
    test rax, rax
    jnz wp_drop_resolve_lnk
    jmp wp_drop_set

wp_drop_resolve_lnk:
    ; Resolve .lnk to target
    lea r8, g_tempBuf
    lea rdx, g_cmdBuf
    lea rcx, g_filePath
    call ResolveLnkPath
    test rax, rax
    jz wp_drop_set

    ; Check if target path is empty
    lea rcx, g_cmdBuf
    call wcslen_w
    test rax, rax
    jz wp_drop_args_only

    ; Append space
    lea rsi, g_cmdBuf
    lea rsi, [rsi + rax*2]
    mov word ptr [rsi], ' '
    add rsi, 2
    mov word ptr [rsi], 0

wp_drop_args_only:
    ; Concatenate arguments
    lea rdx, g_tempBuf
    lea rcx, g_cmdBuf
    call wcscat_w

    ; Set ComboBox text
    lea rdx, g_cmdBuf
    mov rcx, g_hwndEdit
    sub rsp, 32
    call SetWindowTextW
    add rsp, 32
    jmp wp_drop_run

wp_drop_set:
    ; Set ComboBox text to file path
    lea rdx, g_filePath
    mov rcx, g_hwndEdit
    sub rsp, 32
    call SetWindowTextW
    add rsp, 32

wp_drop_run:
    ; Execute dropped command
    mov rcx, [rbp-72]
    call RunCommand
    xor eax, eax
    jmp wp_done

wp_keydown:
    ; Check for ESC key
    mov rax, [rbp-88]
    cmp eax, VK_ESCAPE
    jne wp_defproc
    ; Exit on ESC
    xor ecx, ecx
    sub rsp, 32
    call PostQuitMessage
    add rsp, 32
    xor eax, eax
    jmp wp_done

wp_themechange:
    ; Re-apply theme attributes when system settings change
    mov rcx, [rbp-72]
    sub rsp, 32
    call ApplyWindowTheme
    add rsp, 32
    xor eax, eax
    jmp wp_done

wp_destroy:
    ; Window is being destroyed
    xor ecx, ecx
    sub rsp, 32
    call PostQuitMessage
    add rsp, 32
    xor eax, eax
    jmp wp_done

wp_defproc:
    ; Pass unhandled messages to default handler
    mov r9, [rbp-96]
    mov r8, [rbp-88]
    mov edx, [rbp-80]
    mov rcx, [rbp-72]
    sub rsp, 32
    call DefWindowProcW
    add rsp, 32

wp_done:
    add rsp, 312
    pop r15
    pop r14
    pop r13
    pop r12
    pop rdi
    pop rsi
    pop rbx
    pop rbp
    ret
WndProc endp

; ==============================================================================
; ApplyWindowTheme - Apply DWM theme and backdrop based on system preference
;
; Purpose: Reads the system app theme preference and applies:
;          - DWMWA_USE_IMMERSIVE_DARK_MODE
;          - DWMWA_SYSTEMBACKDROP_TYPE (Mica)
;
; Parameters:
;   RCX = hWnd
;
; Returns:
;   EAX = 0
;
; Stack frame: 80 bytes for registry handles and values
; ==============================================================================
ApplyWindowTheme proc frame
    push rbx
    .pushreg rbx
    sub rsp, 80
    .allocstack 80
    .endprolog

    ; Local variables:
    ; [rsp+32] = hKey (8 bytes)
    ; [rsp+40] = valueData (4 bytes)
    ; [rsp+48] = dataSize (4 bytes)
    ; [rsp+56] = darkFlag (4 bytes)

    mov rbx, rcx                ; Save hWnd
    mov dword ptr [rsp+56], 0   ; Default to light theme

    ; Open theme registry key
    sub rsp, 48
    lea rax, [rsp+48+32]
    mov [rsp+32], rax           ; phkResult
    mov r9d, KEY_READ
    xor r8d, r8d                ; ulOptions
    lea rdx, str_regThemeKey
    mov ecx, HKEY_CURRENT_USER
    call RegOpenKeyExW
    add rsp, 48
    test eax, eax
    jnz awt_apply

    ; Query AppsUseLightTheme (DWORD)
    mov dword ptr [rsp+48], 4   ; dataSize = sizeof(DWORD)
    mov dword ptr [rsp+40], 1   ; Default to light if read fails
    sub rsp, 48
    lea rax, [rsp+48+48]
    mov [rsp+40], rax           ; lpcbData
    lea rax, [rsp+48+40]
    mov [rsp+32], rax           ; lpData
    xor r9d, r9d                ; lpType (NULL)
    xor r8d, r8d                ; lpReserved (NULL)
    lea rdx, str_regAppsUseLightTheme
    mov rcx, [rsp+48+32]        ; hKey
    call RegQueryValueExW
    add rsp, 48
    test eax, eax
    jnz awt_close

    ; If AppsUseLightTheme == 0, enable dark mode
    mov eax, dword ptr [rsp+40]
    test eax, eax
    jnz awt_close
    mov dword ptr [rsp+56], 1

awt_close:
    sub rsp, 32
    mov rcx, [rsp+32]
    call RegCloseKey
    add rsp, 32

awt_apply:
    ; Apply immersive dark mode attribute
    mov eax, dword ptr [rsp+56]
    mov dword ptr [rsp+32], eax
    mov r9d, 4
    lea r8, [rsp+32]
    mov edx, DWMWA_USE_IMMERSIVE_DARK_MODE
    mov rcx, rbx
    sub rsp, 32
    call DwmSetWindowAttribute
    add rsp, 32
    test eax, eax
    jns awt_backdrop

    ; Fallback for older Windows builds
    mov eax, dword ptr [rsp+56]
    mov dword ptr [rsp+32], eax
    mov r9d, 4
    lea r8, [rsp+32]
    mov edx, DWMWA_USE_IMMERSIVE_DARK_MODE_OLD
    mov rcx, rbx
    sub rsp, 32
    call DwmSetWindowAttribute
    add rsp, 32

awt_backdrop:
    ; Apply Mica backdrop (follows light/dark preference)
    mov dword ptr [rsp+32], DWMSBT_MAINWINDOW
    mov r9d, 4
    lea r8, [rsp+32]
    mov edx, DWMWA_SYSTEMBACKDROP_TYPE
    mov rcx, rbx
    sub rsp, 32
    call DwmSetWindowAttribute
    add rsp, 32

    xor eax, eax
    add rsp, 80
    pop rbx
    ret
ApplyWindowTheme endp

; ==============================================================================
; LoadMRU - Load Most Recently Used Commands from Registry
;
; Purpose: Reads up to 5 most recently used commands from the Windows registry
;          and populates the ComboBox dropdown list.
;
; Registry location: HKEY_CURRENT_USER\Software\cmdt
; Value names: "0", "1", "2", "3", "4" (0 = most recent)
;
; Parameters: None
;
; Returns: None
;
; Process:
;   1. Open registry key (HKEY_CURRENT_USER\Software\cmdt)
;   2. Loop through values "0" through "4"
;   3. For each value, read command string
;   4. Add command to ComboBox using CB_ADDSTRING
;   5. Set ComboBox selection to first item (most recent)
;   6. Close registry key
;
; Stack frame: 104 bytes for registry handles and buffers
; ==============================================================================
LoadMRU proc frame
    push rbx
    .pushreg rbx
    push rsi
    .pushreg rsi
    push rdi
    .pushreg rdi
    push r12
    .pushreg r12
    sub rsp, 104
    .allocstack 104
    .endprolog

    ; Local variables:
    ; [rsp+32] = hKey (8 bytes)
    ; [rsp+40] = valName (8 bytes)
    ; [rsp+48] = valDataLen (4 bytes)

    ; Open registry key
    sub rsp, 48
    lea rax, [rsp+48+32]
    mov [rsp+32], rax           ; phkResult
    mov r9d, KEY_READ
    xor r8d, r8d                ; ulOptions
    lea rdx, str_regKey
    mov ecx, HKEY_CURRENT_USER
    call RegOpenKeyExW
    add rsp, 48
    test eax, eax
    jnz lm_done                 ; Key doesn't exist

    ; Loop through MRU values (0-4)
    xor r12d, r12d              ; Counter

lm_loop:
    cmp r12d, MRU_MAX_ITEMS     ; Check if done
    jge lm_setsel

    ; Build value name ("0", "1", etc.)
    mov word ptr [rsp+40], '0'
    add word ptr [rsp+40], r12w
    mov word ptr [rsp+42], 0    ; Null terminator

    ; Reset data length
    mov dword ptr [rsp+48], 1040

    ; Query registry value
    sub rsp, 48
    lea rax, [rsp+48+48]
    mov [rsp+40], rax           ; lpcbData
    lea rax, g_tempBuf
    mov [rsp+32], rax           ; lpData
    xor r9d, r9d                ; lpType (NULL)
    xor r8d, r8d                ; lpReserved (NULL)
    lea rdx, [rsp+48+40]        ; lpValueName
    mov rcx, [rsp+48+32]        ; hKey
    call RegQueryValueExW
    add rsp, 48
    test eax, eax
    jnz lm_setsel               ; No more values

    ; Add to ComboBox
    lea r9, g_tempBuf
    xor r8d, r8d
    mov edx, CB_ADDSTRING
    mov rcx, g_hwndEdit
    sub rsp, 32
    call SendMessageW
    add rsp, 32

    inc r12d
    jmp lm_loop

lm_setsel:
    ; Select first item in ComboBox
    xor r9d, r9d                ; Index 0
    xor r8d, r8d
    mov edx, CB_SETCURSEL
    mov rcx, g_hwndEdit
    sub rsp, 32
    call SendMessageW
    add rsp, 32

    ; Close registry key
    mov rcx, [rsp+32]
    sub rsp, 32
    call RegCloseKey
    add rsp, 32

lm_done:
    add rsp, 104
    pop r12
    pop rdi
    pop rsi
    pop rbx
    ret
LoadMRU endp

; ==============================================================================
; SaveMRU - Save Command to Most Recently Used List
;
; Purpose: Saves the current command to the MRU registry list. Shifts existing
;          entries down and adds new command at position 0. Maintains maximum
;          of 5 entries.
;
; Registry location: HKEY_CURRENT_USER\Software\cmdt
; Value names: "0" (newest) through "4" (oldest)
;
; Parameters: None (reads from g_hwndEdit ComboBox)
;
; Returns: None
;
; Process:
;   1. Get command text from ComboBox
;   2. Open/create registry key
;   3. Delete value "4" (oldest)
;   4. Shift values: "3"→"4", "2"→"3", "1"→"2", "0"→"1"
;   5. Write new command as value "0"
;   6. Add to ComboBox at position 0
;   7. Limit ComboBox to 5 items
;   8. Select first item
;   9. Close registry key
;
; Stack frame: 184 bytes for registry operations
; ==============================================================================
SaveMRU proc frame
    push rbx
    .pushreg rbx
    push rsi
    .pushreg rsi
    push rdi
    .pushreg rdi
    push r12
    .pushreg r12
    push r13
    .pushreg r13
    push r14
    .pushreg r14
    sub rsp, 184
    .allocstack 184
    .endprolog

    ; Local variables:
    ; [rsp+32] = hKey (8 bytes)
    ; [rsp+40] = disp (4 bytes)
    ; [rsp+48] = valName (8 bytes)
    ; [rsp+56] = textLen (4 bytes)

    ; Get text from ComboBox
    mov r8d, 520                ; Buffer size
    lea rdx, g_cmdBuf
    mov rcx, g_hwndEdit
    sub rsp, 32
    call GetWindowTextW
    add rsp, 32
    test eax, eax
    jz sm_done                  ; Empty text, don't save

    ; Calculate byte length (chars * 2 + 2 for null)
    shl eax, 1
    add eax, 2
    mov [rsp+56], eax           ; Save text length in bytes

    ; Create or open registry key
    sub rsp, 80
    lea rax, [rsp+80+40]
    mov [rsp+64], rax           ; lpdwDisposition
    lea rax, [rsp+80+32]
    mov [rsp+56], rax           ; phkResult
    mov qword ptr [rsp+48], 0   ; lpSecurityAttributes
    mov dword ptr [rsp+40], KEY_WRITE or KEY_READ
    mov qword ptr [rsp+32], 0   ; lpClass
    xor r9d, r9d                ; dwOptions
    xor r8d, r8d                ; Reserved
    lea rdx, str_regKey
    mov ecx, HKEY_CURRENT_USER
    call RegCreateKeyExW
    add rsp, 80
    test eax, eax
    jnz sm_done                 ; Failed to create/open key

    ; Delete oldest value "4" to make room
    mov word ptr [rsp+48], '4'
    mov word ptr [rsp+50], 0
    lea rdx, [rsp+48]
    mov rcx, [rsp+32]
    sub rsp, 32
    call RegDeleteValueW
    add rsp, 32                 ; Ignore error if doesn't exist

    ; Shift values down: 3→4, 2→3, 1→2, 0→1
    mov r12d, 3                 ; Start from "3"
sm_shift_loop:
    cmp r12d, 0
    jl sm_write_new             ; Done shifting

    ; Build source value name
    mov word ptr [rsp+48], '0'
    add word ptr [rsp+48], r12w
    mov word ptr [rsp+50], 0

    ; Read value at index r12
    mov dword ptr [rsp+60], 1040
    sub rsp, 48
    lea rax, [rsp+48+60]
    mov [rsp+32], rax           ; lpcbData
    lea r9, g_tempBuf           ; lpData
    xor r8d, r8d                ; lpType (NULL)
    xor edx, edx                ; lpReserved (NULL)
    lea rcx, [rsp+48+48]        ; lpValueName
    mov rbx, [rsp+48+32]        ; hKey
    mov rcx, rbx
    lea rdx, [rsp+48+48]
    xor r8d, r8d
    lea r9, g_tempBuf
    lea rax, [rsp+48+60]
    mov [rsp+32], rax
    call RegQueryValueExW
    add rsp, 48
    test eax, eax
    jnz sm_shift_next           ; Value doesn't exist, skip

    ; Build destination value name (r12+1)
    mov word ptr [rsp+48], '1'
    add word ptr [rsp+48], r12w
    mov word ptr [rsp+50], 0

    ; Write to index r12+1
    mov eax, [rsp+60]           ; Data length
    sub rsp, 48
    mov [rsp+40], eax           ; cbData
    lea rax, g_tempBuf
    mov [rsp+32], rax           ; lpData
    mov r9d, REG_SZ             ; Type
    xor r8d, r8d                ; Reserved
    lea rdx, [rsp+48+48]        ; lpValueName
    mov rcx, [rsp+48+32]        ; hKey
    call RegSetValueExW
    add rsp, 48

sm_shift_next:
    dec r12d
    jmp sm_shift_loop

sm_write_new:
    ; Write new value as "0"
    mov word ptr [rsp+48], '0'
    mov word ptr [rsp+50], 0

    mov eax, [rsp+56]           ; Text length
    sub rsp, 48
    mov [rsp+40], eax           ; cbData
    lea rax, g_cmdBuf
    mov [rsp+32], rax           ; lpData
    mov r9d, REG_SZ             ; Type
    xor r8d, r8d                ; Reserved
    lea rdx, [rsp+48+48]        ; lpValueName
    mov rcx, [rsp+48+32]        ; hKey
    call RegSetValueExW
    add rsp, 48

    ; Insert at position 0 in ComboBox
    lea r9, g_cmdBuf
    xor r8d, r8d                ; Index 0
    mov edx, CB_INSERTSTRING
    mov rcx, g_hwndEdit
    sub rsp, 32
    call SendMessageW
    add rsp, 32

    ; Check ComboBox item count
    mov edx, CB_GETCOUNT
    xor r8d, r8d
    xor r9d, r9d
    mov rcx, g_hwndEdit
    sub rsp, 32
    call SendMessageW
    add rsp, 32
    cmp eax, MRU_MAX_ITEMS
    jle sm_setsel               ; Within limit

    ; Delete last item to maintain max 5
    mov r8d, MRU_MAX_ITEMS
    xor r9d, r9d
    mov edx, CB_DELETESTRING
    mov rcx, g_hwndEdit
    sub rsp, 32
    call SendMessageW
    add rsp, 32

sm_setsel:
    ; Select first item
    xor r9d, r9d                ; Index 0
    xor r8d, r8d
    mov edx, CB_SETCURSEL
    mov rcx, g_hwndEdit
    sub rsp, 32
    call SendMessageW
    add rsp, 32

    ; Close registry key
    mov rcx, [rsp+32]
    sub rsp, 32
    call RegCloseKey
    add rsp, 32

sm_done:
    add rsp, 184
    pop r14
    pop r13
    pop r12
    pop rdi
    pop rsi
    pop rbx
    ret
SaveMRU endp

; ==============================================================================
; ResolveLnkPath - Resolve Windows Shortcut (.lnk) File
;
; Purpose: Uses COM interfaces to resolve a .lnk shortcut file to its target
;          executable path and command-line arguments.
;
; COM Interfaces used:
;   - IShellLinkW: For manipulating Shell Link objects
;   - IPersistFile: For loading .lnk files from disk
;
; Parameters:
;   RCX = Pointer to .lnk file path (input)
;   RDX = Pointer to output buffer for target executable path
;   R8  = Pointer to output buffer for shortcut arguments
;
; Returns:
;   RAX = 1 on success, 0 on failure
;
; Output buffers:
;   - RDX buffer receives the target executable path (e.g., "C:\Windows\notepad.exe")
;   - R8 buffer receives arguments stored in the shortcut
;
; Process:
;   1. Initialize COM (CoInitialize)
;   2. Create IShellLink COM object (CoCreateInstance)
;   3. Query for IPersistFile interface
;   4. Load .lnk file (IPersistFile::Load)
;   5. Get target path (IShellLinkW::GetPath)
;   6. Get arguments (IShellLinkW::GetArguments)
;   7. Release COM interfaces
;   8. Uninitialize COM
;
; Stack frame: 112 bytes for COM interface pointers and return values
; ==============================================================================
ResolveLnkPath proc frame
    push rbx
    .pushreg rbx
    push rsi
    .pushreg rsi
    push rdi
    .pushreg rdi
    push r12
    .pushreg r12
    push r13
    .pushreg r13
    push r14
    .pushreg r14
    push r15
    .pushreg r15
    sub rsp, 112
    .allocstack 112
    .endprolog

    ; Save arguments
    mov r12, rcx                ; r12 = lpLnkPath
    mov r13, rdx                ; r13 = lpOutPath
    mov r14, r8                 ; r14 = lpArgsPath

    ; Local variables on stack:
    ; [rsp+32] = pShellLink (8 bytes)
    ; [rsp+40] = pPersistFile (8 bytes)
    ; [rsp+48] = hr (8 bytes)

    ; Initialize COM
    xor ecx, ecx                ; NULL parameter
    sub rsp, 32
    call CoInitialize
    add rsp, 32

    ; Create IShellLink COM object
    ; CoCreateInstance(&CLSID_ShellLink, NULL, CLSCTX_INPROC_SERVER, &IID_IShellLinkW, &pShellLink)
    lea rax, [rsp+32]           ; &pShellLink
    mov [rsp+32], rax           ; save for later - actually use as temp
    sub rsp, 48
    lea rax, [rsp+48+32]        ; &pShellLink (adjusted for sub rsp)
    mov [rsp+32], rax           ; 5th arg: ppv
    lea r9, IID_IShellLinkW     ; 4th arg: riid
    mov r8d, CLSCTX_INPROC_SERVER ; 3rd arg: dwClsContext
    xor edx, edx                ; 2nd arg: pUnkOuter = NULL
    lea rcx, CLSID_ShellLink    ; 1st arg: rclsid
    call CoCreateInstance
    add rsp, 48
    test eax, eax
    jnz rlp_fail_uninit         ; Failed to create

    ; Query for IPersistFile interface
    ; pShellLink->QueryInterface(&IID_IPersistFile, &pPersistFile)
    mov rbx, [rsp+32]           ; rbx = pShellLink
    mov rax, [rbx]              ; rax = vtable
    lea r8, [rsp+40]            ; 3rd arg: ppvObject = &pPersistFile
    lea rdx, IID_IPersistFile   ; 2nd arg: riid
    mov rcx, rbx                ; 1st arg: this
    sub rsp, 32
    call qword ptr [rax]        ; QueryInterface at vtable[0]
    add rsp, 32
    test eax, eax
    jnz rlp_release_link        ; Failed to query

    ; Load .lnk file
    ; pPersistFile->Load(lpLnkPath, 0)
    mov rbx, [rsp+40]           ; rbx = pPersistFile
    mov rax, [rbx]              ; rax = vtable
    xor r8d, r8d                ; 3rd arg: dwMode = 0
    mov rdx, r12                ; 2nd arg: pszFileName = lpLnkPath
    mov rcx, rbx                ; 1st arg: this
    sub rsp, 32
    call qword ptr [rax+40]     ; Load at vtable[5] = offset 40
    add rsp, 32
    test eax, eax
    jnz rlp_release_both        ; Failed to load

    ; Get target path
    ; pShellLink->GetPath(lpOutPath, 260, NULL, 0)
    mov rbx, [rsp+32]           ; rbx = pShellLink
    mov rax, [rbx]              ; rax = vtable
    sub rsp, 48
    mov qword ptr [rsp+32], 0   ; 5th arg: fFlags = 0
    xor r9d, r9d                ; 4th arg: pfd = NULL
    mov r8d, 260                ; 3rd arg: cch = MAX_PATH
    mov rdx, r13                ; 2nd arg: pszFile = lpOutPath
    mov rcx, rbx                ; 1st arg: this
    call qword ptr [rax+24]     ; GetPath at vtable[3] = offset 24
    add rsp, 48
    mov [rsp+48], eax           ; save hr

    ; Zero out arguments buffer
    mov rdi, r14
    xor eax, eax
    mov ecx, 260
    rep stosw

    ; Get shortcut arguments
    ; pShellLink->GetArguments(lpArgsPath, 520)
    mov rbx, [rsp+32]           ; rbx = pShellLink
    mov rax, [rbx]              ; rax = vtable
    mov r8d, 520                ; 3rd arg: cch
    mov rdx, r14                ; 2nd arg: pszArgs = lpArgsPath
    mov rcx, rbx                ; 1st arg: this
    sub rsp, 32
    call qword ptr [rax+80]     ; GetArguments at vtable[10] = offset 80
    add rsp, 32

    ; Release IPersistFile interface
    mov rbx, [rsp+40]           ; rbx = pPersistFile
    mov rax, [rbx]              ; rax = vtable
    mov rcx, rbx                ; 1st arg: this
    sub rsp, 32
    call qword ptr [rax+16]     ; Release at vtable[2] = offset 16
    add rsp, 32

    ; Release IShellLink interface
    mov rbx, [rsp+32]           ; rbx = pShellLink
    mov rax, [rbx]              ; rax = vtable
    mov rcx, rbx                ; 1st arg: this
    sub rsp, 32
    call qword ptr [rax+16]     ; Release at vtable[2] = offset 16
    add rsp, 32

    ; Uninitialize COM
    sub rsp, 32
    call CoUninitialize
    add rsp, 32

    ; Return based on GetPath result
    mov eax, [rsp+48]           ; hr from GetPath
    test eax, eax
    jnz rlp_fail                ; Failed
    mov eax, 1                  ; Success
    jmp rlp_done

rlp_release_both:
    ; Release both interfaces on error
    mov rbx, [rsp+40]
    mov rax, [rbx]
    mov rcx, rbx
    sub rsp, 32
    call qword ptr [rax+16]
    add rsp, 32

rlp_release_link:
    ; Release ShellLink interface
    mov rbx, [rsp+32]
    mov rax, [rbx]
    mov rcx, rbx
    sub rsp, 32
    call qword ptr [rax+16]
    add rsp, 32

rlp_fail_uninit:
    ; Uninitialize COM on error
    sub rsp, 32
    call CoUninitialize
    add rsp, 32

rlp_fail:
    xor eax, eax                ; Return failure

rlp_done:
    add rsp, 112
    pop r15
    pop r14
    pop r13
    pop r12
    pop rdi
    pop rsi
    pop rbx
    ret
ResolveLnkPath endp


; ==============================================================================
; CreateMainWindow - Create and Initialize Main Application Window
;
; Purpose: Registers the window class and creates the main application window.
;          Extracts application icon from shell32.dll.
;
; Parameters:
;   RCX = Application instance handle (hInstance)
;
; Returns:
;   RAX = Main window handle on success, 0 on failure
;
; Process:
;   1. Initialize WNDCLASSW structure
;   2. Set window procedure to WndProc
;   3. Extract icon #104 from shell32.dll (shield icon)
;   4. Set cursor to standard arrow
;   5. Set background to white brush
;   6. Register window class
;   7. Create overlapped window (420x150 pixels)
;   8. Show and update window
;   9. Return window handle
;
; Window properties:
;   - Class name: "TIRunnerClass"
;   - Title: "Run as TrustedInstaller"
;   - Size: 420x150 pixels at position (100, 100)
;   - Style: Standard overlapped window with system menu
;
; Stack frame: 152 bytes for WNDCLASSW structure
; ==============================================================================
CreateMainWindow proc frame
    push rbx
    .pushreg rbx
    push rsi
    .pushreg rsi
    push rdi
    .pushreg rdi
    push r12
    .pushreg r12
    sub rsp, 152
    .allocstack 152
    .endprolog

    mov r12, rcx                ; Save hInstance

    ; Initialize WNDCLASSW structure (zero it out)
    lea rdi, [rsp+40]
    xor rax, rax
    mov rcx, 10                 ; 10 QWORDs = 80 bytes
@@:
    test rcx, rcx
    jz @F
    mov qword ptr [rdi], rax
    add rdi, 8
    dec rcx
    jmp @B
@@:

    ; Fill in WNDCLASSW fields
    mov dword ptr [rsp+40], CS_HREDRAW or CS_VREDRAW ; style
    lea rax, WndProc
    mov [rsp+40+8], rax         ; lpfnWndProc
    mov [rsp+40+24], r12        ; hInstance

    ; Extract icon from shell32.dll (shield icon at index 104)
    sub rsp, 32
    mov dword ptr [rsp+32], 1   ; nIcons
    lea r9, [rsp+32+136]        ; phiconSmall
    lea r8, [rsp+32+128]        ; phiconLarge
    mov edx, 104                ; nIconIndex
    lea rcx, str_Shell32        ; lpszFile
    call ExtractIconExW
    add rsp, 32

    ; Set icon in window class
    mov rax, [rsp+128]
    mov [rsp+40+32], rax        ; hIcon

    ; Load standard arrow cursor
    mov rdx, IDC_ARROW_ATOM
    xor ecx, ecx
    sub rsp, 32
    call LoadCursorW
    add rsp, 32
    mov [rsp+40+40], rax        ; hCursor

    ; No background brush - allow Mica backdrop to be visible
    xor rax, rax
    mov [rsp+40+48], rax        ; hbrBackground = NULL

    ; Set class name
    lea rax, str_ClassName
    mov [rsp+40+64], rax        ; lpszClassName

    ; Register window class
    lea rcx, [rsp+40]
    sub rsp, 32
    call RegisterClassW
    add rsp, 32
    test eax, eax
    jz cmw_fail                 ; Registration failed

    ; Create main window
    sub rsp, 96
    xor rax, rax
    mov [rsp+88], rax           ; lpParam
    mov [rsp+80], r12           ; hInstance
    xor rax, rax
    mov [rsp+72], rax           ; hMenu
    xor rax, rax
    mov [rsp+64], rax           ; hWndParent
    mov dword ptr [rsp+56], 150 ; nHeight
    mov dword ptr [rsp+48], 420 ; nWidth
    mov dword ptr [rsp+40], 100 ; Y
    mov dword ptr [rsp+32], 100 ; X
    mov r9d, STY_MAINWIN        ; dwStyle
    lea r8, str_Title           ; lpWindowName
    lea rdx, str_ClassName      ; lpClassName
    xor ecx, ecx                ; dwExStyle
    call CreateWindowExW
    add rsp, 96
    test rax, rax
    jz cmw_fail                 ; Creation failed
    mov g_hwndMain, rax         ; Save window handle

    ; Show window
    mov edx, SW_SHOWNORMAL
    mov rcx, rax
    sub rsp, 32
    call ShowWindow
    add rsp, 32

    ; Update window display
    mov rcx, g_hwndMain
    sub rsp, 32
    call UpdateWindow
    add rsp, 32

    ; Apply theme attributes (dark mode + Mica backdrop)
    mov rcx, g_hwndMain
    sub rsp, 32
    call ApplyWindowTheme
    add rsp, 32

    mov rax, g_hwndMain         ; Return window handle
    jmp cmw_done

cmw_fail:
    xor rax, rax                ; Return NULL on failure

cmw_done:
    add rsp, 152
    pop r12
    pop rdi
    pop rsi
    pop rbx
    ret
CreateMainWindow endp

end
