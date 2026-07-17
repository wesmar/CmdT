; WndProc declaration, dispatch, creation and command handlers.
; Continued verbatim by window_events.inc in the same procedure.

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
    cmp eax, WM_ERASEBKGND
    je wp_erasebkgnd
    cmp eax, WM_MEASUREITEM
    je wp_measureitem
    cmp eax, WM_DRAWITEM
    je wp_drawitem
    cmp eax, WM_CTLCOLOREDIT
    je wp_ctlcolor_edit
    cmp eax, WM_CTLCOLORLISTBOX
    je wp_ctlcolor_edit
    cmp eax, WM_CTLCOLORSTATIC
    je wp_ctlcolor_static
    cmp eax, WM_CTLCOLORBTN
    je wp_ctlcolor_static
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

    ; Determine history-enabled state (registry key existence) before
    ; building the menu, so the checkbox starts in the right state and no
    ; registry touch happens if the key was never created.
    lea rax, [rbp-112]
    sub rsp, 48
    mov [rsp+32], rax           ; phkResult -> &[rbp-112]
    mov r9d, KEY_READ
    xor r8d, r8d                ; ulOptions
    lea rdx, str_regKey
    mov ecx, HKEY_CURRENT_USER
    movsxd rcx, ecx             ; predefined HKEY values are signed handles
    call RegOpenKeyExW
    add rsp, 48
    test eax, eax
    jnz wp_hist_disabled
    mov rcx, [rbp-112]
    sub rsp, 32
    call RegCloseKey
    add rsp, 32
    mov dword ptr g_historyEnabled, 1
    jmp wp_hist_state_done
wp_hist_disabled:
    mov dword ptr g_historyEnabled, 0
wp_hist_state_done:

    ; Per-monitor DPI for this window, used to scale every child control's
    ; design-time (96 DPI) coordinates below. Stashed at [rbp-120] since the
    ; control-creation blocks below span many calls that would otherwise
    ; clobber a scratch register. Falls back to 96 (100%, i.e. no scaling)
    ; on Windows 7/8/8.1, where GetDpiForWindow doesn't exist.
    sub rsp, 32
    call InitDpiApis
    add rsp, 32
    sub rsp, 32
    call EnsureThemeBrushes
    add rsp, 32
    mov rax, qword ptr g_pGetDpiForWindow
    test rax, rax
    jz wp_dpi_fallback
    mov rcx, [rbp-72]           ; hWnd
    sub rsp, 32
    call rax
    add rsp, 32
    mov dword ptr [rbp-120], eax
    jmp wp_dpi_done
wp_dpi_fallback:
    mov dword ptr [rbp-120], 96
wp_dpi_done:

    ; Theme state must be known before menu creation because owner-draw flags
    ; are fixed by AppendMenuW. Only read the preference here; applying DWM
    ; and control themes inside WM_CREATE can recursively send theme messages
    ; before child construction is complete.
    mov dword ptr g_darkMode, 0
    lea rax, [rbp-136]                  ; temporary HKEY
    sub rsp, 48
    mov [rsp+32], rax
    mov r9d, KEY_READ
    xor r8d, r8d
    lea rdx, str_regThemeKey
    mov ecx, HKEY_CURRENT_USER
    movsxd rcx, ecx             ; predefined HKEY values are signed handles
    call RegOpenKeyExW
    add rsp, 48
    test eax, eax
    jnz wp_theme_pref_done
    mov dword ptr [rbp-140], 1          ; AppsUseLightTheme default
    mov dword ptr [rbp-144], 4
    sub rsp, 48
    lea rax, [rbp-144]
    mov [rsp+40], rax
    lea rax, [rbp-140]
    mov [rsp+32], rax
    xor r9d, r9d
    xor r8d, r8d
    lea rdx, str_regAppsUseLightTheme
    mov rcx, [rbp-136]
    call RegQueryValueExW
    add rsp, 48
    test eax, eax
    jnz wp_theme_pref_close
    cmp dword ptr [rbp-140], 0
    jne wp_theme_pref_close
    mov dword ptr g_darkMode, 1
wp_theme_pref_close:
    mov rcx, [rbp-136]
    sub rsp, 32
    call RegCloseKey
    add rsp, 32
wp_theme_pref_done:
    ; Defer owner-draw activation until all child windows exist. USER32 may
    ; synchronously measure menu items during SetMenu; doing that in the
    ; middle of WM_CREATE can abort top-level creation on some builds.
    mov eax, dword ptr g_darkMode
    mov dword ptr [rbp-148], eax
    mov dword ptr g_darkMode, 0

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
    mov g_hMenuFile, rax         ; Keep for later CheckMenuItem calls

    ; Add "Browse" menu item
    lea r9, str_MenuBrowse
    mov r8d, ID_FILE_BROWSE
    mov edx, MF_STRING
    cmp dword ptr g_darkMode, 0
    je @F
    or edx, MF_OWNERDRAW
@@:
    mov rcx, r13
    sub rsp, 32
    call AppendMenuW
    add rsp, 32

    ; Add separator
    xor r9d, r9d
    xor r8d, r8d
    mov edx, MF_SEPARATOR
    cmp dword ptr g_darkMode, 0
    je @F
    or edx, MF_OWNERDRAW
@@:
    mov rcx, r13
    sub rsp, 32
    call AppendMenuW
    add rsp, 32

    ; Add "Enable History" checkbox item (checked state reflects whether
    ; the MRU registry key already existed at startup)
    lea r9, str_MenuHistory
    mov r8d, ID_FILE_HISTORY
    mov edx, MF_STRING
    cmp dword ptr g_darkMode, 0
    je @F
    or edx, MF_OWNERDRAW
@@:
    cmp dword ptr g_historyEnabled, 0
    je @F
    or edx, MF_CHECKED
@@:
    mov rcx, r13
    sub rsp, 32
    call AppendMenuW
    add rsp, 32

    ; Add separator
    xor r9d, r9d
    xor r8d, r8d
    mov edx, MF_SEPARATOR
    cmp dword ptr g_darkMode, 0
    je @F
    or edx, MF_OWNERDRAW
@@:
    mov rcx, r13
    sub rsp, 32
    call AppendMenuW
    add rsp, 32

    ; Add "About" menu item
    lea r9, str_MenuAbout
    mov r8d, ID_FILE_ABOUT
    mov edx, MF_STRING
    cmp dword ptr g_darkMode, 0
    je @F
    or edx, MF_OWNERDRAW
@@:
    mov rcx, r13
    sub rsp, 32
    call AppendMenuW
    add rsp, 32

    ; Add separator
    xor r9d, r9d
    xor r8d, r8d
    mov edx, MF_SEPARATOR
    cmp dword ptr g_darkMode, 0
    je @F
    or edx, MF_OWNERDRAW
@@:
    mov rcx, r13
    sub rsp, 32
    call AppendMenuW
    add rsp, 32

    ; Add "Exit" menu item
    lea r9, str_MenuExit
    mov r8d, ID_FILE_EXIT
    mov edx, MF_STRING
    cmp dword ptr g_darkMode, 0
    je @F
    or edx, MF_OWNERDRAW
@@:
    mov rcx, r13
    sub rsp, 32
    call AppendMenuW
    add rsp, 32

    ; Add File submenu to main menu
    lea r9, str_MenuFile
    mov r8, r13                 ; Popup menu handle
    mov edx, MF_POPUP
    cmp dword ptr g_darkMode, 0
    je @F
    or edx, MF_OWNERDRAW
@@:
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
    mov eax, dword ptr [rbp-148]
    mov dword ptr g_darkMode, eax

    ; Create ComboBox control
    sub rsp, 96
    xor rax, rax
    mov [rsp+88], rax           ; lpParam
    mov [rsp+80], r15           ; hInstance
    mov rax, IDC_COMBO
    mov [rsp+72], rax           ; hMenu (control ID)
    mov rax, [rbp-72]
    mov [rsp+64], rax           ; hWndParent
    mov eax, 200
    imul eax, dword ptr [rbp-120]
    mov ecx, 96
    cdq
    idiv ecx
    mov dword ptr [rsp+56], eax ; nHeight
    mov eax, 460
    imul eax, dword ptr [rbp-120]
    mov ecx, 96
    cdq
    idiv ecx
    mov dword ptr [rsp+48], eax ; nWidth
    mov eax, 10
    imul eax, dword ptr [rbp-120]
    mov ecx, 96
    cdq
    idiv ecx
    mov dword ptr [rsp+40], eax ; Y
    mov dword ptr [rsp+32], eax ; X (same 10 design value)
    mov r9d, STY_COMBO          ; dwStyle
    xor r8d, r8d                ; lpWindowName
    lea rdx, str_ComboBox       ; lpClassName
    mov ecx, WS_EX_CLIENTEDGE   ; dwExStyle
    call CreateWindowExW
    add rsp, 96
    test rax, rax
    jz wp_create_fail
    mov g_hwndEdit, rax         ; Save ComboBox handle

    ; Load MRU list into ComboBox (skip entirely when history is off —
    ; no registry touch at all in that case)
    cmp dword ptr g_historyEnabled, 0
    je wp_skip_loadmru
    call LoadMRU
wp_skip_loadmru:

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
    mov eax, 25
    imul eax, dword ptr [rbp-120]
    mov ecx, 96
    cdq
    idiv ecx
    mov dword ptr [rsp+56], eax ; nHeight
    mov eax, 80
    imul eax, dword ptr [rbp-120]
    mov ecx, 96
    cdq
    idiv ecx
    mov dword ptr [rsp+48], eax ; nWidth
    mov eax, 45
    imul eax, dword ptr [rbp-120]
    mov ecx, 96
    cdq
    idiv ecx
    mov dword ptr [rsp+40], eax ; Y
    mov eax, 10
    imul eax, dword ptr [rbp-120]
    mov ecx, 96
    cdq
    idiv ecx
    mov dword ptr [rsp+32], eax ; X
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
    mov eax, 25
    imul eax, dword ptr [rbp-120]
    mov ecx, 96
    cdq
    idiv ecx
    mov dword ptr [rsp+56], eax ; nHeight
    mov eax, 80
    imul eax, dword ptr [rbp-120]
    mov ecx, 96
    cdq
    idiv ecx
    mov dword ptr [rsp+48], eax ; nWidth
    mov eax, 45
    imul eax, dword ptr [rbp-120]
    mov ecx, 96
    cdq
    idiv ecx
    mov dword ptr [rsp+40], eax ; Y
    mov eax, 400
    imul eax, dword ptr [rbp-120]
    mov ecx, 96
    cdq
    idiv ecx
    mov dword ptr [rsp+32], eax ; X
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
    mov eax, 20
    imul eax, dword ptr [rbp-120]
    mov ecx, 96
    cdq
    idiv ecx
    mov dword ptr [rsp+56], eax ; nHeight
    mov eax, 200
    imul eax, dword ptr [rbp-120]
    mov ecx, 96
    cdq
    idiv ecx
    mov dword ptr [rsp+48], eax ; nWidth
    mov eax, 48
    imul eax, dword ptr [rbp-120]
    mov ecx, 96
    cdq
    idiv ecx
    mov dword ptr [rsp+40], eax ; Y
    mov eax, 100
    imul eax, dword ptr [rbp-120]
    mov ecx, 96
    cdq
    idiv ecx
    mov dword ptr [rsp+32], eax ; X
    mov r9d, STY_STATIC         ; dwStyle
    lea r8, str_StatusInit      ; lpWindowName (initial text)
    lea rdx, str_Static         ; lpClassName
    xor ecx, ecx                ; dwExStyle
    call CreateWindowExW
    add rsp, 96
    test rax, rax
    jz wp_create_fail
    mov g_hwndStatus, rax       ; Save status label handle

    ; Convert the completed menu tree to owner-draw only after every child
    ; control exists. This avoids re-entrant measurement during WM_CREATE.
    jmp wp_menu_theme_done       ; native UxTheme dark menu is preferred
    cmp dword ptr g_darkMode, 0
    je wp_menu_theme_done
    sub rsp, 40
    lea rax, str_MenuBrowse
    mov [rsp+32], rax
    mov r9d, ID_FILE_BROWSE
    mov r8d, MF_OWNERDRAW
    mov edx, ID_FILE_BROWSE
    mov rcx, r13
    call ModifyMenuW
    add rsp, 40
    mov r8d, MF_OWNERDRAW
    cmp dword ptr g_historyEnabled, 0
    je @F
    or r8d, MF_CHECKED
@@:
    sub rsp, 40
    lea rax, str_MenuHistory
    mov [rsp+32], rax
    mov r9d, ID_FILE_HISTORY
    mov edx, ID_FILE_HISTORY
    mov rcx, r13
    call ModifyMenuW
    add rsp, 40
    sub rsp, 40
    lea rax, str_MenuAbout
    mov [rsp+32], rax
    mov r9d, ID_FILE_ABOUT
    mov r8d, MF_OWNERDRAW
    mov edx, ID_FILE_ABOUT
    mov rcx, r13
    call ModifyMenuW
    add rsp, 40
    sub rsp, 40
    lea rax, str_MenuExit
    mov [rsp+32], rax
    mov r9d, ID_FILE_EXIT
    mov r8d, MF_OWNERDRAW
    mov edx, ID_FILE_EXIT
    mov rcx, r13
    call ModifyMenuW
    add rsp, 40
    sub rsp, 40
    lea rax, str_MenuFile
    mov [rsp+32], rax
    mov r9, r13
    mov r8d, MF_BYPOSITION or MF_POPUP or MF_OWNERDRAW
    xor edx, edx
    mov rcx, r12
    call ModifyMenuW
    add rsp, 40
wp_menu_theme_done:

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
    cmp r15d, ID_FILE_HISTORY
    je wp_cmd_history
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

wp_cmd_history:
    ; Toggle "Enable History" checkbox
    cmp dword ptr g_historyEnabled, 0
    jne wph_turn_off

    ; Currently off -> turn ON. No registry write here: the key is created
    ; lazily by SaveMRU on the next executed command.
    mov dword ptr g_historyEnabled, 1
    mov r8d, MF_BYCOMMAND or MF_CHECKED
    jmp wph_setcheck

wph_turn_off:
    ; Currently on -> turn OFF and wipe any stored history immediately.
    mov dword ptr g_historyEnabled, 0
    lea rdx, str_regKey
    mov ecx, HKEY_CURRENT_USER
    sub rsp, 32
    call RegDeleteTreeW
    add rsp, 32
    mov r8d, MF_BYCOMMAND or MF_UNCHECKED

wph_setcheck:
    mov edx, ID_FILE_HISTORY
    mov rcx, g_hMenuFile
    sub rsp, 32
    call CheckMenuItem
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

