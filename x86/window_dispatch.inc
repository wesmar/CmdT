; WndProc declaration, dispatch, creation and command handlers.
; Continued verbatim by window_events.inc in the same procedure.

WndProc proc uses ebx esi edi hWnd:DWORD, uMsg:DWORD, wParam:DWORD, lParam:DWORD
    LOCAL rect[4]:DWORD                     ; Client rectangle (RECT structure)
    LOCAL ofn[22]:DWORD                     ; OPENFILENAMEW structure
    LOCAL hMenu:DWORD                       ; Main menu handle
    LOCAL hFileMenu:DWORD                   ; File submenu handle
    LOCAL notifyCode:DWORD                  ; Notification code from WM_COMMAND
    LOCAL hBtnBrowse:DWORD                  ; Browse button handle
    LOCAL dpi:DWORD                         ; Per-monitor DPI (96 = 100%)
    LOCAL sX:DWORD                          ; Scaled X, reused per control
    LOCAL sY:DWORD                          ; Scaled Y, reused per control
    LOCAL sW:DWORD                          ; Scaled width, reused per control
    LOCAL sH:DWORD                          ; Scaled height, reused per control

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
    je wp_theme_repaint
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

    ; Determine history-enabled state (registry key existence) before
    ; building the menu, so the checkbox starts in the right state and no
    ; registry touch happens if the key was never created.
    invoke RegOpenKeyExW, HKEY_CURRENT_USER, offset str_regKey, 0, KEY_READ, addr hMenu
    test eax, eax
    jnz wp_hist_disabled
    invoke RegCloseKey, hMenu
    mov g_historyEnabled, 1
    jmp wp_hist_state_done
wp_hist_disabled:
    mov g_historyEnabled, 0
wp_hist_state_done:

    ; Per-monitor DPI for this window, used to scale every child control's
    ; design-time (96 DPI) coordinates below. Falls back to 96 (no
    ; scaling) on Windows 7/8/8.1, where GetDpiForWindow doesn't exist.
    invoke InitDpiApis
    mov eax, g_pGetDpiForWindow
    test eax, eax
    jz wp_dpi_fallback
    push hWnd
    call eax
    mov dpi, eax
    jmp wp_dpi_done
wp_dpi_fallback:
    mov dpi, 96
wp_dpi_done:

    ; Create main menu
    invoke CreateMenu
    mov hMenu, eax

    ; Create File submenu
    invoke CreatePopupMenu
    mov hFileMenu, eax
    mov g_hMenuFile, eax                    ; Keep for later CheckMenuItem calls

    ; Add menu items to File menu
    invoke AppendMenuW, hFileMenu, MF_STRING, ID_FILE_BROWSE, offset str_MenuBrowse
    invoke AppendMenuW, hFileMenu, MF_SEPARATOR, 0, 0
    mov eax, MF_STRING
    cmp g_historyEnabled, 0
    je @F
    or eax, MF_CHECKED
@@:
    invoke AppendMenuW, hFileMenu, eax, ID_FILE_HISTORY, offset str_MenuHistory
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
    mov eax, 10
    imul eax, dpi
    mov ecx, 96
    cdq
    idiv ecx
    mov sX, eax
    mov sY, eax
    mov eax, 460
    imul eax, dpi
    mov ecx, 96
    cdq
    idiv ecx
    mov sW, eax
    mov eax, 200
    imul eax, dpi
    mov ecx, 96
    cdq
    idiv ecx
    mov sH, eax
    mov eax, lParam
    mov eax, [eax+4]                        ; hInstance
    invoke CreateWindowExW, WS_EX_CLIENTEDGE, offset str_ComboBox, 0, STY_COMBO, sX, sY, sW, sH, hWnd, IDC_COMBO, eax, 0
    test eax, eax
    jz wp_create_fail
    mov g_hwndEdit, eax
    
    ; Load MRU list from registry into ComboBox (skip entirely when history
    ; is off — no registry touch at all in that case)
    cmp g_historyEnabled, 0
    je wp_skip_loadmru
    call LoadMRU
wp_skip_loadmru:
    
    ; Create Browse button
    mov eax, 10
    imul eax, dpi
    mov ecx, 96
    cdq
    idiv ecx
    mov sX, eax
    mov eax, 45
    imul eax, dpi
    mov ecx, 96
    cdq
    idiv ecx
    mov sY, eax
    mov eax, 80
    imul eax, dpi
    mov ecx, 96
    cdq
    idiv ecx
    mov sW, eax
    mov eax, 25
    imul eax, dpi
    mov ecx, 96
    cdq
    idiv ecx
    mov sH, eax
    mov eax, lParam
    mov eax, [eax+4]
    invoke CreateWindowExW, 0, offset str_Button, offset str_BtnBrowse, STY_BUTTON, sX, sY, sW, sH, hWnd, IDC_BTN_BROWSE, eax, 0
    test eax, eax
    jz wp_create_fail

    ; Create Run button (default push button)
    mov eax, 400
    imul eax, dpi
    mov ecx, 96
    cdq
    idiv ecx
    mov sX, eax
    mov eax, 45
    imul eax, dpi
    mov ecx, 96
    cdq
    idiv ecx
    mov sY, eax
    mov eax, 80
    imul eax, dpi
    mov ecx, 96
    cdq
    idiv ecx
    mov sW, eax
    mov eax, 25
    imul eax, dpi
    mov ecx, 96
    cdq
    idiv ecx
    mov sH, eax
    mov eax, lParam
    mov eax, [eax+4]
    invoke CreateWindowExW, 0, offset str_Button, offset str_BtnRun, STY_BUTTON_DEF, sX, sY, sW, sH, hWnd, IDC_BTN_RUN, eax, 0
    test eax, eax
    jz wp_create_fail
    mov g_hwndBtn, eax
    
    ; Create status label (static text control)
    mov eax, 100
    imul eax, dpi
    mov ecx, 96
    cdq
    idiv ecx
    mov sX, eax
    mov eax, 48
    imul eax, dpi
    mov ecx, 96
    cdq
    idiv ecx
    mov sY, eax
    mov eax, 200
    imul eax, dpi
    mov ecx, 96
    cdq
    idiv ecx
    mov sW, eax
    mov eax, 20
    imul eax, dpi
    mov ecx, 96
    cdq
    idiv ecx
    mov sH, eax
    mov eax, lParam
    mov eax, [eax+4]
    invoke CreateWindowExW, 0, offset str_Static, offset str_StatusInit, STY_STATIC, sX, sY, sW, sH, hWnd, IDC_STATIC_STATUS, eax, 0
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
    cmp eax, ID_FILE_HISTORY
    je wp_cmd_history
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
    
wp_cmd_history:
    ; Toggle "Enable History" checkbox
    cmp g_historyEnabled, 0
    jne wph_turn_off

    ; Currently off -> turn ON. No registry write here: the key is created
    ; lazily by SaveMRU on the next executed command.
    mov g_historyEnabled, 1
    invoke CheckMenuItem, g_hMenuFile, ID_FILE_HISTORY, MF_BYCOMMAND or MF_CHECKED
    xor eax, eax
    ret

wph_turn_off:
    ; Currently on -> turn OFF and wipe any stored history immediately.
    mov g_historyEnabled, 0
    invoke RegDeleteTreeW, HKEY_CURRENT_USER, offset str_regKey
    invoke CheckMenuItem, g_hMenuFile, ID_FILE_HISTORY, MF_BYCOMMAND or MF_UNCHECKED
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
    
