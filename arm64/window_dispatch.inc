; ==============================================================================
; window_dispatch.inc  -  Included by window.asm (ARM64 / armasm64)
;
; WndProc prologue, tray delegation, message dispatch, WM_CREATE construction
; and the WM_COMMAND handlers (including the Browse dialog + .lnk resolution).
; The procedure is continued verbatim in window_events.inc, which supplies the
; remaining handlers and the single shared epilogue (wp_done).
;
; hWnd/uMsg/wParam/lParam live in x19/w20/x21/x22 for the whole procedure.
; Ported from x64/window_dispatch.inc.
; ==============================================================================

WndProc PROC
    STP x29, x30, [sp, #-16]!
    MOV x29, sp
    STP x19, x20, [sp, #-16]!
    STP x21, x22, [sp, #-16]!
    STP x23, x24, [sp, #-16]!
    STP x25, x26, [sp, #-16]!
    STP x27, x28, [sp, #-16]!
    SUB sp, sp, #WPF_SIZE

    MOV x19, x0                     ; hWnd
    MOV w20, w1                     ; uMsg
    MOV x21, x2                     ; wParam
    MOV x22, x3                     ; lParam

    ; Tray behavior (Shift+minimize, tray callbacks, TaskbarCreated) is fully
    ; isolated in tray.asm; a nonzero return means the message is consumed.
    MOV x0, x19
    MOV w1, w20
    MOV x2, x21
    MOV x3, x22
    BL TrayHandleWindowMessage
    CBNZ w0, wp_handled

    CMP w20, #WM_CREATE
    B.EQ wp_create
    CMP w20, #WM_COMMAND
    B.EQ wp_command
    CMP w20, #WM_CHAR
    B.EQ wp_char
    CMP w20, #WM_SIZE
    B.EQ wp_size
    CMP w20, #WM_ERASEBKGND
    B.EQ wp_erasebkgnd
    CMP w20, #WM_CTLCOLOREDIT
    B.EQ wp_ctlcolor_edit
    CMP w20, #WM_CTLCOLORLISTBOX
    B.EQ wp_ctlcolor_edit
    CMP w20, #WM_CTLCOLORSTATIC
    B.EQ wp_ctlcolor_static
    CMP w20, #WM_CTLCOLORBTN
    B.EQ wp_ctlcolor_static
    CMP w20, #WM_SETTINGCHANGE
    B.EQ wp_themechange
    CMP w20, #WM_THEMECHANGED
    B.EQ wp_theme_repaint
    CMP w20, #WM_DESTROY
    B.EQ wp_destroy
    CMP w20, #WM_DROPFILES
    B.EQ wp_dropfiles
    CMP w20, #WM_KEYDOWN
    B.EQ wp_keydown
    B wp_defproc

wp_handled
    MOV w0, #0
    B wp_done

; ------------------------------------------------------------------------------
; WM_CREATE
; ------------------------------------------------------------------------------
wp_create
    LDR x23, [x22, #8]              ; x23 = hInstance (CREATESTRUCT.hInstance)

    ; History-enabled = does HKCU\Software\cmdt already exist?
    MOVZ w0, #0x0001
    MOVK w0, #0x8000, LSL #16       ; HKEY_CURRENT_USER
    ADRP x1, str_regKey
    ADD x1, x1, str_regKey
    MOV x2, XZR
    MOVZ w3, #0x0019
    MOVK w3, #0x0002, LSL #16   ; KEY_READ = 0x00020019
    ADD x4, sp, #WP_HKEY
    BL RegOpenKeyExW
    CBNZ w0, wp_hist_off
    LDR x0, [sp, #WP_HKEY]
    BL RegCloseKey
    MOV w9, #1
    ADRP x10, g_historyEnabled
    ADD x10, x10, g_historyEnabled
    STR w9, [x10]
    B wp_hist_done
wp_hist_off
    ADRP x10, g_historyEnabled
    ADD x10, x10, g_historyEnabled
    STR wzr, [x10]
wp_hist_done

    BL InitDpiApis
    BL EnsureThemeBrushes

    ; Per-window DPI (w26 for the rest of WM_CREATE; falls back to 96).
    ADRP x9, g_pGetDpiForWindow
    ADD x9, x9, g_pGetDpiForWindow
    LDR x9, [x9]
    CBZ x9, wp_create_dpi96
    MOV x0, x19
    BLR x9
    MOV w26, w0
    B wp_create_dpi_done
wp_create_dpi96
    MOV w26, #96
wp_create_dpi_done
    STR w26, [sp, #WP_DPI]

    ; Read the theme preference (only read here; applied after controls exist).
    ADRP x10, g_darkMode
    ADD x10, x10, g_darkMode
    STR wzr, [x10]
    MOVZ w0, #0x0001
    MOVK w0, #0x8000, LSL #16
    ADRP x1, str_regThemeKey
    ADD x1, x1, str_regThemeKey
    MOV x2, XZR
    MOVZ w3, #0x0019
    MOVK w3, #0x0002, LSL #16   ; KEY_READ = 0x00020019
    ADD x4, sp, #WP_HKEY2
    BL RegOpenKeyExW
    CBNZ w0, wp_theme_done
    MOV w9, #1
    STR w9, [sp, #WP_DW1]           ; AppsUseLightTheme default = 1
    MOV w9, #4
    STR w9, [sp, #WP_DW2]           ; size
    LDR x0, [sp, #WP_HKEY2]
    ADRP x1, str_regAppsUseLightTheme
    ADD x1, x1, str_regAppsUseLightTheme
    MOV x2, XZR
    MOV x3, XZR
    ADD x4, sp, #WP_DW1
    ADD x5, sp, #WP_DW2
    BL RegQueryValueExW
    CBNZ w0, wp_theme_close
    LDR w9, [sp, #WP_DW1]
    CBNZ w9, wp_theme_close
    MOV w9, #1
    ADRP x10, g_darkMode
    ADD x10, x10, g_darkMode
    STR w9, [x10]
wp_theme_close
    LDR x0, [sp, #WP_HKEY2]
    BL RegCloseKey
wp_theme_done
    ; Defer owner-draw: stash dark state, force light during menu build.
    ADRP x10, g_darkMode
    ADD x10, x10, g_darkMode
    LDR w9, [x10]
    STR w9, [sp, #WP_DARKSV]
    STR wzr, [x10]

    ; Build the menu.  x24 = main menu, x25 = File popup.
    BL CreateMenu
    MOV x24, x0
    BL CreatePopupMenu
    MOV x25, x0
    ADRP x9, g_hMenuFile
    ADD x9, x9, g_hMenuFile
    STR x25, [x9]

    ; Browse
    MOV x0, x25
    MOV w1, #MF_STRING
    MOV x2, #ID_FILE_BROWSE
    ADRP x3, str_MenuBrowse
    ADD x3, x3, str_MenuBrowse
    BL AppendMenuW
    ; separator
    MOV x0, x25
    MOV w1, #MF_SEPARATOR
    MOV x2, XZR
    MOV x3, XZR
    BL AppendMenuW
    ; Enable History (checked reflects current state)
    MOV x0, x25
    MOV w1, #MF_STRING
    ADRP x9, g_historyEnabled
    ADD x9, x9, g_historyEnabled
    LDR w9, [x9]
    CBZ w9, wp_menu_hist
    ORR w1, w1, #MF_CHECKED
wp_menu_hist
    MOV x2, #ID_FILE_HISTORY
    ADRP x3, str_MenuHistory
    ADD x3, x3, str_MenuHistory
    BL AppendMenuW
    ; separator
    MOV x0, x25
    MOV w1, #MF_SEPARATOR
    MOV x2, XZR
    MOV x3, XZR
    BL AppendMenuW
    ; About
    MOV x0, x25
    MOV w1, #MF_STRING
    MOV x2, #ID_FILE_ABOUT
    ADRP x3, str_MenuAbout
    ADD x3, x3, str_MenuAbout
    BL AppendMenuW
    ; separator
    MOV x0, x25
    MOV w1, #MF_SEPARATOR
    MOV x2, XZR
    MOV x3, XZR
    BL AppendMenuW
    ; Exit
    MOV x0, x25
    MOV w1, #MF_STRING
    MOV x2, #ID_FILE_EXIT
    ADRP x3, str_MenuExit
    ADD x3, x3, str_MenuExit
    BL AppendMenuW
    ; File popup -> menu bar
    MOV x0, x24
    MOV w1, #MF_POPUP
    MOV x2, x25
    ADRP x3, str_MenuFile
    ADD x3, x3, str_MenuFile
    BL AppendMenuW
    ; attach
    MOV x0, x19
    MOV x1, x24
    BL SetMenu

    ; restore dark state
    LDR w9, [sp, #WP_DARKSV]
    ADRP x10, g_darkMode
    ADD x10, x10, g_darkMode
    STR w9, [x10]

    ; ---- ComboBox ----
    MOV w0, #WS_EX_CLIENTEDGE
    ADRP x1, str_ComboBox
    ADD x1, x1, str_ComboBox
    MOV x2, XZR
    MOVZ w3, #0x0042
    MOVK w3, #0x5001, LSL #16       ; STY_COMBO
    MOV w9, #10
    MUL w9, w9, w26
    MOV w10, #96
    UDIV w9, w9, w10                ; scaled 10
    MOV w4, w9                      ; X
    MOV w5, w9                      ; Y
    MOV w9, #460
    MUL w9, w9, w26
    UDIV w6, w9, w10                ; nWidth
    MOV w9, #200
    MUL w9, w9, w26
    UDIV w7, w9, w10                ; nHeight
    STR x19, [sp, #WP_OUT + 0]      ; hWndParent
    MOV x9, #IDC_COMBO
    STR x9, [sp, #WP_OUT + 8]       ; hMenu (control id)
    STR x23, [sp, #WP_OUT + 16]     ; hInstance
    STR xzr, [sp, #WP_OUT + 24]     ; lpParam
    BL CreateWindowExW
    CBZ x0, wp_create_fail
    ADRP x9, g_hwndEdit
    ADD x9, x9, g_hwndEdit
    STR x0, [x9]

    ; Load MRU only when history is enabled (no registry touch otherwise).
    ADRP x9, g_historyEnabled
    ADD x9, x9, g_historyEnabled
    LDR w9, [x9]
    CBZ w9, wp_skip_loadmru
    BL LoadMRU
wp_skip_loadmru

    ; ---- Browse button ----
    MOV w0, #0
    ADRP x1, str_Button
    ADD x1, x1, str_Button
    ADRP x2, str_BtnBrowse
    ADD x2, x2, str_BtnBrowse
    MOVZ w3, #0x5001, LSL #16       ; STY_BUTTON
    MOV w10, #96
    MOV w9, #10
    MUL w9, w9, w26
    UDIV w4, w9, w10                ; X = 10
    MOV w9, #45
    MUL w9, w9, w26
    UDIV w5, w9, w10                ; Y = 45
    MOV w9, #80
    MUL w9, w9, w26
    UDIV w6, w9, w10                ; W = 80
    MOV w9, #25
    MUL w9, w9, w26
    UDIV w7, w9, w10                ; H = 25
    STR x19, [sp, #WP_OUT + 0]
    MOV x9, #IDC_BTN_BROWSE
    STR x9, [sp, #WP_OUT + 8]
    STR x23, [sp, #WP_OUT + 16]
    STR xzr, [sp, #WP_OUT + 24]
    BL CreateWindowExW
    CBZ x0, wp_create_fail

    ; ---- Run button (default) ----
    MOV w0, #0
    ADRP x1, str_Button
    ADD x1, x1, str_Button
    ADRP x2, str_BtnRun
    ADD x2, x2, str_BtnRun
    MOVZ w3, #0x0001
    MOVK w3, #0x5001, LSL #16       ; STY_BUTTON_DEF
    MOV w10, #96
    MOV w9, #400
    MUL w9, w9, w26
    UDIV w4, w9, w10                ; X = 400
    MOV w9, #45
    MUL w9, w9, w26
    UDIV w5, w9, w10                ; Y = 45
    MOV w9, #80
    MUL w9, w9, w26
    UDIV w6, w9, w10                ; W = 80
    MOV w9, #25
    MUL w9, w9, w26
    UDIV w7, w9, w10                ; H = 25
    STR x19, [sp, #WP_OUT + 0]
    MOV x9, #IDC_BTN_RUN
    STR x9, [sp, #WP_OUT + 8]
    STR x23, [sp, #WP_OUT + 16]
    STR xzr, [sp, #WP_OUT + 24]
    BL CreateWindowExW
    CBZ x0, wp_create_fail
    ADRP x9, g_hwndBtn
    ADD x9, x9, g_hwndBtn
    STR x0, [x9]

    ; ---- Status label ----
    MOV w0, #0
    ADRP x1, str_Static
    ADD x1, x1, str_Static
    ADRP x2, str_StatusInit
    ADD x2, x2, str_StatusInit
    MOVZ w3, #0x0200
    MOVK w3, #0x5000, LSL #16       ; STY_STATIC
    MOV w10, #96
    MOV w9, #100
    MUL w9, w9, w26
    UDIV w4, w9, w10                ; X = 100
    MOV w9, #48
    MUL w9, w9, w26
    UDIV w5, w9, w10                ; Y = 48
    MOV w9, #200
    MUL w9, w9, w26
    UDIV w6, w9, w10                ; W = 200
    MOV w9, #20
    MUL w9, w9, w26
    UDIV w7, w9, w10                ; H = 20
    STR x19, [sp, #WP_OUT + 0]
    MOV x9, #IDC_STATIC_STATUS
    STR x9, [sp, #WP_OUT + 8]
    STR x23, [sp, #WP_OUT + 16]
    STR xzr, [sp, #WP_OUT + 24]
    BL CreateWindowExW
    CBZ x0, wp_create_fail
    ADRP x9, g_hwndStatus
    ADD x9, x9, g_hwndStatus
    STR x0, [x9]

    ; SetFocus(g_hwndEdit)
    ADRP x0, g_hwndEdit
    ADD x0, x0, g_hwndEdit
    LDR x0, [x0]
    BL SetFocus

    ; DragAcceptFiles(hWnd, TRUE)
    MOV x0, x19
    MOV w1, #1
    BL DragAcceptFiles

    ; Allow WM_DROPFILES / WM_COPYGLOBALDATA through UIPI.
    MOV x0, x19
    MOV w1, #WM_DROPFILES
    MOV w2, #MSGFLT_ALLOW
    MOV x3, XZR
    BL ChangeWindowMessageFilterEx
    MOV x0, x19
    MOV w1, #WM_COPYGLOBALDATA
    MOV w2, #MSGFLT_ALLOW
    MOV x3, XZR
    BL ChangeWindowMessageFilterEx

    MOV w0, #0
    B wp_done

wp_create_fail
    MOV x0, #-1
    B wp_done

; ------------------------------------------------------------------------------
; WM_COMMAND
; ------------------------------------------------------------------------------
wp_command
    UXTH w23, w21                  ; w23 = control ID (low word)
    LSR w24, w21, #16              ; w24 = notification code (high word)

    CMP w23, #IDC_EDIT
    B.EQ wp_check_edit
    CMP w23, #IDC_BTN_RUN
    B.EQ wp_cmd_run
    CMP w23, #1                    ; IDOK
    B.EQ wp_cmd_run
    CMP w23, #IDC_BTN_BROWSE
    B.EQ wp_cmd_browse
    CMP w23, #ID_FILE_BROWSE
    B.EQ wp_cmd_browse
    CMP w23, #ID_FILE_EXIT
    B.EQ wp_cmd_exit
    CMP w23, #ID_FILE_ABOUT
    B.EQ wp_cmd_about
    CMP w23, #ID_FILE_HISTORY
    B.EQ wp_cmd_history
    B wp_defproc

wp_check_edit
    CMP w24, #EN_SETFOCUS
    B.NE wp_defproc
    ; Edit got focus -> make Run the default button.
    ADRP x0, g_hwndBtn
    ADD x0, x0, g_hwndBtn
    LDR x0, [x0]
    MOV w1, #BM_SETSTYLE
    MOV w2, #BS_DEFPUSHBUTTON
    MOV w3, #1
    BL SendMessageW
    MOV w0, #0
    B wp_done

wp_cmd_run
    MOV x0, x19
    BL RunCommand
    B wp_done

wp_cmd_exit
    MOV w0, #0
    BL PostQuitMessage
    MOV w0, #0
    B wp_done

wp_cmd_about
    MOV x0, x19
    ADRP x1, str_AboutText
    ADD x1, x1, str_AboutText
    ADRP x2, str_AboutTitle
    ADD x2, x2, str_AboutTitle
    MOV w3, #MB_ICONINFORMATION
    BL MessageBoxW
    MOV w0, #0
    B wp_done

wp_cmd_history
    ADRP x25, g_historyEnabled
    ADD x25, x25, g_historyEnabled
    LDR w9, [x25]
    CBNZ w9, wp_hist_turn_off
    ; OFF -> ON. Key is created lazily by SaveMRU on the next run.
    MOV w9, #1
    STR w9, [x25]
    MOV w24, #(MF_BYCOMMAND + MF_CHECKED)
    B wp_hist_setcheck
wp_hist_turn_off
    ; ON -> OFF, and wipe stored history immediately.
    STR wzr, [x25]
    MOVZ w0, #0x0001
    MOVK w0, #0x8000, LSL #16       ; HKEY_CURRENT_USER
    ADRP x1, str_regKey
    ADD x1, x1, str_regKey
    BL RegDeleteTreeW
    MOV w24, #(MF_BYCOMMAND + MF_UNCHECKED)
wp_hist_setcheck
    ADRP x0, g_hMenuFile
    ADD x0, x0, g_hMenuFile
    LDR x0, [x0]
    MOV w1, #ID_FILE_HISTORY
    MOV w2, w24
    BL CheckMenuItem
    MOV w0, #0
    B wp_done

; ------------------------------------------------------------------------------
; WM_COMMAND -> Browse (file open dialog + optional .lnk resolution)
; ------------------------------------------------------------------------------
wp_cmd_browse
    ; Zero OPENFILENAMEW (152 bytes = 19 qwords).
    ADD x8, sp, #WP_OFN
    MOV w9, #19
wp_ofn_zero
    STR xzr, [x8], #8
    SUBS w9, w9, #1
    B.NE wp_ofn_zero

    MOV w9, #152
    STR w9, [sp, #WP_OFN + 0]       ; lStructSize
    STR x19, [sp, #WP_OFN + 8]      ; hwndOwner
    ADRP x9, str_Filter
    ADD x9, x9, str_Filter
    STR x9, [sp, #WP_OFN + 24]      ; lpstrFilter
    ADRP x9, g_filePath
    ADD x9, x9, g_filePath
    STR x9, [sp, #WP_OFN + 48]      ; lpstrFile
    MOV w9, #520
    STR w9, [sp, #WP_OFN + 56]      ; nMaxFile
    ADRP x9, str_DefPath
    ADD x9, x9, str_DefPath
    STR x9, [sp, #WP_OFN + 80]      ; lpstrInitialDir
    MOV w9, #(OFN_FILEMUSTEXIST + OFN_PATHMUSTEXIST)
    STR w9, [sp, #WP_OFN + 96]      ; Flags

    ; Clear the path buffer's first char.
    ADRP x9, g_filePath
    ADD x9, x9, g_filePath
    STRH wzr, [x9]

    ADD x0, sp, #WP_OFN
    BL GetOpenFileNameW
    CBZ w0, wp_browse_cancel

    ; .lnk? Need at least 4 chars.
    ADRP x0, g_filePath
    ADD x0, x0, g_filePath
    BL wcslen_p
    MOV x24, x0                     ; length (chars)
    CMP x24, #4
    B.LT wp_browse_set

    ADRP x0, g_filePath
    ADD x0, x0, g_filePath
    ADD x0, x0, x24, LSL #1         ; end of string
    SUB x0, x0, #8                  ; last 4 chars
    ADRP x1, str_extLnk
    ADD x1, x1, str_extLnk
    BL wcscmp_ci
    CBZ w0, wp_browse_set           ; not .lnk

    ; Resolve .lnk -> g_cmdBuf (target) + g_tempBuf (args)
    ADRP x0, g_filePath
    ADD x0, x0, g_filePath
    ADRP x1, g_cmdBuf
    ADD x1, x1, g_cmdBuf
    ADRP x2, g_tempBuf
    ADD x2, x2, g_tempBuf
    BL ResolveLnkPath
    CBZ w0, wp_browse_set

    ; If target non-empty, append a space before the args.
    ADRP x0, g_cmdBuf
    ADD x0, x0, g_cmdBuf
    BL wcslen_p
    CBZ x0, wp_browse_args_only
    ADRP x9, g_cmdBuf
    ADD x9, x9, g_cmdBuf
    ADD x9, x9, x0, LSL #1
    MOV w10, #' '
    STRH w10, [x9]
    STRH wzr, [x9, #2]
wp_browse_args_only
    ADRP x0, g_cmdBuf
    ADD x0, x0, g_cmdBuf
    ADRP x1, g_tempBuf
    ADD x1, x1, g_tempBuf
    BL wcscat_p
    ADRP x0, g_hwndEdit
    ADD x0, x0, g_hwndEdit
    LDR x0, [x0]
    ADRP x1, g_cmdBuf
    ADD x1, x1, g_cmdBuf
    BL SetWindowTextW
    B wp_browse_cancel

wp_browse_set
    ADRP x0, g_hwndEdit
    ADD x0, x0, g_hwndEdit
    LDR x0, [x0]
    ADRP x1, g_filePath
    ADD x1, x1, g_filePath
    BL SetWindowTextW

wp_browse_cancel
    ADRP x0, g_hwndEdit
    ADD x0, x0, g_hwndEdit
    LDR x0, [x0]
    BL SetFocus
    ADRP x0, g_hwndBtn
    ADD x0, x0, g_hwndBtn
    LDR x0, [x0]
    MOV w1, #BM_SETSTYLE
    MOV w2, #BS_DEFPUSHBUTTON
    MOV w3, #1
    BL SendMessageW
    MOV w0, #0
    B wp_done

; ------------------------------------------------------------------------------
; WM_CHAR - Enter in the ComboBox runs the command.
; ------------------------------------------------------------------------------
wp_char
    CMP w21, #13                   ; VK_RETURN
    B.NE wp_defproc
    BL GetFocus
    ADRP x9, g_hwndEdit
    ADD x9, x9, g_hwndEdit
    LDR x9, [x9]
    CMP x0, x9
    B.NE wp_defproc
    MOV x0, x19
    BL RunCommand
    MOV w0, #0
    B wp_done



    END
