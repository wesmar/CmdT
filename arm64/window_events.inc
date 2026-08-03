; ==============================================================================
; window_events.inc  -  Included by window.asm (ARM64 / armasm64)
;
; Continuation of WndProc: layout (WM_SIZE), drag-drop, key input, theming,
; owner-draw colors, teardown, the default handler, and the single shared
; epilogue that every handler branches to via wp_done.
; Ported from x64/window_events.inc.
; ==============================================================================

; ------------------------------------------------------------------------------
; WM_SIZE - relayout every child control for the new client size + DPI.
;   w23 = client width, w26 = DPI, w24 = combo width, w27 = Run X,
;   w28 = status width. Design coordinates are 96-DPI.
; ------------------------------------------------------------------------------
wp_size
    MOV x0, x19
    ADD x1, sp, #WP_RECT
    BL GetClientRect
    LDR w23, [sp, #WP_RECT + 8]     ; width = rect.right

    BL InitDpiApis
    ADRP x9, g_pGetDpiForWindow
    ADD x9, x9, g_pGetDpiForWindow
    LDR x9, [x9]
    CBZ x9, wp_size_dpi96
    MOV x0, x19
    BLR x9
    MOV w26, w0
    B wp_size_dpi_done
wp_size_dpi96
    MOV w26, #96
wp_size_dpi_done

    ; comboWidth = width - scaled(20)
    MOV w10, #96
    MOV w9, #20
    MUL w9, w9, w26
    UDIV w9, w9, w10
    SUB w24, w23, w9

    ; MoveWindow(g_hwndEdit, scaled10, scaled10, comboWidth, scaled25, TRUE)
    ADRP x0, g_hwndEdit
    ADD x0, x0, g_hwndEdit
    LDR x0, [x0]
    MOV w10, #96
    MOV w9, #10
    MUL w9, w9, w26
    UDIV w1, w9, w10               ; X = 10
    MOV w2, w1                     ; Y = 10
    MOV w3, w24                    ; W = comboWidth
    MOV w9, #25
    MUL w9, w9, w26
    UDIV w4, w9, w10               ; H = 25
    MOV w5, #1
    BL MoveWindow

    ; Browse button
    MOV x0, x19
    MOV w1, #IDC_BTN_BROWSE
    BL GetDlgItem
    MOV w10, #96
    MOV w9, #10
    MUL w9, w9, w26
    UDIV w1, w9, w10               ; X = 10
    MOV w9, #45
    MUL w9, w9, w26
    UDIV w2, w9, w10               ; Y = 45
    MOV w9, #80
    MUL w9, w9, w26
    UDIV w3, w9, w10               ; W = 80
    MOV w9, #25
    MUL w9, w9, w26
    UDIV w4, w9, w10               ; H = 25
    MOV w5, #1
    BL MoveWindow

    ; Run button: X = comboWidth - scaled(70)
    MOV w10, #96
    MOV w9, #70
    MUL w9, w9, w26
    UDIV w9, w9, w10
    SUB w27, w24, w9
    ADRP x0, g_hwndBtn
    ADD x0, x0, g_hwndBtn
    LDR x0, [x0]
    MOV w1, w27                    ; X
    MOV w9, #45
    MUL w9, w9, w26
    UDIV w2, w9, w10               ; Y = 45
    MOV w9, #80
    MUL w9, w9, w26
    UDIV w3, w9, w10               ; W = 80
    MOV w9, #25
    MUL w9, w9, w26
    UDIV w4, w9, w10               ; H = 25
    MOV w5, #1
    BL MoveWindow

    ; Status label: X = scaled100, width = RunX - scaled110
    MOV w10, #96
    MOV w9, #110
    MUL w9, w9, w26
    UDIV w9, w9, w10
    SUB w28, w27, w9
    ADRP x0, g_hwndStatus
    ADD x0, x0, g_hwndStatus
    LDR x0, [x0]
    MOV w9, #100
    MUL w9, w9, w26
    UDIV w1, w9, w10              ; X = 100
    MOV w9, #48
    MUL w9, w9, w26
    UDIV w2, w9, w10             ; Y = 48
    MOV w3, w28                   ; W = status width
    MOV w9, #20
    MUL w9, w9, w26
    UDIV w4, w9, w10             ; H = 20
    MOV w5, #1
    BL MoveWindow

    MOV w0, #0
    B wp_done

; ------------------------------------------------------------------------------
; WM_DROPFILES
; ------------------------------------------------------------------------------
wp_dropfiles
    ; DragQueryFileW(hDrop, 0, g_filePath, 520)
    MOV x0, x21                     ; HDROP = wParam
    MOV w1, #0
    ADRP x2, g_filePath
    ADD x2, x2, g_filePath
    MOV w3, #520
    BL DragQueryFileW
    ; DragFinish(hDrop)
    MOV x0, x21
    BL DragFinish

    ADRP x0, g_filePath
    ADD x0, x0, g_filePath
    BL wcslen_p
    MOV x24, x0
    CMP x24, #4
    B.LT wp_drop_set

    ADRP x0, g_filePath
    ADD x0, x0, g_filePath
    ADD x0, x0, x24, LSL #1
    SUB x0, x0, #8
    ADRP x1, str_extLnk
    ADD x1, x1, str_extLnk
    BL wcscmp_ci
    CBZ w0, wp_drop_set             ; not .lnk

    ADRP x0, g_filePath
    ADD x0, x0, g_filePath
    ADRP x1, g_cmdBuf
    ADD x1, x1, g_cmdBuf
    ADRP x2, g_tempBuf
    ADD x2, x2, g_tempBuf
    BL ResolveLnkPath
    CBZ w0, wp_drop_set

    ADRP x0, g_cmdBuf
    ADD x0, x0, g_cmdBuf
    BL wcslen_p
    CBZ x0, wp_drop_args_only
    ADRP x9, g_cmdBuf
    ADD x9, x9, g_cmdBuf
    ADD x9, x9, x0, LSL #1
    MOV w10, #' '
    STRH w10, [x9]
    STRH wzr, [x9, #2]
wp_drop_args_only
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
    B wp_drop_run

wp_drop_set
    ADRP x0, g_hwndEdit
    ADD x0, x0, g_hwndEdit
    LDR x0, [x0]
    ADRP x1, g_filePath
    ADD x1, x1, g_filePath
    BL SetWindowTextW

wp_drop_run
    MOV x0, x19
    BL RunCommand
    MOV w0, #0
    B wp_done

; ------------------------------------------------------------------------------
; WM_KEYDOWN - Esc quits.
; ------------------------------------------------------------------------------
wp_keydown
    CMP w21, #VK_ESCAPE
    B.NE wp_defproc
    MOV w0, #0
    BL PostQuitMessage
    MOV w0, #0
    B wp_done

; ------------------------------------------------------------------------------
; WM_SETTINGCHANGE / WM_THEMECHANGED
; ------------------------------------------------------------------------------
wp_themechange
    MOV x0, x19
    BL ApplyWindowTheme
    MOV w0, #0
    B wp_done

wp_theme_repaint
    ; SetWindowTheme emits WM_THEMECHANGED synchronously; only repaint here to
    ; avoid recursively re-applying the full theme.
    MOV x0, x19
    MOV x1, XZR
    MOV w2, #1
    BL InvalidateRect
    MOV w0, #0
    B wp_done

; ------------------------------------------------------------------------------
; WM_ERASEBKGND (dark mode only)
; ------------------------------------------------------------------------------
wp_erasebkgnd
    ADRP x9, g_darkMode
    ADD x9, x9, g_darkMode
    LDR w9, [x9]
    CBZ w9, wp_defproc
    MOV x0, x19
    ADD x1, sp, #WP_RECT
    BL GetClientRect
    MOV x0, x21                     ; HDC = wParam
    ADD x1, sp, #WP_RECT
    ADRP x2, g_brushBg
    ADD x2, x2, g_brushBg
    LDR x2, [x2]
    BL FillRect
    MOV w0, #1
    B wp_done

; ------------------------------------------------------------------------------
; WM_CTLCOLOREDIT / WM_CTLCOLORLISTBOX (dark mode only)
; ------------------------------------------------------------------------------
wp_ctlcolor_edit
    ADRP x9, g_darkMode
    ADD x9, x9, g_darkMode
    LDR w9, [x9]
    CBZ w9, wp_defproc
    MOV x0, x21                     ; HDC
    MOVZ w1, #0xF2F2
    MOVK w1, #0x00F2, LSL #16       ; 0x00F2F2F2 text
    BL SetTextColor
    MOV x0, x21
    MOVZ w1, #0x2424
    MOVK w1, #0x0024, LSL #16       ; 0x00242424 background
    BL SetBkColor
    ADRP x0, g_brushInput
    ADD x0, x0, g_brushInput
    LDR x0, [x0]
    B wp_done

; ------------------------------------------------------------------------------
; WM_CTLCOLORSTATIC / WM_CTLCOLORBTN (dark mode only)
; ------------------------------------------------------------------------------
wp_ctlcolor_static
    ADRP x9, g_darkMode
    ADD x9, x9, g_darkMode
    LDR w9, [x9]
    CBZ w9, wp_defproc
    MOV x0, x21                     ; HDC
    MOVZ w1, #0xF2F2
    MOVK w1, #0x00F2, LSL #16       ; 0x00F2F2F2 text
    BL SetTextColor
    MOV x0, x21
    MOV w1, #TRANSPARENT
    BL SetBkMode
    ADRP x0, g_brushBg
    ADD x0, x0, g_brushBg
    LDR x0, [x0]
    B wp_done

; ------------------------------------------------------------------------------
; WM_DESTROY
; ------------------------------------------------------------------------------
wp_destroy
    MOV x0, x19
    BL TrayShutdown

    ADRP x0, g_brushBg
    ADD x0, x0, g_brushBg
    LDR x0, [x0]
    CBZ x0, wp_del_surface
    BL DeleteObject
wp_del_surface
    ADRP x0, g_brushSurface
    ADD x0, x0, g_brushSurface
    LDR x0, [x0]
    CBZ x0, wp_del_hover
    BL DeleteObject
wp_del_hover
    ADRP x0, g_brushHover
    ADD x0, x0, g_brushHover
    LDR x0, [x0]
    CBZ x0, wp_del_input
    BL DeleteObject
wp_del_input
    ADRP x0, g_brushInput
    ADD x0, x0, g_brushInput
    LDR x0, [x0]
    CBZ x0, wp_destroy_quit
    BL DeleteObject
wp_destroy_quit
    MOV w0, #0
    BL PostQuitMessage
    MOV w0, #0
    B wp_done

; ------------------------------------------------------------------------------
; Default handler
; ------------------------------------------------------------------------------
wp_defproc
    MOV x0, x19
    MOV w1, w20
    MOV x2, x21
    MOV x3, x22
    BL DefWindowProcW               ; returns LRESULT in x0

; ------------------------------------------------------------------------------
; Shared epilogue (every handler branches here; x0 holds the LRESULT).
; ------------------------------------------------------------------------------
wp_done
    ADD sp, sp, #WPF_SIZE
    LDP x27, x28, [sp], #16
    LDP x25, x26, [sp], #16
    LDP x23, x24, [sp], #16
    LDP x21, x22, [sp], #16
    LDP x19, x20, [sp], #16
    LDP x29, x30, [sp], #16
    RET
WndProc ENDP



    END
