; Continuation of WndProc: layout, input, painting, theme and teardown.

wp_size:
    ; ===== WM_SIZE: Resize controls to fit window =====
    invoke GetClientRect, hWnd, addr rect
    mov esi, [rect+8]                       ; Width (already real pixels)
    mov edi, [rect+12]                      ; Height (already real pixels)

    ; Per-monitor DPI, used to scale every 96-DPI-design margin/size below.
    ; Falls back to 96 (no scaling) on pre-Win10-1607 systems.
    invoke InitDpiApis
    mov eax, g_pGetDpiForWindow
    test eax, eax
    jz wp_size_dpi_fallback
    push hWnd
    call eax
    mov dpi, eax
    jmp wp_size_dpi_done
wp_size_dpi_fallback:
    mov dpi, 96
wp_size_dpi_done:

    ; Row 1: ComboBox (spans full width with margins)
    ; Position: (scaled10, scaled10), Size: (Width - scaled20, scaled25)
    mov eax, 20
    imul eax, dpi
    mov ecx, 96
    cdq
    idiv ecx
    mov ebx, esi
    sub ebx, eax                            ; ebx = combo width

    mov eax, 10
    imul eax, dpi
    mov ecx, 96
    cdq
    idiv ecx
    mov sX, eax
    mov sY, eax

    mov eax, 25
    imul eax, dpi
    mov ecx, 96
    cdq
    idiv ecx
    mov sH, eax                             ; shared row-1/row-2 height

    invoke MoveWindow, g_hwndEdit, sX, sY, ebx, sH, 1

    ; Row 2: Browse button, Status label, Run button
    ; All aligned at y=scaled45

    ; Browse Button: Left-aligned. Position: (scaled10, scaled45), Size: (scaled80, scaled25)
    invoke GetDlgItem, hWnd, IDC_BTN_BROWSE
    mov hBtnBrowse, eax

    mov eax, 45
    imul eax, dpi
    mov ecx, 96
    cdq
    idiv ecx
    mov sY, eax                             ; shared Browse/Run row Y

    mov eax, 80
    imul eax, dpi
    mov ecx, 96
    cdq
    idiv ecx
    mov sW, eax                             ; shared Browse/Run width

    invoke MoveWindow, hBtnBrowse, sX, sY, sW, sH, 1

    ; Run Button: Right-aligned. Position: (Width - scaled90, scaled45)
    mov eax, 90
    imul eax, dpi
    mov ecx, 96
    cdq
    idiv ecx
    mov ebx, esi
    sub ebx, eax                            ; ebx = Run X
    push ebx                                ; Save X position
    invoke MoveWindow, g_hwndBtn, ebx, sY, sW, sH, 1
    pop ebx                                 ; Restore X position

    ; Status Label: Centered between Browse and Run buttons
    ; Position: (scaled100, scaled48), Width: RunBtnX - scaled110
    mov eax, 110
    imul eax, dpi
    mov ecx, 96
    cdq
    idiv ecx
    sub ebx, eax                            ; ebx = status width

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

    mov eax, 20
    imul eax, dpi
    mov ecx, 96
    cdq
    idiv ecx
    mov sH, eax

    invoke MoveWindow, g_hwndStatus, sX, sY, ebx, sH, 1

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

wp_erasebkgnd:
    cmp g_darkMode, 0
    je wp_defproc
    invoke GetClientRect, hWnd, addr rect
    invoke FillRect, wParam, addr rect, g_brushBg
    mov eax, 1
    ret

wp_ctlcolor_edit:
    cmp g_darkMode, 0
    je wp_defproc
    invoke SetTextColor, wParam, 00F2F2F2h
    invoke SetBkColor, wParam, 00242424h
    mov eax, g_brushInput
    ret

wp_ctlcolor_static:
    cmp g_darkMode, 0
    je wp_defproc
    invoke SetTextColor, wParam, 00F2F2F2h
    invoke SetBkMode, wParam, TRANSPARENT
    mov eax, g_brushBg
    ret

wp_themechange:
    invoke ApplyWindowTheme, hWnd
    xor eax, eax
    ret

wp_theme_repaint:
    ; SetWindowTheme may synchronously send WM_THEMECHANGED. The full theme
    ; was already applied by WM_SETTINGCHANGE; repaint without re-entering it.
    invoke InvalidateRect, hWnd, 0, 1
    xor eax, eax
    ret

wp_destroy:
    ; ===== WM_DESTROY: Window is being destroyed =====
    invoke TrayShutdown, hWnd

    cmp g_brushBg, 0
    je @F
    invoke DeleteObject, g_brushBg
@@:
    cmp g_brushInput, 0
    je @F
    invoke DeleteObject, g_brushInput
@@:
    cmp g_brushSurface, 0
    je @F
    invoke DeleteObject, g_brushSurface
@@:
    cmp g_brushHover, 0
    je @F
    invoke DeleteObject, g_brushHover
@@:
    invoke PostQuitMessage, 0
    xor eax, eax
    ret

wp_defproc:
    ; Unhandled message - pass to default window procedure
    invoke DefWindowProcW, hWnd, uMsg, wParam, lParam
    ret
WndProc endp
