; Continuation of WndProc: layout, input, painting, theme and teardown.

wp_size:
    ; Get client area dimensions
    lea rdx, [rsp+40]           ; RECT buffer
    mov rcx, [rbp-72]
    sub rsp, 32
    call GetClientRect
    add rsp, 32

    ; Extract width and height (already real physical pixels for this DPI --
    ; GetClientRect needs no scaling itself)
    mov esi, dword ptr [rsp+40+8]   ; Width
    mov edi, dword ptr [rsp+40+12]  ; Height

    ; Per-monitor DPI, used to scale every 96-DPI-design margin/size below
    ; (the client rect above is already real pixels; the fixed layout
    ; constants like the 20px combo-box margin are not). Falls back to 96
    ; (no scaling) on pre-Win10-1607 systems.
    sub rsp, 32
    call InitDpiApis
    add rsp, 32
    mov rax, qword ptr g_pGetDpiForWindow
    test rax, rax
    jz wp_size_dpi_fallback
    mov rcx, [rbp-72]
    sub rsp, 32
    call rax
    add rsp, 32
    mov dword ptr [rbp-120], eax
    jmp wp_size_dpi_done
wp_size_dpi_fallback:
    mov dword ptr [rbp-120], 96
wp_size_dpi_done:

    ; Calculate ComboBox width (client width - scaled 20px margin)
    mov eax, 20
    imul eax, dword ptr [rbp-120]
    mov ecx, 96
    cdq
    idiv ecx
    mov r12d, esi
    sub r12d, eax

    ; Resize ComboBox
    mov eax, 10
    imul eax, dword ptr [rbp-120]
    mov ecx, 96
    cdq
    idiv ecx
    mov r15d, eax                ; scaled 10 (shared X and Y)
    mov eax, 25
    imul eax, dword ptr [rbp-120]
    mov ecx, 96
    cdq
    idiv ecx
    sub rsp, 48
    mov dword ptr [rsp+40], 1   ; bRepaint
    mov dword ptr [rsp+32], eax ; Height
    mov r9d, r12d               ; Width
    mov r8d, r15d                ; Y
    mov edx, r15d                ; X
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
    mov eax, 80
    imul eax, dword ptr [rbp-120]
    mov ecx, 96
    cdq
    idiv ecx
    mov ebx, eax                 ; scaled 80 (Browse/Run width)
    mov eax, 45
    imul eax, dword ptr [rbp-120]
    mov ecx, 96
    cdq
    idiv ecx
    mov r13d, eax                 ; scaled 45 (Browse/Run Y)
    mov eax, 25
    imul eax, dword ptr [rbp-120]
    mov ecx, 96
    cdq
    idiv ecx
    sub rsp, 48
    mov dword ptr [rsp+40], 1   ; bRepaint
    mov dword ptr [rsp+32], eax ; Height
    mov r9d, ebx                ; Width
    mov r8d, r13d                ; Y
    mov edx, r15d                ; X
    mov rcx, r14
    call MoveWindow
    add rsp, 48

    ; Calculate Run button X position (client width - combo margin - scaled 70px gap)
    mov eax, 70
    imul eax, dword ptr [rbp-120]
    mov ecx, 96
    cdq
    idiv ecx
    mov r14d, r12d
    sub r14d, eax

    ; Reposition Run button
    mov eax, 25
    imul eax, dword ptr [rbp-120]
    mov ecx, 96
    cdq
    idiv ecx
    sub rsp, 48
    mov dword ptr [rsp+40], 1   ; bRepaint
    mov dword ptr [rsp+32], eax ; Height
    mov r9d, ebx                 ; Width
    mov r8d, r13d                 ; Y
    mov edx, r14d                ; X
    mov rcx, g_hwndBtn
    call MoveWindow
    add rsp, 48

    ; Reposition status label. Keep the same scaled10 gap after Browse and
    ; before Run: X=scaled100, width=RunX-scaled110.
    mov eax, 110
    imul eax, dword ptr [rbp-120]
    mov ecx, 96
    cdq
    idiv ecx
    mov ebx, r14d                ; Run button X
    sub ebx, eax                 ; status width = RunX - scaled110
    mov eax, 48
    imul eax, dword ptr [rbp-120]
    mov ecx, 96
    cdq
    idiv ecx
    mov r13d, eax                 ; scaled 48 Y
    mov eax, 100
    imul eax, dword ptr [rbp-120]
    mov ecx, 96
    cdq
    idiv ecx
    mov r14d, eax                 ; scaled 100 X
    mov eax, 20
    imul eax, dword ptr [rbp-120]
    mov ecx, 96
    cdq
    idiv ecx
    sub rsp, 48
    mov dword ptr [rsp+40], 1   ; bRepaint
    mov dword ptr [rsp+32], eax ; Height
    mov r9d, ebx                 ; Width (equal gap on both sides)
    mov r8d, r13d                 ; Y (aligned with buttons)
    mov edx, r14d                 ; X (after Browse button)
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

wp_theme_repaint:
    ; SetWindowTheme synchronously emits WM_THEMECHANGED. Reapplying the
    ; complete theme from that notification recursively exhausts the stack.
    ; The originating WM_SETTINGCHANGE path already updated every attribute;
    ; this secondary notification only needs a repaint.
    xor r9d, r9d
    mov r8d, 1
    xor edx, edx
    mov rcx, [rbp-72]
    sub rsp, 32
    call InvalidateRect
    add rsp, 32
    xor eax, eax
    jmp wp_done

wp_erasebkgnd:
    cmp dword ptr g_darkMode, 0
    je wp_defproc
    lea rdx, [rbp-128]
    mov rcx, [rbp-72]
    sub rsp, 32
    call GetClientRect
    add rsp, 32
    mov r8, g_brushBg
    lea rdx, [rbp-128]
    mov rcx, [rbp-88]           ; HDC from wParam
    sub rsp, 32
    call FillRect
    add rsp, 32
    mov eax, 1
    jmp wp_done

wp_measureitem:
    cmp dword ptr g_darkMode, 0
    je wp_defproc
    mov rsi, [rbp-96]
    cmp dword ptr [rsi], ODT_MENU
    jne wp_defproc
    mov rax, qword ptr [rsi+24]       ; itemData = label pointer
    test rax, rax
    jz wp_measure_separator
    lea rcx, str_MenuFile
    cmp rax, rcx
    jne wp_measure_popup
    mov dword ptr [rsi+12], 52        ; compact top-level File tab
    mov dword ptr [rsi+16], 30
    jmp wp_measure_done
wp_measure_popup:
    mov dword ptr [rsi+12], 210
    mov dword ptr [rsi+16], 32
    jmp wp_measure_done
wp_measure_separator:
    mov dword ptr [rsi+12], 210
    mov dword ptr [rsi+16], 9
wp_measure_done:
    mov eax, 1
    jmp wp_done

wp_drawitem:
    cmp dword ptr g_darkMode, 0
    je wp_defproc
    mov rsi, [rbp-96]
    cmp dword ptr [rsi], ODT_MENU
    jne wp_defproc
    mov r8, g_brushSurface
    test dword ptr [rsi+16], ODS_SELECTED
    jz @F
    mov r8, g_brushHover
@@:
    lea rdx, [rsi+40]                 ; rcItem
    mov rcx, [rsi+32]                 ; hDC
    sub rsp, 32
    call FillRect
    add rsp, 32
    mov rax, qword ptr [rsi+56]       ; itemData
    test rax, rax
    jz wp_draw_separator
    mov edx, 00F2F2F2h
    test dword ptr [rsi+16], ODS_DISABLED
    jz @F
    mov edx, 00909090h
@@:
    mov rcx, [rsi+32]
    sub rsp, 32
    call SetTextColor
    add rsp, 32
    mov edx, TRANSPARENT
    mov rcx, [rsi+32]
    sub rsp, 32
    call SetBkMode
    add rsp, 32
    add dword ptr [rsi+40], 16        ; Notepad-like left padding
    cmp dword ptr [rsi+8], ID_FILE_HISTORY
    jne wp_draw_text
    cmp dword ptr g_historyEnabled, 0
    je wp_draw_text
    ; Draw check mark in the reserved leading column.
    sub dword ptr [rsi+40], 10
    lea r9, [rsi+40]
    mov r8d, -1
    lea rdx, str_MenuCheck
    mov rcx, [rsi+32]
    sub rsp, 40
    mov dword ptr [rsp+32], DT_SINGLELINE or DT_VCENTER or DT_LEFT
    call DrawTextW
    add rsp, 40
    add dword ptr [rsi+40], 10
wp_draw_text:
    lea r9, [rsi+40]
    mov r8d, -1
    mov rdx, qword ptr [rsi+56]
    mov rcx, [rsi+32]
    sub rsp, 40
    mov dword ptr [rsp+32], DT_SINGLELINE or DT_VCENTER or DT_LEFT
    call DrawTextW
    add rsp, 40
    mov eax, 1
    jmp wp_done
wp_draw_separator:
    ; Surface was already filled; a centered one-pixel hover-color band is
    ; visually equivalent to Notepad's subtle #414141 separator.
    mov eax, dword ptr [rsi+44]
    add eax, dword ptr [rsi+52]       ; rc.top + rc.bottom
    shr eax, 1
    mov dword ptr [rbp-144], eax
    inc eax
    mov dword ptr [rbp-136], eax
    mov eax, dword ptr [rsi+40]
    add eax, 12
    mov dword ptr [rbp-148], eax
    mov eax, dword ptr [rsi+48]
    sub eax, 12
    mov dword ptr [rbp-140], eax
    mov r8, g_brushHover
    lea rdx, [rbp-148]
    mov rcx, [rsi+32]
    sub rsp, 32
    call FillRect
    add rsp, 32
    mov eax, 1
    jmp wp_done

wp_ctlcolor_edit:
    cmp dword ptr g_darkMode, 0
    je wp_defproc
    mov edx, 00F2F2F2h
    mov rcx, [rbp-88]
    sub rsp, 32
    call SetTextColor
    add rsp, 32
    mov edx, 00242424h
    mov rcx, [rbp-88]
    sub rsp, 32
    call SetBkColor
    add rsp, 32
    mov rax, g_brushInput
    jmp wp_done

wp_ctlcolor_static:
    cmp dword ptr g_darkMode, 0
    je wp_defproc
    mov edx, 00F2F2F2h
    mov rcx, [rbp-88]
    sub rsp, 32
    call SetTextColor
    add rsp, 32
    mov edx, TRANSPARENT
    mov rcx, [rbp-88]
    sub rsp, 32
    call SetBkMode
    add rsp, 32
    mov rax, g_brushBg
    jmp wp_done

wp_destroy:
    ; Window is being destroyed
    mov rcx, [rbp-72]
    sub rsp, 32
    call TrayShutdown
    add rsp, 32

    mov rcx, g_brushBg
    test rcx, rcx
    jz @F
    sub rsp, 32
    call DeleteObject
    add rsp, 32
@@:
    mov rcx, g_brushSurface
    test rcx, rcx
    jz @F
    sub rsp, 32
    call DeleteObject
    add rsp, 32
@@:
    mov rcx, g_brushHover
    test rcx, rcx
    jz @F
    sub rsp, 32
    call DeleteObject
    add rsp, 32
@@:
    mov rcx, g_brushInput
    test rcx, rcx
    jz @F
    sub rsp, 32
    call DeleteObject
    add rsp, 32
@@:
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
