; Included by window.asm; kept architecture-local intentionally.

ApplyWindowTheme proc frame
    push rbx
    .pushreg rbx
    sub rsp, 80
    .allocstack 80
    .endprolog

    cmp dword ptr g_themeApplying, 0
    jne awt_return
    mov dword ptr g_themeApplying, 1

    ; Local variables:
    ; [rsp+32] = hKey (8 bytes)
    ; [rsp+40] = valueData (4 bytes)
    ; [rsp+48] = dataSize (4 bytes)
    ; [rsp+56] = darkFlag (4 bytes)

    mov rbx, rcx                ; Save hWnd
    ; WM_CREATE already sampled the preference before building the controls.
    ; Preserve that known state if a transient registry query later fails.
    mov eax, dword ptr g_darkMode
    mov dword ptr [rsp+56], eax

    ; Open theme registry key
    sub rsp, 48
    lea rax, [rsp+48+32]
    mov [rsp+32], rax           ; phkResult
    mov r9d, KEY_READ
    xor r8d, r8d                ; ulOptions
    lea rdx, str_regThemeKey
    mov ecx, HKEY_CURRENT_USER
    movsxd rcx, ecx             ; predefined HKEY values are signed handles
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
    mov rcx, [rsp+32+32]          ; hKey in the caller's local frame
    call RegCloseKey
    add rsp, 32

awt_apply:
    mov eax, dword ptr [rsp+56]
    mov dword ptr g_darkMode, eax
    cmp dword ptr g_nativeDarkSupported, 0
    je @F
    lea rcx, str_uxthemeA
    call GetModuleHandleA
    test rax, rax
    jz @F
    mov rcx, rax
    mov edx, 133                    ; AllowDarkModeForWindow(HWND, BOOL)
    call GetProcAddress
    test rax, rax
    jz @F
    mov edx, dword ptr g_darkMode
    mov rcx, rbx
    call rax
    lea rcx, str_uxthemeA
    call GetModuleHandleA
    test rax, rax
    jz @F
    mov rcx, rax
    mov edx, 136                    ; rebuild cached menu theme handles
    call GetProcAddress
    test rax, rax
    jz @F
    call rax
@@:
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

    ; Native controls provide their own animated hover/pressed transitions.
    ; Selecting the Explorer dark theme keeps those transitions while our
    ; WM_CTLCOLOR handlers supply the exact Notepad-like client palette.
    lea rax, str_Explorer
    mov qword ptr [rsp+64], rax
    cmp dword ptr g_darkMode, 0
    je @F
    lea rax, str_DarkExplorer
    mov qword ptr [rsp+64], rax
@@:
    xor r8d, r8d
    mov rdx, qword ptr [rsp+64]
    mov rcx, rbx
    sub rsp, 32
    call SetWindowTheme
    add rsp, 32
    xor r8d, r8d
    mov rdx, qword ptr [rsp+64]
    mov rcx, g_hwndEdit
    test rcx, rcx
    jz @F
    sub rsp, 32
    call SetWindowTheme
    add rsp, 32
@@:
    mov edx, IDC_BTN_BROWSE
    mov rcx, rbx
    sub rsp, 32
    call GetDlgItem
    add rsp, 32
    test rax, rax
    jz @F
    xor r8d, r8d
    mov rdx, qword ptr [rsp+64]
    mov rcx, rax
    sub rsp, 32
    call SetWindowTheme
    add rsp, 32
@@:
    xor r8d, r8d
    mov rdx, qword ptr [rsp+64]
    mov rcx, g_hwndBtn
    test rcx, rcx
    jz @F
    sub rsp, 32
    call SetWindowTheme
    add rsp, 32
@@:
    xor r9d, r9d
    mov r8d, 1
    mov edx, 0
    mov rcx, rbx
    sub rsp, 32
    call InvalidateRect
    add rsp, 32
    mov rcx, rbx
    sub rsp, 32
    call DrawMenuBar
    add rsp, 32

    mov dword ptr g_themeApplying, 0
awt_return:
    xor eax, eax
    add rsp, 80
    pop rbx
    ret
ApplyWindowTheme endp

EnsureThemeBrushes proc frame
    sub rsp, 40
    .allocstack 40
    .endprolog
    cmp qword ptr g_brushBg, 0
    jne etb_done
    mov ecx, 00202020h
    call CreateSolidBrush
    mov g_brushBg, rax
    mov ecx, 002B2B2Bh
    call CreateSolidBrush
    mov g_brushSurface, rax
    mov ecx, 00383838h
    call CreateSolidBrush
    mov g_brushHover, rax
    mov ecx, 00242424h
    call CreateSolidBrush
    mov g_brushInput, rax
etb_done:
    add rsp, 40
    ret
EnsureThemeBrushes endp

ApplyDarkMenuBar proc frame
    push rbx
    .pushreg rbx
    sub rsp, 80
    .allocstack 80
    .endprolog
    cmp dword ptr g_darkMode, 0
    je admb_done
    mov rbx, rcx
    call GetMenu
    test rax, rax
    jz admb_done
    mov qword ptr [rsp+32], rax   ; keep clear of MENUINFO at rsp+40..79
    lea r10, [rsp+40]
    xor rax, rax
    mov qword ptr [r10], rax
    mov qword ptr [r10+8], rax
    mov qword ptr [r10+16], rax
    mov qword ptr [r10+24], rax
    mov qword ptr [r10+32], rax
    mov dword ptr [r10], 40
    mov dword ptr [r10+4], 2              ; MIM_BACKGROUND
    mov rax, g_brushSurface
    mov qword ptr [r10+16], rax           ; hbrBack
    lea rdx, [rsp+40]
    mov rcx, qword ptr [rsp+32]
    call SetMenuInfo
    sub rsp, 40
    lea rax, str_MenuFile
    mov [rsp+32], rax
    mov r9, g_hMenuFile
    mov r8d, MF_BYPOSITION or MF_POPUP or MF_OWNERDRAW
    xor edx, edx
    mov rcx, qword ptr [rsp+40+32]
    call ModifyMenuW
    add rsp, 40
    mov rcx, rbx
    call DrawMenuBar
admb_done:
    add rsp, 80
    pop rbx
    ret
ApplyDarkMenuBar endp

EnableNativeDarkMode proc frame
    push rbx
    .pushreg rbx
    sub rsp, 336
    .allocstack 336
    .endprolog
    lea rcx, str_ntdllA
    call GetModuleHandleA
    test rax, rax
    jz endm_done
    mov rbx, rax
    lea rdx, str_fnRtlGetVersion
    mov rcx, rbx
    call GetProcAddress
    test rax, rax
    jz endm_done
    mov qword ptr [rsp+40], 284
    lea r10, [rsp+48]
    mov ecx, 35
    xor edx, edx
endm_zero:
    mov qword ptr [r10], rdx
    add r10, 8
    dec ecx
    jnz endm_zero
    lea rcx, [rsp+40]
    call rax
    test eax, eax
    js endm_done
    cmp dword ptr [rsp+44], 10
    jb endm_done
    cmp dword ptr [rsp+52], 17763
    jb endm_done
    lea rcx, str_uxthemeA
    call GetModuleHandleA
    test rax, rax
    jz endm_done
    mov rbx, rax
    mov edx, 135
    mov rcx, rbx
    call GetProcAddress
    test rax, rax
    jz endm_done
    ; Ordinal 135 is AllowDarkModeForApp(BOOL) on 1809 and
    ; SetPreferredAppMode(enum) on 1903+. Value 1 means TRUE/AllowDark in
    ; both ABIs; value 2 would be ForceDark on newer Windows and would make
    ; popup menus dark even while AppsUseLightTheme requests the light theme.
    mov ecx, 1
    call rax
    mov dword ptr g_nativeDarkSupported, 1
    mov edx, 136              ; FlushMenuThemes
    mov rcx, rbx
    call GetProcAddress
    test rax, rax
    jz endm_done
    call rax
endm_done:
    add rsp, 336
    pop rbx
    ret
EnableNativeDarkMode endp

