; Included by window.asm; kept architecture-local intentionally.

ApplyWindowTheme proc uses ebx esi hWnd:DWORD
    LOCAL hKey:DWORD
    LOCAL value:DWORD
    LOCAL cb:DWORD
    cmp g_themeApplying, 0
    jne awt32_return
    mov g_themeApplying, 1
    mov g_darkMode, 0
    mov value, 1
    mov cb, 4
    invoke RegOpenKeyExW, HKEY_CURRENT_USER, offset str_regThemeKey, 0, KEY_READ, addr hKey
    test eax, eax
    jnz awt32_apply
    invoke RegQueryValueExW, hKey, offset str_regAppsUseLightTheme, 0, 0, addr value, addr cb
    invoke RegCloseKey, hKey
    cmp value, 0
    jne awt32_apply
    mov g_darkMode, 1
awt32_apply:
    invoke DwmSetWindowAttribute, hWnd, 20, offset g_darkMode, 4
    cmp g_nativeDarkSupported, 0
    je awt32_controls
    invoke GetModuleHandleA, offset str_uxthemeA
    mov ebx, eax
    test eax, eax
    jz awt32_controls
    invoke GetProcAddress, ebx, 133
    test eax, eax
    jz awt32_controls
    push g_darkMode
    push hWnd
    call eax
    invoke GetProcAddress, ebx, 136
    test eax, eax
    jz awt32_controls
    call eax
awt32_controls:
    mov esi, offset str_Explorer
    cmp g_darkMode, 0
    je @F
    mov esi, offset str_DarkExplorer
@@:
    invoke SetWindowTheme, hWnd, esi, 0
    cmp g_hwndEdit, 0
    je @F
    invoke SetWindowTheme, g_hwndEdit, esi, 0
@@:
    cmp g_hwndBtn, 0
    je @F
    invoke SetWindowTheme, g_hwndBtn, esi, 0
@@:
    invoke GetDlgItem, hWnd, IDC_BTN_BROWSE
    test eax, eax
    jz @F
    invoke SetWindowTheme, eax, esi, 0
@@:
    invoke InvalidateRect, hWnd, 0, 1
    invoke DrawMenuBar, hWnd
    mov g_themeApplying, 0
awt32_return:
    ret
ApplyWindowTheme endp

EnsureThemeBrushes proc
    cmp g_brushBg, 0
    jne etb32_done
    invoke CreateSolidBrush, 00202020h
    mov g_brushBg, eax
    invoke CreateSolidBrush, 002B2B2Bh
    mov g_brushSurface, eax
    invoke CreateSolidBrush, 00383838h
    mov g_brushHover, eax
    invoke CreateSolidBrush, 00242424h
    mov g_brushInput, eax
etb32_done:
    ret
EnsureThemeBrushes endp

EnableNativeDarkMode proc uses ebx esi edi
    LOCAL osv[71]:DWORD
    lea edi, osv
    xor eax, eax
    mov ecx, 71
    rep stosd
    mov osv[0], 284
    invoke GetModuleHandleA, offset str_ntdllA
    test eax, eax
    jz endm32_done
    mov ebx, eax
    invoke GetProcAddress, ebx, offset str_fnRtlGetVersion
    test eax, eax
    jz endm32_done
    mov esi, eax
    lea eax, osv
    push eax
    call esi
    test eax, eax
    js endm32_done
    cmp osv[4], 10
    jb endm32_done
    cmp osv[12], 17763
    jb endm32_done
    invoke GetModuleHandleA, offset str_uxthemeA
    test eax, eax
    jz endm32_done
    mov ebx, eax
    invoke GetProcAddress, ebx, 135
    test eax, eax
    jz endm32_done
    ; Compatible with both AllowDarkModeForApp(TRUE) on 1809 and
    ; SetPreferredAppMode(AllowDark) on 1903+. Never force dark globally.
    push 1
    call eax
    mov g_nativeDarkSupported, 1
    invoke GetProcAddress, ebx, 136
    test eax, eax
    jz endm32_done
    call eax
endm32_done:
    ret
EnableNativeDarkMode endp

