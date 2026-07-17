; Included by window.asm; kept architecture-local intentionally.

InitDpiApis proc
    cmp g_dpiApisInit, 0
    jnz idi_done

    invoke GetModuleHandleA, offset str_user32A
    test eax, eax
    jz idi_mark_done             ; user32.dll not found (shouldn't happen) -- leave pointers NULL

    invoke GetProcAddress, eax, offset str_fnGetDpiForSystem
    mov g_pGetDpiForSystem, eax

    invoke GetModuleHandleA, offset str_user32A
    invoke GetProcAddress, eax, offset str_fnGetDpiForWindow
    mov g_pGetDpiForWindow, eax

    invoke GetModuleHandleA, offset str_user32A
    invoke GetProcAddress, eax, offset str_fnAdjustWindowRectExForDpi
    mov g_pAdjustWindowRectExForDpi, eax

idi_mark_done:
    mov g_dpiApisInit, 1

idi_done:
    ret
InitDpiApis endp

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

CreateMainWindow proc uses ebx esi edi hInstance:DWORD
    LOCAL wc[10]:DWORD                      ; WNDCLASSW structure
    LOCAL hIconLarge:DWORD                  ; Large icon handle
    LOCAL hIconSmall:DWORD                  ; Small icon handle
    LOCAL dpi:DWORD                         ; System DPI, scales the window's initial size below
    LOCAL sX:DWORD
    LOCAL sW:DWORD
    LOCAL sH:DWORD
    LOCAL rc[4]:DWORD                       ; RECT for AdjustWindowRectExForDpi

    ; Initialize WNDCLASSW structure
    invoke EnableNativeDarkMode
    invoke EnsureThemeBrushes
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
    ; Falls back to 96 (100%, no scaling) on pre-Win10-1607 systems.
    invoke InitDpiApis
    mov eax, g_pGetDpiForSystem
    test eax, eax
    jz cmw_dpi_fallback
    call eax
    mov dpi, eax
    jmp cmw_dpi_done
cmw_dpi_fallback:
    mov dpi, 96
cmw_dpi_done:

    mov eax, 100
    imul eax, dpi
    mov ecx, 96
    cdq
    idiv ecx
    mov sX, eax                             ; shared X/Y design value

    ; Compute the total window size from the DESIRED CLIENT size via
    ; AdjustWindowRectExForDpi, instead of guessing a total-window pixel
    ; constant -- the caption/menu/border overhead varies by Windows
    ; theme/font scale across machines, and a fixed guess clips the bottom
    ; row of buttons on some of them.
    mov dword ptr rc[0], 0                  ; left
    mov dword ptr rc[4], 0                  ; top
    mov eax, 420
    imul eax, dpi
    mov ecx, 96
    cdq
    idiv ecx
    mov dword ptr rc[8], eax                ; right (scaled desired client width)
    mov eax, 80
    imul eax, dpi
    mov ecx, 96
    cdq
    idiv ecx
    mov dword ptr rc[12], eax               ; bottom (scaled desired client height)

    ; Falls back to a fixed classic non-client-overhead estimate (~16px
    ; width, ~59px height at 96 DPI) when AdjustWindowRectExForDpi isn't
    ; available (pre-Win10-1607). Either path leaves rc[8]/rc[12] ready
    ; for the same width/height extraction below.
    mov eax, g_pAdjustWindowRectExForDpi
    test eax, eax
    jz cmw_awr_fallback
    push dpi
    push 0
    push 1
    push STY_MAINWIN
    lea ecx, rc
    push ecx
    call eax
    jmp cmw_awr_done
cmw_awr_fallback:
    add dword ptr rc[8], 16                 ; right += width overhead
    add dword ptr rc[12], 59                ; bottom += height overhead
cmw_awr_done:

    mov eax, rc[8]
    sub eax, rc[0]                          ; width = right - left
    mov sW, eax
    mov eax, rc[12]
    sub eax, rc[4]                          ; height = bottom - top
    mov sH, eax

    invoke CreateWindowExW, 0, offset str_ClassName, offset str_Title, STY_MAINWIN, sX, sX, sW, sH, 0, 0, hInstance, 0
    test eax, eax
    jz cmw_fail
    mov g_hwndMain, eax

    ; Apply before first ShowWindow to avoid a light-theme composition frame.
    invoke ApplyWindowTheme, eax

    ; Show and update window
    invoke ShowWindow, eax, SW_SHOWNORMAL
    invoke UpdateWindow, g_hwndMain
    mov eax, g_hwndMain
    ret

cmw_fail:
    xor eax, eax                            ; Return NULL on failure
    ret
CreateMainWindow endp

