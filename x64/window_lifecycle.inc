; Included by window.asm; kept architecture-local intentionally.

InitDpiApis proc frame
    push rbx
    .pushreg rbx
    sub rsp, 32
    .allocstack 32
    .endprolog

    cmp dword ptr g_dpiApisInit, 0
    jnz idi_done

    lea rcx, str_user32A
    call GetModuleHandleA
    test rax, rax
    jz idi_mark_done             ; user32.dll not found (shouldn't happen) -- leave pointers NULL
    mov rbx, rax                 ; rbx = hModule

    lea rdx, str_fnGetDpiForSystem
    mov rcx, rbx
    call GetProcAddress
    mov qword ptr g_pGetDpiForSystem, rax

    lea rdx, str_fnGetDpiForWindow
    mov rcx, rbx
    call GetProcAddress
    mov qword ptr g_pGetDpiForWindow, rax

    lea rdx, str_fnAdjustWindowRectExForDpi
    mov rcx, rbx
    call GetProcAddress
    mov qword ptr g_pAdjustWindowRectExForDpi, rax

idi_mark_done:
    mov dword ptr g_dpiApisInit, 1

idi_done:
    add rsp, 32
    pop rbx
    ret
InitDpiApis endp

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

CreateMainWindow proc frame
    push rbx
    .pushreg rbx
    push rsi
    .pushreg rsi
    push rdi
    .pushreg rdi
    push r12
    .pushreg r12
    sub rsp, 216
    .allocstack 216
    .endprolog

    mov r12, rcx                ; Save hInstance

    ; Resolve the DPI APIs once, up front -- WM_CREATE fires synchronously
    ; inside CreateWindowExW below, so by the time it runs the pointers
    ; InitDpiApis fills in are already valid for wp_create to use too.
    sub rsp, 32
    call InitDpiApis
    add rsp, 32

    sub rsp, 32
    call EnableNativeDarkMode
    add rsp, 32

    ; System DPI, used to scale the main window's creation size below.
    ; esi is callee-saved and otherwise unused in this proc, so it survives
    ; every call between here and the CreateWindowExW that needs it. Falls
    ; back to 96 (100%, no scaling) on pre-Win10-1607 systems.
    mov rax, qword ptr g_pGetDpiForSystem
    test rax, rax
    jz cmw_dpi_fallback
    sub rsp, 32
    call rax
    add rsp, 32
    mov esi, eax                ; esi = DPI (96 = 100%)
    jmp cmw_dpi_done
cmw_dpi_fallback:
    mov esi, 96
cmw_dpi_done:

    ; Initialize WNDCLASSW structure (zero it out)
    lea rdi, [rsp+40]
    xor rax, rax
    mov rcx, 10                 ; 10 QWORDs = 80 bytes (covers 72-byte WNDCLASSW)
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
    ; Store extracted handles at [rsp+120] and [rsp+128] (safe local space)
    sub rsp, 32
    mov dword ptr [rsp+32], 1   ; nIcons
    lea r9, [rsp+32+128]        ; phiconSmall
    lea r8, [rsp+32+120]        ; phiconLarge
    mov edx, 104                ; nIconIndex
    lea rcx, str_Shell32        ; lpszFile
    call ExtractIconExW
    add rsp, 32

    ; Set icon in window class
    mov rax, [rsp+120]          ; phiconLarge
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

    ; Compute the total window size from the DESIRED CLIENT size via
    ; AdjustWindowRectExForDpi, instead of guessing a total-window pixel
    ; constant. A fixed constant has to guess the caption/menu/border
    ; overhead, which varies by Windows theme/font scale across machines --
    ; guess too small and the bottom row of buttons gets clipped (as seen
    ; on a real high-DPI laptop), guess too large and there's dead space.
    ; This asks Windows for the exact overhead at this DPI instead.
    ; RECT{0,0,scaledClientW,scaledClientH} at [rsp+176..191]; saved final
    ; width/height at [rsp+200]/[rsp+204] survive the CreateWindowExW
    ; call-frame's own "sub rsp,96" below since that only lowers rsp --
    ; the underlying memory at rsp0+200/204 is unaffected, just re-addressed
    ; as [rsp+96+200]/[rsp+96+204] afterward.
    mov dword ptr [rsp+176], 0  ; left
    mov dword ptr [rsp+180], 0  ; top
    mov eax, 420
    imul eax, esi
    mov ecx, 96
    cdq
    idiv ecx
    mov dword ptr [rsp+184], eax ; right (scaled desired client width)
    mov eax, 80
    imul eax, esi
    mov ecx, 96
    cdq
    idiv ecx
    mov dword ptr [rsp+188], eax ; bottom (scaled desired client height)

    ; Falls back to a fixed classic non-client-overhead estimate (~16px
    ; width, ~59px height at 96 DPI -- caption + menu + thin borders,
    ; matching this window's original pre-DPI-fix hardcoded size) when
    ; AdjustWindowRectExForDpi isn't available (pre-Win10-1607). Either
    ; path leaves right/bottom in [rsp+184]/[rsp+188] ready for the same
    ; width/height extraction below.
    mov rax, qword ptr g_pAdjustWindowRectExForDpi
    test rax, rax
    jz cmw_awr_fallback
    lea rcx, [rsp+176]           ; lpRect
    mov edx, STY_MAINWIN         ; dwStyle
    mov r8d, 1                   ; bMenu = TRUE (window has a menu bar)
    xor r9d, r9d                 ; dwExStyle
    sub rsp, 48
    mov dword ptr [rsp+32], esi  ; dpi
    call rax
    add rsp, 48
    jmp cmw_awr_done
cmw_awr_fallback:
    add dword ptr [rsp+184], 16  ; right += width overhead
    add dword ptr [rsp+188], 59  ; bottom += height overhead
cmw_awr_done:

    mov eax, dword ptr [rsp+184]
    sub eax, dword ptr [rsp+176] ; width = right - left
    mov dword ptr [rsp+200], eax
    mov eax, dword ptr [rsp+188]
    sub eax, dword ptr [rsp+180] ; height = bottom - top
    mov dword ptr [rsp+204], eax

    ; Create main window
    sub rsp, 96
    xor rax, rax
    mov [rsp+88], rax           ; lpParam
    mov [rsp+80], r12           ; hInstance
    xor rax, rax
    mov [rsp+72], rax           ; hMenu
    xor rax, rax
    mov [rsp+64], rax           ; hWndParent
    mov eax, dword ptr [rsp+96+204]
    mov dword ptr [rsp+56], eax ; nHeight (from AdjustWindowRectExForDpi)
    mov eax, dword ptr [rsp+96+200]
    mov dword ptr [rsp+48], eax ; nWidth (from AdjustWindowRectExForDpi)
    mov eax, 100
    imul eax, esi
    mov ecx, 96
    cdq
    idiv ecx
    mov dword ptr [rsp+40], eax ; Y
    mov dword ptr [rsp+32], eax ; X (same 100 design value)
    mov r9d, STY_MAINWIN        ; dwStyle
    lea r8, str_Title           ; lpWindowName
    lea rdx, str_ClassName      ; lpClassName
    xor ecx, ecx                ; dwExStyle
    call CreateWindowExW
    add rsp, 96
    test rax, rax
    jz cmw_fail                 ; Creation failed
    mov g_hwndMain, rax         ; Save window handle

    ; Apply the theme before the first ShowWindow. This removes the bright
    ; client/title flash that otherwise appears for one composition frame.
    mov rcx, rax
    sub rsp, 32
    call ApplyWindowTheme
    add rsp, 32

    ; Show window
    mov edx, SW_SHOWNORMAL
    mov rcx, g_hwndMain
    sub rsp, 32
    call ShowWindow
    add rsp, 32

    ; Update window display
    mov rcx, g_hwndMain
    sub rsp, 32
    call UpdateWindow
    add rsp, 32

    mov rax, g_hwndMain         ; Return window handle
    jmp cmw_done

cmw_fail:
    xor rax, rax                ; Return NULL on failure

cmw_done:
    add rsp, 216
    pop r12
    pop rdi
    pop rsi
    pop rbx
    ret
CreateMainWindow endp

