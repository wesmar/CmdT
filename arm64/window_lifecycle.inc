; ==============================================================================
; window_lifecycle.inc  -  Included by window.asm (ARM64 / armasm64)
;
; Runtime DPI-API resolution and main-window construction.
; Ported from x64/window_lifecycle.inc; the string helpers wcslen_w /
; wcscmp_ci_w / wcscat_w are replaced by strutil's wcslen_p / wcscmp_ci /
; wcscat_p (identical semantics), so only InitDpiApis and CreateMainWindow
; live here.
; ==============================================================================

; ==============================================================================
; InitDpiApis - resolve GetDpiForSystem / GetDpiForWindow /
; AdjustWindowRectExForDpi once. Left NULL on Windows < 10 1607; every call
; site falls back to fixed 96-DPI behavior.
; ==============================================================================
InitDpiApis PROC
    STP x29, x30, [sp, #-16]!
    MOV x29, sp
    STP x19, x20, [sp, #-16]!

    ADRP x9, g_dpiApisInit
    ADD x9, x9, g_dpiApisInit
    LDR w10, [x9]
    CBNZ w10, idi_done

    ADRP x0, str_user32A
    ADD x0, x0, str_user32A
    BL GetModuleHandleA
    CBZ x0, idi_mark               ; user32 missing (impossible): leave NULL
    MOV x19, x0                     ; x19 = hModule

    ADRP x1, str_fnGetDpiForSystem
    ADD x1, x1, str_fnGetDpiForSystem
    MOV x0, x19
    BL GetProcAddress
    ADRP x9, g_pGetDpiForSystem
    ADD x9, x9, g_pGetDpiForSystem
    STR x0, [x9]

    ADRP x1, str_fnGetDpiForWindow
    ADD x1, x1, str_fnGetDpiForWindow
    MOV x0, x19
    BL GetProcAddress
    ADRP x9, g_pGetDpiForWindow
    ADD x9, x9, g_pGetDpiForWindow
    STR x0, [x9]

    ADRP x1, str_fnAdjustWindowRectExForDpi
    ADD x1, x1, str_fnAdjustWindowRectExForDpi
    MOV x0, x19
    BL GetProcAddress
    ADRP x9, g_pAdjustWindowRectExForDpi
    ADD x9, x9, g_pAdjustWindowRectExForDpi
    STR x0, [x9]

idi_mark
    MOV w10, #1
    ADRP x9, g_dpiApisInit
    ADD x9, x9, g_dpiApisInit
    STR w10, [x9]

idi_done
    LDP x19, x20, [sp], #16
    LDP x29, x30, [sp], #16
    RET
InitDpiApis ENDP

; ==============================================================================
; CreateMainWindow - register the class and create/show the main window.
;   x0 = hInstance
;   returns x0 = hWnd (0 on failure)
;
; Frame (locals; sp fixed after prologue):
;   [sp+CMW_OUT ] outgoing stack args for CreateWindowExW (4 * 8)
;   [sp+CMW_WC  ] WNDCLASSW (72 bytes)
;   [sp+CMW_RECT] RECT for AdjustWindowRectExForDpi
;   [sp+CMW_ICONL/ICONS] icon handles from ExtractIconExW
;   [sp+CMW_W/H ] computed total window width/height
; Register use: x19 = hInstance, x20 = system DPI (survives every call).
; ==============================================================================
CMW_OUT     EQU 0
CMW_WC      EQU 32
CMW_RECT    EQU 104
CMW_ICONL   EQU 120
CMW_ICONS   EQU 128
CMW_W       EQU 136
CMW_H       EQU 140
CMW_FRAME   EQU 160

    EXPORT CreateMainWindow
CreateMainWindow PROC
    STP x29, x30, [sp, #-16]!
    MOV x29, sp
    STP x19, x20, [sp, #-16]!
    SUB sp, sp, #CMW_FRAME

    MOV x19, x0                     ; x19 = hInstance

    ; WM_CREATE fires synchronously inside CreateWindowExW below, so resolve the
    ; DPI APIs and native dark mode up front.
    BL InitDpiApis
    BL EnableNativeDarkMode

    ; System DPI (falls back to 96 = 100%).
    ADRP x9, g_pGetDpiForSystem
    ADD x9, x9, g_pGetDpiForSystem
    LDR x9, [x9]
    CBZ x9, cmw_dpi96
    BLR x9
    MOV w20, w0
    B cmw_dpi_done
cmw_dpi96
    MOV w20, #96
cmw_dpi_done

    ; Zero WNDCLASSW (72 bytes -> 9 qwords), then fill.
    ADD x8, sp, #CMW_WC
    STP xzr, xzr, [x8]
    STP xzr, xzr, [x8, #16]
    STP xzr, xzr, [x8, #32]
    STP xzr, xzr, [x8, #48]
    STR xzr, [x8, #64]

    MOV w9, #(CS_HREDRAW + CS_VREDRAW)
    STR w9, [sp, #CMW_WC + 0]      ; style
    ADRP x9, WndProc
    ADD x9, x9, WndProc
    STR x9, [sp, #CMW_WC + 8]      ; lpfnWndProc
    STR x19, [sp, #CMW_WC + 24]    ; hInstance

    ; ExtractIconExW(str_Shell32, 104, &iconLarge, &iconSmall, 1)
    ADRP x0, str_Shell32
    ADD x0, x0, str_Shell32
    MOV w1, #104
    ADD x2, sp, #CMW_ICONL
    ADD x3, sp, #CMW_ICONS
    MOV w4, #1
    BL ExtractIconExW
    LDR x9, [sp, #CMW_ICONL]
    STR x9, [sp, #CMW_WC + 32]     ; hIcon

    ; LoadCursorW(NULL, IDC_ARROW)
    MOV x0, XZR
    MOV w1, #IDC_ARROW_ATOM
    BL LoadCursorW
    STR x0, [sp, #CMW_WC + 40]     ; hCursor

    ; GetStockObject(WHITE_BRUSH) as the light-mode erase brush.
    MOV w0, #WHITE_BRUSH
    BL GetStockObject
    STR x0, [sp, #CMW_WC + 48]     ; hbrBackground

    ADRP x9, str_ClassName
    ADD x9, x9, str_ClassName
    STR x9, [sp, #CMW_WC + 64]     ; lpszClassName

    ADD x0, sp, #CMW_WC
    BL RegisterClassW
    CBZ w0, cmw_fail

    ; RECT{0,0, scaled(420), scaled(80)} = desired CLIENT size at this DPI.
    STR wzr, [sp, #CMW_RECT + 0]   ; left
    STR wzr, [sp, #CMW_RECT + 4]   ; top
    MOV w0, #420
    MUL w0, w0, w20
    MOV w1, #96
    UDIV w0, w0, w1
    STR w0, [sp, #CMW_RECT + 8]    ; right
    MOV w0, #80
    MUL w0, w0, w20
    MOV w1, #96
    UDIV w0, w0, w1
    STR w0, [sp, #CMW_RECT + 12]   ; bottom

    ; AdjustWindowRectExForDpi(&rect, STY_MAINWIN, TRUE, 0, dpi) if available,
    ; else add the classic ~16px/~59px non-client overhead.
    ADRP x9, g_pAdjustWindowRectExForDpi
    ADD x9, x9, g_pAdjustWindowRectExForDpi
    LDR x9, [x9]
    CBZ x9, cmw_awr_fallback
    ADD x0, sp, #CMW_RECT
    MOVZ w1, #0x10CA, LSL #16      ; STY_MAINWIN
    MOV w2, #1                     ; bMenu = TRUE
    MOV w3, #0                     ; dwExStyle
    MOV w4, w20                    ; dpi
    BLR x9
    B cmw_awr_done
cmw_awr_fallback
    LDR w9, [sp, #CMW_RECT + 8]
    ADD w9, w9, #16
    STR w9, [sp, #CMW_RECT + 8]
    LDR w9, [sp, #CMW_RECT + 12]
    ADD w9, w9, #59
    STR w9, [sp, #CMW_RECT + 12]
cmw_awr_done

    ; width = right - left, height = bottom - top
    LDR w9, [sp, #CMW_RECT + 8]
    LDR w10, [sp, #CMW_RECT + 0]
    SUB w9, w9, w10
    STR w9, [sp, #CMW_W]
    LDR w9, [sp, #CMW_RECT + 12]
    LDR w10, [sp, #CMW_RECT + 4]
    SUB w9, w9, w10
    STR w9, [sp, #CMW_H]

    ; CreateWindowExW(0, class, title, STY_MAINWIN, scaled100, scaled100,
    ;                 width, height, NULL, NULL, hInstance, NULL)
    MOV w9, #100
    MUL w9, w9, w20
    MOV w10, #96
    UDIV w9, w9, w10               ; scaled X/Y (both 100 design)
    MOV x0, XZR                    ; dwExStyle
    ADRP x1, str_ClassName
    ADD x1, x1, str_ClassName
    ADRP x2, str_Title
    ADD x2, x2, str_Title
    MOVZ w3, #0x10CA, LSL #16      ; STY_MAINWIN
    MOV w4, w9                     ; X
    MOV w5, w9                     ; Y
    LDR w6, [sp, #CMW_W]           ; nWidth
    LDR w7, [sp, #CMW_H]           ; nHeight
    STR xzr, [sp, #CMW_OUT + 0]    ; hWndParent
    STR xzr, [sp, #CMW_OUT + 8]    ; hMenu
    STR x19, [sp, #CMW_OUT + 16]   ; hInstance
    STR xzr, [sp, #CMW_OUT + 24]   ; lpParam
    BL CreateWindowExW
    CBZ x0, cmw_fail

    ADRP x9, g_hwndMain
    ADD x9, x9, g_hwndMain
    STR x0, [x9]
    MOV x19, x0                     ; reuse x19 = hWnd from here on

    ; Theme before first ShowWindow to avoid a one-frame bright flash.
    MOV x0, x19
    BL ApplyWindowTheme

    MOV x0, x19
    MOV w1, #SW_SHOWNORMAL
    BL ShowWindow

    MOV x0, x19
    BL UpdateWindow

    MOV x0, x19                     ; return hWnd
    B cmw_done

cmw_fail
    MOV x0, XZR

cmw_done
    ADD sp, sp, #CMW_FRAME
    LDP x19, x20, [sp], #16
    LDP x29, x30, [sp], #16
    RET
CreateMainWindow ENDP



    END
