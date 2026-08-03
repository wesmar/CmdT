; ==============================================================================
; window_theme.inc  -  Included by window.asm (ARM64 / armasm64)
;
; Carries three responsibilities for the whole window translation unit:
;   1. The shared EQU constant block (Win32 messages, styles, menu/registry
;      flags, control IDs) used across every window_*.inc fragment.
;   2. IMPORT declarations for the globals that physically live in main.asm.
;   3. The theme engine: ApplyWindowTheme, EnsureThemeBrushes,
;      EnableNativeDarkMode.
;
; Ported from x64/window_theme.inc. Behavior is byte-for-byte equivalent; the
; ARM64 form keeps hot state in registers and uses the barrel shifter for
; address math where it helps.
; ==============================================================================

; ------------------------------------------------------------------------------
; Cross-module globals (defined + exported by main.asm)
; ------------------------------------------------------------------------------
    IMPORT g_hwndMain
    IMPORT g_hwndEdit
    IMPORT g_hwndBtn
    IMPORT g_hwndStatus
    IMPORT g_hMenuFile
    IMPORT g_historyEnabled
    IMPORT g_cmdBuf
    IMPORT g_tempBuf
    IMPORT g_filePath

; String helpers (defined + exported by strutil.asm). They have the same
; semantics the x64 window module's private wcslen_w/wcscmp_ci_w/wcscat_w had,
; so the window code reuses them directly instead of carrying local copies.
    IMPORT wcslen_p
    IMPORT wcscmp_ci
    IMPORT wcscat_p

; ------------------------------------------------------------------------------
; Registry
; ------------------------------------------------------------------------------
HKEY_CURRENT_USER   EQU 0x80000001
KEY_READ            EQU 0x00020019
KEY_WRITE           EQU 0x00020006
REG_SZ              EQU 1

; ------------------------------------------------------------------------------
; Window messages
; ------------------------------------------------------------------------------
WM_CREATE           EQU 0x0001
WM_DESTROY          EQU 0x0002
WM_SIZE             EQU 0x0005
WM_ERASEBKGND       EQU 0x0014
WM_SETTINGCHANGE    EQU 0x001A
WM_CHAR             EQU 0x0102
WM_KEYDOWN          EQU 0x0100
WM_COMMAND          EQU 0x0111
WM_THEMECHANGED     EQU 0x031A
WM_DROPFILES        EQU 0x0233
WM_COPYGLOBALDATA   EQU 0x0049
WM_CTLCOLOREDIT     EQU 0x0133
WM_CTLCOLORLISTBOX  EQU 0x0134
WM_CTLCOLORBTN      EQU 0x0135
WM_CTLCOLORSTATIC   EQU 0x0138

; ------------------------------------------------------------------------------
; DWM / theme
; ------------------------------------------------------------------------------
DWMWA_USE_IMMERSIVE_DARK_MODE     EQU 20
DWMWA_USE_IMMERSIVE_DARK_MODE_OLD EQU 19
DWMWA_SYSTEMBACKDROP_TYPE         EQU 38
DWMSBT_MAINWINDOW                 EQU 2
TRANSPARENT                       EQU 1

; ------------------------------------------------------------------------------
; Window / control styles
; ------------------------------------------------------------------------------
WS_EX_CLIENTEDGE    EQU 0x00000200
STY_COMBO           EQU 0x50010042
STY_BUTTON          EQU 0x50010000
STY_BUTTON_DEF      EQU 0x50010001
STY_STATIC          EQU 0x50000200
STY_MAINWIN         EQU 0x10CA0000
CS_HREDRAW          EQU 0x0002
CS_VREDRAW          EQU 0x0001
WHITE_BRUSH         EQU 0
IDC_ARROW_ATOM      EQU 32512
SW_SHOWNORMAL       EQU 1

; ------------------------------------------------------------------------------
; Control notifications / messages
; ------------------------------------------------------------------------------
EN_SETFOCUS         EQU 0x0100
BM_SETSTYLE         EQU 0x00F4
BS_DEFPUSHBUTTON    EQU 0x0001
CB_ADDSTRING        EQU 0x0143
CB_DELETESTRING     EQU 0x0144
CB_GETCOUNT         EQU 0x0146
CB_INSERTSTRING     EQU 0x014A
CB_SETCURSEL        EQU 0x014E

; ------------------------------------------------------------------------------
; Menu / dialog / misc flags
; ------------------------------------------------------------------------------
MF_STRING           EQU 0x00000000
MF_POPUP            EQU 0x00000010
MF_SEPARATOR        EQU 0x00000800
MF_BYCOMMAND        EQU 0x00000000
MF_CHECKED          EQU 0x00000008
MF_UNCHECKED        EQU 0x00000000
OFN_FILEMUSTEXIST   EQU 0x00001000
OFN_PATHMUSTEXIST   EQU 0x00000800
MB_ICONERROR        EQU 0x00000010
MB_ICONINFORMATION  EQU 0x00000040
MSGFLT_ALLOW        EQU 1
CLSCTX_INPROC_SERVER EQU 1
VK_ESCAPE           EQU 0x1B
MRU_MAX_ITEMS       EQU 5

; ------------------------------------------------------------------------------
; Control / menu command IDs
; ------------------------------------------------------------------------------
IDC_COMBO           EQU 101
IDC_EDIT            EQU 101
IDC_BTN_RUN         EQU 102
IDC_STATIC_STATUS   EQU 103
IDC_BTN_BROWSE      EQU 104
ID_FILE_BROWSE      EQU 201
ID_FILE_EXIT        EQU 203
ID_FILE_ABOUT       EQU 204
ID_FILE_HISTORY     EQU 205

; ------------------------------------------------------------------------------
; WndProc stack frame (shared by window_dispatch.inc + window_events.inc).
; hWnd/uMsg/wParam/lParam live in x19/w20/x21/x22 for the whole procedure;
; the frame below holds only structures too large for registers.
; ------------------------------------------------------------------------------
WP_OUT      EQU 0        ; [0..31]   outgoing stack args (CreateWindowExW etc.)
WP_OFN      EQU 32       ; [32..183] OPENFILENAMEW (152 bytes)
WP_RECT     EQU 184      ; [184..199] RECT
WP_HKEY     EQU 200      ; history/MRU HKEY
WP_HKEY2    EQU 208      ; theme HKEY
WP_DW1      EQU 216      ; AppsUseLightTheme value
WP_DW2      EQU 220      ; query data size
WP_DPI      EQU 224      ; per-window DPI (design coords are 96-DPI)
WP_DARKSV   EQU 228      ; dark-mode value saved across menu construction
WPF_SIZE    EQU 240      ; total locals (16-aligned)

; ==============================================================================
; ApplyWindowTheme - read the OS light/dark preference and apply it to the
; window: immersive dark title bar, Mica backdrop, Explorer control theme.
;   x0 = hWnd
; Guards against the synchronous WM_THEMECHANGED that SetWindowTheme emits.
; ==============================================================================
ApplyWindowTheme PROC
    STP x29, x30, [sp, #-16]!
    MOV x29, sp
    STP x19, x20, [sp, #-16]!
    SUB sp, sp, #0x40

    ; Re-entry guard (g_themeApplying). Set inside the frame so awt_return can
    ; clear it on every exit path.
    ADRP x9, g_themeApplying
    ADD x9, x9, g_themeApplying
    LDR w10, [x9]
    CBNZ w10, awt_return
    MOV w10, #1
    STR w10, [x9]

    MOV x19, x0                     ; x19 = hWnd

    ; Candidate = current g_darkMode; a failed registry read keeps this value.
    ADRP x9, g_darkMode
    ADD x9, x9, g_darkMode
    LDR w10, [x9]
    STR w10, [sp, #0x38]            ; [sp+0x38] = candidate darkFlag

    ; RegOpenKeyExW(HKCU, str_regThemeKey, 0, KEY_READ, &hKey)
    MOVZ w0, #0x0001
    MOVK w0, #0x8000, LSL #16       ; HKEY_CURRENT_USER
    ADRP x1, str_regThemeKey
    ADD x1, x1, str_regThemeKey
    MOV x2, XZR
    MOVZ w3, #0x0019
    MOVK w3, #0x0002, LSL #16   ; KEY_READ = 0x00020019
    ADD x4, sp, #0x28              ; &hKey (local at sp+0x28)
    BL RegOpenKeyExW
    CBNZ w0, awt_apply             ; key missing: apply candidate as-is

    ; RegQueryValueExW(hKey, str_regAppsUseLightTheme, NULL, NULL, &data, &size)
    MOV w8, #1
    STR w8, [sp, #0x20]           ; default value = 1 (light) if read fails
    MOV w8, #4
    STR w8, [sp, #0x24]           ; size = sizeof(DWORD)
    LDR x0, [sp, #0x28]
    ADRP x1, str_regAppsUseLightTheme
    ADD x1, x1, str_regAppsUseLightTheme
    MOV x2, XZR
    MOV x3, XZR
    ADD x4, sp, #0x20             ; &data
    ADD x5, sp, #0x24             ; &size
    BL RegQueryValueExW
    CBNZ w0, awt_close

    ; A successful read is authoritative in both directions: reset candidate to
    ; light, then set dark only when AppsUseLightTheme == 0.
    STR wzr, [sp, #0x38]
    LDR w8, [sp, #0x20]
    CBNZ w8, awt_close
    MOV w8, #1
    STR w8, [sp, #0x38]

awt_close
    LDR x0, [sp, #0x28]
    BL RegCloseKey

awt_apply
    LDR w8, [sp, #0x38]
    ADRP x9, g_darkMode
    ADD x9, x9, g_darkMode
    STR w8, [x9]                   ; g_darkMode = candidate

    ; Native dark menus: AllowDarkModeForWindow(hWnd, dark) then FlushMenuThemes.
    ADRP x9, g_nativeDarkSupported
    ADD x9, x9, g_nativeDarkSupported
    LDR w9, [x9]
    CBZ w9, awt_dwm

    ADRP x0, str_uxthemeA
    ADD x0, x0, str_uxthemeA
    BL GetModuleHandleA
    CBZ x0, awt_dwm                 ; x0 = hModule (uxtheme)
    MOV w1, #133                    ; AllowDarkModeForWindow ordinal
    BL GetProcAddress
    CBZ x0, awt_dwm
    MOV x9, x0                      ; fn ptr
    MOV x0, x19                     ; hWnd
    ADRP x1, g_darkMode
    ADD x1, x1, g_darkMode
    LDR w1, [x1]                    ; dark flag
    BLR x9

    ADRP x0, str_uxthemeA
    ADD x0, x0, str_uxthemeA
    BL GetModuleHandleA
    CBZ x0, awt_dwm
    MOV w1, #136                    ; FlushMenuThemes ordinal
    BL GetProcAddress
    CBZ x0, awt_dwm
    BLR x0

awt_dwm
    ; DwmSetWindowAttribute(hWnd, DWMWA_USE_IMMERSIVE_DARK_MODE, &dark, 4)
    LDR w8, [sp, #0x38]
    STR w8, [sp, #0x30]           ; attribute value scratch
    MOV x0, x19
    MOV w1, #DWMWA_USE_IMMERSIVE_DARK_MODE
    ADD x2, sp, #0x30
    MOV w3, #4
    BL DwmSetWindowAttribute
    TBZ w0, #31, awt_backdrop      ; HRESULT >= 0 (S bit clear): success

    ; Fallback for older Windows 10 builds (attribute 19).
    LDR w8, [sp, #0x38]
    STR w8, [sp, #0x30]
    MOV x0, x19
    MOV w1, #DWMWA_USE_IMMERSIVE_DARK_MODE_OLD
    ADD x2, sp, #0x30
    MOV w3, #4
    BL DwmSetWindowAttribute

awt_backdrop
    ; DwmSetWindowAttribute(hWnd, DWMWA_SYSTEMBACKDROP_TYPE, &DWMSBT_MAINWINDOW, 4)
    MOV w8, #DWMSBT_MAINWINDOW
    STR w8, [sp, #0x30]
    MOV x0, x19
    MOV w1, #DWMWA_SYSTEMBACKDROP_TYPE
    ADD x2, sp, #0x30
    MOV w3, #4
    BL DwmSetWindowAttribute

    ; Choose Explorer vs DarkMode_Explorer theme, store the pointer in x20.
    ADRP x20, str_Explorer
    ADD x20, x20, str_Explorer
    ADRP x9, g_darkMode
    ADD x9, x9, g_darkMode
    LDR w9, [x9]
    CBZ w9, awt_settheme
    ADRP x20, str_DarkExplorer
    ADD x20, x20, str_DarkExplorer

awt_settheme
    ; SetWindowTheme(hWnd, themeName, NULL)
    MOV x0, x19
    MOV x1, x20
    MOV x2, XZR
    BL SetWindowTheme

    ; SetWindowTheme(g_hwndEdit, themeName, NULL)
    ADRP x0, g_hwndEdit
    ADD x0, x0, g_hwndEdit
    LDR x0, [x0]
    CBZ x0, awt_theme_browse
    MOV x1, x20
    MOV x2, XZR
    BL SetWindowTheme

awt_theme_browse
    ; SetWindowTheme(GetDlgItem(hWnd, IDC_BTN_BROWSE), themeName, NULL)
    MOV x0, x19
    MOV w1, #IDC_BTN_BROWSE
    BL GetDlgItem
    CBZ x0, awt_theme_run
    MOV x1, x20
    MOV x2, XZR
    BL SetWindowTheme

awt_theme_run
    ; SetWindowTheme(g_hwndBtn, themeName, NULL)
    ADRP x0, g_hwndBtn
    ADD x0, x0, g_hwndBtn
    LDR x0, [x0]
    CBZ x0, awt_repaint
    MOV x1, x20
    MOV x2, XZR
    BL SetWindowTheme

awt_repaint
    ; InvalidateRect(hWnd, NULL, TRUE)
    MOV x0, x19
    MOV x1, XZR
    MOV w2, #1
    BL InvalidateRect
    ; DrawMenuBar(hWnd)
    MOV x0, x19
    BL DrawMenuBar

    ; Clear the re-entry guard.
    ADRP x9, g_themeApplying
    ADD x9, x9, g_themeApplying
    STR wzr, [x9]

awt_return
    MOV w0, #0
    ADD sp, sp, #0x40
    LDP x19, x20, [sp], #16
    LDP x29, x30, [sp], #16
    RET
ApplyWindowTheme ENDP

; ==============================================================================
; EnsureThemeBrushes - lazily create the four dark-palette GDI brushes.
; ==============================================================================
EnsureThemeBrushes PROC
    STP x29, x30, [sp, #-16]!
    MOV x29, sp

    ADRP x9, g_brushBg
    ADD x9, x9, g_brushBg
    LDR x10, [x9]
    CBNZ x10, etb_done             ; already created

    MOVZ w0, #0x2020
    MOVK w0, #0x0020, LSL #16       ; 0x00202020 main background
    BL CreateSolidBrush
    ADRP x9, g_brushBg
    ADD x9, x9, g_brushBg
    STR x0, [x9]

    MOVZ w0, #0x2B2B
    MOVK w0, #0x002B, LSL #16       ; 0x002B2B2B surface
    BL CreateSolidBrush
    ADRP x9, g_brushSurface
    ADD x9, x9, g_brushSurface
    STR x0, [x9]

    MOVZ w0, #0x3838
    MOVK w0, #0x0038, LSL #16       ; 0x00383838 hover
    BL CreateSolidBrush
    ADRP x9, g_brushHover
    ADD x9, x9, g_brushHover
    STR x0, [x9]

    MOVZ w0, #0x2424
    MOVK w0, #0x0024, LSL #16       ; 0x00242424 input background
    BL CreateSolidBrush
    ADRP x9, g_brushInput
    ADD x9, x9, g_brushInput
    STR x0, [x9]

etb_done
    LDP x29, x30, [sp], #16
    RET
EnsureThemeBrushes ENDP

; ==============================================================================
; EnableNativeDarkMode - probe RtlGetVersion for Windows 10 1809+, then enable
; app-wide dark menus via uxtheme ordinal 135 (+ FlushMenuThemes ordinal 136).
; Sets g_nativeDarkSupported on success.
; ==============================================================================
ENDM_FRAME  EQU 0x140              ; room for OSVERSIONINFOEXW (284 bytes)
ENDM_OSVI   EQU 0x10               ; struct base within frame

EnableNativeDarkMode PROC
    STP x29, x30, [sp, #-16]!
    MOV x29, sp
    STP x19, x20, [sp, #-16]!
    SUB sp, sp, #ENDM_FRAME

    ADRP x0, str_ntdllA
    ADD x0, x0, str_ntdllA
    BL GetModuleHandleA
    CBZ x0, endm_done
    MOV x19, x0
    ADRP x1, str_fnRtlGetVersion
    ADD x1, x1, str_fnRtlGetVersion
    MOV x0, x19
    BL GetProcAddress
    CBZ x0, endm_done
    MOV x20, x0                     ; RtlGetVersion

    ; Zero the 284-byte OSVERSIONINFOEXW, then set dwOSVersionInfoSize.
    ADD x8, sp, #ENDM_OSVI
    MOV w9, #36                     ; 36 * 8 = 288 bytes (covers 284)
endm_zero
    STR xzr, [x8], #8
    SUBS w9, w9, #1
    B.NE endm_zero

    ADD x8, sp, #ENDM_OSVI
    MOV w9, #284
    STR w9, [x8]                    ; dwOSVersionInfoSize
    MOV x0, x8
    BLR x20                         ; RtlGetVersion(&osvi)
    TBNZ w0, #31, endm_done         ; negative NTSTATUS: bail

    ; Require major >= 10 and build >= 17763 (1809).
    LDR w9, [sp, #ENDM_OSVI + 4]    ; dwMajorVersion
    CMP w9, #10
    B.LO endm_done
    LDR w9, [sp, #ENDM_OSVI + 12]   ; dwBuildNumber
    MOV w10, #17763
    CMP w9, w10
    B.LO endm_done

    ADRP x0, str_uxthemeA
    ADD x0, x0, str_uxthemeA
    BL GetModuleHandleA
    CBZ x0, endm_done
    MOV x19, x0
    MOV w1, #135                    ; SetPreferredAppMode / AllowDarkModeForApp
    BL GetProcAddress
    CBZ x0, endm_done
    MOV x9, x0
    MOV w0, #1                      ; TRUE / AllowDark (1 in both ABIs)
    BLR x9

    MOV w8, #1
    ADRP x9, g_nativeDarkSupported
    ADD x9, x9, g_nativeDarkSupported
    STR w8, [x9]

    MOV x0, x19
    MOV w1, #136                    ; FlushMenuThemes
    BL GetProcAddress
    CBZ x0, endm_done
    BLR x0

endm_done
    ADD sp, sp, #ENDM_FRAME
    LDP x19, x20, [sp], #16
    LDP x29, x30, [sp], #16
    RET
EnableNativeDarkMode ENDP



    END
