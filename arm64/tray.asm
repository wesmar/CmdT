; ==============================================================================
; CMDT - System Tray Integration (ARM64 / armasm64)
;
; Author: Marek Wesolowski (wesmar)
; Ported to: Native ARM64 (armasm64) for Windows on ARM
;
; Owns all notification-area state and behavior. Holding Shift while the main
; window is minimized hides it to the tray. Double-click or Restore brings it
; back; the context menu also provides Exit. The icon is republished after an
; Explorer restart, including across the elevated-process UIPI boundary.
; ==============================================================================

; ==============================================================================
; EXTERNAL REFERENCES
; ==============================================================================
    IMPORT Shell_NotifyIconW
    IMPORT GetClassLongPtrW
    IMPORT LoadIconW
    IMPORT ShowWindow
    IMPORT SetForegroundWindow
    IMPORT GetKeyState
    IMPORT RegisterWindowMessageW
    IMPORT ChangeWindowMessageFilterEx
    IMPORT CreatePopupMenu
    IMPORT AppendMenuW
    IMPORT TrackPopupMenu
    IMPORT DestroyMenu
    IMPORT GetCursorPos
    IMPORT PostMessageW
    IMPORT DestroyWindow

; ==============================================================================
; CONSTANTS
; ==============================================================================
NIM_ADD             EQU 0
NIM_DELETE          EQU 2
NIF_MESSAGE         EQU 1
NIF_ICON            EQU 2
NIF_TIP             EQU 4
GCLP_HICON          EQU -14
IDI_APPLICATION     EQU 32512
TPM_RIGHTBUTTON     EQU 0x0002
TPM_RETURNCMD       EQU 0x0100
MSGFLT_ALLOW        EQU 1
IDM_TRAY_RESTORE    EQU 3001
IDM_TRAY_EXIT       EQU 3002

SW_HIDE             EQU 0
SW_RESTORE          EQU 9
VK_SHIFT            EQU 0x10
WM_CREATE           EQU 0x0001
WM_SIZE             EQU 0x0005
SIZE_MINIMIZED      EQU 1
WM_LBUTTONDBLCLK    EQU 0x0203
WM_RBUTTONUP        EQU 0x0205
WM_TRAY             EQU 0x8001
MF_STRING           EQU 0x0000
MF_SEPARATOR        EQU 0x0800

; ==============================================================================
; CONSTANT STRING DATA (READ-ONLY)
; ==============================================================================
    AREA |.const|, DATA, READONLY, ALIGN=3

str_TaskbarCreated  DCW 'T','a','s','k','b','a','r','C','r','e','a','t','e','d', 0
str_TrayRestore     DCW 'R','e','s','t','o','r','e', 0
str_TrayExit        DCW 'E','x','i','t', 0

; ==============================================================================
; MUTABLE DATA
; ==============================================================================
    AREA |.data|, DATA, READWRITE, ALIGN=3

g_trayVisible       DCD 0
g_taskbarCreated    DCD 0

    ALIGN 8
; NOTIFYICONDATAW v1 (cbSize 168): stable callback semantics on every
; supported Windows. hWnd (+8) and hIcon (+32) are filled in at runtime.
tray_nid
    DCD 168, 0                     ; cbSize, (pad)
    DCQ 0                          ; hWnd
    DCD 1                          ; uID
    DCD (NIF_MESSAGE + NIF_ICON + NIF_TIP)  ; uFlags
    DCD WM_TRAY                    ; uCallbackMessage
    DCD 0                          ; (pad)
    DCQ 0                          ; hIcon
    DCW 'C','M','D','T',' ','-',' ','R','u','n',' ','a','s',' '
    DCW 'T','r','u','s','t','e','d','I','n','s','t','a','l','l','e','r', 0
    SPACE 66                       ; pad szTip[64] out to 64 WCHARs

; ==============================================================================
; CODE
; ==============================================================================
    AREA |.text|, CODE, READONLY, ALIGN=4

    EXPORT TrayHandleWindowMessage
    EXPORT TrayShutdown

; ------------------------------------------------------------------------------
; TrayPublish - (re)add the icon. Returns Shell_NotifyIconW result in w0.
;   x0 = hWnd
; ------------------------------------------------------------------------------
TrayPublish PROC
    STP x29, x30, [sp, #-16]!
    MOV x29, sp
    STP x19, x20, [sp, #-16]!

    MOV x19, x0                    ; hWnd

    MOV x0, x19
    MOVN x1, #13                   ; GCLP_HICON = -14
    BL GetClassLongPtrW
    MOV x20, x0
    CBNZ x20, tp_icon_ready
    MOV x0, XZR
    MOV w1, #IDI_APPLICATION
    BL LoadIconW
    MOV x20, x0

tp_icon_ready
    ADRP x9, tray_nid
    ADD x9, x9, tray_nid
    STR x19, [x9, #8]              ; hWnd
    STR x20, [x9, #32]             ; hIcon
    MOV w0, #NIM_ADD
    MOV x1, x9
    BL Shell_NotifyIconW

    LDP x19, x20, [sp], #16
    LDP x29, x30, [sp], #16
    RET
TrayPublish ENDP

; ------------------------------------------------------------------------------
; TrayAdd - hide the window to the tray (idempotent).
;   x0 = hWnd
; ------------------------------------------------------------------------------
TrayAdd PROC
    STP x29, x30, [sp, #-16]!
    MOV x29, sp
    STP x19, x20, [sp, #-16]!

    MOV x19, x0
    ADRP x9, g_trayVisible
    ADD x9, x9, g_trayVisible
    LDR w9, [x9]
    CBNZ w9, ta_done

    MOV x0, x19
    BL TrayPublish
    CBZ w0, ta_done                ; publish failed

    MOV w9, #1
    ADRP x10, g_trayVisible
    ADD x10, x10, g_trayVisible
    STR w9, [x10]
    MOV x0, x19
    MOV w1, #SW_HIDE
    BL ShowWindow

ta_done
    LDP x19, x20, [sp], #16
    LDP x29, x30, [sp], #16
    RET
TrayAdd ENDP

; ------------------------------------------------------------------------------
; TrayRestore - remove the icon and restore the window (idempotent).
;   x0 = hWnd
; ------------------------------------------------------------------------------
TrayRestore PROC
    STP x29, x30, [sp, #-16]!
    MOV x29, sp
    STP x19, x20, [sp, #-16]!

    MOV x19, x0
    ADRP x9, g_trayVisible
    ADD x9, x9, g_trayVisible
    LDR w9, [x9]
    CBZ w9, tr_done

    MOV w0, #NIM_DELETE
    ADRP x1, tray_nid
    ADD x1, x1, tray_nid
    BL Shell_NotifyIconW

    ADRP x9, g_trayVisible
    ADD x9, x9, g_trayVisible
    STR wzr, [x9]

    MOV x0, x19
    MOV w1, #SW_RESTORE
    BL ShowWindow
    MOV x0, x19
    BL SetForegroundWindow

tr_done
    LDP x19, x20, [sp], #16
    LDP x29, x30, [sp], #16
    RET
TrayRestore ENDP

; ------------------------------------------------------------------------------
; TrayContextMenu - show Restore/Exit menu at the cursor.
;   x0 = hWnd     Frame: [sp+0] POINT
; ------------------------------------------------------------------------------
TCM_PT      EQU 0
TCM_FRAME   EQU 16

TrayContextMenu PROC
    STP x29, x30, [sp, #-16]!
    MOV x29, sp
    STP x19, x20, [sp, #-16]!
    STP x21, x22, [sp, #-16]!
    SUB sp, sp, #TCM_FRAME

    MOV x19, x0                    ; hWnd
    BL CreatePopupMenu
    CBZ x0, tcm_done
    MOV x20, x0                    ; menu

    MOV x0, x20
    MOV w1, #MF_STRING
    MOV x2, #IDM_TRAY_RESTORE
    ADRP x3, str_TrayRestore
    ADD x3, x3, str_TrayRestore
    BL AppendMenuW
    MOV x0, x20
    MOV w1, #MF_SEPARATOR
    MOV x2, XZR
    MOV x3, XZR
    BL AppendMenuW
    MOV x0, x20
    MOV w1, #MF_STRING
    MOV x2, #IDM_TRAY_EXIT
    ADRP x3, str_TrayExit
    ADD x3, x3, str_TrayExit
    BL AppendMenuW

    MOV x0, x19
    BL SetForegroundWindow
    ADD x0, sp, #TCM_PT
    BL GetCursorPos

    ; TrackPopupMenu(menu, TPM_RIGHTBUTTON|TPM_RETURNCMD, x, y, 0, hWnd, NULL)
    MOV x0, x20
    MOV w1, #(TPM_RIGHTBUTTON + TPM_RETURNCMD)
    LDR w2, [sp, #TCM_PT + 0]      ; x
    LDR w3, [sp, #TCM_PT + 4]      ; y
    MOV x4, XZR
    MOV x5, x19
    MOV x6, XZR
    BL TrackPopupMenu
    MOV w21, w0                    ; selected command

    MOV x0, x20
    BL DestroyMenu
    MOV x0, x19
    MOV x1, XZR                    ; WM_NULL
    MOV x2, XZR
    MOV x3, XZR
    BL PostMessageW

    CMP w21, #IDM_TRAY_RESTORE
    B.NE tcm_check_exit
    MOV x0, x19
    BL TrayRestore
    B tcm_done
tcm_check_exit
    CMP w21, #IDM_TRAY_EXIT
    B.NE tcm_done
    MOV x0, x19
    BL DestroyWindow

tcm_done
    ADD sp, sp, #TCM_FRAME
    LDP x21, x22, [sp], #16
    LDP x19, x20, [sp], #16
    LDP x29, x30, [sp], #16
    RET
TrayContextMenu ENDP

; ------------------------------------------------------------------------------
; TrayHandleWindowMessage - returns w0 = 1 when the message is fully consumed.
;   x0 = hWnd, w1 = uMsg, x2 = wParam, x3 = lParam
; ------------------------------------------------------------------------------
TrayHandleWindowMessage PROC
    STP x29, x30, [sp, #-16]!
    MOV x29, sp
    STP x19, x20, [sp, #-16]!
    STP x21, x22, [sp, #-16]!

    MOV x19, x0                    ; hWnd
    MOV w20, w1                    ; uMsg
    MOV x21, x2                    ; wParam
    MOV x22, x3                    ; lParam

    CMP w20, #WM_CREATE
    B.NE thwm_size

    ; Register the shell's "TaskbarCreated" broadcast, and let it through UIPI.
    ADRP x0, str_TaskbarCreated
    ADD x0, x0, str_TaskbarCreated
    BL RegisterWindowMessageW
    ADRP x9, g_taskbarCreated
    ADD x9, x9, g_taskbarCreated
    STR w0, [x9]
    CBZ w0, thwm_not_handled
    MOV x0, x19
    ADRP x9, g_taskbarCreated
    ADD x9, x9, g_taskbarCreated
    LDR w1, [x9]                   ; registered message id
    MOV w2, #MSGFLT_ALLOW
    MOV x3, XZR
    BL ChangeWindowMessageFilterEx
    B thwm_not_handled

thwm_size
    CMP w20, #WM_SIZE
    B.NE thwm_callback
    CMP w21, #SIZE_MINIMIZED
    B.NE thwm_not_handled
    MOV w0, #VK_SHIFT
    BL GetKeyState
    TBZ w0, #15, thwm_handled      ; Shift up: ordinary minimize, do nothing
    MOV x0, x19
    BL TrayAdd
    B thwm_handled

thwm_callback
    MOV w9, #WM_TRAY               ; 0x8001 exceeds a CMP immediate; use a reg
    CMP w20, w9
    B.NE thwm_taskbar
    UXTH w9, w22                   ; lParam low word = mouse event
    CMP w9, #WM_LBUTTONDBLCLK
    B.NE thwm_tray_rclick
    MOV x0, x19
    BL TrayRestore
    B thwm_handled
thwm_tray_rclick
    CMP w9, #WM_RBUTTONUP
    B.NE thwm_handled
    MOV x0, x19
    BL TrayContextMenu
    B thwm_handled

thwm_taskbar
    ; Explorer restarted: republish if we were hidden.
    ADRP x9, g_taskbarCreated
    ADD x9, x9, g_taskbarCreated
    LDR w9, [x9]
    CBZ w9, thwm_not_handled
    CMP w20, w9
    B.NE thwm_not_handled
    ADRP x9, g_trayVisible
    ADD x9, x9, g_trayVisible
    LDR w9, [x9]
    CBZ w9, thwm_handled
    MOV x0, x19
    BL TrayPublish

thwm_handled
    MOV w0, #1
    B thwm_ret
thwm_not_handled
    MOV w0, #0

thwm_ret
    LDP x21, x22, [sp], #16
    LDP x19, x20, [sp], #16
    LDP x29, x30, [sp], #16
    RET
TrayHandleWindowMessage ENDP

; ------------------------------------------------------------------------------
; TrayShutdown - remove the icon if visible.
;   x0 = hWnd
; ------------------------------------------------------------------------------
TrayShutdown PROC
    STP x29, x30, [sp, #-16]!
    MOV x29, sp

    ADRP x9, g_trayVisible
    ADD x9, x9, g_trayVisible
    LDR w10, [x9]
    CBZ w10, ts_done

    ADRP x9, tray_nid
    ADD x9, x9, tray_nid
    STR x0, [x9, #8]               ; hWnd
    MOV w0, #NIM_DELETE
    MOV x1, x9
    BL Shell_NotifyIconW

    ADRP x9, g_trayVisible
    ADD x9, x9, g_trayVisible
    STR wzr, [x9]

ts_done
    LDP x29, x30, [sp], #16
    RET
TrayShutdown ENDP

    END
