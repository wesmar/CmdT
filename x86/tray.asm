; ==============================================================================
; CMDT - System Tray Integration (x86)
;
; Architecture counterpart of x64/tray.asm. Shift+Minimize hides the window
; in the notification area; double-click or Restore shows it again.
; ==============================================================================

.586
.model flat, stdcall
option casemap:none
include consts.inc

Shell_NotifyIconW          PROTO :DWORD,:DWORD
GetClassLongW             PROTO :DWORD,:DWORD
LoadIconW                 PROTO :DWORD,:DWORD
ShowWindow                PROTO :DWORD,:DWORD
SetForegroundWindow       PROTO :DWORD
GetKeyState               PROTO :DWORD
RegisterWindowMessageW    PROTO :DWORD
ChangeWindowMessageFilterEx PROTO :DWORD,:DWORD,:DWORD,:DWORD
CreatePopupMenu           PROTO
AppendMenuW               PROTO :DWORD,:DWORD,:DWORD,:DWORD
TrackPopupMenu            PROTO :DWORD,:DWORD,:DWORD,:DWORD,:DWORD,:DWORD,:DWORD
DestroyMenu               PROTO :DWORD
GetCursorPos              PROTO :DWORD
PostMessageW              PROTO :DWORD,:DWORD,:DWORD,:DWORD
DestroyWindow             PROTO :DWORD

NIM_ADD                EQU 0
NIM_DELETE             EQU 2
NIF_MESSAGE            EQU 1
NIF_ICON               EQU 2
NIF_TIP                EQU 4
NID_CBSIZE             EQU 152
GCL_HICON              EQU -14
IDI_APPLICATION        EQU 32512
TPM_RIGHTBUTTON        EQU 0002h
TPM_RETURNCMD          EQU 0100h
MSGFLT_ALLOW           EQU 1
IDM_TRAY_RESTORE       EQU 3001
IDM_TRAY_EXIT          EQU 3002

.const
str_TaskbarCreated dw 'T','a','s','k','b','a','r','C','r','e','a','t','e','d',0
str_TrayRestore    dw 'R','e','s','t','o','r','e',0
str_TrayExit       dw 'E','x','i','t',0

.data
align 4
g_trayVisible      dd 0
g_taskbarCreated   dd 0
tray_nid           dd NID_CBSIZE
                   dd 0                    ; hWnd
                   dd 1                    ; uID
                   dd NIF_MESSAGE or NIF_ICON or NIF_TIP
                   dd WM_TRAY
                   dd 0                    ; hIcon
                   dw 'C','M','D','T',' ','-',' ','R','u','n',' ','a','s',' '
                   dw 'T','r','u','s','t','e','d','I','n','s','t','a','l','l','e','r',0
                   dw 33 dup(0)

.code

TrayPublish proc uses ebx hWnd:DWORD
    mov ebx, hWnd
    invoke GetClassLongW, ebx, GCL_HICON
    test eax, eax
    jnz @F
    invoke LoadIconW, 0, IDI_APPLICATION
@@:
    mov tray_nid+4, ebx
    mov tray_nid+20, eax
    invoke Shell_NotifyIconW, NIM_ADD, offset tray_nid
    ret
TrayPublish endp

TrayAdd proc uses ebx hWnd:DWORD
    mov ebx, hWnd
    cmp g_trayVisible, 0
    jne @F
    invoke TrayPublish, ebx
    test eax, eax
    jz @F
    mov g_trayVisible, 1
    invoke ShowWindow, ebx, SW_HIDE
@@:
    ret
TrayAdd endp

TrayRestore proc uses ebx hWnd:DWORD
    mov ebx, hWnd
    cmp g_trayVisible, 0
    je @F
    invoke Shell_NotifyIconW, NIM_DELETE, offset tray_nid
    mov g_trayVisible, 0
    invoke ShowWindow, ebx, SW_RESTORE
    invoke SetForegroundWindow, ebx
@@:
    ret
TrayRestore endp

TrayContextMenu proc uses ebx esi edi hWnd:DWORD
    LOCAL pt[2]:DWORD
    mov ebx, hWnd
    invoke CreatePopupMenu
    test eax, eax
    jz tcm_done
    mov edi, eax
    invoke AppendMenuW, edi, MF_STRING, IDM_TRAY_RESTORE, offset str_TrayRestore
    invoke AppendMenuW, edi, MF_SEPARATOR, 0, 0
    invoke AppendMenuW, edi, MF_STRING, IDM_TRAY_EXIT, offset str_TrayExit
    invoke SetForegroundWindow, ebx
    invoke GetCursorPos, addr pt
    invoke TrackPopupMenu, edi, TPM_RIGHTBUTTON or TPM_RETURNCMD, pt[0], pt[4], 0, ebx, 0
    mov esi, eax
    invoke DestroyMenu, edi
    invoke PostMessageW, ebx, 0, 0, 0
    cmp esi, IDM_TRAY_RESTORE
    jne @F
    invoke TrayRestore, ebx
    jmp tcm_done
@@:
    cmp esi, IDM_TRAY_EXIT
    jne tcm_done
    invoke DestroyWindow, ebx
tcm_done:
    ret
TrayContextMenu endp

PUBLIC TrayHandleWindowMessage
TrayHandleWindowMessage proc uses ebx esi edi hWnd:DWORD,uMsg:DWORD,wParam:DWORD,lParam:DWORD
    mov ebx, hWnd
    mov esi, uMsg
    cmp esi, WM_CREATE
    jne thwm_size
    invoke RegisterWindowMessageW, offset str_TaskbarCreated
    mov g_taskbarCreated, eax
    test eax, eax
    jz thwm_not_handled
    invoke ChangeWindowMessageFilterEx, ebx, eax, MSGFLT_ALLOW, 0
    jmp thwm_not_handled
thwm_size:
    cmp esi, WM_SIZE
    jne thwm_callback
    cmp wParam, SIZE_MINIMIZED
    jne thwm_not_handled
    invoke GetKeyState, VK_SHIFT
    test ax, 8000h
    jz thwm_handled
    invoke TrayAdd, ebx
    jmp thwm_handled
thwm_callback:
    cmp esi, WM_TRAY
    jne thwm_taskbar
    mov eax, lParam
    and eax, 0FFFFh
    cmp eax, WM_LBUTTONDBLCLK
    jne @F
    invoke TrayRestore, ebx
    jmp thwm_handled
@@:
    cmp eax, WM_RBUTTONUP
    jne thwm_handled
    invoke TrayContextMenu, ebx
    jmp thwm_handled
thwm_taskbar:
    mov eax, g_taskbarCreated
    test eax, eax
    jz thwm_not_handled
    cmp esi, eax
    jne thwm_not_handled
    cmp g_trayVisible, 0
    je thwm_handled
    invoke TrayPublish, ebx
thwm_handled:
    mov eax, 1
    ret
thwm_not_handled:
    xor eax, eax
    ret
TrayHandleWindowMessage endp

PUBLIC TrayShutdown
TrayShutdown proc hWnd:DWORD
    cmp g_trayVisible, 0
    je @F
    mov eax, hWnd
    mov tray_nid+4, eax
    invoke Shell_NotifyIconW, NIM_DELETE, offset tray_nid
    mov g_trayVisible, 0
@@:
    ret
TrayShutdown endp

end
