; ==============================================================================
; CMDT - System Tray Integration (x64)
;
; Owns all notification-area state and behavior. Holding Shift while the main
; window is minimized hides it to the tray. Double-click or Restore brings it
; back; the context menu also provides Exit. The icon is republished after an
; Explorer restart, including across the elevated-process UIPI boundary.
; ==============================================================================

option casemap:none
include consts.inc

EXTRN Shell_NotifyIconW:PROC
EXTRN GetClassLongPtrW:PROC
EXTRN LoadIconW:PROC
EXTRN ShowWindow:PROC
EXTRN SetForegroundWindow:PROC
EXTRN GetKeyState:PROC
EXTRN RegisterWindowMessageW:PROC
EXTRN ChangeWindowMessageFilterEx:PROC
EXTRN CreatePopupMenu:PROC
EXTRN AppendMenuW:PROC
EXTRN TrackPopupMenu:PROC
EXTRN DestroyMenu:PROC
EXTRN GetCursorPos:PROC
EXTRN PostMessageW:PROC
EXTRN DestroyWindow:PROC

NIM_ADD                EQU 0
NIM_DELETE             EQU 2
NIF_MESSAGE            EQU 1
NIF_ICON               EQU 2
NIF_TIP                EQU 4
NID_CBSIZE             EQU 168
GCLP_HICON             EQU -14
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
align 8
g_trayVisible      dd 0
g_taskbarCreated   dd 0

; NOTIFYICONDATAW v1: stable callback semantics on every supported Windows.
tray_nid           dd NID_CBSIZE, 0
                   dq 0                    ; hWnd
                   dd 1                    ; uID
                   dd NIF_MESSAGE or NIF_ICON or NIF_TIP
                   dd WM_TRAY              ; uCallbackMessage
                   dd 0
                   dq 0                    ; hIcon
                   dw 'C','M','D','T',' ','-',' ','R','u','n',' ','a','s',' '
                   dw 'T','r','u','s','t','e','d','I','n','s','t','a','l','l','e','r',0
                   dw 33 dup(0)

.code

TrayPublish proc frame
    push rbx
    .pushreg rbx
    sub rsp, 32
    .allocstack 32
    .endprolog
    mov rbx, rcx

    mov edx, GCLP_HICON
    call GetClassLongPtrW
    test rax, rax
    jnz tp_icon_ready
    mov edx, IDI_APPLICATION
    xor ecx, ecx
    call LoadIconW
tp_icon_ready:
    mov qword ptr tray_nid+8, rbx
    mov qword ptr tray_nid+32, rax
    lea rdx, tray_nid
    mov ecx, NIM_ADD
    call Shell_NotifyIconW

    add rsp, 32
    pop rbx
    ret
TrayPublish endp

TrayAdd proc frame
    push rbx
    .pushreg rbx
    sub rsp, 32
    .allocstack 32
    .endprolog
    mov rbx, rcx
    cmp dword ptr g_trayVisible, 0
    jne ta_done
    call TrayPublish
    test eax, eax
    jz ta_done
    mov dword ptr g_trayVisible, 1
    mov edx, SW_HIDE
    mov rcx, rbx
    call ShowWindow
ta_done:
    add rsp, 32
    pop rbx
    ret
TrayAdd endp

TrayRestore proc frame
    push rbx
    .pushreg rbx
    sub rsp, 32
    .allocstack 32
    .endprolog
    mov rbx, rcx
    cmp dword ptr g_trayVisible, 0
    je tr_done
    lea rdx, tray_nid
    mov ecx, NIM_DELETE
    call Shell_NotifyIconW
    mov dword ptr g_trayVisible, 0
    mov edx, SW_RESTORE
    mov rcx, rbx
    call ShowWindow
    mov rcx, rbx
    call SetForegroundWindow
tr_done:
    add rsp, 32
    pop rbx
    ret
TrayRestore endp

TrayContextMenu proc frame
    push rbx
    .pushreg rbx
    push rsi
    .pushreg rsi
    push rdi
    .pushreg rdi
    sub rsp, 64
    .allocstack 64
    .endprolog
    mov rbx, rcx

    call CreatePopupMenu
    test rax, rax
    jz tcm_done
    mov rdi, rax
    lea r9, str_TrayRestore
    mov r8d, IDM_TRAY_RESTORE
    mov edx, MF_STRING
    mov rcx, rdi
    call AppendMenuW
    xor r9d, r9d
    xor r8d, r8d
    mov edx, MF_SEPARATOR
    mov rcx, rdi
    call AppendMenuW
    lea r9, str_TrayExit
    mov r8d, IDM_TRAY_EXIT
    mov edx, MF_STRING
    mov rcx, rdi
    call AppendMenuW

    mov rcx, rbx
    call SetForegroundWindow
    lea rcx, [rsp+56]
    call GetCursorPos
    mov qword ptr [rsp+48], 0
    mov qword ptr [rsp+40], rbx
    mov dword ptr [rsp+32], 0
    mov r9d, dword ptr [rsp+60]
    mov r8d, dword ptr [rsp+56]
    mov edx, TPM_RIGHTBUTTON or TPM_RETURNCMD
    mov rcx, rdi
    call TrackPopupMenu
    mov esi, eax

    mov rcx, rdi
    call DestroyMenu
    xor r9d, r9d
    xor r8d, r8d
    xor edx, edx
    mov rcx, rbx
    call PostMessageW

    cmp esi, IDM_TRAY_RESTORE
    jne tcm_check_exit
    mov rcx, rbx
    call TrayRestore
    jmp tcm_done
tcm_check_exit:
    cmp esi, IDM_TRAY_EXIT
    jne tcm_done
    mov rcx, rbx
    call DestroyWindow
tcm_done:
    add rsp, 64
    pop rdi
    pop rsi
    pop rbx
    ret
TrayContextMenu endp

; Returns nonzero when the message has been fully consumed.
PUBLIC TrayHandleWindowMessage
TrayHandleWindowMessage proc frame
    push rbx
    .pushreg rbx
    push rsi
    .pushreg rsi
    push rdi
    .pushreg rdi
    push r12
    .pushreg r12
    sub rsp, 40
    .allocstack 40
    .endprolog
    mov rbx, rcx
    mov esi, edx
    mov rdi, r8
    mov r12, r9

    cmp esi, WM_CREATE
    jne thwm_size
    lea rcx, str_TaskbarCreated
    call RegisterWindowMessageW
    mov dword ptr g_taskbarCreated, eax
    test eax, eax
    jz thwm_not_handled
    xor r9d, r9d
    mov r8d, MSGFLT_ALLOW
    mov edx, eax
    mov rcx, rbx
    call ChangeWindowMessageFilterEx
    jmp thwm_not_handled

thwm_size:
    cmp esi, WM_SIZE
    jne thwm_callback
    cmp edi, SIZE_MINIMIZED
    jne thwm_not_handled
    mov ecx, VK_SHIFT
    call GetKeyState
    test ax, 8000h
    jz thwm_handled             ; normal minimize; skip meaningless zero layout
    mov rcx, rbx
    call TrayAdd
    jmp thwm_handled

thwm_callback:
    cmp esi, WM_TRAY
    jne thwm_taskbar
    movzx eax, r12w
    cmp eax, WM_LBUTTONDBLCLK
    jne @F
    mov rcx, rbx
    call TrayRestore
    jmp thwm_handled
@@:
    cmp eax, WM_RBUTTONUP
    jne thwm_handled
    mov rcx, rbx
    call TrayContextMenu
    jmp thwm_handled

thwm_taskbar:
    mov eax, dword ptr g_taskbarCreated
    test eax, eax
    jz thwm_not_handled
    cmp esi, eax
    jne thwm_not_handled
    cmp dword ptr g_trayVisible, 0
    je thwm_handled
    mov rcx, rbx
    call TrayPublish
thwm_handled:
    mov eax, 1
    jmp thwm_done
thwm_not_handled:
    xor eax, eax
thwm_done:
    add rsp, 40
    pop r12
    pop rdi
    pop rsi
    pop rbx
    ret
TrayHandleWindowMessage endp

PUBLIC TrayShutdown
TrayShutdown proc frame
    sub rsp, 40
    .allocstack 40
    .endprolog
    cmp dword ptr g_trayVisible, 0
    je ts_done
    mov qword ptr tray_nid+8, rcx
    lea rdx, tray_nid
    mov ecx, NIM_DELETE
    call Shell_NotifyIconW
    mov dword ptr g_trayVisible, 0
ts_done:
    add rsp, 40
    ret
TrayShutdown endp

end
