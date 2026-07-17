; Included by window.asm; kept architecture-local intentionally.

LoadMRU proc frame
    push rbx
    .pushreg rbx
    push rsi
    .pushreg rsi
    push rdi
    .pushreg rdi
    push r12
    .pushreg r12
    sub rsp, 104
    .allocstack 104
    .endprolog

    ; Local variables:
    ; [rsp+32] = hKey (8 bytes)
    ; [rsp+40] = valName (8 bytes)
    ; [rsp+48] = valDataLen (4 bytes)

    ; Open registry key
    sub rsp, 48
    lea rax, [rsp+48+32]
    mov [rsp+32], rax           ; phkResult
    mov r9d, KEY_READ
    xor r8d, r8d                ; ulOptions
    lea rdx, str_regKey
    mov ecx, HKEY_CURRENT_USER
    call RegOpenKeyExW
    add rsp, 48
    test eax, eax
    jnz lm_done                 ; Key doesn't exist

    ; Loop through MRU values (0-4)
    xor r12d, r12d              ; Counter

lm_loop:
    cmp r12d, MRU_MAX_ITEMS     ; Check if done
    jge lm_setsel

    ; Build value name ("0", "1", etc.)
    mov word ptr [rsp+40], '0'
    add word ptr [rsp+40], r12w
    mov word ptr [rsp+42], 0    ; Null terminator

    ; Reset data length
    mov dword ptr [rsp+48], 1040

    ; Query registry value
    sub rsp, 48
    lea rax, [rsp+48+48]
    mov [rsp+40], rax           ; lpcbData
    lea rax, g_tempBuf
    mov [rsp+32], rax           ; lpData
    xor r9d, r9d                ; lpType (NULL)
    xor r8d, r8d                ; lpReserved (NULL)
    lea rdx, [rsp+48+40]        ; lpValueName
    mov rcx, [rsp+48+32]        ; hKey
    call RegQueryValueExW
    add rsp, 48
    test eax, eax
    jnz lm_setsel               ; No more values

    ; Add to ComboBox
    lea r9, g_tempBuf
    xor r8d, r8d
    mov edx, CB_ADDSTRING
    mov rcx, g_hwndEdit
    sub rsp, 32
    call SendMessageW
    add rsp, 32

    inc r12d
    jmp lm_loop

lm_setsel:
    ; Deselect all items — leave the edit field empty on startup
    mov r9d, -1
    xor r8d, r8d
    mov edx, CB_SETCURSEL
    mov rcx, g_hwndEdit
    sub rsp, 32
    call SendMessageW
    add rsp, 32

    ; Explicitly clear the edit field (CBS_DROPDOWN retains text even with no selection)
    lea rdx, g_cmdBuf           ; g_cmdBuf is zero-initialized (BSS), acts as empty wide string
    mov rcx, g_hwndEdit
    sub rsp, 32
    call SetWindowTextW
    add rsp, 32

    ; Close registry key
    mov rcx, [rsp+32]
    sub rsp, 32
    call RegCloseKey
    add rsp, 32

lm_done:
    add rsp, 104
    pop r12
    pop rdi
    pop rsi
    pop rbx
    ret
LoadMRU endp

SaveMRU proc frame
    push rbx
    .pushreg rbx
    push rsi
    .pushreg rsi
    push rdi
    .pushreg rdi
    push r12
    .pushreg r12
    push r13
    .pushreg r13
    push r14
    .pushreg r14
    sub rsp, 184
    .allocstack 184
    .endprolog

    ; Local variables:
    ; [rsp+32] = hKey (8 bytes)
    ; [rsp+40] = disp (4 bytes)
    ; [rsp+48] = valName (8 bytes)
    ; [rsp+56] = textLen (4 bytes)

    ; History disabled -> do nothing, touch no registry state
    cmp dword ptr g_historyEnabled, 0
    je sm_done

    ; Get text from ComboBox
    mov r8d, 520                ; Buffer size
    lea rdx, g_cmdBuf
    mov rcx, g_hwndEdit
    sub rsp, 32
    call GetWindowTextW
    add rsp, 32
    test eax, eax
    jz sm_done                  ; Empty text, don't save

    ; Calculate byte length (chars * 2 + 2 for null)
    shl eax, 1
    add eax, 2
    mov [rsp+56], eax           ; Save text length in bytes

    ; Create or open registry key
    sub rsp, 80
    lea rax, [rsp+80+40]
    mov [rsp+64], rax           ; lpdwDisposition
    lea rax, [rsp+80+32]
    mov [rsp+56], rax           ; phkResult
    mov qword ptr [rsp+48], 0   ; lpSecurityAttributes
    mov dword ptr [rsp+40], KEY_WRITE or KEY_READ
    mov qword ptr [rsp+32], 0   ; lpClass
    xor r9d, r9d                ; dwOptions
    xor r8d, r8d                ; Reserved
    lea rdx, str_regKey
    mov ecx, HKEY_CURRENT_USER
    call RegCreateKeyExW
    add rsp, 80
    test eax, eax
    jnz sm_done                 ; Failed to create/open key

    ; Delete oldest value "4" to make room
    mov word ptr [rsp+48], '4'
    mov word ptr [rsp+50], 0
    lea rdx, [rsp+48]
    mov rcx, [rsp+32]
    sub rsp, 32
    call RegDeleteValueW
    add rsp, 32                 ; Ignore error if doesn't exist

    ; Shift values down: 3→4, 2→3, 1→2, 0→1
    mov r12d, 3                 ; Start from "3"
sm_shift_loop:
    cmp r12d, 0
    jl sm_write_new             ; Done shifting

    ; Build source value name
    mov word ptr [rsp+48], '0'
    add word ptr [rsp+48], r12w
    mov word ptr [rsp+50], 0

    ; Read value at index r12
    mov dword ptr [rsp+60], 1040
    sub rsp, 48
    lea rax, [rsp+48+60]
    mov [rsp+32], rax           ; lpcbData
    lea r9, g_tempBuf           ; lpData
    xor r8d, r8d                ; lpType (NULL)
    xor edx, edx                ; lpReserved (NULL)
    lea rcx, [rsp+48+48]        ; lpValueName
    mov rbx, [rsp+48+32]        ; hKey
    mov rcx, rbx
    lea rdx, [rsp+48+48]
    xor r8d, r8d
    lea r9, g_tempBuf
    lea rax, [rsp+48+60]
    mov [rsp+32], rax
    call RegQueryValueExW
    add rsp, 48
    test eax, eax
    jnz sm_shift_next           ; Value doesn't exist, skip

    ; Build destination value name (r12+1)
    mov word ptr [rsp+48], '1'
    add word ptr [rsp+48], r12w
    mov word ptr [rsp+50], 0

    ; Write to index r12+1
    mov eax, [rsp+60]           ; Data length
    sub rsp, 48
    mov [rsp+40], eax           ; cbData
    lea rax, g_tempBuf
    mov [rsp+32], rax           ; lpData
    mov r9d, REG_SZ             ; Type
    xor r8d, r8d                ; Reserved
    lea rdx, [rsp+48+48]        ; lpValueName
    mov rcx, [rsp+48+32]        ; hKey
    call RegSetValueExW
    add rsp, 48

sm_shift_next:
    dec r12d
    jmp sm_shift_loop

sm_write_new:
    ; Write new value as "0"
    mov word ptr [rsp+48], '0'
    mov word ptr [rsp+50], 0

    mov eax, [rsp+56]           ; Text length
    sub rsp, 48
    mov [rsp+40], eax           ; cbData
    lea rax, g_cmdBuf
    mov [rsp+32], rax           ; lpData
    mov r9d, REG_SZ             ; Type
    xor r8d, r8d                ; Reserved
    lea rdx, [rsp+48+48]        ; lpValueName
    mov rcx, [rsp+48+32]        ; hKey
    call RegSetValueExW
    add rsp, 48

    ; Insert at position 0 in ComboBox
    lea r9, g_cmdBuf
    xor r8d, r8d                ; Index 0
    mov edx, CB_INSERTSTRING
    mov rcx, g_hwndEdit
    sub rsp, 32
    call SendMessageW
    add rsp, 32

    ; Check ComboBox item count
    mov edx, CB_GETCOUNT
    xor r8d, r8d
    xor r9d, r9d
    mov rcx, g_hwndEdit
    sub rsp, 32
    call SendMessageW
    add rsp, 32
    cmp eax, MRU_MAX_ITEMS
    jle sm_setsel               ; Within limit

    ; Delete last item to maintain max 5
    mov r8d, MRU_MAX_ITEMS
    xor r9d, r9d
    mov edx, CB_DELETESTRING
    mov rcx, g_hwndEdit
    sub rsp, 32
    call SendMessageW
    add rsp, 32

sm_setsel:
    ; Select first item
    xor r9d, r9d                ; Index 0
    xor r8d, r8d
    mov edx, CB_SETCURSEL
    mov rcx, g_hwndEdit
    sub rsp, 32
    call SendMessageW
    add rsp, 32

    ; Close registry key
    mov rcx, [rsp+32]
    sub rsp, 32
    call RegCloseKey
    add rsp, 32

sm_done:
    add rsp, 184
    pop r14
    pop r13
    pop r12
    pop rdi
    pop rsi
    pop rbx
    ret
SaveMRU endp

