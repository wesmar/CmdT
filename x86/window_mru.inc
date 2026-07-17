; Included by window.asm; kept architecture-local intentionally.

LoadMRU proc uses ebx esi edi
    LOCAL hKey:DWORD                        ; Registry key handle
    LOCAL valName[4]:WORD                   ; Value name buffer ("0" to "4")
    LOCAL valDataLen:DWORD                  ; Data length for RegQueryValueExW
    LOCAL idx:DWORD                         ; Loop index

    ; Open registry key for reading
    invoke RegOpenKeyExW, HKEY_CURRENT_USER, offset str_regKey, 0, KEY_READ, addr hKey
    test eax, eax
    jnz lm_done                             ; Key doesn't exist, exit

    ; Loop through MRU values (0-4)
    mov idx, 0
lm_loop:
    cmp idx, MRU_MAX_ITEMS
    jge lm_setsel                           ; All items loaded

    ; Prepare value name ("0", "1", etc.)
    mov ax, '0'
    add ax, word ptr [idx]
    mov word ptr [valName], ax
    mov word ptr [valName+2], 0

    ; Read value from registry
    mov valDataLen, 1040                    ; Buffer size in bytes
    invoke RegQueryValueExW, hKey, addr valName, 0, 0, offset g_tempBuf, addr valDataLen
    test eax, eax
    jnz lm_setsel                           ; Value doesn't exist, stop

    ; Add command to ComboBox
    invoke SendMessageW, g_hwndEdit, CB_ADDSTRING, 0, offset g_tempBuf

    inc idx
    jmp lm_loop

lm_setsel:
    ; Don't pre-select any MRU entry. CB_SETCURSEL -1 clears the dropdown
    ; selection; then SetWindowTextW with an empty string also clears the
    ; CBS_DROPDOWN edit-field text (which CB_SETCURSEL -1 alone may leave
    ; populated). g_cmdBuf is zero-initialized in the BSS, so passing it
    ; is equivalent to passing L"".
    invoke SendMessageW, g_hwndEdit, CB_SETCURSEL, -1, 0
    invoke SetWindowTextW, g_hwndEdit, offset g_cmdBuf
    invoke RegCloseKey, hKey

lm_done:
    ret
LoadMRU endp

SaveMRU proc uses ebx esi edi
    LOCAL hKey:DWORD                        ; Registry key handle
    LOCAL disp:DWORD                        ; Disposition (key created/opened)
    LOCAL valName[4]:WORD                   ; Value name buffer
    LOCAL textLen:DWORD                     ; Command text length in bytes
    LOCAL idx:DWORD                         ; Loop index
    LOCAL dataLen:DWORD                     ; Data length for read operations

    ; History disabled -> do nothing, touch no registry state
    cmp g_historyEnabled, 0
    je sm_done

    ; Get command text from ComboBox
    invoke GetWindowTextW, g_hwndEdit, offset g_cmdBuf, 520
    test eax, eax
    jz sm_done                              ; Empty text, don't save
    
    ; Calculate text length in bytes (including null terminator)
    shl eax, 1                              ; Length in wide chars → bytes
    add eax, 2                              ; Add null terminator
    mov textLen, eax

    ; Open or create registry key
    invoke RegCreateKeyExW, HKEY_CURRENT_USER, offset str_regKey, 0, 0, 0, KEY_WRITE or KEY_READ, 0, addr hKey, addr disp
    test eax, eax
    jnz sm_done                             ; Key creation failed

    ; Delete oldest value "4" to make room (ignore if doesn't exist)
    mov word ptr [valName], '4'
    mov word ptr [valName+2], 0
    invoke RegDeleteValueW, hKey, addr valName

    ; Shift existing values down: 3→4, 2→3, 1→2, 0→1
    mov idx, 3
sm_shift_loop:
    cmp idx, 0
    jl sm_write_new                         ; All shifted, write new

    ; Read value at index
    mov ax, '0'
    add ax, word ptr [idx]
    mov word ptr [valName], ax
    mov word ptr [valName+2], 0

    mov dataLen, 1040
    invoke RegQueryValueExW, hKey, addr valName, 0, 0, offset g_tempBuf, addr dataLen
    test eax, eax
    jnz sm_shift_next                       ; Value doesn't exist, skip

    ; Write to index+1
    mov ax, '1'
    add ax, word ptr [idx]
    mov word ptr [valName], ax
    mov word ptr [valName+2], 0

    invoke RegSetValueExW, hKey, addr valName, 0, REG_SZ, offset g_tempBuf, dataLen

sm_shift_next:
    dec idx
    jmp sm_shift_loop

sm_write_new:
    ; Write new command as value "0"
    mov word ptr [valName], '0'
    mov word ptr [valName+2], 0
    invoke RegSetValueExW, hKey, addr valName, 0, REG_SZ, offset g_cmdBuf, textLen

    ; Add to ComboBox at position 0
    invoke SendMessageW, g_hwndEdit, CB_INSERTSTRING, 0, offset g_cmdBuf

    ; Limit ComboBox to 5 items
    invoke SendMessageW, g_hwndEdit, CB_GETCOUNT, 0, 0
    cmp eax, MRU_MAX_ITEMS
    jle sm_setsel
    invoke SendMessageW, g_hwndEdit, CB_DELETESTRING, MRU_MAX_ITEMS, 0

sm_setsel:
    ; Same no-auto-select policy as LoadMRU — after saving, leave the
    ; edit field empty so the user types fresh next time instead of
    ; accidentally re-running the previous command.
    invoke SendMessageW, g_hwndEdit, CB_SETCURSEL, -1, 0
    invoke SetWindowTextW, g_hwndEdit, offset g_cmdBuf
    invoke RegCloseKey, hKey

sm_done:
    ret
SaveMRU endp

