; ==============================================================================
; window_mru.inc  -  Included by window.asm (ARM64 / armasm64)
;
; Most-Recently-Used command list, persisted under HKCU\Software\cmdt as
; string values "0".."4". Ported from x64/window_mru.inc.
; ==============================================================================

; ==============================================================================
; LoadMRU - populate the ComboBox from the stored values, leaving the edit
; field empty. Silently returns if the key doesn't exist.
;   x19 = loop counter, x20 = hKey
; ==============================================================================
LM_HKEY     EQU 0
LM_VAL      EQU 8        ; value-name buffer ("0".."4" + NUL)
LM_SIZE     EQU 16       ; query data size
LM_FRAME    EQU 32

LoadMRU PROC
    STP x29, x30, [sp, #-16]!
    MOV x29, sp
    STP x19, x20, [sp, #-16]!
    SUB sp, sp, #LM_FRAME

    ; RegOpenKeyExW(HKCU, str_regKey, 0, KEY_READ, &hKey)
    MOVZ w0, #0x0001
    MOVK w0, #0x8000, LSL #16
    ADRP x1, str_regKey
    ADD x1, x1, str_regKey
    MOV x2, XZR
    MOVZ w3, #0x0019
    MOVK w3, #0x0002, LSL #16   ; KEY_READ = 0x00020019
    ADD x4, sp, #LM_HKEY
    BL RegOpenKeyExW
    CBNZ w0, lm_done

    LDR x20, [sp, #LM_HKEY]
    MOV w19, #0

lm_loop
    CMP w19, #MRU_MAX_ITEMS
    B.GE lm_setsel

    ; value name = '0' + counter
    MOV w9, #'0'
    ADD w9, w9, w19
    STRH w9, [sp, #LM_VAL]
    STRH wzr, [sp, #LM_VAL + 2]

    MOV w9, #1040
    STR w9, [sp, #LM_SIZE]

    ; RegQueryValueExW(hKey, valName, NULL, NULL, g_tempBuf, &size)
    MOV x0, x20
    ADD x1, sp, #LM_VAL
    MOV x2, XZR
    MOV x3, XZR
    ADRP x4, g_tempBuf
    ADD x4, x4, g_tempBuf
    ADD x5, sp, #LM_SIZE
    BL RegQueryValueExW
    CBNZ w0, lm_setsel

    ; SendMessageW(g_hwndEdit, CB_ADDSTRING, 0, g_tempBuf)
    ADRP x0, g_hwndEdit
    ADD x0, x0, g_hwndEdit
    LDR x0, [x0]
    MOV w1, #CB_ADDSTRING
    MOV x2, XZR
    ADRP x3, g_tempBuf
    ADD x3, x3, g_tempBuf
    BL SendMessageW

    ADD w19, w19, #1
    B lm_loop

lm_setsel
    ; Deselect all + clear the edit field (CBS_DROPDOWN keeps text otherwise).
    ADRP x0, g_hwndEdit
    ADD x0, x0, g_hwndEdit
    LDR x0, [x0]
    MOV w1, #CB_SETCURSEL
    MOV x2, #-1
    MOV x3, XZR
    BL SendMessageW

    ADRP x0, g_hwndEdit
    ADD x0, x0, g_hwndEdit
    LDR x0, [x0]
    ADRP x1, g_cmdBuf              ; zero-initialized BSS = empty wide string
    ADD x1, x1, g_cmdBuf
    BL SetWindowTextW

    MOV x0, x20
    BL RegCloseKey

lm_done
    ADD sp, sp, #LM_FRAME
    LDP x19, x20, [sp], #16
    LDP x29, x30, [sp], #16
    RET
LoadMRU ENDP

; ==============================================================================
; SaveMRU - push the current edit text to the front of the MRU list, shifting
; older values down and capping at MRU_MAX_ITEMS. No-op when history is off.
;   x19 = loop counter, x20 = hKey
; ==============================================================================
SM_OUT      EQU 0        ; [0..7] outgoing 9th arg (RegCreateKeyExW disposition)
SM_HKEY     EQU 8
SM_DISP     EQU 16
SM_TEXTLEN  EQU 20       ; new value length in bytes
SM_SRC      EQU 24       ; source value name
SM_DST      EQU 28       ; destination value name
SM_QLEN     EQU 32       ; per-value query length
SM_FRAME    EQU 48

SaveMRU PROC
    STP x29, x30, [sp, #-16]!
    MOV x29, sp
    STP x19, x20, [sp, #-16]!
    SUB sp, sp, #SM_FRAME

    ADRP x9, g_historyEnabled
    ADD x9, x9, g_historyEnabled
    LDR w9, [x9]
    CBZ w9, sm_done

    ; GetWindowTextW(g_hwndEdit, g_cmdBuf, 520)
    ADRP x0, g_hwndEdit
    ADD x0, x0, g_hwndEdit
    LDR x0, [x0]
    ADRP x1, g_cmdBuf
    ADD x1, x1, g_cmdBuf
    MOV w2, #520
    BL GetWindowTextW
    CBZ w0, sm_done

    LSL w0, w0, #1
    ADD w0, w0, #2
    STR w0, [sp, #SM_TEXTLEN]       ; byte length incl. NUL

    ; RegCreateKeyExW(HKCU, str_regKey, 0, NULL, 0, KEY_WRITE|KEY_READ,
    ;                 NULL, &hKey, &disp)
    MOVZ w0, #0x0001
    MOVK w0, #0x8000, LSL #16
    ADRP x1, str_regKey
    ADD x1, x1, str_regKey
    MOV x2, XZR
    MOV x3, XZR
    MOV x4, XZR
    MOVZ w5, #0x001F
    MOVK w5, #0x0002, LSL #16       ; KEY_WRITE | KEY_READ = 0x2001F
    MOV x6, XZR
    ADD x7, sp, #SM_HKEY
    ADD x9, sp, #SM_DISP
    STR x9, [sp, #SM_OUT]           ; lpdwDisposition (9th arg)
    BL RegCreateKeyExW
    CBNZ w0, sm_done

    LDR x20, [sp, #SM_HKEY]

    ; Delete oldest value "4" (ignore if absent).
    MOV w9, #'4'
    STRH w9, [sp, #SM_SRC]
    STRH wzr, [sp, #SM_SRC + 2]
    MOV x0, x20
    ADD x1, sp, #SM_SRC
    BL RegDeleteValueW

    ; Shift 3->4, 2->3, 1->2, 0->1.
    MOV w19, #3
sm_shift
    CMP w19, #0
    B.LT sm_write_new

    MOV w9, #'0'
    ADD w9, w9, w19
    STRH w9, [sp, #SM_SRC]
    STRH wzr, [sp, #SM_SRC + 2]

    MOV w9, #1040
    STR w9, [sp, #SM_QLEN]

    ; RegQueryValueExW(hKey, srcName, NULL, NULL, g_tempBuf, &len)
    MOV x0, x20
    ADD x1, sp, #SM_SRC
    MOV x2, XZR
    MOV x3, XZR
    ADRP x4, g_tempBuf
    ADD x4, x4, g_tempBuf
    ADD x5, sp, #SM_QLEN
    BL RegQueryValueExW
    CBNZ w0, sm_shift_next

    ; dstName = '1' + counter
    MOV w9, #'1'
    ADD w9, w9, w19
    STRH w9, [sp, #SM_DST]
    STRH wzr, [sp, #SM_DST + 2]

    ; RegSetValueExW(hKey, dstName, 0, REG_SZ, g_tempBuf, len)
    MOV x0, x20
    ADD x1, sp, #SM_DST
    MOV x2, XZR
    MOV w3, #REG_SZ
    ADRP x4, g_tempBuf
    ADD x4, x4, g_tempBuf
    LDR w5, [sp, #SM_QLEN]
    BL RegSetValueExW

sm_shift_next
    SUB w19, w19, #1
    B sm_shift

sm_write_new
    ; Write the new entry as "0".
    MOV w9, #'0'
    STRH w9, [sp, #SM_SRC]
    STRH wzr, [sp, #SM_SRC + 2]
    MOV x0, x20
    ADD x1, sp, #SM_SRC
    MOV x2, XZR
    MOV w3, #REG_SZ
    ADRP x4, g_cmdBuf
    ADD x4, x4, g_cmdBuf
    LDR w5, [sp, #SM_TEXTLEN]
    BL RegSetValueExW

    ; Insert at ComboBox position 0.
    ADRP x0, g_hwndEdit
    ADD x0, x0, g_hwndEdit
    LDR x0, [x0]
    MOV w1, #CB_INSERTSTRING
    MOV x2, XZR
    ADRP x3, g_cmdBuf
    ADD x3, x3, g_cmdBuf
    BL SendMessageW

    ; Cap the list: delete index MRU_MAX_ITEMS if count exceeds it.
    ADRP x0, g_hwndEdit
    ADD x0, x0, g_hwndEdit
    LDR x0, [x0]
    MOV w1, #CB_GETCOUNT
    MOV x2, XZR
    MOV x3, XZR
    BL SendMessageW
    CMP w0, #MRU_MAX_ITEMS
    B.LE sm_setsel
    ADRP x0, g_hwndEdit
    ADD x0, x0, g_hwndEdit
    LDR x0, [x0]
    MOV w1, #CB_DELETESTRING
    MOV x2, #MRU_MAX_ITEMS
    MOV x3, XZR
    BL SendMessageW

sm_setsel
    ADRP x0, g_hwndEdit
    ADD x0, x0, g_hwndEdit
    LDR x0, [x0]
    MOV w1, #CB_SETCURSEL
    MOV x2, XZR
    MOV x3, XZR
    BL SendMessageW

    MOV x0, x20
    BL RegCloseKey

sm_done
    ADD sp, sp, #SM_FRAME
    LDP x19, x20, [sp], #16
    LDP x29, x30, [sp], #16
    RET
SaveMRU ENDP



    END
