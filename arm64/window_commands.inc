; ==============================================================================
; window_commands.inc  -  Included by window.asm (ARM64 / armasm64)
;
; RunCommand (execute the edit-box command as TrustedInstaller) and
; ResolveLnkPath (resolve a .lnk shortcut to its target + arguments via COM).
; Ported from x64/window_commands.inc.
; ==============================================================================

    EXPORT ResolveLnkPath

; ==============================================================================
; RunCommand - run the ComboBox text as TrustedInstaller, update the status
; label and MRU list.
;   x0 = hWnd
;   returns w0 = 1 success, 0 failure/empty
; ==============================================================================
RunCommand PROC
    STP x29, x30, [sp, #-16]!
    MOV x29, sp
    STP x19, x20, [sp, #-16]!

    MOV x19, x0                     ; x19 = hWnd

    ; GetWindowTextW(g_hwndEdit, g_cmdBuf, 520)
    ADRP x0, g_hwndEdit
    ADD x0, x0, g_hwndEdit
    LDR x0, [x0]
    ADRP x1, g_cmdBuf
    ADD x1, x1, g_cmdBuf
    MOV w2, #520
    BL GetWindowTextW
    CBZ w0, rc_empty                ; nothing entered

    ; SetWindowTextW(g_hwndStatus, str_StatusRunning)
    ADRP x0, g_hwndStatus
    ADD x0, x0, g_hwndStatus
    LDR x0, [x0]
    ADRP x1, str_StatusRunning
    ADD x1, x1, str_StatusRunning
    BL SetWindowTextW

    ; RunAsTrustedInstaller(g_cmdBuf, 1)  (1 = new console window)
    ADRP x0, g_cmdBuf
    ADD x0, x0, g_cmdBuf
    MOV w1, #1
    BL RunAsTrustedInstaller
    CBZ w0, rc_fail

    BL SaveMRU

    ; SetWindowTextW(g_hwndStatus, str_StatusOK)
    ADRP x0, g_hwndStatus
    ADD x0, x0, g_hwndStatus
    LDR x0, [x0]
    ADRP x1, str_StatusOK
    ADD x1, x1, str_StatusOK
    BL SetWindowTextW

    MOV w0, #1
    B rc_done

rc_fail
    ADRP x0, g_hwndStatus
    ADD x0, x0, g_hwndStatus
    LDR x0, [x0]
    ADRP x1, str_StatusFail
    ADD x1, x1, str_StatusFail
    BL SetWindowTextW
    MOV w0, #0
    B rc_done

rc_empty
    ; MessageBoxW(hWnd, str_ErrEmpty, str_TitleErr, MB_ICONERROR)
    MOV x0, x19
    ADRP x1, str_ErrEmpty
    ADD x1, x1, str_ErrEmpty
    ADRP x2, str_TitleErr
    ADD x2, x2, str_TitleErr
    MOV w3, #MB_ICONERROR
    BL MessageBoxW
    MOV w0, #0

rc_done
    LDP x19, x20, [sp], #16
    LDP x29, x30, [sp], #16
    RET
RunCommand ENDP

; ==============================================================================
; ResolveLnkPath - resolve a Windows shortcut to its target and arguments.
;   x0 = lpLnkPath (input .lnk path)
;   x1 = lpOutPath (receives target path, MAX_PATH)
;   x2 = lpArgsPath (receives arguments, up to 520 chars)
;   returns w0 = 1 on success, 0 on failure
;
; Frame:  [sp+0] pShellLink   [sp+8] pPersistFile   [sp+16] GetPath hr
; Regs:   x19 lpLnkPath, x20 lpOutPath, x21 lpArgsPath, x22 pShellLink,
;         x23 pPersistFile
; ==============================================================================
RLP_PSL     EQU 0
RLP_PPF     EQU 8
RLP_HR      EQU 16
RLP_FRAME   EQU 32

ResolveLnkPath PROC
    STP x29, x30, [sp, #-16]!
    MOV x29, sp
    STP x19, x20, [sp, #-16]!
    STP x21, x22, [sp, #-16]!
    STP x23, x24, [sp, #-16]!
    SUB sp, sp, #RLP_FRAME

    MOV x19, x0                     ; lpLnkPath
    MOV x20, x1                     ; lpOutPath
    MOV x21, x2                     ; lpArgsPath

    ; CoInitialize(NULL)
    MOV x0, XZR
    BL CoInitialize

    ; CoCreateInstance(&CLSID_ShellLink, NULL, CLSCTX_INPROC_SERVER,
    ;                  &IID_IShellLinkW, &pShellLink)
    ADRP x0, CLSID_ShellLink
    ADD x0, x0, CLSID_ShellLink
    MOV x1, XZR
    MOV w2, #CLSCTX_INPROC_SERVER
    ADRP x3, IID_IShellLinkW
    ADD x3, x3, IID_IShellLinkW
    ADD x4, sp, #RLP_PSL
    BL CoCreateInstance
    CBNZ w0, rlp_fail_uninit

    LDR x22, [sp, #RLP_PSL]         ; pShellLink

    ; pShellLink->QueryInterface(&IID_IPersistFile, &pPersistFile)  [vtbl+0]
    LDR x8, [x22]
    LDR x9, [x8, #0]
    MOV x0, x22
    ADRP x1, IID_IPersistFile
    ADD x1, x1, IID_IPersistFile
    ADD x2, sp, #RLP_PPF
    BLR x9
    CBNZ w0, rlp_release_link

    LDR x23, [sp, #RLP_PPF]         ; pPersistFile

    ; pPersistFile->Load(lpLnkPath, 0)  [vtbl+40]
    LDR x8, [x23]
    LDR x9, [x8, #40]
    MOV x0, x23
    MOV x1, x19
    MOV x2, XZR
    BLR x9
    CBNZ w0, rlp_release_both

    ; pShellLink->GetPath(lpOutPath, 260, NULL, 0)  [vtbl+24]
    LDR x8, [x22]
    LDR x9, [x8, #24]
    MOV x0, x22
    MOV x1, x20
    MOV w2, #260
    MOV x3, XZR
    MOV x4, XZR
    BLR x9
    STR w0, [sp, #RLP_HR]           ; save GetPath hr

    ; Zero the arguments buffer (260 wchars = 520 bytes) before GetArguments.
    MOV x8, x21
    MOV w9, #65                     ; 65 * 8 = 520 bytes
rlp_zero
    STR xzr, [x8], #8
    SUBS w9, w9, #1
    B.NE rlp_zero

    ; pShellLink->GetArguments(lpArgsPath, 520)  [vtbl+80]
    LDR x8, [x22]
    LDR x9, [x8, #80]
    MOV x0, x22
    MOV x1, x21
    MOV w2, #520
    BLR x9

    ; pPersistFile->Release()  [vtbl+16]
    LDR x8, [x23]
    LDR x9, [x8, #16]
    MOV x0, x23
    BLR x9

    ; pShellLink->Release()  [vtbl+16]
    LDR x8, [x22]
    LDR x9, [x8, #16]
    MOV x0, x22
    BLR x9

    BL CoUninitialize

    ; Success iff GetPath returned S_OK.
    LDR w0, [sp, #RLP_HR]
    CBNZ w0, rlp_fail
    MOV w0, #1
    B rlp_done

rlp_release_both
    LDR x8, [x23]
    LDR x9, [x8, #16]
    MOV x0, x23
    BLR x9

rlp_release_link
    LDR x8, [x22]
    LDR x9, [x8, #16]
    MOV x0, x22
    BLR x9

rlp_fail_uninit
    BL CoUninitialize

rlp_fail
    MOV w0, #0

rlp_done
    ADD sp, sp, #RLP_FRAME
    LDP x23, x24, [sp], #16
    LDP x21, x22, [sp], #16
    LDP x19, x20, [sp], #16
    LDP x29, x30, [sp], #16
    RET
ResolveLnkPath ENDP



    END
