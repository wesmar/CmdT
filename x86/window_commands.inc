; Included by window.asm; kept architecture-local intentionally.

RunCommand proc hWnd:DWORD
    ; Get text from ComboBox edit control
    invoke GetWindowTextW, g_hwndEdit, offset g_cmdBuf, 520
    test eax, eax
    jz rc_empty                             ; No text entered

    ; Update status to show launching
    invoke SetWindowTextW, g_hwndStatus, offset str_StatusRunning

    ; Execute command with TrustedInstaller privileges
    ; Parameter 1: new console window
    invoke FixRegeditPath, offset g_cmdBuf
    invoke RunAsTrustedInstaller, eax, 1
    test eax, eax
    jz rc_fail                              ; Execution failed

    ; Success - save to MRU list
    call SaveMRU

    ; Update status to success
    invoke SetWindowTextW, g_hwndStatus, offset str_StatusOK
    mov eax, 1
    ret

rc_fail:
    ; Execution failed - update status
    invoke SetWindowTextW, g_hwndStatus, offset str_StatusFail
    xor eax, eax
    ret

rc_empty:
    ; No command entered - show error message
    invoke MessageBoxW, hWnd, offset str_ErrEmpty, offset str_TitleErr, MB_ICONERROR
    xor eax, eax
    ret
RunCommand endp

ResolveLnkPath proc uses ebx esi edi lpLnkPath:DWORD, lpOutPath:DWORD, lpArgsPath:DWORD
    LOCAL pShellLink:DWORD                  ; IShellLink interface pointer
    LOCAL pPersistFile:DWORD                ; IPersistFile interface pointer
    LOCAL hr:DWORD                          ; HRESULT from GetPath

    ; Initialize COM library
    invoke CoInitialize, 0

    ; Create ShellLink COM object
    invoke CoCreateInstance, offset CLSID_ShellLink, 0, CLSCTX_INPROC_SERVER, offset IID_IShellLinkW, addr pShellLink
    test eax, eax
    jnz rlp_fail_uninit                     ; COM object creation failed

    ; Query for IPersistFile interface
    mov eax, pShellLink
    mov eax, [eax]                          ; Get vtable pointer
    mov ebx, pShellLink
    lea ecx, pPersistFile
    push ecx                                ; ppvObject
    push offset IID_IPersistFile            ; riid
    push ebx                                ; this
    call dword ptr [eax]                    ; QueryInterface at vtable[0]
    test eax, eax
    jnz rlp_release_link                    ; QI failed

    ; Load .lnk file using IPersistFile::Load
    mov eax, pPersistFile
    mov eax, [eax]                          ; Get vtable
    mov ebx, pPersistFile
    push 0                                  ; grfMode = 0
    push lpLnkPath                          ; pszFileName
    push ebx                                ; this
    call dword ptr [eax+20]                 ; Load at vtable[5] = offset 20
    test eax, eax
    jnz rlp_release_both                    ; Load failed

    ; Get target path using IShellLinkW::GetPath
    mov eax, pShellLink
    mov eax, [eax]                          ; Get vtable
    mov ebx, pShellLink
    push 0                                  ; fFlags = 0
    push 0                                  ; pfd = NULL
    push 260                                ; cch = MAX_PATH
    push lpOutPath                          ; pszFile
    push ebx                                ; this
    call dword ptr [eax+12]                 ; GetPath at vtable[3] = offset 12
    mov hr, eax                             ; Save result

    ; Zero argument buffer before getting arguments
    push edi
    mov edi, lpArgsPath
    xor eax, eax
    mov ecx, 260
    rep stosd                               ; Clear 260 DWORDs (1040 bytes)
    pop edi

    ; Get arguments using IShellLinkW::GetArguments
    mov eax, pShellLink
    mov eax, [eax]                          ; Get vtable
    mov ebx, pShellLink
    push 520                                ; cch (buffer size)
    push lpArgsPath                         ; pszArgs
    push ebx                                ; this
    call dword ptr [eax+40]                 ; GetArguments at vtable[10] = offset 40

    ; Release IPersistFile interface
    mov eax, pPersistFile
    mov eax, [eax]
    mov ebx, pPersistFile
    push ebx
    call dword ptr [eax+8]                  ; Release at vtable[2] = offset 8

    ; Release IShellLink interface
    mov eax, pShellLink
    mov eax, [eax]
    mov ebx, pShellLink
    push ebx
    call dword ptr [eax+8]                  ; Release at vtable[2] = offset 8

    ; Uninitialize COM
    invoke CoUninitialize

    ; Return success/failure based on GetPath result
    mov eax, hr
    test eax, eax
    jnz rlp_fail
    mov eax, 1                              ; Success
    ret

rlp_release_both:
    ; Release both interfaces on error
    mov eax, pPersistFile
    mov eax, [eax]
    mov ebx, pPersistFile
    push ebx
    call dword ptr [eax+8]

rlp_release_link:
    ; Release ShellLink interface
    mov eax, pShellLink
    mov eax, [eax]
    mov ebx, pShellLink
    push ebx
    call dword ptr [eax+8]

rlp_fail_uninit:
    ; Uninitialize COM on error
    invoke CoUninitialize

rlp_fail:
    xor eax, eax                            ; Return failure
    ret
ResolveLnkPath endp

