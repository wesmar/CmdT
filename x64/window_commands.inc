; Included by window.asm; kept architecture-local intentionally.

RunCommand proc frame
    push rbx
    .pushreg rbx
    push rsi
    .pushreg rsi
    push rdi
    .pushreg rdi
    sub rsp, 48
    .allocstack 48
    .endprolog

    mov rbx, rcx                ; Save window handle

    ; Get text from ComboBox edit control
    mov r8d, 520                ; Buffer size
    lea rdx, g_cmdBuf           ; Buffer pointer
    mov rcx, g_hwndEdit         ; ComboBox handle
    sub rsp, 32
    call GetWindowTextW
    add rsp, 32
    test eax, eax
    jz rc_empty                 ; No text entered

    ; Update status to show launching
    lea rdx, str_StatusRunning
    mov rcx, g_hwndStatus
    sub rsp, 32
    call SetWindowTextW
    add rsp, 32

    ; Execute command with TrustedInstaller privileges
    mov edx, 1                  ; new console window flag
    lea rcx, g_cmdBuf           ; Command to execute
    sub rsp, 32
    call RunAsTrustedInstaller
    add rsp, 32
    test eax, eax
    jz rc_fail                  ; Execution failed

    ; Success - save to MRU list
    sub rsp, 32
    call SaveMRU
    add rsp, 32

    ; Update status to success
    lea rdx, str_StatusOK
    mov rcx, g_hwndStatus
    sub rsp, 32
    call SetWindowTextW
    add rsp, 32

    mov eax, 1                  ; Return success
    jmp rc_done

rc_fail:
    ; Execution failed - update status
    lea rdx, str_StatusFail
    mov rcx, g_hwndStatus
    sub rsp, 32
    call SetWindowTextW
    add rsp, 32
    xor eax, eax                ; Return failure
    jmp rc_done

rc_empty:
    ; No command entered - show error message
    mov r9d, MB_ICONERROR
    lea r8, str_TitleErr
    lea rdx, str_ErrEmpty
    mov rcx, rbx
    sub rsp, 32
    call MessageBoxW
    add rsp, 32
    xor eax, eax                ; Return failure

rc_done:
    add rsp, 48
    pop rdi
    pop rsi
    pop rbx
    ret
RunCommand endp

ResolveLnkPath proc frame
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
    push r15
    .pushreg r15
    sub rsp, 112
    .allocstack 112
    .endprolog

    ; Save arguments
    mov r12, rcx                ; r12 = lpLnkPath
    mov r13, rdx                ; r13 = lpOutPath
    mov r14, r8                 ; r14 = lpArgsPath

    ; Local variables on stack:
    ; [rsp+32] = pShellLink (8 bytes)
    ; [rsp+40] = pPersistFile (8 bytes)
    ; [rsp+48] = hr (8 bytes)

    ; Initialize COM
    xor ecx, ecx                ; NULL parameter
    sub rsp, 32
    call CoInitialize
    add rsp, 32

    ; Create IShellLink COM object
    ; CoCreateInstance(&CLSID_ShellLink, NULL, CLSCTX_INPROC_SERVER, &IID_IShellLinkW, &pShellLink)
    lea rax, [rsp+32]           ; &pShellLink
    mov [rsp+32], rax           ; save for later - actually use as temp
    sub rsp, 48
    lea rax, [rsp+48+32]        ; &pShellLink (adjusted for sub rsp)
    mov [rsp+32], rax           ; 5th arg: ppv
    lea r9, IID_IShellLinkW     ; 4th arg: riid
    mov r8d, CLSCTX_INPROC_SERVER ; 3rd arg: dwClsContext
    xor edx, edx                ; 2nd arg: pUnkOuter = NULL
    lea rcx, CLSID_ShellLink    ; 1st arg: rclsid
    call CoCreateInstance
    add rsp, 48
    test eax, eax
    jnz rlp_fail_uninit         ; Failed to create

    ; Query for IPersistFile interface
    ; pShellLink->QueryInterface(&IID_IPersistFile, &pPersistFile)
    mov rbx, [rsp+32]           ; rbx = pShellLink
    mov rax, [rbx]              ; rax = vtable
    lea r8, [rsp+40]            ; 3rd arg: ppvObject = &pPersistFile
    lea rdx, IID_IPersistFile   ; 2nd arg: riid
    mov rcx, rbx                ; 1st arg: this
    sub rsp, 32
    call qword ptr [rax]        ; QueryInterface at vtable[0]
    add rsp, 32
    test eax, eax
    jnz rlp_release_link        ; Failed to query

    ; Load .lnk file
    ; pPersistFile->Load(lpLnkPath, 0)
    mov rbx, [rsp+40]           ; rbx = pPersistFile
    mov rax, [rbx]              ; rax = vtable
    xor r8d, r8d                ; 3rd arg: dwMode = 0
    mov rdx, r12                ; 2nd arg: pszFileName = lpLnkPath
    mov rcx, rbx                ; 1st arg: this
    sub rsp, 32
    call qword ptr [rax+40]     ; Load at vtable[5] = offset 40
    add rsp, 32
    test eax, eax
    jnz rlp_release_both        ; Failed to load

    ; Get target path
    ; pShellLink->GetPath(lpOutPath, 260, NULL, 0)
    mov rbx, [rsp+32]           ; rbx = pShellLink
    mov rax, [rbx]              ; rax = vtable
    sub rsp, 48
    mov qword ptr [rsp+32], 0   ; 5th arg: fFlags = 0
    xor r9d, r9d                ; 4th arg: pfd = NULL
    mov r8d, 260                ; 3rd arg: cch = MAX_PATH
    mov rdx, r13                ; 2nd arg: pszFile = lpOutPath
    mov rcx, rbx                ; 1st arg: this
    call qword ptr [rax+24]     ; GetPath at vtable[3] = offset 24
    add rsp, 48
    mov [rsp+48], eax           ; save hr

    ; Zero out arguments buffer
    mov rdi, r14
    xor eax, eax
    mov ecx, 260
    rep stosw

    ; Get shortcut arguments
    ; pShellLink->GetArguments(lpArgsPath, 520)
    mov rbx, [rsp+32]           ; rbx = pShellLink
    mov rax, [rbx]              ; rax = vtable
    mov r8d, 520                ; 3rd arg: cch
    mov rdx, r14                ; 2nd arg: pszArgs = lpArgsPath
    mov rcx, rbx                ; 1st arg: this
    sub rsp, 32
    call qword ptr [rax+80]     ; GetArguments at vtable[10] = offset 80
    add rsp, 32

    ; Release IPersistFile interface
    mov rbx, [rsp+40]           ; rbx = pPersistFile
    mov rax, [rbx]              ; rax = vtable
    mov rcx, rbx                ; 1st arg: this
    sub rsp, 32
    call qword ptr [rax+16]     ; Release at vtable[2] = offset 16
    add rsp, 32

    ; Release IShellLink interface
    mov rbx, [rsp+32]           ; rbx = pShellLink
    mov rax, [rbx]              ; rax = vtable
    mov rcx, rbx                ; 1st arg: this
    sub rsp, 32
    call qword ptr [rax+16]     ; Release at vtable[2] = offset 16
    add rsp, 32

    ; Uninitialize COM
    sub rsp, 32
    call CoUninitialize
    add rsp, 32

    ; Return based on GetPath result
    mov eax, [rsp+48]           ; hr from GetPath
    test eax, eax
    jnz rlp_fail                ; Failed
    mov eax, 1                  ; Success
    jmp rlp_done

rlp_release_both:
    ; Release both interfaces on error
    mov rbx, [rsp+40]
    mov rax, [rbx]
    mov rcx, rbx
    sub rsp, 32
    call qword ptr [rax+16]
    add rsp, 32

rlp_release_link:
    ; Release ShellLink interface
    mov rbx, [rsp+32]
    mov rax, [rbx]
    mov rcx, rbx
    sub rsp, 32
    call qword ptr [rax+16]
    add rsp, 32

rlp_fail_uninit:
    ; Uninitialize COM on error
    sub rsp, 32
    call CoUninitialize
    add rsp, 32

rlp_fail:
    xor eax, eax                ; Return failure

rlp_done:
    add rsp, 112
    pop r15
    pop r14
    pop r13
    pop r12
    pop rdi
    pop rsi
    pop rbx
    ret
ResolveLnkPath endp

