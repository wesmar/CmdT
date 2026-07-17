; ==============================================================================
; CMDT - Run as TrustedInstaller (x86)
; Non-Admin Output-Relay Module
;
; Author: Marek Wesolowski (wesmar)
; Purpose: Implements the relay path that lets `cmdt -cli <command>` run
;          from a non-admin shell still return its child's stdout/stderr to
;          the caller's redirect target or console. Spawns an elevated copy
;          of cmdt with internal `-outfile` and `-errfile` paths, waits for it,
;          then streams the temp file back to this (non-admin) process's
;          STD_OUTPUT — which cmd.exe wired up before launching us, so
;          `>file` / `|pipe` / `>>file` all work transparently.
;
; Exported routine (stdcall):
;   NonAdminRelayLaunch - Attempt relay. Returns 0 if it declines (e.g. -new
;                         flag conflicts with output capture, temp-file
;                         setup failed); never returns once the elevated
;                         child has been spawned (ExitProcess in every
;                         post-spawn path).
; ==============================================================================

.586
.model flat, stdcall
option casemap:none

include consts.inc
include globals.inc

; --- Cross-module strings owned by main.asm ---
EXTRN str_runas:WORD
EXTRN str_newSwitch:WORD
EXTRN str_outfileFlag:WORD
EXTRN str_errfileFlag:WORD

; --- Win32 APIs ---
GetModuleFileNameW      PROTO :DWORD,:DWORD,:DWORD
GetTempPathW            PROTO :DWORD,:DWORD
GetTempFileNameW        PROTO :DWORD,:DWORD,:DWORD,:DWORD
ShellExecuteExW         PROTO :DWORD
WaitForSingleObject     PROTO :DWORD,:DWORD
GetExitCodeProcess      PROTO :DWORD,:DWORD
CloseHandle             PROTO :DWORD
CreateFileW             PROTO :DWORD,:DWORD,:DWORD,:DWORD,:DWORD,:DWORD,:DWORD
ReadFile                PROTO :DWORD,:DWORD,:DWORD,:DWORD,:DWORD
WriteFile               PROTO :DWORD,:DWORD,:DWORD,:DWORD,:DWORD
DeleteFileW             PROTO :DWORD
GetStdHandle            PROTO :DWORD
ExitProcess             PROTO :DWORD

; --- In-project helpers from strutil.asm ---
wcscpy_p                PROTO :DWORD,:DWORD
wcscat_p                PROTO :DWORD,:DWORD
wcscmp_ci               PROTO :DWORD,:DWORD
skip_spaces             PROTO :DWORD

; --- In-project helper from help.asm ---
NudgeConsolePrompt      PROTO

; --- In-project helper from token.asm ---
RunAsTrustedInstaller   PROTO :DWORD,:DWORD

; ==============================================================================
; CONSTANT STRING DATA - private to this module
; ==============================================================================
.const

; 3-char prefix used by GetTempFileNameW. The API itself appends a hex
; sequence + .TMP.
str_cmdtPrefix  dw 'C','M','D',0

; Fixed prefix injected into the elevated child's argument string. The full
; string built becomes:
;   "-cli -outfile \"<stdout-path>\" -errfile \"<stderr-path>\" <command>"
str_relayPrefix dw '-','c','l','i',' ','-','o','u','t','f','i','l','e',' ','"',0
str_relayMid    dw '"',' ','-','e','r','r','f','i','l','e',' ','"',0
str_relayTail   dw '"',' ',0

; Interactive shells that cannot tolerate the relay path (CREATE_NO_WINDOW +
; redirected stdout to a temp file). When the user runs `cmdt -cli <shell>`
; with no further arguments we decline relay so the plain UAC self-elevate
; fallback gives them a real, attachable console window instead.
str_shell_cmd      dw 'c','m','d',0
str_shell_cmd_exe  dw 'c','m','d','.','e','x','e',0
str_shell_ps       dw 'p','o','w','e','r','s','h','e','l','l',0
str_shell_ps_exe   dw 'p','o','w','e','r','s','h','e','l','l','.','e','x','e',0
str_shell_pwsh     dw 'p','w','s','h',0
str_shell_pwsh_exe dw 'p','w','s','h','.','e','x','e',0

; Null-terminated pointer table walked by the shell-guard loop in
; NonAdminRelayLaunch. Add new shell names by inserting another pointer
; before the terminating zero. 4-byte pointers in x86.
shell_names_table  dd str_shell_cmd, str_shell_cmd_exe, \
                      str_shell_ps,  str_shell_ps_exe,  \
                      str_shell_pwsh, str_shell_pwsh_exe, 0

; ==============================================================================
; CODE SECTION
; ==============================================================================
.code

NonAdminRelayLaunch proc uses ebx esi edi pArgvIn:DWORD, argcIn:DWORD, rawCmd:DWORD
    LOCAL sei[60]:BYTE
    LOCAL hProc:DWORD
    LOCAL hRead:DWORD
    LOCAL bytesRead:DWORD
    LOCAL bytesWritten:DWORD
    LOCAL relayExit:DWORD

    mov relayExit, 1

    ; -cli -new must keep a visible new console, so do not capture it.
    cmp argcIn, 3
    jl narl_setup
    mov esi, pArgvIn
    mov eax, [esi+8]
    invoke wcscmp_ci, eax, offset str_newSwitch
    test eax, eax
    jnz narl_decline

    ; Interactive-shell guard. The relay path uses CREATE_NO_WINDOW plus a
    ; temp-file capture of stdout — fundamentally incompatible with a shell
    ; that expects an attached console (no stdin, no prompt redraw, no
    ; output to the user). When argv looks like exactly `cmdt -cli <shell>`
    ; (argc == 3, no extra tokens like `/c` or `-Command`), bail out so the
    ; caller's plain UAC self-elevate path spawns a real new console.
    ; Anything with extra arguments (e.g. `cmd /c dir`) keeps the relay so
    ; its output is still streamed back to the caller.
    cmp argcIn, 3
    jne narl_setup

    mov ebx, offset shell_names_table
narl_shell_loop:
    mov edx, [ebx]
    test edx, edx
    jz narl_setup                       ; List exhausted, not an interactive shell
    mov esi, pArgvIn
    mov eax, [esi+8]                    ; argv[2]
    invoke wcscmp_ci, eax, edx
    test eax, eax
    jnz narl_decline                    ; Match: interactive shell → fall back to plain UAC
    add ebx, 4
    jmp narl_shell_loop

narl_setup:
    invoke GetModuleFileNameW, 0, offset g_exePath, 260
    invoke GetTempPathW, 260, offset g_tempDirBuf
    test eax, eax
    jz narl_decline
    invoke GetTempFileNameW, offset g_tempDirBuf, offset str_cmdtPrefix, 0, offset g_relayPath
    test eax, eax
    jz narl_decline
    invoke GetTempFileNameW, offset g_tempDirBuf, offset str_cmdtPrefix, 0, offset g_relayErrPath
    test eax, eax
    jz narl_delete_stdout_decline

    ; Build "-cli -outfile "<temp>" " + original rest after -cli.
    invoke wcscpy_p, offset g_relayArgs, offset str_relayPrefix
    invoke wcscat_p, offset g_relayArgs, offset g_relayPath
    invoke wcscat_p, offset g_relayArgs, offset str_relayMid
    invoke wcscat_p, offset g_relayArgs, offset g_relayErrPath
    invoke wcscat_p, offset g_relayArgs, offset str_relayTail

    mov esi, rawCmd
    xor edi, edi
narl_skip_exe:
    mov ax, word ptr [esi]
    test ax, ax
    jz narl_append_rest
    cmp ax, '"'
    jne @F
    xor edi, 1
@@:
    cmp ax, ' '
    jne @F
    test edi, edi
    jnz @F
    add esi, 2
    invoke skip_spaces, esi
    mov esi, eax
    jmp narl_skip_cli
@@:
    add esi, 2
    jmp narl_skip_exe

narl_skip_cli:
    mov ax, word ptr [esi]
    test ax, ax
    jz narl_append_rest
    cmp ax, ' '
    jne @F
    add esi, 2
    invoke skip_spaces, esi
    mov esi, eax
    jmp narl_append_rest
@@:
    add esi, 2
    jmp narl_skip_cli

narl_append_rest:
    invoke wcscat_p, offset g_relayArgs, esi

    ; Zero and fill SHELLEXECUTEINFOW.
    lea edi, sei
    xor eax, eax
    mov ecx, 15
    rep stosd
    lea edi, sei
    mov dword ptr [edi], 60
    mov dword ptr [edi+4], 00000040h        ; SEE_MASK_NOCLOSEPROCESS
    mov dword ptr [edi+12], offset str_runas
    mov dword ptr [edi+16], offset g_exePath
    mov dword ptr [edi+20], offset g_relayArgs
    mov dword ptr [edi+28], SW_HIDE

    invoke ShellExecuteExW, edi
    test eax, eax
    jz narl_delete_exit

    mov eax, dword ptr [edi+56]
    mov hProc, eax
    test eax, eax
    jz narl_open_file
    invoke WaitForSingleObject, hProc, 0FFFFFFFFh
    invoke GetExitCodeProcess, hProc, addr relayExit
    invoke CloseHandle, hProc

narl_open_file:
    invoke CreateFileW, offset g_relayPath, GENERIC_READ, FILE_SHARE_READ, 0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0
    cmp eax, -1
    je narl_delete_exit
    mov hRead, eax

    invoke GetStdHandle, STD_OUTPUT_HANDLE
    mov ebx, eax

narl_copy_loop:
    invoke ReadFile, hRead, offset g_relayReadBuf, 4096, addr bytesRead, 0
    test eax, eax
    jz narl_close_file
    cmp bytesRead, 0
    je narl_close_file
    invoke WriteFile, ebx, offset g_relayReadBuf, bytesRead, addr bytesWritten, 0
    jmp narl_copy_loop

narl_close_file:
    invoke CloseHandle, hRead

    invoke CreateFileW, offset g_relayErrPath, GENERIC_READ, FILE_SHARE_READ, 0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0
    cmp eax, -1
    je narl_delete_exit
    mov hRead, eax
    invoke GetStdHandle, STD_ERROR_HANDLE
    mov ebx, eax
narl_err_copy_loop:
    invoke ReadFile, hRead, offset g_relayReadBuf, 4096, addr bytesRead, 0
    test eax, eax
    jz narl_err_close
    cmp bytesRead, 0
    je narl_err_close
    invoke WriteFile, ebx, offset g_relayReadBuf, bytesRead, addr bytesWritten, 0
    jmp narl_err_copy_loop
narl_err_close:
    invoke CloseHandle, hRead

narl_delete_exit:
    invoke NudgeConsolePrompt
    invoke DeleteFileW, offset g_relayPath
    invoke DeleteFileW, offset g_relayErrPath
    invoke ExitProcess, relayExit

narl_decline:
    xor eax, eax
    ret
narl_delete_stdout_decline:
    invoke DeleteFileW, offset g_relayPath
    jmp narl_decline
NonAdminRelayLaunch endp

; ==============================================================================
; AdminRelayLaunch - Output-relay for -cli when the caller is ALREADY elevated
;
; Same problem as NonAdminRelayLaunch (RunAsTrustedInstaller's
; CreateProcessWithTokenW doesn't reliably hand the child usable inherited
; std handles), but the caller here is already Administrator, so re-elevating
; through ShellExecuteExW("runas") would be wrong -- "runas" can still pop a
; second UAC consent prompt even from an elevated process, which is exactly
; the double-prompt regression this routine exists to avoid. Instead this
; captures output entirely in-process: open a local temp file, set
; g_relayHandle, call RunAsTrustedInstaller directly (token duplication only,
; no shell elevation, no second prompt), then stream the temp file to our own
; STD_OUTPUT_HANDLE, exactly like the non-admin path does after its child
; returns.
;
; Note: unlike NonAdminRelayLaunch's non-admin path, this does not call
; NudgeConsolePrompt -- that helper compensates for cmd.exe not waiting on
; GUI-subsystem child processes, which does not apply here (this build is
; console-subsystem, so cmd.exe already waits for us correctly).
;
; Parameters (stdcall): pArgvIn = argv, argcIn = argc, rawCmd = raw cmdline
;
; Returns:
;   EAX = 0 if declined (-new flag present, or an interactive shell target).
;         Caller should fall through to normal admin dispatch.
;   Never returns once the local run has started -- every path past that
;   point ends in ExitProcess.
; ==============================================================================
AdminRelayLaunch proc uses ebx esi edi pArgvIn:DWORD, argcIn:DWORD, rawCmd:DWORD
    LOCAL hWrite:DWORD
    LOCAL hRead:DWORD
    LOCAL bytesRead:DWORD
    LOCAL bytesWritten:DWORD

    ; -cli -new must keep a visible new console, so do not capture it.
    cmp argcIn, 3
    jl arl_setup
    mov esi, pArgvIn
    mov eax, [esi+8]                    ; argv[2]
    invoke wcscmp_ci, eax, offset str_newSwitch
    test eax, eax
    jnz arl_decline

    ; If the non-admin relay's elevated child lands here (it is itself
    ; already-admin once UAC finishes), argv[2] is the internal "-outfile"
    ; token, not a real command. Decline so it falls through to normal
    ; dispatch, where cli.asm's own -outfile parsing sets g_relayHandle and
    ; RunAsTrustedInstaller writes to that temp file directly.
    mov esi, pArgvIn
    mov eax, [esi+8]                    ; argv[2]
    invoke wcscmp_ci, eax, offset str_outfileFlag
    test eax, eax
    jnz arl_decline

    ; Interactive-shell guard, same as the non-admin path.
    cmp argcIn, 3
    jne arl_setup

    mov ebx, offset shell_names_table
arl_shell_loop:
    mov edx, [ebx]
    test edx, edx
    jz arl_setup
    mov esi, pArgvIn
    mov eax, [esi+8]                    ; argv[2]
    invoke wcscmp_ci, eax, edx
    test eax, eax
    jnz arl_decline
    add ebx, 4
    jmp arl_shell_loop

arl_setup:
    invoke GetTempPathW, 260, offset g_tempDirBuf
    test eax, eax
    jz arl_decline
    invoke GetTempFileNameW, offset g_tempDirBuf, offset str_cmdtPrefix, 0, offset g_relayPath
    test eax, eax
    jz arl_decline
    invoke GetTempFileNameW, offset g_tempDirBuf, offset str_cmdtPrefix, 0, offset g_relayErrPath
    test eax, eax
    jz arl_delete_stdout_decline

    ; Open it for inheritable write access.
    mov g_sa[0], 12
    mov g_sa[4], 0
    mov g_sa[8], 1
    invoke CreateFileW, offset g_relayPath, GENERIC_WRITE, FILE_SHARE_READ, offset g_sa, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, 0
    cmp eax, -1
    je arl_decline
    mov hWrite, eax
    mov g_relayHandle, eax
    invoke CreateFileW, offset g_relayErrPath, GENERIC_WRITE, FILE_SHARE_READ, offset g_sa, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, 0
    cmp eax, -1
    je arl_close_stdout_decline
    mov g_relayErrHandle, eax

    ; Locate REST (everything after "-cli") by walking the raw cmdline, and
    ; copy it straight into g_cmdBuf -- there's no re-spawn here.
    mov esi, rawCmd
    xor edi, edi
arl_skip_exe:
    mov ax, word ptr [esi]
    test ax, ax
    jz arl_run
    cmp ax, '"'
    jne @F
    xor edi, 1
@@:
    cmp ax, ' '
    jne @F
    test edi, edi
    jnz @F
    add esi, 2
    invoke skip_spaces, esi
    mov esi, eax
    jmp arl_skip_cli
@@:
    add esi, 2
    jmp arl_skip_exe

arl_skip_cli:
    mov ax, word ptr [esi]
    test ax, ax
    jz arl_run
    cmp ax, ' '
    jne @F
    add esi, 2
    invoke skip_spaces, esi
    mov esi, eax
    jmp arl_run
@@:
    add esi, 2
    jmp arl_skip_cli

arl_run:
    mov word ptr g_cmdBuf, 0
    invoke wcscpy_p, offset g_cmdBuf, esi

    ; RunAsTrustedInstaller sees g_relayHandle != 0 and runs relay Mode 3:
    ; CREATE_NO_WINDOW, stdout/stderr -> our temp file, waits for the child
    ; internally. No ShellExecuteExW anywhere in this path.
    invoke RunAsTrustedInstaller, offset g_cmdBuf, 0
    push eax                                ; preserve creation BOOL

    ; Close the write handle so the file is flushed before we read it back.
    invoke CloseHandle, hWrite
    mov g_relayHandle, 0
    invoke CloseHandle, g_relayErrHandle
    mov g_relayErrHandle, 0

    invoke CreateFileW, offset g_relayPath, GENERIC_READ, FILE_SHARE_READ, 0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0
    cmp eax, -1
    je arl_delete_exit
    mov hRead, eax

    invoke GetStdHandle, STD_OUTPUT_HANDLE
    mov ebx, eax

arl_copy_loop:
    invoke ReadFile, hRead, offset g_relayReadBuf, 4096, addr bytesRead, 0
    test eax, eax
    jz arl_close_file
    cmp bytesRead, 0
    je arl_close_file
    invoke WriteFile, ebx, offset g_relayReadBuf, bytesRead, addr bytesWritten, 0
    jmp arl_copy_loop

arl_close_file:
    invoke CloseHandle, hRead

    invoke CreateFileW, offset g_relayErrPath, GENERIC_READ, FILE_SHARE_READ, 0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0
    cmp eax, -1
    je arl_delete_exit
    mov hRead, eax
    invoke GetStdHandle, STD_ERROR_HANDLE
    mov ebx, eax
arl_err_copy_loop:
    invoke ReadFile, hRead, offset g_relayReadBuf, 4096, addr bytesRead, 0
    test eax, eax
    jz arl_err_close
    cmp bytesRead, 0
    je arl_err_close
    invoke WriteFile, ebx, offset g_relayReadBuf, bytesRead, addr bytesWritten, 0
    jmp arl_err_copy_loop
arl_err_close:
    invoke CloseHandle, hRead

arl_delete_exit:
    invoke DeleteFileW, offset g_relayPath
    invoke DeleteFileW, offset g_relayErrPath
    pop eax
    test eax, eax
    jz arl_process_failed
    invoke ExitProcess, g_childExitCode
arl_process_failed:
    invoke ExitProcess, 1

arl_decline:
    xor eax, eax
    ret
arl_close_stdout_decline:
    invoke CloseHandle, g_relayHandle
    mov g_relayHandle, 0
arl_delete_stdout_decline:
    invoke DeleteFileW, offset g_relayPath
    jmp arl_decline
AdminRelayLaunch endp

end
