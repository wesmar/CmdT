; ==============================================================================
; CMDT - Run as TrustedInstaller (x86)
; Non-Admin Output-Relay Module
;
; Author: Marek Wesolowski (wesmar)
; Purpose: Implements the relay path that lets `cmdt -cli <command>` run
;          from a non-admin shell still return its child's stdout/stderr to
;          the caller's redirect target or console. Spawns an elevated copy
;          of cmdt with an internal `-outfile <path>` flag, waits for it,
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

; --- Win32 APIs ---
GetModuleFileNameW      PROTO :DWORD,:DWORD,:DWORD
GetTempPathW            PROTO :DWORD,:DWORD
GetTempFileNameW        PROTO :DWORD,:DWORD,:DWORD,:DWORD
ShellExecuteExW         PROTO :DWORD
WaitForSingleObject     PROTO :DWORD,:DWORD
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

; ==============================================================================
; CONSTANT STRING DATA - private to this module
; ==============================================================================
.const

; 3-char prefix used by GetTempFileNameW. The API itself appends a hex
; sequence + .TMP.
str_cmdtPrefix  dw 'C','M','D',0

; Fixed prefix injected into the elevated child's argument string. The full
; string built becomes:
;   "-cli -outfile \"<relay-path>\" <rest of original args after -cli>"
str_relayPrefix dw '-','c','l','i',' ','-','o','u','t','f','i','l','e',' ','"',0
str_relayMid    dw '"',' ',0

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

    ; -cli -new must keep a visible new console, so do not capture it.
    cmp argcIn, 3
    jl narl_setup
    mov esi, pArgvIn
    mov eax, [esi+8]
    invoke wcscmp_ci, eax, offset str_newSwitch
    test eax, eax
    jnz narl_decline

narl_setup:
    invoke GetModuleFileNameW, 0, offset g_exePath, 260
    invoke GetTempPathW, 260, offset g_tempDirBuf
    test eax, eax
    jz narl_decline
    invoke GetTempFileNameW, offset g_tempDirBuf, offset str_cmdtPrefix, 0, offset g_relayPath
    test eax, eax
    jz narl_decline

    ; Build "-cli -outfile "<temp>" " + original rest after -cli.
    invoke wcscpy_p, offset g_relayArgs, offset str_relayPrefix
    invoke wcscat_p, offset g_relayArgs, offset g_relayPath
    invoke wcscat_p, offset g_relayArgs, offset str_relayMid

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

narl_delete_exit:
    invoke DeleteFileW, offset g_relayPath
    invoke ExitProcess, 0

narl_decline:
    xor eax, eax
    ret
NonAdminRelayLaunch endp

end
