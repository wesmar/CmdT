; ==============================================================================
; CMDT - Run as TrustedInstaller (x86)
; Command-Line / File-Run Dispatch Module
;
; Author: Marek Wesolowski (wesmar)
; Purpose: Hosts the two execution-style modes that actually run user
;          commands as TrustedInstaller — the explicit `-cli <command>` flow
;          and the context-menu `cmdt.exe "<file>"` flow. Both end up
;          calling RunAsTrustedInstaller; this module owns the command-line
;          parsing, optional .lnk resolution, the internal -outfile relay
;          protocol, and the regedit WOW64 fallback (FixRegeditPath).
;
; Exported routines:
;   FixRegeditPath  - PROC. Rewrite "regedit"/"regedit.exe" to the SysWOW64
;                     copy so cmdt can resolve regedit on 64-bit Windows
;                     when run as the 32-bit binary. Used by both
;                     mode_cli_setup and window.asm's button handler.
;
; Exported labels (PUBLIC at file scope, reached via JMP from start in
; main.asm — they share start's stack frame for `msg`/`sei`, and rely on
; the globals g_argv/g_argc/g_argv1/g_sa that main promoted from LOCALs):
;   mode_cli_found  - argv[1] == "-cli" / "--cli" / "cli"
;   mode_file_run   - argv[1] looks like a file path (context-menu route)
; ==============================================================================

.586
.model flat, stdcall
option casemap:none

include consts.inc
include globals.inc

; --- Strings owned by main.asm ---
EXTRN str_newSwitch:WORD
EXTRN str_extLnk_m:WORD
EXTRN str_space:WORD
EXTRN str_outfileFlag:WORD

; --- Cross-module proc in main.asm reached on GUI fallback ---
mode_gui                PROTO

; --- Win32 APIs ---
GetCommandLineW         PROTO
LocalFree               PROTO :DWORD
ExitProcess             PROTO :DWORD
CreateFileW             PROTO :DWORD,:DWORD,:DWORD,:DWORD,:DWORD,:DWORD,:DWORD
CloseHandle             PROTO :DWORD

; --- Other modules ---
RunAsTrustedInstaller   PROTO :DWORD,:DWORD
ResolveLnkPath          PROTO :DWORD,:DWORD,:DWORD

; --- String helpers from strutil.asm ---
wcscpy_p                PROTO :DWORD,:DWORD
wcscat_p                PROTO :DWORD,:DWORD
wcscmp_ci               PROTO :DWORD,:DWORD
wcscmp_token            PROTO :DWORD,:DWORD
skip_spaces             PROTO :DWORD
wcslen_p                PROTO :DWORD

; --- Make our dispatch labels visible to main.asm's `start` ---
PUBLIC mode_cli_found
PUBLIC mode_file_run

; ==============================================================================
; CONSTANT STRING DATA
; ==============================================================================
.const

; CLI fallback for regedit on WOW64 systems (FixRegeditPath only)
str_regedit         dw 'r','e','g','e','d','i','t',0
str_regedit_exe     dw 'r','e','g','e','d','i','t','.','e','x','e',0
str_regedit_path    dw 'C',':','\','W','i','n','d','o','w','s','\','S','y','s','W','O','W','6','4','\','r','e','g','e','d','i','t','.','e','x','e',0

; ==============================================================================
; CODE SECTION
; ==============================================================================
.code

; ==============================================================================
; FixRegeditPath - WOW64-safe regedit fallback
;
; Purpose: 32-bit cmdt running on 64-bit Windows can't launch
;          System32\regedit.exe directly (WOW64 redirection sends it to
;          SysWOW64, which historically didn't include regedit; on systems
;          where it exists it's the 32-bit copy anyway). If the command
;          token is "regedit" or "regedit.exe" without a path, rewrite it
;          to the explicit SysWOW64 path so we always run the 32-bit
;          regedit consistent with our process bitness.
;
; Parameters:
;   lpCmd - Pointer to command line string
;
; Returns:
;   EAX = Pointer to command string to execute (original or g_cmdBuf)
; ==============================================================================
FixRegeditPath proc lpCmd:DWORD
    LOCAL endCh:WORD
    LOCAL endPtr:DWORD
    LOCAL afterPtr:DWORD
    LOCAL quoted:DWORD
    LOCAL hasPath:DWORD

    mov esi, lpCmd
    mov eax, esi
    test eax, eax
    jz frp_return_orig

    invoke skip_spaces, esi
    mov esi, eax

    mov quoted, 0
    mov hasPath, 0

    cmp word ptr [esi], '"'
    jne frp_token_start
    mov quoted, 1
    add esi, 2

frp_token_start:
    mov edi, esi
frp_scan:
    mov ax, word ptr [edi]
    test ax, ax
    jz frp_token_end
    cmp ax, '"'
    jne @F
    cmp quoted, 0
    je @F
    jmp frp_token_end
@@:
    cmp ax, ' '
    jne @F
    cmp quoted, 0
    jne @F
    jmp frp_token_end
@@:
    cmp ax, '\'
    je frp_mark_path
    cmp ax, ':'
    jne frp_advance
frp_mark_path:
    mov hasPath, 1
frp_advance:
    add edi, 2
    jmp frp_scan

frp_token_end:
    mov endPtr, edi
    mov eax, edi
    mov afterPtr, eax
    cmp quoted, 0
    je @F
    cmp word ptr [edi], '"'
    jne @F
    add afterPtr, 2
@@:
    cmp hasPath, 0
    jne frp_return_orig

    mov ax, word ptr [edi]
    mov endCh, ax
    mov word ptr [edi], 0

    push offset str_regedit
    push esi
    call wcscmp_ci
    test eax, eax
    jnz frp_match
    push offset str_regedit_exe
    push esi
    call wcscmp_ci
    test eax, eax
    jz frp_restore_return

frp_match:
    mov ax, endCh
    mov word ptr [edi], ax

    invoke skip_spaces, afterPtr
    invoke wcscpy_p, offset g_tempBuf, eax

    invoke wcscpy_p, offset g_cmdBuf, offset str_regedit_path
    invoke wcslen_p, offset g_tempBuf
    test eax, eax
    jz frp_return_buf
    invoke wcscat_p, offset g_cmdBuf, offset str_space
    invoke wcscat_p, offset g_cmdBuf, offset g_tempBuf

frp_return_buf:
    mov eax, offset g_cmdBuf
    ret

frp_restore_return:
    mov ax, endCh
    mov word ptr [edi], ax
frp_return_orig:
    mov eax, lpCmd
    ret
FixRegeditPath endp

; ------------------------------------------------------------------------------
; mode_file_run - Context-menu "Run as TrustedInstaller" file handler
;
; Triggered when argv[1] is not a recognized switch — typically a quoted
; file path delivered by the Explorer context menu we register. We resolve
; .lnk shortcuts ourselves so the target gets the TrustedInstaller token
; rather than launching cmd.exe with a shortcut argument. Falls back to
; mode_gui (in main.asm) when no file argument is found.
; ------------------------------------------------------------------------------
mode_file_run:
    invoke LocalFree, g_argv

    invoke GetCommandLineW
    mov esi, eax
    xor edi, edi

skip_exe_for_file:
    mov ax, word ptr [esi]
    test ax, ax
    jz mode_gui                             ; No file argument found, fall back
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
    jmp run_file_direct
@@:
    add esi, 2
    jmp skip_exe_for_file

run_file_direct:
    mov edx, esi
    cmp word ptr [edx], '"'
    jne @F
    add edx, 2
@@:
    invoke wcscpy_p, offset g_filePath, edx

    invoke wcslen_p, offset g_filePath
    test eax, eax
    jz run_file_exec
    mov edx, offset g_filePath
    cmp word ptr [edx + eax*2 - 2], '"'
    jne @F
    mov word ptr [edx + eax*2 - 2], 0
    dec eax
@@:
    cmp eax, 4
    jl run_file_exec

    mov edx, offset g_filePath
    lea ecx, [edx + eax*2 - 8]
    invoke wcscmp_ci, ecx, offset str_extLnk_m
    test eax, eax
    jz run_file_exec

    push edi
    mov edi, offset g_tempBuf
    xor eax, eax
    mov ecx, 260
    rep stosd
    pop edi

    invoke ResolveLnkPath, offset g_filePath, offset g_cmdBuf, offset g_tempBuf
    test eax, eax
    jz run_file_exec

    invoke wcslen_p, offset g_cmdBuf
    test eax, eax
    jz run_file_lnk_args

    invoke wcslen_p, offset g_tempBuf
    test eax, eax
    jz run_file_lnk_cmd

    invoke wcscat_p, offset g_cmdBuf, offset str_space
    invoke wcscat_p, offset g_cmdBuf, offset g_tempBuf
    jmp run_file_lnk_cmd

run_file_lnk_args:
    invoke wcscpy_p, offset g_cmdBuf, offset g_tempBuf

run_file_lnk_cmd:
    invoke FixRegeditPath, offset g_cmdBuf
    invoke RunAsTrustedInstaller, eax, 0
    invoke ExitProcess, 0

run_file_exec:
    invoke FixRegeditPath, esi
    invoke RunAsTrustedInstaller, eax, 0
    invoke ExitProcess, 0

; ------------------------------------------------------------------------------
; mode_cli_found - Dispatcher target for `cmdt -cli <args>`
;
; Validates argc, recognizes the optional -new flag, then falls into
; mode_cli_setup which parses the actual command from the raw cmdline.
; ------------------------------------------------------------------------------
mode_cli_found:
    cmp g_argc, 3
    jl cli_no_cmd_free

    mov esi, g_argv
    mov eax, [esi+8]                        ; argv[2]
    invoke wcscmp_ci, eax, offset str_newSwitch
    test eax, eax
    jz cli_no_new_flag

    cmp g_argc, 4
    jl cli_no_cmd_free
    mov g_useNewConsole, 1
    jmp cli_free_and_setup

cli_no_new_flag:
    mov g_useNewConsole, 0

cli_free_and_setup:
    invoke LocalFree, g_argv
    jmp mode_cli_setup

cli_no_cmd_free:
    invoke LocalFree, g_argv
    invoke ExitProcess, 1

; ------------------------------------------------------------------------------
; mode_cli_setup - Parse the raw cmdline and dispatch to .lnk-aware runner
;
; Walks the GetCommandLineW string past the exe path, past "-cli", optionally
; honours the internal "-outfile <path>" flag (relay protocol from the
; non-admin parent), optionally skips "-new", then locates the actual user
; command. If the command looks like a path-to-.lnk, it is resolved into the
; real target before RunAsTrustedInstaller is invoked.
; ------------------------------------------------------------------------------
mode_cli_setup:
    invoke GetCommandLineW
    mov esi, eax
    xor ecx, ecx
    mov edi, 0

skip_exe_loop:
    mov ax, word ptr [esi]
    test ax, ax
    jz cli_failed_setup
    cmp ax, '"'
    jne @F
    xor edi, 1
@@:
    cmp ax, ' '
    jne @F
    test edi, edi
    jnz @F
    add esi, 2
    jmp skip_switch_init
@@:
    add esi, 2
    jmp skip_exe_loop

skip_switch_init:
    invoke skip_spaces, esi
    mov esi, eax

    mov edi, 0
skip_switch_loop:
    mov ax, word ptr [esi]
    test ax, ax
    jz cli_failed_setup
    cmp ax, ' '
    jne @F
    add esi, 2
    jmp after_switch
@@:
    add esi, 2
    jmp skip_switch_loop

after_switch:
    invoke skip_spaces, esi
    mov esi, eax

    ; Internal relay mode from non-admin parent:
    ; -cli -outfile "<temp-file>" <command>
    invoke wcscmp_token, esi, offset str_outfileFlag
    test eax, eax
    jz after_outfile

    add esi, 16                             ; Skip "-outfile"
    invoke skip_spaces, esi
    mov esi, eax

    mov edi, esi
    xor ebx, ebx                            ; quoted flag
    cmp word ptr [edi], '"'
    jne outfile_scan_unquoted
    add edi, 2
    mov esi, edi
    mov ebx, 1

outfile_scan_quoted:
    mov ax, word ptr [edi]
    test ax, ax
    jz outfile_copy
    cmp ax, '"'
    je outfile_copy
    add edi, 2
    jmp outfile_scan_quoted

outfile_scan_unquoted:
    mov ax, word ptr [edi]
    test ax, ax
    jz outfile_copy
    cmp ax, ' '
    je outfile_copy
    add edi, 2
    jmp outfile_scan_unquoted

outfile_copy:
    push esi
    push edi
    mov ecx, edi
    sub ecx, esi
    shr ecx, 1
    mov edi, offset g_relayPath
    rep movsw
    mov word ptr [edi], 0
    pop edi
    pop esi

    mov esi, edi
    test ebx, ebx
    jz outfile_advance_spaces
    cmp word ptr [esi], '"'
    jne outfile_advance_spaces
    add esi, 2
outfile_advance_spaces:
    invoke skip_spaces, esi
    mov esi, eax

    ; SECURITY_ATTRIBUTES with bInheritHandle=TRUE so the spawned cmd.exe
    ; inherits the relay file handle as its stdout/stderr.
    mov g_sa[0], 12
    mov g_sa[4], 0
    mov g_sa[8], 1
    invoke CreateFileW, offset g_relayPath, GENERIC_WRITE, FILE_SHARE_READ, offset g_sa, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, 0
    cmp eax, -1
    je after_outfile
    mov g_relayHandle, eax

after_outfile:
    cmp g_useNewConsole, 0
    je run_command

skip_new_token:
    mov ax, word ptr [esi]
    test ax, ax
    jz cli_failed_setup
    cmp ax, ' '
    jne @F
    add esi, 2
    jmp run_command
@@:
    add esi, 2
    jmp skip_new_token

run_command:
    invoke skip_spaces, esi
    mov esi, eax

    invoke wcslen_p, esi
    mov ecx, eax
    cmp ecx, 4
    jl run_no_lnk

    mov edi, esi
    xor ebx, ebx
    xor edx, edx
find_space_or_quote:
    mov ax, word ptr [edi]
    test ax, ax
    jz check_lnk_ext
    cmp ax, '"'
    jne @F
    xor edx, 1
@@:
    cmp ax, ' '
    jne @F
    test edx, edx
    jnz @F
    mov ebx, edi
    jmp check_lnk_ext
@@:
    add edi, 2
    jmp find_space_or_quote

check_lnk_ext:
    test ebx, ebx
    jz check_whole_path

    mov ecx, ebx
    cmp word ptr [ecx-2], '"'
    jne @F
    sub ecx, 2
@@:
    mov edx, esi
    cmp word ptr [edx], '"'
    jne @F
    add edx, 2
@@:
    sub ecx, edx
    shr ecx, 1
    cmp ecx, 4
    jl run_no_lnk

    lea edi, [edx + ecx*2 - 8]
    invoke wcscmp_ci, edi, offset str_extLnk_m
    test eax, eax
    jz run_no_lnk

    push esi
    push ebx

    mov edx, esi
    cmp word ptr [edx], '"'
    jne @F
    add edx, 2
@@:
    mov ecx, ebx
    cmp word ptr [ecx-2], '"'
    jne @F
    sub ecx, 2
@@:
    sub ecx, edx
    shr ecx, 1

    mov esi, edx
    mov edi, offset g_filePath
    rep movsw
    mov word ptr [edi], 0

    pop ebx
    add ebx, 2
    invoke skip_spaces, ebx
    mov esi, eax
    invoke wcscpy_p, offset g_argsBuf, esi

    push edi
    mov edi, offset g_tempBuf
    xor eax, eax
    mov ecx, 260
    rep stosd
    pop edi

    invoke ResolveLnkPath, offset g_filePath, offset g_cmdBuf, offset g_tempBuf
    test eax, eax
    pop esi
    jz run_no_lnk

    invoke wcslen_p, offset g_cmdBuf
    test eax, eax
    jz use_args_only

    mov edi, offset g_cmdBuf
    lea edi, [edi + eax*2]
    mov word ptr [edi], ' '
    add edi, 2
    mov word ptr [edi], 0

use_args_only:
    invoke wcscat_p, offset g_cmdBuf, offset g_tempBuf

    invoke wcslen_p, offset g_argsBuf
    test eax, eax
    jz run_resolved

    invoke wcscat_p, offset g_cmdBuf, offset str_space
    invoke wcscat_p, offset g_cmdBuf, offset g_argsBuf
    jmp run_resolved

check_whole_path:
    mov edx, esi
    mov edi, ecx

    cmp word ptr [edx], '"'
    jne @F
    add edx, 2
    dec edi
@@:
    cmp edi, 1
    jl run_no_lnk
    cmp word ptr [edx + edi*2 - 2], '"'
    jne @F
    dec edi
@@:
    cmp edi, 4
    jl run_no_lnk

    lea eax, [edx + edi*2 - 8]
    invoke wcscmp_ci, eax, offset str_extLnk_m
    test eax, eax
    jz run_no_lnk

    push edi
    mov edi, offset g_tempBuf
    xor eax, eax
    mov ecx, 260
    rep stosd
    pop edi

    mov edx, esi
    cmp word ptr [edx], '"'
    jne @F
    add edx, 2
@@:
    push esi
    mov esi, edx
    mov ecx, edi
    mov edi, offset g_filePath
    rep movsw
    mov word ptr [edi], 0
    pop esi

    invoke ResolveLnkPath, offset g_filePath, offset g_cmdBuf, offset g_tempBuf
    test eax, eax
    jz run_no_lnk

    invoke wcslen_p, offset g_cmdBuf
    test eax, eax
    jz use_lnk_args_only

    mov edi, offset g_cmdBuf
    lea edi, [edi + eax*2]
    mov word ptr [edi], ' '
    add edi, 2
    mov word ptr [edi], 0

use_lnk_args_only:
    invoke wcscat_p, offset g_cmdBuf, offset g_tempBuf
    jmp run_resolved

run_resolved:
    invoke FixRegeditPath, offset g_cmdBuf
    invoke RunAsTrustedInstaller, eax, g_useNewConsole
    jmp run_check_result

run_no_lnk:
    invoke FixRegeditPath, esi
    invoke RunAsTrustedInstaller, eax, g_useNewConsole

run_check_result:
    cmp g_relayHandle, 0
    je @F
    invoke CloseHandle, g_relayHandle
    mov g_relayHandle, 0
@@:
    test eax, eax
    jz cli_failed

    invoke ExitProcess, 0

cli_failed_setup:
cli_failed:
    cmp g_relayHandle, 0
    je @F
    invoke CloseHandle, g_relayHandle
    mov g_relayHandle, 0
@@:
    invoke ExitProcess, 1

end
