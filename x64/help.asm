; ==============================================================================
; CMDT - Run as TrustedInstaller
; Help / Usage Display Module
;
; Author: Marek Wesolowski (wesmar)
; Purpose: Owns everything related to "show me the options" — recognizing
;          all supported help switches and rendering the usage banner. Lives
;          on its own so that the rest of the project never has to touch
;          help-related strings or output formatting.
;
; Exported routines:
;   HelpCheckAndExit - If argv[1] is a known help variant, free argv and
;                      display usage (this never returns). Otherwise returns
;                      to the caller untouched.
;   ShowUsage        - Free argv (optional) and write the usage banner. Picks
;                      WriteConsoleW for console handles and WriteFile for
;                      redirected handles. Never returns.
; ==============================================================================

option casemap:none

include consts.inc

; --- External Win32 dependencies ---
EXTRN AttachConsole:PROC
EXTRN GetStdHandle:PROC
EXTRN GetFileType:PROC
EXTRN WriteConsoleW:PROC
EXTRN WriteConsoleInputW:PROC
EXTRN WriteFile:PROC
EXTRN LocalFree:PROC
EXTRN ExitProcess:PROC

; --- String helpers from strutil.asm ---
EXTRN wcscmp_ci:PROC
EXTRN wcslen_p:PROC

; ==============================================================================
; CONSTANT STRING DATA
; ==============================================================================
.const

; All supported help-switch spellings. The user-visible canonical form is
; "-help"; the rest are convenience aliases so users with muscle memory from
; other tools (POSIX, classic Windows /?, etc.) don't get surprised.
str_helpSwitch      dw '-','h','e','l','p',0
str_helpSwitchH     dw '-','h',0
str_helpSwitchDD    dw '-','-','h','e','l','p',0
str_helpSwitchQ     dw '-','?',0
str_helpSwitchSQ    dw '/','?',0
str_helpSwitchSH    dw '/','h',0
str_helpSwitchSHELP dw '/','h','e','l','p',0

; Usage banner printed by ShowUsage. CRLF line endings keep the layout
; readable in cmd.exe; WriteConsoleW will treat them as a single line break.
str_usage       dw 13,10
                dw 'U','s','a','g','e',':',' ','c','m','d','t','.','e','x','e',' ','[','o','p','t','i','o','n',']',13,10
                dw 13,10
                dw ' ',' ','-','c','l','i',' ','<','c','m','d','>'
                dw ' ',' ',' ',' ',' ',' ',' ',' ',' ',' '
                dw 'R','u','n',' ','c','o','m','m','a','n','d',13,10
                dw ' ',' ','-','c','l','i',' ','-','n','e','w',' ','<','c','m','d','>'
                dw ' ',' ',' ',' ',' '
                dw 'R','u','n',' ','i','n',' ','n','e','w',' ','c','o','n','s','o','l','e',13,10
                dw ' ',' ','-','i','n','s','t','a','l','l'
                dw ' ',' ',' ',' ',' ',' ',' ',' ',' ',' ',' ',' '
                dw 'A','d','d',' ','c','o','n','t','e','x','t',' ','m','e','n','u',13,10
                dw ' ',' ','-','u','n','i','n','s','t','a','l','l'
                dw ' ',' ',' ',' ',' ',' ',' ',' ',' ',' '
                dw 'R','e','m','o','v','e',' ','c','o','n','t','e','x','t',' ','m','e','n','u',13,10
                dw ' ',' ','-','s','h','i','f','t'
                dw ' ',' ',' ',' ',' ',' ',' ',' ',' ',' ',' ',' ',' ',' '
                dw 'H','o','o','k',' ','s','e','t','h','c','.','e','x','e',13,10
                dw ' ',' ','-','u','n','s','h','i','f','t'
                dw ' ',' ',' ',' ',' ',' ',' ',' ',' ',' ',' ',' '
                dw 'U','n','h','o','o','k',' ','s','e','t','h','c','.','e','x','e',13,10
                dw ' ',' ','-','h','e','l','p',',',' ','-','h',',',' ','-','-','h','e','l','p',',',' ','-','?',',',' ','/','?'
                dw ' ',' '
                dw 'S','h','o','w',' ','t','h','i','s',' ','h','e','l','p',13,10
                dw 13,10
                dw ' ',' ','N','o',' ','a','r','g','s',' ','t','o',' ','s','t','a','r','t',' ','G','U','I','.',13,10
                dw 0

; ==============================================================================
; CODE SECTION
; ==============================================================================
.code

; ==============================================================================
; HelpCheckAndExit - Detect help switches and hand off to ShowUsage
;
; Purpose: Run during entry, before any UAC self-elevation, so that the
;          usage banner reaches the user's original shell (its redirect
;          target or its console). If a help switch is found, this routine
;          frees the argv buffer and never returns. Otherwise it leaves the
;          caller's state untouched and returns.
;
; Parameters:
;   RCX = argc (only the low DWORD is meaningful)
;   RDX = argv array (LocalFree-able pointer from CommandLineToArgvW), or NULL
;
; Returns:
;   Returns to caller if no help switch is present.
;   Never returns if a help switch is present (calls ShowUsage -> ExitProcess).
;
; Modifies: RAX, RCX, RDX, R8, R9, R10, R11 (volatile registers).
;           RDI, RSI are preserved by wcscmp_ci and skip_spaces internally.
; ==============================================================================
HelpCheckAndExit proc frame
    push rbx
    .pushreg rbx
    push r12
    .pushreg r12
    sub rsp, 40                 ; 32 shadow + 8 alignment
    .allocstack 40
    .endprolog

    mov r12, rdx                ; r12 = argv (preserved across calls)

    ; Need at least argv[0] + argv[1] to have any switch to check.
    cmp ecx, 2
    jl hcae_no_help

    test r12, r12
    jz hcae_no_help

    mov rbx, [r12+8]            ; rbx = argv[1] (preserved across calls)

    ; Check each supported help spelling. We compare instead of building a
    ; table because there are only seven and inlining keeps the lookup branch-
    ; predictable for the common "no help" path that takes all branches.
    lea rdx, str_helpSwitch
    mov rcx, rbx
    call wcscmp_ci
    test rax, rax
    jnz hcae_match

    lea rdx, str_helpSwitchH
    mov rcx, rbx
    call wcscmp_ci
    test rax, rax
    jnz hcae_match

    lea rdx, str_helpSwitchDD
    mov rcx, rbx
    call wcscmp_ci
    test rax, rax
    jnz hcae_match

    lea rdx, str_helpSwitchQ
    mov rcx, rbx
    call wcscmp_ci
    test rax, rax
    jnz hcae_match

    lea rdx, str_helpSwitchSQ
    mov rcx, rbx
    call wcscmp_ci
    test rax, rax
    jnz hcae_match

    lea rdx, str_helpSwitchSH
    mov rcx, rbx
    call wcscmp_ci
    test rax, rax
    jnz hcae_match

    lea rdx, str_helpSwitchSHELP
    mov rcx, rbx
    call wcscmp_ci
    test rax, rax
    jnz hcae_match

hcae_no_help:
    ; No help switch detected - return to caller, frame intact.
    add rsp, 40
    pop r12
    pop rbx
    ret

hcae_match:
    ; Help switch found. Hand argv off to ShowUsage so it can LocalFree the
    ; buffer before exiting the process. We `call` (not `jmp`) to keep stack
    ; alignment sane; ShowUsage never returns, so we never restore the frame.
    mov rcx, r12                ; argv to free
    call ShowUsage
    int 3                       ; unreachable
HelpCheckAndExit endp

; ==============================================================================
; ShowUsage - Print usage banner to stdout and exit
;
; Purpose: Selects the right output API based on the stdout handle type:
;          WriteConsoleW for a console (UTF-16 native), WriteFile for files
;          or pipes (raw UTF-16 LE bytes). After printing it nudges cmd.exe
;          into redrawing its prompt by posting a fake VK_RETURN to stdin
;          when stdout is a real console — without that, cmd shows the
;          prompt before our output and the cursor sits idle until the user
;          presses a key.
;
; Parameters:
;   RCX = argv buffer pointer to free, or NULL.
;
; Returns: Does not return (calls ExitProcess with code 1).
; ==============================================================================
ShowUsage proc frame
    push rbx
    .pushreg rbx
    push rdi
    .pushreg rdi
    push r12
    .pushreg r12
    sub rsp, 96
    .allocstack 96
    .endprolog

    ; Free argv if caller handed us one.
    test rcx, rcx
    jz su_no_free
    sub rsp, 32
    call LocalFree
    add rsp, 32
su_no_free:

    ; STD_OUTPUT_HANDLE. If invalid or NULL we bail silently (e.g. launched
    ; from Explorer with no parent console).
    mov ecx, STD_OUTPUT_HANDLE
    sub rsp, 32
    call GetStdHandle
    add rsp, 32
    mov rbx, rax
    test rbx, rbx
    jz su_exit
    cmp rbx, -1
    je su_exit

    ; WCHAR count of the usage text (needed for both output paths).
    lea rcx, str_usage
    call wcslen_p
    mov r12, rax

    ; Pick the right API: WriteConsoleW for true consoles, WriteFile for
    ; everything else (file or pipe). WriteConsoleW silently fails when
    ; stdout points at a file, which is why naive `cmdt -help > out.txt`
    ; previously produced an empty file.
    mov rcx, rbx
    sub rsp, 32
    call GetFileType
    add rsp, 32
    cmp eax, 2                  ; FILE_TYPE_CHAR = console
    jne su_writefile

    sub rsp, 48
    mov qword ptr [rsp+32], 0   ; lpReserved
    lea r9, [rsp+48+64]         ; lpNumberOfCharsWritten (use local slot)
    mov r8, r12                 ; nNumberOfCharsToWrite
    lea rdx, str_usage
    mov rcx, rbx
    call WriteConsoleW
    add rsp, 48
    jmp su_post

su_writefile:
    sub rsp, 32+8
    mov qword ptr [rsp+32], 0   ; lpOverlapped
    lea r9, [rsp+40+24]         ; lpNumberOfBytesWritten
    mov rax, r12
    shl rax, 1                  ; WCHAR count -> byte count
    mov r8d, eax                ; nNumberOfBytesToWrite
    lea rdx, str_usage
    mov rcx, rbx
    call WriteFile
    add rsp, 32+8
    jmp su_exit

su_post:
    ; Post a fake Enter into the attached console so cmd.exe redraws its
    ; prompt after we exit. Without this, the prompt is printed before our
    ; output appears and the cursor sits there idle until a key is pressed.
    mov ecx, STD_INPUT_HANDLE
    sub rsp, 32
    call GetStdHandle
    add rsp, 32
    test rax, rax
    jz su_exit
    mov rdi, rax

    ; INPUT_RECORD (20 bytes) at [rsp+0..rsp+19] inside our 96-byte frame.
    mov word ptr  [rsp+0],  1   ; EventType = KEY_EVENT
    mov word ptr  [rsp+2],  0   ; padding
    mov dword ptr [rsp+4],  1   ; bKeyDown = TRUE
    mov word ptr  [rsp+8],  1   ; wRepeatCount = 1
    mov word ptr  [rsp+10], 0Dh ; wVirtualKeyCode = VK_RETURN
    mov word ptr  [rsp+12], 1Ch ; wVirtualScanCode
    mov word ptr  [rsp+14], 0Dh ; uChar.UnicodeChar = '\r'
    mov dword ptr [rsp+16], 0   ; dwControlKeyState

    sub rsp, 32
    lea r9, [rsp+32+72]         ; lpNumberOfEventsWritten (scratch slot)
    mov r8d, 1                  ; nLength = 1
    lea rdx, [rsp+32]           ; lpBuffer = INPUT_RECORD
    mov rcx, rdi
    call WriteConsoleInputW
    add rsp, 32

su_exit:
    mov ecx, 1                  ; Exit with status 1 (usage-shown convention)
    sub rsp, 32
    call ExitProcess
    add rsp, 32
ShowUsage endp

end
