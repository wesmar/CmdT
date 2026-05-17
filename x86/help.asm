; ==============================================================================
; CMDT - Run as TrustedInstaller (x86)
; Help / Usage Display Module
;
; Author: Marek Wesolowski (wesmar)
; Purpose: Owns help-switch recognition and the usage banner. Lives in its
;          own translation unit so the rest of the project never has to
;          touch help-related strings or output formatting.
;
; Exported routines (stdcall):
;   IsHelpSwitch     - Return 1 if a wide-string argument matches any of the
;                      supported help-switch spellings, 0 otherwise.
;   ShowUsageAndExit - Free the argv buffer (optional), pick the appropriate
;                      output API for STDOUT (WriteConsoleW for a console,
;                      WriteFile for a redirected file/pipe), write the
;                      usage banner, then call ExitProcess. Never returns.
; ==============================================================================

.586
.model flat, stdcall
option casemap:none

include consts.inc

; --- Win32 dependencies ---
AttachConsole       PROTO :DWORD
GetStdHandle        PROTO :DWORD
GetFileType         PROTO :DWORD
WriteConsoleW       PROTO :DWORD,:DWORD,:DWORD,:DWORD,:DWORD
WriteFile           PROTO :DWORD,:DWORD,:DWORD,:DWORD,:DWORD
WriteConsoleInputW  PROTO :DWORD,:DWORD,:DWORD,:DWORD
LocalFree           PROTO :DWORD
ExitProcess         PROTO :DWORD

; --- String helpers from strutil.asm ---
wcscmp_ci           PROTO :DWORD,:DWORD
wcslen_p            PROTO :DWORD

; ==============================================================================
; CONSTANT STRING DATA
; ==============================================================================
.const

; All supported help-switch spellings. Canonical form is "-help"; the rest are
; convenience aliases so users with muscle memory from POSIX tools or classic
; Windows /? don't get surprised.
str_helpSwitch      dw '-','h','e','l','p',0
str_helpSwitchH     dw '-','h',0
str_helpSwitchDD    dw '-','-','h','e','l','p',0
str_helpSwitchQ     dw '-','?',0
str_helpSwitchSQ    dw '/','?',0
str_helpSwitchSH    dw '/','h',0
str_helpSwitchSHELP dw '/','h','e','l','p',0

; Usage banner. CRLF line endings keep the layout readable in cmd.exe.
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
; IsHelpSwitch - Return 1 if a command-line token is a supported help switch
; ==============================================================================
IsHelpSwitch proc lpArg:DWORD
    invoke wcscmp_ci, lpArg, offset str_helpSwitch
    test eax, eax
    jnz ihs_yes
    invoke wcscmp_ci, lpArg, offset str_helpSwitchH
    test eax, eax
    jnz ihs_yes
    invoke wcscmp_ci, lpArg, offset str_helpSwitchDD
    test eax, eax
    jnz ihs_yes
    invoke wcscmp_ci, lpArg, offset str_helpSwitchQ
    test eax, eax
    jnz ihs_yes
    invoke wcscmp_ci, lpArg, offset str_helpSwitchSQ
    test eax, eax
    jnz ihs_yes
    invoke wcscmp_ci, lpArg, offset str_helpSwitchSH
    test eax, eax
    jnz ihs_yes
    invoke wcscmp_ci, lpArg, offset str_helpSwitchSHELP
    test eax, eax
    jnz ihs_yes
    xor eax, eax
    ret
ihs_yes:
    mov eax, 1
    ret
IsHelpSwitch endp

; ==============================================================================
; ShowUsageAndExit - Print usage to stdout/redirect target and exit
;
; The parent shell may have wired up STDOUT to a console, a file (>out.txt),
; or a pipe (|findstr). WriteConsoleW silently fails on non-console handles,
; which is why a naive `cmdt -help > out.txt` previously produced an empty
; file — so we ask GetFileType which case we're in and pick the right API.
;
; If our STDOUT is invalid (process started without one — e.g. launched by
; Explorer with no parent console) we try to attach to the parent process
; console and re-query; if that also fails, we bail out silently.
; ==============================================================================
ShowUsageAndExit proc uses ebx edi pFreeArgv:DWORD
    LOCAL written:DWORD
    LOCAL usageChars:DWORD
    LOCAL inputRec[20]:BYTE                 ; INPUT_RECORD for fake VK_RETURN

    mov eax, pFreeArgv
    test eax, eax
    jz sue_no_free
    invoke LocalFree, eax
sue_no_free:

    invoke GetStdHandle, STD_OUTPUT_HANDLE
    mov ebx, eax
    test ebx, ebx
    jz sue_try_attach
    cmp ebx, -1
    jne sue_have_stdout

sue_try_attach:
    invoke AttachConsole, ATTACH_PARENT_PROCESS
    invoke GetStdHandle, STD_OUTPUT_HANDLE
    mov ebx, eax
    test ebx, ebx
    jz sue_exit
    cmp ebx, -1
    je sue_exit

sue_have_stdout:
    invoke wcslen_p, offset str_usage
    mov usageChars, eax
    invoke GetFileType, ebx
    cmp eax, FILE_TYPE_CHAR
    jne sue_file

    invoke WriteConsoleW, ebx, offset str_usage, usageChars, addr written, 0

    ; Post a fake Enter into the attached console so cmd.exe redraws its
    ; prompt after we exit. Without this, the prompt is printed before our
    ; output appears and the cursor sits idle until a key is pressed.
    invoke GetStdHandle, STD_INPUT_HANDLE
    test eax, eax
    jz sue_exit
    cmp eax, -1
    je sue_exit
    mov edi, eax

    ; Build INPUT_RECORD: KEY_EVENT, bKeyDown=TRUE, VK_RETURN, ASCII '\r'.
    lea ebx, inputRec
    mov word ptr [ebx+0],  1                ; EventType = KEY_EVENT
    mov word ptr [ebx+2],  0                ; padding
    mov dword ptr [ebx+4], 1                ; bKeyDown = TRUE
    mov word ptr [ebx+8],  1                ; wRepeatCount = 1
    mov word ptr [ebx+10], 0Dh              ; wVirtualKeyCode = VK_RETURN
    mov word ptr [ebx+12], 1Ch              ; wVirtualScanCode
    mov word ptr [ebx+14], 0Dh              ; uChar.UnicodeChar = '\r'
    mov dword ptr [ebx+16], 0               ; dwControlKeyState
    invoke WriteConsoleInputW, edi, ebx, 1, addr written
    jmp sue_exit

sue_file:
    mov ecx, usageChars
    shl ecx, 1
    invoke WriteFile, ebx, offset str_usage, ecx, addr written, 0

sue_exit:
    invoke ExitProcess, 1
ShowUsageAndExit endp

; ==============================================================================
; NudgeConsolePrompt - Ask cmd.exe to redraw its prompt after CLI output
;
; Purpose: GUI-subsystem programs (like cmdt_x86.exe) launched from cmd.exe
;          are not waited on as console programs would be. By the time we
;          finish streaming relay output to STDOUT, cmd.exe has already
;          printed its next prompt one or more lines too high — leaving the
;          cursor parked over our last line of output with no visible
;          prompt. Posting a single VK_RETURN into the console input queue
;          makes cmd redraw a fresh prompt below our output.
;
;          No-op when STDOUT is not a console (redirected file or pipe) —
;          there is nothing to nudge.
;
; Parameters: None
; Returns: None
; ==============================================================================
NudgeConsolePrompt proc uses ebx edi
    LOCAL ncpRec[20]:BYTE
    LOCAL ncpWritten:DWORD

    invoke GetStdHandle, STD_OUTPUT_HANDLE
    test eax, eax
    jz ncp_done
    cmp eax, -1
    je ncp_done

    invoke GetFileType, eax
    cmp eax, FILE_TYPE_CHAR
    jne ncp_done

    invoke GetStdHandle, STD_INPUT_HANDLE
    test eax, eax
    jz ncp_done
    cmp eax, -1
    je ncp_done
    mov edi, eax

    lea ebx, ncpRec
    mov word ptr [ebx+0],  1                ; EventType = KEY_EVENT
    mov word ptr [ebx+2],  0                ; padding
    mov dword ptr [ebx+4], 1                ; bKeyDown = TRUE
    mov word ptr [ebx+8],  1                ; wRepeatCount = 1
    mov word ptr [ebx+10], 0Dh              ; wVirtualKeyCode = VK_RETURN
    mov word ptr [ebx+12], 1Ch              ; wVirtualScanCode
    mov word ptr [ebx+14], 0Dh              ; uChar.UnicodeChar = '\r'
    mov dword ptr [ebx+16], 0               ; dwControlKeyState
    invoke WriteConsoleInputW, edi, ebx, 1, addr ncpWritten

ncp_done:
    ret
NudgeConsolePrompt endp

end
