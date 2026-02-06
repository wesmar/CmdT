; ==============================================================================
; CMDT - Run as TrustedInstaller
; Main Entry Point and Command-Line Processing Module
; 
; Author: Marek Wesołowski (wesmar)
; Purpose: Provides the entry point and command-line argument parsing for the
;          TrustedInstaller privilege elevation utility.
; ==============================================================================

option casemap:none

include consts.inc

; External function declarations
EXTRN CreateMainWindow:PROC
EXTRN RunAsTrustedInstaller:PROC
EXTRN ResolveLnkPath:PROC
EXTRN GetModuleHandleW:PROC
EXTRN GetMessageW:PROC
EXTRN TranslateMessage:PROC
EXTRN DispatchMessageW:PROC
EXTRN IsDialogMessageW:PROC
EXTRN ExitProcess:PROC
EXTRN GetCommandLineW:PROC
EXTRN CommandLineToArgvW:PROC
EXTRN LocalFree:PROC
EXTRN SetFocus:PROC

; ==============================================================================
; CONSTANT STRING DATA
; ==============================================================================
.const
; Registry key for storing application settings and MRU list
str_regKey      dw 'S','o','f','t','w','a','r','e','\','c','m','d','t',0

; Command-line switches for CLI mode
str_cliSwitch1  dw '-','c','l','i',0           ; Short form with dash
str_cliSwitch2  dw '-','-','c','l','i',0       ; Long form with double dash
str_cliSwitch3  dw 'c','l','i',0               ; Bare form without dash

; Switch to request new console window
str_newSwitch   dw '-','n','e','w',0

; File extension for Windows shortcuts
str_extLnk_m    dw '.','l','n','k',0

; Space character for string concatenation
str_space       dw ' ',0

; ==============================================================================
; PRIVILEGE NAME STRINGS
; These strings are used to construct full privilege names by combining with
; the "Se" prefix and "Privilege" suffix (e.g., "SeDebugPrivilege")
; ==============================================================================
privStr_0  dw 'A','s','s','i','g','n','P','r','i','m','a','r','y','T','o','k','e','n',0
privStr_1  dw 'B','a','c','k','u','p',0
privStr_2  dw 'R','e','s','t','o','r','e',0
privStr_3  dw 'D','e','b','u','g',0
privStr_4  dw 'I','m','p','e','r','s','o','n','a','t','e',0
privStr_5  dw 'T','a','k','e','O','w','n','e','r','s','h','i','p',0
privStr_6  dw 'L','o','a','d','D','r','i','v','e','r',0
privStr_7  dw 'S','y','s','t','e','m','E','n','v','i','r','o','n','m','e','n','t',0
privStr_8  dw 'M','a','n','a','g','e','V','o','l','u','m','e',0
privStr_9  dw 'S','e','c','u','r','i','t','y',0
privStr_10 dw 'S','h','u','t','d','o','w','n',0
privStr_11 dw 'S','y','s','t','e','m','t','i','m','e',0
privStr_12 dw 'T','c','b',0
privStr_13 dw 'I','n','c','r','e','a','s','e','Q','u','o','t','a',0
privStr_14 dw 'A','u','d','i','t',0
privStr_15 dw 'C','h','a','n','g','e','N','o','t','i','f','y',0
privStr_16 dw 'U','n','d','o','c','k',0
privStr_17 dw 'C','r','e','a','t','e','T','o','k','e','n',0
privStr_18 dw 'L','o','c','k','M','e','m','o','r','y',0
privStr_19 dw 'C','r','e','a','t','e','P','a','g','e','f','i','l','e',0
privStr_20 dw 'C','r','e','a','t','e','P','e','r','m','a','n','e','n','t',0
privStr_21 dw 'S','y','s','t','e','m','P','r','o','f','i','l','e',0
privStr_22 dw 'P','r','o','f','i','l','e','S','i','n','g','l','e','P','r','o','c','e','s','s',0
privStr_23 dw 'C','r','e','a','t','e','G','l','o','b','a','l',0
privStr_24 dw 'T','i','m','e','Z','o','n','e',0
privStr_25 dw 'C','r','e','a','t','e','S','y','m','b','o','l','i','c','L','i','n','k',0
privStr_26 dw 'I','n','c','r','e','a','s','e','B','a','s','e','P','r','i','o','r','i','t','y',0
privStr_27 dw 'R','e','m','o','t','e','S','h','u','t','d','o','w','n',0
privStr_28 dw 'I','n','c','r','e','a','s','e','W','o','r','k','i','n','g','S','e','t',0
privStr_29 dw 'R','e','l','a','b','e','l',0
privStr_30 dw 'D','e','l','e','g','a','t','e','S','e','s','s','i','o','n','U','s','e','r','I','m','p','e','r','s','o','n','a','t','e',0
privStr_31 dw 'T','r','u','s','t','e','d','C','r','e','d','M','a','n','A','c','c','e','s','s',0
privStr_32 dw 'E','n','a','b','l','e','D','e','l','e','g','a','t','i','o','n',0
privStr_33 dw 'S','y','n','c','A','g','e','n','t',0

; Prefix and suffix for privilege name construction
privPrefix dw 'S','e',0
privSuffix dw 'P','r','i','v','i','l','e','g','e',0

; ==============================================================================
; INITIALIZED DATA SECTION
; ==============================================================================
.data
    align 8

; Privilege table: Array of pointers to privilege name strings
; Used by token manipulation code to enable all available privileges
PUBLIC g_privTable
g_privTable dq offset privStr_0,offset privStr_1,offset privStr_2,offset privStr_3,offset privStr_4,offset privStr_5
            dq offset privStr_6,offset privStr_7,offset privStr_8,offset privStr_9,offset privStr_10,offset privStr_11
            dq offset privStr_12,offset privStr_13,offset privStr_14,offset privStr_15,offset privStr_16,offset privStr_17
            dq offset privStr_18,offset privStr_19,offset privStr_20,offset privStr_21,offset privStr_22,offset privStr_23
            dq offset privStr_24,offset privStr_25,offset privStr_26,offset privStr_27,offset privStr_28,offset privStr_29
            dq offset privStr_30,offset privStr_31,offset privStr_32,offset privStr_33

; Global variables exported to other modules
PUBLIC g_cachedToken, g_tokenTime, g_hwndMain, g_hwndEdit, g_hwndBtn, g_hwndStatus, g_hConsoleOut, g_hInstance
PUBLIC privPrefix, privSuffix

; Cached TrustedInstaller token handle (for performance optimization)
g_cachedToken   dq 0

; Timestamp of when the cached token was obtained (in milliseconds)
g_tokenTime     dd 0
                dd 0            ; Padding for alignment

; Window handles for GUI controls
g_hwndMain      dq 0            ; Main window handle
g_hwndEdit      dq 0            ; ComboBox edit control handle
g_hwndBtn       dq 0            ; Run button handle
g_hwndStatus    dq 0            ; Status label handle

; Console output handle (for CLI mode)
g_hConsoleOut   dq 0

; Flag indicating whether to create a new console window
g_useNewConsole dd 0
                dd 0            ; Padding for alignment

; Application instance handle
g_hInstance     dq 0

; ==============================================================================
; UNINITIALIZED DATA SECTION
; ==============================================================================
.data?

; Export buffer declarations for use by other modules
PUBLIC g_cmdBuf, g_statusBuf, g_filePath, g_argsBuf, g_tempBuf

; Buffer for command line text (520 WCHARs = 1040 bytes)
g_cmdBuf        dw 520 dup(?)

; Buffer for status text (520 WCHARs)
g_statusBuf     dw 520 dup(?)

; Buffer for file paths (520 WCHARs)
g_filePath      dw 520 dup(?)

; Buffer for command-line arguments (520 WCHARs)
g_argsBuf       dw 520 dup(?)

; Temporary buffer for various operations (1040 WCHARs = 2080 bytes)
g_tempBuf       dw 1040 dup(?)

; ==============================================================================
; CODE SECTION
; ==============================================================================
.code

; ==============================================================================
; wcscpy_p - Wide Character String Copy (Private Implementation)
;
; Purpose: Copies a null-terminated wide character string from source to dest
;
; Parameters:
;   RCX = Destination buffer pointer
;   RDX = Source string pointer
;
; Returns: None
;
; Modifies: RAX, RDI, RSI (saved/restored), destination buffer
; ==============================================================================
wcscpy_p proc
    push rsi
    push rdi
    mov rdi, rcx                ; RDI = destination pointer
    mov rsi, rdx                ; RSI = source pointer
@@:
    mov ax, word ptr [rsi]      ; Read wide character from source
    mov word ptr [rdi], ax      ; Write to destination
    test ax, ax                 ; Check for null terminator
    jz @F                       ; Exit if null terminator found
    add rsi, 2                  ; Advance source pointer
    add rdi, 2                  ; Advance destination pointer
    jmp @B                      ; Continue copying
@@:
    pop rdi
    pop rsi
    ret
wcscpy_p endp

; ==============================================================================
; wcscat_p - Wide Character String Concatenate (Private Implementation)
;
; Purpose: Appends source string to the end of destination string
;
; Parameters:
;   RCX = Destination buffer pointer
;   RDX = Source string pointer
;
; Returns: None
;
; Modifies: RAX, RDI, RSI (saved/restored), destination buffer
; ==============================================================================
wcscat_p proc
    push rsi
    push rdi
    mov rdi, rcx                ; RDI = destination pointer
@@:
    cmp word ptr [rdi], 0       ; Check for null terminator
    je @F                       ; Found end of destination string
    add rdi, 2                  ; Move to next character
    jmp @B                      ; Continue searching
@@:
    mov rsi, rdx                ; RSI = source pointer
@@:
    mov ax, word ptr [rsi]      ; Read wide character from source
    mov word ptr [rdi], ax      ; Write to destination
    test ax, ax                 ; Check for null terminator
    jz @F                       ; Exit if null terminator found
    add rsi, 2                  ; Advance source pointer
    add rdi, 2                  ; Advance destination pointer
    jmp @B                      ; Continue copying
@@:
    pop rdi
    pop rsi
    ret
wcscat_p endp

; ==============================================================================
; wcscmp_ci - Wide Character String Compare (Case-Insensitive)
;
; Purpose: Compares two wide character strings ignoring case differences
;
; Parameters:
;   RCX = First string pointer
;   RDX = Second string pointer
;
; Returns:
;   RAX = 1 if strings are equal (case-insensitive), 0 otherwise
;
; Modifies: RAX, RDX, RSI, RDI (saved/restored)
; ==============================================================================
wcscmp_ci proc
    push rsi
    push rdi
    mov rsi, rcx                ; RSI = first string
    mov rdi, rdx                ; RDI = second string
wci_loop:
    mov ax, word ptr [rsi]      ; Load character from first string
    mov dx, word ptr [rdi]      ; Load character from second string
    
    ; Convert first character to lowercase if uppercase
    cmp ax, 'A'
    jb wci_skip1
    cmp ax, 'Z'
    ja wci_skip1
    add ax, 32                  ; Convert A-Z to a-z
wci_skip1:
    
    ; Convert second character to lowercase if uppercase
    cmp dx, 'A'
    jb wci_skip2
    cmp dx, 'Z'
    ja wci_skip2
    add dx, 32                  ; Convert A-Z to a-z
wci_skip2:
    
    cmp ax, dx                  ; Compare normalized characters
    jne not_eq                  ; Characters differ
    test ax, ax                 ; Check if end of strings
    jz equal                    ; Both null terminators reached
    add rsi, 2                  ; Advance first string pointer
    add rdi, 2                  ; Advance second string pointer
    jmp wci_loop                ; Continue comparison
equal:
    pop rdi
    pop rsi
    mov rax, 1                  ; Return 1 (strings equal)
    ret
not_eq:
    pop rdi
    pop rsi
    xor rax, rax                ; Return 0 (strings differ)
    ret
wcscmp_ci endp

; ==============================================================================
; skip_spaces - Skip Leading Whitespace in Wide String
;
; Purpose: Advances a string pointer past any leading space characters
;
; Parameters:
;   RCX = String pointer
;
; Returns:
;   RAX = Pointer to first non-space character
;
; Modifies: RAX
; ==============================================================================
skip_spaces proc
    mov rax, rcx                ; RAX = input pointer
@@:
    cmp word ptr [rax], ' '     ; Check if current character is space
    jne @F                      ; Exit if non-space found
    add rax, 2                  ; Skip this space
    jmp @B                      ; Continue checking
@@:
    ret
skip_spaces endp

; ==============================================================================
; wcslen_p - Wide Character String Length (Private Implementation)
;
; Purpose: Calculates the length of a null-terminated wide character string
;
; Parameters:
;   RCX = String pointer
;
; Returns:
;   RAX = Number of wide characters (excluding null terminator)
;
; Modifies: RAX, R8
; ==============================================================================
wcslen_p proc
    mov rax, rcx                ; RAX = string pointer
    xor r8, r8                  ; R8 = character count
@@:
    cmp word ptr [rax + r8*2], 0 ; Check for null terminator
    je @F                       ; Exit if found
    inc r8                      ; Increment count
    jmp @B                      ; Continue counting
@@:
    mov rax, r8                 ; Return count in RAX
    ret
wcslen_p endp

; ==============================================================================
; mainCRTStartup - Application Entry Point
;
; Purpose: Main entry point for the application. Parses command-line arguments
;          and either runs in CLI mode or launches the GUI window.
;
; Command-line syntax:
;   - GUI mode: <exe> (no arguments)
;   - CLI mode: <exe> -cli <command>
;   - CLI mode with new console: <exe> -cli -new <command>
;
; Returns: Does not return (calls ExitProcess)
;
; Stack frame: 312 bytes local variables
; ==============================================================================
mainCRTStartup proc frame
    push rbp
    .pushreg rbp
    mov rbp, rsp
    .setframe rbp, 0
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
    sub rsp, 312
    .allocstack 312
    .endprolog

    cld                         ; Clear direction flag for string operations

    ; Get the full command line string
    sub rsp, 32
    call GetCommandLineW
    add rsp, 32
    mov r12, rax                ; R12 = command line string

    ; Parse command line into argv array
    lea rdx, [rbp-64]           ; RDX = pointer to argc output
    mov rcx, r12                ; RCX = command line string
    sub rsp, 32
    call CommandLineToArgvW
    add rsp, 32
    mov r13, rax                ; R13 = argv array

    ; Check if at least 2 arguments (exe + switch)
    mov eax, dword ptr [rbp-64] ; EAX = argc
    cmp eax, 2
    jl mode_gui_free            ; Less than 2 args: GUI mode

    ; Get second argument (first after exe path)
    mov rax, [r13+8]            ; RAX = argv[1]
    mov r14, rax                ; R14 = argv[1]

    ; Check if argv[1] matches "-cli"
    lea rdx, str_cliSwitch1
    mov rcx, r14
    call wcscmp_ci
    test rax, rax
    jnz mode_cli_found

    ; Check if argv[1] matches "--cli"
    lea rdx, str_cliSwitch2
    mov rcx, r14
    call wcscmp_ci
    test rax, rax
    jnz mode_cli_found

    ; Check if argv[1] matches "cli"
    lea rdx, str_cliSwitch3
    mov rcx, r14
    call wcscmp_ci
    test rax, rax
    jnz mode_cli_found

    jmp mode_gui_free           ; No CLI switch found: GUI mode

mode_cli_found:
    ; CLI mode detected - check for minimum arguments
    mov eax, dword ptr [rbp-64] ; EAX = argc
    cmp eax, 3
    jl cli_no_cmd_free          ; Error: no command specified

    ; Check if argv[2] is "-new" (new console flag)
    mov rax, [r13+16]           ; RAX = argv[2]
    lea rdx, str_newSwitch
    mov rcx, rax
    call wcscmp_ci
    test rax, rax
    jz cli_no_new_flag          ; Not "-new"

    ; "-new" flag found: need at least 4 args (exe, switch, -new, command)
    mov eax, dword ptr [rbp-64]
    cmp eax, 4
    jl cli_no_cmd_free          ; Error: no command after -new
    mov dword ptr g_useNewConsole, 1 ; Set new console flag
    jmp cli_free_and_setup

cli_no_new_flag:
    mov dword ptr g_useNewConsole, 0 ; Clear new console flag

cli_free_and_setup:
    ; Free argv array allocated by CommandLineToArgvW
    mov rcx, r13
    sub rsp, 32
    call LocalFree
    add rsp, 32
    jmp mode_cli_setup

mode_gui_free:
    ; Free argv array and proceed to GUI mode
    mov rcx, r13
    sub rsp, 32
    call LocalFree
    add rsp, 32
    jmp mode_gui

cli_no_cmd_free:
    ; Error: insufficient arguments for CLI mode
    mov rcx, r13
    sub rsp, 32
    call LocalFree
    add rsp, 32
    mov ecx, 1                  ; Exit code 1 (error)
    sub rsp, 32
    call ExitProcess
    add rsp, 32

mode_cli_setup:
    ; Parse the raw command line to extract the actual command
    ; Need to skip: exe path, CLI switch, and optionally "-new"
    
    sub rsp, 32
    call GetCommandLineW
    add rsp, 32
    mov rsi, rax                ; RSI = command line pointer
    xor r8, r8
    xor rdi, rdi                ; RDI = quote state flag

skip_exe_loop:
    ; Skip past the executable path (which may be quoted)
    mov ax, word ptr [rsi]
    test ax, ax
    jz cli_failed_setup         ; Unexpected end of command line
    cmp ax, '"'
    jne @F
    xor edi, 1                  ; Toggle quote state
@@:
    cmp ax, ' '
    jne @F
    test edi, edi               ; Are we inside quotes?
    jnz @F                      ; Yes: space is part of path
    add rsi, 2                  ; No: space ends exe path
    jmp skip_switch_init
@@:
    add rsi, 2                  ; Continue scanning
    jmp skip_exe_loop

skip_switch_init:
    ; Skip leading spaces before CLI switch
    mov rcx, rsi
    call skip_spaces
    mov rsi, rax

    xor edi, edi
skip_switch_loop:
    ; Skip the CLI switch argument
    mov ax, word ptr [rsi]
    test ax, ax
    jz cli_failed_setup
    cmp ax, ' '
    jne @F
    add rsi, 2
    jmp after_switch
@@:
    add rsi, 2
    jmp skip_switch_loop

after_switch:
    ; Skip spaces after CLI switch
    mov rcx, rsi
    call skip_spaces
    mov rsi, rax

    ; Check if we need to skip the "-new" token
    cmp dword ptr g_useNewConsole, 0
    je run_command              ; No "-new" flag: proceed to command

skip_new_token:
    ; Skip the "-new" token
    mov ax, word ptr [rsi]
    test ax, ax
    jz cli_failed_setup
    cmp ax, ' '
    jne @F
    add rsi, 2                  ; Skip space after "-new"
    jmp run_command
@@:
    add rsi, 2
    jmp skip_new_token

run_command:
    ; Skip spaces before the actual command
    mov rcx, rsi
    call skip_spaces
    mov rsi, rax                ; RSI now points to the command

    ; Check if the command is a .lnk file (shortcut)
    mov rcx, rsi
    call wcslen_p
    mov r8, rax                 ; R8 = command length
    cmp r8, 4                   ; Minimum length for ".lnk"
    jl run_no_lnk               ; Too short to be .lnk

    ; Scan for space or quote to find end of path
    mov rdi, rsi
    xor rbx, rbx                ; RBX = pointer to first space (if found)
    xor r9, r9                  ; R9 = quote state
find_space_or_quote:
    mov ax, word ptr [rdi]
    test ax, ax
    jz check_lnk_ext            ; End of string
    cmp ax, '"'
    jne @F
    xor r9d, 1                  ; Toggle quote state
@@:
    cmp ax, ' '
    jne @F
    test r9d, r9d               ; Inside quotes?
    jnz @F                      ; Yes: ignore space
    mov rbx, rdi                ; No: mark first space position
    jmp check_lnk_ext            ; Stop at first space
@@:
    add rdi, 2
    jmp find_space_or_quote

check_lnk_ext:
    ; Check if path (before first space) ends with .lnk
    test rbx, rbx
    jz check_whole_path         ; No space found: check entire string

    ; Compute clean bounds (strip surrounding quotes from path)
    mov r10, rsi                ; R10 = start of path
    mov r11, rbx                ; R11 = end of path (space position)
    cmp word ptr [r10], '"'
    jne @F
    add r10, 2                  ; Skip opening quote
@@:
    cmp word ptr [r11-2], '"'
    jne @F
    sub r11, 2                  ; Skip closing quote
@@:
    mov rcx, r11
    sub rcx, r10
    shr rcx, 1                  ; RCX = path length in characters
    cmp rcx, 4
    jl run_no_lnk               ; Too short for .lnk

    ; Check if last 4 characters are ".lnk"
    lea rdi, [r10 + rcx*2 - 8]  ; Point to last 4 chars
    lea rdx, str_extLnk_m
    mov rcx, rdi
    call wcscmp_ci
    test rax, rax
    jz run_no_lnk               ; Not a .lnk file

    ; Save original RSI and RBX
    push rsi
    push rbx

    ; Copy clean path (without quotes) to g_filePath
    mov rcx, r11
    sub rcx, r10
    shr rcx, 1                  ; RCX = number of characters
    mov r8, rcx
    mov rsi, r10
    lea rdi, g_filePath
@@:
    test r8, r8
    jz @F
    mov ax, word ptr [rsi]
    mov word ptr [rdi], ax
    add rsi, 2
    add rdi, 2
    dec r8
    jmp @B
@@:
    mov word ptr [rdi], 0       ; Null terminate

    ; Extract arguments after the path
    pop rbx
    add rbx, 2                  ; Skip the space
    mov rcx, rbx
    call skip_spaces
    mov rsi, rax
    mov rdx, rsi
    lea rcx, g_argsBuf
    xchg rcx, rdx
    call wcscpy_p               ; Copy arguments to g_argsBuf

    ; Clear g_tempBuf (for shortcut arguments from .lnk)
    lea rdi, g_tempBuf
    xor rax, rax
    mov rcx, 260
@@:
    test rcx, rcx
    jz @F
    mov qword ptr [rdi], rax
    add rdi, 8
    dec rcx
    jmp @B
@@:

    ; Resolve the .lnk file to get target path and built-in arguments
    lea r8, g_tempBuf           ; R8 = buffer for .lnk arguments
    lea rdx, g_cmdBuf           ; RDX = buffer for target path
    lea rcx, g_filePath         ; RCX = .lnk file path
    sub rsp, 32
    call ResolveLnkPath
    add rsp, 32
    test rax, rax
    pop rsi
    jz run_no_lnk               ; Resolution failed

    ; Build final command: target + .lnk args + user args
    lea rcx, g_cmdBuf
    call wcslen_p
    test rax, rax
    jz use_args_only            ; No target path

    ; Append space after target path
    lea rdi, g_cmdBuf
    lea rdi, [rdi + rax*2]
    mov word ptr [rdi], ' '
    add rdi, 2
    mov word ptr [rdi], 0

use_args_only:
    ; Append .lnk arguments
    lea rdx, g_tempBuf
    lea rcx, g_cmdBuf
    call wcscat_p

    ; Check if user provided additional arguments
    lea rcx, g_argsBuf
    call wcslen_p
    test rax, rax
    jz run_resolved             ; No user arguments

    ; Append space and user arguments
    lea rdx, str_space
    lea rcx, g_cmdBuf
    call wcscat_p
    lea rdx, g_argsBuf
    lea rcx, g_cmdBuf
    call wcscat_p
    jmp run_resolved

check_whole_path:
    ; No space found: entire command might be a .lnk file
    ; Compute clean bounds (strip surrounding quotes)
    mov r10, rsi                ; R10 = start
    mov r11, r8                 ; R11 = length
    cmp word ptr [r10], '"'
    jne @F
    add r10, 2                  ; Skip opening quote
    dec r11
@@:
    cmp r11, 1
    jl run_no_lnk
    cmp word ptr [r10 + r11*2 - 2], '"'
    jne @F
    dec r11                     ; Exclude closing quote
@@:
    cmp r11, 4
    jl run_no_lnk               ; Too short for .lnk

    ; Check if last 4 characters are ".lnk"
    lea rdi, [r10 + r11*2 - 8]
    lea rdx, str_extLnk_m
    mov rcx, rdi
    call wcscmp_ci
    test rax, rax
    jz run_no_lnk               ; Not a .lnk file

    ; Clear g_tempBuf
    lea rdi, g_tempBuf
    xor rax, rax
    mov rcx, 260
@@:
    test rcx, rcx
    jz @F
    mov qword ptr [rdi], rax
    add rdi, 8
    dec rcx
    jmp @B
@@:

    ; Copy clean path to g_filePath
    mov rcx, r11                ; RCX = length
    lea rdi, g_filePath
    mov rdx, r10
@@:
    test rcx, rcx
    jz @F
    mov ax, word ptr [rdx]
    mov word ptr [rdi], ax
    add rdx, 2
    add rdi, 2
    dec rcx
    jmp @B
@@:
    mov word ptr [rdi], 0       ; Null terminate

    ; Resolve the .lnk file
    lea r8, g_tempBuf
    lea rdx, g_cmdBuf
    lea rcx, g_filePath
    sub rsp, 32
    call ResolveLnkPath
    add rsp, 32
    test rax, rax
    jz run_no_lnk               ; Resolution failed

    ; Build command from resolved target and arguments
    lea rcx, g_cmdBuf
    call wcslen_p
    test rax, rax
    jz use_lnk_args_only        ; No target path

    ; Append space after target path
    lea rdi, g_cmdBuf
    lea rdi, [rdi + rax*2]
    mov word ptr [rdi], ' '
    add rdi, 2
    mov word ptr [rdi], 0

use_lnk_args_only:
    ; Append .lnk arguments
    lea rdx, g_tempBuf
    lea rcx, g_cmdBuf
    call wcscat_p
    jmp run_resolved

run_resolved:
    ; Execute the resolved command as TrustedInstaller
    mov edx, dword ptr g_useNewConsole
    lea rcx, g_cmdBuf
    sub rsp, 32
    call RunAsTrustedInstaller
    add rsp, 32
    jmp run_check_result

run_no_lnk:
    ; Execute the original command (not a .lnk file)
    mov edx, dword ptr g_useNewConsole
    mov rcx, rsi
    sub rsp, 32
    call RunAsTrustedInstaller
    add rsp, 32

run_check_result:
    ; Check execution result
    test rax, rax
    jz cli_failed               ; Execution failed

    ; Success: exit with code 0
    xor ecx, ecx
    sub rsp, 32
    call ExitProcess
    add rsp, 32

cli_failed_setup:
cli_failed:
    ; Failure: exit with code 1
    mov ecx, 1
    sub rsp, 32
    call ExitProcess
    add rsp, 32

mode_gui:
    ; GUI Mode: Create and show main window
    
    ; Get application instance handle
    xor ecx, ecx
    sub rsp, 32
    call GetModuleHandleW
    add rsp, 32
    mov g_hInstance, rax

    ; Create the main application window
    mov rcx, rax
    sub rsp, 32
    call CreateMainWindow
    add rsp, 32
    test rax, rax
    jz exit_app                 ; Window creation failed

msg_loop:
    ; Main message loop
    lea rcx, [rbp-200]          ; RCX = &msg structure
    xor edx, edx                ; EDX = hWnd (NULL = all windows)
    xor r8d, r8d                ; R8D = wMsgFilterMin (0 = all messages)
    xor r9d, r9d                ; R9D = wMsgFilterMax (0 = all messages)
    sub rsp, 32
    call GetMessageW
    add rsp, 32
    test eax, eax
    jz exit_app                 ; WM_QUIT received

    ; Check for ESC key to exit application
    mov eax, dword ptr [rbp-200+8]  ; EAX = message
    cmp eax, WM_KEYDOWN
    jne msg_not_esc
    mov rax, qword ptr [rbp-200+16] ; RAX = wParam
    cmp eax, VK_ESCAPE
    je exit_app                 ; ESC pressed: exit

msg_not_esc:
    ; Process dialog messages (for tab navigation, etc.)
    lea rdx, [rbp-200]
    mov rcx, g_hwndMain
    sub rsp, 32
    call IsDialogMessageW
    add rsp, 32
    test eax, eax
    jnz msg_loop                ; Message was processed

    ; Translate and dispatch regular window messages
    lea rcx, [rbp-200]
    sub rsp, 32
    call TranslateMessage
    add rsp, 32

    lea rcx, [rbp-200]
    sub rsp, 32
    call DispatchMessageW
    add rsp, 32
    jmp msg_loop

exit_app:
    ; Exit the application
    xor ecx, ecx
    sub rsp, 32
    call ExitProcess
    add rsp, 32

    ; Cleanup and return (unreachable, but proper epilog)
    add rsp, 312
    pop r15
    pop r14
    pop r13
    pop r12
    pop rdi
    pop rsi
    pop rbx
    pop rbp
    ret
mainCRTStartup endp

end
