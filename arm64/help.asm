; ==============================================================================
; CMDT - Run as TrustedInstaller
; Help / Usage Display Module
;
; Author: Marek Wesolowski (wesmar)
; Ported to: Native ARM64 (armasm64) for Windows on ARM
;
; Purpose: Owns everything related to "show me the options" -- recognizing all
;          supported help switches and rendering the usage banner. A faithful
;          port of x64/help.asm: identical switches, identical banner text, and
;          the same WriteConsoleW (console) / WriteFile (redirected) selection
;          so output is correct in both a terminal and a pipe/file.
;
; Exported routines:
;   HelpCheckAndExit(argc, argv) - If argv[1] is a known help variant, display
;                                  usage and exit (never returns). Otherwise
;                                  returns to the caller untouched.
;   ShowUsage(argv)              - Free argv (if non-NULL) and print the banner,
;                                  then ExitProcess(1). Never returns.
; ==============================================================================

; ==============================================================================
; EXTERNAL REFERENCES
; ==============================================================================
    IMPORT GetStdHandle
    IMPORT GetFileType
    IMPORT WriteConsoleW
    IMPORT WriteFile
    IMPORT LocalFree
    IMPORT ExitProcess
    IMPORT wcscmp_ci
    IMPORT wcslen_p

; ==============================================================================
; CONSTANT STRING DATA (READ-ONLY)
; ==============================================================================
    AREA |.const|, DATA, READONLY, ALIGN=3

    EXPORT HelpCheckAndExit
    EXPORT ShowUsage

; All supported help-switch spellings. Canonical form is "-help"; the rest are
; convenience aliases (POSIX -h/--help, classic Windows /? etc.).
str_helpSwitch      DCW '-','h','e','l','p', 0
str_helpSwitchH     DCW '-','h', 0
str_helpSwitchDD    DCW '-','-','h','e','l','p', 0
str_helpSwitchQ     DCW '-','?', 0
str_helpSwitchSQ    DCW '/','?', 0
str_helpSwitchSH    DCW '/','h', 0
str_helpSwitchSHELP DCW '/','h','e','l','p', 0

; Usage banner. CRLF (13,10) line endings keep the layout readable in cmd.exe;
; WriteConsoleW treats them as a single line break.
str_usage
    DCW 13,10
    DCW 'U','s','a','g','e',':',' ','c','m','d','t','.','e','x','e',' ','[','o','p','t','i','o','n',']',13,10
    DCW 13,10
    DCW ' ',' ','-','c','l','i',' ','<','c','m','d','>',' ',' ',' ',' ',' ',' ',' ',' ',' ',' ','R','u','n',' ','c','o','m','m','a','n','d',13,10
    DCW ' ',' ','-','c','l','i',' ','-','n','e','w',' ','<','c','m','d','>',' ',' ',' ',' ',' ','R','u','n',' ','i','n',' ','n','e','w',' ','c','o','n','s','o','l','e',13,10
    DCW ' ',' ','-','i','n','s','t','a','l','l',' ',' ',' ',' ',' ',' ',' ',' ',' ',' ',' ',' ','A','d','d',' ','c','o','n','t','e','x','t',' ','m','e','n','u',13,10
    DCW ' ',' ','-','u','n','i','n','s','t','a','l','l',' ',' ',' ',' ',' ',' ',' ',' ',' ',' ','R','e','m','o','v','e',' ','c','o','n','t','e','x','t',' ','m','e','n','u',13,10
    DCW ' ',' ','-','s','h','i','f','t',' ',' ',' ',' ',' ',' ',' ',' ',' ',' ',' ',' ',' ',' ','H','o','o','k',' ','s','e','t','h','c','.','e','x','e',13,10
    DCW ' ',' ','-','u','n','s','h','i','f','t',' ',' ',' ',' ',' ',' ',' ',' ',' ',' ',' ',' ','U','n','h','o','o','k',' ','s','e','t','h','c','.','e','x','e',13,10
    DCW ' ',' ','-','h','i','s','t','o','r','y','-','c','l','e','a','r',' ',' ',' ',' ',' ',' ','C','l','e','a','r',' ','s','a','v','e','d',' ','c','o','m','m','a','n','d',' ','h','i','s','t','o','r','y',13,10
    DCW ' ',' ','-','h','e','l','p',',',' ','-','h',',',' ','-','-','h','e','l','p',',',' ','-','?',',',' ','/','?',' ',' ','S','h','o','w',' ','t','h','i','s',' ','h','e','l','p',13,10
    DCW 13,10
    DCW ' ',' ','N','o',' ','a','r','g','s',' ','t','o',' ','s','t','a','r','t',' ','G','U','I','.',13,10
    DCW 0

; ==============================================================================
; CODE SECTION
; ==============================================================================
    AREA |.text|, CODE, READONLY, ALIGN=4

; ==============================================================================
; HelpCheckAndExit - Detect help switches and hand off to ShowUsage.
;   x0 = argc (low DWORD meaningful)
;   x1 = argv (LocalFree-able pointer, or NULL)
; Returns to caller only when no help switch is present.
; ==============================================================================
HelpCheckAndExit PROC
    STP x29, x30, [sp, #-16]!
    MOV x29, sp
    STP x19, x20, [sp, #-16]!

    MOV x19, x1                     ; x19 = argv (preserved across calls)

    CMP w0, #2
    B.LT hce_no_help
    CBZ x19, hce_no_help

    LDR x20, [x19, #8]             ; x20 = argv[1] (preserved across calls)

    ; wcscmp_ci returns 1 on match -> CBNZ (x64 was 'jnz hcae_match').
    ADRP x1, str_helpSwitch
    ADD x1, x1, str_helpSwitch
    MOV x0, x20
    BL wcscmp_ci
    CBNZ w0, hce_match

    ADRP x1, str_helpSwitchH
    ADD x1, x1, str_helpSwitchH
    MOV x0, x20
    BL wcscmp_ci
    CBNZ w0, hce_match

    ADRP x1, str_helpSwitchDD
    ADD x1, x1, str_helpSwitchDD
    MOV x0, x20
    BL wcscmp_ci
    CBNZ w0, hce_match

    ADRP x1, str_helpSwitchQ
    ADD x1, x1, str_helpSwitchQ
    MOV x0, x20
    BL wcscmp_ci
    CBNZ w0, hce_match

    ADRP x1, str_helpSwitchSQ
    ADD x1, x1, str_helpSwitchSQ
    MOV x0, x20
    BL wcscmp_ci
    CBNZ w0, hce_match

    ADRP x1, str_helpSwitchSH
    ADD x1, x1, str_helpSwitchSH
    MOV x0, x20
    BL wcscmp_ci
    CBNZ w0, hce_match

    ADRP x1, str_helpSwitchSHELP
    ADD x1, x1, str_helpSwitchSHELP
    MOV x0, x20
    BL wcscmp_ci
    CBNZ w0, hce_match

hce_no_help
    LDP x19, x20, [sp], #16
    LDP x29, x30, [sp], #16
    RET

hce_match
    ; Help switch found: hand argv to ShowUsage to free + display, then exit.
    MOV x0, x19
    BL ShowUsage                    ; never returns
    BRK #0                          ; unreachable
HelpCheckAndExit ENDP

; ==============================================================================
; ShowUsage - Print the usage banner to stdout and exit.
;   x0 = argv buffer to free (or NULL)
; Console handles use WriteConsoleW (native UTF-16); files/pipes use WriteFile
; with raw UTF-16LE bytes. Never returns (ExitProcess with code 1).
; ==============================================================================
ShowUsage PROC
    STP x29, x30, [sp, #-16]!
    MOV x29, sp
    STP x19, x20, [sp, #-16]!
    SUB sp, sp, #16                 ; local: lpNumberOfCharsWritten at [sp+0]

    ; Free argv if the caller handed us one.
    CBZ x0, su_no_free
    BL LocalFree
su_no_free

    ; STD_OUTPUT_HANDLE (-11). Bail silently if NULL / INVALID_HANDLE_VALUE.
    MOVN x0, #10
    BL GetStdHandle
    MOV x19, x0
    CBZ x19, su_exit
    MOVN x9, #0                     ; -1 = INVALID_HANDLE_VALUE
    CMP x19, x9
    B.EQ su_exit

    ; WCHAR count of the banner (both output paths need it).
    ADRP x0, str_usage
    ADD x0, x0, str_usage
    BL wcslen_p
    MOV x20, x0

    ; Console -> WriteConsoleW; otherwise WriteFile (file/pipe).
    MOV x0, x19
    BL GetFileType
    CMP w0, #2                      ; FILE_TYPE_CHAR
    B.NE su_writefile

    MOV x0, x19
    ADRP x1, str_usage
    ADD x1, x1, str_usage
    MOV x2, x20                     ; nNumberOfCharsToWrite
    ADD x3, sp, #0                  ; lpNumberOfCharsWritten
    MOV x4, XZR                     ; lpReserved
    BL WriteConsoleW
    B su_exit

su_writefile
    MOV x0, x19
    ADRP x1, str_usage
    ADD x1, x1, str_usage
    LSL x2, x20, #1                 ; WCHAR count -> byte count
    ADD x3, sp, #0                  ; lpNumberOfBytesWritten
    MOV x4, XZR                     ; lpOverlapped
    BL WriteFile

su_exit
    MOV w0, #1                      ; usage-shown convention
    BL ExitProcess
    BRK #0                          ; unreachable
ShowUsage ENDP

    END
