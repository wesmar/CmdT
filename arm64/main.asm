; ==============================================================================
; CMDT - Run as TrustedInstaller
; Main Entry Point and Command-Line Processing Module
; 
; Author: Marek Wesołowski (wesmar)
; Ported to: Native ARM64 (armasm64) for Windows on ARM
; Purpose: Provides the entry point and command-line argument parsing for the
;          TrustedInstaller privilege elevation utility.
; ==============================================================================

    AREA |.text|, CODE, READONLY, ALIGN=4

; ==============================================================================
; EXTERNAL FUNCTION DECLARATIONS
; ==============================================================================
    IMPORT CreateMainWindow
    IMPORT RunAsTrustedInstaller
    IMPORT ResolveLnkPath
    IMPORT GetModuleHandleW
    IMPORT GetMessageW
    IMPORT TranslateMessage
    IMPORT DispatchMessageW
    IMPORT IsDialogMessageW
    IMPORT ExitProcess
    IMPORT GetCommandLineW
    IMPORT CommandLineToArgvW
    IMPORT LocalFree
    IMPORT SetFocus
    IMPORT IsUserAnAdmin
    IMPORT ShellExecuteExW
    IMPORT GetModuleFileNameW
    IMPORT AttachConsole
    IMPORT GetStdHandle
    IMPORT GetFileType
    IMPORT WaitForSingleObject
    IMPORT CloseHandle
    IMPORT CreateFileW
    IMPORT ReadFile
    IMPORT WriteFile
    IMPORT DeleteFileW
    IMPORT GetTempPathW
    IMPORT GetTempFileNameW
    IMPORT RegDeleteTreeW
    IMPORT GetConsoleWindow
    IMPORT ShowWindow
    IMPORT GetConsoleProcessList

; MRU registry key path string (defined in window.asm)
    IMPORT str_regKey

; String helpers (defined in strutil.asm)
    IMPORT DecryptWideStr
    IMPORT wcscpy_p
    IMPORT wcscat_p
    IMPORT wcscmp_ci
    IMPORT wcscmp_token
    IMPORT skip_spaces
    IMPORT wcslen_p

; Help / usage display (defined in help.asm)
    IMPORT HelpCheckAndExit
    IMPORT ShowUsage

; Installation / hook management (defined in install.asm)
    IMPORT InstallContextMenu
    IMPORT UninstallContextMenu
    IMPORT InstallShift
    IMPORT UninstallShift

; Non-admin output relay (defined in relay.asm)
    IMPORT NonAdminRelayLaunch
    IMPORT AdminRelayLaunch

; CLI / file-run dispatch (defined in cli.asm)
    IMPORT mode_cli_found
    IMPORT mode_file_run

; ==============================================================================
; CONSTANT STRING DATA (READ-ONLY)
; ==============================================================================
    AREA |.const|, DATA, READONLY, ALIGN=3

    EXPORT str_runas
    EXPORT str_newSwitch
    EXPORT str_extLnk_m
    EXPORT str_space
    EXPORT str_outfileFlag
    EXPORT str_errfileFlag

str_cliSwitch1      DCW '-','c','l','i', 0
str_outfileFlag     DCW '-','o','u','t','f','i','l','e', 0
str_errfileFlag     DCW '-','e','r','r','f','i','l','e', 0
str_newSwitch       DCW '-','n','e','w', 0
str_extLnk_m        DCW '.','l','n','k', 0
str_space           DCW ' ', 0
str_runas           DCW 'r','u','n','a','s', 0
str_installSwitch   DCW '-','i','n','s','t','a','l','l', 0
str_uninstallSwitch DCW '-','u','n','i','n','s','t','a','l','l', 0
str_shiftSwitch     DCW '-','s','h','i','f','t', 0
str_unshiftSwitch   DCW '-','u','n','s','h','i','f','t', 0
str_historyClearSwitch DCW '-','h','i','s','t','o','r','y','-','c','l','e','a','r', 0

; ==============================================================================
; PRIVILEGE NAME STRINGS
; ==============================================================================
privStr_0   DCW 'A','s','s','i','g','n','P','r','i','m','a','r','y','T','o','k','e','n', 0
privStr_1   DCW 'B','a','c','k','u','p', 0
privStr_2   DCW 'R','e','s','t','o','r','e', 0
privStr_3   DCW 'D','e','b','u','g', 0
privStr_4   DCW 'I','m','p','e','r','s','o','n','a','t','e', 0
privStr_5   DCW 'T','a','k','e','O','w','n','e','r','s','h','i','p', 0
privStr_6   DCW 'L','o','a','d','D','r','i','v','e','r', 0
privStr_7   DCW 'S','y','s','t','e','m','E','n','v','i','r','o','n','m','e','n','t', 0
privStr_8   DCW 'M','a','n','a','g','e','V','o','l','u','m','e', 0
privStr_9   DCW 'S','e','c','u','r','i','t','y', 0
privStr_10  DCW 'S','h','u','t','d','o','w','n', 0
privStr_11  DCW 'S','y','s','t','e','m','t','i','m','e', 0
privStr_12  DCW 'T','c','b', 0
privStr_13  DCW 'I','n','c','r','e','a','s','e','Q','u','o','t','a', 0
privStr_14  DCW 'A','u','d','i','t', 0
privStr_15  DCW 'C','h','a','n','g','e','N','o','t','i','f','y', 0
privStr_16  DCW 'U','n','d','o','c','k', 0
privStr_17  DCW 'C','r','e','a','t','e','T','o','k','e','n', 0
privStr_18  DCW 'L','o','c','k','M','e','m','o','r','y', 0
privStr_19  DCW 'C','r','e','a','t','e','P','a','g','e','f','i','l','e', 0
privStr_20  DCW 'C','r','e','a','t','e','P','e','r','m','a','n','e','n','t', 0
privStr_21  DCW 'S','y','s','t','e','m','P','r','o','f','i','l','e', 0
privStr_22  DCW 'P','r','o','f','i','l','e','S','i','n','g','l','e','P','r','o','c','e','s','s', 0
privStr_23  DCW 'C','r','e','a','t','e','G','l','o','b','a','l', 0
privStr_24  DCW 'T','i','m','e','Z','o','n','e', 0
privStr_25  DCW 'C','r','e','a','t','e','S','y','m','b','o','l','i','c','L','i','n','k', 0
privStr_26  DCW 'I','n','c','r','e','a','s','e','B','a','s','e','P','r','i','o','r','i','t','y', 0
privStr_27  DCW 'R','e','m','o','t','e','S','h','u','t','d','o','w','n', 0
privStr_28  DCW 'I','n','c','r','e','a','s','e','W','o','r','k','i','n','g','S','e','t', 0
privStr_29  DCW 'R','e','l','a','b','e','l', 0
privStr_30  DCW 'D','e','l','e','g','a','t','e','S','e','s','s','i','o','n','U','s','e','r','I','m','p','e','r','s','o','n','a','t','e', 0
privStr_31  DCW 'T','r','u','s','t','e','d','C','r','e','d','M','a','n','A','c','c','e','s','s', 0
privStr_32  DCW 'E','n','a','b','l','e','D','e','l','e','g','a','t','i','o','n', 0
privStr_33  DCW 'S','y','n','c','A','g','e','n','t', 0

privPrefix  DCW 'S','e', 0
privSuffix  DCW 'P','r','i','v','i','l','e','g','e', 0

; ==============================================================================
; INITIALIZED DATA SECTION
; ==============================================================================
    AREA |.data|, DATA, READWRITE, ALIGN=3

; Privilege table: Array of 64-bit pointers to privilege name strings
    EXPORT g_privTable
g_privTable
    DCQ privStr_0, privStr_1, privStr_2, privStr_3, privStr_4, privStr_5
    DCQ privStr_6, privStr_7, privStr_8, privStr_9, privStr_10, privStr_11
    DCQ privStr_12, privStr_13, privStr_14, privStr_15, privStr_16, privStr_17
    DCQ privStr_18, privStr_19, privStr_20, privStr_21, privStr_22, privStr_23
    DCQ privStr_24, privStr_25, privStr_26, privStr_27, privStr_28, privStr_29
    DCQ privStr_30, privStr_31, privStr_32, privStr_33

; Global variables exported to other modules
    EXPORT g_cachedToken
    EXPORT g_tokenTime
    EXPORT g_hwndMain
    EXPORT g_hwndEdit
    EXPORT g_hwndBtn
    EXPORT g_hwndStatus
    EXPORT g_hConsoleOut
    EXPORT g_hInstance
    EXPORT g_useNewConsole
    EXPORT g_historyEnabled
    EXPORT g_hMenuFile
    EXPORT privPrefix
    EXPORT privSuffix
    EXPORT g_relayHandle
    EXPORT g_relayErrHandle
    EXPORT g_childExitCode

g_cachedToken     DCQ 0
g_tokenTime       DCD 0, 0                  ; 8-byte alignment padding
g_hwndMain        DCQ 0
g_hwndEdit        DCQ 0
g_hwndBtn         DCQ 0
g_hwndStatus      DCQ 0
g_hConsoleOut     DCQ 0
g_useNewConsole   DCD 0, 0
g_hInstance       DCQ 0
g_relayHandle     DCQ 0
g_relayErrHandle  DCQ 0
g_childExitCode   DCD 1, 0                  ; Conservative default exit code
g_historyEnabled  DCD 0, 0
g_hMenuFile       DCQ 0

; ==============================================================================
; UNINITIALIZED DATA SECTION (BSS)
; ==============================================================================
    AREA |.bss|, DATA, READWRITE, NOINIT, ALIGN=3

    EXPORT g_cmdBuf
    EXPORT g_statusBuf
    EXPORT g_filePath
    EXPORT g_argsBuf
    EXPORT g_tempBuf
    EXPORT g_exePath
    EXPORT g_decryptBuf
    EXPORT g_relayPath
    EXPORT g_relayErrPath
    EXPORT g_tempDirBuf
    EXPORT g_relayArgs
    EXPORT g_relayReadBuf

g_cmdBuf          SPACE 65536          ; 64 KB buffer for full Win32 command-line
g_statusBuf       SPACE 1040
g_filePath        SPACE 1040
g_argsBuf         SPACE 1040
g_tempBuf         SPACE 2080
g_exePath         SPACE 520
g_decryptBuf      SPACE 1040
g_tempDirBuf      SPACE 520
g_relayPath       SPACE 520
g_relayErrPath    SPACE 520
g_relayArgs       SPACE 131072          ; Cmdline plus two quoted relay paths
g_relayReadBuf    SPACE 4096           ; Scratch buffer for streaming temp-file bytes

; ==============================================================================
; CODE SECTION
; ==============================================================================
    AREA |.text|, CODE, READONLY, ALIGN=4

; ------------------------------------------------------------------------------
; HidePrivateConsoleForGui
; Hides a console created solely for a no-argument GUI process.
; A console shared with cmd.exe or another parent is never touched.
; ------------------------------------------------------------------------------
HidePrivateConsoleForGui PROC
    STP x29, x30, [sp, #-16]!       ; Standard ARM64 prologue
    MOV x29, sp
    SUB sp, sp, #32                 ; Allocate 32 bytes (shadow space + local)

    ; GetConsoleProcessList(&buffer, 1) -- the RETURN VALUE (w0) is the number
    ; of processes attached to the console, not the buffer contents. The buffer
    ; only receives the PID list (here just the first PID). x64 tested eax; the
    ; earlier port wrongly reloaded the buffer, comparing a PID against 1.
    ADD x0, sp, #16                 ; x0 = &buffer (one PID slot)
    MOV x1, #1                      ; x1 = 1 (buffer capacity)
    BL GetConsoleProcessList

    CMP w0, #1
    B.NE hpcfg_done                 ; more than 1 process -> shared console, keep it

    BL GetConsoleWindow
    CBZ x0, hpcfg_done              ; If no console window exists, skip

    MOV x1, #0                      ; SW_HIDE = 0
    BL ShowWindow

hpcfg_done
    ADD sp, sp, #32
    LDP x29, x30, [sp], #16         ; Standard ARM64 epilogue
    RET
HidePrivateConsoleForGui ENDP

; ==============================================================================
; mainCRTStartup - Application Entry Point
; ==============================================================================
    EXPORT mainCRTStartup
mainCRTStartup PROC
    ; Prologue: Save frame pointer, link register, and callee-saved registers
    STP x29, x30, [sp, #-16]!
    MOV x29, sp
    SUB sp, sp, #336                ; 320 bytes locals + 16 for 16-byte stack alignment
    STP x19, x20, [sp, #16]
    STP x21, x22, [sp, #32]
    STP x23, x24, [sp, #48]
    STP x25, x26, [sp, #64]
    STP x27, x28, [sp, #80]

    ; Register allocation map for main logic:
    ; x19 = argc, x20 = argv, x21 = cmdline, x22 = is_admin, x23 = argv[1]

    ; ===== Attach to parent shell's console (best effort) =====
    MOVN x0, #10                     ; STD_OUTPUT_HANDLE = -11 (0xFFFFFFF5)
    BL GetStdHandle
    CBZ x0, early_do_attach
    CMP x0, #0xFFFFFFFFFFFFFFFF
    B.EQ early_do_attach

    MOV x19, x0                     ; Temporarily save handle
    BL GetFileType
    CMP w0, #2                      ; FILE_TYPE_CHAR = 2 (console/character device)
    B.EQ early_do_attach
    B early_attach_done

early_do_attach
    MOVN x0, #0                      ; ATTACH_PARENT_PROCESS = -1
    BL AttachConsole

early_attach_done
    ; ===== Early help-switch check (runs BEFORE UAC self-elevation) =====
    BL GetCommandLineW
    MOV x21, x0                     ; x21 = cmdline

    MOV x0, x21                     ; cmdline
    ADD x1, sp, #96                 ; x1 = &argc (stored at sp+96)
    BL CommandLineToArgvW
    MOV x20, x0                     ; x20 = argv
    CBZ x20, early_help_skip

    ; Hide the non-elevated parent's private GUI console before UAC relaunch
    LDR w1, [sp, #96]
    CMP w1, #1
    B.NE early_not_gui
    BL HidePrivateConsoleForGui

early_not_gui
    ; HelpCheckAndExit(argc, argv) - never returns if help switch is found
    LDR w0, [sp, #96]               ; argc
    MOV x1, x20                     ; argv
    BL HelpCheckAndExit

early_help_skip
    ; ===== Admin check =====
    BL IsUserAnAdmin
    MOV w22, w0                     ; w22 = is_admin (writing Wn zero-extends x22)

    ; ===== CLI output relay =====
    CBZ x20, cli_relay_done
    LDR w1, [sp, #96]
    CMP w1, #2
    B.LT cli_relay_done

    LDR x23, [x20, #8]              ; x23 = argv[1]
    ADRP x1, str_cliSwitch1
    ADD x1, x1, str_cliSwitch1
    MOV x0, x23
    BL wcscmp_ci

    ; wcscmp_ci returns 1 on match, 0 on mismatch. The x64 original was
    ; 'test rax,rax / jz cli_relay_done' -- jz branches when rax==0, i.e. when
    ; argv[1] does NOT equal "-cli". The direct ARM64 equivalent is CBZ.
    CBZ x0, cli_relay_done          ; If argv[1] != "-cli", skip relay

    ; Dispatch to appropriate relay function based on elevation state
    CMP x22, #0
    B.EQ cli_relay_nonadmin

    ; AdminRelayLaunch(argc, argv, cmdline)
    LDR w0, [sp, #96]
    MOV x1, x20
    MOV x2, x21
    BL AdminRelayLaunch
    B cli_relay_ret

cli_relay_nonadmin
    ; NonAdminRelayLaunch(argc, argv, cmdline)
    LDR w0, [sp, #96]
    MOV x1, x20
    MOV x2, x21
    BL NonAdminRelayLaunch

cli_relay_ret
    ; Relay functions may exit internally, but if they return, we continue

cli_relay_done
    CMP x22, #0
    B.NE admin_dispatch

    ; ===== Non-admin: decide between relay and plain UAC =====
    CBZ x20, nonadmin_plain
    LDR w1, [sp, #96]
    CMP w1, #2
    B.LT nonadmin_plain

    LDR x23, [x20, #8]              ; argv[1]
    ADRP x1, str_cliSwitch1
    ADD x1, x1, str_cliSwitch1
    MOV x0, x23
    BL wcscmp_ci
    CBZ x0, nonadmin_plain          ; If argv[1] != "-cli", plain UAC (x64: jz)

    ; NonAdminRelayLaunch(argc, argv, cmdline)
    LDR w0, [sp, #96]
    MOV x1, x20
    MOV x2, x21
    BL NonAdminRelayLaunch
    ; If it returns 0 (declined), fall through to nonadmin_plain

nonadmin_plain
    ; Free early-parse argv
    CBZ x20, nonadmin_plain_setup
    MOV x0, x20
    BL LocalFree
    MOV x20, XZR

nonadmin_plain_setup
    ; ===== Plain UAC self-elevate =====
    MOV x0, XZR
    ADRP x1, g_exePath
    ADD x1, x1, g_exePath
    MOV w2, #260
    BL GetModuleFileNameW

    BL GetCommandLineW
    MOV x21, x0                     ; x21 = cmdline (reuse)
    MOV x24, x0                     ; x24 = scan pointer (rsi equivalent)
    MOV x25, XZR                    ; x25 = quote flag (edi equivalent)

uac_skip_exe
    LDRH w1, [x24]
    CBZ w1, uac_no_args
    CMP w1, #'"'
    B.NE uac_check_space
    EOR x25, x25, #1                ; Toggle quote flag

uac_check_space
    CMP w1, #' '
    B.NE uac_next_char
    CBNZ x25, uac_next_char         ; If in quotes, don't skip space

    ; skip_spaces
    MOV x0, x24
    BL skip_spaces
    MOV x26, x0                     ; x26 = args pointer (r15 equivalent)
    B uac_launch

uac_next_char
    ADD x24, x24, #2
    B uac_skip_exe

uac_no_args
    MOV x26, x24                    ; Points to null terminator

uac_launch
    ; Zero SHELLEXECUTEINFOW at [sp, #112] (112 bytes = 14 * 8)
    ADD x0, sp, #112
    MOV x1, XZR
    MOV x2, #14
uac_zero
    STR x1, [x0], #8
    SUBS x2, x2, #1
    B.NE uac_zero

    ; Fill SHELLEXECUTEINFOW fields
    MOV w0, #112
    STR w0, [sp, #112]              ; cbSize (offset 0)

    ADRP x0, str_runas
    ADD x0, x0, str_runas
    STR x0, [sp, #128]              ; lpVerb (offset 16)

    ADRP x0, g_exePath
    ADD x0, x0, g_exePath
    STR x0, [sp, #136]              ; lpFile (offset 24)

    STR x26, [sp, #144]             ; lpParameters (offset 32)

    MOV w0, #1                      ; SW_SHOWNORMAL
    STR w0, [sp, #160]              ; nShow (offset 48)

    ADD x0, sp, #112
    BL ShellExecuteExW

    ; Exit - elevated instance takes over, or user cancelled UAC
    MOV x0, XZR
    BL ExitProcess

admin_dispatch
    ; Free early-parse argv (will re-parse to match original behavior)
    CBZ x20, uac_already_admin
    MOV x0, x20
    BL LocalFree
    MOV x20, XZR

uac_already_admin
    BL GetCommandLineW
    MOV x21, x0

    MOV x0, x21
    ADD x1, sp, #96                 ; &argc
    BL CommandLineToArgvW
    MOV x20, x0                     ; argv
    CBZ x20, mode_gui_free

    LDR w1, [sp, #96]
    CMP w1, #2
    B.LT mode_gui_free

    LDR x23, [x20, #8]              ; argv[1]

    ; Check "-cli" (x64: jnz mode_cli_found -> branch on match). mode_cli_found
    ; lives in cli.asm; a compare-branch can't carry a cross-module relocation,
    ; so invert to a local skip and reach it with an unconditional B.
    ADRP x1, str_cliSwitch1
    ADD x1, x1, str_cliSwitch1
    MOV x0, x23
    BL wcscmp_ci
    CBZ x0, dispatch_chk_install     ; not "-cli" -> next check
    B mode_cli_found                 ; "-cli" -> tail-branch into cli.asm

dispatch_chk_install
    ; Check "-install"
    ADRP x1, str_installSwitch
    ADD x1, x1, str_installSwitch
    MOV x0, x23
    BL wcscmp_ci
    CBNZ x0, mode_install_found

    ; Check "-uninstall"
    ADRP x1, str_uninstallSwitch
    ADD x1, x1, str_uninstallSwitch
    MOV x0, x23
    BL wcscmp_ci
    CBNZ x0, mode_uninstall_found

    ; Check "-shift"
    ADRP x1, str_shiftSwitch
    ADD x1, x1, str_shiftSwitch
    MOV x0, x23
    BL wcscmp_ci
    CBNZ x0, mode_shift_found

    ; Check "-unshift"
    ADRP x1, str_unshiftSwitch
    ADD x1, x1, str_unshiftSwitch
    MOV x0, x23
    BL wcscmp_ci
    CBNZ x0, mode_unshift_found

    ; Check "-history-clear"
    ADRP x1, str_historyClearSwitch
    ADD x1, x1, str_historyClearSwitch
    MOV x0, x23
    BL wcscmp_ci
    CBNZ x0, mode_historyclear_found

    ; Unknown switch starting with '-' or '/' shows usage
    LDRH w1, [x23]
    CMP w1, #'-'
    B.EQ dispatch_show_usage
    CMP w1, #'/'
    B.EQ dispatch_show_usage
    B mode_file_run

dispatch_show_usage
    MOV x0, x20
    BL ShowUsage                    ; Never returns

mode_install_found
    MOV x0, x20
    BL LocalFree
    BL InstallContextMenu
    MOV x0, XZR
    BL ExitProcess

mode_uninstall_found
    MOV x0, x20
    BL LocalFree
    BL UninstallContextMenu
    MOV x0, XZR
    BL ExitProcess

mode_shift_found
    MOV x0, x20
    BL LocalFree
    BL InstallShift
    MOV x0, XZR
    BL ExitProcess

mode_unshift_found
    MOV x0, x20
    BL LocalFree
    BL UninstallShift
    MOV x0, XZR
    BL ExitProcess

mode_historyclear_found
    MOV x0, x20
    BL LocalFree

    ADRP x1, str_regKey
    ADD x1, x1, str_regKey
    MOVZ x0, #0x0001
    MOVK x0, #0x8000, LSL #16       ; HKEY_CURRENT_USER = 0x80000001
    BL RegDeleteTreeW

    MOV x0, XZR
    BL ExitProcess

mode_gui_free
    MOV x0, x20
    BL LocalFree

    BL HidePrivateConsoleForGui
    B mode_gui                      ; Transfer into GUI proc (never returns)

    ; Epilogue (unreachable, kept for unwind-table correctness)
    LDP x27, x28, [sp, #80]
    LDP x25, x26, [sp, #64]
    LDP x23, x24, [sp, #48]
    LDP x21, x22, [sp, #32]
    LDP x19, x20, [sp, #16]
    ADD sp, sp, #336
    LDP x29, x30, [sp], #16
    RET
mainCRTStartup ENDP

; ==============================================================================
; mode_gui - GUI Mode entry point
; ==============================================================================
    EXPORT mode_gui
mode_gui PROC
    STP x29, x30, [sp, #-16]!
    MOV x29, sp
    SUB sp, sp, #224                ; 208 bytes locals + 16 for alignment

    ; GetModuleHandleW(NULL)
    MOV x0, XZR
    BL GetModuleHandleW
    ADRP x1, g_hInstance
    ADD x1, x1, g_hInstance
    STR x0, [x1]

    ; CreateMainWindow(hInstance)
    BL CreateMainWindow
    CBZ x0, gui_exit                ; Window creation failed
    ADRP x1, g_hwndMain
    ADD x1, x1, g_hwndMain
    STR x0, [x1]

gui_msg_loop
    ; GetMessageW(&msg, NULL, 0, 0)
    ; MSG structure (48 bytes) is located at sp+16
    ADD x0, sp, #16
    MOV x1, XZR
    MOV x2, XZR
    MOV x3, XZR
    BL GetMessageW
    CBZ w0, gui_exit                ; WM_QUIT received

    ; Check for ESC key to exit application
    LDR w1, [sp, #24]               ; msg.message (offset 8 from sp+16)
    CMP w1, #0x0100                 ; WM_KEYDOWN
    B.NE gui_msg_not_esc
    LDR w1, [sp, #32]               ; msg.wParam (offset 16 from sp+16)
    CMP w1, #0x1B                   ; VK_ESCAPE
    B.EQ gui_exit

gui_msg_not_esc
    ; IsDialogMessageW(g_hwndMain, &msg)
    ADRP x0, g_hwndMain
    ADD x0, x0, g_hwndMain
    LDR x0, [x0]
    ADD x1, sp, #16
    BL IsDialogMessageW
    CBNZ w0, gui_msg_loop           ; Message was consumed

    ; TranslateMessage(&msg)
    ADD x0, sp, #16
    BL TranslateMessage

    ; DispatchMessageW(&msg)
    ADD x0, sp, #16
    BL DispatchMessageW
    B gui_msg_loop

gui_exit
    MOV x0, XZR
    BL ExitProcess

    ; Epilogue (unreachable)
    ADD sp, sp, #224
    LDP x29, x30, [sp], #16
    RET
mode_gui ENDP

    END