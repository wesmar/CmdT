; ==============================================================================
; CMDT - Run as TrustedInstaller
; Main Entry Point and Command-Line Interface Module
; 
; Author: Marek Wesołowski (wesmar)
; Purpose: Implements the main entry point for the application with support for
;          both GUI and CLI modes. Handles command-line argument parsing,
;          .lnk shortcut resolution, and process creation with TrustedInstaller
;          privileges.
;
; Features:
;          - Dual mode operation: GUI (default) or CLI (with -cli flag)
;          - Command-line argument parsing with multiple switch formats
;          - Optional new console window creation (-new flag)
;          - Windows shortcut (.lnk) file resolution
;          - Automatic command and argument extraction from shortcuts
;          - Integration with TrustedInstaller token acquisition
; ==============================================================================

.586                            ; Target 80586 instruction set
.model flat, stdcall            ; 32-bit flat memory model, stdcall convention
option casemap:none             ; Case-sensitive symbol names

include consts.inc              ; Windows API constants and structures

; ==============================================================================
; EXTERNAL FUNCTION PROTOTYPES
; ==============================================================================

; Window and GUI functions
CreateMainWindow        PROTO :DWORD

; Process and token management
RunAsTrustedInstaller   PROTO :DWORD,:DWORD

; Shortcut (.lnk) file resolution
ResolveLnkPath          PROTO :DWORD,:DWORD,:DWORD

; Wide string manipulation utilities (defined in strutil.asm)
wcscpy_p                PROTO :DWORD,:DWORD
wcscat_p                PROTO :DWORD,:DWORD
wcscmp_ci               PROTO :DWORD,:DWORD
wcscmp_token            PROTO :DWORD,:DWORD
skip_spaces             PROTO :DWORD
wcslen_p                PROTO :DWORD

; Help / usage display (defined in help.asm)
IsHelpSwitch            PROTO :DWORD
ShowUsageAndExit        PROTO :DWORD

; Windows API functions
GetModuleHandleW        PROTO :DWORD
GetMessageW             PROTO :DWORD,:DWORD,:DWORD,:DWORD
TranslateMessage        PROTO :DWORD
DispatchMessageW        PROTO :DWORD
IsDialogMessageW        PROTO :DWORD,:DWORD
ExitProcess             PROTO :DWORD
GetCommandLineW         PROTO
CommandLineToArgvW      PROTO :DWORD,:DWORD
LocalFree               PROTO :DWORD
SetFocus                PROTO :DWORD
IsUserAnAdmin           PROTO
ShellExecuteExW         PROTO :DWORD
GetModuleFileNameW      PROTO :DWORD,:DWORD,:DWORD
RegCreateKeyExW         PROTO :DWORD,:DWORD,:DWORD,:DWORD,:DWORD,:DWORD,:DWORD,:DWORD,:DWORD
RegSetValueExW          PROTO :DWORD,:DWORD,:DWORD,:DWORD,:DWORD,:DWORD
RegDeleteKeyW           PROTO :DWORD,:DWORD
RegCloseKey             PROTO :DWORD
AttachConsole           PROTO :DWORD
GetStdHandle            PROTO :DWORD
GetFileType             PROTO :DWORD
WriteConsoleW           PROTO :DWORD,:DWORD,:DWORD,:DWORD,:DWORD
WriteFile               PROTO :DWORD,:DWORD,:DWORD,:DWORD,:DWORD
WriteConsoleInputW      PROTO :DWORD,:DWORD,:DWORD,:DWORD
CreateFileW             PROTO :DWORD,:DWORD,:DWORD,:DWORD,:DWORD,:DWORD,:DWORD
ReadFile                PROTO :DWORD,:DWORD,:DWORD,:DWORD,:DWORD
DeleteFileW             PROTO :DWORD
GetTempPathW            PROTO :DWORD,:DWORD
GetTempFileNameW        PROTO :DWORD,:DWORD,:DWORD,:DWORD
RegDeleteValueW         PROTO :DWORD,:DWORD
RegOpenKeyExW           PROTO :DWORD,:DWORD,:DWORD,:DWORD,:DWORD
WaitForSingleObject     PROTO :DWORD,:DWORD
CloseHandle             PROTO :DWORD

; Installation / hook management (defined in install.asm)
InstallContextMenu      PROTO
UninstallContextMenu    PROTO
InstallShift            PROTO
UninstallShift          PROTO

; Non-admin output relay (defined in relay.asm)
NonAdminRelayLaunch     PROTO :DWORD,:DWORD,:DWORD

; CLI / file-run dispatch targets (defined in cli.asm). These are reached
; via JMP rather than CALL — they share start's stack frame for sei
; and rely on the promoted globals (g_argv, g_argc, g_argv1, g_sa).
EXTRN mode_cli_found:PROC
EXTRN mode_file_run:PROC

; mode_gui lives below as its own proc; declared here so we can invoke it
; from inside `start` (the dispatch's mode_gui_free wrapper).
mode_gui                PROTO

; ==============================================================================
; CONSTANT STRING DATA
; ==============================================================================
.const

; Strings shared with relay.asm and (later) cli.asm. PUBLIC so cross-module
; references resolve at link time. Help / install switches stay private —
; they are only consulted in the dispatcher inside `start`.
PUBLIC str_runas, str_newSwitch, str_extLnk_m, str_space, str_outfileFlag

; Command-line switch variations for CLI mode
str_cliSwitch1  dw '-','c','l','i',0      ; Standard Unix-style switch
str_cliSwitch2  dw '-','-','c','l','i',0  ; GNU-style long option
str_cliSwitch3  dw 'c','l','i',0          ; Bare switch without hyphen

; Switch for new console window creation
str_newSwitch   dw '-','n','e','w',0

; Internal relay switch used by non-admin -cli parent.
str_outfileFlag dw '-','o','u','t','f','i','l','e',0

; File extension strings
str_extLnk_m    dw '.','l','n','k',0      ; Windows shortcut extension
str_space       dw ' ',0                   ; Space character for string building

; CLI fallback for regedit on WOW64-only systems

; UAC self-elevation verb
str_runas           dw 'r','u','n','a','s',0

; Context menu registration switches
str_installSwitch   dw '-','i','n','s','t','a','l','l',0
str_uninstallSwitch dw '-','u','n','i','n','s','t','a','l','l',0

; Sticky Keys (sethc.exe) IFEO switches
str_shiftSwitch     dw '-','s','h','i','f','t',0
str_unshiftSwitch   dw '-','u','n','s','h','i','f','t',0

; ==============================================================================
; PRIVILEGE STRING DEFINITIONS
; Windows privileges without "Se" prefix and "Privilege" suffix
; These are combined with privPrefix and privSuffix to form complete names
; ==============================================================================

; Privilege name parts (core names only)
privStr_0  dw 'A','s','s','i','g','n','P','r','i','m','a','r','y','T','o','k','e','n',0      ; SeAssignPrimaryTokenPrivilege
privStr_1  dw 'B','a','c','k','u','p',0                                                        ; SeBackupPrivilege
privStr_2  dw 'R','e','s','t','o','r','e',0                                                    ; SeRestorePrivilege
privStr_3  dw 'D','e','b','u','g',0                                                            ; SeDebugPrivilege
privStr_4  dw 'I','m','p','e','r','s','o','n','a','t','e',0                                    ; SeImpersonatePrivilege
privStr_5  dw 'T','a','k','e','O','w','n','e','r','s','h','i','p',0                            ; SeTakeOwnershipPrivilege
privStr_6  dw 'L','o','a','d','D','r','i','v','e','r',0                                        ; SeLoadDriverPrivilege
privStr_7  dw 'S','y','s','t','e','m','E','n','v','i','r','o','n','m','e','n','t',0            ; SeSystemEnvironmentPrivilege
privStr_8  dw 'M','a','n','a','g','e','V','o','l','u','m','e',0                                ; SeManageVolumePrivilege
privStr_9  dw 'S','e','c','u','r','i','t','y',0                                                ; SeSecurityPrivilege
privStr_10 dw 'S','h','u','t','d','o','w','n',0                                                ; SeShutdownPrivilege
privStr_11 dw 'S','y','s','t','e','m','t','i','m','e',0                                        ; SeSystemtimePrivilege
privStr_12 dw 'T','c','b',0                                                                    ; SeTcbPrivilege
privStr_13 dw 'I','n','c','r','e','a','s','e','Q','u','o','t','a',0                            ; SeIncreaseQuotaPrivilege
privStr_14 dw 'A','u','d','i','t',0                                                            ; SeAuditPrivilege
privStr_15 dw 'C','h','a','n','g','e','N','o','t','i','f','y',0                                ; SeChangeNotifyPrivilege
privStr_16 dw 'U','n','d','o','c','k',0                                                        ; SeUndockPrivilege
privStr_17 dw 'C','r','e','a','t','e','T','o','k','e','n',0                                    ; SeCreateTokenPrivilege
privStr_18 dw 'L','o','c','k','M','e','m','o','r','y',0                                        ; SeLockMemoryPrivilege
privStr_19 dw 'C','r','e','a','t','e','P','a','g','e','f','i','l','e',0                        ; SeCreatePagefilePrivilege
privStr_20 dw 'C','r','e','a','t','e','P','e','r','m','a','n','e','n','t',0                    ; SeCreatePermanentPrivilege
privStr_21 dw 'S','y','s','t','e','m','P','r','o','f','i','l','e',0                            ; SeSystemProfilePrivilege
privStr_22 dw 'P','r','o','f','i','l','e','S','i','n','g','l','e','P','r','o','c','e','s','s',0 ; SeProfileSingleProcessPrivilege
privStr_23 dw 'C','r','e','a','t','e','G','l','o','b','a','l',0                                ; SeCreateGlobalPrivilege
privStr_24 dw 'T','i','m','e','Z','o','n','e',0                                                ; SeTimeZonePrivilege
privStr_25 dw 'C','r','e','a','t','e','S','y','m','b','o','l','i','c','L','i','n','k',0        ; SeCreateSymbolicLinkPrivilege
privStr_26 dw 'I','n','c','r','e','a','s','e','B','a','s','e','P','r','i','o','r','i','t','y',0 ; SeIncreaseBasePriorityPrivilege
privStr_27 dw 'R','e','m','o','t','e','S','h','u','t','d','o','w','n',0                        ; SeRemoteShutdownPrivilege
privStr_28 dw 'I','n','c','r','e','a','s','e','W','o','r','k','i','n','g','S','e','t',0        ; SeIncreaseWorkingSetPrivilege
privStr_29 dw 'R','e','l','a','b','e','l',0                                                    ; SeRelabelPrivilege
privStr_30 dw 'D','e','l','e','g','a','t','e','S','e','s','s','i','o','n','U','s','e','r','I','m','p','e','r','s','o','n','a','t','e',0 ; SeDelegateSessionUserImpersonatePrivilege
privStr_31 dw 'T','r','u','s','t','e','d','C','r','e','d','M','a','n','A','c','c','e','s','s',0 ; SeTrustedCredManAccessPrivilege
privStr_32 dw 'E','n','a','b','l','e','D','e','l','e','g','a','t','i','o','n',0                ; SeEnableDelegationPrivilege
privStr_33 dw 'S','y','n','c','A','g','e','n','t',0                                            ; SeSyncAgentPrivilege

; Privilege name prefix and suffix for building complete privilege names
privPrefix dw 'S','e',0                    ; Standard Windows privilege prefix
privSuffix dw 'P','r','i','v','i','l','e','g','e',0 ; Standard Windows privilege suffix

; ==============================================================================
; INITIALIZED DATA SECTION
; ==============================================================================
.data
    align 4                                 ; Align to 4-byte boundary for performance

; Privilege table: array of pointers to privilege name strings
; Used for iterating through all privileges when enabling them
PUBLIC g_privTable
g_privTable dd offset privStr_0,offset privStr_1,offset privStr_2,offset privStr_3,offset privStr_4,offset privStr_5
            dd offset privStr_6,offset privStr_7,offset privStr_8,offset privStr_9,offset privStr_10,offset privStr_11
            dd offset privStr_12,offset privStr_13,offset privStr_14,offset privStr_15,offset privStr_16,offset privStr_17
            dd offset privStr_18,offset privStr_19,offset privStr_20,offset privStr_21,offset privStr_22,offset privStr_23
            dd offset privStr_24,offset privStr_25,offset privStr_26,offset privStr_27,offset privStr_28,offset privStr_29
            dd offset privStr_30,offset privStr_31,offset privStr_32,offset privStr_33

; Global variables exported for use in other modules
PUBLIC g_cachedToken, g_tokenTime, g_hwndMain, g_hwndEdit, g_hwndBtn, g_hwndStatus, g_hConsoleOut, g_hInstance
PUBLIC g_useNewConsole, g_relayHandle
PUBLIC g_argv, g_argc, g_argv1, g_sa
PUBLIC privPrefix, privSuffix
; FixRegeditPath moved to cli.asm (PUBLIC declared there)

g_cachedToken   dd 0                        ; Cached TrustedInstaller token handle
g_tokenTime     dd 0                        ; Timestamp of cached token (for expiration)
g_hwndMain      dd 0                        ; Main window handle
g_hwndEdit      dd 0                        ; Edit/ComboBox control handle
g_hwndBtn       dd 0                        ; Run button handle
g_hwndStatus    dd 0                        ; Status label handle
g_hConsoleOut   dd 0                        ; Console output handle (CLI mode)
g_useNewConsole dd 0                        ; Flag: create new console window
g_hInstance     dd 0                        ; Application instance handle
g_relayHandle   dd 0                        ; Output relay file handle for elevated child

; Command-line parse outputs from `start` — promoted from LOCALs to globals
; so cli.asm can refer to them after the dispatch labels were extracted.
g_argv          dd 0                        ; Argument vector (LocalFree-able)
g_argc          dd 0                        ; Argument count
g_argv1         dd 0                        ; Pointer to argv[1] (or NULL)

; SECURITY_ATTRIBUTES used by mode_cli_setup when it opens the relay file
; with bInheritHandle=TRUE. Three DWORDs: nLength, lpSecurityDescriptor,
; bInheritHandle.
g_sa            dd 3 dup(0)

; ==============================================================================
; UNINITIALIZED DATA SECTION
; Large buffers for strings and temporary data
; ==============================================================================
.data?
PUBLIC g_cmdBuf, g_statusBuf, g_filePath, g_argsBuf, g_tempBuf
PUBLIC g_exePath, g_tempDirBuf, g_relayPath, g_relayArgs, g_relayReadBuf

g_cmdBuf        dw 520 dup(?)               ; Command line buffer (1040 bytes)
g_statusBuf     dw 520 dup(?)               ; Status message buffer (1040 bytes)
g_filePath      dw 520 dup(?)               ; File path buffer (1040 bytes)
g_argsBuf       dw 520 dup(?)               ; Arguments buffer (1040 bytes)
g_tempBuf       dw 1040 dup(?)              ; Temporary work buffer (2080 bytes)
g_exePath       dw 260 dup(?)               ; Exe path buffer (UAC and context menu)
g_tempDirBuf    dw 260 dup(?)               ; Relay temp directory buffer
g_relayPath     dw 260 dup(?)               ; Relay temp file path
g_relayArgs     dw 1040 dup(?)              ; Relay child argument string
g_relayReadBuf  db 4096 dup(?)              ; Relay read/copy buffer

; ==============================================================================
; CODE SECTION
; ==============================================================================
.code


; ==============================================================================
; FixRegeditPath - WOW64-safe regedit fallback
;
; Purpose: On some systems, System32\regedit.exe is missing and only the
;          SysWOW64 version exists. If the command token is "regedit" or
;          "regedit.exe" without a path, rewrite the command to use:
;          C:\Windows\SysWOW64\regedit.exe
;
; Parameters:
;   lpCmd - Pointer to command line string
;
; Returns:
;   EAX = Pointer to command string to execute (original or g_cmdBuf)
;
; Registers modified: EAX, EBX, ECX, EDX, ESI, EDI
; ==============================================================================


; ==============================================================================
; start - Main Entry Point
;
; Purpose: Application entry point. Parses command-line arguments to determine
;          operating mode (GUI vs CLI), processes .lnk shortcuts if necessary,
;          and either launches the GUI or executes commands in CLI mode.
;
; Command-line usage:
;   GUI mode (default):
;     cmdt.exe
;
;   CLI mode:
;     cmdt.exe -cli <command>           - Run command, inherit console
;     cmdt.exe -cli -new <command>      - Run command in new console window
;     cmdt.exe --cli <command>          - Alternative CLI switch format
;     cmdt.exe cli <command>            - Short form CLI switch
;
; Process flow:
;   1. Parse command line using CommandLineToArgvW
;   2. Check for CLI switches (-cli, --cli, cli)
;   3. Check for -new flag (new console window)
;   4. Extract command portion from arguments
;   5. Check if command is a .lnk file
;   6. If .lnk: resolve to target executable and arguments
;   7. Execute via RunAsTrustedInstaller or launch GUI
;   8. Clean up and exit
;
; Local variables:
;   g_argv   - Pointer to argument vector array
;   argc    - Argument count
;   msg     - Windows message structure (for GUI mode)
;   g_argv1   - Pointer to first argument (for switch checking)
;
; Returns: Does not return (calls ExitProcess)
; ==============================================================================
start proc
    LOCAL sei[60]:BYTE                      ; SHELLEXECUTEINFOW for UAC elevation
    ; pArgv/argc/argv1/sa promoted to globals (g_argv, g_argc, g_argv1, g_sa)
    ; so that cli.asm dispatch labels can reach them after extraction.
    ; msg (MSG struct) lives in mode_gui's own frame now that it's a proc.

    cld                                     ; Clear direction flag (forward string ops)

    ; Parse once before UAC so help can print without elevation and non-admin
    ; CLI can decide whether to use the output relay.
    invoke GetCommandLineW
    mov ebx, eax                            ; EBX = raw command line
    invoke CommandLineToArgvW, ebx, offset g_argc
    mov g_argv, eax
    test eax, eax
    jz early_after_help

    cmp g_argc,2
    jl early_after_help
    mov esi, g_argv
    mov eax, [esi+4]
    mov g_argv1, eax
    invoke IsHelpSwitch, g_argv1
    test eax, eax
    jz early_after_help
    invoke ShowUsageAndExit, g_argv

early_after_help:
    ; Check if running as administrator
    invoke IsUserAnAdmin
    test eax, eax
    jnz uac_already_admin

    ; Non-admin -cli path: use relay so stdout/stderr survive UAC and `>>`.
    cmp g_argv, 0
    je nonadmin_plain
    cmp g_argc,2
    jl nonadmin_plain
    mov esi, g_argv
    mov eax, [esi+4]
    mov g_argv1, eax
    invoke wcscmp_ci, g_argv1, offset str_cliSwitch1
    test eax, eax
    jnz nonadmin_try_relay
    invoke wcscmp_ci, g_argv1, offset str_cliSwitch2
    test eax, eax
    jnz nonadmin_try_relay
    invoke wcscmp_ci, g_argv1, offset str_cliSwitch3
    test eax, eax
    jz nonadmin_plain

nonadmin_try_relay:
    invoke NonAdminRelayLaunch, g_argv, g_argc, ebx

nonadmin_plain:
    cmp g_argv, 0
    je nonadmin_plain_setup
    invoke LocalFree, g_argv
    mov g_argv, 0

nonadmin_plain_setup:
    ; Not admin - relaunch with UAC elevation prompt
    invoke GetModuleFileNameW, 0, offset g_exePath, 260

    ; Extract arguments from command line (skip exe path)
    invoke GetCommandLineW
    mov esi, eax
    xor edi, edi
uac_skip_exe:
    mov ax, word ptr [esi]
    test ax, ax
    jz uac_no_args
    cmp ax, '"'
    jne @F
    xor edi, 1
@@:
    cmp ax, ' '
    jne @F
    test edi, edi
    jnz @F
    invoke skip_spaces, esi
    mov ebx, eax
    jmp uac_launch
@@:
    add esi, 2
    jmp uac_skip_exe

uac_no_args:
    mov ebx, esi                ; Points to null terminator (empty args)

uac_launch:
    ; Zero SHELLEXECUTEINFOW (60 bytes)
    lea edi, sei
    xor eax, eax
    mov ecx, 15
    rep stosd

    ; Fill SHELLEXECUTEINFOW fields
    lea edi, sei
    mov dword ptr [edi], 60                     ; cbSize
    mov dword ptr [edi+12], offset str_runas    ; lpVerb
    mov dword ptr [edi+16], offset g_exePath    ; lpFile
    mov dword ptr [edi+20], ebx                 ; lpParameters
    mov dword ptr [edi+28], SW_SHOWNORMAL       ; nShow

    invoke ShellExecuteExW, edi

    ; Exit - elevated instance takes over, or user cancelled UAC
    invoke ExitProcess, 0

uac_already_admin:
    cmp g_argv, 0
    je mode_gui

    ; Check argument count: need at least 2 for CLI mode (exe + switch)
    cmp g_argc,2
    jl mode_gui_free                        ; Less than 2 args â†’ GUI mode
    
    ; Retrieve argv[1] (first argument after executable name)
    mov esi, g_argv
    mov eax, [esi+4]                        ; argv[1] pointer
    mov g_argv1, eax
    
    ; Check if argv[1] == "-cli" (standard switch)
    push offset str_cliSwitch1
    push g_argv1
    call wcscmp_ci                          ; Case-insensitive comparison
    test eax, eax
    jnz mode_cli_found                      ; Non-zero = match found

    ; Check if argv[1] == "--cli" (GNU-style long option)
    push offset str_cliSwitch2
    push g_argv1
    call wcscmp_ci
    test eax, eax
    jnz mode_cli_found                      ; Match found

    ; Check if argv[1] == "cli" (bare switch without hyphens)
    push offset str_cliSwitch3
    push g_argv1
    call wcscmp_ci
    test eax, eax
    jnz mode_cli_found                      ; Match found

    ; Check if argv[1] matches "-install"
    invoke wcscmp_ci, g_argv1, offset str_installSwitch
    test eax, eax
    jnz mode_install_found

    ; Check if argv[1] matches "-uninstall"
    invoke wcscmp_ci, g_argv1, offset str_uninstallSwitch
    test eax, eax
    jnz mode_uninstall_found

    ; Check if argv[1] matches "-shift"
    invoke wcscmp_ci, g_argv1, offset str_shiftSwitch
    test eax, eax
    jnz mode_shift_found

    ; Check if argv[1] matches "-unshift"
    invoke wcscmp_ci, g_argv1, offset str_unshiftSwitch
    test eax, eax
    jnz mode_unshift_found

    ; No recognized switch: check if argv[1] starts with '-'
    mov eax, g_argv1
    cmp word ptr [eax], '-'
    je show_usage               ; Unknown switch, display available options
    cmp word ptr [eax], '/'
    je show_usage
    jmp mode_file_run           ; Not a switch, treat as file path


mode_install_found:
    invoke LocalFree, g_argv
    call InstallContextMenu
    invoke ExitProcess, 0

mode_uninstall_found:
    invoke LocalFree, g_argv
    call UninstallContextMenu
    invoke ExitProcess, 0

mode_shift_found:
    invoke LocalFree, g_argv
    call InstallShift
    invoke ExitProcess, 0

mode_unshift_found:
    invoke LocalFree, g_argv
    call UninstallShift
    invoke ExitProcess, 0

show_usage:
    ; Unknown switch detected, display available options and exit
    invoke ShowUsageAndExit, g_argv

mode_gui_free:
    invoke LocalFree, g_argv                 ; Free argv and go to GUI mode
    invoke mode_gui                          ; never returns (ExitProcess inside)
    ret                                      ; unreachable but keeps unwind clean
start endp

; ==============================================================================
; mode_gui - GUI Mode entry point
;
; Purpose: Creates the main application window and runs the message loop
;          until WM_QUIT (or ESC). Reached either from start's dispatch when
;          no recognized switch was supplied, or from cli.asm's
;          skip_exe_for_file when the file-run path had no file argument.
;
; Parameters: None
;
; Returns: Never (every exit path goes through ExitProcess).
; ==============================================================================
mode_gui proc
    LOCAL msg:MSG

    invoke GetModuleHandleW, 0
    mov g_hInstance, eax
    invoke CreateMainWindow, eax
    test eax, eax
    jz gui_exit

gui_msg_loop:
    invoke GetMessageW, addr msg, 0, 0, 0
    test eax, eax
    jz gui_exit                             ; WM_QUIT received

    cmp [msg.message], WM_KEYDOWN
    jne gui_msg_not_esc
    cmp [msg.wParam], VK_ESCAPE
    je gui_exit                             ; ESC exits

gui_msg_not_esc:
    invoke IsDialogMessageW, g_hwndMain, addr msg
    test eax, eax
    jnz gui_msg_loop                        ; Message consumed by IsDialogMessageW

    invoke TranslateMessage, addr msg
    invoke DispatchMessageW, addr msg
    jmp gui_msg_loop

gui_exit:
    invoke ExitProcess, 0
    ret                                     ; unreachable
mode_gui endp

; ==============================================================================
; InstallContextMenu - Register Explorer context menu entries
;
; Purpose: Creates registry keys under HKEY_CLASSES_ROOT for context menu
;          entries that allow running executables and opening directories
;          with TrustedInstaller privileges.
;
; Registry locations created:
;   - Directory\Background\shell\CMDT (background right-click in folders)
;   - Directory\shell\CMDT (right-click on folder icons)
;   - exefile\shell\CMDT (right-click on .exe files)
;   - lnkfile\shell\CMDT (right-click on .lnk shortcut files)
;
; Each entry includes:
;   - Default value: Menu text ("Open CMD as TrustedInstaller" or "Run as TrustedInstaller")
;   - Icon value: shell32.dll,104 (UAC shield icon)
;   - command subkey: Command line to execute when menu item is selected
;
; Commands generated:
;   - For directories: "<exepath>" -cli -new cmd.exe /k cd /d "%V"
;   - For files: "<exepath>" "%1"
;
; Parameters: None (uses global g_exePath and g_tempBuf buffers)
;
; Returns: None (ignores errors to allow partial installation)
; ==============================================================================

end start
