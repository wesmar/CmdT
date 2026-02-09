; ==============================================================================
; CMDT - Run as TrustedInstaller
; Constants and Structure Definitions (32-bit version)
; 
; Author: Marek Wesołowski (wesmar)
; Purpose: Defines Windows API constants, flags, and structure definitions
;          used throughout the 32-bit version of the application. These
;          constants match Windows SDK definitions and are used for API calls,
;          window styles, and system interactions.
; ==============================================================================

; ==============================================================================
; TOKEN ACCESS RIGHTS
; Constants for specifying desired access when opening security tokens
; ==============================================================================
TOKEN_QUERY                 EQU 0008h    ; Query token information
TOKEN_ADJUST_PRIVILEGES     EQU 0020h    ; Adjust token privileges
TOKEN_DUPLICATE             EQU 0002h    ; Duplicate token
TOKEN_QUERY_ADJUST          EQU 0028h    ; Query + Adjust combined (0x08 | 0x20)
TOKEN_DUP_QUERY             EQU 000Ah    ; Duplicate + Query (0x02 | 0x08)
TOKEN_DUP_QUERY_ADJ         EQU 002Ah    ; Duplicate + Query + Adjust (0x02 | 0x08 | 0x20)
MAXIMUM_ALLOWED             EQU 02000000h ; Request maximum access rights

; ==============================================================================
; TOKEN AND SECURITY CONSTANTS
; ==============================================================================
SE_PRIVILEGE_ENABLED        EQU 00000002h ; Enable a privilege in token
SECURITY_IMPERSONATION_LVL  EQU 2         ; SecurityImpersonation level
TOKEN_TYPE_IMPERSONATION    EQU 2         ; Impersonation token type

; ==============================================================================
; PROCESS ACCESS RIGHTS
; Constants for specifying desired access when opening processes
; ==============================================================================
PROCESS_QUERY_INFORMATION   EQU 0400h    ; Query process information
PROCESS_DUP_HANDLE          EQU 0040h    ; Duplicate handles from process
PROCESS_QUERY_DUP           EQU 0440h    ; Query + Duplicate (0x0400 | 0x0040)

; ==============================================================================
; PROCESS CREATION FLAGS
; Flags for CreateProcess family of functions
; ==============================================================================
CREATE_NO_WINDOW            EQU 08000000h ; Create process without console window
CREATE_NEW_CONSOLE          EQU 00000010h ; Create new console window for process
CREATE_UNICODE_ENVIRONMENT  EQU 00000400h ; Environment block is Unicode

; ==============================================================================
; TOOLHELP32 CONSTANTS
; Used for process enumeration via Toolhelp32 snapshot
; ==============================================================================
TH32CS_SNAPPROCESS          EQU 00000002h ; Include all processes in snapshot
PROCESSENTRY32W_SIZE        EQU 556       ; Size of PROCESSENTRY32W structure

; ==============================================================================
; SERVICE CONTROL MANAGER CONSTANTS
; For interacting with Windows services
; ==============================================================================
SERVICE_STATUS_PROCESS_SIZE EQU 36        ; Size of SERVICE_STATUS_PROCESS
SERVICE_RUNNING             EQU 4         ; Service is running state
SERVICE_STOPPED             EQU 1         ; Service is stopped state
SC_MANAGER_CONNECT          EQU 0001h    ; Connect to SCM access right
SERVICE_QUERY_STATUS        EQU 0004h    ; Query service status access right
SERVICE_START               EQU 0010h    ; Start service access right
SERVICE_QS                  EQU 0014h    ; Query + Start (0x04 | 0x10)
SC_STATUS_PROCESS_INFO      EQU 0         ; Query process information level

; ==============================================================================
; STARTUPINFO CONSTANTS
; For configuring process startup parameters
; ==============================================================================
STARTUPINFOW_SIZE           EQU 68        ; Size of STARTUPINFOW structure (32-bit)
STARTF_USESHOWWINDOW        EQU 00000001h ; Use wShowWindow field
SW_HIDE                     EQU 0         ; Hide window on startup

; ==============================================================================
; WINDOW STYLES
; Basic window style flags
; ==============================================================================
WS_OVERLAPPEDWINDOW         EQU 00CF0000h ; Standard overlapped window
WS_VISIBLE                  EQU 10000000h ; Window is initially visible
WS_CHILD                    EQU 40000000h ; Child window
WS_TABSTOP                  EQU 00010000h ; Tab stop control
WS_BORDER                   EQU 00800000h ; Window has border

; ==============================================================================
; EXTENDED WINDOW STYLES
; Extended style flags for CreateWindowEx
; ==============================================================================
WS_EX_CLIENTEDGE            EQU 00000200h ; 3D sunken edge border
WS_EX_ACCEPTFILES           EQU 00000010h ; Accept drag-drop files

; ==============================================================================
; EDIT CONTROL STYLES
; Style flags specific to edit controls
; ==============================================================================
ES_LEFT                     EQU 00000000h ; Left-aligned text
ES_AUTOHSCROLL              EQU 00000080h ; Automatic horizontal scrolling

; ==============================================================================
; COMBO BOX STYLES
; Style flags for combo box controls
; ==============================================================================
CBS_DROPDOWN                EQU 00000002h ; Dropdown combo box (editable)
CBS_AUTOHSCROLL             EQU 00000040h ; Automatic horizontal scrolling

; ==============================================================================
; BUTTON STYLES
; Style flags for button controls
; ==============================================================================
BS_PUSHBUTTON               EQU 00000000h ; Standard push button
BS_DEFPUSHBUTTON            EQU 00000001h ; Default push button (bold outline)

; ==============================================================================
; STATIC CONTROL STYLES
; Style flags for static text controls
; ==============================================================================
SS_LEFT                     EQU 00000000h ; Left-aligned text

; ==============================================================================
; COMBINED CONTROL STYLES (for convenience)
; Pre-combined style flags for common control configurations
; ==============================================================================
WS_CHILD_VISIBLE            EQU 50000000h ; WS_CHILD | WS_VISIBLE
STY_EDIT                    EQU 50810080h ; Edit control standard style
STY_COMBO                   EQU 50010042h ; ComboBox dropdown standard style
STY_BUTTON                  EQU 50010000h ; Button standard style
STY_BUTTON_DEF              EQU 50010001h ; Default button standard style
STY_STATIC                  EQU 50000000h ; Static text standard style
STY_MAINWIN                 EQU 10CF0000h ; Main window standard style

; ==============================================================================
; WINDOW CLASS STYLES
; Style flags for window class registration
; ==============================================================================
CS_HREDRAW                  EQU 0002h    ; Redraw on horizontal size change
CS_VREDRAW                  EQU 0001h    ; Redraw on vertical size change

; ==============================================================================
; WINDOW MESSAGES
; Standard Windows message identifiers
; ==============================================================================
WM_CREATE                   EQU 0001h    ; Window creation message
WM_DESTROY                  EQU 0002h    ; Window destruction message
WM_SIZE                     EQU 0005h    ; Window size changed
WM_CHAR                     EQU 0102h    ; Character input message
WM_COMMAND                  EQU 0111h    ; Command from menu or control
WM_CLOSE                    EQU 0010h    ; Window close request
WM_KEYDOWN                  EQU 0100h    ; Key pressed message
WM_GETTEXT                  EQU 000Dh    ; Get window text
WM_GETTEXTLENGTH            EQU 000Eh    ; Get window text length
WM_SETTEXT                  EQU 000Ch    ; Set window text
WM_DROPFILES                EQU 0233h    ; Files dropped on window
WM_COPYGLOBALDATA           EQU 0049h    ; Copy global data (for drag-drop)

; ==============================================================================
; MESSAGE BOX FLAGS
; Flags for MessageBox function
; ==============================================================================
MB_OK                       EQU 00000000h ; OK button only
MB_ICONERROR                EQU 00000010h ; Error icon (X)
MB_ICONINFORMATION          EQU 00000040h ; Information icon (i)

; ==============================================================================
; EDIT CONTROL NOTIFICATIONS
; Notification codes sent by edit controls via WM_COMMAND
; ==============================================================================
EN_SETFOCUS                 EQU 0100h    ; Edit control received focus

; ==============================================================================
; BUTTON MESSAGES
; Messages sent to button controls
; ==============================================================================
BM_SETSTYLE                 EQU 00F4h    ; Set button style

; ==============================================================================
; COMBO BOX MESSAGES
; Messages for interacting with combo box controls
; ==============================================================================
CB_ADDSTRING                EQU 0143h    ; Add string to list
CB_DELETESTRING             EQU 0144h    ; Delete string from list
CB_GETCOUNT                 EQU 0146h    ; Get item count
CB_GETCURSEL                EQU 0147h    ; Get current selection index
CB_GETLBTEXT                EQU 0148h    ; Get list box text
CB_INSERTSTRING             EQU 014Ah    ; Insert string at specific position
CB_RESETCONTENT             EQU 014Bh    ; Clear all items
CB_SETCURSEL                EQU 014Eh    ; Set current selection
CB_FINDSTRINGEXACT          EQU 0158h    ; Find exact string match

; ==============================================================================
; COMBO BOX NOTIFICATIONS
; Notification codes sent by combo boxes via WM_COMMAND
; ==============================================================================
CBN_SELCHANGE               EQU 1        ; Selection changed
CBN_EDITCHANGE              EQU 5        ; Edit text changed
CBN_SETFOCUS                EQU 3        ; Combo box received focus

; ==============================================================================
; MRU (MOST RECENTLY USED) CONFIGURATION
; ==============================================================================
MRU_MAX_ITEMS               EQU 5        ; Maximum number of MRU entries

; ==============================================================================
; CONTROL IDs
; Resource identifiers for controls
; ==============================================================================
IDC_COMBO                   EQU 101      ; ComboBox control ID
IDC_EDIT                    EQU 101      ; Edit control ID (same as combo)
IDC_BTN_RUN                 EQU 102      ; Run button ID
IDC_STATIC_STATUS           EQU 103      ; Status label ID
IDC_BTN_BROWSE              EQU 104      ; Browse button ID

; ==============================================================================
; MENU COMMAND IDs
; Resource identifiers for menu items
; ==============================================================================
ID_FILE_BROWSE              EQU 201      ; File -> Browse menu command
ID_FILE_CONSOLE             EQU 202      ; File -> Console menu command
ID_FILE_EXIT                EQU 203      ; File -> Exit menu command
ID_FILE_ABOUT               EQU 204      ; File -> About menu command

; ==============================================================================
; STANDARD RESOURCES
; Predefined resource identifiers
; ==============================================================================
IDC_ARROW_ATOM              EQU 32512    ; Standard arrow cursor (MAKEINTRESOURCE)
WHITE_BRUSH                 EQU 0        ; White background brush
SW_SHOWNORMAL               EQU 1        ; Show window in normal state

; ==============================================================================
; STANDARD HANDLES
; Console standard handle identifiers
; ==============================================================================
STD_INPUT_HANDLE            EQU -10      ; Standard input (stdin)
STD_OUTPUT_HANDLE           EQU -11      ; Standard output (stdout)
STD_ERROR_HANDLE            EQU -12      ; Standard error (stderr)
STARTF_USESTDHANDLES        EQU 00000100h ; Use hStd* fields in STARTUPINFO

; ==============================================================================
; MENU FLAGS
; Flags for AppendMenu and related functions
; ==============================================================================
MF_STRING                   EQU 00000000h ; Menu item contains string
MF_POPUP                    EQU 00000010h ; Menu item opens popup menu
MF_SEPARATOR                EQU 00000800h ; Menu separator line

; ==============================================================================
; OPEN FILE DIALOG FLAGS
; Flags for GetOpenFileName function
; ==============================================================================
OFN_FILEMUSTEXIST           EQU 00001000h ; File must already exist
OFN_PATHMUSTEXIST           EQU 00000800h ; Path must already exist

; ==============================================================================
; COM CONSTANTS
; Component Object Model constants
; ==============================================================================
CLSCTX_INPROC_SERVER        EQU 1        ; In-process server class context
MSGFLT_ALLOW                EQU 1        ; Allow message through UIPI filter

; ==============================================================================
; VIRTUAL KEY CODES
; Keyboard virtual key identifiers
; ==============================================================================
VK_ESCAPE                   EQU 1Bh      ; ESC key code

; ==============================================================================
; CONSOLE ATTACH FLAGS
; ==============================================================================
ATTACH_PARENT_PROCESS       EQU 0FFFFFFFFh ; Attach to parent process console

; ==============================================================================
; REGISTRY CONSTANTS
; Windows Registry access flags and types
; ==============================================================================
HKEY_CURRENT_USER           EQU 80000001h ; Current user registry hive
HKEY_CLASSES_ROOT           EQU 80000000h ; Classes root registry hive
HKEY_LOCAL_MACHINE          EQU 80000002h ; Local machine registry hive
KEY_READ                    EQU 00020019h ; Read access to registry key
KEY_WRITE                   EQU 00020006h ; Write access to registry key
REG_SZ                      EQU 00000001h ; Null-terminated string value type

; ==============================================================================
; MESSAGE STRUCTURE
; Definition of Windows MSG structure for message loop
; ==============================================================================
MSG STRUCT
    hwnd    DWORD ?         ; Window handle that message is for
    message DWORD ?         ; Message identifier (WM_*)
    wParam  DWORD ?         ; First message parameter
    lParam  DWORD ?         ; Second message parameter
    time    DWORD ?         ; Time message was posted
    pt_x    DWORD ?         ; Cursor X position when message posted
    pt_y    DWORD ?         ; Cursor Y position when message posted
MSG ENDS
