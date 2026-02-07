; ==============================================================================
; CMDT - Run as TrustedInstaller
; Constants and Structure Definitions
; 
; Author: Marek Wesołowski (wesmar)
; Purpose: Defines Windows API constants, flags, and structure definitions
;          used throughout the application.
; ==============================================================================

; ==============================================================================
; TOKEN ACCESS RIGHTS
; ==============================================================================
TOKEN_QUERY                 EQU 0008h    ; Query token information
TOKEN_ADJUST_PRIVILEGES     EQU 0020h    ; Adjust token privileges
TOKEN_DUPLICATE             EQU 0002h    ; Duplicate token
TOKEN_QUERY_ADJUST          EQU 0028h    ; Query + Adjust combined
TOKEN_DUP_QUERY             EQU 000Ah    ; Duplicate + Query combined
TOKEN_DUP_QUERY_ADJ         EQU 002Ah    ; Duplicate + Query + Adjust combined
MAXIMUM_ALLOWED             EQU 02000000h ; Maximum access rights

; ==============================================================================
; TOKEN AND SECURITY CONSTANTS
; ==============================================================================
SE_PRIVILEGE_ENABLED        EQU 00000002h ; Enable a privilege
SECURITY_IMPERSONATION_LVL  EQU 2         ; Security impersonation level
TOKEN_TYPE_IMPERSONATION    EQU 2         ; Impersonation token type

; ==============================================================================
; PROCESS ACCESS RIGHTS
; ==============================================================================
PROCESS_QUERY_INFORMATION   EQU 0400h    ; Query process information
PROCESS_DUP_HANDLE          EQU 0040h    ; Duplicate handles
PROCESS_QUERY_DUP           EQU 0440h    ; Query + Duplicate combined

; ==============================================================================
; PROCESS CREATION FLAGS
; ==============================================================================
CREATE_NO_WINDOW            EQU 08000000h ; Create process without window
CREATE_NEW_CONSOLE          EQU 00000010h ; Create new console window
CREATE_UNICODE_ENVIRONMENT  EQU 00000400h ; Use Unicode environment block

; ==============================================================================
; TOOLHELP32 CONSTANTS
; ==============================================================================
TH32CS_SNAPPROCESS          EQU 00000002h ; Include all processes
PROCESSENTRY32W_SIZE        EQU 568       ; Size of PROCESSENTRY32W structure

; ==============================================================================
; SERVICE CONTROL MANAGER CONSTANTS
; ==============================================================================
SERVICE_STATUS_PROCESS_SIZE EQU 36        ; Size of SERVICE_STATUS_PROCESS
SERVICE_RUNNING             EQU 4         ; Service is running
SERVICE_STOPPED             EQU 1         ; Service is stopped
SC_MANAGER_CONNECT          EQU 0001h    ; Connect to SCM
SERVICE_QUERY_STATUS        EQU 0004h    ; Query service status
SERVICE_START               EQU 0010h    ; Start service
SERVICE_QS                  EQU 0014h    ; Query + Start combined
SC_STATUS_PROCESS_INFO      EQU 0         ; Query process info

; ==============================================================================
; STARTUPINFO CONSTANTS
; ==============================================================================
STARTUPINFOW_SIZE           EQU 104       ; Size of STARTUPINFOW structure
STARTF_USESHOWWINDOW        EQU 00000001h ; Use wShowWindow field
SW_HIDE                     EQU 0         ; Hide window

; ==============================================================================
; WINDOW STYLES
; ==============================================================================
WS_OVERLAPPEDWINDOW         EQU 00CF0000h ; Standard overlapped window
WS_VISIBLE                  EQU 10000000h ; Window is visible
WS_CHILD                    EQU 40000000h ; Child window
WS_TABSTOP                  EQU 00010000h ; Tab stop
WS_BORDER                   EQU 00800000h ; Window has border

; ==============================================================================
; EXTENDED WINDOW STYLES
; ==============================================================================
WS_EX_CLIENTEDGE            EQU 00000200h ; 3D sunken edge
WS_EX_ACCEPTFILES           EQU 00000010h ; Accept drag-drop files

; ==============================================================================
; EDIT CONTROL STYLES
; ==============================================================================
ES_LEFT                     EQU 00000000h ; Left-aligned text
ES_AUTOHSCROLL              EQU 00000080h ; Auto horizontal scroll

; ==============================================================================
; COMBO BOX STYLES
; ==============================================================================
CBS_DROPDOWN                EQU 00000002h ; Dropdown combo box
CBS_AUTOHSCROLL             EQU 00000040h ; Auto horizontal scroll

; ==============================================================================
; BUTTON STYLES
; ==============================================================================
BS_PUSHBUTTON               EQU 00000000h ; Push button
BS_DEFPUSHBUTTON            EQU 00000001h ; Default push button

; ==============================================================================
; STATIC CONTROL STYLES
; ==============================================================================
SS_LEFT                     EQU 00000000h ; Left-aligned text

; ==============================================================================
; COMBINED CONTROL STYLES (for convenience)
; ==============================================================================
WS_CHILD_VISIBLE            EQU 50000000h ; Child + Visible
STY_EDIT                    EQU 50810080h ; Edit control style
STY_COMBO                   EQU 50010042h ; ComboBox style
STY_BUTTON                  EQU 50010000h ; Button style
STY_BUTTON_DEF              EQU 50010001h ; Default button style
STY_STATIC                  EQU 50000000h ; Static text style
STY_MAINWIN                 EQU 10CF0000h ; Main window style

; ==============================================================================
; WINDOW CLASS STYLES
; ==============================================================================
CS_HREDRAW                  EQU 0002h    ; Redraw on horizontal resize
CS_VREDRAW                  EQU 0001h    ; Redraw on vertical resize

; ==============================================================================
; WINDOW MESSAGES
; ==============================================================================
WM_CREATE                   EQU 0001h    ; Window creation
WM_DESTROY                  EQU 0002h    ; Window destruction
WM_SIZE                     EQU 0005h    ; Window size changed
WM_CHAR                     EQU 0102h    ; Character input
WM_COMMAND                  EQU 0111h    ; Command from menu or control
WM_CLOSE                    EQU 0010h    ; Window close request
WM_KEYDOWN                  EQU 0100h    ; Key pressed
WM_GETTEXT                  EQU 000Dh    ; Get window text
WM_GETTEXTLENGTH            EQU 000Eh    ; Get text length
WM_SETTEXT                  EQU 000Ch    ; Set window text
WM_DROPFILES                EQU 0233h    ; Files dropped
WM_COPYGLOBALDATA           EQU 0049h    ; Copy global data

; ==============================================================================
; MESSAGE BOX FLAGS
; ==============================================================================
MB_OK                       EQU 00000000h ; OK button only
MB_ICONERROR                EQU 00000010h ; Error icon
MB_ICONINFORMATION          EQU 00000040h ; Information icon

; ==============================================================================
; EDIT CONTROL NOTIFICATIONS
; ==============================================================================
EN_SETFOCUS                 EQU 0100h    ; Edit control got focus

; ==============================================================================
; BUTTON MESSAGES
; ==============================================================================
BM_SETSTYLE                 EQU 00F4h    ; Set button style

; ==============================================================================
; COMBO BOX MESSAGES
; ==============================================================================
CB_ADDSTRING                EQU 0143h    ; Add string
CB_DELETESTRING             EQU 0144h    ; Delete string
CB_GETCOUNT                 EQU 0146h    ; Get item count
CB_GETCURSEL                EQU 0147h    ; Get current selection
CB_GETLBTEXT                EQU 0148h    ; Get list box text
CB_INSERTSTRING             EQU 014Ah    ; Insert string at position
CB_RESETCONTENT             EQU 014Bh    ; Clear all items
CB_SETCURSEL                EQU 014Eh    ; Set current selection
CB_FINDSTRINGEXACT          EQU 0158h    ; Find exact string match

; ==============================================================================
; COMBO BOX NOTIFICATIONS
; ==============================================================================
CBN_SELCHANGE               EQU 1        ; Selection changed
CBN_EDITCHANGE              EQU 5        ; Edit text changed
CBN_SETFOCUS                EQU 3        ; Combo box got focus

; ==============================================================================
; MRU (MOST RECENTLY USED) CONFIGURATION
; ==============================================================================
MRU_MAX_ITEMS               EQU 5        ; Maximum MRU list items

; ==============================================================================
; CONTROL IDs
; ==============================================================================
IDC_COMBO                   EQU 101      ; ComboBox control ID
IDC_EDIT                    EQU 101      ; Edit control ID (same as combo)
IDC_BTN_RUN                 EQU 102      ; Run button ID
IDC_STATIC_STATUS           EQU 103      ; Status label ID
IDC_BTN_BROWSE              EQU 104      ; Browse button ID

; ==============================================================================
; MENU COMMAND IDs
; ==============================================================================
ID_FILE_BROWSE              EQU 201      ; File -> Browse menu item
ID_FILE_CONSOLE             EQU 202      ; File -> Console menu item
ID_FILE_EXIT                EQU 203      ; File -> Exit menu item
ID_FILE_ABOUT               EQU 204      ; File -> About menu item

; ==============================================================================
; STANDARD RESOURCES
; ==============================================================================
IDC_ARROW_ATOM              EQU 32512    ; Standard arrow cursor
WHITE_BRUSH                 EQU 0        ; White brush for background
SW_SHOWNORMAL               EQU 1        ; Show window normally

; ==============================================================================
; STANDARD HANDLES
; ==============================================================================
STD_INPUT_HANDLE            EQU -10      ; Standard input handle
STD_OUTPUT_HANDLE           EQU -11      ; Standard output handle
STD_ERROR_HANDLE            EQU -12      ; Standard error handle
STARTF_USESTDHANDLES        EQU 00000100h ; Use standard handles

; ==============================================================================
; MENU FLAGS
; ==============================================================================
MF_STRING                   EQU 00000000h ; Menu item is string
MF_POPUP                    EQU 00000010h ; Menu item is popup
MF_SEPARATOR                EQU 00000800h ; Menu separator

; ==============================================================================
; OPEN FILE DIALOG FLAGS
; ==============================================================================
OFN_FILEMUSTEXIST           EQU 00001000h ; File must exist
OFN_PATHMUSTEXIST           EQU 00000800h ; Path must exist

; ==============================================================================
; COM CONSTANTS
; ==============================================================================
CLSCTX_INPROC_SERVER        EQU 1        ; In-process server
MSGFLT_ALLOW                EQU 1        ; Allow message through filter

; ==============================================================================
; VIRTUAL KEY CODES
; ==============================================================================
VK_ESCAPE                   EQU 1Bh      ; ESC key

; ==============================================================================
; CONSOLE ATTACH FLAGS
; ==============================================================================
ATTACH_PARENT_PROCESS       EQU 0FFFFFFFFh ; Attach to parent console

; ==============================================================================
; REGISTRY CONSTANTS
; ==============================================================================
HKEY_CURRENT_USER           EQU 80000001h ; Current user registry hive
HKEY_CLASSES_ROOT           EQU 80000000h ; Classes root registry hive
KEY_READ                    EQU 00020019h ; Read access to registry key
KEY_WRITE                   EQU 00020006h ; Write access to registry key
REG_SZ                      EQU 00000001h ; String registry value type

; ==============================================================================
; ERROR CONSTANTS
; ==============================================================================
INVALID_HANDLE_VALUE        EQU -1       ; Invalid handle value

; ==============================================================================
; MESSAGE STRUCTURE
; ==============================================================================
MSG STRUCT
    hwnd    QWORD ?         ; Window handle
    message DWORD ?         ; Message identifier
    pad1    DWORD ?         ; Padding for alignment
    wParam  QWORD ?         ; First message parameter
    lParam  QWORD ?         ; Second message parameter
    time    DWORD ?         ; Message time
    pt_x    DWORD ?         ; Cursor X position
    pt_y    DWORD ?         ; Cursor Y position
    pad2    DWORD ?         ; Padding for alignment
MSG ENDS
