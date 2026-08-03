; ==============================================================================
; CMDT - Run as TrustedInstaller
; String Utility Module
;
; Author: Marek Wesolowski (wesmar)
; Ported to: Native ARM64 (armasm64) for Windows on ARM
;
; Purpose: Self-contained wide-character string helpers shared across modules.
;          No external symbols required; functions are pure manipulations of
;          caller-provided buffers and produce no side effects on globals.
;
; Exported routines:
;   DecryptWideStr  - XOR-decrypt a wide string into a destination buffer
;   wcscpy_p        - Copy a null-terminated wide string
;   wcscat_p        - Concatenate a wide string onto an existing buffer
;   wcscmp_ci       - Case-insensitive wide-string comparison
;   wcscmp_token    - Like wcscmp_ci, but treats space as token terminator
;   skip_spaces     - Advance a wide-string pointer past leading spaces
;   wcslen_p        - Length of a null-terminated wide string, in characters
;
; ARM64 Port Notes:
;   - x64 RSI/RDI (callee-saved) mapped to x19/x20.
;   - x64 byte-level XOR in DecryptWideStr uses LDRB/STRB (byte load/store).
;   - x64 word-level operations use LDRH/STRH (halfword = 16-bit wide char).
;   - skip_spaces and wcslen_p are leaf functions with no callee-saved
;     register usage — no prologue/epilogue needed.
;   - All functions preserve the exact return-value semantics of the x64
;     originals (wcscmp_ci returns 1 on match, 0 on mismatch).
; ==============================================================================

; ==============================================================================
; CODE SECTION
; ==============================================================================
    AREA |.text|, CODE, READONLY, ALIGN=4

    EXPORT DecryptWideStr
    EXPORT wcscpy_p
    EXPORT wcscat_p
    EXPORT wcscmp_ci
    EXPORT wcscmp_token
    EXPORT skip_spaces
    EXPORT wcslen_p

; ==============================================================================
; DecryptWideStr - XOR Decrypt Wide String In-Place
;
; Purpose: Decrypts a XOR-encrypted wide character string into destination
;          buffer. Uses simple XOR with single-byte key applied to each byte.
;
; Parameters:
;   x0 = Pointer to encrypted source string
;   x1 = Pointer to destination buffer
;
; Returns:
;   x0 = Pointer to destination buffer (same as x1 input)
;
; Register allocation:
;   x19 = source pointer (encrypted)
;   x20 = destination pointer
;   w2  = scratch byte register
;
; Notes:
;   - XOR key is hardcoded as 0xAA
;   - Decryption stops at null terminator (0x0000)
;   - Each byte of the wide string is XORed independently
; ==============================================================================
DecryptWideStr PROC
    STP x29, x30, [sp, #-16]!
    MOV x29, sp
    STP x19, x20, [sp, #-16]!

    MOV x19, x0                 ; x19 = source (encrypted)
    MOV x20, x1                 ; x20 = destination
    MOV w4, #0xAA               ; w4 = XOR key (0xAA is not a valid logical
                                ; immediate, so hold it in a register)

dws_loop
    ; Decrypt first byte of wide char
    LDRB w2, [x19]             ; Load byte from source
    EOR w2, w2, w4             ; XOR with key
    STRB w2, [x20]             ; Store to destination

    ; Decrypt second byte of wide char
    LDRB w2, [x19, #1]         ; Load second byte
    EOR w2, w2, w4             ; XOR with key
    STRB w2, [x20, #1]         ; Store second byte

    ; Check if we hit null terminator
    LDRH w2, [x20]
    CBZ w2, dws_done           ; Null terminator reached

    ; Move to next character
    ADD x19, x19, #2
    ADD x20, x20, #2
    B dws_loop

dws_done
    MOV x0, x20                 ; Return destination pointer

    LDP x19, x20, [sp], #16
    LDP x29, x30, [sp], #16
    RET
DecryptWideStr ENDP

; ==============================================================================
; wcscpy_p - Wide Character String Copy (Private Implementation)
;
; Purpose: Copies a null-terminated wide character string from source to dest.
;
; Parameters:
;   x0 = Destination buffer pointer
;   x1 = Source string pointer
;
; Returns: None
;
; Register allocation:
;   x19 = destination pointer
;   x20 = source pointer
;   w2  = scratch character register
; ==============================================================================
wcscpy_p PROC
    STP x29, x30, [sp, #-16]!
    MOV x29, sp
    STP x19, x20, [sp, #-16]!

    MOV x19, x0                 ; x19 = destination
    MOV x20, x1                 ; x20 = source

wcp_loop
    LDRH w2, [x20]             ; Read wide char from source
    STRH w2, [x19]             ; Write to destination
    CBZ w2, wcp_done           ; Exit if null terminator
    ADD x20, x20, #2           ; Advance source
    ADD x19, x19, #2           ; Advance destination
    B wcp_loop

wcp_done
    LDP x19, x20, [sp], #16
    LDP x29, x30, [sp], #16
    RET
wcscpy_p ENDP

; ==============================================================================
; wcscat_p - Wide Character String Concatenate (Private Implementation)
;
; Purpose: Appends source string to the end of destination string.
;
; Parameters:
;   x0 = Destination buffer pointer
;   x1 = Source string pointer
;
; Returns: None
;
; Register allocation:
;   x19 = destination pointer (advances to end, then appends)
;   x20 = source pointer
;   w2  = scratch character register
; ==============================================================================
wcscat_p PROC
    STP x29, x30, [sp, #-16]!
    MOV x29, sp
    STP x19, x20, [sp, #-16]!

    MOV x19, x0                 ; x19 = destination

    ; Find end of destination string
wca_find_end
    LDRH w2, [x19]
    CBZ w2, wca_copy           ; Found null terminator
    ADD x19, x19, #2
    B wca_find_end

wca_copy
    MOV x20, x1                 ; x20 = source

wca_copy_loop
    LDRH w2, [x20]             ; Read wide char from source
    STRH w2, [x19]             ; Write to destination
    CBZ w2, wca_done           ; Exit if null terminator
    ADD x20, x20, #2           ; Advance source
    ADD x19, x19, #2           ; Advance destination
    B wca_copy_loop

wca_done
    LDP x19, x20, [sp], #16
    LDP x29, x30, [sp], #16
    RET
wcscat_p ENDP

; ==============================================================================
; wcscmp_ci - Wide Character String Compare (Case-Insensitive)
;
; Purpose: Compares two wide character strings ignoring case differences.
;
; Parameters:
;   x0 = First string pointer
;   x1 = Second string pointer
;
; Returns:
;   x0 = 1 if strings are equal (case-insensitive), 0 otherwise
;
; Register allocation:
;   x19 = first string pointer
;   x20 = second string pointer
;   w2  = character from first string (normalized)
;   w3  = character from second string (normalized)
;
; Normalization: ASCII uppercase A-Z (0x41-0x5A) converted to lowercase
; by adding 32 (0x20). Non-ASCII characters are compared as-is.
; ==============================================================================
wcscmp_ci PROC
    STP x29, x30, [sp, #-16]!
    MOV x29, sp
    STP x19, x20, [sp, #-16]!

    MOV x19, x0                 ; x19 = first string
    MOV x20, x1                 ; x20 = second string

wci_loop
    LDRH w2, [x19]             ; Load char from first string
    LDRH w3, [x20]             ; Load char from second string

    ; Convert first character to lowercase if uppercase (A-Z)
    CMP w2, #'A'
    B.LT wci_skip1
    CMP w2, #'Z'
    B.GT wci_skip1
    ADD w2, w2, #32            ; Convert A-Z to a-z

wci_skip1
    ; Convert second character to lowercase if uppercase (A-Z)
    CMP w3, #'A'
    B.LT wci_skip2
    CMP w3, #'Z'
    B.GT wci_skip2
    ADD w3, w3, #32            ; Convert A-Z to a-z

wci_skip2
    CMP w2, w3                  ; Compare normalized characters
    B.NE wci_not_eq            ; Characters differ

    CBZ w2, wci_equal          ; Both null terminators reached

    ADD x19, x19, #2           ; Advance first string
    ADD x20, x20, #2           ; Advance second string
    B wci_loop

wci_equal
    MOV w0, #1                  ; Return 1 (strings equal)
    B wci_done

wci_not_eq
    MOV w0, WZR                 ; Return 0 (strings differ)

wci_done
    LDP x19, x20, [sp], #16
    LDP x29, x30, [sp], #16
    RET
wcscmp_ci ENDP

; ==============================================================================
; wcscmp_token - Compare command-line token against a null-terminated literal
;
; Purpose: Like wcscmp_ci but treats a space character in the first argument
;          as an end-of-string marker. Used to recognize switches inside a
;          raw command-line buffer without temporarily mutating the buffer.
;
; Parameters:
;   x0 = command-line pointer (a token followed by space or null)
;   x1 = null-terminated literal to match (e.g. str_outfileFlag)
;
; Returns:
;   x0 = 1 if the literal matches the token (case-insensitive), 0 otherwise
;
; Register allocation:
;   x19 = cmdline cursor
;   x20 = literal pointer
;   w2  = literal character (normalized)
;   w3  = token character (normalized)
;
; Matching semantics:
;   - Literal is consumed character by character.
;   - Token ends at space or null. If literal ends but token continues
;     with a non-space character, it's NOT a match.
;   - If both end simultaneously (or token ends with space), it IS a match.
; ==============================================================================
wcscmp_token PROC
    STP x29, x30, [sp, #-16]!
    MOV x29, sp
    STP x19, x20, [sp, #-16]!

    MOV x19, x0                 ; x19 = cmdline cursor
    MOV x20, x1                 ; x20 = literal

wctk_loop
    LDRH w2, [x20]             ; Load char from literal
    CBZ w2, wctk_lit_end       ; Literal exhausted

    LDRH w3, [x19]             ; Load char from token
    CBZ w3, wctk_no            ; Token ended early (null)
    CMP w3, #' '
    B.EQ wctk_no               ; Token ended early (space)

    ; Normalize literal character (A-Z → a-z)
    CMP w2, #'A'
    B.LT wctk_skip1
    CMP w2, #'Z'
    B.GT wctk_skip1
    ADD w2, w2, #32

wctk_skip1
    ; Normalize token character (A-Z → a-z)
    CMP w3, #'A'
    B.LT wctk_skip2
    CMP w3, #'Z'
    B.GT wctk_skip2
    ADD w3, w3, #32

wctk_skip2
    CMP w2, w3
    B.NE wctk_no               ; Characters differ

    ADD x19, x19, #2
    ADD x20, x20, #2
    B wctk_loop

wctk_lit_end
    ; Literal finished. Check if token also finished (space or null).
    LDRH w3, [x19]
    CBZ w3, wctk_yes           ; Token ended with null
    CMP w3, #' '
    B.EQ wctk_yes              ; Token ended with space

wctk_no
    MOV w0, WZR                 ; Return 0 (no match)
    B wctk_done

wctk_yes
    MOV w0, #1                  ; Return 1 (match)

wctk_done
    LDP x19, x20, [sp], #16
    LDP x29, x30, [sp], #16
    RET
wcscmp_token ENDP

; ==============================================================================
; skip_spaces - Skip Leading Whitespace in Wide String
;
; Purpose: Advances a string pointer past any leading space characters.
;
; Parameters:
;   x0 = String pointer
;
; Returns:
;   x0 = Pointer to first non-space character
;
; Note: This is a leaf function with no callee-saved register usage.
;       No prologue/epilogue needed.
; ==============================================================================
skip_spaces PROC
ss_loop
    LDRH w1, [x0]              ; Load current character
    CMP w1, #' '               ; Check if space
    B.NE ss_done               ; Non-space found
    ADD x0, x0, #2             ; Skip this space
    B ss_loop

ss_done
    RET
skip_spaces ENDP

; ==============================================================================
; wcslen_p - Wide Character String Length (Private Implementation)
;
; Purpose: Calculates the length of a null-terminated wide character string.
;
; Parameters:
;   x0 = String pointer
;
; Returns:
;   x0 = Number of characters (excluding null terminator)
;
; Register allocation:
;   x0 = string pointer (input), length (output)
;   x1 = length counter (scratch)
;   w2 = scratch character register
;
; Note: This is a leaf function with no callee-saved register usage.
;       No prologue/epilogue needed.
; ==============================================================================
wcslen_p PROC
    MOV x1, XZR                 ; x1 = length counter = 0

wl_loop
    LDRH w2, [x0, x1, LSL #1] ; Load char at offset x1*2
    CBZ w2, wl_done            ; Null terminator found
    ADD x1, x1, #1             ; Increment length
    B wl_loop

wl_done
    MOV x0, x1                  ; Return length in x0
    RET
wcslen_p ENDP

    END