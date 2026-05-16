; ==============================================================================
; CMDT - Run as TrustedInstaller (x86)
; String Utility Module
;
; Author: Marek Wesolowski (wesmar)
; Purpose: Self-contained wide-character string helpers shared across modules.
;          No external symbols required; functions are pure manipulations of
;          caller-provided buffers and produce no side effects on globals.
;
; Exported routines (stdcall):
;   wcscpy_p        - Copy a null-terminated wide string
;   wcscat_p        - Concatenate a wide string onto an existing buffer
;   wcscmp_ci       - Case-insensitive wide-string comparison
;   wcscmp_token    - Like wcscmp_ci, but treats space as token terminator
;   skip_spaces     - Advance a wide-string pointer past leading spaces
;   wcslen_p        - Length of a null-terminated wide string, in characters
; ==============================================================================

.586
.model flat, stdcall
option casemap:none

.code

; ==============================================================================
; wcscpy_p - Wide Character String Copy
; ==============================================================================
wcscpy_p proc dest:DWORD, src:DWORD
    push esi
    push edi
    mov edi, dest
    mov esi, src
@@:
    mov ax, word ptr [esi]
    mov word ptr [edi], ax
    test ax, ax
    jz @F
    add esi, 2
    add edi, 2
    jmp @B
@@:
    pop edi
    pop esi
    ret
wcscpy_p endp

; ==============================================================================
; wcscat_p - Wide Character String Concatenate
; ==============================================================================
wcscat_p proc dest:DWORD, src:DWORD
    push esi
    push edi
    mov edi, dest
@@:
    cmp word ptr [edi], 0
    je @F
    add edi, 2
    jmp @B
@@:
    mov esi, src
@@:
    mov ax, word ptr [esi]
    mov word ptr [edi], ax
    test ax, ax
    jz @F
    add esi, 2
    add edi, 2
    jmp @B
@@:
    pop edi
    pop esi
    ret
wcscat_p endp

; ==============================================================================
; wcscmp_ci - Wide Character String Compare (Case-Insensitive)
; ==============================================================================
wcscmp_ci proc str1:DWORD, str2:DWORD
    push esi
    push edi
    mov esi, str1
    mov edi, str2
wci_loop:
    mov ax, word ptr [esi]
    mov dx, word ptr [edi]

    cmp ax, 'A'
    jb wci_skip1
    cmp ax, 'Z'
    ja wci_skip1
    add ax, 32
wci_skip1:
    cmp dx, 'A'
    jb wci_skip2
    cmp dx, 'Z'
    ja wci_skip2
    add dx, 32
wci_skip2:
    cmp ax, dx
    jne wci_not_eq
    test ax, ax
    jz wci_equal
    add esi, 2
    add edi, 2
    jmp wci_loop
wci_equal:
    pop edi
    pop esi
    mov eax, 1
    ret
wci_not_eq:
    pop edi
    pop esi
    xor eax, eax
    ret
wcscmp_ci endp

; ==============================================================================
; skip_spaces - Skip Leading Whitespace
; ==============================================================================
skip_spaces proc lpStr:DWORD
    mov eax, lpStr
@@:
    cmp word ptr [eax], ' '
    jne @F
    add eax, 2
    jmp @B
@@:
    ret
skip_spaces endp

; ==============================================================================
; wcslen_p - Wide Character String Length
; ==============================================================================
wcslen_p proc lpStr:DWORD
    mov eax, lpStr
    xor ecx, ecx
@@:
    cmp word ptr [eax + ecx*2], 0
    je @F
    inc ecx
    jmp @B
@@:
    mov eax, ecx
    ret
wcslen_p endp

; ==============================================================================
; wcscmp_token - Compare a raw command-line token with a literal
;
; Treats a space in lpToken as end-of-token. Returns 1 on match, 0 otherwise.
; Used to detect cmdline switches inside a buffer that can't be temporarily
; null-terminated.
; ==============================================================================
wcscmp_token proc uses esi edi lpToken:DWORD, lpLiteral:DWORD
    mov esi, lpToken
    mov edi, lpLiteral
wct_loop:
    mov ax, word ptr [edi]
    test ax, ax
    jz wct_lit_end
    mov dx, word ptr [esi]
    test dx, dx
    jz wct_no
    cmp dx, ' '
    je wct_no

    cmp ax, 'A'
    jb @F
    cmp ax, 'Z'
    ja @F
    add ax, 32
@@:
    cmp dx, 'A'
    jb @F
    cmp dx, 'Z'
    ja @F
    add dx, 32
@@:
    cmp ax, dx
    jne wct_no
    add esi, 2
    add edi, 2
    jmp wct_loop

wct_lit_end:
    mov dx, word ptr [esi]
    test dx, dx
    jz wct_yes
    cmp dx, ' '
    je wct_yes
wct_no:
    xor eax, eax
    ret
wct_yes:
    mov eax, 1
    ret
wcscmp_token endp

end
