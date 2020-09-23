;Extract.asm  by Robert Rayment  Nov 2002

; ASM layout done to roughly match VB layout

; VB
; Public ptrMC, ptrStruc    ' Ptrs to Machine Code & Structure
; 'MCode Structure
; Public Type PStruc
;   picWd As Long
;   picHt As Long
;   Ptrpic1mem As Long      ' pic1mem(1-4,1-picWd,1-picHt) bytes
;   Ptrpic2mem As Long      ' pic2mem(1-4,1-picWd,1-picHt) bytes
;   StartColor As Long
;   PtrDELRGB As Long       ' DELRGB(0,1,2,3) Grey, +/1 R,G,B
; End Type
; Public DStruc As PStruc
; EG
; Public ExtractMC() As Byte
; Public Sub ASM_Extract()
; res = CallWindowProc(ptrMC, ptrStruc, 2&, 3&, 4&)
;                             8         12  16  20
; End Sub

%macro movab 2      ;name & num of parameters
  push dword %2     ;2nd param
  pop dword %1      ;1st param
%endmacro           ;use  movab %1,%2
;Allows eg  movab bmW,[ebx+4]

;Define names to match VB code
%define picWd          [ebp-4]
%define picHt          [ebp-8]
%define Ptrpic1mem     [ebp-12]
%define Ptrpic2mem     [ebp-16]
%define StartColor     [ebp-20]
%define PtrDELRGB      [ebp-24]

; Some variables
%define R1         [ebp-28]
%define G1         [ebp-32]
%define B1         [ebp-36]
%define R          [ebp-52]
%define G          [ebp-56]
%define B          [ebp-60]
%define R1minus    [ebp-64]    
%define G1minus    [ebp-68]    
%define B1minus    [ebp-72]    
%define R1plus     [ebp-76]    
%define G1plus     [ebp-80]    
%define B1plus     [ebp-84]

[bits 32]

    push ebp
    mov ebp,esp
    sub esp,84
    push edi
    push esi
    push ebx

    ;Fill structure
    mov ebx,            [ebp+8]
    movab picWd,        [ebx]
    movab picHt,        [ebx+4]
    movab Ptrpic1mem,   [ebx+8]
    movab Ptrpic2mem,   [ebx+12]
    movab StartColor,   [ebx+16]
    movab PtrDELRGB,    [ebx+20]
;----------------------------

;   Get R1,G1,B1 from StartColor
    xor eax,eax
    mov R1,eax
    mov G1,eax
    mov B1,eax
    mov eax,StartColor
    mov R1,AL
    mov G1,AH
    bswap eax
    mov B1,AH

;   Get R1minus/plus,G1minus/plus,B1minus/plus
    
    mov edi,PtrDELRGB

    mov eax,R1
    mov ebx,[edi+4] ; DELRGB(1) +/- red
    sub eax,ebx     ; R1-DELRGB(1)
    mov R1minus,eax
    add eax,ebx
    add eax,ebx
    mov R1plus,eax      ; R1+DELRGB(1)

    mov eax,G1
    mov ebx,[edi+8] ; DELRGB(2) +/- green
    sub eax,ebx     ; G1-DELRGB(2)
    mov G1minus,eax
    add eax,ebx
    add eax,ebx
    mov G1plus,eax      ; G1+DELRGB(2)

    mov eax,B1
    mov ebx,[edi+12]    ; DELRGB(3) +/- blue
    sub eax,ebx     ; B1-DELRGB(3)
    mov B1minus,eax
    add eax,ebx
    add eax,ebx
    mov B1plus,eax      ; B1+DELRGB(3)


;   Test each pixel's RGB against 
;   R1,G1,B1 +/-  DELRGB(1,2,3)

    mov eax,picHt
    mov ebx,picWd
    mul ebx
    mov ecx,eax         ; picHt*picWd no. of 4-bytes chunks 
                        ; in pic1mem & pic2mem
    
    mov esi,Ptrpic1mem  ; esi->pic1mem(1,1,1)
    mov edi,Ptrpic2mem  ; edi->pic2mem(1,1,1)

ForChunk:

    xor eax,eax
    mov R,eax
    mov G,eax
    mov B,eax

    mov eax,[esi]       ; Lo-BGRA-Hi 32bit order in pic1mem
    mov B,AL
    mov G,AH
    bswap eax
    mov R,AH

    ; TestR
    mov eax,R
    cmp eax,R1minus
    jl Blacken
    cmp eax,R1plus
    jg Blacken

    ; TestG
    mov eax,G
    cmp eax,G1minus
    jl Blacken
    cmp eax,G1plus
    jg Blacken

    ; TestB
    mov eax,B
    cmp eax,B1minus
    jl Blacken
    cmp eax,B1plus
    jg Blacken

    ; Whiten
    mov eax,0FFFFFFh
    jmp FillPic2mem
Blacken:
    xor eax,eax
FillPic2mem:
    mov [edi],eax

NexChunk:
    mov eax,4
    add esi,eax
    add edi,eax
    dec ecx
    jnz ForChunk


GETOUT:
    pop ebx
    pop esi
    pop edi
    mov esp,ebp
    pop ebp
    ret 16

;=====================================================================
