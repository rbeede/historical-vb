;InpOut.DLL
;Inp and Out Visual Basic Keyword replacements
;Copyright 1991 Crescent Software
;Written by Jay Munro
;
.286
.Model Medium
Public Inp,Out,Wep
Extrn UnlockSegment:Proc        ;use LibW.LIB from Win SDK
Extrn LocalInit:Proc

WinProlog Macro
  Push DS                       ;useless setup code
  Pop  AX
  Nop
  Inc  BP                       ;adjust BP
  Push BP
  Mov  BP,SP                    ;set up stack frame
  Push DS
  Mov  DS,AX
EndM

WinEpilog Macro
  Dec  BP
  Dec  BP
  Mov  SP,BP
  Pop  DS
  Pop  BP
  Dec  BP
EndM

.Data
  Required_Data_Header DB 16 dup (?)     ;as required by Windows for .ASM Dlls 

INIT_TEXT  SEGMENT BYTE PUBLIC 'CODE'
Assume CS:INIT_TEXT

LibEntry Proc Far
        WinProlog

        push    di               ; handle of the module instance
        push    ds               ; library data segment
        push    cx               ; heap size

        ; if we have some heap then initialize it
        jcxz    CallMain         ; jump if no heap specified

        ; call the Windows function LocalInit() to set up the heap
        ; LocalInit((LPSTR)start, WORD cbHeap);

        xor     ax,ax
        Push    DS
        Push    AX
        Push    CX
        Call    LocalInit
        Or      ax,ax            ; did it do it ok ?
        Jz      error            ; quit if it failed

        ; invoke the asm routine to do any special initialization

CallMain:
        Call Far Ptr   LibMain          ; invoke the startup routine (result in AX)
        Jmp short exit           ; LibMain is responsible for stack clean up

error:
        pop     cx
        pop     ds
        pop     di

exit:
        WinEpilog
        Ret
LibEntry EndP
INIT_TEXT EndS

.Code

;  Libmain gets these parameters
;    parmW hInstance              ;handle [bp+14]
;    parmW hDataSeg               ;word   [bp+12]
;    parmW cbHeapSize             ;word   [bp+10]

LibMain Proc Far
   WinProlog                    ;windows prolog code
   Push -1                      ;unlock data segment (just in case)
   Call UnlockSegment
   Mov  AX,1                    ;return a 1 to caller
   WinEpilog                    ;windows epilog code
   Ret 6                       ;3 variables (6 bytes)

LibMain EndP

Wep Proc Far
   WinProlog
   Mov  AX,1
   WinEpilog
   Ret
Wep EndP

;Declare Function Inp Lib "InpOut.DLL" (Port%) as Integer
;PortNumb% = &h378
;PortValue% = Inp(PortNumb%)

Inp Proc Far
   WinProlog                    ;standard prolog macro
   Mov   DX,[BP+6]              ;load DX with port to check out
   In    AL,DX                  ;input a byte    
   Xor   AH,AH                  ;clear high byte for return in AX
   WinEpilog                    ;standard epilog macro
   Ret 2
Inp EndP

;Declare Sub Out Lib "InpOut.DLL" (Port%,Value%)
;PortNumb% = &h378                      ;for example only, don't try this
;Value% = 65
;Out PortNumb%,Value%

Out Proc Far    ;Warning errors can be ignored!
   WinProlog                    ;standard prolog macro
   Mov  AL,[BP+6]               ;get data value into AL (bytes only)
   Mov  DX,[BP+8]               ;get port value into DX
   Out  DX,AL                   ;output the byte
   WinEpilog                    ;standard epilog macro
   Ret 4                        ;return clearing 4 bytes off stack
Out EndP

End
