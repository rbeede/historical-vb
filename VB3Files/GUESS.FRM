VERSION 2.00
Begin Form frmGuess 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Double
   Caption         =   "Break the code!"
   ClientHeight    =   2145
   ClientLeft      =   3225
   ClientTop       =   2970
   ClientWidth     =   2220
   ControlBox      =   0   'False
   Height          =   2550
   KeyPreview      =   -1  'True
   Left            =   3165
   LinkTopic       =   "Form1"
   ScaleHeight     =   2145
   ScaleWidth      =   2220
   Top             =   2625
   Width           =   2340
   Begin CommandButton cmdExit 
      Caption         =   "&End"
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   1320
      Width           =   1695
   End
   Begin CommandButton cmdguess 
      Caption         =   "&Guess"
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   720
      Width           =   1695
   End
   Begin SpinButton Spinner3 
      Height          =   375
      Left            =   1680
      Top             =   240
      Width           =   255
   End
   Begin SpinButton Spinner2 
      Height          =   375
      Left            =   1080
      Top             =   240
      Width           =   255
   End
   Begin SpinButton Spinner1 
      Height          =   375
      Left            =   480
      Top             =   240
      Width           =   255
   End
   Begin Label lblNumber3 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   240
      Width           =   255
   End
   Begin Label lblNumber2 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   240
      Width           =   255
   End
   Begin Label lblNumber1 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   255
   End
End
'Declares Variables Used in Form
Dim Number1, Number2, Number3, Number As Integer

Sub cmdExit_Click ()
    End 'Ends program
End Sub

Sub cmdguess_Click ()
    Dim Guess As String ' Declare Variable

    'Get guess stored
    Guess = Number1 & Number2 & Number3
    
    'If guess correct tell user correct guessif wrong tell user wrong guess
    If Guess Like Number Then
       MsgBox "Good job you are right.  Try a new one now.", 64, "Good Guess"
    Else
       MsgBox "Try again with a new number.", 48, "Wrong Guess"
    End If

    'Get a new number
    Call Form_Load

    'Reset numbers and there captions
    Number1 = 0
    Number2 = 0
    Number3 = 0

    lblNumber1.Caption = "0"
    lblNumber2.Caption = "0"
    lblNumber3.Caption = "0"
End Sub

Sub Form_KeyDown (KeyCode As Integer, Shift As Integer)
    If Shift = 7 Then MsgBox Str$(Number), , "Cheat"
End Sub

Sub Form_Load ()
    'Stores current number to guess
    Number1 = 0
    Number2 = 0
    Number3 = 0

    'Get a random number for code
    Randomize
    Number = Rnd(999) * 999 / 1
End Sub

Sub Spinner1_SpinDown ()
    
    'Subtract one from number one and show in caption
    Number1 = Number1 - 1
    If Number1 < 0 Then Number1 = 9
    lblNumber1.Caption = Number1
    DoEvents 'Let windows catch up
End Sub

Sub Spinner1_SpinUp ()
    
    'Add one to number one and show in caption
    Number1 = Number1 + 1
    If Number1 > 9 Then Number1 = 0
    lblNumber1.Caption = Number1
    DoEvents 'Let windows catch up
End Sub

Sub Spinner2_SpinDown ()
    
    'Subtract one from number two and show in caption
    Number2 = Number2 - 1
    If Number2 < 0 Then Number2 = 9
    lblNumber2.Caption = Number2
    DoEvents 'Let windows catch up
End Sub

Sub Spinner2_SpinUp ()
    
    'Add one to number two and show in caption
    Number2 = Number2 + 1
    If Number2 > 9 Then Number2 = 0
    lblNumber2.Caption = Number2
    DoEvents 'Let windows catch up
End Sub

Sub Spinner3_SpinDown ()
    
    'Subtract one from number three and show in caption
    Number3 = Number3 - 1
    If Number3 < 0 Then Number3 = 9
    lblNumber3.Caption = Number3
    DoEvents 'Let windows catch up
End Sub

Sub Spinner3_SpinUp ()
    
    'Add one to number three and show in caption
    Number3 = Number3 + 1
    If Number3 > 9 Then Number3 = 0
    lblNumber3.Caption = Number3
    DoEvents 'Let windows catch up
End Sub

