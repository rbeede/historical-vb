VERSION 5.00
Object = "{B16553C3-06DB-101B-85B2-0000C009BE81}#1.0#0"; "SPIN32.OCX"
Begin VB.Form frmGuess 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Break the code!"
   ClientHeight    =   2145
   ClientLeft      =   3225
   ClientTop       =   2970
   ClientWidth     =   2220
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2145
   ScaleWidth      =   2220
   Begin VB.CommandButton cmdExit 
      Appearance      =   0  'Flat
      Caption         =   "&End"
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   1320
      Width           =   1695
   End
   Begin VB.CommandButton cmdguess 
      Appearance      =   0  'Flat
      Caption         =   "&Guess"
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   720
      Width           =   1695
   End
   Begin Spin.SpinButton Spinner3 
      Height          =   375
      Left            =   1680
      TabIndex        =   5
      Top             =   240
      Width           =   255
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   1323
      _StockProps     =   73
   End
   Begin Spin.SpinButton Spinner2 
      Height          =   375
      Left            =   1080
      TabIndex        =   6
      Top             =   240
      Width           =   255
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   1323
      _StockProps     =   73
   End
   Begin Spin.SpinButton Spinner1 
      Height          =   375
      Left            =   480
      TabIndex        =   7
      Top             =   240
      Width           =   255
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   1323
      _StockProps     =   73
   End
   Begin VB.Label lblNumber3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   240
      Width           =   255
   End
   Begin VB.Label lblNumber2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   240
      Width           =   255
   End
   Begin VB.Label lblNumber1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   255
   End
End
Attribute VB_Name = "frmGuess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Declares Variables Used in Form
Dim Number1, Number2, Number3, Number As Integer

Private Sub cmdExit_Click()
    End 'Ends program
End Sub

Private Sub cmdguess_Click()
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

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 7 Then MsgBox Str$(Number), , "Cheat"
End Sub

Private Sub Form_Load()
    'Stores current number to guess
    Number1 = 0
    Number2 = 0
    Number3 = 0

    'Get a random number for code
    Randomize
    Number = Rnd(999) * 999 / 1
End Sub

Private Sub Spinner1_SpinDown()
    
    'Subtract one from number one and show in caption
    Number1 = Number1 - 1
    If Number1 < 0 Then Number1 = 9
    lblNumber1.Caption = Number1
    DoEvents 'Let windows catch up
End Sub

Private Sub Spinner1_SpinUp()
    
    'Add one to number one and show in caption
    Number1 = Number1 + 1
    If Number1 > 9 Then Number1 = 0
    lblNumber1.Caption = Number1
    DoEvents 'Let windows catch up
End Sub

Private Sub Spinner2_SpinDown()
    
    'Subtract one from number two and show in caption
    Number2 = Number2 - 1
    If Number2 < 0 Then Number2 = 9
    lblNumber2.Caption = Number2
    DoEvents 'Let windows catch up
End Sub

Private Sub Spinner2_SpinUp()
    
    'Add one to number two and show in caption
    Number2 = Number2 + 1
    If Number2 > 9 Then Number2 = 0
    lblNumber2.Caption = Number2
    DoEvents 'Let windows catch up
End Sub

Private Sub Spinner3_SpinDown()
    
    'Subtract one from number three and show in caption
    Number3 = Number3 - 1
    If Number3 < 0 Then Number3 = 9
    lblNumber3.Caption = Number3
    DoEvents 'Let windows catch up
End Sub

Private Sub Spinner3_SpinUp()
    
    'Add one to number three and show in caption
    Number3 = Number3 + 1
    If Number3 > 9 Then Number3 = 0
    lblNumber3.Caption = Number3
    DoEvents 'Let windows catch up
End Sub

