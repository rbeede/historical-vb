VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "The Program Flasher"
   ClientHeight    =   1485
   ClientLeft      =   2625
   ClientTop       =   1485
   ClientWidth     =   3615
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
   Icon            =   "Title Flasher.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1485
   ScaleWidth      =   3615
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      Caption         =   "End"
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   840
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "FlashWindow"
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   1560
      Top             =   840
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "<-- Just turn on and move mouse over any program."
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   1560
      TabIndex        =   2
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Timer1.Enabled = Not (Timer1.Enabled)
End Sub

Private Sub Command2_Click()
    MsgBox "Program by Rodney Beede.  E-Mail me at rodney_beede@hotmail.com", 64, "The Program Flasher"
    End
End Sub

Private Sub Form_Load()
    Me.Top = Screen.Height / 2 - Me.Height / 2
    Me.Left = Screen.Width / 2 - Me.Width / 2
End Sub

Private Sub Timer1_Timer()
    Dim rc As Integer
    
    GetCursorPos MP
    X1% = MP.X
    Y1% = MP.Y
    Windy% = WindowFromPoint(X1%, Y1%)
    
    rc = FlashWindow(Windy%, 1)
End Sub

