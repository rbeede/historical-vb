VERSION 2.00
Begin Form Form1 
   BorderStyle     =   3  'Fixed Double
   Caption         =   "The Program Flasher"
   ClientHeight    =   1485
   ClientLeft      =   2625
   ClientTop       =   1485
   ClientWidth     =   3450
   Height          =   1890
   Icon            =   FLASHER.FRX:0000
   Left            =   2565
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1485
   ScaleWidth      =   3450
   Top             =   1140
   Width           =   3570
   Begin CommandButton Command2 
      Caption         =   "End"
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   840
      Width           =   1215
   End
   Begin CommandButton Command1 
      Caption         =   "FlashWindow"
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   1320
      Top             =   840
   End
   Begin Label Label1 
      Caption         =   "<-- Just turn on and move mouse over any program."
      Height          =   615
      Left            =   1320
      TabIndex        =   2
      Top             =   120
      Width           =   1935
   End
End

Sub Command1_Click ()
    Timer1.Enabled = Not (Timer1.Enabled)
End Sub

Sub Command2_Click ()
    MsgBox "Program by Rodney Beede.  E-Mail me at rodney_beede@hotmail.com", 64, "The Program Flasher"
    End
End Sub

Sub Form_Load ()
    Me.Top = screen.Height / 2 - Me.Height / 2
    Me.Left = screen.Width / 2 - Me.Width / 2
End Sub

Sub Timer1_Timer ()
    GetCursorPos MP
    X1% = MP.X
    Y1% = MP.Y
    Windy% = WindowFromPoint(X1%, Y1%)
    
    Dim rc As Integer
    rc = FlashWindow(Windy%, 1)
End Sub

