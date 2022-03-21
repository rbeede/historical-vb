VERSION 5.00
Begin VB.Form frmCode 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Copy this code for your table"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "Web Page Calender Code Maker Code Form.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtCode 
      Height          =   495
      Left            =   480
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "Copy"
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   495
      Left            =   1800
      TabIndex        =   0
      Top             =   2040
      Width           =   1215
   End
End
Attribute VB_Name = "frmCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdCopy_Click()
    txtCode.SelStart = 0
    txtCode.SelLength = Len(txtCode.Text)
    
    Clipboard.Clear
    Clipboard.SetText (txtCode.Text)
    
    txtCode.SetFocus
End Sub

Private Sub Form_Load()
    Me.Top = 0
    Me.Left = 0
    Me.Height = Screen.Height
    Me.Width = Screen.Width

    cmdCopy.Top = Me.Height - cmdCopy.Height - 500
    cmdClose.Top = cmdCopy.Top
    
    cmdClose.Left = (Me.Width / 2) + 200
    cmdCopy.Left = 200

    cmdCopy.Width = Me.Width / 2 - 200
    cmdClose.Width = cmdCopy.Width - 200

    txtCode.Left = cmdCopy.Left
    txtCode.Top = 0
    txtCode.Width = Me.Width - 400
    txtCode.Height = Me.Height - cmdCopy.Height - 1000
End Sub
