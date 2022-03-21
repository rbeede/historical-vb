VERSION 5.00
Begin VB.Form frmdaybirth 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Days Till Birthdays"
   ClientHeight    =   3705
   ClientLeft      =   2010
   ClientTop       =   1545
   ClientWidth     =   3360
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
   Icon            =   "DAYS UNTIL BIRTHDAY.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3705
   ScaleWidth      =   3360
   Begin VB.CommandButton cmdtiffany 
      Appearance      =   0  'Flat
      Caption         =   "&Tiffany"
      Height          =   495
      Left            =   1800
      TabIndex        =   7
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton cmdsandy 
      Appearance      =   0  'Flat
      Caption         =   "&Sandy"
      Height          =   495
      Left            =   1800
      TabIndex        =   6
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton cmddavid 
      Appearance      =   0  'Flat
      Caption         =   "&David"
      Height          =   495
      Left            =   1800
      TabIndex        =   5
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton cmdjason 
      Appearance      =   0  'Flat
      Caption         =   "Jas&on"
      Height          =   495
      Left            =   360
      TabIndex        =   4
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton cmdrodney 
      Appearance      =   0  'Flat
      Caption         =   "&Rodney"
      Height          =   495
      Left            =   360
      TabIndex        =   3
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton cmdanna 
      Appearance      =   0  'Flat
      Caption         =   "&Anna"
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton cmdexit 
      Appearance      =   0  'Flat
      Caption         =   "E&xit"
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   2880
      Width           =   2655
   End
   Begin VB.Label lbltodaysdate 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "lbltodaysdate"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   1140
   End
End
Attribute VB_Name = "frmdaybirth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdanna_Click()
    Dim Msg As String

    MsgBox "The number of days till Anna's birthday is " & DateDiff("d", Now, "01/14"), 64, Me.Caption

End Sub

Private Sub cmddavid_Click()
    Dim Msg As String

    MsgBox "The number of days till David's birthday is " & DateDiff("d", Now, "12/14"), 64, Me.Caption
    
End Sub

Private Sub cmdexit_Click()
    End
End Sub

Private Sub cmdjason_Click()
    Dim Msg As String

    MsgBox "The number of days till Jasons's birthday is " & DateDiff("d", Now, "10/12"), 64, Me.Caption

End Sub

Private Sub cmdrodney_Click()
    Dim Msg As String

    MsgBox "The number of days till Rodney's birthday is " & DateDiff("d", Now, "08/20"), 64, Me.Caption

End Sub

Private Sub cmdsandy_Click()
    Dim Msg As String

    MsgBox "The number of days till Sandy's birthday is " & DateDiff("d", Now, "6/30"), 64, Me.Caption

End Sub

Private Sub cmdtiffany_Click()
    Dim Msg As String

    MsgBox "The number of days till Tiffany's birthday is " & DateDiff("d", Now, "01/5"), 64, Me.Caption

End Sub

Private Sub Form_Load()
    lbltodaysdate.Caption = "Todays date is " & Date$
End Sub

