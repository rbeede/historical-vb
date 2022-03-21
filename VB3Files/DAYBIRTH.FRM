VERSION 2.00
Begin Form frmdaybirth 
   BorderStyle     =   3  'Fixed Double
   Caption         =   "Days Till Birthdays"
   ClientHeight    =   3705
   ClientLeft      =   2010
   ClientTop       =   1545
   ClientWidth     =   3360
   Height          =   4110
   Icon            =   DAYBIRTH.FRX:0000
   Left            =   1950
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   3360
   Top             =   1200
   Width           =   3480
   Begin CommandButton cmdtiffany 
      Caption         =   "&Tiffany"
      Height          =   495
      Left            =   1800
      TabIndex        =   7
      Top             =   2160
      Width           =   1215
   End
   Begin CommandButton cmdsandy 
      Caption         =   "&Sandy"
      Height          =   495
      Left            =   1800
      TabIndex        =   6
      Top             =   1440
      Width           =   1215
   End
   Begin CommandButton cmddavid 
      Caption         =   "&David"
      Height          =   495
      Left            =   1800
      TabIndex        =   5
      Top             =   720
      Width           =   1215
   End
   Begin CommandButton cmdjason 
      Caption         =   "Jas&on"
      Height          =   495
      Left            =   360
      TabIndex        =   4
      Top             =   2160
      Width           =   1215
   End
   Begin CommandButton cmdrodney 
      Caption         =   "&Rodney"
      Height          =   495
      Left            =   360
      TabIndex        =   3
      Top             =   1440
      Width           =   1215
   End
   Begin CommandButton cmdanna 
      Caption         =   "&Anna"
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin CommandButton cmdexit 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   2880
      Width           =   2655
   End
   Begin Label lbltodaysdate 
      AutoSize        =   -1  'True
      Caption         =   "lbltodaysdate"
      Height          =   195
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   1140
   End
End

Sub cmdanna_Click ()
    Dim Msg As String

    MsgBox "The number of days till Anna's birthday is " & DateDiff("d", Now, "01/14"), 64, Me.Caption

End Sub

Sub cmddavid_Click ()
    Dim Msg As String

    MsgBox "The number of days till David's birthday is " & DateDiff("d", Now, "12/14"), 64, Me.Caption
    
End Sub

Sub cmdexit_Click ()
    End
End Sub

Sub cmdjason_Click ()
    Dim Msg As String

    MsgBox "The number of days till Jasons's birthday is " & DateDiff("d", Now, "10/12"), 64, Me.Caption

End Sub

Sub cmdrodney_Click ()
    Dim Msg As String

    MsgBox "The number of days till Rodney's birthday is " & DateDiff("d", Now, "08/20"), 64, Me.Caption

End Sub

Sub cmdsandy_Click ()
    Dim Msg As String

    MsgBox "The number of days till Sandy's birthday is " & DateDiff("d", Now, "6/30"), 64, Me.Caption

End Sub

Sub cmdtiffany_Click ()
    Dim Msg As String

    MsgBox "The number of days till Tiffany's birthday is " & DateDiff("d", Now, "01/5"), 64, Me.Caption

End Sub

Sub Form_Load ()
    lblTodaysDate.Caption = "Todays date is " & Date$
End Sub

