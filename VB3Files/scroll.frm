VERSION 2.00
Begin Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Scroll Bar Example"
   ClientHeight    =   3765
   ClientLeft      =   1095
   ClientTop       =   1485
   ClientWidth     =   3195
   Height          =   4170
   Left            =   1035
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   3195
   Top             =   1140
   Width           =   3315
   Begin CommandButton Command1 
      Caption         =   "Quit"
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   2880
      Width           =   1215
   End
   Begin CommandButton cmdAdd 
      Caption         =   "&Add Label"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   1680
      Width           =   1215
   End
   Begin VScrollBar Scroll 
      Height          =   1695
      Left            =   2520
      Max             =   6
      TabIndex        =   0
      Top             =   1680
      Width           =   255
   End
   Begin Label lblWhy 
      Caption         =   "This example shows how to make a scrowing area for controls.  They fall within a certain part of the form."
      Height          =   1335
      Left            =   360
      TabIndex        =   3
      Top             =   120
      Width           =   2415
   End
   Begin Label label 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   195
      Index           =   0
      Left            =   1680
      TabIndex        =   1
      Top             =   1680
      Width           =   585
   End
End
Dim NumOfLbl As Integer
Dim X As Integer

Sub cmdAdd_Click ()

Scroll.Value = 0

Label(0).Top = Scroll.Top

NumOfLbl = NumOfLbl + 1
Load Label(NumOfLbl)

For X = 1 To NumOfLbl
   Label(X).Top = Label(X - 1).Top + Label(X - 1).Height + 100
   Label(X).Caption = "Label" & X + 1
   Label(X).Visible = True
Next X

If NumOfLbl > 5 Then
   For X = 6 To NumOfLbl
      Label(X).Visible = False
   Next X
End If

Scroll.Max = NumOfLbl
End Sub

Sub Command1_Click ()
End
End Sub

Sub Form_Load ()
Me.Top = (Screen.Height - Me.Height) / 2
Me.Left = (Screen.Width - Me.Width) / 2

Label(0).Top = Scroll.Top

For X = 1 To 5
   Load Label(X)
   Label(X).Top = Label(X - 1).Top + Label(X - 1).Height + 100
   Label(X).Caption = "Label" & X + 1
   Label(X).Visible = True
   NumOfLbl = X
Next X

Scroll.Max = NumOfLbl
End Sub

Sub label_Click (Index As Integer)
MsgBox Str$(Index), , Label(Index).Caption
End Sub

Sub Scroll_Change ()
On Error Resume Next

For X = 6 To 1 Step -1
   Label(Scroll.Value - X).Visible = False
Next X

Label(Scroll.Value).Top = Scroll.Top
Label(Scroll.Value).Visible = True

For X = 1 To 5 Step 1
   Label(Scroll.Value + X).Top = Label(Scroll.Value + (X - 1)).Top + 100 + Label(Scroll.Value).Height
   Label(Scroll.Value + X).Visible = True
Next X

Label(Scroll.Value + 6).Top = Label(Scroll.Value + 5).Top + 100 + Label(Scroll.Value).Height
Label(Scroll.Value + 6).Visible = False

For X = 0 To NumOfLbl
   Label(X).Caption = "Label" & X + 1
   Label(X).Refresh
Next X
End Sub

