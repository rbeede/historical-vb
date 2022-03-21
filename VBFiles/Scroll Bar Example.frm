VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Scroll Bar Example"
   ClientHeight    =   3765
   ClientLeft      =   1095
   ClientTop       =   1485
   ClientWidth     =   3195
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3765
   ScaleWidth      =   3195
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "Quit"
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton cmdAdd 
      Appearance      =   0  'Flat
      Caption         =   "&Add Label"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   1680
      Width           =   1215
   End
   Begin VB.VScrollBar Scroll 
      Height          =   1695
      Left            =   2520
      Max             =   6
      TabIndex        =   0
      Top             =   1680
      Width           =   255
   End
   Begin VB.Label lblWhy 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "This example shows how to make a scrowing area for controls.  They fall within a certain part of the form."
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   360
      TabIndex        =   3
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label label 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Label1"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   1680
      TabIndex        =   1
      Top             =   1680
      Width           =   585
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim NumOfLbl As Integer
Dim X As Integer

Private Sub cmdAdd_Click()

Scroll.Value = 0

label(0).Top = Scroll.Top

NumOfLbl = NumOfLbl + 1
Load label(NumOfLbl)

For X = 1 To NumOfLbl
   label(X).Top = label(X - 1).Top + label(X - 1).Height + 100
   label(X).Caption = "Label" & X + 1
   label(X).Visible = True
Next X

If NumOfLbl > 5 Then
   For X = 6 To NumOfLbl
      label(X).Visible = False
   Next X
End If

Scroll.Max = NumOfLbl
End Sub

Private Sub Command1_Click()
End
End Sub

Private Sub Form_Load()
Me.Top = (Screen.Height - Me.Height) / 2
Me.Left = (Screen.Width - Me.Width) / 2

label(0).Top = Scroll.Top

For X = 1 To 5
   Load label(X)
   label(X).Top = label(X - 1).Top + label(X - 1).Height + 100
   label(X).Caption = "Label" & X + 1
   label(X).Visible = True
   NumOfLbl = X
Next X

Scroll.Max = NumOfLbl
End Sub

Private Sub label_Click(Index As Integer)
MsgBox Str$(Index), , label(Index).Caption
End Sub

Private Sub Scroll_Change()
On Error Resume Next

For X = 6 To 1 Step -1
   label(Scroll.Value - X).Visible = False
Next X

label(Scroll.Value).Top = Scroll.Top
label(Scroll.Value).Visible = True

For X = 1 To 5 Step 1
   label(Scroll.Value + X).Top = label(Scroll.Value + (X - 1)).Top + 100 + label(Scroll.Value).Height
   label(Scroll.Value + X).Visible = True
Next X

label(Scroll.Value + 6).Top = label(Scroll.Value + 5).Top + 100 + label(Scroll.Value).Height
label(Scroll.Value + 6).Visible = False

For X = 0 To NumOfLbl
   label(X).Caption = "Label" & X + 1
   label(X).Refresh
Next X
End Sub

