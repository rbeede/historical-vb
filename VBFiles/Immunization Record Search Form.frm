VERSION 5.00
Begin VB.Form frmSearch 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Search"
   ClientHeight    =   4005
   ClientLeft      =   1095
   ClientTop       =   1485
   ClientWidth     =   4230
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4005
   ScaleWidth      =   4230
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   240
      TabIndex        =   3
      Top             =   480
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   2760
      TabIndex        =   2
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "Go To"
      Default         =   -1  'True
      Height          =   495
      Left            =   2760
      TabIndex        =   1
      Top             =   480
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   2955
      Left            =   240
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   840
      Width           =   2295
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "NAME:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   600
   End
End
Attribute VB_Name = "FrmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ByPass As Integer
Dim ListSelect As Integer

Private Sub Command1_Click()
Dim Passed As Integer, i As Integer

Passed = False
ByPass = True

List1.Enabled = False
For i = 0 To List1.ListCount - 1
  List1.ListIndex = i
  List1.Selected(List1.ListIndex) = True
  If Text1.Text = List1.Text Then Passed = True
Next i
ByPass = False
List1.Enabled = True

If Text1.Text <> "" And Passed = True Then
   FrmMain!Data1.Recordset.MoveFirst
   Do Until Text1.Text = FrmMain!Text1.Text
      FrmMain!Data1.Recordset.MoveNext
   Loop
   Me.Hide
   Unload Me
   FrmMain.Enabled = True
   FrmMain.Show
End If
End Sub

Private Sub Command2_Click()
FrmMain.Enabled = True
Unload Me
End Sub

Private Sub Form_Load()

On Error Resume Next
FrmMain!Data1.Refresh
FrmMain!Data1.Recordset.MoveFirst
If FrmMain!Data1.Recordset.BOF And FrmMain!Data1.Recordset.EOF Then Unload Me

List1.Clear
Do While Not FrmMain!Data1.Recordset.EOF
  List1.AddItem FrmMain!Text1.Text
  FrmMain!Data1.Recordset.MoveNext
Loop
FrmMain!Data1.Refresh
ByPass = False
Text1.SetFocus
Me.Show
End Sub

Private Sub List1_Click()
If ByPass = False Then
    ListSelect = True
    Text1.Text = List1.Text
    ListSelect = False
End If
End Sub

Private Sub Text1_Change()
  If ByPass = False Then
    If Text1.Text = "" And ListSelect = False Then
         List1.Selected(List1.ListIndex) = False
    End If
  End If
End Sub

