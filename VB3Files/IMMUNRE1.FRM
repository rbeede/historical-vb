VERSION 2.00
Begin Form FrmSearch 
   BorderStyle     =   3  'Fixed Double
   Caption         =   "Search"
   ClientHeight    =   4005
   ClientLeft      =   1095
   ClientTop       =   1485
   ClientWidth     =   4230
   ControlBox      =   0   'False
   Height          =   4410
   Left            =   1035
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4005
   ScaleWidth      =   4230
   Top             =   1140
   Width           =   4350
   Begin TextBox Text1 
      Height          =   285
      Left            =   240
      TabIndex        =   3
      Top             =   480
      Width           =   2295
   End
   Begin CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   2760
      TabIndex        =   2
      Top             =   1080
      Width           =   1215
   End
   Begin CommandButton Command1 
      Caption         =   "Go To"
      Default         =   -1  'True
      Height          =   495
      Left            =   2760
      TabIndex        =   1
      Top             =   480
      Width           =   1215
   End
   Begin ListBox List1 
      Height          =   2955
      Left            =   240
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   840
      Width           =   2295
   End
   Begin Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "NAME:"
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   600
   End
End
Declare Function SendMessage& Lib "User" (ByVal hWnd%, ByVal wMsg%, ByVal wParam%, lParam As Any)
Declare Function SendMessageByNum& Lib "User" Alias "SendMessage" (ByVal hWnd%, ByVal wMsg%, ByVal wParam%, ByVal lParam&)
Declare Function SendMessageByString& Lib "User" Alias "SendMessage" (ByVal hWnd%, ByVal wMsg%, ByVal wParam%, ByVal lParam$)

Const WM_USER = &H400
Const LB_ERR = -1
Const LB_SELECTSTRING = (WM_USER + 13)
Const LB_FINDSTRING = (WM_USER + 16)
Const LB_SETTABSTOPS = (WM_USER + 19)
Const LB_SETTOPINDEX = (WM_USER + 24)
Const LB_FINDSTRINGEXACT = (WM_USER + 35)

Dim ByPass As Integer
Dim ListSelect As Integer

Sub Command1_Click ()
Dim Passed As Integer, I As Integer

Passed = False
ByPass = True

list1.Enabled = False
For I = 0 To list1.ListCount - 1
  list1.ListIndex = I
  list1.Selected(list1.ListIndex) = True
  If Text1.Text = list1.Text Then Passed = True
Next I
ByPass = False
list1.Enabled = True

If Text1.Text <> "" And Passed = True Then
   frmMain!Data1.Recordset.MoveFirst
   Do Until Text1.Text = frmMain!Text1.Text
      frmMain!Data1.Recordset.MoveNext
   Loop
   Me.Hide
   Unload Me
   frmMain.Enabled = True
End If
End Sub

Sub Command2_Click ()
frmMain.Enabled = True
Unload Me
End Sub

Sub Form_Load ()

On Error Resume Next
frmMain!Data1.Refresh
frmMain!Data1.Recordset.MoveFirst
If frmMain!Data1.Recordset.BOF And frmMain!Data1.Recordset.EOF Then Unload Me

list1.Clear
Do While Not frmMain!Data1.Recordset.EOF
  list1.AddItem frmMain!Text1.Text
  frmMain!Data1.Recordset.MoveNext
Loop
frmMain!Data1.Refresh
ByPass = False
Text1.SetFocus
Me.Show
End Sub

Sub List1_Click ()
If ByPass = False Then
    ListSelect = True
    Text1.Text = list1.Text
    ListSelect = False
End If
End Sub

Sub Text1_Change ()
Dim sSearchString, nStringPosition As Integer
Dim rc As Integer
  If ByPass = False Then
    If Text1.Text = "" And ListSelect = False Then
         list1.Selected(list1.ListIndex) = False
    End If

    sSearchString = Text1.Text
    If (sSearchString <> "") Then
        nStringPosition = SendMessageByString(list1.hWnd, LB_FINDSTRING, 0, sSearchString)
        If nStringPosition <> LB_ERR Then
            rc = SendMessageByNum(list1.hWnd, LB_SETTOPINDEX, nStringPosition, 0)
        End If
    End If
  End If
End Sub

