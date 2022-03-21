VERSION 2.00
Begin Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6690
   ClientLeft      =   1020
   ClientTop       =   1365
   ClientWidth     =   3195
   ClipControls    =   0   'False
   Height          =   7095
   Left            =   960
   LinkTopic       =   "Form1"
   ScaleHeight     =   6690
   ScaleWidth      =   3195
   Top             =   1020
   Width           =   3315
   Begin CommandButton Command2 
      Caption         =   "API Search"
      Height          =   495
      Left            =   1800
      TabIndex        =   3
      Top             =   4440
      Width           =   1215
   End
   Begin CommandButton Command1 
      Caption         =   "VB Search"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   4440
      Width           =   1215
   End
   Begin TextBox Text1 
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   3840
      Width           =   2895
   End
   Begin ListBox List1 
      Height          =   2370
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2895
   End
   Begin Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Type in item to be searched and click on search method. Item ranges from ITEM1 to ITEM5000. Time is displayed in milliseconds. The VB clock only increments every 55 ms."
      Height          =   1215
      Left            =   120
      TabIndex        =   6
      Top             =   2520
      Width           =   2895
   End
   Begin Label Label2 
      Caption         =   "Label2"
      Height          =   495
      Left            =   1800
      TabIndex        =   5
      Top             =   5040
      Width           =   1215
   End
   Begin Label Label1 
      Caption         =   "Label1"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   5040
      Width           =   1215
   End
End
Option Explicit
' current time
Declare Function GetCurrentTime Lib "User" () As Long

' message sender function
Declare Function SendMessage Lib "User" (ByVal hWnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, lParam As Any) As Long

Const WM_USER = &H400
Const LB_FINDSTRINGEXACT = (WM_USER + 35)

Sub Command1_Click ()
' search for string using VB methods

' reset timer label
label1.Caption = ""
DoEvents
' time variable
Dim lTime As Long
Dim sFind As String
sFind = Trim$(Text1.Text)
' save start time
lTime = GetCurrentTime()
' search list
Dim j As Integer
For j = 0 To List1.ListCount - 1
  ' match?
  If List1.List(j) = sFind Then
    Exit For
  End If
Next j

' if found
If j < List1.ListCount Then
  ' display elapsed time
  label1.Caption = GetCurrentTime() - lTime
Else
  label1.Caption = "No Find"
End If

End Sub

Sub Command2_Click ()
' search for string using API function

' reset label timer
Label2.Caption = ""
DoEvents

' time variable
Dim lTime As Long
' save starting time
lTime = GetCurrentTime()
' search list
Dim iNdx As Integer
' store txt value in temporary string
Dim sFind As String
sFind = Text1.Text
iNdx = SendMessage(List1.hWnd, LB_FINDSTRINGEXACT, -1, ByVal sFind)
If iNdx > -1 Then
  ' found string. display elapsed time
  Label2.Caption = GetCurrentTime() - lTime
Else
  ' string not found
  Label2.Caption = "No Find"
End If

End Sub

Sub Form_Activate ()
Text1.SetFocus
End Sub

Sub Form_Load ()
' reset the labels
label1.Caption = ""
Label2.Caption = ""
' fill the list box with elements
Dim j As Integer
For j = 1 To 5000
  List1.AddItem "ITEM" & Format$(j)
Next j
End Sub

