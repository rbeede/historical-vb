VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Get Handle"
      Height          =   495
      Left            =   2280
      TabIndex        =   3
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Run Calc"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Show"
      Height          =   495
      Left            =   2280
      TabIndex        =   1
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Hide"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   1320
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long

Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
    
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long

Dim ProgHwnd

Const WM_CLOSE = &H10
Const SW_HIDE = 0
Const SW_MAXIMIZE = 3
Const SW_SHOW = 5
Const SW_MINIMIZE = 6

Sub WindowHandle(win, cas As Long)


    'by storm
    'Case 0 = CloseWindow
    'Case 1 = Show Win
    'Case 2 = Hide Win
    'Case 3 = Max Win
    'Case 4 = Min Win


    Select Case cas
        Case 0:
        Dim X%
        X% = SendMessage(win, WM_CLOSE, 0, 0)
        Case 1:
        X = ShowWindow(win, SW_SHOW)
        Case 2:
        X = ShowWindow(win, SW_HIDE)
        Case 3:
        X = ShowWindow(win, SW_MAXIMIZE)
        Case 4:
        X = ShowWindow(win, SW_MINIMIZE)
    End Select


'any questions e-mail me at storm@n2.com
End Sub

Private Sub Command1_Click()
    Call WindowHandle(ProgHwnd, 2)
End Sub

Private Sub Command2_Click()
    Call WindowHandle(ProgHwnd, 1)
End Sub

Private Sub Command3_Click()
    Dim X
    
    X = Shell("c:\windows\calc.exe", 6)
End Sub

Private Sub Command4_Click()
    
    ProgHwnd = GetModuleHandle("c:\windows\calc.exe")
    Form1.Print Str$(GetModuleHandle("c:\windows\notepad.exe"))
End Sub
