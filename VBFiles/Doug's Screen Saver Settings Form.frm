VERSION 5.00
Begin VB.Form FrmSettings 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Settings"
   ClientHeight    =   3450
   ClientLeft      =   1095
   ClientTop       =   1485
   ClientWidth     =   7365
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
   ScaleHeight     =   3450
   ScaleWidth      =   7365
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      Caption         =   "Click Here To Save These Settings"
      Height          =   495
      Left            =   3240
      TabIndex        =   8
      Top             =   2880
      Width           =   3135
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   855
      Left            =   6480
      TabIndex        =   7
      Top             =   2040
      Width           =   735
   End
   Begin VB.HScrollBar Scroll 
      Height          =   255
      LargeChange     =   20
      Left            =   3240
      Max             =   520
      Min             =   20
      SmallChange     =   20
      TabIndex        =   4
      Top             =   1680
      Value           =   20
      Width           =   3975
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3240
      TabIndex        =   2
      Text            =   "c:\windows\*.bmp"
      Top             =   480
      Width           =   3975
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3000
      Left            =   120
      ScaleHeight     =   2970
      ScaleWidth      =   2970
      TabIndex        =   0
      Top             =   120
      Width           =   3000
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Screen saver that moves a picture around the screen bouncing it at the edges."
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   3240
      TabIndex        =   6
      Top             =   2040
      Width           =   3015
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Slow-----------------------------------------------------Fast"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3240
      TabIndex        =   5
      Top             =   1440
      Width           =   3960
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Speed to move picture."
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3240
      TabIndex        =   3
      Top             =   1080
      Width           =   1995
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Path and File Name for picture."
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3240
      TabIndex        =   1
      Top             =   120
      Width           =   2670
   End
End
Attribute VB_Name = "FrmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Declares for WindowsDirectory
Private Declare Function GetWindowsDirectory Lib "Kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Private Sub command1_click()
    Unload Me
End Sub

Private Sub Command2_Click()
    Call iniFile
    Call command1_click
End Sub

Private Function Exists%(F$)
On Error Resume Next
X& = FileLen(F$)
If X& Then Exists% = True
End Function

Private Sub Form_Load()
    Call openini
    Me.Show
End Sub

Private Sub iniFile()
    Dim WinPath As String

    WinPath = WindowsDirectory()

    On Error Resume Next
    Open WinPath + "\bouncer.ini" For Output As #2
    

    Write #2, Text1.Text, Scroll.Value
    
    Close
    
End Sub

Private Sub openini()
    Dim WinPath As String

    WinPath = WindowsDirectory()

    On Error Resume Next
    Open WinPath + "\scrnpic.ini" For Input As #1
    

    Input #1, Pic, Speed
    
    If Err = 52 Then
       Text1.Text = "c:\windows\*.bmp"
       Scroll.Value = 60
    ElseIf Not (Err) Then
       Text1.Text = Pic
       Scroll.Value = Speed
    End If
    Close

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    
    If KeyAscii = 13 Then
       KeyAscii = 0
       
        dummy$ = Text1.Text
        Existab = Exists%(dummy$)
        
        If Existab = True Then
           Picture1.Picture = LoadPicture(Text1.Text)
        Else
           MsgBox "File does not exist.", 16, "Error"
        End If

    End If
End Sub

Private Function WindowsDirectory() As String
Dim WinPath As String
    WinPath = String(145, Chr(0))
    WindowsDirectory = Left(WinPath, GetWindowsDirectory(WinPath, Len(WinPath)))
End Function

