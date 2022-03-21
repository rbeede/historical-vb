VERSION 2.00
Begin Form FrmSettings 
   BorderStyle     =   3  'Fixed Double
   Caption         =   "Settings"
   ClientHeight    =   3450
   ClientLeft      =   1095
   ClientTop       =   1485
   ClientWidth     =   7365
   ControlBox      =   0   'False
   Height          =   3855
   Left            =   1035
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3450
   ScaleWidth      =   7365
   Top             =   1140
   Width           =   7485
   Begin CommandButton Command2 
      Caption         =   "Click Here To Save These Settings"
      Height          =   495
      Left            =   3240
      TabIndex        =   8
      Top             =   2880
      Width           =   3135
   End
   Begin CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   855
      Left            =   6480
      TabIndex        =   7
      Top             =   2040
      Width           =   735
   End
   Begin HScrollBar Scroll 
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
   Begin TextBox Text1 
      Height          =   285
      Left            =   3240
      TabIndex        =   2
      Text            =   "c:\windows\*.bmp"
      Top             =   480
      Width           =   3975
   End
   Begin PictureBox Picture1 
      Height          =   3000
      Left            =   120
      ScaleHeight     =   2970
      ScaleWidth      =   2970
      TabIndex        =   0
      Top             =   120
      Width           =   3000
   End
   Begin Label Label4 
      Caption         =   "Screen saver that moves a picture around the screen bouncing it at the edges."
      Height          =   735
      Left            =   3240
      TabIndex        =   6
      Top             =   2040
      Width           =   3015
   End
   Begin Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Slow-----------------------------------------------------Fast"
      Height          =   195
      Left            =   3240
      TabIndex        =   5
      Top             =   1440
      Width           =   3960
   End
   Begin Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Speed to move picture."
      Height          =   195
      Left            =   3240
      TabIndex        =   3
      Top             =   1080
      Width           =   1995
   End
   Begin Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Path and File Name for picture."
      Height          =   195
      Left            =   3240
      TabIndex        =   1
      Top             =   120
      Width           =   2670
   End
End
'Declares for WindowsDirectory
Declare Function GetWindowsDirectory Lib "Kernel" (ByVal P$, ByVal S%) As Integer

Sub command1_click ()
    Unload Me
End Sub

Sub Command2_Click ()
    Call iniFile
    Call command1_click
End Sub

Function Exists% (F$)
On Error Resume Next
X& = FileLen(F$)
If X& Then Exists% = True
End Function

Sub Form_Load ()
    Call openini
    Me.Show
End Sub

Sub iniFile ()
    Dim WinPath As String

    WinPath = WindowsDirectory()

    On Error Resume Next
    Open WinPath + "\bouncer.ini" For Output As #2
    

    Write #2, text1.Text, scroll.Value
    
    Close
    
End Sub

Sub openini ()
    Dim WinPath As String

    WinPath = WindowsDirectory()

    On Error Resume Next
    Open WinPath + "\scrnpic.ini" For Input As #1
    

    Input #1, Pic, Speed
    
    If Err = 52 Then
       text1.Text = "c:\windows\*.bmp"
       scroll.Value = 60
    ElseIf Not (Err) Then
       text1.Text = Pic
       scroll.Value = Speed
    End If
    Close

End Sub

Sub Text1_KeyPress (KeyAscii As Integer)
    If KeyAscii = 13 Then
       KeyAscii = 0
       If InStr(text1.Text, ".bmp") Or InStr(text1.Text, ".ico") Or InStr(text1.Text, ".pcx") Then
          dummy$ = text1.Text
          Existab = Exists%(dummy$)
          If Existab = True Then
             picture1.Picture = LoadPicture(text1.Text)
          Else
             MsgBox "File does not exist.", 16, "Error"
          End If
       End If
    End If
End Sub

Function WindowsDirectory () As String
Dim WinPath As String
    WinPath = String(145, Chr(0))
    WindowsDirectory = Left(WinPath, GetWindowsDirectory(WinPath, Len(WinPath)))
End Function

