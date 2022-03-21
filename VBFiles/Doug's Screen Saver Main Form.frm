VERSION 5.00
Begin VB.Form myForm 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "myForm"
   ClientHeight    =   3060
   ClientLeft      =   1050
   ClientTop       =   1380
   ClientWidth     =   4830
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H000000FF&
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3060
   ScaleWidth      =   4830
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      DrawStyle       =   5  'Transparent
      ForeColor       =   &H80000008&
      Height          =   720
      Left            =   1800
      Picture         =   "Doug's Screen Saver Main Form.frx":0000
      ScaleHeight     =   720
      ScaleWidth      =   720
      TabIndex        =   0
      Top             =   1320
      Width           =   720
   End
   Begin VB.Timer Timer2 
      Interval        =   500
      Left            =   3240
      Top             =   0
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "myForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim QuitFlag As Integer
Dim CursorCount As Integer
Dim T As Integer
Dim Motion As Integer
Dim Direction As Integer
Dim MoveSpeed As Integer


Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Private Sub Form_Click()
QuitFlag = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
QuitFlag = True
End Sub

Private Sub Form_Load()
   'Record original mouse-pointer show count
    CursorCount = 1
    
    'Hide mouse pointer
    Do While ShowCursor(False) >= -1
    Loop
    Do While ShowCursor(True) < -1
    Loop

    'Don't allow multiple instances of program
    If App.PrevInstance = True Then
        Unload Me
        Exit Sub
    End If

    'Process Setup button of Desktop control panel
    If Command$ = "/c" Then

        'Temporarily reshow mouse pointer
        X% = ShowCursor(True)

        Load FrmSettings


        'Hide mouse pointer
        X% = ShowCursor(False)

        'Don't do any graphics during setup
        Unload myForm
        Exit Sub
    End If
    T = 1
    Motion = Int(Rnd(4) * 4) ' Set random direction to move it
    Call iniFile
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Static Xlast, Ylast

    'Get current position
    Xnow = X
    Ynow = Y

    'On first move, simply record position
    If Xlast = 0 And Ylast = 0 Then
        Xlast = Xnow
        Ylast = Ynow
        Exit Sub
    End If

    'Quit only if mouse actually changes position
    If Xnow <> Xlast Or Ynow <> Ylast Then
        QuitFlag = True
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Do While ShowCursor(False) >= CursorCount
    Loop
    Do While ShowCursor(True) < CursorCount
    Loop
End Sub

Private Sub iniFile()
    Dim WinPath As String

    WinPath = WindowsDirectory()

    On Error Resume Next
    Open WinPath + "\bouncer.ini" For Input As #1
    

    Input #1, Pic, Speed
    
    If Err = 52 Then
       Picture1.Picture = Picture1.Picture
       MoveSpeed = 20
    ElseIf Not (Err) Then
       Picture1.Picture = LoadPicture(Pic)
       MoveSpeed = Speed
    End If
    Close
    
End Sub

Private Sub MovePicture(Motion)
   Static MoveSpeedPic As Integer
   MoveSpeedPic = MoveSpeed
   Select Case Motion
    Case 1
        ' Move graphic left/up by 20 twips using Move method
        Picture1.Move Picture1.Left - MoveSpeed, Picture1.Top - MoveSpeed
        ' If graphic reaches left edge of form, move it right/up
        If Picture1.Left <= 0 Then
            Motion = 2
        ' If graphic reaches top edge of form, move it left/down
        ElseIf Picture1.Top <= 0 Then
            Motion = 4
        End If
    Case 2
        ' Move graphic right/up by 20 twips
        Picture1.Move Picture1.Left + MoveSpeed, Picture1.Top - MoveSpeed
        ' If the graphic reaches right edge of form, move left/up.
        ' Routine determines right edge of form by subtracting graphic
        ' width from form width
        If Picture1.Left >= (myForm.Width - Picture1.Width) Then
            Motion = 1
        ' If graphic reaches top edge of form, move right/down
        ElseIf Picture1.Top <= 0 Then
            Motion = 3
        End If
    Case 3
        ' Move graphic right/down by 20 twips
        Picture1.Move Picture1.Left + MoveSpeed, Picture1.Top + MoveSpeed
        ' If graphic reaches right edge of form, move left/down
        If Picture1.Left >= (myForm.Width - Picture1.Width) Then
            Motion = 4
        ' If graphic reaches bottom edge of form, move right/up.
        ' Routine determines bottom of form by subtracting
        ' graphic height from form height less 680 twips for height
        ' of title bar and menu bar
        ElseIf Picture1.Top >= (myForm.Height - Picture1.Height) - 680 Then
            Motion = 2
        End If
    Case 4
        ' Move the graphic left/down by 20 twips
        Picture1.Move Picture1.Left - MoveSpeed, Picture1.Top + MoveSpeed
        ' If graphic reaches left edge of form, move right/down
        If Picture1.Left <= 0 Then
            Motion = 3
        ' If graphic reaches bottom edge of the form, move left/up
        ElseIf Picture1.Top >= (myForm.Height - Picture1.Height) - 680 Then
            Motion = 1
        End If
    End Select
End Sub

Private Sub Timer1_Timer()
  Unload myForm
End Sub

Private Sub Timer2_Timer()
' this is where the screen saver events should happen
If QuitFlag = True Then
   Timer1.Enabled = True
   Timer2.Enabled = False
End If
If Me.Visible = False Then QuitFlag = True

Call MovePicture(Motion) ' Run sub MovePicture
End Sub

Private Function WindowsDirectory() As String
Dim WinPath As String
    WinPath = String(145, Chr(0))
    WindowsDirectory = Left(WinPath, GetWindowsDirectory(WinPath, Len(WinPath)))
End Function

