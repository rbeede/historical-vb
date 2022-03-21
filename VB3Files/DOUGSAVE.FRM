VERSION 2.00
Begin Form myForm 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "myForm"
   ClientHeight    =   3060
   ClientLeft      =   1050
   ClientTop       =   1380
   ClientWidth     =   4830
   ForeColor       =   &H000000FF&
   Height          =   3465
   Left            =   990
   LinkTopic       =   "Form1"
   ScaleHeight     =   3060
   ScaleWidth      =   4830
   Top             =   1035
   Width           =   4950
   WindowState     =   2  'Maximized
   Begin PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      DrawStyle       =   5  'Transparent
      Height          =   720
      Left            =   1800
      Picture         =   DOUGSAVE.FRX:0000
      ScaleHeight     =   720
      ScaleWidth      =   720
      TabIndex        =   0
      Top             =   1320
      Width           =   720
   End
   Begin Timer Timer2 
      Interval        =   500
      Left            =   3240
      Top             =   0
   End
   Begin Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   0
   End
End
Dim QuitFlag As Integer
Dim CursorCount As Integer
Dim T As Integer
Dim Motion As Integer
Dim Direction As Integer
Dim MoveSpeed As Integer


Declare Function ShowCursor Lib "User" (ByVal bShow As Integer) As Integer
Declare Function GetWindowsDirectory Lib "Kernel" (ByVal P$, ByVal S%) As Integer

Sub Form_Click ()
QuitFlag = True
End Sub

Sub Form_KeyDown (KeyCode As Integer, Shift As Integer)
QuitFlag = True
End Sub

Sub Form_Load ()
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

        Load frmsettings


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

Sub Form_MouseMove (Button As Integer, Shift As Integer, X As Single, Y As Single)
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

Sub Form_Unload (Cancel As Integer)
    Do While ShowCursor(False) >= CursorCount
    Loop
    Do While ShowCursor(True) < CursorCount
    Loop
End Sub

Sub iniFile ()
    Dim WinPath As String

    WinPath = WindowsDirectory()

    On Error Resume Next
    Open WinPath + "\bouncer.ini" For Input As #1
    

    Input #1, Pic, Speed
    
    If Err = 52 Then
       picture1.Picture = picture1.Picture
       MoveSpeed = 20
    ElseIf Not (Err) Then
       picture1.Picture = LoadPicture(Pic)
       MoveSpeed = Speed
    End If
    Close
    
End Sub

Sub MovePicture (Motion)
   Static MoveSpeedPic As Integer
   MoveSpeedPic = MoveSpeed
   Select Case Motion
    Case 1
        ' Move graphic left/up by 20 twips using Move method
        picture1.Move picture1.Left - MoveSpeed, picture1.Top - MoveSpeed
        ' If graphic reaches left edge of form, move it right/up
        If picture1.Left <= 0 Then
            Motion = 2
        ' If graphic reaches top edge of form, move it left/down
        ElseIf picture1.Top <= 0 Then
            Motion = 4
        End If
    Case 2
        ' Move graphic right/up by 20 twips
        picture1.Move picture1.Left + MoveSpeed, picture1.Top - MoveSpeed
        ' If the graphic reaches right edge of form, move left/up.
        ' Routine determines right edge of form by subtracting graphic
        ' width from form width
        If picture1.Left >= (myForm.Width - picture1.Width) Then
            Motion = 1
        ' If graphic reaches top edge of form, move right/down
        ElseIf picture1.Top <= 0 Then
            Motion = 3
        End If
    Case 3
        ' Move graphic right/down by 20 twips
        picture1.Move picture1.Left + MoveSpeed, picture1.Top + MoveSpeed
        ' If graphic reaches right edge of form, move left/down
        If picture1.Left >= (myForm.Width - picture1.Width) Then
            Motion = 4
        ' If graphic reaches bottom edge of form, move right/up.
        ' Routine determines bottom of form by subtracting
        ' graphic height from form height less 680 twips for height
        ' of title bar and menu bar
        ElseIf picture1.Top >= (myForm.Height - picture1.Height) - 680 Then
            Motion = 2
        End If
    Case 4
        ' Move the graphic left/down by 20 twips
        picture1.Move picture1.Left - MoveSpeed, picture1.Top + MoveSpeed
        ' If graphic reaches left edge of form, move right/down
        If picture1.Left <= 0 Then
            Motion = 3
        ' If graphic reaches bottom edge of the form, move left/up
        ElseIf picture1.Top >= (myForm.Height - picture1.Height) - 680 Then
            Motion = 1
        End If
    End Select
End Sub

Sub Timer1_Timer ()
  Unload myForm
End Sub

Sub Timer2_Timer ()
' this is where the screen saver events should happen
If QuitFlag = True Then
   timer1.Enabled = True
   timer2.Enabled = False
End If
If Me.Visible = False Then QuitFlag = True

Call MovePicture(Motion) ' Run sub MovePicture
End Sub

Function WindowsDirectory () As String
Dim WinPath As String
    WinPath = String(145, Chr(0))
    WindowsDirectory = Left(WinPath, GetWindowsDirectory(WinPath, Len(WinPath)))
End Function

