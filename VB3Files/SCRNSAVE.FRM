VERSION 2.00
Begin Form myForm 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "myForm"
   ClientHeight    =   2145
   ClientLeft      =   1050
   ClientTop       =   1380
   ClientWidth     =   3690
   ForeColor       =   &H000000FF&
   Height          =   2550
   Left            =   990
   LinkTopic       =   "Form1"
   ScaleHeight     =   2145
   ScaleWidth      =   3690
   Top             =   1035
   Width           =   3810
   WindowState     =   2  'Maximized
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

Declare Function ShowCursor Lib "User" (ByVal bShow As Integer) As Integer

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

        'Do any user interaction
        MsgBox "No setup options for this screen saver"

        'Hide mouse pointer
        X% = ShowCursor(False)

        'Don't do any graphics during setup
        Unload myform
        Exit Sub
    End If
    T = 1
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

Sub Timer1_Timer ()
  Unload myform
End Sub

Sub Timer2_Timer ()
' this is where the screen saver events should happen
If QuitFlag = True Then
   timer1.Enabled = True
   timer2.Enabled = False
End If
If Me.Visible = False Then QuitFlag = True


End Sub

