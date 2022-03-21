VERSION 5.00
Begin VB.Form myForm 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "myForm"
   ClientHeight    =   3450
   ClientLeft      =   1050
   ClientTop       =   1380
   ClientWidth     =   5745
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
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3450
   ScaleWidth      =   5745
   WindowState     =   2  'Maximized
   Begin VB.Timer tmrEvent 
      Interval        =   500
      Left            =   0
      Top             =   0
   End
   Begin VB.Label lblInstructions 
      AutoSize        =   -1  'True
      Caption         =   $"Screen Saver Template.frx":0000
      Height          =   2340
      Left            =   720
      TabIndex        =   0
      Top             =   840
      Visible         =   0   'False
      Width           =   4245
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "myForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim QuitFlag As Integer  'Flag to tell if terminating screen saver
Dim CursorCount As Integer  '

'Windows API function to show and hide mouse
Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long


Private Sub Form_Click()
    QuitFlag = True  'Time to quit because of mouse click
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    QuitFlag = True  'Time to quit because a key was pressed
End Sub

Private Sub Form_Load()
    'Each time the ShowCursor(False) is called it sets the
    '"Cursor Visible Count" down one, a value of one or
    'greater means it is visible, zero or less than zero
    'means hidden, if another program is hidding it as well
    'than the screen saver program should only up it by one
    'more so the other program has to show it, usually this
    'is only done if there was no mouse installed on the
    'machine

    'Store the current count in Windows
    CursorCount = ShowCursor(False) + 1

    'Loop until the cursor is hidden
    Do While ShowCursor(False) >= -1
    Loop

    'Don't allow multiple instances of program
    If App.PrevInstance = True Then
        Unload Me  'Unload program
    End If

    'Process Setup button of Desktop control panel
    If InStr(Command$, "/c") Then
        'Re-show mouse pointer
        Do While ShowCursor(True) < 0
        Loop

        'You can put a screen saver configuration here
        'I made another form once and had it configure
        'the screen saver but for now the user is told
        'that there are no options
        
        MsgBox "No setup options for this screen saver"

        'Don't do any graphics during setup
        Unload myForm  'Unload screen saver
    End If

    'Check for preview command
    If InStr(Command$, "/p") <> False Then
        'I haven't added the code to put the screen
        'saver in the little window under the display
        'properties but it can be done using Win API
        
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Static Xlast, Ylast  'Remember values after leaving

    'Get current position of mouse
    Xnow = X
    Ynow = Y

    'On first move (usually right at startup),
    'simply record position of mouse
    If Xlast = 0 And Ylast = 0 Then
        Xlast = Xnow  'Horizontal position
        Ylast = Ynow  'Vertical position
        Exit Sub  'Leave
    End If

    'Quit only if mouse actually changes position
    If Xnow <> Xlast Or Ynow <> Ylast Then
        QuitFlag = True  'Set flag to quit
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Show cursor again
    Do While ShowCursor(True) <= CursorCount
    Loop
    
    End  'Make sure screen saver terminates
End Sub

Private Sub tmrEvent_Timer()
'This timer makes the screen saver events occur, you
'just automate with controls what you want to happen,
'I have made a picture bounce around

'Check to see if screen saver needs to unload
If QuitFlag = True Then
    Unload myForm  'Unload the form
End If

'Start adding screen saver routines here
End Sub
