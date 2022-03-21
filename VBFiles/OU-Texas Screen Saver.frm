VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form myForm 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "myForm"
   ClientHeight    =   6435
   ClientLeft      =   1050
   ClientTop       =   1380
   ClientWidth     =   10305
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
   ScaleHeight     =   6435
   ScaleWidth      =   10305
   WindowState     =   2  'Maximized
   Begin MCI.MMControl sndPlayer 
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   4440
      Visible         =   0   'False
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   873
      _Version        =   327681
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.Timer tmrEvent 
      Interval        =   500
      Left            =   0
      Top             =   0
   End
   Begin VB.Image imgOUHall 
      Appearance      =   0  'Flat
      Height          =   3135
      Left            =   600
      Picture         =   "OU-Texas Screen Saver.frx":0000
      Top             =   2640
      Visible         =   0   'False
      Width           =   2130
   End
   Begin VB.Image imgOUTexas 
      Height          =   1200
      Left            =   240
      Picture         =   "OU-Texas Screen Saver.frx":79D2
      Top             =   4560
      Visible         =   0   'False
      Width           =   8865
   End
   Begin VB.Image imgOUCenter 
      Height          =   1830
      Left            =   0
      Picture         =   "OU-Texas Screen Saver.frx":13714
      Top             =   840
      Visible         =   0   'False
      Width           =   1830
   End
   Begin VB.Image imgTexasTitle 
      Height          =   1350
      Left            =   480
      Picture         =   "OU-Texas Screen Saver.frx":15EBB
      Top             =   2400
      Visible         =   0   'False
      Width           =   4950
   End
   Begin VB.Image imgTexasLonghornTitle 
      Height          =   1200
      Left            =   840
      Picture         =   "OU-Texas Screen Saver.frx":1D7B5
      Top             =   840
      Visible         =   0   'False
      Width           =   8865
   End
   Begin VB.Image imgOUTitle 
      Height          =   885
      Left            =   1560
      Picture         =   "OU-Texas Screen Saver.frx":294F7
      Top             =   240
      Visible         =   0   'False
      Width           =   5730
   End
   Begin VB.Image imgTexasBand 
      Height          =   2220
      Left            =   3000
      Picture         =   "OU-Texas Screen Saver.frx":2F1B9
      Top             =   960
      Visible         =   0   'False
      Width           =   6000
   End
   Begin VB.Image imgTexasFootballPlayer 
      Height          =   7560
      Left            =   3720
      Picture         =   "OU-Texas Screen Saver.frx":3DD3B
      Top             =   960
      Visible         =   0   'False
      Width           =   8250
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

Private Sub RunIt()
    Do Until QuitFlag = True
        DoEvents
    
        'Center the Texas sign
        imgTexasTitle.Top = 100
        imgTexasTitle.Left = (myForm.Width / 2) - (imgTexasTitle.Width / 2)
        imgTexasTitle.Visible = True
    
        'Pop on the texas band
        imgTexasBand.Top = myForm.Height / 2 - imgTexasBand.Height / 2
        imgTexasBand.Left = (myForm.Width / 2) - (imgTexasBand.Width / 2)
        imgTexasBand.Visible = True
    
        'Pop on the final texas sign
        imgTexasLonghornTitle.Top = myForm.Height - imgTexasLonghornTitle.Height - 200
        imgTexasLonghornTitle.Left = (myForm.Width / 2) - (imgTexasLonghornTitle.Width / 2)
        imgTexasLonghornTitle.Visible = True
    
        'Play that great Texas fight song for 1:15
        sndPlayer.Command = "close"
        sndPlayer.DeviceType = "WaveAudio"
        sndPlayer.filename = App.Path + "\texas.wav"
        sndPlayer.Command = "Open"
        
        sndPlayer.Command = "Play"
        
        Do  'just loop until either the song is over or the screen saver is killed
                DoEvents

            If QuitFlag = True Or Not sndPlayer.Mode = mciModePlay Then Exit Do
        Loop
    
        'Now for OU!!!
        sndPlayer.Command = "Close"
        
        imgTexasTitle.Visible = False
        imgTexasBand.Visible = False
        imgTexasLonghornTitle.Visible = False
        
        'Show the scared Texas football player
        imgTexasFootballPlayer.Top = Me.Height / 2 - (imgTexasFootballPlayer.Height / 2)
        imgTexasFootballPlayer.Left = Me.Width / 2 - imgTexasFootballPlayer.Width / 2
        imgTexasFootballPlayer.Visible = True
        
        'Only show football player for like 3 secs
        Dim startTime
        
        startTime = Timer
        Do Until QuitFlat = True Or Timer - startTime > 3
            DoEvents
        Loop
        
        imgTexasFootballPlayer.Visible = False
    
        'Lets see the REAL stuff
        imgOUTitle.Top = 100
        imgOUTitle.Left = myForm.Width / 2 - imgOUTitle.Width / 2
        imgOUTitle.Visible = True
    
        imgOUHall.Top = myForm.Height / 2 - imgOUHall.Height / 2
        imgOUHall.Left = myForm.Width / 2 - imgOUHall.Width / 2
        imgOUHall.Visible = True
        
        imgOUCenter.Top = imgOUHall.Top - imgOUCenter.Height
        imgOUCenter.Left = myForm.Width / 2 - imgOUCenter.Width / 2
        imgOUCenter.Visible = True
    
        imgOUTexas.Top = myForm.Height - imgOUTexas.Height - 200
        imgOUTexas.Left = myForm.Width / 2 - imgOUTexas.Width / 2
        imgOUTexas.Visible = True
        
        'Play Boomer
        sndPlayer.filename = App.Path + "\boomer.wav"
        sndPlayer.Command = "Open"
        
        sndPlayer.Command = "Play"
        
        Do  'just loop until either the song is over or the screen saver is killed
                DoEvents

            If QuitFlag = True Or Not sndPlayer.Mode = mciModePlay Then Exit Do
        Loop
        
        imgOUTitle.Visible = False
        imgOUHall.Visible = False
        imgOUCenter.Visible = False
        imgOUTexas.Visible = False
        
        startTime = Timer
        Do Until QuitFlat = True Or Timer - startTime > 2
            DoEvents
        Loop
    Loop
End Sub

Private Sub Form_Click()
    QuitFlag = True  'Time to quit because of mouse click
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 83 Then
        sndPlayer.Command = "Stop"
        Exit Sub
    End If
    
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

    'Process Setup button of Desktop control panel
    If Command$ = "/c" Then

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
    
        'Still haven't done it, just give no preview
        Unload myForm
    End If

    Me.Show  'Show me

    'Run the screen saver stuff
    Call RunIt

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
    'Hide it all
    imgTexasTitle.Visible = False
    imgTexasBand.Visible = False
    imgTexasLonghornTitle.Visible = False
    
    imgTexasFootballPlayer.Visible = False
    
    imgOUTitle.Visible = False
    imgOUHall.Visible = False
    imgOUCenter.Visible = False
    imgOUTexas.Visible = False

    
    'Spout off my name
    myForm.FontSize = 24
    myForm.AutoRedraw = True
    myForm.Print "Made by Rodney Beede."
    myForm.Print "rodney_beede@hotmail.com"
    startTime = Timer
    Do Until QuitFlat = True Or Timer - startTime > 2
        DoEvents
    Loop
    
    myForm.Cls
    
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
End Sub
