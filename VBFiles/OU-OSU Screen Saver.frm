VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form myForm 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "myForm"
   ClientHeight    =   5880
   ClientLeft      =   1050
   ClientTop       =   1380
   ClientWidth     =   9630
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5880
   ScaleWidth      =   9630
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.TextBox Rodney 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   4935
      Left            =   7200
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "OU-OSU Screen Saver.frx":0000
      Top             =   5040
      Visible         =   0   'False
      Width           =   9015
   End
   Begin MCI.MMControl sndPlayer 
      Height          =   615
      Left            =   960
      TabIndex        =   0
      Top             =   5160
      Visible         =   0   'False
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   1085
      _Version        =   327681
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.Timer tmrEvent 
      Interval        =   500
      Left            =   0
      Top             =   0
   End
   Begin VB.Image Title2 
      Height          =   2700
      Left            =   0
      Picture         =   "OU-OSU Screen Saver.frx":008B
      Top             =   -120
      Visible         =   0   'False
      Width           =   7500
   End
   Begin VB.Image Pete2 
      Height          =   3300
      Index           =   1
      Left            =   4680
      Picture         =   "OU-OSU Screen Saver.frx":0DEF
      Top             =   1920
      Visible         =   0   'False
      Width           =   2010
   End
   Begin VB.Image Pete 
      Height          =   3300
      Index           =   1
      Left            =   4680
      Picture         =   "OU-OSU Screen Saver.frx":2277
      Top             =   1080
      Visible         =   0   'False
      Width           =   2010
   End
   Begin VB.Image OUPride 
      Height          =   5370
      Left            =   5160
      Picture         =   "OU-OSU Screen Saver.frx":3722
      Top             =   3120
      Visible         =   0   'False
      Width           =   4440
   End
   Begin VB.Image OSUPregame2 
      Height          =   3600
      Left            =   600
      Picture         =   "OU-OSU Screen Saver.frx":B105
      Top             =   1560
      Visible         =   0   'False
      Width           =   7125
   End
   Begin VB.Image OSUPregame 
      Height          =   3600
      Left            =   840
      Picture         =   "OU-OSU Screen Saver.frx":27387
      Top             =   960
      Visible         =   0   'False
      Width           =   7125
   End
   Begin VB.Image Pete2 
      Height          =   3300
      Index           =   0
      Left            =   7200
      Picture         =   "OU-OSU Screen Saver.frx":43609
      Top             =   1440
      Visible         =   0   'False
      Width           =   2010
   End
   Begin VB.Image Pete 
      Height          =   3300
      Index           =   0
      Left            =   7560
      Picture         =   "OU-OSU Screen Saver.frx":44A91
      Top             =   600
      Visible         =   0   'False
      Width           =   2010
   End
   Begin VB.Image OUTitle 
      BorderStyle     =   1  'Fixed Single
      Height          =   945
      Left            =   3600
      Picture         =   "OU-OSU Screen Saver.frx":45F3C
      Top             =   120
      Visible         =   0   'False
      Width           =   5790
   End
   Begin VB.Image Title4 
      Height          =   2700
      Left            =   1080
      Picture         =   "OU-OSU Screen Saver.frx":4BBFE
      Top             =   2040
      Visible         =   0   'False
      Width           =   5250
   End
   Begin VB.Image Title3 
      Height          =   2700
      Left            =   -480
      Picture         =   "OU-OSU Screen Saver.frx":4C4CC
      Top             =   1560
      Visible         =   0   'False
      Width           =   7500
   End
   Begin VB.Image Title1 
      Height          =   2700
      Left            =   0
      Picture         =   "OU-OSU Screen Saver.frx":4D21F
      Top             =   0
      Visible         =   0   'False
      Width           =   7500
   End
End
Attribute VB_Name = "myForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim QuitFlag As Integer  'Flag to tell if terminating screen saver
Dim CursorCount As Integer  '

'Colors for background
Const OSUColor = &H80FF&
Const OUColor = &HC0&

'Windows API function to show and hide mouse
Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long


Private Sub RunIt()
    Dim WaitTime
    
    Do Until QuitFlag = True
        DoEvents  'Let windows process

        'Change background to OSU color
        myForm.BackColor = OSUColor
        
        'Align the OSU title
        Title1.Left = Me.Width / 2 - Title1.Width / 2
        Title1.Top = 100
        
        'Align cartoon pete, pregame, cartoon pete
        Pete(0).Left = Me.Width / 2 - (Pete(0).Width + 250 + OSUPregame.Width + 250 + Pete(1).Width) / 2
        Pete(0).Top = Title1.Top + Title1.Height + 250
        OSUPregame.Left = Pete(0).Left + Pete(0).Width + 250
        OSUPregame.Top = Pete(0).Top
        Pete(1).Left = OSUPregame.Left + OSUPregame.Width + 250
        Pete(1).Top = Pete(0).Top
        
        'Show off OSU
        Title1.Visible = True
        Pete(0).Visible = True
        OSUPregame.Visible = True
        Pete(1).Visible = True
    
        'Let's here that great OSU song, hehe
        sndPlayer.Command = "close"
        sndPlayer.DeviceType = "WaveAudio"
        sndPlayer.filename = App.Path + "\cowboys.wav"
        sndPlayer.Command = "Open"
        
        sndPlayer.Command = "Play"
        
        Do  'just loop until either the song is over or the screen saver is killed
            DoEvents

            If QuitFlag = True Or Not sndPlayer.Mode = mciModePlay Then Exit Do
        
            'Within last 4 secs of song, change OSU title to squeezed, not red yet
            If (sndPlayer.Position / 1000) > 25 And Title2.Visible = False Then
                Title1.Visible = False
                Title2.Left = Title1.Left
                Title2.Top = Title1.Top
                Title2.Visible = True
            End If
        Loop
    
        'Now for OU!!!
        sndPlayer.Command = "Close"
           
        'SHUT off OSU
        Title1.Visible = False
        Pete(0).Visible = False
        OSUPregame.Visible = False
        Pete(1).Visible = False
            
        
        'Lets give the nice OU stuff now
        'Change background color to dark red
        myForm.BackColor = OUColor
        
        'Align cartoon pete2, pregame2, cartoon pete2
        Pete2(0).Left = Pete(0).Left
        Pete2(0).Top = Pete(0).Top
        OSUPregame2.Left = OSUPregame.Left
        OSUPregame2.Top = OSUPregame.Top
        Pete2(1).Left = Pete(1).Left
        Pete2(1).Top = Pete(1).Top
        
        'Add the OU title
        OUTitle.Left = Me.Width / 2 - OUTitle.Width / 2
        OUTitle.Top = OSUPregame2.Top + OSUPregame2.Height + 250
        
        'Lets show the nice OU stuff now
        Pete2(0).Visible = True
        OSUPregame2.Visible = True
        Pete2(1).Visible = True
        OUTitle.Visible = True
        
        'BOOMER SOONER, BOOMER SOONER!!!
        sndPlayer.Command = "close"
        sndPlayer.DeviceType = "WaveAudio"
        sndPlayer.filename = App.Path + "\boomer.wav"
        sndPlayer.Command = "Open"
        
        sndPlayer.Command = "Play"
        
        Do  'just loop until either the song is over or the screen saver is killed
            DoEvents

            If QuitFlag = True Or Not sndPlayer.Mode = mciModePlay Then Exit Do
        
            'After 6 secs in song, OSU squeezed in red, after 12 secs, just OU
            If ((sndPlayer.Position / 1000) > 12) And Title4.Visible = False Then
                Title3.Visible = False
                Title4.Left = Me.Width / 2 - Title4.Width / 2
                Title4.Top = Title1.Top
                Title4.Visible = True
            ElseIf ((sndPlayer.Position / 1000) > 6) And Title3.Visible = False And Title4.Visible = False Then
                Title2.Visible = False
                Title3.Left = Title2.Left
                Title3.Top = Title2.Top
                Title3.Visible = True
            End If
        
            'At 25 seconds show the "Pride" image instead of OSU makeover
            If (sndPlayer.Position / 1000) > 37 Then
                'Back to the OSU makeover!!!
                Pete2(0).Visible = True
                OSUPregame2.Visible = True
                Pete2(1).Visible = True
                OUTitle.Visible = True
                
                OUPride.Visible = False
            ElseIf (sndPlayer.Position / 1000) > 24 Then
                Pete2(0).Visible = False
                OSUPregame2.Visible = False
                Pete2(1).Visible = False
                OUTitle.Visible = False
            
                'Show it!!!
                OUPride.Left = Me.Width / 2 - OUPride.Width / 2
                OUPride.Top = Title4.Top + Title4.Height + 250
                OUPride.Visible = True
            End If
        Loop
    
        'Still going, guess need to hide all OU stuff now, hehe
        Title4.Visible = False
        Pete2(0).Visible = False
        OSUPregame2.Visible = False
        Pete2(1).Visible = False
        OUTitle.Visible = False
    
        'Lets show off my name, hehe
        Rodney.Left = Me.Width / 2 - Rodney.Width / 2
        Rodney.Top = Me.Height / 2 - Rodney.Height / 2
        Rodney.Visible = True
    
        WaitTime = Timer
        'Show my little message
        Do Until QuitFlag = True Or (Timer - WaitTime) > 10
            DoEvents
        Loop
    
        Rodney.Visible = False
    Loop
End Sub

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
        Exit Sub  'hehe
    End If

    'Check for preview command
    If InStr(Command$, "/p") <> False Then
        'I haven't added the code to put the screen
        'saver in the little window under the display
        'properties but it can be done using Win API
        
        'Still not done!!!
        Unload Me
    End If

    Me.Show  'Show this form
    
    Call RunIt  'Start all the pretty screen saver stuff
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
'Check to see if screen saver needs to unload
If QuitFlag = True Then
    Unload myForm  'Unload the form
End If
End Sub
