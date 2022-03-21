VERSION 2.00
Begin Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bell System Demo Version"
   ClientHeight    =   4020
   ClientLeft      =   480
   ClientTop       =   1395
   ClientWidth     =   8940
   Height          =   4425
   Icon            =   BELLSYSD.FRX:0000
   KeyPreview      =   -1  'True
   Left            =   420
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   8940
   Top             =   1050
   Width           =   9060
   Begin SSCheck chkWeekBellTime 
      Caption         =   "Use weekly time schedule"
      Font3D          =   1  'Raised w/light shading
      Height          =   255
      Left            =   6360
      TabIndex        =   38
      Top             =   3000
      Width           =   2655
   End
   Begin CommonDialog CMD 
      Left            =   3240
      Top             =   5760
   End
   Begin SSCommand cmdhelp 
      BevelWidth      =   1
      Caption         =   "&Help"
      Font3D          =   1  'Raised w/light shading
      Height          =   495
      Left            =   4080
      TabIndex        =   15
      Top             =   4200
      Width           =   1815
   End
   Begin Timer tmrTime 
      Interval        =   500
      Left            =   1320
      Top             =   5760
   End
   Begin Timer tmrPlayBellSound 
      Interval        =   500
      Left            =   1800
      Top             =   5760
   End
   Begin TextBox txttimes 
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   1
      Left            =   2520
      MaxLength       =   11
      TabIndex        =   1
      Top             =   1200
      Width           =   1215
   End
   Begin Timer tmrMinuteWait 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   2280
      Top             =   5760
   End
   Begin SSCommand cmdintercom 
      BevelWidth      =   1
      Caption         =   "&Intercom"
      Font3D          =   1  'Raised w/light shading
      Height          =   495
      Left            =   4080
      TabIndex        =   12
      Top             =   2400
      Width           =   1815
   End
   Begin TextBox txttimes 
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   8
      Left            =   2520
      MaxLength       =   11
      TabIndex        =   8
      Top             =   5400
      Width           =   1215
   End
   Begin SSPanel pnllblRing 
      Alignment       =   4  'Right Justify - MIDDLE
      Caption         =   "Last Bell:"
      FloodShowPct    =   0   'False
      Font3D          =   1  'Raised w/light shading
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   9.75
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   255
      Index           =   8
      Left            =   360
      TabIndex        =   35
      Top             =   5400
      Width           =   2025
   End
   Begin TextBox txttimes 
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   7
      Left            =   2520
      MaxLength       =   11
      TabIndex        =   7
      Top             =   4800
      Width           =   1215
   End
   Begin SSPanel pnllblRing 
      Alignment       =   4  'Right Justify - MIDDLE
      Caption         =   "Lunch Bell:"
      FloodShowPct    =   0   'False
      Font3D          =   1  'Raised w/light shading
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   9.75
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   255
      Index           =   7
      Left            =   360
      TabIndex        =   34
      Top             =   4800
      Width           =   2025
   End
   Begin SSPanel pnllblRing 
      Alignment       =   4  'Right Justify - MIDDLE
      Caption         =   "Seventh Bell:"
      FloodShowPct    =   0   'False
      Font3D          =   1  'Raised w/light shading
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   9.75
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   255
      Index           =   6
      Left            =   360
      TabIndex        =   33
      Top             =   4200
      Width           =   2025
   End
   Begin SSPanel pnllblRing 
      Alignment       =   4  'Right Justify - MIDDLE
      Caption         =   "Sixth Bell:"
      FloodShowPct    =   0   'False
      Font3D          =   1  'Raised w/light shading
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   9.75
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   255
      Index           =   5
      Left            =   360
      TabIndex        =   32
      Top             =   3600
      Width           =   2025
   End
   Begin SSPanel pnllblRing 
      Alignment       =   4  'Right Justify - MIDDLE
      Caption         =   "Fifth Bell:"
      FloodShowPct    =   0   'False
      Font3D          =   1  'Raised w/light shading
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   9.75
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   255
      Index           =   4
      Left            =   360
      TabIndex        =   31
      Top             =   3000
      Width           =   2025
   End
   Begin SSPanel pnllblRing 
      Alignment       =   4  'Right Justify - MIDDLE
      Caption         =   "Fourth Bell:"
      FloodShowPct    =   0   'False
      Font3D          =   1  'Raised w/light shading
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   9.75
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   255
      Index           =   3
      Left            =   360
      TabIndex        =   30
      Top             =   2400
      Width           =   2025
   End
   Begin SSPanel pnllblRing 
      Alignment       =   4  'Right Justify - MIDDLE
      Caption         =   "Third Bell:"
      FloodShowPct    =   0   'False
      Font3D          =   1  'Raised w/light shading
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   9.75
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   29
      Top             =   1800
      Width           =   2025
   End
   Begin SSPanel pnllblRing 
      Alignment       =   4  'Right Justify - MIDDLE
      Caption         =   "Second Bell:"
      FloodShowPct    =   0   'False
      Font3D          =   1  'Raised w/light shading
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   9.75
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   28
      Top             =   1200
      Width           =   2025
   End
   Begin MMControl MCIPlayer 
      DeviceType      =   "WaveAudio"
      Height          =   495
      Left            =   360
      RecordMode      =   1  'Overwrite
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   6240
      UpdateInterval  =   1
      Visible         =   0   'False
      Width           =   8670
   End
   Begin SSCommand cmdAbout 
      BevelWidth      =   1
      Caption         =   "&About"
      Font3D          =   1  'Raised w/light shading
      Height          =   495
      Left            =   4080
      TabIndex        =   16
      Top             =   4800
      Width           =   1815
   End
   Begin SSCommand cmdSetDefault 
      BevelWidth      =   1
      Caption         =   "Set To &Defaults"
      Font3D          =   1  'Raised w/light shading
      Height          =   495
      Left            =   4080
      TabIndex        =   14
      Top             =   3600
      Width           =   1815
   End
   Begin SSCommand cmdConPan 
      BevelWidth      =   1
      Caption         =   "Start &Control Panel"
      Font3D          =   1  'Raised w/light shading
      Height          =   495
      Left            =   4080
      TabIndex        =   13
      Top             =   3000
      Width           =   1815
   End
   Begin SSCommand cmdshutdown 
      AutoSize        =   1  'Adjust Picture Size To Button
      BevelWidth      =   1
      Caption         =   "Sh&ut Down System"
      Font3D          =   1  'Raised w/light shading
      Height          =   495
      Left            =   4080
      TabIndex        =   17
      Top             =   5400
      Width           =   1815
   End
   Begin SSCommand cmdStorm 
      BevelWidth      =   1
      Caption         =   "&Storm Bell"
      Font3D          =   1  'Raised w/light shading
      Height          =   495
      Left            =   4080
      TabIndex        =   11
      Top             =   1800
      Width           =   1815
   End
   Begin SSCommand cmdFire 
      BevelWidth      =   1
      Caption         =   "&Fire Bell"
      Font3D          =   1  'Raised w/light shading
      Height          =   495
      Left            =   4080
      TabIndex        =   10
      Top             =   1200
      Width           =   1815
   End
   Begin SSCommand cmdSetTime 
      BevelWidth      =   1
      Caption         =   "Set &Time"
      Font3D          =   1  'Raised w/light shading
      Height          =   495
      Left            =   4080
      TabIndex        =   9
      Top             =   600
      Width           =   1815
   End
   Begin SSFrame framebell 
      Caption         =   "Ring minute bell after"
      Font3D          =   3  'Inset w/light shading
      Height          =   1815
      Left            =   6240
      TabIndex        =   24
      Top             =   4080
      Width           =   2175
      Begin SpinButton spintime 
         BackColor       =   &H00808080&
         Height          =   975
         Left            =   120
         SpinBackColor   =   &H00808080&
         Top             =   720
         Width           =   1935
      End
      Begin Label lbltimering 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "5 minute(s)"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   12
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   300
         Left            =   120
         TabIndex        =   25
         Top             =   360
         Width           =   1365
      End
   End
   Begin SSPanel pnllblRing 
      Alignment       =   4  'Right Justify - MIDDLE
      Caption         =   "First Bell:"
      FloodShowPct    =   0   'False
      Font3D          =   1  'Raised w/light shading
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   9.75
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   23
      Top             =   600
      Width           =   2025
   End
   Begin TextBox txttimes 
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   6
      Left            =   2520
      MaxLength       =   11
      TabIndex        =   6
      Top             =   4200
      Width           =   1215
   End
   Begin TextBox txttimes 
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   5
      Left            =   2520
      MaxLength       =   11
      TabIndex        =   5
      Top             =   3600
      Width           =   1215
   End
   Begin TextBox txttimes 
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   4
      Left            =   2520
      MaxLength       =   11
      TabIndex        =   4
      Top             =   3000
      Width           =   1215
   End
   Begin TextBox txttimes 
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   3
      Left            =   2520
      MaxLength       =   11
      TabIndex        =   3
      Top             =   2400
      Width           =   1215
   End
   Begin TextBox txttimes 
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   2
      Left            =   2520
      MaxLength       =   11
      TabIndex        =   2
      Top             =   1800
      Width           =   1215
   End
   Begin TextBox txttimes 
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   0
      Left            =   2520
      MaxLength       =   11
      TabIndex        =   0
      Top             =   600
      Width           =   1215
   End
   Begin SSOption optnormal 
      Caption         =   "Normal day with 7 periods"
      Font3D          =   1  'Raised w/light shading
      Height          =   375
      Left            =   6360
      TabIndex        =   18
      Top             =   840
      Value           =   -1  'True
      Width           =   2535
   End
   Begin SSOption optTestDay 
      Caption         =   "Test day with 4 period(s)"
      Font3D          =   1  'Raised w/light shading
      Height          =   375
      Left            =   6360
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   1320
      Width           =   2415
   End
   Begin SSFrame Frameopt 
      Caption         =   "Ring for"
      Font3D          =   3  'Inset w/light shading
      Height          =   3255
      Left            =   6240
      TabIndex        =   22
      Top             =   600
      Width           =   3015
      Begin SSCommand cmdSetWeekTimes 
         BevelWidth      =   1
         Caption         =   "Set &Weekly Times"
         Enabled         =   0   'False
         Font3D          =   1  'Raised w/light shading
         Height          =   375
         Left            =   120
         TabIndex        =   39
         Top             =   2760
         Width           =   2535
      End
      Begin SpinButton spinperiod 
         BackColor       =   &H00808080&
         Height          =   375
         Left            =   2640
         SpinBackColor   =   &H00808080&
         Top             =   720
         Width           =   255
      End
      Begin SSCheck chklast 
         Caption         =   "Last bell"
         Font3D          =   1  'Raised w/light shading
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   1920
         Value           =   -1  'True
         Width           =   1455
      End
      Begin SSCheck chklunch 
         Caption         =   "Lunch bell"
         Font3D          =   1  'Raised w/light shading
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   1560
         Value           =   -1  'True
         Width           =   2175
      End
      Begin Line Line1 
         X1              =   120
         X2              =   2520
         Y1              =   2280
         Y2              =   2280
      End
      Begin Label lblinstruct 
         BackStyle       =   0  'Transparent
         Caption         =   "Click spinner to change period number"
         Height          =   495
         Left            =   120
         TabIndex        =   37
         Top             =   1080
         Width           =   2775
         WordWrap        =   -1  'True
      End
   End
   Begin Label lblDummyTime 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "lblDummyTime"
      Height          =   195
      Left            =   4080
      TabIndex        =   36
      Top             =   6000
      Visible         =   0   'False
      Width           =   1215
   End
   Begin Label lblTime 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "The current time is"
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   18
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   435
      Left            =   360
      TabIndex        =   26
      Top             =   0
      Width           =   3285
   End
End

Sub chklast_Click (Value As Integer)
    If chkLast.Value = True Then
       txttimes(8).Visible = True 'Show last bell time box
    Else
       txttimes(8).Visible = False 'Hide last bell time box
    End If
End Sub

Sub chklunch_Click (Value As Integer)
    If chkLunch.Value = True Then
       txttimes(7).Visible = True 'Show lunch bell time box
    Else
       txttimes(7).Visible = False 'Hide lunch bell time box
    End If
End Sub

Sub chkWeekBellTime_Click (Value As Integer)

MsgBox "Register" & Chr$(13) & Chr$(10) & "See the BellSysD.txt file for more information!", 48, "Feature disabled"

'Disabled in demo version
'    Dim OldPeriods As Integer 'Variable for storing
'
    'Check to see if using weekly time schedule
'    If Value = True Then 'User is using weekly schedule
'       cmdSetWeekTimes.Enabled = True 'Enable command button
'       tmrPlayBellSound.Enabled = False 'Disable normal timer
'       tmrWeekBellTime.Enabled = True 'Enable weekly timer
       'Hide boxes
'       For I = 0 To 8
'          frmMain!txttimes(I).Visible = False
'       Next I
       'Set lunch and last bell checkmarks
'       chkLunch.Value = False
'       chkLast.Value = False
       'Disable Ring for controls
'       optNormal.Enabled = False 'Normal day option
'       optTestDay.Enabled = False 'Test day option
'       SpinPeriod.Enabled = False 'Spinner for period number
'       lblInstruct.Enabled = False 'Instructions to spinner
'       chkLunch.Enabled = False 'Check box for lunch bell
'       chkLast.Enabled = False 'Check box for last bell
'    Else
'       OldPeriods = Periods 'Store old number
       'Check to see if using weekly time schedule
'       If Value = False Then 'User is not using weekly schedule
'          cmdSetWeekTimes.Enabled = False 'Disable command button
'          tmrPlayBellSound.Enabled = True 'Enable non-weekly timer
'          tmrWeekBellTime.Enabled = False 'Disable weekly timer
'          If optNormal.Value = True Then 'Check to see if normal day was used
'             Periods = 7 'Set to show all periods
'             BoxByPass = True 'Bypass normal checking in sub
'             Call ShowHideBoxes 'Call sub to change periods
'             Periods = OldPeriods 'Reset period number
'             BoxByPass = False 'Turn off bypass flag
'          Else
'             Call ShowHideBoxes 'Call sub to change periods
'          End If
          'Set lunch and last bell checkmarks
'          chkLunch.Value = True
'          chkLast.Value = True
          'Enable Ring For controls
'          optNormal.Enabled = True 'Normal day option
'          optTestDay.Enabled = True 'Test day option
'          SpinPeriod.Enabled = True 'Spinner for period number
'          lblInstruct.Enabled = True 'Instructions to spinner
'          chkLunch.Enabled = True 'Check box for lunch bell
'          chkLast.Enabled = True 'Check box for last bell
'       End If
'    End If
End Sub

Sub cmdAbout_Click ()
    Dim Msg As String 'Variable for message
    Dim CTRL As String 'Variable that stands for character line feed

    CTRL = Chr$(13) + Chr$(10) 'Set up character line feed

    'Set up message
    Msg = "Bell System Demo Version" + CTRL
    Msg = Msg + "Programed by Rodney Beede." + CTRL
    Msg = Msg + "Published by Infinisoft." + CTRL
    Msg = Msg + "E-mail me at rodney_beede@hotmail.com" + CTRL
    Msg = Msg + "Read the BellSys.txt file for more information."
    
    'Show message box
    MsgBox Msg, 64, "About"

End Sub

Sub cmdConPan_Click ()
    Dim ShellDummy As Integer 'Used to start shell
    Dim WinPath As String 'Variable for windows path
    
    'We need to start control panel and program will assume it is in windows
    'directory where it is usually installed

    WinPath = WindowsDirectory() 'Get windows directory
    
    ShellDummy = Shell(WinPath + "\control.exe", 1)'Start control panel with normal size
End Sub

Sub cmdFire_Click ()
    Call PlaySounds("Fire", "Playing") 'Go to sub to play sound
End Sub

Sub cmdhelp_Click ()
    'Demo verson tells user to register
    Beep
    MsgBox "See the Bellsysd.txt for more information", 48, "Register this program!"
    
    
    'Bypass errors
    On Error Resume Next
    
    'Get programs path
    If Len(App.Path) > 3 Then
       VPath = App.Path + "\" 'Add \ character to path
    Else
       VPath = App.Path 'Set path
    End If

    'Check to see if help already started
    AppActivate "Bell System Help"
    
    'Check for error saying help is not already started
    If Err = 5 Then
       'Need to start help
       'Use Common Dialog Box to start help
       CMD.HelpFile = VPath + "BellSys.HLP"
       CMD.HelpCommand = &H3
       CMD.Action = 6
    End If
End Sub

Sub cmdintercom_Click ()
MsgBox "Register" & Chr$(13) & Chr$(10) & "See the BellSysD.txt file for more information!", 48, "Feature disabled"

'This is a demo and so this is not enabled
'    Me.Enabled = False 'Disable form
'    Load frmInterCom 'Make sure it loads
'    frmInterCom.Show 'Show intercom form
End Sub

Sub cmdSetDefault_Click ()
   Dim Response As Integer 'Response variable
   Dim Msg As String 'Message variable

   'May be making ini file so determine if ByPass variable = true
   If ByPass = True Then GoTo ByPassed 'Is true skip message and set to default
   
   'Set up message
   Msg = "Setting to default settings will erase current!" + Chr$(13)
   
   'Make message box show up
   Response = MsgBox(Msg, 17, "Warning")

   'Determine response of user
   If Response = 2 Then Exit Sub ' User canceled

ByPassed: 'Line label
   'If user did not cancel then set times, options, minutes, and checkmarks
   'Start with times
   txttimes(0).Text = "8:20:00 AM" 'First period
   txttimes(1).Text = "9:15:00 AM" 'Second period
   txttimes(2).Text = "10:10:00 AM" 'Third period
   txttimes(3).Text = "11:05:00 AM" 'Forth period
   txttimes(7).Text = "12:00:00 AM" 'Lunch
   txttimes(4).Text = "12:30:00 PM" 'Fifth period
   txttimes(5).Text = "1:25:00 PM" 'Sixth period
   txttimes(6).Text = "2:20:00 PM" 'Seventh period
   txttimes(8).Text = "3:10:00 PM" 'Last bell

   optNormal.Value = True 'Set options to Normal Day

   chkLast.Value = True 'Put a checkmark in the Last Bell box
   chkLunch.Value = True 'Put a checkmark in the Lunch Bell box
   chkWeekBellTime.Value = False 'Erase checkmark in Week box

   'Set minutes till second bell rings
   lblTimeRing.Caption = "5 minute(s)"

End Sub

Sub cmdSetTime_Click ()
    Dim TimeEntered As String 'Variable for Time Entered

    'Ask for user to type in time
    TimeEntered = InputBox$("Type in a new time.", "Set Time", Time)

    'Check out time entered
    If TimeEntered = "" Then Exit Sub 'User Canceled

    If IsDate(TimeEntered) Then 'Valid time entered
       Time = TimeEntered  'Set time to time entered
    Else
       'Not valid time tell user
       MsgBox "Time not valid.", 16, "Warning"
    End If
End Sub

Sub cmdSetWeekTimes_Click ()
'Disabled in demo version
'  frmMain.Enabled = False 'Disable this form
'  frmWeekBellTime.Caption = "Set Weekly Times" 'Reset caption on form
'  frmWeekBellTime.Enabled = True 'Make sure enabled
'  Load frmWeekBellTime 'Load form for getting bell times
'  frmWeekBellTime.Show 'Show form
'  tmrWeekBellTime.Enabled = False 'Disable timer
End Sub

Sub cmdshutdown_Click ()
Dim Response As Integer 'Response variable from users answer

'Show message and wait for user response
Response = MsgBox("Warning! About to shut down system.", 17, "Warning")

If Response = 2 Then Exit Sub 'User canceled

Dim Dummy As Integer 'Dummy variable for function call
    Dummy = ExitWindows(0, 0) 'Close windows

End Sub

Sub cmdStorm_Click ()
Call PlaySounds("Storm", "Playing") 'Go to sub to play sound
End Sub

Sub Form_Activate ()
       'Get what user entered as period(s) number for later use
       If InStr(optTestDay.Caption, "6") Then
           Periods = 6 '6 periods
       ElseIf InStr(optTestDay.Caption, "5") Then
           Periods = 5 '5 periods
       ElseIf InStr(optTestDay.Caption, "4") Then
           Periods = 4 '4 periods
       ElseIf InStr(optTestDay.Caption, "3") Then
           Periods = 3 '3 periods
       ElseIf InStr(optTestDay.Caption, "2") Then
           Periods = 2 '2 periods
       ElseIf InStr(optTestDay.Caption, "1") Then
           Periods = 1 '1 period
       End If
End Sub

Sub Form_Load ()
    
    'This is the first thing to be processed by the program start up
    
    'Set up form position, width, and height
    
    Me.Left = 0 'Put left of form to left of screen
    
    Me.Top = 0  'Put top of form to top of screen
    
    Me.Height = Screen.Height 'Make form's height same as screen
    
    Me.Width = Screen.Width   'Make form's width same as screen

    'Determine if \ character needs to be added to path for later use
    If Len(App.Path) > 3 Then 'It does give it path plus \ character
       VPath = App.Path + "\"
    Else
       VPath = App.Path 'It does not just give it path
    End If
    
    ByPass = False 'Set flag to false

    Call iniFile(True, False) 'Open ini file for input
    
    Call PlaySounds("Dummy", "Testing") 'Test mci device
End Sub

Sub Form_Unload (Cancel As Integer)
   Call iniFile(False, True) 'Save ini file
   MCIPlayer.Command = "Close" 'Close MCI Player
   End 'Make sure closes
End Sub

Sub optNormal_Click (Value As Integer)
    Dim Number As Integer 'Used for counter
    
    'Check to see if selected
    If Value Then 'It is selected
       'Make all of the time boxes appear
       For Number = 0 To 8
           txttimes(Number).Visible = True
       Next Number
    End If

End Sub

Sub optTestDay_Click (Value As Integer)
    Call ShowHideBoxes 'Call sub
End Sub

Sub spinperiod_SpinDown ()
    'Change number
    Periods = Periods - 1

    'Check number
    If Periods < 1 Then Periods = 6

    'Set caption of option button
    optTestDay.Caption = "Test day with " & Periods & " period(s)"

    'Wait for windows
    DoEvents

    Call ShowHideBoxes 'Call sub
End Sub

Sub spinperiod_SpinUp ()
    'Change number
    Periods = Periods + 1

    'Check number
    If Periods > 6 Then Periods = 1

    'Set caption of option button
    optTestDay.Caption = "Test day with " & Periods & " period(s)"

    'Wait for windows
    DoEvents

    Call ShowHideBoxes 'Call sub
End Sub

Sub spintime_SpinDown ()
    Dim Number As Integer ' Declare variable number for current times to ring
    
    'Get number of times to ring
    Number = Left(lblTimeRing.Caption, 2)
    
    'Change number and check it
    Number = Number - 1
    If Number < 1 Then Number = 10
    
    'Change labels caption when spun down
    lblTimeRing.Caption = Number & " minute(s)"
    
    'Make sure label changes
    DoEvents

End Sub

Sub spintime_SpinUp ()
    Dim Number As Integer ' Declare variable number for current times to ring
    
    'Get number of times to ring
    Number = Left(lblTimeRing.Caption, 2)
    
    'Change number and check it
    Number = Number + 1
    If Number > 10 Then Number = 1
    
    'Change labels caption when spun down
    lblTimeRing.Caption = Number & " minute(s)"
    
    'Make sure label changes
    DoEvents

End Sub

Sub tmrMinuteWait_Timer ()
    'Store number
    Wait = Left(lblTimeRing, 2)
    
    'Add number by one
    MinPassed = MinPassed + 1
    
    'Check to see if amount of time for minute bell is up
    If MinPassed >= Wait Then 'It is time to play sound
       Call PlaySounds("Bell", "Playing") 'Play sound
       SpinTime.Enabled = True 'Enable spinner
       tmrMinuteWait.Enabled = False 'Disable this timer
       tmrPlayBellSound.Enabled = True 'Enable non-weekly timer
       MinPassed = 0 'Reset number of minutes
    End If
End Sub

Sub tmrPlayBellSound_Timer ()
    Dim counter As Integer 'Used in counter
    
    'Set time in dummy label
    lblDummyTime.Caption = Format$(Time, "h:mm:ss AM/PM")

    'Determine if it is time to sound bell
    'Make sure it is a normal day
    If optNormal.Value = True Then 'It is so check time now
       For counter = 0 To 6 'Start counter
           If txttimes(counter).Text = lblDummyTime.Caption Then 'It is time
              tmrMinuteWait.Enabled = True 'Enable timer
              Call PlaySounds("Bell", "Playing") 'Go to sub to play sound
              GoTo MinuteBell 'Goto line labeled MinuteBell
           End If
       Next counter
    ElseIf optTestDay.Value = True Then 'Test day option one was selected
       Call Get_Period_Number 'Get period number
       
       'Start counter
       For counter = 0 To Periods - 1
           'Check to see if it is time yet
           If txttimes(counter).Text = lblDummyTime.Caption Then 'It is time
              tmrMinuteWait.Enabled = True 'Enable timer
              Call PlaySounds("Bell", "Playing") 'Go to sub to play sound
              GoTo MinuteBell 'Goto line labeled MinuteBell:
           End If
       Next counter
    End If

    'Check to see if it is time to play lunch or last bell
    If chkLunch.Value = True Then 'Lunch bell is to be played
       If txttimes(7).Text = lblDummyTime Then 'It is time
          tmrMinuteWait.Enabled = True 'Enable timer
          Call PlaySounds("Bell", "Playing") 'Go to sub to play sound
          GoTo MinuteBell 'Goto line labeled MinuteBell:
       End If
    End If
    If chkLast.Value = True Then 'Last bell is to be played
       If txttimes(8).Text = lblDummyTime Then 'It is time
          tmrMinuteWait.Enabled = True 'Enable timer
          Call PlaySounds("Bell", "Playing") 'Go to sub to play sound
          GoTo MinuteBell 'Goto line labeled MinuteBell:
       End If
    End If
    
    Exit Sub 'Exits sub
MinuteBell: 'Line label
    tmrMinuteWait.Enabled = True 'Enable timer
    SpinTime.Enabled = False 'Disable spin button for minutes
    tmrPlayBellSound.Enabled = False 'Disable timer
End Sub

Sub tmrTime_Timer ()
    'Set date and time in label
    lblTime.Caption = "The current time is " + Time + "."
End Sub

Sub txttimes_LostFocus (Index As Integer)
    'To bypass possible errors
    On Error Resume Next
    
    'Set to normal time format in case entered time is in army time
    txttimes(Index).Text = TimeValue(txttimes(Index).Text)
    'Check out time entered
    If IsDate(txttimes(Index).Text) Then 'Valid time entered
       'Add seconds part to time to ring
       txttimes(Index).Text = Format$(txttimes(Index).Text, "h:mm:ss AM/PM")
    Else
       'Not valid time tell user
       MsgBox "Time not valid.", 16, "Warning"
       'Highlight and give box focus back
       txttimes(Index).SetFocus
       txttimes(Index).SelStart = 0
       txttimes(Index).SelLength = Len(txttimes(Index).Text)
    End If
End Sub

