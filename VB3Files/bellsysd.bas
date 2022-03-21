'Declares for WindowsDirectory
Declare Function GetWindowsDirectory Lib "Kernel" (ByVal P$, ByVal S%) As Integer
'Declares for ExitWindows
Declare Function ExitWindows% Lib "User" (ByVal dwReturnCode&, ByVal reserved%)

Global ByPass As Integer 'Declare flag used for ini file saving

Global VPath As String 'Variable used to store apps path

Global Passed As Integer 'Flag for bell timers

Global MinPassed As Integer 'Used to count minutes passed

Global Periods As Integer 'Used to store number of periods

'Flag used to see if a sound file is to be deleted
Global FDelete As Integer

'Used to show boxes in ShowHideBoxes sub for weekly schedule
Global BoxByPass As Integer

Global Shifting As Integer 'Flag for grid movement

Global WBTI(50) As String 'Array for storing weekly bell time information

'This is where the program branches off to the sub Form_Load

Sub Get_Period_Number ()
       'Get what user entered as period(s) number
       If InStr(frmMain!optTestDay.Caption, "6") Then
           Periods = 6 '6 periods
       ElseIf InStr(frmMain!optTestDay.Caption, "5") Then
           Periods = 5 '5 periods
       ElseIf InStr(frmMain!optTestDay.Caption, "4") Then
           Periods = 4 '4 periods
       ElseIf InStr(frmMain!optTestDay.Caption, "3") Then
           Periods = 3 '3 periods
       ElseIf InStr(frmMain!optTestDay.Caption, "2") Then
           Periods = 2 '2 periods
       ElseIf InStr(frmMain!optTestDay.Caption, "1") Then
           Periods = 1 '1 period
       End If
End Sub

Sub iniFile (OpenFile As Integer, SaveFile As Integer)
    Dim counter As Integer 'Used in counter
    Dim SecondBellRing As String 'Second bell ring time variable
    Dim UseWeekTimes As Integer 'Tell if using weekly set times
    Static DayTimes(10) As String 'Used to store normal day

    'This is for the .ini file to be used when program is started

    On Error GoTo ErrorHandler 'If error occurs go on to next line
    
    'Determine if user needs to open ini file or save it
    If OpenFile = True And SaveFile = False Then 'User needs to open it
       
       'Open Bellsys.ini file for input
       Open VPath + "Bellsys.ini" For Input As #1

       'Store file data in arrays and variable
       Input #1, SecondBellRing
       Input #1, UseWeekTimes
       For counter = 0 To 8
         Input #1, DayTimes(counter)
       Next counter
       For counter = 0 To 45
         Input #1, WBTI(counter)
       Next counter
       
       'Set up normal day times
       For counter = 0 To 8
         frmMain!txttimes(counter).Text = DayTimes(counter)
       Next counter

       'Set second bell ring time
       frmMain.lbltimering.Caption = SecondBellRing & " minute(s)"
       
       'Check to see if user used week time settings
       'If UseWeekTimes = True Then frmMain!chkWeekBellTime.Value = True
       
       Close 'Close file

    Else
       'Save file with bell times and second bell time
       
       'Open file for output
       Open VPath + "Bellsys.ini" For Output As #2

       'Write data to file using counter
       Write #2, Left(frmMain.lbltimering.Caption, 2)
       Write #2, frmMain!chkWeekBellTime.Value
       For counter = 0 To 8
         Write #2, frmMain.txttimes(counter).Text
       Next counter
       For counter = 0 To 45
         Write #2, WBTI(counter)
       Next counter
       
       Close 'Close file
    End If

ErrorHandler: 'Line label
    'Check for any errors
    If Err Then 'There is a error
       If Err = 53 Then 'Assume first time started
          'Use default settings by calling cmdsetdefult_click
          ByPass = True 'Set flag to true
          frmMain!cmdSetDefault.Value = True
          ByPass = False 'Set flag to false
       Else 'Tell user about error
          'Make beep and message show up
          Beep
          MsgBox "Error " & Err & "." + Chr$(13) + Error$ + ".", 16, "ini Error"
       End If
    End If

Exit Sub 'Leave sub
Resume 'Needed for error handler

End Sub

Sub PlaySounds (Sound As String, Doing As String)
    Dim MCI As MMControl 'Variable used for mciplayer object
    
    Set MCI = frmMain!MCIPlayer 'Set variable so it is mci object
    
    'Set up mci player device for use
    MCI.Wait = True 'Wait till finished playing
    MCI.Notify = False 'Do not notify when done
    MCI.UpdateInterval = 1000 'Update every second
    MCI.TimeFormat = 0  'Set time format
    MCI.DeviceType = "WaveAudio" 'Wave device
    
    If Doing = "Testing" Then 'Check to see if testing device
        MCI.Command = "Open" 'Open device
        'Check to see if device can record for intercom
        If MCI.CanRecord = True Then 'It can
           'Do nothing
        Else 'It can not
           frmMain.cmdintercom.Enabled = False 'Disable intercom button
        End If
    
        'Check to see if device can play
        If MCI.CanPlay = True Then 'It can
           'Do nothing
        Else 'It can not
           Beep 'Beep computer speaker
           MsgBox "Device can not play.", 16, "Warning" 'Tell user
           End 'Terminate program
        End If
    
        MCI.Command = "Close" 'Close device
        Exit Sub 'Leave sub
    End If

    'Determine what sound to play
    If Sound = "Bell" Then
       MCI.FileName = VPath + "Bellsond.wav" 'Set file
       MCI.Command = "Open" 'Open device
    ElseIf Sound = "Fire" Then
       MCI.Command = "Close" 'For priority reasons
       MCI.FileName = VPath + "FireBell.wav" 'Set file
       MCI.Command = "Open" 'Open device
    ElseIf Sound = "Storm" Then
       MCI.Command = "Close" 'For priority reasons
       MCI.FileName = VPath + "Stormbel.wav" 'Set file
       MCI.Command = "Open" 'Open device
    ElseIf Sound = "Record" Then
       MCI.FileName = VPath + "Intercom.wav" 'Set file
       MCI.Command = "Open" 'Open device
    End If

    'Check to if we are playing, recording, stoping, or saving a file and
    'then do we should be doing

    If Doing = "Playing" Then
       MCI.From = 0 'Start from beginning
       MCI.To = CDbl(MCI.Length) 'Play till end
       MCI.Command = "Play" 'Play sound
    ElseIf Doing = "Recording" Then
       MCI.From = 0 'Start at beginning
       MCI.Command = "Record" 'Start recording
       Exit Sub 'Leave sub
    ElseIf Doing = "Stoping" Then
       MCI.Command = "Stop" 'Stop recording
       Exit Sub 'Leave sub
    ElseIf Doing = "Saving" Then
       MCI.Command = "Save" 'Save file
       FDelete = True 'Set flag so can delete intercom.wav
       MCI.Command = "Close" 'Close device
       Exit Sub 'Leave sub
    End If

    'Need to pause so sound can finish before closing
    Do
       DoEvents 'Let windows do its stuff
    Loop Until MCI.Position = CDbl(MCI.Length)
    
    MCI.Command = "Close" 'Close device
    
    If FDelete = True Then 'Need to delete Intercom.wav
       Kill VPath + "Intercom.wav" 'Delete file
       FDelete = False 'Toggle flag
    End If
End Sub

Sub ShowHideBoxes ()
    Dim I As Integer 'Counter variable

    'Hide or show correct number of boxes for times

    'Check to see if option selected for test day
    If frmMain!optTestDay.Value = False Then
       If BoxByPass = False Then Exit Sub 'Leave sub
    End If
    
    'Reset boxes
    For I = 0 To 6
       frmMain!txttimes(I).Visible = False
    Next I

    'Show correct boxes
    For I = 0 To Periods - 1
        frmMain!txttimes(I).Visible = True
    Next I
End Sub

Sub spinperiod_SpinDown ()
    'Change number
    Periods = Periods - 1

    'Check number
    If Periods < 1 Then Periods = 6

    'Set caption of option button
    frmMain!optTestDay.Caption = "Test day with " & Periods & " period(s)"

    'Wait for windows
    DoEvents

    Call ShowHideBoxes 'Call sub
End Sub

Function WindowsDirectory () As String
Dim WinPath As String
    WinPath = String(145, Chr(0))
    WindowsDirectory = Left(WinPath, GetWindowsDirectory(WinPath, Len(WinPath)))
End Function

