VERSION 2.00
Begin Form frmMain 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Windows Shell"
   ClientHeight    =   4020
   ClientLeft      =   1095
   ClientTop       =   1485
   ClientWidth     =   7365
   ForeColor       =   &H00000000&
   Height          =   4425
   Icon            =   WINSHEL1.FRX:0000
   Left            =   1035
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   7365
   Top             =   1140
   Width           =   7485
   Begin Frame fraPrograms 
      BackColor       =   &H00C0C0C0&
      Caption         =   "SELECT A PROGRAM"
      Height          =   2055
      Left            =   720
      TabIndex        =   3
      Top             =   1440
      Width           =   4575
      Begin CommandButton cmdProgram 
         Height          =   495
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
   End
   Begin CommandButton cmdProgramManager 
      Caption         =   "Program Manager"
      Height          =   495
      Left            =   4080
      TabIndex        =   2
      Top             =   0
      Width           =   2055
   End
   Begin CommandButton cmdDeleteProgram 
      Caption         =   "Delete Program"
      Height          =   495
      Left            =   2040
      TabIndex        =   1
      Top             =   0
      Width           =   2055
   End
   Begin CommandButton cmdAddProgram 
      Caption         =   "Add Program"
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2055
   End
End
Option Explicit

Dim Programs() As String  'Array to hold program paths
Dim ProgramNames() As String  'Array to hold program names

Dim Deleting As Integer  'Flag to tell if deleting button

Sub CleanArray ()
    Dim CountA As Integer  'For counter
    Dim CurrArrayCount As Integer  'For place in array
    Dim TmpArray() As String  'Holds copy of a array

    'Clean out Programs array first
    ReDim TmpArray(UBound(Programs))  'Size array
    
    'Make copy of array
    For CountA = 0 To UBound(Programs)
        TmpArray(CountA) = Programs(CountA)
    Next CountA
         
    CurrArrayCount = 1  'Set first position

    Erase Programs  'Clear out array
    ReDim Programs(0)  'Setup array
    
    'Refill in original array compacted
    For CountA = 1 To UBound(TmpArray)
        'Check to see if their is a blank spot
        If Not TmpArray(CountA) = "" Then
            'Their isn't, fill in data
            ReDim Preserve Programs(CurrArrayCount)  'Resize array
            
            Programs(CurrArrayCount) = TmpArray(CountA)
            CurrArrayCount = CurrArrayCount + 1  'Increment position
        End If
    Next CountA
    
    'Clean out ProgramNames array
    ReDim TmpArray(UBound(ProgramNames))  'Size array
    
    'Make copy of array
    For CountA = 0 To UBound(ProgramNames)
        TmpArray(CountA) = ProgramNames(CountA)
    Next CountA

    CurrArrayCount = 1  'Set first position

    Erase ProgramNames  'Clear out array
    ReDim ProgramNames(0)  'Setup array
    
    'Refill in original array compacted
    For CountA = 1 To UBound(TmpArray)
        'Check to see if their is a blank spot
        If Not TmpArray(CountA) = "" Then
            'Their isn't, fill in data
            ReDim Preserve ProgramNames(CurrArrayCount)  'Resize array
            ProgramNames(CurrArrayCount) = TmpArray(CountA)
            CurrArrayCount = CurrArrayCount + 1  'Increment position
        End If
    Next CountA
End Sub

Sub cmdAddProgram_Click ()
    Dim ProgramPath As String  'Holds program location
    Dim ProgramName As String  'Holds program name

    'Do a security check
    frmSecurity.Show 1

    'Check if password was valid
    If Passed = False Then  'It wasn't
        Exit Sub  'Leave
    Else
        Passed = False  'Clear flag and go on
    End If
    
    'Ask user for program path
    ProgramPath = InputBox$("Enter the path and filename of the program.", "Program Location")
    
    'Check to see if user canceled
    If ProgramPath = "" Then Exit Sub  'User did

    'Ask user for program name
    ProgramName = InputBox$("Enter the name for the program.", "Program Name")
    
    'Check to see if user canceled
    If ProgramName = "" Then Exit Sub  'User did

    'Resize arrays
    ReDim Preserve Programs(UBound(Programs) + 1)
    ReDim Preserve ProgramNames(UBound(ProgramNames) + 1)

    'Add data to array
    Programs(UBound(Programs)) = ProgramPath
    ProgramNames(UBound(ProgramNames)) = ProgramName

    'Clean up arrays
    Call CleanArray
    
    'Setup the buttons
    Call SetupButtons
End Sub

Sub cmdDeleteProgram_Click ()
    'Determine if deleting
    If Deleting = True Then
        'User canceled delete, tell user
        MsgBox "Delete was canceled.", MB_ICONINFORMATION, "Delete Program Canceled"
        
        'Enable appriopate buttons
        cmdAddProgram.Enabled = True
        cmdProgramManager.Enabled = True
    Else  'Tell user how to delete
        'Do a security check
        frmSecurity.Show 1
    
        'Check if password was valid
        If Passed = False Then  'It wasn't
            Exit Sub  'Leave
        Else
            Passed = False  'Clear flag and go on
        End If

        MsgBox "Click on the button you wish to delete.  To cancel click on the DELETE button again.", MB_ICONINFORMATION, "Delete Program"
    
        'Disable appriopate buttons
        cmdAddProgram.Enabled = False
        cmdProgramManager.Enabled = False
    End If

    Deleting = Not Deleting  'Toggle flag
End Sub

Sub cmdProgram_Click (Index As Integer)
    Dim appID As Integer  'For holding shell id
    Dim i As Integer  'For counter
    
    On Error Resume Next  'Skip over any errors
    
    'Check to see if this program needs to be deleted
    If Deleting = True Then  'Delete this program
        'Clear out entries in arrays
        Programs(Index + 1) = ""
        ProgramNames(Index + 1) = ""
    
        'Clean up arrays
        Call CleanArray

        Call SetupButtons  'Reset buttons

        Deleting = Not Deleting  'Toggle flag
        
        'Enable appriopate buttons
        cmdAddProgram.Enabled = True
        cmdProgramManager.Enabled = True
        
        Exit Sub  'Leave
    End If  'End of delete section

    'Strip out the directory and set it as the working directory
    For i = Len(Programs(Index + 1)) To 1 Step -1
        If Mid$(Programs(Index + 1), i, 1) = "\" Then
            'Found end of execution name, everything else is directory
            'so set the current working directory
            'Check to see if new working directory is the root of drive
            If InStr(Left$(Programs(Index + 1), i - 1), "\") Then 'Not root of drive
                ChDir Left$(Programs(Index + 1), i - 1)  'Change to directory
            Else  'Root of drive  ("C:" is the path looked at)
                ChDir Left$(Programs(Index + 1), i)'Change to root
            End If
            
            Exit For  'Leave loop
        End If
    Next i
  
    'Run the program
    appID = Shell(Programs(Index + 1), 1)
    
    If Err <> 0 Then  'Error occured, state error
        MsgBox Error(Err), MB_ICONEXCLAMATION, "Error number " + Str$(Err)
    End If
End Sub

Sub cmdProgramManager_Click ()
    Dim appID As Integer  'For holding app ID when ran

    'Do a security check
    frmSecurity.Show 1

    'Check if password was valid
    If Passed = False Then  'It wasn't
        Exit Sub  'Leave
    Else
        Passed = False  'Clear flag and go on
    End If
    
    'Run program manager
    appID = Shell("c:\windows\progman.exe", 1)
End Sub

Sub Form_Load ()
    'The Programs array goes from 1 - UBOUND(Programs)
    'The ProgramNames array goes from 1 - UBOUND(ProgramNames)
    'The cmdProgram array goes from 0 - (UBOUND(Programs) - 1)
    'UBOUND(Programs) == UBOUND(ProgramNames) at all times
    
    Dim returnVal As Integer  'Holds return value of API call
    Dim pOld  'Just a dummy variable for API call

    ReDim Programs(0)  'Setup array
    ReDim ProgramNames(0)  'Setup array
    
    'Size form as big as screen and put at top
    Me.Height = Screen.Height
    Me.Width = Screen.Width
    Me.Top = 0
    Me.Left = 0

    'Size the control bar
    cmdAddProgram.Left = 0
    cmdAddProgram.Width = Me.Width / 3
    cmdDeleteProgram.Left = cmdAddProgram.Left + cmdAddProgram.Width
    cmdDeleteProgram.Width = Me.Width / 3
    cmdProgramManager.Left = cmdDeleteProgram.Left + cmdDeleteProgram.Width
    cmdProgramManager.Width = Me.Width / 3

    'Position the frame
    fraPrograms.Left = 0
    fraPrograms.Top = cmdAddProgram.Top + cmdAddProgram.Height + 250
    fraPrograms.Height = Me.Height - fraPrograms.Top - 310
    fraPrograms.Width = Me.Width - 20

    Call iniFile("Read")    'Read information from inilization file

    Call SetupButtons  'Setup buttons on screen

    frmProgramManager.Show  'Show program manager form
End Sub

Sub Form_Resize ()
    'Check to see if the form is visible and not minimized
    'if not then still loading program
    If Me.Visible = True And Me.WindowState <> 1 Then  'Need to put back at top
        Me.Top = 0
        Me.Left = 0
    End If
End Sub

Sub Form_Unload (Cancel As Integer)
    Dim Response As Integer  'To hold user response

    'Ask user if they want to terminate Windows
    Response = MsgBox("Continuing will end your Windows session.", MB_OKCANCEL + MB_ICONEXCLAMATION + MB_SYSTEMMODAL, "Exiting Windows")

    'Determine user response
    If Response = IDCANCEL Then  'User canceled
        Cancel = True  'Stop program from ending
        Exit Sub  'Leave
    End If

    'Program is going to end, write out inilization file
    Call iniFile("Write")

    End  'Make sure program terminates
End Sub

Sub iniFile (Action As String)
    Dim InputData As String  'Place to hold read in data
    Dim CountA As Integer  'For counter
    
    On Error Resume Next  'Skip past any file errors
    
    'Determine if reading or writing inilization file
    If Action = "Read" Then  'Reading it
        'Open up the file, or at least try
        Open App.Path + "\WinShell.ini" For Input As #1

        'Check if their was an error, if so file probably doesn't
        'exist so just leave
        If Err <> 0 Then Exit Sub

        Do Until EOF(1)  'Read in data until end of file
            Input #1, InputData  'Get line
            
            If InputData = "" Then
                'Blank line, ignore
            Else
                'Resize programs array and add inputed data
                ReDim Preserve Programs(UBound(Programs) + 1)
                Programs(UBound(Programs)) = InputData
            
                'Read in the program name and store it
                Input #1, InputData
                ReDim Preserve ProgramNames(UBound(ProgramNames) + 1)
                ProgramNames(UBound(ProgramNames)) = InputData
            End If
        Loop
    ElseIf Action = "Write" Then  'Writing data
        'Try to open a file for writing
        Open App.Path + "\WinShell.ini" For Output As #1

        'Check if to see if their was a error, if so tell user
        If Err <> 0 Then
            MsgBox Error(Err), 48, "Error number " & Err
            Exit Sub  'Leave sub
        End If

        'Write programs into file
        For CountA = (LBound(Programs) + 1) To UBound(Programs)
            Write #1, Programs(CountA)
            Write #1, ProgramNames(CountA)
        Next CountA
    End If

    Close #1  'Close the file
End Sub

Sub SetupButtons ()
    Dim CountA As Integer  'For counter
    
    On Error Resume Next  'Incase control doesn't exist
    
    'Unload any buttons that are already loaded
    For CountA = 1 To UBound(Programs)
        Unload cmdProgram(CountA)
    Next CountA
    
    'Check to see if their are no programs
    If Programs(1) = "" Then  'None
        cmdProgram(0).Visible = False  'Hide the only visible button
        Exit Sub  'Leave
    Else  'Setup the first button
        cmdProgram(0).Caption = ProgramNames(1)

        cmdProgram(0).Top = 255
        cmdProgram(0).Left = 100
        cmdProgram(0).Width = Len(ProgramNames(1)) * 150
        cmdProgram(0).Visible = True
    End If

    'Make buttons for programs
    For CountA = 2 To UBound(Programs)
        'Check to see if their is a blank spot in the array
        'if not then add the program button
        If Not ProgramNames(CountA) = "" Then
            
            'Make a new program command button
            Load cmdProgram(CountA - 1)
            
            'Setup the new button
            cmdProgram(CountA - 1).Caption = ProgramNames(CountA)
            cmdProgram(CountA - 1).Visible = True  'Show button
            cmdProgram(CountA - 1).Width = Len(ProgramNames(CountA)) * 150
            
            'Check to see if a new row needs to be started
            If (cmdProgram(CountA - 2).Left + cmdProgram(CountA - 2).Width + cmdProgram(CountA - 1).Width) > Me.Width Then
                'Next button would go out of screen, move it to next row instead
                cmdProgram(CountA - 1).Left = 100
                cmdProgram(CountA - 1).Top = cmdProgram(CountA - 2).Top + cmdProgram(CountA - 2).Height
            Else  'Just put button in current row
                cmdProgram(CountA - 1).Left = cmdProgram(CountA - 2).Left + cmdProgram(CountA - 2).Width
                cmdProgram(CountA - 1).Top = cmdProgram(CountA - 2).Top
            End If
        End If
    Next CountA
End Sub

