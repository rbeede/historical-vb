VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Screen Saver Controller Configuration"
   ClientHeight    =   4470
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5895
   Icon            =   "Screen Saver Controller Main Form.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4470
   ScaleWidth      =   5895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSetPassword 
      Caption         =   "&Password Protect Screen Saver"
      Height          =   495
      Left            =   240
      TabIndex        =   7
      Top             =   3720
      Width           =   2655
   End
   Begin MSComDlg.CommonDialog cdgAddFile 
      Left            =   2040
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdHide 
      Caption         =   "&Close Configuration"
      Height          =   495
      Left            =   3000
      TabIndex        =   6
      Top             =   3720
      Width           =   2655
   End
   Begin VB.CommandButton cmdSettings 
      Caption         =   "&Setup Selected Screen Saver"
      Height          =   495
      Left            =   3000
      TabIndex        =   5
      Top             =   3120
      Width           =   2655
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "P&review Selected Screen Saver"
      Height          =   495
      Left            =   3000
      TabIndex        =   4
      Top             =   2520
      Width           =   2655
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "&Remove Screen Saver"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   3120
      Width           =   2655
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add Screen Saver"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   2520
      Width           =   2655
   End
   Begin VB.ListBox lstScreenSavers 
      Height          =   1815
      Left            =   240
      TabIndex        =   0
      ToolTipText     =   "This is the list of screen savers controlled by this program."
      Top             =   600
      Width           =   5415
   End
   Begin VB.Label lblCurrentScreenSavers 
      AutoSize        =   -1  'True
      Caption         =   "Current Screen Savers:"
      Height          =   195
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   1650
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub iniFile(Action As Boolean)
    Dim i As Integer  'For counters
    Dim tmpData As String  'Array for holding read data
    
    On Error GoTo ErrorHandler  'Incase error occurs with file
    
    'Determine what action to take
    If Action = 0 Then  'Read program .ini file
        'Open the config file and read the program settings
        Open App.Path + "\Screen Saver Controller.ini" For Input As #1
        
        Input #1, tmpData  'Read the first data from the config file
        
        'Loop until all the data has been read
        Do Until tmpData = ""
            frmMain!lstScreenSavers.AddItem tmpData  'Add the screen saver
            
            'Check to see if the end of file has been reached, if so leave the loop
            If EOF(1) = True Then Exit Do
            
            Input #1, tmpData 'Get the next screen saver
        Loop
    ElseIf Action = -1 Then  'Write program .ini file
        'Open the config file and write the program settings
        Open App.Path + "\Screen Saver Controller.ini" For Output As #1

        'Loop through the list of screen savers
        For i = 0 To frmMain!lstScreenSavers.ListCount - 1
            Write #1, frmMain!lstScreenSavers.List(i) 'Save the screen saver
        Next i
    End If
    
    Close #1  'Close the file
    
    Exit Sub  'Leave this sub, error handling isn't necessary since no errors occured
    
ErrorHandler:  'Label to jump to incase of erros
    'Determine what was trying to be done
    If Action = 0 And Err.Number = 53 Then
        'The file wasn't found, probably first time program has ever run
        'just ignore it and continue loading the program
        Err.Number = 0  'Reset error to nothing
    ElseIf Action = 0 And Err.Number = 62 Then
        'End of file was reached during reading
        Resume Next  'Continue with the next statment after where the error occured
    ElseIf Action = 0 And Err.Number <> 53 Then
        'Some kind of other error reading the configuration file, tell the user
        MsgBox "Error Number " & Err.Number & " reading configuration file." + vbCrLf + Err.Description, vbCritical, "File Error"
    ElseIf Action = 1 Then
        'Some kind of error writing the configuration file, tell the user
        MsgBox "Error Number " & Err.Number & " writing configuration file." + vbCrLf + Err.Description + vbCrLf + "Settings may now be lost!", vbCritical, "File Error"
    End If
        
    Close  'Close any open files
End Sub

Private Sub cmdAdd_Click()
    
    'Set the file filters for common dialog box file open display
    cdgAddFile.Filter = "Screen Saver (*.scr)|*.scr|All Files (*.*)|*.*"
    
    cdgAddFile.filename = ""  'Clear out any old entry
    
    cdgAddFile.ShowOpen  'Pop open a file box
    
    If cdgAddFile.filename = "" Then Exit Sub  'User canceled
    
    'Add screen saver to list
    lstScreenSavers.AddItem cdgAddFile.filename
End Sub

Private Sub cmdHide_Click()
    Me.WindowState = vbMinimized  'Minimize this form to hide it
End Sub

Private Sub cmdPreview_Click()
    Dim taskID  'For holding return value from shell command
    
    If lstScreenSavers.ListIndex = -1 Then  'No screen saver selected
        'Tell user to select one first
        MsgBox "Please select a screen saver first.", vbExclamation, "Error"
        Exit Sub  'Leave sub
    End If
    
    'Execute screen saver
    taskID = Shell(lstScreenSavers.List(lstScreenSavers.ListIndex) + " /s", vbNormalFocus)
    
    'Check to see if screen saver ran
    If taskID = 0 Then  'It did not run
        MsgBox "Error running screen saver:  " + lstScreenSavers.List(lstScreenSavers.ListIndex), vbExclamation
    End If
End Sub

Private Sub cmdRemove_Click()
    If lstScreenSavers.ListIndex = -1 Then  'Nothing selected
        MsgBox "You must select a screen saver first.", vbExclamation, "Error"
        Exit Sub  'Leave sub
    End If

    'Remove the screen saver
    lstScreenSavers.RemoveItem lstScreenSavers.ListIndex
End Sub

Private Sub cmdSetPassword_Click()
    'Tell user not working
    MsgBox "This feature doesn't work yet." + vbCrLf + "Please e-mail me if you know how to make VB set the standard Window's screen saver password and how to check for passwords.", vbExclamation, "Unsupported"
End Sub

Private Sub cmdSettings_Click()
    Dim taskID  'For holding return value from shell command
    
    If lstScreenSavers.ListIndex = -1 Then  'No screen saver selected
        'Tell user to select one first
        MsgBox "Please select a screen saver first.", vbExclamation, "Error"
        Exit Sub  'Leave sub
    End If
    
    'Execute screen saver with config option
    taskID = Shell(lstScreenSavers.List(lstScreenSavers.ListIndex) + " /c", vbNormalFocus)
    
    'Check to see if screen saver ran
    If taskID = 0 Then  'It did not run
        MsgBox "Error running screen saver setup:  " + lstScreenSavers.List(lstScreenSavers.ListIndex), vbExclamation
    End If
End Sub

Private Sub Form_Load()
    Load frmSysTray  'Load icon in system tray

    Me.Visible = False  'Don't show this at startup

    'Read the configuration file (if one exists)
    Call iniFile(0)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'Check to see if Windows is shutting down
    If UnloadMode = vbAppWindows Then
        'It is, terminate program without asking
        End  'End program
    End If
End Sub

Private Sub Form_Resize()
    'Check to see if window is minimized
    If Me.WindowState = vbMinimized Then
        'Hide window from user and "Restore" back to original
        'size for next time it is shown
        Me.Visible = False  'Hide
        Me.WindowState = vbNormal  'Restore window
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim Response As Integer  'Stores user response
   
    'Ask user if they really wish to quit
    Response = MsgBox("Are you sure you wish to terminate this program.  Your screen savers will not come on unless you restart the program.  If you just wish to close the window minimize it instead and it will return to your system tray.", vbQuestion + vbYesNo)
    
    If Response = vbYes Then  'User wants to terminate program
        Call iniFile(-1)  'Save configuration

        Unload frmSysTray  'Unload the system tray form
        
        End  'Terminate program
    Else
        Cancel = True  'Abort unload
    End If
End Sub
