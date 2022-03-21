VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "File Serializer"
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6015
   Icon            =   "File Serializer Main Form.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   6015
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   495
      Left            =   3120
      TabIndex        =   12
      Top             =   5280
      Width           =   2775
   End
   Begin VB.CommandButton cmdSerialize 
      Caption         =   "&Serialize"
      Height          =   495
      Left            =   120
      TabIndex        =   11
      Top             =   5280
      Width           =   2775
   End
   Begin VB.TextBox txtFilePrefix 
      Height          =   315
      Left            =   3120
      MaxLength       =   3
      TabIndex        =   10
      Top             =   4800
      Width           =   2775
   End
   Begin VB.DirListBox dirDestination 
      Height          =   1665
      Left            =   3120
      TabIndex        =   8
      Top             =   2400
      Width           =   2775
   End
   Begin VB.DriveListBox drvDestination 
      Height          =   315
      Left            =   3120
      TabIndex        =   7
      Top             =   4080
      Width           =   2775
   End
   Begin VB.ComboBox cboFilePattern 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   4800
      Width           =   2775
   End
   Begin VB.DirListBox dirSource 
      Height          =   1665
      Left            =   120
      TabIndex        =   2
      Top             =   2400
      Width           =   2775
   End
   Begin VB.DriveListBox drvSource 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   4080
      Width           =   2775
   End
   Begin VB.Label lblFilePrefix 
      AutoSize        =   -1  'True
      Caption         =   "File Prefix:"
      Height          =   195
      Left            =   3120
      TabIndex        =   9
      Top             =   4560
      Width           =   720
   End
   Begin VB.Label lblDestination 
      AutoSize        =   -1  'True
      Caption         =   "Destination Directory:"
      Height          =   195
      Left            =   3120
      TabIndex        =   6
      Top             =   2160
      Width           =   1515
   End
   Begin VB.Label lblFilePattern 
      AutoSize        =   -1  'True
      Caption         =   "File Pattern:"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   4560
      Width           =   840
   End
   Begin VB.Label lblSource 
      AutoSize        =   -1  'True
      Caption         =   "Source Directory:"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   2160
      Width           =   1230
   End
   Begin VB.Label lblInstructions 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "lblInstructions"
      Height          =   1815
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5775
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim FilePattern As String  'For file pattern
Const FileLimit As Integer = 9999  'Set serialization limit
Const FileLimitPlaces As Integer = 4  'For number of digits

Sub Busy(State As Integer)
    'Determine state and setup form correctly
    If State = True Then  'Program busy
        Screen.MousePointer = vbArrowHourglass  'Change mouse cursor
        
        'Disable parts of form
        cboFilePattern.Enabled = False  'File pattern box
        txtFilePrefix.Enabled = False  'File prefix box
        dirSource.Enabled = False  'Source Directory
        dirDestination.Enabled = False  'Dest directory
        drvSource.Enabled = False  'Drive source
        drvDestination.Enabled = False  'Drive dest
        cmdSerialize.Enabled = False  'Serialization button
    Else  'Program isn't busy
        Screen.MousePointer = vbNormal  'Change mouse cursor
        
        'Enable parts of form
        cboFilePattern.Enabled = True  'File pattern box
        txtFilePrefix.Enabled = True  'File prefix box
        dirSource.Enabled = True  'Source Directory
        dirDestination.Enabled = True  'Dest directory
        drvSource.Enabled = True  'Drive source
        drvDestination.Enabled = True  'Drive dest
        cmdSerialize.Enabled = True  'Serialization button
    End If
End Sub

Private Sub cboFilePattern_Click()
    'Determine file pattern to use
    Select Case cboFilePattern.ListIndex
        Case 0:  'JPEG File
            FilePattern = "*.jpg"  'Set file pattern variable
        Case 1:  'Gif File
            FilePattern = "*.gif"  'Set file pattern variable
        Case 2:  'Bitmap File
            FilePattern = "*.bmp"  'Set file pattern variable
        Case 3:  'Any File
            FilePattern = "*.*"  'Set file pattern variable
    End Select
End Sub

Private Sub cmdExit_Click()
    End  'Terminate program
End Sub

Private Sub cmdSerialize_Click()
    Dim SourceFileList() As String  'Array to hold file listing
    Dim i As Long  'For counter
    Dim currFile As String  'For current file
    Dim SerialNum As Long  'For storing file serial number
    
    On Error Resume Next  'Skip past file errors
    
    'Show user program is busy
    Call Busy(True)
     
    'Get a directory listing
    ReDim SourceFileList(0)  'Resize array
    SourceFileList(0) = Dir(dirSource.Path + "\" + FilePattern)  'Get first entry

    If SourceFileList(0) = "" Then  'No files were found, tell user
        'Show user program is no longer busy
        Call Busy(False)
        
        MsgBox "No files were found in the source directory with the file pattern you specified.  Try again.", vbCritical, "Error"
        Exit Sub  'Leave sub
    Else
        currFile = Dir  'Get next entry
    End If
    
    'Get the rest of the entries
    Do While currFile <> ""
        ReDim Preserve SourceFileList(UBound(SourceFileList) + 1)  'Resize array
        SourceFileList(UBound(SourceFileList)) = currFile  'Store current listing
        currFile = Dir  'Get next entry
    Loop

    'Check the destination directory for any earlier serializations
    'Get a directory listing for any files matching prefix and number format
    currFile = Dir(dirDestination.Path + "\" + txtFilePrefix.Text + String(FileLimitPlaces, "?") + Right$(FilePattern, 4))

    If currFile <> "" Then
        'Their are old ones, run through until the last one is found
        
        For i = FileLimit To 0 Step -1
            currFile = Dir$(dirDestination.Path + "\" + txtFilePrefix.Text + Format$(Str$(i), String(FileLimitPlaces, "0")) + Right$(FilePattern, 4))

            'Check for error
            If currFile <> "" Then
                'File exists, serial number after this files is the place to start
                SerialNum = i + 1  'Set serial number
                Exit For  'Leave loop
            End If
        
            DoEvents  'Allow Windows to go on
        Next i
        
        'Check to see if to many files are in here now (over 99,999)
        If SerialNum > FileLimit Then  'Too many, tell user
            'Show user program is no longer busy
            Call Busy(False)
            
            MsgBox "Program found previous serialized files and their were too many.  Try a different prefix or move the files.", vbExclamation, "Error"
            Exit Sub  'Leave sub
        Else
            'Tell user what happened
            MsgBox "Program found previous serialized files and will continue starting with " + Str$(SerialNum), vbInformation, "File Serialization"
        End If
    Else
        SerialNum = 0  'No previous files, start at zero
    End If

    'Start serialization and moving of files
    For i = 0 To UBound(SourceFileList)
        'Rename & move the file
        Name dirSource.Path + "\" + SourceFileList(i) As dirDestination.Path + "\" + txtFilePrefix.Text + Format$(Str$(SerialNum), String(FileLimitPlaces, "0")) + Right$(FilePattern, 4)
   
        'Check for any errors
        If Err.Number <> 0 Then  'Error, tell user
            MsgBox Err.Description, vbExclamation, "Error " + Str$(Err.Number)
        End If
    
        SerialNum = SerialNum + 1  'Increment current serial number
        
        'Check to see if there are to many files (99,999 is max)
        If SerialNum > FileLimit Then  'To many, tell user
            'Show user program is no longer busy
            Call Busy(False)
            
            MsgBox "There are to many serialized files in the destination folder.  Try a different prefix or move some of the serialized files from the destination.  Stopped at " + SourceFileList(i), vbCritical, "Error"
            Exit Sub  'Leave sub
        End If
    Next i

    'Show user program is no longer busy
    Call Busy(False)
    
    'Tell user program is done
    MsgBox "Serialization is complete.", vbInformation, "File Serializer"
End Sub

Private Sub drvDestination_Change()
    On Error Resume Next  'Incase of device error
    
    dirDestination.Path = drvDestination.Drive  'Change the directory listing

    'Check for a error
    If Err.Number <> 0 Then  'Error, tell user
        MsgBox Err.Description, vbCritical, "Error " + Str$(Err.Number)
        drvDestination.Drive = dirDestination.Path  'Reset current drive
    End If
End Sub

Private Sub drvSource_Change()
    On Error Resume Next  'Incase of device error
    
    dirSource.Path = drvSource.Drive  'Change the directory listing

    'Check for a error
    If Err.Number <> 0 Then  'Error, tell user
        MsgBox Err.Description, vbCritical, "Error " + Str$(Err.Number)
        drvSource.Drive = dirSource.Path  'Reset current drive
    End If
End Sub

Private Sub Form_Load()
    'Fill in the instructions
    lblInstructions = "Instructions:" + vbCrLf
    lblInstructions = lblInstructions + "Step 1:  Select the drive and path with the source files." + vbCrLf
    lblInstructions = lblInstructions + "Step 2:  Select the file format pattern to look for." + vbCrLf
    lblInstructions = lblInstructions + "Step 3:  Select the destination folder for the serialized files to be placed in." + vbCrLf
    lblInstructions = lblInstructions + "Step 4:  Enter in a file prefix.  Three characters at max.  This step optional." + vbCrLf
    lblInstructions = lblInstructions + "Step 5:  Click on the Serialize button to serialize the files." + vbCrLf
    lblInstructions = lblInstructions + vbCrLf + "Files will be renamed and moved to the destination folder when you begin the process."

    'Add file patterns to combo box
    cboFilePattern.AddItem "JPEG Files (*.jpg)"  'JPEGs
    cboFilePattern.AddItem "GIF Files (*.gif)"  'GIFs
    cboFilePattern.AddItem "Bitmap Files (*.bmp)"  'Bitmaps
    cboFilePattern.AddItem "All Files (*.*)"  'Anything

    'Select a default file pattern
    cboFilePattern.ListIndex = 0
    Call cboFilePattern_Click  'Fill in a file pattern
End Sub

