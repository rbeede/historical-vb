VERSION 2.00
Begin Form frmFileOpen 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Double
   Caption         =   "Select File to Play"
   ClientHeight    =   4140
   ClientLeft      =   2400
   ClientTop       =   1710
   ClientWidth     =   4470
   ControlBox      =   0   'False
   Height          =   4545
   Left            =   2340
   LinkTopic       =   "Form1"
   ScaleHeight     =   4140
   ScaleWidth      =   4470
   Top             =   1365
   Width           =   4590
   Begin FileListBox File 
      Height          =   2955
      Left            =   120
      Pattern         =   "*.Wav"
      TabIndex        =   2
      Top             =   480
      Width           =   1935
   End
   Begin CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   435
      Left            =   2280
      TabIndex        =   5
      Top             =   3600
      Width           =   975
   End
   Begin CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   435
      Left            =   1080
      TabIndex        =   4
      Top             =   3600
      Width           =   975
   End
   Begin DirListBox SoundDir 
      Height          =   2505
      Left            =   2280
      TabIndex        =   1
      Top             =   480
      Width           =   2055
   End
   Begin DriveListBox SoundDrive 
      Height          =   315
      Left            =   2280
      TabIndex        =   0
      Top             =   3120
      Width           =   2055
   End
   Begin Label lblPathFile 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   75
   End
End

Sub cmdCancel_Click ()
    Me.Hide ' Hide Form
    frmSound.Enabled = True ' Enabled sound form
End Sub

Sub cmdOK_Click ()
    'Declare needed variable
    Dim Ctr As Integer

    ' Clear out file to play
    PathWavFile = ""

    ' If no wav file then exit sub
    If lblPathFile.Caption = "" Then Exit Sub

    Ctr = 11 ' reset number
    
    ' Set .Wav file path and name
    PathWavFile = lblPathFile.Caption
    
    ' Get file only to show on sound form
    WavFile = Right$(PathWavFile, 12)
    Do
      If InStr(WavFile, "\") Then
         WavFile = Right$(PathWavFile, Ctr)
      Else
         Exit Do
      End If
      Ctr = Ctr - 1
    Loop

    ' Show file on sound form
    frmSound.lblWavFile.Caption = UCase$(WavFile)
    
    Me.Hide ' Remove form
    frmSound.Enabled = True ' Enable sound form
    frmSound.SetFocus ' Set focus on sound form
End Sub

Sub File_Click ()
    ' Declare needed variable
    Dim vPath As String

    If Len(SoundDir.Path) > 3 Then ' Path needs \
       vPath = SoundDir.Path + "\"
    Else ' Path is root
       vPath = SoundDir.Path
    End If

    ' Show path and file in label box
    lblPathFile = vPath + File.FileName
End Sub

Sub File_DblClick ()
    Call cmdOK_Click ' command button ok
End Sub

Sub SoundDir_Change ()
    File.Path = SoundDir.Path ' Set file list path up
    lblPathFile.Caption = ""
End Sub

Sub SoundDrive_Change ()
    On Error Resume Next ' If there is a error just go on

    ' Set directory box new path
    SoundDir.Path = SoundDrive.Drive

    If Err Then ' Error occured
       MsgBox Error$ + ".", 16, "Error " & Err
       SoundDrive.Drive = SoundDir.Path
    End If
End Sub

