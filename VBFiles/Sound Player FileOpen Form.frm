VERSION 5.00
Begin VB.Form frmFileOpen 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select File to Play"
   ClientHeight    =   4140
   ClientLeft      =   2400
   ClientTop       =   1710
   ClientWidth     =   4470
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4140
   ScaleWidth      =   4470
   Begin VB.FileListBox File 
      Appearance      =   0  'Flat
      Height          =   2955
      Left            =   120
      Pattern         =   "*.Wav"
      TabIndex        =   2
      Top             =   480
      Width           =   1935
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   435
      Left            =   2280
      TabIndex        =   5
      Top             =   3600
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Appearance      =   0  'Flat
      Caption         =   "&OK"
      Height          =   435
      Left            =   1080
      TabIndex        =   4
      Top             =   3600
      Width           =   975
   End
   Begin VB.DirListBox SoundDir 
      Appearance      =   0  'Flat
      Height          =   2505
      Left            =   2280
      TabIndex        =   1
      Top             =   480
      Width           =   2055
   End
   Begin VB.DriveListBox SoundDrive 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2280
      TabIndex        =   0
      Top             =   3120
      Width           =   2055
   End
   Begin VB.Label lblPathFile 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   75
   End
End
Attribute VB_Name = "frmFileOpen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCancel_Click()
    Me.Hide ' Hide Form
    frmSound.Enabled = True ' Enabled sound form
End Sub

Private Sub cmdOK_Click()
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

Private Sub File_Click()
    ' Declare needed variable
    Dim vPath As String

    If Len(SoundDir.Path) > 3 Then ' Path needs \
       vPath = SoundDir.Path + "\"
    Else ' Path is root
       vPath = SoundDir.Path
    End If

    ' Show path and file in label box
    lblPathFile = vPath + File.filename
End Sub

Private Sub File_DblClick()
    Call cmdOK_Click ' command button ok
End Sub

Private Sub SoundDir_Change()
    File.Path = SoundDir.Path ' Set file list path up
    lblPathFile.Caption = ""
End Sub

Private Sub SoundDrive_Change()
    On Error Resume Next ' If there is a error just go on

    ' Set directory box new path
    SoundDir.Path = SoundDrive.Drive

    If Err Then ' Error occured
       MsgBox Error$ + ".", 16, "Error " & Err
       SoundDrive.Drive = SoundDir.Path
    End If
End Sub

