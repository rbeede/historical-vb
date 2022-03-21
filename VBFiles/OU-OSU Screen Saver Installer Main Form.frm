VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Screen Saver Installer"
   ClientHeight    =   3975
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6225
   Icon            =   "OU-OSU Screen Saver Installer Main Form.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   6225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3240
      TabIndex        =   3
      Top             =   3000
      Width           =   2775
   End
   Begin VB.CommandButton cmdInstall 
      Caption         =   "&Install"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   3000
      Width           =   2775
   End
   Begin MCI.MMControl MMControl1 
      Height          =   495
      Left            =   1080
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   873
      _Version        =   327681
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.Label lblInfo 
      Caption         =   "lblInfo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   5655
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Private Sub cmdExit_Click()
    End
End Sub

Private Sub cmdInstall_Click()
    Dim WinDir As String  'For holding Windows Directory
    Dim X  'For holding result of API call
    Dim Source As String, Dest As String
    
    WinDir = Space(144)  'Fill in data to give string size
    
    'Call procedure, passing variable to be filled and max size it can take
    X = GetWindowsDirectory(WinDir, Len(WinDir))

    WinDir = Trim$(WinDir)  'Cut out extra spaces
    WinDir = Left$(WinDir, Len(WinDir) - 1)  'Cut out invalid character at end

    Screen.MousePointer = vbHourglass
    'Copy the files to the Windows directory
    'Kill WinDir + "\boomer.wav"
    Source = App.Path + "\boomer.wav"
    Dest = WinDir + "\boomer.wav"
    FileCopy Source, Dest
    Source = App.Path + "\cowboys.wav"
    Dest = WinDir + "\cowboys.wav"
    FileCopy Source, Dest
    Source = App.Path + "\OU-OSU Screen Saver.scr"
    Dest = WinDir + "\OU-OSU Screen Saver.scr"
    FileCopy Source, Dest
    
    Screen.MousePointer = vbNormal
    
    MsgBox "Finished copying files.  You may uninstall this program if you wish.", vbInformation, "Successful Copy"
    
    End
End Sub

Private Sub Form_Load()
    'Fill in the information about what to do next
    lblInfo.Caption = "You have installed the screen saver on your computer."
    lblInfo.Caption = lblInfo.Caption + vbCrLf
    lblInfo.Caption = lblInfo.Caption + vbCrLf + "To make it easy to select your screen saver in Windows this program will copy the screen saver files into your Windows directory."
    lblInfo.Caption = lblInfo.Caption + vbCrLf
    lblInfo.Caption = lblInfo.Caption + vbCrLf + "Once the program is finished you can tell Windows to use the screen saver by going into Control Panel, double-clicking on Display, click on the Screen Saver tab, and select the OU-OSU screen saver from the list."
    lblInfo.Caption = lblInfo.Caption + vbCrLf
    lblInfo.Caption = lblInfo.Caption + vbCrLf + "Click on Install to install the screen saver into your Windows directory or click on Exit to quit this program."
End Sub
