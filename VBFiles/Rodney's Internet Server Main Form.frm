VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rodney is"
   ClientHeight    =   1350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2910
   Icon            =   "Rodney's Internet Server Main Form.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1350
   ScaleWidth      =   2910
   StartUpPosition =   3  'Windows Default
   WindowState     =   1  'Minimized
   Begin VB.Timer tmrPause 
      Interval        =   500
      Left            =   360
      Top             =   0
   End
   Begin MSWinsockLib.Winsock sckServer 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   327681
   End
   Begin VB.CommandButton cmdMinimize 
      Caption         =   "&Minimize Window"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   2415
   End
   Begin VB.Shape shpLight 
      BackStyle       =   1  'Opaque
      Height          =   255
      Left            =   2400
      Shape           =   3  'Circle
      Top             =   240
      Width           =   255
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Rodney is"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2130
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub Pause()
    'Use a timer since it will always wait the correct
    'amount of time
    
    tmrPause.Enabled = True  'One second pause
    
    Do  'Loop until the timer disables itself after 1 second
        DoEvents  'Let the timer count down
    Loop Until tmrPause.Enabled = False
End Sub

Private Sub cmdMinimize_Click()
    Me.WindowState = vbMinimized  'Minimize window
End Sub

Private Sub Form_Load()
    'Fill in caption and label with status
    Me.Caption = "Rodney is offline"
    lblStatus.Caption = "Rodney is offline"

    'Set the options up for the connection
    sckServer.LocalPort = 9999  'Set local port number
    sckServer.Listen  'Listen for connections
    Call Pause  'Short pause

    Me.Show  'Make sure form pops up
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'Check to see if windows is being shutdown
    If UnloadMode = vbAppWindows Then  'Windows is ending
        End  'Terminate program without asking
    Else
        Cancel = True  'Don't allow program to unload
        
        'Tell user this program cannot be terminated
        MsgBox "You may not terminate this program.  This program will only end when you shutdown Windows.", vbCritical + vbOKOnly, "Cannot Comply"
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Tell user this program cannot be terminated
    MsgBox "You may not terminate this program.  This program will only end when you shutdown Windows.", vbCritical + vbOKOnly, "Cannot Comply"
    
    'Stop unload
    Cancel = True
End Sub

Private Sub sckServer_Close()
    'Other end closed so listen again
    
    sckServer.Close  'Close socket
    sckServer.LocalPort = 0  'Clear out address control
    sckServer.LocalPort = 9999  'Set local port number
    
    sckServer.Listen  'Listen for connections
    Call Pause  'Short pause

    'Say Rodney is Offline
    Me.Caption = "Rodney is offline"
    lblStatus.Caption = "Rodney is offline"
End Sub

Private Sub sckServer_ConnectionRequest(ByVal requestID As Long)
    'Other program is connecting
    
    sckServer.Close  'Close socket
    sckServer.Accept requestID  'Accept connection request
    Call Pause  'Short pause
End Sub

Private Sub sckServer_DataArrival(ByVal bytesTotal As Long)
    Dim IncomingData As String  'Place to store incoming data
    
    sckServer.GetData IncomingData, vbString  'Get data
    Call Pause  'Short pause
    
    'Check string for current status
    If Left$(IncomingData, 9) = "/~Status@" Then
        'Read off new status
        If Mid$(IncomingData, 10, Len(IncomingData)) = "Online" Then  'Show online status
            Me.Caption = "Rodney is online"
            lblStatus.Caption = "Rodney is online"
        ElseIf Mid$(IncomingData, 10, Len(IncomingData)) = "Offline" Then  'Show offline status
            Me.Caption = "Rodney is offline"
            lblStatus.Caption = "Rodney is offline"
        End If
    Else  'Just dummy data, flash "Active Light"
        shpLight.BackColor = &HFF00&  'Light Green
        Call Pause  'Wait
        
        'Send back dummy response
        If sckServer.State = 7 Then
            sckServer.SendData "Hello World!!!!!!!!!!!!!"
            DoEvents  'Go On
        End If
        
        shpLight.BackColor = &HFFFFFF  'White
    End If
End Sub

Private Sub sckServer_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    'Attempt to restart listening
    sckServer.Close  'Close socket
    sckServer.LocalPort = 0  'Clear out address
    sckServer.LocalPort = 9999  'Set local port
    sckServer.Listen  'Listen for connections
    Call Pause  'Short pause
End Sub

Private Sub tmrPause_Timer()
    'Disable timer since it has allowed 1 second to pass
    tmrPause.Enabled = False
End Sub
