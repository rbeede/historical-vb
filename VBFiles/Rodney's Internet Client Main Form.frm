VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rodney's Internet Status"
   ClientHeight    =   1470
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2880
   Icon            =   "Rodney's Internet Client Main Form.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1470
   ScaleWidth      =   2880
   StartUpPosition =   2  'CenterScreen
   Begin MSWinsockLib.Winsock sckClient 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton cmdOffline 
      Caption         =   "O&ffline"
      Height          =   495
      Left            =   1440
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton cmdOnline 
      Caption         =   "&Online"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
   Begin VB.Shape shpLight 
      BackColor       =   &H00FFFFFF&
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
      Width           =   2085
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub Wait()
    Dim Start As Single
    
    Start = Timer  'Get the time first started waiting
    
    'Loop for one second
    Do While Timer < Start + 1
        DoEvents  'Allow Windows to process
    Loop
End Sub
Private Sub cmdOffline_Click()
    'Send Offline packet
    sckClient.SendData "/~Status@Offline"
    DoEvents  'Allow Windows to process

    'Update display
    lblStatus.Caption = "Rodney is offline"
End Sub

Private Sub cmdOnline_Click()
    'Send online packet
    sckClient.SendData "/~Status@Online"
    DoEvents  'Allow Windows to process
    
    'Update display
    lblStatus.Caption = "Rodney is online"
End Sub

Private Sub Form_Load()
    On Error Resume Next  'Skip past expected errors
    
    'Try to connect
    sckClient.Close  'Close socket
    sckClient.LocalPort = 0  'Clear out address
    sckClient.RemoteHost = InputBox$("Enter the address")
    sckClient.RemotePort = 9999  'Set remote port
    sckClient.Connect  'Attempt to connect
    
    DoEvents  'Allow Windows to process
    
    'Indicate that I am offline
    lblStatus.Caption = "Rodney is offline"

    'Wait until connected
    Do While sckClient.State <> 7
        DoEvents  'Allow Windows to go on
        
        If sckClient.State = 9 Then
            Me.Caption = sckClient.State
            Exit Do  'Error
        End If
    Loop

    'Send the first dummy signal
    sckClient.SendData "Hello Dummy!!!"
    Call Wait  'Force a pause

    Me.Show  'Show form
End Sub

Private Sub Form_Unload(Cancel As Integer)
    sckClient.Close  'Close connection
    sckClient.LocalPort = 0  'Release address
    DoEvents  'Allow Windows to process

    End  'Terminate program
End Sub

Private Sub sckClient_Close()
    'Other end closed, quit program
    Unload Me
End Sub

Private Sub sckClient_DataArrival(ByVal bytesTotal As Long)
    Dim Dummy As String  'To hold useless data
    
    'Read incoming data out of buffer
    sckClient.GetData Dummy, vbString

    'Flash "Active Light" to tell user computers are listening
    shpLight.BackColor = &HFF00&  'Light Green
        
    Call Wait  'Force a wait
    
    'Send back a dummy reply
    If sckClient.State = 7 Then
        sckClient.SendData "Hello World!"
        DoEvents  'Go on
    End If
    
    shpLight.BackColor = &HFFFFFF  'White
    
    Me.Refresh  'Make sure gets redrawn
End Sub

Private Sub sckClient_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    'Their was a error, try to reconnect
    
    'Try to connect
    sckClient.Close  'Close socket
    sckClient.LocalPort = 0  'Clear out address
    sckClient.RemoteHost = "Pavilion"  'Set remote host
    sckClient.RemotePort = 9999  'Set remote port
    sckClient.Connect  'Attempt to connect
    
    DoEvents  'Allow Windows to process
    
    'Indicate that I am offline
    lblStatus.Caption = "Rodney is offline"

    'Wait until connected
    Do While sckClient.State <> 7
        DoEvents  'Allow Windows to go on
        
        If sckClient.State = 9 Then
            Me.Caption = sckClient.State
            Exit Do  'Error
        End If
    Loop

    'Send the first dummy signal
    sckClient.SendData "Hello Dummy!!!"
    Call Wait  'Force a pause
End Sub
