VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Phone Status"
   ClientHeight    =   795
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3660
   ControlBox      =   0   'False
   Icon            =   "Phone Status Main Form.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   795
   ScaleWidth      =   3660
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Timer tmrCheckStatus 
      Interval        =   1000
      Left            =   2520
      Top             =   600
   End
   Begin MSCommLib.MSComm MSComm 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   327681
      CommPort        =   2
      DTREnable       =   -1  'True
   End
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      Caption         =   "The phone is ringing!!!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   3210
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long


Private Sub Form_Load()
    Dim Instring As String  'Buffer to hold input string
    
    On Error Resume Next  'Skip past any errors
    
    Load frmSysTray  'Load System Tray Icon
    
    MSComm.CommPort = 2  'Use COM2
    
    '9600 baud, no parity, 8 data, and 1 stop bit
    MSComm.Settings = "9600,N,8,1"
    
    'Tell the control to read entire buffer when Input is used
    MSComm.InputLen = 0
    
    MSComm.PortOpen = True  'Open the port
    
    MSComm.Output = "AT Z" + Chr$(13)  'Send the attention command to the modem
    
    'Wait for data to come back to the serial port.
    Do
        DoEvents  'Allow Windows to go on
    Loop Until MSComm.InBufferCount >= 2
    
    'Read the "OK" response data in the serial port.
    Instring = MSComm.Input
        
    If Err.Number <> 0 Then  'Tell user error
        MsgBox "Error number:  " & Err.Number & "  " & Err.Description, vbCritical, "Modem Error"
        Unload Me  'Terminate program
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbAppWindows Then  'Windows is closing
        Unload frmSysTray  'Remove icon from tray
        
        End  'terminate program
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Tell user this program cannot be ended
    
    MsgBox "You may not terminate this program!", vbExclamation, "Phone Status"
    
    Cancel = True  'Stop unload
End Sub

Private Sub tmrCheckStatus_Timer()
    Dim InData As String  'For incoming data
    Dim i As Integer  'For counter
    
    MSComm.InputLen = 0  'Clear buffer
    
    'Check the input buffer for a RING
    Do
        DoEvents
    Loop Until MSComm.InBufferCount >= 2
    
    InData = MSComm.Input  'Read in buffer
    
    If Not InStr(1, InData, "RING") = 0 Then  'Phone is ringing
        'Pop up program and alert that phone is ringing
        Me.Visible = True  'Show form
        lblStatus.Visible = True  'Show message
        DoEvents  'Windows goes on
            
        AppActivate "Phone Status"
        
        'Wake up monitor hopefully
        SendKeys "A", False
        
        DoEvents  'Allow Windows to go on
        
        'Play a ringing sound
        i = PlaySound(ByVal CStr(App.Path + "\ring.wav"), 1, 0)
        
        Me.Visible = False  'Hide form
        lblStatus.Visible = False  'Hide message
        DoEvents  'Windows goes on
    Else
        Print InData
    End If
End Sub
