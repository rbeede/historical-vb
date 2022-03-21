VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Chat!"
   ClientHeight    =   4095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7500
   Icon            =   "Chat! Main Form.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   7500
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstConnections 
      Enabled         =   0   'False
      Height          =   2400
      Left            =   5280
      Sorted          =   -1  'True
      TabIndex        =   7
      Top             =   240
      Width           =   1935
   End
   Begin VB.CommandButton cmdSendFile 
      Caption         =   "Send File"
      Height          =   285
      Left            =   5280
      TabIndex        =   6
      Top             =   2760
      Width           =   1935
   End
   Begin VB.TextBox txtChat 
      Height          =   2415
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   240
      Width           =   4815
   End
   Begin MSWinsockLib.Winsock sckTCP 
      Index           =   0
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   327681
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   495
      Left            =   3600
      TabIndex        =   5
      Top             =   3360
      Width           =   1455
   End
   Begin VB.CommandButton cmdHostStopHost 
      Caption         =   "&Host"
      Height          =   495
      Left            =   1920
      TabIndex        =   4
      Top             =   3360
      Width           =   1455
   End
   Begin VB.CommandButton cmdSendMsg 
      Caption         =   "&Send"
      Default         =   -1  'True
      Height          =   285
      Left            =   4320
      TabIndex        =   2
      Top             =   2760
      Width           =   735
   End
   Begin VB.TextBox txtSendMsg 
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Top             =   2760
      Width           =   4095
   End
   Begin VB.CommandButton cmdConnectDisconnect 
      Caption         =   "&Connect"
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   3360
      Width           =   1455
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type Connection_Info
    sckNum As Integer  'Socket Number
    ChatName As String  'Name
End Type

Dim Connections() As Connection_Info
Dim ChatName As String  'Chatroom name

Private Sub Pause()
    Dim i As Integer  'For counter
    
    For i = 0 To 500  'Pause for half a second
        DoEvents  'Allow windows to continue
    Next i
End Sub
Private Sub cmdConnectDisconnect_Click()
    Dim i As Integer  'For counter
    Dim Response As Integer  'For users response
    
    'Need to skip past errors since we could get the error '340'
    'because a socket didn't exist
    On Error Resume Next
    
    'Determine if button is suppose to connect or disconnect
    If cmdConnectDisconnect.Caption = "&Connect" Then  'Connect
        cmdConnectDisconnect.Caption = "&Disconnect"  'Change button caption
        cmdHostStopHost.Enabled = False  'Disable host button
    
        'Ask for the remote computer address
        sckTCP(0).Tag = InputBox$("Enter the computer name or ip address of the machine you wish to connect.", "Remote Address", sckTCP(0).RemoteHost)
        
        'Check to make sure user didn't cancel
        If sckTCP(0).Tag = "" Then  'User canceled
            cmdConnectDisconnect.Caption = "&Connect"  'Reset button caption
            cmdHostStopHost.Enabled = True  'Enable host button
            Exit Sub  'Leave sub
        Else  'User entered a address
            sckTCP(0).RemoteHost = sckTCP(0).Tag  'Set address
            sckTCP(0).Tag = ""  'Clear out tag
        End If
            
        'Ask for the remote computer port number
        sckTCP(0).Tag = InputBox$("Enter the remote computer port number.", "Remote Port Number", Str$(sckTCP(0).RemotePort))
        
        'Check to see if user canceled
        If sckTCP(0).Tag = "" Then  'User canceled
            cmdConnectDisconnect.Caption = "&Connect"  'Reset button caption
            cmdHostStopHost.Enabled = True  'Enable host button
            Exit Sub  'Leave sub
        Else  'User entered a port number
            sckTCP(0).RemotePort = Val(sckTCP(0).Tag)  'Save value
            sckTCP(0).Tag = ""  'Clear out tag
        End If
        
        'Ask for a chat name
        ChatName = InputBox$("Enter the name you would like to be called.", "Name", ChatName)
        
        'Check name to see if it is blank or if user canceled
        If ChatName = "" Then  'User canceled
            cmdConnectDisconnect.Caption = "&Connect"  'Reset button caption
            cmdHostStopHost.Enabled = True  'Enable host button
            Exit Sub  'Leave sub
        End If
                
        'Make a new instance of the winsock that can be unloaded when
        'the connection is broken, that way the program won't get a
        'address error if it tries to connect again
                
        Unload sckTCP(1)  'Make sure doesn't exist
        
        Load sckTCP(1)  'Load new control
        
        sckTCP(1).RemoteHost = sckTCP(0).RemoteHost  'Set remote host
        sckTCP(1).RemotePort = sckTCP(0).RemotePort  'Set remote port
        sckTCP(1).LocalPort = 0  'Clear out local port number
        sckTCP(1).Connect  'Attempt to connect to the remote computer
        Call Pause  'Short pause
        
        'Tell user that a connection is being attempted
        txtChat.Text = txtChat.Text + vbCrLf + "Attempting to connect to " + sckTCP(1).RemoteHost + " on port number" + Str$(sckTCP(1).RemotePort) + "..."
        
        'Pause to allow connection to catch up
        Call Pause
    Else  'Disconnect
        'Ask user if they really wish to disconnect
        Response = MsgBox("Are you sure you wish to disconnect?", vbQuestion + vbYesNo, "Chat!")
        
        'Determine response
        If Response = vbYes Then  'User said yes
            sckTCP(1).Close  'Close connection
            
            Unload sckTCP(1)  'Unload control
            
            'Tell user connection is closed
            txtChat.Text = txtChat.Text + vbCrLf + "Connection closed."
                        
            'Change button caption
            cmdConnectDisconnect.Caption = "&Connect"
            
            cmdHostStopHost.Enabled = True  'Enable host button
            
            'Disable send box and button
            txtSendMsg.Enabled = False
            cmdSendMsg.Enabled = False
            
            'Clear out send text box and color it to look disabled
            txtSendMsg.Text = ""
            txtSendMsg.BackColor = frmMain.BackColor
        
            'Clear out list of connections and disable it
            lstConnections.Clear
            lstConnections.Enabled = False
        End If
    End If
End Sub

Private Sub cmdExit_Click()
    Unload frmMain  'Unload this form
End Sub

Private Sub cmdHostStopHost_Click()
    Dim i As Integer  'For counter
    Dim Response As Integer  'For users response
    
    'Need to skip past errors since we could get the error '340'
    'because a socket didn't exist
    On Error Resume Next
    
    'Determine if program should start or stop hosting
    If cmdHostStopHost.Caption = "&Host" Then  'Start hosting
        cmdHostStopHost.Caption = "&Stop &Hosting"  'Reset caption
        
        'Ask user what port to listen on
        sckTCP(0).Tag = InputBox$("Enter the local port number to listen for connections on.", "Local Port Number", sckTCP(0).LocalPort)
        
        'Check to make sure user didn't cancel
        If sckTCP(0).Tag = "" Then  'User canceled
            cmdConnectDisconnect.Enabled = True  'Enable connect button
            cmdHostStopHost.Caption = "&Host"  'Reset caption
            Exit Sub  'Leave sub
        End If
        
        sckTCP(0).LocalPort = Val(sckTCP(0).Tag)  'Set port number
        
        'Ask for a chat name
        ChatName = InputBox$("Enter the name you would like to be called.", "Name", ChatName)
        
        If ChatName = "" Then  'User canceled
            cmdHostStopHost.Caption = "&Host"  'Reset button caption
            cmdConnectDisconnect.Enabled = True  'Enable connect button
            Exit Sub  'Leave sub
        End If
        
        'Set socket to listen for connectins
        sckTCP(0).Listen
    
        'Tell user what program is doing
        txtChat.Text = txtChat.Text + vbCrLf + "Listening for connections on port number" + Str$(sckTCP(0).LocalPort) + "."

        'Enable and color in send text box and button
        txtSendMsg.Enabled = True
        txtSendMsg.BackColor = vbWhite
        cmdSendMsg.Enabled = True

        cmdConnectDisconnect.Enabled = False  'Disable connect button

        'Resize array
        ReDim Connections(1)
        
        'Fill in server entry for user name and fill in list
        Connections(0).ChatName = ChatName
        Connections(0).sckNum = 0

        lstConnections.Enabled = True  'Enable list
        lstConnections.Clear  'Clear out list

        For i = 0 To sckTCP.UBound + 1
            If Not Connections(i).ChatName = "" Then  'Not empty
                lstConnections.AddItem Connections(i).ChatName
            End If
        Next i
    Else  'Stop hosting
        'Ask user if they really want to disconnect all sockets
        Response = MsgBox("Are you sure you wish to disconnect all connections?", vbQuestion + vbYesNo, "Chat!")
        
        'Determine response
        If Response = vbYes Then  'Disconnect them all
            'Loop through any socket controls above index zero
            For i = 1 To sckTCP.UBound
                sckTCP(i).Close  'Close socket
                Unload sckTCP(i)  'Unload socket control
            Next i
    
            Erase Connections  'Clear out connections array

            sckTCP(0).Close  'Close primary socket
        
            'Disable send text box and button
            txtSendMsg.Enabled = False
            txtSendMsg.Text = ""
            txtSendMsg.BackColor = frmMain.BackColor
            cmdSendMsg.Enabled = False
            
            'Disable list of connections and clear it
            lstConnections.Enabled = False
            lstConnections.Clear
            
            cmdHostStopHost.Caption = "&Host"  'Reset caption
        
            cmdConnectDisconnect.Enabled = True  'Enable connect button
        
            'Tell user not listening
            txtChat.Text = txtChat.Text + vbCrLf + "Not listening for connections."
        End If
    End If
End Sub


Private Sub cmdSendFile_Click()
    Dim i As Integer  'For counter
    Dim FilePart() As String  'Holds parts of file
    Dim NumParts As Integer  'Number of parts of file
    Dim FileSize As Long  'Size of the file
    Dim WholeFile As String  'Stores entire file
    Dim WholeFile2nd As String  'Stores second copy of file
    
    'Open file for reading
    Open "C:\WINDOWS\PBRUSH.EXE" For Binary As #1
    
        FileSize = LOF(1)  'Get file size
    
        'Size WholeFile string variable to the file size
        WholeFile = String(FileSize, " ")
    
        Get #1, , WholeFile  'Read entire file into memory
    
    Close #1  'Close file
    
    
    NumParts = (FileSize / 5000)  'Get number of parts in 5KB blocks

    ReDim FilePart(NumParts + 1)  'Dimension FilePart array
    
    'Split up file into 5KB blocks
    For i = 0 To NumParts
        FilePart(i) = Mid(WholeFile, i * 5000& + 1, 5000)
    Next i

    
    'Open up new file for writing
    Open "C:\TEMP\TEST.EXE" For Binary As #1
    
        'Put all the pieces together in one string
        For i = 0 To NumParts
            WholeFile2nd = WholeFile2nd + FilePart(i)
        Next i
        
        'Write entire file all at once
        Put #1, 1, WholeFile2nd
    
    Close #1  'Close file

    DoEvents  'Allow Windows to process
    
    'Print out debugging info
    Print "File size should be " & FileSize
    Print "New file size is " & FileLen("C:\TEMP\TEST.EXE")

End Sub

Private Sub cmdSendMsg_Click()
    Dim i As Integer  'For counter
    
    'Need to skip past errors since we could get the error '340'
    'because a socket didn't exist
    On Error Resume Next
    
    If txtSendMsg.Text = "" Then Exit Sub  'Nothing to send
    
    'If the program is hosting then send the message to everyone else
    If sckTCP(0).State = sckListening Then  'It is
        For i = 1 To sckTCP.UBound
            sckTCP(i).SendData ChatName + " says:  " + txtSendMsg.Text
            Call Pause  'Allow data to be sent
        Next i
        
        'Show message in chat text box
        txtChat.Text = txtChat.Text + vbCrLf + ChatName + " says:  " + txtSendMsg.Text
    Else  'Not hosting, just send message
        sckTCP(1).SendData txtSendMsg.Text
    End If

    txtSendMsg.Text = ""  'Clear out text box
End Sub

Private Sub Form_Load()
    'Write the instructions on the chat display text box
    txtChat.Text = "Click on the CONNECT button to connect to a host.  Click on the HOST button to host for connections.  You can copy the text in this display by selecting the text and pressing CTRL+C.  You can clear the text in this display by pressing the DEL key."

    'Color in text box to look disabled
    txtSendMsg.BackColor = frmMain.BackColor
    
    'Disable text box and send button
    txtSendMsg.Enabled = False
    cmdSendMsg.Enabled = False

    'Fill in default information
    sckTCP(0).RemoteHost = "localhost"
    sckTCP(0).RemotePort = "20000"
    sckTCP(0).LocalPort = "20000"
    ChatName = "Anonymous"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer  'For counter
    
    'Need to skip past errors since we could get the error '340'
    'because a socket didn't exist
    On Error Resume Next
    
    'Figure out which sockets are still open and close them all
    For i = sckTCP.LBound To sckTCP.UBound
        If sckTCP(i).State <> sckClosed Then  'Open socket
            sckTCP(i).Close  'Close the socket
            
            'Pause until socket closes
            Do While sckTCP(i).State <> sckClosed
                DoEvents  'Allow windows to continue
                
                'Check to make sure an error didn't occur
                If Not Err.Number = 0 Then  'Error occured
                    Exit Do  'Exit out of loop
                End If
            Loop
        End If
    Next i

    End  'Make sure program terminates
End Sub

Private Sub sckTCP_Close(Index As Integer)
    Dim i, j, k As Integer 'For counters
    Dim tmpData As String  'For holding temporary data
    
    'Need to skip past errors since we could get the error '340'
    'because a socket didn't exist
    On Error Resume Next
    
    If sckTCP(0).State = sckListening Then  'We are in host mode
        sckTCP(Index).Close  'Close socket
       
        'Determine who disconnected
        For i = 1 To sckTCP.Count
            If Connections(i).sckNum = Index Then  'Match between connection
                'Display user who left
                txtChat.Text = txtChat.Text + vbCrLf + Connections(i).ChatName + " has disconnected."
                
                'Tell everyone else who left
                For j = 1 To sckTCP.UBound
                    sckTCP(j).SendData Connections(i).ChatName + " has disconnected."
                    Call Pause  'Allow data to be sent
                Next j
                
                'Remove user from array by pushing all others down one
                For j = i + 1 To sckTCP.Count
                    Connections(j - 1).ChatName = Connections(j).ChatName
                    Connections(j - 1).sckNum = Connections(j).sckNum
                Next j
                
                'Remove extra entry at end
                Connections(sckTCP.Count).ChatName = ""
                Connections(sckTCP.Count).sckNum = ""
            
                'Remove name from list by rewriting entire list
                lstConnections.Clear
                
                For j = 0 To sckTCP.UBound
                    'Check to make sure not adding empty entry
                    If Not Connections(j).ChatName = "" Then  'Not empty
                        lstConnections.AddItem Connections(j).ChatName
                    End If
                Next j
                
                'Tell everyone current connections
                For j = 1 To sckTCP.UBound
                    'Send data for a list of all connections to clients
                    tmpData = "/~NameList@"
                    
                    For k = 0 To UBound(Connections)
                        If Not Connections(k).ChatName = "" Then  'Entry filled, add it
                            tmpData = tmpData + Connections(k).ChatName + "@"
                        End If
                    Next k
                    
                    sckTCP(j).SendData tmpData
                    Call Pause  'Allow data to be sent
                Next j
                
                Exit For  'Leave loop
            End If
        Next i
        
        Unload sckTCP(Index)  'Unload socket
        
        ReDim Preserve Connections(sckTCP.Count)  'Resize array
    Else
        'Host broke connection, tell user
        txtChat.Text = txtChat.Text + vbCrLf + "Host broke connection."
        
        sckTCP(1).Close  'Close connection
        Unload sckTCP(1)  'Unload socket
    
        txtSendMsg.Text = ""  'Clear out send text box
        txtSendMsg.BackColor = frmMain.BackColor  'Color send box to look disabled
        txtSendMsg.Enabled = False  'Disable text box
        cmdSendMsg.Enabled = False  'Disable send button
    
        lstConnections.Clear  'Clear out list of connections
        lstConnections.Enabled = False  'Disable list
        
        'Change Disconnect button caption
        cmdConnectDisconnect.Caption = "&Connect"
        
        cmdHostStopHost.Enabled = True  'Enable host button
    End If
End Sub

Private Sub sckTCP_Connect(Index As Integer)
    'Tell user that program connected
    txtChat.Text = txtChat.Text + vbCrLf + "Connected."
    
    'Send chat name to host
    sckTCP(1).SendData "/~Name@" + ChatName

    'Enable send text box and button
    txtSendMsg.Enabled = True
    cmdSendMsg.Enabled = True
    
    txtSendMsg.BackColor = vbWhite  'Make text box look enabled
End Sub

Private Sub sckTCP_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    Dim i As Integer  'For counter
    Dim FreeSocket As Integer  'For free socket number
    
    'Need to skip past errors since we could get the error '340'
    'because a socket didn't exist
    On Error Resume Next
    
    'Figure out what socket is avaiable if any
    For i = 1 To sckTCP.UBound + 1
        'Try to generate a '340 object doesn't exist error'
        'to find a free socket space
        
        FreeSocket = sckTCP(i).Index  'Set free socket
        
        'Check to see if a socket didn't exist so we
        'can make it exist and use it
        If Err.Number = 340 Then  'Control didn't exist
            FreeSocket = i  'Set free socket number
            Exit For  'Leave i loop
        End If
    Next i
            
    Load sckTCP(FreeSocket)  'Load new socket
    sckTCP(FreeSocket).LocalPort = 0  'Assign a random port number
    sckTCP(FreeSocket).Accept requestID  'Accept connection
    
    ReDim Preserve Connections(sckTCP.Count + 1) 'Resize array of connections
    
    Connections(sckTCP.Count).sckNum = FreeSocket  'Store connection socket number
    Connections(sckTCP.Count).ChatName = "Unconnected"  'Store the default name
End Sub

Private Sub sckTCP_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim IncomingData As String
    Dim i, j, k As Integer 'For counters
    Dim Start As Integer  'For position in string
    Dim tmpData As String  'For holding temporary data
    
    'Need to skip past errors since we could get the error '340'
    'because a socket didn't exist
    On Error Resume Next
    
    'Get incoming data
    sckTCP(Index).GetData IncomingData, vbString
    
    'Give a short pause to make sure data gets read
    Call Pause
        
    'Check to see if the host is getting information
    If sckTCP(0).State = sckListening Then  'Host is
        'Determine if it is the chat name of the client
        If Left$(IncomingData, 6) = "/~Name" Then  'Name is coming
            'Loop through all the connections
            For i = 0 To sckTCP.Count
                'Check to see if the current socket is in the array
                If Index = Connections(i).sckNum Then  'Found match
                    'Give that connection a name
                    Connections(i).ChatName = Mid$(IncomingData, 8, Len(IncomingData))
                    
                    'Tell everyone new connections name
                    For j = 1 To sckTCP.UBound
                        sckTCP(j).SendData Connections(i).ChatName + " has connected."
                        Call Pause  'Allow data to be sent
                        
                        'Send data for a list of all connections to clients
                        tmpData = "/~NameList@"
                        
                        For k = 0 To UBound(Connections)
                            If Not Connections(k).ChatName = "" Then  'Entry filled, add it
                                tmpData = tmpData + Connections(k).ChatName + "@"
                            End If
                        Next k
                        
                        sckTCP(j).SendData tmpData
                        Call Pause  'Allow data to be sent
                    Next j
                    
                    'Show who connected to user
                    txtChat.Text = txtChat.Text + vbCrLf + Connections(i).ChatName + " has connected."
                
                    'Clear out current list of people
                    lstConnections.Clear
                    
                    'Refresh entire list of people in chat
                    For j = 0 To sckTCP.Count
                        'Check to make sure not adding empty space
                        If Not Connections(j).ChatName = "" Then  'Not empty
                            lstConnections.AddItem Connections(j).ChatName
                        End If
                    Next j
                    
                    lstConnections.Enabled = True  'Enable list
                    
                    Exit For  'Leave loop
                End If
            Next i
        Else
            'Information is something someone typed, so display it
            'Loop through and figure out who said it
            For i = 0 To sckTCP.Count
                'Check for a match between socket numbers
                If Index = Connections(i).sckNum Then  'Match
                    'Display their name and message to all
                    txtChat.Text = txtChat.Text + vbCrLf + Connections(i).ChatName + " says:  " + IncomingData
                    
                    'Send data to everyone else
                    For j = 1 To sckTCP.UBound
                        sckTCP(j).SendData Connections(i).ChatName + " says:  " + IncomingData
                        Call Pause  'Allow data to be sent
                    Next j
                
                    Exit For  'Leave loop
                End If
            Next i
        End If
    Else  'Client is getting data
        If Left$(IncomingData, 11) = "/~NameList@" Then   'Receive list of people here
            lstConnections.Clear  'Clear out current list
            lstConnections.Enabled = True  'Enable list
            
            Start = 12  'Set position in string
            
            Do  'Loop until everything is read
                For i = Start To Len(IncomingData)
                    If Mid$(IncomingData, i, 1) = "@" Then
                        'Found a name, put it in the list
                        lstConnections.AddItem Mid$(IncomingData, Start, i - Start)

                        Start = i + 1  'Reset next position
                        
                        'Check to see if done reading everything
                        If Start > Len(IncomingData) Then Exit Do 'Done
                        
                        Exit For  'Leave for-loop
                    End If
                Next i
            Loop
        Else
            'Show incoming data
            txtChat.Text = txtChat.Text + vbCrLf + IncomingData
        End If
    End If
End Sub

Private Sub sckTCP_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    Dim i As Integer  'For counter
    
    'Need to skip past errors since we could get the error '340'
    'because a socket didn't exist
    On Error Resume Next
    
    'Tell user what the error was and that call connections are closed
    txtChat.Text = txtChat.Text + vbCrLf + "Error " & Number & ":  " & Description & "." + vbCrLf + "All connections closed."
    
    'Clear out send text box and color it to be disabled
    txtSendMsg.Text = ""
    txtSendMsg.Enabled = False
    txtSendMsg.BackColor = frmMain.BackColor

    cmdSendMsg.Enabled = False  'Disable send message button
    
    'Reset connect/disconnect and host/stop hosting buttons
    cmdConnectDisconnect.Caption = "&Connect"
    cmdHostStopHost.Caption = "&Host"
    cmdConnectDisconnect.Enabled = True
    cmdHostStopHost.Enabled = True
    
    sckTCP(0).Close  'Close primary socket
    
    'Close all other sockets and unload them
    For i = 1 To sckTCP.UBound
        sckTCP(i).Close
        Unload sckTCP(i)
    Next i
End Sub

Private Sub txtChat_Change()
    txtChat.SelStart = Len(txtChat.Text)
End Sub

Private Sub txtChat_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then  'Delete key was pressed
        txtChat.Text = ""  'Clear out text
    End If
End Sub

