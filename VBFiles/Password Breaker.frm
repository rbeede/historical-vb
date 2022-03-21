VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Password Breaker"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7395
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   7395
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdBreak 
      Caption         =   "Break It!"
      Height          =   3255
      Left            =   5760
      TabIndex        =   1
      Top             =   600
      Width           =   1575
   End
   Begin VB.TextBox txtCodes 
      Height          =   4335
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "Password Breaker.frx":0000
      Top             =   0
      Width           =   5775
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBreak_Click()
    Dim i, j As Integer  'For counter
    Dim Password As String  'Password to break
    Dim PasswordPart(8) As String  'Array to hold password parts
    Dim Guess As String  'Holds entire guess
    Dim Fast As Integer  'Flag for speed
    Dim StartTimeDate As String  'Starting time & date
    Dim StopTimeDate As String  'Stop time & date
    Const MAXCHARS As Integer = 126  'Maximum charaters to go thorough
    
    'Ask for a password
    Password = InputBox$("Enter a password to break", "Password Breaker", "Password")
    
    If Password = "" Then Exit Sub  'User canceled
    
    'Ask if the user wants fast mode
    Fast = MsgBox("Do you want fast break?  Fast break will not display guesses but will run a lot faster.", vbYesNo, "Password Breaker")
    
    If Fast = vbNo Then
        Fast = False 'Set flag
        txtCodes.Text = "Displaying Guesses"  'Tell user no fast
    Else
        'Tell user no guesses shown
        txtCodes.Text = "Not Displaying Guesses"
    End If

    StartTimeDate = Time$ + " " + Date$  'Set start time & date
    Screen.MousePointer = vbArrowHourglass  'Set mouse pointer
    cmdBreak.Enabled = False  'Disable button
    Do  'Loop trying every possiblity of a password
        For i = 32 To MAXCHARS  'All possible characters
            PasswordPart(0) = Chr$(i)  'Set first position
            
            Guess = ""  'Clear out last guess
            
            'Assemble new part into a guess
            For j = 0 To 7
                Guess = Guess + PasswordPart(j)
            Next j
            
            If Fast = False Then  'Show guesses
                txtCodes.Text = txtCodes.Text + vbTab + Guess
                txtCodes.SelStart = Len(txtCodes.Text) - 1
            End If
            
            DoEvents  'Allow windows to go on

            'Check to see if their is a match
            If Guess = Password Then  'Match
                StopTimeDate = Time$ + " " + Date$  'Set stop time and date
                
                Screen.MousePointer = vbDefault  'Reset mouse pointer
                
                cmdBreak.Enabled = True
                
                'Tell user
                MsgBox "Your password is " + Guess + "." + vbCrLf + "Start time and date was " + StartTimeDate + " and stop time and date was " + StopTimeDate + ".", vbOKOnly + vbInformation, "Password Breaker"
                Exit Sub  'Leave subroutine
            End If
        Next i
        
        'Check to see if empty
        If PasswordPart(1) = "" Then
            PasswordPart(1) = Chr$(31)  'Get ready for push from nothing
        End If
        
        PasswordPart(1) = Chr$(Asc(PasswordPart(1)) + 1)  'Move to next part

        'Check for knock ups in next parts
        For i = 1 To 7
            
            If Asc(PasswordPart(i)) > MAXCHARS Then

                PasswordPart(i) = Chr$(32)  'Reset to bottom

                'Make sure that next push is not empty
                If PasswordPart(i + 1) = "" Then
                    PasswordPart(i + 1) = Chr$(31) 'Get ready for push from nothing
                End If
                
                PasswordPart(i + 1) = Chr$(Asc(PasswordPart(i + 1)) + 1) 'Push up one
            Else
                Exit For  'Don't need to knock up next section
            End If
        Next i
    
        If Not PasswordPart(8) = "" Then
            StopTimeDate = Time$ + " " + Date$  'Set stop time and date
            
            cmdBreak.Enabled = True  'Enable button
            
            Screen.MousePointer = vbDefault  'Reset mouse pointer
            
            MsgBox "Could not break password." + vbCrLf + "Start time and date was " + StartTimeDate + " and stop time and date was " + StopTimeDate + ".", vbExclamation + vbOKOnly, "Password Breaker"
            Exit Sub
        End If
    Loop
End Sub

Private Sub Form_Load()
    'Size form and controls
    Me.Height = Screen.Height
    Me.Width = Screen.Width
    txtCodes.Width = Me.Width
    txtCodes.Height = Me.Height - 2000
    cmdBreak.Top = txtCodes.Top + txtCodes.Height
    cmdBreak.Left = 0
    cmdBreak.Width = txtCodes.Width
    cmdBreak.Height = Me.Height - txtCodes.Height - 300
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End  'Terminate program
End Sub
