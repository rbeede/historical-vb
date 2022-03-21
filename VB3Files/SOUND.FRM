VERSION 2.00
Begin Form frmSound 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Double
   Caption         =   "The Sound Player"
   ClientHeight    =   2625
   ClientLeft      =   2430
   ClientTop       =   2640
   ClientWidth     =   2175
   ControlBox      =   0   'False
   Height          =   3030
   Icon            =   SOUND.FRX:0000
   Left            =   2370
   LinkTopic       =   "Form1"
   ScaleHeight     =   2625
   ScaleWidth      =   2175
   Top             =   2295
   Width           =   2295
   Begin CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   495
      Left            =   480
      TabIndex        =   2
      Top             =   1920
      Width           =   1335
   End
   Begin CommandButton cmdSelect 
      Caption         =   "&Select Sound"
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   1200
      Width           =   1335
   End
   Begin CommandButton cmdPlay 
      Caption         =   "&Play Sound"
      Enabled         =   0   'False
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   1335
   End
   Begin Image Image1 
      Height          =   480
      Left            =   0
      Picture         =   SOUND.FRX:0302
      Top             =   1920
      Width           =   480
   End
   Begin Image imgBell 
      Height          =   480
      Left            =   0
      Picture         =   SOUND.FRX:0604
      Top             =   1200
      Width           =   480
   End
   Begin Image imgSpeaker 
      Height          =   330
      Left            =   0
      Picture         =   SOUND.FRX:0906
      Top             =   600
      Width           =   360
   End
   Begin Label lblWavFile 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   75
   End
End

Sub cmdExit_Click ()
    End ' End program
End Sub

Sub cmdPlay_Click ()
    ' Declare variables used to get response from user, to use in counter,
    ' and use in message box
    Dim Response As Integer, Times As Integer
    Dim Msg As String, CRLF As String, Continue As Integer
    
    ' Ask user how many times to play sound
    Response = Val(InputBox("Type number of times to play sound.", "Play Sound", "1"))

    ' Check answer
    If Response = Null Then Exit Sub ' User canceled
    If Response = 0 Then Exit Sub ' User did not want to play any sounds
    ' Program can play that many times but will still ask user if ok to wait
    ' that long
    If Response > 10 Then
       
       CRLF = Chr$(13) + Chr$(10) ' Character Line Feed
       
       ' Setup message
       Msg = "Time to play sound over ten times will take" + CRLF
       Msg = Msg + "some time to finish before control is returned." + CRLF
       Msg = Msg + "Do you wish to continue?"
       
       ' Ask message
       Continue = MsgBox(Msg, 308, "Warning")
    
       ' See what the response was
       If Continue = 7 Then Exit Sub ' User said no
    End If
    
    ' Start counter to play sound correct number of times
    For Times = 1 To Response
        
        ' Call playsound to play sound
        Call PlaySound(PathWavFile)
    
    ' Go to start of counter and add one untill number of reponses is reached
    Next Times
End Sub

Sub cmdSelect_Click ()
    ' Call function to get windows directory
    WindowsDir = WindowsDirectory()
    
    Me.Enabled = False ' Disable this form
    frmFileOpen.Show ' Show file form

    ' Set Sound File default path to windows directory
    frmFileOpen.SoundDir.Path = WindowsDir
End Sub

Sub Form_Load ()
Show ' Display stupid form
End Sub

Sub lblWavFile_Change ()
    ' User chose a file enable play button for use
    cmdPlay.Enabled = True
End Sub

