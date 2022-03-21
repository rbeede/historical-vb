VERSION 2.00
Begin Form frmInterCom 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Double
   Caption         =   "InterCom"
   ClientHeight    =   1830
   ClientLeft      =   3840
   ClientTop       =   1485
   ClientWidth     =   2055
   ControlBox      =   0   'False
   Height          =   2235
   Left            =   3780
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1830
   ScaleWidth      =   2055
   Top             =   1140
   Width           =   2175
   Begin SSCommand cmdClose 
      BevelWidth      =   1
      Caption         =   "&Close InterCom"
      Font3D          =   1  'Raised w/light shading
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   1575
   End
   Begin SSCommand cmdOnOff 
      BevelWidth      =   1
      Caption         =   "InterCom &On"
      Font3D          =   1  'Raised w/light shading
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1575
   End
End

Sub cmdClose_Click ()
    cmdOnOff.Caption = "Intercom &On" 'Change caption
    Me.Enabled = False 'Disable form
    frmMain.Enabled = True 'Enable main form
    frmInterCom.Hide 'Hide intercom form
    Unload frmInterCom 'Make sure it unloads
End Sub

Sub cmdOnOff_Click ()
    Dim Response As Integer 'Declare varialbe

    'Check to see what caption of button says
    If cmdOnOff.Caption = "InterCom &On" Then 'Open up intercom
       
       'Make sure they want to do this
       Response = MsgBox("Click on OK and speak into the microphone.", 49, "Intercom")

       If Response = 2 Then Exit Sub 'They don't

       Call PlaySounds("Record", "Recording") 'Open intercom
    
       cmdOnOff.Caption = "Intercom O&ff" 'Change caption
       cmdClose.Enabled = False 'Disable close button
    Else
       Call PlaySounds("Dummy", "Stoping")
       Call PlaySounds("Dummy", "Saving")

       cmdOnOff.Caption = "InterCom &On"
       cmdClose.Enabled = True
    End If
End Sub

