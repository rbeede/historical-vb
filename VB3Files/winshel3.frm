VERSION 2.00
Begin Form frmProgramManager 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Program Manager"
   ClientHeight    =   1260
   ClientLeft      =   1095
   ClientTop       =   1485
   ClientWidth     =   3375
   Height          =   1665
   Icon            =   WINSHEL3.FRX:0000
   Left            =   1035
   LinkMode        =   1  'Source
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1260
   ScaleWidth      =   3375
   Top             =   1140
   Width           =   3495
   WindowState     =   1  'Minimized
   Begin Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "This isn't really Program Manager.  This was placed here by Windows Shell for compatibility with other programs."
      Height          =   855
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3255
   End
End
Option Explicit

Sub Form_QueryUnload (Cancel As Integer, UnloadMode As Integer)
    'Don't allow this program to end unless Windows is shutting down
    If UnloadMode <> 2 Then
        'Tell user they can't unload this form
        MsgBox "You may not close this Window.  Minimize it instead.", MB_ICONEXCLAMATION, "Windows Shell"

        Cancel = True  'Stop form from unloading
    End If
End Sub

