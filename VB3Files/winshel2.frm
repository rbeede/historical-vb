VERSION 2.00
Begin Form frmSecurity 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Double
   Caption         =   "Security Check"
   ClientHeight    =   1590
   ClientLeft      =   1110
   ClientTop       =   1500
   ClientWidth     =   3495
   ControlBox      =   0   'False
   Height          =   1995
   Left            =   1050
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1590
   ScaleWidth      =   3495
   Top             =   1155
   Width           =   3615
   Begin CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   1800
      TabIndex        =   2
      Top             =   840
      Width           =   1455
   End
   Begin CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   1455
   End
   Begin TextBox txtPassword 
      Height          =   285
      Left            =   120
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   480
      Width           =   3135
   End
   Begin Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Please enter the password:"
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   12
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   300
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   3285
   End
End
Option Explicit

Sub cmdCancel_Click ()
    Passed = False  'Set flag
    txtPassword.Text = ""  'Clear out password box
    Me.Hide  'Hide this form
End Sub

Sub cmdOK_Click ()
    Passed = False  'Clear out flag
    
    'Check the password
    If txtPassword.Text = Password Then
        'Valid password, allow user to go on
        Passed = True  'Set flag
        txtPassword.Text = ""  'Clear out password box
        Me.Hide  'Hide this form
    Else
        'Tell user invalid password
        MsgBox "Invalid password.", MB_ICONEXCLAMATION, "Error"
        txtPassword.Text = ""  'Clear out password box
        txtPassword.SetFocus  'Give focus to password box
    End If
End Sub

