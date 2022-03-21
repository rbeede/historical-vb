VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Automatic Multi-Program Opener/Closer"
   ClientHeight    =   4335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4575
   Icon            =   "Automatic Multi-Program OpenerCloser Main Form.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   4575
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   495
      Left            =   3120
      TabIndex        =   11
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close Them"
      Enabled         =   0   'False
      Height          =   495
      Left            =   1680
      TabIndex        =   10
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "&Open Them"
      Height          =   495
      Left            =   240
      TabIndex        =   9
      Top             =   3600
      Width           =   1215
   End
   Begin VB.TextBox txtApp 
      Height          =   285
      Index           =   7
      Left            =   240
      TabIndex        =   8
      Top             =   3000
      Width           =   4095
   End
   Begin VB.TextBox txtApp 
      Height          =   285
      Index           =   6
      Left            =   240
      TabIndex        =   7
      Top             =   2640
      Width           =   4095
   End
   Begin VB.TextBox txtApp 
      Height          =   285
      Index           =   5
      Left            =   240
      TabIndex        =   6
      Top             =   2280
      Width           =   4095
   End
   Begin VB.TextBox txtApp 
      Height          =   285
      Index           =   4
      Left            =   240
      TabIndex        =   5
      Top             =   1920
      Width           =   4095
   End
   Begin VB.TextBox txtApp 
      Height          =   285
      Index           =   3
      Left            =   240
      TabIndex        =   4
      Top             =   1560
      Width           =   4095
   End
   Begin VB.TextBox txtApp 
      Height          =   285
      Index           =   2
      Left            =   240
      TabIndex        =   3
      Top             =   1200
      Width           =   4095
   End
   Begin VB.TextBox txtApp 
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   4095
   End
   Begin VB.TextBox txtApp 
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   4095
   End
   Begin VB.Label lblApplications 
      AutoSize        =   -1  'True
      Caption         =   "Enter in the path and filename of the programs"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3960
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim AppIDs(8)

Private Sub cmdClose_Click()
    Dim i As Integer  'For counter
    
    'Close all the programs
    For i = 0 To 7
        If Not AppIDs(i) = 0 Then  'Program has id number
            AppActivate AppIDs(i)  'Activate program
            DoEvents  'Allow Windows to process
            SendKeys "%{F4}"  'Send ALT+F4 keystroke
            DoEvents  'Allow Windows to process
        End If
    Next i
        
    cmdClose.Enabled = False  'Disable this button
    cmdOpen.Enabled = True  'Enable open button
End Sub

Private Sub cmdExit_Click()
    Unload Me  'Unload form
End Sub

Private Sub cmdOpen_Click()
    Dim i As Integer  'For counter
    
    On Error Resume Next  'Skip over expected errors
    
    'Start programs
    For i = 0 To 7
        If Not txtApp(i).Text = "" Then  'Field isn't blank
            AppIDs(i) = Shell(txtApp(i).Text)    'Run program
            DoEvents  'Allow Windows to process
        End If
    Next i

    cmdOpen.Enabled = False  'Disable this button
    cmdClose.Enabled = True  'Enable close button
End Sub

Private Sub Form_Load()
    Dim Programs(8) As String  'Place to store program paths and filenames
    Dim i As Integer  'For counter
    
    On Error Resume Next  'Skip over expected errors
    
    'Try to open the initiation file
    Open App.Path + "\" + "Automatic Multi-Program OpenerCloser.INI" For Input As #1
    
    If Err.Number = 53 Then  'First time program has been run
        'Tell the user how to use the progam
        MsgBox "To use this program type in the directory paths and filenames of the programs you wish to automatically open and close.  Clicking on the Open Them button will open all the programs listed in the boxes.  Clicking on the Close Them button will close all the programs that were opened.  In order for the programs to close they must be able to be terminated by pressing ALT+F4.  Clicking on the Exit button ends the program and does not close any open programs.  You will not see this message the next time you start this program." + vbCrLf + vbCrLf + "Program made by Rodney Beede." + vbCrLf + "E-mail me at rodney_beede@hotmail.com" + vbCrLf + vbCrLf + "I am at no way responsible for any kind of loss so you use this program at your own risk.", vbOKOnly + vbInformation, "How to use this program"
    Else
        'Read out the program paths and filenames
        For i = 0 To 7
            Line Input #1, Programs(i)  'Store program
        Next i
    
        'Fill in text boxes with paths and filenames
        For i = 0 To 7
            txtApp(i).Text = Programs(i)
        Next i
    End If
    
    Close #1  'Close file
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer  'For counter
    
    On Error Resume Next  'Skip over expected errors
    
    'Attempt to open a file to save data
    Open App.Path + "\" + "Automatic Multi-Program OpenerCloser.INI" For Output As #1
    
    If Err.Number <> 0 Then  'Can't save file, tell user
        MsgBox "Unable to make file to save data, data will be lost.", vbExclamation + vbOKOnly, "Error Number " & Err.Number & ": " + Err.Description
        
        End  'Terminate program
    End If
    
    'Write program paths and filenames
    For i = 0 To 7
        Print #1, txtApp(i).Text
    Next i
    
    Close #1  'Close file
    
    End  'Terminate program
End Sub

Private Sub txtApp_GotFocus(Index As Integer)
    'Select everything in text box when focus is given
    txtApp(Index).SelStart = 0
    txtApp(Index).SelLength = Len(txtApp(Index).Text)
End Sub
