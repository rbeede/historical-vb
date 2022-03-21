VERSION 5.00
Begin VB.Form frmABC 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "THE ALPHABET ABC's"
   ClientHeight    =   6540
   ClientLeft      =   1080
   ClientTop       =   1755
   ClientWidth     =   9510
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "Alphabet.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6540
   ScaleWidth      =   9510
   Begin VB.Label lblletter 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Z"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   585
      Index           =   25
      Left            =   1080
      TabIndex        =   4
      Top             =   480
      Width           =   360
   End
   Begin VB.Label lblletter 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Y"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   585
      Index           =   24
      Left            =   5280
      TabIndex        =   5
      Top             =   1320
      Width           =   390
   End
   Begin VB.Label lblletter 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   585
      Index           =   23
      Left            =   4560
      TabIndex        =   6
      Top             =   480
      Width           =   375
   End
   Begin VB.Label lblletter 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "W"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   585
      Index           =   22
      Left            =   3960
      TabIndex        =   7
      Top             =   1320
      Width           =   525
   End
   Begin VB.Label lblletter 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "V"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   585
      Index           =   21
      Left            =   3480
      TabIndex        =   8
      Top             =   480
      Width           =   390
   End
   Begin VB.Label lblletter 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "U"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   585
      Index           =   20
      Left            =   3360
      TabIndex        =   9
      Top             =   1320
      Width           =   405
   End
   Begin VB.Label lblletter 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "T"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   585
      Index           =   19
      Left            =   2280
      TabIndex        =   10
      Top             =   1320
      Width           =   360
   End
   Begin VB.Label lblletter 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "S"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   585
      Index           =   18
      Left            =   1680
      TabIndex        =   11
      Top             =   1320
      Width           =   375
   End
   Begin VB.Label lblletter 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "R"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   585
      Index           =   17
      Left            =   1680
      TabIndex        =   12
      Top             =   480
      Width           =   390
   End
   Begin VB.Label lblletter 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Q"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   585
      Index           =   16
      Left            =   2280
      TabIndex        =   13
      Top             =   480
      Width           =   435
   End
   Begin VB.Label lblletter 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "P"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   585
      Index           =   15
      Left            =   7680
      TabIndex        =   14
      Top             =   480
      Width           =   375
   End
   Begin VB.Label lblletter 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   585
      Index           =   14
      Left            =   8280
      TabIndex        =   15
      Top             =   480
      Width           =   435
   End
   Begin VB.Label lblletter 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "N"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   585
      Index           =   13
      Left            =   8880
      TabIndex        =   16
      Top             =   480
      Width           =   420
   End
   Begin VB.Label lblletter 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "M"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   585
      Index           =   12
      Left            =   7080
      TabIndex        =   17
      Top             =   480
      Width           =   450
   End
   Begin VB.Label lblletter 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "L"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   585
      Index           =   11
      Left            =   5520
      TabIndex        =   18
      Top             =   480
      Width           =   330
   End
   Begin VB.Label lblletter 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "K"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   585
      Index           =   10
      Left            =   6000
      TabIndex        =   19
      Top             =   480
      Width           =   375
   End
   Begin VB.Label lblletter 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "J"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   585
      Index           =   9
      Left            =   6600
      TabIndex        =   20
      Top             =   480
      Width           =   300
   End
   Begin VB.Label lblletter 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "I"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   585
      Index           =   8
      Left            =   5160
      TabIndex        =   21
      Top             =   480
      Width           =   180
   End
   Begin VB.Label lblletter 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "H"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   585
      Index           =   7
      Left            =   4680
      TabIndex        =   22
      Top             =   1320
      Width           =   405
   End
   Begin VB.Label lblletter 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "G"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   585
      Index           =   6
      Left            =   3960
      TabIndex        =   23
      Top             =   480
      Width           =   435
   End
   Begin VB.Label lblletter 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "F"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   585
      Index           =   5
      Left            =   2760
      TabIndex        =   24
      Top             =   1320
      Width           =   360
   End
   Begin VB.Label lblletter 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "E"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   585
      Index           =   4
      Left            =   2880
      TabIndex        =   25
      Top             =   480
      Width           =   375
   End
   Begin VB.Label lblinstruct 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      Caption         =   "lblinstruct"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   300
      Left            =   0
      TabIndex        =   28
      Top             =   6250
      Width           =   9615
   End
   Begin VB.Label lblword 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "lblword"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   480
      TabIndex        =   27
      Top             =   5760
      Width           =   1320
   End
   Begin VB.Image imgclear 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   8160
      Picture         =   "Alphabet.frx":030A
      Top             =   2520
      Width           =   480
   End
   Begin VB.Image imgstop 
      Appearance      =   0  'Flat
      Height          =   735
      Left            =   8040
      Picture         =   "Alphabet.frx":0614
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   735
   End
   Begin VB.Label lblletters 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "lblletters"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   29.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   2895
      Left            =   480
      TabIndex        =   26
      Top             =   2520
      Width           =   7365
   End
   Begin VB.Label lblletter 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "D"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   585
      Index           =   3
      Left            =   480
      TabIndex        =   3
      Top             =   1320
      Width           =   405
   End
   Begin VB.Label lblletter 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   585
      Index           =   2
      Left            =   1080
      TabIndex        =   2
      Top             =   1320
      Width           =   405
   End
   Begin VB.Label lblletter 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   585
      Index           =   1
      Left            =   5880
      TabIndex        =   1
      Top             =   1320
      Width           =   375
   End
   Begin VB.Label lblletter 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   585
      Index           =   0
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   390
   End
   Begin VB.Menu mnustart 
      Caption         =   "&Start"
   End
   Begin VB.Menu mnustop 
      Caption         =   "S&top"
   End
   Begin VB.Menu mnucheat 
      Caption         =   "&Cheat"
   End
   Begin VB.Menu mnuabout 
      Caption         =   "&About"
   End
   Begin VB.Menu mnuend 
      Caption         =   "&Exit"
   End
End
Attribute VB_Name = "frmABC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim QuitFlag As Integer
Dim CursorCount As Integer
Private Declare Function ShowCursor Lib "User32" (ByVal bShow As Integer) As Integer

Private Sub form_keydown(KeyCode As Integer, Shift As Integer)
If KeyCode = &H1B Then
       Call HideMouse(False)
       en = MsgBox("Do you wish to end this program?", 36, "The Alphabet ABC's")
       If en = 7 Then
          
       Else
          Unload Me
       End If
End If
End Sub

Private Sub Form_Load()
    Left = 0
    Top = 0
    
    Height = Screen.Height
    Width = Screen.Width
    
    lblinstruct.Top = Me.Height - lblinstruct.Height - 650
    lblinstruct.Width = Me.Width
    
    If App.PrevInstance = True Then
       Unload Me
       Exit Sub
    End If
    lblletters.Visible = False
    lblword.Visible = False
    lblinstruct.Caption = "Click on start menu to begin."
    imgstop.Visible = False
    imgclear.Visible = False
    mnustop.Enabled = False
    mnucheat.Enabled = False
    MousePointer = 10
    Screen.MousePointer = 10
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Call HideMouse(False)
End Sub

Private Sub HideMouse(ShowHide As Integer)
   CursorCount = 1
    
   If ShowHide = True Then
    Do While ShowCursor(False) >= -1
    Loop
    Do While ShowCursor(True) < -1
    Loop
   Else
    Do While ShowCursor(False) >= CursorCount
    Loop
    Do While ShowCursor(True) < CursorCount
    Loop
   End If
End Sub

Private Sub imgclear_Click()
    If MsgBox("Do you wish to erase letters?", 17 + 256, "Erase Letters") = 1 Then
       lblletters.Caption = ""
    End If
End Sub

Private Sub imgstop_Click()
    Call StopCheck
End Sub

Private Sub lblletter_Click(Index As Integer)
    Select Case Index
      Case 0
       lblletters.Caption = lblletters.Caption + "A "
      Case 1
       lblletters.Caption = lblletters.Caption + "B "
      Case 2
       lblletters.Caption = lblletters.Caption + "C "
      Case 3
       lblletters.Caption = lblletters.Caption + "D "
      Case 4
       lblletters.Caption = lblletters.Caption + "E "
      Case 5
       lblletters.Caption = lblletters.Caption + "F "
      Case 6
       lblletters.Caption = lblletters.Caption + "G "
      Case 7
       lblletters.Caption = lblletters.Caption + "H "
      Case 8
       lblletters.Caption = lblletters.Caption + "I "
      Case 9
       lblletters.Caption = lblletters.Caption + "J "
      Case 10
       lblletters.Caption = lblletters.Caption + "K "
      Case 11
       lblletters.Caption = lblletters.Caption + "L "
      Case 12
       lblletters.Caption = lblletters.Caption + "M "
      Case 13
       lblletters.Caption = lblletters.Caption + "N "
      Case 14
       lblletters.Caption = lblletters.Caption + "O "
      Case 15
       lblletters.Caption = lblletters.Caption + "P "
      Case 16
       lblletters.Caption = lblletters.Caption + "Q "
      Case 17
       lblletters.Caption = lblletters.Caption + "R "
      Case 18
       lblletters.Caption = lblletters.Caption + "S "
      Case 19
       lblletters.Caption = lblletters.Caption + "T "
      Case 20
       lblletters.Caption = lblletters.Caption + "U "
      Case 21
       lblletters.Caption = lblletters.Caption + "V "
      Case 22
       lblletters.Caption = lblletters.Caption + "W "
      Case 23
       lblletters.Caption = lblletters.Caption + "X "
      Case 24
       lblletters.Caption = lblletters.Caption + "Y "
      Case 25
       lblletters.Caption = lblletters.Caption + "Z "
    End Select
End Sub

Private Sub mnuabout_Click()
    Dim Msg As String

    Msg = "Alphabet ABC's" + Chr$(13)
    Msg = Msg + "Programed and designed by Rodney Beede." + Chr$(13)
    Msg = Msg + "Published by Infinisoft." + Chr$(13)
    Msg = Msg + "E-Mail me at rodney_beede@hotmail.com." + Chr$(13)
    Msg = Msg + "Read the Alphabet.txt file for more information"
    
    MsgBox Msg, 64, "About Alphabet ABC's"
End Sub

Private Sub mnucheat_Click()
Dim Msg As String
    Msg = "The alphabet is: " + Chr$(13)
    Msg = Msg + "A B C D E F H I J K L M N O P Q R S T U V W X Y Z."
    MsgBox Msg, 48, "Cheat!"
End Sub

Private Sub mnuend_Click()
    Call form_keydown(&H1B, 0)
End Sub

Private Sub mnustart_Click()
    lblword.Caption = "Click the letters of the alphabet in order."
    lblinstruct.Caption = "Click stop when done.  Click the eraser to erase letters."
    lblletters.Caption = ""
    lblletters.Visible = True
    lblinstruct.Visible = True
    lblword.Visible = True
    imgstop.Visible = True
    imgclear.Visible = True
    mnustop.Enabled = True
    mnucheat.Enabled = True
    mnustart.Enabled = False
End Sub

Private Sub mnustop_Click()
    lblletters.Visible = False
    lblword.Visible = False
    lblinstruct.Caption = "Click on start menu to begin."
    imgstop.Visible = False
    imgclear.Visible = False
    mnustop.Enabled = False
    mnucheat.Enabled = False
    mnustart.Enabled = True
    MousePointer = 10
    Screen.MousePointer = 10
End Sub

Private Sub StopCheck()
    If InStr(lblword.Caption, "alphabet") Then
       If lblletters.Caption = "A B C D E F G H I J K L M N O P Q R S T U V W X Y Z " Then
          MsgBox "You are correct!", 48, "Good Job"
       Else
          MsgBox "You are wrong!", 48, "Please try again"
       End If
    End If
    Call mnustop_Click
End Sub

