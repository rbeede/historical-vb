VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Modem Calculator"
   ClientHeight    =   3030
   ClientLeft      =   1095
   ClientTop       =   1485
   ClientWidth     =   5415
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
   Icon            =   "CALULATE MODEM SPEED.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3030
   ScaleWidth      =   5415
   Begin VB.CommandButton cmdMB 
      Caption         =   "1 MB"
      Height          =   255
      Left            =   4560
      TabIndex        =   7
      Top             =   720
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      Caption         =   "End"
      Height          =   495
      Left            =   2760
      TabIndex        =   3
      Top             =   2280
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "Calculate"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   2280
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1920
      MaxLength       =   30
      TabIndex        =   1
      Top             =   720
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3120
      TabIndex        =   0
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   240
      TabIndex        =   6
      Top             =   1200
      Width           =   4935
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   " Modem Bits Per Second Speed: "
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   240
      TabIndex        =   5
      Top             =   240
      Width           =   2820
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   " File Size in Bytes: "
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   720
      Width           =   1650
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdMB_Click()
    Text2.Text = Str$(2 ^ 20)
End Sub

Private Sub Command1_Click()
Dim Bytes_a_Second As Double
Dim Msg As String
Dim ResultSecs As Double
Dim ResultMins As Double
Dim ResultHours As Double

On Error Resume Next

Text1.Text = Format$(Text1.Text, "#,###")
Text2.Text = Format$(Text2.Text, "#,###")

Bytes_a_Second = Text1.Text / 8
ResultSecs = Format(Text2.Text / Bytes_a_Second, "0,000.00")
ResultMins = Format(ResultSecs / 60, "0,000.00")
ResultHours = Format(ResultMins / 60, "0,000.00")

Msg = "The time needed to move a file the size of "
Msg = Msg + Format$(Text2.Text, "#,###")
Msg = Msg + " bytes when at a speed of " & Format(Bytes_a_Second, "#,###")
Msg = Msg + " bytes a second is " & ResultMins & " minutes or " & ResultHours & " hours."
Label3.Caption = Msg

If Err Then Label3.Caption = "Error Calculating"

End Sub

Private Sub Command2_Click()
End
End Sub

