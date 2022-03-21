VERSION 2.00
Begin Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Modem Calculator"
   ClientHeight    =   3030
   ClientLeft      =   1095
   ClientTop       =   1485
   ClientWidth     =   5415
   Height          =   3435
   Icon            =   CALMODEM.FRX:0000
   Left            =   1035
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   5415
   Top             =   1140
   Width           =   5535
   Begin CommandButton Command2 
      Caption         =   "End"
      Height          =   495
      Left            =   2760
      TabIndex        =   3
      Top             =   2280
      Width           =   2415
   End
   Begin CommandButton Command1 
      Caption         =   "Calculate"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   2280
      Width           =   2415
   End
   Begin TextBox Text2 
      Height          =   285
      Left            =   1920
      MaxLength       =   30
      TabIndex        =   1
      Top             =   720
      Width           =   3255
   End
   Begin TextBox Text1 
      Height          =   285
      Left            =   3120
      MaxLength       =   7
      TabIndex        =   0
      Top             =   240
      Width           =   855
   End
   Begin Label Label3 
      Height          =   855
      Left            =   240
      TabIndex        =   6
      Top             =   1200
      Width           =   4935
   End
   Begin Label Label2 
      AutoSize        =   -1  'True
      Caption         =   " Modem Bits Per Second Speed: "
      Height          =   195
      Left            =   240
      TabIndex        =   5
      Top             =   240
      Width           =   2820
   End
   Begin Label Label1 
      AutoSize        =   -1  'True
      Caption         =   " File Size in Bytes: "
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   720
      Width           =   1650
   End
End
Sub Command1_Click ()
Dim Bytes_a_Second As Double
Dim Msg As String
Dim Result As Double

On Error Resume Next

Bytes_a_Second = Text1.Text / 8
Result = (Text2.Text / Bytes_a_Second) / 60

Msg = "The time needed to move a file the size of "
Msg = Msg + Format$(Text2.Text, "###,###,###,###,###,###,###,###,###,###")
Msg = Msg + " bytes when at a speed of " & Bytes_a_Second
Msg = Msg + " bytes a second is " & Result & " minutes."
Label3.Caption = Msg

If Err Then Label3.Caption = "Error Calculating"

End Sub

Sub Command2_Click ()
End
End Sub

