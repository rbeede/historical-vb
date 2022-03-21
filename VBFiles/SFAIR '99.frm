VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3825
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5430
   LinkTopic       =   "Form1"
   ScaleHeight     =   3825
   ScaleWidth      =   5430
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   1200
      TabIndex        =   0
      Top             =   120
      Width           =   3735
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   1200
      TabIndex        =   1
      Top             =   480
      Width           =   3735
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   2
      Left            =   1200
      TabIndex        =   2
      Top             =   840
      Width           =   3735
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   3
      Left            =   1200
      TabIndex        =   3
      Top             =   1200
      Width           =   3735
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   4
      Left            =   1200
      TabIndex        =   4
      Top             =   1560
      Width           =   3735
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   5
      Left            =   1200
      TabIndex        =   5
      Top             =   1920
      Width           =   3735
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   6
      Left            =   1200
      TabIndex        =   6
      Top             =   2280
      Width           =   3735
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   7
      Left            =   1200
      TabIndex        =   7
      Top             =   2640
      Width           =   3735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add"
      Height          =   495
      Left            =   1200
      TabIndex        =   8
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Save"
      Height          =   495
      Left            =   3720
      TabIndex        =   9
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Category"
      Height          =   195
      Left            =   450
      TabIndex        =   17
      Top             =   120
      Width           =   630
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Question"
      Height          =   195
      Left            =   450
      TabIndex        =   16
      Top             =   480
      Width           =   630
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Answer A"
      Height          =   195
      Left            =   405
      TabIndex        =   15
      Top             =   840
      Width           =   675
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Answer B"
      Height          =   195
      Left            =   405
      TabIndex        =   14
      Top             =   1200
      Width           =   675
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Answer C"
      Height          =   195
      Left            =   405
      TabIndex        =   13
      Top             =   1560
      Width           =   675
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Answer D"
      Height          =   195
      Left            =   390
      TabIndex        =   12
      Top             =   1920
      Width           =   690
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Value"
      Height          =   195
      Left            =   675
      TabIndex        =   11
      Top             =   2280
      Width           =   405
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Correct Answer"
      Height          =   195
      Left            =   0
      TabIndex        =   10
      Top             =   2640
      Width           =   1080
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type ProbDef
    Answers(3) As String  'Array of multiple choice answers
    Category As String  'Question category
    CorrectAnswer As Byte  'The correct answer
    Question As String  'Question (or problem)
    Value As Integer  'Amount of money for correct answer
End Type

Dim Problems() As ProbDef

Dim ArraySize As Integer

Private Sub Command1_Click()
    Dim i As Byte
    
    ArraySize = ArraySize + 1
    
    ReDim Preserve Problems(ArraySize)
    
    Problems(ArraySize).Category = Text1(0).Text
    Problems(ArraySize).Question = Text1(1).Text
    For i = 2 To 5
        Problems(ArraySize).Answers(i - 2) = Text1(i).Text
    Next i
    Problems(ArraySize).Value = Val(Text1(6).Text)
    Problems(ArraySize).CorrectAnswer = Val(Text1(7).Text)

    For i = 0 To 7
        Text1(i).Text = ""
    Next i
    
End Sub

Private Sub Command2_Click()
    Dim i As Integer
    
    Open "c:\temp\problems.dat" For Random As #1 Len = 500
    
    For i = 0 To UBound(Problems)
        Put #1, , Problems(i)
    Next i
        
    Close #1
End Sub

Private Sub Form_Load()
    ArraySize = -1
End Sub
