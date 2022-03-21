VERSION 2.00
Begin Form frmQuestions 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Double
   Caption         =   "Answer the problem."
   ClientHeight    =   1455
   ClientLeft      =   1095
   ClientTop       =   1515
   ClientWidth     =   6525
   ControlBox      =   0   'False
   Height          =   1860
   Left            =   1035
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1455
   ScaleWidth      =   6525
   Top             =   1170
   Width           =   6645
   Begin CommandButton cmdEnd 
      Caption         =   "&Quit"
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   24
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   1095
      Left            =   5160
      TabIndex        =   5
      Top             =   240
      Width           =   1095
   End
   Begin CommandButton cmdCalculator 
      Caption         =   "&Calculator"
      Height          =   495
      Left            =   3600
      TabIndex        =   4
      Top             =   240
      Width           =   1095
   End
   Begin TextBox txtAnswer 
      Height          =   495
      Left            =   2280
      MaxLength       =   9
      TabIndex        =   1
      Top             =   240
      Width           =   1095
   End
   Begin Timer tmrQuestionTime 
      Interval        =   1000
      Left            =   2280
      Top             =   960
   End
   Begin Label lblWrong 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Answer incorrect: #"
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "Times New Roman"
      FontSize        =   12
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   285
      Left            =   2880
      TabIndex        =   3
      Top             =   960
      Width           =   1950
   End
   Begin Label lblRight 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Answers correct:  #"
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "Times New Roman"
      FontSize        =   12
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   1965
   End
   Begin Label lblProblem 
      Caption         =   "2 + 2 = X"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1935
   End
End
'Declare variables
Dim Problem As Integer 'For which problem is being done
Dim SecondsLeft As Integer 'For storing how many seconds are left
Dim Numbera As Integer 'For storing random number one
Dim Numberb As Integer 'For storing random number two
Dim Numberc As Integer 'For storing random number three
Dim Correct As Integer 'For storing how many problems are correct
'For storing how many problems are incorrect
Dim Incorrect As Integer

Sub Check_Answer ()
    
    If SecondsLeft <= 0 Then 'Time has run out
       'Add one to number of shots computer gets
       PCNumberShots = PCNumberShots + 1
       Incorrect = Incorrect + 1 'Update how many are incorrect
       'Update label for how many are incorrect
       lblWrong.Caption = "Answer's incorrect: " + Str$(Incorrect)
       
       'Tell user time ran out
       MsgBox "Out of time!", 48, "Time ran out"
       
    Else
       'Makes answer correct, for debugging
       If txtAnswer = "Bank_Shot" Then
          Correct = Correct + 1
          MsgBox "Correct = Correct + 1", , "OK RODNEY"
          GoTo ByPass
       End If
       
       Select Case Problem 'Determine which problem is being done
          Case 1
             'Check problem
             If txtAnswer = Format$((Numbera + Numberb) / Numberc, "######.##") Then
                'Add one to number correct
                Correct = Correct + 1
                MsgBox "Correct!", 48, "GOOD" 'Tell user answer was correct
             Else
                'Add one to number incorrect
                Incorrect = Incorrect + 1
                'Tell user answer was wrong
                MsgBox "Incorrect!", 48, "BETTER LUCK NEXT TIME"
             End If
          Case 2
             If txtAnswer = Format$((Numbera + Numberb) / Numberc, "######.##") Then
                'Add one to number correct
                Correct = Correct + 1
                MsgBox "Correct!", 48, "GOOD" 'Tell user answer was correct
             Else
                'Add one to number incorrect
                Incorrect = Incorrect + 1
                'Tell user answer was wrong
                MsgBox "Incorrect!", 48, "BETTER LUCK NEXT TIME"
             End If
          Case 3
             If txtAnswer = Format$((Numberc * Numberb) - Numbera, "######.##") Then
                'Add one to number correct
                Correct = Correct + 1
                MsgBox "Correct!", 48, "GOOD" 'Tell user answer was correct
             Else
                'Add one to number incorrect
                Incorrect = Incorrect + 1
                'Tell user answer was wrong
                MsgBox "Incorrect!", 48, "BETTER LUCK NEXT TIME"
             End If
          Case 4
             If txtAnswer = Format$(Sqr(Numbera), "######.##") Then
                'Add one to number correct
                Correct = Correct + 1
                MsgBox "Correct!", 48, "GOOD" 'Tell user answer was correct
             Else
                'Add one to number incorrect
                Incorrect = Incorrect + 1
                'Tell user answer was wrong
                MsgBox "Incorrect!", 48, "BETTER LUCK NEXT TIME"
             End If
          Case 5
             If txtAnswer = Format$(Numbera * (Numberb - Numberc), "######.##") Then
                'Add one to number correct
                Correct = Correct + 1
                MsgBox "Correct!", 48, "GOOD" 'Tell user answer was correct
             Else
                'Add one to number incorrect
                Incorrect = Incorrect + 1
                'Tell user answer was wrong
                MsgBox "Incorrect!", 48, "BETTER LUCK NEXT TIME"
             End If
       
       End Select
    End If

ByPass: 'Line label
    
    If Correct + Incorrect = 5 Then 'Out of questions
       MyNumberShots = Correct 'Set number of shots for user
       PCNumberShots = Incorrect 'Set number of shots for pc
       frmMain.Enabled = True 'Enable main form
       'Update labels
       frmMain!lblMyShots.Caption = "Shots you get: " & Correct
       frmMain!lblPCShots.Caption = "Shots I get: " & Incorrect
       frmMain!lblMyShots.Visible = True
       frmMain!lblPCShots.Visible = True
       Shooting = True 'Say shooting
       'Update label
       frmMain!lblInstructions.Caption = "Click here to begin"
       frmMain.Refresh 'Forces redraw on main form
       Unload frmQuestions 'Unload this form
       Exit Sub 'Leave this sub
    Else
       'Update labels
       lblRight.Caption = "Answers correct: " & Correct
       lblWrong.Caption = "Answers incorrect: " & Incorrect
       
       'Reset time left according to skill
       If Skill = 0 Then 'Easy skill
          SecondsLeft = 180
       ElseIf Skill = 1 Then 'Normal Skill
          SecondsLeft = 120
       ElseIf Skill = 2 Then 'Hard Skill
          SecondsLeft = 60
       End If
       
       Problem = Problem + 1 'Add one to number of problem
       Call Make_Problem 'Make new problem
    End If

    tmrQuestionTime.Enabled = True 'Enable timer

End Sub

Sub cmdCalculator_Click ()
    Dim X As Integer 'For controling shell command

    On Error Resume Next 'Skip errors

    AppActivate "Calculator" 'Activate calculator
    
    'If activate didn't work start calculator
    If Err <> 0 Then X = Shell("Calc.exe", 1)

    'Unable to start calculator
    If Err = 53 Then MsgBox "Unable to start calculator!", 16, "ERROR"
End Sub

Sub cmdEnd_Click ()
    End 'End program
End Sub

Sub Form_Load ()
    Correct = 0 'Clear out variable
    Incorrect = 0 'Clear out variable
    lblProblem.Caption = "" 'Clear out label
    lblRight.Caption = "Answers right: 0" 'Reset label
    lblWrong.Caption = "Answers wrong: 0" 'Reset label

    'Set position in center of screen
    Me.Top = (Screen.Height / 2) - (Me.Height / 2)
    Me.Left = (Screen.Width / 2) - (Me.Width / 2)

    Me.Show 'Show form
    Me.Enabled = True 'Enable form
    
    frmMain.Enabled = False 'Disable main form

    Problem = 1 'Set to first problem

    'Set time according to skill
    If Skill = 0 Then 'Easy skill
       SecondsLeft = 180
    ElseIf Skill = 1 Then 'Normal Skill
       SecondsLeft = 120
    ElseIf Skill = 2 Then 'Hard Skill
       SecondsLeft = 60
    End If

    Call Make_Problem 'Make a problem
End Sub

Sub Make_Problem ()
    
    txtAnswer.Text = "" 'Clear out answer box
    
    'Get random numbers
    Numbera = Rnd * 99
    Numberb = Rnd * 99
    Numberc = Rnd * 99

    'Can't have zero as a number
    If Numbera = 0 Then Numbera = Numbera + 1
    If Numberb = 0 Then Numberb = Numberb + 1
    If Numberc = 0 Then Numberc = Numberc + 1

    Select Case Problem 'Determine which problem to do
       Case 1
          'Update label with problem
          lblProblem.Caption = "(" + Str$(Numbera) + " + " + Str$(Numberb) + " ) / " + Str$(Numberc) + " = ?"
       Case 2
          'Update label with problem
          lblProblem.Caption = "(" + Str$(Numbera) + " + " + Str$(Numberb) + ") / ? = " + Str$(Numberc)
       Case 3
          'Update label with problem
          lblProblem.Caption = "(? + " + Str$(Numbera) + ") /" + Str$(Numberb) + " = " + Str$(Numberc)
       Case 4
          'Update label with problem
          lblProblem.Caption = "? * " + Str$(Sqr(Numbera)) + " = " + Str$(Numbera)
       Case 5
          'Update label with problem
          lblProblem.Caption = Str$(Numbera) + " * (" + Str$(Numberb) + " - " + Str$(Numberc) + ") = ?"
    End Select
End Sub

Sub tmrQuestionTime_Timer ()
    If SecondsLeft <= 0 Then 'Ran out of time
       Call Check_Answer 'Call sub to tell user
    Else
       SecondsLeft = SecondsLeft - 1 'Decrease time left
       'Update form caption to tell user how much time left
       Me.Caption = "You have " + Str$(SecondsLeft) + " seconds left to answer the problem."
    End If

End Sub

Sub txtAnswer_KeyDown (KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 And Not txtAnswer.Text = "" Then 'Enter key was pressed
       'Format player answer to round decimal
       txtAnswer.Text = Format$(txtAnswer.Text, "######.##")
       tmrQuestionTime.Enabled = False 'Disable timer
       Call Check_Answer 'Call sub to check answer
    End If
End Sub

