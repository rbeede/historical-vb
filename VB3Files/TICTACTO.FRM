VERSION 2.00
Begin Form frmmain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tic Tac Toe"
   ClientHeight    =   4020
   ClientLeft      =   1875
   ClientTop       =   1950
   ClientWidth     =   3510
   Height          =   4710
   Icon            =   TICTACTO.FRX:0000
   Left            =   1815
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   3510
   Top             =   1320
   Width           =   3630
   Begin Label lblmessage 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Select Start from Game Menu."
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   480
      TabIndex        =   9
      Top             =   120
      Width           =   2655
   End
   Begin Line cross 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Index           =   3
      X1              =   2160
      X2              =   2160
      Y1              =   600
      Y2              =   3720
   End
   Begin Line cross 
      BorderWidth     =   2
      Index           =   2
      X1              =   1320
      X2              =   1320
      Y1              =   600
      Y2              =   3720
   End
   Begin Line cross 
      BorderWidth     =   2
      Index           =   1
      X1              =   360
      X2              =   3120
      Y1              =   2640
      Y2              =   2640
   End
   Begin Line cross 
      BorderWidth     =   2
      Index           =   0
      X1              =   360
      X2              =   3120
      Y1              =   1680
      Y2              =   1680
   End
   Begin Label lblplace 
      Alignment       =   1  'Right Justify
      Caption         =   "O"
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   24
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   975
      Index           =   8
      Left            =   2280
      TabIndex        =   8
      Top             =   2760
      Visible         =   0   'False
      Width           =   855
   End
   Begin Label lblplace 
      Alignment       =   2  'Center
      Caption         =   "O"
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   24
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   975
      Index           =   7
      Left            =   1440
      TabIndex        =   7
      Top             =   2760
      Visible         =   0   'False
      Width           =   645
   End
   Begin Label lblplace 
      Caption         =   "O"
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   24
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   975
      Index           =   6
      Left            =   360
      TabIndex        =   6
      Top             =   2760
      Visible         =   0   'False
      Width           =   885
   End
   Begin Label lblplace 
      Alignment       =   1  'Right Justify
      Caption         =   "O"
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   24
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   795
      Index           =   5
      Left            =   2280
      TabIndex        =   5
      Top             =   1800
      Visible         =   0   'False
      Width           =   885
   End
   Begin Label lblplace 
      Alignment       =   2  'Center
      Caption         =   "O"
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   24
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   735
      Index           =   4
      Left            =   1440
      TabIndex        =   4
      Top             =   1800
      Visible         =   0   'False
      Width           =   615
   End
   Begin Label lblplace 
      Caption         =   "O"
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   24
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   735
      Index           =   3
      Left            =   360
      TabIndex        =   3
      Top             =   1800
      Visible         =   0   'False
      Width           =   885
   End
   Begin Label lblplace 
      Alignment       =   1  'Right Justify
      Caption         =   "O"
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   24
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   975
      Index           =   2
      Left            =   2280
      TabIndex        =   2
      Top             =   600
      Visible         =   0   'False
      Width           =   885
   End
   Begin Label lblplace 
      Alignment       =   2  'Center
      Caption         =   "O"
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   24
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   975
      Index           =   1
      Left            =   1440
      TabIndex        =   1
      Top             =   600
      Visible         =   0   'False
      Width           =   615
   End
   Begin Label lblplace 
      Caption         =   "O"
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   24
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   975
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Visible         =   0   'False
      Width           =   885
   End
   Begin Menu mnugame 
      Caption         =   "&Game"
      Begin Menu mnugamestart 
         Caption         =   "&Start"
      End
      Begin Menu mnugameend 
         Caption         =   "&End"
         Enabled         =   0   'False
      End
      Begin Menu mnugamecolors 
         Caption         =   "Co&lors"
         Begin Menu mnugamecolorsbackground 
            Caption         =   "&Back Ground"
            Begin Menu mnugamecolorsbackgroundred 
               Caption         =   "&Red"
            End
            Begin Menu mnugamecolorsbackgroundwhite 
               Caption         =   "&White"
            End
            Begin Menu mnugamecolorsbackgroundlightgreen 
               Caption         =   "&Light Green"
            End
         End
         Begin Menu mnugamecolorslinescolor 
            Caption         =   "&Lines Color"
            Begin Menu mnugamecolorslinescolorred 
               Caption         =   "&Red"
            End
            Begin Menu mnugamecolorslinescolorblack 
               Caption         =   "&Black"
            End
            Begin Menu mnugamecolorslinescolorlightgreen 
               Caption         =   "&Light Green"
            End
         End
      End
      Begin Menu mnugamesepbar1 
         Caption         =   "-"
      End
      Begin Menu mnugameexit 
         Caption         =   "E&xit"
      End
   End
   Begin Menu mnuAbout 
      Caption         =   "&About"
   End
End
Dim Started As Integer
Dim Playing As String

Sub check ()
    lblmessage.Caption = "Checking move."
    
    If lblplace(0).Caption = "O" And lblplace(1).Caption = "O" And lblplace(2).Caption = "O" Then
       Beep
       MsgBox "O's win the game!", 48, "Winner"
       GoTo stopgame
    ElseIf lblplace(0).Caption = "O" And lblplace(3).Caption = "O" And lblplace(6).Caption = "O" Then
       Beep
       MsgBox "O's win the game!", 48, "Winner"
       GoTo stopgame
    ElseIf lblplace(0).Caption = "O" And lblplace(4).Caption = "O" And lblplace(8).Caption = "O" Then
       Beep
       MsgBox "O's win the game!", 48, "Winner"
       GoTo stopgame
    ElseIf lblplace(2).Caption = "O" And lblplace(4).Caption = "O" And lblplace(6).Caption = "O" Then
       Beep
       MsgBox "O's win the game!", 48, "Winner"
       GoTo stopgame
    ElseIf lblplace(1).Caption = "O" And lblplace(4).Caption = "O" And lblplace(7).Caption = "O" Then
       Beep
       MsgBox "O's win the game!", 48, "Winner"
       GoTo stopgame
    ElseIf lblplace(2).Caption = "O" And lblplace(5).Caption = "O" And lblplace(8).Caption = "O" Then
       Beep
       MsgBox "O's win the game!", 48, "Winner"
       GoTo stopgame
    ElseIf lblplace(3).Caption = "O" And lblplace(4).Caption = "O" And lblplace(5).Caption = "O" Then
       Beep
       MsgBox "O's win the game!", 48, "Winner"
       GoTo stopgame
    ElseIf lblplace(6).Caption = "O" And lblplace(7).Caption = "O" And lblplace(8).Caption = "O" Then
       Beep
       MsgBox "O's win the game!", 48, "Winner"
       GoTo stopgame
    End If

    If lblplace(0).Caption = "X" And lblplace(1).Caption = "X" And lblplace(2).Caption = "X" Then
       Beep
       MsgBox "X's win the game!", 48, "Winner"
       GoTo stopgame
    ElseIf lblplace(0).Caption = "X" And lblplace(3).Caption = "X" And lblplace(6).Caption = "X" Then
       Beep
       MsgBox "X's win the game!", 48, "Winner"
       GoTo stopgame
    ElseIf lblplace(0).Caption = "X" And lblplace(4).Caption = "X" And lblplace(8).Caption = "X" Then
       Beep
       MsgBox "X's win the game!", 48, "Winner"
       GoTo stopgame
    ElseIf lblplace(2).Caption = "X" And lblplace(4).Caption = "X" And lblplace(6).Caption = "X" Then
       Beep
       MsgBox "X's win the game!", 48, "Winner"
       GoTo stopgame
    ElseIf lblplace(1).Caption = "X" And lblplace(4).Caption = "X" And lblplace(7).Caption = "X" Then
       Beep
       MsgBox "X's win the game!", 48, "Winner"
       GoTo stopgame
    ElseIf lblplace(2).Caption = "X" And lblplace(5).Caption = "X" And lblplace(8).Caption = "X" Then
       Beep
       MsgBox "X's win the game!", 48, "Winner"
       GoTo stopgame
    ElseIf lblplace(3).Caption = "X" And lblplace(4).Caption = "X" And lblplace(5).Caption = "X" Then
       Beep
       MsgBox "X's win the game!", 48, "Winner"
       GoTo stopgame
    ElseIf lblplace(6).Caption = "X" And lblplace(7).Caption = "X" And lblplace(8).Caption = "X" Then
       Beep
       MsgBox "X's win the game!", 48, "Winner"
       GoTo stopgame
    End If

    If lblplace(0).Caption <> "" And lblplace(1).Caption <> "" And lblplace(2).Caption <> "" Then
     If lblplace(3).Caption <> "" And lblplace(4).Caption <> "" And lblplace(5).Caption <> "" Then
      If lblplace(6).Caption <> "" And lblplace(7).Caption <> "" And lblplace(8).Caption <> "" Then
       
       lblmessage.Caption = "KAT!  KAT!"
       For wait = 1 To 5000
         DoEvents
       Next wait
       GoTo stopgame
      End If
     End If
    End If

Exit Sub
stopgame:
    Started = False
    For i = 0 To 8
        lblplace(i).Caption = ""
    Next i
    mnugamestart.Enabled = True
    mnugameend.Enabled = False
    lblmessage.Caption = "Select Start from Game Menu."
End Sub

Sub Form_Load ()
    Dim spath As String, bColor As Long, i As Integer
    Dim lColor As Long

    If Len(app.Path) > 3 Then
       spath = app.Path + "\"
    Else
       spath = app.Path
    End If

    On Error Resume Next
    
    Open spath + "tictacto.ini" For Input As #2

    Input #2, bColor, lColor
    
    Close

    If Err Then
       bColor = &HFFFFFF
       lColor = &H80000008
    End If

    Me.BackColor = bColor
    For i = 0 To 3
        cross(i).BorderColor = lColor
    Next i
    
    lblmessage.BackColor = Me.BackColor

    Started = False
End Sub

Sub Form_Unload (Cancel As Integer)
    Call inifile
End Sub

Sub inifile ()
    Dim spath As String

    If Len(app.Path) > 3 Then
       spath = app.Path + "\"
    Else
       spath = app.Path
    End If

    On Error Resume Next
    
    Open spath + "tictacto.ini" For Output As #1

    Write #1, Me.BackColor
    Write #1, cross(0).BorderColor

    Close

    If Err Then
       MsgBox "Error " & Err & "." & Chr$(13) & Error$ + ".", 16, "Read/Write error to .ini file."
    End If
End Sub

Sub lblplace_Click (Index As Integer)
  If lblplace(Index).Caption > "" Then Exit Sub
  
  If Started = True Then
    If Playing = "X" Then
       lblplace(Index).ForeColor = QBColor(12)
    Else
       lblplace(Index).ForeColor = QBColor(9)
    End If
    
    lblplace(Index).Caption = Playing
    Call check
    
    If Started = False Then Exit Sub

    If Playing = "X" Then
       Playing = "O"
       lblmessage.Caption = "It's O's turn."
    Else
       Playing = "X"
       lblmessage.Caption = "It's X's turn."
    End If
  End If
End Sub

Sub mnuAbout_Click ()
    Dim Msg As String
    Dim NL As String

    NL = Chr$(13) + Chr$(10)

    Msg = "Program by Rodney Beede" & NL
    Msg = Msg + "E-mail me at rodney_beede@hotmail.com" + NL
    Msg = Msg + NL
    Msg = Msg + "Published by Infinisoft." + NL
    Msg = Msg + "Read the tictacto.txt file for more information."

    MsgBox Msg, 64, "Program Information"
End Sub

Sub mnugamecolorsbackgroundlightgreen_Click ()
    mnugamecolorsbackgroundlightgreen.Checked = True
    mnugamecolorsbackgroundwhite.Checked = False
    mnugamecolorsbackgroundred.Checked = False
    Me.BackColor = QBColor(10)
    Dim i As Integer
    For i = 0 To 8
        lblplace(i).BackColor = Me.BackColor
        lblmessage.BackColor = Me.BackColor
    Next i
End Sub

Sub mnugamecolorsbackgroundred_Click ()
    mnugamecolorsbackgroundlightgreen.Checked = False
    mnugamecolorsbackgroundwhite.Checked = False
    mnugamecolorsbackgroundred.Checked = True
    Me.BackColor = QBColor(12)

    Dim i As Integer
    For i = 0 To 8
        lblplace(i).BackColor = Me.BackColor
        lblmessage.BackColor = Me.BackColor
    Next i
End Sub

Sub mnugamecolorsbackgroundwhite_Click ()
    mnugamecolorsbackgroundlightgreen.Checked = False
    mnugamecolorsbackgroundwhite.Checked = True
    mnugamecolorsbackgroundred.Checked = False
    Me.BackColor = QBColor(15)
    Dim i As Integer
    For i = 0 To 8
        lblplace(i).BackColor = Me.BackColor
        lblmessage.BackColor = Me.BackColor
    Next i
End Sub

Sub mnugamecolorslinescolorblack_Click ()
    mnugamecolorslinescolorblack.Checked = True
    mnugamecolorslinescolorred.Checked = False
    mnugamecolorslinescolorlightgreen.Checked = False
    Dim i As Integer
    For i = 0 To 3
      cross(i).BorderColor = QBColor(0)
    Next i
   
End Sub

Sub mnugamecolorslinescolorlightgreen_Click ()
    mnugamecolorslinescolorblack.Checked = False
    mnugamecolorslinescolorred.Checked = False
    mnugamecolorslinescolorlightgreen.Checked = True
    Dim i As Integer
    For i = 0 To 3
      cross(i).BorderColor = QBColor(10)
    Next i
End Sub

Sub mnugamecolorslinescolorred_Click ()
    mnugamecolorslinescolorblack.Checked = False
    mnugamecolorslinescolorred.Checked = True
    mnugamecolorslinescolorlightgreen.Checked = False
    
    Dim i As Integer
    For i = 0 To 3
      cross(i).BorderColor = QBColor(12)
    Next i
End Sub

Sub mnugameend_click ()
    Dim response As Integer, i As Integer

    response = MsgBox("Do you wish to end game?", 36, "End Tic Tac Toe Game?")
    If response = 6 Then
       Started = False
       For i = 0 To 8
          lblplace(i).Caption = ""
       Next i
       mnugamestart.Enabled = True
       mnugameend.Enabled = False
       lblmessage.Caption = "Select Start from Game Menu."
    End If
End Sub

Sub mnugameexit_Click ()
    Call inifile
    End
End Sub

Sub mnugamestart_Click ()
    Dim Msg As String, i As Integer
    
    Started = True
    
    Msg = "First player will be X's." + Chr$(13)
    Msg = Msg + "Second player will be X's." + Chr$(13)
    Msg = Msg + "X's will go first."
    MsgBox Msg, 64, "Tic Tac Toe"
    Playing = "X"
    For i = 0 To 8
        lblplace(i).BackColor = Me.BackColor
        lblplace(i).Caption = ""
        lblplace(i).Visible = True
    Next i
    mnugamestart.Enabled = False
    mnugameend.Enabled = True
    lblmessage.Caption = "It's X's turn."
End Sub

