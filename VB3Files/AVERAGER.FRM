VERSION 2.00
Begin Form mainform 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Averager"
   ClientHeight    =   4830
   ClientLeft      =   1065
   ClientTop       =   1800
   ClientWidth     =   6510
   Height          =   5520
   Icon            =   AVERAGER.FRX:0000
   KeyPreview      =   -1  'True
   Left            =   1005
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4830
   ScaleWidth      =   6510
   Top             =   1170
   Width           =   6630
   Begin CommandButton Command3 
      Caption         =   "Quit"
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   9.75
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   495
      Left            =   4200
      TabIndex        =   13
      Top             =   3960
      Width           =   1095
   End
   Begin CommandButton Command2 
      Caption         =   "Print"
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   9.75
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   495
      Left            =   2760
      TabIndex        =   12
      Top             =   3960
      Width           =   1215
   End
   Begin CommandButton Command1 
      Caption         =   "Average"
      Default         =   -1  'True
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   9.75
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   495
      Left            =   1200
      TabIndex        =   11
      Top             =   3960
      Width           =   1215
   End
   Begin TextBox Text10 
      Enabled         =   0   'False
      Height          =   495
      Left            =   4800
      MaxLength       =   3
      TabIndex        =   9
      Top             =   1080
      Width           =   495
   End
   Begin TextBox Text9 
      Enabled         =   0   'False
      Height          =   495
      Left            =   3960
      MaxLength       =   3
      TabIndex        =   8
      Top             =   1080
      Width           =   495
   End
   Begin TextBox Text8 
      Enabled         =   0   'False
      Height          =   495
      Left            =   3000
      MaxLength       =   3
      TabIndex        =   7
      Top             =   1080
      Width           =   495
   End
   Begin TextBox Text7 
      Enabled         =   0   'False
      Height          =   495
      Left            =   2040
      MaxLength       =   3
      TabIndex        =   6
      Top             =   1080
      Width           =   495
   End
   Begin TextBox Text6 
      Enabled         =   0   'False
      Height          =   495
      Left            =   1200
      MaxLength       =   3
      TabIndex        =   5
      Top             =   1080
      Width           =   495
   End
   Begin TextBox Text5 
      Enabled         =   0   'False
      Height          =   495
      Left            =   4800
      MaxLength       =   3
      TabIndex        =   4
      Top             =   240
      Width           =   495
   End
   Begin TextBox Text4 
      Enabled         =   0   'False
      Height          =   495
      Left            =   3960
      MaxLength       =   3
      TabIndex        =   3
      Top             =   240
      Width           =   495
   End
   Begin TextBox Text3 
      Enabled         =   0   'False
      Height          =   495
      Left            =   3000
      MaxLength       =   3
      TabIndex        =   2
      Top             =   240
      Width           =   495
   End
   Begin TextBox Text2 
      Enabled         =   0   'False
      Height          =   495
      Left            =   2040
      MaxLength       =   3
      TabIndex        =   1
      Top             =   240
      Width           =   495
   End
   Begin TextBox Text1 
      Enabled         =   0   'False
      Height          =   495
      Left            =   1200
      MaxLength       =   3
      TabIndex        =   0
      Top             =   240
      Width           =   495
   End
   Begin Label Label1 
      BorderStyle     =   1  'Fixed Single
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   12
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   1695
      Left            =   360
      TabIndex        =   10
      Top             =   1920
      Width           =   5775
      WordWrap        =   -1  'True
   End
   Begin Image Image4 
      BorderStyle     =   1  'Fixed Single
      Height          =   510
      Left            =   5640
      Picture         =   AVERAGER.FRX:0302
      Top             =   3960
      Width           =   510
   End
   Begin Image Image3 
      BorderStyle     =   1  'Fixed Single
      Height          =   510
      Left            =   5640
      Picture         =   AVERAGER.FRX:0604
      Top             =   240
      Width           =   510
   End
   Begin Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   510
      Left            =   360
      Picture         =   AVERAGER.FRX:0906
      Top             =   3960
      Width           =   510
   End
   Begin Image Image1 
      Height          =   480
      Left            =   360
      Picture         =   AVERAGER.FRX:0C08
      Top             =   240
      Width           =   480
   End
   Begin Menu mnufile 
      Caption         =   "&File"
      Begin Menu mnufilenewreportcard 
         Caption         =   "New reportcard..."
      End
      Begin Menu mnufileopenreportcard 
         Caption         =   "Open reportcard..."
         Shortcut        =   ^O
      End
      Begin Menu mnufilesavereportcardas 
         Caption         =   "Save reportcard as..."
         Shortcut        =   ^S
      End
      Begin Menu mnufilesepbar 
         Caption         =   "-"
      End
      Begin Menu mnufilequit 
         Caption         =   "Quit"
      End
   End
   Begin Menu mnuhelp 
      Caption         =   "&Help"
      Begin Menu mnuhelponaverager 
         Caption         =   "... on Averager"
         Shortcut        =   ^H
      End
      Begin Menu mnuhelpabout 
         Caption         =   "About"
      End
   End
End
Sub Command1_Click ()
    g1 = Val(text1.Text)
    g2 = Val(text2.Text)
    g3 = Val(text3.Text)
    g4 = Val(text4.Text)
    g5 = Val(text5.Text)
    g6 = Val(text6.Text)
    g7 = Val(text7.Text)
    g8 = Val(text8.Text)
    g9 = Val(text9.Text)
    g10 = Val(text10.Text)
    gradet = g1 + g2 + g3 + g4 + g5 + g6 + g7 + g8 + g9 + g10
    average = gradet \ nogrades
    msgrade = "Date/Time: " + Date$ + " " + Time$ + Chr$(13)
    msgrade = msgrade + "Student: " & UCase(student) + Chr$(13)
    msgrade = msgrade + "Grades: " + text1.Text + " " + text2.Text + " " + text3.Text + " " + text4.Text + " " + text5.Text + " " + text6.Text + " " + text7.Text + " " + text8.Text + " " + text9.Text + " " + text10.Text + Chr$(13)
    msgrade = msgrade + "The grade is " & average
    label1.Caption = msgrade
End Sub

Sub command2_click ()
    printque = MsgBox("Do you wish to print this?", 36, "Averager")
    If printque = 6 Then
                      screen.MousePointer = 11
                      printer.Print label1.Caption
                      printer.EndDoc
                      screen.MousePointer = 0
    End If
End Sub

Sub Command3_Click ()
    quitnow = MsgBox("Do you wish to quit", 36, "Averager")
    If quitnow = 6 Then End
End Sub

Sub Form_Load ()
    Show
    On Error Resume Next
    student = InputBox("Type the student's name.", "Averager", "Student")
    If student = "" Then End
    nogrades = InputBox("Type the number of grades.", "Version 1.0", "7")
    If nogrades = 0 Then End
    If nogrades = 1 Then text1.Enabled = True
    If nogrades = 2 Then text1.Enabled = True: text2.Enabled = True
    If nogrades = 3 Then text1.Enabled = True: text2.Enabled = True: text3.Enabled = True
    If nogrades = 4 Then text1.Enabled = True: text2.Enabled = True: text3.Enabled = True: text4.Enabled = True
    If nogrades = 5 Then text1.Enabled = True: text2.Enabled = True: text3.Enabled = True: text4.Enabled = True: text5.Enabled = True
    If nogrades = 6 Then text1.Enabled = True: text2.Enabled = True: text3.Enabled = True: text4.Enabled = True: text5.Enabled = True: text6.Enabled = True
    If nogrades = 7 Then text1.Enabled = True: text2.Enabled = True: text3.Enabled = True: text4.Enabled = True: text5.Enabled = True: text6.Enabled = True: text7.Enabled = True
    If nogrades = 8 Then text1.Enabled = True: text2.Enabled = True: text3.Enabled = True: text4.Enabled = True: text5.Enabled = True: text6.Enabled = True: text7.Enabled = True: text8.Enabled = True
    If nogrades = 9 Then text1.Enabled = True: text2.Enabled = True: text3.Enabled = True: text4.Enabled = True: text5.Enabled = True: text6.Enabled = True: text7.Enabled = True: text8.Enabled = True: text9.Enabled = True
    If nogrades = 10 Then text1.Enabled = True: text2.Enabled = True: text3.Enabled = True: text4.Enabled = True: text5.Enabled = True: text6.Enabled = True: text7.Enabled = True: text8.Enabled = True: text9.Enabled = True: text10.Enabled = True
    If nogrades > 10 Then MsgBox "To many grades!", , "": nogrades = 10: If nogrades = 10 Then text1.Enabled = True: text2.Enabled = True: text3.Enabled = True: text4.Enabled = True: text5.Enabled = True: text6.Enabled = True: text7.Enabled = True: text8.Enabled = True: text9.Enabled = True: text10.Enabled = True
    If label1.Caption = "" Then command2.Enabled = False
End Sub

Sub Image1_MouseUp (button As Integer, shift As Integer, X As Single, Y As Single)
If shift <> 0 And button Then MsgBox "PC1", , "Message"
End Sub

Sub Image2_MouseUp (button As Integer, shift As Integer, X As Single, Y As Single)
If shift <> 0 And button Then MsgBox "PC2", , "Message"
End Sub

Sub Image3_MouseUp (button As Integer, shift As Integer, X As Single, Y As Single)
If shift <> 0 And button Then MsgBox "PC3", , "Message"
End Sub

Sub Image4_MouseUp (button As Integer, shift As Integer, X As Single, Y As Single)
If shift <> 0 And button Then MsgBox "PC4", , "Message"
End Sub

Sub Label1_Change ()
    If label1.Caption > "" Then command2.Enabled = True
    If label1.Caption = "" Then command2.Enabled = False
End Sub

Sub mnufilenewreportcard_Click ()
    On Error Resume Next
    student = InputBox("Type the student's name.", "Averager", "Student")
    If student = "" Then Exit Sub
    nogrades = InputBox("Type the number of grades.", "Version 1.0", "7")
    text1.Enabled = False
    text2.Enabled = False
    text3.Enabled = False
    text4.Enabled = False
    text5.Enabled = False
    text6.Enabled = False
    text7.Enabled = False
    text8.Enabled = False
    text9.Enabled = False
    text10.Enabled = False
    If nogrades = 0 Then Exit Sub
    If nogrades = 1 Then text1.Enabled = True
    If nogrades = 2 Then text1.Enabled = True: text2.Enabled = True
    If nogrades = 3 Then text1.Enabled = True: text2.Enabled = True: text3.Enabled = True
    If nogrades = 4 Then text1.Enabled = True: text2.Enabled = True: text3.Enabled = True: text4.Enabled = True
    If nogrades = 5 Then text1.Enabled = True: text2.Enabled = True: text3.Enabled = True: text4.Enabled = True: text5.Enabled = True
    If nogrades = 6 Then text1.Enabled = True: text2.Enabled = True: text3.Enabled = True: text4.Enabled = True: text5.Enabled = True: text6.Enabled = True
    If nogrades = 7 Then text1.Enabled = True: text2.Enabled = True: text3.Enabled = True: text4.Enabled = True: text5.Enabled = True: text6.Enabled = True: text7.Enabled = True
    If nogrades = 8 Then text1.Enabled = True: text2.Enabled = True: text3.Enabled = True: text4.Enabled = True: text5.Enabled = True: text6.Enabled = True: text7.Enabled = True: text8.Enabled = True
    If nogrades = 9 Then text1.Enabled = True: text2.Enabled = True: text3.Enabled = True: text4.Enabled = True: text5.Enabled = True: text6.Enabled = True: text7.Enabled = True: text8.Enabled = True: text9.Enabled = True
    If nogrades = 10 Then text1.Enabled = True: text2.Enabled = True: text3.Enabled = True: text4.Enabled = True: text5.Enabled = True: text6.Enabled = True: text7.Enabled = True: text8.Enabled = True: text9.Enabled = True: text10.Enabled = True
    text1.Text = "0"
    text2.Text = "0"
    text3.Text = "0"
    text4.Text = "0"
    text5.Text = "0"
    text6.Text = "0"
    text7.Text = "0"
    text8.Text = "0"
    text9.Text = "0"
    text10.Text = "0"
    label1.Caption = ""
End Sub

Sub mnufileopenreportcard_Click ()
    saveopenform.Caption = "Open Reportcard"
    mainform.Enabled = False
    saveopenform.Visible = True
End Sub

Sub mnufilequit_Click ()
    quitnow = MsgBox("Do you wish to quit", 36, "Averager")
    If quitnow = 6 Then End
End Sub

Sub mnufilesavereportcardas_Click ()
    saveopenform.Caption = "Save reportcard as"
    saveopenform.Visible = True
    mainform.Enabled = False
End Sub

Sub mnuhelpabout_Click ()
    Dim msg As String

    msg = "Averager ver 1.0" + Chr$(13)
    msg = msg + "Program writing by Rodney Beede." + Chr$(13)
    msg = msg + "Program designed by Rodney Beede." + Chr$(13)
    msg = msg + "Program made in Visual Basic 3.0 for MS-Windows 3.x" + Chr$(13)
    msg = msg + "________________________________" + Chr$(13)
    msg = msg + "Questions?  Write to 900 Gst Snyder,OK 73566"
    MsgBox msg, 64, "About Averager"
End Sub

Sub mnuhelponaverager_Click ()
Dim msg As String
   msg = "To activate text boxes just move the mouse pointer over it and click." + Chr$(13)
   msg = msg + "You type the grades in the text boxes and " + Chr$(13)
   msg = msg + "then click on the Average button to average the grades." + Chr$(13)
   msg = msg + "_____________________________________________" + Chr$(13)
   msg = msg + "                        The Buttons" + Chr$(13)
   msg = msg + "The average button adds all of the " + Chr$(13)
   msg = msg + "text boxes then divides them by the number of grades." + Chr$(13)
   msg = msg + "The print button lets you print the reportcard on the printer." + Chr$(13)
   msg = msg + "The quit buttton lets you end the program." + Chr$(13)
   msg = msg + "                        The File menu." + Chr$(13)
   msg = msg + "The File menu lets you open and save reportcards." + Chr$(13)
   msg = msg + "It also lets you end the program." + Chr$(13)
   msg = msg + "                        The Text boxes and grades." + Chr$(13)
   msg = msg + "The text boxes are what you type the grades in." + Chr$(13)
   msg = msg + "  The grade range is 0 to 120. If you type in" + Chr$(13)
   msg = msg + "   a grade larger than 120 " '+ Chr$(13)
   msg = msg + " you will receive an error message."
   MsgBox msg, , "Help on Averager."
End Sub

Sub Text1_Change ()
If text1.Text = "" Or text1.Text Like "." Then Exit Sub
If Val(text1.Text) > 120 Then
                      MsgBox "To high grade!", 48, "": text1.Text = "120"
End If
GoTo 10
10 End Sub



Sub Text1_GotFocus ()
text1.SelStart = 0
text1.SelLength = Len(text1.Text)
End Sub

Sub Text10_Change ()
If text10.Text = "" Or text10.Text Like "." Then Exit Sub
If Val(text10.Text) > 120 Then
                      MsgBox "To high grade!", 48, "": text10.Text = "120"
End If
GoTo 19
19 End Sub

Sub Text10_GotFocus ()
text10.SelStart = 0
text10.SelLength = Len(text1.Text)
End Sub

Sub text2_change ()
If text2.Text = "" Or text2.Text Like "." Then Exit Sub
If Val(text2.Text) > 120 Then
                      MsgBox "To high grade!", 48, "": text2.Text = "120"
End If
GoTo 11
11 End Sub

Sub text2_gotfocus ()
text2.SelStart = 0
text2.SelLength = Len(text1.Text)
End Sub

Sub Text3_Change ()
If text3.Text = "" Or text3.Text Like "." Then Exit Sub
If Val(text3.Text) > 120 Then
                      MsgBox "To high grade!", 48, "": text3.Text = "120"
End If
Exit Sub
End Sub

Sub Text3_GotFocus ()
text3.SelStart = 0
text3.SelLength = Len(text1.Text)

End Sub

Sub Text4_Change ()
If text4.Text = "" Or text4.Text Like "." Then Exit Sub
If Val(text4.Text) > 120 Then
                      MsgBox "To high grade!", 48, "": text4.Text = "120"
End If
Exit Sub
End Sub

Sub Text4_GotFocus ()
text4.SelStart = 0
text4.SelLength = Len(text1.Text)

End Sub

Sub Text5_Change ()
If text5.Text = "" Or text5.Text Like "." Then Exit Sub
If Val(text5.Text) > 120 Then
                      MsgBox "To high grade!", 48, "": text5.Text = "120"
End If
Exit Sub
End Sub

Sub Text5_GotFocus ()
text5.SelStart = 0
text5.SelLength = Len(text1.Text)

End Sub

Sub Text6_Change ()
If text6.Text = "" Or text6.Text Like "." Then Exit Sub
If Val(text6.Text) > 120 Then
                      MsgBox "To high grade!", 48, "": text6.Text = "120"
End If
Exit Sub
End Sub

Sub Text6_GotFocus ()
text6.SelStart = 0
text6.SelLength = Len(text1.Text)

End Sub

Sub Text7_Change ()
If text7.Text = "" Or text7.Text Like "." Then Exit Sub
If Val(text7.Text) > 120 Then
                      MsgBox "To high grade!", 48, "": text7.Text = "120"
End If
Exit Sub
End Sub

Sub Text7_GotFocus ()
text7.SelStart = 0
text7.SelLength = Len(text1.Text)

End Sub

Sub Text8_Change ()
If text8.Text = "" Or text8.Text Like "." Then Exit Sub
If Val(text8.Text) > 120 Then
                      MsgBox "To high grade!", 48, "": text8.Text = "120"
End If
Exit Sub
End Sub

Sub Text8_GotFocus ()
text8.SelStart = 0
text8.SelLength = Len(text1.Text)

End Sub

Sub Text9_Change ()
If text9.Text = "" Or text9.Text Like "." Then Exit Sub
If Val(text9.Text) > 120 Then
                      MsgBox "To high grade!", 48, "": text9.Text = "120"
End If
Exit Sub
End Sub

Sub Text9_GotFocus ()
text9.SelStart = 0
text9.SelLength = Len(text1.Text)

End Sub

