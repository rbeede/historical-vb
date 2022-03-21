VERSION 2.00
Begin Form saveopenform 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Double
   ClientHeight    =   4065
   ClientLeft      =   2370
   ClientTop       =   1710
   ClientWidth     =   4575
   Height          =   4470
   Left            =   2310
   LinkTopic       =   "Form1"
   ScaleHeight     =   4065
   ScaleWidth      =   4575
   Top             =   1365
   Width           =   4695
   Begin CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   315
      Left            =   2280
      TabIndex        =   6
      Top             =   3480
      Width           =   975
   End
   Begin CommandButton Command1 
      Caption         =   "OK"
      Height          =   315
      Left            =   1140
      TabIndex        =   5
      Top             =   3480
      Width           =   975
   End
   Begin TextBox Text1 
      Height          =   330
      Left            =   60
      MaxLength       =   12
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   60
      Width           =   4395
   End
   Begin FileListBox File1 
      Height          =   2565
      Left            =   2280
      Pattern         =   "*.rpt"
      TabIndex        =   2
      Top             =   660
      Width           =   2175
   End
   Begin DirListBox Dir1 
      Height          =   2055
      Left            =   60
      TabIndex        =   1
      Top             =   660
      Width           =   2055
   End
   Begin DriveListBox Drive1 
      Height          =   315
      Left            =   60
      TabIndex        =   0
      Top             =   2910
      Width           =   2055
   End
   Begin Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   195
      Left            =   60
      TabIndex        =   3
      Top             =   420
      Width           =   585
   End
End
Dim filelist() As String

Sub check ()
    mainform.Text1.Enabled = False
    mainform.Text2.Enabled = False
    mainform.Text3.Enabled = False
    mainform.Text4.Enabled = False
    mainform.Text5.Enabled = False
    mainform.Text6.Enabled = False
    mainform.Text7.Enabled = False
    mainform.Text8.Enabled = False
    mainform.Text9.Enabled = False
    mainform.Text10.Enabled = False
    If nogrades = 0 Then Exit Sub
    If nogrades = 1 Then mainform.Text1.Enabled = True
    If nogrades = 2 Then mainform.Text1.Enabled = True: mainform.Text2.Enabled = True
    If nogrades = 3 Then mainform.Text1.Enabled = True: mainform.Text2.Enabled = True: mainform.Text3.Enabled = True
    If nogrades = 4 Then mainform.Text1.Enabled = True: mainform.Text2.Enabled = True: mainform.Text3.Enabled = True: mainform.Text4.Enabled = True
    If nogrades = 5 Then mainform.Text1.Enabled = True: mainform.Text2.Enabled = True: mainform.Text3.Enabled = True: mainform.Text4.Enabled = True: mainform.Text5.Enabled = True
    If nogrades = 6 Then mainform.Text1.Enabled = True: mainform.Text2.Enabled = True: mainform.Text3.Enabled = True: mainform.Text4.Enabled = True: mainform.Text5.Enabled = True: mainform.Text6.Enabled = True
    If nogrades = 7 Then mainform.Text1.Enabled = True: mainform.Text2.Enabled = True: mainform.Text3.Enabled = True: mainform.Text4.Enabled = True: mainform.Text5.Enabled = True: mainform.Text6.Enabled = True: mainform.Text7.Enabled = True
    If nogrades = 8 Then mainform.Text1.Enabled = True: mainform.Text2.Enabled = True: mainform.Text3.Enabled = True: mainform.Text4.Enabled = True: mainform.Text5.Enabled = True: mainform.Text6.Enabled = True: mainform.Text7.Enabled = True: mainform.Text8.Enabled = True
    If nogrades = 9 Then mainform.Text1.Enabled = True: mainform.Text2.Enabled = True: mainform.Text3.Enabled = True: mainform.Text4.Enabled = True: mainform.Text5.Enabled = True: mainform.Text6.Enabled = True: mainform.Text7.Enabled = True: mainform.Text8.Enabled = True: mainform.Text9.Enabled = True
    If nogrades = 10 Then mainform.Text1.Enabled = True: mainform.Text2.Enabled = True: mainform.Text3.Enabled = True: mainform.Text4.Enabled = True: mainform.Text5.Enabled = True: mainform.Text6.Enabled = True: mainform.Text7.Enabled = True: mainform.Text8.Enabled = True: mainform.Text9.Enabled = True: mainform.Text10.Enabled = True
End Sub

Sub Command1_Click ()
    On Error GoTo erro
       If saveopenform.caption = "Save reportcard as" Then
        If Text1.Text Like "*[.]*" Then Text1.Text = Text1.Text Else Text1.Text = Text1.Text + ".rpt"
        If file1.ListIndex > -1 Then q = MsgBox("Warning.  This will overwrite this file.", 17, "Warning")
           If q = 2 Then Exit Sub
        If Len(dir1.Path) > 3 Then
           fileplace$ = dir1.Path + "\" + Text1.Text
        Else
           fileplace$ = dir1.Path + Text1.Text
        End If
        Open fileplace$ For Output As #1
        Print #1, mainform.label1.caption
        Print #1, student
        Print #1, nogrades
        Close #1
        saveopenform.Hide
        mainform.Enabled = True
        Exit Sub
       End If
       If saveopenform.caption = "Open Reportcard" Then
           If Len(dir1.Path) > 3 Then
              fileplace$ = dir1.Path + "\" + Text1.Text
           Else
              fileplace$ = dir1.Path + Text1.Text
           End If
           Open fileplace$ For Input As #2
           Do Until EOF(2)
            Input #2, rd$, studnt$, grades$, avge$, studenta$, nogradesa
           mainform.label1.caption = rd$ + Chr$(13) + studnt$ + Chr$(13) + grades$ + Chr$(13) + avge$
           nogrades = nogradesa
           student = studenta$
           Loop
           mainform.label1.caption = rd$ + Chr$(13) + studnt$ + Chr$(13) + grades$ + Chr$(13) + avge$
           nogrades = nogradesa
           student = studenta$
           Close #2
           Call check
           saveopenform.Hide
           mainform.Enabled = True
          Exit Sub
         End If
erro: MsgBox "Error " & Err & " " & Error$, 16, "Error"
mainform.label1.caption = ""
Close
Exit Sub
End Sub

Sub command2_click ()
    mainform.Enabled = True
    saveopenform.Visible = False
End Sub

Sub Dir1_Change ()
    file1.Path = dir1.Path

    
    If Len(dir1.Path) > 3 Then
       label1.caption = dir1.Path + "\" + file1.FileName
    Else
       label1.caption = dir1.Path + file1.Pattern
    End If

    
    If file1.FileName <> "" Then
       Text1.Text = file1.FileName
    Else
       Text1.Text = file1.Pattern
    End If
End Sub

Sub Dir1_KeyPress (keyascii As Integer)
    If keyascii = 13 Then
       keyascii = 0
       dir1.Path = dir1.List(dir1.ListIndex)
    End If
End Sub

Sub Drive1_Change ()
    Dim msg As String
    
    
    On Error Resume Next

    
    screen.MousePointer = 11
    dir1.Path = drive1.Drive
    screen.MousePointer = 0

    
    If Err Then
       msg = "Error " & Err & ": " & Error$ & "."
       MsgBox msg, 48, "Drive/Directory Error"
       drive1.Drive = dir1.Path
       End If
       On Error GoTo 0
End Sub

Sub File1_Click ()
   If file1.ListIndex = -1 Then Text1.Text = Text1.Text Else Text1.Text = file1.FileName
End Sub

Sub file1_dblclick ()
    Command1_Click
End Sub

Sub File1_KeyPress (keyascii As Integer)
    If keyascii = 13 Then
       keyascii = 0
       file1_dblclick
    End If
End Sub

Sub File1_PathChange ()
    If Len(file1.Path) > 3 Then
      label1.caption = file1.Path + "\" + file1.Pattern
    Else
      label1.caption = file1.Path + file1.Pattern
    End If

    
    If dir1.Path <> file1.Path Then dir1.Path = file1.Path
    If drive1.Drive <> Left$(file1.Path, 1) Then
       drive1.Drive = Left$(file1.Path, 1)
    End If
    If file1.ListCount = 0 Then Text1.Text = file1.Pattern
End Sub

Sub File1_PatternChange ()
    If Len(file1.Path) > 3 Then
      label1.caption = file1.Path + "\" + file1.Pattern
    Else
      label1.caption = file1.Path + file1.Pattern
    End If
    

    If file1.ListCount = 0 Then
       Text1.Text = file1.Pattern
    Else
       file1.ListIndex = 0
       Text1.Text = file1.FileName
    End If

    Text1.SelStart = 0
    Text1.SelLength = Len(Text1.Text)
    Text1.SetFocus
End Sub

Sub Form_Load ()
    drive1.Drive = CurDir$
    dir1.Path = CurDir$
    label1.caption = dir1.Path
    file1.Path = CurDir$
    If file1.FileName <> "" Then
       Text1.Text = file1.FileName
    Else
       Text1.Text = file1.Pattern
       End If
End Sub

Sub Text1_Change ()
    If Text1.Text Like "*[>]*" Or Text1.Text Like "*[ ,+/]*" Then
       msg = Text1.Text + " is an invalid file name."
       MsgBox msg, 48, "Bad File Name"
    End If
End Sub

Sub Text1_Click ()
    file1.FileName = ""
    file1.ListIndex = -1
End Sub

Sub Text1_GotFocus ()
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1.Text)
End Sub

Sub Text1_KeyPress (keyascii As Integer)
    Dim msg As String

    If keyascii <> 13 Then Exit Sub
    keyascii = 0

    If Text1.Text = file1.List(file1.ListIndex) Then
       file1.Selected(file1.ListIndex) = True
       Command1_Click
       Exit Sub
    End If


    If Text1.Text Like "*[;>]*" Or Text1.Text Like "*[ ,+:/]*" Then
       msg = Text1.Text + " is an invalid file name."
       MsgBox msg, 48, "Bad File Name"
       Text1.Text = file1.FileName
       Text1.SetFocus
       Text1.SelStart = 0
       Text1.SelLength = Len(Text1.Text)
       Exit Sub

       
    Else
       On Error Resume Next
       file1.FileName = Text1.Text


       If Err Then
          msg = "Error " & Err & ": " + Error$ + "."
          MsgBox msg, 48, "File Error"
       End If
    On Error GoTo 0


    If file1.ListCount > 0 Then
       Text1.Text = file1.FileName
    Else
       Text1.Text = file1.Pattern
    End If


    Text1.SelStart = 0
    Text1.SelLength = Len(Text1.Text)
    Text1.SetFocus
  End If
End Sub

