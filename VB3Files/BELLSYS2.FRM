VERSION 2.00
Begin Form FrmWeekBellTime 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Double
   Caption         =   "Set Weekly Times"
   ClientHeight    =   3255
   ClientLeft      =   1005
   ClientTop       =   1530
   ClientWidth     =   7710
   ControlBox      =   0   'False
   Height          =   3660
   Left            =   945
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   7710
   Top             =   1185
   Width           =   7830
   Begin TextBox txtTime 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   2640
      Visible         =   0   'False
      Width           =   300
   End
   Begin Grid grdTimes 
      BorderStyle     =   0  'None
      Cols            =   6
      Height          =   2415
      HighLight       =   0   'False
      Left            =   240
      Rows            =   10
      ScrollBars      =   0  'None
      TabIndex        =   0
      Top             =   120
      Width           =   7295
   End
   Begin CommandButton CmdCancel 
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   3960
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   2640
      Width           =   3615
   End
   Begin CommandButton cmdDone 
      Caption         =   "&Done"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2640
      Width           =   3615
   End
   Begin Line LineY 
      BorderWidth     =   3
      X1              =   0
      X2              =   120
      Y1              =   1800
      Y2              =   3120
   End
   Begin Line LineX 
      BorderWidth     =   3
      X1              =   120
      X2              =   0
      Y1              =   120
      Y2              =   1560
   End
End

Sub cmdCancel_click ()
  frmMain!tmrWeekBellTime.Enabled = True
  frmMain.Enabled = True
  Me.Enabled = False
  Unload Me
End Sub

Sub cmdDone_Click ()
  Dim i As Integer, c As Integer, ci As Integer
  
  grdTimes.Visible = False

  Me.Caption = "Saving data.  Please wait."

  i = 0
  For c = 1 To grdTimes.Cols - 1
    For ci = 1 To grdTimes.Rows - 1
      grdTimes.Row = ci
      grdTimes.Col = c
      WBTI(i) = grdTimes.Text
      i = i + 1
    Next ci
  Next c
  
  grdTimes.Visible = True

  Call cmdCancel_click
End Sub

Sub Form_Load ()
    Dim counter As Integer, c As Integer, ci As Integer, i As Integer
    
    grdTimes.Rows = 10
    grdTimes.Cols = 6
    
    grdTimes.FixedCols = 0
    grdTimes.FixedRows = 0
    
    grdTimes.Row = 0
    grdTimes.Col = 1
    grdTimes.Text = "Monday     "
    grdTimes.Col = 2
    grdTimes.Text = "Tuesday    "
    grdTimes.Col = 3
    grdTimes.Text = "Wednesday  "
    grdTimes.Col = 4
    grdTimes.Text = "Thursday   "
    grdTimes.Col = 5
    grdTimes.Text = "Friday     "
    
    grdTimes.Col = 0
    grdTimes.Row = 1
    grdTimes.Text = "1st"
    grdTimes.Row = 2
    grdTimes.Text = "2nd"
    grdTimes.Row = 3
    grdTimes.Text = "3rd"
    grdTimes.Row = 4
    grdTimes.Text = "4th"
    grdTimes.Row = 5
    grdTimes.Text = "5th"
    grdTimes.Row = 6
    grdTimes.Text = "6th"
    grdTimes.Row = 7
    grdTimes.Text = "7th"
    grdTimes.Row = 8
    grdTimes.Text = "Lunch"
    grdTimes.Row = 9
    grdTimes.Text = "Last"
     
    For counter = 1 To grdTimes.Cols - 1
        grdTimes.Row = 0
        grdTimes.Col = counter
        grdTimes.ColWidth(counter) = Len(grdTimes.Text) * 120
    Next counter

    grdTimes.ColWidth(0) = grdTimes.ColWidth(0) + 50

    grdTimes.Width = 0
    For counter = 0 To 5
        grdTimes.Width = grdTimes.Width + grdTimes.ColWidth(counter)
    Next counter
    grdTimes.Width = grdTimes.Width + 5

    grdTimes.Height = 0
    For counter = 0 To 9
        grdTimes.Height = grdTimes.Height + grdTimes.RowHeight(counter) - 5
    Next counter

    grdTimes.FixedRows = 1
    grdTimes.FixedCols = 1

    LineX.X1 = grdTimes.Left - 1
    LineX.Y1 = grdTimes.Top - 1
    LineX.X2 = grdTimes.Width + 200
    LineX.Y2 = grdTimes.Top - 1

    LineY.X1 = grdTimes.Left - 1
    LineY.Y1 = grdTimes.Top - 1
    LineY.X2 = grdTimes.Left - 1
    LineY.Y2 = grdTimes.Height + 90
    
    LineX.ZOrder
    LineY.ZOrder

    i = 0
    For c = 1 To grdTimes.Cols - 1
      For ci = 1 To grdTimes.Rows - 1
        grdTimes.Row = ci
        grdTimes.Col = c
        grdTimes.Text = WBTI(i)
        i = i + 1
      Next ci
    Next c
    
    Me.Top = (screen.Height - Me.Height) / 2
    Me.Left = (screen.Width - Me.Width) / 2
End Sub

Sub grdTimes_DblClick ()
  If grdTimes.Row > 0 And grdTimes.Col > 0 Then grdTimes_KeyPress 13
End Sub

Sub grdTimes_KeyDown (KeyCode As Integer, Shift As Integer)
  txtTime.Visible = False
  If Shift = 1 Then Shifting = True Else Shifting = False
End Sub

Sub grdTimes_KeyPress (KeyAscii As Integer)
  Dim Char As String

  Select Case KeyAscii
  Case 27 'Escape key was pressed

  Case 9 'Tab
  If Shifting Then
    If grdTimes.Col > 1 Then
      grdTimes.Col = grdTimes.Col - 1
    End If
  Else
    If grdTimes.Col < (grdTimes.Cols - 1) Then
      grdTimes.Col = grdTimes.Col + 1
    End If
  End If
  
  Case Else
    If KeyAscii = 13 Then
      txtTime.Text = grdTimes.Text
      txtTime.SelStart = Len(txtTime.Text)
    Else
      Char = Chr$(KeyAscii)
      txtTime.Text = Char
      txtTime.SelStart = 1
    End If
    ShowTextBox
    KeyAscii = 0
  End Select
End Sub

Sub grdTimes_MouseDown (Button As Integer, Shift As Integer, X As Single, Y As Single)
  txtTime.Visible = False
End Sub

Sub grdTimes_RowColChange ()
  txtTime.Text = grdTimes.Text
End Sub

Sub txtTime_KeyDown (KeyCode As Integer, Shift As Integer)
  If Shift = 1 Then
    Shifting = True
  Else
    Shifting = False
  End If

  Select Case KeyCode
  Case 38 'Up
    txtTime_KeyPress 13 'Simulate enter key
    SendKeys "{UP}"
  Case 40 'Down
    txtTime_KeyPress 13 'Simulate enter key
    SendKeys "{DOWN}"
  End Select
End Sub

Sub txtTime_KeyPress (KeyAscii As Integer)
  Select Case KeyAscii
    Case 13
    grdTimes.Text = txtTime.Text
    txtTime.Visible = False
    grdTimes.SetFocus
    KeyAscii = 0
    
    Case 27
    KeyAscii = 0
    txtTime.Visible = False
    grdTimes.SetFocus
  End Select
End Sub

Sub txtTime_LostFocus ()
    'To bypass possible errors
    On Error Resume Next
    
    'Set to normal time format in case entered time is in army time
    grdTimes.Text = TimeValue(grdTimes.Text)
    'Check out time entered
    If IsDate(grdTimes.Text) Then 'Valid time entered
       'Add seconds part to time to ring
       grdTimes.Text = Format$(grdTimes.Text, "h:mm:ss AM/PM")
    ElseIf grdTimes.Text = "" Then
       'Do nothing
    Else
       'Not valid time tell user
       MsgBox "Time not valid.", 16, "Warning"
       'Highlight and give box focus back
       grdTimes_KeyPress 13
       txtTime.SelStart = 0
       txtTime.SelLength = Len(txtTime.Text)
    End If
End Sub

