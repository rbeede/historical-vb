VERSION 5.00
Object = "{A8B3B723-0B5A-101B-B22E-00AA0037B2FC}#1.0#0"; "GRID32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Web Page Calander Code Maker"
   ClientHeight    =   4770
   ClientLeft      =   2385
   ClientTop       =   2040
   ClientWidth     =   5055
   Icon            =   "Web Page Calender Code Maker Main Form.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4770
   ScaleWidth      =   5055
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   1560
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   2760
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtCaption 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Text            =   "Caption Here"
      Top             =   120
      Width           =   4815
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "&Quit"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   3720
      Width           =   1335
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear &Grid"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   3240
      Width           =   1335
   End
   Begin VB.CommandButton cmdCode 
      Caption         =   "&Code It"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   2760
      Width           =   1335
   End
   Begin MSGrid.Grid grdCalender 
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   4815
      _Version        =   65536
      _ExtentX        =   8493
      _ExtentY        =   5318
      _StockProps     =   77
      BackColor       =   16777215
      Rows            =   5
      Cols            =   7
      FixedRows       =   0
      FixedCols       =   0
      ScrollBars      =   0
      HighLight       =   0   'False
   End
   Begin VB.Image imgNumber 
      BorderStyle     =   1  'Fixed Single
      Height          =   330
      Index           =   30
      Left            =   4080
      Picture         =   "Web Page Calender Code Maker Main Form.frx":0442
      Top             =   4320
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgNumber 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   29
      Left            =   3720
      Picture         =   "Web Page Calender Code Maker Main Form.frx":04D4
      Top             =   4320
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image imgNumber 
      BorderStyle     =   1  'Fixed Single
      Height          =   270
      Index           =   28
      Left            =   3360
      Picture         =   "Web Page Calender Code Maker Main Form.frx":055E
      Top             =   4320
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image imgNumber 
      BorderStyle     =   1  'Fixed Single
      Height          =   270
      Index           =   27
      Left            =   3000
      Picture         =   "Web Page Calender Code Maker Main Form.frx":05E0
      Top             =   4320
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image imgNumber 
      BorderStyle     =   1  'Fixed Single
      Height          =   330
      Index           =   26
      Left            =   2640
      Picture         =   "Web Page Calender Code Maker Main Form.frx":0662
      Top             =   4320
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image imgNumber 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   25
      Left            =   2280
      Picture         =   "Web Page Calender Code Maker Main Form.frx":06F4
      Top             =   4320
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image imgNumber 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   24
      Left            =   1920
      Picture         =   "Web Page Calender Code Maker Main Form.frx":077E
      Top             =   4320
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Image imgNumber 
      BorderStyle     =   1  'Fixed Single
      Height          =   330
      Index           =   23
      Left            =   1560
      Picture         =   "Web Page Calender Code Maker Main Form.frx":080C
      Top             =   4320
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image imgNumber 
      BorderStyle     =   1  'Fixed Single
      Height          =   330
      Index           =   22
      Left            =   4800
      Picture         =   "Web Page Calender Code Maker Main Form.frx":089E
      Top             =   3960
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Image imgNumber 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Index           =   21
      Left            =   4440
      Picture         =   "Web Page Calender Code Maker Main Form.frx":0930
      Top             =   3960
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Image imgNumber 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Index           =   20
      Left            =   4080
      Picture         =   "Web Page Calender Code Maker Main Form.frx":09C6
      Top             =   3960
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image imgNumber 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   19
      Left            =   3720
      Picture         =   "Web Page Calender Code Maker Main Form.frx":0A5C
      Top             =   3960
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgNumber 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   18
      Left            =   3360
      Picture         =   "Web Page Calender Code Maker Main Form.frx":0AEA
      Top             =   3960
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image imgNumber 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   17
      Left            =   3000
      Picture         =   "Web Page Calender Code Maker Main Form.frx":0B78
      Top             =   3960
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image imgNumber 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   16
      Left            =   2640
      Picture         =   "Web Page Calender Code Maker Main Form.frx":0C06
      Top             =   3960
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image imgNumber 
      BorderStyle     =   1  'Fixed Single
      Height          =   270
      Index           =   15
      Left            =   2280
      Picture         =   "Web Page Calender Code Maker Main Form.frx":0C90
      Top             =   3960
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image imgNumber 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   14
      Left            =   1920
      Picture         =   "Web Page Calender Code Maker Main Form.frx":0D12
      Top             =   3960
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image imgNumber 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   13
      Left            =   1560
      Picture         =   "Web Page Calender Code Maker Main Form.frx":0D98
      Top             =   3960
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image imgNumber 
      BorderStyle     =   1  'Fixed Single
      Height          =   270
      Index           =   12
      Left            =   4680
      Picture         =   "Web Page Calender Code Maker Main Form.frx":0E16
      Top             =   3600
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image imgNumber 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   11
      Left            =   4440
      Picture         =   "Web Page Calender Code Maker Main Form.frx":0E98
      Top             =   3600
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image imgNumber 
      BorderStyle     =   1  'Fixed Single
      Height          =   270
      Index           =   10
      Left            =   4080
      Picture         =   "Web Page Calender Code Maker Main Form.frx":0F16
      Top             =   3600
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Image imgNumber 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   9
      Left            =   3720
      Picture         =   "Web Page Calender Code Maker Main Form.frx":0F98
      Top             =   3600
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image imgNumber 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   8
      Left            =   3480
      Picture         =   "Web Page Calender Code Maker Main Form.frx":101E
      Top             =   3600
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgNumber 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   7
      Left            =   3240
      Picture         =   "Web Page Calender Code Maker Main Form.frx":10A4
      Top             =   3600
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Image imgNumber 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Index           =   6
      Left            =   3000
      Picture         =   "Web Page Calender Code Maker Main Form.frx":112E
      Top             =   3600
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imgNumber 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   5
      Left            =   2640
      Picture         =   "Web Page Calender Code Maker Main Form.frx":11C4
      Top             =   3600
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Image imgNumber 
      BorderStyle     =   1  'Fixed Single
      Height          =   270
      Index           =   4
      Left            =   2400
      Picture         =   "Web Page Calender Code Maker Main Form.frx":1242
      Top             =   3600
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image imgNumber 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   3
      Left            =   2160
      Picture         =   "Web Page Calender Code Maker Main Form.frx":12C4
      Top             =   3600
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgNumber 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   2
      Left            =   1920
      Picture         =   "Web Page Calender Code Maker Main Form.frx":134E
      Top             =   3600
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image imgNumber 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   1
      Left            =   1680
      Picture         =   "Web Page Calender Code Maker Main Form.frx":13DC
      Top             =   3600
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgNumber 
      BorderStyle     =   1  'Fixed Single
      Height          =   270
      Index           =   0
      Left            =   1560
      Picture         =   "Web Page Calender Code Maker Main Form.frx":146A
      Top             =   3600
      Visible         =   0   'False
      Width           =   150
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim GridDate(42) As String

Private Sub ShowTextBox()
  Dim TextX As Integer, TestY As Integer 'Test variables
  Dim C As Integer 'Loop counter
  Dim frmM As frmMain

  Set frmM = frmMain 'For easier use

  'Hide the text box and make it two lines tall and wide
  frmM.txtInfo.Visible = False
  frmM.txtInfo.Height = frmM!grdCalender.RowHeight(frmM!grdCalender.Row) - (Screen.TwipsPerPixelY * 2)
  frmM.txtInfo.Width = frmM!grdCalender.ColWidth(frmM!grdCalender.Col) - (Screen.TwipsPerPixelX * 2)
  
  'Determine X coordinate of the current cell

  TestX = frmM!grdCalender.Left + (Screen.TwipsPerPixelX * 3)

  'Sum all column widths
  For C = frmM!grdCalender.LeftCol To frmM!grdCalender.Col - 1
    TestX = TestX + frmM!grdCalender.ColWidth(C) + Screen.TwipsPerPixelX
  Next C

  'Determine Y coordinate of the current cell

  TestY = frmM!grdCalender.Top + frmM!grdCalender.RowHeight(0) + (Screen.TwipsPerPixelY * 3)
   
  'Sum all column heights
  For C = frmM!grdCalender.TopRow To frmM!grdCalender.Row - 1
    TestY = TestY + frmM!grdCalender.RowHeight(C) + Screen.TwipsPerPixelY
  Next C
   
  'Position the text box control
  frmM!txtInfo.Left = TestX
  frmM!txtInfo.Top = TestY
  
  frmM!txtInfo.ZOrder  'Make sure on top of grid
  frmM!txtInfo.Visible = True 'Show text box
  frmM!txtInfo.SetFocus 'Give focus to text box
End Sub
 
Private Sub cmdClear_Click()
    Dim I As Integer, C As Integer, R As Integer
    Dim calDate As String, calMonth As String
    Dim calYear As String, Temp As String
    Dim LeapYear As Integer, DayOfWeek As Integer
    Dim NumDays As Integer
    
    For I = 1 To 6
        For C = 0 To 6
            grdCalender.Row = I
            grdCalender.Col = C
            grdCalender.Text = ""
            grdCalender.Picture = LoadPicture()
        Next C
    Next I

    Do Until calDate <> ""
        calDate = InputBox$("Please enter a date for calender.  It should look something like 01-01-1997", "Enter a date", Date$)
    Loop
    
    If IsDate(calDate) = False Then calDate = "01-01-1997"
    
    calMonth = Month(calDate)
    calYear = Year(calDate)
    
    Select Case calMonth
        Case 1
            txtCaption.Text = "January"
        Case 2
            txtCaption.Text = "February"
        Case 3
            txtCaption.Text = "March"
        Case 4
            txtCaption.Text = "April"
        Case 5
            txtCaption.Text = "May"
        Case 6
            txtCaption.Text = "June"
        Case 7
            txtCaption.Text = "July"
        Case 8
            txtCaption.Text = "August"
        Case 9
            txtCaption.Text = "September"
        Case 10
            txtCaption.Text = "October"
        Case 11
            txtCaption.Text = "November"
        Case 12
            txtCaption.Text = "December"
    End Select

    DayOfWeek = WeekDay(calMonth & "-" & "01-" & calYear)
        
    If calYear Mod 4 = 0 Then LeapYear = True

    NumDays = 31
    Select Case calMonth
        Case 2
            NumDays = 28
            If LeapYear = True Then NumDays = 29
        Case 4
            NumDays = 30
        Case 6
            NumDays = 30
        Case 9
            NumDays = 30
        Case 11
            NumDays = 30
    End Select
    
    grdCalender.Row = 1
    grdCalender.Col = 0
    Temp = DayOfWeek - 1
    For I = 1 To NumDays
        If grdCalender.Col = 6 And grdCalender.Row = 1 Then
           grdCalender.Row = 2
           Temp = 0
        ElseIf grdCalender.Col = 6 And grdCalender.Row = 2 Then
           grdCalender.Row = 3
           Temp = 0
        ElseIf grdCalender.Col = 6 And grdCalender.Row = 3 Then
           grdCalender.Row = 4
           Temp = 0
        ElseIf grdCalender.Col = 6 And grdCalender.Row = 4 Then
           grdCalender.Row = 5
           Temp = 0
        ElseIf grdCalender.Col = 6 And grdCalender.Row = 5 Then
           grdCalender.Row = 6
           Temp = 0
        End If
        
        grdCalender.Col = Temp
        grdCalender.Picture = imgNumber(I - 1).Picture
        Temp = Temp + 1
    Next I

    For I = 0 To 42
        GridDate(I) = ""
    Next I
    
    Temp = 1
    For I = DayOfWeek - 1 To 41
        If Temp <= NumDays Then
            GridDate(I) = Temp
            Temp = Temp + 1
        End If
    Next I

End Sub

Private Sub cmdCode_Click()
    Dim CRLF As String, Code As String, Temp As String
    Dim I As Integer, R As Integer
    
    CRLF = Chr$(13) + Chr$(10)
    
    Code = ""
    Code = Code + "<TABLE BORDER>" + CRLF
    Code = Code + "  <CAPTION ALIGN=CENTER VALIGN=TOP>" + frmMain!txtCaption.Text + "</CAPTION>" + CRLF
    Code = Code + "  <TR BGCOLOR=orange VALIGN=TOP>" + CRLF
    Code = Code + "    <TH>Sunday</TH><TH>Monday</TH><TH>Tuesday</TH><TH>Wednesday</TH><TH>Thursday</TH><TH>Friday</TH><TH>Saturday</TH>" + CRLF
    Code = Code + "  </TR>" + CRLF
    Code = Code + "" + CRLF
     
    Temp = 0
    For R = 1 To 6
        Code = Code + "  <TR VALIGN=TOP>" + CRLF
        For I = 0 To 6
            grdCalender.Row = R
            grdCalender.Col = I
            If grdCalender.Text = "" Then
                Code = Code + "    <TD ALIGN=RIGHT><BIG>" + GridDate(I + Temp) + "</BIG>" + CRLF
            Else
                Code = Code + "    <TD ALIGN=RIGHT><BIG>" + GridDate(I + Temp) + "</BIG><BR>" + CRLF
            End If
            
            If grdCalender.Text = "" Then grdCalender.Text = "  "
            Code = Code + "      " + grdCalender.Text + CRLF
            Code = Code + "    </TD>" + CRLF
            Code = Code + "" + CRLF
        Next I
        Code = Code + "  </TR>" + CRLF
        Code = Code + "" + CRLF
        Temp = Temp + 7
    Next R
    Code = Code + "</TABLE>"
    


    Load frmCode
    frmCode!txtCode.Text = Code
    frmCode.Show
End Sub

Private Sub cmdQuit_Click()
    End
End Sub

Private Sub Form_Load()
    Dim Temp As String, TB As String
    Dim I As Integer, C As Integer, R As Integer
    
    Me.Width = Screen.Width
    
    TB = Chr$(9)
    
    grdCalender.FixedRows = 0
    grdCalender.FixedCols = 0
    grdCalender.Rows = 1
    grdCalender.Cols = 7
    
    Temp = "Sunday" & TB & "Monday" & TB & "Tuesday" & TB & "Wednesday" & TB & "Thursday" & TB & "Friday" & TB & "Saturday"
    grdCalender.AddItem Temp
    grdCalender.RemoveItem 0
    
    grdCalender.Rows = 7
    grdCalender.FixedRows = 1

    grdCalender.Row = 0
    For C = 0 To 6
        grdCalender.Col = C
        grdCalender.ColWidth(C) = 1080
    Next C

    grdCalender.Col = 0
    For R = 1 To 6
        grdCalender.Row = R
        grdCalender.RowHeight(R) = 540
    Next R

    grdCalender.Height = 540 * 6 + grdCalender.RowHeight(0) + 300
    grdCalender.Width = 1080 * 7 + 300
         
    Me.Height = grdCalender.Top + grdCalender.Height + 500
    
    cmdCode.Top = grdCalender.Top
    cmdCode.Left = grdCalender.Left + grdCalender.Width + 120

    cmdClear.Top = cmdCode.Top + cmdCode.Height + 120
    cmdClear.Left = cmdCode.Left
    
    cmdQuit.Top = cmdClear.Top + cmdClear.Height + 120
    cmdQuit.Left = cmdClear.Left

    txtCaption.Width = grdCalender.Width
   
    Call cmdClear_Click
End Sub

Private Sub grdCalender_DblClick()
    Call grdCalender_KeyPress(13)
End Sub

Private Sub grdCalender_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
      txtInfo.Text = grdCalender.Text
      txtInfo.SelStart = Len(txtInfo.Text)
    Else
      Char = Chr$(KeyAscii)
      txtInfo.Text = Char
      txtInfo.SelStart = 1
    End If
    ShowTextBox
    KeyAscii = 0
End Sub

Private Sub grdCalender_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    txtInfo.Visible = False
End Sub

Private Sub txtInfo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then
        txtInfo.Text = txtInfo.Text + "<BR>"
        txtInfo.SelStart = Len(txtInfo.Text)
        KeyAscii = 0  'Clear out key
    End If
End Sub

Private Sub txtInfo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        grdCalender.Text = txtInfo.Text
        txtInfo.Visible = False
        grdCalender.SetFocus
        KeyAscii = 0  'Clear out key
    End If
End Sub

