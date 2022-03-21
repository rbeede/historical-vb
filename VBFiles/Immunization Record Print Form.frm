VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Begin VB.Form frmPrint 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Print"
   ClientHeight    =   1905
   ClientLeft      =   1095
   ClientTop       =   1485
   ClientWidth     =   3480
   ControlBox      =   0   'False
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1905
   ScaleWidth      =   3480
   Begin VB.CommandButton cmdoptions 
      Appearance      =   0  'Flat
      Caption         =   "&Setup Printer"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CommandButton cmdcancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2280
      TabIndex        =   4
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton cmdok 
      Appearance      =   0  'Flat
      Caption         =   "&OK"
      Height          =   375
      Left            =   2280
      TabIndex        =   3
      Top             =   120
      Width           =   1095
   End
   Begin VB.Frame frameoptions 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Print"
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1935
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Current Record"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   1695
      End
      Begin VB.OptionButton optall 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "All Records"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Value           =   -1  'True
         Width           =   1455
      End
   End
   Begin MSComDlg.CommonDialog cdbox 
      Left            =   2280
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327680
      FontSize        =   0
      MaxFileSize     =   256
   End
End
Attribute VB_Name = "frmPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim printall As Integer

Private Sub cmdcancel_click()
    FrmMain.Enabled = True
    Me.Hide
End Sub

Private Sub cmdok_Click()
    
    If optall.Value = True Then ' Print all option was selected
       
       printall = True ' Set flag to true

    Else ' Print current option was selected

       printall = False  ' Set flag to false

    End If

    Call cmdcancel_click ' Go to sub cmdcancel click event to close form

    Call printstudents ' go to sub printstudent to print out student(s)

End Sub

Private Sub cmdoptions_Click()
    
    ' Printer flag

    cdbox.Flags = pd_printsetup ' Tell dialog box to use printer setup
    
    cdbox.Action = 5 ' Tell dialog box to show printer flag
End Sub

Private Sub printstudents()
Dim i As Integer

   FrmMain.CmdAdd.Visible = False
   FrmMain.cmdSave.Visible = False
   FrmMain.cmdDelete.Visible = False
   FrmMain.cmdNext.Visible = False
   FrmMain.cmdBack.Visible = False
   FrmMain.cmdSearch.Visible = False
   FrmMain.cmdPrint.Visible = False
   FrmMain.cmdAbout.Visible = False
   FrmMain.cmdHelp.Visible = False
   FrmMain.cmdClose.Visible = False
   
   On Error Resume Next
   If printall = True Then

       FrmMain.Data1.Recordset.MoveFirst ' Move to first of database
       
       Do While Not FrmMain.Data1.Recordset.EOF ' Loop until end of database
          
          FrmMain.Refresh
          FrmMain.PrintForm

          FrmMain.Data1.Recordset.MoveNext ' Move to next record

       Loop ' Go back to do

       FrmMain.Data1.Refresh ' Refresh database
    
    Else

       FrmMain.Refresh
       For X% = 0 To 1000
          DoEvents
       Next X%
       FrmMain.PrintForm
    End If
   
   FrmMain.CmdAdd.Visible = True
   FrmMain.cmdSave.Visible = True
   FrmMain.cmdDelete.Visible = True
   FrmMain.cmdNext.Visible = True
   FrmMain.cmdBack.Visible = True
   FrmMain.cmdSearch.Visible = True
   FrmMain.cmdPrint.Visible = True
   FrmMain.cmdAbout.Visible = True
   FrmMain.cmdHelp.Visible = True
   FrmMain.cmdClose.Visible = True

End Sub

