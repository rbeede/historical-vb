VERSION 2.00
Begin Form frmPrint 
   BorderStyle     =   3  'Fixed Double
   Caption         =   "Print"
   ClientHeight    =   1905
   ClientLeft      =   1095
   ClientTop       =   1485
   ClientWidth     =   3480
   ControlBox      =   0   'False
   Height          =   2310
   Left            =   1035
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1905
   ScaleWidth      =   3480
   Top             =   1140
   Width           =   3600
   Begin CommandButton cmdoptions 
      Caption         =   "&Setup Printer"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   1335
   End
   Begin CommandButton cmdcancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2280
      TabIndex        =   4
      Top             =   600
      Width           =   1095
   End
   Begin CommandButton cmdok 
      Caption         =   "&OK"
      Height          =   375
      Left            =   2280
      TabIndex        =   3
      Top             =   120
      Width           =   1095
   End
   Begin Frame frameoptions 
      Caption         =   "Print"
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1935
      Begin OptionButton Option1 
         Caption         =   "Current Record"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   1695
      End
      Begin OptionButton optall 
         Caption         =   "All Records"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Value           =   -1  'True
         Width           =   1455
      End
   End
   Begin CommonDialog cdbox 
      Left            =   2280
      Top             =   1080
   End
End
Dim printall As Integer

Sub cmdcancel_click ()
    frmMain.Enabled = True
    Me.Hide
End Sub

Sub cmdok_Click ()
    
    If optall.Value = True Then ' Print all option was selected
       
       printall = True ' Set flag to true

    Else ' Print current option was selected

       printall = False  ' Set flag to false

    End If

    Call cmdcancel_click ' Go to sub cmdcancel click event to close form

    Call printstudents ' go to sub printstudent to print out student(s)

End Sub

Sub cmdoptions_Click ()
    
    ' Printer flag

    cdbox.Flags = pd_printsetup ' Tell dialog box to use printer setup
    
    cdbox.Action = 5 ' Tell dialog box to show printer flag
End Sub

Sub printstudents ()
Dim i As Integer

   frmMain.CmdAdd.Visible = False
   frmMain.cmdSave.Visible = False
   frmMain.cmdDelete.Visible = False
   frmMain.cmdNext.Visible = False
   frmMain.cmdBack.Visible = False
   frmMain.cmdSearch.Visible = False
   frmMain.cmdPrint.Visible = False
   frmMain.cmdAbout.Visible = False
   frmMain.cmdHelp.Visible = False
   frmMain.cmdClose.Visible = False
   
   On Error Resume Next
   If printall = True Then

       frmMain.Data1.Recordset.MoveFirst ' Move to first of database
       
       Do While Not frmMain.Data1.Recordset.EOF ' Loop until end of database
          
          frmMain.Refresh
          frmMain.PrintForm

          frmMain.Data1.Recordset.MoveNext ' Move to next record

       Loop ' Go back to do

       frmMain.Data1.Refresh ' Refresh database
    
    Else

       frmMain.Refresh
       frmMain.PrintForm
    End If
   
   frmMain.CmdAdd.Visible = True
   frmMain.cmdSave.Visible = True
   frmMain.cmdDelete.Visible = True
   frmMain.cmdNext.Visible = True
   frmMain.cmdBack.Visible = True
   frmMain.cmdSearch.Visible = True
   frmMain.cmdPrint.Visible = True
   frmMain.cmdAbout.Visible = True
   frmMain.cmdHelp.Visible = True
   frmMain.cmdClose.Visible = True

End Sub

