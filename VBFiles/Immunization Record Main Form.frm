VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Immunization Record"
   ClientHeight    =   6825
   ClientLeft      =   1020
   ClientTop       =   105
   ClientWidth     =   7290
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
   Icon            =   "Immunization Record Main Form.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6825
   ScaleWidth      =   7290
   Begin VB.CommandButton cmdClose 
      Appearance      =   0  'Flat
      Caption         =   "Close"
      Height          =   975
      Left            =   6600
      TabIndex        =   56
      Top             =   5520
      Width           =   785
   End
   Begin VB.CommandButton cmdAbout 
      Appearance      =   0  'Flat
      Caption         =   "About"
      Height          =   975
      Left            =   6600
      TabIndex        =   54
      Top             =   4560
      Width           =   785
   End
   Begin VB.CommandButton cmdHelp 
      Appearance      =   0  'Flat
      Caption         =   "Help"
      Height          =   975
      Left            =   6600
      TabIndex        =   55
      Top             =   3600
      Width           =   785
   End
   Begin VB.CommandButton cmdBack 
      Appearance      =   0  'Flat
      Caption         =   "Back"
      Height          =   975
      Left            =   0
      TabIndex        =   51
      Top             =   5520
      Width           =   735
   End
   Begin VB.CommandButton cmdNext 
      Appearance      =   0  'Flat
      Caption         =   "Next"
      Height          =   975
      Left            =   0
      TabIndex        =   50
      Top             =   4560
      Width           =   735
   End
   Begin VB.CommandButton cmdDelete 
      Appearance      =   0  'Flat
      Caption         =   "Delete"
      Height          =   975
      Left            =   0
      TabIndex        =   49
      Top             =   3600
      Width           =   735
   End
   Begin MSComDlg.CommonDialog CMD 
      Left            =   120
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327681
      FontSize        =   0
      MaxFileSize     =   256
      PrinterDefault  =   0   'False
   End
   Begin VB.CommandButton cmdPrint 
      Appearance      =   0  'Flat
      Caption         =   "Print"
      Height          =   975
      Left            =   6600
      TabIndex        =   53
      Top             =   2640
      Width           =   785
   End
   Begin VB.CommandButton cmdSave 
      Appearance      =   0  'Flat
      Caption         =   "Save"
      Enabled         =   0   'False
      Height          =   975
      Left            =   0
      TabIndex        =   48
      Top             =   2640
      Width           =   735
   End
   Begin MSMask.MaskEdBox MaskedEdit2 
      DataSource      =   "Data1"
      Height          =   285
      Index           =   20
      Left            =   2760
      TabIndex        =   45
      Top             =   6480
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      _Version        =   327681
      Appearance      =   0
      BackColor       =   14745599
      ForeColor       =   0
      PromptInclude   =   0   'False
      MaxLength       =   8
      Mask            =   "##-##-##"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MaskedEdit2 
      DataSource      =   "Data1"
      Height          =   285
      Index           =   19
      Left            =   2760
      TabIndex        =   43
      Top             =   6240
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      _Version        =   327681
      Appearance      =   0
      BackColor       =   14745599
      ForeColor       =   0
      PromptInclude   =   0   'False
      MaxLength       =   8
      Mask            =   "##-##-##"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MaskedEdit2 
      DataSource      =   "Data1"
      Height          =   285
      Index           =   18
      Left            =   2760
      TabIndex        =   41
      Top             =   6000
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      _Version        =   327681
      Appearance      =   0
      BackColor       =   14745599
      ForeColor       =   0
      PromptInclude   =   0   'False
      MaxLength       =   8
      Mask            =   "##-##-##"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MaskedEdit2 
      DataSource      =   "Data1"
      Height          =   285
      Index           =   17
      Left            =   2760
      TabIndex        =   39
      Top             =   5760
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      _Version        =   327681
      Appearance      =   0
      BackColor       =   14745599
      ForeColor       =   0
      PromptInclude   =   0   'False
      MaxLength       =   8
      Mask            =   "##-##-##"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MaskedEdit2 
      DataSource      =   "Data1"
      Height          =   285
      Index           =   16
      Left            =   2760
      TabIndex        =   37
      Top             =   5520
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      _Version        =   327681
      Appearance      =   0
      BackColor       =   14745599
      ForeColor       =   0
      PromptInclude   =   0   'False
      MaxLength       =   8
      Mask            =   "##-##-##"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MaskedEdit2 
      DataSource      =   "Data1"
      Height          =   285
      Index           =   15
      Left            =   2760
      TabIndex        =   35
      Top             =   5280
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      _Version        =   327681
      Appearance      =   0
      BackColor       =   14745599
      ForeColor       =   0
      PromptInclude   =   0   'False
      MaxLength       =   8
      Mask            =   "##-##-##"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MaskedEdit2 
      DataSource      =   "Data1"
      Height          =   285
      Index           =   14
      Left            =   2760
      TabIndex        =   32
      Top             =   5040
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      _Version        =   327681
      Appearance      =   0
      BackColor       =   14745599
      ForeColor       =   0
      PromptInclude   =   0   'False
      MaxLength       =   8
      Mask            =   "##-##-##"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MaskedEdit2 
      DataSource      =   "Data1"
      Height          =   285
      Index           =   13
      Left            =   2760
      TabIndex        =   30
      Top             =   4800
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      _Version        =   327681
      Appearance      =   0
      BackColor       =   14745599
      ForeColor       =   0
      PromptInclude   =   0   'False
      MaxLength       =   8
      Mask            =   "##-##-##"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MaskedEdit2 
      DataSource      =   "Data1"
      Height          =   285
      Index           =   12
      Left            =   2760
      TabIndex        =   28
      Top             =   4560
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      _Version        =   327681
      Appearance      =   0
      BackColor       =   14745599
      ForeColor       =   0
      PromptInclude   =   0   'False
      MaxLength       =   8
      Mask            =   "##-##-##"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MaskedEdit2 
      DataSource      =   "Data1"
      Height          =   285
      Index           =   11
      Left            =   2760
      TabIndex        =   26
      Top             =   4320
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      _Version        =   327681
      Appearance      =   0
      BackColor       =   14745599
      ForeColor       =   0
      PromptInclude   =   0   'False
      MaxLength       =   8
      Mask            =   "##-##-##"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MaskedEdit2 
      DataSource      =   "Data1"
      Height          =   285
      Index           =   10
      Left            =   2760
      TabIndex        =   24
      Top             =   4080
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      _Version        =   327681
      Appearance      =   0
      BackColor       =   14745599
      ForeColor       =   0
      PromptInclude   =   0   'False
      MaxLength       =   8
      Mask            =   "##-##-##"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MaskedEdit2 
      DataSource      =   "Data1"
      Height          =   285
      Index           =   9
      Left            =   2760
      TabIndex        =   22
      Top             =   3840
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      _Version        =   327681
      Appearance      =   0
      BackColor       =   14745599
      ForeColor       =   0
      PromptInclude   =   0   'False
      MaxLength       =   8
      Mask            =   "##-##-##"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MaskedEdit2 
      DataSource      =   "Data1"
      Height          =   285
      Index           =   8
      Left            =   2760
      TabIndex        =   20
      Top             =   3600
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      _Version        =   327681
      Appearance      =   0
      BackColor       =   14745599
      ForeColor       =   0
      PromptInclude   =   0   'False
      MaxLength       =   8
      Mask            =   "##-##-##"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MaskedEdit2 
      DataSource      =   "Data1"
      Height          =   285
      Index           =   7
      Left            =   2760
      TabIndex        =   18
      Top             =   3360
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      _Version        =   327681
      Appearance      =   0
      BackColor       =   14745599
      ForeColor       =   0
      PromptInclude   =   0   'False
      MaxLength       =   8
      Mask            =   "##-##-##"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MaskedEdit2 
      DataSource      =   "Data1"
      Height          =   285
      Index           =   6
      Left            =   2760
      TabIndex        =   16
      Top             =   3120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      _Version        =   327681
      Appearance      =   0
      BackColor       =   14745599
      ForeColor       =   0
      PromptInclude   =   0   'False
      MaxLength       =   8
      Mask            =   "##-##-##"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MaskedEdit2 
      DataSource      =   "Data1"
      Height          =   285
      Index           =   5
      Left            =   2760
      TabIndex        =   14
      Top             =   2880
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      _Version        =   327681
      Appearance      =   0
      BackColor       =   14745599
      ForeColor       =   0
      PromptInclude   =   0   'False
      MaxLength       =   8
      Mask            =   "##-##-##"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MaskedEdit2 
      DataSource      =   "Data1"
      Height          =   285
      Index           =   4
      Left            =   2760
      TabIndex        =   12
      Top             =   2640
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      _Version        =   327681
      Appearance      =   0
      BackColor       =   14745599
      ForeColor       =   0
      PromptInclude   =   0   'False
      MaxLength       =   8
      Mask            =   "##-##-##"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MaskedEdit2 
      DataSource      =   "Data1"
      Height          =   285
      Index           =   3
      Left            =   2760
      TabIndex        =   10
      Top             =   2400
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      _Version        =   327681
      Appearance      =   0
      BackColor       =   14745599
      ForeColor       =   0
      PromptInclude   =   0   'False
      MaxLength       =   8
      Mask            =   "##-##-##"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MaskedEdit2 
      DataSource      =   "Data1"
      Height          =   285
      Index           =   2
      Left            =   2760
      TabIndex        =   8
      Top             =   2160
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      _Version        =   327681
      Appearance      =   0
      BackColor       =   14745599
      ForeColor       =   0
      PromptInclude   =   0   'False
      MaxLength       =   8
      Mask            =   "##-##-##"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MaskedEdit2 
      DataSource      =   "Data1"
      Height          =   285
      Index           =   1
      Left            =   2760
      TabIndex        =   6
      Top             =   1920
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      _Version        =   327681
      Appearance      =   0
      BackColor       =   14745599
      ForeColor       =   0
      PromptInclude   =   0   'False
      MaxLength       =   8
      Mask            =   "##-##-##"
      PromptChar      =   "_"
   End
   Begin VB.CommandButton cmdSearch 
      Appearance      =   0  'Flat
      Caption         =   "Search"
      Height          =   975
      Left            =   6600
      TabIndex        =   52
      Top             =   1680
      Width           =   785
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0FFFF&
      DataSource      =   "Data1"
      Height          =   285
      Index           =   20
      Left            =   3960
      MaxLength       =   22
      TabIndex        =   46
      Top             =   6480
      Width           =   2655
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0FFFF&
      DataSource      =   "Data1"
      Height          =   285
      Index           =   19
      Left            =   3960
      MaxLength       =   22
      TabIndex        =   44
      Top             =   6240
      Width           =   2655
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0FFFF&
      DataSource      =   "Data1"
      Height          =   285
      Index           =   18
      Left            =   3960
      MaxLength       =   22
      TabIndex        =   42
      Top             =   6000
      Width           =   2655
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0FFFF&
      DataSource      =   "Data1"
      Height          =   285
      Index           =   17
      Left            =   3960
      MaxLength       =   22
      TabIndex        =   40
      Top             =   5760
      Width           =   2655
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0FFFF&
      DataSource      =   "Data1"
      Height          =   285
      Index           =   16
      Left            =   3960
      MaxLength       =   22
      TabIndex        =   38
      Top             =   5520
      Width           =   2655
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0FFFF&
      DataSource      =   "Data1"
      Height          =   285
      Index           =   15
      Left            =   3960
      MaxLength       =   22
      TabIndex        =   36
      Top             =   5280
      Width           =   2655
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0FFFF&
      DataSource      =   "Data1"
      Height          =   285
      Index           =   14
      Left            =   3960
      MaxLength       =   22
      TabIndex        =   33
      Top             =   5040
      Width           =   2655
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0FFFF&
      DataSource      =   "Data1"
      Height          =   285
      Index           =   13
      Left            =   3960
      MaxLength       =   22
      TabIndex        =   31
      Top             =   4800
      Width           =   2655
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0FFFF&
      DataSource      =   "Data1"
      Height          =   285
      Index           =   12
      Left            =   3960
      MaxLength       =   22
      TabIndex        =   29
      Top             =   4560
      Width           =   2655
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0FFFF&
      DataSource      =   "Data1"
      Height          =   285
      Index           =   11
      Left            =   3960
      MaxLength       =   22
      TabIndex        =   27
      Top             =   4320
      Width           =   2655
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0FFFF&
      DataSource      =   "Data1"
      Height          =   285
      Index           =   10
      Left            =   3960
      MaxLength       =   22
      TabIndex        =   25
      Top             =   4080
      Width           =   2655
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0FFFF&
      DataSource      =   "Data1"
      Height          =   285
      Index           =   9
      Left            =   3960
      MaxLength       =   22
      TabIndex        =   23
      Top             =   3840
      Width           =   2655
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0FFFF&
      DataSource      =   "Data1"
      Height          =   285
      Index           =   8
      Left            =   3960
      MaxLength       =   22
      TabIndex        =   21
      Top             =   3600
      Width           =   2655
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0FFFF&
      DataSource      =   "Data1"
      Height          =   285
      Index           =   7
      Left            =   3960
      MaxLength       =   22
      TabIndex        =   19
      Top             =   3360
      Width           =   2655
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0FFFF&
      DataSource      =   "Data1"
      Height          =   285
      Index           =   6
      Left            =   3960
      MaxLength       =   22
      TabIndex        =   17
      Top             =   3120
      Width           =   2655
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0FFFF&
      DataSource      =   "Data1"
      Height          =   285
      Index           =   5
      Left            =   3960
      MaxLength       =   22
      TabIndex        =   15
      Top             =   2880
      Width           =   2655
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0FFFF&
      DataSource      =   "Data1"
      Height          =   285
      Index           =   4
      Left            =   3960
      MaxLength       =   22
      TabIndex        =   13
      Top             =   2640
      Width           =   2655
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0FFFF&
      DataSource      =   "Data1"
      Height          =   285
      Index           =   3
      Left            =   3960
      MaxLength       =   22
      TabIndex        =   11
      Top             =   2400
      Width           =   2655
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0FFFF&
      DataSource      =   "Data1"
      Height          =   285
      Index           =   2
      Left            =   3960
      MaxLength       =   22
      TabIndex        =   9
      Top             =   2160
      Width           =   2655
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0FFFF&
      DataSource      =   "Data1"
      Height          =   285
      Index           =   1
      Left            =   3960
      MaxLength       =   22
      TabIndex        =   7
      Top             =   1920
      Width           =   2655
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0FFFF&
      DataSource      =   "Data1"
      Height          =   285
      Index           =   0
      Left            =   3960
      MaxLength       =   22
      TabIndex        =   5
      Top             =   1680
      Width           =   2655
   End
   Begin MSMask.MaskEdBox MaskedEdit2 
      DataSource      =   "Data1"
      Height          =   285
      Index           =   0
      Left            =   2760
      TabIndex        =   4
      Top             =   1680
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      _Version        =   327681
      Appearance      =   0
      BackColor       =   14745599
      ForeColor       =   0
      PromptInclude   =   0   'False
      MaxLength       =   8
      Mask            =   "##-##-##"
      PromptChar      =   "_"
   End
   Begin VB.Data Data1 
      Appearance      =   0  'Flat
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   270
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   720
      Visible         =   0   'False
      Width           =   1140
   End
   Begin MSMask.MaskEdBox MaskedEdit1 
      DataSource      =   "Data1"
      Height          =   285
      Left            =   2160
      TabIndex        =   1
      Top             =   360
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   503
      _Version        =   327681
      Appearance      =   0
      BackColor       =   14745599
      ForeColor       =   0
      PromptInclude   =   0   'False
      MaxLength       =   8
      Mask            =   "##-##-##"
      PromptChar      =   "_"
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0FFFF&
      DataSource      =   "Data1"
      Height          =   285
      Left            =   2160
      MaxLength       =   40
      TabIndex        =   3
      Top             =   1080
      Width           =   4455
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0FFFF&
      DataSource      =   "Data1"
      Height          =   285
      Left            =   2160
      MaxLength       =   40
      TabIndex        =   2
      Top             =   720
      Width           =   4455
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0FFFF&
      DataField       =   "FullName"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   2160
      MaxLength       =   40
      TabIndex        =   0
      Top             =   0
      Width           =   4455
   End
   Begin VB.CommandButton CmdAdd 
      Appearance      =   0  'Flat
      Caption         =   "Add"
      Height          =   975
      Left            =   0
      TabIndex        =   47
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Other"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   20
      Left            =   720
      TabIndex        =   83
      Top             =   6480
      Width           =   2055
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Other"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   19
      Left            =   720
      TabIndex        =   82
      Top             =   6240
      Width           =   2055
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Influenza         4"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   18
      Left            =   720
      TabIndex        =   81
      Top             =   6000
      Width           =   2055
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Influenza         3"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   17
      Left            =   720
      TabIndex        =   80
      Top             =   5760
      Width           =   2055
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Influenza         2"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   16
      Left            =   720
      TabIndex        =   79
      Top             =   5520
      Width           =   2055
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Influenza         1"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   15
      Left            =   720
      TabIndex        =   78
      Top             =   5280
      Width           =   2055
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Pneumococcal"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   14
      Left            =   720
      TabIndex        =   77
      Top             =   5040
      Width           =   2055
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Hepatitis B       3"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   13
      Left            =   720
      TabIndex        =   76
      Top             =   4800
      Width           =   2055
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Hepatitis B       2"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   12
      Left            =   720
      TabIndex        =   75
      Top             =   4560
      Width           =   2055
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Hepatitis B       1"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   11
      Left            =   720
      TabIndex        =   74
      Top             =   4320
      Width           =   2055
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "HIB"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   10
      Left            =   720
      TabIndex        =   73
      Top             =   4080
      Width           =   2055
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "MMR/MR"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   9
      Left            =   720
      TabIndex        =   72
      Top             =   3840
      Width           =   2055
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "OPV/IPV           4"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   8
      Left            =   720
      TabIndex        =   71
      Top             =   3600
      Width           =   2055
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "OPV/IPV           3"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   7
      Left            =   720
      TabIndex        =   70
      Top             =   3360
      Width           =   2055
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "OPV/IPV           2"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   6
      Left            =   720
      TabIndex        =   69
      Top             =   3120
      Width           =   2055
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "OPV/IPV           1"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   5
      Left            =   720
      TabIndex        =   68
      Top             =   2880
      Width           =   2055
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DTP/PED DT/ Td    5"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   720
      TabIndex        =   67
      Top             =   2640
      Width           =   2055
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DTP/PED DT/ Td    4"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   720
      TabIndex        =   66
      Top             =   2400
      Width           =   2055
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DTP/PED DT/ Td    3"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   720
      TabIndex        =   65
      Top             =   2160
      Width           =   2055
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DTP/PED DT/ Td    2"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   720
      TabIndex        =   64
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DTP/PED DT/ Td    1"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   720
      TabIndex        =   63
      Top             =   1680
      Width           =   2055
   End
   Begin VB.Line Line4 
      X1              =   2760
      X2              =   2760
      Y1              =   1440
      Y2              =   6735
   End
   Begin VB.Line Line3 
      X1              =   3960
      X2              =   3960
      Y1              =   1440
      Y2              =   6735
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   720
      X2              =   6600
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00E0FFFF&
      Caption         =   "Doctor or Clinic"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   5040
      TabIndex        =   62
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00E0FFFF&
      Caption         =   "Date Given"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2880
      TabIndex        =   61
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00E0FFFF&
      Caption         =   "Vaccine"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   960
      TabIndex        =   60
      Top             =   1440
      Width           =   705
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Drug Sensitivity:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   720
      TabIndex        =   59
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Address:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1320
      TabIndex        =   58
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Date of Birth:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   960
      TabIndex        =   57
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Full Name:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1200
      TabIndex        =   34
      Top             =   0
      Width           =   975
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Adding As Integer
Dim Closing As Integer
Option Explicit

Private Sub cmdAbout_Click()
Dim Msg As String

Msg = "Program written by Rodney Beede." + Chr$(13)
Msg = Msg & "Send questions and comments to rodney_beede@hotmail.com" & Chr$(13)
Msg = Msg & "Program published by Infinisoft." & Chr$(13)
Msg = Msg & Chr$(13)
Msg = Msg & "Support my work, register this program!" & Chr$(13)
Msg = Msg & "Read the file ImmunRec.txt for more info!"

MsgBox Msg, 64, "About Immunization Record"
End Sub

Private Sub CmdAdd_Click()
   Dim i As Integer
   
   Adding = True
   Data1.Recordset.AddNew
   cmdDelete.Enabled = False
   cmdNext.Enabled = False
   cmdBack.Enabled = False
   cmdSearch.Enabled = False
   cmdPrint.Enabled = False
   CmdAdd.Enabled = False
   cmdSave.Enabled = True
   Text1.Enabled = True
   Text1.BackColor = &HE0FFFF
   MaskedEdit1.Enabled = True
   MaskedEdit1.BackColor = &HE0FFFF
   Text2.Enabled = True
   Text2.BackColor = &HE0FFFF
   Text3.Enabled = True
   Text3.BackColor = &HE0FFFF
   For i = 1 To 21
      MaskedEdit2(i - 1).Enabled = True
      MaskedEdit2(i - 1).BackColor = &HE0FFFF
      Text4(i - 1).Enabled = True
      Text4(i - 1).BackColor = &HE0FFFF
   Next i
   Text1.SetFocus
End Sub

Private Sub cmdBack_Click()
On Error Resume Next
Data1.Recordset.MovePrevious
If Data1.Recordset.BOF Then Data1.Recordset.MoveNext
End Sub

Private Sub cmdClose_Click()
   Unload Me
End Sub

Private Sub cmdDelete_Click()
   Dim result As Integer
   On Error Resume Next
   Beep
   result = MsgBox("Are you sure you want to delete this record?", 36, "Confirm")

   If result = 6 Then Data1.Recordset.Delete
   Data1.Recordset.MoveNext
   If Data1.Recordset.EOF Then Data1.Recordset.MovePrevious
   If Data1.Recordset.BOF And Data1.Recordset.EOF Then Call data1_reposition
End Sub

Private Sub cmdHelp_Click()
    'Bypass errors
    On Error Resume Next
    
    Dim VPath As String
    
    'Get programs path
    If Len(App.Path) > 3 Then
       VPath = App.Path + "\" 'Add \ character to path
    Else
       VPath = App.Path 'Set path
    End If

    'Check to see if help already started
    AppActivate "Immunization Record Help"
    
    'Check for error saying help is not already started
    If Err = 5 Then
       'Need to start help
       'Use Common Dialog Box to start help
       CMD.HelpFile = VPath + "Immunization Record.HLP"
       CMD.HelpCommand = &H3
       CMD.Action = 6
    End If

End Sub

Private Sub cmdNext_Click()
On Error Resume Next
Data1.Recordset.MoveNext
If Data1.Recordset.EOF Then Data1.Recordset.MovePrevious
End Sub

Private Sub cmdPrint_Click()
    If Data1.Recordset.BOF And Data1.Recordset.EOF Then
       
       MsgBox "Nothing to print.", 64, "Error"
    
       Exit Sub ' Leave sub
    End If

    frmPrint.Show ' Display the print form
    Me.Enabled = False ' Disable form
End Sub

Private Sub cmdSave_Click()
      On Error Resume Next
      Data1.Recordset.Update
      Data1.Refresh
      cmdDelete.Enabled = True
      cmdNext.Enabled = True
      cmdBack.Enabled = True
      cmdSearch.Enabled = True
      cmdPrint.Enabled = True
      CmdAdd.Enabled = True
      cmdPrint.Enabled = True
      cmdSave.Enabled = False
      Adding = False
End Sub

Private Sub cmdSearch_Click()
Me.Enabled = False
Load FrmSearch
End Sub

Private Sub data1_reposition()
Dim i As Integer

If Data1.Recordset.BOF And Data1.Recordset.EOF And Not Adding = True Then
   If Closing = False Then
      MsgBox "No move records in database.  Add one", 16, "Error"
   End If
   cmdDelete.Enabled = False
   cmdNext.Enabled = False
   cmdBack.Enabled = False
   cmdSearch.Enabled = False
   cmdPrint.Enabled = False
   Text1.Enabled = False
   Text1.BackColor = &HC0C0C0
   MaskedEdit1.Enabled = False
   MaskedEdit1.BackColor = &HC0C0C0
   Text2.Enabled = False
   Text2.BackColor = &HC0C0C0
   Text3.Enabled = False
   Text3.BackColor = &HC0C0C0
   For i = 1 To 21
      MaskedEdit2(i - 1).Enabled = False
      MaskedEdit2(i - 1).BackColor = &HC0C0C0
      Text4(i - 1).Enabled = False
      Text4(i - 1).BackColor = &HC0C0C0
   Next i
End If
End Sub


Private Sub Form_Load()
    Dim i As Integer
    Me.Top = -60
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Height = Screen.Height + 60

    Adding = False
    
    Text1.DataField = "FullName"
    MaskedEdit1.DataField = "DateofBirth"
    Text2.DataField = "Address"
    Text3.DataField = "DrugSen"
    For i = 1 To 21
       MaskedEdit2(i - 1).DataField = "Date" + Chr$(64 + i)
       Text4(i - 1).DataField = "Doctor" + Chr$(64 + i)
    Next i
    
    Data1.DatabaseName = App.Path + "\Immunization Record.mdb"
    Data1.RecordSource = "ShotRecords"
    Data1.Refresh
    
    Closing = True
    Data1.Refresh
    Closing = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    
    Closing = True
    Data1.Recordset.Update
    Data1.Refresh
    Data1.Recordset.Close
    Closing = False

    If Not Err = 3020 Then MsgBox Error, 16, "Error"

    End
End Sub

