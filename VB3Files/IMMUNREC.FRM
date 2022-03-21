VERSION 2.00
Begin Form FrmMain 
   BackColor       =   &H00E0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Immunization Record"
   ClientHeight    =   6915
   ClientLeft      =   1020
   ClientTop       =   105
   ClientWidth     =   7320
   Height          =   7320
   Icon            =   IMMUNREC.FRX:0000
   Left            =   960
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6915
   ScaleWidth      =   7320
   Top             =   -240
   Width           =   7440
   Begin CommonDialog CMD 
      Left            =   120
      PrinterDefault  =   0   'False
      Top             =   120
   End
   Begin CommandButton cmdPrint 
      Caption         =   "Print"
      Height          =   975
      Left            =   6600
      TabIndex        =   55
      Top             =   2640
      Width           =   735
   End
   Begin CommandButton cmdSave 
      Caption         =   "Save"
      Enabled         =   0   'False
      Height          =   975
      Left            =   0
      TabIndex        =   50
      Top             =   2640
      Width           =   735
   End
   Begin CommandButton cmdAbout 
      Caption         =   "About"
      Height          =   975
      Left            =   6600
      TabIndex        =   56
      Top             =   4560
      Width           =   735
   End
   Begin CommandButton cmdHelp 
      Caption         =   "Help"
      Height          =   975
      Left            =   6600
      TabIndex        =   57
      Top             =   3600
      Width           =   735
   End
   Begin MaskEdBox MaskedEdit2 
      BackColor       =   &H00E0FFFF&
      DataSource      =   "Data1"
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   21
      Left            =   2760
      Mask            =   "##-##-##"
      MaxLength       =   8
      PromptChar      =   "_"
      PromptInclude   =   0   'False
      TabIndex        =   47
      Top             =   6720
      Width           =   1215
   End
   Begin MaskEdBox MaskedEdit2 
      BackColor       =   &H00E0FFFF&
      DataSource      =   "Data1"
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   20
      Left            =   2760
      Mask            =   "##-##-##"
      MaxLength       =   8
      PromptChar      =   "_"
      PromptInclude   =   0   'False
      TabIndex        =   45
      Top             =   6480
      Width           =   1215
   End
   Begin MaskEdBox MaskedEdit2 
      BackColor       =   &H00E0FFFF&
      DataSource      =   "Data1"
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   19
      Left            =   2760
      Mask            =   "##-##-##"
      MaxLength       =   8
      PromptChar      =   "_"
      PromptInclude   =   0   'False
      TabIndex        =   43
      Top             =   6240
      Width           =   1215
   End
   Begin MaskEdBox MaskedEdit2 
      BackColor       =   &H00E0FFFF&
      DataSource      =   "Data1"
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   18
      Left            =   2760
      Mask            =   "##-##-##"
      MaxLength       =   8
      PromptChar      =   "_"
      PromptInclude   =   0   'False
      TabIndex        =   41
      Top             =   6000
      Width           =   1215
   End
   Begin MaskEdBox MaskedEdit2 
      BackColor       =   &H00E0FFFF&
      DataSource      =   "Data1"
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   17
      Left            =   2760
      Mask            =   "##-##-##"
      MaxLength       =   8
      PromptChar      =   "_"
      PromptInclude   =   0   'False
      TabIndex        =   39
      Top             =   5760
      Width           =   1215
   End
   Begin MaskEdBox MaskedEdit2 
      BackColor       =   &H00E0FFFF&
      DataSource      =   "Data1"
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   16
      Left            =   2760
      Mask            =   "##-##-##"
      MaxLength       =   8
      PromptChar      =   "_"
      PromptInclude   =   0   'False
      TabIndex        =   37
      Top             =   5520
      Width           =   1215
   End
   Begin MaskEdBox MaskedEdit2 
      BackColor       =   &H00E0FFFF&
      DataSource      =   "Data1"
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   15
      Left            =   2760
      Mask            =   "##-##-##"
      MaxLength       =   8
      PromptChar      =   "_"
      PromptInclude   =   0   'False
      TabIndex        =   35
      Top             =   5280
      Width           =   1215
   End
   Begin MaskEdBox MaskedEdit2 
      BackColor       =   &H00E0FFFF&
      DataSource      =   "Data1"
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   14
      Left            =   2760
      Mask            =   "##-##-##"
      MaxLength       =   8
      PromptChar      =   "_"
      PromptInclude   =   0   'False
      TabIndex        =   32
      Top             =   5040
      Width           =   1215
   End
   Begin MaskEdBox MaskedEdit2 
      BackColor       =   &H00E0FFFF&
      DataSource      =   "Data1"
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   13
      Left            =   2760
      Mask            =   "##-##-##"
      MaxLength       =   8
      PromptChar      =   "_"
      PromptInclude   =   0   'False
      TabIndex        =   30
      Top             =   4800
      Width           =   1215
   End
   Begin MaskEdBox MaskedEdit2 
      BackColor       =   &H00E0FFFF&
      DataSource      =   "Data1"
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   12
      Left            =   2760
      Mask            =   "##-##-##"
      MaxLength       =   8
      PromptChar      =   "_"
      PromptInclude   =   0   'False
      TabIndex        =   28
      Top             =   4560
      Width           =   1215
   End
   Begin MaskEdBox MaskedEdit2 
      BackColor       =   &H00E0FFFF&
      DataSource      =   "Data1"
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   11
      Left            =   2760
      Mask            =   "##-##-##"
      MaxLength       =   8
      PromptChar      =   "_"
      PromptInclude   =   0   'False
      TabIndex        =   26
      Top             =   4320
      Width           =   1215
   End
   Begin MaskEdBox MaskedEdit2 
      BackColor       =   &H00E0FFFF&
      DataSource      =   "Data1"
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   10
      Left            =   2760
      Mask            =   "##-##-##"
      MaxLength       =   8
      PromptChar      =   "_"
      PromptInclude   =   0   'False
      TabIndex        =   24
      Top             =   4080
      Width           =   1215
   End
   Begin MaskEdBox MaskedEdit2 
      BackColor       =   &H00E0FFFF&
      DataSource      =   "Data1"
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   9
      Left            =   2760
      Mask            =   "##-##-##"
      MaxLength       =   8
      PromptChar      =   "_"
      PromptInclude   =   0   'False
      TabIndex        =   22
      Top             =   3840
      Width           =   1215
   End
   Begin MaskEdBox MaskedEdit2 
      BackColor       =   &H00E0FFFF&
      DataSource      =   "Data1"
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   8
      Left            =   2760
      Mask            =   "##-##-##"
      MaxLength       =   8
      PromptChar      =   "_"
      PromptInclude   =   0   'False
      TabIndex        =   20
      Top             =   3600
      Width           =   1215
   End
   Begin MaskEdBox MaskedEdit2 
      BackColor       =   &H00E0FFFF&
      DataSource      =   "Data1"
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   7
      Left            =   2760
      Mask            =   "##-##-##"
      MaxLength       =   8
      PromptChar      =   "_"
      PromptInclude   =   0   'False
      TabIndex        =   18
      Top             =   3360
      Width           =   1215
   End
   Begin MaskEdBox MaskedEdit2 
      BackColor       =   &H00E0FFFF&
      DataSource      =   "Data1"
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   6
      Left            =   2760
      Mask            =   "##-##-##"
      MaxLength       =   8
      PromptChar      =   "_"
      PromptInclude   =   0   'False
      TabIndex        =   16
      Top             =   3120
      Width           =   1215
   End
   Begin MaskEdBox MaskedEdit2 
      BackColor       =   &H00E0FFFF&
      DataSource      =   "Data1"
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   5
      Left            =   2760
      Mask            =   "##-##-##"
      MaxLength       =   8
      PromptChar      =   "_"
      PromptInclude   =   0   'False
      TabIndex        =   14
      Top             =   2880
      Width           =   1215
   End
   Begin MaskEdBox MaskedEdit2 
      BackColor       =   &H00E0FFFF&
      DataSource      =   "Data1"
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   4
      Left            =   2760
      Mask            =   "##-##-##"
      MaxLength       =   8
      PromptChar      =   "_"
      PromptInclude   =   0   'False
      TabIndex        =   12
      Top             =   2640
      Width           =   1215
   End
   Begin MaskEdBox MaskedEdit2 
      BackColor       =   &H00E0FFFF&
      DataSource      =   "Data1"
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   3
      Left            =   2760
      Mask            =   "##-##-##"
      MaxLength       =   8
      PromptChar      =   "_"
      PromptInclude   =   0   'False
      TabIndex        =   10
      Top             =   2400
      Width           =   1215
   End
   Begin MaskEdBox MaskedEdit2 
      BackColor       =   &H00E0FFFF&
      DataSource      =   "Data1"
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   2
      Left            =   2760
      Mask            =   "##-##-##"
      MaxLength       =   8
      PromptChar      =   "_"
      PromptInclude   =   0   'False
      TabIndex        =   8
      Top             =   2160
      Width           =   1215
   End
   Begin MaskEdBox MaskedEdit2 
      BackColor       =   &H00E0FFFF&
      DataSource      =   "Data1"
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   1
      Left            =   2760
      Mask            =   "##-##-##"
      MaxLength       =   8
      PromptChar      =   "_"
      PromptInclude   =   0   'False
      TabIndex        =   6
      Top             =   1920
      Width           =   1215
   End
   Begin CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   975
      Left            =   6600
      TabIndex        =   58
      Top             =   5520
      Width           =   735
   End
   Begin CommandButton cmdSearch 
      Caption         =   "Search"
      Height          =   975
      Left            =   6600
      TabIndex        =   54
      Top             =   1680
      Width           =   735
   End
   Begin CommandButton cmdBack 
      Caption         =   "Back"
      Height          =   975
      Left            =   0
      TabIndex        =   53
      Top             =   5520
      Width           =   735
   End
   Begin CommandButton cmdNext 
      Caption         =   "Next"
      Height          =   975
      Left            =   0
      TabIndex        =   52
      Top             =   4560
      Width           =   735
   End
   Begin CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   975
      Left            =   0
      TabIndex        =   51
      Top             =   3600
      Width           =   735
   End
   Begin CommandButton CmdAdd 
      Caption         =   "Add"
      Height          =   975
      Left            =   0
      TabIndex        =   49
      Top             =   1680
      Width           =   735
   End
   Begin TextBox Text4 
      BackColor       =   &H00E0FFFF&
      DataSource      =   "Data1"
      Height          =   285
      Index           =   21
      Left            =   3960
      MaxLength       =   22
      TabIndex        =   48
      Top             =   6720
      Width           =   2655
   End
   Begin TextBox Text4 
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
   Begin TextBox Text4 
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
   Begin TextBox Text4 
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
   Begin TextBox Text4 
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
   Begin TextBox Text4 
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
   Begin TextBox Text4 
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
   Begin TextBox Text4 
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
   Begin TextBox Text4 
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
   Begin TextBox Text4 
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
   Begin TextBox Text4 
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
   Begin TextBox Text4 
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
   Begin TextBox Text4 
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
   Begin TextBox Text4 
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
   Begin TextBox Text4 
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
   Begin TextBox Text4 
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
   Begin TextBox Text4 
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
   Begin TextBox Text4 
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
   Begin TextBox Text4 
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
   Begin TextBox Text4 
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
   Begin TextBox Text4 
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
   Begin TextBox Text4 
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
   Begin MaskEdBox MaskedEdit2 
      BackColor       =   &H00E0FFFF&
      DataSource      =   "Data1"
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   0
      Left            =   2760
      Mask            =   "##-##-##"
      MaxLength       =   8
      PromptChar      =   "_"
      PromptInclude   =   0   'False
      TabIndex        =   4
      Top             =   1680
      Width           =   1215
   End
   Begin Data Data1 
      Caption         =   "Data1"
      Connect         =   ""
      DatabaseName    =   ""
      Exclusive       =   0   'False
      Height          =   270
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordSource    =   ""
      Top             =   720
      Visible         =   0   'False
      Width           =   1140
   End
   Begin MaskEdBox MaskedEdit1 
      BackColor       =   &H00E0FFFF&
      DataSource      =   "Data1"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   2160
      Mask            =   "##-##-##"
      MaxLength       =   8
      PromptChar      =   "_"
      PromptInclude   =   0   'False
      TabIndex        =   1
      Top             =   360
      Width           =   4455
   End
   Begin TextBox Text3 
      BackColor       =   &H00E0FFFF&
      DataSource      =   "Data1"
      Height          =   285
      Left            =   2160
      MaxLength       =   40
      TabIndex        =   3
      Top             =   1080
      Width           =   4455
   End
   Begin TextBox Text2 
      BackColor       =   &H00E0FFFF&
      DataSource      =   "Data1"
      Height          =   285
      Left            =   2160
      MaxLength       =   40
      TabIndex        =   2
      Top             =   720
      Width           =   4455
   End
   Begin TextBox Text1 
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
   Begin Label Label8 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Other                         "
      Height          =   255
      Index           =   21
      Left            =   720
      TabIndex        =   86
      Top             =   6720
      Width           =   2055
   End
   Begin Label Label8 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Other                         "
      Height          =   255
      Index           =   20
      Left            =   720
      TabIndex        =   85
      Top             =   6480
      Width           =   2055
   End
   Begin Label Label8 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Other                         "
      Height          =   255
      Index           =   19
      Left            =   720
      TabIndex        =   84
      Top             =   6240
      Width           =   2055
   End
   Begin Label Label8 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Influenza                  4"
      Height          =   255
      Index           =   18
      Left            =   720
      TabIndex        =   83
      Top             =   6000
      Width           =   2055
   End
   Begin Label Label8 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Influenza                  3"
      Height          =   255
      Index           =   17
      Left            =   720
      TabIndex        =   82
      Top             =   5760
      Width           =   2055
   End
   Begin Label Label8 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Influenza                  2"
      Height          =   255
      Index           =   16
      Left            =   720
      TabIndex        =   81
      Top             =   5520
      Width           =   2055
   End
   Begin Label Label8 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Influenza                  1"
      Height          =   255
      Index           =   15
      Left            =   720
      TabIndex        =   80
      Top             =   5280
      Width           =   2055
   End
   Begin Label Label8 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Pneumococcal            "
      Height          =   255
      Index           =   14
      Left            =   720
      TabIndex        =   79
      Top             =   5040
      Width           =   2055
   End
   Begin Label Label8 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Hepatitis B                3 "
      Height          =   255
      Index           =   13
      Left            =   720
      TabIndex        =   78
      Top             =   4800
      Width           =   2055
   End
   Begin Label Label8 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Hepatitis B                2 "
      Height          =   255
      Index           =   12
      Left            =   720
      TabIndex        =   77
      Top             =   4560
      Width           =   2055
   End
   Begin Label Label8 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Hepatitis B                1 "
      Height          =   255
      Index           =   11
      Left            =   720
      TabIndex        =   76
      Top             =   4320
      Width           =   2055
   End
   Begin Label Label8 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "HIB                            "
      Height          =   255
      Index           =   10
      Left            =   720
      TabIndex        =   75
      Top             =   4080
      Width           =   2055
   End
   Begin Label Label8 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "MMR/MR                    "
      Height          =   255
      Index           =   9
      Left            =   720
      TabIndex        =   74
      Top             =   3840
      Width           =   2055
   End
   Begin Label Label8 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "OPV/IPV                   4"
      Height          =   255
      Index           =   8
      Left            =   720
      TabIndex        =   73
      Top             =   3600
      Width           =   2055
   End
   Begin Label Label8 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "OPV/IPV                   3"
      Height          =   255
      Index           =   7
      Left            =   720
      TabIndex        =   72
      Top             =   3360
      Width           =   2055
   End
   Begin Label Label8 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "OPV/IPV                   2"
      Height          =   255
      Index           =   6
      Left            =   720
      TabIndex        =   71
      Top             =   3120
      Width           =   2055
   End
   Begin Label Label8 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "OPV/IPV                   1"
      Height          =   255
      Index           =   5
      Left            =   720
      TabIndex        =   70
      Top             =   2880
      Width           =   2055
   End
   Begin Label Label8 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DTP/PED DT/ Td      5"
      Height          =   255
      Index           =   4
      Left            =   720
      TabIndex        =   69
      Top             =   2640
      Width           =   2055
   End
   Begin Label Label8 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DTP/PED DT/ Td      4"
      Height          =   255
      Index           =   3
      Left            =   720
      TabIndex        =   68
      Top             =   2400
      Width           =   2055
   End
   Begin Label Label8 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DTP/PED DT/ Td      3"
      Height          =   255
      Index           =   2
      Left            =   720
      TabIndex        =   67
      Top             =   2160
      Width           =   2055
   End
   Begin Label Label8 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DTP/PED DT/ Td      2"
      Height          =   255
      Index           =   1
      Left            =   720
      TabIndex        =   66
      Top             =   1920
      Width           =   2055
   End
   Begin Label Label8 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DTP/PED DT/ Td      1"
      Height          =   255
      Index           =   0
      Left            =   720
      TabIndex        =   65
      Top             =   1680
      Width           =   2055
   End
   Begin Line Line4 
      X1              =   2760
      X2              =   2760
      Y1              =   1440
      Y2              =   6960
   End
   Begin Line Line3 
      X1              =   3960
      X2              =   3960
      Y1              =   1440
      Y2              =   6960
   End
   Begin Line Line1 
      BorderWidth     =   2
      X1              =   720
      X2              =   6600
      Y1              =   1440
      Y2              =   1440
   End
   Begin Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0FFFF&
      Caption         =   "Doctor or Clinic"
      Height          =   195
      Left            =   5040
      TabIndex        =   64
      Top             =   1440
      Width           =   1335
   End
   Begin Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0FFFF&
      Caption         =   "Date Given"
      Height          =   195
      Left            =   2880
      TabIndex        =   63
      Top             =   1440
      Width           =   975
   End
   Begin Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0FFFF&
      Caption         =   "Vaccine"
      Height          =   195
      Left            =   960
      TabIndex        =   62
      Top             =   1440
      Width           =   705
   End
   Begin Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Drug Sensitivity:"
      Height          =   255
      Left            =   720
      TabIndex        =   61
      Top             =   1080
      Width           =   1455
   End
   Begin Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address:"
      Height          =   255
      Left            =   1320
      TabIndex        =   60
      Top             =   720
      Width           =   855
   End
   Begin Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date of Birth:"
      Height          =   255
      Left            =   960
      TabIndex        =   59
      Top             =   360
      Width           =   1215
   End
   Begin Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Full Name:"
      Height          =   255
      Left            =   1200
      TabIndex        =   34
      Top             =   0
      Width           =   975
   End
End
Dim Adding As Integer
Dim Closing As Integer
Option Explicit
'
' declare Windows API calls
'

Declare Function GetFreeSystemResources% Lib "User" (ByVal fuSysResources%)
Declare Function GetFreeSpace Lib "Kernel" (ByVal wFlags As Integer) As Long

Sub cmdAbout_Click ()
'
' calculate and display the #of bytes available in the
' global heap.
'
Dim dwFreeHeapSpace As Long
Dim Msg As String
Dim sysres
Dim gdires
Dim userres


dwFreeHeapSpace = GetFreeSpace(0)
sysres = GetFreeSystemResources(0)
gdires = GetFreeSystemResources(1)
userres = GetFreeSystemResources(2)
    
Msg = "SYSTEM resources  = " + sysres + "%" + Chr$(13)
Msg = Msg & "GDI resources          = " + gdires + "%" + Chr$(13)
Msg = Msg & "USER resources       = " + userres + "%" + Chr$(13)
Msg = Msg & "Global Heap Space available: " & Format$(dwFreeHeapSpace, "###,###,##0") & " bytes" + Chr$(13)
Msg = Msg + Chr$(13)
Msg = Msg & "Program written by Rodney Beede." + Chr$(13)
Msg = Msg & "Send questions and comments to rodney_beede@hotmail.com" & Chr$(13)
Msg = Msg & "Program published by Infinisoft." & Chr$(13)
Msg = Msg & Chr$(13)
Msg = Msg & "Support my work, register this program!" & Chr$(13)
Msg = Msg & "Read the file ImmunRec.txt for more info!"

MsgBox Msg, 64, "About Immunization Record"
End Sub

Sub CmdAdd_Click ()
   Dim i As Integer
   
   Adding = True
   data1.Recordset.AddNew
   cmdDelete.Enabled = False
   cmdNext.Enabled = False
   cmdBack.Enabled = False
   cmdSearch.Enabled = False
   cmdPrint.Enabled = False
   cmdAdd.Enabled = False
   cmdSave.Enabled = True
   Text1.Enabled = True
   MaskedEdit1.Enabled = True
   Text2.Enabled = True
   Text3.Enabled = True
   For i = 1 To 22
      MaskedEdit2(i - 1).Enabled = True
      Text4(i - 1).Enabled = True
   Next i
   Text1.SetFocus
End Sub

Sub cmdBack_Click ()
On Error Resume Next
data1.Recordset.MovePrevious
If data1.Recordset.BOF Then data1.Recordset.MoveNext
End Sub

Sub cmdClose_Click ()
   Unload Me
End Sub

Sub cmdDelete_Click ()
   Dim result As Integer
   On Error Resume Next
   Beep
   result = MsgBox("Are you sure you want to delete this record?", 36, "Confirm")

   If result = 6 Then data1.Recordset.Delete
   data1.Recordset.MoveNext
   If data1.Recordset.EOF Then data1.Recordset.MovePrevious
   If data1.Recordset.BOF And data1.Recordset.EOF Then Call data1_reposition
End Sub

Sub cmdHelp_Click ()
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
       CMD.HelpFile = VPath + "ImmunRec.HLP"
       CMD.HelpCommand = &H3
       CMD.Action = 6
    End If

End Sub

Sub cmdNext_Click ()
On Error Resume Next
data1.Recordset.MoveNext
If data1.Recordset.EOF Then data1.Recordset.MovePrevious
End Sub

Sub cmdPrint_Click ()
    If data1.Recordset.BOF And data1.Recordset.EOF Then
       
       MsgBox "Nothing to print.", 64, "Error"
    
       Exit Sub ' Leave sub
    End If

    frmPrint.Show ' Display the print form
    Me.Enabled = False ' Disable form
End Sub

Sub cmdSave_Click ()
      On Error Resume Next
      data1.Recordset.Update
      data1.Refresh
      cmdDelete.Enabled = True
      cmdNext.Enabled = True
      cmdBack.Enabled = True
      cmdSearch.Enabled = True
      cmdPrint.Enabled = True
      cmdAdd.Enabled = True
      cmdPrint.Enabled = True
      cmdSave.Enabled = False
      Adding = False
End Sub

Sub cmdSearch_Click ()
Me.Enabled = False
Load frmSearch
End Sub

Sub data1_reposition ()
Dim i As Integer

If data1.Recordset.BOF And data1.Recordset.EOF And Not Adding = True Then
   If Closing = False Then
      MsgBox "No move records in database.  Add one", 16, "Error"
   End If
   cmdDelete.Enabled = False
   cmdNext.Enabled = False
   cmdBack.Enabled = False
   cmdSearch.Enabled = False
   cmdPrint.Enabled = False
   Text1.Enabled = False
   MaskedEdit1.Enabled = False
   Text2.Enabled = False
   Text3.Enabled = False
   For i = 1 To 22
      MaskedEdit2(i - 1).Enabled = False
      Text4(i - 1).Enabled = False
   Next i
End If
End Sub

Sub Form_Load ()
    Dim i As Integer
    Me.Top = -60
    Me.Left = (screen.Width - Me.Width) / 2
    Me.Height = screen.Height + 60

    Adding = False
    data1.DatabaseName = App.Path + "\immunrec.mdb"
    data1.RecordSource = "ShotRecords"
    data1.Refresh
    
    Text1.DataField = "FullName"
    MaskedEdit1.DataField = "DateofBirth"
    Text2.DataField = "Address"
    Text3.DataField = "DrugSen"
    For i = 1 To 22
       MaskedEdit2(i - 1).DataField = "Date" + Chr$(64 + i)
       Text4(i - 1).DataField = "Doctor" + Chr$(64 + i)
    Next i

    Closing = True
    data1.Refresh
    Closing = False
End Sub

Sub Form_Unload (Cancel As Integer)
    Closing = True
    data1.Recordset.Update
    data1.Refresh
    data1.Recordset.Close
    Closing = False
End Sub

