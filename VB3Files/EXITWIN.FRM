VERSION 2.00
Begin Form Form1 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "System"
   ClientHeight    =   4335
   ClientLeft      =   2205
   ClientTop       =   1875
   ClientWidth     =   2295
   ClipControls    =   0   'False
   DrawMode        =   11  'Nop
   Height          =   4740
   Icon            =   EXITWIN.FRX:0000
   Left            =   2145
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   2295
   Top             =   1530
   Width           =   2415
   Begin SSPanel pn3GDI 
      BackColor       =   &H00C0C0C0&
      Caption         =   "GDI"
      FloodType       =   1  'Left To Right
      Font3D          =   0  'None
      Height          =   375
      Left            =   0
      Outline         =   -1  'True
      TabIndex        =   8
      Top             =   3000
      Width           =   2295
   End
   Begin SSPanel pn3System 
      BackColor       =   &H00C0C0C0&
      Caption         =   "System"
      FloodType       =   1  'Left To Right
      Font3D          =   0  'None
      Height          =   375
      Left            =   0
      Outline         =   -1  'True
      TabIndex        =   7
      Top             =   3480
      Width           =   2295
   End
   Begin SSPanel pn3User 
      BackColor       =   &H00C0C0C0&
      Caption         =   " "
      FloodType       =   1  'Left To Right
      Font3D          =   0  'None
      Height          =   375
      Left            =   0
      Outline         =   -1  'True
      TabIndex        =   6
      Top             =   3960
      Width           =   2295
   End
   Begin CommandButton cmdGlobalHeapSpace 
      Caption         =   "Global &Heap Space"
      Height          =   495
      Left            =   0
      TabIndex        =   5
      Top             =   1920
      Width           =   2295
   End
   Begin CommandButton Command5 
      Caption         =   "&Quit Windows"
      Height          =   495
      Left            =   0
      TabIndex        =   3
      Top             =   960
      Width           =   2295
   End
   Begin CommandButton Command4 
      Caption         =   "System &Info"
      Height          =   495
      Left            =   0
      TabIndex        =   4
      Top             =   1440
      Width           =   2295
   End
   Begin CommandButton Command3 
      Caption         =   "Re&Boot System"
      Height          =   495
      Left            =   0
      TabIndex        =   2
      Top             =   480
      Width           =   2295
   End
   Begin CommandButton Command2 
      Caption         =   "&Restart Windows"
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   2295
   End
   Begin CommandButton Command1 
      Caption         =   "&Done"
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   2400
      Width           =   2295
   End
End
Option Explicit
'
' declare Windows API calls
'
Declare Function ExitWindows% Lib "User" (ByVal dwReturnCode&, ByVal reserved%)
Declare Function GetFreeSystemResources% Lib "User" (ByVal fuSysResources%)
Declare Function GetFreeSpace Lib "Kernel" (ByVal wFlags As Integer) As Long

Sub cmdGlobalHeapSpace_Click ()
'
' calculate and display the #of bytes available in the
' global heap.
'
Dim dwFreeHeapSpace As Long
Dim msg As String

dwFreeHeapSpace = GetFreeSpace(0)
    
msg = "Global Heap Space available: " & Format$(dwFreeHeapSpace, "###,###,##0") & " bytes"

MsgBox msg, 0, "Heap Space"

End Sub

Sub Command1_Click ()
    End
End Sub

Sub Command2_Click ()
Dim X As Integer
    X = ExitWindows(66, 0)
End Sub

Sub Command3_Click ()
Dim X As Integer
    X = ExitWindows(67, 0)
End Sub

Sub Command4_Click ()
'
' display Windows system resources available
'
Dim sysres
Dim gdires
Dim userres
Dim msg As String

sysres = GetFreeSystemResources(0)
gdires = GetFreeSystemResources(1)
userres = GetFreeSystemResources(2)
    
msg = "SYSTEM resources  = " + sysres + "%" + Chr$(13)
msg = msg & "GDI resources          = " + gdires + "%" + Chr$(13)
msg = msg & "USER resources       = " + userres + "%"
'
' display graphs
'
pn3System.FloodPercent = sysres
pn3GDI.FloodPercent = gdires
pn3User.FloodPercent = userres

MsgBox msg, 0, "System Info"

End Sub

Sub Command5_Click ()
Dim X As Integer
    X = ExitWindows(0, 0)
End Sub

