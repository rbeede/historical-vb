VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1800
      TabIndex        =   0
      Top             =   1320
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Dim i As Integer  'For counter
    Dim FilePart() As String  'Holds parts of file
    Dim NumParts As Integer  'Number of parts of file
    Dim FileSize As Long  'Size of the file
    Dim WholeFile As String  'Stores entire file
    Dim WholeFile2nd As String  'Stores second copy of file
    
    'Open file for reading
    Open "C:\WINDOWS\PBRUSH.EXE" For Binary As #1
    
        FileSize = LOF(1)  'Get file size
    
        'Size WholeFile string variable to the file size
        WholeFile = String(FileSize, " ")
    
        Get #1, , WholeFile  'Read entire file into memory
    
    Close #1  'Close file
    
    
    NumParts = (FileSize / 5000)  'Get number of parts in 5KB blocks

    ReDim FilePart(NumParts + 1)  'Dimension FilePart array
    
    'Split up file into 5KB blocks
    For i = 0 To NumParts
        FilePart(i) = Mid(WholeFile, i * 5000& + 1, 5000)
    Next i

    
    'Open up new file for writing
    Open "C:\TEMP\TEST.EXE" For Binary As #1
    
        'Put all the pieces together in one string
        For i = 0 To NumParts
            WholeFile2nd = WholeFile2nd + FilePart(i)
        Next i
        
        'Write entire file all at once
        Put #1, 1, WholeFile2nd
    
    Close #1  'Close file

    DoEvents  'Allow Windows to process
    
    'Print out debugging info
    Print "File size should be " & FileSize
    Print "New file size is " & FileLen("C:\TEMP\TEST.EXE")
End Sub
