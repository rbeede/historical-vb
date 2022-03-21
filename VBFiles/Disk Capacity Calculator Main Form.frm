VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Disk Capacity Calculator"
   ClientHeight    =   2790
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4710
   Icon            =   "Disk Capacity Calculator Main Form.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   4710
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtResults 
      Height          =   1215
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Text            =   "Disk Capacity Calculator Main Form.frx":000C
      Top             =   240
      Width           =   4215
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   495
      Left            =   2393
      TabIndex        =   3
      Top             =   2040
      Width           =   2055
   End
   Begin VB.CommandButton cmdCalculate 
      Caption         =   "&Calculate"
      Default         =   -1  'True
      Height          =   495
      Left            =   233
      TabIndex        =   2
      Top             =   2040
      Width           =   2055
   End
   Begin VB.TextBox txtDriveSize 
      Height          =   285
      Left            =   1080
      TabIndex        =   0
      Top             =   1560
      Width           =   3375
   End
   Begin VB.Label lblDriveSize 
      AutoSize        =   -1  'True
      Caption         =   "Drive Size"
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   1560
      Width           =   720
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCalculate_Click()
    'Set the size in decimal and binary of different bytes
    Const BinDecByteSize As Byte = 1  'Decimal and Binary byte size
    
    Const BinKByteSize As Integer = 1024  'Binary Kilobyte size
    Const DecKByteSize As Integer = 1000  'Decimal Kilobyte size
    
    Const BinMByteSize As Long = 1048576  'Binary Megabyte size
    Const DecMByteSize As Long = 1000000  'Decimal Megabyte size
    
    Const BinGByteSize As Long = 1073741824  'Binary Gigabyte size
    Const DecGByteSize As Long = 1000000000  'Decimal Gigabyte size
    
    Const BinTByteSize As Double = 1099511627776#  'Binary Terabyte size
    Const DecTByteSize As Double = 1000000000000#  'Decimal Terabyte size
    
    Dim DecSize As Variant  'To hold decimal size of calculation (base 10)
    Dim BinSize As Variant  'To hold binary size of calculation (base 2)
    Dim DriveSize As Double  'Holds drive size
    Dim ByteType As Integer  'Holds information about the size entered (MB, GB, etc.)
    Dim i As Integer  'For counter
    
    'Check to see if the data was entered correctly
    If InStr(txtDriveSize.Text, " ") <> 0 And InStr(txtDriveSize.Text, "B") Then
        'Correct format up to now, try to read in drive size
        For i = 1 To Len(txtDriveSize.Text)
            If Mid$(txtDriveSize.Text, i, 1) = " " Then
                'Found space separating number from number size
                'Read in number
                DriveSize = Val(Left$(txtDriveSize.Text, i - 1))
    
                'Determine what the byte type is
                'Have to start at top since looking for "B" for byte would
                'automatically kick us out
                If InStr(txtDriveSize.Text, "TB") <> 0 Then  'Terabyte
                    ByteType = 4  'Terabyte type
                ElseIf InStr(txtDriveSize.Text, "GB") <> 0 Then  'Gigabyte
                    ByteType = 3  'Gigabyte type
                ElseIf InStr(txtDriveSize.Text, "MB") <> 0 Then  'Megabyte
                    ByteType = 2  'Megabyte type
                ElseIf InStr(txtDriveSize.Text, "KB") <> 0 Then  'Kilobyte
                    ByteType = 1  'Kilobyte type
                ElseIf InStr(txtDriveSize.Text, "B") <> 0 Then  'Byte
                    ByteType = 0  'Byte type
                End If
            End If
        Next i
    Else  'No correct data was entered
        'Tell user
        MsgBox "No valid data was entered", vbCritical + vbOKOnly, "Error"
        txtDriveSize.SetFocus  'Give focus to text box
        Exit Sub  'Leave sub
    End If

    'Give header information
    txtResults.Text = "Results for a drive the size of " & DriveSize
    
    'Calculate the size in every possible byte size
    Select Case ByteType
        Case 0  'Byte type
            'Fill in the rest of the header informatin
            txtResults.Text = txtResults.Text + " Bytes:" + vbCrLf
            
            'Show size of drive in binary & decimal Byte size
            txtResults.Text = txtResults.Text + "Binary size in Bytes is " & (DriveSize) & vbCrLf
            txtResults.Text = txtResults.Text + "Decimal size in Bytes is " & (DriveSize) & vbCrLf
            
            'Show size of drive in binary & decimal KiloByte size
            txtResults.Text = txtResults.Text + "Binary size in KiloBytes is " & (DriveSize / BinKByteSize) & vbCrLf
            txtResults.Text = txtResults.Text + "Decimal size in KiloBytes is " & (DriveSize / DecKByteSize) & vbCrLf
    
            'Show size of drive in binary & decimal MegaByte size
            txtResults.Text = txtResults.Text + "Binary size in MegaBytes is " & (DriveSize / BinMByteSize) & vbCrLf
            txtResults.Text = txtResults.Text + "Decimal size in MegaBytes is " & (DriveSize / DecMByteSize) & vbCrLf
    
            'Show size of drive in binary & decimal GigaByte size
            txtResults.Text = txtResults.Text + "Binary size in GigaBytes is " & (DriveSize / BinGByteSize) & vbCrLf
            txtResults.Text = txtResults.Text + "Decimal size in GigaBytes is " & (DriveSize / DecGByteSize) & vbCrLf
    
            'Show size of drive in binary & decimal TeraByte size
            txtResults.Text = txtResults.Text + "Binary size in TeraBytes is " & (DriveSize / BinTByteSize) & vbCrLf
            txtResults.Text = txtResults.Text + "Decimal size in TeraBytes is " & (DriveSize / DecTByteSize) & vbCrLf
        Case 1  'KiloBytes
            'Fill in the rest of the header informatin
            txtResults.Text = txtResults.Text + " KiloBytes:" + vbCrLf
            
            'Show size of drive in binary & decimal Byte size
            txtResults.Text = txtResults.Text + "Binary size in Bytes is " & (DriveSize * BinKByteSize / BinDecByteSize) & vbCrLf
            txtResults.Text = txtResults.Text + "Decimal size in Bytes is " & (DriveSize * DecKByteSize / BinDecByteSize) & vbCrLf
            
            'Show size of drive in binary & decimal KiloByte size
            txtResults.Text = txtResults.Text + "Binary size in KiloBytes is " & (DriveSize) & vbCrLf
            txtResults.Text = txtResults.Text + "Decimal size in KiloBytes is " & (DriveSize) & vbCrLf
    
            'Show size of drive in binary & decimal MegaByte size
            txtResults.Text = txtResults.Text + "Binary size in MegaBytes is " & (DriveSize * BinKByteSize / BinMByteSize) & vbCrLf
            txtResults.Text = txtResults.Text + "Decimal size in MegaBytes is " & (DriveSize * DecKByteSize / DecMByteSize) & vbCrLf
    
            'Show size of drive in binary & decimal GigaByte size
            txtResults.Text = txtResults.Text + "Binary size in GigaBytes is " & (DriveSize * BinKByteSize / BinGByteSize) & vbCrLf
            txtResults.Text = txtResults.Text + "Decimal size in GigaBytes is " & (DriveSize * DecKByteSize / DecGByteSize) & vbCrLf
    
            'Show size of drive in binary & decimal TeraByte size
            txtResults.Text = txtResults.Text + "Binary size in TeraBytes is " & (DriveSize * BinKByteSize / BinTByteSize) & vbCrLf
            txtResults.Text = txtResults.Text + "Decimal size in TeraBytes is " & (DriveSize * DecKByteSize / DecTByteSize) & vbCrLf
        Case 2  'MegaBytes
            'Fill in the rest of the header informatin
            txtResults.Text = txtResults.Text + " MegaBytes:" + vbCrLf
            
            'Show size of drive in binary & decimal Byte size
            txtResults.Text = txtResults.Text + "Binary size in Bytes is " & (DriveSize * BinMByteSize / BinDecByteSize) & vbCrLf
            txtResults.Text = txtResults.Text + "Decimal size in Bytes is " & (DriveSize * DecMByteSize / BinDecByteSize) & vbCrLf
            
            'Show size of drive in binary & decimal KiloByte size
            txtResults.Text = txtResults.Text + "Binary size in KiloBytes is " & (DriveSize * BinMByteSize / BinKByteSize) & vbCrLf
            txtResults.Text = txtResults.Text + "Decimal size in KiloBytes is " & (DriveSize * DecMByteSize / DecKByteSize) & vbCrLf
    
            'Show size of drive in binary & decimal MegaByte size
            txtResults.Text = txtResults.Text + "Binary size in MegaBytes is " & (DriveSize * BinMByteSize / BinMByteSize) & vbCrLf
            txtResults.Text = txtResults.Text + "Decimal size in MegaBytes is " & (DriveSize * DecMByteSize / DecMByteSize) & vbCrLf
    
            'Show size of drive in binary & decimal GigaByte size
            txtResults.Text = txtResults.Text + "Binary size in GigaBytes is " & (DriveSize * BinMByteSize / BinGByteSize) & vbCrLf
            txtResults.Text = txtResults.Text + "Decimal size in GigaBytes is " & (DriveSize * DecMByteSize / DecGByteSize) & vbCrLf
    
            'Show size of drive in binary & decimal TeraByte size
            txtResults.Text = txtResults.Text + "Binary size in TeraBytes is " & (DriveSize * BinMByteSize / BinTByteSize) & vbCrLf
            txtResults.Text = txtResults.Text + "Decimal size in TeraBytes is " & (DriveSize * DecMByteSize / DecTByteSize) & vbCrLf
        Case 3  'GigaBytes
            'Fill in the rest of the header informatin
            txtResults.Text = txtResults.Text + " GigaBytes:" + vbCrLf
            
            'Show size of drive in binary & decimal Byte size
            txtResults.Text = txtResults.Text + "Binary size in Bytes is " & (DriveSize * BinGByteSize / BinDecByteSize) & vbCrLf
            txtResults.Text = txtResults.Text + "Decimal size in Bytes is " & (DriveSize * DecGByteSize / BinDecByteSize) & vbCrLf
            
            'Show size of drive in binary & decimal KiloByte size
            txtResults.Text = txtResults.Text + "Binary size in KiloBytes is " & (DriveSize * BinGByteSize / BinKByteSize) & vbCrLf
            txtResults.Text = txtResults.Text + "Decimal size in KiloBytes is " & (DriveSize * DecGByteSize / DecKByteSize) & vbCrLf
    
            'Show size of drive in binary & decimal MegaByte size
            txtResults.Text = txtResults.Text + "Binary size in MegaBytes is " & (DriveSize * BinGByteSize / BinMByteSize) & vbCrLf
            txtResults.Text = txtResults.Text + "Decimal size in MegaBytes is " & (DriveSize * DecGByteSize / DecMByteSize) & vbCrLf
    
            'Show size of drive in binary & decimal GigaByte size
            txtResults.Text = txtResults.Text + "Binary size in GigaBytes is " & (DriveSize * BinGByteSize / BinGByteSize) & vbCrLf
            txtResults.Text = txtResults.Text + "Decimal size in GigaBytes is " & (DriveSize * DecGByteSize / DecGByteSize) & vbCrLf
    
            'Show size of drive in binary & decimal TeraByte size
            txtResults.Text = txtResults.Text + "Binary size in TeraBytes is " & (DriveSize * BinGByteSize / BinTByteSize) & vbCrLf
            txtResults.Text = txtResults.Text + "Decimal size in TeraBytes is " & (DriveSize * DecGByteSize / DecTByteSize) & vbCrLf
        Case 4  'TeraBytes
            'Fill in the rest of the header informatin
            txtResults.Text = txtResults.Text + " TeraBytes:" + vbCrLf
            
            'Show size of drive in binary & decimal Byte size
            txtResults.Text = txtResults.Text + "Binary size in Bytes is " & (DriveSize * BinTByteSize / BinDecByteSize) & vbCrLf
            txtResults.Text = txtResults.Text + "Decimal size in Bytes is " & (DriveSize * DecTByteSize / BinDecByteSize) & vbCrLf
            
            'Show size of drive in binary & decimal KiloByte size
            txtResults.Text = txtResults.Text + "Binary size in KiloBytes is " & (DriveSize * BinTByteSize / BinKByteSize) & vbCrLf
            txtResults.Text = txtResults.Text + "Decimal size in KiloBytes is " & (DriveSize * DecTByteSize / DecKByteSize) & vbCrLf
    
            'Show size of drive in binary & decimal MegaByte size
            txtResults.Text = txtResults.Text + "Binary size in MegaBytes is " & (DriveSize * BinTByteSize / BinMByteSize) & vbCrLf
            txtResults.Text = txtResults.Text + "Decimal size in MegaBytes is " & (DriveSize * DecTByteSize / DecMByteSize) & vbCrLf
    
            'Show size of drive in binary & decimal GigaByte size
            txtResults.Text = txtResults.Text + "Binary size in GigaBytes is " & (DriveSize * BinTByteSize / BinGByteSize) & vbCrLf
            txtResults.Text = txtResults.Text + "Decimal size in GigaBytes is " & (DriveSize * DecTByteSize / DecGByteSize) & vbCrLf
    
            'Show size of drive in binary & decimal TeraByte size
            txtResults.Text = txtResults.Text + "Binary size in TeraBytes is " & (DriveSize * BinTByteSize / BinTByteSize) & vbCrLf
            txtResults.Text = txtResults.Text + "Decimal size in TeraBytes is " & (DriveSize * DecTByteSize / DecTByteSize) & vbCrLf
    End Select
End Sub

Private Sub cmdExit_Click()
    End  'Terminate program
End Sub
