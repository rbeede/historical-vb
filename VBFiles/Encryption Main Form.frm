VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Encryption Using the XOR Operator"
   ClientHeight    =   4335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4695
   Icon            =   "Encryption Main Form.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   4695
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtEncrypted 
      Height          =   1215
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   1920
      Width           =   4455
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear Everything"
      Height          =   495
      Left            =   3240
      TabIndex        =   6
      Top             =   3720
      Width           =   1335
   End
   Begin VB.CommandButton cmdDecrypt 
      Caption         =   "Decrypt"
      Height          =   495
      Left            =   1680
      TabIndex        =   5
      Top             =   3720
      Width           =   1335
   End
   Begin VB.CommandButton cmdEncrypt 
      Caption         =   "Encrypt"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   3720
      Width           =   1335
   End
   Begin VB.TextBox txtKey 
      Height          =   285
      Left            =   1800
      MaxLength       =   10
      TabIndex        =   3
      Top             =   3240
      Width           =   2775
   End
   Begin VB.TextBox txtDecrypted 
      Height          =   1215
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   360
      Width           =   4455
   End
   Begin VB.Label lblEncryptedText 
      AutoSize        =   -1  'True
      Caption         =   "Encrypted Text:"
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   1680
      Width           =   1125
   End
   Begin VB.Label lblKey 
      AutoSize        =   -1  'True
      Caption         =   "Type in a key to use:"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   3240
      Width           =   1485
   End
   Begin VB.Label lblDecryptedText 
      AutoSize        =   -1  'True
      Caption         =   "Decrypted Text:"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1140
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClear_Click()
    'Clear out text boxes
    txtDecrypted.Text = ""
    txtEncrypted.Text = ""
    txtKey.Text = ""
    
    'Set focus to decrypted text box
    txtDecrypted.SetFocus
End Sub

Private Sub cmdDecrypt_Click()
    Dim i As Long  'For counter
    Dim Key As Long  'Stores key
    Dim StartPlace As Integer  'For place when searching text
    Dim EndPlace As Integer  'For end place when searching text
    
    On Error Resume Next  'Incase bad key was used
    
    'Check to see if their is anything to Decrypt
    If txtEncrypted.Text = "" Then  'Nothing to Decrypt
        MsgBox "There is nothing to Decrypt!", vbExclamation, "Error"
        Exit Sub  'Leave sub
    End If
    
    'Check to see if a key was entered
    If txtKey.Text = "" Then
        'No key entered, tell user one will be given
        MsgBox "You didn't enter a Decryption key, one will be made for you.", vbInformation, "Decrypt"
        txtKey.Text = Str$(CInt(Rnd * 10))  'Put a random number for key
    End If
    
    'The key must be a number
    If Not IsNumeric(txtKey.Text) Then
        'Tell user key is not valid
        MsgBox "You have entered a invalid key, the key can only be a number.", vbExclamation, "Error"
        Exit Sub  'Leave sub
    End If

    'Check to see if key is to big or to small
    If (Val(txtKey.Text) < 0) Or (Val(txtKey.Text) > 2147483647) Then
        'Tell user that key will not work
        MsgBox "Your key is either to small or to big.  Valid ranges are 0 to 2,147,483,647.  You cannot use commas.", vbExclamation, "Error"
        Exit Sub  'Leave sub
    Else  'Key is in valid range
        Key = Val(txtKey.Text)  'Store key
    End If

    'Decrypt each character and put it in the Decrypted text box
    txtDecrypted.Text = ""  'Clear out text box
    
    StartPlace = 1  'Start at beginning
    EndPlace = Len(txtEncrypted.Text)  'Set end place to look at
    
    'Have to loop through pulling out each number
    Do
        'Get the beginning of next number
        For i = StartPlace To EndPlace
            If Mid$(txtEncrypted.Text, i, 1) = " " Then  'Found beginning of number
                StartPlace = i + 1  'Mark place
                Exit For  'Leave loop
            End If
        Next i
        
        'Get the end of the next number
        For i = StartPlace To EndPlace
            If Mid$(txtEncrypted.Text, i, 1) = " " Then  'Found end of number
                EndPlace = i  'Mark Place
                Exit For  'Leave loop
            End If
        Next i

        'Decrypt number back to character and display
        txtDecrypted.Text = txtDecrypted.Text + Chr$((Val(Mid$(txtEncrypted.Text, StartPlace, EndPlace - StartPlace + 1)) Xor Key))
        
        StartPlace = EndPlace  'Reset new starting place for next number
        EndPlace = Len(txtEncrypted.Text)  'Reset ending place
    Loop Until StartPlace >= Len(txtEncrypted.Text)
    
    txtEncrypted.Text = ""  'Remove Encrypted text
End Sub

Private Sub cmdEncrypt_Click()
    Dim i As Long  'For counter
    Dim Key As Long  'Stores key
    
    On Error Resume Next  'Incase of bad key
    
    'Check to see if their is anything to encrypt
    If txtDecrypted.Text = "" Then  'Nothing to encrypt
        MsgBox "There is nothing to encrypt!", vbExclamation, "Error"
        Exit Sub  'Leave sub
    End If
    
    'Check to see if a key was entered
    If txtKey.Text = "" Then
        'No key entered, tell user one will be given
        MsgBox "You didn't enter a encryption key, one will be made for you.", vbInformation, "Encrypt"
        txtKey.Text = Str$(CInt(Rnd * 10))  'Put a random number for key
    End If
    
    'The key must be a number
    If Not IsNumeric(txtKey.Text) Then
        'Tell user key is not valid
        MsgBox "You have entered a invalid key, the key can only be a number.", vbExclamation, "Error"
        Exit Sub  'Leave sub
    End If

    'Check to see if key is to big or to small
    If (Val(txtKey.Text) < 0) Or (Val(txtKey.Text) > 2147483647) Then
        'Tell user that key will not work
        MsgBox "Your key is either to small or to big.  Valid ranges are 0 to 2,147,483,647.  You cannot use commas.", vbExclamation, "Error"
        Exit Sub  'Leave sub
    Else  'Key is in valid range
        Key = Val(txtKey.Text)  'Store key
    End If

    'Encrypt each character and put it in the encrypted text box
    txtEncrypted.Text = ""  'Clear out text box
    For i = 1 To Len(txtDecrypted.Text)
        txtEncrypted.Text = txtEncrypted.Text + Str$(Asc(Mid$(txtDecrypted.Text, i, 1)) Xor Key)
    Next i

    txtDecrypted.Text = ""  'Remove decrypted text
End Sub
