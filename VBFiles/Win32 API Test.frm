VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long

Private Sub Form_Click()
    Dim Used() As Boolean
    Dim ReturnColor As Long
    Dim x As Integer
    Dim y As Integer
    Dim c1 As Long
    Dim c2 As Long
    Dim c3 As Long
    Dim Start As Single
    
    Me.ScaleMode = vbPixels
    
    ReDim Used(Screen.Width, Screen.Height)

    Start = -1
    
    Do
retry:
        x = Rnd * Me.ScaleWidth
        y = Rnd * Me.ScaleHeight
        
        If (x < 1) Or (x > Me.ScaleWidth) Then GoTo retry
        If (y < 1) Or (y > Me.ScaleHeight) Then GoTo retry
        
        If Used(x, y) = True Then
            Randomize
            GoTo retry
        End If
        
        Used(x, y) = True
        
        c1 = Rnd * 255 + 1
        c2 = Rnd * 255 + 1
        c3 = Rnd * 255 + 1
        
        ReturnColor = SetPixel(Form1.hdc, x, y, RGB(c1, c2, c3))
    
        If ReturnColor = -1 Then Stop
        
        If Start = -1 Then Start = Timer
        
        If Timer - Start > 1 Then
            DoEvents
            Form1.Refresh
            DoEvents
            Start = -1
            c2 = Rnd * 255 + 1
        End If
    
        DoEvents
    Loop
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub
