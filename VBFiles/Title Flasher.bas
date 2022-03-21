Attribute VB_Name = "FLASHER1"
Declare Function FlashWindow Lib "user32" (ByVal hwnd As Long, ByVal bInvert As Long) As Long

Type PointAPI
     X As Integer
     Y As Integer
End Type
Global MP As PointAPI

Declare Function GetCursorPos Lib "user32" (lpPoint As PointAPI) As Long
Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long

