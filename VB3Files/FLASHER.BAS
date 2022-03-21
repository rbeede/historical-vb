Declare Function FlashWindow Lib "User" (ByVal hWnd As Integer, ByVal bInvert As Integer) As Integer
Type PointAPI
     Y As Integer
     X As Integer
End Type
Global MP As PointAPI
Declare Sub GetCursorPos Lib "User" (P As PointAPI)
Declare Function WindowFromPoint Lib "User" (ByVal X As Integer, ByVal Y As Integer) As Integer

