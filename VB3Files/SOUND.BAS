' Forces all variables to be declared
Option Explicit

'Declare to get windows directory in variable
Global WindowsDir As String

'Declare for Path and File for .Wav
Global PathWavFile As String

'Declare wav file to show on sound form
Global WavFile As String

'Declares for .WAV Sound Player
Declare Function sndPlaySound Lib "MMSystem" (lpsound As Any, ByVal flag As Integer) As Integer

'Declares for WindowsDirectory
Declare Function GetWindowsDirectory Lib "Kernel" (ByVal P$, ByVal S%) As Integer


Sub PlaySound (DirWavFile As String)

       ' Declare variabled needed
       Dim R As Integer, sPath As String
       Const SYNC = 1
       
       On Error Resume Next ' If a error occurs go on
       
       ' Play sound
       R = sndPlaySound(ByVal CStr(DirWavFile), SYNC)

       If Err Then ' Error occured
          MsgBox "Error playing sound.", 16, "Error"
       End If
End Sub

Function WindowsDirectory () As String
Dim WinPath As String
    WinPath = String(145, Chr(0))
    WindowsDirectory = Left(WinPath, GetWindowsDirectory(WinPath, Len(WinPath)))
End Function

