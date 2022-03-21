'----------------------------------------------------------------
'     Frame animation routines for Visual Basic applications
'----------------------------------------------------------------
' Source file:  PAN02.BAS
' Release version:  1.00                   Programmer:  Lee Adams
' Type:  Visual Basic functional module for Windows applications
' Compiler:  Microsoft Visual Basic 1.00
' Output and features:  Provides routines to manage the authoring
'   process and playback engine for real-time animation.
' Publication:  Contains material from Windcrest/McGraw-Hill
'   book 4224 published by TAB BOOKS Div. of McGraw-Hill Inc.
' License:  As purchaser of the book you are granted a
'   royalty-free license to distribute executable files
'   generated uSing this code provided that you accept the
'   conditions of the License Agreement and Limited Warranty
'   described in the book and on the companion disk.  Government
'   users:  This software and documentation are subject to
'   restrictions set forth in The Rights in Technical Data and
'   Computer Software clause at 252.227-7013 and elsewhere.
'----------------------------------------------------------------
'     (c) Copyright 1992 Lee Adams.  All rights reserved.
'          Lee Adams(tm) is a trademark of Lee Adams.
'----------------------------------------------------------------
'
'

'----------------------------------------------------------------
'                    Display the next frame.
'  This procedure is intelligent enough to discern between
'  RAM-BASED FRAME ANIMATION (where all frames have already been
'  loaded from disk and stored in RAM) and DISK-BASED FRAME
'  ANIMATION (where each frame must be loaded from disk during
'  playback).
'----------------------------------------------------------------
Sub zShowNextFrame ()              'uses global variable FrameNum
If UseDisk = True GoTo DISK_PLAYBACK

RAM_PLAYBACK:         'if all frames have been loaded into RAM...
If AnimationReady = False Then
  Exit Sub
End If
If AnimationLoaded = False Then
  Exit Sub
End If
If Pause = True Then
  Exit Sub
End If
If FrameDirection = zFORWARD Then
  FrameNum = FrameNum + 1             'increment the frame number
  If FrameNum > zFINALFRAME Then
    FrameNum = zFIRSTFRAME
  End If
End If
If FrameDirection = zREVERSE Then
  FrameNum = FrameNum - 1             'decrement the frame number
  If FrameNum < zFIRSTFRAME Then
    FrameNum = zFINALFRAME
  End If
End If
Select Case FrameNum                'select the appropriate frame
  Case 1:  RetVal = SelectObject(hFDC, hF1)
  Case 2:  RetVal = SelectObject(hFDC, hF2)
  Case 3:  RetVal = SelectObject(hFDC, hF3)
  Case 4:  RetVal = SelectObject(hFDC, hF4)
  Case 5:  RetVal = SelectObject(hFDC, hF5)
  Case 6:  RetVal = SelectObject(hFDC, hF6)
  Case 7:  RetVal = SelectObject(hFDC, hF7)
  Case 8:  RetVal = SelectObject(hFDC, hF8)
  Case 9:  RetVal = SelectObject(hFDC, hF9)
  Case 10:  RetVal = SelectObject(hFDC, hF10)
  Case 11:  RetVal = SelectObject(hFDC, hF11)
  Case 12:  RetVal = SelectObject(hFDC, hF12)
  Case 13:  RetVal = SelectObject(hFDC, hF13)
  Case 14:  RetVal = SelectObject(hFDC, hF14)
  Case 15:  RetVal = SelectObject(hFDC, hF15)
  Case 16:  RetVal = SelectObject(hFDC, hF16)
  Case 17:  RetVal = SelectObject(hFDC, hF17)
  Case 18:  RetVal = SelectObject(hFDC, hF18)
  Case 19:  RetVal = SelectObject(hFDC, hF19)
  Case 20:  RetVal = SelectObject(hFDC, hF18)
  Case 21:  RetVal = SelectObject(hFDC, hF17)
  Case 22:  RetVal = SelectObject(hFDC, hF16)
  Case 23:  RetVal = SelectObject(hFDC, hF15)
  Case 24:  RetVal = SelectObject(hFDC, hF14)
  Case 25:  RetVal = SelectObject(hFDC, hF13)
  Case 26:  RetVal = SelectObject(hFDC, hF12)
  Case 27:  RetVal = SelectObject(hFDC, hF11)
  Case 28:  RetVal = SelectObject(hFDC, hF10)
  Case 29:  RetVal = SelectObject(hFDC, hF9)
  Case 30:  RetVal = SelectObject(hFDC, hF8)
  Case 31:  RetVal = SelectObject(hFDC, hF7)
  Case 32:  RetVal = SelectObject(hFDC, hF6)
  Case 33:  RetVal = SelectObject(hFDC, hF5)
  Case 34:  RetVal = SelectObject(hFDC, hF4)
  Case 35:  RetVal = SelectObject(hFDC, hF3)
  Case 36:  RetVal = SelectObject(hFDC, hF2)
End Select
RetVal = BitBlt(Form1.hDC, 0, 0, zFRAMEWIDE, zFRAMEHIGH, hFDC, 0, 0, SRCCOPY)
Exit Sub

DISK_PLAYBACK: 'if each frame must be loaded from disk as needed
If AnimationHalted = True Then
  Exit Sub
End If
If Pause = True Then
  Exit Sub
End If
If FrameDirection = zFORWARD Then
  FrameNum = FrameNum + 1             'increment the frame number
  If FrameNum > zFINALFRAME Then
    FrameNum = zFIRSTFRAME
  End If
End If
If FrameDirection = zREVERSE Then
  FrameNum = FrameNum - 1             'decrement the frame number
  If FrameNum < zFIRSTFRAME Then
    FrameNum = zFINALFRAME
  End If
End If
Select Case FrameNum
  Case 1:  FileName = "PAN1.BMP": zLoadFrame
  Case 2:  FileName = "PAN2.BMP": zLoadFrame
  Case 3:  FileName = "PAN3.BMP": zLoadFrame
  Case 4:  FileName = "PAN4.BMP": zLoadFrame
  Case 5:  FileName = "PAN5.BMP": zLoadFrame
  Case 6:  FileName = "PAN6.BMP": zLoadFrame
  Case 7:  FileName = "PAN7.BMP": zLoadFrame
  Case 8:  FileName = "PAN8.BMP": zLoadFrame
  Case 9:  FileName = "PAN9.BMP": zLoadFrame
  Case 10:  FileName = "PAN10.BMP": zLoadFrame
  Case 11:  FileName = "PAN11.BMP": zLoadFrame
  Case 12:  FileName = "PAN12.BMP": zLoadFrame
  Case 13:  FileName = "PAN13.BMP": zLoadFrame
  Case 14:  FileName = "PAN14.BMP": zLoadFrame
  Case 15:  FileName = "PAN15.BMP": zLoadFrame
  Case 16:  FileName = "PAN16.BMP": zLoadFrame
  Case 17:  FileName = "PAN17.BMP": zLoadFrame
  Case 18:  FileName = "PAN18.BMP": zLoadFrame
  Case 19:  FileName = "PAN19.BMP": zLoadFrame
  Case 20:  FileName = "PAN18.BMP": zLoadFrame
  Case 21:  FileName = "PAN17.BMP": zLoadFrame
  Case 22:  FileName = "PAN16.BMP": zLoadFrame
  Case 23:  FileName = "PAN15.BMP": zLoadFrame
  Case 24:  FileName = "PAN14.BMP": zLoadFrame
  Case 25:  FileName = "PAN13.BMP": zLoadFrame
  Case 26:  FileName = "PAN12.BMP": zLoadFrame
  Case 27:  FileName = "PAN11.BMP": zLoadFrame
  Case 28:  FileName = "PAN10.BMP": zLoadFrame
  Case 29:  FileName = "PAN9.BMP": zLoadFrame
  Case 30:  FileName = "PAN8.BMP": zLoadFrame
  Case 31:  FileName = "PAN7.BMP": zLoadFrame
  Case 32:  FileName = "PAN6.BMP": zLoadFrame
  Case 33:  FileName = "PAN5.BMP": zLoadFrame
  Case 34:  FileName = "PAN4.BMP": zLoadFrame
  Case 35:  FileName = "PAN3.BMP": zLoadFrame
  Case 36:  FileName = "PAN2.BMP": zLoadFrame
End Select
If FrameLoaded = False Then
  AnimationHalted = True
  Beep
  MsgBox "Unable to load next frame from disk.  Animation halted.", MB_OK, "Unexpected animation condition"
  Exit Sub
End If
End Sub

'----------------------------------------------------------------
'                      Load a frame from disk
'----------------------------------------------------------------
Sub zLoadFrame ()
If FrameReady = False Then
  MsgBox "Hidden frame not ready.", MB_OK, "Unexpected animation condition"
  Exit Sub
End If
LoadingFrame = True                    'disable refresh procedure
On Error GoTo LoadError:                   'enable error trapping
Form1.Picture = LoadPicture(FileName)            'load .BMP image
On Error GoTo 0                           'disable error trapping
LoadingFrame = False                    'enable refresh procedure
FrameLoaded = True
'-------- select persistent bitmap, clear it, deselect it -------
Form1.AutoRedraw = True
RetVal = PatBlt(Form1.hDC, 0, 0, zFRAMEWIDE, zFRAMEHIGH, WHITENESS)
Form1.AutoRedraw = False
Exit Sub
'-------------------------- error-handler -----------------------
LoadError:
  On Error GoTo 0
  LoadingFrame = False
  FrameLoaded = False
  Beep
  MsgBox "Unable to load the .BMP file.  Is system initialized?  Does file exist on disk?", MB_OK, "Unexpected animation condition"
  Exit Sub
End Sub

