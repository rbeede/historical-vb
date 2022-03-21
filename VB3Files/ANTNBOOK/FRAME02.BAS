'----------------------------------------------------------------
'     Frame animation routines for Visual Basic applications
'----------------------------------------------------------------
' Source file:  FRAME02.BAS
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
  Case 1:  FileName = "FRAME1.BMP": zLoadFrame
  Case 2:  FileName = "FRAME2.BMP": zLoadFrame
  Case 3:  FileName = "FRAME3.BMP": zLoadFrame
  Case 4:  FileName = "FRAME4.BMP": zLoadFrame
  Case 5:  FileName = "FRAME5.BMP": zLoadFrame
  Case 6:  FileName = "FRAME6.BMP": zLoadFrame
  Case 7:  FileName = "FRAME7.BMP": zLoadFrame
  Case 8:  FileName = "FRAME8.BMP": zLoadFrame
  Case 9:  FileName = "FRAME9.BMP": zLoadFrame
  Case 10:  FileName = "FRAME10.BMP": zLoadFrame
  Case 11:  FileName = "FRAME11.BMP": zLoadFrame
  Case 12:  FileName = "FRAME12.BMP": zLoadFrame
  Case 13:  FileName = "FRAME13.BMP": zLoadFrame
  Case 14:  FileName = "FRAME14.BMP": zLoadFrame
  Case 15:  FileName = "FRAME15.BMP": zLoadFrame
  Case 16:  FileName = "FRAME16.BMP": zLoadFrame
  Case 17:  FileName = "FRAME17.BMP": zLoadFrame
  Case 18:  FileName = "FRAME18.BMP": zLoadFrame
  Case 19:  FileName = "FRAME19.BMP": zLoadFrame
End Select
If FrameLoaded = False Then
  AnimationHalted = True
  Beep
  MsgBox "Unable to load next frame from disk.  Animation halted.", MB_OK, "Unexpected animation condition"
  Exit Sub
End If
End Sub

'----------------------------------------------------------------
'          Draw one key drawing and place cel on frame
'----------------------------------------------------------------
Sub zDoKeyDrawing ()
  Dim Repeat As Integer                             'loop counter
  Dim DrawPtr As Integer      'dynamic index into array of coords
  Dim ColorBrush As Integer, PrevBrush As Integer        'brushes
  Dim ColorPen As Integer, PrevPen As Integer               'pens
  Static Poly4(3, 1) As Integer                'array of 4 points
  Static Poly5(4, 1) As Integer                'array of 5 points
ColorPen = CreatePen(PS_SOLID, 1, RGB(255, 255, 0))   'create pen
ColorBrush = CreateSolidBrush(RGB(255, 255, 0))     'create brush
PrevPen = SelectObject(Form1.hDC, ColorPen)           'select pen
PrevBrush = SelectObject(Form1.hDC, ColorBrush)     'select brush
'------------------- branch to appropriate code -----------------
Select Case DrawingNum
  Case 1:  GoTo DRAW_CEL_1
  Case 2:  GoTo DRAW_CEL_2
  Case 3:  GoTo DRAW_CEL_3
  Case 4:  GoTo DRAW_CEL_4
  Case 5:  GoTo DRAW_CEL_5
  Case 6:  GoTo DRAW_CEL_6
  Case 7:  GoTo DRAW_CEL_7
End Select
MsgBox "Internal error in Select Case structure in zDoKeyDrawing( ).", MB_OK, "Book 4224 error report"
Exit Sub                                  'trap any coding errors
'--------------------------- draw cel 1 -------------------------
DRAW_CEL_1:
RetVal = Ellipse(Form1.hDC, image1(0) - 10, image1(1) - 10, image1(0) + 10, image1(1) + 10)
DrawPtr = 2                 'update the index into database array
Poly4(0, 0) = image1(DrawPtr): DrawPtr = DrawPtr + 1
Poly4(1, 0) = image1(DrawPtr): DrawPtr = DrawPtr + 1
Poly4(2, 0) = image1(DrawPtr): DrawPtr = DrawPtr + 1
Poly4(3, 0) = image1(DrawPtr): DrawPtr = DrawPtr + 1
Poly4(0, 1) = image1(DrawPtr): DrawPtr = DrawPtr + 1
Poly4(1, 1) = image1(DrawPtr): DrawPtr = DrawPtr + 1
Poly4(2, 1) = image1(DrawPtr): DrawPtr = DrawPtr + 1
Poly4(3, 1) = image1(DrawPtr): DrawPtr = DrawPtr + 1
RetVal = Polygon(Form1.hDC, Poly4(0, 0), 4)        'draw the neck
Poly5(0, 0) = image1(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(1, 0) = image1(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(2, 0) = image1(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(3, 0) = image1(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(4, 0) = image1(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(0, 1) = image1(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(1, 1) = image1(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(2, 1) = image1(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(3, 1) = image1(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(4, 1) = image1(DrawPtr): DrawPtr = DrawPtr + 1
RetVal = Polygon(Form1.hDC, Poly5(0, 0), 5)       'draw the torso
For Repeat = 1 To 13
  Poly4(0, 0) = image1(DrawPtr): DrawPtr = DrawPtr + 1
  Poly4(1, 0) = image1(DrawPtr): DrawPtr = DrawPtr + 1
  Poly4(2, 0) = image1(DrawPtr): DrawPtr = DrawPtr + 1
  Poly4(3, 0) = image1(DrawPtr): DrawPtr = DrawPtr + 1
  Poly4(0, 1) = image1(DrawPtr): DrawPtr = DrawPtr + 1
  Poly4(1, 1) = image1(DrawPtr): DrawPtr = DrawPtr + 1
  Poly4(2, 1) = image1(DrawPtr): DrawPtr = DrawPtr + 1
  Poly4(3, 1) = image1(DrawPtr): DrawPtr = DrawPtr + 1
  RetVal = Polygon(Form1.hDC, Poly4(0, 0), 4)'draw remaining parts
Next Repeat
GoTo All_DONE
'--------------------------- draw cel 2 -------------------------
DRAW_CEL_2:
RetVal = Ellipse(Form1.hDC, image2(0) - 10, image2(1) - 10, image2(0) + 10, image2(1) + 10)
DrawPtr = 2
Poly4(0, 0) = image2(DrawPtr): DrawPtr = DrawPtr + 1
Poly4(1, 0) = image2(DrawPtr): DrawPtr = DrawPtr + 1
Poly4(2, 0) = image2(DrawPtr): DrawPtr = DrawPtr + 1
Poly4(3, 0) = image2(DrawPtr): DrawPtr = DrawPtr + 1
Poly4(0, 1) = image2(DrawPtr): DrawPtr = DrawPtr + 1
Poly4(1, 1) = image2(DrawPtr): DrawPtr = DrawPtr + 1
Poly4(2, 1) = image2(DrawPtr): DrawPtr = DrawPtr + 1
Poly4(3, 1) = image2(DrawPtr): DrawPtr = DrawPtr + 1
RetVal = Polygon(Form1.hDC, Poly4(0, 0), 4)
Poly5(0, 0) = image2(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(1, 0) = image2(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(2, 0) = image2(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(3, 0) = image2(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(4, 0) = image2(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(0, 1) = image2(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(1, 1) = image2(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(2, 1) = image2(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(3, 1) = image2(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(4, 1) = image2(DrawPtr): DrawPtr = DrawPtr + 1
RetVal = Polygon(Form1.hDC, Poly5(0, 0), 5)
For Repeat = 1 To 13
  Poly4(0, 0) = image2(DrawPtr): DrawPtr = DrawPtr + 1
  Poly4(1, 0) = image2(DrawPtr): DrawPtr = DrawPtr + 1
  Poly4(2, 0) = image2(DrawPtr): DrawPtr = DrawPtr + 1
  Poly4(3, 0) = image2(DrawPtr): DrawPtr = DrawPtr + 1
  Poly4(0, 1) = image2(DrawPtr): DrawPtr = DrawPtr + 1
  Poly4(1, 1) = image2(DrawPtr): DrawPtr = DrawPtr + 1
  Poly4(2, 1) = image2(DrawPtr): DrawPtr = DrawPtr + 1
  Poly4(3, 1) = image2(DrawPtr): DrawPtr = DrawPtr + 1
  RetVal = Polygon(Form1.hDC, Poly4(0, 0), 4)
Next Repeat
GoTo All_DONE
'--------------------------- draw cel 3 -------------------------
DRAW_CEL_3:
RetVal = Ellipse(Form1.hDC, image3(0) - 10, image3(1) - 10, image3(0) + 10, image3(1) + 10)
DrawPtr = 2
Poly4(0, 0) = image3(DrawPtr): DrawPtr = DrawPtr + 1
Poly4(1, 0) = image3(DrawPtr): DrawPtr = DrawPtr + 1
Poly4(2, 0) = image3(DrawPtr): DrawPtr = DrawPtr + 1
Poly4(3, 0) = image3(DrawPtr): DrawPtr = DrawPtr + 1
Poly4(0, 1) = image3(DrawPtr): DrawPtr = DrawPtr + 1
Poly4(1, 1) = image3(DrawPtr): DrawPtr = DrawPtr + 1
Poly4(2, 1) = image3(DrawPtr): DrawPtr = DrawPtr + 1
Poly4(3, 1) = image3(DrawPtr): DrawPtr = DrawPtr + 1
RetVal = Polygon(Form1.hDC, Poly4(0, 0), 4)
Poly5(0, 0) = image3(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(1, 0) = image3(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(2, 0) = image3(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(3, 0) = image3(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(4, 0) = image3(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(0, 1) = image3(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(1, 1) = image3(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(2, 1) = image3(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(3, 1) = image3(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(4, 1) = image3(DrawPtr): DrawPtr = DrawPtr + 1
RetVal = Polygon(Form1.hDC, Poly5(0, 0), 5)
For Repeat = 1 To 13
  Poly4(0, 0) = image3(DrawPtr): DrawPtr = DrawPtr + 1
  Poly4(1, 0) = image3(DrawPtr): DrawPtr = DrawPtr + 1
  Poly4(2, 0) = image3(DrawPtr): DrawPtr = DrawPtr + 1
  Poly4(3, 0) = image3(DrawPtr): DrawPtr = DrawPtr + 1
  Poly4(0, 1) = image3(DrawPtr): DrawPtr = DrawPtr + 1
  Poly4(1, 1) = image3(DrawPtr): DrawPtr = DrawPtr + 1
  Poly4(2, 1) = image3(DrawPtr): DrawPtr = DrawPtr + 1
  Poly4(3, 1) = image3(DrawPtr): DrawPtr = DrawPtr + 1
  RetVal = Polygon(Form1.hDC, Poly4(0, 0), 4)
Next Repeat
GoTo All_DONE
'--------------------------- draw cel 4 -------------------------
DRAW_CEL_4:
RetVal = Ellipse(Form1.hDC, image4(0) - 10, image4(1) - 10, image4(0) + 10, image4(1) + 10)
DrawPtr = 2
Poly4(0, 0) = image4(DrawPtr): DrawPtr = DrawPtr + 1
Poly4(1, 0) = image4(DrawPtr): DrawPtr = DrawPtr + 1
Poly4(2, 0) = image4(DrawPtr): DrawPtr = DrawPtr + 1
Poly4(3, 0) = image4(DrawPtr): DrawPtr = DrawPtr + 1
Poly4(0, 1) = image4(DrawPtr): DrawPtr = DrawPtr + 1
Poly4(1, 1) = image4(DrawPtr): DrawPtr = DrawPtr + 1
Poly4(2, 1) = image4(DrawPtr): DrawPtr = DrawPtr + 1
Poly4(3, 1) = image4(DrawPtr): DrawPtr = DrawPtr + 1
RetVal = Polygon(Form1.hDC, Poly4(0, 0), 4)
Poly5(0, 0) = image4(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(1, 0) = image4(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(2, 0) = image4(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(3, 0) = image4(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(4, 0) = image4(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(0, 1) = image4(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(1, 1) = image4(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(2, 1) = image4(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(3, 1) = image4(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(4, 1) = image4(DrawPtr): DrawPtr = DrawPtr + 1
RetVal = Polygon(Form1.hDC, Poly5(0, 0), 5)
For Repeat = 1 To 13
  Poly4(0, 0) = image4(DrawPtr): DrawPtr = DrawPtr + 1
  Poly4(1, 0) = image4(DrawPtr): DrawPtr = DrawPtr + 1
  Poly4(2, 0) = image4(DrawPtr): DrawPtr = DrawPtr + 1
  Poly4(3, 0) = image4(DrawPtr): DrawPtr = DrawPtr + 1
  Poly4(0, 1) = image4(DrawPtr): DrawPtr = DrawPtr + 1
  Poly4(1, 1) = image4(DrawPtr): DrawPtr = DrawPtr + 1
  Poly4(2, 1) = image4(DrawPtr): DrawPtr = DrawPtr + 1
  Poly4(3, 1) = image4(DrawPtr): DrawPtr = DrawPtr + 1
  RetVal = Polygon(Form1.hDC, Poly4(0, 0), 4)
Next Repeat
GoTo All_DONE
'--------------------------- draw cel 5 -------------------------
DRAW_CEL_5:
RetVal = Ellipse(Form1.hDC, image5(0) - 10, image5(1) - 10, image5(0) + 10, image5(1) + 10)
DrawPtr = 2
Poly4(0, 0) = image5(DrawPtr): DrawPtr = DrawPtr + 1
Poly4(1, 0) = image5(DrawPtr): DrawPtr = DrawPtr + 1
Poly4(2, 0) = image5(DrawPtr): DrawPtr = DrawPtr + 1
Poly4(3, 0) = image5(DrawPtr): DrawPtr = DrawPtr + 1
Poly4(0, 1) = image5(DrawPtr): DrawPtr = DrawPtr + 1
Poly4(1, 1) = image5(DrawPtr): DrawPtr = DrawPtr + 1
Poly4(2, 1) = image5(DrawPtr): DrawPtr = DrawPtr + 1
Poly4(3, 1) = image5(DrawPtr): DrawPtr = DrawPtr + 1
RetVal = Polygon(Form1.hDC, Poly4(0, 0), 4)
Poly5(0, 0) = image5(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(1, 0) = image5(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(2, 0) = image5(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(3, 0) = image5(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(4, 0) = image5(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(0, 1) = image5(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(1, 1) = image5(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(2, 1) = image5(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(3, 1) = image5(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(4, 1) = image5(DrawPtr): DrawPtr = DrawPtr + 1
RetVal = Polygon(Form1.hDC, Poly5(0, 0), 5)
For Repeat = 1 To 13
  Poly4(0, 0) = image5(DrawPtr): DrawPtr = DrawPtr + 1
  Poly4(1, 0) = image5(DrawPtr): DrawPtr = DrawPtr + 1
  Poly4(2, 0) = image5(DrawPtr): DrawPtr = DrawPtr + 1
  Poly4(3, 0) = image5(DrawPtr): DrawPtr = DrawPtr + 1
  Poly4(0, 1) = image5(DrawPtr): DrawPtr = DrawPtr + 1
  Poly4(1, 1) = image5(DrawPtr): DrawPtr = DrawPtr + 1
  Poly4(2, 1) = image5(DrawPtr): DrawPtr = DrawPtr + 1
  Poly4(3, 1) = image5(DrawPtr): DrawPtr = DrawPtr + 1
  RetVal = Polygon(Form1.hDC, Poly4(0, 0), 4)
Next Repeat
GoTo All_DONE
'--------------------------- draw cel 6 -------------------------
DRAW_CEL_6:
RetVal = Ellipse(Form1.hDC, image6(0) - 10, image6(1) - 10, image6(0) + 10, image6(1) + 10)
DrawPtr = 2
Poly4(0, 0) = image6(DrawPtr): DrawPtr = DrawPtr + 1
Poly4(1, 0) = image6(DrawPtr): DrawPtr = DrawPtr + 1
Poly4(2, 0) = image6(DrawPtr): DrawPtr = DrawPtr + 1
Poly4(3, 0) = image6(DrawPtr): DrawPtr = DrawPtr + 1
Poly4(0, 1) = image6(DrawPtr): DrawPtr = DrawPtr + 1
Poly4(1, 1) = image6(DrawPtr): DrawPtr = DrawPtr + 1
Poly4(2, 1) = image6(DrawPtr): DrawPtr = DrawPtr + 1
Poly4(3, 1) = image6(DrawPtr): DrawPtr = DrawPtr + 1
RetVal = Polygon(Form1.hDC, Poly4(0, 0), 4)
Poly5(0, 0) = image6(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(1, 0) = image6(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(2, 0) = image6(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(3, 0) = image6(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(4, 0) = image6(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(0, 1) = image6(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(1, 1) = image6(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(2, 1) = image6(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(3, 1) = image6(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(4, 1) = image6(DrawPtr): DrawPtr = DrawPtr + 1
RetVal = Polygon(Form1.hDC, Poly5(0, 0), 5)
For Repeat = 1 To 13
  Poly4(0, 0) = image6(DrawPtr): DrawPtr = DrawPtr + 1
  Poly4(1, 0) = image6(DrawPtr): DrawPtr = DrawPtr + 1
  Poly4(2, 0) = image6(DrawPtr): DrawPtr = DrawPtr + 1
  Poly4(3, 0) = image6(DrawPtr): DrawPtr = DrawPtr + 1
  Poly4(0, 1) = image6(DrawPtr): DrawPtr = DrawPtr + 1
  Poly4(1, 1) = image6(DrawPtr): DrawPtr = DrawPtr + 1
  Poly4(2, 1) = image6(DrawPtr): DrawPtr = DrawPtr + 1
  Poly4(3, 1) = image6(DrawPtr): DrawPtr = DrawPtr + 1
  RetVal = Polygon(Form1.hDC, Poly4(0, 0), 4)
Next Repeat
GoTo All_DONE
'--------------------------- draw cel 7 -------------------------
DRAW_CEL_7:
RetVal = Ellipse(Form1.hDC, image7(0) - 10, image7(1) - 10, image7(0) + 10, image7(1) + 10)
DrawPtr = 2
Poly4(0, 0) = image7(DrawPtr): DrawPtr = DrawPtr + 1
Poly4(1, 0) = image7(DrawPtr): DrawPtr = DrawPtr + 1
Poly4(2, 0) = image7(DrawPtr): DrawPtr = DrawPtr + 1
Poly4(3, 0) = image7(DrawPtr): DrawPtr = DrawPtr + 1
Poly4(0, 1) = image7(DrawPtr): DrawPtr = DrawPtr + 1
Poly4(1, 1) = image7(DrawPtr): DrawPtr = DrawPtr + 1
Poly4(2, 1) = image7(DrawPtr): DrawPtr = DrawPtr + 1
Poly4(3, 1) = image7(DrawPtr): DrawPtr = DrawPtr + 1
RetVal = Polygon(Form1.hDC, Poly4(0, 0), 4)
Poly5(0, 0) = image7(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(1, 0) = image7(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(2, 0) = image7(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(3, 0) = image7(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(4, 0) = image7(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(0, 1) = image7(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(1, 1) = image7(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(2, 1) = image7(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(3, 1) = image7(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(4, 1) = image7(DrawPtr): DrawPtr = DrawPtr + 1
RetVal = Polygon(Form1.hDC, Poly5(0, 0), 5)
For Repeat = 1 To 13
  Poly4(0, 0) = image7(DrawPtr): DrawPtr = DrawPtr + 1
  Poly4(1, 0) = image7(DrawPtr): DrawPtr = DrawPtr + 1
  Poly4(2, 0) = image7(DrawPtr): DrawPtr = DrawPtr + 1
  Poly4(3, 0) = image7(DrawPtr): DrawPtr = DrawPtr + 1
  Poly4(0, 1) = image7(DrawPtr): DrawPtr = DrawPtr + 1
  Poly4(1, 1) = image7(DrawPtr): DrawPtr = DrawPtr + 1
  Poly4(2, 1) = image7(DrawPtr): DrawPtr = DrawPtr + 1
  Poly4(3, 1) = image7(DrawPtr): DrawPtr = DrawPtr + 1
  RetVal = Polygon(Form1.hDC, Poly4(0, 0), 4)
Next Repeat
'----------------------- tidy up and return ---------------------
All_DONE:
RetVal = SelectObject(Form1.hDC, PrevPen)           'deselect pen
RetVal = SelectObject(Form1.hDC, PrevBrush)       'deselect brush
RetVal = DeleteObject(ColorPen)                       'delete pen
RetVal = DeleteObject(ColorBrush)                   'delete brush
End Sub

'----------------------------------------------------------------
'        Initialize the database of drawing coordinates
'----------------------------------------------------------------
Sub zInitDatabase ()
  Dim Num As Integer             'index into array of coordinates
'--------------- create the database for cel 1 ------------------
image1(0) = 36: image1(1) = 16                            'circle
image1(2) = 45: image1(3) = 23: image1(4) = 49              'neck
image1(5) = 27: image1(6) = 47: image1(7) = 28
image1(8) = 42: image1(9) = 25
image1(10) = 54: image1(11) = 26: image1(12) = 78          'torso
image1(13) = 54: image1(14) = 76: image1(15) = 60
image1(16) = 57: image1(17) = 57: image1(18) = 42
image1(19) = 30
image1(20) = 15: image1(21) = 25: image1(22) = 24           'hand
image1(23) = 27: image1(24) = 23: image1(25) = 33
image1(26) = 12: image1(27) = 32
image1(28) = 22: image1(29) = 32                 'leading forearm
image1(30) = 30: image1(31) = 44: image1(32) = 26
image1(33) = 46: image1(34) = 20: image1(35) = 35
image1(36) = 45: image1(37) = 35               'leading upper arm
image1(38) = 48: image1(39) = 38: image1(40) = 25
image1(41) = 46: image1(42) = 23: image1(43) = 44
image1(44) = 55: image1(45) = 28              'trailing upper arm
image1(46) = 69: image1(47) = 20: image1(48) = 70
image1(49) = 23: image1(50) = 58: image1(51) = 31
image1(52) = 68: image1(53) = 20                'trailing forearm
image1(54) = 91: image1(55) = 20: image1(56) = 91
image1(57) = 23: image1(58) = 70: image1(59) = 23
image1(60) = 91: image1(61) = 20                   'trailing hand
image1(62) = 100: image1(63) = 15: image1(64) = 100
image1(65) = 18: image1(66) = 91: image1(67) = 23
image1(68) = 31: image1(69) = 52                   'leading thigh
image1(70) = 65: image1(71) = 55: image1(72) = 71
image1(73) = 60: image1(74) = 35: image1(75) = 55
image1(76) = 40: image1(77) = 57                    'leading calf
image1(78) = 51: image1(79) = 75: image1(80) = 41
image1(81) = 74: image1(82) = 31: image1(83) = 52
image1(84) = 46: image1(85) = 68                    'leading foot
image1(86) = 51: image1(87) = 75: image1(88) = 22
image1(89) = 72: image1(90) = 40: image1(91) = 68
image1(92) = 76: image1(93) = 60                  'trailing thigh
image1(94) = 106: image1(95) = 92: image1(96) = 101
image1(97) = 94: image1(98) = 69: image1(99) = 60
image1(100) = 76: image1(101) = 60                 'trailing calf
image1(102) = 106: image1(103) = 92: image1(104) = 101
image1(105) = 94: image1(106) = 69: image1(107) = 60
image1(108) = 104: image1(109) = 92                'trailing foot
image1(110) = 111: image1(111) = 92: image1(112) = 111
image1(113) = 108: image1(114) = 104: image1(115) = 106
image1(116) = 111: image1(117) = 108                'trailing toe
image1(118) = 105: image1(119) = 115: image1(120) = 100
image1(121) = 115: image1(122) = 104: image1(123) = 107
'--------------- create the database for cel 2 ------------------
Num = 0
image2(Num) = 36
Num = Num + 1: image2(Num) = 13: Num = Num + 1: image2(Num) = 44
Num = Num + 1: image2(Num) = 20: Num = Num + 1: image2(Num) = 48
Num = Num + 1: image2(Num) = 24: Num = Num + 1: image2(Num) = 43
Num = Num + 1: image2(Num) = 24: Num = Num + 1: image2(Num) = 41
Num = Num + 1: image2(Num) = 21: Num = Num + 1: image2(Num) = 51
Num = Num + 1: image2(Num) = 22: Num = Num + 1: image2(Num) = 77
Num = Num + 1: image2(Num) = 55: Num = Num + 1: image2(Num) = 71
Num = Num + 1: image2(Num) = 58: Num = Num + 1: image2(Num) = 54
Num = Num + 1: image2(Num) = 56: Num = Num + 1: image2(Num) = 40
Num = Num + 1: image2(Num) = 25: Num = Num + 1: image2(Num) = 20
Num = Num + 1: image2(Num) = 31: Num = Num + 1: image2(Num) = 20
Num = Num + 1: image2(Num) = 39: Num = Num + 1: image2(Num) = 12
Num = Num + 1: image2(Num) = 40: Num = Num + 1: image2(Num) = 10
Num = Num + 1: image2(Num) = 33: Num = Num + 1: image2(Num) = 20
Num = Num + 1: image2(Num) = 36: Num = Num + 1: image2(Num) = 34
Num = Num + 1: image2(Num) = 43: Num = Num + 1: image2(Num) = 35
Num = Num + 1: image2(Num) = 45: Num = Num + 1: image2(Num) = 20
Num = Num + 1: image2(Num) = 39: Num = Num + 1: image2(Num) = 44
Num = Num + 1: image2(Num) = 32: Num = Num + 1: image2(Num) = 46
Num = Num + 1: image2(Num) = 36: Num = Num + 1: image2(Num) = 35
Num = Num + 1: image2(Num) = 45: Num = Num + 1: image2(Num) = 34
Num = Num + 1: image2(Num) = 42: Num = Num + 1: image2(Num) = 54
Num = Num + 1: image2(Num) = 25: Num = Num + 1: image2(Num) = 76
Num = Num + 1: image2(Num) = 25: Num = Num + 1: image2(Num) = 76
Num = Num + 1: image2(Num) = 28: Num = Num + 1: image2(Num) = 54
Num = Num + 1: image2(Num) = 28: Num = Num + 1: image2(Num) = 78
Num = Num + 1: image2(Num) = 25: Num = Num + 1: image2(Num) = 95
Num = Num + 1: image2(Num) = 33: Num = Num + 1: image2(Num) = 95
Num = Num + 1: image2(Num) = 37: Num = Num + 1: image2(Num) = 74
Num = Num + 1: image2(Num) = 28: Num = Num + 1: image2(Num) = 95
Num = Num + 1: image2(Num) = 33: Num = Num + 1: image2(Num) = 105
Num = Num + 1: image2(Num) = 30: Num = Num + 1: image2(Num) = 104
Num = Num + 1: image2(Num) = 35: Num = Num + 1: image2(Num) = 95
Num = Num + 1: image2(Num) = 37: Num = Num + 1: image2(Num) = 52
Num = Num + 1: image2(Num) = 52: Num = Num + 1: image2(Num) = 54
Num = Num + 1: image2(Num) = 55: Num = Num + 1: image2(Num) = 32
Num = Num + 1: image2(Num) = 60: Num = Num + 1: image2(Num) = 28
Num = Num + 1: image2(Num) = 59: Num = Num + 1: image2(Num) = 28
Num = Num + 1: image2(Num) = 59: Num = Num + 1: image2(Num) = 32
Num = Num + 1: image2(Num) = 60: Num = Num + 1: image2(Num) = 32
Num = Num + 1: image2(Num) = 85: Num = Num + 1: image2(Num) = 28
Num = Num + 1: image2(Num) = 78: Num = Num + 1: image2(Num) = 32
Num = Num + 1: image2(Num) = 78: Num = Num + 1: image2(Num) = 32
Num = Num + 1: image2(Num) = 85: Num = Num + 1: image2(Num) = 10
Num = Num + 1: image2(Num) = 75: Num = Num + 1: image2(Num) = 28
Num = Num + 1: image2(Num) = 78: Num = Num + 1: image2(Num) = 75
Num = Num + 1: image2(Num) = 56: Num = Num + 1: image2(Num) = 92
Num = Num + 1: image2(Num) = 73: Num = Num + 1: image2(Num) = 88
Num = Num + 1: image2(Num) = 76: Num = Num + 1: image2(Num) = 70
Num = Num + 1: image2(Num) = 57: Num = Num + 1: image2(Num) = 92
Num = Num + 1: image2(Num) = 74: Num = Num + 1: image2(Num) = 116
Num = Num + 1: image2(Num) = 86: Num = Num + 1: image2(Num) = 112
Num = Num + 1: image2(Num) = 90: Num = Num + 1: image2(Num) = 88
Num = Num + 1: image2(Num) = 76: Num = Num + 1: image2(Num) = 116
Num = Num + 1: image2(Num) = 86: Num = Num + 1: image2(Num) = 120
Num = Num + 1: image2(Num) = 85: Num = Num + 1: image2(Num) = 120
Num = Num + 1: image2(Num) = 108: Num = Num + 1: image2(Num) = 112
Num = Num + 1: image2(Num) = 90: Num = Num + 1: image2(Num) = 116
Num = Num + 1: image2(Num) = 86: Num = Num + 1: image2(Num) = 120
Num = Num + 1: image2(Num) = 85: Num = Num + 1: image2(Num) = 120
Num = Num + 1: image2(Num) = 108: Num = Num + 1: image2(Num) = 112
Num = Num + 1: image2(Num) = 90
'--------------- create the database for cel 3 ------------------
Num = 0
image3(Num) = 38
Num = Num + 1: image3(Num) = 10: Num = Num + 1: image3(Num) = 45
Num = Num + 1: image3(Num) = 16: Num = Num + 1: image3(Num) = 48
Num = Num + 1: image3(Num) = 21: Num = Num + 1: image3(Num) = 44
Num = Num + 1: image3(Num) = 22: Num = Num + 1: image3(Num) = 41
Num = Num + 1: image3(Num) = 17: Num = Num + 1: image3(Num) = 52
Num = Num + 1: image3(Num) = 20: Num = Num + 1: image3(Num) = 75
Num = Num + 1: image3(Num) = 55: Num = Num + 1: image3(Num) = 70
Num = Num + 1: image3(Num) = 56: Num = Num + 1: image3(Num) = 53
Num = Num + 1: image3(Num) = 54: Num = Num + 1: image3(Num) = 40
Num = Num + 1: image3(Num) = 22: Num = Num + 1: image3(Num) = 20
Num = Num + 1: image3(Num) = 36: Num = Num + 1: image3(Num) = 24
Num = Num + 1: image3(Num) = 44: Num = Num + 1: image3(Num) = 17
Num = Num + 1: image3(Num) = 47: Num = Num + 1: image3(Num) = 13
Num = Num + 1: image3(Num) = 40: Num = Num + 1: image3(Num) = 22
Num = Num + 1: image3(Num) = 40: Num = Num + 1: image3(Num) = 40
Num = Num + 1: image3(Num) = 43: Num = Num + 1: image3(Num) = 40
Num = Num + 1: image3(Num) = 47: Num = Num + 1: image3(Num) = 22
Num = Num + 1: image3(Num) = 44: Num = Num + 1: image3(Num) = 37
Num = Num + 1: image3(Num) = 43: Num = Num + 1: image3(Num) = 44
Num = Num + 1: image3(Num) = 31: Num = Num + 1: image3(Num) = 47
Num = Num + 1: image3(Num) = 36: Num = Num + 1: image3(Num) = 40
Num = Num + 1: image3(Num) = 47: Num = Num + 1: image3(Num) = 55
Num = Num + 1: image3(Num) = 25: Num = Num + 1: image3(Num) = 76
Num = Num + 1: image3(Num) = 29: Num = Num + 1: image3(Num) = 74
Num = Num + 1: image3(Num) = 31: Num = Num + 1: image3(Num) = 55
Num = Num + 1: image3(Num) = 29: Num = Num + 1: image3(Num) = 76
Num = Num + 1: image3(Num) = 29: Num = Num + 1: image3(Num) = 90
Num = Num + 1: image3(Num) = 45: Num = Num + 1: image3(Num) = 85
Num = Num + 1: image3(Num) = 48: Num = Num + 1: image3(Num) = 74
Num = Num + 1: image3(Num) = 31: Num = Num + 1: image3(Num) = 90
Num = Num + 1: image3(Num) = 45: Num = Num + 1: image3(Num) = 100
Num = Num + 1: image3(Num) = 45: Num = Num + 1: image3(Num) = 95
Num = Num + 1: image3(Num) = 49: Num = Num + 1: image3(Num) = 85
Num = Num + 1: image3(Num) = 49: Num = Num + 1: image3(Num) = 53
Num = Num + 1: image3(Num) = 54: Num = Num + 1: image3(Num) = 55
Num = Num + 1: image3(Num) = 58: Num = Num + 1: image3(Num) = 35
Num = Num + 1: image3(Num) = 72: Num = Num + 1: image3(Num) = 30
Num = Num + 1: image3(Num) = 69: Num = Num + 1: image3(Num) = 35
Num = Num + 1: image3(Num) = 72: Num = Num + 1: image3(Num) = 16
Num = Num + 1: image3(Num) = 96: Num = Num + 1: image3(Num) = 17
Num = Num + 1: image3(Num) = 85: Num = Num + 1: image3(Num) = 30
Num = Num + 1: image3(Num) = 69: Num = Num + 1: image3(Num) = 17
Num = Num + 1: image3(Num) = 85: Num = Num + 1: image3(Num) = 20
Num = Num + 1: image3(Num) = 90: Num = Num + 1: image3(Num) = 17
Num = Num + 1: image3(Num) = 96: Num = Num + 1: image3(Num) = 0
Num = Num + 1: image3(Num) = 80: Num = Num + 1: image3(Num) = 75
Num = Num + 1: image3(Num) = 55: Num = Num + 1: image3(Num) = 92
Num = Num + 1: image3(Num) = 70: Num = Num + 1: image3(Num) = 88
Num = Num + 1: image3(Num) = 74: Num = Num + 1: image3(Num) = 70
Num = Num + 1: image3(Num) = 57: Num = Num + 1: image3(Num) = 92
Num = Num + 1: image3(Num) = 70: Num = Num + 1: image3(Num) = 121
Num = Num + 1: image3(Num) = 77: Num = Num + 1: image3(Num) = 122
Num = Num + 1: image3(Num) = 81: Num = Num + 1: image3(Num) = 87
Num = Num + 1: image3(Num) = 73: Num = Num + 1: image3(Num) = 121
Num = Num + 1: image3(Num) = 77: Num = Num + 1: image3(Num) = 126
Num = Num + 1: image3(Num) = 75: Num = Num + 1: image3(Num) = 139
Num = Num + 1: image3(Num) = 95: Num = Num + 1: image3(Num) = 122
Num = Num + 1: image3(Num) = 81: Num = Num + 1: image3(Num) = 121
Num = Num + 1: image3(Num) = 77: Num = Num + 1: image3(Num) = 126
Num = Num + 1: image3(Num) = 75: Num = Num + 1: image3(Num) = 139
Num = Num + 1: image3(Num) = 95: Num = Num + 1: image3(Num) = 122
Num = Num + 1: image3(Num) = 81
'--------------- create the database for cel 4 ------------------
Num = 0
image4(Num) = 38
Num = Num + 1: image4(Num) = 13: Num = Num + 1: image4(Num) = 45
Num = Num + 1: image4(Num) = 19: Num = Num + 1: image4(Num) = 48
Num = Num + 1: image4(Num) = 23: Num = Num + 1: image4(Num) = 44
Num = Num + 1: image4(Num) = 25: Num = Num + 1: image4(Num) = 41
Num = Num + 1: image4(Num) = 20: Num = Num + 1: image4(Num) = 52
Num = Num + 1: image4(Num) = 22: Num = Num + 1: image4(Num) = 70
Num = Num + 1: image4(Num) = 56: Num = Num + 1: image4(Num) = 64
Num = Num + 1: image4(Num) = 60: Num = Num + 1: image4(Num) = 46
Num = Num + 1: image4(Num) = 60: Num = Num + 1: image4(Num) = 40
Num = Num + 1: image4(Num) = 26: Num = Num + 1: image4(Num) = 32
Num = Num + 1: image4(Num) = 43: Num = Num + 1: image4(Num) = 36
Num = Num + 1: image4(Num) = 46: Num = Num + 1: image4(Num) = 34
Num = Num + 1: image4(Num) = 51: Num = Num + 1: image4(Num) = 27
Num = Num + 1: image4(Num) = 48: Num = Num + 1: image4(Num) = 36
Num = Num + 1: image4(Num) = 46: Num = Num + 1: image4(Num) = 50
Num = Num + 1: image4(Num) = 44: Num = Num + 1: image4(Num) = 50
Num = Num + 1: image4(Num) = 48: Num = Num + 1: image4(Num) = 36
Num = Num + 1: image4(Num) = 49: Num = Num + 1: image4(Num) = 36
Num = Num + 1: image4(Num) = 46: Num = Num + 1: image4(Num) = 50
Num = Num + 1: image4(Num) = 44: Num = Num + 1: image4(Num) = 50
Num = Num + 1: image4(Num) = 48: Num = Num + 1: image4(Num) = 36
Num = Num + 1: image4(Num) = 49: Num = Num + 1: image4(Num) = 55
Num = Num + 1: image4(Num) = 35: Num = Num + 1: image4(Num) = 69
Num = Num + 1: image4(Num) = 45: Num = Num + 1: image4(Num) = 62
Num = Num + 1: image4(Num) = 50: Num = Num + 1: image4(Num) = 55
Num = Num + 1: image4(Num) = 40: Num = Num + 1: image4(Num) = 55
Num = Num + 1: image4(Num) = 35: Num = Num + 1: image4(Num) = 69
Num = Num + 1: image4(Num) = 45: Num = Num + 1: image4(Num) = 62
Num = Num + 1: image4(Num) = 50: Num = Num + 1: image4(Num) = 55
Num = Num + 1: image4(Num) = 40: Num = Num + 1: image4(Num) = 55
Num = Num + 1: image4(Num) = 35: Num = Num + 1: image4(Num) = 69
Num = Num + 1: image4(Num) = 45: Num = Num + 1: image4(Num) = 62
Num = Num + 1: image4(Num) = 50: Num = Num + 1: image4(Num) = 55
Num = Num + 1: image4(Num) = 40: Num = Num + 1: image4(Num) = 46
Num = Num + 1: image4(Num) = 59: Num = Num + 1: image4(Num) = 51
Num = Num + 1: image4(Num) = 60: Num = Num + 1: image4(Num) = 31
Num = Num + 1: image4(Num) = 75: Num = Num + 1: image4(Num) = 26
Num = Num + 1: image4(Num) = 73: Num = Num + 1: image4(Num) = 26
Num = Num + 1: image4(Num) = 73: Num = Num + 1: image4(Num) = 31
Num = Num + 1: image4(Num) = 76: Num = Num + 1: image4(Num) = 25
Num = Num + 1: image4(Num) = 105: Num = Num + 1: image4(Num) = 22
Num = Num + 1: image4(Num) = 97: Num = Num + 1: image4(Num) = 22
Num = Num + 1: image4(Num) = 97: Num = Num + 1: image4(Num) = 26
Num = Num + 1: image4(Num) = 99: Num = Num + 1: image4(Num) = 25
Num = Num + 1: image4(Num) = 105: Num = Num + 1: image4(Num) = 0
Num = Num + 1: image4(Num) = 95: Num = Num + 1: image4(Num) = 58
Num = Num + 1: image4(Num) = 60: Num = Num + 1: image4(Num) = 64
Num = Num + 1: image4(Num) = 60: Num = Num + 1: image4(Num) = 58
Num = Num + 1: image4(Num) = 78: Num = Num + 1: image4(Num) = 50
Num = Num + 1: image4(Num) = 78: Num = Num + 1: image4(Num) = 58
Num = Num + 1: image4(Num) = 78: Num = Num + 1: image4(Num) = 75
Num = Num + 1: image4(Num) = 85: Num = Num + 1: image4(Num) = 74
Num = Num + 1: image4(Num) = 89: Num = Num + 1: image4(Num) = 50
Num = Num + 1: image4(Num) = 78: Num = Num + 1: image4(Num) = 75
Num = Num + 1: image4(Num) = 85: Num = Num + 1: image4(Num) = 80
Num = Num + 1: image4(Num) = 85: Num = Num + 1: image4(Num) = 80
Num = Num + 1: image4(Num) = 103: Num = Num + 1: image4(Num) = 74
Num = Num + 1: image4(Num) = 88: Num = Num + 1: image4(Num) = 75
Num = Num + 1: image4(Num) = 85: Num = Num + 1: image4(Num) = 80
Num = Num + 1: image4(Num) = 85: Num = Num + 1: image4(Num) = 80
Num = Num + 1: image4(Num) = 103: Num = Num + 1: image4(Num) = 74
Num = Num + 1: image4(Num) = 88
'--------------- create the database for cel 5 ------------------
Num = 0
image5(Num) = 38
Num = Num + 1: image5(Num) = 18: Num = Num + 1: image5(Num) = 45
Num = Num + 1: image5(Num) = 25: Num = Num + 1: image5(Num) = 49
Num = Num + 1: image5(Num) = 28: Num = Num + 1: image5(Num) = 43
Num = Num + 1: image5(Num) = 30: Num = Num + 1: image5(Num) = 42
Num = Num + 1: image5(Num) = 26: Num = Num + 1: image5(Num) = 54
Num = Num + 1: image5(Num) = 28: Num = Num + 1: image5(Num) = 56
Num = Num + 1: image5(Num) = 64: Num = Num + 1: image5(Num) = 45
Num = Num + 1: image5(Num) = 66: Num = Num + 1: image5(Num) = 34
Num = Num + 1: image5(Num) = 64: Num = Num + 1: image5(Num) = 40
Num = Num + 1: image5(Num) = 31: Num = Num + 1: image5(Num) = 27
Num = Num + 1: image5(Num) = 58: Num = Num + 1: image5(Num) = 26
Num = Num + 1: image5(Num) = 63: Num = Num + 1: image5(Num) = 18
Num = Num + 1: image5(Num) = 58: Num = Num + 1: image5(Num) = 21
Num = Num + 1: image5(Num) = 55: Num = Num + 1: image5(Num) = 28
Num = Num + 1: image5(Num) = 58: Num = Num + 1: image5(Num) = 40
Num = Num + 1: image5(Num) = 56: Num = Num + 1: image5(Num) = 40
Num = Num + 1: image5(Num) = 60: Num = Num + 1: image5(Num) = 25
Num = Num + 1: image5(Num) = 63: Num = Num + 1: image5(Num) = 28
Num = Num + 1: image5(Num) = 58: Num = Num + 1: image5(Num) = 40
Num = Num + 1: image5(Num) = 56: Num = Num + 1: image5(Num) = 40
Num = Num + 1: image5(Num) = 60: Num = Num + 1: image5(Num) = 25
Num = Num + 1: image5(Num) = 63: Num = Num + 1: image5(Num) = 54
Num = Num + 1: image5(Num) = 28: Num = Num + 1: image5(Num) = 57
Num = Num + 1: image5(Num) = 48: Num = Num + 1: image5(Num) = 60
Num = Num + 1: image5(Num) = 46: Num = Num + 1: image5(Num) = 54
Num = Num + 1: image5(Num) = 38: Num = Num + 1: image5(Num) = 60
Num = Num + 1: image5(Num) = 48: Num = Num + 1: image5(Num) = 66
Num = Num + 1: image5(Num) = 48: Num = Num + 1: image5(Num) = 52
Num = Num + 1: image5(Num) = 53: Num = Num + 1: image5(Num) = 50
Num = Num + 1: image5(Num) = 50: Num = Num + 1: image5(Num) = 60
Num = Num + 1: image5(Num) = 48: Num = Num + 1: image5(Num) = 66
Num = Num + 1: image5(Num) = 48: Num = Num + 1: image5(Num) = 52
Num = Num + 1: image5(Num) = 53: Num = Num + 1: image5(Num) = 50
Num = Num + 1: image5(Num) = 50: Num = Num + 1: image5(Num) = 34
Num = Num + 1: image5(Num) = 65: Num = Num + 1: image5(Num) = 40
Num = Num + 1: image5(Num) = 66: Num = Num + 1: image5(Num) = 14
Num = Num + 1: image5(Num) = 76: Num = Num + 1: image5(Num) = 12
Num = Num + 1: image5(Num) = 74: Num = Num + 1: image5(Num) = 22
Num = Num + 1: image5(Num) = 74: Num = Num + 1: image5(Num) = 40
Num = Num + 1: image5(Num) = 78: Num = Num + 1: image5(Num) = 38
Num = Num + 1: image5(Num) = 82: Num = Num + 1: image5(Num) = 12
Num = Num + 1: image5(Num) = 76: Num = Num + 1: image5(Num) = 40
Num = Num + 1: image5(Num) = 78: Num = Num + 1: image5(Num) = 46
Num = Num + 1: image5(Num) = 76: Num = Num + 1: image5(Num) = 45
Num = Num + 1: image5(Num) = 96: Num = Num + 1: image5(Num) = 38
Num = Num + 1: image5(Num) = 82: Num = Num + 1: image5(Num) = 40
Num = Num + 1: image5(Num) = 68: Num = Num + 1: image5(Num) = 45
Num = Num + 1: image5(Num) = 66: Num = Num + 1: image5(Num) = 27
Num = Num + 1: image5(Num) = 85: Num = Num + 1: image5(Num) = 20
Num = Num + 1: image5(Num) = 84: Num = Num + 1: image5(Num) = 27
Num = Num + 1: image5(Num) = 85: Num = Num + 1: image5(Num) = 33
Num = Num + 1: image5(Num) = 114: Num = Num + 1: image5(Num) = 25
Num = Num + 1: image5(Num) = 107: Num = Num + 1: image5(Num) = 20
Num = Num + 1: image5(Num) = 84: Num = Num + 1: image5(Num) = 25
Num = Num + 1: image5(Num) = 107: Num = Num + 1: image5(Num) = 31
Num = Num + 1: image5(Num) = 107: Num = Num + 1: image5(Num) = 32
Num = Num + 1: image5(Num) = 115: Num = Num + 1: image5(Num) = 5
Num = Num + 1: image5(Num) = 115: Num = Num + 1: image5(Num) = 25
Num = Num + 1: image5(Num) = 107: Num = Num + 1: image5(Num) = 31
Num = Num + 1: image5(Num) = 107: Num = Num + 1: image5(Num) = 32
Num = Num + 1: image5(Num) = 115: Num = Num + 1: image5(Num) = 5
Num = Num + 1: image5(Num) = 115
'--------------- create the database for cel 6 ------------------
Num = 0
image6(Num) = 38
Num = Num + 1: image6(Num) = 22: Num = Num + 1: image6(Num) = 41
Num = Num + 1: image6(Num) = 29: Num = Num + 1: image6(Num) = 45
Num = Num + 1: image6(Num) = 35: Num = Num + 1: image6(Num) = 39
Num = Num + 1: image6(Num) = 35: Num = Num + 1: image6(Num) = 39
Num = Num + 1: image6(Num) = 30: Num = Num + 1: image6(Num) = 38
Num = Num + 1: image6(Num) = 36: Num = Num + 1: image6(Num) = 49
Num = Num + 1: image6(Num) = 33: Num = Num + 1: image6(Num) = 60
Num = Num + 1: image6(Num) = 68: Num = Num + 1: image6(Num) = 52
Num = Num + 1: image6(Num) = 70: Num = Num + 1: image6(Num) = 39
Num = Num + 1: image6(Num) = 70: Num = Num + 1: image6(Num) = 12
Num = Num + 1: image6(Num) = 48: Num = Num + 1: image6(Num) = 16
Num = Num + 1: image6(Num) = 54: Num = Num + 1: image6(Num) = 11
Num = Num + 1: image6(Num) = 59: Num = Num + 1: image6(Num) = 5
Num = Num + 1: image6(Num) = 51: Num = Num + 1: image6(Num) = 15
Num = Num + 1: image6(Num) = 52: Num = Num + 1: image6(Num) = 34
Num = Num + 1: image6(Num) = 55: Num = Num + 1: image6(Num) = 38
Num = Num + 1: image6(Num) = 59: Num = Num + 1: image6(Num) = 15
Num = Num + 1: image6(Num) = 56: Num = Num + 1: image6(Num) = 33
Num = Num + 1: image6(Num) = 55: Num = Num + 1: image6(Num) = 40
Num = Num + 1: image6(Num) = 40: Num = Num + 1: image6(Num) = 40
Num = Num + 1: image6(Num) = 50: Num = Num + 1: image6(Num) = 38
Num = Num + 1: image6(Num) = 60: Num = Num + 1: image6(Num) = 50
Num = Num + 1: image6(Num) = 38: Num = Num + 1: image6(Num) = 65
Num = Num + 1: image6(Num) = 50: Num = Num + 1: image6(Num) = 60
Num = Num + 1: image6(Num) = 50: Num = Num + 1: image6(Num) = 50
Num = Num + 1: image6(Num) = 43: Num = Num + 1: image6(Num) = 60
Num = Num + 1: image6(Num) = 50: Num = Num + 1: image6(Num) = 65
Num = Num + 1: image6(Num) = 50: Num = Num + 1: image6(Num) = 60
Num = Num + 1: image6(Num) = 67: Num = Num + 1: image6(Num) = 58
Num = Num + 1: image6(Num) = 63: Num = Num + 1: image6(Num) = 60
Num = Num + 1: image6(Num) = 50: Num = Num + 1: image6(Num) = 65
Num = Num + 1: image6(Num) = 50: Num = Num + 1: image6(Num) = 60
Num = Num + 1: image6(Num) = 67: Num = Num + 1: image6(Num) = 58
Num = Num + 1: image6(Num) = 63: Num = Num + 1: image6(Num) = 38
Num = Num + 1: image6(Num) = 66: Num = Num + 1: image6(Num) = 39
Num = Num + 1: image6(Num) = 70: Num = Num + 1: image6(Num) = 16
Num = Num + 1: image6(Num) = 80: Num = Num + 1: image6(Num) = 15
Num = Num + 1: image6(Num) = 76: Num = Num + 1: image6(Num) = 22
Num = Num + 1: image6(Num) = 77: Num = Num + 1: image6(Num) = 49
Num = Num + 1: image6(Num) = 85: Num = Num + 1: image6(Num) = 36
Num = Num + 1: image6(Num) = 85: Num = Num + 1: image6(Num) = 16
Num = Num + 1: image6(Num) = 80: Num = Num + 1: image6(Num) = 40
Num = Num + 1: image6(Num) = 83: Num = Num + 1: image6(Num) = 49
Num = Num + 1: image6(Num) = 86: Num = Num + 1: image6(Num) = 32
Num = Num + 1: image6(Num) = 102: Num = Num + 1: image6(Num) = 36
Num = Num + 1: image6(Num) = 86: Num = Num + 1: image6(Num) = 45
Num = Num + 1: image6(Num) = 72: Num = Num + 1: image6(Num) = 52
Num = Num + 1: image6(Num) = 70: Num = Num + 1: image6(Num) = 43
Num = Num + 1: image6(Num) = 91: Num = Num + 1: image6(Num) = 38
Num = Num + 1: image6(Num) = 90: Num = Num + 1: image6(Num) = 41
Num = Num + 1: image6(Num) = 89: Num = Num + 1: image6(Num) = 62
Num = Num + 1: image6(Num) = 104: Num = Num + 1: image6(Num) = 54
Num = Num + 1: image6(Num) = 103: Num = Num + 1: image6(Num) = 38
Num = Num + 1: image6(Num) = 92: Num = Num + 1: image6(Num) = 53
Num = Num + 1: image6(Num) = 104: Num = Num + 1: image6(Num) = 63
Num = Num + 1: image6(Num) = 104: Num = Num + 1: image6(Num) = 54
Num = Num + 1: image6(Num) = 115: Num = Num + 1: image6(Num) = 49
Num = Num + 1: image6(Num) = 112: Num = Num + 1: image6(Num) = 52
Num = Num + 1: image6(Num) = 112: Num = Num + 1: image6(Num) = 53
Num = Num + 1: image6(Num) = 115: Num = Num + 1: image6(Num) = 43
Num = Num + 1: image6(Num) = 115: Num = Num + 1: image6(Num) = 49
Num = Num + 1: image6(Num) = 112
'--------------- create the database for cel 7 ------------------
Num = 0
image7(Num) = 38
Num = Num + 1: image7(Num) = 19: Num = Num + 1: image7(Num) = 45
Num = Num + 1: image7(Num) = 25: Num = Num + 1: image7(Num) = 48
Num = Num + 1: image7(Num) = 30: Num = Num + 1: image7(Num) = 43
Num = Num + 1: image7(Num) = 31: Num = Num + 1: image7(Num) = 42
Num = Num + 1: image7(Num) = 26: Num = Num + 1: image7(Num) = 51
Num = Num + 1: image7(Num) = 30: Num = Num + 1: image7(Num) = 69
Num = Num + 1: image7(Num) = 62: Num = Num + 1: image7(Num) = 60
Num = Num + 1: image7(Num) = 65: Num = Num + 1: image7(Num) = 49
Num = Num + 1: image7(Num) = 61: Num = Num + 1: image7(Num) = 40
Num = Num + 1: image7(Num) = 32: Num = Num + 1: image7(Num) = 14
Num = Num + 1: image7(Num) = 33: Num = Num + 1: image7(Num) = 15
Num = Num + 1: image7(Num) = 41: Num = Num + 1: image7(Num) = 6
Num = Num + 1: image7(Num) = 42: Num = Num + 1: image7(Num) = 4
Num = Num + 1: image7(Num) = 34: Num = Num + 1: image7(Num) = 15
Num = Num + 1: image7(Num) = 38: Num = Num + 1: image7(Num) = 30
Num = Num + 1: image7(Num) = 45: Num = Num + 1: image7(Num) = 30
Num = Num + 1: image7(Num) = 50: Num = Num + 1: image7(Num) = 15
Num = Num + 1: image7(Num) = 41: Num = Num + 1: image7(Num) = 30
Num = Num + 1: image7(Num) = 45: Num = Num + 1: image7(Num) = 41
Num = Num + 1: image7(Num) = 38: Num = Num + 1: image7(Num) = 42
Num = Num + 1: image7(Num) = 42: Num = Num + 1: image7(Num) = 30
Num = Num + 1: image7(Num) = 50: Num = Num + 1: image7(Num) = 53
Num = Num + 1: image7(Num) = 32: Num = Num + 1: image7(Num) = 73
Num = Num + 1: image7(Num) = 36: Num = Num + 1: image7(Num) = 71
Num = Num + 1: image7(Num) = 39: Num = Num + 1: image7(Num) = 55
Num = Num + 1: image7(Num) = 36: Num = Num + 1: image7(Num) = 73
Num = Num + 1: image7(Num) = 37: Num = Num + 1: image7(Num) = 83
Num = Num + 1: image7(Num) = 53: Num = Num + 1: image7(Num) = 80
Num = Num + 1: image7(Num) = 55: Num = Num + 1: image7(Num) = 71
Num = Num + 1: image7(Num) = 40: Num = Num + 1: image7(Num) = 83
Num = Num + 1: image7(Num) = 53: Num = Num + 1: image7(Num) = 93
Num = Num + 1: image7(Num) = 55: Num = Num + 1: image7(Num) = 89
Num = Num + 1: image7(Num) = 57: Num = Num + 1: image7(Num) = 80
Num = Num + 1: image7(Num) = 55: Num = Num + 1: image7(Num) = 49
Num = Num + 1: image7(Num) = 61: Num = Num + 1: image7(Num) = 60
Num = Num + 1: image7(Num) = 65: Num = Num + 1: image7(Num) = 22
Num = Num + 1: image7(Num) = 65: Num = Num + 1: image7(Num) = 23
Num = Num + 1: image7(Num) = 61: Num = Num + 1: image7(Num) = 31
Num = Num + 1: image7(Num) = 65: Num = Num + 1: image7(Num) = 53
Num = Num + 1: image7(Num) = 81: Num = Num + 1: image7(Num) = 40
Num = Num + 1: image7(Num) = 80: Num = Num + 1: image7(Num) = 23
Num = Num + 1: image7(Num) = 65: Num = Num + 1: image7(Num) = 48
Num = Num + 1: image7(Num) = 77: Num = Num + 1: image7(Num) = 53
Num = Num + 1: image7(Num) = 81: Num = Num + 1: image7(Num) = 27
Num = Num + 1: image7(Num) = 89: Num = Num + 1: image7(Num) = 40
Num = Num + 1: image7(Num) = 79: Num = Num + 1: image7(Num) = 61
Num = Num + 1: image7(Num) = 65: Num = Num + 1: image7(Num) = 67
Num = Num + 1: image7(Num) = 63: Num = Num + 1: image7(Num) = 68
Num = Num + 1: image7(Num) = 84: Num = Num + 1: image7(Num) = 63
Num = Num + 1: image7(Num) = 85: Num = Num + 1: image7(Num) = 68
Num = Num + 1: image7(Num) = 83: Num = Num + 1: image7(Num) = 84
Num = Num + 1: image7(Num) = 99: Num = Num + 1: image7(Num) = 80
Num = Num + 1: image7(Num) = 101: Num = Num + 1: image7(Num) = 63
Num = Num + 1: image7(Num) = 85: Num = Num + 1: image7(Num) = 81
Num = Num + 1: image7(Num) = 98: Num = Num + 1: image7(Num) = 89
Num = Num + 1: image7(Num) = 98: Num = Num + 1: image7(Num) = 85
Num = Num + 1: image7(Num) = 113: Num = Num + 1: image7(Num) = 80
Num = Num + 1: image7(Num) = 111: Num = Num + 1: image7(Num) = 89
Num = Num + 1: image7(Num) = 111: Num = Num + 1: image7(Num) = 85
Num = Num + 1: image7(Num) = 112: Num = Num + 1: image7(Num) = 80
Num = Num + 1: image7(Num) = 115: Num = Num + 1: image7(Num) = 76
Num = Num + 1: image7(Num) = 115
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

