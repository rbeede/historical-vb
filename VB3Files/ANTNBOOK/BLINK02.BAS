'----------------------------------------------------------------
'     Frame animation routines for Visual Basic applications
'----------------------------------------------------------------
' Source file:  BLINK02.BAS
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
  Case 2:  RetVal = SelectObject(hFDC, hF1)
  Case 3:  RetVal = SelectObject(hFDC, hF1)
  Case 4:  RetVal = SelectObject(hFDC, hF1)
  Case 5:  RetVal = SelectObject(hFDC, hF1)
  Case 6:  RetVal = SelectObject(hFDC, hF1)
  Case 7:  RetVal = SelectObject(hFDC, hF1)
  Case 8:  RetVal = SelectObject(hFDC, hF1)
  Case 9:  RetVal = SelectObject(hFDC, hF1)
  Case 10:  RetVal = SelectObject(hFDC, hF1)
  Case 11:  RetVal = SelectObject(hFDC, hF1)
  Case 12:  RetVal = SelectObject(hFDC, hF1)
  Case 13:  RetVal = SelectObject(hFDC, hF1)
  Case 14:  RetVal = SelectObject(hFDC, hF1)
  Case 15:  RetVal = SelectObject(hFDC, hF1)
  Case 16:  RetVal = SelectObject(hFDC, hF1)
  Case 17:  RetVal = SelectObject(hFDC, hF1)
  Case 18:  RetVal = SelectObject(hFDC, hF2)
  Case 19:  RetVal = SelectObject(hFDC, hF3)
  Case 20:  RetVal = SelectObject(hFDC, hF3)
  Case 21:  RetVal = SelectObject(hFDC, hF3)
  Case 22:  RetVal = SelectObject(hFDC, hF3)
  Case 23:  RetVal = SelectObject(hFDC, hF4)
  Case 24:  RetVal = SelectObject(hFDC, hF4)
  Case 25:  RetVal = SelectObject(hFDC, hF4)
  Case 26:  RetVal = SelectObject(hFDC, hF4)
  Case 27:  RetVal = SelectObject(hFDC, hF4)
  Case 28:  RetVal = SelectObject(hFDC, hF4)
  Case 29:  RetVal = SelectObject(hFDC, hF5)
  Case 30:  RetVal = SelectObject(hFDC, hF6)
  Case 31:  RetVal = SelectObject(hFDC, hF7)
  Case 32:  RetVal = SelectObject(hFDC, hF8)
  Case 33:  RetVal = SelectObject(hFDC, hF9)
  Case 34:  RetVal = SelectObject(hFDC, hF10)
  Case 35:  RetVal = SelectObject(hFDC, hF10)
  Case 36:  RetVal = SelectObject(hFDC, hF10)
  Case 37:  RetVal = SelectObject(hFDC, hF10)
  Case 38:  RetVal = SelectObject(hFDC, hF10)
  Case 39:  RetVal = SelectObject(hFDC, hF10)
  Case 40:  RetVal = SelectObject(hFDC, hF10)
  Case 41:  RetVal = SelectObject(hFDC, hF10)
  Case 42:  RetVal = SelectObject(hFDC, hF10)
  Case 43:  RetVal = SelectObject(hFDC, hF10)
  Case 44:  RetVal = SelectObject(hFDC, hF10)
  Case 45:  RetVal = SelectObject(hFDC, hF11)
  Case 46:  RetVal = SelectObject(hFDC, hF11)
  Case 47:  RetVal = SelectObject(hFDC, hF12)
  Case 48:  RetVal = SelectObject(hFDC, hF12)
  Case 49:  RetVal = SelectObject(hFDC, hF13)
  Case 50:  RetVal = SelectObject(hFDC, hF13)
  Case 51:  RetVal = SelectObject(hFDC, hF13)
  Case 52:  RetVal = SelectObject(hFDC, hF13)
  Case 53:  RetVal = SelectObject(hFDC, hF13)
  Case 54:  RetVal = SelectObject(hFDC, hF13)
  Case 55:  RetVal = SelectObject(hFDC, hF13)
  Case 56:  RetVal = SelectObject(hFDC, hF13)
  Case 57:  RetVal = SelectObject(hFDC, hF13)
  Case 58:  RetVal = SelectObject(hFDC, hF13)
  Case 59:  RetVal = SelectObject(hFDC, hF13)
  Case 60:  RetVal = SelectObject(hFDC, hF13)
  Case 61:  RetVal = SelectObject(hFDC, hF13)
  Case 62:  RetVal = SelectObject(hFDC, hF13)
  Case 63:  RetVal = SelectObject(hFDC, hF13)
  Case 64:  RetVal = SelectObject(hFDC, hF13)
  Case 65:  RetVal = SelectObject(hFDC, hF13)
  Case 66:  RetVal = SelectObject(hFDC, hF14)
  Case 67:  RetVal = SelectObject(hFDC, hF15)
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
  Case 1:  FileName = "BLINK1.BMP": zLoadFrame
  Case 2:  FileName = "BLINK2.BMP": zLoadFrame
  Case 3:  FileName = "BLINK3.BMP": zLoadFrame
  Case 4:  FileName = "BLINK4.BMP": zLoadFrame
  Case 5:  FileName = "BLINK5.BMP": zLoadFrame
  Case 6:  FileName = "BLINK6.BMP": zLoadFrame
  Case 7:  FileName = "BLINK7.BMP": zLoadFrame
  Case 8:  FileName = "BLINK8.BMP": zLoadFrame
  Case 9:  FileName = "BLINK9.BMP": zLoadFrame
  Case 10:  FileName = "BLINK10.BMP": zLoadFrame
  Case 11:  FileName = "BLINK11.BMP": zLoadFrame
  Case 12:  FileName = "BLINK12.BMP": zLoadFrame
  Case 13:  FileName = "BLINK13.BMP": zLoadFrame
  Case 14:  FileName = "BLINK14.BMP": zLoadFrame
  Case 15:  FileName = "BLINK15.BMP": zLoadFrame
End Select
If FrameLoaded = False Then
  AnimationHalted = True
  Beep
  MsgBox "Unable to load next frame from disk.  Animation halted.", MB_OK, "Unexpected animation condition"
  Exit Sub
End If
End Sub

'----------------------------------------------------------------
'        Initialize the database of drawing coordinates
'----------------------------------------------------------------
Sub zInitDatabase ()
'   Note:  each array also contains drawing codes that give
'   meaning to the coordinates in the database.  For example,
'   an initial value of 1 indicates the start of a new line.  The
'   next element in the array contains the number of endpoints
'   in the line, followed by the coordinates for the startpoint
'   and each subsequent vertex.  An initial value of 2 indicates
'   an ellipse, followed by the diagonal bounding coordinates.
'   A value of 0 indicates the end of the array.
'--------------------- declare local variable -------------------
  Dim Num As Integer                            'index into array
'--------------------- database for cel 1 -----------------------
Num = 0                        'reset index to beginning of array
Cel1(Num) = 1                         'polyline (with 1 endpoint)
Num = Num + 1: Cel1(Num) = 1
Num = Num + 1: Cel1(Num) = 175
Num = Num + 1: Cel1(Num) = 164
Num = Num + 1: Cel1(Num) = 184
Num = Num + 1: Cel1(Num) = 157
Num = Num + 1: Cel1(Num) = 1          'polyline (with 1 endpoint)
Num = Num + 1: Cel1(Num) = 1
Num = Num + 1: Cel1(Num) = 214
Num = Num + 1: Cel1(Num) = 157
Num = Num + 1: Cel1(Num) = 220
Num = Num + 1: Cel1(Num) = 163
Num = Num + 1: Cel1(Num) = 1         'polyline (with 6 endpoints)
Num = Num + 1: Cel1(Num) = 6
Num = Num + 1: Cel1(Num) = 181
Num = Num + 1: Cel1(Num) = 161
Num = Num + 1: Cel1(Num) = 187
Num = Num + 1: Cel1(Num) = 170
Num = Num + 1: Cel1(Num) = 194
Num = Num + 1: Cel1(Num) = 175
Num = Num + 1: Cel1(Num) = 200
Num = Num + 1: Cel1(Num) = 176
Num = Num + 1: Cel1(Num) = 206
Num = Num + 1: Cel1(Num) = 174
Num = Num + 1: Cel1(Num) = 212
Num = Num + 1: Cel1(Num) = 168
Num = Num + 1: Cel1(Num) = 216
Num = Num + 1: Cel1(Num) = 163
Num = Num + 1: Cel1(Num) = 2                             'ellipse
Num = Num + 1: Cel1(Num) = 188
Num = Num + 1: Cel1(Num) = 107
Num = Num + 1: Cel1(Num) = 195
Num = Num + 1: Cel1(Num) = 116
Num = Num + 1: Cel1(Num) = 2                             'ellipse
Num = Num + 1: Cel1(Num) = 204
Num = Num + 1: Cel1(Num) = 107
Num = Num + 1: Cel1(Num) = 211
Num = Num + 1: Cel1(Num) = 116
Num = Num + 1: Cel1(Num) = 1        'polyline (with 23 endpoints)
Num = Num + 1: Cel1(Num) = 23
Num = Num + 1: Cel1(Num) = 200
Num = Num + 1: Cel1(Num) = 104
Num = Num + 1: Cel1(Num) = 204
Num = Num + 1: Cel1(Num) = 100
Num = Num + 1: Cel1(Num) = 210
Num = Num + 1: Cel1(Num) = 99
Num = Num + 1: Cel1(Num) = 215
Num = Num + 1: Cel1(Num) = 100
Num = Num + 1: Cel1(Num) = 217
Num = Num + 1: Cel1(Num) = 105
Num = Num + 1: Cel1(Num) = 218
Num = Num + 1: Cel1(Num) = 110
Num = Num + 1: Cel1(Num) = 216
Num = Num + 1: Cel1(Num) = 118
Num = Num + 1: Cel1(Num) = 215
Num = Num + 1: Cel1(Num) = 120
Num = Num + 1: Cel1(Num) = 210
Num = Num + 1: Cel1(Num) = 122
Num = Num + 1: Cel1(Num) = 205
Num = Num + 1: Cel1(Num) = 121
Num = Num + 1: Cel1(Num) = 202
Num = Num + 1: Cel1(Num) = 120
Num = Num + 1: Cel1(Num) = 200
Num = Num + 1: Cel1(Num) = 118
Num = Num + 1: Cel1(Num) = 198
Num = Num + 1: Cel1(Num) = 120
Num = Num + 1: Cel1(Num) = 193
Num = Num + 1: Cel1(Num) = 122
Num = Num + 1: Cel1(Num) = 188
Num = Num + 1: Cel1(Num) = 122
Num = Num + 1: Cel1(Num) = 185
Num = Num + 1: Cel1(Num) = 120
Num = Num + 1: Cel1(Num) = 183
Num = Num + 1: Cel1(Num) = 115
Num = Num + 1: Cel1(Num) = 182
Num = Num + 1: Cel1(Num) = 110
Num = Num + 1: Cel1(Num) = 183
Num = Num + 1: Cel1(Num) = 105
Num = Num + 1: Cel1(Num) = 186
Num = Num + 1: Cel1(Num) = 100
Num = Num + 1: Cel1(Num) = 190
Num = Num + 1: Cel1(Num) = 99
Num = Num + 1: Cel1(Num) = 194
Num = Num + 1: Cel1(Num) = 100
Num = Num + 1: Cel1(Num) = 198
Num = Num + 1: Cel1(Num) = 101
Num = Num + 1: Cel1(Num) = 200
Num = Num + 1: Cel1(Num) = 104
Num = Num + 1: Cel1(Num) = 0
'--------------------- database for cel 2 -----------------------
Num = 0                        'reset index to beginning of array
Cel2(Num) = 1
Num = Num + 1: Cel2(Num) = 1
Num = Num + 1: Cel2(Num) = 178
Num = Num + 1: Cel2(Num) = 168
Num = Num + 1: Cel2(Num) = 183
Num = Num + 1: Cel2(Num) = 162
Num = Num + 1: Cel2(Num) = 1
Num = Num + 1: Cel2(Num) = 6
Num = Num + 1: Cel2(Num) = 180
Num = Num + 1: Cel2(Num) = 165
Num = Num + 1: Cel2(Num) = 190
Num = Num + 1: Cel2(Num) = 170
Num = Num + 1: Cel2(Num) = 195
Num = Num + 1: Cel2(Num) = 171
Num = Num + 1: Cel2(Num) = 200
Num = Num + 1: Cel2(Num) = 171
Num = Num + 1: Cel2(Num) = 206
Num = Num + 1: Cel2(Num) = 170
Num = Num + 1: Cel2(Num) = 212
Num = Num + 1: Cel2(Num) = 167
Num = Num + 1: Cel2(Num) = 215
Num = Num + 1: Cel2(Num) = 164
Num = Num + 1: Cel2(Num) = 1
Num = Num + 1: Cel2(Num) = 1
Num = Num + 1: Cel2(Num) = 215
Num = Num + 1: Cel2(Num) = 160
Num = Num + 1: Cel2(Num) = 220
Num = Num + 1: Cel2(Num) = 168
Num = Num + 1: Cel2(Num) = 2
Num = Num + 1: Cel2(Num) = 188
Num = Num + 1: Cel2(Num) = 107
Num = Num + 1: Cel2(Num) = 195
Num = Num + 1: Cel2(Num) = 116
Num = Num + 1: Cel2(Num) = 2
Num = Num + 1: Cel2(Num) = 204
Num = Num + 1: Cel2(Num) = 107
Num = Num + 1: Cel2(Num) = 211
Num = Num + 1: Cel2(Num) = 116
Num = Num + 1: Cel2(Num) = 1
Num = Num + 1: Cel2(Num) = 1
Num = Num + 1: Cel2(Num) = 195
Num = Num + 1: Cel2(Num) = 96
Num = Num + 1: Cel2(Num) = 199
Num = Num + 1: Cel2(Num) = 105
Num = Num + 1: Cel2(Num) = 1
Num = Num + 1: Cel2(Num) = 1
Num = Num + 1: Cel2(Num) = 205
Num = Num + 1: Cel2(Num) = 97
Num = Num + 1: Cel2(Num) = 201
Num = Num + 1: Cel2(Num) = 104
Num = Num + 1: Cel2(Num) = 1
Num = Num + 1: Cel2(Num) = 10
Num = Num + 1: Cel2(Num) = 200
Num = Num + 1: Cel2(Num) = 110
Num = Num + 1: Cel2(Num) = 205
Num = Num + 1: Cel2(Num) = 108
Num = Num + 1: Cel2(Num) = 210
Num = Num + 1: Cel2(Num) = 107
Num = Num + 1: Cel2(Num) = 215
Num = Num + 1: Cel2(Num) = 108
Num = Num + 1: Cel2(Num) = 219
Num = Num + 1: Cel2(Num) = 112
Num = Num + 1: Cel2(Num) = 215
Num = Num + 1: Cel2(Num) = 120
Num = Num + 1: Cel2(Num) = 212
Num = Num + 1: Cel2(Num) = 122
Num = Num + 1: Cel2(Num) = 207
Num = Num + 1: Cel2(Num) = 122
Num = Num + 1: Cel2(Num) = 201
Num = Num + 1: Cel2(Num) = 121
Num = Num + 1: Cel2(Num) = 200
Num = Num + 1: Cel2(Num) = 118
Num = Num + 1: Cel2(Num) = 200
Num = Num + 1: Cel2(Num) = 109
Num = Num + 1: Cel2(Num) = 1
Num = Num + 1: Cel2(Num) = 7
Num = Num + 1: Cel2(Num) = 200
Num = Num + 1: Cel2(Num) = 117
Num = Num + 1: Cel2(Num) = 195
Num = Num + 1: Cel2(Num) = 121
Num = Num + 1: Cel2(Num) = 190
Num = Num + 1: Cel2(Num) = 122
Num = Num + 1: Cel2(Num) = 185
Num = Num + 1: Cel2(Num) = 120
Num = Num + 1: Cel2(Num) = 180
Num = Num + 1: Cel2(Num) = 114
Num = Num + 1: Cel2(Num) = 186
Num = Num + 1: Cel2(Num) = 108
Num = Num + 1: Cel2(Num) = 192
Num = Num + 1: Cel2(Num) = 108
Num = Num + 1: Cel2(Num) = 200
Num = Num + 1: Cel2(Num) = 109
Num = Num + 1: Cel2(Num) = 0
'--------------------- database for cel 3 -----------------------
Num = 0
Cel3(Num) = 1
Num = Num + 1: Cel3(Num) = 1
Num = Num + 1: Cel3(Num) = 179
Num = Num + 1: Cel3(Num) = 166
Num = Num + 1: Cel3(Num) = 184
Num = Num + 1: Cel3(Num) = 161
Num = Num + 1: Cel3(Num) = 1
Num = Num + 1: Cel3(Num) = 1
Num = Num + 1: Cel3(Num) = 215
Num = Num + 1: Cel3(Num) = 159
Num = Num + 1: Cel3(Num) = 218
Num = Num + 1: Cel3(Num) = 166
Num = Num + 1: Cel3(Num) = 1
Num = Num + 1: Cel3(Num) = 4
Num = Num + 1: Cel3(Num) = 181
Num = Num + 1: Cel3(Num) = 164
Num = Num + 1: Cel3(Num) = 190
Num = Num + 1: Cel3(Num) = 167
Num = Num + 1: Cel3(Num) = 200
Num = Num + 1: Cel3(Num) = 167
Num = Num + 1: Cel3(Num) = 210
Num = Num + 1: Cel3(Num) = 165
Num = Num + 1: Cel3(Num) = 216
Num = Num + 1: Cel3(Num) = 163
Num = Num + 1: Cel3(Num) = 1
Num = Num + 1: Cel3(Num) = 2
Num = Num + 1: Cel3(Num) = 190
Num = Num + 1: Cel3(Num) = 99
Num = Num + 1: Cel3(Num) = 196
Num = Num + 1: Cel3(Num) = 102
Num = Num + 1: Cel3(Num) = 199
Num = Num + 1: Cel3(Num) = 110
Num = Num + 1: Cel3(Num) = 1
Num = Num + 1: Cel3(Num) = 2
Num = Num + 1: Cel3(Num) = 211
Num = Num + 1: Cel3(Num) = 97
Num = Num + 1: Cel3(Num) = 205
Num = Num + 1: Cel3(Num) = 103
Num = Num + 1: Cel3(Num) = 202
Num = Num + 1: Cel3(Num) = 109
Num = Num + 1: Cel3(Num) = 1
Num = Num + 1: Cel3(Num) = 3
Num = Num + 1: Cel3(Num) = 178
Num = Num + 1: Cel3(Num) = 115
Num = Num + 1: Cel3(Num) = 183
Num = Num + 1: Cel3(Num) = 111
Num = Num + 1: Cel3(Num) = 190
Num = Num + 1: Cel3(Num) = 110
Num = Num + 1: Cel3(Num) = 198
Num = Num + 1: Cel3(Num) = 113
Num = Num + 1: Cel3(Num) = 1
Num = Num + 1: Cel3(Num) = 3
Num = Num + 1: Cel3(Num) = 201
Num = Num + 1: Cel3(Num) = 113
Num = Num + 1: Cel3(Num) = 210
Num = Num + 1: Cel3(Num) = 110
Num = Num + 1: Cel3(Num) = 216
Num = Num + 1: Cel3(Num) = 110
Num = Num + 1: Cel3(Num) = 221
Num = Num + 1: Cel3(Num) = 114
Num = Num + 1: Cel3(Num) = 1
Num = Num + 1: Cel3(Num) = 3
Num = Num + 1: Cel3(Num) = 187
Num = Num + 1: Cel3(Num) = 121
Num = Num + 1: Cel3(Num) = 189
Num = Num + 1: Cel3(Num) = 119
Num = Num + 1: Cel3(Num) = 194
Num = Num + 1: Cel3(Num) = 115
Num = Num + 1: Cel3(Num) = 196
Num = Num + 1: Cel3(Num) = 115
Num = Num + 1: Cel3(Num) = 1
Num = Num + 1: Cel3(Num) = 3
Num = Num + 1: Cel3(Num) = 202
Num = Num + 1: Cel3(Num) = 115
Num = Num + 1: Cel3(Num) = 208
Num = Num + 1: Cel3(Num) = 115
Num = Num + 1: Cel3(Num) = 211
Num = Num + 1: Cel3(Num) = 118
Num = Num + 1: Cel3(Num) = 213
Num = Num + 1: Cel3(Num) = 121
Num = Num + 1: Cel3(Num) = 0
'--------------------- database for cel 4 -----------------------
Num = 0
Cel4(Num) = 2
Num = Num + 1: Cel4(Num) = 195
Num = Num + 1: Cel4(Num) = 155
Num = Num + 1: Cel4(Num) = 206
Num = Num + 1: Cel4(Num) = 190
Num = Num + 1: Cel4(Num) = 2
Num = Num + 1: Cel4(Num) = 194
Num = Num + 1: Cel4(Num) = 107
Num = Num + 1: Cel4(Num) = 200
Num = Num + 1: Cel4(Num) = 118
Num = Num + 1: Cel4(Num) = 2
Num = Num + 1: Cel4(Num) = 201
Num = Num + 1: Cel4(Num) = 107
Num = Num + 1: Cel4(Num) = 207
Num = Num + 1: Cel4(Num) = 118
Num = Num + 1: Cel4(Num) = 1
Num = Num + 1: Cel4(Num) = 18
Num = Num + 1: Cel4(Num) = 190
Num = Num + 1: Cel4(Num) = 121
Num = Num + 1: Cel4(Num) = 187
Num = Num + 1: Cel4(Num) = 121
Num = Num + 1: Cel4(Num) = 182
Num = Num + 1: Cel4(Num) = 116
Num = Num + 1: Cel4(Num) = 182
Num = Num + 1: Cel4(Num) = 105
Num = Num + 1: Cel4(Num) = 184
Num = Num + 1: Cel4(Num) = 96
Num = Num + 1: Cel4(Num) = 190
Num = Num + 1: Cel4(Num) = 91
Num = Num + 1: Cel4(Num) = 197
Num = Num + 1: Cel4(Num) = 91
Num = Num + 1: Cel4(Num) = 199
Num = Num + 1: Cel4(Num) = 92
Num = Num + 1: Cel4(Num) = 200
Num = Num + 1: Cel4(Num) = 95
Num = Num + 1: Cel4(Num) = 205
Num = Num + 1: Cel4(Num) = 92
Num = Num + 1: Cel4(Num) = 210
Num = Num + 1: Cel4(Num) = 92
Num = Num + 1: Cel4(Num) = 215
Num = Num + 1: Cel4(Num) = 95
Num = Num + 1: Cel4(Num) = 218
Num = Num + 1: Cel4(Num) = 100
Num = Num + 1: Cel4(Num) = 219
Num = Num + 1: Cel4(Num) = 105
Num = Num + 1: Cel4(Num) = 219
Num = Num + 1: Cel4(Num) = 110
Num = Num + 1: Cel4(Num) = 217
Num = Num + 1: Cel4(Num) = 116
Num = Num + 1: Cel4(Num) = 214
Num = Num + 1: Cel4(Num) = 121
Num = Num + 1: Cel4(Num) = 210
Num = Num + 1: Cel4(Num) = 122
Num = Num + 1: Cel4(Num) = 206
Num = Num + 1: Cel4(Num) = 121
Num = Num + 1: Cel4(Num) = 0
'--------------------- database for cel 5 -----------------------
Num = 0
Cel5(Num) = 2
Num = Num + 1: Cel5(Num) = 193
Num = Num + 1: Cel5(Num) = 160
Num = Num + 1: Cel5(Num) = 204
Num = Num + 1: Cel5(Num) = 180
Num = Num + 1: Cel5(Num) = 2
Num = Num + 1: Cel5(Num) = 194
Num = Num + 1: Cel5(Num) = 105
Num = Num + 1: Cel5(Num) = 200
Num = Num + 1: Cel5(Num) = 116
Num = Num + 1: Cel5(Num) = 2
Num = Num + 1: Cel5(Num) = 202
Num = Num + 1: Cel5(Num) = 105
Num = Num + 1: Cel5(Num) = 209
Num = Num + 1: Cel5(Num) = 116
Num = Num + 1: Cel5(Num) = 1
Num = Num + 1: Cel5(Num) = 16
Num = Num + 1: Cel5(Num) = 192
Num = Num + 1: Cel5(Num) = 120
Num = Num + 1: Cel5(Num) = 188
Num = Num + 1: Cel5(Num) = 120
Num = Num + 1: Cel5(Num) = 184
Num = Num + 1: Cel5(Num) = 116
Num = Num + 1: Cel5(Num) = 184
Num = Num + 1: Cel5(Num) = 110
Num = Num + 1: Cel5(Num) = 185
Num = Num + 1: Cel5(Num) = 100
Num = Num + 1: Cel5(Num) = 188
Num = Num + 1: Cel5(Num) = 96
Num = Num + 1: Cel5(Num) = 193
Num = Num + 1: Cel5(Num) = 95
Num = Num + 1: Cel5(Num) = 199
Num = Num + 1: Cel5(Num) = 97
Num = Num + 1: Cel5(Num) = 200
Num = Num + 1: Cel5(Num) = 100
Num = Num + 1: Cel5(Num) = 206
Num = Num + 1: Cel5(Num) = 96
Num = Num + 1: Cel5(Num) = 213
Num = Num + 1: Cel5(Num) = 96
Num = Num + 1: Cel5(Num) = 217
Num = Num + 1: Cel5(Num) = 98
Num = Num + 1: Cel5(Num) = 219
Num = Num + 1: Cel5(Num) = 105
Num = Num + 1: Cel5(Num) = 218
Num = Num + 1: Cel5(Num) = 114
Num = Num + 1: Cel5(Num) = 216
Num = Num + 1: Cel5(Num) = 120
Num = Num + 1: Cel5(Num) = 211
Num = Num + 1: Cel5(Num) = 122
Num = Num + 1: Cel5(Num) = 208
Num = Num + 1: Cel5(Num) = 121
Num = Num + 1: Cel5(Num) = 0
'--------------------- database for cel 6 -----------------------
Num = 0
Cel6(Num) = 1
Num = Num + 1: Cel6(Num) = 7
Num = Num + 1: Cel6(Num) = 184
Num = Num + 1: Cel6(Num) = 168
Num = Num + 1: Cel6(Num) = 190
Num = Num + 1: Cel6(Num) = 168
Num = Num + 1: Cel6(Num) = 213
Num = Num + 1: Cel6(Num) = 165
Num = Num + 1: Cel6(Num) = 208
Num = Num + 1: Cel6(Num) = 168
Num = Num + 1: Cel6(Num) = 203
Num = Num + 1: Cel6(Num) = 172
Num = Num + 1: Cel6(Num) = 197
Num = Num + 1: Cel6(Num) = 173
Num = Num + 1: Cel6(Num) = 192
Num = Num + 1: Cel6(Num) = 171
Num = Num + 1: Cel6(Num) = 184
Num = Num + 1: Cel6(Num) = 168
Num = Num + 1: Cel6(Num) = 2
Num = Num + 1: Cel6(Num) = 188
Num = Num + 1: Cel6(Num) = 107
Num = Num + 1: Cel6(Num) = 195
Num = Num + 1: Cel6(Num) = 116
Num = Num + 1: Cel6(Num) = 2
Num = Num + 1: Cel6(Num) = 204
Num = Num + 1: Cel6(Num) = 107
Num = Num + 1: Cel6(Num) = 211
Num = Num + 1: Cel6(Num) = 116
Num = Num + 1: Cel6(Num) = 1
Num = Num + 1: Cel6(Num) = 23
Num = Num + 1: Cel6(Num) = 200
Num = Num + 1: Cel6(Num) = 104
Num = Num + 1: Cel6(Num) = 204
Num = Num + 1: Cel6(Num) = 100
Num = Num + 1: Cel6(Num) = 210
Num = Num + 1: Cel6(Num) = 99
Num = Num + 1: Cel6(Num) = 215
Num = Num + 1: Cel6(Num) = 100
Num = Num + 1: Cel6(Num) = 217
Num = Num + 1: Cel6(Num) = 105
Num = Num + 1: Cel6(Num) = 218
Num = Num + 1: Cel6(Num) = 110
Num = Num + 1: Cel6(Num) = 216
Num = Num + 1: Cel6(Num) = 118
Num = Num + 1: Cel6(Num) = 215
Num = Num + 1: Cel6(Num) = 120
Num = Num + 1: Cel6(Num) = 210
Num = Num + 1: Cel6(Num) = 122
Num = Num + 1: Cel6(Num) = 205
Num = Num + 1: Cel6(Num) = 121
Num = Num + 1: Cel6(Num) = 202
Num = Num + 1: Cel6(Num) = 120
Num = Num + 1: Cel6(Num) = 200
Num = Num + 1: Cel6(Num) = 118
Num = Num + 1: Cel6(Num) = 198
Num = Num + 1: Cel6(Num) = 120
Num = Num + 1: Cel6(Num) = 193
Num = Num + 1: Cel6(Num) = 122
Num = Num + 1: Cel6(Num) = 188
Num = Num + 1: Cel6(Num) = 122
Num = Num + 1: Cel6(Num) = 185
Num = Num + 1: Cel6(Num) = 120
Num = Num + 1: Cel6(Num) = 183
Num = Num + 1: Cel6(Num) = 115
Num = Num + 1: Cel6(Num) = 182
Num = Num + 1: Cel6(Num) = 110
Num = Num + 1: Cel6(Num) = 183
Num = Num + 1: Cel6(Num) = 105
Num = Num + 1: Cel6(Num) = 186
Num = Num + 1: Cel6(Num) = 100
Num = Num + 1: Cel6(Num) = 190
Num = Num + 1: Cel6(Num) = 99
Num = Num + 1: Cel6(Num) = 194
Num = Num + 1: Cel6(Num) = 100
Num = Num + 1: Cel6(Num) = 198
Num = Num + 1: Cel6(Num) = 101
Num = Num + 1: Cel6(Num) = 200
Num = Num + 1: Cel6(Num) = 104
Num = Num + 1: Cel6(Num) = 0
'--------------------- database for cel 7 -----------------------
Num = 0
Cel7(Num) = 1
Num = Num + 1: Cel7(Num) = 9
Num = Num + 1: Cel7(Num) = 181
Num = Num + 1: Cel7(Num) = 165
Num = Num + 1: Cel7(Num) = 190
Num = Num + 1: Cel7(Num) = 166
Num = Num + 1: Cel7(Num) = 195
Num = Num + 1: Cel7(Num) = 166
Num = Num + 1: Cel7(Num) = 210
Num = Num + 1: Cel7(Num) = 164
Num = Num + 1: Cel7(Num) = 218
Num = Num + 1: Cel7(Num) = 164
Num = Num + 1: Cel7(Num) = 210
Num = Num + 1: Cel7(Num) = 166
Num = Num + 1: Cel7(Num) = 201
Num = Num + 1: Cel7(Num) = 169
Num = Num + 1: Cel7(Num) = 195
Num = Num + 1: Cel7(Num) = 169
Num = Num + 1: Cel7(Num) = 190
Num = Num + 1: Cel7(Num) = 167
Num = Num + 1: Cel7(Num) = 181
Num = Num + 1: Cel7(Num) = 165
Num = Num + 1: Cel7(Num) = 2
Num = Num + 1: Cel7(Num) = 188
Num = Num + 1: Cel7(Num) = 107
Num = Num + 1: Cel7(Num) = 195
Num = Num + 1: Cel7(Num) = 116
Num = Num + 1: Cel7(Num) = 2
Num = Num + 1: Cel7(Num) = 204
Num = Num + 1: Cel7(Num) = 107
Num = Num + 1: Cel7(Num) = 211
Num = Num + 1: Cel7(Num) = 116
Num = Num + 1: Cel7(Num) = 1
Num = Num + 1: Cel7(Num) = 17
Num = Num + 1: Cel7(Num) = 182
Num = Num + 1: Cel7(Num) = 102
Num = Num + 1: Cel7(Num) = 218
Num = Num + 1: Cel7(Num) = 102
Num = Num + 1: Cel7(Num) = 218
Num = Num + 1: Cel7(Num) = 110
Num = Num + 1: Cel7(Num) = 216
Num = Num + 1: Cel7(Num) = 115
Num = Num + 1: Cel7(Num) = 214
Num = Num + 1: Cel7(Num) = 120
Num = Num + 1: Cel7(Num) = 211
Num = Num + 1: Cel7(Num) = 122
Num = Num + 1: Cel7(Num) = 208
Num = Num + 1: Cel7(Num) = 122
Num = Num + 1: Cel7(Num) = 202
Num = Num + 1: Cel7(Num) = 120
Num = Num + 1: Cel7(Num) = 201
Num = Num + 1: Cel7(Num) = 119
Num = Num + 1: Cel7(Num) = 200
Num = Num + 1: Cel7(Num) = 115
Num = Num + 1: Cel7(Num) = 199
Num = Num + 1: Cel7(Num) = 119
Num = Num + 1: Cel7(Num) = 195
Num = Num + 1: Cel7(Num) = 121
Num = Num + 1: Cel7(Num) = 192
Num = Num + 1: Cel7(Num) = 122
Num = Num + 1: Cel7(Num) = 189
Num = Num + 1: Cel7(Num) = 122
Num = Num + 1: Cel7(Num) = 184
Num = Num + 1: Cel7(Num) = 118
Num = Num + 1: Cel7(Num) = 182
Num = Num + 1: Cel7(Num) = 113
Num = Num + 1: Cel7(Num) = 182
Num = Num + 1: Cel7(Num) = 107
Num = Num + 1: Cel7(Num) = 182
Num = Num + 1: Cel7(Num) = 102
Num = Num + 1: Cel7(Num) = 0
'--------------------- database for cel 8 -----------------------
Num = 0
Cel8(Num) = 1
Num = Num + 1: Cel8(Num) = 1
Num = Num + 1: Cel8(Num) = 175
Num = Num + 1: Cel8(Num) = 168
Num = Num + 1: Cel8(Num) = 179
Num = Num + 1: Cel8(Num) = 162
Num = Num + 1: Cel8(Num) = 1
Num = Num + 1: Cel8(Num) = 5
Num = Num + 1: Cel8(Num) = 178
Num = Num + 1: Cel8(Num) = 164
Num = Num + 1: Cel8(Num) = 186
Num = Num + 1: Cel8(Num) = 167
Num = Num + 1: Cel8(Num) = 191
Num = Num + 1: Cel8(Num) = 167
Num = Num + 1: Cel8(Num) = 200
Num = Num + 1: Cel8(Num) = 165
Num = Num + 1: Cel8(Num) = 210
Num = Num + 1: Cel8(Num) = 164
Num = Num + 1: Cel8(Num) = 218
Num = Num + 1: Cel8(Num) = 165
Num = Num + 1: Cel8(Num) = 2
Num = Num + 1: Cel8(Num) = 186
Num = Num + 1: Cel8(Num) = 102
Num = Num + 1: Cel8(Num) = 192
Num = Num + 1: Cel8(Num) = 112
Num = Num + 1: Cel8(Num) = 2
Num = Num + 1: Cel8(Num) = 205
Num = Num + 1: Cel8(Num) = 105
Num = Num + 1: Cel8(Num) = 212
Num = Num + 1: Cel8(Num) = 114
Num = Num + 1: Cel8(Num) = 1
Num = Num + 1: Cel8(Num) = 19
Num = Num + 1: Cel8(Num) = 180
Num = Num + 1: Cel8(Num) = 103
Num = Num + 1: Cel8(Num) = 186
Num = Num + 1: Cel8(Num) = 101
Num = Num + 1: Cel8(Num) = 194
Num = Num + 1: Cel8(Num) = 101
Num = Num + 1: Cel8(Num) = 200
Num = Num + 1: Cel8(Num) = 103
Num = Num + 1: Cel8(Num) = 202
Num = Num + 1: Cel8(Num) = 104
Num = Num + 1: Cel8(Num) = 215
Num = Num + 1: Cel8(Num) = 103
Num = Num + 1: Cel8(Num) = 220
Num = Num + 1: Cel8(Num) = 104
Num = Num + 1: Cel8(Num) = 219
Num = Num + 1: Cel8(Num) = 110
Num = Num + 1: Cel8(Num) = 216
Num = Num + 1: Cel8(Num) = 116
Num = Num + 1: Cel8(Num) = 213
Num = Num + 1: Cel8(Num) = 120
Num = Num + 1: Cel8(Num) = 209
Num = Num + 1: Cel8(Num) = 121
Num = Num + 1: Cel8(Num) = 206
Num = Num + 1: Cel8(Num) = 121
Num = Num + 1: Cel8(Num) = 202
Num = Num + 1: Cel8(Num) = 118
Num = Num + 1: Cel8(Num) = 201
Num = Num + 1: Cel8(Num) = 111
Num = Num + 1: Cel8(Num) = 198
Num = Num + 1: Cel8(Num) = 119
Num = Num + 1: Cel8(Num) = 193
Num = Num + 1: Cel8(Num) = 121
Num = Num + 1: Cel8(Num) = 189
Num = Num + 1: Cel8(Num) = 121
Num = Num + 1: Cel8(Num) = 184
Num = Num + 1: Cel8(Num) = 116
Num = Num + 1: Cel8(Num) = 181
Num = Num + 1: Cel8(Num) = 111
Num = Num + 1: Cel8(Num) = 181
Num = Num + 1: Cel8(Num) = 103
Num = Num + 1: Cel8(Num) = 0
'--------------------- database for cel 9 -----------------------
Num = 0
Cel9(Num) = 1
Num = Num + 1: Cel9(Num) = 5
Num = Num + 1: Cel9(Num) = 180
Num = Num + 1: Cel9(Num) = 164
Num = Num + 1: Cel9(Num) = 185
Num = Num + 1: Cel9(Num) = 164
Num = Num + 1: Cel9(Num) = 197
Num = Num + 1: Cel9(Num) = 162
Num = Num + 1: Cel9(Num) = 205
Num = Num + 1: Cel9(Num) = 159
Num = Num + 1: Cel9(Num) = 212
Num = Num + 1: Cel9(Num) = 159
Num = Num + 1: Cel9(Num) = 221
Num = Num + 1: Cel9(Num) = 162
Num = Num + 1: Cel9(Num) = 2
Num = Num + 1: Cel9(Num) = 185
Num = Num + 1: Cel9(Num) = 100
Num = Num + 1: Cel9(Num) = 191
Num = Num + 1: Cel9(Num) = 111
Num = Num + 1: Cel9(Num) = 1
Num = Num + 1: Cel9(Num) = 5
Num = Num + 1: Cel9(Num) = 207
Num = Num + 1: Cel9(Num) = 105
Num = Num + 1: Cel9(Num) = 208
Num = Num + 1: Cel9(Num) = 110
Num = Num + 1: Cel9(Num) = 209
Num = Num + 1: Cel9(Num) = 111
Num = Num + 1: Cel9(Num) = 211
Num = Num + 1: Cel9(Num) = 111
Num = Num + 1: Cel9(Num) = 213
Num = Num + 1: Cel9(Num) = 109
Num = Num + 1: Cel9(Num) = 213
Num = Num + 1: Cel9(Num) = 104
Num = Num + 1: Cel9(Num) = 1
Num = Num + 1: Cel9(Num) = 4
Num = Num + 1: Cel9(Num) = 177
Num = Num + 1: Cel9(Num) = 102
Num = Num + 1: Cel9(Num) = 190
Num = Num + 1: Cel9(Num) = 99
Num = Num + 1: Cel9(Num) = 204
Num = Num + 1: Cel9(Num) = 105
Num = Num + 1: Cel9(Num) = 207
Num = Num + 1: Cel9(Num) = 105
Num = Num + 1: Cel9(Num) = 212
Num = Num + 1: Cel9(Num) = 100
Num = Num + 1: Cel9(Num) = 1
Num = Num + 1: Cel9(Num) = 3
Num = Num + 1: Cel9(Num) = 207
Num = Num + 1: Cel9(Num) = 105
Num = Num + 1: Cel9(Num) = 213
Num = Num + 1: Cel9(Num) = 104
Num = Num + 1: Cel9(Num) = 220
Num = Num + 1: Cel9(Num) = 104
Num = Num + 1: Cel9(Num) = 225
Num = Num + 1: Cel9(Num) = 108
Num = Num + 1: Cel9(Num) = 1
Num = Num + 1: Cel9(Num) = 7
Num = Num + 1: Cel9(Num) = 182
Num = Num + 1: Cel9(Num) = 101
Num = Num + 1: Cel9(Num) = 181
Num = Num + 1: Cel9(Num) = 109
Num = Num + 1: Cel9(Num) = 183
Num = Num + 1: Cel9(Num) = 114
Num = Num + 1: Cel9(Num) = 187
Num = Num + 1: Cel9(Num) = 120
Num = Num + 1: Cel9(Num) = 194
Num = Num + 1: Cel9(Num) = 121
Num = Num + 1: Cel9(Num) = 198
Num = Num + 1: Cel9(Num) = 121
Num = Num + 1: Cel9(Num) = 203
Num = Num + 1: Cel9(Num) = 115
Num = Num + 1: Cel9(Num) = 204
Num = Num + 1: Cel9(Num) = 105
Num = Num + 1: Cel9(Num) = 1
Num = Num + 1: Cel9(Num) = 6
Num = Num + 1: Cel9(Num) = 203
Num = Num + 1: Cel9(Num) = 115
Num = Num + 1: Cel9(Num) = 207
Num = Num + 1: Cel9(Num) = 120
Num = Num + 1: Cel9(Num) = 212
Num = Num + 1: Cel9(Num) = 122
Num = Num + 1: Cel9(Num) = 216
Num = Num + 1: Cel9(Num) = 121
Num = Num + 1: Cel9(Num) = 220
Num = Num + 1: Cel9(Num) = 115
Num = Num + 1: Cel9(Num) = 222
Num = Num + 1: Cel9(Num) = 110
Num = Num + 1: Cel9(Num) = 223
Num = Num + 1: Cel9(Num) = 104
Num = Num + 1: Cel9(Num) = 0
'-------------------- database for cel 10 -----------------------
Num = 0
Cel10(Num) = 1
Num = Num + 1: Cel10(Num) = 6
Num = Num + 1: Cel10(Num) = 180
Num = Num + 1: Cel10(Num) = 170
Num = Num + 1: Cel10(Num) = 185
Num = Num + 1: Cel10(Num) = 170
Num = Num + 1: Cel10(Num) = 200
Num = Num + 1: Cel10(Num) = 161
Num = Num + 1: Cel10(Num) = 207
Num = Num + 1: Cel10(Num) = 158
Num = Num + 1: Cel10(Num) = 214
Num = Num + 1: Cel10(Num) = 158
Num = Num + 1: Cel10(Num) = 220
Num = Num + 1: Cel10(Num) = 162
Num = Num + 1: Cel10(Num) = 222
Num = Num + 1: Cel10(Num) = 167
Num = Num + 1: Cel10(Num) = 1
Num = Num + 1: Cel10(Num) = 2
Num = Num + 1: Cel10(Num) = 222
Num = Num + 1: Cel10(Num) = 171
Num = Num + 1: Cel10(Num) = 225
Num = Num + 1: Cel10(Num) = 165
Num = Num + 1: Cel10(Num) = 225
Num = Num + 1: Cel10(Num) = 161
Num = Num + 1: Cel10(Num) = 2
Num = Num + 1: Cel10(Num) = 180
Num = Num + 1: Cel10(Num) = 99
Num = Num + 1: Cel10(Num) = 188
Num = Num + 1: Cel10(Num) = 108
Num = Num + 1: Cel10(Num) = 1
Num = Num + 1: Cel10(Num) = 4
Num = Num + 1: Cel10(Num) = 205
Num = Num + 1: Cel10(Num) = 108
Num = Num + 1: Cel10(Num) = 207
Num = Num + 1: Cel10(Num) = 111
Num = Num + 1: Cel10(Num) = 210
Num = Num + 1: Cel10(Num) = 112
Num = Num + 1: Cel10(Num) = 212
Num = Num + 1: Cel10(Num) = 110
Num = Num + 1: Cel10(Num) = 212
Num = Num + 1: Cel10(Num) = 106
Num = Num + 1: Cel10(Num) = 1
Num = Num + 1: Cel10(Num) = 7
Num = Num + 1: Cel10(Num) = 175
Num = Num + 1: Cel10(Num) = 101
Num = Num + 1: Cel10(Num) = 181
Num = Num + 1: Cel10(Num) = 98
Num = Num + 1: Cel10(Num) = 185
Num = Num + 1: Cel10(Num) = 98
Num = Num + 1: Cel10(Num) = 192
Num = Num + 1: Cel10(Num) = 100
Num = Num + 1: Cel10(Num) = 203
Num = Num + 1: Cel10(Num) = 107
Num = Num + 1: Cel10(Num) = 205
Num = Num + 1: Cel10(Num) = 108
Num = Num + 1: Cel10(Num) = 222
Num = Num + 1: Cel10(Num) = 105
Num = Num + 1: Cel10(Num) = 226
Num = Num + 1: Cel10(Num) = 111
Num = Num + 1: Cel10(Num) = 1
Num = Num + 1: Cel10(Num) = 3
Num = Num + 1: Cel10(Num) = 204
Num = Num + 1: Cel10(Num) = 105
Num = Num + 1: Cel10(Num) = 209
Num = Num + 1: Cel10(Num) = 104
Num = Num + 1: Cel10(Num) = 210
Num = Num + 1: Cel10(Num) = 100
Num = Num + 1: Cel10(Num) = 210
Num = Num + 1: Cel10(Num) = 96
Num = Num + 1: Cel10(Num) = 1
Num = Num + 1: Cel10(Num) = 2
Num = Num + 1: Cel10(Num) = 210
Num = Num + 1: Cel10(Num) = 106
Num = Num + 1: Cel10(Num) = 213
Num = Num + 1: Cel10(Num) = 103
Num = Num + 1: Cel10(Num) = 214
Num = Num + 1: Cel10(Num) = 99
Num = Num + 1: Cel10(Num) = 1
Num = Num + 1: Cel10(Num) = 15
Num = Num + 1: Cel10(Num) = 180
Num = Num + 1: Cel10(Num) = 106
Num = Num + 1: Cel10(Num) = 180
Num = Num + 1: Cel10(Num) = 114
Num = Num + 1: Cel10(Num) = 183
Num = Num + 1: Cel10(Num) = 120
Num = Num + 1: Cel10(Num) = 190
Num = Num + 1: Cel10(Num) = 123
Num = Num + 1: Cel10(Num) = 195
Num = Num + 1: Cel10(Num) = 121
Num = Num + 1: Cel10(Num) = 200
Num = Num + 1: Cel10(Num) = 118
Num = Num + 1: Cel10(Num) = 203
Num = Num + 1: Cel10(Num) = 112
Num = Num + 1: Cel10(Num) = 203
Num = Num + 1: Cel10(Num) = 107
Num = Num + 1: Cel10(Num) = 203
Num = Num + 1: Cel10(Num) = 112
Num = Num + 1: Cel10(Num) = 204
Num = Num + 1: Cel10(Num) = 118
Num = Num + 1: Cel10(Num) = 210
Num = Num + 1: Cel10(Num) = 121
Num = Num + 1: Cel10(Num) = 215
Num = Num + 1: Cel10(Num) = 121
Num = Num + 1: Cel10(Num) = 219
Num = Num + 1: Cel10(Num) = 120
Num = Num + 1: Cel10(Num) = 221
Num = Num + 1: Cel10(Num) = 115
Num = Num + 1: Cel10(Num) = 222
Num = Num + 1: Cel10(Num) = 110
Num = Num + 1: Cel10(Num) = 222
Num = Num + 1: Cel10(Num) = 104
Num = Num + 1: Cel10(Num) = 0
'-------------------- database for cel 11 -----------------------
Num = 0
Cel11(Num) = 1
Num = Num + 1: Cel11(Num) = 6
Num = Num + 1: Cel11(Num) = 175
Num = Num + 1: Cel11(Num) = 161
Num = Num + 1: Cel11(Num) = 185
Num = Num + 1: Cel11(Num) = 164
Num = Num + 1: Cel11(Num) = 190
Num = Num + 1: Cel11(Num) = 163
Num = Num + 1: Cel11(Num) = 203
Num = Num + 1: Cel11(Num) = 160
Num = Num + 1: Cel11(Num) = 210
Num = Num + 1: Cel11(Num) = 158
Num = Num + 1: Cel11(Num) = 207
Num = Num + 1: Cel11(Num) = 160
Num = Num + 1: Cel11(Num) = 225
Num = Num + 1: Cel11(Num) = 162
Num = Num + 1: Cel11(Num) = 1
Num = Num + 1: Cel11(Num) = 1
Num = Num + 1: Cel11(Num) = 226
Num = Num + 1: Cel11(Num) = 158
Num = Num + 1: Cel11(Num) = 227
Num = Num + 1: Cel11(Num) = 168
Num = Num + 1: Cel11(Num) = 2
Num = Num + 1: Cel11(Num) = 187
Num = Num + 1: Cel11(Num) = 100
Num = Num + 1: Cel11(Num) = 194
Num = Num + 1: Cel11(Num) = 111
Num = Num + 1: Cel11(Num) = 1
Num = Num + 1: Cel11(Num) = 5
Num = Num + 1: Cel11(Num) = 210
Num = Num + 1: Cel11(Num) = 107
Num = Num + 1: Cel11(Num) = 209
Num = Num + 1: Cel11(Num) = 110
Num = Num + 1: Cel11(Num) = 210
Num = Num + 1: Cel11(Num) = 113
Num = Num + 1: Cel11(Num) = 214
Num = Num + 1: Cel11(Num) = 114
Num = Num + 1: Cel11(Num) = 216
Num = Num + 1: Cel11(Num) = 110
Num = Num + 1: Cel11(Num) = 216
Num = Num + 1: Cel11(Num) = 105
Num = Num + 1: Cel11(Num) = 1
Num = Num + 1: Cel11(Num) = 13
Num = Num + 1: Cel11(Num) = 222
Num = Num + 1: Cel11(Num) = 109
Num = Num + 1: Cel11(Num) = 222
Num = Num + 1: Cel11(Num) = 115
Num = Num + 1: Cel11(Num) = 220
Num = Num + 1: Cel11(Num) = 119
Num = Num + 1: Cel11(Num) = 216
Num = Num + 1: Cel11(Num) = 120
Num = Num + 1: Cel11(Num) = 209
Num = Num + 1: Cel11(Num) = 120
Num = Num + 1: Cel11(Num) = 205
Num = Num + 1: Cel11(Num) = 118
Num = Num + 1: Cel11(Num) = 202
Num = Num + 1: Cel11(Num) = 115
Num = Num + 1: Cel11(Num) = 199
Num = Num + 1: Cel11(Num) = 121
Num = Num + 1: Cel11(Num) = 192
Num = Num + 1: Cel11(Num) = 122
Num = Num + 1: Cel11(Num) = 189
Num = Num + 1: Cel11(Num) = 121
Num = Num + 1: Cel11(Num) = 182
Num = Num + 1: Cel11(Num) = 119
Num = Num + 1: Cel11(Num) = 180
Num = Num + 1: Cel11(Num) = 112
Num = Num + 1: Cel11(Num) = 180
Num = Num + 1: Cel11(Num) = 104
Num = Num + 1: Cel11(Num) = 181
Num = Num + 1: Cel11(Num) = 99
Num = Num + 1: Cel11(Num) = 1
Num = Num + 1: Cel11(Num) = 6
Num = Num + 1: Cel11(Num) = 175
Num = Num + 1: Cel11(Num) = 102
Num = Num + 1: Cel11(Num) = 181
Num = Num + 1: Cel11(Num) = 99
Num = Num + 1: Cel11(Num) = 185
Num = Num + 1: Cel11(Num) = 98
Num = Num + 1: Cel11(Num) = 203
Num = Num + 1: Cel11(Num) = 105
Num = Num + 1: Cel11(Num) = 207
Num = Num + 1: Cel11(Num) = 104
Num = Num + 1: Cel11(Num) = 210
Num = Num + 1: Cel11(Num) = 101
Num = Num + 1: Cel11(Num) = 210
Num = Num + 1: Cel11(Num) = 95
Num = Num + 1: Cel11(Num) = 1
Num = Num + 1: Cel11(Num) = 2
Num = Num + 1: Cel11(Num) = 210
Num = Num + 1: Cel11(Num) = 105
Num = Num + 1: Cel11(Num) = 204
Num = Num + 1: Cel11(Num) = 102
Num = Num + 1: Cel11(Num) = 205
Num = Num + 1: Cel11(Num) = 99
Num = Num + 1: Cel11(Num) = 1
Num = Num + 1: Cel11(Num) = 3
Num = Num + 1: Cel11(Num) = 204
Num = Num + 1: Cel11(Num) = 108
Num = Num + 1: Cel11(Num) = 222
Num = Num + 1: Cel11(Num) = 103
Num = Num + 1: Cel11(Num) = 225
Num = Num + 1: Cel11(Num) = 105
Num = Num + 1: Cel11(Num) = 227
Num = Num + 1: Cel11(Num) = 110
Num = Num + 1: Cel11(Num) = 0
'-------------------- database for cel 12 -----------------------
Num = 0
Cel12(Num) = 1
Num = Num + 1: Cel12(Num) = 4
Num = Num + 1: Cel12(Num) = 170
Num = Num + 1: Cel12(Num) = 153
Num = Num + 1: Cel12(Num) = 180
Num = Num + 1: Cel12(Num) = 161
Num = Num + 1: Cel12(Num) = 188
Num = Num + 1: Cel12(Num) = 163
Num = Num + 1: Cel12(Num) = 195
Num = Num + 1: Cel12(Num) = 162
Num = Num + 1: Cel12(Num) = 227
Num = Num + 1: Cel12(Num) = 152
Num = Num + 1: Cel12(Num) = 1
Num = Num + 1: Cel12(Num) = 1
Num = Num + 1: Cel12(Num) = 226
Num = Num + 1: Cel12(Num) = 145
Num = Num + 1: Cel12(Num) = 228
Num = Num + 1: Cel12(Num) = 157
Num = Num + 1: Cel12(Num) = 1
Num = Num + 1: Cel12(Num) = 6
Num = Num + 1: Cel12(Num) = 194
Num = Num + 1: Cel12(Num) = 102
Num = Num + 1: Cel12(Num) = 192
Num = Num + 1: Cel12(Num) = 108
Num = Num + 1: Cel12(Num) = 193
Num = Num + 1: Cel12(Num) = 110
Num = Num + 1: Cel12(Num) = 195
Num = Num + 1: Cel12(Num) = 112
Num = Num + 1: Cel12(Num) = 199
Num = Num + 1: Cel12(Num) = 112
Num = Num + 1: Cel12(Num) = 200
Num = Num + 1: Cel12(Num) = 110
Num = Num + 1: Cel12(Num) = 200
Num = Num + 1: Cel12(Num) = 107
Num = Num + 1: Cel12(Num) = 1
Num = Num + 1: Cel12(Num) = 6
Num = Num + 1: Cel12(Num) = 214
Num = Num + 1: Cel12(Num) = 105
Num = Num + 1: Cel12(Num) = 213
Num = Num + 1: Cel12(Num) = 108
Num = Num + 1: Cel12(Num) = 213
Num = Num + 1: Cel12(Num) = 111
Num = Num + 1: Cel12(Num) = 215
Num = Num + 1: Cel12(Num) = 113
Num = Num + 1: Cel12(Num) = 218
Num = Num + 1: Cel12(Num) = 113
Num = Num + 1: Cel12(Num) = 220
Num = Num + 1: Cel12(Num) = 110
Num = Num + 1: Cel12(Num) = 220
Num = Num + 1: Cel12(Num) = 104
Num = Num + 1: Cel12(Num) = 1
Num = Num + 1: Cel12(Num) = 13
Num = Num + 1: Cel12(Num) = 222
Num = Num + 1: Cel12(Num) = 108
Num = Num + 1: Cel12(Num) = 221
Num = Num + 1: Cel12(Num) = 117
Num = Num + 1: Cel12(Num) = 220
Num = Num + 1: Cel12(Num) = 119
Num = Num + 1: Cel12(Num) = 208
Num = Num + 1: Cel12(Num) = 118
Num = Num + 1: Cel12(Num) = 204
Num = Num + 1: Cel12(Num) = 116
Num = Num + 1: Cel12(Num) = 201
Num = Num + 1: Cel12(Num) = 113
Num = Num + 1: Cel12(Num) = 200
Num = Num + 1: Cel12(Num) = 120
Num = Num + 1: Cel12(Num) = 196
Num = Num + 1: Cel12(Num) = 121
Num = Num + 1: Cel12(Num) = 190
Num = Num + 1: Cel12(Num) = 121
Num = Num + 1: Cel12(Num) = 184
Num = Num + 1: Cel12(Num) = 120
Num = Num + 1: Cel12(Num) = 180
Num = Num + 1: Cel12(Num) = 115
Num = Num + 1: Cel12(Num) = 179
Num = Num + 1: Cel12(Num) = 108
Num = Num + 1: Cel12(Num) = 180
Num = Num + 1: Cel12(Num) = 102
Num = Num + 1: Cel12(Num) = 181
Num = Num + 1: Cel12(Num) = 100
Num = Num + 1: Cel12(Num) = 1
Num = Num + 1: Cel12(Num) = 6
Num = Num + 1: Cel12(Num) = 176
Num = Num + 1: Cel12(Num) = 103
Num = Num + 1: Cel12(Num) = 181
Num = Num + 1: Cel12(Num) = 100
Num = Num + 1: Cel12(Num) = 185
Num = Num + 1: Cel12(Num) = 99
Num = Num + 1: Cel12(Num) = 190
Num = Num + 1: Cel12(Num) = 100
Num = Num + 1: Cel12(Num) = 203
Num = Num + 1: Cel12(Num) = 107
Num = Num + 1: Cel12(Num) = 208
Num = Num + 1: Cel12(Num) = 106
Num = Num + 1: Cel12(Num) = 211
Num = Num + 1: Cel12(Num) = 95
Num = Num + 1: Cel12(Num) = 1
Num = Num + 1: Cel12(Num) = 2
Num = Num + 1: Cel12(Num) = 208
Num = Num + 1: Cel12(Num) = 105
Num = Num + 1: Cel12(Num) = 213
Num = Num + 1: Cel12(Num) = 104
Num = Num + 1: Cel12(Num) = 215
Num = Num + 1: Cel12(Num) = 99
Num = Num + 1: Cel12(Num) = 1
Num = Num + 1: Cel12(Num) = 4
Num = Num + 1: Cel12(Num) = 208
Num = Num + 1: Cel12(Num) = 105
Num = Num + 1: Cel12(Num) = 214
Num = Num + 1: Cel12(Num) = 105
Num = Num + 1: Cel12(Num) = 220
Num = Num + 1: Cel12(Num) = 104
Num = Num + 1: Cel12(Num) = 223
Num = Num + 1: Cel12(Num) = 104
Num = Num + 1: Cel12(Num) = 228
Num = Num + 1: Cel12(Num) = 107
Num = Num + 1: Cel12(Num) = 0
'-------------------- database for cel 13 -----------------------
Num = 0
Cel13(Num) = 1
Num = Num + 1: Cel13(Num) = 1
Num = Num + 1: Cel13(Num) = 166
Num = Num + 1: Cel13(Num) = 151
Num = Num + 1: Cel13(Num) = 171
Num = Num + 1: Cel13(Num) = 142
Num = Num + 1: Cel13(Num) = 1
Num = Num + 1: Cel13(Num) = 1
Num = Num + 1: Cel13(Num) = 226
Num = Num + 1: Cel13(Num) = 133
Num = Num + 1: Cel13(Num) = 233
Num = Num + 1: Cel13(Num) = 140
Num = Num + 1: Cel13(Num) = 1
Num = Num + 1: Cel13(Num) = 4
Num = Num + 1: Cel13(Num) = 170
Num = Num + 1: Cel13(Num) = 148
Num = Num + 1: Cel13(Num) = 180
Num = Num + 1: Cel13(Num) = 161
Num = Num + 1: Cel13(Num) = 190
Num = Num + 1: Cel13(Num) = 165
Num = Num + 1: Cel13(Num) = 220
Num = Num + 1: Cel13(Num) = 154
Num = Num + 1: Cel13(Num) = 230
Num = Num + 1: Cel13(Num) = 135
Num = Num + 1: Cel13(Num) = 1
Num = Num + 1: Cel13(Num) = 6
Num = Num + 1: Cel13(Num) = 197
Num = Num + 1: Cel13(Num) = 106
Num = Num + 1: Cel13(Num) = 196
Num = Num + 1: Cel13(Num) = 110
Num = Num + 1: Cel13(Num) = 197
Num = Num + 1: Cel13(Num) = 113
Num = Num + 1: Cel13(Num) = 200
Num = Num + 1: Cel13(Num) = 115
Num = Num + 1: Cel13(Num) = 203
Num = Num + 1: Cel13(Num) = 114
Num = Num + 1: Cel13(Num) = 204
Num = Num + 1: Cel13(Num) = 110
Num = Num + 1: Cel13(Num) = 204
Num = Num + 1: Cel13(Num) = 108
Num = Num + 1: Cel13(Num) = 1
Num = Num + 1: Cel13(Num) = 5
Num = Num + 1: Cel13(Num) = 218
Num = Num + 1: Cel13(Num) = 106
Num = Num + 1: Cel13(Num) = 217
Num = Num + 1: Cel13(Num) = 111
Num = Num + 1: Cel13(Num) = 219
Num = Num + 1: Cel13(Num) = 113
Num = Num + 1: Cel13(Num) = 222
Num = Num + 1: Cel13(Num) = 113
Num = Num + 1: Cel13(Num) = 223
Num = Num + 1: Cel13(Num) = 111
Num = Num + 1: Cel13(Num) = 223
Num = Num + 1: Cel13(Num) = 106
Num = Num + 1: Cel13(Num) = 1
Num = Num + 1: Cel13(Num) = 1
Num = Num + 1: Cel13(Num) = 202
Num = Num + 1: Cel13(Num) = 120
Num = Num + 1: Cel13(Num) = 213
Num = Num + 1: Cel13(Num) = 122
Num = Num + 1: Cel13(Num) = 1
Num = Num + 1: Cel13(Num) = 2
Num = Num + 1: Cel13(Num) = 205
Num = Num + 1: Cel13(Num) = 116
Num = Num + 1: Cel13(Num) = 215
Num = Num + 1: Cel13(Num) = 118
Num = Num + 1: Cel13(Num) = 225
Num = Num + 1: Cel13(Num) = 120
Num = Num + 1: Cel13(Num) = 1
Num = Num + 1: Cel13(Num) = 5
Num = Num + 1: Cel13(Num) = 201
Num = Num + 1: Cel13(Num) = 115
Num = Num + 1: Cel13(Num) = 200
Num = Num + 1: Cel13(Num) = 120
Num = Num + 1: Cel13(Num) = 190
Num = Num + 1: Cel13(Num) = 120
Num = Num + 1: Cel13(Num) = 185
Num = Num + 1: Cel13(Num) = 118
Num = Num + 1: Cel13(Num) = 180
Num = Num + 1: Cel13(Num) = 111
Num = Num + 1: Cel13(Num) = 181
Num = Num + 1: Cel13(Num) = 102
Num = Num + 1: Cel13(Num) = 1
Num = Num + 1: Cel13(Num) = 6
Num = Num + 1: Cel13(Num) = 177
Num = Num + 1: Cel13(Num) = 103
Num = Num + 1: Cel13(Num) = 186
Num = Num + 1: Cel13(Num) = 99
Num = Num + 1: Cel13(Num) = 197
Num = Num + 1: Cel13(Num) = 106
Num = Num + 1: Cel13(Num) = 204
Num = Num + 1: Cel13(Num) = 108
Num = Num + 1: Cel13(Num) = 218
Num = Num + 1: Cel13(Num) = 106
Num = Num + 1: Cel13(Num) = 223
Num = Num + 1: Cel13(Num) = 106
Num = Num + 1: Cel13(Num) = 228
Num = Num + 1: Cel13(Num) = 106
Num = Num + 1: Cel13(Num) = 1
Num = Num + 1: Cel13(Num) = 1
Num = Num + 1: Cel13(Num) = 207
Num = Num + 1: Cel13(Num) = 107
Num = Num + 1: Cel13(Num) = 211
Num = Num + 1: Cel13(Num) = 97
Num = Num + 1: Cel13(Num) = 1
Num = Num + 1: Cel13(Num) = 2
Num = Num + 1: Cel13(Num) = 207
Num = Num + 1: Cel13(Num) = 107
Num = Num + 1: Cel13(Num) = 213
Num = Num + 1: Cel13(Num) = 105
Num = Num + 1: Cel13(Num) = 215
Num = Num + 1: Cel13(Num) = 100
Num = Num + 1: Cel13(Num) = 0
'-------------------- database for cel 14 -----------------------
Num = 0
Cel14(Num) = 1
Num = Num + 1: Cel14(Num) = 1
Num = Num + 1: Cel14(Num) = 168
Num = Num + 1: Cel14(Num) = 158
Num = Num + 1: Cel14(Num) = 172
Num = Num + 1: Cel14(Num) = 149
Num = Num + 1: Cel14(Num) = 1
Num = Num + 1: Cel14(Num) = 2
Num = Num + 1: Cel14(Num) = 226
Num = Num + 1: Cel14(Num) = 141
Num = Num + 1: Cel14(Num) = 230
Num = Num + 1: Cel14(Num) = 145
Num = Num + 1: Cel14(Num) = 230
Num = Num + 1: Cel14(Num) = 150
Num = Num + 1: Cel14(Num) = 1
Num = Num + 1: Cel14(Num) = 5
Num = Num + 1: Cel14(Num) = 171
Num = Num + 1: Cel14(Num) = 151
Num = Num + 1: Cel14(Num) = 185
Num = Num + 1: Cel14(Num) = 161
Num = Num + 1: Cel14(Num) = 199
Num = Num + 1: Cel14(Num) = 161
Num = Num + 1: Cel14(Num) = 210
Num = Num + 1: Cel14(Num) = 158
Num = Num + 1: Cel14(Num) = 220
Num = Num + 1: Cel14(Num) = 152
Num = Num + 1: Cel14(Num) = 228
Num = Num + 1: Cel14(Num) = 144
Num = Num + 1: Cel14(Num) = 2
Num = Num + 1: Cel14(Num) = 191
Num = Num + 1: Cel14(Num) = 104
Num = Num + 1: Cel14(Num) = 198
Num = Num + 1: Cel14(Num) = 114
Num = Num + 1: Cel14(Num) = 2
Num = Num + 1: Cel14(Num) = 212
Num = Num + 1: Cel14(Num) = 104
Num = Num + 1: Cel14(Num) = 219
Num = Num + 1: Cel14(Num) = 114
Num = Num + 1: Cel14(Num) = 1
Num = Num + 1: Cel14(Num) = 10
Num = Num + 1: Cel14(Num) = 224
Num = Num + 1: Cel14(Num) = 112
Num = Num + 1: Cel14(Num) = 220
Num = Num + 1: Cel14(Num) = 119
Num = Num + 1: Cel14(Num) = 216
Num = Num + 1: Cel14(Num) = 121
Num = Num + 1: Cel14(Num) = 210
Num = Num + 1: Cel14(Num) = 120
Num = Num + 1: Cel14(Num) = 203
Num = Num + 1: Cel14(Num) = 115
Num = Num + 1: Cel14(Num) = 200
Num = Num + 1: Cel14(Num) = 119
Num = Num + 1: Cel14(Num) = 195
Num = Num + 1: Cel14(Num) = 120
Num = Num + 1: Cel14(Num) = 188
Num = Num + 1: Cel14(Num) = 120
Num = Num + 1: Cel14(Num) = 183
Num = Num + 1: Cel14(Num) = 115
Num = Num + 1: Cel14(Num) = 181
Num = Num + 1: Cel14(Num) = 110
Num = Num + 1: Cel14(Num) = 182
Num = Num + 1: Cel14(Num) = 103
Num = Num + 1: Cel14(Num) = 1
Num = Num + 1: Cel14(Num) = 8
Num = Num + 1: Cel14(Num) = 179
Num = Num + 1: Cel14(Num) = 103
Num = Num + 1: Cel14(Num) = 190
Num = Num + 1: Cel14(Num) = 100
Num = Num + 1: Cel14(Num) = 198
Num = Num + 1: Cel14(Num) = 100
Num = Num + 1: Cel14(Num) = 202
Num = Num + 1: Cel14(Num) = 103
Num = Num + 1: Cel14(Num) = 205
Num = Num + 1: Cel14(Num) = 107
Num = Num + 1: Cel14(Num) = 210
Num = Num + 1: Cel14(Num) = 103
Num = Num + 1: Cel14(Num) = 217
Num = Num + 1: Cel14(Num) = 102
Num = Num + 1: Cel14(Num) = 222
Num = Num + 1: Cel14(Num) = 103
Num = Num + 1: Cel14(Num) = 225
Num = Num + 1: Cel14(Num) = 107
Num = Num + 1: Cel14(Num) = 1
Num = Num + 1: Cel14(Num) = 1
Num = Num + 1: Cel14(Num) = 205
Num = Num + 1: Cel14(Num) = 107
Num = Num + 1: Cel14(Num) = 209
Num = Num + 1: Cel14(Num) = 100
Num = Num + 1: Cel14(Num) = 0
'-------------------- database for cel 15 -----------------------
Num = 0
Cel15(Num) = 1
Num = Num + 1: Cel15(Num) = 1
Num = Num + 1: Cel15(Num) = 171
Num = Num + 1: Cel15(Num) = 165
Num = Num + 1: Cel15(Num) = 175
Num = Num + 1: Cel15(Num) = 158
Num = Num + 1: Cel15(Num) = 1
Num = Num + 1: Cel15(Num) = 1
Num = Num + 1: Cel15(Num) = 220
Num = Num + 1: Cel15(Num) = 152
Num = Num + 1: Cel15(Num) = 225
Num = Num + 1: Cel15(Num) = 159
Num = Num + 1: Cel15(Num) = 1
Num = Num + 1: Cel15(Num) = 4
Num = Num + 1: Cel15(Num) = 174
Num = Num + 1: Cel15(Num) = 161
Num = Num + 1: Cel15(Num) = 188
Num = Num + 1: Cel15(Num) = 169
Num = Num + 1: Cel15(Num) = 200
Num = Num + 1: Cel15(Num) = 170
Num = Num + 1: Cel15(Num) = 210
Num = Num + 1: Cel15(Num) = 165
Num = Num + 1: Cel15(Num) = 221
Num = Num + 1: Cel15(Num) = 156
Num = Num + 1: Cel15(Num) = 2
Num = Num + 1: Cel15(Num) = 189
Num = Num + 1: Cel15(Num) = 107
Num = Num + 1: Cel15(Num) = 195
Num = Num + 1: Cel15(Num) = 116
Num = Num + 1: Cel15(Num) = 2
Num = Num + 1: Cel15(Num) = 209
Num = Num + 1: Cel15(Num) = 107
Num = Num + 1: Cel15(Num) = 215
Num = Num + 1: Cel15(Num) = 116
Num = Num + 1: Cel15(Num) = 1
Num = Num + 1: Cel15(Num) = 19
Num = Num + 1: Cel15(Num) = 201
Num = Num + 1: Cel15(Num) = 103
Num = Num + 1: Cel15(Num) = 210
Num = Num + 1: Cel15(Num) = 102
Num = Num + 1: Cel15(Num) = 217
Num = Num + 1: Cel15(Num) = 103
Num = Num + 1: Cel15(Num) = 219
Num = Num + 1: Cel15(Num) = 106
Num = Num + 1: Cel15(Num) = 219
Num = Num + 1: Cel15(Num) = 112
Num = Num + 1: Cel15(Num) = 216
Num = Num + 1: Cel15(Num) = 120
Num = Num + 1: Cel15(Num) = 212
Num = Num + 1: Cel15(Num) = 122
Num = Num + 1: Cel15(Num) = 207
Num = Num + 1: Cel15(Num) = 122
Num = Num + 1: Cel15(Num) = 203
Num = Num + 1: Cel15(Num) = 120
Num = Num + 1: Cel15(Num) = 200
Num = Num + 1: Cel15(Num) = 118
Num = Num + 1: Cel15(Num) = 198
Num = Num + 1: Cel15(Num) = 121
Num = Num + 1: Cel15(Num) = 192
Num = Num + 1: Cel15(Num) = 122
Num = Num + 1: Cel15(Num) = 188
Num = Num + 1: Cel15(Num) = 121
Num = Num + 1: Cel15(Num) = 184
Num = Num + 1: Cel15(Num) = 120
Num = Num + 1: Cel15(Num) = 181
Num = Num + 1: Cel15(Num) = 114
Num = Num + 1: Cel15(Num) = 181
Num = Num + 1: Cel15(Num) = 107
Num = Num + 1: Cel15(Num) = 185
Num = Num + 1: Cel15(Num) = 102
Num = Num + 1: Cel15(Num) = 190
Num = Num + 1: Cel15(Num) = 100
Num = Num + 1: Cel15(Num) = 196
Num = Num + 1: Cel15(Num) = 101
Num = Num + 1: Cel15(Num) = 201
Num = Num + 1: Cel15(Num) = 103
Num = Num + 1: Cel15(Num) = 0
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

