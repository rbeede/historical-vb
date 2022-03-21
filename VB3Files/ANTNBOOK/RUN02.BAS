'----------------------------------------------------------------
'     Frame animation routines for Visual Basic applications
'----------------------------------------------------------------
' Source file:  RUN02.BAS
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
  Case 1:  FileName = "RUN1.BMP": zLoadFrame
  Case 2:  FileName = "RUN2.BMP": zLoadFrame
  Case 3:  FileName = "RUN3.BMP": zLoadFrame
  Case 4:  FileName = "RUN4.BMP": zLoadFrame
  Case 5:  FileName = "RUN5.BMP": zLoadFrame
  Case 6:  FileName = "RUN6.BMP": zLoadFrame
  Case 7:  FileName = "RUN7.BMP": zLoadFrame
  Case 8:  FileName = "RUN8.BMP": zLoadFrame
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
'-------------------- database for cel 1 ------------------------
image1(0) = 120: image1(1) = 110: image1(2) = 24            'head
image1(3) = 143: image1(4) = 119: image1(5) = 160          'spine
image1(6) = 123: image1(7) = 211: image1(8) = 112
image1(9) = 248: image1(10) = 114: image1(11) = 270
image1(12) = 118: image1(13) = 295: image1(14) = 110
image1(15) = 10                                             'tail
image1(16) = 282: image1(17) = 141: image1(18) = 315 'far hindleg
image1(19) = 151: image1(20) = 336: image1(21) = 174
image1(22) = 350: image1(23) = 178
image1(24) = 196: image1(25) = 137: image1(26) = 198 'far foreleg
image1(27) = 174: image1(28) = 211: image1(29) = 193
image1(30) = 222: image1(31) = 196
image1(32) = 193: image1(33) = 147: image1(34) = 161'near foreleg
image1(35) = 162: image1(36) = 162: image1(37) = 187
image1(38) = 170: image1(39) = 191
image1(40) = 258: image1(41) = 158: image1(42) = 276'near hindleg
image1(43) = 188: image1(44) = 270: image1(45) = 213
image1(46) = 260: image1(47) = 213

'-------------------- database for cel 2 ------------------------
image2(0) = 120: image2(1) = 103: image2(2) = 24
image2(3) = 144: image2(4) = 111: image2(5) = 161
image2(6) = 115: image2(7) = 215: image2(8) = 117
image2(9) = 250: image2(10) = 112: image2(11) = 272
image2(12) = 111: image2(13) = 300: image2(14) = 109
image2(15) = 10
image2(16) = 290: image2(17) = 129: image2(18) = 327
image2(19) = 120: image2(20) = 354: image2(21) = 130
image2(22) = 368: image2(23) = 133
image2(24) = 176: image2(25) = 150: image2(26) = 146
image2(27) = 170: image2(28) = 142: image2(29) = 193
image2(30) = 146: image2(31) = 204
image2(32) = 163: image2(33) = 158: image2(34) = 126
image2(35) = 163: image2(36) = 110: image2(37) = 182
image2(38) = 108: image2(39) = 196
image2(40) = 283: image2(41) = 140: image2(42) = 320
image2(43) = 148: image2(44) = 340: image2(45) = 166
image2(46) = 342: image2(47) = 181

'-------------------- database for cel 3 ------------------------
image3(0) = 120: image3(1) = 98: image3(2) = 24
image3(3) = 145: image3(4) = 103: image3(5) = 161
image3(6) = 108: image3(7) = 212: image3(8) = 101
image3(9) = 249: image3(10) = 96: image3(11) = 270
image3(12) = 97: image3(13) = 296: image3(14) = 100
image3(15) = 10
image3(16) = 284: image3(17) = 116: image3(18) = 321
image3(19) = 116: image3(20) = 346: image3(21) = 133
image3(22) = 360: image3(23) = 138
image3(24) = 142: image3(25) = 140: image3(26) = 107
image3(27) = 140: image3(28) = 96: image3(29) = 165
image3(30) = 94: image3(31) = 178
image3(32) = 152: image3(33) = 145: image3(34) = 118
image3(35) = 152: image3(36) = 93: image3(37) = 163
image3(38) = 83: image3(39) = 175
image3(40) = 289: image3(41) = 110: image3(42) = 323
image3(43) = 100: image3(44) = 353: image3(45) = 104
image3(46) = 370: image3(47) = 106

'-------------------- database for cel 4 ------------------------
image4(0) = 120: image4(1) = 106: image4(2) = 24
image4(3) = 144: image4(4) = 116: image4(5) = 161
image4(6) = 121: image4(7) = 213: image4(8) = 110
image4(9) = 250: image4(10) = 109: image4(11) = 273
image4(12) = 110: image4(13) = 297: image4(14) = 116
image4(15) = 10
image4(16) = 273: image4(17) = 148: image4(18) = 307
image4(19) = 143: image4(20) = 318: image4(21) = 170
image4(22) = 328: image4(23) = 180
image4(24) = 177: image4(25) = 160: image4(26) = 152
image4(27) = 185: image4(28) = 127: image4(29) = 190
image4(30) = 111: image4(31) = 190
image4(32) = 169: image4(33) = 160: image4(34) = 152
image4(35) = 195: image4(36) = 133: image4(37) = 208
image4(38) = 119: image4(39) = 208
image4(40) = 285: image4(41) = 135: image4(42) = 321
image4(43) = 135: image4(44) = 345: image4(45) = 155
image4(46) = 360: image4(47) = 158

'-------------------- database for cel 5 ------------------------
image5(0) = 120: image5(1) = 128: image5(2) = 24
image5(3) = 145: image5(4) = 133: image5(5) = 164
image5(6) = 131: image5(7) = 201: image5(8) = 96
image5(9) = 231: image5(10) = 80: image5(11) = 251
image5(12) = 76: image5(13) = 278: image5(14) = 76
image5(15) = 10
image5(16) = 229: image5(17) = 122: image5(18) = 260
image5(19) = 109: image5(20) = 270: image5(21) = 131
image5(22) = 282: image5(23) = 140
image5(24) = 186: image5(25) = 160: image5(26) = 180
image5(27) = 195: image5(28) = 170: image5(29) = 212
image5(30) = 160: image5(31) = 212
image5(32) = 201: image5(33) = 146: image5(34) = 212
image5(35) = 181: image5(36) = 209: image5(37) = 203
image5(38) = 200: image5(39) = 203
image5(40) = 247: image5(41) = 122: image5(42) = 280
image5(43) = 110: image5(44) = 304: image5(45) = 124
image5(46) = 320: image5(47) = 124

'-------------------- database for cel 6 ------------------------
image6(0) = 120: image6(1) = 122: image6(2) = 24
image6(3) = 145: image6(4) = 131: image6(5) = 161
image6(6) = 135: image6(7) = 200: image6(8) = 102
image6(9) = 233: image6(10) = 98: image6(11) = 255
image6(12) = 97: image6(13) = 279: image6(14) = 90
image6(15) = 10
image6(16) = 228: image6(17) = 140: image6(18) = 220
image6(19) = 171: image6(20) = 208: image6(21) = 194
image6(22) = 195: image6(23) = 195
image6(24) = 190: image6(25) = 163: image6(26) = 216
image6(27) = 190: image6(28) = 235: image6(29) = 205
image6(30) = 248: image6(31) = 208
image6(32) = 201: image6(33) = 149: image6(34) = 232
image6(35) = 166: image6(36) = 254: image6(37) = 176
image6(38) = 266: image6(39) = 176
image6(40) = 215: image6(41) = 131: image6(42) = 250
image6(43) = 138: image6(44) = 245: image6(45) = 163
image6(46) = 247: image6(47) = 175

'-------------------- database for cel 7 ------------------------
image7(0) = 120: image7(1) = 119: image7(2) = 24
image7(3) = 145: image7(4) = 128: image7(5) = 162
image7(6) = 131: image7(7) = 210: image7(8) = 110
image7(9) = 244: image7(10) = 114: image7(11) = 265
image7(12) = 120: image7(13) = 282: image7(14) = 103
image7(15) = 10
image7(16) = 230: image7(17) = 155: image7(18) = 233
image7(19) = 191: image7(20) = 219: image7(21) = 212
image7(22) = 205: image7(23) = 214
image7(24) = 198: image7(25) = 152: image7(26) = 222
image7(27) = 181: image7(28) = 245: image7(29) = 192
image7(30) = 258: image7(31) = 191
image7(32) = 202: image7(33) = 139: image7(34) = 210
image7(35) = 175: image7(36) = 220: image7(37) = 199
image7(38) = 229: image7(39) = 200
image7(40) = 205: image7(41) = 132: image7(42) = 212
image7(43) = 170: image7(44) = 188: image7(45) = 180
image7(46) = 184: image7(47) = 191

'-------------------- database for cel 8 ------------------------
image8(0) = 120: image8(1) = 115: image8(2) = 24
image8(3) = 145: image8(4) = 124: image8(5) = 163
image8(6) = 127: image8(7) = 212: image8(8) = 111
image8(9) = 245: image8(10) = 115: image8(11) = 268
image8(12) = 120: image8(13) = 290: image8(14) = 109
image8(15) = 10
image8(16) = 270: image8(17) = 153: image8(18) = 301
image8(19) = 170: image8(20) = 308: image8(21) = 200
image8(22) = 319: image8(23) = 205
image8(24) = 199: image8(25) = 143: image8(26) = 192
image8(27) = 180: image8(28) = 215: image8(29) = 192
image8(30) = 229: image8(31) = 193
image8(32) = 201: image8(33) = 132: image8(34) = 179
image8(35) = 161: image8(36) = 179: image8(37) = 186
image8(38) = 182: image8(39) = 195
image8(40) = 237: image8(41) = 160: image8(42) = 264
image8(43) = 182: image8(44) = 253: image8(45) = 212
image8(46) = 240: image8(47) = 214
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

