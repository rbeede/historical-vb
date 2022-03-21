'----------------------------------------------------------------
'   Real-time animation routines for Visual Basic applications
'----------------------------------------------------------------
' Source file:  REAL02.BAS
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
'                 Create the bitmaps for the cels.
'----------------------------------------------------------------
Sub zBuildCels ()
  Dim Num As Integer             'index into array of coordinates
If CelsReady = True Then                 'if cels already created
  MsgBox "The cels have already been created.", MB_OK, "Animation ready"
  Exit Sub
End If
If CelsCalled = True Then       'if this procedure already called
  MsgBox "A previous attempt to create the cels has already failed, indicating a possible shortage of memory.", MB_OK, "Animation not ready"
  Exit Sub
End If
CelsCalled = True              'this procedure callable once only
RetLong = GlobalCompact(-1)           'maximize contiguous memory
'--------------- create the database for cel 1 ------------------
keyframe1(0) = 36: keyframe1(1) = 16                      'circle
keyframe1(2) = 45: keyframe1(3) = 23: keyframe1(4) = 49     'neck
keyframe1(5) = 27: keyframe1(6) = 47: keyframe1(7) = 28
keyframe1(8) = 42: keyframe1(9) = 25
keyframe1(10) = 54: keyframe1(11) = 26: keyframe1(12) = 78 'torso
keyframe1(13) = 54: keyframe1(14) = 76: keyframe1(15) = 60
keyframe1(16) = 57: keyframe1(17) = 57: keyframe1(18) = 42
keyframe1(19) = 30
keyframe1(20) = 15: keyframe1(21) = 25: keyframe1(22) = 24  'hand
keyframe1(23) = 27: keyframe1(24) = 23: keyframe1(25) = 33
keyframe1(26) = 12: keyframe1(27) = 32
keyframe1(28) = 22: keyframe1(29) = 32           'leading forearm
keyframe1(30) = 30: keyframe1(31) = 44: keyframe1(32) = 26
keyframe1(33) = 46: keyframe1(34) = 20: keyframe1(35) = 35
keyframe1(36) = 45: keyframe1(37) = 35         'leading upper arm
keyframe1(38) = 48: keyframe1(39) = 38: keyframe1(40) = 25
keyframe1(41) = 46: keyframe1(42) = 23: keyframe1(43) = 44
keyframe1(44) = 55: keyframe1(45) = 28        'trailing upper arm
keyframe1(46) = 69: keyframe1(47) = 20: keyframe1(48) = 70
keyframe1(49) = 23: keyframe1(50) = 58: keyframe1(51) = 31
keyframe1(52) = 68: keyframe1(53) = 20          'trailing forearm
keyframe1(54) = 91: keyframe1(55) = 20: keyframe1(56) = 91
keyframe1(57) = 23: keyframe1(58) = 70: keyframe1(59) = 23
keyframe1(60) = 91: keyframe1(61) = 20             'trailing hand
keyframe1(62) = 100: keyframe1(63) = 15: keyframe1(64) = 100
keyframe1(65) = 18: keyframe1(66) = 91: keyframe1(67) = 23
keyframe1(68) = 31: keyframe1(69) = 52             'leading thigh
keyframe1(70) = 65: keyframe1(71) = 55: keyframe1(72) = 71
keyframe1(73) = 60: keyframe1(74) = 35: keyframe1(75) = 55
keyframe1(76) = 40: keyframe1(77) = 57              'leading calf
keyframe1(78) = 51: keyframe1(79) = 75: keyframe1(80) = 41
keyframe1(81) = 74: keyframe1(82) = 31: keyframe1(83) = 52
keyframe1(84) = 46: keyframe1(85) = 68              'leading foot
keyframe1(86) = 51: keyframe1(87) = 75: keyframe1(88) = 22
keyframe1(89) = 72: keyframe1(90) = 40: keyframe1(91) = 68
keyframe1(92) = 76: keyframe1(93) = 60            'trailing thigh
keyframe1(94) = 106: keyframe1(95) = 92: keyframe1(96) = 101
keyframe1(97) = 94: keyframe1(98) = 69: keyframe1(99) = 60
keyframe1(100) = 76: keyframe1(101) = 60           'trailing calf
keyframe1(102) = 106: keyframe1(103) = 92: keyframe1(104) = 101
keyframe1(105) = 94: keyframe1(106) = 69: keyframe1(107) = 60
keyframe1(108) = 104: keyframe1(109) = 92          'trailing foot
keyframe1(110) = 111: keyframe1(111) = 92: keyframe1(112) = 111
keyframe1(113) = 108: keyframe1(114) = 104: keyframe1(115) = 106
keyframe1(116) = 111: keyframe1(117) = 108          'trailing toe
keyframe1(118) = 105: keyframe1(119) = 115: keyframe1(120) = 100
keyframe1(121) = 115: keyframe1(122) = 104: keyframe1(123) = 107
'--------------- create the database for cel 2 ------------------
Num = 0
keyframe2(Num) = 36
Num = Num + 1: keyframe2(Num) = 13: Num = Num + 1: keyframe2(Num) = 44
Num = Num + 1: keyframe2(Num) = 20: Num = Num + 1: keyframe2(Num) = 48
Num = Num + 1: keyframe2(Num) = 24: Num = Num + 1: keyframe2(Num) = 43
Num = Num + 1: keyframe2(Num) = 24: Num = Num + 1: keyframe2(Num) = 41
Num = Num + 1: keyframe2(Num) = 21: Num = Num + 1: keyframe2(Num) = 51
Num = Num + 1: keyframe2(Num) = 22: Num = Num + 1: keyframe2(Num) = 77
Num = Num + 1: keyframe2(Num) = 55: Num = Num + 1: keyframe2(Num) = 71
Num = Num + 1: keyframe2(Num) = 58: Num = Num + 1: keyframe2(Num) = 54
Num = Num + 1: keyframe2(Num) = 56: Num = Num + 1: keyframe2(Num) = 40
Num = Num + 1: keyframe2(Num) = 25: Num = Num + 1: keyframe2(Num) = 20
Num = Num + 1: keyframe2(Num) = 31: Num = Num + 1: keyframe2(Num) = 20
Num = Num + 1: keyframe2(Num) = 39: Num = Num + 1: keyframe2(Num) = 12
Num = Num + 1: keyframe2(Num) = 40: Num = Num + 1: keyframe2(Num) = 10
Num = Num + 1: keyframe2(Num) = 33: Num = Num + 1: keyframe2(Num) = 20
Num = Num + 1: keyframe2(Num) = 36: Num = Num + 1: keyframe2(Num) = 34
Num = Num + 1: keyframe2(Num) = 43: Num = Num + 1: keyframe2(Num) = 35
Num = Num + 1: keyframe2(Num) = 45: Num = Num + 1: keyframe2(Num) = 20
Num = Num + 1: keyframe2(Num) = 39: Num = Num + 1: keyframe2(Num) = 44
Num = Num + 1: keyframe2(Num) = 32: Num = Num + 1: keyframe2(Num) = 46
Num = Num + 1: keyframe2(Num) = 36: Num = Num + 1: keyframe2(Num) = 35
Num = Num + 1: keyframe2(Num) = 45: Num = Num + 1: keyframe2(Num) = 34
Num = Num + 1: keyframe2(Num) = 42: Num = Num + 1: keyframe2(Num) = 54
Num = Num + 1: keyframe2(Num) = 25: Num = Num + 1: keyframe2(Num) = 76
Num = Num + 1: keyframe2(Num) = 25: Num = Num + 1: keyframe2(Num) = 76
Num = Num + 1: keyframe2(Num) = 28: Num = Num + 1: keyframe2(Num) = 54
Num = Num + 1: keyframe2(Num) = 28: Num = Num + 1: keyframe2(Num) = 78
Num = Num + 1: keyframe2(Num) = 25: Num = Num + 1: keyframe2(Num) = 95
Num = Num + 1: keyframe2(Num) = 33: Num = Num + 1: keyframe2(Num) = 95
Num = Num + 1: keyframe2(Num) = 37: Num = Num + 1: keyframe2(Num) = 74
Num = Num + 1: keyframe2(Num) = 28: Num = Num + 1: keyframe2(Num) = 95
Num = Num + 1: keyframe2(Num) = 33: Num = Num + 1: keyframe2(Num) = 105
Num = Num + 1: keyframe2(Num) = 30: Num = Num + 1: keyframe2(Num) = 104
Num = Num + 1: keyframe2(Num) = 35: Num = Num + 1: keyframe2(Num) = 95
Num = Num + 1: keyframe2(Num) = 37: Num = Num + 1: keyframe2(Num) = 52
Num = Num + 1: keyframe2(Num) = 52: Num = Num + 1: keyframe2(Num) = 54
Num = Num + 1: keyframe2(Num) = 55: Num = Num + 1: keyframe2(Num) = 32
Num = Num + 1: keyframe2(Num) = 60: Num = Num + 1: keyframe2(Num) = 28
Num = Num + 1: keyframe2(Num) = 59: Num = Num + 1: keyframe2(Num) = 28
Num = Num + 1: keyframe2(Num) = 59: Num = Num + 1: keyframe2(Num) = 32
Num = Num + 1: keyframe2(Num) = 60: Num = Num + 1: keyframe2(Num) = 32
Num = Num + 1: keyframe2(Num) = 85: Num = Num + 1: keyframe2(Num) = 28
Num = Num + 1: keyframe2(Num) = 78: Num = Num + 1: keyframe2(Num) = 32
Num = Num + 1: keyframe2(Num) = 78: Num = Num + 1: keyframe2(Num) = 32
Num = Num + 1: keyframe2(Num) = 85: Num = Num + 1: keyframe2(Num) = 10
Num = Num + 1: keyframe2(Num) = 75: Num = Num + 1: keyframe2(Num) = 28
Num = Num + 1: keyframe2(Num) = 78: Num = Num + 1: keyframe2(Num) = 75
Num = Num + 1: keyframe2(Num) = 56: Num = Num + 1: keyframe2(Num) = 92
Num = Num + 1: keyframe2(Num) = 73: Num = Num + 1: keyframe2(Num) = 88
Num = Num + 1: keyframe2(Num) = 76: Num = Num + 1: keyframe2(Num) = 70
Num = Num + 1: keyframe2(Num) = 57: Num = Num + 1: keyframe2(Num) = 92
Num = Num + 1: keyframe2(Num) = 74: Num = Num + 1: keyframe2(Num) = 116
Num = Num + 1: keyframe2(Num) = 86: Num = Num + 1: keyframe2(Num) = 112
Num = Num + 1: keyframe2(Num) = 90: Num = Num + 1: keyframe2(Num) = 88
Num = Num + 1: keyframe2(Num) = 76: Num = Num + 1: keyframe2(Num) = 116
Num = Num + 1: keyframe2(Num) = 86: Num = Num + 1: keyframe2(Num) = 120
Num = Num + 1: keyframe2(Num) = 85: Num = Num + 1: keyframe2(Num) = 120
Num = Num + 1: keyframe2(Num) = 108: Num = Num + 1: keyframe2(Num) = 112
Num = Num + 1: keyframe2(Num) = 90: Num = Num + 1: keyframe2(Num) = 116
Num = Num + 1: keyframe2(Num) = 86: Num = Num + 1: keyframe2(Num) = 120
Num = Num + 1: keyframe2(Num) = 85: Num = Num + 1: keyframe2(Num) = 120
Num = Num + 1: keyframe2(Num) = 108: Num = Num + 1: keyframe2(Num) = 112
Num = Num + 1: keyframe2(Num) = 90
'--------------- create the database for cel 3 ------------------
Num = 0
keyframe3(Num) = 38
Num = Num + 1: keyframe3(Num) = 10: Num = Num + 1: keyframe3(Num) = 45
Num = Num + 1: keyframe3(Num) = 16: Num = Num + 1: keyframe3(Num) = 48
Num = Num + 1: keyframe3(Num) = 21: Num = Num + 1: keyframe3(Num) = 44
Num = Num + 1: keyframe3(Num) = 22: Num = Num + 1: keyframe3(Num) = 41
Num = Num + 1: keyframe3(Num) = 17: Num = Num + 1: keyframe3(Num) = 52
Num = Num + 1: keyframe3(Num) = 20: Num = Num + 1: keyframe3(Num) = 75
Num = Num + 1: keyframe3(Num) = 55: Num = Num + 1: keyframe3(Num) = 70
Num = Num + 1: keyframe3(Num) = 56: Num = Num + 1: keyframe3(Num) = 53
Num = Num + 1: keyframe3(Num) = 54: Num = Num + 1: keyframe3(Num) = 40
Num = Num + 1: keyframe3(Num) = 22: Num = Num + 1: keyframe3(Num) = 20
Num = Num + 1: keyframe3(Num) = 36: Num = Num + 1: keyframe3(Num) = 24
Num = Num + 1: keyframe3(Num) = 44: Num = Num + 1: keyframe3(Num) = 17
Num = Num + 1: keyframe3(Num) = 47: Num = Num + 1: keyframe3(Num) = 13
Num = Num + 1: keyframe3(Num) = 40: Num = Num + 1: keyframe3(Num) = 22
Num = Num + 1: keyframe3(Num) = 40: Num = Num + 1: keyframe3(Num) = 40
Num = Num + 1: keyframe3(Num) = 43: Num = Num + 1: keyframe3(Num) = 40
Num = Num + 1: keyframe3(Num) = 47: Num = Num + 1: keyframe3(Num) = 22
Num = Num + 1: keyframe3(Num) = 44: Num = Num + 1: keyframe3(Num) = 37
Num = Num + 1: keyframe3(Num) = 43: Num = Num + 1: keyframe3(Num) = 44
Num = Num + 1: keyframe3(Num) = 31: Num = Num + 1: keyframe3(Num) = 47
Num = Num + 1: keyframe3(Num) = 36: Num = Num + 1: keyframe3(Num) = 40
Num = Num + 1: keyframe3(Num) = 47: Num = Num + 1: keyframe3(Num) = 55
Num = Num + 1: keyframe3(Num) = 25: Num = Num + 1: keyframe3(Num) = 76
Num = Num + 1: keyframe3(Num) = 29: Num = Num + 1: keyframe3(Num) = 74
Num = Num + 1: keyframe3(Num) = 31: Num = Num + 1: keyframe3(Num) = 55
Num = Num + 1: keyframe3(Num) = 29: Num = Num + 1: keyframe3(Num) = 76
Num = Num + 1: keyframe3(Num) = 29: Num = Num + 1: keyframe3(Num) = 90
Num = Num + 1: keyframe3(Num) = 45: Num = Num + 1: keyframe3(Num) = 85
Num = Num + 1: keyframe3(Num) = 48: Num = Num + 1: keyframe3(Num) = 74
Num = Num + 1: keyframe3(Num) = 31: Num = Num + 1: keyframe3(Num) = 90
Num = Num + 1: keyframe3(Num) = 45: Num = Num + 1: keyframe3(Num) = 100
Num = Num + 1: keyframe3(Num) = 45: Num = Num + 1: keyframe3(Num) = 95
Num = Num + 1: keyframe3(Num) = 49: Num = Num + 1: keyframe3(Num) = 85
Num = Num + 1: keyframe3(Num) = 49: Num = Num + 1: keyframe3(Num) = 53
Num = Num + 1: keyframe3(Num) = 54: Num = Num + 1: keyframe3(Num) = 55
Num = Num + 1: keyframe3(Num) = 58: Num = Num + 1: keyframe3(Num) = 35
Num = Num + 1: keyframe3(Num) = 72: Num = Num + 1: keyframe3(Num) = 30
Num = Num + 1: keyframe3(Num) = 69: Num = Num + 1: keyframe3(Num) = 35
Num = Num + 1: keyframe3(Num) = 72: Num = Num + 1: keyframe3(Num) = 16
Num = Num + 1: keyframe3(Num) = 96: Num = Num + 1: keyframe3(Num) = 17
Num = Num + 1: keyframe3(Num) = 85: Num = Num + 1: keyframe3(Num) = 30
Num = Num + 1: keyframe3(Num) = 69: Num = Num + 1: keyframe3(Num) = 17
Num = Num + 1: keyframe3(Num) = 85: Num = Num + 1: keyframe3(Num) = 20
Num = Num + 1: keyframe3(Num) = 90: Num = Num + 1: keyframe3(Num) = 17
Num = Num + 1: keyframe3(Num) = 96: Num = Num + 1: keyframe3(Num) = 0
Num = Num + 1: keyframe3(Num) = 80: Num = Num + 1: keyframe3(Num) = 75
Num = Num + 1: keyframe3(Num) = 55: Num = Num + 1: keyframe3(Num) = 92
Num = Num + 1: keyframe3(Num) = 70: Num = Num + 1: keyframe3(Num) = 88
Num = Num + 1: keyframe3(Num) = 74: Num = Num + 1: keyframe3(Num) = 70
Num = Num + 1: keyframe3(Num) = 57: Num = Num + 1: keyframe3(Num) = 92
Num = Num + 1: keyframe3(Num) = 70: Num = Num + 1: keyframe3(Num) = 121
Num = Num + 1: keyframe3(Num) = 77: Num = Num + 1: keyframe3(Num) = 122
Num = Num + 1: keyframe3(Num) = 81: Num = Num + 1: keyframe3(Num) = 87
Num = Num + 1: keyframe3(Num) = 73: Num = Num + 1: keyframe3(Num) = 121
Num = Num + 1: keyframe3(Num) = 77: Num = Num + 1: keyframe3(Num) = 126
Num = Num + 1: keyframe3(Num) = 75: Num = Num + 1: keyframe3(Num) = 139
Num = Num + 1: keyframe3(Num) = 95: Num = Num + 1: keyframe3(Num) = 122
Num = Num + 1: keyframe3(Num) = 81: Num = Num + 1: keyframe3(Num) = 121
Num = Num + 1: keyframe3(Num) = 77: Num = Num + 1: keyframe3(Num) = 126
Num = Num + 1: keyframe3(Num) = 75: Num = Num + 1: keyframe3(Num) = 139
Num = Num + 1: keyframe3(Num) = 95: Num = Num + 1: keyframe3(Num) = 122
Num = Num + 1: keyframe3(Num) = 81
'--------------- create the database for cel 4 ------------------
Num = 0
keyframe4(Num) = 38
Num = Num + 1: keyframe4(Num) = 13: Num = Num + 1: keyframe4(Num) = 45
Num = Num + 1: keyframe4(Num) = 19: Num = Num + 1: keyframe4(Num) = 48
Num = Num + 1: keyframe4(Num) = 23: Num = Num + 1: keyframe4(Num) = 44
Num = Num + 1: keyframe4(Num) = 25: Num = Num + 1: keyframe4(Num) = 41
Num = Num + 1: keyframe4(Num) = 20: Num = Num + 1: keyframe4(Num) = 52
Num = Num + 1: keyframe4(Num) = 22: Num = Num + 1: keyframe4(Num) = 70
Num = Num + 1: keyframe4(Num) = 56: Num = Num + 1: keyframe4(Num) = 64
Num = Num + 1: keyframe4(Num) = 60: Num = Num + 1: keyframe4(Num) = 46
Num = Num + 1: keyframe4(Num) = 60: Num = Num + 1: keyframe4(Num) = 40
Num = Num + 1: keyframe4(Num) = 26: Num = Num + 1: keyframe4(Num) = 32
Num = Num + 1: keyframe4(Num) = 43: Num = Num + 1: keyframe4(Num) = 36
Num = Num + 1: keyframe4(Num) = 46: Num = Num + 1: keyframe4(Num) = 34
Num = Num + 1: keyframe4(Num) = 51: Num = Num + 1: keyframe4(Num) = 27
Num = Num + 1: keyframe4(Num) = 48: Num = Num + 1: keyframe4(Num) = 36
Num = Num + 1: keyframe4(Num) = 46: Num = Num + 1: keyframe4(Num) = 50
Num = Num + 1: keyframe4(Num) = 44: Num = Num + 1: keyframe4(Num) = 50
Num = Num + 1: keyframe4(Num) = 48: Num = Num + 1: keyframe4(Num) = 36
Num = Num + 1: keyframe4(Num) = 49: Num = Num + 1: keyframe4(Num) = 36
Num = Num + 1: keyframe4(Num) = 46: Num = Num + 1: keyframe4(Num) = 50
Num = Num + 1: keyframe4(Num) = 44: Num = Num + 1: keyframe4(Num) = 50
Num = Num + 1: keyframe4(Num) = 48: Num = Num + 1: keyframe4(Num) = 36
Num = Num + 1: keyframe4(Num) = 49: Num = Num + 1: keyframe4(Num) = 55
Num = Num + 1: keyframe4(Num) = 35: Num = Num + 1: keyframe4(Num) = 69
Num = Num + 1: keyframe4(Num) = 45: Num = Num + 1: keyframe4(Num) = 62
Num = Num + 1: keyframe4(Num) = 50: Num = Num + 1: keyframe4(Num) = 55
Num = Num + 1: keyframe4(Num) = 40: Num = Num + 1: keyframe4(Num) = 55
Num = Num + 1: keyframe4(Num) = 35: Num = Num + 1: keyframe4(Num) = 69
Num = Num + 1: keyframe4(Num) = 45: Num = Num + 1: keyframe4(Num) = 62
Num = Num + 1: keyframe4(Num) = 50: Num = Num + 1: keyframe4(Num) = 55
Num = Num + 1: keyframe4(Num) = 40: Num = Num + 1: keyframe4(Num) = 55
Num = Num + 1: keyframe4(Num) = 35: Num = Num + 1: keyframe4(Num) = 69
Num = Num + 1: keyframe4(Num) = 45: Num = Num + 1: keyframe4(Num) = 62
Num = Num + 1: keyframe4(Num) = 50: Num = Num + 1: keyframe4(Num) = 55
Num = Num + 1: keyframe4(Num) = 40: Num = Num + 1: keyframe4(Num) = 46
Num = Num + 1: keyframe4(Num) = 59: Num = Num + 1: keyframe4(Num) = 51
Num = Num + 1: keyframe4(Num) = 60: Num = Num + 1: keyframe4(Num) = 31
Num = Num + 1: keyframe4(Num) = 75: Num = Num + 1: keyframe4(Num) = 26
Num = Num + 1: keyframe4(Num) = 73: Num = Num + 1: keyframe4(Num) = 26
Num = Num + 1: keyframe4(Num) = 73: Num = Num + 1: keyframe4(Num) = 31
Num = Num + 1: keyframe4(Num) = 76: Num = Num + 1: keyframe4(Num) = 25
Num = Num + 1: keyframe4(Num) = 105: Num = Num + 1: keyframe4(Num) = 22
Num = Num + 1: keyframe4(Num) = 97: Num = Num + 1: keyframe4(Num) = 22
Num = Num + 1: keyframe4(Num) = 97: Num = Num + 1: keyframe4(Num) = 26
Num = Num + 1: keyframe4(Num) = 99: Num = Num + 1: keyframe4(Num) = 25
Num = Num + 1: keyframe4(Num) = 105: Num = Num + 1: keyframe4(Num) = 0
Num = Num + 1: keyframe4(Num) = 95: Num = Num + 1: keyframe4(Num) = 58
Num = Num + 1: keyframe4(Num) = 60: Num = Num + 1: keyframe4(Num) = 64
Num = Num + 1: keyframe4(Num) = 60: Num = Num + 1: keyframe4(Num) = 58
Num = Num + 1: keyframe4(Num) = 78: Num = Num + 1: keyframe4(Num) = 50
Num = Num + 1: keyframe4(Num) = 78: Num = Num + 1: keyframe4(Num) = 58
Num = Num + 1: keyframe4(Num) = 78: Num = Num + 1: keyframe4(Num) = 75
Num = Num + 1: keyframe4(Num) = 85: Num = Num + 1: keyframe4(Num) = 74
Num = Num + 1: keyframe4(Num) = 89: Num = Num + 1: keyframe4(Num) = 50
Num = Num + 1: keyframe4(Num) = 78: Num = Num + 1: keyframe4(Num) = 75
Num = Num + 1: keyframe4(Num) = 85: Num = Num + 1: keyframe4(Num) = 80
Num = Num + 1: keyframe4(Num) = 85: Num = Num + 1: keyframe4(Num) = 80
Num = Num + 1: keyframe4(Num) = 103: Num = Num + 1: keyframe4(Num) = 74
Num = Num + 1: keyframe4(Num) = 88: Num = Num + 1: keyframe4(Num) = 75
Num = Num + 1: keyframe4(Num) = 85: Num = Num + 1: keyframe4(Num) = 80
Num = Num + 1: keyframe4(Num) = 85: Num = Num + 1: keyframe4(Num) = 80
Num = Num + 1: keyframe4(Num) = 103: Num = Num + 1: keyframe4(Num) = 74
Num = Num + 1: keyframe4(Num) = 88
'--------------- create the database for cel 5 ------------------
Num = 0
keyframe5(Num) = 38
Num = Num + 1: keyframe5(Num) = 18: Num = Num + 1: keyframe5(Num) = 45
Num = Num + 1: keyframe5(Num) = 25: Num = Num + 1: keyframe5(Num) = 49
Num = Num + 1: keyframe5(Num) = 28: Num = Num + 1: keyframe5(Num) = 43
Num = Num + 1: keyframe5(Num) = 30: Num = Num + 1: keyframe5(Num) = 42
Num = Num + 1: keyframe5(Num) = 26: Num = Num + 1: keyframe5(Num) = 54
Num = Num + 1: keyframe5(Num) = 28: Num = Num + 1: keyframe5(Num) = 56
Num = Num + 1: keyframe5(Num) = 64: Num = Num + 1: keyframe5(Num) = 45
Num = Num + 1: keyframe5(Num) = 66: Num = Num + 1: keyframe5(Num) = 34
Num = Num + 1: keyframe5(Num) = 64: Num = Num + 1: keyframe5(Num) = 40
Num = Num + 1: keyframe5(Num) = 31: Num = Num + 1: keyframe5(Num) = 27
Num = Num + 1: keyframe5(Num) = 58: Num = Num + 1: keyframe5(Num) = 26
Num = Num + 1: keyframe5(Num) = 63: Num = Num + 1: keyframe5(Num) = 18
Num = Num + 1: keyframe5(Num) = 58: Num = Num + 1: keyframe5(Num) = 21
Num = Num + 1: keyframe5(Num) = 55: Num = Num + 1: keyframe5(Num) = 28
Num = Num + 1: keyframe5(Num) = 58: Num = Num + 1: keyframe5(Num) = 40
Num = Num + 1: keyframe5(Num) = 56: Num = Num + 1: keyframe5(Num) = 40
Num = Num + 1: keyframe5(Num) = 60: Num = Num + 1: keyframe5(Num) = 25
Num = Num + 1: keyframe5(Num) = 63: Num = Num + 1: keyframe5(Num) = 28
Num = Num + 1: keyframe5(Num) = 58: Num = Num + 1: keyframe5(Num) = 40
Num = Num + 1: keyframe5(Num) = 56: Num = Num + 1: keyframe5(Num) = 40
Num = Num + 1: keyframe5(Num) = 60: Num = Num + 1: keyframe5(Num) = 25
Num = Num + 1: keyframe5(Num) = 63: Num = Num + 1: keyframe5(Num) = 54
Num = Num + 1: keyframe5(Num) = 28: Num = Num + 1: keyframe5(Num) = 57
Num = Num + 1: keyframe5(Num) = 48: Num = Num + 1: keyframe5(Num) = 60
Num = Num + 1: keyframe5(Num) = 46: Num = Num + 1: keyframe5(Num) = 54
Num = Num + 1: keyframe5(Num) = 38: Num = Num + 1: keyframe5(Num) = 60
Num = Num + 1: keyframe5(Num) = 48: Num = Num + 1: keyframe5(Num) = 66
Num = Num + 1: keyframe5(Num) = 48: Num = Num + 1: keyframe5(Num) = 52
Num = Num + 1: keyframe5(Num) = 53: Num = Num + 1: keyframe5(Num) = 50
Num = Num + 1: keyframe5(Num) = 50: Num = Num + 1: keyframe5(Num) = 60
Num = Num + 1: keyframe5(Num) = 48: Num = Num + 1: keyframe5(Num) = 66
Num = Num + 1: keyframe5(Num) = 48: Num = Num + 1: keyframe5(Num) = 52
Num = Num + 1: keyframe5(Num) = 53: Num = Num + 1: keyframe5(Num) = 50
Num = Num + 1: keyframe5(Num) = 50: Num = Num + 1: keyframe5(Num) = 34
Num = Num + 1: keyframe5(Num) = 65: Num = Num + 1: keyframe5(Num) = 40
Num = Num + 1: keyframe5(Num) = 66: Num = Num + 1: keyframe5(Num) = 14
Num = Num + 1: keyframe5(Num) = 76: Num = Num + 1: keyframe5(Num) = 12
Num = Num + 1: keyframe5(Num) = 74: Num = Num + 1: keyframe5(Num) = 22
Num = Num + 1: keyframe5(Num) = 74: Num = Num + 1: keyframe5(Num) = 40
Num = Num + 1: keyframe5(Num) = 78: Num = Num + 1: keyframe5(Num) = 38
Num = Num + 1: keyframe5(Num) = 82: Num = Num + 1: keyframe5(Num) = 12
Num = Num + 1: keyframe5(Num) = 76: Num = Num + 1: keyframe5(Num) = 40
Num = Num + 1: keyframe5(Num) = 78: Num = Num + 1: keyframe5(Num) = 46
Num = Num + 1: keyframe5(Num) = 76: Num = Num + 1: keyframe5(Num) = 45
Num = Num + 1: keyframe5(Num) = 96: Num = Num + 1: keyframe5(Num) = 38
Num = Num + 1: keyframe5(Num) = 82: Num = Num + 1: keyframe5(Num) = 40
Num = Num + 1: keyframe5(Num) = 68: Num = Num + 1: keyframe5(Num) = 45
Num = Num + 1: keyframe5(Num) = 66: Num = Num + 1: keyframe5(Num) = 27
Num = Num + 1: keyframe5(Num) = 85: Num = Num + 1: keyframe5(Num) = 20
Num = Num + 1: keyframe5(Num) = 84: Num = Num + 1: keyframe5(Num) = 27
Num = Num + 1: keyframe5(Num) = 85: Num = Num + 1: keyframe5(Num) = 33
Num = Num + 1: keyframe5(Num) = 114: Num = Num + 1: keyframe5(Num) = 25
Num = Num + 1: keyframe5(Num) = 107: Num = Num + 1: keyframe5(Num) = 20
Num = Num + 1: keyframe5(Num) = 84: Num = Num + 1: keyframe5(Num) = 25
Num = Num + 1: keyframe5(Num) = 107: Num = Num + 1: keyframe5(Num) = 31
Num = Num + 1: keyframe5(Num) = 107: Num = Num + 1: keyframe5(Num) = 32
Num = Num + 1: keyframe5(Num) = 115: Num = Num + 1: keyframe5(Num) = 5
Num = Num + 1: keyframe5(Num) = 115: Num = Num + 1: keyframe5(Num) = 25
Num = Num + 1: keyframe5(Num) = 107: Num = Num + 1: keyframe5(Num) = 31
Num = Num + 1: keyframe5(Num) = 107: Num = Num + 1: keyframe5(Num) = 32
Num = Num + 1: keyframe5(Num) = 115: Num = Num + 1: keyframe5(Num) = 5
Num = Num + 1: keyframe5(Num) = 115
'--------------- create the database for cel 6 ------------------
Num = 0
keyframe6(Num) = 38
Num = Num + 1: keyframe6(Num) = 22: Num = Num + 1: keyframe6(Num) = 41
Num = Num + 1: keyframe6(Num) = 29: Num = Num + 1: keyframe6(Num) = 45
Num = Num + 1: keyframe6(Num) = 35: Num = Num + 1: keyframe6(Num) = 39
Num = Num + 1: keyframe6(Num) = 35: Num = Num + 1: keyframe6(Num) = 39
Num = Num + 1: keyframe6(Num) = 30: Num = Num + 1: keyframe6(Num) = 38
Num = Num + 1: keyframe6(Num) = 36: Num = Num + 1: keyframe6(Num) = 49
Num = Num + 1: keyframe6(Num) = 33: Num = Num + 1: keyframe6(Num) = 60
Num = Num + 1: keyframe6(Num) = 68: Num = Num + 1: keyframe6(Num) = 52
Num = Num + 1: keyframe6(Num) = 70: Num = Num + 1: keyframe6(Num) = 39
Num = Num + 1: keyframe6(Num) = 70: Num = Num + 1: keyframe6(Num) = 12
Num = Num + 1: keyframe6(Num) = 48: Num = Num + 1: keyframe6(Num) = 16
Num = Num + 1: keyframe6(Num) = 54: Num = Num + 1: keyframe6(Num) = 11
Num = Num + 1: keyframe6(Num) = 59: Num = Num + 1: keyframe6(Num) = 5
Num = Num + 1: keyframe6(Num) = 51: Num = Num + 1: keyframe6(Num) = 15
Num = Num + 1: keyframe6(Num) = 52: Num = Num + 1: keyframe6(Num) = 34
Num = Num + 1: keyframe6(Num) = 55: Num = Num + 1: keyframe6(Num) = 38
Num = Num + 1: keyframe6(Num) = 59: Num = Num + 1: keyframe6(Num) = 15
Num = Num + 1: keyframe6(Num) = 56: Num = Num + 1: keyframe6(Num) = 33
Num = Num + 1: keyframe6(Num) = 55: Num = Num + 1: keyframe6(Num) = 40
Num = Num + 1: keyframe6(Num) = 40: Num = Num + 1: keyframe6(Num) = 40
Num = Num + 1: keyframe6(Num) = 50: Num = Num + 1: keyframe6(Num) = 38
Num = Num + 1: keyframe6(Num) = 60: Num = Num + 1: keyframe6(Num) = 50
Num = Num + 1: keyframe6(Num) = 38: Num = Num + 1: keyframe6(Num) = 65
Num = Num + 1: keyframe6(Num) = 50: Num = Num + 1: keyframe6(Num) = 60
Num = Num + 1: keyframe6(Num) = 50: Num = Num + 1: keyframe6(Num) = 50
Num = Num + 1: keyframe6(Num) = 43: Num = Num + 1: keyframe6(Num) = 60
Num = Num + 1: keyframe6(Num) = 50: Num = Num + 1: keyframe6(Num) = 65
Num = Num + 1: keyframe6(Num) = 50: Num = Num + 1: keyframe6(Num) = 60
Num = Num + 1: keyframe6(Num) = 67: Num = Num + 1: keyframe6(Num) = 58
Num = Num + 1: keyframe6(Num) = 63: Num = Num + 1: keyframe6(Num) = 60
Num = Num + 1: keyframe6(Num) = 50: Num = Num + 1: keyframe6(Num) = 65
Num = Num + 1: keyframe6(Num) = 50: Num = Num + 1: keyframe6(Num) = 60
Num = Num + 1: keyframe6(Num) = 67: Num = Num + 1: keyframe6(Num) = 58
Num = Num + 1: keyframe6(Num) = 63: Num = Num + 1: keyframe6(Num) = 38
Num = Num + 1: keyframe6(Num) = 66: Num = Num + 1: keyframe6(Num) = 39
Num = Num + 1: keyframe6(Num) = 70: Num = Num + 1: keyframe6(Num) = 16
Num = Num + 1: keyframe6(Num) = 80: Num = Num + 1: keyframe6(Num) = 15
Num = Num + 1: keyframe6(Num) = 76: Num = Num + 1: keyframe6(Num) = 22
Num = Num + 1: keyframe6(Num) = 77: Num = Num + 1: keyframe6(Num) = 49
Num = Num + 1: keyframe6(Num) = 85: Num = Num + 1: keyframe6(Num) = 36
Num = Num + 1: keyframe6(Num) = 85: Num = Num + 1: keyframe6(Num) = 16
Num = Num + 1: keyframe6(Num) = 80: Num = Num + 1: keyframe6(Num) = 40
Num = Num + 1: keyframe6(Num) = 83: Num = Num + 1: keyframe6(Num) = 49
Num = Num + 1: keyframe6(Num) = 86: Num = Num + 1: keyframe6(Num) = 32
Num = Num + 1: keyframe6(Num) = 102: Num = Num + 1: keyframe6(Num) = 36
Num = Num + 1: keyframe6(Num) = 86: Num = Num + 1: keyframe6(Num) = 45
Num = Num + 1: keyframe6(Num) = 72: Num = Num + 1: keyframe6(Num) = 52
Num = Num + 1: keyframe6(Num) = 70: Num = Num + 1: keyframe6(Num) = 43
Num = Num + 1: keyframe6(Num) = 91: Num = Num + 1: keyframe6(Num) = 38
Num = Num + 1: keyframe6(Num) = 90: Num = Num + 1: keyframe6(Num) = 41
Num = Num + 1: keyframe6(Num) = 89: Num = Num + 1: keyframe6(Num) = 62
Num = Num + 1: keyframe6(Num) = 104: Num = Num + 1: keyframe6(Num) = 54
Num = Num + 1: keyframe6(Num) = 103: Num = Num + 1: keyframe6(Num) = 38
Num = Num + 1: keyframe6(Num) = 92: Num = Num + 1: keyframe6(Num) = 53
Num = Num + 1: keyframe6(Num) = 104: Num = Num + 1: keyframe6(Num) = 63
Num = Num + 1: keyframe6(Num) = 104: Num = Num + 1: keyframe6(Num) = 54
Num = Num + 1: keyframe6(Num) = 115: Num = Num + 1: keyframe6(Num) = 49
Num = Num + 1: keyframe6(Num) = 112: Num = Num + 1: keyframe6(Num) = 52
Num = Num + 1: keyframe6(Num) = 112: Num = Num + 1: keyframe6(Num) = 53
Num = Num + 1: keyframe6(Num) = 115: Num = Num + 1: keyframe6(Num) = 43
Num = Num + 1: keyframe6(Num) = 115: Num = Num + 1: keyframe6(Num) = 49
Num = Num + 1: keyframe6(Num) = 112
'--------------- create the database for cel 7 ------------------
Num = 0
keyframe7(Num) = 38
Num = Num + 1: keyframe7(Num) = 19: Num = Num + 1: keyframe7(Num) = 45
Num = Num + 1: keyframe7(Num) = 25: Num = Num + 1: keyframe7(Num) = 48
Num = Num + 1: keyframe7(Num) = 30: Num = Num + 1: keyframe7(Num) = 43
Num = Num + 1: keyframe7(Num) = 31: Num = Num + 1: keyframe7(Num) = 42
Num = Num + 1: keyframe7(Num) = 26: Num = Num + 1: keyframe7(Num) = 51
Num = Num + 1: keyframe7(Num) = 30: Num = Num + 1: keyframe7(Num) = 69
Num = Num + 1: keyframe7(Num) = 62: Num = Num + 1: keyframe7(Num) = 60
Num = Num + 1: keyframe7(Num) = 65: Num = Num + 1: keyframe7(Num) = 49
Num = Num + 1: keyframe7(Num) = 61: Num = Num + 1: keyframe7(Num) = 40
Num = Num + 1: keyframe7(Num) = 32: Num = Num + 1: keyframe7(Num) = 14
Num = Num + 1: keyframe7(Num) = 33: Num = Num + 1: keyframe7(Num) = 15
Num = Num + 1: keyframe7(Num) = 41: Num = Num + 1: keyframe7(Num) = 6
Num = Num + 1: keyframe7(Num) = 42: Num = Num + 1: keyframe7(Num) = 4
Num = Num + 1: keyframe7(Num) = 34: Num = Num + 1: keyframe7(Num) = 15
Num = Num + 1: keyframe7(Num) = 38: Num = Num + 1: keyframe7(Num) = 30
Num = Num + 1: keyframe7(Num) = 45: Num = Num + 1: keyframe7(Num) = 30
Num = Num + 1: keyframe7(Num) = 50: Num = Num + 1: keyframe7(Num) = 15
Num = Num + 1: keyframe7(Num) = 41: Num = Num + 1: keyframe7(Num) = 30
Num = Num + 1: keyframe7(Num) = 45: Num = Num + 1: keyframe7(Num) = 41
Num = Num + 1: keyframe7(Num) = 38: Num = Num + 1: keyframe7(Num) = 42
Num = Num + 1: keyframe7(Num) = 42: Num = Num + 1: keyframe7(Num) = 30
Num = Num + 1: keyframe7(Num) = 50: Num = Num + 1: keyframe7(Num) = 53
Num = Num + 1: keyframe7(Num) = 32: Num = Num + 1: keyframe7(Num) = 73
Num = Num + 1: keyframe7(Num) = 36: Num = Num + 1: keyframe7(Num) = 71
Num = Num + 1: keyframe7(Num) = 39: Num = Num + 1: keyframe7(Num) = 55
Num = Num + 1: keyframe7(Num) = 36: Num = Num + 1: keyframe7(Num) = 73
Num = Num + 1: keyframe7(Num) = 37: Num = Num + 1: keyframe7(Num) = 83
Num = Num + 1: keyframe7(Num) = 53: Num = Num + 1: keyframe7(Num) = 80
Num = Num + 1: keyframe7(Num) = 55: Num = Num + 1: keyframe7(Num) = 71
Num = Num + 1: keyframe7(Num) = 40: Num = Num + 1: keyframe7(Num) = 83
Num = Num + 1: keyframe7(Num) = 53: Num = Num + 1: keyframe7(Num) = 93
Num = Num + 1: keyframe7(Num) = 55: Num = Num + 1: keyframe7(Num) = 89
Num = Num + 1: keyframe7(Num) = 57: Num = Num + 1: keyframe7(Num) = 80
Num = Num + 1: keyframe7(Num) = 55: Num = Num + 1: keyframe7(Num) = 49
Num = Num + 1: keyframe7(Num) = 61: Num = Num + 1: keyframe7(Num) = 60
Num = Num + 1: keyframe7(Num) = 65: Num = Num + 1: keyframe7(Num) = 22
Num = Num + 1: keyframe7(Num) = 65: Num = Num + 1: keyframe7(Num) = 23
Num = Num + 1: keyframe7(Num) = 61: Num = Num + 1: keyframe7(Num) = 31
Num = Num + 1: keyframe7(Num) = 65: Num = Num + 1: keyframe7(Num) = 53
Num = Num + 1: keyframe7(Num) = 81: Num = Num + 1: keyframe7(Num) = 40
Num = Num + 1: keyframe7(Num) = 80: Num = Num + 1: keyframe7(Num) = 23
Num = Num + 1: keyframe7(Num) = 65: Num = Num + 1: keyframe7(Num) = 48
Num = Num + 1: keyframe7(Num) = 77: Num = Num + 1: keyframe7(Num) = 53
Num = Num + 1: keyframe7(Num) = 81: Num = Num + 1: keyframe7(Num) = 27
Num = Num + 1: keyframe7(Num) = 89: Num = Num + 1: keyframe7(Num) = 40
Num = Num + 1: keyframe7(Num) = 79: Num = Num + 1: keyframe7(Num) = 61
Num = Num + 1: keyframe7(Num) = 65: Num = Num + 1: keyframe7(Num) = 67
Num = Num + 1: keyframe7(Num) = 63: Num = Num + 1: keyframe7(Num) = 68
Num = Num + 1: keyframe7(Num) = 84: Num = Num + 1: keyframe7(Num) = 63
Num = Num + 1: keyframe7(Num) = 85: Num = Num + 1: keyframe7(Num) = 68
Num = Num + 1: keyframe7(Num) = 83: Num = Num + 1: keyframe7(Num) = 84
Num = Num + 1: keyframe7(Num) = 99: Num = Num + 1: keyframe7(Num) = 80
Num = Num + 1: keyframe7(Num) = 101: Num = Num + 1: keyframe7(Num) = 63
Num = Num + 1: keyframe7(Num) = 85: Num = Num + 1: keyframe7(Num) = 81
Num = Num + 1: keyframe7(Num) = 98: Num = Num + 1: keyframe7(Num) = 89
Num = Num + 1: keyframe7(Num) = 98: Num = Num + 1: keyframe7(Num) = 85
Num = Num + 1: keyframe7(Num) = 113: Num = Num + 1: keyframe7(Num) = 80
Num = Num + 1: keyframe7(Num) = 111: Num = Num + 1: keyframe7(Num) = 89
Num = Num + 1: keyframe7(Num) = 111: Num = Num + 1: keyframe7(Num) = 85
Num = Num + 1: keyframe7(Num) = 112: Num = Num + 1: keyframe7(Num) = 80
Num = Num + 1: keyframe7(Num) = 115: Num = Num + 1: keyframe7(Num) = 76
Num = Num + 1: keyframe7(Num) = 115
'------------------------ create the cels -----------------------
hMemoryDC = CreateCompatibleDC(Form1.hDC) 'memory-display context
Cel1 = CreateCompatibleBitmap(Form1.hDC, zWIDE, zHIGH)
If Cel1 = Null Then
  MsgBox "Unable to allocate memory for cel 1.", MB_OK, "Unexpected animation condition"
  RetVal = DeleteDC(hMemoryDC)
  RetVal = SelectObject(hFrameDC, hPrevFrame)
  RetVal = DeleteObject(hFrame)
  RetVal = DeleteDC(hFrameDC)
  TimerExists = False
  CelsReady = False
  Selected = False
  Exit Sub
End If
Cel2 = CreateCompatibleBitmap(Form1.hDC, zWIDE, zHIGH)
If Cel2 = Null Then
  MsgBox "Unable to allocate memory for cel 2.", MB_OK, "Unexpected animation condition"
  RetVal = DeleteObject(Cel1)
  RetVal = DeleteDC(hMemoryDC)
  RetVal = SelectObject(hFrameDC, hPrevFrame)
  RetVal = DeleteObject(hFrame)
  RetVal = DeleteDC(hFrameDC)
  TimerExists = False
  CelsReady = False
  Selected = False
  Exit Sub
End If
Cel3 = CreateCompatibleBitmap(Form1.hDC, zWIDE, zHIGH)
If Cel3 = Null Then
  MsgBox "Unable to allocate memory for cel 3.", MB_OK, "Unexpected animation condition"
  RetVal = DeleteObject(Cel1)
  RetVal = DeleteObject(Cel2)
  RetVal = DeleteDC(hMemoryDC)
  RetVal = SelectObject(hFrameDC, hPrevFrame)
  RetVal = DeleteObject(hFrame)
  RetVal = DeleteDC(hFrameDC)
  TimerExists = False
  CelsReady = False
  Selected = False
  Exit Sub
End If
Cel4 = CreateCompatibleBitmap(Form1.hDC, zWIDE, zHIGH)
If Cel4 = Null Then
  MsgBox "Unable to allocate memory for cel 4.", MB_OK, "Unexpected animation condition"
  RetVal = DeleteObject(Cel1)
  RetVal = DeleteObject(Cel2)
  RetVal = DeleteObject(Cel3)
  RetVal = DeleteDC(hMemoryDC)
  RetVal = SelectObject(hFrameDC, hPrevFrame)
  RetVal = DeleteObject(hFrame)
  RetVal = DeleteDC(hFrameDC)
  TimerExists = False
  CelsReady = False
  Selected = False
  Exit Sub
End If
Cel5 = CreateCompatibleBitmap(Form1.hDC, zWIDE, zHIGH)
If Cel5 = Null Then
  MsgBox "Unable to allocate memory for cel 5.", MB_OK, "Unexpected animation condition"
  RetVal = DeleteObject(Cel1)
  RetVal = DeleteObject(Cel2)
  RetVal = DeleteObject(Cel3)
  RetVal = DeleteObject(Cel4)
  RetVal = DeleteDC(hMemoryDC)
  RetVal = SelectObject(hFrameDC, hPrevFrame)
  RetVal = DeleteObject(hFrame)
  RetVal = DeleteDC(hFrameDC)
  TimerExists = False
  CelsReady = False
  Selected = False
  Exit Sub
End If
Cel6 = CreateCompatibleBitmap(Form1.hDC, zWIDE, zHIGH)
If Cel6 = Null Then
  MsgBox "Unable to allocate memory for cel 6.", MB_OK, "Unexpected animation condition"
  RetVal = DeleteObject(Cel1)
  RetVal = DeleteObject(Cel2)
  RetVal = DeleteObject(Cel3)
  RetVal = DeleteObject(Cel4)
  RetVal = DeleteObject(Cel5)
  RetVal = DeleteDC(hMemoryDC)
  RetVal = SelectObject(hFrameDC, hPrevFrame)
  RetVal = DeleteObject(hFrame)
  RetVal = DeleteDC(hFrameDC)
  TimerExists = False
  CelsReady = False
  Selected = False
  Exit Sub
End If
Cel7 = CreateCompatibleBitmap(Form1.hDC, zWIDE, zHIGH)
If Cel7 = Null Then
  MsgBox "Unable to allocate memory for cel 7.", MB_OK, "Unexpected animation condition"
  RetVal = DeleteObject(Cel1)
  RetVal = DeleteObject(Cel2)
  RetVal = DeleteObject(Cel3)
  RetVal = DeleteObject(Cel4)
  RetVal = DeleteObject(Cel5)
  RetVal = DeleteObject(Cel6)
  RetVal = DeleteDC(hMemoryDC)
  RetVal = SelectObject(hFrameDC, hPrevFrame)
  RetVal = DeleteObject(hFrame)
  RetVal = DeleteDC(hFrameDC)
  TimerExists = False
  CelsReady = False
  Selected = False
  Exit Sub
End If
MsgBox "Memory successfully allocated for cels.", MB_OK, "Animation status"

'---------------------- draw the 7 cels -------------------------
ArrayPtr = 1                                'point to first array
zDrawOneCartoon (ArrayPtr)                          'draw the cel
hPrevBitmap = SelectObject(hMemoryDC, Cel1)    'select the bitmap
RetVal = BitBlt(hMemoryDC, 0, 0, zWIDE, zHIGH, Form1.hDC, 0, 0, SRCCOPY)
zBlank                                  'clear the display window
ArrayPtr = 2
zDrawOneCartoon (ArrayPtr)
RetVal = SelectObject(hMemoryDC, Cel2)
RetVal = BitBlt(hMemoryDC, 0, 0, zWIDE, zHIGH, Form1.hDC, 0, 0, SRCCOPY)
zBlank
ArrayPtr = 3
zDrawOneCartoon (ArrayPtr)
RetVal = SelectObject(hMemoryDC, Cel3)
RetVal = BitBlt(hMemoryDC, 0, 0, zWIDE, zHIGH, Form1.hDC, 0, 0, SRCCOPY)
zBlank
ArrayPtr = 4
zDrawOneCartoon (ArrayPtr)
RetVal = SelectObject(hMemoryDC, Cel4)
RetVal = BitBlt(hMemoryDC, 0, 0, zWIDE, zHIGH, Form1.hDC, 0, 0, SRCCOPY)
zBlank
ArrayPtr = 5
zDrawOneCartoon (ArrayPtr)
RetVal = SelectObject(hMemoryDC, Cel5)
RetVal = BitBlt(hMemoryDC, 0, 0, zWIDE, zHIGH, Form1.hDC, 0, 0, SRCCOPY)
zBlank
ArrayPtr = 6
zDrawOneCartoon (ArrayPtr)
RetVal = SelectObject(hMemoryDC, Cel6)
RetVal = BitBlt(hMemoryDC, 0, 0, zWIDE, zHIGH, Form1.hDC, 0, 0, SRCCOPY)
zBlank
ArrayPtr = 7
zDrawOneCartoon (ArrayPtr)
RetVal = SelectObject(hMemoryDC, Cel7)
RetVal = BitBlt(hMemoryDC, 0, 0, zWIDE, zHIGH, Form1.hDC, 0, 0, SRCCOPY)
zBlank
'---------------------- tidy up and return ----------------------
CelsReady = True
RetVal = SelectObject(hMemoryDC, Cel1)
Selected = True
FrameNum = 1
MsgBox "Cel images have been successfully created and stored in memory.", MB_OK, "Animation status"
End Sub

'----------------------------------------------------------------
'                          Draw one cel.
'----------------------------------------------------------------
Sub zDrawOneCartoon (ByVal Num As Integer)
  Dim Repeat As Integer                             'loop counter
  Dim DrawPtr As Integer      'dynamic index into array of coords
  Dim ColorBrush As Integer, PrevBrush As Integer        'brushes
  Dim ColorPen As Integer, PrevPen As Integer               'pens
  Static Poly4(3, 1) As Integer                'array of 4 points
  Static Poly5(4, 1) As Integer                'array of 5 points
ColorPen = CreatePen(PS_SOLID, 1, RGB(0, 128, 0))     'create pen
ColorBrush = CreateSolidBrush(RGB(0, 128, 0))       'create brush
PrevPen = SelectObject(Form1.hDC, ColorPen)           'select pen
PrevBrush = SelectObject(Form1.hDC, ColorBrush)     'select brush
'------------------- branch to appropriate code -----------------
Select Case Num
  Case 1:  GoTo DRAW_CEL_1
  Case 2:  GoTo DRAW_CEL_2
  Case 3:  GoTo DRAW_CEL_3
  Case 4:  GoTo DRAW_CEL_4
  Case 5:  GoTo DRAW_CEL_5
  Case 6:  GoTo DRAW_CEL_6
  Case 7:  GoTo DRAW_CEL_7
End Select
MsgBox "Internal error in Select Case structure in zDrawOneCartoon( ).", MB_OK, "Book 4224 error report"
Exit Sub                                  'trap any coding errors
'--------------------------- draw cel 1 -------------------------
DRAW_CEL_1:
RetVal = Ellipse(Form1.hDC, keyframe1(0) - 10, keyframe1(1) - 10, keyframe1(0) + 10, keyframe1(1) + 10)
DrawPtr = 2                 'update the index into database array
Poly4(0, 0) = keyframe1(DrawPtr): DrawPtr = DrawPtr + 1
Poly4(1, 0) = keyframe1(DrawPtr): DrawPtr = DrawPtr + 1
Poly4(2, 0) = keyframe1(DrawPtr): DrawPtr = DrawPtr + 1
Poly4(3, 0) = keyframe1(DrawPtr): DrawPtr = DrawPtr + 1
Poly4(0, 1) = keyframe1(DrawPtr): DrawPtr = DrawPtr + 1
Poly4(1, 1) = keyframe1(DrawPtr): DrawPtr = DrawPtr + 1
Poly4(2, 1) = keyframe1(DrawPtr): DrawPtr = DrawPtr + 1
Poly4(3, 1) = keyframe1(DrawPtr): DrawPtr = DrawPtr + 1
RetVal = Polygon(Form1.hDC, Poly4(0, 0), 4)        'draw the neck
Poly5(0, 0) = keyframe1(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(1, 0) = keyframe1(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(2, 0) = keyframe1(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(3, 0) = keyframe1(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(4, 0) = keyframe1(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(0, 1) = keyframe1(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(1, 1) = keyframe1(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(2, 1) = keyframe1(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(3, 1) = keyframe1(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(4, 1) = keyframe1(DrawPtr): DrawPtr = DrawPtr + 1
RetVal = Polygon(Form1.hDC, Poly5(0, 0), 5)       'draw the torso
For Repeat = 1 To 13
  Poly4(0, 0) = keyframe1(DrawPtr): DrawPtr = DrawPtr + 1
  Poly4(1, 0) = keyframe1(DrawPtr): DrawPtr = DrawPtr + 1
  Poly4(2, 0) = keyframe1(DrawPtr): DrawPtr = DrawPtr + 1
  Poly4(3, 0) = keyframe1(DrawPtr): DrawPtr = DrawPtr + 1
  Poly4(0, 1) = keyframe1(DrawPtr): DrawPtr = DrawPtr + 1
  Poly4(1, 1) = keyframe1(DrawPtr): DrawPtr = DrawPtr + 1
  Poly4(2, 1) = keyframe1(DrawPtr): DrawPtr = DrawPtr + 1
  Poly4(3, 1) = keyframe1(DrawPtr): DrawPtr = DrawPtr + 1
  RetVal = Polygon(Form1.hDC, Poly4(0, 0), 4)'draw remaining parts
Next Repeat
GoTo All_DONE
'--------------------------- draw cel 2 -------------------------
DRAW_CEL_2:
RetVal = Ellipse(Form1.hDC, keyframe2(0) - 10, keyframe2(1) - 10, keyframe2(0) + 10, keyframe2(1) + 10)
DrawPtr = 2
Poly4(0, 0) = keyframe2(DrawPtr): DrawPtr = DrawPtr + 1
Poly4(1, 0) = keyframe2(DrawPtr): DrawPtr = DrawPtr + 1
Poly4(2, 0) = keyframe2(DrawPtr): DrawPtr = DrawPtr + 1
Poly4(3, 0) = keyframe2(DrawPtr): DrawPtr = DrawPtr + 1
Poly4(0, 1) = keyframe2(DrawPtr): DrawPtr = DrawPtr + 1
Poly4(1, 1) = keyframe2(DrawPtr): DrawPtr = DrawPtr + 1
Poly4(2, 1) = keyframe2(DrawPtr): DrawPtr = DrawPtr + 1
Poly4(3, 1) = keyframe2(DrawPtr): DrawPtr = DrawPtr + 1
RetVal = Polygon(Form1.hDC, Poly4(0, 0), 4)
Poly5(0, 0) = keyframe2(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(1, 0) = keyframe2(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(2, 0) = keyframe2(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(3, 0) = keyframe2(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(4, 0) = keyframe2(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(0, 1) = keyframe2(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(1, 1) = keyframe2(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(2, 1) = keyframe2(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(3, 1) = keyframe2(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(4, 1) = keyframe2(DrawPtr): DrawPtr = DrawPtr + 1
RetVal = Polygon(Form1.hDC, Poly5(0, 0), 5)
For Repeat = 1 To 13
  Poly4(0, 0) = keyframe2(DrawPtr): DrawPtr = DrawPtr + 1
  Poly4(1, 0) = keyframe2(DrawPtr): DrawPtr = DrawPtr + 1
  Poly4(2, 0) = keyframe2(DrawPtr): DrawPtr = DrawPtr + 1
  Poly4(3, 0) = keyframe2(DrawPtr): DrawPtr = DrawPtr + 1
  Poly4(0, 1) = keyframe2(DrawPtr): DrawPtr = DrawPtr + 1
  Poly4(1, 1) = keyframe2(DrawPtr): DrawPtr = DrawPtr + 1
  Poly4(2, 1) = keyframe2(DrawPtr): DrawPtr = DrawPtr + 1
  Poly4(3, 1) = keyframe2(DrawPtr): DrawPtr = DrawPtr + 1
  RetVal = Polygon(Form1.hDC, Poly4(0, 0), 4)
Next Repeat
GoTo All_DONE
'--------------------------- draw cel 3 -------------------------
DRAW_CEL_3:
RetVal = Ellipse(Form1.hDC, keyframe3(0) - 10, keyframe3(1) - 10, keyframe3(0) + 10, keyframe3(1) + 10)
DrawPtr = 2
Poly4(0, 0) = keyframe3(DrawPtr): DrawPtr = DrawPtr + 1
Poly4(1, 0) = keyframe3(DrawPtr): DrawPtr = DrawPtr + 1
Poly4(2, 0) = keyframe3(DrawPtr): DrawPtr = DrawPtr + 1
Poly4(3, 0) = keyframe3(DrawPtr): DrawPtr = DrawPtr + 1
Poly4(0, 1) = keyframe3(DrawPtr): DrawPtr = DrawPtr + 1
Poly4(1, 1) = keyframe3(DrawPtr): DrawPtr = DrawPtr + 1
Poly4(2, 1) = keyframe3(DrawPtr): DrawPtr = DrawPtr + 1
Poly4(3, 1) = keyframe3(DrawPtr): DrawPtr = DrawPtr + 1
RetVal = Polygon(Form1.hDC, Poly4(0, 0), 4)
Poly5(0, 0) = keyframe3(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(1, 0) = keyframe3(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(2, 0) = keyframe3(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(3, 0) = keyframe3(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(4, 0) = keyframe3(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(0, 1) = keyframe3(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(1, 1) = keyframe3(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(2, 1) = keyframe3(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(3, 1) = keyframe3(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(4, 1) = keyframe3(DrawPtr): DrawPtr = DrawPtr + 1
RetVal = Polygon(Form1.hDC, Poly5(0, 0), 5)
For Repeat = 1 To 13
  Poly4(0, 0) = keyframe3(DrawPtr): DrawPtr = DrawPtr + 1
  Poly4(1, 0) = keyframe3(DrawPtr): DrawPtr = DrawPtr + 1
  Poly4(2, 0) = keyframe3(DrawPtr): DrawPtr = DrawPtr + 1
  Poly4(3, 0) = keyframe3(DrawPtr): DrawPtr = DrawPtr + 1
  Poly4(0, 1) = keyframe3(DrawPtr): DrawPtr = DrawPtr + 1
  Poly4(1, 1) = keyframe3(DrawPtr): DrawPtr = DrawPtr + 1
  Poly4(2, 1) = keyframe3(DrawPtr): DrawPtr = DrawPtr + 1
  Poly4(3, 1) = keyframe3(DrawPtr): DrawPtr = DrawPtr + 1
  RetVal = Polygon(Form1.hDC, Poly4(0, 0), 4)
Next Repeat
GoTo All_DONE
'--------------------------- draw cel 4 -------------------------
DRAW_CEL_4:
RetVal = Ellipse(Form1.hDC, keyframe4(0) - 10, keyframe4(1) - 10, keyframe4(0) + 10, keyframe4(1) + 10)
DrawPtr = 2
Poly4(0, 0) = keyframe4(DrawPtr): DrawPtr = DrawPtr + 1
Poly4(1, 0) = keyframe4(DrawPtr): DrawPtr = DrawPtr + 1
Poly4(2, 0) = keyframe4(DrawPtr): DrawPtr = DrawPtr + 1
Poly4(3, 0) = keyframe4(DrawPtr): DrawPtr = DrawPtr + 1
Poly4(0, 1) = keyframe4(DrawPtr): DrawPtr = DrawPtr + 1
Poly4(1, 1) = keyframe4(DrawPtr): DrawPtr = DrawPtr + 1
Poly4(2, 1) = keyframe4(DrawPtr): DrawPtr = DrawPtr + 1
Poly4(3, 1) = keyframe4(DrawPtr): DrawPtr = DrawPtr + 1
RetVal = Polygon(Form1.hDC, Poly4(0, 0), 4)
Poly5(0, 0) = keyframe4(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(1, 0) = keyframe4(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(2, 0) = keyframe4(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(3, 0) = keyframe4(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(4, 0) = keyframe4(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(0, 1) = keyframe4(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(1, 1) = keyframe4(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(2, 1) = keyframe4(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(3, 1) = keyframe4(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(4, 1) = keyframe4(DrawPtr): DrawPtr = DrawPtr + 1
RetVal = Polygon(Form1.hDC, Poly5(0, 0), 5)
For Repeat = 1 To 13
  Poly4(0, 0) = keyframe4(DrawPtr): DrawPtr = DrawPtr + 1
  Poly4(1, 0) = keyframe4(DrawPtr): DrawPtr = DrawPtr + 1
  Poly4(2, 0) = keyframe4(DrawPtr): DrawPtr = DrawPtr + 1
  Poly4(3, 0) = keyframe4(DrawPtr): DrawPtr = DrawPtr + 1
  Poly4(0, 1) = keyframe4(DrawPtr): DrawPtr = DrawPtr + 1
  Poly4(1, 1) = keyframe4(DrawPtr): DrawPtr = DrawPtr + 1
  Poly4(2, 1) = keyframe4(DrawPtr): DrawPtr = DrawPtr + 1
  Poly4(3, 1) = keyframe4(DrawPtr): DrawPtr = DrawPtr + 1
  RetVal = Polygon(Form1.hDC, Poly4(0, 0), 4)
Next Repeat
GoTo All_DONE
'--------------------------- draw cel 5 -------------------------
DRAW_CEL_5:
RetVal = Ellipse(Form1.hDC, keyframe5(0) - 10, keyframe5(1) - 10, keyframe5(0) + 10, keyframe5(1) + 10)
DrawPtr = 2
Poly4(0, 0) = keyframe5(DrawPtr): DrawPtr = DrawPtr + 1
Poly4(1, 0) = keyframe5(DrawPtr): DrawPtr = DrawPtr + 1
Poly4(2, 0) = keyframe5(DrawPtr): DrawPtr = DrawPtr + 1
Poly4(3, 0) = keyframe5(DrawPtr): DrawPtr = DrawPtr + 1
Poly4(0, 1) = keyframe5(DrawPtr): DrawPtr = DrawPtr + 1
Poly4(1, 1) = keyframe5(DrawPtr): DrawPtr = DrawPtr + 1
Poly4(2, 1) = keyframe5(DrawPtr): DrawPtr = DrawPtr + 1
Poly4(3, 1) = keyframe5(DrawPtr): DrawPtr = DrawPtr + 1
RetVal = Polygon(Form1.hDC, Poly4(0, 0), 4)
Poly5(0, 0) = keyframe5(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(1, 0) = keyframe5(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(2, 0) = keyframe5(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(3, 0) = keyframe5(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(4, 0) = keyframe5(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(0, 1) = keyframe5(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(1, 1) = keyframe5(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(2, 1) = keyframe5(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(3, 1) = keyframe5(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(4, 1) = keyframe5(DrawPtr): DrawPtr = DrawPtr + 1
RetVal = Polygon(Form1.hDC, Poly5(0, 0), 5)
For Repeat = 1 To 13
  Poly4(0, 0) = keyframe5(DrawPtr): DrawPtr = DrawPtr + 1
  Poly4(1, 0) = keyframe5(DrawPtr): DrawPtr = DrawPtr + 1
  Poly4(2, 0) = keyframe5(DrawPtr): DrawPtr = DrawPtr + 1
  Poly4(3, 0) = keyframe5(DrawPtr): DrawPtr = DrawPtr + 1
  Poly4(0, 1) = keyframe5(DrawPtr): DrawPtr = DrawPtr + 1
  Poly4(1, 1) = keyframe5(DrawPtr): DrawPtr = DrawPtr + 1
  Poly4(2, 1) = keyframe5(DrawPtr): DrawPtr = DrawPtr + 1
  Poly4(3, 1) = keyframe5(DrawPtr): DrawPtr = DrawPtr + 1
  RetVal = Polygon(Form1.hDC, Poly4(0, 0), 4)
Next Repeat
GoTo All_DONE
'--------------------------- draw cel 6 -------------------------
DRAW_CEL_6:
RetVal = Ellipse(Form1.hDC, keyframe6(0) - 10, keyframe6(1) - 10, keyframe6(0) + 10, keyframe6(1) + 10)
DrawPtr = 2
Poly4(0, 0) = keyframe6(DrawPtr): DrawPtr = DrawPtr + 1
Poly4(1, 0) = keyframe6(DrawPtr): DrawPtr = DrawPtr + 1
Poly4(2, 0) = keyframe6(DrawPtr): DrawPtr = DrawPtr + 1
Poly4(3, 0) = keyframe6(DrawPtr): DrawPtr = DrawPtr + 1
Poly4(0, 1) = keyframe6(DrawPtr): DrawPtr = DrawPtr + 1
Poly4(1, 1) = keyframe6(DrawPtr): DrawPtr = DrawPtr + 1
Poly4(2, 1) = keyframe6(DrawPtr): DrawPtr = DrawPtr + 1
Poly4(3, 1) = keyframe6(DrawPtr): DrawPtr = DrawPtr + 1
RetVal = Polygon(Form1.hDC, Poly4(0, 0), 4)
Poly5(0, 0) = keyframe6(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(1, 0) = keyframe6(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(2, 0) = keyframe6(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(3, 0) = keyframe6(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(4, 0) = keyframe6(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(0, 1) = keyframe6(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(1, 1) = keyframe6(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(2, 1) = keyframe6(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(3, 1) = keyframe6(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(4, 1) = keyframe6(DrawPtr): DrawPtr = DrawPtr + 1
RetVal = Polygon(Form1.hDC, Poly5(0, 0), 5)
For Repeat = 1 To 13
  Poly4(0, 0) = keyframe6(DrawPtr): DrawPtr = DrawPtr + 1
  Poly4(1, 0) = keyframe6(DrawPtr): DrawPtr = DrawPtr + 1
  Poly4(2, 0) = keyframe6(DrawPtr): DrawPtr = DrawPtr + 1
  Poly4(3, 0) = keyframe6(DrawPtr): DrawPtr = DrawPtr + 1
  Poly4(0, 1) = keyframe6(DrawPtr): DrawPtr = DrawPtr + 1
  Poly4(1, 1) = keyframe6(DrawPtr): DrawPtr = DrawPtr + 1
  Poly4(2, 1) = keyframe6(DrawPtr): DrawPtr = DrawPtr + 1
  Poly4(3, 1) = keyframe6(DrawPtr): DrawPtr = DrawPtr + 1
  RetVal = Polygon(Form1.hDC, Poly4(0, 0), 4)
Next Repeat
GoTo All_DONE
'--------------------------- draw cel 7 -------------------------
DRAW_CEL_7:
RetVal = Ellipse(Form1.hDC, keyframe7(0) - 10, keyframe7(1) - 10, keyframe7(0) + 10, keyframe7(1) + 10)
DrawPtr = 2
Poly4(0, 0) = keyframe7(DrawPtr): DrawPtr = DrawPtr + 1
Poly4(1, 0) = keyframe7(DrawPtr): DrawPtr = DrawPtr + 1
Poly4(2, 0) = keyframe7(DrawPtr): DrawPtr = DrawPtr + 1
Poly4(3, 0) = keyframe7(DrawPtr): DrawPtr = DrawPtr + 1
Poly4(0, 1) = keyframe7(DrawPtr): DrawPtr = DrawPtr + 1
Poly4(1, 1) = keyframe7(DrawPtr): DrawPtr = DrawPtr + 1
Poly4(2, 1) = keyframe7(DrawPtr): DrawPtr = DrawPtr + 1
Poly4(3, 1) = keyframe7(DrawPtr): DrawPtr = DrawPtr + 1
RetVal = Polygon(Form1.hDC, Poly4(0, 0), 4)
Poly5(0, 0) = keyframe7(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(1, 0) = keyframe7(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(2, 0) = keyframe7(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(3, 0) = keyframe7(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(4, 0) = keyframe7(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(0, 1) = keyframe7(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(1, 1) = keyframe7(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(2, 1) = keyframe7(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(3, 1) = keyframe7(DrawPtr): DrawPtr = DrawPtr + 1
Poly5(4, 1) = keyframe7(DrawPtr): DrawPtr = DrawPtr + 1
RetVal = Polygon(Form1.hDC, Poly5(0, 0), 5)
For Repeat = 1 To 13
  Poly4(0, 0) = keyframe7(DrawPtr): DrawPtr = DrawPtr + 1
  Poly4(1, 0) = keyframe7(DrawPtr): DrawPtr = DrawPtr + 1
  Poly4(2, 0) = keyframe7(DrawPtr): DrawPtr = DrawPtr + 1
  Poly4(3, 0) = keyframe7(DrawPtr): DrawPtr = DrawPtr + 1
  Poly4(0, 1) = keyframe7(DrawPtr): DrawPtr = DrawPtr + 1
  Poly4(1, 1) = keyframe7(DrawPtr): DrawPtr = DrawPtr + 1
  Poly4(2, 1) = keyframe7(DrawPtr): DrawPtr = DrawPtr + 1
  Poly4(3, 1) = keyframe7(DrawPtr): DrawPtr = DrawPtr + 1
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
'                    REAL-TIME ANIMATION ENGINE.
' This function builds the next frame of the animation sequence
' on a hidden bitmap.  When completed, the finished frame is
' then copied to the display window.
'----------------------------------------------------------------
Sub zShowNextFrame ()              'uses global variable FrameNum
If CelsReady = False Then               'cancel if cels not ready
  Exit Sub
End If
If Pause = True Then                  'cancel if animation paused
  Exit Sub
End If
If FrameDirection = zFORWARD Then
  FrameNum = FrameNum + 1             'increment the frame number
  If FrameNum > zFINALFRAME Then
    FrameNum = zFIRSTFRAME
  End If
  SX1 = SX1 - zMOVEHORIZ                 'adjust the cel position
  If SX1 < zMAXLEFT Then
    SX1 = zMAXRIGHT
  End If
End If
If FrameDirection = zREVERSE Then
  FrameNum = FrameNum - 1             'decrement the frame number
  If FrameNum < zFIRSTFRAME Then
    FrameNum = zFINALFRAME
  End If
  SX1 = SX1 + zMOVEHORIZ                 'adjust the cel position
  If SX1 > zMAXRIGHT Then
    SX1 = zMAXLEFT
  End If
End If

DISPLAY_FRAME:
'------------ draw the next image on the hidden frame -----------
RetVal = PatBlt(hFrameDC, 0, 0, zFRAMEWIDE, zFRAMEHIGH, WHITENESS)
Select Case FrameNum
  Case 1:  RetVal = SelectObject(hMemoryDC, Cel1)
           RetVal = BitBlt(hFrameDC, SX1, SY1, zWIDE, zHIGH, hMemoryDC, 0, 0, SRCCOPY)
  Case 2:  RetVal = SelectObject(hMemoryDC, Cel2)
           RetVal = BitBlt(hFrameDC, SX1, SY1, zWIDE, zHIGH, hMemoryDC, 0, 0, SRCCOPY)
  Case 3:  RetVal = SelectObject(hMemoryDC, Cel3)
           RetVal = BitBlt(hFrameDC, SX1, SY1, zWIDE, zHIGH, hMemoryDC, 0, 0, SRCCOPY)
  Case 4:  RetVal = SelectObject(hMemoryDC, Cel4)
           RetVal = BitBlt(hFrameDC, SX1, SY1, zWIDE, zHIGH, hMemoryDC, 0, 0, SRCCOPY)
  Case 5:  RetVal = SelectObject(hMemoryDC, Cel5)
           RetVal = BitBlt(hFrameDC, SX1, SY1, zWIDE, zHIGH, hMemoryDC, 0, 0, SRCCOPY)
  Case 6:  RetVal = SelectObject(hMemoryDC, Cel6)
           RetVal = BitBlt(hFrameDC, SX1, SY1, zWIDE, zHIGH, hMemoryDC, 0, 0, SRCCOPY)
  Case 7:  RetVal = SelectObject(hMemoryDC, Cel7)
           RetVal = BitBlt(hFrameDC, SX1, SY1, zWIDE, zHIGH, hMemoryDC, 0, 0, SRCCOPY)
End Select
'------------ copy hidden frame to the display window -----------
RetVal = BitBlt(Form1.hDC, 0, 0, zFRAMEWIDE, zFRAMEHIGH, hFrameDC, 0, 0, SRCCOPY)
End Sub

'----------------------------------------------------------------
'                  Blank the display window
'----------------------------------------------------------------
Sub zBlank ()
RetVal = PatBlt(Form1.hDC, 0, 0, zFRAMEWIDE, zFRAMEHIGH, WHITENESS)
End Sub

