'----------------------------------------------------------------
'  Reusable global module for Visual Basic graphics applications
'              that call Windows API functions.
'----------------------------------------------------------------
'  Source file:  BLINK.BAS
'  Release version:  1.00                  Programmer:  Lee Adams
'  Type:  Visual Basic global module for Windows applications.
'  Compiler:  Microsoft Visual Basic 1.00
'  Dependencies:  BLINK.FRM primary module
'                 BLINK01.BAS module contains Main()
'                 BLINK02.BAS module contains animation routines
'  Output and features:  Provides declarations for Windows API
'    (Application Programming Interface) functions callable by
'    Visual Basic applications at runtime, including routines
'    from Windows' GDI, USER, and KERNEL DLLs (dynamic link
'    libraries).  Also provides declarations of various variables
'    and constants.  Functions, variables, and constants declared
'    in this global module are visible throughout the project.
'  Publication:  Contains material from Windcrest/McGraw-Hill
'    book 4224 published by TAB BOOKS Div. of McGraw-Hill Inc.
'  License:  As purchaser of the book you are granted a
'    royalty-free license to distribute executable files
'    generated using this code provided that you accept the
'    conditions of the License Agreement and Limited Warranty
'    described in the book and on the companion disk.  Government
'    users:  This software and documentation are subject to
'    restrictions set forth in The Rights in Technical Data and
'    Computer Software clause at 252.227-7013 and elsewhere.
'----------------------------------------------------------------
'      (c) Copyright 1992 Lee Adams.  All rights reserved.
'           Lee Adams(tm) is a trademark of Lee Adams.
'----------------------------------------------------------------

'----------------------------------------------------------------
'              General constants and variables
'----------------------------------------------------------------
Global Const True = -1                          'boolean constant
Global Const False = 0                          'boolean constant
Global Const MB_OK = 0                 'OK button for message box
Global Const PIXELS = 3                        'pixel coordinates
Global Const Null = 0                                  'undefined
Global StartUp As Integer                'tracks the startup code
Global IgnoreRefresh As Integer          'tracks refresh activity

'----------------------------------------------------------------
'                   Window specifications
'----------------------------------------------------------------
Global Const zWINDOW_WIDTH = 408                 'width of window
Global Const zWINDOW_HEIGHT = 346               'height of window
Global Const zFRAMEWIDE = 400               'width of client area
Global Const zFRAMEHIGH = 300              'height of client area
Global HorizTwipsPixel As Single    'twips-per-pixel (horizontal)
Global VertTwipsPixel As Single       'twips-per-pixel (vertical)
Global Window_Width As Single            'runtime width of window
Global Window_Height As Single          'runtime height of window

'----------------------------------------------------------------
'                     Runtime conditions
'----------------------------------------------------------------
Global hDesktopWnd As Integer                  'handle to desktop
Global hDCcaps As Integer            'display-context for desktop
Global DisplayWidth As Integer      'horizontal screen resolution
Global DisplayHeight As Integer       'vertical screen resolution
Global DisplayBits As Integer           'number of bits-per-pixel
Global DisplayPlanes As Integer              'number of bitplanes
Global MemoryMode As Long                    'runtime memory mode
Global RetVal As Integer           'will receive GDI return value
Global RetLong As Long             'will receive GDI return value
Global MousePresent As Integer                     'mouse active?
Global LoadingFrame As Integer                 'loading an image?

'----------------------------------------------------------------
'                    Hidden frame operations
'----------------------------------------------------------------
Global hFrameDC As Integer      'display-context for hidden-frame
Global hFrame As Integer           'handle to hidden-frame bitmap
Global hPrevFrame As Integer         'default bitmap for hFrameDC
Global FrameReady As Integer               'hidden-frame created?

'----------------------------------------------------------------
'                      Timer operations
'----------------------------------------------------------------
Global Const zTIMER_PAUSE = 3                  'for slow machines
Global TimerCounter As Integer                 'for slow machines
Global TimerExists As Integer                   'timer activated?

'----------------------------------------------------------------
'                         Drawings
'----------------------------------------------------------------
Global Cel1(88) As Integer                'database for drawing 1
Global Cel2(92) As Integer                'database for drawing 2
Global Cel3(80) As Integer                'database for drawing 3
Global Cel4(55) As Integer                'database for drawing 4
Global Cel5(51) As Integer                'database for drawing 5
Global Cel6(78) As Integer                'database for drawing 6
Global Cel7(70) As Integer                'database for drawing 7
Global Cel8(72) As Integer                'database for drawing 8
Global Cel9(89) As Integer                'database for drawing 9
Global Cel10(111) As Integer             'database for drawing 10
Global Cel11(106) As Integer             'database for drawing 11
Global Cel12(116) As Integer             'database for drawing 12
Global Cel13(112) As Integer             'database for drawing 13
Global Cel14(88) As Integer              'database for drawing 14
Global Cel15(76) As Integer              'database for drawing 15

'----------------------------------------------------------------
'                       Animation engine
'----------------------------------------------------------------
Global Pause As Integer                       'animation running?
Global wFrameRate As Long             'arbitrary rate of 18.2 fps
Global Const zFORWARD = 1            'indicates forward animation
Global Const zREVERSE = 0            'indicates reverse animation
Global FrameDirection As Integer              'forward or reverse
Global FrameNum As Integer              'current frame identifier
Global Const zFIRSTFRAME = 1   'first frame in animation sequence
Global Const zFINALFRAME = 67  'final frame in animation sequence

'----------------------------------------------------------------
'                     Variables for disk I/O
'----------------------------------------------------------------
Global FileName As String              'name of binary image file
Global DrawingNum As Integer                'which drawing to use
Global hFDC As Integer       'memory display-context for playback
Global hF1 As Integer        'handles to bitmaps for playback...
Global hF2 As Integer, hF3 As Integer, hF4 As Integer
Global hF5 As Integer, hF6 As Integer, hF7 As Integer
Global hF8 As Integer, hF9 As Integer, hF10 As Integer
Global hF11 As Integer, hF12 As Integer, hF13 As Integer
Global hF14 As Integer, hF15 As Integer
Global hPrevF As Integer                          'default bitmap
Global FrameSaved As Integer                'frame saved to disk?
Global FrameLoaded As Integer            'frame loaded from disk?
Global AnimationSaved As Integer        'animation saved to disk?
Global AnimationLoaded As Integer    'animation loaded from disk?
Global AnimationReady As Integer   'animation ready for playback?
Global PrevSaveAttempt As Integer    'previous save attempt made?
Global PrevLoadAttempt As Integer    'previous load attempt made?
Global UseDisk As Integer             'load each frame as needed?
Global AnimationHalted As Integer   'disk error during animation?

'----------------------------------------------------------------
'              Constants for Windows API functions
'----------------------------------------------------------------
Global Const SRCCOPY = &HCC0020                   'for bitblts...
Global Const SRCINVERT = &H660046
Global Const SRCPAINT = &HEE0086
Global Const WHITENESS = &HFF0062
Global Const BLACKNESS = &H42&
Global Const ALTERNATE = 1                        'for filling...
Global Const WINDING = 2
Global Const R2_COPYPEN = 13                     'for pen mode...
Global Const R2_XORPEN = 7
Global Const TRANSPARENT = 1              'for background mode...
Global Const OPAQUE = 2
Global Const PS_SOLID = 0                          'for solid pen
Global Const PS_NULL = 5                     'for transparent pen
Global Const BLACKONWHITE = 1              'for bitblt scaling...
Global Const WHITEONBLACK = 2
Global Const COLORONCOLOR = 3
Global Const HORZRES = 8             'args for GetDeviceCaps()...
Global Const VERTRES = 10
Global Const BITSPIXEL = 12
Global Const PLANES = 14
Global Const SM_MOUSEPRESENT = 19         'for GetSystemMetrics()
Global Const WF_ENHANCED = &H20             'for GetWinFlags()...
Global Const WF_STANDARD = &H10

'----------------------------------------------------------------
'             GDI functions for display-contexts
'----------------------------------------------------------------
Declare Function GetDC Lib "USER" (ByVal hWnd As Integer) As Integer
Declare Function ReleaseDC Lib "USER" (ByVal hWnd As Integer, ByVal hDC As Integer) As Integer

'----------------------------------------------------------------
'               GDI functions for the desktop
'----------------------------------------------------------------
Declare Function GetDesktopWindow Lib "USER" () As Integer

'----------------------------------------------------------------
'        GDI functions for creating drawing objects
'----------------------------------------------------------------
Declare Function CreatePen Lib "GDI" (ByVal PenStyle As Integer, ByVal Wd As Integer, ByVal Color As Long) As Integer
Declare Function CreateSolidBrush Lib "GDI" (ByVal Color As Long) As Integer

'----------------------------------------------------------------
'            GDI functions for selecting objects
'----------------------------------------------------------------
Declare Function SelectObject Lib "GDI" (ByVal hDC As Integer, ByVal hObject As Integer) As Integer
Declare Function DeleteObject Lib "GDI" (ByVal hObject As Integer) As Integer

'----------------------------------------------------------------
'            GDI functions for bitmaps and bitblts
'----------------------------------------------------------------
Declare Function CreateCompatibleDC Lib "GDI" (ByVal hDC As Integer) As Integer
Declare Function CreateCompatibleBitmap Lib "GDI" (ByVal hDC As Integer, ByVal Wd As Integer, ByVal Ht As Integer) As Integer
Declare Function PatBlt Lib "GDI" (ByVal hDC As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal Wd As Integer, ByVal Ht As Integer, ByVal RasOp As Long) As Integer
Declare Function BitBlt Lib "GDI" (ByVal hDestDC As Integer, ByVal DestX As Integer, ByVal DestY As Integer, ByVal Wd As Integer, ByVal Ht As Integer, ByVal hSrcDC As Integer, ByVal SrcX As Integer, ByVal SrcY As Integer, ByVal RastOp As Long) As Integer
Declare Function DeleteDC Lib "GDI" (ByVal hDC As Integer) As Integer

'----------------------------------------------------------------
'           GDI functions for drawing mode operations
'----------------------------------------------------------------
Declare Function SetROP2 Lib "GDI" (ByVal hDC As Integer, ByVal RasMode As Integer) As Integer
Declare Function SetBkColor Lib "GDI" (ByVal hDC As Integer, ByVal Color As Long) As Long
Declare Function SetBkMode Lib "GDI" (ByVal hDC As Integer, ByVal BkMode As Integer) As Integer
Declare Function SetPolyFillMode Lib "GDI" (ByVal hDC As Integer, ByVal PolyMode As Integer) As Integer
Declare Function SetStretchBltMode Lib "GDI" (ByVal hDC As Integer, ByVal StretchMode As Integer) As Integer

'----------------------------------------------------------------
'             GDI functions for drawing operations
'----------------------------------------------------------------
Declare Function MoveTo Lib "GDI" (ByVal hDC As Integer, ByVal X As Integer, ByVal Y As Integer) As Long
Declare Function LineTo Lib "GDI" (ByVal hDC As Integer, ByVal X As Integer, ByVal Y As Integer) As Integer
Declare Function Polygon Lib "GDI" (ByVal hDC As Integer, FirstPt As Integer, ByVal Count As Integer) As Integer
Declare Function Rectangle Lib "GDI" (ByVal hDC As Integer, ByVal X1 As Integer, ByVal Y1 As Integer, ByVal X2 As Integer, ByVal Y2 As Integer) As Integer
Declare Function Ellipse Lib "GDI" (ByVal hDC As Integer, ByVal X1 As Integer, ByVal Y1 As Integer, ByVal X2 As Integer, ByVal Y2 As Integer) As Integer
Declare Function FloodFill Lib "GDI" (ByVal hDC As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal Color As Long) As Integer
Declare Function SetPixel Lib "GDI" (ByVal hDC As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal Color As Long) As Long
Declare Function GetPixel Lib "GDI" (ByVal hDC As Integer, ByVal X As Integer, ByVal Y As Integer) As Long

'----------------------------------------------------------------
'                 GDI functions for regions
'----------------------------------------------------------------
Declare Function PaintRgn Lib "GDI" (ByVal hDC As Integer, ByVal hRgn As Integer) As Integer
Declare Function CreatePolygonRgn Lib "GDI" (FirstPoint As Integer, ByVal Count As Integer, ByVal PolyFillMode As Integer) As Integer
Declare Function SetViewportOrg Lib "GDI" (ByVal hDC As Integer, ByVal X As Integer, ByVal Y As Integer) As Long

'----------------------------------------------------------------
'   GDI, USER, KERNEL functions for various runtime conditions
'----------------------------------------------------------------
Declare Function GetDeviceCaps Lib "GDI" (ByVal hDC As Integer, ByVal Item As Integer) As Integer
Declare Function GetWinFlags Lib "KERNEL" () As Long
Declare Function GetSystemMetrics Lib "USER" (ByVal Item As Integer) As Integer
Declare Function GlobalCompact Lib "KERNEL" (ByVal NumBytes As Long) As Long
Declare Function SetCapture Lib "USER" (ByVal hWnd As Integer) As Integer
Declare Sub ReleaseCapture Lib "USER" ()

'----------------------------------------------------------------
'                     End of global module.
'----------------------------------------------------------------

