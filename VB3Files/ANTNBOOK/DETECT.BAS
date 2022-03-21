'----------------------------------------------------------------
'  Reusable global module for Visual Basic graphics applications
'              that call Windows API functions.
'----------------------------------------------------------------
'  Source file:  DETECT.BAS
'  Release version:  1.00                  Programmer:  Lee Adams
'  Type:  Visual Basic global module for Windows applications.
'  Compiler:  Microsoft Visual Basic 1.00
'  Dependencies:  DETECT.FRM primary module
'                 DETECT01.BAS module containing Main()
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

'----------------------------------------------------------------
'                   Hidden frame operations
'----------------------------------------------------------------
Global hFrameDC As Integer      'display-context for hidden-frame
Global hFrame As Integer           'handle to hidden-frame bitmap
Global hPrevFrame As Integer         'default bitmap for hFrameDC
Global FrameReady As Integer               'hidden-frame created?

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
Declare Function CreatePolygonRgn Lib "GDI" (FirstPt As Integer, ByVal Count As Integer, ByVal PolyFillMode As Integer) As Integer

'----------------------------------------------------------------
'   GDI, USER, KERNEL functions for various runtime conditions
'----------------------------------------------------------------
Declare Function GetDeviceCaps Lib "GDI" (ByVal hDC As Integer, ByVal Item As Integer) As Integer
Declare Function GetWinFlags Lib "KERNEL" () As Long
Declare Function GetSystemMetrics Lib "USER" (ByVal Item As Integer) As Integer
Declare Function GlobalCompact Lib "KERNEL" (ByVal NumBytes As Long) As Long

'----------------------------------------------------------------
'                     End of global module.
'----------------------------------------------------------------

