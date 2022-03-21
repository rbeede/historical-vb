'----------------------------------------------------------------
'      3D routines for Visual Basic Windows applications
'----------------------------------------------------------------
' Source file:  LIB3D.BAS
' Release version:  1.00                   Programmer:  Lee Adams
' Type:  Visual Basic functional module for Windows applications
' Compiler:  Microsoft Visual Basic 1.00
' Output and features:  Provides routines to generate 3D solids
'   with automatic hidden surface removal uSing the backplane
'   algorithm.  Each facet of a solid is shaded according to
'   a point light source.  The camera position and direction can
'   be adjusted independently of the subject position and
'   orientation.
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
Const zMAX_LOCATION = 100                        'max translation
Const zMAX_HEADING = 360                      'max camera heading
Const zMIN_HEADING = 0                        'min camera heading
Const zMAX_PITCH = 360                          'max camera pitch
Const zMIN_PITCH = 270                          'min camera pitch
Const zMAX_DISTANCE = 700         'max camera-to-subject distance
Const zMIN_DISTANCE = 100         'min camera-to-subject distance

'----------------------------------------------------------------
'           Declaration of variables for 3D modeling
'----------------------------------------------------------------
Dim DomainWidth As Double, DomainDepth As Double
Dim x As Double, y As Double, z As Double
Dim xc1 As Double, xc2 As Double, xc3 As Double, xc4 As Double
Dim xc5 As Double, xc6 As Double, xc7 As Double
Dim yc1 As Double, yc2 As Double, yc3 As Double, yc4 As Double
Dim yc5 As Double, yc6 As Double, yc7 As Double
Dim zc1 As Double, zc2 As Double, zc3 As Double, zc4 As Double
Dim zc5 As Double, zc6 As Double, zc7 As Double
Dim sx1 As Double, sx2 As Double, sx3 As Double
Dim sx4 As Double, sx5 As Double
Dim sy1 As Double, sy2 As Double, sy3 As Double
Dim sy4 As Double, sy5 As Double
Dim xw1 As Double, xw2 As Double, xw3 As Double
Dim yw1 As Double, yw2 As Double, yw3 As Double
Dim zw1 As Double, zw2 As Double, zw3 As Double
Dim sx3D As Double, sy3D As Double
Dim cursorx As Double, cursory As Double, cursorz As Double
Dim xa As Double, ya As Double, za As Double
Dim focal_length As Double
Dim ObjYaw As Double, ObjRoll As Double, ObjPitch As Double
Dim sOYaw As Double, cOYaw As Double
Dim sORoll As Double, cORoll As Double
Dim sOPitch As Double, cOPitch As Double
Dim xObj As Double, yObj As Double, zObj As Double
Dim CamYaw As Double, CamRoll As Double, CamPitch As Double
Dim sCYaw As Double, sCRoll As Double, sCPitch As Double
Dim cCYaw As Double, cCRoll As Double, cCPitch As Double
Dim xCam As Double, yCam As Double, zCam As Double
Dim rx As Double, ry As Double
Dim hcenter As Double, vcenter As Double
Dim viewheight As Double
Dim dist As Double, yawdist As Double
Dim pitchheading As Integer, yawheading As Integer
Dim viewchg As Integer
Dim yawdelta As Double, pitchdelta As Double
Dim planex As Double, planey As Double, planez As Double
Dim cursorxchg As Double, cursorychg As Double
Dim cursorzchg As Double
Dim signmx As Double, signmy As Double, signmz As Double
Dim cubeObj(7, 2) As Double               'cube xyz object coords
Dim cubeWorld(7, 2) As Double           'xw1,yw1,zw1 world coords
Dim camcoords(7, 2) As Double          'xc1,yc1,zc1 camera coords
Dim displaycoords(7, 2) As Double         'sx1,sy1 display coords

'----------------------------------------------------------------
'    Declaration of variables: rendering & backplane removal
'----------------------------------------------------------------
Dim visible As Double
Dim sp1 As Double, sp2 As Double, sp3 As Double
Dim xLight As Double, yLight As Double, zLight As Double
Dim illum_range As Double
Dim normalized_illum As Double
Dim xu As Double, yu As Double, zu As Double
Dim xv As Double, yv As Double, zv As Double
Dim x_surf_normal As Double
Dim y_surf_normal As Double
Dim z_surf_normal As Double
Dim v1 As Double, v2 As Double, v3 As Double
Dim x_unit_vector As Double
Dim y_unit_vector As Double
Dim z_unit_vector As Double
Dim zDeviceIllum As Integer
Dim zShadingColor As Integer

'----------------------------------------------------------------
'             Declaration of pixel-based variables
'----------------------------------------------------------------
Dim bInitialized As Integer
Dim clipx1 As Integer, clipy1 As Integer
Dim clipx2 As Integer, clipy2 As Integer
Dim ViewportWidth As Integer
Dim VirtualWidth As Integer
Dim ViewportDepth As Integer
Dim ViewportTemp As Integer
Dim Points(3, 1) As Integer

'
'----------------------------------------------------------------
'                     3 D    R O U T I N E S
'----------------------------------------------------------------
'

'----------------------------------------------------------------
'          Initialize and configure the 3D environment
'----------------------------------------------------------------
Sub zInitialize3D ()
If bInitialized = True Then
  Exit Sub
End If
clipx1 = 0: clipy1 = 0                     'coords of viewport...
clipx2 = zFRAMEWIDE - 1: clipy2 = zFRAMEHIGH - 1
'--------------------------- viewport ---------------------------
ViewportWidth = clipx2 - clipx1                'width of viewport
ViewportDepth = clipy2 - clipy1                'depth of viewport
ViewportTemp = ViewportDepth * 1.3333                  '4:3 ratio
VirtualWidth = ViewportTemp            'width of virtual viewport
DomainWidth = 800: DomainDepth = 600           'size of 3D domain
rx = VirtualWidth / DomainWidth                    'mapping ratio
ry = ViewportDepth / DomainDepth                   'mapping ratio
hcenter = ViewportWidth / 2                   'center of viewport
vcenter = ViewportDepth / 2                   'center of viewport
'------------------------ illumination --------------------------
xLight = -.21131                                 'light source...
yLight = .86603
zLight = .45315
illum_range = 255                    'surface brightness 0 to 255
zShadingColor = 4                          'current shading color
'--------------------------- camera -----------------------------
focal_length = 1200#                  'angular perspective factor
viewheight = 0             '0 for camera, else 5 or 6 for android
dist = 360: yawdist = 360             'camera-to-subject distance
CamYaw = 0#                                      'camera angle...
CamRoll = 0#
CamPitch = 6.28319
pitchheading = 360: yawheading = 0              'camera direction
viewchg = 2
yawdelta = 0: pitchdelta = 0
zSetCamAngle             'call procedure to initialize the camera
xCam = 0#: yCam = 0#: zCam = -360#               'camera location
signmx = 1: signmy = -1: signmz = -1
'--------------------------- models -----------------------------
ObjYaw = 0#                                    'model orientation
ObjRoll = 0#
ObjPitch = 0#
zSetObjAngle         'call procedure to initialize model rotation
xObj = 0#: yObj = 0#: zObj = 0#                   'model location

'------------ set default coords for parallelepiped -------------
cubeObj(0, 0) = 10: cubeObj(0, 1) = -10: cubeObj(0, 2) = 10
cubeObj(1, 0) = 10: cubeObj(1, 1) = 10: cubeObj(1, 2) = 10
cubeObj(2, 0) = -10: cubeObj(2, 1) = 10: cubeObj(2, 2) = 10
cubeObj(3, 0) = -10: cubeObj(3, 1) = -10: cubeObj(3, 2) = 10
cubeObj(4, 0) = 10: cubeObj(4, 1) = 10: cubeObj(4, 2) = -10
cubeObj(5, 0) = -10: cubeObj(5, 1) = 10: cubeObj(5, 2) = -10
cubeObj(6, 0) = -10: cubeObj(6, 1) = -10: cubeObj(6, 2) = -10
cubeObj(7, 0) = 10: cubeObj(7, 1) = -10: cubeObj(7, 2) = -10
'----------------------- virtual 3D cursors ---------------------
cursorx = 15: cursory = 15: cursorz = 15    '3D volumetric cursor
planex = 100: planey = 0: planez = 75             '3D groundplane
'--------------------- set a status indicator -------------------
bInitialized = True
End Sub

'----------------------------------------------------------------
'                      Set the camera heading.
'----------------------------------------------------------------
Sub zSetCameraHeading (ByVal Heading As Integer)
              'call with Heading in degrees ranging from 0 to 360
If bInitialized = False Then
  zShowMessage
  Exit Sub
End If
If Heading > zMAX_HEADING Then
  Exit Sub
End If
If Heading < zMIN_HEADING Then
  Exit Sub
End If
yawheading = Heading
CamYaw = yawheading * .0175433
If yawheading = 360 Then
  CamYaw = 6.28319
End If
If yawheading = 0 Then
  CamYaw = 0#
End If
zSetCamAngle
If CamYaw >= 4.71239 Then
  If CamYaw <= 6.28319 Then
    signmx = -1
    signmz = -1
    yawdelta = 6.28319 - CamYaw
    GoTo calccamyaw1
  End If
End If
If CamYaw >= 0 Then
  If CamYaw < 1.57079 Then
    signmx = 1
    signmz = -1
    yawdelta = CamYaw
    GoTo calccamyaw1
  End If
End If
If CamYaw >= 1.57079 Then
  If CamYaw < 3.14159 Then
    signmx = 1
    signmz = 1
    yawdelta = 3.14159 - CamYaw
    GoTo calccamyaw1
  End If
End If
If CamYaw >= 3.14159 Then
  If CamYaw < 4.71239 Then
    signmx = -1
    signmz = 1
    yawdelta = CamYaw - 3.14159
    GoTo calccamyaw1
  End If
End If

calccamyaw1:
xCam = Sin(yawdelta) * yawdist * signmx
zCam = Cos(yawdelta) * yawdist * signmz
End Sub

'----------------------------------------------------------------
'                       Set the camera pitch.
'----------------------------------------------------------------
Sub zSetCameraPitch (ByVal Pitch As Integer)
                'call with Pitch in range 270 to 360 (horizontal)
If bInitialized = False Then
  zShowMessage
  Exit Sub
End If
If Pitch > zMAX_PITCH Then
  Exit Sub                          'do not penetrate groundplane
End If
If Pitch < zMIN_PITCH Then
  Exit Sub                                'do not exceed vertical
End If
pitchheading = Pitch
CamPitch = pitchheading * .0174533            'convert to radians
If pitchheading = 360 Then
  CamPitch = 6.28319
End If
If pitchheading = 0 Then
  CamPitch = 0#
End If
zSetCamAngle
pitchdelta = 6.28319 - CamPitch    'change in pitch from start-up
yCam = Sin(pitchdelta) * dist * signmy         'new y translation
yawdist = Sqr((dist * dist) - (yCam * yCam))          'hypotenuse
xCam = Sin(yawdelta) * yawdist * signmx        'new x translation
zCam = Sqr((yawdist * yawdist) - (xCam * xCam)) * signmz   'new z
End Sub

'----------------------------------------------------------------
'                Set the camera-to-subject distance.
'----------------------------------------------------------------
Sub zSetCameraDistance (ByVal Range As Integer)
If bInitialized = False Then
  zShowMessage
  Exit Sub
End If
If Range < zMIN_DISTANCE Then
  Range = zMIN_DISTANCE
End If
If Range > zMAX_DISTANCE Then
  Range = zMAX_DISTANCE
End If
dist = Range
yCam = Sin(pitchdelta) * dist * signmy         'new y translation
yawdist = Sqr((dist * dist) - (yCam * yCam))          'hypotenuse
xCam = Sin(yawdelta) * yawdist * signmx        'new x translation
zCam = Sqr((yawdist * yawdist) - (xCam * xCam)) * signmz   'new z
End Sub

'----------------------------------------------------------------
'                 Set the translation of the subject.
'----------------------------------------------------------------
Sub zSetSubjectLocation (ByVal SSLx As Integer, ByVal SSLy As Integer, ByVal SSLz As Integer)
If bInitialized = False Then
  zShowMessage
  Exit Sub
End If
If SSLx > zMAX_LOCATION Then
  Exit Sub
End If
If SSLx < -zMAX_LOCATION Then
  Exit Sub
End If
If SSLy > zMAX_LOCATION Then
  Exit Sub
End If
If SSLy < -zMAX_LOCATION Then
  Exit Sub
End If
If SSLz > zMAX_LOCATION Then
  Exit Sub
End If
If SSLz < -zMAX_LOCATION Then
  Exit Sub
End If
xObj = SSLx
yObj = SSLy
zObj = SSLz
End Sub

'----------------------------------------------------------------
'                  Set the attitude of the subject.
'----------------------------------------------------------------
Sub zSetSubjectAttitude (ByVal Yaw As Integer, ByVal Roll As Integer, ByVal Pitch As Integer)
If bInitialized = False Then
  zShowMessage
  Exit Sub
End If
If Yaw < 0 Then
  Exit Sub
End If
If Yaw > 360 Then
  Exit Sub
End If
If Roll < 0 Then
  Exit Sub
End If
If Roll > 360 Then
  Exit Sub
End If
If Pitch < 0 Then
  Exit Sub
End If
If Pitch > 360 Then
  Exit Sub
End If
ObjYaw = Yaw * .0175433                    'convert to radians...
ObjRoll = Roll * .0175433
ObjPitch = Pitch * .0175433
If Yaw = 360 Then
  ObjYaw = 6.28319                    'tidy up boundary values...
End If
If Yaw = 0 Then
  ObjYaw = 0#
End If
If Roll = 360 Then
  ObjRoll = 6.28319
End If
If Roll = 0 Then
  ObjRoll = 0#
End If
If Pitch = 360 Then
  ObjPitch = 6.28319
End If
If Pitch = 0 Then
  ObjPitch = 0#
End If
zSetObjAngle                         'set Sine and Cosine factors
End Sub

'----------------------------------------------------------------
'       Display debugging message If system not initialized.
'----------------------------------------------------------------
Sub zShowMessage ()
MsgBox "3D routines not yet initialized by a call to zInitialize3D().", MB_OK, "Book 4224 3D programming error"
End Sub

'----------------------------------------------------------------
'              Perform the backplane visibility test
'----------------------------------------------------------------
Sub zVisibilityTest ()
  'Enter with 3 vertices as camera coords.
  'Exit with visibility token.
sp1 = xc1 * (yc2 * zc3 - yc3 * zc2)
sp1 = (-1) * sp1
sp2 = xc2 * (yc3 * zc1 - yc1 * zc3)
sp3 = xc3 * (yc1 * zc2 - yc2 * zc1)
visible = sp1 - sp2 - sp3
End Sub

'----------------------------------------------------------------
'       Calculate object Sine and Cosine rotation factors
'----------------------------------------------------------------
Sub zSetObjAngle ()
  'Enter with ObjYaw,ObjRoll,ObjPitch object rotation angles.
  'Exit with Sine, Cosine object rotation factors.
sOYaw = Sin(ObjYaw)
cOYaw = Cos(ObjYaw)
sORoll = Sin(ObjRoll)
cORoll = Cos(ObjRoll)
sOPitch = Sin(ObjPitch)
cOPitch = Cos(ObjPitch)
End Sub

'----------------------------------------------------------------
'       Calculate camera Sine and Cosine rotation factors
'----------------------------------------------------------------
Sub zSetCamAngle ()
  'Enter with Yaw,Roll,Pitch world rotation angles.
  'Exit with Sine, Cosine world rotation factors.
sCYaw = Sin(CamYaw)
sCRoll = Sin(CamRoll)
sCPitch = Sin(CamPitch)
cCYaw = Cos(CamYaw)
cCRoll = Cos(CamRoll)
cCPitch = Cos(CamPitch)
End Sub

'----------------------------------------------------------------
'             Transform object coords to screen coords
'----------------------------------------------------------------
Sub zPutObjToScreen ()
  'Enter with xyz object coordinates.  This routine transforms
  'the obj coords to world coords to image plane coords to
  'sx3D,sy3D physical screen coords.
zGetWorldCoords
zGetCameraCoords
zGetImageCoords
zGetScreenCoords
End Sub

'----------------------------------------------------------------
'             Transform world coords to screen coords
'----------------------------------------------------------------
Sub zPutWorldToScreen ()
  'Enter with xyz world coordinates.  This routine transforms
  'the world coords to image plane coords to sx3D,sy3D physical
  'screen coords.
zGetCameraCoords
zGetImageCoords
zGetScreenCoords
End Sub

'----------------------------------------------------------------
'           Calculate world coords from object coords
'----------------------------------------------------------------
Sub zGetWorldCoords ()
  'Enter with xyz unclipped object coordinates.
  'Exit with unclipped xyz world coordinates.
xa = cORoll * x + sORoll * y                      'roll rotate...
ya = cORoll * y - sORoll * x
x = cOYaw * xa - sOYaw * z                         'yaw rotate...
za = sOYaw * xa + cOYaw * z
z = cOPitch * za - sOPitch * ya                  'pitch rotate...
y = sOPitch * za + cOPitch * ya
x = x + xObj                                 'lateral movement...
y = y + yObj
z = z + zObj
End Sub

'----------------------------------------------------------------
'             Calculate camera coords from world coords
'----------------------------------------------------------------
Sub zGetCameraCoords ()
  'Enter with unclipped xyz world coordinates.
  'Exit with unclipped xyz camera coordinates.
x = (-1) * x            'adjust for cartesian coords of 2D screen
y = y - viewheight       'adjust world coords to height of viewer
x = x - xCam                                 'lateral movement...
y = y + yCam
z = z + zCam
xa = cCYaw * x - sCYaw * z                         'yaw rotate...
za = sCYaw * x + cCYaw * z
z = cCPitch * za - sCPitch * y                   'pitch rotate...
ya = sCPitch * za + cCPitch * y
x = cCRoll * xa + sCRoll * ya                     'roll rotate...
y = cCRoll * ya - sCRoll * xa
End Sub

'----------------------------------------------------------------
'           Calculate display coords from camera coords
'----------------------------------------------------------------
Sub zGetImageCoords ()
  'Enter with clipped xyz camera coordinates.
  'Exit with unclipped sx3D,sy3D display coordinates.
sx3D = focal_length * (x / z)
sy3D = focal_length * (y / z)
End Sub

'----------------------------------------------------------------
'           Calculate screen coords from display coords
'----------------------------------------------------------------
Sub zGetScreenCoords ()
  'Enter with unclipped sx3D,sy3D display coordinates.
  'Exit with sx3D,sy3D device-dependent display coordinates
  'scaled to the world range with correct aspect ratio.
sx3D = sx3D * rx
sy3D = sy3D * ry
sx3D = sx3D + hcenter
sy3D = sy3D + vcenter
End Sub

'----------------------------------------------------------------
'                   Draw a four-sided polygon
'----------------------------------------------------------------
Sub zDrawFacet ()
  Dim hPrevBrush As Integer, hFacetBrush As Integer
  Dim hPrevRegion As Integer, hFacetRegion As Integer
Points(0, 0) = sx1      'load polygon region array with coords...
Points(1, 0) = sy1
Points(2, 0) = sx2
Points(3, 0) = sy2
Points(0, 1) = sx3
Points(1, 1) = sy3
Points(2, 1) = sx4
Points(3, 1) = sy4
zGetBrightness                    'get brightness factor of facet
Select Case zShadingColor           'set the dithering pattern...
  Case zRED
    hFacetBrush = CreateSolidBrush(RGB(zDeviceIllum, 0, 0))
  Case zGREEN
    hFacetBrush = CreateSolidBrush(RGB(0, zDeviceIllum, 0))
  Case zBROWN
    hFacetBrush = CreateSolidBrush(RGB(zDeviceIllum, zDeviceIllum, 0))
  Case zBLUE
    hFacetBrush = CreateSolidBrush(RGB(0, 0, zDeviceIllum))
  Case zMAGENTA
    hFacetBrush = CreateSolidBrush(RGB(zDeviceIllum, 0, zDeviceIllum))
  Case zCYAN
    hFacetBrush = CreateSolidBrush(RGB(0, zDeviceIllum, zDeviceIllum))
  Case zGRAY
    hFacetBrush = CreateSolidBrush(RGB(zDeviceIllum, zDeviceIllum, zDeviceIllum))
  Case Else
    hFacetBrush = CreateSolidBrush(RGB(0, 0, zDeviceIllum))
End Select
hPrevBrush = SelectObject(Form1.hDC, hFacetBrush)   'select brush
hFacetRegion = CreatePolygonRgn(Points(0, 0), 4, WINDING) 'region
'      PROGRAMMER'S NOTE:  An idiosyncracy in VBRUN100.DLL
'      causes the created polygon region to be automatically
'      selected into the DC.  Attempting to use the API to
'      select and deselect the polygon region will corrupt
'      the API's internal table of handles to regions.
RetVal = PaintRgn(Form1.hDC, hFacetRegion)       'fill the region
RetVal = DeleteObject(hFacetRegion)               'destroy region
RetVal = SelectObject(Form1.hDC, hPrevBrush)      'deselect brush
RetVal = DeleteObject(hFacetBrush)                 'destroy brush
End Sub

'----------------------------------------------------------------
'            Calculate the brightness level of a facet
'----------------------------------------------------------------
Sub zGetBrightness ()
  'Enter with facet world coordinates.
  'Exit with illumination level token.
xu = xw2 - xw1                  'vector from vertex 1 to vertex 2
yu = yw2 - yw1
zu = zw2 - zw1
xv = xw3 - xw1                  'vector from vertex 1 to vertex 3
yv = yw3 - yw1
zv = zw3 - zw1
x_surf_normal = (yu * zv) - (zu * yv)
y_surf_normal = (zu * xv) - (xu * zv)
z_surf_normal = (xu * yv) - (yu * xv)
y_surf_normal = y_surf_normal * (-1)
z_surf_normal = z_surf_normal * (-1)        'use cartesian system
v1 = (x_surf_normal * x_surf_normal) + (y_surf_normal * y_surf_normal) + (z_surf_normal * z_surf_normal)
v2 = Sqr(v1)           'magnitude of surface perpendicular vector
v3 = 1 / v2          'ratio of magnitude to length of unit vector
x_unit_vector = v3 * x_surf_normal
y_unit_vector = v3 * y_surf_normal
z_unit_vector = v3 * z_surf_normal      'surface perp unit vector
normalized_illum = (x_unit_vector * xLight) + (y_unit_vector * yLight) + (z_unit_vector * zLight)
normalized_illum = normalized_illum * illum_range       '0 to 255
zDeviceIllum = normalized_illum                'change to integer
If zDeviceIllum < 47 Then
  zDeviceIllum = 47                       'inhibit deepest shadow
End If
End Sub

'----------------------------------------------------------------
'             Set the current shading color for facets
'----------------------------------------------------------------
Sub zSetShadingColor (ByVal iHue As Integer)
If bInitialized = False Then
  zShowMessage
  Exit Sub
End If
If iHue < 1 Then
  Exit Sub
End If
If iHue > 7 Then
  Exit Sub
End If
zShadingColor = iHue
End Sub

'----------------------------------------------------------------
'              Draw an instance of a parallelepiped
'----------------------------------------------------------------
Sub zGetCubeCoords ()                      'called by zDrawCube()
  Dim Count As Integer
  Dim negx As Double, negy As Double, negz As Double
negx = (-1) * (cursorx)
negy = (-1) * (cursory)
negz = (-1) * (cursorz)
cubeObj(0, 0) = cursorx
cubeObj(0, 1) = negy
cubeObj(0, 2) = cursorz
cubeObj(1, 0) = cursorx
cubeObj(1, 1) = cursory
cubeObj(1, 2) = cursorz
cubeObj(2, 0) = negx
cubeObj(2, 1) = cursory
cubeObj(2, 2) = cursorz
cubeObj(3, 0) = negx
cubeObj(3, 1) = negy
cubeObj(3, 2) = cursorz
cubeObj(4, 0) = cursorx
cubeObj(4, 1) = cursory
cubeObj(4, 2) = negz
cubeObj(5, 0) = negx
cubeObj(5, 1) = cursory
cubeObj(5, 2) = negz
cubeObj(6, 0) = negx
cubeObj(6, 1) = negy
cubeObj(6, 2) = negz
cubeObj(7, 0) = cursorx
cubeObj(7, 1) = negy
cubeObj(7, 2) = negz
For Count = 0 To 7
  x = cubeObj(Count, 0)
  y = cubeObj(Count, 1)
  z = cubeObj(Count, 2)
  zGetWorldCoords
  cubeWorld(Count, 0) = x
  cubeWorld(Count, 1) = y
  cubeWorld(Count, 2) = z
  zGetCameraCoords
  camcoords(Count, 0) = x
  camcoords(Count, 1) = y
  camcoords(Count, 2) = z
  zGetImageCoords
  zGetScreenCoords
  displaycoords(Count, 0) = sx3D
  displaycoords(Count, 1) = sy3D
Next Count
End Sub

'----------------------------------------------------------------
'      Draw a parallelepiped at current position in 3D domain
'----------------------------------------------------------------
'
Sub zDrawCube ()                 'draw box at current 3D position
If bInitialized = False Then
  zShowMessage
  Exit Sub
End If
zGetCubeCoords              'get camera coords and display coords
surface0:
  xc1 = camcoords(7, 0)
  yc1 = camcoords(7, 1)
  zc1 = camcoords(7, 2)
  xc2 = camcoords(0, 0)
  yc2 = camcoords(0, 1)
  zc2 = camcoords(0, 2)
  xc3 = camcoords(3, 0)
  yc3 = camcoords(3, 1)
  zc3 = camcoords(3, 2)
  xc4 = camcoords(6, 0)
  yc4 = camcoords(6, 1)
  zc4 = camcoords(6, 2)
  zVisibilityTest
  If visible > 0 Then
    GoTo surface1
  End If
  sx1 = displaycoords(7, 0)
  sy1 = displaycoords(7, 1)
  sx2 = displaycoords(0, 0)
  sy2 = displaycoords(0, 1)
  sx3 = displaycoords(3, 0)
  sy3 = displaycoords(3, 1)
  sx4 = displaycoords(6, 0)
  sy4 = displaycoords(6, 1)
  xw3 = cubeWorld(7, 0)
  yw3 = cubeWorld(7, 1)
  zw3 = cubeWorld(7, 2)
  xw2 = cubeWorld(0, 0)
  yw2 = cubeWorld(0, 1)
  zw2 = cubeWorld(0, 2)
  xw1 = cubeWorld(3, 0)
  yw1 = cubeWorld(3, 1)
  zw1 = cubeWorld(3, 2)
  zDrawFacet
surface1:
  xc1 = camcoords(6, 0)
  yc1 = camcoords(6, 1)
  zc1 = camcoords(6, 2)
  xc2 = camcoords(5, 0)
  yc2 = camcoords(5, 1)
  zc2 = camcoords(5, 2)
  xc3 = camcoords(4, 0)
  yc3 = camcoords(4, 1)
  zc3 = camcoords(4, 2)
  xc4 = camcoords(7, 0)
  yc4 = camcoords(7, 1)
  zc4 = camcoords(7, 2)
  zVisibilityTest
  If visible > 0 Then
    GoTo surface2
  End If
  sx1 = displaycoords(6, 0)
  sy1 = displaycoords(6, 1)
  sx2 = displaycoords(5, 0)
  sy2 = displaycoords(5, 1)
  sx3 = displaycoords(4, 0)
  sy3 = displaycoords(4, 1)
  sx4 = displaycoords(7, 0)
  sy4 = displaycoords(7, 1)
  xw3 = cubeWorld(6, 0)
  yw3 = cubeWorld(6, 1)
  zw3 = cubeWorld(6, 2)
  xw2 = cubeWorld(5, 0)
  yw2 = cubeWorld(5, 1)
  zw2 = cubeWorld(5, 2)
  xw1 = cubeWorld(4, 0)
  yw1 = cubeWorld(4, 1)
  zw1 = cubeWorld(4, 2)
  zDrawFacet
surface2:
  xc1 = camcoords(3, 0)
  yc1 = camcoords(3, 1)
  zc1 = camcoords(3, 2)
  xc2 = camcoords(2, 0)
  yc2 = camcoords(2, 1)
  zc2 = camcoords(2, 2)
  xc3 = camcoords(5, 0)
  yc3 = camcoords(5, 1)
  zc3 = camcoords(5, 2)
  xc4 = camcoords(6, 0)
  yc4 = camcoords(6, 1)
  zc4 = camcoords(6, 2)
  zVisibilityTest
  If visible > 0 Then
    GoTo surface3
  End If
  sx1 = displaycoords(3, 0)
  sy1 = displaycoords(3, 1)
  sx2 = displaycoords(2, 0)
  sy2 = displaycoords(2, 1)
  sx3 = displaycoords(5, 0)
  sy3 = displaycoords(5, 1)
  sx4 = displaycoords(6, 0)
  sy4 = displaycoords(6, 1)
  xw3 = cubeWorld(3, 0)
  yw3 = cubeWorld(3, 1)
  zw3 = cubeWorld(3, 2)
  xw2 = cubeWorld(2, 0)
  yw2 = cubeWorld(2, 1)
  zw2 = cubeWorld(2, 2)
  xw1 = cubeWorld(5, 0)
  yw1 = cubeWorld(5, 1)
  zw1 = cubeWorld(5, 2)
  zDrawFacet
surface3:
  xc1 = camcoords(0, 0)
  yc1 = camcoords(0, 1)
  zc1 = camcoords(0, 2)
  xc2 = camcoords(1, 0)
  yc2 = camcoords(1, 1)
  zc2 = camcoords(1, 2)
  xc3 = camcoords(2, 0)
  yc3 = camcoords(2, 1)
  zc3 = camcoords(2, 2)
  xc4 = camcoords(3, 0)
  yc4 = camcoords(3, 1)
  zc4 = camcoords(3, 2)
  zVisibilityTest
  If visible > 0 Then
    GoTo surface4
  End If
  sx1 = displaycoords(0, 0)
  sy1 = displaycoords(0, 1)
  sx2 = displaycoords(1, 0)
  sy2 = displaycoords(1, 1)
  sx3 = displaycoords(2, 0)
  sy3 = displaycoords(2, 1)
  sx4 = displaycoords(3, 0)
  sy4 = displaycoords(3, 1)
  xw3 = cubeWorld(0, 0)
  yw3 = cubeWorld(0, 1)
  zw3 = cubeWorld(0, 2)
  xw2 = cubeWorld(1, 0)
  yw2 = cubeWorld(1, 1)
  zw2 = cubeWorld(1, 2)
  xw1 = cubeWorld(2, 0)
  yw1 = cubeWorld(2, 1)
  zw1 = cubeWorld(2, 2)
  zDrawFacet
surface4:
  xc1 = camcoords(7, 0)
  yc1 = camcoords(7, 1)
  zc1 = camcoords(7, 2)
  xc2 = camcoords(4, 0)
  yc2 = camcoords(4, 1)
  zc2 = camcoords(4, 2)
  xc3 = camcoords(1, 0)
  yc3 = camcoords(1, 1)
  zc3 = camcoords(1, 2)
  xc4 = camcoords(0, 0)
  yc4 = camcoords(0, 1)
  zc4 = camcoords(0, 2)
  zVisibilityTest
  If visible > 0 Then
    GoTo surface5
  End If
  sx1 = displaycoords(7, 0)
  sy1 = displaycoords(7, 1)
  sx2 = displaycoords(4, 0)
  sy2 = displaycoords(4, 1)
  sx3 = displaycoords(1, 0)
  sy3 = displaycoords(1, 1)
  sx4 = displaycoords(0, 0)
  sy4 = displaycoords(0, 1)
  xw3 = cubeWorld(7, 0)
  yw3 = cubeWorld(7, 1)
  zw3 = cubeWorld(7, 2)
  xw2 = cubeWorld(4, 0)
  yw2 = cubeWorld(4, 1)
  zw2 = cubeWorld(4, 2)
  xw1 = cubeWorld(1, 0)
  yw1 = cubeWorld(1, 1)
  zw1 = cubeWorld(1, 2)
  zDrawFacet
surface5:
  xc1 = camcoords(1, 0)
  yc1 = camcoords(1, 1)
  zc1 = camcoords(1, 2)
  xc2 = camcoords(4, 0)
  yc2 = camcoords(4, 1)
  zc2 = camcoords(4, 2)
  xc3 = camcoords(5, 0)
  yc3 = camcoords(5, 1)
  zc3 = camcoords(5, 2)
  xc4 = camcoords(2, 0)
  yc4 = camcoords(2, 1)
  zc4 = camcoords(2, 2)
  zVisibilityTest
  If visible > 0 Then
    GoTo surfaces_done
  End If
  sx1 = displaycoords(1, 0)
  sy1 = displaycoords(1, 1)
  sx2 = displaycoords(4, 0)
  sy2 = displaycoords(4, 1)
  sx3 = displaycoords(5, 0)
  sy3 = displaycoords(5, 1)
  sx4 = displaycoords(2, 0)
  sy4 = displaycoords(2, 1)
  xw3 = cubeWorld(1, 0)
  yw3 = cubeWorld(1, 1)
  zw3 = cubeWorld(1, 2)
  xw2 = cubeWorld(4, 0)
  yw2 = cubeWorld(4, 1)
  zw2 = cubeWorld(4, 2)
  xw1 = cubeWorld(5, 0)
  yw1 = cubeWorld(5, 1)
  zw1 = cubeWorld(5, 2)
  zDrawFacet
surfaces_done:
End Sub

'----------------------------------------------------------------
'                 Set the extrusion of the subject.
'----------------------------------------------------------------
Sub zSetSubjectSize (ByVal iWidth As Integer, ByVal iHeight As Integer, ByVal iDepth As Integer)
If bInitialized = False Then
zShowMessage
  Exit Sub
End If
If iWidth < 2 Then
  iWidth = 2                               'inhibit the range...
End If
If iWidth > 100 Then
  iWidth = 100
End If
If iHeight < 2 Then
  iHeight = 2
End If
If iHeight > 100 Then
  iHeight = 100
End If
If iDepth < 2 Then
  iDepth = 2
End If
If iDepth > 100 Then
  iDepth = 100
End If
cursorx = iWidth                    'set the extrusion factors...
cursory = iHeight
cursorz = iDepth
End Sub

