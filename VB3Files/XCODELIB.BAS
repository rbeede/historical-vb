'TrackPopupMenu declares
Declare Function TrackPopupMenu% Lib "user" (ByVal hMenu%, ByVal wFlags%, ByVal X%, ByVal Y%, ByVal r2%, ByVal hWnd%, ByVal r1&)
Declare Function GetMenu% Lib "user" (ByVal hWnd%)
Declare Function GetSubMenu% Lib "user" (ByVal hMenu%, ByVal nPos%)
'Declares for TestLength
Global Const MB_ICONEXCLAMATION = 48


'Declares for SearchWindowLIst
Declare Function GetWindow% Lib "USER" (ByVal hWnd%, ByVal wCmd%)
Global Const GW_HWNDFIRST = 0
Global Const GW_HWNDNEXT = 2
Declare Function GetWindowText Lib "User" (ByVal hWnd As Integer, ByVal lpString As String, ByVal aint As Integer) As Integer


'Declares for RestoreApp
Declare Function IsIconic Lib "user" (ByVal hWnd As Any)


'Declares for Loaded
Declare Function FindWindow Lib "user" (ByVal CName As Any, ByVal Caption As Any)
'Declares for IconExtractor
Const GWW_HINSTANCE = (-6)
Declare Function GetWindowWord Lib "User" (ByVal hWnd As Integer, ByVal nIndex As Integer) As Integer
Declare Function ExtractIcon Lib "shell" (ByVal lpHandle As Integer, ByVal lpExe As String, ByVal lpiconindex As Integer) As Integer
Declare Function DrawIcon Lib "USER" (ByVal lpHandle As Integer, ByVal xcoord As Integer, ByVal ycoord As Integer, ByVal Hicon As Integer) As Integer

'Declares for GetSystemDir
Declare Function GetSystemDirectory Lib "Kernel" (ByVal lpBuffer As String, ByVal nSize As Integer) As Integer
Option Explicit
'Declares for Find Parent
Declare Function GetParent Lib "User" (ByVal hWnd As Integer) As Integer
Declare Sub SetWindowPos Lib "User" (ByVal hWnd As Integer, ByVal hWndInsertAfter As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal cx As Integer, ByVal cy As Integer, ByVal wFlags As Integer)
'Declares for FrameControl
Global Const HiColor = &HFFFFFF
Global Const LoColor = &H808080
Const NIL = ""



Sub AlertBox (Mess$)
' 12/31/94: * creates an Alert box with an OK button
MsgBox Mess$, 48, App.Title
End Sub

Sub CenterForm (F As Form)
' 5/24/95
F.Move (SCREEN.Width - F.Width) / 2, (SCREEN.Height - F.Height) / 2
End Sub

Function CNull (v As Variant) As String
' 4/27/94: if v is NULL then return "" else return its string value
If IsNull(v) Then CNull = "" Else CNull = CStr(v)
End Function

Function CompactDB (db As Database, sPath As String) As Integer
' 12/31/94  compact a db
' in: db - handle to database named in sPath
' in: sPath - fully qualified path to db
' returns: TRUE if compacting was successful else FALSE
'
Const DB_LANG_GENERAL = ";LANGID=0x0809;CP=1252;COUNTRY=0"
Const HOURGLASS = 11
Const DEFAULT = 0
Dim sTmpDB As String
Dim dbTmp As Database
'
' create Temp filename. It ends with .TMP
sTmpDB = TmpFileName()
' now change suffix from .TMP to .MDB
Mid$(sTmpDB, InStr(1, sTmpDB, "TMP")) = "MDB"

SCREEN.MousePointer = HOURGLASS
' make sure on multi-user system that everyone is off
db.Close
' open in exclusive mode (single-user mode) so noone else can have access
Set dbTmp = OpenDatabase(sPath, True, True)
dbTmp.Close
CompactDatabase sPath, sTmpDB, DB_LANG_GENERAL
If Dir$(sTmpDB) <> NIL Then
  ' successfully created new db
  Kill sPath                ' erase old db
  Name sTmpDB As sPath      ' rename and move as new db
  InfoBox (sPath & " compacted")
  CompactDB = True
Else
  AlertBox ("Compacting " & sPath & "failed")
  CompactDB = False
End If
SCREEN.MousePointer = DEFAULT

End Function

Function confirm% (Ask$)
If MsgBox(Ask$, 52, App.Title) = 6 Then confirm% = True
End Function

Function ConfirmNO (Ask$, Title$)
' 308 = 4+48+256
If MsgBox(Ask$, 308, Title$) = 6 Then ConfirmNO = True
End Function

Function CreatePath (ByVal DestPath$) As Integer
'---------------------------------------------
' Create the path contained in DestPath$
' First char must be drive letter, followed by
' a ":\" followed by the path, if any.
'---------------------------------------------
Dim BackPos As Integer
Dim forePos As Integer
Dim sTemp As String

    SCREEN.MousePointer = 11

    '---------------------------------------------
    ' Add slash to end of path if not there already
    '---------------------------------------------
    If Right$(DestPath$, 1) <> "\" Then
        DestPath$ = DestPath$ + "\"
    End If
          

    '-----------------------------------
    ' Change to the root dir of the drive
    '-----------------------------------
    On Error Resume Next
    ChDrive DestPath$
    If Err <> 0 Then GoTo errorOut
    ChDir "\"

    '-------------------------------------------------
    ' Attempt to make each directory, then change to it
    '-------------------------------------------------
    BackPos = 3
    forePos = InStr(4, DestPath$, "\")
    Do While forePos <> 0
        sTemp = Mid$(DestPath$, BackPos + 1, forePos - BackPos - 1)

        Err = 0
        MkDir sTemp
        If Err <> 0 And Err <> 75 Then GoTo errorOut

        Err = 0
        ChDir sTemp
        If Err <> 0 Then GoTo errorOut

        BackPos = forePos
        forePos = InStr(BackPos + 1, DestPath$, "\")
    Loop
                 
    CreatePath = True
    SCREEN.MousePointer = 0
    Exit Function
                 
errorOut:
    MsgBox "Error While Attempting to Create Directories on Destination Drive.", 48, "SETUP"
    CreatePath = False
    SCREEN.MousePointer = 0

End Function

Sub CreateProgManGroup (X As Form, GroupName$, GroupPath$)
'-------------------------------------------------------------
' Procedure: CreateProgManGroup
' Arguments: X           The Form where a Label1 exist
'            GroupName$  A string that contains the group name
'            GroupPath$  A string that contains the group file
'                        name  ie 'myapp.grp'
'-------------------------------------------------------------
    
    SCREEN.MousePointer = 11
    
    '----------------------------------------------------------------------
    ' Windows requires DDE in order to create a program group and item.
    ' Here, a Visual Basic label control is used to generate the DDE messages
    '----------------------------------------------------------------------
    On Error Resume Next

    
    '--------------------------------
    ' Set LinkTopic to PROGRAM MANAGER
    '--------------------------------
    X.Label1.LinkTopic = "ProgMan|Progman"
    X.Label1.LinkMode = 2
    For i% = 1 To 10                                         ' Loop to ensure that there is enough time to
      z% = DoEvents()                                        ' process DDE Execute.  This is redundant but needed
    Next                                                     ' for debug windows.
    X.Label1.LinkTimeout = 100


    '---------------------
    ' Create program group
    '---------------------
    X.Label1.LinkExecute "[CreateGroup(" + GroupName$ + Chr$(44) + GroupPath$ + ")]"


    '-----------------
    ' Reset properties
    '-----------------
    X.Label1.LinkTimeout = 50
    X.Label1.LinkMode = 0
    
    SCREEN.MousePointer = 0
End Sub

Sub CreateProgManItem (X As Form, CmdLine$, IconTitle$)

'----------------------------------------------------------
' Procedure: CreateProgManItem
'
' Arguments: X           The form where Label1 exists
'
'            CmdLine$    A string that contains the command
'                        line for the item/icon.
'                        ie 'c:\myapp\setup.exe'
'
'            IconTitle$  A string that contains the item's
'                        caption
'----------------------------------------------------------
    
    SCREEN.MousePointer = 11
    
    '----------------------------------------------------------------------
    ' Windows requires DDE in order to create a program group and item.
    ' Here, a Visual Basic label control is used to generate the DDE messages
    '----------------------------------------------------------------------
    On Error Resume Next


    '---------------------------------
    ' Set LinkTopic to PROGRAM MANAGER
    '---------------------------------
    X.Label1.LinkTopic = "ProgMan|Progman"
    X.Label1.LinkMode = 2
    For i% = 1 To 10                                         ' Loop to ensure that there is enough time to
      z% = DoEvents()                                        ' process DDE Execute.  This is redundant but needed
    Next                                                     ' for debug windows.
    X.Label1.LinkTimeout = 100

    
    '------------------------------------------------
    ' Create Program Item, one of the icons to launch
    ' an application from Program Manager
    '------------------------------------------------
    X.Label1.LinkExecute "[AddItem(" + CmdLine$ + Chr$(44) + IconTitle$ + Chr$(44) + ",,)]"
    
    '-----------------
    ' Reset properties
    '-----------------
    X.Label1.LinkTimeout = 50
    X.Label1.LinkMode = 0
    
    SCREEN.MousePointer = 0
End Sub

Function Cvtchar (X As String, Y As String, z As String) As String
'  4/23/94: convert all occurrences of char y to char z in string x

Dim iLen As Integer     ' len of string x
Dim J As Integer
Dim sTmp As String     ' converted string
Dim ch As String * 1

iLen = Len(X)
sTmp = NIL
For J = 1 To iLen
  ch = Mid$(X, J, 1)
  If ch <> Y Then sTmp = sTmp & ch Else sTmp = sTmp & z
Next J
Cvt = sTmp
End Function

Function dquote (inpstr As String) As String
    ' Return the argument, with double quotes around it.
    dquote = Chr$(34) + inpstr + Chr$(34)
End Function

Function Exists% (F$)
On Error Resume Next
X& = FileLen(F$)
If X& Then Exists% = True
End Function

Function famt (sStr As String, sDelim As String) As Integer
' 2/3/95: count #of fields separated by sDelim.
Dim iAmt As Integer
Dim iLen As Integer
Dim J As Integer

iLen = Len(sStr)
If iLen Then
  ' we have at least 1 field
  iAmt = 1
  For J = 1 To iLen
    If Mid$(sStr, J, 1) = sDelim Then iAmt = iAmt + 1
  Next J
Else
  iAmt = 0
End If
famt = iAmt
End Function

Sub FindAndRestorePrevInstance (Cap$)
Dim X%
If App.PrevInstance Then
   AppActivate Cap$
   SendKeys ("% R")
   End
End If
End Sub

Function FindProgentor (WinHand As Integer) As Integer
     Parent% = GetParent(WinHand%)
     OldParent% = Parent%
     'Get the parent of the parent if any
     Do While Parent%
          OldParent% = Parent%
          Parent% = GetParent%(OldParent%)
         ' Debug.Print Parent%
     Loop
Parent% = OldParent%
FindProgenitor = Parent%
End Function

Function FixAPIString$ (ByVal Test$)
FixAPIString$ = Trim(Left$(Test$, InStr(Test$, Chr$(0)) - 1))
End Function

Function FixPath$ (Test$)
'sticks a backslash on the end of test$ if there's
'not one there already
Dim T$
T$ = Test$
If Right$(T$, 1) <> "\" Then T$ = T$ + "\"
FixPath$ = T$
End Function

Sub FormStayOnTop (handle%, OnTop%)
Const Swp_Nosize = &H1
Const SWP_Nomove = &H2
Const Swp_NoActivate = &H10
Const Swp_ShowWindow = &H40
Const Hwnd_TopMost = -1
Const Hwnd_NoTopMost = -2
wFlags = SWP_Nomove Or Swp_Nosize Or Swp_ShowWindow Or Swp_NoActivate
Select Case OnTop%
     Case True
        PosFlag = Hwnd_TopMost
     Case False
         PosFlag = Hwnd_NoTopMost
     End Select
SetWindowPos handle%, PosFlag, 0, 0, 0, 0, wFlags
End Sub

Sub FrameControl (F As Form, C As Control, OffSet As Integer, iWidth As Integer)
'Declares for FormStayOnTop
F.DrawWidth = iWidth
F.forecolor = &HFFFFFF
'bottom:
F.Line (C.Left, C.Top + C.Height + OffSet)-(C.Left + C.Width, C.Top + C.Height + OffSet)
'right:
F.Line (C.Left + C.Width + OffSet, C.Top)-(C.Left + C.Width + OffSet, C.Top + C.Height + OffSet)
F.forecolor = &H808080
'top:
F.Line (C.Left - OffSet * 1.5, C.Top - OffSet * 1.5)-(C.Left + C.Width + OffSet * 1.5, C.Top - OffSet * 1.5)
'left:
F.Line (C.Left - OffSet * 1.5, C.Top - OffSet * 1.5)-(C.Left - OffSet * 1.5, C.Top + C.Height + OffSet)

End Sub

Sub GetSystemDir (SystemPath$)
Dim Sys As String * 256
X = GetSystemDirectory(Sys, Len(Sys))
X = InStr(1, Sys, Chr$(0))
SystemPath$ = Left$(Sys, InStr(Sys, Chr$(0)) - 1)
End Sub

Function HIWORD% (LongVal&)
HIWORD% = LongVal& \ 65536 ' (note: '\', not '/')
End Function

Sub IconExtractor (ExeFile$, F As Form, Pic As Picture)
handle = F.hWnd
z = SCREEN.Height
    Select Case z
        Case 7000
            X = 2: Y = 1
        Case 7200
            X = 3: Y = 0
        Case 9000
            X = 3: Y = 0
        Case Is > 9000
            X = 8: Y = 4
    End Select
                
    Static Looper
    Looper = Looper + 1
    Inst = GetWindowWord(handle, GWW_HINSTANCE)
    Hicon = ExtractIcon(Inst, ExeFile$, Looper - 1)
    If Hicon = 0 Then
        If Looper > 0 Then
            Hicon = ExtractIcon(Inst, ExeFile$, 0)
            Looper = 1
        Else Beep: Exit Sub
        End If
    End If
    F.Pic.Cls
    Draw = DrawIcon(F.Pic.hDC, X, Y, Hicon)
End Sub

Sub InfoBox (Mess$)
' 9/5/94  creates an Info box with an OK button
MsgBox Mess$, 64, App.Title
End Sub

Function IsAlpha (ch As String) As Integer
' 9/2/94: see if ch is an alphabetic char a-z or A-Z
Dim iCh As Integer

iCh = Asc(UCase$(ch))
If (65 <= iCh And iCh <= 90) Then
  IsAlpha = True
Else
  IsAlpha = False
End If
End Function

Function IsDigit (ch As String) As Integer
' 9/2/94: 48 = asc("0"), 57 = asc("9")
Dim iCh As Integer

iCh = Asc(ch)
If (48 <= iCh) And (iCh <= 57) Then
  IsDigit = True
Else
  IsDigit = False
End If
End Function

'------------------------------------------------------
' Function:   IsValidPath as integer
' arguments:  DestPath$         a string that is a full path
'             DefaultDrive$     the default drive.  eg.  "C:"
'
'  If DestPath$ does not include a drive specification,
'  IsValidPath uses Default Drive
'
'  When IsValidPath is finished, DestPath$ is reformated
'  to the format "X:\dir\dir\dir\"
'
' Result:  True (-1) if path is valid.
'          False (0) if path is invalid
'-------------------------------------------------------
Function IsValidPath (DestPath$, ByVal DefaultDrive$) As Integer

    '----------------------------
    ' Remove left and right spaces
    '----------------------------
    DestPath$ = RTrim$(LTrim$(DestPath$))
    

    '-----------------------------
    ' Check Default Drive Parameter
    '-----------------------------
    If Right$(DefaultDrive$, 1) <> ":" Or Len(DefaultDrive$) <> 2 Then
        MsgBox "Bad default drive parameter specified in IsValidPath Function.  You passed,  """ + DefaultDrive$ + """.  Must be one drive letter and "":"".  For example, ""C:"", ""D:""...", 64, "Setup Kit Error"
        GoTo parseErr
    End If
    

    '-------------------------------------------------------
    ' Insert default drive if path begins with root backslash
    '-------------------------------------------------------
    If Left$(DestPath$, 1) = "\" Then
        DestPath$ = DefaultDrive + DestPath$
    End If
    
    '-----------------------------
    ' check for invalid characters
    '-----------------------------
    On Error Resume Next
    tmp$ = Dir$(DestPath$)
    If Err <> 0 Then
        GoTo parseErr
    End If
    

    '-----------------------------------------
    ' Check for wildcard characters and spaces
    '-----------------------------------------
    If (InStr(DestPath$, "*") <> 0) GoTo parseErr
    If (InStr(DestPath$, "?") <> 0) GoTo parseErr
    If (InStr(DestPath$, " ") <> 0) GoTo parseErr
         
    
    '------------------------------------------
    ' Make Sure colon is in second char position
    '------------------------------------------
    If Mid$(DestPath$, 2, 1) <> Chr$(58) Then GoTo parseErr
    

    '-------------------------------
    ' Insert root backslash if needed
    '-------------------------------
    If Len(DestPath$) > 2 Then
      If Right$(Left$(DestPath$, 3), 1) <> "\" Then
        DestPath$ = Left$(DestPath$, 2) + "\" + Right$(DestPath$, Len(DestPath$) - 2)
      End If
    End If

    '-------------------------
    ' Check drive to install on
    '-------------------------
    drive$ = Left$(DestPath$, 1)
    ChDrive (drive$)                                                        ' Try to change to the dest drive
    If Err <> 0 Then GoTo parseErr
    
    '-----------
    ' Add final \
    '-----------
    If Right$(DestPath$, 1) <> "\" Then
        DestPath$ = DestPath$ + "\"
    End If
    

    '-------------------------------------
    ' Root dir is a valid dir
    '-------------------------------------
    If Len(DestPath$) = 3 Then
        If Right$(DestPath$, 2) = ":\" Then
            GoTo ParseOK
        End If
    End If
    

    '------------------------
    ' Check for repeated Slash
    '------------------------
    If InStr(DestPath$, "\\") <> 0 Then GoTo parseErr
        
    '--------------------------------------
    ' Check for illegal directory names
    '--------------------------------------
    legalChar$ = "!#$%&'()-0123456789@ABCDEFGHIJKLMNOPQRSTUVWXYZ^_`{}~."
    BackPos = 3
    forePos = InStr(4, DestPath$, "\")
    Do
        temp$ = Mid$(DestPath$, BackPos + 1, forePos - BackPos - 1)
        
        '----------------------------
        ' Test for illegal characters
        '----------------------------
        For i = 1 To Len(temp$)
            If InStr(legalChar$, UCase$(Mid$(temp$, i, 1))) = 0 Then GoTo parseErr
        Next i

        '-------------------------------------------
        ' Check combinations of periods and lengths
        '-------------------------------------------
        periodPos = InStr(temp$, ".")
        length = Len(temp$)
        If periodPos = 0 Then
            If length > 8 Then GoTo parseErr                         ' Base too long
        Else
            If periodPos > 9 Then GoTo parseErr                      ' Base too long
            If length > periodPos + 3 Then GoTo parseErr             ' Extension too long
            If InStr(periodPos + 1, temp$, ".") <> 0 Then GoTo parseErr' Two periods not allowed
        End If

        BackPos = forePos
        forePos = InStr(BackPos + 1, DestPath$, "\")
    Loop Until forePos = 0

ParseOK:
    IsValidPath = True
    Exit Function

parseErr:
    IsValidPath = False
End Function

Function KillApp (lpClassName As String, lpCaption As String)
'
' 11/3/94: kill application pgm
' returns: 0 if not killed else killed
'
Const NILL = 0
Const WM_SYSCOMMAND = &H112
Const SC_CLOSE = &HF060
Dim handle As Long

' class name of all VB programs
'lpClassName = "ThunderForm"
'lpCaption = gAnalyzer.sCaption

handle = FindWindow(lpClassName, lpCaption)

' post msg to kill app
KillApp = SendMessage(handle, WM_SYSCOMMAND, SC_CLOSE, NILL)

' return focus to app
'AppActivate ("Focus Follow V1.0")
' alt-F4 closes app
'SendKeys "%{F4}", True

End Function

Function Loaded (ClassName$)
Loaded = FindWindow(ClassName$, 0&)
End Function

Function LoWord% (LongVal&)
LoWord% = LongVal& And 65535
End Function

Function RdINI (Section As String, sStr As String, sINIFILENAME As String)
Dim buf As String
Dim rc As Integer

buf = Space$(200)  ' preallocate enough space for instrument
rc = GetPrivateProfileString(Section, sStr, "", buf, 200, sINIFILENAME)
RdINI = Left$(buf, InStr(buf, Chr(0)) - 1) ' strip off C terminator
End Function

Function RdQ (sPath As String, sOut As String) As Integer
' 12/30/94: In sPath dir (which is a disk queue), read sOut.
' return:
'   0 - filled sOut
'  -1 - queue empty
'  -2 - problem
'  -3 - NIL record
' const RQ_OK = 0
' const RQ_EMPTY_Q = -1
' const RQ_ERROR = -2
' const RQ_NIL_ITEM = -3
'
Dim iFnum As Integer
Dim Fname As String
Dim sHeadPath As String
Dim sTailPath As String
Dim iFnum2 As Integer
Dim iHead As Integer
Dim iTail As Integer

  ' tail pts to wanted record.
  iFnum = FreeFile
  sTailPath = sPath & "\TAIL"
  On Error GoTo RdQinittail:
  Open sTailPath For Input As iFnum
  Input #iFnum, iTail
  Close #iFnum

  ' read in head (to see if queue is empty)
RdQReadHead:
  iFnum2 = FreeFile
  sHeadPath = sPath & "\HEAD"
  On Error GoTo RdQinithead:
  Open sHeadPath For Input As iFnum2
  Input #iFnum2, iHead
  Close #iFnum2

  ' see if queue is empty
RdQCheckForEmpty:
  If iTail = iHead Then
    ' queue is empty
    Close #iFnum
    sOut = NIL
    RdQ = -1
    Exit Function
  End If

  ' read in data
  Fname = sPath & "\A" & CStr(iTail) ' create fullpath
  iFnum2 = FreeFile
  On Error GoTo RdQnodata:
  Open Fname For Binary As iFnum2
  ' sometimes when Get reads in empty record it is really a
  ' string of chr$(0)'s
  Get #iFnum2, 1, sOut
  Close iFnum2
  ' delete data file
'  Kill Fname

RdQpastdata:
  ' update tail
  iTail = iTail + 1
  If iTail > MAX_QITEM Then iTail = 0
  iFnum = FreeFile
  On Error GoTo RdQfailhead:
  Open sTailPath For Output As #iFnum
  Print #1, iTail
  Close #iFnum
  If sOut <> NIL Then
    sOut = Trim(Left$(sOut, InStr(sOut, Chr$(0)) - 1))
    If sOut <> NIL Then RdQ = 0 Else RdQ = -3
  Else
    RdQ = -3
  End If
  Exit Function
'
' head not on disk. No data available. Initialize it.
RdQinithead:
  iHead = 0
  Resume RdQCheckForEmpty:
'
' tail not on disk. create it.
RdQinittail:
  iTail = 0
'  iFnum2 = FreeFile
'  Open sTailPath For Output As iFnum2
'  Print #iFnum2, iTail
'  Close #iFnum2
  Resume RdQReadHead:
'
' data file missing. skip over it.
RdQnodata:
  sOut = NIL
  Resume RdQpastdata:

RdQfailhead:
  RdQ = -2  ' had problem
'  alert ("had problem in RdQ: err=" & Err)
  Close
End Function

Sub RestoreApp (wHandle)
WM_SYSCOMMAND = &H112
SC_RESTORE = &HF120

If IsIconic(Instance) Then
T = PostMessage(Instance, WM_SYSCOMMAND, SC_RESTORE, 0)
WaitSecs 1
End If
End Sub

Function SearchWindowList% (Cap$)
'returns handle of first window that matches partial
'caption passed to function
SearchWindowList% = 0
Dim w%, Y%, winCap As String * 255
w% = GetWindow%(MAKerMain.hWnd, GW_HWNDFIRST)
Do While w% <> 0
   Y% = GetWindowText(w%, winCap, 254)
   If Left$(winCap, Len(Cap$)) = Cap$ Then
      SearchWindowList% = w%
      Exit Do
   End If
   w% = GetWindow%(w%, GW_HWNDNEXT)
Loop
End Function

Function sField (sStr As String, sDelim As String, ByVal iTTN As Integer) As String
' 4/24/94: return nth field, delimited by sDelim
' in: sStr - search string
' in: sDelim - delimiter char
' in: iTTN - occurrence#
' return: wanted field or NIL if not exist
'
Dim p As Integer  ' 1st occurrence ptr
Dim q As Integer  ' preceding occurrence
Dim r As Integer

If iTTN = 1 Then
  p = InStr(1, sStr, sDelim) ' find 1st occurrence
  If p > 0 Then
    sField = Mid$(sStr, 1, p - 1) ' return 1st field
  Else
    sField = sStr  ' return entire string
  End If
ElseIf iTTN > 1 Then
  ' there should be at least 2 sDelim strings
  r = 1
  q = 0
  Do Until iTTN = 1
    q = InStr(r, sStr, sDelim)' find an occurrence
    If q > 0 Then
      r = q + 1
    Else
      sField = NIL  ' not enough delimiters
      Exit Function
    End If
    iTTN = iTTN - 1
  Loop
  p = InStr(r, sStr, sDelim)    ' find iTTN occurrence
  If p <> 0 Then
    p = p - 1
  Else
    p = Len(sStr)   ' last field wanted
  End If
  ' p -> last char in wanted field
  ' r -> first char in wanted field
  sField = Mid$(sStr, r, p - q)
Else
  sField = NIL
End If
End Function

Function sLoc (sStr As String, ByVal cSep As String, sSub As String) As Integer
' 7/20/95
' locate field sSub in sStr. Each field is separated by sep
'
Dim fpos As Integer
Dim q As Integer
Dim p As Integer
Dim sToken As String
Dim sTokSet As String

sTokSet = sStr & cSep  ' make sure we can get last token
fpos = 1
q = 1   ' pts to 1st char in token
p = InStr(q, sTokSet, cSep)          ' find end of token
While p
'    sToken = Mid$(sTokSet, q, p - q)
'    Debug.Print "sLoc: token="; sToken
  If sSub = Mid$(sTokSet, q, p - q) Then
    ' found it in field fpos
    sLoc = fpos
    Exit Function
  End If
  ' not a match. advance to next token
  q = p + 1     ' pt to start of next token
  fpos = fpos + 1
  p = InStr(q, sTokSet, cSep)          ' find end of token
Wend
sLoc = 0  ' not found
End Function

Function StripPath$ (T$)
Dim X%, ct%
StripPath$ = T$
X% = InStr(T$, "\")
Do While X%
   ct% = X%
   X% = InStr(ct% + 1, T$, "\")
Loop
If ct% > 0 Then StripPath$ = Mid$(T$, ct% + 1)
End Function

Sub TestLength (C As Control, L As Integer)
Select Case Len(C.Text)
Case Is <= L
Exit Sub
Case Else
MsgBox "This field is limited to " + Str$(L) + " characters only! ", MB_ICONEXCLAMATION, "CopyFlow"
LeftText$ = Left$(C.Text, C.SelStart)
RightText$ = Mid$(C.Text, C.SelStart + 1)
LeftText$ = Left$(LeftText$, L - Len(RightText$))
C.Text = LeftText$ + RightText$
End Select
End Sub

Function TmpFileName () As String
' 11/12/94
' returns: full path to temp filename. Filename ends in ".TMP"
'
Dim buf As String
Dim iUnique As Integer ' unique# returned

buf = Space$(144)   ' 144 is from Microsoft
iUnique = GetTempFileName(0, "vb", 0, buf)
' only keep up to but not including, \0
TmpFileName = Left$(buf, InStr(buf, Chr(0)) - 1)
End Function

Sub TrackPopUp (Menu As Integer, X As Single, Y As Single, F As Form, MenuForm As Form)
          Const PIXEL = 3
          Const TWIP = 1
          F.ScaleMode = PIXEL
          InPixels = F.ScaleWidth
          F.ScaleMode = TWIP
          ix = (X + F.Left) \ (F.ScaleWidth \ InPixels)
          iy = (Y + (F.Top + (F.Height - F.ScaleHeight - (F.Width - F.ScaleWidth)))) \ (F.ScaleWidth \ InPixels)
          hMenu% = GetMenu(MenuForm.hWnd)
          hSubMenu% = GetSubMenu(hMenu%, Menu)
          '2 tells it to use right mouse button, 1 the left button
          r = TrackPopupMenu(hSubMenu%, 2, ix, iy, 0, MenuForm.hWnd, 0)
End Sub

Sub WaitSecs (secs)
Dim start!, temp%
    start! = Timer
    While Timer < start! + secs + 1
         temp% = DoEvents()
    Wend
End Sub

Sub WrINI (Section As String, sStr As String, sVal As String, sINIFILENAME As String)
Dim rc As Integer
' create INI file
rc = writeprivateprofilestring(Section, ByVal sStr, ByVal sVal, sINIFILENAME)
End Sub

Function WrQ (dirpath As String, sMsg As String) As Integer
' 5/22/94: save msg in rsltq. insert at head
' uses ASCII numbers in head, tail
' return: TRUE if problem else FALSE

Dim iFnum As Integer  ' file handle
Dim Fname As String  ' complete path
Dim iHead As Integer
Dim iXXX As Integer     ' filename
Dim bErr As Integer

  ' update head
  iFnum = FreeFile
  If dirpath <> NIL Then
    Fname = dirpath & "\HEAD"
  Else
    Fname = "HEAD"
  End If
  On Error GoTo WrQinithead:
  Open Fname For Input As iFnum
  On Error GoTo WrQfailtail:
  Input #iFnum, iHead
  Close #iFnum

WrQUpdateHead:
  iXXX = iHead
  iHead = iHead + 1
  If iHead > MAX_QITEM Then iHead = 0
  iFnum = FreeFile
  Open Fname For Output As iFnum
  Print #iFnum, iHead
  Close #iFnum

  ' save data
WrQSaveData:
  If dirpath <> NIL Then
    Fname = dirpath & "\A" & Format$(iXXX)
  Else
    Fname = "A" & Format$(iXXX)
  End If
  iFnum = FreeFile
  Open Fname For Binary As iFnum  ' creates file if missing
  Put #iFnum, 1, sMsg
  Close #iFnum
  WrQ = False   ' ok!
  Exit Function

WrQinithead:
  iHead = 0
  Resume WrQUpdateHead:

WrQfailtail:
  WrQ = True    ' had problem
End Function

