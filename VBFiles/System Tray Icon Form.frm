VERSION 5.00
Begin VB.Form frmSysTray 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "System Tray Icon Demo"
   ClientHeight    =   3195
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Menu mnuPopUp 
      Caption         =   "PopUpMenu"
      Begin VB.Menu mnuPopUpSub 
         Caption         =   "Sub Menu"
      End
   End
End
Attribute VB_Name = "frmSysTray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim xShell As Long

Select Case X
    Case 7680 'MouseMove
    
    Case 7695 'Left MouseDown
    'Place code here to be executed when left mouse button is pressed.
    
    
    Case 7710 'Left MouseUp
    
    Case 7725 'Left DoubleClick
    
    Case 7740 'Right MouseDown
    'Place code here to be executed when right mouse button is pressed.
    
    Me.PopupMenu Me.mnuPopUp  'Show popup menu
    
    Case 7755 'Right MouseUp
    
    Case 7770 'Right DoubleClick
    
End Select

End Sub


Private Function setNOTIFYICONDATA(hwnd As Long, ID As Long, Flags As Long, CallbackMessage As Long, Icon As Long, Tip As String) As NOTIFYICONDATA

Dim nidTemp As NOTIFYICONDATA

nidTemp.cbSize = Len(nidTemp)
nidTemp.hwnd = hwnd
nidTemp.uID = ID
nidTemp.uFlags = Flags
nidTemp.uCallbackMessage = CallbackMessage
nidTemp.hIcon = Icon
nidTemp.szTip = Tip & Chr$(0)

setNOTIFYICONDATA = nidTemp

End Function

Private Sub Form_Unload(Cancel As Integer)

Dim i As Integer
Dim nid As NOTIFYICONDATA

nid = setNOTIFYICONDATA(hwnd:=frmSysTray.hwnd, ID:=vbNull, Flags:=NIF_MESSAGE Or NIF_ICON Or NIF_TIP, CallbackMessage:=vbNull, Icon:=frmSysTray.Icon, Tip:="")
i = Shell_NotifyIconA(NIM_DELETE, nid)

End Sub

Private Sub Form_Load()
    Dim i As Integer  'For holding return value
    Dim s As String  'For storing ToolTip
    Dim nid As NOTIFYICONDATA
    
    'Set tool tip here
    s = "This is the tooltip"
    
    'Icon used is frmMain.Icon
    nid = setNOTIFYICONDATA(hwnd:=frmSysTray.hwnd, ID:=vbNull, Flags:=NIF_MESSAGE Or NIF_ICON Or NIF_TIP, CallbackMessage:=&H200, Icon:=frmMain.Icon, Tip:=s)
    i = Shell_NotifyIconA(NIM_ADD, nid)
    
    Me.Visible = False  'Hide this form
End Sub
