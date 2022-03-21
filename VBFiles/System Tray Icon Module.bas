Attribute VB_Name = "SysTrayIcon_Module"
Option Explicit

Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Global Const NIM_ADD = 0
Global Const NIM_MODIFY = 1
Global Const NIM_DELETE = 2
Global Const NIF_MESSAGE = 1
Global Const NIF_ICON = 2
Global Const NIF_TIP = 4

Declare Function Shell_NotifyIconA Lib "SHELL32" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Integer

Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_LBUTTONDBLCLICK = &H203
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_RBUTTONDBLCLICK = &H206

