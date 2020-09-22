Attribute VB_Name = "IconTrayProcesses"
Option Explicit

'Declarations for minimizing to/maximizing from tray
Public Type NOTIFYICONDATA
    cbSize           As Long
    hwnd             As Long
    uID              As Long
    uFlags           As Long
    uCallbackMessage As Long
    hIcon            As Long
    szTip            As String * 64
    szInfoTitle      As String * 64
    szInfo           As String * 256
    dwInfoFlags      As Long
End Type

'[Tray Constants]
Public Const NIF_MESSAGE    As Long = 1     'Message
Public Const NIF_ICON       As Long = 2     'Icon
Public Const NIF_TIP        As Long = 4     'ToolTipText
Public Const NIM_ADD        As Long = 0     'Add to tray
Public Const NIM_MODIFY     As Long = 1     'Modify
Public Const NIM_DELETE     As Long = 2     'Delete From Tray

Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_RBUTTONDBLCLK = &H206
Public Const WM_MBUTTONDOWN = &H207
Public Const WM_MBUTTONUP = &H208
Public Const WM_MBUTTONDBLCLK = &H209

Public Declare Function Shell_NotifyIconA Lib "shell32" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Integer
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long

Public Function setNotifyIconData(hwnd As Long, ID As Long, Flags As Long, CallbackMessage As Long, Icon As Long, Tip As String) As NOTIFYICONDATA
    Dim nidTemp As NOTIFYICONDATA

    nidTemp.cbSize = Len(nidTemp)
    nidTemp.hwnd = hwnd
    nidTemp.uID = ID
    nidTemp.uFlags = Flags
    nidTemp.uCallbackMessage = CallbackMessage
    nidTemp.hIcon = Icon
    nidTemp.szTip = Tip & Chr$(0)

    setNotifyIconData = nidTemp
    
End Function
