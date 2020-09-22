Attribute VB_Name = "IconTray"
Option Explicit
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function DrawAnimatedRects Lib "user32" (ByVal hwnd As Long, ByVal idAni As Long, lprcFrom As RECT, lprcTo As RECT) As Long
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
Type RECT
 Left As Long
 Top As Long
 Right As Long
 Bottom As Long
End Type
Private Const GW_CHILD = 5
Private Const GW_HWNDFIRST = 0
Private Const GW_HWNDLAST = 1
Private Const GW_HWNDNEXT = 2
Private Const GW_HWNDPREV = 3
Private Const GWL_WNDPROC = (-4)
Private Const IDANI_OPEN = &H1
Private Const IDANI_CLOSE = &H2
Private Const IDANI_CAPTION = &H3
Private Const WM_USER = &H400
Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type
'Mouse_Move X Constants
Public Const RightDoubleClick = 7770
Public Const LeftMouseUp = 7710
Public Const RightMouseUp = 7755
Public Const RightMouseDown = 7740
Public Const LeftDoubleClick = 7725
Public Const MouseMove = 7680
Public Const LeftMouseDown = 7695
Private TrayIcon As NOTIFYICONDATA
Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_RBUTTONDBLCLK = &H206
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_MOUSEMOVE = &H200
Private rctFrom As RECT
Private rctTo As RECT
Private lngTrayHand As Long
Private lngStartMenuHand As Long
Private lngChildHand As Long
Private strClass As String * 255
Private lngClassNameLen As Long
Private lngRetVal As Long
Public Sub ChangeIcon(frm As Form, newIcon As Object)
TrayIcon.cbSize = Len(TrayIcon)
TrayIcon.hwnd = frm.hwnd
TrayIcon.uID = vbNull
TrayIcon.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
TrayIcon.uCallbackMessage = WM_MOUSEMOVE
TrayIcon.hIcon = newIcon
TrayIcon.szTip = frm.Caption + Chr(0)
Shell_NotifyIcon 1, TrayIcon
End Sub

Public Sub ChangeToolTip(frm As Form, Tip As String)
TrayIcon.cbSize = Len(TrayIcon)
TrayIcon.hwnd = frm.hwnd
TrayIcon.uID = vbNull
TrayIcon.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
TrayIcon.uCallbackMessage = WM_MOUSEMOVE
TrayIcon.hIcon = frm.Icon
TrayIcon.szTip = Tip + Chr(0)
Shell_NotifyIcon 1, TrayIcon
End Sub

Public Function TitleToTray(frm As Form)
lngStartMenuHand = FindWindow("Shell_TrayWnd", vbNullString)
lngChildHand = GetWindow(lngStartMenuHand, GW_CHILD)
Do
lngClassNameLen = GetClassName(lngChildHand, strClass, Len(strClass))
If InStr(1, strClass, "TrayNotifyWnd") Then
lngTrayHand = lngChildHand
Exit Do
End If
lngChildHand = GetWindow(lngChildHand, GW_HWNDNEXT)
Loop
lngRetVal = GetWindowRect(frm.hwnd, rctFrom)
lngRetVal = GetWindowRect(lngTrayHand, rctTo)
lngRetVal = DrawAnimatedRects(frm.hwnd, IDANI_OPEN Or IDANI_CAPTION, rctFrom, rctTo)
End Function

Public Function TrayToTitle(frm As Form)
lngStartMenuHand = FindWindow("Shell_TrayWnd", vbNullString)
lngChildHand = GetWindow(lngStartMenuHand, GW_CHILD)
Do
lngClassNameLen = GetClassName(lngChildHand, strClass, Len(strClass))
If InStr(1, strClass, "TrayNotifyWnd") Then
lngTrayHand = lngChildHand
Exit Do
End If
lngChildHand = GetWindow(lngChildHand, GW_HWNDNEXT)
Loop
lngRetVal = GetWindowRect(frm.hwnd, rctFrom)
lngRetVal = GetWindowRect(lngTrayHand, rctTo)
lngRetVal = DrawAnimatedRects(frm.hwnd, IDANI_CLOSE Or IDANI_CAPTION, rctTo, rctFrom)
End Function






Public Sub PlaceIcon(ByRef frm As Form)
TrayIcon.cbSize = Len(TrayIcon)
TrayIcon.hwnd = frm.hwnd
TrayIcon.uID = vbNull
TrayIcon.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
TrayIcon.uCallbackMessage = WM_MOUSEMOVE
TrayIcon.hIcon = frm.Icon
TrayIcon.szTip = frm.Caption + Chr(0)
Shell_NotifyIcon 0, TrayIcon
End Sub
Public Sub DestroyIcon(frm As Form)
TrayIcon.cbSize = Len(TrayIcon)
TrayIcon.hwnd = frm.hwnd
TrayIcon.uID = vbNull
TrayIcon.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
TrayIcon.uCallbackMessage = WM_MOUSEMOVE
TrayIcon.hIcon = frm.Icon
TrayIcon.szTip = Chr(0)
Shell_NotifyIcon 2, TrayIcon
End Sub
