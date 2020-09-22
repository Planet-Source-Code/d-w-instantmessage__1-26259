Attribute VB_Name = "Globals"
Option Explicit
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
Public Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hwnd As _
    Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As _
    Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long

Public Const SE_ERR_FNF = 2&
Public Const SE_ERR_PNF = 3&
Public Const SE_ERR_ACCESSDENIED = 5&
Public Const SE_ERR_OOM = 8&
Public Const SE_ERR_DLLNOTFOUND = 32&
Public Const SE_ERR_SHARE = 26&
Public Const SE_ERR_ASSOCINCOMPLETE = 27&
Public Const SE_ERR_DDETIMEOUT = 28&
Public Const SE_ERR_DDEFAIL = 29&
Public Const SE_ERR_DDEBUSY = 30&
Public Const SE_ERR_NOASSOC = 31&
Public Const ERROR_BAD_FORMAT = 11&
Public Const SW_HIDE = 0
Public Const SW_NORMAL = 1
Public Const SW_SHOWNORMAL = 1
Public Const SW_SHOWMINIMIZED = 2
Public Const SW_SHOWMAXIMIZED = 3
Public Const SW_RESTORE = 9
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1

Public LoginName As String
Public LoggedOn As Boolean
Public Trying As Boolean
Public Item As ListItem
Public MNum As Integer
Public Messages() As Form
Public MessageLeft As Single
Public MessageTop As Single
Public MessageFile() As String
Public Function GoodName(Name As String) As Boolean
Dim i As Integer
Dim Char(8) As String
Char(0) = "\"
Char(1) = "/"
Char(2) = ":"
Char(3) = "*"
Char(4) = "?"
Char(5) = """"
Char(6) = "<"
Char(7) = ">"
Char(8) = "|"
If Trim(Name) = "" Then
GoodName = False
Exit Function
End If
For i = 0 To 8
If InStr(1, Name, Char(i), vbTextCompare) > 0 Then
GoodName = False
Exit Function
End If
Next
GoodName = True
End Function


Public Sub LogOff()
On Error Resume Next
Kill NetPath & LoginName & "\*.*"
RmDir NetPath & LoginName
Buddy.BuddyList.ListItems.Clear
Set Item = Buddy.BuddyList.ListItems.Add(, , "Co-workers of:   " & LoginName, , 1)
Set Item = Buddy.BuddyList.ListItems.Add
LoggedOn = False
Trying = False
End Sub


Public Function OnList(Name As String) As Boolean
Dim NameLength As Long
Dim i As Integer
Dim Buffer As String * 25
Dim Found As String
Do
DoEvents
i = i + 1
NameLength = GetPrivateProfileString("Buddy List", _
    CStr(i), "", Buffer, Len(Buffer), TheIni)
Found = UCase(Left(Buffer, NameLength))
If UCase(Name) = Found Then
OnList = True
Exit Function
End If
Loop While NameLength > 0
OnList = False
End Function

Public Function Resize() As Single
Select Case Screen.Width
Case 9600
Resize = 1
Case 12000
Resize = 1.25
Case 15360
Resize = 1.6
Case 19200
Resize = 2
Case Else
Resize = 1
End Select
End Function



Public Function NetPath() As String
NetPath = "X:\Drive Copy\Im\"
End Function

Public Function LogOnNetwork() As Boolean
DoEvents
If Dir(NetPath, vbDirectory) = "" Then
GoTo Err1:
End If
If Dir(NetPath & LoginName, vbDirectory) = "" Then
MkDir NetPath & LoginName
End If
LogOnNetwork = True
LoggedOn = True
Exit Function
Err1:
LoggedOn = False
LogOnNetwork = False
Trying = True
End Function
Public Function OldLogon() As String
Dim Buffer As String * 15
Dim BufferLength As Integer
BufferLength = GetPrivateProfileString("Logon", _
  "Name", "", Buffer, Len(Buffer), TheIni)
OldLogon = Left(Buffer, BufferLength)
OldLogon = PrettyName(OldLogon)
End Function
Public Sub CheckBuddyLogon()

On Error GoTo TheEnd
Dim Selected() As String
Dim TheArray() As String
Dim Found As String
Dim i As Integer
Dim j As Integer
Found = Dir(NetPath, vbDirectory)
ReDim TheArray(0)
Do
DoEvents
If Found <> "." And Found <> ".." Then
   If (GetAttr(NetPath & Found) And vbDirectory) = vbDirectory Then
       If Buddy.menuHideThem.Checked And Not OnList(Found) Then
       GoTo Skip:
       End If
   ReDim Preserve TheArray(UBound(TheArray) + 1)
   TheArray(UBound(TheArray)) = Found
Skip:
   End If
End If
Found = Dir
Loop While Found > ""

ReDim Selected(2)
For i = 2 To Buddy.BuddyList.ListItems.Count
ReDim Preserve Selected(i + 1)
If Buddy.BuddyList.ListItems.Item(i).Selected = True Then
Selected(i) = Buddy.BuddyList.ListItems.Item(i).Text
End If
Next

Buddy.BuddyList.ListItems.Clear
Set Item = Buddy.BuddyList.ListItems.Add(, , "Co-workers of:   " & LoginName, , 1)
Set Item = Buddy.BuddyList.ListItems.Add(, , "   Everyone", , 3)
For i = 1 To (UBound(TheArray))
Set Item = Buddy.BuddyList.ListItems.Add(, , TheArray(i), , 3)
Next

For j = 2 To UBound(Selected)
    For i = 2 To Buddy.BuddyList.ListItems.Count
        If Selected(j) = Buddy.BuddyList.ListItems.Item(i).Text Then
        Buddy.BuddyList.ListItems.Item(i).Selected = True
        End If
    Next
Next
Exit Sub
TheEnd:

End Sub
Private Function MsgBx(MsgText As String, Optional Flags As Variant, Optional Title As Variant) As Long
If IsMissing(Title) Then
Title = App.ExeName
End If
If IsMissing(Flags) Then
Flags = vbOKOnly
End If
MsgBx = MessageBox(Screen.ActiveForm.hwnd, MsgText, Title, Flags)
End Function
Public Function TheIni() As String
Select Case Right(App.Path, 1)
Case "\"
TheIni = App.Path & "Im.ini"
Case Else
TheIni = App.Path & "\Im.ini"
End Select
End Function
Public Sub NoTopZ(frm As Form)
Dim lRetVal As Long
lRetVal = SetWindowPos(frm.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOSIZE Or SWP_NOMOVE)
End Sub

Public Sub TopZ(frm As Form)
Dim lRetVal As Long
lRetVal = SetWindowPos(frm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOSIZE Or SWP_NOMOVE)
End Sub
