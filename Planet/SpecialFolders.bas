Attribute VB_Name = "SpecialFolders"
Option Explicit

Private Declare Function GetWindowsDirectory Lib "kernel32.dll" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetSystemDirectory Lib "kernel32.dll" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetTempPath Lib "kernel32.dll" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long

Private Type SHITEMID
    cb As Long
    abID As Byte
End Type

Private Type ITEMIDLIST
    mkid As SHITEMID
End Type

Public Enum Folder
WINDOWS = vbNull
WINSYSTEM = -1
DESKTOP = 0
PROGRAMS = 2
DOCUMENTS = 5
FAVORITES = 6
STARTUP = 7
RECENT = 8
SENDTO = 9
STARTMENU = 11
DESKTOPUSER = 16
NETHOOD = 19
FONTFOLDER = 20
SHELLNEW = 21
PRINTHOOD = 27
TEMPINTERNET = 32
COOKIES = 33
HISTORY = 34
TEMP = 99 'puts a backslash on the end
End Enum

Public Function SpecialFolder(Optional TheFolder As Folder = vbNull) As String
Dim ID As ITEMIDLIST
Dim LngRet As Long
Dim ThePath As String
Dim TheLength As Long
ThePath = Space(255)
Select Case TheFolder
Case WINDOWS
TheLength = GetWindowsDirectory(ThePath, 255)
ThePath = Left(ThePath, TheLength)
Case WINSYSTEM
TheLength = GetSystemDirectory(ThePath, 255)
ThePath = Left(ThePath, TheLength)
Case TEMP
TheLength = GetTempPath(255, ThePath)
ThePath = Left(ThePath, TheLength)
Case Else
LngRet = SHGetSpecialFolderLocation(0, TheFolder, ID)
If LngRet = 0 Then
If SHGetPathFromIDList(ID.mkid.cb, ThePath) <> 0 Then
ThePath = Left(ThePath, InStr(ThePath, vbNullChar) - 1)
End If
End If
End Select
SpecialFolder = Trim(ThePath)
End Function



