Attribute VB_Name = "OSVersion"
Option Explicit
Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Public Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

Public Const VER_INFO_SIZE& = 148
Public Const VER_PLATFORM_WIN32_NT& = 2
Public Const VER_PLATFORM_WIN32_WINDOWS& = 1

Public myVer As OSVERSIONINFO

Public Function OS() As Integer
Dim RtnVal As Long
myVer.dwOSVersionInfoSize = VER_INFO_SIZE
RtnVal = GetVersionEx(myVer)
If myVer.dwPlatformId = VER_PLATFORM_WIN32_WINDOWS Then
OS = 1
ElseIf myVer.dwPlatformId = VER_PLATFORM_WIN32_NT Then
OS = 2
Else
OS = 3
End If
End Function



