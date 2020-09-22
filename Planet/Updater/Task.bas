Attribute VB_Name = "Task"
Option Explicit

Private Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szexeFile As String * MAX_PATH
End Type

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type
Private Declare Function TerminateProcess Lib "kernel32" (ByVal ApphProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal blnheritHandle As Long, ByVal dwAppProcessId As Long) As Long
Private Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, lProcessID As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function RegisterServiceProcess Lib "kernel32" (ByVal ProcessID As Long, ByVal ServiceFlags As Long) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Function GetVersionEx Lib "kernel32.dll" Alias "GetVersionExA" (OSInfo As OSVERSIONINFO) As Long
Private Declare Function EnumProcesses Lib "psapi.dll" (ByRef lpidProcess As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Private Declare Function GetModuleFileNameExA Lib "psapi.dll" (ByVal hProcess As Long, ByVal hModule As Long, ByVal ModuleName As String, ByVal nSize As Long) As Long
Private Declare Function EnumProcessModules Lib "psapi.dll" (ByVal hProcess As Long, ByRef lphModule As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long

Private Const MAX_PATH& = 260
Private Const PROCESS_ALL_ACCESS = 0
Private Const TH32CS_SNAPPROCESS As Long = 2&


Public TheArray() As String
Public TheId() As Long
Private A_Process As PROCESSENTRY32
Private TheProcess As Long
Dim lngCbNeeded2 As Long
Dim lngCb As Long
Dim lngCbNeeded As Long
Dim Modules(1 To 200) As Long
Dim lRet As Long
Dim TheModule As String
Dim lngHProcess As Long
Public Function EndTask(TaskName As String) As Boolean
Dim x As Integer
For x = 1 To UBound(TheArray) - 1
If Right(TheArray(x), Len(TaskName)) = LCase(TaskName) Then
TheProcess = OpenProcess(PROCESS_ALL_ACCESS, False, TheId(x))
TerminateProcess TheProcess, 0
CloseHandle TheProcess
EndTask = True
Exit Function
End If
Next
EndTask = False
End Function

Public Sub Initialize()
Dim ProcessFound As Long
Dim TheSnapshot As Long
Dim ExeName As String
Dim i As Integer

If OS = 1 Then
A_Process.dwSize = Len(A_Process)
TheSnapshot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
ProcessFound = ProcessFirst(TheSnapshot, A_Process)
ReDim Preserve TheArray(1)
ReDim Preserve TheId(1)
Do While ProcessFound
i = InStr(1, A_Process.szexeFile, Chr(0))
ExeName = LCase(Left(A_Process.szexeFile, i - 1))
TheArray(UBound(TheArray)) = ExeName
TheId(UBound(TheId)) = A_Process.th32ProcessID
ProcessFound = ProcessNext(TheSnapshot, A_Process)
ReDim Preserve TheArray(UBound(TheArray) + 1)
ReDim Preserve TheId(UBound(TheId) + 1)
Loop
CloseHandle TheSnapshot
Else
lngCb = 8
    lngCbNeeded = 96
    Do While lngCb <= lngCbNeeded
       lngCb = lngCb * 2
       ReDim TheId(lngCb / 4) As Long
       lRet = EnumProcesses(TheId(1), lngCb, lngCbNeeded)
    Loop
    For ProcessFound = 1 To lngCbNeeded / 4
       lngHProcess = OpenProcess(PROCESS_QUERY_INFORMATION _
          Or PROCESS_VM_READ, 0, TheId(ProcessFound))
       If lngHProcess <> 0 Then
           lRet = EnumProcessModules(lngHProcess, Modules(1), 200, lngCbNeeded2)
           If lRet <> 0 Then
           TheModule = Space(MAX_PATH)
           lRet = GetModuleFileNameExA(lngHProcess, Modules(1), TheModule, 500)
           ReDim Preserve TheArray(i)
           TheArray(i) = Trim(TheModule)
           i = i + 1
           End If
       End If
     CloseHandle lngHProcess
    Next
End If
End Sub

Private Function OS() As Byte
   
Dim typOSInfo As OSVERSIONINFO
On Error GoTo Err_OS

typOSInfo.dwOSVersionInfoSize = Len(typOSInfo)

If GetVersionEx(typOSInfo) = 0 Then
    OS = 0
    Exit Function
End If

With typOSInfo
   If .dwPlatformId = 1 And _
                .dwMajorVersion = 4 And _
                .dwMinorVersion = 0 Then
          OS = 1                             ' Windows 95
   ElseIf .dwPlatformId = 1 And _
                .dwMajorVersion = 4 And _
                .dwMinorVersion = 10 Then
          OS = 1                             ' Windows 98
   ElseIf .dwPlatformId = 1 And _
                .dwMajorVersion = 4 And _
                .dwMinorVersion > 10 Then
          OS = 1                             ' Windows Me
   ElseIf .dwPlatformId = 2 And _
                .dwMajorVersion = 3 And _
                .dwMinorVersion = 51 Then
          OS = 2                             ' Windows NT 3.51
   ElseIf .dwPlatformId = 2 And _
                .dwMajorVersion = 4 And _
                .dwMinorVersion = 0 Then
          OS = 2                            ' Windows NT 4.0
   ElseIf .dwPlatformId = 2 And _
                .dwMajorVersion = 5 And _
                .dwMinorVersion = 0 Then
          OS = 2                            ' Windows NT 5.0  / Windows 2000
   ElseIf .dwPlatformId = 2 And _
                .dwMajorVersion = 5 And _
                .dwMinorVersion > 0 Then
          OS = 2
   Else
          OS = 0                            ' Unknown
   End If
End With
Exit Function

Err_OS:
OS = 0
Err.Clear
End Function



