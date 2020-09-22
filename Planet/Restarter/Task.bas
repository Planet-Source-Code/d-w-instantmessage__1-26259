Attribute VB_Name = "Task"
Option Explicit

Private Declare Function TerminateProcess Lib "kernel32" (ByVal ApphProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal blnheritHandle As Long, ByVal dwAppProcessId As Long) As Long
Private Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, lProcessID As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function RegisterServiceProcess Lib "kernel32" (ByVal ProcessID As Long, ByVal ServiceFlags As Long) As Long
Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Function Thread32First Lib "kernel32" (ByVal hSnapshot As Long, uProcess As THREADENTRY32) As Long
Private Declare Function Thread32Next Lib "kernel32" (ByVal hSnapshot As Long, uThread As THREADENTRY32) As Long
Private Declare Function IsHungThread Lib "user32" (ByVal lThreadID As Long) As Long

Private Const MAX_PATH& = 260
Private Const PROCESS_ALL_ACCESS = 0
Private Const TH32CS_SNAPPROCESS As Long = 2&

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
Private Type THREADENTRY32
    lSize As Long
    lUsage As Long
    lThreadID As Long
    lOwnerProcessID As Long
    lBasePri As Long
    lDeltaPri As Long
    lFlags As Long
End Type
Public TheArray() As String
Public TheId() As Long
Public TheThread() As Long
Private TheProcess As Long

Public Function CheckTask(TaskName As String) As Boolean
On Error GoTo TheEnd
Dim x As Integer
For x = 1 To UBound(TheArray) - 1
If Right(TheArray(x), Len(TaskName)) = LCase(TaskName) Then
If TheThread(x) <> 0 Then
TheProcess = OpenProcess(PROCESS_ALL_ACCESS, False, TheId(x))
TerminateProcess TheProcess, 0
CloseHandle TheProcess
CheckTask = False
Exit Function
End If
CheckTask = True
Exit Function
End If
Next
CheckTask = False
TheEnd:
End Function
Public Function EndTask(TaskName As String) As Boolean
Dim x As Integer
On Error GoTo TheEnd:
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
TheEnd:
End Function

Public Sub Initialize()

Dim A_Process As PROCESSENTRY32
Dim A_Thread As THREADENTRY32
Dim ProcessFound As Long
Dim ThreadFound As Long
Dim TheSnapshot As Long
Dim ExeName As String
Dim i As Integer
On Error GoTo TheEnd:

A_Process.dwSize = Len(A_Process)
TheSnapshot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
ProcessFound = ProcessFirst(TheSnapshot, A_Process)
ThreadFound = Thread32First(TheSnapshot, A_Thread)
ReDim Preserve TheArray(1)
ReDim Preserve TheId(1)
ReDim Preserve TheThread(1)

Do While ProcessFound
i = InStr(1, A_Process.szexeFile, Chr(0))
ExeName = LCase(Left(A_Process.szexeFile, i - 1))
TheArray(UBound(TheArray)) = ExeName
TheId(UBound(TheId)) = A_Process.th32ProcessID
TheThread(UBound(TheThread)) = IsHungThread(A_Thread.lThreadID)
ProcessFound = ProcessNext(TheSnapshot, A_Process)
ThreadFound = Thread32Next(TheSnapshot, A_Thread)

ReDim Preserve TheArray(UBound(TheArray) + 1)
ReDim Preserve TheId(UBound(TheId) + 1)
ReDim Preserve TheThread(UBound(TheThread) + 1)
Loop

CloseHandle TheSnapshot
TheEnd:
End Sub


