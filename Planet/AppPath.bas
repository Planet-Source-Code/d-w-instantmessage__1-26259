Attribute VB_Name = "AppPath"
Option Explicit

Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" _
    (ByVal hkey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, _
    ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" _
    (ByVal hkey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, _
    lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Byte, _
    lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hkey As Long) As Long

Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const KEY_ALL_ACCESS = &H3F

Public Function LongName(ShortName As String) As String
    
Dim Temp As String
Dim NewString As String
Dim Searched As Boolean
Dim i As Integer

If Len(ShortName) = 0 Then Exit Function

Temp = ShortName
If Right(Temp, 1) = "\" Then
Temp = Left(Temp, Len(Temp) - 1)
Searched = True
End If

On Error GoTo NoFile:
If InStr(Temp, "\") Then
    NewString = ""
    Do While InStr(Temp, "\")
        If Len(NewString) Then
        NewString = Dir(Temp, 55) & "\" & NewString
        Else
        NewString = Dir(Temp, 55)
            If NewString = "" Then
            LongName = ShortName
            Exit Function
            End If
        End If
        On Error Resume Next
        For i = Len(Temp) To 1 Step -1
            If ("\" = Mid(Temp, i, 1)) Then
            Exit For
            End If
        Next
        Temp = Left(Temp, i - 1)
    Loop
    NewString = Temp & "\" & NewString
Else
NewString = Dir(Temp, 55)
End If

Here:
If Searched Then
NewString = NewString & "\"
End If

LongName = PrettyPath(NewString)
Exit Function
NoFile:
NewString = ""
Resume Here:
End Function

Public Function GetAppPath(ByVal AppName As String) As String
On Error GoTo TheEnd:
Dim TheResult As Long
Dim Index As Long
Dim TheEntry As String
Dim EntryLength As Long
Dim TheDataType As Long
Dim TheByteArray(1 To 1024) As Byte
Dim DataLength As Long
Dim ByteValue As String
Dim i As Integer
Dim MainKey As Long
Dim SubKey As String
Dim mKey As Long

If LCase(Right(AppName, 4)) <> ".exe" Then
AppName = AppName & ".exe"
End If

MainKey = HKEY_LOCAL_MACHINE
SubKey = "Software\Microsoft\Windows\CurrentVersion\App Paths\" & AppName

TheResult = RegOpenKeyEx(MainKey, SubKey, 0, KEY_ALL_ACCESS, mKey)

If TheResult <> 0 Then Exit Function
'looked for it and failed

Index = 0
Do
EntryLength = 1024
DataLength = 1024
TheEntry = Space(EntryLength)
TheResult = RegEnumValue(mKey, Index, TheEntry, EntryLength, 0, _
     TheDataType, TheByteArray(1), DataLength)
'looks like we just have to pass just the first element
'of the array to have it filled...
If TheResult <> 0 Then Exit Do

TheEntry = Left(TheEntry, EntryLength)

If Len(TheEntry) = 0 Then
'looking for (Default), empty string

    ByteValue = ""
    For i = 1 To DataLength - 1
    ByteValue = ByteValue & Chr(TheByteArray(i))
    Next
    
    If ByteValue <> "" Then
    GetAppPath = LongName(ByteValue)
    RegCloseKey mKey
    Exit Function
    End If

End If
Index = Index + 1
Loop
GetAppPath = ""
RegCloseKey mKey
Exit Function
TheEnd:
GetAppPath = ""
End Function

Public Function PrettyName(TheName As String) As String
On Error GoTo TheEnd:
'This is designed for human names not path names.
'Middle initial not supported but could be added.
Dim Name As String
Dim Spot As Integer
Dim First As String
Dim Last As String
Dim i As Integer

Name = TheName
Name = Replace(Name, "    ", " ", , , vbTextCompare)
Name = Replace(Name, "   ", " ", , , vbTextCompare)
Name = Replace(Name, "  ", " ", , , vbTextCompare)

First = Left(Name, 1)
Name = LCase(Name)
First = UCase(First)


If InStr(1, Name, ",") > 0 Then
Spot = InStr(1, Name, ",")
Last = Mid(Name, Spot + 1, 1)
    Do
        i = i + 1
        If Last = " " Then
        Last = Mid(Name, Spot + i, 1)
        End If
    Loop Until Last <> " " Or i = 3
Last = UCase(Last)
Name = First & Mid(Name, 2, Spot - 2 + i) & Last & Mid(Name, Spot + 1 + i, Len(Name) - Spot + i)
PrettyName = Name

Else
    If InStr(1, Name, " ") = 0 Then
    Name = First & Mid(Name, 2)
    PrettyName = Name
    Else
    Spot = InStr(1, Name, " ")
    Last = Mid(Name, Spot + 1, 1)
    Last = UCase(Last)
    Name = First & Mid(Name, 2, Spot - 1) & Last & Mid(Name, Spot + 2, Len(Name) - Spot + 1)
    PrettyName = Name
    End If
End If
Exit Function
TheEnd:
PrettyName = TheName
End Function

Public Function PrettyPath(ThePath As String) As String

On Error GoTo TheEnd:

Dim Path As String
Dim Start As Integer
Dim Temp As String

Path = ThePath
Path = LCase(Path)

Temp = Left(Path, 1)
Temp = UCase(Temp)
Path = Temp & Right(Path, Len(Path) - 1)

Start = 1
Do
Start = InStr(Start, Path, "\")
If Start = 0 Then Exit Do
Mid(Path, Start + 1, 1) = UCase(Mid(Path, Start + 1, 1))
Start = Start + 1
Loop While Start < Len(ThePath)

Start = 1
Do
Start = InStr(Start, Path, " ")
If Start = 0 Then Exit Do
Mid(Path, Start + 1, 1) = UCase(Mid(Path, Start + 1, 1))
Start = Start + 1
Loop While Start < Len(Path)

PrettyPath = Path
Exit Function
TheEnd:
PrettyPath = ThePath
End Function

Public Function EnvironPath(AppName As String) As String
On Error GoTo TheEnd:
Dim i As Integer
Dim EnvString As String
Dim Test As String
Dim Paths() As String
Dim Start As Long
Dim Length As Long


EnvString = Environ("Path")
EnvString = EnvString & ";" 'to tell us where the end is

Start = 1
Length = InStr(1, EnvString, ";", vbBinaryCompare) - 1
Do 'start parsing the Path environment
i = i + 1
ReDim Preserve Paths(i)
Paths(i) = Mid(EnvString, Start, Length)
Start = Start + Len(Paths(i)) + 1
Length = InStr(Start, EnvString, ";", vbBinaryCompare) - Start
Loop While Length > 0

For i = 1 To UBound(Paths)
    If InStr(1, AppName, ".") = 0 Then
        If Dir(Paths(i) & "\" & AppName & ".exe", 55) <> "" Then
        EnvironPath = Paths(i) & "\" & AppName & ".exe"
        ElseIf Dir(Paths(i) & "\" & AppName & ".com", 55) <> "" Then
        EnvironPath = Paths(i) & "\" & AppName & ".com"
        ElseIf Dir(Paths(i) & "\" & AppName & ".bat", 55) <> "" Then
        EnvironPath = Paths(i) & "\" & AppName & ".bat"
        ElseIf Dir(Paths(i) & "\" & AppName & ".pif", 55) <> "" Then
        EnvironPath = Paths(i) & "\" & AppName & ".pif"
        ElseIf Dir(Paths(i) & "\" & AppName & ".scr", 55) <> "" Then
        EnvironPath = Paths(i) & "\" & AppName & ".scr"
        End If
    Else
        If Dir(Paths(i) & "\" & AppName, 55) <> "" Then
        EnvironPath = Paths(i) & "\" & AppName
        End If
    End If
Next
EnvironPath = PrettyPath(EnvironPath)
Exit Function
TheEnd:

End Function


