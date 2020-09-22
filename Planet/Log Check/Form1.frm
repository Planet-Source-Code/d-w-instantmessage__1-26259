VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   1920
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   3255
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1920
   ScaleWidth      =   3255
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1440
      Top             =   1200
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TheArray() As String
Private Sub FixLogOff()
On Error Resume Next
Dim Found As String
Dim KillList() As String
Dim MessageFile As String
Dim i As Integer
Dim KIA As Integer
Found = Dir(NetPath, vbDirectory)
ReDim KillList(0)
Do
DoEvents
If Found <> "." And Found <> ".." Then
   If (GetAttr(NetPath & Found) And vbDirectory) = vbDirectory Then
   ReDim Preserve KillList(UBound(KillList) + 1)
   KillList(UBound(KillList)) = Found
   End If
End If
Found = Dir
Loop While Found > ""

For i = 1 To UBound(KillList)
MessageFile = Dir(NetPath & KillList(i) & "\*.*")
Do While MessageFile <> ""
If MessageFile = "test" Then
KIA = KIA + 1
Kill NetPath & KillList(i) & "\*.*"
RmDir NetPath & KillList(i)
Exit Do
End If
MessageFile = Dir
Loop
Next

MsgBox "Killed " & KIA & " folders."
End Sub


Public Sub SendLogonTest()

On Error GoTo TheEnd

Dim FileNumber As Integer
Dim TheArray() As String
Dim Found As String
Dim i As Integer
Dim j As Integer
FileNumber = FreeFile
Found = Dir(NetPath, vbDirectory)
ReDim TheArray(0)
Do
DoEvents
If Found <> "." And Found <> ".." Then
   If (GetAttr(NetPath & Found) And vbDirectory) = vbDirectory Then
   ReDim Preserve TheArray(UBound(TheArray) + 1)
   TheArray(UBound(TheArray)) = Found
   End If
End If
Found = Dir
Loop While Found > ""


For i = 1 To (UBound(TheArray))
Open NetPath & TheArray(i) & "\test" For Output As FileNumber
Print #FileNumber, "goto: test"
Close #FileNumber
Next

MsgBox "Found " & UBound(TheArray) & " folders."
Exit Sub
TheEnd:

End Sub

Public Function NetPath() As String
NetPath = "X:\Drive Copy\Im\"
End Function
Private Sub Form_Load()
MousePointer = vbHourglass
Caption = "IM Logon Check"
Me.AutoRedraw = True
Me.Print
Me.Print
Me.Print
Me.Print "            Checking IM logon status..."
SendLogonTest
End Sub


Private Sub Timer1_Timer()
Static Timeit As Integer
Timeit = Timeit + 1
If Timeit = 10 Then
FixLogOff
End
End If
End Sub


