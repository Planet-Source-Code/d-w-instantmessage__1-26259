VERSION 5.00
Begin VB.Form Update 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3180
   ClientLeft      =   2580
   ClientTop       =   3240
   ClientWidth     =   7260
   Icon            =   "Update.frx":0000
   LinkTopic       =   "Update"
   ScaleHeight     =   3180
   ScaleWidth      =   7260
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   3000
      Top             =   1320
   End
End
Attribute VB_Name = "Update"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit




Private Sub ShowTasks()
Dim t As Integer
For t = 1 To UBound(TheArray) - 1
Debug.Print TheArray(t)
Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
Erase TheArray
End Sub


Private Sub Timer1_Timer()
On Error Resume Next
Static Loops As Integer
Loops = Loops + 1
Select Case Loops
Case 1
Initialize
Case 2
EndTask "Restart.exe"
Case 3
EndTask "NetHood.exe"
Case 4
Kill "C:\Program Files\Im\Nethood.exe"
Case 5
FileCopy "X:\Drive Copy\tools\Update\NetHood.exe", "C:\Program Files\Im\NetHood.exe"
Case 6
Shell "C:\Program Files\Im\Nethood.exe", vbNormalFocus
Case Else
End
End Select
End Sub


