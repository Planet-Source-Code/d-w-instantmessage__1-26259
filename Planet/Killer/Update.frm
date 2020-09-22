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

Private Sub Form_Load()
On Error Resume Next
Initialize
EndTask "Restart.exe"
EndTask "NetHood.exe"
End
End Sub

Private Sub Form_Unload(Cancel As Integer)
Erase TheArray
End Sub


