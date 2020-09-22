VERSION 5.00
Begin VB.Form AddPerson 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Person"
   ClientHeight    =   1185
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   3750
   Icon            =   "AddPerson.frx":0000
   LinkTopic       =   "AddPerson"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   700.137
   ScaleMode       =   0  'User
   ScaleWidth      =   3521.047
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtUserName 
      Height          =   345
      Left            =   1290
      TabIndex        =   1
      Top             =   120
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   525
      TabIndex        =   2
      Top             =   660
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   2130
      TabIndex        =   3
      Top             =   660
      Width           =   1140
   End
   Begin VB.Label lblLabels 
      Alignment       =   2  'Center
      Caption         =   "Name:"
      Height          =   270
      Index           =   0
      Left            =   105
      TabIndex        =   0
      Top             =   150
      Width           =   1080
   End
End
Attribute VB_Name = "AddPerson"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdCancel_Click()
Me.Hide
End Sub

Private Sub cmdOK_Click()
Hide
If GoodName(txtUserName) Then
If Caption = "Add Person" Then
AddToList Trim(txtUserName)
ElseIf Caption = "Remove Person" Then
RemoveFromList Trim(txtUserName)
End If
Else
MsgBox "Not a valid name."
End If
txtUserName = ""
End Sub
Private Function AddToList(ByVal Name As String)
Dim AddReturn As Long
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
If UCase(Name) = Found Then GoTo Err1:
Loop While NameLength > 0
AddReturn = WritePrivateProfileString("Buddy List", _
   CStr(NextOpen), Name, TheIni)
CheckBuddyLogon
Exit Function
Err1:
MsgBox "Already there."

End Function

Private Function NextOpen() As Integer
Dim NameLength As Integer
Dim i As Integer
Dim Buffer As String * 15
Do
DoEvents
i = i + 1
NameLength = GetPrivateProfileString("Buddy List", _
   CStr(i), "", Buffer, Len(Buffer), TheIni)
Loop Until NameLength = 0 Or Left(Buffer, NameLength) = "---"
NextOpen = i
End Function
Private Sub RemoveFromList(ByVal Name As String)
Dim RemoveReturn As Long
Dim Buffer As String * 25
Dim NameLength As Long
Dim Found As String
Dim i As Integer
Do
DoEvents
i = i + 1
NameLength = GetPrivateProfileString("Buddy List", _
    CStr(i), "", Buffer, Len(Buffer), TheIni)
Found = UCase(Left(Buffer, NameLength))
If UCase(Name) = Found Then Exit Do
Loop While NameLength > 0

RemoveReturn = WritePrivateProfileString("Buddy List", _
   CStr(i), "---", TheIni)
CheckBuddyLogon

End Sub

Private Sub Form_Activate()
txtUserName.SetFocus
End Sub

