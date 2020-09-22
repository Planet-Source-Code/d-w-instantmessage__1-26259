VERSION 5.00
Begin VB.Form Message 
   Caption         =   "Instant Message"
   ClientHeight    =   3630
   ClientLeft      =   60
   ClientTop       =   3360
   ClientWidth     =   5040
   Icon            =   "Message.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   5040
   Begin VB.CommandButton Clearer 
      Caption         =   "CLEAR"
      Height          =   405
      Left            =   45
      TabIndex        =   8
      Top             =   3225
      Width           =   1095
   End
   Begin VB.TextBox From 
      Height          =   285
      Left            =   3570
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   30
      Width           =   1425
   End
   Begin VB.TextBox ToWho 
      Height          =   285
      Left            =   945
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   30
      Width           =   1425
   End
   Begin VB.CommandButton Exiter 
      Caption         =   "&CLOSE"
      Height          =   405
      Left            =   3840
      TabIndex        =   3
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton Send 
      Caption         =   "SEND"
      Enabled         =   0   'False
      Height          =   405
      Left            =   2670
      TabIndex        =   2
      Top             =   3240
      Width           =   1095
   End
   Begin VB.TextBox Sending 
      Height          =   705
      Left            =   30
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   2460
      Width           =   4965
   End
   Begin VB.TextBox Conversation 
      Height          =   1845
      Left            =   30
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   360
      Width           =   4965
   End
   Begin VB.Label SentDate 
      Alignment       =   2  'Center
      Height          =   195
      Left            =   60
      TabIndex        =   9
      Top             =   2235
      Width           =   4680
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "From:"
      Height          =   255
      Index           =   1
      Left            =   2850
      TabIndex        =   7
      Top             =   60
      Width           =   675
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "To:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   75
      Width           =   825
   End
End
Attribute VB_Name = "Message"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim OrgHeight As Single
Dim OrgWidth As Single
Dim SendingTop As Single
Dim ButtonTop As Single
Dim ConHeight As Single
Dim BoxWidth As Single
Dim ExitLeft As Single
Dim SendLeft As Single
Dim ToWhom As String
Dim ClearLeft As Single
Dim DateTop As Single

Private Sub Clearer_Click()
Conversation = ""
SentDate = ""
End Sub

Private Sub Conversation_Change()
Conversation.SelStart = Len(Conversation)
End Sub

Private Sub Exiter_Click()
Hide
Sending = ""
End Sub

Private Sub Form_DblClick()
Sending = Lamerized(Sending)
End Sub

Private Sub Form_Load()
OrgHeight = Height
OrgWidth = Width
SendingTop = Sending.Top
ButtonTop = Send.Top
ConHeight = Conversation.Height
BoxWidth = Sending.Width
ExitLeft = Exiter.Left
SendLeft = Send.Left
ClearLeft = Clearer.Left
DateTop = SentDate.Top
End Sub


Private Sub Form_Paint()
Conversation.SelStart = Len(Conversation)
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = 0 Then
Cancel = True
Hide
Sending = ""
End If
End Sub

Private Sub Form_Resize()
On Error Resume Next
If WindowState = vbMinimized Then Exit Sub
Conversation.Height = ConHeight + Height - OrgHeight
Sending.Top = SendingTop + Height - OrgHeight
Exiter.Top = ButtonTop + Height - OrgHeight
Exiter.Left = ExitLeft + Width - OrgWidth
Send.Top = ButtonTop + Height - OrgHeight
Send.Left = SendLeft + Width - OrgWidth
Clearer.Left = ClearLeft + Width - OrgWidth
Clearer.Top = ButtonTop + Height - OrgHeight
Sending.Width = BoxWidth + Width - OrgWidth
Conversation.Width = BoxWidth + Width - OrgWidth
SentDate.Top = DateTop + Height - OrgHeight
End Sub


Private Sub Send_Click()
If SendIt Then
If Conversation = "" Then
Conversation = LoginName & " :>  " & Sending & vbCrLf
Sending = ""
Else
Conversation = Conversation & LoginName & _
   " :>  " & Sending & vbCrLf
Sending = ""
End If
Sending.SetFocus
End If
End Sub

Private Function SendIt() As Boolean
On Error GoTo Err1:
Dim i As Integer
Dim Found As String
Dim FileNumber As Integer
FileNumber = FreeFile
If Me.Tag <> "Multiple" Then
    If ToWho <> "Everyone" Then
    Open NetPath & ToWho & "\" & LoginName For Output As FileNumber
    Print #FileNumber, Sending
    Close #FileNumber
    Else
    Found = Dir(NetPath, vbDirectory)
    Do While Found <> ""
    DoEvents
    If Found <> "." And Found <> ".." Then
       If (GetAttr(NetPath & Found) And vbDirectory) = vbDirectory Then
           If Buddy.menuHideThem.Checked And Not OnList(Found) Then
           GoTo Skip:
           End If
       Open NetPath & Found & "\" & LoginName For Output As FileNumber
       Print #FileNumber, Sending
       Close #FileNumber
Skip:
       End If
    End If
    Found = Dir
    Loop
    End If
SendIt = True
Else
    For i = 1 To UBound(MessageFile)
    Open NetPath & MessageFile(i) & "\" & LoginName For Output As FileNumber
    Print #FileNumber, Sending
    Close #FileNumber
    Next
Erase MessageFile
SendIt = True
End If
Exit Function
Err1:
Close
MsgBox "Error: message not sent."
SendIt = False
End Function
 












 
Private Sub Sending_Change()
If Sending <> "" Then Send.Enabled = True
If Sending = "" Then Send.Enabled = False
End Sub


Public Sub Sending_Click()

End Sub


Private Sub Sending_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If Send.Enabled = True Then
Send_Click
KeyAscii = 0
End If
ElseIf KeyAscii = 27 Then
KeyAscii = 0
Exiter_Click
End If
End Sub

 

