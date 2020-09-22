VERSION 5.00
Begin VB.Form Tray 
   Caption         =   "Instant Message"
   ClientHeight    =   3180
   ClientLeft      =   165
   ClientTop       =   765
   ClientWidth     =   4680
   Icon            =   "Tray.frx":0000
   LinkTopic       =   "Tray"
   ScaleHeight     =   3180
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   1  'Minimized
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuLogon 
         Caption         =   "Logon"
      End
      Begin VB.Menu mnuShow 
         Caption         =   "Show List"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuList 
      Caption         =   "List"
      Begin VB.Menu mnuMess 
         Caption         =   "Send Message"
      End
      Begin VB.Menu mnuAddTo 
         Caption         =   "Add To List"
      End
      Begin VB.Menu mnuRem 
         Caption         =   "Remove From List"
      End
   End
End
Attribute VB_Name = "Tray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
PlaceIcon Me
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim Msg As Long
Msg = x / Screen.TwipsPerPixelX
Select Case Msg
Case WM_LBUTTONDOWN
Case WM_LBUTTONUP
TrayToTitle Buddy
Buddy.ZOrder
Buddy.Show
TopZ Buddy
NoTopZ Buddy
Case WM_LBUTTONDBLCLK
Case WM_RBUTTONDOWN
Case WM_RBUTTONUP
PopupMenu Tray.mnuFile
Case WM_RBUTTONDBLCLK
End Select
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
DestroyIcon Me
End Sub


Private Sub mnuAddTo_Click()
AddPerson.Caption = "Add Person"
AddPerson.txtUserName = Buddy.BuddyList.SelectedItem
AddPerson.Show
AddPerson.txtUserName.SetFocus
End Sub

Private Sub mnuExit_Click()
Unload Buddy
Unload Me
End
End Sub

Private Sub mnuLogon_Click()
With Login.txtUserName
.Text = OldLogon
.SelStart = 0
.SelLength = Len(Login.txtUserName)
End With
Login.Show vbModal
If LoggedOn Then
CheckBuddyLogon
End If
End Sub

Private Sub mnuMess_Click()
On Error GoTo TheEnd:
Dim i As Integer
Dim MessageFile As String
If Buddy.BuddyList.SelectedItem.Index < 2 Then Exit Sub
MessageFile = Trim(Buddy.BuddyList.SelectedItem.Text)
If Buddy.BuddyList.SelectedItem.Index > 2 Then
If Dir(NetPath & MessageFile, vbDirectory) = "" Then
MsgBox "That user has logged off."
Exit Sub
End If
End If
For i = 1 To MNum
If Messages(i).Caption = "Instant Message - " & MessageFile Then
Messages(i).Show
Messages(i).Conversation.SelStart = Len(Messages(i).Text)
Messages(i).Sending.SetFocus
Messages(i).Sending.SelStart = 1
TopZ Messages(i)
NoTopZ Messages(i)
Exit Sub
End If
Next

MNum = MNum + 1
ReDim Preserve Messages(MNum)
Set Messages(MNum) = New Message
Messages(MNum).Left = MessageLeft + (200 * (MNum - 1))
Messages(MNum).Top = MessageTop - (200 * (MNum - 1))
Messages(MNum).Caption = "Instant Message - " & MessageFile
Messages(MNum).ToWho = MessageFile
Messages(MNum).From = LoginName
Messages(MNum).Show
Messages(MNum).Sending.SetFocus

Exit Sub
TheEnd:
End Sub


Private Sub mnuRem_Click()
AddPerson.Caption = "Remove Person"
AddPerson.txtUserName = Buddy.BuddyList.SelectedItem
AddPerson.Show
AddPerson.txtUserName.SetFocus
End Sub


Private Sub mnuShow_Click()
TrayToTitle Buddy
Buddy.ZOrder
Buddy.Show
TopZ Buddy
NoTopZ Buddy
End Sub


