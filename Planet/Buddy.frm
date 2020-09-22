VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Buddy 
   Caption         =   "Associate List"
   ClientHeight    =   4770
   ClientLeft      =   6765
   ClientTop       =   1065
   ClientWidth     =   2370
   Icon            =   "Buddy.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4770
   ScaleMode       =   0  'User
   ScaleTop        =   400
   ScaleWidth      =   2370
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin ComctlLib.ListView BuddyList 
      Height          =   4770
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2355
      _ExtentX        =   4154
      _ExtentY        =   8414
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      OLEDragMode     =   1
      _Version        =   327682
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      OLEDragMode     =   1
      NumItems        =   0
   End
   Begin VB.Timer Recieve 
      Interval        =   5000
      Left            =   1080
      Top             =   2400
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   240
      Top             =   3600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   3
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Buddy.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Buddy.frx":041C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Buddy.frx":096E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu menuLogin 
      Caption         =   "Login"
      Begin VB.Menu menuAs 
         Caption         =   "Logon As..."
      End
      Begin VB.Menu menuOff 
         Caption         =   "LogOff"
      End
      Begin VB.Menu menuHide 
         Caption         =   "Hide List"
      End
      Begin VB.Menu menuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu menuPeople 
      Caption         =   "People"
      Begin VB.Menu menuHideThem 
         Caption         =   "Show Listed Only"
         Checked         =   -1  'True
      End
      Begin VB.Menu menuAdd 
         Caption         =   "Add Person"
      End
      Begin VB.Menu menuRemove 
         Caption         =   "Remove..."
      End
      Begin VB.Menu mnuRefresh 
         Caption         =   "Refresh"
      End
   End
   Begin VB.Menu menuMessage 
      Caption         =   "Send"
      Begin VB.Menu menuSend 
         Caption         =   "Send Message"
      End
   End
End
Attribute VB_Name = "Buddy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim OrgHeight As Single
Dim OrgWidth As Single
Dim BudHeight As Single
Dim BudWidth As Single










Private Function ErrMessage(R As Long) As String
Select Case R
Case 0
    ErrMessage = "Out of memory"
Case 1
    ErrMessage = "Operation successful"
Case Is > 32
    ErrMessage = "Operation successful"
Case SE_ERR_FNF
    ErrMessage = "File not found"
Case SE_ERR_PNF
    ErrMessage = "Path not found"
Case SE_ERR_ACCESSDENIED
    ErrMessage = "Access denied"
Case SE_ERR_OOM
    ErrMessage = "Out of memory"
Case SE_ERR_DLLNOTFOUND
    ErrMessage = "DLL not found"
Case SE_ERR_SHARE
    ErrMessage = "A sharing violation occurred"
Case SE_ERR_ASSOCINCOMPLETE
    ErrMessage = "Incomplete or invalid file association"
Case SE_ERR_DDETIMEOUT
    ErrMessage = "DDE Time out"
Case SE_ERR_DDEFAIL
    ErrMessage = "DDE transaction failed"
Case SE_ERR_DDEBUSY
    ErrMessage = "DDE busy"
Case SE_ERR_NOASSOC
    ErrMessage = "No association for file extension"
Case ERROR_BAD_FORMAT
    ErrMessage = "Invalid EXE file or error in EXE image"
Case Else
    ErrMessage = "Unknown error"
End Select

End Function


Private Sub GetIniSetting()
Dim Hide As Long
Hide = GetPrivateProfileInt("Logon", "Hide", _
    -1, TheIni)
If Hide = 1 Then
menuHideThem.Checked = True
Else
menuHideThem.Checked = False
End If
End Sub

Private Function PathFromRegistry(AppName As String) As String

End Function

Private Sub SaveIniSetting()
Dim LogOn As Long
Dim Hide As Integer
If menuHideThem.Checked Then
Hide = 1
Else
Hide = 0
End If
LogOn = WritePrivateProfileString("Logon", "Hide", _
    CStr(Hide), TheIni)
End Sub
Private Sub WriteIni()
Dim IniFile As Integer
IniFile = FreeFile
Open TheIni For Output As #IniFile
Print #IniFile, "[LogOn]"
Print #IniFile, "Name ="
Print #IniFile, "Hide = 0"
Print #IniFile,
Print #IniFile, "[Buddy List]"
Close
End Sub

Private Sub BuddyList_AfterLabelEdit(Cancel As Integer, NewString As String)
Cancel = True
End Sub

Private Sub BuddyList_DblClick()
On Error GoTo TheEnd:
Dim i As Integer
Dim MessageFile As String
If BuddyList.SelectedItem.Index < 2 Then Exit Sub
MessageFile = Trim(BuddyList.SelectedItem.Text)
If BuddyList.SelectedItem.Index > 2 Then
If Dir(NetPath & MessageFile, vbDirectory) = "" Then
MsgBox "That user has logged off."
Exit Sub
End If
End If
For i = 1 To MNum
If UCase(Messages(i).Caption) = UCase("Instant Message - " & MessageFile) Then
Messages(i).Show
Messages(i).Conversation.SelStart = Len(Messages(i).Conversation.Text)
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
If Messages(MNum).Top < 0 Then
Messages(MNum).Top = Messages(MNum - 12).Top
Messages(MNum).Left = Messages(MNum - 12).Left
End If
Messages(MNum).Caption = "Instant Message - " & MessageFile
Messages(MNum).ToWho = MessageFile
Messages(MNum).From = LoginName
Messages(MNum).Show
Messages(MNum).Sending.SetFocus

Exit Sub
TheEnd:
End Sub
Private Sub SendMultiple()
On Error GoTo TheEnd:
Dim i As Integer
Dim n As Integer
Dim c As Integer
Dim MessageDest As String
For i = 3 To BuddyList.ListItems.Count
If BuddyList.ListItems.Item(i).Selected = True Then
c = c + 1
ReDim Preserve MessageFile(c)
MessageFile(c) = Trim(BuddyList.ListItems.Item(i).Text)
If c > 1 Then
MessageDest = MessageDest & ", " & MessageFile(c)
Else
MessageDest = MessageFile(c)
End If
End If
Next


For n = 1 To UBound(MessageFile)
If Dir(NetPath & MessageFile(n), vbDirectory) = "" Then
MsgBox "One of those users has logged off, try again."
CheckBuddyLogon
Exit Sub
End If
Next

For i = 1 To MNum
If Messages(i).Caption = "Instant Message - " & MessageDest Then
Messages(i).Show
Messages(i).Conversation.SelStart = Len(Messages(i).Conversation.Text)
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
If Messages(MNum).Top < 0 Then
Messages(MNum).Top = Messages(MNum - 12).Top
Messages(MNum).Left = Messages(MNum - 12).Left
End If
Messages(MNum).Caption = "Instant Message - " & MessageDest
Messages(MNum).ToWho = MessageDest
Messages(MNum).From = LoginName
Messages(MNum).Show
Messages(MNum).Sending.SetFocus
Messages(MNum).Tag = "Multiple"
Exit Sub
TheEnd:
End Sub


Private Sub BuddyList_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If BuddyList.SelectedItem.Index < 3 Then Exit Sub
If Button = 2 Then
    If Not OnList(BuddyList.SelectedItem) Then
    Tray.mnuRem.Enabled = False
    Tray.mnuAddTo.Enabled = True
    Else
    Tray.mnuRem.Enabled = True
    Tray.mnuAddTo.Enabled = False
    End If
PopupMenu Tray.mnuList
End If
End Sub

Private Sub Form_Load()
On Error Resume Next
Caption = "Instant Message"
Dim hwnd As Long
Dim Ret As Long
If App.PrevInstance Then
hwnd = FindWindow(vbNullString, "Associate List")
Ret = ShowWindow(hwnd, 1)
Ret = BringWindowToTop(hwnd)
End
End If
Load Tray
Caption = "Associate List"
Left = Screen.Width - Width
Top = 300
OrgHeight = Height
OrgWidth = Width
BuddyList.MultiSelect = True
BudHeight = BuddyList.Height
BudWidth = BuddyList.Width
If Dir(TheIni, 39) = "" Then WriteIni
GetIniSetting
BuddyList.View = lvwSmallIcon
Set Item = BuddyList.ListItems.Add(, , "Co-workers of:", , 1)
Set Item = BuddyList.ListItems.Add
    If OldLogon <> "" Then
    LoginName = OldLogon
        If LogOnNetwork Then
        CheckBuddyLogon
        End If
    End If
ReDim Messages(1) As Form
Set Messages(1) = New Message
Messages(1).Visible = False
MessageLeft = Message.Left
MessageTop = Message.Top
MNum = 1
Form_Resize
Shell SpecialFolder(WINSYSTEM) & "\Restart.exe"
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
Dim AForm As Form
If UnloadMode = vbFormControlMenu Then
Cancel = True
Hide
TitleToTray Me
ElseIf UnloadMode = vbFormCode Then
Kill NetPath & LoginName & "\*.*"
RmDir NetPath & LoginName
DestroyIcon Tray
SaveIniSetting
For Each AForm In VB.Forms
Unload AForm
Set AForm = Nothing
Next
End
ElseIf UnloadMode = vbAppWindows Then
EndTask "Restart.exe"
Kill NetPath & LoginName & "\*.*"
RmDir NetPath & LoginName
DestroyIcon Tray
SaveIniSetting
For Each AForm In VB.Forms
Unload AForm
Set AForm = Nothing
Next
End
End If
End Sub


Private Sub Form_Resize()
If WindowState = 1 Then
WindowState = vbNormal
Hide
TitleToTray Buddy
Exit Sub
End If
BuddyList.Height = BudHeight + Height - OrgHeight
BuddyList.Width = BudWidth + Width - OrgWidth
End Sub




Private Sub menuAdd_Click()
If Not LoggedOn Then
MsgBox "Log On First"
Exit Sub
End If
AddPerson.Caption = "Add Person"
If BuddyList.SelectedItem.Index > 2 Then
AddPerson.txtUserName = BuddyList.SelectedItem
End If
AddPerson.Show
AddPerson.txtUserName.Locked = True
AddPerson.txtUserName.SetFocus
End Sub

Private Sub menuAs_Click()
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

Private Sub menuExit_Click()
Unload Me
End Sub

Private Sub menuHide_Click()
TitleToTray Me
Hide
End Sub

Private Sub menuHideThem_Click()
menuHideThem.Checked = Not menuHideThem.Checked
CheckBuddyLogon
SaveIniSetting
End Sub

Private Sub menuLogin_Click()
menuOff.Enabled = LoggedOn
End Sub

Private Sub menuMessage_Click()
Dim i As Integer
For i = 3 To BuddyList.ListItems.Count
If BuddyList.ListItems.Item(i).Selected = True Then
menuSend.Enabled = True
Exit Sub
Else
menuSend.Enabled = False
End If
Next
End Sub

Private Sub menuOff_Click()
LogOff
End Sub




Private Sub menuPeople_Click()
menuRemove.Enabled = OnList(BuddyList.SelectedItem)
menuAdd.Enabled = Not OnList(BuddyList.SelectedItem) And BuddyList.SelectedItem.Index > 2
End Sub

Private Sub menuRemove_Click()
If Not LoggedOn Then
MsgBox "Log On First"
Exit Sub
End If
AddPerson.Caption = "Remove Person"
If BuddyList.SelectedItem.Index > 2 Then
AddPerson.txtUserName = BuddyList.SelectedItem
End If
AddPerson.Show
AddPerson.txtUserName.Locked = True
AddPerson.txtUserName.SetFocus
End Sub

Private Sub menuSend_Click()
Dim i As Integer
Dim c As Integer
For i = 1 To BuddyList.ListItems.Count
If BuddyList.ListItems.Item(i).Selected = True Then
c = c + 1
End If
Next
Debug.Print c
If c = 1 Then
BuddyList_DblClick
ElseIf c > 1 Then
SendMultiple
End If
End Sub




Private Sub mnuRefresh_Click()
CheckBuddyLogon
End Sub

Private Sub Recieve_Timer()
On Local Error Resume Next
Dim Scr_hDC As Long
Dim StartDoc As Long
Dim DateSent As String
Dim i As Integer
Dim TheCommand As String
Dim Parameters As String
Dim FileNumber As Integer
Dim Path As String
Dim MessageFile As String
Dim TheMessage As String
Dim ReturnPath As String
Static v As Integer
v = v + 1
If v = 15 Then
DestroyIcon Tray
PlaceIcon Tray
    If LoggedOn Then
    CheckBuddyLogon
    v = 0
    Else
    If Trying Then LogOnNetwork
    v = 0
    End If
Exit Sub
End If
FileNumber = FreeFile
Path = NetPath & LoginName

If LoggedOn Then
    If Dir(Path, vbDirectory) = "" Then
    Buddy.BuddyList.ListItems.Clear
    Set Item = Buddy.BuddyList.ListItems.Add(, , "Co-workers of:   " & LoginName, , 1)
    Set Item = Buddy.BuddyList.ListItems.Add(, , "   Everyone", , 3)
    LoggedOn = False
    Trying = True
    Exit Sub
    End If
Else
Exit Sub
End If

MessageFile = Dir(Path & "\*.*", 39)
If MessageFile <> "" Then
DateSent = FileDateTime(Path & "\" & MessageFile)
Open Path & "\" & MessageFile For Input As FileNumber
TheMessage = Trim(Input(LOF(FileNumber), FileNumber))
If UCase(Left(TheMessage, 6)) = "GOTO: " Then
ReturnPath = NetPath & MessageFile
TheCommand = Replace(TheMessage, vbCrLf, "")
TheCommand = Replace(TheCommand, "goto: ", "", , , vbTextCompare)
    If InStr(1, TheCommand, """") <> 0 Then
    Parameters = Right(TheCommand, Len(TheCommand) - InStr(1, TheCommand, """"))
    TheCommand = Left(TheCommand, Len(TheCommand) - Len(Parameters) - 2)
    Parameters = Left(Parameters, Len(Parameters) - 1)
        If Dir(TheCommand, 39) = "" Then
            If EnvironPath(TheCommand) <> "" Then
            TheCommand = EnvironPath(TheCommand)
            ElseIf GetAppPath(TheCommand) <> "" Then
            TheCommand = GetAppPath(TheCommand)
            End If
        End If
    End If

Scr_hDC = GetDesktopWindow()
StartDoc = ShellExecute(Scr_hDC, "Open", TheCommand, _
    Parameters, "C:\", SW_SHOWNORMAL)
If Dir(Parameters, 39) <> "" Then Parameters = PrettyPath(LongName(Parameters))
Close
Kill Path & "\" & MessageFile
Open ReturnPath & "\" & LoginName For Output As FileNumber
Print #FileNumber, TheCommand & " " & Parameters & " : " & ErrMessage(StartDoc)
Close
Exit Sub
End If

Close
Kill Path & "\" & MessageFile

For i = 1 To MNum
If UCase(Messages(i).Caption) = UCase("Instant Message - " & MessageFile) Then
Messages(i).Conversation = Messages(i).Conversation & _
   MessageFile & " :>  " & TheMessage & vbCrLf
Messages(i).SentDate = "Message date: " & DateSent
Messages(i).Show
Messages(i).Sending.SetFocus
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
If Messages(MNum).Top < 0 Then
Messages(MNum).Top = Messages(MNum - 12).Top
Messages(MNum).Left = Messages(MNum - 12).Left
End If
Messages(MNum).Conversation = MessageFile & " :>  " & TheMessage & vbCrLf
Messages(MNum).ToWho = MessageFile
Messages(MNum).From = LoginName
Messages(MNum).Caption = "Instant Message - " & MessageFile
Messages(MNum).SentDate = "Message date: " & DateSent
Messages(MNum).Show
Messages(MNum).Sending.SetFocus
TopZ Messages(MNum)
NoTopZ Messages(MNum)
End If
Exit Sub
TheEnd:
Close
End Sub

