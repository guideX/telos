VERSION 5.00
Begin VB.MDIForm mdiTelos 
   BackColor       =   &H00000000&
   Caption         =   "Nexgen Telos"
   ClientHeight    =   6360
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   9900
   Icon            =   "mdiTelos.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   9840
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   9900
      Begin VB.FileListBox filLogs 
         Height          =   1065
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   1335
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuNew 
         Caption         =   "New"
         Begin VB.Menu mnuTextFile 
            Caption         =   "Text File"
            Shortcut        =   ^N
         End
         Begin VB.Menu mnuSep3872983692 
            Caption         =   "-"
         End
         Begin VB.Menu mnuServer1 
            Caption         =   "New Server"
            Shortcut        =   {F4}
         End
         Begin VB.Menu mnuServerRange 
            Caption         =   "New Server Range"
            Shortcut        =   ^R
         End
         Begin VB.Menu mnuSep398729863 
            Caption         =   "-"
         End
         Begin VB.Menu mnuClientWindow 
            Caption         =   "New Client Window"
         End
         Begin VB.Menu mnuConnectionToServer 
            Caption         =   "New Telnet Connection"
         End
         Begin VB.Menu mnuSep837298362 
            Caption         =   "-"
         End
         Begin VB.Menu mnuTELNETEXE 
            Caption         =   "New Telnet Window"
         End
      End
      Begin VB.Menu mnuSep386283682753 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLogFiles 
         Caption         =   "Open"
         Begin VB.Menu mnuOpenTextFile 
            Caption         =   "Text File"
            Shortcut        =   ^O
         End
         Begin VB.Menu mnuSep3872896392 
            Caption         =   "-"
         End
         Begin VB.Menu mnuLogFile 
            Caption         =   ""
            Index           =   0
         End
      End
      Begin VB.Menu mnuFoldersAccessed 
         Caption         =   "Accessed Folders"
         Visible         =   0   'False
         Begin VB.Menu mnuFolderName 
            Caption         =   ""
            Index           =   0
         End
      End
      Begin VB.Menu mnuAddUser 
         Caption         =   "Add User"
         Shortcut        =   ^U
      End
      Begin VB.Menu mnuRemoveUser 
         Caption         =   "Remove User"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSep8369276392 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save &As ..."
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuSaveSelection 
         Caption         =   "Save Selection"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSep123456789 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "Print"
         Shortcut        =   ^P
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSep12345 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
      Begin VB.Menu txtEditor 
         Caption         =   "Text Editor"
      End
      Begin VB.Menu mnuUsersTEXT 
         Caption         =   "Userlist"
      End
      Begin VB.Menu mnuSep836287563892 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClear 
         Caption         =   "Clear"
      End
      Begin VB.Menu mnuSep836798263 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBackcolor 
         Caption         =   "Change Color"
         Shortcut        =   ^B
      End
   End
   Begin VB.Menu mnuServer 
      Caption         =   "Server"
      Begin VB.Menu mnuLogin 
         Caption         =   "Login"
         Shortcut        =   {F8}
      End
      Begin VB.Menu mnuSep8378926392 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSettings 
         Caption         =   "Settings"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuStatus 
         Caption         =   "Status"
         Begin VB.Menu mnuServerActive 
            Caption         =   "Connect"
            Shortcut        =   {F5}
         End
         Begin VB.Menu mnuServerNotActive 
            Caption         =   "Disconnect"
            Shortcut        =   {F6}
         End
      End
      Begin VB.Menu mnuFiles 
         Caption         =   "Events"
         Begin VB.Menu mnuConnect 
            Caption         =   "On connect"
         End
         Begin VB.Menu mnuMotdTEXT 
            Caption         =   "On motd"
         End
         Begin VB.Menu mnuWelcomeText 
            Caption         =   "When logged in"
         End
      End
   End
   Begin VB.Menu mnuWindows 
      Caption         =   "Windows"
      Begin VB.Menu mnuTileH 
         Caption         =   "Tile (Horizontel)"
      End
      Begin VB.Menu mnuTileVerticle 
         Caption         =   "Tile (Verticle)"
      End
      Begin VB.Menu mnuCascade 
         Caption         =   "Cascade"
      End
      Begin VB.Menu mnu8758585785 
         Caption         =   "-"
      End
      Begin VB.Menu mnuArrangeIcons 
         Caption         =   "Arrange Icons"
      End
      Begin VB.Menu mnuSep32897389 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRestoreAll 
         Caption         =   "Restore All"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuMinimize 
         Caption         =   "Minimize"
      End
      Begin VB.Menu mnuMinimizeAll 
         Caption         =   "Minimize  All"
      End
      Begin VB.Menu mnuSep893798263 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCloseWindow 
         Caption         =   "Close Window"
         Shortcut        =   ^W
      End
      Begin VB.Menu mnuCloseAllWindows 
         Caption         =   "Close All Windows"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuRegister 
         Caption         =   "Register"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About Telos"
      End
   End
   Begin VB.Menu mnuNicklist 
      Caption         =   "Nicklist"
      Visible         =   0   'False
      Begin VB.Menu mnuRemoveFromServer 
         Caption         =   "Remove (Kick)"
      End
      Begin VB.Menu mnuSendMessage 
         Caption         =   "Send Message"
      End
      Begin VB.Menu mnuInformation 
         Caption         =   "Information"
      End
   End
End
Attribute VB_Name = "mdiTelos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub ActiveateSave()
On Local Error Resume Next
mnuSave_Click
End Sub

Public Sub ActiveateSaveAs()
On Local Error Resume Next
mnuSaveAs_Click
End Sub

Private Sub MDIForm_Load()
On Local Error Resume Next
Dim f As Form, i As Integer
lSettings.sIniFile = App.Path & "\telos.ini"
Caption = "Nexgen Telos v" & App.Major & "." & App.Revision
ResetLogs
frmSplash.Show
lSettings.sAcceptingConnections = ReadINI(lSettings.sIniFile, "Settings", "AcceptingConnections", True)
lSettings.sAcceptingNewUsers = ReadINI(lSettings.sIniFile, "Settings", "AcceptingNewUsers", True)
lSettings.sAdministrator = ReadINI(lSettings.sIniFile, "Settings", "Administrator", "Administrator")
lSettings.sClosePause = ReadINI(lSettings.sIniFile, "Settings", "ClosePause", False)
lSettings.sColoredWindows = ReadINI(lSettings.sIniFile, "Settings", "ColoredWindows", False)
lSettings.sConnectedOnStartup = ReadINI(lSettings.sIniFile, "Settings", "ConnectedOnStartup", True)
lSettings.sEchoText = ReadINI(lSettings.sIniFile, "Settings", "EchoText", False)
lSettings.sEMail = ReadINI(lSettings.sIniFile, "Settings", "EMail", "not@specified.nothing")
lSettings.sFolderChat = ReadINI(lSettings.sIniFile, "Settings", "FolderChat", True)
lSettings.sFSTelnet = ReadINI(lSettings.sIniFile, "Settings", "FSTelnet", False)
lSettings.sKillDelay = ReadINI(lSettings.sIniFile, "Settings", "KillDelay", 1)
lSettings.sLoadServerOnStartup = ReadINI(lSettings.sIniFile, "Settings", "LoadServerOnStartup", True)
lSettings.sMaxPasswordRetries = ReadINI(lSettings.sIniFile, "Settings", "MaxPasswordRetries", 3)
lSettings.sMonitorChat = ReadINI(lSettings.sIniFile, "Settings", "MonitorChat", True)
lSettings.sName = ReadINI(lSettings.sIniFile, "Settings", "Name", "")
lSettings.sPassword = ReadINI(lSettings.sIniFile, "Settings", "Password", "")
lSettings.sRootDir = ReadINI(lSettings.sIniFile, "Settings", "RootDir", App.Path & "\fs")
lSettings.sRootDirLen = Len(lSettings.sRootDir)
lSettings.sServerName = ReadINI(lSettings.sIniFile, "Settings", "ServerName", "Telos")
lSettings.sShareOnlyHomeDir = ReadINI(lSettings.sIniFile, "Settings", "ShareOnlyHomeDir", True)
lSettings.sShowUsersInDir = ReadINI(lSettings.sIniFile, "Settings", "ShowUsersInDir", False)
lSettings.sTelnetOnStartup = ReadINI(lSettings.sIniFile, "Settings", "TelnetOnStartup", False)
lSettings.sWhiteBlack = ReadINI(lSettings.sIniFile, "Settings", "WhiteBlack", False)
lStrings.sAccessDenied = ReadINI(lSettings.sIniFile, "Strings", "AccessDenied", "Access Denied")
lStrings.sBadPassword = ReadINI(lSettings.sIniFile, "Strings", "BadPassword", "That is not the password. Please, go away")
lStrings.sLoggedOut = ReadINI(lSettings.sIniFile, "Strings", "LoggedOut", "You have been logged out by the system administrator")
lStrings.sNick = ReadINI(lSettings.sIniFile, "Strings", "Nick", " is now known as ")
lStrings.sDirectoryError = ReadINI(lSettings.sIniFile, "Strings", "DirectoryError", "Directory does not exist")
lStrings.sEnterUsername = ReadINI(lSettings.sIniFile, "Strings", "EnterUsername", "Username: ")
lStrings.sEnterPassword = ReadINI(lSettings.sIniFile, "Strings", "EnterPassword", "Password: ")
lStrings.sShuttingDown = ReadINI(lSettings.sIniFile, "Strings", "ShuttingDown", " will be shutting down in ")
lStrings.sShuttingDownNow = ReadINI(lSettings.sIniFile, "Strings", "ShuttingDownNow", " will be shutting down now")
If Len(lSettings.sName) <> 0 And Len(lSettings.sPassword) <> 0 And KeyGen(lSettings.sName, "pickles", 1) = lSettings.sPassword Then
    lSettings.sRegistered = True
    lSettings.sMaxUsers = 1000
Else
    lSettings.sRegistered = False
    lSettings.sMaxUsers = 35
End If
If lSettings.sLoadServerOnStartup = True Then
    Set f = New frmTelos
    f.Show
End If
End Sub

Private Sub MDIForm_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then PopupMenu mnuServer
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
On Local Error Resume Next
lSettings.sGlobalClose = True
If lSettings.sFSTelnet = True Then
    Unload frmFSTelnet
End If
End
End Sub

Private Sub mnuAbout_Click()
On Local Error Resume Next
Dim i As Form
Set i = New frmTxtEditor
i.Show
i.txtIncoming.SelText = "Telos" & vbCrLf & "By Team Nexgen" & vbCrLf & "Telos is a command line operating system that can be accessed online from any computer (mac, pc, more) capable of connecting to a telnet server. Telos was developed by Leon J Aiossa for the purpose of learning the command-line linux operating system for an operating system class he attended at St. Paul College without actually running linux."
i.txtIncoming.SelText = i.txtIncoming.SelText & vbCrLf & vbCrLf & "Programmer: Leon J Aiossa" & vbCrLf & "Also known as |guideX|"
i.txtIncoming.Locked = True
i.Tag = "About"
i.Caption = "About"
i.chkModifiedSinceLastSave.Value = 0
End Sub

Private Sub mnuAddUser_Click()
On Local Error Resume Next
frmNewUser.Show 1
End Sub

Private Sub mnuArrangeIcons_Click()
On Local Error Resume Next
mdiTelos.Arrange vbArrangeIcons
End Sub

Private Sub mnuBackcolor_Click()
On Local Error Resume Next
Dim msg As String
msg = GetRnd(14700000)
If ActiveForm.Tag = "Telnet" Then
    ActiveForm.txtIncoming.BackColor = msg
    ActiveForm.txtOutgoing.BackColor = msg
ElseIf ActiveForm.Name = "frmTxtEditor" Then
    ActiveForm.txtIncoming.BackColor = msg
Else
    ActiveForm.txtOutgoing.BackColor = msg
    ActiveForm.console.BackColor = msg
    ActiveForm.lstUsers.BackColor = msg
End If
If Err.Number <> 0 Then
    Err.Clear
    mdiTelos.BackColor = msg
End If
End Sub

Private Sub mnuCascade_Click()
On Local Error Resume Next
mdiTelos.Arrange vbCascade
End Sub

Private Sub mnuClear_Click()
On Local Error Resume Next
If ActiveForm.Tag = "Telnet" Then
    ActiveForm.txtIncoming.text = ""
Else
    ActiveForm.console.Clear
End If
End Sub

Private Sub mnuClientWindow_Click()
On Local Error Resume Next
LoadClient
End Sub

Private Sub mnuCloseAllWindows_Click()
On Error GoTo ErrDump
Dim i As Integer
For i = 1 To mdiTelos.Count
    If Len(ActiveForm.Tag) <> 0 Then
        Unload ActiveForm
    End If
Next i
ErrDump:
    Err.Number = 0
End Sub

Private Sub mnuCloseWindow_Click()
On Local Error Resume Next
Unload ActiveForm
End Sub

Private Sub mnuConnect_Click()
On Local Error Resume Next
OpenTextFile App.Path & "\files\connect.txt"
End Sub

Private Sub mnuConnectionToServer_Click()
On Local Error Resume Next
Dim msg As String
lSettings.sClientConnectionIp = InputBox("Enter IP Address:", "", lSettings.sIpAddress)
If Len(lSettings.sClientConnectionIp) <> 0 Then
    frmTelnet.Show
End If
End Sub

Private Sub mnuExit_Click()
On Local Error Resume Next
Unload Me
End Sub

Private Sub mnuFolderName_Click(Index As Integer)
On Local Error Resume Next
Dim msg As String
msg = mnuFolderName(Index).Caption
Shell "EXPLORER.EXE " & msg, vbNormalFocus
End Sub

Private Sub mnuInformation_Click()
On Local Error Resume Next
With frmUserInfo
    .txtHost.text = lUsers.uUser(ActiveForm.lstUsers.ListIndex + 1).uHost
    .txtInput.text = lUsers.uUser(ActiveForm.lstUsers.ListIndex + 1).uInput
    .txtNickname.text = lUsers.uUser(ActiveForm.lstUsers.ListIndex + 1).uName
    .txtRelDIR.text = lUsers.uUser(ActiveForm.lstUsers.ListIndex + 1).uRelDir
    .txtSock.text = lUsers.uUser(ActiveForm.lstUsers.ListIndex + 1).uSock
    .txtSuperUser.text = lUsers.uUser(ActiveForm.lstUsers.ListIndex + 1).uOperator
    .txtWhat.text = lUsers.uUser(ActiveForm.lstUsers.ListIndex + 1).uWhat
    .txtWTF.text = lUsers.uUser(ActiveForm.lstUsers.ListIndex + 1).uWTF
    .Show
End With
End Sub

Private Sub mnuLogFile_Click(Index As Integer)
On Local Error Resume Next
OpenTextFile App.Path & "\logs\" & mnuLogFile(Index).Caption
End Sub

Private Sub mnuLogin_Click()
On Local Error Resume Next
LoadClient
End Sub

Private Sub mnuMinimize_Click()
On Local Error Resume Next
mdiTelos.ActiveForm.WindowState = vbMinimized
End Sub

Private Sub mnuMinimizeAll_Click()
On Local Error Resume Next
Dim i As Integer
For i = 1 To mdiTelos.Count
    ActiveForm.WindowState = vbMinimized
Next i
End Sub

Private Sub mnuMotdTEXT_Click()
On Local Error Resume Next
OpenTextFile App.Path & "\files\motd.txt"
End Sub

Private Sub mnuOpenTextFile_Click()
On Local Error Resume Next
Dim msg As String, msg2 As String, f As Form
msg = OpenDialog(Me, "Text Files (*.txt)|*.txt|All Files (*.*)|*.*|", "Open Text File", CurDir)
If Len(msg) <> 0 Then
    msg2 = ReadFile(msg)
    If Len(msg2) <> 0 Then
        Set f = New frmTxtEditor
        f.Show
        f.Tag = "Text Editor"
        f.Caption = msg
        f.txtIncoming.SelText = msg2
        f.chkModifiedSinceLastSave.Value = 0
    End If
End If
End Sub

Private Sub mnuRegister_Click()
On Local Error Resume Next
frmRegister.Show 1
End Sub

Private Sub mnuRemoveFromServer_Click()
On Local Error Resume Next
LogoutUser ActiveForm.lstUsers.ListIndex + 1, ActiveForm, InputBox("Reason:", "", "Systemic Abuse")
End Sub

Private Sub mnuRestoreAll_Click()
On Local Error Resume Next
Dim i As Integer
For i = 1 To mdiTelos.Count
    ActiveForm.WindowState = vbNormal
Next i
End Sub

Private Sub mnuSave_Click()
On Local Error Resume Next
Dim msg As String, msg2 As String, i As Integer
If ActiveForm.Tag = "Telnet" Then
    msg = ActiveForm.txtIncoming.text
    msg2 = App.Path & "\logs\" & Right(Format(Time, "##.##.##.##"), Len(Format(Time, "##.##.##.##")) - 1) & ".log"
    SaveFile msg2, msg
    ActiveForm.AddText vbCrLf & "Saved as " & msg2 & vbCrLf, vbBlue
ElseIf ActiveForm.Tag = "Text Editor" Then
    If DoesFileExist(ActiveForm.Caption) = True Then
        msg = ActiveForm.txtIncoming.text
        msg2 = ActiveForm.Caption
        SaveFile msg2, msg
    Else
        ActiveateSaveAs
    End If
Else
    For i = 0 To ActiveForm.console.ListIndex
        If i = 0 Then
            msg = ActiveForm.console.List(i)
        Else
            msg = msg & vbCrLf & ActiveForm.console.List(i)
        End If
    Next i
    msg2 = App.Path & "\logs\" & Right(Format(Time, "##.##.##.##"), Len(Format(Time, "##.##.##.##")) - 1) & ".log"
    SaveFile msg2, msg
    ConsolePrint "Saved as " & msg2, ActiveForm
    DoEvents
    ResetLogs
End If
End Sub

Private Sub mnuSaveAs_Click()
On Local Error Resume Next
Dim msg As String, msg2 As String, i As Integer, msg3 As String
If ActiveForm.Tag = "Telnet" Then
    msg = ActiveForm.txtIncoming.text
    msg2 = SaveDialog(mdiTelos, "Log Files (*.log)|*.log|", "Save as ...", App.Path & "\logs\")
    msg2 = Left(msg2, Len(msg2) - 1) & ".log"
    SaveFile msg2, msg
    ActiveForm.AddText "Saved as " & msg2, vbBlue
ElseIf ActiveForm.Tag = "Text Editor" Then
    msg = ActiveForm.txtIncoming.text
    msg3 = ActiveForm.Caption
    msg3 = GetFileTitle(msg3)
    msg2 = SaveDialog(mdiTelos, "Log Files (*.txt)|*.txt|", "Save as ...", Left(ActiveForm.Caption, Len(ActiveForm.Caption) - Len(msg3)))
    msg2 = Left(msg2, Len(msg2) - 2) & ".txt"
    ActiveForm.chkModifiedSinceLastSave.Value = 0
    SaveFile msg2, msg
Else
    For i = 0 To ActiveForm.console.ListIndex
        If i = 0 Then
            msg = ActiveForm.console.List(i)
        Else
            msg = msg & vbCrLf & ActiveForm.console.List(i)
        End If
    Next i
    msg2 = SaveDialog(mdiTelos, "Log Files (*.log)|*.log|", "Save as ...", App.Path & "\logs\")
    If Len(msg2) <> 0 Then
        msg2 = Left(msg2, Len(msg2) - 1) & ".log"
        SaveFile msg2, msg
        ConsolePrint "Saved as " & msg2, ActiveForm
    End If
    DoEvents
    ResetLogs
End If
End Sub

Private Sub mnuSendMessage_Click()
On Local Error Resume Next
ActiveForm.Send vbCrLf & "<" & lSettings.sAdministrator & "> " & InputBox("Enter message:", "", "") & vbCrLf, ActiveForm.lstUsers.ListIndex + 1
End Sub

Private Sub mnuServer1_Click()
On Local Error Resume Next
Dim i As Form, msg As String
Set i = New frmTelos
i.Show
If lSettings.sColoredWindows = True Then
    msg = GetRnd(14700000)
    i.lstUsers.BackColor = msg
    i.console.BackColor = msg
    i.txtOutgoing.BackColor = msg
End If
End Sub

Private Sub mnuServerActive_Click()
On Local Error Resume Next
ActivateServerToggle True, ActiveForm
End Sub

Private Sub mnuServerNotActive_Click()
On Local Error Resume Next
ActivateServerToggle False, ActiveForm
End Sub

Private Sub mnuServerRange_Click()
On Local Error Resume Next
frmNewServerRange.Show
End Sub

Private Sub mnuSettings_Click()
On Local Error Resume Next
frmSettings.Show 1
End Sub

Private Sub mnuTELNETEXE_Click()
On Local Error Resume Next
Shell "telnet " & lSettings.sIpAddress, vbNormalFocus
End Sub

Private Sub mnuTextFile_Click()
On Local Error Resume Next
Dim i As Form
Set i = New frmTxtEditor
i.Show
End Sub

Private Sub mnuTileH_Click()
On Local Error Resume Next
mdiTelos.Arrange vbTileHorizontal
End Sub

Private Sub mnuTileVerticle_Click()
On Local Error Resume Next
mdiTelos.Arrange vbTileVertical
End Sub

Private Sub mnuUsersTEXT_Click()
On Local Error Resume Next
Dim f As Form
Set f = New frmTxtEditor
f.Show
f.txtIncoming.text = ReadFile(App.Path & "\files\users.txt")
f.Caption = App.Path & "\files\users.txt"
f.Tag = "Text Editor"
f.chkModifiedSinceLastSave.Value = 0
End Sub

Private Sub mnuWelcomeTEXT_Click()
On Local Error Resume Next
OpenTextFile App.Path & "\files\welcome.txt"
End Sub

Private Sub txtEditor_Click()
frmTxtEditor.Show
End Sub
