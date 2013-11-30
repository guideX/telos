Attribute VB_Name = "mdlMain"
Option Explicit
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Const HWND_NOTOPMOST = -2
Public Const HWND_TOPMOST = -1
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_SHOWWINDOW = &H40
Public waitft As Boolean
Public iniclass As New INI
Public Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long

Public Sub AlwaysOnTop(myfrm As Form, SetOnTop As Boolean)
On Local Error Resume Next
Dim lFlag As Integer
If SetOnTop Then
    lFlag = HWND_TOPMOST
Else
    lFlag = HWND_NOTOPMOST
End If
SetWindowPos myfrm.hwnd, lFlag, myfrm.Left / Screen.TwipsPerPixelX, myfrm.Top / Screen.TwipsPerPixelY, myfrm.Width / Screen.TwipsPerPixelX, myfrm.Height / Screen.TwipsPerPixelY, SWP_NOACTIVATE Or SWP_SHOWWINDOW
End Sub

Public Sub Surf(lUrl As String, lHwnd As Long)
On Local Error Resume Next
Dim msg As Long
msg = ShellExecute(lHwnd, vbNullString, lUrl, vbNullString, "c:\", SW_SHOWNORMAL)
If Err.Number <> 0 Then ErrorHandle Err.Number, Err.Description, "Public Sub Surf(lUrl As String, lHwnd As Long)"
End Sub

Sub Main()
On Local Error Resume Next
CurDir = App.Path & "\fs"
If Err.Number <> 0 Then ErrorHandle Err.Number, Err.Description, "Sub Main()"
End Sub

Public Function DoesFileExist(lFilename As String) As Boolean
On Local Error Resume Next
Dim msg As String, i As Integer, dr As String
If Len(lFilename) <> 0 Then
    msg = Dir(lFilename)
    If msg <> "" Then
        DoesFileExist = True
    Else
        DoesFileExist = False
    End If
End If
If Err.Number <> 0 Then ErrorHandle Err.Number, Err.Description, "Public Function DoesFileExist(lFilename As String) As Boolean"
End Function

Public Function ReadFile(lFile As String) As String
On Local Error Resume Next
Dim o As Integer, msg As String
o = FreeFile
If DoesFileExist(lFile) = True Then
    Open lFile For Input As #o
        ReadFile = StrConv(InputB(LOF(o), o), vbUnicode)
    Close #o
End If
If Err.Number <> 0 Then ErrorHandle Err.Number, Err.Description, "Public Function ReadFile(lFile As String) As String"
End Function

Public Function GetRnd(Num As Long) As Long
Randomize Timer
GetRnd = Int((Num * Rnd) + 1)
If Err.Number <> 0 Then ErrorHandle Err.Number, Err.Description, "Public Function GetRnd(Num As Long) As Long"
End Function

Public Sub LogoutUser(lUserIndex As Integer, lForm As Form, Optional lReason As String)
On Local Error Resume Next
Dim i As Integer
With lUsers.uUser(lUserIndex)
    For i = 0 To lSettings.sMaxUsers
        If Len(lUsers.uUser(i).uName) <> 0 And Len(lUsers.uUser(i).uHost) <> 0 Then
            If Len(lReason) <> 0 Then
                lForm.Send vbCrLf & lUsers.uUser(lUserIndex).uName & " was logged out by " & lSettings.sAdministrator & " (" & lReason & ")", i
            End If
        End If
    Next i
    lForm.Send vbCrLf & vbCrLf & lStrings.sLoggedOut, lUserIndex
    Sleep 2, True
    lForm.pol(.uSock).Close
    ConsolePrint .uName & " disconnected", lForm
    .uName = Empty
    .uInput = Empty
    .uHost = Empty
    .uWhat = Empty
    .uSock = Empty
    .uOperator = False
    .uWTF = False
    .uRelDir = Empty
    lForm.Update
End With
If Err.Number <> 0 Then ErrorHandle Err.Number, Err.Description, "Public Sub LogoutUser(lUserIndex As Integer, lForm As Form, Optional lReason As String)"
End Sub

Public Function SaveFile(lFilename As String, lText As String) As Boolean
On Local Error Resume Next
If Len(lFilename) <> 0 And Len(lText) <> 0 Then
    Open lFilename For Output As #1
    Print #1, lText
    Close #1
End If
If Err.Number <> 0 Then ErrorHandle Err.Number, Err.Description, "Public Function SaveFile(lFilename As String, lText As String) As Boolean"
End Function

Public Sub LoadClient()
On Local Error Resume Next
Dim i As Form, msg As String
Set i = New frmTelnet
If lSettings.sFSTelnet = True Then
    mdiTelos.WindowState = vbMinimized
    frmFSTelnet.Show
    i.Show
    frmFSTelnet.txtOutgoing.SetFocus
Else
    i.Show
End If
If lSettings.sColoredWindows = True Then
    msg = GetRnd(14700000)
    i.txtOutgoing.BackColor = msg
    i.txtIncoming.BackColor = msg
End If
If Err.Number <> 0 Then ErrorHandle Err.Number, Err.Description, "Public Sub LoadClient()"
End Sub

Public Sub ErrorHandle(lErrorNumber As Integer, lErrorDescription As String, lSub As String)
On Local Error Resume Next
Err = 0
End Sub

Public Sub ResetLogs()
On Local Error Resume Next
Dim i As Integer
mdiTelos.filLogs.Path = App.Path & "\logs\"
mdiTelos.filLogs.Refresh
For i = 0 To mdiTelos.mnuLogFile.Count - 1
    If i <> 0 Then
        Unload mdiTelos.mnuLogFile(i)
        mdiTelos.mnuLogFile(i).Caption = ""
    End If
Next i
For i = 0 To mdiTelos.filLogs.ListCount
    If Len(mdiTelos.filLogs.List(i)) <> 0 Then
        If i <> 0 Then Load mdiTelos.mnuLogFile(i)
        If Err.Number <> 0 Then Err = 0
        mdiTelos.mnuLogFile(i).Caption = mdiTelos.filLogs.List(i)
        mdiTelos.mnuLogFile(i).Visible = True
    End If
Next i
If Err.Number <> 0 Then ErrorHandle Err.Number, Err.Description, "Public Sub ResetLogs()"
End Sub

Public Function CommandParse(lCommand As String, lArguement As String, IsInternet As Boolean, lForm As Form, Optional lUserID, Optional SU As Boolean = False) As Boolean
On Local Error GoTo ErrorSpot
Dim m As Integer, k As Menu, msg As String, bob As Integer, l As Integer, temp As String, i As Integer
If (lCommand = "logout" And IsInternet = True) Or (lCommand = "exit" And IsInternet = True) Then
    lForm.Send vbCrLf, lUserID
    With lUsers.uUser(lUserID)
        lForm.pol(.uSock).Close
        ConsolePrint .uName & " disconnected.", lForm
        For m = 1 To lForm.pol.Count
            lForm.Send vbCrLf & .uName & " disconnected" & vbCrLf, m
        Next m
        .uName = Empty
        .uInput = Empty
        .uHost = Empty
        .uWhat = Empty
        .uSock = Empty
        .uOperator = False
        .uWTF = False
        .uRelDir = Empty
        lForm.Update
        CommandParse = True
        Exit Function
    End With
ElseIf lCommand = "shutdown" Then
    lForm.Send vbCrLf, lUserID
    With lUsers.uUser(lUserID)
        If .uOperator = True Or SU = True Then
            BroadcastKillMessage Val(lArguement), lForm
            CommandParse = True
            Exit Function
        Else
            lForm.Send lStrings.sAccessDenied & vbCrLf, lUserID
            CommandParse = False
            Exit Function
        End If
    End With
ElseIf lCommand = "nick" Then
    lForm.Send vbCrLf, lUserID
    With lUsers.uUser(lUserID)
        For i = 1 To lSettings.sMaxUsers
            If Len(lUsers.uUser(i).uName) <> 0 Then lForm.Send .uName & lStrings.sNick & lArguement & vbCrLf, i
        Next i
        .uName = lArguement
    End With
ElseIf lCommand = "list" Then
    lForm.Send vbCrLf, lUserID
    With lUsers.uUser(lUserID)
        lForm.Send vbCrLf & "Who's online at " & lSettings.sServerName, lUserID
        lForm.Send vbCrLf & "Node #0" & ": [@] " & lForm.acc.LocalIP & ": " & lSettings.sAdministrator & " (" & lSettings.sEMail & ")", lUserID
        For bob = 1 To lSettings.sMaxUsers
            If Len(lForm.pol(bob).RemoteHostIP) <> 0 And Len(lUsers.uUser(bob).uName) <> 0 Then
                If .uOperator = True Then
                    lForm.Send vbCrLf & "Node #" & bob & ": [@] " & lForm.pol(bob).RemoteHostIP & ": " & lUsers.uUser(bob).uName, lUserID
                Else
                    lForm.Send vbCrLf & "Node #" & bob & ": [-] " & lForm.pol(bob).RemoteHostIP & ": " & lUsers.uUser(bob).uName, lUserID
                End If
            End If
        Next bob
    End With
ElseIf lCommand = "motd" Then
    lForm.Send vbCrLf, lUserID
    lForm.Send vbCrLf & "Sending (motd.txt) ..." & vbCrLf, lUserID
    lForm.SendFile App.Path & "\files\motd.txt", lUserID
ElseIf lCommand = "createuser" Then
    lForm.Send vbCrLf, lUserID
    With lUsers.uUser(lUserID)
        If .uOperator = True Or SU = True Then
            lForm.Send vbCrLf & lStrings.sEnterUsername, lUserID
            temp = Trim(lForm.WaitForText(lUserID))
            temp = temp & ","
            lForm.Send vbCrLf & lStrings.sEnterPassword
            temp = temp & Trim(lForm.WaitForText(lUserID))
            temp = temp & ","
            lForm.Send vbCrLf & "Administrator? [Y-N]:", lUserID
            If lForm.WaitForText(lUserID) = "y" Then
                temp = temp & "1,"
            Else
                temp = temp & "0,"
            End If
            Open App.Path & "\files\users.txt" For Append As #1
            temp = temp
            Print #1, temp
            Close #1
            lForm.Send vbCrLf & "User created!" & vbCrLf, lUserID
            CommandParse = True
            Exit Function
        Else
            lForm.Send vbCrLf & lStrings.sAccessDenied, lUserID
            CommandParse = False
            Exit Function
        End If
    End With
ElseIf lCommand = "ver" Then
    lForm.Send vbCrLf, lUserID
    lForm.Send vbCrLf & "Nexgen Telos Version " & App.Major & "." & App.Minor & App.Revision & " - " & App.ThreadID & vbCrLf & "Telnet Operating System", lUserID
'ElseIf lCommand = "dir/w" Then
'WIDE:
    'lForm.Send vbCrLf, lUserID
    'For bob = 0 To lForm.folders(lUserID).ListCount - 1
    '    If bob = 0 Then
    '        lForm.Send "[" & JustFile(lForm.folders(lUserID).List(bob)) & "]     ", lUserID
    '    Else
    '        lForm.Send "[" & JustFile(lForm.folders(lUserID).List(bob)) & "]     ", lUserID
    '    End If
    '    l = l + 1
    'Next bob
    'lForm.Send vbCrLf & vbCrLf, lUserID
    'For bob = 0 To lForm.files(lUserID).ListCount - 1
    '    If lForm.files(lUserID).List(bob) <> "ownership.ini" Then
    '        lForm.Send lForm.files(lUserID).List(bob) & "     ", lUserID
    '        m = m + 1
    '    End If
    'Next bob
    'If m <> 0 Then
    '    msg = m & " file(s)"
    '    If l <> 0 Then msg = vbCrLf & msg & ", " & l & " folder(s)"
    'Else
    '    If l <> 0 Then msg = vbCrLf & l & " folder(s)"
    'End If
    'lForm.Send msg & vbCrLf, lUserID
    'lForm.Send l & " file(s)"
    'CommandParse = True
    'Exit Function
ElseIf lCommand = "ls" Or lCommand = "dir" Then
    'If LCase(Trim(lArguement)) = "/w" Then GoTo WIDE
    lForm.Send vbCrLf, lUserID
    For bob = 0 To lForm.folders(lUserID).ListCount - 1
        If bob = 0 Then
            lForm.Send vbCrLf & "[dr][" & Format(GetAttr(lForm.folders(lUserID).List(bob)), "000") & "] " & JustFile(lForm.folders(lUserID).List(bob)) & vbCrLf, lUserID
        Else
            lForm.Send "[dr][" & Format(GetAttr(lForm.folders(lUserID).List(bob)), "000") & "] " & JustFile(lForm.folders(lUserID).List(bob)) & vbCrLf, lUserID
        End If
        l = l + 1
    Next bob
    For bob = 0 To lForm.files(lUserID).ListCount - 1
        If lForm.files(lUserID).List(bob) <> "ownership.ini" Then
            lForm.Send "[fl][" & Format(GetAttr(lForm.folders(lUserID).Path & "\" & lForm.files(lUserID).List(bob)), "000") & "] " & lForm.files(lUserID).List(bob) & " [" & Format(FileSystem.FileLen(lForm.folders(lUserID).Path & "\" & lForm.files(lUserID).List(bob)), "###,###,###") & "]" & vbCrLf, lUserID
            m = m + 1
        End If
    Next bob
    If lSettings.sShowUsersInDir = True Then
        For bob = 0 To lSettings.sMaxUsers
            If Len(lUsers.uUser(bob).uName) <> 0 Then
                If LCase(lUsers.uUser(bob).uRelDir = LCase(lUsers.uUser(lUserID).uRelDir)) Then
                    If bob = 0 Then
                        i = i + 1
                        lForm.Send vbCrLf & "[ur][" & Format(i, "000") & "] " & lUsers.uUser(i).uName & vbCrLf, lUserID
                    Else
                        i = i + 1
                        lForm.Send "[ur][" & Format(i, "000") & "] " & lUsers.uUser(i).uName & vbCrLf, lUserID
                    End If
                End If
            End If
        Next bob
    End If
    If m <> 0 Then
        If m = 1 Then
            msg = "1 file"
        Else
            msg = m & " files"
        End If
    End If
    If l <> 0 Then
        If l = 1 Then
            msg = msg & " 1 folder"
        Else
            msg = msg & " " & l & " folders"
        End If
    End If
    If lSettings.sShowUsersInDir = True Then
        If i <> 0 Then
            If i = 1 Then
                msg = msg & " 1 user"
            Else
                msg = msg & " " & i & " users"
            End If
        End If
    End If
    lForm.Send Trim(msg) & vbCrLf, lUserID
    CommandParse = True
    Exit Function
ElseIf lCommand = "man" Then
    lForm.Send vbCrLf, lUserID
    PrintManPage lArguement, lUserID, lForm
    CommandParse = True
    Exit Function
ElseIf lCommand = "pwd" Then
    lForm.Send vbCrLf, lUserID
    With lUsers.uUser(lUserID)
        lForm.Send vbCrLf & .uRelDir & vbCrLf, lUserID
        CommandParse = True
        Exit Function
    End With
ElseIf lCommand = "cd" Then
    lForm.Send vbCrLf, lUserID
    With lUsers.uUser(lUserID)
        If lSettings.sShareOnlyHomeDir = True And lForm.folders(lUserID).Path = lSettings.sRootDir And lArguement = ".." Then
            lForm.Send vbCrLf & lStrings.sAccessDenied & vbCrLf, lUserID
            Exit Function
        End If
        If Right(lForm.folders(lUserID).Path, 1) = "\" And Len(lForm.folders(lUserID).Path) = 3 Then
            CommandParse = True
            lForm.files(lUserID).Path = lForm.folders(lUserID).Path & lArguement
            lForm.folders(lUserID).Path = lForm.files(lUserID).Path
            .uRelDir = lForm.folders(lUserID).Path
            Exit Function
        End If
        If IfDirectoryExists(lForm.folders(lUserID).Path & "\" & lArguement) = True Then
            CommandParse = True
            If lForm.folders(lUserID).Path = lSettings.sRootDir And lArguement = ".." Then
                lForm.files(lUserID).Path = lForm.folders(lUserID).Path & "\" & lArguement
                lForm.folders(lUserID).Path = lForm.files(lUserID).Path
                .uRelDir = lForm.folders(lUserID).Path
                .uRelDir = Right$(.uRelDir, Len(.uRelDir) - lSettings.sRootDirLen)
            Else
                lForm.files(lUserID).Path = lForm.folders(lUserID).Path & "\" & lArguement
                lForm.folders(lUserID).Path = lForm.files(lUserID).Path
                .uRelDir = lForm.folders(lUserID).Path
                .uRelDir = Right$(.uRelDir, Len(.uRelDir) - lSettings.sRootDirLen)
            End If
            
            lSettings.sFolderCount = lSettings.sFolderCount + 1
            Load mdiTelos.mnuFolderName(lSettings.sFolderCount)
            mdiTelos.mnuFolderName(lSettings.sFolderCount).Caption = lForm.folders(lUserID).Path
            mdiTelos.mnuFolderName(lSettings.sFolderCount).Visible = True
            mdiTelos.mnuFolderName(0).Visible = False
            mdiTelos.mnuFoldersAccessed.Visible = True
            Exit Function
        Else
            lForm.Send lStrings.sDirectoryError, lUserID
        End If
    End With
ElseIf lCommand = "help" Then
    If Len(lArguement) = 0 Then
        lForm.SendFile App.Path & "\files\basichelp.txt", lUserID
    Else
        Select Case LCase(lArguement)
        Case "cd"
            lForm.Send "cd Command" & vbCrLf & "Usage:" & vbCrLf & "cd [folder]:" & vbCrLf & "cd is the command to change the current folder." & vbCrLf, lUserID
        Case "createuser"
            lForm.Send "Createuser Command" & vbCrLf & "Usage:" & vbCrLf & "Createuser" & vbCrLf & "Createuser is the command to create a new user." & vbCrLf & "You will be prompted for a user name, password, and a yes/no answer" & vbCrLf & "for weather or not the new user will have superuser powers.", lUserID
        End Select
    End If
ElseIf lCommand = "get" Or lCommand = "type" Or lCommand = "download" Then
    'iniclass.AppName = lArguement
    'iniclass.DefaultReturn = ""
    'owner = iniclass.GPPS("owner", lForm.folders(lUserID).Path & "\" & "ownership.ini")
    'meinfo = iniclass.GPPS("me", lForm.folders(lUserID).Path & "\" & "ownership.ini")
    'theminfo = iniclass.GPPS("them", lForm.folders(lUserID).Path & "\" & "ownership.ini")
    'If Ac_Name(lUserID) = owner And InStr(1, meinfo, "r") Then
    lForm.SendFile lForm.folders(lUserID).Path & "\" & lArguement, lUserID
    CommandParse = True
    'Exit Function
    'Else
    'End If
    'If InStr(1, theminfo, "r") Or Ac_Name(lUserID) = "root" Then
    'lForm.SendFile lForm.folders(lUserID).Path & "\" & lArguement, lUserID
    'CommandParse = True
    'Exit Function
    'Else
    'End If
    Exit Function
ElseIf lCommand = "home" Or lCommand = "root" Then
    With lUsers.uUser(lUserID)
        If IfDirectoryExists(lSettings.sRootDir) = True Then
            CommandParse = True
            'If lForm.folders(lUserID).Path = lSettings.sRootDir And lArguement = ".." Then Exit Function
            lForm.files(lUserID).Path = lSettings.sRootDir
            lForm.folders(lUserID).Path = lSettings.sRootDir
            .uRelDir = ""
            Exit Function
        End If
    End With
ElseIf lCommand = "delete" Then
    With lUsers.uUser(lUserID)
        Dim owner As String, theminfo As String, meinfo As String
        iniclass.AppName = lArguement
        iniclass.DefaultReturn = ""
        owner = iniclass.GPPS("owner", lForm.folders(lUserID).Path & "\" & "ownership.ini")
        meinfo = iniclass.GPPS("me", lForm.folders(lUserID).Path & "\" & "ownership.ini")
        theminfo = iniclass.GPPS("them", lForm.folders(lUserID).Path & "\" & "ownership.ini")
        If .uName = owner And InStr(1, meinfo, "w") Then
            DeleteFile lForm.folders(lUserID).Path & "\" & lArguement
            iniclass.WPPS "owner", "", lForm.folders(lUserID).Path & "\" & "ownership.ini"
            iniclass.WPPS "me", "", lForm.folders(lUserID).Path & "\" & "ownership.ini"
            iniclass.WPPS "them", "", lForm.folders(lUserID).Path & "\" & "ownership.ini"
            lForm.files(lUserID).Refresh
            CommandParse = True
            Exit Function
        End If
        If InStr(1, theminfo, "w") Or .uName = "root" Then
            DeleteFile lForm.folders(lUserID).Path & "\" & lArguement
            iniclass.WPPS "owner", "", lForm.folders(lUserID).Path & "\" & "ownership.ini"
            iniclass.WPPS "me", "", lForm.folders(lUserID).Path & "\" & "ownership.ini"
            iniclass.WPPS "them", "", lForm.folders(lUserID).Path & "\" & "ownership.ini"
            lForm.files(lUserID).Refresh
            CommandParse = True
            Exit Function
        End If
        CommandParse = False
        Exit Function
    End With
Else
ErrorSpot:
    CommandParse = False
    If Err.Number <> 0 Then
        'ConsolePrint "Error in CommandParse! " & Err.Description
        'lForm.Send "Error Occured - " & Err.Description & vbCrLf, lUserID
        ErrorHandle Err.Number, Err.Description, "Public Function CommandParse(lCommand As String, lArguement As String, IsInternet As Boolean, lForm As Form, Optional lUserID, Optional SU As Boolean = False) As Boolean"
    Else
        With lUsers.uUser(lUserID)
            If Len(lCommand) <> 0 Then
                If lSettings.sMonitorChat = True Then
                    If Len(.uRelDir) <> 0 Then
                        ConsolePrint "[" & .uRelDir & "] <" & .uName & "> " & lCommand & " " & lArguement, lForm
                    Else
                        ConsolePrint "[Root] <" & .uName & "> " & lCommand & " " & lArguement, lForm
                    End If
                Else
                    
                End If
                If lForm.pol.Count <> 0 Then
                    For i = 1 To lForm.pol.Count
                        If Len(.uName) <> 0 Then
                            If lSettings.sFolderChat = True Then
                                If LCase(.uRelDir) = LCase(lUsers.uUser(i).uRelDir) Then
                                    If Len(.uRelDir) <> 0 Then
                                        lForm.Send vbCrLf & "[" & JustFile2(lForm.folders(lUserID).Path) & "] <" & .uName & "> " & lCommand & " " & lArguement, i
                                    Else
                                        lForm.Send vbCrLf & "[Home] <" & .uName & "> " & lCommand & " " & lArguement, i
                                    End If
                                End If
                            Else
                                lForm.Send vbCrLf & "<" & .uName & "> " & lCommand & " " & lArguement & vbCrLf, i
                            End If
                        End If
                    Next i
                End If
            End If
        End With
    End If
End If
If Err.Number <> 0 Then ErrorHandle Err.Number, Err.Description, "Public Function CommandParse(lCommand As String, lArguement As String, IsInternet As Boolean, lForm As Form, Optional lUserID, Optional SU As Boolean = False) As Boolean"
End Function

Public Sub OpenTextFile(lFilename As String)
On Local Error Resume Next
Dim f As Form
Set f = New frmTxtEditor
f.Show
f.txtIncoming.text = ReadFile(lFilename)
f.Caption = lFilename
f.Tag = "Text Editor"
f.chkModifiedSinceLastSave.Value = 0
End Sub

Public Sub SaveCurrentSettings()
On Local Error Resume Next
WriteINI lSettings.sIniFile, "Strings", "DirectoryError", lStrings.sDirectoryError
WriteINI lSettings.sIniFile, "Strings", "AccessDenied", lStrings.sAccessDenied
WriteINI lSettings.sIniFile, "Strings", "BadPassword", lStrings.sBadPassword
WriteINI lSettings.sIniFile, "Strings", "LoggedOut", lStrings.sLoggedOut
WriteINI lSettings.sIniFile, "Strings", "Nick", lStrings.sNick
WriteINI lSettings.sIniFile, "Strings", "EnterUsername", lStrings.sEnterUsername
WriteINI lSettings.sIniFile, "Strings", "EnterPassword", lStrings.sEnterPassword
WriteINI lSettings.sIniFile, "Strings", "ShuttingDown", lStrings.sShuttingDown
WriteINI lSettings.sIniFile, "Strings", "ShuttingDownNow", lStrings.sShuttingDownNow
WriteINI lSettings.sIniFile, "Settings", "ShowUsersInDir", lSettings.sShowUsersInDir
WriteINI lSettings.sIniFile, "Settings", "ClosePause", lSettings.sClosePause
WriteINI lSettings.sIniFile, "Settings", "MaxPasswordRetries", lSettings.sMaxPasswordRetries
WriteINI lSettings.sIniFile, "Settings", "FolderChat", lSettings.sFolderChat
WriteINI lSettings.sIniFile, "Settings", "EMail", lSettings.sEMail
WriteINI lSettings.sIniFile, "Settings", "ColoredWindows", lSettings.sColoredWindows
WriteINI lSettings.sIniFile, "Settings", "MonitorChat", lSettings.sMonitorChat
WriteINI lSettings.sIniFile, "Settings", "RootDir", lSettings.sRootDir
WriteINI lSettings.sIniFile, "Settings", "KillDelay", lSettings.sKillDelay
WriteINI lSettings.sIniFile, "Settings", "ServerName", lSettings.sServerName
WriteINI lSettings.sIniFile, "Settings", "Administrator", lSettings.sAdministrator
WriteINI lSettings.sIniFile, "Settings", "ConnectedOnStartup", lSettings.sConnectedOnStartup
WriteINI lSettings.sIniFile, "Settings", "AcceptingConnections", lSettings.sAcceptingConnections
WriteINI lSettings.sIniFile, "Settings", "LoadServerOnStartup", lSettings.sLoadServerOnStartup
WriteINI lSettings.sIniFile, "Settings", "EchoText", lSettings.sEchoText
WriteINI lSettings.sIniFile, "Settings", "ShareOnlyHomeDir", lSettings.sShareOnlyHomeDir
WriteINI lSettings.sIniFile, "Settings", "TelnetOnStartup", lSettings.sTelnetOnStartup
WriteINI lSettings.sIniFile, "Settings", "FSTelnet", lSettings.sFSTelnet
WriteINI lSettings.sIniFile, "Settings", "WhiteBlack", lSettings.sWhiteBlack
WriteINI lSettings.sIniFile, "Settings", "AcceptingNewUsers", lSettings.sAcceptingNewUsers
End Sub

Public Sub FileAttributes(lFilename As String, lAttribute As VbFileAttribute)
On Local Error Resume Next
If DoesFileExist(lFilename) = True Then
    SetAttr lFilename, lAttribute
End If
If Err.Number <> 0 Then ErrorHandle Err.Number, Err.Description, "Public Sub FileAttributes(lFilename As String, lAttribute As VbFileAttribute)"
End Sub

Public Sub BroadcastKillMessage(HowLong, lForm As Form)
On Local Error Resume Next
Dim i As Integer, OldTimer As Single
If lSettings.sClosePause = True Then
    For i = 1 To lSettings.sMaxUsers
        If Len(lUsers.uUser(i).uName) <> 0 Then lForm.Send vbCrLf & lSettings.sServerName & lStrings.sShuttingDown & HowLong & " second(s)", i
    Next
    ConsolePrint "Closing server in " & HowLong & " seconds.", lForm
    OldTimer = Timer
    Do While (Timer - OldTimer) < HowLong
        DoEvents
    Loop
    lForm.Tag = "done"
    Unload lForm
Else
    For i = 1 To lSettings.sMaxUsers
        If Len(lUsers.uUser(i).uName) <> 0 Then lForm.Send vbCrLf & lSettings.sServerName & lStrings.sShuttingDownNow, i
    Next
    ConsolePrint "Closing server now", lForm
    lForm.Tag = "done"
    Unload lForm
End If
If Err.Number <> 0 Then ErrorHandle Err.Number, Err.Description, "Public Sub BroadcastKillMessage(HowLong, lForm As Form)"
End Sub

Public Sub Sleep(Seconds As Single, EventEnable As Boolean)
On Error GoTo ErrHndl
Dim OldTimer As Single
OldTimer = Timer
Do While (Timer - OldTimer) < Seconds
    If EventEnable Then DoEvents
Loop
Exit Sub

ErrHndl:
If Err.Number <> 0 Then ErrorHandle Err.Number, Err.Description, "Public Sub Sleep(Seconds As Single, EventEnable As Boolean)"
Err.Clear
End Sub

Public Function IsAlpha(sData As String) As Boolean
On Local Error Resume Next
If sData = "" Then Exit Function
sData = Mid(sData, 1, 1)
If Asc(sData) >= 65 And Asc(sData) <= 90 Then
    IsAlpha = False
    Exit Function
ElseIf Asc(sData) >= 97 And Asc(sData) <= 122 Then
    IsAlpha = True
    Exit Function
ElseIf Asc(sData) >= 48 And Asc(sData) <= 57 Then
    IsAlpha = False
    Exit Function
End If
IsAlpha = False
If Err.Number <> 0 Then ErrorHandle Err.Number, Err.Description, "Public Function IsAlpha(sData As String) As Boolean"
End Function

Public Function IsAlphaNum(sData As String) As Boolean
On Local Error Resume Next
If sData = "" Then Exit Function
sData = Mid(sData, 1, 1)
If Asc(sData) >= 65 And Asc(sData) <= 90 Then
    IsAlphaNum = True
    Exit Function
ElseIf Asc(sData) >= 97 And Asc(sData) <= 122 Then
    IsAlphaNum = True
    Exit Function
ElseIf Asc(sData) >= 48 And Asc(sData) <= 57 Then
    IsAlphaNum = True
    Exit Function
End If
IsAlphaNum = False
If Err.Number <> 0 Then ErrorHandle Err.Number, Err.Description, "Public Function IsAlphaNum(sData As String) As Boolean"
End Function

Function JustFile(Path As String)
On Local Error Resume Next
Dim Pos As Integer
Pos = InStr(1, StrReverse(Path), "\")
JustFile = Right(Path, Pos - 1)
If Err.Number <> 0 Then ErrorHandle Err.Number, Err.Description, "Function JustFile(Path As String)"
End Function

Function JustFile2(Path As String)
On Local Error Resume Next
Dim Pos As Integer
Pos = InStr(1, StrReverse(Path), "\")
JustFile2 = Right(Path, Pos - 1)
If JustFile2 = "fs" Then
    JustFile2 = "\"
End If
If Err.Number <> 0 Then ErrorHandle Err.Number, Err.Description, "Function JustFile2(Path As String)"
End Function

Public Sub ActivateServerToggle(lEnabled As Boolean, lForm As Form, Optional lPortNumber As Long)
On Local Error Resume Next
If LCase(Trim(lForm.Tag)) = "telnet" Then
    Select Case lEnabled
    Case True
        If lForm.wskTelnet.State = sckConnected Then
            lForm.AddText "Already connected" & vbCrLf, vbBlue
            Exit Sub
        Else
            lForm.wskTelnet.Connect lSettings.sIpAddress, "23"
            Exit Sub
        End If
    Case False
        If lForm.wskTelnet.State = sckConnected Then
            lForm.wskTelnet.Close
            lForm.AddText vbCrLf & "Disconnected ", vbBlue
            Exit Sub
        End If
    End Select
Else
    lForm.Tag = str(lPortNumber)
    If lEnabled = True Then
        If lForm.acc.State = 2 Then
            ConsolePrint "Already connected", lForm
            Exit Sub
        End If
        lForm.Caption = "Telos - Status (" & lForm.acc.LocalIP & ") Connected"
        If lForm.acc.LocalPort = lPortNumber Then
            If lForm.acc.State = 2 Then
                Exit Sub
            End If
        End If
        If lPortNumber <> 0 Then
            lForm.acc.LocalPort = lPortNumber
            lForm.acc.Bind
        Else
            lForm.acc.LocalPort = 23
            lForm.acc.Bind
        End If
        If Err.Number = 10048 Then
            ConsolePrint "Unable to activate this server, reached address in use error", lForm
            ConsolePrint "To activate this server with another port, type ...", lForm
            ConsolePrint "'/server <port number>' in the status window.", lForm
            Exit Sub
        End If
        lForm.acc.Listen
        ConsolePrint "Server active", lForm
    ElseIf lEnabled = False Then
        lForm.Caption = "Telos - Status (" & lForm.acc.LocalIP & ") Disconnected"
        lForm.acc.Close
        ConsolePrint "Server Closed", lForm
    End If
    Exit Sub
End If
If Err.Number <> 0 Then ErrorHandle Err.Number, Err.Description, "Public Sub ActivateServerToggle(lEnabled As Boolean, lForm As Form, Optional lPortNumber As Long)"
End Sub

Sub PrintManPage(cmd, refid, lForm As Form)
On Local Error Resume Next
lForm.SendFile "files\man\" & cmd & ".txt", refid
If Err.Number <> 0 Then ErrorHandle Err.Number, Err.Description, "Sub PrintManPage(cmd, refid, lForm As Form)"
End Sub

Function IfDirectoryExists(dirpath As String) As Boolean
On Local Error Resume Next
Dim f As String, dirFolder As String
f$ = dirpath
dirFolder = Dir(f$, vbDirectory)
If dirFolder <> "" Then
    IfDirectoryExists = True
End If
If Err.Number <> 0 Then ErrorHandle Err.Number, Err.Description, "Function IfDirectoryExists(dirpath As String) As Boolean"
End Function

Sub StartConsole(lForm As Form)
On Local Error Resume Next
If lSettings.sRegistered = False Then
    lForm.Visible = True
    ConsolePrint "Telos v" & App.Major & "." & App.Revision & " Unregistered 35 node version", lForm
    Sleep 2, True
Else
    lForm.Visible = True
    ConsolePrint "Telos v" & App.Major & "." & App.Revision & " Registered 1000 node version", lForm
End If
If Err.Number <> 0 Then ErrorHandle Err.Number, Err.Description, "Sub StartConsole(lForm As Form)"
End Sub

Public Sub ConsolePrint(szOut As String, lForm As Form)
On Local Error Resume Next
Trim (szOut)
lForm.console.AddItem szOut
lForm.console.Selected(lForm.console.ListCount - 1) = True
If Err.Number <> 0 Then ErrorHandle Err.Number, Err.Description, "Public Sub ConsolePrint(szOut As String, lForm As Form)"
End Sub
