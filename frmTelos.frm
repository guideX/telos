VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form frmTelos 
   Caption         =   "Telos - Disconnected"
   ClientHeight    =   2055
   ClientLeft      =   1020
   ClientTop       =   1305
   ClientWidth     =   7155
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTelos.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   2055
   ScaleWidth      =   7155
   Begin VB.TextBox txtOutgoing 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   250
      Left            =   0
      TabIndex        =   0
      Top             =   1560
      Width           =   7095
   End
   Begin VB.ListBox console 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1455
      IntegralHeight  =   0   'False
      Left            =   1680
      TabIndex        =   5
      Top             =   0
      Width           =   5415
   End
   Begin VB.ListBox lstUsers 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      IntegralHeight  =   0   'False
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1575
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   3480
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.DriveListBox drives 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   0
      TabIndex        =   4
      Top             =   3000
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.FileListBox files 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Index           =   1
      Left            =   0
      TabIndex        =   3
      Top             =   3840
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.DirListBox folders 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   0
      TabIndex        =   2
      Top             =   5040
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSWinsockLib.Winsock acc 
      Left            =   0
      Top             =   2160
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock pol 
      Index           =   0
      Left            =   480
      Top             =   2160
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.ListBox lstProcess 
      Appearance      =   0  'Flat
      Height          =   420
      ItemData        =   "frmTelos.frx":0CCA
      Left            =   0
      List            =   "frmTelos.frx":0CCC
      TabIndex        =   7
      Top             =   0
      Width           =   615
   End
End
Attribute VB_Name = "frmTelos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WaitText As String

Private Sub acc_ConnectionRequest(ByVal requestID As Long)
On Local Error Resume Next
Dim scan As Integer, refid As String
lWinsockCount = lWinsockCount + 1
Load pol(lWinsockCount)
Load files(lWinsockCount)
Load folders(lWinsockCount)
If Err.Number <> 0 Then Err = 0
Me.folders(lWinsockCount).Path = lSettings.sRootDir
Me.files(lWinsockCount).Path = lSettings.sRootDir
pol(lWinsockCount).Close
pol(lWinsockCount).Accept requestID
acc.Close
acc.Listen
For scan = 1 To lSettings.sMaxUsers
    With lUsers.uUser(scan)
        If .uName = Empty Then
            'If Ac_Name(scan) = Empty Then
            refid = scan
            Exit For
        End If
    End With
Next scan
With lUsers.uUser(refid)
    .uName = "unknown"
    .uHost = pol(lWinsockCount).RemoteHostIP
    .uWhat = "login"
    .uSock = lWinsockCount
    JustFile(Me.folders(refid).Path) = "\"
    ConsolePrint pol(lWinsockCount).RemoteHostIP & " connected", Me
    SendFile "files\connect.txt", refid
    Send Crt, refid
    Send lStrings.sEnterUsername, refid
    'Send "Username: ", refid
    Update
End With
End Sub

Sub Update()
On Local Error Resume Next
Dim scan As Integer, p As Integer
lstUsers.Clear
For scan = 1 To lSettings.sMaxUsers
    With lUsers.uUser(scan)
        If .uName <> Empty Then
            If .uOperator = False Then
                lstUsers.AddItem .uName & " - " & .uHost
            Else
                lstUsers.AddItem "@" & .uName & " - " & .uHost
            End If
            p = p + 1
        End If
    End With
Next scan
Form_Resize
Caption = "Telos - Status (" & pol(0).LocalIP & ") - " & Trim(p) & " connection(s)"
End Sub

Private Sub Form_Load()
On Local Error Resume Next
Dim msg As String
If lSettings.sWhiteBlack = True Then
    lstUsers.BackColor = vbBlack
    lstUsers.ForeColor = vbWhite
    console.BackColor = vbBlack
    console.ForeColor = vbWhite
    txtOutgoing.BackColor = vbBlack
    txtOutgoing.ForeColor = vbWhite
ElseIf lSettings.sColoredWindows = True Then
    msg = GetRnd(14700000)
    txtOutgoing.BackColor = msg
    console.BackColor = msg
    lstUsers.BackColor = msg
End If
lSettings.sIpAddress = acc.LocalIP
Crt = vbCrLf
StartConsole Me
If lSettings.sConnectedOnStartup = True Then
    ActivateServerToggle True, Me
    If lSettings.sTelnetOnStartup = True Then LoadClient
Else
    ConsolePrint "Waiting for " & lSettings.sAdministrator & " to start the server", Me
    ConsolePrint "To load the server, type '/server 23'", Me
End If
End Sub

Private Sub Form_Resize()
On Local Error Resume Next
If lstUsers.ListCount <> 0 Then
    lstUsers.Visible = True
    lstUsers.Top = 0
    lstUsers.Left = 0
    lstUsers.Width = (Me.Width / 4)
    If Me.ScaleHeight <> 0 Then lstUsers.Height = (Me.ScaleHeight - txtOutgoing.Height)
    console.Top = 0
    If Me.ScaleWidth <> 0 Then console.Width = (Me.ScaleWidth - lstUsers.Width)
    If Me.ScaleHeight <> 0 Then console.Height = (Me.ScaleHeight - txtOutgoing.Height)
    console.Left = lstUsers.Width
    txtOutgoing.Left = 0
    txtOutgoing.Width = Me.ScaleWidth
    txtOutgoing.Top = (Me.ScaleHeight - txtOutgoing.Height)
Else
    lstUsers.Visible = False
    console.Top = 0
    console.Left = 0
    If Me.ScaleWidth <> 0 Then console.Width = (Me.ScaleWidth)
    If Me.ScaleHeight <> 0 Then console.Height = (Me.ScaleHeight - txtOutgoing.Height)
    txtOutgoing.Left = 0
    txtOutgoing.Width = Me.ScaleWidth
    txtOutgoing.Top = (Me.ScaleHeight - txtOutgoing.Height)
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Local Error Resume Next
Dim TM As Single
TM = lSettings.sKillDelay
If TM <> 0 Then
    BroadcastKillMessage TM, Me
    If Me.Tag <> "done" Then Cancel = 1
End If
End Sub

Private Sub lstUsers_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
On Local Error Resume Next
If Button = 2 Then
    If Len(lstUsers.text) <> 0 Then PopupMenu mdiTelos.mnuNicklist
End If
End Sub

Private Sub pol_Close(Index As Integer)
On Local Error Resume Next
Dim scan As Integer, refid As String
For scan = 1 To lSettings.sMaxUsers
    With lUsers.uUser(scan)
        If .uSock = Index Then
            refid = scan
            Exit For
        End If
    End With
Next scan
With lUsers.uUser(refid)
    ConsolePrint .uName & " disconnected.", Me
    .uName = Empty
    .uInput = Empty
    .uHost = Empty
    .uWhat = Empty
    .uSock = Empty
    .uWTF = Empty
    .uRelDir = Empty
End With
pol(Index).Close
Update
End Sub

Sub SendFile(ByVal FileName As String, ByVal person As Integer)
On Error GoTo ErrSpot
Dim temp As String
If InStr(1, FileName, ":") Then
Else
    FileName = App.Path & "\" & FileName
End If
If DoesFileExist(FileName) = True Then
    Open FileName For Input As #1
    Do
        If EOF(1) Then Exit Do
        Line Input #1, temp
'        Send temp & Chr(10) & Chr(13), person
        Send vbCrLf & temp, person
    Loop
    Close #1
Else
    Send "File not found" & vbCrLf, person
End If
ErrSpot:
End Sub

Sub Send(ByVal text As String, ByVal person As Integer)
On Local Error Resume Next
With lUsers.uUser(person)
    If .uName = "" Then Exit Sub
    If pol(.uSock).State = sckConnected Then
        pol(.uSock).SendData text
    End If
End With
End Sub

Private Sub pol_DataArrival(Index As Integer, ByVal bytesTotal As Long)
On Local Error Resume Next
Dim text As String, scan As Integer, refid As String, stack As String, h As Integer, pg As String, goodcom As Boolean, i_command As String, i_arg As String, temp As String, g As Integer, load_name As String, load_password As String, load_su, msg As String
pol(Index).GetData text, vbString
For scan = 1 To lSettings.sMaxUsers
    If lUsers.uUser(scan).uSock = Index Then
        refid = scan
        Exit For
    End If
Next scan
stack = ""
If refid = 0 Then
    pol(Index).Close
    Exit Sub
End If
With lUsers.uUser(refid)
    For h = 1 To Len(text)
        pg = Mid(text, h, 1)
        If pg = Chr(13) Then
            If .uWhat = "prompt" Then
                .uInput = Trim(.uInput)
                'Send Crt, 1
                If .uInput = Empty Then goodcom = True
                For scan = 1 To Len(.uInput)
                    If Mid(.uInput, scan, 1) = " " Then
                        i_command = Mid(.uInput, 1, scan - 1)
                        i_arg = Mid(.uInput, scan + 1, 100)
                        Exit For
                    End If
                Next scan
                If i_command = "" Then i_command = .uInput
                If .uWTF = True Then
                    WaitText = i_command & " " & i_arg
                Else
                    goodcom = CommandParse(i_command, i_arg, True, Me, refid)
                End If
                If goodcom = False Then
                    'stack = stack & vbCrLf & i_command & ": " & reason & Crt
                    stack = stack '& vbCrLf
                End If
                .uInput = Empty
                If .uWTF = False Then
                    stack = stack & vbCrLf & "[" & .uName & "@" & lSettings.sServerName & "] [" & JustFile2(Me.folders(refid).Path) & "]: "
                    Send stack, refid
                End If
                Exit Sub
            End If
            If LCase(.uWhat) = "new1" Then
                msg = .uInput
                If Len(msg) > 8 Or IsAlpha(msg) = False Or InStr(msg, " ") Then
                    Send vbCrLf & "  - Username is too long, or not alpha charecters only", refid
                    Send vbCrLf & "Enter new username: ", refid
                    .uInput = ""
                    .uWhat = "new1"
                    Exit Sub
                End If
                .uName = .uInput
                .uWhat = "new2"
                .uInput = ""
                Send vbCrLf & "  - You are now known as " & .uName, refid
                Send vbCrLf & "Enter new password: ", refid
                Exit Sub
            ElseIf LCase(.uWhat) = "new2" Then
                msg = .uInput
                If Len(msg) > 8 Or IsAlpha(msg) = False Then
                    .uInput = ""
                    Send vbCrLf & "  - Password is too long, or not alpha charecters" & .uName, refid
                    Send vbCrLf & "Enter new password: ", refid
                    Exit Sub
                End If
                .uPassword = .uInput
                .uWhat = "new3"
                .uInput = ""
                Send vbCrLf & "Re-enter new password: ", refid
                Exit Sub
            ElseIf LCase(.uWhat) = "new3" Then
                If .uPassword = .uInput Then
                    Send vbCrLf & "  Passwords match", refid
                    .uWhat = "new4"
                    .uInput = ""
                    SendFile App.Path & "\files\newsignup.txt", refid
                Else
                    Send vbCrLf & "  Passwords do not match", refid
                    Send vbCrLf & "Enter new password: ", refid
                    .uWhat = "new2"
                    .uInput = ""
                End If
                Exit Sub
            ElseIf LCase(.uWhat) = "new4" Then
                If LCase(.uInput) = "y" Then
                    Send vbCrLf & "User " & .uName & " created", refid
                    DoEvents
                    temp = .uName & "," & .uPassword & ",0,"
                    Open App.Path & "\files\users.txt" For Append As #1
                    Print #1, temp
                    Close #1
                    .uWhat = "login"
                    .uInput = ""
                    SendFile App.Path & "\files\connect.txt", refid
                    'Send vbCrLf & vbCrLf & "Username: ", refid
                    Send vbCrLf & vbCrLf & lStrings.sEnterUsername, refid
                    Exit Sub
                ElseIf LCase(.uInput) = "n" Then
                    LogoutUser Int(refid), Me
                Else
                    .uInput = ""
                    Exit Sub
                End If
            End If
            If .uWhat = "login" Then
                If .uInput = Empty Then
                    stack = stack & Crt
                    stack = stack & Crt
                    'stack = stack & "Username: "
                    stack = stack & lStrings.sEnterUsername
                    Send stack, refid
                    Exit Sub
                End If
                If LCase(.uInput) = "new" Then
                    If lSettings.sAcceptingNewUsers = True Then
                        .uInput = ""
                        .uWhat = "new1"
                        Send vbCrLf & "Enter new username: ", refid
                        Exit Sub
                    Else
                        Send vbCrLf & "Not accepting new users at this time, try back later", refid
                        LogoutUser Index, Me, "Not accepting new users"
                    End If
                End If
                .uName = .uInput
                .uInput = Empty
                stack = stack & Crt
                'stack = stack & "Password: "
                stack = stack & lStrings.sEnterPassword
                .uWhat = "password"
                Send stack, refid
                Exit Sub
            End If
            If .uWhat = "new" Then
            
            End If
            If .uWhat = "password" Then
                Open App.Path & "\files\users.txt" For Input As #1
                Do
                    If EOF(1) Then Exit Do
                    Line Input #1, temp
                    If Mid(temp, 1, 1) <> "#" Then
                        g = 0
rscan:
                        For scan = 1 To Len(temp)
                            If Mid(temp, scan, 1) = "," Then
                                g = g + 1
                                If g = 1 Then
                                    load_name = Mid(temp, 1, scan - 1)
                                    temp = Mid(temp, scan + 1, 100)
                                    GoTo rscan
                                End If
                                If g = 2 Then
                                    load_password = Mid(temp, 1, scan - 1)
                                    temp = Mid(temp, scan + 1, 100)
                                    GoTo rscan
                                End If
                                If g = 3 Then
                                    load_su = Mid(temp, 1, scan - 1)
                                    temp = Mid(temp, scan + 1, 100)
                                End If
                                If lSettings.sAcceptingConnections = True Then
                                    If load_name = .uName Then
                                        If load_password = .uInput Then
                                            Dim m As Integer
                                            For m = 1 To pol.Count
                                                Send vbCrLf & .uName & " logged in", m
                                            Next m
                                            ConsolePrint pol(Index).RemoteHostIP & " is now known as " & .uName, Me
                                            'stack = stack & Crt
                                            'stack = stack & Ac_Name(refid) & "@" & lSettings.sServerName & " " & JustFile2(frmTelos.folders(refid).Path) & "> "
                                            stack = stack & "[" & .uName & "@" & lSettings.sServerName & "] [" & JustFile2(Me.folders(refid).Path) & "]: "
                                            .uWhat = "prompt"
                                            .uInput = Empty
                                            .uOperator = False
                                            If load_su = "1" Then
                                                .uOperator = True
                                            End If
                                            Close #1
                                            Send vbCrLf, refid
                                            SendFile "files\welcome.txt", refid
                                            Send vbCrLf & stack, refid
                                            Update
                                            Exit Sub
                                        End If
                                        .uInput = Empty
                                    End If
                                Else
                                    Send vbCrLf & "This server is currently not accepting connections" & vbCrLf, refid
                                End If
                            End If
                        Next scan
                    End If
                Loop
                Close #1
                .uInput = Empty
                stack = stack & Crt
                If .uPasswordRetries = lSettings.sMaxPasswordRetries Then
                    Send vbCrLf & "Maximum password retries used, goodbye" & vbCrLf, refid
                    DoEvents
                    'Sleep 2, False
                    pol(Index).Close
                    Exit Sub
                End If
                .uPasswordRetries = .uPasswordRetries + 1
                stack = vbCrLf & "Retry: " & .uPasswordRetries & vbCrLf & " " & lStrings.sBadPassword & vbCrLf & lStrings.sEnterUsername
                Send stack, refid
                .uWhat = "login"
                Exit Sub
            End If
        End If
        If pg = Chr(8) Then
            If .uInput <> "" Then
                .uInput = Mid(.uInput, 1, Len(.uInput) - 1)
                If .uWhat <> "password" Then
                    Send " " & Chr(8), refid
                End If
            End If
            Exit Sub
        End If
        If pg = Chr(21) Then
            If .uInput <> "" Then
                For g = 1 To Len(.uInput)
                    Send Chr(8) & " " & Chr(8), refid
                Next g
            End If
            .uInput = ""
            Exit Sub
        End If
        If .uWhat <> "password" Then
            If lSettings.sEchoText = True Then
                'makes text double in xp/2000 where telnet echo settings is on
                Send pg, refid
            End If
        End If
        If .uWTF = True Then
            WaitText = WaitText & text
        End If
        .uInput = .uInput & pg
    Next h
End With
End Sub

Private Sub pol_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
On Local Error Resume Next
Dim scan As Integer, refid As String
For scan = 1 To lSettings.sMaxUsers
    With lUsers.uUser(scan)
        If .uSock = Index Then
            refid = scan
            Exit For
        End If
    End With
Next scan
With lUsers.uUser(refid)
    .uName = Empty
    .uInput = Empty
    .uHost = Empty
    .uWhat = Empty
    .uSock = Empty
    .uRelDir = Empty
End With
pol(Index).Close
Update
End Sub

Public Function WaitForText(refid) As String
On Local Error Resume Next
With lUsers.uUser(refid)
    .uInput = ""
    .uWTF = True
    Do
    DoEvents
    Loop Until WaitText <> ""
        WaitForText = WaitText
        WaitText = ""
        .uWTF = False
End With
End Function

Sub RootManPage(cmd As String)
On Local Error Resume Next
Dim TempVar As String
Open "files\man\" & cmd & ".txt" For Input As #1
While EOF(1) = False
Line Input #1, TempVar
ConsolePrint TempVar, Me
Wend
Close #1
End Sub

Private Sub txtOutgoing_KeyDown(KeyCode As Integer, Shift As Integer)
On Local Error Resume Next
Static LastSent(1 To 20) As String
Static l As Integer
If KeyCode = 38 Then
    l = l + 1
    If l > 10 Then l = 1
    txtOutgoing.text = LastSent(l)
End If
If KeyCode = 40 Then
    If l <> 0 Then
        l = l - 1
        txtOutgoing.text = LastSent(l)
        Exit Sub
    End If
End If
If KeyCode = 13 Then
    l = 0
    LastSent(10) = LastSent(9)
    LastSent(9) = LastSent(8)
    LastSent(8) = LastSent(7)
    LastSent(7) = LastSent(6)
    LastSent(6) = LastSent(5)
    LastSent(5) = LastSent(4)
    LastSent(4) = LastSent(3)
    LastSent(3) = LastSent(2)
    LastSent(2) = LastSent(1)
    LastSent(1) = txtOutgoing.text
'    txtOutgoing.text = ""
End If
If Err.Number <> 0 Then ErrorHandle Err.Number, Err.Description, "Private Sub txtOutgoing_KeyDown(KeyCode As Integer, Shift As Integer)"
End Sub

Private Sub txtOutgoing_KeyPress(KeyAscii As Integer)
On Local Error Resume Next
Dim i As Integer, lPortNumber As Long, t As Form
If KeyAscii = 27 Then
    Me.WindowState = vbMinimized
End If
If KeyAscii = 13 Then
    If Len(txtOutgoing.text) <> 0 Then
        If Left(LCase(txtOutgoing.text), 8) = "/server " Then
            lPortNumber = Int(Right(txtOutgoing.text, Len(txtOutgoing.text) - 8))
            console.Clear
            ActivateServerToggle True, Me, lPortNumber
            txtOutgoing.text = ""
            KeyAscii = 0
            Exit Sub
        End If
        Select Case LCase(txtOutgoing.text)
        Case "/clear"
            Me.console.Clear
            txtOutgoing.text = ""
            KeyAscii = 0
            Exit Sub
        Case "/end"
            End
            Exit Sub
        Case "/help"
            ConsolePrint "Commands:", Me
            ConsolePrint "/end - Removes Telos from memory", Me
            ConsolePrint "/exit - Unloads Telos", Me
            ConsolePrint "/clear - Clears the console of your server", Me
            ConsolePrint "/l or /login - Connects you to your server as a client", Me
            ConsolePrint "/server <port number> - Resets your listening socket to your specified port number", Me
            txtOutgoing.text = ""
            KeyAscii = 0
            Exit Sub
        Case "/exit"
            Unload Me
            Exit Sub
        Case "/l"
            LoadClient
            KeyAscii = 0
            txtOutgoing.text = ""
            Exit Sub
        Case "/login"
            LoadClient
            KeyAscii = 0
            txtOutgoing.text = ""
            Exit Sub
        End Select
        For i = 1 To pol.Count
            With lUsers.uUser(i)
                If Len(.uName) <> 0 Then
                    Send vbCrLf & "<" & lSettings.sAdministrator & "> " & txtOutgoing.text, i
                End If
            End With
        Next i
        ConsolePrint "[" & lSettings.sServerName & "] <" & lSettings.sAdministrator & "> " & txtOutgoing.text, Me
        KeyAscii = 0
        txtOutgoing.text = ""
        Exit Sub
    End If
End If
End Sub
