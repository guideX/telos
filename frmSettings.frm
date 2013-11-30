VERSION 5.00
Object = "{B3FF7FA6-B059-4900-8BEC-5C65E3D9C033}#1.0#0"; "xplook.ocx"
Begin VB.Form frmSettings 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Telos - Settings"
   ClientHeight    =   5235
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4560
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5235
   ScaleWidth      =   4560
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton optSettings 
      Caption         =   "Users"
      Height          =   375
      Index           =   1
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.OptionButton optSettings 
      Caption         =   "Strings"
      Height          =   375
      Index           =   2
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   0
      Width           =   855
   End
   Begin VB.OptionButton optSettings 
      Caption         =   "General"
      Height          =   375
      Index           =   0
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   855
   End
   Begin VB.Frame fraSettings 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   4335
      Index           =   2
      Left            =   0
      TabIndex        =   33
      Top             =   360
      Visible         =   0   'False
      Width           =   4575
      Begin VB.TextBox txtShuttingDownNow 
         Height          =   285
         Left            =   1320
         TabIndex        =   50
         Top             =   3000
         Width           =   3135
      End
      Begin VB.TextBox txtAccessDenied 
         Height          =   285
         Left            =   1320
         TabIndex        =   41
         Top             =   120
         Width           =   3135
      End
      Begin VB.TextBox txtBadPassword 
         Height          =   285
         Left            =   1320
         TabIndex        =   40
         Top             =   480
         Width           =   3135
      End
      Begin VB.TextBox txtLoggedOut 
         Height          =   285
         Left            =   1320
         TabIndex        =   39
         Top             =   840
         Width           =   3135
      End
      Begin VB.TextBox txtNick 
         Height          =   285
         Left            =   1320
         TabIndex        =   38
         Top             =   1200
         Width           =   3135
      End
      Begin VB.TextBox txtDirectoryError 
         Height          =   285
         Left            =   1320
         TabIndex        =   37
         Top             =   1560
         Width           =   3135
      End
      Begin VB.TextBox txtEnterUsername 
         Height          =   285
         Left            =   1320
         TabIndex        =   36
         Top             =   1920
         Width           =   3135
      End
      Begin VB.TextBox txtEnterPassword 
         Height          =   285
         Left            =   1320
         TabIndex        =   35
         Top             =   2280
         Width           =   3135
      End
      Begin VB.TextBox txtShuttingDown 
         Height          =   285
         Left            =   1320
         TabIndex        =   34
         Top             =   2640
         Width           =   3135
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Shutdown Now:"
         Height          =   255
         Left            =   120
         TabIndex        =   51
         Top             =   3000
         Width           =   1215
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Access Denied:"
         Height          =   255
         Left            =   120
         TabIndex        =   49
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Bad Password:"
         Height          =   255
         Left            =   120
         TabIndex        =   48
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Logged Out:"
         Height          =   255
         Left            =   120
         TabIndex        =   47
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Nick Change:"
         Height          =   255
         Left            =   120
         TabIndex        =   46
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Directory Error:"
         Height          =   255
         Left            =   120
         TabIndex        =   45
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Username:"
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Password:"
         Height          =   255
         Left            =   120
         TabIndex        =   43
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Shutting Down:"
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   2640
         Width           =   1215
      End
   End
   Begin OsenXPCntrl.OsenXPButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   3600
      TabIndex        =   32
      Top             =   4800
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Cancel"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   0
      FCOLO           =   16711680
      MCOL            =   12632256
      MPTR            =   0
      MICON           =   "frmSettings.frx":0CCA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin OsenXPCntrl.OsenXPButton cmdOK 
      Height          =   375
      Left            =   2640
      TabIndex        =   31
      Top             =   4800
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "OK"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   0
      FCOLO           =   16711680
      MCOL            =   12632256
      MPTR            =   0
      MICON           =   "frmSettings.frx":0CE6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame fraSettings 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "General"
      Height          =   4095
      Index           =   0
      Left            =   0
      TabIndex        =   3
      Top             =   360
      Visible         =   0   'False
      Width           =   4455
      Begin VB.CheckBox chkShareOnlyHomeDir 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Share home dir only"
         ForeColor       =   &H80000008&
         Height          =   250
         Left            =   2280
         TabIndex        =   26
         Top             =   2520
         Width           =   1815
      End
      Begin VB.CheckBox chkMonitorChat 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Monitor chat"
         ForeColor       =   &H80000008&
         Height          =   250
         Left            =   2280
         TabIndex        =   25
         Top             =   2280
         Width           =   1335
      End
      Begin VB.CheckBox chkFolderChat 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Use folders as chatrooms"
         ForeColor       =   &H80000008&
         Height          =   250
         Left            =   2280
         TabIndex        =   30
         Top             =   3240
         Width           =   2175
      End
      Begin VB.CheckBox chkWhiteBlack 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Use only black and white"
         ForeColor       =   &H80000008&
         Height          =   250
         Left            =   2280
         TabIndex        =   29
         Top             =   3480
         Width           =   2175
      End
      Begin VB.CheckBox chkTelnetOnStartup 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Telnet client on startup"
         ForeColor       =   &H80000008&
         Height          =   250
         Left            =   2280
         TabIndex        =   28
         Top             =   3000
         Width           =   2055
      End
      Begin VB.CheckBox chkShowUsersInDir 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Show users in dir"
         ForeColor       =   &H80000008&
         Height          =   250
         Left            =   2280
         TabIndex        =   27
         Top             =   2760
         Width           =   1575
      End
      Begin VB.CheckBox chkLoadServerOnStartup 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Load server on startup"
         ForeColor       =   &H80000008&
         Height          =   250
         Left            =   120
         TabIndex        =   24
         Top             =   3720
         Width           =   2055
      End
      Begin VB.CheckBox chkFSTelnet 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Full screen Telnet"
         ForeColor       =   &H80000008&
         Height          =   250
         Left            =   120
         TabIndex        =   23
         Top             =   3480
         Width           =   1695
      End
      Begin VB.CheckBox chkEchoText 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Echo text"
         ForeColor       =   &H80000008&
         Height          =   250
         Left            =   120
         TabIndex        =   22
         Top             =   3240
         Width           =   1095
      End
      Begin VB.CheckBox chkColoredWindows 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Colored windows"
         ForeColor       =   &H80000008&
         Height          =   250
         Left            =   120
         TabIndex        =   21
         Top             =   2760
         Width           =   1695
      End
      Begin VB.CheckBox chkConnectedOnStartup 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Connected on startup"
         ForeColor       =   &H80000008&
         Height          =   250
         Left            =   120
         TabIndex        =   20
         Top             =   3000
         Width           =   1935
      End
      Begin VB.CheckBox chkAcceptingNewUsers 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Accepting new users"
         ForeColor       =   &H80000008&
         Height          =   250
         Left            =   120
         TabIndex        =   19
         Top             =   2520
         Width           =   1935
      End
      Begin VB.CheckBox chkAcceptingConnections 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Accepting connections"
         ForeColor       =   &H80000008&
         Height          =   250
         Left            =   120
         TabIndex        =   18
         Top             =   2280
         Width           =   2055
      End
      Begin VB.ComboBox cboKillDelay 
         Height          =   315
         ItemData        =   "frmSettings.frx":0D02
         Left            =   1560
         List            =   "frmSettings.frx":0D04
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   1560
         Width           =   2775
      End
      Begin VB.CheckBox chkClosePause 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Close Pause"
         ForeColor       =   &H80000008&
         Height          =   250
         Left            =   120
         TabIndex        =   15
         Top             =   1580
         Width           =   1335
      End
      Begin VB.ComboBox cboPasswordRetries 
         Height          =   315
         ItemData        =   "frmSettings.frx":0D06
         Left            =   1560
         List            =   "frmSettings.frx":0D10
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   1920
         Width           =   2775
      End
      Begin VB.TextBox txtServerName 
         Height          =   285
         Left            =   1560
         TabIndex        =   12
         Top             =   1200
         Width           =   2775
      End
      Begin VB.CommandButton cmdChangeDir 
         Caption         =   "..."
         Height          =   255
         Left            =   3840
         TabIndex        =   10
         Top             =   840
         Width           =   495
      End
      Begin VB.TextBox txtPath 
         Height          =   285
         Left            =   1560
         TabIndex        =   9
         Top             =   840
         Width           =   2175
      End
      Begin VB.TextBox txtEMail 
         Height          =   285
         Left            =   1560
         TabIndex        =   7
         Top             =   480
         Width           =   2775
      End
      Begin VB.TextBox txtAdministrator 
         Height          =   285
         Left            =   1560
         TabIndex        =   5
         Top             =   120
         Width           =   2775
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Password Retries:"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   1950
         Width           =   1455
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Server Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Path:"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   840
         Width           =   2055
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "E-Mail:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Administrator:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   2055
      End
   End
   Begin VB.Frame fraSettings 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   4335
      Index           =   1
      Left            =   0
      TabIndex        =   4
      Top             =   360
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.Image Image1 
      Height          =   600
      Left            =   0
      Picture         =   "frmSettings.frx":0D1A
      Stretch         =   -1  'True
      Top             =   4680
      Width           =   4560
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub SettingsChange()
On Local Error Resume Next
Dim i As Integer
lStrings.sDirectoryError = txtDirectoryError.text
lStrings.sAccessDenied = txtAccessDenied.text
lStrings.sBadPassword = txtBadPassword.text
lStrings.sLoggedOut = txtLoggedOut.text
lStrings.sNick = txtNick.text
lStrings.sEnterUsername = txtEnterUsername.text
lStrings.sEnterPassword = txtEnterPassword.text
lStrings.sShuttingDown = txtShuttingDown.text
lStrings.sShuttingDownNow = txtShuttingDownNow.text
lSettings.sRootDir = txtPath.text
lSettings.sRootDirLen = Len(txtPath.text)
lSettings.sKillDelay = cboKillDelay.ListIndex + 1
lSettings.sMaxPasswordRetries = cboPasswordRetries.ListIndex + 1
lSettings.sServerName = txtServerName.text
lSettings.sAdministrator = txtAdministrator.text
lSettings.sEMail = txtEMail.text
If chkShowUsersInDir.Value = 1 Then
    lSettings.sShowUsersInDir = True
Else
    lSettings.sShowUsersInDir = False
End If
If chkAcceptingNewUsers.Value = 1 Then
    lSettings.sAcceptingNewUsers = True
Else
    lSettings.sAcceptingNewUsers = False
End If
If chkClosePause.Value = 1 Then
    lSettings.sClosePause = True
Else
    lSettings.sClosePause = False
End If
If chkFSTelnet.Value = 1 Then
    lSettings.sFSTelnet = True
Else
    lSettings.sFSTelnet = False
End If
If chkWhiteBlack.Value = 1 Then
    lSettings.sWhiteBlack = True
Else
    lSettings.sWhiteBlack = False
End If
If chkFolderChat.Value = 1 Then
    lSettings.sFolderChat = True
Else
    lSettings.sFolderChat = False
End If
If chkColoredWindows.Value = 1 Then
    lSettings.sColoredWindows = True
Else
    lSettings.sColoredWindows = False
End If
If chkTelnetOnStartup.Value = 1 Then
    lSettings.sTelnetOnStartup = True
Else
    lSettings.sTelnetOnStartup = False
End If
If chkShareOnlyHomeDir.Value = 1 Then
    lSettings.sShareOnlyHomeDir = True
Else
    lSettings.sShareOnlyHomeDir = False
End If
If chkLoadServerOnStartup.Value = 1 Then
    lSettings.sLoadServerOnStartup = True
Else
    lSettings.sLoadServerOnStartup = False
End If
If chkMonitorChat.Value = 1 Then
    lSettings.sMonitorChat = True
Else
    lSettings.sMonitorChat = False
End If
If chkEchoText.Value = 1 Then
    lSettings.sEchoText = True
Else
    lSettings.sEchoText = False
End If
If chkConnectedOnStartup.Value = 1 Then
    lSettings.sConnectedOnStartup = True
Else
    lSettings.sConnectedOnStartup = False
End If
If chkAcceptingConnections.Value = 1 Then
    lSettings.sAcceptingConnections = True
Else
    lSettings.sAcceptingConnections = False
End If
SaveCurrentSettings
If lSettings.sWhiteBlack = True Then
    For i = 0 To mdiTelos.Count
        If mdiTelos.ActiveForm.Tag = "Telnet" Then
            mdiTelos.ActiveForm.txtIncoming.BackColor = vbBlack
            mdiTelos.ActiveForm.txtOutgoing.BackColor = vbBlack
            
            mdiTelos.ActiveForm.WindowState = vbMinimized
        Else
            mdiTelos.ActiveForm.lstUsers.BackColor = vbBlack
            mdiTelos.ActiveForm.lstUsers.ForeColor = vbWhite
            mdiTelos.ActiveForm.console.BackColor = vbBlack
            mdiTelos.ActiveForm.console.ForeColor = vbWhite
            mdiTelos.ActiveForm.WindowState = vbMinimized
        End If
    Next i
    For i = 0 To mdiTelos.Count
        mdiTelos.ActiveForm.WindowState = vbNormal
    Next i
Else
    
End If
End Sub

Private Sub chkClosePause_Click()
On Local Error Resume Next
If chkClosePause.Value = 1 Then
    cboKillDelay.Enabled = True
Else
    cboKillDelay.Enabled = False
End If
End Sub

Private Sub cmdCancel_Click()
On Local Error Resume Next
Unload Me
End Sub

Private Sub cmdChangeDir_Click()
On Local Error Resume Next
Dim msg As String
msg = InputBox("Select Directory:")
txtPath.text = msg
End Sub

Private Sub cmdClose_Click()
On Local Error Resume Next
Unload Me
End Sub

Private Sub cmdOK_Click()
On Local Error Resume Next
SettingsChange
Unload Me
End Sub

Private Sub Form_Load()
On Local Error Resume Next
Dim i As Integer, m As Integer, s As Integer
Me.Icon = mdiTelos.Icon
fraSettings(0).Visible = True
optSettings(0).Value = True
txtAccessDenied.text = lStrings.sAccessDenied
txtLoggedOut.text = lStrings.sLoggedOut
txtBadPassword.text = lStrings.sBadPassword
txtNick.text = lStrings.sNick
txtDirectoryError.text = lStrings.sDirectoryError
txtEnterPassword.text = lStrings.sEnterPassword
txtEnterUsername.text = lStrings.sEnterUsername
txtShuttingDown.text = lStrings.sShuttingDown
txtShuttingDownNow.text = lStrings.sShuttingDownNow
cboKillDelay.Clear
cboPasswordRetries.Clear
For i = 1 To 200
    If i = 1 Then
        cboPasswordRetries.AddItem i & " retry"
    Else
        cboPasswordRetries.AddItem i & " retries"
    End If
Next i
For i = 1 To 300
    If i = 1 Then
        cboKillDelay.AddItem i & " second"
    Else
        cboKillDelay.AddItem i & " seconds"
    End If
Next i
If lSettings.sShowUsersInDir = True Then chkShowUsersInDir.Value = 1
If lSettings.sFolderChat = True Then chkFolderChat.Value = 1
If lSettings.sColoredWindows = True Then chkColoredWindows.Value = 1
If lSettings.sTelnetOnStartup = True Then chkTelnetOnStartup.Value = 1
If lSettings.sLoadServerOnStartup = True Then chkLoadServerOnStartup.Value = 1
If lSettings.sConnectedOnStartup = True Then chkConnectedOnStartup.Value = 1
If lSettings.sMonitorChat = True Then chkMonitorChat.Value = 1
If lSettings.sAcceptingConnections = True Then chkAcceptingConnections.Value = 1
If lSettings.sShareOnlyHomeDir = True Then chkShareOnlyHomeDir.Value = 1
If lSettings.sEchoText = True Then chkEchoText.Value = 1
If lSettings.sWhiteBlack = True Then chkWhiteBlack.Value = 1
If lSettings.sFSTelnet = True Then chkFSTelnet.Value = 1
If lSettings.sAcceptingNewUsers = True Then chkAcceptingNewUsers.Value = 1
If lSettings.sClosePause = True Then
    cboKillDelay.Enabled = True
    chkClosePause.Value = 1
Else
    cboKillDelay.Enabled = False
    chkClosePause.Value = 0
End If
txtEMail.text = lSettings.sEMail
txtServerName.text = lSettings.sServerName
txtPath.text = lSettings.sRootDir
txtAdministrator.text = lSettings.sAdministrator
cboPasswordRetries.ListIndex = lSettings.sMaxPasswordRetries - 1
cboKillDelay.ListIndex = lSettings.sKillDelay - 1
End Sub

Private Sub optSettings_Click(Index As Integer)
On Local Error Resume Next
Dim i As Integer
For i = 0 To fraSettings.Count
    fraSettings(i).Visible = False
Next i
fraSettings(Index).Visible = True
End Sub

Private Sub OsenXPButton1_Click()

End Sub

Private Sub OsenXPButton2_Click()

End Sub
