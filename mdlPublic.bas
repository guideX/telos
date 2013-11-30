Attribute VB_Name = "mdlPublic"
Option Explicit
Public Crt As String
Public lWinsockCount As Long
Private Type gUser
    uName As String
    uInput As String
    uHost As String
    uWhat As String
    uSock As Integer
    uOperator As Boolean
    uWTF As Boolean
    uRelDir As String
    uPassword As String
    uPasswordRetries As Integer
End Type
Private Type gUsers
    uUser(1000) As gUser
    uCount As Integer
End Type
Private Type gFontProporties
    fFace As String
    fSize As Integer
    fBold As Boolean
    fItalic As Boolean
End Type
Private Type gStrings
    sBadPassword As String
    sLoggedOut As String
    sAccessDenied As String
    sNick As String
    sDirectoryError As String
    sEnterUsername As String
    sEnterPassword As String
    sShuttingDown As String
    sShuttingDownNow As String
End Type
Private Type gSettings
    sFolderCount As Integer
    sWhiteBlack As Boolean
    sShareOnlyHomeDir As Boolean
    sLoadServerOnStartup As Boolean
    sMonitorChat As Boolean
    sGlobalClose As Boolean
    sIpAddress As String
    sKillDelay As Long
    sFontProporties As gFontProporties
    sTelnetOnStartup As Boolean
    sServerName As String
    sRootDir As String
    sRootDirLen As Integer
    sIniFile As String
    sAdministrator As String
    sConnectedOnStartup As Boolean
    sAcceptingConnections As Boolean
    sEchoText As Boolean
    sClientConnectionIp As String
    sColoredWindows As Boolean
    sEMail As String
    sFolderChat As Boolean
    sFSTelnet As Boolean
    sMaxUsers As Long
    sName As String
    sPassword As String
    sRegistered As Boolean
    sClosePause As Boolean
    sAcceptingNewUsers As Boolean
    sMaxPasswordRetries As Integer
    sShowUsersInDir As Boolean
End Type
Global lStrings As gStrings
Global lSettings As gSettings
Global lUsers As gUsers
