VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.Form frmTelnet 
   Caption         =   "Telos - Telnet"
   ClientHeight    =   3450
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4155
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTelnet.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3450
   ScaleWidth      =   4155
   Begin RichTextLib.RichTextBox txtIncoming 
      Height          =   2295
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   4048
      _Version        =   393217
      BackColor       =   16777215
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      MousePointer    =   1
      Appearance      =   0
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmTelnet.frx":0CCA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSWinsockLib.Winsock wskTelnet 
      Left            =   0
      Top             =   2880
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox txtOutgoing 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   2400
      Width           =   3975
   End
End
Attribute VB_Name = "frmTelnet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub AddText(lText As String, lColor As ColorConstants)
On Local Error Resume Next
txtIncoming.SelStart = Len(txtIncoming)
If lSettings.sWhiteBlack = True Then
    txtIncoming.SelColor = vbWhite
Else
    txtIncoming.SelColor = lColor
End If
If Len(txtIncoming.SelText) <> 0 Then
    txtIncoming.SelText = txtIncoming.SelText & vbCrLf & lText
Else
    
    txtIncoming.SelText = lText
End If
End Sub

Private Sub Form_Load()
On Local Error Resume Next
Me.Tag = "Telnet"
If Len(lSettings.sClientConnectionIp) <> 0 Then
    wskTelnet.Close
    wskTelnet.Connect lSettings.sClientConnectionIp, "23"
    lSettings.sClientConnectionIp = ""
Else
    wskTelnet.Close
    wskTelnet.Connect lSettings.sIpAddress, "23"
End If
If lSettings.sWhiteBlack = True Then
    txtIncoming.BackColor = vbBlack
    txtOutgoing.BackColor = vbBlack
    txtOutgoing.ForeColor = vbWhite
End If
End Sub

Private Sub Form_Resize()
On Local Error Resume Next
If Me.ScaleHeight > 270 Then
    txtIncoming.Height = (Me.ScaleHeight - txtOutgoing.Height)
    txtIncoming.Width = Me.ScaleWidth
    txtOutgoing.Top = (Me.ScaleHeight - txtOutgoing.Height)
    txtOutgoing.Width = Me.ScaleWidth
End If
End Sub

Private Sub txtIncoming_Change()
On Local Error Resume Next
txtIncoming.SelStart = Len(txtIncoming)
End Sub

Private Sub txtOutgoing_GotFocus()
On Local Error Resume Next
txtIncoming.SelStart = Len(txtIncoming)
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
End If
If Err.Number <> 0 Then ErrorHandle Err.Number, Err.Description, "Private Sub txtOutgoing_KeyDown(KeyCode As Integer, Shift As Integer)"
End Sub

Private Sub txtOutgoing_KeyPress(KeyAscii As Integer)
On Local Error Resume Next
Dim msg As String
If KeyAscii = 27 Then
    Me.WindowState = vbMinimized
End If
If KeyAscii = 13 Then
    If Left(LCase(txtOutgoing.text), 9) = "/connect " Then
        wskTelnet.Close
        msg = Right(txtOutgoing.text, Len(txtOutgoing.text) - 9)
        AddText "Connecting to " & msg & vbCrLf, vbBlue
        wskTelnet.Connect msg, 23
        KeyAscii = 0
        txtOutgoing.text = ""
        Exit Sub
    ElseIf Left(LCase(txtOutgoing.text), 10) = "/backcolor" Or Left(LCase(txtOutgoing.text), 5) = "/back" Then
        msg = GetRnd(14700000)
        txtOutgoing.BackColor = msg
        txtIncoming.BackColor = msg
        KeyAscii = 0
    ElseIf Left(LCase(txtOutgoing.text), 12) = "/disconnect " Then
        wskTelnet.Close
    Else
        If Len(txtOutgoing.text) = 0 Then
            KeyAscii = 0
            Exit Sub
        End If
        msg = txtOutgoing.text & vbCrLf
        If wskTelnet.State = sckConnected Then
            wskTelnet.SendData msg
            KeyAscii = 0
            txtOutgoing.text = ""
            Exit Sub
        Else
            KeyAscii = 0
            txtOutgoing.text = ""
            AddText "Unable to send data" & vbCrLf, vbBlue
            Exit Sub
        End If
    End If
End If
End Sub

Private Sub wskTelnet_Close()
On Local Error Resume Next
If lSettings.sFSTelnet = True Then
    Unload mdiTelos.ActiveForm
    Unload frmFSTelnet
Else
    AddText vbCrLf & "Socket Disconnected (" & wskTelnet.RemoteHostIP & ")" & vbCrLf, vbBlue
End If
End Sub

Private Sub wskTelnet_Connect()
On Local Error Resume Next
AddText "Socket Connected (" & wskTelnet.RemoteHostIP & ")" & vbCrLf, vbBlue
End Sub

Private Sub wskTelnet_ConnectionRequest(ByVal requestID As Long)
On Local Error Resume Next
AddText "A connection request was made and declined. Id: " & requestID & vbCrLf, vbBlue
End Sub

Private Sub wskTelnet_DataArrival(ByVal bytesTotal As Long)
On Local Error Resume Next
Dim text As String
wskTelnet.GetData text, vbString
If lSettings.sWhiteBlack = True Then
    txtIncoming.SelColor = vbWhite
Else
    txtIncoming.SelColor = vbBlack
End If
If lSettings.sFSTelnet = True Then
    frmFSTelnet.AddText text, vbWhite
Else
    txtIncoming.SelText = txtIncoming.SelText & text
End If
End Sub

Private Sub wskTelnet_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
On Local Error Resume Next
AddText "Error: " & Description & vbCrLf, vbBlue
End Sub
