VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.Form frmFSTelnet 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2880
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3570
   LinkTopic       =   "Form1"
   ScaleHeight     =   2880
   ScaleWidth      =   3570
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Tag             =   "FS"
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtOutgoing 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   2520
      Width           =   3255
   End
   Begin RichTextLib.RichTextBox txtIncoming 
      Height          =   2415
      Left            =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   4260
      _Version        =   393217
      BackColor       =   0
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmFullScreenChat.frx":0000
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
End
Attribute VB_Name = "frmFSTelnet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub AddText(lText As String, lColor As ColorConstants)
On Local Error Resume Next
txtIncoming.SelColor = lColor
txtIncoming.SelText = txtIncoming.SelText & vbCrLf & lText
End Sub

Private Sub Form_Resize()
On Local Error Resume Next
txtIncoming.Width = Me.ScaleWidth + 20
txtIncoming.Height = (Me.ScaleHeight - txtOutgoing.Height)
txtOutgoing.Top = (Me.ScaleHeight - txtOutgoing.Height)
txtOutgoing.Width = Me.ScaleWidth + 20
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Local Error Resume Next
mdiTelos.WindowState = vbNormal
End Sub

Private Sub txtIncoming_GotFocus()
On Local Error Resume Next
txtOutgoing.SetFocus
End Sub

Private Sub txtOutgoing_KeyPress(KeyAscii As Integer)
On Local Error Resume Next
Dim msg As String
If KeyAscii = 13 Then
    If Len(txtOutgoing.text) = 0 Then
        KeyAscii = 0
        Exit Sub
    End If
    If mdiTelos.ActiveForm.wskTelnet.State = sckConnected Then
        mdiTelos.ActiveForm.wskTelnet.SendData txtOutgoing.text & vbCrLf
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
End Sub
