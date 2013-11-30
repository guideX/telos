VERSION 5.00
Begin VB.Form frmChat 
   Caption         =   "Telos - Chat"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmChat.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   Begin VB.TextBox txtOutgoing 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Top             =   2160
      Width           =   4215
   End
   Begin VB.ListBox lstChat 
      Appearance      =   0  'Flat
      Height          =   1950
      IntegralHeight  =   0   'False
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4215
   End
End
Attribute VB_Name = "frmChat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Resize()
On Local Error Resume Next
lstChat.Width = Me.ScaleWidth
lstChat.Height = (Me.ScaleHeight - txtOutgoing.Height)
txtOutgoing.Top = (Me.ScaleHeight - txtOutgoing.Height)
txtOutgoing.Width = Me.ScaleWidth
End Sub

Private Sub txtOutgoing_KeyPress(KeyAscii As Integer)
'Dim i As Integer
'If Len(txtOutgoing.text) <> 0 And KeyAscii = 13 Then
'    For i = 1 To frmTelos.pol.Count
'        If Len(Ac_Name(i)) <> 0 Then
'            frmTelos.Send vbCrLf & "<" & lSettings.sAdministrator & "> " & txtOutgoing.text & vbCrLf, i
'        End If
'    Next i
'    frmChat.lstChat.AddItem "<" & lSettings.sAdministrator & "> " & txtOutgoing.text
'    KeyAscii = 0
'    txtOutgoing.text = ""
'End If
End Sub
