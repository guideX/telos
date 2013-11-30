VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.Form frmTxtEditor 
   Caption         =   "Telos - Text Editor"
   ClientHeight    =   1680
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   2145
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTxtEditor.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   1680
   ScaleWidth      =   2145
   Tag             =   "Text Editor"
   Begin RichTextLib.RichTextBox txtIncoming 
      Height          =   1335
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   2355
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      ScrollBars      =   3
      Appearance      =   0
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmTxtEditor.frx":0CCA
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
   Begin VB.CheckBox chkModifiedSinceLastSave 
      Caption         =   "Modified since last save"
      Height          =   195
      Left            =   0
      TabIndex        =   1
      Top             =   1440
      Width           =   2415
   End
End
Attribute VB_Name = "frmTxtEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
On Local Error Resume Next
Dim msg As String
If lSettings.sColoredWindows = True Then
    msg = GetRnd(14700000)
    txtIncoming.BackColor = msg
End If

End Sub

Private Sub Form_Resize()
On Local Error Resume Next
txtIncoming.Width = Me.ScaleWidth
txtIncoming.Height = Me.ScaleHeight
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Local Error Resume Next
Dim m As VbMsgBoxResult
If Len(txtIncoming.text) <> 0 Then
    If chkModifiedSinceLastSave.Value = 1 Then
        m = MsgBox("You have not yet saved this text, do you wish to save it now?", vbYesNoCancel + vbQuestion, App.Title)
        If m = vbYes Then
            mdiTelos.ActiveateSave
            DoEvents
            Unload Me
        ElseIf m = vbNo Then
        ElseIf m = vbCancel Then
            Cancel = 1
        End If
    End If
End If
End Sub

Private Sub txtIncoming_Change()
On Local Error Resume Next
chkModifiedSinceLastSave.Value = 1
End Sub
