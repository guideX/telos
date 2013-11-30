VERSION 5.00
Begin VB.Form frmRemoveUser 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Telos - Remove User"
   ClientHeight    =   3975
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3255
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRemoveUser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   3255
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Remove"
      Default         =   -1  'True
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   3480
      Width           =   1095
   End
   Begin VB.ListBox lstUsers 
      Height          =   3180
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "frmRemoveUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type gUser
    uAdministrator As Boolean
    uNickname As String
End Type
Private Type gUsers
    uUsers(300) As gUser
    uCount As Integer
End Type
Dim lUsers As gUser

Private Sub cmdRemove_Click()
On Local Error Resume Next
Dim i As Integer

End Sub

Private Sub Form_Load()
On Local Error Resume Next
Dim msg As String, m As String
msg = ReadFile(App.Path & "\files\users.txt")
If Len(msg) <> "" Then
AGAIN:
    If InStr(msg, Chr(13)) Then
'        MsgBox msg
    End If
End If
End Sub
