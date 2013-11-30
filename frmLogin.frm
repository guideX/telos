VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Telos - Login"
   ClientHeight    =   1320
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5370
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1320
   ScaleWidth      =   5370
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4080
      TabIndex        =   5
      Top             =   840
      Width           =   855
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3120
      TabIndex        =   4
      Top             =   840
      Width           =   855
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      Left            =   1680
      TabIndex        =   3
      Top             =   480
      Width           =   3255
   End
   Begin VB.TextBox txtNickname 
      Height          =   285
      Left            =   1680
      TabIndex        =   2
      Top             =   120
      Width           =   3255
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter your username/pass to login"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   2775
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      Height          =   255
      Left            =   720
      TabIndex        =   1
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Nickname:"
      Height          =   255
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmLogin.frx":0000
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
On Local Error Resume Next
Dim i As Integer
If Len(txtNickname.text) <> 0 And Len(txtPassword.text) <> 0 Then
    For i = 1 To mdiTelos.Count
        
    Next i
End If
Unload Me
End Sub
