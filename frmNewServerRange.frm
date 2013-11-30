VERSION 5.00
Begin VB.Form frmNewServerRange 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Telos - Port Range"
   ClientHeight    =   1335
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3855
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmNewServerRange.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1335
   ScaleWidth      =   3855
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtEnd 
      Height          =   285
      Left            =   1920
      TabIndex        =   5
      Text            =   "28"
      Top             =   480
      Width           =   1815
   End
   Begin VB.TextBox txtStart 
      Height          =   285
      Left            =   1920
      TabIndex        =   3
      Text            =   "24"
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   840
      Width           =   855
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1920
      TabIndex        =   0
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Ending Port Range:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Starting Port Range:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "frmNewServerRange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
On Local Error Resume Next
Dim i As Long, m As Form, b As Boolean, msg As String
Me.Visible = False
b = lSettings.sConnectedOnStartup
lSettings.sConnectedOnStartup = False
For i = Int(txtStart.text) To Int(txtEnd.text)
    Set m = New frmTelos
    ActivateServerToggle True, m, i
    If lSettings.sColoredWindows = True Then
        msg = GetRnd(14700000)
        m.lstUsers.BackColor = msg
        m.console.BackColor = msg
        m.txtOutgoing.BackColor = msg
    End If
Next i
lSettings.sConnectedOnStartup = b
Unload Me
End Sub
