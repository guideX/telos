VERSION 5.00
Begin VB.Form frmUserInfo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Telos - User information"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4575
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmUserInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4575
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Default         =   -1  'True
      Height          =   375
      Left            =   3360
      TabIndex        =   15
      Top             =   2640
      Width           =   1095
   End
   Begin VB.TextBox txtRelDIR 
      Height          =   285
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   2280
      Width           =   1575
   End
   Begin VB.TextBox txtWTF 
      Height          =   285
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   1920
      Width           =   1575
   End
   Begin VB.TextBox txtSuperUser 
      Height          =   285
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   1560
      Width           =   1575
   End
   Begin VB.TextBox txtSock 
      Height          =   285
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   1200
      Width           =   1575
   End
   Begin VB.TextBox txtWhat 
      Height          =   285
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   840
      Width           =   1575
   End
   Begin VB.TextBox txtHost 
      Height          =   285
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   480
      Width           =   1575
   End
   Begin VB.TextBox txtInput 
      Height          =   2445
      Left            =   2760
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   120
      Width           =   1695
   End
   Begin VB.TextBox txtNickname 
      Height          =   285
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label8 
      Caption         =   "RelDir:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label Label7 
      Caption         =   "WTF:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Label Label6 
      Caption         =   "Super User:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Label Label5 
      Caption         =   "Sock:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "What:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Host:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Nickname:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmUserInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
Unload Me
End Sub
