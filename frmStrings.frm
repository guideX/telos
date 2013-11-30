VERSION 5.00
Begin VB.Form frmStrings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Telos - String Editor"
   ClientHeight    =   945
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4110
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmStrings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   945
   ScaleWidth      =   4110
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cboStringObjects 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1320
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   110
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "String Objects:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmStrings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
'lStrings.sAccessDenied
'lStrings.sBadPassword
'lStrings.sLoggedOut
'lstrings.sNick
End Sub
