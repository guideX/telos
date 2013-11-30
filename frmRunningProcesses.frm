VERSION 5.00
Begin VB.Form frmRunningProcesses 
   Caption         =   "Telos - Running Processes"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6765
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRunningProcesses.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   6765
   Visible         =   0   'False
   Begin VB.ListBox lstProcess 
      Appearance      =   0  'Flat
      Height          =   1785
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3855
   End
End
Attribute VB_Name = "frmRunningProcesses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
lstProcess.Clear
EnumWindows AddressOf EnumWindowsProc, 0
End Sub

Private Sub Form_Resize()
On Local Error Resume Next
lstProcess.Width = (Me.ScaleWidth)
lstProcess.Height = (Me.ScaleHeight)
lstProcess.Top = 0
End Sub
