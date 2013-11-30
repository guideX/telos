VERSION 5.00
Begin VB.Form frmSplash 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2265
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   3000
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSplash.frx":0000
   ScaleHeight     =   2265
   ScaleWidth      =   3000
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrUnload 
      Interval        =   2000
      Left            =   120
      Top             =   1800
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
AlwaysOnTop Me, True
End Sub

Private Sub tmrUnload_Timer()
On Local Error Resume Next
tmrUnload.Enabled = False
AlwaysOnTop Me, False
Unload Me
End Sub
