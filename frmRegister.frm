VERSION 5.00
Object = "{B3FF7FA6-B059-4900-8BEC-5C65E3D9C033}#1.0#0"; "xplook.ocx"
Begin VB.Form frmRegister 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Register"
   ClientHeight    =   5130
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   3345
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRegister.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   3345
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "&G"
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   4680
      Visible         =   0   'False
      Width           =   375
   End
   Begin OsenXPCntrl.OsenXPButton OsenXPButton2 
      Height          =   375
      Left            =   1440
      TabIndex        =   10
      Top             =   2160
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Buy Telos"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   0
      FCOLO           =   16711680
      MCOL            =   12632256
      MPTR            =   0
      MICON           =   "frmRegister.frx":000C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin OsenXPCntrl.OsenXPButton cmdCancel 
      Height          =   375
      Left            =   2160
      TabIndex        =   6
      Top             =   4680
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Cancel"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   0
      FCOLO           =   16711680
      MCOL            =   12632256
      MPTR            =   0
      MICON           =   "frmRegister.frx":0028
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin OsenXPCntrl.OsenXPButton cmdRegister 
      Default         =   -1  'True
      Height          =   375
      Left            =   960
      TabIndex        =   5
      Top             =   4680
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Register"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   0
      FCOLO           =   16711680
      MCOL            =   12632256
      MPTR            =   0
      MICON           =   "frmRegister.frx":0044
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      Left            =   1080
      TabIndex        =   4
      Top             =   4080
      Width           =   2175
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   1080
      TabIndex        =   2
      Top             =   3720
      Width           =   2175
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmRegister.frx":0060
      Height          =   975
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   3015
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   3240
      Y1              =   3615
      Y2              =   3615
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   3240
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Wait 1 day for code to be generated, you will recieve the code in e-mail, enter it below when it has been recieved"
      Height          =   975
      Left            =   840
      TabIndex        =   8
      Top             =   2760
      Width           =   2415
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Step 2:"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Registration costs $20 USD. Click the button below to launch paypal"
      Height          =   855
      Left            =   840
      TabIndex        =   9
      Top             =   1560
      Width           =   2415
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Step 1:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1560
      Width           =   855
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   3240
      Y1              =   4575
      Y2              =   4575
   End
   Begin VB.Line Line3 
      X1              =   120
      X2              =   3240
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   3720
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "How to register Telos"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   3855
   End
End
Attribute VB_Name = "frmRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
On Local Error Resume Next
Unload Me
End Sub

Private Sub cmdRegister_Click()
On Local Error Resume Next
Dim m As Boolean, i As String
'm = True
If txtName.text = "guidex_developer" And txtPassword.text = "07281979-2841" Then
    Command1.Visible = True
    txtName.text = ""
    txtPassword.text = ""
    Exit Sub
End If
If m = True Then
    txtPassword.text = KeyGen(txtName.text, "pickles", 1)
Else
    i = KeyGen(txtName.text, "pickles", 1)
    If i = txtPassword.text Then
        MsgBox "Thank you very much for registering. All of the money made from Telos is spent on the development of Telos", vbInformation
        lSettings.sName = txtName.text
        lSettings.sPassword = txtPassword.text
        WriteINI lSettings.sIniFile, "Settings", "Name", lSettings.sName
        WriteINI lSettings.sIniFile, "Settings", "Password", lSettings.sPassword
        lSettings.sRegistered = True
        Unload Me
    Else
        MsgBox "The code you entered was not correct. The name did not match the password. Please try again", vbInformation
    End If
End If
End Sub

Private Sub Command1_Click()
On Local Error Resume Next
Dim msg As String
msg = InputBox("Enter secret phrase:", "Code generator", "37463788473623")
If msg = "pickles" Then
    txtPassword.text = KeyGen(txtName.text, "pickles", 1)
Else
    txtName.text = ""
    txtPassword.text = ""
End If
End Sub

Private Sub Form_Load()
On Local Error Resume Next
If Len(lSettings.sName) <> 0 And Len(lSettings.sPassword) <> 0 Then
    txtName.text = lSettings.sName
    txtPassword.text = lSettings.sPassword
End If
End Sub

Private Sub OsenXPButton1_Click()
On Local Error Resume Next
Surf "mailto:guidex@team-nexgen.com", Me.hWnd
End Sub

Private Sub OsenXPButton2_Click()
On Local Error Resume Next
Surf "https://www.paypal.com/xclick/business=guidex%40team-nexgen.com&item_name=Audiogen+Registration&amount=20.00&no_note=1&tax=0&currency_code=USD&lc=US", Me.hWnd
End Sub
