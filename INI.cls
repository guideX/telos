VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "INI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private mvarAppName As String
Private mvarDefaultReturn As String

Public Property Let DefaultReturn(ByVal vData As String)
    mvarDefaultReturn = vData
End Property

Public Property Get DefaultReturn() As String
    DefaultReturn = mvarDefaultReturn
End Property

Public Property Let AppName(ByVal vData As String)
    mvarAppName = vData
End Property

Public Property Get AppName() As String
    AppName = mvarAppName
End Property

Public Sub WPPS(KeyName As String, txt As String, FileName As String)
WritePrivateProfileString mvarAppName, KeyName, txt, FileName
End Sub

Public Function GPPS(KeyName As String, FileName As String) As String
Dim str As String * 500
GetPrivateProfileString mvarAppName, KeyName, mvarDefaultReturn, str, Len(str), FileName
GPPS = Trim(str)
End Function
