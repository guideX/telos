Attribute VB_Name = "mdlNexINCLUSIVE"
Option Explicit
Declare Function EnumWindows Lib "user32.dll" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Declare Function GetWindowTextLength Lib "user32.dll" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Declare Function GetWindowText Lib "user32.dll" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal nMaxCount As Long) As Long
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Declare Function CloseWindow Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Public Type POINTAPI
    x As Long
    y As Long
End Type
Public Type msg
    hwnd As Long
    message As Long
    wParam As Long
    lParam As Long
    time As Long
    pt As POINTAPI
End Type

Public Sub IncludeTelnet(lForm As Form)
On Local Error Resume Next
Dim i As Integer, h As Long
ListAllWindows lForm
For i = 0 To lForm.lstProcess.ListCount
    If Left(LCase(lForm.lstProcess.List(i)), 6) = "telnet" Then
        lForm.lstProcess.ListIndex = i
        h = lForm.lstProcess.ItemData(lForm.lstProcess.ListIndex)
        SetParent h, mdiTelos.hwnd
        CloseWindow h
        mdiTelos.SetFocus
    End If
Next i
End Sub

Public Function EnumWindowsProc(ByVal hwnd As Long, ByVal lParam As Long, lForm As Form) As Long
Dim SLength As Long, Buffer As String
Dim RetVal As Long, progcap As String
Dim onlist As Boolean
onlist = False
Static WinNum As Integer
WinNum = WinNum + 1
SLength = GetWindowTextLength(hwnd) + 1
If SLength > 1 Then
    Buffer = Space(SLength)
    RetVal = GetWindowText(hwnd, Buffer, SLength)
    progcap = Left(Buffer, SLength - 1)
    lForm.lstProcess.AddItem progcap
    lForm.lstProcess.ItemData(lForm.lstProcess.NewIndex) = hwnd
End If
EnumWindowsProc = 1

End Function
