Attribute VB_Name = "mdlCrypt"
Option Explicit

Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public softwareCode As String
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const HTCAPTION = 2

Function KeyGen(kName As String, kPass As String, kType As Integer) As String
On Error Resume Next
Dim cTable(512) As Integer
Dim nKeys(16) As Integer
Dim s0(512) As Integer
Dim nArray(16) As Integer
Dim pArray(16) As Integer
Dim n As Integer
Dim nPtr As Integer
Dim cPtr As Integer
Dim cFlip As Boolean
Dim sIni As Integer
Dim temp As Integer
Dim rtn As Integer
Dim gKey As String
Dim nLen As Integer
Dim pLen As Integer
Dim kPtr As Integer
Dim sPtr As Integer
Dim nOffset As Integer
Dim pOffset As Integer
Dim tOffset As Integer
Const nXor As Integer = 18
Const pXor As Integer = 25
Const cLw As Integer = 65
Const nLw As Integer = 48
Const sOffset As Integer = 0
nLen = Len(kName)
pLen = Len(kPass)
nKeys(1) = 52
nKeys(2) = 69
nKeys(3) = 149
nKeys(4) = 37
nKeys(5) = 403
nKeys(6) = 20
nKeys(7) = 58
nKeys(8) = 29
nKeys(9) = 123
nKeys(10) = 84
nKeys(11) = 201
nKeys(12) = 202
nKeys(13) = 34
nKeys(14) = 38
nKeys(15) = 73
nKeys(16) = 30
sIni = 0
For n = 0 To 512
    s0(n) = n
Next n
For n = 0 To 512
    sIni = (sOffset + sIni + n) Mod 256
    temp = s0(n)
    s0(n) = s0(sIni)
    s0(sIni) = temp
Next n
If kType = 1 Then
    nPtr = 0
    For n = 0 To 512
        cTable(s0(n)) = (nLw + (nPtr))
        nPtr = nPtr + 1
        If nPtr = 10 Then nPtr = 0
    Next n
    gKey = String(16, " ")
ElseIf kType = 2 Then
    nPtr = 0
    cPtr = 0
    cFlip = False
    For n = 0 To 512
        If cFlip Then
            cTable(s0(n)) = (nLw + nPtr)
            nPtr = nPtr + 1
            If nPtr = 10 Then nPtr = 0
            cFlip = False
        Else
            cTable(s0(n)) = (cLw + cPtr)
            cPtr = cPtr + 1
            If cPtr = 26 Then cPtr = 0
            cFlip = True
        End If
    Next n
    gKey = String(16, " ")
Else
    gKey = String(19, " ")
End If
kPtr = 1
For n = 1 To nLen
  nArray(kPtr) = nArray(kPtr) + Asc(Mid(kName, n, 1)) Xor nXor
  nOffset = nOffset + nArray(kPtr)
  kPtr = kPtr + 1
    If kPtr = 9 Then kPtr = 1
Next n
For n = 1 To pLen
  pArray(kPtr) = pArray(kPtr) + Asc(Mid(kPass, n, 1)) Xor pXor
  pOffset = pOffset + pArray(kPtr)
  kPtr = kPtr + 1
    If kPtr = 9 Then kPtr = 1
Next n
tOffset = (nOffset + pOffset) Mod 512
kPtr = 1
sPtr = 1
For n = 1 To 16
  pArray(n) = pArray(n) Xor nKeys(n)
  rtn = Abs(((nArray(n) Xor pArray(n)) Mod 512) - tOffset)
  If kType = 3 Then
        If rtn < 16 Then
            Mid(gKey, kPtr, 2) = "0" & Hex(rtn)
        Else
            Mid(gKey, kPtr, 2) = Hex(rtn)
        End If
            If sPtr = 2 And kPtr < 18 Then
                kPtr = kPtr + 1
                Mid(gKey, kPtr + 1, 1) = "-"
            End If
        kPtr = kPtr + 2
        sPtr = sPtr + 1
        If sPtr = 3 Then sPtr = 1
  Else
    Mid(gKey, n, 1) = Chr(cTable(rtn))
  End If
Next
KeyGen = gKey
End Function
