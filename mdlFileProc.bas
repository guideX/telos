Attribute VB_Name = "mdlFileProc"
Option Explicit

Public Function GetFileTitle(lFilename As String) As String
On Local Error Resume Next
If Len(lFilename) <> 0 Then
AGAIN:
    If InStr(lFilename, "\") Then
        lFilename = Right(lFilename, Len(lFilename) - InStr(lFilename, "\"))
        If InStr(lFilename, "\") Then
            GoTo AGAIN
        Else
            GetFileTitle = lFilename
        End If
    Else
        GetFileTitle = lFilename
    End If
End If
End Function


