Attribute VB_Name = "mdlConsole"
Sub StartConsole()
ConsolePrint "Telos " & App.Major & "." & App.Minor
End Sub

Public Sub ConsolePrint(szOut As String)
Trim (szOut)
frmTelos.console.AddItem szOut
frmTelos.console.Selected(frmTelos.console.ListCount - 1) = True
End Sub
