

'Opens File Containing Script
Set objShell = CreateObject("Wscript.Shell")

strPath = Wscript.ScriptFullName

Set objFSO = CreateObject("Scripting.FileSystemObject")

Set objFile = objFSO.GetFile(strPath)
strFolder = objFSO.GetParentFolderName(objFile) 

strPath = "explorer.exe /e," & strFolder
'objShell.Run strPath


'Opens Current Directory
Set objShell2 = CreateObject("Wscript.Shell")
strFolder2 = objShell2.CurrentDirectory

strPath2 = "explorer.exe /e," & strFolder2
'objShell2.Run strPath2


If strFolder = strFolder2 Then
	strEqTest = "These paths are the same."
Else
	strEqTest = "These paths are NOT the same."
End If


MsgBox("The folder path containing the script is '" & strFolder &"'." & vbCRLF & "The current directory is '" & strFolder2 & "'." & vbCRLF & strEqTest)
