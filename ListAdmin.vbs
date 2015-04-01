Set objNet = CreateObject("WScript.NetWork") 
strComputer = objNet.ComputerName

strInfo = ""
Set objGroup = GetObject("WinNT://" & strComputer & "/Administrators")
For Each objUser in objGroup.Members
    strInfo = strInfo & objUser.Name & chr(13) & chr(10)

Next

'Opens File Containing Script
Set objShell = CreateObject("Wscript.Shell")

strPath = Wscript.ScriptFullName

Set objFSO = CreateObject("Scripting.FileSystemObject")

Set objFile = objFSO.GetFile(strPath)
strFolder = objFSO.GetParentFolderName(objFile)
 
Set objFileToWrite = CreateObject("Scripting.FileSystemObject").OpenTextFile(strFolder & "\AdminList.txt",2,true)
objFileToWrite.WriteLine(strInfo)
objFileToWrite.Close
Set objFileToWrite = Nothing
