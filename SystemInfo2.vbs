' Gets Username, ComputerName and UserDomain


Set objNet = CreateObject("WScript.NetWork") 

Set objExcel = CreateObject("Excel.Application")
strAlias= objExcel.UserName
objExcel.Quit

Set objWord = CreateObject("Word.Application")
strInitials = objWord.UserInitials
objWord.Quit

strComputer = objNet.ComputerName
strUser = objNet.UserName

isAdministrator = false




Set objGroup = GetObject("WinNT://" & strComputer & "/Administrators")
For Each objUser in objGroup.Members
    If objUser.Name = strUser Then
        isAdministrator = true        
    End If
Next

If isAdministrator Then
    strAdmin = strUser & " is a local administrator on " & strComputer
Else
    strAdmin = strUser & " is not a local administrator on " & strComputer
End If

If IsAdmin Then output = output &vbnewline & "The User is an Administrator on this Session"
If NOT IsAdmin Then output = output &vbnewline & "The User is NOT an Administrator on this Session"

strDate = FormatDateTime(Now(), 1)
strTime = FormatDateTime(Now(), 4)

strInfo = strDate & vbCRLF & strTime & vbCRLF & "User Name: " & objNet.UserName & vbCRLF & _
                "Domain Name: " & objNet.UserDomain & vbCRLF &_
				"Computer Name: " & objNet.ComputerName & vbCRLF & _
				"Computer Serial Number: " & mid(objNet.ComputerName, 5, 1000) & vbCRLF &_
				"User Alias: " & strAlias & vbCRLF &_
				"User Initials: " & strInitials & vbCRLF &_
				strAdmin & vbCRLF &_ 
				output

'MsgBox strInfo

''Writes info to txt file on desktop
'Set objFileToWrite = CreateObject("Scripting.FileSystemObject").OpenTextFile("C:\Users\" & objNet.UserName & "." & objNet.UserDomain & "\Desktop\SystemInfo.txt",2,true)
'objFileToWrite.WriteLine(strInfo)
'objFileToWrite.Close
'Set objFileToWrite = Nothing



'Writes info to txt file in script's path

'Opens File Containing Script
Set objShell = CreateObject("Wscript.Shell")

strPath = Wscript.ScriptFullName

Set objFSO = CreateObject("Scripting.FileSystemObject")

Set objFile = objFSO.GetFile(strPath)
strFolder = objFSO.GetParentFolderName(objFile)
 
Set objFileToWrite = CreateObject("Scripting.FileSystemObject").OpenTextFile(strFolder & "\SystemInfo.txt",2,true)
objFileToWrite.WriteLine(strInfo)
objFileToWrite.Close
Set objFileToWrite = Nothing

Function IsAdmin()
  IsAdmin = False
  On Error Resume Next
  key = CreateObject("WScript.Shell").RegRead("HKEY_USERS\S-1-5-19\Environment\TEMP")
  If err.number = 0 Then IsAdmin = True
End Function