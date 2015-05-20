'	Uncomments specific code
'	Useful for batching multiple files which may have commented code.
'	The reverse would also be easy
'	(c) J. Ditzel, May 2015
'
'	Usage CMD line:
'	ReplaceText FileName
'
'	Example 
'			ReplaceText "C:\folder\subfolder\file.ext"
'

Dim oArgs, strOldText, strNewText, oFSO, oFile, strText

Set oArgs = WScript.Arguments

strOpenComment = "<!-- "
strEndComment = " -->"

'Set the code to be uncommented here
'	Use chr(34) to escape double quotes
strCode = "<font face=" & chr(34) & "Segoe UI" & chr(34) & ">"

strCommented = strOpenComment & strCode & strEndComment
strUncommented = strFont

Set oFSO = CreateObject("Scripting.FileSystemObject") 

If Not oFSO.FileExists(oArgs(0)) Then
    WScript.Echo "Specified file does not exist."
Else
    Set oFile = oFSO.OpenTextFile(oArgs(0), 1)
    strText = oFile.ReadAll
    oFile.Close

    strText = Replace(strText, strCommented, strUncommented, 1, -1, intCaseSensitive) 

    Set oFile = oFSO.OpenTextFile(oArgs(0), 2)
    oFile.WriteLine strText
    oFile.Close
End If


