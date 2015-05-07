'This file is meant to be run from the command line
'It needs at least one argument in the form of the path of the html file
'


Dim htmlfile
Dim fso, ts, OutApp, OutMail

Set objArgs = Wscript.Arguments
On Error Resume Next

htmlfile = objArgs(0)


Set fso = CreateObject("Scripting.FileSystemObject")
Set ts = fso.GetFile(htmlfile).OpenAsTextStream(1, -2)
HTMLtoString = ts.readall
	ts.close
HTMLtoString = Replace(HTMLtoString, "align=center x:publishsource=", "align=left x:publishsource=")

Set OutApp = CreateObject("Outlook.Application")
Set OutMail = OutApp.CreateItem(0)

    'On Error Resume Next
    With OutMail
		.To = "test@somecompany.com"
		.CC = ""
		.BCC = ""
		.Subject = "This is the Subject line"
		.HTMLBody = HTMLtoString
		'.Send   'or use .Display
		.Display
    End With
Set ts = Nothing
Set fso = Nothing
Set OutMail = Nothing
Set OutApp = Nothing

