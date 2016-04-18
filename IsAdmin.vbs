' (c) 2014
' J. Ditzel


If IsAdmin Then output = output &vbnewline & "The User is an Administrator on this Session"
If NOT IsAdmin Then output = output &vbnewline & "The User is NOT an Administrator on this Session"

wscript.echo output

Function IsAdmin()
  IsAdmin = False
  On Error Resume Next
  key = CreateObject("WScript.Shell").RegRead("HKEY_USERS\S-1-5-19\Environment\TEMP")
  If err.number = 0 Then IsAdmin = True
End Function