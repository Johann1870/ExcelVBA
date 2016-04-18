Sub FormatSubStrings()
'(c) J. Ditzel, May 2015
' Conditionally formats substrings in a range in Excel


Dim intA, intB, intC, intD, intE As Integer
Dim strA, strB, strC, strD, strE As String
Dim rng As Range

'Set range here
Set rng = Sheet1.Range("A2:A195")

'Set substrings here
strA = "xxxxx" 'Sets 5 char length
strB = "GLUTEN FREE"
strC = "SM"
strD = "LG"
strE = "XLG"

'Sets intA to intE to zero. These are the starting positions in the string.
intA = 0
intB = 0
intC = 0
intD = 0
intE = 0

'Set your range here
For Each c In rng

'Sets the starting positions in the string
    intA = 1 'This one is simply set to start at the beginning of the string. You could also set 5 chars from the end by using intA = LEN(strA) - 5
    intB = InStr(c, strB)
    intC = InStr(c, strC)
    intD = InStr(c, strD)
    intE = InStr(c, strE)
    
    'strA
    If intA > 0 Then
        With c.Characters(Start:=1, Length:=Len(strA)).Font
            .FontStyle = "Bold"
        End With
    End If
    
     'strB
    If intB > 0 Then
        With c.Characters(Start:=intB, Length:=Len(strB)).Font
            .FontStyle = "Bold"
            .Color = RGB(255, 128, 128)
        End With
    End If
    
    'strC
    If intC > 0 Then
        With c.Characters(Start:=intC, Length:=Len(strC)).Font
            .FontStyle = "Bold"
            .Color = RGB(255, 128, 0)
        End With
    End If
    
    'strD
    If intD > 0 Then
        With c.Characters(Start:=intD, Length:=Len(strD)).Font
            .FontStyle = "Bold"
            .Color = RGB(0, 200, 0)
        End With
    End If
    
    'strE
    If intE > 0 Then
        With c.Characters(Start:=intE, Length:=Len(strE)).Font
            .FontStyle = "Bold"
            .Color = RGB(0, 128, 255)
        End With
    End If
    
    
Next
End Sub

