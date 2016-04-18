Attribute VB_Name = "FormatChemicals"
Sub FormatChemicals()
'(c) J. Ditzel, May 2015
'Formats Chemical subscripts in Excel


Dim intA As Integer
Dim strA As String
Dim rng As Range

'Set range here
Set rng = Sheet1.Range("A1:A3")

intA = 0

'Set your range here
For Each c In rng
strA = c.Value
'Sets the starting positions in the string
    For i = 1 To Len(strA)
        If IsNumeric(Mid(c.Value, i, 1)) Then
            
            With c.Characters(Start:=i, Length:=1).Font
                .Subscript = True
            End With
        End If
    Next
Next

End Sub

Sub UnFormatChemicals()
'(c) J. Ditzel, May 2015
'Formats Chemical subscripts in Excel


Dim intA As Integer
Dim strA As String
Dim rng As Range

'Set range here
Set rng = Sheet1.Range("A1:A3")

For Each c In rng
    With c.Font
        .Superscript = False
        .Subscript = False
    End With
Next

End Sub
