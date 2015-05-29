Attribute VB_Name = "Module1"
Sub BatchRenameFiles()
' (c) J. Ditzel, May 2015
' Renames files listed in Column A to names in Column B
' Use "dir /b > dir.csv" to get a list of the files
'

Dim xDir As String
Dim xFile As String
Dim xRow As Long

'Set the directory here
    xDir = "J:\test\"
    
    xFile = Dir(xDir & Application.PathSeparator & "*")
    Do Until xFile = ""
        xRow = 0
        On Error Resume Next
        xRow = Application.Match(xFile, Range("A:A"), 0)
        If xRow > 0 Then
            Name xDir & Application.PathSeparator & xFile As _
            xDir & Application.PathSeparator & Cells(xRow, "B").Value
        End If
        xFile = Dir
    Loop

End Sub
