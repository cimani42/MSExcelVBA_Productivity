Sub MakeDateDirectory()
' This code creates a new directory, saved as the current date - Today's date.
' If the current date already exists - in the same format - no new folder will be saved.
' Folders should be saved without have a / or \.

Dim str As String
Dim fol As String
Dim DateFormat As String
Dim Period As String

DateFormat = "yyyy-mm-dd"
'Debug.Print DateFormat

Period = Worksheets("Sheet1").Range("A3").Value
'Debug.Print Period


'Make sure a valid date format is used
If DateFormat Like "[\/]" Then
    MsgBox "Unauthorized Date Format used. Please see code Documentation.", vbCritical
Else
    'Refactoring DateFormat. Using in-built date and formatting
    DateFormat = Format$(Date, DateFormat)
    'Debug.Print DateFormat
    
    str = "C:\MyFiles\FY20XY_XZ\" & Period & "\Emails\" & DateFormat & "\"
	'Concatenating with text string with ampersand to make variable dynamic.
    fol = Dir(str, vbDirectory)
    If fol = "" Then MkDir str
        'If fol returns as "." then the file exists and nothing needs to be done.
End If
End Sub