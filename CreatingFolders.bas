Attribute VB_Name = "CreatingFolders"
Option Explicit
Sub CurrentPeriod()
    ' This macro selects the current period, for cell "A2" in the Sheet1 worksheet
    ' Given the start and end dates as given below. This  applies to FY 20XX-XY
        ' PERIOD        START DATE  END DATE
        'P01 January	AA/BB/20XX  AA/12/20XY
        'P02 February	AA/BB/20XX  AA/12/20XY
        'P03 March	AA/BB/20XX  AA/12/20XY
        'P04 April	AA/BB/20XX  AA/12/20XY
        'P05 May	AA/BB/20XX  AA/12/20XY

Dim myRange As Range
Dim TODAY

Dim JanStartDate
Dim JanEndDate
Dim FebStartDate
Dim FebEndDate
Dim MarStartDate
Dim MarEndDate
Dim AprStartDate
Dim AprEndDate
Dim MayStartDate
Dim MayEndDate
Dim FY20XY_XZ

Set myRange = Worksheets("Sheet1").Range("A2")
TODAY = Date
    'Date contains the current system date.
'TODAY = Worksheets("Sheet1").Cells(11, 2) ' Used as check for output to cell.

'Creating the variables for the start and end dates of each period.
DecStartDate = DateSerial(20XX, BB, AA)
DecEndDate = DateSerial(20XX, BB, AC)

JanStartDate = DateSerial(20XY, DD, EE)
JanEndDate = DateSerial(20XY, DD, EF)

FebStartDate = DateSerial(20XY, DD, EE)
FebEndDate = DateSerial(20XY, DD, EF)

MarStartDate = DateSerial(20XY, DD, EE)
MarEndDate = DateSerial(20XY, DD, EF)

AprStartDate = DateSerial(20XY, DD, EE)
AprEndDate = DateSerial(20XY, DD, EF)

FY2023_24 = DateSerial(20XZ, GG, HH)

'Worksheets("Sheet1").Cells(7, 2).Value = DecStartDate 'checking the output

' Using if and elseif statements to populate chosen cell with period

If TODAY >= MayStartDate And TODAY <= MayEndDate Then
    myRange = "P05 May"
ElseIf TODAY >= JanStartDate And TODAY <= JanEndDate Then
    myRange = "P01 January"
ElseIf TODAY >= FebStartDate And TODAY <= FebEndDate Then
    myRange = "P02 February"
ElseIf TODAY >= MarStartDate And TODAY <= MarEndDate Then
    myRange = "P03 March"
ElseIf TODAY >= AprStartDate And TODAY <= AprEndDate Then
    myRange = "P04 April"
ElseIf TODAY >= FY20XY_XZ Then
    MsgBox "Date outside of current fiscal period. Work in the correct period."
Else
    MsgBox "Date given outside scope accounted for. Please refer to the code documentation."
    
End If
End Sub

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

Sub CreateFolder()

Call CurrentPeriod
Call MakeDateDirectory

End Sub

