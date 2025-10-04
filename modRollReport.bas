Attribute VB_Name = "modRollReport"
Rem Attribute VBA_ModuleType=VBAModule
'''Option VBASupport 1  '''For LibeOffice
Option Explicit

' ============================================
' Module: RollReport
' Purpose: Rolls forward reporting date, saves workbook, and clears investment data
' Dependencies: Named ranges on the Control sheet (PDate, LDate, BasePath, FilePath, FileName)
'               LDate drives folder naming and file save location.
' Notes:
' - User input for business date is expected in UK format: dd/mm/yyyy.
' - Folder structure is generated using LDate in "YYYY\mm.Mmm" format (e.g., "2025\09.Sep")
'   to ensure chronological sorting and readability.
' Author:
' Last Updated: 2025-09-01
' ============================================

Sub RollReport()
    On Error GoTo FatalError
    Application.ScreenUpdating = False

    Dim strFullPath As String
    Dim fsFolderPath As Object
    Dim FolderPath As String
    Dim LDateValue As Date
    Dim BasePath As String
    Dim YearFolder As String
    Dim MonthFolder As String

    ' Validate named ranges individually
    On Error GoTo NamedRangeError
    [PDate] = [LDate]
    On Error GoTo 0

    Call doLDate
    ActiveSheet.Calculate

    On Error GoTo NamedRangeError
    LDateValue = Range("LDate").Value
    On Error GoTo 0

    Set fsFolderPath = CreateObject("Scripting.FileSystemObject")

    On Error GoTo NamedRangeError
    BasePath = Range("BasePath").Value
    On Error GoTo 0

    If Right(BasePath, 1) <> "\" Then BasePath = BasePath & "\"

    If Not fsFolderPath.FolderExists(BasePath) Then
        MsgBox "Static base path does not exist: " & BasePath, vbCritical
        GoTo Cleanup
    End If

    YearFolder = Format(LDateValue, "YYYY")
    FolderPath = BasePath & YearFolder
    If Not fsFolderPath.FolderExists(FolderPath) Then
        fsFolderPath.CreateFolder FolderPath
    End If

    MonthFolder = Format(LDateValue, "mm.Mmm")
    FolderPath = FolderPath & "\" & MonthFolder
    If Not fsFolderPath.FolderExists(FolderPath) Then
        fsFolderPath.CreateFolder FolderPath
    End If

    On Error GoTo NamedRangeError
    strFullPath = Range("FilePath").Value & Range("FileName").Value
    On Error GoTo 0

    On Error GoTo SaveError
    ThisWorkbook.SaveAs strFullPath
    On Error GoTo 0

    Call ClearInvestmentsData

    Sheets("Control").Range("LDate").Activate
    MsgBox "Finished rolling report!"

Cleanup:
    Application.ScreenUpdating = True
    Exit Sub

NamedRangeError:
    MsgBox "Error accessing named range or formula-driven cell: " & Err.Description, vbCritical
    Resume Cleanup

SaveError:
    MsgBox "Error saving workbook to path: " & strFullPath & vbCrLf & "Details: " & Err.Description, vbCritical
    Resume Cleanup

FatalError:
    MsgBox "Unexpected error: " & Err.Description, vbCritical
    Resume Cleanup
End Sub

Private Sub doLDate()
    Dim TodayDate As Integer
    Dim LDate As String

    TodayDate = Weekday(Date)
    If TodayDate = 2 Then
        LDate = Format(Date - 3, "MM/DD/YYYY")
    Else
        LDate = Format(Date - 1, "MM/DD/YYYY")
    End If

    LDate = InputBox("Please enter business date", , LDate)
    Range("LDate").Value = LDate
End Sub

Private Sub ClearInvestmentsData()
    Dim ws As Worksheet
    Dim rngInvestments As Range
    Dim LastCol As Long
    Dim LastRow As Long
    
    Set ws = ThisWorkbook.Worksheets("Investments")
    
    ' Find the last column with data in row 1, starting from B1
    LastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If LastCol < 2 Then LastCol = 2 ' Ensure we don't go left of column B
    
    ' Find the last row with data in column B
    LastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    If LastRow <= 1 Then Exit Sub ' Exit if no data below header
    
    ' Set the range from B2 to the last column and last row
    Set rngInvestments = ws.Range(ws.Cells(2, 2), ws.Cells(LastRow, LastCol))
    
    ' Clear the contents of the range
    rngInvestments.ClearContents
End Sub
