Attribute VB_Name = "modImportInvestments"
Rem Attribute VBA_ModuleType=VBAModule
'''References Microsoft ActiveX Data Objects 6.1 Library
'''Option VBASupport 1  '''for LibreOffice
Option Explicit

' === Main Entry Point ===
Public Sub ImportInvestments()
    On Error GoTo Err_Import
    Dim startRow As Long, endRow As Long
    startRow = 0: endRow = 0

    TogglePerformance False

    Call RetrieveInvestmentsData(SHEET_INVESTMENTS, startRow, endRow)
    Call UpdateInvestmentsWithVLOOKUP(SHEET_INVESTMENTS, startRow, endRow)

    MsgBox "Import of Investments complete.", vbInformation

Exit_Import:
    TogglePerformance True
    Exit Sub

Err_Import:
    MsgBox "Error: " & Err.Description, vbCritical
    Resume Exit_Import
End Sub

' === Performance Toggle ===
Public Sub TogglePerformance(ByVal enable As Boolean)
    With Application
        .ScreenUpdating = enable
        .EnableEvents = enable
        .Calculation = IIf(enable, xlCalculationAutomatic, xlCalculationManual)
    End With
End Sub

' === Data Retrieval ===
Private Sub RetrieveInvestmentsData(sheetName As String, ByRef startRow As Long, ByRef endRow As Long)
    Dim cn As Object, rs As Object
    Dim strFile As String, strSQL As String
    Dim ws As Worksheet, wbSource As Workbook, wsSource As Worksheet
    Dim LastRow As Long, dataRange As String
    Dim dataArray() As Variant, rowIndex As Long
    Dim dateCell As Range
    Dim startCol As Long, startRowSource As Long

    startRow = 2: endRow = 0
    Set ws = ThisWorkbook.Sheets(sheetName)
    If ws Is Nothing Then
        MsgBox "Sheet '" & sheetName & "' not found!", vbCritical
        Exit Sub
    End If

    On Error Resume Next
    strFile = ThisWorkbook.Names("InvestmentsPath").RefersToRange.Value
    On Error GoTo 0

    If strFile = "" Or Dir(strFile) = "" Then
        MsgBox "Invalid or missing source file path!", vbCritical
        Exit Sub
    End If

    Set wbSource = Workbooks.Open(strFile)
    If wbSource Is Nothing Then
        MsgBox "Failed to open source file: " & strFile, vbCritical
        Exit Sub
    End If

    Set wsSource = wbSource.Sheets(1)
    Set dateCell = wsSource.Cells.Find("Investment", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=True)
    If dateCell Is Nothing Then
        MsgBox "Header 'Investment' not found!", vbCritical
        wbSource.Close False
        Exit Sub
    End If

    startRowSource = dateCell.Row
    startCol = dateCell.Column

    Dim expectedHeaders As Variant
    expectedHeaders = Array("Investment", "Identifier", "Quantity Held", "Last Price", "Last Price CCY", _
                            "Value", "Value CCY", "FX Rate", "Last Price (p)", "Value (£)", "Book Cost", _
                            "Book Cost CCY", "Average FX Rate", "Book Cost (£)", "% Change")

    Dim i As Long
    For i = 0 To HEADER_COUNT_INVESTMENTS - 1
        If Trim(wsSource.Cells(startRowSource, startCol + i).Value) <> expectedHeaders(i) Then
            MsgBox "Header mismatch in column " & Chr(65 + startCol - 1 + i) & ": Expected '" & expectedHeaders(i) & _
                   "', found '" & wsSource.Cells(startRowSource, startCol + i).Value & "'", vbCritical
            wbSource.Close False
            Exit Sub
        End If
    Next i

    LastRow = wsSource.Cells.Find("*", , xlValues, , xlByRows, xlPrevious).Row
    If LastRow < startRowSource + 1 Then
        MsgBox "No data found in source file below headers!", vbCritical
        wbSource.Close False
        Exit Sub
    End If

    Dim startColLetter As String, endColLetter As String
    startColLetter = Split(Cells(1, startCol).Address, "$")(1)
    endColLetter = Split(Cells(1, startCol + HEADER_COUNT_INVESTMENTS - 1).Address, "$")(1)
    dataRange = startColLetter & startRowSource & ":" & endColLetter & LastRow

    Set cn = OpenADOConnection(strFile)
    If cn Is Nothing Then
        wbSource.Close False
        Exit Sub
    End If

    strSQL = BuildSQLQuery(wsSource.Name, dataRange)
    Set rs = ExecuteADOQuery(cn, strSQL)
    If rs Is Nothing Then
        CleanupADO Nothing, cn, wbSource
        Exit Sub
    End If

    If Not rs.EOF Then
        ReDim dataArray(1 To rs.RecordCount, 1 To HEADER_COUNT_INVESTMENTS)
        rowIndex = 1
        Do While Not rs.EOF
            dataArray(rowIndex, 1) = rs.Fields("Investment").Value
            dataArray(rowIndex, 2) = rs.Fields("Identifier").Value
            dataArray(rowIndex, 3) = NullToZero(rs.Fields("Quantity Held").Value, startRowSource + rowIndex)
            dataArray(rowIndex, 4) = NullToZero(rs.Fields("Last Price").Value, startRowSource + rowIndex)
            dataArray(rowIndex, 5) = rs.Fields("Last Price CCY").Value
            dataArray(rowIndex, 6) = NullToZero(rs.Fields("Value").Value, startRowSource + rowIndex)
            dataArray(rowIndex, 7) = rs.Fields("Value CCY").Value
            dataArray(rowIndex, 8) = NullToZero(rs.Fields("FX Rate").Value, startRowSource + rowIndex)
            dataArray(rowIndex, 9) = NullToZero(rs.Fields("Last Price (p)").Value, startRowSource + rowIndex)
            dataArray(rowIndex, 10) = NullToZero(rs.Fields("Value (£)").Value, startRowSource + rowIndex)
            dataArray(rowIndex, 11) = NullToZero(rs.Fields("Book Cost").Value, startRowSource + rowIndex)
            dataArray(rowIndex, 12) = rs.Fields("Book Cost CCY").Value
            dataArray(rowIndex, 13) = NullToZero(rs.Fields("Average FX Rate").Value, startRowSource + rowIndex)
            dataArray(rowIndex, 14) = NullToZero(rs.Fields("Book Cost (£)").Value, startRowSource + rowIndex)
            dataArray(rowIndex, 15) = NullToZero(rs.Fields("% Change").Value, startRowSource + rowIndex)
            rowIndex = rowIndex + 1
            rs.MoveNext
        Loop
        endRow = startRow + rs.RecordCount - 1
        ws.Range(ws.Cells(startRow, 2), ws.Cells(endRow, HEADER_COUNT_INVESTMENTS)).Value = dataArray
    Else
        MsgBox "No data retrieved from source file!", vbInformation
    End If

    CleanupADO rs, cn, wbSource
End Sub

' === VLOOKUP Injection ===
Private Sub UpdateInvestmentsWithVLOOKUP(sheetName As String, startRow As Long, endRow As Long)
    On Error GoTo Err_Update
    Dim ws As Worksheet, i As Long

    If startRow <= 0 Or endRow < startRow Or Len(sheetName) = 0 Then
        MsgBox "Invalid row range or sheet name.", vbExclamation
        Exit Sub
    End If

    Set ws = ThisWorkbook.Sheets(sheetName)

    For i = startRow To endRow
        With ws.Cells(i, 1)
            .FormulaR1C1 = VLOOKUP_FORMULA
            .Interior.ThemeColor = xlThemeColorDark2
        End With
    Next i
    Exit Sub

Err_Update:
    MsgBox "Error in UpdateInvestmentsWithVLOOKUP: " & Err.Description, vbCritical
End Sub

