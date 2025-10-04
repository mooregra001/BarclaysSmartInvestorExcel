Attribute VB_Name = "modImportTransactions"
Rem Attribute VBA_ModuleType=VBAModule
'''References Microsoft ActiveX Data Objects 6.1 Library
'''Option VBASupport 1  '''for LibreOffice
Option Explicit

Public Sub Import()
    On Error GoTo Err_Import
    Dim startRow As Long, endRow As Long
    'Initialize to avoid unassigned variable issues
    startRow = 0
    endRow = 0
    
    TogglePerformance False
    
    Call RetrieveAndAppendData(SHEET_TRANSACTIONS, startRow, endRow)
    
    If startRow > 0 And endRow >= startRow Then
        
        Call PopulateQuantity(SHEET_TRANSACTIONS, startRow, endRow)
        Call PopulateDividendQuantity(SHEET_TRANSACTIONS, startRow, endRow)
        Call PopulatePrice(SHEET_TRANSACTIONS, startRow, endRow)
        Call PopulateModifiedDetails(SHEET_TRANSACTIONS, startRow, endRow)
        Call UpdateBlankModifiedDetailsWithVLOOKUP(SHEET_TRANSACTIONS, startRow, endRow)
        Call UpdateBlankConsolidatedDetailsWithVLOOKUP(SHEET_TRANSACTIONS, startRow, endRow)
        
    End If
    
    MsgBox "Import complete. Data appended, processed, and VLOOKUP applied.", vbInformation
    TogglePerformance True

Exit_Import:
    Exit Sub

Err_Import:
    MsgBox "Error: " & Err.Description, vbCritical
    Resume Exit_Import
End Sub

Private Sub RetrieveAndAppendData(SHEET_TRANSACTIONS As String, ByRef startRow As Long, ByRef endRow As Long)
    Dim ws As Worksheet, wbSource As Workbook, wsSource As Worksheet
    Dim strFile As String, startCol As Long, startRowSource As Long
    Dim dataArray() As Variant, preImportLastRow As Long
    Dim rs As Object, cn As Object

    startRow = 0: endRow = 0
    Set ws = ThisWorkbook.Sheets(SHEET_TRANSACTIONS)

    strFile = GetDataPath()
    If strFile = "" Then Exit Sub
    If Dir(strFile) = "" Then MsgBox "Source file not found: " & strFile, vbCritical: Exit Sub

    Set wbSource = SafeOpenWorkbook(strFile)
    If wbSource Is Nothing Then Exit Sub
    Set wsSource = wbSource.Sheets(1)

    If Not LocateHeader(wsSource, "Date", startRowSource, startCol) Then
        wbSource.Close False: Exit Sub
    End If

    If Not ValidateHeaders(wsSource, startRowSource, startCol) Then
        wbSource.Close False: Exit Sub
    End If

    Dim lastRowSource As Long
    lastRowSource = GetLastRow(wsSource)
    If lastRowSource < startRowSource + 1 Then
        MsgBox "No data found in source file below headers!", vbCritical
        wbSource.Close False: Exit Sub
    End If

    Dim dataRange As String
    dataRange = BuildRangeAddress(startCol, startRowSource, lastRowSource)

    Set cn = OpenADOConnection(strFile)
    If cn Is Nothing Then wbSource.Close False: Exit Sub

    Dim strSQL As String
    strSQL = BuildSQLQuery(wsSource.Name, dataRange)
    Set rs = ExecuteADOQuery(cn, strSQL)
    If rs Is Nothing Then CleanupADO Nothing, cn, wbSource: Exit Sub

    preImportLastRow = GetLastRow(ws)
    If preImportLastRow < 2 Then preImportLastRow = 2

    If Not rs.EOF Then
        dataArray = ConvertRecordsetToArray(rs, startRowSource)
        startRow = preImportLastRow + 1
        endRow = preImportLastRow + rs.RecordCount
        ws.Range("A" & startRow & ":E" & endRow).Value = dataArray
        SortByDate ws, startRow, endRow
    Else
        MsgBox "No data retrieved from source file!", vbInformation
    End If

    CleanupADO rs, cn, wbSource
End Sub

Private Function GetDataPath() As String
    On Error Resume Next
    GetDataPath = ThisWorkbook.Names("DataPath").RefersToRange.Value
    On Error GoTo 0
    If GetDataPath = "" Then MsgBox "DataPath named range is empty or not found!", vbCritical
End Function

Private Function SafeOpenWorkbook(path As String) As Workbook
    On Error Resume Next
    Set SafeOpenWorkbook = Workbooks.Open(path)
    On Error GoTo 0
    If SafeOpenWorkbook Is Nothing Then MsgBox "Failed to open source file: " & path, vbCritical
End Function

Private Function LocateHeader(ws As Worksheet, headerText As String, ByRef rowOut As Long, ByRef colOut As Long) As Boolean
    Dim cell As Range
    Set cell = ws.Cells.Find(headerText, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=True)
    If cell Is Nothing Then
        MsgBox "Header '" & headerText & "' not found in source file!", vbCritical
        LocateHeader = False
    Else
        rowOut = cell.Row
        colOut = cell.Column
        LocateHeader = True
    End If
End Function

Private Function ValidateHeaders(ws As Worksheet, headerRow As Long, startCol As Long) As Boolean
    Dim expectedHeaders As Variant: expectedHeaders = Array("Date", "Details", "Account", "Paid In", "Withdrawn")
    Dim i As Long
    For i = 0 To UBound(expectedHeaders)
        If Trim(ws.Cells(headerRow, startCol + i).Value) <> expectedHeaders(i) Then
            MsgBox "Header mismatch in column " & Chr(65 + startCol - 1 + i) & ": Expected '" & expectedHeaders(i) & "', found '" & ws.Cells(headerRow, startCol + i).Value & "'", vbCritical
            ValidateHeaders = False: Exit Function
        End If
    Next i
    ValidateHeaders = True
End Function

Private Function GetLastRow(ws As Worksheet) As Long
    GetLastRow = ws.Cells.Find("*", , xlValues, , xlByRows, xlPrevious).Row
End Function

Private Function BuildRangeAddress(startCol As Long, startRow As Long, endRow As Long) As String
    Dim startColLetter As String, endColLetter As String
    startColLetter = Split(Cells(1, startCol).Address, "$")(1)
    endColLetter = Split(Cells(1, startCol + 4).Address, "$")(1)
    BuildRangeAddress = startColLetter & startRow & ":" & endColLetter & endRow
End Function

Private Function ConvertRecordsetToArray(rs As Object, baseRow As Long) As Variant
    Dim arr() As Variant, i As Long
    ReDim arr(1 To rs.RecordCount, 1 To HEADER_COUNT_TRANSACTIONS)
    i = 1
    Do While Not rs.EOF
        arr(i, 1) = IIf(IsDate(rs.Fields("Date").Value), CDate(rs.Fields("Date").Value), vbNullString)
        arr(i, 2) = rs.Fields("Details").Value
        arr(i, 3) = rs.Fields("Account").Value
        arr(i, 4) = NullToZero(rs.Fields("Paid In").Value, baseRow + i)
        arr(i, 5) = NullToZero(rs.Fields("Withdrawn").Value, baseRow + i)
        i = i + 1
        rs.MoveNext
    Loop
    ConvertRecordsetToArray = arr
End Function

Private Sub SortByDate(ws As Worksheet, startRow As Long, endRow As Long)
    With ws.Sort
        .SortFields.Clear
        .SortFields.Add Key:=ws.Range("A" & startRow & ":A" & endRow), _
            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SetRange ws.Range("A" & startRow & ":E" & endRow)
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply
    End With
End Sub



