Attribute VB_Name = "modFunctions"
Rem Attribute VBA_ModuleType=VBAModule
'''Option VBASupport 1  '''For LibreOffice
Option Explicit

'------------------------------------------------------------
' Function: BuildSQLQuery
' Purpose: Constructs an ADO SQL query string for a given sheet and range
' Input:
'   - sheetName: Name of the worksheet
'   - dataRange: Excel range in A1-style (e.g., "A2:E100")
' Output: SQL string for use with ADO (e.g., "SELECT * FROM [Sheet1$A2:E100]")
'------------------------------------------------------------
Public Function BuildSQLQuery(sheetName As String, dataRange As String) As String
    ' Construct SQL query for ADO to retrieve data from the specified range
    BuildSQLQuery = "SELECT * FROM [" & sheetName & "$" & dataRange & "]"
End Function

'------------------------------------------------------------
' Function: NullToZero
' Purpose: Converts null, empty, or non-numeric values to 0 for numeric processing
' Input:
'   - val: Variant value from worksheet or recordset
'   - rowNum: Row number for logging/debugging
' Output: Double value (0 if invalid)
'------------------------------------------------------------
Public Function NullToZero(val As Variant, rowNum As Long) As Double
    ' Convert Null, empty, or non-numeric values to 0
    If IsNull(val) Or Trim(val & "") = "" Then
        Debug.Print "Null or empty value at row " & rowNum
        NullToZero = 0
        Exit Function
    End If
    
    Dim cleanVal As String
    cleanVal = Replace(Trim(val & ""), ",", "")
    
    If IsNumeric(cleanVal) Then
        On Error Resume Next
        NullToZero = CDbl(cleanVal)
        If Err.Number <> 0 Then
            Debug.Print "Invalid number at row " & rowNum & ": " & val
            NullToZero = 0
        End If
        On Error GoTo 0
    Else
        Debug.Print "Non-numeric value at row " & rowNum & ": " & val
        NullToZero = 0
    End If
End Function
