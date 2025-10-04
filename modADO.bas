Attribute VB_Name = "modADO"
Rem Attribute VBA_ModuleType=VBAModule
'''Option VBASupport 1  '''For LibreOffice
Option Explicit

'@Description Opens an ADO connection to an Excel file using ACE OLEDB
'@Param filePath - Full path to the Excel file
'@Return ADODB.Connection object or Nothing on failure
Public Function OpenADOConnection(filePath As String) As Object
    Dim cn As Object
    Set cn = CreateObject("ADODB.Connection")
    
    If Len(Trim(filePath)) = 0 Then
        MsgBox "Invalid file path", vbExclamation
        Set OpenADOConnection = Nothing
        Exit Function
    End If

    On Error Resume Next
    cn.Open "Provider=Microsoft.ACE.OLEDB.12.0;" & _
            "Data Source=" & filePath & ";" & _
            "Extended Properties='Excel 12.0 Xml;HDR=YES';"
    If Err.Number <> 0 Then
        MsgBox "Failed to connect to source file: " & Err.Description, vbCritical
        Set OpenADOConnection = Nothing
        Exit Function
    End If
    On Error GoTo 0
    Set OpenADOConnection = cn
End Function

'@Description Executes a SQL query against an open ADO connection
'@Param cn - Active ADODB.Connection object
'@Param sqlQuery - SQL query string to execute
'@Return ADODB.Recordset object or Nothing on failure
Public Function ExecuteADOQuery(cn As Object, sqlQuery As String) As Object
    Dim rs As Object
    Set rs = CreateObject("ADODB.Recordset")
    
    If cn Is Nothing Then
        MsgBox "Connection object is not initialized.", vbExclamation
        Set ExecuteADOQuery = Nothing
        Exit Function
    End If

    If Len(Trim(sqlQuery)) = 0 Then
        MsgBox "SQL query is empty.", vbExclamation
        Set ExecuteADOQuery = Nothing
        Exit Function
    End If

    On Error Resume Next
    rs.Open sqlQuery, cn, 3, 3 ' adOpenStatic, adLockOptimistic
    If Err.Number <> 0 Then
        MsgBox "Error executing query: " & Err.Description & vbCrLf & "Query: " & sqlQuery, vbCritical
        Set ExecuteADOQuery = Nothing
        Exit Function
    End If
    On Error GoTo 0
    Set ExecuteADOQuery = rs
End Function

'@Description Closes and releases ADO objects and optionally a workbook
'@Param rs - ADODB.Recordset object to close
'@Param cn - ADODB.Connection object to close
'@Param wb - Workbook object to close without saving
Public Sub CleanupADO(rs As Object, cn As Object, wb As Workbook)
    If Not rs Is Nothing Then rs.Close
    If Not cn Is Nothing Then cn.Close
    If Not wb Is Nothing Then wb.Close False
    Set rs = Nothing
    Set cn = Nothing
    Set wb = Nothing
End Sub

