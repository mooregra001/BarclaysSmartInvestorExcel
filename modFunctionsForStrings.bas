Attribute VB_Name = "modFunctionsForStrings"
Rem Attribute VBA_ModuleType=VBAModule
'''References Microsoft ActiveX Data Objects 6.1 Library
'''Option VBASupport 1  '''for LibreOffice
Option Explicit

Public Sub PopulateQuantity(SHEET_TRANSACTIONS As String, startRow As Long, endRow As Long)
    Const COL_QUANTITY As Long = 6
    Dim ws As Worksheet
    Dim i As Long
    Dim Details As String
    Dim quantity As Long

    Set ws = ThisWorkbook.Sheets(SHEET_TRANSACTIONS)

    If ws.Cells(1, COL_QUANTITY).Value = "" Then
        ws.Cells(1, COL_QUANTITY).Value = "Quantity"
    End If

    If startRow < 2 Then startRow = 2
    If endRow < startRow Then
        MsgBox "No data to process!", vbExclamation
        Exit Sub
    End If

    For i = startRow To endRow
        Details = ws.Cells(i, 2).Value
        quantity = 0

        quantity = ExtractQuantity(Details, "Bought", False)
        If quantity = 0 Then
            quantity = ExtractQuantity(Details, "Sold", True)
        End If

        ws.Cells(i, COL_QUANTITY).Value = quantity
    Next i
End Sub

Private Function ExtractQuantity(ByVal Details As String, ByVal keyword As String, ByVal isNegative As Boolean) As Long
    Dim keywordPos As Long, spacePos As Long, nextSpacePos As Long
    Dim qtyStr As String

    keywordPos = InStr(1, Details, keyword, vbTextCompare)
    If keywordPos = 0 Then Exit Function

    spacePos = keywordPos + Len(keyword) + 1
    nextSpacePos = InStr(spacePos, Details, " ")
    If nextSpacePos = 0 Then nextSpacePos = Len(Details) + 1

    qtyStr = Mid(Details, spacePos, nextSpacePos - spacePos)
    If IsNumeric(qtyStr) Then
        ExtractQuantity = CLng(qtyStr)
        If isNegative Then ExtractQuantity = -ExtractQuantity
    End If
End Function

Public Sub PopulateDividendQuantity(SHEET_TRANSACTIONS As String, startRow As Long, endRow As Long)
    Const COL_DETAILS As Long = 2
    Const COL_QUANTITY As Long = 6
    Dim ws As Worksheet
    Dim i As Long
    Dim Details As String
    Dim quantity As Long

    ' Set reference to worksheet
    Set ws = ThisWorkbook.Sheets(SHEET_TRANSACTIONS)

    ' Ensure Quantity header exists
    If ws.Cells(1, COL_QUANTITY).Value = "" Then
        ws.Cells(1, COL_QUANTITY).Value = "Quantity"
    End If

    ' Validate row range
    If startRow < 2 Then startRow = 2
    If endRow < startRow Then
        MsgBox "No dividend data to process!", vbExclamation
        Exit Sub
    End If

    ' Process rows
    For i = startRow To endRow
        Details = ws.Cells(i, COL_DETAILS).Value
        quantity = ExtractDividendQuantity(Details)
        If quantity > 0 Then
            ws.Cells(i, COL_QUANTITY).Value = quantity
        End If
    Next i
End Sub

Private Function ExtractDividendQuantity(ByVal Details As String) As Long
    Const keyword As String = "Automatic dividend reinvest - purchase"
    Dim keywordPos As Long
    Dim spacePos As Long
    Dim nextSpacePos As Long
    Dim qtyStr As String

    ExtractDividendQuantity = 0 ' Default

    keywordPos = InStr(1, Details, keyword, vbTextCompare)
    If keywordPos > 0 Then
        spacePos = keywordPos + Len(keyword) + 1
        nextSpacePos = InStr(spacePos, Details, " ")
        If nextSpacePos = 0 Then nextSpacePos = Len(Details) + 1

        qtyStr = Mid(Details, spacePos, nextSpacePos - spacePos)
        If IsNumeric(qtyStr) Then
            ExtractDividendQuantity = CLng(qtyStr)
        End If
    End If
End Function

Public Function StripBuySell(Details As Variant) As String
    Dim result As String
    
    If IsNull(Details) Or Details = "" Then
        StripBuySell = ""
        Exit Function
    End If
    
    result = Trim(Details)
    If InStr(1, result, " Buy", vbTextCompare) > 0 Then
        result = Left(result, InStr(1, result, " Buy", vbTextCompare) - 1)
    ElseIf InStr(1, result, " Sell", vbTextCompare) > 0 Then
        result = Left(result, InStr(1, result, " Sell", vbTextCompare) - 1)
    End If
    
    StripBuySell = result
End Function

Public Sub PopulatePrice(SHEET_TRANSACTIONS As String, startRow As Long, endRow As Long)
    Const COL_QUANTITY As Long = 6
    Const COL_PRICE As Long = 7
    Const COL_PAIDIN As Long = 4
    Const COL_WITHDRAWN As Long = 5

    Dim ws As Worksheet
    Dim i As Long
    Dim quantity As Long
    Dim paidIn As Double
    Dim withdrawn As Double
    Dim price As Double

    ' Set reference to worksheet
    Set ws = ThisWorkbook.Sheets(SHEET_TRANSACTIONS)

    ' Ensure Price header exists
    If ws.Cells(1, COL_PRICE).Value = "" Then
        ws.Cells(1, COL_PRICE).Value = "Price"
    End If

    ' Validate row range
    If startRow < 2 Then startRow = 2
    If endRow < startRow Then
        MsgBox "No data to process for Price!", vbExclamation
        Exit Sub
    End If

    ' Calculate price for each row
    For i = startRow To endRow
        quantity = ws.Cells(i, COL_QUANTITY).Value
        paidIn = ws.Cells(i, COL_PAIDIN).Value
        withdrawn = ws.Cells(i, COL_WITHDRAWN).Value

        price = CalculatePrice(quantity, paidIn, withdrawn)
        ws.Cells(i, COL_PRICE).Value = price
    Next i
End Sub

Private Function CalculatePrice(quantity As Long, paidIn As Double, withdrawn As Double) As Double
    If quantity = 0 Then
        CalculatePrice = 0
    ElseIf quantity > 0 Then
        CalculatePrice = -withdrawn / quantity
    Else
        CalculatePrice = paidIn / -quantity
    End If
End Function

Public Sub PopulateModifiedDetails(SHEET_TRANSACTIONS As String, startRow As Long, endRow As Long)
    Const COL_DETAILS As Long = 2
    Const COL_MODIFIED As Long = 8
    Dim ws As Worksheet
    Dim i As Long
    Dim Details As String
    Dim result As String
    
    ' Set reference to worksheet
    Set ws = ThisWorkbook.Sheets(SHEET_TRANSACTIONS)
    
    ' Ensure ModifiedDetails header exists
    If ws.Cells(1, COL_MODIFIED).Value = "" Then
        ws.Cells(1, COL_MODIFIED).Value = "ModifiedDetails"
    End If
    
    ' Validate row range
    If startRow < 2 Then startRow = 2
    If endRow < startRow Then
        MsgBox "No data to process for ModifiedDetails!", vbExclamation
        Exit Sub
    End If
    
    ' Process rows
    For i = startRow To endRow
        Details = ws.Cells(i, COL_DETAILS).Value ' Details in column B
        If Len(Details) > 0 Then
            ' Apply the functions in sequence
            result = StripTransactionFee(Details)
            result = StripAfterQuantity(result)
            result = StripAfterDividend(result)
            result = HandleAdminFees(result)
            result = HandleInterest(result)
            
            ' Check if result matches original Details (no transformation)
            If Trim(result) = Trim(Details) Then
                ws.Cells(i, COL_MODIFIED).Value = "" ' Leave blank for further processing
            Else
                ws.Cells(i, COL_MODIFIED).Value = result ' Update with transformed result
            End If
        Else
            ws.Cells(i, COL_MODIFIED).Value = ""
        End If
    Next i
End Sub

Public Function StripTransactionFee(Details As Variant) As String
    Const FEE_PREFIX As String = "Online transaction fee "
    Dim result As String
    Dim feePos As Long

    ' Handle null or empty input
    If IsEmpty(Details) Or Details = "" Then
        StripTransactionFee = ""
        Exit Function
    End If

    ' Clean input
    result = Trim(Details)

    ' Remove fee prefix if present
    feePos = InStr(1, result, FEE_PREFIX, vbTextCompare)
    If feePos > 0 Then
        result = Trim(Mid(result, feePos + Len(FEE_PREFIX)))
    End If

    ' Remove trailing verb suffix
    result = RemoveTrailingVerb(result)

    StripTransactionFee = result
End Function

Private Function RemoveTrailingVerb(ByVal text As String) As String
    Const VERB_BUY As String = " Buy"
    Const VERB_SELL As String = " Sell"
    Dim verbPos As Long

    verbPos = InStr(1, text, VERB_BUY, vbTextCompare)
    If verbPos > 0 Then
        RemoveTrailingVerb = Left(text, verbPos - 1)
        Exit Function
    End If

    verbPos = InStr(1, text, VERB_SELL, vbTextCompare)
    If verbPos > 0 Then
        RemoveTrailingVerb = Left(text, verbPos - 1)
        Exit Function
    End If

    RemoveTrailingVerb = text
End Function

Public Function StripAfterQuantity(Details As Variant) As String
    Const KW_BOUGHT As String = "Bought"
    Const KW_SOLD As String = "Sold"
    Const DELIMITER_AT As String = " @"
    
    Dim result As String
    Dim keywordFound As String
    Dim keywordPos As Long

    ' Handle null or empty input
    If IsEmpty(Details) Or Details = "" Then
        StripAfterQuantity = ""
        Exit Function
    End If

    result = Trim(Details)
    keywordFound = GetTransactionKeyword(result, KW_BOUGHT, KW_SOLD)

    If keywordFound <> "" Then
        keywordPos = InStr(1, result, keywordFound, vbTextCompare)
        result = ExtractAfterQuantity(result, keywordPos, keywordFound, DELIMITER_AT)
    End If

    StripAfterQuantity = result
End Function

Private Function GetTransactionKeyword(ByVal text As String, ByVal kw1 As String, ByVal kw2 As String) As String
    If InStr(1, text, kw1, vbTextCompare) > 0 Then
        GetTransactionKeyword = kw1
    ElseIf InStr(1, text, kw2, vbTextCompare) > 0 Then
        GetTransactionKeyword = kw2
    Else
        GetTransactionKeyword = ""
    End If
End Function

Private Function ExtractAfterQuantity(ByVal text As String, ByVal keywordPos As Long, ByVal keyword As String, ByVal delimiter As String) As String
    Dim spacePos As Long, nextSpacePos As Long, quantityText As String, atPos As Long

    spacePos = keywordPos + Len(keyword)
    If spacePos > Len(text) Then Exit Function
    If Mid(text, spacePos, 1) <> " " Then Exit Function

    spacePos = spacePos + 1
    nextSpacePos = InStr(spacePos, text, " ")
    If nextSpacePos = 0 Then nextSpacePos = Len(text) + 1

    quantityText = Trim(Mid(text, spacePos, nextSpacePos - spacePos))
    If Not IsNumeric(quantityText) Then Exit Function

    ExtractAfterQuantity = Trim(Mid(text, nextSpacePos + 1))
    atPos = InStr(1, ExtractAfterQuantity, delimiter)
    If atPos > 0 Then
        ExtractAfterQuantity = Trim(Left(ExtractAfterQuantity, atPos - 1))
    End If
End Function

Public Function StripAfterDividend(Details As Variant) As String
    Const KW_DIVIDEND As String = "Dividend: "
    Const DELIMITER_PAREN As String = " ("
    Dim result As String
    Dim dividendPos As Long
    Dim spacePos As Long
    Dim atPos As Long
    
    ' Handle null or empty input
    If IsEmpty(Details) Or Details = "" Then
        StripAfterDividend = ""
        Exit Function
    End If
    
    ' Clean input
    result = Trim(Details)
    
    ' Check if string begins with "Dividend: "
    dividendPos = InStr(1, result, KW_DIVIDEND, vbTextCompare)
    If dividendPos = 1 Then
        ' Find the space after "Dividend: "
        spacePos = dividendPos + Len(KW_DIVIDEND)
        If spacePos <= Len(result) Then
            ' Return string AFTER "Dividend: "
            result = Trim(Mid(result, spacePos))
            ' Find position of " (" and remove everything to the right
            atPos = InStr(1, result, DELIMITER_PAREN)
            If atPos > 0 Then
                result = Trim(Left(result, atPos - 1))
            End If
        End If
    End If
    
    StripAfterDividend = result
End Function

Public Function HandleAdminFees(Details As Variant) As String
    Dim feeKeywords As Variant
    feeKeywords = Array("Customer fee", "fee payment", "SippdealAdminFee", "SIPP Admin Fee", "Administration Fee")

    HandleAdminFees = NormalizeByKeywords(Details, feeKeywords, "Admin Fee")
End Function

Public Function NormalizeByKeywords(ByVal inputText As Variant, ByVal keywords As Variant, ByVal normalizedLabel As String) As String
    Dim i As Long
    Dim cleanedText As String

    If IsNull(inputText) Or Trim(inputText) = "" Then
        NormalizeByKeywords = ""
        Exit Function
    End If

    cleanedText = Trim(inputText)

    For i = LBound(keywords) To UBound(keywords)
        If InStr(1, cleanedText, keywords(i), vbTextCompare) > 0 Then
            NormalizeByKeywords = normalizedLabel
            Exit Function
        End If
    Next i

    NormalizeByKeywords = cleanedText
End Function

Public Function HandleInterest(Details As Variant) As String
    Dim interestKeywords As Variant
    interestKeywords = Array("Interest") ' Expandable if needed

    HandleInterest = NormalizeByKeywords(Details, interestKeywords, "Interest")
End Function

Public Sub UpdateBlankModifiedDetailsWithVLOOKUP(SHEET_TRANSACTIONS As String, startRow As Long, endRow As Long)
    ' Purpose: Updates blank cells in the ModifiedDetails  with a VLOOKUP formula
    '          referencing Details and MapMiscTransactions sheet.
    
    On Error GoTo Err_Update
    Const COL_MODIFIED_DETAILS As Long = 8
    Dim ws As Worksheet
    Dim i As Long
    
    ' Validate inputs
    If startRow <= 0 Or endRow < startRow Or Len(SHEET_TRANSACTIONS) = 0 Then
        MsgBox "Invalid row range or sheet name.", vbExclamation
        Exit Sub
    End If
    
    ' Set worksheet
    Set ws = ThisWorkbook.Sheets(SHEET_TRANSACTIONS)
    
    ' Loop through rows and update blank ModifiedDetails
    For i = startRow To endRow
        If IsEmpty(ws.Cells(i, COL_MODIFIED_DETAILS)) Or ws.Cells(i, COL_MODIFIED_DETAILS).Value = "" Then
            ws.Cells(i, COL_MODIFIED_DETAILS).FormulaR1C1 = VLOOKUP_FORMULA_MISC
        End If
    Next i
        
    Exit Sub

Err_Update:
    MsgBox "Error in UpdateBlankModifiedDetailsWithVLOOKUP: " & Err.Description, vbCritical
End Sub

Public Sub UpdateBlankConsolidatedDetailsWithVLOOKUP(SHEET_TRANSACTIONS As String, startRow As Long, endRow As Long)
    ' Purpose: Updates blank cells in the ConsolidatedDetails  with a VLOOKUP formula
    '          referencing ModifiedDetails and MapConsolidatedDetails sheet.
    
    On Error GoTo Err_Update
    Const COL_CONSOLIDATED_DETAILS As Long = 9
    Dim ws As Worksheet
    Dim i As Long
    
    ' Validate inputs
    If startRow <= 0 Or endRow < startRow Or Len(SHEET_TRANSACTIONS) = 0 Then
        MsgBox "Invalid row range or sheet name.", vbExclamation
        Exit Sub
    End If
    
    ' Set worksheet
    Set ws = ThisWorkbook.Sheets(SHEET_TRANSACTIONS)
    
    ' Loop through rows and update MapConsolidatedDetails
    For i = startRow To endRow
        If IsEmpty(ws.Cells(i, COL_CONSOLIDATED_DETAILS)) Or ws.Cells(i, COL_CONSOLIDATED_DETAILS).Value = "" Then
            With ws.Cells(i, COL_CONSOLIDATED_DETAILS)
                .FormulaR1C1 = VLOOKUP_FORMULA_CONSOLIDATED
                .Interior.ThemeColor = xlThemeColorDark2
            End With
        End If
    Next i
    
    Exit Sub

Err_Update:
    MsgBox "Error in UpdateBlankModifiedDetailsWithVLOOKUP: " & Err.Description, vbCritical
End Sub

