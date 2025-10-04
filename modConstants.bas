Attribute VB_Name = "modConstants"
Rem Attribute VBA_ModuleType=VBAModule
'''Option VBASupport 1  '''for LibreOffice
Option Explicit

' === Transactions Constants ===
Public Const SHEET_TRANSACTIONS As String = "Transactions"
Public Const HEADER_COUNT_TRANSACTIONS As Long = 5
Public Const VLOOKUP_FORMULA_CONSOLIDATED As String = "=VLOOKUP(trim(RC[-1]),MapConsolidatedDetails!C1:C2,2,0)"
Public Const VLOOKUP_FORMULA_MISC As String = "=VLOOKUP(trim(RC[-6]),MapMiscTransactions!C1:C2,2,0)"

' === Investments Constants ===
Public Const SHEET_INVESTMENTS As String = "Investments"
Public Const HEADER_COUNT_INVESTMENTS As Long = 15
Public Const VLOOKUP_FORMULA As String = "=VLOOKUP(TRIM(RC[1]),MapConsolidatedDetails!C1:C2,2,FALSE)"


