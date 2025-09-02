Attribute VB_Name = "tool9b"
'Author/Developer: Vadim Krifuks
'Collaborators: Hui Zeng, Man Ming Tse
'Last Updated: 5Feb2025

Option Explicit


Sub XToTotals(ByVal unitRatesRng As Range, ByVal totalsRng As Range)
'this subroutine turns "x" into totals effectively unblinding the budget
'skips all blanks and non-numerical values
 
    Dim totalRows As Integer
    Dim totalColumns As Integer
    
    totalRows = unitRatesRng.rows.count
    totalColumns = totalsRng.Columns.count

    Dim curRow As Integer
    Dim curColumn As Integer
    
    Dim totalsValue As Variant
    Dim totalsFormulaString As String
    Dim unitRateAddressString As String
    Dim unitRate As Variant
    
    For curRow = 1 To totalRows
        unitRate = unitRatesRng.Cells(curRow).Value
        unitRateAddressString = unitRatesRng.Cells(curRow, 1).Address(RowAbsolute:=False)
        If IsNumeric(unitRate) Then
            For curColumn = 1 To totalColumns
                totalsValue = totalsRng.Cells(curRow, curColumn).Value
                If totalsValue Like "*x" Then
                    If totalsValue = "x" Then
                        totalsFormulaString = "=round(" & unitRateAddressString & ",2)"
                    ElseIf totalsValue Like "*#x" Then
                        totalsFormulaString = "=round(" & unitRateAddressString & "*" _
                        & Left(totalsValue, Len(totalsValue) - 1) _
                        & ",2)"
                    End If
                    
                    totalsRng.Cells(curRow, curColumn).formula = totalsFormulaString
                End If
            Next curColumn
        End If
    Next curRow
End Sub

Sub tool9b_UnblindBudgetGrid()
'user selects two ranges and if ranges are valid, program runs to unblind the budget
    
    Dim unitRatesRng As Range
    Dim totalsRng As Range
    Dim selectTwoRangesReturn As Integer
    
    selectTwoRangesReturn = tool9_utilities.SelectTwoRanges(unitRatesRng, totalsRng, "unblind (convert x to $)")
    
    'if ranges are not valid, execution stops
    If selectTwoRangesReturn = 1 Then
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Call XToTotals(unitRatesRng, totalsRng)
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

End Sub

