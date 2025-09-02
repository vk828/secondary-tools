Attribute VB_Name = "tool9a"
'Author/Developer: Vadim Krifuks
'Collaborators: Hui Zeng, Man Ming Tse
'Last Updated: 5Feb2025

Option Explicit

Sub TotalsToX(ByVal unitRatesRng As Range, ByVal totalsRng As Range)
' this subroutine turns totals into "x" effectively blinding the budget
' skips all blanks and non-numerical values
     
    Dim totalRows As Integer
    Dim totalColumns As Integer
    
    totalRows = unitRatesRng.rows.count
    totalColumns = totalsRng.Columns.count

    Dim curRow As Integer
    Dim curColumn As Integer
    
    Dim totalsValue As Variant
    Dim unitRate As Variant
    Dim frequency As Single
    Dim roundMultiple As Single
    
    'NOTE TO USER: adjust this multiple if you require different rounding
    '0.25 means that totals would round to 0.25x, 0.5x, 0.75x, x, 1.25x, etc.
    roundMultiple = 0.25
    
    For curRow = 1 To totalRows
    
        unitRate = unitRatesRng.Cells(curRow, 1).Value
        If IsNumeric(unitRate) And unitRate > 0 Then
            For curColumn = 1 To totalColumns
                totalsValue = totalsRng.Cells(curRow, curColumn).Value
                If IsNumeric(totalsValue) And totalsValue > 0 Then
                    'frequency is rounded
                    frequency = Application.WorksheetFunction.MRound(totalsValue / unitRate, roundMultiple)
                    If frequency = 1 Then
                        totalsRng.Cells(curRow, curColumn) = "x"
                    Else
                        totalsRng.Cells(curRow, curColumn) = Trim(str(frequency)) & "x"
                    End If
                End If
            Next curColumn
        End If
    Next curRow
End Sub

Sub tool9a_BlindBudgetGrid()
'user selects two ranges and if ranges are valid, program runs to blind the budget

    Dim unitRatesRng As Range
    Dim totalsRng As Range
    Dim selectTwoRangesReturn As Integer
    
    selectTwoRangesReturn = tool9_utilities.SelectTwoRanges(unitRatesRng, totalsRng, "blind (convert $ to x)")
    
    'if ranges are not valid, execution stops
    If selectTwoRangesReturn = 1 Then
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Call TotalsToX(unitRatesRng, totalsRng)
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

End Sub
