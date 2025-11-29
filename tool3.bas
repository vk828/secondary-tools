Attribute VB_Name = "tool3"
'Author/Developer: Vadim Krifuks
'Collaborators: Hui Zeng, Man Ming Tse
'Last Updated: 20October2025

Option Explicit


'*****CONSTANTS*****
'input ranges
'row/column locations on a sheet where input ranges are stored
'data range is recalculated based on ranges provided for ids and visit names
'the other ranges can be put in manually or selected with an input box during program execution
Const column_ib As Integer = 8
Const column_sb As Integer = column_ib + 1

Const row_workbookName As Integer = 15
Const row_sheetName As Integer = row_workbookName + 1
Const row_idsRange As Integer = row_workbookName + 2
Const row_proceduresRange As Integer = row_workbookName + 3
Const row_negRatesRange As Integer = row_workbookName + 4
Const row_visitNamesRange As Integer = row_workbookName + 5
Const row_dataRange As Integer = row_workbookName + 6

'formulas
'row/column locations on a sheet where main formulas are put in
'formulas are updated every time the program executes
Const row_visitNamesFormula = 28
Const column_visitNamesFormula = 9

Const column_idsFormula = 5
Const column_proceduresFormula = column_idsFormula + 2
Const column_negRatesFormula = column_idsFormula + 3
Const column_gridTopLeftFormula = column_idsFormula + 4

Const row_ibSectionFiveFormulas = 101
Const row_sbSectionFiveFormulas = 500

Const rowsToClearInSectionFive = row_sbSectionFiveFormulas - row_ibSectionFiveFormulas - 1       'max field is 397 rows; we are clearing 398 rows
Const columnsToClear = 500    'arbitrary number

Sub tool3_HarmonizeIBwithSB()
'subroutine that executes the main program upon user's button click
'button may be called ADD/UPDATE FORMULAS

    Call IbSbTotalsAlignment
End Sub

Sub tool3_RemoveFormulas()
'subroutine to clear the sheet from all added formulas; initiated by a button click
'the button may be called REMOVE FORMULAS

    Call ClearSheet(ActiveSheet)
End Sub

Private Sub IbSbTotalsAlignment()
'main subroutine that makes user select ranges and add formulas to the
'appropriate sections on a sheet

    Dim toolSheet As Worksheet
    
    'the variable, buffer, is used to add grid formulas to a few additional rows and columns
    'to provide a seamless experience for the user who might add additional rows and columns
    'while aligning IB and SB
    Dim buffer As Integer

    Set toolSheet = ActiveSheet
    buffer = 10

    '*****STEP 1 - SELECT and ADD Data Ranges*****
    If Not SelectFourRangesAndSetDataRange(column_ib, "IB", toolSheet) Then
        Exit Sub
    End If
    
    If Not SelectFourRangesAndSetDataRange(column_sb, "SB", toolSheet) Then
        Exit Sub
    End If

    '*****STEP 2 - CLEAR CELLS*****
    Call ClearSheet(toolSheet)

    '*****STEP 3 - ADD FORMULA TO SECTION 1. Visit Names*****
    Call SetVisitNamesFormula(toolSheet)

    '*****STEP 4 - ADD FORMULAS TO SECTION 5. Raw Data: Internal and Sponsor Budgets*****
    Call SetSectionFive(row_ibSectionFiveFormulas, _
                                column_ib, _
                                toolSheet, _
                                buffer)

    Call SetSectionFive(row_sbSectionFiveFormulas, _
                                column_sb, _
                                toolSheet, _
                                buffer)
End Sub

Private Sub SetSectionFive(row_formula As Integer, _
                                column_input As Integer, _
                                toolSheet As Worksheet, _
                                buffer As Integer)
'this subroutine calls other subroutines to setup section 5

    Dim lastRow As Integer
    Dim lastColumn As Integer

    'STEP 1: add IDs formula
    Call SetIdsFormula(row_formula, column_input, toolSheet)
    
    'STEP 2: add Procedures formula
    Call SetProceduresFormula(row_formula, column_input, toolSheet)

    'STEP 3: add Negotiated Rates formula
    Call SetNegRatesFormula(row_formula, column_input, toolSheet)

    'STEP 4: determine lastRow and lastColumn
    'start from the lowest row in the section and find the last row going up
    With toolSheet

        If IsError(.Cells(row_formula, column_proceduresFormula)) Then
            With .Cells(row_formula, column_proceduresFormula)
                .value = "can't proceed until error in " & .Offset(0, -2).Address(False, False) & " is fixed"
                .WrapText = False
                .Font.color = RGB(255, 0, 0)
            End With
            GoTo Done
        ElseIf .Cells(row_formula, column_proceduresFormula).HasSpill And .Cells(row_formula, column_proceduresFormula).SpillingToRange.Columns.count = 1 Then
            lastRow = row_formula - 1 + .Cells(row_formula, column_proceduresFormula).SpillingToRange.Rows.count
        Else
            With .Cells(row_formula, column_gridTopLeftFormula)
                .value = "can't proceed untill values from " & .Offset(0, -2).Address(False, False) & " formula spill to multiple rows and stay within one column"
                .WrapText = False
                .Font.color = RGB(255, 0, 0)
            End With
            GoTo Done
        End If

    End With

    'adding buffer rows
    'make sure they don't go into next section
    If lastRow + buffer > row_formula + rowsToClearInSectionFive - 1 Then
        lastRow = row_formula + rowsToClearInSectionFive - 1
    Else
        lastRow = lastRow + buffer
    End If

    lastColumn = Cells(row_visitNamesFormula, column_visitNamesFormula).End(xlToRight).column

    'adding buffer to columns
    'if lastColumn jumps too far to the right, lastColumn = startColumn
    If lastColumn > 5000 Then
        lastColumn = column_visitNamesFormula
    Else
        lastColumn = lastColumn + buffer
    End If

    'STEP 5: add Grid formula
    Call SetGridFormula(row_formula, column_input, toolSheet)
    
    'STEP 6: add grid formulas to all visits, procedures and buffer
    Call Utilities.FillFormulas(toolSheet, row_formula, column_gridTopLeftFormula, lastRow, lastColumn)
Done:

End Sub

Private Sub SetVisitNamesFormula(toolSheet As Worksheet)
'this subroutine writes a formula to add visit names to section 1
'note that this formula has one implicit intersection operator: @sbOnlyVisitNamesArray
'it is added to make sure we are looking at the top left cell of an potential array
'if user selects multiple rows, the formula will take the first row ONLY

'=LET(
'    ibVisitNamesArray, TRIM(CLEAN(CHOOSEROWS('[Internal_Budget_25086_Vo_KinLET_vk_v5_19Oct2025.xlsx]Budget Details_P1_Coh1'!$AM$32:$CZ$33, 1))),
'    sbVisitNamesArray, TRIM(CLEAN(CHOOSEROWS('[25086 Sponsor Budget draft_vk_v5_19Oct25.xlsx]United States-Cohort 1'!$G$141:$DT$154,1))),
'    sbOnlyVisitNamesArray, UNIQUE(HSTACK(ibVisitNamesArray, ibVisitNamesArray, sbVisitNamesArray), TRUE, TRUE),
'    sbOnlyVisitNamesArrayEmpty, IFERROR(@sbOnlyVisitNamesArray, TRUE),
'    uniqueVisitNamesArray, IF(sbOnlyVisitNamesArrayEmpty = TRUE, ibVisitNamesArray, HSTACK(ibVisitNamesArray, sbOnlyVisitNamesArray)),
'    IF(uniqueVisitNamesArray = "", "", uniqueVisitNamesArray)
')

    Dim ibVisitNamesRange As String
    Dim sbVisitNamesRange As String
        
    ibVisitNamesRange = Utilities.AssembleRangeComponentsToAddressString(column_ib, row_workbookName, row_sheetName, row_visitNamesRange, toolSheet)
    sbVisitNamesRange = Utilities.AssembleRangeComponentsToAddressString(column_sb, row_workbookName, row_sheetName, row_visitNamesRange, toolSheet)
    
    Dim formula As String

    formula = "=LET(" & Chr(10) _
                    & String(4, Chr(32)) & "ibVisitNamesArray, TRIM(CLEAN(CHOOSEROWS(" & ibVisitNamesRange & ", 1)))," & Chr(10) _
                    & String(4, Chr(32)) & "sbVisitNamesArray, TRIM(CLEAN(CHOOSEROWS(" & sbVisitNamesRange & ",1)))," & Chr(10) _
                    & String(4, Chr(32)) & "sbOnlyVisitNamesArray, UNIQUE(HSTACK(ibVisitNamesArray, ibVisitNamesArray, sbVisitNamesArray), TRUE, TRUE)," & Chr(10) _
                    & String(4, Chr(32)) & "sbOnlyVisitNamesArrayEmpty, IFERROR(@sbOnlyVisitNamesArray, TRUE)," & Chr(10) _
                    & String(4, Chr(32)) & "uniqueVisitNamesArray, IF(sbOnlyVisitNamesArrayEmpty = TRUE, ibVisitNamesArray, HSTACK(ibVisitNamesArray, sbOnlyVisitNamesArray))," & Chr(10) _
                    & String(4, Chr(32)) & "IF(uniqueVisitNamesArray = """", """", uniqueVisitNamesArray)" & Chr(10) _
                & String(0, Chr(32)) & ")" & Chr(10)
                
    toolSheet.Cells(row_visitNamesFormula, column_visitNamesFormula).Formula2 = formula

End Sub

Private Sub SetIdsFormula(row_formula As Integer, _
                          column_input As Integer, _
                          toolSheet As Worksheet)
'this subroutine writes a formula to add IDs to section 5

'=LET(
'    errorMsg, CONCAT("Make sure workbook listed in H15 cell is open"),
'    id, IFERROR(TRIM(CLEAN('[Internal Budget_CC246510_Cheng_Merck_vk_v4.1_8Jan2025.xlsm]Budget_Details_ADJ_DBL'!$G$16:$G$198)), errorMsg),
'    IF(id = "",
'        "",
'        ID
'    )
')

    Dim idsRange As String
    idsRange = Utilities.AssembleRangeComponentsToAddressString(column_input, row_workbookName, row_sheetName, row_idsRange, toolSheet)

    Dim formula As String
    
    Dim inputCell As String
    
    inputCell = toolSheet.Cells(row_workbookName, column_input).Address(False, False)

    formula = "=LET(" & Chr(10) _
                    & String(4, Chr(32)) & "errorMsg, CONCAT(""Make sure workbook listed in " & inputCell & " cell is open"")," & Chr(10) _
                    & String(4, Chr(32)) & "id, IFERROR(TRIM(CLEAN(" & idsRange & ")), errorMsg)," & Chr(10) _
                    & String(4, Chr(32)) & "IF(id = """"," & Chr(10) _
                        & String(8, Chr(32)) & """""," & Chr(10) _
                        & String(8, Chr(32)) & "id" & Chr(10) _
                    & String(4, Chr(32)) & ")" & Chr(10) _
                & String(0, Chr(32)) & ")" & Chr(10)
                
    toolSheet.Cells(row_formula, column_idsFormula).Formula2 = formula

End Sub

Private Sub SetProceduresFormula(row_formula As Integer, _
                                 column_input As Integer, _
                                 toolSheet As Worksheet)
'this subroutine writes a formula to add Procedures to section 5

'=LET(
'    errorMsg, CONCAT("Make sure workbook listed in H15 cell is open"),
'    procedure, IFERROR(LEFT(TRIM(CLEAN('[Internal Budget_CC246510_Cheng_Merck_vk_v4.1_8Jan2025.xlsm]Budget_Details_ADJ_DBL'!$I$16:$I$198)),255),errorMsg),
'    IF(procedure = "",
'        "",
'        procedure
'    )
')

    Dim proceduresRange As String
    proceduresRange = Utilities.AssembleRangeComponentsToAddressString(column_input, row_workbookName, row_sheetName, row_proceduresRange, toolSheet)

    Dim formula As String
    
    Dim inputCell As String
    
    inputCell = toolSheet.Cells(row_workbookName, column_input).Address(False, False)
    
    formula = "=LET(" & Chr(10) _
                    & String(4, Chr(32)) & "errorMsg, CONCAT(""Make sure workbook listed in " & inputCell & " cell is open"")," & Chr(10) _
                    & String(4, Chr(32)) & "procedure, IFERROR(LEFT(TRIM(CLEAN(" & proceduresRange & ")),255),errorMsg)," & Chr(10) _
                    & String(4, Chr(32)) & "IF(procedure = """"," & Chr(10) _
                        & String(8, Chr(32)) & """""," & Chr(10) _
                        & String(8, Chr(32)) & "procedure" & Chr(10) _
                    & String(4, Chr(32)) & ")" & Chr(10) _
                & String(0, Chr(32)) & ")" & Chr(10)
                
    toolSheet.Cells(row_formula, column_proceduresFormula).Formula2 = formula

End Sub

Private Sub SetNegRatesFormula(row_formula As Integer, _
                               column_input As Integer, _
                               toolSheet As Worksheet)
'this subroutine writes a formula to add NegRates to section 5

'=LET(
'    errorMsg, CONCAT("Make sure workbook listed in H15 cell is open"),
'    negRate,
'        IFERROR(
'            IF(ISBLANK('[Internal Budget_CC246510_Cheng_Merck_vk_v4.1_8Jan2025.xlsm]Budget_Details_ADJ_DBL'!$L$16:$L$198),
'                "",
'                '[Internal Budget_CC246510_Cheng_Merck_vk_v4.1_8Jan2025.xlsm]Budget_Details_ADJ_DBL'!$L$16:$L$198
'            ),
'            errorMsg
'        ),
'    IF(negRate = "",
'        "",
'        negRate
'    )
')

    If toolSheet.Cells(row_negRatesRange, column_input) = "" Then
        toolSheet.Cells(row_formula, column_negRatesFormula).ClearContents
        GoTo Done
    End If
    
    Dim negotiatedRatesRange As String
        
    negotiatedRatesRange = Utilities.AssembleRangeComponentsToAddressString(column_input, row_workbookName, row_sheetName, row_negRatesRange, toolSheet)

    Dim formula As String
    
    Dim inputCell As String
    
    inputCell = toolSheet.Cells(row_workbookName, column_input).Address(False, False)
    
    formula = "=LET(" & Chr(10) _
                    & String(4, Chr(32)) & "errorMsg, CONCAT(""Make sure workbook listed in " & inputCell & " cell is open"")," & Chr(10) _
                    & String(4, Chr(32)) & "negRate," & Chr(10) _
                        & String(8, Chr(32)) & "IFERROR(" & Chr(10) _
                            & String(12, Chr(32)) & "if(isblank(" & negotiatedRatesRange & ")," & Chr(10) _
                                & String(16, Chr(32)) & """""," & Chr(10) _
                                & String(16, Chr(32)) & negotiatedRatesRange & Chr(10) _
                            & String(12, Chr(32)) & ")," & Chr(10) _
                            & String(12, Chr(32)) & "errorMsg" & Chr(10) _
                        & String(8, Chr(32)) & ")," & Chr(10) _
                    & String(4, Chr(32)) & "IF(negRate = """"," & Chr(10) _
                        & String(8, Chr(32)) & """""," & Chr(10) _
                        & String(8, Chr(32)) & "negRate" & Chr(10) _
                    & String(4, Chr(32)) & ")" & Chr(10) _
                & String(0, Chr(32)) & ")" & Chr(10)
                
    toolSheet.Cells(row_formula, column_negRatesFormula).Formula2 = formula
Done:

End Sub


Private Sub SetGridFormula(row_formula As Integer, _
                           column_input As Integer, _
                           toolSheet As Worksheet)
'this subroutine writes a formula to add to the top left corner of the grid in section 5

'=LET(
'    curIdCell, $E500,
'    curProcedureCell, $G500,
'    curVisitCell, I$499,
'    curNegRateCell,
'    IF(
'        ISNUMBER(SEARCH("inr", curIdCell)),
'        "",
'        $H500
'    ),
'
'    IF(OR(AND(curProcedureCell = "", curIdCell = ""), curVisitCell = ""),
'        "",
'        LET(
'            curId, TRIM(CLEAN(curIdCell)),
'            curProcedure, TRIM(CLEAN(curProcedureCell)),
'            curIdProcedure, LEFT(curId & curProcedure, 255),
'            curVisit, TRIM(CLEAN(curVisitCell)),
'            dataRange, '[25086 Sponsor Budget draft_vk_v5_19Oct25.xlsx]United States-Cohort 1'!$G$142:$DT$283,
'            idRange, TRIM(CLEAN('[25086 Sponsor Budget draft_vk_v5_19Oct25.xlsx]United States-Cohort 1'!$DV$142:$DV$283)),
'            procedureRange, TRIM(CLEAN('[25086 Sponsor Budget draft_vk_v5_19Oct25.xlsx]United States-Cohort 1'!$B$142:$B$283)),
'            idProcedureRange, LEFT(idRange & procedureRange, 255),
'            visitRange, TRIM(CLEAN(CHOOSEROWS('[25086 Sponsor Budget draft_vk_v5_19Oct25.xlsx]United States-Cohort 1'!$G$141:$DT$154, 1))),
'            indexRow, MATCH(curIdProcedure, idProcedureRange, 0),
'            indexColumn, MATCH(curVisit, visitRange, 0),
'            curTimePoint, INDEX(dataRange, indexRow, indexColumn),
'            rawOutput,
'            IF(
'                ISNA(curTimePoint),
'                "no result",
'                IF(
'                    curTimePoint = "",
'                    "",
'                    IF(
'                        curNegRateCell = "",
'                        curTimePoint,
'                        IF(
'                            AND(ISNUMBER(VALUE(curNegRateCell)), ISNUMBER(VALUE(curTimePoint))),
'                            curNegRateCell * curTimePoint,
'                            IF(
'                                AND(NOT(ISNUMBER(VALUE(curNegRateCell))), ISNUMBER(VALUE(curTimePoint))),
'                                IF(
'                                    curTimePoint > 1,
'                                    CONCAT(curTimePoint," x ", curNegRateCell),
'                                    Concat (curNegRateCell)
'                                ),
'                                IF(
'                                    AND(ISNUMBER(VALUE(curNegRateCell)), NOT(ISNUMBER(VALUE(curTimePoint)))),
'                                    CONCAT(curTimePoint, " @ ", DOLLAR(curNegRateCell)),
'                                    CONCAT(curTimePoint, " @ ", curNegRateCell)
'                                )
'                            )
'                        )
'                    )
'                )
'            ),
'            IF(
'                AND(ISNUMBER(SEARCH("inv", curId)), rawOutput <> "", NOT(ISNA(curTimePoint))),
'                IF(
'                    ISNUMBER(rawOutput),
'                    CONCAT("INV: ", DOLLAR(rawOutput)),
'                    CONCAT("INV: ", rawOutput)
'                ),
'                rawOutput
'            )
'        )
'    )
')
'indent     - String(12, Chr(32))
'new line   - Chr (10)
'empty string - """" or chr(34) & chr (34)


    Dim curIdCell, curProcedureCell, curNegRateCell As String
    Dim curVisitCell As String
    
    With toolSheet
        curIdCell = .Cells(row_formula, column_idsFormula).Address(RowAbsolute:=False)
        curProcedureCell = .Cells(row_formula, column_proceduresFormula).Address(RowAbsolute:=False)
        curNegRateCell = .Cells(row_formula, column_negRatesFormula).Address(RowAbsolute:=False)
        
        curVisitCell = .Cells(row_formula - 1, column_gridTopLeftFormula).Address(columnabsolute:=False)
    End With
    
    Dim idRange As String
    Dim procedureRange As String
    Dim visitRange As String
    Dim dataRange As String
    
    idRange = Utilities.AssembleRangeComponentsToAddressString(column_input, row_workbookName, row_sheetName, row_idsRange, toolSheet)
    procedureRange = Utilities.AssembleRangeComponentsToAddressString(column_input, row_workbookName, row_sheetName, row_proceduresRange, toolSheet)
    visitRange = Utilities.AssembleRangeComponentsToAddressString(column_input, row_workbookName, row_sheetName, row_visitNamesRange, toolSheet)
    dataRange = Utilities.AssembleRangeComponentsToAddressString(column_input, row_workbookName, row_sheetName, row_dataRange, toolSheet)
    
    Dim formula As String

    formula = "=LET(" & Chr(10) _
                & String(4, Chr(32)) & "curIdCell, " & curIdCell & "," & Chr(10) _
                & String(4, Chr(32)) & "curProcedureCell, " & curProcedureCell & "," & Chr(10) _
                & String(4, Chr(32)) & "curVisitCell, " & curVisitCell & "," & Chr(10) _
                & String(4, Chr(32)) & "curNegRateCell, " & Chr(10) _
                & String(4, Chr(32)) & "IF(" & Chr(10) _
                    & String(8, Chr(32)) & "ISNUMBER(SEARCH(""inr"", curIdCell))," & Chr(10) _
                    & String(8, Chr(32)) & """"" ," & Chr(10) _
                    & String(8, Chr(32)) & curNegRateCell & Chr(10) _
                & String(4, Chr(32)) & ")," & Chr(10) _
                & Chr(10) _
                & String(4, Chr(32)) & "IF(OR(AND(curProcedureCell = """", curIdCell = """"), curVisitCell = """")," & Chr(10) _
                    & String(8, Chr(32)) & """"" ," & Chr(10) _
                    & String(8, Chr(32)) & "LET(" & Chr(10)
    
    formula = formula _
                    & String(12, Chr(32)) & "curId, TRIM(CLEAN(curIdCell))," & Chr(10) _
                    & String(12, Chr(32)) & "curProcedure, TRIM(CLEAN(curProcedureCell))," & Chr(10) _
                    & String(12, Chr(32)) & "curIdProcedure, LEFT(curId & curProcedure, 255)," & Chr(10) _
                    & String(12, Chr(32)) & "curVisit, TRIM(CLEAN(curVisitCell))," & Chr(10) _
                    & String(12, Chr(32)) & "dataRange, " & dataRange & "," & Chr(10) _
                    & String(12, Chr(32)) & "idRange, TRIM(CLEAN(" & idRange & "))," & Chr(10) _
                    & String(12, Chr(32)) & "procedureRange, TRIM(CLEAN(" & procedureRange & "))," & Chr(10) _
                    & String(12, Chr(32)) & "idProcedureRange, LEFT(idRange & procedureRange, 255)," & Chr(10) _
                    & String(12, Chr(32)) & "visitRange, TRIM(CLEAN(CHOOSEROWS(" & visitRange & ", 1)))," & Chr(10) _
                    & String(12, Chr(32)) & "indexRow, MATCH(curIdProcedure, idProcedureRange, 0)," & Chr(10) _
                    & String(12, Chr(32)) & "indexColumn, MATCH(curVisit, visitRange, 0)," & Chr(10) _
                    & String(12, Chr(32)) & "curTimePoint, INDEX(dataRange, indexRow, indexColumn)," & Chr(10)
                    
    formula = formula _
                    & String(12, Chr(32)) & "rawOutput," & Chr(10) _
                    & String(12, Chr(32)) & "IF(" & Chr(10) _
                        & String(16, Chr(32)) & "ISNA(curTimePoint)," & Chr(10) _
                        & String(16, Chr(32)) & """no result""," & Chr(10) _
                        & String(16, Chr(32)) & "IF(" & Chr(10) _
                            & String(20, Chr(32)) & "curTimePoint = """"," & Chr(10) _
                            & String(20, Chr(32)) & """""," & Chr(10) _
                            & String(20, Chr(32)) & "IF(" & Chr(10) _
                                & String(24, Chr(32)) & "curNegRateCell = """"," & Chr(10) _
                                & String(24, Chr(32)) & "curTimePoint," & Chr(10) _
                                & String(24, Chr(32)) & "IF(" & Chr(10) _
                                    & String(28, Chr(32)) & "AND(ISNUMBER(VALUE(curNegRateCell)), ISNUMBER(VALUE(curTimePoint)))," & Chr(10) _
                                    & String(28, Chr(32)) & "curNegRateCell * curTimePoint," & Chr(10) _
                                    & String(28, Chr(32)) & "IF(" & Chr(10) _
                                    & String(32, Chr(32)) & "AND(NOT(ISNUMBER(VALUE(curNegRateCell))), ISNUMBER(VALUE(curTimePoint)))," & Chr(10) _
                                    & String(32, Chr(32)) & "IF(" & Chr(10) _
                                        & String(36, Chr(32)) & "curTimePoint > 1," & Chr(10) _
                                        & String(36, Chr(32)) & "CONCAT(curTimePoint,"" x "", curNegRateCell)," & Chr(10) _
                                        & String(36, Chr(32)) & "CONCAT(curNegRateCell)" & Chr(10) _
                                    & String(32, Chr(32)) & ")," & Chr(10)
                                    
    formula = formula _
                                    & String(32, Chr(32)) & "IF(" & Chr(10) _
                                        & String(36, Chr(32)) & "AND(ISNUMBER(VALUE(curNegRateCell)), NOT(ISNUMBER(VALUE(curTimePoint))))," & Chr(10) _
                                        & String(36, Chr(32)) & "CONCAT(curTimePoint, "" @ "", DOLLAR(curNegRateCell))," & Chr(10) _
                                        & String(36, Chr(32)) & "CONCAT(curTimePoint, "" @ "", curNegRateCell)" & Chr(10) _
                                    & String(32, Chr(32)) & ")" & Chr(10) _
                                & String(28, Chr(32)) & ")" & Chr(10) _
                            & String(24, Chr(32)) & ")" & Chr(10) _
                        & String(20, Chr(32)) & ")" & Chr(10) _
                    & String(16, Chr(32)) & ")" & Chr(10) _
                & String(12, Chr(32)) & ")," & Chr(10)
                
    formula = formula _
                & String(12, Chr(32)) & "IF(" & Chr(10) _
                    & String(16, Chr(32)) & "AND(ISNUMBER(SEARCH(""inv"", curId)), rawOutput <> """", NOT(ISNA(curTimePoint)))," & Chr(10) _
                    & String(16, Chr(32)) & "IF(" & Chr(10) _
                        & String(20, Chr(32)) & "ISNUMBER(rawOutput)," & Chr(10) _
                        & String(20, Chr(32)) & "CONCAT(""INV: "", DOLLAR(rawOutput))," & Chr(10) _
                        & String(20, Chr(32)) & "CONCAT(""INV: "", rawOutput)" & Chr(10) _
                    & String(16, Chr(32)) & ")," & Chr(10) _
                    & String(16, Chr(32)) & "rawOutput" & Chr(10) _
                & String(12, Chr(32)) & ")" & Chr(10) _
            & String(8, Chr(32)) & ")" & Chr(10) _
        & String(4, Chr(32)) & ")" & Chr(10) _
    & String(0, Chr(32)) & ")" & Chr(10) _

    toolSheet.Cells(row_formula, column_gridTopLeftFormula).Formula2 = formula

End Sub

Private Sub ClearSheet(toolSheet As Worksheet)
'this subroutine clears the sheet from added formulas

    With toolSheet
    
        'section 1 - visit names
        .Range(.Cells(row_visitNamesFormula, column_visitNamesFormula), _
                .Cells(row_visitNamesFormula, column_visitNamesFormula + columnsToClear - 1)) _
                .ClearContents

        'section 5 - IB
        .Range(.Cells(row_ibSectionFiveFormulas, column_idsFormula), _
                .Cells(row_ibSectionFiveFormulas + rowsToClearInSectionFive - 1, column_idsFormula)) _
                .ClearContents
                
        .Range(.Cells(row_ibSectionFiveFormulas, column_proceduresFormula), _
                .Cells(row_ibSectionFiveFormulas + rowsToClearInSectionFive - 1, column_proceduresFormula)) _
                .ClearContents
        
        .Range(.Cells(row_ibSectionFiveFormulas, column_negRatesFormula), _
                .Cells(row_ibSectionFiveFormulas + rowsToClearInSectionFive - 1, column_negRatesFormula)) _
                .ClearContents
        
        .Range(.Cells(row_ibSectionFiveFormulas, column_gridTopLeftFormula), _
                .Cells(row_ibSectionFiveFormulas + rowsToClearInSectionFive - 1, _
                        column_gridTopLeftFormula + columnsToClear - 1)) _
                .ClearContents

        'section 5 - SB
        .Range(.Cells(row_sbSectionFiveFormulas, column_idsFormula), _
                .Cells(row_sbSectionFiveFormulas + rowsToClearInSectionFive - 1, column_idsFormula)) _
                .ClearContents
                
        .Range(.Cells(row_sbSectionFiveFormulas, column_proceduresFormula), _
                .Cells(row_sbSectionFiveFormulas + rowsToClearInSectionFive - 1, column_proceduresFormula)) _
                .ClearContents
        
        .Range(.Cells(row_sbSectionFiveFormulas, column_negRatesFormula), _
                .Cells(row_sbSectionFiveFormulas + rowsToClearInSectionFive - 1, column_negRatesFormula)) _
                .ClearContents
        
        .Range(.Cells(row_sbSectionFiveFormulas, column_gridTopLeftFormula), _
                .Cells(row_sbSectionFiveFormulas + rowsToClearInSectionFive - 1, _
                        column_gridTopLeftFormula + columnsToClear - 1)) _
                .ClearContents
    End With
    
End Sub


Private Function SelectFourRangesAndSetDataRange(column_allComponents As Integer, _
                                                 source As String, _
                                                 toolSheet As Worksheet) As Boolean
'this function let's the user select four ranges and sets one based on the selection
'returns true if data range is set, false otherwise
'data range is set based on ids and visit names ranges

    Dim areNonDataRangesSet, isDataRangeSet As Boolean

    'Step1: call to let the user select four input ranges
    areNonDataRangesSet = SelectIdsProceduresNegRatesAndVisitNamesRanges(column_allComponents, _
                                                                         source, toolSheet)
    If areNonDataRangesSet Then
        'Step2: attempt to calculate and write data range. if it fails, the function is set to false
        isDataRangeSet = Utilities.SetDataRange(row_workbookName, row_sheetName, _
                                                row_visitNamesRange, row_idsRange, _
                                                row_dataRange, column_allComponents, toolSheet)
        
        'Step3: switch view back to toolsheet
        toolSheet.Activate
        
        SelectFourRangesAndSetDataRange = isDataRangeSet
    Else
        SelectFourRangesAndSetDataRange = False
    End If

End Function

Private Function SelectIdsProceduresNegRatesAndVisitNamesRanges(column_allComponents As Integer, _
                                                                source As String, _
                                                                toolSheet As Worksheet) As Boolean
'this function is responsible for having the user select and record components of
'ids, procedures, negotiated rates, and visit names ranges
'returns false if unsuccessful, true otherwise

    Dim idsRng, proceduresRng, negRatesRng, visitNamesRng As Range
    
    Dim title, titlePart1, titlePart2, titlePart3 As String
    Dim prompt, promptPart1, promptPart2, promptPart3, promptPart4, promptPart5, _
                                                                    promptPart6 As String
    SelectIdsProceduresNegRatesAndVisitNamesRanges = True
    
    'RANGE ONE - Sponsor IDs
    titlePart1 = "Select " & source & " "
    titlePart2 = "SPONSOR IDs"
    titlePart3 = " Range"
    title = titlePart1 & titlePart2 & titlePart3
    
    promptPart1 = "You have two options:" & Chr(10) & _
                  "  1) select a new " & source & " "
    promptPart2 = "Sponsor IDs"
    promptPart3 = " range (one "
    promptPart4 = "column up to 398 rows)"
    promptPart5 = " and click OK, or " & Chr(10) & _
                "  2) click Cancel to keep the old range"
    promptPart6 = Chr(10) & Chr(10) & "Note: Sponsor IDs, Procedures, and Neg Rates " _
                  & "ranges adjust to first columns and same start/end rows"
    prompt = promptPart1 & promptPart2 & promptPart3 & promptPart4 & promptPart5 & promptPart6
    
    'get a range input from the user
    Set idsRng = Utilities.SelectRange(title, prompt)
    
    If Not (idsRng Is Nothing) Then
        Call Utilities.WriteSelectedRangeComponentsToCells(row_workbookName, _
                                                            row_sheetName, _
                                                            column_allComponents, _
                                                            row_idsRange, _
                                                            idsRng, _
                                                            toolSheet)
                                                            
    ElseIf Not CheckWorkBookAndSheetAndActivateSheets(column_allComponents, toolSheet) Then
        SelectIdsProceduresNegRatesAndVisitNamesRanges = False
        Exit Function
    End If

    'RANGE TWO - Procedures
    titlePart2 = "PROCEDURES"
    title = titlePart1 & titlePart2 & titlePart3
    
    promptPart2 = Application.WorksheetFunction.Proper(titlePart2)
    prompt = promptPart1 & promptPart2 & promptPart3 & promptPart4 & promptPart5 & promptPart6
    
    'get a range input from the user
    Set proceduresRng = Utilities.SelectRange(title, prompt)

    If Not (proceduresRng Is Nothing) Then
        Call Utilities.WriteSelectedRangeComponentsToCells(row_workbookName, _
                                                            row_sheetName, _
                                                            column_allComponents, _
                                                            row_proceduresRange, _
                                                            proceduresRng, _
                                                            toolSheet)

    ElseIf Not CheckWorkBookAndSheetAndActivateSheets(column_allComponents, toolSheet) Then
        SelectIdsProceduresNegRatesAndVisitNamesRanges = False
        Exit Function
    End If
    
    'RANGE THREE - Negotiated Rates
    
    Dim answer As String
    Dim msgBoxPrompt As String
    
    msgBoxPrompt = "If the " & source & " grid shows totals and you decide against adding INR " _
                    & "to each individual Sponsor ID, you can click YES to set all Negotiated " & _
                    "Rates cells to empty " & Chr(10) & _
                    "  -YES for budgets that show TOTALS" & Chr(10) & _
                    "  -NO for budgets that show FREQUENCIES or both TOTALS and FREQUENCIES"

    answer = MsgBox(msgBoxPrompt, vbYesNo, _
                    title:=source & " - Ignore Negotiated Rates (INR)?")
    
    If answer = vbYes Then
        toolSheet.Cells(row_negRatesRange, column_allComponents).ClearContents
        GoTo afterNegRatesSelection
    End If
    
    titlePart2 = "NEGOTIATED RATES"
    title = titlePart1 & titlePart2 & titlePart3

    promptPart2 = Application.WorksheetFunction.Proper(titlePart2)
    prompt = promptPart1 & promptPart2 & promptPart3 & promptPart4 & promptPart5 & promptPart6
    
    'get a range input from the user
    Set negRatesRng = Utilities.SelectRange(title, prompt)

    If Not (negRatesRng Is Nothing) Then
        Call Utilities.WriteSelectedRangeComponentsToCells(row_workbookName, _
                                                            row_sheetName, _
                                                            column_allComponents, _
                                                            row_negRatesRange, _
                                                            negRatesRng, _
                                                            toolSheet)
    
    ElseIf Not CheckWorkBookAndSheetAndActivateSheets(column_allComponents, toolSheet) Then
        SelectIdsProceduresNegRatesAndVisitNamesRanges = False
        Exit Function
    End If

afterNegRatesSelection:
    
    'adjusts ranges so that the user doesn't have to worry about it
    'make single column ranges that start and end at the same rows
    Call AdjustIdsProceduresNegRatesRanges(column_allComponents, toolSheet)
    
    'RANGE FOUR - Visit Names
    titlePart2 = "SPONSOR VISIT NAMES"
    title = titlePart1 & titlePart2 & titlePart3
    
    promptPart2 = Application.WorksheetFunction.Proper(titlePart2)
    promptPart4 = "row up to 150 columns)"
    prompt = promptPart1 & promptPart2 & promptPart3 & promptPart4 & promptPart5
    
    'get a range input from the user
    Set visitNamesRng = Utilities.SelectRange(title, prompt)

    If Not (visitNamesRng Is Nothing) Then
        Call Utilities.WriteSelectedRangeComponentsToCells(row_workbookName, _
                                                            row_sheetName, _
                                                            column_allComponents, _
                                                            row_visitNamesRange, _
                                                            visitNamesRng, _
                                                            toolSheet)
    
    ElseIf Not CheckWorkBookAndSheetAndActivateSheets(column_allComponents, toolSheet) Then
        SelectIdsProceduresNegRatesAndVisitNamesRanges = False
    End If

End Function

Private Function GetWorkbookName(column As Integer, toolSheet As Worksheet) As String
'this function returns workbook name listed on the tool sheet
    
    GetWorkbookName = CStr(toolSheet.Cells(row_workbookName, column))
End Function

Private Function GetSheetName(column As Integer, toolSheet As Worksheet) As String
'this function returns sheet name listed on the tool sheet

    GetSheetName = CStr(toolSheet.Cells(row_sheetName, column))
End Function

Private Function CheckWorkBookAndSheetAndActivateSheets(column As Integer, toolSheet As Worksheet) As Boolean
'this function:
'  1) checks the validity of the workbook and worksheet,
'  2) activates the sheet of the selected range,
'  3) switches back to the toolSheet
'returns true if workbook and sheet are open, false otherwise

    Dim areValid As Boolean
    Dim workbookName, sheetName As String
       
    workbookName = GetWorkbookName(column, toolSheet)
    sheetName = GetSheetName(column, toolSheet)
    
    areValid = Utilities.AreWorkbookAndWorksheetValid(workbookName, sheetName)

    'if workbook and sheet names are valid activate the sheet and switch back to toolSheet
'    If areValid Then
'
'        'switch to sheet of the selection
'        Workbooks(workbookName).Worksheets(sheetName).Activate
'
'        'wait so that user can see where what sheet the selection is made on
'        Application.Wait (Now + TimeValue("0:00:02"))
'
'        'switch to the tool sheet
'        toolSheet.Activate
'
'    End If
    
    CheckWorkBookAndSheetAndActivateSheets = areValid

End Function

Private Sub AdjustIdsProceduresNegRatesRanges(column_allComponents As Integer, _
                                              toolSheet As Worksheet)
'this subroutine looks at three ranges and adjusts them to
'1) single columns and 2) same start and end rows

    Dim idsRng, proceduresRng, negRatesRng As Range
    Dim negRatesRngAddressString As String
    Dim column_idsRng, column_proceduresRng, column_negRatesRng As Integer
    Dim row_start, row_end As Integer
    
    With toolSheet
        Set idsRng = .Range(.Cells(row_idsRange, column_allComponents))
        Set proceduresRng = .Range(.Cells(row_proceduresRange, column_allComponents))
        
        negRatesRngAddressString = .Cells(row_negRatesRange, column_allComponents)
                
        column_idsRng = idsRng.Cells(1, 1).column
        column_proceduresRng = proceduresRng.Cells(1, 1).column
        
        'find start row
        If idsRng.Cells(1, 1).Row < proceduresRng.Cells(1, 1).Row Then
            row_start = idsRng.Cells(1, 1).Row
        Else
            row_start = proceduresRng.Cells(1, 1).Row
        End If

        'find end row
        If idsRng.Cells(idsRng.Rows.count, 1).Row > proceduresRng.Cells(proceduresRng.Rows.count, 1).Row Then
            
            row_end = idsRng.Cells(idsRng.Rows.count, 1).Row
        Else
            row_end = proceduresRng.Cells(proceduresRng.Rows.count, 1).Row
        End If

        'compare start and end rows against rows of negRatesRng if negRatesRng is provided
        If negRatesRngAddressString <> "" Then
            Set negRatesRng = .Range(.Cells(row_negRatesRange, column_allComponents))
            column_negRatesRng = negRatesRng.Cells(1, 1).column

            'find start row
            If row_start > negRatesRng.Cells(1, 1).Row Then
                row_start = negRatesRng.Cells(1, 1).Row
            End If

            'find end row
            If row_end < negRatesRng.Cells(negRatesRng.Rows.count, 1).Row Then
                row_end = negRatesRng.Cells(negRatesRng.Rows.count, 1).Row
            End If
            
            're-write negRatesRange
            .Cells(row_negRatesRange, column_allComponents) = _
                        Range(Cells(row_start, column_negRatesRng), _
                                Cells(row_end, column_negRatesRng)).Address(False, False)
            
        End If
    
        're-write idsRange and proceduresRange ranges
        .Cells(row_idsRange, column_allComponents) = _
                    Range(Cells(row_start, column_idsRng), _
                            Cells(row_end, column_idsRng)).Address(False, False)
        
        .Cells(row_proceduresRange, column_allComponents) = _
                    Range(Cells(row_start, column_proceduresRng), _
                            Cells(row_end, column_proceduresRng)).Address(False, False)
        
    
    End With
    
End Sub


