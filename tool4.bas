Attribute VB_Name = "tool4"
'Author/Developer: Vadim Krifuks
'Collaborators: Hui Zeng, Man Ming Tse
'Last Updated: 19January2025

Option Explicit


'*****CONSTANTS*****
'input ranges
'row/column locations on a sheet where input ranges are stored
'data range is recalculated based on ranges provided for visit names and procedures
'the other ranges can be put in manually or selected with an input box during program execution
Const column_ib As Integer = 3
Const column_onCore As Integer = column_ib + 1

Const row_workbookName As Integer = 5
Const row_sheetName As Integer = row_workbookName + 1
Const row_proceduresRange As Integer = row_workbookName + 2
Const row_visitNamesRange As Integer = row_workbookName + 3
Const row_dataRange As Integer = row_workbookName + 4

'equivalent pairs
'row/column locations on a sheet where equivalent pairs are stored
'this is adjusted by user manually to give flexibility
'there are times when IB may say Effort while OnCore may say 1 or RC and S1
'the pairs are equivalent
Const column_ibEquivalentPair As Integer = 9
Const column_onCoreEquivalentPair As Integer = column_ibEquivalentPair + 1

Const row_equivalentPair1 As Integer = 5
Const row_equivalentPair2 As Integer = row_equivalentPair1 + 1
Const row_equivalentPair3 As Integer = row_equivalentPair1 + 2
Const row_equivalentPair4 As Integer = row_equivalentPair1 + 3
Const row_equivalentPair5 As Integer = row_equivalentPair1 + 4
Const row_equivalentPair6 As Integer = row_equivalentPair1 + 5
Const row_equivalentPair7 As Integer = row_equivalentPair1 + 6
Const row_equivalentPair8 As Integer = row_equivalentPair1 + 7
Const row_equivalentPair9 As Integer = row_equivalentPair1 + 8
Const row_equivalentPair10 As Integer = row_equivalentPair1 + 9

'formulas
'row/column locations on a sheet where main formulas and counters are put in
'these are updated every time the program executes
Const column_visitNamesFormula = 3
Const column_proceduresFormula = column_visitNamesFormula - 1
Const column_gridFormula = column_visitNamesFormula
Const column_numbersVertical = column_visitNamesFormula - 2
Const column_numbersHorizontal = column_visitNamesFormula

Const row_visitNamesFormula = 19
Const row_proceduresFormula = row_visitNamesFormula + 1
Const row_gridFormula = row_proceduresFormula
Const row_numbersVertical = row_proceduresFormula
Const row_numbersHorizontal = row_visitNamesFormula - 2


Sub tool4_HarmonizeIBwithOnCore()
'user clicks the button to start the program execution
'the button may be called REBUILD GRID

    Call IbOncoreGridCompare
    
End Sub

Sub tool4_RemoveFormulas()
'user clicks the button to clear the field
'this button may be called CLEAR GRID and used before
'user decides to save the tool file to reduce the file size

    Call ClearCells(ActiveSheet)

End Sub

Private Sub IbOncoreGridCompare()
'main subroutine
    
    'sheet where the tool is located
    Dim toolSheet As Worksheet
    
    Dim lastRow, lastColumn As Integer
    
    'the assumption is that a user calls the program by clicking a button
    'located on the tool sheet
    Set toolSheet = ActiveSheet

    Call ClearCells(toolSheet)

    'IB Related - select two ranges and set data range
    'if data range can't be set, execution stops
    If Not SelectTwoRangesAndSetDataRange(column_ib, "Internal Budget", toolSheet) Then
        Exit Sub
    End If
    
    'OnCore Related - select two ranges and set data range
    'if data range can't be set, execution stops
    If Not SelectTwoRangesAndSetDataRange(column_onCore, "OnCore", toolSheet) Then
        Exit Sub
    End If
    
    'add formulas for visit names and procedures
    Call SetVisitNamesFormula(toolSheet)
    Call SetProceduresFormula(toolSheet)
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    'add main grid formula to the top left corner
    Call SetGridFormula(toolSheet)
    
    'if there are more than 1 procedure and 1 visit
    If Not (IsEmpty(Cells(row_proceduresFormula + 1, column_proceduresFormula).value) Or _
            IsEmpty(Cells(row_visitNamesFormula, column_visitNamesFormula + 1).value)) Then
        
        'find lastRow/lastColumn
        With toolSheet
            lastRow = .Cells(row_proceduresFormula, column_proceduresFormula).End(xlDown).row
            lastColumn = .Cells(row_visitNamesFormula, column_visitNamesFormula).End(xlToRight).column
        End With
        
        'copy/paste main formula to the entire grid
        Call Utilities.FillFormulas(toolSheet, row_gridFormula, column_gridFormula, lastRow, lastColumn)

        'recalculate field
        Application.Calculate
        
        'highlight all non empty cells on the grid
        Call HighlightNonEmptyCells(toolSheet, row_gridFormula, column_gridFormula, lastRow, lastColumn)

        'add formatting to the sheet
        Call FormatField(toolSheet, row_visitNamesFormula, column_proceduresFormula, lastRow, lastColumn)

        'add visit and procedure numbers
        Call addNumbers(toolSheet, row_numbersHorizontal, column_numbersHorizontal, row_numbersHorizontal, lastColumn)
        Call addNumbers(toolSheet, row_numbersVertical, column_numbersVertical, lastRow, column_numbersVertical)
    End If

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

End Sub

Private Function SelectTwoRangesAndSetDataRange(column_allComponents As Integer, _
                                                source As String, _
                                                toolSheet As Worksheet) As Boolean
'this function let's the user select two ranges and sets the third one based on the first two
'returns true if the third range is set, false otherwise

    Dim isDataRangeSet As Boolean

    'Step1: call to let the user select two input ranges
    Call SelectProceduresAndVisitNamesRanges(column_allComponents, source, toolSheet)
    
    'Step2: attempt to calculate and write data range. if it fails, the function is set to false
    isDataRangeSet = Utilities.SetDataRange(row_workbookName, row_sheetName, _
                                            row_visitNamesRange, row_proceduresRange, _
                                            row_dataRange, column_allComponents, toolSheet)
    
    'Step3: switch view back to toolsheet
    toolSheet.Parent.Activate
    
    SelectTwoRangesAndSetDataRange = isDataRangeSet

End Function

Private Sub HighlightNonEmptyCells(curSheet As Worksheet, firstRow As Integer, firstCol As Integer, ByVal lastRow As Integer, lastCol As Integer)
'this subroutine highlights cells that are not empty within a range

    Dim cel As Range
    Dim selectedRange As Range
        
    With curSheet
        Set selectedRange = .Range(.Cells(firstRow, firstCol), .Cells(lastRow, lastCol))
    End With
    
    'remove any old highlights from the selected cells
    'this is not needed for the entire range but is important for the top left corner cell
    selectedRange.Interior.color = xlNone
    
    For Each cel In selectedRange.Cells
        If cel.value <> "" Then
            cel.Interior.color = vbYellow
        End If
    Next cel

   
End Sub

Private Sub SetVisitNamesFormula(toolSheet As Worksheet)
'this subroutine adds formula for visit names to the toolSheet

    Dim visitNamesRangeString, formulaString As String
    Dim wkbCellLocation, wkshCellLocation, rngCellLocation As String
    
    visitNamesRangeString = AssembleRangeComponentsToAddressString(column_ib, row_workbookName, row_sheetName, _
                                                                        row_visitNamesRange, toolSheet)
    With toolSheet
        wkbCellLocation = .Cells(row_workbookName, column_ib).Address(False, False)
        wkshCellLocation = .Cells(row_sheetName, column_ib).Address(False, False)
        rngCellLocation = .Cells(row_visitNamesRange, column_ib).Address(False, False)
    
        formulaString = "=IFERROR(TRIM(CLEAN(" & visitNamesRangeString _
                        & ")),""Check workbook, worksheet, and visit names in cells " _
                        & wkbCellLocation & ", " _
                        & wkshCellLocation & ", and " _
                        & rngCellLocation & ", respectively, " _
                        & "and make sure the referenced Internal Budget file is open"")"

        .Cells(row_visitNamesFormula, column_visitNamesFormula).Formula2 = formulaString

    End With

End Sub

Private Sub SetProceduresFormula(toolSheet As Worksheet)
'this subroutine adds formula for procedures to the toolSheet

    Dim proceduresRangeString, formulaString As String
    Dim wkbCellLocation, wkshCellLocation, rngCellLocation As String
    
    proceduresRangeString = AssembleRangeComponentsToAddressString(column_onCore, row_workbookName, row_sheetName, _
                                                                        row_proceduresRange, toolSheet)
    With toolSheet
        wkbCellLocation = .Cells(row_workbookName, column_ib).Address(False, False)
        wkshCellLocation = .Cells(row_sheetName, column_ib).Address(False, False)
        rngCellLocation = .Cells(row_proceduresRange, column_ib).Address(False, False)
    
        formulaString = "=IFERROR(TRIM(CLEAN(" & proceduresRangeString _
                        & ")),""Check workbook, worksheet, and procedures in cells " _
                        & wkbCellLocation & ", " _
                        & wkshCellLocation & ", and " _
                        & rngCellLocation & ", respectively, " _
                        & "and make sure the referenced OnCore file is open"")"

        .Cells(row_proceduresFormula, column_proceduresFormula).Formula2 = formulaString

    End With

End Sub

Private Sub SetGridFormula(toolSheet As Worksheet)
'this subroutine adds the main grid formula to the toolSheet

'=IF(
'    OR(C$14="",$B15=""),
'    "",
'    LET(
'        curVisit, C$14,
'        curProcedure, $B15,
'        ibVisitsRange, RIGHT,
'        ibProceduresRange, RIGHT(FORMULATEXT($C$6),LEN(FORMULATEXT($C$6))-1),
'        ibSchedRange, RIGHT(FORMULATEXT($C$9),LEN(FORMULATEXT($C$9))-1),
'        oncoreVisitsRange, RIGHT(FORMULATEXT($D$5),LEN(FORMULATEXT($D$5))-1),
'        oncoreProceduresRange, RIGHT(FORMULATEXT($D$6),LEN(FORMULATEXT($D$6))-1),
'        oncoreSchedRange, RIGHT(FORMULATEXT($D$9),LEN(FORMULATEXT($D$9))-1),
'        visitNotFound, "visit N/A",
'        procedureNotFound, "procedure N/A",
'        visitAndProcedureNotFound, "visit & procedure N/A",
'        ibMatchRow, MATCH(curProcedure,TRIM(CLEAN(INDIRECT(ibProceduresRange))),0),
'        ibMatchColumn, MATCH(curVisit,TRIM(CLEAN(INDIRECT(ibVisitsRange))),0),
'        oncoreMatchRow, MATCH(curProcedure,TRIM(CLEAN(INDIRECT(oncoreProceduresRange))),0),
'        oncoreMatchColumn, MATCH(curVisit,TRIM(CLEAN(INDIRECT(oncoreVisitsRange))),0),
'        ibLookup, IF(AND(ISNA(ibMatchRow),ISNA(ibMatchColumn)),
'                                 visitAndProcedureNotFound,
'                                 IF(ISNA(MATCH(curProcedure,TRIM(CLEAN(INDIRECT(ibProceduresRange))),0)),
'                                     procedureNotFound,
'                                     IF(ISNA(MATCH(curVisit,TRIM(CLEAN(INDIRECT(ibVisitsRange))),0)),
'                                         visitNotFound,
'                                         INDEX(INDIRECT(TRIM(CLEAN(ibSchedRange))), ibMatchRow, ibMatchColumn)
'                                     )
'                                 )
'                             ),
'        oncoreLookup, IF(AND(ISNA(oncoreMatchRow),ISNA(oncoreMatchColumn)),
'                                            visitAndProcedureNotFound,
'                                            IF(ISNA(MATCH(curProcedure,TRIM(CLEAN(INDIRECT(oncoreProceduresRange))),0)),
'                                                procedureNotFound,
'                                                IF(ISNA(MATCH(curVisit,TRIM(CLEAN(INDIRECT(oncoreVisitsRange))),0)),
'                                                    visitNotFound,
'                                                    INDEX(INDIRECT(TRIM(CLEAN(oncoreSchedRange))), oncoreMatchRow, oncoreMatchColumn)
'                                                )
'                                            )
'                                        ),
'        output, CONCAT(
'                                             "IB: ",
'                                                  IF(
'                                                      LEN(ibLookup)=0,
'                                                      "[empty]",
'                                                      ibLookup
'                                              ),
'                                              CHAR(10),
'                                              "OnCore: ",
'                                               IF(
'                                                      LEN(oncoreLookup)=0,
'                                                      "[empty]",
'                                                      oncoreLookup
'                                              )
'                          ),
'        check1, AND(OR(ibLookup = visitNotFound, ibLookup = procedureNotFound, ibLookup = visitAndProcedureNotFound), OR(oncoreLookup = visitNotFound, oncoreLookup = procedureNotFound, oncoreLookup = visitAndProcedureNotFound)),
'        check2, AND(OR(ibLookup = visitNotFound, ibLookup = procedureNotFound, ibLookup = visitAndProcedureNotFound), OR(oncoreLookup = "", CONCAT(oncoreLookup) = "0")),
'        check3, AND(OR(oncoreLookup = visitNotFound, oncoreLookup = procedureNotFound, oncoreLookup = visitAndProcedureNotFound), OR(ibLookup = "", CONCAT(ibLookup) = "0")),
'        check4, ibLookup = oncoreLookup,
'        IF(OR(check1, check2, check3, check4),
'            "",
'            output
'        )
'    )
')


    Dim ibVisitsRange, ibProceduresRange, ibSchedRange, oncoreVisitsRange, oncoreProceduresRange, oncoreSchedRange As String
    Dim curVisit, curProcedure As String
    Dim ibEquivalent1, ibEquivalent2, ibEquivalent3, ibEquivalent4, ibEquivalent5, _
        ibEquivalent6, ibEquivalent7, ibEquivalent8, ibEquivalent9, ibEquivalent10 As String
    Dim onCoreEquivalent1, onCoreEquivalent2, onCoreEquivalent3, onCoreEquivalent4, onCoreEquivalent5, _
        onCoreEquivalent6, onCoreEquivalent7, onCoreEquivalent8, onCoreEquivalent9, onCoreEquivalent10 As String
    
    Dim formulaString As String
    
    ibVisitsRange = AssembleRangeComponentsToAddressString(column_ib, row_workbookName, row_sheetName, _
                                                                        row_visitNamesRange, toolSheet)
    ibProceduresRange = AssembleRangeComponentsToAddressString(column_ib, row_workbookName, row_sheetName, _
                                                                        row_proceduresRange, toolSheet)
    ibSchedRange = AssembleRangeComponentsToAddressString(column_ib, row_workbookName, row_sheetName, _
                                                                        row_dataRange, toolSheet)
    
    oncoreVisitsRange = AssembleRangeComponentsToAddressString(column_onCore, row_workbookName, row_sheetName, _
                                                                        row_visitNamesRange, toolSheet)
    oncoreProceduresRange = AssembleRangeComponentsToAddressString(column_onCore, row_workbookName, row_sheetName, _
                                                                        row_proceduresRange, toolSheet)
    oncoreSchedRange = AssembleRangeComponentsToAddressString(column_onCore, row_workbookName, row_sheetName, _
                                                                        row_dataRange, toolSheet)

    With toolSheet
        curVisit = .Cells(row_visitNamesFormula, column_visitNamesFormula).Address(True, False)
        curProcedure = .Cells(row_proceduresFormula, column_proceduresFormula).Address(False, True)
        ibEquivalent1 = .Cells(row_equivalentPair1, column_ibEquivalentPair).Address
        ibEquivalent2 = .Cells(row_equivalentPair2, column_ibEquivalentPair).Address
        ibEquivalent3 = .Cells(row_equivalentPair3, column_ibEquivalentPair).Address
        ibEquivalent4 = .Cells(row_equivalentPair4, column_ibEquivalentPair).Address
        ibEquivalent5 = .Cells(row_equivalentPair5, column_ibEquivalentPair).Address
        ibEquivalent6 = .Cells(row_equivalentPair6, column_ibEquivalentPair).Address
        ibEquivalent7 = .Cells(row_equivalentPair7, column_ibEquivalentPair).Address
        ibEquivalent8 = .Cells(row_equivalentPair8, column_ibEquivalentPair).Address
        ibEquivalent9 = .Cells(row_equivalentPair9, column_ibEquivalentPair).Address
        ibEquivalent10 = .Cells(row_equivalentPair10, column_ibEquivalentPair).Address
        onCoreEquivalent1 = .Cells(row_equivalentPair1, column_onCoreEquivalentPair).Address
        onCoreEquivalent2 = .Cells(row_equivalentPair2, column_onCoreEquivalentPair).Address
        onCoreEquivalent3 = .Cells(row_equivalentPair3, column_onCoreEquivalentPair).Address
        onCoreEquivalent4 = .Cells(row_equivalentPair4, column_onCoreEquivalentPair).Address
        onCoreEquivalent5 = .Cells(row_equivalentPair5, column_onCoreEquivalentPair).Address
        onCoreEquivalent6 = .Cells(row_equivalentPair6, column_onCoreEquivalentPair).Address
        onCoreEquivalent7 = .Cells(row_equivalentPair7, column_onCoreEquivalentPair).Address
        onCoreEquivalent8 = .Cells(row_equivalentPair8, column_onCoreEquivalentPair).Address
        onCoreEquivalent9 = .Cells(row_equivalentPair9, column_onCoreEquivalentPair).Address
        onCoreEquivalent10 = .Cells(row_equivalentPair10, column_onCoreEquivalentPair).Address
        
    End With
        
    formulaString = "=IF(" & Chr(10) _
                        & String(4, Chr(32)) & "OR(" & curVisit & "=""""," & curProcedure & "="""")," & Chr(10) _
                        & String(4, Chr(32)) & """""," & Chr(10) _
                        & String(4, Chr(32)) & "LET(" & Chr(10) _
                            & String(8, Chr(32)) & "curVisit, " & curVisit & "," & Chr(10) _
                            & String(8, Chr(32)) & "curProcedure, " & curProcedure & "," & Chr(10) _
                            & String(8, Chr(32)) & "ibVisitsRange, " & ibVisitsRange & "," & Chr(10) _
                            & String(8, Chr(32)) & "ibProceduresRange, " & ibProceduresRange & "," & Chr(10) _
                            & String(8, Chr(32)) & "ibSchedRange, " & ibSchedRange & "," & Chr(10) _
                            & String(8, Chr(32)) & "oncoreVisitsRange, " & oncoreVisitsRange & "," & Chr(10) _
                            & String(8, Chr(32)) & "oncoreProceduresRange, " & oncoreProceduresRange & "," & Chr(10) _
                            & String(8, Chr(32)) & "oncoreSchedRange, " & oncoreSchedRange & "," & Chr(10) _
                            & String(8, Chr(32)) & "visitNotFound, ""visit NOT found""," & Chr(10) _
                            & String(8, Chr(32)) & "procedureNotFound, ""procedure NOT found""," & Chr(10) _
                            & String(8, Chr(32)) & "visitAndProcedureNotFound, ""visit & procedure NOT found""," & Chr(10) _
                            & String(8, Chr(32)) & "ibMatchRow, MATCH(curProcedure,TRIM(CLEAN(ibProceduresRange)),0)," & Chr(10) _
                            & String(8, Chr(32)) & "ibMatchColumn, MATCH(curVisit,TRIM(CLEAN(ibVisitsRange)),0)," & Chr(10) _
                            & String(8, Chr(32)) & "oncoreMatchRow, MATCH(curProcedure,TRIM(CLEAN(oncoreProceduresRange)),0)," & Chr(10) _
                            & String(8, Chr(32)) & "oncoreMatchColumn, MATCH(curVisit,TRIM(CLEAN(oncoreVisitsRange)),0)," & Chr(10)
                            
    formulaString = formulaString _
                            & String(8, Chr(32)) & "ibLookup, IF(AND(ISNA(ibMatchRow),ISNA(ibMatchColumn))," & Chr(10) _
                                & String(12, Chr(32)) & "visitAndProcedureNotFound," & Chr(10) _
                                & String(12, Chr(32)) & "IF(ISNA(MATCH(curProcedure,TRIM(CLEAN(ibProceduresRange)),0))," & Chr(10) _
                                    & String(16, Chr(32)) & "procedureNotFound," & Chr(10) _
                                    & String(16, Chr(32)) & "IF(ISNA(MATCH(curVisit,TRIM(CLEAN(ibVisitsRange)),0))," & Chr(10) _
                                        & String(20, Chr(32)) & "visitNotFound," & Chr(10) _
                                        & String(20, Chr(32)) & "INDEX(TRIM(CLEAN(ibSchedRange)), ibMatchRow, ibMatchColumn)" & Chr(10) _
                                    & String(16, Chr(32)) & ")" & Chr(10) _
                                & String(12, Chr(32)) & ")" & Chr(10) _
                            & String(8, Chr(32)) & ")," & Chr(10)

    formulaString = formulaString _
                        & String(8, Chr(32)) & "oncoreLookup, IF(AND(ISNA(oncoreMatchRow),ISNA(oncoreMatchColumn))," & Chr(10) _
                            & String(12, Chr(32)) & "visitAndProcedureNotFound," & Chr(10) _
                            & String(12, Chr(32)) & "IF(ISNA(MATCH(curProcedure,TRIM(CLEAN(oncoreProceduresRange)),0))," & Chr(10) _
                                & String(16, Chr(32)) & "procedureNotFound," & Chr(10) _
                                & String(16, Chr(32)) & "IF(ISNA(MATCH(curVisit,TRIM(CLEAN(oncoreVisitsRange)),0))," & Chr(10) _
                                    & String(20, Chr(32)) & "visitNotFound," & Chr(10) _
                                    & String(20, Chr(32)) & "INDEX(TRIM(CLEAN(oncoreSchedRange)), oncoreMatchRow, oncoreMatchColumn)" & Chr(10) _
                                & String(16, Chr(32)) & ")" & Chr(10) _
                            & String(12, Chr(32)) & ")" & Chr(10) _
                        & String(8, Chr(32)) & ")," & Chr(10)

    formulaString = formulaString _
                        & String(8, Chr(32)) & "output, CONCAT(" & Chr(10) _
                            & String(12, Chr(32)) & """IB: ""," & Chr(10) _
                            & String(12, Chr(32)) & "IF(LEN(ibLookup)=0,""[empty]"",ibLookup)," & Chr(10) _
                            & String(12, Chr(32)) & "CHAR(10)," & Chr(10) _
                            & String(12, Chr(32)) & """OnCore: ""," & Chr(10) _
                            & String(12, Chr(32)) & "IF(LEN(oncoreLookup)=0,""[empty]"",oncoreLookup)" & Chr(10) _
                        & String(8, Chr(32)) & ")," & Chr(10)
    
    'as of 14Dec24 check1 is removed from IF(OR(check1, check2, check3, check4)
    formulaString = formulaString _
                        & String(8, Chr(32)) & "check1, AND(OR(ibLookup = visitNotFound, ibLookup = procedureNotFound, ibLookup = visitAndProcedureNotFound), OR(oncoreLookup = visitNotFound, oncoreLookup = procedureNotFound, oncoreLookup = visitAndProcedureNotFound))," & Chr(10) _
                        & String(8, Chr(32)) & "check2, AND(OR(ibLookup = visitNotFound, ibLookup = procedureNotFound, ibLookup = visitAndProcedureNotFound), OR(oncoreLookup = """", CONCAT(oncoreLookup) = ""0""))," & Chr(10) _
                        & String(8, Chr(32)) & "check3, AND(OR(oncoreLookup = visitNotFound, oncoreLookup = procedureNotFound, oncoreLookup = visitAndProcedureNotFound), OR(ibLookup = """", CONCAT(ibLookup) = ""0""))," & Chr(10) _
                        & String(8, Chr(32)) & "check4, ibLookup = oncoreLookup," & Chr(10) _
                        & String(8, Chr(32)) & "check5, and(ibLookup = lower(" & ibEquivalent1 & "), oncoreLookup = lower(" & onCoreEquivalent1 & "))," & Chr(10) _
                        & String(8, Chr(32)) & "check6, and(ibLookup = lower(" & ibEquivalent2 & "), oncoreLookup = lower(" & onCoreEquivalent2 & "))," & Chr(10) _
                        & String(8, Chr(32)) & "check7, and(ibLookup = lower(" & ibEquivalent3 & "), oncoreLookup = lower(" & onCoreEquivalent3 & "))," & Chr(10) _
                        & String(8, Chr(32)) & "check8, and(ibLookup = lower(" & ibEquivalent4 & "), oncoreLookup = lower(" & onCoreEquivalent4 & "))," & Chr(10) _
                        & String(8, Chr(32)) & "check9, and(ibLookup = lower(" & ibEquivalent5 & "), oncoreLookup = lower(" & onCoreEquivalent5 & "))," & Chr(10) _
                        & String(8, Chr(32)) & "check10, and(ibLookup = lower(" & ibEquivalent6 & "), oncoreLookup = lower(" & onCoreEquivalent6 & "))," & Chr(10) _
                        & String(8, Chr(32)) & "check11, and(ibLookup = lower(" & ibEquivalent7 & "), oncoreLookup = lower(" & onCoreEquivalent7 & "))," & Chr(10) _
                        & String(8, Chr(32)) & "check12, and(ibLookup = lower(" & ibEquivalent8 & "), oncoreLookup = lower(" & onCoreEquivalent8 & "))," & Chr(10) _
                        & String(8, Chr(32)) & "check13, and(ibLookup = lower(" & ibEquivalent9 & "), oncoreLookup = lower(" & onCoreEquivalent9 & "))," & Chr(10) _
                        & String(8, Chr(32)) & "check14, and(ibLookup = lower(" & ibEquivalent10 & "), oncoreLookup = lower(" & onCoreEquivalent10 & "))," & Chr(10) _
                        & String(8, Chr(32)) & "IF(OR(check2, check3, check4, check5, check6, check7, check8, check9, check10, check11, check12, check13, check14)," & Chr(10) _
                            & String(12, Chr(32)) & """""," & Chr(10) _
                            & String(12, Chr(32)) & "output" & Chr(10) _
                        & String(8, Chr(32)) & ")" & Chr(10) _
                    & String(4, Chr(32)) & ")" & Chr(10) _
                & String(0, Chr(32)) & ")" & Chr(10) _

    toolSheet.Cells(row_gridFormula, column_gridFormula).Formula2 = formulaString

End Sub

Private Sub SelectProceduresAndVisitNamesRanges(column_allComponents As Integer, _
                                                source As String, toolSheet As Worksheet)
'this subroutine is responsible for having the user select and record components of
'the procedures and visit names ranges

    Dim visitNamesRng, proceduresRng As Range
    
    Dim title, titlePart1, titlePart2, titlePart3 As String
    Dim prompt, promptPart1, promptPart2, promptPart3, promptPart4, promptPart5 As String
    
    
    titlePart1 = "Select " & source & " "
    titlePart2 = "Procedures"
    titlePart3 = " Range"
    title = titlePart1 & titlePart2 & titlePart3
    
    promptPart1 = "You have two options:" & Chr(10) & _
                  "  1) select a new " & source & " "
    promptPart2 = titlePart2
    promptPart3 = " range (make sure you select ONE "
    promptPart4 = "COLUMN"
    promptPart5 = " [up to 500 cells] and click OK), or " & Chr(10) & _
                "  2) click Cancel to keep the old range"
    prompt = promptPart1 & promptPart2 & promptPart3 & promptPart4 & promptPart5
    
    'get a range input from the user
    Set proceduresRng = Utilities.SelectRange(title, prompt)

    If Not (proceduresRng Is Nothing) Then
        Call Utilities.WriteSelectedRangeComponentsToCells(row_workbookName, _
                                                            row_sheetName, _
                                                            column_allComponents, _
                                                            row_proceduresRange, _
                                                            proceduresRng, _
                                                            toolSheet)
    End If

    titlePart2 = "Visit Names"

    title = titlePart1 & titlePart2 & titlePart3
    
    promptPart2 = titlePart2
    promptPart4 = "ROW"
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
    End If

End Sub

Private Sub FormatField(toolSheet As Worksheet, sourceRow As Integer, sourceColumn As Integer, _
                        ByVal endRow As Integer, endColumn As Integer)
'this subroutine adds formating to the field
    
    With toolSheet
        'applies to all field
        With .Range(.Cells(sourceRow, sourceColumn), .Cells(endRow, endColumn))
            
            .WrapText = True
            
            With .Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .ColorIndex = xlAutomatic
                .TintAndShade = 0
                .Weight = xlMedium
            End With
            With .Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .ColorIndex = xlAutomatic
                .TintAndShade = 0
                .Weight = xlMedium
            End With
            With .Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .ColorIndex = xlAutomatic
                .TintAndShade = 0
                .Weight = xlMedium
            End With
            With .Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .ColorIndex = xlAutomatic
                .TintAndShade = 0
                .Weight = xlMedium
            End With
            With .Borders(xlInsideVertical)
                .LineStyle = xlContinuous
                .ColorIndex = xlAutomatic
                .TintAndShade = 0
                .Weight = xlMedium
            End With
            With .Borders(xlInsideHorizontal)
                .LineStyle = xlContinuous
                .ColorIndex = xlAutomatic
                .TintAndShade = 0
                .Weight = xlThin
            End With
            
        End With
        
        'applies to visit names ONLY
        With .Range(.Cells(sourceRow, sourceColumn), .Cells(sourceRow, endColumn))
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlBottom
            .WrapText = True
        End With
        
        'applies to procedures names ONLY
        With .Range(.Cells(sourceRow, sourceColumn), .Cells(endRow, sourceColumn))
            .Font.Bold = True
            .WrapText = True
        End With
    End With
    


End Sub
Private Sub addNumbers(toolSheet As Worksheet, sourceRow As Integer, sourceColumn As Integer, _
                        ByVal endRow As Integer, endColumn As Integer)
'this subroutine adds numbers from 1 to n to a provided range on a sheet

    Dim count As Integer
    Dim i As Integer
    Dim vertical As Boolean
    
    vertical = False
    
    'determine whether this is a vertical or horizontal range
    'and find n; n is count
    If (endRow - sourceRow) > 0 Then
        count = endRow - sourceRow + 1
        vertical = True
    Else
        count = endColumn - sourceColumn + 1
        
    End If
    
    
    If vertical Then
        With toolSheet
            .Range(.Cells(sourceRow, sourceColumn), .Cells(sourceRow + 500, sourceColumn)).Clear
            For i = 1 To count
                .Cells(sourceRow - 1 + i, sourceColumn).value = i
            Next
        End With
    Else
        With toolSheet
            .Range(.Cells(sourceRow, sourceColumn), .Cells(sourceRow, sourceColumn + 500)).Clear
            For i = 1 To count
                .Cells(sourceRow, sourceColumn - 1 + i).value = i
            Next
        End With
    End If
        
End Sub


Private Sub ClearCells(toolSheet As Worksheet)
'this subroutine clears field

    Dim cellsToClear As Integer
    cellsToClear = 500

    With toolSheet
        'clear from the second visit down and to the right
        .Range(.Cells(row_numbersHorizontal, column_numbersHorizontal + 1), _
                .Cells(row_numbersHorizontal + cellsToClear, column_numbersHorizontal + cellsToClear)).Clear
        
        'clear from the second procedure down and to the right
        .Range(.Cells(row_numbersVertical + 1, column_numbersVertical), _
                .Cells(row_numbersVertical + cellsToClear, column_numbersVertical + cellsToClear)).Clear

        'clear contents ONLY of the cell containing Visit Names formula
        .Cells(row_visitNamesFormula, column_visitNamesFormula).ClearContents
        
        'clear contents ONLY of the cell containing Procedures formula
        .Cells(row_proceduresFormula, column_proceduresFormula).ClearContents
        
        'clear contents ONLY of the cell containing top left Grid formula
        .Cells(row_gridFormula, column_gridFormula).ClearContents
        
    End With

End Sub


