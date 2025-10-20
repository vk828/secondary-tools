Attribute VB_Name = "tool2"
'Author/Developer: Vadim Krifuks
'Collaborators: Man Ming Tse
'Last Updated: 1 October 2025

Option Explicit

'*****CONSTANTS*****
'input ranges
'row/column locations on a sheet where input ranges are stored
'data range is recalculated based on ranges provided for visit names and procedures
'the other ranges can be put in manually or selected with an input box during program execution
Const column_ib As Integer = 3

Const row_workbookName As Integer = 5
Const row_sheetName As Integer = row_workbookName + 1
Const row_proceduresRange As Integer = row_workbookName + 2
Const row_visitNamesRange As Integer = row_workbookName + 3
Const row_dataRange As Integer = row_workbookName + 4

Sub tool2_UpdateIntBdgtGridToOncore()
'main subroutine
    
    'sheet where the tool is located
    Dim toolSheet As Worksheet
    
    'Int Bdgt ranges
    Dim ib_visitsRng As Range, ib_proceduresRng As Range, ib_gridRng As Range
    
    'OnCore ranges
    Dim rngCollection As New Collection         'collection of oncore ranges that
    Dim oncore_visitsRng As Range, oncore_proceduresRng As Range, oncore_gridRng As Range
      
    'the assumption is that a user calls the program by clicking a button
    'located on the tool sheet
    Set toolSheet = ActiveSheet
    
    '### STEP 1 ###
    'get INTERNAL BUDGET RANGES
    'exit sub if not successful
    If Not tool2_selectIntBdgtComponents.SelectIntBdgtComponents(toolSheet, _
                                                                column_ib, _
                                                                row_workbookName, _
                                                                row_sheetName, _
                                                                row_proceduresRange, _
                                                                row_visitNamesRange) Then
        Exit Sub
    End If
    
    'set Int Bdgt ranges
    Call AssembleIntBdgtRanges(ib_visitsRng, _
                            ib_proceduresRng, _
                            ib_gridRng, _
                            toolSheet)
    
    'switch view to internal budget
    ib_gridRng.Worksheet.Parent.Activate
    ib_gridRng.Worksheet.Activate

    '### STEP 2 ###
    'check Visit and Procedure ranges for duplicates
    'if there are duplicates, highlight them and exit sub
    If tool2_duplicates.FindAndHighlightDuplicatesTwoRanges(ib_proceduresRng, ib_visitsRng) Then
        ActiveWindow.WindowState = xlMaximized
        MsgBox ("Please ensure all visit and procedure names are unique. Tool2 has identified " _
                & "duplicates, highlighted them in red, and halted execution. Kindly resolve " _
                & "these duplicates and rerun the tool. Thank you.")
        Exit Sub
    End If

    '### STEP 3 ###
    'convert billing grid to internal budget grid type and get ONCORE RANGES
    'exit sub if not successful
    Set rngCollection = tool2_oncore.GetOncoreRanges(ib_gridRng.Worksheet.name)
    
    If rngCollection.count = 0 Then Exit Sub
    
    Set oncore_proceduresRng = rngCollection(1)
    Set oncore_visitsRng = rngCollection(2)
    Set oncore_gridRng = rngCollection(3)
    
    '### STEP 4 ###
    'update visit and procedure names on the internal budget
    'exit sub if user terminates the program by clicking 'Exit'
    If tool2_overwriteNames.OverwriteProcedureAndVisitNames(ib_proceduresRng, _
                                                            ib_visitsRng, _
                                                            oncore_proceduresRng, _
                                                            oncore_visitsRng) Then Exit Sub
    '### STEP 5 ###
    'process internal budget grid
    Call ProcessGrids(ib_visitsRng, _
                            ib_proceduresRng, _
                            ib_gridRng, _
                            oncore_visitsRng, _
                            oncore_proceduresRng, _
                            oncore_gridRng)

End Sub

'Private Sub ProcessGrids(ib_visitsRng As Range, _
'                            ib_proceduresRng As Range, _
'                            ib_gridRng As Range, _
'                            oncore_visitsRng As Range, _
'                            oncore_proceduresRng As Range, _
'                            oncore_gridRng As Range)
''subroutine to process grids; takes three ranges from Int Bdgt and three from OnCore
'
'    Dim dateStampStr As String
'    Dim msg As String
'    dateStampStr = "[" & Format(Date, "ddmmmyy") & " tool2 execution] "
'
'    Dim ib_rows As Integer, ib_currRow As Integer, ib_columns As Integer, ib_currColumn As Integer
'    Dim oncore_rows As Integer, oncore_columns As Integer
'    Dim oncore_currRow As Variant, oncore_currColumn As Variant 'Variant type because Variant can hold CVErr that might be returned by Application.Match
'
'    Dim ib_value As Variant
'    Dim oncore_value As Variant
'
'    Dim procedureName As String
'    Dim visitName As String
'
'    Dim response As Integer
'
'    Dim fillColor_sameEmpty As Long             'cell was empty before and is empty now (curr oncore = prev int bdgt)
'    Dim fillColor_sameValue As Long             'cell before and now has the same value (curr oncore = prev int bdgt)
'    Dim fillColor_updatedToOncore As Long       'cell updated to OnCore value (prev int bdgt <> oncore; updated to oncore)
'    Dim fillColor_differentFromOncore As Long   'prev int bdgt value is kept; oncore is different. OR procedure/visit are not in oncore
'
'    Dim visitNotFoundArray() As Boolean            'declare dynamic array (no size yet)
'
'    fillColor_sameEmpty = RGB(217, 217, 217)            'grey
'    fillColor_sameValue = xlNone                        'none
'    fillColor_updatedToOncore = RGB(255, 255, 0)        'yellow
'    fillColor_differentFromOncore = RGB(187, 225, 250)  'misty blue
'
'    ib_rows = ib_proceduresRng.rows.count
'    ib_columns = ib_visitsRng.Columns.count
'
'    ReDim visitNotFoundArray(1 To ib_columns)      'set a size of the array; all elements are initialized to False
'
'    'loop through procedures on internal budget
'    For ib_currRow = 1 To ib_rows
'
'        'procedure name from internal budget
'        procedureName = Application.WorksheetFunction.Trim(Application.WorksheetFunction.Clean(ib_proceduresRng.Cells(ib_currRow, 1).Value))
'
'        'find row where the internal budget procedure is located on oncore document
'        'if not found, Application.Match returns CVErr
'        oncore_currRow = Application.Match(procedureName, oncore_proceduresRng, 0)
'
'        'if procedure is not found
'        If IsError(oncore_currRow) Then
'            Call tool2_cases.ProcedureNotFound(ib_proceduresRng.Cells(ib_currRow, 1), ib_gridRng.rows(ib_currRow), fillColor_differentFromOncore)
'            GoTo nextProcedure
'        End If
'
'        'loop through visits on internal budget
'        For ib_currColumn = 1 To ib_columns
'
'            'visit name from internal budget
'            visitName = Application.WorksheetFunction.Trim(Application.WorksheetFunction.Clean(ib_visitsRng.Cells(1, ib_currColumn).Value))
'
'            'find column where the internal budget visit is located on oncore document
'            'if not found, Application.Match returns CVErr
'            oncore_currColumn = Application.Match(visitName, oncore_visitsRng, 0)
'
'            'if visit is not found
'            If IsError(oncore_currColumn) Then
'
'                'process this one time per visit ONLY
'                If visitNotFoundArray(ib_currColumn) = False Then
'                    Call tool2_cases.VisitNotFound(ib_visitsRng.Cells(1, ib_currColumn), ib_gridRng.Columns(ib_currColumn), fillColor_differentFromOncore)
'                    visitNotFoundArray(ib_currColumn) = True
'                End If
'
'                GoTo nextVisit
'            End If
'
'            ib_value = ib_gridRng.Cells(ib_currRow, ib_currColumn).Value
'            oncore_value = oncore_gridRng.Cells(oncore_currRow, oncore_currColumn).Value
'
'            'case1 - ib = "" and oncore = ""
'            If ib_value = "" And oncore_value = "" Then
'
'                Call tool2_cases.PrevAndCurrEqualNothing(ib_gridRng.Cells(ib_currRow, ib_currColumn), fillColor_sameEmpty)
'
'            'case2 - ib = oncore
'            ElseIf ib_value = oncore_value Then
'
'                Call tool2_cases.PrevAndCurrEqualX(ib_gridRng.Cells(ib_currRow, ib_currColumn), fillColor_sameValue)
'
'            'case3 - ib = "" and oncore <> ""
'            ElseIf ib_value = "" And oncore_value <> "" Then
'
'                Call tool2_cases.PrevNothingCurrX(ib_gridRng.Cells(ib_currRow, ib_currColumn), fillColor_updatedToOncore, ib_value, oncore_value)
'
'            'case4 - ib = "inv" or "effort" and oncore = 1
'            ElseIf isPrevEffortCurrOne(ib_value, oncore_value) Or isPrevInvoiceCurrOne(ib_value, oncore_value) Then
'
'                Call tool2_cases.PrevEquivalentToOneCurrOne(ib_gridRng.Cells(ib_currRow, ib_currColumn), fillColor_sameValue, ib_value, oncore_value)
'
'            'case5 - ib <> oncore
'            Else
'
'                If tool2_cases.PrevXCurrY(ib_gridRng.Cells(ib_currRow, ib_currColumn), _
'                                            oncore_gridRng.Cells(oncore_currRow, oncore_currColumn), _
'                                            visitName, _
'                                            procedureName, _
'                                            fillColor_updatedToOncore, _
'                                            fillColor_differentFromOncore, _
'                                            ib_value, _
'                                            oncore_value) = 1 Then
'                    oncore_gridRng.Worksheet.Parent.Close SaveChanges:=False 'close workbook
'                    Exit Sub
'                End If
'
'            End If
'nextVisit:
'        Next ib_currColumn
'
'nextProcedure:
'    Next ib_currRow
'
'Call Done(oncore_gridRng)
'
'End Sub

Private Sub ProcessGrids(ib_visitsRng As Range, _
                            ib_proceduresRng As Range, _
                            ib_gridRng As Range, _
                            oncore_visitsRng As Range, _
                            oncore_proceduresRng As Range, _
                            oncore_gridRng As Range)
'subroutine to process grids; takes three ranges from Int Bdgt and three from OnCore

    Dim dateStampStr As String
    Dim msg As String
    dateStampStr = "[" & Format(Date, "ddmmmyy") & " tool2 execution] "
    
    Dim ib_rows As Integer, ib_currRow As Integer, ib_columns As Integer, ib_currColumn As Integer
    Dim oncore_rows As Integer, oncore_columns As Integer
    Dim oncore_currRow As Variant, oncore_currColumn As Variant 'Variant type because Variant can hold CVErr that might be returned by Application.Match
    
    Dim ib_value As Variant
    Dim oncore_value As Variant

    Dim procedureName As String
    Dim visitName As String
    
    Dim response As Integer
    
    Dim fillColor_sameEmpty As Long             'cell was empty before and is empty now (curr oncore = prev int bdgt)
    Dim fillColor_sameValue As Long             'cell before and now has the same value (curr oncore = prev int bdgt)
    Dim fillColor_updatedToOncore As Long       'cell updated to OnCore value (prev int bdgt <> oncore; updated to oncore)
    Dim fillColor_differentFromOncore As Long   'prev int bdgt value is kept; oncore is different. OR procedure/visit are not in oncore

    Dim visitNotFoundArray() As Boolean            'declare dynamic array (no size yet)

    fillColor_sameEmpty = RGB(217, 217, 217)            'grey
    fillColor_sameValue = xlNone                        'none
    fillColor_updatedToOncore = RGB(255, 255, 0)        'yellow
    fillColor_differentFromOncore = RGB(187, 225, 250)  'misty blue

    'ib_rows = ib_proceduresRng.rows.count
    'ib_columns = ib_visitsRng.Columns.count
    
    ReDim visitNotFoundArray(1 To ib_visitsRng.count)      'set a size of the array; all elements are initialized to False
    
    Dim hCell As Range, vCell As Range
    Dim visitNumber As Integer
    
    'loop through procedures on internal budget
    'For ib_currRow = 1 To ib_rows
    For Each hCell In ib_proceduresRng
        
        visitNumber = 1
        ib_currRow = hCell.row
        
        'procedure name from internal budget
        procedureName = Application.WorksheetFunction.Trim(Application.WorksheetFunction.Clean(hCell.Value))
        
        'find row where the internal budget procedure is located on oncore document
        'if not found, Application.Match returns CVErr
        oncore_currRow = Application.Match(procedureName, oncore_proceduresRng, 0)
        
        'if procedure is not found
        If IsError(oncore_currRow) Then
            Call tool2_cases.ProcedureNotFound(hCell, _
                    Application.Intersect(ib_gridRng, ib_gridRng.Worksheet.rows(ib_currRow)), _
                    fillColor_differentFromOncore)
            GoTo nextProcedure
        End If
        
        'loop through visits on internal budget
        For Each vCell In ib_visitsRng
            
            visitNumber = visitNumber + 1
            ib_currColumn = vCell.column
            
            'visit name from internal budget
            visitName = Application.WorksheetFunction.Trim(Application.WorksheetFunction.Clean(vCell.Value))
            
            'find column where the internal budget visit is located on oncore document
            'if not found, Application.Match returns CVErr
            oncore_currColumn = Application.Match(visitName, oncore_visitsRng, 0)
            
            'if visit is not found
            If IsError(oncore_currColumn) Then
                
                'process this one time per visit ONLY
                If visitNotFoundArray(visitNumber) = False Then
                    Call tool2_cases.VisitNotFound(vCell, _
                            Application.Intersect(ib_gridRng, ib_gridRng.Worksheet.Columns(ib_currColumn)), _
                            fillColor_differentFromOncore)
                    visitNotFoundArray(visitNumber) = True
                End If
                
                GoTo nextVisit
            End If
            
            ib_value = ib_gridRng.Worksheet.Cells(ib_currRow, ib_currColumn).Value
            oncore_value = oncore_gridRng.Cells(oncore_currRow, oncore_currColumn).Value

            'case1 - ib = "" and oncore = ""
            If ib_value = "" And oncore_value = "" Then

                Call tool2_cases.PrevAndCurrEqualNothing(ib_gridRng.Worksheet.Cells(ib_currRow, ib_currColumn), fillColor_sameEmpty)

            'case2 - ib = oncore
            ElseIf ib_value = oncore_value Then

                Call tool2_cases.PrevAndCurrEqualX(ib_gridRng.Worksheet.Cells(ib_currRow, ib_currColumn), fillColor_sameValue)

            'case3 - ib = "" and oncore <> ""
            ElseIf ib_value = "" And oncore_value <> "" Then

                Call tool2_cases.PrevNothingCurrX(ib_gridRng.Worksheet.Cells(ib_currRow, ib_currColumn), fillColor_updatedToOncore, ib_value, oncore_value)

            'case4 - ib = "inv" or "effort" and oncore = 1
            ElseIf isPrevEffortCurrOne(ib_value, oncore_value) Or isPrevInvoiceCurrOne(ib_value, oncore_value) Then

                Call tool2_cases.PrevEquivalentToOneCurrOne(ib_gridRng.Worksheet.Cells(ib_currRow, ib_currColumn), fillColor_sameValue, ib_value, oncore_value)

            'case5 - ib <> oncore
            Else
            
                If tool2_cases.PrevXCurrY(ib_gridRng.Worksheet.Cells(ib_currRow, ib_currColumn), _
                                            oncore_gridRng.Cells(oncore_currRow, oncore_currColumn), _
                                            visitName, _
                                            procedureName, _
                                            fillColor_updatedToOncore, _
                                            fillColor_differentFromOncore, _
                                            ib_value, _
                                            oncore_value) = 1 Then
                    oncore_gridRng.Worksheet.Parent.Close SaveChanges:=False 'close workbook
                    Exit Sub
                End If
            
            End If
nextVisit:
        Next vCell
    
nextProcedure:
    Next hCell
        
Call Done(oncore_gridRng)

End Sub


Private Sub Done(rng As Range)
'shows message that the program is done
    
    rng.Worksheet.Parent.Close SaveChanges:=False 'close workbook
    MsgBox ("Tool2 finished updating the grid.")

End Sub

'Private Function SelectRanges(toolSheet As Worksheet) As Boolean
''function selects two ranges and sets a data range; returns true if successful, false otherwise
'
'    'if data range can't be set, execution stops
'    If Not SelectIntBdgtRanges(toolSheet) Then
'        SelectRanges = False
'        Exit Function
'    End If
'
'    SelectRanges = True
'
'End Function
'
'Private Function SelectIntBdgtRanges(toolSheet As Worksheet) As Boolean
'
'    'Int Bdgt Related - select two ranges and set data range
'    'if data range can't be set, execution stops
'    If Not SelectTwoRangesAndSetDataRange(column_ib, "Internal Budget", toolSheet) Then
'        SelectIntBdgtRanges = False
'        Exit Function
'    End If
'
'    SelectIntBdgtRanges = True
'
'End Function

Private Function AssembleIntBdgtRanges(ByRef ib_visits_rng As Range, _
                                    ByRef ib_procedures_rng As Range, _
                                    ByRef ib_grid_rng As Range, _
                                    toolSheet As Worksheet)

    Set ib_visits_rng = Utilities.AssembleRangeComponentsToRange(column_ib, row_workbookName, row_sheetName, row_visitNamesRange, toolSheet)
    Set ib_procedures_rng = Utilities.AssembleRangeComponentsToRange(column_ib, row_workbookName, row_sheetName, row_proceduresRange, toolSheet)
    Set ib_grid_rng = Utilities.AssembleRangeComponentsToRange(column_ib, row_workbookName, row_sheetName, row_dataRange, toolSheet)

End Function

'Private Function SelectTwoRangesAndSetDataRange(column_allComponents As Integer, _
'                                                source As String, _
'                                                toolSheet As Worksheet) As Boolean
''this function let's the user select two ranges and sets the third one based on the first two
''returns true if the third range is set, false otherwise
'
'    Dim isDataRangeSet As Boolean
'
'    'Step1: call to let the user select two input ranges
'    Call SelectProceduresAndVisitNamesRanges(column_allComponents, source, toolSheet)
'
'    'Step2: attempt to calculate and write data range. if it fails, the function is set to false
'    isDataRangeSet = Utilities.SetDataRange(row_workbookName, row_sheetName, _
'                                            row_visitNamesRange, row_proceduresRange, _
'                                            row_dataRange, column_allComponents, toolSheet)
'
'    'Step3: switch view back to toolsheet
'    toolSheet.Parent.Activate
'
'    SelectTwoRangesAndSetDataRange = isDataRangeSet
'
'End Function
'
'
'Private Sub SelectProceduresAndVisitNamesRanges(column_allComponents As Integer, _
'                                                source As String, toolSheet As Worksheet)
''this subroutine is responsible for having the user select and record components of
''the procedures and visit names ranges
'
'    Dim visitNamesRng, proceduresRng As Range
'
'    Dim title, titlePart1, titlePart2, titlePart3 As String
'    Dim prompt, promptPart1, promptPart2, promptPart3, promptPart4, promptPart5 As String
'
'
'    titlePart1 = "Select " & source & " "
'    titlePart2 = "Procedures"
'    titlePart3 = " Range"
'    title = titlePart1 & titlePart2 & titlePart3
'
'    promptPart1 = "You have two options:" & Chr(10) & _
'                  "  1) select a new " & source & " "
'    promptPart2 = titlePart2
'    promptPart3 = " range (make sure you select ONE "
'    promptPart4 = "COLUMN"
'    promptPart5 = " [up to 500 cells] and click OK), or " & Chr(10) & _
'                "  2) click Cancel to keep the old range"
'    prompt = promptPart1 & promptPart2 & promptPart3 & promptPart4 & promptPart5
'
'    'get a range input from the user
'    Set proceduresRng = Utilities.SelectRange(title, prompt)
'
'    If Not (proceduresRng Is Nothing) Then
'        Call Utilities.WriteSelectedRangeComponentsToCells(row_workbookName, _
'                                                            row_sheetName, _
'                                                            column_allComponents, _
'                                                            row_proceduresRange, _
'                                                            proceduresRng, _
'                                                            toolSheet)
'    End If
'
'    titlePart2 = "Visit Names"
'
'    title = titlePart1 & titlePart2 & titlePart3
'
'    promptPart2 = titlePart2
'    promptPart4 = "ROW"
'    prompt = promptPart1 & promptPart2 & promptPart3 & promptPart4 & promptPart5
'
'    'get a range input from the user
'    Set visitNamesRng = Utilities.SelectRange(title, prompt)
'
'    If Not (visitNamesRng Is Nothing) Then
'        Call Utilities.WriteSelectedRangeComponentsToCells(row_workbookName, _
'                                                            row_sheetName, _
'                                                            column_allComponents, _
'                                                            row_visitNamesRange, _
'                                                            visitNamesRng, _
'                                                            toolSheet)
'    End If
'
'End Sub
 
Private Function isPrevInvoiceCurrOne(prev As Variant, curr As Variant) As Boolean

    If StrComp(CStr(prev), "inv", vbTextCompare) = 0 And IsNumeric(curr) Then
        If CInt(curr) = 1 Then
            isPrevInvoiceCurrOne = True
        Else
            isPrevInvoiceCurrOne = False
        End If
    Else
        isPrevInvoiceCurrOne = False
    End If

End Function

Private Function isPrevEffortCurrOne(prev As Variant, curr As Variant) As Boolean

    If StrComp(CStr(prev), "effort", vbTextCompare) = 0 And IsNumeric(curr) Then
        If CInt(curr) = 1 Then
            isPrevEffortCurrOne = True
        Else
            isPrevEffortCurrOne = False
        End If
    Else
        isPrevEffortCurrOne = False
    End If

End Function

