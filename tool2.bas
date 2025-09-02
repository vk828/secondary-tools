Attribute VB_Name = "tool2"
'Author/Developer: Vadim Krifuks
'Collaborators: Man Ming Tse
'Last Updated: 12July2025

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

Sub tool2_UpdateIntBdgtGridToUserDefinedGrid()
'user clicks the button to start the program execution
'the button may be called UPDATE INTERNAL BUDGET GRID to A GRID
'user selects four ranges

    Call UpdateIntBdgtGridToUserDefinedGrid
    
End Sub

Sub tool2_UpdateIntBdgtGridToOncore()
'user clicks the button to start the program execution
'the button may be called UPDATE INTERNAL BUDGET GRID to ONCORE
'user selects
'  1) two ranges on the internal budget file
'  2) billing grid file
'  3) an arm within the billing grid file

    Call tool2_oncore.UpdateIntBdgtGridToOncore
    
End Sub

Private Sub UpdateIntBdgtGridToUserDefinedGrid()
'main subroutine
    
    'sheet where the tool is located
    Dim toolSheet As Worksheet
    
    Dim ib_visitsRng As Range, ib_proceduresRng As Range, ib_gridRng As Range
    Dim oncore_visitsRng As Range, oncore_proceduresRng As Range, oncore_gridRng As Range
  
    Dim oncore_visitsArr As Variant, oncore_proceduresArr As Variant
  
    'the assumption is that a user calls the program by clicking a button
    'located on the tool sheet
    Set toolSheet = ActiveSheet
    
    If Not SelectRanges(toolSheet) Then
        Exit Sub
    End If
    
    Call AssembleSixRanges(ib_visitsRng, _
                            ib_proceduresRng, _
                            ib_gridRng, _
                            oncore_visitsRng, _
                            oncore_proceduresRng, _
                            oncore_gridRng, _
                            toolSheet)
    
    'switch to internal budget
    ib_gridRng.Worksheet.Parent.Activate
    ib_gridRng.Worksheet.Activate
    
    oncore_visitsArr = ConvertRangeToArray(oncore_visitsRng)
    oncore_proceduresArr = ConvertRangeToArray(oncore_proceduresRng)
    
' NOTE TO VADIM: replace oncore_visitsRng and oncore_proceduresRng with oncore_visitsArr and oncore_proceduresArr
    
    Call ProcessGrids(ib_visitsRng, _
                            ib_proceduresRng, _
                            ib_gridRng, _
                            oncore_visitsRng, _
                            oncore_proceduresRng, _
                            oncore_gridRng)
    
End Sub

Private Function ConvertRangeToArray(rng) As Variant

    Dim arr() As Variant
    arr = rng.Value
    
    Dim rows As Long, cols As Long
    Dim row As Long, col As Long

    rows = UBound(arr, 1) - LBound(arr, 1) + 1  ' Number of rows
    cols = UBound(arr, 2) - LBound(arr, 2) + 1  ' Number of columns

    'iterate through columns and clean up the names
    For col = 1 To cols
        arr(1, col) = Left(Trim(Application.WorksheetFunction.Clean(CStr(arr(1, col)))), 255)
    Next col

    'iterate through rows and clean up the names
    For row = 1 To rows
        arr(row, 1) = Left(Trim(Application.WorksheetFunction.Clean(CStr(arr(row, 1)))), 255)
    Next row

    ConvertRangeToArray = arr

End Function

Private Sub ProcessGrids(ib_visitsRng As Range, _
                            ib_proceduresRng As Range, _
                            ib_gridRng As Range, _
                            oncore_visitsRng As Range, _
                            oncore_proceduresRng As Range, _
                            oncore_gridRng As Range)
    
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

    ib_rows = ib_proceduresRng.rows.count
    ib_columns = ib_visitsRng.Columns.count
    
    ReDim visitNotFoundArray(1 To ib_columns)      'set a size of the array; all elements are initialized to False
    
    'loop through procedures on internal budget
    For ib_currRow = 1 To ib_rows
        
        'procedure name from internal budget
        procedureName = ib_proceduresRng.Cells(ib_currRow, 1).Value
        
        'find row where the internal budget procedure is located on oncore document
        'if not found, Application.Match returns CVErr
        oncore_currRow = Application.Match(procedureName, oncore_proceduresRng, 0)
        
        'if procedure is not found
        If IsError(oncore_currRow) Then
            ib_proceduresRng.Cells(ib_currRow, 1).Interior.color = fillColor_differentFromOncore
            Call AddComment(ib_proceduresRng.Cells(ib_currRow, 1), dateStampStr & "procedure not found in OnCore; row skipped")
            
            GoTo nextProcedure
        End If
        
        'loop through visits on internal budget
        For ib_currColumn = 1 To ib_columns
            
            'visit name from internal budget
            visitName = ib_visitsRng.Cells(1, ib_currColumn).Value
            
            'find column where the internal budget visit is located on oncore document
            'if not found, Application.Match returns CVErr
            oncore_currColumn = Application.Match(visitName, oncore_visitsRng, 0)
            
            'if visit is not found
            If IsError(oncore_currColumn) Then
                
                ib_visitsRng.Cells(1, ib_currColumn).Interior.color = fillColor_differentFromOncore

                'comment is added only one time
                If visitNotFoundArray(ib_currColumn) = False Then
                    Call AddComment(ib_visitsRng.Cells(1, ib_currColumn), dateStampStr & "visit not found in OnCore; column skipped")
                    visitNotFoundArray(ib_currColumn) = True
                End If
                
                GoTo nextVisit
            End If
            
            ib_value = ib_gridRng.Cells(ib_currRow, ib_currColumn).Value
            oncore_value = oncore_gridRng.Cells(oncore_currRow, oncore_currColumn).Value

            'case1 - ib = "" and oncore = ""
            If ib_value = "" And oncore_value = "" Then

                Call tool2_cases.PrevAndCurrEqualNothing(ib_gridRng.Cells(ib_currRow, ib_currColumn), fillColor_sameEmpty)

            'case2 - ib = oncore
            ElseIf ib_value = oncore_value Then

                Call tool2_cases.PrevAndCurrEqualX(ib_gridRng.Cells(ib_currRow, ib_currColumn), fillColor_sameValue)

            'case3 - ib = "" and oncore <> ""
            ElseIf ib_value = "" And oncore_value <> "" Then

                Call tool2_cases.PrevNothingCurrX(ib_gridRng.Cells(ib_currRow, ib_currColumn), fillColor_updatedToOncore, ib_value, oncore_value)

            'case4 - ib = "inv" and oncore = number
            ElseIf isPrevInvoiceCurrOne(ib_value, oncore_value) Then

                Call tool2_cases.PrevInvoiceCurrOne(ib_gridRng.Cells(ib_currRow, ib_currColumn), fillColor_sameValue, ib_value, oncore_value)

            'case5 - ib <> oncore
            Else
            
                If tool2_cases.PrevXCurrY(ib_gridRng.Cells(ib_currRow, ib_currColumn), _
                                            oncore_gridRng.Cells(oncore_currRow, oncore_currColumn), _
                                            visitName, _
                                            procedureName, _
                                            fillColor_updatedToOncore, _
                                            fillColor_differentFromOncore, _
                                            ib_value, _
                                            oncore_value) = 1 Then
                    Exit Sub
                End If
            
            End If
nextVisit:
        Next ib_currColumn
    
nextProcedure:
    Next ib_currRow
        
Call Done


'    Application.ScreenUpdating = False
'    Application.Calculation = xlCalculationManual
'
'
'    Application.Calculation = xlCalculationAutomatic
'    Application.ScreenUpdating = True

End Sub

Function IsComment(rng As Range)
'function that returns true/false telling the user if there is comment in a cell
    
    Dim cmt As CommentThreaded
    
    Set cmt = rng.CommentThreaded
    
    If cmt Is Nothing Then
        IsComment = False
    Else
        IsComment = True
    End If

End Function

Sub OpenForm()
'opens user form at a specified location on the screen

    With form_amds
        'manual positioning
        .StartUpPosition = 0
        
        'set left and top location
        'currently set to bottom right of excel app
        .Left = Application.Left + (1 * Application.Width) - (1 * .Width)
        .Top = Application.Top + (1 * Application.Height) - (1 * .Height)
        
        'open in modless mode
        .Show vbModeless
    End With

End Sub

Private Sub Done()
'shows message that the program is done

MsgBox ("Tool2 finished updating the grid.")

End Sub

Private Sub UpdateIntBdgtGridToOncore()
'main subroutine


End Sub

Private Function SelectRanges(toolSheet As Worksheet) As Boolean
    
    'Int Bdgt Related - select two ranges and set data range
    'if data range can't be set, execution stops
    If Not SelectIntBdgtRanges(toolSheet) Then
        SelectRanges = False
        Exit Function
    End If
    
    'OnCore Related - select two ranges and set data range
    'if data range can't be set, execution stops
    If Not SelectTwoRangesAndSetDataRange(column_onCore, "OnCore", toolSheet) Then
        SelectRanges = False
        Exit Function
    End If
    
    SelectRanges = True

End Function

Private Function SelectIntBdgtRanges(toolSheet As Worksheet) As Boolean
    
    'Int Bdgt Related - select two ranges and set data range
    'if data range can't be set, execution stops
    If Not SelectTwoRangesAndSetDataRange(column_ib, "Internal Budget", toolSheet) Then
        SelectIntBdgtRanges = False
        Exit Function
    End If
        
    SelectIntBdgtRanges = True

End Function

Private Function AssembleSixRanges(ByRef ib_visits_rng As Range, _
                                    ByRef ib_procedures_rng As Range, _
                                    ByRef ib_grid_rng As Range, _
                                    ByRef oncore_visits_rng As Range, _
                                    ByRef oncore_procedures_rng As Range, _
                                    ByRef oncore_grid_rng As Range, _
                                    toolSheet As Worksheet)

    Set ib_visits_rng = AssembleRangeComponentsToRange(column_ib, row_workbookName, row_sheetName, row_visitNamesRange, toolSheet)
    Set ib_procedures_rng = AssembleRangeComponentsToRange(column_ib, row_workbookName, row_sheetName, row_proceduresRange, toolSheet)
    Set ib_grid_rng = AssembleRangeComponentsToRange(column_ib, row_workbookName, row_sheetName, row_dataRange, toolSheet)
    
    
    Set oncore_visits_rng = AssembleRangeComponentsToRange(column_onCore, row_workbookName, row_sheetName, row_visitNamesRange, toolSheet)
    Set oncore_procedures_rng = AssembleRangeComponentsToRange(column_onCore, row_workbookName, row_sheetName, row_proceduresRange, toolSheet)
    Set oncore_grid_rng = AssembleRangeComponentsToRange(column_onCore, row_workbookName, row_sheetName, row_dataRange, toolSheet)

End Function

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
 
Private Sub ConvertNoteToComment(cell As Range)
    
    Dim commentText As String
    
    If Not cell.comment Is Nothing Then
        commentText = cell.comment.Text
        cell.comment.Delete
        cell.AddCommentThreaded ("[previous note]" & commentText)
    End If
    
End Sub

Function AssembleComment(middleStr As String, prev As Variant, curr As Variant) As String
    Dim startStr As String
    Dim endStr As String
    
    If prev = "" Then prev = "[empty]"
    If curr = "" Then curr = "[empty]"
    
    startStr = "[" & Format(Date, "ddmmmyy") & " tool2 execution] "

    endStr = Chr(10) & _
            " -prev int bdgt value: " & CStr(prev) & Chr(10) & _
            " -curr onCore value: " & CStr(curr)
    AssembleComment = startStr & middleStr & endStr

End Function

Sub AddComment(cell As Range, commentText As String)
    'if there is a note, convert it to comment
    Call ConvertNoteToComment(cell)
    
    'add comment
    With cell
        If .CommentThreaded Is Nothing Then
            .AddCommentThreaded (commentText)
        Else
            .CommentThreaded.AddReply (commentText)
        End If
    End With
End Sub

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


