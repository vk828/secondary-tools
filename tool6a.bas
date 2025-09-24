Attribute VB_Name = "tool6a"
'Author/Developer: Vadim Krifuks
'Collaborators: Hui Zeng, Man Ming Tse
'Last Updated: 19January2025

Option Explicit

Sub tool6a_AddVertLookupFormulas()
'this subroutine adds vertical lookup formulas for Values
'user selects the appropriate ranges, and the subroutine adds lookup formulas

    Dim sourceRanges() As Range
    Dim destinationRanges() As Range
    Dim title1, prompt1, title2, prompt2 As String
    Dim sourceProceduresRng, sourceValuesRng, destinationProceduresRng, destinationValuesRng As Range
    Dim outputMsg As String
    
    'STEP 1 - select and adjust SOURCE RANGES
    
    title1 = "Select Source Procedures Range"
    prompt1 = "Select SOURCE PROCEDURES range (can be one cell) on the sheet you'd like to " & _
              "look up Values FROM." _
              & Chr(10) & Chr(10) & _
              "Note: Procedures and Values ranges adjust to first selected columns of their " & _
              "respective selections and same start/end rows."


    title2 = "Select Source Values Range"
    prompt2 = "Select SOURCE VALUES range (can be one cell) on the sheet you'd like to look " & _
              "up Values FROM." _
              & Chr(10) & Chr(10) & _
              "Note: Procedures and Values ranges adjust to first selected columns of their " & _
              "respective selections and same start/end rows."
    
    sourceRanges = SelectProceduresAndValuesRanges(title1, prompt1, title2, prompt2)
    'if user cancelled selection, stop execution
    If (sourceRanges(0) Is Nothing) Or (sourceRanges(1) Is Nothing) Then
        Exit Sub
    End If
    
    sourceRanges = AdjustProceduresAndValuesRanges(sourceRanges)
    
    Set sourceProceduresRng = sourceRanges(0)
    Set sourceValuesRng = sourceRanges(1)
    
    'STEP 2 - select and adjust DESTINATION RANGES
    
    title1 = "Select Loookup Procedures Range"
    prompt1 = "Select LOOKUP PROCEDURES range (can be one cell) on the sheet where you'd like to " & _
              "add lookup formulas TO." _
              & Chr(10) & Chr(10) & _
              "Note: Procedures and Formulas ranges adjust to first selected columns of their " & _
              "respective selections and same start/end rows."

    title2 = "Select Lookup Formulas Range"
    prompt2 = "Select a range of empty cells (can be one cell) to add " & _
              "lookup formulas TO." _
              & Chr(10) & Chr(10) & _
              "Note: Procedures and Formulas ranges adjust to first selected columns of their " & _
              "respective selections and same start/end rows."
              
    destinationRanges = SelectProceduresAndValuesRanges(title1, prompt1, title2, prompt2)
    'if user cancelled selection, stop execution
    If (destinationRanges(0) Is Nothing) Or (destinationRanges(1) Is Nothing) Then
        Exit Sub
    End If
    
    destinationRanges = AdjustProceduresAndValuesRanges(destinationRanges)
    
    Set destinationProceduresRng = destinationRanges(0)
    Set destinationValuesRng = destinationRanges(1)
    
    'STEP 3 - add LOOKUP FORMULAS
    Call GenerateAndFillVerticalLookupFormulas(sourceProceduresRng, sourceValuesRng, _
                                       destinationProceduresRng, destinationValuesRng)
    
    'STEP 4 - generate and show the message to user
    outputMsg = "You've added lookup formulas TO:" & Chr(10) & _
                "  -workbook:  " & destinationProceduresRng.Parent.Parent.Name & Chr(10) & _
                "  -sheet:  " & destinationProceduresRng.Parent.Name & Chr(10) & _
                "  -lookup procedures range:  " & destinationProceduresRng.Address(False, False) & Chr(10) & _
                "  -lookup formulas range:  " & destinationValuesRng.Address(False, False) & Chr(10) & _
                Chr(10) & _
                "FROM:" & Chr(10) & _
                "  -source workbook:  " & sourceProceduresRng.Parent.Parent.Name & Chr(10) & _
                "  -source sheet:  " & sourceProceduresRng.Parent.Name & Chr(10) & _
                "  -source procedures range:  " & sourceProceduresRng.Address(False, False) & Chr(10) & _
                "  -source values range:  " & sourceValuesRng.Address(False, False) & "."

    MsgBox outputMsg, vbInformation


End Sub

Private Function SelectProceduresAndValuesRanges(ByVal titleFirstSelection As String, _
                                              ByVal promptFirstSelection As String, _
                                              ByVal titleSecondSelection As String, _
                                              ByVal promptSecondSelection As String) As Range()
'this funciton: lets the user select Procedures and Values ranges, and
'returns an array that holds two selected ranges

    Dim returnArray(0 To 1) As Range
    Dim check As Boolean

    check = False
    
    'loop until either user cancels selection or until ranges are selected correctly
    Do
        
        'select first range
        Set returnArray(0) = Utilities.SelectRange(titleFirstSelection, promptFirstSelection)
        
        'if user cancelled selection; exit function
        If returnArray(0) Is Nothing Then
            Set returnArray(1) = Nothing
            GoTo Done:
        End If
        
        'select second range
        Set returnArray(1) = Utilities.SelectRange(titleSecondSelection, promptSecondSelection)
        
        'if user cancelled selection; exit function
        If returnArray(1) Is Nothing Then
            GoTo Done:
        End If
        
        'stop looping if workbooks or sheets don't match
        If returnArray(0).Parent.Parent.Name <> returnArray(1).Parent.Parent.Name Then
            MsgBox ("Procedures and Values ranges must be part of the same workbook and " _
                    & "sheet. You selected two different workbooks. Please try again.")
        ElseIf returnArray(0).Parent.Name <> returnArray(1).Parent.Name Then
            MsgBox ("Procedures and Values ranges must be part of the same workbook and " _
                    & "sheet. You selected two different sheets. Please try again.")
        Else
            check = True
        End If
    Loop Until check = True
    
Done:
    SelectProceduresAndValuesRanges = returnArray

End Function

Private Function AdjustProceduresAndValuesRanges(rawRange() As Range) As Range()
'this function adjusts ranges to their first respective columns and same start/end (min/max) rows
'input: an array with two raw ranges; output: an array with two adjusted ranges
    
    Dim rowTop, rowBottom, rowTopTemp, rowBottomTemp, columnProcedures, columnValues As Long
    Dim outputRange(0 To 1) As Range

    With rawRange(0)
        rowTop = .Cells(1, 1).row
        rowBottom = .Cells(.rows.count, 1).row
        columnProcedures = .Cells(1, 1).column
    End With
    
    With rawRange(1)
        rowTopTemp = .Cells(1, 1).row
        rowBottomTemp = .Cells(.rows.count, 1).row
        columnValues = .Cells(1, 1).column
    End With

    If rowTop > rowTopTemp Then
        rowTop = rowTopTemp
    End If
    
    If rowBottom < rowBottomTemp Then
        rowBottom = rowBottomTemp
    End If
    
    With rawRange(0).Parent
        Set outputRange(0) = .Range(Cells(rowTop, columnProcedures), Cells(rowBottom, columnProcedures))
        Set outputRange(1) = .Range(Cells(rowTop, columnValues), Cells(rowBottom, columnValues))
    End With
    
    AdjustProceduresAndValuesRanges = outputRange
End Function

Sub GenerateAndFillVerticalLookupFormulas(ByVal sourceProceduresRng As Range, _
                                                  ByVal sourceLookupRng As Range, _
                                                  ByVal destinationProceduresRng As Range, _
                                                  ByVal destinationLookupRng As Range)
'this subroutine generates the the lookup formula for the top cell and then uses it
'to generate the entire range
'if lookup value is "", value is ""

'=IF($I16="",
'     "",
'    IF(ISNA(MATCH(LEFT($I16,255),LEFT(Budget_Details_ADJ_DBL!$I$16:$I$199,255),0)),
'        "NO RESULT FOR " & $I16,
'        IF(INDEX(Budget_Details_ADJ_DBL!$G$16:$G$199, MATCH(LEFT($I16,255), LEFT( Budget_Details_ADJ_DBL!$I$16:$I$199,255), 0))="",
'             "",
'            (INDEX(Budget_Details_ADJ_DBL!$G$16:$G$199, MATCH(LEFT($I16,255), LEFT( Budget_Details_ADJ_DBL!$I$16:$I$199,255), 0)))
'        )
'    )
')

    Dim sourceProceduresAddress, sourceValuesAddress, destinationTopProcedureAddress As String
    Dim formula As String
    Dim notFoundMsg As String
    
    notFoundMsg = """NO RESULT for "" & "

    sourceProceduresAddress = sourceProceduresRng.Address(external:=True)
    sourceValuesAddress = sourceLookupRng.Address(external:=True)
    destinationTopProcedureAddress = destinationProceduresRng.Cells(1, 1).Address(RowAbsolute:=False)
    
    formula = "=IF(" & destinationTopProcedureAddress & "=""""," & Chr(10) & _
                  String(4, Chr(32)) & """""," & Chr(10) & _
                  String(4, Chr(32)) & "IF(ISNA(MATCH(LEFT(" & destinationTopProcedureAddress & ",255),LEFT(" & sourceProceduresAddress & ",255),0))," & Chr(10) & _
                      String(8, Chr(32)) & notFoundMsg & destinationTopProcedureAddress & "," & Chr(10) & _
                      String(8, Chr(32)) & "IF(INDEX(" & sourceValuesAddress & ", MATCH(LEFT(" & destinationTopProcedureAddress & ",255), LEFT( " & sourceProceduresAddress & ",255), 0))=""""," & Chr(10) & _
                          String(12, Chr(32)) & """""," & Chr(10) & _
                          String(12, Chr(32)) & "(INDEX(" & sourceValuesAddress & ", MATCH(LEFT(" & destinationTopProcedureAddress & ",255), LEFT( " & sourceProceduresAddress & ",255), 0)))" & Chr(10) & _
                      String(8, Chr(32)) & ")" & Chr(10) & _
                  String(4, Chr(32)) & ")" & Chr(10) & _
              ")"
    
    'add formula to the top cell
    destinationLookupRng.Cells(1, 1).Formula2 = formula
    
    'add formula to the entire range
    destinationLookupRng.Cells(1, 1).Copy
    destinationLookupRng.PasteSpecial Paste:=xlPasteFormulas
    
End Sub


