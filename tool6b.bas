Attribute VB_Name = "tool6b"
'Author/Developer: Vadim Krifuks
'Collaborators: Hui Zeng, Man Ming Tse
'Last Updated: 29November2025

Option Explicit

Sub tool6b_AddHorizLookupFormulas()
'this subroutine adds horizontal lookup formulas for Values
'user selects the appropriate ranges, and the subroutine adds lookup formulas

    Dim sourceRanges() As Range
    Dim destinationRanges() As Range
    Dim title1, prompt1, title2, prompt2, note As String
    Dim sourceVisitsRng, sourceValuesRng, destinationVisitsRng, destinationValuesRng As Range
    Dim outputMsg As String
    
    Call Instructions
    
    'STEP 1 - select and adjust SOURCE RANGES
    
    note = "Note: Visits and Values ranges auto-adjust to the same start/end columns."
    
    title1 = "Select Source Visits Range"
    prompt1 = "Select SOURCE VISITS range (can be one cell) on the sheet you'd like to " & _
              "look up Values FROM." _
              & Chr(10) & Chr(10) & _
              note


    title2 = "Select Source Values Range"
    prompt2 = "Select SOURCE VALUES range (can be one cell) on the sheet you'd like to look " & _
              "up Values FROM." _
              & Chr(10) & Chr(10) & _
              note
    
    sourceRanges = SelectVisitsAndValuesRanges(title1, prompt1, title2, prompt2)
    'if user cancelled selection, stop execution
    If (sourceRanges(0) Is Nothing) Or (sourceRanges(1) Is Nothing) Then
        Exit Sub
    End If
    
    sourceRanges = LineUpVisitsAndValuesColumns(sourceRanges)
    
    Set sourceVisitsRng = sourceRanges(0)
    Set sourceValuesRng = sourceRanges(1)
    
    'STEP 2 - select and adjust DESTINATION RANGES
    
    note = "Note: Visits and Values ranges auto-adjust to the same start/end columns. " & _
            "Lookup Visits range is resized to match rows of Source Visits range."

    title1 = "Select Lookup Visits Range"
    prompt1 = "Select LOOKUP VISITS range (can be one cell) on the sheet where you'd like to " & _
              "add lookup formulas TO." _
              & Chr(10) & Chr(10) & _
              note

    title2 = "Select Lookup Formulas Range"
    prompt2 = "Select a range of empty cells (can be one cell) to add " & _
              "lookup formulas TO." _
              & Chr(10) & Chr(10) & _
              note
              
    destinationRanges = SelectVisitsAndValuesRanges(title1, prompt1, title2, prompt2)
    'if user cancelled selection, stop execution
    If (destinationRanges(0) Is Nothing) Or (destinationRanges(1) Is Nothing) Then
        Exit Sub
    End If
    
    destinationRanges = LineUpVisitsAndValuesColumns(destinationRanges)
    
    Set destinationVisitsRng = ResizeDestinationVisitsRows(sourceRanges(0), destinationRanges(0))
    Set destinationValuesRng = destinationRanges(1)
    
    'STEP 3 - add LOOKUP FORMULAS
    Call GenerateAndFillHorizontalLookupFormulas(sourceVisitsRng, sourceValuesRng, _
                                       destinationVisitsRng, destinationValuesRng)
    
    'STEP 4 - generate and show the message to user
    outputMsg = "You've added lookup formulas TO:" & Chr(10) & _
                "  -workbook:  " & destinationVisitsRng.Parent.Parent.name & Chr(10) & _
                "  -sheet:  " & destinationVisitsRng.Parent.name & Chr(10) & _
                "  -lookup visits range:  " & destinationVisitsRng.Address(False, False) & Chr(10) & _
                "  -lookup formulas range:  " & destinationValuesRng.Address(False, False) & Chr(10) & _
                Chr(10) & _
                "FROM:" & Chr(10) & _
                "  -source workbook:  " & sourceVisitsRng.Parent.Parent.name & Chr(10) & _
                "  -source sheet:  " & sourceVisitsRng.Parent.name & Chr(10) & _
                "  -source visits range:  " & sourceVisitsRng.Address(False, False) & Chr(10) & _
                "  -source values range:  " & sourceValuesRng.Address(False, False) & "."

    MsgBox outputMsg, vbInformation


End Sub

Private Sub Instructions()
' routine that provides information about the tool to its user

    Dim message As String
    
    message = "Tool adds formulas to a single cell or group of cells to lookup values from " & _
            "horizontally organized data based on a common label." & _
            Chr(10) & _
            Chr(10) & _
            "User selects four ranges:" & Chr(10) & _
            " - lookup_array (source visits)" & Chr(10) & _
            " - return_array (source values)" & Chr(10) & _
            " - lookup_value (lookup visits)" & Chr(10) & _
            " - lookup_formula (lookup formulas)" & _
            Chr(10) & _
            Chr(10) & _
            "The lookup_array and return_array ranges will automatically extend so they " & _
            "share the same left and right column boundaries. Similarly, the lookup_value " & _
            "and lookup_formula ranges will also automatically extend to share the same " & _
            "left and right column boundaries. Additionally, " & _
            "lookup_value will adjust to include the same number of rows as " & _
            "lookup_array (single or multiple rows)."
    
    MsgBox message, vbInformation, "Tool Info"

End Sub


Private Function SelectVisitsAndValuesRanges(ByVal titleFirstSelection As String, _
                                              ByVal promptFirstSelection As String, _
                                              ByVal titleSecondSelection As String, _
                                              ByVal promptSecondSelection As String) As Range()
'this function: lets the user select Visits and Values ranges, and
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
        If returnArray(0).Parent.Parent.name <> returnArray(1).Parent.Parent.name Then
            MsgBox ("Visits and Values ranges must be part of the same workbook and " _
                    & "sheet. You selected two different WORKBOOKS. Please try again.")
        ElseIf returnArray(0).Parent.name <> returnArray(1).Parent.name Then
            MsgBox ("Visits and Values ranges must be part of the same workbook and " _
                    & "sheet. You selected two different SHEETS. Please try again.")
        Else
            check = True
        End If
    Loop Until check = True
    
Done:
    SelectVisitsAndValuesRanges = returnArray

End Function

Private Function LineUpVisitsAndValuesColumns(rawRange() As Range) As Range()
' Repositions both ranges to share identical start/end columns while preserving row extents
    
    Dim columnLeft, columnRight, columnLeftTemp, columnRightTemp As Long
    Dim outputRange(0 To 1) As Range

    ' Find overall leftmost and rightmost columns across both ranges
    With rawRange(0)
        columnLeft = .Cells(1, 1).column
        columnRight = .Cells(1, .Columns.count).column
    End With
    
    With rawRange(1)
        columnLeftTemp = .Cells(1, 1).column
        columnRightTemp = .Cells(1, .Columns.count).column
    End With
    
    ' Use the absolute leftmost start and rightmost end
    If columnLeftTemp < columnLeft Then columnLeft = columnLeftTemp
    If columnRightTemp > columnRight Then columnRight = columnRightTemp
    
    ' Reposition both ranges to start at columnLeft, with full width
    With rawRange(0).Parent
        Set outputRange(0) = .Range(.Cells(rawRange(0).Row, columnLeft), _
                                   .Cells(rawRange(0).Row + rawRange(0).Rows.count - 1, columnRight))
        Set outputRange(1) = .Range(.Cells(rawRange(1).Row, columnLeft), _
                                   .Cells(rawRange(1).Row + rawRange(1).Rows.count - 1, columnRight))
    End With
    
    LineUpVisitsAndValuesColumns = outputRange
End Function

Private Function ResizeDestinationVisitsRows(sourceRng As Range, destinationRng As Range) As Range
' Resizes destination range to be the same number of rows as source range
    Set ResizeDestinationVisitsRows = destinationRng.Resize(sourceRng.Rows.count)
End Function

Private Sub GenerateAndFillHorizontalLookupFormulas(ByVal sourceVisitsRng As Range, _
                                                  ByVal sourceValuesRng As Range, _
                                                  ByVal destinationVisitsRng As Range, _
                                                  ByVal destinationValuesRng As Range)
'this subroutine generates the the lookup formula for the leftmost cell and then uses it
'to generate the entire range

    Dim sourceVisitsAddress, sourceValuesAddress As String
    Dim formula As String
    Dim notFoundMsg As String
    Dim destinationCurrVisitAddress As String           'used for single row headers and simple formula
    Dim destinationVisitsCurrColumnAddress As String    'used for multi row headers and complex formula
    Dim headerRowsCount As Integer                      'number of rows in the selected headers
    
    notFoundMsg = """NO RESULT for "" & "

    sourceVisitsAddress = sourceVisitsRng.Address(external:=True)
    sourceValuesAddress = sourceValuesRng.Address(external:=True)
        
    headerRowsCount = sourceVisitsRng.Rows.count
    
    If headerRowsCount = 1 Then
        destinationCurrVisitAddress = destinationVisitsRng.Cells(1, 1).Address(columnabsolute:=False)
    
        '=LET(
        '    currentVisit, LEFT(TRIM(CLEAN(BA$31)), 255),
        '    sourceVisits, LEFT(TRIM(CLEAN('[CABA-201-002 JIIM_Myositis Budget V4.1_PI Kim_UCSF_FINAL_25Nov2025.xlsx]Post Infusion Follow Up JIIM'!$D$17:$AH$17)), 255),
        '    sourceValues, '[CABA-201-002 JIIM_Myositis Budget V4.1_PI Kim_UCSF_FINAL_25Nov2025.xlsx]Post Infusion Follow Up JIIM'!$D$71:$AH$71,
        '    IF(currentVisit = "",
        '        0,
        '        LET(
        '            firstMatchColumnNumber, MATCH(currentVisit, sourceVisits, 0),
        '            firstValue, INDEX(sourceValues, firstMatchColumnNumber),
        '            IF(ISNA(firstMatchColumnNumber),
        '                "NO RESULT for " & currentVisit,
        '                firstValue
        '            )
        '        )
        '    )
        ')
        
        formula = Space(0) & "=LET(" & Chr(10) & _
                Space(4) & "currentVisit, LEFT(TRIM(CLEAN(" & destinationCurrVisitAddress & ")), 255)," & Chr(10) & _
                Space(4) & "sourceVisits, LEFT(TRIM(CLEAN(" & sourceVisitsAddress & ")), 255)," & Chr(10) & _
                Space(4) & "sourceValues, " & sourceValuesAddress & "," & Chr(10) & _
                Space(4) & "IF(currentVisit = """"," & Chr(10) & _
                    Space(8) & "0," & Chr(10) & _
                    Space(8) & "LET(" & Chr(10) & _
                        Space(12) & "firstMatchColumnNumber, MATCH(currentVisit, sourceVisits, 0)," & Chr(10) & _
                        Space(12) & "firstValue, INDEX(sourceValues, firstMatchColumnNumber)," & Chr(10) & _
                        Space(12) & "IF(ISNA(firstMatchColumnNumber)," & Chr(10) & _
                            Space(16) & notFoundMsg & "currentVisit," & Chr(10) & _
                            Space(16) & "firstValue" & Chr(10) & _
                        Space(12) & ")" & Chr(10) & _
                    Space(8) & ")" & Chr(10) & _
                Space(4) & ")" & Chr(10) & _
            Space(0) & ")"


    Else
        destinationVisitsCurrColumnAddress = destinationVisitsRng.Columns(1).Address(columnabsolute:=False)

        '=LET(
        '    currentVisitKeys, TRIM(CLEAN(AK$31:AK$33)),
        '    sourceVisitKeysBlock, TRIM(CLEAN('[CABA-201-002 JIIM_Myositis Budget V4.1_PI Kim_UCSF_FINAL_25Nov2025.xlsx]Screening_Treatment JIIM'!$D$16:$N$18)),
        '    sourceValues, '[CABA-201-002 JIIM_Myositis Budget V4.1_PI Kim_UCSF_FINAL_25Nov2025.xlsx]Screening_Treatment JIIM'!$D$103:$N$103,
        '    currentVisitFullName, CONCAT(currentVisitKeys),
        '    IF(currentVisitFullName = "",
        '        0,
        '        LET(
        '            matchMatrix, (sourceVisitKeysBlock = currentVisitKeys) * 1,
        '            matchMatrixTotals, MMULT(TRANSPOSE(matchMatrix), SEQUENCE(ROWS(matchMatrix),1,1,0)),
        '            requiredMatchTotal, ROWS(sourceVisitKeysBlock),
        '            firstMatchColumnNumber, MATCH(requiredMatchTotal, matchMatrixTotals, 0),
        '            firstValue, INDEX(sourceValues, firstMatchColumnNumber),
        '            IF(ISNA(firstMatchColumnNumber),
        '                "NO RESULT for " & currentVisitFullName,
        '                firstValue
        '            )
        '        )
        '    )
        ')

        formula = Space(0) & "=LET(" & Chr(10) & _
                    Space(4) & "currentVisitKeys, TRIM(CLEAN(" & destinationVisitsCurrColumnAddress & "))," & Chr(10) & _
                    Space(4) & "sourceVisitKeysBlock, TRIM(CLEAN(" & sourceVisitsAddress & "))," & Chr(10) & _
                    Space(4) & "sourceValues, " & sourceValuesAddress & "," & Chr(10) & _
                    Space(4) & "currentVisitFullName, CONCAT(currentVisitKeys)," & Chr(10) & _
                    Space(4) & "IF(currentVisitFullName = """"," & Chr(10) & _
                        Space(8) & "0," & Chr(10) & _
                        Space(8) & "LET(" & Chr(10) & _
                            Space(12) & "matchMatrix, (sourceVisitKeysBlock = currentVisitKeys) * 1," & Chr(10) & _
                            Space(12) & "matchMatrixTotals, MMULT(TRANSPOSE(matchMatrix), SEQUENCE(ROWS(matchMatrix),1,1,0))," & Chr(10) & _
                            Space(12) & "requiredMatchTotal, ROWS(sourceVisitKeysBlock)," & Chr(10) & _
                            Space(12) & "firstMatchColumnNumber, MATCH(requiredMatchTotal, matchMatrixTotals, 0)," & Chr(10) & _
                            Space(12) & "firstValue, INDEX(sourceValues, firstMatchColumnNumber)," & Chr(10) & _
                            Space(12) & "IF(ISNA(firstMatchColumnNumber)," & Chr(10) & _
                                Space(16) & notFoundMsg & "currentVisitFullName," & Chr(10) & _
                                Space(16) & "firstValue" & Chr(10) & _
                            Space(12) & ")" & Chr(10) & _
                        Space(8) & ")" & Chr(10) & _
                    Space(4) & ")" & Chr(10) & _
                Space(0) & ")"
    End If

    'add formula to the leftmost cell
    destinationValuesRng.Cells(1, 1).Formula2 = formula
    
    'add formula to the entire range
    destinationValuesRng.Formula2 = destinationValuesRng.Cells(1, 1).Formula2

End Sub




