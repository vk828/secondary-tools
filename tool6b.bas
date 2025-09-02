Attribute VB_Name = "tool6b"
'Author/Developer: Vadim Krifuks
'Collaborators: Hui Zeng, Man Ming Tse
'Last Updated: 20January2025

Option Explicit

Sub tool6b_AddHorizLookupFormulas()
'this subroutine adds horizontal lookup formulas for Values
'user selects the appropriate ranges, and the subroutine adds lookup formulas

    Dim sourceRanges() As Range
    Dim destinationRanges() As Range
    Dim title1, prompt1, title2, prompt2, note As String
    Dim sourceVisitsRng, sourceValuesRng, destinationVisitsRng, destinationValuesRng As Range
    Dim outputMsg As String
    
    'STEP 1 - select and adjust SOURCE RANGES
    
    note = "Note: Visits and Values ranges adjust to top selected rows of their " & _
           "respective selections and same start/end columns."
    
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
    
    sourceRanges = AdjustVisitsAndValuesRanges(sourceRanges)
    
    Set sourceVisitsRng = sourceRanges(0)
    Set sourceValuesRng = sourceRanges(1)
    
    'STEP 2 - select and adjust DESTINATION RANGES
    
    title1 = "Select Loookup Visits Range"
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
    
    destinationRanges = AdjustVisitsAndValuesRanges(destinationRanges)
    
    Set destinationVisitsRng = destinationRanges(0)
    Set destinationValuesRng = destinationRanges(1)
    
    'STEP 3 - add LOOKUP FORMULAS
    Call GenerateAndFillHorizontalLookupFormulas(sourceVisitsRng, sourceValuesRng, _
                                       destinationVisitsRng, destinationValuesRng)
    
    'STEP 4 - generate and show the message to user
    outputMsg = "You've added lookup formulas TO:" & Chr(10) & _
                "  -workbook:  " & destinationVisitsRng.Parent.Parent.Name & Chr(10) & _
                "  -sheet:  " & destinationVisitsRng.Parent.Name & Chr(10) & _
                "  -lookup visits range:  " & destinationVisitsRng.Address(False, False) & Chr(10) & _
                "  -lookup formulas range:  " & destinationValuesRng.Address(False, False) & Chr(10) & _
                Chr(10) & _
                "FROM:" & Chr(10) & _
                "  -source workbook:  " & sourceVisitsRng.Parent.Parent.Name & Chr(10) & _
                "  -source sheet:  " & sourceVisitsRng.Parent.Name & Chr(10) & _
                "  -source visits range:  " & sourceVisitsRng.Address(False, False) & Chr(10) & _
                "  -source values range:  " & sourceValuesRng.Address(False, False) & "."

    MsgBox outputMsg, vbInformation


End Sub

Private Function SelectVisitsAndValuesRanges(ByVal titleFirstSelection As String, _
                                              ByVal promptFirstSelection As String, _
                                              ByVal titleSecondSelection As String, _
                                              ByVal promptSecondSelection As String) As Range()
'this funciton: lets the user select Visits and Values ranges, and
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
            MsgBox ("Visits and Values ranges must be part of the same workbook and " _
                    & "sheet. You selected two different workbooks. Please try again.")
        ElseIf returnArray(0).Parent.Name <> returnArray(1).Parent.Name Then
            MsgBox ("Visits and Values ranges must be part of the same workbook and " _
                    & "sheet. You selected two different sheets. Please try again.")
        Else
            check = True
        End If
    Loop Until check = True
    
Done:
    SelectVisitsAndValuesRanges = returnArray

End Function

Private Function AdjustVisitsAndValuesRanges(rawRange() As Range) As Range()
'this function adjusts ranges to their top respective rows and same start/end (left/right) columns
'input: an array with two raw ranges; output: an array with two adjusted ranges
    
    Dim rowVisits, rowValues As Long
    Dim columnLeft, columnRight, columnLeftTemp, columnRightTemp As Long
    Dim outputRange(0 To 1) As Range

    With rawRange(0)
        rowVisits = .Cells(1, 1).row
        columnLeft = .Cells(1, 1).column
        columnRight = .Cells(1, .Columns.count).column
    End With
    
    With rawRange(1)
        rowValues = .Cells(1, 1).row
        columnLeftTemp = .Cells(1, 1).column
        columnRightTemp = .Cells(1, .Columns.count).column
    End With

    If columnLeft > columnLeftTemp Then
        columnLeft = columnLeftTemp
    End If
    
    If columnRight < columnRightTemp Then
        columnRight = columnRightTemp
    End If
    
    With rawRange(0).Parent
        Set outputRange(0) = .Range(Cells(rowVisits, columnLeft), Cells(rowVisits, columnRight))
        Set outputRange(1) = .Range(Cells(rowValues, columnLeft), Cells(rowValues, columnRight))
    End With
    
    AdjustVisitsAndValuesRanges = outputRange
End Function

Private Sub GenerateAndFillHorizontalLookupFormulas(ByVal sourceVisitsRng As Range, _
                                                  ByVal sourceLookupRng As Range, _
                                                  ByVal destinationProceduresRng As Range, _
                                                  ByVal destinationLookupRng As Range)
'this subroutine generates the the lookup formula for the most left cell and then uses it
'to generate the entire range

'=IF(M$15="",
'     0,
'    IF(ISNA(MATCH(LEFT(M$15,255),LEFT(Budget_Details_ADJ_DBL!$M$15:$BA$15,255),0)),
'        "NO RESULT FOR " & M$15,
'        (INDEX(Budget_Details_ADJ_DBL!$M$175:$BA$175,,MATCH(LEFT(M$15,255),LEFT(Budget_Details_ADJ_DBL!$M$15:$BA$15,255), 0)))
'    )
')


    Dim sourceVisitsAddress, sourceValuesAddress, destinationLeftVisitAddress As String
    Dim formula As String
    Dim notFoundMsg As String
    
    notFoundMsg = """NO RESULT for "" & "

    sourceVisitsAddress = sourceVisitsRng.Address(external:=True)
    sourceValuesAddress = sourceLookupRng.Address(external:=True)
    destinationLeftVisitAddress = destinationProceduresRng.Cells(1, 1).Address(ColumnAbsolute:=False)
    
    formula = "=IF(" & destinationLeftVisitAddress & "=""""," & Chr(10) & _
                  String(4, Chr(32)) & 0 & "," & Chr(10) & _
                  String(4, Chr(32)) & "IF(ISNA(MATCH(LEFT(" & destinationLeftVisitAddress & ",255),LEFT(" & sourceVisitsAddress & ",255),0))," & Chr(10) & _
                      String(8, Chr(32)) & notFoundMsg & destinationLeftVisitAddress & "," & Chr(10) & _
                      String(8, Chr(32)) & "(INDEX(" & sourceValuesAddress & ",, MATCH(LEFT(" & destinationLeftVisitAddress & ",255), LEFT( " & sourceVisitsAddress & ",255), 0)))" & Chr(10) & _
                  String(4, Chr(32)) & ")" & Chr(10) & _
              ")"
    
    'add formula to the left cell
    destinationLookupRng.Cells(1, 1).Formula2 = formula
    
    'add formula to the entire range
    destinationLookupRng.Cells(1, 1).Copy
    destinationLookupRng.PasteSpecial Paste:=xlPasteFormulas

End Sub
