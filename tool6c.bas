Attribute VB_Name = "tool6c"
'Author/Developer: Vadim Krifuks
'Collaborators: Hui Zeng, Man Ming Tse
'Last Updated: 21January2025

Option Explicit

Sub tool6c_AddTwoDimensionalLookupFormulas()
'this subroutine adds two-dimensional lookup formulas for Values
'user selects the appropriate ranges, and the subroutine adds lookup formulas

    Dim sourceRanges() As Range
    Dim destinationRanges() As Range
    Dim title1, prompt1, title2, prompt2, note As String
    Dim sourceProceduresRng, sourceVisitsRng, sourceValuesRng As Range
    Dim destinationProceduresRng, destinationVisitsRng, destinationValuesRng As Range
    Dim outputMsg As String
    
    'STEP 1 - select and adjust SOURCE PROCEDURES and VISITS RANGES
    
    note = "Note: Procedures and Visit ranges adjust to the first selected column and " & _
           "the top row, respectively. Values range is set by intersecting " & _
           "Procedures and Visits ranges."
    
    title1 = "Select Source Procedures Range"
    prompt1 = "Select SOURCE PROCEDURES range on the sheet you'd like to " & _
              "look up Values FROM." _
              & Chr(10) & Chr(10) & _
              note

    title2 = "Select Source Visits Range"
    prompt2 = "Select SOURCE VISITS range on the sheet you'd like to look " & _
              "up Values FROM." _
              & Chr(10) & Chr(10) & _
              note
              
    sourceRanges = SelectProceduresAndVisitsRanges(title1, prompt1, title2, prompt2)
    'if user cancelled selection, stop execution
    If (sourceRanges(0) Is Nothing) Or (sourceRanges(1) Is Nothing) Then
        Exit Sub
    End If
    
    sourceRanges = AdjustProceduresAndVisitsRanges(sourceRanges)
    
    Set sourceProceduresRng = sourceRanges(0)
    Set sourceVisitsRng = sourceRanges(1)
    
    'STEP 2 - set SOURCE VALUES RANGE
    
    Set sourceValuesRng = SetValuesRange(sourceRanges)
    
    'STEP 3 - select and adjust DESTINATION PROCEDURES and VISITS RANGES
    
    title1 = "Select Loookup Procedures Range"
    prompt1 = "Select LOOKUP PROCEDURES range on the sheet where you'd like to " & _
              "add lookup formulas TO." _
              & Chr(10) & Chr(10) & _
              note

    title2 = "Select Lookup Visits Range"
    prompt2 = "Select LOOKUP VISITS range on the sheet where you'd like to " & _
              "add lookup formulas TO." _
              & Chr(10) & Chr(10) & _
              note
              
    destinationRanges = SelectProceduresAndVisitsRanges(title1, prompt1, title2, prompt2)
    'if user cancelled selection, stop execution
    If (destinationRanges(0) Is Nothing) Or (destinationRanges(1) Is Nothing) Then
        Exit Sub
    End If
    
    destinationRanges = AdjustProceduresAndVisitsRanges(destinationRanges)
    
    Set destinationProceduresRng = destinationRanges(0)
    Set destinationVisitsRng = destinationRanges(1)
    
    'STEP 4 - set DESTINATION VALUES RANGE
    
    Set destinationValuesRng = SetValuesRange(destinationRanges)
    
    
    'STEP 5 - add LOOKUP FORMULAS
    
    Call GenerateAndFillTwoDimensionalLookupFormulas(sourceProceduresRng, _
                                                     sourceVisitsRng, _
                                                     sourceValuesRng, _
                                                     destinationProceduresRng, _
                                                     destinationVisitsRng, _
                                                     destinationValuesRng)
    
    'STEP 6 - generate and show the message to user
    
    outputMsg = "You've added lookup formulas TO:" & Chr(10) & _
                "  -workbook:  " & destinationProceduresRng.Parent.Parent.name & Chr(10) & _
                "  -sheet:  " & destinationProceduresRng.Parent.name & Chr(10) & _
                "  -lookup procedures range:  " & destinationProceduresRng.Address(False, False) & Chr(10) & _
                "  -lookup visits range:  " & destinationVisitsRng.Address(False, False) & Chr(10) & _
                "  -lookup formulas range:  " & destinationValuesRng.Address(False, False) & Chr(10) & _
                Chr(10) & _
                "FROM:" & Chr(10) & _
                "  -source workbook:  " & sourceProceduresRng.Parent.Parent.name & Chr(10) & _
                "  -source sheet:  " & sourceProceduresRng.Parent.name & Chr(10) & _
                "  -source procedures range:  " & sourceProceduresRng.Address(False, False) & Chr(10) & _
                "  -source visits range:  " & sourceVisitsRng.Address(False, False) & Chr(10) & _
                "  -source values range:  " & sourceValuesRng.Address(False, False) & "."

    MsgBox outputMsg, vbInformation


End Sub

Private Function SelectProceduresAndVisitsRanges(ByVal titleFirstSelection As String, _
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
        If returnArray(0).Parent.Parent.name <> returnArray(1).Parent.Parent.name Then
            MsgBox ("Procedures and Visits ranges must be part of the same workbook and " _
                    & "sheet. You selected two different workbooks. Please try again.")
        ElseIf returnArray(0).Parent.name <> returnArray(1).Parent.name Then
            MsgBox ("Procedures and Visits ranges must be part of the same workbook and " _
                    & "sheet. You selected two different sheets. Please try again.")
        Else
            check = True
        End If
    Loop Until check = True
    
Done:
    SelectProceduresAndVisitsRanges = returnArray

End Function

Private Function AdjustProceduresAndVisitsRanges(rawRange() As Range) As Range()
'this function adjusts first range to the first column and the second to the top row
'input: an array with two raw ranges; output: an array with two adjusted ranges
    
    Dim outputRange(0 To 1) As Range
    
    Set outputRange(0) = rawRange(0).Columns(1)
    Set outputRange(1) = rawRange(1).Rows(1)
    
    AdjustProceduresAndVisitsRanges = outputRange
End Function

Private Function SetValuesRange(rng() As Range) As Range
'this function sets a range by intersecting input ranges
'return the resulting range
    
    Dim outputRange As Range
    Dim topRow, bottomRow, leftColumn, rightColumn As Long
    
    With rng(0)
        topRow = .Rows(1).Row
        bottomRow = .Rows(.Rows.count).Row
    End With
    
    With rng(1)
        leftColumn = .Columns(1).column
        rightColumn = .Columns(.Columns.count).column
        
        Set outputRange = .Parent.Range(Cells(topRow, leftColumn), Cells(bottomRow, rightColumn))
    End With
    
    Set SetValuesRange = outputRange
End Function

Private Sub GenerateAndFillTwoDimensionalLookupFormulas(ByVal sourceProceduresRng As Range, _
                                                        ByVal sourceVisitsRng As Range, _
                                                        ByVal sourceValuesRng As Range, _
                                                        ByVal destinationProceduresRng As Range, _
                                                        ByVal destinationVisitsRng As Range, _
                                                        ByVal destinationValuesRng As Range)
'this subroutine generates the the lookup formula for the top left cell and then uses it
'to generate the entire range
'if lookup value is "", value is ""

'=IF(OR(B$1="",$A2=""),
'    "",
'    IF(AND(ISNA(MATCH(B$1,Sheet1!$C$4:$T$4,0)),ISNA(MATCH($A2,Sheet1!$B$5:$B$36,0))),
'        "NO RESULT",
'        IF(ISNA(MATCH(B$1,Sheet1!$C$4:$T$4,0)),
'            "NO RESULT for " & B$1,
'            IF(ISNA(MATCH($A2,Sheet1!$B$5:$B$36,0)),
'                "NO RESULT for " & $A2,
'                INDEX(Sheet1!$C$5:$T$36,MATCH($A2,Sheet1!$B$5:$B$36,0),MATCH(B$1,Sheet1!$C$4:$T$4,0))
'            )
'        )
'    )
')


    Dim sourceProceduresAddress, sourceVisitsAddress, sourceValuesAddress As String
    Dim lookupFirstProcedureAddress, lookupFirstVisitAddress As String
    Dim formula As String
    Dim procedureNorVisitFoundMsg, noProcedureFoundMsg, noVisitFoundMsg As String

    sourceProceduresAddress = sourceProceduresRng.Address(external:=True)
    sourceVisitsAddress = sourceVisitsRng.Address(external:=True)
    sourceValuesAddress = sourceValuesRng.Address(external:=True)
    
    lookupFirstProcedureAddress = destinationProceduresRng.Cells(1, 1).Address(RowAbsolute:=False)
    lookupFirstVisitAddress = destinationVisitsRng.Cells(1, 1).Address(columnabsolute:=False)
    
    procedureNorVisitFoundMsg = """NO RESULT"""
    noProcedureFoundMsg = """NO RESULT for ""&" & lookupFirstProcedureAddress
    noVisitFoundMsg = """NO RESULT for ""&" & lookupFirstVisitAddress
    
    formula = "=IF(OR(" & lookupFirstProcedureAddress & "=""""," & lookupFirstVisitAddress & "="""")," & Chr(10) & _
                    String(4, Chr(32)) & """""," & Chr(10) & _
                    String(4, Chr(32)) & "IF(AND(ISNA(MATCH(LEFT(" & lookupFirstProcedureAddress & ",255),LEFT(" & sourceProceduresAddress & ",255),0))," & _
                                                "ISNA(MATCH(LEFT(" & lookupFirstVisitAddress & ",255),LEFT(" & sourceVisitsAddress & ",255),0)))," & Chr(10) & _
                        String(8, Chr(32)) & procedureNorVisitFoundMsg & "," & Chr(10) & _
                        String(8, Chr(32)) & "IF(ISNA(MATCH(LEFT(" & lookupFirstProcedureAddress & ",255),LEFT(" & sourceProceduresAddress & ",255),0))," & Chr(10) & _
                            String(12, Chr(32)) & noProcedureFoundMsg & "," & Chr(10) & _
                            String(12, Chr(32)) & "IF(ISNA(MATCH(LEFT(" & lookupFirstVisitAddress & ",255),LEFT(" & sourceVisitsAddress & ",255),0))," & Chr(10) & _
                                String(16, Chr(32)) & noVisitFoundMsg & "," & Chr(10) & _
                                String(16, Chr(32)) & "IF(INDEX(" & _
                                                            sourceValuesAddress & ", " & _
                                                            "MATCH(LEFT(" & lookupFirstProcedureAddress & ",255),LEFT(" & sourceProceduresAddress & ",255),0), " & _
                                                            "MATCH(LEFT(" & lookupFirstVisitAddress & ",255),LEFT(" & sourceVisitsAddress & ",255),0)" & _
                                                         ")" & _
                                                         "=""""," & Chr(10) & _
                                    String(20, Chr(32)) & """""," & Chr(10) & _
                                    String(20, Chr(32)) & "INDEX(" & _
                                                            sourceValuesAddress & ", " & _
                                                            "MATCH(LEFT(" & lookupFirstProcedureAddress & ",255),LEFT(" & sourceProceduresAddress & ",255),0), " & _
                                                            "MATCH(LEFT(" & lookupFirstVisitAddress & ",255),LEFT(" & sourceVisitsAddress & ",255),0)" & _
                                                                ")" & Chr(10) & _
                                String(16, Chr(32)) & ")" & Chr(10) & _
                            String(12, Chr(32)) & ")" & Chr(10) & _
                        String(8, Chr(32)) & ")" & Chr(10)
                        
    formula = formula & _
                    String(4, Chr(32)) & ")" & Chr(10) & _
                ")"
    
    'add formula to the top cell
    destinationValuesRng.Cells(1, 1).Formula2 = formula
    
    'add formula to the entire range
    destinationValuesRng.Cells(1, 1).Copy
    destinationValuesRng.PasteSpecial Paste:=xlPasteFormulas
    
End Sub


