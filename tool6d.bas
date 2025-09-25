Attribute VB_Name = "tool6d"
'Author/Developer: Vadim Krifuks
'Collaborators: Hui Zeng, Man Ming Tse
'Last Updated: 19January2025

Option Explicit

Sub tool6d_Add3ColAT3LookupFormulas()
'this subroutine adds lookup formulas for CptCodes, Costs, and NegRates
'user selects the appropriate ranges, and the subroutine adds lookup formulas

    Dim selectedRanges() As Range
    Dim title1, prompt1, title2, prompt2 As String
    Dim sourceProceduresRng, sourceCptCodesRng, sourceCostsRng, sourceNegRatesRng As Range
    Dim destinationProceduresRng, destinationCptCodesRng, destinationCostsRng, _
        destinationNegRatesRng As Range
    
    'STEP 1 - select and adjust SOURCE RANGES
    
    title1 = "Select Source Procedures, CPT Codes, Costs, and Negotiated Rates"
    prompt1 = "Select SOURCE PROCEDURES, CPT CODEs, COSTs, and NEGOTIATED RATEs range on the " & _
              "sheet you'd like to look up CPT CODEs, COSTs, NEGOTIATED RATEs FROM." _
              & Chr(10) & Chr(10) & _
              "Note: select a single four column range"


    title2 = "Select Destination Procedures, CPT Codes, Costs, and Negotiated Rates"
    prompt2 = "Select PROCEDURES, CPT CODEs, COSTs, NEGOTIATED RATEs range " & _
              "where you'd like to add lookup formulas for CPT CODEs, COSTs, and NEGOTIATED RATEs." _
              & Chr(10) & Chr(10) & _
              "Note: select a single four column range"

    selectedRanges = SelectTwoRanges(title1, prompt1, title2, prompt2)
    'if user cancelled selection, stop execution
    If (selectedRanges(0) Is Nothing) Or (selectedRanges(1) Is Nothing) Then
        Exit Sub
    End If

    Set sourceProceduresRng = selectedRanges(0).Columns(1)
    Set sourceCptCodesRng = selectedRanges(0).Columns(2)
    Set sourceCostsRng = selectedRanges(0).Columns(3)
    Set sourceNegRatesRng = selectedRanges(0).Columns(4)
    
    Set destinationProceduresRng = selectedRanges(1).Columns(1)
    Set destinationCptCodesRng = selectedRanges(1).Columns(2)
    Set destinationCostsRng = selectedRanges(1).Columns(3)
    Set destinationNegRatesRng = selectedRanges(1).Columns(4)
    
    'STEP 3 - add LOOKUP FORMULAS
    Call tool6a.GenerateAndFillVerticalLookupFormulas(sourceProceduresRng, _
                                                      sourceCptCodesRng, _
                                                      destinationProceduresRng, _
                                                      destinationCptCodesRng)
                                            
    Call GenerateAndFillVerticalLookupFormulas(sourceProceduresRng, sourceCostsRng, _
                                            destinationProceduresRng, destinationCostsRng)

    Call GenerateAndFillVerticalLookupFormulas(sourceProceduresRng, sourceNegRatesRng, _
                                            destinationProceduresRng, destinationNegRatesRng)
    
    Call AddColorToRange(destinationNegRatesRng, RGB(204, 204, 255))

End Sub

Private Function SelectTwoRanges(ByVal titleFirstSelection As String, _
                                 ByVal promptFirstSelection As String, _
                                 ByVal titleSecondSelection As String, _
                                 ByVal promptSecondSelection As String) As Range()
'this funciton: lets the user select two ranges, and
'returns an array that holds two selected ranges

    Dim returnArray(0 To 1) As Range
   
    'loop until either user cancels selection or until ranges are selected correctly
        
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

Done:
    SelectTwoRanges = returnArray

End Function

Private Sub GenerateAndFillVerticalLookupFormulas(ByVal sourceProceduresRng As Range, _
                                                  ByVal sourceLookupRng As Range, _
                                                  ByVal destinationProceduresRng As Range, _
                                                  ByVal destinationLookupRng As Range)
'this subroutine generates the the lookup formula for the top cell and then uses it
'to generate the entire range
'if lookup value is "", value is 0

'=IF($I16="",
'     0,
'    IF(ISNA(MATCH(LEFT($I16,255),LEFT(Budget_Details_ADJ_DBL!$I$16:$I$199,255),0)),
'        "NO RESULT FOR " & $I16,
'        (INDEX(Budget_Details_ADJ_DBL!$G$16:$G$199, MATCH(LEFT($I16,255), LEFT( Budget_Details_ADJ_DBL!$I$16:$I$199,255), 0)))
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
                  String(4, Chr(32)) & 0 & "," & Chr(10) & _
                  String(4, Chr(32)) & "IF(ISNA(MATCH(LEFT(" & destinationTopProcedureAddress & ",255),LEFT(" & sourceProceduresAddress & ",255),0))," & Chr(10) & _
                      String(8, Chr(32)) & notFoundMsg & destinationTopProcedureAddress & "," & Chr(10) & _
                      String(8, Chr(32)) & "(INDEX(" & sourceValuesAddress & ", MATCH(LEFT(" & destinationTopProcedureAddress & ",255), LEFT( " & sourceProceduresAddress & ",255), 0)))" & Chr(10) & _
                  String(4, Chr(32)) & ")" & Chr(10) & _
              ")"
    
    'add formula to the top cell
    destinationLookupRng.Cells(1, 1).Formula2 = formula
    
    'add formula to the entire range
    destinationLookupRng.Cells(1, 1).Copy
    destinationLookupRng.PasteSpecial Paste:=xlPasteFormulas
    
End Sub

Private Sub AddColorToRange(Rng As Range, colorCode As Long)
'this subroutine sets a specified background color to a range

    Rng.Interior.color = colorCode
End Sub

