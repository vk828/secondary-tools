Attribute VB_Name = "tool9_utilities"
'Author/Developer: Vadim Krifuks
'Collaborators: Hui Zeng, Man Ming Tse
'Last Updated: 5Feb2025

Option Explicit


Function SelectTwoRanges(ByRef unitRatesRng As Range, ByRef totalsRng As Range, str As String)
'this function asks user to select two ranges and returns them if they are valid

    Dim rowsVerificationResult As Integer

'when user selects wrong ranges, rows don't match, there is an option to retry
'clicking retry takes the program execution here
startAgain:

    'on Error Resume Next takes the program to the next line in case the user cancels selecting a range
    On Error Resume Next
    'get a range input from the user for Unit Rates
    Set unitRatesRng = Utilities.SelectRange("Unit Rates Range", "Select a range that contains Unit Rates and click OK.")
    
    'on Error Resume Next takes the program to the next line in case the user cancels selecting a range
    On Error Resume Next
    'get a range input from the user for Totals
    Set totalsRng = Utilities.SelectRange("Totals Range", _
                                          "Select a range that contains totals that you'd like " & _
                                          "to " & str & " and click OK.")
                    
    'if user clicked canceled in either or both of the input boxes, execution stops because
    'the returned range from inputBox is Nothing
    If unitRatesRng Is Nothing Or totalsRng Is Nothing Then
        SelectTwoRanges = 1
        Exit Function
    End If

    'rows verification to make sure rows of two selected ranges align
    rowsVerificationResult = VerifyRowsOfTwoRangesAlign("Unit Rates", unitRatesRng, "Totals", totalsRng)
    '0 means ranges align
    '4 means ranges don't align and user decided to select new ranges
    '2 means ranges don't align and user decided to cancel (either clicked cancel button or closed the window)
    Select Case rowsVerificationResult
    Case 0:
    Case 4:
        GoTo startAgain
    Case 2:
        SelectTwoRanges = 1
        Exit Function
    End Select
        
    SelectTwoRanges = 0
    
End Function

Function VerifyRowsOfTwoRangesAlign(rangeOneName As String, firstRng As Range, rangeTwoName As String, secondRng As Range)
'this function compares rows of two ranges and returns 0 if they align
'if they don't, returns 4 if the user clicks retry, or 2 for cancel
        
    Dim unitRatesRngTopRow As Integer
    Dim unitRatesRngBottomRow As Integer
    Dim totalsRngTopRow As Integer
    Dim totalsRngBottomRow As Integer
    
    Dim errorString As String
    
    unitRatesRngTopRow = firstRng.Cells(1).Row
    unitRatesRngBottomRow = firstRng.Cells(firstRng.Rows.count, 1).Row
    
    totalsRngTopRow = secondRng.Cells(1).Row
    totalsRngBottomRow = secondRng.Cells(secondRng.Rows.count, 1).Row
    
    If unitRatesRngTopRow <> totalsRngTopRow Or unitRatesRngBottomRow <> totalsRngBottomRow Then
        
        errorString = "You selected the following ranges" & Chr(10) _
                        & "  " & rangeOneName & ": " & firstRng.Address(False, False, external:=True) & Chr(10) _
                        & "  " & rangeTwoName & ": " & secondRng.Address(False, False, external:=True) & Chr(10) _
                        & "The rows between selected ranges MUST match. Please try again."
        
        VerifyRowsOfTwoRangesAlign = MsgBox(prompt:=errorString, Buttons:=vbRetryCancel + vbExclamation, title:="Error in Rows")
    Else
        VerifyRowsOfTwoRangesAlign = 0
    End If
End Function
