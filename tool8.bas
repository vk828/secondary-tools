Attribute VB_Name = "tool8"
Option Explicit


Sub tool8_RemoveFootnotes()
'this subroutine removes footnotes from a user selected range

    Dim Rng As Range
    
    'on Error Resume Next takes the program to the next line in case the user cancels selecting a range
    On Error Resume Next
    'get a range input from the user
    Set Rng = Utilities.SelectRange("Footnotes Range", "Select a range of cell that contains footnotes (characters in superscript)" & _
                                    " you'd like to remove and click OK.")
    
    'test to ensure the user did not cancel
    If Rng Is Nothing Then Exit Sub
  
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Call Utilities.RemoveFootnotesFromSelectedRange(Rng)
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

End Sub

