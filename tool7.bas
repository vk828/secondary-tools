Attribute VB_Name = "tool7"
Option Explicit


Sub tool7_FillEmptyNoFillCellsGray()
'this subroutine lets the user select a range and then goes cell by cell of this range and
'add background color if cell is empty and doesn't have a background color

    Dim rng As Range
    Dim title, prompt As String
    
    title = "Select Range"
    prompt = "Select a range of cells that you'd like the program to process. It will highlight " & _
             "empty cells that don't have a background color and highlight them with gray."

    Set rng = Utilities.SelectRange(title, prompt)
    
    'if user cancelled selection, exit sub
    If rng Is Nothing Then
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Call FillBlank(rng, RGB(217, 217, 217))
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

End Sub

Private Sub FillBlank(rng As Range, color As Long)
'this subroutine highlights cells that are empty and don't hve a background color
    Dim cel As Range

    For Each cel In rng.Cells
        If cel.Value = "" And cel.Interior.ColorIndex = xlNone Then
            cel.Interior.color = color
        End If
    Next cel
   
End Sub
