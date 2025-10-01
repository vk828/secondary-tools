Attribute VB_Name = "tool2_duplicates"
'Author/Developer: Vadim Krifuks
'Last Updated: 30 September 2025

Option Explicit
Option Private Module

Function FindAndHighlightDuplicatesTwoRanges(proceduresRng As Range, visitsRng As Range) As Boolean
'if either of the ranges have duplicates, the function hightlights them and returns true

    Dim areProceduresRepeat As Boolean
    Dim areVisitsRepeat As Boolean
    
    areProceduresRepeat = FindAndHighlightDuplicates(proceduresRng)
    areVisitsRepeat = FindAndHighlightDuplicates(visitsRng)

    If areProceduresRepeat Or areVisitsRepeat Then FindAndHighlightDuplicatesTwoRanges = True

End Function

Private Function FindAndHighlightDuplicates(rng As Range) As Boolean
'this function finds duplicates, highlights them, and returns true if they exist

    Dim cell As Range
    Dim dict As Object
    Dim key As String
    Dim cellValue As String
    
    'late binding dictionary
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' Count occurrences of cleaned & trimmed values
    For Each cell In rng
        If Not IsEmpty(cell.Value) Then
            cellValue = Application.WorksheetFunction.Trim(Application.WorksheetFunction.Clean(cell.Value))
            If dict.exists(cellValue) Then
                dict(cellValue) = dict(cellValue) + 1
            Else
                dict.Add cellValue, 1
            End If
        End If
    Next cell
    
    ' Highlight cells whose cleaned & trimmed values appear more than once and if so set a return boolean to true
    For Each cell In rng
        If Not IsEmpty(cell.Value) Then
            cellValue = Application.WorksheetFunction.Trim(Application.WorksheetFunction.Clean(cell.Value))
            If cellValue <> "" And dict.exists(cellValue) Then
                If dict(cellValue) > 1 Then
                    cell.Select
                    cell.Interior.color = vbRed
                    FindAndHighlightDuplicates = True
                End If
            End If
        End If
    Next cell
End Function
