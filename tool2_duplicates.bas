Attribute VB_Name = "tool2_duplicates"
'Author/Developer: Vadim Krifuks
'Collaborators: Man Ming Tse
'Last Updated: 13 October 2026

Option Explicit
Option Private Module

Function FindAndHighlightDuplicatesTwoRanges(proceduresRng As Range, visitsRng As Range) As Boolean
'if either of the ranges have duplicates, the function hightlights them and returns true

    Dim areProceduresRepeat As Boolean
    Dim areVisitsRepeat As Boolean
    
    areProceduresRepeat = FindAndHighlightProcedureDuplicates(proceduresRng)
    areVisitsRepeat = FindAndHighlightVisitDuplicates(visitsRng)

    If areProceduresRepeat Or areVisitsRepeat Then FindAndHighlightDuplicatesTwoRanges = True

End Function

Private Function FindAndHighlightProcedureDuplicates(rng As Range) As Boolean
'this function finds duplicates, highlights them, and returns true if they exist

    Dim cell As Range
    Dim dict As Object
    Dim cellValue As String
    Dim duplicateFound As Boolean
    
    'late binding dictionary
    Set dict = CreateObject("Scripting.Dictionary")
    
    duplicateFound = False
    
    ' Count occurrences of cleaned & trimmed values
    For Each cell In rng
        If Not IsEmpty(cell.value) Then
            cellValue = Application.WorksheetFunction.Trim(Application.WorksheetFunction.Clean(cell.value))
            If dict.exists(cellValue) Then
                dict(cellValue) = dict(cellValue) + 1
            Else
                dict.Add cellValue, 1
            End If
        End If
    Next cell
    
    ' Highlight cells whose cleaned & trimmed values appear more than once and if so set a return boolean to true
    For Each cell In rng
        If Not IsEmpty(cell.value) Then
            cellValue = Application.WorksheetFunction.Trim(Application.WorksheetFunction.Clean(cell.value))
            If cellValue <> "" And dict.exists(cellValue) Then
                If dict(cellValue) > 1 Then
                    cell.Select
                    cell.Interior.color = vbRed
                    duplicateFound = True
                End If
            End If
        End If
    Next cell

    FindAndHighlightProcedureDuplicates = duplicateFound
End Function

Private Function FindAndHighlightVisitDuplicates(rng As Range) As Boolean
'this function finds duplicates, highlights them, and returns true if they exist

    Dim visitCell As Range
    Dim dict As Object
    Dim visitName As String
    Dim segmentName As String
    Dim combinedSegmentVisitName As String
    Dim duplicateFound As Boolean
    
    'late binding dictionary
    Set dict = CreateObject("Scripting.Dictionary")
    
    duplicateFound = False
    
    ' Count occurrences of cleaned & trimmed values
    For Each visitCell In rng
        If Not IsEmpty(visitCell.value) Then
            visitName = visitCell.value
            segmentName = visitCell.Offset(-1, 0).value
            If segmentName = "" Then
                combinedSegmentVisitName = visitName
            Else
                combinedSegmentVisitName = visitName & " " & Split(segmentName, Chr(10))(0)
            End If
                
            combinedSegmentVisitName = Application.WorksheetFunction.Trim(Application.WorksheetFunction.Clean(combinedSegmentVisitName))
            
            If dict.exists(combinedSegmentVisitName) Then
                dict(combinedSegmentVisitName) = dict(combinedSegmentVisitName) + 1
            Else
                dict.Add combinedSegmentVisitName, 1
            End If
        End If
    Next visitCell
    
    ' Highlight cells whose cleaned & trimmed values appear more than once and if so set a return boolean to true
    For Each visitCell In rng
        If Not IsEmpty(visitCell.value) Then
            visitName = visitCell.value
            segmentName = visitCell.Offset(-1, 0).value
            
            If segmentName = "" Then
                combinedSegmentVisitName = visitName
            Else
                combinedSegmentVisitName = visitName & " " & Split(segmentName, Chr(10))(0)
            End If
            
            combinedSegmentVisitName = Application.WorksheetFunction.Trim(Application.WorksheetFunction.Clean(combinedSegmentVisitName))
            
            If combinedSegmentVisitName <> "" And dict.exists(combinedSegmentVisitName) Then
                If dict(combinedSegmentVisitName) > 1 Then
                    visitCell.Select
                    visitCell.Interior.color = vbRed
                    visitCell.Offset(-1, 0).Interior.color = vbRed
                    duplicateFound = True
                End If
            End If
        End If
    Next visitCell

    FindAndHighlightVisitDuplicates = duplicateFound
End Function
