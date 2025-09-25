Attribute VB_Name = "tool2_overwriteNames"
Option Explicit
Option Private Module

Sub OverwriteProcedureAndVisitNames(ib_proceduresRng As Range, ib_visitsRng As Range, _
                                    oncore_proceduresRng As Range, oncore_visitsRng As Range)

    Dim ib_unpairedProceduresColl As Collection
    Dim oncore_unpairedProceduresColl As Collection

    Dim ib_unpairedVisitsColl As Collection
    Dim oncore_unpairedVisitsColl As Collection

    Dim collArray As Variant
    
    'get unpaired procedures lists
    collArray = GetUnpairedLists(ib_proceduresRng, oncore_proceduresRng)
    Set ib_unpairedProceduresColl = collArray(1)
    Set oncore_unpairedProceduresColl = collArray(2)
    
    'get unpaired visits lists
    collArray = GetUnpairedLists(ib_visitsRng, oncore_visitsRng)
    Set ib_unpairedVisitsColl = collArray(1)
    Set oncore_unpairedVisitsColl = collArray(2)
    
    'open user form and give user options to choose pairs
    'when a pair is selected update a procedure or visit name on int bdgt
    'remove the pair from both lists

End Sub

Private Function GetUnpairedLists(firstRng As Range, secondRng As Range) As Variant
'this function takes two ranges and returns an array of two collections containing only unpaired items

    Dim cell As Range
    
    Dim collArray(1 To 2) As Collection
    Set collArray(1) = New Collection
    Set collArray(2) = New Collection
    
    
    Dim collOneItem As String, collTwoItem As String
    Dim i As Integer, j As Integer
    
    'convert range to collection; values are cleaned and trimmed before they are stored in a collection
    For Each cell In firstRng.Cells
        collArray(1).Add Application.WorksheetFunction.Trim(Application.WorksheetFunction.Clean(cell.Value))
    Next cell
    
    'convert range to collection; values are cleaned and trimmed before they are stored in a collection
    For Each cell In secondRng.Cells
        collArray(2).Add Application.WorksheetFunction.Trim(Application.WorksheetFunction.Clean(cell.Value))
    Next cell
    
    'compare item by item and remove matching pairs from the collections
    For i = collArray(1).count To 1 Step -1
        collOneItem = collArray(1).item(i)
        For j = collArray(2).count To 1 Step -1
            collTwoItem = collArray(2).item(j)
            If collOneItem = collTwoItem Then
                collArray(1).Remove i
                collArray(2).Remove j
                Exit For
            End If
        Next j
    Next i
        
    GetUnpairedLists = collArray
    
End Function


