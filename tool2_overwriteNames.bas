Attribute VB_Name = "tool2_overwriteNames"
'Author/Developer: Vadim Krifuks
'Last Updated: 1 October 2025

Option Explicit
Option Private Module

Private frmProcedures As frmTool2PairNames
Private frmVisits As frmTool2PairNames

Private wbReport As Workbook

Function OverwriteProcedureAndVisitNames(ib_proceduresRng As Range, _
                                    ib_visitsRng As Range, _
                                    oncore_proceduresRng As Range, _
                                    oncore_visitsRng As Range) As Boolean
'function overwrites procedures and visit names and returns true if user decides
'to terminate the program early, false otherwise

    Dim ib_unpairedProceduresColl As Collection
    Dim oncore_unpairedProceduresColl As Collection

    Dim ib_unpairedVisitsColl As Collection
    Dim oncore_unpairedVisitsColl As Collection
    
    Dim instrStart As String
    Dim instrEnd As String
        
    Dim terminateFlag As Boolean
    
    instrStart = "The following names do not have a match. You can manually select the appropriate pair(s) one by one and " _
                & "click 'Pair Selected' to overwrite the "
    instrEnd = " on the internal budget. When you are ready to start updating the grid, please click 'Done.' You can also " _
                & "generate a list of unpaired items or abort the program altogether."

    Dim collArray As Variant
    
    '### STEP1 ###
    'get unpaired procedures lists
    collArray = GetUnpairedLists(ib_proceduresRng, oncore_proceduresRng)
    Set ib_unpairedProceduresColl = collArray(1)
    Set oncore_unpairedProceduresColl = collArray(2)
    
    '### STEP2 ###
    'get unpaired visits lists
    collArray = GetUnpairedLists(ib_visitsRng, oncore_visitsRng)
    Set ib_unpairedVisitsColl = collArray(1)
    Set oncore_unpairedVisitsColl = collArray(2)
    
    '### STEP3 ###
    'switch active sheet, so the forms open in that window
    Call SwitchActiveSheet(ib_proceduresRng)
    
    '### STEP4 ###
    'pass Int Bdgt ranges to the appropriate user forms to enable the progrma to update names on Int Bdgt
    Set frmProcedures = New frmTool2PairNames
    Set frmVisits = New frmTool2PairNames
    Set frmProcedures.RangeToForm = ib_proceduresRng
    Set frmVisits.RangeToForm = ib_visitsRng

    '### STEP5 ###
    'set comboboxes on procedures user form
    Call frmProcedures.SetCbos(ib_unpairedProceduresColl, _
                                oncore_unpairedProceduresColl, _
                                "Update Procedure Names on Internal Budget", _
                                instrStart & "procedure name(s)" & instrEnd, _
                                "PROCEED to Pairing Visit Names")
    
    'set comboboxes on visits user form
    Call frmVisits.SetCbos(ib_unpairedVisitsColl, _
                            oncore_unpairedVisitsColl, _
                            "Update Visit Names on Internal Budget", _
                            instrStart & "visit name(s)" & instrEnd, _
                            "PROCEED to Updating Grid")
    
    'show the user form
    Call ShowModelessFormAndPause(frmProcedures)
    'if true unload both forms and exit sub
    If frmProcedures.IsTerminated Then
        terminateFlag = CloseAndCleanUp(frmProcedures, frmVisits, oncore_visitsRng)
        GoTo SkipVisitNames
    End If
    
    'show the user form
    Call ShowModelessFormAndPause(frmVisits)
    'if true unload both forms and exit sub
    If frmVisits.IsTerminated Then
        terminateFlag = CloseAndCleanUp(frmProcedures, frmVisits, oncore_visitsRng)
    End If
    
SkipVisitNames:
    'if 'exit' was pressed, forms have been unloaded
    If terminateFlag Then
        OverwriteProcedureAndVisitNames = terminateFlag
    'otherwise unload the forms
    Else
        Unload frmProcedures
        Unload frmVisits
        Set wbReport = Nothing
    End If
End Function

Private Function IsWorkbookOpen(wbName As String) As Boolean
    Dim wb As Workbook
    On Error Resume Next
    Set wb = Workbooks(wbName)
    IsWorkbookOpen = Not wb Is Nothing
    Set wb = Nothing
    On Error GoTo 0
End Function

Sub ReportUnpaired()
'subroutine creates a workbook and puts unpaired procedures and visits on it

    Dim i As Integer
    'if this is not the the first time getting a report, check if the worbook is open,
    'if it is, clear the report workbook and populate it with information
    'otherwise, create a report
    If Not wbReport Is Nothing Then
        'is it safe to operate on wbReport
        If IsWorkbookOpen(wbReport.name) Then
            wbReport.Sheets(1).Columns("A:E").Clear
            wbReport.Activate
        End If
    Else
        Set wbReport = Workbooks.Add
    End If
    
    With wbReport.Sheets(1)
    
        .Cells(1, 1).Value = "Unpaired Procedure Names"
        .Range(.Cells(1, 1), .Cells(1, 2)).Merge
        .Cells(2, 1).Value = "Internal Budget"
        .Cells(2, 2).Value = "OnCore"
        
        For i = 1 To frmProcedures.IbCollection.count
            .Cells(i + 2, 1).Value = frmProcedures.IbCollection(i)
        Next i
        
        For i = 1 To frmProcedures.OncoreCollection.count
            .Cells(i + 2, 2).Value = frmProcedures.OncoreCollection(i)
        Next i
        
        .Cells(1, 4).Value = "Unpaired Visit Names"
        .Range(.Cells(1, 4), .Cells(1, 5)).Merge
        .Cells(2, 4).Value = "Internal Budget"
        .Cells(2, 5).Value = "OnCore"
        
        For i = 1 To frmVisits.IbCollection.count
            .Cells(i + 2, 4).Value = frmVisits.IbCollection(i)
        Next i
        
        For i = 1 To frmVisits.OncoreCollection.count
            .Cells(i + 2, 5).Value = frmVisits.OncoreCollection(i)
        Next i
        
        With .rows("1:2")
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
        End With
        
        .Columns("A:B").ColumnWidth = 60
        .Columns("D:E").ColumnWidth = 60
        
        .Columns("A:B").WrapText = True
        .Columns("D:E").WrapText = True
        
    End With
    
End Sub

Private Function CloseAndCleanUp(frmOne As frmTool2PairNames, frmTwo As frmTool2PairNames, rng As Range) As Boolean
    Unload frmOne
    Unload frmTwo
    Set wbReport = Nothing
    rng.Worksheet.Parent.Close SaveChanges:=False 'close workbook
    CloseAndCleanUp = True
End Function

Private Sub ShowModelessFormAndPause(frm As frmTool2PairNames)
    
    With frm
        'manual positioning
        .StartUpPosition = 0
        
        'set left and top location
        'currently set to bottom right of excel app
        .Left = Application.Left + (1 * Application.Width) - (1 * .Width)
        .Top = Application.Top + (1 * Application.Height) - (1 * .Height)
        
        .Show vbModeless
    End With
    
    Do While frm.Visible
        DoEvents  ' Allows Excel to respond to events
    Loop
    
    ' Code here will run after the UserForm is closed or hidden
    'MsgBox "Form closed, code resumes."
End Sub

Private Sub SwitchActiveSheet(rng As Range)

    'activates workbook
    rng.Worksheet.Parent.Activate
    
    'activates worksheet
    rng.Worksheet.Activate

End Sub

Private Function GetUnpairedLists(firstRng As Range, secondRng As Range) As Variant
'this function takes two ranges and returns an array of two collections containing only unpaired items

    Dim cell As Range
    Dim name As String
    
    Dim collArray(1 To 2) As Collection
    Set collArray(1) = New Collection
    Set collArray(2) = New Collection
    
    Dim collOneItem As String, collTwoItem As String
    Dim i As Integer, j As Integer
    
    'convert range to collection; values are cleaned and trimmed before they are stored in a collection
    'all "" are skipped
    For Each cell In firstRng.Cells
        name = Application.WorksheetFunction.Trim(Application.WorksheetFunction.Clean(cell.Value))
        If name <> "" Then collArray(1).Add name
    Next cell
    
    'convert range to collection; values are cleaned and trimmed before they are stored in a collection
    'all "" are skipped
    For Each cell In secondRng.Cells
        name = Application.WorksheetFunction.Trim(Application.WorksheetFunction.Clean(cell.Value))
        If name <> "" Then collArray(2).Add name
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


