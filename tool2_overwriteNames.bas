Attribute VB_Name = "tool2_overwriteNames"
'Author/Developer: Vadim Krifuks
'Collaborators: Man Ming Tse
'Last Updated: 13 October 2026

Option Explicit
Option Private Module

'forms to pair procedures and visits
Private frmProcedures As frmTool2PairIntBdgtNames
Private frmVisits As frmTool2PairIntBdgtNames
Private ib_visitsRow As Range
Private oncore_visitsRow As Range

'wbReport holds reference to 'unpaired procedures/visits' workbook
'if this workbook is open, variable is NOT nothing
Private wbReport As Workbook

Function OverwriteProcedureAndVisitNames(ib_proceduresArray() As Variant, _
                                    ib_visitsArray() As Variant, _
                                    oncore_proceduresArray() As Variant, _
                                    oncore_visitsArray() As Variant, _
                                    ib_proceduresRng As Range, _
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
    
    'this variable is true if CloseAndCleanUp was called
    Dim terminateFlag As Boolean
    
    instrStart = "The following names do not have a match. You can manually select the appropriate pair(s) one by one and " _
                & "click 'Pair Selected' to overwrite the "
    instrEnd = " on the internal budget. When you are ready to start updating the grid, please click 'PROCEED to Updating Grid'. You can also " _
                & "generate a list of unpaired items or abort the program altogether."

    Dim collArray As Variant
    
    '### STEP1 ###
    'get unpaired procedures lists
    collArray = GetUnpairedLists(ib_proceduresArray, oncore_proceduresArray)
    Set ib_unpairedProceduresColl = collArray(1)
    Set oncore_unpairedProceduresColl = collArray(2)
    
    '### STEP2 ###
    'get unpaired visits lists
    collArray = GetUnpairedLists(ib_visitsArray, oncore_visitsArray)
    Set ib_unpairedVisitsColl = collArray(1)
    Set oncore_unpairedVisitsColl = collArray(2)
    
    '### STEP3 ###
    'switch active sheet, so the forms open in that window
    Call Utilities.SwitchActiveSheet(ib_proceduresRng)
    
    '### STEP4 ###
    'pass Int Bdgt ranges to the appropriate user forms to enable the program to update names on Int Bdgt
    Set frmProcedures = New frmTool2PairIntBdgtNames
    Set frmVisits = New frmTool2PairIntBdgtNames
'    Set frmProcedures.RangeToForm = ib_proceduresRng
'    Set frmVisits.RangeToForm = ib_visitsRng

    '### STEP5 ###
    'set comboboxes on procedures user form
    Call frmProcedures.SetCbos(True, _
                                ib_unpairedProceduresColl, _
                                oncore_unpairedProceduresColl, _
                                ib_proceduresArray, _
                                oncore_proceduresArray, _
                                ib_proceduresRng, _
                                oncore_proceduresRng, _
                                "Update Procedure Names on Internal Budget", _
                                instrStart & "procedure name(s)" & instrEnd, _
                                "PROCEED to Pairing Visit Names")
    
    'set comboboxes on visits user form
    Call frmVisits.SetCbos(False, _
                            ib_unpairedVisitsColl, _
                            oncore_unpairedVisitsColl, _
                            ib_visitsArray, _
                            oncore_visitsArray, _
                            ib_visitsRng, _
                            oncore_visitsRng, _
                            "Update Visit and Segment Names on Internal Budget", _
                            instrStart & "visit and segment name(s)" & instrEnd, _
                            "PROCEED to Updating Grid")
    
    'show the user form
    Call ShowModelessFormAndPause(frmProcedures)
    'if true unload both forms and exit sub
    If frmProcedures.IsTerminated Then
        terminateFlag = CloseAndCleanUp(frmProcedures, frmVisits, oncore_proceduresRng)
        GoTo SkipVisitNames
    End If
    
    'show the user form
    Call ShowModelessFormAndPause_Visits(frmVisits)
    'if true unload both forms and exit sub
    If frmVisits.IsTerminated Then
        terminateFlag = CloseAndCleanUp(frmProcedures, frmVisits, oncore_proceduresRng)
    End If
    
SkipVisitNames:
    'if 'exit' was pressed, forms have been unloaded
    If terminateFlag Then
        'function return is set to TRUE if form(s) were exited intentionaly
        OverwriteProcedureAndVisitNames = terminateFlag
    
    'otherwise, conditionally update report and unload the forms
    Else
        'if wbReport exists, update it
        If Not wbReport Is Nothing Then
            ReportUnpaired
        End If

        ib_visitsArray = frmVisits.IbArray
        oncore_visitsArray = frmVisits.OncoreArray

        Unload frmProcedures
        Unload frmVisits
        Set wbReport = Nothing
    End If
End Function

Private Function IsWorkbookOpen(wbName As String) As Boolean
'function returns true if workbook is open, false otherwise

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
    
    Dim shUnpairedList As String
    Dim shUnpairedMap As String
    Dim ws As Worksheet
    Dim note As String
    Dim visitsFirstCell As Range
    Dim proceduresFirstCell As Range
    Dim totalVisits As Integer
    Dim totalProcedures As Integer
    
    shUnpairedList = "Unpaired Proc and Visits List"
    shUnpairedMap = "Unpaired - OnCore Map"
    
    
    If Not wbReport Is Nothing Then
        'is it safe to operate on wbReport
        If IsWorkbookOpen(wbReport.name) Then
        
            ' Temporarily ignore errors
            On Error Resume Next
            Set ws = wbReport.Sheets(shUnpairedList)
            On Error GoTo 0 ' Turn standard error tracking back on immediately
            
            ' If ws is Nothing, the sheet does not exist
            If ws Is Nothing Then
                Set ws = wbReport.Sheets.Add(before:=wbReport.Sheets(1))
                ws.name = shUnpairedList
            End If
            
            ' Temporarily ignore errors
            On Error Resume Next
            Set ws = wbReport.Sheets(shUnpairedMap)
            On Error GoTo 0 ' Turn standard error tracking back on immediately
            
            ' If ws is Nothing, the sheet does not exist
            If ws Is Nothing Then
                Set ws = wbReport.Sheets.Add(after:=wbReport.Sheets(shUnpairedList))
                ws.name = shUnpairedMap
            End If
            
            'clear content on both sheets
            wbReport.Sheets(shUnpairedList).Columns("A:G").Clear
            wbReport.Sheets(shUnpairedMap).Cells.Clear
            wbReport.Activate
        End If
    Else
        Set wbReport = Workbooks.Add
        Set ws = wbReport.Sheets.Add(before:=wbReport.Sheets(1))
        ws.name = shUnpairedList
        Set ws = wbReport.Sheets.Add(after:=wbReport.Sheets(shUnpairedList))
        ws.name = shUnpairedMap
    End If
        
    ' *** UPDAGE shUnpairedList sheet ***
    With wbReport.Sheets(shUnpairedList)
    
        .Cells(1, 1).value = "Unpaired Procedure Names"
        .Range(.Cells(1, 1), .Cells(1, 2)).Merge
        .Cells(2, 1).value = "Internal Budget"
        .Cells(2, 2).value = "OnCore"
        .Cells(2, 1).Resize(2).Merge
        .Cells(2, 2).Resize(2).Merge
        
        For i = 1 To frmProcedures.IbCollection.count
            .Cells(i + 3, 1).value = frmProcedures.IbArray(3, frmProcedures.IbCollection(i))
        Next i
        
        For i = 1 To frmProcedures.OncoreCollection.count
            .Cells(i + 3, 2).value = frmProcedures.OncoreArray(3, frmProcedures.OncoreCollection(i))
        Next i
        
        .Cells(1, 4).value = "Unpaired Visit Names"
        .Range(.Cells(1, 4), .Cells(1, 7)).Merge
        .Cells(2, 4).value = "Internal Budget"
        .Range(.Cells(2, 4), .Cells(2, 5)).Merge
        .Cells(2, 6).value = "OnCore"
        .Range(.Cells(2, 6), .Cells(2, 7)).Merge
        .Range(.Cells(3, 4), .Cells(3, 6)).value = "Segment"
        .Range(.Cells(3, 5), .Cells(3, 7)).value = "Visit"
        
        For i = 1 To frmVisits.IbCollection.count
            .Cells(i + 3, 4).value = frmVisits.IbArray(1, frmVisits.IbCollection(i))
            .Cells(i + 3, 5).value = frmVisits.IbArray(2, frmVisits.IbCollection(i))
        Next i
        
        For i = 1 To frmVisits.OncoreCollection.count
            .Cells(i + 3, 6).value = frmVisits.OncoreArray(1, frmVisits.OncoreCollection(i))
            .Cells(i + 3, 7).value = frmVisits.OncoreArray(2, frmVisits.OncoreCollection(i))
        Next i
        
        With .Rows("1:3")
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
        End With
        
        .Range("A:B, D:G").ColumnWidth = 45
        .Columns("C").ColumnWidth = 5
        
        With .Columns("A:B")
            .WrapText = True
            .VerticalAlignment = xlCenter
        End With

    End With
    
    ' *** UPDAGE shUnpairedMap sheet ***
    ' insert note
    With wbReport.Sheets(shUnpairedMap)
        note = "Note to Analyst: The following are OnCore headers. Procedure(s) and/or Visit(s) in bold black font currently do NOT have a matching pair on the Internal Budget." _
        & Chr(10) & "Internal Budget options:" _
        & Chr(10) & " - pair equivalents: use the tool to pair equivalent procedure or visit names" _
        & Chr(10) & " - expand selection: run the tool again and adjust and expand the selected ranges to capture the items" _
        & Chr(10) & " - add new rows/columns: insert the missing procedures or visits directly into the Internal Budget and run the tool again" _
        & Chr(10) & " - do nothing: ignore if these specific bold items are intentionally meant to be excluded"
        
        With .Cells(1.1)
            .ColumnWidth = 110
            .Font.color = vbRed
            .Font.Italic = True
            .value = note
        End With
    End With
    
    'insert all procedure names from oncore
    Set proceduresFirstCell = wbReport.Sheets(shUnpairedMap).Cells(4, 1)
    totalProcedures = UBound(frmProcedures.OncoreArray, 2)
    
    With proceduresFirstCell.Resize(totalProcedures, 1)
        .value = Application.Transpose(Application.index(frmProcedures.OncoreArray, 3, 0))
        .Font.color = RGB(166, 166, 166)
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    
    'highlight visit-segment pairs that don't have a pair on internal budget
    With proceduresFirstCell
        For i = 1 To totalProcedures
            If frmProcedures.OncoreArray(5, i) = 0 Then
                With .Offset(i - 1, 0).Font
                    .color = vbBlack
                    .Bold = True
                End With
            End If
        Next i
    End With
    
    'insert all visit-segment names from oncore
    Set visitsFirstCell = wbReport.Sheets(shUnpairedMap).Cells(3, 2)
    totalVisits = UBound(frmVisits.OncoreArray, 2)
    
    With visitsFirstCell.Offset(-1).Resize(2, totalVisits)
        .value = frmVisits.OncoreArray
        .Font.color = RGB(166, 166, 166)
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    
    'highlight visit-segment pairs that don't have a pair on internal budget
    With visitsFirstCell
        For i = 1 To totalVisits
            If frmVisits.OncoreArray(5, i) = 0 Then
                With .Offset(-1, i - 1).Resize(2, 1).Font
                    .color = vbBlack
                    .Bold = True
                End With
            End If
        Next i
    End With
    
End Sub

Private Function CloseAndCleanUp(frmOne As frmTool2PairIntBdgtNames, frmTwo As frmTool2PairIntBdgtNames, rng As Range) As Boolean
'this function conditionally updates 'unpaired report' workbook if it is open, unloads forms, and closes oncore workbook
'returns true upon completion
    
    'if wbReport exists, update it
    If Not wbReport Is Nothing Then
        ReportUnpaired
    End If
    
    Unload frmOne
    Unload frmTwo
    
    Set wbReport = Nothing
    
    'close onCore workbook
    rng.Worksheet.Parent.Close SaveChanges:=False
    
    CloseAndCleanUp = True
End Function

Private Sub ShowModelessFormAndPause(frm As frmTool2PairIntBdgtNames)
'show form
    
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

Private Sub ShowModelessFormAndPause_Visits(frm As frmTool2PairIntBdgtNames)
'show form
    
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

Private Function GetUnpairedLists(ib_array() As Variant, oncore_array() As Variant) As Variant

    Dim cell As Range
    Dim name As String
    
    Dim collArray(1 To 2) As Collection
    Set collArray(1) = New Collection
    Set collArray(2) = New Collection
    
    Dim collOneItem As String, collTwoItem As String
    Dim i As Integer, j As Integer
    
    Dim ib_index As Integer
    Dim oncore_index As Integer
    
    For i = 1 To UBound(ib_array, 2)
        name = ib_array(3, i)
        
        If name = "" Then
            ' a number in row 5 greater than 0 means there is a matching visit in OnCore
            ' -1 means this column is not a visit
            ib_array(5, i) = -1
        Else
        
            For j = 1 To UBound(oncore_array, 2)
                If LCase(Trim(name)) = LCase(Trim(oncore_array(3, j))) Then
                    
                    ' 1 means that this visit has a match on the internal budget
                    oncore_array(5, j) = 1
                    
                    ' maps column of the matching visit on OnCore
                    ' a number in row 5 greater than 0 means there is a matching visit in OnCore
                    ib_array(5, i) = oncore_array(4, j)
                End If
            Next j
        End If
    Next i
    
    'build a collection of indices for visits listed on the internal budget that don't have a matching pair on oncore
    For i = 1 To UBound(ib_array, 2)
        If ib_array(5, i) = 0 Then collArray(1).Add i
    Next i
    
    'build a collection of indices for visits listed on oncore that don't have a matching pair on the internal budget
    For i = 1 To UBound(oncore_array, 2)
        If oncore_array(5, i) = 0 Then collArray(2).Add i
    Next i
        
    GetUnpairedLists = collArray

End Function


