Attribute VB_Name = "tool2_oncore"
'Author/Developer: Vadim Krifuks
'Last Updated: September 2025

Option Explicit
Option Private Module

Function GetOncoreRanges(ibSheetName As String) As Collection
    Dim wbOncore As Workbook
    Dim wsOncoreName As String
    Dim wsOncore As Worksheet
    
    Dim firstProcedureRow As Integer
    Dim lastProcedureRow As Integer
    Dim procedureNamesCol As Integer
    
    Dim segmentNamesRow As Integer
    Dim visitNamesRow As Integer
    Dim uniqueVisitNamesRow As Integer
    Dim firstVisitCol As Integer
    Dim lastVisitCol As Integer
    
    Dim rngCollection As New Collection
        
    'STEP1: open file
    'exit Sub if system didn't open Billing Grid
    If OpenFileWithStatus = -1 Then Exit Function
    
    Set wbOncore = ActiveWorkbook
    
    'STEP2: choose arm
    'Billing Grid worksheet NAME to be used to process this amendment
    wsOncoreName = ChooseArm(ibSheetName)
    
    'exit Sub and close Billing Grid if user clicks 'Exit'
    If wsOncoreName = "" Then
        wbOncore.Close
        Exit Function
    End If
    
    'Billing Grid WORKSHEET to be used for this amendment
    Set wsOncore = wbOncore.Sheets(wsOncoreName)
    
    segmentNamesRow = 3
    visitNamesRow = segmentNamesRow + 1
    firstProcedureRow = visitNamesRow + 1
    
    procedureNamesCol = 1
    firstVisitCol = procedureNamesCol + 1

    With wsOncore
        lastVisitCol = .Cells(visitNamesRow, .Columns.count).End(xlToLeft).column
    End With

    Call OptimizeStart
    
    'STEP3: remove unnecessary rows
    lastProcedureRow = RemoveRows(wsOncore, firstProcedureRow, procedureNamesCol)

    'STEP4: remove footnotes and update visit names
    
    'if sheet was already processed, skip UpdateVisitNames
    If wsOncore.Cells(firstProcedureRow, procedureNamesCol) = "" Then
        uniqueVisitNamesRow = firstProcedureRow
    Else
        uniqueVisitNamesRow = UpdateVisitNames(wsOncore, firstVisitCol, lastVisitCol, segmentNamesRow, visitNamesRow)
    End If
    
    'adjust procedure row numbers as needed
    If uniqueVisitNamesRow <> visitNamesRow Then
        If uniqueVisitNamesRow <= firstProcedureRow Then
            firstProcedureRow = firstProcedureRow + 1
            lastProcedureRow = lastProcedureRow + 1
        End If
    End If
    
    'STEP5: remove footnotes and update procedure names
    Call UpdateProcedureNames(wsOncore, firstProcedureRow, lastProcedureRow, procedureNamesCol)
    
    'STEP6: remove footnotes and update grid
    Call UpdateGrid(wsOncore, firstProcedureRow, lastProcedureRow, firstVisitCol, lastVisitCol)
    
    Call OptimizeEnd

    'STEP7: return three objects
    rngCollection.Add GetRngObj(wsOncore, firstProcedureRow, lastProcedureRow, procedureNamesCol, procedureNamesCol)
    rngCollection.Add GetRngObj(wsOncore, uniqueVisitNamesRow, uniqueVisitNamesRow, firstVisitCol, lastVisitCol)
    rngCollection.Add GetRngObj(wsOncore, firstProcedureRow, lastProcedureRow, firstVisitCol, lastVisitCol)
    Set GetOncoreRanges = rngCollection
    
End Function

Private Function GetRngObj(ws As Worksheet, _
                            firstRow As Integer, lastRow As Integer, _
                            firstCol As Integer, lastCol As Integer) As Range
    With ws
        Set GetRngObj = .Range(.Cells(firstRow, firstCol), .Cells(lastRow, lastCol))
    End With
    
End Function

Private Sub UpdateGrid(ws As Worksheet, _
                        firstRow As Integer, lastRow As Integer, _
                        firstCol As Integer, lastCol As Integer)

    Dim rng As Range
    Dim cell As Range

    With ws
        Set rng = .Range(.Cells(firstRow, firstCol), .Cells(lastRow, lastCol))
    End With
    
    'remove footnotes from grid
    Call Utilities.RemoveFootnotesFromSelectedRange(rng)
    
    Call UpdateCAToIntBdgtGrid(rng, rng)

End Sub

Private Sub UpdateCAToIntBdgtGrid(source_range As Range, target_range As Range)
'this subroutine is a simplified verstion of the subroutine copied from the Internal Budget Template macro

    Dim myArray As Variant                                          'declare Array
    Dim UB_row_myArray As Integer                                   'number of rows (upper bound) in myArray
    Dim UB_column_myArray As Integer                                'number of columns (upper bound) in myArray
    Dim myArray_freq() As Variant                                   'Array to store frequency
    Dim cValue As String
    Dim cValueLength As Integer
    Dim m, n As Integer
                
    myArray = source_range                                          'copy values from Range to myArray
        
    UB_row_myArray = UBound(myArray, 1)                             '1 indicates the first dimension of myArray
    UB_column_myArray = UBound(myArray, 2)                          '2 indicates the second dimension of myArray
        
    'resize Dynamic Array
    ReDim myArray_freq(1 To UB_row_myArray, 1 To UB_column_myArray)
    
    For m = 1 To UB_row_myArray                                     'loop through rows
        For n = 1 To UB_column_myArray                              'loop through columns

            cValue = myArray(m, n)
            cValueLength = Len(cValue)

            'case1: cell is empty
            If cValueLength = 0 Then
                myArray_freq(m, n) = Empty
                
            'case2a: "R"
            ElseIf cValueLength = 1 And cValue Like "R" Then
                myArray_freq(m, n) = 1
                
            'case2b: "number*R"
            ElseIf cValueLength <= 3 And cValue Like "#*R" Then
                myArray_freq(m, n) = CInt(Left(cValue, cValueLength - 1))

            'case3a: "R(F)"
            ElseIf cValueLength = 4 And cValue Like "R(F)" Then
                myArray_freq(m, n) = 1

           'case3b: "number*R(F)"
            ElseIf cValueLength <= 6 And cValue Like "#*R(F)" Then
                myArray_freq(m, n) = CInt(Left(cValue, cValueLength - 4))
                
            'case4a: "R(CL)"
            ElseIf cValueLength = 5 And cValue Like "R(CL)" Then
                myArray_freq(m, n) = 1
                
            'case4b: "number*R(CL)"
            ElseIf cValueLength <= 7 And cValue Like "#*R(CL)" Then
                myArray_freq(m, n) = CInt(Left(cValue, cValueLength - 5))
                
            'case5: anything else
            Else
                myArray_freq(m, n) = cValue
                
            End If

        Next n
    Next m

    'copy values from myArray_freq to Range
    target_range = myArray_freq

End Sub

Private Sub UpdateProcedureNames(ws As Worksheet, firstRow As Integer, lastRow As Integer, col As Integer)

    Dim rng As Range
    Dim cell As Range

    With ws
        Set rng = .Range(.Cells(firstRow, col), .Cells(lastRow, col))
    End With
    
    'remove footnotes from procedures names
    Call Utilities.RemoveFootnotesFromSelectedRange(rng)
    
    For Each cell In rng
        cell.Value = Application.WorksheetFunction.Trim(Application.WorksheetFunction.Clean(cell.Value))
    Next cell

End Sub

Private Sub OptimizeEnd()

    With Application
        .Calculation = xlCalculationAutomatic
        .ScreenUpdating = True
    End With

End Sub


Private Sub OptimizeStart()

    With Application
        .Calculation = xlManual
        .ScreenUpdating = False
    End With

End Sub


Private Function UpdateVisitNames(ws As Worksheet, _
                            firstCol As Integer, _
                            lastCol As Integer, _
                            segmentNamesRow As Integer, _
                            visitNamesRow As Integer) As Integer
'returns unique visit names row

    Dim uniqueNamesRow As Integer

    Dim rng As Range
    Dim cell As Range
    Dim prevSegmentName As String
    Dim curSegmentName As String
    
    Dim segmentCellAddress As String
    Dim visitCellAddress As String
    
    uniqueNamesRow = visitNamesRow + 1
    
    'remove footnotes and unmerge segment names
    With ws
        Set rng = .Range(.Cells(segmentNamesRow, firstCol), .Cells(segmentNamesRow, lastCol))
    End With
    
    Call Utilities.RemoveFootnotesFromSelectedRange(rng)
    rng.UnMerge
    
    'add missing segment names
    prevSegmentName = rng.Cells(1).Value
    For Each cell In rng
        curSegmentName = cell.Value
        If curSegmentName = "" Then
            cell.Value = prevSegmentName
        Else
            prevSegmentName = curSegmentName
        End If
    Next cell
    
    'remove footnotes from visit names
    With ws
        Set rng = .Range(.Cells(visitNamesRow, firstCol), .Cells(visitNamesRow, lastCol))
    End With
    
    Call Utilities.RemoveFootnotesFromSelectedRange(rng)
    
    'insert row for unique names
    ws.Cells(uniqueNamesRow, firstCol).EntireRow.Insert
    
    'add formulas into the unique names row and make them easier to read
    '=trim(clean(CONCAT(AM$36, " ", LEFT(AM$35,SEARCH(CHAR(10),AM$35)-1))))
    
    With ws
        segmentCellAddress = .Cells(segmentNamesRow, firstCol).Address(False, False)
        visitCellAddress = .Cells(visitNamesRow, firstCol).Address(False, False)
        
        .Cells(uniqueNamesRow, firstCol).formula = _
                 "=TRIM(CLEAN(CONCAT(" & visitCellAddress & ", CHAR(32), LEFT(" & segmentCellAddress & ",SEARCH(CHAR(10)," & segmentCellAddress & ")-1))))"
                 
        Range(.Cells(uniqueNamesRow, firstCol), .Cells(uniqueNamesRow, lastCol)).formula = .Cells(uniqueNamesRow, firstCol).formula
        
        With Range(.Cells(uniqueNamesRow, firstCol), .Cells(uniqueNamesRow, lastCol))
            .WrapText = True
            .VerticalAlignment = xlTop
        End With
    End With
    
    'hide rows to make it easier to read for a human
    With ws
        .rows(segmentNamesRow).EntireRow.Hidden = True
        .rows(visitNamesRow).EntireRow.Hidden = True
    End With
    
    UpdateVisitNames = uniqueNamesRow
    
End Function

Private Function RemoveRows(ws As Worksheet, firstRow, col) As Integer
'returns number of last row

    Dim lastRow As Integer
    
    Dim startOfRemovedRows As Variant
    Dim cell As Range
    Dim curProcedureName As String
    Dim rng As Range
    Dim i, j As Integer
    
    startOfRemovedRows = Array("-", "(INV)")
    
    With ws
        lastRow = .Cells(.rows.count, col).End(xlUp).row
        Set rng = .Range(.Cells(firstRow, col), .Cells(lastRow, col))
    End With
        
    For j = rng.Cells.count To 1 Step -1
        curProcedureName = Trim(Application.WorksheetFunction.Clean(rng.Cells(j).Value))
        For i = LBound(startOfRemovedRows) To UBound(startOfRemovedRows)
            If InStr(1, curProcedureName, startOfRemovedRows(i)) = 1 Then
                rng.Cells(j).EntireRow.Delete
                lastRow = lastRow - 1
                Exit For
            End If
        Next i
    Next j
    
    RemoveRows = lastRow
    
End Function

Private Function ChooseArm(name As String) As String
'function opens a user form and let's the user pick a billing grid arm
'it returns the arm name; if user clicks 'Exit' or 'X', the return is "" (empty string)
    
    Dim uf As New frmTool2ChooseArm
    uf.DetailInstructions (name)
    uf.Show
    
    ChooseArm = uf.SelectedSheet
    Unload uf

End Function

Private Function OpenFileWithStatus()
' returns -1 if unsuccessful, 0 if successful

    Dim FileToOpen As Variant
    FileToOpen = Application.GetOpenFilename( _
        title:="Please choose a Billing Grid to open", _
        FileFilter:="Excel Files (*.xls*),*.xls*", _
        FilterIndex:=1, _
        ButtonText:="Open", _
        MultiSelect:=False)
    
    Application.WindowState = xlMaximized
        
    If FileToOpen = False Then
        MsgBox "No file selected.", vbExclamation
        OpenFileWithStatus = -1
    Else
        On Error GoTo ErrorHandler
        Workbooks.Open fileName:=FileToOpen
        On Error GoTo 0 ' Reset error handling
        OpenFileWithStatus = 0
    End If
    
    Exit Function
    
ErrorHandler:
    MsgBox "Error opening file: " & Err.Description, vbCritical
    OpenFileWithStatus = -1
End Function
