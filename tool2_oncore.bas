Attribute VB_Name = "tool2_oncore"
'Author/Developer: Vadim Krifuks
'Collaborators: Man Ming Tse
'Last Updated: 13 October 2026

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
    Dim firstVisitCol As Integer
    Dim lastVisitCol As Integer
    
    Dim rngCollection As New Collection
        
    'STEP1: open file
    'exit Sub if system didn't open Billing Grid
    If OpenFileWithStatus = -1 Then Exit Function
    
    'turn billing grid file autosave off
    Call Utilities.SetWorkbookAutoSave(ActiveWorkbook, False)
    
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
    Call UpdateVisitNames(wsOncore, firstVisitCol, lastVisitCol, segmentNamesRow, visitNamesRow)

    'STEP5: remove footnotes and update procedure names
    Call UpdateProcedureNames(wsOncore, firstProcedureRow, lastProcedureRow, procedureNamesCol)
    
    'STEP6: remove footnotes and update grid
    Call UpdateGrid(wsOncore, firstProcedureRow, lastProcedureRow, firstVisitCol, lastVisitCol)
    
    Call OptimizeEnd

    'STEP7: return three objects
    rngCollection.Add GetRngObj(wsOncore, firstProcedureRow, lastProcedureRow, procedureNamesCol, procedureNamesCol)
    rngCollection.Add GetRngObj(wsOncore, visitNamesRow, visitNamesRow, firstVisitCol, lastVisitCol)
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
        cell.value = Application.WorksheetFunction.Trim(Application.WorksheetFunction.Clean(cell.value))
    Next cell

End Sub

Private Sub UpdateVisitNames(ws As Worksheet, _
                            firstCol As Integer, lastCol As Integer, _
                            segmentNamesRow As Integer, visitNamesRow As Integer)
' removes footnotes from segment and visit names
    Dim rng As Range
    
    With ws
        Set rng = .Range(.Cells(segmentNamesRow, firstCol), .Cells(visitNamesRow, lastCol))
    End With
    
    Call Utilities.RemoveFootnotesFromSelectedRange(rng)
    
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
        lastRow = .Cells(.Rows.count, col).End(xlUp).Row
        Set rng = .Range(.Cells(firstRow, col), .Cells(lastRow, col))
    End With
        
    For j = rng.Cells.count To 1 Step -1
        curProcedureName = Trim(Application.WorksheetFunction.Clean(rng.Cells(j).value))
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
    
    Dim uf As New frmTool2ChooseOnCoreArm
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
