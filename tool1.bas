Attribute VB_Name = "tool1"
' [VK 16July2023] tool1_ConvBillingGridToCAorIB was developed by
' Vadim Krifuks in close collaboration with Man Ming Tse, Hui Zeng, and other HT2 Team members
' a few functions that open/close file were copied from OCTA Internal Budget template file

Option Explicit

Sub tool1_ConvBillingGridToCAorIB()
' this subroutine opens userform
' userform subrotines are in form_choices

    form_choices.Show

End Sub

Sub FillBlank(curSheet As Worksheet, firstRow As Integer, firstCol As Integer, lastRow As Integer, lastCol As Integer)
Attribute FillBlank.VB_ProcData.VB_Invoke_Func = " \n14"
'this subroutine highlights cells that are empty within a range
       
    Dim cel As Range
    Dim selectedRange As Range
    
    curSheet.Activate
    
    Set selectedRange = curSheet.Range(Cells(firstRow, firstCol), Cells(lastRow, lastCol))

    For Each cel In selectedRange.Cells
        If cel.Value = "" And cel.Interior.ColorIndex = xlNone Then
            cel.ClearContents
            cel.Interior.color = RGB(217, 217, 217)
        End If
    Next cel
   
End Sub

Sub FillCellsOfInterest(curSheet As Worksheet, caDesignCol As Integer, firstRow As Integer, lastRow As Integer, _
                        firstCol As Integer, lastCol As Integer, ByRef designationsArray() As String)
'this surbroutine highlights cells within a range that match parameters passed in the array
       
    Dim cel As Range
    Dim selectedRange As Range
    Dim celValue As String
    
    curSheet.Activate
    
    Set selectedRange = Union(Range(Cells(firstRow, caDesignCol), Cells(lastRow, caDesignCol)), _
                              Range(Cells(firstRow, firstCol), Cells(lastRow, lastCol)))

    For Each cel In selectedRange.Cells
        
        celValue = cel.Value
        
        If celValue <> "" Then
            If celValue Like "*" & designationsArray(0) & "*" Or _
                celValue Like "*" & designationsArray(1) & "*" Or _
                (celValue Like "*" & designationsArray(2) & "*" And Not celValue Like "*" & designationsArray(2) & "(*") Or _
                celValue Like "*" & designationsArray(3) & "*" Or _
                celValue Like "*" & designationsArray(4) & "*" Or _
                celValue Like "*" & designationsArray(5) & "*" Or _
                celValue Like "*" & designationsArray(6) & "*" Or _
                celValue Like "*" & designationsArray(7) & "*" Or _
                celValue Like "*" & designationsArray(8) & "*" Or _
                celValue Like "*" & designationsArray(9) & "*" Or _
                celValue Like "*" & designationsArray(10) & "*" Then
                    cel.Interior.color = RGB(241, 156, 187)
            End If
        End If
    Next cel
    
End Sub

'this subroutine is copied from Internal Budget Template macro
'the only modification is at the end of the subrotine where cells are filled with color
Private Sub schedule_transfer(source_range As Range, target_range As Range)

    Dim myArray As Variant                                          'declare Array
    Dim UB_row_myArray As Integer                                   'number of rows (upper bound) in myArray
    Dim UB_column_myArray As Integer                                'number of columns (upper bound) in myArray
    Dim myArray_freq() As Variant                                   'Array to store frequency
    Dim myArray_color() As Variant                                  'Array to store cell background color
    Dim cValue As String
    Dim cValueLength As String
    Dim m, n As Integer
    
    Dim blank_color As Long
    Dim not_important_cell As Integer
    Dim not_important_cell_font_color As Long
    Dim important_cell As Integer
    'Dim important_cell_font_color As Long
    
    'Dim forms_color, central_lab_color As Long
    'forms_color = RGB(0, 255, 255)
    'central_lab_color = RGB(255, 153, 0)
    
    blank_color = RGB(217, 217, 217)
    not_important_cell = -2
    important_cell = -1
    'important_cell_font_color = RGB(0, 191, 255)
    not_important_cell_font_color = RGB(165, 165, 165)
        
    myArray = source_range 'copy values from Range to myArray
        
    UB_row_myArray = UBound(myArray, 1)                             '1 indicates the first dimension of myArray
    UB_column_myArray = UBound(myArray, 2)                          '2 indicates the second dimension of myArray
        
    'resize Dynamic Arrays
    ReDim myArray_freq(1 To UB_row_myArray, 1 To UB_column_myArray)
    ReDim myArray_color(1 To UB_row_myArray, 1 To UB_column_myArray)

    For m = 1 To UB_row_myArray                                     'loop through rows
        For n = 1 To UB_column_myArray                              'loop through columns

            cValue = myArray(m, n)
            cValueLength = Len(cValue)

            'case1: cell is empty
            'color tells the user that the cells was processed
            If cValueLength = 0 Then
               myArray_color(m, n) = blank_color

            'case2a: "R"
            ElseIf cValueLength = 1 And cValue Like "R" Then
                myArray_freq(m, n) = 1
                
            'case2b: "number*R"
            ElseIf cValueLength > 1 And cValue Like "#*R" Then
                myArray_freq(m, n) = CInt(Left(cValue, cValueLength - 1))

            'case3a: "R(F)"
            ElseIf cValueLength = 4 And cValue Like "R(F)" Then
                myArray_freq(m, n) = 1
                'myArray_color(m, n) = forms_color

           'case3b: "number*R(F)"
            ElseIf cValueLength > 4 And cValue Like "#*R(F)" Then
                myArray_freq(m, n) = CInt(Left(cValue, cValueLength - 4))
                'myArray_color(m, n) = forms_color

            'case4a: "R(CL)"
            ElseIf cValueLength = 5 And cValue Like "R(CL)" Then
                myArray_freq(m, n) = 1
                'myArray_color(m, n) = central_lab_color

            'case4b: "number*R(CL)"
            ElseIf cValueLength > 5 And cValue Like "#*R(CL)" Then
                myArray_freq(m, n) = CInt(Left(cValue, cValueLength - 5))
                'myArray_color(m, n) = central_lab_color
            
            'case 5: all that end in "S1", "S0", and "N(NR)" - [VK 4Aug2023]
            ElseIf cValue Like "*S1" Or cValue Like "*S0" Or cValue Like "*N(NR)" Then
                myArray_freq(m, n) = cValue
                myArray_color(m, n) = important_cell

            'case5a: "S1"
            'ElseIf cValueLength = 2 And cValue Like "S1" Then
            '    myArray_freq(m, n) = "Routine Care"

            'case5b: "number*S1"
            'ElseIf cValueLength > 2 And cValue Like "#*S1" Then
            '    myArray_freq(m, n) = "Routine Care (x" & Left(cValue, cValueLength - 2) & ")"
                                
            'case6a: "N(NA)"
            'ElseIf cValueLength = 5 And cValue Like "N(NA)" Then
            '    myArray_freq(m, n) = "Bundled Service"
            
            'case6b: "number*N(NA)"
            'ElseIf cValueLength > 5 And cValue Like "#*N(NA)" Then
            '    myArray_freq(m, n) = "Bundled Service (x" & Left(cValue, cValueLength - 5) & ")"
                
            'case7a: "N(NB)"
            'ElseIf cValueLength = 5 And cValue Like "N(NB)" Then
            '    myArray_freq(m, n) = "Not Billable"
                
            'case7b: "number*N(NB)"
            'ElseIf cValueLength > 5 And cValue Like "#*N(NB)" Then
            '    myArray_freq(m, n) = "Not Billable (x" & Left(cValue, cValueLength - 5) & ")"
                
            'case8a: "N(NR)"
            'ElseIf cValueLength = 5 And cValue Like "N(NR)" Then
            '    myArray_freq(m, n) = "Not Part of the Research Study"
                
            'case8b: "number*N(NR)"
            'ElseIf cValueLength > 5 And cValue Like "#*N(NR)" Then
            '    myArray_freq(m, n) = "Not Part of the Research Study (x" & Left(cValue, cValueLength - 5) & ")"
                
            'case9a: "N(NU)"
            'ElseIf cValueLength = 5 And cValue Like "N(NU)" Then
            '    myArray_freq(m, n) = "Not used/performed at this visit"
                                
            'case9b: "number*N(NU)"
            'ElseIf cValueLength > 5 And cValue Like "#*N(NU)" Then
            '    myArray_freq(m, n) = "Not used/performed at this visit (x" & Left(cValue, cValueLength - 5) & ")"
                                
            'case10a: "N(WO)"
            'ElseIf cValueLength = 5 And cValue Like "N(WO)" Then
            '    myArray_freq(m, n) = "Write Off"

            'case10b: "N(WO)"
            'ElseIf cValueLength > 5 And cValue Like "#*N(WO)" Then
            '    myArray_freq(m, n) = "Write Off (x" & Left(cValue, cValueLength - 5) & ")"
                
            'case11: anything else
            Else
                myArray_freq(m, n) = cValue
                myArray_color(m, n) = not_important_cell
            
            End If

        Next n
    Next m


    With target_range
    
        'copy values from myArray_freq to Range
        .Resize(UB_row_myArray, UB_column_myArray) = myArray_freq
                
        'loop through Array to add colors
        
        For m = 1 To UB_row_myArray                                                 'rows
            For n = 1 To UB_column_myArray                                          'columns
                If IsEmpty(myArray_color(m, n)) Then                                'modifications for for blank cells
                    .Offset(m - 1, n - 1).Interior.color = xlNone
                ElseIf myArray_color(m, n) = not_important_cell Then                'modifications for cells that are not important
                    .Offset(m - 1, n - 1).Font.color = not_important_cell_font_color
                ElseIf myArray_color(m, n) = important_cell Then                    'modifications for important cells
                    .Offset(m - 1, n - 1).Font.Bold = True
                    '.Offset(m - 1, n - 1).Font.Color = important_cell_font_color
                Else                                                                'fill color for all other cells
                    .Offset(m - 1, n - 1).Interior.color = myArray_color(m, n)
                
                End If
            Next n
        Next m
        
    End With


End Sub
                

Sub ConvertInternalBudgetGridT1toT2(curSheet As Worksheet, startRow As Integer, startCol As Integer, endRow As Integer, endCol As Integer, divChar As String)
'this subroutine converts InternalBudget Grid Type1 to Type2

    Dim myArray As Variant                                          'declare Array
    Dim UB_row_myArray As Integer                                   'number of rows (upper bound) in myArray
    Dim UB_column_myArray As Integer                                'number of columns (upper bound) in myArray
    
    Dim cValue As Variant
    Dim cValueLength As Integer
    Dim m, n As Integer
    
    Dim curUnitRateCell As String
    
    With curSheet
        myArray = .Range(.Cells(startRow, startCol), .Cells(endRow, endCol)) 'copy values from Range to myArray
    End With
    
    UB_row_myArray = UBound(myArray, 1)                             '1 indicates the first dimension of myArray
    UB_column_myArray = UBound(myArray, 2)                          '2 indicates the second dimension of myArray
        
    For m = 1 To UB_row_myArray                                     'loop through rows
        For n = 1 To UB_column_myArray                              'loop through columns

            cValue = myArray(m, n)
            cValueLength = Len(cValue)

        If cValueLength = 0 Then
            GoTo nextCell
        End If
                
        If IsNumeric(cValue) Then
            With curSheet
                With .Cells(startRow, startCol)
                    .Offset(m - 1, n - 1).formula = "=IF(ISERROR(VALUE(" & curUnitRateCell & ")*" & cValue & "),CONCAT(" & curUnitRateCell & "," & Chr(34) & " (x" & Chr(34) & "," & cValue & "," & Chr(34) & ")" & Chr(34) & ")," & curUnitRateCell & "*" & cValue & ")"
                End With
            End With
        ElseIf cValue = divChar Then
            With curSheet
                With .Cells(startRow, startCol)
                    .Offset(m - 1, n - 1).Value = 1
                    curUnitRateCell = .Offset(m - 1, n - 1).Address(RowAbsolute:=False, ColumnAbsolute:=False)
                    If m = 1 Then
                        .Offset(m - 1, n - 1).ColumnWidth = 6
                        With .Offset(m - 2, n - 1)
                            .Value = "Arm Unit Rate"
                            .Font.Bold = True
                            .WrapText = True
                        End With
                    End If
                End With
            End With
            
        End If

nextCell:
        Next n
    Next m
    
    With curSheet
        With .Range(.Cells(startRow, startCol), .Cells(endRow, endCol))
            .NumberFormat = "$#,##0.00"
            .WrapText = True
        End With
    End With

End Sub

Sub VisitNamesReconciliationRow(curSheet As Worksheet, rowLocation As Integer, nameColumn As Integer, commentColumn As Integer)
'this function adds a row that is used to reconcile visit names between versions
'synkronizer uses this row to compare columns

    curSheet.rows(rowLocation).Insert
    
    With curSheet.rows(rowLocation)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
    
    With curSheet.Cells(rowLocation, nameColumn)
    
       .Value = "Manually Assigned Visit Name Unique IDs (see comment on the far right)"
       .HorizontalAlignment = xlLeft
       .Font.color = vbWhite
    
    End With
    
    With curSheet.Cells(rowLocation, commentColumn)
        .Value = "***VERY IMPORTANT*** this row, Visit Name Unique IDs, is used by Synkronizer to correctly compare columns. If Synkronizer instead of comparing columns " _
                    & "shows them as deleted on the source workbook (the Billing Grid with a lower calendar version) and added on the target workbook " _
                    & "(the Billing with a higher calendar version), please add Unique Column IDs to the target workbook in the format of ID_1, " _
                    & "ID_2, ID_3, etc. and then transfer them over to the appropriate columns on the source workbook."
        .Font.Bold = False
        .Font.color = vbRed
        .WrapText = True
        .RowHeight = 60
        .ColumnWidth = 120
        .VerticalAlignment = xlVAlignTop
        .HorizontalAlignment = xlLeft
        
    End With
    
End Sub

Sub ZoomAndFreezePanes(curSheet As Worksheet, r As Integer, c As Integer)
    curSheet.Activate
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
    ActiveWindow.FreezePanes = False
    curSheet.Cells(r, c).Select
    ActiveWindow.Zoom = 70
    ActiveWindow.FreezePanes = True

End Sub

Sub SaveWB(destinationSheet As Worksheet, sourceSheet As Worksheet, destinationRangeFileName As Range, destinationRangeProtocolInfo As Range)
'this subroutine

    Dim ProtocolNo As String
    Dim PI As String
    Dim CalendarVersion As String
    Dim Sponsor As String
    Dim ShortTitle As String
    Dim Today As String
    Dim fileName As String
    Dim fileNameBG As String
    Dim PandCANumbers As String
    Dim TargetAccrual As String
    Dim AccruedToDate As String
    Dim Status As String
        
    'set billing grid specific variables
    With sourceSheet
        ProtocolNo = Right(.Range("B1").Value, Len(.Range("B1").Value) - 14)
        PI = Right(.Range("B2").Value, Len(.Range("B2").Value) - 4)
        PI = Left(PI, InStr(PI, ",") - 1)
        CalendarVersion = Right(.Range("A4").Value, Len(.Range("A4").Value) - 18)
        Sponsor = Right(.Range("C1").Value, Len(.Range("C1").Value) - 9)
        
        TargetAccrual = .Range("A2").Value
        AccruedToDate = .Range("C2").Value
        Status = .Range("C3").Value
        
        If Len(.Range("A3").Value) = 12 Then
            ShortTitle = ""
        Else
            ShortTitle = Right(.Range("A3").Value, Len(.Range("A3").Value) - 13)
        End If

        
        If Len(.Range("A1").Value) < 15 Then
            PandCANumbers = ""
        Else
            PandCANumbers = Right(.Range("A1").Value, Len(.Range("A1").Value) - 14) + " | "
        End If
        
        
    End With
        
    
    'Defined string with today's date to be saved into filename
    Today = Format(Date, "ddmmmyy")
       
    'write in study/calendar identifiers
    With destinationRangeProtocolInfo
        .Value = ProtocolNo + " | " + PI + " | " + Sponsor + " | " + ShortTitle
        .RowHeight = 45
        .Offset(1, 0).Value = PandCANumbers + TargetAccrual + " | " + AccruedToDate + " | " + Status
        .Offset(2, 0).Value = "Calendar Version: " + CalendarVersion
        With .Resize(3, 1)
            .Font.Bold = True
            .WrapText = True
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlVAlignTop
        End With
    End With
    
    'shorten CalendarVersion
    'cut to "budget?version?"
    If InStr(1, CalendarVersion, "budget version:", vbTextCompare) > 0 Then
        CalendarVersion = Left(CalendarVersion, Application.WorksheetFunction.Search("budget?version?", CalendarVersion) - 2)
    End If
    'cut after second "v"
    If InStr(2, CalendarVersion, "v", vbTextCompare) > 0 Then
        CalendarVersion = Right(CalendarVersion, Len(CalendarVersion) - Application.WorksheetFunction.Search("v", CalendarVersion, 2))
    End If
    
    'shorten Sponsor name
    If Len(Sponsor) > 15 Then
        If InStr(1, Sponsor, " ", vbTextCompare) > 0 Then
            Sponsor = Left(Sponsor, Application.WorksheetFunction.Search(" ", Sponsor) - 1)
        Else
            Sponsor = Left(Sponsor, 15)
        End If
    End If
    
    'shorten Short Title for filename
    If Len(ShortTitle) > 0 Then
        ShortTitle = Left(ShortTitle, 20)
    End If
    
'    'write in a proposed filename
'    With destinationRangeFileName
'        .Value = "Proposed Filename: " & fileName
'        .Characters(19).Font.Bold = True
'        .Offset(1, 0).Value = "Note to Analyst: review the proposed filename, edit/remove words in parenthesis as needed, and use it for the filename."
'        .Offset(1, 0).Font.Italic = True
'        .Resize(2, 1).Font.Color = vbRed
'        .Resize(2, 1).WrapText = False
'        .Select
'    End With
    
    'make all font size on the sheet uniform
    destinationSheet.Cells.Font.Size = 11
    
    If InStr(1, destinationSheet.Name, "internal budget grid", vbTextCompare) > 0 Then
        fileName = "Internal Budget Grid_" & ProtocolNo & "_" & PI & "_" & ShortTitle & "_" & Sponsor & "_" & "PAv" & CalendarVersion & "_" & Today
    Else
        fileName = "CA_" & ProtocolNo & "_" & PI & "_" & ShortTitle & "_" & Sponsor & "_" & "PAv" & CalendarVersion & "_" & Today
    End If
    
    fileName = CleanString(fileName)
    fileNameBG = ActiveWorkbook.Name
    
    On Error GoTo SkipSave
        'Save the file
        ActiveWorkbook.SaveAs fileName:=fileName & ".xlsx", FileFormat:=xlOpenXMLWorkbook 'updated from FileFormat:=xlWorkbookDefault to allow save on Mac 1/18/21
        Exit Sub
SkipSave:
    MsgBox ("Renaming and saving the original file " & fileNameBG & " to " & fileName & " was skipped. Please rename/save the file manually before closing.")
    
End Sub

Private Function CleanString(str As String)
' [VK 12July2023] the following function is an adopted function from
' https://stackoverflow.com/questions/24356993/removing-special-characters-vba-excel

    Dim regEx As Object
    Dim strPattern As String: strPattern = "[^a-zA-Z0-9\-\.]" 'The regex pattern to find special characters
    Dim strReplace As String: strReplace = "_" 'The replacement for the special characters
    Set regEx = CreateObject("vbscript.regexp") 'Initialize the regex object
    
    ' Configure the regex object
    With regEx
        .Global = True
        .MultiLine = True
        .IgnoreCase = False
        .Pattern = strPattern
    End With
    
    ' Perform the regex replacement
    CleanString = regEx.Replace(str, strReplace)
    
End Function

Sub CopyRenameSheet(sourceSheet As Worksheet, destinationPlace As Worksheet, renamedSheetName As String)
'this subroutine copies a sheet to a workbook and renames it

    sourceSheet.Copy Before:=destinationPlace
    destinationPlace.Parent.Worksheets(sourceSheet.Name).Name = renamedSheetName


End Sub

Sub InsertArmDivider(curSheet As Worksheet, armNumberRow As Integer, firstRow As Integer, curCol As Integer, lastRow As Integer, counter As Integer, divChar As String)
'this subroutine adds a column of values that separate arms

    With curSheet
        
        .Cells(armNumberRow, curCol).ClearFormats
        .Range(.Cells(armNumberRow, curCol), .Cells(armNumberRow + 1, curCol)).Merge
        
        With .Cells(armNumberRow, curCol)
        
            .Value = "Arm_" & counter
            .ColumnWidth = 7
            .VerticalAlignment = xlVAlignCenter
            .HorizontalAlignment = xlHAlignCenter
        
        End With
        
        With .Range(.Cells(firstRow, curCol), .Cells(lastRow, curCol))
            .Value = divChar
            .Font.color = RGB(165, 165, 165)
            .HorizontalAlignment = xlCenter
        End With
    
    End With
    

End Sub

Sub CopySchedule(destinationSheet As Worksheet, sourceSheet As Worksheet, _
                sourceStartRow As Integer, sourceStartCol As Integer, sourceLastRow As Integer, sourceLastCol As Integer, _
                destinationStartRow As Integer, destinationStartCol As Integer)
'this subroutine copies schedule from source sheet to destination sheet
    
    With sourceSheet
        .Range(.Cells(sourceStartRow, sourceStartCol), .Cells(sourceLastRow, sourceLastCol)).Copy destinationSheet.Cells(destinationStartRow, destinationStartCol)
    End With
    
    destinationSheet.Cells(destinationStartRow, destinationStartCol).Value = sourceSheet.Cells(sourceStartRow, sourceStartCol - 1)
    
    'group columns
    With destinationSheet
        .Columns(destinationStartCol).Resize(, sourceLastCol - sourceStartCol + 1).Group
    End With
    
End Sub

Sub AddSynkVisitNames(curSheet As Worksheet, curSheetName As String, synkRow As Integer, visitNameRow As Integer, armFirstVisitColumn As Integer, armLastVisitColumn As Integer, _
                        checkbox_removeFootnotes As Boolean)
                        
    Dim rng As Range
    Dim numberOfVisits As Integer
    Dim i As Integer    ' used to iterate through arrays
    Dim pos As Integer  ' used to position of find chr(10)
    
    numberOfVisits = armLastVisitColumn - armFirstVisitColumn + 1
    
    ' All elements are initialized to by default to an empty string ""
    Dim segmentNamesArray As Variant
    Dim visitNamesArray As Variant
    ReDim synkVisitNamesArray(1 To numberOfVisits) As String

    With curSheet
        Set rng = .Range(.Cells(synkRow, armFirstVisitColumn), .Cells(synkRow, armLastVisitColumn))
    End With
    
    'STEP 1 - prepare segment names
    'copy/paste segment names
    With curSheet
        curSheet.Range(.Cells(visitNameRow - 1, armFirstVisitColumn), .Cells(visitNameRow - 1, armLastVisitColumn)).Copy
    End With
    rng.PasteSpecial xlPasteAllExceptBorders
    Application.CutCopyMode = False 'empties clipboard after copying
    'remove footnotes
    Call Utilities.RemoveFootnotesFromSelectedRange(rng)
    'copy segment names into array
    segmentNamesArray = rng.Value
    
    For i = 1 To numberOfVisits
        
        'if there is no segment name, copy previous
        If (segmentNamesArray(1, i) = "" And i > 1) Then
            segmentNamesArray(1, i) = segmentNamesArray(1, i - 1)
        
        'if there is a segment name, cut it before the line feed
        Else
            pos = InStr(segmentNamesArray(1, i), Chr(10))
            If pos > 0 Then
                segmentNamesArray(1, i) = Left(segmentNamesArray(1, i), pos - 1)
            End If
        End If
    Next i
            
        
    'STEP 2 - prepare visit names
    'copy/paste visit names
    With curSheet
        .Range(.Cells(visitNameRow, armFirstVisitColumn), .Cells(visitNameRow, armLastVisitColumn)).Copy
    End With
    rng.PasteSpecial xlPasteAllExceptBorders
    Application.CutCopyMode = False 'empties clipboard after copying
    'remove footnotes
    Call Utilities.RemoveFootnotesFromSelectedRange(rng)
    'copy segment names into array
    visitNamesArray = rng.Value
            
    'STEP 3 - prepare synk visit names
    For i = 1 To numberOfVisits
        synkVisitNamesArray(i) = visitNamesArray(1, i) & " " & segmentNamesArray(1, i)
    
    Next i
        
        
    'STEP 4 - write to rng
    
    rng.Value = synkVisitNamesArray
   
    'STEP 5 - add borders for each arm
    With rng
        .Interior.color = RGB(198, 224, 180)
        
        With .Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
            
        With .Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
                    
        With .Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
            
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
                
        With .Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
    End With
End Sub

Sub AddSynkProcedureNames(curSheet As Worksheet, synkColumn As Integer, firstProcedureRow As Integer, procedureColumn As Integer, numberOfProcedures As Integer, _
                            checkbox_trimProcedureNames As Boolean, checkbox_removeFootnotes As Boolean)

    Dim lastProcedureRow As Integer
    Dim curRow As Integer
    
    lastProcedureRow = firstProcedureRow + numberOfProcedures - 1
    curRow = firstProcedureRow
    
    With curSheet
        .Range(.Cells(firstProcedureRow, procedureColumn), .Cells(lastProcedureRow, procedureColumn)).Copy
        .Range(.Cells(firstProcedureRow, synkColumn), .Cells(lastProcedureRow, synkColumn)).PasteSpecial xlPasteAllExceptBorders
        Application.CutCopyMode = False 'empties clipboard after copying
        
        'If checkbox_removeFootnotes Then
            Call Utilities.RemoveFootnotesFromSelectedRange(.Range(.Cells(firstProcedureRow, synkColumn), .Cells(lastProcedureRow, synkColumn)))
        'End If
        
        'If checkbox_trimProcedureNames Then
            Call TrimCleanNames(.Range(.Cells(firstProcedureRow, synkColumn), .Cells(lastProcedureRow, synkColumn)))
        'End If
        
        .Range(.Cells(firstProcedureRow, synkColumn), .Cells(lastProcedureRow, synkColumn)).Interior.color = RGB(198, 224, 180)
    End With

End Sub


Sub PrepareForCopy(destinationSheet As Worksheet, sourceSheet As Worksheet, _
                sourceStartRow As Integer, sourceLastRow As Integer, _
                destinationStartRow As Integer)
'this subroutine adds rows to a sheet

    With destinationSheet
        .Range(.Cells(destinationStartRow + 2, 1), .Cells(destinationStartRow + sourceLastRow - sourceStartRow, 1)).EntireRow.Insert xlShiftDown
    End With

    destinationSheet.Cells(destinationStartRow, 1).EntireRow.Copy
    With destinationSheet
        .Range(.Cells(destinationStartRow, 1), .Cells(destinationStartRow + sourceLastRow - sourceStartRow, 1)).EntireRow.PasteSpecial xlFormulas
        Application.CutCopyMode = False 'empties clipboard after copying
    End With

End Sub

Sub CopyCol(destinationSheet As Worksheet, sourceSheet As Worksheet, _
                sourceStartRow As Integer, sourceCol As Integer, sourceLastRow As Integer, _
                destinationStartRow As Integer, destinationCol As Integer)
'this subroutine copies a column from one sheet to another
    
    With sourceSheet
        .Range(.Cells(sourceStartRow, sourceCol), .Cells(sourceLastRow, sourceCol)).Copy
    End With
    
    With destinationSheet
        .Cells(destinationStartRow, destinationCol).PasteSpecial xlPasteAllExceptBorders
        Application.CutCopyMode = False 'empties clipboard after copying
        .Columns(destinationCol).WrapText = True
    End With

End Sub

Sub AggregateCADesignations(curSheet As Worksheet, _
                            startRow As Integer, formulaCol As Integer, numberOfProcedures As Integer, _
                            scheduleStartColumn As Integer, scheduleEndColumn As Integer, divChar As String)
'this subroutine looks at a row and creates a string of uniquie designations from it

    curSheet.Activate
    
    Dim rowCounter As Integer
    Dim rowArray As String
    Dim celDesignation As String
    Dim cel As Range
    Dim counter As Integer
    Dim countDesignations As Integer
    
    
    For rowCounter = 1 To numberOfProcedures
        
        'reset for new row
        rowArray = ""
        countDesignations = 0
        
        For Each cel In curSheet.Range(Cells(startRow - 1 + rowCounter, scheduleStartColumn), _
                                        Cells(startRow - 1 + rowCounter, scheduleEndColumn))
            
            counter = Len(cel.Value)
            If counter = 0 Or cel.Value = divChar Then
                celDesignation = ""
                GoTo ReadyToGoToNextCell
            Else
                
                Do While counter > 0 And cel.Characters(counter, 1).Font.Superscript = True
                    counter = counter - 1
                Loop
                
                celDesignation = Left(cel.Value, counter)
            End If

        
            Do While Len(celDesignation) > 0 And IsNumeric(Left(celDesignation, 1))
                celDesignation = Right(celDesignation, Len(celDesignation) - 1)
            Loop
            
            If celDesignation = "" Then
                GoTo ReadyToGoToNextCell
            End If
                    
        
        'this makes start of rowArray be: rowArray = "
        If rowArray = "" Then
            rowArray = Chr(34)
        End If
        
        'max array for textjoin is 252 elements. This esentially doesn't allow to add more elements which could be a problem.
        'the solution below works in chunks of 200 elements
        
        If countDesignations < 200 Then
            'R", "
            rowArray = rowArray & celDesignation & Chr(34) & Chr(44) & Chr(32) & Chr(34)
            countDesignations = countDesignations + 1
        
        'if count is 200, reset array
        Else
            
            'remove [, "] from tail of rowArray
            If Right(rowArray, 3) = Chr(44) & Chr(32) & Chr(34) Then
                rowArray = Left(rowArray, Len(rowArray) - 3)
            End If
            
            'keep only unique values in rowArray in a form of string
            Cells(startRow - 1 + rowCounter, formulaCol).formula = "=TEXTJOIN(" & Chr(34) & Chr(34) & Chr(34) & Chr(44) & Chr(32) & Chr(34) & Chr(34) & Chr(34) & ", TRUE,sort(UNIQUE({" & rowArray & "}, TRUE, FALSE)))"
            
            'since calculation is turned off, formulas won't produce a value unless forced to calculate
            Application.Calculate
            
            Cells(startRow - 1 + rowCounter, formulaCol).Value = Cells(startRow - 1 + rowCounter, formulaCol).Value
            rowArray = Chr(34) & Cells(startRow - 1 + rowCounter, formulaCol).Value & Chr(34) & Chr(44) & Chr(32) & Chr(34)
            
            'reset count
            countDesignations = 0
        End If
ReadyToGoToNextCell:
        
        Next
        
        If Len(rowArray) > 1 And Right(rowArray, 3) = Chr(44) & Chr(32) & Chr(34) Then
            rowArray = Left(rowArray, Len(rowArray) - 3)
            
        End If
        
        'Chr(34) = "
        'Chr(32) = [space]
        'Chr(44) = ,
        'Chr(123) = {
        'Chr(125) = }
        
        If Len(rowArray) > 1 Then
            Cells(startRow - 1 + rowCounter, formulaCol).formula = "=TEXTJOIN(" & Chr(34) & Chr(44) & Chr(32) & Chr(34) & ", TRUE,sort(UNIQUE({" & rowArray & "}, TRUE, FALSE)))"
            
            'since calculation is turned off, formulas won't produce a value unless forced to calculate
            Application.Calculate
            
            Cells(startRow - 1 + rowCounter, formulaCol).Value = Cells(startRow - 1 + rowCounter, formulaCol).Value
        End If
        
        If Cells(startRow - 1 + rowCounter, formulaCol).Value = "" Then
            Cells(startRow - 1 + rowCounter, formulaCol).Value = "floating item; Official Comments column must include OnCore billing designation"
        End If
    Next

                                                
End Sub

Sub ExtractEventCodes(curSheet As Worksheet, _
                numberOfProcedures As Integer, _
                startRow As Integer, procedureColumn As Integer)
                
    Dim eventCodesColumn As Integer
    eventCodesColumn = procedureColumn + 1
    
    curSheet.Activate
    Cells(startRow, eventCodesColumn).formula = "=IFERROR(LEFT(IF(FIND(""(88"",TRIM(R[1]C[-1])),RIGHT(TRIM(R[1]C[-1]),9),""""),8),"""")"
    Cells(startRow, eventCodesColumn).AutoFill destination:=Range(Cells(startRow, eventCodesColumn), Cells(startRow + numberOfProcedures - 1, eventCodesColumn)), Type:=xlFillDefault
    
    'since calculation is turned off, formulas won't produce a value unless forced to calculate
    Application.Calculate
    
    Range(Cells(startRow, eventCodesColumn), Cells(startRow + numberOfProcedures - 1, eventCodesColumn)).Copy
    Cells(startRow, eventCodesColumn).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False 'empties clipboard after copying

End Sub

Sub AddOriginalRowOrderIDs(curSheet As Worksheet, _
                numberOfProcedures As Integer, _
                startRow As Integer, procedureColumn As Integer, IdColumn As Integer)
    
    Dim parentCounter As Integer
    Dim currentID As Integer
    Dim procedureName As String
    Dim currentProcedureCount As Integer
    
    curSheet.Activate
    parentCounter = 100

    currentID = parentCounter * 10
    
    For currentProcedureCount = 0 To numberOfProcedures - 1
            
        procedureName = Cells(startRow + currentProcedureCount, procedureColumn).Value
        
        If (Left(Trim(procedureName), 1) = "-") Then
            currentID = currentID + 1
        Else
            currentID = parentCounter * 10
            parentCounter = parentCounter + 1
        End If
        
        Cells(startRow + currentProcedureCount, IdColumn).Value = currentID
    
    Next currentProcedureCount
                
End Sub

Sub AddSynkronizerRowOrderIDs(curSheet As Worksheet, _
                numberOfProcedures As Integer, _
                startRow As Integer, procedureColumn As Integer, originalRowOrderIdColumn, synkronizerRowOrderIdColumn As Integer, sortBackFlag As Boolean)
                
    Dim i As Integer
    
    curSheet.Activate
    
    'sort
    Range(Cells(startRow, 1), Cells(startRow + numberOfProcedures - 1, 1)).EntireRow.sort key1:=Cells(startRow, procedureColumn), Order1:=xlAscending, Header:=xlNo
                 
    'add ids
    For i = 1 To numberOfProcedures
        Cells(startRow + i - 1, synkronizerRowOrderIdColumn).Value = i
    Next
    
    If sortBackFlag Then
        'sort back to the original order
        Range(Cells(startRow, 1), Cells(startRow + numberOfProcedures - 1, 1)).EntireRow.sort key1:=Cells(startRow, originalRowOrderIdColumn), Order1:=xlAscending, Header:=xlNo
    End If
                
End Sub


Sub RemoveChildRows(curSheet As Worksheet, lastRow As Integer, firstRow As Integer, procedureColumn As Integer, eventCodeColumn As Integer)

    Dim i As Integer
    Dim procedure As String
    
    curSheet.Activate
    
    'Deleting Child Rows
    For i = lastRow To firstRow Step -1
        procedure = Trim(Cells(i, procedureColumn).Value)
        If Left(procedure, 1) = "-" Then
            If Len(Cells(i, eventCodeColumn)) = 0 Then
                Cells(i, procedureColumn).EntireRow.Delete
            Else
                Cells(i - 1, eventCodeColumn).Value = Cells(i - 1, eventCodeColumn).Value & ", " & Cells(i, eventCodeColumn)
                Cells(i, procedureColumn).EntireRow.Delete
            End If
        End If
    Next

End Sub

Sub FormatArm(curSheet As Worksheet, firstProcedureRow As Integer, scheduleStartColumn As Integer, lastRow As Integer, lastCol As Integer)

    With curSheet
    
    .Cells.WrapText = False
            
        With .Cells(1, 1)
            .WrapText = True
            .RowHeight = 45
            .ColumnWidth = 60
        End With
        
        With .rows(1)
            .VerticalAlignment = xlVAlignTop
            .Font.Bold = True
        End With
        
        With .rows("2:4")
            .HorizontalAlignment = xlCenter
            .WrapText = True
            .VerticalAlignment = xlVAlignTop
            .Font.Bold = True
        End With
        
        With .Columns(lastCol)
            .ColumnWidth = 100
            .WrapText = True
        End With
               
        With .Range(.Cells(1, 1), .Cells(lastRow, lastCol))
            With .Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            
            With .Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
                    
            With .Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            
            With .Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
                
            With .Borders(xlInsideVertical)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            
            With .Borders(xlInsideHorizontal)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
        
        End With
    
    End With
    Call FillBlank(curSheet, firstProcedureRow, scheduleStartColumn, lastRow, lastCol - 1)
    Call ZoomAndFreezePanes(curSheet, firstProcedureRow, scheduleStartColumn)

End Sub

Sub CheckWorkbookValidity(wb As Workbook, renamedTemplateSheetName As String)
    
    If IsSheetFound(renamedTemplateSheetName) Then
        wb.Close
        MsgBox "You asked Billing Grid Converter to add a sheet named """ & renamedTemplateSheetName _
                & """ to the file you selected. A sheet with the same name already exists in the file. Billing Grid Coverter aborted your request. " _
                & "Please try again with a different file.", vbExclamation, "Error - Try Another File "
                
        'End - Terminates execution immediately. Never required by itself but may be placed anywhere in a procedure to end code execution
        End
    End If
    
End Sub

Sub OverwriteBillingDesignations(source As Worksheet, destination As Worksheet, startRow As Integer, startColumn As Integer, numberOfRows As Integer)
'this subroutine is used to copy/paste billing grid designation legend from one place to another
    
    destination.Columns("A:B").EntireColumn.Delete
    
    With source
        .Range(.Cells(startRow, startColumn), .Cells(startRow + numberOfRows - 1, startColumn)).Copy destination.Cells(1, 1)
    End With
    
    destination.Cells(1, 1).ColumnWidth = 115
    
End Sub

Sub FormatFootnoteLegend(destination As Worksheet)
'this subroutine formats footnote legend sheet

    With destination
        .Cells(1, 2).ColumnWidth = 230
        .Columns(2).WrapText = True
        
    End With
        

End Sub
 
Sub BillingGridToCAFileConverter(ByRef designationsArray() As String, _
                                Optional checkbox_removeFootnotes As Boolean = False, _
                                Optional synkronizerSortFlag As Boolean = False, _
                                Optional convertToInternalBudgetScheduleType1 As Boolean = False, _
                                Optional convertToInternalBudgetScheduleType2 As Boolean = False, _
                                Optional checkbox_trimProcedureNames As Boolean = False, _
                                Optional checkbox_sort As Boolean = False)
'this is the main subroutine

    Dim lastRow As Integer
    Dim lastCol As Integer
    Dim curSheetName As String
    Dim curSheet As Worksheet
    Dim wb As Workbook
    Dim i As Integer
    Dim j As Boolean
    Dim k As Integer
    Dim armCounter As Integer
    Dim divChar As String
    
    Dim templateSheet As Worksheet
    Dim originalTemplateSheetName As String
    Dim renamedTemplateSheetName As String
    Dim footnoteLegendSheetName As String
    
    Dim billingGridOriginalFirstProcedureRow As Integer                    'row where procedures start on the original Billing Grid
    Dim billingGridOriginalFirstProcedureColumn As Integer                 'column where procedures start on the original Billing Grid
    Dim billingGridOriginalScheduleStartColumn As Integer
    Dim billingGridOriginalNumberOfProcedureRows As Integer
    Dim curArmStartColumn As Integer
    Dim synkRow As Integer                  'row for visit names that are used by Synkronizer to compare workbooks
    Dim synkColumn As Integer               'column for procedure names that are used by Synkronizer to compare workbooks
    Dim firstProcedureRow As Integer
    Dim visitNameRow As Integer
    Dim scheduleStartRow As Integer
    Dim scheduleStartColumn As Integer
    Dim procedureColumn As Integer
    Dim oncoreCommentColumn As Integer
    Dim eventCodeColumn As Integer
    Dim aggregatedDesignationsColumn As Integer
    Dim originalRowOrderIdColumn As Integer
    Dim synkronizerRowOrderIdColumn As Integer
    Dim billingDesignationLegendRow As Integer
    Dim armFirstVisitColumn As Integer
    Dim armLastVisitColumn As Integer
    
    'original billing grid variables
    billingGridOriginalFirstProcedureRow = 5                    'row where procedures start on the original Billing Grid
    billingGridOriginalFirstProcedureColumn = 1
    billingGridOriginalScheduleStartColumn = billingGridOriginalFirstProcedureColumn + 1
    
    'new sheet variables
    synkRow = 1
    synkColumn = 1
    
    'fixed rows on the new sheet
    scheduleStartRow = synkRow + 1
    visitNameRow = scheduleStartRow + 3
    firstProcedureRow = scheduleStartRow + 4
    billingDesignationLegendRow = scheduleStartRow + 9
    
    'fixed columns on the new sheet
    procedureColumn = synkColumn + 1
    eventCodeColumn = procedureColumn + 1
    aggregatedDesignationsColumn = procedureColumn + 18
    oncoreCommentColumn = procedureColumn + 19
    scheduleStartColumn = procedureColumn + 24
    originalRowOrderIdColumn = scheduleStartColumn - 3
    synkronizerRowOrderIdColumn = originalRowOrderIdColumn + 1

    'moving columns on the new sheet
    curArmStartColumn = scheduleStartColumn
    armLastVisitColumn = curArmStartColumn - 2
    
    originalTemplateSheetName = "CA_generated on_Template"
    
    If convertToInternalBudgetScheduleType1 Or convertToInternalBudgetScheduleType2 Then
        renamedTemplateSheetName = "Internal Budget Grid v" & Format(Date, "ddmmmyy")
    Else
        renamedTemplateSheetName = "CA_generated on " & Format(Date, "ddmmmyy")
    End If
    
    footnoteLegendSheetName = "Footnote Legend"
    
    Set templateSheet = ActiveWorkbook.Worksheets(originalTemplateSheetName)
    
    j = False   'flag to run a few operations ONCE:
                                                '1)add rows to a template sheet,
                                                '2)copy procedure names,
                                                '3)copy comments,
                                                '4)extract Event Codes,
                                                '5)add IDs for sorting rows
    
    Call Utilities.OpenNewWorkbook
    
    Set wb = ActiveWorkbook
    
    'quick check to confirm that the selected workbook has NOT been processed through this macro
    Call CheckWorkbookValidity(wb, renamedTemplateSheetName)
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
        
    Call OverwriteBillingDesignations(templateSheet, wb.Worksheets("Billing Designation Legend"), billingDesignationLegendRow, procedureColumn, 11)
    
    If IsSheetFound(footnoteLegendSheetName) Then
        Call FormatFootnoteLegend(wb.Worksheets(footnoteLegendSheetName))
    End If
    
    Call CopyRenameSheet(templateSheet, wb.Worksheets("Protocol Information"), renamedTemplateSheetName)
        
    armCounter = 0
        
    For i = 1 To Sheets.count
        wb.Sheets(i).Activate
        Set curSheet = ActiveWorkbook.Worksheets(i)
        
        If curSheet.Name Like renamedTemplateSheetName _
                        Or curSheet.Name Like "Protocol Information" _
                        Or curSheet.Name Like "Billing Designation Legend" _
                        Or curSheet.Name Like footnoteLegendSheetName _
                        Or curSheet.Name Like "QCT Checklist" _
                        Or LCase(curSheet.Name) Like LCase("*CA_generated*") _
                        Or LCase(curSheet.Name) Like LCase("*Internal Budget Grid*") Then
            GoTo NextIteration
        End If
        
        armCounter = armCounter + 1
        curSheetName = curSheet.Name

        'original billing grid file
        'lastCol and lastRow of an arm based on 1st row and 1st column
        lastCol = Cells(1, Columns.count).End(xlToLeft).column      'last column of an arm sheet (not new sheet that we added to the Billing Grid)
        lastRow = Cells(rows.count, 1).End(xlUp).row                'last row of an arm sheet (not new sheet that we added to the Billing Grid)
        billingGridOriginalNumberOfProcedureRows = lastRow - billingGridOriginalFirstProcedureRow + 1
        
        If j = False Then
            
            'add rows to the template sheet
            Call PrepareForCopy(wb.Worksheets(renamedTemplateSheetName), curSheet, _
                billingGridOriginalFirstProcedureRow, lastRow, _
                firstProcedureRow)
            
            'copy procedure names
            Call CopyCol(wb.Worksheets(renamedTemplateSheetName), curSheet, _
                billingGridOriginalFirstProcedureRow, billingGridOriginalFirstProcedureColumn, lastRow, _
                firstProcedureRow, procedureColumn)
                        
            'add unique procedure name labels for synkronizer
            'these labels should be trimmed and have no footnotes
            'VK keept the capability to keep the original state in case this is needed in the future
            Call AddSynkProcedureNames(wb.Worksheets(renamedTemplateSheetName), synkColumn, firstProcedureRow, procedureColumn, billingGridOriginalNumberOfProcedureRows, _
                                        checkbox_trimProcedureNames, checkbox_removeFootnotes)
                        
            'copy current OnCore comments
            Call CopyCol(wb.Worksheets(renamedTemplateSheetName), curSheet, _
                billingGridOriginalFirstProcedureRow, lastCol, lastRow, _
                firstProcedureRow, oncoreCommentColumn)
            
            'add rowIDs to keep track of the original order of procedures on the Billing Grid
            'will be used for sorting the procedures back to the original state after Syncronizer highlights differences
            Call AddOriginalRowOrderIDs(wb.Worksheets(renamedTemplateSheetName), _
                billingGridOriginalNumberOfProcedureRows, firstProcedureRow, procedureColumn, originalRowOrderIdColumn)
                
            If synkronizerSortFlag Then
                
                
                If checkbox_removeFootnotes Then
                    'remove footnotes from procedure column
                    Call Utilities.RemoveFootnotesFromSelectedRange(wb.Worksheets(renamedTemplateSheetName).Range(Cells(firstProcedureRow, procedureColumn), Cells(lastRow, procedureColumn)))
                End If
                
                If checkbox_trimProcedureNames Then
                    'trim procedure names
                    Call TrimCleanNames(wb.Worksheets(renamedTemplateSheetName).Range(Cells(firstProcedureRow, procedureColumn), _
                                                                                Cells(firstProcedureRow + billingGridOriginalNumberOfProcedureRows - 1, procedureColumn)))
                End If
                
                'add rowIDs to prepare rows for comparison in Synkronizer
                'will be used for sorting the procedures before running through Synkronizer
                Call AddSynkronizerRowOrderIDs(wb.Worksheets(renamedTemplateSheetName), _
                    billingGridOriginalNumberOfProcedureRows, firstProcedureRow, synkColumn, originalRowOrderIdColumn, synkronizerRowOrderIdColumn, True)
                        
            End If
            
            'extract Event Codes from child rows
            Call ExtractEventCodes(wb.Worksheets(renamedTemplateSheetName), _
                billingGridOriginalNumberOfProcedureRows, _
                firstProcedureRow, procedureColumn)
            
            k = i
            j = True
        End If
                        
        'apply specific formatting to an arm
        Call FormatArm(curSheet, billingGridOriginalFirstProcedureRow, billingGridOriginalScheduleStartColumn, lastRow, lastCol)
        
        'add a divider between arms
        divChar = "|"
        Call InsertArmDivider(wb.Worksheets(renamedTemplateSheetName), synkRow, firstProcedureRow, curArmStartColumn - 1, _
                                firstProcedureRow + billingGridOriginalNumberOfProcedureRows - 1, armCounter, divChar)
         
        armFirstVisitColumn = curArmStartColumn
         
        'copy schedule from arm to template
        Call CopySchedule(wb.Worksheets(renamedTemplateSheetName), curSheet, 1, billingGridOriginalScheduleStartColumn, lastRow, lastCol - 1, scheduleStartRow, curArmStartColumn)
        
        curArmStartColumn = curArmStartColumn + (lastCol - 1 - billingGridOriginalScheduleStartColumn) + 2
        
        armLastVisitColumn = curArmStartColumn - 2
        
        'add unique visit name labels for synkronizer
        'these labels should have no footnotes
        'VK keept the capability to keep the original state in case this is needed in the future
        Call AddSynkVisitNames(wb.Worksheets(renamedTemplateSheetName), curSheetName, synkRow, visitNameRow, armFirstVisitColumn, armLastVisitColumn, checkbox_removeFootnotes)
    
NextIteration:
    Next
    
    'enter formulas to agregate all CA designations from all arms
    Call AggregateCADesignations(wb.Worksheets(renamedTemplateSheetName), firstProcedureRow, aggregatedDesignationsColumn, _
                                billingGridOriginalNumberOfProcedureRows, scheduleStartColumn, armLastVisitColumn, divChar)
    
    
    'remove child rows from the template sheet
    Call RemoveChildRows(wb.Worksheets(renamedTemplateSheetName), firstProcedureRow + billingGridOriginalNumberOfProcedureRows - 1, firstProcedureRow, _
                        procedureColumn, eventCodeColumn)
    
    'lastRow is now a variable of the new sheet
    lastRow = Cells(rows.count, synkColumn).End(xlUp).row
    
    
    If checkbox_sort Then
        'sort rows to prepare for Synkronizer
        Range(Cells(visitNameRow, 1), Cells(lastRow, 1)).EntireRow.sort key1:=Cells(visitNameRow, synkronizerRowOrderIdColumn), Order1:=xlAscending, Header:=xlYes
    End If
    
    
    'this code is executed if the user requested to remove footnotes from the template sheet
    If checkbox_removeFootnotes Then
       
       
        With wb.Worksheets(renamedTemplateSheetName)
            
            'remove footnotes from synk and procedure colum
            Call Utilities.RemoveFootnotesFromSelectedRange(.Range(.Cells(firstProcedureRow, procedureColumn), .Cells(lastRow, procedureColumn)))
            
            'remove footnotes from all visits and visit headers
            Call Utilities.RemoveFootnotesFromSelectedRange(.Range(.Cells(visitNameRow, scheduleStartColumn), .Cells(lastRow, armLastVisitColumn)))
        End With
    End If
    
    Call FillCellsOfInterest(wb.Worksheets(renamedTemplateSheetName), aggregatedDesignationsColumn, firstProcedureRow, lastRow, _
                            scheduleStartColumn, armLastVisitColumn, designationsArray)
                       
    'this code converts CA designations to a format used on the Internal Budget (ie R to 1, 5R(CL) to 5) and allows to schedule directly against the Internal Budget
    If convertToInternalBudgetScheduleType1 = True Then
        With wb.Worksheets(renamedTemplateSheetName)
            '"schedule_transfer" subroutine is a direct copy from the Internal Budget Template macro with one modification where macro fills cells with colors
            'this modification is done to save time
            'first argument - range of the schedule
            'second argument - top left corner of the range where to place the schedule. In this case, we are replacing schedules in place
            Call schedule_transfer(.Range(.Cells(firstProcedureRow, scheduleStartColumn), .Cells(lastRow, armLastVisitColumn)), .Cells(firstProcedureRow, scheduleStartColumn))
            
            .Range(.Cells(firstProcedureRow, scheduleStartColumn), .Cells(lastRow, armLastVisitColumn)).WrapText = True
            
            
            'renamedTemplateSheetName = "Internal Budget Grid v" + Right(renamedTemplateSheetName, 7)
            '.Name = renamedTemplateSheetName
            
        End With
        
        If convertToInternalBudgetScheduleType2 = True Then
        
            Call ConvertInternalBudgetGridT1toT2(wb.Worksheets(renamedTemplateSheetName), _
                                                firstProcedureRow, scheduleStartColumn - 1, lastRow, armLastVisitColumn, divChar)
        End If
        
    End If
    
    'Call VisitNamesReconciliationRow(wb.Worksheets(renamedTemplateSheetName), scheduleStartRow + 3, procedureColumn, curArmStartColumn - 1)
    
    lastRow = Cells(rows.count, procedureColumn).End(xlUp).row
        
    With wb.Worksheets(renamedTemplateSheetName)
    
        Call SaveWB(wb.Worksheets(renamedTemplateSheetName), _
                    wb.Worksheets("Protocol Information"), _
                    wb.Worksheets(renamedTemplateSheetName).Range(.Cells(lastRow + 16, procedureColumn), .Cells(lastRow + 16, procedureColumn)), _
                    wb.Worksheets(renamedTemplateSheetName).Range(.Cells(scheduleStartRow, procedureColumn), .Cells(scheduleStartRow, procedureColumn)))
    
    End With
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
        
End Sub


Sub BillingGridToCAFileConverterSynkronizerPrepared(ByRef designationsArray() As String, checkbox_removeFootnotes As Boolean, _
                                                checkbox_trimProcedureNames As Boolean, checkbox_sort As Boolean)

    Call BillingGridToCAFileConverter(designationsArray, _
                                        checkbox_removeFootnotes:=checkbox_removeFootnotes, _
                                        synkronizerSortFlag:=True, _
                                        checkbox_trimProcedureNames:=checkbox_trimProcedureNames, _
                                        checkbox_sort:=checkbox_sort)

End Sub

Sub BillingGridToInternalBudgetGridFrequencies(ByRef designationsArray() As String, checkbox_removeFootnotes As Boolean, _
                                                checkbox_trimProcedureNames As Boolean, checkbox_sort As Boolean)
        
    Call BillingGridToCAFileConverter(designationsArray, _
                                        checkbox_removeFootnotes:=checkbox_removeFootnotes, _
                                        synkronizerSortFlag:=True, _
                                        convertToInternalBudgetScheduleType1:=True, _
                                        checkbox_trimProcedureNames:=checkbox_trimProcedureNames, _
                                        checkbox_sort:=checkbox_sort)

End Sub

Sub BillingGridToInternalBudgetGridTotals(ByRef designationsArray() As String, checkbox_removeFootnotes As Boolean, _
                                                checkbox_trimProcedureNames As Boolean, checkbox_sort As Boolean)
        
    Call BillingGridToCAFileConverter(designationsArray, _
                                        checkbox_removeFootnotes:=checkbox_removeFootnotes, _
                                        synkronizerSortFlag:=True, _
                                        convertToInternalBudgetScheduleType1:=True, _
                                        convertToInternalBudgetScheduleType2:=True, _
                                        checkbox_trimProcedureNames:=checkbox_trimProcedureNames, _
                                        checkbox_sort:=checkbox_sort)

End Sub


Sub TrimCleanNames(userRange As Range)

    Dim cRange As Range

        'Trim formula removes all spaces before and after the text as well as consecutive spaces in the middle of a string
        'Clean formula deletes any and all of the first 32 non-printing characters in the 7-bit ASCII set (values 0 through 31)
        ' including line break (value 10) and tab (value 9)
        For Each cRange In userRange.Cells
            cRange.Value = Trim(Application.WorksheetFunction.Clean(cRange.Value))

        Next
End Sub

Private Function IsSheetFound(ByVal sheetName As String)
'this function returns true if a sheet with sheetName exists and false otherwise
'the function is case-insesitive

Dim curSheet As Worksheet

sheetName = LCase(sheetName)

For Each curSheet In Worksheets
            
    If LCase(curSheet.Name) = sheetName Then
        IsSheetFound = True
        Exit Function
    End If
    
Next

IsSheetFound = False

End Function

