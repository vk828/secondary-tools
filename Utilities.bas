Attribute VB_Name = "Utilities"
'Author/Developer: Vadim Krifuks
'Last Updated: 26December2024

Option Explicit
Option Private Module

Function SelectRange(ByVal title As String, ByVal prompt As String)
'this function 1) lets user select a range, and 2) makes the sheet of the selected range active
'the second action is necessary because otherwise excel seems to be confused about what
'ActiveSheet it is on

    Dim rng As Range
    
    'on Error Resume Next takes the program to the next line in case the user cancels selecting a range
    On Error Resume Next
    'get a range input from the user for Unit Rates
    Set rng = Application.InputBox( _
                title:=title, _
                prompt:=prompt, _
                Type:=8)

    If Not (rng Is Nothing) Then
        rng.Parent.Activate
    End If

    'MsgBox prompt:="you selected " & rng.Address(external:=True)

    Set SelectRange = rng

End Function

Sub RemoveFootnotesFromSelectedRange(userRange As Range)
'this subroutine looks at each cell in a range and removes superscript

    Dim cRange As Range
    Dim counter As Integer
 
    'Remove Footnotes from provided range
    For Each cRange In userRange.Cells
            
        counter = Len(cRange.value)
        
        'if cell is empty, go to next cell
        If counter = 0 Then
            GoTo NextIteration
        
        'otherwise go from right to left and see if there are superscripts
        'this assumes that all possible superscripts are all on the right side of the string
        Else
            Do While counter > 0 And cRange.Characters(counter, 1).Font.Superscript = True
                counter = counter - 1
            Loop
                    
            cRange.value = Left(cRange.value, counter)
        End If
NextIteration:
    Next
End Sub

Function AreWorkbookAndWorksheetValid(ByVal workbookName As String, ByVal worksheetName As String)
'this function returns true workbook is open and worksheet exists within this workbook
'returns false otherwise and shows a message to the user to give feedback on what needs to be fixed

    If IsWorkbookOpen(workbookName) Then
        If IsSheetFound(Workbooks(workbookName), worksheetName) Then
            AreWorkbookAndWorksheetValid = True
        Else
            AreWorkbookAndWorksheetValid = False
            MsgBox "Can't find " & worksheetName & " within " & workbookName & " workbook. Please update the worksheet and try again."
        End If
    Else
        AreWorkbookAndWorksheetValid = False
        MsgBox workbookName & " workbook is NOT open. Please open it and try again."
    End If
End Function

Function IsWorkbookOpen(workbookName As String) As Boolean
'this function returns true if workbook is open, false otherwise

    Dim wkb As Workbook
    On Error Resume Next
    Set wkb = Workbooks(workbookName)
    On Error GoTo 0
    If wkb Is Nothing Then
        IsWorkbookOpen = False
    Else
        IsWorkbookOpen = True
    End If
End Function

Function IsSheetFound(ByVal wkb As Workbook, ByVal sheetName As String)
'this function returns true if a sheet with sheetName exists and false otherwise
'the function is case-insesitive

    Dim curSheet As Worksheet
    sheetName = LCase(sheetName)
    
    For Each curSheet In wkb.Worksheets
        If LCase(curSheet.name) = sheetName Then
            IsSheetFound = True
            Exit Function
        End If
    Next
    IsSheetFound = False
End Function

Function IsRangeStringValid(ws As Worksheet, rngString As String) As Boolean
'this function returns true if range string is valid and false otherwise

    Dim testRange As Range
    On Error GoTo InvalidRange
    Set testRange = ws.Range(rngString)
    On Error GoTo 0
    IsRangeStringValid = True
    Exit Function

InvalidRange:
    IsRangeStringValid = False
End Function

Function IsSingleSheet(rng As Range) As Boolean
'this function returns true if all areas of a non-contiguous range come from a single
'worksheet; false otherwise
    Dim area As Range
    Dim ws As Worksheet
    Dim firstSheet As Worksheet
    
    If rng.Areas.count = 0 Then
        ' nothing range case: consider this valid (on one sheet)
        IsSingleSheet = True
        Exit Function
    End If
    
    Set firstSheet = rng.Areas(1).Worksheet
    
    For Each area In rng.Areas
        Set ws = area.Worksheet
        If Not ws Is firstSheet Then
            IsSingleSheet = False
            Exit Function
        End If
    Next area
    
    IsSingleSheet = True
End Function

Function AssembleRangeComponentsToRange(columnIndex As Integer, _
                                        rowIndex_wkb As Integer, _
                                        rowIndex_wksh As Integer, _
                                        rowIndex_rng As Integer, _
                                        sourceSheet As Worksheet) As Range
'this function returns a range built from range components listed on a sheet

    With sourceSheet
        Set AssembleRangeComponentsToRange = Workbooks(CStr(.Cells(rowIndex_wkb, columnIndex))). _
                                                Worksheets(CStr(.Cells(rowIndex_wksh, columnIndex))). _
                                                Range(CStr(.Cells(rowIndex_rng, columnIndex)))
    End With
End Function

Function AssembleRangeComponentsToAddressString(columnIndex As Integer, rowIndex_wkb As Integer, _
                                                rowIndex_wksh As Integer, rowIndex_rng As Integer, _
                                                sourceSheet As Worksheet) As String
'this function returns an external absolute row column address string built from range
'components listed on a sheet

        AssembleRangeComponentsToAddressString = AssembleRangeComponentsToRange(columnIndex, _
                                                                                rowIndex_wkb, _
                                                                                rowIndex_wksh, _
                                                                                rowIndex_rng, _
                                                                                sourceSheet). _
                                                                                Address(external:=True)

End Function

Sub WriteSelectedRangeComponentsToCells(row_workbookName As Integer, _
                                        row_sheetName As Integer, _
                                        column_allComponents As Integer, _
                                        row_range As Integer, _
                                        ByVal inputRange As Range, _
                                        toolSheet As Worksheet)
'this subroutine:
'  1) takes a range object and writes its workbook name, worksheet name, and range to
'  cells on a sheet
'  2) it also alarms the user if previously written workbook name and/or worksheet name changed

    Dim savedWorkbookName, savedSheetName, selectedWorkbookName, selectedSheetName As String
    Dim promptPart1, promptPart2, promptPart3 As String
    
    With toolSheet
        savedWorkbookName = .Cells(row_workbookName, column_allComponents)
        savedSheetName = .Cells(row_sheetName, column_allComponents)
    End With
    
    selectedWorkbookName = inputRange.Parent.Parent.name
    selectedSheetName = inputRange.Parent.name
    
    promptPart1 = "Your selection changed" & Chr(10)
    
    promptPart2 = "  workbook" & Chr(10) _
                & "    from: " & savedWorkbookName & Chr(10) _
                & "    to: " & selectedWorkbookName & Chr(10)
    
    promptPart3 = "  sheet" & Chr(10) _
                & "    from: " & savedSheetName & Chr(10) _
                & "    to: " & selectedSheetName & Chr(10)
    
    If LCase(savedWorkbookName) <> LCase(selectedWorkbookName) _
        And LCase(savedSheetName) <> LCase(selectedSheetName) Then
        MsgBox promptPart1 & promptPart2 & promptPart3
        
    ElseIf LCase(savedWorkbookName) <> LCase(selectedWorkbookName) Then
        MsgBox promptPart1 & promptPart2
        
    ElseIf LCase(savedSheetName) <> LCase(selectedSheetName) Then
        MsgBox promptPart1 & promptPart3
    End If
    
    With toolSheet
        .Cells(row_workbookName, column_allComponents) = selectedWorkbookName
        .Cells(row_sheetName, column_allComponents) = selectedSheetName
    
        .Cells(row_range, column_allComponents) = inputRange.Address(False, False)
    End With

End Sub

Function SelectDataRngFromNoncontiguousHeaders(ByVal visitsRng As Range, ByVal proceduresRng As Range) As Range
'this function takes visit and procedure ranges and returns a data range
'it works for noncontiguous ranges

    Dim singleAreaRng As Range
    Dim allColumns As Range, allRows As Range
        
    'build union of all columns from headers
    For Each singleAreaRng In visitsRng.Areas
        If allColumns Is Nothing Then
            Set allColumns = singleAreaRng.EntireColumn
        Else
            Set allColumns = Union(allColumns, singleAreaRng.EntireColumn)
        End If
    Next singleAreaRng

    'build union of all rows from headers
    For Each singleAreaRng In proceduresRng.Areas
        If allRows Is Nothing Then
            Set allRows = singleAreaRng.EntireRow
        Else
            Set allRows = Union(allRows, singleAreaRng.EntireRow)
        End If
    Next singleAreaRng

    Set SelectDataRngFromNoncontiguousHeaders = Application.Intersect(allRows, allColumns)

End Function

Function SetDataRange(row_workbookName As Integer, _
                      row_sheetName As Integer, _
                      row_visitNamesRange As Integer, _
                      row_proceduresRange As Integer, _
                      row_dataRange As Integer, _
                      column_allComponents As Integer, _
                      toolSheet As Worksheet) As Boolean
'this function attempts to set data range based on provided components, and
'writes it to sheet and returns true if data range was set,
'or writes nothing and returns false otherwise

    Dim wkbNameStr, wkshNameStr, visitNamesAddrString, proceduresAddrString As String
    Dim visitNamesRng, proceduresRng, dataRng As Range
    
    With toolSheet
    
        wkbNameStr = .Cells(row_workbookName, column_allComponents).value
        wkshNameStr = .Cells(row_sheetName, column_allComponents).value
        visitNamesAddrString = .Cells(row_visitNamesRange, column_allComponents).value
        proceduresAddrString = .Cells(row_proceduresRange, column_allComponents).value
    
    End With

    SetDataRange = False
    toolSheet.Cells(row_dataRange, column_allComponents).ClearContents
    
    If Utilities.AreWorkbookAndWorksheetValid(wkbNameStr, wkshNameStr) Then
            
        With Workbooks(wkbNameStr).Worksheets(wkshNameStr)
            
            Set visitNamesRng = .Range(visitNamesAddrString)
            Set proceduresRng = .Range(proceduresAddrString)
    
            'Set dataRng = .Range(.Cells(proceduresRng.Cells(1, 1).row, visitNamesRng.Cells(1, 1).column), _
                                 .Cells(proceduresRng.Cells(proceduresRng.rows.count, 1).row, _
                                                visitNamesRng.Cells(1, visitNamesRng.Columns.count).column))
                                                
            Set dataRng = SelectDataRngFromNoncontiguousHeaders(visitNamesRng, proceduresRng)
            
        End With

        toolSheet.Cells(row_dataRange, column_allComponents) = dataRng.Address(False, False)
        
        SetDataRange = True

    End If

End Function

Sub FillFormulas(toolSheet As Worksheet, sourceRow As Integer, sourceColumn As Integer, ByVal endRow As Integer, endColumn As Integer)
'this subrotine copies top left corner cells and pastes it to the entire field

    With toolSheet

        .Cells(sourceRow, sourceColumn).Copy
        
        With .Range(.Cells(sourceRow, sourceColumn), .Cells(endRow, endColumn))
            .PasteSpecial xlPasteFormulasAndNumberFormats
            .WrapText = True
            .Font.ColorIndex = 0
        End With
    
    End With

End Sub

Private Function SelectFileOrFilesMac()
'this function gets a file descriptor for Mac
'Source derived from MS website: https://docs.microsoft.com/en-us/previous-versions/office/developer/office-2010/hh710200(v=office.14)?redirectedfrom=MSDN
    
    Dim MyPath As String
    Dim MyScript As String
    Dim MyFile As String
    Dim Fname As String
    Dim mybook As Workbook
    
    On Error Resume Next
    MyPath = MacScript("return (path to documents folder) as String")
    'Or use MyPath = "Macintosh HD:Users:Ron:Desktop:TestFolder:"
    
    ' In the following statement, change true to false in the line "multiple
    ' selections allowed true" if you do not want to be able to select more
    ' than one file. Additionally, if you want to filter for multiple files, change
    ' {""com.microsoft.Excel.xls""} to
    ' {""com.microsoft.excel.xls"",""public.comma-separated-values-text""}
    ' if you want to filter on xls and csv files, for example.
    MyScript = _
        "set applescript's text item delimiters to "","" " & vbNewLine & _
        "set theFiles to (choose file of type " & _
        " {""org.openxmlformats.spreadsheetml.sheet""} " & _
        "with prompt ""Please select a file or files"" default location alias """ & _
        MyPath & """ multiple selections allowed false) as string" & vbNewLine & _
        "set applescript's text item delimiters to """" " & vbNewLine & _
        "return theFiles"
        
    MyFile = MacScript(MyScript)
    
    SelectFileOrFilesMac = MyFile
   
End Function

Sub OpenNewWorkbook()
'sub to open a file

    Dim strFile As String
    Dim wb As Workbook
    
    'Opening source file from location
    'Logic updated to determine OS is Win or Mac then use appropriate method for select file dialogue
    If Not Application.OperatingSystem Like "*Mac*" Then
        ' Is Windows.
        strFile = Application.GetOpenFilename("Excel-files,*.xlsx", 1, "Select the Billing Grid file", "Open", False)
    Else
        ' Is a Mac and will test if running Excel 2011 or higher.
        If val(Application.Version) > 16 Then
            strFile = SelectFileOrFilesMac
            strFile = Replace(strFile, ":", "/")
            strFile = Replace(strFile, "Macintosh HD", "")
        ElseIf val(Application.Version) > 14 Then
            strFile = SelectFileOrFilesMac
        End If
    End If
   
    On Error GoTo ErrorHandler
    Set wb = Workbooks.Open(strFile)
    
    wb.Activate
    Exit Sub
    
ErrorHandler:
End

End Sub

Sub SwitchActiveSheet(rng As Range)

    'activates workbook
    rng.Worksheet.Parent.Activate
    
    'activates worksheet
    rng.Worksheet.Activate

End Sub
