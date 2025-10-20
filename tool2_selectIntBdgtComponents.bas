Attribute VB_Name = "tool2_selectIntBdgtComponents"
'Author/Developer: Vadim Krifuks
'Collaborators: Man Ming Tse
'Last Updated: 14 October 2025

Option Explicit
Option Private Module


Function SelectIntBdgtComponents(toolSheet As Worksheet, _
                                column_ibComponents As Integer, _
                                row_workbookName As Integer, _
                                row_sheetName As Integer, _
                                row_proceduresRange As Integer, _
                                row_visitNamesRange As Integer) As Boolean

    Dim uf As New frmTool2SelectIntBdgtComponents
    
    Call LoadDefaults(toolSheet, _
                        column_ibComponents, _
                        row_workbookName, _
                        row_sheetName, _
                        row_proceduresRange, _
                        row_visitNamesRange, _
                        uf)
    
    uf.Show


End Function


Sub LoadDefaults(toolSheet As Worksheet, _
                    column_ibComponents As Integer, _
                    row_workbookName As Integer, _
                    row_sheetName As Integer, _
                    row_proceduresRange As Integer, _
                    row_visitNamesRange As Integer, _
                    uf As frmTool2SelectIntBdgtComponents)
    
    Dim wkb As Workbook
    Dim sh As Worksheet

    Dim wkbNameStr As String
    Dim shNameStr As String
    Dim proceduresRngStr As String
    Dim visitsRngStr As String
    
    With toolSheet
        wkbNameStr = .Cells(row_workbookName, column_ibComponents).Value
        shNameStr = .Cells(row_sheetName, column_ibComponents).Value
        proceduresRngStr = .Cells(row_proceduresRange, column_ibComponents).Value
        visitsRngStr = .Cells(row_visitNamesRange, column_ibComponents).Value
    End With
    
    ' check that workbook is open, sheet exists and ranges are valid
    If Utilities.IsWorkbookOpen(wkbNameStr) Then
        Set wkb = Workbooks(wkbNameStr)
        If Utilities.IsSheetFound(wkb, shNameStr) Then
            Set sh = wkb.Worksheets(shNameStr)
            If Utilities.IsRangeStringValid(sh, proceduresRngStr) And _
                Utilities.IsRangeStringValid(sh, visitsRngStr) Then
                
'                uf.cboFileName.Text = wkbNameStr
'                uf.cboSheetName.Text = shNameStr
                uf.tbxProceduresRange.Text = proceduresRngStr
                uf.tbxVisitsRange.Text = visitsRngStr
            End If
        End If
    End If

End Sub
