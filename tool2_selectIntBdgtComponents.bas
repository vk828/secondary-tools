Attribute VB_Name = "tool2_selectIntBdgtComponents"
'Author/Developer: Vadim Krifuks
'Collaborators: Man Ming Tse
'Last Updated: 25 October 2025

Option Explicit
Option Private Module

Dim currUf As frmTool2SelectIntBdgtComponents
Dim isDefaultLoad As Boolean

Function SelectIntBdgtComponents(toolSheet As Worksheet, _
                                column_ibComponents As Integer, _
                                row_workbookName As Integer, _
                                row_sheetName As Integer, _
                                row_proceduresRange As Integer, _
                                row_visitNamesRange As Integer) As Boolean
    
    Call LoadDefaults(toolSheet, _
                        column_ibComponents, _
                        row_workbookName, _
                        row_sheetName, _
                        row_proceduresRange, _
                        row_visitNamesRange)
    
End Function

Sub ChangeIntBdgFile(fileName As String, uf As frmTool2SelectIntBdgtComponents)

    Dim wb As Workbook
    'Dim currUf As frmTool2SelectIntBdgtComponents
    
    Set wb = Workbooks(fileName)

    ' Unload previous form if loaded
    If Not uf Is Nothing Then
        Unload uf
    End If
    
    wb.Activate
    
    Set uf = New frmTool2SelectIntBdgtComponents
    Load uf
    Set currUf = uf
    
    uf.IgnoreEvents = True  ' Disable events before updating ComboBox
    uf.cboFileName.value = fileName
    uf.IgnoreEvents = False ' Re-enable events after update

    uf.LoadCboSheetNamesAndPickActive
    
    'show user form after initial load
    'user form is shown from initial load function on initial load
    If Not isDefaultLoad Then Call ShowUserForm(currUf)
    
End Sub

Sub ShowUserForm(uf As frmTool2SelectIntBdgtComponents)
    
    With uf
        .StartUpPosition = 0 ' Manual positioning
        .Left = Application.Left + (Application.Width - .Width) / 2
        .Top = Application.Top + (Application.Height - .Height) / 2
        .Show
    End With

End Sub

Sub LoadDefaults(toolSheet As Worksheet, _
                    column_ibComponents As Integer, _
                    row_workbookName As Integer, _
                    row_sheetName As Integer, _
                    row_proceduresRange As Integer, _
                    row_visitNamesRange As Integer)
    
    Dim wkb As Workbook
    Dim sh As Worksheet
    Dim uf As frmTool2SelectIntBdgtComponents
    Dim isDefaultLoadInvalid As Boolean

    Set uf = New frmTool2SelectIntBdgtComponents
    Load uf
    Set currUf = uf
    
    isDefaultLoad = True
    
    Dim wkbNameStr As String
    Dim shNameStr As String
    Dim proceduresRngStr As String
    Dim visitsRngStr As String
    
    With toolSheet
        wkbNameStr = .Cells(row_workbookName, column_ibComponents).value
        shNameStr = .Cells(row_sheetName, column_ibComponents).value
        proceduresRngStr = .Cells(row_proceduresRange, column_ibComponents).value
        visitsRngStr = .Cells(row_visitNamesRange, column_ibComponents).value
    End With
    
    ' check that workbook is open, sheet exists and ranges are valid
    If Utilities.IsWorkbookOpen(wkbNameStr) Then
        Set wkb = Workbooks(wkbNameStr)
        If Utilities.IsSheetFound(wkb, shNameStr) Then
            Set sh = wkb.Worksheets(shNameStr)
            If Utilities.IsRangeStringValid(sh, proceduresRngStr) And _
                Utilities.IsRangeStringValid(sh, visitsRngStr) Then
                
                currUf.cboFileName.value = wkbNameStr
                currUf.cboSheetName.value = shNameStr
                currUf.tbxProceduresRange.value = proceduresRngStr
                currUf.tbxVisitsRange.value = visitsRngStr
                
            Else
            ' ranges are not valid
                isDefaultLoadInvalid = True
            
            End If
        Else
        ' worksheet doesn't exist
            isDefaultLoadInvalid = True
        
        End If
    Else
    ' workbook is not open
        isDefaultLoadInvalid = True
    
    End If

    'if loading defaults fail, make the first file be selected in the combobox
    If isDefaultLoadInvalid Then
        If currUf.cboFileName.ListCount > 0 Then currUf.cboFileName.ListIndex = 0
    End If

    isDefaultLoad = False
    Call ShowUserForm(currUf)

End Sub
