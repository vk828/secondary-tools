VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTool2SelectIntBdgtComponents 
   Caption         =   "Select Internal Budget Components | One Arm at a Time"
   ClientHeight    =   6255
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15120
   OleObjectBlob   =   "frmTool2SelectIntBdgtComponents.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmTool2SelectIntBdgtComponents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bIgnoreEvents As Boolean


Public Property Get IgnoreEvents() As Boolean
    IgnoreEvents = bIgnoreEvents
End Property

Public Property Let IgnoreEvents(value As Boolean)
    bIgnoreEvents = value
End Property

Private Sub btnDone_Click()
    If (Me.cboFileName.value <> "" And Me.cboSheetName.value <> "" And _
        Me.tbxProceduresRange.value <> "" And Me.tbxVisitsRange.value <> "") Then
    
        Call tool2.WriteComponentsToExcel(Me.cboFileName.value, _
                                        Me.cboSheetName.value, _
                                        Me.tbxProceduresRange.value, _
                                        Me.tbxVisitsRange.value)
        Unload Me
    End If
End Sub

Private Sub btnExit_Click()
    tool2.AbortSelectingRanges
    Unload Me
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        ' The X button was pressed
        tool2.AbortSelectingRanges
        Unload Me
    End If
End Sub

Private Sub btnSelectProceduresRng_Click()
    Me.Hide
    Me.tbxProceduresRange.value = GetSelectedRangeAddress(False, True)
    Me.Show
    'returns the view to where the user form is
    'Me.tbxProceduresRange.SetFocus
End Sub

Private Sub btnSelectVisitsRng_Click()
    Me.Hide
    Me.tbxVisitsRange.value = GetSelectedRangeAddress(True, False)
    Me.Show
    'returns the view to where the user form is
    'Me.tbxVisitsRange.SetFocus
End Sub

Private Sub cboFileName_Change()
    If bIgnoreEvents Then Exit Sub ' Skip event processing
    Call tool2_selectIntBdgtComponents.ChangeIntBdgFile(Me.cboFileName.value, Me)
    
End Sub


Sub LoadCboSheetNamesAndPickActive()
'load active workbook sheet names into combobox and select active sheet

    Dim wsh As Worksheet
    Dim wb As Workbook
    
    Set wb = ActiveWorkbook
    
    Me.cboSheetName.Clear
    
    For Each wsh In wb.Worksheets
        Me.cboSheetName.AddItem wsh.name
    Next wsh
    
    'show the active sheet
    Me.cboSheetName.value = wb.ActiveSheet.name

End Sub

Private Sub cboSheetName_Change()
    Workbooks(Me.cboFileName.value).Activate
    Worksheets(Me.cboSheetName.value).Activate
    Me.tbxProceduresRange.value = ""
    Me.tbxVisitsRange.value = ""
    
End Sub

Private Sub tbxProceduresRange_Change()
    CheckIfReady
End Sub

Private Sub tbxVisitsRange_Change()
    CheckIfReady
End Sub

Private Sub UserForm_Initialize()
'at initialization load all open workbooks in cboFileName and set font size to 11

    Dim wb As Workbook
    
    Me.cboFileName.Clear
    
    For Each wb In Application.Workbooks
        'add all workbooks except one that contains this macro
        If Not wb.FullName = ThisWorkbook.FullName Then
            Me.cboFileName.AddItem wb.name
        End If
    Next wb
    
    'adjust font sizes in the comboboxes and textboxes; the original, 8, is too small
    Me.cboFileName.Font.Size = 11
    Me.cboSheetName.Font.Size = 11
    Me.tbxProceduresRange.Font.Size = 11
    Me.tbxVisitsRange.Font.Size = 11

End Sub

Function GetSelectedRangeAddress(intersectWithFirstRow As Boolean, intersectWithFirstColumn As Boolean) As String
    Dim rng As Range
    
    On Error Resume Next
    Set rng = Application.InputBox("Select a range", "Range Selection", Type:=8)
    On Error GoTo 0
    If Not rng Is Nothing Then
        If rng.Worksheet.Parent.name = Me.cboFileName.value And rng.Worksheet.name = Me.cboSheetName.value Then
            If intersectWithFirstRow Then
                Set rng = Application.Intersect(rng, rng.Cells(1, 1).EntireRow)
            ElseIf intersectWithFirstColumn Then
                Set rng = Application.Intersect(rng, rng.Cells(1, 1).EntireColumn)
            End If
            
            GetSelectedRangeAddress = rng.Address(False, False)
        Else
            GetSelectedRangeAddress = ""
        End If
    Else
        GetSelectedRangeAddress = ""
    End If
End Function

Private Sub CheckIfReady()
    If Len(Me.tbxProceduresRange.Text) > 0 And Len(Me.tbxVisitsRange.Text) > 0 Then
        Me.btnDone.Enabled = True
    Else
        Me.btnDone.Enabled = False
    End If
End Sub
