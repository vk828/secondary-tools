VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTool2ChooseOnCoreArm 
   Caption         =   "Select OnCore Arm"
   ClientHeight    =   2910
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14265
   OleObjectBlob   =   "frmTool2ChooseOnCoreArm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmTool2ChooseOnCoreArm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private wbOncore As Workbook
Private selectedSheetName As String

Private Sub UserForm_Initialize()
'this sub initializes the form

    Dim ws As Worksheet
    Dim wsName As String
    Dim skipSheets As Variant
    Dim sheetName As String
    Dim skipSheetFound As Boolean
    Dim i As Integer
    
    Set wbOncore = ActiveWorkbook
    
    ' List sheet names you want to exclude
    skipSheets = Array("Protocol Information", "Billing Designation Legend", "Footnote Legend", "QCT Checklist", "CA_generated on *", "Internal Budget Grid v*")
    
    For Each ws In wbOncore.Worksheets
        wsName = ws.name
        skipSheetFound = False
        For i = LBound(skipSheets) To UBound(skipSheets)
            If wsName Like skipSheets(i) Then
                skipSheetFound = True
                Exit For
            End If
        Next i
        
        If Not skipSheetFound Then
            Me.cboArms.AddItem wsName
        End If
    Next ws
    
    ' Set default to first item to avoid empty display
    Me.cboArms.ListIndex = 0
    
    'font size in the combobox; original is 8 and it is too small
    Me.cboArms.Font.Size = 11

End Sub

Private Sub cboArms_Click()
'sub that is called every time combobox selection is changed

    Dim targetSheet As String
    targetSheet = cboArms.value
    If targetSheet <> "" Then
        wbOncore.Activate
        Worksheets(targetSheet).Activate
        Cells(1, 1).Select
        'btnSubmit.Caption = cboArms.Value & Chr(10) & "Submit"
    End If
End Sub

Private Sub btnSubmit_Click()
'sub that is called when 'Submit' button is clicked

    selectedSheetName = cboArms.value
    Me.Hide   'hides the form and keeps it in the memory
End Sub

Private Sub btnExit_Click()
'sub that is called when 'Exit' button is clicked

    Me.Hide   'hides the form and keeps it in the memory
    'Unload Me   'closes the form and clears memory
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
'this sub controls what happens when a user clicks 'X' button in the right top corner
'mimic behavior of Private Sub btnExit_Click()

    If CloseMode = vbFormControlMenu Then ' Detect click on the "X" button
        Cancel = True    ' Cancel the default close action
        Me.Hide          ' Hide the form, just like btnExit_Click does
        ' Unload Me     ' Use this instead if you want to completely close and free memory
    End If
End Sub

Public Property Get SelectedSheet() As String
'property to read the selected arm

    SelectedSheet = selectedSheetName
End Property

Sub DetailInstructions(newWord As String)

    Dim oldWord As String
    oldWord = "[REPLACE]"

    Me.lblPrompt.Caption = Replace(Me.lblPrompt.Caption, oldWord, "'" & newWord & "'")
End Sub
