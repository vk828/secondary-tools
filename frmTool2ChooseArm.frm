VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTool2ChooseArm 
   Caption         =   "UserForm1"
   ClientHeight    =   7170
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14250
   OleObjectBlob   =   "frmTool2ChooseArm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmTool2ChooseArm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private wbOncore As Workbook
Private selectedSheetName As String

Private Sub UserForm_Initialize()
    Dim ws As Worksheet
    Dim wsName As String
    Dim skipSheets As Variant
    Dim sheetName As String
    Dim skipSheetFound As Boolean
    Dim i As Integer
    
    Set wbOncore = ActiveWorkbook
    
    ' List sheet names you want to exclude
    skipSheets = Array("Protocol Information", "Billing Designation Legend", "Footnote Legend", "QCT Checklist")
    
    For Each ws In wbOncore.Worksheets
        wsName = ws.Name
        skipSheetFound = False
        For i = LBound(skipSheets) To UBound(skipSheets)
            If wsName = skipSheets(i) Then
                skipSheetFound = True
                Exit For
            End If
        Next i
        
        If Not skipSheetFound Then
            cboArms.AddItem wsName
        End If
    Next ws
    
    ' Set default to first item to avoid empty display
    cboArms.ListIndex = 0

End Sub

Private Sub cboArms_Click()
    Dim targetSheet As String
    targetSheet = cboArms.Value
    If targetSheet <> "" Then
        wbOncore.Activate
        Worksheets(targetSheet).Activate
        Cells(1, 1).Select
        'btnSubmit.Caption = cboArms.Value & Chr(10) & "Submit"
    End If
End Sub

Private Sub btnSubmit_Click()
    selectedSheetName = cboArms.Value
    Me.Hide   'hides the form and keeps it in the memory
End Sub

Private Sub btnAbort_Click()
    Me.Hide   'hides the form and keeps it in the memory
    'Unload Me   'closes the form and clears memory
End Sub

Public Property Get SelectedSheet() As String
    SelectedSheet = selectedSheetName
End Property

