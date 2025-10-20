VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTool2PairNames 
   Caption         =   "Update Internal Budget Visit/Procedure Names"
   ClientHeight    =   4725
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14265
   OleObjectBlob   =   "frmTool2PairNames.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmTool2PairNames"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private terminateFlag As Boolean 'initialized to false

Private ibRng As Range  'ib range that might have names to be overwritten

Private ibColl As Collection        'collection of unpaired Int Bdgt names
Private oncoreColl As Collection    'collection of unpaired OnCore names


Sub SetCbos(collOne As Collection, collTwo As Collection, _
            frmCaption As String, lblInstructions As String, _
            doneBtnLbl As String)
'subroutine sets initial comboboxes

    Dim item As Variant
    
    Set ibColl = collOne
    Set oncoreColl = collTwo
    
    Me.cboIntBdgtName.Clear
    Me.cboOncoreName.Clear
    
    Me.Caption = frmCaption
    Me.lblInstruction = lblInstructions
    Me.btnDone.Caption = doneBtnLbl

    For Each item In ibColl
        Me.cboIntBdgtName.AddItem item
    Next item

    For Each item In oncoreColl
        Me.cboOncoreName.AddItem item
    Next item

    ' Set cbo view to the first item to avoid empty display
    Call UpdatedCboView(Me.cboIntBdgtName, 0)
    Call UpdatedCboView(Me.cboOncoreName, 0)
    
    'font size in the combobox; original is 8 and it is too small
    Me.cboIntBdgtName.Font.Size = 11
    Me.cboOncoreName.Font.Size = 11

End Sub

Private Sub btnExit_Click()
'sub that is called when 'Exit' button is clicked
    terminateFlag = True
    Me.Hide   'hides the form and keeps it in the memory
End Sub

Private Sub btnReportUnpaired_Click()
    Call tool2_overwriteNames.ReportUnpaired
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
'this sub controls what happens when a user clicks 'X' button in the right top corner
'mimic behavior of Private Sub btnExit_Click()

    If CloseMode = vbFormControlMenu Then ' Detect click on the "X" button
        Cancel = True    ' Cancel the default close action
        terminateFlag = True
        Me.Hide          ' Hide the form, just like btnExit_Click does
    End If
End Sub

Public Property Get IsTerminated() As Boolean
'property to read the selected arm
    IsTerminated = terminateFlag
End Property

Public Property Set RangeToForm(rng As Range)
'property that sets ibRng
    Set ibRng = rng
End Property

Public Property Get IbCollection() As Collection
    Set IbCollection = ibColl
End Property

Public Property Get OncoreCollection() As Collection
    Set OncoreCollection = oncoreColl
End Property

Private Sub btnDone_Click()
    Me.Hide
End Sub

Private Sub btnPairSelected_Click()
'this sub assumes that there are NO duplicates
'when 'Pair' button is clicked, this function is called

    Dim ibName As String
    Dim oncoreName As String
    Dim cell As Range
    Dim ibIndex As Integer
    Dim oncoreIndex As Integer
    
    ibIndex = Me.cboIntBdgtName.ListIndex
    oncoreIndex = Me.cboOncoreName.ListIndex
    
    'if something is selected in both comboboxes, remove these items
    If ibIndex <> -1 And oncoreIndex <> -1 Then
        
        ibName = Me.cboIntBdgtName.Value
        oncoreName = Me.cboOncoreName.Value
        
        'remove item from combobox
        Me.cboIntBdgtName.RemoveItem ibIndex
        
        'remove item from collection
        ibColl.Remove ibIndex + 1
        
        'update combobox view
        Call UpdatedCboView(Me.cboIntBdgtName, ibIndex)
            
        'remove item from combobox
        Me.cboOncoreName.RemoveItem oncoreIndex
        
        'remove item from collection
        oncoreColl.Remove oncoreIndex + 1
        
        'update combobox view
        Call UpdatedCboView(Me.cboOncoreName, oncoreIndex)
        
        'update name on the internal budget
        Set cell = FindCell(ibName)
        With cell
            'activate workbook
            .Worksheet.Parent.Activate
            'activate worksheet
            .Worksheet.Activate
            'select
            .Select
            .Value = oncoreName
            .Interior.color = RGB(255, 255, 0)  'fill yellow color
        End With
    End If

End Sub

Private Sub cboIntBdgtName_Change()
    Dim cell As Range
    Set cell = FindCell(Me.cboIntBdgtName.Value)
    
    If Not cell Is Nothing Then
        With cell
            'activate workbook
            .Worksheet.Parent.Activate
            'activate worksheet
            .Worksheet.Activate
            'select
            .Select
        End With
    End If
End Sub

Private Function FindCell(ibName As String) As Range

    Dim cell As Range
    
    For Each cell In ibRng.Cells
        If Application.WorksheetFunction.Trim(Application.WorksheetFunction.Clean(cell.Value)) = ibName Then
            Set FindCell = cell
            Exit For 'ensures that the first found instance is returned
        End If
    Next cell

End Function

Private Sub UpdatedCboView(cbo As ComboBox, index As Integer)

    Dim cboEmptyMsg As String
    
    cboEmptyMsg = "all names have a pair"
    
    'if combobox is empty, 1)show a message, 2)disable combobox, and 3)disable 'Pair' button
    If cbo.ListCount = 0 Then
        cbo.Text = cboEmptyMsg
        cbo.Enabled = False
        Me.btnPairSelected.Enabled = False
    Else
        If cbo.ListCount > index Then
            cbo.ListIndex = index
        Else
            cbo.ListIndex = Application.WorksheetFunction.Min(index - 1, cbo.ListCount - 1)
        End If
    End If

End Sub
