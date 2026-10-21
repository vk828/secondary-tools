VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTool2PairIntBdgtVisitNames 
   Caption         =   "Update Internal Budget Visit/Procedure Names"
   ClientHeight    =   4725
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14265
   OleObjectBlob   =   "frmTool2PairIntBdgtVisitNames.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmTool2PairIntBdgtVisitNames"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private terminateFlag As Boolean 'initialized to false

Private ibRng As Range  'ib range that might have names to be overwritten

Private ibColl As Collection        'collection of unpaired Int Bdgt names
Private oncoreColl As Collection    'collection of unpaired OnCore names
Private ibArr As Variant
Private oncoreArr As Variant
Private ib_visitsRow As Range
Private oncore_visitsRow As Range


Sub SetCbos(collOne As Collection, collTwo As Collection, _
            arrayOne() As Variant, arrayTwo() As Variant, _
            rngOne As Range, rngTwo As Range, _
            frmCaption As String, lblInstructions As String, _
            doneBtnLbl As String)
'subroutine sets initial comboboxes

    Dim item As Variant
    
    Set ibColl = collOne
    Set oncoreColl = collTwo
    
    Me.cboIntBdgtName.Clear
    Me.cboOncoreName.Clear
    
    ibArr = arrayOne
    oncoreArr = arrayTwo
    
    Set ib_visitsRow = rngOne
    Set oncore_visitsRow = rngTwo
    
    ' Set up Internal Budget ComboBox
    Me.cboIntBdgtName.ColumnCount = 2
    Me.cboIntBdgtName.BoundColumn = 1  ' .Value will return the index
    Me.cboIntBdgtName.TextColumn = 2   ' The textbox part will display the unique visit name
    Me.cboIntBdgtName.ColumnWidths = "0pt;657pt" ' hides index

    ' Set up OnCore ComboBox
    Me.cboOncoreName.ColumnCount = 2
    Me.cboOncoreName.BoundColumn = 1
    Me.cboOncoreName.TextColumn = 2
    Me.cboOncoreName.ColumnWidths = "0pt;657pt"
    
    Me.Caption = frmCaption
    Me.lblInstruction = lblInstructions
    Me.btnDone.Caption = doneBtnLbl

    For Each item In ibColl
        Me.cboIntBdgtName.AddItem
        Me.cboIntBdgtName.List(Me.cboIntBdgtName.ListCount - 1, 0) = item
        Me.cboIntBdgtName.List(Me.cboIntBdgtName.ListCount - 1, 1) = ibArr(3, item)
    Next item

    For Each item In oncoreColl
        Me.cboOncoreName.AddItem
        Me.cboOncoreName.List(Me.cboOncoreName.ListCount - 1, 0) = item
        Me.cboOncoreName.List(Me.cboOncoreName.ListCount - 1, 1) = oncoreArr(3, item)
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
    Call tool2_overwriteNames.ReportUnpaired(OncoreArray)
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

Public Property Get IbArray() As Variant
    IbArray = ibArr
End Property

Public Property Get OncoreArray() As Variant
    OncoreArray = oncoreArr
End Property

Private Sub btnDone_Click()
    Me.Hide
End Sub

Private Sub btnPairSelected_Click()
'this sub assumes that there are NO duplicates
'when 'Pair' button is clicked, this function is called

    Dim ibArrayIndex As String
    Dim oncoreArrayIndex As String
    Dim cell As Range
    Dim ibCboIndex As Integer
    Dim oncoreCboIndex As Integer
    
    ibCboIndex = Me.cboIntBdgtName.ListIndex
    oncoreCboIndex = Me.cboOncoreName.ListIndex
    
    'if something is selected in both comboboxes, remove these items
    If ibCboIndex <> -1 And oncoreCboIndex <> -1 Then
        
        ibArrayIndex = Me.cboIntBdgtName.value
        oncoreArrayIndex = Me.cboOncoreName.value
        
        'remove item from combobox
        Me.cboIntBdgtName.RemoveItem ibCboIndex
        
        'remove item from collection
        ibColl.Remove ibCboIndex + 1
        
        'update combobox view
        Call UpdatedCboView(Me.cboIntBdgtName, ibCboIndex)
            
        'remove item from combobox
        Me.cboOncoreName.RemoveItem oncoreCboIndex
        
        'remove item from collection
        oncoreColl.Remove oncoreCboIndex + 1
        
        'update combobox view
        Call UpdatedCboView(Me.cboOncoreName, oncoreCboIndex)
        
        ' update map for paired item: update names and corresponding column in oncore
        ibArr(5, ibArrayIndex) = oncoreArr(4, oncoreArrayIndex)
        ibArr(1, ibArrayIndex) = oncoreArr(1, oncoreArrayIndex)
        ibArr(2, ibArrayIndex) = oncoreArr(2, oncoreArrayIndex)
        
        ' update oncoreArray to show there is a match
        oncoreArr(5, oncoreArrayIndex) = 1
                
        'update name on the internal budget

        ' visit
        Set cell = ib_visitsRow.Worksheet.Cells(ib_visitsRow.Row, ibArr(4, ibArrayIndex))
    
        If cell.value <> ibArr(2, ibArrayIndex) Then
            With cell
                .value = ibArr(2, ibArrayIndex)
                .Interior.color = RGB(255, 255, 0)  'fill yellow color
                'activate workbook
                .Worksheet.Parent.Activate
                'activate worksheet
                .Worksheet.Activate
                'select
                .Select
            End With
        End If
        
        ' segment
        Set cell = cell.Offset(-1, 0)
        If cell.value <> ibArr(1, ibArrayIndex) Then
            With cell
                .value = ibArr(1, ibArrayIndex)
                .Interior.color = RGB(255, 255, 0)  'fill yellow color
                'activate workbook
                .Worksheet.Parent.Activate
                'activate worksheet
                .Worksheet.Activate
                'select
                .Select
            End With
        End If
    End If

End Sub

Private Sub cboIntBdgtName_Change()
'    Dim cell As Range
'    Set cell = FindCell(Me.cboIntBdgtName.value)
'
'    If Not cell Is Nothing Then
'        With cell
'            'activate workbook
'            .Worksheet.Parent.Activate
'            'activate worksheet
'            .Worksheet.Activate
'            'select
'            .Select
'        End With
'    End If
End Sub

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


