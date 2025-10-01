Attribute VB_Name = "tool2_cases"
'Author/Developer: Vadim Krifuks
'Last Updated: 22 September 2025

Option Explicit
Option Private Module

Sub ProcedureNotFound(procedureRng As Range, gridRowRng As Range, fillColor As Long)

    If Not (AreAllCellsEmpty(procedureRng) And AreAllCellsEmpty(gridRowRng)) Then
        Dim msg As String
        msg = AssembleComment("procedure not found in OnCore; row skipped")
        
        Call AddComment(procedureRng, msg)
        
        procedureRng.Interior.color = fillColor
        gridRowRng.Interior.color = fillColor
    End If
    
End Sub

Sub VisitNotFound(visitRng As Range, gridColRng As Range, fillColor As Long)

    If Not (AreAllCellsEmpty(visitRng) And AreAllCellsEmpty(gridColRng)) Then
        Dim msg As String
        msg = AssembleComment("visit not found in OnCore; column skipped")
    
        Call AddComment(visitRng, msg)
        
        visitRng.Interior.color = fillColor
        gridColRng.Interior.color = fillColor
    End If

End Sub

Sub PrevAndCurrEqualNothing(cell As Range, fillColor As Long)
    cell.Interior.color = fillColor
End Sub

Function AreAllCellsEmpty(rng As Range) As Boolean
    Dim cell As Range
    For Each cell In rng.Cells
        If CStr(cell.Value) <> "" Then
            AreAllCellsEmpty = False
            Exit Function
        End If
    Next cell
    AreAllCellsEmpty = True
End Function

Sub PrevAndCurrEqualX(cell As Range, fillColor As Long)
    cell.Interior.color = fillColor
End Sub

Sub PrevNothingCurrX(cell As Range, fillColor As Long, prevValue As Variant, currValue As Variant)
    Dim msg As String
    msg = AssembleComment("auto-updated to current onCore value", prevValue, currValue)

    With cell
        .Interior.color = fillColor
        .Value = currValue
    End With
                
    Call AddComment(cell, msg)

End Sub

Sub PrevInvoiceCurrOne(cell As Range, fillColor As Long, prevValue As Variant, currValue As Variant)
    Dim msg As String
    msg = AssembleComment("auto-kept previous internal budget value", prevValue, currValue)
    
    cell.Interior.color = fillColor
    
    Call AddComment(cell, msg)
    
End Sub

Function PrevXCurrY(ibCell As Range, _
                oncoreCell As Range, _
                visitName As String, _
                procedureName As String, _
                updateFillColor As Long, _
                keepFillColor As Long, _
                prevValue As Variant, _
                currValue As Variant) As Integer
    
    Dim ib_sheetName As String
    Dim oncore_sheetName As String
    Dim ib_valueStr As String
    Dim oncore_valueStr As String
    Dim targetAddressStr As String
    Dim cmtFlagStr As String
    Dim userFormMsg As String
    Dim msg As String

    'get sheet names
    oncore_sheetName = oncoreCell.Worksheet.name
    ib_sheetName = ibCell.Worksheet.name
    
    'adjust ib and oncore values for readibility
    If CStr(prevValue) = "" Then
        ib_valueStr = "[empty]"
    Else
        ib_valueStr = CStr(prevValue)
    End If
    
    If CStr(currValue) = "" Then
        oncore_valueStr = "[empty]"
    Else
        oncore_valueStr = CStr(currValue)
    End If
    
    'get an address of a visit/procedure cell
    targetAddressStr = ibCell.Address(RowAbsolute:=False, ColumnAbsolute:=False)
        
    'activate visit/procedure cell
    With ibCell
        ' to select a cell, workbook and worsheet need to be active
        ' 1) we activate a workbook the range belongs to
        .Worksheet.Parent.Activate
        ' 2) we activate the worksheet the range belongs to
        .Worksheet.Activate
        ' 3) select the cell
        .Select
    End With
    
    'check if there is a comment in visit/procedure cell
    If IsComment(ibCell) Then
        cmtFlagStr = "Yes"
    Else
        cmtFlagStr = "No"
    End If
    
    userFormMsg = "OnCore Billing Grid (source) sheet: " & oncore_sheetName & Chr(10) _
                    & "Internal Budget (target) sheet: " & ib_sheetName & Chr(10) _
                    & Chr(10) _
                    & "Procedure: " & procedureName & Chr(10) _
                    & "Visit: " & visitName & Chr(10) _
                    & Chr(10) _
                    & "Cell: " & targetAddressStr & Chr(10) _
                    & "Comment: " & cmtFlagStr & Chr(10) _
                    & Chr(10) _
                    & "OnCore value: " & oncore_valueStr & Chr(10) _
                    & "Internal budget value: " & ib_valueStr & Chr(10) _
                    & Chr(10) _
                    & "Would you like to update to OnCore value?"
    
    'Display the form modelessly
    form_amds.UserResponse = ""
    form_amds.Label1 = userFormMsg
    Call OpenForm

    'Wait for the user to respond (polling loop)
    Do While form_amds.Visible
        DoEvents
    Loop
    
    'Check the response
    'case_a: update
    If form_amds.UserResponse = "update" Then
        With ibCell
            .Interior.color = updateFillColor
            .Value = currValue
        End With
        
        msg = AssembleComment("analyst updated to current onCore value", prevValue, currValue)
        Call AddComment(ibCell, msg)
    'case_b: keep
    ElseIf form_amds.UserResponse = "keep" Then
        ibCell.Interior.color = keepFillColor
        
        msg = AssembleComment("analyst kept previous internal budget value", prevValue, currValue)
        Call AddComment(ibCell, msg)
        
    'case_c: exit
    ElseIf form_amds.UserResponse = "" Then
        PrevXCurrY = 1
        Exit Function
    End If
    
    Unload form_amds
    PrevXCurrY = 0
    
End Function

Function IsComment(rng As Range)
'function that returns true/false telling the user if there is comment in a cell
    
    Dim cmt As CommentThreaded
    
    Set cmt = rng.CommentThreaded
    
    If cmt Is Nothing Then
        IsComment = False
    Else
        IsComment = True
    End If

End Function

Sub OpenForm()
'opens user form at a specified location on the screen

    With form_amds
        'manual positioning
        .StartUpPosition = 0
        
        'set left and top location
        'currently set to bottom right of excel app
        .Left = Application.Left + (1 * Application.Width) - (1 * .Width)
        .Top = Application.Top + (1 * Application.Height) - (1 * .Height)
        
        'open in modless mode
        .Show vbModeless
    End With

End Sub

Function AssembleComment(middleStr As String, Optional prev As Variant, Optional curr As Variant) As String

    Dim startStr As String
    Dim endStr As String
    
    startStr = "[" & Format(Date, "ddmmmyy") & " tool2 execution] "
    
    If Not (IsMissing(prev) And IsMissing(curr)) Then
    
        If prev = "" Then prev = "[empty]"
        If curr = "" Then curr = "[empty]"
    
        endStr = Chr(10) & _
                " -prev int bdgt value: " & CStr(prev) & Chr(10) & _
                " -curr onCore value: " & CStr(curr)
    End If
    
    AssembleComment = startStr & middleStr & endStr

End Function

Sub AddComment(cell As Range, commentText As String)

    'if there is a note, convert it to comment
    Call ConvertNoteToComment(cell)
    
    'add comment
    With cell
        If .CommentThreaded Is Nothing Then
            .AddCommentThreaded (commentText)
        Else
            .CommentThreaded.AddReply (commentText)
        End If
    End With
End Sub

Sub ConvertNoteToComment(cell As Range)
    
    Dim commentText As String
    
    If Not cell.comment Is Nothing Then
        commentText = cell.comment.Text
        cell.comment.Delete
        cell.AddCommentThreaded ("[previous note]" & commentText)
    End If
    
End Sub
