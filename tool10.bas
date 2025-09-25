Attribute VB_Name = "tool10"
Option Explicit


Sub tool10_RemoveReplyFromThreadedCommentByText()

'    Dim ws As Worksheet
    Dim Rng As Range
    Dim cell As Range
    Dim threadedComment As CommentThreaded
    Dim reply As CommentThreaded
    Dim searchText As String
    Dim i As Long

    Set Rng = SelectRange
    If Rng Is Nothing Then Exit Sub

    ' The text to search for within the replies
    searchText = GetUserText
    If searchText = "" Then Exit Sub

    For Each cell In Rng
        ' Check if the cell has a threaded comment
        If Not cell.CommentThreaded Is Nothing Then
            Set threadedComment = cell.CommentThreaded
            If InStr(1, threadedComment.Text, searchText, vbTextCompare) > 0 Then
                threadedComment.Delete
            Else
                ' Loop through the replies in reverse order (important when deleting items from a collection)
                For i = threadedComment.Replies.count To 1 Step -1
                    Set reply = threadedComment.Replies.item(i)
        
                    ' Check if the reply's text contains the search text
                    If InStr(1, reply.Text, searchText, vbTextCompare) > 0 Then
                        ' If found, delete the reply
                        reply.Delete
                        'MsgBox "Reply containing '" & searchText & "' deleted from cell " & cell.Address, vbInformation
                        'Exit For ' Exit the loop after deleting the first matching reply
                    End If
                Next i
            End If
        Else
            'MsgBox "Cell " & rng.Address & " does not contain a threaded comment.", vbInformation
        End If
    Next cell

End Sub

Private Function SelectRange() As Range
    Dim Rng As Range
    On Error Resume Next
    Set Rng = Application.InputBox( _
        prompt:="Select a range you'd like the utility to process", _
        title:="Comment Find/Delete Range Selector", _
        Type:=8)
    On Error GoTo 0
    
    If Rng Is Nothing Then
        MsgBox "You did not select a range."
        Exit Function
    End If
    

    Set SelectRange = Rng

End Function

Private Function GetUserText() As String
    Dim userText As String

    userText = InputBox( _
            prompt:="Please enter the search text you'd like to look for in the comments:", _
            title:="Comment Search Text Input")
    
    If userText = "" Then
        MsgBox "No input was provided."
    End If
    
    GetUserText = userText
    
End Function

