Attribute VB_Name = "tool2_oncore"
'open file
'choose arm
'remove footnotes
'remove childrows
'copy to array
'update visit names
'update procedure names
'update grid

Sub UpdateIntBdgtGridToOncore()
Call OpenFileWithStatus
End Sub

Function OpenFileWithStatus()
' returns -1 if unsuccessful, 0 if successful

    Dim FileToOpen As Variant
    FileToOpen = Application.GetOpenFilename( _
        title:="Please choose a Billing Grid to open", _
        FileFilter:="Excel Files (*.xls*),*.xls*", _
        FilterIndex:=1, _
        ButtonText:="Open", _
        MultiSelect:=False)
        
    If FileToOpen = False Then
        MsgBox "No file selected.", vbExclamation
        OpenFileWithStatus = -1
    Else
        On Error GoTo ErrorHandler
        Workbooks.Open fileName:=FileToOpen
        On Error GoTo 0 ' Reset error handling
        OpenFileWithStatus = 0
    End If
    
    Exit Function
    
ErrorHandler:
    MsgBox "Error opening file: " & Err.Description, vbCritical
    OpenFileWithStatus = -1
End Function

