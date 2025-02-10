Option Explicit

Sub FilterBySelected()
    ' Enable error handling
    On Error GoTo ErrorHandler
    
    Dim criteria As Range           ' Variable for the filtering criteria (cell)
    Dim col As Integer              ' Column number of the active cell
    Dim ws As Worksheet             ' Active worksheet
    Dim tbl As ListObject           ' Table object (first table in the worksheet)
    Dim current_Cell As Range       ' The currently active cell
    
    ' Set the active worksheet
    Set ws = ActiveSheet
    
    ' Check if the active cell is not empty
    If Not IsEmpty(ActiveCell.Value) Then
        Set criteria = ActiveCell       ' Use the active cell as the criteria
        col = ActiveCell.Column         ' Get the column number of the active cell
        Set current_Cell = ActiveCell   ' Save the active cell for later reference
    Else
        ' Notify the user to select a cell with filtering criteria
        MsgBox "Please select a cell containing the filtering criteria.", vbExclamation
        Exit Sub
    End If
    
    ' Attempt to set the table (first ListObject in the worksheet)
    On Error Resume Next
    Set tbl = ws.ListObjects(1)
    On Error GoTo 0
    
    ' If no table is found, notify the user and exit
    If tbl Is Nothing Then
        MsgBox "No table found on the current worksheet.", vbExclamation
        Exit Sub
    End If
    
    ' Apply the AutoFilter on the table using the criteria from the active cell
    tbl.Range.AutoFilter Field:=col, Criteria1:=criteria.Value
    
    ' Reselect the active cell to maintain focus after filtering
    ws.Cells(1, col).Select
    current_Cell.Select
        
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    ' Display an error message if an error occurs
    MsgBox "Error: " & Err.Description, vbCritical
End Sub
