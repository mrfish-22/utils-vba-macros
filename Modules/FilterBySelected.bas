Attribute VB_Name = "FilterBySelected"
Option Explicit

Sub FilterBySelected()
Attribute FilterBySelected.VB_ProcData.VB_Invoke_Func = "Y\n14"

    On Error GoTo ErrorHandler
    
    Dim criteria As Range
    Dim col As Integer
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim current_Cell As Range
    
    Set ws = ActiveSheet
    
    If Not IsEmpty(ActiveCell.Value) Then
        Set criteria = ActiveCell
        col = ActiveCell.Column
        Set current_Cell = ActiveCell
    Else
        MsgBox "Zaznacz komórkê zawieraj¹c¹ kryterium.", vbExclamation
        Exit Sub
    End If
    
    On Error Resume Next
    Set tbl = ws.ListObjects(1)
    On Error GoTo 0
    
    If tbl Is Nothing Then
        MsgBox "Brak tabeli na bie¿¹cym arkuszu.", vbExclamation
        Exit Sub
    End If
    
    tbl.Range.AutoFilter Field:=col, Criteria1:=criteria.Value
    
    ws.Cells(1, col).Select
    current_Cell.Select
        
    Application.ScreenUpdating = True

    Exit Sub
    
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical

End Sub

