Option Explicit

Sub FilterBySelectedExclude()
    ' Enable error handling
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet             ' Active worksheet
    Dim tbl As ListObject           ' Table object in the worksheet
    Dim criteriaCell As Range       ' Cell that contains the filtering criteria
    Dim fieldIndex As Long          ' Index of the column in the table corresponding to the criteria cell
    Dim colWorksheet As Long        ' Column number of the criteria cell in the worksheet
    Dim dataRange As Range          ' Range of data in the table column
    Dim cell As Range               ' Loop variable for cells in the data range
    Dim uniqueValues As Object      ' Dictionary to store unique values in the column
    Dim allowedValues As Collection ' Collection for values allowed (not excluded)
    Dim arrAllowed() As Variant     ' Array of allowed values for use with AutoFilter
    Dim i As Long                   ' Loop counter
    Dim critVal As Variant          ' The filtering criteria value
    
    ' Use the active worksheet
    Set ws = ActiveSheet
    
    ' Check if the active cell contains a criteria value
    If ActiveCell Is Nothing Or IsEmpty(ActiveCell.Value) Then
        MsgBox "Please select a cell containing the filtering criteria.", vbExclamation
        Exit Sub
    End If
    Set criteriaCell = ActiveCell
    critVal = criteriaCell.Value
    
    ' Get the first table from the worksheet
    On Error Resume Next
    Set tbl = ws.ListObjects(1)
    On Error GoTo 0
    If tbl Is Nothing Then
        MsgBox "No table found on the current worksheet.", vbExclamation
        Exit Sub
    End If
    
    ' Calculate the field index (column in the table) corresponding to the active cell
    colWorksheet = criteriaCell.Column
    fieldIndex = colWorksheet - tbl.Range.Columns(1).Column + 1
    
    ' Check if any filter is currently active in the table
    Dim filterActive As Boolean
    filterActive = False
    Dim f As Variant
    If Not tbl.AutoFilter Is Nothing Then
        For Each f In tbl.AutoFilter.Filters
            If Not f Is Nothing Then
                If f.On Then
                    filterActive = True
                    Exit For
                End If
            End If
        Next f
    End If
    
    ' Static dictionary to store excluded values for specific columns
    Static exclusionDict As Object
    If exclusionDict Is Nothing Then
        Set exclusionDict = CreateObject("Scripting.Dictionary")
    Else
        ' If the table is fully displayed (no active filters), reset the dictionary
        If Not filterActive Then
            Set exclusionDict = CreateObject("Scripting.Dictionary")
        End If
    End If
    
    Dim colExclusions As Object
    ' Get or create the exclusions dictionary for the current field index
    If exclusionDict.exists(fieldIndex) Then
        Set colExclusions = exclusionDict(fieldIndex)
    Else
        Set colExclusions = CreateObject("Scripting.Dictionary")
        exclusionDict.Add fieldIndex, colExclusions
    End If
    
    ' Add the selected criteria value to the exclusions list (if it doesn't already exist)
    If Not colExclusions.exists(critVal) Then
        colExclusions.Add critVal, critVal
    End If
    
    ' Retrieve all unique values from the specified column of the table (entire DataBodyRange)
    Set uniqueValues = CreateObject("Scripting.Dictionary")
    Set dataRange = tbl.ListColumns(fieldIndex).DataBodyRange
    For Each cell In dataRange.Cells
        If Not uniqueValues.exists(cell.Value) Then
            uniqueValues.Add cell.Value, cell.Value
        End If
    Next cell
    
    ' Build a collection of values to display (all unique values excluding those in the exclusions list)
    Set allowedValues = New Collection
    Dim key As Variant
    For Each key In uniqueValues.Keys
        If Not colExclusions.exists(key) Then
            allowedValues.Add key
        End If
    Next key
    
    ' If no values are left to display, notify the user and optionally reset the filter
    If allowedValues.Count = 0 Then
        MsgBox "No values to display after filtering.", vbInformation
        tbl.AutoFilter.ShowAllData
        Exit Sub
    End If
    
    ' Convert the collection to an array – AutoFilter with the xlFilterValues operator requires an array
    ReDim arrAllowed(0 To allowedValues.Count - 1)
    For i = 1 To allowedValues.Count
        arrAllowed(i - 1) = allowedValues(i)
    Next i
    
    ' Apply the filter: display only the rows where the column value is in the allowed values array
    tbl.Range.AutoFilter Field:=fieldIndex, Criteria1:=arrAllowed, Operator:=xlFilterValues
    
    ' Optionally, restore the selection to the criteria cell
    ws.Cells(tbl.Range.Row, colWorksheet).Select
    criteriaCell.Select
    
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    ' Display an error message if something goes wrong
    MsgBox "Error: " & Err.Description, vbCritical
End Sub
