Attribute VB_Name = "FilterBySelectedExclude"
Option Explicit

Sub FilterBySelectedExclude()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim criteriaCell As Range
    Dim fieldIndex As Long
    Dim colWorksheet As Long
    Dim dataRange As Range
    Dim cell As Range
    Dim uniqueValues As Object
    Dim allowedValues As Collection
    Dim arrAllowed() As Variant
    Dim i As Long
    Dim critVal As Variant
    
    ' U¿ywamy bie¿¹cego arkusza
    Set ws = ActiveSheet
    
    ' SprawdŸ, czy aktywna komórka zawiera kryterium
    If ActiveCell Is Nothing Or IsEmpty(ActiveCell.Value) Then
        MsgBox "Zaznacz komórkê zawieraj¹c¹ kryterium.", vbExclamation
        Exit Sub
    End If
    Set criteriaCell = ActiveCell
    critVal = criteriaCell.Value
    
    ' Pobierz pierwsz¹ tabelê z arkusza
    On Error Resume Next
    Set tbl = ws.ListObjects(1)
    On Error GoTo 0
    If tbl Is Nothing Then
        MsgBox "Brak tabeli na bie¿¹cym arkuszu.", vbExclamation
        Exit Sub
    End If
    
    ' Oblicz indeks pola (kolumny w tabeli) odpowiadaj¹cy aktywnej komórce.
    colWorksheet = criteriaCell.Column
    fieldIndex = colWorksheet - tbl.Range.Columns(1).Column + 1
    
    ' SprawdŸ, czy w tabeli jest aktywny jakikolwiek filtr
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
    
    ' Statyczny s³ownik do przechowywania wykluczonych wartoœci dla poszczególnych kolumn
    Static exclusionDict As Object
    If exclusionDict Is Nothing Then
        Set exclusionDict = CreateObject("Scripting.Dictionary")
    Else
        ' Jeœli tabela jest wyœwietlona w ca³oœci (bez aktywnych filtrów), resetujemy s³ownik
        If Not filterActive Then
            Set exclusionDict = CreateObject("Scripting.Dictionary")
        End If
    End If
    
    Dim colExclusions As Object
    If exclusionDict.exists(fieldIndex) Then
        Set colExclusions = exclusionDict(fieldIndex)
    Else
        Set colExclusions = CreateObject("Scripting.Dictionary")
        exclusionDict.Add fieldIndex, colExclusions
    End If
    
    ' Dodaj wybrane kryterium do listy wykluczeñ (jeœli jeszcze nie istnieje)
    If Not colExclusions.exists(critVal) Then
        colExclusions.Add critVal, critVal
    End If
    
    ' Pobierz wszystkie unikalne wartoœci z danej kolumny tabeli (ca³y DataBodyRange)
    Set uniqueValues = CreateObject("Scripting.Dictionary")
    Set dataRange = tbl.ListColumns(fieldIndex).DataBodyRange
    For Each cell In dataRange.Cells
        If Not uniqueValues.exists(cell.Value) Then
            uniqueValues.Add cell.Value, cell.Value
        End If
    Next cell
    
    ' Zbuduj kolekcjê wartoœci do wyœwietlenia (wszystkie unikalne, poza tymi wykluczonymi)
    Set allowedValues = New Collection
    Dim key As Variant
    For Each key In uniqueValues.Keys
        If Not colExclusions.exists(key) Then
            allowedValues.Add key
        End If
    Next key
    
    ' Jeœli nie ma wartoœci do wyœwietlenia, poinformuj u¿ytkownika i opcjonalnie przywróæ pe³ny widok
    If allowedValues.Count = 0 Then
        MsgBox "Brak wartoœci do wyœwietlenia po odfiltrowaniu.", vbInformation
        tbl.AutoFilter.ShowAllData
        Exit Sub
    End If
    
    ' Konwertuj kolekcjê do tablicy – AutoFilter z operatorem xlFilterValues wymaga tablicy
    ReDim arrAllowed(0 To allowedValues.Count - 1)
    For i = 1 To allowedValues.Count
        arrAllowed(i - 1) = allowedValues(i)
    Next i
    
    ' Ustaw filtr: wyœwietl tylko te wiersze, których wartoœæ w danej kolumnie znajduje siê w tablicy arrAllowed
    tbl.Range.AutoFilter Field:=fieldIndex, Criteria1:=arrAllowed, Operator:=xlFilterValues
    
    ' Opcjonalnie: przywróæ zaznaczenie do komórki z kryterium
    ws.Cells(tbl.Range.Row, colWorksheet).Select
    criteriaCell.Select
    
    Application.ScreenUpdating = True
    Exit Sub
    
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
End Sub


