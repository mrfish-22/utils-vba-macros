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
    
    ' U�ywamy bie��cego arkusza
    Set ws = ActiveSheet
    
    ' Sprawd�, czy aktywna kom�rka zawiera kryterium
    If ActiveCell Is Nothing Or IsEmpty(ActiveCell.Value) Then
        MsgBox "Zaznacz kom�rk� zawieraj�c� kryterium.", vbExclamation
        Exit Sub
    End If
    Set criteriaCell = ActiveCell
    critVal = criteriaCell.Value
    
    ' Pobierz pierwsz� tabel� z arkusza
    On Error Resume Next
    Set tbl = ws.ListObjects(1)
    On Error GoTo 0
    If tbl Is Nothing Then
        MsgBox "Brak tabeli na bie��cym arkuszu.", vbExclamation
        Exit Sub
    End If
    
    ' Oblicz indeks pola (kolumny w tabeli) odpowiadaj�cy aktywnej kom�rce.
    colWorksheet = criteriaCell.Column
    fieldIndex = colWorksheet - tbl.Range.Columns(1).Column + 1
    
    ' Sprawd�, czy w tabeli jest aktywny jakikolwiek filtr
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
    
    ' Statyczny s�ownik do przechowywania wykluczonych warto�ci dla poszczeg�lnych kolumn
    Static exclusionDict As Object
    If exclusionDict Is Nothing Then
        Set exclusionDict = CreateObject("Scripting.Dictionary")
    Else
        ' Je�li tabela jest wy�wietlona w ca�o�ci (bez aktywnych filtr�w), resetujemy s�ownik
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
    
    ' Dodaj wybrane kryterium do listy wyklucze� (je�li jeszcze nie istnieje)
    If Not colExclusions.exists(critVal) Then
        colExclusions.Add critVal, critVal
    End If
    
    ' Pobierz wszystkie unikalne warto�ci z danej kolumny tabeli (ca�y DataBodyRange)
    Set uniqueValues = CreateObject("Scripting.Dictionary")
    Set dataRange = tbl.ListColumns(fieldIndex).DataBodyRange
    For Each cell In dataRange.Cells
        If Not uniqueValues.exists(cell.Value) Then
            uniqueValues.Add cell.Value, cell.Value
        End If
    Next cell
    
    ' Zbuduj kolekcj� warto�ci do wy�wietlenia (wszystkie unikalne, poza tymi wykluczonymi)
    Set allowedValues = New Collection
    Dim key As Variant
    For Each key In uniqueValues.Keys
        If Not colExclusions.exists(key) Then
            allowedValues.Add key
        End If
    Next key
    
    ' Je�li nie ma warto�ci do wy�wietlenia, poinformuj u�ytkownika i opcjonalnie przywr�� pe�ny widok
    If allowedValues.Count = 0 Then
        MsgBox "Brak warto�ci do wy�wietlenia po odfiltrowaniu.", vbInformation
        tbl.AutoFilter.ShowAllData
        Exit Sub
    End If
    
    ' Konwertuj kolekcj� do tablicy � AutoFilter z operatorem xlFilterValues wymaga tablicy
    ReDim arrAllowed(0 To allowedValues.Count - 1)
    For i = 1 To allowedValues.Count
        arrAllowed(i - 1) = allowedValues(i)
    Next i
    
    ' Ustaw filtr: wy�wietl tylko te wiersze, kt�rych warto�� w danej kolumnie znajduje si� w tablicy arrAllowed
    tbl.Range.AutoFilter Field:=fieldIndex, Criteria1:=arrAllowed, Operator:=xlFilterValues
    
    ' Opcjonalnie: przywr�� zaznaczenie do kom�rki z kryterium
    ws.Cells(tbl.Range.Row, colWorksheet).Select
    criteriaCell.Select
    
    Application.ScreenUpdating = True
    Exit Sub
    
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
End Sub


