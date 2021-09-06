Option Explicit

Sub CreateTable(rangeName, tableName)

    Dim startRange As Object
    Dim endRange As Object
    Dim i As Integer
    Dim col As Variant
    Dim columns(3) As String
    Dim columnsIndexes(3) As Integer
    
    columns(0) = "Наименование"
    columns(1) = "Цена"
    columns(2) = "Сайт"
    columns(3) = "Ссылка"
    
    For i = 0 To 3
        columnsIndexes(i) = i
    Next i
    
    Set startRange = ActiveSheet.range(rangeName)
    Set endRange = startRange.Offset(0, 3)

    ActiveSheet.ListObjects.Add(xlSrcRange, range(startRange.Address & ":" & endRange.Address), , xlYes).name = tableName
    ActiveSheet.ListObjects(tableName).TableStyle = "TableStyleLight8"
    
    For Each col In columnsIndexes
        If col = 0 Then
            startRange.Value = columns(col)
        Else
            startRange.Offset(0, col).Value = columns(col)
        End If
    Next col
    
End Sub

Sub addRow(tableName, parseData)
    Dim table As Object
    Dim row As Object
    
    Set table = ActiveSheet.ListObjects(tableName)
    
    If table Is Nothing Then
        MsgBox ("Нет таблицы с таким именем")
        Exit Sub
    End If
    
    Set row = table.ListRows.Add
    
    row.range.Item(1).Value = parseData(0)
'    If (row.range.Item(1).ColumnWidth < Len(name)) Then
'        row.range.Item(1).ColumnWidth = Len(name)
'    End If
    row.range.Item(2).Value = parseData(1)
    If (row.range.Item(2).ColumnWidth < Len(parseData(1))) Then
        row.range.Item(2).ColumnWidth = Len(parseData(1))
    End If
    row.range.Item(3).Value = parseData(2)
    If (row.range.Item(3).ColumnWidth < Len(parseData(2))) Then
        row.range.Item(3).ColumnWidth = Len(parseData(2))
    End If
    row.range.Item(4).Value = parseData(3)
    If (row.range.Item(4).ColumnWidth < Len(parseData(3))) Then
        row.range.Item(4).ColumnWidth = Len(parseData(3))
    End If
End Sub

Sub updateRow(row)
    Dim parseData
    
    parseData = Parsers.SendParser(row.range.Item(4), row.range.Item(3))
    
    row.range.Item(1).Value = parseData(0)
    row.range.Item(2).Value = parseData(1)
    If (row.range.Item(2).ColumnWidth < Len(parseData(2))) Then
        row.range.Item(2).ColumnWidth = Len(parseData(2))
    End If
End Sub