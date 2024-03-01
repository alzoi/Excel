# VBA макрос для вставки и удаления строк
```vba
    '' Вставка строк в общую таблицу
    Sheets("Лист1").Select
    Range("A3").Select
    Selection.CurrentRegion.Select
    Selection.Copy
    Sheets("Лист2").Select
    Range("A2").Select
    Selection.Insert Shift:=xlDown
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Range("A1").Select
    
    '' Удаление строк в которых в ячейке B нет значения
    Dim ws As Worksheet
    Dim rng As Range
    Dim lastRow As Long

    Set ws = ActiveWorkbook.Sheets("Лист2")

    lastRow = ws.Range("B" & ws.Rows.Count).End(xlUp).Row

    Set rng = ws.Range("B1:B" & lastRow)
    
    ' Фильтруем и удаляем ячейки
    With rng
        .AutoFilter Field:=1, Criteria1:="="
        .Offset(1, 0).SpecialCells(xlCellTypeVisible).EntireRow.Delete
    End With

    ws.AutoFilterMode = False
```