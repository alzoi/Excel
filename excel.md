# VBA макрос для вставки и удаления строк
```vba
    Application.ScreenUpdating = False
               
    ' Запрещаем повторное добавление.
    Dim Sheet_gr As Worksheet
    Set Sheet_gr = ActiveWorkbook.Sheets("Группа1")
    
    If Sheet_gr.Range("B1").FormulaR1C1 = "" Then
        Application.ScreenUpdating = True
        Exit Sub
    End If
        
    ' Копируем данные группы за день
    Sheet_gr.Range("A4").CurrentRegion.Copy
            
    ' Последняя строка в таблице итогов.
    Dim ws As Worksheet
    Dim lastRow As Long
    
    ' Вставляем данные в лист Итоги
    Dim range_last As Range
    Set ws = ActiveWorkbook.Sheets("Итоги")
    lastRow = ws.Range("A" & ws.Rows.Count).End(xlUp).Row + 1
    
    Set range_last = Range("A" & lastRow)
    range_last.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=True, Transpose:=False
    
    Application.CutCopyMode = False

    ' Удаление строк в которых в ячейке D "Сумма" нет значения
    Dim rng As Range
    lastRow = ws.Range("D" & ws.Rows.Count).End(xlUp).Row
    Set rng = ws.Range("D1:D" & lastRow)
    
    ' Фильтруем и удаляем ячейки
    With rng
        .AutoFilter Field:=1, Criteria1:="="
        .Offset(1, 0).SpecialCells(xlCellTypeVisible).EntireRow.Delete
    End With

    ws.AutoFilterMode = False
        
    ' Метка о записи данных в таблицу
    Sheet_gr.Range("B1").FormulaR1C1 = ""
        
    ' Обнуляем сумму
    Sheet_gr.Range("D4:D30").ClearContents

    'Dim row_num As Integer
    'row_num = 3
    'Do While row_num < 30
    '    row_num = row_num + 1
    '    Range("C" & row_num).Select
    '    ActiveCell.FormulaR1C1 = ""
    'Loop
    
    'Range("A1").Select
    
    Application.ScreenUpdating = True
```
![image](https://github.com/alzoi/Excel/assets/20499566/e5a7decc-ee68-4d93-a554-7f8ac56c5308)

