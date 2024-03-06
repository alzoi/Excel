# VBA макрос для вставки и удаления строк

## Ссылки
https://learn.microsoft.com/ru-ru/office/vba/api/overview/excel

```vba
' Добавить данные групп в таблицу итогов
Sub Add_data_groups_to_result()
        
  Application.ScreenUpdating = False
      
  Dim num_group As Long
  
  num_group = 0
  
  Do While num_group < 2
          
    num_group = num_group + 1
        
    ' Если текущие данные группы ещё не добавлены.
    Dim Sheet_gr As Worksheet
    Set Sheet_gr = ActiveWorkbook.Worksheets(num_group)
    
    If Sheet_gr.Range("B1").FormulaR1C1 <> "" Then
        
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
    
    End If
          
  Loop
  
  Range("A1").Select
  
  Application.ScreenUpdating = True

End Sub
```
![image](https://github.com/alzoi/Excel/assets/20499566/e5a7decc-ee68-4d93-a554-7f8ac56c5308)

