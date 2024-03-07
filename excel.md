# VBA макрос для вставки и удаления строк

## Ссылки
https://learn.microsoft.com/ru-ru/office/vba/api/overview/excel  
[Работа со словарём](http://perfect-excel.ru/publ/excel/makrosy_i_programmy_vba/ischerpyvajushhee_opisanie_obekta_dictionary/7-1-0-101)  

```vba
Sub Add_data_groups_to_result()
' Добавить данные групп в таблицу итогов
        
  Application.ScreenUpdating = False
      
  Dim num_group As Long
  Dim num_row As Integer
  
  num_group = 0
  
  Do While num_group < 2
          
    num_group = num_group + 1
        
    ' Если текущие данные группы ещё не добавлены B1 <> "".
    Dim Sheet_gr As Worksheet
    Set Sheet_gr = Worksheets(num_group)
    
    If Sheet_gr.Range("B1").FormulaR1C1 <> "" Then
        
      ' Копируем данные группы за день
      Sheet_gr.Range("A4:F30").Copy
      
      ' Считаем итоги аванса
      Total_prepayment Sheet_gr
              
      ' Последняя строка в таблице итогов.
      Dim ws As Worksheet
      Dim lastRow As Long
      
      ' Вставляем данные в лист Итоги
      Dim range_last As Range
      Set ws = ActiveWorkbook.Sheets("Итоги")
      lastRow = ws.Range("A" & ws.Rows.Count).End(xlUp).Row + 1
      
      'ws.AutoFilterMode = False
      
      Set range_last = Range("A" & lastRow)
      range_last.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=True, Transpose:=False
      
      Application.CutCopyMode = False
            
      ' Удаление строк с нулевой суммой
      num_row = lastRow + 30 - 1
      Do While num_row >= lastRow
        If Range("E" & num_row).value = "" And Range("F" & num_row).value = "" Then
          Range("E" & num_row).EntireRow.Delete
        End If
        num_row = num_row - 1
      Loop
                
      ' Метка о записи данных в таблицу
      Sheet_gr.Range("B1").FormulaR1C1 = ""
          
      ' Обнуляем сумму в группе
      Sheet_gr.Range("E4:F30").ClearContents
          
    End If
          
  Loop
  
  Range("A1").Select
  
  Application.ScreenUpdating = True

End Sub
Sub Total_prepayment(Sheet_gr As Worksheet)
' Расчёт остатка аванса

  Dim num_row As Integer
  
  Dim sum, prep, sum_prep  As Double
    
  For num_row = 4 To 30
    
    sum = 0
    prep = 0
    sum_prep = 0
    
    sum = Sheet_gr.Range("E" & num_row).value
    prep = Sheet_gr.Range("F" & num_row).value
    sum_prep = Sheet_gr.Range("G" & num_row).value
    
    sum_prep = sum_prep + prep - sum
    
    If sum_prep < 0 Then
      sum_prep = 0
    End If
    
    Sheet_gr.Range("G" & num_row).value = sum_prep
    
  Next num_row

End Sub
```
![image](https://github.com/alzoi/Excel/assets/20499566/e5a7decc-ee68-4d93-a554-7f8ac56c5308)

