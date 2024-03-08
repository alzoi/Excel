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

# Журнал
```vba
Dim НачалоДанных, ОкончаниеДанных As Integer

Sub ВыполнитьПроводкуВсеГруппы()

  Application.ScreenUpdating = False
  
  НачалоДанных = 4
  ОкончаниеДанных = 30
  
  Dim НомерГруппы As Long
  Dim Ошибка As Boolean
  
  ' Список групп
  For НомерГруппы = 1 To 2
    
    Dim ЛистГруппы As Worksheet
    Set ЛистГруппы = Worksheets(НомерГруппы)
        
    ' Если текущие данные группы ещё не добавлены B1 <> "".
    If Ошибка = False And ЛистГруппы.range("B1").FormulaR1C1 <> "" Then
          
      Ошибка = ВыполнитьПроводкуГруппы(ЛистГруппы)
    
    End If
      
  Next НомерГруппы
    
  If Ошибка = False Then
    range("A1").Select
  End If
  
  Application.ScreenUpdating = True

End Sub
Function ВыполнитьПроводкуГруппы(ЛистГруппы As Worksheet) As Boolean

  Dim Ошибка As Boolean
  Dim НомерСтроки As Integer
  Dim ИсходныйОстатокАванса As Double
  
  Dim Сумма, Аванс, ОстатокАванса As range
  
  ' Получаем новый остаток Аванса
  For НомерСтроки = НачалоДанных To ОкончаниеДанных
    
    Set Аванс = ЛистГруппы.range("F" & НомерСтроки)
    
    Set ОстатокАванса = ЛистГруппы.range("G" & НомерСтроки)
    
    ИсходныйОстатокАванса = ОстатокАванса.value
    
    If Аванс.value > 0 Then
                              
      ОстатокАванса.value = ИсходныйОстатокАванса + Аванс.value
      
    End If
    
    Set Сумма = ЛистГруппы.range("E" & НомерСтроки)
    
    ' Если сумму нельзя списать из аванса
    If Сумма.value < 0 _
    And ОстатокАванса.value > 0 And Abs(Сумма.value) > ОстатокАванса.value Then
      
      ОстатокАванса.value = ИсходныйОстатокАванса
      
      ЛистГруппы.Activate
      
      ЛистГруппы.range("E" & НомерСтроки).Select
      
      Dim Response
      Response = MsgBox("Аванса недостаточно, лист ''" & ЛистГруппы.Name & _
        "'', строка " & НомерСтроки, vbExclamation)
      
      ВыполнитьПроводкуГруппы = True
      Exit Function
            
    End If
    
    ' Если сумму можно списать из аванса
    If Сумма.value < 0 _
    And ОстатокАванса.value > 0 And Abs(Сумма.value) <= ОстатокАванса.value Then
      ОстатокАванса.value = ОстатокАванса.value + Сумма.value
    End If
    
  Next НомерСтроки
  
  ' Копируем данные группы за день
  ЛистГруппы.range("A" & НачалоДанных & ":F" & ОкончаниеДанных).Copy
  
  ' Вставляем данные в лист Итоги
  Dim ЛистИтоги As Worksheet
  Set ЛистИтоги = ActiveWorkbook.Sheets("Итоги")
  
  Dim НоваяСтрока As Long
  НоваяСтрока = ПолучитьСтрокуВставки(ЛистИтоги)
        
  Dim range_last As range
  Set range_last = range("A" & НоваяСтрока)
  range_last.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=True, Transpose:=False
  
  Application.CutCopyMode = False
        
  ' Удаление строк с нулевой суммой
  Dim i As Integer
  For i = НоваяСтрока + ОкончаниеДанных - НачалоДанных To НоваяСтрока Step -1
    
    If range("E" & i).value = "" And range("F" & i).value = "" Then
      
      range("E" & i).EntireRow.Delete
    
    ElseIf range("E" & i).value < 0 Then
      
      ' Перенос суммы в использование аванса
      range("G" & i).value = range("E" & i).value
      range("E" & i).value = 0
    
    End If
  
  Next i
            
  ' Метка о записи данных в таблицу
  ЛистГруппы.range("B1").FormulaR1C1 = ""
      
  ' Обнуляем введённые данные
  ЛистГруппы.range("E" & НачалоДанных & ":F" & ОкончаниеДанных).ClearContents

End Function
Function ПолучитьСтрокуВставки(ws As Worksheet) As Long
    
  Dim ran As range
  Set ran = ws.range("A:A").Find("", LookIn:=xlValues)
  If Not ran Is Nothing Then
    ПолучитьСтрокуВставки = ran.Row
  End If

End Function
```

![image](https://github.com/alzoi/Excel/assets/20499566/7e251cdc-7593-48cc-bf85-0cb4e7cf92a8)  
![image](https://github.com/alzoi/Excel/assets/20499566/44d4a133-8ddd-4731-a7e4-35b42eb4ddc9)  



