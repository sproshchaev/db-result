Attribute VB_Name = "Module_Динамика"
' *** Лист Динамика ***

' *** Глобальные переменные ***
' Public numStr_Лист8 As Integer
' ***                       ***

' Динамика продаж по годам, кварталам, месяцам
Sub Динамика_продаж()
Attribute Динамика_продаж.VB_ProcData.VB_Invoke_Func = " \n14"
Dim число_расчетных_месяцев As Integer

  ' Запрос
  If MsgBox("Сформировать выбор данных?", vbYesNo) = vbYes Then
    
    ' Строка "Форма Dyn_1"
    row_Форма_Dyn_1 = rowByValue(ThisWorkbook.Name, "Динамика", "Форма Dyn_1", 100, 100)
    
    ' Очистка формы
    Call clearСontents2(ThisWorkbook.Name, "Динамика", "C9", "I23")
    
    ' Строка "Форма Dyn_2"
    row_Форма_Dyn_2 = rowByValue(ThisWorkbook.Name, "Динамика", "Форма Dyn_2", 100, 100)
    
    ' Очистка формы 2
    Call clearСontents2(ThisWorkbook.Name, "Динамика", "A" + CStr(row_Форма_Dyn_2 + 3), "P93")
    
    
    ' Открываем таблицы:
    ' Открываем BASE\Sales
    OpenBookInBase ("Sales_Office")
    
    ' Открываем BASE\Products
    OpenBookInBase ("Products")
        
    ' Номер продукта на Листе Динамика
    Номер_продукта_на_Листе_Динамика = 0
        
    Номер_заполненной_строки_Форма_Dyn_2 = row_Форма_Dyn_2 + 2
        
    ' Выборка данных
    rowCount = row_Форма_Dyn_1 + 3
    Do While (ThisWorkbook.Sheets("Динамика").Cells(rowCount, 2).Value <> "") And (rowCount < ThisWorkbook.Sheets("Динамика").Range("M6").Value)
    
      ' Номер продукта на Листе Динамика
      Номер_продукта_на_Листе_Динамика = Номер_продукта_на_Листе_Динамика + 1
      
      ' Ед.изм.
      ThisWorkbook.Sheets("Динамика").Cells(rowCount, 3).Value = Product_Name_to_Unit(ThisWorkbook.Sheets("Динамика").Cells(rowCount, 2).Value)
      ThisWorkbook.Sheets("Динамика").Cells(rowCount, 3).NumberFormat = "@"
      ThisWorkbook.Sheets("Динамика").Cells(rowCount, 3).HorizontalAlignment = xlCenter

      ' "Форма Dyn_2"
      Номер_заполненной_строки_Форма_Dyn_2 = Номер_заполненной_строки_Форма_Dyn_2 + 1
      '
      ThisWorkbook.Sheets("Динамика").Cells(Номер_заполненной_строки_Форма_Dyn_2, 1).Value = ThisWorkbook.Sheets("Динамика").Cells(rowCount, 1).Value
      ThisWorkbook.Sheets("Динамика").Cells(Номер_заполненной_строки_Форма_Dyn_2, 1).NumberFormat = "@"
      ThisWorkbook.Sheets("Динамика").Cells(Номер_заполненной_строки_Форма_Dyn_2, 1).HorizontalAlignment = xlCenter
      '
      ThisWorkbook.Sheets("Динамика").Cells(Номер_заполненной_строки_Форма_Dyn_2, 2).Value = ThisWorkbook.Sheets("Динамика").Cells(rowCount, 2).Value
      ThisWorkbook.Sheets("Динамика").Cells(Номер_заполненной_строки_Форма_Dyn_2, 2).NumberFormat = "@"
      ThisWorkbook.Sheets("Динамика").Cells(Номер_заполненной_строки_Форма_Dyn_2, 2).HorizontalAlignment = xlLeft
      '
      ThisWorkbook.Sheets("Динамика").Cells(Номер_заполненной_строки_Форма_Dyn_2, 3).Value = Product_Name_to_Unit(ThisWorkbook.Sheets("Динамика").Cells(rowCount, 2).Value)
      ThisWorkbook.Sheets("Динамика").Cells(Номер_заполненной_строки_Форма_Dyn_2, 3).NumberFormat = "@"
      ThisWorkbook.Sheets("Динамика").Cells(Номер_заполненной_строки_Форма_Dyn_2, 3).HorizontalAlignment = xlCenter
      
      
      ' Цикл по 3-м годам
      For i = 1 To 3
        ' Номера офисов от 1 до 5
        Select Case i
          Case 1 ' 2019
            curr_year = 2019
            column_Факт_year = 4
            число_расчетных_месяцев = 12
          Case 2 ' 2020
            curr_year = 2020
            column_Факт_year = 6
            число_расчетных_месяцев = 12
          Case 3 ' 2021
            curr_year = 2021
            column_Факт_year = 8
            число_расчетных_месяцев = ThisWorkbook.Sheets("Динамика").Range("Q6").Value ' 8
        End Select
        
        ' Заголовки
        If Year(Date) <> curr_year Then
          ThisWorkbook.Sheets("Динамика").Cells(row_Форма_Dyn_1 + 2, column_Факт_year).Value = "Факт '" + Mid(CStr(curr_year), 3, 2)
          ThisWorkbook.Sheets("Динамика").Cells(row_Форма_Dyn_1 + 2, column_Факт_year + 1).Value = "Исп.плана '" + Mid(CStr(curr_year), 3, 2)
        Else
          
          ThisWorkbook.Sheets("Динамика").Cells(row_Форма_Dyn_1 + 2, column_Факт_year).Value = "Факт на " + strDDMM(Date_last_day_month(CDate("01." + CStr(число_расчетных_месяцев) + "." + CStr(Year(Date))))) ' ThisWorkbook.Sheets("Лист8").Range("F9").Value
          
          ' Если месяц равен число_расчетных_месяцев, то дату факта берем с Лист8
          If Month(CDate(Mid(ThisWorkbook.Sheets("Лист8").Range("B5").Value, 52, 10))) = число_расчетных_месяцев Then
            ThisWorkbook.Sheets("Динамика").Cells(row_Форма_Dyn_1 + 2, column_Факт_year).Value = ThisWorkbook.Sheets("Лист8").Range("F9").Value
          End If
          
        End If
        
        ' Рисуем код продукта и год во второй таблице
        Номер_заполненной_строки_Форма_Dyn_2 = Номер_заполненной_строки_Форма_Dyn_2 + 1
        ThisWorkbook.Sheets("Динамика").Cells(Номер_заполненной_строки_Форма_Dyn_2, 2).Value = Product_Name_to_Product_Code(ThisWorkbook.Sheets("Динамика").Cells(rowCount, 2).Value) + "_" + CStr(curr_year)
        ThisWorkbook.Sheets("Динамика").Cells(Номер_заполненной_строки_Форма_Dyn_2, 2).NumberFormat = "@"
        ThisWorkbook.Sheets("Динамика").Cells(Номер_заполненной_строки_Форма_Dyn_2, 2).HorizontalAlignment = xlRight

        
        ' Обработка по месяцам года
        Факт_года_РОО = 0
        План_года_РОО = 0
        
        curr_month = 1
        Do While curr_month <= число_расчетных_месяцев ' 12
        
          ' Итого факт за месяц
          Факт_месяца_РОО = 0
          План_месяца_РОО = 0
          For office_number = 1 To 5
            
            ' Берем показатель из столбца 2
            In_Product_Code = Product_Name_to_Product_Code(ThisWorkbook.Sheets("Динамика").Cells(rowCount, 2).Value)
           
            ' Определяем Факт месяца
            date_Var = CDate("01." + CStr(curr_month) + "." + CStr(curr_year))
            Факт_М_Var = Факт_М(date_Var, office_number, In_Product_Code)
            Факт_месяца_РОО = Факт_месяца_РОО + Факт_М_Var
            Факт_года_РОО = Факт_года_РОО + Факт_М_Var
            ' План месяца - План_М
            План_М_Var = План_М(date_Var, office_number, In_Product_Code)
            План_месяца_РОО = План_месяца_РОО + План_М_Var
            План_года_РОО = План_года_РОО + План_М_Var
            
          Next office_number
          
          ' Итоги по 5-ти офисам за месяц вставляем
          ' Номер_заполненной_строки_Форма_Dyn_2 = Номер_заполненной_строки_Форма_Dyn_2 + 1
          ThisWorkbook.Sheets("Динамика").Cells(Номер_заполненной_строки_Форма_Dyn_2, 3 + curr_month).Value = Факт_месяца_РОО
          ThisWorkbook.Sheets("Динамика").Cells(Номер_заполненной_строки_Форма_Dyn_2, 3 + curr_month).NumberFormat = "#,##0"
          ThisWorkbook.Sheets("Динамика").Cells(Номер_заполненной_строки_Форма_Dyn_2, 3 + curr_month).HorizontalAlignment = xlRight

          
          ' Следующая запись
          Application.StatusBar = "Обработано: " + ThisWorkbook.Sheets("Динамика").Cells(rowCount, 2).Value + "..."
          curr_month = curr_month + 1
          DoEventsInterval (rowCount)
    
        Loop ' Обработка с 1 по 12 месяцев года
        
        
        ' Вывод итогов по году
        ' Факт
        ThisWorkbook.Sheets("Динамика").Cells(rowCount, column_Факт_year).Value = Факт_года_РОО
        ThisWorkbook.Sheets("Динамика").Cells(rowCount, column_Факт_year).NumberFormat = "#,##0"
        ThisWorkbook.Sheets("Динамика").Cells(rowCount, column_Факт_year).HorizontalAlignment = xlRight
        
        ' Если это Штат, то делим на число_расчетных_месяцев
        If InStr(ThisWorkbook.Sheets("Динамика").Cells(rowCount, 2).Value, "Штат") <> 0 Then
          ThisWorkbook.Sheets("Динамика").Cells(rowCount, column_Факт_year).Value = ThisWorkbook.Sheets("Динамика").Cells(rowCount, column_Факт_year).Value / число_расчетных_месяцев
        End If
        
        '  Исполнение плана
        ThisWorkbook.Sheets("Динамика").Cells(rowCount, column_Факт_year + 1).Value = РассчетДоли(План_года_РОО, Факт_года_РОО, 3)
        ThisWorkbook.Sheets("Динамика").Cells(rowCount, column_Факт_year + 1).NumberFormat = "0%"
        ThisWorkbook.Sheets("Динамика").Cells(rowCount, column_Факт_year + 1).HorizontalAlignment = xlRight
        ' Окраска ячейки СФЕТОФОР ' 70. Заливка ячейки цветом "светофор" - значение In_Value в %, In_Target в %
        Call Full_Color_RangeII("Динамика", rowCount, column_Факт_year + 1, ThisWorkbook.Sheets("Динамика").Cells(rowCount, column_Факт_year + 1).Value, 1)
        
        ' Итоги года во второй таблице
        Column_Итого = 16
        ThisWorkbook.Sheets("Динамика").Cells(Номер_заполненной_строки_Форма_Dyn_2, Column_Итого).Value = Факт_года_РОО
        ThisWorkbook.Sheets("Динамика").Cells(Номер_заполненной_строки_Форма_Dyn_2, Column_Итого).NumberFormat = "#,##0"
        ThisWorkbook.Sheets("Динамика").Cells(Номер_заполненной_строки_Форма_Dyn_2, Column_Итого).HorizontalAlignment = xlRight
        
      Next i
        
      
      ' Следующая запись
      Application.StatusBar = "Обработано: " + ThisWorkbook.Sheets("Динамика").Cells(rowCount, 2).Value + "..."
      rowCount = rowCount + 1
      DoEventsInterval (rowCount)
    
    Loop

    
    ' Закрываем таблицы:
    ' Закрываем BASE\Products
    CloseBook ("Products")
    
    ' Закрываем BASE\Sales
    CloseBook ("Sales_Office")

    
    
    MsgBox ("Выборка сформирована!")
  
  End If

  
  
End Sub

' Лист Динамика
Sub SaveFromЛист_Динамика()

  ' Копируем Лист2
  ThisWorkbook.Sheets("Динамика").Copy

  '
  ' Workbooks("Книга1").Sheets("Лист1").Paste

End Sub


' Создать график по данным из таблицы
Sub Создать_график_Лист_Динамика()

    ' Выбор диапазона с данныйми по Y Range("C8:I9").Select
    ThisWorkbook.Sheets("Charts").Range("C9:I9").Select
    
    ' Добавление графика
    ActiveSheet.Shapes.AddChart2(332, xlLineMarkers, 1000, 150).Select
    
    ' ActiveChart.SetSourceData Source:=Range("Графики!$C$8:$I$9")
    ActiveChart.SetSourceData Source:=ThisWorkbook.Sheets("Charts").Range("Charts!$C$8:$I$9")
    
    ActiveChart.ChartTitle.Select
    ActiveChart.Axes(xlValue).MajorGridlines.Select
    
    ' Надпись наименования графика
    ActiveChart.ChartTitle.Text = "Выдача ПК"
    ActiveChart.ChartTitle.Format.TextFrame2.TextRange.Characters.Text = "Выдача ПК"
    
    ' Наименование первого ряда
    ActiveChart.FullSeriesCollection(1).Name = "=""Ряд_1"""
    
    ' Добавление второго ряда
    ActiveChart.PlotArea.Select
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.FullSeriesCollection(2).Name = "=""Ряд_2"""
    ActiveChart.FullSeriesCollection(2).Values = "=Charts!$C$10:$I$10"
    
    ' Добавление третьего ряда - индекс ряда нужно увеличить FullSeriesCollection(2, затем 3 и т.д.)
    ActiveChart.PlotArea.Select
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.FullSeriesCollection(3).Name = "=""Ряд_3"""
    ActiveChart.FullSeriesCollection(3).Values = "=Charts!$C$11:$I$11"
    
    ' Добавление легенды
    ' ActiveSheet.ChartObjects("Диаграмма 23").Activate
    ActiveChart.SetElement (msoElementLegendRight)
    ' ActiveSheet.ChartObjects("Диаграмма 23").Activate
    ' ActiveChart.Legend.Select
    ' ActiveChart.Legend.LegendEntries(1).Select
    ' Application.CommandBars("Format Object").Visible = False

    
End Sub