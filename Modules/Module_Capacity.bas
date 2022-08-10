Attribute VB_Name = "Module_Capacity"
' Лист "Capacity"

' Шаблон_обработки_5_офисов
Sub Отчет_Capacity_New()
  
' Описание переменных
Dim ReportName_String, officeNameInReport, CheckFormatReportResult As String
Dim i, rowCount As Integer
Dim finishProcess As Boolean
    
  ' Открыть файл с отчетом
  FileName = Application.GetOpenFilename("Excel Files (*.xlsx), *.xlsx", , "Открытие файла с отчетом")

  ' Если файл был выбран
  If (Len(FileName) > 5) Then
  
    ' Строка статуса
    Application.StatusBar = "Обработка отчета..."
  
    ' Переменная начала обработки
    finishProcess = False

    ' Выводим для инфо данные об имени файла
    ReportName_String = Dir(FileName)
  
    ' Строка статуса
    Application.StatusBar = "Открытие файла Capacity..."
  
    ' Открываем выбранную книгу (UpdateLinks:=0)
    Workbooks.Open FileName, 0
      
    ' Переходим на окно DB
    ThisWorkbook.Sheets("Capacity").Activate

    ' Проверка формы отчета
    CheckFormatReportResult = CheckFormatReport(ReportName_String, "Тепловая карта", 19, Date)
    
    If CheckFormatReportResult = "OK" Then
      
      ' Выставляем фильтры на вкладке "Продажи"
      ' Call Открытие_сводных_Capacity_New_Продажи(ReportName_String, "Продажи")
      
      ' Клиенты (кросс)
      Application.StatusBar = "Открытие Клиенты (кросс)..."
      Call Открытие_сводных_Capacity_New_Клиенты_кросс(ReportName_String, "Клиенты (кросс)")
      
      ' PA ПК
      Application.StatusBar = "Открытие PA ПК..."
      Call Открытие_сводных_Capacity_New_Клиенты_кросс(ReportName_String, "PA ПК")
      
      ' PA KK
      Application.StatusBar = "Открытие PA KK..."
      Call Открытие_сводных_Capacity_New_Клиенты_кросс(ReportName_String, "PA KK")
      
      ' ДК и Пенс
      Application.StatusBar = "Открытие ДК и Пенс..."
      Call Открытие_сводных_Capacity_New_Клиенты_кросс(ReportName_String, "ДК и Пенс")
      
      ' Строка статуса
      Application.StatusBar = "Определение столбцов..."
      
      row_ЛистCapacity_Форма61 = rowByValue(ThisWorkbook.Name, "Capacity", "Форма 6.1", 100, 100)
      
      ' Определение столбцов на "Клиенты (кросс)"
      ' Клиентов
      column_Capacity_КлиентыКросс_Клиентов = ColumnByValue(ReportName_String, "Клиенты (кросс)", "Клиентов", 300, 300)
      ' Заказчиков КК
      column_Capacity_КлиентыКросс_ЗаказчиковКК = ColumnByValue(ReportName_String, "Клиенты (кросс)", "Заказчиков КК", 300, 300)
      ' Доля Заказчиков КК
      column_Capacity_КлиентыКросс_ДоляЗаказчиковКК = ColumnByValue(ReportName_String, "Клиенты (кросс)", "Доля Заказчиков КК", 300, 300)
      ' Заявок на кредит
      column_Capacity_КлиентыКросс_ЗаявокНаКредит = ColumnByValue(ReportName_String, "Клиенты (кросс)", "Заявок на кредит", 300, 300)
      ' Доля заявок на кредит
      column_Capacity_КлиентыКросс_ДоляЗаявокНаКредит = ColumnByValue(ReportName_String, "Клиенты (кросс)", "Доля заявок на кредит", 300, 300)
            
               
            
      ' Обрабатываем отчет
      ' Цикл по 5-ти офисам
      ' Обработка отчета
      For i = 1 To 5
        ' Номера офисов от 1 до 5
        Select Case i
          Case 1 ' ОО «Тюменский»
            officeNameInReport = "Тюменский"
          Case 2 ' ОО «Сургутский»
            officeNameInReport = "Сургутский"
          Case 3 ' ОО «Нижневартовский»
            officeNameInReport = "Нижневартовский"
          Case 4 ' ОО «Новоуренгойский»
            officeNameInReport = "Новоуренгойский"
          Case 5 ' ОО «Тарко-Сале»
            officeNameInReport = "Тарко-Сале"
        End Select

        ' 1) Клиенты (кросс)
        rowCount = 1
        Do While InStr(Workbooks(ReportName_String).Sheets("Клиенты (кросс)").Cells(rowCount, 1).Value, "Общий итог") = 0
        
          ' Если это текущий офис
          If (InStr(Workbooks(ReportName_String).Sheets("Клиенты (кросс)").Cells(rowCount, 1).Value, officeNameInReport) <> 0) And (InStr(Workbooks(ReportName_String).Sheets("Клиенты (кросс)").Cells(rowCount, 1).Value, "ОО") <> 0) Then
            
            ' Вставляем row_ЛистCapacity_Форма61
            ' №
            ThisWorkbook.Sheets("Capacity").Cells(row_ЛистCapacity_Форма61 + 2 + i, 1).Value = CStr(i)
            ThisWorkbook.Sheets("Capacity").Cells(row_ЛистCapacity_Форма61 + 2 + i, 1).NumberFormat = "#,##0"
            ThisWorkbook.Sheets("Capacity").Cells(row_ЛистCapacity_Форма61 + 2 + i, 1).HorizontalAlignment = xlCenter

            ' Офис
            ThisWorkbook.Sheets("Capacity").Cells(row_ЛистCapacity_Форма61 + 2 + i, 2).Value = getNameOfficeByNumber(i)
            ThisWorkbook.Sheets("Capacity").Cells(row_ЛистCapacity_Форма61 + 2 + i, 2).NumberFormat = "@"
            ThisWorkbook.Sheets("Capacity").Cells(row_ЛистCapacity_Форма61 + 2 + i, 2).HorizontalAlignment = xlLeft
            
            ' Клиенты
            ThisWorkbook.Sheets("Capacity").Cells(row_ЛистCapacity_Форма61 + 2 + i, 3).Value = Workbooks(ReportName_String).Sheets("Клиенты (кросс)").Cells(rowCount, column_Capacity_КлиентыКросс_Клиентов).Value
            ThisWorkbook.Sheets("Capacity").Cells(row_ЛистCapacity_Форма61 + 2 + i, 3).NumberFormat = "#,##0"
            ThisWorkbook.Sheets("Capacity").Cells(row_ЛистCapacity_Форма61 + 2 + i, 3).HorizontalAlignment = xlRight
            
            ' Заявки КК
            ThisWorkbook.Sheets("Capacity").Cells(row_ЛистCapacity_Форма61 + 2 + i, 4).Value = Workbooks(ReportName_String).Sheets("Клиенты (кросс)").Cells(rowCount, column_Capacity_КлиентыКросс_ЗаказчиковКК).Value
            ThisWorkbook.Sheets("Capacity").Cells(row_ЛистCapacity_Форма61 + 2 + i, 4).NumberFormat = "#,##0"
            ThisWorkbook.Sheets("Capacity").Cells(row_ЛистCapacity_Форма61 + 2 + i, 4).HorizontalAlignment = xlRight
            
            ' Доля
            ThisWorkbook.Sheets("Capacity").Cells(row_ЛистCapacity_Форма61 + 2 + i, 5).Value = Workbooks(ReportName_String).Sheets("Клиенты (кросс)").Cells(rowCount, column_Capacity_КлиентыКросс_ДоляЗаказчиковКК).Value
            ThisWorkbook.Sheets("Capacity").Cells(row_ЛистCapacity_Форма61 + 2 + i, 5).NumberFormat = "0%"
            ThisWorkbook.Sheets("Capacity").Cells(row_ЛистCapacity_Форма61 + 2 + i, 5).HorizontalAlignment = xlRight

            
                
          End If
        
        
          ' Следующая запись
          rowCount = rowCount + 1
          Application.StatusBar = officeNameInReport + ": " + CStr(rowCount) + "..."
          DoEventsInterval (rowCount)
        Loop
   
        ' Выводим данные по офису
      
        ' 2) ДК и Пенс
        ' rowCount = 1
        ' Do While InStr(Workbooks(ReportName_String).Sheets("ДК и Пенс").Cells(rowCount, 1).Value, "Общий итог") = 0
        
          ' Если это текущий офис
        '  If (InStr(Workbooks(ReportName_String).Sheets("ДК и Пенс").Cells(rowCount, 1).Value, officeNameInReport) <> 0) And (InStr(Workbooks(ReportName_String).Sheets("ДК и Пенс").Cells(rowCount, 1).Value, "ОО") <> 0) Then
            
            ' Вставляем row_ЛистCapacity_Форма61
            ' №
        '    ThisWorkbook.Sheets("Capacity").Cells(row_ЛистCapacity_Форма61 + 2 + i, 1).Value = CStr(i)
        '    ThisWorkbook.Sheets("Capacity").Cells(row_ЛистCapacity_Форма61 + 2 + i, 1).NumberFormat = "#,##0"
        '    ThisWorkbook.Sheets("Capacity").Cells(row_ЛистCapacity_Форма61 + 2 + i, 1).HorizontalAlignment = xlCenter

            ' Офис
        '    ThisWorkbook.Sheets("Capacity").Cells(row_ЛистCapacity_Форма61 + 2 + i, 2).Value = getNameOfficeByNumber(i)
        '    ThisWorkbook.Sheets("Capacity").Cells(row_ЛистCapacity_Форма61 + 2 + i, 2).NumberFormat = "@"
        '    ThisWorkbook.Sheets("Capacity").Cells(row_ЛистCapacity_Форма61 + 2 + i, 2).HorizontalAlignment = xlLeft
            
            ' Клиенты
        '    ThisWorkbook.Sheets("Capacity").Cells(row_ЛистCapacity_Форма61 + 2 + i, 3).Value = Workbooks(ReportName_String).Sheets("Клиенты (кросс)").Cells(rowCount, column_Capacity_КлиентыКросс_Клиентов).Value
        '    ThisWorkbook.Sheets("Capacity").Cells(row_ЛистCapacity_Форма61 + 2 + i, 3).NumberFormat = "#,##0"
        '    ThisWorkbook.Sheets("Capacity").Cells(row_ЛистCapacity_Форма61 + 2 + i, 3).HorizontalAlignment = xlRight
            
            ' Заявки КК
         '   ThisWorkbook.Sheets("Capacity").Cells(row_ЛистCapacity_Форма61 + 2 + i, 4).Value = Workbooks(ReportName_String).Sheets("Клиенты (кросс)").Cells(rowCount, column_Capacity_КлиентыКросс_ЗаказчиковКК).Value
         '   ThisWorkbook.Sheets("Capacity").Cells(row_ЛистCapacity_Форма61 + 2 + i, 4).NumberFormat = "#,##0"
         '   ThisWorkbook.Sheets("Capacity").Cells(row_ЛистCapacity_Форма61 + 2 + i, 4).HorizontalAlignment = xlRight
            
            ' Доля
         '   ThisWorkbook.Sheets("Capacity").Cells(row_ЛистCapacity_Форма61 + 2 + i, 5).Value = Workbooks(ReportName_String).Sheets("Клиенты (кросс)").Cells(rowCount, column_Capacity_КлиентыКросс_ДоляЗаказчиковКК).Value
         '   ThisWorkbook.Sheets("Capacity").Cells(row_ЛистCapacity_Форма61 + 2 + i, 5).NumberFormat = "0%"
         '   ThisWorkbook.Sheets("Capacity").Cells(row_ЛистCapacity_Форма61 + 2 + i, 5).HorizontalAlignment = xlRight

            
                
         ' End If
        
        
          ' Следующая запись
         ' rowCount = rowCount + 1
        '  Application.StatusBar = officeNameInReport + ": " + CStr(rowCount) + "..."
       '   DoEventsInterval (rowCount)
      '  Loop
   
        ' Выводим данные по офису
      
      
     Next i ' Следующий офис
      
      ' Выводим итоги обработки
      
      ' Сохранение изменений
      ThisWorkbook.Save
    
      ' Переменная завершения обработки
      finishProcess = True
    Else
      ' Сообщение о неверном формате отчета или даты
      MsgBox ("Проверьте отчет: " + CheckFormatReportResult + "!")
    End If ' Проверка формы отчета

    ' Закрываем файл с отчетом без сохранения изменений (параметр SaveChanges:=False)
    ' Workbooks(Dir(FileName)).Close SaveChanges:=False - отладка
    
    ' Переходим в ячейку M2
    ThisWorkbook.Sheets("Capacity").Range("A1").Select

    ' Строка статуса
    Application.StatusBar = ""

    ' Зачеркиваем пункт меню на стартовой страницы
    ' Call ЗачеркиваемТекстВячейке("Лист0", "D9")
    ' Call ЗачеркиваемТекстВячейке("Лист0", RangeByValue(ThisWorkbook.Name, "Лист0", "Оперативная справка по _________________", 100, 100))
    
    ' Итоговое сообщение
    If finishProcess = True Then
      MsgBox ("Обработка " + Dir(ReportName_String) + " завершена!")
    Else
      MsgBox ("Обработка отчета была прервана!")
    End If

  End If ' Если файл был выбран

End Sub

' Открытие сводных таблиц в Capacity на вкладке "Клиенты (кросс)"
Sub Открытие_сводных_Capacity_New_Клиенты_кросс(In_ReportName_String, In_Sheet)
      
    ' Range("A11").Select
    Workbooks(In_ReportName_String).Sheets(In_Sheet).PivotTables("Сводная таблица2").PivotFields("[Куратор].[Куратор].[Куратор]").PivotItems("[Куратор].[Куратор].&[Данилов Александр Сергеевич]").DrilledDown = True
    
    ' Range("A17").Select
    Workbooks(In_ReportName_String).Sheets(In_Sheet).PivotTables("Сводная таблица2").PivotFields("[Филиал].[Регион].[Регион]").PivotItems("[Филиал].[Регион].&[Тюменский]").DrilledDown = True

End Sub

' Открытие сводных таблиц в Capacity на вкладке "PRE-APPROVED"
Sub Открытие_сводных_Capacity_New_PREAPPROVED(In_ReportName_String, In_Sheet)

    ' Добавление ClientIdRetail
    Workbooks(In_ReportName_String).Sheets(In_Sheet).PivotTables("Сводная таблица2").CubeFields ("[Клиент].[ClientIdRetail]")
    Workbooks(In_ReportName_String).Sheets(In_Sheet).PivotTables("Сводная таблица2").Orientation = xlRowField
    Workbooks(In_ReportName_String).Sheets(In_Sheet).PivotTables("Сводная таблица2").Position = 4
    
    ' Открываем Тюменский РОО
    Workbooks(In_ReportName_String).Sheets(In_Sheet).PivotTables("Сводная таблица2").PivotFields("[Филиал].[Регион].[Регион]").VisibleItemsList = Array("[Филиал].[Регион].&[Тюменский]")
    Workbooks(In_ReportName_String).Sheets(In_Sheet).PivotTables("Сводная таблица2").PivotFields("[Филиал].[Регион].[Регион]").PivotItems("[Филиал].[Регион].&[Тюменский]").DrilledDown = True
    
    ' Открытие 5 офисов
    Workbooks(In_ReportName_String).Sheets(In_Sheet).PivotTables("Сводная таблица2").PivotFields("[Филиал].[ДО].[ДО]").PivotItems("[Филиал].[ДО].&[ОО ""Нижневартовский""]").DrilledDown = True
    Workbooks(In_ReportName_String).Sheets(In_Sheet).PivotTables("Сводная таблица2").PivotFields("[Филиал].[ДО].[ДО]").PivotItems("[Филиал].[ДО].&[ОО ""Новоуренгойский""]").DrilledDown = True
    Workbooks(In_ReportName_String).Sheets(In_Sheet).PivotTables("Сводная таблица2").PivotFields("[Филиал].[ДО].[ДО]").PivotItems("[Филиал].[ДО].&[ОО ""Сургутский""]").DrilledDown = True
    Workbooks(In_ReportName_String).Sheets(In_Sheet).PivotTables("Сводная таблица2").PivotFields("[Филиал].[ДО].[ДО]").PivotItems("[Филиал].[ДО].&[ОО ""Тарко-Сале"" Уральского филиала ПАО ""Промсвязьбанк""]").DrilledDown = True
    Workbooks(In_ReportName_String).Sheets(In_Sheet).PivotTables("Сводная таблица2").PivotFields("[Филиал].[ДО].[ДО]").PivotItems("[Филиал].[ДО].&[ОО ""Тюменский""]").DrilledDown = True
    
End Sub

' Открытие сводных таблиц в Capacity на вкладке "Продажи"
Sub Открытие_сводных_Capacity_New_Продажи(In_ReportName_String, In_Sheet)
  
  ' Workbooks(In_ReportName_String).Sheets(In_Sheet).PivotTables("Сводная таблица2").PivotFields("[Филиал].[Регион].[Регион]").PivotItems("[Филиал].[Регион].&[Ивановский]").DrilledDown = True
  ' Workbooks(In_ReportName_String).Sheets(In_Sheet).PivotTables("Сводная таблица2").PivotFields("[Филиал].[Регион].[Регион]").PivotItems("[Филиал].[Регион].&[Владимирский]").DrilledDown = True

    Workbooks(In_ReportName_String).Sheets(In_Sheet).PivotTables("Сводная таблица2").PivotFields("[Куратор].[Куратор].[Куратор]").PivotItems("[Куратор].[Куратор].&[Данилов Александр Сергеевич]").DrilledDown = True
    Workbooks(In_ReportName_String).Sheets(In_Sheet).PivotTables("Сводная таблица2").PivotFields("[Филиал].[Регион].[Регион]").PivotItems("[Филиал].[Регион].&[Тюменский]").DrilledDown = True


End Sub


' Выставляем фильтры на вкладке "Продажи"
Sub setFilter_Capacity_Продажи()
  
  
  
End Sub

