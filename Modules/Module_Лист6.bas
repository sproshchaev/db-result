Attribute VB_Name = "Module_Лист6"
' Отчет Capacity
' Планы доработки:
' "Цикл №5",
' "Вставляем информацию о наличии PA ПК и КК в BASE\Clients: Capacity_PA_Date, Capacity_PA, Capacity_PA_CC_Date, Capacity_PA_CC"
' Выводить "Кол-во заявок на кредит" "Доля заявок на Кредит" в форму отчета
Sub Отчет_Capacity()
  
' Описание переменных
Dim ReportName_String, officeNameInReport, CheckFormatReportResult, allNameStr, currentNameStr_Range, cellSum As String
Dim i, rowCount, countClient, countCredCard, Клиенты, Клиенты_с_PA_КК, Выдано_PA, Заказ_PA_КК, с_ПР_PA_КК, Выдано_ПР_PA_КК, row_Тюменский_ОО1, Клиенты_с_PA, Клиенты_Без_ДК, Заявки_ДК, Количество_пенсионеров, Пенсионеры_Без_ДК, Пенсионеры_выдано_ДК, Пенсионеры_Заявл_ПФР As Integer
Dim rowForWriteInSheet6, currentNameStr_Row, currentNameStr_Column, row_Итого_по_РОО As Byte
Dim finishProcess As Boolean
' Dim  As Double
Dim dateReportCapacity As Date
        
  ' Открыть файл с отчетом
  FileName = Application.GetOpenFilename("Excel Files (*.xlsb), *.xlsb", , "Открытие файла с отчетом")

  ' Если файл был выбран
  If (Len(FileName) > 5) Then
  
    ' Строка статуса
    Application.StatusBar = "Обработка отчета..."
  
    ' Переменная начала обработки
    finishProcess = False

    ' Выводим для инфо данные об имени файла
    ReportName_String = Dir(FileName)
  
    ' Открываем выбранную книгу (UpdateLinks:=0)
    Workbooks.Open FileName, 0
      
    ' Переходим на окно DB
    ThisWorkbook.Sheets("Лист6").Activate
    ThisWorkbook.Sheets("Лист6").Range("A1").Select

    ' Проверка формы отчета
    CheckFormatReportResult = CheckFormatReport(ReportName_String, "Клиенты (Кросс)", 7, Date)
    
    ' Если формат отчета корректен, то запускаем обработку
    If CheckFormatReportResult = "OK" Then
      
      ' Открываем BASE\Clients
      OpenBookInBase ("Clients")
            
      ' Получение даты отчета Capacity. Лист "Тип операции". Строка "ДО - Менеджер - Тип операции - Категория операции - Операция".
      dateReportCapacity = dateReportFromCapacity(ReportName_String, "Тип операции")
      ' в B2 Jnxtn c __ по ___
      ThisWorkbook.Sheets("Лист6").Range("B2").Value = "Отчет кросс-продажи с " + strDDMM(monthStartDate(dateReportCapacity)) + " по " + CStr(dateReportCapacity) + " г."
      ' Неделя
      ThisWorkbook.Sheets("Лист6").Cells(rowByValue(ThisWorkbook.Name, "Лист6", "Неделя:", 100, 100), ColumnByValue(ThisWorkbook.Name, "Лист6", "Неделя:", 100, 100) + 1).Value = WeekNumber(dateReportCapacity)
      ' Тема:
      ThisWorkbook.Sheets("Лист6").Cells(rowByValue(ThisWorkbook.Name, "Лист6", "Тема:", 100, 100), ColumnByValue(ThisWorkbook.Name, "Лист6", "Тема:", 100, 100) + 1).Value = ThisWorkbook.Sheets("Лист6").Range("B2").Value
      ' Список получателей:
      ThisWorkbook.Sheets("Лист6").Cells(rowByValue(ThisWorkbook.Name, "Лист6", "Список получателей:", 100, 100), ColumnByValue(ThisWorkbook.Name, "Лист6", "Список получателей:", 100, 100) + 2).Value = getFromAddrBook("УДО2,УДО3,УДО4,УДО5,НОРПиКО1,НОРПиКО2,НОРПиКО3,НОРПиКО4,НОРПиКО5", 2)
      
      ' Показываем все скрытые ячейки по горизонтали
      ThisWorkbook.Sheets("Лист6").Rows("6:20").EntireRow.Hidden = False

      ' Очищаем ячейки отчета
      Call clearСontents2(ThisWorkbook.Name, "Лист6", "A6", "AI40")
      
      ' Цикл №1 прошивка КК на входящем потоке
            
      ' Переменные открытия списка сводной таблицы
      список_открыт_Клиенты_Кросс = False
      ' Номер строки для вывода на Листе 6 (как "Заявки КК на потоке" +2)
      rowForWriteInSheet6 = rowByValue(ThisWorkbook.Name, "Лист6", "Заявки КК на потоке", 100, 100) + 2
      ' Весь список сотрудников
      allNameStr = ""
      ' Сумма ячеек по офисам: С6+С11+...
      cellSum = ""
      
      ' Переходим на вкладку "Клиенты (Кросс)" и открываем все сводные таблицы
      Call openPivotTables_Capacity_Клиенты_Кросс(ReportName_String)
                
      ' Обрабатываем отчет
      ' Цикл по 5-ти офисам
      For i = 1 To 5
        
        ' Номера офисов от 1 до 5
        Select Case i
          Case 1 ' ОО «Тюменский» "ОО ""Тюменский"""
            officeNameInReport = "ОО ""Тюменский""" ' "Тюменский"
            officeNameInReport2 = "Тюменский"
          Case 2 ' ОО «Сургутский» "ОО2""Сургутский"""
            officeNameInReport = "ОО2""Сургутский""" ' "Сургутский"
            officeNameInReport2 = "Сургутский"
          Case 3 ' ОО «Нижневартовский» "ОО2 ""Нижневартовский"""
            officeNameInReport = "ОО2 ""Нижневартовский""" ' "Нижневартовский"
            officeNameInReport2 = "Нижневартовский"
          Case 4 ' ОО «Новоуренгойский» "ОО2""Новоуренгойский"""
            officeNameInReport = "ОО2""Новоуренгойский""" ' "Новоуренгойский"
            officeNameInReport2 = "Новоуренгойский"
          Case 5 ' ОО «Тарко-Сале» "ОО2 ""Тарко-Сале"""
            officeNameInReport = "ОО2 ""Тарко-Сале""" ' "Тарко-Сале"
            officeNameInReport2 = "Тарко-Сале"
        End Select
                
        ' Офис i
        ThisWorkbook.Sheets("Лист6").Cells(rowForWriteInSheet6, 1).Value = CStr(i)
        ' Офис наименование
        ThisWorkbook.Sheets("Лист6").Cells(rowForWriteInSheet6, 2).Value = getNameOfficeByNumber(i)
        ThisWorkbook.Sheets("Лист6").Cells(rowForWriteInSheet6, 2).HorizontalAlignment = xlLeft
        
        ' Цвет всей строки
        ' Call setColorCells(ThisWorkbook.Name, "Лист6", rowForWriteInSheet6, 2, rowForWriteInSheet6, 26)
        Call setColorCells(ThisWorkbook.Name, "Лист6", rowForWriteInSheet6, 2, rowForWriteInSheet6, 34)
        
        ' Сумма
        If cellSum = "" Then
          cellSum = ConvertToLetter(3) + CStr(rowForWriteInSheet6)
        Else
          cellSum = cellSum + "+" + ConvertToLetter(3) + CStr(rowForWriteInSheet6)
        End If
        ' Номер строки вывода на Листе 6
        rowForWriteInSheet6 = rowForWriteInSheet6 + 1

        ' Формируем список НОРПиКО и МРК по офису из Addr.Book
        rowCount = rowByValue(ThisWorkbook.Name, "Addr.Book", "ФИО", 100, 100) + 2
        Do While ThisWorkbook.Sheets("Addr.Book").Cells(rowCount, 3).Value <> ""
          
          ' Если это ООi и НОРПиКОi или МРКi + или УДОi
          If (ThisWorkbook.Sheets("Addr.Book").Cells(rowCount, 4).Value = "ОО" + CStr(i)) And ((ThisWorkbook.Sheets("Addr.Book").Cells(rowCount, 3).Value = "НОРПиКО" + CStr(i)) Or (ThisWorkbook.Sheets("Addr.Book").Cells(rowCount, 3).Value = "МРК" + CStr(i)) Or (ThisWorkbook.Sheets("Addr.Book").Cells(rowCount, 3).Value = "УДО" + CStr(i))) Then
            
            ' Выводим в Лист 6
            ' ФИО сотрудника
            ThisWorkbook.Sheets("Лист6").Cells(rowForWriteInSheet6, 2).Value = Фамилия_и_Имя(ThisWorkbook.Sheets("Addr.Book").Cells(rowCount, 2).Value, 3)
            allNameStr = allNameStr + ThisWorkbook.Sheets("Лист6").Cells(rowForWriteInSheet6, 2).Value + ","
            ThisWorkbook.Sheets("Лист6").Cells(rowForWriteInSheet6, 2).HorizontalAlignment = xlRight
            
            ' Дублируем ФИО сотрудника с правого конца для удобства анализа продаж ИБ и НС
            ThisWorkbook.Sheets("Лист6").Cells(rowForWriteInSheet6, 35).Value = ThisWorkbook.Sheets("Лист6").Cells(rowForWriteInSheet6, 2).Value
            ThisWorkbook.Sheets("Лист6").Cells(rowForWriteInSheet6, 35).HorizontalAlignment = xlLeft
            
            ' Клиенты
            ' ThisWorkbook.Sheets("Лист6").Cells(rowForWriteInSheet6, 3).Value = 0
            ThisWorkbook.Sheets("Лист6").Cells(rowForWriteInSheet6, 3).HorizontalAlignment = xlRight
            ' Заявки КК
            ' ThisWorkbook.Sheets("Лист6").Cells(rowForWriteInSheet6, 4).Value = 0
            ThisWorkbook.Sheets("Лист6").Cells(rowForWriteInSheet6, 4).HorizontalAlignment = xlRight
            ' Доля
            ' ThisWorkbook.Sheets("Лист6").Cells(rowForWriteInSheet6, 5).Value = 0
            ThisWorkbook.Sheets("Лист6").Cells(rowForWriteInSheet6, 5).HorizontalAlignment = xlRight
            rowForWriteInSheet6 = rowForWriteInSheet6 + 1
          End If
         
          ' Следующая запись
          rowCount = rowCount + 1
          Application.StatusBar = "Формирование списков " + officeNameInReport + ": " + CStr(rowCount) + "..."
          DoEventsInterval (rowCount)
        Loop
                
        ' Контроль значения строки со всеми МРК в AQ2
        ThisWorkbook.Sheets("Лист6").Range("AQ2").Value = allNameStr
                
        ' Обработка списка
        ' Кол-во клиентов (шт)
        countClient = 0
        ' Кол-во заявок на КК (шт)
        countCredCard = 0
        
        ' Начинаем со строки "Неделя/Куратор" +1
        rowCount = rowByValue(ReportName_String, "Клиенты (Кросс)", "Неделя/Куратор", 100, 100) + 1
        
        ' Определяем номера столбцов на листе "Клиенты (Кросс)"
        Column_Колво_клиентов = ColumnByValue(ReportName_String, "Клиенты (Кросс)", "Кол-во клиентов", 100, 100) ' 2
        Column_Колво_заявок_на_КК = ColumnByValue(ReportName_String, "Клиенты (Кросс)", "Кол-во заявок на КК", 100, 100) ' 3
        
        Column_Имеют_ИБ = ColumnByValue(ReportName_String, "Клиенты (Кросс)", "Имеют ИБ", 100, 100) ' 9
        Column_Подключение_ИБ = ColumnByValue(ReportName_String, "Клиенты (Кросс)", " Подключение ИБ", 100, 100) ' 10
        Column_ИБ_активный = ColumnByValue(ReportName_String, "Клиенты (Кросс)", " ИБ активный", 100, 100) ' 11
        Column_Не_имеют_ИБ = ColumnByValue(ReportName_String, "Клиенты (Кросс)", " Не имеют ИБ", 100, 100) ' 12
        Column_Потенциал_к_реактивации = ColumnByValue(ReportName_String, "Клиенты (Кросс)", "Потенциал к реактивации ", 100, 100)  ' 13
        
        Column_Накопительные_счета = ColumnByValue(ReportName_String, "Клиенты (Кросс)", "Накопительные счета", 100, 100)
        
        ' Нужная неделя для вывода в отчет
        beginWeekPeriod = False
        
        ' Do While Not IsEmpty(Workbooks(ReportName_String).Sheets("Клиенты (Кросс)").Cells(rowCount, 1).Value)
        Do While InStr(Workbooks(ReportName_String).Sheets("Клиенты (Кросс)").Cells(rowCount, 1).Value, "Общий итог") = 0
          
          ' Формат недели: Для Capacity 11.03.2020 г.: 10.02.03.20-08.03.20, 11.09.03.20-15.03.20 (один дефиз и 5 точек)
          If rowOfWeekPeriod(Workbooks(ReportName_String).Sheets("Клиенты (Кросс)").Cells(rowCount, 1).Value) <> 0 Then
                        
            ' Нужная неделя для вывода в отчет
            If rowOfWeekPeriod(Workbooks(ReportName_String).Sheets("Клиенты (Кросс)").Cells(rowCount, 1).Value) = WeekNumber(Date) Then
              beginWeekPeriod = True
            End If
            
          End If
          
          ' Если это один из офисов
          If InStr(Workbooks(ReportName_String).Sheets("Клиенты (Кросс)").Cells(rowCount, 1).Value, officeNameInReport) <> 0 Then
          ' If InStr(Workbooks(ReportName_String).Sheets("Клиенты (Кросс)").Cells(rowCount, 1).Value, officeNameInReport2) <> 0 Then
            countClient = countClient + Workbooks(ReportName_String).Sheets("Клиенты (Кросс)").Cells(rowCount, Column_Колво_клиентов).Value
            countCredCard = countCredCard + Workbooks(ReportName_String).Sheets("Клиенты (Кросс)").Cells(rowCount, Column_Колво_заявок_на_КК).Value
          End If
                    
          ' Следующая запись
          rowCount = rowCount + 1
          Application.StatusBar = officeNameInReport + ": " + CStr(rowCount) + "..."
          DoEventsInterval (rowCount)
        Loop
        
        ' Запоминаем переменную - номер строки на вкладке "Клиенты (Кросс)", где есть "Общий итог"
        rowCount_КлиентыКросс_Общий_итог = rowCount
        
        Application.StatusBar = ""
        
        ' Выводим данные по офису
        rowCount = rowByValue(ThisWorkbook.Name, "Лист6", getNameOfficeByNumber(i), 100, 100)
        ThisWorkbook.Sheets("Лист6").Cells(rowCount, 3).Value = countClient
        ThisWorkbook.Sheets("Лист6").Cells(rowCount, 3).HorizontalAlignment = xlCenter
        ThisWorkbook.Sheets("Лист6").Cells(rowCount, 4).Value = countCredCard
        ThisWorkbook.Sheets("Лист6").Cells(rowCount, 4).HorizontalAlignment = xlCenter
        If countClient <> 0 Then
          ThisWorkbook.Sheets("Лист6").Cells(rowCount, 5).Value = Round(countCredCard / countClient, 3)
        Else
          ThisWorkbook.Sheets("Лист6").Cells(rowCount, 5).Value = 0
        End If
        ' ThisWorkbook.Sheets("Лист6").Cells(RowCount, 5).NumberFormat = "0.0%"
        ThisWorkbook.Sheets("Лист6").Cells(rowCount, 5).NumberFormat = cellsNumberFormat(ThisWorkbook.Sheets("Лист6").Cells(rowCount, 5).Value)
        ThisWorkbook.Sheets("Лист6").Cells(rowCount, 5).HorizontalAlignment = xlCenter

      Next i ' Следующий офис
      
      ' Выводим итоги по сотрудникам
      ' Начинаем со строки "Неделя/Куратор" +1
      rowCount = rowByValue(ReportName_String, "Клиенты (Кросс)", "Неделя/Куратор", 100, 100) + 1
      
      ' Do While Not IsEmpty(Workbooks(ReportName_String).Sheets("Клиенты (Кросс)").Cells(rowCount, 1).Value)
      Do While InStr(Workbooks(ReportName_String).Sheets("Клиенты (Кросс)").Cells(rowCount, 1).Value, "Общий итог") = 0
        
        ' Если в ячейки нет символов: ", ф-л, -
        currentNameInCapacityStr = Workbooks(ReportName_String).Sheets("Клиенты (Кросс)").Cells(rowCount, 1).Value
        
        ' If (InStr(currentNameInCapacityStr, Chr(34)) = 0) And (InStr(currentNameInCapacityStr, "ф-л") = 0) And (InStr(currentNameInCapacityStr, "-") = 0) Then
        ' Вариант 2
        If (InStr(currentNameInCapacityStr, Chr(34)) = 0) And (InStr(currentNameInCapacityStr, "ф-л") = 0) And (InStr(currentNameInCapacityStr, "-") = 0) And (currentNameInCapacityStr <> "") Then
          
          ' Если это ФИО, то проверяем в подстроке всех сотрудников Тюмени
          currentNameInCapacityStr = Фамилия_и_Имя(currentNameInCapacityStr, 3)
          
          ' Дебаг
          ' If InStr(currentNameInCapacityStr, "Мельник") <> 0 Then
          '   t = 1
          ' End If
          
          If InStr(allNameStr, currentNameInCapacityStr) <> 0 Then
            '
            currentNameStr_Range = RangeByValue(ThisWorkbook.Name, "Лист6", currentNameInCapacityStr, 100, 100)
            currentNameStr_Row = ThisWorkbook.Sheets("Лист6").Range(currentNameStr_Range).Row
            currentNameStr_Column = ThisWorkbook.Sheets("Лист6").Range(currentNameStr_Range).Column
            
            ' Заносим данные по сотруднику:
            ' Клиенты
            ThisWorkbook.Sheets("Лист6").Cells(currentNameStr_Row, currentNameStr_Column + 1).Value = ThisWorkbook.Sheets("Лист6").Cells(currentNameStr_Row, currentNameStr_Column + 1).Value + Workbooks(ReportName_String).Sheets("Клиенты (Кросс)").Cells(rowCount, Column_Колво_клиентов).Value
            ' Заявки КК
            ThisWorkbook.Sheets("Лист6").Cells(currentNameStr_Row, currentNameStr_Column + 2).Value = ThisWorkbook.Sheets("Лист6").Cells(currentNameStr_Row, currentNameStr_Column + 2).Value + Workbooks(ReportName_String).Sheets("Клиенты (Кросс)").Cells(rowCount, Column_Колво_заявок_на_КК).Value
            ' Доля
            ThisWorkbook.Sheets("Лист6").Cells(currentNameStr_Row, currentNameStr_Column + 3).Value = Round(ThisWorkbook.Sheets("Лист6").Cells(currentNameStr_Row, currentNameStr_Column + 2).Value / ThisWorkbook.Sheets("Лист6").Cells(currentNameStr_Row, currentNameStr_Column + 1).Value, 3)
            ' ThisWorkbook.Sheets("Лист6").Cells(currentNameStr_Row, currentNameStr_Column + 3).NumberFormat = "0.0%"
            ThisWorkbook.Sheets("Лист6").Cells(currentNameStr_Row, currentNameStr_Column + 3).NumberFormat = cellsNumberFormat(ThisWorkbook.Sheets("Лист6").Cells(currentNameStr_Row, currentNameStr_Column + 3).Value)
            ' Окраска ячейки СФЕТОФОР: если
            targetVar = (ThisWorkbook.Sheets("Лист6").Cells(3, 5).Value) * 100
            Call Full_Color_RangeII("Лист6", currentNameStr_Row, currentNameStr_Column + 3, (РассчетДоли(ThisWorkbook.Sheets("Лист6").Cells(currentNameStr_Row, currentNameStr_Column + 1).Value, ThisWorkbook.Sheets("Лист6").Cells(currentNameStr_Row, currentNameStr_Column + 2).Value, 3) * 100), targetVar)
             
            ' *** Блок Интернет-Банка ***
            ' Интернет-Банк: Имеют ИБ (клиенты, у которых есть ИБ)
            ThisWorkbook.Sheets("Лист6").Cells(currentNameStr_Row, currentNameStr_Column + 26).Value = ThisWorkbook.Sheets("Лист6").Cells(currentNameStr_Row, currentNameStr_Column + 26).Value + Workbooks(ReportName_String).Sheets("Клиенты (Кросс)").Cells(rowCount, Column_Имеют_ИБ).Value ' столбец 9
            ThisWorkbook.Sheets("Лист6").Cells(currentNameStr_Row, currentNameStr_Column + 26).HorizontalAlignment = xlRight
            
            ' (Имеют ИБ) в т.ч. ИБ активный
            ThisWorkbook.Sheets("Лист6").Cells(currentNameStr_Row, currentNameStr_Column + 27).Value = ThisWorkbook.Sheets("Лист6").Cells(currentNameStr_Row, currentNameStr_Column + 27).Value + Workbooks(ReportName_String).Sheets("Клиенты (Кросс)").Cells(rowCount, Column_ИБ_активный).Value ' столбец 11
            ThisWorkbook.Sheets("Лист6").Cells(currentNameStr_Row, currentNameStr_Column + 27).HorizontalAlignment = xlRight
            
            ' (Имеют ИБ) в т.ч. Потенциал к реактивации
            ThisWorkbook.Sheets("Лист6").Cells(currentNameStr_Row, currentNameStr_Column + 28).Value = ThisWorkbook.Sheets("Лист6").Cells(currentNameStr_Row, currentNameStr_Column + 28).Value + Workbooks(ReportName_String).Sheets("Клиенты (Кросс)").Cells(rowCount, Column_Потенциал_к_реактивации).Value ' столбец 14
            ThisWorkbook.Sheets("Лист6").Cells(currentNameStr_Row, currentNameStr_Column + 28).HorizontalAlignment = xlRight
            
            ' Не имеют ИБ
            ThisWorkbook.Sheets("Лист6").Cells(currentNameStr_Row, currentNameStr_Column + 29).Value = ThisWorkbook.Sheets("Лист6").Cells(currentNameStr_Row, currentNameStr_Column + 29).Value + Workbooks(ReportName_String).Sheets("Клиенты (Кросс)").Cells(rowCount, Column_Не_имеют_ИБ).Value ' столбец 12
            ThisWorkbook.Sheets("Лист6").Cells(currentNameStr_Row, currentNameStr_Column + 29).HorizontalAlignment = xlRight
            
            ' Подключение ИБ
            ThisWorkbook.Sheets("Лист6").Cells(currentNameStr_Row, currentNameStr_Column + 30).Value = ThisWorkbook.Sheets("Лист6").Cells(currentNameStr_Row, currentNameStr_Column + 30).Value + Workbooks(ReportName_String).Sheets("Клиенты (Кросс)").Cells(rowCount, Column_Подключение_ИБ).Value ' столбец 10
            ThisWorkbook.Sheets("Лист6").Cells(currentNameStr_Row, currentNameStr_Column + 30).HorizontalAlignment = xlRight
            
            ' Упущено ИБ с 29.12 делаем расчет, так как этого столбца нет в Capacity: Клиенты - Имеют ИБ - Подключение ИБ
            ' почему-то дает с минусами, пока отключил
            ' ThisWorkbook.Sheets("Лист6").Cells(currentNameStr_Row, currentNameStr_Column + 30).Value = ThisWorkbook.Sheets("Лист6").Cells(currentNameStr_Row, currentNameStr_Column + 1).Value - ThisWorkbook.Sheets("Лист6").Cells(currentNameStr_Row, currentNameStr_Column + 26).Value - ThisWorkbook.Sheets("Лист6").Cells(currentNameStr_Row, currentNameStr_Column + 27).Value ' ThisWorkbook.Sheets("Лист6").Cells(currentNameStr_Row, currentNameStr_Column + 30).Value + Workbooks(ReportName_String).Sheets("Клиенты (Кросс)").Cells(rowCount, 13).Value
            ' ThisWorkbook.Sheets("Лист6").Cells(currentNameStr_Row, currentNameStr_Column + 30).HorizontalAlignment = xlRight
            ' *** Блок Интернет-Банка ***
            
            ' Накопительные счета
            ThisWorkbook.Sheets("Лист6").Cells(currentNameStr_Row, currentNameStr_Column + 32).Value = ThisWorkbook.Sheets("Лист6").Cells(currentNameStr_Row, currentNameStr_Column + 32).Value + Workbooks(ReportName_String).Sheets("Клиенты (Кросс)").Cells(rowCount, Column_Накопительные_счета).Value ' был столбец 18
            ThisWorkbook.Sheets("Лист6").Cells(currentNameStr_Row, currentNameStr_Column + 32).HorizontalAlignment = xlRight

           End If
        End If
        ' Следующая запись
        rowCount = rowCount + 1
        Application.StatusBar = "Сотрудники: " + CStr(rowCount) + "..."
        DoEventsInterval (rowCount)
      Loop
      
      ' Итоги по РОО клиентам и по заявкам в столбце C и D
      row_Итого_по_РОО = rowForWriteInSheet6
      ThisWorkbook.Sheets("Лист6").Cells(row_Итого_по_РОО, 2).Value = "Итого по РОО: "
      ThisWorkbook.Sheets("Лист6").Cells(row_Итого_по_РОО, 3).Formula = "=SUM(" + cellSum + ")"
      ThisWorkbook.Sheets("Лист6").Cells(row_Итого_по_РОО, 3).HorizontalAlignment = xlCenter
      ThisWorkbook.Sheets("Лист6").Cells(row_Итого_по_РОО, 4).Formula = "=SUM(" + Replace(cellSum, "C", "D") + ")"
      ThisWorkbook.Sheets("Лист6").Cells(row_Итого_по_РОО, 4).HorizontalAlignment = xlCenter
      ThisWorkbook.Sheets("Лист6").Cells(row_Итого_по_РОО, 5).Value = Round(ThisWorkbook.Sheets("Лист6").Cells(row_Итого_по_РОО, 4).Value / ThisWorkbook.Sheets("Лист6").Cells(row_Итого_по_РОО, 3).Value, 2)
      '
      ' ThisWorkbook.Sheets("Лист6").Cells(row_Итого_по_РОО, 5).NumberFormat = "0.0%"
      ThisWorkbook.Sheets("Лист6").Cells(row_Итого_по_РОО, 5).NumberFormat = cellsNumberFormat(ThisWorkbook.Sheets("Лист6").Cells(row_Итого_по_РОО, 5).Value)
      
      ThisWorkbook.Sheets("Лист6").Cells(row_Итого_по_РОО, 5).HorizontalAlignment = xlCenter
      Application.StatusBar = ""
      
      ' --- Цикл №1 прошивка КК на входящем потоке ---


      ' Цикл №2 Pre-Approved КК на входящем потоке: Лист "PA_KK" ---
      
      ' Переходим на вкладку "PA_KK"
      Call openPivotTables_Capacity_PA_KK(ReportName_String)
      
      ' Определение столбцов
      ' Кол-во клиентов
      Column_Колво_клиентов = ColumnByValue(ReportName_String, "PA_KK", "Кол-во клиентов", 100, 100) ' 2
      ' Клиенты с КК
      Column_Клиенты_с_КК = ColumnByValue(ReportName_String, "PA_KK", "Клиенты с КК", 100, 100) ' 3
      ' Выдано РА КК
      Column_Выдано_РА_КК = ColumnByValue(ReportName_String, "PA_KK", "Выдано РА КК", 100, 100) ' 6
      '  Заказано РА_КК
      Column_Заказано_РА_КК = ColumnByValue(ReportName_String, "PA_KK", " Заказано РА_КК", 100, 100) ' 8
      ' Клиенты с предв. решением по РА_КК
      Column_Клиенты_с_предв_решением_по_РА_КК = ColumnByValue(ReportName_String, "PA_KK", "Клиенты с предв. решением по РА_КК", 100, 100) ' 9
      ' Выдано PA_KK Предложение
      Column_Выдано_PA_KK_Предложение = ColumnByValue(ReportName_String, "PA_KK", "Выдано PA_KK Предложение", 100, 100) ' 10
      
      ' Обработка 1
      ' Начинаем со строки "Филиал/ДО" +1
      rowCount = rowByValue(ReportName_String, "PA_KK", "Филиал/ДО", 100, 100) + 1
      Do While InStr(Workbooks(ReportName_String).Sheets("PA_KK").Cells(rowCount, 1).Value, "Общий итог") = 0
        
        ' Если в ячейки нет символов: ", ф-л, -, "", Москва
        currentNameInCapacityStr = Workbooks(ReportName_String).Sheets("PA_KK").Cells(rowCount, 1).Value
        
        ' Проверяем содержимое ячейки
        If (currentNameInCapacityStr <> "") And (InStr(currentNameInCapacityStr, Chr(34)) = 0) And (InStr(currentNameInCapacityStr, "ф-л") = 0) And (InStr(currentNameInCapacityStr, "-") = 0) And (InStr(currentNameInCapacityStr, "Москва") = 0) And (InStr(currentNameInCapacityStr, "ГО") = 0) Then
          
          ' Если это ФИО, то проверяем в подстроке всех сотрудников Тюмени
          currentNameInCapacityStr = Фамилия_и_Имя(currentNameInCapacityStr, 3)
          If InStr(allNameStr, currentNameInCapacityStr) <> 0 Then
            '
            currentNameStr_Range = RangeByValue(ThisWorkbook.Name, "Лист6", currentNameInCapacityStr, 100, 100)
            currentNameStr_Row = ThisWorkbook.Sheets("Лист6").Range(currentNameStr_Range).Row
            currentNameStr_Column = ThisWorkbook.Sheets("Лист6").Range(currentNameStr_Range).Column
            ' Заносим данные по сотруднику
            ' Кол-во клиентов (из B2) в Клиенты (в F5)
            ThisWorkbook.Sheets("Лист6").Cells(currentNameStr_Row, currentNameStr_Column + 4).Value = Workbooks(ReportName_String).Sheets("PA_KK").Cells(rowCount, Column_Колво_клиентов).Value ' 2
            ThisWorkbook.Sheets("Лист6").Cells(currentNameStr_Row, currentNameStr_Column + 4).HorizontalAlignment = xlRight
            ' Клиенты с КК (из B3) в с PA КК (в G6)
            ThisWorkbook.Sheets("Лист6").Cells(currentNameStr_Row, currentNameStr_Column + 5).Value = Workbooks(ReportName_String).Sheets("PA_KK").Cells(rowCount, Column_Клиенты_с_КК).Value ' 3
            ThisWorkbook.Sheets("Лист6").Cells(currentNameStr_Row, currentNameStr_Column + 5).HorizontalAlignment = xlRight
            ' Выдано РА КК
            ThisWorkbook.Sheets("Лист6").Cells(currentNameStr_Row, currentNameStr_Column + 6).Value = Workbooks(ReportName_String).Sheets("PA_KK").Cells(rowCount, Column_Выдано_РА_КК).Value ' 6
            ThisWorkbook.Sheets("Лист6").Cells(currentNameStr_Row, currentNameStr_Column + 6).HorizontalAlignment = xlRight
            
            ' Доля
            ThisWorkbook.Sheets("Лист6").Cells(currentNameStr_Row, currentNameStr_Column + 7).Value = РассчетДоли(ThisWorkbook.Sheets("Лист6").Cells(currentNameStr_Row, currentNameStr_Column + 5).Value, ThisWorkbook.Sheets("Лист6").Cells(currentNameStr_Row, currentNameStr_Column + 6).Value, 3)
            ' ThisWorkbook.Sheets("Лист6").Cells(currentNameStr_Row, currentNameStr_Column + 7).NumberFormat = "0.0%"
            ThisWorkbook.Sheets("Лист6").Cells(currentNameStr_Row, currentNameStr_Column + 7).NumberFormat = cellsNumberFormat(ThisWorkbook.Sheets("Лист6").Cells(currentNameStr_Row, currentNameStr_Column + 7).Value)
            ThisWorkbook.Sheets("Лист6").Cells(currentNameStr_Row, currentNameStr_Column + 7).HorizontalAlignment = xlRight
            ' Заказано
            ThisWorkbook.Sheets("Лист6").Cells(currentNameStr_Row, currentNameStr_Column + 8).Value = Workbooks(ReportName_String).Sheets("PA_KK").Cells(rowCount, Column_Заказано_РА_КК).Value ' 8
            ThisWorkbook.Sheets("Лист6").Cells(currentNameStr_Row, currentNameStr_Column + 8).HorizontalAlignment = xlRight
            ' Клиенты с предв. решением по РА_КК (9)
            ThisWorkbook.Sheets("Лист6").Cells(currentNameStr_Row, currentNameStr_Column + 9).Value = Workbooks(ReportName_String).Sheets("PA_KK").Cells(rowCount, Column_Клиенты_с_предв_решением_по_РА_КК).Value ' 9
            ThisWorkbook.Sheets("Лист6").Cells(currentNameStr_Row, currentNameStr_Column + 9).HorizontalAlignment = xlRight
            ' Выдано PA_KK Предложение (10)
            ThisWorkbook.Sheets("Лист6").Cells(currentNameStr_Row, currentNameStr_Column + 10).Value = Workbooks(ReportName_String).Sheets("PA_KK").Cells(rowCount, Column_Выдано_PA_KK_Предложение).Value ' 10
            ThisWorkbook.Sheets("Лист6").Cells(currentNameStr_Row, currentNameStr_Column + 10).HorizontalAlignment = xlRight
            ' Доля 2()
            ThisWorkbook.Sheets("Лист6").Cells(currentNameStr_Row, currentNameStr_Column + 11).Value = РассчетДоли(ThisWorkbook.Sheets("Лист6").Cells(currentNameStr_Row, currentNameStr_Column + 9).Value, ThisWorkbook.Sheets("Лист6").Cells(currentNameStr_Row, currentNameStr_Column + 10).Value, 3)
            ' ThisWorkbook.Sheets("Лист6").Cells(currentNameStr_Row, currentNameStr_Column + 11).NumberFormat = "0.0%"
            ThisWorkbook.Sheets("Лист6").Cells(currentNameStr_Row, currentNameStr_Column + 11).NumberFormat = cellsNumberFormat(ThisWorkbook.Sheets("Лист6").Cells(currentNameStr_Row, currentNameStr_Column + 11).Value)
            ThisWorkbook.Sheets("Лист6").Cells(currentNameStr_Row, currentNameStr_Column + 11).HorizontalAlignment = xlRight
            
           End If
        End If
        
        ' Если в ячейке "Тюменский ОО1", то запоминаем позицию для последующей обработки упущенных клиентов - Открывается Лист1
        If InStr(Workbooks(ReportName_String).Sheets("PA_KK").Cells(rowCount, 1).Value, "Тюменский ОО1") <> 0 Then
          
          row_Тюменский_ОО1 = rowCount
                    
          ' Workbooks(ReportName_String).Activate ' !!!
          Workbooks(ReportName_String).Sheets("PA_KK").Cells(rowCount, Column_Клиенты_с_КК).ShowDetail = True ' 3
          
          ' Workbooks(ReportName_String).Sheets("PA_KK").Cells(RowCount, 3).ShowDetail = True
          '
          ' Range("Таблица1[[#Headers],[PA_KK]]").Select ' !!!
          
          ' ActiveSheet.ListObjects("Таблица1").Range.AutoFilter Field:=7, Criteria1:="1"
          
          ' Открылся Лист1 в Capacity - Выгружаем клиентов
          ' PA_KK = "1"
          
          ' KK-GR = "1" - это не выгружаем пока, там где KK-GR=1 и PA_KK=1
          
          ' Возврат на лист "PA_KK"
          ' Workbooks(ReportName_String).Sheets("PA_KK").Select
          
          ' Переходим на окно DB
          ThisWorkbook.Sheets("Лист6").Activate ' !!!
 
        End If
        
        ' Следующая запись
        rowCount = rowCount + 1
        Application.StatusBar = "Сотрудники: " + CStr(rowCount) + "..."
        DoEventsInterval (rowCount)
      Loop

      ' Обработка 2: Итоги по офисам и по РОО (идем снизу вверх)
      rowCount = rowByValue(ThisWorkbook.Name, "Лист6", "Итого по РОО:", 100, 100) - 1
      ' Переменные итогов по ОО
      Клиенты = 0
      Клиенты_с_PA_КК = 0
      Выдано_PA = 0
      Заказ_PA_КК = 0
      с_ПР_PA_КК = 0
      Выдано_ПР_PA_КК = 0
      '
      Имеют_ИБ = 0
      Подключен_ИБ = 0
      ИБ_активен = 0
      Нет_ИБ = 0
      Упущено_ИБ = 0
      Потенциал_реакт = 0
      '
      Накопительные_счета = 0

      Do While ThisWorkbook.Sheets("Лист6").Cells(rowCount, 2).Value <> ""
        
        ' Если в строке "OO"
        If InStr(ThisWorkbook.Sheets("Лист6").Cells(rowCount, 2).Value, "ОО") <> 0 Then
          ' Заносим итоги
          ThisWorkbook.Sheets("Лист6").Cells(rowCount, 6).Value = Клиенты
          ThisWorkbook.Sheets("Лист6").Cells(rowCount, 6).HorizontalAlignment = xlCenter
          ThisWorkbook.Sheets("Лист6").Cells(rowCount, 7).Value = Клиенты_с_PA_КК
          ThisWorkbook.Sheets("Лист6").Cells(rowCount, 7).HorizontalAlignment = xlCenter
          ThisWorkbook.Sheets("Лист6").Cells(rowCount, 8).Value = Выдано_PA
          ThisWorkbook.Sheets("Лист6").Cells(rowCount, 8).HorizontalAlignment = xlCenter
          ' Доля
          ThisWorkbook.Sheets("Лист6").Cells(rowCount, 9).Value = РассчетДоли(ThisWorkbook.Sheets("Лист6").Cells(rowCount, 7).Value, ThisWorkbook.Sheets("Лист6").Cells(rowCount, 8).Value, 3)
          ' ThisWorkbook.Sheets("Лист6").Cells(RowCount, 9).NumberFormat = "0.0%"
          ThisWorkbook.Sheets("Лист6").Cells(rowCount, 9).NumberFormat = cellsNumberFormat(ThisWorkbook.Sheets("Лист6").Cells(rowCount, 9).Value)
          ThisWorkbook.Sheets("Лист6").Cells(rowCount, 9).HorizontalAlignment = xlCenter
          '
          ThisWorkbook.Sheets("Лист6").Cells(rowCount, 10).Value = Заказ_PA_КК
          ThisWorkbook.Sheets("Лист6").Cells(rowCount, 10).HorizontalAlignment = xlCenter
          ThisWorkbook.Sheets("Лист6").Cells(rowCount, 11).Value = с_ПР_PA_КК
          ThisWorkbook.Sheets("Лист6").Cells(rowCount, 11).HorizontalAlignment = xlCenter
          ThisWorkbook.Sheets("Лист6").Cells(rowCount, 12).Value = Выдано_ПР_PA_КК
          ThisWorkbook.Sheets("Лист6").Cells(rowCount, 12).HorizontalAlignment = xlCenter
          ' Доля
          ThisWorkbook.Sheets("Лист6").Cells(rowCount, 13).Value = РассчетДоли(ThisWorkbook.Sheets("Лист6").Cells(rowCount, 11).Value, ThisWorkbook.Sheets("Лист6").Cells(rowCount, 12).Value, 3)
          ' ThisWorkbook.Sheets("Лист6").Cells(RowCount, 13).NumberFormat = "0.0%"
          ThisWorkbook.Sheets("Лист6").Cells(rowCount, 13).NumberFormat = cellsNumberFormat(ThisWorkbook.Sheets("Лист6").Cells(rowCount, 13).Value)
          ThisWorkbook.Sheets("Лист6").Cells(rowCount, 13).HorizontalAlignment = xlCenter
          
          ' ИБ выводим итоги по ОО
          ThisWorkbook.Sheets("Лист6").Cells(rowCount, 28).Value = Имеют_ИБ
          ThisWorkbook.Sheets("Лист6").Cells(rowCount, 28).HorizontalAlignment = xlCenter
          ThisWorkbook.Sheets("Лист6").Cells(rowCount, 29).Value = Подключен_ИБ
          ThisWorkbook.Sheets("Лист6").Cells(rowCount, 29).HorizontalAlignment = xlCenter
          ThisWorkbook.Sheets("Лист6").Cells(rowCount, 30).Value = ИБ_активен
          ThisWorkbook.Sheets("Лист6").Cells(rowCount, 30).HorizontalAlignment = xlCenter
          ThisWorkbook.Sheets("Лист6").Cells(rowCount, 31).Value = Нет_ИБ
          ThisWorkbook.Sheets("Лист6").Cells(rowCount, 31).HorizontalAlignment = xlCenter
          ThisWorkbook.Sheets("Лист6").Cells(rowCount, 32).Value = Упущено_ИБ
          ThisWorkbook.Sheets("Лист6").Cells(rowCount, 32).HorizontalAlignment = xlCenter
          ThisWorkbook.Sheets("Лист6").Cells(rowCount, 33).Value = Потенциал_реакт
          ThisWorkbook.Sheets("Лист6").Cells(rowCount, 33).HorizontalAlignment = xlCenter
          
          ' НС выводим итоги по ОО
          ThisWorkbook.Sheets("Лист6").Cells(rowCount, 34).Value = Накопительные_счета
          ThisWorkbook.Sheets("Лист6").Cells(rowCount, 34).HorizontalAlignment = xlCenter
                    
          ' Выводим в примечание к DB (Лист8)
          ' ThisWorkbook.Sheets("Лист8").Range(Range_Лист8_Интернет_банк(ThisWorkbook.Sheets("Лист6").Cells(rowCount, 2).Value)).Value = "Capacity: Упущены " + CStr(Упущено_ИБ) + ", Потенц. реактив. " + CStr(Потенциал_реакт) + " Итого: " + CStr(Упущено_ИБ + Потенциал_реакт) + " шт."
                    
          ' Обнуляем переменные
          Клиенты = 0
          Клиенты_с_PA_КК = 0
          Выдано_PA = 0
          Заказ_PA_КК = 0
          с_ПР_PA_КК = 0
          Выдано_ПР_PA_КК = 0
          ' ИБ
          Имеют_ИБ = 0
          Подключен_ИБ = 0
          ИБ_активен = 0
          Нет_ИБ = 0
          Упущено_ИБ = 0
          Потенциал_реакт = 0
          ' НС
          Накопительные_счета = 0
        
        Else
          ' Суммируем
          Клиенты = Клиенты + ThisWorkbook.Sheets("Лист6").Cells(rowCount, 6).Value
          Клиенты_с_PA_КК = Клиенты_с_PA_КК + ThisWorkbook.Sheets("Лист6").Cells(rowCount, 7).Value
          Выдано_PA = Выдано_PA + ThisWorkbook.Sheets("Лист6").Cells(rowCount, 8).Value
          Заказ_PA_КК = Заказ_PA_КК + ThisWorkbook.Sheets("Лист6").Cells(rowCount, 10).Value
          с_ПР_PA_КК = с_ПР_PA_КК + ThisWorkbook.Sheets("Лист6").Cells(rowCount, 11).Value
          Выдано_ПР_PA_КК = Выдано_ПР_PA_КК + ThisWorkbook.Sheets("Лист6").Cells(rowCount, 12).Value
          ' ИБ
          Имеют_ИБ = Имеют_ИБ + ThisWorkbook.Sheets("Лист6").Cells(rowCount, 28).Value
          Подключен_ИБ = Подключен_ИБ + ThisWorkbook.Sheets("Лист6").Cells(rowCount, 29).Value
          ИБ_активен = ИБ_активен + ThisWorkbook.Sheets("Лист6").Cells(rowCount, 30).Value
          Нет_ИБ = Нет_ИБ + ThisWorkbook.Sheets("Лист6").Cells(rowCount, 31).Value
          Упущено_ИБ = Упущено_ИБ + ThisWorkbook.Sheets("Лист6").Cells(rowCount, 32).Value
          Потенциал_реакт = Потенциал_реакт + ThisWorkbook.Sheets("Лист6").Cells(rowCount, 33).Value
          ' НС
          Накопительные_счета = Накопительные_счета + ThisWorkbook.Sheets("Лист6").Cells(rowCount, 34).Value
        
        End If
        
        ' Следующая запись
        rowCount = rowCount - 1
        Application.StatusBar = "Сотрудники: " + CStr(rowCount) + "..."
        DoEventsInterval (rowCount)
      Loop
           
      ' Итоги по РОО
      ThisWorkbook.Sheets("Лист6").Cells(row_Итого_по_РОО, 6).Formula = "=SUM(" + Replace(cellSum, "C", "F") + ")"
      ThisWorkbook.Sheets("Лист6").Cells(row_Итого_по_РОО, 6).HorizontalAlignment = xlCenter
      '
      ThisWorkbook.Sheets("Лист6").Cells(row_Итого_по_РОО, 7).Formula = "=SUM(" + Replace(cellSum, "C", "G") + ")"
      ThisWorkbook.Sheets("Лист6").Cells(row_Итого_по_РОО, 7).HorizontalAlignment = xlCenter
      '
      ThisWorkbook.Sheets("Лист6").Cells(row_Итого_по_РОО, 8).Formula = "=SUM(" + Replace(cellSum, "C", "H") + ")"
      ThisWorkbook.Sheets("Лист6").Cells(row_Итого_по_РОО, 8).HorizontalAlignment = xlCenter
      ' Доля
      ThisWorkbook.Sheets("Лист6").Cells(row_Итого_по_РОО, 9).Value = РассчетДоли(ThisWorkbook.Sheets("Лист6").Cells(row_Итого_по_РОО, 7).Value, ThisWorkbook.Sheets("Лист6").Cells(row_Итого_по_РОО, 8).Value, 3)
      ' ThisWorkbook.Sheets("Лист6").Cells(row_Итого_по_РОО, 9).NumberFormat = "0.0%"
      ThisWorkbook.Sheets("Лист6").Cells(row_Итого_по_РОО, 9).NumberFormat = cellsNumberFormat(ThisWorkbook.Sheets("Лист6").Cells(row_Итого_по_РОО, 9).Value)
      ThisWorkbook.Sheets("Лист6").Cells(row_Итого_по_РОО, 9).HorizontalAlignment = xlCenter
      '
      ThisWorkbook.Sheets("Лист6").Cells(row_Итого_по_РОО, 10).Formula = "=SUM(" + Replace(cellSum, "C", "J") + ")"
      ThisWorkbook.Sheets("Лист6").Cells(row_Итого_по_РОО, 10).HorizontalAlignment = xlCenter
      '
      ThisWorkbook.Sheets("Лист6").Cells(row_Итого_по_РОО, 11).Formula = "=SUM(" + Replace(cellSum, "C", "K") + ")"
      ThisWorkbook.Sheets("Лист6").Cells(row_Итого_по_РОО, 11).HorizontalAlignment = xlCenter
      '
      ThisWorkbook.Sheets("Лист6").Cells(row_Итого_по_РОО, 12).Formula = "=SUM(" + Replace(cellSum, "C", "L") + ")"
      ThisWorkbook.Sheets("Лист6").Cells(row_Итого_по_РОО, 12).HorizontalAlignment = xlCenter
      ' Доля 2
      ThisWorkbook.Sheets("Лист6").Cells(row_Итого_по_РОО, 13).Value = РассчетДоли(ThisWorkbook.Sheets("Лист6").Cells(row_Итого_по_РОО, 11).Value, ThisWorkbook.Sheets("Лист6").Cells(row_Итого_по_РОО, 12).Value, 3)
      ' ThisWorkbook.Sheets("Лист6").Cells(row_Итого_по_РОО, 13).NumberFormat = "0.0%"
      ThisWorkbook.Sheets("Лист6").Cells(row_Итого_по_РОО, 13).NumberFormat = cellsNumberFormat(ThisWorkbook.Sheets("Лист6").Cells(row_Итого_по_РОО, 13).Value)
      ThisWorkbook.Sheets("Лист6").Cells(row_Итого_по_РОО, 13).HorizontalAlignment = xlCenter
      
      ' ИБ выводим итоги по РОО
      ThisWorkbook.Sheets("Лист6").Cells(row_Итого_по_РОО, 28).Formula = "=SUM(" + Replace(cellSum, "C", "AB") + ")"
      ThisWorkbook.Sheets("Лист6").Cells(row_Итого_по_РОО, 28).HorizontalAlignment = xlCenter
      '
      ThisWorkbook.Sheets("Лист6").Cells(row_Итого_по_РОО, 29).Formula = "=SUM(" + Replace(cellSum, "C", "AC") + ")"
      ThisWorkbook.Sheets("Лист6").Cells(row_Итого_по_РОО, 29).HorizontalAlignment = xlCenter
      '
      ThisWorkbook.Sheets("Лист6").Cells(row_Итого_по_РОО, 30).Formula = "=SUM(" + Replace(cellSum, "C", "AD") + ")"
      ThisWorkbook.Sheets("Лист6").Cells(row_Итого_по_РОО, 30).HorizontalAlignment = xlCenter
      '
      ThisWorkbook.Sheets("Лист6").Cells(row_Итого_по_РОО, 31).Formula = "=SUM(" + Replace(cellSum, "C", "AE") + ")"
      ThisWorkbook.Sheets("Лист6").Cells(row_Итого_по_РОО, 31).HorizontalAlignment = xlCenter
      '
      ThisWorkbook.Sheets("Лист6").Cells(row_Итого_по_РОО, 32).Formula = "=SUM(" + Replace(cellSum, "C", "AF") + ")"
      ThisWorkbook.Sheets("Лист6").Cells(row_Итого_по_РОО, 32).HorizontalAlignment = xlCenter
      '
      ' ThisWorkbook.Sheets("Лист6").Cells(row_Итого_по_РОО, 33).Formula = "=SUM(" + Replace(cellSum, "C", "AG") + ")"
      ' ThisWorkbook.Sheets("Лист6").Cells(row_Итого_по_РОО, 33).HorizontalAlignment = xlCenter
      
      
      ' НС выводим итоги по РОО
      ThisWorkbook.Sheets("Лист6").Cells(row_Итого_по_РОО, 34).Formula = "=SUM(" + Replace(cellSum, "C", "AH") + ")"
      ThisWorkbook.Sheets("Лист6").Cells(row_Итого_по_РОО, 34).HorizontalAlignment = xlCenter

      
      ' --- Цикл №3 выгрузка всех упущенных клиентов с PA КК и с ПР PA КК в отдельный файл для проработки в рамках акции
      
      ' (выше выгрузка)
      
      ' --- Цикл №2 Pre-Approved КК на входящем потоке ---
      
      ' Цикл №3 Pre-Approved ПК на входящем потоке ---
      
      ' Открытие сводных таблиц в Capacity на Листе "Pre-Approved"
      Call openPivotTables_Capacity_PA_ПК(ReportName_String)

      ' --- Цикл №3 Pre-Approved ПК на входящем потоке ---
      
      ' Определяем столбцы на Лист "Pre-Approved" в Книге Capacity
      Column_Колво_клиентов = ColumnByValue(ReportName_String, "Pre-Approved", "Кол-во клиентов", 100, 100) ' 2
      
      ' "Клиентов с Pre-Approved" (стоит символ переноса строки)
      ' Column_Клиентов_с_PreApproved = ColumnByValue(ReportName_String, "Pre-Approved", "с Pre-Approved", 100, 100) ' 3
      ' Column_Клиентов_с_PreApproved = ColumnByValue(ReportName_String, "Pre-Approved", "Клиентов" + Chr(32) + "с Pre-Approved", 100, 100) ' 3
      ' не находит - есть скрытый знак в ячейке
      Column_Клиентов_с_PreApproved = Column_Колво_клиентов + 1
      
      ' "Выдано PA"
      Column_Выдано_PA = ColumnByValue(ReportName_String, "Pre-Approved", "Выдано PA", 100, 100) ' 6
      
      ' Обработка 1 (Pre-Approved ПК)
      ' Начинаем со строки "Филиал/ДО" +1
      rowCount = rowByValue(ReportName_String, "Pre-Approved", "Филиал/ДО", 100, 100) + 1
      Do While InStr(Workbooks(ReportName_String).Sheets("Pre-Approved").Cells(rowCount, 1).Value, "Общий итог") = 0
        
        ' Если в ячейки нет символов: ", ф-л, -, "", Москва
        currentNameInCapacityStr = Workbooks(ReportName_String).Sheets("Pre-Approved").Cells(rowCount, 1).Value
        
        ' Проверяем содержимое ячейки
        If (currentNameInCapacityStr <> "") And (InStr(currentNameInCapacityStr, Chr(34)) = 0) And (InStr(currentNameInCapacityStr, "ф-л") = 0) And (InStr(currentNameInCapacityStr, "-") = 0) And (InStr(currentNameInCapacityStr, "Москва") = 0) And (InStr(currentNameInCapacityStr, "ГО") = 0) Then
          
          ' Если это ФИО, то проверяем в подстроке всех сотрудников Тюмени
          currentNameInCapacityStr = Фамилия_и_Имя(currentNameInCapacityStr, 3)
          
          If InStr(allNameStr, currentNameInCapacityStr) <> 0 Then
            '
            currentNameStr_Range = RangeByValue(ThisWorkbook.Name, "Лист6", currentNameInCapacityStr, 100, 100)
            currentNameStr_Row = ThisWorkbook.Sheets("Лист6").Range(currentNameStr_Range).Row
            currentNameStr_Column = ThisWorkbook.Sheets("Лист6").Range(currentNameStr_Range).Column
            ' Заносим данные по сотруднику
            
            ' Кол-во клиентов (Pre-Approved ПК)
            ThisWorkbook.Sheets("Лист6").Cells(currentNameStr_Row, currentNameStr_Column + 12).Value = Workbooks(ReportName_String).Sheets("Pre-Approved").Cells(rowCount, Column_Колво_клиентов).Value ' 2
            ThisWorkbook.Sheets("Лист6").Cells(currentNameStr_Row, currentNameStr_Column + 12).HorizontalAlignment = xlRight
            ' Клиентов с Pre - Approved (Pre-Approved ПК)
            ThisWorkbook.Sheets("Лист6").Cells(currentNameStr_Row, currentNameStr_Column + 13).Value = Workbooks(ReportName_String).Sheets("Pre-Approved").Cells(rowCount, Column_Клиентов_с_PreApproved).Value
            ThisWorkbook.Sheets("Лист6").Cells(currentNameStr_Row, currentNameStr_Column + 13).HorizontalAlignment = xlRight
            ' Выдано PA (Pre-Approved ПК)
            ThisWorkbook.Sheets("Лист6").Cells(currentNameStr_Row, currentNameStr_Column + 14).Value = Workbooks(ReportName_String).Sheets("Pre-Approved").Cells(rowCount, Column_Выдано_PA).Value ' 6
            ThisWorkbook.Sheets("Лист6").Cells(currentNameStr_Row, currentNameStr_Column + 14).HorizontalAlignment = xlRight
            
            ' Доля (Pre-Approved ПК) ФАКТ
            ThisWorkbook.Sheets("Лист6").Cells(currentNameStr_Row, currentNameStr_Column + 15).Value = РассчетДоли(ThisWorkbook.Sheets("Лист6").Cells(currentNameStr_Row, currentNameStr_Column + 13).Value, ThisWorkbook.Sheets("Лист6").Cells(currentNameStr_Row, currentNameStr_Column + 14).Value, 3)
            ' ThisWorkbook.Sheets("Лист6").Cells(currentNameStr_Row, currentNameStr_Column + 15).NumberFormat = "0.0%"
            ThisWorkbook.Sheets("Лист6").Cells(currentNameStr_Row, currentNameStr_Column + 15).NumberFormat = cellsNumberFormat(ThisWorkbook.Sheets("Лист6").Cells(currentNameStr_Row, currentNameStr_Column + 15).Value)
            ThisWorkbook.Sheets("Лист6").Cells(currentNameStr_Row, currentNameStr_Column + 15).HorizontalAlignment = xlRight
            
            ' Окраска ячейки СФЕТОФОР: если
            targetVar = (ThisWorkbook.Sheets("Лист6").Cells(3, 17).Value) * 100 ' Pre-Approved ПК
            Call Full_Color_RangeII("Лист6", currentNameStr_Row, currentNameStr_Column + 15, (РассчетДоли(ThisWorkbook.Sheets("Лист6").Cells(currentNameStr_Row, currentNameStr_Column + 13).Value, ThisWorkbook.Sheets("Лист6").Cells(currentNameStr_Row, currentNameStr_Column + 14).Value, 3) * 100), targetVar)
                   
           End If
        End If
        
        ' Если в ячейке "Тюменский ОО1", то запоминаем позицию для последующей обработки упущенных клиентов
        If InStr(Workbooks(ReportName_String).Sheets("Pre-Approved").Cells(rowCount, 1).Value, "Тюменский ОО1") <> 0 Then
          
          row_Тюменский_ОО1 = rowCount
                  
          ' Workbooks(ReportName_String).Activate ' !!!
          Workbooks(ReportName_String).Sheets("Pre-Approved").Cells(rowCount, 3).ShowDetail = True
          
          ' Workbooks(ReportName_String).Sheets("Pre-Approved").Cells(RowCount, 3).ShowDetail = True
          ' (Pre-Approved ПК)
          ' Range("Таблица2[[#Headers],[PA]]").Select
          
          ' ActiveSheet.ListObjects("Таблица2").Range.AutoFilter Field:=4, Criteria1:="1"
          
          ' Открылся Лист2 в Capacity - Выгружаем клиентов
          ' PA = "1"
          ' ID клиента (Retail) PA  Date of activ   Выдача PA
          ' 8397429              1   20200303          1       - это пример выданного PA ПК
                                           
          ' Переходим на окно DB
          ThisWorkbook.Sheets("Лист6").Activate ' !!!

        End If
        
        ' Следующая запись
        rowCount = rowCount + 1
        Application.StatusBar = "Сотрудники: " + CStr(rowCount) + "..."
        DoEventsInterval (rowCount)
      Loop

      ' Вот здесь можно вставить выгрузку клиентов по ПК и КК с Лист1 сводной таблицы, потому как и в случае PA ПК и PA KK - идет открытие Лист1 и Лист2 с одинаковым набором данных
      
      ' Создаем выходную книгу для выгрузки PA
      OutBookName = ThisWorkbook.Path + "\Out\Pre-Approved_" + strDDMMYYYY(dateReportCapacity) + ".xlsx"
      ' Вложение2
      ThisWorkbook.Sheets("Лист6").Range("AO3").Value = OutBookName
      ' Создать файл
      Call createBook_out_PA(OutBookName)
      
      ' Переходим на окно DB
      ThisWorkbook.Sheets("Лист6").Activate

      ' ===========================================================================================================================
      ' Поиск PA на Лист1
      
          ' Столбцы - сделать замену на поиск в сводной таблице номера столбца по имени
          Column_ID_клиента_Retail = ColumnByName(ReportName_String, "Лист1", 1, "ID клиента (Retail)") ' 4
          Column_PA = ColumnByName(ReportName_String, "Лист1", 1, "PA") ' 5
          Column_DateOfActiv = ColumnByName(ReportName_String, "Лист1", 1, "Date of activ") ' 6
          Column_Выдача_PA = ColumnByName(ReportName_String, "Лист1", 1, "Выдача PA") ' 7
          Column_PA_KK = ColumnByName(ReportName_String, "Лист1", 1, "PA_KK") ' 8
          Column_Выдача_РА_КК = ColumnByName(ReportName_String, "Лист1", 1, "Выдача_РА_КК") ' 9
          Column_chan = ColumnByName(ReportName_String, "Лист1", 1, "chan") ' 10
          Column_ФИО_сотрудника = ColumnByName(ReportName_String, "Лист1", 1, "ФИО сотрудника") ' 11
          Column_РегОфис = ColumnByName(ReportName_String, "Лист1", 1, "Рег. офис") ' 12
          Column_Допофис = ColumnByName(ReportName_String, "Лист1", 1, "Доп. офис") ' 13
          
          rowCount_Лист1 = 2
          Do While Not IsEmpty(Workbooks(ReportName_String).Sheets("Лист1").Cells(rowCount_Лист1, 1).Value)
            
            ' Обработка Тюменских Capacity.Лист1
            If Workbooks(ReportName_String).Sheets("Лист1").Cells(rowCount_Лист1, Column_РегОфис).Value = "Тюменский ОО1" Then
              
              ' Если это PA или PA КК
              If (Workbooks(ReportName_String).Sheets("Лист1").Cells(rowCount_Лист1, Column_PA).Value = "1") Or (Workbooks(ReportName_String).Sheets("Лист1").Cells(rowCount_Лист1, Column_PA_KK).Value = "1") Then
              
                ' Вставляем клиента: ' Поля: ID_клиента_Retail, PA, DateOfActiv, Выдача_PA, PA_KK, Выдача_РА_КК, chan, ФИО_сотрудника, РегОфис, ДопОфис
                Call InsertRecordInBook(Dir(OutBookName), "Лист1", "ID_клиента_Retail", Workbooks(ReportName_String).Sheets("Лист1").Cells(rowCount_Лист1, Column_ID_клиента_Retail).Value, _
                                              "ID_клиента_Retail", Workbooks(ReportName_String).Sheets("Лист1").Cells(rowCount_Лист1, Column_ID_клиента_Retail).Value, _
                                                "PA", Workbooks(ReportName_String).Sheets("Лист1").Cells(rowCount_Лист1, Column_PA).Value, _
                                                  "DateOfActiv", Workbooks(ReportName_String).Sheets("Лист1").Cells(rowCount_Лист1, Column_DateOfActiv).Value, _
                                                    "Выдача_PA", Workbooks(ReportName_String).Sheets("Лист1").Cells(rowCount_Лист1, Column_Выдача_PA).Value, _
                                                      "PA_KK", Workbooks(ReportName_String).Sheets("Лист1").Cells(rowCount_Лист1, Column_PA_KK).Value, _
                                                        "Выдача_РА_КК", Workbooks(ReportName_String).Sheets("Лист1").Cells(rowCount_Лист1, Column_Выдача_РА_КК).Value, _
                                                          "chan", Workbooks(ReportName_String).Sheets("Лист1").Cells(rowCount_Лист1, Column_chan).Value, _
                                                            "ФИО_сотрудника", Workbooks(ReportName_String).Sheets("Лист1").Cells(rowCount_Лист1, Column_ФИО_сотрудника).Value, _
                                                              "РегОфис", Workbooks(ReportName_String).Sheets("Лист1").Cells(rowCount_Лист1, Column_РегОфис).Value, _
                                                                "ДопОфис", cityOfficeName(Workbooks(ReportName_String).Sheets("Лист1").Cells(rowCount_Лист1, Column_Допофис).Value), _
                                                                  "", "", _
                                                                    "", "", _
                                                                      "", "", _
                                                                        "", "", _
                                                                          "", "", _
                                                                            "", "", _
                                                                              "", "", _
                                                                                "", "", _
                                                                                  "", "", _
                                                                                    "", "")
                
                ' Вставляем информацию о наличии PA ПК и КК в BASE\Clients: Capacity_PA_Date, Capacity_PA, Capacity_PA_CC_Date, Capacity_PA_CC

                                                                                    

              End If ' Если это PA или PA КК
              
            End If ' Если это офис Тюменский
            
            rowCount_Лист1 = rowCount_Лист1 + 1
            Application.StatusBar = "Обработка PA: " + CStr(rowCount_Лист1) + "..."
          Loop

      
      ' ===========================================================================================================================
      

      ' Обработка 2: Итоги по офисам и по РОО (Pre-Approved ПК) (идем снизу вверх)
      rowCount = rowByValue(ThisWorkbook.Name, "Лист6", "Итого по РОО:", 100, 100) - 1
      
      ' Переменные итогов по ОО (Pre-Approved ПК)
      Клиенты = 0
      Клиенты_с_PA = 0
      Выдано_PA = 0
      
      Do While ThisWorkbook.Sheets("Лист6").Cells(rowCount, 2).Value <> ""
        
        ' Если в строке "OO"
        If InStr(ThisWorkbook.Sheets("Лист6").Cells(rowCount, 2).Value, "ОО") <> 0 Then
          ' Заносим итоги - Клиенты (РА_ПК) (Pre-Approved ПК)
          ThisWorkbook.Sheets("Лист6").Cells(rowCount, 14).Value = Клиенты
          ThisWorkbook.Sheets("Лист6").Cells(rowCount, 14).HorizontalAlignment = xlCenter
          ' Клиенты_с_PA (Pre-Approved ПК)
          ThisWorkbook.Sheets("Лист6").Cells(rowCount, 15).Value = Клиенты_с_PA
          ThisWorkbook.Sheets("Лист6").Cells(rowCount, 15).HorizontalAlignment = xlCenter
          ' Выдано_PA (Pre-Approved ПК)
          ThisWorkbook.Sheets("Лист6").Cells(rowCount, 16).Value = Выдано_PA
          ThisWorkbook.Sheets("Лист6").Cells(rowCount, 16).HorizontalAlignment = xlCenter
          ' Доля (Pre-Approved ПК)
          ThisWorkbook.Sheets("Лист6").Cells(rowCount, 17).Value = РассчетДоли(ThisWorkbook.Sheets("Лист6").Cells(rowCount, 15).Value, ThisWorkbook.Sheets("Лист6").Cells(rowCount, 16).Value, 3)
          ' ThisWorkbook.Sheets("Лист6").Cells(RowCount, 17).NumberFormat = "0.0%"
          ThisWorkbook.Sheets("Лист6").Cells(rowCount, 17).NumberFormat = cellsNumberFormat(ThisWorkbook.Sheets("Лист6").Cells(rowCount, 17).Value)
          ThisWorkbook.Sheets("Лист6").Cells(rowCount, 17).HorizontalAlignment = xlCenter
          
          ' Обнуляем переменные (Pre-Approved ПК)
          Клиенты = 0
          Клиенты_с_PA = 0
          Выдано_PA = 0
        
        Else
        
          ' Суммируем (Pre-Approved ПК)
          Клиенты = Клиенты + ThisWorkbook.Sheets("Лист6").Cells(rowCount, 14).Value
          Клиенты_с_PA = Клиенты_с_PA + ThisWorkbook.Sheets("Лист6").Cells(rowCount, 15).Value
          Выдано_PA = Выдано_PA + ThisWorkbook.Sheets("Лист6").Cells(rowCount, 16).Value
          
        End If
        
        ' Следующая запись (Pre-Approved ПК)
        rowCount = rowCount - 1
        Application.StatusBar = "Сотрудники: " + CStr(rowCount) + "..."
        DoEventsInterval (rowCount)
      Loop
           
      ' Итоги по РОО
      ' Клиенты (Pre-Approved ПК)
      ThisWorkbook.Sheets("Лист6").Cells(row_Итого_по_РОО, 14).Formula = "=SUM(" + Replace(cellSum, "C", "N") + ")"
      ThisWorkbook.Sheets("Лист6").Cells(row_Итого_по_РОО, 14).HorizontalAlignment = xlCenter
      ' Клиенты с PA (Pre-Approved ПК)
      ThisWorkbook.Sheets("Лист6").Cells(row_Итого_по_РОО, 15).Formula = "=SUM(" + Replace(cellSum, "C", "O") + ")"
      ThisWorkbook.Sheets("Лист6").Cells(row_Итого_по_РОО, 15).HorizontalAlignment = xlCenter
      ' Выдано PA (Pre-Approved ПК)
      ThisWorkbook.Sheets("Лист6").Cells(row_Итого_по_РОО, 16).Formula = "=SUM(" + Replace(cellSum, "C", "P") + ")"
      ThisWorkbook.Sheets("Лист6").Cells(row_Итого_по_РОО, 16).HorizontalAlignment = xlCenter
      ' Доля (Pre-Approved ПК)
      ThisWorkbook.Sheets("Лист6").Cells(row_Итого_по_РОО, 17).Value = РассчетДоли(ThisWorkbook.Sheets("Лист6").Cells(row_Итого_по_РОО, 15).Value, ThisWorkbook.Sheets("Лист6").Cells(row_Итого_по_РОО, 16).Value, 3)
      ' ThisWorkbook.Sheets("Лист6").Cells(row_Итого_по_РОО, 17).NumberFormat = "0.0%"
      ThisWorkbook.Sheets("Лист6").Cells(row_Итого_по_РОО, 17).NumberFormat = cellsNumberFormat(ThisWorkbook.Sheets("Лист6").Cells(row_Итого_по_РОО, 17).Value)
      ThisWorkbook.Sheets("Лист6").Cells(row_Итого_по_РОО, 17).HorizontalAlignment = xlCenter
      
      
      ' Цикл №4 ДК на входящем потоке + Цикл №5 Пенсионеры на входящем потоке ---
      
      ' Открытие сводных таблиц в Capacity на Листе "Дет. Продаж ДК"
      Call openPivotTables_Capacity_Дет_Продаж_ДК(ReportName_String)
      
      ' Обработка 1 (Дет. Продаж ДК)
      ' "Названия строк"
      Row_ДетПродажДК_НазванияCтрок = rowByValue(ReportName_String, "Дет. Продаж ДК", "Названия строк", 100, 100)
      
      ' Определяем номера столбцов:
      ' Кол-во клиентов (Дет. Продаж ДК) Column_ДетПродажДК_КоличествоКлиентов
      Column_ДетПродажДК_КоличествоКлиентов = ColumnByValue(ReportName_String, "Дет. Продаж ДК", "Количество клиентов", 100, 100)
      
      ' Без ДК (Дет. Продаж ДК) Column_ДетПродажДК_НетDK
      Column_ДетПродажДК_НетDK = ColumnByValue(ReportName_String, "Дет. Продаж ДК", "Нет DK", 100, 100)
      
      ' Заявки ДК (Дет. Продаж ДК) Column_ДетПродажДК_ЗаявокНаДК
      Column_ДетПродажДК_ЗаявокНаДК = ColumnByValue(ReportName_String, "Дет. Продаж ДК", "Заявок на ДК", 100, 100)
      
      ' Количество_пенсионеров Column_ДетПродажДК_КоличествоПенсионеров
      Column_ДетПродажДК_КоличествоПенсионеров = ColumnByValue(ReportName_String, "Дет. Продаж ДК", "Количество пенсионеров", 100, 100)
      
      ' Пенсионеры_Без_ДК Column_ДетПродажДК_УпущенныеПенсионеры
      Column_ДетПродажДК_УпущенныеПенсионеры = ColumnByValue(ReportName_String, "Дет. Продаж ДК", "Упущенные пенсионеры", 100, 100)
      
      ' Пенсионеры_выдано_ДК Column_ДетПродажДК_ВыданоМоментальныхПенсионныхКарт
      Column_ДетПродажДК_ВыданоМоментальныхПенсионныхКарт = ColumnByValue(ReportName_String, "Дет. Продаж ДК", "Выдано моментальных пенсионных карт", 100, 100)
      
      ' Пенсионеры_Заявл_ПФР Column_ДетПродажДК_НаправленоЗаявленийВПФР
      Column_ДетПродажДК_НаправленоЗаявленийВПФР = ColumnByValue(ReportName_String, "Дет. Продаж ДК", "Направлено заявлений в ПФР", 100, 100)
      
      ' Начинаем со строки "Названия строк" +1
      ' RowCount = RowByValue(ReportName_String, "Дет. Продаж ДК", "Названия строк", 100, 100) + 1
      rowCount = Row_ДетПродажДК_НазванияCтрок + 1
      
      Do While InStr(Workbooks(ReportName_String).Sheets("Дет. Продаж ДК").Cells(rowCount, 1).Value, "Общий итог") = 0
        
        ' Если в ячейки нет символов: ", ф-л, -, "", Москва, 0, 1, 2, 3, 4, 5, 6, 7, 8, 9
        ' (InStr(currentNameInCapacityStr, "0") = 0)And(InStr(currentNameInCapacityStr, "0") = 0)And(InStr(currentNameInCapacityStr, "1") = 0)And(InStr(currentNameInCapacityStr, "2") = 0)And(InStr(currentNameInCapacityStr, "3") = 0)And(InStr(currentNameInCapacityStr, "4") = 0)And(InStr(currentNameInCapacityStr, "5") = 0)And(InStr(currentNameInCapacityStr, "6") = 0)And(InStr(currentNameInCapacityStr, "7") = 0)And(InStr(currentNameInCapacityStr, "8") = 0)And(InStr(currentNameInCapacityStr, "9") = 0)
        currentNameInCapacityStr = Workbooks(ReportName_String).Sheets("Дет. Продаж ДК").Cells(rowCount, 1).Value
        
        ' Проверяем содержимое ячейки
        If (currentNameInCapacityStr <> "") And (InStr(currentNameInCapacityStr, Chr(34)) = 0) And (InStr(currentNameInCapacityStr, "ф-л") = 0) And (InStr(currentNameInCapacityStr, "-") = 0) And (InStr(currentNameInCapacityStr, "Москва") = 0) And (InStr(currentNameInCapacityStr, "0") = 0) And (InStr(currentNameInCapacityStr, "0") = 0) And (InStr(currentNameInCapacityStr, "1") = 0) And (InStr(currentNameInCapacityStr, "2") = 0) And (InStr(currentNameInCapacityStr, "3") = 0) And (InStr(currentNameInCapacityStr, "4") = 0) And (InStr(currentNameInCapacityStr, "5") = 0) And (InStr(currentNameInCapacityStr, "6") = 0) And (InStr(currentNameInCapacityStr, "7") = 0) And (InStr(currentNameInCapacityStr, "8") = 0) And (InStr(currentNameInCapacityStr, "9") = 0) Then
          
          ' Если это ФИО, то проверяем в подстроке всех сотрудников Тюмени
          currentNameInCapacityStr = Фамилия_и_Имя(currentNameInCapacityStr, 3)
          
          If InStr(allNameStr, currentNameInCapacityStr) <> 0 Then
            '
            currentNameStr_Range = RangeByValue(ThisWorkbook.Name, "Лист6", currentNameInCapacityStr, 100, 100)
            currentNameStr_Row = ThisWorkbook.Sheets("Лист6").Range(currentNameStr_Range).Row
            currentNameStr_Column = ThisWorkbook.Sheets("Лист6").Range(currentNameStr_Range).Column
            
            ' Заносим данные по сотруднику
            ' Кол-во клиентов (Дет. Продаж ДК)
            ThisWorkbook.Sheets("Лист6").Cells(currentNameStr_Row, currentNameStr_Column + 16).Value = Workbooks(ReportName_String).Sheets("Дет. Продаж ДК").Cells(rowCount, Column_ДетПродажДК_КоличествоКлиентов).Value
            ThisWorkbook.Sheets("Лист6").Cells(currentNameStr_Row, currentNameStr_Column + 16).HorizontalAlignment = xlRight
            ' Без ДК (Дет. Продаж ДК)
            ThisWorkbook.Sheets("Лист6").Cells(currentNameStr_Row, currentNameStr_Column + 17).Value = Workbooks(ReportName_String).Sheets("Дет. Продаж ДК").Cells(rowCount, Column_ДетПродажДК_НетDK).Value
            ThisWorkbook.Sheets("Лист6").Cells(currentNameStr_Row, currentNameStr_Column + 17).HorizontalAlignment = xlRight
            ' Заявки ДК (Дет. Продаж ДК)
            ThisWorkbook.Sheets("Лист6").Cells(currentNameStr_Row, currentNameStr_Column + 18).Value = Workbooks(ReportName_String).Sheets("Дет. Продаж ДК").Cells(rowCount, Column_ДетПродажДК_ЗаявокНаДК).Value
            ThisWorkbook.Sheets("Лист6").Cells(currentNameStr_Row, currentNameStr_Column + 18).HorizontalAlignment = xlRight
            ' Доля (Дет. Продаж ДК)
            ThisWorkbook.Sheets("Лист6").Cells(currentNameStr_Row, currentNameStr_Column + 19).Value = РассчетДоли(ThisWorkbook.Sheets("Лист6").Cells(currentNameStr_Row, currentNameStr_Column + 17).Value, ThisWorkbook.Sheets("Лист6").Cells(currentNameStr_Row, currentNameStr_Column + 18).Value, 3)
            ' ThisWorkbook.Sheets("Лист6").Cells(currentNameStr_Row, currentNameStr_Column + 19).NumberFormat = "0.0%"
            ThisWorkbook.Sheets("Лист6").Cells(currentNameStr_Row, currentNameStr_Column + 19).NumberFormat = cellsNumberFormat(ThisWorkbook.Sheets("Лист6").Cells(currentNameStr_Row, currentNameStr_Column + 19).Value)
            ThisWorkbook.Sheets("Лист6").Cells(currentNameStr_Row, currentNameStr_Column + 19).HorizontalAlignment = xlRight
            
            ' Окраска ячейки СФЕТОФОР
            Call Full_Color_RangeII("Лист6", currentNameStr_Row, currentNameStr_Column + 19, (РассчетДоли(ThisWorkbook.Sheets("Лист6").Cells(currentNameStr_Row, currentNameStr_Column + 17).Value, ThisWorkbook.Sheets("Лист6").Cells(currentNameStr_Row, currentNameStr_Column + 18).Value, 3) * 100), 15)
         
            ' Количество_пенсионеров
            ThisWorkbook.Sheets("Лист6").Cells(currentNameStr_Row, currentNameStr_Column + 20).Value = Workbooks(ReportName_String).Sheets("Дет. Продаж ДК").Cells(rowCount, Column_ДетПродажДК_КоличествоПенсионеров).Value
            ThisWorkbook.Sheets("Лист6").Cells(currentNameStr_Row, currentNameStr_Column + 20).HorizontalAlignment = xlRight
            
            ' Пенсионеры_Без_ДК
            ThisWorkbook.Sheets("Лист6").Cells(currentNameStr_Row, currentNameStr_Column + 21).Value = Workbooks(ReportName_String).Sheets("Дет. Продаж ДК").Cells(rowCount, Column_ДетПродажДК_УпущенныеПенсионеры).Value
            ThisWorkbook.Sheets("Лист6").Cells(currentNameStr_Row, currentNameStr_Column + 21).HorizontalAlignment = xlRight
            
            ' Пенсионеры_выдано_ДК
            ThisWorkbook.Sheets("Лист6").Cells(currentNameStr_Row, currentNameStr_Column + 22).Value = Workbooks(ReportName_String).Sheets("Дет. Продаж ДК").Cells(rowCount, Column_ДетПродажДК_ВыданоМоментальныхПенсионныхКарт).Value
            ThisWorkbook.Sheets("Лист6").Cells(currentNameStr_Row, currentNameStr_Column + 22).HorizontalAlignment = xlRight
            
            ' Доля
            ThisWorkbook.Sheets("Лист6").Cells(currentNameStr_Row, currentNameStr_Column + 23).Value = РассчетДоли(ThisWorkbook.Sheets("Лист6").Cells(currentNameStr_Row, currentNameStr_Column + 21).Value, ThisWorkbook.Sheets("Лист6").Cells(currentNameStr_Row, currentNameStr_Column + 22).Value, 3)
            ' ThisWorkbook.Sheets("Лист6").Cells(currentNameStr_Row, currentNameStr_Column + 23).NumberFormat = "0.0%"
            ThisWorkbook.Sheets("Лист6").Cells(currentNameStr_Row, currentNameStr_Column + 23).NumberFormat = cellsNumberFormat(ThisWorkbook.Sheets("Лист6").Cells(currentNameStr_Row, currentNameStr_Column + 23).Value)
            ThisWorkbook.Sheets("Лист6").Cells(currentNameStr_Row, currentNameStr_Column + 23).HorizontalAlignment = xlRight
            
            ' Окраска ячейки СФЕТОФОР
            Call Full_Color_RangeII("Лист6", currentNameStr_Row, currentNameStr_Column + 23, (РассчетДоли(ThisWorkbook.Sheets("Лист6").Cells(currentNameStr_Row, currentNameStr_Column + 21).Value, ThisWorkbook.Sheets("Лист6").Cells(currentNameStr_Row, currentNameStr_Column + 22).Value, 3) * 100), 15)
            
            ' Пенсионеры_Заявл_ПФР
            ThisWorkbook.Sheets("Лист6").Cells(currentNameStr_Row, currentNameStr_Column + 24).Value = Workbooks(ReportName_String).Sheets("Дет. Продаж ДК").Cells(rowCount, Column_ДетПродажДК_НаправленоЗаявленийВПФР).Value
            
            ' Ставим ноль, если не было заявления, так как в Капасити пустая ячейка
            If IsEmpty(ThisWorkbook.Sheets("Лист6").Cells(currentNameStr_Row, currentNameStr_Column + 24).Value) Then
              ThisWorkbook.Sheets("Лист6").Cells(currentNameStr_Row, currentNameStr_Column + 24).Value = 0
            End If
            
            ThisWorkbook.Sheets("Лист6").Cells(currentNameStr_Row, currentNameStr_Column + 24).HorizontalAlignment = xlRight
            
            ' Доля ФАКТ по пенсионерам ПФР
            ThisWorkbook.Sheets("Лист6").Cells(currentNameStr_Row, currentNameStr_Column + 25).Value = РассчетДоли(ThisWorkbook.Sheets("Лист6").Cells(currentNameStr_Row, currentNameStr_Column + 20).Value, ThisWorkbook.Sheets("Лист6").Cells(currentNameStr_Row, currentNameStr_Column + 24).Value, 3)
            ThisWorkbook.Sheets("Лист6").Cells(currentNameStr_Row, currentNameStr_Column + 25).NumberFormat = cellsNumberFormat(ThisWorkbook.Sheets("Лист6").Cells(currentNameStr_Row, currentNameStr_Column + 25).Value)
            ThisWorkbook.Sheets("Лист6").Cells(currentNameStr_Row, currentNameStr_Column + 25).HorizontalAlignment = xlRight
            ' Окраска ячейки СФЕТОФОР - вариант 1 (красим в красный всех, кто с нулями)
            ' If ThisWorkbook.Sheets("Лист6").Cells(currentNameStr_Row, currentNameStr_Column + 25).Value = 0 Then
            '   Call Full_Color_RangeII("Лист6", currentNameStr_Row, currentNameStr_Column + 25, (РассчетДоли(ThisWorkbook.Sheets("Лист6").Cells(currentNameStr_Row, currentNameStr_Column + 20).Value, ThisWorkbook.Sheets("Лист6").Cells(currentNameStr_Row, currentNameStr_Column + 24).Value, 3) * 100), 20)
            ' End If
            ' Окраска ячейки СФЕТОФОР - вариант 2 (сфетофор на цель 20%)
            Call Full_Color_RangeII("Лист6", currentNameStr_Row, currentNameStr_Column + 25, (РассчетДоли(ThisWorkbook.Sheets("Лист6").Cells(currentNameStr_Row, currentNameStr_Column + 20).Value, ThisWorkbook.Sheets("Лист6").Cells(currentNameStr_Row, currentNameStr_Column + 24).Value, 3) * 100), 20)
            
       
           End If
        End If
        
        ' Если в ячейке "Тюменский ОО1", то запоминаем позицию для последующей обработки упущенных клиентов
        If InStr(Workbooks(ReportName_String).Sheets("Дет. Продаж ДК").Cells(rowCount, 1).Value, "Тюменский ОО1") <> 0 Then
          
          ' Смотрим упущенных пенсионеров для ПФР?
          ' row_Тюменский_ОО1 = RowCount
                  
          ' Workbooks(ReportName_String).Activate
          ' Workbooks(ReportName_String).Sheets("Дет. Продаж ДК").Cells(RowCount, 3).Select  ' было 1
          ' Workbooks(ReportName_String).Sheets("Дет. Продаж ДК").Cells(RowCount, 3).ShowDetail = True
          ' (Pre-Approved ПК)
          ' Range("Таблица2[[#Headers],[PA]]").Select
          ' ActiveSheet.ListObjects("Таблица2").Range.AutoFilter Field:=4, Criteria1:="1"
                                        
          ' Переходим на окно DB
          ' ThisWorkbook.Sheets("Лист6").Activate

          ' Столбец "Количество пенсионеров" - Раскрываем сводную таблицу. После раскрытия в Книге должен появиться новый "Лист3"
          Workbooks(ReportName_String).Sheets("Дет. Продаж ДК").Cells(rowCount, 11).ShowDetail = True
          ThisWorkbook.Sheets("Лист6").Activate
          
          ' Обработка пенсионеров на Лист1. Поля: "ID клиента (Retail)", "Дата поручения" (Integer), "Доп. офис", "ФИО сотрудника", "Пенсионер", "имеет пенс карту", "Актив пенс карта"
          ' Поля в BASE\Clients: Номер_клиента, Офис (Тюменский, Сургутский, Нижневартовский, Новоуренгойский, Тарко-Сале), Capacity_pens_Date, Capacity_pensioner, ФИО_сотрудника, Имеет_пенс_карту, Актив_пенс_карта
          ' Обработка Лист1 с занесением в таблицу Клиентов
          
          ' Столбцы - сделать замену на поиск в сводной таблице номера столбца по имени
          Column_Рег_офис = 5
          Column_ID_клиента_Retail = 13
          Column_Дата_поручения = 4
          Column_Доп_офис = 6
          Column_ФИО_сотрудника = 10
          Column_Пенсионер = 14
          Column_имеет_пенс_карту = 16
          Column_Актив_пенс_карта = 17
          
          RowCount_Лист3 = 2
          Do While Not IsEmpty(Workbooks(ReportName_String).Sheets("Лист3").Cells(RowCount_Лист3, 1).Value)
            
            ' Обработка пенсионеров на Capacity.Лист1
            If Workbooks(ReportName_String).Sheets("Лист3").Cells(RowCount_Лист3, Column_Пенсионер).Value = 1 Then
              
              ' Вставляем пенсионера в BASE\Clients
              Call InsertRecordInBook("Clients", "Лист1", "Номер_клиента", Workbooks(ReportName_String).Sheets("Лист3").Cells(RowCount_Лист3, Column_ID_клиента_Retail).Value, _
                                            "Номер_клиента", Workbooks(ReportName_String).Sheets("Лист3").Cells(RowCount_Лист3, Column_ID_клиента_Retail).Value, _
                                              "Офис", cityOfficeName(Workbooks(ReportName_String).Sheets("Лист3").Cells(RowCount_Лист3, Column_Доп_офис).Value), _
                                                "Capacity_pens_Date", CDate(Workbooks(ReportName_String).Sheets("Лист3").Cells(RowCount_Лист3, Column_Дата_поручения).Value), _
                                                  "Capacity_pensioner", "1", _
                                                    "ФИО_сотрудника", Фамилия_и_Имя(Workbooks(ReportName_String).Sheets("Лист3").Cells(RowCount_Лист3, Column_ФИО_сотрудника).Value, 3), _
                                                      "Имеет_пенс_карту", Workbooks(ReportName_String).Sheets("Лист3").Cells(RowCount_Лист3, Column_имеет_пенс_карту).Value, _
                                                        "Актив_пенс_карта", Workbooks(ReportName_String).Sheets("Лист3").Cells(RowCount_Лист3, Column_Актив_пенс_карта).Value, _
                                                          "", "", _
                                                            "", "", _
                                                              "", "", _
                                                                "", "", _
                                                                  "", "", _
                                                                    "", "", _
                                                                      "", "", _
                                                                        "", "", _
                                                                          "", "", _
                                                                            "", "", _
                                                                              "", "", _
                                                                                "", "", _
                                                                                  "", "")

              
            End If ' Если это пенсионер
            
            RowCount_Лист3 = RowCount_Лист3 + 1
            Application.StatusBar = "Обработка пенсионеров: " + CStr(RowCount_Лист3) + "..."
          Loop
          

        End If ' Если в ячейке "Тюменский ОО1", то запоминаем позицию для последующей обработки упущенных клиентов
        
        ' Следующая запись
        rowCount = rowCount + 1
        Application.StatusBar = "Сотрудники: " + CStr(rowCount) + "..."
        DoEventsInterval (rowCount)
      
      Loop
      

      ' Обработка 2: Итоги по офисам и по РОО (Дет. Продаж ДК) (идем снизу вверх)
      rowCount = rowByValue(ThisWorkbook.Name, "Лист6", "Итого по РОО:", 100, 100) - 1
      
      ' Переменные итогов по ОО (Дет. Продаж ДК)
      Клиенты = 0
      Клиенты_Без_ДК = 0
      Заявки_ДК = 0
      '
      Количество_пенсионеров = 0
      Пенсионеры_Без_ДК = 0
      Пенсионеры_выдано_ДК = 0
      Пенсионеры_Заявл_ПФР = 0

      Do While ThisWorkbook.Sheets("Лист6").Cells(rowCount, 2).Value <> ""
        
        ' Если в строке "OO"
        If InStr(ThisWorkbook.Sheets("Лист6").Cells(rowCount, 2).Value, "ОО") <> 0 Then
          
          ' Заносим итоги - Клиенты (Дет. Продаж ДК)
          ThisWorkbook.Sheets("Лист6").Cells(rowCount, 18).Value = Клиенты
          ThisWorkbook.Sheets("Лист6").Cells(rowCount, 18).HorizontalAlignment = xlCenter
          
          ' Клиенты_Без_ДК (Дет. Продаж ДК)
          ThisWorkbook.Sheets("Лист6").Cells(rowCount, 19).Value = Клиенты_Без_ДК
          ThisWorkbook.Sheets("Лист6").Cells(rowCount, 19).HorizontalAlignment = xlCenter
          
          ' Заявки_ДК (Дет. Продаж ДК)
          ThisWorkbook.Sheets("Лист6").Cells(rowCount, 20).Value = Заявки_ДК
          ThisWorkbook.Sheets("Лист6").Cells(rowCount, 20).HorizontalAlignment = xlCenter
          
          ' Доля (Дет. Продаж ДК)
          ThisWorkbook.Sheets("Лист6").Cells(rowCount, 21).Value = РассчетДоли(ThisWorkbook.Sheets("Лист6").Cells(rowCount, 19).Value, ThisWorkbook.Sheets("Лист6").Cells(rowCount, 20).Value, 3)
          ' ThisWorkbook.Sheets("Лист6").Cells(RowCount, 21).NumberFormat = "0.0%"
          ThisWorkbook.Sheets("Лист6").Cells(rowCount, 21).NumberFormat = cellsNumberFormat(ThisWorkbook.Sheets("Лист6").Cells(rowCount, 21).Value)
          ThisWorkbook.Sheets("Лист6").Cells(rowCount, 21).HorizontalAlignment = xlCenter
          
          ' Количество_пенсионеров
          ThisWorkbook.Sheets("Лист6").Cells(rowCount, 22).Value = Количество_пенсионеров
          ThisWorkbook.Sheets("Лист6").Cells(rowCount, 22).HorizontalAlignment = xlCenter
          
          ' Пенсионеры_Без_ДК
          ThisWorkbook.Sheets("Лист6").Cells(rowCount, 23).Value = Пенсионеры_Без_ДК
          ThisWorkbook.Sheets("Лист6").Cells(rowCount, 23).HorizontalAlignment = xlCenter
          
          ' Пенсионеры_выдано_ДК
          ThisWorkbook.Sheets("Лист6").Cells(rowCount, 24).Value = Пенсионеры_выдано_ДК
          ThisWorkbook.Sheets("Лист6").Cells(rowCount, 24).HorizontalAlignment = xlCenter
          
          ' Доля
          ThisWorkbook.Sheets("Лист6").Cells(rowCount, 25).Value = РассчетДоли(ThisWorkbook.Sheets("Лист6").Cells(rowCount, 23).Value, ThisWorkbook.Sheets("Лист6").Cells(rowCount, 24).Value, 3)
          ' ThisWorkbook.Sheets("Лист6").Cells(RowCount, 25).NumberFormat = "0.0%"
          ThisWorkbook.Sheets("Лист6").Cells(rowCount, 25).NumberFormat = cellsNumberFormat(ThisWorkbook.Sheets("Лист6").Cells(rowCount, 25).Value)
          ThisWorkbook.Sheets("Лист6").Cells(rowCount, 25).HorizontalAlignment = xlCenter
          
          ' Пенсионеры_Заявл_ПФР
          ThisWorkbook.Sheets("Лист6").Cells(rowCount, 26).Value = Пенсионеры_Заявл_ПФР
          ThisWorkbook.Sheets("Лист6").Cells(rowCount, 26).HorizontalAlignment = xlCenter
          
          ' Факт ПФР
          ThisWorkbook.Sheets("Лист6").Cells(rowCount, 27).Value = РассчетДоли(ThisWorkbook.Sheets("Лист6").Cells(rowCount, 22).Value, ThisWorkbook.Sheets("Лист6").Cells(rowCount, 26).Value, 3)
          ThisWorkbook.Sheets("Лист6").Cells(rowCount, 27).NumberFormat = cellsNumberFormat(ThisWorkbook.Sheets("Лист6").Cells(rowCount, 27).Value)
          ThisWorkbook.Sheets("Лист6").Cells(rowCount, 27).HorizontalAlignment = xlCenter
       
          ' Обнуляем переменные (Дет. Продаж ДК)
          Клиенты = 0
          Клиенты_Без_ДК = 0
          Заявки_ДК = 0
          '
          Количество_пенсионеров = 0
          Пенсионеры_Без_ДК = 0
          Пенсионеры_выдано_ДК = 0
          Пенсионеры_Заявл_ПФР = 0
        
        Else
        
          ' Суммируем (Дет. Продаж ДК)
          Клиенты = Клиенты + ThisWorkbook.Sheets("Лист6").Cells(rowCount, 18).Value
          Клиенты_Без_ДК = Клиенты_Без_ДК + ThisWorkbook.Sheets("Лист6").Cells(rowCount, 19).Value
          Заявки_ДК = Заявки_ДК + ThisWorkbook.Sheets("Лист6").Cells(rowCount, 20).Value
          '
          Количество_пенсионеров = Количество_пенсионеров + ThisWorkbook.Sheets("Лист6").Cells(rowCount, 22).Value
          Пенсионеры_Без_ДК = Пенсионеры_Без_ДК + ThisWorkbook.Sheets("Лист6").Cells(rowCount, 23).Value
          Пенсионеры_выдано_ДК = Пенсионеры_выдано_ДК + ThisWorkbook.Sheets("Лист6").Cells(rowCount, 24).Value
          Пенсионеры_Заявл_ПФР = Пенсионеры_Заявл_ПФР + ThisWorkbook.Sheets("Лист6").Cells(rowCount, 26).Value

          
        End If
        
        ' Следующая запись (Дет. Продаж ДК)
        rowCount = rowCount - 1
        Application.StatusBar = "Сотрудники: " + CStr(rowCount) + "..."
        DoEventsInterval (rowCount)
      Loop
           
      ' Итоги по РОО
      ' Клиенты (Дет. Продаж ДК)
      ThisWorkbook.Sheets("Лист6").Cells(row_Итого_по_РОО, 18).Formula = "=SUM(" + Replace(cellSum, "C", "R") + ")"
      ThisWorkbook.Sheets("Лист6").Cells(row_Итого_по_РОО, 18).HorizontalAlignment = xlCenter
      
      ' Клиенты_Без_ДК (Дет. Продаж ДК)
      ThisWorkbook.Sheets("Лист6").Cells(row_Итого_по_РОО, 19).Formula = "=SUM(" + Replace(cellSum, "C", "S") + ")"
      ThisWorkbook.Sheets("Лист6").Cells(row_Итого_по_РОО, 19).HorizontalAlignment = xlCenter
      
      ' Заявки_ДК (Дет. Продаж ДК)
      ThisWorkbook.Sheets("Лист6").Cells(row_Итого_по_РОО, 20).Formula = "=SUM(" + Replace(cellSum, "C", "T") + ")"
      ThisWorkbook.Sheets("Лист6").Cells(row_Итого_по_РОО, 20).HorizontalAlignment = xlCenter
      
      ' Доля (Дет. Продаж ДК)
      ThisWorkbook.Sheets("Лист6").Cells(row_Итого_по_РОО, 21).Value = РассчетДоли(ThisWorkbook.Sheets("Лист6").Cells(row_Итого_по_РОО, 19).Value, ThisWorkbook.Sheets("Лист6").Cells(row_Итого_по_РОО, 20).Value, 3)
      ' ThisWorkbook.Sheets("Лист6").Cells(row_Итого_по_РОО, 21).NumberFormat = "0.0%"
      ThisWorkbook.Sheets("Лист6").Cells(row_Итого_по_РОО, 21).NumberFormat = cellsNumberFormat(ThisWorkbook.Sheets("Лист6").Cells(row_Итого_по_РОО, 21).Value)
      ThisWorkbook.Sheets("Лист6").Cells(row_Итого_по_РОО, 21).HorizontalAlignment = xlCenter
      
      ' Количество_пенсионеров
      ThisWorkbook.Sheets("Лист6").Cells(row_Итого_по_РОО, 22).Formula = "=SUM(" + Replace(cellSum, "C", "V") + ")"
      ThisWorkbook.Sheets("Лист6").Cells(row_Итого_по_РОО, 22).HorizontalAlignment = xlCenter
      
      ' Пенсионеры_Без_ДК
      ThisWorkbook.Sheets("Лист6").Cells(row_Итого_по_РОО, 23).Formula = "=SUM(" + Replace(cellSum, "C", "W") + ")"
      ThisWorkbook.Sheets("Лист6").Cells(row_Итого_по_РОО, 23).HorizontalAlignment = xlCenter
      
      ' Пенсионеры_выдано_ДК
      ThisWorkbook.Sheets("Лист6").Cells(row_Итого_по_РОО, 24).Formula = "=SUM(" + Replace(cellSum, "C", "X") + ")"
      ThisWorkbook.Sheets("Лист6").Cells(row_Итого_по_РОО, 24).HorizontalAlignment = xlCenter

      ' Доля
      ThisWorkbook.Sheets("Лист6").Cells(row_Итого_по_РОО, 25).Value = РассчетДоли(ThisWorkbook.Sheets("Лист6").Cells(row_Итого_по_РОО, 23).Value, ThisWorkbook.Sheets("Лист6").Cells(row_Итого_по_РОО, 24).Value, 3)
      ThisWorkbook.Sheets("Лист6").Cells(row_Итого_по_РОО, 25).NumberFormat = cellsNumberFormat(ThisWorkbook.Sheets("Лист6").Cells(row_Итого_по_РОО, 25).Value)
      ThisWorkbook.Sheets("Лист6").Cells(row_Итого_по_РОО, 25).HorizontalAlignment = xlCenter
      
      ' Пенсионеры_Заявл_ПФР
      ThisWorkbook.Sheets("Лист6").Cells(row_Итого_по_РОО, 26).Formula = "=SUM(" + Replace(cellSum, "C", "Z") + ")"
      ThisWorkbook.Sheets("Лист6").Cells(row_Итого_по_РОО, 26).HorizontalAlignment = xlCenter
      
      ' Факт Заявл_ПФР в %
      ThisWorkbook.Sheets("Лист6").Cells(row_Итого_по_РОО, 27).Value = РассчетДоли(ThisWorkbook.Sheets("Лист6").Cells(row_Итого_по_РОО, 22).Value, ThisWorkbook.Sheets("Лист6").Cells(row_Итого_по_РОО, 26).Value, 3)
      ThisWorkbook.Sheets("Лист6").Cells(row_Итого_по_РОО, 27).NumberFormat = cellsNumberFormat(ThisWorkbook.Sheets("Лист6").Cells(row_Итого_по_РОО, 25).Value)
      ThisWorkbook.Sheets("Лист6").Cells(row_Итого_по_РОО, 27).HorizontalAlignment = xlCenter
      
      ' --- Цикл №4 ДК на входящем потоке + Цикл №5 Пенсионеры на входящем потоке ---
      
      ' --- Цикл №5 Потенциал для Интернет-банка (ИБ) ---
      ' Не так! На вкладке "Клиенты (Кросс)" проходим по листу и находим куратора на каждой неделе и в каждом нахождении раскрываем список и обрабатываем клиентов с наличием IB
      ' Делаем так: 1) Переходим на "Клиенты (Кросс)"
      '             2) Находим строку rowCount_КлиентыКросс_Общий_итог "Общий итог" и раскрываем сводную таблицу в столбце Column_Колво_клиентов - это будет Лист4
      '             3) В таблице выбираем клиентов
      ' Если есть IB, но не активный, то это потенциал к реактивации
      ' Создаем пустой файл и в него выгружаем: 1) Клиентов без ИБ где не подулючен, и 2) где есть, но не активен IB
      ' так же заносим в файл Clients в поля: Capacity_IB, Capacity_IB_Active, Capacity_IB_DateUpdate

      
      ' Открываем сводную таблицу
      Workbooks(ReportName_String).Sheets("Клиенты (Кросс)").Cells(rowCount_КлиентыКросс_Общий_итог, Column_Колво_клиентов).ShowDetail = True
      ' Переходим на Лист6
      ThisWorkbook.Sheets("Лист6").Activate

      ' Создаем выходную книгу для выгрузки ИБ
      OutBookName_IB = ThisWorkbook.Path + "\Out\IB_" + strDDMMYYYY(dateReportCapacity) + ".xlsx"
      ' Вложение2
      ThisWorkbook.Sheets("Лист6").Range("AQ3").Value = OutBookName_IB
      ' Создать файл
      Call createBook_out_ИБ(OutBookName_IB)

      ' Определяем столбцы
      Column_Лист4_Регофис = ColumnByName(ReportName_String, "Лист4", 1, "Рег. офис") ' "Рег. офис"
      Column_Лист4_ID_клиента_Retail = ColumnByName(ReportName_String, "Лист4", 1, "ID клиента (Retail)") ' "ID клиента (Retail)"
      Column_Лист4_Допофис = ColumnByName(ReportName_String, "Лист4", 1, "Доп. офис") ' "Доп. офис"
      Column_Лист4_ФИО_сотрудника = ColumnByName(ReportName_String, "Лист4", 1, "ФИО сотрудника") ' "ФИО сотрудника"
      Column_Лист4_Тип_клиента2 = ColumnByName(ReportName_String, "Лист4", 1, "Тип клиента2") ' "Тип клиента2"
      Column_Лист4_IB = ColumnByName(ReportName_String, "Лист4", 1, "IB") ' "IB"
      Column_Лист4_Net_IB = ColumnByName(ReportName_String, "Лист4", 1, "Net IB") ' "Net IB"
      Column_Лист4_ИБ_активный = ColumnByName(ReportName_String, "Лист4", 1, "ИБ активный") ' "ИБ активный"

      ' Копируем клиентов ИБ с Лист4 в выходной файл
      RowCount_Лист4 = 2
      Do While Not IsEmpty(Workbooks(ReportName_String).Sheets("Лист4").Cells(RowCount_Лист4, 1).Value)
            
            ' Обработка Тюменских Capacity.Лист1
            If Workbooks(ReportName_String).Sheets("Лист4").Cells(RowCount_Лист4, Column_Лист4_Регофис).Value = "Тюменский ОО1" Then
              
                ' Потенциал реактивации ИБ
                If (Workbooks(ReportName_String).Sheets("Лист4").Cells(RowCount_Лист4, Column_Лист4_IB).Value = 1) And (Workbooks(ReportName_String).Sheets("Лист4").Cells(RowCount_Лист4, Column_Лист4_ИБ_активный).Value = 0) Then
                  Потенциал_реактивации_ИБ_Var = 1
                Else
                  Потенциал_реактивации_ИБ_Var = 0
                End If
              
                ' Вставляем клиента:
                Call InsertRecordInBook(Dir(OutBookName_IB), "Лист1", "ID_клиента_Retail", Workbooks(ReportName_String).Sheets("Лист4").Cells(RowCount_Лист4, Column_Лист4_ID_клиента_Retail).Value, _
                                              "ID_клиента_Retail", Workbooks(ReportName_String).Sheets("Лист4").Cells(RowCount_Лист4, Column_Лист4_ID_клиента_Retail).Value, _
                                                "ФИО", "", _
                                                  "Тип_клиента", Workbooks(ReportName_String).Sheets("Лист4").Cells(RowCount_Лист4, Column_Лист4_Тип_клиента2).Value, _
                                                    "ИБ", Workbooks(ReportName_String).Sheets("Лист4").Cells(RowCount_Лист4, Column_Лист4_IB).Value, _
                                                      "Нет ИБ", Workbooks(ReportName_String).Sheets("Лист4").Cells(RowCount_Лист4, Column_Лист4_Net_IB).Value, _
                                                        "ИБ активный", Workbooks(ReportName_String).Sheets("Лист4").Cells(RowCount_Лист4, Column_Лист4_ИБ_активный).Value, _
                                                          "Потенциал реактивации ИБ", Потенциал_реактивации_ИБ_Var, _
                                                            "ФИО_сотрудника", Фамилия_и_Имя(Workbooks(ReportName_String).Sheets("Лист4").Cells(RowCount_Лист4, Column_Лист4_ФИО_сотрудника).Value, 3), _
                                                              "РегОфис", Workbooks(ReportName_String).Sheets("Лист4").Cells(RowCount_Лист4, Column_Лист4_Регофис).Value, _
                                                                "ДопОфис", cityOfficeName(Workbooks(ReportName_String).Sheets("Лист4").Cells(RowCount_Лист4, Column_Лист4_Допофис).Value), _
                                                                  "", "", _
                                                                    "", "", _
                                                                      "", "", _
                                                                        "", "", _
                                                                          "", "", _
                                                                            "", "", _
                                                                              "", "", _
                                                                                "", "", _
                                                                                  "", "", _
                                                                                    "", "")
                
            End If ' Если это офис Тюменский
            
            RowCount_Лист4 = RowCount_Лист4 + 1
            Application.StatusBar = "Обработка IB: " + CStr(RowCount_Лист4) + "..."
            DoEventsInterval (RowCount_Лист4)
          Loop

           

      ' --- Конец Цикл №5 Потенциал для Интернет-банка (ИБ) ---
      
      ' Закрываем базу BASE\Clients
      CloseBook ("Clients")
      
      ' Закрываем выходную книгу с выгрузкой PA
      Workbooks(Dir(OutBookName)).Close SaveChanges:=True
      
      ' Закрываем выходную книгу с выгрузкой IB
      Workbooks(Dir(OutBookName_IB)).Close SaveChanges:=True
      
      ' Сохранение изменений
      ThisWorkbook.Save
    
      ' Переменная завершения обработки
      finishProcess = True
      
    Else
      ' Сообщение о неверном формате отчета или даты
      MsgBox ("Проверьте отчет: " + CheckFormatReportResult + "!")
    End If ' Проверка формы отчета

    ' Закрываем файл с отчетом без сохранения изменений (параметр SaveChanges:=False)
    Workbooks(Dir(FileName)).Close SaveChanges:=False
    
    ' Переходим на окно DB
    ' Переходим в ячейку M2
    ' ThisWorkbook.Activate
    ThisWorkbook.Sheets("Лист6").Activate ' !!!
    ThisWorkbook.Sheets("Лист6").Range("A1").Select
    
    ' Строка статуса
    Application.StatusBar = "Подготовка к копированию..."

    ' Копируем в исходящий файл
    Call copyDBToSend_Sheet6

    ' Строка статуса
    Application.StatusBar = ""


    ' Зачеркиваем пункт меню на стартовой страницы
    Call ЗачеркиваемТекстВячейке("Лист0", RangeByValue(ThisWorkbook.Name, "Лист0", "Отчет кросс-продажи Capacity Model", 100, 100))
    
    ' Отправка письма: отправляю шаблон самому себе для последующей отправки в сеть письма на его основе
    Call Отправка_Lotus_Notes_Лист6_Capacity
        
    ' Отправка письма с отработкой Pre-Approved_ДДММГГГГ
    Call Отправка_Lotus_Notes_Лист6_Pre_Approved
                
    ' Отправка письма с отработкой IB
    Call Отправка_Lotus_Notes_Лист6_IB
                
    ' Итоговое сообщение
    If finishProcess = True Then
      MsgBox ("Обработка " + Dir(ReportName_String) + " завершена!")
    Else
      MsgBox ("Обработка отчета была прервана!")
    End If

  End If ' Если файл был выбран

End Sub


' Открытие сводных таблиц в Capacity на Листе "Клиенты (Кросс)"
Sub openPivotTables_Capacity_Клиенты_Кросс(In_ReportName_String)
Dim список_открыт_Клиенты_Кросс As Boolean
        
  ' Workbooks(In_ReportName_String).Activate
  ' Sheets("Клиенты (Кросс)").Select
                
          ' Открываем все ячейки с "Валяев Сергей Николаевич" в столбце A (1)
          rowCount = 1
          список_открыт_Клиенты_Кросс = False
          Do While (Workbooks(In_ReportName_String).Sheets("Клиенты (Кросс)").Cells(rowCount, 1).Value <> "Общий итог") And (список_открыт_Клиенты_Кросс = False)
            ' Проверяем ячейку
            If (Trim(Workbooks(In_ReportName_String).Sheets("Клиенты (Кросс)").Cells(rowCount, 1).Value) = "Валяев Сергей Николаевич") Or (Trim(Workbooks(In_ReportName_String).Sheets("Клиенты (Кросс)").Cells(rowCount, 1).Value) = "Данилов Александр Сергеевич") Then
              
              ' Раскрываем сводную таблицу
              Workbooks(In_ReportName_String).Sheets("Клиенты (Кросс)").Cells(rowCount, 1).ShowDetail = True
              
              ' Открытие сводной таблицы
              ' ActiveSheet.PivotTables("SASApp:TEMP.CAP_CROSS").PivotFields("КураторРБ").PivotItems("Валяев Сергей Николаевич").ShowDetail = True
              ' Переменная открытия списка
              список_открыт_Клиенты_Кросс = True
            End If
            ' Следующая запись
            rowCount = rowCount + 1
          Loop

  ' Переходим на окно DB
  ThisWorkbook.Sheets("Лист6").Activate

End Sub

' Открытие сводных таблиц в Capacity на Листе "PA_KK"
Sub openPivotTables_Capacity_PA_KK(In_ReportName_String)

  ' Workbooks(In_ReportName_String).Activate
  ' Sheets("PA_KK").Select
      
      
      rowCount = 1
      
      Do While (Workbooks(In_ReportName_String).Sheets("PA_KK").Cells(rowCount, 1).Value <> "Общий итог")
        
        ' Тюменский ОО1
        If Workbooks(In_ReportName_String).Sheets("PA_KK").Cells(rowCount, 1).Value = "Тюменский ОО1" Then
          ' Раскрываем сводную таблицу
          Workbooks(In_ReportName_String).Sheets("PA_KK").Cells(rowCount, 1).ShowDetail = True
        End If
        
        ' Проверяем ячейку - если это Тюменский (без ОО1), Сургутский, Нижневартовский, Новоуренгойский, Тарко-Сале
        If Workbooks(In_ReportName_String).Sheets("PA_KK").Cells(rowCount, 1).Value = "ОО ""Тюменский""" Then
          ' Раскрываем сводную таблицу
          Workbooks(In_ReportName_String).Sheets("PA_KK").Cells(rowCount, 1).ShowDetail = True
          ' Открытие сводной таблицы
          ' ActiveSheet.PivotTables("SASApp:TEMP.CAPACITY_PA").PivotFields("Доп. офис").PivotItems("ОО ""Тюменский""").ShowDetail = True
        End If
        ' Сургутский
        If Workbooks(In_ReportName_String).Sheets("PA_KK").Cells(rowCount, 1).Value = "ОО2""Сургутский""" Then
          ' Раскрываем сводную таблицу
          Workbooks(In_ReportName_String).Sheets("PA_KK").Cells(rowCount, 1).ShowDetail = True
          ' Открытие сводной таблицы
          ' ActiveSheet.PivotTables("SASApp:TEMP.CAPACITY_PA").PivotFields("Доп. офис").PivotItems("ОО2""Сургутский""").ShowDetail = True
        End If
        ' Нижневартовский
        If Workbooks(In_ReportName_String).Sheets("PA_KK").Cells(rowCount, 1).Value = "ОО2 ""Нижневартовский""" Then
          ' Раскрываем сводную таблицу
          Workbooks(In_ReportName_String).Sheets("PA_KK").Cells(rowCount, 1).ShowDetail = True
          ' Открытие сводной таблицы
          ' ActiveSheet.PivotTables("SASApp:TEMP.CAPACITY_PA").PivotFields("Доп. офис").PivotItems("ОО2 ""Нижневартовский""").ShowDetail = True
        End If
        ' Новоуренгойский
        If Workbooks(In_ReportName_String).Sheets("PA_KK").Cells(rowCount, 1).Value = "ОО2""Новоуренгойский""" Then
          ' Раскрываем сводную таблицу
          Workbooks(In_ReportName_String).Sheets("PA_KK").Cells(rowCount, 1).ShowDetail = True
          ' Открытие сводной таблицы
          ' ActiveSheet.PivotTables("SASApp:TEMP.CAPACITY_PA").PivotFields("Доп. офис").PivotItems("ОО2""Новоуренгойский""").ShowDetail = True
        End If
        ' Тарко-Сале
        If Workbooks(In_ReportName_String).Sheets("PA_KK").Cells(rowCount, 1).Value = "ОО2 ""Тарко-Сале""" Then
          ' Раскрываем сводную таблицу
          Workbooks(In_ReportName_String).Sheets("PA_KK").Cells(rowCount, 1).ShowDetail = True
          ' Открытие сводной таблицы
          ' ActiveSheet.PivotTables("SASApp:TEMP.CAPACITY_PA").PivotFields("Доп. офис").PivotItems("ОО2 ""Тарко-Сале""").ShowDetail = True
        End If
                   
        ' Следующая запись
        rowCount = rowCount + 1
      Loop

  ' Переходим на окно DB
  ThisWorkbook.Sheets("Лист6").Activate

End Sub

' Открытие сводных таблиц в Capacity на Листе "Pre-Approved"
Sub openPivotTables_Capacity_PA_ПК(In_ReportName_String)

  ' Workbooks(In_ReportName_String).Activate
  ' Sheets("Pre-Approved").Select
      
      rowCount = 1
      ' список_открыт_Клиенты_Кросс = False
      Do While (Workbooks(In_ReportName_String).Sheets("Pre-Approved").Cells(rowCount, 1).Value <> "Общий итог")
        
        
        ' Открываем главный список "Тюменский ОО1"
        If Workbooks(In_ReportName_String).Sheets("Pre-Approved").Cells(rowCount, 1).Value = "Тюменский ОО1" Then
          
          ' Раскрываем сводную таблицу
          Workbooks(In_ReportName_String).Sheets("Pre-Approved").Cells(rowCount, 1).ShowDetail = True
          ' Открытие сводной таблицы
          ' ActiveSheet.PivotTables("SASApp:TEMP.CAPACITY_PA").PivotFields("Доп. офис").PivotItems("ОО ""Тюменский""").ShowDetail = True
          
        End If
        
        
        ' Проверяем ячейку - если это Тюменский (без ОО1), Сургутский, Нижневартовский, Новоуренгойский, Тарко-Сале
        If Workbooks(In_ReportName_String).Sheets("Pre-Approved").Cells(rowCount, 1).Value = "ОО ""Тюменский""" Then
          ' Раскрываем сводную таблицу
          Workbooks(In_ReportName_String).Sheets("Pre-Approved").Cells(rowCount, 1).ShowDetail = True
          ' Открытие сводной таблицы
          ' ActiveSheet.PivotTables("SASApp:TEMP.CAPACITY_PA").PivotFields("Доп. офис").PivotItems("ОО ""Тюменский""").ShowDetail = True
        End If
        ' Сургутский
        If Workbooks(In_ReportName_String).Sheets("Pre-Approved").Cells(rowCount, 1).Value = "ОО2""Сургутский""" Then
          ' Раскрываем сводную таблицу
          Workbooks(In_ReportName_String).Sheets("Pre-Approved").Cells(rowCount, 1).ShowDetail = True
          ' Открытие сводной таблицы
          ' ActiveSheet.PivotTables("SASApp:TEMP.CAPACITY_PA").PivotFields("Доп. офис").PivotItems("ОО2""Сургутский""").ShowDetail = True
        End If
        ' Нижневартовский
        If Workbooks(In_ReportName_String).Sheets("Pre-Approved").Cells(rowCount, 1).Value = "ОО2 ""Нижневартовский""" Then
          ' Раскрываем сводную таблицу
          Workbooks(In_ReportName_String).Sheets("Pre-Approved").Cells(rowCount, 1).ShowDetail = True
          ' Открытие сводной таблицы
          ' ActiveSheet.PivotTables("SASApp:TEMP.CAPACITY_PA").PivotFields("Доп. офис").PivotItems("ОО2 ""Нижневартовский""").ShowDetail = True
        End If
        ' Новоуренгойский
        If Workbooks(In_ReportName_String).Sheets("Pre-Approved").Cells(rowCount, 1).Value = "ОО2""Новоуренгойский""" Then
          ' Раскрываем сводную таблицу
          Workbooks(In_ReportName_String).Sheets("Pre-Approved").Cells(rowCount, 1).ShowDetail = True
          ' Открытие сводной таблицы
          ' ActiveSheet.PivotTables("SASApp:TEMP.CAPACITY_PA").PivotFields("Доп. офис").PivotItems("ОО2""Новоуренгойский""").ShowDetail = True
        End If
        ' Тарко-Сале
        If Workbooks(In_ReportName_String).Sheets("Pre-Approved").Cells(rowCount, 1).Value = "ОО2 ""Тарко-Сале""" Then
          ' Раскрываем сводную таблицу
          Workbooks(In_ReportName_String).Sheets("Pre-Approved").Cells(rowCount, 1).ShowDetail = True
          ' Открытие сводной таблицы
          ' ActiveSheet.PivotTables("SASApp:TEMP.CAPACITY_PA").PivotFields("Доп. офис").PivotItems("ОО2 ""Тарко-Сале""").ShowDetail = True
        End If
                   
        ' Следующая запись
        rowCount = rowCount + 1
      Loop

  ' Переходим на окно DB
  ThisWorkbook.Sheets("Лист6").Activate

End Sub

' Открытие сводной таблицы openPivotTables_Capacity_Дет_Продаж_ДК
Sub openPivotTables_Capacity_Дет_Продаж_ДК(In_ReportName_String)
Dim список_открыт_Клиенты_Кросс As Boolean
                
  ' Workbooks(In_ReportName_String).Activate
  ' Sheets("Дет. Продаж ДК").Select
                  
          ' Открываем все ячейки с "Валяев Сергей Николаевич" в столбце A (1)
          rowCount = 1
          список_открыт_Клиенты_Кросс = False
          Do While (Workbooks(In_ReportName_String).Sheets("Дет. Продаж ДК").Cells(rowCount, 1).Value <> "Общий итог") And (список_открыт_Клиенты_Кросс = False)
            
            ' Проверяем ячейку - Валяев Сергей Николаевич
            If (Trim(Workbooks(In_ReportName_String).Sheets("Дет. Продаж ДК").Cells(rowCount, 1).Value) = "Валяев Сергей Николаевич") Or ((Trim(Workbooks(In_ReportName_String).Sheets("Дет. Продаж ДК").Cells(rowCount, 1).Value) = "Данилов Александр Сергеевич")) Then
              ' Раскрываем сводную таблицу
              Workbooks(In_ReportName_String).Sheets("Дет. Продаж ДК").Cells(rowCount, 1).ShowDetail = True
              ' Открытие сводной таблицы
              ' ActiveSheet.PivotTables("SASApp:TEMP.DET_OF_CLIENT").PivotFields("КураторРБ").PivotItems("Валяев Сергей Николаевич").ShowDetail = True
            End If

            ' Проверяем ячейку 2 - Тюменский ОО1
            If Trim(Workbooks(In_ReportName_String).Sheets("Дет. Продаж ДК").Cells(rowCount, 1).Value) = "Тюменский ОО1" Then
              
              ' Раскрываем сводную таблицу
              Workbooks(In_ReportName_String).Sheets("Дет. Продаж ДК").Cells(rowCount, 1).ShowDetail = True
              ' ActiveSheet.PivotTables("SASApp:TEMP.DET_OF_CLIENT").PivotFields("Рег. офис").PivotItems("Тюменский ОО1").ShowDetail = True
              
              ' Переменная открытия списка
              ' пока убираем - список_открыт_Клиенты_Кросс = True
            
            End If
            
            ' Далее нужно по очереди открыть офисы
            ' ОО "Тюменский"
            If InStr(Workbooks(In_ReportName_String).Sheets("Дет. Продаж ДК").Cells(rowCount, 1).Value, "ОО ""Тюменский""") <> 0 Then
              
              ' Раскрываем сводную таблицу
              Workbooks(In_ReportName_String).Sheets("Дет. Продаж ДК").Cells(rowCount, 1).ShowDetail = True
              ' ActiveSheet.PivotTables("SASApp:TEMP.DET_OF_CLIENT").PivotFields("Доп. офис").PivotItems("ОО ""Тюменский""").ShowDetail = True
            
            End If
            
            ' ОО2"Сургутский"
            If InStr(Workbooks(In_ReportName_String).Sheets("Дет. Продаж ДК").Cells(rowCount, 1).Value, "ОО2""Сургутский""") <> 0 Then
              
              ' Раскрываем сводную таблицу
              Workbooks(In_ReportName_String).Sheets("Дет. Продаж ДК").Cells(rowCount, 1).ShowDetail = True
              ' ActiveSheet.PivotTables("SASApp:TEMP.DET_OF_CLIENT").PivotFields("Доп. офис").PivotItems("ОО2""Сургутский""").ShowDetail = True
            
            End If
            
            ' ОО2 "Нижневартовский"
            If InStr(Workbooks(In_ReportName_String).Sheets("Дет. Продаж ДК").Cells(rowCount, 1).Value, "ОО2 ""Нижневартовский""") <> 0 Then
              
              ' Раскрываем сводную таблицу
              Workbooks(In_ReportName_String).Sheets("Дет. Продаж ДК").Cells(rowCount, 1).ShowDetail = True
              ' ActiveSheet.PivotTables("SASApp:TEMP.DET_OF_CLIENT").PivotFields("Доп. офис").PivotItems("ОО2 ""Нижневартовский""").ShowDetail = True
            
            End If
            
            ' ОО2"Новоуренгойский"
            If InStr(Workbooks(In_ReportName_String).Sheets("Дет. Продаж ДК").Cells(rowCount, 1).Value, "ОО2""Новоуренгойский""") <> 0 Then
              
              ' Раскрываем сводную таблицу
              Workbooks(In_ReportName_String).Sheets("Дет. Продаж ДК").Cells(rowCount, 1).ShowDetail = True
              ' ActiveSheet.PivotTables("SASApp:TEMP.DET_OF_CLIENT").PivotFields("Доп. офис").PivotItems("ОО2""Новоуренгойский""").ShowDetail = True
            
            End If
            
            ' ОО2 "Тарко-Сале"
            If InStr(Workbooks(In_ReportName_String).Sheets("Дет. Продаж ДК").Cells(rowCount, 1).Value, "ОО2 ""Тарко-Сале""") <> 0 Then
              
              ' Раскрываем сводную таблицу
              Workbooks(In_ReportName_String).Sheets("Дет. Продаж ДК").Cells(rowCount, 1).ShowDetail = True
              ' ActiveSheet.PivotTables("SASApp:TEMP.DET_OF_CLIENT").PivotFields("Доп. офис").PivotItems("ОО2 ""Тарко-Сале""").ShowDetail = True
            
            End If
            
    
            ' Следующая запись
            rowCount = rowCount + 1
          Loop

  ' Переходим на окно DB
  ThisWorkbook.Sheets("Лист6").Activate

End Sub

      
' Получение даты отчета Capacity. Лист "Тип операции". Строка "ДО - Менеджер - Тип операции - Категория операции - Операция".
Function dateReportFromCapacity(In_Workbooks, In_Sheet) As Date
Dim строка_ДО_Менеджер, Итог_Column, ColumnCount As Byte

  ' Строка с ДО - Менеджер - Тип операции - Категория операции - Операция
  строка_ДО_Менеджер = rowByValue(In_Workbooks, In_Sheet, "ДО - Менеджер - Тип операции - Категория операции - Операция", 100, 100)
  
  ' Столбец Итог
  Итог_Column = ColumnByValue(In_Workbooks, In_Sheet, "Итог", 100, 100)
  
  ' Обработка столбца
  ColumnCount = ColumnByValue(In_Workbooks, In_Sheet, "ДО - Менеджер - Тип операции - Категория операции - Операция", 100, 100) + 1
  
  Do While ColumnCount < Итог_Column
    
    ' Если в ячейке не пусто
    ' If Workbooks(In_Workbooks).Sheets(In_Sheet).Cells(строка_ДО_Менеджер, ColumnCount).Value <> "" Then
    If Not IsEmpty(Workbooks(In_Workbooks).Sheets(In_Sheet).Cells(строка_ДО_Менеджер, ColumnCount).Value) Then
      
      dateReportFromCapacity = CDate(Workbooks(In_Workbooks).Sheets(In_Sheet).Cells(строка_ДО_Менеджер, ColumnCount).Value)
      
    End If
    ' Следующий столбец
    ColumnCount = ColumnCount + 1
  Loop
  
  ' t = dateReportFromCapacity
  
End Function


' Отправка письма: отправляю шаблон самому себе для последующей отправки в сеть письма на его основе:
Sub Отправка_Lotus_Notes_Лист6_Capacity()
Dim темаПисьма, текстПисьма, hashTag, attachmentFile As String
Dim i As Byte
  
  ' Подтвержение
  If MsgBox("Отправить себе Шаблон письма с вложением Capacity Model?", vbYesNo) = vbYes Then
    
    ' Тема письма - Тема:
    темаПисьма = subjectFromSheet("Лист6")

    ' hashTag - Хэштэг:
    hashTag = hashTagFromSheet("Лист6")
    
    ' Файл-вложение из "Вложение1"
    attachmentFile = ThisWorkbook.Sheets("Лист6").Range("AM3").Value
 
    ' Текст письма
    текстПисьма = "" + Chr(13)
    текстПисьма = текстПисьма + "" + ThisWorkbook.Sheets("Лист6").Cells(rowByValue(ThisWorkbook.Name, "Лист6", "Список получателей:", 100, 100), ColumnByValue(ThisWorkbook.Name, "Лист6", "Список получателей:", 100, 100) + 2).Value + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "" + getFromAddrBook("РД", 2) + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "Уважаемые руководители," + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "Отработка сотрудниками клиентов на входящем потоке с начала месяца." + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    ' Визитка (подпись С Ув., )
    текстПисьма = текстПисьма + ПодписьВПисьме()
    ' Хэштег
    текстПисьма = текстПисьма + createBlankStr(20) + hashTag
    
    ' Вызов
    Call send_Lotus_Notes(темаПисьма, "Sergey Fedorovich Proschaev/Tyumen/PSBank/Ru", "Sergey Fedorovich Proschaev/Tyumen/PSBank/Ru", текстПисьма, attachmentFile)

    ' Сообщение
    MsgBox ("Письмо отправлено!")
          
  End If
  
End Sub

' Создание книги с PA
Sub createBook_out_PA(In_OutBookName)

    ' Поля: ID_клиента_Retail, PA, DateOfActiv, Выдача_PA, PA_KK, Выдача_РА_КК, chan, ФИО_сотрудника, РегОфис, ДопОфис

    Workbooks.Add
    ActiveWorkbook.SaveAs FileName:=In_OutBookName
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Activate
    
    ' Форматирование полей
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 1).Value = "ID_клиента_Retail"
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns("A:A").EntireColumn.ColumnWidth = 20
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 1).HorizontalAlignment = xlCenter
    
    ' ФИО
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 2).Value = "ФИО"
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns("B:B").EntireColumn.ColumnWidth = 25
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 2).HorizontalAlignment = xlCenter
    
    ' PA
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 3).Value = "PA"
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns("C:C").EntireColumn.ColumnWidth = 10
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 3).HorizontalAlignment = xlCenter
 
    ' DateOfActiv
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 4).Value = "DateOfActiv"
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns("D:D").EntireColumn.ColumnWidth = 20
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 4).HorizontalAlignment = xlCenter
    
    ' Выдача_PA
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 5).Value = "Выдача_PA"
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns("E:E").EntireColumn.ColumnWidth = 10
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 5).HorizontalAlignment = xlCenter
    
    ' PA_KK
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 6).Value = "PA_KK"
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns("F:F").EntireColumn.ColumnWidth = 10
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 6).HorizontalAlignment = xlCenter
    
    ' Выдача_РА_КК
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 7).Value = "Выдача_РА_КК"
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns("G:G").EntireColumn.ColumnWidth = 15
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 7).HorizontalAlignment = xlCenter
    
    ' chan
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 8).Value = "chan"
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns("H:H").EntireColumn.ColumnWidth = 15
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 8).HorizontalAlignment = xlCenter
    
    ' ФИО_сотрудника
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 9).Value = "ФИО_сотрудника"
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns("I:I").EntireColumn.ColumnWidth = 25
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 9).HorizontalAlignment = xlCenter
    
    ' РегОфис
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 10).Value = "РегОфис"
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns("J:J").EntireColumn.ColumnWidth = 15
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 10).HorizontalAlignment = xlCenter
    
    ' ДопОфис
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 11).Value = "ДопОфис"
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns("K:K").EntireColumn.ColumnWidth = 25
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 11).HorizontalAlignment = xlCenter

End Sub

' Создание файла для отправки в офисы
Sub copyDBToSend_Sheet6()
Dim TemplatesFile As String

  Application.StatusBar = "Копирование..."

  ' Открываем "Отчет кросс-продажи.xlsx"
  If Dir(ThisWorkbook.Path + "\Templates\" + "Отчет кросс-продажи.xlsx") <> "" Then
    ' Открываем шаблон Templates\Ежедневный отчет по продажам
    TemplatesFileName = "Отчет кросс-продажи"
  End If
              
  ' Открываем шаблон Протокола из C:\Users\...\Documents\#VBA\DB_Result\Templates
  Workbooks.Open (ThisWorkbook.Path + "\Templates\" + TemplatesFileName + ".xlsx")
           
  ' Переходим на окно DB
  ThisWorkbook.Sheets("Лист6").Activate

  ' Обновляем список получателей
  ' ThisWorkbook.Sheets("Лист8").Cells(rowByValue(ThisWorkbook.Name, "Лист8", "Список получателей:", 100, 100), ColumnByValue(ThisWorkbook.Name, "Лист8", "Список получателей:", 100, 100) + 2).Value = _
  '   getFromAddrBook("УДО2,УДО3,УДО4,УДО5,НОРПиКО1,НОРПиКО2,НОРПиКО3,НОРПиКО4,НОРПиКО5,ПМ,МРК1,МРК2,МРК3,МРК4,МРК5,НОКП,РРКК,МПП", 2)

  ' Имя нового файла
  FileCapacityName = Replace(Mid(ThisWorkbook.Sheets("Лист6").Range("B2").Value, 1, 41), ".", "-") + ".xlsx"
  Workbooks(TemplatesFileName).SaveAs FileName:=ThisWorkbook.Path + "\Out\" + FileCapacityName, FileFormat:=xlOpenXMLWorkbook, createBackUp:=False
  ' Вложение1
  ThisWorkbook.Sheets("Лист6").Range("AM3").Value = ThisWorkbook.Path + "\Out\" + FileCapacityName
            
  ' *** Копирование данных ***
   
  ' Отчет кросс-продажи с по г.
  ThisWorkbook.Sheets("Лист6").Range("B1").Copy Destination:=Workbooks(FileCapacityName).Sheets("Capacity Model").Range("B1")
   
  ' Копируем цели
  For j = 1 To 34
    ThisWorkbook.Sheets("Лист6").Cells(3, j).Copy Destination:=Workbooks(FileCapacityName).Sheets("Capacity Model").Cells(3, j)
  Next j
      
  ' Копируем данные по офисам
  For i = 6 To 40
      
    For j = 1 To 34
      ThisWorkbook.Sheets("Лист6").Cells(i, j).Copy Destination:=Workbooks(FileCapacityName).Sheets("Capacity Model").Cells(i, j)
    Next j
      
  Next i
  
  ' ***
                    
  ' Закрытие файла
  Workbooks(FileCapacityName).Close SaveChanges:=True

  ' Копирование завершено
  Application.StatusBar = "Скопировано!"

End Sub

' Отправка_Lotus_Notes_Лист6_Pre_Approved Отправка письма: отправляю шаблон самому себе для последующей отправки в сеть письма на его основе:
Sub Отправка_Lotus_Notes_Лист6_Pre_Approved()
Dim темаПисьма, текстПисьма, hashTag, attachmentFile As String
Dim i As Byte
  
  ' Подтвержение
  If MsgBox("Отправить себе Шаблон письма с вложением Pre-Approved?", vbYesNo) = vbYes Then
    
    ' Тема письма - Тема:
    темаПисьма = "Клиенты с Pre-Approved на " + Mid(ThisWorkbook.Sheets("Лист6").Range("B2").Value, 32, 10)

    ' hashTag - Хэштэг:
    hashTag = hashTagFromSheet("Лист6") + " #Pre-Approved"
    
    ' Файл-вложение из "Вложение2"
    attachmentFile = ThisWorkbook.Sheets("Лист6").Range("AO3").Value
 
    ' Текст письма
    текстПисьма = "" + Chr(13)
    текстПисьма = текстПисьма + "" + ThisWorkbook.Sheets("Лист6").Cells(rowByValue(ThisWorkbook.Name, "Лист6", "Список получателей:", 100, 100), ColumnByValue(ThisWorkbook.Name, "Лист6", "Список получателей:", 100, 100) + 2).Value + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "" + getFromAddrBook("РД", 2) + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "Уважаемые руководители," + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "Направляю список клиентов с упущенными готовыми решениями по потребкредитам и КК." + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "Прошу организовать отработку в срок до " + CStr(weekEndDate(Date) - 2) + " с конверсией не менее " + CStr(ThisWorkbook.Sheets("Лист6").Range("Q3").Value * 100) + "% " + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    ' Визитка (подпись С Ув., )
    текстПисьма = текстПисьма + ПодписьВПисьме()
    ' Хэштег
    текстПисьма = текстПисьма + createBlankStr(20) + hashTag
    
    ' Вызов
    Call send_Lotus_Notes(темаПисьма, "Sergey Fedorovich Proschaev/Tyumen/PSBank/Ru", "Sergey Fedorovich Proschaev/Tyumen/PSBank/Ru", текстПисьма, attachmentFile)

    ' Сообщение
    MsgBox ("Письмо отправлено!")
          
  End If
  
End Sub

' Получение ячейки в которую записываем комменатрий на Лист8 для In_officeName по Интернет-банк (ИБ)
Function Range_Лист8_Интернет_банк(In_officeName)
        
        Select Case In_officeName
          Case "ОО «Тюменский»"
            Range_Лист8_Интернет_банк = "M15"
          Case "ОО «Сургутский»"
            Range_Лист8_Интернет_банк = "M53"
          Case "ОО «Нижневартовский»"
            Range_Лист8_Интернет_банк = "M91"
          Case "ОО «Новоуренгойский»"
            Range_Лист8_Интернет_банк = "M129"
          Case "ОО «Тарко-Сале»"
            Range_Лист8_Интернет_банк = "M167"
        End Select
  
End Function

' Создание книги с ИБ
Sub createBook_out_ИБ(In_OutBookName)

    Workbooks.Add
    ActiveWorkbook.SaveAs FileName:=In_OutBookName
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Activate
    
    ' Форматирование полей
    
    ' ID_клиента_Retail
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 1).Value = "ID_клиента_Retail"
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns("A:A").EntireColumn.ColumnWidth = 20
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 1).HorizontalAlignment = xlCenter
    
    ' ФИО
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 2).Value = "ФИО"
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns("B:B").EntireColumn.ColumnWidth = 25
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 2).HorizontalAlignment = xlCenter
    
    ' Тип_клиента
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 3).Value = "Тип_клиента"
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns("C:C").EntireColumn.ColumnWidth = 18
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 3).HorizontalAlignment = xlCenter
    
    ' IB
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 4).Value = "ИБ"
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns("D:D").EntireColumn.ColumnWidth = 8
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 4).HorizontalAlignment = xlCenter
 
    ' Net IB
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 5).Value = "Нет ИБ"
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns("E:E").EntireColumn.ColumnWidth = 12
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 5).HorizontalAlignment = xlCenter
    
    ' ИБ активный
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 6).Value = "ИБ активный"
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns("F:F").EntireColumn.ColumnWidth = 15
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 6).HorizontalAlignment = xlCenter
    
    ' Потенциал реактивации ИБ
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 7).Value = "Потенциал реактивации ИБ"
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns("G:G").EntireColumn.ColumnWidth = 30
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 7).HorizontalAlignment = xlCenter
    
    ' ФИО_сотрудника
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 8).Value = "ФИО_сотрудника"
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns("H:H").EntireColumn.ColumnWidth = 20
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 8).HorizontalAlignment = xlCenter
    
    ' ДопОфис
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 9).Value = "ДопОфис"
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns("I:I").EntireColumn.ColumnWidth = 15
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 9).HorizontalAlignment = xlCenter
    
    ' РегОфис
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 10).Value = "РегОфис"
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns("J:J").EntireColumn.ColumnWidth = 15
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 10).HorizontalAlignment = xlCenter

    ' Комментарий
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 11).Value = "Комментарий"
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns("K:K").EntireColumn.ColumnWidth = 50
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 11).HorizontalAlignment = xlCenter


End Sub

' Отправка_Lotus_Notes_Лист6_IB
Sub Отправка_Lotus_Notes_Лист6_IB()
Dim темаПисьма, текстПисьма, hashTag, attachmentFile As String
Dim i As Byte
  
  ' Подтвержение
  If MsgBox("Отправить себе Шаблон письма с вложением IB?", vbYesNo) = vbYes Then
    
    ' Тема письма - Тема:
    темаПисьма = "Клиенты с потенциалом подключения Интернет-банка на " + Mid(ThisWorkbook.Sheets("Лист6").Range("B2").Value, 32, 10)

    ' hashTag - Хэштэг:
    hashTag = hashTagFromSheet("Лист6") + " #IB"
    
    ' Файл-вложение из "Вложение2"
    attachmentFile = ThisWorkbook.Sheets("Лист6").Range("AQ3").Value
 
    ' Текст письма
    текстПисьма = "" + Chr(13)
    текстПисьма = текстПисьма + "" + ThisWorkbook.Sheets("Лист6").Cells(rowByValue(ThisWorkbook.Name, "Лист6", "Список получателей:", 100, 100), ColumnByValue(ThisWorkbook.Name, "Лист6", "Список получателей:", 100, 100) + 2).Value + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "" + getFromAddrBook("РД", 2) + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "Уважаемые руководители," + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "Направляю список клиентов с потенциалом подключения и реактивации ИБ." + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "Прошу отработать клиентов по фильтрам:" + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "- столбец " + Chr(34) + "ИБ" + Chr(34) + " = 0 - клиенты, которым можно подключить ИБ" + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "- столбец " + Chr(34) + "Потенциал реактивации ИБ" + Chr(34) + " = 1 - клиенты у которых возможна реактивация ИБ" + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    ' Визитка (подпись С Ув., )
    текстПисьма = текстПисьма + ПодписьВПисьме()
    ' Хэштег
    текстПисьма = текстПисьма + createBlankStr(20) + hashTag
    
    ' Вызов
    Call send_Lotus_Notes(темаПисьма, "Sergey Fedorovich Proschaev/Tyumen/PSBank/Ru", "Sergey Fedorovich Proschaev/Tyumen/PSBank/Ru", текстПисьма, attachmentFile)

    ' Сообщение
    MsgBox ("Письмо отправлено!")
          
  End If
  
End Sub

