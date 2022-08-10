Attribute VB_Name = "Module_Лист5"
' Лист 5 Оперативная бизнес-справка (карточные продукты) на ДД.ММ.ГГГГ

' Проодажи_Дебетовых_и_Кредитных_карт Макрос
Sub Проодажи_Дебетовых_и_Кредитных_карт()
Attribute Проодажи_Дебетовых_и_Кредитных_карт.VB_ProcData.VB_Invoke_Func = " \n14"

Dim maxDateInReport, dateBeginMonth, dateBeginWeek, dateEndWeek As Date
Dim rowCount, countDayDebetCard, countWeekDebetCard, countMonthDebetCard, countDayCreditCard, countWeekCreditCard, countMonthCreditCard, workingDaysMonth, Пенсионные_карты_месяц As Integer
Dim planDayDebetCard, planDayCreditCard, lagDebetCard, lagCreditCard As Double
Dim ЗаноситьКартуВБазу As Boolean
Dim officeNameInReport, Программа_выпуска_карты As String
' Имя ячейки (например G41) в которой находтся заданный К_пор (например ЗДК1)
Dim RangeК_пор As String
Dim RangeК_пор_Row, RangeК_пор_Column As Byte
' Индиктор вызова процедур setК_порInЕСУП, currentК_порInЕСУП
Dim call_setК_порInЕСУП, call_currentК_порInЕСУП As Boolean
Dim finishProcess As Boolean
Dim CheckFormatReportResult As String
Dim DateTimeStart, DateTimeEnd As Date

  ' Запрос - сегодня понедельник (системная дата) - формировать отчет за дату сегодня?
  ' Убираем
  ' If Weekday(Date, vbMonday) = 1 Then
  '   MsgBox ("Сегодня понедельник, сформировать отчет с прогнозом на неделю!")
  '  ' Устанавливаем O8=1 P8=Date:
  '   ThisWorkbook.Sheets("Лист5").Cells(8, 15).Value = 1
  '   ThisWorkbook.Sheets("Лист5").Cells(8, 16).Value = Date
  ' Else
    ' Устанавливаем O8=0
  '   ThisWorkbook.Sheets("Лист5").Cells(8, 15).Value = 0
  ' End If

  ' Объем кредитного портфеля с изменениями за период
  ' Открыть файл с отчетом
  FileName = Application.GetOpenFilename("Excel Files (*.xml), *.xml", , "Открытие файла с отчетом")

  ' Если файл был выбран
  If (Len(FileName) > 5) Then
  
    ' Строка статуса
    Application.StatusBar = "Обработка отчета..."
  
    DateTimeStart = Now
    ThisWorkbook.Sheets("Лист5").Cells(15, 19).Value = CStr(DateTimeStart)
  
    ' Выводим для инфо данные об имени файла
    DBstrName_String = Dir(FileName)
  
    ' Открываем выбранную книгу (UpdateLinks:=0)
    Workbooks.Open FileName, 0

    ' Переходим на окно DB
    ThisWorkbook.Sheets("Лист5").Activate

    ' Проверка формы отчета
    CheckFormatReportResult = CheckFormatReport(DBstrName_String, "Список", 5, periodFromSheet("Лист5"))
    
    If CheckFormatReportResult = "OK" Then

    ' Переменная процесса обработки
    finishProcess = True

    ' Открываем BASE\Cards
    OpenBookInBase ("Cards")

    ' Открываем BASE\Tasks
    OpenBookInBase ("Tasks")
    
    ' Открываем BASE\Clients
    OpenBookInBase ("Clients")

    ' Обрабатываем отчет
    maxDateInReport = CDate(Mid(Workbooks(DBstrName_String).Sheets("Список").Cells(2, 1).Value, 28, 10)) - 1
    ' Пробуем брать дату отчета
    ' maxDateInReport = CDate(Mid(Workbooks(DBstrName_String).Sheets("Список").Cells(2, 1).Value, 28, 10))
    
    ' Определяем maxDateInReport: если сегодня понедельник, то берем сегодняшний день, если любой другой день, то берем из отчета эмиссии
    ' If ThisWorkbook.Sheets("Лист5").Cells(8, 15).Value = 1 Then
      ' Берем дату из ячейки P8
    ' maxDateInReport = CDate(ThisWorkbook.Sheets("Лист5").Cells(8, 16).Value)
    ' Else
      ' Берем максимальную дату из отчета А2=" За период с 01.01.2020 по 12.02.2020" и  вычитаем один день!
      ' maxDateInReport = CDate(Mid(Workbooks(DBstrName_String).Sheets("Список").Cells(2, 1).Value, 28, 10)) - 1
      ' На период разработки - берем дату 10.02.2020
      ' maxDateInReport = CDate("17.02.2020")
    ' End If
        
    ' maxDateInReport = CDate("21.02.2020")
    ' MsgBox ("Внимание! режим отладки, используется всегда дата " + CStr(maxDateInReport))
        
    ' Дата начала месяца
    dateBeginMonth = CDate("01." + Mid(CStr(maxDateInReport), 4, 7))
    
    ' Дата начала недели
    dateBeginWeek = weekStartDate(maxDateInReport)
    
    ' Дата конца недели
    dateEndWeek = weekEndDate(maxDateInReport)
    
    ' Число рабочих дней в месяце - MsgBox ("Число рабочих дней в месяце " + CStr(workingDaysMonth))
    workingDaysMonth = Working_days_in_the_FullMonth(maxDateInReport, 6)
    
    ' Выводим данные в заголовки
    ' Неделя
    ThisWorkbook.Sheets("Лист5").Cells(2, 10).Value = CStr(WeekNumber(maxDateInReport))
    
    ' ThisWorkbook.Sheets("Лист5").Cells(5, 7).Value = "Дебетовые карты, неделя с " + strDDMM(dateBeginWeek) + " по " + strDDMM(dateEndWeek)
    ' Дебетовые карты,           неделя (48) с 10.02 по 16.02
    ThisWorkbook.Sheets("Лист5").Cells(5, 7).Value = "Дебетовые карты,           неделя (" + CStr(WeekNumber(maxDateInReport)) + ") с " + strDDMM(dateBeginWeek) + " по " + strDDMM(dateEndWeek)
    
    ' ThisWorkbook.Sheets("Лист5").Cells(5, 9).Value = "Кредитные карты, неделя с " + strDDMM(dateBeginWeek) + " по " + strDDMM(dateEndWeek)
    ' Кредитные карты,           неделя (48) с 10.02 по 16.02
    ThisWorkbook.Sheets("Лист5").Cells(5, 9).Value = "Кредитные карты,           неделя (" + CStr(WeekNumber(maxDateInReport)) + ") с " + strDDMM(dateBeginWeek) + " по " + strDDMM(dateEndWeek)
    
    '
    ThisWorkbook.Sheets("Лист5").Cells(2, 2).Value = "Оперативная бизнес-справка (карточные продукты) на " + CStr(maxDateInReport) + " г."
    ThisWorkbook.Sheets("Лист5").Cells(5, 11).Value = "Заявки за " + Mid(CStr(maxDateInReport), 1, 5)
    ThisWorkbook.Sheets("Лист5").Cells(2, 18).Value = "Оперативная бизнес-справка (карточные продукты) на " + CStr(maxDateInReport) + " г."

    ' Индиктор вызова процедур setК_порInЕСУП, currentК_порInЕСУП
    call_setК_порInЕСУП = False
    call_currentК_порInЕСУП = False

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

      ' Обработка отчета
      rowCount = 5
      countDayDebetCard = 0
      countDayCreditCard = 0
      countWeekDebetCard = 0
      countWeekCreditCard = 0
      countMonthDebetCard = 0
      countMonthCreditCard = 0
      Пенсионные_карты_месяц = 0
      
      Do While Not IsEmpty(Workbooks(DBstrName_String).Sheets("Список").Cells(rowCount, 1).Value)
        
        ' Дата оформления заказа (11)
        ' If CDate(Workbooks(DBstrName_String).Sheets("Список").Cells(RowCount, 11).Value) = maxDateInReport Then
          
          ' Переменная заносить карту в BASE\Cards
          ЗаноситьКартуВБазу = False
          ' Строка статуса
          Application.StatusBar = "Обработка отчета: " + officeNameInReport + " " + CStr(rowCount)

          ' Если это текущий офис - Подразделение обслуживания (10)
          If (InStr(Workbooks(DBstrName_String).Sheets("Список").Cells(rowCount, 10).Value, officeNameInReport) <> 0) Then
            
            ' Режим заказа (19): Заказ новой карты, ???, Выдача карты моментального выпуска
            If (Workbooks(DBstrName_String).Sheets("Список").Cells(rowCount, 19).Value = "Заказ новой карты") Or (Workbooks(DBstrName_String).Sheets("Список").Cells(rowCount, 19).Value = "???") Or (Workbooks(DBstrName_String).Sheets("Список").Cells(rowCount, 19).Value = "Выдача карты моментального выпуска") Then
            
              ' Программа выпуска
              Программа_выпуска_карты = Workbooks(DBstrName_String).Sheets("Список").Cells(rowCount, 4).Value
            
              ' Дебетовая карта - Программа выпуска (4): В движении, Хорошее настроение, Пенсионная карта, Дебетовая карта "Карта мира без границ", Твой ПСБ, Твой кэшбэк
              If (Программа_выпуска_карты = "В движении") Or (Программа_выпуска_карты = "Хорошее настроение") Or (Программа_выпуска_карты = "Пенсионная карта") Or (InStr(Программа_выпуска_карты, "Карта мира без границ") <> 0) Or (Программа_выпуска_карты = "Твой ПСБ") Or (Программа_выпуска_карты = "Твой кэшбэк") Then
              
                ' Дата оформления заказа (11)
                ' День
                If CDate(Workbooks(DBstrName_String).Sheets("Список").Cells(rowCount, 11).Value) = maxDateInReport Then
                  countDayDebetCard = countDayDebetCard + 1
                End If
                ' Неделя
                If (CDate(Workbooks(DBstrName_String).Sheets("Список").Cells(rowCount, 11).Value) >= dateBeginWeek) And (CDate(Workbooks(DBstrName_String).Sheets("Список").Cells(rowCount, 11).Value) <= maxDateInReport) Then
                  countWeekDebetCard = countWeekDebetCard + 1
                End If
                ' Месяц
                If (CDate(Workbooks(DBstrName_String).Sheets("Список").Cells(rowCount, 11).Value) >= dateBeginMonth) And (CDate(Workbooks(DBstrName_String).Sheets("Список").Cells(rowCount, 11).Value) <= maxDateInReport) Then
                  countMonthDebetCard = countMonthDebetCard + 1
                  ЗаноситьКартуВБазу = True
                  ' Пенсионные карты с начала месяца
                  If Программа_выпуска_карты = "Пенсионная карта" Then
                    Пенсионные_карты_месяц = Пенсионные_карты_месяц + 1
                  End If
                End If ' Месяц
                
                ' Заносим пенсионера в BASE\Clients
                If (Программа_выпуска_карты = "Пенсионная карта") Then
                
                  НК_Клиента = ПреобразованиеФИОиНК2(Workbooks(DBstrName_String).Sheets("Список").Cells(rowCount, 16).Value)
                
                  Call InsertRecordInBook("Clients", "Лист1", "Номер_клиента", НК_Клиента, _
                                            "Номер_клиента", НК_Клиента, _
                                              "Офис", cityOfficeName(officeNameInReport), _
                                                "ФИО", Workbooks(DBstrName_String).Sheets("Список").Cells(rowCount, 14).Value, _
                                                  "Фамилия", ПреобразованиеФИОиНК3(Workbooks(DBstrName_String).Sheets("Список").Cells(rowCount, 14).Value), _
                                                    "Имя", ПреобразованиеФИОиНК4(Workbooks(DBstrName_String).Sheets("Список").Cells(rowCount, 14).Value), _
                                                      "Отчество", ПреобразованиеФИОиНК5(Workbooks(DBstrName_String).Sheets("Список").Cells(rowCount, 14).Value), _
                                                        "Контакты", Workbooks(DBstrName_String).Sheets("Список").Cells(rowCount, 15).Value, _
                                                          "Cards_Pensioner_Date", Workbooks(DBstrName_String).Sheets("Список").Cells(rowCount, 13).Value, _
                                                            "Cards_Pensioner", Программа_выпуска_карты, _
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
                End If ' Заносим пенсионера в BASE\Clients
                
              End If ' Дебетовая карта
            
              ' Кредитная карта - Программа выпуска (4): Двойной кэшбэк, Кредитная карта 100+, "Кредитная карта "Платинум", Суперкарта
              ' If (Workbooks(DBstrName_String).Sheets("Список").Cells(RowCount, 4).Value = "Двойной кэшбэк") Or (Workbooks(DBstrName_String).Sheets("Список").Cells(RowCount, 4).Value = "Кредитная карта 100+") Or (InStr(Workbooks(DBstrName_String).Sheets("Список").Cells(RowCount, 4).Value, "Платинум") <> 0) Or (Workbooks(DBstrName_String).Sheets("Список").Cells(RowCount, 4).Value = "Суперкарта") Then
              If (Программа_выпуска_карты = "Двойной кэшбэк") Or (Программа_выпуска_карты = "Кредитная карта 100+") Or (InStr(Программа_выпуска_карты, "Платинум") <> 0) Or (Программа_выпуска_карты = "Суперкарта") Then
                ' День
                If CDate(Workbooks(DBstrName_String).Sheets("Список").Cells(rowCount, 11).Value) = maxDateInReport Then
                  countDayCreditCard = countDayCreditCard + 1
                End If
                ' Неделя
                If (CDate(Workbooks(DBstrName_String).Sheets("Список").Cells(rowCount, 11).Value) >= dateBeginWeek) And (CDate(Workbooks(DBstrName_String).Sheets("Список").Cells(rowCount, 11).Value) <= maxDateInReport) Then
                  countWeekCreditCard = countWeekCreditCard + 1
                End If
                ' Месяц
                If (CDate(Workbooks(DBstrName_String).Sheets("Список").Cells(rowCount, 11).Value) >= dateBeginMonth) And (CDate(Workbooks(DBstrName_String).Sheets("Список").Cells(rowCount, 11).Value) <= maxDateInReport) Then
                  countMonthCreditCard = countMonthCreditCard + 1
                  ЗаноситьКартуВБазу = True
                End If
                
              End If ' Кредитная карта
              
              ' Заносить карту в Базу
              If ЗаноситьКартуВБазу = True Then
                ' Заносим в BASE\Cards
                Call InsertRecordInBook("Cards", "Лист1", "Номер_заказа", Workbooks(DBstrName_String).Sheets("Список").Cells(rowCount, 1).Value, _
                                            "Дата_оформления_заказа", Workbooks(DBstrName_String).Sheets("Список").Cells(rowCount, 11).Value, _
                                              "Подразделение_обслуживания", officeNameInReport, _
                                                "Номер_заказа", Workbooks(DBstrName_String).Sheets("Список").Cells(rowCount, 1).Value, _
                                                  "Номер_карты", Workbooks(DBstrName_String).Sheets("Список").Cells(rowCount, 2).Value, _
                                                    "Номер_договора", Workbooks(DBstrName_String).Sheets("Список").Cells(rowCount, 3).Value, _
                                                      "Программа_выпуска", Программа_выпуска_карты, _
                                                        "Карточный_продукт", Workbooks(DBstrName_String).Sheets("Список").Cells(rowCount, 5).Value, _
                                                          "Группа", Workbooks(DBstrName_String).Sheets("Список").Cells(rowCount, 6).Value, _
                                                            "Лимит_кредитования", Workbooks(DBstrName_String).Sheets("Список").Cells(rowCount, 7).Value, _
                                                              "Состояние", Workbooks(DBstrName_String).Sheets("Список").Cells(rowCount, 8).Value, _
                                                                "Место_получения_карты", Workbooks(DBstrName_String).Sheets("Список").Cells(rowCount, 9).Value, _
                                                                  "Дата_последнего_изменения", Workbooks(DBstrName_String).Sheets("Список").Cells(rowCount, 12).Value, _
                                                                    "Дата_выдачи", Workbooks(DBstrName_String).Sheets("Список").Cells(rowCount, 13).Value, _
                                                                      "Владелец_счета", ПреобразованиеФИОиНК(Workbooks(DBstrName_String).Sheets("Список").Cells(rowCount, 14).Value), _
                                                                        "Держатель_карты", ПреобразованиеФИОиНК(Workbooks(DBstrName_String).Sheets("Список").Cells(rowCount, 16).Value), _
                                                                          "Канал_заказа", Workbooks(DBstrName_String).Sheets("Список").Cells(rowCount, 18).Value, _
                                                                            "Режим_заказа", Workbooks(DBstrName_String).Sheets("Список").Cells(rowCount, 19).Value, _
                                                                              "", "", _
                                                                                "", "", _
                                                                                  "", "")
                                                                     
              End If ' Заносить карту в Базу
              
            End If ' Режим заказа
          
          End If
        
        ' End If
      
        ' Следующая запись
        DoEventsInterval (rowCount)
        rowCount = rowCount + 1
      Loop
   
      ' Выводим данные по офису:
      
      ' Дк
      ThisWorkbook.Sheets("Лист5").Cells(6 + i, 4).Value = countMonthDebetCard
      ThisWorkbook.Sheets("Лист5").Cells(6 + i, 8).Value = countWeekDebetCard
      ThisWorkbook.Sheets("Лист5").Cells(6 + i, 11).Value = countDayDebetCard
      ' КК
      ThisWorkbook.Sheets("Лист5").Cells(6 + i, 6).Value = countMonthCreditCard
      ThisWorkbook.Sheets("Лист5").Cells(6 + i, 10).Value = countWeekCreditCard
      ThisWorkbook.Sheets("Лист5").Cells(6 + i, 12).Value = countDayCreditCard
      
      ' План на 1 день - убрал 12-03
      ' planDayDebetCard = ThisWorkbook.Sheets("Лист5").Cells(6 + i, 3).Value / workingDaysMonth
      ' planDayCreditCard = ThisWorkbook.Sheets("Лист5").Cells(6 + i, 5).Value / workingDaysMonth
      
      ' Отставание на начало недели - убрал 12-03
      ' lagDebetCard = planDayDebetCard * Working_days_between_dates(dateBeginMonth, dateBeginWeek - 1, 6) - (ThisWorkbook.Sheets("Лист5").Cells(6 + i, 4).Value - ThisWorkbook.Sheets("Лист5").Cells(6 + i, 8).Value)
      
      ' MsgBox ("Отставание на начало недели ДК: " + CStr(lagDebetCard)) - убрал 12-03
      ' lagCreditCard = planDayCreditCard * Working_days_between_dates(dateBeginMonth, dateBeginWeek - 1, 6) - (ThisWorkbook.Sheets("Лист5").Cells(6 + i, 6).Value - ThisWorkbook.Sheets("Лист5").Cells(6 + i, 10).Value)
      
      ' План на неделю по ДК
      ' ThisWorkbook.Sheets("Лист5").Cells(6 + i, 7).Value = Round((planDayDebetCard * Working_days_between_dates(dateBeginWeek, dateEndWeek, 6)) + ((lagDebetCard / Working_days_between_dates(dateBeginWeek, Date_last_day_month(maxDateInReport), 6)) * Working_days_between_dates(dateBeginWeek, dateEndWeek, 6)), 0)
      ' План на неделю по КК
      ' ThisWorkbook.Sheets("Лист5").Cells(6 + i, 9).Value = Round((planDayCreditCard * Working_days_between_dates(dateBeginWeek, dateEndWeek, 6)) + ((lagCreditCard / Working_days_between_dates(dateBeginWeek, Date_last_day_month(maxDateInReport), 6)) * Working_days_between_dates(dateBeginWeek, dateEndWeek, 6)), 0)
            
      ' План на неделю по ДК (вариант 2) - 6 дней
      ' ThisWorkbook.Sheets("Лист5").Cells(6 + i, 7).Value = Round(ПланНаНеделю(ThisWorkbook.Sheets("Лист5").Cells(6 + i, 3).Value, ThisWorkbook.Sheets("Лист5").Cells(6 + i, 4).Value - ThisWorkbook.Sheets("Лист5").Cells(6 + i, 8).Value, maxDateInReport, 6), 0)
      ' План на неделю по ДК (вариант 2) - 5 дней
      ThisWorkbook.Sheets("Лист5").Cells(6 + i, 7).Value = Round(ПланНаНеделю(ThisWorkbook.Sheets("Лист5").Cells(6 + i, 3).Value, ThisWorkbook.Sheets("Лист5").Cells(6 + i, 4).Value - ThisWorkbook.Sheets("Лист5").Cells(6 + i, 8).Value, maxDateInReport, 5), 0)
      
      ' План на неделю по КК (вариант 2) - 6 дней
      ' ThisWorkbook.Sheets("Лист5").Cells(6 + i, 9).Value = Round(ПланНаНеделю(ThisWorkbook.Sheets("Лист5").Cells(6 + i, 5).Value, ThisWorkbook.Sheets("Лист5").Cells(6 + i, 6).Value - ThisWorkbook.Sheets("Лист5").Cells(6 + i, 10).Value, maxDateInReport, 6), 0)
      ' План на неделю по КК (вариант 2) - 5 дней
      ThisWorkbook.Sheets("Лист5").Cells(6 + i, 9).Value = Round(ПланНаНеделю(ThisWorkbook.Sheets("Лист5").Cells(6 + i, 5).Value, ThisWorkbook.Sheets("Лист5").Cells(6 + i, 6).Value - ThisWorkbook.Sheets("Лист5").Cells(6 + i, 10).Value, maxDateInReport, 5), 0)
            
      ' Вывод по категориям заведенных карт
      ' Пенсионные_карты_месяц
      ThisWorkbook.Sheets("Лист5").Cells(6 + 16 + i, 3).Value = Пенсионные_карты_месяц
            
      ' Заносим Поручения в Лист "ЕСУП"
      ' План на неделю по заявкам на неделю ДК
      If НеделяНаЛистеN("ЕСУП") = WeekNumber(maxDateInReport) Then
        ThisWorkbook.Sheets("ЕСУП").Cells(6 + i, 7).Value = ThisWorkbook.Sheets("Лист5").Cells(6 + i, 7).Value
        ' Поручение офису по ЗДКi
        Call setК_порInЕСУП(ThisWorkbook.Name, "ЕСУП", "ЗДК" + CStr(i), ThisWorkbook.Sheets("Лист5").Cells(6 + i, 7).Value, dateBeginWeek, "шт.", "")
        ' Переменная, что процедуру setК_порInЕСУП вызывали
        call_setК_порInЕСУП = True
      End If
      
      ' Факт по заявкам ДК недели
      If НеделяНаЛистеN("ЕСУП") = WeekNumber(maxDateInReport) Then
        ThisWorkbook.Sheets("ЕСУП").Cells(6, 8).Value = "Факт нед.(" + CStr(WeekNumber(maxDateInReport)) + ")"
        ThisWorkbook.Sheets("ЕСУП").Cells(6 + i, 8).Value = ThisWorkbook.Sheets("Лист5").Cells(6 + i, 8).Value
        ' --- Заносим факт исполнения ЕСУП
        Call currentК_порInЕСУП(ThisWorkbook.Name, "ЕСУП", "ЗДК" + CStr(i), maxDateInReport, ThisWorkbook.Sheets("Лист5").Cells(6 + i, 8).Value, "шт.")
        ' Переменная, что процедуру currentК_порInЕСУП вызывали
        call_currentК_порInЕСУП = True
        ' --- Конец Заносим факт исполнения ЕСУП
      End If
      
      ' Факт по заявкам КК недели
      If НеделяНаЛистеN("ЕСУП") = WeekNumber(maxDateInReport) Then
        ThisWorkbook.Sheets("ЕСУП").Cells(6, 10).Value = "Факт нед.(" + CStr(WeekNumber(maxDateInReport)) + ")"
        ThisWorkbook.Sheets("ЕСУП").Cells(6 + i, 10).Value = ThisWorkbook.Sheets("Лист5").Cells(6 + i, 10).Value
      End If
      
      ' План на неделю по заявкам на неделю КК
      If НеделяНаЛистеN("ЕСУП") = WeekNumber(maxDateInReport) Then
        ThisWorkbook.Sheets("ЕСУП").Cells(6 + i, 9).Value = ThisWorkbook.Sheets("Лист5").Cells(6 + i, 9).Value
        ' Поручение офису ЗКК (строку находим по ЗККi) ЗККi
        Call setК_порInЕСУП(ThisWorkbook.Name, "ЕСУП", "ЗКК" + CStr(i), ThisWorkbook.Sheets("Лист5").Cells(6 + i, 9).Value, dateBeginWeek, "шт.", "")
        ' Переменная, что процедуру setК_порInЕСУП вызывали
        call_setК_порInЕСУП = True
      End If
      
      ' --- Заносим факт исполнения ЕСУП
      If НеделяНаЛистеN("ЕСУП") = WeekNumber(maxDateInReport) Then
        Call currentК_порInЕСУП(ThisWorkbook.Name, "ЕСУП", "ЗКК" + CStr(i), maxDateInReport, ThisWorkbook.Sheets("Лист5").Cells(6 + i, 10).Value, "шт.")
        ' Переменная, что процедуру currentК_порInЕСУП вызывали
        call_currentК_порInЕСУП = True
      End If
      ' --- Конец Заносим факт исполнения ЕСУП
      
    Next i ' Следующий офис

    ' Формируем список для отправки (в "Список получателей:"):
    ' ThisWorkbook.Sheets("Лист5").Cells(rowByValue(ThisWorkbook.Name, "Лист5", "Список получателей:", 100, 100), ColumnByValue(ThisWorkbook.Name, "Лист5", "Список получателей:", 100, 100) + 2).Value = getFromAddrBook("УДО2,УДО3,УДО4,УДО5,НОРПиКО1,НОРПиКО2,НОРПиКО3,НОРПиКО4,НОРПиКО5,ПМ,МРК1,МРК2,МРК3,МРК4,МРК5,НОКП,РРКК,МПП", 2)
    ThisWorkbook.Sheets("Лист5").Cells(rowByValue(ThisWorkbook.Name, "Лист5", "Список получателей:", 100, 100), ColumnByValue(ThisWorkbook.Name, "Лист5", "Список получателей:", 100, 100) + 2).Value = getFromAddrBook("УДО2,УДО3,УДО4,УДО5,НОРПиКО1,НОРПиКО2,НОРПиКО3,НОРПиКО4,НОРПиКО5,ПМ,НОКП,РРКК,МПП", 2)

    ' Закрываем BASE\Cards
    CloseBook ("Cards")

    ' Закрываем базу BASE\Tasks
    CloseBook ("Tasks")

    ' Закрываем базу BASE\Clients
    CloseBook ("Clients")

    ' Строка статуса
    Application.StatusBar = "Завершение ..."

    ' Переходим в ячейку M2
    ThisWorkbook.Sheets("Лист5").Cells(4, 15).Select

    ' Строка статуса
    Application.StatusBar = ""

    ' Зачеркиваем пункт меню на стартовой страницы
    
    ' Проверить зачеркивание в зависимости от активности
    If (ThisWorkbook.Sheets("Лист0").Cells(rowByValue(ThisWorkbook.Name, "Лист0", "Регламентная карта:", 100, 100), ColumnByValue(ThisWorkbook.Name, "Лист0", "Регламентная карта:", 100, 100) + 4).Value = "(понедельник)") Or (CStr(ThisWorkbook.Sheets("Лист0").Cells(rowByValue(ThisWorkbook.Name, "Лист0", "Первый день недели:", 100, 100), ColumnByValue(ThisWorkbook.Name, "Лист0", "Первый день недели:", 100, 100) + 2).Value) = "1") Then
      ' Если был отчет за неделю "Оперативная справка по активам за неделю", т.е.: Лист0-F2 = "(понедельник)" или Лист0-L2 = "1"
      Call ЗачеркиваемТекстВячейке("Лист0", RangeByValue(ThisWorkbook.Name, "Лист0", "Оперативная справка по заявкам на карты за неделю", 100, 100))
    Else
      ' Если был отчет за текущий день недели
      Call ЗачеркиваемТекстВячейке("Лист0", RangeByValue(ThisWorkbook.Name, "Лист0", "Оперативная справка по заявкам на карты", 100, 100))
    End If

    ' Сохранение изменений
    ThisWorkbook.Save

    ' Переменная процесса обработки
    finishProcess = True

    DateTimeEnd = Now
    ThisWorkbook.Sheets("Лист5").Cells(16, 19).Value = CStr(DateTimeEnd)
    ThisWorkbook.Sheets("Лист5").Cells(17, 19).Value = CStr(DateTimeEnd - DateTimeStart)
    
    Else
      
      ' Сообщение о неверном формате отчета или даты
      MsgBox ("Проверьте отчет: " + CheckFormatReportResult + "!")
    
    End If ' Проверка формы отчета

   

    ' Закрываем файл с отчетом без сохранения изменений (параметр SaveChanges:=False)
    Workbooks(Dir(FileName)).Close SaveChanges:=False
    
    ' Итоговое сообщение
    If finishProcess = True Then
      MsgBox ("Обработка " + Dir(DBstrName_String) + " за " + CStr(maxDateInReport) + " завершена!")
    Else
      MsgBox ("Обработка отчета была прервана!")
    End If



  End If


End Sub

' Отправка письма: отправляю шаблон самому себе для последующей отправки в сеть письма на его основе:
Sub Отправка_Lotus_Notes_Лист5_Карты()
Dim темаПисьма, текстПисьма, hashTag As String
Dim i As Byte
  
  If MsgBox("Отправить себе Шаблон письма?", vbYesNo) = vbYes Then
    
    ' Тема письма - Тема:
    ' темаПисьма = ThisWorkbook.Sheets("Лист5").Cells(RowByValue(ThisWorkbook.Name, "Лист5", "Тема:", 100, 100), ColumnByValue(ThisWorkbook.Name, "Лист5", "Тема:", 100, 100) + 1).Value
    темаПисьма = subjectFromSheet("Лист5")
    
    ' hashTag - Хэштэг:
    ' hashTag = ThisWorkbook.Sheets("Лист5").Cells(RowByValue(ThisWorkbook.Name, "Лист5", "Хэштэг:", 100, 100), ColumnByValue(ThisWorkbook.Name, "Лист5", "Хэштэг:", 100, 100) + 1).Value
    ' hashTag - Хэштэг:
    hashTag = hashTagFromSheet("Лист5")

    ' Текст письма
    текстПисьма = "" + Chr(13)
    текстПисьма = текстПисьма + "" + ThisWorkbook.Sheets("Лист5").Cells(rowByValue(ThisWorkbook.Name, "Лист5", "Список получателей:", 100, 100), ColumnByValue(ThisWorkbook.Name, "Лист5", "Список получателей:", 100, 100) + 2).Value + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "" + getFromAddrBook("РД", 2) + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "Уважаемые сотрудники," + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "Заявки на дебетовые и кредитные карты." + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    ' Визитка (подпись С Ув., )
    текстПисьма = текстПисьма + ПодписьВПисьме()
    ' Хэштег
    текстПисьма = текстПисьма + createBlankStr(20) + hashTag
    ' Вызов
    Call send_Lotus_Notes(темаПисьма, "Sergey Fedorovich Proschaev/Tyumen/PSBank/Ru", "Sergey Fedorovich Proschaev/Tyumen/PSBank/Ru", текстПисьма, "")
  
    ' Сообщение
    MsgBox ("Письмо отправлено!")
     
  End If
  
End Sub

' Проверка отчета (число карт)
Sub Проверка_отчета_по_эмиссии()

' Описание переменных
Dim ReportName_String, officeNameInReport, CheckFormatReportResult As String
Dim i, rowCount As Integer
Dim finishProcess As Boolean
    
  ' Открыть файл с отчетом
  FileName = Application.GetOpenFilename("Excel Files (*.xml), *.xml", , "Открытие файла с отчетом")

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
    ThisWorkbook.Sheets("Лист3").Activate

    ' Проверка формы отчета
    CheckFormatReportResult = CheckFormatReport(ReportName_String, "Список", 5, Date)
    
    If CheckFormatReportResult = "OK" Then
      
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

        rowCount = 0
        Do While Not IsEmpty(Workbooks(ReportName_String).Sheets("Список").Cells(rowCount, 1).Value)
        
          ' Следующая запись
          rowCount = rowCount + 1
          Application.StatusBar = officeNameInReport + ": " + CStr(rowCount) + "..."
          DoEventsInterval (rowCount)
        Loop
   
        ' Выводим данные по офису
      
      Next i ' Следующий офис
      
      ' Выводим итоги обработки
      
      ' Зачеркиваем пункт меню на стартовой страницы
      ' Call ЗачеркиваемТекстВячейке("Лист0", "D9")
      Call ЗачеркиваемТекстВячейке("Лист0", RangeByValue(ThisWorkbook.Name, "Лист0", "Оперативная справка по заявкам на карты", 100, 100))
       
      ' Переменная завершения обработки
      finishProcess = True
    
    Else
      ' Сообщение о неверном формате отчета или даты
      MsgBox ("Проверьте отчет: " + CheckFormatReportResult + "!")
    End If ' Проверка формы отчета

    ' Закрываем файл с отчетом без сохранения изменений (параметр SaveChanges:=False)
    Workbooks(Dir(FileName)).Close SaveChanges:=False
    
    ' Переходим в ячейку M2
    ThisWorkbook.Sheets("Лист3").Range("L78").Select

    ' Строка статуса
    Application.StatusBar = ""

    ' Итоговое сообщение
    If finishProcess = True Then
      MsgBox ("Обработка " + Dir(ReportName_String) + " завершена!")
    Else
      MsgBox ("Обработка отчета была прервана!")
    End If

  End If ' Если файл был выбран


End Sub

' Отчет Cards_emisssion
Sub Cards_emisssion()

' Описание переменных
Dim ReportName_String, officeNameInReport, CheckFormatReportResult As String
Dim i, rowCount As Integer
Dim finishProcess As Boolean
    
  ' Открыть файл с отчетом
  FileName = Application.GetOpenFilename("Excel Files (*.xlsm), *.xlsm", , "Открытие файла с отчетом Cards_emisssion")

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
    ThisWorkbook.Sheets("Лист5").Activate

    ' Формируем список для отправки (в "Список получателей:"):
    ThisWorkbook.Sheets("Лист5").Cells(rowByValue(ThisWorkbook.Name, "Лист5", "Список получателей:", 100, 100), ColumnByValue(ThisWorkbook.Name, "Лист5", "Список получателей:", 100, 100) + 2).Value = getFromAddrBook("УДО2,УДО3,УДО4,УДО5,НОРПиКО1,НОРПиКО2,НОРПиКО3,НОРПиКО4,НОРПиКО5,ПМ,НОКП,РРКК,МПП", 2)

    ' Проверка формы отчета
    CheckFormatReportResult = CheckFormatReport(ReportName_String, "F1", 12, Date)
    If CheckFormatReportResult = "OK" Then
      
      ' Дата отчета Cards_emisssion_29_07_20_(2019)_2
      dateReport_CardsEmission = CDate(Mid(ReportName_String, 17, 2) + "." + Mid(ReportName_String, 20, 2) + ".20" + Mid(ReportName_String, 23, 2))
      ThisWorkbook.Sheets("Лист5").Range("B35").Value = "Эмиссия банковских карт на " + CStr(dateReport_CardsEmission) + " г."
      ThisWorkbook.Sheets("Лист5").Range("R35").Value = "Остатки карт в сейфах и потенциал для активации на " + CStr(dateReport_CardsEmission) + " г."
    
      ' Создаем 1-ый файл исходящий с картами в сейфе
      OutBookName = ThisWorkbook.Path + "\Out\CardsInSafe_" + strDDMMYYYY(dateReport_CardsEmission) + ".xlsx"
      Call createBook_CardsInSafe(OutBookName)
      ThisWorkbook.Sheets("Лист5").Range("R36").Value = OutBookName
            
      ' Создаем 2-ой файл исходящий с картами для активации
      OutBookName2 = ThisWorkbook.Path + "\Out\CardsForActive_" + strDDMMYYYY(dateReport_CardsEmission) + ".xlsx"
      Call createBook_CardsForActive(OutBookName2)
      ThisWorkbook.Sheets("Лист5").Range("T36").Value = OutBookName2
      
      ' Число строк в итоговом OutBookName
      rowCount3 = 1
      
      
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


        rowCount = 7
        Do While Not IsEmpty(Workbooks(ReportName_String).Sheets("Карты в сейфах").Cells(rowCount, 2).Value)
        
          ' Если это текущий офис
          If InStr(Workbooks(ReportName_String).Sheets("Карты в сейфах").Cells(rowCount, 2).Value, officeNameInReport) <> 0 Then
            
            ' Раскрываем столбец 6-ой F "Общий итог"
            Workbooks(ReportName_String).Sheets("Карты в сейфах").Cells(rowCount, 6).ShowDetail = True
          
            ' Переходим на окно DB
            ThisWorkbook.Sheets("Лист5").Activate
          
            ' Номера столбцов - определяем нумерацию:
            ' Офис - DP4_отчет (6) Текст
            Column_DP4 = ColumnByName(Workbooks(ReportName_String).Name, "Лист" + CStr(i), 1, "DP4_отчет")
            ' Номер_клиента - CLIENTPSBID (1) Число
            Column_CLIENTPSBID = ColumnByName(Workbooks(ReportName_String).Name, "Лист" + CStr(i), 1, "CLIENTPSBID")
            ' Дата_выпуска - DATE_ISSUE (23) Дата
            Column_DATE_ISSUE = ColumnByName(Workbooks(ReportName_String).Name, "Лист" + CStr(i), 1, "DATE_ISSUE")
            ' Номер_договора - CONTRACT_NUMBER (3)
            Column_CONTRACT_NUMBER = ColumnByName(Workbooks(ReportName_String).Name, "Лист" + CStr(i), 1, "CONTRACT_NUMBER")
            ' Тариф - TARIFFID (8)
            Column_TARIFFID = ColumnByName(Workbooks(ReportName_String).Name, "Лист" + CStr(i), 1, "TARIFFID")
            ' Продукт - product_name (10)
            column_Product_Name = ColumnByName(Workbooks(ReportName_String).Name, "Лист" + CStr(i), 1, "product_name")
            ' Месяц_выпуска - Month_ISSUE (13)
            Column_Month_ISSUE = ColumnByName(Workbooks(ReportName_String).Name, "Лист" + CStr(i), 1, "Month_ISSUE")
            ' Год_выпуска - Year (14)
            Column_Month_Year = ColumnByName(Workbooks(ReportName_String).Name, "Лист" + CStr(i), 1, "Year")
            ' Вид_заказа - ORDERMODE (15)
            Column_ORDERMODE = ColumnByName(Workbooks(ReportName_String).Name, "Лист" + CStr(i), 1, "ORDERMODE")
            ' ClaimID - CLAIMID (2)
            Column_CLAIMID = ColumnByName(Workbooks(ReportName_String).Name, "Лист" + CStr(i), 1, "CLAIMID")
            ' Статус - CARDSTATUS (16)
            Column_CARDSTATUS = ColumnByName(Workbooks(ReportName_String).Name, "Лист" + CStr(i), 1, "CARDSTATUS")
            ' Плат_система - CARDSORT (17)
            Column_CARDSORT = ColumnByName(Workbooks(ReportName_String).Name, "Лист" + CStr(i), 1, "CARDSORT")
            ' ClaimStatus - CLAIMSTATUS (19)
            Column_CLAIMSTATUS = ColumnByName(Workbooks(ReportName_String).Name, "Лист" + CStr(i), 1, "CLAIMSTATUS")
            ' CardType - CARDTYPE (20)
            Column_CARDTYPE = ColumnByName(Workbooks(ReportName_String).Name, "Лист" + CStr(i), 1, "CARDTYPE")
            ' SalesChannel - SALESCHANNEL (22)
            Column_SALESCHANNEL = ColumnByName(Workbooks(ReportName_String).Name, "Лист" + CStr(i), 1, "SALESCHANNEL")
                      
            ' Число ДК и КК в сейфе
            countDC = 0
            countCC = 0
          
            ' Обрабатываем и выводим в OutBookName карты, находящиеся в сейфах офисов. Для Офиса Тюменский - Лист1, Сургут - Лист2, Нижневартовск - Лист3, Новый Уренгой - Лист4, Тарко-Сале - Лист5
            rowCount2 = 2
            Do While Not IsEmpty(Workbooks(ReportName_String).Sheets("Лист" + CStr(i)).Cells(rowCount2, 1).Value)
              
              ' Строка в итоговом OutBookName
              rowCount3 = rowCount3 + 1
                            
              ' Офис - DP4_отчет (6) Текст
              Workbooks(Dir(OutBookName)).Sheets("Лист1").Cells(rowCount3, 1).Value = cityOfficeName(Workbooks(ReportName_String).Sheets("Лист" + CStr(i)).Cells(rowCount2, Column_DP4).Value)
              ' Номер_клиента - CLIENTPSBID (1) Число
              Workbooks(Dir(OutBookName)).Sheets("Лист1").Cells(rowCount3, 2).Value = Workbooks(ReportName_String).Sheets("Лист" + CStr(i)).Cells(rowCount2, Column_CLIENTPSBID).Value
              ' Дата_выпуска - DATE_ISSUE (23) Дата
              Workbooks(Dir(OutBookName)).Sheets("Лист1").Cells(rowCount3, 3).Value = Workbooks(ReportName_String).Sheets("Лист" + CStr(i)).Cells(rowCount2, Column_DATE_ISSUE).Value
              ' Номер_договора - CONTRACT_NUMBER (3)
              Workbooks(Dir(OutBookName)).Sheets("Лист1").Cells(rowCount3, 4).Value = Workbooks(ReportName_String).Sheets("Лист" + CStr(i)).Cells(rowCount2, Column_CONTRACT_NUMBER).Value
              ' Тариф - TARIFFID (8)
              Workbooks(Dir(OutBookName)).Sheets("Лист1").Cells(rowCount3, 5).Value = Workbooks(ReportName_String).Sheets("Лист" + CStr(i)).Cells(rowCount2, Column_TARIFFID).Value
              ' Продукт - product_name (10)
              Workbooks(Dir(OutBookName)).Sheets("Лист1").Cells(rowCount3, 6).Value = Workbooks(ReportName_String).Sheets("Лист" + CStr(i)).Cells(rowCount2, column_Product_Name).Value
              ' Месяц_выпуска - Month_ISSUE (13)
              Workbooks(Dir(OutBookName)).Sheets("Лист1").Cells(rowCount3, 7).Value = Workbooks(ReportName_String).Sheets("Лист" + CStr(i)).Cells(rowCount2, Column_Month_ISSUE).Value
              ' Год_выпуска - Year (14)
              Workbooks(Dir(OutBookName)).Sheets("Лист1").Cells(rowCount3, 8).Value = Workbooks(ReportName_String).Sheets("Лист" + CStr(i)).Cells(rowCount2, Column_Month_Year).Value
              ' Вид_заказа - ORDERMODE (15)
              Workbooks(Dir(OutBookName)).Sheets("Лист1").Cells(rowCount3, 9).Value = Workbooks(ReportName_String).Sheets("Лист" + CStr(i)).Cells(rowCount2, Column_ORDERMODE).Value
              ' ClaimID - CLAIMID (2)
              Workbooks(Dir(OutBookName)).Sheets("Лист1").Cells(rowCount3, 10).Value = Workbooks(ReportName_String).Sheets("Лист" + CStr(i)).Cells(rowCount2, Column_CLAIMID).Value
              ' Статус - CARDSTATUS (16)
              Workbooks(Dir(OutBookName)).Sheets("Лист1").Cells(rowCount3, 11).Value = Workbooks(ReportName_String).Sheets("Лист" + CStr(i)).Cells(rowCount2, Column_CARDSTATUS).Value
              ' Плат_система - CARDSORT (17)
              Workbooks(Dir(OutBookName)).Sheets("Лист1").Cells(rowCount3, 12).Value = Workbooks(ReportName_String).Sheets("Лист" + CStr(i)).Cells(rowCount2, Column_CARDSORT).Value
              ' ClaimStatus - CLAIMSTATUS (19)
              Workbooks(Dir(OutBookName)).Sheets("Лист1").Cells(rowCount3, 13).Value = Workbooks(ReportName_String).Sheets("Лист" + CStr(i)).Cells(rowCount2, Column_CLAIMSTATUS).Value
              ' CardType - CARDTYPE (20)
              Workbooks(Dir(OutBookName)).Sheets("Лист1").Cells(rowCount3, 14).Value = Workbooks(ReportName_String).Sheets("Лист" + CStr(i)).Cells(rowCount2, Column_CARDTYPE).Value
              ' SalesChannel - SALESCHANNEL (22)
              Workbooks(Dir(OutBookName)).Sheets("Лист1").Cells(rowCount3, 15).Value = Workbooks(ReportName_String).Sheets("Лист" + CStr(i)).Cells(rowCount2, Column_SALESCHANNEL).Value

              ' Считаем карты по типу
              If Workbooks(Dir(OutBookName)).Sheets("Лист1").Cells(rowCount3, 14).Value = "Дебетовая" Then
                countDC = countDC + 1
              End If
              '
              If Workbooks(Dir(OutBookName)).Sheets("Лист1").Cells(rowCount3, 14).Value = "Кредитная" Then
                countCC = countCC + 1
              End If

              
              ' Следующая запись
              rowCount2 = rowCount2 + 1
              Application.StatusBar = officeNameInReport + ": " + CStr(rowCount) + " (" + CStr(rowCount2) + ")..."
              DoEventsInterval (rowCount)
            
            Loop
            
            ' Заносим в Таблицу на "Лист5"
            ' Число дебетовых карт в сейфе
            ThisWorkbook.Sheets("Лист5").Cells(38 + i, 3).Value = countDC
            
            ' Число кредитных карт в сейфе
            ThisWorkbook.Sheets("Лист5").Cells(38 + i, 4).Value = countCC
            
          End If
        
          ' Следующая запись
          rowCount = rowCount + 1
          ' Application.StatusBar = officeNameInReport + ": " + CStr(rowCount) + "..."
          DoEventsInterval (rowCount)
        Loop
   
        ' Выводим данные по офису
      
      Next i ' Следующий офис
      
      Application.StatusBar = ""
      
      ' Закрываем файл с отчетом без сохранения изменений (параметр SaveChanges:=False)
      Workbooks(Dir(FileName)).Close SaveChanges:=False ' - отладка не закрываем
        
      ' Закрываем первый сформированный файлы с картами в сейфе
      Workbooks(Dir(OutBookName)).Close SaveChanges:=True
            
      ' Выводим итоги обработки
      
      ' ******** Второй цикл обработки *********
      
      ' Открываем снова Книгу Card_Emission
      ' Открываем выбранную книгу (UpdateLinks:=0)
      Workbooks.Open FileName, 0
      
      ' Переходим на окно DB
      ThisWorkbook.Sheets("Лист5").Activate

      
      ' ******** На листе "Потенциальные для активации" установка фильтра "ВСЕ" (и ДК и КК)
      Workbooks(ReportName_String).Sheets("Потенциальные для активации").PivotTables("SASApp:CARDS.TRUKHACHEV_EMISS_ACT_ISSUE").PivotFields("Tip").ClearAllFilters
      Workbooks(ReportName_String).Sheets("Потенциальные для активации").PivotTables("SASApp:CARDS.TRUKHACHEV_EMISS_ACT_ISSUE").PivotFields("Tip").CurrentPage = "(All)"
      ' ********
      
      ' Номер записи в выходной книги с картами для активации
      rowCount3 = 1
      ' Номер открываемого листа в PivotTables
      Номер_Листа = 0
            
      ' Цикл по 5-ти офисам (второй для потенциала для активации)
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

      
        ' 1. Первый цикл обработки "Потенциальные для активации"
        строка_Офис = rowByValue(Workbooks(ReportName_String).Name, "Потенциальные для активации", "Офис", 100, 100) ' 12
        столбец_Офис = ColumnByValue(Workbooks(ReportName_String).Name, "Потенциальные для активации", "Офис", 100, 100) ' 2
        столбец_Месяц_выдачи = ColumnByValue(Workbooks(ReportName_String).Name, "Потенциальные для активации", "Месяц выдачи", 100, 100) ' 4
        
        ' Число карт, потенциальных для активации в офисе по типам:
        countDC = 0
        countCC = 0
      
        rowCount = строка_Офис + 1
        Do While Not IsEmpty(Workbooks(ReportName_String).Sheets("Потенциальные для активации").Cells(rowCount, столбец_Офис).Value)
        
          ' Если это текущий офис
          If InStr(Workbooks(ReportName_String).Sheets("Потенциальные для активации").Cells(rowCount, столбец_Офис).Value, officeNameInReport) <> 0 Then
            
            
            ' Движемся по горизонтали вправо до пустой ячейки
            ColumnCount = столбец_Месяц_выдачи
            Do While Not IsEmpty(Workbooks(ReportName_String).Sheets("Потенциальные для активации").Cells(rowCount, ColumnCount).Value)
            
            
              ' Раскрываем столбец
              Workbooks(ReportName_String).Sheets("Потенциальные для активации").Cells(rowCount, ColumnCount).ShowDetail = True
              ' Номер открывшегося листа: Лист1, Лист2, Лист3, ...
              Номер_Листа = Номер_Листа + 1
              ' Переходим на окно DB
              ThisWorkbook.Sheets("Лист5").Activate

              ' Номера столбцов - определяем нумерацию:
              ' Офис (1) - DP4 (16)
              Column_DP4 = ColumnByName(Workbooks(ReportName_String).Name, "Лист" + CStr(Номер_Листа), 1, "DP4")
              ' Номер_клиента (2) - CLIENTPSBID (13)
              Column_CLIENTPSBID = ColumnByName(Workbooks(ReportName_String).Name, "Лист" + CStr(Номер_Листа), 1, "CLIENTPSBID")
              ' Дата_выдачи (3) - DATEISSUEONHAND (9)
              Column_DATEISSUEONHAND = ColumnByName(Workbooks(ReportName_String).Name, "Лист" + CStr(Номер_Листа), 1, "DATEISSUEONHAND")
              ' Номер_договора (4) - CONTRACTNUMBER (14)
              Column_CONTRACTNUMBER = ColumnByName(Workbooks(ReportName_String).Name, "Лист" + CStr(Номер_Листа), 1, "CONTRACTNUMBER")
              ' Тариф (5) - TARIFFID (8)
              Column_TARIFFID = ColumnByName(Workbooks(ReportName_String).Name, "Лист" + CStr(Номер_Листа), 1, "TARIFFID")
              ' Продукт (6) - product_name (12)
              column_Product_Name = ColumnByName(Workbooks(ReportName_String).Name, "Лист" + CStr(Номер_Листа), 1, "product_name")
              ' Месяц_выдачи (7) - Месяц выдачи (1)
              Column_Месяц_выдачи = ColumnByName(Workbooks(ReportName_String).Name, "Лист" + CStr(Номер_Листа), 1, "Месяц выдачи")
              ' CardType (8) - Tip (10)
              Column_Tip = ColumnByName(Workbooks(ReportName_String).Name, "Лист" + CStr(Номер_Листа), 1, "Tip")
              ' SalesChannel (9) - Канал продаж (11)
              Column_Канал_продаж = ColumnByName(Workbooks(ReportName_String).Name, "Лист" + CStr(Номер_Листа), 1, "Канал продаж")
              ' SalesChannel2 (10) - SALESCHANNEL (22)
              Column_SALESCHANNEL = ColumnByName(Workbooks(ReportName_String).Name, "Лист" + CStr(Номер_Листа), 1, "SALESCHANNEL")
              ' Сотрудник (11) - NAMEUSR3 (20)
              Column_NAMEUSR3 = ColumnByName(Workbooks(ReportName_String).Name, "Лист" + CStr(Номер_Листа), 1, "NAMEUSR3")

              ' Обрабатываем текущий открывшийся Лист с номером Номер_Листа и Заносим карты во вторую исходящую таблицу
              rowCount2 = 2
              Do While Not IsEmpty(Workbooks(ReportName_String).Sheets("Лист" + CStr(Номер_Листа)).Cells(rowCount2, 1).Value)
            
                ' Номер записи в выходной книги с картами для активации
                rowCount3 = rowCount3 + 1
                  
                ' Офис (1) - DP4 (16)
                Workbooks(Dir(OutBookName2)).Sheets("Лист1").Cells(rowCount3, 1).Value = cityOfficeName(Workbooks(ReportName_String).Sheets("Лист" + CStr(Номер_Листа)).Cells(rowCount2, Column_DP4).Value)
                ' Номер_клиента (2) - CLIENTPSBID (13)
                Workbooks(Dir(OutBookName2)).Sheets("Лист1").Cells(rowCount3, 2).Value = Workbooks(ReportName_String).Sheets("Лист" + CStr(Номер_Листа)).Cells(rowCount2, Column_CLIENTPSBID).Value
                ' Дата_выдачи (3) - DATEISSUEONHAND (9)
                Workbooks(Dir(OutBookName2)).Sheets("Лист1").Cells(rowCount3, 3).Value = Workbooks(ReportName_String).Sheets("Лист" + CStr(Номер_Листа)).Cells(rowCount2, Column_DATEISSUEONHAND).Value
                ' Номер_договора (4) - CONTRACTNUMBER (14)
                Workbooks(Dir(OutBookName2)).Sheets("Лист1").Cells(rowCount3, 4).Value = Workbooks(ReportName_String).Sheets("Лист" + CStr(Номер_Листа)).Cells(rowCount2, Column_CONTRACTNUMBER).Value
                ' Тариф (5) - TARIFFID (8)
                Workbooks(Dir(OutBookName2)).Sheets("Лист1").Cells(rowCount3, 5).Value = Workbooks(ReportName_String).Sheets("Лист" + CStr(Номер_Листа)).Cells(rowCount2, Column_TARIFFID).Value
                ' Продукт (6) - product_name (12)
                Workbooks(Dir(OutBookName2)).Sheets("Лист1").Cells(rowCount3, 6).Value = Workbooks(ReportName_String).Sheets("Лист" + CStr(Номер_Листа)).Cells(rowCount2, column_Product_Name).Value
                ' Месяц_выдачи (7) - Месяц выдачи (1)
                Workbooks(Dir(OutBookName2)).Sheets("Лист1").Cells(rowCount3, 7).Value = Workbooks(ReportName_String).Sheets("Лист" + CStr(Номер_Листа)).Cells(rowCount2, Column_Месяц_выдачи).Value
                ' CardType (8) - Tip (10)
                Workbooks(Dir(OutBookName2)).Sheets("Лист1").Cells(rowCount3, 8).Value = Workbooks(ReportName_String).Sheets("Лист" + CStr(Номер_Листа)).Cells(rowCount2, Column_Tip).Value
                ' SalesChannel (9) - Канал продаж (11)
                Workbooks(Dir(OutBookName2)).Sheets("Лист1").Cells(rowCount3, 9).Value = Workbooks(ReportName_String).Sheets("Лист" + CStr(Номер_Листа)).Cells(rowCount2, Column_Канал_продаж).Value
                ' SalesChannel2 (10) - SALESCHANNEL (22)
                Workbooks(Dir(OutBookName2)).Sheets("Лист1").Cells(rowCount3, 10).Value = Workbooks(ReportName_String).Sheets("Лист" + CStr(Номер_Листа)).Cells(rowCount2, Column_SALESCHANNEL).Value
                ' Сотрудник (11) - NAMEUSR3 (20)
                Workbooks(Dir(OutBookName2)).Sheets("Лист1").Cells(rowCount3, 11).Value = Фамилия_и_Имя(Workbooks(ReportName_String).Sheets("Лист" + CStr(Номер_Листа)).Cells(rowCount2, Column_NAMEUSR3).Value, 2)
        
                ' Для заполнения (12) - Дата планируемой активации

                ' Для заполнения (13) - Примечание
                
                ' Считаем карты по типу
                If Workbooks(Dir(OutBookName2)).Sheets("Лист1").Cells(rowCount3, 8).Value = "Дебетовые" Then
                  countDC = countDC + 1
                End If
                '
                If Workbooks(Dir(OutBookName2)).Sheets("Лист1").Cells(rowCount3, 8).Value = "Кредитные" Then
                  countCC = countCC + 1
                End If

                ' Следующая запись в ЛистN
                rowCount2 = rowCount2 + 1
                Application.StatusBar = "Потенциальные для активации " + officeNameInReport + ": " + CStr(rowCount2) + "..."
                DoEventsInterval (rowCount2)
              Loop

              ' Следующая запись
              ColumnCount = ColumnCount + 1
              ' Application.StatusBar = "Потенциальные для активации " + officeNameInReport + ": " + CStr(rowCount) + "..."
              DoEventsInterval (rowCount)
            Loop

          End If
        
          ' Следующая запись
          rowCount = rowCount + 1
          ' Application.StatusBar = "Потенциальные для активации " + officeNameInReport + ": " + CStr(rowCount) + "..."
          DoEventsInterval (rowCount)
        Loop
        
        ' Заносим в Таблицу на "Лист5"
        ' Число дебетовых карт для активации
        ThisWorkbook.Sheets("Лист5").Cells(38 + i, 6).Value = countDC
        ' Число кредитных карт для активации
        ThisWorkbook.Sheets("Лист5").Cells(38 + i, 7).Value = countCC

      Next i ' Следующий офис
      
      Application.StatusBar = "Завершение ..."
      
      ' Закрываем файл с отчетом без сохранения изменений (параметр SaveChanges:=False)
      Workbooks(Dir(FileName)).Close SaveChanges:=False ' - отладка не закрываем
      
      ' Закрываем сформированный второй сформированный файл - Потенциал для активации
      Workbooks(Dir(OutBookName2)).Close SaveChanges:=True

      ' Сохранение изменений
      ThisWorkbook.Save
    
      ' Переменная завершения обработки
      finishProcess = True
    Else
      ' Сообщение о неверном формате отчета или даты
      MsgBox ("Проверьте отчет: " + CheckFormatReportResult + "!")
    End If ' Проверка формы отчета

    
    ' Переходим в ячейку M2
    ThisWorkbook.Sheets("Лист5").Range("P42").Select
    ' Workbooks("DB_Result").Sheets("Лист5").Range("P42").Select

    ' Строка статуса
    Application.StatusBar = ""

    ' Зачеркиваем пункт меню на стартовой страницы
    ' Call ЗачеркиваемТекстВячейке("Лист0", RangeByValue(ThisWorkbook.Name, "Лист0", "Оперативная справка по _________________", 100, 100))
    
    ' Перемещение на листе
    Call Карты_к_началу
    Call Карты_к_Cards_emisssion
    
    ' Итоговое сообщение
    If finishProcess = True Then
      MsgBox ("Обработка " + Dir(ReportName_String) + " завершена!")
    Else
      MsgBox ("Обработка отчета была прервана!")
    End If

  End If ' Если файл был выбран

End Sub

' Переход к началу Листа на ПК
Sub Карты_к_началу()
  ThisWorkbook.Sheets("Лист5").Cells(3, 13).Select
End Sub

' Переход к Форме 2 Листа на Карты
Sub Карты_к_Cards_emisssion()
  ThisWorkbook.Sheets("Лист5").Activate
  Call Карты_к_началу
  ActiveWindow.SmallScroll Down:=32
  ThisWorkbook.Sheets("Лист5").Cells(56, 13).Select
End Sub

' Создание книги с остатками в сейфе
Sub createBook_CardsInSafe(In_OutBookName)

    Workbooks.Add
    ActiveWorkbook.SaveAs FileName:=In_OutBookName
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Activate
    
    ' Форматирование полей
    ' Офис - DP4_отчет (6) Текст
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 1).Value = "Офис"
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns("A:A").EntireColumn.ColumnWidth = 20
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 1).HorizontalAlignment = xlCenter
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns("A:A").HorizontalAlignment = xlCenter
    
    ' Номер_клиента - CLIENTPSBID (1) Число
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 2).Value = "Номер_клиента"
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns("B:B").EntireColumn.ColumnWidth = 15
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 2).HorizontalAlignment = xlCenter
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns("B:B").HorizontalAlignment = xlCenter
    
    ' Дата_выпуска - DATE_ISSUE (23) Дата (NumberFormat = "m/d/yyyy")
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 3).Value = "Дата_выпуска"
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns("C:C").EntireColumn.ColumnWidth = 16
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 3).HorizontalAlignment = xlCenter
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns("C:C").HorizontalAlignment = xlCenter
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns("C:C").NumberFormat = "m/d/yyyy"
   
    ' Номер_договора - CONTRACT_NUMBER (3)
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 4).Value = "Номер_договора"
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns("D:D").EntireColumn.ColumnWidth = 18
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 4).HorizontalAlignment = xlCenter
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns("D:D").HorizontalAlignment = xlCenter
        
    ' Тариф - TARIFFID (8)
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 5).Value = "Тариф"
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns("E:E").EntireColumn.ColumnWidth = 10
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 5).HorizontalAlignment = xlCenter
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns("E:E").HorizontalAlignment = xlCenter
    
    ' Продукт - product_name (10)
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 6).Value = "Продукт"
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns("F:F").EntireColumn.ColumnWidth = 32
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 6).HorizontalAlignment = xlCenter
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns("F:F").HorizontalAlignment = xlCenter

    ' Месяц_выпуска - Month_ISSUE (13)
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 7).Value = "Месяц_выпуска"
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns("G:G").EntireColumn.ColumnWidth = 8
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns("G:G").HorizontalAlignment = xlCenter
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 7).HorizontalAlignment = xlLeft
    
    ' Год_выпуска - Year (14)
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 8).Value = "Год_выпуска"
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns("H:H").EntireColumn.ColumnWidth = 8
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns("H:H").HorizontalAlignment = xlCenter
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 8).HorizontalAlignment = xlLeft

    ' Вид_заказа - ORDERMODE (15)
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 9).Value = "Вид_заказа"
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns("I:I").EntireColumn.ColumnWidth = 17
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 9).HorizontalAlignment = xlCenter
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns("I:I").HorizontalAlignment = xlCenter
    
    ' ClaimID - CLAIMID (2)
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 10).Value = "ClaimID"
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns("J:J").EntireColumn.ColumnWidth = 10
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 10).HorizontalAlignment = xlCenter
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns("J:J").HorizontalAlignment = xlCenter
    
    ' Статус - CARDSTATUS (16)
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 11).Value = "Статус"
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns("K:K").EntireColumn.ColumnWidth = 17
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 11).HorizontalAlignment = xlCenter
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns("K:K").HorizontalAlignment = xlCenter
    
    ' Плат_система - CARDSORT (17)
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 12).Value = "Плат_система"
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns("L:L").EntireColumn.ColumnWidth = 20
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 12).HorizontalAlignment = xlCenter
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns("L:L").HorizontalAlignment = xlCenter
    
    ' ClaimStatus - CLAIMSTATUS (19)
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 13).Value = "ClaimStatus"
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns("M:M").EntireColumn.ColumnWidth = 21
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 13).HorizontalAlignment = xlCenter
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns("M:M").HorizontalAlignment = xlCenter
    
    ' CardType - CARDTYPE (20)
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 14).Value = "CardType"
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns("N:N").EntireColumn.ColumnWidth = 12
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 14).HorizontalAlignment = xlCenter
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns("N:N").HorizontalAlignment = xlCenter
    
    ' SalesChannel - SALESCHANNEL (22)
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 15).Value = "SalesChannel"
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns("O:O").EntireColumn.ColumnWidth = 12
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 15).HorizontalAlignment = xlCenter
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns("O:O").HorizontalAlignment = xlCenter
        
    ' Для заполнения - Дата планируемой выдачи
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 16).Value = "Дата планируемой выдачи"
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns("P:P").EntireColumn.ColumnWidth = 25
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 16).HorizontalAlignment = xlCenter
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns("P:P").HorizontalAlignment = xlCenter

    ' Для заполнения - Примечание
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 17).Value = "Комментарий"
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns("Q:Q").EntireColumn.ColumnWidth = 60
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 17).HorizontalAlignment = xlCenter
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns("Q:Q").HorizontalAlignment = xlCenter
    
    
    ' ActiveCell.Offset(0, -4).Columns("A:A").EntireColumn.Select
    ' Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Range("C:C").Select
    ' Числовой
    ' Selection.NumberFormat = "0"
    ' Текстовый
    ' Selection.NumberFormat = "@"

    ' Установка фильтров
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Range("A1:Q1").Select
    Selection.AutoFilter
    
End Sub

' Создание книги с картами для активации
Sub createBook_CardsForActive(In_OutBookName)

    Workbooks.Add
    ActiveWorkbook.SaveAs FileName:=In_OutBookName
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Activate
    
    ' Форматирование полей
    ' Офис - DP4 (16)
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 1).Value = "Офис"
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns("A:A").EntireColumn.ColumnWidth = 20
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 1).HorizontalAlignment = xlCenter
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns("A:A").HorizontalAlignment = xlCenter
    
    ' Номер_клиента - CLIENTPSBID (13)
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 2).Value = "Номер_клиента"
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns("B:B").EntireColumn.ColumnWidth = 20
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 2).HorizontalAlignment = xlCenter
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns("B:B").HorizontalAlignment = xlCenter
    
    ' Дата_выдачи - DATEISSUEONHAND (9)
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 3).Value = "Дата_выдачи"
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns("C:C").EntireColumn.ColumnWidth = 15
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 3).HorizontalAlignment = xlCenter
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns("C:C").HorizontalAlignment = xlCenter
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns("C:C").NumberFormat = "m/d/yyyy"
   
    ' Номер_договора - CONTRACTNUMBER (14)
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 4).Value = "Номер_договора"
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns("D:D").EntireColumn.ColumnWidth = 20
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 4).HorizontalAlignment = xlCenter
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns("D:D").HorizontalAlignment = xlCenter
        
    ' Тариф - TARIFFID (8)
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 5).Value = "Тариф"
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns("E:E").EntireColumn.ColumnWidth = 10
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 5).HorizontalAlignment = xlCenter
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns("E:E").HorizontalAlignment = xlCenter
    
    ' Продукт - product_name (12)
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 6).Value = "Продукт"
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns("F:F").EntireColumn.ColumnWidth = 32
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 6).HorizontalAlignment = xlCenter
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns("F:F").HorizontalAlignment = xlCenter

    ' Месяц_выдачи - Месяц выдачи (1)
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 7).Value = "Месяц_выдачи"
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns("G:G").EntireColumn.ColumnWidth = 15
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 7).HorizontalAlignment = xlCenter
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns("G:G").HorizontalAlignment = xlCenter
    
    ' CardType - Tip (10)
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 8).Value = "CardType"
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns("H:H").EntireColumn.ColumnWidth = 15
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 8).HorizontalAlignment = xlCenter
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns("H:H").HorizontalAlignment = xlCenter
    
    ' SalesChannel - Канал продаж (11)
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 9).Value = "SalesChannel"
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns("I:I").EntireColumn.ColumnWidth = 15
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 9).HorizontalAlignment = xlCenter
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns("I:I").HorizontalAlignment = xlCenter
        
    ' SalesChannel2 - SALESCHANNEL (22)
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 10).Value = "SalesChannel2"
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns("J:J").EntireColumn.ColumnWidth = 15
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 10).HorizontalAlignment = xlCenter
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns("J:J").HorizontalAlignment = xlCenter
        
    ' Сотрудник - NAMEUSR3 (20)
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 11).Value = "Сотрудник"
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns("K:K").EntireColumn.ColumnWidth = 20
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns("K:K").HorizontalAlignment = xlLeft
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 11).HorizontalAlignment = xlCenter
        
    ' Для заполнения - Дата планируемой активации
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 12).Value = "Дата планируемой активации"
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns("L:L").EntireColumn.ColumnWidth = 30
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 12).HorizontalAlignment = xlCenter
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns("P:P").HorizontalAlignment = xlCenter

    ' Для заполнения - Примечание
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 13).Value = "Комментарий"
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns("M:M").EntireColumn.ColumnWidth = 60
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 13).HorizontalAlignment = xlCenter
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns("Q:Q").HorizontalAlignment = xlCenter
    
    
    ' ActiveCell.Offset(0, -4).Columns("A:A").EntireColumn.Select
    ' Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Range("C:C").Select
    ' Числовой
    ' Selection.NumberFormat = "0"
    ' Текстовый
    ' Selection.NumberFormat = "@"

    ' Установка фильтров
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Range("A1:Q1").Select
    Selection.AutoFilter

    
End Sub

' Отправка письма: отправляю шаблон самому себе для последующей отправки в сеть письма на его основе:
Sub Отправка_Lotus_Notes_Лист5_Cards_emisssion()
Dim темаПисьма, текстПисьма, hashTag, attachmentFile As String
Dim i As Byte
  
  If MsgBox("Отправить себе Шаблон письма?", vbYesNo) = vbYes Then
    
    ' Тема письма - Тема:
    ' темаПисьма = ThisWorkbook.Sheets("Лист5").Cells(RowByValue(ThisWorkbook.Name, "Лист5", "Тема:", 100, 100), ColumnByValue(ThisWorkbook.Name, "Лист5", "Тема:", 100, 100) + 1).Value
    темаПисьма = subjectFromSheetII("Лист5", 2)
    
    ' hashTag - Хэштэг:
    ' hashTag = ThisWorkbook.Sheets("Лист5").Cells(RowByValue(ThisWorkbook.Name, "Лист5", "Хэштэг:", 100, 100), ColumnByValue(ThisWorkbook.Name, "Лист5", "Хэштэг:", 100, 100) + 1).Value
    ' hashTag - Хэштэг:
    hashTag = hashTagFromSheetII("Лист5", 2)

    ' Файл-вложение (!!!)
    ' attachmentFile = ThisWorkbook.Sheets("Лист5").Range("R36").Value ' только один файл работает + " " + ThisWorkbook.Sheets("Лист5").Range("T36").Value
    attachmentFile = ""
    
    ' Заготовка:  * - для активации карты используем приветственный Welcome-бонус в размере скидки в 300 руб. на первый платеж в 1000 руб. в теч. первых 14 дней (бонус сохраняется и в период акции первого бесплатного года обслуживания для КК).
    
    ' Текст письма
    текстПисьма = "" + Chr(13)
    текстПисьма = текстПисьма + "" + ThisWorkbook.Sheets("Лист5").Cells(rowByValue(ThisWorkbook.Name, "Лист5", "Список получателей:", 100, 100), ColumnByValue(ThisWorkbook.Name, "Лист5", "Список получателей:", 100, 100) + 2).Value + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "" + getFromAddrBook("РД", 2) + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "Уважаемые руководители," + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "Направляю информацию по остаткам карт в сейфах на " + Mid(ThisWorkbook.Sheets("Лист5").Range("B35").Value, 28, 13) + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "По данным спискам необходимо:" + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "1. Выдать карты из сейфов клиентам:" + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "2. По выданным картам, но не активированным - провести работу по их активации:" + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "3. По невостребованным картам - провести уничтожение (не реже 1 раза в месяц)" + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)

    ' текстПисьма = текстПисьма + "Цели: перекрыть дефицит плана за счет активации карт у клиентов на руках" + Chr(13)
    
    ' For i = 39 To 43
      
    '  ДК
    '   If ThisWorkbook.Sheets("Лист5").Cells(i, 9).Value > 0 Then
        
    '   End If
    '  КК
    '   If ThisWorkbook.Sheets("Лист5").Cells(i, 10).Value > 0 Then
    '
    '   End If
      
      ' Выводим в текст письма
      ' текстПисьма = текстПисьма + "- " + ThisWorkbook.Sheets("Лист5").Cells(i, 2).Value + " ДК шт., КК шт." + Chr(13)
      
    ' Next i
    
    текстПисьма = текстПисьма + "" + Chr(13)
    
    ' Визитка (подпись С Ув., )
    текстПисьма = текстПисьма + ПодписьВПисьме()
    ' Хэштег
    текстПисьма = текстПисьма + createBlankStr(20) + hashTag
    ' Вызов
    ' Call send_Lotus_Notes(темаПисьма, "Sergey Fedorovich Proschaev/Tyumen/PSBank/Ru", "Sergey Fedorovich Proschaev/Tyumen/PSBank/Ru", текстПисьма, "")
    Call send_Lotus_Notes(темаПисьма, "Sergey Fedorovich Proschaev/Tyumen/PSBank/Ru", "Sergey Fedorovich Proschaev/Tyumen/PSBank/Ru", текстПисьма, attachmentFile)
  
    ' Сообщение
    MsgBox ("Письмо отправлено!")
     
  End If
  
End Sub

