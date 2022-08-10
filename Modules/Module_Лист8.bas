Attribute VB_Name = "Module_Лист8"
' *** Лист 8 ***

' *** Глобальные переменные ***
Public numStr_Лист8 As Integer
Public Порядковый_Номер_продукта_на_Лист8 As Byte
Public Порядковый_Номер_продукта_Дробь_на_Лист8 As Byte
Public Строка_нет_листа_в_DB As String
Public Строка_нет_столбца_на_листе_в_DB As String
Public dateDB As Date
Public Первый_день_недели As Boolean
Public Первый_день_недели_Date As Date


' ***                       ***

' Показатели из DB
Sub Показатели_из_DB()
      
' Описание переменных
Dim ReportName_String, officeNameInReport, CheckFormatReportResult, ID_RecVar, StringInSheet, SheetName_String, Наименование_столбца_план As String
Dim i, rowCount, row_DP3_отчет, column_TAB_OK, column_ФИО, column_DP3_отчет, column_DP4_отчет, recInЛист7, порядковый_номер, ном_стр_офис As Integer
Dim finishProcess As Boolean
' Dim dateDB As Date
        
        
  ' Открыть файл с отчетом
  FileName = Application.GetOpenFilename("Excel Files (*.xlsm), *.xlsm", , "Открытие файла с отчетом")

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
    ThisWorkbook.Sheets("Лист8").Activate
    
    ' Проверка формы отчета
    CheckFormatReportResult = CheckFormatReport(ReportName_String, "Оглавление", 1, Date)
    If CheckFormatReportResult = "OK" Then
      
      ' Дата Дашбоарда
      dateDB = CDate(Mid(Workbooks(ReportName_String).Sheets("Оглавление").Cells(1, 1).Value, 23, 10))
      dateDB_Лист7 = CDate(Mid(ThisWorkbook.Sheets("Лист7").Range("B5").Value, 40, 10))

      ' Выставляем - Цель "На неделю:" в "M9"
      Call Цель_на_неделю_Лист8
      ' В N9
      ThisWorkbook.Sheets("Лист8").Range("N9").Value = "Исп.неделя"

      ' Заголовок отчета (из A1 - Отчет по состоянию на 07.07.2020)
      ThisWorkbook.Sheets("Лист8").Cells(5, 2).Value = "Ежедневный отчет по продажам (Dashboard_РБ_New) от " + Mid(Workbooks(ReportName_String).Sheets("Оглавление").Cells(1, 1).Value, 23, 10) + " г."
      ThisWorkbook.Sheets("Лист8").Range("P2").Value = Replace(ThisWorkbook.Sheets("Лист8").Cells(5, 2).Value, " (Dashboard_РБ_New)", "")
            
      ' В M8 остаток дней от текущей даты
      ' Остаток рабочих дней определяем число рабочих дней с понеделника до конца месяца Working_days_between_dates(In_DateStart, In_DateEnd, In_working_days_in_the_week) As Integer
      Остаток_рабочих_дней = Working_days_between_dates(Date, Date_last_day_month(Date), 5)
      '
      ThisWorkbook.Sheets("Лист8").Range("M8").Value = "Дней: " + CStr(Остаток_рабочих_дней)
        
      ' Неделя по текущей дате
      ThisWorkbook.Sheets("Лист8").Range("L5").Value = WeekNumber(Date)
        
      ' Инициализация переменной
      Строка_нет_листа_в_DB = ""
      Строка_нет_столбца_на_листе_в_DB = ""
        
      ' Загружаем изменения за неделю
      ' Ориентир - Месяц
      ' ThisWorkbook.Sheets("Лист8").Range("O9").Value = PreviousWeek(dateDB)
      ' Ориентир - Q
      ThisWorkbook.Sheets("Лист8").Range("O9").Value = PreviousWeek2(dateDB)
        
      ' Очищаем ячейки отчета
      ' Рейтинг (было A-M, сделал до T)
      Call clearСontents2(ThisWorkbook.Name, "Лист8", "A" + CStr(getRowFromSheet8("Интегральный рейтинг по офисам", "Интегральный рейтинг по офисам") + 3), "T" + CStr(getRowFromSheet8("Итого по РОО «Тюменский»", "Итого по РОО «Тюменский»") - 2))
      
      ' 1. ОО «Тюменский»
      Call clearСontents2(ThisWorkbook.Name, "Лист8", "A" + CStr(getRowFromSheet8("ОО «Тюменский»", "ОО «Тюменский»") + 3), "T" + CStr(getRowFromSheet8("ОО «Сургутский»", "ОО «Сургутский»") - 1))
      
      ' 2. ОО «Сургутский»
      Call clearСontents2(ThisWorkbook.Name, "Лист8", "A" + CStr(getRowFromSheet8("ОО «Сургутский»", "ОО «Сургутский»") + 3), "T" + CStr(getRowFromSheet8("ОО «Нижневартовский»", "ОО «Нижневартовский»") - 1))
      
      ' 3. ОО «Нижневартовский»
      Call clearСontents2(ThisWorkbook.Name, "Лист8", "A" + CStr(getRowFromSheet8("ОО «Нижневартовский»", "ОО «Нижневартовский»") + 3), "T" + CStr(getRowFromSheet8("ОО «Новоуренгойский»", "ОО «Новоуренгойский»") - 1))
      
      ' 4. ОО «Новоуренгойский»
      Call clearСontents2(ThisWorkbook.Name, "Лист8", "A" + CStr(getRowFromSheet8("ОО «Новоуренгойский»", "ОО «Новоуренгойский»") + 3), "T" + CStr(getRowFromSheet8("ОО «Тарко-Сале»", "ОО «Тарко-Сале»") - 1))
      
      ' 5. ОО «Тарко-Сале»
      Call clearСontents2(ThisWorkbook.Name, "Лист8", "A" + CStr(getRowFromSheet8("ОО «Тарко-Сале»", "ОО «Тарко-Сале»") + 3), "T" + CStr(getRowFromSheet8("Интегральный рейтинг по офисам", "Интегральный рейтинг по офисам") - 2))
      
      ' 6. РОО Тюменский
      Call clearСontents2(ThisWorkbook.Name, "Лист8", "A" + CStr(getRowFromSheet8("Итого по РОО «Тюменский»", "Итого по РОО «Тюменский»") + 3), "T" + CStr(getRowFromSheet8("Итого по РОО «Тюменский»", "Итого по РОО «Тюменский»") + (getRowFromSheet8("ОО «Сургутский»", "ОО «Сургутский»") - getRowFromSheet8("ОО «Тюменский»", "ОО «Тюменский»"))))
      
      ' Обнуление итоговых значений по Офисам и РОО
      Call Обнуление_итоговых_значений_по_Офисам_и_РОО
              
      ' Открытие офисов РОО на листе "Интегральный рейтинг_Регионы"?
      StringInSheet = "Интегральный рейтинг_Регионы"
      SheetName_String = FindNameSheet(ReportName_String, StringInSheet)
      If SheetName_String <> "" Then
    
        ' Открываем на листе "3. Интегральный рейтинг_Регионы" данные по офисам ОО
        Workbooks(ReportName_String).Sheets(SheetName_String).PivotTables("Сводная таблица1").PivotFields("DP3_отчет_new").PivotItems("Тюменский ОО1").ShowDetail = True
        
        ' Открываем сводную "Лист1" по показателям
        row_Тюменский_ОО1 = rowByValue(ReportName_String, SheetName_String, "Тюменский ОО1", 100, 100)
        column_Тюменский_ОО1 = ColumnByValue(ReportName_String, SheetName_String, "Тюменский ОО1", 300, 300)
        Workbooks(ReportName_String).Sheets(SheetName_String).Cells(row_Тюменский_ОО1, column_Тюменский_ОО1 + 1).ShowDetail = True
        
      Else
        ' Если в DB Лист не найден
        Call в_DB_Лист_не_найден(StringInSheet)
      End If
        
      ' Заголовки
      For i = 1 To 6
        ' Номера офисов от 1 до 5
        Select Case i
          Case 1 ' ОО «Тюменский»
            officeNameInReport = "ОО «Тюменский»"
          Case 2 ' ОО «Сургутский»
            officeNameInReport = "ОО «Сургутский»"
          Case 3 ' ОО «Нижневартовский»
            officeNameInReport = "ОО «Нижневартовский»"
          Case 4 ' ОО «Новоуренгойский»
            officeNameInReport = "ОО «Новоуренгойский»"
          Case 5 ' ОО «Тарко-Сале»
            officeNameInReport = "ОО «Тарко-Сале»"
          Case 6 ' Итого по РОО «Тюменский»
            officeNameInReport = "Итого по РОО «Тюменский»"
            
        End Select
        
        ' Находим номер строки с наименованием офиса
        row_офис = getRowFromSheet8(officeNameInReport, officeNameInReport)
        
        ThisWorkbook.Sheets("Лист8").Cells(row_офис + 1, 5).Value = quarterName(CDate(Mid(Workbooks(ReportName_String).Sheets("Оглавление").Cells(1, 1).Value, 23, 10))) ' 3 кв. 2020 г.
        ThisWorkbook.Sheets("Лист8").Cells(row_офис + 1, 9).Value = "Месяц (" + ИмяМесяца(CDate(Mid(Workbooks(ReportName_String).Sheets("Оглавление").Cells(1, 1).Value, 23, 10))) + ")" 'Месяц (сентябрь)
        ThisWorkbook.Sheets("Лист8").Cells(row_офис + 2, 6).Value = "Факт на " + Mid(Workbooks(ReportName_String).Sheets("Оглавление").Cells(1, 1).Value, 23, 5)
        ThisWorkbook.Sheets("Лист8").Cells(row_офис + 2, 10).Value = ThisWorkbook.Sheets("Лист8").Cells(9, 6).Value
        
      Next i
      
      ' Открываем BASE\Sales
      OpenBookInBase ("Sales_Office")
            
      ' Открываем BASE\Sales
      OpenBookInBase ("Sales")
            
      ' Открываем BASE\Products
      OpenBookInBase ("Products")
            
            
      ' Находим номер строки с наименованием офиса
      row_ОО_Тюменский = getRowFromSheet8("ОО «Тюменский»", "ОО «Тюменский»")
      row_ОО_Сургутский = getRowFromSheet8("ОО «Сургутский»", "ОО «Сургутский»")
      Размер_блока_офиса = row_ОО_Сургутский - row_ОО_Тюменский
            
      ' Цикл по 5-ти офисам
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

        ' Порядковый номер строки в блоке офиса (начинается с нуля для каждого нового офиса)
        ном_стр_офис = 0
      
        ' Нумерация осуществляется через описанную функцию НумерацияПунктов
        Порядковый_Номер_продукта_на_Лист8 = 0
      
        ' Показатель №1 Зарплатные карты (ном_стр_офис = 1)
        StringInSheet = "Зарплатные карты"
        SheetName_String = FindNameSheet(ReportName_String, StringInSheet) ' "3.11 Зарплатные карты"
        If SheetName_String <> "" Then
          
          ном_стр_офис = ном_стр_офис + 1
          Call DB_UniversalSheetInDB(ReportName_String, _
                                     SheetName_String, _
                                       officeNameInReport, _
                                         (row_ОО_Тюменский + 2 + ном_стр_офис) + ((i - 1) * Размер_блока_офиса), _
                                           НумерацияПунктов("Зарплатные карты 18+"), _
                                             "Зарплатные карты 18+", _
                                               "ЗП", _
                                                 "шт.", _
                                                    0, _
                                                     "", _
                                                       "Продажи ЗП 18+, шт._Квартал ", _
                                                         0, _
                                                           "Алтайский ОО1", _
                                                             0, _
                                                               0, _
                                                                 0, _
                                                                   0, 1, 1)
        
          ном_стр_офис = ном_стр_офис + 1
          Call DB_UniversalSheetInDB(ReportName_String, _
                                     SheetName_String, _
                                       officeNameInReport, _
                                         (row_ОО_Тюменский + 2 + ном_стр_офис) + ((i - 1) * Размер_блока_офиса), _
                                           НумерацияПунктов("Зарплатные карты 18+"), _
                                             "Портфель ЗП 18+, шт._Квартал ", _
                                               "Портфель_ЗП", _
                                                 "шт.", _
                                                   0.2, _
                                                     "", _
                                                       "Портфель ЗП 18+, шт._Квартал ", _
                                                         0, _
                                                           "Алтайский ОО1", _
                                                             0, _
                                                               0, _
                                                                 0, _
                                                                   0, 1, 1)
        
        
        Else
          ' Если в DB Лист не найден
          Call в_DB_Лист_не_найден(StringInSheet)
          ' Заносим StringInSheet в переменную Строка_нет_листа_в_DB
          ' Call в_DB_Лист_не_найден(StringInSheet)
        End If

        
        ' ----------------------------------------------------------------------------------------------------------------------------------
        ' Чертим горизонтальную линию 2 (указываем предидущее значение строки + 1)
        Call gorizontalLineII(ThisWorkbook.Name, "Лист8", (row_ОО_Тюменский + 2 + ном_стр_офис) + ((i - 1) * Размер_блока_офиса) + 1, 2, 12)
        ' ----------------------------------------------------------------------------------------------------------------------------------

        ' Показатель №2 ПК (ном_стр_офис = 2)
        StringInSheet = "Потребительские  кредиты"
        SheetName_String = FindNameSheet(ReportName_String, StringInSheet) ' "3.1 Потребительские  кредиты"
        If SheetName_String <> "" Then
        
          ном_стр_офис = ном_стр_офис + 1
          Call DB_UniversalSheetInDB(ReportName_String, _
                                     SheetName_String, _
                                       officeNameInReport, _
                                         (row_ОО_Тюменский + 2 + ном_стр_офис) + ((i - 1) * Размер_блока_офиса), _
                                           НумерацияПунктов("Потребительские кредиты"), _
                                             "Потребительские кредиты", _
                                               "ПК", _
                                                 "тыс.руб.", _
                                                   0.2, _
                                                     "Выдачи, тыс.руб._Месяц", _
                                                       "Выдачи, тыс.руб._Квартал", _
                                                         4, _
                                                           "Алтайский ОО1", _
                                                             0, _
                                                               0, _
                                                                 0, _
                                                                   0, 1, 1)
          
          ' В т.ч. ПК DSA
          ном_стр_офис = ном_стр_офис + 1
          Call DB_UniversalSheetInDB(ReportName_String, _
                                     SheetName_String, _
                                       officeNameInReport, _
                                         (row_ОО_Тюменский + 2 + ном_стр_офис) + ((i - 1) * Размер_блока_офиса), _
                                           НумерацияПунктов("в т.ч. ПК DSA"), _
                                             "в т.ч. ПК DSA", _
                                               "ПК_DSA", _
                                                 "тыс.руб.", _
                                                    0, _
                                                     " DSA_Выдачи, тыс.руб._Месяц", _
                                                       " DSA Выдачи, тыс.руб._Квартал", _
                                                         6, _
                                                           "Алтайский ОО1", _
                                                             0, _
                                                               0, _
                                                                 0, _
                                                                   0, 1, 1)
                                                                   
          ' ПК Выдачи, шт.
          ном_стр_офис = ном_стр_офис + 1
          Call DB_UniversalSheetInDB(ReportName_String, _
                                     SheetName_String, _
                                       officeNameInReport, _
                                         (row_ОО_Тюменский + 2 + ном_стр_офис) + ((i - 1) * Размер_блока_офиса), _
                                           НумерацияПунктов("Выдачи ПК"), _
                                             "Выдачи ПК", _
                                               "Выдачи_ПК_шт", _
                                                 "шт.", _
                                                    0, _
                                                     "Выдачи, шт._Месяц", _
                                                       "Выдачи, шт._Квартал", _
                                                         4, _
                                                           "Алтайский ОО1", _
                                                             0, _
                                                               0, _
                                                                 0, _
                                                                   0, 1, 1)

          
          ' (3) Делаем расчет План/Факт по текущему МЕСЯЦ с Лист7 в офисном канале для текущего officeNameInReport
          ' Пока выключаем
          If False Then
          
            ном_стр_офис = ном_стр_офис + 1 ' (3)
            ' Проверяем условие - даты обработки Дашбоарда на Лист7 = Лист8
            ' If dateDB = dateDB_Лист7 Then
          
            Call План_Факт_ПК_Лист7(officeNameInReport, _
                                      (row_ОО_Тюменский + 2 + ном_стр_офис) + ((i - 1) * Размер_блока_офиса), _
                                        "2.1", _
                                          dateDB)
          
            ' (3) Делаем расчет План/Факт по КВАРТАЛ из BASE\Sales в офисном канале для текущего officeNameInReport
            Call План_Факт_Q_ПК_Sales(officeNameInReport, _
                                      (row_ОО_Тюменский + 2 + ном_стр_офис) + ((i - 1) * Размер_блока_офиса), _
                                        "2.1", _
                                          dateDB)
          End If
        
          
        Else
          ' Если в DB Лист не найден
          Call в_DB_Лист_не_найден(StringInSheet)
        End If
                                                                   
                                                                   
                                                                   
        ' Показатель "Заявки ПК"
        StringInSheet = "Потребительские  кредиты"
        SheetName_String = FindNameSheet(ReportName_String, StringInSheet) ' "3.1 Потребительские  кредиты"
        If SheetName_String <> "" Then
        
          ' Если на листе "Потребительские  кредиты" найдена " % Проникновение страховок_Итого * (Заявка Офис-Выдача Офис+Заявка Офис-Выдача ИБ+Заявка ИБ-Выдача Офис)", то это новая версия отчета
          If ColumnByValue(ReportName_String, SheetName_String, " % Проникновение страховок_Итого * (Заявка Офис-Выдача Офис+Заявка Офис-Выдача ИБ+Заявка ИБ-Выдача Офис)", 1000, 1000) <> 0 Then

            ном_стр_офис = ном_стр_офис + 1
            Call DB_UniversalSheetInDB(ReportName_String, _
                                     SheetName_String, _
                                       officeNameInReport, _
                                         (row_ОО_Тюменский + 2 + ном_стр_офис) + ((i - 1) * Размер_блока_офиса), _
                                           НумерацияПунктов("Заявки ПК"), _
                                             "Заявки ПК", _
                                               "Заявки_ПК", _
                                                 "шт.", _
                                                   0, _
                                                     "Заявки_Месяц", _
                                                       "Заявки_Квартал", _
                                                         0, _
                                                           "Филиал", _
                                                             0, _
                                                               0, _
                                                                 0, _
                                                                   0, 1, 1)
          End If
          
        Else
          ' Если в DB Лист не найден
          Call в_DB_Лист_не_найден(StringInSheet)
        End If
             
             
        ' Показатель "ПК PA_Готовое_решение" - пропал в DB с конца мая - начала июня 2021
        ' StringInSheet = "Потребительские  кредиты"
        ' SheetName_String = FindNameSheet(ReportName_String, StringInSheet) ' "3.1 Потребительские  кредиты"
        ' If SheetName_String <> "" Then
        
        '   ' Если на листе "Потребительские  кредиты" найдена " % Проникновение страховок_Итого * (Заявка Офис-Выдача Офис+Заявка Офис-Выдача ИБ+Заявка ИБ-Выдача Офис)", то это новая версия отчета
        '   If ColumnByValue(ReportName_String, SheetName_String, " % Проникновение страховок_Итого * (Заявка Офис-Выдача Офис+Заявка Офис-Выдача ИБ+Заявка ИБ-Выдача Офис)", 1000, 1000) <> 0 Then
        
        '     ном_стр_офис = ном_стр_офис + 1
        '     Call DB_UniversalSheetInDB(ReportName_String, _
        '                              SheetName_String, _
        '                                officeNameInReport, _
        '                                  (row_ОО_Тюменский + 2 + ном_стр_офис) + ((i - 1) * Размер_блока_офиса), _
        '                                    НумерацияПунктов("ПК PA_Готовое_решение"), _
        '                                      "ПК PA_Готовое_решение", _
        '                                        "ПА_PA_Гот_реш", _
        '                                          "шт.", _
        '                                            0, _
        '                                              "PA_Готовое решение", _
        '                                                "", _
        '                                                  0, _
        '                                                    "Филиал", _
        '                                                      0, _
        '                                                        0, _
        '                                                          0, _
        '                                                            0, 1, 1)
        '   End If
          
        ' Else
          ' Если в DB Лист не найден
        '   Call в_DB_Лист_не_найден(StringInSheet)
        ' End If

        ' Показатель "Проникновение страховок в ПК"
        StringInSheet = "Потребительские  кредиты"
        SheetName_String = FindNameSheet(ReportName_String, StringInSheet) ' "3.1 Потребительские  кредиты"
        If SheetName_String <> "" Then
        
          ' Если на листе "Потребительские  кредиты" найдена " % Проникновение страховок_Итого * (Заявка Офис-Выдача Офис+Заявка Офис-Выдача ИБ+Заявка ИБ-Выдача Офис)", то это новая версия отчета
          If ColumnByValue(ReportName_String, SheetName_String, " % Проникновение страховок_Итого * (Заявка Офис-Выдача Офис+Заявка Офис-Выдача ИБ+Заявка ИБ-Выдача Офис)", 1000, 1000) <> 0 Then

            ном_стр_офис = ном_стр_офис + 1
            ' Call DB_UniversalSheetInDB(ReportName_String, _
            '                          SheetName_String, _
            '                           officeNameInReport, _
            '                             (row_ОО_Тюменский + 2 + ном_стр_офис) + ((i - 1) * Размер_блока_офиса), _
            '                               НумерацияПунктов("Проникновение СЖиЗ в ПК"), _
            '                                 "Проникновение СЖиЗ в ПК", _
            '                                   "СЖиЗ_ПК_%", _
            '                                     "%", _
            '                                       0, _
            '                                         " % Проникновение страховок_Итого * (Заявка Офис-Выдача Офис+Заявка Офис-Выдача ИБ+Заявка ИБ-Выдача Офис)", _
            '                                           " % Проникновение страховок_Итого * (Заявка Офис-Выдача Офис+Заявка Офис-Выдача ИБ+Заявка ИБ-Выдача Офис) ", _
            '                                             0, _
            '                                               "Филиал", _
            '                                                 -1, _
            '                                                   19, _
            '                                                     85, _
            '                                                       85, 1, 1)
            
            Call DB_UniversalSheetInDB(ReportName_String, _
                                     SheetName_String, _
                                       officeNameInReport, _
                                         (row_ОО_Тюменский + 2 + ном_стр_офис) + ((i - 1) * Размер_блока_офиса), _
                                           НумерацияПунктов("Проникновение СЖиЗ в ПК"), _
                                             "Проникновение СЖиЗ в ПК", _
                                               "СЖиЗ_ПК_%", _
                                                 "%", _
                                                   0, _
                                                     "Заявки_Квартал", _
                                                       "Заявки_Квартал", _
                                                         0, _
                                                           "Алтайский ОО1", _
                                                              34, _
                                                                47, _
                                                                  85, _
                                                                    85, 1, 1)
          
          End If
          
        Else
          ' Если в DB Лист не найден
          Call в_DB_Лист_не_найден(StringInSheet)
        End If


        ' Показатель " AR %_Офис " (месяц)
        StringInSheet = "Потребительские  кредиты"
        SheetName_String = FindNameSheet(ReportName_String, StringInSheet) '
        If SheetName_String <> "" Then

          ном_стр_офис = ном_стр_офис + 1
          НумерацияПунктов_ПК_AR_офис = НумерацияПунктов("ПК AR офис")
          Call DB_getParamFromUniversalSheetInDB(ReportName_String, _
                                                   SheetName_String, _
                                                     "Алтайский ОО1", _
                                                       officeNameInReport, _
                                                         (row_ОО_Тюменский + 2 + ном_стр_офис) + ((i - 1) * Размер_блока_офиса), _
                                                           НумерацияПунктов_ПК_AR_офис, _
                                                             "ПК AR офис", _
                                                               " AR %_Офис ", _
                                                                 "ПК_AR_офис", _
                                                                   "%", _
                                                                     0, _
                                                                      "Месяц")
          
          Call DB_getParamFromUniversalSheetInDB2(ReportName_String, _
                                                   SheetName_String, _
                                                     "Алтайский ОО1", _
                                                       officeNameInReport, _
                                                         (row_ОО_Тюменский + 2 + ном_стр_офис) + ((i - 1) * Размер_блока_офиса), _
                                                           НумерацияПунктов_ПК_AR_офис, _
                                                             "ПК AR офис", _
                                                               " AR %_Офис ", _
                                                                 "ПК_AR_офис", _
                                                                   "%", _
                                                                     0, _
                                                                      "Квартал", _
                                                                        14)
        
        Else
          ' Если в DB Лист не найден
          Call в_DB_Лист_не_найден(StringInSheet)
        End If

        ' Показатель "Утилизация лимита %"
        StringInSheet = "Потребительские  кредиты"
        SheetName_String = FindNameSheet(ReportName_String, StringInSheet) '
        If SheetName_String <> "" Then

          ном_стр_офис = ном_стр_офис + 1
          Call DB_getParamFromUniversalSheetInDB(ReportName_String, _
                                                   SheetName_String, _
                                                     "Алтайский ОО1", _
                                                       officeNameInReport, _
                                                         (row_ОО_Тюменский + 2 + ном_стр_офис) + ((i - 1) * Размер_блока_офиса), _
                                                           НумерацияПунктов("Утилизация лимита %"), _
                                                             "Утилизация лимита ПК", _
                                                               " Утилизация лимита %", _
                                                                 "Утилизация_лимита_ПК", _
                                                                   "%", _
                                                                     0, _
                                                                      "Квартал")
        
        Else
          ' Если в DB Лист не найден
          Call в_DB_Лист_не_найден(StringInSheet)
        End If


        ' ----------------------------------------------------------------------------------------------------------------------------------
        ' Чертим горизонтальную линию 2 (указываем предидущее значение строки + 1)
        Call gorizontalLineII(ThisWorkbook.Name, "Лист8", (row_ОО_Тюменский + 2 + ном_стр_офис) + ((i - 1) * Размер_блока_офиса) + 1, 2, 12)
        ' ----------------------------------------------------------------------------------------------------------------------------------
            
            
        ' Показатель "Кредитные карты"
        StringInSheet = "Кредитные карты"
        SheetName_String = FindNameSheet(ReportName_String, StringInSheet) ' "3.2 Кредитные карты": "Активированные карты, шт._Месяц", "Активированные карты, шт._Квартал"
        If SheetName_String <> "" Then
                
          ном_стр_офис = ном_стр_офис + 1
          Call DB_UniversalSheetInDB(ReportName_String, _
                                     SheetName_String, _
                                       officeNameInReport, _
                                         (row_ОО_Тюменский + 2 + ном_стр_офис) + ((i - 1) * Размер_блока_офиса), _
                                           НумерацияПунктов("Кредитные карты (актив.)"), _
                                             "Кредитные карты (актив.)", _
                                               "КК", _
                                                 "шт.", _
                                                   0.1, _
                                                     "Активированные карты, шт._Месяц_Итого", _
                                                       "Активированные карты, шт._Квартал_Итого", _
                                                        4, _
                                                          "Алтайский ОО1", _
                                                             0, _
                                                               0, _
                                                                 0, _
                                                                   0, 1, 1)
        
          ном_стр_офис = ном_стр_офис + 1
          Call DB_UniversalSheetInDB(ReportName_String, _
                                     SheetName_String, _
                                       officeNameInReport, _
                                         (row_ОО_Тюменский + 2 + ном_стр_офис) + ((i - 1) * Размер_блока_офиса), _
                                           НумерацияПунктов("Кредитные карты (актив.)"), _
                                             "в т.ч. КК сеть", _
                                               "КК_сеть", _
                                                 "шт.", _
                                                   0, _
                                                     "Активированные карты, шт._Месяц_Сеть (в т.ч. CRM, DIGITAL, КЦ)", _
                                                       "Активированные карты, шт._Квартал_Сеть (в т.ч. CRM, DIGITAL, КЦ)", _
                                                        6, _
                                                          "Алтайский ОО1", _
                                                             0, _
                                                               0, _
                                                                 0, _
                                                                   0, 1, 1)
        
          ном_стр_офис = ном_стр_офис + 1
          Call DB_UniversalSheetInDB(ReportName_String, _
                                     SheetName_String, _
                                       officeNameInReport, _
                                         (row_ОО_Тюменский + 2 + ном_стр_офис) + ((i - 1) * Размер_блока_офиса), _
                                           НумерацияПунктов("Кредитные карты (актив.)"), _
                                             "           КК DSA", _
                                               "КК_DSA", _
                                                 "шт.", _
                                                   0, _
                                                     "Активированные карты, шт._Месяц_DSA", _
                                                       "Активированные карты, шт._Квартал_DSA", _
                                                        4, _
                                                          "Алтайский ОО1", _
                                                             0, _
                                                               0, _
                                                                 0, _
                                                                   0, 1, 1)
        
          ном_стр_офис = ном_стр_офис + 1
          Call DB_UniversalSheetInDB(ReportName_String, _
                                     SheetName_String, _
                                       officeNameInReport, _
                                         (row_ОО_Тюменский + 2 + ном_стр_офис) + ((i - 1) * Размер_блока_офиса), _
                                           НумерацияПунктов("Кредитные карты (актив.)"), _
                                             "           КК OPC", _
                                               "КК_OPC", _
                                                 "шт.", _
                                                   0, _
                                                     "Активированные карты, шт._Месяц_OPC", _
                                                       "Активированные карты, шт._Квартал_OPC", _
                                                        4, _
                                                          "Алтайский ОО1", _
                                                             0, _
                                                               0, _
                                                                 0, _
                                                                   0, 1, 1)
        
          ном_стр_офис = ном_стр_офис + 1
          Call DB_UniversalSheetInDB(ReportName_String, _
                                     SheetName_String, _
                                       officeNameInReport, _
                                         (row_ОО_Тюменский + 2 + ном_стр_офис) + ((i - 1) * Размер_блока_офиса), _
                                           НумерацияПунктов("КК к ЗП (актив.)"), _
                                             "           КК к ЗП", _
                                               "КК_ЗП_актив", _
                                                 "шт.", _
                                                    0, _
                                                     "Активированные карты, шт._Месяц_КК к ЗП", _
                                                       "Активированные карты, шт._Квартал_КК к ЗП", _
                                                        4, _
                                                          "Алтайский ОО1", _
                                                             0, _
                                                               0, _
                                                                 0, _
                                                                   0, 1, 1)
        
          ном_стр_офис = ном_стр_офис + 1
          Call DB_UniversalSheetInDB(ReportName_String, _
                                     SheetName_String, _
                                       officeNameInReport, _
                                         (row_ОО_Тюменский + 2 + ном_стр_офис) + ((i - 1) * Размер_блока_офиса), _
                                           НумерацияПунктов("КК к ЗП (актив.)"), _
                                             "           КК к Ипотеке", _
                                               "КК_Ипотека", _
                                                 "шт.", _
                                                    0, _
                                                     "Активированные карты, шт._Месяц_КК к Ипотеке", _
                                                       "Активированные карты, шт._Квартал_КК к Ипотеке", _
                                                        4, _
                                                          "Алтайский ОО1", _
                                                             0, _
                                                               0, _
                                                                 0, _
                                                                   0, 1, 1)
        
        
        
        Else
          ' Если в DB Лист не найден
          Call в_DB_Лист_не_найден(StringInSheet)
        End If


        ' Показатель "Сплиты к ПК" - выдача
        StringInSheet = "Кредитные карты"
        SheetName_String = FindNameSheet(ReportName_String, StringInSheet) ' "3.2 Кредитные карты"
        If SheetName_String <> "" Then
                
          ном_стр_офис = ном_стр_офис + 1
          Call DB_UniversalSheetInDB(ReportName_String, _
                                     SheetName_String, _
                                       officeNameInReport, _
                                         (row_ОО_Тюменский + 2 + ном_стр_офис) + ((i - 1) * Размер_блока_офиса), _
                                           НумерацияПунктов("Сплиты к ПК"), _
                                             "Сплиты к ПК", _
                                               "ПК_КК", _
                                                 "шт.", _
                                                    0, _
                                                     "Сплиты к ПК_Месяц", _
                                                       "Сплиты к ПК_Квартал", _
                                                         0, _
                                                           "Алтайский ОО1", _
                                                             0, _
                                                               0, _
                                                                 0, _
                                                                   0, 1, 1)
        
        Else
          ' Если в DB Лист не найден
          Call в_DB_Лист_не_найден(StringInSheet)
        End If

        ' Показатель "Сплиты к ПК" (КК) - активация
        StringInSheet = "Кредитные карты"
        SheetName_String = FindNameSheet(ReportName_String, StringInSheet) ' "3.2 Кредитные карты"
        If SheetName_String <> "" Then
                
          ном_стр_офис = ном_стр_офис + 1
          Call DB_UniversalSheetInDB(ReportName_String, _
                                     SheetName_String, _
                                       officeNameInReport, _
                                         (row_ОО_Тюменский + 2 + ном_стр_офис) + ((i - 1) * Размер_блока_офиса), _
                                           НумерацияПунктов("Сплиты к ПК (актив.)"), _
                                             "Сплиты к ПК (актив.)", _
                                               "ПК_КК_актив", _
                                                 "шт.", _
                                                    0, _
                                                     "Сплиты к ПК_Месяц", _
                                                       "Сплиты к ПК_Квартал", _
                                                        0, _
                                                          "Алтайский ОО1", _
                                                             1, _
                                                               1, _
                                                                 0, _
                                                                   0, 2, 2)
        
        Else
          ' Если в DB Лист не найден
          Call в_DB_Лист_не_найден(StringInSheet)
        End If


        ' Показатель "Заявки КК"
        StringInSheet = "Кредитные карты"
        SheetName_String = FindNameSheet(ReportName_String, StringInSheet) ' "3.2 Кредитные карты"
        If SheetName_String <> "" Then
                
          ном_стр_офис = ном_стр_офис + 1
          Call DB_UniversalSheetInDB(ReportName_String, _
                                     SheetName_String, _
                                       officeNameInReport, _
                                         (row_ОО_Тюменский + 2 + ном_стр_офис) + ((i - 1) * Размер_блока_офиса), _
                                           НумерацияПунктов("Заявки на Кредитные карты"), _
                                             "Заявки на Кредитные карты", _
                                               "Заявки_КК", _
                                                 "шт.", _
                                                    0, _
                                                     "Заявки_Месяц", _
                                                       "Заявки_Квартал", _
                                                         4, _
                                                          "Алтайский ОО1", _
                                                             0, _
                                                               0, _
                                                                 0, _
                                                                   0, 1, 1)
        
        Else
          ' Если в DB Лист не найден
          Call в_DB_Лист_не_найден(StringInSheet)
        End If

        ' Показатель "КК Кол-во карт в сейфе "
        StringInSheet = "Кредитные карты"
        SheetName_String = FindNameSheet(ReportName_String, StringInSheet) '
        If SheetName_String <> "" Then

          ном_стр_офис = ном_стр_офис + 1
          Call DB_getParamFromUniversalSheetInDB(ReportName_String, _
                                                   SheetName_String, _
                                                     "Алтайский ОО1", _
                                                       officeNameInReport, _
                                                         (row_ОО_Тюменский + 2 + ном_стр_офис) + ((i - 1) * Размер_блока_офиса), _
                                                           НумерацияПунктов("КК кол-во карт в сейфе"), _
                                                             "КК кол-во карт в сейфе", _
                                                               "Кол-во карт в сейфе ", _
                                                                 "КК_Сейф_Всего", _
                                                                   "шт.", _
                                                                     0, _
                                                                      "Месяц")
            
        Else
          ' Если в DB Лист не найден
          Call в_DB_Лист_не_найден(StringInSheet)
        End If

        ' Показатель КК Кол-во выданных карт за месяц (ном_стр_офис=40) & ' Показатель №28 КК Кол-во выданных карт за квартал (ном_стр_офис=40) - надо ли?
        StringInSheet = "Кредитные карты"
        SheetName_String = FindNameSheet(ReportName_String, StringInSheet) '
        If SheetName_String <> "" Then

          ном_стр_офис = ном_стр_офис + 1
          Call DB_getParamFromUniversalSheetInDB(ReportName_String, _
                                                   SheetName_String, _
                                                     "Алтайский ОО1", _
                                                       officeNameInReport, _
                                                         (row_ОО_Тюменский + 2 + ном_стр_офис) + ((i - 1) * Размер_блока_офиса), _
                                                           НумерацияПунктов("КК кол-во выданных"), _
                                                             "КК кол-во выданных", _
                                                               "Кол-во выданных карт за месяц", _
                                                                 "КК_Выдано_Месяц", _
                                                                   "шт.", _
                                                                     0, _
                                                                      "Месяц")
        Else
          ' Если в DB Лист не найден
          Call в_DB_Лист_не_найден(StringInSheet)
        End If


        ' ----------------------------------------------------------------------------------------------------------------------------------
        ' Чертим горизонтальную линию 2 (указываем предидущее значение строки + 1)
        Call gorizontalLineII(ThisWorkbook.Name, "Лист8", (row_ОО_Тюменский + 2 + ном_стр_офис) + ((i - 1) * Размер_блока_офиса) + 1, 2, 12)
        ' ----------------------------------------------------------------------------------------------------------------------------------

            
        ' Показатель ДК
        StringInSheet = "Дебетовые карты"
        SheetName_String = FindNameSheet(ReportName_String, StringInSheet) ' "3.3 Дебетовые карты"
        If SheetName_String <> "" Then
        
          ном_стр_офис = ном_стр_офис + 1
          ' Call DB_UniversalSheetInDB(ReportName_String, _
          '                            SheetName_String, _
          '                              officeNameInReport, _
          '                                (row_ОО_Тюменский + 2 + ном_стр_офис) + ((i - 1) * Размер_блока_офиса), _
          '                                  НумерацияПунктов("Дебетовые карты (актив.)"), _
          '                                    "Дебетовые карты (актив.)", _
          '                                      "ДК", _
          '                                        "шт.", _
          '                                          0, _
          '                                            "Активированные карты, шт._Месяц", _
          '                                              "Активированные карты, шт._Квартал", _
          '                                                4, _
          '                                                  "Алтайский ОО1", _
          '                                                    0, _
          '                                                      0, _
          '                                                        0, _
          '                                                          0, 1, 1)
        
          ' Call DB_UniversalSheetInDB(ReportName_String, _
          '                           SheetName_String, _
          '                             officeNameInReport, _
          '                               (row_ОО_Тюменский + 2 + ном_стр_офис) + ((i - 1) * Размер_блока_офиса), _
          '                                 НумерацияПунктов("Дебетовые карты (актив.)"), _
          '                                   "Дебетовые карты (актив.)", _
          '                                     "ДК", _
          '                                       "шт.", _
          '                                         0, _
          '                                           "Месяц", _
          '                                             "Квартал", _
          '                                               4, _
          '                                                 "Алтайский ОО1", _
          '                                                   5, _
          '                                                     5, _
          '                                                       0, _
          '                                                         0, 1, 1)
        
          Call DB_UniversalSheetInDB(ReportName_String, _
                                     SheetName_String, _
                                       officeNameInReport, _
                                         (row_ОО_Тюменский + 2 + ном_стр_офис) + ((i - 1) * Размер_блока_офиса), _
                                           НумерацияПунктов("Дебетовые карты (актив.)"), _
                                             "Дебетовые карты (актив.)", _
                                               "ДК", _
                                                 "шт.", _
                                                   0, _
                                                     "Активированные карты, шт._Месяц  ", _
                                                       "Активированные карты, шт._Квартал ", _
                                                         4, _
                                                           "Алтайский ОО1", _
                                                             15, _
                                                               15, _
                                                                 0, _
                                                                   0, 1, 1)
        
    
        Else
          ' Если в DB Лист не найден
          Call в_DB_Лист_не_найден(StringInSheet)
        End If

        ' Показатель "ИЗП индивидуальный зарплатный проект"
        StringInSheet = "ИЗП"
        SheetName_String = FindNameSheet(ReportName_String, StringInSheet) ' 8.13 ИЗП
        If SheetName_String <> "" Then
                
          ном_стр_офис = ном_стр_офис + 1

          ' Первый вариант - через DB_getParamFromUniversalSheetInDB - работает! Sub DB_getParamFromUniversalSheetInDB(In_ReportName_String, In_Sheets, In_Заголовок_столбца_офисы, In_officeNameInReport, In_Row_Лист8, In_N, In_Product_Name, In_Param_Name_In_DB, In_Product_Code, In_Unit, In_Weight, In_Period)
          ' Call DB_getParamFromUniversalSheetInDB(ReportName_String, _
          '                                          SheetName_String, _
          '                                            "Алтайский ОО1", _
          '                                              officeNameInReport, _
          '                                                (row_ОО_Тюменский + 2 + ном_стр_офис) + ((i - 1) * Размер_блока_офиса), _
          '                                                  НумерацияПунктов("ИЗП"), _
          '                                                    "ИЗП", _
          '                                                      "Подключено ИЗП, шт. (за вычетом пенсионных карт)", _
          '                                                        "ИЗП", _
          '                                                          "шт.", _
          '                                                            0, _
          '                                                             "Квартал")
        

          Call DB_UniversalSheetInDB(ReportName_String, _
                                     SheetName_String, _
                                       officeNameInReport, _
                                         (row_ОО_Тюменский + 2 + ном_стр_офис) + ((i - 1) * Размер_блока_офиса), _
                                           НумерацияПунктов("ИЗП"), _
                                             "ИЗП", _
                                               "ИЗП", _
                                                 "шт.", _
                                                    0, _
                                                     "", _
                                                       "Выдано ДК, шт. (за вычетом пенсионных карт)", _
                                                         0, _
                                                           "Алтайский ОО1", _
                                                             0, _
                                                               0, _
                                                                 0, _
                                                                   0, 1, 1)


        
        Else
          ' Если в DB Лист не найден
          Call в_DB_Лист_не_найден(StringInSheet)
        End If


        ' Показатель "Заявки на Дебетовые карты"
        StringInSheet = "Дебетовые карты"
        SheetName_String = FindNameSheet(ReportName_String, StringInSheet) ' "3.3 Дебетовые карты"
        If SheetName_String <> "" Then
                
          ном_стр_офис = ном_стр_офис + 1
          Call DB_UniversalSheetInDB(ReportName_String, _
                                     SheetName_String, _
                                       officeNameInReport, _
                                         (row_ОО_Тюменский + 2 + ном_стр_офис) + ((i - 1) * Размер_блока_офиса), _
                                           НумерацияПунктов("Заявки на Дебетовые карты"), _
                                             "Заявки на Дебетовые карты", _
                                               "Заявки_ДК", _
                                                 "шт.", _
                                                    0, _
                                                     "Заявки_Месяц", _
                                                       "Заявки_Квартал", _
                                                         0, _
                                                          "Алтайский ОО1", _
                                                             0, _
                                                               0, _
                                                                 0, _
                                                                   0, 1, 1)
        
        Else
          ' Если в DB Лист не найден
          Call в_DB_Лист_не_найден(StringInSheet)
        End If




        ' Показатель "ДК Кол-во карт в сейфе "
        StringInSheet = "Дебетовые карты"
        SheetName_String = FindNameSheet(ReportName_String, StringInSheet) '
        If SheetName_String <> "" Then

          ном_стр_офис = ном_стр_офис + 1
          Call DB_getParamFromUniversalSheetInDB(ReportName_String, _
                                                   SheetName_String, _
                                                     "Алтайский ОО1", _
                                                       officeNameInReport, _
                                                         (row_ОО_Тюменский + 2 + ном_стр_офис) + ((i - 1) * Размер_блока_офиса), _
                                                           НумерацияПунктов("ДК кол-во карт в сейфе"), _
                                                             "ДК кол-во карт в сейфе", _
                                                               "Кол-во карт в сейфе ", _
                                                                 "ДК_Сейф_Всего", _
                                                                   "шт.", _
                                                                     0, _
                                                                      "Месяц")
            
        Else
          ' Если в DB Лист не найден
          Call в_DB_Лист_не_найден(StringInSheet)
        End If

        ' Показатель "ДК Кол-во выданных карт за месяц"  & ' Показатель "ДК Кол-во выданных карт за квартал" - надо ли?
        StringInSheet = "Дебетовые карты"
        SheetName_String = FindNameSheet(ReportName_String, StringInSheet) '
        If SheetName_String <> "" Then

          ном_стр_офис = ном_стр_офис + 1
          Call DB_getParamFromUniversalSheetInDB(ReportName_String, _
                                                   SheetName_String, _
                                                     "Алтайский ОО1", _
                                                       officeNameInReport, _
                                                         (row_ОО_Тюменский + 2 + ном_стр_офис) + ((i - 1) * Размер_блока_офиса), _
                                                           НумерацияПунктов("ДК кол-во выданных"), _
                                                             "ДК кол-во выданных", _
                                                               "Кол-во выданных карт за месяц", _
                                                                 "ДК_Выдано_Месяц", _
                                                                   "шт.", _
                                                                     0, _
                                                                      "Месяц")
            
        Else
          ' Если в DB Лист не найден
          Call в_DB_Лист_не_найден(StringInSheet)
        End If


        

        ' ----------------------------------------------------------------------------------------------------------------------------------
        ' Чертим горизонтальную линию 2 (указываем предидущее значение строки + 1)
        Call gorizontalLineII(ThisWorkbook.Name, "Лист8", (row_ОО_Тюменский + 2 + ном_стр_офис) + ((i - 1) * Размер_блока_офиса) + 1, 2, 12)
        ' ----------------------------------------------------------------------------------------------------------------------------------

        
        ' Показатель №5 (6) ИБ (ном_стр_офис=6)
        StringInSheet = "ИБ"
        SheetName_String = FindNameSheet(ReportName_String, StringInSheet) ' "3.10 ИБ"
        If SheetName_String <> "" Then

          ном_стр_офис = ном_стр_офис + 1
          Call DB_UniversalSheetInDB(ReportName_String, _
                                     SheetName_String, _
                                       officeNameInReport, _
                                         (row_ОО_Тюменский + 2 + ном_стр_офис) + ((i - 1) * Размер_блока_офиса), _
                                           НумерацияПунктов("Интернет-банк"), _
                                             "Интернет-банк", _
                                               "ИБ", _
                                                 "шт.", _
                                                   0, _
                                                     "ИБ_Месяц", _
                                                       "ИБ_Квартал", _
                                                         4, _
                                                           "Филиал", _
                                                             0, _
                                                               0, _
                                                                 0, _
                                                                   0, 1, 1)
        
        Else
          ' Если в DB Лист не найден
          Call в_DB_Лист_не_найден(StringInSheet)
        End If


            
        ' Показатель НС
        StringInSheet = "Накопительные счета"
        SheetName_String = FindNameSheet(ReportName_String, StringInSheet) ' "3.5 Накопительные счета"
        If SheetName_String <> "" Then
        
          ном_стр_офис = ном_стр_офис + 1
          Call DB_UniversalSheetInDB(ReportName_String, _
                                      SheetName_String, _
                                       officeNameInReport, _
                                         (row_ОО_Тюменский + 2 + ном_стр_офис) + ((i - 1) * Размер_блока_офиса), _
                                           НумерацияПунктов("Накопительные счета"), _
                                             "Накопительные счета", _
                                               "НС", _
                                                 "шт.", _
                                                   0, _
                                                     "Накопительные счета, открытые в офисе, шт._Месяц", _
                                                       "Накопительные счета, открытые в офисе, шт._Квартал", _
                                                         4, _
                                                           "Алтайский ОО1", _
                                                             0, _
                                                               0, _
                                                                 0, _
                                                                   0, 1, 1)
                                                                   
        Else
          ' Если в DB Лист не найден
          Call в_DB_Лист_не_найден(StringInSheet)
        End If


        ' ----------------------------------------------------------------------------------------------------------------------------------
        ' Чертим горизонтальную линию 2 (указываем предидущее значение строки + 1)
        Call gorizontalLineII(ThisWorkbook.Name, "Лист8", (row_ОО_Тюменский + 2 + ном_стр_офис) + ((i - 1) * Размер_блока_офиса) + 1, 2, 12)
        ' ----------------------------------------------------------------------------------------------------------------------------------


        ' Показатель №7 (8) OPC (ном_стр_офис=8)
        StringInSheet = "OPC"
        SheetName_String = FindNameSheet(ReportName_String, StringInSheet) ' "3.9 OPC"
        If SheetName_String <> "" Then
        
          ном_стр_офис = ном_стр_офис + 1
          Call DB_UniversalSheetInDB(ReportName_String, _
                                     SheetName_String, _
                                       officeNameInReport, _
                                         (row_ОО_Тюменский + 2 + ном_стр_офис) + ((i - 1) * Размер_блока_офиса), _
                                           НумерацияПунктов("Orange Premium Club"), _
                                             "Orange Premium Club", _
                                               "OPC", _
                                                 "шт.", _
                                                   0, _
                                                     "OPC, шт._Месяц", _
                                                       "OPC, шт._Квартал", _
                                                         4, _
                                                           "Алтайский ОО1", _
                                                             0, _
                                                               0, _
                                                                 0, _
                                                                   0, 1, 1)
          ном_стр_офис = ном_стр_офис + 1
          Call DB_getParamFromUniversalSheetInDB2(ReportName_String, _
                                                   SheetName_String, _
                                                     "Алтайский ОО1", _
                                                       officeNameInReport, _
                                                         (row_ОО_Тюменский + 2 + ном_стр_офис) + ((i - 1) * Размер_блока_офиса), _
                                                           НумерацияПунктов("OPC_Портфель"), _
                                                             "Портфель OPC", _
                                                               "OPC, шт._Квартал", _
                                                                 "OPC_Портфель", _
                                                                   "шт.", _
                                                                     0, _
                                                                      "Квартал", _
                                                                        6)

          
        Else
          ' Если в DB Лист не найден
          Call в_DB_Лист_не_найден(StringInSheet)
        End If


        ' ----------------------------------------------------------------------------------------------------------------------------------
        ' Чертим горизонтальную линию 2 (указываем предидущее значение строки + 1)
        Call gorizontalLineII(ThisWorkbook.Name, "Лист8", (row_ОО_Тюменский + 2 + ном_стр_офис) + ((i - 1) * Размер_блока_офиса) + 1, 2, 12)
        ' ----------------------------------------------------------------------------------------------------------------------------------
       
                                                               
        
        ' Показатель №12 (13) Срочные вклады (ном_стр_офис=13) - заменить на Ипотеку с "Лист3"
        ' StringInSheet = "Срочные вклады"
        ' SheetName_String = FindNameSheet(ReportName_String, StringInSheet) ' "3.4 Срочные вклады"
        ' If SheetName_String <> "" Then
        
        '   ном_стр_офис = ном_стр_офис + 1
        '   Call DB_UniversalSheetInDB(ReportName_String, _
        '                              SheetName_String, _
        '                                officeNameInReport, _
        '                                  (row_ОО_Тюменский + 2 + ном_стр_офис) + ((i - 1) * Размер_блока_офиса), _
        '                                    НумерацияПунктов("Срочные вклады"), _
        '                                      "Срочные вклады", _
        '                                        "ВКЛ", _
        '                                          "тыс.руб.", _
        '                                            0, _
        '                                              "Портфель, тыс.руб._Месяц", _
        '                                                "Портфель, тыс.руб._Квартал", _
        '                                                  0, _
        '                                                    "Алтайский ОО1", _
        '                                                      0, _
        '                                                        0, _
        '                                                          0, _
        '                                                            0, 1, 1)
        '
        '
        ' Else
        '   ' Если в DB Лист не найден
        '   Call в_DB_Лист_не_найден(StringInSheet)
        ' End If
                                                                   

        ' ----------------------------------------------------------------------------------------------------------------------------------
        ' Чертим горизонтальную линию 2 (указываем предидущее значение строки + 1)
        ' Call gorizontalLineII(ThisWorkbook.Name, "Лист8", (row_ОО_Тюменский + 2 + ном_стр_офис) + ((i - 1) * Размер_блока_офиса) + 1, 2, 12)
        ' ----------------------------------------------------------------------------------------------------------------------------------


        ' Показатель №13 (14) Ядро (3.10 ИБ) (ном_стр_офис=14)
        StringInSheet = "ИБ"
        SheetName_String = FindNameSheet(ReportName_String, StringInSheet) ' "3.10 ИБ"
        If SheetName_String <> "" Then
        
          ном_стр_офис = ном_стр_офис + 1
          Call DB_UniversalSheetInDB(ReportName_String, _
                                     SheetName_String, _
                                       officeNameInReport, _
                                         (row_ОО_Тюменский + 2 + ном_стр_офис) + ((i - 1) * Размер_блока_офиса), _
                                           НумерацияПунктов("Ядро клиентов"), _
                                             "Ядро клиентов", _
                                               "ЯД", _
                                                 "шт.", _
                                                   0, _
                                                     "Данные за Месяц", _
                                                       "Данные за квартал", _
                                                         0, _
                                                           "Филиал", _
                                                             0, _
                                                               0, _
                                                                 0, _
                                                                   0, 1, 1)
        Else
          ' Если в DB Лист не найден
          Call в_DB_Лист_не_найден(StringInSheet)
        End If
                                                                   

        ' ----------------------------------------------------------------------------------------------------------------------------------
        ' Чертим горизонтальную линию 2 (указываем предидущее значение строки + 1)
        Call gorizontalLineII(ThisWorkbook.Name, "Лист8", (row_ОО_Тюменский + 2 + ном_стр_офис) + ((i - 1) * Размер_блока_офиса) + 1, 2, 12)
        ' ----------------------------------------------------------------------------------------------------------------------------------

        ' Показатель №14 (15) Комиссионный доход
        StringInSheet = "Ком доход"
        SheetName_String = FindNameSheet(ReportName_String, StringInSheet) ' "2. Ком доход"
        If SheetName_String <> "" Then
          
                    
          
          ' Если на листе "Ком доход" найдена "Страховка к кредитам", то это старая версия отчета
          If ColumnByValue(ReportName_String, SheetName_String, "Месяц", 1000, 1000) = 0 Then
            ном_стр_офис = ном_стр_офис + 1 ' (ном_стр_офис=15)
            ' (старая версия DB 2019 года)
            MsgBox ("Внимание! Старая версия листа Ком доход!")
            ' Call DB_UniversalSheetInDB(ReportName_String, SheetName_String, officeNameInReport, (row_ОО_Тюменский + 2 + ном_стр_офис) + ((i - 1) * Размер_блока_офиса), 14, "Комиссионный доход", "КД", "тыс.руб.", 0.2, "", "Ком.доход, тыс. руб. План квартал", 0, "Алтайский ОО1", 0, 0, 0, 0)
            ' Call DB_UniversalSheetInDB(ReportName_String, SheetName_String, officeNameInReport, (row_ОО_Тюменский + 2 + ном_стр_офис) + ((i - 1) * Размер_блока_офиса), 14, "Комиссионный доход", "КД", "тыс.руб.", 0.2, "", "Итог План квартал", 0, "Алтайский ОО1", 0, 0, 0, 0)
            Call DB_UniversalSheetInDB(ReportName_String, SheetName_String, officeNameInReport, (row_ОО_Тюменский + 2 + ном_стр_офис) + ((i - 1) * Размер_блока_офиса), 14, "Комиссионный доход", "КД", "тыс.руб.", 0.2, "", "Итого", 0, "Алтайский ОО1", 0, 0, 0, 0, 1, 1)
          End If ' Если на листе "Ком доход" найдена "Страховка к кредитам", то это старая версия отчета
          
          
          ' Если на листе "Ком доход" найдена "Страховки к ПК", то это новая версия отчета (ном_стр_офис=15)
          If ColumnByValue(ReportName_String, SheetName_String, "Месяц", 1000, 1000) <> 0 Then
          
            ' (новая версия DB 2020 года)
            ном_стр_офис = ном_стр_офис + 1
            Call DB_UniversalSheetInDB(ReportName_String, _
                                     SheetName_String, _
                                       officeNameInReport, _
                                         (row_ОО_Тюменский + 2 + ном_стр_офис) + ((i - 1) * Размер_блока_офиса), _
                                           НумерацияПунктов("Комиссионный доход"), _
                                             "Комиссионный доход", _
                                               "КД", _
                                                 "тыс.руб.", _
                                                   0.15, _
                                                     "Месяц", _
                                                       "Квартал", _
                                                         3, _
                                                           "Алтайский ОО1", _
                                                             39, _
                                                               39, _
                                                                 0, _
                                                                   0, 1, 1)
        
        
            ' Показатель №14.1 (16) Комиссионный доход - Страховки к ПК "2. Ком доход" (новая версия DB 2020 года) (ном_стр_офис=16)
            ном_стр_офис = ном_стр_офис + 1
            Call DB_UniversalSheetInDB(ReportName_String, _
                                     SheetName_String, _
                                       officeNameInReport, _
                                         (row_ОО_Тюменский + 2 + ном_стр_офис) + ((i - 1) * Размер_блока_офиса), _
                                           НумерацияПунктов("в т.ч. страховки к ПК"), _
                                             "в т.ч. страховки к ПК", _
                                               "КД_БС", _
                                                 "тыс.руб.", _
                                                   0, _
                                                     "Месяц", _
                                                       "Квартал", _
                                                         3, _
                                                           "Алтайский ОО1", _
                                                             0, _
                                                               0, _
                                                                 0, _
                                                                   0, 1, 1)
        
           ' Показатель №14.2 (17) Комиссионный доход - ИСЖ_МАСС "2. Ком доход" (новая версия DB 2020 года) (ном_стр_офис=17)
           ном_стр_офис = ном_стр_офис + 1
           Call DB_UniversalSheetInDB(ReportName_String, _
                                     SheetName_String, _
                                       officeNameInReport, _
                                         (row_ОО_Тюменский + 2 + ном_стр_офис) + ((i - 1) * Размер_блока_офиса), _
                                           НумерацияПунктов("           ИСЖ_МАСС"), _
                                             "           ИСЖ_МАСС", _
                                               "КД_ИСЖ_МАСС", _
                                                 "тыс.руб.", _
                                                   0, _
                                                     "Месяц", _
                                                       "Квартал", _
                                                         3, _
                                                           "Алтайский ОО1", _
                                                             4, _
                                                               4, _
                                                                 0, _
                                                                   0, 1, 1)
        
            
            ' Показатель №14.3 (18) Комиссионный доход - НСЖ_МАСС "2. Ком доход" (новая версия DB 2020 года) (ном_стр_офис=18)
            ном_стр_офис = ном_стр_офис + 1
            Call DB_UniversalSheetInDB(ReportName_String, _
                                     SheetName_String, _
                                       officeNameInReport, _
                                         (row_ОО_Тюменский + 2 + ном_стр_офис) + ((i - 1) * Размер_блока_офиса), _
                                           НумерацияПунктов("           НСЖ_МАСС"), _
                                             "           НСЖ_МАСС", _
                                               "КД_ИСЖ_НАСС", _
                                                 "тыс.руб.", _
                                                   0, _
                                                     "Месяц", _
                                                       "Квартал", _
                                                         3, _
                                                           "Алтайский ОО1", _
                                                             8, _
                                                               8, _
                                                                 0, _
                                                                   0, 1, 1)

        
          ' Показатель №14.4 (19) Комиссионный доход - Коробочное страхование "2. Ком доход" (новая версия DB 2020 года) (ном_стр_офис=19)
          ном_стр_офис = ном_стр_офис + 1
          Call DB_UniversalSheetInDB(ReportName_String, _
                                     SheetName_String, _
                                       officeNameInReport, _
                                         (row_ОО_Тюменский + 2 + ном_стр_офис) + ((i - 1) * Размер_блока_офиса), _
                                           НумерацияПунктов("           КС"), _
                                             "           КС", _
                                               "КД_КС", _
                                                 "тыс.руб.", _
                                                   0, _
                                                     "Месяц", _
                                                       "Квартал", _
                                                         9, _
                                                           "Алтайский ОО1", _
                                                             12, _
                                                               12, _
                                                                 0, _
                                                                   0, 1, 1)
          
     
          
          ' C DB 16.08.2021 его нет
          ' Показатель №14.5 (20) Комиссионный доход - Личный адвокат "2. Ком доход" (новая версия DB 2020 года) (ном_стр_офис=20)
          ' ном_стр_офис = ном_стр_офис + 1
          ' Call DB_UniversalSheetInDB(ReportName_String, _
          '                            SheetName_String, _
          '                              officeNameInReport, _
          '                               (row_ОО_Тюменский + 2 + ном_стр_офис) + ((i - 1) * Размер_блока_офиса), _
          '                                 НумерацияПунктов("           ЛА"), _
          '                                   "           ЛА", _
          '                                     "КД_ЛА", _
          '                                       "тыс.руб.", _
          '                                         0, _
          '                                           "Месяц", _
          '                                             "Квартал", _
          '                                               3, _
          '                                                 "Алтайский ОО1", _
          '                                                   16, _
          '                                                     16, _
          '                                                       0, _
          '                                                         0, 1, 1)
        
          ' Показатель №14.5 (20+) Комиссионный доход - УК (MASS)
          ном_стр_офис = ном_стр_офис + 1
          Call DB_UniversalSheetInDB(ReportName_String, _
                                     SheetName_String, _
                                       officeNameInReport, _
                                         (row_ОО_Тюменский + 2 + ном_стр_офис) + ((i - 1) * Размер_блока_офиса), _
                                           НумерацияПунктов("           УК (MASS)"), _
                                             "           УК (MASS)", _
                                               "КД_УК_MASS", _
                                                 "тыс.руб.", _
                                                   0, _
                                                     "Месяц", _
                                                       "Квартал", _
                                                         0, _
                                                           "Алтайский ОО1", _
                                                             26, _
                                                               26, _
                                                                 0, _
                                                                   0, 1, 1)
        
           
            ' Показатель №14.6 (21) Комиссионный доход - ИСЖ, НСЖ, КС (Affluent) "2. Ком доход" (новая версия DB 2020 года) (ном_стр_офис=21)
            ном_стр_офис = ном_стр_офис + 1
            Call DB_UniversalSheetInDB(ReportName_String, _
                                     SheetName_String, _
                                       officeNameInReport, _
                                         (row_ОО_Тюменский + 2 + ном_стр_офис) + ((i - 1) * Размер_блока_офиса), _
                                           НумерацияПунктов("           ИСЖ, НСЖ, КС (Affluent)"), _
                                             "           ИСЖ, НСЖ, КС (Affluent)", _
                                               "КД_ИСЖ_НСЖ_КС_Affluent", _
                                                 "тыс.руб.", _
                                                   0, _
                                                     "Месяц", _
                                                       "Квартал", _
                                                         3, _
                                                           "Алтайский ОО1", _
                                                             22, _
                                                               22, _
                                                                 0, _
                                                                   0, 1, 1)
        

            ' Показатель №14.7 (22) Комиссионный доход - УК (Affluent) "2. Ком доход" (новая версия DB 2020 года) (ном_стр_офис=22)
            ном_стр_офис = ном_стр_офис + 1
            Call DB_UniversalSheetInDB(ReportName_String, _
                                     SheetName_String, _
                                       officeNameInReport, _
                                         (row_ОО_Тюменский + 2 + ном_стр_офис) + ((i - 1) * Размер_блока_офиса), _
                                           НумерацияПунктов("           УК (Affluent)"), _
                                             "           УК (Affluent)", _
                                               "КД_УК_Affluent", _
                                                 "тыс.руб.", _
                                                   0, _
                                                     "Месяц", _
                                                       "Квартал", _
                                                         0, _
                                                           "Алтайский ОО1", _
                                                             29, _
                                                               29, _
                                                                 0, _
                                                                   0, 1, 1)
                                                                                                                              
          End If ' Если это новая версия комдохода на листе
          
        Else
          ' Если в DB Лист не найден
          Call в_DB_Лист_не_найден(StringInSheet)
        End If ' "2. Ком доход"
                                                                   
        ' ----------------------------------------------------------------------------------------------------------------------------------
        ' Чертим горизонтальную линию 2 (указываем предидущее значение строки + 1)
        Call gorizontalLineII(ThisWorkbook.Name, "Лист8", (row_ОО_Тюменский + 2 + ном_стр_офис) + ((i - 1) * Размер_блока_офиса) + 1, 2, 12)
        ' ----------------------------------------------------------------------------------------------------------------------------------
                                                                   
                                                                   
        ' Показатель "Коробки+ЛА" в штуках
        StringInSheet = "Коробки+Юрист24" ' "Коробки+ЛА"
        SheetName_String = FindNameSheet(ReportName_String, StringInSheet) ' с 08-07-2021 "8.9 Коробки+Юрист24", "3.8 Коробки+ЛА"
        If SheetName_String <> "" Then

          ном_стр_офис = ном_стр_офис + 1
          Call DB_UniversalSheetInDB(ReportName_String, _
                                    SheetName_String, _
                                      officeNameInReport, _
                                        (row_ОО_Тюменский + 2 + ном_стр_офис) + ((i - 1) * Размер_блока_офиса), _
                                          НумерацияПунктов("Коробки+Личный адвокат"), _
                                            "Коробки+Личный адвокат", _
                                               "КЛА", _
                                                 "шт.", _
                                                   0, _
                                                     "Месяц", _
                                                       "Квартал", _
                                                         3, _
                                                           "Алтайский ОО1", _
                                                             0, _
                                                               0, _
                                                                 0, _
                                                                   0, 1, 1)
            
            
        Else
          ' Если в DB Лист не найден
          Call в_DB_Лист_не_найден(StringInSheet)
        End If

        ' Показатель "Коробки+ЛА" (премия)
        StringInSheet = "Коробки+Юрист24" ' "Коробки+ЛА"
        SheetName_String = FindNameSheet(ReportName_String, StringInSheet) ' "3.8 Коробки+ЛА"
        If SheetName_String <> "" Then

          ном_стр_офис = ном_стр_офис + 1
          Call DB_UniversalSheetInDB(ReportName_String, _
                                     SheetName_String, _
                                       officeNameInReport, _
                                         (row_ОО_Тюменский + 2 + ном_стр_офис) + ((i - 1) * Размер_блока_офиса), _
                                           НумерацияПунктов("Коробки+Личный адвокат (премия)"), _
                                             "Коробки+Личный адвокат (премия)", _
                                               "КЛА_Премия", _
                                                 "тыс.руб.", _
                                                   0, _
                                                     "Месяц", _
                                                       "Квартал", _
                                                         4, _
                                                           "Алтайский ОО1", _
                                                             47, _
                                                               47, _
                                                                 0, _
                                                                   0, 1, 1)
            
            
          ' "в т.ч. КЛА_Премия_Офис"
          ном_стр_офис = ном_стр_офис + 1
          Call DB_UniversalSheetInDB(ReportName_String, _
                                     SheetName_String, _
                                       officeNameInReport, _
                                         (row_ОО_Тюменский + 2 + ном_стр_офис) + ((i - 1) * Размер_блока_офиса), _
                                           НумерацияПунктов("в т.ч. КЛА Премия Офис"), _
                                             "в т.ч. КЛА премия Офис", _
                                               "КЛА_Премия_Офис", _
                                                 "тыс.руб.", _
                                                   0, _
                                                     "Месяц", _
                                                       "Квартал", _
                                                         4, _
                                                           "Алтайский ОО1", _
                                                             21, _
                                                               21, _
                                                                 0, _
                                                                   0, 1, 1)
          ' "           КЛА Премия ИЦ"
          ном_стр_офис = ном_стр_офис + 1
          Call DB_UniversalSheetInDB(ReportName_String, _
                                     SheetName_String, _
                                       officeNameInReport, _
                                         (row_ОО_Тюменский + 2 + ном_стр_офис) + ((i - 1) * Размер_блока_офиса), _
                                           НумерацияПунктов("           КЛА премия ИЦ"), _
                                             "           КЛА премия ИЦ", _
                                               "КЛА_Премия_ИЦ", _
                                                 "тыс.руб.", _
                                                   0, _
                                                     "Месяц", _
                                                       "Квартал", _
                                                         4, _
                                                           "Алтайский ОО1", _
                                                             26, _
                                                               26, _
                                                                 0, _
                                                                   0, 1, 1)
          
          ' "           КЛА Премия OPC"
          ном_стр_офис = ном_стр_офис + 1
          Call DB_UniversalSheetInDB(ReportName_String, _
                                     SheetName_String, _
                                       officeNameInReport, _
                                         (row_ОО_Тюменский + 2 + ном_стр_офис) + ((i - 1) * Размер_блока_офиса), _
                                           НумерацияПунктов("           КЛА премия OPC"), _
                                             "           КЛА премия OPC", _
                                               "КЛА_Премия_OPC", _
                                                 "тыс.руб.", _
                                                   0, _
                                                     "Месяц", _
                                                       "Квартал", _
                                                         4, _
                                                           "Алтайский ОО1", _
                                                             31, _
                                                               31, _
                                                                 0, _
                                                                   0, 1, 1)
            
            
        Else
          ' Если в DB Лист не найден
          Call в_DB_Лист_не_найден(StringInSheet)
        End If


        ' ----------------------------------------------------------------------------------------------------------------------------------
        ' Чертим горизонтальную линию 2 (указываем предидущее значение строки + 1)
        Call gorizontalLineII(ThisWorkbook.Name, "Лист8", (row_ОО_Тюменский + 2 + ном_стр_офис) + ((i - 1) * Размер_блока_офиса) + 1, 2, 12)
        ' ----------------------------------------------------------------------------------------------------------------------------------


        ' Показатель "ИСЖ_МАСС"
        StringInSheet = "ИСЖ_МАСС"
        SheetName_String = FindNameSheet(ReportName_String, StringInSheet) ' "3.6 ИСЖ_МАСС"
        If SheetName_String <> "" Then
        
          ном_стр_офис = ном_стр_офис + 1
          Call DB_UniversalSheetInDB(ReportName_String, _
                                     SheetName_String, _
                                       officeNameInReport, _
                                         (row_ОО_Тюменский + 2 + ном_стр_офис) + ((i - 1) * Размер_блока_офиса), _
                                           НумерацияПунктов("Премия ИСЖ МАСС"), _
                                             "Премия ИСЖ МАСС", _
                                               "ИСЖ_МАСС", _
                                                 "тыс.руб.", _
                                                   0, _
                                                     "Премия, тыс.руб._Месяц", _
                                                       "Премия, тыс.руб._Квартал", _
                                                        3, _
                                                          "Алтайский ОО1", _
                                                             0, _
                                                               0, _
                                                                 0, _
                                                                   0, 1, 1)
        
        Else
          ' Если в DB Лист не найден
          Call в_DB_Лист_не_найден(StringInSheet)
        End If
                                                                   
        
        ' Показатель "НСЖ_МАСС"
        StringInSheet = "НСЖ_МАСС"
        SheetName_String = FindNameSheet(ReportName_String, StringInSheet) ' "3.7 НСЖ_МАСС"
        If SheetName_String <> "" Then
        
          ном_стр_офис = ном_стр_офис + 1
          Call DB_UniversalSheetInDB(ReportName_String, _
                                     SheetName_String, _
                                       officeNameInReport, _
                                         (row_ОО_Тюменский + 2 + ном_стр_офис) + ((i - 1) * Размер_блока_офиса), _
                                           НумерацияПунктов("Премия НСЖ МАСС"), _
                                             "Премия НСЖ МАСС", _
                                               "НСЖ_МАСС", _
                                                 "тыс.руб.", _
                                                   0, _
                                                     "Премия, тыс.руб._Месяц", _
                                                       "Премия, тыс.руб._Квартал", _
                                                        3, _
                                                          "Алтайский ОО1", _
                                                             0, _
                                                               0, _
                                                                 0, _
                                                                   0, 1, 1)
        
        Else
          ' Если в DB Лист не найден
          Call в_DB_Лист_не_найден(StringInSheet)
        End If

        ' Кол-во закрытых вкладов
        ' Показатель "Кол-во закрытых вкладов"
        StringInSheet = "Прон-е ИСЖ в закр. вклады"
        SheetName_String = FindNameSheet(ReportName_String, StringInSheet) '
        If SheetName_String <> "" Then

          ном_стр_офис = ном_стр_офис + 1
          Call DB_getParamFromUniversalSheetInDB(ReportName_String, _
                                                   SheetName_String, _
                                                     "Филиал", _
                                                       officeNameInReport, _
                                                         (row_ОО_Тюменский + 2 + ном_стр_офис) + ((i - 1) * Размер_блока_офиса), _
                                                           НумерацияПунктов("Кол-во закр. вкладов (все)"), _
                                                             "Кол-во закр. вкладов (все)", _
                                                               "Кол-во закрытых вкладов", _
                                                                 "Закрытые_Вклады_все_шт", _
                                                                   "шт.", _
                                                                     0, _
                                                                      "Месяц")
        Else
          ' Если в DB Лист не найден
          Call в_DB_Лист_не_найден(StringInSheet)
        End If


        ' Показатель "Сумма закрытых вкладов, тыс.руб."
        StringInSheet = "Прон-е ИСЖ в закр. вклады"
        SheetName_String = FindNameSheet(ReportName_String, StringInSheet) '
        If SheetName_String <> "" Then

          ном_стр_офис = ном_стр_офис + 1
          Call DB_getParamFromUniversalSheetInDB(ReportName_String, _
                                                   SheetName_String, _
                                                     "Филиал", _
                                                       officeNameInReport, _
                                                         (row_ОО_Тюменский + 2 + ном_стр_офис) + ((i - 1) * Размер_блока_офиса), _
                                                           НумерацияПунктов("Закрытые вклады (все)"), _
                                                             "Закрытые вклады (все)", _
                                                               "Сумма закрытых вкладов, тыс.руб.", _
                                                                 "Закрытые_Вклады_все", _
                                                                   "тыс.руб.", _
                                                                     0, _
                                                                      "Месяц")
        Else
          ' Если в DB Лист не найден
          Call в_DB_Лист_не_найден(StringInSheet)
        End If


        ' Кол-во закрытых вкладов, в т.ч. OPC
        ' Показатель "Кол-во закрытых вкладов, в т.ч. OPC"
        StringInSheet = "Прон-е ИСЖ в закр. вклады"
        SheetName_String = FindNameSheet(ReportName_String, StringInSheet) '
        If SheetName_String <> "" Then

          ном_стр_офис = ном_стр_офис + 1
          Call DB_getParamFromUniversalSheetInDB(ReportName_String, _
                                                   SheetName_String, _
                                                     "Филиал", _
                                                       officeNameInReport, _
                                                         (row_ОО_Тюменский + 2 + ном_стр_офис) + ((i - 1) * Размер_блока_офиса), _
                                                           НумерацияПунктов("Кол-во закр. вкладов OPC"), _
                                                             "Кол-во закр. вкладов OPC", _
                                                               " Кол-во закрытых вкладов, в т.ч. OPC", _
                                                                 "Закрытые_Вклады_OPC_шт", _
                                                                   "шт.", _
                                                                     0, _
                                                                      "Месяц")
        Else
          ' Если в DB Лист не найден
          Call в_DB_Лист_не_найден(StringInSheet)
        End If

        ' Сумма закрытых вкладов, тыс.руб. в т.ч OPC
        ' Показатель "Сумма закрытых вкладов, тыс.руб. в т.ч OPC"
        StringInSheet = "Прон-е ИСЖ в закр. вклады"
        SheetName_String = FindNameSheet(ReportName_String, StringInSheet) '
        If SheetName_String <> "" Then

          ном_стр_офис = ном_стр_офис + 1
          Call DB_getParamFromUniversalSheetInDB(ReportName_String, _
                                                   SheetName_String, _
                                                     "Филиал", _
                                                       officeNameInReport, _
                                                         (row_ОО_Тюменский + 2 + ном_стр_офис) + ((i - 1) * Размер_блока_офиса), _
                                                           НумерацияПунктов("Закрытые вклады OPC"), _
                                                             "Закрытые вклады OPC", _
                                                               "Сумма закрытых вкладов, тыс.руб. в т.ч OPC", _
                                                                 "Закрытые_Вклады_OPC", _
                                                                   "тыс.руб.", _
                                                                     0, _
                                                                      "Месяц")
        Else
          ' Если в DB Лист не найден
          Call в_DB_Лист_не_найден(StringInSheet)
        End If


        ' Показатель "Прон-е ИСЖ в закр. вклады"
        ' StringInSheet = "Прон-е ИСЖ в закр. вклады"
        ' SheetName_String = FindNameSheet(ReportName_String, StringInSheet) '
        ' If SheetName_String <> "" Then

        '   ном_стр_офис = ном_стр_офис + 1
        '   Call DB_getParamFromUniversalSheetInDB(ReportName_String, _
        '                                            SheetName_String, _
        '                                             "Филиал", _
        '                                               officeNameInReport, _
        '                                                 (row_ОО_Тюменский + 2 + ном_стр_офис) + ((i - 1) * Размер_блока_офиса), _
        '                                                   НумерацияПунктов("Прон-е ИСЖ в закр.вклады"), _
        '                                                     "Прон-е ИСЖ в закр.вклады", _
        '                                                       " % Проникновение_тыс.руб.", _
        '                                                         "Проникн_ИСЖ_в_закр_вклады", _
        '                                                           "%", _
        '                                                             0, _
        '                                                              "Месяц")
        ' Else
          ' Если в DB Лист не найден
        '  Call в_DB_Лист_не_найден(StringInSheet)
        ' End If


        ' ----------------------------------------------------------------------------------------------------------------------------------
        ' Чертим горизонтальную линию 2 (указываем предидущее значение строки + 1)
        Call gorizontalLineII(ThisWorkbook.Name, "Лист8", (row_ОО_Тюменский + 2 + ном_стр_офис) + ((i - 1) * Размер_блока_офиса) + 1, 2, 12)
        ' ----------------------------------------------------------------------------------------------------------------------------------

        
        
        ' Показатель №15 (23) ДВС (ном_стр_офис=23). Вкладка "ДВС" действовала до февраля 2021 года, затем была переименована в "Портфель пассивов"
        StringInSheet = "ДВС"
        SheetName_String = FindNameSheet(ReportName_String, StringInSheet) ' "3.5.1 ДВС"
        ' Если вкладки "ДВС" нет, то проверяем вкладку "Портфель пассивов"
        StringInSheet = "Портфель пассивов"
        If SheetName_String = "" Then
          SheetName_String = FindNameSheet(ReportName_String, StringInSheet) ' "3.5.1 Портфель пассивов"
        End If
        ' Обработка
        If SheetName_String <> "" Then

          ' ном_стр_офис = ном_стр_офис + 1
          ' Call DB_UniversalSheetInDB(ReportName_String, _
          '                            SheetName_String, _
          '                              officeNameInReport, _
          '                                (row_ОО_Тюменский + 2 + ном_стр_офис) + ((i - 1) * Размер_блока_офиса), _
          '                                  НумерацияПунктов("ДВС"), _
          '                                    "ДВС", _
          '                                      "ДВС", _
          '                                        "тыс.руб.", _
          '                                          0, _
          '                                            "Портфель, тыс.руб._Месяц", _
          '                                              "Портфель, тыс.руб._Квартал", _
          '                                                0, _
          '                                                  "Филиал", _
          '                                                    12, _
          '                                                      12, _
          '                                                        0, _
          '                                                          0, 1, 1)
                                                                   
        ном_стр_офис = ном_стр_офис + 1
        Call DB_UniversalSheetInDB(ReportName_String, _
                                    SheetName_String, _
                                      officeNameInReport, _
                                        (row_ОО_Тюменский + 2 + ном_стр_офис) + ((i - 1) * Размер_блока_офиса), _
                                          НумерацияПунктов("Пассивы"), _
                                            "Пассивы", _
                                              "Пассивы", _
                                                "тыс.руб.", _
                                                  0.2, _
                                                    "Портфель, тыс.руб._Месяц", _
                                                      "Портфель, тыс.руб._Квартал", _
                                                        0, _
                                                          "Филиал", _
                                                            24, _
                                                              24, _
                                                                0, _
                                                                  0, 1, 1)
                   
                   
        ' Показатель "Срочные вклады"
        ном_стр_офис = ном_стр_офис + 1
        Call DB_UniversalSheetInDB(ReportName_String, _
                                     SheetName_String, _
                                       officeNameInReport, _
                                         (row_ОО_Тюменский + 2 + ном_стр_офис) + ((i - 1) * Размер_блока_офиса), _
                                           НумерацияПунктов("в т.ч. Срочные вклады"), _
                                             "в т.ч. Срочные вклады", _
                                               "ВКЛ", _
                                                 "тыс.руб.", _
                                                   0, _
                                                     "Портфель, тыс.руб._Месяц", _
                                                       "Портфель, тыс.руб._Квартал", _
                                                         0, _
                                                           "Филиал", _
                                                             0, _
                                                               0, _
                                                                 0, _
                                                                   0, 1, 1)
                   
                   
        ' Показатель "Накопительный счет"
        ном_стр_офис = ном_стр_офис + 1
        Call DB_UniversalSheetInDB(ReportName_String, _
                                     SheetName_String, _
                                       officeNameInReport, _
                                         (row_ОО_Тюменский + 2 + ном_стр_офис) + ((i - 1) * Размер_блока_офиса), _
                                           НумерацияПунктов("           Накопительный счет"), _
                                             "           Накопительный счет", _
                                               "ДВС_НС", _
                                                 "тыс.руб.", _
                                                   0, _
                                                     "Портфель, тыс.руб._Месяц", _
                                                       "Портфель, тыс.руб._Квартал", _
                                                         0, _
                                                           "Филиал", _
                                                             3, _
                                                               3, _
                                                                 0, _
                                                                   0, 1, 1)
        
          ' Показатель "СКС"
          ном_стр_офис = ном_стр_офис + 1
          Call DB_UniversalSheetInDB(ReportName_String, _
                                     SheetName_String, _
                                       officeNameInReport, _
                                         (row_ОО_Тюменский + 2 + ном_стр_офис) + ((i - 1) * Размер_блока_офиса), _
                                           НумерацияПунктов("           СКС"), _
                                             "           СКС", _
                                               "ДВС_СКС", _
                                                 "тыс.руб.", _
                                                   0, _
                                                     "Портфель, тыс.руб._Месяц", _
                                                       "Портфель, тыс.руб._Квартал", _
                                                         0, _
                                                           "Филиал", _
                                                             6, _
                                                               6, _
                                                                 0, _
                                                                   0, 1, 1)
        
          ' Показатель "Прочие ДВС"
          ном_стр_офис = ном_стр_офис + 1
          Call DB_UniversalSheetInDB(ReportName_String, _
                                     SheetName_String, _
                                       officeNameInReport, _
                                         (row_ОО_Тюменский + 2 + ном_стр_офис) + ((i - 1) * Размер_блока_офиса), _
                                           НумерацияПунктов("           Прочие ДВС"), _
                                             "           Прочие ДВС", _
                                               "ДВС_ПрДВС", _
                                                 "тыс.руб.", _
                                                   0, _
                                                     "Портфель, тыс.руб._Месяц", _
                                                       "Портфель, тыс.руб._Квартал", _
                                                         0, _
                                                           "Филиал", _
                                                             9, _
                                                               9, _
                                                                 0, _
                                                                   0, 1, 1)
        
          ' Показатель "Аккредитивы"
          ном_стр_офис = ном_стр_офис + 1
          Call DB_UniversalSheetInDB(ReportName_String, _
                                     SheetName_String, _
                                       officeNameInReport, _
                                         (row_ОО_Тюменский + 2 + ном_стр_офис) + ((i - 1) * Размер_блока_офиса), _
                                           НумерацияПунктов("           Аккредитивы"), _
                                             "           Аккредитивы", _
                                               "ДВС_АК", _
                                                 "тыс.руб.", _
                                                   0, _
                                                     "Портфель, тыс.руб._Месяц", _
                                                       "Портфель, тыс.руб._Квартал", _
                                                         0, _
                                                           "Филиал", _
                                                             12, _
                                                               12, _
                                                                 0, _
                                                                   0, 1, 1)
                                                                   
          ' Показатель "Брокер"
          ном_стр_офис = ном_стр_офис + 1
          Call DB_UniversalSheetInDB(ReportName_String, _
                                     SheetName_String, _
                                       officeNameInReport, _
                                         (row_ОО_Тюменский + 2 + ном_стр_офис) + ((i - 1) * Размер_блока_офиса), _
                                           НумерацияПунктов("           Брокер"), _
                                             "           Брокер", _
                                               "Пассивы_Брокер", _
                                                 "тыс.руб.", _
                                                   0, _
                                                     "Портфель, тыс.руб._Месяц", _
                                                       "Портфель, тыс.руб._Квартал", _
                                                         0, _
                                                           "Филиал", _
                                                             15, _
                                                               15, _
                                                                 0, _
                                                                   0, 1, 1)
                                                                   
          ' Показатель "УК"
          ном_стр_офис = ном_стр_офис + 1
          Call DB_UniversalSheetInDB(ReportName_String, _
                                     SheetName_String, _
                                       officeNameInReport, _
                                         (row_ОО_Тюменский + 2 + ном_стр_офис) + ((i - 1) * Размер_блока_офиса), _
                                           НумерацияПунктов("           УК"), _
                                             "           УК", _
                                               "Пассивы_УК", _
                                                 "тыс.руб.", _
                                                   0, _
                                                     "Портфель, тыс.руб._Месяц", _
                                                       "Портфель, тыс.руб._Квартал", _
                                                         0, _
                                                           "Филиал", _
                                                             18, _
                                                               18, _
                                                                 0, _
                                                                   0, 1, 1)
                                                                   
          ' Показатель "АУМ"
          ном_стр_офис = ном_стр_офис + 1
          Call DB_UniversalSheetInDB(ReportName_String, _
                                     SheetName_String, _
                                       officeNameInReport, _
                                         (row_ОО_Тюменский + 2 + ном_стр_офис) + ((i - 1) * Размер_блока_офиса), _
                                           НумерацияПунктов("           АУМ"), _
                                             "           АУМ", _
                                               "Пассивы_АУМ", _
                                                 "тыс.руб.", _
                                                   0, _
                                                     "Портфель, тыс.руб._Месяц", _
                                                       "Портфель, тыс.руб._Квартал", _
                                                         0, _
                                                           "Филиал", _
                                                             21, _
                                                               21, _
                                                                 0, _
                                                                   0, 1, 1)
                                                                   
                                                                   
        Else
          ' Если в DB Лист не найден
          Call в_DB_Лист_не_найден(StringInSheet)
        End If
                                                                   
                                                               
        ' ----------------------------------------------------------------------------------------------------------------------------------
        ' Чертим горизонтальную линию 2 (указываем предидущее значение строки + 1)
        Call gorizontalLineII(ThisWorkbook.Name, "Лист8", (row_ОО_Тюменский + 2 + ном_стр_офис) + ((i - 1) * Размер_блока_офиса) + 1, 2, 12)
        ' ----------------------------------------------------------------------------------------------------------------------------------
                                                               
                                                               
        ' Показатель №16 (28) Инвесты - показатель квартальный (ном_стр_офис=28)
        StringInSheet = " ИНВЕСТ"
        SheetName_String = FindNameSheet(ReportName_String, StringInSheet) ' "10. ИНВЕСТ"
        If SheetName_String <> "" Then

          ' 1) ИНВЕСТ MASS
          ном_стр_офис = ном_стр_офис + 1

          ' Наименование_столбца_план = "Итог План, тыс. руб."
          Наименование_столбца_план = "План, тыс. руб."
          
          ' Месяц план/факт
          Call DB_swith_to_MonthQuarter2(ReportName_String, SheetName_String, "Месяц")
          НумерацияПункта_Инвест = НумерацияПунктов("Инвест")
          Call DB_UniversalSheetInDB(ReportName_String, _
                                     SheetName_String, _
                                       officeNameInReport, _
                                         (row_ОО_Тюменский + 2 + ном_стр_офис) + ((i - 1) * Размер_блока_офиса), _
                                           НумерацияПункта_Инвест, _
                                             "Инвест", _
                                               "ИНВ", _
                                                 "тыс. руб.", _
                                                   0, _
                                                     Наименование_столбца_план, _
                                                       "", _
                                                         0, _
                                                           "Алтайский ОО1", _
                                                             0, _
                                                               0, _
                                                                 0, _
                                                                   0, 1, 1)
          
          
          
          ' Квартальный план/факт
          Call DB_swith_to_MonthQuarter2(ReportName_String, SheetName_String, "Квартал")
          Call DB_UniversalSheetInDB(ReportName_String, _
                                     SheetName_String, _
                                       officeNameInReport, _
                                         (row_ОО_Тюменский + 2 + ном_стр_офис) + ((i - 1) * Размер_блока_офиса), _
                                           НумерацияПункта_Инвест, _
                                             "Инвест", _
                                               "ИНВ", _
                                                 "тыс. руб.", _
                                                   0, _
                                                     "", _
                                                       Наименование_столбца_план, _
                                                         0, _
                                                           "Алтайский ОО1", _
                                                             0, _
                                                               0, _
                                                                 0, _
                                                                   0, 1, 1)
                                                                   
                                                                   
          ' 1) ИНВЕСТ OPC
          ном_стр_офис = ном_стр_офис + 1

          
          ' "Продукты УК "Промсвязь" _Affluent"
          ' "Продукты УК "Промсвязь" _Affluent**"
          Наименование_столбца_план = "Продукты УК " + Chr(34) + "Промсвязь" + Chr(34) + " _Affluent**"
          
          ' Месяц план/факт
          Call DB_swith_to_MonthQuarter2(ReportName_String, SheetName_String, "Месяц")
          НумерацияПункта_Инвест = НумерацияПунктов("Инвест OPC")
          Call DB_UniversalSheetInDB(ReportName_String, _
                                     SheetName_String, _
                                       officeNameInReport, _
                                         (row_ОО_Тюменский + 2 + ном_стр_офис) + ((i - 1) * Размер_блока_офиса), _
                                           НумерацияПункта_Инвест, _
                                             "Инвест OPC", _
                                               "ИНВ_OPC", _
                                                 "тыс. руб.", _
                                                   0, _
                                                     Наименование_столбца_план, _
                                                       "", _
                                                         0, _
                                                           "Алтайский ОО1", _
                                                             0, _
                                                               0, _
                                                                 0, _
                                                                   0, 1, 1)
          
          
          
          ' Квартальный план/факт
          Call DB_swith_to_MonthQuarter2(ReportName_String, SheetName_String, "Квартал")
          Call DB_UniversalSheetInDB(ReportName_String, _
                                     SheetName_String, _
                                       officeNameInReport, _
                                         (row_ОО_Тюменский + 2 + ном_стр_офис) + ((i - 1) * Размер_блока_офиса), _
                                           НумерацияПункта_Инвест, _
                                             "Инвест OPC", _
                                               "ИНВ_OPC", _
                                                 "тыс. руб.", _
                                                   0, _
                                                     "", _
                                                       Наименование_столбца_план, _
                                                         0, _
                                                           "Алтайский ОО1", _
                                                             0, _
                                                               0, _
                                                                 0, _
                                                                   0, 1, 1)
                                                                            
                                                                   
          ' 3) ИНВЕСТ Брокерские счета
          ном_стр_офис = ном_стр_офис + 1
          
          Наименование_столбца_план = "Брокерское обслуживание и ИИС*_MASS"
          
          ' Месяц план/факт
          Call DB_swith_to_MonthQuarter2(ReportName_String, SheetName_String, "Месяц")
          НумерацияПункта_Инвест = НумерацияПунктов("Инвест Брокер обслуж")
          Call DB_UniversalSheetInDB(ReportName_String, _
                                     SheetName_String, _
                                       officeNameInReport, _
                                         (row_ОО_Тюменский + 2 + ном_стр_офис) + ((i - 1) * Размер_блока_офиса), _
                                           НумерацияПункта_Инвест, _
                                             "Инвест Брокер обслуж", _
                                               "ИНВ_БО", _
                                                 "шт.", _
                                                   0, _
                                                     Наименование_столбца_план, _
                                                       "", _
                                                         0, _
                                                           "Алтайский ОО1", _
                                                             0, _
                                                               0, _
                                                                 0, _
                                                                   0, 1, 1)
          
          
          
          ' Квартальный план/факт
          Call DB_swith_to_MonthQuarter2(ReportName_String, SheetName_String, "Квартал")
          Call DB_UniversalSheetInDB(ReportName_String, _
                                     SheetName_String, _
                                       officeNameInReport, _
                                         (row_ОО_Тюменский + 2 + ном_стр_офис) + ((i - 1) * Размер_блока_офиса), _
                                           НумерацияПункта_Инвест, _
                                             "Инвест Брокер обслуж", _
                                               "ИНВ_БО", _
                                                 "шт.", _
                                                   0, _
                                                     "", _
                                                       Наименование_столбца_план, _
                                                         0, _
                                                           "Алтайский ОО1", _
                                                             0, _
                                                               0, _
                                                                 0, _
                                                                   0, 1, 1)
                                                                   
                                                                   
                                                                   
          
          ' 4) ИНВЕСТ Брокерские счета Affluent
          ном_стр_офис = ном_стр_офис + 1
          
          Наименование_столбца_план = "Брокерское обслуживание и ИИС*_Affluent"
          
          ' Месяц план/факт
          Call DB_swith_to_MonthQuarter2(ReportName_String, SheetName_String, "Месяц")
          НумерацияПункта_Инвест = НумерацияПунктов("Инвест Брокер обслуж OPC")
          Call DB_UniversalSheetInDB(ReportName_String, _
                                     SheetName_String, _
                                       officeNameInReport, _
                                         (row_ОО_Тюменский + 2 + ном_стр_офис) + ((i - 1) * Размер_блока_офиса), _
                                           НумерацияПункта_Инвест, _
                                             "Инвест Брокер обслуж OPC", _
                                               "ИНВ_БО_OPC", _
                                                 "шт.", _
                                                   0, _
                                                     Наименование_столбца_план, _
                                                       "", _
                                                         0, _
                                                           "Алтайский ОО1", _
                                                             0, _
                                                               0, _
                                                                 0, _
                                                                   0, 1, 1)
          
          
          
          ' Квартальный план/факт
          Call DB_swith_to_MonthQuarter2(ReportName_String, SheetName_String, "Квартал")
          Call DB_UniversalSheetInDB(ReportName_String, _
                                     SheetName_String, _
                                       officeNameInReport, _
                                         (row_ОО_Тюменский + 2 + ном_стр_офис) + ((i - 1) * Размер_блока_офиса), _
                                           НумерацияПункта_Инвест, _
                                             "Инвест Брокер обслуж OPC", _
                                               "ИНВ_БО_OPC", _
                                                 "шт.", _
                                                   0, _
                                                     "", _
                                                       Наименование_столбца_план, _
                                                         0, _
                                                           "Алтайский ОО1", _
                                                             0, _
                                                               0, _
                                                                 0, _
                                                                   0, 1, 1)
                                                                   
        Else
          ' Если в DB Лист не найден
          Call в_DB_Лист_не_найден(StringInSheet)
        End If
                                                                   
                                                               
        ' Показатель №17 (29) ОФЗ (ном_стр_офис=29)
        StringInSheet = "ОФЗ"
        SheetName_String = FindNameSheet(ReportName_String, StringInSheet) ' "7.ОФЗ"
        If SheetName_String <> "" Then
        
          ном_стр_офис = ном_стр_офис + 1
        
          ' Месяц "7.ОФЗ"
          Call DB_swith_to_MonthQuarter(ReportName_String, SheetName_String, 1, "Срез_Период3")
          НумерацияПунктов_ОФЗ = НумерацияПунктов("ОФЗ (в т.ч.OPC)")
          Call DB_UniversalSheetInDB(ReportName_String, _
                                     SheetName_String, _
                                       officeNameInReport, _
                                         (row_ОО_Тюменский + 2 + ном_стр_офис) + ((i - 1) * Размер_блока_офиса), _
                                           НумерацияПунктов_ОФЗ, _
                                             "ОФЗ (в т.ч.OPC)", _
                                               "ОФЗ", _
                                                 "руб.", _
                                                   0, _
                                                     "Итог План, руб.", _
                                                       "", _
                                                         6, _
                                                           "Филиал", _
                                                             0, _
                                                               0, _
                                                                 0, _
                                                                   0, 1, 1)

        
          ' Квартал "7.ОФЗ"
          Call DB_swith_to_MonthQuarter(ReportName_String, SheetName_String, 2, "Срез_Период3")
          Call DB_UniversalSheetInDB(ReportName_String, _
                                     SheetName_String, _
                                       officeNameInReport, _
                                         (row_ОО_Тюменский + 2 + ном_стр_офис) + ((i - 1) * Размер_блока_офиса), _
                                           НумерацияПунктов_ОФЗ, _
                                             "ОФЗ (в т.ч.OPC)", _
                                               "ОФЗ", _
                                                 "руб.", _
                                                   0, _
                                                     "", _
                                                       "Итог План, руб.", _
                                                         6, _
                                                           "Филиал", _
                                                             0, _
                                                               0, _
                                                                 0, _
                                                                   0, 1, 1)
                                                                   
        Else
          ' Если в DB Лист не найден
          Call в_DB_Лист_не_найден(StringInSheet)
        End If
                                                                   
                                                               
        ' ----------------------------------------------------------------------------------------------------------------------------------
        ' Чертим горизонтальную линию 2 (указываем предидущее значение строки + 1)
        Call gorizontalLineII(ThisWorkbook.Name, "Лист8", (row_ОО_Тюменский + 2 + ном_стр_офис) + ((i - 1) * Размер_блока_офиса) + 1, 2, 12)
        ' ----------------------------------------------------------------------------------------------------------------------------------
                                                                   
                                                                   
        ' Показатель "Штат" - вкладка "Комплектность продаж" исключена в марте 2021 (может и ранее)
        ' StringInSheet = "Комплектность продаж"
        ' SheetName_String = FindNameSheet(ReportName_String, StringInSheet) ' "3.1 Потребительские  кредиты"
        ' If SheetName_String <> "" Then
        '    ном_стр_офис = ном_стр_офис + 1
        '    Call DB_UniversalSheetInDB(ReportName_String, _
        '                             SheetName_String, _
        '                               officeNameInReport, _
        '                                 (row_ОО_Тюменский + 2 + ном_стр_офис) + ((i - 1) * Размер_блока_офиса), _
        '                                   НумерацияПунктов("Штат"), _
        '                                     "Штат", _
        '                                       "Штат", _
        '                                         "ед.", _
        '                                           0, _
        '                                             "Штат", _
        '                                               "", _
        '                                                 0, _
        '                                                   "Филиал", _
        '                                                     0, _
        '                                                       0, _
        '                                                         0, _
        '                                                           0, 1, 1)
        '
        'Else
        '  ' Если в DB Лист не найден
        '  Call в_DB_Лист_не_найден(StringInSheet)
        'End If
        
        ' Вариант расчета Штата - берем с Лист7
        ном_стр_офис = ном_стр_офис + 1
        ' DB_Штат
        Call DB_Штат(ReportName_String, _
                                     SheetName_String, _
                                       officeNameInReport, _
                                         (row_ОО_Тюменский + 2 + ном_стр_офис) + ((i - 1) * Размер_блока_офиса), _
                                           НумерацияПунктов("Штат"), _
                                             "Штат", _
                                               "Штат", _
                                                 "ед.", _
                                                   0)
        
        
        ' *** *** ***
        ' Ипотека с листа "Интегральный рейтинг_Регионы"
        StringInSheet = "Интегральный рейтинг_Регионы"
        SheetName_String = FindNameSheet(ReportName_String, StringInSheet)
        If SheetName_String <> "" Then
          
          ' Нумерация
          ном_стр_офис = ном_стр_офис + 1
          Call DB_Ипотека(ReportName_String, _
                                     SheetName_String, _
                                       officeNameInReport, _
                                         (row_ОО_Тюменский + 2 + ном_стр_офис) + ((i - 1) * Размер_блока_офиса), _
                                           НумерацияПунктов("Ипотека"), _
                                             "Ипотека", _
                                               "Ипотека", _
                                                 "тыс.руб.", _
                                                   0.15)
          
                                                                   
        Else
          ' Если в DB Лист не найден
          Call в_DB_Лист_не_найден(StringInSheet)
        End If


   
        ' Выводим данные по офису
      
      Next i ' Следующий офис
            
      ' Строка статуса
      Application.StatusBar = "РОО Тюменский..."
            
      ' *** Вывести свод по РОО в отдельную процедуру

      ' Формирование свода по РОО на основании данных по каждому офису
      Call Свод_по_РОО
  
      ' *** Вывести свод по РОО в отдельную процедуру
                    
      ' Строка статуса
      Application.StatusBar = "Интег-ый рейтинг по офисам..."
      
      
      
      ' Интегральный рейтинг по офисам
      ' Вкладка "1.1 Интег-ый рейтинг  по офисам" действовала по 14.02.2021
      If dateDB <= CDate("14.02.2021") Then
      
          StringInSheet = "Интег-ый рейтинг  по офисам"
          SheetName_String = FindNameSheet(ReportName_String, StringInSheet) ' "1.1 Интег-ый рейтинг  по офисам"
      
          If SheetName_String <> "" Then
            ' Обработка версии ИР офисов до 14.02.2021
            Call DB_rating(ReportName_String, _
                         SheetName_String, _
                           getRowFromSheet8("Интегральный рейтинг по офисам", "Интегральный рейтинг по офисам") + 3, _
                             "Филиал")
          Else
            ' Если в DB Лист не найден
            Call в_DB_Лист_не_найден(StringInSheet)
          End If
      End If
      
      ' Итоги обработки
      
      ' Загружаем изменения за неделю
      ' ThisWorkbook.Sheets("Лист8").Range("O9").Value=... (дата определена в начале)
      Application.StatusBar = "Изменения за неделю..."
      
      ' Старая версия - использует изменение за месяц
      ' Call Загрузить_факт_на_дату
      
      ' Новая версия - использует новую функцию "Факт_Q_на_дату"
      Call Загрузить_факт_на_дату2
      
      ' Строка статуса
      Application.StatusBar = "Копирование итогов..."

      ' Формирование отчета в почту по исполнению ИПЗ
      Call Выполнение_ИПЗ_ГО

      ' Формирование рейтинга регионов
      Call Формирование_рейтинга_регионов(ReportName_String)

      ' Оперативная справка за неделю
      If (ДеньНедели(Date) = "понедельник") Or (ThisWorkbook.Sheets("Лист0").Range("L2").Value = "1") Then
        Call Оперативная_справка_за_неделю
      End If

      ' Копируем итоговый отчет в Книгу для отправки
      Call copyDBToSend

      ' Обрабатываем апдейт по всем дням - пробуем так
      Call Отклонения_по_офисам_Update
      
      ' Если сегодня понедельник, либо Лист_0_L2 = "1", то формируем "Цели на неделю" и отправляем Письма
      ' If (ДеньНедели(Date) = "понедельник") Or (ThisWorkbook.Sheets("Лист0").Range("L2").Value = "1") Then
      If (ДеньНедели(CDate(ThisWorkbook.Sheets("Лист0").Range("E2").Value)) = "понедельник") Or (ThisWorkbook.Sheets("Лист0").Range("L2").Value = "1") Then
        ' Определяем отклонения по офисам и формируем цели на неделю
        Call Отклонения_по_офисам
        Call Отклонения_по_ОКП
      Else
        ' Обновление исполнения плана по целям по свежему DB (обновляет BASE\TargetWeek и оттуда затем в Отклонения_по_офисам, Отклонения_по_ОКП берем информацию по исполнению прошедшей недели)
        ' Call Отклонения_по_офисам_Update - пробуем так
      End If
      
      ' Вносим в S9 диапазон календарной недели, для 15.09 это будет "13.09.2021-19.09.2021"
      ThisWorkbook.Sheets("Лист8").Range("S9").Value = CStr(weekStartDate(dateDB)) + "-" + CStr(dateDB)
      Call Продажи_квартала_за_период
      
      ' Динамика
      ' Call Динамика_продаж
      
      ' Строка статуса
      Application.StatusBar = "Завершение..."

      ' Сохранение изменений
      ThisWorkbook.Save
        
      ' Закрываем BASE\Sales
      CloseBook ("Sales_Office")
                
      ' Закрываем BASE\Sales
      CloseBook ("Sales")
      
      ' Закрываем BASE\Products
      CloseBook ("Products")
                
      ' Переменная завершения обработки
      finishProcess = True
      
      ' Строка статуса
      Application.StatusBar = ""
           
    Else
      ' Сообщение о неверном формате отчета или даты
      MsgBox ("Проверьте отчет: " + CheckFormatReportResult + "!")
    End If ' Проверка формы отчета

    ' Закрываем файл с отчетом без сохранения изменений (параметр SaveChanges:=False)
    Workbooks(Dir(FileName)).Close SaveChanges:=False ' тестирование
    
    ' Переходим в ячейку M2
    ThisWorkbook.Sheets("Лист8").Range("A1").Select
  
    ' Строка статуса
    Application.StatusBar = ""

    ' Зачеркиваем пункт меню на стартовой страницы
    ' Call ЗачеркиваемТекстВячейке("Лист0", "D9")
    ' Call ЗачеркиваемТекстВячейке("Лист0", RangeByValue(ThisWorkbook.Name, "Лист0", "Оперативная справка по _________________", 100, 100))
    
    ' Итоговое сообщение
    If finishProcess = True Then
      ' Отправка сообщения
      Call Отправка_Lotus_Notes_Лист8
      ' Сообщение
      MsgBox ("Обработка " + Dir(ReportName_String) + " завершена!")
    Else
      MsgBox ("Обработка отчета была прервана!")
    End If

  End If ' Если файл был выбран

End Sub

' Показатель из вкладки DB
Sub DB_UniversalSheetInDB(In_ReportName_String, In_Sheets, In_officeNameInReport, In_Row_Лист8, In_N, In_Product_Name, In_Product_Code, In_Unit, In_Weight, In_ColumnNameMonth, In_ColumnNameQuarter, In_DeltaPrediction, In_Заголовок_столбца_офисы, In_ColumnNameMonth_смещение_План, In_ColumnNameQuarter_смещение_План, In_PlanMonth, In_PlanQuarter, In_Fact_Plan_displacement_Month, In_Fact_Plan_displacement_Quarter)
Dim dateDB As Date
    
  ' ***
  ' In_ColumnNameMonth - наименование столбца с планом месяца, например "Премия, тыс.руб._Месяц" для "3.6 ИСЖ_МАСС". Если планов на месяц нет, то In_ColumnNameMonth=""
  ' In_ColumnNameQuarter - наименование столбца с планом квартала, например "Премия, тыс.руб._Квартал" для "3.6 ИСЖ_МАСС"
  ' In_DeltaPrediction - + число столбцов от столбца План (месяца или квартала) в котором находится прогноз выполнения в %, например для "3.6 ИСЖ_МАСС" In_DeltaPrediction=3 ("План", "Факт" (+1), "% Вып-е" (+2), "% Вып-е_Прог" (+3) ). Если столбца "Прогноз" нет, то In_DeltaPrediction = 0
  ' In_Заголовок_столбца_офисы - наименование заголовка на листе, под которым идут филиалы: Алтайский ОО1, Архангельский ОО1, Астраханский ОО1 ...
  ' In_ColumnNameMonth_смещение_План - смещение относительно столбца In_ColumnNameMonth через которое выходим на "План месяца", например для "3.6 ИСЖ_МАСС" это смещение = 0, а для "3.5.1 ДВС" при In_ColumnNameMonth="Портфель, тыс.руб._Месяц" чтобы выйти на "ДВС_Итого-План" нужно In_ColumnNameMonth_смещение_План=12
  ' In_ColumnNameQuarter_смещение_План - смещение относительно столбца In_ColumnNameQuarter через которое выходим на "План квартала", например для, например для "3.6 ИСЖ_МАСС" это смещение = 0, а для "3.5.1 ДВС" при In_ColumnNameMonth="Портфель, тыс.руб._Квартал" чтобы выйти на "ДВС_Итого-План" нужно In_ColumnNameMonth_смещение_План=12
  ' In_PlanMonth - значение плана месяц цифрой, например 80% проникновения в страховки. Если 0, то берем из DB. Примечание - смещение In_ColumnNameMonth_смещение_План тогда = -1
  ' In_PlanQuarter - значение плана квартала цифрой, например 80% проникновения в страховки. Если 0, то берем из DB. Примечание - смещение In_ColumnNameQuarter_смещение_План = -1
  ' In_Fact_Plan_displacement_Month - смещение Факта относительно плана по Месяцу. По умолчанию = 1
  ' In_Fact_Plan_displacement_Quarter - смещение Факта относительно плана по Кварталу. По умолчанию = 1
  ' ***
    
  ' Дата DB
  dateDB = CDate(Mid(Workbooks(In_ReportName_String).Sheets("Оглавление").Cells(1, 1).Value, 23, 10))
  ' Дата DB с Лист8 (должны совпадать)
  dateDB_Лист8 = CDate(Mid(ThisWorkbook.Sheets("Лист8").Range("B5").Value, 52, 10))

  ' Апдейтим таблицу BASE\Products
  Call Update_BASE_Products(In_Product_Name, In_Product_Code, In_Unit)
  
  ' Вкладка In_Sheets
  ' 42
  Row_Заголовок_столбца_офисы = rowByValue(In_ReportName_String, In_Sheets, In_Заголовок_столбца_офисы, 300, 300) ' было 1000 1000
  ' 2
  Column_Заголовок_столбца_офисы = ColumnByValue(In_ReportName_String, In_Sheets, In_Заголовок_столбца_офисы, 300, 300)
  
  ' Выдачи_тыс_руб_Месяц - столбец "Выдачи, тыс.руб._Месяц" (в строке "Показатель")
  If In_ColumnNameMonth <> "" Then
    
    ' План (BK) 63
    Column_Продажи_Месяц_План = ColumnByValue(In_ReportName_String, In_Sheets, In_ColumnNameMonth, 500, 500) + In_ColumnNameMonth_смещение_План  ' "Выдачи, тыс.руб._Месяц" было 1000 1000
    ' Функция ColumnByValue3 - без удаления пробелов в строке поиска. Попробовал - не работает на ОФЗ! Вернул
    ' Column_Продажи_Месяц_План = ColumnByValue3(In_ReportName_String, In_Sheets, In_ColumnNameMonth, 500, 500) + In_ColumnNameMonth_смещение_План  ' "Выдачи, тыс.руб._Месяц" было 1000 1000
    
    ' Если столбец не найден - выдаем сообщение:
    If Column_Продажи_Месяц_План = 0 Then
      
      ' Заносим StringInSheet в переменную Строка_нет_листа_в_DB
      If InStr(Строка_нет_столбца_на_листе_в_DB, In_ColumnNameMonth) = 0 Then
    
        Строка_нет_столбца_на_листе_в_DB = Строка_нет_столбца_на_листе_в_DB + In_ColumnNameMonth + ", "
        ' Выводим сообщение
        MsgBox ("Внимание! По " + In_Product_Name + " не найден " + In_ColumnNameMonth + "!")

      End If
    
    End If
    
    ' Факт (BL) 64
    ' Column_Продажи_Месяц_Факт = Column_Продажи_Месяц_План + 1
    Column_Продажи_Месяц_Факт = Column_Продажи_Месяц_План + In_Fact_Plan_displacement_Month
    
    ' Прогноз (BO) 67
    If In_DeltaPrediction <> 0 Then
      Column_Продажи_Месяц_Прогноз = Column_Продажи_Месяц_План + In_DeltaPrediction ' (+ 4) параметр In_DeltaPrediction - это через сколько столбец с прогнозом в %
    End If
    
  End If
  
  ' Выдачи_тыс_руб_Квартал - столбец "Выдачи, тыс.руб._Квартал" (в строке "Показатель")
  ' План (CP) 94
  Column_Продажи_Квартал_План = ColumnByValue(In_ReportName_String, In_Sheets, In_ColumnNameQuarter, 500, 500) + In_ColumnNameQuarter_смещение_План ' "Выдачи, тыс.руб._Квартал" было 1000 1000
  ' Без удаления пробелов в поиске - ColumnByValue3. Не работает на ОФЗ, вернул!
  ' Column_Продажи_Квартал_План = ColumnByValue3(In_ReportName_String, In_Sheets, In_ColumnNameQuarter, 500, 500) + In_ColumnNameQuarter_смещение_План ' "Выдачи, тыс.руб._Квартал" было 1000 1000
  
  
  ' Факт (CQ) 95
  ' Column_Продажи_Квартал_Факт = Column_Продажи_Квартал_План + 1
  Column_Продажи_Квартал_Факт = Column_Продажи_Квартал_План + In_Fact_Plan_displacement_Quarter
   
  ' Прогноз (CT) 98
  If In_DeltaPrediction <> 0 Then
    Column_Продажи_Квартал_Прогноз = Column_Продажи_Квартал_План + In_DeltaPrediction ' (+ 4) параметр In_DeltaPrediction - это через сколько столбец с прогнозом в %
  End If
  
  ' Заносим наименование продукта на Лист8
  ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 1).NumberFormat = "@"
  ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 1).Value = In_N
  ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 1).HorizontalAlignment = xlCenter
  '
  ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 2).Value = In_Product_Name
  ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 2).HorizontalAlignment = xlLeft
  ' Вес выводим, если он не нулевой
  If In_Weight <> 0 Then
    ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 3).Value = In_Weight
    ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 3).NumberFormat = "0.0%"
    ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 3).HorizontalAlignment = xlCenter
  End If
  '
  ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 4).Value = In_Unit
  ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 4).HorizontalAlignment = xlCenter

  ' Контрольный показатель - "Тюменский ОО1"
  Офис_найден = False
  
  ' Контрольный показатель - In_officeNameInReport ("ОО2") найден в "Тюменский ОО1"
  ОО2_найден = False

  ' Находим в с столбце "Тюменский ОО1"
  rowCount = Row_Заголовок_столбца_офисы + 1
  Do While (InStr(Workbooks(In_ReportName_String).Sheets(In_Sheets).Cells(rowCount, Column_Заголовок_столбца_офисы).Value, "Общий итог") = 0) And (Not IsEmpty(Workbooks(In_ReportName_String).Sheets(In_Sheets).Cells(rowCount, Column_Заголовок_столбца_офисы).Value))
    
    ' Если это "Тюменский ОО1" - Раскрываем список
    If InStr(Workbooks(In_ReportName_String).Sheets(In_Sheets).Cells(rowCount, Column_Заголовок_столбца_офисы).Value, "Тюменский ОО1") <> 0 Then
      Workbooks(In_ReportName_String).Sheets(In_Sheets).Cells(rowCount, Column_Заголовок_столбца_офисы).ShowDetail = True
      Офис_найден = True
    End If
              
    ' Если это текущий офис
    If (InStr(Workbooks(In_ReportName_String).Sheets(In_Sheets).Cells(rowCount, Column_Заголовок_столбца_офисы).Value, In_officeNameInReport) <> 0) And (InStr(Workbooks(In_ReportName_String).Sheets(In_Sheets).Cells(rowCount, Column_Заголовок_столбца_офисы).Value, "ОО1") = 0) Then
      
      ' Контрольный показатель - In_officeNameInReport ("ОО2") найден в "Тюменский ОО1"
      ОО2_найден = True
      
      ' Берем из этой строки данные и копируем на Лист8
      
      ' Квартал:
      ' If (In_ColumnNameQuarter <> "") Then
      If (In_ColumnNameQuarter <> "") And (Column_Продажи_Квартал_План <> 0) Then ' 21.09 для обработки прошлых DB
        
        ' Квартал - план
        If In_PlanQuarter = 0 Then
          ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 5).Value = Workbooks(In_ReportName_String).Sheets(In_Sheets).Cells(rowCount, Column_Продажи_Квартал_План).Value
        Else
          ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 5).Value = In_PlanQuarter
        End If
        ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 5).NumberFormat = "#,##0"
        ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 5).HorizontalAlignment = xlRight
        

        ' Квартал - факт
        ' Если измерение в %
        If In_Unit <> "%" Then
          
          ' Квартал факт
          ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 6).Value = Workbooks(In_ReportName_String).Sheets(In_Sheets).Cells(rowCount, Column_Продажи_Квартал_Факт).Value
          
          ' === Сюда вставляем цель на неделю - сколько надо прирасти, чтобы выйти на прогноз Q в 100%
          If Прогноз_квартала_проц(dateDB, ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 5).Value, ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 6).Value, 5, 0) < 1 Then
            
            ' Первый вариант: Считаем какой должен быть прогноз - на текущая дата DB + 7
            ' Факт_на_дату_для_прогноза_квартала_Var = Факт_на_дату_для_прогноза_квартала(dateDB + 7, ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 5).Value, 1, 5, 0)
          
            ' Первая дата из M9. Пример: для "План на неделю: (09.08-15.08.21)" это будет 09.08.2021 (пнд). Для нее DB это четверг - 05.08.2021 (минус 4 дня)
            date1FromM9 = CDate(Mid(ThisWorkbook.Sheets("Лист8").Range("M9").Value, 18, 5) + ".20" + Mid(ThisWorkbook.Sheets("Лист8").Range("M9").Value, 30, 2))
          
            ' Второй вариант: Считаем какой должен быть прогноз - из M9 "План на неделю: (02.08-08.08.21)" берем вторую дату
            date2FromM9 = CDate(Mid(ThisWorkbook.Sheets("Лист8").Range("M9").Value, 24, 6) + "20" + Mid(ThisWorkbook.Sheets("Лист8").Range("M9").Value, 30, 2))
            
            ' Отставание DB: от воскресенья (конец недели ) - 3 дня = четверг!
            date2FromM9 = date2FromM9 - 3
            Факт_на_дату_для_прогноза_квартала_Var = Факт_на_дату_для_прогноза_квартала(date2FromM9, ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 5).Value, 1, 5, 0)
            
            ' Факт_квартала: если сегодня понедельник, то берем из 6-го столбца Лист8, если вторник и т.д., то берем из Факт_Q_на_дату
            ' Если дата начала из M9 - 4 (отставание отчетности) = dateDB_Лист8, то факт Q, берем из 6-го столбца Лист8, а если нет, то берем из Факт_Q_на_дату
            If (date1FromM9 - 4) = dateDB_Лист8 Then
              Факт_квартала_Var = ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 6).Value
            Else
              Факт_квартала_Var = Факт_Q_на_дату(getNumberOfficeByName(In_officeNameInReport), In_Product_Code, (date1FromM9 - 4))
            End If
            
            ' Если Факт для выхода на прогноз Q больше, чем текущий Факт Q, то считаем прирост
            If Факт_на_дату_для_прогноза_квартала_Var > ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 6).Value Then
              
              ' В 13-ый столбец пишем план недели
              ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 13).Value = Факт_на_дату_для_прогноза_квартала_Var - Факт_квартала_Var ' ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 6).Value
              ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 13).NumberFormat = "#,##0"
              ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 13).HorizontalAlignment = xlRight
              
              ' В 14-ый столбец пишем исполнение Плана недели (из 13-го столбца)
              ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 14).Value = ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 6).Value - Факт_квартала_Var ' ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 6).Value
              ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 14).NumberFormat = "#,##0"
              ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 14).HorizontalAlignment = xlRight
            
            End If
          
          End If
          ' ===
          
        Else
          ' Если это %, то умножаем на 100
          ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 6).Value = (Workbooks(In_ReportName_String).Sheets(In_Sheets).Cells(rowCount, Column_Продажи_Квартал_Факт).Value * 100)
        End If
        ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 6).NumberFormat = "#,##0"
        ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 6).HorizontalAlignment = xlRight
        

        ' Квартал - исполнение (в %)
        ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 7).Value = РассчетДоли(ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 5).Value, ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 6).Value, 3)
        ' ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 7).NumberFormat = "0.0%"
        ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 7).NumberFormat = "0%"
        ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 7).HorizontalAlignment = xlRight
        
        ' Если столбца "Прогноз" нет (In_DeltaPrediction = 0), то Окраска ячейки СФЕТОФОР ' 70. Заливка ячейки цветом "светофор" - значение In_Value в %, In_Target в %
        If (In_DeltaPrediction = 0) And (ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 5).Value <> 0) Then
          ' Окраска ячейки СФЕТОФОР ' 70. Заливка ячейки цветом "светофор" - значение In_Value в %, In_Target в %
          Call Full_Color_RangeII("Лист8", In_Row_Лист8, 7, ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 7).Value, 1)
        End If
      
        ' Квартал - прогноз
        If (In_DeltaPrediction <> 0) And (Workbooks(In_ReportName_String).Sheets(In_Sheets).Cells(rowCount, Column_Продажи_Квартал_План).Value <> 0) Then
          ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 8).Value = Workbooks(In_ReportName_String).Sheets(In_Sheets).Cells(rowCount, Column_Продажи_Квартал_Прогноз).Value
          ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 8).NumberFormat = "0%"
          ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 8).HorizontalAlignment = xlRight
          ' Окраска ячейки СФЕТОФОР ' 70. Заливка ячейки цветом "светофор" - значение In_Value в %, In_Target в %
          Call Full_Color_RangeII("Лист8", In_Row_Лист8, 8, ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 8).Value, 1)
          
          
        End If
        
        ' ***
        ' Тестирование Функции "Прогноз_квартала" по всем позициям, если измерение не в %
        If ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 4).Value <> "%" Then
          ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 20).Value = Прогноз_квартала_проц(dateDB, ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 5).Value, ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 6).Value, 5, 0)
          ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 20).NumberFormat = "0%"
          ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 20).HorizontalAlignment = xlRight
        End If
        ' ***
        
        ' Если по продукту есть квартальный показатель, критерий: (In_ColumnNameMonth = "") AND (In_ColumnNameQuarter <>"")
        If (In_ColumnNameMonth = "") And (In_ColumnNameQuarter <> "") Then
        
          ' Заносим в Sales_Office
          '  Идентификатор ID_Rec:
          ID_RecVar = CStr(CStr(getNumberOfficeByName(In_officeNameInReport)) + "-" + strNQYY(dateDB) + "-" + In_Product_Code)
                        
          ' Текущие значения в месяце: Date_01 (N), Date_02 (O), Date_03 (P), Date_04 (Q), Date_05 Date_06 Date_07 Date_08 Date_09 Date_10 Date_11 Date_12 Date_13 Date_14 Date_15 Date_16 Date_17 Date_18 Date_19 Date_20 Date_21 Date_22 Date_23 Date_24 Date_25 Date_26 Date_27 Date_28 Date_29 Date_30 Date_31
          ' Номер месяца в квартале: 1-"", 2-"2", 3-"3"
          M_num = Nom_mes_quarter_str(dateDB)
          curr_Day_Month_Q = "Date" + M_num + "_" + Mid(dateDB, 1, 2)
                                      
          ' Вносим данные в BASE\Sales_Office по ПК.
          Call InsertRecordInBook("Sales_Office", "Лист1", "ID_Rec", ID_RecVar, _
                                            "ID_Rec", ID_RecVar, _
                                              "Оffice_Number", getNumberOfficeByName(In_officeNameInReport), _
                                                "Product_Name", In_Product_Name, _
                                                  "Оffice", In_officeNameInReport, _
                                                    "MMYY", strNQYY(dateDB), _
                                                      "Update_Date", dateDB, _
                                                       "Product_Code", In_Product_Code, _
                                                         "Plan", ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 5).Value, _
                                                            "Unit", In_Unit, _
                                                              "Fact", ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 6).Value, _
                                                                "Percent_Completion", ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 7).Value, _
                                                                  curr_Day_Month_Q, ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 6).Value, _
                                                                    "", "", _
                                                                      "", "", _
                                                                        "", "", _
                                                                          "", "", _
                                                                            "", "", _
                                                                              "", "", _
                                                                                "", "", _
                                                                                  "", "")

        
        End If
        
        
        
      End If
                  
      ' Месяц:
      If (In_ColumnNameMonth <> "") And (Column_Продажи_Месяц_План <> 0) Then ' 21.09 для обработки прошлых DB
        
        ' Месяц - план
        If In_PlanMonth = 0 Then
          ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 9).Value = Workbooks(In_ReportName_String).Sheets(In_Sheets).Cells(rowCount, Column_Продажи_Месяц_План).Value
        Else
          ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 9).Value = In_PlanMonth
        End If
        ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 9).NumberFormat = "#,##0"
        ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 9).HorizontalAlignment = xlRight
        
        
        ' Месяц - факт
        ' Если измерение в %
        If In_Unit <> "%" Then
          ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 10).Value = Workbooks(In_ReportName_String).Sheets(In_Sheets).Cells(rowCount, Column_Продажи_Месяц_Факт).Value
        Else
          ' Если это %, то умножаем на 100
          ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 10).Value = (Workbooks(In_ReportName_String).Sheets(In_Sheets).Cells(rowCount, Column_Продажи_Месяц_Факт).Value * 100)
        End If
        ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 10).NumberFormat = "#,##0"
        ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 10).HorizontalAlignment = xlRight
            
        ' Месяц - исполнение
        ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 11).Value = РассчетДоли(ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 9).Value, ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 10).Value, 3)
        ' ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 11).NumberFormat = "0.0%"
        ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 11).NumberFormat = "0%"
        ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 11).HorizontalAlignment = xlRight
        ' Если столбца "Прогноз" нет (In_DeltaPrediction = 0), то Окраска ячейки СФЕТОФОР ' 70. Заливка ячейки цветом "светофор" - значение In_Value в %, In_Target в %
        If (In_DeltaPrediction = 0) And (ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 9).Value <> 0) Then
          ' Окраска ячейки СФЕТОФОР ' 70. Заливка ячейки цветом "светофор" - значение In_Value в %, In_Target в %
          Call Full_Color_RangeII("Лист8", In_Row_Лист8, 11, ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 11).Value, 1)
        End If

        ' Месяц - прогноз (штуки, тыс.руб и т.п.) делаем расчет
        If (In_DeltaPrediction <> 0) And (Workbooks(In_ReportName_String).Sheets(In_Sheets).Cells(rowCount, Column_Продажи_Месяц_План).Value <> 0) Then
      
          PredictionVar = (ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 9).Value) * Workbooks(In_ReportName_String).Sheets(In_Sheets).Cells(rowCount, Column_Продажи_Месяц_Прогноз).Value
                
          ' Месяц - прогноз, %
          ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 12).Value = Workbooks(In_ReportName_String).Sheets(In_Sheets).Cells(rowCount, Column_Продажи_Месяц_Прогноз).Value
          ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 12).NumberFormat = "0%"
          ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 12).HorizontalAlignment = xlRight
          PredictionPercent = ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 12).Value
          ' Окраска ячейки СФЕТОФОР ' 70. Заливка ячейки цветом "светофор" - значение In_Value в %, In_Target в %
          Call Full_Color_RangeII("Лист8", In_Row_Лист8, 12, ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 12).Value, 1)
        Else
          
          ' Если прогноза нет по продукту в DB
          PredictionVar = 0
          PredictionPercent = 0
        
        End If
      
        ' Заносим в Sales_Office
        '  Идентификатор ID_Rec:
        ID_RecVar = CStr(CStr(getNumberOfficeByName(In_officeNameInReport)) + "-" + strMMYY(dateDB) + "-" + In_Product_Code)
            
        ' Текущие значения в месяце: Date_01 (N), Date_02 (O), Date_03 (P), Date_04 (Q), Date_05 Date_06 Date_07 Date_08 Date_09 Date_10 Date_11 Date_12 Date_13 Date_14 Date_15 Date_16 Date_17 Date_18 Date_19 Date_20 Date_21 Date_22 Date_23 Date_24 Date_25 Date_26 Date_27 Date_28 Date_29 Date_30 Date_31
        curr_Day_Month = "Date_" + Mid(dateDB, 1, 2)
            
        ' Вносим данные в BASE\Sales_Office по ПК.
        Call InsertRecordInBook("Sales_Office", "Лист1", "ID_Rec", ID_RecVar, _
                                            "ID_Rec", ID_RecVar, _
                                              "Оffice_Number", getNumberOfficeByName(In_officeNameInReport), _
                                                "Product_Name", In_Product_Name, _
                                                  "Оffice", In_officeNameInReport, _
                                                    "MMYY", strMMYY(dateDB), _
                                                      "Update_Date", dateDB, _
                                                       "Product_Code", In_Product_Code, _
                                                         "Plan", ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 9).Value, _
                                                            "Unit", In_Unit, _
                                                              "Fact", ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 10).Value, _
                                                                "Percent_Completion", ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 11).Value, _
                                                                  "Prediction", PredictionVar, _
                                                                    "Percent_Prediction", PredictionPercent, _
                                                                      curr_Day_Month, ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 10).Value, _
                                                                        "", "", _
                                                                          "", "", _
                                                                            "", "", _
                                                                              "", "", _
                                                                                "", "", _
                                                                                  "", "")

      End If ' If In_ColumnNameMonth <> "" Then
      
    End If
    
    ' Следующая запись
    Application.StatusBar = In_Product_Code + " " + In_officeNameInReport + ": " + CStr(rowCount) + "..."
    rowCount = rowCount + 1
    DoEventsInterval (rowCount)
    
  Loop
  
  ' Контрольный показатель - если все 4 равны нулю, то данные из DB взяты не корректно
  If (Офис_найден = False) Then
    
    ' Если в DB Лист не найден
    MsgBox ("Внимание! По " + In_Product_Name + " не найдены Офисы!")

  End If

  ' Контрольный показатель - In_officeNameInReport ("ОО2") найден в "Тюменский ОО1"
  If ОО2_найден = False Then
    ' Не найден офис ОО2
    ' t = 0
  
    ' Анализируем переданные переменные In_ColumnNameMonth, In_ColumnNameQuarter
    If In_ColumnNameMonth <> "" Then
      ' Заносим по месяцу нули План=0, Факт=0, Исп.=0%, Прогноз=0%
      ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 9).Value = 0
      ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 10).Value = 0
      ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 11).Value = 0
      ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 12).Value = 0
    End If
  
    ' Анализируем переданные переменные In_ColumnNameMonth, In_ColumnNameQuarter
    If In_ColumnNameQuarter <> "" Then
      ' Заносим по кварталу нули План=0, Факт=0, Исп.=0%, Прогноз=0%
      ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 5).Value = 0
      ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 6).Value = 0
      ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 7).Value = 0
      ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 8).Value = 0
      
    End If
  
  
  End If


  
End Sub



' Создание файла для отправки в офисы
Sub copyDBToSend_with_Msg()
  
  ' Запрос
  If MsgBox("Сформировать файл для отправки?", vbYesNo) = vbYes Then
    Call copyDBToSend
    MsgBox ("Данные скопированы!")
  End If

End Sub


' Создание файла для отправки в офисы
Sub copyDBToSend()
Dim TemplatesFile As String

  Application.StatusBar = "Копирование..."

  ' Открываем шаблон "Ежедневный отчет.xlsx"
  If Dir(ThisWorkbook.Path + "\Templates\" + "Ежедневный отчет.xlsx") <> "" Then
    ' Открываем шаблон Templates\Ежедневный отчет по продажам
    TemplatesFileName = "Ежедневный отчет"
  End If
              
  ' Открываем шаблон Протокола из C:\Users\...\Documents\#VBA\DB_Result\Templates
  Workbooks.Open (ThisWorkbook.Path + "\Templates\" + TemplatesFileName + ".xlsx")
           
  ' Переходим на окно DB
  ThisWorkbook.Sheets("Лист8").Activate

  ' Обновляем список получателей
  ThisWorkbook.Sheets("Лист8").Cells(rowByValue(ThisWorkbook.Name, "Лист8", "Список получателей:", 100, 100), ColumnByValue(ThisWorkbook.Name, "Лист8", "Список получателей:", 100, 100) + 2).Value = _
    getFromAddrBook("УДО2,УДО3,УДО4,УДО5,НОРПиКО1,НОРПиКО2,НОРПиКО3,НОРПиКО4,НОРПиКО5,ПМ,МРК1,МРК2,МРК3,МРК4,МРК5,НОКП,РРКК,МПП,РИЦ,СотрИЦ", 2)

  ' Имя нового файла
  FileDBName = Replace(Mid(ThisWorkbook.Sheets("Лист8").Range("B5").Value, 1, 61), ".", "-") + ".xlsx"
  
  ' Проверяем - если файл есть, то удаляем его
  Call deleteFile(ThisWorkbook.Path + "\Out\" + FileDBName)
  
  Workbooks(TemplatesFileName).SaveAs FileName:=ThisWorkbook.Path + "\Out\" + FileDBName, FileFormat:=xlOpenXMLWorkbook, createBackUp:=False
  ThisWorkbook.Sheets("Лист8").Range("Q3").Value = ThisWorkbook.Path + "\Out\" + FileDBName
            
  ' *** Копирование данных ***
 
  ' Находим номер строки с наименованием офиса
  row_ОО_Тюменский = getRowFromSheet8("ОО «Тюменский»", "ОО «Тюменский»")
  row_ОО_Сургутский = getRowFromSheet8("ОО «Сургутский»", "ОО «Сургутский»")
  row_Итого_по_РОО_Тюменский = getRowFromSheet8("Итого по РОО «Тюменский»", "Итого по РОО «Тюменский»")
  row_Интегральный_рейтинг_по_офисам = getRowFromSheet8("Интегральный рейтинг по офисам", "Интегральный рейтинг по офисам")
  Размер_блока_офиса = row_ОО_Сургутский - row_ОО_Тюменский
  
  row_Лист1_Интегральный_рейтинг_по_офисам = rowByValue(FileDBName, "Лист1", "1. Интегральный рейтинг по офисам", 100, 100)

  ' Копируем Интегральный рейтинг по офисам
  countString_Лист1 = 0
  
  ' Дата в Изм. Интегрального рейтинга
  Workbooks(FileDBName).Sheets("Лист1").Cells(5, 15).Value = "Дата " + CStr(ThisWorkbook.Sheets("Лист8").Range("O9").Value)
  
  
  ' For i = (row_Интегральный_рейтинг_по_офисам + 3) To (row_Интегральный_рейтинг_по_офисам + 3) + 5
  '
  ' Счетчик строк
  '  countString_Лист1 = countString_Лист1 + 1
  '
  '  For j = 1 To 16 ' с учетом Изм.
  '    ThisWorkbook.Sheets("Лист8").Cells(i, j).Copy Destination:=Workbooks(FileDBName).Sheets("Лист1").Cells((row_Лист1_Интегральный_рейтинг_по_офисам + 2) + countString_Лист1, j)
  '  Next j
  '
  ' Next i
  
  For i = (row_Интегральный_рейтинг_по_офисам + 1) To (row_Интегральный_рейтинг_по_офисам + 1) + 1 + 5
    
    ' Счетчик строк
    countString_Лист1 = countString_Лист1 + 1
  
    For j = 1 To 9 ' с учетом Изм.
      ThisWorkbook.Sheets("Лист8").Cells(i, j).Copy Destination:=Workbooks(FileDBName).Sheets("Лист1").Cells((row_Лист1_Интегральный_рейтинг_по_офисам) + countString_Лист1, j)
    Next j
  
  Next i
  
  
  ' *** Копируем интегральный рейтинг по сотрудникам с Листа7
  ' Проверяем даты загрузки на Лист7 и Лист8. Копируем, если они совпадают
  dateDB_Лист8 = CDate(Mid(ThisWorkbook.Sheets("Лист8").Range("B5").Value, 52, 10))
  dateDB_Лист7 = CDate(Mid(ThisWorkbook.Sheets("Лист7").Range("B5").Value, 40, 10))
  
  If dateDB_Лист8 = dateDB_Лист7 Then
    
    ' Заголовки ИР по сотрудникам
    Workbooks(FileDBName).Sheets("Лист1").Cells(14, 1).Value = "ИР"
    Workbooks(FileDBName).Sheets("Лист1").Cells(15, 1).Value = "%"
    
    ' Наименования продуктов
    For i = 1 To 10
      Workbooks(FileDBName).Sheets("Лист1").Cells(14, 4 + 5 * (i - 1)).Value = ThisWorkbook.Sheets("Лист7").Cells(7, 6 + 5 * (i - 1)).Value
      ' Workbooks(FileDBName).Sheets("Лист1").Cells(14, 9) = ThisWorkbook.Sheets("Лист7").Cells(7, 11).Value
    Next i
    
    For i = 9 To 24
      
      ' Интегральный рейтинг
      For j = 60 To 60
        ThisWorkbook.Sheets("Лист7").Cells(i, j).Copy Destination:=Workbooks(FileDBName).Sheets("Лист1").Cells(i + 7, j - 59)
      Next j
            
      ' Показатели
      For j = 3 To 54
        ThisWorkbook.Sheets("Лист7").Cells(i, j).Copy Destination:=Workbooks(FileDBName).Sheets("Лист1").Cells(i + 7, j - 1)
      Next j
      
      ' Число продуктов и примечание
      For j = 61 To 63
        ThisWorkbook.Sheets("Лист7").Cells(i, j).Copy Destination:=Workbooks(FileDBName).Sheets("Лист1").Cells(i + 7, j - 6)
      Next j
      
    Next i
  End If
  ' *** Копируем интегральный рейтинг по сотрудникам с Листа7
 
  ' Копируем заголовки
  ThisWorkbook.Sheets("Лист8").Cells(5, 2).Copy Destination:=Workbooks(FileDBName).Sheets("Лист1").Cells(2, 1)
  
  ' Здесь 33 это номер строки в шаблоне ежедневного отчета
  
  For k = 1 To 5
    
    ' Наименование офиса
    Workbooks(FileDBName).Sheets("Лист1").Cells(33 + ((k - 1) * Размер_блока_офиса), 2).Value = Workbooks(FileDBName).Sheets("Лист1").Cells(33 + ((k - 1) * Размер_блока_офиса), 2).Value + ThisWorkbook.Sheets("Лист8").Cells(7 + ((k - 1) * Размер_блока_офиса), 2).Value
    ' Заголовок офиса
    ' Квартал
    Workbooks(FileDBName).Sheets("Лист1").Cells(33 + 1 + ((k - 1) * Размер_блока_офиса), 5).Value = ThisWorkbook.Sheets("Лист8").Cells(7 + 1 + ((k - 1) * Размер_блока_офиса), 5).Value
    ' Месяц
    Workbooks(FileDBName).Sheets("Лист1").Cells(33 + 1 + ((k - 1) * Размер_блока_офиса), 9).Value = ThisWorkbook.Sheets("Лист8").Cells(7 + 1 + ((k - 1) * Размер_блока_офиса), 9).Value
    ' Факт на
    Workbooks(FileDBName).Sheets("Лист1").Cells(33 + 1 + 1 + ((k - 1) * Размер_блока_офиса), 6).Value = ThisWorkbook.Sheets("Лист8").Cells(7 + 1 + 1 + ((k - 1) * Размер_блока_офиса), 6).Value
    Workbooks(FileDBName).Sheets("Лист1").Cells(33 + 1 + 1 + ((k - 1) * Размер_блока_офиса), 10).Value = ThisWorkbook.Sheets("Лист8").Cells(7 + 1 + 1 + ((k - 1) * Размер_блока_офиса), 10).Value
        
    ' Прошлый период и Динамика
    Workbooks(FileDBName).Sheets("Лист1").Cells(33 + 1 + ((k - 1) * Размер_блока_офиса), 15).Value = "Дата " + CStr(ThisWorkbook.Sheets("Лист8").Range("O9").Value)
        
  Next k
 
   
   
  ' Копируем ячейки
  For k = 1 To 5
    
    ' For i = 10 + ((k - 1) * Размер_блока_офиса) To 44 + ((k - 1) * Размер_блока_офиса)
    For i = (row_ОО_Тюменский + 3) + ((k - 1) * Размер_блока_офиса) To (row_ОО_Сургутский - 1) + ((k - 1) * Размер_блока_офиса)
      
      ' Показатели текущего периода
      For j = 1 To 12
        ThisWorkbook.Sheets("Лист8").Cells(i, j).Copy Destination:=Workbooks(FileDBName).Sheets("Лист1").Cells(i + 26, j)
      Next j
      
      ' Показатели прошлого периода и Динамики
      For j = 15 To 18
        ThisWorkbook.Sheets("Лист8").Cells(i, j).Copy Destination:=Workbooks(FileDBName).Sheets("Лист1").Cells(i + 26, j)
      Next j
      
      
    Next i
      
      ' Строка статуса
      Application.StatusBar = "Копирование итогов " + CStr(k) + "..."
      DoEvents
    
  Next k
  
  ' Заголовок Итого по РОО «Тюменский» (адаптирован!)
  
  ' Находим на Лист1 "8. Продажи "
  row_8_Продажи = rowByValue(Workbooks(FileDBName).Name, "Лист1", "8. Продажи ", 1000, 3)
  
  Workbooks(FileDBName).Sheets("Лист1").Cells(row_8_Продажи, 2).Value = ThisWorkbook.Sheets("Лист8").Cells(row_Итого_по_РОО_Тюменский, 2).Value
  
  ' Квартал
  Workbooks(FileDBName).Sheets("Лист1").Cells(row_8_Продажи + 1, 5).Value = ThisWorkbook.Sheets("Лист8").Cells(row_Итого_по_РОО_Тюменский + 1, 5).Value
  ' Месяц
  Workbooks(FileDBName).Sheets("Лист1").Cells(row_8_Продажи + 1, 9).Value = ThisWorkbook.Sheets("Лист8").Cells(row_Итого_по_РОО_Тюменский + 1, 9).Value
  ' Факт на
  Workbooks(FileDBName).Sheets("Лист1").Cells(row_8_Продажи + 2, 6).Value = ThisWorkbook.Sheets("Лист8").Cells(row_Итого_по_РОО_Тюменский + 2, 6).Value
  Workbooks(FileDBName).Sheets("Лист1").Cells(row_8_Продажи + 2, 10).Value = ThisWorkbook.Sheets("Лист8").Cells(row_Итого_по_РОО_Тюменский + 2, 10).Value
  ' Прошлый период и Динамика
  Workbooks(FileDBName).Sheets("Лист1").Cells(row_8_Продажи + 1, 15).Value = "Дата " + CStr(ThisWorkbook.Sheets("Лист8").Range("O9").Value)

  ' Копируем по ячейки Итого по РОО «Тюменский»
  For i = (row_Итого_по_РОО_Тюменский + 3) To (row_Итого_по_РОО_Тюменский + Размер_блока_офиса - 1)
      
      ' Показатели текущего периода
      For j = 1 To 12
        ThisWorkbook.Sheets("Лист8").Cells(i, j).Copy Destination:=Workbooks(FileDBName).Sheets("Лист1").Cells(i + 16, j)
      Next j
      
      ' Показатели прошлого периода и Динамики
      For j = 15 To 18
        ThisWorkbook.Sheets("Лист8").Cells(i, j).Copy Destination:=Workbooks(FileDBName).Sheets("Лист1").Cells(i + 16, j)
      Next j
      
      ' Строка статуса
      Application.StatusBar = "Копирование итогов РОО " + CStr(i) + "..."
      DoEvents
      
  Next i

  ' Копируем лист Capacity (Лист6)
  
  
  ' ***
                    
  ' Закрытие файла
  Workbooks(FileDBName).Close SaveChanges:=True

  ' Копирование завершено
  Application.StatusBar = "Скопировано!"
  Application.StatusBar = ""

End Sub

' Интегральный рейтинг по офисам
Sub DB_rating(In_ReportName_String, In_Sheets, In_Row_Лист8, In_Заголовок_столбца_офисы)
Dim dateDB As Date
        
  ' ***
  ' In_Заголовок_столбца_офисы - наименование заголовка на листе, под которым идут филиалы: Алтайский ОО1, Архангельский ОО1, Астраханский ОО1 ...
  ' ***
    
  ' Из A1 "Отчет по состоянию на 02.09.2020" берем дату, если ее нет на текущем листе, то берем с листа "Оглавление"
  If Workbooks(In_ReportName_String).Sheets(In_Sheets).Cells(1, 1).Value <> "" Then
    dateDB = CDate(Mid(Workbooks(In_ReportName_String).Sheets(In_Sheets).Cells(1, 1).Value, 23, 10))
  Else
    dateDB = CDate(Mid(Workbooks(In_ReportName_String).Sheets("Оглавление").Cells(1, 1).Value, 23, 10))
  End If
  
  ' Вкладка In_Sheets
  ' 42
  Row_Заголовок_столбца_офисы = rowByValue(In_ReportName_String, In_Sheets, In_Заголовок_столбца_офисы, 1000, 1000)
  ' 2
  Column_Заголовок_столбца_офисы = ColumnByValue(In_ReportName_String, In_Sheets, In_Заголовок_столбца_офисы, 1000, 1000)
  
  ' Находим в с столбце "Тюменский ОО1"
  rowCount = Row_Заголовок_столбца_офисы + 3
  Do While InStr(Workbooks(In_ReportName_String).Sheets(In_Sheets).Cells(rowCount, Column_Заголовок_столбца_офисы - 2).Value, "Общий итог") = 0
                
    ' Если это текущий офис
    If (InStr(Workbooks(In_ReportName_String).Sheets(In_Sheets).Cells(rowCount, Column_Заголовок_столбца_офисы).Value, "Тюменский") <> 0) Or _
          (InStr(Workbooks(In_ReportName_String).Sheets(In_Sheets).Cells(rowCount, Column_Заголовок_столбца_офисы).Value, "Сургутский") <> 0) Or _
            (InStr(Workbooks(In_ReportName_String).Sheets(In_Sheets).Cells(rowCount, Column_Заголовок_столбца_офисы).Value, "Нижневартовский") <> 0) Or _
              (InStr(Workbooks(In_ReportName_String).Sheets(In_Sheets).Cells(rowCount, Column_Заголовок_столбца_офисы).Value, "Новоуренгойский") <> 0) Or _
                (InStr(Workbooks(In_ReportName_String).Sheets(In_Sheets).Cells(rowCount, Column_Заголовок_столбца_офисы).Value, "Тарко-Сале") <> 0) Then
      
      ' Выводим показатели интегрального рейтинга в строку In_Row_Лист8
      ' Место
      Workbooks(In_ReportName_String).Sheets(In_Sheets).Cells(rowCount, Column_Заголовок_столбца_офисы - 3).Copy Destination:=ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 1)
      ' Убираем рамки
      ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 1).Borders(xlDiagonalDown).LineStyle = xlNone
      ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 1).Borders(xlDiagonalUp).LineStyle = xlNone
      ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 1).Borders(xlEdgeLeft).LineStyle = xlNone
      ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 1).Borders(xlEdgeTop).LineStyle = xlNone
      ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 1).Borders(xlEdgeBottom).LineStyle = xlNone
      ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 1).Borders(xlEdgeRight).LineStyle = xlNone
      ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 1).Borders(xlInsideVertical).LineStyle = xlNone
      ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 1).Borders(xlInsideHorizontal).LineStyle = xlNone
      ' Офис
      Workbooks(In_ReportName_String).Sheets(In_Sheets).Cells(rowCount, Column_Заголовок_столбца_офисы).Copy Destination:=ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 2)
      ' Показатели (11 шт.)
      ' Офис
      Workbooks(In_ReportName_String).Sheets(In_Sheets).Cells(rowCount, Column_Заголовок_столбца_офисы + 1).Copy Destination:=ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 3)
      Workbooks(In_ReportName_String).Sheets(In_Sheets).Cells(rowCount, Column_Заголовок_столбца_офисы + 2).Copy Destination:=ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 4)
      Workbooks(In_ReportName_String).Sheets(In_Sheets).Cells(rowCount, Column_Заголовок_столбца_офисы + 3).Copy Destination:=ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 5)
      Workbooks(In_ReportName_String).Sheets(In_Sheets).Cells(rowCount, Column_Заголовок_столбца_офисы + 4).Copy Destination:=ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 6)
      Workbooks(In_ReportName_String).Sheets(In_Sheets).Cells(rowCount, Column_Заголовок_столбца_офисы + 5).Copy Destination:=ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 7)
      Workbooks(In_ReportName_String).Sheets(In_Sheets).Cells(rowCount, Column_Заголовок_столбца_офисы + 6).Copy Destination:=ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 8)
      Workbooks(In_ReportName_String).Sheets(In_Sheets).Cells(rowCount, Column_Заголовок_столбца_офисы + 7).Copy Destination:=ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 9)
      Workbooks(In_ReportName_String).Sheets(In_Sheets).Cells(rowCount, Column_Заголовок_столбца_офисы + 8).Copy Destination:=ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 10)
      Workbooks(In_ReportName_String).Sheets(In_Sheets).Cells(rowCount, Column_Заголовок_столбца_офисы + 9).Copy Destination:=ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 11)
      Workbooks(In_ReportName_String).Sheets(In_Sheets).Cells(rowCount, Column_Заголовок_столбца_офисы + 10).Copy Destination:=ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 12)
      
      ' Интегральный рейтинг офиса, который должен быть не менее 90%
      Workbooks(In_ReportName_String).Sheets(In_Sheets).Cells(rowCount, Column_Заголовок_столбца_офисы + 11).Copy Destination:=ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 13)

      ' Всставляем в показатели продаж BASE\Sales_Office
      In_officeNameInReport = Workbooks(In_ReportName_String).Sheets(In_Sheets).Cells(rowCount, Column_Заголовок_столбца_офисы).Value
      In_Product_Name = "Интегральный рейтинг"
      In_Product_Code = "ИнтРейт"
      Факт_ИР_Офиса = Round(ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 13) * 100, 2)
      
      '  Идентификатор ID_Rec:
      ID_RecVar = CStr(CStr(getNumberOfficeByName2(In_officeNameInReport)) + "-" + strMMYY(dateDB) + "-" + In_Product_Code)
            
      ' Текущие значения в месяце: Date_01 (N), Date_02 (O), Date_03 (P), Date_04 (Q), Date_05 Date_06 Date_07 Date_08 Date_09 Date_10 Date_11 Date_12 Date_13 Date_14 Date_15 Date_16 Date_17 Date_18 Date_19 Date_20 Date_21 Date_22 Date_23 Date_24 Date_25 Date_26 Date_27 Date_28 Date_29 Date_30 Date_31
      curr_Day_Month = "Date_" + Mid(dateDB, 1, 2)
            
      ' Вносим данные в BASE\Sales_Office по ПК.
      Call InsertRecordInBook("Sales_Office", "Лист1", "ID_Rec", ID_RecVar, _
                                            "ID_Rec", ID_RecVar, _
                                              "Оffice_Number", getNumberOfficeByName2(In_officeNameInReport), _
                                                "Product_Name", In_Product_Name, _
                                                  "Оffice", getShortNameOfficeByName(In_officeNameInReport), _
                                                    "MMYY", strMMYY(dateDB), _
                                                      "Update_Date", dateDB, _
                                                        "Product_Code", In_Product_Code, _
                                                          "Plan", "90", _
                                                             "Unit", "%", _
                                                               "Fact", Факт_ИР_Офиса, _
                                                                 "Percent_Completion", "", _
                                                                   "Prediction", "", _
                                                                     "Percent_Prediction", "", _
                                                                       curr_Day_Month, Факт_ИР_Офиса, _
                                                                         "", "", _
                                                                           "", "", _
                                                                             "", "", _
                                                                               "", "", _
                                                                                 "", "", _
                                                                                   "", "")


      ' Увеличиваем значение вывода следующей строки
      In_Row_Лист8 = In_Row_Лист8 + 1
      
    End If
    
    ' Следующая запись
    Application.StatusBar = "Интегральный рейтинг по офисам " + CStr(rowCount) + "..."
    rowCount = rowCount + 1
    DoEventsInterval (rowCount)
    
  Loop
  
  
End Sub


' Отправка письма: отправляю шаблон самому себе для последующей отправки в сеть письма на его основе:
Sub Отправка_Lotus_Notes_Лист8()
Dim темаПисьма, текстПисьма, hashTag, attachmentFile As String
Dim i As Byte
  
  ' Строка статуса
  Application.StatusBar = "Отправка письма с фокусами контроля..."
  
  ' Период контроля показателей
  If ThisWorkbook.Sheets("Лист8").Range("N7").Value = 1 Then
    ПериодКонтроля = "Месяц"
  Else
    ПериодКонтроля = "Квартал"
  End If
  
  ' Запрос
  ' If MsgBox("Отправить себе Шаблон письма с фокусами контроля '" + ПериодКонтроля + "'?", vbYesNo) = vbYes Then
    
    ' Тема письма - Тема:
    ' темаПисьма = ThisWorkbook.Sheets("Лист8").Cells(RowByValue(ThisWorkbook.Name, "Лист8", "Тема:", 100, 100), ColumnByValue(ThisWorkbook.Name, "Лист8", "Тема:", 100, 100) + 1).Value
    темаПисьма = subjectFromSheet("Лист8")

    ' hashTag - Хэштэг:
    ' hashTag = ThisWorkbook.Sheets("Лист8").Cells(RowByValue(ThisWorkbook.Name, "Лист8", "Хэштэг:", 100, 100), ColumnByValue(ThisWorkbook.Name, "Лист8", "Хэштэг:", 100, 100) + 1).Value
    ' hashTag - Хэштэг:
    hashTag = hashTagFromSheet("Лист8")

    ' Файл-вложение (!!!)
    attachmentFile = ThisWorkbook.Sheets("Лист8").Cells(3, 17).Value
    
    ' Текст письма
    текстПисьма = "" + Chr(13)
    текстПисьма = текстПисьма + "" + ThisWorkbook.Sheets("Лист8").Cells(rowByValue(ThisWorkbook.Name, "Лист8", "Список получателей:", 100, 100), ColumnByValue(ThisWorkbook.Name, "Лист8", "Список получателей:", 100, 100) + 2).Value + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "" + getFromAddrBook("РД", 2) + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "Уважаемые сотрудники," + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "Ежедневный отчет по продажам в разрезе офисов и сотрудников (файл во вложении)." + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "Продукты, входящие в интегральный рейтинг, с прогнозом исполнения плана " + Квартал_месяц_план + " менее 100% на " + Завершающий_день_квартал_месяц(DashboardDate()) + " г.:" + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + Фокусы_контроля() + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "Продажи по сотрудникам, с прогнозом исполнения плана месяца менее 90% на " + CStr(Date_last_day_month(DashboardDate())) + " г.:"
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + Фокусы_контроля2() + Chr(13)
    ' Визитка (подпись С Ув., )
    текстПисьма = текстПисьма + ПодписьВПисьме()
    ' Хэштег
    текстПисьма = текстПисьма + createBlankStr(27) + hashTag
    ' Вызов
    Call send_Lotus_Notes(темаПисьма, "Sergey Fedorovich Proschaev/Tyumen/PSBank/Ru", "Sergey Fedorovich Proschaev/Tyumen/PSBank/Ru", текстПисьма, attachmentFile)
  
    ' Зачеркнуть
    Call ЗачеркиваемТекстВячейке("Лист0", RangeByValue(ThisWorkbook.Name, "Лист0", "DashBoard (при наличии)", 100, 100))
  
    ' Сообщение
    ' MsgBox ("Письмо отправлено!")
     
    ' Строка статуса
    Application.StatusBar = ""
     
  ' End If
  
End Sub


' Переключение выборки в DB Месяц/Квартал на листе
Sub DB_swith_to_MonthQuarter(In_ReportName_String, In_Sheets, In_Period, In_Срез_Период)
        
        ' ***
        ' In_Period = 1 - месяц, In_Period = 2 - квартал
        ' In_Срез_Период = Срез_Период5 - для Инвестов, In_Срез_Период = Срез_Период3 - для ОФЗ
        ' ***
                
        ' Переключение Месяц/Квартал -> Месяц
        If In_Period = 1 Then

          ' Месяц
          Select Case In_Срез_Период
          
          Case "Срез_Период3" ' ОФЗ
            
            Workbooks(In_ReportName_String).SlicerCaches(In_Срез_Период).SlicerItems("Месяц").Selected = True
            Workbooks(In_ReportName_String).SlicerCaches(In_Срез_Период).SlicerItems("Квартал").Selected = False
            
          Case "Срез_Период5" ' Инвест
        
            ' Первой идет то, что True - иначе неверно выводит данные
            Workbooks(In_ReportName_String).SlicerCaches(In_Срез_Период).SlicerItems("Месяц").Selected = True
            Workbooks(In_ReportName_String).SlicerCaches(In_Срез_Период).SlicerItems("Квартал").Selected = False
            Workbooks(In_ReportName_String).SlicerCaches(In_Срез_Период).SlicerItems("(пусто)").Selected = False
        
          End Select

          
        
        End If
        
        ' Переключение Месяц/Квартал -> Квартал
        If In_Period = 2 Then
          
          ' Квартал
          Select Case In_Срез_Период
          
          Case "Срез_Период3" ' ОФЗ
            
            ' Первой идет то, что True - иначе неверно выводит данные
            Workbooks(In_ReportName_String).SlicerCaches(In_Срез_Период).SlicerItems("Квартал").Selected = True
            Workbooks(In_ReportName_String).SlicerCaches(In_Срез_Период).SlicerItems("Месяц").Selected = False
            
          Case "Срез_Период5" ' Инвест
        
            Workbooks(In_ReportName_String).SlicerCaches(In_Срез_Период).SlicerItems("Квартал").Selected = True
            Workbooks(In_ReportName_String).SlicerCaches(In_Срез_Период).SlicerItems("Месяц").Selected = False
            Workbooks(In_ReportName_String).SlicerCaches(In_Срез_Период).SlicerItems("(пусто)").Selected = False
        
          End Select

        
        End If
        
               
    ' With ActiveWorkbook.SlicerCaches("Срез_Период3")
    '     .SlicerItems("Месяц").Selected = True
    '     .SlicerItems("Квартал").Selected = False
    ' End With
    
    ' With ActiveWorkbook.SlicerCaches("Срез_Период3")
    '     .SlicerItems("Квартал").Selected = True
    '     .SlicerItems("Месяц").Selected = False
    ' End With

        
End Sub

' С Листа8 - все, что в прогнозе менее 100%
Function Фокусы_контроля() As String
Dim rowCount As Integer
 
  ' Минимальный норматив прогноза или факта
  Мин_норматив = 1 ' 0.9
 
  ' Обработка DB
  Фокусы_контроля = ""
  Строка_к_выводу = ""
  
  ' Если контроль месячный N7="1"
  If ThisWorkbook.Sheets("Лист8").Range("N7").Value = 1 Then
    ' Месяц
    Столбец_прогноза = 12
  Else
    ' Квартал
    Столбец_прогноза = 8
  End If
  
  
  ' 1. Отдел корпоративных продаж
  Строка_к_выводу = Строка_к_выводу + "1. Отдел корпоративных продаж: "
  
  ' Зарплатные карты 18+
  row_Зарплатные_карты_18 = getRowFromSheet8("Итого по РОО «Тюменский»", "Зарплатные карты 18+")
  Строка_к_выводу = Строка_к_выводу + Сокр(ThisWorkbook.Sheets("Лист8").Cells(row_Зарплатные_карты_18, 2).Value) + " (Факт " + CStr(Round(ThisWorkbook.Sheets("Лист8").Cells(row_Зарплатные_карты_18, 7).Value * 100, 0)) + "%, Прогноз " + CStr(Round(ThisWorkbook.Sheets("Лист8").Cells(row_Зарплатные_карты_18, 20).Value * 100, 0)) + "%), "
  
  ' Портфель ЗП 18+, шт._Квартал
  row_Портфель_ЗП_18_шт_Квартал = getRowFromSheet8("Итого по РОО «Тюменский»", "Портфель ЗП 18+, шт._Квартал ")
  Строка_к_выводу = Строка_к_выводу + Сокр(ThisWorkbook.Sheets("Лист8").Cells(row_Портфель_ЗП_18_шт_Квартал, 2).Value) + " (" + CStr(Round(ThisWorkbook.Sheets("Лист8").Cells(row_Портфель_ЗП_18_шт_Квартал, 7).Value * 100, 0)) + "%), "
  
  '            КК к ЗП
  row_КК_к_ЗП = getRowFromSheet8("Итого по РОО «Тюменский»", "           КК к ЗП")
  Строка_к_выводу = Строка_к_выводу + Сокр(ThisWorkbook.Sheets("Лист8").Cells(row_КК_к_ЗП, 2).Value) + " (" + CStr(Round(ThisWorkbook.Sheets("Лист8").Cells(row_КК_к_ЗП, 20).Value * 100, 0)) + "%)" + Chr(13) + Chr(13)
    
  ' 2. Канал ПВО
  Строка_к_выводу = Строка_к_выводу + "2. Канал ПВО: "
  
  ' в т.ч. ПК DSA
  row_ПК_DSA = getRowFromSheet8("Итого по РОО «Тюменский»", "в т.ч. ПК DSA")
  If ThisWorkbook.Sheets("Лист8").Cells(row_ПК_DSA, Столбец_прогноза).Value < Мин_норматив Then
    Строка_к_выводу = Строка_к_выводу + Сокр(ThisWorkbook.Sheets("Лист8").Cells(row_ПК_DSA, 2).Value) + " (" + CStr(Round(ThisWorkbook.Sheets("Лист8").Cells(row_ПК_DSA, Столбец_прогноза).Value * 100, 0)) + "%), "
  End If
  
  '            КК DSA
  row_КК_DSA = getRowFromSheet8("Итого по РОО «Тюменский»", "           КК DSA")
  If ThisWorkbook.Sheets("Лист8").Cells(row_КК_DSA, Столбец_прогноза).Value < Мин_норматив Then
    Строка_к_выводу = Строка_к_выводу + Сокр(ThisWorkbook.Sheets("Лист8").Cells(row_КК_DSA, 2).Value) + " (" + CStr(Round(ThisWorkbook.Sheets("Лист8").Cells(row_КК_DSA, Столбец_прогноза).Value * 100, 0)) + "%)"
  End If
  
  Строка_к_выводу = Строка_к_выводу + Chr(13) + Chr(13)
  
  ' 3. Офисы. Обработка текущих параметров
  Строка_к_выводу = Строка_к_выводу + "3. Офисы: "
  
  rowCount = rowByValue(ThisWorkbook.Name, "Лист8", "Тюменский РОО", 100, 100) + 2
  Do While (InStr(ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value, "Интегральный рейтинг по офисам") = 0)
    
    ' Если начинается раздел офиса
    If (InStr(ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value, "Тюменский") <> 0) Or (InStr(ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value, "Сургутский") <> 0) Or (InStr(ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value, "Нижневартовский") Or (InStr(ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value, "Новоуренгойский")) Or (InStr(ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value, "Тарко-Сале") <> 0)) Then
      
      Фокусы_контроля = Фокусы_контроля + Строка_к_выводу + Chr(13) + Chr(13)
      Строка_к_выводу = "- " + cityOfficeName(ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value) + ": "
    
    End If
    
    ' Если контроль месячный N7="1"
    If ThisWorkbook.Sheets("Лист8").Range("N7").Value = 1 Then
      ' Месяц
      Столбец_прогноза = 12
    Else
      ' Квартал
      Столбец_прогноза = 8
    End If
    
    
    ' Набираем показатели контроля в строку по кварталу
    ' 100%
    ' If (ThisWorkbook.Sheets("Лист8").Cells(rowCount, 8).Value < 1) And (ThisWorkbook.Sheets("Лист8").Cells(rowCount, 3).Value <> "") Then
    ' 90%
    ' If ((ThisWorkbook.Sheets("Лист8").Cells(rowCount, 8).Value < 0.9) And (ThisWorkbook.Sheets("Лист8").Cells(rowCount, 3).Value <> "")) Then
    If ((ThisWorkbook.Sheets("Лист8").Cells(rowCount, Столбец_прогноза).Value < Мин_норматив) And (ThisWorkbook.Sheets("Лист8").Cells(rowCount, 3).Value <> "")) Then
      
      Строка_к_выводу = Строка_к_выводу + Сокр(ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value) + " (" + CStr(Round(ThisWorkbook.Sheets("Лист8").Cells(rowCount, Столбец_прогноза).Value * 100, 0)) + "%), "
    
    End If
    
    ' Пассивы
    If (ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value = "Пассивы") And ((ThisWorkbook.Sheets("Лист8").Cells(rowCount, 7).Value < 0.9)) Then

      Строка_к_выводу = Строка_к_выводу + Сокр(ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value) + " (" + CStr(Round(ThisWorkbook.Sheets("Лист8").Cells(rowCount, 7).Value * 100, 0)) + "%), "
      
    End If
    
    
    ' Вывод исп плана по коробкам в месяц
    ' Набираем показатели контроля в строку по кварталу
    ' If (ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value = "Коробки+Личный адвокат") And ((ThisWorkbook.Sheets("Лист8").Cells(rowCount, 12).Value < 0.9)) Then
    If (ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value = "Коробки+Личный адвокат (премия)") And ((ThisWorkbook.Sheets("Лист8").Cells(rowCount, Столбец_прогноза).Value < Мин_норматив)) Then

      Строка_к_выводу = Строка_к_выводу + Сокр(ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value) + " (" + CStr(Round(ThisWorkbook.Sheets("Лист8").Cells(rowCount, Столбец_прогноза).Value * 100, 0)) + "%), "
      
    End If
    
    ' ИСЖ
    If (ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value = "Премия ИСЖ МАСС") And (ThisWorkbook.Sheets("Лист8").Cells(rowCount, Столбец_прогноза).Value < Мин_норматив) Then
          
      Строка_к_выводу = Строка_к_выводу + Сокр(ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value) + " (" + CStr(Round(ThisWorkbook.Sheets("Лист8").Cells(rowCount, Столбец_прогноза).Value * 100, 0)) + "%), "
      
    End If
    
    ' НСЖ
    If (ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value = "Премия НСЖ МАСС") And (ThisWorkbook.Sheets("Лист8").Cells(rowCount, Столбец_прогноза).Value < Мин_норматив) Then
          
      Строка_к_выводу = Строка_к_выводу + Сокр(ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value) + " (" + CStr(Round(ThisWorkbook.Sheets("Лист8").Cells(rowCount, Столбец_прогноза).Value * 100, 0)) + "%), "
      
    End If
    
    
    ' Инвесты
    ' Набираем показатели контроля в строку по кварталу
    If (ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value = "Инвест") And (ThisWorkbook.Sheets("Лист8").Cells(rowCount, Столбец_прогноза).Value < Мин_норматив) Then
          
      Строка_к_выводу = Строка_к_выводу + Сокр(ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value) + " (" + CStr(Round(ThisWorkbook.Sheets("Лист8").Cells(rowCount, Столбец_прогноза).Value * 100, 0)) + "%), "
      
    End If
    
    ' Инвест Брокер обслуж
    If (ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value = "Инвест Брокер обслуж") And (ThisWorkbook.Sheets("Лист8").Cells(rowCount, Столбец_прогноза).Value < Мин_норматив) Then
          
      Строка_к_выводу = Строка_к_выводу + Сокр(ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value) + " (" + CStr(Round(ThisWorkbook.Sheets("Лист8").Cells(rowCount, Столбец_прогноза).Value * 100, 0)) + "%), "
      
    End If
    
    ' Следующая запись
    Application.StatusBar = "Анализ прогнозов " + CStr(rowCount) + "..."
    rowCount = rowCount + 1
    DoEventsInterval (rowCount)
  
  Loop
  
  ' Итоги
  Фокусы_контроля = Фокусы_контроля + Строка_к_выводу + Chr(13)
  
  Application.StatusBar = ""
  
  
End Function

' Сокращение для отправки в почте
Function Сокр(In_Str) As String
  
  Сокр = In_Str
 
        Select Case In_Str
          
          Case "Зарплатные карты 18+"
            Сокр = "ЗП карты"
          Case "Портфель ЗП 18+, шт._Квартал "
            Сокр = "Портфель ЗП"
          Case "           КК к ЗП"
            Сокр = "КК к ЗП"
          Case "в т.ч. ПК DSA"
            Сокр = "ПК DSA"
          Case "           КК DSA"
            Сокр = "КК DSA"
          Case "           КК к Ипотеке"
            Сокр = "КК к Ипотеке"
          Case "Потребительские кредиты"
            Сокр = "ПК"
          Case "Кредитные карты (актив.)"
            Сокр = "КК"
          Case "Дебетовые карты (актив.)"
            Сокр = "ДК"
          Case "Интернет-банк"
            Сокр = "ИБ"
          Case "Orange Premium Club"
            Сокр = "OPC"
          Case "Комиссионный доход"
            Сокр = "Ком.доход"
          Case "Коробки+Личный адвокат"
            Сокр = "Коробки"
          Case "           КЛА премия ИЦ"
            Сокр = "КСП ЛА ИЦ"
          Case "Накопительные счета"
            Сокр = "НС"
          Case "Зарплатные карты 18+"
            Сокр = "Зарпл.карты"
          Case "Инвест Брокер обслуж"
            Сокр = "Брокер.счета"
          Case "           КК OPC"
            Сокр = "КК OPC"
          Case "Инвест Брокер обслуж OPC"
            Сокр = "Брокер.счета OPC"
            
        End Select
  

End Function

' Сокращение для отправки в почте
Function Сокр2(In_Str) As String
  
  Сокр2 = In_Str
 
        Select Case In_Str
          Case "Потребительские кредиты"
            Сокр2 = "Потреб.кредиты"
          Case "Кредитные карты (актив.)"
            Сокр2 = "Кред.карты"
          Case "Дебетовые карты (актив.)"
            Сокр2 = "ДК"
          Case "Интернет-банк"
            Сокр2 = "ИБ"
          Case "Orange Premium Club"
            Сокр2 = "OPC"
          Case "Комиссионный доход"
            Сокр2 = "Ком.доход"
          Case "Коробки+Личный адвокат"
            Сокр2 = "Коробки"
          Case "Накопительные счета"
            Сокр2 = "НС"
          Case "Зарплатные карты 18+"
            Сокр2 = "Зарпл.карты"
          Case "Инвест Брокер обслуж"
            Сокр2 = "Брок.счета"
          Case "Портфель ЗП 18+, шт._Квартал "
            Сокр2 = "Портф.ЗП"
            
        End Select
  

End Function


' Переход к началу Листа на ПК
Sub Лист8_к_началу()
  ThisWorkbook.Sheets("Лист8").Range("M8").Select
End Sub

' Переход к Сургуту на Листе8
Sub Лист8_к_Сургуту()
  
  ' Делаем расчет блока
  row_ОО_Сургутский = getRowFromSheet8("ОО «Сургутский»", "ОО «Сургутский»")
  
  ' Перемещаемся
  ActiveWindow.SmallScroll Down:=row_ОО_Сургутский - 1
  
End Sub

' Переход к Нижневартовску на Листе8
Sub Лист8_к_Нижневартовску()
  row_ОО_Нижневартовский = getRowFromSheet8("ОО «Нижневартовский»", "ОО «Нижневартовский»")
  ' Перемещаемся
  ActiveWindow.SmallScroll Down:=row_ОО_Нижневартовский - 1
End Sub

' Переход к Новый Уренгой на Листе8
Sub Лист8_к_Новый_Уренгой()
  row_ОО_Новоуренгойский = getRowFromSheet8("ОО «Новоуренгойский»", "ОО «Новоуренгойский»")
  ActiveWindow.SmallScroll Down:=row_ОО_Новоуренгойский - 1
End Sub

' Переход к Тарко-Сале на Листе8
Sub Лист8_к_Тарко_Сале()
  row_ОО_Тарко_Сале = getRowFromSheet8("ОО «Тарко-Сале»", "ОО «Тарко-Сале»")
  ' Перемещаемся
  ActiveWindow.SmallScroll Down:=row_ОО_Тарко_Сале - 1
End Sub

' Переход к РОО на Листе8
Sub Лист8_к_РОО_Тюменский()
  row_Итого_по_РОО_Тюменский = getRowFromSheet8("Итого по РОО «Тюменский»", "Итого по РОО «Тюменский»")
  ' Перемещаемся
  ActiveWindow.SmallScroll Down:=row_Итого_по_РОО_Тюменский - 1
End Sub


' Дата с листа Дашбоарда c Листа 8
Function DashboardDate() As Date
  
  ' Тюменский РОО
  rowVar = rowByValue(ThisWorkbook.Name, "Лист8", "Тюменский РОО", 1000, 1000)
  columnVar = ColumnByValue(ThisWorkbook.Name, "Лист8", "Тюменский РОО", 1000, 1000)
  
  DashboardDate = CDate(Mid(ThisWorkbook.Sheets("Лист8").Cells(rowVar + 1, columnVar).Value, 52, 10))
  
End Function

' С Листа7 - все, что в прогнозе менее 100%
Function Фокусы_контроля2() As String
Dim rowCount As Byte
 
  ' Столбец "Прим."
  Column_Прим = ColumnByValue(ThisWorkbook.Name, "Лист7", "Прим.", 100, 100)
 
  ' Обработка DB
  Фокусы_контроля2 = ""
  Строка_к_выводу = ""
  rowCount = rowByValue(ThisWorkbook.Name, "Лист7", "Тюменский РОО", 100, 100) + 5
  Do While Not IsEmpty(ThisWorkbook.Sheets("Лист7").Cells(rowCount, 1).Value)
    
    ' Если в столбце "Прим." пусто
    If IsEmpty(ThisWorkbook.Sheets("Лист7").Cells(rowCount, Column_Прим).Value) Then
    
      ' ФИО
      Строка_к_выводу = Строка_к_выводу + " - " + ThisWorkbook.Sheets("Лист7").Cells(rowCount, 3).Value + ": "
    
      ' Потреб кредитование
      If ThisWorkbook.Sheets("Лист7").Cells(rowCount, 9).Value < 0.9 Then
        Строка_к_выводу = Строка_к_выводу + Сокр("Потребительские кредиты") + " (" + CStr(Round(ThisWorkbook.Sheets("Лист7").Cells(rowCount, 9).Value * 100, 0)) + "%), "
      End If
      
      ' БС
      If ThisWorkbook.Sheets("Лист7").Cells(rowCount, 14).Value < 0.9 Then
        Строка_к_выводу = Строка_к_выводу + Сокр("СЖиЗ к ПК") + " (" + CStr(Round(ThisWorkbook.Sheets("Лист7").Cells(rowCount, 14).Value * 100, 0)) + "%), "
      End If
      
      ' КК
      If ThisWorkbook.Sheets("Лист7").Cells(rowCount, 19).Value < 0.9 Then
        Строка_к_выводу = Строка_к_выводу + Сокр("Кредитные карты (актив.)") + " (" + CStr(Round(ThisWorkbook.Sheets("Лист7").Cells(rowCount, 19).Value * 100, 0)) + "%), "
      End If
      
      ' ДК
      If ThisWorkbook.Sheets("Лист7").Cells(rowCount, 24).Value < 0.9 Then
        Строка_к_выводу = Строка_к_выводу + Сокр("Дебетовые карты (актив.)") + " (" + CStr(Round(ThisWorkbook.Sheets("Лист7").Cells(rowCount, 24).Value * 100, 0)) + "%), "
      End If
      
      ' Интернет-банк
      If ThisWorkbook.Sheets("Лист7").Cells(rowCount, 29).Value < 0.9 Then
        Строка_к_выводу = Строка_к_выводу + Сокр("Интернет-банк") + " (" + CStr(Round(ThisWorkbook.Sheets("Лист7").Cells(rowCount, 29).Value * 100, 0)) + "%), "
      End If
      
      ' Накопительные счета
      ' If ThisWorkbook.Sheets("Лист7").Cells(rowCount, 34).Value < 0.9 Then
      '   Строка_к_выводу = Строка_к_выводу + Сокр("Накопительные счета") + " (" + CStr(Round(ThisWorkbook.Sheets("Лист7").Cells(rowCount, 34).Value * 100, 0)) + "%), "
      ' End If
      
      ' ИСЖ_МАСС (Премия, тыс.руб.)
      If ThisWorkbook.Sheets("Лист7").Cells(rowCount, 39).Value < 0.9 Then
        Строка_к_выводу = Строка_к_выводу + Сокр("ИСЖ") + " (" + CStr(Round(ThisWorkbook.Sheets("Лист7").Cells(rowCount, 39).Value * 100, 0)) + "%), "
      End If
      
      ' НСЖ_МАСС (комиссионный доход)
      If ThisWorkbook.Sheets("Лист7").Cells(rowCount, 44).Value < 0.9 Then
        Строка_к_выводу = Строка_к_выводу + Сокр("НСЖ") + " (" + CStr(Round(ThisWorkbook.Sheets("Лист7").Cells(rowCount, 44).Value * 100, 0)) + "%), "
      End If
      
      ' Коробочное страхование
      If ThisWorkbook.Sheets("Лист7").Cells(rowCount, 49).Value < 0.9 Then
        Строка_к_выводу = Строка_к_выводу + Сокр("КС") + " (" + CStr(Round(ThisWorkbook.Sheets("Лист7").Cells(rowCount, 49).Value * 100, 0)) + "%), "
      End If
      
      Строка_к_выводу = Строка_к_выводу + Chr(13) + Chr(13)
      
    End If ' Если в столбце "Прим." пусто
      
    ' Следующая запись
    Application.StatusBar = "Анализ прогнозов 2" + CStr(rowCount) + "..."
    rowCount = rowCount + 1
    DoEventsInterval (rowCount)
  
  Loop
  
  ' Итоги
  Фокусы_контроля2 = Фокусы_контроля2 + Строка_к_выводу + Chr(13)
  
  Application.StatusBar = ""
  
  
End Function

' Загрузить планы из файла "Декомпозиция планов продаж_4кв.2020_Сеть"
Sub Загрузить_Декомпозицию_планов()
Dim row_Наименование As Integer

  ' Запрос на исполнение процедуры
  If MsgBox("Загрузить Декомпозицию планов продаж?", vbYesNo) = vbYes Then
    
    ' Открыть файл с отчетом
    FileName = Application.GetOpenFilename("Excel Files (*.xlsx), *.xlsx", , "Открытие файла с Декомпозицией")

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
      ThisWorkbook.Sheets("Лист8").Activate

      ' Проверка формы отчета
      CheckFormatReportResult = CheckFormatReport(ReportName_String, "ПК", 10, Date)
    
      If CheckFormatReportResult = "OK" Then
      
      
        ' Очищаем квартальные показатели
        
      
        ' Обрабатываем отчет
        ' Цикл по 5-ти офисам
        ' Обработка отчета
        For i = 1 To 5
          ' Номера офисов от 1 до 5
          Select Case i
            Case 1 ' ОО «Тюменский»
              officeNameInReport = "Тюменский"
              row_Наименование = getRowFromSheet8("Тюменский", "Тюменский") ' 9
            Case 2 ' ОО «Сургутский»
              officeNameInReport = "Сургутский"
              row_Наименование = getRowFromSheet8("Сургутский", "Сургутский") ' 47
            Case 3 ' ОО «Нижневартовский»
              officeNameInReport = "Нижневартовский"
              row_Наименование = getRowFromSheet8("Нижневартовский", "Нижневартовский") ' 85
            Case 4 ' ОО «Новоуренгойский»
              officeNameInReport = "Новоуренгойский"
              row_Наименование = getRowFromSheet8("Новоуренгойский", "Новоуренгойский") ' 123
            Case 5 ' ОО «Тарко-Сале»
              officeNameInReport = "Тарко-Сале"
              row_Наименование = getRowFromSheet8("Тарко-Сале", "Тарко-Сале") ' 161
          End Select

          ' *** Циклы обработки показателей ИР ***
          ' 1) 1. Зарплатные карты 18+
          Call getDataFromDecompositionPlans(officeNameInReport, _
                                               ReportName_String, _
                                                 "ЗП_карты", _
                                                   4, _
                                                     5, _
                                                       6, _
                                                         7, _
                                                           8, _
                                                             row_Наименование + 1)
          
          ' 2) 2. Потребительские кредиты
          Call getDataFromDecompositionPlans(officeNameInReport, _
                                               ReportName_String, _
                                                 "ПК", _
                                                   4, _
                                                     5, _
                                                       6, _
                                                         7, _
                                                           8, _
                                                             row_Наименование + 2)
          
          ' 3) 3. Кредитные карты (актив.)
          Call getDataFromDecompositionPlans(officeNameInReport, _
                                               ReportName_String, _
                                                 "КК", _
                                                   4, _
                                                     5, _
                                                       6, _
                                                         7, _
                                                           8, _
                                                             row_Наименование + 3)
          
          ' 4) 4. Дебетовые карты (актив.)
          Call getDataFromDecompositionPlans(officeNameInReport, _
                                               ReportName_String, _
                                                 "ДК", _
                                                   4, _
                                                     5, _
                                                       6, _
                                                         7, _
                                                           8, _
                                                             row_Наименование + 4)
          
          ' 5) 5. Интернет -банк
          Call getDataFromDecompositionPlans(officeNameInReport, _
                                               ReportName_String, _
                                                 "ИБ", _
                                                   4, _
                                                     5, _
                                                       6, _
                                                         7, _
                                                           8, _
                                                             row_Наименование + 5)
          
          ' 6) 7. Orange Premium Club
          Call getDataFromDecompositionPlans(officeNameInReport, _
                                               ReportName_String, _
                                                 "OPC", _
                                                   4, _
                                                     5, _
                                                       6, _
                                                         7, _
                                                           8, _
                                                             row_Наименование + 7)
                                                             
          ' 7) 8. Кредитные карты Affluent
          Call getDataFromDecompositionPlans(officeNameInReport, _
                                               ReportName_String, _
                                                 "OPC", _
                                                   4, _
                                                     20, _
                                                       21, _
                                                         22, _
                                                           23, _
                                                             row_Наименование + 8)
                                                   
          
          
          ' 9)  14.1    в т.ч. страховки к ПК
          Call getDataFromDecompositionPlans(officeNameInReport, _
                                               ReportName_String, _
                                                 "Ком доход", _
                                                   4, _
                                                     36, _
                                                       37, _
                                                         38, _
                                                           39, _
                                                             row_Наименование + 15)
          
          
          ' 10)  14.2               ИСЖ_МАСС
          Call getDataFromDecompositionPlans(officeNameInReport, _
                                               ReportName_String, _
                                                 "Ком доход", _
                                                   4, _
                                                     6, _
                                                       7, _
                                                         8, _
                                                           9, _
                                                             row_Наименование + 16)
          
          ' 11) 14.3               НСЖ_МАСС
          Call getDataFromDecompositionPlans(officeNameInReport, _
                                               ReportName_String, _
                                                 "Ком доход", _
                                                   4, _
                                                     41, _
                                                       42, _
                                                         43, _
                                                           44, _
                                                             row_Наименование + 17)
          
          
          ' 12) 14.4               КС
          Call getDataFromDecompositionPlans(officeNameInReport, _
                                               ReportName_String, _
                                                 "Ком доход", _
                                                   4, _
                                                     16, _
                                                       17, _
                                                         18, _
                                                           19, _
                                                             row_Наименование + 18)
          
          ' 13) 14.5               ЛА
          Call getDataFromDecompositionPlans(officeNameInReport, _
                                               ReportName_String, _
                                                 "Ком доход", _
                                                   4, _
                                                     26, _
                                                       27, _
                                                         28, _
                                                           29, _
                                                             row_Наименование + 19)
          
          ' 14) 14.6               ИСЖ, НСЖ, КС (Affluent)
          Call getDataFromDecompositionPlans(officeNameInReport, _
                                               ReportName_String, _
                                                 "OPC", _
                                                   4, _
                                                     15, _
                                                       16, _
                                                         17, _
                                                           18, _
                                                             row_Наименование + 20)
          
          
          ' 15) 14.7               УК (Affluent)
          Call getDataFromDecompositionPlans(officeNameInReport, _
                                               ReportName_String, _
                                                 "OPC", _
                                                   4, _
                                                     10, _
                                                       11, _
                                                         12, _
                                                           13, _
                                                             row_Наименование + 21)
          
          
          
          ' 8) 14. Комиссионный доход (Сумма)
          commissionIncomeVar = ThisWorkbook.Sheets("Лист8").Cells(row_Наименование + 15, 5).Value _
                                  + ThisWorkbook.Sheets("Лист8").Cells(row_Наименование + 16, 5).Value _
                                    + ThisWorkbook.Sheets("Лист8").Cells(row_Наименование + 17, 5).Value _
                                      + ThisWorkbook.Sheets("Лист8").Cells(row_Наименование + 18, 5).Value _
                                        + ThisWorkbook.Sheets("Лист8").Cells(row_Наименование + 19, 5).Value _
                                          + ThisWorkbook.Sheets("Лист8").Cells(row_Наименование + 20, 5).Value _
                                            + ThisWorkbook.Sheets("Лист8").Cells(row_Наименование + 21, 5).Value
          
          ThisWorkbook.Sheets("Лист8").Cells(row_Наименование + 14, 5).Value = commissionIncomeVar
           
          ' 16) 16. Инвест
          Call getDataFromDecompositionPlans(officeNameInReport, _
                                               ReportName_String, _
                                                 "ИНВЕСТ", _
                                                   4, _
                                                     5, _
                                                       6, _
                                                         7, _
                                                           8, _
                                                             row_Наименование + 27)

          
          
          ' *** (конец) Циклы обработки показателей ИР ***
            
   
          ' Выводим данные по офису
      
        Next i ' Следующий офис
            
        ' Выводим итоги обработки
      
        ' Переменная завершения обработки
        finishProcess = True
      Else
        ' Сообщение о неверном формате отчета или даты
        MsgBox ("Проверьте отчет: " + CheckFormatReportResult + "!")
      End If ' Проверка формы отчета

      ' Закрываем файл с отчетом без сохранения изменений (параметр SaveChanges:=False)
      Workbooks(Dir(FileName)).Close SaveChanges:=False
    
      ' Переходим в ячейку M2
      ThisWorkbook.Sheets("Лист8").Range("A1").Select

      ' Строка статуса
      Application.StatusBar = ""
    
      ' Сохранение изменений
      ThisWorkbook.Save
          
      ' Итоговое сообщение
      If finishProcess = True Then
        MsgBox ("Обработка " + Dir(ReportName_String) + " завершена!")
      Else
        MsgBox ("Обработка отчета была прервана!")
      End If

    End If ' Если файл был выбран
 
  End If
  
End Sub

' Поручение данных из файла Декомпозиция планов
Sub getDataFromDecompositionPlans(In_officeNameInReport, In_ReportName_String, In_Sheets, In_ColOffice, In_Col1, In_Col2, In_Col3, In_Col4, In_rowSheet8)

  ' Офис был найден
  Офис_найден = False

  ' Найти строку ДО/ОО2
  rowCount = rowByValue(In_ReportName_String, In_Sheets, "ДО/ОО2", 100, 100)
  ' rowCount = 3
  
  Do While (Not IsEmpty(Workbooks(In_ReportName_String).Sheets(In_Sheets).Cells(rowCount, In_ColOffice).Value)) And (Офис_найден = False)
  ' Do While Not IsEmpty(Workbooks(In_ReportName_String).Sheets(In_Sheets).Cells(rowCount, In_ColOffice).Value)
        
    ' Если это текущий офис
    If InStr(Workbooks(In_ReportName_String).Sheets(In_Sheets).Cells(rowCount, In_ColOffice).Value, In_officeNameInReport) <> 0 Then
            
      ' План Квартала
      If In_Col4 <> 0 Then
        ' Берем из позиции квартала и вставляем в Лист8 столбец 5 (План квартала)
        ThisWorkbook.Sheets("Лист8").Cells(In_rowSheet8, 5).Value = Workbooks(In_ReportName_String).Sheets(In_Sheets).Cells(rowCount, In_Col4).Value
      Else
        ' План квартала получаем из суммы плана 3-х месяцев: In_Col1, In_Col2, In_Col3
        ThisWorkbook.Sheets("Лист8").Cells(In_rowSheet8, 5).Value = Workbooks(In_ReportName_String).Sheets(In_Sheets).Cells(rowCount, In_Col1).Value + Workbooks(In_ReportName_String).Sheets(In_Sheets).Cells(rowCount, In_Col2).Value + Workbooks(In_ReportName_String).Sheets(In_Sheets).Cells(rowCount, In_Col3).Value
      End If
      
      ' Центруем ячейку
      ThisWorkbook.Sheets("Лист8").Cells(In_rowSheet8, 5).HorizontalAlignment = xlRight
      
      ' Отмечаем у номера точкой + "."
      ' ThisWorkbook.Sheets("Лист8").Cells(In_rowSheet8, 1).NumberFormat = "@"
      ThisWorkbook.Sheets("Лист8").Cells(In_rowSheet8, 1).Value = CStr(ThisWorkbook.Sheets("Лист8").Cells(In_rowSheet8, 1).Value) + "."
      
      ' Офис был найден
      Офис_найден = True
      
    End If
        
    ' Следующая запись
    rowCount = rowCount + 1
    Application.StatusBar = In_officeNameInReport + ": " + CStr(rowCount) + "..."
    DoEventsInterval (rowCount)
  Loop


End Sub


' Сформировать ИПЗ для управляющих
Sub Сформировать_ПЗ_УДО()
  
  If MsgBox("Сформировать ПЗ для Управляющих и НОРПиКО на " + ThisWorkbook.Sheets("Лист8").Range("E8").Value + "?", vbYesNo) = vbYes Then
        
    ' Дата формирования документа
    ДатаПЗ = Date
        
    ' Число ПЗ
    CountПЗ = 0
    
    ' *** ***
    ' Цикл по 5-ти офисам
    ' Обработка отчета
    For i = 1 To 9
      
      ' Номера офисов от 1 до 5
      Select Case i
        Case 1 ' ОО «Тюменский», НОРП1
          officeNameInReport = "Тюменский"
          row_Наименование = getRowFromSheet8("Тюменский", "Наименование") ' 9
          functionVar = "НОРПиКО1"
          row_End = "ОО «Сургутский»"
        Case 2 ' ОО «Сургутский», УДО2
          officeNameInReport = "Сургутский"
          row_Наименование = getRowFromSheet8("Сургутский", "Наименование") ' 47
          functionVar = "УДО2"
          row_End = "ОО «Нижневартовский»"
        Case 3 ' ОО «Сургутский», НОРП2
          officeNameInReport = "Сургутский"
          row_Наименование = getRowFromSheet8("Сургутский", "Наименование") ' 47
          functionVar = "НОРПиКО2"
          row_End = "ОО «Нижневартовский»"
        Case 4 ' ОО «Нижневартовский», УДО3
          officeNameInReport = "Нижневартовский"
          row_Наименование = getRowFromSheet8("Нижневартовский", "Наименование") ' 85
          functionVar = "УДО3"
          row_End = "ОО «Новоуренгойский»"
        Case 5 ' ОО «Нижневартовский», НОРП3
          officeNameInReport = "Нижневартовский"
          row_Наименование = getRowFromSheet8("Нижневартовский", "Наименование") ' 85
          functionVar = "НОРПиКО3"
          row_End = "ОО «Новоуренгойский»"
        Case 6 ' ОО «Новоуренгойский», УДО4
          officeNameInReport = "Новоуренгойский"
          row_Наименование = getRowFromSheet8("Новоуренгойский", "Наименование") ' 123
          functionVar = "УДО4"
          row_End = "ОО «Тарко-Сале»"
        Case 7 ' ОО «Новоуренгойский», НОРП4
          officeNameInReport = "Новоуренгойский"
          row_Наименование = getRowFromSheet8("Новоуренгойский", "Наименование") ' 123
          functionVar = "НОРПиКО4"
          row_End = "ОО «Тарко-Сале»"
        Case 8 ' ОО «Тарко-Сале», УДО5
          officeNameInReport = "Тарко-Сале"
          row_Наименование = getRowFromSheet8("Тарко-Сале", "Наименование") ' 161
          functionVar = "УДО5"
          row_End = "Интегральный рейтинг по офисам"
        Case 9 ' ОО «Тарко-Сале», НОРП5
          officeNameInReport = "Тарко-Сале"
          row_Наименование = getRowFromSheet8("Тарко-Сале", "Наименование") ' 161
          functionVar = "НОРПиКО5"
          row_End = "Интегральный рейтинг по офисам"
      End Select
                 
      ' Открываем шаблон ИПЗ
      Workbooks.Open (ThisWorkbook.Path + "\Templates\ПЗ.xlsx")
         
      ' Переход на Лист8
      ThisWorkbook.Sheets("Лист8").Activate
         
      ' Из B5 "Интегральный рейтинг по сотрудникам на 15.07.2020 г." берем дату
      ДатаПЗDB = Mid(ThisWorkbook.Sheets("Лист8").Range("B5").Value, 40, 10)
         
      ' Имя файла с ИПЗ
      ' ОТВЕТЫ НА ЧАСТО ЗАДАВАЕМЫЕ ВОПРОСЫ ПО ЕСУП: КАК НАЗЫВАТЬ ФАЙЛЫ. ВНИМАНИЕ! ИЗМЕНЕНИЕ МЕТОДОЛОГИИ! Если это общее (командное) мероприятие, то название файла должно быть следующего формата: «Наименование ИФР _ наименование ДО _ дата», например: «Протокол _ ДО Звездный_01.09.2018». Если это индивидуальная встреча /активность/мероприятие, то название файла должно быть следующего формата: «Наименование ИФР _ ФИО _ дата», например: «Карта достижений _ Иванов_ 01.09.2019 », «ЛИР_Петров_01.08.2019», «ИПР_Сидоров_01.03.2019»
    
      NameVar = getFromAddrBook(functionVar, 3)
      
      ' Если NameVar<>""
      If NameVar <> "" Then
      
        CountПЗ = CountПЗ + 1
        Application.StatusBar = CStr(CountПЗ) + ". " + NameVar + "..."
      
        FileIPZName = "ПЗ_" + NameVar + "_" + CStr(ДатаПЗ) + ".xlsx"
        
        ' Проверяем - если файл есть, то удаляем его
        Call deleteFile(ThisWorkbook.Path + "\Out\" + FileIPZName)
        
        Workbooks("ПЗ.xlsx").SaveAs FileName:=ThisWorkbook.Path + "\Out\" + FileIPZName, FileFormat:=xlOpenXMLWorkbook, createBackUp:=False
            
        ' Должность + ФИО
        Workbooks(FileIPZName).Sheets("Лист1").Range("F7").Value = getFromAddrBook(functionVar, 1)
        ' ФИО
        Workbooks(FileIPZName).Sheets("Лист1").Range("H40").Value = NameVar
        Workbooks(FileIPZName).Sheets("Лист1").Range("H40").HorizontalAlignment = xlLeft
        ' Офис
        ' Workbooks(FileIPZName).Sheets("Лист1").Range("F9").Value = "ОО «" + officeNameInReport + "»"
        ' Дата - "17" июля 2020 г.
        Workbooks(FileIPZName).Sheets("Лист1").Range("G10").Value = ДеньМесяцГод(ДатаПЗ)
        
        ' Текст ПЗ
        ' Если это УДО
        If (Mid(functionVar, 1, 3) = "УДО") Then
          Workbooks(FileIPZName).Sheets("Лист1").Range("A13").Value = "               С целью выполнения планов " + quarterName(ДатаПЗ) + " розничного бизнеса РОО «Тюменский» прошу принять к исполнению плановое задание ОО «" + officeNameInReport + "»"
        End If
        ' Если это НОРПиКО
        If (Mid(functionVar, 1, 7) = "НОРПиКО") Then
          Workbooks(FileIPZName).Sheets("Лист1").Range("A13").Value = "               С целью выполнения планов " + quarterName(ДатаПЗ) + " розничного бизнеса РОО «Тюменский» прошу принять к исполнению плановое задание ОРПиКО ОО «" + officeNameInReport + "»"
        End If
        
        ' Заголовок - План            4 кв. 2020
        Workbooks(FileIPZName).Sheets("Лист1").Range("G16").Value = "План            " + quarterName(ДатаПЗ)
        Workbooks(FileIPZName).Sheets("Лист1").Range("G16").WrapText = True
        Workbooks(FileIPZName).Sheets("Лист1").Range("G16").ColumnWidth = 12
        
        ' Число пунктов в ПЗ
        ЧислоПунктовПЗ = 0
        ' Нумерация в ПЗ
        НумерацияПЗ = 0
        
        ' ====
        rowCount = row_Наименование + 1
        ' Do While Not IsEmpty(ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value)
        ' Do While ThisWorkbook.Sheets("Лист8").Cells(rowCount, 1).Value <> ""
        Do While ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value <> row_End
            
          ' Выводим позицию
          ВыводимПозицию = False
            
          ' Если это УДО
          If (Mid(functionVar, 1, 3) = "УДО") And ((ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value = "Зарплатные карты 18+") Or (ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value = "Портфель ЗП 18+, шт._Квартал ")) Then
            
            ' Выводим позицию
            ВыводимПозицию = True
            
          End If
      
          ' Выводим
          If (ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value = "Потребительские кредиты") Or _
               (ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value = "Кредитные карты (актив.)") Or _
                 (ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value = "Дебетовые карты (актив.)") Or _
                   (ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value = "Интернет-банк") Or _
                     (ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value = "Orange Premium Club") Or _
                       (ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value = "Кредитные карты Affluent") Or _
                         (ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value = "Комиссионный доход") Or _
                           (ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value = "в т.ч. страховки к ПК") Or _
                             (ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value = "           ИСЖ_МАСС") Or _
                               (ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value = "           НСЖ_МАСС") Or _
                                 (ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value = "           КС") Or _
                                   (ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value = "           ЛА") Or _
                                     (ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value = "           ИСЖ, НСЖ, КС (Affluent)") Or _
                                       (ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value = "           УК (Affluent)") Or _
                                         (ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value = "Пассивы") Or _
                                           (ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value = "Инвест") Or _
                                             (ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value = "Инвест Брокер обслуж") _
          Then
            
            ' Выводим позицию
            ВыводимПозицию = True
            
          End If
          
          ' Выводим позицию
          If ВыводимПозицию = True Then
          
            ' Число пунктов в ПЗ
            ЧислоПунктовПЗ = ЧислоПунктовПЗ + 1
          
            ' Выводим планы по продуктам
            ' №
            If Not ((ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value = "в т.ч. страховки к ПК") Or _
                 (ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value = "           ИСЖ_МАСС") Or _
                    (ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value = "           НСЖ_МАСС") Or _
                       (ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value = "           КС") Or _
                          (ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value = "           ЛА") Or _
                             (ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value = "           ИСЖ, НСЖ, КС (Affluent)") Or _
                                (ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value = "           УК (Affluent)")) _
            Then
              ' Нумерация в ПЗ
              НумерацияПЗ = НумерацияПЗ + 1
              Workbooks(FileIPZName).Sheets("Лист1").Cells(ЧислоПунктовПЗ + 16, 2).Value = CStr(НумерацияПЗ)
            End If
            
            ' Проукт
            Workbooks(FileIPZName).Sheets("Лист1").Cells(ЧислоПунктовПЗ + 16, 3).Value = ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value
            Workbooks(FileIPZName).Sheets("Лист1").Cells(ЧислоПунктовПЗ + 16, 3).RowHeight = 15
            Workbooks(FileIPZName).Sheets("Лист1").Cells(ЧислоПунктовПЗ + 16, 3).HorizontalAlignment = xlLeft
            ' Единица измерения
            Workbooks(FileIPZName).Sheets("Лист1").Cells(ЧислоПунктовПЗ + 16, 6).Value = ThisWorkbook.Sheets("Лист8").Cells(rowCount, 4).Value
            ' План
            Workbooks(FileIPZName).Sheets("Лист1").Cells(ЧислоПунктовПЗ + 16, 7).Value = ThisWorkbook.Sheets("Лист8").Cells(rowCount, 5).Value
            Workbooks(FileIPZName).Sheets("Лист1").Cells(ЧислоПунктовПЗ + 16, 7).HorizontalAlignment = xlRight
          End If
                  
          ' Следующая запись
          rowCount = rowCount + 1
          DoEventsInterval (rowCount)

        Loop
    
        ' Закрытие файла
        Workbooks(FileIPZName).Close SaveChanges:=True
    
      End If ' Если ФИО из справочника <>""
    
    Next i
    ' ***
    
    ' Статус
    Application.StatusBar = ""

    ' =====
    
    MsgBox ("ПЗ в количестве " + CStr(CountПЗ) + " шт. сформированы!")
    
    ' Перенести файл протокола в каталог ЕСУП? - https://www.excel-vba.ru/chto-umeet-excel/kak-sredstvami-vba-pereimenovatperemestitskopirovat-fajl/
    If MsgBox("Скопировать файлы ПЗ сотрудников в каталог ЕСУП (Индивидуальные встречи)?", vbYesNo) = vbYes Then
  
      ' Строка статуса
      Application.StatusBar = "Копирование в каталог ЕСУП..."
    
      CountПЗ = 0
    
      ' ***
      ' Цикл по 5-ти офисам
      ' Обработка отчета
      For i = 1 To 9
      
        ' Номера офисов от 1 до 5
        Select Case i
          Case 1 ' ОО «Тюменский», НОРП1
            functionVar = "НОРПиКО1"
          Case 2 ' ОО «Сургутский», УДО2
            functionVar = "УДО2"
          Case 3 ' ОО «Сургутский», НОРП2
            functionVar = "НОРПиКО2"
          Case 4 ' ОО «Нижневартовский», УДО3
            functionVar = "УДО3"
          Case 5 ' ОО «Нижневартовский», НОРП3
            functionVar = "НОРПиКО3"
          Case 6 ' ОО «Новоуренгойский», УДО4
            functionVar = "УДО4"
          Case 7 ' ОО «Новоуренгойский», НОРП4
            functionVar = "НОРПиКО4"
          Case 8 ' ОО «Тарко-Сале», УДО5
            functionVar = "УДО5"
          Case 9 ' ОО «Тарко-Сале», НОРП5
            functionVar = "НОРПиКО5"
        End Select
                    
        NameVar = getFromAddrBook(functionVar, 3)
      
        ' Если NameVar<>""
        If NameVar <> "" Then
                    
          CountПЗ = CountПЗ + 1
        
          ' Имя файла с ПЗ
          FileIPZName = ThisWorkbook.Path + "\Out\" + "ПЗ_" + NameVar + "_" + CStr(ДатаПЗ) + ".xlsx"

          ' Строка статуса
          Application.StatusBar = CStr(CountПЗ) + " Копирование " + FileIPZName + "..."
           
          ' ОТВЕТЫ НА ЧАСТО ЗАДАВАЕМЫЕ ВОПРОСЫ ПО ЕСУП: КАК НАЗЫВАТЬ ФАЙЛЫ. ВНИМАНИЕ! ИЗМЕНЕНИЕ МЕТОДОЛОГИИ! Если это общее (командное) мероприятие, то название файла должно быть следующего формата: «Наименование ИФР _ наименование ДО _ дата», например: «Протокол _ ДО Звездный_01.09.2018». Если это индивидуальная встреча /активность/мероприятие, то название файла должно быть следующего формата: «Наименование ИФР _ ФИО _ дата», например: «Карта достижений _ Иванов_ 01.09.2019 », «ЛИР_Петров_01.08.2019», «ИПР_Сидоров_01.03.2019»
          FileCopy FileIPZName, "\\probank\DavWWWRoot\drp\DocLib1\Тюменский ОО1\Управленческие процедуры\Индивидуальные встречи\" + "ПЗ_" + NameVar + "_" + CStr(ДатаПЗ) + ".xlsx"
   
        End If ' Если NameVar<>""
   
      Next i
      ' ***
   
      Application.StatusBar = "Скопировано!"
        
      ' Строка статуса
      Application.StatusBar = ""

      ' Сообщение
      MsgBox ("ПЗ в количестве " + CStr(CountПЗ) + " шт. перенесены в каталог ЕСУП!")

    End If ' Перенос в ЕСУП
        
        
  End If
    
End Sub

' Делаем расчет План/Факт по текущему месяцу с Лист7 в офисном канале для текущего officeNameInReport
Sub План_Факт_ПК_Лист7(In_officeNameInReport, In_Row_Лист8, In_N, In_dateDB)
    
  ' ***
  In_Product_Code = "ПК_Офис"
  In_Product_Name = "в т.ч. Офисный канал"
  In_Unit = "тыс.руб."
  curr_Day_Month = "Date_" + Mid(In_dateDB, 1, 2)
  ' ***
      
  План_офис_ПК = 0
  Факт_Офис_ПК = 0
  Прогноз_Офис_ПК_тыс_руб = 0
    
  ' ====
  rowCount = 9
  Do While Not IsEmpty(ThisWorkbook.Sheets("Лист7").Cells(rowCount, 3).Value)
      
    ' Индикация
    Application.StatusBar = "Расчет ПК (офис)" + In_officeNameInReport + "..."
    
    ' Если текущая запись - это наш офис
    If ThisWorkbook.Sheets("Лист7").Cells(rowCount, 4).Value = In_officeNameInReport Then
      План_офис_ПК = План_офис_ПК + ThisWorkbook.Sheets("Лист7").Cells(rowCount, 5).Value
      Факт_Офис_ПК = Факт_Офис_ПК + ThisWorkbook.Sheets("Лист7").Cells(rowCount, 6).Value
      Прогноз_Офис_ПК_тыс_руб = Прогноз_Офис_ПК_тыс_руб + ThisWorkbook.Sheets("Лист7").Cells(rowCount, 8).Value
    End If
      
    ' Следующая запись
    rowCount = rowCount + 1
    DoEventsInterval (rowCount)

  Loop

  ' Выводим показатели в In_Row_Лист8
  ' Наименование продукта на Лист8
  ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 1).NumberFormat = "@"
  ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 1).Value = In_N
  ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 1).HorizontalAlignment = xlCenter
  '
  ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 2).Value = "в т.ч. Офисный канал"
  ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 2).HorizontalAlignment = xlLeft
  ' Unit
  ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 4).Value = "тыс.руб."
  ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 4).HorizontalAlignment = xlCenter
  ' Месяц - план
  ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 9).Value = План_офис_ПК
  ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 9).NumberFormat = "#,##0"
  ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 9).HorizontalAlignment = xlRight
  ' Месяц - факт
  ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 10).Value = Факт_Офис_ПК
  ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 10).NumberFormat = "#,##0"
  ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 10).HorizontalAlignment = xlRight
  ' Месяц - исполнение
  ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 11).Value = РассчетДоли(ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 9).Value, ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 10).Value, 3)
  ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 11).NumberFormat = "0%"
  ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 11).HorizontalAlignment = xlRight
  ' Прогноз в %
  ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 12).Value = РассчетДоли(План_офис_ПК, Прогноз_Офис_ПК_тыс_руб, 3)
  ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 12).NumberFormat = "0%"
  ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 12).HorizontalAlignment = xlRight
  
  ' Окраска ячейки СФЕТОФОР ' 70. Заливка ячейки цветом "светофор" - значение In_Value в %, In_Target в %
  Call Full_Color_RangeII("Лист8", In_Row_Лист8, 12, ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 12).Value, 1)

  ' Месяц - прогноз (штуки, тыс.руб и т.п.) делаем расчет
  PredictionVar = (ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 9).Value * ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 12).Value)
                
  ' Заносим в Sales_Office
  '  Идентификатор ID_Rec:
  ID_RecVar = CStr(CStr(getNumberOfficeByName(In_officeNameInReport)) + "-" + strMMYY(In_dateDB) + "-" + In_Product_Code)
           
  ' Вносим данные в BASE\Sales_Office по ПК.
  Call InsertRecordInBook("Sales_Office", "Лист1", "ID_Rec", ID_RecVar, _
                                            "ID_Rec", ID_RecVar, _
                                              "Оffice_Number", getNumberOfficeByName(In_officeNameInReport), _
                                                "Product_Name", In_Product_Name, _
                                                  "Оffice", In_officeNameInReport, _
                                                    "MMYY", strMMYY(In_dateDB), _
                                                      "Update_Date", In_dateDB, _
                                                       "Product_Code", In_Product_Code, _
                                                         "Plan", ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 9).Value, _
                                                            "Unit", In_Unit, _
                                                              "Fact", ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 10).Value, _
                                                                "Percent_Completion", ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 11).Value, _
                                                                  "Prediction", PredictionVar, _
                                                                    "Percent_Prediction", ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 12).Value, _
                                                                      curr_Day_Month, ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 10).Value, _
                                                                        "", "", _
                                                                          "", "", _
                                                                            "", "", _
                                                                              "", "", _
                                                                                "", "", _
                                                                                  "", "")


End Sub

' Сформировать поручения на неделю для НОРПиКО и УДО по данным DB
Sub createTaskFroWeekOfficeFromDB()
  
  ' Сообщение о неверном формате отчета или даты
  ' MsgBox ("Внимание! Необходимо обновить Cadr Emission перед началом!")
  
  ' Запрос на формирование
  If MsgBox("Сформировать поручения по Дашбоард до " + CStr(weekEndDate(Date) - 2) + " для НОРПиКО и УДО?", vbYesNo) = vbYes Then
    
    
    ' Счетчик поручений
    Счетчик_поручений = 0
    
    ' Цикл по 5-ти офисам
    For i = 1 To 5
      
      ' Номера офисов от 1 до 5
      Select Case i
        Case 1 ' ОО «Тюменский»
          officeNameInReport = "Тюменский"
          responsibleName = getFromAddrBook("НОРПиКО1", 3)
          ' row_КД_Лист8 = getRowFromSheet8("Тюменский", ) ' 25 ' в т.ч. страховки к ПК
          ' row_ЗаявкиКК_Лист8 = getRowFromSheet8("Тюменский", ) ' 43
          row_Карты_в_сейфе_Лист5 = 39
          ' по DB
          ' row_Потребительские_кредиты = getRowFromSheet8("Тюменский", ) ' 11
          ' row_Дебетовые_карты_актив = getRowFromSheet8("Тюменский", ) ' 14
          ' row_Интернет_банк = getRowFromSheet8("Тюменский", ) ' 15
          ' row_Накопительные_счета = getRowFromSheet8("Тюменский", ) ' 16
          ' row_Коробки_Личный_адвокат_премия = getRowFromSheet8("Тюменский", ) ' 44
          ' row_OPC = getRowFromSheet8("Тюменский", ) ' 17
          ' row_Комиссионный_доход = getRowFromSheet8("Тюменский", ) ' 24
          ' row_Кредитные_карты_актив = getRowFromSheet8("Тюменский", ) ' 13
        Case 2 ' ОО «Сургутский»
          officeNameInReport = "Сургутский"
          responsibleName = getFromAddrBook("УДО2", 3)
          ' row_КД_Лист8 = getRowFromSheet8("Сургутский", ) ' 63 ' в т.ч. страховки к ПК
          ' row_ЗаявкиКК_Лист8 = getRowFromSheet8("Сургутский", ) ' 81
          row_Карты_в_сейфе_Лист5 = 40
          ' по DB
          ' row_Потребительские_кредиты = getRowFromSheet8("Сургутский", ) ' 49
          ' row_Дебетовые_карты_актив = getRowFromSheet8("Сургутский", ) ' 52
          ' row_Интернет_банк = getRowFromSheet8("Сургутский", ) ' 53
          ' row_Накопительные_счета = getRowFromSheet8("Сургутский", ) ' 54
          ' row_Коробки_Личный_адвокат_премия = getRowFromSheet8("Сургутский", ) ' 82
          ' row_OPC = getRowFromSheet8("Сургутский", ) ' 55
          ' row_Комиссионный_доход = getRowFromSheet8("Сургутский", ) ' 62
          ' row_Кредитные_карты_актив = getRowFromSheet8("Сургутский", ) ' 51
        Case 3 ' ОО «Нижневартовский»
          officeNameInReport = "Нижневартовский"
          responsibleName = getFromAddrBook("НОРПиКО3", 3)
          ' row_КД_Лист8 = getRowFromSheet8("Нижневартовский", ) ' 101 ' в т.ч. страховки к ПК
          ' row_ЗаявкиКК_Лист8 = getRowFromSheet8("Нижневартовский", ) ' 119
          row_Карты_в_сейфе_Лист5 = 41
          ' по DB
          ' row_Потребительские_кредиты = getRowFromSheet8("Нижневартовский", ) ' 87
          ' row_Дебетовые_карты_актив = getRowFromSheet8("Нижневартовский", ) ' 90
          ' row_Интернет_банк = getRowFromSheet8("Нижневартовский", ) ' 91
          ' row_Накопительные_счета = getRowFromSheet8("Нижневартовский", ) ' 92
          ' row_Коробки_Личный_адвокат_премия = getRowFromSheet8("Нижневартовский", ) ' 120
          ' row_OPC = getRowFromSheet8("Нижневартовский", ) ' 93
          ' row_Комиссионный_доход = getRowFromSheet8("Нижневартовский", ) ' 100
          ' row_Кредитные_карты_актив = getRowFromSheet8("Нижневартовский", ) ' 89
        Case 4 ' ОО «Новоуренгойский»
          officeNameInReport = "Новоуренгойский"
          responsibleName = getFromAddrBook("УДО4", 3)
          ' row_КД_Лист8 = getRowFromSheet8("Новоуренгойский", ) ' 139 ' в т.ч. страховки к ПК
          ' row_ЗаявкиКК_Лист8 = getRowFromSheet8("Новоуренгойский", ) ' 157
          row_Карты_в_сейфе_Лист5 = 42
          ' по DB
          ' row_Потребительские_кредиты = getRowFromSheet8("Новоуренгойский", ) ' 125
          ' row_Дебетовые_карты_актив = getRowFromSheet8("Новоуренгойский", ) ' 128
          ' row_Интернет_банк = getRowFromSheet8("Новоуренгойский", ) ' 129
          ' row_Накопительные_счета = getRowFromSheet8("Новоуренгойский", ) ' 130
          ' row_Коробки_Личный_адвокат_премия = getRowFromSheet8("Новоуренгойский", ) ' 158
          ' row_OPC = getRowFromSheet8("Новоуренгойский", ) ' 131
          ' row_Комиссионный_доход = getRowFromSheet8("Новоуренгойский", ) ' 138
          ' row_Кредитные_карты_актив = getRowFromSheet8("Новоуренгойский", ) ' 127
        Case 5 ' ОО «Тарко-Сале»
          officeNameInReport = "Тарко-Сале"
          responsibleName = getFromAddrBook("НОРПиКО5", 3)
          ' row_КД_Лист8 = getRowFromSheet8("Тарко-Сале", ) ' 177 ' в т.ч. страховки к ПК
          ' row_ЗаявкиКК_Лист8 =  getRowFromSheet8("Тарко-Сале", ) ' 195
          row_Карты_в_сейфе_Лист5 = 43
          ' по DB
          ' row_Потребительские_кредиты =  getRowFromSheet8("Тарко-Сале", ) ' 163
          ' row_Дебетовые_карты_актив =  getRowFromSheet8("Тарко-Сале", ) ' 166
          ' row_Интернет_банк =  getRowFromSheet8("Тарко-Сале", ) ' 167
          ' row_Накопительные_счета =  getRowFromSheet8("Тарко-Сале", ) ' 168
          ' row_Коробки_Личный_адвокат_премия =  getRowFromSheet8("Тарко-Сале", ) ' 196
          ' row_OPC =  getRowFromSheet8("Тарко-Сале", ) ' 169
          ' row_Комиссионный_доход =  getRowFromSheet8("Тарко-Сале", ) ' 176
          ' row_Кредитные_карты_актив =  getRowFromSheet8("Тарко-Сале", ) ' 165
      End Select
    
    ' В новой версии переносим сюда
    row_КД_Лист8 = getRowFromSheet8(officeNameInReport, "в т.ч. страховки к ПК") ' 177 ' в т.ч. страховки к ПК
    row_ЗаявкиКК_Лист8 = getRowFromSheet8(officeNameInReport, "Заявки на Кредитные карты")   ' 195
    ' row_Карты_в_сейфе_Лист5 = getRowFromSheet8(officeNameInReport, "")  ' 43
    ' по DB
    row_Потребительские_кредиты = getRowFromSheet8(officeNameInReport, "Потребительские кредиты")  ' 163
    row_Дебетовые_карты_актив = getRowFromSheet8(officeNameInReport, "Дебетовые карты (актив.)")  ' 166
    row_Интернет_банк = getRowFromSheet8(officeNameInReport, "Интернет-банк")  ' 167
    row_Накопительные_счета = getRowFromSheet8(officeNameInReport, "Накопительные счета")  ' 168
    row_Коробки_Личный_адвокат_премия = getRowFromSheet8(officeNameInReport, "Коробки+Личный адвокат (премия)")  ' 196
    row_OPC = getRowFromSheet8(officeNameInReport, "Orange Premium Club")  ' 169
    row_Комиссионный_доход = getRowFromSheet8(officeNameInReport, "Комиссионный доход")  ' 176
    row_Кредитные_карты_актив = getRowFromSheet8(officeNameInReport, "Кредитные карты (актив.)")  ' 165
    
    ' Обнуляем переменные
    Дефицит_ПК = 0
    Дефицит_КК = 0
    Дефицит_ДК = 0
    Дефицит_ИБ = 0
    Дефицит_НС = 0
    Дефицит_КСП = 0
    Дефицит_OPC = 0
    План_Заявки_КК_неделя = 0
    
    
    Application.StatusBar = "Формирование пакета поручений " + CStr(i) + "..."

    ' Дата начала недели
    Дата_начала_недели = weekStartDate(Date)
    ' Дата окончания недели
    Дата_окончания_недели = Дата_начала_недели + 4
    ' Остаток рабочих дней определяем число рабочих дней с понеделника до конца месяца Working_days_between_dates(In_DateStart, In_DateEnd, In_working_days_in_the_week) As Integer
    Остаток_рабочих_дней = Working_days_between_dates(Дата_начала_недели, Date_last_day_month(Дата_начала_недели), 5)
          
    ' Выводим поручения по текущему офису
    ' 1) ПК
    Дефицит_ПК = Round(ThisWorkbook.Sheets("Лист8").Cells(row_Потребительские_кредиты, 9).Value - ThisWorkbook.Sheets("Лист8").Cells(row_Потребительские_кредиты, 10).Value, 0)
    План_ПК_неделя = Round(Дефицит_ПК / Остаток_рабочих_дней, 0) * (Дата_окончания_недели - Дата_начала_недели + 1)
    If План_ПК_неделя > 0 Then
      Call Вставка_строки_в_Поручения_участникам(responsibleName, Дата_окончания_недели, "Обеспечить выдачу потребительских кредитов на сумму не менее " + CStr(План_ПК_неделя) + " тыс.руб.")
      ' Значение плана на неделю в столбце "M" (13)
      ThisWorkbook.Sheets("Лист8").Cells(row_Потребительские_кредиты, 13).Value = План_ПК_неделя
      ThisWorkbook.Sheets("Лист8").Cells(row_Потребительские_кредиты, 13).NumberFormat = "#,##0"
      ThisWorkbook.Sheets("Лист8").Cells(row_Потребительские_кредиты, 13).HorizontalAlignment = xlRight
      ' Счетчик поручений
      Счетчик_поручений = Счетчик_поручений + 1
    End If
    
    ' 2) КК
    ' Заявки КК на месяц
    Дефицит_Заявки_КК = Round(ThisWorkbook.Sheets("Лист8").Cells(row_ЗаявкиКК_Лист8, 9).Value - ThisWorkbook.Sheets("Лист8").Cells(row_ЗаявкиКК_Лист8, 10).Value, 0)
    План_Заявки_КК_неделя = Round(Дефицит_Заявки_КК / Остаток_рабочих_дней, 0) * (Дата_окончания_недели - Дата_начала_недели + 1)
    If План_Заявки_КК_неделя > 0 Then
      Call Вставка_строки_в_Поручения_участникам(responsibleName, Дата_окончания_недели, "Завести заявки на кредитные карты не менее " + CStr(План_Заявки_КК_неделя) + " шт.")
      ' Значение плана на неделю в столбце "M" (13)
      ThisWorkbook.Sheets("Лист8").Cells(row_ЗаявкиКК_Лист8, 13).Value = План_Заявки_КК_неделя
      ThisWorkbook.Sheets("Лист8").Cells(row_ЗаявкиКК_Лист8, 13).NumberFormat = "#,##0"
      ThisWorkbook.Sheets("Лист8").Cells(row_ЗаявкиКК_Лист8, 13).HorizontalAlignment = xlRight
      ' Счетчик поручений
      Счетчик_поручений = Счетчик_поручений + 1
    End If
        
    
    ' 2.1) КК 2
    ' Кредитные карты (актив.)
    Дефицит_Кредитные_карты_актив = Round(ThisWorkbook.Sheets("Лист8").Cells(row_Кредитные_карты_актив, 9).Value - ThisWorkbook.Sheets("Лист8").Cells(row_Кредитные_карты_актив, 10).Value, 0)
    План_Кредитные_карты_актив = Round(Дефицит_Кредитные_карты_актив / Остаток_рабочих_дней, 0) * (Дата_окончания_недели - Дата_начала_недели + 1)
    If План_Кредитные_карты_актив > 0 Then
      Call Вставка_строки_в_Поручения_участникам(responsibleName, Дата_окончания_недели, "Активировать кредитные карты не менее " + CStr(План_Кредитные_карты_актив) + " шт.")
      ' Значение плана на неделю в столбце "M" (13)
      ThisWorkbook.Sheets("Лист8").Cells(row_Кредитные_карты_актив, 13).Value = План_Кредитные_карты_актив
      ThisWorkbook.Sheets("Лист8").Cells(row_Кредитные_карты_актив, 13).NumberFormat = "#,##0"
      ThisWorkbook.Sheets("Лист8").Cells(row_Кредитные_карты_актив, 13).HorizontalAlignment = xlRight
      ' Счетчик поручений
      Счетчик_поручений = Счетчик_поручений + 1
    End If
        
        
    ' 3) ДК
    Дефицит_ДК = Round(ThisWorkbook.Sheets("Лист8").Cells(row_Дебетовые_карты_актив, 9).Value - ThisWorkbook.Sheets("Лист8").Cells(row_Дебетовые_карты_актив, 10).Value, 0)
    План_ДК_неделя = Round(Дефицит_ДК / Остаток_рабочих_дней, 0) * (Дата_окончания_недели - Дата_начала_недели + 1)
    If План_ДК_неделя > 0 Then
      Call Вставка_строки_в_Поручения_участникам(responsibleName, Дата_окончания_недели, "Обеспечить заведение заявок дебетовых карт не менее " + CStr(План_ДК_неделя) + " шт.")
      ' Значение плана на неделю в столбце "M" (13)
      ThisWorkbook.Sheets("Лист8").Cells(row_Дебетовые_карты_актив, 13).Value = План_ДК_неделя
      ThisWorkbook.Sheets("Лист8").Cells(row_Дебетовые_карты_актив, 13).NumberFormat = "#,##0"
      ThisWorkbook.Sheets("Лист8").Cells(row_Дебетовые_карты_актив, 13).HorizontalAlignment = xlRight
      ' Счетчик поручений
      Счетчик_поручений = Счетчик_поручений + 1
    End If
    
    ' 4) ИБ
    Дефицит_ИБ = Round(ThisWorkbook.Sheets("Лист8").Cells(row_Интернет_банк, 9).Value - ThisWorkbook.Sheets("Лист8").Cells(row_Интернет_банк, 10).Value, 0)
    План_ИБ_неделя = Round(Дефицит_ИБ / Остаток_рабочих_дней, 0) * (Дата_окончания_недели - Дата_начала_недели + 1)
    If План_ИБ_неделя > 0 Then
      Call Вставка_строки_в_Поручения_участникам(responsibleName, Дата_окончания_недели, "Обеспечить подключение Интернет-банка в кол-ве не менее " + CStr(План_ИБ_неделя) + " шт.")
      ' Значение плана на неделю в столбце "M" (13)
      ThisWorkbook.Sheets("Лист8").Cells(row_Интернет_банк, 13).Value = План_ИБ_неделя
      ThisWorkbook.Sheets("Лист8").Cells(row_Интернет_банк, 13).NumberFormat = "#,##0"
      ThisWorkbook.Sheets("Лист8").Cells(row_Интернет_банк, 13).HorizontalAlignment = xlRight
      ' Счетчик поручений
      Счетчик_поручений = Счетчик_поручений + 1
    End If
    
    ' 5) НС
    Дефицит_НС = Round(ThisWorkbook.Sheets("Лист8").Cells(row_Накопительные_счета, 9).Value - ThisWorkbook.Sheets("Лист8").Cells(row_Накопительные_счета, 10).Value, 0)
    План_НС_неделя = Round(Дефицит_НС / Остаток_рабочих_дней, 0) * (Дата_окончания_недели - Дата_начала_недели + 1)
    If План_НС_неделя > 0 Then
      Call Вставка_строки_в_Поручения_участникам(responsibleName, Дата_окончания_недели, "Открыть Накопительные счета не менее " + CStr(План_НС_неделя) + " шт.")
      ' Значение плана на неделю в столбце "M" (13)
      ThisWorkbook.Sheets("Лист8").Cells(row_Накопительные_счета, 13).Value = План_НС_неделя
      ThisWorkbook.Sheets("Лист8").Cells(row_Накопительные_счета, 13).NumberFormat = "#,##0"
      ThisWorkbook.Sheets("Лист8").Cells(row_Накопительные_счета, 13).HorizontalAlignment = xlRight
      ' Счетчик поручений
      Счетчик_поручений = Счетчик_поручений + 1
    End If
    
    ' 6) КСП (премия)
    Дефицит_КСП = Round(ThisWorkbook.Sheets("Лист8").Cells(row_Коробки_Личный_адвокат_премия, 9).Value - ThisWorkbook.Sheets("Лист8").Cells(row_Коробки_Личный_адвокат_премия, 10).Value, 0)
    План_КСП_неделя = Round(Дефицит_КСП / Остаток_рабочих_дней, 0) * (Дата_окончания_недели - Дата_начала_недели + 1)
    If План_НС_неделя > 0 Then
      Call Вставка_строки_в_Поручения_участникам(responsibleName, Дата_окончания_недели, "Обеспечить премию с продажи КСП не менее " + CStr(План_КСП_неделя) + " тыс.руб.")
      ' Значение плана на неделю в столбце "M" (13)
      ThisWorkbook.Sheets("Лист8").Cells(row_Коробки_Личный_адвокат_премия, 13).Value = План_КСП_неделя
      ThisWorkbook.Sheets("Лист8").Cells(row_Коробки_Личный_адвокат_премия, 13).NumberFormat = "#,##0"
      ThisWorkbook.Sheets("Лист8").Cells(row_Коробки_Личный_адвокат_премия, 13).HorizontalAlignment = xlRight
      ' Счетчик поручений
      Счетчик_поручений = Счетчик_поручений + 1
    End If
    
    ' 7) КД ' (из Лист8 берем по каждому офису 5 позиций включая row_КД_Лист8 = 177 вниз ' в т.ч. страховки к ПК)
    Дефицит_КД = Round(ThisWorkbook.Sheets("Лист8").Cells(row_КД_Лист8, 9).Value _
                   + ThisWorkbook.Sheets("Лист8").Cells(row_КД_Лист8 + 1, 9).Value _
                     + ThisWorkbook.Sheets("Лист8").Cells(row_КД_Лист8 + 2, 9).Value _
                       + ThisWorkbook.Sheets("Лист8").Cells(row_КД_Лист8 + 3, 9).Value _
                         + ThisWorkbook.Sheets("Лист8").Cells(row_КД_Лист8 + 4, 9).Value _
                           - (ThisWorkbook.Sheets("Лист8").Cells(row_КД_Лист8, 10).Value _
                               + ThisWorkbook.Sheets("Лист8").Cells(row_КД_Лист8 + 1, 10).Value _
                                 + ThisWorkbook.Sheets("Лист8").Cells(row_КД_Лист8 + 2, 10).Value _
                                   + ThisWorkbook.Sheets("Лист8").Cells(row_КД_Лист8 + 3, 10).Value _
                                     + ThisWorkbook.Sheets("Лист8").Cells(row_КД_Лист8 + 4, 10).Value), 0)

                           
    
    План_КД_неделя = Round(Дефицит_КД / Остаток_рабочих_дней, 0) * (Дата_окончания_недели - Дата_начала_недели + 1)
    If План_КД_неделя > 0 Then
      Call Вставка_строки_в_Поручения_участникам(responsibleName, Дата_окончания_недели, "Обеспечить получение комиссионного дохода не менее " + CStr(План_КД_неделя) + " тыс.руб.")
      ' Значение плана на неделю в столбце "M" (13)
      ThisWorkbook.Sheets("Лист8").Cells(row_Комиссионный_доход, 13).Value = План_КД_неделя
      ThisWorkbook.Sheets("Лист8").Cells(row_Комиссионный_доход, 13).NumberFormat = "#,##0"
      ThisWorkbook.Sheets("Лист8").Cells(row_Комиссионный_доход, 13).HorizontalAlignment = xlRight
      ' Счетчик поручений
      Счетчик_поручений = Счетчик_поручений + 1
    End If
              
    ' 8) OPC
    Дефицит_OPC = ThisWorkbook.Sheets("Лист8").Cells(row_OPC, 9).Value - ThisWorkbook.Sheets("Лист8").Cells(row_OPC, 10).Value
    If Дефицит_OPC > 0 Then
      
      План_OPC_неделя = (Дефицит_OPC / Остаток_рабочих_дней) * (Дата_окончания_недели - Дата_начала_недели + 1)
      ' Если дефицит есть, то ставим минимум 1 пакет
      If (План_OPC_неделя > 0) And (План_OPC_неделя < 1) Then
        План_OPC_неделя = 1
      Else
        План_OPC_неделя = Round(План_OPC_неделя, 0)
      End If
      
      If План_OPC_неделя > 0 Then
        Call Вставка_строки_в_Поручения_участникам(responsibleName, Дата_окончания_недели, "Обеспечить выдачу пакетов OPC не менее " + CStr(План_OPC_неделя) + " шт.")
        ' Значение плана на неделю в столбце "M" (13)
        ThisWorkbook.Sheets("Лист8").Cells(row_OPC, 13).Value = План_OPC_неделя
        ThisWorkbook.Sheets("Лист8").Cells(row_OPC, 13).NumberFormat = "#,##0"
        ThisWorkbook.Sheets("Лист8").Cells(row_OPC, 13).HorizontalAlignment = xlRight
        ' Счетчик поручений
        Счетчик_поручений = Счетчик_поручений + 1
      End If
    End If
              
    ' 9) Заявки на ПК используем метрику 1 заявка на 1 МРК в день
    Call Вставка_строки_в_Поручения_участникам(responsibleName, Дата_окончания_недели, "Обеспечить заведение минимум 1 заявки на 1 МРК в день.")
                           
    Application.StatusBar = ""
              
              
    Next i
              
    ' В M8 остаток дней
    ThisWorkbook.Sheets("Лист8").Range("M8").Value = "Дней: " + CStr(Остаток_рабочих_дней)
              
    Application.StatusBar = ""
        
    ' Сообщение о неверном формате отчета или даты
    MsgBox ("Поручения в количестве " + CStr(Счетчик_поручений) + " сформированы!")

    ' Перейти на Лист ЕСУП
    Call goToSheetЕСУП
  
    ' Переход в часть листа с Поручениями
    ThisWorkbook.Sheets("ЕСУП").Range("AF77").Select

  End If
  
End Sub


' Выгрузить файл с дневным планом продаж по форме данилова Templates\Ежедневная форма отчёта (куратор).xlsx
Sub Выгрузить_план_дневных_продаж()
Dim FileNewVar As String

  ' Формируем цели на день по форме Данилова
  
  ' Запрос на формирование
  If MsgBox("Сформировать поручения на день для офисов?", vbYesNo) = vbYes Then
      
    ' Открываем шаблон C:\Users\proschaevsf\Documents\#DB_Result\Templates\Ежедневная форма отчёта (куратор).xlsx
    Workbooks.Open (ThisWorkbook.Path + "\Templates\Ежедневная форма отчёта (куратор).xlsx")
           
    ' Переходим на окно DB
    ThisWorkbook.Sheets("Лист8").Activate

    ' Дата формирования - если сегодня понедельник, то формируем за пятницу
    ' Если текущая дата это понедельник, то формируем отчет за пятницу
    If Weekday(CurrDate, vbMonday) = 1 Then
      dateReport = Date - 3
    Else
      dateReport = Date
    End If

    ' Имя нового файла
    FileNewVar = "Ежедневная форма отчёта_" + strДД_MM_YY(dateReport) + ".xlsx"
    
    ' Проверяем - если файл есть, то удаляем его
    Call deleteFile(ThisWorkbook.Path + "\Out\" + FileNewVar)
    
    Workbooks("Ежедневная форма отчёта (куратор).xlsx").SaveAs FileName:=ThisWorkbook.Path + "\Out\" + FileNewVar, FileFormat:=xlOpenXMLWorkbook, createBackUp:=False
    
    ' Остаток рабочих дней определяем число рабочих дней с понеделника до конца месяца Working_days_between_dateReports(In_dateReportStart, In_dateReportEnd, In_working_days_in_the_week) As Integer
    Остаток_рабочих_дней = Working_days_between_dates(dateReport - 1, Date_last_day_month(dateReport), 5)

    ' Проходим по Листу8 и заполняем планы:
    For i = 1 To 5
        ' Номера офисов от 1 до 5
        Select Case i
          Case 1 ' ОО «Тюменский»
            officeNameInReport = "ОО «Тюменский»"
            row_Лист1 = 29
          Case 2 ' ОО «Сургутский»
            officeNameInReport = "ОО «Сургутский»"
            row_Лист1 = 33
          Case 3 ' ОО «Нижневартовский»
            officeNameInReport = "ОО «Нижневартовский»"
            row_Лист1 = 30
          Case 4 ' ОО «Новоуренгойский»
            officeNameInReport = "ОО «Новоуренгойский»"
            row_Лист1 = 32
          Case 5 ' ОО «Тарко-Сале»
            officeNameInReport = "ОО «Тарко-Сале»"
            row_Лист1 = 31
        End Select
        
        ' Заполняем
        ' 1. ОО "Тюменский"
        ' ПК (29, 5)
        row_ПК = getRowFromSheet8(officeNameInReport, "Потребительские кредиты")
        If Round(((ThisWorkbook.Sheets("Лист8").Cells(row_ПК, 9).Value - ThisWorkbook.Sheets("Лист8").Cells(row_ПК, 10).Value) / Остаток_рабочих_дней), 0) > 0 Then
          Workbooks(FileNewVar).Sheets("Лист1").Cells(row_Лист1, 5).Value = Round(((ThisWorkbook.Sheets("Лист8").Cells(row_ПК, 9).Value - ThisWorkbook.Sheets("Лист8").Cells(row_ПК, 10).Value) / Остаток_рабочих_дней), 0)
        Else
          Workbooks(FileNewVar).Sheets("Лист1").Cells(row_Лист1, 5).Value = 0
        End If
        
        ' КСП + ЛА (премия)
        row_КСП_ЛА = getRowFromSheet8(officeNameInReport, "Коробки+Личный адвокат (премия)")
        If Round(((ThisWorkbook.Sheets("Лист8").Cells(row_КСП_ЛА, 9).Value - ThisWorkbook.Sheets("Лист8").Cells(row_КСП_ЛА, 10).Value) / Остаток_рабочих_дней), 0) > 0 Then
          Workbooks(FileNewVar).Sheets("Лист1").Cells(row_Лист1, 9).Value = Round(((ThisWorkbook.Sheets("Лист8").Cells(row_КСП_ЛА, 9).Value - ThisWorkbook.Sheets("Лист8").Cells(row_КСП_ЛА, 10).Value) / Остаток_рабочих_дней), 0)
        Else
          Workbooks(FileNewVar).Sheets("Лист1").Cells(row_Лист1, 9).Value = 0
        End If
        
        ' КК
        row_КК = getRowFromSheet8(officeNameInReport, "Кредитные карты (актив.)")
        If Round(((ThisWorkbook.Sheets("Лист8").Cells(row_КК, 9).Value - ThisWorkbook.Sheets("Лист8").Cells(row_КК, 10).Value) / Остаток_рабочих_дней), 0) > 0 Then
          Workbooks(FileNewVar).Sheets("Лист1").Cells(row_Лист1, 13).Value = Round(((ThisWorkbook.Sheets("Лист8").Cells(row_КК, 9).Value - ThisWorkbook.Sheets("Лист8").Cells(row_КК, 10).Value) / Остаток_рабочих_дней), 0)
        Else
          Workbooks(FileNewVar).Sheets("Лист1").Cells(row_Лист1, 13).Value = 0
          ' Ставим минимум 1
          Workbooks(FileNewVar).Sheets("Лист1").Cells(row_Лист1, 13).Value = 1
        End If
        
        ' ДК
        row_ДК = getRowFromSheet8(officeNameInReport, "Дебетовые карты (актив.)")
        If Round(((ThisWorkbook.Sheets("Лист8").Cells(row_ДК, 9).Value - ThisWorkbook.Sheets("Лист8").Cells(row_ДК, 10).Value) / Остаток_рабочих_дней), 0) > 0 Then
          Workbooks(FileNewVar).Sheets("Лист1").Cells(row_Лист1, 17).Value = Round(((ThisWorkbook.Sheets("Лист8").Cells(row_ДК, 9).Value - ThisWorkbook.Sheets("Лист8").Cells(row_ДК, 10).Value) / Остаток_рабочих_дней), 0)
        Else
          Workbooks(FileNewVar).Sheets("Лист1").Cells(row_Лист1, 17).Value = 0
          ' Ставим минимум 1
          Workbooks(FileNewVar).Sheets("Лист1").Cells(row_Лист1, 17).Value = 1
        End If
        
        ' ИСЖ
        row_ИСЖ = getRowFromSheet8(officeNameInReport, "Премия ИСЖ МАСС")
        If Round(((ThisWorkbook.Sheets("Лист8").Cells(row_ИСЖ, 9).Value - ThisWorkbook.Sheets("Лист8").Cells(row_ИСЖ, 10).Value) / Остаток_рабочих_дней), 0) > 0 Then
          Workbooks(FileNewVar).Sheets("Лист1").Cells(row_Лист1, 19).Value = Round(((ThisWorkbook.Sheets("Лист8").Cells(row_ИСЖ, 9).Value - ThisWorkbook.Sheets("Лист8").Cells(row_ИСЖ, 10).Value) / Остаток_рабочих_дней), 0)
        Else
          Workbooks(FileNewVar).Sheets("Лист1").Cells(row_Лист1, 19).Value = 0
        End If
        
        ' НСЖ
        row_НСЖ = getRowFromSheet8(officeNameInReport, "Премия НСЖ МАСС")
        If Round(((ThisWorkbook.Sheets("Лист8").Cells(row_НСЖ, 9).Value - ThisWorkbook.Sheets("Лист8").Cells(row_НСЖ, 10).Value) / Остаток_рабочих_дней), 0) > 0 Then
          Workbooks(FileNewVar).Sheets("Лист1").Cells(row_Лист1, 21).Value = Round(((ThisWorkbook.Sheets("Лист8").Cells(row_НСЖ, 9).Value - ThisWorkbook.Sheets("Лист8").Cells(row_НСЖ, 10).Value) / Остаток_рабочих_дней), 0)
        Else
          Workbooks(FileNewVar).Sheets("Лист1").Cells(row_Лист1, 21).Value = 0
        End If
        
        ' НС > 8 000 руб
        row_НС = getRowFromSheet8(officeNameInReport, "Накопительные счета")
        If Round(((ThisWorkbook.Sheets("Лист8").Cells(row_НС, 9).Value - ThisWorkbook.Sheets("Лист8").Cells(row_НС, 10).Value) / Остаток_рабочих_дней), 0) > 0 Then
          Workbooks(FileNewVar).Sheets("Лист1").Cells(row_Лист1, 23).Value = Round(((ThisWorkbook.Sheets("Лист8").Cells(row_НС, 9).Value - ThisWorkbook.Sheets("Лист8").Cells(row_НС, 10).Value) / Остаток_рабочих_дней), 0)
        Else
          Workbooks(FileNewVar).Sheets("Лист1").Cells(row_Лист1, 23).Value = 0
        End If
        
    Next i
    
    
    
    ' Закрываем файл
    Workbooks(FileNewVar).Close SaveChanges:=True

    MsgBox ("Сформирован файл " + ThisWorkbook.Path + "\Out\" + FileNewVar + "!")

    ' Отправка в почте в офисы
    Call Отправка_Lotus_Notes_Выгр_день_Лист8(ThisWorkbook.Path + "\Out\" + FileNewVar, dateReport)
      
    ' Письмо данилову: #t121121134 ">>: Re: Обновленная форма еженевного отчета"
      
  End If
  
End Sub

' Формирование свода по РОО на основании данных по каждому офису
Sub Свод_по_РОО()

      ' Вместо 209
      row_Итого_по_РОО_Тюменский = getRowFromSheet8("Итого по РОО «Тюменский»", "Итого по РОО «Тюменский»") + 2
      ' Вместо 9
      row_ОО_Тюменский = getRowFromSheet8("ОО «Тюменский»", "ОО «Тюменский»") + 2
      ' Вместо 47
      row_ОО_Сургутский = getRowFromSheet8("ОО «Сургутский»", "ОО «Сургутский»") + 2
      ' Вместо 85
      row_ОО_Нижневартовский = getRowFromSheet8("ОО «Нижневартовский»", "ОО «Нижневартовский»") + 2
      ' Вместо 123
      row_ОО_Новоуренгойский = getRowFromSheet8("ОО «Новоуренгойский»", "ОО «Новоуренгойский»") + 2
      ' Вместо 161
      row_ОО_Тарко_Сале = getRowFromSheet8("ОО «Тарко-Сале»", "ОО «Тарко-Сале»") + 2

      ' Дата DB с Лист8 (должны совпадать)
      dateDB_Лист8 = CDate(Mid(ThisWorkbook.Sheets("Лист8").Range("B5").Value, 52, 10))

      ' Определяем Число_показателей
      Число_показателей = 0
      rowCount = row_ОО_Тюменский + 1
      ' Do While (ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value <> "") And (InStr(ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value, "ОО «Сургутский»") = 0)
      Do While (InStr(ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value, "ОО «Сургутский»") = 0)
          
        Число_показателей = Число_показателей + 1
        
        ' Следующая запись
        Application.StatusBar = "Расчет показателей " + CStr(rowCount) + "..."
        rowCount = rowCount + 1
        DoEventsInterval (rowCount)
  
      Loop

      Application.StatusBar = ""


      ' РОО Тюменский - 35 строк показателей
      For i = 1 To Число_показателей
        
        ' Квартал: Вариант 1: когда в каждом офисе есть план и факт не равный нулю
        Вариант_расчета_прогноза = "1"
        
        ' Месяц: Вариант 1: когда в каждом офисе есть план и факт не равный нулю
        Вариант_расчета_прогноза_М = "1"
        
        ' №
        ThisWorkbook.Sheets("Лист8").Cells(row_Итого_по_РОО_Тюменский + i, 1).Value = ThisWorkbook.Sheets("Лист8").Cells(row_ОО_Тюменский + i, 1).Value
        ThisWorkbook.Sheets("Лист8").Cells(row_Итого_по_РОО_Тюменский + i, 1).NumberFormat = "@"
        ThisWorkbook.Sheets("Лист8").Cells(row_Итого_по_РОО_Тюменский + i, 1).HorizontalAlignment = xlCenter
        
        ' Наименование
        ThisWorkbook.Sheets("Лист8").Cells(row_Итого_по_РОО_Тюменский + i, 2).Value = ThisWorkbook.Sheets("Лист8").Cells(row_ОО_Тюменский + i, 2).Value
        ThisWorkbook.Sheets("Лист8").Cells(row_Итого_по_РОО_Тюменский + i, 2).HorizontalAlignment = xlLeft
        
        ' Вес
        ThisWorkbook.Sheets("Лист8").Cells(row_Итого_по_РОО_Тюменский + i, 3).Value = ThisWorkbook.Sheets("Лист8").Cells(row_ОО_Тюменский + i, 3).Value
        ThisWorkbook.Sheets("Лист8").Cells(row_Итого_по_РОО_Тюменский + i, 3).NumberFormat = "0.0%"
        ThisWorkbook.Sheets("Лист8").Cells(row_Итого_по_РОО_Тюменский + i, 3).HorizontalAlignment = xlCenter
        
        ' Ед.изм.
        ThisWorkbook.Sheets("Лист8").Cells(row_Итого_по_РОО_Тюменский + i, 4).Value = ThisWorkbook.Sheets("Лист8").Cells(row_ОО_Тюменский + i, 4).Value
        ThisWorkbook.Sheets("Лист8").Cells(row_Итого_по_РОО_Тюменский + i, 4).HorizontalAlignment = xlCenter
        
        ' Если Ед.изм. не в %, то суммируем по каждому офису
        If ThisWorkbook.Sheets("Лист8").Cells(row_Итого_по_РОО_Тюменский + i, 4).Value <> "%" Then
        
          ' Суммируем:
          ' Квартал - План
          If Not IsEmpty(ThisWorkbook.Sheets("Лист8").Cells(row_ОО_Тюменский + i, 5).Value) Then
            
            ' На пустых возникает ошибка при использовании Round
            ThisWorkbook.Sheets("Лист8").Cells(row_Итого_по_РОО_Тюменский + i, 5).Value = CheckData(ThisWorkbook.Sheets("Лист8").Cells(row_ОО_Тюменский + i, 5).Value) + CheckData(ThisWorkbook.Sheets("Лист8").Cells(row_ОО_Сургутский + i, 5).Value) + CheckData(ThisWorkbook.Sheets("Лист8").Cells(row_ОО_Нижневартовский + i, 5).Value) + CheckData(ThisWorkbook.Sheets("Лист8").Cells(row_ОО_Новоуренгойский + i, 5).Value) + CheckData(ThisWorkbook.Sheets("Лист8").Cells(row_ОО_Тарко_Сале + i, 5).Value)
            
            ' Формат ячейки
            ThisWorkbook.Sheets("Лист8").Cells(row_Итого_по_РОО_Тюменский + i, 5).HorizontalAlignment = xlRight
            ThisWorkbook.Sheets("Лист8").Cells(row_Итого_по_РОО_Тюменский + i, 5).NumberFormat = "#,##0"
          End If
          
          ' Квартал - Факт
          If Not IsEmpty(ThisWorkbook.Sheets("Лист8").Cells(row_ОО_Тюменский + i, 6).Value) Then
            
            ' На пустых возникает ошибка при использовании Round
            ThisWorkbook.Sheets("Лист8").Cells(row_Итого_по_РОО_Тюменский + i, 6).Value = CheckData(ThisWorkbook.Sheets("Лист8").Cells(row_ОО_Тюменский + i, 6).Value) + CheckData(ThisWorkbook.Sheets("Лист8").Cells(row_ОО_Сургутский + i, 6).Value) + CheckData(ThisWorkbook.Sheets("Лист8").Cells(row_ОО_Нижневартовский + i, 6).Value) + CheckData(ThisWorkbook.Sheets("Лист8").Cells(row_ОО_Новоуренгойский + i, 6).Value) + CheckData(ThisWorkbook.Sheets("Лист8").Cells(row_ОО_Тарко_Сале + i, 6).Value)
            
            ' Формат ячейки
            ThisWorkbook.Sheets("Лист8").Cells(row_Итого_по_РОО_Тюменский + i, 6).HorizontalAlignment = xlRight
            ThisWorkbook.Sheets("Лист8").Cells(row_Итого_по_РОО_Тюменский + i, 6).NumberFormat = "#,##0"
            
          End If
        
          ' Квартал - Исп (если План заполнен)
          If Not IsEmpty(ThisWorkbook.Sheets("Лист8").Cells(row_Итого_по_РОО_Тюменский + i, 5).Value) Then
            
            ThisWorkbook.Sheets("Лист8").Cells(row_Итого_по_РОО_Тюменский + i, 7).Value = РассчетДоли(ThisWorkbook.Sheets("Лист8").Cells(row_Итого_по_РОО_Тюменский + i, 5).Value, ThisWorkbook.Sheets("Лист8").Cells(row_Итого_по_РОО_Тюменский + i, 6).Value, 2)
            ' Формат ячейки - %
            ThisWorkbook.Sheets("Лист8").Cells(row_Итого_по_РОО_Тюменский + i, 7).HorizontalAlignment = xlRight
            ThisWorkbook.Sheets("Лист8").Cells(row_Итого_по_РОО_Тюменский + i, 7).NumberFormat = "0%"
          
          End If
        
        
          ' Квартал - Прогноз (если Прогноз заполнен)
          ' Офис 1
          ' План 1
          If ThisWorkbook.Sheets("Лист8").Cells(row_ОО_Тюменский + i, 5).Value <> "" Then
            План_по_продукту_Q_Офис_1 = CLng(ThisWorkbook.Sheets("Лист8").Cells(row_ОО_Тюменский + i, 5).Value)
          Else
            План_по_продукту_Q_Офис_1 = 0
          End If
          ' Прогноз 1 в %
          If ThisWorkbook.Sheets("Лист8").Cells(row_ОО_Тюменский + i, 8).Value <> "" Then
            Прогноз_по_продукту_Q_Офис_1_процент = CDbl(ThisWorkbook.Sheets("Лист8").Cells(row_ОО_Тюменский + i, 8).Value) ' * 100
          Else
            Прогноз_по_продукту_Q_Офис_1_процент = 0
          End If
          ' Прогноз 1 (в шт, тыс.руб.)
          If (План_по_продукту_Q_Офис_1 <> 0) And (Прогноз_по_продукту_Q_Офис_1_процент <> 0) Then
            Прогноз_по_продукту_Q_Офис_1 = План_по_продукту_Q_Офис_1 * Прогноз_по_продукту_Q_Офис_1_процент
          Else
            Прогноз_по_продукту_Q_Офис_1 = 0
          End If
          ' Вариант расчета прогноза "Вариант 1": когда в каждом офисе есть план и факт не равный нулю и "Вариант 2", если план=0,а факт<>0
          If (ThisWorkbook.Sheets("Лист8").Cells(row_ОО_Тюменский + i, 5).Value = 0) And (ThisWorkbook.Sheets("Лист8").Cells(row_ОО_Тюменский + i, 6).Value <> 0) Then
            Вариант_расчета_прогноза = "2"
          End If
        
          ' Офис 2
          ' План 1
          If ThisWorkbook.Sheets("Лист8").Cells(row_ОО_Сургутский + i, 5).Value <> "" Then
            План_по_продукту_Q_Офис_2 = CLng(ThisWorkbook.Sheets("Лист8").Cells(row_ОО_Сургутский + i, 5).Value)
          Else
            План_по_продукту_Q_Офис_2 = 0
          End If
          ' Прогноз 1 в %
          If ThisWorkbook.Sheets("Лист8").Cells(row_ОО_Сургутский + i, 8).Value <> "" Then
            Прогноз_по_продукту_Q_Офис_2_процент = CDbl(ThisWorkbook.Sheets("Лист8").Cells(row_ОО_Сургутский + i, 8).Value) ' * 100
          Else
            Прогноз_по_продукту_Q_Офис_2_процент = 0
          End If
          ' Прогноз 1 (в шт, тыс.руб.)
          If (План_по_продукту_Q_Офис_2 <> 0) And (Прогноз_по_продукту_Q_Офис_2_процент <> 0) Then
            Прогноз_по_продукту_Q_Офис_2 = План_по_продукту_Q_Офис_2 * Прогноз_по_продукту_Q_Офис_2_процент
          Else
            Прогноз_по_продукту_Q_Офис_2 = 0
          End If
          ' Вариант расчета прогноза "Вариант 1": когда в каждом офисе есть план и факт не равный нулю и "Вариант 2", если план=0,а факт<>0
          If (ThisWorkbook.Sheets("Лист8").Cells(row_ОО_Сургутский + i, 5).Value = 0) And (ThisWorkbook.Sheets("Лист8").Cells(row_ОО_Сургутский + i, 6).Value <> 0) Then
            Вариант_расчета_прогноза = "2"
          End If

          ' Офис 3
          ' План 1
          If ThisWorkbook.Sheets("Лист8").Cells(row_ОО_Нижневартовский + i, 5).Value <> "" Then
            План_по_продукту_Q_Офис_3 = CLng(ThisWorkbook.Sheets("Лист8").Cells(row_ОО_Нижневартовский + i, 5).Value)
          Else
            План_по_продукту_Q_Офис_3 = 0
          End If
          ' Прогноз 1 в %
          If ThisWorkbook.Sheets("Лист8").Cells(row_ОО_Нижневартовский + i, 8).Value <> "" Then
            Прогноз_по_продукту_Q_Офис_3_процент = CDbl(ThisWorkbook.Sheets("Лист8").Cells(row_ОО_Нижневартовский + i, 8).Value) ' * 100
          Else
            Прогноз_по_продукту_Q_Офис_3_процент = 0
          End If
          ' Прогноз 1 (в шт, тыс.руб.)
          If (План_по_продукту_Q_Офис_3 <> 0) And (Прогноз_по_продукту_Q_Офис_3_процент <> 0) Then
            Прогноз_по_продукту_Q_Офис_3 = План_по_продукту_Q_Офис_3 * Прогноз_по_продукту_Q_Офис_3_процент
          Else
            Прогноз_по_продукту_Q_Офис_3 = 0
          End If
          ' Вариант расчета прогноза "Вариант 1": когда в каждом офисе есть план и факт не равный нулю и "Вариант 2", если план=0,а факт<>0
          If (ThisWorkbook.Sheets("Лист8").Cells(row_ОО_Нижневартовский + i, 5).Value = 0) And (ThisWorkbook.Sheets("Лист8").Cells(row_ОО_Нижневартовский + i, 6).Value <> 0) Then
            Вариант_расчета_прогноза = "2"
          End If
                
          ' Офис 4
          ' План 1
          If ThisWorkbook.Sheets("Лист8").Cells(row_ОО_Новоуренгойский + i, 5).Value <> "" Then
            План_по_продукту_Q_Офис_4 = CLng(ThisWorkbook.Sheets("Лист8").Cells(row_ОО_Новоуренгойский + i, 5).Value)
          Else
            План_по_продукту_Q_Офис_4 = 0
          End If
          ' Прогноз 1 в %
          If ThisWorkbook.Sheets("Лист8").Cells(row_ОО_Новоуренгойский + i, 8).Value <> "" Then
            Прогноз_по_продукту_Q_Офис_4_процент = CDbl(ThisWorkbook.Sheets("Лист8").Cells(row_ОО_Новоуренгойский + i, 8).Value) ' * 100
          Else
            Прогноз_по_продукту_Q_Офис_4_процент = 0
          End If
          ' Прогноз 1 (в шт, тыс.руб.)
          If (План_по_продукту_Q_Офис_4 <> 0) And (Прогноз_по_продукту_Q_Офис_4_процент <> 0) Then
            Прогноз_по_продукту_Q_Офис_4 = План_по_продукту_Q_Офис_4 * Прогноз_по_продукту_Q_Офис_4_процент
          Else
            Прогноз_по_продукту_Q_Офис_4 = 0
          End If
          ' Вариант расчета прогноза "Вариант 1": когда в каждом офисе есть план и факт не равный нулю и "Вариант 2", если план=0,а факт<>0
          If (ThisWorkbook.Sheets("Лист8").Cells(row_ОО_Новоуренгойский + i, 5).Value = 0) And (ThisWorkbook.Sheets("Лист8").Cells(row_ОО_Новоуренгойский + i, 6).Value <> 0) Then
            Вариант_расчета_прогноза = "2"
          End If
        
          ' Офис 5
          ' План 1
          If ThisWorkbook.Sheets("Лист8").Cells(row_ОО_Тарко_Сале + i, 5).Value <> "" Then
            План_по_продукту_Q_Офис_5 = CLng(ThisWorkbook.Sheets("Лист8").Cells(row_ОО_Тарко_Сале + i, 5).Value)
          Else
            План_по_продукту_Q_Офис_5 = 0
          End If
          ' Прогноз 1 в %
          If ThisWorkbook.Sheets("Лист8").Cells(row_ОО_Тарко_Сале + i, 8).Value <> "" Then
            Прогноз_по_продукту_Q_Офис_5_процент = CDbl(ThisWorkbook.Sheets("Лист8").Cells(row_ОО_Тарко_Сале + i, 8).Value) ' * 100
          Else
            Прогноз_по_продукту_Q_Офис_5_процент = 0
          End If
          ' Прогноз 1 (в шт, тыс.руб.)
          If (План_по_продукту_Q_Офис_5 <> 0) And (Прогноз_по_продукту_Q_Офис_5_процент <> 0) Then
            Прогноз_по_продукту_Q_Офис_5 = План_по_продукту_Q_Офис_5 * Прогноз_по_продукту_Q_Офис_5_процент
          Else
            Прогноз_по_продукту_Q_Офис_5 = 0
          End If
          ' Вариант расчета прогноза "Вариант 1": когда в каждом офисе есть план и факт не равный нулю и "Вариант 2", если план=0,а факт<>0
          If (ThisWorkbook.Sheets("Лист8").Cells(row_ОО_Тарко_Сале + i, 5).Value = 0) And (ThisWorkbook.Sheets("Лист8").Cells(row_ОО_Тарко_Сале + i, 6).Value <> 0) Then
            Вариант_расчета_прогноза = "2"
          End If
   
   
          ' Прогноз по продукту
          Прогноз_по_продукту_Q = Прогноз_по_продукту_Q_Офис_1 + _
                                    Прогноз_по_продукту_Q_Офис_2 + _
                                      Прогноз_по_продукту_Q_Офис_3 + _
                                        Прогноз_по_продукту_Q_Офис_4 + _
                                          Прогноз_по_продукту_Q_Офис_5
          
          ' План по всем офисам
          If ThisWorkbook.Sheets("Лист8").Cells(row_Итого_по_РОО_Тюменский + i, 5).Value <> "" Then
          План_по_всем_офисам_Q = ThisWorkbook.Sheets("Лист8").Cells(row_Итого_по_РОО_Тюменский + i, 5).Value
          Else
            План_по_всем_офисам_Q = 0
          End If
          
          ' Прогноз по продукту
          If (Прогноз_по_продукту_Q <> 0) And (План_по_всем_офисам_Q <> 0) Then
            
            
            ' Вариант 1: когда в каждом офисе есть план и факт не равный нулю
            If Вариант_расчета_прогноза = "1" Then
              ThisWorkbook.Sheets("Лист8").Cells(row_Итого_по_РОО_Тюменский + i, 8).Value = (Прогноз_по_продукту_Q / План_по_всем_офисам_Q) ' * 100
            End If
            
            ' Вариант 2: когда в одном из офисов план=0, но есть факт не равный нулю
            If Вариант_расчета_прогноза = "2" Then
              ThisWorkbook.Sheets("Лист8").Cells(row_Итого_по_РОО_Тюменский + i, 8).Value = Прогноз_квартала_проц(dateDB_Лист8, ThisWorkbook.Sheets("Лист8").Cells(row_Итого_по_РОО_Тюменский + i, 5).Value, ThisWorkbook.Sheets("Лист8").Cells(row_Итого_по_РОО_Тюменский + i, 6).Value, 5, 0) ' (Прогноз_по_продукту_Q / План_по_всем_офисам_Q) ' * 100
            End If
            
            ' Окраска ячейки СФЕТОФОР ' 70. Заливка ячейки цветом "светофор" - значение In_Value в %, In_Target в %
            Call Full_Color_RangeII("Лист8", row_Итого_по_РОО_Тюменский + i, 8, ThisWorkbook.Sheets("Лист8").Cells(row_Итого_по_РОО_Тюменский + i, 8).Value, 1)
            
            
          End If
          ' Формат ячейки - %
          ThisWorkbook.Sheets("Лист8").Cells(row_Итого_по_РОО_Тюменский + i, 8).HorizontalAlignment = xlRight
          ThisWorkbook.Sheets("Лист8").Cells(row_Итого_по_РОО_Тюменский + i, 8).NumberFormat = "0%"

          ' Если прогноза нет, то красим в светофор исполнение (7-ой столбец)
          If (Прогноз_по_продукту_Q = 0) And (План_по_всем_офисам_Q <> 0) Then
            ' Окраска ячейки СФЕТОФОР ' 70. Заливка ячейки цветом "светофор" - значение In_Value в %, In_Target в %
            Call Full_Color_RangeII("Лист8", row_Итого_по_РОО_Тюменский + i, 7, ThisWorkbook.Sheets("Лист8").Cells(row_Итого_по_РОО_Тюменский + i, 7).Value, 1)
          End If

          ' Месяц - План
          If Not IsEmpty(ThisWorkbook.Sheets("Лист8").Cells(row_ОО_Тюменский + i, 9).Value) Then
          ' If ThisWorkbook.Sheets("Лист8").Cells(row_ОО_Тюменский + i, 9).Value <> "" Then
            
            ' Проверяем наличие данных в ячейках. Бывает, что у части офисов отсутствует план, факт. При этом у некоторых есть - пример: Dashboard_new_РБ_05.12.2021
            ' If ThisWorkbook.Sheets("Лист8").Cells(row_ОО_Тюменский + i, 9).Value <> "" Then var_1_9 = ThisWorkbook.Sheets("Лист8").Cells(row_ОО_Тюменский + i, 9).Value Else var_1_9 = 0
            ' If ThisWorkbook.Sheets("Лист8").Cells(row_ОО_Сургутский + i, 9).Value <> "" Then var_2_9 = ThisWorkbook.Sheets("Лист8").Cells(row_ОО_Сургутский + i, 9).Value Else var_2_9 = 0
            ' If ThisWorkbook.Sheets("Лист8").Cells(row_ОО_Нижневартовский + i, 9).Value <> "" Then var_3_9 = ThisWorkbook.Sheets("Лист8").Cells(row_ОО_Нижневартовский + i, 9).Value Else var_3_9 = 0
            ' If ThisWorkbook.Sheets("Лист8").Cells(row_ОО_Новоуренгойский + i, 9).Value <> "" Then var_4_9 = ThisWorkbook.Sheets("Лист8").Cells(row_ОО_Новоуренгойский + i, 9).Value Else var_4_9 = 0
            ' If ThisWorkbook.Sheets("Лист8").Cells(row_ОО_Тарко_Сале + i, 9).Value <> "" Then var_5_9 = ThisWorkbook.Sheets("Лист8").Cells(row_ОО_Тарко_Сале + i, 9).Value Else var_5_9 = 0
            ' Суммируем проверенные на null значения
            ' ThisWorkbook.Sheets("Лист8").Cells(row_Итого_по_РОО_Тюменский + i, 9).Value = var_1_9 + var_2_9 + var_3_9 + var_4_9 + var_5_9
            
            ThisWorkbook.Sheets("Лист8").Cells(row_Итого_по_РОО_Тюменский + i, 9).Value = ThisWorkbook.Sheets("Лист8").Cells(row_ОО_Тюменский + i, 9).Value + ThisWorkbook.Sheets("Лист8").Cells(row_ОО_Сургутский + i, 9).Value + ThisWorkbook.Sheets("Лист8").Cells(row_ОО_Нижневартовский + i, 9).Value + ThisWorkbook.Sheets("Лист8").Cells(row_ОО_Новоуренгойский + i, 9).Value + ThisWorkbook.Sheets("Лист8").Cells(row_ОО_Тарко_Сале + i, 9).Value
            
            
            ' Формат ячейки
            ThisWorkbook.Sheets("Лист8").Cells(row_Итого_по_РОО_Тюменский + i, 9).HorizontalAlignment = xlRight
            ThisWorkbook.Sheets("Лист8").Cells(row_Итого_по_РОО_Тюменский + i, 9).NumberFormat = "#,##0"
          End If
          
          ' Месяц - Факт
          If Not IsEmpty(ThisWorkbook.Sheets("Лист8").Cells(row_ОО_Тюменский + i, 10).Value) Then
            ThisWorkbook.Sheets("Лист8").Cells(row_Итого_по_РОО_Тюменский + i, 10).Value = ThisWorkbook.Sheets("Лист8").Cells(row_ОО_Тюменский + i, 10).Value + ThisWorkbook.Sheets("Лист8").Cells(row_ОО_Сургутский + i, 10).Value + ThisWorkbook.Sheets("Лист8").Cells(row_ОО_Нижневартовский + i, 10).Value + ThisWorkbook.Sheets("Лист8").Cells(row_ОО_Новоуренгойский + i, 10).Value + ThisWorkbook.Sheets("Лист8").Cells(row_ОО_Тарко_Сале + i, 10).Value
            ' Формат ячейки
            ThisWorkbook.Sheets("Лист8").Cells(row_Итого_по_РОО_Тюменский + i, 10).HorizontalAlignment = xlRight
            ThisWorkbook.Sheets("Лист8").Cells(row_Итого_по_РОО_Тюменский + i, 10).NumberFormat = "#,##0"
          End If
        
          ' Месяц - Исп (если План заполнен)
          If Not IsEmpty(ThisWorkbook.Sheets("Лист8").Cells(row_Итого_по_РОО_Тюменский + i, 9).Value) Then
        
            ThisWorkbook.Sheets("Лист8").Cells(row_Итого_по_РОО_Тюменский + i, 11).Value = РассчетДоли(ThisWorkbook.Sheets("Лист8").Cells(row_Итого_по_РОО_Тюменский + i, 9).Value, ThisWorkbook.Sheets("Лист8").Cells(row_Итого_по_РОО_Тюменский + i, 10).Value, 2)
            ' Формат ячейки - %
            ThisWorkbook.Sheets("Лист8").Cells(row_Итого_по_РОО_Тюменский + i, 11).HorizontalAlignment = xlRight
            ThisWorkbook.Sheets("Лист8").Cells(row_Итого_по_РОО_Тюменский + i, 11).NumberFormat = "0%"

          End If
        
          ' Месяц - Прогноз (если Прогноз заполнен)
     
          ' Офис 1
          ' План 1
          If ThisWorkbook.Sheets("Лист8").Cells(row_ОО_Тюменский + i, 9).Value <> "" Then
            План_по_продукту_M_Офис_1 = CLng(ThisWorkbook.Sheets("Лист8").Cells(row_ОО_Тюменский + i, 9).Value)
          Else
            План_по_продукту_M_Офис_1 = 0
          End If
          ' Прогноз 1 в %
          If ThisWorkbook.Sheets("Лист8").Cells(row_ОО_Тюменский + i, 12).Value <> "" Then
            Прогноз_по_продукту_M_Офис_1_процент = CDbl(ThisWorkbook.Sheets("Лист8").Cells(row_ОО_Тюменский + i, 12).Value) ' * 100
          Else
            Прогноз_по_продукту_M_Офис_1_процент = 0
          End If
          ' Прогноз 1 (в шт, тыс.руб.)
          If (План_по_продукту_M_Офис_1 <> 0) And (Прогноз_по_продукту_M_Офис_1_процент <> 0) Then
            Прогноз_по_продукту_M_Офис_1 = План_по_продукту_M_Офис_1 * Прогноз_по_продукту_M_Офис_1_процент
          Else
            Прогноз_по_продукту_M_Офис_1 = 0
          End If
          ' Месяц: Вариант расчета прогноза "Вариант 1": когда в каждом офисе есть план и факт не равный нулю и "Вариант 2", если план=0,а факт<>0
          If (ThisWorkbook.Sheets("Лист8").Cells(row_ОО_Тюменский + i, 9).Value = 0) And (ThisWorkbook.Sheets("Лист8").Cells(row_ОО_Тюменский + i, 10).Value <> 0) Then
            Вариант_расчета_прогноза_М = "2"
          End If
        
        
          ' Офис 2
          ' План 1
          If ThisWorkbook.Sheets("Лист8").Cells(row_ОО_Сургутский + i, 9).Value <> "" Then
            План_по_продукту_M_Офис_2 = CLng(ThisWorkbook.Sheets("Лист8").Cells(row_ОО_Сургутский + i, 9).Value)
          Else
            План_по_продукту_M_Офис_2 = 0
          End If
          ' Прогноз 1 в %
          If ThisWorkbook.Sheets("Лист8").Cells(row_ОО_Сургутский + i, 12).Value <> "" Then
            Прогноз_по_продукту_M_Офис_2_процент = CDbl(ThisWorkbook.Sheets("Лист8").Cells(row_ОО_Сургутский + i, 12).Value) ' * 100
          Else
            Прогноз_по_продукту_M_Офис_2_процент = 0
          End If
          ' Прогноз 1 (в шт, тыс.руб.)
          If (План_по_продукту_M_Офис_2 <> 0) And (Прогноз_по_продукту_M_Офис_2_процент <> 0) Then
            Прогноз_по_продукту_M_Офис_2 = План_по_продукту_M_Офис_2 * Прогноз_по_продукту_M_Офис_2_процент
          Else
            Прогноз_по_продукту_M_Офис_2 = 0
          End If
          ' Месяц: Вариант расчета прогноза "Вариант 1": когда в каждом офисе есть план и факт не равный нулю и "Вариант 2", если план=0,а факт<>0
          If (ThisWorkbook.Sheets("Лист8").Cells(row_ОО_Сургутский + i, 9).Value = 0) And (ThisWorkbook.Sheets("Лист8").Cells(row_ОО_Сургутский + i, 10).Value <> 0) Then
            Вариант_расчета_прогноза_М = "2"
          End If

          ' Офис 3
          ' План 1
          If ThisWorkbook.Sheets("Лист8").Cells(row_ОО_Нижневартовский + i, 9).Value <> "" Then
            План_по_продукту_M_Офис_3 = CLng(ThisWorkbook.Sheets("Лист8").Cells(row_ОО_Нижневартовский + i, 9).Value)
          Else
            План_по_продукту_M_Офис_3 = 0
          End If
          ' Прогноз 1 в %
          If ThisWorkbook.Sheets("Лист8").Cells(row_ОО_Нижневартовский + i, 12).Value <> "" Then
            Прогноз_по_продукту_M_Офис_3_процент = CDbl(ThisWorkbook.Sheets("Лист8").Cells(row_ОО_Нижневартовский + i, 12).Value) ' * 100
          Else
            Прогноз_по_продукту_M_Офис_3_процент = 0
          End If
          ' Прогноз 1 (в шт, тыс.руб.)
          If (План_по_продукту_M_Офис_3 <> 0) And (Прогноз_по_продукту_M_Офис_3_процент <> 0) Then
            Прогноз_по_продукту_M_Офис_3 = План_по_продукту_M_Офис_3 * Прогноз_по_продукту_M_Офис_3_процент
          Else
            Прогноз_по_продукту_M_Офис_3 = 0
          End If
          ' Месяц: Вариант расчета прогноза "Вариант 1": когда в каждом офисе есть план и факт не равный нулю и "Вариант 2", если план=0,а факт<>0
          If (ThisWorkbook.Sheets("Лист8").Cells(row_ОО_Нижневартовский + i, 9).Value = 0) And (ThisWorkbook.Sheets("Лист8").Cells(row_ОО_Нижневартовский + i, 10).Value <> 0) Then
            Вариант_расчета_прогноза_М = "2"
          End If
        
        
          ' Офис 4
          ' План 1
          If ThisWorkbook.Sheets("Лист8").Cells(row_ОО_Новоуренгойский + i, 9).Value <> "" Then
            План_по_продукту_M_Офис_4 = CLng(ThisWorkbook.Sheets("Лист8").Cells(row_ОО_Новоуренгойский + i, 9).Value)
          Else
            План_по_продукту_M_Офис_4 = 0
          End If
          ' Прогноз 1 в %
          If ThisWorkbook.Sheets("Лист8").Cells(row_ОО_Новоуренгойский + i, 12).Value <> "" Then
            Прогноз_по_продукту_M_Офис_4_процент = CDbl(ThisWorkbook.Sheets("Лист8").Cells(row_ОО_Новоуренгойский + i, 12).Value) ' * 100
          Else
            Прогноз_по_продукту_M_Офис_4_процент = 0
          End If
          ' Прогноз 1 (в шт, тыс.руб.)
          If (План_по_продукту_M_Офис_4 <> 0) And (Прогноз_по_продукту_M_Офис_4_процент <> 0) Then
            Прогноз_по_продукту_M_Офис_4 = План_по_продукту_M_Офис_4 * Прогноз_по_продукту_M_Офис_4_процент
          Else
            Прогноз_по_продукту_M_Офис_4 = 0
          End If
          ' Месяц: Вариант расчета прогноза "Вариант 1": когда в каждом офисе есть план и факт не равный нулю и "Вариант 2", если план=0,а факт<>0
          If (ThisWorkbook.Sheets("Лист8").Cells(row_ОО_Новоуренгойский + i, 9).Value = 0) And (ThisWorkbook.Sheets("Лист8").Cells(row_ОО_Новоуренгойский + i, 10).Value <> 0) Then
            Вариант_расчета_прогноза_М = "2"
          End If

        
          ' Офис 5
          ' План 1
          If ThisWorkbook.Sheets("Лист8").Cells(row_ОО_Тарко_Сале + i, 9).Value <> "" Then
            План_по_продукту_M_Офис_5 = CLng(ThisWorkbook.Sheets("Лист8").Cells(row_ОО_Тарко_Сале + i, 9).Value)
          Else
            План_по_продукту_M_Офис_5 = 0
          End If
          ' Прогноз 1 в %
          If ThisWorkbook.Sheets("Лист8").Cells(row_ОО_Тарко_Сале + i, 12).Value <> "" Then
            Прогноз_по_продукту_M_Офис_5_процент = CDbl(ThisWorkbook.Sheets("Лист8").Cells(row_ОО_Тарко_Сале + i, 12).Value) ' * 100
          Else
            Прогноз_по_продукту_M_Офис_5_процент = 0
          End If
          ' Прогноз 1 (в шт, тыс.руб.)
          If (План_по_продукту_M_Офис_5 <> 0) And (Прогноз_по_продукту_M_Офис_5_процент <> 0) Then
            Прогноз_по_продукту_M_Офис_5 = План_по_продукту_M_Офис_5 * Прогноз_по_продукту_M_Офис_5_процент
          Else
            Прогноз_по_продукту_M_Офис_5 = 0
          End If
          ' Месяц: Вариант расчета прогноза "Вариант 1": когда в каждом офисе есть план и факт не равный нулю и "Вариант 2", если план=0,а факт<>0
          If (ThisWorkbook.Sheets("Лист8").Cells(row_ОО_Тарко_Сале + i, 9).Value = 0) And (ThisWorkbook.Sheets("Лист8").Cells(row_ОО_Тарко_Сале + i, 10).Value <> 0) Then
            Вариант_расчета_прогноза_М = "2"
          End If
   
   
          ' Прогноз по продукту
          Прогноз_по_продукту_M = Прогноз_по_продукту_M_Офис_1 + _
                                    Прогноз_по_продукту_M_Офис_2 + _
                                      Прогноз_по_продукту_M_Офис_3 + _
                                        Прогноз_по_продукту_M_Офис_4 + _
                                          Прогноз_по_продукту_M_Офис_5
          
          ' План по всем офисам
          If ThisWorkbook.Sheets("Лист8").Cells(row_Итого_по_РОО_Тюменский + i, 9).Value <> "" Then
            План_по_всем_офисам_М = ThisWorkbook.Sheets("Лист8").Cells(row_Итого_по_РОО_Тюменский + i, 9).Value
          Else
            План_по_всем_офисам_М = 0
          End If
          
          If (Прогноз_по_продукту_M <> 0) And (План_по_всем_офисам_М <> 0) Then
            
            ' Вариант 1: когда в каждом офисе есть план и факт не равный нулю
            If Вариант_расчета_прогноза_М = "1" Then
              ThisWorkbook.Sheets("Лист8").Cells(row_Итого_по_РОО_Тюменский + i, 12).Value = (Прогноз_по_продукту_M / План_по_всем_офисам_М) ' * 100
            End If
            
            ' Вариант 2: когда в одном из офисов план=0, но есть факт не равный нулю
            If Вариант_расчета_прогноза_М = "2" Then
              ThisWorkbook.Sheets("Лист8").Cells(row_Итого_по_РОО_Тюменский + i, 12).Value = Прогноз_месяца_проц(dateDB_Лист8, ThisWorkbook.Sheets("Лист8").Cells(row_Итого_по_РОО_Тюменский + i, 9).Value, ThisWorkbook.Sheets("Лист8").Cells(row_Итого_по_РОО_Тюменский + i, 6).Value, 10, 0)
            End If
            
            ' Окраска ячейки СФЕТОФОР ' 70. Заливка ячейки цветом "светофор" - значение In_Value в %, In_Target в %
            Call Full_Color_RangeII("Лист8", row_Итого_по_РОО_Тюменский + i, 12, ThisWorkbook.Sheets("Лист8").Cells(row_Итого_по_РОО_Тюменский + i, 12).Value, 1)
          
          End If
        
        
          ' Если столбца "Прогноз" нет (In_DeltaPrediction = 0), то Окраска ячейки СФЕТОФОР ' 70. Заливка ячейки цветом "светофор" - значение In_Value в %, In_Target в %
          If (Прогноз_по_продукту_M = 0) And (План_по_всем_офисам_М <> 0) Then
            ' Окраска ячейки СФЕТОФОР ' 70. Заливка ячейки цветом "светофор" - значение In_Value в %, In_Target в %
            Call Full_Color_RangeII("Лист8", row_Итого_по_РОО_Тюменский + i, 11, ThisWorkbook.Sheets("Лист8").Cells(row_Итого_по_РОО_Тюменский + i, 11).Value, 1)
          End If

        
          ' Формат ячейки - %
          ThisWorkbook.Sheets("Лист8").Cells(row_Итого_по_РОО_Тюменский + i, 12).HorizontalAlignment = xlRight
          ThisWorkbook.Sheets("Лист8").Cells(row_Итого_по_РОО_Тюменский + i, 12).NumberFormat = "0%"
    
          ' ***
          ' Тестирование Функции "Прогноз_квартала" по всем позициям, если измерение не в %
          ' If ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 4).Value <> "%" Then
          
            t = row_Итого_по_РОО_Тюменский + i
            t2 = dateDB
            t3 = ThisWorkbook.Sheets("Лист8").Cells(row_Итого_по_РОО_Тюменский + i, 5).Value
            t4 = ThisWorkbook.Sheets("Лист8").Cells(row_Итого_по_РОО_Тюменский + i, 6).Value
            t5 = 0
          
            ThisWorkbook.Sheets("Лист8").Cells(row_Итого_по_РОО_Тюменский + i, 20).Value = Прогноз_квартала_проц(dateDB, ThisWorkbook.Sheets("Лист8").Cells(row_Итого_по_РОО_Тюменский + i, 5).Value, ThisWorkbook.Sheets("Лист8").Cells(row_Итого_по_РОО_Тюменский + i, 6).Value, 5, 0)
            ThisWorkbook.Sheets("Лист8").Cells(row_Итого_по_РОО_Тюменский + i, 20).NumberFormat = "0%"
            ThisWorkbook.Sheets("Лист8").Cells(row_Итого_по_РОО_Тюменский + i, 20).HorizontalAlignment = xlRight
            
            ' Делаем рассчет показателей на неделю, если данный показатель не в процентах
            If ThisWorkbook.Sheets("Лист8").Cells(row_Итого_по_РОО_Тюменский + i, 4).Value <> "%" Then
              
              ' === Сюда вставляем цель на неделю - сколько надо прирасти, чтобы выйти на прогноз Q в 100%
              If Прогноз_квартала_проц(dateDB, ThisWorkbook.Sheets("Лист8").Cells(row_Итого_по_РОО_Тюменский + i, 5).Value, ThisWorkbook.Sheets("Лист8").Cells(row_Итого_по_РОО_Тюменский + i, 6).Value, 5, 0) < 1 Then
            
                ' Считаем какой должен быть прогноз
                Факт_на_дату_для_прогноза_квартала_Var = Факт_на_дату_для_прогноза_квартала(dateDB + 7, ThisWorkbook.Sheets("Лист8").Cells(row_Итого_по_РОО_Тюменский + i, 5).Value, 1, 5, 0)
          
                ' Если Факт для выхода на прогноз Q больше, чем текущий Факт Q, то считаем прирост
                If Факт_на_дату_для_прогноза_квартала_Var > ThisWorkbook.Sheets("Лист8").Cells(row_Итого_по_РОО_Тюменский + i, 6).Value Then
                  
                  ThisWorkbook.Sheets("Лист8").Cells(row_Итого_по_РОО_Тюменский + i, 13).Value = Факт_на_дату_для_прогноза_квартала_Var - ThisWorkbook.Sheets("Лист8").Cells(row_Итого_по_РОО_Тюменский + i, 6).Value
                  ThisWorkbook.Sheets("Лист8").Cells(row_Итого_по_РОО_Тюменский + i, 13).NumberFormat = "#,##0"
                  ThisWorkbook.Sheets("Лист8").Cells(row_Итого_по_РОО_Тюменский + i, 13).HorizontalAlignment = xlRight
                  
                  '
                  ' В 14-ый столбец пишем исполнение Плана недели (из 13-го столбца)
                  ThisWorkbook.Sheets("Лист8").Cells(row_Итого_по_РОО_Тюменский + i, 14).Value = CheckData(ThisWorkbook.Sheets("Лист8").Cells(getRowFromSheet8("ОО «Тюменский»", ThisWorkbook.Sheets("Лист8").Cells(row_Итого_по_РОО_Тюменский + i, 2).Value), 14).Value) + _
                                                                                 CheckData(ThisWorkbook.Sheets("Лист8").Cells(getRowFromSheet8("ОО «Сургутский»", ThisWorkbook.Sheets("Лист8").Cells(row_Итого_по_РОО_Тюменский + i, 2).Value), 14).Value) + _
                                                                                   CheckData(ThisWorkbook.Sheets("Лист8").Cells(getRowFromSheet8("ОО «Нижневартовский»", ThisWorkbook.Sheets("Лист8").Cells(row_Итого_по_РОО_Тюменский + i, 2).Value), 14).Value) + _
                                                                                     CheckData(ThisWorkbook.Sheets("Лист8").Cells(getRowFromSheet8("ОО «Новоуренгойский»", ThisWorkbook.Sheets("Лист8").Cells(row_Итого_по_РОО_Тюменский + i, 2).Value), 14).Value) + _
                                                                                       CheckData(ThisWorkbook.Sheets("Лист8").Cells(getRowFromSheet8("ОО «Тарко-Сале»", ThisWorkbook.Sheets("Лист8").Cells(row_Итого_по_РОО_Тюменский + i, 2).Value), 14).Value)
                  ThisWorkbook.Sheets("Лист8").Cells(row_Итого_по_РОО_Тюменский + i, 14).NumberFormat = "#,##0"
                  ThisWorkbook.Sheets("Лист8").Cells(row_Итого_по_РОО_Тюменский + i, 14).HorizontalAlignment = xlRight

                  
                End If
            
              End If
              ' ===
              
            End If
          
          ' ***
        
        
        Else
          
          ' Если текущий показатель измеряется в %, то берем цифру из Тюмерский РОО в DB
          ' Это наименование показателя ThisWorkbook.Sheets("Лист8").Cells(row_ОО_Тюменский + i, 2).Value
          ' Его берем сумму для 5 офисов по кварталу, делим на 5 и вставляем в %
          
          ' % Квартал (план)
          If ThisWorkbook.Sheets("Лист8").Cells(getRowFromSheet8("ОО «Тюменский»", ThisWorkbook.Sheets("Лист8").Cells(row_ОО_Тюменский + i, 2).Value), 5).Value <> "" Then
            ThisWorkbook.Sheets("Лист8").Cells(row_Итого_по_РОО_Тюменский + i, 5).Value = ThisWorkbook.Sheets("Лист8").Cells(getRowFromSheet8("ОО «Тюменский»", ThisWorkbook.Sheets("Лист8").Cells(row_ОО_Тюменский + i, 2).Value), 5).Value
            ' Формат
            ThisWorkbook.Sheets("Лист8").Cells(row_Итого_по_РОО_Тюменский + i, 5).NumberFormat = "#,##0"
            ThisWorkbook.Sheets("Лист8").Cells(row_Итого_по_РОО_Тюменский + i, 5).HorizontalAlignment = xlRight
          End If
          
          ' % Месяц (план)
          If ThisWorkbook.Sheets("Лист8").Cells(getRowFromSheet8("ОО «Тюменский»", ThisWorkbook.Sheets("Лист8").Cells(row_ОО_Тюменский + i, 2).Value), 9).Value <> "" Then
            ThisWorkbook.Sheets("Лист8").Cells(row_Итого_по_РОО_Тюменский + i, 9).Value = ThisWorkbook.Sheets("Лист8").Cells(getRowFromSheet8("ОО «Тюменский»", ThisWorkbook.Sheets("Лист8").Cells(row_ОО_Тюменский + i, 2).Value), 9).Value
            ' Формат
            ThisWorkbook.Sheets("Лист8").Cells(row_Итого_по_РОО_Тюменский + i, 9).NumberFormat = "#,##0"
            ThisWorkbook.Sheets("Лист8").Cells(row_Итого_по_РОО_Тюменский + i, 9).HorizontalAlignment = xlRight
          End If
          
          ' % Квартал (факт)
          Офис_1_Q = ThisWorkbook.Sheets("Лист8").Cells(getRowFromSheet8("ОО «Тюменский»", ThisWorkbook.Sheets("Лист8").Cells(row_ОО_Тюменский + i, 2).Value), 6).Value
          If Офис_1_Q = "" Then
            Офис_1_Q = 0
          End If
          
          Офис_2_Q = ThisWorkbook.Sheets("Лист8").Cells(getRowFromSheet8("ОО «Сургутский»", ThisWorkbook.Sheets("Лист8").Cells(row_ОО_Тюменский + i, 2).Value), 6).Value
          If Офис_2_Q = "" Then
            Офис_2_Q = 0
          End If
          
          Офис_3_Q = ThisWorkbook.Sheets("Лист8").Cells(getRowFromSheet8("ОО «Нижневартовский»", ThisWorkbook.Sheets("Лист8").Cells(row_ОО_Тюменский + i, 2).Value), 6).Value
          If Офис_3_Q = "" Then
            Офис_3_Q = 0
          End If
          
          Офис_4_Q = ThisWorkbook.Sheets("Лист8").Cells(getRowFromSheet8("ОО «Новоуренгойский»", ThisWorkbook.Sheets("Лист8").Cells(row_ОО_Тюменский + i, 2).Value), 6).Value
          If Офис_4_Q = "" Then
            Офис_4_Q = 0
          End If
          
          Офис_5_Q = ThisWorkbook.Sheets("Лист8").Cells(getRowFromSheet8("ОО «Тарко-Сале»", ThisWorkbook.Sheets("Лист8").Cells(row_ОО_Тюменский + i, 2).Value), 6).Value
          If Офис_5_Q = "" Then
            Офис_5_Q = 0
          End If
          
          ThisWorkbook.Sheets("Лист8").Cells(row_Итого_по_РОО_Тюменский + i, 6).Value = (Офис_1_Q + Офис_2_Q + Офис_3_Q + Офис_4_Q + Офис_5_Q) / 5
          ' Формат
          ThisWorkbook.Sheets("Лист8").Cells(row_Итого_по_РОО_Тюменский + i, 6).NumberFormat = "#,##0"
          ThisWorkbook.Sheets("Лист8").Cells(row_Итого_по_РОО_Тюменский + i, 6).HorizontalAlignment = xlRight
          
          ' % Месяц (факт)
          Офис_1_M = ThisWorkbook.Sheets("Лист8").Cells(getRowFromSheet8("ОО «Тюменский»", ThisWorkbook.Sheets("Лист8").Cells(row_ОО_Тюменский + i, 2).Value), 10).Value
          If Офис_1_M = "" Then
            Офис_1_M = 0
          End If
          
          Офис_2_M = ThisWorkbook.Sheets("Лист8").Cells(getRowFromSheet8("ОО «Сургутский»", ThisWorkbook.Sheets("Лист8").Cells(row_ОО_Тюменский + i, 2).Value), 10).Value
          If Офис_2_M = "" Then
            Офис_2_M = 0
          End If
          
          Офис_3_M = ThisWorkbook.Sheets("Лист8").Cells(getRowFromSheet8("ОО «Нижневартовский»", ThisWorkbook.Sheets("Лист8").Cells(row_ОО_Тюменский + i, 2).Value), 10).Value
          If Офис_3_M = "" Then
            Офис_3_M = 0
          End If
          
          Офис_4_M = ThisWorkbook.Sheets("Лист8").Cells(getRowFromSheet8("ОО «Новоуренгойский»", ThisWorkbook.Sheets("Лист8").Cells(row_ОО_Тюменский + i, 2).Value), 10).Value
          If Офис_4_M = "" Then
            Офис_4_M = 0
          End If
          
          Офис_5_M = ThisWorkbook.Sheets("Лист8").Cells(getRowFromSheet8("ОО «Тарко-Сале»", ThisWorkbook.Sheets("Лист8").Cells(row_ОО_Тюменский + i, 2).Value), 10).Value
          If Офис_5_M = "" Then
            Офис_5_M = 0
          End If
          
          ThisWorkbook.Sheets("Лист8").Cells(row_Итого_по_РОО_Тюменский + i, 10).Value = (Офис_1_M + Офис_2_M + Офис_3_M + Офис_4_M + Офис_5_M) / 5
          ' Формат
          ThisWorkbook.Sheets("Лист8").Cells(row_Итого_по_РОО_Тюменский + i, 10).NumberFormat = "#,##0"
          ThisWorkbook.Sheets("Лист8").Cells(row_Итого_по_РОО_Тюменский + i, 10).HorizontalAlignment = xlRight
          
        End If ' Если Ед.изм. не в %, то суммируем по каждому офису
        
      Next i



  ' Чертим горизонтальные линии
  

  ' ----------------------------------------------------------------------------------------------------------------------------------
  ' Чертим горизонтальную линию 2 (указываем предидущее значение строки + 1)
  Call gorizontalLineII(ThisWorkbook.Name, "Лист8", getRowFromSheet8("Итого по РОО «Тюменский»", "Портфель ЗП 18+") + 1, 2, 12)
  Call gorizontalLineII(ThisWorkbook.Name, "Лист8", getRowFromSheet8("Итого по РОО «Тюменский»", "Утилизация лимита ПК") + 1, 2, 12)
  Call gorizontalLineII(ThisWorkbook.Name, "Лист8", getRowFromSheet8("Итого по РОО «Тюменский»", "КК кол-во выданных") + 1, 2, 12)
  Call gorizontalLineII(ThisWorkbook.Name, "Лист8", getRowFromSheet8("Итого по РОО «Тюменский»", "ДК кол-во выданных") + 1, 2, 12)
  Call gorizontalLineII(ThisWorkbook.Name, "Лист8", getRowFromSheet8("Итого по РОО «Тюменский»", "Накопительные счета") + 1, 2, 12)
  Call gorizontalLineII(ThisWorkbook.Name, "Лист8", getRowFromSheet8("Итого по РОО «Тюменский»", "Orange Premium Club") + 1, 2, 12)
  Call gorizontalLineII(ThisWorkbook.Name, "Лист8", getRowFromSheet8("Итого по РОО «Тюменский»", "Ядро клиентов") + 1, 2, 12)
  
  ' If getRowFromSheet8("Итого по РОО «Тюменский»", "           УК (Affluent)") <> 0 Then
  '   Call gorizontalLineII(ThisWorkbook.Name, "Лист8", getRowFromSheet8("Итого по РОО «Тюменский»", "           УК (Affluent)") + 1, 2, 12)
  ' End If
  
  Call gorizontalLineII(ThisWorkbook.Name, "Лист8", getRowFromSheet8("Итого по РОО «Тюменский»", "           УК (Affluent)") + 1, 2, 12)
  
  Call gorizontalLineII(ThisWorkbook.Name, "Лист8", getRowFromSheet8("Итого по РОО «Тюменский»", "           КЛА премия OPC") + 1, 2, 12)
  Call gorizontalLineII(ThisWorkbook.Name, "Лист8", getRowFromSheet8("Итого по РОО «Тюменский»", "Закрытые вклады OPC") + 1, 2, 12)
  Call gorizontalLineII(ThisWorkbook.Name, "Лист8", getRowFromSheet8("Итого по РОО «Тюменский»", "           АУМ") + 1, 2, 12)
  Call gorizontalLineII(ThisWorkbook.Name, "Лист8", getRowFromSheet8("Итого по РОО «Тюменский»", "ОФЗ (в т.ч.OPC)") + 1, 2, 12)
  
  ' ----------------------------------------------------------------------------------------------------------------------------------

  ' Считаем Цель на неделю
  

End Sub


' Отправка письма: отправляю шаблон самому себе для последующей отправки в сеть письма на его основе:
Sub Отправка_Lotus_Notes_Выгр_день_Лист8(In_attachmentFile, In_DateFile)
Dim темаПисьма, текстПисьма, hashTag, attachmentFile As String
Dim i As Byte
  
  
  ' Запрос
  If MsgBox("Отправить себе Шаблон письма с планами на день?", vbYesNo) = vbYes Then
    
    ' Тема письма - Тема:
    ' темаПисьма = ThisWorkbook.Sheets("Лист8").Cells(RowByValue(ThisWorkbook.Name, "Лист8", "Тема:", 100, 100), ColumnByValue(ThisWorkbook.Name, "Лист8", "Тема:", 100, 100) + 1).Value
    темаПисьма = "Планы продаж на " + CStr(In_DateFile)

    ' hashTag - Хэштэг:
    ' hashTag = ThisWorkbook.Sheets("Лист8").Cells(RowByValue(ThisWorkbook.Name, "Лист8", "Хэштэг:", 100, 100), ColumnByValue(ThisWorkbook.Name, "Лист8", "Хэштэг:", 100, 100) + 1).Value
    ' hashTag - Хэштэг:
    hashTag = "#ежедневныйотчет"

    ' Файл-вложение (!!!)
    attachmentFile = In_attachmentFile
    
    ' Текст письма
    текстПисьма = "" + Chr(13)
    текстПисьма = текстПисьма + "" + getFromAddrBook("УДО2,УДО3,УДО4,УДО5,НОРПиКО1,НОРПиКО2,НОРПиКО3,НОРПиКО4,НОРПиКО5", 2) + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "" + getFromAddrBook("РД", 2) + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "Уважаемые руководители," + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "План продаж на " + CStr(In_DateFile) + Chr(13)
    ' текстПисьма = текстПисьма + "" + Chr(13)
    ' текстПисьма = текстПисьма + "" + Chr(13)
    ' текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    ' Визитка (подпись С Ув., )
    текстПисьма = текстПисьма + ПодписьВПисьме()
    ' Хэштег
    текстПисьма = текстПисьма + createBlankStr(27) + hashTag
    ' Вызов
    Call send_Lotus_Notes(темаПисьма, "Sergey Fedorovich Proschaev/Tyumen/PSBank/Ru", "Sergey Fedorovich Proschaev/Tyumen/PSBank/Ru", текстПисьма, attachmentFile)
  
    ' Сообщение
    MsgBox ("Письмо отправлено!")
     
  End If
  
End Sub

' Загрузить факт на дату с диалогами
Sub Загрузить_факт_на_дату_диалог()
      
  ' Запрос
  If MsgBox("Загрузить факт на " + CStr(ThisWorkbook.Sheets("Лист8").Range("O9").Value) + "?", vbYesNo) = vbYes Then
      
    ' Обнуление итоговых значений по Офисам и РОО
    Call Обнуление_итоговых_значений_по_Офисам_и_РОО
      
    ' Открываем BASE\Sales
    OpenBookInBase ("Sales_Office")
  
    ' Вызов Месяц
    ' Call Загрузить_факт_на_дату
        
    ' Вызов Q
    Call Загрузить_факт_на_дату2
        
        
    ' Закрываем BASE\Sales
    CloseBook ("Sales_Office")
    
    ' Сообщение
    MsgBox ("Данные загружены!")
    
  End If

End Sub

' Загрузить факт на дату
Sub Загрузить_факт_на_дату()
Dim dateForLoad As Date
  
  ' Дата для загрузки данных
  dateForLoad = ThisWorkbook.Sheets("Лист8").Range("O9").Value
  ' Дата месяца ММГГ
  strMMYYVar = strMMYY(dateForLoad)
  ' Дата Q
  strNQYYVar = strNQYY(dateDB)
  
    ' Эти дни считаем для ПК у МРК
    ' Число отработанных дней WorkDayLeft для расчета прогноза в офисном канале (с учетом рабочих дней)
    WorkDayLeft = Working_days_between_datesII(monthStartDate(dateForLoad), dateForLoad, 5)
    ' Число оставшихся дней WorkDayRight для расчета прогноза в офисном канале ((с учетом рабочих дней))
    WorkDayRight = Working_days_between_datesII(dateForLoad + 1, Date_last_day_month(dateForLoad), 5)
          
    ' Эти дни считаем для ПК
    ' Число отработанных дней WorkDayLeft для расчета прогноза в офисном канале (с учетом рабочих дней)
    WorkDayLeftCalendar = Working_days_between_dates(monthStartDate(dateForLoad), dateForLoad, 5)
    ' Число оставшихся дней WorkDayRight для расчета прогноза в офисном канале ((с учетом рабочих дней))
    WorkDayRightCalendar = Working_days_between_dates(dateForLoad + 1, Date_last_day_month(dateForLoad), 5)
              
          
    ' Переходим на окно DB
    ThisWorkbook.Sheets("Лист8").Activate
          
    ' Обнуление итоговых значений по Офисам и РОО
    ' Call Обнуление_итоговых_значений_по_Офисам_и_РОО
         
    ' Столбец Оffice_Number
    Column_Оffice_Number = ColumnByName("Sales_Office", "Лист1", 1, "Оffice_Number")
    ' Столбец периода
    column_MMYY = ColumnByName("Sales_Office", "Лист1", 1, "MMYY")
    ' Находим номер столбца Date_ДД для месяца
    Column_Date_ДД = ColumnByName("Sales_Office", "Лист1", 1, "Date_" + CStr(Mid(dateForLoad, 1, 2)))
    
    ' Находим номер столбца DateN_ДД для квартала
    M_num = Nom_mes_quarter_str(dateForLoad)
    curr_Day_Month_Q = "Date" + M_num + "_" + Mid(dateForLoad, 1, 2)
    Column_DateN_ДД = ColumnByName("Sales_Office", "Лист1", 1, curr_Day_Month_Q)
    
    ' Находим номер столбца Product_Name
    column_Product_Name = ColumnByName("Sales_Office", "Лист1", 1, "Product_Name")
         
    rowCount = 2
    Do While Not IsEmpty(Workbooks("Sales_Office").Sheets("Лист1").Cells(rowCount, 1).Value)
      
      ' Индикация
      Application.StatusBar = CStr(rowCount) + "..."
    
      ' Если текущая запись - это наш месяц
      If Workbooks("Sales_Office").Sheets("Лист1").Cells(rowCount, column_MMYY).Value = strMMYYVar Then
        
        ' Если это НЕ интегральный рейтинг (Столбец 3)
        If Workbooks("Sales_Office").Sheets("Лист1").Cells(rowCount, 3).Value <> "Интегральный рейтинг" Then
        
          ' Здесь проверяем продукт - Потребительские кредиты считаем календарные дни в прогнозе, т.к. выдачи через ИБ
          If Workbooks("Sales_Office").Sheets("Лист1").Cells(rowCount, column_Product_Name).Value = "Потребительские кредиты" Then
            ' Для Потребительские кредиты считаем календарные минус суббота и воскр.
            WorkDayLeft_Var = WorkDayLeftCalendar
            WorkDayRight_Var = WorkDayRightCalendar
          Else
            ' Для остальных учитываем нерабочие дни из BASE\NonWorkingDays
            WorkDayLeft_Var = WorkDayLeft
            WorkDayRight_Var = WorkDayRight
          End If
        
          ' Вызываем процедуру записи данных на Лист8 (In_Оffice, In_Product_Name, In_Fact)
          Call Записать_данные_на_Лист8(Workbooks("Sales_Office").Sheets("Лист1").Cells(rowCount, Column_Оffice_Number).Value, _
                                          Workbooks("Sales_Office").Sheets("Лист1").Cells(rowCount, column_Product_Name).Value, _
                                            Workbooks("Sales_Office").Sheets("Лист1").Cells(rowCount, Column_Date_ДД).Value, _
                                              WorkDayLeft_Var, _
                                                WorkDayRight_Var)
        
        End If ' Если это НЕ интегральный рейтинг (Столбец 3)
        
        ' Если это интегральный рейтинг (Столбец 3)
        If Workbooks("Sales_Office").Sheets("Лист1").Cells(rowCount, 3).Value = "Интегральный рейтинг" Then
        
          ' Вызываем процедуру записи данных на Лист8 (In_Оffice, In_Product_Name, In_Fact)
          Call Записать_данные_ИР_на_Лист8(Workbooks("Sales_Office").Sheets("Лист1").Cells(rowCount, 4).Value, _
                                             Workbooks("Sales_Office").Sheets("Лист1").Cells(rowCount, column_Product_Name).Value, _
                                               Workbooks("Sales_Office").Sheets("Лист1").Cells(rowCount, Column_Date_ДД).Value, _
                                                 WorkDayLeft, _
                                                   WorkDayRight)
        
        End If ' Если это интегральный рейтинг (Столбец 3)
        
      End If
      
      ' Если текущая запись - это наш Q, то записываем его на Лист8
      If Workbooks("Sales_Office").Sheets("Лист1").Cells(rowCount, column_MMYY).Value = strNQYYVar Then
          
          ' Для квартала
          WorkDayLeft_Var = Working_days_between_dates(Date_begin_day_quarter(dateForLoad), dateForLoad, 5)
          WorkDayRight_Var = Working_days_between_dates(dateForLoad + 1, Date_last_day_quarter(dateForLoad), 5)
          
          ' Вызываем процедуру записи данных на Лист8 (In_Оffice, In_Product_Name, In_Fact)
          Call Записать_данные_на_Лист8_Q(Workbooks("Sales_Office").Sheets("Лист1").Cells(rowCount, Column_Оffice_Number).Value, _
                                          Workbooks("Sales_Office").Sheets("Лист1").Cells(rowCount, column_Product_Name).Value, _
                                            Workbooks("Sales_Office").Sheets("Лист1").Cells(rowCount, Column_DateN_ДД).Value, _
                                              WorkDayLeft_Var, _
                                                WorkDayRight_Var)
      
      End If
    ' Следующая запись
    rowCount = rowCount + 1
    DoEventsInterval (rowCount)

  Loop
         
    Application.StatusBar = ""
  
End Sub

' Вызываем процедуру записи данных на Лист8 для Квартала
Sub Записать_данные_на_Лист8_Q(In_Оffice, In_Product_Name, In_Fact, In_WorkDayLeft, In_WorkDayRight)
  
  ' Для офиса стартовая позиция
  Select Case In_Оffice
    Case 1 ' ОО «Тюменский»
      rowCount_Лист8_Q = getRowFromSheet8("ОО «Тюменский»", In_Product_Name)
    Case 2 ' ОО «Сургутский»
      rowCount_Лист8_Q = getRowFromSheet8("ОО «Сургутский»", In_Product_Name)
    Case 3 ' ОО «Нижневартовский»
      rowCount_Лист8_Q = getRowFromSheet8("ОО «Нижневартовский»", In_Product_Name)
    Case 4 ' ОО «Новоуренгойский»
      rowCount_Лист8_Q = getRowFromSheet8("ОО «Новоуренгойский»", In_Product_Name)
    Case 5 ' ОО «Тарко-Сале»
      rowCount_Лист8_Q = getRowFromSheet8("ОО «Тарко-Сале»", In_Product_Name)
  End Select
  
  
  ' Заносим значение
  ThisWorkbook.Sheets("Лист8").Cells(rowCount_Лист8_Q, 15).Value = In_Fact
  ThisWorkbook.Sheets("Лист8").Cells(rowCount_Лист8_Q, 15).NumberFormat = "#,##0"
  ThisWorkbook.Sheets("Лист8").Cells(rowCount_Лист8_Q, 15).HorizontalAlignment = xlRight

  ' Заносим изменение факта по Кварталу
  If (ThisWorkbook.Sheets("Лист8").Cells(rowCount_Лист8_Q, 6).Value <> "") Then
    ThisWorkbook.Sheets("Лист8").Cells(rowCount_Лист8_Q, 16).Value = Round(ThisWorkbook.Sheets("Лист8").Cells(rowCount_Лист8_Q, 6).Value - In_Fact, 0)
    ThisWorkbook.Sheets("Лист8").Cells(rowCount_Лист8_Q, 16).NumberFormat = "#,##0"
    ThisWorkbook.Sheets("Лист8").Cells(rowCount_Лист8_Q, 16).HorizontalAlignment = xlRight
  End If

End Sub

' Вызываем процедуру записи данных на Лист8 для месяца
Sub Записать_данные_на_Лист8(In_Оffice, In_Product_Name, In_Fact, In_WorkDayLeft, In_WorkDayRight)
  
  ' Для офиса стартовая позиция
  Select Case In_Оffice
    Case 1 ' ОО «Тюменский»
      rowCount = getRowFromSheet8("ОО «Тюменский»", "ОО «Тюменский»") + 3
    Case 2 ' ОО «Сургутский»
      rowCount = getRowFromSheet8("ОО «Сургутский»", "ОО «Сургутский»") + 3
    Case 3 ' ОО «Нижневартовский»
      rowCount = getRowFromSheet8("ОО «Нижневартовский»", "ОО «Нижневартовский»") + 3
    Case 4 ' ОО «Новоуренгойский»
      rowCount = getRowFromSheet8("ОО «Новоуренгойский»", "ОО «Новоуренгойский»") + 3
    Case 5 ' ОО «Тарко-Сале»
      rowCount = getRowFromSheet8("ОО «Тарко-Сале»", "ОО «Тарко-Сале»") + 3
  End Select
  
  ' Число отработанных дней In_WorkDayLeft
  
  ' Число оставшихся дней In_WorkDayRight
    
  ' Переменная
  Данные_внесены = False
  
  ' Определяем размер блока
  Размер_блока = getRowFromSheet8("ОО «Сургутский»", "ОО «Сургутский»") - getRowFromSheet8("ОО «Тюменский»", "ОО «Тюменский»") - 3
  ' Сколько строк прошли в блоке
  Прошли_строк_в_блоке = 0
  
  ' Ищем продукт на Лист8
  Do While (Прошли_строк_в_блоке <= Размер_блока) And (Данные_внесены = False)
  
    ' Проверяем Наименование продукта
    If ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value = In_Product_Name Then
      
      ' Заносим значение
      ThisWorkbook.Sheets("Лист8").Cells(rowCount, 15).Value = In_Fact
      ThisWorkbook.Sheets("Лист8").Cells(rowCount, 15).NumberFormat = "#,##0"
      ThisWorkbook.Sheets("Лист8").Cells(rowCount, 15).HorizontalAlignment = xlRight
      
      ' Заносим изменение факта по месяцу
      If (ThisWorkbook.Sheets("Лист8").Cells(rowCount, 10).Value <> "") And (ThisWorkbook.Sheets("Лист8").Cells(rowCount, 15).Value <> "") Then
        ThisWorkbook.Sheets("Лист8").Cells(rowCount, 16).Value = Round(ThisWorkbook.Sheets("Лист8").Cells(rowCount, 10).Value - In_Fact, 0)
        ThisWorkbook.Sheets("Лист8").Cells(rowCount, 16).NumberFormat = "#,##0"
        ThisWorkbook.Sheets("Лист8").Cells(rowCount, 16).HorizontalAlignment = xlRight
      End If
      
      ' Расчет прогноза на загруженную Дату по Офису (если в 12-ом столбце есть прогноз месяца)
      If (ThisWorkbook.Sheets("Лист8").Cells(rowCount, 12).Value <> "") And (In_Fact <> Empty) Then
        
        ThisWorkbook.Sheets("Лист8").Cells(rowCount, 17).Value = (In_Fact + ((In_Fact / In_WorkDayLeft) * In_WorkDayRight)) / ThisWorkbook.Sheets("Лист8").Cells(rowCount, 9).Value
        ThisWorkbook.Sheets("Лист8").Cells(rowCount, 17).NumberFormat = "0%"
        ThisWorkbook.Sheets("Лист8").Cells(rowCount, 17).HorizontalAlignment = xlRight
        
        ' Динамика прогноза (столбец R (18))
        ThisWorkbook.Sheets("Лист8").Cells(rowCount, 18).Value = ThisWorkbook.Sheets("Лист8").Cells(rowCount, 12).Value - ThisWorkbook.Sheets("Лист8").Cells(rowCount, 17).Value
        ThisWorkbook.Sheets("Лист8").Cells(rowCount, 18).NumberFormat = "0%"
        ThisWorkbook.Sheets("Лист8").Cells(rowCount, 18).HorizontalAlignment = xlRight
        
        ' Окраска ячейки СФЕТОФОР: в красный, если отрицательная динамика и исполнее менее 1
        If (ThisWorkbook.Sheets("Лист8").Cells(rowCount, 18).Value < 0) And (ThisWorkbook.Sheets("Лист8").Cells(rowCount, 11).Value < 1) Then
          Call Full_Color_RangeII("Лист8", rowCount, 18, 0, 100)
        End If

        ' Окраска ячейки СФЕТОФОР: в зеленый, если положительная динамика
        If ThisWorkbook.Sheets("Лист8").Cells(rowCount, 18).Value > 0 Then
          Call Full_Color_RangeII("Лист8", rowCount, 18, 100, 100)
        End If
        
      End If
      
      ' Переменная
      Данные_внесены = True
    
    End If
    
    ' Следующая запись
    rowCount = rowCount + 1
    Прошли_строк_в_блоке = Прошли_строк_в_блоке + 1
  
  Loop
  
  ' Добавляем для РОО
  rowCount = getRowFromSheet8("Итого по РОО «Тюменский»", "Итого по РОО «Тюменский»")
  
  ' Переменная
  Данные_внесены = False
  
  ' Сколько строк прошли в блоке
  Прошли_строк_в_блоке = 0

  
  ' Ищем продукт на Лист8
  Do While (Прошли_строк_в_блоке <= Размер_блока) And (Данные_внесены = False)
  
    ' Проверяем Наименование продукта
    If ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value = In_Product_Name Then
      
      ' Если единица измерения не в %
      If ThisWorkbook.Sheets("Лист8").Cells(rowCount, 4).Value <> "%" Then
      
        ' Заносим значение
        If ThisWorkbook.Sheets("Лист8").Cells(rowCount, 15).Value <> "" Then
          ThisWorkbook.Sheets("Лист8").Cells(rowCount, 15).Value = ThisWorkbook.Sheets("Лист8").Cells(rowCount, 15).Value + CDec(In_Fact)
        Else
          ThisWorkbook.Sheets("Лист8").Cells(rowCount, 15).Value = CDec(In_Fact)
        End If
        ThisWorkbook.Sheets("Лист8").Cells(rowCount, 15).NumberFormat = "#,##0"
        ThisWorkbook.Sheets("Лист8").Cells(rowCount, 15).HorizontalAlignment = xlRight
      
        ' Заносим изменение по месяцу
        If ThisWorkbook.Sheets("Лист8").Cells(rowCount, 10).Value <> "" Then
          Месяц_Факт = ThisWorkbook.Sheets("Лист8").Cells(rowCount, 10).Value
        Else
          Месяц_Факт = 0
        End If
        
        ' ThisWorkbook.Sheets("Лист8").Cells(rowCount, 16).Value = Round(ThisWorkbook.Sheets("Лист8").Cells(rowCount, 10).Value - ThisWorkbook.Sheets("Лист8").Cells(rowCount, 15).Value, 0)
        ThisWorkbook.Sheets("Лист8").Cells(rowCount, 16).Value = Round(Месяц_Факт - ThisWorkbook.Sheets("Лист8").Cells(rowCount, 15).Value, 0)
        ThisWorkbook.Sheets("Лист8").Cells(rowCount, 16).NumberFormat = "#,##0"
        ThisWorkbook.Sheets("Лист8").Cells(rowCount, 16).HorizontalAlignment = xlRight
      
        ' Расчет прогноза на загруженную Дату по РОО
        If (ThisWorkbook.Sheets("Лист8").Cells(rowCount, 12).Value <> "") And (In_Fact <> Empty) Then
      
          ThisWorkbook.Sheets("Лист8").Cells(rowCount, 17).Value = (ThisWorkbook.Sheets("Лист8").Cells(rowCount, 15).Value + ((ThisWorkbook.Sheets("Лист8").Cells(rowCount, 15).Value / In_WorkDayLeft) * In_WorkDayRight)) / ThisWorkbook.Sheets("Лист8").Cells(rowCount, 9).Value
          ThisWorkbook.Sheets("Лист8").Cells(rowCount, 17).NumberFormat = "0%"
          ThisWorkbook.Sheets("Лист8").Cells(rowCount, 17).HorizontalAlignment = xlRight
        
          ' Динамика прогноза (столбец R (18))
          ThisWorkbook.Sheets("Лист8").Cells(rowCount, 18).Value = ThisWorkbook.Sheets("Лист8").Cells(rowCount, 12).Value - ThisWorkbook.Sheets("Лист8").Cells(rowCount, 17).Value
          ThisWorkbook.Sheets("Лист8").Cells(rowCount, 18).NumberFormat = "0%"
          ThisWorkbook.Sheets("Лист8").Cells(rowCount, 18).HorizontalAlignment = xlRight

          ' Окраска ячейки СФЕТОФОР: в красный, если отрицательная динамика
          If (ThisWorkbook.Sheets("Лист8").Cells(rowCount, 18).Value < 0) And (ThisWorkbook.Sheets("Лист8").Cells(rowCount, 11).Value < 1) Then
            Call Full_Color_RangeII("Лист8", rowCount, 18, 0, 100)
          End If

          ' Окраска ячейки СФЕТОФОР: в зеленый, если положительная динамика
          If ThisWorkbook.Sheets("Лист8").Cells(rowCount, 18).Value > 0 Then
            Call Full_Color_RangeII("Лист8", rowCount, 18, 100, 100)
          End If

        End If
      Else
        ' Если единица измерения в %
        
      End If ' Если единица измерения не в %
      
      ' Переменная
      Данные_внесены = True
    End If
    
    ' Следующая запись
    rowCount = rowCount + 1
    Прошли_строк_в_блоке = Прошли_строк_в_блоке + 1
  Loop
  
  
End Sub


' Обнуление итоговых значений по Офисам и РОО
Sub Обнуление_итоговых_значений_по_Офисам_и_РОО()
    
    Call clearСontents2(ThisWorkbook.Name, "Лист8", "O" + CStr(getRowFromSheet8("ОО «Тюменский»", "ОО «Тюменский»") + 3), "R" + CStr(getRowFromSheet8("ОО «Сургутский»", "ОО «Сургутский»") - 1))
    
    Call clearСontents2(ThisWorkbook.Name, "Лист8", "O" + CStr(getRowFromSheet8("ОО «Сургутский»", "ОО «Сургутский»") + 3), "R" + CStr(getRowFromSheet8("ОО «Нижневартовский»", "ОО «Нижневартовский»") - 1))
    
    Call clearСontents2(ThisWorkbook.Name, "Лист8", "O" + CStr(getRowFromSheet8("ОО «Нижневартовский»", "ОО «Нижневартовский»") + 3), "R" + CStr(getRowFromSheet8("ОО «Новоуренгойский»", "ОО «Новоуренгойский»") - 1))
    
    Call clearСontents2(ThisWorkbook.Name, "Лист8", "O" + CStr(getRowFromSheet8("ОО «Новоуренгойский»", "ОО «Новоуренгойский»") + 3), "R" + CStr(getRowFromSheet8("ОО «Тарко-Сале»", "ОО «Тарко-Сале»") - 1))
    
    Call clearСontents2(ThisWorkbook.Name, "Лист8", "O" + CStr(getRowFromSheet8("ОО «Тарко-Сале»", "ОО «Тарко-Сале»") + 3), "R" + CStr(getRowFromSheet8("Интегральный рейтинг по офисам", "Интегральный рейтинг по офисам") - 2))
    
    Call clearСontents2(ThisWorkbook.Name, "Лист8", "O" + CStr(getRowFromSheet8("Итого по РОО «Тюменский»", "Итого по РОО «Тюменский»") + 3), "R" + CStr(getRowFromSheet8("Итого по РОО «Тюменский»", "Итого по РОО «Тюменский»") + (getRowFromSheet8("ОО «Сургутский»", "ОО «Сургутский»") - getRowFromSheet8("ОО «Тюменский»", "ОО «Тюменский»"))))

    ' Интегральный рейтинг по офисам
    Call clearСontents2(ThisWorkbook.Name, "Лист8", "O" + CStr(getRowFromSheet8("Интегральный рейтинг по офисам", "Интегральный рейтинг по офисам") + 3), "R" + CStr(getRowFromSheet8("Итого по РОО «Тюменский»", "Итого по РОО «Тюменский»") - 1))

End Sub

' Получение номера строки на Лист9 по Офису (Тюменский, Сургутский, Нижневартовский, Новоуренгойский, Тарко-Сале) и Наименованию продукта
Function getRowFromSheet8(In_Office, In_ProductName)
  
  ' Итоговое значение
  getRowFromSheet8 = 0
  
  ' Берем с листа ОО «Тюменский»
  rowCount = rowByValue(ThisWorkbook.Name, "Лист8", "ОО «Тюменский»", 100, 100)
  
  '  Переменная секции офиса
  Это_нужный_офис = False
  Значение_определено = False
  
  ' Обрабатываем Лист - ищем Сначала Офис, если находим офис, то ищем позицию с наименованием продукта
  ' Do While (InStr(ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value, "Интегральный рейтинг по офисам") = 0) And (Значение_определено = False)
  Do While (rowCount <= 1000) And (Значение_определено = False)
  
    ' Проверяем офис
    If InStr(ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value, In_Office) <> 0 Then
      Это_нужный_офис = True
    End If
    
    '  Проверяем наименование продукта
    If (Это_нужный_офис = True) And (InStr(ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value, In_ProductName) <> 0) Then
      
      getRowFromSheet8 = rowCount
      
      ' Останавливаем дальнейший поиск
      Значение_определено = True
    End If
    
    
    ' Следующая запись
    rowCount = rowCount + 1
  Loop
    
End Function

' Получение показателя офиса с Лист8
Function getValueFromSheet8(In_Office, In_ProductName, In_Column)
  
  ' Итоговое значение
  getValueFromSheet8 = ThisWorkbook.Sheets("Лист8").Cells( _
                                                          getRowFromSheet8(In_Office, In_ProductName), _
                                                            In_Column).Value
  
End Function

' Как убрать: Будьте внимательны! в документе могут быть персональные данные, которые невозможно удалить с помощью инспектора документов.
' https://www.planetaexcel.ru/forum/index.php?PAGE_NAME=message&FID=1&TID=58709&TITLE_SEO=58709-opoveshchenie-pered-sokhraneniem&MID=993712#message993712
' В параметрах конфиденциальности снять галку "Удалять личные сведения..."


' Переключение Инвестов месяц/квартал
Sub DB_swith_to_MonthQuarter2(In_ReportName_String, In_Sheets, In_Period)

'   ActiveSheet.PivotTables("Сводная таблица5").PivotFields("Период").ClearAllFilters
'   ActiveSheet.PivotTables("Сводная таблица5").PivotFields("Период").CurrentPage = "Квартал"
    
    On Error Resume Next
    ' Сводная таблица5
    ' Workbooks(In_ReportName_String).Sheets(In_Sheets).PivotTables("Сводная таблица5").PivotFields("Период").ClearAllFilters
    ' Workbooks(In_ReportName_String).Sheets(In_Sheets).PivotTables("Сводная таблица5").PivotFields("Период").CurrentPage = In_Period
    
    ' Сводная таблица3
    Workbooks(In_ReportName_String).Sheets(In_Sheets).PivotTables("Сводная таблица3").PivotFields("Период").ClearAllFilters
    Workbooks(In_ReportName_String).Sheets(In_Sheets).PivotTables("Сводная таблица3").PivotFields("Период").CurrentPage = In_Period

'    ActiveSheet.PivotTables("Сводная таблица3").PivotFields("Период").ClearAllFilters
'    ActiveSheet.PivotTables("Сводная таблица3").PivotFields("Период").CurrentPage = "Месяц"
      
      
   ' От 21.04
   If In_Period = "Месяц" Then
     
     ' With ActiveWorkbook.SlicerCaches("Срез_Период5")
     '      .SlicerItems("Месяц").Selected = True
     '      .SlicerItems("Квартал").Selected = False
     '      .SlicerItems("(пусто)").Selected = False
     '  End With
     
     Workbooks(In_ReportName_String).SlicerCaches("Срез_Период5").SlicerItems("Месяц").Selected = True
     Workbooks(In_ReportName_String).SlicerCaches("Срез_Период5").SlicerItems("Квартал").Selected = False
     Workbooks(In_ReportName_String).SlicerCaches("Срез_Период5").SlicerItems("(пусто)").Selected = False

   End If
   
   If In_Period = "Квартал" Then
   
     '  With ActiveWorkbook.SlicerCaches("Срез_Период5")
     '      .SlicerItems("Квартал").Selected = True
     '      .SlicerItems("Месяц").Selected = False
     '      .SlicerItems("(пусто)").Selected = False
     '  End With
     
     Workbooks(In_ReportName_String).SlicerCaches("Срез_Период5").SlicerItems("Квартал").Selected = True
     Workbooks(In_ReportName_String).SlicerCaches("Срез_Период5").SlicerItems("Месяц").Selected = False
     Workbooks(In_ReportName_String).SlicerCaches("Срез_Период5").SlicerItems("(пусто)").Selected = False

   End If
      
      
End Sub

' Предидущая неделя текущего месяца
Function PreviousWeek(In_Date) As Date
  
  ' Минус 7 дней назад
  PreviousWeek = In_Date - 7
  
  ' Проверяем - если мы ушли в прошлый месяц, то оставляем 01 число текущего месяца
  If Month(PreviousWeek) < Month(In_Date) Then
    PreviousWeek = Date_begin_day_month(In_Date)
  End If
  
End Function

' Предидущая неделя текущего Q
Function PreviousWeek2(In_Date) As Date
  
  ' Минус 7 дней назад
  PreviousWeek2 = In_Date - 7
  
  ' Проверяем - если мы ушли в прошлый Q, то оставляем 01 число 1-го месяца в текущем Q
  If quarterName3(PreviousWeek2) <> quarterName3(In_Date) Then
    PreviousWeek2 = Date_begin_day_quarter(In_Date)
  End If
  
End Function


' Вызываем процедуру записи данных на Лист8
Sub Записать_данные_ИР_на_Лист8(In_Оffice, In_Product_Name, In_Fact, In_WorkDayLeft, In_WorkDayRight)
  
  ' Для офиса стартовая позиция
  rowCount = getRowFromSheet8("Интегральный рейтинг по офисам", "Интегральный рейтинг по офисам") + 3
  
  ' Переменная
  Данные_внесены = False
  
  ' Ищем продукт на Лист8
  Do While (Данные_внесены = False) And (Not IsEmpty(ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value)) And (InStr(ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value, "Итого по РОО «Тюменский»") = 0)
    
    ' Если это текущий офис
    If InStr(ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value, In_Оffice) <> 0 Then
    
      ' Заносим значение
      ThisWorkbook.Sheets("Лист8").Cells(rowCount, 15).Value = (In_Fact / 100)
      ThisWorkbook.Sheets("Лист8").Cells(rowCount, 15).NumberFormat = "0%"
      ThisWorkbook.Sheets("Лист8").Cells(rowCount, 15).HorizontalAlignment = xlRight
      
      ' Заносим изменение факта по месяцу
      ThisWorkbook.Sheets("Лист8").Cells(rowCount, 16).Value = ThisWorkbook.Sheets("Лист8").Cells(rowCount, 13).Value - (In_Fact / 100)
      ThisWorkbook.Sheets("Лист8").Cells(rowCount, 16).NumberFormat = "0%"
      ThisWorkbook.Sheets("Лист8").Cells(rowCount, 16).HorizontalAlignment = xlRight
        
      ' Светофор в зависимости от + (зеленый)/- (красный)
      If ThisWorkbook.Sheets("Лист8").Cells(rowCount, 16).Value > 0 Then
        ' Зеленый
        Call Full_Color_RangeII("Лист8", rowCount, 16, 100, 100)
      End If
      If ThisWorkbook.Sheets("Лист8").Cells(rowCount, 16).Value < 0 Then
        ' Красный
        Call Full_Color_RangeII("Лист8", rowCount, 16, 0, 100)
      End If
      
      ' Переменная
      Данные_внесены = True
    
    End If
    
    ' Следующая запись
    rowCount = rowCount + 1
  Loop
  
End Sub

' (3) Делаем расчет План/Факт по Кварталу из BASE\Sales в офисном канале для текущего officeNameInReport
Sub План_Факт_Q_ПК_Sales(In_officeNameInReport, In_Row_Лист8, In_N, In_dateDB)
    
  ' ***
  In_Product_Code = "ПК_Офис"
  In_Product_Name = "в т.ч. Офисный канал"
  In_Unit = "тыс.руб."
  curr_Day_Month = "Date_" + Mid(In_dateDB, 1, 2)
  ' ***
      
  План_офис_ПК = 0
  Факт_Офис_ПК = 0
  Прогноз_Офис_ПК_тыс_руб = 0
    
  ' Месяц квартала 1
  ДатаНачалаКвартала = quarterStartDate(In_dateDB)
  месQ1 = strMMYY(ДатаНачалаКвартала)
  ' Месяц квартала 2
  месQ2 = strMMYY(quarterSecondMonthStartDate(In_dateDB))
  ' Месяц квартала 3
  ДатаКонцаКвартала = Date_last_day_quarter(In_dateDB)
  месQ3 = strMMYY(ДатаКонцаКвартала)
  ' ====
  rowCount = 2
  Do While Not IsEmpty(Workbooks("Sales").Sheets("Лист1").Cells(rowCount, 1).Value)
      
    ' Индикация
    Application.StatusBar = "Расчет ПК (офис) Q " + In_officeNameInReport + ": " + CStr(rowCount) + " ..."
    
    ' Если текущая запись - это наш офис
    If (Workbooks("Sales").Sheets("Лист1").Cells(rowCount, 4).Value = In_officeNameInReport) _
         And ((Workbooks("Sales").Sheets("Лист1").Cells(rowCount, 5).Value = месQ1) Or (Workbooks("Sales").Sheets("Лист1").Cells(rowCount, 5).Value = месQ2) Or (Workbooks("Sales").Sheets("Лист1").Cells(rowCount, 5).Value = месQ3)) _
           And (Workbooks("Sales").Sheets("Лист1").Cells(rowCount, 7).Value = "ПК") Then
      
      План_офис_ПК = План_офис_ПК + Workbooks("Sales").Sheets("Лист1").Cells(rowCount, 8).Value
      Факт_Офис_ПК = Факт_Офис_ПК + Workbooks("Sales").Sheets("Лист1").Cells(rowCount, 10).Value
    
    End If
      
    ' Следующая запись
    rowCount = rowCount + 1
    DoEventsInterval (rowCount)

  Loop

  ' Выводим показатели в In_Row_Лист8
  ' Квартал - план
  ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 5).Value = План_офис_ПК
  ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 5).NumberFormat = "#,##0"
  ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 5).HorizontalAlignment = xlRight
  ' Квартал - факт
  ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 6).Value = Факт_Офис_ПК
  ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 6).NumberFormat = "#,##0"
  ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 6).HorizontalAlignment = xlRight
  ' Месяц - исполнение
  ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 7).Value = РассчетДоли(ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 5).Value, ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 6).Value, 3)
  ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 7).NumberFormat = "0%"
  ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 7).HorizontalAlignment = xlRight
  
  ' Делаем расчет прогноза между ДатаНачалаКвартала и ДатаКонцаКвартала
  Число_прошедших_раб_дней = Working_days_between_datesII(ДатаНачалаКвартала, In_dateDB, 5)
  Число_раб_дней_квартал = Working_days_between_datesII(ДатаНачалаКвартала, ДатаКонцаКвартала, 5)
  
  Прогноз_Офис_ПК_тыс_руб = (Факт_Офис_ПК / Число_прошедших_раб_дней) * Число_раб_дней_квартал
  
  ' Прогноз в %
  ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 8).Value = РассчетДоли(План_офис_ПК, Прогноз_Офис_ПК_тыс_руб, 3)
  ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 8).NumberFormat = "0%"
  ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 8).HorizontalAlignment = xlRight
  
  ' Окраска ячейки СФЕТОФОР ' 70. Заливка ячейки цветом "светофор" - значение In_Value в %, In_Target в %
  Call Full_Color_RangeII("Лист8", In_Row_Лист8, 8, ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 8).Value, 1)

  ' Индикация
  Application.StatusBar = ""

End Sub

' Единичный показатель из вкладки DB (без план/факта)
Sub DB_getParamFromUniversalSheetInDB(In_ReportName_String, In_Sheets, In_Заголовок_столбца_офисы, In_officeNameInReport, In_Row_Лист8, In_N, In_Product_Name, In_Param_Name_In_DB, In_Product_Code, In_Unit, In_Weight, In_Period)
  
  dateDB = CDate(Mid(Workbooks(In_ReportName_String).Sheets("Оглавление").Cells(1, 1).Value, 23, 10))
  
  ' Апдейтим таблицу BASE\Products
  Call Update_BASE_Products(In_Product_Name, In_Product_Code, In_Unit)
  
  ' Вкладка In_Sheets
  Row_Заголовок_столбца_офисы = rowByValue(In_ReportName_String, In_Sheets, In_Заголовок_столбца_офисы, 1000, 1000)
  Column_Заголовок_столбца_офисы = ColumnByValue(In_ReportName_String, In_Sheets, In_Заголовок_столбца_офисы, 1000, 1000)
  
  ' Находим столбец в котором есть In_Param_Name_In_DB
  column_Product_Name = ColumnByValue(In_ReportName_String, In_Sheets, In_Param_Name_In_DB, 1000, 1000)
  
  ' Находим в с столбце "Тюменский ОО1"
  rowCount = Row_Заголовок_столбца_офисы + 1

  Do While (InStr(Workbooks(In_ReportName_String).Sheets(In_Sheets).Cells(rowCount, Column_Заголовок_столбца_офисы).Value, "Общий итог") = 0) And (Not IsEmpty(Workbooks(In_ReportName_String).Sheets(In_Sheets).Cells(rowCount, Column_Заголовок_столбца_офисы).Value))
    
    ' Если это "Тюменский ОО1" - Раскрываем список
    If InStr(Workbooks(In_ReportName_String).Sheets(In_Sheets).Cells(rowCount, Column_Заголовок_столбца_офисы).Value, "Тюменский ОО1") <> 0 Then
      Workbooks(In_ReportName_String).Sheets(In_Sheets).Cells(rowCount, Column_Заголовок_столбца_офисы).ShowDetail = True
      Офис_найден = True
    End If
              
    ' Если это текущий офис
    If (InStr(Workbooks(In_ReportName_String).Sheets(In_Sheets).Cells(rowCount, Column_Заголовок_столбца_офисы).Value, In_officeNameInReport) <> 0) And (InStr(Workbooks(In_ReportName_String).Sheets(In_Sheets).Cells(rowCount, Column_Заголовок_столбца_офисы).Value, "ОО1") = 0) Then
      
      ' Берем из этой строки данные и копируем на Лист8
      Найденное_значение = Workbooks(In_ReportName_String).Sheets(In_Sheets).Cells(rowCount, column_Product_Name).Value

      ' Если измерение в %
      If In_Unit = "%" Then
        ' Если это %, то умножаем на 100
        Найденное_значение = (Найденное_значение * 100)
      End If

    End If
    
    ' Следующая запись
    Application.StatusBar = In_Product_Code + " " + In_officeNameInReport + ": " + CStr(rowCount) + "..."
    rowCount = rowCount + 1
    DoEventsInterval (rowCount)
    
  Loop
 
  ' Заносим показатель на Лист8 в позицию Факт In_Period = "Месяц"/"Квартал"
  If In_Period = "Месяц" Then
    
    Column_In_Лист8 = 10
    
    ' Идентификатор ID_Rec:
    ID_RecVar = CStr(CStr(getNumberOfficeByName(In_officeNameInReport)) + "-" + strMMYY(dateDB) + "-" + In_Product_Code)
    
    ' Текущие значения в месяце: Date_01 (N), Date_02 (O), Date_03 (P), Date_04 (Q), Date_05 Date_06 Date_07 Date_08 Date_09 Date_10 Date_11 Date_12 Date_13 Date_14 Date_15 Date_16 Date_17 Date_18 Date_19 Date_20 Date_21 Date_22 Date_23 Date_24 Date_25 Date_26 Date_27 Date_28 Date_29 Date_30 Date_31
    curr_Day_Month = "Date_" + Mid(dateDB, 1, 2)
    
    ' Период в BASE\Sale_Office
    MMYY_Var = strMMYY(dateDB)
  
  End If
  
  ' Для Квартала
  If In_Period = "Квартал" Then
    
    Column_In_Лист8 = 6
    
    ' Идентификатор ID_Rec:
    ID_RecVar = CStr(CStr(getNumberOfficeByName(In_officeNameInReport)) + "-" + strNQYY(dateDB) + "-" + In_Product_Code)
    
    ' Текущие значения в месяце: Date_01 (N), Date_02 (O), Date_03 (P), Date_04 (Q), Date_05 Date_06 Date_07 Date_08 Date_09 Date_10 Date_11 Date_12 Date_13 Date_14 Date_15 Date_16 Date_17 Date_18 Date_19 Date_20 Date_21 Date_22 Date_23 Date_24 Date_25 Date_26 Date_27 Date_28 Date_29 Date_30 Date_31
    M_num = Nom_mes_quarter_str(dateDB)
    curr_Day_Month = "Date" + M_num + "_" + Mid(dateDB, 1, 2)
    
    ' Период в BASE\Sale_Office
    MMYY_Var = strNQYY(dateDB)
    
  End If
  
  ' Заносим наименование продукта на Лист8
  ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 1).NumberFormat = "@"
  ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 1).Value = In_N
  ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 1).HorizontalAlignment = xlCenter
  ' Наименование
  ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 2).Value = In_Product_Name
  ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 2).HorizontalAlignment = xlLeft
  
  ' Вес выводим, если он не нулевой
  If In_Weight <> 0 Then
    ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 3).Value = In_Weight
    ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 3).NumberFormat = "0.0%"
    ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 3).HorizontalAlignment = xlCenter
  End If
  '
  ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 4).Value = In_Unit
  ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 4).HorizontalAlignment = xlCenter
  
  ' Заносим найденное значение на Лист8
  ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, Column_In_Лист8).Value = Найденное_значение
  ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, Column_In_Лист8).NumberFormat = "#,##0"
  ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, Column_In_Лист8).HorizontalAlignment = xlRight
  
  ' Заносим найденную переменную в BASE\Sale_Office
  Call InsertRecordInBook("Sales_Office", "Лист1", "ID_Rec", ID_RecVar, _
                                            "ID_Rec", ID_RecVar, _
                                              "Оffice_Number", getNumberOfficeByName(In_officeNameInReport), _
                                                "Product_Name", In_Product_Name, _
                                                  "Оffice", In_officeNameInReport, _
                                                    "MMYY", MMYY_Var, _
                                                      "Update_Date", dateDB, _
                                                       "Product_Code", In_Product_Code, _
                                                         "Plan", "", _
                                                            "Unit", In_Unit, _
                                                              "Fact", Найденное_значение, _
                                                                "Percent_Completion", "", _
                                                                  "Prediction", "", _
                                                                    "Percent_Prediction", "", _
                                                                      curr_Day_Month, Найденное_значение, _
                                                                        "", "", _
                                                                          "", "", _
                                                                            "", "", _
                                                                              "", "", _
                                                                                "", "", _
                                                                                  "", "")
  
End Sub

' Единичный показатель из вкладки DB (без план/факта) со возможным сдвигом по столбцам
Sub DB_getParamFromUniversalSheetInDB2(In_ReportName_String, In_Sheets, In_Заголовок_столбца_офисы, In_officeNameInReport, In_Row_Лист8, In_N, In_Product_Name, In_Param_Name_In_DB, In_Product_Code, In_Unit, In_Weight, In_Period, In_DeltaPrediction)
  
  ' In_DeltaPrediction - + число столбцов от столбца In_Param_Name_In_DB в котором находится значение переменной
  
  dateDB = CDate(Mid(Workbooks(In_ReportName_String).Sheets("Оглавление").Cells(1, 1).Value, 23, 10))
  
  ' Апдейтим таблицу BASE\Products
  Call Update_BASE_Products(In_Product_Name, In_Product_Code, In_Unit)
  
  ' Вкладка In_Sheets
  Row_Заголовок_столбца_офисы = rowByValue(In_ReportName_String, In_Sheets, In_Заголовок_столбца_офисы, 1000, 1000)
  Column_Заголовок_столбца_офисы = ColumnByValue(In_ReportName_String, In_Sheets, In_Заголовок_столбца_офисы, 1000, 1000)
  
  ' Находим столбец в котором есть In_Param_Name_In_DB
  column_Product_Name = ColumnByValue(In_ReportName_String, In_Sheets, In_Param_Name_In_DB, 1000, 1000) + In_DeltaPrediction
  
  ' Находим в с столбце "Тюменский ОО1"
  rowCount = Row_Заголовок_столбца_офисы + 1

  Do While (InStr(Workbooks(In_ReportName_String).Sheets(In_Sheets).Cells(rowCount, Column_Заголовок_столбца_офисы).Value, "Общий итог") = 0) And (Not IsEmpty(Workbooks(In_ReportName_String).Sheets(In_Sheets).Cells(rowCount, Column_Заголовок_столбца_офисы).Value))
    
    ' Если это "Тюменский ОО1" - Раскрываем список
    If InStr(Workbooks(In_ReportName_String).Sheets(In_Sheets).Cells(rowCount, Column_Заголовок_столбца_офисы).Value, "Тюменский ОО1") <> 0 Then
      Workbooks(In_ReportName_String).Sheets(In_Sheets).Cells(rowCount, Column_Заголовок_столбца_офисы).ShowDetail = True
      Офис_найден = True
    End If
              
    ' Если это текущий офис
    If (InStr(Workbooks(In_ReportName_String).Sheets(In_Sheets).Cells(rowCount, Column_Заголовок_столбца_офисы).Value, In_officeNameInReport) <> 0) And (InStr(Workbooks(In_ReportName_String).Sheets(In_Sheets).Cells(rowCount, Column_Заголовок_столбца_офисы).Value, "ОО1") = 0) Then
      
      ' Берем из этой строки данные и копируем на Лист8
      Найденное_значение = Workbooks(In_ReportName_String).Sheets(In_Sheets).Cells(rowCount, column_Product_Name).Value

      ' Если измерение в %
      If In_Unit = "%" Then
        ' Если это %, то умножаем на 100
        Найденное_значение = (Найденное_значение * 100)
      End If

    End If
    
    ' Следующая запись
    Application.StatusBar = In_Product_Code + " " + In_officeNameInReport + ": " + CStr(rowCount) + "..."
    rowCount = rowCount + 1
    DoEventsInterval (rowCount)
    
  Loop
 
  ' Заносим показатель на Лист8 в позицию Факт In_Period = "Месяц"/"Квартал"
  If In_Period = "Месяц" Then
    
    Column_In_Лист8 = 10
    
    ' Идентификатор ID_Rec:
    ID_RecVar = CStr(CStr(getNumberOfficeByName(In_officeNameInReport)) + "-" + strMMYY(dateDB) + "-" + In_Product_Code)
    
    ' Текущие значения в месяце: Date_01 (N), Date_02 (O), Date_03 (P), Date_04 (Q), Date_05 Date_06 Date_07 Date_08 Date_09 Date_10 Date_11 Date_12 Date_13 Date_14 Date_15 Date_16 Date_17 Date_18 Date_19 Date_20 Date_21 Date_22 Date_23 Date_24 Date_25 Date_26 Date_27 Date_28 Date_29 Date_30 Date_31
    curr_Day_Month = "Date_" + Mid(dateDB, 1, 2)
    
    ' Период в BASE\Sale_Office
    MMYY_Var = strMMYY(dateDB)
  
  End If
  
  ' Для Квартала
  If In_Period = "Квартал" Then
    
    Column_In_Лист8 = 6
    
    ' Идентификатор ID_Rec:
    ID_RecVar = CStr(CStr(getNumberOfficeByName(In_officeNameInReport)) + "-" + strNQYY(dateDB) + "-" + In_Product_Code)
  
    ' Текущие значения в месяце: Date_01 (N), Date_02 (O), Date_03 (P), Date_04 (Q), Date_05 Date_06 Date_07 Date_08 Date_09 Date_10 Date_11 Date_12 Date_13 Date_14 Date_15 Date_16 Date_17 Date_18 Date_19 Date_20 Date_21 Date_22 Date_23 Date_24 Date_25 Date_26 Date_27 Date_28 Date_29 Date_30 Date_31
    M_num = Nom_mes_quarter_str(dateDB)
    curr_Day_Month = "Date" + M_num + "_" + Mid(dateDB, 1, 2)
    
    ' Период в BASE\Sale_Office
    MMYY_Var = strNQYY(dateDB)
  
  End If
  
  ' Заносим наименование продукта на Лист8
  ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 1).NumberFormat = "@"
  ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 1).Value = In_N
  ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 1).HorizontalAlignment = xlCenter
  ' Наименование
  ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 2).Value = In_Product_Name
  ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 2).HorizontalAlignment = xlLeft
  
  ' Вес выводим, если он не нулевой
  If In_Weight <> 0 Then
    ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 3).Value = In_Weight
    ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 3).NumberFormat = "0.0%"
    ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 3).HorizontalAlignment = xlCenter
  End If
  '
  ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 4).Value = In_Unit
  ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 4).HorizontalAlignment = xlCenter
  
  ' Заносим найденное значение на Лист8
  ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, Column_In_Лист8).Value = Найденное_значение
  ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, Column_In_Лист8).NumberFormat = "#,##0"
  ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, Column_In_Лист8).HorizontalAlignment = xlRight
    
  ' Заносим найденную переменную в BASE\Sale_Office
  Call InsertRecordInBook("Sales_Office", "Лист1", "ID_Rec", ID_RecVar, _
                                            "ID_Rec", ID_RecVar, _
                                              "Оffice_Number", getNumberOfficeByName(In_officeNameInReport), _
                                                "Product_Name", In_Product_Name, _
                                                  "Оffice", In_officeNameInReport, _
                                                    "MMYY", MMYY_Var, _
                                                      "Update_Date", dateDB, _
                                                       "Product_Code", In_Product_Code, _
                                                         "Plan", "", _
                                                            "Unit", In_Unit, _
                                                              "Fact", Найденное_значение, _
                                                                "Percent_Completion", "", _
                                                                  "Prediction", "", _
                                                                    "Percent_Prediction", "", _
                                                                      curr_Day_Month, Найденное_значение, _
                                                                        "", "", _
                                                                          "", "", _
                                                                            "", "", _
                                                                              "", "", _
                                                                                "", "", _
                                                                                  "", "")
  
End Sub


' Выгрузить продукт по офисам с Листа8
Sub Выгрузить_продукт_по_офисам_Лист8()
  
  ' Строка
  не_брать_прогноз_из_20_столбца_по_этим_продуктам_str = "Портфель ЗП 18+, шт._Квартал , Пассивы, в т.ч. Срочные вклады,            Накопительный счет,            СКС,            Прочие ДВС,            Аккредитивы,            Брокер,            УК,           АУМ"
  
  ' Определяем, где находится текущая ячейка. Должен быть диапазон A62:N90 (в относительных от "Повестка_дня" координатах)
  Ячейка_ОО_Тюменский = RangeByValue(ThisWorkbook.Name, "Лист8", "ОО «Тюменский»", 100, 100)
  НомерСтроки_ОО_Тюменский = ThisWorkbook.Sheets("Лист8").Range(Ячейка_ОО_Тюменский).Row
  НомерСтолбца_ОО_Тюменский = ThisWorkbook.Sheets("Лист8").Range(Ячейка_ОО_Тюменский).Column
  
  Ячейка_ОО_Сургутский = RangeByValue(ThisWorkbook.Name, "Лист8", "ОО «Сургутский»", 100, 100)
  НомерСтроки_ОО_Сургутский = ThisWorkbook.Sheets("Лист8").Range(Ячейка_ОО_Сургутский).Row
  НомерСтолбца_ОО_Сургутский = ThisWorkbook.Sheets("Лист8").Range(Ячейка_ОО_Сургутский).Column
  
  ' Проверка где находится активная ячейка
  If (ActiveCell.Row >= НомерСтроки_ОО_Тюменский + 3) And (ActiveCell.Row <= НомерСтроки_ОО_Сургутский - 1) And (ActiveCell.Column >= НомерСтолбца_ОО_Тюменский - 1) And ((ActiveCell.Column <= НомерСтолбца_ОО_Тюменский + 10)) Then
  
    ' Наименование показателя
    Показатель_наименование = ThisWorkbook.Sheets("Лист8").Cells(ActiveCell.Row, НомерСтолбца_ОО_Тюменский).Value
  
    ' Запрос
    If MsgBox("Выгрузить «" + Показатель_наименование + "» по офисам?", vbYesNo) = vbYes Then
    
      ' Формируем выгрузку:
      ' Открываем шаблон
      If Dir(ThisWorkbook.Path + "\Templates\" + "Продажи по продукту.xlsx") <> "" Then
        ' Открываем шаблон Templates\Продажи по продукту
        TemplatesFileName = "Продажи по продукту"
      End If
              
      ' Открываем шаблон из C:\Users\...\Documents\#VBA\DB_Result\Templates
      Workbooks.Open (ThisWorkbook.Path + "\Templates\" + TemplatesFileName + ".xlsx")
           
      ' Переходим на окно DB
      ThisWorkbook.Sheets("Лист8").Activate

      ' Дата DB
      dateDB_Лист8 = CDate(Mid(ThisWorkbook.Sheets("Лист8").Range("B5").Value, 52, 10))

      ' Имя нового файла
      FileDBName = "Продажи " + Показатель_наименование + " " + Replace(CStr(dateDB_Лист8), ".", "-") + ".xlsx"
       
      ' Проверяем - если файл есть, то удаляем его
      Call deleteFile(ThisWorkbook.Path + "\Out\" + FileDBName)

      Workbooks(TemplatesFileName).SaveAs FileName:=ThisWorkbook.Path + "\Out\" + FileDBName, FileFormat:=xlOpenXMLWorkbook, createBackUp:=False
    
      ' Выгрузка в файл данных:
      ' В "B3" и "B4" вставляем Наименование продукта
      Workbooks(FileDBName).Sheets("Лист1").Range("B3").Value = "Продажи по продукту «" + Показатель_наименование + "» на " + CStr(dateDB_Лист8) + " г."
      Workbooks(FileDBName).Sheets("Лист1").Range("B4").Value = "Показатель_наименование"
    
      ' Заголовки таблицы
      Workbooks(FileDBName).Sheets("Лист1").Cells(4, 5).Value = quarterName(dateDB_Лист8)
      Workbooks(FileDBName).Sheets("Лист1").Cells(4, 9).Value = "Месяц (" + ИмяМесяца(dateDB_Лист8) + ")" 'Месяц (сентябрь)
      Workbooks(FileDBName).Sheets("Лист1").Cells(5, 6).Value = "Факт на " + strDDMM(dateDB_Лист8)
      Workbooks(FileDBName).Sheets("Лист1").Cells(5, 10).Value = Workbooks(FileDBName).Sheets("Лист1").Cells(5, 6).Value
      Workbooks(FileDBName).Sheets("Лист1").Cells(4, 15).Value = "Дата " + CStr(ThisWorkbook.Sheets("Лист8").Range("O9").Value)
    
      ' Итоги
      Итого_Квартал_План = 0
      Итого_Квартал_Факт = 0
      Итого_Месяц_План = 0
      Итого_Месяц_Факт = 0
      Число_офисов = 0
    
      ' Выводим продажи продукта по офису
      For i = 1 To 5
        ' Номера офисов от 1 до 5
        Select Case i
          Case 1 ' ОО «Тюменский»
            officeNameInReport = "ОО «Тюменский»"
          Case 2 ' ОО «Сургутский»
            officeNameInReport = "ОО «Сургутский»"
          Case 3 ' ОО «Нижневартовский»
            officeNameInReport = "ОО «Нижневартовский»"
          Case 4 ' ОО «Новоуренгойский»
            officeNameInReport = "ОО «Новоуренгойский»"
          Case 5 ' ОО «Тарко-Сале»
            officeNameInReport = "ОО «Тарко-Сале»"
        End Select
        
        ' Function getRowFromSheet8(In_Office, In_ProductName)
        RowFromSheet8 = getRowFromSheet8(officeNameInReport, Показатель_наименование)
      
        Число_офисов = Число_офисов + 1
      
        ' №
        Workbooks(FileDBName).Sheets("Лист1").Cells(5 + i, 1).Value = CStr(i)
        ' Офис
        Workbooks(FileDBName).Sheets("Лист1").Cells(5 + i, 2).Value = officeNameInReport
        ' Вес
        ThisWorkbook.Sheets("Лист8").Cells(RowFromSheet8, 3).Copy Destination:=Workbooks(FileDBName).Sheets("Лист1").Cells(5 + i, 3)
        ' Ед.изм.
        ThisWorkbook.Sheets("Лист8").Cells(RowFromSheet8, 4).Copy Destination:=Workbooks(FileDBName).Sheets("Лист1").Cells(5 + i, 4)
        Ед_изм_Var = Workbooks(FileDBName).Sheets("Лист1").Cells(5 + i, 4).Value
        
        ' Квартал: План
        ThisWorkbook.Sheets("Лист8").Cells(RowFromSheet8, 5).Copy Destination:=Workbooks(FileDBName).Sheets("Лист1").Cells(5 + i, 5)
        If Workbooks(FileDBName).Sheets("Лист1").Cells(5 + i, 5).Value <> "" Then
          Итого_Квартал_План = Итого_Квартал_План + Workbooks(FileDBName).Sheets("Лист1").Cells(5 + i, 5).Value
        End If
        
        ' Квартал: Факт на ____
        ThisWorkbook.Sheets("Лист8").Cells(RowFromSheet8, 6).Copy Destination:=Workbooks(FileDBName).Sheets("Лист1").Cells(5 + i, 6)
        If Workbooks(FileDBName).Sheets("Лист1").Cells(5 + i, 6).Value <> "" Then
          Итого_Квартал_Факт = Итого_Квартал_Факт + Workbooks(FileDBName).Sheets("Лист1").Cells(5 + i, 6).Value
        End If
        
        ' Квартал: Исп.
        ThisWorkbook.Sheets("Лист8").Cells(RowFromSheet8, 7).Copy Destination:=Workbooks(FileDBName).Sheets("Лист1").Cells(5 + i, 7)
        
        ' Квартал: Прогноз - если он есть в столбце 8, иначе из 20-го
        If ThisWorkbook.Sheets("Лист8").Cells(RowFromSheet8, 8).Value <> "" Then
          ThisWorkbook.Sheets("Лист8").Cells(RowFromSheet8, 8).Copy Destination:=Workbooks(FileDBName).Sheets("Лист1").Cells(5 + i, 8)
        Else
          ' Из 20-го столбца
          If InStr(не_брать_прогноз_из_20_столбца_по_этим_продуктам_str, Показатель_наименование) = 0 Then
            
            ThisWorkbook.Sheets("Лист8").Cells(RowFromSheet8, 20).Copy Destination:=Workbooks(FileDBName).Sheets("Лист1").Cells(5 + i, 8)
            ' Если значение в столбце 8 не пустое
            If Workbooks(FileDBName).Sheets("Лист1").Cells(5 + i, 8).Value <> "" Then
              ' Окраска ячейки СФЕТОФОР ' 70. Заливка ячейки цветом "светофор" - значение In_Value в %, In_Target в %
              Call Full_Color_RangeV(FileDBName, "Лист1", 5 + i, 8, Workbooks(FileDBName).Sheets("Лист1").Cells(5 + i, 8).Value, 1)
              ' Убираем цвет в столбце 7
              Call Убрать_заливку_цветом(FileDBName, "Лист1", 5 + i, 7)
            End If
          
          End If
          
        End If
        
        ' Месяц: План
        ThisWorkbook.Sheets("Лист8").Cells(RowFromSheet8, 9).Copy Destination:=Workbooks(FileDBName).Sheets("Лист1").Cells(5 + i, 9)
        If Workbooks(FileDBName).Sheets("Лист1").Cells(5 + i, 9).Value <> "" Then
          Итого_Месяц_План = Итого_Месяц_План + Workbooks(FileDBName).Sheets("Лист1").Cells(5 + i, 9).Value
        End If
        
        ' Месяц: Факт на ____
        ThisWorkbook.Sheets("Лист8").Cells(RowFromSheet8, 10).Copy Destination:=Workbooks(FileDBName).Sheets("Лист1").Cells(5 + i, 10)
        If Workbooks(FileDBName).Sheets("Лист1").Cells(5 + i, 10).Value <> "" Then
          Итого_Месяц_Факт = Итого_Месяц_Факт + Workbooks(FileDBName).Sheets("Лист1").Cells(5 + i, 10).Value
        End If
        
        ' Месяц: Исп.
        ThisWorkbook.Sheets("Лист8").Cells(RowFromSheet8, 11).Copy Destination:=Workbooks(FileDBName).Sheets("Лист1").Cells(5 + i, 11)
        
        ' Месяц: Прогноз
        ThisWorkbook.Sheets("Лист8").Cells(RowFromSheet8, 12).Copy Destination:=Workbooks(FileDBName).Sheets("Лист1").Cells(5 + i, 12)

        ' Убираем рамки в ячейках строки
        For columnNum = 3 To 12
          ' Форматируем
          Call Убрать_рамку(FileDBName, "Лист1", 5 + i, columnNum)
        Next columnNum
        
        
        ' Изменения: Факт
        ThisWorkbook.Sheets("Лист8").Cells(RowFromSheet8, 15).Copy Destination:=Workbooks(FileDBName).Sheets("Лист1").Cells(5 + i, 15)
        
        ' Изменения: Изм.
        ThisWorkbook.Sheets("Лист8").Cells(RowFromSheet8, 16).Copy Destination:=Workbooks(FileDBName).Sheets("Лист1").Cells(5 + i, 16)
        
        ' Изменения: Прогн.
        ThisWorkbook.Sheets("Лист8").Cells(RowFromSheet8, 17).Copy Destination:=Workbooks(FileDBName).Sheets("Лист1").Cells(5 + i, 17)
        
        ' Изменения: Динамика %
        ThisWorkbook.Sheets("Лист8").Cells(RowFromSheet8, 18).Copy Destination:=Workbooks(FileDBName).Sheets("Лист1").Cells(5 + i, 18)
        
    
        
      Next i

      ' Рисуем черту перед итогом
      Call gorizontalLine(FileDBName, "Лист1", 5 + i)

      ' Если показатель в %, то в итогах рассчет среднего
      If Ед_изм_Var = "%" Then
        Итого_Квартал_План = Итого_Квартал_План / Число_офисов
        Итого_Квартал_Факт = Итого_Квартал_Факт / Число_офисов
        '
        Итого_Месяц_План = Итого_Месяц_План / Число_офисов
        Итого_Месяц_Факт = Итого_Месяц_Факт / Число_офисов
      End If

      ' Итоги Итого по РОО:
      Workbooks(FileDBName).Sheets("Лист1").Cells(5 + i, 2).Value = "Итого по РОО:"
      ' Квартал: План
      Workbooks(FileDBName).Sheets("Лист1").Cells(5 + i, 5).Value = Итого_Квартал_План
      Workbooks(FileDBName).Sheets("Лист1").Cells(5 + i, 5).NumberFormat = "#,##0"
      ' Квартал: Факт на ____
      Workbooks(FileDBName).Sheets("Лист1").Cells(5 + i, 6).Value = Итого_Квартал_Факт
      Workbooks(FileDBName).Sheets("Лист1").Cells(5 + i, 6).NumberFormat = "#,##0"
      ' Квартал: Исп.
      Workbooks(FileDBName).Sheets("Лист1").Cells(5 + i, 7).Value = РассчетДоли(Итого_Квартал_План, Итого_Квартал_Факт, 3)
      Workbooks(FileDBName).Sheets("Лист1").Cells(5 + i, 7).NumberFormat = "0%"
      ' Квартал: Прогноз
      If Ед_изм_Var <> "%" Then
        
        ' Если из 20-го столбца не берем расчетный прогноз
        If InStr(не_брать_прогноз_из_20_столбца_по_этим_продуктам_str, Показатель_наименование) = 0 Then
          Workbooks(FileDBName).Sheets("Лист1").Cells(5 + i, 8).Value = Прогноз_квартала_проц(dateDB_Лист8, Итого_Квартал_План, Итого_Квартал_Факт, 5, 0)
          Workbooks(FileDBName).Sheets("Лист1").Cells(5 + i, 8).NumberFormat = "0%"
        End If
        
      End If
      
      '
      ' Месяц: План
      Workbooks(FileDBName).Sheets("Лист1").Cells(5 + i, 9).Value = Итого_Месяц_План
      Workbooks(FileDBName).Sheets("Лист1").Cells(5 + i, 9).NumberFormat = "#,##0"

      ' Месяц: Факт на ____
      Workbooks(FileDBName).Sheets("Лист1").Cells(5 + i, 10).Value = Итого_Месяц_Факт
      Workbooks(FileDBName).Sheets("Лист1").Cells(5 + i, 10).NumberFormat = "#,##0"
      
      ' Месяц: Исп.
      Workbooks(FileDBName).Sheets("Лист1").Cells(5 + i, 11).Value = РассчетДоли(Итого_Месяц_План, Итого_Месяц_Факт, 3)
      Workbooks(FileDBName).Sheets("Лист1").Cells(5 + i, 11).NumberFormat = "0%"
            
      ' Месяц: Прогноз
      Workbooks(FileDBName).Sheets("Лист1").Cells(5 + i, 12).Value = ""
      Workbooks(FileDBName).Sheets("Лист1").Cells(5 + i, 12).NumberFormat = "0%"
      
    
      ' Закрытие файла
      Workbooks(FileDBName).Close SaveChanges:=True
    
      ' Сообщение о том, что выгрузка завершена
      MsgBox ("Данные по «" + Показатель_наименование + "» выгружены в файл!")
  
      ' Запрос - открыть файл?
      If MsgBox("Открыть сформированный файл " + Dir(FileDBName) + "?", vbYesNo) = vbYes Then
        Workbooks.Open (ThisWorkbook.Path + "\Out\" + FileDBName)
      End If
  
      ' Запрос - Отправить файл в почту?
      If MsgBox("Отправить сформированный файл " + Dir(FileDBName) + " в почту?", vbYesNo) = vbYes Then
        Call Отправка_Lotus_Notes_Показатель_Лист8(dateDB_Лист8, Показатель_наименование, ThisWorkbook.Path + "\Out\" + FileDBName)
      End If
  
    End If
  
  Else
  
    MsgBox ("Перейдите в блок офиса!")
  
  End If ' Проверка нахождения в блоке офиса
  
End Sub

' Чертим горизонтальную линию
Sub gorizontalLine(In_FileDBName, In_Sheets, In_Row)

  Workbooks(In_FileDBName).Sheets(In_Sheets).Range("B" + CStr(In_Row) + ":L" + CStr(In_Row)).Borders(xlDiagonalDown).LineStyle = xlNone
    
  Workbooks(In_FileDBName).Sheets(In_Sheets).Range("B" + CStr(In_Row) + ":L" + CStr(In_Row)).Borders(xlDiagonalUp).LineStyle = xlNone
  Workbooks(In_FileDBName).Sheets(In_Sheets).Range("B" + CStr(In_Row) + ":L" + CStr(In_Row)).Borders(xlEdgeLeft).LineStyle = xlNone
  With Workbooks(In_FileDBName).Sheets(In_Sheets).Range("B" + CStr(In_Row) + ":L" + CStr(In_Row)).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
  End With
  Workbooks(In_FileDBName).Sheets(In_Sheets).Range("B" + CStr(In_Row) + ":L" + CStr(In_Row)).Borders(xlEdgeBottom).LineStyle = xlNone
  Workbooks(In_FileDBName).Sheets(In_Sheets).Range("B" + CStr(In_Row) + ":L" + CStr(In_Row)).Borders(xlEdgeRight).LineStyle = xlNone
  Workbooks(In_FileDBName).Sheets(In_Sheets).Range("B" + CStr(In_Row) + ":L" + CStr(In_Row)).Borders(xlInsideVertical).LineStyle = xlNone
  Workbooks(In_FileDBName).Sheets(In_Sheets).Range("B" + CStr(In_Row) + ":L" + CStr(In_Row)).Borders(xlInsideHorizontal).LineStyle = xlNone

End Sub



' Чертим горизонтальную линию 2
Sub gorizontalLineII(In_FileDBName, In_Sheets, In_Row, In_ColumnBegin, In_ColumnEnd)

  letterColumnBegin = ConvertToLetter(In_ColumnBegin)
  letterColumnEnd = ConvertToLetter(In_ColumnEnd)

  Workbooks(In_FileDBName).Sheets(In_Sheets).Range(letterColumnBegin + CStr(In_Row) + ":" + letterColumnEnd + CStr(In_Row)).Borders(xlDiagonalDown).LineStyle = xlNone
    
  Workbooks(In_FileDBName).Sheets(In_Sheets).Range(letterColumnBegin + CStr(In_Row) + ":" + letterColumnEnd + CStr(In_Row)).Borders(xlDiagonalUp).LineStyle = xlNone
  Workbooks(In_FileDBName).Sheets(In_Sheets).Range(letterColumnBegin + CStr(In_Row) + ":" + letterColumnEnd + CStr(In_Row)).Borders(xlEdgeLeft).LineStyle = xlNone
  With Workbooks(In_FileDBName).Sheets(In_Sheets).Range(letterColumnBegin + CStr(In_Row) + ":" + letterColumnEnd + CStr(In_Row)).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
  End With
  Workbooks(In_FileDBName).Sheets(In_Sheets).Range(letterColumnBegin + CStr(In_Row) + ":" + letterColumnEnd + CStr(In_Row)).Borders(xlEdgeBottom).LineStyle = xlNone
  Workbooks(In_FileDBName).Sheets(In_Sheets).Range(letterColumnBegin + CStr(In_Row) + ":" + letterColumnEnd + CStr(In_Row)).Borders(xlEdgeRight).LineStyle = xlNone
  Workbooks(In_FileDBName).Sheets(In_Sheets).Range(letterColumnBegin + CStr(In_Row) + ":" + letterColumnEnd + CStr(In_Row)).Borders(xlInsideVertical).LineStyle = xlNone
  Workbooks(In_FileDBName).Sheets(In_Sheets).Range(letterColumnBegin + CStr(In_Row) + ":" + letterColumnEnd + CStr(In_Row)).Borders(xlInsideHorizontal).LineStyle = xlNone

End Sub


' Отправка письма с показателем
Sub Отправка_Lotus_Notes_Показатель_Лист8(In_dateDB_Лист8, In_Показатель_наименование, In_fileName)
Dim темаПисьма, текстПисьма, hashTag, attachmentFile As String
Dim i As Byte
  
    
    ' Тема письма - Тема:
    темаПисьма = In_Показатель_наименование + " на " + CStr(In_dateDB_Лист8)

    ' hashTag - Хэштэг:
    hashTag = hashTagFromSheet("Лист8")

    ' Файл-вложение (!!!)
    attachmentFile = In_fileName
    
    ' Текст письма
    текстПисьма = "" + Chr(13)
    текстПисьма = текстПисьма + "" + getFromAddrBook("УДО2,УДО3,УДО4,УДО5,НОРПиКО1,НОРПиКО2,НОРПиКО3,НОРПиКО4,НОРПиКО5,ПМ,МРК1,МРК2,МРК3,МРК4,МРК5", 2) + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "" + getFromAddrBook("РД", 2) + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "Уважаемые руководители," + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + In_Показатель_наименование + " на " + CStr(In_dateDB_Лист8) + " г." + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    ' Визитка (подпись С Ув., )
    текстПисьма = текстПисьма + ПодписьВПисьме()
    ' Хэштег
    текстПисьма = текстПисьма + createBlankStr(27) + hashTag
    
    ' Вызов
    Call send_Lotus_Notes(темаПисьма, "Sergey Fedorovich Proschaev/Tyumen/PSBank/Ru", "Sergey Fedorovich Proschaev/Tyumen/PSBank/Ru", текстПисьма, attachmentFile)
    
    ' Сообщение
    MsgBox ("Письмо отправлено!")
  
End Sub


' Выполнение ИЗП ГО по месяцу
Sub Выполнение_ИПЗ_ГО()
      
      ' Строка статуса
      Application.StatusBar = "Выполнение ИЗП ГО: Формирование отчета..."
      
      ' Формируем выгрузку:
      ' Открываем шаблон
      If Dir(ThisWorkbook.Path + "\Templates\" + "Выполнение ИПЗ ГО.xlsx") <> "" Then
        ' Открываем шаблон Templates\Продажи по продукту
        TemplatesFileName = "Выполнение ИПЗ ГО"
      End If
              
      ' Открываем шаблон из C:\Users\...\Documents\#VBA\DB_Result\Templates
      Workbooks.Open (ThisWorkbook.Path + "\Templates\" + TemplatesFileName + ".xlsx")
           
      ' Переходим на окно DB
      ThisWorkbook.Sheets("Лист8").Activate

      ' Дата DB
      dateDB_Лист8 = CDate(Mid(ThisWorkbook.Sheets("Лист8").Range("B5").Value, 52, 10))

      ' Имя нового файла
      FileDBName = "Выполнение ИПЗ " + Replace(CStr(dateDB_Лист8), ".", "-") + ".xlsx"
      
      ' Проверяем - если файл есть, то удаляем его
      Call deleteFile(ThisWorkbook.Path + "\Out\" + FileDBName)
      
      Workbooks(TemplatesFileName).SaveAs FileName:=ThisWorkbook.Path + "\Out\" + FileDBName, FileFormat:=xlOpenXMLWorkbook, createBackUp:=False
    
      ' В "B2" заголовок
      Workbooks(FileDBName).Sheets("Лист1").Range("A2").Value = "Выполнение ИПЗ офисами за " + ИмяМесяца(dateDB_Лист8) + " на " + CStr(dateDB_Лист8) + " г."
    
      ' Прогноз по месяцу или кварталу
      If ThisWorkbook.Sheets("Лист8").Range("I3").Value = 1 Then
        ' 1 = Месяц
        cloumn_Лист8_План = 9
        cloumn_Лист8_Факт = 10
        cloumn_Лист8_Исп% = 11
        cloumn_Лист8_Прогноз = 12
      Else
        ' 2 = Квартал
        cloumn_Лист8_План = 5
        cloumn_Лист8_Факт = 6
        cloumn_Лист8_Исп% = 7
        cloumn_Лист8_Прогноз = 8
      End If
      
    
      ' Выводим продажи продукта по офису
      For i = 1 To 5
        ' Номера офисов от 1 до 5
        Select Case i
          Case 1 ' ОО «Тюменский»
            officeNameInReport = "ОО «Тюменский»"
            officeNameInReportIR = "ОО " + Chr(34) + "Тюменский" + Chr(34)
            officeNameЛист7 = "Тюменский"
          Case 2 ' ОО «Сургутский»
            officeNameInReport = "ОО «Сургутский»"
            officeNameInReportIR = "ОО2" + Chr(34) + "Сургутский" + Chr(34)
            officeNameЛист7 = "Сургутский"
          Case 3 ' ОО «Нижневартовский»
            officeNameInReport = "ОО «Нижневартовский»"
            officeNameInReportIR = "ОО2 " + Chr(34) + "Нижневартовский" + Chr(34)
            officeNameЛист7 = "Нижневартовский"
          Case 4 ' ОО «Новоуренгойский»
            officeNameInReport = "ОО «Новоуренгойский»"
            officeNameInReportIR = "ОО2" + Chr(34) + "Новоуренгойский" + Chr(34)
            officeNameЛист7 = "Новоуренгойский"
          Case 5 ' ОО «Тарко-Сале»
            officeNameInReport = "ОО «Тарко-Сале»"
            officeNameInReportIR = "ОО2 " + Chr(34) + "Тарко-Сале" + Chr(34)
            officeNameЛист7 = "Тарко-Сале"
        End Select
        
        ' Function getRowFromSheet8(In_Office, In_ProductName)
        RowFromSheet8 = getRowFromSheet8(officeNameInReport, Показатель_наименование)
      
        ' №
        Workbooks(FileDBName).Sheets("Лист1").Cells(5 + i, 1).Value = CStr(i)
        Workbooks(FileDBName).Sheets("Лист1").Cells(5 + i, 1).HorizontalAlignment = xlCenter
        ' Офис
        Workbooks(FileDBName).Sheets("Лист1").Cells(5 + i, 2).Value = officeNameInReport
        
        ' Потреб. кредитование: Прогноз, %
        Workbooks(FileDBName).Sheets("Лист1").Cells(5 + i, 3).Value = ThisWorkbook.Sheets("Лист8").Cells(getRowFromSheet8(officeNameInReport, "Потребительские кредиты"), cloumn_Лист8_Прогноз).Value
        Workbooks(FileDBName).Sheets("Лист1").Cells(5 + i, 3).NumberFormat = "0%"
        ' Окраска ячейки СФЕТОФОР ' 70. Заливка ячейки цветом "светофор" - значение In_Value в %, In_Target в %
        Call Full_Color_RangeV(FileDBName, "Лист1", 5 + i, 3, Workbooks(FileDBName).Sheets("Лист1").Cells(5 + i, 3).Value, 1)

        ' Потреб. кредитование: Конверсия PA, %
        Workbooks(FileDBName).Sheets("Лист1").Cells(5 + i, 4).Value = ThisWorkbook.Sheets("Лист6").Cells(rowByValue(ThisWorkbook.Name, "Лист6", officeNameInReport, 100, 100), 17).Value
        Workbooks(FileDBName).Sheets("Лист1").Cells(5 + i, 4).NumberFormat = "0%"

        ' Проникновение в ЗП, %
        Workbooks(FileDBName).Sheets("Лист1").Cells(5 + i, 5).Value = ThisWorkbook.Sheets("КроссЗП").Cells(rowByValue(ThisWorkbook.Name, "КроссЗП", officeNameInReport, 100, 100), 7).Value
        Workbooks(FileDBName).Sheets("Лист1").Cells(5 + i, 5).NumberFormat = "0.0%"

        ' Комиссионный доход: Прогноз, %
        Workbooks(FileDBName).Sheets("Лист1").Cells(5 + i, 6).Value = ThisWorkbook.Sheets("Лист8").Cells(getRowFromSheet8(officeNameInReport, "Комиссионный доход"), cloumn_Лист8_Прогноз).Value
        Workbooks(FileDBName).Sheets("Лист1").Cells(5 + i, 6).NumberFormat = "0%"
        ' Окраска ячейки СФЕТОФОР ' 70. Заливка ячейки цветом "светофор" - значение In_Value в %, In_Target в %
        Call Full_Color_RangeV(FileDBName, "Лист1", 5 + i, 6, Workbooks(FileDBName).Sheets("Лист1").Cells(5 + i, 6).Value, 1)

        ' Комиссионный доход: ИСЖ+НСЖ Прогноз, % !!! добавить НСЖ !!!
        ' Это просто прогноз ИСЖ Workbooks(FileDBName).Sheets("Лист1").Cells(5 + i, 7).Value = ThisWorkbook.Sheets("Лист8").Cells(getRowFromSheet8(officeNameInReport, "Премия ИСЖ МАСС"), 12).Value
        ' Берем план ИСЖ + план НСЖ
        План_ИСЖ_НСЖ = ThisWorkbook.Sheets("Лист8").Cells(getRowFromSheet8(officeNameInReport, "Премия ИСЖ МАСС"), cloumn_Лист8_План).Value + ThisWorkbook.Sheets("Лист8").Cells(getRowFromSheet8(officeNameInReport, "Премия НСЖ МАСС"), cloumn_Лист8_План).Value
        Факт_ИСЖ_НСЖ = ThisWorkbook.Sheets("Лист8").Cells(getRowFromSheet8(officeNameInReport, "Премия ИСЖ МАСС"), cloumn_Лист8_Факт).Value + ThisWorkbook.Sheets("Лист8").Cells(getRowFromSheet8(officeNameInReport, "Премия НСЖ МАСС"), cloumn_Лист8_Факт).Value
        
        ' Делаем расчет прогноза
        Число_прошедших_раб_дней_месяц = Working_days_between_dates(Date_begin_day_month(dateDB_Лист8), dateDB_Лист8, 5)
        Число_раб_дней_месяц = Working_days_between_dates(Date_begin_day_month(dateDB_Лист8), Date_last_day_month(dateDB_Лист8), 5)
        Прогноз_ИСЖ_НСЖ_тыс_руб = (Факт_ИСЖ_НСЖ / Число_прошедших_раб_дней_месяц) * Число_раб_дней_месяц
        Прогноз_ИСЖ_НСЖ_проц = РассчетДоли(План_ИСЖ_НСЖ, Прогноз_ИСЖ_НСЖ_тыс_руб, 3)
        Workbooks(FileDBName).Sheets("Лист1").Cells(5 + i, 7).Value = Прогноз_ИСЖ_НСЖ_проц
        Workbooks(FileDBName).Sheets("Лист1").Cells(5 + i, 7).NumberFormat = "0%"

        ' ПК с БС и КСП, % (N=60%)
        ПК_факт_шт = getDataFromSheet7(officeNameЛист7, "Потреб кредитование (кредиты сотрудникам Банка исключены)")
        КСП_факт_шт = getDataFromSheet7(officeNameЛист7, "Коробочное страхование")
        Workbooks(FileDBName).Sheets("Лист1").Cells(5 + i, 8).Value = РассчетДоли(ПК_факт_шт, КСП_факт_шт, 3)
        Workbooks(FileDBName).Sheets("Лист1").Cells(5 + i, 8).NumberFormat = "0%"

        ' Кредитные карты: Прогноз, %
        Workbooks(FileDBName).Sheets("Лист1").Cells(5 + i, 9).Value = ThisWorkbook.Sheets("Лист8").Cells(getRowFromSheet8(officeNameInReport, "Кредитные карты (актив.)"), cloumn_Лист8_Прогноз).Value
        Workbooks(FileDBName).Sheets("Лист1").Cells(5 + i, 9).NumberFormat = "0%"
        ' Окраска ячейки СФЕТОФОР ' 70. Заливка ячейки цветом "светофор" - значение In_Value в %, In_Target в %
        Call Full_Color_RangeV(FileDBName, "Лист1", 5 + i, 9, Workbooks(FileDBName).Sheets("Лист1").Cells(5 + i, 9).Value, 1)

        ' Кредитные карты: Заявки на потоке, %
        Workbooks(FileDBName).Sheets("Лист1").Cells(5 + i, 10).Value = ThisWorkbook.Sheets("Лист6").Cells(rowByValue(ThisWorkbook.Name, "Лист6", officeNameInReport, 100, 100), 5).Value
        Workbooks(FileDBName).Sheets("Лист1").Cells(5 + i, 10).NumberFormat = "0%"

        ' Сплиты к ПК, %
        Workbooks(FileDBName).Sheets("Лист1").Cells(5 + i, 11).Value = ThisWorkbook.Sheets("Лист8").Cells(getRowFromSheet8(officeNameInReport, "Сплиты к ПК"), cloumn_Лист8_Исп%).Value
        Workbooks(FileDBName).Sheets("Лист1").Cells(5 + i, 11).NumberFormat = "0%"

        ' Инвесты: Прогноз, %
        Workbooks(FileDBName).Sheets("Лист1").Cells(5 + i, 12).Value = ThisWorkbook.Sheets("Лист8").Cells(getRowFromSheet8(officeNameInReport, "Инвест"), cloumn_Лист8_Исп%).Value
        Workbooks(FileDBName).Sheets("Лист1").Cells(5 + i, 12).NumberFormat = "0%"
        ' Окраска ячейки СФЕТОФОР ' 70. Заливка ячейки цветом "светофор" - значение In_Value в %, In_Target в %
        ' Call Full_Color_RangeV(FileDBName, "Лист1", 5 + i, 12, Workbooks(FileDBName).Sheets("Лист1").Cells(5 + i, 12).Value, 1)

        ' Активные ЗП карты 18+
        Workbooks(FileDBName).Sheets("Лист1").Cells(5 + i, 13).Value = ThisWorkbook.Sheets("Лист8").Cells(getRowFromSheet8(officeNameInReport, "Зарплатные карты 18+"), 7).Value
        Workbooks(FileDBName).Sheets("Лист1").Cells(5 + i, 13).NumberFormat = "0%"
  
        ' Интегр рейтинг
        ' Workbooks(FileDBName).Sheets("Лист1").Cells(5 + i, 14).Value = ThisWorkbook.Sheets("Лист8").Cells(getRowFromSheet8("Интегральный рейтинг по офисам", officeNameInReportIR), 13).Value
        ' Workbooks(FileDBName).Sheets("Лист1").Cells(5 + i, 14).NumberFormat = "0%"
        
      Next i

      ' Рисуем черту перед итогом
      Call gorizontalLineII(FileDBName, "Лист1", 5 + i, 2, 14)

      ' Итоги Итого по РОО:
      Workbooks(FileDBName).Sheets("Лист1").Cells(5 + i, 2).Value = "Итого по РОО:"
      ' Квартал: План
      ' Workbooks(FileDBName).Sheets("Лист1").Cells(5 + i, 5).Value = Итого_Квартал_План
      ' Workbooks(FileDBName).Sheets("Лист1").Cells(5 + i, 5).NumberFormat = "#,##0"
    
      ' Среднее значение Конверсия PA, %
      Workbooks(FileDBName).Sheets("Лист1").Cells(5 + i, 4).FormulaR1C1 = "=AVERAGE(R[-5]C:R[-1]C)"
      
      ' Среднее значение Проникновение в ЗП, %
      Workbooks(FileDBName).Sheets("Лист1").Cells(5 + i, 5).FormulaR1C1 = "=AVERAGE(R[-5]C:R[-1]C)"
      
      ' Среднее значение ПК с БС и КСП, % (N=60%)
      Workbooks(FileDBName).Sheets("Лист1").Cells(5 + i, 8).FormulaR1C1 = "=AVERAGE(R[-5]C:R[-1]C)"
      
      ' Среднее значение Заявки на потоке, %
      Workbooks(FileDBName).Sheets("Лист1").Cells(5 + i, 10).FormulaR1C1 = "=AVERAGE(R[-5]C:R[-1]C)"
      
      ' Среднее значение Интегрального рейтинга
      Workbooks(FileDBName).Sheets("Лист1").Cells(5 + i, 14).FormulaR1C1 = "=AVERAGE(R[-5]C:R[-1]C)"
    
      ' Закрытие файла
      Workbooks(FileDBName).Close SaveChanges:=True
    
      ' Строка статуса
      Application.StatusBar = "Выполнение ИПЗ ГО: Отправка отчета..."
  
      ' Отправка файла в почту
      Call Отправка_Lotus_Notes_Выполнение_ИПЗ_ГО_Лист8(ThisWorkbook.Path + "\Out\" + FileDBName, dateDB_Лист8)
  
      ' Строка статуса
      Application.StatusBar = "Выполнение ИПЗ ГО: Отчет отправлен"
      Application.StatusBar = ""
  
  
End Sub

' Отправка письма: отправляю шаблон самому себе для последующей отправки в сеть письма на его основе:
Sub Отправка_Lotus_Notes_Выполнение_ИПЗ_ГО_Лист8(In_fileName, In_dateDB)
Dim темаПисьма, текстПисьма, hashTag, attachmentFile As String
Dim i As Byte
  
  
  ' Запрос
  ' If MsgBox("Отправить себе Шаблон письма с фокусами контроля '" + ПериодКонтроля + "'?", vbYesNo) = vbYes Then
    
    ' Тема письма - Тема:
    ' темаПисьма = ThisWorkbook.Sheets("Лист8").Cells(RowByValue(ThisWorkbook.Name, "Лист8", "Тема:", 100, 100), ColumnByValue(ThisWorkbook.Name, "Лист8", "Тема:", 100, 100) + 1).Value
    темаПисьма = "Выполнение ИПЗ 1 кв. 2021 на " + CStr(In_dateDB) + " г."

    ' hashTag - Хэштэг:
    ' hashTag = ThisWorkbook.Sheets("Лист8").Cells(RowByValue(ThisWorkbook.Name, "Лист8", "Хэштэг:", 100, 100), ColumnByValue(ThisWorkbook.Name, "Лист8", "Хэштэг:", 100, 100) + 1).Value
    ' hashTag - Хэштэг:
    hashTag = "#Выполнение_ИПЗ_ГО"

    ' Файл-вложение (!!!)
    attachmentFile = In_fileName
    
    ' Текст письма
    текстПисьма = "" + Chr(13)
    текстПисьма = текстПисьма + "" + getFromAddrBook("УДО2,УДО3,УДО4,УДО5,НОРПиКО1,НОРПиКО2,НОРПиКО3,НОРПиКО4,НОРПиКО5,НОКП,РРКК", 2) + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "" + getFromAddrBook("РД", 2) + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "Уважаемые руководители," + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "В рамках исполнения поручения п.3 протокола Собрания №6-01022021 направляю факт исполнения ИПЗ на " + strDDMM(In_dateDB) + "*" + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    ' текстПисьма = текстПисьма + "* - Показатель: ПК с БС и КСП будет сформирован дополнительно" + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    ' Визитка (подпись С Ув., )
    текстПисьма = текстПисьма + ПодписьВПисьме()
    ' Хэштег
    текстПисьма = текстПисьма + createBlankStr(27) + hashTag
    ' Вызов
    Call send_Lotus_Notes(темаПисьма, "Sergey Fedorovich Proschaev/Tyumen/PSBank/Ru", "Sergey Fedorovich Proschaev/Tyumen/PSBank/Ru", текстПисьма, attachmentFile)
  
    ' Зачеркнуть
    ' Call ЗачеркиваемТекстВячейке("Лист0", RangeByValue(ThisWorkbook.Name, "Лист0", "DashBoard (при наличии)", 100, 100))
  
    ' Сообщение
    ' MsgBox ("Письмо отправлено!")
     
  ' End If
  
End Sub

' Выгрузить файл с дневным планом продаж по форме обновленной форме (на основе формы данилова)
' Templates\Ежедневная форма отчёта (куратор) 2.xlsx
Sub Выгрузить_план_дневных_продаж2()
Dim FileNewVar As String

  ' Формируем цели на день по форме Данилова
  
  ' Запрос на формирование
  If MsgBox("Сформировать поручения на день для офисов?", vbYesNo) = vbYes Then
      

    ' Дата формирования - если сегодня понедельник, то формируем за пятницу
    ' Если текущая дата это понедельник, то формируем отчет за пятницу
    If Weekday(CurrDate, vbMonday) = 1 Then
      dateReport = Date - 3
    Else
      dateReport = Date
    End If

    ' Остаток рабочих дней определяем число рабочих дней с понеделника до конца месяца Working_days_between_dateReports(In_dateReportStart, In_dateReportEnd, In_working_days_in_the_week) As Integer
    Остаток_рабочих_дней = Working_days_between_dates(dateReport - 1, Date_last_day_month(dateReport), 5)

    ' Наименование листа в файле (TemplateSheets)
    TS = "Ежедневный отчет"

    ' Строка с именами файлов для архивирования
    strFileNewVar_Office = ""

    ' Проходим по Листу8 и заполняем планы:
    For i = 1 To 5
        ' Номера офисов от 1 до 5
        Select Case i
          Case 1 ' ОО «Тюменский»
            officeNameInReport = "ОО «Тюменский»"
          Case 2 ' ОО «Сургутский»
            officeNameInReport = "ОО «Сургутский»"
          Case 3 ' ОО «Нижневартовский»
            officeNameInReport = "ОО «Нижневартовский»"
          Case 4 ' ОО «Новоуренгойский»
            officeNameInReport = "ОО «Новоуренгойский»"
          Case 5 ' ОО «Тарко-Сале»
            officeNameInReport = "ОО «Тарко-Сале»"
        End Select
        
        ' Сообщение
        Application.StatusBar = "Формирование по " + officeNameInReport + "..."
                
        ' Открываем шаблон C:\Users\proschaevsf\Documents\#DB_Result\Templates\Ежедневная форма отчёта (куратор) 2.xlsx
        fileTemplatesName = "Ежедневная форма отчёта (куратор) 2.xlsx"
        Workbooks.Open (ThisWorkbook.Path + "\Templates\" + fileTemplatesName)
        
        ' Имя нового файла
        FileNewVar = "Ежедневная_форма_отчёта_" + cityOfficeNameByNumber(i) + "_" + strДД_MM_YYYY(dateReport) + ".xlsx"
         
        ' Проверяем - если файл есть, то удаляем его
        Call deleteFile(ThisWorkbook.Path + "\Out\" + FileNewVar)
       
        Workbooks(fileTemplatesName).SaveAs FileName:=ThisWorkbook.Path + "\Out\" + FileNewVar, FileFormat:=xlOpenXMLWorkbook, createBackUp:=False
        
        ' Строка с именами файлов для архивирования
        strFileNewVar_Office = strFileNewVar_Office + ThisWorkbook.Path + "\Out\" + FileNewVar + " "
        ' Переходим на окно DB
        ThisWorkbook.Sheets("Лист8").Activate
    
        ' Заголовок в исходящем файле
        Workbooks(FileNewVar).Sheets(TS).Range("A1").Value = "Продажи за: " + CStr(dateReport) + " (ост.дней " + CStr(Остаток_рабочих_дней) + ")"
                
        ' Текущая строка офиса в отчете
        row_TS = rowByValue(FileNewVar, TS, officeNameInReport, 100, 100)
        
        ' Обрабатываем столбцы в fileTemplatesName в горизонтальном направлении
        ColumnCount = 1
        Do While (ColumnCount <= 100)
          
          ' Если находим # в ячейке
          If InStr(Workbooks(FileNewVar).Sheets(TS).Cells(1, ColumnCount).Value, "#") <> 0 Then
            
            ' Текущий продукт в форме отчета
            currProductName = Mid(Workbooks(FileNewVar).Sheets(TS).Cells(1, ColumnCount).Value, 2)
            
            ' Находим Текущий продукт на Лист8 для текущего офиса
            Row_Лист8 = getRowFromSheet8(officeNameInReport, currProductName)
            
            ' Расчет плана дня
            If Round(((ThisWorkbook.Sheets("Лист8").Cells(Row_Лист8, 9).Value - ThisWorkbook.Sheets("Лист8").Cells(Row_Лист8, 10).Value) / Остаток_рабочих_дней), 0) > 0 Then
              Workbooks(FileNewVar).Sheets(TS).Cells(row_TS, ColumnCount).Value = Round(((ThisWorkbook.Sheets("Лист8").Cells(Row_Лист8, 9).Value - ThisWorkbook.Sheets("Лист8").Cells(Row_Лист8, 10).Value) / Остаток_рабочих_дней), 0)
            Else
              Workbooks(FileNewVar).Sheets(TS).Cells(row_TS, ColumnCount).Value = 0
            End If

            ' Формат ячейки плана
            Workbooks(FileNewVar).Sheets(TS).Cells(row_TS, ColumnCount).NumberFormat = "#,##0"
            
          End If ' Если находим # в ячейке
          
          ' Следующий столбец
          ' Application.StatusBar = In_Product_Code + " " + In_officeNameInReport + ": " + CStr(rowCount) + "..."
          ColumnCount = ColumnCount + 1
          DoEventsInterval (ColumnCount)
        Loop
                
        ' Строка статуса
        Application.StatusBar = "Сохранение " + officeNameInReport + "..."
                
        ' Закрываем файл
        Workbooks(FileNewVar).Close SaveChanges:=True

        Application.StatusBar = "Сформирован файл " + ThisWorkbook.Path + "\Out\" + FileNewVar
                
        ' Переходим на окно DB
        ThisWorkbook.Sheets("Лист8").Activate

        ' Строка статуса
        Application.StatusBar = ""
        
    Next i


    Application.StatusBar = "Создание архива..."

    ' Запускаем архиватор этого файла, Справка https://www.dmosk.ru/miniinstruktions.php?mini=7zip-cmd
    ' Имя файла архива
    File7zipName = "Ежедневная_форма_отчёта_" + strДД_MM_YYYY(dateReport) + ".zip"
    Shell ("C:\Program Files\7-Zip\7z a -tzip -ssw -mx9 C:\Users\PROSCHAEVSF\Documents\#DB_Result\Out\" + File7zipName + " " + strFileNewVar_Office)

    Application.StatusBar = "Архив создан!"

    MsgBox ("Сформирован файл " + ThisWorkbook.Path + "\Out\" + FileNewVar + "!")

    Application.StatusBar = "Подготовка сообщения к отправке..."

    ' Отправка в почте в офисы
    ' Call Отправка_Lotus_Notes_Выгр_день_Лист8(ThisWorkbook.Path + "\Out\" + FileNewVar, DateReport)
    Call Отправка_Lotus_Notes_Выгр_день_Лист8(ThisWorkbook.Path + "\Out\" + File7zipName, dateReport)
      
    ' Строка статуса
    Application.StatusBar = "Сообщение отправлено!"
    ' Строка статуса
    Application.StatusBar = ""
      
  End If
  
End Sub



' Нумерация осуществляется через описанную функцию НумерацияПунктов
Function НумерацияПунктов(In_ProductName)
  
  ' Если In_ProductName начинается с "в т.ч. страховки к ПК" или ("           ИСЖ_МАСС")
  If (Mid(In_ProductName, 1, 6) = "в т.ч.") Or (Mid(In_ProductName, 1, 1) = " ") Then
    ' Выводим X.Y - как строку
    ' Мы не увеличиваем Порядковый_Номер_продукта_на_Лист8
    ' Дробная часть увеличивается на единицу
    Порядковый_Номер_продукта_Дробь_на_Лист8 = Порядковый_Номер_продукта_Дробь_на_Лист8 + 1
    ' Выводим как текст
    НумерацияПунктов = CStr(Порядковый_Номер_продукта_на_Лист8) + "." + CStr(Порядковый_Номер_продукта_Дробь_на_Лист8)
  Else
    ' Дробная часть обнуляется
    Порядковый_Номер_продукта_Дробь_на_Лист8 = 0
    ' Целую часть увеличиваем на единицу
    Порядковый_Номер_продукта_на_Лист8 = Порядковый_Номер_продукта_на_Лист8 + 1
    ' Выводим как X - это число
    НумерацияПунктов = CInt(Порядковый_Номер_продукта_на_Лист8)
  End If
  
End Function

' Заливка на Лист8 блока продукта синим цветом
Sub fillBlockBlue(In_RowStart, In_RowEnd)
    
    Range("B16:B19").Select
    Application.CutCopyMode = False
    Selection.Interior.Pattern = xlSolid
    Selection.Interior.PatternColorIndex = xlAutomatic
    Selection.Interior.ThemeColor = xlThemeColorAccent1
    Selection.Interior.TintAndShade = 0.799981688894314
    Selection.Interior.PatternTintAndShade = 0
    
End Sub

' Печать отчета
Sub Печать_Лист8()
    
  ' Запрос
  If MsgBox("Напечатать данные по офисам?", vbYesNo) = vbYes Then
    
      ' 1. ОО «Тюменский»
      Range("A" + CStr(getRowFromSheet8("ОО «Тюменский»", "ОО «Тюменский»")) + ":T" + CStr(getRowFromSheet8("ОО «Сургутский»", "ОО «Сургутский»") - 1)).Select
      Selection.PrintOut Copies:=1, Collate:=True
      
      ' 2. ОО «Сургутский»
      Range("A" + CStr(getRowFromSheet8("ОО «Сургутский»", "ОО «Сургутский»")) + ":T" + CStr(getRowFromSheet8("ОО «Нижневартовский»", "ОО «Нижневартовский»") - 1)).Select
      Selection.PrintOut Copies:=1, Collate:=True
      
      ' 3. ОО «Нижневартовский»
      Range("A" + CStr(getRowFromSheet8("ОО «Нижневартовский»", "ОО «Нижневартовский»")) + ":T" + CStr(getRowFromSheet8("ОО «Новоуренгойский»", "ОО «Новоуренгойский»") - 1)).Select
      Selection.PrintOut Copies:=1, Collate:=True
      
      ' 4. ОО «Новоуренгойский»
      Range("A" + CStr(getRowFromSheet8("ОО «Новоуренгойский»", "ОО «Новоуренгойский»")) + ":T" + CStr(getRowFromSheet8("ОО «Тарко-Сале»", "ОО «Тарко-Сале»") - 1)).Select
      Selection.PrintOut Copies:=1, Collate:=True
      
      ' 5. ОО «Тарко-Сале»
      Range("A" + CStr(getRowFromSheet8("ОО «Тарко-Сале»", "ОО «Тарко-Сале»")) + ":T" + CStr(getRowFromSheet8("Интегральный рейтинг по офисам", "Интегральный рейтинг по офисам") - 1)).Select
      Selection.PrintOut Copies:=1, Collate:=True
      
      ' 6. РОО Тюменский
      Range("A" + CStr(getRowFromSheet8("Итого по РОО «Тюменский»", "Итого по РОО «Тюменский»")) + ":T" + CStr(getRowFromSheet8("Итого по РОО «Тюменский»", "Итого по РОО «Тюменский»") + (getRowFromSheet8("ОО «Сургутский»", "ОО «Сургутский»") - getRowFromSheet8("ОО «Тюменский»", "ОО «Тюменский»")))).Select
      Selection.PrintOut Copies:=1, Collate:=True

      ' Переходим к началу Листа
      Call Лист8_к_началу
    
  End If
    
End Sub

' DB_Ипотека
Sub DB_Ипотека(In_ReportName_String, In_Sheets, In_officeNameInReport, In_Row_Лист8, In_N, In_Product_Name, In_Product_Code, In_Unit, In_Weight)
Dim dateDB As Date
  
  dateDB = CDate(Mid(Workbooks(In_ReportName_String).Sheets("Оглавление").Cells(1, 1).Value, 23, 10))
    
  ' Апдейтим таблицу BASE\Products
  Call Update_BASE_Products(In_Product_Name, In_Product_Code, In_Unit)
       
  ' В DB на Лист1 определяем столбцы
  ' DP4_отчет (Офисы): ОО "Тюменский", ОО2 "Нижневартовский", ОО2 "Тарко-Сале", ОО2"Новоуренгойский", ОО2"Сургутский"
  column_DB_Лист1_DP4_отчет = ColumnByValue(In_ReportName_String, "Лист1", "DP4_отчет", 100, 100)
  ' План_руб_Q_Ипотека
  column_DB_Лист1_План_руб_Q_Ипотека = ColumnByValue(In_ReportName_String, "Лист1", "План_руб_Q_Ипотека", 100, 100)
  ' Факт_руб_Q_Ипотека
  column_DB_Лист1_Факт_руб_Q_Ипотека = ColumnByValue(In_ReportName_String, "Лист1", "Факт_руб_Q_Ипотека", 100, 100)
  ' Прог_руб_Q_Ипотека
  column_DB_Лист1_Прог_руб_Q_Ипотека = ColumnByValue(In_ReportName_String, "Лист1", "Прог_руб_Q_Ипотека", 100, 100)
  
  rowCount_DB_Лист1 = 2
  Do While Not IsEmpty(Workbooks(In_ReportName_String).Sheets("Лист1").Cells(rowCount_DB_Лист1, column_DB_Лист1_DP4_отчет).Value)
  
    ' Если это текущий офис
    If InStr(Workbooks(In_ReportName_String).Sheets("Лист1").Cells(rowCount_DB_Лист1, column_DB_Лист1_DP4_отчет).Value, getShortNameOfficeByName(In_officeNameInReport)) <> 0 Then
      
      ' 1. Заносим на Лист 8 данные по ипотеке в квартал
      ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 1).NumberFormat = "@"
      ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 1).Value = In_N
      ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 1).HorizontalAlignment = xlCenter
      '
      ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 2).Value = In_Product_Name
      ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 2).HorizontalAlignment = xlLeft
      ' Вес выводим, если он не нулевой
      If In_Weight <> 0 Then
        ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 3).Value = In_Weight
        ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 3).NumberFormat = "0.0%"
        ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 3).HorizontalAlignment = xlCenter
      End If
      '
      ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 4).Value = In_Unit
      ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 4).HorizontalAlignment = xlCenter
      
      ' Квартал - план
      ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 5).Value = Workbooks(In_ReportName_String).Sheets("Лист1").Cells(rowCount_DB_Лист1, column_DB_Лист1_План_руб_Q_Ипотека).Value
      ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 5).NumberFormat = "#,##0"
      ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 5).HorizontalAlignment = xlRight

      ' Квартал - факт
      ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 6).Value = Workbooks(In_ReportName_String).Sheets("Лист1").Cells(rowCount_DB_Лист1, column_DB_Лист1_Факт_руб_Q_Ипотека).Value
      ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 6).NumberFormat = "#,##0"
      ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 6).HorizontalAlignment = xlRight

      ' Квартал - исполнение (в %)
      ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 7).Value = РассчетДоли(ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 5).Value, ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 6).Value, 3)
      ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 7).NumberFormat = "0%"
      ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 7).HorizontalAlignment = xlRight
        
      ' Квартал - прогноз
      If ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 5).Value <> 0 Then
        
        ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 8).Value = Workbooks(In_ReportName_String).Sheets("Лист1").Cells(rowCount_DB_Лист1, column_DB_Лист1_Прог_руб_Q_Ипотека).Value / Workbooks(In_ReportName_String).Sheets("Лист1").Cells(rowCount_DB_Лист1, column_DB_Лист1_План_руб_Q_Ипотека).Value
        ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 8).NumberFormat = "0%"
        ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 8).HorizontalAlignment = xlRight
      
        ' Окраска ячейки СФЕТОФОР ' 70. Заливка ячейки цветом "светофор" - значение In_Value в %, In_Target в %
        Call Full_Color_RangeII("Лист8", In_Row_Лист8, 8, ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 8).Value, 1)
        
      End If
      
      ' 2. Заносим в БД
      
      ' Заносим в Sales_Office
      '  Идентификатор ID_Rec:
      ID_RecVar = CStr(CStr(getNumberOfficeByName(In_officeNameInReport)) + "-" + strNQYY(dateDB) + "-" + In_Product_Code)
                        
      ' Текущие значения в месяце: Date_01 (N), Date_02 (O), Date_03 (P), Date_04 (Q), Date_05 Date_06 Date_07 Date_08 Date_09 Date_10 Date_11 Date_12 Date_13 Date_14 Date_15 Date_16 Date_17 Date_18 Date_19 Date_20 Date_21 Date_22 Date_23 Date_24 Date_25 Date_26 Date_27 Date_28 Date_29 Date_30 Date_31
      ' Номер месяца в квартале: 1-"", 2-"2", 3-"3"
      M_num = Nom_mes_quarter_str(dateDB)
      curr_Day_Month_Q = "Date" + M_num + "_" + Mid(dateDB, 1, 2)
                                      
      ' Вносим данные в BASE\Sales_Office по ПК.
      Call InsertRecordInBook("Sales_Office", "Лист1", "ID_Rec", ID_RecVar, _
                                            "ID_Rec", ID_RecVar, _
                                              "Оffice_Number", getNumberOfficeByName(In_officeNameInReport), _
                                                "Product_Name", In_Product_Name, _
                                                  "Оffice", In_officeNameInReport, _
                                                    "MMYY", strNQYY(dateDB), _
                                                      "Update_Date", dateDB, _
                                                       "Product_Code", In_Product_Code, _
                                                         "Plan", ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 5).Value, _
                                                            "Unit", In_Unit, _
                                                              "Fact", ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 6).Value, _
                                                                "Percent_Completion", ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 7).Value, _
                                                                  curr_Day_Month_Q, ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 6).Value, _
                                                                    "", "", _
                                                                      "", "", _
                                                                        "", "", _
                                                                          "", "", _
                                                                            "", "", _
                                                                              "", "", _
                                                                                "", "", _
                                                                                  "", "")

      
    End If
  
    ' Следующая запись
    rowCount_DB_Лист1 = rowCount_DB_Лист1 + 1
    DoEventsInterval (rowCount)
    
  Loop
  
End Sub


' Обработка исключения "Если в DB Лист не найден"
Sub в_DB_Лист_не_найден(In_StringInSheet)
          
  ' Заносим StringInSheet в переменную Строка_нет_листа_в_DB
  If InStr(Строка_нет_листа_в_DB, In_StringInSheet) = 0 Then
    
    Строка_нет_листа_в_DB = Строка_нет_листа_в_DB + In_StringInSheet + ", "
    ' Call в_DB_Лист_не_найден(StringInSheet)
    
    ' Если в DB Лист не найден - выводим сообщение
    MsgBox ("Не найден Лист " + Chr(34) + In_StringInSheet + Chr(34)) ' + " в " + ReportName_String)

  End If
  
End Sub


' Выдает строку "квартала" / "месяца"
Function Квартал_месяц_план() As String
    
    ' Если контроль месячный N7="1"
    If ThisWorkbook.Sheets("Лист8").Range("N7").Value = 1 Then
      ' Месяц
      Квартал_месяц_план = "месяца"
    Else
      ' Квартал
      Квартал_месяц_план = "квартала"
    End If

End Function

' Выдает строку с датой окончания квартала или месяца
Function Завершающий_день_квартал_месяц(In_Date) As String
    
    ' Если контроль месячный N7="1"
    If ThisWorkbook.Sheets("Лист8").Range("N7").Value = 1 Then
      ' Месяц
      Завершающий_день_квартал_месяц = CStr(Date_last_day_month(In_Date))
    Else
      ' Квартал
      Завершающий_день_квартал_месяц = CStr(Date_last_day_quarter(In_Date))
    End If

End Function

' Формирование_рейтинга_регионов
Sub Формирование_рейтинга_регионов(In_ReportName_String)
      
      ' Строка статуса
      Application.StatusBar = "Рейтинг регионов: формирование..."
      
      ' Открываем шаблон
      If Dir(ThisWorkbook.Path + "\Templates\" + "Интегральный рейтинг регионы.xlsx") <> "" Then
        ' Открываем шаблон
        TemplatesFileName = "Интегральный рейтинг регионы"
      End If
              
      ' Открываем шаблон из C:\Users\...\Documents\#VBA\DB_Result\Templates
      Workbooks.Open (ThisWorkbook.Path + "\Templates\" + TemplatesFileName + ".xlsx")
           
      ' Переходим на окно DB
      ThisWorkbook.Sheets("Лист8").Activate

      ' Дата DB
      dateDB_Лист8 = CDate(Mid(ThisWorkbook.Sheets("Лист8").Range("B5").Value, 52, 10))

      ' Имя нового файла
      FileDBName = "Рейтинг регионов " + Replace(CStr(dateDB_Лист8), ".", "-") + ".xlsx"
      
      ' Проверяем - если файл есть, то удаляем его
      Call deleteFile(ThisWorkbook.Path + "\Out\" + FileDBName)
      
      Workbooks(TemplatesFileName).SaveAs FileName:=ThisWorkbook.Path + "\Out\" + FileDBName, FileFormat:=xlOpenXMLWorkbook, createBackUp:=False
    
      ' В "A1" заголовок
      Workbooks(FileDBName).Sheets("Лист1").Range("A1").Value = "Интегральные рейтинги на " + CStr(dateDB_Лист8) + " г."
      
      ' Есть ли лист "Интегральный рейтинг_Регионы"?
      StringInSheet = "Интегральный рейтинг_Регионы"
      SheetName_String = FindNameSheet(In_ReportName_String, StringInSheet)
      If SheetName_String <> "" Then
    
        ' I. Копируем из DB данные по позициям 1-18 с листа "3. Интегральный рейтинг_Регионы" - ОО
        ' Сводную таблицу открывем в самом начале обработки DB, здесь ремарим Workbooks(In_ReportName_String).Sheets(SheetName_String).PivotTables("Сводная таблица1").PivotFields("DP3_отчет_new").PivotItems("Тюменский ОО1").ShowDetail = True
        
        row_Тюменский_ОО1 = rowByValue(In_ReportName_String, SheetName_String, "Тюменский ОО1", 100, 100)
        column_Тюменский_ОО1 = ColumnByValue(In_ReportName_String, SheetName_String, "Тюменский ОО1", 300, 300)
        row_Лист1 = 5

        ' I.1 Копируем блок 5 строк и 8 столбцов
        For i = row_Тюменский_ОО1 + 1 To row_Тюменский_ОО1 + 5
      
          row_Лист1 = row_Лист1 + 1
      
          For j = column_Тюменский_ОО1 To column_Тюменский_ОО1 + 7
            
            ' Workbooks(In_ReportName_String).Sheets(SheetName_String).Cells(i, j).Copy Destination:=Workbooks(FileDBName).Sheets("Лист1").Cells(i - 20, j - 1)
            Workbooks(In_ReportName_String).Sheets(SheetName_String).Cells(i, j).Copy Destination:=Workbooks(FileDBName).Sheets("Лист1").Cells(row_Лист1, j - 1)
            
            ' Если текущий столбец = (column_Тюменский_ОО1 + 7), то вставляем в BASE\Sales_Office
            If j = (column_Тюменский_ОО1 + 7) Then
              
              ' Вставляем в BASE\Sales_Office. Наименование офисов ОО2 "Тарко-Сале", ОО2 "Нижневартовский", ОО2"Новоуренгойский", ОО "Тюменский", ОО2"Сургутский"
              In_officeNameInReport = Workbooks(FileDBName).Sheets("Лист1").Cells(row_Лист1, 2).Value
              In_Product_Name = "Интегральный рейтинг"
              In_Product_Code = "ИнтРейт"
              Факт_ИР_Офиса = Round(Workbooks(FileDBName).Sheets("Лист1").Cells(row_Лист1, 9).Value * 100, 2)
      
              '  Идентификатор ID_Rec:
              ID_RecVar = CStr(CStr(getNumberOfficeByName2(In_officeNameInReport)) + "-" + strMMYY(dateDB_Лист8) + "-" + In_Product_Code)
            
              ' Текущие значения в месяце: Date_01 (N), Date_02 (O), Date_03 (P), Date_04 (Q), Date_05 Date_06 Date_07 Date_08 Date_09 Date_10 Date_11 Date_12 Date_13 Date_14 Date_15 Date_16 Date_17 Date_18 Date_19 Date_20 Date_21 Date_22 Date_23 Date_24 Date_25 Date_26 Date_27 Date_28 Date_29 Date_30 Date_31
              curr_Day_Month = "Date_" + Mid(dateDB_Лист8, 1, 2)
      
              ' Вносим данные в BASE\Sales_Office по ПК.
              Call InsertRecordInBook("Sales_Office", "Лист1", "ID_Rec", ID_RecVar, _
                                            "ID_Rec", ID_RecVar, _
                                              "Оffice_Number", getNumberOfficeByName2(In_officeNameInReport), _
                                                "Product_Name", In_Product_Name, _
                                                  "Оffice", getShortNameOfficeByName(In_officeNameInReport), _
                                                    "MMYY", strMMYY(dateDB_Лист8), _
                                                      "Update_Date", dateDB_Лист8, _
                                                        "Product_Code", In_Product_Code, _
                                                          "Plan", "100", _
                                                             "Unit", "%", _
                                                               "Fact", Факт_ИР_Офиса, _
                                                                 "Percent_Completion", "", _
                                                                   "Prediction", "", _
                                                                     "Percent_Prediction", "", _
                                                                       curr_Day_Month, Факт_ИР_Офиса, _
                                                                         "", "", _
                                                                           "", "", _
                                                                             "", "", _
                                                                               "", "", _
                                                                                 "", "", _
                                                                                   "", "")

              
            End If ' Если текущий столбец = (column_Тюменский_ОО1 + 7), то вставляем в BASE\Sales_Office
            
            DoEvents
            
          Next j
      
        Next i

        
        ' I.2 Форматируем рейтинг Офисов: Нумеруем пункты первого рейтинга, апдейтим наименования офисов, убираем дробную часть у %
        For i = 1 To 5
          
          ' № - Нумеруем пункты первого рейтинга
          Workbooks(FileDBName).Sheets("Лист1").Cells(i + 5, 1).Value = CStr(i)
          Workbooks(FileDBName).Sheets("Лист1").Cells(i + 5, 1).HorizontalAlignment = xlCenter
          ' Апдейтим наименования офисов
          Workbooks(FileDBName).Sheets("Лист1").Cells(i + 5, 2).Value = updateNameOfficeByName(Workbooks(FileDBName).Sheets("Лист1").Cells(i + 5, 2))
          ' Форматируем (убираем дробную чать у %)
          For j = 3 To 9
            Workbooks(FileDBName).Sheets("Лист1").Cells(i + 5, j).NumberFormat = "0%"
          Next j
          
        Next i
        
        ' I.3 Копируем рейтинг офисов в таблицу на Листе8
        row_Лист8_Интегральный_рейтинг_по_офисам = getRowFromSheet8("Интегральный рейтинг по офисам", "Интегральный рейтинг по офисам")
        ' Заголовки
        For j = 3 To 9
          ThisWorkbook.Sheets("Лист8").Cells(row_Лист8_Интегральный_рейтинг_по_офисам + 1, j).Value = Workbooks(FileDBName).Sheets("Лист1").Cells(4, j).Value
          ThisWorkbook.Sheets("Лист8").Cells(row_Лист8_Интегральный_рейтинг_по_офисам + 1, j).NumberFormat = "0%"
          ThisWorkbook.Sheets("Лист8").Cells(row_Лист8_Интегральный_рейтинг_по_офисам + 2, j).Value = Workbooks(FileDBName).Sheets("Лист1").Cells(5, j).Value
        Next j
        
        ' Показатели рейтинга ОО на Лист8 копируем (5 строк и 9 столбцов)
        For i = 1 To 5
          For j = 1 To 9
            
            Workbooks(FileDBName).Sheets("Лист1").Cells(i + 5, j).Copy Destination:=ThisWorkbook.Sheets("Лист8").Cells(row_Лист8_Интегральный_рейтинг_по_офисам + 2 + i, j)
            DoEvents
            
          Next j
        Next i
         

        ' II. Копируем из DB данные по позициям 1-18 с листа "3. Интегральный рейтинг_Регионы" - Филиалы
        
        ' Закрываем наши офисы
        ' Открываем сводную "Лист1" по показателям
        ' row_Тюменский_ОО1 = rowByValue(In_ReportName_String, SheetName_String, "Тюменский ОО1", 100, 100)
        ' column_Тюменский_ОО1 = ColumnByValue(In_ReportName_String, SheetName_String, "Тюменский ОО1", 300, 300)
        
        Workbooks(In_ReportName_String).Sheets(SheetName_String).Cells(row_Тюменский_ОО1, column_Тюменский_ОО1).ShowDetail = False

        
        
        ' Если такой лист есть - находим ячейку "№ п/п"
        row_№_пп = rowByValue(In_ReportName_String, SheetName_String, "№ п/п", 100, 100)
        column_№_пп = ColumnByValue(In_ReportName_String, SheetName_String, "№ п/п", 300, 300)
        row_Лист1 = 16
         
        ' Копируем блок 18 строк и 8 столбцов
        For i = row_№_пп + 1 To row_№_пп + 18
          
          row_Лист1 = row_Лист1 + 1
          
          For j = column_№_пп To column_№_пп + 8
            
            ' Workbooks(In_ReportName_String).Sheets(SheetName_String).Cells(i, j).Copy Destination:=Workbooks(FileDBName).Sheets("Лист1").Cells(i + 9, j - 1)
            Workbooks(In_ReportName_String).Sheets(SheetName_String).Cells(i, j).Copy Destination:=Workbooks(FileDBName).Sheets("Лист1").Cells(row_Лист1, j - 1)
            DoEvents
          Next j
        Next i
        
        ' Форматируем рейтинг регионов РФ: Убираем выделение наименование филиала (если это не Тюменский ОО1), убираем дробную часть у %
        For i = 17 To 34
          
          ' Убираем выделение наименование филиала (если это не Тюменский ОО1)
          If InStr(Workbooks(FileDBName).Sheets("Лист1").Cells(i, 2).Value, "Тюменский ОО1") = 0 Then
            
            For j = 1 To 9
              ' Формат ячейки в рейтинге Регионов
              Call Формат_ячейки_Рейтинга_регионов(FileDBName, "Лист1", i, j)
            Next j
            
          End If
                     
          ' Форматируем (убираем дробную чать у %) - для всех!
          For j = 3 To 9
            Workbooks(FileDBName).Sheets("Лист1").Cells(i, j).NumberFormat = "0%"
          Next j
          
        Next i
        
      
      Else
        ' Если в DB Лист не найден
        ' Сообщение, что листа нет!
      End If
   
    
      ' Закрытие файла
      Workbooks(FileDBName).Close SaveChanges:=True
    
      ' Строка статуса
      Application.StatusBar = "Рейтинг регионов: Отправка отчета..."
  
      ' Отправка файла в почту
      Call Отправка_Lotus_Notes_Рейтинг_регионов(ThisWorkbook.Path + "\Out\" + FileDBName, dateDB_Лист8)
  
      ' Строка статуса
      Application.StatusBar = "Рейтинг регионов: Отчет отправлен!"
      Application.StatusBar = ""
  
End Sub

' Отправка письма: отправляю шаблон самому себе для последующей отправки в сеть письма на его основе:
Sub Отправка_Lotus_Notes_Рейтинг_регионов(In_fileName, In_dateDB)
Dim темаПисьма, текстПисьма, hashTag, attachmentFile As String
Dim i As Byte
  
  
  ' Запрос
  ' If MsgBox("Отправить себе Шаблон письма с фокусами контроля '" + ПериодКонтроля + "'?", vbYesNo) = vbYes Then
    
    ' Тема письма - Тема:
    ' темаПисьма = ThisWorkbook.Sheets("Лист8").Cells(RowByValue(ThisWorkbook.Name, "Лист8", "Тема:", 100, 100), ColumnByValue(ThisWorkbook.Name, "Лист8", "Тема:", 100, 100) + 1).Value
    темаПисьма = "Рейтинг регионов и офисов на " + CStr(In_dateDB) + " г."

    ' hashTag - Хэштэг:
    ' hashTag = ThisWorkbook.Sheets("Лист8").Cells(RowByValue(ThisWorkbook.Name, "Лист8", "Хэштэг:", 100, 100), ColumnByValue(ThisWorkbook.Name, "Лист8", "Хэштэг:", 100, 100) + 1).Value
    ' hashTag - Хэштэг:
    hashTag = "#рейтинг_регионов"

    ' Файл-вложение (!!!)
    attachmentFile = In_fileName
    
    ' Текст письма
    текстПисьма = "" + Chr(13)
    текстПисьма = текстПисьма + "" + getFromAddrBook("УДО2,УДО3,УДО4,УДО5,НОРПиКО1,НОРПиКО2,НОРПиКО3,НОРПиКО4,НОРПиКО5,НОКП,РРКК,РИЦ", 2) + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "" + getFromAddrBook("РД", 2) + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "Уважаемые руководители," + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "Направляю рейтинг регионов по состоянию на " + strDDMM(In_dateDB) + " в прогнозе исполнения планов квартала." + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    ' текстПисьма = текстПисьма + "* - Показатель: ПК с БС и КСП будет сформирован дополнительно" + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    ' Визитка (подпись С Ув., )
    текстПисьма = текстПисьма + ПодписьВПисьме()
    ' Хэштег
    текстПисьма = текстПисьма + createBlankStr(27) + hashTag
    ' Вызов
    Call send_Lotus_Notes(темаПисьма, "Sergey Fedorovich Proschaev/Tyumen/PSBank/Ru", "Sergey Fedorovich Proschaev/Tyumen/PSBank/Ru", текстПисьма, attachmentFile)
  
    ' Зачеркнуть
    ' Call ЗачеркиваемТекстВячейке("Лист0", RangeByValue(ThisWorkbook.Name, "Лист0", "DashBoard (при наличии)", 100, 100))
  
    ' Сообщение
    ' MsgBox ("Письмо отправлено!")
     
  ' End If
  
End Sub

' Форматирование ячейки Рейтинга регионов
Sub Формат_ячейки_Рейтинга_регионов(In_Workbooks, In_Sheets, In_Row, In_Col)
  
  Workbooks(In_Workbooks).Sheets(In_Sheets).Cells(In_Row, In_Col).Font.Bold = False
  Workbooks(In_Workbooks).Sheets(In_Sheets).Cells(In_Row, In_Col).Borders(xlDiagonalDown).LineStyle = xlNone
  Workbooks(In_Workbooks).Sheets(In_Sheets).Cells(In_Row, In_Col).Borders(xlDiagonalUp).LineStyle = xlNone
  Workbooks(In_Workbooks).Sheets(In_Sheets).Cells(In_Row, In_Col).Borders(xlEdgeLeft).LineStyle = xlNone
  Workbooks(In_Workbooks).Sheets(In_Sheets).Cells(In_Row, In_Col).Borders(xlEdgeTop).LineStyle = xlNone
  Workbooks(In_Workbooks).Sheets(In_Sheets).Cells(In_Row, In_Col).Borders(xlEdgeBottom).LineStyle = xlNone
  Workbooks(In_Workbooks).Sheets(In_Sheets).Cells(In_Row, In_Col).Borders(xlEdgeRight).LineStyle = xlNone
  Workbooks(In_Workbooks).Sheets(In_Sheets).Cells(In_Row, In_Col).Borders(xlInsideVertical).LineStyle = xlNone
  Workbooks(In_Workbooks).Sheets(In_Sheets).Cells(In_Row, In_Col).Borders(xlInsideHorizontal).LineStyle = xlNone

End Sub

' Отклонения по офисам сообщение
Sub Отклонения_по_офисам_with_Msg()
  
  ' Запрос
  If MsgBox("Сформировать отклонения по офисам?", vbYesNo) = vbYes Then
    Call Отклонения_по_офисам
    MsgBox ("Письма с отклонениями сформированы!")
  End If

End Sub


' Отклонения по офисам сообщение
Sub Отклонения_по_офисам()
Dim темаПисьма, текстПисьма, hashTag, attachmentFile As String
' Dim симв As Byte

  ' Строка статуса
  Application.StatusBar = "Отклонения по Офисам..."

  ' Открываем BASE\Sales
  OpenBookInBase ("MBO")
  ThisWorkbook.Sheets("Лист8").Activate

  OpenBookInBase ("TargetWeek")
  ThisWorkbook.Sheets("Лист8").Activate

  OpenBookInBase ("Products")
  ThisWorkbook.Sheets("Лист8").Activate


  ' Проходим по 5-ти офисам
  ' Заголовки
  For i = 1 To 5 ' 5 - на период отладки
    ' Номера офисов от 1 до 5
    Select Case i
      Case 1 ' ОО «Тюменский»
        officeNameInReport = "ОО «Тюменский»"
        Адресат_письма = getFromAddrBook("НОРПиКО1,НОКП,РИЦ", 2)
      Case 2 ' ОО «Сургутский»
        officeNameInReport = "ОО «Сургутский»"
        Адресат_письма = getFromAddrBook("УДО" + CStr(i), 2)
      Case 3 ' ОО «Нижневартовский»
        officeNameInReport = "ОО «Нижневартовский»"
        Адресат_письма = getFromAddrBook("УДО" + CStr(i), 2)
      Case 4 ' ОО «Новоуренгойский»
        officeNameInReport = "ОО «Новоуренгойский»"
        Адресат_письма = getFromAddrBook("УДО" + CStr(i), 2)
      Case 5 ' ОО «Тарко-Сале»
        officeNameInReport = "ОО «Тарко-Сале»"
        Адресат_письма = getFromAddrBook("УДО" + CStr(i), 2)
    End Select
  
    ' Открываем файл с шаблоном "Цели на неделю.xlsx"
    If Dir(ThisWorkbook.Path + "\Templates\" + "Цели на неделю.xlsx") <> "" Then
      ' Открываем шаблон Templates\Ежедневный отчет по продажам
      TemplatesFileName = "Цели на неделю"
    End If
              
    ' Открываем шаблон Протокола из C:\Users\...\Documents\#VBA\DB_Result\Templates
    Workbooks.Open (ThisWorkbook.Path + "\Templates\" + TemplatesFileName + ".xlsx")
           
    ' Переходим на окно DB
    ThisWorkbook.Sheets("Лист8").Activate

    ' Имя нового файла
    FileTargetWeekName = getShortNameOfficeByName(officeNameInReport) + "_" + CStr(dateDB_Лист_8) + ".xlsx"
    
    ' Проверяем - если файл есть, то удаляем его
    Call deleteFile(ThisWorkbook.Path + "\Out\" + FileTargetWeekName)
    
    ' Переименовываем шаблон в новый файл
    Workbooks(TemplatesFileName).SaveAs FileName:=ThisWorkbook.Path + "\Out\" + FileTargetWeekName, FileFormat:=xlOpenXMLWorkbook, createBackUp:=False
    ' ThisWorkbook.Sheets("Лист8").Range("Q3").Value = ThisWorkbook.Path + "\Out\" + FileTargetWeekName
    
    ' Заголовки в новом файле (период, офис)
    ' Цели на неделю в A2
    Workbooks(FileTargetWeekName).Sheets("Лист1").Range("A2").Value = "Цели на неделю c " + strDDMM(weekStartDate(Date)) + " по " + CStr(weekEndDate(Date) - 2)
    ' Офис
    Workbooks(FileTargetWeekName).Sheets("Лист1").Range("B3").Value = officeNameInReport
    ' Квартал
    Workbooks(FileTargetWeekName).Sheets("Лист1").Range("E4").Value = ThisWorkbook.Sheets("Лист8").Range("E8").Value
    ' Факт
    Workbooks(FileTargetWeekName).Sheets("Лист1").Range("F5").Value = ThisWorkbook.Sheets("Лист8").Range("F9").Value
    ' Цели на неделю
    Workbooks(FileTargetWeekName).Sheets("Лист1").Range("I4").Value = "Цели на неделю " + strDDMM(weekStartDate(Date)) + "-" + strDDMM(weekEndDate(Date) - 2)
    ' Итоги прошедшей недели
    Workbooks(FileTargetWeekName).Sheets("Лист1").Range("L4").Value = "Итоги прошедшей недели " + strDDMM(weekStartDate(Date - 7)) + "-" + strDDMM(weekEndDate(Date - 7) - 2)
    
    ' Находим отклонения по позициям ИР
  
    ' *** Формируем письмо ***
    
    ' Тема письма
    темаПисьма = "Исполнение БП " + officeNameInReport + " по РБ на " + CStr(dateDB_Лист_8)

    ' hashTag - Хэштэг:
    hashTag = "#БП_РБ #БП_" + getShortNameOfficeByName(officeNameInReport)

    ' Файл-вложение
    attachmentFile = ThisWorkbook.Path + "\Out\" + FileTargetWeekName
    
    ' Текст письма
    текстПисьма = "" + Chr(13)
    ' текстПисьма = текстПисьма + "" + getFromAddrBook("УДО" + CStr(i), 2) + Chr(13)
    текстПисьма = текстПисьма + "" + Адресат_письма + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "" + getFromAddrBook("РД", 2) + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    
    
    ' *** Обращение ***
    ' Если текущая дата четная, то Имя Отчество
    ' текстПисьма = текстПисьма + getFromAddrBook("УДО" + CStr(i), 6) + Добрый_утро_день_вечер(Time()) + Chr(13)
    ' если нечетная, то по Имени
    текстПисьма = текстПисьма + getFromAddrBook("УДО" + CStr(i), 6) + ", " + Добрый_утро_день_вечер(Time(), "д") + ", " + Chr(13)
    ' *** Обращение ***
    
    ' *** Начало письма ***
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + Промежуточный_срез_результатов + " " + officeNameInReport + " на " + strDDMM(dateDB_Лист_8) + ":" + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    
    ' Номер пункта
    Номер_пункта = 0
    
    ' 1. Показатели в "зеленой зоне"
    Показатели_Зеленая_зона_Q_Var = Показатели_Зеленая_Желтая_Красная_зона_Q(officeNameInReport, 1, 1000, FileTargetWeekName) ' Показатели_Зеленая_зона_Q(officeNameInReport)
    If Показатели_Зеленая_зона_Q_Var <> "" Then
      Номер_пункта = Номер_пункта + 1
      текстПисьма = текстПисьма + CStr(Номер_пункта) + ". Показатели в ЗЕЛЕНОЙ ЗОНЕ: " + Показатели_Зеленая_зона_Q_Var + Chr(13)
    Else
      ' Если нет ни одного показателя в зеленой зоне
      Номер_пункта = Номер_пункта + 1
      текстПисьма = текстПисьма + CStr(Номер_пункта) + ". Нет ни одного показателя в ЗЕЛЕНОЙ ЗОНЕ!" + Chr(13)
    End If
    текстПисьма = текстПисьма + "" + Chr(13)
    
    ' 2. Показатели в "желтой зоне"
    Номер_пункта = Номер_пункта + 1
    Показатели_Желтая_зона_Q_Var = Показатели_Зеленая_Желтая_Красная_зона_Q(officeNameInReport, 0.9, 1, FileTargetWeekName)
    текстПисьма = текстПисьма + CStr(Номер_пункта) + ". Показатели в ЖЕЛТОЙ ЗОНЕ: " + Показатели_Желтая_зона_Q_Var + Chr(13)
    ' текстПисьма = текстПисьма + "- " + Показатели_Желтая_зона_Q(officeNameInReport) + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    
    ' 3. Показатели в "красной зоне"
    Номер_пункта = Номер_пункта + 1
    Показатели_Красная_зона_Q_Var = Показатели_Зеленая_Желтая_Красная_зона_Q(officeNameInReport, 0, 0.8999, FileTargetWeekName)
    текстПисьма = текстПисьма + CStr(Номер_пункта) + ". Показатели в КРАСНОЙ ЗОНЕ: " + Показатели_Красная_зона_Q_Var + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    
    ' 4. Результативность сотрудников
    ' Номер_пункта = Номер_пункта + 1
    ' текстПисьма = текстПисьма + CStr(Номер_пункта) + ". Результативность сотрудников:" + Chr(13)
    ' текстПисьма = текстПисьма + "" + Chr(13)
    
    ' 5. Отработка МК
    ' Номер_пункта = Номер_пункта + 1
    ' текстПисьма = текстПисьма + CStr(Номер_пункта) + ". Отработка МК:" + Chr(13)
    ' текстПисьма = текстПисьма + "" + Chr(13)
    
    ' 6. Цели на неделю
    Номер_пункта = Номер_пункта + 1
    текстПисьма = текстПисьма + CStr(Номер_пункта) + ". Цели на неделю и факт исполнения поручений прошлой недели:" + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    
    
    
    ' *** Начало письма ***
    
    
    
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    ' текстПисьма = текстПисьма + "* - Показатель: ПК с БС и КСП будет сформирован дополнительно" + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    ' Визитка (подпись С Ув., )
    текстПисьма = текстПисьма + ПодписьВПисьме()
    ' Хэштег
    текстПисьма = текстПисьма + createBlankStr(27) + hashTag
    
    ' Закрытие файла
    Workbooks(FileTargetWeekName).Close SaveChanges:=True
    
    ' Вызов
    Call send_Lotus_Notes(темаПисьма, "Sergey Fedorovich Proschaev/Tyumen/PSBank/Ru", "Sergey Fedorovich Proschaev/Tyumen/PSBank/Ru", текстПисьма, attachmentFile)
    
    ' *** Формируем письмо ***
  
  Next i
  
  ' Закрываем таблицу MBO
  ' Закрываем BASE\Sales
  CloseBook ("MBO")
  ThisWorkbook.Sheets("Лист8").Activate

  ' Закрываем таблицу MBO
  ' Закрываем BASE\Sales
  CloseBook ("TargetWeek")
  
  ' Закрываем BASE\Sales
  OpenBookInBase ("Products")
  
  ThisWorkbook.Sheets("Лист8").Activate

  ' Строка статуса
  Application.StatusBar = ""
  
End Sub

' Отклонения по офисам сообщение
Sub Отклонения_по_ОКП()
Dim темаПисьма, текстПисьма, hashTag, attachmentFile As String
' Dim симв As Byte

  ' Строка статуса
  Application.StatusBar = "Отклонения по ОКП..."

  ' Открываем BASE\Sales
  OpenBookInBase ("MBO")
  ThisWorkbook.Sheets("Лист8").Activate

  OpenBookInBase ("TargetWeek")
  ThisWorkbook.Sheets("Лист8").Activate

  ' Наименование офиса
  officeNameInReport = "ОКП (+ПВО)"
  
  ' Открываем файл с шаблоном "Цели на неделю.xlsx"
  If Dir(ThisWorkbook.Path + "\Templates\" + "Цели на неделю.xlsx") <> "" Then
    ' Открываем шаблон Templates\Ежедневный отчет по продажам
    TemplatesFileName = "Цели на неделю"
  End If
              
  ' Открываем шаблон Протокола из C:\Users\...\Documents\#VBA\DB_Result\Templates
  Workbooks.Open (ThisWorkbook.Path + "\Templates\" + TemplatesFileName + ".xlsx")
           
  ' Переходим на окно DB
  ThisWorkbook.Sheets("Лист8").Activate

  ' Имя нового файла
  FileTargetWeekName = officeNameInReport + "_" + CStr(dateDB_Лист_8) + ".xlsx"
    
  ' Проверяем - если файл есть, то удаляем его
  Call deleteFile(ThisWorkbook.Path + "\Out\" + FileTargetWeekName)
    
  ' Переименовываем шаблон в новый файл
  Workbooks(TemplatesFileName).SaveAs FileName:=ThisWorkbook.Path + "\Out\" + FileTargetWeekName, FileFormat:=xlOpenXMLWorkbook, createBackUp:=False
  
    
  ' Заголовки в новом файле (период, офис)
  ' Цели на неделю в A2
  Workbooks(FileTargetWeekName).Sheets("Лист1").Range("A2").Value = "Цели на неделю c " + strDDMM(weekStartDate(Date)) + " по " + CStr(weekEndDate(Date) - 2)
  ' Офис
  Workbooks(FileTargetWeekName).Sheets("Лист1").Range("B3").Value = officeNameInReport
  ' Квартал
  Workbooks(FileTargetWeekName).Sheets("Лист1").Range("E4").Value = ThisWorkbook.Sheets("Лист8").Range("E8").Value
  ' Факт
  Workbooks(FileTargetWeekName).Sheets("Лист1").Range("F5").Value = ThisWorkbook.Sheets("Лист8").Range("F9").Value
  ' Цели на неделю
  Workbooks(FileTargetWeekName).Sheets("Лист1").Range("I4").Value = "Цели на неделю " + strDDMM(weekStartDate(Date)) + "-" + strDDMM(weekEndDate(Date) - 2)
  ' Итоги прошедшей недели
  Workbooks(FileTargetWeekName).Sheets("Лист1").Range("L4").Value = "Итоги прошедшей недели " + strDDMM(weekStartDate(Date - 7)) + "-" + strDDMM(weekEndDate(Date - 7) - 2)

  ' Находим отклонения по позициям ИР
  
  ' *** Формируем письмо ***
    
  ' Тема письма
  темаПисьма = "Исполнение БП " + officeNameInReport + " на " + CStr(dateDB_Лист_8)

  ' hashTag - Хэштэг:
  hashTag = "#БП_РБ #БП_" + officeNameInReport

  ' Файл-вложение
  attachmentFile = ThisWorkbook.Path + "\Out\" + FileTargetWeekName
    
  ' Текст письма
  текстПисьма = "" + Chr(13)
  текстПисьма = текстПисьма + "" + getFromAddrBook("НОКП" + CStr(i), 2) + ", Marat Albertovich Timergaliev/Tyumen/PSBank/Ru, Yuriy Vladimirovich Martyuchenko/Tyumen/PSBank/Ru" + Chr(13)
  текстПисьма = текстПисьма + "" + Chr(13)
  текстПисьма = текстПисьма + "" + getFromAddrBook("РД", 2) + Chr(13)
  текстПисьма = текстПисьма + "" + Chr(13)
    
    
  ' *** Обращение ***
  ' Если текущая дата четная, то Имя Отчество
  ' текстПисьма = текстПисьма + getFromAddrBook("УДО" + CStr(i), 6) + Добрый_утро_день_вечер(Time()) + Chr(13)
  ' если нечетная, то по Имени
  ' текстПисьма = текстПисьма + getFromAddrBook("ОКП" + CStr(i), 6) + ", " + Добрый_утро_день_вечер(Time(), "д") + ", " + Chr(13)
  
  текстПисьма = текстПисьма + "Уважаемые руководители, " + Chr(13)
  
  
  ' *** Обращение ***
    
  ' *** Начало письма ***
  текстПисьма = текстПисьма + "" + Chr(13)
  текстПисьма = текстПисьма + Промежуточный_срез_результатов + " " + officeNameInReport + " (в т.ч. канал ПВО) на " + strDDMM(dateDB_Лист_8) + " и цели на неделю:" + Chr(13)
  текстПисьма = текстПисьма + "" + Chr(13)
    
  ' Номер пункта
  Номер_пункта = 0
    
  ' Контроля для ОКП берем по "Итого по РОО «Тюменский»":
  ' 1) "Зарплатные карты 18+"
  ' 2) "Портфель ЗП 18+, шт._Квартал "
  ' 3) "в т.ч. ПК DSA"
  ' 4) "           КК DSA"
  ' 5) "           КК к ЗП"
  
  ' Находим строку заголовка "Показатель" в выходном файле
  row_Лист1_Показатель = rowByValue(FileTargetWeekName, "Лист1", "Показатель", 50, 50)
  
  For i = 1 To 5
  
    ' Номера офисов от 1 до 5
    Select Case i
      Case 1 ' "Зарплатные карты 18+"
        productName = "Зарплатные карты 18+"
        Наимнование_показателя = "Зарплатные карты 18+"
        row_вес_продукта = getRowFromSheet8("Итого по РОО «Тюменский»", "Зарплатные карты 18+")
      Case 2 ' "Портфель ЗП 18+, шт._Квартал "
        productName = "Портфель ЗП 18+, шт._Квартал "
        Наимнование_показателя = "Портфель ЗП 18+, шт._Квартал "
        row_вес_продукта = getRowFromSheet8("Итого по РОО «Тюменский»", "Портфель ЗП 18+, шт._Квартал ")
      Case 3 ' "в т.ч. ПК DSA"
        productName = "в т.ч. ПК DSA"
        Наимнование_показателя = "Потреб кредиты DSA"
        row_вес_продукта = getRowFromSheet8("Итого по РОО «Тюменский»", "Потребительские кредиты")
      Case 4 ' "           КК DSA"
        productName = "           КК DSA"
        Наимнование_показателя = "Кредитные карты DSA"
        row_вес_продукта = getRowFromSheet8("Итого по РОО «Тюменский»", "Кредитные карты (актив.)")
      Case 5 ' "           КК к ЗП"
        productName = "           КК к ЗП"
        Наимнование_показателя = "Кредитные карты к ЗП"
        row_вес_продукта = getRowFromSheet8("Итого по РОО «Тюменский»", "Кредитные карты (актив.)")
    End Select

    ' Находим номер строки где расположен продукт на Лист8 для РОО Тюменский
    row_Лист8_productName = getRowFromSheet8("Итого по РОО «Тюменский»", productName)
  
    ' Вставляем #1
    ' Номер
    Workbooks(FileTargetWeekName).Sheets("Лист1").Cells(row_Лист1_Показатель + 1 + i, 1).Value = i
        
    ' Наименование показателя
    ThisWorkbook.Sheets("Лист8").Cells(row_Лист8_productName, 2).Copy Destination:=Workbooks(FileTargetWeekName).Sheets("Лист1").Cells(row_Лист1_Показатель + 1 + i, 2)
    Workbooks(FileTargetWeekName).Sheets("Лист1").Cells(row_Лист1_Показатель + 1 + i, 2).Value = Наимнование_показателя
        
    ' Вес
    ThisWorkbook.Sheets("Лист8").Cells(row_вес_продукта, 3).Copy Destination:=Workbooks(FileTargetWeekName).Sheets("Лист1").Cells(row_Лист1_Показатель + 1 + i, 3)
    Call Формат_ячейки_Цели_на_неделю(FileTargetWeekName, "Лист1", row_Лист1_Показатель + 1 + i, 3)
        
    ' Ед.изм.
    ThisWorkbook.Sheets("Лист8").Cells(row_Лист8_productName, 4).Copy Destination:=Workbooks(FileTargetWeekName).Sheets("Лист1").Cells(row_Лист1_Показатель + 1 + i, 4)
        
    ' План
    ThisWorkbook.Sheets("Лист8").Cells(row_Лист8_productName, 5).Copy Destination:=Workbooks(FileTargetWeekName).Sheets("Лист1").Cells(row_Лист1_Показатель + 1 + i, 5)
        
    ' Факт
    ThisWorkbook.Sheets("Лист8").Cells(row_Лист8_productName, 6).Copy Destination:=Workbooks(FileTargetWeekName).Sheets("Лист1").Cells(row_Лист1_Показатель + 1 + i, 6)
        
    ' Исполнение Q
    ThisWorkbook.Sheets("Лист8").Cells(row_Лист8_productName, 7).Copy Destination:=Workbooks(FileTargetWeekName).Sheets("Лист1").Cells(row_Лист1_Показатель + 1 + i, 7)
        
    ' Если это Портфель или Пассивы, то Прогноз=Факт
    If (InStr(ThisWorkbook.Sheets("Лист8").Cells(row_Лист8_productName, 2).Value, "Портфель") <> 0) Or (InStr(ThisWorkbook.Sheets("Лист8").Cells(row_Лист8_productName, 2).Value, "Пассивы") <> 0) Or (InStr(ThisWorkbook.Sheets("Лист8").Cells(row_Лист8_productName, 2).Value, "Инвест") <> 0) Then
      ' Факт
      column_Прогноз = 7
    Else
      ' Расчетный прогноз
      column_Прогноз = 20
    End If
        
    ' Прогноз (если прогноз не берем со  столбца 7 на Лист8)
    If column_Прогноз <> 7 Then
          
      ThisWorkbook.Sheets("Лист8").Cells(row_Лист8_productName, column_Прогноз).Copy Destination:=Workbooks(FileTargetWeekName).Sheets("Лист1").Cells(row_Лист1_Показатель + 1 + i, 8)
                    
    End If
        
    ' Форматируем Исполнение квартала и Прогноз (цвет)
    ' Если "Прогноз" есть, то заливаем его Светофор, а в "Исп. Q" убираем цвет
    If Not IsEmpty(Workbooks(FileTargetWeekName).Sheets("Лист1").Cells(row_Лист1_Показатель + 1 + i, 8).Value) Then
      
      ' Убираем заливку в "Исп. Q"
      Call Убрать_заливку_цветом(FileTargetWeekName, "Лист1", row_Лист1_Показатель + 1 + i, 7)
      ' Окраска ячейки СФЕТОФОР ' 70. Заливка ячейки цветом "светофор" - значение In_Value в %, In_Target в %
      Call Full_Color_RangeV(FileTargetWeekName, "Лист1", row_Лист1_Показатель + 1 + i, 8, Workbooks(FileTargetWeekName).Sheets("Лист1").Cells(row_Лист1_Показатель + 1 + i, 8).Value, 1)
    End If
        
    ' Цель недели
    ' Формируем расчет Цель на неделю:
    Дата_начала_недели_по_DB = CDate(Mid(ThisWorkbook.Sheets("Лист8").Range("B5").Value, 52, 10))
    Дата_окончания_недели_по_DB = Дата_начала_недели_по_DB + 7
    ' Дата_окончания_недели_по_DB = Дата_начала_недели_по_DB + 6
          
    ' Если неделя переходит на новый квартал, то берем последний день квартала
    If Дата_окончания_недели_по_DB > Date_last_day_quarter(Дата_начала_недели_по_DB) Then
      Дата_окончания_недели_по_DB = Date_last_day_quarter(Дата_начала_недели_по_DB)
    End If
          
    ' Если показатель перевыполняется, то задача сохранить динамику, иначе стремимся на 100%
    If Workbooks(FileTargetWeekName).Sheets("Лист1").Cells(row_Лист1_Показатель + 1 + i, 8).Value > 1 Then
            
      ' Проверять, чтобы было не пусто! иначе ошибка
      ' Цель_прогноза_квартала = Workbooks(FileTargetWeekName).Sheets("Лист1").Cells(row_Лист1_Показатель, 8).Value / 100
        
    Else
      Цель_прогноза_квартала = 1
    End If
          
    ' Тестируемая функция, которая работает! ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 20).Value = Прогноз_квартала_проц(dateDB, ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 5).Value, ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 6).Value, 5, 0)
    Факт_на_дату_для_прогноза_квартала_Var = Факт_на_дату_для_прогноза_квартала(Дата_окончания_недели_по_DB, _
                                                                                      Workbooks(FileTargetWeekName).Sheets("Лист1").Cells(row_Лист1_Показатель + 1 + i, 5).Value, _
                                                                                        Цель_прогноза_квартала, _
                                                                                          5, _
                                                                                            0)
                                                                                            
    Необходимый_прирост_за_неделю = Факт_на_дату_для_прогноза_квартала_Var - Workbooks(FileTargetWeekName).Sheets("Лист1").Cells(row_Лист1_Показатель + 1 + i, 6).Value
          
    ' Пока прописываем обязательства на неделю для всех показателей<100% в прогнозе
    If Необходимый_прирост_за_неделю > 0 Then
          
      ' Если это Портфель или Пассивы, то Прогноз=Факт
      If (InStr(ThisWorkbook.Sheets("Лист8").Cells(row_Лист8_productName, 2).Value, "Портфель") = 0) And (InStr(ThisWorkbook.Sheets("Лист8").Cells(row_Лист8_productName, 2).Value, "Пассивы") = 0) And (InStr(ThisWorkbook.Sheets("Лист8").Cells(row_Лист8_productName, 2).Value, "Инвест") = 0) Then
          
        ' Копируем параметры ячейки (из План)
        Workbooks(FileTargetWeekName).Sheets("Лист1").Cells(row_Лист1_Показатель + 1 + i, 5).Copy Destination:=Workbooks(FileTargetWeekName).Sheets("Лист1").Cells(row_Лист1_Показатель + 1 + i, 9)
        Workbooks(FileTargetWeekName).Sheets("Лист1").Cells(row_Лист1_Показатель + 1 + i, 9) = Необходимый_прирост_за_неделю
          
        ' Расчетный прогноз на конец недели
        Workbooks(FileTargetWeekName).Sheets("Лист1").Cells(row_Лист1_Показатель + 1 + i, 10).Value = Прогноз_квартала_проц(Дата_окончания_недели_по_DB, _
                                                                                                                     Workbooks(FileTargetWeekName).Sheets("Лист1").Cells(row_Лист1_Показатель + 1 + i, 5).Value, _
                                                                                                                       Workbooks(FileTargetWeekName).Sheets("Лист1").Cells(row_Лист1_Показатель + 1 + i, 6).Value + Workbooks(FileTargetWeekName).Sheets("Лист1").Cells(row_Лист1_Показатель + 1 + i, 9).Value, _
                                                                                                                        5, _
                                                                                                                         0)
        Workbooks(FileTargetWeekName).Sheets("Лист1").Cells(row_Лист1_Показатель + 1 + i, 10).NumberFormat = "0%"
        Workbooks(FileTargetWeekName).Sheets("Лист1").Cells(row_Лист1_Показатель + 1 + i, 10).HorizontalAlignment = xlRight
                              
      End If
                              
    End If
    
    ' Вставляем сюда прирост за прошедшую неделю по текущему показателю!
    ' "Итоги прошедшей недели " + strDDMM(weekStartDate(Date - 7)) + "-" + strDDMM(weekEndDate(Date - 7) - 2)
    ' Исполнение Q
    Workbooks(FileTargetWeekName).Sheets("Лист1").Cells(row_Лист1_Показатель + 1 + i, 12).Value = ThisWorkbook.Sheets("Лист8").Cells(row_Лист8_productName, 16).Value
    Workbooks(FileTargetWeekName).Sheets("Лист1").Cells(row_Лист1_Показатель + 1 + i, 12).NumberFormat = "#,##0"
 
  Next i
  ' *** Начало письма ***
    
    
    
  текстПисьма = текстПисьма + "" + Chr(13)
  текстПисьма = текстПисьма + "" + Chr(13)
  ' текстПисьма = текстПисьма + "* - Показатель: ПК с БС и КСП будет сформирован дополнительно" + Chr(13)
  текстПисьма = текстПисьма + "" + Chr(13)
  текстПисьма = текстПисьма + "" + Chr(13)
  ' Визитка (подпись С Ув., )
  текстПисьма = текстПисьма + ПодписьВПисьме()
  ' Хэштег
  текстПисьма = текстПисьма + createBlankStr(27) + hashTag
    
  ' Закрытие файла
  Workbooks(FileTargetWeekName).Close SaveChanges:=True
    
  ' Вызов
  Call send_Lotus_Notes(темаПисьма, "Sergey Fedorovich Proschaev/Tyumen/PSBank/Ru", "Sergey Fedorovich Proschaev/Tyumen/PSBank/Ru", текстПисьма, attachmentFile)
    
  ' *** Формируем письмо ***
  
  ' Закрываем таблицу MBO
  ' Закрываем BASE\Sales
  CloseBook ("MBO")
  ThisWorkbook.Sheets("Лист8").Activate

  ' Закрываем таблицу MBO
  ' Закрываем BASE\Sales
  CloseBook ("TargetWeek")
  ThisWorkbook.Sheets("Лист8").Activate
  
  ' Строка статуса
  Application.StatusBar = ""
    
End Sub


' Коды символов
Sub AscW_ChrW()
  ThisWorkbook.Sheets("Лист8").Range("X7").Value = AscW(ThisWorkbook.Sheets("Лист8").Range("V7").Value)
End Sub


' Показатель MBO
Function Показатель_MBO(In_Product_Name) As Boolean
      
  Показатель_MBO = False
  
  ' Открываем таблицу MBO
  ' Открываем BASE\Sales
  ' OpenBookInBase ("MBO")
  ' ThisWorkbook.Sheets("Лист8").Activate
  
  показатель_найден = False
  column_Product_Name = 2
  rowCount = 2
  
  ' Обрабатываем Лист - ищем Сначала Офис, если находим офис, то ищем позицию с наименованием продукта
  Do While (Not IsEmpty(Workbooks("MBO").Sheets("Лист1").Cells(rowCount, column_Product_Name).Value)) And (показатель_найден = False)
  
    ' Проверяем параметр с весом
    If Workbooks("MBO").Sheets("Лист1").Cells(rowCount, column_Product_Name).Value = In_Product_Name Then
      Показатель_MBO = True
      показатель_найден = True
    End If
    
    ' Следующая запись
    rowCount = rowCount + 1
  Loop

  ' Закрываем таблицу MBO
  ' Закрываем BASE\Sales
  ' CloseBook ("MBO")
  ' ThisWorkbook.Sheets("Лист8").Activate
  
End Function


' Показатели в Зеленой_Желтой_Красной зоны
Function Показатели_Зеленая_Желтая_Красная_зона_Q(In_officeNameInReport, In_LowRange, In_UpperRange, In_FileTargetWeekName) As String
      
  Показатели_Зеленая_Желтая_Красная_зона_Q = ""
  
  ' Берем с листа ОО «Тюменский»
  rowCount = rowByValue(ThisWorkbook.Name, "Лист8", In_officeNameInReport, 1000, 100) + 3
  
  ' Берем с Лист1
  row_Показатель_Лист1 = rowByValue(In_FileTargetWeekName, "Лист1", "Показатель", 100, 100) + 2
  число_строк_Лист1 = 1
  
  ' Обрабатываем Лист - ищем Сначала Офис, если находим офис, то ищем позицию с наименованием продукта
  Do While (InStr(ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value, "ОО «") = 0) And (InStr(ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value, "Интегральный рейтинг по офисам") = 0)
  
    ' Проверяем параметр с весом
    If (ThisWorkbook.Sheets("Лист8").Cells(rowCount, 3).Value <> "") Or (Показатель_MBO(ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value) = True) Then
      
      ' Прогноз есть в столобце 8 или берем из 20-го
      If ThisWorkbook.Sheets("Лист8").Cells(rowCount, 8).Value <> "" Then
        column_Прогноз = 8
      Else
       
       
      ' Если это Портфель или Пассивы, то Прогноз=Факт
      If (InStr(ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value, "Портфель") <> 0) Or (InStr(ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value, "Пассивы") <> 0) Or (InStr(ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value, "Инвест") <> 0) Then
        ' Факт
        column_Прогноз = 7
      Else
        ' Расчетный прогноз
        column_Прогноз = 20
      End If
       
      End If
      
      ' Проверяем прогноз квартала
      If (ThisWorkbook.Sheets("Лист8").Cells(rowCount, column_Прогноз).Value >= In_LowRange) And ((ThisWorkbook.Sheets("Лист8").Cells(rowCount, column_Прогноз).Value <= In_UpperRange)) Then
      
        If Показатели_Зеленая_Желтая_Красная_зона_Q <> "" Then
          Показатели_Зеленая_Желтая_Красная_зона_Q = Показатели_Зеленая_Желтая_Красная_зона_Q + ", "
        End If
        
        Показатели_Зеленая_Желтая_Красная_зона_Q = Показатели_Зеленая_Желтая_Красная_зона_Q + Сокр2(ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value) + " (" + CStr(Round(ThisWorkbook.Sheets("Лист8").Cells(rowCount, column_Прогноз).Value * 100, 0)) + "%)"
        
        ' Вставляем строку в исходящую таблицу FileTargetWeekName
        Do While Not IsEmpty(Workbooks(In_FileTargetWeekName).Sheets("Лист1").Cells(row_Показатель_Лист1, 1).Value)
          число_строк_Лист1 = число_строк_Лист1 + 1
          row_Показатель_Лист1 = row_Показатель_Лист1 + 1
        Loop
        
        ' Вставляем #1
        ' Номер
        Workbooks(In_FileTargetWeekName).Sheets("Лист1").Cells(row_Показатель_Лист1, 1).Value = число_строк_Лист1
        
        ' Наименование показателя
        ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Copy Destination:=Workbooks(In_FileTargetWeekName).Sheets("Лист1").Cells(row_Показатель_Лист1, 2)
        
        ' Вес
        ThisWorkbook.Sheets("Лист8").Cells(rowCount, 3).Copy Destination:=Workbooks(In_FileTargetWeekName).Sheets("Лист1").Cells(row_Показатель_Лист1, 3)
        
        ' Ед.изм.
        ThisWorkbook.Sheets("Лист8").Cells(rowCount, 4).Copy Destination:=Workbooks(In_FileTargetWeekName).Sheets("Лист1").Cells(row_Показатель_Лист1, 4)
        
        ' План
        ThisWorkbook.Sheets("Лист8").Cells(rowCount, 5).Copy Destination:=Workbooks(In_FileTargetWeekName).Sheets("Лист1").Cells(row_Показатель_Лист1, 5)
        
        ' Факт
        ThisWorkbook.Sheets("Лист8").Cells(rowCount, 6).Copy Destination:=Workbooks(In_FileTargetWeekName).Sheets("Лист1").Cells(row_Показатель_Лист1, 6)
        
        ' Исполнение
        ThisWorkbook.Sheets("Лист8").Cells(rowCount, 7).Copy Destination:=Workbooks(In_FileTargetWeekName).Sheets("Лист1").Cells(row_Показатель_Лист1, 7)
        
        ' Прогноз (если прогноз не берем со  столбца 7 на Лист8)
        If column_Прогноз <> 7 Then
          
          ThisWorkbook.Sheets("Лист8").Cells(rowCount, column_Прогноз).Copy Destination:=Workbooks(In_FileTargetWeekName).Sheets("Лист1").Cells(row_Показатель_Лист1, 8)
                    
        End If
        
        ' Цель недели
        ' Формируем расчет Цель на неделю:
        Дата_начала_недели_по_DB = CDate(Mid(ThisWorkbook.Sheets("Лист8").Range("B5").Value, 52, 10))
        Дата_окончания_недели_по_DB = Дата_начала_недели_по_DB + 7
        ' Дата_окончания_недели_по_DB = Дата_начала_недели_по_DB + 6
          
        ' Если неделя переходит на новый квартал, то берем последний день квартала
        If Дата_окончания_недели_по_DB > Date_last_day_quarter(Дата_начала_недели_по_DB) Then
          Дата_окончания_недели_по_DB = Date_last_day_quarter(Дата_начала_недели_по_DB)
        End If
          
        ' Если показатель перевыполняется, то задача сохранить динамику, иначе стремимся на 100%
        If Workbooks(In_FileTargetWeekName).Sheets("Лист1").Cells(row_Показатель_Лист1, 8).Value > 1 Then
          ' Проверять, чтобы было не пусто! иначе ошибка
          ' Цель_прогноза_квартала = Workbooks(In_FileTargetWeekName).Sheets("Лист1").Cells(row_Показатель_Лист1, 8).Value / 100
        Else
          Цель_прогноза_квартала = 1
        End If
          
        ' Если у параметра есть прогноз в DB
        If Workbooks(In_FileTargetWeekName).Sheets("Лист1").Cells(row_Показатель_Лист1, 8).Value <> "" Then
          
          ' Делаем расчет, чтобы прогноз был 100%                                                                     Тестируемая функция, которая работает! ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 20).Value = Прогноз_квартала_проц(dateDB, ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 5).Value, ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 6).Value, 5, 0)
          Факт_на_дату_для_прогноза_квартала_Var = Факт_на_дату_для_прогноза_квартала(Дата_окончания_недели_по_DB, _
                                                                                        Workbooks(In_FileTargetWeekName).Sheets("Лист1").Cells(row_Показатель_Лист1, 5).Value, _
                                                                                          Цель_прогноза_квартала, _
                                                                                            5, _
                                                                                              0)
          Необходимый_прирост_за_неделю = Факт_на_дату_для_прогноза_квартала_Var - Workbooks(In_FileTargetWeekName).Sheets("Лист1").Cells(row_Показатель_Лист1, 6).Value
        
        Else
          
          ' Если у показателя нет прогноза, то делаем расчет на неделю равномерно на остаток дней
          План_Var = Workbooks(In_FileTargetWeekName).Sheets("Лист1").Cells(row_Показатель_Лист1, 5).Value
          Факт_Var = Workbooks(In_FileTargetWeekName).Sheets("Лист1").Cells(row_Показатель_Лист1, 6).Value
          ' Считаем с сегодня Date по календарю
          Число_рабочих_дней_до_конца_Q = Working_days_between_dates(Date, Date_last_day_quarter(Date), 5)
          Число_рабочих_дней_на_этой_неделе = 5
          Необходимый_прирост_за_неделю = ((План_Var - Факт_Var) / Число_рабочих_дней_до_конца_Q) * Число_рабочих_дней_на_этой_неделе
          
        End If
          
                    
        ' Если есть необходимый расчетный прирост за неделю
        If Необходимый_прирост_за_неделю > 0 Then
          
          ' Копируем Формат параметров ячейки (из План)
          Workbooks(In_FileTargetWeekName).Sheets("Лист1").Cells(row_Показатель_Лист1, 5).Copy Destination:=Workbooks(In_FileTargetWeekName).Sheets("Лист1").Cells(row_Показатель_Лист1, 9)
          
          ' Цели на неделю - прирост
          Workbooks(In_FileTargetWeekName).Sheets("Лист1").Cells(row_Показатель_Лист1, 9) = Необходимый_прирост_за_неделю
          
          ' Если это не Портфель и не Пассивы, то ставим расчетный прирост
          ' If (InStr(ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value, "Портфель") = 0) And (InStr(ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value, "Пассивы") = 0) And (InStr(ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value, "Инвест") = 0) Then
          
          ' Если у параметра есть прогноз в DB
          If Workbooks(In_FileTargetWeekName).Sheets("Лист1").Cells(row_Показатель_Лист1, 8).Value <> "" Then
          
            ' Расчетный прогноз на конец недели
            Workbooks(In_FileTargetWeekName).Sheets("Лист1").Cells(row_Показатель_Лист1, 10).Value = Прогноз_квартала_проц(Дата_окончания_недели_по_DB, _
                                                                                                                     Workbooks(In_FileTargetWeekName).Sheets("Лист1").Cells(row_Показатель_Лист1, 5).Value, _
                                                                                                                       Workbooks(In_FileTargetWeekName).Sheets("Лист1").Cells(row_Показатель_Лист1, 6).Value + Workbooks(In_FileTargetWeekName).Sheets("Лист1").Cells(row_Показатель_Лист1, 9).Value, _
                                                                                                                        5, _
                                                                                                                         0)
            Workbooks(In_FileTargetWeekName).Sheets("Лист1").Cells(row_Показатель_Лист1, 10).NumberFormat = "0%"
            Workbooks(In_FileTargetWeekName).Sheets("Лист1").Cells(row_Показатель_Лист1, 10).HorizontalAlignment = xlRight
                              
          End If ' Если это не Портфель и не Пассивы
        
        End If ' If Необходимый_прирост_за_неделю > 0 Then

        
        ' Итоги прошедшей недели ДД.ММ-ДД.ММ - здесь проверять getDataFrom_BASE_Workbook на <> "не найден", иначе  - ошибка!
        ' If False Then
          
          ' Прирост (вычисляем как разность Факт_Q_Update_Date - Факт_Q)
          getDataFrom_BASE_Workbook_Var1 = CheckData(getDataFrom_BASE_Workbook("TargetWeek", "Лист1", "ID_Rec", ID_Rec_TargetWeek(weekStartDate(Date - 7), Workbooks(In_FileTargetWeekName).Sheets("Лист1").Cells(row_Показатель_Лист1, 2).Value, In_officeNameInReport), "Факт_Q_Update_Date", 0))
          getDataFrom_BASE_Workbook_Var2 = CheckData(getDataFrom_BASE_Workbook("TargetWeek", "Лист1", "ID_Rec", ID_Rec_TargetWeek(weekStartDate(Date - 7), Workbooks(In_FileTargetWeekName).Sheets("Лист1").Cells(row_Показатель_Лист1, 2).Value, In_officeNameInReport), "Факт_Q", 0))
          Workbooks(In_FileTargetWeekName).Sheets("Лист1").Cells(row_Показатель_Лист1, 12).Value = getDataFrom_BASE_Workbook_Var1 - getDataFrom_BASE_Workbook_Var2
          Workbooks(In_FileTargetWeekName).Sheets("Лист1").Cells(row_Показатель_Лист1, 12).NumberFormat = "#,##0"
          Workbooks(In_FileTargetWeekName).Sheets("Лист1").Cells(row_Показатель_Лист1, 12).HorizontalAlignment = xlRight
              
        
          ' Исполнение цели - Если была установлена цель по данному продукту на прошлой неделе, то выводим Исполнение цели
          If getDataFrom_BASE_Workbook("TargetWeek", "Лист1", "ID_Rec", ID_Rec_TargetWeek(weekStartDate(Date - 7), Workbooks(In_FileTargetWeekName).Sheets("Лист1").Cells(row_Показатель_Лист1, 2).Value, In_officeNameInReport), "TargetWeek", 0) > 0 Then
          
            Workbooks(In_FileTargetWeekName).Sheets("Лист1").Cells(row_Показатель_Лист1, 13).Value = CheckData(getDataFrom_BASE_Workbook("TargetWeek", "Лист1", "ID_Rec", ID_Rec_TargetWeek(weekStartDate(Date - 7), Workbooks(In_FileTargetWeekName).Sheets("Лист1").Cells(row_Показатель_Лист1, 2).Value, In_officeNameInReport), "Исп_TargetWeek", 0))
            Workbooks(In_FileTargetWeekName).Sheets("Лист1").Cells(row_Показатель_Лист1, 13).NumberFormat = "0%"
            Workbooks(In_FileTargetWeekName).Sheets("Лист1").Cells(row_Показатель_Лист1, 13).HorizontalAlignment = xlRight
            ' Окраска ячейки СФЕТОФОР ' 70. Заливка ячейки цветом "светофор" - значение In_Value в %, In_Target в %
            Call Full_Color_RangeV(In_FileTargetWeekName, "Лист1", row_Показатель_Лист1, 13, Workbooks(In_FileTargetWeekName).Sheets("Лист1").Cells(row_Показатель_Лист1, 13).Value, 1)

          End If
        ' End If
        
        ' Форматируем строку на Лист1 - убираем все линии в ячейках
        For столбцы_на_Лист1 = 1 To 10
          Call Формат_ячейки_Цели_на_неделю(In_FileTargetWeekName, "Лист1", row_Показатель_Лист1, столбцы_на_Лист1)
        Next столбцы_на_Лист1

        ' Вставляем #2
        ' ID_Rec = 1-2321-ЗП Офис-НеделяГод-Продукт
        WWYY_Var = CStr(WeekNumber(Дата_окончания_недели_по_DB)) + strYY(Дата_окончания_недели_по_DB)
        Product_Code_Var = Product_Name_to_Product_Code(Workbooks(In_FileTargetWeekName).Sheets("Лист1").Cells(row_Показатель_Лист1, 2).Value)
        ID_Rec_TargetWeek_Var = CStr(getNumberOfficeByName2(In_officeNameInReport)) + "-" + WWYY_Var + "-" + Product_Code_Var
        
        ' Вносим данные в BASE\TargetWeek
        Call InsertRecordInBook("TargetWeek", "Лист1", "ID_Rec", ID_Rec_TargetWeek_Var, _
                                            "ID_Rec", ID_Rec_TargetWeek_Var, _
                                              "Оffice_Number", getNumberOfficeByName2(In_officeNameInReport), _
                                                "Оffice", getShortNameOfficeByName(In_officeNameInReport), _
                                                  "Product_Name", Workbooks(In_FileTargetWeekName).Sheets("Лист1").Cells(row_Показатель_Лист1, 2).Value, _
                                                    "Product_Code", Product_Code_Var, _
                                                      "Unit", Workbooks(In_FileTargetWeekName).Sheets("Лист1").Cells(row_Показатель_Лист1, 4).Value, _
                                                        "WWYY", WWYY_Var, _
                                                          "Update_Date", Дата_начала_недели_по_DB, _
                                                            "DB_Start", Дата_начала_недели_по_DB, _
                                                              "DB_End", Дата_окончания_недели_по_DB, _
                                                                "План_Q", Workbooks(In_FileTargetWeekName).Sheets("Лист1").Cells(row_Показатель_Лист1, 5).Value, _
                                                                  "Факт_Q", Workbooks(In_FileTargetWeekName).Sheets("Лист1").Cells(row_Показатель_Лист1, 6).Value, _
                                                                    "Исп_Q", Workbooks(In_FileTargetWeekName).Sheets("Лист1").Cells(row_Показатель_Лист1, 7).Value, _
                                                                      "Прог_Q", Workbooks(In_FileTargetWeekName).Sheets("Лист1").Cells(row_Показатель_Лист1, 8).Value, _
                                                                        "TargetWeek", Необходимый_прирост_за_неделю, _
                                                                          "TargetWeek_Прог_Q", Workbooks(In_FileTargetWeekName).Sheets("Лист1").Cells(row_Показатель_Лист1, 10).Value, _
                                                                            "Факт_Q_Update_Date", Workbooks(In_FileTargetWeekName).Sheets("Лист1").Cells(row_Показатель_Лист1, 6).Value, _
                                                                              "Исп_TargetWeek", РассчетДоли(Необходимый_прирост_за_неделю, Workbooks(In_FileTargetWeekName).Sheets("Лист1").Cells(row_Показатель_Лист1, 6).Value - Workbooks(In_FileTargetWeekName).Sheets("Лист1").Cells(row_Показатель_Лист1, 6).Value, 3), _
                                                                                "Fact_WeekStart", weekStartDate(Дата_окончания_недели_по_DB), _
                                                                                  "Fact_WeekEnd", weekEndDate(Дата_окончания_недели_по_DB))
      
      End If ' Проверяем прогноз квартала
      
    End If
    
    ' Следующая запись
    rowCount = rowCount + 1
  Loop

      
End Function

' Генерация строки
Function Промежуточный_срез_результатов() As String

  If Cегодня_четное_число(Now) = True Then
    Промежуточный_срез_результатов = "Промежуточный срез результатов"
  Else
    Промежуточный_срез_результатов = "Промежуточные итоги"
  End If
  
  
  ' Промежуточный_срез_результатов = "Промежуточные итоги"
  ' Промежуточный_срез_результатов = "Результаты"
  
  
  
End Function

' Четная или нечетная дата в числе клендаря
Function Cегодня_четное_число(In_Date) As Boolean
  
  ' Если сегодня четное число
  If CInt(Mid(CStr(In_Date), 1, 2)) Mod 2 = 0 Then
    Cегодня_четное_число = True
  Else
    Cегодня_четное_число = False
  End If
  
  
End Function

' Квартал сколько дней прошло и сколько всего (для расчета прогноза)
Function Квартал_днейпрошло_днейвсего() As String
  
  dateDB_Лист8 = CDate(Mid(ThisWorkbook.Sheets("Лист8").Range("B5").Value, 52, 10))
  
  Дата_начала_квартала = quarterStartDate(dateDB_Лист8)
  
  Дата_конца_квартала = Date_last_day_quarter(dateDB_Лист8)
  
  Дней_прошло = Working_days_between_dates(Дата_начала_квартала, dateDB_Лист8, 5)
  
  Дней_всего = Working_days_between_dates(Дата_начала_квартала, Дата_конца_квартала, 5)
  
  Квартал_днейвсего_прошлодней = "Дней_прошло: " + CStr(Дней_прошло) + ", Дней_всего: " + CStr(Дней_всего)
  
  ThisWorkbook.Sheets("Лист8").Range("T8").Value = Квартал_днейвсего_прошлодней
  
End Function

' Обновление исполнения плана по целям по свежему DB
Sub Отклонения_по_офисам_Update()
Dim дата_для_расчета_недели As Date
  
  ' Строка статуса
  Application.StatusBar = "Отклонения_по_офисам_Update..."

  ' Открываем BASE\Sales
  OpenBookInBase ("MBO")
  ThisWorkbook.Sheets("Лист8").Activate

  OpenBookInBase ("TargetWeek")
  ThisWorkbook.Sheets("Лист8").Activate

  OpenBookInBase ("Products")
  ThisWorkbook.Sheets("Лист8").Activate

  ' Определяем столбцы в "TargetWeek"
  column_Update_Date = 8
  column_Факт_Q_Update_Date = 17
  column_Факт_Исп_TargetWeek = 18

  ' Дата DB
  dateDB_Лист8 = CDate(Mid(ThisWorkbook.Sheets("Лист8").Range("B5").Value, 52, 10))

  ' Определяем дату для рсчета недели в TargetWeek
  Select Case Weekday(dateDB_Лист8, vbMonday)
    Case 1 ' Понедельник
      дата_для_расчета_недели = dateDB_Лист8
    Case 2 ' Вторник
      дата_для_расчета_недели = dateDB_Лист8
    Case 3 ' Среда
      дата_для_расчета_недели = dateDB_Лист8
    Case 4 ' Четверг
      ' дата_для_расчета_недели = dateDB_Лист8 + 4
      ' c 12.07 пробуем оставить результат четверга на прошлой неделе. По ДБ это с 08.07
      дата_для_расчета_недели = dateDB_Лист8
    Case 5 ' Пятница
      дата_для_расчета_недели = dateDB_Лист8 + 3
    Case 6 ' Суббота
      дата_для_расчета_недели = dateDB_Лист8 + 2
    Case 7 ' Воскресенье
      дата_для_расчета_недели = dateDB_Лист8 + 1
  End Select
  
  ' Определяем неделю в TargetWeek
  WWYY_Var = CStr(WeekNumber(дата_для_расчета_недели)) + strYY(дата_для_расчета_недели)

  ' Проходим по 5-ти офисам
  ' Заголовки
  For i = 1 To 5 ' 5 - на период отладки
    ' Номера офисов от 1 до 5
    Select Case i
      Case 1 ' ОО «Тюменский»
        officeNameInReport = "ОО «Тюменский»"
      Case 2 ' ОО «Сургутский»
        officeNameInReport = "ОО «Сургутский»"
      Case 3 ' ОО «Нижневартовский»
        officeNameInReport = "ОО «Нижневартовский»"
      Case 4 ' ОО «Новоуренгойский»
        officeNameInReport = "ОО «Новоуренгойский»"
      Case 5 ' ОО «Тарко-Сале»
        officeNameInReport = "ОО «Тарко-Сале»"
    End Select
  
    ' Берем с листа ОО «Тюменский»
    rowCount = rowByValue(ThisWorkbook.Name, "Лист8", officeNameInReport, 1000, 100) + 3
  
    ' Берем с Лист1
    ' row_Показатель_Лист1 = rowByValue(In_FileTargetWeekName, "Лист1", "Показатель", 100, 100) + 2
    ' число_строк_Лист1 = 1
  
    ' Обрабатываем Лист - ищем Сначала Офис, если находим офис, то ищем позицию с наименованием продукта
    Do While (InStr(ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value, "ОО «") = 0) And (InStr(ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value, "Интегральный рейтинг по офисам") = 0)
  
      ' Проверяем параметр с весом или целью MBO
      If (ThisWorkbook.Sheets("Лист8").Cells(rowCount, 3).Value <> "") Or (Показатель_MBO(ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value) = True) Then
        
        
        Product_Code_Var = Product_Name_to_Product_Code(ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value)
        
        ID_Rec_TargetWeek_Var = CStr(getNumberOfficeByName2(officeNameInReport)) + "-" + WWYY_Var + "-" + Product_Code_Var
        
        ' Переходим в "TargetWeek"
        ' Выполняем поиск - Текущая_дата_рассчета есть в BASE\NonWorkingDays?
        Set searchResults = Workbooks("TargetWeek").Sheets("Лист1").Columns("A:A").Find(ID_Rec_TargetWeek_Var, LookAt:=xlWhole)
  
        ' Проверяем - есть ли такая запись в "TargetWeek"
        If searchResults Is Nothing Then
          ' Если не найдена
        Else
          
          ' Если найдена
          ' Апдейтим найденную запись: дату дашборда Update_Date, факт по продукту Факт_Q_Update_Date и считаем Исп_TargetWeek
          ' Update_Date
          Workbooks("TargetWeek").Sheets("Лист1").Cells(searchResults.Row, column_Update_Date).Value = dateDB_Лист8
          Workbooks("TargetWeek").Sheets("Лист1").Cells(searchResults.Row, column_Факт_Q_Update_Date).Value = ThisWorkbook.Sheets("Лист8").Cells(rowCount, 6).Value
          Workbooks("TargetWeek").Sheets("Лист1").Cells(searchResults.Row, column_Факт_Исп_TargetWeek).Value = РассчетДоли(Workbooks("TargetWeek").Sheets("Лист1").Cells(searchResults.Row, 15).Value, _
                                                                                                                             Workbooks("TargetWeek").Sheets("Лист1").Cells(searchResults.Row, 17).Value - Workbooks("TargetWeek").Sheets("Лист1").Cells(searchResults.Row, 12).Value, 3)
        End If
        
      End If ' Проверяем параметр с весом или целью MBO
    
      ' Следующая запись
      rowCount = rowCount + 1
    Loop
  
  Next i


  ' Закрываем таблицу MBO
  ' Закрываем BASE\Sales
  CloseBook ("MBO")
  ThisWorkbook.Sheets("Лист8").Activate

  ' Закрываем таблицу MBO
  ' Закрываем BASE\Sales
  CloseBook ("TargetWeek")
  ThisWorkbook.Sheets("Лист8").Activate
    
  ' Закрываем таблицу MBO
  ' Закрываем BASE\Sales
  CloseBook ("Products")
  ThisWorkbook.Sheets("Лист8").Activate

  
  ' Строка статуса
  Application.StatusBar = ""


End Sub

' Форматирование ячейки
Sub Формат_ячейки_Цели_на_неделю(In_Workbooks, In_Sheets, In_Row, In_Col)
  
  Workbooks(In_Workbooks).Sheets(In_Sheets).Cells(In_Row, In_Col).Font.Bold = False
  Workbooks(In_Workbooks).Sheets(In_Sheets).Cells(In_Row, In_Col).Borders(xlDiagonalDown).LineStyle = xlNone
  Workbooks(In_Workbooks).Sheets(In_Sheets).Cells(In_Row, In_Col).Borders(xlDiagonalUp).LineStyle = xlNone
  Workbooks(In_Workbooks).Sheets(In_Sheets).Cells(In_Row, In_Col).Borders(xlEdgeLeft).LineStyle = xlNone
  Workbooks(In_Workbooks).Sheets(In_Sheets).Cells(In_Row, In_Col).Borders(xlEdgeTop).LineStyle = xlNone
  Workbooks(In_Workbooks).Sheets(In_Sheets).Cells(In_Row, In_Col).Borders(xlEdgeBottom).LineStyle = xlNone
  Workbooks(In_Workbooks).Sheets(In_Sheets).Cells(In_Row, In_Col).Borders(xlEdgeRight).LineStyle = xlNone
  Workbooks(In_Workbooks).Sheets(In_Sheets).Cells(In_Row, In_Col).Borders(xlInsideVertical).LineStyle = xlNone
  Workbooks(In_Workbooks).Sheets(In_Sheets).Cells(In_Row, In_Col).Borders(xlInsideHorizontal).LineStyle = xlNone

End Sub

' Тест даты в номер недели
Sub дата_в_неделю()
  
  ThisWorkbook.Sheets("Лист8").Range("X6").Value = Номер_недели(CDate(ThisWorkbook.Sheets("Лист8").Range("V6").Value))

End Sub

' Генерация ID_Rec для BASE\TargetWeek по Дате недели и наименованию продукта In_Product_Name и офису
Function ID_Rec_TargetWeek(In_Date, In_Product_Name, In_officeNameInReport) As String
  
  ' ID_Rec = 1-2321-ЗП Офис-НеделяГод-Продукт
  WWYY_Var = CStr(WeekNumber(In_Date)) + strYY(In_Date)
  
  Product_Code_Var = Product_Name_to_Product_Code(In_Product_Name)
  
  ID_Rec_TargetWeek = CStr(getNumberOfficeByName2(In_officeNameInReport)) + "-" + WWYY_Var + "-" + Product_Code_Var

End Function

' Апдейтим таблицу BASE\Products
Sub Update_BASE_Products(In_Product_Name, In_Product_Code, In_Unit)
  
  ' Ищем Product_Code
  
  ' Выполняем поиск - Текущая_дата_рассчета есть в BASE\NonWorkingDays?
  Set searchResults = Workbooks("Products").Sheets("Лист1").Columns("B:B").Find(In_Product_Code, LookAt:=xlWhole)
  
  ' Проверяем - есть ли такая дата, если нет, то добавляем
  If searchResults Is Nothing Then
    
    ' Если не найдена - вставляем
    ' Вносим данные в BASE\Products
    Call InsertRecordInBook("Products", "Лист1", "Product_Code", In_Product_Code, _
                                            "Product_Code", In_Product_Code, _
                                              "Product_Name", In_Product_Name, _
                                                "Unit", In_Unit, _
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
                                                                          "", "", _
                                                                            "", "", _
                                                                              "", "", _
                                                                                "", "", _
                                                                                  "", "")
  
  Else
    ' Если найдена
  
  End If

  
End Sub


' Новая версия - использует новую функцию "Факт_Q_на_дату"
Sub Загрузить_факт_на_дату2()
  
  ' Написать новый загрузчик Факта и из прогноза квартала
  
  ' Факт_Q_на_дату - загружает по продукту факт на любую дату
  
  ' Прогноз_квартала_проц - делает расчет прогноза на дату по факту и плану
  
  ' Product_Name_to_Product_Code - преобразует "Потребительские кредиты" в "ПК"
  
  ' Дата, на которую загружаем данные
  dateForLoad = CDate(ThisWorkbook.Sheets("Лист8").Range("O9").Value)
  
  ' Открываем таблицы, которые нужны:
  ' Открываем BASE\Products
  OpenBookInBase ("Products")

  
  ' 1) Обработка офисов
  For i = 1 To 5
    
    ' Номера офисов от 1 до 5
    Select Case i
      Case 1 ' ОО «Тюменский»
        officeNameInReport = "ОО «Тюменский»"
      Case 2 ' ОО «Сургутский»
        officeNameInReport = "ОО «Сургутский»"
      Case 3 ' ОО «Нижневартовский»
        officeNameInReport = "ОО «Нижневартовский»"
      Case 4 ' ОО «Новоуренгойский»
        officeNameInReport = "ОО «Новоуренгойский»"
      Case 5 ' ОО «Тарко-Сале»
        officeNameInReport = "ОО «Тарко-Сале»"
            
    End Select
        
    ' Находим номер строки с наименованием офиса
    row_офис = getRowFromSheet8(officeNameInReport, officeNameInReport)
        
    ' Обрабатываем блок офиса
    rowCount = row_офис + 3
    Do While (ThisWorkbook.Sheets("Лист8").Cells(rowCount, 1).Value <> "") And (ThisWorkbook.Sheets("Лист8").Cells(rowCount, 1).Value <> ".")

      ' Код продукта
      Product_Name_to_Product_Code_Var = Product_Name_to_Product_Code(ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value)
      
      ' Если по продукту найден короткий Код продукта
      If Product_Name_to_Product_Code_Var <> "" Then
      
        ' 1) Заполняем столбец "O" - факт на дату
        ThisWorkbook.Sheets("Лист8").Cells(rowCount, 15).Value = Факт_Q_на_дату(i, _
                                                                                  Product_Name_to_Product_Code_Var, _
                                                                                    dateForLoad)
        ThisWorkbook.Sheets("Лист8").Cells(rowCount, 15).NumberFormat = "#,##0"
        ThisWorkbook.Sheets("Лист8").Cells(rowCount, 15).HorizontalAlignment = xlRight
        
        ' 2) Заполняем столбец "P" - изменения фактов Квартала
        Факт_текущий = 0
        Факт_прошлый = 0
        ' Если в столбце 6 есть значение, то берем его
        If ThisWorkbook.Sheets("Лист8").Cells(rowCount, 6).Value <> "" Then
          Факт_текущий = ThisWorkbook.Sheets("Лист8").Cells(rowCount, 6).Value
        Else
          ' Если в Квартале было "", то Проверяем месяц
          If ThisWorkbook.Sheets("Лист8").Cells(rowCount, 10).Value <> "" Then
            Факт_текущий = ThisWorkbook.Sheets("Лист8").Cells(rowCount, 10).Value
          End If
        End If
        ' Если в столбце 16 есть значение, то берем его
        If ThisWorkbook.Sheets("Лист8").Cells(rowCount, 15).Value <> "" Then
          Факт_прошлый = ThisWorkbook.Sheets("Лист8").Cells(rowCount, 15).Value
        End If
        ' Записываем значения
        ThisWorkbook.Sheets("Лист8").Cells(rowCount, 16).Value = Факт_текущий - Факт_прошлый
        ThisWorkbook.Sheets("Лист8").Cells(rowCount, 16).NumberFormat = "#,##0"
        ThisWorkbook.Sheets("Лист8").Cells(rowCount, 16).HorizontalAlignment = xlRight
        
        ' 3) Заполняем столбец "Q" - прогноз Квартала на дату
        If (ThisWorkbook.Sheets("Лист8").Cells(rowCount, 5).Value <> "") And (ThisWorkbook.Sheets("Лист8").Cells(rowCount, 15).Value <> "") And (ThisWorkbook.Sheets("Лист8").Cells(rowCount, 4).Value <> "%") Then
                    
          ' Если по продукту нет прогноза в столбце 8
          If ThisWorkbook.Sheets("Лист8").Cells(rowCount, 8).Value <> "" Then
            
            ThisWorkbook.Sheets("Лист8").Cells(rowCount, 17).Value = Прогноз_квартала_проц(dateForLoad, _
                                                                                             ThisWorkbook.Sheets("Лист8").Cells(rowCount, 5).Value, _
                                                                                               ThisWorkbook.Sheets("Лист8").Cells(rowCount, 15).Value, _
                                                                                                 5, 0)
                                                                                                 
            ThisWorkbook.Sheets("Лист8").Cells(rowCount, 17).NumberFormat = "0%"
            ThisWorkbook.Sheets("Лист8").Cells(rowCount, 17).HorizontalAlignment = xlRight
          
          End If
                                                                                               
        End If
        
        ' 4) Заполняем столбец "R" - Динамика прогноза
        If (ThisWorkbook.Sheets("Лист8").Cells(rowCount, 8).Value <> "") And (ThisWorkbook.Sheets("Лист8").Cells(rowCount, 17).Value <> "") Then
          
          ThisWorkbook.Sheets("Лист8").Cells(rowCount, 18).Value = ThisWorkbook.Sheets("Лист8").Cells(rowCount, 8).Value - ThisWorkbook.Sheets("Лист8").Cells(rowCount, 17).Value
          ThisWorkbook.Sheets("Лист8").Cells(rowCount, 18).NumberFormat = "0%"
          ThisWorkbook.Sheets("Лист8").Cells(rowCount, 18).HorizontalAlignment = xlRight
        
          ' Окраска ячейки СФЕТОФОР: в красный, если отрицательная динамика и исполнее менее 1
          If (ThisWorkbook.Sheets("Лист8").Cells(rowCount, 18).Value < 0) And (ThisWorkbook.Sheets("Лист8").Cells(rowCount, 11).Value < 1) Then
            Call Full_Color_RangeII("Лист8", rowCount, 18, 0, 100)
          End If

          ' Окраска ячейки СФЕТОФОР: в зеленый, если положительная динамика
          If ThisWorkbook.Sheets("Лист8").Cells(rowCount, 18).Value > 0 Then
            Call Full_Color_RangeII("Лист8", rowCount, 18, 100, 100)
          End If
        End If

      End If
      
      ' Следующая запись
      Application.StatusBar = officeNameInReport + ": " + CStr(ThisWorkbook.Sheets("Лист8").Cells(rowCount, 1).Value) + "..."
      rowCount = rowCount + 1
      DoEventsInterval (rowCount)
    
    Loop
        
  Next i
  
  ' 2) Свод по РОО
  Application.StatusBar = "РОО..."
  
  ' Находим номер строки с наименованием офиса
  row_офис = getRowFromSheet8("Итого по РОО «Тюменский»", "Итого по РОО «Тюменский»")
        
  ' Обрабатываем блок офиса
  rowCount = row_офис + 3
  Do While (ThisWorkbook.Sheets("Лист8").Cells(rowCount, 1).Value <> "")
  
    ' 1) Заполняем столбец "O" - факт на дату
    ThisWorkbook.Sheets("Лист8").Cells(rowCount, 15).Value = ThisWorkbook.Sheets("Лист8").Cells(getRowFromSheet8("ОО «Тюменский»", ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value), 15).Value + _
                                                                 ThisWorkbook.Sheets("Лист8").Cells(getRowFromSheet8("ОО «Сургутский»", ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value), 15).Value + _
                                                                   ThisWorkbook.Sheets("Лист8").Cells(getRowFromSheet8("ОО «Нижневартовский»", ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value), 15).Value + _
                                                                     ThisWorkbook.Sheets("Лист8").Cells(getRowFromSheet8("ОО «Новоуренгойский»", ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value), 15).Value + _
                                                                       ThisWorkbook.Sheets("Лист8").Cells(getRowFromSheet8("ОО «Тарко-Сале»", ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value), 15).Value
    ' Если это %, то делим на 5 офисов - находим ср.арифметическое
    If ThisWorkbook.Sheets("Лист8").Cells(rowCount, 4).Value = "%" Then
      ' Если это %
      ThisWorkbook.Sheets("Лист8").Cells(rowCount, 15).Value = ThisWorkbook.Sheets("Лист8").Cells(rowCount, 15).Value / 5
    End If
    
    ThisWorkbook.Sheets("Лист8").Cells(rowCount, 15).NumberFormat = "#,##0"
    ThisWorkbook.Sheets("Лист8").Cells(rowCount, 15).HorizontalAlignment = xlRight
        
    ' 2) Заполняем столбец "P" - изменения фактов Квартала
    Факт_текущий = 0
    Факт_прошлый = 0
    ' Если в столбце 6 есть значение, то берем его
    If ThisWorkbook.Sheets("Лист8").Cells(rowCount, 6).Value <> "" Then
      Факт_текущий = ThisWorkbook.Sheets("Лист8").Cells(rowCount, 6).Value
    Else
      ' Если в Квартале было "", то Проверяем месяц
      If ThisWorkbook.Sheets("Лист8").Cells(rowCount, 10).Value <> "" Then
        Факт_текущий = ThisWorkbook.Sheets("Лист8").Cells(rowCount, 10).Value
      End If
    End If
    ' Если в столбце 16 есть значение, то берем его
    If ThisWorkbook.Sheets("Лист8").Cells(rowCount, 15).Value <> "" Then
      Факт_прошлый = ThisWorkbook.Sheets("Лист8").Cells(rowCount, 15).Value
    End If
    ' Записываем значения
    ThisWorkbook.Sheets("Лист8").Cells(rowCount, 16).Value = Факт_текущий - Факт_прошлый
    ThisWorkbook.Sheets("Лист8").Cells(rowCount, 16).NumberFormat = "#,##0"
    ThisWorkbook.Sheets("Лист8").Cells(rowCount, 16).HorizontalAlignment = xlRight

    ' 3) Заполняем столбец "Q" - прогноз Квартала на дату
    If (ThisWorkbook.Sheets("Лист8").Cells(rowCount, 5).Value <> "") And (ThisWorkbook.Sheets("Лист8").Cells(rowCount, 15).Value <> "") And (ThisWorkbook.Sheets("Лист8").Cells(rowCount, 4).Value <> "%") Then
      
      ' Если по продукту нет прогноза в столбце 8
      If ThisWorkbook.Sheets("Лист8").Cells(rowCount, 8).Value <> "" Then
      
        ThisWorkbook.Sheets("Лист8").Cells(rowCount, 17).Value = Прогноз_квартала_проц(dateForLoad, _
                                                                                             ThisWorkbook.Sheets("Лист8").Cells(rowCount, 5).Value, _
                                                                                               ThisWorkbook.Sheets("Лист8").Cells(rowCount, 15).Value, _
                                                                                                 5, 0)
        ThisWorkbook.Sheets("Лист8").Cells(rowCount, 17).NumberFormat = "0%"
        ThisWorkbook.Sheets("Лист8").Cells(rowCount, 17).HorizontalAlignment = xlRight
      End If
                                                                                               
    End If
        
    ' 4) Заполняем столбец "R" - Динамика прогноза
    If (ThisWorkbook.Sheets("Лист8").Cells(rowCount, 8).Value <> "") And (ThisWorkbook.Sheets("Лист8").Cells(rowCount, 17).Value <> "") Then
          
      ThisWorkbook.Sheets("Лист8").Cells(rowCount, 18).Value = ThisWorkbook.Sheets("Лист8").Cells(rowCount, 8).Value - ThisWorkbook.Sheets("Лист8").Cells(rowCount, 17).Value
      ThisWorkbook.Sheets("Лист8").Cells(rowCount, 18).NumberFormat = "0%"
      ThisWorkbook.Sheets("Лист8").Cells(rowCount, 18).HorizontalAlignment = xlRight
        
      ' Окраска ячейки СФЕТОФОР: в красный, если отрицательная динамика и исполнее менее 1
      If (ThisWorkbook.Sheets("Лист8").Cells(rowCount, 18).Value < 0) And (ThisWorkbook.Sheets("Лист8").Cells(rowCount, 11).Value < 1) Then
        Call Full_Color_RangeII("Лист8", rowCount, 18, 0, 100)
      End If

      ' Окраска ячейки СФЕТОФОР: в зеленый, если положительная динамика
      If ThisWorkbook.Sheets("Лист8").Cells(rowCount, 18).Value > 0 Then
        Call Full_Color_RangeII("Лист8", rowCount, 18, 100, 100)
      End If
    End If

  
    ' Следующая запись
    Application.StatusBar = "РОО: " + CStr(ThisWorkbook.Sheets("Лист8").Cells(rowCount, 1).Value) + "..."
    rowCount = rowCount + 1
    DoEventsInterval (rowCount)
    
  Loop
  
    
  ' Завершение
  Application.StatusBar = "Завершение..."
  
  ' Закрываем таблицы, использовавшиеся в функциях
  ' Закрываем BASE\Products
  CloseBook ("Products")
  
  Application.StatusBar = ""
  
  
End Sub

' Выставляем - Цель "На неделю:" в "M9"
Sub Цель_на_неделю_Лист8()
      
      ' Проверка - первый день недели
      Первый_день_недели = False
      
      ' Если сегодня понедельник
      If Weekday(Date, vbMonday) = 1 Then
        Первый_день_недели = True
      End If
      
      ' Если на Лист0 выставлен первый день недели
      If CStr(ThisWorkbook.Sheets("Лист0").Cells(rowByValue(ThisWorkbook.Name, "Лист0", "Первый день недели:", 100, 100), ColumnByValue(ThisWorkbook.Name, "Лист0", "Первый день недели:", 100, 100) + 2).Value) = "1" Then
        ' Обработать как первый день недели?
        If MsgBox("Сформировать задачи первого дня недели?", vbYesNo) = vbYes Then
          Первый_день_недели = True
        Else
          Первый_день_недели = False
        End If
      End If
      
      ' Заголовок
      If Первый_день_недели = True Then
        ' ThisWorkbook.Sheets("Лист8").Range("M9").Value = "На неделю:       (до " + CStr(dateDB + 7) + ")"
        '
        ThisWorkbook.Sheets("Лист8").Range("M9").Value = "План на неделю: (" + strDDMM(Date) + "-" + strДД_MM_YY2(Date + 6) + ")"
      End If

End Sub

' Чек-лист Офисные продажи
Sub Чек_лист_Офисные_продажи()
  
  
  
End Sub

' Оперативная справка за неделю
Sub Оперативная_справка_за_неделю()
  
  ' Строка статуса
  Application.StatusBar = "Оперативная справка за неделю..."
  
  ' Дата DB
  dateDB_Лист8 = CDate(Mid(ThisWorkbook.Sheets("Лист8").Range("B5").Value, 52, 10))
 
  ' Открываем таблицы, которые нужны:
  ' Открываем BASE\Products
  OpenBookInBase ("Products")
  
  ' Шаблон "Оперативная справка за неделю"
  ' Открываем шаблон
  If Dir(ThisWorkbook.Path + "\Templates\" + "Оперативная справка за неделю.xlsx") <> "" Then
    ' Открываем шаблон Templates\Оперативная справка за неделю
    TemplatesFileName = "Оперативная справка за неделю"
  End If
              
  ' Открываем шаблон Протокола из C:\Users\...\Documents\#VBA\DB_Result\Templates
  Workbooks.Open (ThisWorkbook.Path + "\Templates\" + TemplatesFileName + ".xlsx")
           
  ' Переходим на окно DB
  ThisWorkbook.Sheets("Лист8").Activate

  ' Имя нового файла
  FileDBName_OSp = "Оперативная справка за неделю" + Replace(Mid(ThisWorkbook.Sheets("Лист8").Range("B5").Value, 48, 14), ".", "-") + ".xlsx"
  
  ' Проверяем - если файл есть, то удаляем его
  Call deleteFile(ThisWorkbook.Path + "\Out\" + FileDBName_OSp)
  
  Workbooks(TemplatesFileName).SaveAs FileName:=ThisWorkbook.Path + "\Out\" + FileDBName_OSp, FileFormat:=xlOpenXMLWorkbook, createBackUp:=False
  ' Записываем в S3 на Лист8
  ThisWorkbook.Sheets("Лист8").Range("S3").Value = ThisWorkbook.Path + "\Out\" + FileDBName_OSp
  
  ' Строка статуса
  Application.StatusBar = "Оперативная справка за неделю: определение столбцов..."
    
  ' Находим столбец #офис
  column_офис = ColumnByValue(FileDBName_OSp, "Лист1", "#офис", 100, 100)
  
  ' Находим столбец #Product_Code
  column_Product_Code = ColumnByValue(FileDBName_OSp, "Лист1", "#Product_Code", 100, 100)
  
  ' Находим столбец #прогноз
  column_рассчетный_прогноз = ColumnByValue(FileDBName_OSp, "Лист1", "#прогноз", 100, 100)

  ' Определение столбцов
  ' Продукт
  Column_Продукт = ColumnByValue(FileDBName_OSp, "Лист1", "Продукт", 100, 100)
  
  ' Ед.изм.
  column_Ед_изм = ColumnByValue(FileDBName_OSp, "Лист1", "Ед.изм.", 100, 100)
  
  ' План
  column_План = ColumnByValue(FileDBName_OSp, "Лист1", "План", 100, 100)
  
  ' Факт
  column_Факт = ColumnByValue(FileDBName_OSp, "Лист1", "Факт на ____", 100, 100)
  
  ' Исп.
  column_Исп = ColumnByValue(FileDBName_OSp, "Лист1", "Исп.", 100, 100)
  
  ' Прогноз
  column_Прогноз = ColumnByValue(FileDBName_OSp, "Лист1", "Прогноз", 100, 100)
  
  ' Прирост за прошлую неделю
  ' Факт
  column_Факт_прошлая_неделя = ColumnByValue(FileDBName_OSp, "Лист1", "Факт _", 100, 100)

  ' Изм.
  column_Изм = ColumnByValue(FileDBName_OSp, "Лист1", "Изм.", 100, 100)
  
  ' Прогн.Q
  column_Прогн_Q = ColumnByValue(FileDBName_OSp, "Лист1", "Прогн.Q", 100, 100)
  
  ' Динамика %
  column_Динамика = ColumnByValue(FileDBName_OSp, "Лист1", "%", 100, 100)
  
  ' Поручение
  ' Дата исполнения  ____
  column_Поручение = ColumnByValue(FileDBName_OSp, "Лист1", "Дата исполнения  ____", 100, 100)
  
  ' Находим строку "Подразделение"
  row_Подразделение = rowByValue(FileDBName_OSp, "Лист1", "Подразделение", 100, 100)
  
  ' Строка статуса
  Application.StatusBar = "Оперативная справка за неделю: обработка..."
  
  ' Прописываем заголовки в шапке таблице:
  ' F4 - "__ кв. 2020 г." по dateDB
  Workbooks(FileDBName_OSp).Sheets("Лист1").Range("F4").Value = quarterName(dateDB_Лист8)
  ' G5 - "Факт на ____"
  Workbooks(FileDBName_OSp).Sheets("Лист1").Range("G5").Value = "Факт на " + strDDMM(dateDB_Лист8)
  ' K4 - "Прирост за прошлую неделю"
  Workbooks(FileDBName_OSp).Sheets("Лист1").Range("K4").Value = "Прирост за прошлую неделю " + strDDMM(dateDB_Лист8 - 7) + "-" + strDDMM(dateDB_Лист8)
  ' K5 - "Факт"
  Workbooks(FileDBName_OSp).Sheets("Лист1").Range("K5").Value = "Факт " + strDDMM(dateDB_Лист8 - 7)
  ' P5 - "Дата исполнения  ____"
  Workbooks(FileDBName_OSp).Sheets("Лист1").Range("P5").Value = "Дата исполнения " + strDDMM(dateDB_Лист8 + 7)
  
  ' Идем по таблице сверху вниз и определяем в #офис и #Product_Code
  rowCount = row_Подразделение + 1
  Do While Workbooks(FileDBName_OSp).Sheets("Лист1").Cells(rowCount, Column_Продукт).Value <> ""
    
      
    ' Продукт
    ' column_Продукт = ColumnByValue(FileDBName_OSp, "Лист1", "Продукт", 100, 100)
  
    ' Ед.изм.
    ' column_Ед_изм = ColumnByValue(FileDBName_OSp, "Лист1", "Ед.изм.", 100, 100)
    
    Row_Лист8 = getRowFromSheet8(getNameOfficeByNumber(Workbooks(FileDBName_OSp).Sheets("Лист1").Cells(rowCount, column_офис).Value), Product_Code_to_Product_Name(Workbooks(FileDBName_OSp).Sheets("Лист1").Cells(rowCount, column_Product_Code).Value))
    
    Workbooks(FileDBName_OSp).Sheets("Лист1").Cells(rowCount, column_Ед_изм).Value = ThisWorkbook.Sheets("Лист8").Cells(Row_Лист8, 4).Value
    Workbooks(FileDBName_OSp).Sheets("Лист1").Cells(rowCount, column_Ед_изм).NumberFormat = "@" ' "#,##0"
    Workbooks(FileDBName_OSp).Sheets("Лист1").Cells(rowCount, column_Ед_изм).HorizontalAlignment = xlCenter ' xlRight

    ' План
    ' column_План = ColumnByValue(FileDBName_OSp, "Лист1", "План", 100, 100)
    Workbooks(FileDBName_OSp).Sheets("Лист1").Cells(rowCount, column_План).Value = ThisWorkbook.Sheets("Лист8").Cells(Row_Лист8, 5).Value
    Workbooks(FileDBName_OSp).Sheets("Лист1").Cells(rowCount, column_План).NumberFormat = "#,##0"
    Workbooks(FileDBName_OSp).Sheets("Лист1").Cells(rowCount, column_План).HorizontalAlignment = xlRight
  
    ' *UPDATE План Интенсив по ИЗП на 3 кв. 2021. Необходимо организовать в офисах : 1) цели по ИЗП до 30.09.2021 ОО "Тюменский" - 20 шт., ОО "Сургутский" - 12 шт., ОО "Нижневартовский" - 12 шт., ОО "Новоуренгойский" - 18 шт., ОО "Тарко-Сале" - 18 шт. 2) провести срез знаний (ссылка на материалы направлена в письме) 3) ежедневный норматив - 1 заявка на ИЗП в день от офиса
    If (quarterName2(dateDB_Лист8) = "3Q 2021 г.") And (Workbooks(FileDBName_OSp).Sheets("Лист1").Cells(rowCount, column_Product_Code).Value = "ИЗП") Then
      
      ' ОО "Тюменский" - 20 шт., ОО "Сургутский" - 12 шт., ОО "Нижневартовский" - 12 шт., ОО "Новоуренгойский" - 18 шт., ОО "Тарко-Сале" - 18 шт.
      Call UPDATE_План_Интенсив_ИЗП_3Q2021(FileDBName_OSp, "Лист1", rowCount, column_План, column_офис)
      
    End If
  
    ' Факт
    ' column_Факт = ColumnByValue(FileDBName_OSp, "Лист1", "Факт на ____", 100, 100)
    Workbooks(FileDBName_OSp).Sheets("Лист1").Cells(rowCount, column_Факт).Value = ThisWorkbook.Sheets("Лист8").Cells(Row_Лист8, 6).Value
    Workbooks(FileDBName_OSp).Sheets("Лист1").Cells(rowCount, column_Факт).NumberFormat = "#,##0"
    Workbooks(FileDBName_OSp).Sheets("Лист1").Cells(rowCount, column_Факт).HorizontalAlignment = xlRight
  
    ' Исп.
    ' column_Исп = ColumnByValue(FileDBName_OSp, "Лист1", "Исп.", 100, 100)
    Workbooks(FileDBName_OSp).Sheets("Лист1").Cells(rowCount, column_Исп).Value = ThisWorkbook.Sheets("Лист8").Cells(Row_Лист8, 7).Value
    Workbooks(FileDBName_OSp).Sheets("Лист1").Cells(rowCount, column_Исп).NumberFormat = "0%"
    Workbooks(FileDBName_OSp).Sheets("Лист1").Cells(rowCount, column_Исп).HorizontalAlignment = xlRight
  
    ' Прогноз
    ' column_Прогноз = ColumnByValue(FileDBName_OSp, "Лист1", "Прогноз", 100, 100)
    ' Если в column_рассчетный_прогноз стоит "Лист8_column17", то берем рассчетный прогноз из 17 столбца Лист8
    If Workbooks(FileDBName_OSp).Sheets("Лист1").Cells(rowCount, column_рассчетный_прогноз).Value = "Лист8_column20" Then
      столбец_прогноза_Лист8 = 20
    Else
      столбец_прогноза_Лист8 = 8
    End If
    Workbooks(FileDBName_OSp).Sheets("Лист1").Cells(rowCount, column_Прогноз).Value = ThisWorkbook.Sheets("Лист8").Cells(Row_Лист8, столбец_прогноза_Лист8).Value
    Workbooks(FileDBName_OSp).Sheets("Лист1").Cells(rowCount, column_Прогноз).NumberFormat = "0%"
    Workbooks(FileDBName_OSp).Sheets("Лист1").Cells(rowCount, column_Прогноз).HorizontalAlignment = xlRight
    ' Если в column_Прогноз есть значение - красим в Светофор
    If Workbooks(FileDBName_OSp).Sheets("Лист1").Cells(rowCount, column_Прогноз).Value <> "" Then
      ' Окраска ячейки СФЕТОФОР ' 70. Заливка ячейки цветом "светофор" - значение In_Value в %, In_Target в %
      Call Full_Color_RangeV(FileDBName_OSp, "Лист1", rowCount, column_Прогноз, Workbooks(FileDBName_OSp).Sheets("Лист1").Cells(rowCount, column_Прогноз).Value, 1)
    End If
     
    ' Прирост за прошлую неделю
    ' Факт
    ' column_Факт_прошлая_неделя = ColumnByValue(FileDBName_OSp, "Лист1", "Факт _", 100, 100)
    Workbooks(FileDBName_OSp).Sheets("Лист1").Cells(rowCount, column_Факт_прошлая_неделя).Value = ThisWorkbook.Sheets("Лист8").Cells(Row_Лист8, 15).Value
    Workbooks(FileDBName_OSp).Sheets("Лист1").Cells(rowCount, column_Факт_прошлая_неделя).NumberFormat = "#,##0"
    Workbooks(FileDBName_OSp).Sheets("Лист1").Cells(rowCount, column_Факт_прошлая_неделя).HorizontalAlignment = xlRight

    ' Изм.
    ' column_Изм = ColumnByValue(FileDBName_OSp, "Лист1", "Изм.", 100, 100)
    Workbooks(FileDBName_OSp).Sheets("Лист1").Cells(rowCount, column_Изм).Value = ThisWorkbook.Sheets("Лист8").Cells(Row_Лист8, 16).Value
    Workbooks(FileDBName_OSp).Sheets("Лист1").Cells(rowCount, column_Изм).NumberFormat = "#,##0"
    Workbooks(FileDBName_OSp).Sheets("Лист1").Cells(rowCount, column_Изм).HorizontalAlignment = xlRight
  
    ' Прогн.Q
    ' column_Прогн_Q = ColumnByValue(FileDBName_OSp, "Лист1", "Прогн.Q", 100, 100)
    Workbooks(FileDBName_OSp).Sheets("Лист1").Cells(rowCount, column_Прогн_Q).Value = ThisWorkbook.Sheets("Лист8").Cells(Row_Лист8, 17).Value
    Workbooks(FileDBName_OSp).Sheets("Лист1").Cells(rowCount, column_Прогн_Q).NumberFormat = "0%"
    Workbooks(FileDBName_OSp).Sheets("Лист1").Cells(rowCount, column_Прогн_Q).HorizontalAlignment = xlRight
  
    ' Динамика %
    ' column_Динамика = ColumnByValue(FileDBName_OSp, "Лист1", "%", 100, 100)
    Workbooks(FileDBName_OSp).Sheets("Лист1").Cells(rowCount, column_Динамика).Value = ThisWorkbook.Sheets("Лист8").Cells(Row_Лист8, 18).Value
    Workbooks(FileDBName_OSp).Sheets("Лист1").Cells(rowCount, column_Динамика).NumberFormat = "0%"
    Workbooks(FileDBName_OSp).Sheets("Лист1").Cells(rowCount, column_Динамика).HorizontalAlignment = xlRight
    ' Окраска Динамики:
    ' Окраска ячейки СФЕТОФОР: в красный, если отрицательная динамика и исполнее менее 1
    If (Workbooks(FileDBName_OSp).Sheets("Лист1").Cells(rowCount, column_Динамика).Value < 0) And (Workbooks(FileDBName_OSp).Sheets("Лист1").Cells(rowCount, column_Динамика).Value < 1) Then
      Call Full_Color_RangeV(FileDBName_OSp, "Лист1", rowCount, column_Динамика, 0, 100)
    End If

    ' Окраска ячейки СФЕТОФОР: в зеленый, если положительная динамика
    If Workbooks(FileDBName_OSp).Sheets("Лист1").Cells(rowCount, column_Динамика).Value > 0 Then
      Call Full_Color_RangeV(FileDBName_OSp, "Лист1", rowCount, column_Динамика, 100, 100)
    End If
  
    ' Поручение - "Поручение на неделю для выхода на Прогн.Q 100%"
    ' Дата исполнения  ____
    ' column_Поручение = ColumnByValue(FileDBName_OSp, "Лист1", "Дата исполнения  ____", 100, 100)
    
    ' === Сюда вставляем цель на неделю - сколько надо прирасти, чтобы выйти на прогноз Q в 100%
    If Workbooks(FileDBName_OSp).Sheets("Лист1").Cells(rowCount, column_Прогноз).Value < 1 Then
            
            ' Первый вариант: Считаем какой должен быть прогноз - на текущая дата DB + 7
            Факт_на_дату_для_прогноза_квартала_Var = Факт_на_дату_для_прогноза_квартала(dateDB_Лист8 + 7, Workbooks(FileDBName_OSp).Sheets("Лист1").Cells(rowCount, column_План).Value, 1, 5, 0)
            
            ' Если Факт для выхода на прогноз Q больше, чем текущий Факт Q, то считаем прирост
            If Факт_на_дату_для_прогноза_квартала_Var > Workbooks(FileDBName_OSp).Sheets("Лист1").Cells(rowCount, column_Факт).Value Then
              
              ' Если цель на неделю > 0
              If (Факт_на_дату_для_прогноза_квартала_Var - Workbooks(FileDBName_OSp).Sheets("Лист1").Cells(rowCount, column_Факт).Value) > 0 Then
              
                 Workbooks(FileDBName_OSp).Sheets("Лист1").Cells(rowCount, column_Поручение).Value = Факт_на_дату_для_прогноза_квартала_Var - Workbooks(FileDBName_OSp).Sheets("Лист1").Cells(rowCount, column_Факт).Value
                 Workbooks(FileDBName_OSp).Sheets("Лист1").Cells(rowCount, column_Поручение).NumberFormat = "#,##0"
                 Workbooks(FileDBName_OSp).Sheets("Лист1").Cells(rowCount, column_Поручение).HorizontalAlignment = xlRight
                 '
                 ' По просьбе шевелева добавляем ЦО дня в столбец "T" (20)
                 Workbooks(FileDBName_OSp).Sheets("Лист1").Cells(rowCount, 20).Value = (Факт_на_дату_для_прогноза_квартала_Var - Workbooks(FileDBName_OSp).Sheets("Лист1").Cells(rowCount, column_Факт).Value) / 5
                 Workbooks(FileDBName_OSp).Sheets("Лист1").Cells(rowCount, 20).NumberFormat = "#,##0"
                 Workbooks(FileDBName_OSp).Sheets("Лист1").Cells(rowCount, 20).HorizontalAlignment = xlRight
              
              End If
            End If
          
    End If
    ' ===
    
    ' === Если прогноз не рассчитан (например портфели)
    If Workbooks(FileDBName_OSp).Sheets("Лист1").Cells(rowCount, column_Прогноз).Value = "" Then
            
      ' Если цель на неделю > 0
      If (Workbooks(FileDBName_OSp).Sheets("Лист1").Cells(rowCount, column_План).Value - Workbooks(FileDBName_OSp).Sheets("Лист1").Cells(rowCount, column_Факт).Value) > 0 Then
      
        ' Вставляем в цели недели разницу между планом и фактом
        Workbooks(FileDBName_OSp).Sheets("Лист1").Cells(rowCount, column_Поручение).Value = Workbooks(FileDBName_OSp).Sheets("Лист1").Cells(rowCount, column_План).Value - Workbooks(FileDBName_OSp).Sheets("Лист1").Cells(rowCount, column_Факт).Value
        Workbooks(FileDBName_OSp).Sheets("Лист1").Cells(rowCount, column_Поручение).NumberFormat = "#,##0"
        Workbooks(FileDBName_OSp).Sheets("Лист1").Cells(rowCount, column_Поручение).HorizontalAlignment = xlRight
                 
        ' По просьбе шевелева добавляем ЦО дня в столбец "T" (20)
        Workbooks(FileDBName_OSp).Sheets("Лист1").Cells(rowCount, 20).Value = Workbooks(FileDBName_OSp).Sheets("Лист1").Cells(rowCount, column_Поручение).Value / 5
        Workbooks(FileDBName_OSp).Sheets("Лист1").Cells(rowCount, 20).NumberFormat = "#,##0"
        Workbooks(FileDBName_OSp).Sheets("Лист1").Cells(rowCount, 20).HorizontalAlignment = xlRight
      
      End If
      
    End If
    ' ===


    
    ' Следующая запись
    Application.StatusBar = "Оперативная справка за неделю: " + CStr(rowCount) + "..."
    rowCount = rowCount + 1
    DoEventsInterval (rowCount)
    
  Loop
  
  ' Строка статуса
  Application.StatusBar = "Оперативная справка за неделю: завершение..."
  
  ' Закрытие файла
  Workbooks(FileDBName_OSp).Close SaveChanges:=True
  
  ' Закрываем таблицы, использовавшиеся в функциях
  ' Закрываем BASE\Products
  CloseBook ("Products")
  
  ' Строка статуса
  Application.StatusBar = "Оперативная справка за неделю: отправка..."
  
  Call Отправка_Lotus_Notes_Оперативка_Лист8
  
  ' Строка статуса
  Application.StatusBar = ""
   
End Sub


' *UPDATE План Интенсив по ИЗП на 3 кв. 2021. Необходимо организовать в офисах : 1) цели по ИЗП до 30.09.2021 ОО "Тюменский" - 20 шт., ОО "Сургутский" - 12 шт., ОО "Нижневартовский" - 12 шт., ОО "Новоуренгойский" - 18 шт., ОО "Тарко-Сале" - 18 шт. 2) провести срез знаний (ссылка на материалы направлена в письме) 3) ежедневный норматив - 1 заявка на ИЗП в день от офиса
Sub UPDATE_План_Интенсив_ИЗП_3Q2021(In_FileDBName_OSp, In_Sheets, In_rowCount, In_column_План, In_column_офис)
      
      ' ОО "Тюменский" - 20 шт.,
      If Workbooks(In_FileDBName_OSp).Sheets(In_Sheets).Cells(In_rowCount, In_column_офис).Value = 1 Then
        Workbooks(In_FileDBName_OSp).Sheets(In_Sheets).Cells(In_rowCount, In_column_План).Value = 20
      End If
      
      ' ОО "Сургутский" - 12 шт.,
      If Workbooks(In_FileDBName_OSp).Sheets(In_Sheets).Cells(In_rowCount, In_column_офис).Value = 2 Then
        Workbooks(In_FileDBName_OSp).Sheets(In_Sheets).Cells(In_rowCount, In_column_План).Value = 12
      End If
      
      ' ОО "Нижневартовский" - 12 шт.,
      If Workbooks(In_FileDBName_OSp).Sheets(In_Sheets).Cells(In_rowCount, In_column_офис).Value = 3 Then
        Workbooks(In_FileDBName_OSp).Sheets(In_Sheets).Cells(In_rowCount, In_column_План).Value = 12
      End If
      
      ' ОО "Новоуренгойский" - 18 шт.,
      If Workbooks(In_FileDBName_OSp).Sheets(In_Sheets).Cells(In_rowCount, In_column_офис).Value = 4 Then
        Workbooks(In_FileDBName_OSp).Sheets(In_Sheets).Cells(In_rowCount, In_column_План).Value = 18
      End If
      
      ' ОО "Тарко-Сале" - 18 шт.
      If Workbooks(In_FileDBName_OSp).Sheets(In_Sheets).Cells(In_rowCount, In_column_офис).Value = 5 Then
        Workbooks(In_FileDBName_OSp).Sheets(In_Sheets).Cells(In_rowCount, In_column_План).Value = 18
      End If
    
End Sub

' Отправка письма: отправляю шаблон самому себе для последующей отправки в сеть письма на его основе:
Sub Отправка_Lotus_Notes_Оперативка_Лист8()
Dim темаПисьма, текстПисьма, hashTag, attachmentFile As String
Dim i As Byte
  
  ' Строка статуса
  Application.StatusBar = "Отправка письма..."
  
    
    dateDB_Лист8 = CDate(Mid(ThisWorkbook.Sheets("Лист8").Range("B5").Value, 52, 10))
   
    ' Тема письма - Тема:
    ' темаПисьма = ThisWorkbook.Sheets("Лист8").Cells(RowByValue(ThisWorkbook.Name, "Лист8", "Тема:", 100, 100), ColumnByValue(ThisWorkbook.Name, "Лист8", "Тема:", 100, 100) + 1).Value
    темаПисьма = "Оперативная справка по итогам недели на " + CStr(dateDB_Лист8) + " г."

    ' hashTag - Хэштэг:
    hashTag = "#оперативная_справка #поручения"

    ' Файл-вложение (!!!)
    attachmentFile = ThisWorkbook.Sheets("Лист8").Range("S3").Value
    
    ' Текст письма
    текстПисьма = "" + Chr(13)
    текстПисьма = текстПисьма + "" + getFromAddrBook("НОКП,РРКК,РИЦ,НОРПиКО1,УДО2,НОРПиКО2,УДО3,НОРПиКО3,УДО4,НОРПиКО4,УДО5,НОРПиКО5", 2) + Chr(13) + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "" + getFromAddrBook("РД", 2) + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "Уважаемые руководители," + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "Направляю оперативную справку по итогам прошлой недели на " + CStr(dateDB_Лист8) + " г. (файл во вложении)" + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "В столбце " + кавычки() + "P" + кавычки() + " - поручения руководителям подразделений по объемам прироста показателей на неделю (" + strDDMM(Date) + "-" + strДД_MM_YY2(Date + 6) + "), " + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "необходимым для выхода на прогноз исп. БП " + quarterName3(dateDB_Лист8) + " Тюменского РОО не менее 100%" + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "(продолжение в файле)" + Chr(13)
    ' Визитка (подпись С Ув., )
    текстПисьма = текстПисьма + ПодписьВПисьме()
    ' Хэштег
    текстПисьма = текстПисьма + createBlankStr(27) + hashTag
    ' Вызов
    Call send_Lotus_Notes(темаПисьма, "Sergey Fedorovich Proschaev/Tyumen/PSBank/Ru", "Sergey Fedorovich Proschaev/Tyumen/PSBank/Ru", текстПисьма, attachmentFile)
  
    ' Сообщение
    ' MsgBox ("Письмо отправлено!")
     
    ' Строка статуса
    Application.StatusBar = ""
     
  
End Sub


' Не доджелал!!! Не Запускать!!!
' Проадейтить планы на неделю по столбцу M на Лист8 по датам из M9
Sub UPDATE_Лист8_column_M9()
  
  ' Запрос на исполнение процедуры
  If MsgBox("Запустит Update столбца M? Не доджелал!!! Не Запускать!!!", vbYesNo) = vbYes Then
  
  
  ' Открываем таблицы, которые нужны:
  ' Открываем BASE\Products
  OpenBookInBase ("Products")

  
  ' 1) Обработка офисов
  For i = 1 To 5
    
    ' Номера офисов от 1 до 5
    Select Case i
      Case 1 ' ОО «Тюменский»
        officeNameInReport = "ОО «Тюменский»"
      Case 2 ' ОО «Сургутский»
        officeNameInReport = "ОО «Сургутский»"
      Case 3 ' ОО «Нижневартовский»
        officeNameInReport = "ОО «Нижневартовский»"
      Case 4 ' ОО «Новоуренгойский»
        officeNameInReport = "ОО «Новоуренгойский»"
      Case 5 ' ОО «Тарко-Сале»
        officeNameInReport = "ОО «Тарко-Сале»"
            
    End Select
        
    ' Находим номер строки с наименованием офиса
    row_офис = getRowFromSheet8(officeNameInReport, officeNameInReport)
        
    ' Обрабатываем блок офиса
    rowCount = row_офис + 3
    Do While (ThisWorkbook.Sheets("Лист8").Cells(rowCount, 1).Value <> "") And (ThisWorkbook.Sheets("Лист8").Cells(rowCount, 1).Value <> ".")
' ===================
  
  
    ' Квартал - факт
    ' Если измерение в %
    If In_Unit <> "%" Then
          
          ' Квартал факт
          ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 6).Value = Workbooks(In_ReportName_String).Sheets(In_Sheets).Cells(rowCount, Column_Продажи_Квартал_Факт).Value
          
          ' === Сюда вставляем цель на неделю - сколько надо прирасти, чтобы выйти на прогноз Q в 100%
          If Прогноз_квартала_проц(dateDB, ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 5).Value, ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 6).Value, 5, 0) < 1 Then
            
            ' Первый вариант: Считаем какой должен быть прогноз - на текущая дата DB + 7
            ' Факт_на_дату_для_прогноза_квартала_Var = Факт_на_дату_для_прогноза_квартала(dateDB + 7, ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 5).Value, 1, 5, 0)
          
            ' Второй вариант: Считаем какой должен быть прогноз - из M9 "План на неделю: (02.08-08.08.21)" берем вторую дату
            date2FromM9 = CDate(Mid(ThisWorkbook.Sheets("Лист8").Range("M9").Value, 24, 6) + "20" + Mid(ThisWorkbook.Sheets("Лист8").Range("M9").Value, 30, 2))
            
            ' Отставание DB: от воскресенья (конец недели ) - 3 дня = четверг!
            date2FromM9 = date2FromM9 - 3
            Факт_на_дату_для_прогноза_квартала_Var = Факт_на_дату_для_прогноза_квартала(date2FromM9, ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 5).Value, 1, 5, 0)
            
            ' Если Факт для выхода на прогноз Q больше, чем текущий Факт Q, то считаем прирост
            If Факт_на_дату_для_прогноза_квартала_Var > ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 6).Value Then
              ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 13).Value = Факт_на_дату_для_прогноза_квартала_Var - ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 6).Value
              ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 13).NumberFormat = "#,##0"
              ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 13).HorizontalAlignment = xlRight
            End If
          
          End If
          ' ===
    End If

  
' =============================================================================
      ' Следующая запись
      Application.StatusBar = officeNameInReport + ": " + CStr(ThisWorkbook.Sheets("Лист8").Cells(rowCount, 1).Value) + "..."
      rowCount = rowCount + 1
      DoEventsInterval (rowCount)
    
    Loop
        
  Next i
  
  End If ' Запрос
  
  
End Sub


' Отправка письма: отправляю шаблон самому себе для последующей отправки в сеть письма на его основе:
Sub Отправка_Lotus_Notes_Лист8_ИПЗ()
Dim темаПисьма, текстПисьма, hashTag, attachmentFile As String
Dim i As Byte
  
  ' Строка статуса
  Application.StatusBar = "Отправка письма с ИПЗ..."
  
  
  ' Запрос
  ' If MsgBox("Отправить себе Шаблон письма с фокусами контроля '" + ПериодКонтроля + "'?", vbYesNo) = vbYes Then
    
    ' Тема письма - Тема:
    ' темаПисьма = ThisWorkbook.Sheets("Лист8").Cells(RowByValue(ThisWorkbook.Name, "Лист8", "Тема:", 100, 100), ColumnByValue(ThisWorkbook.Name, "Лист8", "Тема:", 100, 100) + 1).Value
    темаПисьма = "ИПЗ на 3Q 2021 (УДО, НОРПиКО) офисные продажи, привлечение"

    ' hashTag - Хэштэг:
    ' hashTag = ThisWorkbook.Sheets("Лист8").Cells(RowByValue(ThisWorkbook.Name, "Лист8", "Хэштэг:", 100, 100), ColumnByValue(ThisWorkbook.Name, "Лист8", "Хэштэг:", 100, 100) + 1).Value
    ' hashTag - Хэштэг:
    hashTag = "#ипз #ипз_удо #ипз_норпико"

    ' Файл-вложение (!!!)
    attachmentFile = ""
    
    ' Текст письма
    текстПисьма = "" + Chr(13)
    текстПисьма = текстПисьма + "" + getFromAddrBook("НОРПиКО1,УДО2,НОРПиКО2,УДО3,НОРПиКО3,УДО4,НОРПиКО4,УДО5,НОРПиКО5", 2) + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "" + getFromAddrBook("РД", 2) + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "Уважаемые руководители," + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "Направляю Индивидуальные плановые задания на 3Q 2021 г." + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "Прошу принять к исполнению." + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    ' Визитка (подпись С Ув., )
    текстПисьма = текстПисьма + ПодписьВПисьме()
    ' Хэштег
    текстПисьма = текстПисьма + createBlankStr(27) + hashTag
    ' Вызов
    Call send_Lotus_Notes(темаПисьма, "Sergey Fedorovich Proschaev/Tyumen/PSBank/Ru", "Sergey Fedorovich Proschaev/Tyumen/PSBank/Ru", текстПисьма, attachmentFile)
  
    ' Зачеркнуть
    Call ЗачеркиваемТекстВячейке("Лист0", RangeByValue(ThisWorkbook.Name, "Лист0", "DashBoard (при наличии)", 100, 100))
  
    ' Сообщение
    ' MsgBox ("Письмо отправлено!")
     
    ' Строка статуса
    Application.StatusBar = ""
     
  ' End If
  
End Sub



' Запрос для Продажи_квартала_за_период()
Sub Продажи_квартала_за_период_with_Msg()
  
  ' Запрос "12.08.2021-19.08.2021"
  If MsgBox("Рассчитать продажи с " + CStr(Дата1(ThisWorkbook.Sheets("Лист8").Range("S9").Value)) + " по " + CStr(Дата2(ThisWorkbook.Sheets("Лист8").Range("S9").Value)) + " ?", vbYesNo) = vbYes Then
    Call Продажи_квартала_за_период
    MsgBox ("Данные загружены!")
  End If
  
  
End Sub

' Продажи_квартала_за_период()
Sub Продажи_квартала_за_период()
  
  ' Открываем BASE\Sales
  OpenBookInBase ("Sales_Office")
  
  ' Открываем BASE\Products
  OpenBookInBase ("Products")

  ' Дата начала
  Дата1_Var = Дата1(ThisWorkbook.Sheets("Лист8").Range("S9").Value)
  ' Дата окончания
  Дата2_Var = Дата2(ThisWorkbook.Sheets("Лист8").Range("S9").Value)
  
  ' Очищаем столбец
  ' 1. ОО «Тюменский»
  Call clearСontents2(ThisWorkbook.Name, "Лист8", "S" + CStr(getRowFromSheet8("ОО «Тюменский»", "ОО «Тюменский»") + 3), "S" + CStr(getRowFromSheet8("ОО «Сургутский»", "ОО «Сургутский»") - 1))
  ' 2. ОО «Сургутский»
  Call clearСontents2(ThisWorkbook.Name, "Лист8", "S" + CStr(getRowFromSheet8("ОО «Сургутский»", "ОО «Сургутский»") + 3), "S" + CStr(getRowFromSheet8("ОО «Нижневартовский»", "ОО «Нижневартовский»") - 1))
  ' 3. ОО «Нижневартовский»
  Call clearСontents2(ThisWorkbook.Name, "Лист8", "S" + CStr(getRowFromSheet8("ОО «Нижневартовский»", "ОО «Нижневартовский»") + 3), "S" + CStr(getRowFromSheet8("ОО «Новоуренгойский»", "ОО «Новоуренгойский»") - 1))
  ' 4. ОО «Новоуренгойский»
  Call clearСontents2(ThisWorkbook.Name, "Лист8", "S" + CStr(getRowFromSheet8("ОО «Новоуренгойский»", "ОО «Новоуренгойский»") + 3), "S" + CStr(getRowFromSheet8("ОО «Тарко-Сале»", "ОО «Тарко-Сале»") - 1))
  ' 5. ОО «Тарко-Сале»
  Call clearСontents2(ThisWorkbook.Name, "Лист8", "S" + CStr(getRowFromSheet8("ОО «Тарко-Сале»", "ОО «Тарко-Сале»") + 3), "S" + CStr(getRowFromSheet8("Интегральный рейтинг по офисам", "Интегральный рейтинг по офисам") - 2))
  ' 6. РОО Тюменский
  Call clearСontents2(ThisWorkbook.Name, "Лист8", "S" + CStr(getRowFromSheet8("Итого по РОО «Тюменский»", "Итого по РОО «Тюменский»") + 3), "S" + CStr(getRowFromSheet8("Итого по РОО «Тюменский»", "Итого по РОО «Тюменский»") + (getRowFromSheet8("ОО «Сургутский»", "ОО «Сургутский»") - getRowFromSheet8("ОО «Тюменский»", "ОО «Тюменский»"))))
  
  ' Обработка по 5 офисам
  For i = 1 To 5
    
    ' Номера офисов от 1 до 5
    Select Case i
      Case 1 ' ОО «Тюменский»
        officeNameInReport = "ОО «Тюменский»"
      Case 2 ' ОО «Сургутский»
        officeNameInReport = "ОО «Сургутский»"
      Case 3 ' ОО «Нижневартовский»
        officeNameInReport = "ОО «Нижневартовский»"
      Case 4 ' ОО «Новоуренгойский»
        officeNameInReport = "ОО «Новоуренгойский»"
      Case 5 ' ОО «Тарко-Сале»
        officeNameInReport = "ОО «Тарко-Сале»"
    End Select
  
    ' Обработка блока
    row_офис_N = getRowFromSheet8(officeNameInReport, officeNameInReport) + 3
    Do While (ThisWorkbook.Sheets("Лист8").Cells(row_офис_N, 1).Value <> "") And (ThisWorkbook.Sheets("Лист8").Cells(row_офис_N, 1).Value <> ".")
      
      ' Рассчет значений
      If ThisWorkbook.Sheets("Лист8").Cells(row_офис_N, 4).Value <> "%" Then
        
        ThisWorkbook.Sheets("Лист8").Cells(row_офис_N, 19).Value = Продажи_Q_за_период(i, Product_Name_to_Product_Code(ThisWorkbook.Sheets("Лист8").Cells(row_офис_N, 2).Value), Дата1_Var, Дата2_Var)
        ThisWorkbook.Sheets("Лист8").Cells(row_офис_N, 19).NumberFormat = "#,##0"
        ThisWorkbook.Sheets("Лист8").Cells(row_офис_N, 19).HorizontalAlignment = xlRight
        
      End If
      
      ' Следующая запись
      Application.StatusBar = officeNameInReport + ": " + CStr(row_офис_N) + "..."
      row_офис_N = row_офис_N + 1
      DoEventsInterval (row_офис_N)
      
    Loop
    
  Next i
  
  ' Свод по РОО
  officeNameInReport = "Итого по РОО «Тюменский»"
  
  ' Обработка блока
  row_офис_N = getRowFromSheet8(officeNameInReport, officeNameInReport) + 3
  Do While (ThisWorkbook.Sheets("Лист8").Cells(row_офис_N, 1).Value <> "")
      
      ' Рассчет значений
      If ThisWorkbook.Sheets("Лист8").Cells(row_офис_N, 4).Value <> "%" Then
        
        ThisWorkbook.Sheets("Лист8").Cells(row_офис_N, 19).Value = CheckData(ThisWorkbook.Sheets("Лист8").Cells(getRowFromSheet8("ОО «Тюменский»", ThisWorkbook.Sheets("Лист8").Cells(row_офис_N, 2).Value), 19).Value) + _
                                                                     CheckData(ThisWorkbook.Sheets("Лист8").Cells(getRowFromSheet8("ОО «Сургутский»", ThisWorkbook.Sheets("Лист8").Cells(row_офис_N, 2).Value), 19).Value) + _
                                                                       CheckData(ThisWorkbook.Sheets("Лист8").Cells(getRowFromSheet8("ОО «Нижневартовский»", ThisWorkbook.Sheets("Лист8").Cells(row_офис_N, 2).Value), 19).Value) + _
                                                                         CheckData(ThisWorkbook.Sheets("Лист8").Cells(getRowFromSheet8("ОО «Новоуренгойский»", ThisWorkbook.Sheets("Лист8").Cells(row_офис_N, 2).Value), 19).Value) + _
                                                                           CheckData(ThisWorkbook.Sheets("Лист8").Cells(getRowFromSheet8("ОО «Тарко-Сале»", ThisWorkbook.Sheets("Лист8").Cells(row_офис_N, 2).Value), 19).Value)
        ThisWorkbook.Sheets("Лист8").Cells(row_офис_N, 19).NumberFormat = "#,##0"
        ThisWorkbook.Sheets("Лист8").Cells(row_офис_N, 19).HorizontalAlignment = xlRight
        
      End If
      
    ' Следующая запись
    Application.StatusBar = officeNameInReport + ": " + CStr(row_офис_N) + "..."
    row_офис_N = row_офис_N + 1
    DoEventsInterval (row_офис_N)
      
  Loop
  
  
  ' Закрываем BASE\Sales
  CloseBook ("Sales_Office")
  
  ' Закрываем BASE\Products
  CloseBook ("Products")
  
  ' Строка статуса
  Application.StatusBar = ""
  
  
End Sub

' DB_Штат
Sub DB_Штат(In_ReportName_String, In_Sheets, In_officeNameInReport, In_Row_Лист8, In_N, In_Product_Name, In_Product_Code, In_Unit, In_Weight)
Dim dateDB As Date
  
  dateDB = CDate(Mid(Workbooks(In_ReportName_String).Sheets("Оглавление").Cells(1, 1).Value, 23, 10))
    
  ' Апдейтим таблицу BASE\Products
  Call Update_BASE_Products(In_Product_Name, In_Product_Code, In_Unit)
       
  ' Определяем число сотрудников из функции Число_сотрудников_в_офисе_Лист7
       
      
  ' 1. Заносим на Лист 8 данные по ипотеке в квартал
  ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 1).NumberFormat = "@"
  ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 1).Value = In_N
  ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 1).HorizontalAlignment = xlCenter
  '
  ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 2).Value = In_Product_Name
  ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 2).HorizontalAlignment = xlLeft
  ' Вес выводим, если он не нулевой
  If In_Weight <> 0 Then
    ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 3).Value = In_Weight
    ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 3).NumberFormat = "0.0%"
    ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 3).HorizontalAlignment = xlCenter
  End If
  '
  ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 4).Value = In_Unit
  ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 4).HorizontalAlignment = xlCenter
      
  ' Месяц - план
  ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 9).Value = План_штат_МРК_офис(getNumberOfficeByName2(In_officeNameInReport)) ' Workbooks(In_ReportName_String).Sheets("Лист1").Cells(rowCount_DB_Лист1, column_DB_Лист1_План_руб_Q_Ипотека).Value
  ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 9).NumberFormat = "#,##0"
  ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 9).HorizontalAlignment = xlRight

  ' Месяц - факт
  ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 10).Value = Число_сотрудников_в_офисе_Лист7(getNumberOfficeByName2(In_officeNameInReport)) ' Workbooks(In_ReportName_String).Sheets("Лист1").Cells(rowCount_DB_Лист1, column_DB_Лист1_Факт_руб_Q_Ипотека).Value
  ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 10).NumberFormat = "#,##0"
  ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 10).HorizontalAlignment = xlRight

  ' Месяц - исполнение (в %)
  ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 11).Value = РассчетДоли(ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 9).Value, ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 10).Value, 3)
  ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 11).NumberFormat = "0%"
  ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 11).HorizontalAlignment = xlRight
        
  ' 2. Заносим в БД
      
  ' Заносим в Sales_Office
  '  Идентификатор ID_Rec:
  ' ID_RecVar = CStr(CStr(getNumberOfficeByName(In_officeNameInReport)) + "-" + strNQYY(dateDB) + "-" + In_Product_Code)
                        
  ' Текущие значения в месяце: Date_01 (N), Date_02 (O), Date_03 (P), Date_04 (Q), Date_05 Date_06 Date_07 Date_08 Date_09 Date_10 Date_11 Date_12 Date_13 Date_14 Date_15 Date_16 Date_17 Date_18 Date_19 Date_20 Date_21 Date_22 Date_23 Date_24 Date_25 Date_26 Date_27 Date_28 Date_29 Date_30 Date_31
  ' Номер месяца в квартале: 1-"", 2-"2", 3-"3"
  ' M_num = Nom_mes_quarter_str(dateDB)
  ' curr_Day_Month_Q = "Date" + M_num + "_" + Mid(dateDB, 1, 2)
         
  '  Идентификатор ID_Rec:
  ID_RecVar = CStr(CStr(getNumberOfficeByName(In_officeNameInReport)) + "-" + strMMYY(dateDB) + "-" + In_Product_Code)
            
  ' Текущие значения в месяце: Date_01 (N), Date_02 (O), Date_03 (P), Date_04 (Q), Date_05 Date_06 Date_07 Date_08 Date_09 Date_10 Date_11 Date_12 Date_13 Date_14 Date_15 Date_16 Date_17 Date_18 Date_19 Date_20 Date_21 Date_22 Date_23 Date_24 Date_25 Date_26 Date_27 Date_28 Date_29 Date_30 Date_31
  curr_Day_Month = "Date_" + Mid(dateDB, 1, 2)
         
  ' Вносим данные в BASE\Sales_Office по ПК.
  Call InsertRecordInBook("Sales_Office", "Лист1", "ID_Rec", ID_RecVar, _
                                        "ID_Rec", ID_RecVar, _
                                          "Оffice_Number", getNumberOfficeByName(In_officeNameInReport), _
                                            "Product_Name", In_Product_Name, _
                                              "Оffice", In_officeNameInReport, _
                                                "MMYY", strMMYY(dateDB), _
                                                  "Update_Date", dateDB, _
                                                    "Product_Code", In_Product_Code, _
                                                      "Plan", ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 9).Value, _
                                                        "Unit", In_Unit, _
                                                          "Fact", ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 10).Value, _
                                                            "Percent_Completion", ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 11).Value, _
                                                              curr_Day_Month, ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 10).Value, _
                                                                "", "", _
                                                                  "", "", _
                                                                    "", "", _
                                                                      "", "", _
                                                                        "", "", _
                                                                          "", "", _
                                                                            "", "", _
                                                                              "", "")

      
  
End Sub


