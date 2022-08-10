Attribute VB_Name = "Module_Лист7"
' Интегральный_рейтинг_по_сотрудникам
Sub Интегральный_рейтинг_по_сотрудникам()
  
' Описание переменных
Dim ReportName_String, officeNameInReport, CheckFormatReportResult, ID_RecVar As String
Dim i, rowCount, row_DP3_отчет, column_TAB_OK, column_ФИО, column_DP3_отчет, column_DP4_отчет, recInЛист7, порядковый_номер As Integer
Dim finishProcess As Boolean
Dim dateDB As Date
    
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
    ThisWorkbook.Sheets("Лист7").Activate

    ' Проверка формы отчета
    CheckFormatReportResult = CheckFormatReport(ReportName_String, "Оглавление", 1, Date)
    If CheckFormatReportResult = "OK" Then
      
      ' Наименование листа Интегральный рей-г по сотруд
      StringInSheet = "Интегральный рей-г по сотруд"
      SheetName_String = FindNameSheet(ReportName_String, StringInSheet) ' "4. Интегральный рей-г по сотруд"
      If SheetName_String <> "" Then
      
      ' Очищаем ячейки отчета
      Call clearСontents2(ThisWorkbook.Name, "Лист7", "A9", "BK50")
      
      ' Заголовки
      ThisWorkbook.Sheets("Лист7").Cells(7, 6).Value = "Продукт1"
      ThisWorkbook.Sheets("Лист7").Cells(7, 11).Value = "Продукт2"
      ThisWorkbook.Sheets("Лист7").Cells(7, 16).Value = "Продукт3"
      ThisWorkbook.Sheets("Лист7").Cells(7, 21).Value = "Продукт4"
      ThisWorkbook.Sheets("Лист7").Cells(7, 26).Value = "Продукт5"
      ThisWorkbook.Sheets("Лист7").Cells(7, 31).Value = "Продукт6"
      ThisWorkbook.Sheets("Лист7").Cells(7, 36).Value = "Продукт7"
      ThisWorkbook.Sheets("Лист7").Cells(7, 41).Value = "Продукт8"
      ThisWorkbook.Sheets("Лист7").Cells(7, 46).Value = "Продукт9"
      ThisWorkbook.Sheets("Лист7").Cells(7, 51).Value = "Продукт10"
      ThisWorkbook.Sheets("Лист7").Cells(7, 56).Value = "Продукт11"

      ' Открываем BASE\Sales
      OpenBookInBase ("Sales")
            
      ' Открываем BASE\ActiveStaff
      OpenBookInBase ("ActiveStaff")


      ' На вкладке DB определяем:
      column_TAB_OK = ColumnByValue(Workbooks(ReportName_String).Name, SheetName_String, "TAB_OK", 100, 100) ' 1
      column_ФИО = ColumnByValue(Workbooks(ReportName_String).Name, SheetName_String, "ФИО", 100, 100) ' 2
      row_DP3_отчет = rowByValue(Workbooks(ReportName_String).Name, SheetName_String, "DP3_отчет", 100, 100) ' 9
      column_DP3_отчет = ColumnByValue(Workbooks(ReportName_String).Name, SheetName_String, "DP3_отчет", 100, 100) ' 4
      column_DP4_отчет = ColumnByValue(Workbooks(ReportName_String).Name, SheetName_String, "DP4_отчет", 100, 100) ' 5
      
      recInЛист7 = 8
      порядковый_номер = 0
      
      ' Заголовок отчета (из A1 - Отчет по состоянию на 07.07.2020)
      ThisWorkbook.Sheets("Лист7").Cells(5, 2).Value = "Интегральный рейтинг по сотрудникам на " + Mid(Workbooks(ReportName_String).Sheets("Оглавление").Cells(1, 1).Value, 23, 10) + " г."
      
     ' Из B5 "Интегральный рейтинг по сотрудникам на 15.07.2020 г." берем дату
      dateDB = CDate(Mid(ThisWorkbook.Sheets("Лист7").Range("B5").Value, 40, 10))

      ' Продукты: 10 шт.:
      
      ' 1. Потреб кредитование в тыс.руб.
      column_Наименование_Продукт1 = ColumnByValue(Workbooks(ReportName_String).Name, SheetName_String, "Потреб кредитование", 100, 100) ' 7
      Продукт1_Product_Code = "ПК"
      Продукт1_Unit = "тыс. руб."
      
      ' Если не найден, то проверяем "План_ПК, шт."
      If column_Наименование_Продукт1 = 0 Then
        column_Наименование_Продукт1 = ColumnByValue(Workbooks(ReportName_String).Name, SheetName_String, "План_ПК, шт.", 100, 100) ' 7
        Продукт1_Product_Code = "ПК_шт"
        Продукт1_Unit = "шт."
      End If
      
      column_Продукт1_План = column_Наименование_Продукт1 ' 7
      column_Продукт1_Факт = column_Наименование_Продукт1 + 1 ' 8
      column_Продукт1_Прогноз = column_Наименование_Продукт1 + 2 ' 9
      column_Продукт1_Вып_проц = column_Наименование_Продукт1 + 3 ' 10
      column_Продукт1_Прогноз_проц = column_Наименование_Продукт1 + 4 ' 11
      
      ' 2. Страховки к ПК
      column_Наименование_Продукт2 = ColumnByValue(Workbooks(ReportName_String).Name, SheetName_String, "Страховки к ПК", 100, 100) ' 12
      column_Продукт2_План = ColumnByValue(Workbooks(ReportName_String).Name, SheetName_String, "План_Страховки к ПК, тыс. руб. ", 100, 100) ' column_Наименование_Продукт2 ' 12
      column_Продукт2_Факт = ColumnByValue(Workbooks(ReportName_String).Name, SheetName_String, "Факт_Страховки к ПК, тыс. руб. ", 100, 100) ' column_Наименование_Продукт2 + 1 ' 13
      column_Продукт2_Прогноз = ColumnByValue(Workbooks(ReportName_String).Name, SheetName_String, "Прогноз_Страховки к ПК, тыс. руб. ", 100, 100) ' column_Наименование_Продукт2 + 2 ' 14
      column_Продукт2_Вып_проц = ColumnByValue(Workbooks(ReportName_String).Name, SheetName_String, "% Вып_Страховки к ПК_Факт", 100, 100) ' column_Наименование_Продукт2 + 3 ' 15
      column_Продукт2_Прогноз_проц = ColumnByValue(Workbooks(ReportName_String).Name, SheetName_String, "% Вып_Страховки к ПК_Прог", 100, 100) ' column_Наименование_Продукт2 + 4 ' 16

      ' 3. Кредитные карты
      column_Наименование_Продукт3 = ColumnByValue(Workbooks(ReportName_String).Name, SheetName_String, "Кредитные карты", 100, 100) ' 17
      column_Продукт3_План = column_Наименование_Продукт3 ' 17
      column_Продукт3_Факт = column_Наименование_Продукт3 + 1 ' 18
      column_Продукт3_Прогноз = column_Наименование_Продукт3 + 2 ' 19
      column_Продукт3_Вып_проц = column_Наименование_Продукт3 + 3 ' 20
      column_Продукт3_Прогноз_проц = column_Наименование_Продукт3 + 4 ' 21

      ' 4. Дебетовые карты
      column_Наименование_Продукт4 = ColumnByValue(Workbooks(ReportName_String).Name, SheetName_String, "Дебетовые карты", 100, 100)  ' 22
      column_Продукт4_План = column_Наименование_Продукт4 ' 22
      column_Продукт4_Факт = column_Наименование_Продукт4 + 1 ' 23
      column_Продукт4_Прогноз = column_Наименование_Продукт4 + 2 '24
      column_Продукт4_Вып_проц = column_Наименование_Продукт4 + 3 ' 25
      column_Продукт4_Прогноз_проц = column_Наименование_Продукт4 + 4 ' 26

      ' 5. Интернет Банк
      column_Наименование_Продукт5 = ColumnByValue(Workbooks(ReportName_String).Name, SheetName_String, "Интернет Банк", 100, 100) ' 27
      column_Продукт5_План = column_Наименование_Продукт5 ' 27
      column_Продукт5_Факт = column_Наименование_Продукт5 + 1 ' 28
      column_Продукт5_Прогноз = column_Наименование_Продукт5 + 2 ' 29
      column_Продукт5_Вып_проц = column_Наименование_Продукт5 + 3 ' 30
      column_Продукт5_Прогноз_проц = column_Наименование_Продукт5 + 4 ' 31

      ' 6. Портфель пассивов, тыс. руб. (ранее - Накопительный счет)
      column_Наименование_Продукт6 = ColumnByValue(Workbooks(ReportName_String).Name, SheetName_String, "Портфель пассивов+АУМ, тыс. руб.", 100, 100) ' "Портфель пассивов, тыс. руб.", 100, 100) ' ColumnByValue(Workbooks(ReportName_String).Name, SheetName_String, "Накопительный счет", 10, 100) ' 32
      column_Продукт6_План = column_Наименование_Продукт6 ' 32
      column_Продукт6_Факт = column_Наименование_Продукт6 + 1 ' 33
      column_Продукт6_Прогноз = 0 ' column_Наименование_Продукт6 + 2 ' 34
      column_Продукт6_Вып_проц = column_Наименование_Продукт6 + 2 ' column_Наименование_Продукт6 + 3 ' 35
      column_Продукт6_Прогноз_проц = 0 ' column_Наименование_Продукт6 + 4 ' 36

      ' 7. ИСЖ_МАСС (Премия, тыс.руб.)
      column_Наименование_Продукт7 = ColumnByValue(Workbooks(ReportName_String).Name, SheetName_String, "ИСЖ_МАСС (Премия, тыс.руб.)", 100, 100) ' 37
      column_Продукт7_План = column_Наименование_Продукт7 ' 37
      column_Продукт7_Факт = column_Наименование_Продукт7 + 1 ' 38
      column_Продукт7_Прогноз = column_Наименование_Продукт7 + 2 ' 39
      column_Продукт7_Вып_проц = column_Наименование_Продукт7 + 3 ' 40
      column_Продукт7_Прогноз_проц = column_Наименование_Продукт7 + 4 ' 41

      ' 8. НСЖ_МАСС (комиссионный доход) или НСЖ_МАСС (Премия, тыс.руб.)
      column_Наименование_Продукт8 = ColumnByValue(Workbooks(ReportName_String).Name, SheetName_String, "НСЖ_МАСС (комиссионный доход)", 100, 100) ' 42
      Продукт8_Product_Code = "НСЖ"
      Продукт8_Unit = "тыс. руб."
      If column_Наименование_Продукт8 = 0 Then
        column_Наименование_Продукт8 = ColumnByValue(Workbooks(ReportName_String).Name, SheetName_String, "НСЖ_МАСС (Премия, тыс.руб.)", 100, 100)
        Продукт8_Product_Code = "НСЖ_Премия"
        Продукт8_Unit = "тыс. руб."
      End If
      
      column_Продукт8_План = column_Наименование_Продукт8 ' 42
      column_Продукт8_Факт = column_Наименование_Продукт8 + 1 ' 43
      column_Продукт8_Прогноз = column_Наименование_Продукт8 + 2 ' 44
      column_Продукт8_Вып_проц = column_Наименование_Продукт8 + 3 ' 45
      column_Продукт8_Прогноз_проц = column_Наименование_Продукт8 + 4 ' 46

      ' 9. Коробочное страхование
      column_Наименование_Продукт9 = ColumnByValue(Workbooks(ReportName_String).Name, SheetName_String, "Коробочное страхование", 100, 100) ' 47
      column_Продукт9_План = column_Наименование_Продукт9 ' 47
      column_Продукт9_Факт = column_Наименование_Продукт9 + 1 ' 48
      column_Продукт9_Прогноз = column_Наименование_Продукт9 + 2 ' 49
      column_Продукт9_Вып_проц = column_Наименование_Продукт9 + 3 ' 50
      column_Продукт9_Прогноз_проц = column_Наименование_Продукт9 + 4 ' 51

      ' 10. Коробочное страхование (Антивирус + Ваша защита)
      column_Наименование_Продукт10 = 0 ' ColumnByValue(Workbooks(ReportName_String).Name, SheetName_String, "Коробочное страхование (Антивирус + Ваша защита)", 10, 100) ' 52
      column_Продукт10_План = column_Наименование_Продукт10 ' 52
      column_Продукт10_Факт = column_Наименование_Продукт10 + 1 ' 53
      column_Продукт10_Прогноз = column_Наименование_Продукт10 + 2 ' 54
      column_Продукт10_Вып_проц = column_Наименование_Продукт10 + 3 ' 55
      column_Продукт10_Прогноз_проц = column_Наименование_Продукт10 + 4 ' 56

      ' 11. Коробочное страхование: "Будьте здоровы"+"Юрист24"
      t = "Коробочное страхование: " + кавычки() + "Будьте здоровы" + кавычки() + "+" + кавычки() + "Юрист24" + кавычки()
      column_Наименование_Продукт11 = ColumnByValue(Workbooks(ReportName_String).Name, SheetName_String, "Коробочное страхование: " + кавычки() + "Будьте здоровы" + кавычки() + "+" + кавычки() + "Юрист24" + кавычки(), 100, 100)
      column_Продукт11_План = column_Наименование_Продукт11
      column_Продукт11_Факт = column_Наименование_Продукт11 + 1
      column_Продукт11_Прогноз = column_Наименование_Продукт11 + 2
      column_Продукт11_Вып_проц = column_Наименование_Продукт11 + 3
      column_Продукт11_Прогноз_проц = column_Наименование_Продукт11 + 4
      
      ' 12. Число продаж
      column_Число_продаж = ColumnByValue(Workbooks(ReportName_String).Name, SheetName_String, "Кол-во продаж сотрудника", 100, 100)
      column_Число_продаж_План = column_Число_продаж
      column_Число_продаж_Факт = column_Число_продаж + 1
      column_Число_продаж_Прогноз = column_Число_продаж + 2
      column_Число_продаж_Вып_проц = column_Число_продаж + 3
      column_Число_продаж_Прогноз_проц = column_Число_продаж + 4
      
      
      ' ***                                            ***
      ' * Инвесты - подготовка сводных таблиц            *
      StringInSheet = " ИНВЕСТ"
      SheetName_String_Инвест = FindNameSheet(ReportName_String, StringInSheet) ' "10. ИНВЕСТ"
      If SheetName_String_Инвест <> "" Then

        ' Устанавливаем период "Месяц"
        Call DB_swith_to_MonthQuarter2(ReportName_String, SheetName_String_Инвест, "Месяц")
      
        ' Находим "Тюменский ОО1"
        row_Тюменский_ОО1 = rowByValue(ReportName_String, SheetName_String_Инвест, "Тюменский ОО1", 300, 300) ' было 1000 1000
        row_ИНВЕСТ_Тюменский_ОО1 = row_Тюменский_ОО1
        column_Тюменский_ОО1 = ColumnByValue(ReportName_String, SheetName_String_Инвест, "Тюменский ОО1", 300, 300)
        
        ' В 19.08.2021 нет "Тюменский ОО1"! Проверяем
        If row_Тюменский_ОО1 <> 0 Then
          ' Открываем сводную таблицу
          Workbooks(ReportName_String).Sheets(SheetName_String_Инвест).Cells(row_Тюменский_ОО1, column_Тюменский_ОО1 + 1).ShowDetail = True
        
          ' Должен открыться Лист1
        
          ' "Факт,тыс. руб." - продажи ПИФ
          column_DB_Факт_тыс_руб = ColumnByValue(Workbooks(ReportName_String).Name, "Лист1", "Факт,тыс. руб.", 10, 100)
          ' "Брокер, шт." - число открытых Брокерских счетов
          column_DB_Брокер_шт = ColumnByValue(Workbooks(ReportName_String).Name, "Лист1", "Брокер, шт.", 10, 100)
          ' "Табельный номер" - табельный номер сотрудника
          column_DB_Табельный_номер = ColumnByValue(Workbooks(ReportName_String).Name, "Лист1", "Табельный номер", 10, 100)
          ' "DP4_отчет" - офис
          column_DB_DP4_отчет = ColumnByValue(Workbooks(ReportName_String).Name, "Лист1", "DP4_отчет", 10, 100)
        
        End If
        
      Else
        ' Если в DB Лист не найден - выводим сообщение
        MsgBox ("Не найден Лист " + Chr(34) + StringInSheet + Chr(34)) ' + " в " + ReportName_String)
      End If
      
      ' Переходим на окно DB
      ThisWorkbook.Sheets("Лист7").Activate
      
      ' *** конец: Инвесты - подготовка сводных таблиц ***


      ' Интегральный рейтинг
      column_Интегральный_рейтинг = ColumnByValue(Workbooks(ReportName_String).Name, SheetName_String, "Интегральный рейтинг", 10, 100) ' 57
            
      ' Кол-во продуктов
      column_Кол_продуктов = ColumnByValue(Workbooks(ReportName_String).Name, SheetName_String, "Кол-во продаж сотрудника", 10, 100) ' 58

      ' Обрабатываем отчет Цикл по 5-ти офисам
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

        rowCount = row_DP3_отчет
        Do While Not IsEmpty(Workbooks(ReportName_String).Sheets(SheetName_String).Cells(rowCount, column_DP3_отчет).Value)
        
          ' Если в столбце DP3_отчет есть текущий офис, то выводим сотрудника
          If InStr(Workbooks(ReportName_String).Sheets(SheetName_String).Cells(rowCount, column_DP4_отчет).Value, officeNameInReport) <> 0 Then
          
            ' Номер записи на Лист7 в списке
            recInЛист7 = recInЛист7 + 1
            ' Номер порядковый
            порядковый_номер = порядковый_номер + 1
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 1).Value = CStr(порядковый_номер) + "."
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 1).NumberFormat = "@"
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 1).HorizontalAlignment = xlCenter

            ' Номер табельный
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 2).Value = Workbooks(ReportName_String).Sheets(SheetName_String).Cells(rowCount, column_TAB_OK).Value
            ' ФИО
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 3).Value = Фамилия_и_Имя(Workbooks(ReportName_String).Sheets(SheetName_String).Cells(rowCount, column_ФИО).Value, 1)
            ' Офис
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 4).Value = officeNameInReport
            
            ' Текущие значения в месяце: Date_01 (N), Date_02 (O), Date_03 (P), Date_04 (Q), Date_05 Date_06 Date_07 Date_08 Date_09 Date_10 Date_11 Date_12 Date_13 Date_14 Date_15 Date_16 Date_17 Date_18 Date_19 Date_20 Date_21 Date_22 Date_23 Date_24 Date_25 Date_26 Date_27 Date_28 Date_29 Date_30 Date_31
            curr_Day_Month = "Date_" + Mid(dateDB, 1, 2)

            
            ' 1. Потреб кредитование (ПК)
            ' -----------------
            ThisWorkbook.Sheets("Лист7").Cells(7, 6).Value = Workbooks(ReportName_String).Sheets(SheetName_String).Cells((row_DP3_отчет - 2), column_Наименование_Продукт1).Value
            ' План
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 5).Value = Workbooks(ReportName_String).Sheets(SheetName_String).Cells(rowCount, column_Продукт1_План).Value
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 5).NumberFormat = "0"
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 5).HorizontalAlignment = xlRight
            ' Факт
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 6).Value = Workbooks(ReportName_String).Sheets(SheetName_String).Cells(rowCount, column_Продукт1_Факт).Value
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 6).NumberFormat = "0"
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 6).HorizontalAlignment = xlRight
            ' Вып_проц
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 7).Value = Workbooks(ReportName_String).Sheets(SheetName_String).Cells(rowCount, column_Продукт1_Вып_проц).Value
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 7).NumberFormat = "0%"
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 7).HorizontalAlignment = xlRight
            ' Прогноз
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 8).Value = Workbooks(ReportName_String).Sheets(SheetName_String).Cells(rowCount, column_Продукт1_Прогноз).Value
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 8).NumberFormat = "0"
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 8).HorizontalAlignment = xlRight
            ' Прогноз_проц
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 9).Value = Workbooks(ReportName_String).Sheets(SheetName_String).Cells(rowCount, column_Продукт1_Прогноз_проц).Value
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 9).NumberFormat = "0%"
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 9).HorizontalAlignment = xlRight
            ' Окраска ячейки СФЕТОФОР ' 70. Заливка ячейки цветом "светофор" - значение In_Value в %, In_Target в %
            Call Full_Color_RangeII("Лист7", recInЛист7, 9, ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 9).Value, 1)
            
            '  Идентификатор ID_Rec:
            ID_RecVar = CStr(ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 2).Value) + "-" + strMMYY(dateDB) + "-ПК"
            
            ' "Product_Code", "ПК" и "Unit", "тыс. руб."
            ' Продукт1_Product_Code = "ПК"
            ' Продукт1_Unit = "тыс. руб."

            ' Вносим данные в BASE\Sales по ПК.
            Call InsertRecordInBook("Sales", "Лист1", "ID_Rec", ID_RecVar, _
                                            "ID_Rec", ID_RecVar, _
                                              "Personnel_Number", ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 2).Value, _
                                                "Saler_Name", Workbooks(ReportName_String).Sheets(SheetName_String).Cells(rowCount, column_ФИО).Value, _
                                                  "Оffice", ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 4).Value, _
                                                    "Product_Code", Продукт1_Product_Code, _
                                                      "Update_Date", dateDB, _
                                                        "Plan", ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 5).Value, _
                                                          "Unit", Продукт1_Unit, _
                                                            "Fact", ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 6).Value, _
                                                              "Percent_Completion", ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 7).Value, _
                                                                "Prediction", ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 8).Value, _
                                                                  "Percent_Prediction", ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 9).Value, _
                                                                    "MMYY", strMMYY(dateDB), _
                                                                      curr_Day_Month, ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 6).Value, _
                                                                        "", "", _
                                                                          "", "", _
                                                                            "", "", _
                                                                              "", "", _
                                                                                "", "", _
                                                                                  "", "")


            
            
            
            ' 2. Страховки к ПК (БС)
            ' -----------------
            ThisWorkbook.Sheets("Лист7").Cells(7, 11).Value = Workbooks(ReportName_String).Sheets(SheetName_String).Cells((row_DP3_отчет - 2), column_Наименование_Продукт2).Value
            ' План
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 10).Value = Workbooks(ReportName_String).Sheets(SheetName_String).Cells(rowCount, column_Продукт2_План).Value
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 10).NumberFormat = "0"
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 10).HorizontalAlignment = xlRight
            ' Факт
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 11).Value = Workbooks(ReportName_String).Sheets(SheetName_String).Cells(rowCount, column_Продукт2_Факт).Value
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 11).NumberFormat = "0"
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 11).HorizontalAlignment = xlRight
            ' Вып_проц
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 12).Value = Workbooks(ReportName_String).Sheets(SheetName_String).Cells(rowCount, column_Продукт2_Вып_проц).Value
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 12).NumberFormat = "0%"
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 12).HorizontalAlignment = xlRight
            ' Прогноз
            If column_Продукт2_Прогноз <> 0 Then
              ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 13).Value = Workbooks(ReportName_String).Sheets(SheetName_String).Cells(rowCount, column_Продукт2_Прогноз).Value
            Else
              ' Делаем расчет
              ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 13).Value = (ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 10).Value / 100) * (Workbooks(ReportName_String).Sheets(SheetName_String).Cells(rowCount, column_Продукт2_Прогноз_проц).Value * 100)
            End If
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 13).NumberFormat = "0"
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 13).HorizontalAlignment = xlRight
            
            ' Прогноз_проц
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 14).Value = Workbooks(ReportName_String).Sheets(SheetName_String).Cells(rowCount, column_Продукт2_Прогноз_проц).Value
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 14).NumberFormat = "0%"
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 14).HorizontalAlignment = xlRight
            ' Окраска ячейки СФЕТОФОР ' 70. Заливка ячейки цветом "светофор" - значение In_Value в %, In_Target в %
            Call Full_Color_RangeII("Лист7", recInЛист7, 14, ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 14).Value, 1)

            '  Идентификатор ID_Rec:
            ID_RecVar = CStr(ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 2).Value) + "-" + strMMYY(dateDB) + "-БС"

            
            ' Вносим данные в BASE\Sales по ПК.
            Call InsertRecordInBook("Sales", "Лист1", "ID_Rec", ID_RecVar, _
                                            "ID_Rec", ID_RecVar, _
                                              "Personnel_Number", ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 2).Value, _
                                                "Saler_Name", Workbooks(ReportName_String).Sheets(SheetName_String).Cells(rowCount, column_ФИО).Value, _
                                                  "Оffice", ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 4).Value, _
                                                    "Product_Code", "БС", _
                                                      "Update_Date", dateDB, _
                                                        "Plan", ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 10).Value, _
                                                          "Unit", "тыс. руб.", _
                                                            "Fact", ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 11).Value, _
                                                              "Percent_Completion", ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 12).Value, _
                                                                "Prediction", ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 13).Value, _
                                                                  "Percent_Prediction", ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 14).Value, _
                                                                    "MMYY", strMMYY(dateDB), _
                                                                      curr_Day_Month, ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 11).Value, _
                                                                        "", "", _
                                                                          "", "", _
                                                                            "", "", _
                                                                              "", "", _
                                                                                "", "", _
                                                                                  "", "")


            ' 3. Кредитные карты (КК)
            ' ------------------
            ThisWorkbook.Sheets("Лист7").Cells(7, 16).Value = Workbooks(ReportName_String).Sheets(SheetName_String).Cells((row_DP3_отчет - 2), column_Наименование_Продукт3).Value
            ' План
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 15).Value = Workbooks(ReportName_String).Sheets(SheetName_String).Cells(rowCount, column_Продукт3_План).Value
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 15).NumberFormat = "0"
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 15).HorizontalAlignment = xlRight
            ' Факт
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 16).Value = Workbooks(ReportName_String).Sheets(SheetName_String).Cells(rowCount, column_Продукт3_Факт).Value
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 16).NumberFormat = "0"
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 16).HorizontalAlignment = xlRight
            ' Вып_проц
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 17).Value = Workbooks(ReportName_String).Sheets(SheetName_String).Cells(rowCount, column_Продукт3_Вып_проц).Value
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 17).NumberFormat = "0%"
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 17).HorizontalAlignment = xlRight
            ' Прогноз
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 18).Value = Workbooks(ReportName_String).Sheets(SheetName_String).Cells(rowCount, column_Продукт3_Прогноз).Value
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 18).NumberFormat = "0"
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 18).HorizontalAlignment = xlRight
            ' Прогноз_проц
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 19).Value = Workbooks(ReportName_String).Sheets(SheetName_String).Cells(rowCount, column_Продукт3_Прогноз_проц).Value
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 19).NumberFormat = "0%"
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 19).HorizontalAlignment = xlRight
            ' Окраска ячейки СФЕТОФОР ' 70. Заливка ячейки цветом "светофор" - значение In_Value в %, In_Target в %
            Call Full_Color_RangeII("Лист7", recInЛист7, 19, ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 19).Value, 1)

            '  Идентификатор ID_Rec:
            ID_RecVar = CStr(ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 2).Value) + "-" + strMMYY(dateDB) + "-КК"
            
            ' Вносим данные в BASE\Sales по ПК.
            Call InsertRecordInBook("Sales", "Лист1", "ID_Rec", ID_RecVar, _
                                            "ID_Rec", ID_RecVar, _
                                              "Personnel_Number", ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 2).Value, _
                                                "Saler_Name", Workbooks(ReportName_String).Sheets(SheetName_String).Cells(rowCount, column_ФИО).Value, _
                                                  "Оffice", ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 4).Value, _
                                                    "Product_Code", "КК", _
                                                      "Update_Date", dateDB, _
                                                        "Plan", ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 15).Value, _
                                                          "Unit", "шт.", _
                                                            "Fact", ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 16).Value, _
                                                              "Percent_Completion", ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 17).Value, _
                                                                "Prediction", ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 18).Value, _
                                                                  "Percent_Prediction", ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 19).Value, _
                                                                    "MMYY", strMMYY(dateDB), _
                                                                      curr_Day_Month, ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 16).Value, _
                                                                        "", "", _
                                                                          "", "", _
                                                                            "", "", _
                                                                              "", "", _
                                                                                "", "", _
                                                                                  "", "")


            ' 4. Дебетовые карты (ДК)
            ' ------------------
            ThisWorkbook.Sheets("Лист7").Cells(7, 21).Value = Workbooks(ReportName_String).Sheets(SheetName_String).Cells((row_DP3_отчет - 2), column_Наименование_Продукт4).Value
            ' План
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 20).Value = Workbooks(ReportName_String).Sheets(SheetName_String).Cells(rowCount, column_Продукт4_План).Value
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 20).NumberFormat = "0"
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 20).HorizontalAlignment = xlRight
            ' Факт
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 21).Value = Workbooks(ReportName_String).Sheets(SheetName_String).Cells(rowCount, column_Продукт4_Факт).Value
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 21).NumberFormat = "0"
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 21).HorizontalAlignment = xlRight
            ' Вып_проц
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 22).Value = Workbooks(ReportName_String).Sheets(SheetName_String).Cells(rowCount, column_Продукт4_Вып_проц).Value
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 22).NumberFormat = "0%"
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 22).HorizontalAlignment = xlRight
            ' Прогноз
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 23).Value = Workbooks(ReportName_String).Sheets(SheetName_String).Cells(rowCount, column_Продукт4_Прогноз).Value
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 23).NumberFormat = "0"
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 23).HorizontalAlignment = xlRight
            ' Прогноз_проц
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 24).Value = Workbooks(ReportName_String).Sheets(SheetName_String).Cells(rowCount, column_Продукт4_Прогноз_проц).Value
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 24).NumberFormat = "0%"
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 24).HorizontalAlignment = xlRight
            ' Окраска ячейки СФЕТОФОР ' 70. Заливка ячейки цветом "светофор" - значение In_Value в %, In_Target в %
            Call Full_Color_RangeII("Лист7", recInЛист7, 24, ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 24).Value, 1)

            '  Идентификатор ID_Rec:
            ID_RecVar = CStr(ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 2).Value) + "-" + strMMYY(dateDB) + "-ДК"
            
            ' Вносим данные в BASE\Sales по ПК.
            Call InsertRecordInBook("Sales", "Лист1", "ID_Rec", ID_RecVar, _
                                            "ID_Rec", ID_RecVar, _
                                              "Personnel_Number", ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 2).Value, _
                                                "Saler_Name", Workbooks(ReportName_String).Sheets(SheetName_String).Cells(rowCount, column_ФИО).Value, _
                                                  "Оffice", ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 4).Value, _
                                                    "Product_Code", "ДК", _
                                                      "Update_Date", dateDB, _
                                                        "Plan", ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 20).Value, _
                                                          "Unit", "шт.", _
                                                            "Fact", ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 21).Value, _
                                                              "Percent_Completion", ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 22).Value, _
                                                                "Prediction", ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 23).Value, _
                                                                  "Percent_Prediction", ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 24).Value, _
                                                                    "MMYY", strMMYY(dateDB), _
                                                                      curr_Day_Month, ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 21).Value, _
                                                                        "", "", _
                                                                          "", "", _
                                                                            "", "", _
                                                                              "", "", _
                                                                                "", "", _
                                                                                  "", "")


            ' 5. Интернет Банк (ИБ)
            ' ------------------
            ThisWorkbook.Sheets("Лист7").Cells(7, 26).Value = Workbooks(ReportName_String).Sheets(SheetName_String).Cells((row_DP3_отчет - 2), column_Наименование_Продукт5).Value
            ' План
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 25).Value = Workbooks(ReportName_String).Sheets(SheetName_String).Cells(rowCount, column_Продукт5_План).Value
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 25).NumberFormat = "0"
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 25).HorizontalAlignment = xlRight
            ' Факт
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 26).Value = Workbooks(ReportName_String).Sheets(SheetName_String).Cells(rowCount, column_Продукт5_Факт).Value
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 26).NumberFormat = "0"
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 26).HorizontalAlignment = xlRight
            ' Вып_проц
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 27).Value = Workbooks(ReportName_String).Sheets(SheetName_String).Cells(rowCount, column_Продукт5_Вып_проц).Value
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 27).NumberFormat = "0%"
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 27).HorizontalAlignment = xlRight
            ' Прогноз
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 28).Value = Workbooks(ReportName_String).Sheets(SheetName_String).Cells(rowCount, column_Продукт5_Прогноз).Value
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 28).NumberFormat = "0"
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 28).HorizontalAlignment = xlRight
            ' Прогноз_проц
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 29).Value = Workbooks(ReportName_String).Sheets(SheetName_String).Cells(rowCount, column_Продукт5_Прогноз_проц).Value
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 29).NumberFormat = "0%"
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 29).HorizontalAlignment = xlRight
            ' Окраска ячейки СФЕТОФОР ' 70. Заливка ячейки цветом "светофор" - значение In_Value в %, In_Target в %
            Call Full_Color_RangeII("Лист7", recInЛист7, 29, ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 29).Value, 1)

            '  Идентификатор ID_Rec:
            ID_RecVar = CStr(ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 2).Value) + "-" + strMMYY(dateDB) + "-ИБ"
            
            ' Вносим данные в BASE\Sales по ПК.
            Call InsertRecordInBook("Sales", "Лист1", "ID_Rec", ID_RecVar, _
                                            "ID_Rec", ID_RecVar, _
                                              "Personnel_Number", ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 2).Value, _
                                                "Saler_Name", Workbooks(ReportName_String).Sheets(SheetName_String).Cells(rowCount, column_ФИО).Value, _
                                                  "Оffice", ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 4).Value, _
                                                    "Product_Code", "ИБ", _
                                                      "Update_Date", dateDB, _
                                                        "Plan", ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 25).Value, _
                                                          "Unit", "шт.", _
                                                            "Fact", ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 26).Value, _
                                                              "Percent_Completion", ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 27).Value, _
                                                                "Prediction", ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 28).Value, _
                                                                  "Percent_Prediction", ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 29).Value, _
                                                                    "MMYY", strMMYY(dateDB), _
                                                                      curr_Day_Month, ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 26).Value, _
                                                                        "", "", _
                                                                          "", "", _
                                                                            "", "", _
                                                                              "", "", _
                                                                                "", "", _
                                                                                  "", "")

            ' 6. Портфель пассивов, тыс. руб. (ранее Накопительный счет (НС))
            ' ------------------
            If column_Наименование_Продукт6 <> 0 Then
            
            ThisWorkbook.Sheets("Лист7").Cells(7, 31).Value = Workbooks(ReportName_String).Sheets(SheetName_String).Cells((row_DP3_отчет - 2), column_Наименование_Продукт6).Value
            ' План
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 30).Value = Workbooks(ReportName_String).Sheets(SheetName_String).Cells(rowCount, column_Продукт6_План).Value
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 30).NumberFormat = "#,##0"
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 30).HorizontalAlignment = xlRight
            ' Факт
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 31).Value = Workbooks(ReportName_String).Sheets(SheetName_String).Cells(rowCount, column_Продукт6_Факт).Value
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 31).NumberFormat = "#,##0"
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 31).HorizontalAlignment = xlRight
            ' Вып_проц
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 32).Value = Workbooks(ReportName_String).Sheets(SheetName_String).Cells(rowCount, column_Продукт6_Вып_проц).Value
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 32).NumberFormat = "0%"
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 32).HorizontalAlignment = xlRight
            ' Прогноз
            If column_Продукт6_Прогноз <> 0 Then
              ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 33).Value = Workbooks(ReportName_String).Sheets(SheetName_String).Cells(rowCount, column_Продукт6_Прогноз).Value
              ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 33).NumberFormat = "0"
              ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 33).HorizontalAlignment = xlRight
            End If
            ' Прогноз_проц
            If column_Продукт6_Прогноз_проц <> 0 Then
              ' Заголовок
              ThisWorkbook.Sheets("Лист7").Cells(8, 34).Value = "Прогноз %"
              ' Значения
              ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 34).Value = Workbooks(ReportName_String).Sheets(SheetName_String).Cells(rowCount, column_Продукт6_Прогноз_проц).Value
              ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 34).NumberFormat = "0%"
              ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 34).HorizontalAlignment = xlRight
              ' Окраска ячейки СФЕТОФОР ' 70. Заливка ячейки цветом "светофор" - значение In_Value в %, In_Target в %
              Call Full_Color_RangeII("Лист7", recInЛист7, 34, ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 34).Value, 1)
            Else
              ' Заголовок
              ThisWorkbook.Sheets("Лист7").Cells(8, 34).Value = "Вып. %"
              ' Значения
              ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 34).Value = ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 32).Value
              ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 34).NumberFormat = "0%"
              ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 34).HorizontalAlignment = xlRight
              ' Окраска ячейки СФЕТОФОР ' 70. Заливка ячейки цветом "светофор" - значение In_Value в %, In_Target в %
              Call Full_Color_RangeII("Лист7", recInЛист7, 34, ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 34).Value, 1)
              ' Убираем в 32
              ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 32).Value = ""
              ThisWorkbook.Sheets("Лист7").Cells(8, 32).Value = ""
              ThisWorkbook.Sheets("Лист7").Cells(8, 33).Value = ""
            End If
    
            '  Идентификатор ID_Rec:
            ID_RecVar = CStr(ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 2).Value) + "-" + strMMYY(dateDB) + "-НС"
            
            ' Вносим данные в BASE\Sales по ПК.
            Call InsertRecordInBook("Sales", "Лист1", "ID_Rec", ID_RecVar, _
                                            "ID_Rec", ID_RecVar, _
                                              "Personnel_Number", ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 2).Value, _
                                                "Saler_Name", Workbooks(ReportName_String).Sheets(SheetName_String).Cells(rowCount, column_ФИО).Value, _
                                                  "Оffice", ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 4).Value, _
                                                    "Product_Code", "ПП", _
                                                      "Update_Date", dateDB, _
                                                        "Plan", ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 30).Value, _
                                                          "Unit", "тыс.руб.", _
                                                            "Fact", ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 31).Value, _
                                                              "Percent_Completion", ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 32).Value, _
                                                                "Prediction", ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 33).Value, _
                                                                  "Percent_Prediction", ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 34).Value, _
                                                                    "MMYY", strMMYY(dateDB), _
                                                                      curr_Day_Month, ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 31).Value, _
                                                                        "", "", _
                                                                          "", "", _
                                                                            "", "", _
                                                                              "", "", _
                                                                                "", "", _
                                                                                  "", "")
    
    
            End If ' Если продукт не найден
            
            ' 7. ИСЖ_МАСС (Премия, тыс.руб.)
            ' ------------------
            ThisWorkbook.Sheets("Лист7").Cells(7, 36).Value = Workbooks(ReportName_String).Sheets(SheetName_String).Cells((row_DP3_отчет - 2), column_Наименование_Продукт7).Value
            ' План
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 35).Value = Workbooks(ReportName_String).Sheets(SheetName_String).Cells(rowCount, column_Продукт7_План).Value
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 35).NumberFormat = "0"
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 35).HorizontalAlignment = xlRight
            ' Факт
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 36).Value = Workbooks(ReportName_String).Sheets(SheetName_String).Cells(rowCount, column_Продукт7_Факт).Value
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 36).NumberFormat = "0"
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 36).HorizontalAlignment = xlRight
            ' Вып_проц
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 37).Value = Workbooks(ReportName_String).Sheets(SheetName_String).Cells(rowCount, column_Продукт7_Вып_проц).Value
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 37).NumberFormat = "0%"
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 37).HorizontalAlignment = xlRight
            ' Прогноз
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 38).Value = Workbooks(ReportName_String).Sheets(SheetName_String).Cells(rowCount, column_Продукт7_Прогноз).Value
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 38).NumberFormat = "0"
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 38).HorizontalAlignment = xlRight
            ' Прогноз_проц
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 39).Value = Workbooks(ReportName_String).Sheets(SheetName_String).Cells(rowCount, column_Продукт7_Прогноз_проц).Value
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 39).NumberFormat = "0%"
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 39).HorizontalAlignment = xlRight
            ' Окраска ячейки СФЕТОФОР ' 70. Заливка ячейки цветом "светофор" - значение In_Value в %, In_Target в %
            Call Full_Color_RangeII("Лист7", recInЛист7, 39, ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 39).Value, 1)
    
            '  Идентификатор ID_Rec:
            ID_RecVar = CStr(ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 2).Value) + "-" + strMMYY(dateDB) + "-ИСЖ"
            
            ' Вносим данные в BASE\Sales по ПК.
            Call InsertRecordInBook("Sales", "Лист1", "ID_Rec", ID_RecVar, _
                                            "ID_Rec", ID_RecVar, _
                                              "Personnel_Number", ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 2).Value, _
                                                "Saler_Name", Workbooks(ReportName_String).Sheets(SheetName_String).Cells(rowCount, column_ФИО).Value, _
                                                  "Оffice", ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 4).Value, _
                                                    "Product_Code", "ИСЖ", _
                                                      "Update_Date", dateDB, _
                                                        "Plan", ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 35).Value, _
                                                          "Unit", "тыс. руб.", _
                                                            "Fact", ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 36).Value, _
                                                              "Percent_Completion", ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 37).Value, _
                                                                "Prediction", ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 38).Value, _
                                                                  "Percent_Prediction", ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 39).Value, _
                                                                    "MMYY", strMMYY(dateDB), _
                                                                      curr_Day_Month, ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 36).Value, _
                                                                        "", "", _
                                                                          "", "", _
                                                                            "", "", _
                                                                              "", "", _
                                                                                "", "", _
                                                                                  "", "")
    
            
            ' 8. НСЖ_МАСС (комиссионный доход) или НСЖ_МАСС (Премия, тыс.руб.)
            ' ------------------
            ThisWorkbook.Sheets("Лист7").Cells(7, 41).Value = Workbooks(ReportName_String).Sheets(SheetName_String).Cells((row_DP3_отчет - 2), column_Наименование_Продукт8).Value
            ' План
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 40).Value = Workbooks(ReportName_String).Sheets(SheetName_String).Cells(rowCount, column_Продукт8_План).Value
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 40).NumberFormat = "0"
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 40).HorizontalAlignment = xlRight
            ' Факт
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 41).Value = Workbooks(ReportName_String).Sheets(SheetName_String).Cells(rowCount, column_Продукт8_Факт).Value
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 41).NumberFormat = "0"
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 41).HorizontalAlignment = xlRight
            ' Вып_проц
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 42).Value = Workbooks(ReportName_String).Sheets(SheetName_String).Cells(rowCount, column_Продукт8_Вып_проц).Value
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 42).NumberFormat = "0%"
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 42).HorizontalAlignment = xlRight
            ' Прогноз
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 43).Value = Workbooks(ReportName_String).Sheets(SheetName_String).Cells(rowCount, column_Продукт8_Прогноз).Value
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 43).NumberFormat = "0"
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 43).HorizontalAlignment = xlRight
            ' Прогноз_проц
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 44).Value = Workbooks(ReportName_String).Sheets(SheetName_String).Cells(rowCount, column_Продукт8_Прогноз_проц).Value
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 44).NumberFormat = "0%"
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 44).HorizontalAlignment = xlRight
            ' Окраска ячейки СФЕТОФОР ' 70. Заливка ячейки цветом "светофор" - значение In_Value в %, In_Target в %
            Call Full_Color_RangeII("Лист7", recInЛист7, 44, ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 44).Value, 1)
            
            '  Идентификатор ID_Rec:
            ID_RecVar = CStr(ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 2).Value) + "-" + strMMYY(dateDB) + "-НСЖ"
            
            ' Продукт8_Product_Code
            ' Продукт8_Unit
            
            ' Вносим данные в BASE\Sales по ПК.
            Call InsertRecordInBook("Sales", "Лист1", "ID_Rec", ID_RecVar, _
                                            "ID_Rec", ID_RecVar, _
                                              "Personnel_Number", ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 2).Value, _
                                                "Saler_Name", Workbooks(ReportName_String).Sheets(SheetName_String).Cells(rowCount, column_ФИО).Value, _
                                                  "Оffice", ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 4).Value, _
                                                    "Product_Code", Продукт8_Product_Code, _
                                                      "Update_Date", dateDB, _
                                                        "Plan", ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 40).Value, _
                                                          "Unit", Продукт8_Unit, _
                                                            "Fact", ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 41).Value, _
                                                              "Percent_Completion", ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 42).Value, _
                                                                "Prediction", ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 43).Value, _
                                                                  "Percent_Prediction", ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 44).Value, _
                                                                    "MMYY", strMMYY(dateDB), _
                                                                      curr_Day_Month, ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 41).Value, _
                                                                        "", "", _
                                                                          "", "", _
                                                                            "", "", _
                                                                              "", "", _
                                                                                "", "", _
                                                                                  "", "")
            
            
            ' 9. Коробочное страхование
            ' -------------------------
            ThisWorkbook.Sheets("Лист7").Cells(7, 46).Value = Workbooks(ReportName_String).Sheets(SheetName_String).Cells((row_DP3_отчет - 2), column_Наименование_Продукт9).Value
            ' План
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 45).Value = Workbooks(ReportName_String).Sheets(SheetName_String).Cells(rowCount, column_Продукт9_План).Value
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 45).NumberFormat = "0"
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 45).HorizontalAlignment = xlRight
            ' Факт
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 46).Value = Workbooks(ReportName_String).Sheets(SheetName_String).Cells(rowCount, column_Продукт9_Факт).Value
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 46).NumberFormat = "0"
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 46).HorizontalAlignment = xlRight
            ' Вып_проц
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 47).Value = Workbooks(ReportName_String).Sheets(SheetName_String).Cells(rowCount, column_Продукт9_Вып_проц).Value
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 47).NumberFormat = "0%"
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 47).HorizontalAlignment = xlRight
            ' Прогноз
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 48).Value = Workbooks(ReportName_String).Sheets(SheetName_String).Cells(rowCount, column_Продукт9_Прогноз).Value
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 48).NumberFormat = "0"
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 48).HorizontalAlignment = xlRight
            ' Прогноз_проц
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 49).Value = Workbooks(ReportName_String).Sheets(SheetName_String).Cells(rowCount, column_Продукт9_Прогноз_проц).Value
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 49).NumberFormat = "0%"
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 49).HorizontalAlignment = xlRight
            ' Окраска ячейки СФЕТОФОР ' 70. Заливка ячейки цветом "светофор" - значение In_Value в %, In_Target в %
            Call Full_Color_RangeII("Лист7", recInЛист7, 49, ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 49).Value, 1)
            
            '  Идентификатор ID_Rec:
            ID_RecVar = CStr(ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 2).Value) + "-" + strMMYY(dateDB) + "-КС"
            
            ' Вносим данные в BASE\Sales по ПК.
            Call InsertRecordInBook("Sales", "Лист1", "ID_Rec", ID_RecVar, _
                                            "ID_Rec", ID_RecVar, _
                                              "Personnel_Number", ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 2).Value, _
                                                "Saler_Name", Workbooks(ReportName_String).Sheets(SheetName_String).Cells(rowCount, column_ФИО).Value, _
                                                  "Оffice", ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 4).Value, _
                                                    "Product_Code", "КС", _
                                                      "Update_Date", dateDB, _
                                                        "Plan", ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 45).Value, _
                                                          "Unit", "шт.", _
                                                            "Fact", ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 46).Value, _
                                                              "Percent_Completion", ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 47).Value, _
                                                                "Prediction", ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 48).Value, _
                                                                  "Percent_Prediction", ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 49).Value, _
                                                                    "MMYY", strMMYY(dateDB), _
                                                                      curr_Day_Month, ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 46).Value, _
                                                                        "", "", _
                                                                          "", "", _
                                                                            "", "", _
                                                                              "", "", _
                                                                                "", "", _
                                                                                  "", "")
            
            
            ' 10. ПИФ ' Коробочное страхование (Антивирус + Ваша защита)
            ' -------------------------
            ' If column_Наименование_Продукт10 <> 0 Then
            If row_ИНВЕСТ_Тюменский_ОО1 <> 0 Then
            
              ThisWorkbook.Sheets("Лист7").Cells(7, 51).Value = "ПИФ+ИИС ДУ (тыс. руб.)"
              ' План
              ThisWorkbook.Sheets("Лист7").Cells(8, 50).Value = "" ' ' Заголовок
              ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 50).Value = "" ' Workbooks(ReportName_String).Sheets(SheetName_String).Cells(rowCount, column_Продукт10_План).Value
              ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 50).NumberFormat = "0"
              ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 50).HorizontalAlignment = xlRight
              
              ' Считаем все проданные ПИФ-ы у сотрудника на Лист1 в DB по ТН=ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 2).Value
              Итого_продажи_ПИФ_месяц = 0
              Итого_открыто_Брок_счетов_месяц = 0
              
              rowCount_Лист1 = 2
              Do While Not IsEmpty(Workbooks(ReportName_String).Sheets("Лист1").Cells(rowCount_Лист1, column_DB_DP4_отчет).Value)
              
                ' Если это табельный номер сотрудника
                If Workbooks(ReportName_String).Sheets("Лист1").Cells(rowCount_Лист1, column_DB_Табельный_номер).Value = ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 2).Value Then
                  Итого_продажи_ПИФ_месяц = Итого_продажи_ПИФ_месяц + Workbooks(ReportName_String).Sheets("Лист1").Cells(rowCount_Лист1, column_DB_Факт_тыс_руб).Value
                  Итого_открыто_Брок_счетов_месяц = Итого_открыто_Брок_счетов_месяц + Workbooks(ReportName_String).Sheets("Лист1").Cells(rowCount_Лист1, column_DB_Брокер_шт).Value
                End If
              
                ' Следующая запись
                rowCount_Лист1 = rowCount_Лист1 + 1
                Application.StatusBar = CStr(ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 2).Value) + " расчет ПИФ..."
                DoEventsInterval (rowCount_Лист1)
          
              Loop

              
              ' Факт
              ThisWorkbook.Sheets("Лист7").Cells(8, 51).Value = "" ' ' Заголовок
              ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 51).Value = "" ' Workbooks(ReportName_String).Sheets(SheetName_String).Cells(rowCount, column_Продукт10_Факт).Value
              ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 51).NumberFormat = "0"
              ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 51).HorizontalAlignment = xlRight
              ' Вып_проц
              ThisWorkbook.Sheets("Лист7").Cells(8, 52).Value = "" ' ' Заголовок
              ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 52).Value = "" ' Workbooks(ReportName_String).Sheets(SheetName_String).Cells(rowCount, column_Продукт10_Вып_проц).Value
              ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 52).NumberFormat = "0%"
              ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 52).HorizontalAlignment = xlRight
              ' Прогноз
              ThisWorkbook.Sheets("Лист7").Cells(8, 53).Value = "" ' ' Заголовок
              ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 53).Value = "" ' Workbooks(ReportName_String).Sheets(SheetName_String).Cells(rowCount, column_Продукт10_Прогноз).Value
              ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 53).NumberFormat = "0"
              ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 53).HorizontalAlignment = xlRight
              ' Факт ' ex Прогноз_проц
              ThisWorkbook.Sheets("Лист7").Cells(8, 54).Value = "Факт" ' ' Заголовок
              ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 54).Value = Итого_продажи_ПИФ_месяц ' Workbooks(ReportName_String).Sheets(SheetName_String).Cells(rowCount, column_Продукт10_Прогноз_проц).Value
              ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 54).NumberFormat = "#,##0" ' "0%"
              ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 54).HorizontalAlignment = xlRight
              
              ' В примечание заносим Итого_открыто_Брок_счетов_месяц
              ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 63).Value = "БрСч " + ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 62).Value + CStr(Итого_открыто_Брок_счетов_месяц) + " шт."
              
              ' Окраска ячейки СФЕТОФОР ' 70. Заливка ячейки цветом "светофор" - значение In_Value в %, In_Target в %
              ' Call Full_Color_RangeII("Лист7", recInЛист7, 54, ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 54).Value, 1)
            
              '  Идентификатор ID_Rec:
              ID_RecVar = CStr(ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 2).Value) + "-" + strMMYY(dateDB) + "ПИФ"
            
              ' Вносим данные в BASE\Sales по ПК.
              Call InsertRecordInBook("Sales", "Лист1", "ID_Rec", ID_RecVar, _
                                            "ID_Rec", ID_RecVar, _
                                              "Personnel_Number", ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 2).Value, _
                                                "Saler_Name", Workbooks(ReportName_String).Sheets(SheetName_String).Cells(rowCount, column_ФИО).Value, _
                                                  "Оffice", ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 4).Value, _
                                                    "Product_Code", "ПИФ", _
                                                      "Update_Date", dateDB, _
                                                        "Plan", "", _
                                                          "Unit", "тыс.руб.", _
                                                            "Fact", Итого_продажи_ПИФ_месяц, _
                                                              "Percent_Completion", "", _
                                                                "Prediction", "", _
                                                                  "Percent_Prediction", "", _
                                                                    "MMYY", strMMYY(dateDB), _
                                                                      curr_Day_Month, Итого_продажи_ПИФ_месяц, _
                                                                        "", "", _
                                                                          "", "", _
                                                                            "", "", _
                                                                              "", "", _
                                                                                "", "", _
                                                                                  "", "")
            End If
            
            
            ' 11. Коробочное страхование: "Будьте здоровы"+"Юрист24"
            ThisWorkbook.Sheets("Лист7").Cells(7, 56).Value = Workbooks(ReportName_String).Sheets(SheetName_String).Cells((row_DP3_отчет - 2), column_Наименование_Продукт11).Value
            ' План
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 55).Value = Workbooks(ReportName_String).Sheets(SheetName_String).Cells(rowCount, column_Продукт11_План).Value
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 55).NumberFormat = "0"
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 55).HorizontalAlignment = xlRight
            ' Факт
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 56).Value = Workbooks(ReportName_String).Sheets(SheetName_String).Cells(rowCount, column_Продукт11_Факт).Value
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 56).NumberFormat = "0"
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 56).HorizontalAlignment = xlRight
            ' Вып_проц
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 57).Value = Workbooks(ReportName_String).Sheets(SheetName_String).Cells(rowCount, column_Продукт11_Вып_проц).Value
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 57).NumberFormat = "0%"
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 57).HorizontalAlignment = xlRight
            ' Прогноз
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 58).Value = Workbooks(ReportName_String).Sheets(SheetName_String).Cells(rowCount, column_Продукт11_Прогноз).Value
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 58).NumberFormat = "0"
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 58).HorizontalAlignment = xlRight
            ' Прогноз_проц
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 59).Value = Workbooks(ReportName_String).Sheets(SheetName_String).Cells(rowCount, column_Продукт11_Прогноз_проц).Value
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 59).NumberFormat = "0%"
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 59).HorizontalAlignment = xlRight
            ' Окраска ячейки СФЕТОФОР ' 70. Заливка ячейки цветом "светофор" - значение In_Value в %, In_Target в %
            Call Full_Color_RangeII("Лист7", recInЛист7, 59, ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 59).Value, 1)
            
            '  Идентификатор ID_Rec:
            ID_RecVar = CStr(ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 2).Value) + "-" + strMMYY(dateDB) + "-КС_БЗ_Ю24"
            
            ' Вносим данные в BASE\Sales по ПК.
            Call InsertRecordInBook("Sales", "Лист1", "ID_Rec", ID_RecVar, _
                                            "ID_Rec", ID_RecVar, _
                                              "Personnel_Number", ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 2).Value, _
                                                "Saler_Name", Workbooks(ReportName_String).Sheets(SheetName_String).Cells(rowCount, column_ФИО).Value, _
                                                  "Оffice", ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 4).Value, _
                                                    "Product_Code", "КС_БЗ_Ю24", _
                                                      "Update_Date", dateDB, _
                                                        "Plan", ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 55).Value, _
                                                          "Unit", "шт.", _
                                                            "Fact", ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 56).Value, _
                                                              "Percent_Completion", ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 57).Value, _
                                                                "Prediction", ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 58).Value, _
                                                                  "Percent_Prediction", ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 59).Value, _
                                                                    "MMYY", strMMYY(dateDB), _
                                                                      curr_Day_Month, ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 56).Value, _
                                                                        "", "", _
                                                                          "", "", _
                                                                            "", "", _
                                                                              "", "", _
                                                                                "", "", _
                                                                                  "", "")
            
            
            ' Интегральный рейтинг из column_Интегральный_рейтинг
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 60).Value = Workbooks(ReportName_String).Sheets(SheetName_String).Cells(rowCount, column_Интегральный_рейтинг).Value
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 60).NumberFormat = "0%"
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 60).HorizontalAlignment = xlRight
            
            ' Окраска ячейки СФЕТОФОР ' 70. Заливка ячейки цветом "светофор" - значение In_Value в %, In_Target в %
            ' Call Full_Color_RangeII("Лист7", recInЛист7, 60, ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 60).Value, 1)
            
            ' Здесь в светофоре от 85 жетлая зона
            ' Call Full_Color_RangeIII("Лист7", recInЛист7, 60, ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 60).Value, 1, 80)
            
            ' Если ИР<50%, то Красным
            If ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 60).Value < 0.5 Then
              Call Full_Color_RangeII("Лист7", recInЛист7, 60, 0, 100)
            End If
            
            ' Кол-во продуктов из column_Кол_продуктов
            If column_Кол_продуктов <> 0 Then
              
              ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 61).Value = Workbooks(ReportName_String).Sheets(SheetName_String).Cells(rowCount, column_Кол_продуктов + 1).Value
              ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 61).NumberFormat = "0"
              ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 61).HorizontalAlignment = xlRight
              
              ' Вставляем в
              '  Идентификатор ID_Rec:
              ID_RecVar = CStr(ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 2).Value) + "-" + strMMYY(dateDB) + "-Число_продаж"
            
              ' Вносим данные в BASE\Sales по ПК.
              Call InsertRecordInBook("Sales", "Лист1", "ID_Rec", ID_RecVar, _
                                            "ID_Rec", ID_RecVar, _
                                              "Personnel_Number", ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 2).Value, _
                                                "Saler_Name", Workbooks(ReportName_String).Sheets(SheetName_String).Cells(rowCount, column_ФИО).Value, _
                                                  "Оffice", ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 4).Value, _
                                                    "Product_Code", "Число_продаж", _
                                                      "Update_Date", dateDB, _
                                                        "Plan", Workbooks(ReportName_String).Sheets(SheetName_String).Cells(rowCount, column_Число_продаж_План).Value, _
                                                          "Unit", "шт.", _
                                                            "Fact", Workbooks(ReportName_String).Sheets(SheetName_String).Cells(rowCount, column_Число_продаж_Факт).Value, _
                                                              "Percent_Completion", Workbooks(ReportName_String).Sheets(SheetName_String).Cells(rowCount, column_Число_продаж_Вып_проц).Value, _
                                                                "Prediction", Workbooks(ReportName_String).Sheets(SheetName_String).Cells(rowCount, column_Число_продаж_Прогноз).Value, _
                                                                  "Percent_Prediction", Workbooks(ReportName_String).Sheets(SheetName_String).Cells(rowCount, column_Число_продаж_Прогноз_проц).Value, _
                                                                    "MMYY", strMMYY(dateDB), _
                                                                      curr_Day_Month, Workbooks(ReportName_String).Sheets(SheetName_String).Cells(rowCount, column_Число_продаж_Факт).Value, _
                                                                        "", "", _
                                                                          "", "", _
                                                                            "", "", _
                                                                              "", "", _
                                                                                "", "", _
                                                                                  "", "")
              
            End If ' Кол-во продуктов из column_Кол_продуктов

            ' Заносим в BASE\ActiveStaff текущего сотрудника по табельному номеру
            Call InsertRecordInBook("ActiveStaff", "Лист1", "ID", ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 2).Value, _
                                            "ID", ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 2).Value, _
                                              "Name", ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 3).Value, _
                                                "Office", ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 4).Value, _
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


            ' Здесь же - берем из BASE\ActiveStaff инфо об увольнении сотрудника. Прим. - ставим: уволен, отпуск, больничный
            ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 62).Value = getInfoFromActiveStaff(ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 2).Value)
                               
          End If
                  
          ' Следующая запись
          rowCount = rowCount + 1
          Application.StatusBar = officeNameInReport + ": " + CStr(rowCount) + "..."
          DoEventsInterval (rowCount)
          
        Loop
   
        ' Выводим данные по офису
      
      Next i ' Следующий офис
      
      ' Выводим итоги обработки
      
      ' Сохранение изменений
      ThisWorkbook.Save
        
      ' Закрываем BASE\Sales
      CloseBook ("Sales")
            
      ' Закрываем BASE\ActiveStaff
      CloseBook ("ActiveStaff")
    
      ' Переменная завершения обработки
      finishProcess = True
      
      Else
        ' Сообщение о неверном формате отчета или даты
        MsgBox ("Не найден лист 'Интегральный рей-г по сотруд'!")
      End If ' Поиск Наименование листа Интегральный рей-г по сотруд (все, что выше сделать +2 отступа )
      
    Else
      ' Сообщение о неверном формате отчета или даты
      MsgBox ("Проверьте отчет: " + CheckFormatReportResult + "!")
    End If ' Проверка формы отчета

    ' Закрываем файл с отчетом без сохранения изменений (параметр SaveChanges:=False)
    Workbooks(Dir(FileName)).Close SaveChanges:=False
    
    ' Переходим в ячейку M2
    ThisWorkbook.Sheets("Лист7").Range("A1").Select
  
    ' Строка статуса
    Application.StatusBar = ""

    ' Зачеркиваем пункт меню на стартовой страницы
    ' Call ЗачеркиваемТекстВячейке("Лист0", "D9")
    ' Call ЗачеркиваемТекстВячейке("Лист0", RangeByValue(ThisWorkbook.Name, "Лист0", "Оперативная справка по _________________", 100, 100))
    
    ' Итоговое сообщение
    If finishProcess = True Then
      MsgBox ("Обработка " + Dir(ReportName_String) + " завершена!")
      
      ' Если сегодня календарный понедельник, то предлагаем сформировать ЛЦО
      If Weekday(Date, vbMonday) = 1 Then

        ' Запустить формирование ЛЦО на неделю?
        If MsgBox("Запустить процесс формирования ЛЦО на неделю?", vbYesNo) = vbYes Then
          Call Сформировать_ЛЦО
        End If
        
      End If
      
    Else
      MsgBox ("Обработка отчета была прервана!")
    End If

  End If ' Если файл был выбран

End Sub


' Свернуть столбцы и оставить только прогноз, %
Sub СвернутьСтолбцыЛист7()
  
    ' Columns("E:H").Select
    ' Selection.Columns.Group

    ' 1. ПК
    Columns("E:H").Select
    Selection.EntireColumn.Hidden = True
    ThisWorkbook.Sheets("Лист7").Cells(7, 9).Value = "ПК"
    
    ' 2. СЖиЗ
    Columns("J:M").Select
    Selection.EntireColumn.Hidden = True
    ThisWorkbook.Sheets("Лист7").Cells(7, 14).Value = "СЖиЗ"
        
    ' 3. Кредитные карты
    Columns("O:R").Select
    Selection.EntireColumn.Hidden = True
    ThisWorkbook.Sheets("Лист7").Cells(7, 19).Value = "КК"
    
    ' 4. Дебетовые карты
    Columns("T:W").Select
    Selection.EntireColumn.Hidden = True
    ThisWorkbook.Sheets("Лист7").Cells(7, 24).Value = "ДК"
    
    ' 5. Интернет Банк
    Columns("Y:AB").Select
    Selection.EntireColumn.Hidden = True
    ThisWorkbook.Sheets("Лист7").Cells(7, 29).Value = "ИБ"
    
    ' 6. Портфель пассивов (ранее Накопительный счет)
    Columns("AD:AG").Select
    Selection.EntireColumn.Hidden = True
    ThisWorkbook.Sheets("Лист7").Cells(7, 34).Value = "ПП" ' "НС"
    
    ' 7. ИСЖ_МАСС (Премия, тыс.руб.)
    Columns("AI:AL").Select
    Selection.EntireColumn.Hidden = True
    ThisWorkbook.Sheets("Лист7").Cells(7, 39).Value = "ИСЖ"
    
    ' 8. НСЖ_МАСС (комиссионный доход)
    Columns("AN:AQ").Select
    Selection.EntireColumn.Hidden = True
    ThisWorkbook.Sheets("Лист7").Cells(7, 44).Value = "НСЖ"
    
    ' 9. Коробочное страхование
    Columns("AS:AV").Select
    Selection.EntireColumn.Hidden = True
    ThisWorkbook.Sheets("Лист7").Cells(7, 49).Value = "КС"
    
    ' 10. Паевые инвестиционные фонды
    Columns("AX:BA").Select
    Selection.EntireColumn.Hidden = True
    ThisWorkbook.Sheets("Лист7").Cells(7, 54).Value = "ПИФ" ' "КС(А+B)"
    
    ' 11. Коробочное страхование: "Будьте здоровы"+"Юрист24"
    Columns("BC:BF").Select
    Selection.EntireColumn.Hidden = True
    ThisWorkbook.Sheets("Лист7").Cells(7, 59).Value = "БЗ+Ю24"

End Sub

' Развернуть столбцы - показать все
Sub РазвернутьСтолбцыЛист7()
Attribute РазвернутьСтолбцыЛист7.VB_ProcData.VB_Invoke_Func = " \n14"
    
    ' 1. ПК
    ' ThisWorkbook.Sheets("Лист7").Columns("C:BF").Select
    ThisWorkbook.Sheets("Лист7").Columns("E:H").Select
    Selection.EntireColumn.Hidden = False
    ThisWorkbook.Sheets("Лист7").Cells(7, 9).Value = ""
    
    ' 2. СЖиЗ
    Columns("J:M").Select
    Selection.EntireColumn.Hidden = False
    ThisWorkbook.Sheets("Лист7").Cells(7, 14).Value = ""
        
    ' 3. Кредитные карты
    Columns("O:R").Select
    Selection.EntireColumn.Hidden = False
    ThisWorkbook.Sheets("Лист7").Cells(7, 19).Value = ""
    
    ' 4. Дебетовые карты
    Columns("T:W").Select
    Selection.EntireColumn.Hidden = False
    ThisWorkbook.Sheets("Лист7").Cells(7, 24).Value = ""
    
    ' 5. Интернет Банк
    Columns("Y:AB").Select
    Selection.EntireColumn.Hidden = False
    ThisWorkbook.Sheets("Лист7").Cells(7, 29).Value = ""
    
    ' 6. Накопительный счет
    Columns("AD:AG").Select
    Selection.EntireColumn.Hidden = False
    ThisWorkbook.Sheets("Лист7").Cells(7, 34).Value = ""
    
    ' 7. ИСЖ_МАСС (Премия, тыс.руб.)
    Columns("AI:AL").Select
    Selection.EntireColumn.Hidden = False
    ThisWorkbook.Sheets("Лист7").Cells(7, 39).Value = ""
    
    ' 8. НСЖ_МАСС (комиссионный доход)
    Columns("AN:AQ").Select
    Selection.EntireColumn.Hidden = False
    ThisWorkbook.Sheets("Лист7").Cells(7, 44).Value = ""
    
    ' 9. Коробочное страхование
    Columns("AS:AV").Select
    Selection.EntireColumn.Hidden = False
    ThisWorkbook.Sheets("Лист7").Cells(7, 49).Value = ""
    
    ' 10. Коробочное страхование (Антивирус + Ваша защита)
    Columns("AX:BA").Select
    Selection.EntireColumn.Hidden = False
    ThisWorkbook.Sheets("Лист7").Cells(7, 54).Value = ""
     
    ' 11.
    Columns("BC:BG").Select
    Selection.EntireColumn.Hidden = False
    ThisWorkbook.Sheets("Лист7").Cells(7, 59).Value = ""
     

    ThisWorkbook.Sheets("Лист7").Range("C2").Select
    
End Sub

' Перемещение в конец
Sub Sheet7_to_rigth_Table()
  ThisWorkbook.Sheets("Лист7").Range("BJ2").Select
End Sub

' Перемещение в начало
Sub Sheet7_to_left_Table()
  ThisWorkbook.Sheets("Лист7").Range("A1").Select
End Sub

' Сформировать ИПЗ
Sub Сформировать_ИПЗ()
Dim rowCount As Integer

  ' Из B5 "Интегральный рейтинг по сотрудникам на 15.07.2020 г." берем дату
  dateDB = CDate(Mid(ThisWorkbook.Sheets("Лист7").Range("B5").Value, 40, 10))

  ' Запрос на формирование ИЗП
  If MsgBox("Сформировать ИПЗ на " + ИмяМесяцаГод(dateDB) + "?", vbYesNo) = vbYes Then
    
    CountИПЗ = 0
    
    ' ====
    rowCount = 9
    ' Do While Not IsEmpty(ThisWorkbook.Sheets("Лист7").Cells(rowCount, 3).Value)
    Do While ThisWorkbook.Sheets("Лист7").Cells(rowCount, 3).Value <> ""
            
      CountИПЗ = CountИПЗ + 1
            
      ' Открываем шаблон ИЗП
      Workbooks.Open (ThisWorkbook.Path + "\Templates\ИПЗ.xlsx")
         
      ' Переход на Лист7
      ThisWorkbook.Sheets("Лист7").Activate
         
      ' Из B5 "Интегральный рейтинг по сотрудникам на 15.07.2020 г." берем дату
      dateDB = Mid(ThisWorkbook.Sheets("Лист7").Range("B5").Value, 40, 10)
         
      ' Имя файла с ИЗП
      ' FileIPZName = "ИПЗ _ РОО Тюменский_" + ИмяМесяцаГод(dateDB) + "_" + Фамилия_и_Имя(ThisWorkbook.Sheets("Лист7").Cells(RowCount, 3).Value, 2) + ".xlsx"
      ' FileIPZName = "ИПЗ _ РОО Тюменский_" + ИмяМесяцаГод(dateDB) + "_" + Фамилия_и_Имя(ThisWorkbook.Sheets("Лист7").Cells(RowCount, 3).Value, 4) + ".xlsx"
      ' Workbooks("ИПЗ.xlsx").SaveAs FileName:=ThisWorkbook.Path + "\Out\" + FileIPZName, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
 
      ' ОТВЕТЫ НА ЧАСТО ЗАДАВАЕМЫЕ ВОПРОСЫ ПО ЕСУП: КАК НАЗЫВАТЬ ФАЙЛЫ. ВНИМАНИЕ! ИЗМЕНЕНИЕ МЕТОДОЛОГИИ! Если это общее (командное) мероприятие, то название файла должно быть следующего формата: «Наименование ИФР _ наименование ДО _ дата», например: «Протокол _ ДО Звездный_01.09.2018». Если это индивидуальная встреча /активность/мероприятие, то название файла должно быть следующего формата: «Наименование ИФР _ ФИО _ дата», например: «Карта достижений _ Иванов_ 01.09.2019 », «ЛИР_Петров_01.08.2019», «ИПР_Сидоров_01.03.2019»
      FileIPZName = "ИПЗ_" + Фамилия_и_Имя(ThisWorkbook.Sheets("Лист7").Cells(rowCount, 3).Value, 5) + "_" + CStr(dateDB) + ".xlsx"
      
      ' Проверяем - если файл есть, то удаляем его
      Call deleteFile(ThisWorkbook.Path + "\Out\" + FileIPZName)
     
      Workbooks("ИПЗ.xlsx").SaveAs FileName:=ThisWorkbook.Path + "\Out\" + FileIPZName, FileFormat:=xlOpenXMLWorkbook, createBackUp:=False
        
      '   FileCopy FileIPZName, "\\probank\DavWWWRoot\drp\DocLib1\Тюменский ОО1\Управленческие процедуры\Индивидуальные встречи\" + "ИПЗ_" + Фамилия_и_Имя(ThisWorkbook.Sheets("Лист7").Cells(RowCount, 3).Value, 5) + CStr(dateDB) + ".xlsx"

    
      ' Должность + ФИО
      Workbooks(FileIPZName).Sheets("Лист1").Range("F8").Value = getFromAddrBook2(ThisWorkbook.Sheets("Лист7").Cells(rowCount, 2).Value, 3) ' Фамилия_и_Имя(ThisWorkbook.Sheets("Лист7").Cells(RowCount, 3).Value, 2)
      ' ФИО
      Workbooks(FileIPZName).Sheets("Лист1").Range("H30").Value = Фамилия_и_Имя(ThisWorkbook.Sheets("Лист7").Cells(rowCount, 3).Value, 2)
      ' Офис
      Workbooks(FileIPZName).Sheets("Лист1").Range("F9").Value = "ОО «" + ThisWorkbook.Sheets("Лист7").Cells(rowCount, 4).Value + "»"
      ' Табельный номер
      Workbooks(FileIPZName).Sheets("Лист1").Range("F10").Value = "Табельный номер: " + CStr(ThisWorkbook.Sheets("Лист7").Cells(rowCount, 2).Value)
      ' Дата - "17" июля 2020 г.
      Workbooks(FileIPZName).Sheets("Лист1").Range("G12").Value = ДеньМесяцГод(dateDB)
      ' Текст ИПЗ
      Workbooks(FileIPZName).Sheets("Лист1").Range("A14").Value = "               С целью выполнения планов " + quarterName(dateDB) + " розничного бизнеса РОО «Тюменский» прошу принять к исполнению индивидуальное плановое задание на " + ИмяМесяцаГод(dateDB) + ":"
      
      ' Выводим планы по продуктам
      ' 1. ПК
      ' Наименование услуги Банка
      Workbooks(FileIPZName).Sheets("Лист1").Range("B18").Value = ThisWorkbook.Sheets("Лист7").Cells(7, 6).Value ' "Потребительские кредиты"
      Workbooks(FileIPZName).Sheets("Лист1").Range("B18").RowHeight = 45
      ' Показатель
      Workbooks(FileIPZName).Sheets("Лист1").Range("E18").Value = "Объем"
      ' Измер-ие
      Workbooks(FileIPZName).Sheets("Лист1").Range("F18").Value = "шт."
      ' Норматив
      Workbooks(FileIPZName).Sheets("Лист1").Range("G18").Value = ThisWorkbook.Sheets("Лист7").Cells(rowCount, 5).Value
           
      ' 2. БС
      ' Наименование услуги Банка
      Workbooks(FileIPZName).Sheets("Лист1").Range("B19").Value = ThisWorkbook.Sheets("Лист7").Cells(7, 11).Value ' "Банкострахование"
      Workbooks(FileIPZName).Sheets("Лист1").Range("B19").RowHeight = 45
      ' Показатель
      Workbooks(FileIPZName).Sheets("Лист1").Range("E19").Value = "Объем"
      ' Измер-ие
      Workbooks(FileIPZName).Sheets("Лист1").Range("F19").Value = "тыс.руб."
      ' Норматив
      Workbooks(FileIPZName).Sheets("Лист1").Range("G19").Value = ThisWorkbook.Sheets("Лист7").Cells(rowCount, 10).Value
      ' Workbooks(FileIPZName).Sheets("Лист1").Range("G19").Value = "80% от ПК"
           
      ' 3. КК
      ' Наименование услуги Банка
      Workbooks(FileIPZName).Sheets("Лист1").Range("B20").Value = ThisWorkbook.Sheets("Лист7").Cells(7, 16).Value ' "Кредитные карты (активные)"
      Workbooks(FileIPZName).Sheets("Лист1").Range("B20").RowHeight = 45
      ' Показатель
      Workbooks(FileIPZName).Sheets("Лист1").Range("E20").Value = "Объем"
      ' Измер-ие
      Workbooks(FileIPZName).Sheets("Лист1").Range("F20").Value = "шт."
      ' Норматив
      Workbooks(FileIPZName).Sheets("Лист1").Range("G20").Value = ThisWorkbook.Sheets("Лист7").Cells(rowCount, 15).Value
           
      ' 4. Дебетовые карты
      ' Наименование услуги Банка
      Workbooks(FileIPZName).Sheets("Лист1").Range("B21").Value = ThisWorkbook.Sheets("Лист7").Cells(7, 21).Value ' "Дебетовые карты (активные)"
      Workbooks(FileIPZName).Sheets("Лист1").Range("B21").RowHeight = 45
      ' Показатель
      Workbooks(FileIPZName).Sheets("Лист1").Range("E21").Value = "Объем"
      ' Измер-ие
      Workbooks(FileIPZName).Sheets("Лист1").Range("F21").Value = "шт."
      ' Норматив
      Workbooks(FileIPZName).Sheets("Лист1").Range("G21").Value = ThisWorkbook.Sheets("Лист7").Cells(rowCount, 20).Value
           
      ' 5. Интернет Банк
      ' Наименование услуги Банка
      Workbooks(FileIPZName).Sheets("Лист1").Range("B22").Value = ThisWorkbook.Sheets("Лист7").Cells(7, 26).Value ' "Интернет-банк"
      Workbooks(FileIPZName).Sheets("Лист1").Range("B22").RowHeight = 45
      ' Показатель
      Workbooks(FileIPZName).Sheets("Лист1").Range("E22").Value = "Объем"
      ' Измер-ие
      Workbooks(FileIPZName).Sheets("Лист1").Range("F22").Value = "шт."
      ' Норматив
      Workbooks(FileIPZName).Sheets("Лист1").Range("G22").Value = ThisWorkbook.Sheets("Лист7").Cells(rowCount, 25).Value
           
      ' 6. Накопительные счета
      ' Наименование услуги Банка
      Workbooks(FileIPZName).Sheets("Лист1").Range("B23").Value = ThisWorkbook.Sheets("Лист7").Cells(7, 31).Value ' "Портфель пассивов"
      Workbooks(FileIPZName).Sheets("Лист1").Range("B23").RowHeight = 45
      ' Показатель
      Workbooks(FileIPZName).Sheets("Лист1").Range("E23").Value = "Объем"
      ' Измер-ие
      Workbooks(FileIPZName).Sheets("Лист1").Range("F23").Value = "шт."
      ' Норматив
      Workbooks(FileIPZName).Sheets("Лист1").Range("G23").Value = ThisWorkbook.Sheets("Лист7").Cells(rowCount, 30).Value
           
      ' 7. ИСЖ
      ' Наименование услуги Банка
      Workbooks(FileIPZName).Sheets("Лист1").Range("B24").Value = ThisWorkbook.Sheets("Лист7").Cells(7, 36).Value ' "ИСЖ (премия)"
      Workbooks(FileIPZName).Sheets("Лист1").Range("B24").RowHeight = 45
      ' Показатель
      Workbooks(FileIPZName).Sheets("Лист1").Range("E24").Value = "Объем"
      ' Измер-ие
      Workbooks(FileIPZName).Sheets("Лист1").Range("F24").Value = "тыс.руб."
      ' Норматив
      Workbooks(FileIPZName).Sheets("Лист1").Range("G24").Value = ThisWorkbook.Sheets("Лист7").Cells(rowCount, 35).Value
           
      ' 8. НСЖ
      ' Наименование услуги Банка
      Workbooks(FileIPZName).Sheets("Лист1").Range("B25").Value = ThisWorkbook.Sheets("Лист7").Cells(7, 41).Value ' "НСЖ (премия)"
      Workbooks(FileIPZName).Sheets("Лист1").Range("B25").RowHeight = 45
      ' Показатель
      Workbooks(FileIPZName).Sheets("Лист1").Range("E25").Value = "Объем"
      ' Измер-ие
      Workbooks(FileIPZName).Sheets("Лист1").Range("F25").Value = "тыс.руб."
      ' Норматив
      Workbooks(FileIPZName).Sheets("Лист1").Range("G25").Value = ThisWorkbook.Sheets("Лист7").Cells(rowCount, 40).Value
           
      ' 9. Коробочное страхование
      ' Наименование услуги Банка
      Workbooks(FileIPZName).Sheets("Лист1").Range("B26").Value = ThisWorkbook.Sheets("Лист7").Cells(7, 46).Value ' "Коробочное страхование, включая юр.сертификат «Личный Адвокат»"
      Workbooks(FileIPZName).Sheets("Лист1").Range("B26").RowHeight = 45
      ' Показатель
      Workbooks(FileIPZName).Sheets("Лист1").Range("E26").Value = "Объем"
      ' Измер-ие
      Workbooks(FileIPZName).Sheets("Лист1").Range("F26").Value = "шт."
      ' Норматив
      Workbooks(FileIPZName).Sheets("Лист1").Range("G26").Value = ThisWorkbook.Sheets("Лист7").Cells(rowCount, 45).Value
           
      ' 10. Паевые инвестиционные фонды
      ' Наименование услуги Банка
      Workbooks(FileIPZName).Sheets("Лист1").Range("B27").Value = ThisWorkbook.Sheets("Лист7").Cells(7, 51).Value ' "Паевые инвестиционные фонды"
      Workbooks(FileIPZName).Sheets("Лист1").Range("B27").RowHeight = 45
      ' Показатель
      Workbooks(FileIPZName).Sheets("Лист1").Range("E27").Value = "Объем"
      ' Измер-ие
      Workbooks(FileIPZName).Sheets("Лист1").Range("F27").Value = "тыс.руб."
      ' Норматив
      Workbooks(FileIPZName).Sheets("Лист1").Range("G27").Value = ThisWorkbook.Sheets("Лист7").Cells(rowCount, 50).Value
           
      ' 11. Коробочное страхование: "Будьте здоровы"+"Юрист24"
      ' Наименование услуги Банка
      Workbooks(FileIPZName).Sheets("Лист1").Range("B28").Value = ThisWorkbook.Sheets("Лист7").Cells(7, 56).Value ' "Паевые инвестиционные фонды"
      Workbooks(FileIPZName).Sheets("Лист1").Range("B28").RowHeight = 45
      ' Показатель
      Workbooks(FileIPZName).Sheets("Лист1").Range("E28").Value = "Объем"
      ' Измер-ие
      Workbooks(FileIPZName).Sheets("Лист1").Range("F28").Value = "шт."
      ' Норматив
      Workbooks(FileIPZName).Sheets("Лист1").Range("G28").Value = ThisWorkbook.Sheets("Лист7").Cells(rowCount, 55).Value
           
           
      ' Закрытие файла
      Workbooks(FileIPZName).Close SaveChanges:=True
    
      ' Следующая запись
      rowCount = rowCount + 1
      Application.StatusBar = CStr(CountИПЗ) + ". " + ThisWorkbook.Sheets("Лист7").Cells(rowCount, 3).Value + "..."
      DoEventsInterval (rowCount)

    Loop
    
    ' =====
    
    MsgBox ("ИПЗ в количестве " + CStr(CountИПЗ) + " шт. сформированы!")
    
    ' Перенести файл протокола в каталог ЕСУП? - https://www.excel-vba.ru/chto-umeet-excel/kak-sredstvami-vba-pereimenovatperemestitskopirovat-fajl/
    If MsgBox("Скопировать файлы ИПЗ сотрудников в каталог ЕСУП (Индивидуальные встречи)?", vbYesNo) = vbYes Then
  
      ' Строка статуса
      Application.StatusBar = "Копирование в каталог ЕСУП ..."
    
      CountИПЗ = 0
    
      ' ====
      rowCount = 9
      Do While Not IsEmpty(ThisWorkbook.Sheets("Лист7").Cells(rowCount, 3).Value)
            
        CountИПЗ = CountИПЗ + 1
        
        ' Имя файла с ИПЗ
        FileIPZName = ThisWorkbook.Path + "\Out\" + "ИПЗ_" + Фамилия_и_Имя(ThisWorkbook.Sheets("Лист7").Cells(rowCount, 3).Value, 5) + "_" + CStr(dateDB) + ".xlsx"

        ' Строка статуса
        Application.StatusBar = CStr(CountИПЗ) + " Копирование " + FileIPZName + "..."
           
        ' ОТВЕТЫ НА ЧАСТО ЗАДАВАЕМЫЕ ВОПРОСЫ ПО ЕСУП: КАК НАЗЫВАТЬ ФАЙЛЫ. ВНИМАНИЕ! ИЗМЕНЕНИЕ МЕТОДОЛОГИИ! Если это общее (командное) мероприятие, то название файла должно быть следующего формата: «Наименование ИФР _ наименование ДО _ дата», например: «Протокол _ ДО Звездный_01.09.2018». Если это индивидуальная встреча /активность/мероприятие, то название файла должно быть следующего формата: «Наименование ИФР _ ФИО _ дата», например: «Карта достижений _ Иванов_ 01.09.2019 », «ЛИР_Петров_01.08.2019», «ИПР_Сидоров_01.03.2019»
        FileCopy FileIPZName, "\\probank\DavWWWRoot\drp\DocLib1\Тюменский ОО1\Управленческие процедуры\Индивидуальные встречи\" + "ИПЗ_" + Фамилия_и_Имя(ThisWorkbook.Sheets("Лист7").Cells(rowCount, 3).Value, 5) + "_" + CStr(dateDB) + ".xlsx"
   
        Application.StatusBar = "Скопировано!"
      
        ' Следующая запись
        rowCount = rowCount + 1
        DoEventsInterval (rowCount)
  
      Loop
  
      ' Отправка ИПЗ
      Call Отправка_Lotus_Notes_Лист7_ИПЗ
  
      ' Строка статуса
      Application.StatusBar = ""

      ' Сообщение
      MsgBox ("ИПЗ в количестве " + CStr(CountИПЗ) + " шт. перенесены в каталог ЕСУП!")

    End If ' Перенос в ЕСУП

    
  End If ' Запрос на формирование

  
End Sub

' Сформировать ИПР
Sub Сформировать_ИПР()
Dim rowCount, CountИПР As Integer

  ' Из B5 "Интегральный рейтинг по сотрудникам на 15.07.2020 г." берем дату
  dateDB = CDate(Mid(ThisWorkbook.Sheets("Лист7").Range("B5").Value, 40, 10))

  ' Запрос на формирование ИЗП
  If MsgBox("Сформировать Индивидуальный план развития (ИПР) на " + ИмяМесяцаГод(dateDB) + "?", vbYesNo) = vbYes Then
    
    ' Число сформированных ИПР
    CountИПР = 0
    
    ' Открываем BASE\ProductCode
    OpenBookInBase ("ProductCode")

    ' Открываем BASE\Sales
    OpenBookInBase ("Sales")

    ' ====
    rowCount = 9
    Do While Not IsEmpty(ThisWorkbook.Sheets("Лист7").Cells(rowCount, 3).Value)
                        
      ' Открываем шаблон ИПР
      Workbooks.Open (ThisWorkbook.Path + "\Templates\ИПР.xlsx")
         
      ' Открываем шаблон файла продаж
      Workbooks.Open (ThisWorkbook.Path + "\Templates\Продажи.xlsx")
                  
      ' Переход на Лист7
      ThisWorkbook.Sheets("Лист7").Activate
         
      ' Из B5 "Интегральный рейтинг по сотрудникам на 15.07.2020 г." берем дату
      dateDB = Mid(ThisWorkbook.Sheets("Лист7").Range("B5").Value, 40, 10)
         
      ' Если сотрудник действующий
      If ThisWorkbook.Sheets("Лист7").Cells(rowCount, 62).Value = "" Then
         
        CountИПР = CountИПР + 1
         
        ' Имя файла с ИЗП
        ' ОТВЕТЫ НА ЧАСТО ЗАДАВАЕМЫЕ ВОПРОСЫ ПО ЕСУП: КАК НАЗЫВАТЬ ФАЙЛЫ. ВНИМАНИЕ! ИЗМЕНЕНИЕ МЕТОДОЛОГИИ! Если это общее (командное) мероприятие, то название файла должно быть следующего формата: «Наименование ИФР _ наименование ДО _ дата», например: «Протокол _ ДО Звездный_01.09.2018». Если это индивидуальная встреча /активность/мероприятие, то название файла должно быть следующего формата: «Наименование ИФР _ ФИО _ дата», например: «Карта достижений _ Иванов_ 01.09.2019 », «ЛИР_Петров_01.08.2019», «ИПР_Сидоров_01.03.2019»
        FileIPRName = "ИПР_" + Фамилия_и_Имя(ThisWorkbook.Sheets("Лист7").Cells(rowCount, 3).Value, 5) + "_" + CStr(dateDB) + ".xlsx"
        Workbooks("ИПР.xlsx").SaveAs FileName:=ThisWorkbook.Path + "\Out\" + FileIPRName, FileFormat:=xlOpenXMLWorkbook, createBackUp:=False
        
        ' Имя файла с продажами (.xlsx)
        FileSale = "Продажи_" + Фамилия_и_Имя(ThisWorkbook.Sheets("Лист7").Cells(rowCount, 3).Value, 5) + "_" + CStr(dateDB) + ".xlsx"
        Workbooks("Продажи.xlsx").SaveAs FileName:=ThisWorkbook.Path + "\Out\" + FileSale, FileFormat:=xlOpenXMLWorkbook, createBackUp:=False
        
        ' ФИО
        Workbooks(FileIPRName).Sheets("ИПР").Range("C2").Value = getFromAddrBook2(ThisWorkbook.Sheets("Лист7").Cells(rowCount, 2).Value, 2)
        Workbooks(FileSale).Sheets("Лист1").Range("B2").Value = Workbooks(FileIPRName).Sheets("ИПР").Range("C2").Value
        
        ' Должность
        Workbooks(FileIPRName).Sheets("ИПР").Range("C3").Value = getFromAddrBook2(ThisWorkbook.Sheets("Лист7").Cells(rowCount, 2).Value, 5)
        Workbooks(FileSale).Sheets("Лист1").Range("B3").Value = Workbooks(FileIPRName).Sheets("ИПР").Range("C3").Value
        
        ' Табельный номер
        Workbooks(FileIPRName).Sheets("ИПР").Range("C4").Value = "Табельный номер: " + CStr(ThisWorkbook.Sheets("Лист7").Cells(rowCount, 2).Value)
        Workbooks(FileSale).Sheets("Лист1").Range("B4").Value = CStr(ThisWorkbook.Sheets("Лист7").Cells(rowCount, 2).Value)
        
        ' Город
        Workbooks(FileIPRName).Sheets("ИПР").Range("B5").Value = cityOfficeName(ThisWorkbook.Sheets("Лист7").Cells(rowCount, 4).Value)
        
        ' Подразделение (Офис)
        Workbooks(FileIPRName).Sheets("ИПР").Range("C6").Value = "ОО «" + ThisWorkbook.Sheets("Лист7").Cells(rowCount, 4).Value + "»"
        Workbooks(FileSale).Sheets("Лист1").Range("B1").Value = Workbooks(FileIPRName).Sheets("ИПР").Range("C6").Value
        
        ' ФИО руководителя
        '
        ' Дата - "17" июля 2020 г.
        Workbooks(FileIPRName).Sheets("ИПР").Range("G18").Value = ДеньМесяцГод(dateDB)
      
        ' Руководитель
        ' В Продажи
        Workbooks(FileSale).Sheets("Лист1").Range("B5").Value = getFromAddrBook(РуководительМРК(ThisWorkbook.Sheets("Лист7").Cells(rowCount, 4).Value), 1)
        ' В ИПР
        Workbooks(FileIPRName).Sheets("ИПР").Range("C7").Value = Workbooks(FileSale).Sheets("Лист1").Range("B5").Value
      
        ' Год
        Workbooks(FileSale).Sheets("Лист1").Range("B6").Value = Year(dateDB)
        
        ' Переходим в Книгу ProductCode и обрабатываем результаты продаж по каждому продукту за период с 0120-0620
        ' Дата начала - первый месяц года от даты DB
        beginPeriod = firstMonthYear_strMMYY(CDate(dateDB)) ' "0120" strMMYY из dateDB ДД.ММ.ГГГГ
        
        
        ' Если это декабрь, то берем 1220
        If Month(dateDB) = 12 Then
                
          endPeriod = strMMYY(dateDB)
                
        Else
          
          ' Дата окончания периода - от текущего месяца (DB) минус 1
          endPeriod = pastMonth_strMMYY(CDate(dateDB))   ' "0620"
        
        End If
        
        rowCount2 = 2
        Do While Not IsEmpty(Workbooks("ProductCode").Sheets("Лист1").Cells(rowCount2, 1).Value)
   
          ' Итоги нарастающим за год
          Итого_план_по_продукту = 0
          Итого_факт_по_продукту = 0
   
          ' Первая итерация - Переходим в Книгу Sales и выбираем продажи по этому продукту за период с "0120" по "0620". Каждый Факт=0 считаем как 1 балл.
          '-----------------------------------------------------------------------------------------------------------------------------------------------
          rowCount3 = 2
          Отсутствие_продаж = 0
          Do While Not IsEmpty(Workbooks("Sales").Sheets("Лист1").Cells(rowCount3, 1).Value)
            
         
            ' Если это продажа текущего продавца по Лист7, за заданный период и по текущему продукту из ProductCode
            ' If (Workbooks("Sales").Sheets("Лист1").Cells(rowCount3, 2).Value = ThisWorkbook.Sheets("Лист7").Cells(rowCount, 2).Value) And (CInt(Workbooks("Sales").Sheets("Лист1").Cells(rowCount3, 5).Value) >= CInt(beginPeriod)) And (CInt(endPeriod) >= CInt(Workbooks("Sales").Sheets("Лист1").Cells(rowCount3, 5).Value)) And (Workbooks("Sales").Sheets("Лист1").Cells(rowCount3, 7).Value = Workbooks("ProductCode").Sheets("Лист1").Cells(rowCount2, 1).Value) Then
            
            If (Workbooks("Sales").Sheets("Лист1").Cells(rowCount3, 2).Value = ThisWorkbook.Sheets("Лист7").Cells(rowCount, 2).Value) And _
                 (dateBeginFromStrMMYY(Workbooks("Sales").Sheets("Лист1").Cells(rowCount3, 5).Value) >= dateBeginFromStrMMYY(beginPeriod)) And _
                   (dateEndFromStrMMYY(endPeriod) >= dateEndFromStrMMYY(Workbooks("Sales").Sheets("Лист1").Cells(rowCount3, 5).Value)) And _
                     (Workbooks("Sales").Sheets("Лист1").Cells(rowCount3, 7).Value = Workbooks("ProductCode").Sheets("Лист1").Cells(rowCount2, 1).Value) Then
                            
              ' Если в текущем месяце не было продаж вообще (10-ый столбец Fact = 0)
              If CDbl(Workbooks("Sales").Sheets("Лист1").Cells(rowCount3, 10).Value) = 0 Then
                Отсутствие_продаж = Отсутствие_продаж + 1
              End If
              
              ' Заносим в файл Продажи:
              
              ' План
              Workbooks(FileSale).Sheets("Лист1").Cells(decodeMMYY(Workbooks("Sales").Sheets("Лист1").Cells(rowCount3, 5).Value) + 8, rowCount2_toColumn(rowCount2)).Value = Workbooks("Sales").Sheets("Лист1").Cells(rowCount3, 8).Value
              Итого_план_по_продукту = Итого_план_по_продукту + Workbooks("Sales").Sheets("Лист1").Cells(rowCount3, 8).Value
              
              ' Факт
              Workbooks(FileSale).Sheets("Лист1").Cells(decodeMMYY(Workbooks("Sales").Sheets("Лист1").Cells(rowCount3, 5).Value) + 8, rowCount2_toColumn(rowCount2) + 1).Value = Workbooks("Sales").Sheets("Лист1").Cells(rowCount3, 10).Value
              Итого_факт_по_продукту = Итого_факт_по_продукту + Workbooks("Sales").Sheets("Лист1").Cells(rowCount3, 10).Value
              
              ' % исп.
              Workbooks(FileSale).Sheets("Лист1").Cells(decodeMMYY(Workbooks("Sales").Sheets("Лист1").Cells(rowCount3, 5).Value) + 8, rowCount2_toColumn(rowCount2) + 2).Value = Workbooks("Sales").Sheets("Лист1").Cells(rowCount3, 11).Value
              
              ' Окраска ячейки СФЕТОФОР ' 70. Заливка ячейки цветом "светофор" - значение In_Value в %, In_Target в %
              Call Full_Color_RangeIV(FileSale, "Лист1", decodeMMYY(Workbooks("Sales").Sheets("Лист1").Cells(rowCount3, 5).Value) + 8, rowCount2_toColumn(rowCount2) + 2, Workbooks(FileSale).Sheets("Лист1").Cells(decodeMMYY(Workbooks("Sales").Sheets("Лист1").Cells(rowCount3, 5).Value) + 8, rowCount2_toColumn(rowCount2) + 1).Value, Workbooks(FileSale).Sheets("Лист1").Cells(decodeMMYY(Workbooks("Sales").Sheets("Лист1").Cells(rowCount3, 5).Value) + 8, rowCount2_toColumn(rowCount2)).Value, 90)

              ' Отработано Working_days из Norm_days
              Workbooks(FileSale).Sheets("Лист1").Cells(decodeMMYY(Workbooks("Sales").Sheets("Лист1").Cells(rowCount3, 5).Value) + 8, 32).Value = "Отработано " + CStr(Workbooks("Sales").Sheets("Лист1").Cells(rowCount3, 15).Value) + " из " + CStr(Workbooks("Sales").Sheets("Лист1").Cells(rowCount3, 17).Value) + " дней"
              
            End If
            
            ' Следующая запись в Sales
            rowCount3 = rowCount3 + 1
            Application.StatusBar = CStr(ThisWorkbook.Sheets("Лист7").Cells(rowCount, 1).Value) + ". " + ThisWorkbook.Sheets("Лист7").Cells(rowCount, 3).Value + " " + CStr(rowCount2) + "-" + CStr(rowCount3) + "..."
            DoEventsInterval (rowCount3)

          Loop
            
          ' Обработка продаж по продавцу по текущему продукту завершена
          ' Заносим итоговый балл в ProductCode.No_Sales_tmp
          Workbooks("ProductCode").Sheets("Лист1").Cells(rowCount2, 6).Value = Отсутствие_продаж
          Workbooks("ProductCode").Sheets("Лист1").Cells(rowCount2, 7).Value = ThisWorkbook.Sheets("Лист7").Cells(rowCount, 3).Value
          
          ' Итоги по продукту в книгу "Продажи" (файл Продажи)
          ' Итого План
          Workbooks(FileSale).Sheets("Лист1").Cells(21, rowCount2_toColumn(rowCount2)).Value = Итого_план_по_продукту
          ' Итого Факт
          Workbooks(FileSale).Sheets("Лист1").Cells(21, rowCount2_toColumn(rowCount2) + 1).Value = Итого_факт_по_продукту
          ' % исп.
          Workbooks(FileSale).Sheets("Лист1").Cells(21, rowCount2_toColumn(rowCount2) + 2).Value = РассчетДоли(Итого_план_по_продукту, Итого_факт_по_продукту, 2)
          ' Окраска ячейки СФЕТОФОР ' 70. Заливка ячейки цветом "светофор" - значение In_Value в %, In_Target в %
          Call Full_Color_RangeIV(FileSale, "Лист1", 21, rowCount2_toColumn(rowCount2) + 2, Итого_факт_по_продукту, Итого_план_по_продукту, 90)
          
          
          ' Следующая запись в ProductCode
          rowCount2 = rowCount2 + 1
          ' Application.StatusBar = CStr(CountИПР) + ". " + ThisWorkbook.Sheets("Лист7").Cells(rowCount, 3).Value + "..."
          DoEventsInterval (rowCount2)

        Loop
   
        
   
        ' Вторая итерация - анализируем ProductCode и выбираем 3 худших результата и вносим их в ИПР
        '-------------------------------------------------------------------------------------------
        add_to_course = ""
        Product = ""
        Full_Product_Name = ""
        PSB_University_course = ""
        PSB_University_URL = ""
        
        For i = 1 To 3
        
          rowCount2 = 2
          max_0 = 0
          curr_product = ""
          Do While Not IsEmpty(Workbooks("ProductCode").Sheets("Лист1").Cells(rowCount2, 1).Value)
           
            If (CInt(Workbooks("ProductCode").Sheets("Лист1").Cells(rowCount2, 6).Value) > max_0) And (InStr(add_to_course, Workbooks("ProductCode").Sheets("Лист1").Cells(rowCount2, 1).Value) = 0) Then
              
              max_0 = CInt(Workbooks("ProductCode").Sheets("Лист1").Cells(rowCount2, 6).Value)
              curr_product = Workbooks("ProductCode").Sheets("Лист1").Cells(rowCount2, 1).Value
              Full_Product_Name = Workbooks("ProductCode").Sheets("Лист1").Cells(rowCount2, 3).Value
              PSB_University_course = Workbooks("ProductCode").Sheets("Лист1").Cells(rowCount2, 4).Value
              PSB_University_URL = Workbooks("ProductCode").Sheets("Лист1").Cells(rowCount2, 5).Value
              
            End If
           
            ' Следующая запись в ProductCode
            rowCount2 = rowCount2 + 1
            ' Application.StatusBar = CStr(CountИПР) + ". " + ThisWorkbook.Sheets("Лист7").Cells(rowCount, 3).Value + " " + CStr(rowCount2)
            Application.StatusBar = CStr(ThisWorkbook.Sheets("Лист7").Cells(rowCount, 1).Value) + ". " + ThisWorkbook.Sheets("Лист7").Cells(rowCount, 3).Value + " " + CStr(rowCount2)
            DoEventsInterval (rowCount2)

          Loop
                                      
          ' Первый максимум по продукту - выводим в план
          add_to_course = add_to_course + curr_product + " "
          
          ' Выводим в ИПР
          ' Области развития
          Workbooks(FileIPRName).Sheets("ИПР").Cells(8 + (2 * i), 5).Value = Full_Product_Name + ". Отсутствие продаж в " + CStr(max_0) + " отчетн. периодах с " + decodeMMYY2(beginPeriod) + " по " + decodeMMYY2(endPeriod)
          ' Мероприятия
          Workbooks(FileIPRName).Sheets("ИПР").Cells(8 + (2 * i), 8).Value = PSB_University_course + " " + PSB_University_URL
          ' Срок
          Workbooks(FileIPRName).Sheets("ИПР").Cells(8 + (2 * i), 10).Value = CStr(Date + 14)
          
        Next ' i
                                      
        ' Закрытие файла
        Workbooks(FileIPRName).Close SaveChanges:=True
        ' Открываем шаблон файла продаж
        Workbooks(FileSale).Close SaveChanges:=True
    
      End If ' Если сотрудник действующий
    
      ' Следующая запись
      rowCount = rowCount + 1
      ' Application.StatusBar = CStr(CountИПР) + ". " + ThisWorkbook.Sheets("Лист7").Cells(rowCount, 3).Value + "..."
      DoEventsInterval (rowCount)

    Loop
    
    ' =====
    Application.StatusBar = ""
    
    ' Закрываем BASE\ProductCode
    CloseBook ("ProductCode")

    ' Закрываем BASE\Sales
    CloseBook ("Sales")
   
    MsgBox ("ИПР в количестве " + CStr(CountИПР) + " шт. сформированы!")
    
    ' Перенести файл протокола в каталог ЕСУП? - https://www.excel-vba.ru/chto-umeet-excel/kak-sredstvami-vba-pereimenovatperemestitskopirovat-fajl/
    If MsgBox("Скопировать файлы ИПР сотрудников в каталог ЕСУП (Индивидуальные встречи)?", vbYesNo) = vbYes Then
  
      ' Строка статуса
      Application.StatusBar = "Копирование в каталог ЕСУП ..."
    
      CountИПР = 0
    
      ' ====
      rowCount = 9
      Do While Not IsEmpty(ThisWorkbook.Sheets("Лист7").Cells(rowCount, 3).Value)
            
        CountИПР = CountИПР + 1
        
        ' Имя файла с ИПР
        FileIPRName = ThisWorkbook.Path + "\Out\" + "ИПР_" + Фамилия_и_Имя(ThisWorkbook.Sheets("Лист7").Cells(rowCount, 3).Value, 5) + "_" + CStr(dateDB) + ".xlsx"

        ' Строка статуса
        Application.StatusBar = CStr(CountИПР) + " Копирование " + FileIPRName + "..."
           
        ' ОТВЕТЫ НА ЧАСТО ЗАДАВАЕМЫЕ ВОПРОСЫ ПО ЕСУП: КАК НАЗЫВАТЬ ФАЙЛЫ. ВНИМАНИЕ! ИЗМЕНЕНИЕ МЕТОДОЛОГИИ! Если это общее (командное) мероприятие, то название файла должно быть следующего формата: «Наименование ИФР _ наименование ДО _ дата», например: «Протокол _ ДО Звездный_01.09.2018». Если это индивидуальная встреча /активность/мероприятие, то название файла должно быть следующего формата: «Наименование ИФР _ ФИО _ дата», например: «Карта достижений _ Иванов_ 01.09.2019 », «ЛИР_Петров_01.08.2019», «ИПР_Сидоров_01.03.2019»
        FileCopy FileIPRName, "\\probank\DavWWWRoot\drp\DocLib1\Тюменский ОО1\Управленческие процедуры\Индивидуальные встречи\" + "ИПР_" + Фамилия_и_Имя(ThisWorkbook.Sheets("Лист7").Cells(rowCount, 3).Value, 5) + "_" + CStr(dateDB) + ".xlsx"
   
        Application.StatusBar = "Скопировано!"
      
        ' Следующая запись
        rowCount = rowCount + 1
        DoEventsInterval (rowCount)
  
      Loop
  
      ' Строка статуса
      Application.StatusBar = ""

      ' Сообщение
      MsgBox ("ИПР в количестве " + CStr(CountИПР) + " шт. перенесены в каталог ЕСУП!")

    End If ' Перенос в ЕСУП

    
  End If ' Запрос на формирование

  
End Sub


' Сформировать ЛЦО - Заполняем шаблон "Приложение 2. ЛЦО МРК"
Sub Сформировать_ЛЦО()
Dim FileLCOName As String
Dim rowCount As Integer

  ' Из B5 "Интегральный рейтинг по сотрудникам на 15.07.2020 г." берем дату
  dateDB = CDate(Mid(ThisWorkbook.Sheets("Лист7").Range("B5").Value, 40, 10))

  ' Дата начала недели
  Дата_начала_недели = weekStartDate(Date)
  ' Дата_начала_недели = weekStartDate(CDate("02.11.2020")) ' Отладка !!!
    
  ' Дата окончания недели
  Дата_окончания_недели = Дата_начала_недели + 4

  ' Определяем столбцы на Лист7
  column_Лист7_Прим = ColumnByValue(ThisWorkbook.Name, "Лист7", "Прим.", 300, 300)


  ' Запрос на формирование ИЗП
  If MsgBox("Сформировать ЛЦО с " + CStr(Дата_начала_недели) + " по " + CStr(Дата_окончания_недели) + "?", vbYesNo) = vbYes Then
    
    CountЛЦО = 0
    
    ' ====
    rowCount = 9
    ' Do While Not IsEmpty(ThisWorkbook.Sheets("Лист7").Cells(rowCount, 3).Value)
    Do While ThisWorkbook.Sheets("Лист7").Cells(rowCount, 3).Value <> ""
    
      ' Если в столбце "Прим." есть "Уволен", то не формируем ЛЦО по сотруднику
      If InStr(ThisWorkbook.Sheets("Лист7").Cells(rowCount, column_Лист7_Прим).Value, "Уволен") = 0 Then
      
      CountЛЦО = CountЛЦО + 1
            
      ' Открываем шаблон ИЗП
      Workbooks.Open (ThisWorkbook.Path + "\Templates\Приложение 2. ЛЦО МРК.xls")
      LCOSheetsName = "Лист целевых ориентиров"
         
      ' Переход на Лист7
      ThisWorkbook.Sheets("Лист7").Activate
                  
      ' Из B5 "Интегральный рейтинг по сотрудникам на 15.07.2020 г." берем дату
      ' dateDB = Mid(ThisWorkbook.Sheets("Лист7").Range("B5").Value, 40, 10)
         
      ' Имя файла с ИЗП
      ' FileLCOName = "ЛЦО _ РОО Тюменский_" + ИмяМесяцаГод(dateDB) + "_" + Фамилия_и_Имя(ThisWorkbook.Sheets("Лист7").Cells(RowCount, 3).Value, 2) + ".xlsx"
      ' FileLCOName = "ЛЦО _ РОО Тюменский_" + ИмяМесяцаГод(dateDB) + "_" + Фамилия_и_Имя(ThisWorkbook.Sheets("Лист7").Cells(RowCount, 3).Value, 4) + ".xlsx"
      ' Workbooks("ЛЦО.xlsx").SaveAs FileName:=ThisWorkbook.Path + "\Out\" + FileLCOName, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    
      
    
      ' ОТВЕТЫ НА ЧАСТО ЗАДАВАЕМЫЕ ВОПРОСЫ ПО ЕСУП: КАК НАЗЫВАТЬ ФАЙЛЫ. ВНИМАНИЕ! ИЗМЕНЕНИЕ МЕТОДОЛОГИИ! Если это общее (командное) мероприятие, то название файла должно быть следующего формата: «Наименование ИФР _ наименование ДО _ дата», например: «Протокол _ ДО Звездный_01.09.2018». Если это индивидуальная встреча /активность/мероприятие, то название файла должно быть следующего формата: «Наименование ИФР _ ФИО _ дата», например: «Карта достижений _ Иванов_ 01.09.2019 », «ЛИР_Петров_01.08.2019», «ИПР_Сидоров_01.03.2019»
      FileLCOName = "ЛЦО_" + Фамилия_и_Имя(ThisWorkbook.Sheets("Лист7").Cells(rowCount, 3).Value, 5) + "_" + CStr(dateDB) + ".xls"
      
      ' Проверяем - если файл есть, то удаляем его
      Call deleteFile(ThisWorkbook.Path + "\Out\" + FileLCOName)
      
      ' Workbooks("Приложение 2. ЛЦО МРК.xls").SaveAs FileName:=ThisWorkbook.Path + "\Out\" + FileLCOName, FileFormat:=xlOpenXMLWorkbook, createBackUp:=False
      Workbooks("Приложение 2. ЛЦО МРК.xls").SaveAs FileName:=ThisWorkbook.Path + "\Out\" + FileLCOName, createBackUp:=False
        
      '   FileCopy FileLCOName, "\\probank\DavWWWRoot\drp\DocLib1\Тюменский ОО1\Управленческие процедуры\Индивидуальные встречи\" + "ЛЦО_" + Фамилия_и_Имя(ThisWorkbook.Sheets("Лист7").Cells(RowCount, 3).Value, 5) + CStr(dateDB) + ".xlsx"

    
      ' Должность + ФИО
      ' Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E1").Value = getFromAddrBook2(ThisWorkbook.Sheets("Лист7").Cells(rowCount, 2).Value, 3) ' Фамилия_и_Имя(ThisWorkbook.Sheets("Лист7").Cells(RowCount, 3).Value, 2)
      ' ФИО
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D1").Value = "ФИО сотрудника: " + Фамилия_и_Имя(ThisWorkbook.Sheets("Лист7").Cells(rowCount, 3).Value, 2)
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D1").Font.Name = "Calibri"
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D1").Font.Size = 18

      ' Офис
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("G1").Value = "Офис: " + "ОО «" + ThisWorkbook.Sheets("Лист7").Cells(rowCount, 4).Value + "»"
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("G1").Font.Name = "Calibri"
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("G1").Font.Size = 18
      
      ' Табельный номер
      ' Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("F10").Value = "Табельный номер: " + CStr(ThisWorkbook.Sheets("Лист7").Cells(rowCount, 2).Value)
      ' Дата - "17" июля 2020 г.
      ' Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("G12").Value = ДеньМесяцГод(dateDB)
      ' Текст ЛЦО
      ' Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("A14").Value = "               С целью выполнения планов " + quarterName(dateDB) + " розничного бизнеса РОО «Тюменский» прошу принять к исполнению индивидуальное плановое задание на " + ИмяМесяцаГод(dateDB) + ":"
      
      ' Остаток рабочих дней определяем число рабочих дней с понеделника до конца месяца Working_days_between_dates(In_DateStart, In_DateEnd, In_working_days_in_the_week) As Integer
      Остаток_рабочих_дней = Working_days_between_dates(Дата_начала_недели, Date_last_day_month(Дата_начала_недели), 5)
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("K1").Value = "Остаток рабочих дней: " + CStr(Остаток_рабочих_дней)
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("K1").Font.Name = "Calibri"
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("K1").Font.Size = 18
      
      
      ' Понедельник B16
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("B16").Value = ДеньМесяцГод(Дата_начала_недели)
      ' Вторник B30
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("B30").Value = ДеньМесяцГод(Дата_начала_недели + 1)
      ' Среда B45
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("B45").Value = ДеньМесяцГод(Дата_начала_недели + 2)
      ' Четверг B58
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("B58").Value = ДеньМесяцГод(Дата_начала_недели + 3)
      ' Пятница B71
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("B71").Value = ДеньМесяцГод(Дата_начала_недели + 4)
      
      ' Факт за месяц (нарастающим итогом)
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("K2").Value = "Факт  за месяц (нарастающий итог) на " + CStr(Дата_начала_недели)
      
      ' Выводим планы по продуктам
      ' 1. ПК
      ' Наименование услуги Банка
      ' Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D5").Value = "Потребительские кредиты, тыс.руб."
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D5").Value = "Потребительские кредиты, шт."
      Call setFontInRange(FileLCOName, LCOSheetsName, "D5", "Calibri", 12)
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("J5").Value = Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D5").Value
      Call setFontInRange(FileLCOName, LCOSheetsName, "J5", "Calibri", 12)
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D19").Value = Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D5").Value
      Call setFontInRange(FileLCOName, LCOSheetsName, "D19", "Calibri", 12)
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D33").Value = Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D5").Value
      Call setFontInRange(FileLCOName, LCOSheetsName, "D33", "Calibri", 12)
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D47").Value = Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D5").Value
      Call setFontInRange(FileLCOName, LCOSheetsName, "D47", "Calibri", 12)
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D60").Value = Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D5").Value
      Call setFontInRange(FileLCOName, LCOSheetsName, "D60", "Calibri", 12)
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D73").Value = Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D5").Value
      Call setFontInRange(FileLCOName, LCOSheetsName, "D73", "Calibri", 12)
      
      ' Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("B18").RowHeight = 15
      ' Показатель
      ' Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E18").Value = "Объем"
      ' Измер-ие
      ' Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("F18").Value = "тыс.руб."
      ' Норматив
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E5").Value = Round(ThisWorkbook.Sheets("Лист7").Cells(rowCount, 5).Value, 0)
      Call setFontInRange(FileLCOName, LCOSheetsName, "E5", "Calibri", 18)
      ' Факт
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("K5").Value = Round(ThisWorkbook.Sheets("Лист7").Cells(rowCount, 6).Value, 0)
      Call setFontInRange(FileLCOName, LCOSheetsName, "K5", "Calibri", 18)
      
      ' ПК Норматив на день - понедельник
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("F19").Value = ОкруглениеБольше((Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E5").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("K5").Value) / Остаток_рабочих_дней)
      Call setFontInRange(FileLCOName, LCOSheetsName, "F19", "Calibri", 18)
      
      ' ПК Норматив на день - вторник
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("F33").Value = ОкруглениеБольше((Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E5").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("K5").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("F19").Value) / (Остаток_рабочих_дней - 1))
      Call setFontInRange(FileLCOName, LCOSheetsName, "F33", "Calibri", 18)
      
      ' ПК Норматив на день - среда
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("F47").Value = ОкруглениеБольше((Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E5").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("K5").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("F19").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("F33").Value) / (Остаток_рабочих_дней - 2))
      Call setFontInRange(FileLCOName, LCOSheetsName, "F47", "Calibri", 18)
      
      ' ПК Норматив на день - четверг
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("F60").Value = ОкруглениеБольше((Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E5").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("K5").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("F19").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("F33").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("F47").Value) / (Остаток_рабочих_дней - 3))
      Call setFontInRange(FileLCOName, LCOSheetsName, "F60", "Calibri", 18)
      
      ' ПК Норматив на день - пятница
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("F73").Value = ОкруглениеБольше((Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E5").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("K5").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("F19").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("F33").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("F47").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("F60").Value) / (Остаток_рабочих_дней - 4))
      Call setFontInRange(FileLCOName, LCOSheetsName, "F73", "Calibri", 18)
      
      ' ПК Цель на неделю
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("G5").Value = Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("F19").Value + Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("F33").Value + Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("F47").Value + Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("F60").Value + Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("F73").Value
      Call setFontInRange(FileLCOName, LCOSheetsName, "G5", "Calibri", 18)

      
      ' 2. БС
      ' Наименование услуги Банка
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D6").Value = "Банкострахование, тыс.руб."
      Call setFontInRange(FileLCOName, LCOSheetsName, "D6", "Calibri", 12)
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("J6").Value = Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D6").Value
      Call setFontInRange(FileLCOName, LCOSheetsName, "J6", "Calibri", 12)
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D20").Value = Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D6").Value
      Call setFontInRange(FileLCOName, LCOSheetsName, "D20", "Calibri", 12)
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D34").Value = Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D6").Value
      Call setFontInRange(FileLCOName, LCOSheetsName, "D34", "Calibri", 12)
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D48").Value = Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D6").Value
      Call setFontInRange(FileLCOName, LCOSheetsName, "D48", "Calibri", 12)
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D61").Value = Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D6").Value
      Call setFontInRange(FileLCOName, LCOSheetsName, "D61", "Calibri", 12)
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D74").Value = Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D6").Value
      Call setFontInRange(FileLCOName, LCOSheetsName, "D74", "Calibri", 12)
      

      
      ' Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("B19").RowHeight = 15
      ' Показатель
      ' Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E19").Value = "Объем"
      ' Измер-ие
      ' Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("F19").Value = "тыс.руб."
      ' Норматив
      ' Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("G19").Value = ThisWorkbook.Sheets("Лист7").Cells(rowCount, 10).Value
      ' Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E6").Value = "80% от ПК"
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E6").Value = Round(ThisWorkbook.Sheets("Лист7").Cells(rowCount, 10).Value, 0)
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E6").NumberFormat = "#,##0"
      Call setFontInRange(FileLCOName, LCOSheetsName, "E6", "Calibri", 18)
      
      ' План месяц
      План_месяц_Var = ThisWorkbook.Sheets("Лист7").Cells(rowCount, 10).Value
      ' Факт месяц
      Факт_месяц_Var = Round(ThisWorkbook.Sheets("Лист7").Cells(rowCount, 11).Value, 0)
      ' Цель до конца месяца
      Сделать_до_конца_месяца = План_месяц_Var - Факт_месяц_Var
      ' Цель на день
      Цель_на_день = Round(Сделать_до_конца_месяца / Остаток_рабочих_дней, 0)
      
      ' Факт
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("K6").Value = Round(ThisWorkbook.Sheets("Лист7").Cells(rowCount, 11).Value, 0)
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("K6").NumberFormat = "#,##0"
      Call setFontInRange(FileLCOName, LCOSheetsName, "K6", "Calibri", 18)
      
      ' БС Норматив на день - понедельник
      ' Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E20").Value = Round((Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("F19").Value / 100) * 80, 0)
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E20").Value = Цель_на_день
      Call setFontInRange(FileLCOName, LCOSheetsName, "E20", "Calibri", 18)
      
      ' БС Норматив на день - вторник
      ' Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E34").Value = Round((Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("F33").Value / 100) * 80, 0)
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E34").Value = Цель_на_день
      Call setFontInRange(FileLCOName, LCOSheetsName, "E34", "Calibri", 18)
      
      ' БС Норматив на день - среда
      ' Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E48").Value = Round((Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("F47").Value / 100) * 80, 0)
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E48").Value = Цель_на_день
      Call setFontInRange(FileLCOName, LCOSheetsName, "E48", "Calibri", 18)
      
      ' БС Норматив на день - четверг
      ' Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E61").Value = Round((Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("F60").Value / 100) * 80, 0)
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E61").Value = Цель_на_день
      Call setFontInRange(FileLCOName, LCOSheetsName, "E61", "Calibri", 18)
      
      ' БС Норматив на день - пятница
      ' Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E74").Value = Round((Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("F73").Value / 100) * 80, 0)
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E74").Value = Цель_на_день
      Call setFontInRange(FileLCOName, LCOSheetsName, "E74", "Calibri", 18)
      
      ' БС Цель на неделю
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("G6").Value = Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E20").Value + Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E34").Value + Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E48").Value + Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E61").Value + Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E74").Value
      Call setFontInRange(FileLCOName, LCOSheetsName, "G6", "Calibri", 18)
     
           
      ' 3. Интернет Банк
      ' Наименование услуги Банка
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D7").Value = "Интернет-банк, шт."
      Call setFontInRange(FileLCOName, LCOSheetsName, "D7", "Calibri", 12)
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("J7").Value = Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D7").Value
      Call setFontInRange(FileLCOName, LCOSheetsName, "J7", "Calibri", 12)
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D21").Value = Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D7").Value
      Call setFontInRange(FileLCOName, LCOSheetsName, "D21", "Calibri", 12)
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D35").Value = Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D7").Value
      Call setFontInRange(FileLCOName, LCOSheetsName, "D35", "Calibri", 12)
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D49").Value = Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D7").Value
      Call setFontInRange(FileLCOName, LCOSheetsName, "D49", "Calibri", 12)
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D62").Value = Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D7").Value
      Call setFontInRange(FileLCOName, LCOSheetsName, "D62", "Calibri", 12)
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D75").Value = Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D7").Value
      Call setFontInRange(FileLCOName, LCOSheetsName, "D75", "Calibri", 12)
      
      ' Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("B22").RowHeight = 15
      ' Показатель
      ' Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E22").Value = "Объем"
      ' Измер-ие
      ' Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("F22").Value = "шт."
      ' Норматив
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E7").Value = ThisWorkbook.Sheets("Лист7").Cells(rowCount, 25).Value
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E7").NumberFormat = "#,##0"
      Call setFontInRange(FileLCOName, LCOSheetsName, "E7", "Calibri", 18)
      ' Факт
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("K7").Value = Round(ThisWorkbook.Sheets("Лист7").Cells(rowCount, 26).Value, 0)
      Call setFontInRange(FileLCOName, LCOSheetsName, "K7", "Calibri", 18)
      
      ' ИБ Норматив на день - понедельник
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E21").Value = ОкруглениеБольше((Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E7").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("K7").Value) / Остаток_рабочих_дней)
      Call setFontInRange(FileLCOName, LCOSheetsName, "E21", "Calibri", 18)
      
      ' ИБ Норматив на день - вторник
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E35").Value = ОкруглениеБольше((Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E7").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("K7").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E21").Value) / (Остаток_рабочих_дней - 1))
      Call setFontInRange(FileLCOName, LCOSheetsName, "E35", "Calibri", 18)
      
      ' ИБ Норматив на день - среда
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E49").Value = ОкруглениеБольше((Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E7").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("K7").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E21").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E35").Value) / (Остаток_рабочих_дней - 2))
      Call setFontInRange(FileLCOName, LCOSheetsName, "E49", "Calibri", 18)
      
      ' ИБ Норматив на день - четверг
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E62").Value = ОкруглениеБольше((Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E7").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("K7").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E21").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E35").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E49").Value) / (Остаток_рабочих_дней - 3))
      Call setFontInRange(FileLCOName, LCOSheetsName, "E62", "Calibri", 18)
      
      ' ИБ Норматив на день - пятница
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E75").Value = ОкруглениеБольше((Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E7").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("K7").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E21").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E35").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E49").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E62").Value) / (Остаток_рабочих_дней - 4))
      Call setFontInRange(FileLCOName, LCOSheetsName, "E75", "Calibri", 18)
           
      ' ИБ Цель на неделю
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("G7").Value = Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E21").Value + Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E35").Value + Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E49").Value + Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E62").Value + Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E75").Value
      Call setFontInRange(FileLCOName, LCOSheetsName, "G7", "Calibri", 18)
           
      ' 4. Дебетовые карты
      ' Наименование услуги Банка
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D8").Value = "Дебетовые карты (активные), шт."
      Call setFontInRange(FileLCOName, LCOSheetsName, "D8", "Calibri", 12)
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("J8").Value = Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D8").Value
      Call setFontInRange(FileLCOName, LCOSheetsName, "J8", "Calibri", 12)
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D22").Value = Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D8").Value
      Call setFontInRange(FileLCOName, LCOSheetsName, "D22", "Calibri", 12)
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D36").Value = Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D8").Value
      Call setFontInRange(FileLCOName, LCOSheetsName, "D36", "Calibri", 12)
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D50").Value = Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D8").Value
      Call setFontInRange(FileLCOName, LCOSheetsName, "D50", "Calibri", 12)
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D63").Value = Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D8").Value
      Call setFontInRange(FileLCOName, LCOSheetsName, "D63", "Calibri", 12)
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D76").Value = Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D8").Value
      Call setFontInRange(FileLCOName, LCOSheetsName, "D76", "Calibri", 12)

      ' Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("B21").RowHeight = 15
      ' Показатель
      ' Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E21").Value = "Объем"
      ' Измер-ие
      ' Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("F21").Value = "шт."
      ' Норматив
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E8").Value = ThisWorkbook.Sheets("Лист7").Cells(rowCount, 20).Value
      Call setFontInRange(FileLCOName, LCOSheetsName, "E8", "Calibri", 18)
      ' Факт
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("K8").Value = Round(ThisWorkbook.Sheets("Лист7").Cells(rowCount, 21).Value, 0)
      Call setFontInRange(FileLCOName, LCOSheetsName, "K8", "Calibri", 18)
      
      ' ДК Норматив на день - понедельник
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("F22").Value = ОкруглениеБольше((Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E8").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("K8").Value) / Остаток_рабочих_дней)
      Call setFontInRange(FileLCOName, LCOSheetsName, "F22", "Calibri", 18)
      
      ' ДК Норматив на день - вторник
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("F36").Value = ОкруглениеБольше((Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E8").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("K8").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("F22").Value) / (Остаток_рабочих_дней - 1))
      Call setFontInRange(FileLCOName, LCOSheetsName, "F36", "Calibri", 18)
      
      ' ДК Норматив на день - среда
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("F50").Value = ОкруглениеБольше((Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E8").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("K8").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("F22").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("F36").Value) / (Остаток_рабочих_дней - 2))
      Call setFontInRange(FileLCOName, LCOSheetsName, "F50", "Calibri", 18)
      
      ' ДК Норматив на день - четверг
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("F63").Value = ОкруглениеБольше((Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E8").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("K8").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("F22").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("F36").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("F50").Value) / (Остаток_рабочих_дней - 3))
      Call setFontInRange(FileLCOName, LCOSheetsName, "F63", "Calibri", 18)
      
      ' ДК Норматив на день - пятница (с учетом округлений)
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("F76").Value = ОкруглениеБольше((Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E8").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("K8").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("F22").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("F36").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("F50").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("F63").Value) / (Остаток_рабочих_дней - 4))
      Call setFontInRange(FileLCOName, LCOSheetsName, "F76", "Calibri", 18)
      
      ' ДК Цель на неделю
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("G8").Value = Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("F22").Value + Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("F36").Value + Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("F50").Value + Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("F63").Value + Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("F76").Value
      Call setFontInRange(FileLCOName, LCOSheetsName, "G8", "Calibri", 18)
      
           
      ' 5. КК
      ' Наименование услуги Банка
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D9").Value = "Кредитные карты (активные), шт."
      Call setFontInRange(FileLCOName, LCOSheetsName, "D9", "Calibri", 12)
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("J9").Value = Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D9").Value
      Call setFontInRange(FileLCOName, LCOSheetsName, "J9", "Calibri", 12)
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D23").Value = Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D9").Value
      Call setFontInRange(FileLCOName, LCOSheetsName, "D23", "Calibri", 12)
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D37").Value = Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D9").Value
      Call setFontInRange(FileLCOName, LCOSheetsName, "D37", "Calibri", 12)
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D51").Value = Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D9").Value
      Call setFontInRange(FileLCOName, LCOSheetsName, "D51", "Calibri", 12)
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D64").Value = Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D9").Value
      Call setFontInRange(FileLCOName, LCOSheetsName, "D64", "Calibri", 12)
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D77").Value = Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D9").Value
      Call setFontInRange(FileLCOName, LCOSheetsName, "D77", "Calibri", 12)
      
      ' Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("B20").RowHeight = 15
      ' Показатель
      ' Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E20").Value = "Объем"
      ' Измер-ие
      ' Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("F20").Value = "шт."
      ' Норматив
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E9").Value = ThisWorkbook.Sheets("Лист7").Cells(rowCount, 15).Value
      Call setFontInRange(FileLCOName, LCOSheetsName, "E9", "Calibri", 18)
      ' Факт
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("K9").Value = Round(ThisWorkbook.Sheets("Лист7").Cells(rowCount, 16).Value, 0)
      Call setFontInRange(FileLCOName, LCOSheetsName, "K9", "Calibri", 18)
      
      ' КК Норматив на день - понедельник
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("F23").Value = ОкруглениеБольше((Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E9").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("K9").Value) / Остаток_рабочих_дней)
      Call setFontInRange(FileLCOName, LCOSheetsName, "F23", "Calibri", 18)
      
      ' КК Норматив на день - вторник
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("F37").Value = ОкруглениеБольше((Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E9").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("K9").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("F23").Value) / (Остаток_рабочих_дней - 1))
      Call setFontInRange(FileLCOName, LCOSheetsName, "F37", "Calibri", 18)
      
      ' КК Норматив на день - среда
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("F51").Value = ОкруглениеБольше((Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E9").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("K9").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("F23").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("F37").Value) / (Остаток_рабочих_дней - 2))
      Call setFontInRange(FileLCOName, LCOSheetsName, "F51", "Calibri", 18)
      
      ' КК Норматив на день - четверг
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("F64").Value = ОкруглениеБольше((Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E9").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("K9").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("F23").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("F37").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("F51").Value) / (Остаток_рабочих_дней - 3))
      Call setFontInRange(FileLCOName, LCOSheetsName, "F64", "Calibri", 18)
      
      ' КК Норматив на день - пятница
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("F77").Value = ОкруглениеБольше((Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E9").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("K9").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("F23").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("F37").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("F51").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("F64").Value) / (Остаток_рабочих_дней - 4))
      Call setFontInRange(FileLCOName, LCOSheetsName, "F77", "Calibri", 18)
        
      ' КК Цель на неделю
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("G9").Value = Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("F23").Value + Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("F37").Value + Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("F51").Value + Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("F64").Value + Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("F77").Value
      Call setFontInRange(FileLCOName, LCOSheetsName, "G9", "Calibri", 18)
        
        
      ' 6. Накопительные счета
      ' Наименование услуги Банка
      If False Then
      
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D10").Value = "Накопительные счета, шт."
      Call setFontInRange(FileLCOName, LCOSheetsName, "D10", "Calibri", 12)
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("J10").Value = Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D10").Value
      Call setFontInRange(FileLCOName, LCOSheetsName, "J10", "Calibri", 12)
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D24").Value = Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D10").Value
      Call setFontInRange(FileLCOName, LCOSheetsName, "D24", "Calibri", 12)
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D38").Value = Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D10").Value
      Call setFontInRange(FileLCOName, LCOSheetsName, "D38", "Calibri", 12)
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D52").Value = Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D10").Value
      Call setFontInRange(FileLCOName, LCOSheetsName, "D52", "Calibri", 12)
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D65").Value = Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D10").Value
      Call setFontInRange(FileLCOName, LCOSheetsName, "D65", "Calibri", 12)
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D78").Value = Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D10").Value
      Call setFontInRange(FileLCOName, LCOSheetsName, "D78", "Calibri", 12)
      
      ' Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("B23").RowHeight = 15
      ' Показатель
      ' Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E23").Value = "Объем"
      ' Измер-ие
      ' Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("F23").Value = "шт."
      ' Норматив
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E10").Value = ThisWorkbook.Sheets("Лист7").Cells(rowCount, 30).Value
      Call setFontInRange(FileLCOName, LCOSheetsName, "E10", "Calibri", 18)
      ' Факт
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("K10").Value = Round(ThisWorkbook.Sheets("Лист7").Cells(rowCount, 31).Value, 0)
      Call setFontInRange(FileLCOName, LCOSheetsName, "K10", "Calibri", 18)
      
      ' НС Норматив на день - понедельник
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E24").Value = ОкруглениеБольше((Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E10").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("K10").Value) / Остаток_рабочих_дней)
      Call setFontInRange(FileLCOName, LCOSheetsName, "E24", "Calibri", 18)
      
      ' НС Норматив на день - вторник
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E38").Value = ОкруглениеБольше((Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E10").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("K10").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E24").Value) / (Остаток_рабочих_дней - 1))
      Call setFontInRange(FileLCOName, LCOSheetsName, "E38", "Calibri", 18)
      
      ' НС Норматив на день - среда
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E52").Value = ОкруглениеБольше((Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E10").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("K10").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E24").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E38").Value) / (Остаток_рабочих_дней - 2))
      Call setFontInRange(FileLCOName, LCOSheetsName, "E52", "Calibri", 18)
      
      ' НС Норматив на день - четверг
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E65").Value = ОкруглениеБольше((Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E10").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("K10").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E24").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E38").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E52").Value) / (Остаток_рабочих_дней - 3))
      Call setFontInRange(FileLCOName, LCOSheetsName, "E65", "Calibri", 18)
      
      ' НС Норматив на день - пятница
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E78").Value = ОкруглениеБольше((Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E10").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("K10").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E24").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E38").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E52").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E65").Value) / (Остаток_рабочих_дней - 4))
      Call setFontInRange(FileLCOName, LCOSheetsName, "E78", "Calibri", 18)
     
      ' НС Цель на неделю
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("G10").Value = Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E24").Value + Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E38").Value + Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E52").Value + Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E65").Value + Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E78").Value
      Call setFontInRange(FileLCOName, LCOSheetsName, "G10", "Calibri", 18)
      
      Else
        
        ' Очищаем в форме
        Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D10").Value = ""
        Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("J10").Value = ""
        Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D24").Value = ""
        Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D38").Value = ""
        Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D52").Value = ""
        Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D65").Value = ""
        Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D78").Value = ""
        
      End If
       
      ' 6. Портфель пассивов+АУМ, тыс. руб.
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D10").Value = "Портфель пассивов+АУМ, тыс. руб."
      Call setFontInRange(FileLCOName, LCOSheetsName, "D10", "Calibri", 12)
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("J10").Value = Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D10").Value
      Call setFontInRange(FileLCOName, LCOSheetsName, "J10", "Calibri", 12)
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D24").Value = Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D10").Value
      Call setFontInRange(FileLCOName, LCOSheetsName, "D24", "Calibri", 12)
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D38").Value = Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D10").Value
      Call setFontInRange(FileLCOName, LCOSheetsName, "D38", "Calibri", 12)
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D52").Value = Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D10").Value
      Call setFontInRange(FileLCOName, LCOSheetsName, "D52", "Calibri", 12)
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D65").Value = Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D10").Value
      Call setFontInRange(FileLCOName, LCOSheetsName, "D65", "Calibri", 12)
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D78").Value = Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D10").Value
      Call setFontInRange(FileLCOName, LCOSheetsName, "D78", "Calibri", 12)
      
      ' Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("B23").RowHeight = 15
      ' Показатель
      ' Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E23").Value = "Объем"
      ' Измер-ие
      ' Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("F23").Value = "шт."
      ' Портфель пассивов+АУМ Норматив
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E10").Value = ThisWorkbook.Sheets("Лист7").Cells(rowCount, 30).Value
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E10").NumberFormat = "#,##0"
      Call setFontInRange(FileLCOName, LCOSheetsName, "E10", "Calibri", 18)
      
      ' План месяц
      План_месяц_Var = ThisWorkbook.Sheets("Лист7").Cells(rowCount, 30).Value
      ' Факт месяц
      Факт_месяц_Var = Round(ThisWorkbook.Sheets("Лист7").Cells(rowCount, 31).Value, 0)
      ' Цель до конца месяца
      Сделать_до_конца_месяца = План_месяц_Var - Факт_месяц_Var
      ' Цель на день
      Цель_на_день = Round(Сделать_до_конца_месяца / Остаток_рабочих_дней, 0)
      
      ' Портфель пассивов+АУМ Факт
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("K10").Value = Round(ThisWorkbook.Sheets("Лист7").Cells(rowCount, 31).Value, 0)
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("K10").NumberFormat = "#,##0"
      Call setFontInRange(FileLCOName, LCOSheetsName, "K10", "Calibri", 18)
      
      ' Портфель пассивов+АУМ НС Норматив на день - понедельник
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E24").Value = Цель_на_день
      Call setFontInRange(FileLCOName, LCOSheetsName, "E24", "Calibri", 18)
      
      ' Портфель пассивов+АУМ Норматив на день - вторник
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E38").Value = Цель_на_день
      Call setFontInRange(FileLCOName, LCOSheetsName, "E38", "Calibri", 18)
      
      ' Портфель пассивов+АУМ Норматив на день - среда
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E52").Value = Цель_на_день
      Call setFontInRange(FileLCOName, LCOSheetsName, "E52", "Calibri", 18)
      
      ' Портфель пассивов+АУМ Норматив на день - четверг
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E65").Value = Цель_на_день
      Call setFontInRange(FileLCOName, LCOSheetsName, "E65", "Calibri", 18)
      
      ' Портфель пассивов+АУМ Норматив на день - пятница
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E78").Value = Цель_на_день
      Call setFontInRange(FileLCOName, LCOSheetsName, "E78", "Calibri", 18)
     
      ' Портфель пассивов+АУМ Цель на неделю
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("G10").Value = Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E24").Value + Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E38").Value + Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E52").Value + Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E65").Value + Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E78").Value
      Call setFontInRange(FileLCOName, LCOSheetsName, "G10", "Calibri", 18)
       
     
      ' 7. ИСЖ
      ' Наименование услуги Банка
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D11").Value = "ИСЖ (премия), тыс.руб."
      Call setFontInRange(FileLCOName, LCOSheetsName, "D11", "Calibri", 12)
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("J11").Value = Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D11").Value
      Call setFontInRange(FileLCOName, LCOSheetsName, "J11", "Calibri", 12)
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D25").Value = Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D11").Value
      Call setFontInRange(FileLCOName, LCOSheetsName, "D25", "Calibri", 12)
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D39").Value = Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D11").Value
      Call setFontInRange(FileLCOName, LCOSheetsName, "D39", "Calibri", 12)
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D53").Value = Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D11").Value
      Call setFontInRange(FileLCOName, LCOSheetsName, "D53", "Calibri", 12)
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D66").Value = Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D11").Value
      Call setFontInRange(FileLCOName, LCOSheetsName, "D53", "Calibri", 12)
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D79").Value = Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D11").Value
      Call setFontInRange(FileLCOName, LCOSheetsName, "D79", "Calibri", 12)
      
      ' Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("B24").RowHeight = 15
      ' Показатель
      ' Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E24").Value = "Объем"
      ' Измер-ие
      ' Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("F24").Value = "тыс.руб."
      ' Норматив
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E11").Value = ThisWorkbook.Sheets("Лист7").Cells(rowCount, 35).Value
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E11").NumberFormat = "#,##0"
      Call setFontInRange(FileLCOName, LCOSheetsName, "E11", "Calibri", 18)
      ' Факт
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("K11").Value = Round(ThisWorkbook.Sheets("Лист7").Cells(rowCount, 36).Value, 0)
      Call setFontInRange(FileLCOName, LCOSheetsName, "K11", "Calibri", 18)
      
      ' ИСЖ Норматив на день - понедельник
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E25").Value = ОкруглениеБольше((Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E11").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("K11").Value) / Остаток_рабочих_дней)
      Call setFontInRange(FileLCOName, LCOSheetsName, "E25", "Calibri", 18)
      
      ' ИСЖ Норматив на день - вторник
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E39").Value = ОкруглениеБольше((Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E11").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("K11").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E25").Value) / (Остаток_рабочих_дней - 1))
      Call setFontInRange(FileLCOName, LCOSheetsName, "E39", "Calibri", 18)
      
      ' ИСЖ Норматив на день - среда
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E53").Value = ОкруглениеБольше((Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E11").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("K11").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E25").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E39").Value) / (Остаток_рабочих_дней - 2))
      Call setFontInRange(FileLCOName, LCOSheetsName, "E53", "Calibri", 18)
      
      ' ИСЖ Норматив на день - четверг
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E66").Value = ОкруглениеБольше((Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E11").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("K11").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E25").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E39").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E53").Value) / (Остаток_рабочих_дней - 3))
      Call setFontInRange(FileLCOName, LCOSheetsName, "E66", "Calibri", 18)
      
      ' ИСЖ Норматив на день - пятница
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E79").Value = ОкруглениеБольше((Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E11").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("K11").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E25").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E39").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E53").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E66").Value) / (Остаток_рабочих_дней - 4))
      Call setFontInRange(FileLCOName, LCOSheetsName, "E79", "Calibri", 18)
      
      ' ИСЖ Цель на неделю
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("G11").Value = Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E25").Value + Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E39").Value + Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E53").Value + Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E66").Value + Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E79").Value
      Call setFontInRange(FileLCOName, LCOSheetsName, "G11", "Calibri", 18)
      
      
      ' 8. НСЖ
      ' Наименование услуги Банка
      ' Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D12").Value = "НСЖ (комиссионный доход), тыс.руб."
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D12").Value = "НСЖ (премия), тыс.руб."
      Call setFontInRange(FileLCOName, LCOSheetsName, "D12", "Calibri", 12)
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("J12").Value = Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D12").Value
      Call setFontInRange(FileLCOName, LCOSheetsName, "J12", "Calibri", 12)
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D26").Value = Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D12").Value
      Call setFontInRange(FileLCOName, LCOSheetsName, "D26", "Calibri", 12)
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D40").Value = Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D12").Value
      Call setFontInRange(FileLCOName, LCOSheetsName, "D40", "Calibri", 12)
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D54").Value = Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D12").Value
      Call setFontInRange(FileLCOName, LCOSheetsName, "D54", "Calibri", 12)
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D67").Value = Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D12").Value
      Call setFontInRange(FileLCOName, LCOSheetsName, "D67", "Calibri", 12)
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D80").Value = Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D12").Value
      Call setFontInRange(FileLCOName, LCOSheetsName, "D80", "Calibri", 12)
      
      ' Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("B25").RowHeight = 15
      ' Показатель
      ' Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E25").Value = "Объем"
      ' Измер-ие
      ' Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("F25").Value = "тыс.руб."
      ' Норматив
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E12").Value = ThisWorkbook.Sheets("Лист7").Cells(rowCount, 40).Value
      Call setFontInRange(FileLCOName, LCOSheetsName, "E12", "Calibri", 18)
      ' Факт
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("K12").Value = Round(ThisWorkbook.Sheets("Лист7").Cells(rowCount, 41).Value, 0)
      Call setFontInRange(FileLCOName, LCOSheetsName, "K12", "Calibri", 18)
      
      ' НСЖ Норматив на день - понедельник
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E26").Value = ОкруглениеБольше((Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E12").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("K12").Value) / Остаток_рабочих_дней)
      Call setFontInRange(FileLCOName, LCOSheetsName, "E26", "Calibri", 18)
      
      ' НСЖ Норматив на день - вторник
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E40").Value = ОкруглениеБольше((Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E12").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("K12").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E26").Value) / (Остаток_рабочих_дней - 1))
      Call setFontInRange(FileLCOName, LCOSheetsName, "E40", "Calibri", 18)
      
      ' НСЖ Норматив на день - среда
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E54").Value = ОкруглениеБольше((Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E12").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("K12").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E26").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E40").Value) / (Остаток_рабочих_дней - 2))
      Call setFontInRange(FileLCOName, LCOSheetsName, "E54", "Calibri", 18)
      
      ' НСЖ Норматив на день - четверг
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E67").Value = ОкруглениеБольше((Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E12").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("K12").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E26").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E40").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E54").Value) / (Остаток_рабочих_дней - 3))
      Call setFontInRange(FileLCOName, LCOSheetsName, "E67", "Calibri", 18)
      
      ' НСЖ Норматив на день - пятница
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E80").Value = ОкруглениеБольше((Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E12").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("K12").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E26").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E40").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E54").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E67").Value) / (Остаток_рабочих_дней - 4))
      Call setFontInRange(FileLCOName, LCOSheetsName, "E80", "Calibri", 18)
      
      ' НСЖ Цель на неделю
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("G12").Value = Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E26").Value + Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E40").Value + Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E54").Value + Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E67").Value + Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E67").Value
      Call setFontInRange(FileLCOName, LCOSheetsName, "G12", "Calibri", 18)
      
      
      ' 9. Коробочное страхование
      ' Наименование услуги Банка
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D13").Value = "Коробочное страхование, шт."
      Call setFontInRange(FileLCOName, LCOSheetsName, "D13", "Calibri", 12)
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("J13").Value = Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D13").Value
      Call setFontInRange(FileLCOName, LCOSheetsName, "J13", "Calibri", 12)
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D27").Value = Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D13").Value
      Call setFontInRange(FileLCOName, LCOSheetsName, "D27", "Calibri", 12)
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D41").Value = Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D13").Value
      Call setFontInRange(FileLCOName, LCOSheetsName, "D41", "Calibri", 12)
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D55").Value = Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D13").Value
      Call setFontInRange(FileLCOName, LCOSheetsName, "D55", "Calibri", 12)
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D68").Value = Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D13").Value
      Call setFontInRange(FileLCOName, LCOSheetsName, "D68", "Calibri", 12)
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D81").Value = Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D13").Value
      Call setFontInRange(FileLCOName, LCOSheetsName, "D81", "Calibri", 12)
      
      ' Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("B26").RowHeight = 45
      ' Показатель
      ' Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E26").Value = "Объем"
      ' Измер-ие
      ' Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("F26").Value = "шт."
      ' Норматив
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E13").Value = ThisWorkbook.Sheets("Лист7").Cells(rowCount, 45).Value
      Call setFontInRange(FileLCOName, LCOSheetsName, "E13", "Calibri", 18)
      ' Факт
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("K13").Value = Round(ThisWorkbook.Sheets("Лист7").Cells(rowCount, 46).Value, 0)
      Call setFontInRange(FileLCOName, LCOSheetsName, "K13", "Calibri", 18)
      
      ' Коробочное страхование Норматив на день - понедельник
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E27").Value = ОкруглениеБольше((Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E13").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("K13").Value) / Остаток_рабочих_дней)
      Call setFontInRange(FileLCOName, LCOSheetsName, "E27", "Calibri", 18)
      
      ' Коробочное страхование Норматив на день - вторник
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E41").Value = ОкруглениеБольше((Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E13").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("K13").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E27").Value) / (Остаток_рабочих_дней - 1))
      Call setFontInRange(FileLCOName, LCOSheetsName, "E41", "Calibri", 18)
      
      ' Коробочное страхование Норматив на день - среда
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E55").Value = ОкруглениеБольше((Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E13").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("K13").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E27").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E41").Value) / (Остаток_рабочих_дней - 2))
      Call setFontInRange(FileLCOName, LCOSheetsName, "E55", "Calibri", 18)
      
      ' Коробочное страхование Норматив на день - четверг
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E68").Value = ОкруглениеБольше((Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E13").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("K13").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E27").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E41").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E55").Value) / (Остаток_рабочих_дней - 3))
      Call setFontInRange(FileLCOName, LCOSheetsName, "E68", "Calibri", 18)
      
      ' Коробочное страхование Норматив на день - пятница (с учетом округлений)
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E81").Value = ОкруглениеБольше((Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E13").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("K13").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E27").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E41").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E55").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E68").Value) / (Остаток_рабочих_дней - 4))
      Call setFontInRange(FileLCOName, LCOSheetsName, "E81", "Calibri", 18)
           
      ' Коробочное страхование Цель на неделю
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("G13").Value = Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E27").Value + Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E41").Value + Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E55").Value + Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E68").Value + Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E81").Value
      Call setFontInRange(FileLCOName, LCOSheetsName, "G13", "Calibri", 18)
           
           
      ' 10. Коробочное страхование
      ' №
      ' Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("C14").Value = ""
      ' Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("C28").Value = ""
      ' Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("C42").Value = ""
      ' Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("C56").Value = ""
      ' Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("C69").Value = ""
      ' Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("C82").Value = ""
      
      ' Наименование услуги Банка
      ' Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("B27").Value = "Коробочное страхование «Антивирус», «Ваша защита»"
      ' Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("B27").RowHeight = 45
      ' Показатель
      ' Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E27").Value = "Объем"
      ' Измер-ие
      ' Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("F27").Value = "шт."
      ' Норматив
      ' Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("G27").Value = ThisWorkbook.Sheets("Лист7").Cells(rowCount, 50).Value
           
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D14").Value = "Коробки: Будьте здоровы+Юрист24"
      Call setFontInRange(FileLCOName, LCOSheetsName, "D14", "Calibri", 12)
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("J14").Value = Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D14").Value
      Call setFontInRange(FileLCOName, LCOSheetsName, "J14", "Calibri", 12)
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D28").Value = Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D14").Value
      Call setFontInRange(FileLCOName, LCOSheetsName, "D28", "Calibri", 12)
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D42").Value = Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D14").Value
      Call setFontInRange(FileLCOName, LCOSheetsName, "D42", "Calibri", 12)
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D56").Value = Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D14").Value
      Call setFontInRange(FileLCOName, LCOSheetsName, "D56", "Calibri", 12)
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D69").Value = Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D14").Value
      Call setFontInRange(FileLCOName, LCOSheetsName, "D69", "Calibri", 12)
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D82").Value = Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("D14").Value
      Call setFontInRange(FileLCOName, LCOSheetsName, "D82", "Calibri", 12)
      
      ' Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("B26").RowHeight = 45
      ' Показатель
      ' Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E26").Value = "Объем"
      ' Измер-ие
      ' Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("F26").Value = "шт."
      ' Норматив
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E14").Value = ThisWorkbook.Sheets("Лист7").Cells(rowCount, 55).Value
      Call setFontInRange(FileLCOName, LCOSheetsName, "E14", "Calibri", 18)
      ' Факт
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("K14").Value = Round(ThisWorkbook.Sheets("Лист7").Cells(rowCount, 56).Value, 0)
      Call setFontInRange(FileLCOName, LCOSheetsName, "K14", "Calibri", 18)
      
      ' Коробочное страхование Норматив на день - понедельник
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E28").Value = ОкруглениеБольше((Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E14").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("K14").Value) / Остаток_рабочих_дней)
      Call setFontInRange(FileLCOName, LCOSheetsName, "E28", "Calibri", 18)
      
      ' Коробочное страхование Норматив на день - вторник
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E42").Value = ОкруглениеБольше((Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E14").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("K14").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E28").Value) / (Остаток_рабочих_дней - 1))
      Call setFontInRange(FileLCOName, LCOSheetsName, "E42", "Calibri", 18)
      
      ' Коробочное страхование Норматив на день - среда
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E56").Value = ОкруглениеБольше((Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E14").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("K14").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E28").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E42").Value) / (Остаток_рабочих_дней - 2))
      Call setFontInRange(FileLCOName, LCOSheetsName, "E56", "Calibri", 18)
      
      ' Коробочное страхование Норматив на день - четверг
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E69").Value = ОкруглениеБольше((Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E14").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("K14").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E28").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E42").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E56").Value) / (Остаток_рабочих_дней - 3))
      Call setFontInRange(FileLCOName, LCOSheetsName, "E69", "Calibri", 18)
      
      ' Коробочное страхование Норматив на день - пятница (с учетом округлений)
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E82").Value = ОкруглениеБольше((Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E14").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("K14").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E28").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E42").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E56").Value - Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E69").Value) / (Остаток_рабочих_дней - 4))
      Call setFontInRange(FileLCOName, LCOSheetsName, "E82", "Calibri", 18)
           
      ' Коробочное страхование Цель на неделю
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("G14").Value = Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E28").Value + Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E42").Value + Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E56").Value + Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E69").Value + Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("E82").Value
      Call setFontInRange(FileLCOName, LCOSheetsName, "G14", "Calibri", 18)
           
           
      ' ===========================================================================================================================
           
      ' Перейти в ячейку
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Activate
      Workbooks(FileLCOName).Sheets(LCOSheetsName).Range("C2").Select
           
      ' Закрытие файла
      Workbooks(FileLCOName).Close SaveChanges:=True
    
      ' Переход на Лист7
      ThisWorkbook.Sheets("Лист7").Activate

      End If ' Если в "Прим" есть "Уволен"

      ' Следующая запись
      rowCount = rowCount + 1
      Application.StatusBar = CStr(CountЛЦО) + ". " + ThisWorkbook.Sheets("Лист7").Cells(rowCount, 3).Value + "..."
      DoEventsInterval (rowCount)

    Loop
    
    ' =====
    
    ' Отправка сообщения в почту
    ' Call Отправка_Лист7_ЛЦО_в_почту
    
    ' Строка статуса
    Application.StatusBar = ""
    

    ' Сообщение
    MsgBox ("ЛЦО в количестве " + CStr(CountЛЦО) + " шт. сформированы!")
    
    ' Перенести файл протокола в каталог ЕСУП? - https://www.excel-vba.ru/chto-umeet-excel/kak-sredstvami-vba-pereimenovatperemestitskopirovat-fajl/
    If MsgBox("Скопировать файлы ЛЦО сотрудников в каталог ЕСУП (Индивидуальные встречи)?", vbYesNo) = vbYes Then
  
      ' Строка статуса
      Application.StatusBar = "Копирование в каталог ЕСУП ..."
    
      CountЛЦО = 0
    
      ' ====
      rowCount = 9
      ' Do While Not IsEmpty(ThisWorkbook.Sheets("Лист7").Cells(rowCount, 3).Value)
      Do While ThisWorkbook.Sheets("Лист7").Cells(rowCount, 3).Value <> ""
            
        
        ' Если в столбце "Прим." есть "Уволен", то не формируем ЛЦО по сотруднику
        If InStr(ThisWorkbook.Sheets("Лист7").Cells(rowCount, column_Лист7_Прим).Value, "Уволен") = 0 Then

          CountЛЦО = CountЛЦО + 1
        
          ' Имя файла с ЛЦО
          FileLCOName = ThisWorkbook.Path + "\Out\" + "ЛЦО_" + Фамилия_и_Имя(ThisWorkbook.Sheets("Лист7").Cells(rowCount, 3).Value, 5) + "_" + CStr(dateDB) + ".xls"

          ' Строка статуса
          Application.StatusBar = CStr(CountЛЦО) + " Копирование " + FileLCOName + "..."
           
          ' ОТВЕТЫ НА ЧАСТО ЗАДАВАЕМЫЕ ВОПРОСЫ ПО ЕСУП: КАК НАЗЫВАТЬ ФАЙЛЫ. ВНИМАНИЕ! ИЗМЕНЕНИЕ МЕТОДОЛОГИИ! Если это общее (командное) мероприятие, то название файла должно быть следующего формата: «Наименование ИФР _ наименование ДО _ дата», например: «Протокол _ ДО Звездный_01.09.2018». Если это индивидуальная встреча /активность/мероприятие, то название файла должно быть следующего формата: «Наименование ИФР _ ФИО _ дата», например: «Карта достижений _ Иванов_ 01.09.2019 », «ЛИР_Петров_01.08.2019», «ИПР_Сидоров_01.03.2019»
          FileCopy FileLCOName, "\\probank\DavWWWRoot\drp\DocLib1\Тюменский ОО1\Управленческие процедуры\Индивидуальные встречи\" + "ЛЦО_" + Фамилия_и_Имя(ThisWorkbook.Sheets("Лист7").Cells(rowCount, 3).Value, 5) + "_" + CStr(dateDB) + ".xls"
   
          Application.StatusBar = "Скопировано!"
      
        End If
      
        ' Следующая запись
        rowCount = rowCount + 1
        DoEventsInterval (rowCount)
  
      Loop
  
      ' Строка статуса
      Application.StatusBar = ""

      ' Сообщение
      MsgBox ("ЛЦО в количестве " + CStr(CountЛЦО) + " шт. перенесены в каталог ЕСУП!")

    End If ' Перенос в ЕСУП


    ' Отправить ЛЦО в почту?
    If MsgBox("Отправить файлы ЛЦО сотрудников в почту?", vbYesNo) = vbYes Then
    
      ' Отправка
      Call Отправка_Lotus_Notes_Лист7_ЛЦО
    
    End If ' Отправить ЛЦО в почту?

  End If ' Запрос на формирование

  
  
End Sub



' Преобразование номер строки ProductCode в Столбец файла продаж (ПК в ПК, БС в БС и тд)
Function rowCount2_toColumn(In_rowCount2) As Integer
rowCount2_toColumn = 0
        Select Case In_rowCount2
          Case 2 ' Потреб кредитование
            rowCount2_toColumn = 2
          Case 3 ' Страховки к ПК
            rowCount2_toColumn = 5
          Case 4 ' Кредитные карты
            rowCount2_toColumn = 8
          Case 5 ' Дебетовые карты
            rowCount2_toColumn = 11
          Case 6 ' Интернет Банк
            rowCount2_toColumn = 14
          Case 7 ' Накопительный счет
            rowCount2_toColumn = 17
          Case 8 ' ИСЖ_МАСС (Премия, тыс.руб.)
            rowCount2_toColumn = 20
          Case 9 ' НСЖ_МАСС (комиссионный доход)
            rowCount2_toColumn = 23
          Case 10 ' Коробочное страхование
            rowCount2_toColumn = 26
          Case 11 ' Коробочное страхование (Антивирус + Ваша защита)
            rowCount2_toColumn = 29
        End Select
End Function

' Руководитель МРК
Function РуководительМРК(In_City) As String
        
        РуководительМРК = ""
        
        Select Case In_City
          Case "Тюменский"
            РуководительМРК = "НОРПиКО1"
          Case "Сургутский"
            РуководительМРК = "УДО2"
          Case "Нижневартовский"
            РуководительМРК = "УДО3"
          Case "Новоуренгойский"
            РуководительМРК = "УДО4"
          Case "Тарко-Сале"
            РуководительМРК = "УДО5"
        End Select
        
End Function

' Обработать Табель (кадры)
Sub Обработка_Табеля()
  
' Описание переменных
Dim ReportName_String, officeNameInReport, CheckFormatReportResult As String
Dim i, rowCount As Integer
Dim finishProcess As Boolean
Dim Пр_календарь_ГГГГ As String
    
  ' Открыть файл с отчетом
  FileName = Application.GetOpenFilename("Excel Files (*.XLS), *.XLS", , "Открытие файла с отчетом")

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
    ThisWorkbook.Sheets("Лист7").Activate

    ' Проверка формы отчета
    ' CheckFormatReportResult = CheckFormatReport(ReportName_String, "___", 6, Date)
    ' If CheckFormatReportResult = "OK" Then
    If True Then
      
      ' Открываем BASE\Sales
      OpenBookInBase ("Sales")
      
      ' Открываем BASE\TimeSheets
      OpenBookInBase ("TimeSheets")
      
      ' Открываем BASE\Пр. календарь ГГГГ
      ' OpenBookInBase ("Пр. календарь 2020")
      
      ' Обрабатываем отчет
      ' Цикл по 5-ти офисам
      ' Обработка отчета
      ' For i = 1 To 5
        ' Номера офисов от 1 до 5
      '  Select Case i
      '    Case 1 ' ОО «Тюменский»
      '      officeNameInReport = "Тюменский"
      '    Case 2 ' ОО «Сургутский»
      '      officeNameInReport = "Сургутский"
      '    Case 3 ' ОО «Нижневартовский»
      '      officeNameInReport = "Нижневартовский"
      '    Case 4 ' ОО «Новоуренгойский»
      '      officeNameInReport = "Новоуренгойский"
      '    Case 5 ' ОО «Тарко-Сале»
      '      officeNameInReport = "Тарко-Сале"
      '  End Select

        ' Row Табельный номер
        Row_Табельный_номер = rowByValue(ReportName_String, "Табель", "Табельный номер", 1000, 1000)
        
        ' Column Табельный номер
        column_Табельный_номер = ColumnByValue(ReportName_String, "Табель", "Табельный номер", 1000, 1000)

        ' Фамилия, инициалы, профессия (должность)
        Сolumn_ФИО = ColumnByValue(ReportName_String, "Табель", "Фамилия, инициалы, профессия (должность)", 1000, 1000)

        ' Отработано за
        Сolumn_Отработано = ColumnByValue(ReportName_String, "Табель", "Отработано за", 1000, 1000)

        ' Отчетный период
        Row_Отчетный_период = rowByValue(ReportName_String, "Табель", "Отчетный период", 1000, 1000)
        Сolumn_Отчетный_период = ColumnByValue(ReportName_String, "Табель", "Отчетный период", 1000, 1000)

        ' Месяц и год табеля
        period_MMYY = strMMYY(CDate(Workbooks(ReportName_String).Sheets("Табель").Cells(Row_Отчетный_период + 2, Сolumn_Отчетный_период).Value))

        ' Открываем соответствующий Производственный календарь BASE\Пр. календарь ГГГГ
        Пр_календарь_ГГГГ = "Пр. календарь " + CStr(Year(CDate(Workbooks(ReportName_String).Sheets("Табель").Cells(Row_Отчетный_период + 2, Сolumn_Отчетный_период).Value)))
        OpenBookInBase (Пр_календарь_ГГГГ)

        ' Берем нормативы Norm_hours, Norm_days для period_MMYY из производственного календаря
        Norm_hours = Производственный_календарь(Пр_календарь_ГГГГ, "Норма", ИмяМесяца3(CDate(Workbooks(ReportName_String).Sheets("Табель").Cells(Row_Отчетный_период + 2, Сolumn_Отчетный_период).Value)), "hours")
        Norm_days = Производственный_календарь(Пр_календарь_ГГГГ, "Норма", ИмяМесяца3(CDate(Workbooks(ReportName_String).Sheets("Табель").Cells(Row_Отчетный_период + 2, Сolumn_Отчетный_период).Value)), "days")
        
        ' Обработка Табеля
        rowCount = Row_Табельный_номер + 9
        ' Do While InStr(Workbooks(ReportName_String).Sheets("Табель").Cells(rowCount, Сolumn_ФИО).Value, "Ответственное лицо") = 0
        Do While (InStr(Workbooks(ReportName_String).Sheets("Табель").Cells(rowCount, Сolumn_ФИО).Value, "Ответственное лицо") = 0) Or (rowCount > 500)
        
          ' Если это табельный номер
          If Not IsEmpty(Workbooks(ReportName_String).Sheets("Табель").Cells(rowCount, column_Табельный_номер).Value) Then
            
            ' Фамилия, инициалы,
            FullName = Workbooks(ReportName_String).Sheets("Табель").Cells(rowCount, Сolumn_ФИО).Value
            
            ' Профессия (должность)
            Position = Workbooks(ReportName_String).Sheets("Табель").Cells(rowCount + 2, Сolumn_ФИО).Value
            
            ' Табельный номер
            Personnel_Number = Workbooks(ReportName_String).Sheets("Табель").Cells(rowCount, column_Табельный_номер).Value
            
            ' Отработано за, дни
            Working_days = CInt(Workbooks(ReportName_String).Sheets("Табель").Cells(rowCount, Сolumn_Отработано + 1).Value)
                        
            ' Отработано за, часы
            Working_hours = CDbl(Workbooks(ReportName_String).Sheets("Табель").Cells(rowCount + 1, Сolumn_Отработано + 1).Value)
            
            '  Идентификатор ID_Rec:
            ID_RecVar = CStr(Personnel_Number) + "-" + period_MMYY ' Пример: 5100313-0120

            ' Апдейтим BASE\Sales
            Call UpdateSales(Personnel_Number, period_MMYY, Working_hours, Working_days, Norm_hours, Norm_days)
            
            ' Офис берем из BASE\Sales
            Office_Var = getOfficeFromSales(Personnel_Number)
            
            ' Добавляем запись в BASE\TimeSheets
            Call InsertRecordInBook("TimeSheets", "Лист1", "ID_Rec", ID_RecVar, _
                                            "ID_Rec", ID_RecVar, _
                                              "Personnel_Number", Personnel_Number, _
                                                "Name", FullName, _
                                                  "Office", Office_Var, _
                                                    "MMYY", period_MMYY, _
                                                      "Working_hours", Working_hours, _
                                                        "Working_days", Working_days, _
                                                          "Norm_days", Norm_days, _
                                                            "Norm_hours", Norm_hours, _
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

                
          End If
        
        
          ' Следующая запись
          rowCount = rowCount + 1
          Application.StatusBar = "Сотрудник " + CStr(Personnel_Number) + "..."
          DoEventsInterval (rowCount)
        Loop
   
        ' Выводим данные по офису
      
      ' Next i ' Следующий офис
      
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
    Workbooks(Dir(FileName)).Close SaveChanges:=False
    
    ' Закрываем BASE\Sales
    CloseBook ("Sales")
        
    ' Закрываем BASE\TimeSheets
    CloseBook ("TimeSheets")
        
    ' Закрываем BASE\Пр. календарь 2020
    CloseBook (Пр_календарь_ГГГГ)
        
    ' Переходим в ячейку M2
    ThisWorkbook.Sheets("Лист7").Range("A1").Select

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

' Функция получения данных из Производственного календаря
Function Производственный_календарь(In_Book, In_Sheet, In_Month, In_Params) As Double
  
  ' Инициализация
  Производственный_календарь = 0
  
  ' Row Табельный номер
  Row_Месяц = rowByValue(In_Book, In_Sheet, In_Month, 100, 100)
        
  ' Column Табельный номер
  Column_Месяц = ColumnByValue(In_Book, In_Sheet, In_Month, 100, 100)
  
  ' Если ищем дни
  If In_Params = "days" Then
    Производственный_календарь = CDbl(Workbooks(In_Book).Sheets(In_Sheet).Cells(Row_Месяц + 3, Column_Месяц).Value)
  End If
  
  ' Если ищем часы
  If In_Params = "hours" Then
    Производственный_календарь = CDbl(Workbooks(In_Book).Sheets(In_Sheet).Cells(Row_Месяц + 6, Column_Месяц).Value)
  End If
  
End Function

' Апдейтим BASE\Sales
Sub UpdateSales(In_Personnel_Number, In_MMYY, In_Working_hours, In_Working_days, In_Norm_hours, In_Norm_days)
  
  ' Обработка Табеля
  rowCount = 2
  Do While Not IsEmpty(Workbooks("Sales").Sheets("Лист1").Cells(rowCount, 1).Value)
    
    ' Апдейтим
    If (Workbooks("Sales").Sheets("Лист1").Cells(rowCount, 2).Value = In_Personnel_Number) And (Workbooks("Sales").Sheets("Лист1").Cells(rowCount, 5).Value = In_MMYY) Then
      
      ' Working_hours
      Workbooks("Sales").Sheets("Лист1").Cells(rowCount, 14).Value = In_Working_hours
      ' Working_days
      Workbooks("Sales").Sheets("Лист1").Cells(rowCount, 15).Value = In_Working_days
      ' Norm_hours
      Workbooks("Sales").Sheets("Лист1").Cells(rowCount, 16).Value = In_Norm_hours
      ' Norm_days
      Workbooks("Sales").Sheets("Лист1").Cells(rowCount, 17).Value = In_Norm_days

    End If
    
    ' Следующая запись
    rowCount = rowCount + 1
    DoEventsInterval (rowCount)
    
  Loop

End Sub


' Офис берем из BASE\Sales
Function getOfficeFromSales(In_Personnel_Number)

  getOfficeFromSales = ""
  
  ' Ищем сотрудника
  
  ' Проверяем наличие записи In_FieldKeyName - In_FieldKeyValue
  ' Литера_столбца = ConvertToLetter(ColumnByName(In_BookName, In_Sheet, 1, "Personnel_Number"))
  ' Set searchResults = Workbooks(In_BookName).Sheets(In_Sheet).Columns(Литера_столбца + ":" + Литера_столбца).Find(In_FieldKeyValue, LookAt:=xlWhole)
  
  Set searchResults = Workbooks("Sales").Sheets("Лист1").Columns("B:B").Find(In_Personnel_Number, LookAt:=xlWhole)
  
  ' Проверяем - есть ли такая дата, если нет, то добавляем
  If searchResults Is Nothing Then
    ' Если не найдена - вставляем
  Else
    ' Если найдена, то апдейтим
    ' rowCount = searchResults.Row
    getOfficeFromSales = Workbooks("Sales").Sheets("Лист1").Cells(searchResults.Row, 4).Value
  End If


End Function

' Установка шрифта в ячейке
Sub setFontInRange(In_FileLCOName, In_LCOSheetsName, In_Range, In_FontName, In_FontSize)
      
  Workbooks(In_FileLCOName).Sheets(In_LCOSheetsName).Range(In_Range).Font.Name = In_FontName
  Workbooks(In_FileLCOName).Sheets(In_LCOSheetsName).Range(In_Range).Font.Size = In_FontSize
  Workbooks(In_FileLCOName).Sheets(In_LCOSheetsName).Range(In_Range).NumberFormat = "#,##0"
  Workbooks(In_FileLCOName).Sheets(In_LCOSheetsName).Range(In_Range).Font.ThemeColor = xlThemeColorLight1

End Sub

' Сформировать поручения на неделю для НОРПиКО и УДО
Sub createTaskFroWeekOffice()
  
  ' Запрос на формирование ИЗП
  If MsgBox("Сформировать поручения до " + CStr(weekEndDate(Date) - 2) + " для НОРПиКО и УДО?", vbYesNo) = vbYes Then
    
    ' ====
    currOffice = ""
    ' Обнуляем переменные
    Дефицит_Потреб_кредитование = 0
    
    ' Строка начала сотрудников
    rowCount = 9
    Do While Not IsEmpty(ThisWorkbook.Sheets("Лист7").Cells(rowCount, 3).Value)
    
      ' Текущий офис
      If currOffice <> ThisWorkbook.Sheets("Лист7").Cells(rowCount, 4).Value Then
        
        ' Если currOffice <> ""
        If currOffice <> "" Then
          
          ' *** Выводим поручения по офису ***
          ' 1) ПК
          If Дефицит_Потреб_кредитование > 0 Then
            Call Вставка_строки_в_Поручения_участникам("Иванов И.", Date, "Выдать кредит по " + currOffice + " на сумму " + CStr(Round(Дефицит_Потреб_кредитование, 0)))
          End If
          
          ' Обнуляем переменные
          Дефицит_Потреб_кредитование = 0
        End If
        
        ' Новое значение Офиса
        currOffice = ThisWorkbook.Sheets("Лист7").Cells(rowCount, 4).Value
      End If
      
      ' Если сейчас строка текущего офиса, то суммируем дефицит
      If currOffice = ThisWorkbook.Sheets("Лист7").Cells(rowCount, 4).Value Then
          
          ' Делаем ресчет дефицита по продуктам:
          ' 1) ПК
          Дефицит_Потреб_кредитование = Дефицит_Потреб_кредитование + (ThisWorkbook.Sheets("Лист7").Cells(rowCount, 5).Value - ThisWorkbook.Sheets("Лист7").Cells(rowCount, 6).Value)
          
      End If
      
      ' Следующая запись
      rowCount = rowCount + 1
      Application.StatusBar = "Расчёт " + CStr(ThisWorkbook.Sheets("Лист7").Cells(rowCount, 1).Value) + "..."
      DoEventsInterval (rowCount)

    Loop
    
    ' Выводим поручения по последнему офису
    Call Вставка_строки_в_Поручения_участникам("Иванов И.", Date, "Выдать кредит по " + currOffice + " на сумму " + CStr(Round(Дефицит_Потреб_кредитование, 0)))
         
    Application.StatusBar = ""
    
    
    ' Сообщение о неверном формате отчета или даты
    MsgBox ("Поручения сформированы!")

  End If
  
End Sub

' Сформировать поручения на неделю для НОРПиКО и УДО (второй вариант)
Sub createTaskFroWeekOffice2()
  
  ' Сообщение о неверном формате отчета или даты
  MsgBox ("Внимание! Необходимо обновить Cadr Emission перед началом!")
  
  ' Запрос на формирование ИЗП
  If MsgBox("Сформировать поручения по офисному каналу до " + CStr(weekEndDate(Date) - 2) + " для НОРПиКО и УДО?", vbYesNo) = vbYes Then
    
    
    ' Счетчик поручений
    Счетчик_поручений = 0
    
    ' Цикл по 5-ти офисам
    For i = 1 To 5
      
      ' Номера офисов от 1 до 5
      Select Case i
        Case 1 ' ОО «Тюменский»
          officeNameInReport = "Тюменский"
          responsibleName = getFromAddrBook("НОРПиКО1", 3)
          row_КД_Лист8 = getRowFromSheet8("Тюменский", "в т.ч. страховки к ПК") ' 25 ' в т.ч. страховки к ПК
          row_ЗаявкиКК_Лист8 = getRowFromSheet8("Тюменский", "Заявки на Кредитные карты") ' 43
          row_Карты_в_сейфе_Лист5 = 39
        Case 2 ' ОО «Сургутский»
          officeNameInReport = "Сургутский"
          responsibleName = getFromAddrBook("УДО2", 3)
          row_КД_Лист8 = getRowFromSheet8("Сургутский", "в т.ч. страховки к ПК") ' 63 в т.ч. страховки к ПК
          row_ЗаявкиКК_Лист8 = getRowFromSheet8("Сургутский", "Заявки на Кредитные карты") ' 81
          row_Карты_в_сейфе_Лист5 = 40
        Case 3 ' ОО «Нижневартовский»
          officeNameInReport = "Нижневартовский"
          responsibleName = getFromAddrBook("НОРПиКО3", 3)
          row_КД_Лист8 = getRowFromSheet8("Нижневартовский", "в т.ч. страховки к ПК") ' 101 ' в т.ч. страховки к ПК
          row_ЗаявкиКК_Лист8 = getRowFromSheet8("Нижневартовский", "Заявки на Кредитные карты") ' 119
          row_Карты_в_сейфе_Лист5 = 41
        Case 4 ' ОО «Новоуренгойский»
          officeNameInReport = "Новоуренгойский"
          responsibleName = getFromAddrBook("НОРПиКО4", 3)
          row_КД_Лист8 = getRowFromSheet8("Новоуренгойский", "в т.ч. страховки к ПК") ' 139 ' в т.ч. страховки к ПК
          row_ЗаявкиКК_Лист8 = getRowFromSheet8("Новоуренгойский", "Заявки на Кредитные карты") ' 157
          row_Карты_в_сейфе_Лист5 = 42
        Case 5 ' ОО «Тарко-Сале»
          officeNameInReport = "Тарко-Сале"
          responsibleName = getFromAddrBook("НОРПиКО5", 3)
          row_КД_Лист8 = getRowFromSheet8("Тарко-Сале", "в т.ч. страховки к ПК") ' 177 ' в т.ч. страховки к ПК
          row_ЗаявкиКК_Лист8 = getRowFromSheet8("Тарко-Сале", "Заявки на Кредитные карты") ' 195
          row_Карты_в_сейфе_Лист5 = 43
      End Select
    
    
    ' Обнуляем переменные
    Дефицит_ПК = 0
    Дефицит_КК = 0
    Дефицит_ДК = 0
    Дефицит_ИБ = 0
    Дефицит_НС = 0
    Дефицит_КСП = 0
    План_Заявки_КК_неделя = 0
    
    ' Строка начала сотрудников
    rowCount = 9
    Do While Not IsEmpty(ThisWorkbook.Sheets("Лист7").Cells(rowCount, 3).Value)
    
      ' Если это текущий офис
      If ThisWorkbook.Sheets("Лист7").Cells(rowCount, 4).Value = officeNameInReport Then
        
        ' Расчет дефицита по продуктам
        ' 1) ПК
        Дефицит_ПК = Дефицит_ПК + (ThisWorkbook.Sheets("Лист7").Cells(rowCount, 5).Value - ThisWorkbook.Sheets("Лист7").Cells(rowCount, 6).Value)
        ' 2) Кредитные карты
        Дефицит_КК = Дефицит_КК + (ThisWorkbook.Sheets("Лист7").Cells(rowCount, 15).Value - ThisWorkbook.Sheets("Лист7").Cells(rowCount, 16).Value)
        ' 3) Дебетовые карты
        Дефицит_ДК = Дефицит_ДК + (ThisWorkbook.Sheets("Лист7").Cells(rowCount, 20).Value - ThisWorkbook.Sheets("Лист7").Cells(rowCount, 21).Value)
        ' 4) ИБ
        Дефицит_ИБ = Дефицит_ИБ + (ThisWorkbook.Sheets("Лист7").Cells(rowCount, 25).Value - ThisWorkbook.Sheets("Лист7").Cells(rowCount, 26).Value)
        ' 5) НС
        Дефицит_НС = Дефицит_НС + (ThisWorkbook.Sheets("Лист7").Cells(rowCount, 30).Value - ThisWorkbook.Sheets("Лист7").Cells(rowCount, 31).Value)
        ' 6) КСП
        Дефицит_КСП = Дефицит_КСП + (ThisWorkbook.Sheets("Лист7").Cells(rowCount, 45).Value - ThisWorkbook.Sheets("Лист7").Cells(rowCount, 46).Value)
        
        
      End If
      
      ' Следующая запись
      rowCount = rowCount + 1
      Application.StatusBar = "Расчёт " + officeNameInReport + ": " + CStr(rowCount) + "..."
      DoEventsInterval (rowCount)

    Loop
    
    Application.StatusBar = "Формирование пакета поручений " + CStr(i) + "..."

    ' Дата начала недели
    Дата_начала_недели = weekStartDate(Date)
    ' Дата окончания недели
    Дата_окончания_недели = Дата_начала_недели + 4
    ' Остаток рабочих дней определяем число рабочих дней с понеделника до конца месяца Working_days_between_dates(In_DateStart, In_DateEnd, In_working_days_in_the_week) As Integer
    Остаток_рабочих_дней = Working_days_between_dates(Дата_начала_недели, Date_last_day_month(Дата_начала_недели), 5)
          
    ' Выводим поручения по текущему офису
    ' 1) ПК
    План_ПК_неделя = Round(Дефицит_ПК / Остаток_рабочих_дней, 0) * (Дата_окончания_недели - Дата_начала_недели + 1)
    If План_ПК_неделя > 0 Then
      Call Вставка_строки_в_Поручения_участникам(responsibleName, Дата_окончания_недели, "Обеспечить выдачу потребительских кредитов в офисном канале на сумму не менее " + CStr(План_ПК_неделя) + " тыс.руб.")
      ' Счетчик поручений
      Счетчик_поручений = Счетчик_поручений + 1
    End If
    
    ' 2) КК
    ' Заявки КК на месяц
    Дефицит_Заявки_КК = Round(ThisWorkbook.Sheets("Лист8").Cells(row_ЗаявкиКК_Лист8, 9).Value - ThisWorkbook.Sheets("Лист8").Cells(row_ЗаявкиКК_Лист8, 10).Value, 0)
    План_Заявки_КК_неделя = Round(Дефицит_Заявки_КК / Остаток_рабочих_дней, 0) * (Дата_окончания_недели - Дата_начала_недели + 1)
    If План_Заявки_КК_неделя > 0 Then
      Call Вставка_строки_в_Поручения_участникам(responsibleName, Дата_окончания_недели, "Завести заявки на кредитные карты не менее " + CStr(План_Заявки_КК_неделя) + " шт.")
      ' Счетчик поручений
      Счетчик_поручений = Счетчик_поручений + 1
    End If
    
    ' Выдать карты из сейфа row_Карты_в_сейфе_Лист5
    
    
    ' Активация КК карт
    ' План_КК_неделя = Round(Дефицит_КК / Остаток_рабочих_дней, 0) * (Дата_окончания_недели - Дата_начала_недели + 1)
    ' If План_КК_неделя > 0 Then
    '   Call Вставка_строки_в_Поручения_участникам(responsibleName, Дата_окончания_недели, "Произвести активацию кредитных карт не менее " + CStr(План_КК_неделя) + " шт. Выдать Сплиты к ПК не менее 50%.")
      ' Счетчик поручений
    '   Счетчик_поручений = Счетчик_поручений + 1
    ' End If
     
    
    ' 3) ДК
    План_ДК_неделя = Round(Дефицит_ДК / Остаток_рабочих_дней, 0) * (Дата_окончания_недели - Дата_начала_недели + 1)
    If План_ДК_неделя > 0 Then
      Call Вставка_строки_в_Поручения_участникам(responsibleName, Дата_окончания_недели, "Обеспечить заведение заявок дебетовых карт не менее " + CStr(План_ДК_неделя) + " шт.")
      ' Счетчик поручений
      Счетчик_поручений = Счетчик_поручений + 1
    End If
    
    ' 4) ИБ
    План_ИБ_неделя = Round(Дефицит_ИБ / Остаток_рабочих_дней, 0) * (Дата_окончания_недели - Дата_начала_недели + 1)
    If План_ИБ_неделя > 0 Then
      Call Вставка_строки_в_Поручения_участникам(responsibleName, Дата_окончания_недели, "Обеспечить подключение Интернет-банка в кол-ве не менее " + CStr(План_ИБ_неделя) + " шт.")
      ' Счетчик поручений
      Счетчик_поручений = Счетчик_поручений + 1
    End If
    
    ' 5) НС
    План_НС_неделя = Round(Дефицит_НС / Остаток_рабочих_дней, 0) * (Дата_окончания_недели - Дата_начала_недели + 1)
    If План_НС_неделя > 0 Then
      Call Вставка_строки_в_Поручения_участникам(responsibleName, Дата_окончания_недели, "Открыть Накопительные счета не менее " + CStr(План_НС_неделя) + " шт.")
      ' Счетчик поручений
      Счетчик_поручений = Счетчик_поручений + 1
    End If
    ' 6) КСП
    План_КСП_неделя = Round(Дефицит_КСП / Остаток_рабочих_дней, 0) * (Дата_окончания_недели - Дата_начала_недели + 1)
    If План_НС_неделя > 0 Then
      Call Вставка_строки_в_Поручения_участникам(responsibleName, Дата_окончания_недели, "Обеспечить продажу КСП не менее " + CStr(План_КСП_неделя) + " шт.")
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
      ' Счетчик поручений
      Счетчик_поручений = Счетчик_поручений + 1
    End If
              
    Application.StatusBar = ""
              
              
    Next i
         
    Application.StatusBar = ""
        
    ' Сообщение о неверном формате отчета или даты
    MsgBox ("Поручения в количестве " + CStr(Счетчик_поручений) + " сформированы!")

    ' Перейти на Лист ЕСУП
    Call goToSheetЕСУП
  
    ' Переход в часть листа с Поручениями
    ThisWorkbook.Sheets("ЕСУП").Range("AF77").Select

  End If
  
End Sub

' Расчет факта по офису и продукту на Лист7
Function getDataFromSheet7(In_Office, In_ProductName_Лист7)
  
  ' Итоговое значение
  getDataFromSheet7 = 0
  
  ' Берем с листа ОО «Тюменский»
  rowCount = rowByValue(ThisWorkbook.Name, "Лист7", "Форма 7.1", 100, 100) + 3
  
  ' Находим столбец с продуктом
  ColumnCount = ColumnByValue(ThisWorkbook.Name, "Лист7", In_ProductName_Лист7, 100, 100)
  
  ' Обрабатываем Лист - ищем Сначала Офис, если находим офис, то ищем позицию с наименованием продукта
  Do While Not IsEmpty(ThisWorkbook.Sheets("Лист7").Cells(rowCount, 1).Value)
    
    ' Проверяем офис
    If InStr(ThisWorkbook.Sheets("Лист7").Cells(rowCount, 4).Value, In_Office) <> 0 Then
      ' Суммируем Итоговое значение
      getDataFromSheet7 = getDataFromSheet7 + ThisWorkbook.Sheets("Лист7").Cells(rowCount, ColumnCount).Value
    End If
    
    ' Следующая запись
    rowCount = rowCount + 1
  Loop
    
End Function


' Отправка письма: отправляю шаблон самому себе для последующей отправки в сеть письма на его основе:
Sub Отправка_Lotus_Notes_Лист7_ЛЦО()
Dim темаПисьма, текстПисьма, hashTag, attachmentFile As String
Dim i As Byte
  
  ' Строка статуса
  Application.StatusBar = "Отправка письма с ЛЦО..."
  
  
  ' Запрос
  ' If MsgBox("Отправить себе Шаблон письма с фокусами контроля '" + ПериодКонтроля + "'?", vbYesNo) = vbYes Then
    
    ' Тема письма - Тема:
    ' темаПисьма = ThisWorkbook.Sheets("Лист8").Cells(RowByValue(ThisWorkbook.Name, "Лист8", "Тема:", 100, 100), ColumnByValue(ThisWorkbook.Name, "Лист8", "Тема:", 100, 100) + 1).Value
    темаПисьма = "Листы целевых ориентиров МРК/ВМРК по офисным продажам " + strDDMM(Date) + "-" + strDDMM(Date + 4) ' strДД_MM_YY2(Date + 4)

    ' hashTag - Хэштэг:
    ' hashTag = ThisWorkbook.Sheets("Лист8").Cells(RowByValue(ThisWorkbook.Name, "Лист8", "Хэштэг:", 100, 100), ColumnByValue(ThisWorkbook.Name, "Лист8", "Хэштэг:", 100, 100) + 1).Value
    ' hashTag - Хэштэг:
    hashTag = "#лцо #лцо_неделя"

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
    текстПисьма = текстПисьма + "Направляю индивидуальные Листы целевых ориентиров (ЛЦО) для МРК/ВМРК по ключевым продуктам на период с " + strDDMM(Date) + " по " + strDDMM(Date + 4) + " *. " + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "Прошу сегодня разобрать с каждым, внести актуальные данные на " + strDDMM(Date) + " и ежедневно контролировать отклонения!" + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "* - ЛЦО сформированы на основе факта продаж по Интегральному рейтингу от " + ((Mid(ThisWorkbook.Sheets("Лист7").Range("B5").Value, 40, 10))) + " г." + Chr(13)
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

' Отправка ЛЦО в почту
Sub Отправка_Лист7_ЛЦО_в_почту()

    ' Отправить ЛЦО в почту?
    If MsgBox("Отправить ЛЦО в почту?", vbYesNo) = vbYes Then
    
      ' Отправка
      Call Отправка_Lotus_Notes_Лист7_ЛЦО
    
    End If ' Отправить ЛЦО в почту?


End Sub

' Отправка письма: отправляю шаблон самому себе для последующей отправки в сеть письма на его основе:
Sub Отправка_Lotus_Notes_Лист7_ИПЗ()
Dim темаПисьма, текстПисьма, hashTag, attachmentFile As String
Dim i As Byte
  
  ' Строка статуса
  Application.StatusBar = "Отправка письма с ИПЗ..."
  
  
  ' Запрос
  ' If MsgBox("Отправить себе Шаблон письма с фокусами контроля '" + ПериодКонтроля + "'?", vbYesNo) = vbYes Then
    
    ' Тема письма - Тема:
    ' темаПисьма = ThisWorkbook.Sheets("Лист8").Cells(RowByValue(ThisWorkbook.Name, "Лист8", "Тема:", 100, 100), ColumnByValue(ThisWorkbook.Name, "Лист8", "Тема:", 100, 100) + 1).Value
    темаПисьма = "Индивидуальные плановые задания МРК/ВМРК по офисным продажам на " + ИмяМесяцаГод(CDate(Mid(ThisWorkbook.Sheets("Лист7").Range("B5").Value, 40, 10)))

    ' hashTag - Хэштэг:
    ' hashTag = ThisWorkbook.Sheets("Лист8").Cells(RowByValue(ThisWorkbook.Name, "Лист8", "Хэштэг:", 100, 100), ColumnByValue(ThisWorkbook.Name, "Лист8", "Хэштэг:", 100, 100) + 1).Value
    ' hashTag - Хэштэг:
    hashTag = "#ИПЗ #ИПЗ_месяц"

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
    текстПисьма = текстПисьма + "Направляю индивидуальные плановые задания (ИПЗ) для МРК/ВМРК по ключевым продуктам на " + ИмяМесяцаГод(CDate(Mid(ThisWorkbook.Sheets("Лист7").Range("B5").Value, 40, 10))) + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "Прошу довести до каждого сотрудника." + Chr(13)
    ' текстПисьма = текстПисьма + "" + Chr(13)
    ' текстПисьма = текстПисьма + "" + Chr(13)
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

' Отправка ИПЗ в почту
Sub Отправка_Лист7_ИПЗ_в_почту()

    ' Отправить ИПЗ в почту?
    If MsgBox("Отправить ИПЗ в почту?", vbYesNo) = vbYes Then
    
      ' Отправка
      Call Отправка_Lotus_Notes_Лист7_ИПЗ
    
    End If ' Отправить ИПЗ в почту?


End Sub

' Число сотрудников в офисе на Листе7
Function Число_сотрудников_в_офисе_Лист7(In_NumberOffice) As Integer
  
  Число_сотрудников_в_офисе_Лист7 = 0
  
  ' "Форма 7.1"
  row_Форма_7_1 = rowByValue(ThisWorkbook.Name, "Лист7", "Форма 7.1", 10, 10)
  
  ' 1 - Тюменский
  getNameOfficeByNumber2_Var = getNameOfficeByNumber2(In_NumberOffice)
  
  rowCount_Лист7 = row_Форма_7_1 + 3
  Do While (ThisWorkbook.Sheets("Лист7").Cells(rowCount_Лист7, 1).Value) <> ""
  
    ' t1 = ThisWorkbook.Sheets("Лист7").Cells(rowCount_Лист7, 4).Value
  
    If InStr(ThisWorkbook.Sheets("Лист7").Cells(rowCount_Лист7, 4).Value, getNameOfficeByNumber2_Var) <> 0 Then
      Число_сотрудников_в_офисе_Лист7 = Число_сотрудников_в_офисе_Лист7 + 1
    End If
  
    ' Следующая запись
    rowCount_Лист7 = rowCount_Лист7 + 1
    ' Application.StatusBar = CStr(ThisWorkbook.Sheets("Лист7").Cells(recInЛист7, 2).Value) + "..."
    ' DoEventsInterval (rowCount_Лист1)
          
  Loop
  
  ' t = 0
  
End Function

' План по сотрудникам в офисе (МРК/ВМРК)
Function План_штат_МРК_офис(In_NumberOffice) As Integer
  
  План_штат_МРК_офис = 0
  
  Select Case In_NumberOffice
    Case 1
      План_штат_МРК_офис = 3
    Case 2
      План_штат_МРК_офис = 2
    Case 3
      План_штат_МРК_офис = 2
    Case 4
      План_штат_МРК_офис = 3
    Case 5
      План_штат_МРК_офис = 3
  End Select

  
End Function

