Attribute VB_Name = "Module_КроссЗП"
' *** Лист КроссЗП ***

' *** Глобальные переменные ***
' Public numStr_Лист8 As Integer
' ***                       ***

' Обработка отчета http://isrb.psbnk.msk.ru/inf/6601/6622/otchet_zp_org/
Sub Анализ_ЗП_организаций_по_РА_кураторам()

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
  
    ' Открываем выбранную книгу (UpdateLinks:=0)
    Workbooks.Open FileName, 0
      
    ' Переходим на окно DB
    ThisWorkbook.Sheets("КроссЗП").Activate

    ' Проверка формы отчета
    CheckFormatReportResult = CheckFormatReport(ReportName_String, "сводная по ТП", 17, Date)
    If CheckFormatReportResult = "OK" Then
      
      ' Очистка данных
      Call clearСontents2(ThisWorkbook.Name, "КроссЗП", "C6", "N11")
      
      ' Дата отчета c "Реестр компаний" "B2"
      dateReport = CDate(Workbooks(ReportName_String).Sheets("Реестр компаний").Range("B2").Value)
      
      ' Тема письма - Тема:
      ThisWorkbook.Sheets("КроссЗП").Cells(rowByValue(ThisWorkbook.Name, "КроссЗП", "Тема:", 100, 100), ColumnByValue(ThisWorkbook.Name, "КроссЗП", "Тема:", 100, 100) + 1).Value = "Проникновение продуктов в ЗП-организации на " + strДД_MM_YYYY(dateReport)
      
      ' Запись обработки таблицы
      startRow = rowByValue(ReportName_String, "сводная по ТП", "бакет по численности", 100, 100) + 1
      
      ' Workbooks(ReportName_String).Sheets("сводная по ТП").Activate
      ' Workbooks(ReportName_String).Sheets("сводная по ТП").PivotTables("Сводная таблица8").PivotFields("Филиал").PivotItems("Тюменский ОО1").ShowDetail = True
      Раскрыт_список_офисов = False
      Определены_столбцы_на_Лист1 = False
      
      ' Создаем выходную таблицу для выгрузки
      OutBookName = ThisWorkbook.Path + "\Out\КроссЗП_" + strДД_MM_YYYY(dateReport) + ".xlsx"
      
      ' Проверяем - если файл есть, то удаляем его
      Call deleteFile(OutBookName)

      ' Вложение2
      ThisWorkbook.Sheets("КроссЗП").Range("U3").Value = OutBookName
      ' Создать файл
      Call createBook_out_КроссЗП(OutBookName)

      ' Переходим на окно DB
      ThisWorkbook.Sheets("КроссЗП").Activate

      
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

        ' Номер строки офиса на "КроссЗП"
        RowOfficeInSheet = rowByValue(ThisWorkbook.Name, "КроссЗП", "ОО «" + officeNameInReport + "»", 100, 100)
        rowCount = startRow
        Do While Not IsEmpty(Workbooks(ReportName_String).Sheets("сводная по ТП").Cells(rowCount, 1).Value)
        
          ' Раскрыть список офисов
          If (InStr(Workbooks(ReportName_String).Sheets("сводная по ТП").Cells(rowCount, 1).Value, "Тюменский ОО1") <> 0) And (Раскрыт_список_офисов = False) Then
            
            ' Старая версия
            ' Раскрытие списка офисов
            ' Workbooks(ReportName_String).Sheets("сводная по ТП").PivotTables("Сводная таблица8").PivotFields("Филиал").PivotItems("Тюменский ОО1").ShowDetail = True
            ' Раскрытие списка организаций c Лист1
            Workbooks(ReportName_String).Sheets("сводная по ТП").Cells(rowCount, 2).ShowDetail = True
            
            ' Новая версия
            ' Range("B65").Select
            ' Selection.ShowDetail = True
            
            ' Переход на страницу
            ThisWorkbook.Sheets("КроссЗП").Activate
            
            ' Переменная
            Раскрыт_список_офисов = True
          
          End If
          
          ' Если это текущий офис
          If (InStr(Workbooks(ReportName_String).Sheets("сводная по ТП").Cells(rowCount, 1).Value, officeNameInReport) <> 0) And (InStr(Workbooks(ReportName_String).Sheets("сводная по ТП").Cells(rowCount, 1).Value, "ОО1") = 0) Then
            
            ' Кол-во орг
            ThisWorkbook.Sheets("КроссЗП").Cells(RowOfficeInSheet, 3).Value = Workbooks(ReportName_String).Sheets("сводная по ТП").Cells(rowCount, 2).Value
            ThisWorkbook.Sheets("КроссЗП").Cells(RowOfficeInSheet, 3).HorizontalAlignment = xlRight
                
          End If
        
        
          ' Следующая запись
          rowCount = rowCount + 1
          Application.StatusBar = officeNameInReport + ": " + CStr(rowCount) + "..."
          DoEventsInterval (rowCount)
        Loop
   
        ' Обработка данных по офису по ЗП организациям на Лист2
        If Определены_столбцы_на_Лист2 = False Then
          Application.StatusBar = "Определение столбцов..."
          column_Филиал = ColumnByValue(ReportName_String, "Лист2", "Филиал", 1000, 1000)
          column_офис = ColumnByValue(ReportName_String, "Лист2", "Офис", 1000, 1000)
          column_ИНН = ColumnByValue(ReportName_String, "Лист2", "ИНН", 1000, 1000)
          column_Организация = ColumnByValue(ReportName_String, "Лист2", "Организация", 1000, 1000)
          column_SALORGCD = ColumnByValue(ReportName_String, "Лист2", "SALORGCD", 1000, 1000)
          column_Кол_во_клиентов_с_наличием_зачислений_в_теч_6_последних_мес = ColumnByValue(ReportName_String, "Лист2", "Кол-во клиентов с наличием  зачислений в теч 6 последних мес.", 1000, 1000)
          column_ядро_ЗП = ColumnByValue(ReportName_String, "Лист2", "ядро ЗП", 1000, 1000)
          column_есть_открытый_ПК = ColumnByValue(ReportName_String, "Лист2", "есть открытый ПК", 1000, 1000)
          column_есть_КК = ColumnByValue(ReportName_String, "Лист2", "есть КК", 1000, 1000)
          column_есть_НС = ColumnByValue(ReportName_String, "Лист2", "есть НС", 1000, 1000)
          column_акт_РА_ПК = ColumnByValue(ReportName_String, "Лист2", "акт РА-ПК", 1000, 1000)
          column_акт_CRM_ПК_решение_или_предложение = ColumnByValue(ReportName_String, "Лист2", "акт CRM-ПК (решение или предложение)", 1000, 1000)
          column_акт_РА_КК = ColumnByValue(ReportName_String, "Лист2", "акт РА-КК", 1000, 1000)
          column_средний_AR_по_клиентам = ColumnByValue(ReportName_String, "Лист2", "средний AR по клиентам", 1000, 1000)
          column_категория = ColumnByValue(ReportName_String, "Лист2", "категория", 1000, 1000)
          column_средняя_ставка = ColumnByValue(ReportName_String, "Лист2", "средняя ставка", 1000, 1000)
          Application.StatusBar = "Столбцы определены!"
          Определены_столбцы_на_Лист2 = True
        End If
        
        ' Обнуление переменных
        Итого_ЗП_проекты = 0
        Итого_Ядро_ЗП = 0
        Итого_Число_клиентов_с_зачислениями_в_Ядре_ЗП = 0
        Итого_Наличие_ПК = 0
        Итого_Активный_PA_ПК = 0
        Итого_Активный_CRM_ПК = 0
        Итого_Наличие_КК = 0
        Итого_Активный_PA_КК = 0
        Итого_НС = 0
        
        rowCount = 2
        Do While Not IsEmpty(Workbooks(ReportName_String).Sheets("Лист2").Cells(rowCount, 1).Value)
                  
          ' Если это ЗП текущего офиса
          If (InStr(Workbooks(ReportName_String).Sheets("Лист2").Cells(rowCount, column_офис).Value, officeNameInReport) <> 0) Then
            
            ' ЗП_проекты
            Итого_ЗП_проекты = Итого_ЗП_проекты + 1
            ' Ядро ЗП
            Итого_Ядро_ЗП = Итого_Ядро_ЗП + Workbooks(ReportName_String).Sheets("Лист2").Cells(rowCount, column_ядро_ЗП).Value
            ' Число клиентов с зачислениями в Ядре ЗП
            Итого_Число_клиентов_с_зачислениями_в_Ядре_ЗП = Итого_Число_клиентов_с_зачислениями_в_Ядре_ЗП + Workbooks(ReportName_String).Sheets("Лист2").Cells(rowCount, column_Кол_во_клиентов_с_наличием_зачислений_в_теч_6_последних_мес).Value
            ' Число клиентов с Наличием ПК
            Итого_Наличие_ПК = Итого_Наличие_ПК + Workbooks(ReportName_String).Sheets("Лист2").Cells(rowCount, column_есть_открытый_ПК).Value
            ' Активный PA ПК
            Итого_Активный_PA_ПК = Итого_Активный_PA_ПК + Workbooks(ReportName_String).Sheets("Лист2").Cells(rowCount, column_акт_РА_ПК).Value
            ' Активный CRM-ПК
            Итого_Активный_CRM_ПК = Итого_Активный_CRM_ПК + Workbooks(ReportName_String).Sheets("Лист2").Cells(rowCount, column_акт_CRM_ПК_решение_или_предложение).Value
            ' Наличие КК
            Итого_Наличие_КК = Итого_Наличие_КК + Workbooks(ReportName_String).Sheets("Лист2").Cells(rowCount, column_есть_КК).Value
            ' Активный PA КК
            Итого_Активный_PA_КК = Итого_Активный_PA_КК + Workbooks(ReportName_String).Sheets("Лист2").Cells(rowCount, column_акт_РА_КК).Value
            ' НС
            Итого_НС = Итого_НС + Workbooks(ReportName_String).Sheets("Лист2").Cells(rowCount, column_есть_НС).Value
            ' Итого_... = Итого_... + Workbooks(ReportName_String).Sheets("Лист2").Cells(rowCount, ...).Value
            
            ' Выгружаем в таблицу
            Call InsertRecordInBook(Dir(OutBookName), "Лист1", "SalOrgCD", Workbooks(ReportName_String).Sheets("Лист2").Cells(rowCount, column_SALORGCD).Value, _
                                              "Офис", officeNameInReport, _
                                                "Организация_ИНН", Workbooks(ReportName_String).Sheets("Лист2").Cells(rowCount, column_ИНН).Value, _
                                                  "Организация_Наименование", Workbooks(ReportName_String).Sheets("Лист2").Cells(rowCount, column_Организация).Value, _
                                                    "SalOrgCD", Workbooks(ReportName_String).Sheets("Лист2").Cells(rowCount, column_SALORGCD).Value, _
                                                      "Категория", Workbooks(ReportName_String).Sheets("Лист2").Cells(rowCount, column_категория).Value, _
                                                        "Ядро_ЗП_шт", Workbooks(ReportName_String).Sheets("Лист2").Cells(rowCount, column_ядро_ЗП).Value, _
                                                          "Клиенты_с_зачислениями_в_теч_последних_6_мес_шт", Workbooks(ReportName_String).Sheets("Лист2").Cells(rowCount, column_Кол_во_клиентов_с_наличием_зачислений_в_теч_6_последних_мес).Value, _
                                                            "Открытый_ПК_шт", Workbooks(ReportName_String).Sheets("Лист2").Cells(rowCount, column_есть_открытый_ПК).Value, _
                                                              "Cредняя_ставка", Workbooks(ReportName_String).Sheets("Лист2").Cells(rowCount, column_средняя_ставка).Value, _
                                                                "Активные_РА_ПК_шт", Workbooks(ReportName_String).Sheets("Лист2").Cells(rowCount, column_акт_РА_ПК).Value, _
                                                                  "Активные_CRM_ПК_решение_предложение_шт", Workbooks(ReportName_String).Sheets("Лист2").Cells(rowCount, column_акт_CRM_ПК_решение_или_предложение).Value, _
                                                                    "Наличие_КК_шт", Workbooks(ReportName_String).Sheets("Лист2").Cells(rowCount, column_есть_КК).Value, _
                                                                      "Активные_РА_КК_шт", Workbooks(ReportName_String).Sheets("Лист2").Cells(rowCount, column_акт_РА_КК).Value, _
                                                                        "Cредний_AR_по_клиентам", Workbooks(ReportName_String).Sheets("Лист2").Cells(rowCount, column_средний_AR_по_клиентам).Value, _
                                                                          "Наличие_НС_шт", Workbooks(ReportName_String).Sheets("Лист2").Cells(rowCount, column_есть_НС).Value, _
                                                                            "", "", _
                                                                              "", "", _
                                                                                "", "", _
                                                                                  "", "", _
                                                                                    "", "")

            
          End If ' Если это ЗП текущего офиса
          
          ' Следующая запись
          rowCount = rowCount + 1
          Application.StatusBar = officeNameInReport + " ЗП проекты: " + CStr(rowCount) + "..."
          DoEventsInterval (rowCount)
        Loop
      
        ' Выводим итоги обработки:
        ' ЗП проекты
        ThisWorkbook.Sheets("КроссЗП").Cells(RowOfficeInSheet, 3).Value = Итого_ЗП_проекты
        ThisWorkbook.Sheets("КроссЗП").Cells(RowOfficeInSheet, 3).NumberFormat = "#,##0"
        ThisWorkbook.Sheets("КроссЗП").Cells(RowOfficeInSheet, 3).HorizontalAlignment = xlRight
        ' Ядро ЗП
        ThisWorkbook.Sheets("КроссЗП").Cells(RowOfficeInSheet, 4).Value = Итого_Ядро_ЗП
        ThisWorkbook.Sheets("КроссЗП").Cells(RowOfficeInSheet, 4).NumberFormat = "#,##0"
        ThisWorkbook.Sheets("КроссЗП").Cells(RowOfficeInSheet, 4).HorizontalAlignment = xlRight
        ' Число клиентов с зачислениями в Ядре ЗП
        ThisWorkbook.Sheets("КроссЗП").Cells(RowOfficeInSheet, 5).Value = Итого_Число_клиентов_с_зачислениями_в_Ядре_ЗП
        ThisWorkbook.Sheets("КроссЗП").Cells(RowOfficeInSheet, 5).NumberFormat = "#,##0"
        ThisWorkbook.Sheets("КроссЗП").Cells(RowOfficeInSheet, 5).HorizontalAlignment = xlRight
        ' Число клиентов с Наличием ПК
        ThisWorkbook.Sheets("КроссЗП").Cells(RowOfficeInSheet, 6).Value = Итого_Наличие_ПК
        ThisWorkbook.Sheets("КроссЗП").Cells(RowOfficeInSheet, 6).NumberFormat = "#,##0"
        ThisWorkbook.Sheets("КроссЗП").Cells(RowOfficeInSheet, 6).HorizontalAlignment = xlRight
        ' Проникновение ПК в Ядро
        ThisWorkbook.Sheets("КроссЗП").Cells(RowOfficeInSheet, 7).Value = Итого_Наличие_ПК / Итого_Ядро_ЗП
        ThisWorkbook.Sheets("КроссЗП").Cells(RowOfficeInSheet, 7).NumberFormat = "0.0%"
        ThisWorkbook.Sheets("КроссЗП").Cells(RowOfficeInSheet, 7).HorizontalAlignment = xlRight
        ' Активный PA ПК
        ThisWorkbook.Sheets("КроссЗП").Cells(RowOfficeInSheet, 8).Value = Итого_Активный_PA_ПК
        ThisWorkbook.Sheets("КроссЗП").Cells(RowOfficeInSheet, 8).NumberFormat = "#,##0"
        ThisWorkbook.Sheets("КроссЗП").Cells(RowOfficeInSheet, 8).HorizontalAlignment = xlRight
        ' Активный CRM-ПК
        ThisWorkbook.Sheets("КроссЗП").Cells(RowOfficeInSheet, 9).Value = Итого_Активный_CRM_ПК
        ThisWorkbook.Sheets("КроссЗП").Cells(RowOfficeInSheet, 9).NumberFormat = "#,##0"
        ThisWorkbook.Sheets("КроссЗП").Cells(RowOfficeInSheet, 9).HorizontalAlignment = xlRight
        ' Наличие КК
        ThisWorkbook.Sheets("КроссЗП").Cells(RowOfficeInSheet, 10).Value = Итого_Наличие_КК
        ThisWorkbook.Sheets("КроссЗП").Cells(RowOfficeInSheet, 10).NumberFormat = "#,##0"
        ThisWorkbook.Sheets("КроссЗП").Cells(RowOfficeInSheet, 10).HorizontalAlignment = xlRight
        ' Проникновение КК в Ядро
        ThisWorkbook.Sheets("КроссЗП").Cells(RowOfficeInSheet, 11).Value = Итого_Наличие_КК / Итого_Ядро_ЗП
        ThisWorkbook.Sheets("КроссЗП").Cells(RowOfficeInSheet, 11).NumberFormat = "0.0%"
        ThisWorkbook.Sheets("КроссЗП").Cells(RowOfficeInSheet, 11).HorizontalAlignment = xlRight
        ' Активный PA КК
        ThisWorkbook.Sheets("КроссЗП").Cells(RowOfficeInSheet, 12).Value = Итого_Активный_PA_КК
        ThisWorkbook.Sheets("КроссЗП").Cells(RowOfficeInSheet, 12).NumberFormat = "#,##0"
        ThisWorkbook.Sheets("КроссЗП").Cells(RowOfficeInSheet, 12).HorizontalAlignment = xlRight
        ' НС
        ThisWorkbook.Sheets("КроссЗП").Cells(RowOfficeInSheet, 13).Value = Итого_НС
        ThisWorkbook.Sheets("КроссЗП").Cells(RowOfficeInSheet, 13).NumberFormat = "#,##0"
        ThisWorkbook.Sheets("КроссЗП").Cells(RowOfficeInSheet, 13).HorizontalAlignment = xlRight
        ' Проникновение НС в Ядро
        ThisWorkbook.Sheets("КроссЗП").Cells(RowOfficeInSheet, 14).Value = Итого_НС / Итого_Ядро_ЗП
        ThisWorkbook.Sheets("КроссЗП").Cells(RowOfficeInSheet, 14).NumberFormat = "0.0%"
        ThisWorkbook.Sheets("КроссЗП").Cells(RowOfficeInSheet, 14).HorizontalAlignment = xlRight
        
      
      Next i ' Следующий офис
      
      ' Итоги по офисам
      RowOfficeInSheet = RowOfficeInSheet + 1
      ThisWorkbook.Sheets("КроссЗП").Cells(RowOfficeInSheet, 4).Value = "Итого по РОО:"
      ' ЗП проекты, шт.
      ThisWorkbook.Sheets("КроссЗП").Cells(RowOfficeInSheet, 3).FormulaR1C1 = "=R[-5]C+R[-4]C+R[-3]C+R[-2]C+R[-1]C"
      ThisWorkbook.Sheets("КроссЗП").Cells(RowOfficeInSheet, 3).NumberFormat = "#,##0"
      ThisWorkbook.Sheets("КроссЗП").Cells(RowOfficeInSheet, 3).HorizontalAlignment = xlRight
      ' Ядро ЗП
      ThisWorkbook.Sheets("КроссЗП").Cells(RowOfficeInSheet, 4).FormulaR1C1 = "=R[-5]C+R[-4]C+R[-3]C+R[-2]C+R[-1]C"
      ThisWorkbook.Sheets("КроссЗП").Cells(RowOfficeInSheet, 4).NumberFormat = "#,##0"
      ThisWorkbook.Sheets("КроссЗП").Cells(RowOfficeInSheet, 4).HorizontalAlignment = xlRight
      ' Число клиентов с зачислениями в Ядре ЗП
      ThisWorkbook.Sheets("КроссЗП").Cells(RowOfficeInSheet, 5).FormulaR1C1 = "=R[-5]C+R[-4]C+R[-3]C+R[-2]C+R[-1]C"
      ThisWorkbook.Sheets("КроссЗП").Cells(RowOfficeInSheet, 5).NumberFormat = "#,##0"
      ThisWorkbook.Sheets("КроссЗП").Cells(RowOfficeInSheet, 5).HorizontalAlignment = xlRight
      ' Число клиентов с Наличием ПК
      ThisWorkbook.Sheets("КроссЗП").Cells(RowOfficeInSheet, 6).FormulaR1C1 = "=R[-5]C+R[-4]C+R[-3]C+R[-2]C+R[-1]C"
      ThisWorkbook.Sheets("КроссЗП").Cells(RowOfficeInSheet, 6).NumberFormat = "#,##0"
      ThisWorkbook.Sheets("КроссЗП").Cells(RowOfficeInSheet, 6).HorizontalAlignment = xlRight
      ' Проникновение ПК в Ядро
      ThisWorkbook.Sheets("КроссЗП").Cells(RowOfficeInSheet, 7).FormulaR1C1 = "=AVERAGE(R[-5]C:R[-1]C)"
      ThisWorkbook.Sheets("КроссЗП").Cells(RowOfficeInSheet, 7).NumberFormat = "0.0%"
      ThisWorkbook.Sheets("КроссЗП").Cells(RowOfficeInSheet, 7).HorizontalAlignment = xlRight
      ' Активный PA ПК
      ThisWorkbook.Sheets("КроссЗП").Cells(RowOfficeInSheet, 8).FormulaR1C1 = "=R[-5]C+R[-4]C+R[-3]C+R[-2]C+R[-1]C"
      ThisWorkbook.Sheets("КроссЗП").Cells(RowOfficeInSheet, 8).NumberFormat = "#,##0"
      ThisWorkbook.Sheets("КроссЗП").Cells(RowOfficeInSheet, 8).HorizontalAlignment = xlRight
      ' Активный CRM-ПК
      ThisWorkbook.Sheets("КроссЗП").Cells(RowOfficeInSheet, 9).FormulaR1C1 = "=R[-5]C+R[-4]C+R[-3]C+R[-2]C+R[-1]C"
      ThisWorkbook.Sheets("КроссЗП").Cells(RowOfficeInSheet, 9).NumberFormat = "#,##0"
      ThisWorkbook.Sheets("КроссЗП").Cells(RowOfficeInSheet, 9).HorizontalAlignment = xlRight
      ' Наличие КК
      ThisWorkbook.Sheets("КроссЗП").Cells(RowOfficeInSheet, 10).FormulaR1C1 = "=R[-5]C+R[-4]C+R[-3]C+R[-2]C+R[-1]C"
      ThisWorkbook.Sheets("КроссЗП").Cells(RowOfficeInSheet, 10).NumberFormat = "#,##0"
      ThisWorkbook.Sheets("КроссЗП").Cells(RowOfficeInSheet, 10).HorizontalAlignment = xlRight
      ' Проникновение КК в Ядро
      ThisWorkbook.Sheets("КроссЗП").Cells(RowOfficeInSheet, 11).FormulaR1C1 = "=R[-5]C+R[-4]C+R[-3]C+R[-2]C+R[-1]C"
      ThisWorkbook.Sheets("КроссЗП").Cells(RowOfficeInSheet, 11).NumberFormat = "0.0%"
      ThisWorkbook.Sheets("КроссЗП").Cells(RowOfficeInSheet, 11).HorizontalAlignment = xlRight
      ' Активный PA КК
      ThisWorkbook.Sheets("КроссЗП").Cells(RowOfficeInSheet, 12).FormulaR1C1 = "=R[-5]C+R[-4]C+R[-3]C+R[-2]C+R[-1]C"
      ThisWorkbook.Sheets("КроссЗП").Cells(RowOfficeInSheet, 12).NumberFormat = "#,##0"
      ThisWorkbook.Sheets("КроссЗП").Cells(RowOfficeInSheet, 12).HorizontalAlignment = xlRight
      ' НС
      ThisWorkbook.Sheets("КроссЗП").Cells(RowOfficeInSheet, 13).FormulaR1C1 = "=R[-5]C+R[-4]C+R[-3]C+R[-2]C+R[-1]C"
      ThisWorkbook.Sheets("КроссЗП").Cells(RowOfficeInSheet, 13).NumberFormat = "#,##0"
      ThisWorkbook.Sheets("КроссЗП").Cells(RowOfficeInSheet, 13).HorizontalAlignment = xlRight
      ' Проникновение НС в Ядро
      ThisWorkbook.Sheets("КроссЗП").Cells(RowOfficeInSheet, 14).FormulaR1C1 = "=AVERAGE(R[-5]C:R[-1]C)"
      ThisWorkbook.Sheets("КроссЗП").Cells(RowOfficeInSheet, 14).NumberFormat = "0.0%"
      ThisWorkbook.Sheets("КроссЗП").Cells(RowOfficeInSheet, 14).HorizontalAlignment = xlRight
     
     
      ' Сохранение изменений
      ThisWorkbook.Save
    
      ' Закрываем выходную книгу с выгрузкой КроссЗП
      Workbooks(Dir(OutBookName)).Close SaveChanges:=True
    
      ' Строка статуса
      Application.StatusBar = "Отправка сообщения..."
    
      ' Отправка сообщения
      Call Отправка_Lotus_Notes_КроссЗП
    
      ' Строка статуса
      Application.StatusBar = ""
    
      ' Переменная завершения обработки
      finishProcess = True
    Else
      ' Сообщение о неверном формате отчета или даты
      MsgBox ("Проверьте отчет: " + CheckFormatReportResult + "!")
    End If ' Проверка формы отчета

    ' Закрываем файл с отчетом без сохранения изменений (параметр SaveChanges:=False)
    Workbooks(Dir(FileName)).Close SaveChanges:=False
    
    ' Переходим в ячейку M2
    ThisWorkbook.Sheets("КроссЗП").Range("A1").Select

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


' Создание книги с Кросс ЗП
Sub createBook_out_КроссЗП(In_OutBookName)

    Application.StatusBar = "Создание " + In_OutBookName + "..."

    Workbooks.Add
    ActiveWorkbook.SaveAs FileName:=In_OutBookName
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Activate
    
    ' Форматирование полей
    ' Офис
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 1).Value = "Офис"
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns("A:A").EntireColumn.ColumnWidth = 17
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns("A:A").HorizontalAlignment = xlCenter
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 1).HorizontalAlignment = xlCenter
    
    ' Организация_ИНН
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 2).Value = "Организация_ИНН"
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns("B:B").EntireColumn.ColumnWidth = 19
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns("B:B").NumberFormat = "0"
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns("B:B").HorizontalAlignment = xlLeft
    
    ' Организация_Наименование
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 3).Value = "Организация_Наименование"
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns("C:C").EntireColumn.ColumnWidth = 40
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 3).HorizontalAlignment = xlCenter
    
    ' SalOrgCD
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 4).Value = "SalOrgCD"
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns("D:D").EntireColumn.ColumnWidth = 11
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 4).HorizontalAlignment = xlLeft
 
    ' Категория
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 5).Value = "Категория"
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns("E:E").EntireColumn.ColumnWidth = 12
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 5).HorizontalAlignment = xlLeft
    
    ' Ядро_ЗП_шт
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 6).Value = "Ядро_ЗП_шт"
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns("F:F").EntireColumn.ColumnWidth = 14
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 6).HorizontalAlignment = xlLeft
    
    ' Клиенты_с_зачислениями_в_теч_последних_6_мес_шт
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 7).Value = "Клиенты_с_зачислениями_в_теч_последних_6_мес_шт"
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns("G:G").EntireColumn.ColumnWidth = 20
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 7).HorizontalAlignment = xlLeft
    
    ' Открытый_ПК_шт
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 8).Value = "Открытый_ПК_шт"
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns("H:H").EntireColumn.ColumnWidth = 19
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 8).HorizontalAlignment = xlLeft
    
    ' Cредняя_ставка
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 9).Value = "Cредняя_ставка"
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns("I:I").EntireColumn.ColumnWidth = 17
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns("I:I").NumberFormat = "0.0%"
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 9).HorizontalAlignment = xlLeft
    
    ' Активные_РА_ПК_шт
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 10).Value = "Активные_РА_ПК_шт"
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns("J:J").EntireColumn.ColumnWidth = 21
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 10).HorizontalAlignment = xlLeft

    ' Активные_CRM_ПК_решение_предложение_шт
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 11).Value = "Активные_CRM_ПК_решение_предложение_шт"
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns("K:K").EntireColumn.ColumnWidth = 21
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 11).HorizontalAlignment = xlLeft

    ' Наличие_КК_шт
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 12).Value = "Наличие_КК_шт"
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns("L:L").EntireColumn.ColumnWidth = 17
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 12).HorizontalAlignment = xlLeft
    
    ' Активные_РА_КК_шт
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 13).Value = "Активные_РА_КК_шт"
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns("M:M").EntireColumn.ColumnWidth = 22
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 13).HorizontalAlignment = xlLeft

    ' Cредний_AR_по_клиентам
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 14).Value = "Cредний_AR_по_клиентам"
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns("N:N").EntireColumn.ColumnWidth = 27
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns("N:N").NumberFormat = "0.0%"
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 14).HorizontalAlignment = xlCenter

    ' Наличие_НС_шт
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 15).Value = "Наличие_НС_шт"
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns("O:O").EntireColumn.ColumnWidth = 17
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 15).HorizontalAlignment = xlLeft

    ' Включить автофильтр
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 1).AutoFilter

    Application.StatusBar = In_OutBookName + " создан!"

End Sub

' Отправка письма: отправляю шаблон самому себе для последующей отправки в сеть письма на его основе:
Sub Отправка_Lotus_Notes_КроссЗП()
Dim темаПисьма, текстПисьма, hashTag, attachmentFile As String
Dim i As Byte
  
  
  ' Запрос
  ' If MsgBox("Отправить себе Шаблон письма с фокусами контроля '" + ПериодКонтроля + "'?", vbYesNo) = vbYes Then
    
    ' Тема письма - Тема:
    темаПисьма = ThisWorkbook.Sheets("КроссЗП").Cells(rowByValue(ThisWorkbook.Name, "КроссЗП", "Тема:", 100, 100), ColumnByValue(ThisWorkbook.Name, "КроссЗП", "Тема:", 100, 100) + 1).Value

    ' hashTag - Хэштэг:
    hashTag = ThisWorkbook.Sheets("КроссЗП").Cells(rowByValue(ThisWorkbook.Name, "КроссЗП", "Хэштэг:", 100, 100), ColumnByValue(ThisWorkbook.Name, "КроссЗП", "Хэштэг:", 100, 100) + 1).Value

    ' Файл-вложение (!!!)
    attachmentFile = ThisWorkbook.Sheets("КроссЗП").Cells(rowByValue(ThisWorkbook.Name, "КроссЗП", "Вложение:", 100, 100), ColumnByValue(ThisWorkbook.Name, "КроссЗП", "Вложение:", 100, 100) + 1).Value
    
    ' Текст письма
    текстПисьма = "" + Chr(13)
    текстПисьма = текстПисьма + "" + getFromAddrBook("УДО2,УДО3,УДО4,УДО5,НОРПиКО1,НОРПиКО2,НОРПиКО3,НОРПиКО4,НОРПиКО5,НОКП,РРКК,МПП", 2) + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "" + getFromAddrBook("РД", 2) + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "Уважаемые руководители," + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + Replace(темаПисьма, "-", ".") + " г." + Chr(13)
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
  
    ' Зачеркнуть
    ' Call ЗачеркиваемТекстВячейке("Лист0", RangeByValue(ThisWorkbook.Name, "Лист0", "DashBoard (при наличии)", 100, 100))
  
    ' Сообщение
    ' MsgBox ("Письмо отправлено!")
     
  ' End If
  
End Sub

