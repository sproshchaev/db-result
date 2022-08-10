Attribute VB_Name = "Module_DB_ЗП"
' *** Лист DB_ЗП ***

' *** Глобальные переменные ***
' Public dateDB_DB_ЗП As Date
' ***                       ***

' Показатели из DB_ЗП
Sub Показатели_из_DB_ЗП()

' Описание переменных
Dim ReportName_String, officeNameInReport, CheckFormatReportResult As String
Dim i, rowCount As Integer
Dim finishProcess As Boolean
    
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
    ThisWorkbook.Sheets("DB_ЗП").Activate

    ' Проверка формы отчета
    CheckFormatReportResult = CheckFormatReport(ReportName_String, "Сотр. прод.", 20, Date)
    
    If CheckFormatReportResult = "OK" Then
      
      ' Дата отчета в "B5" из имени файла "Отчет по ЗП_3кв_2021_29.07.2021.xlsb"
      ' dateDB_DB_ЗП_str = Mid(ReportName_String, InStr(ReportName_String, ".xlsb") - 10, 10)
      ' 22 и 10
      dateDB_DB_ЗП_str = Mid(ReportName_String, 22, 10)
      dateDB_DB_ЗП = CDate(dateDB_DB_ЗП_str)
      ThisWorkbook.Sheets("DB_ЗП").Range("B5").Value = "Привлечение Карты 18+ офисы и ОКП от " + dateDB_DB_ЗП_str
      
      ' Тема письма
      ThisWorkbook.Sheets("DB_ЗП").Range("Q2").Value = "Привлечение ЗП 18+ офисы и ОКП от " + dateDB_DB_ЗП_str
      
      ' Очищаем таблицу "Форма DB_ЗП"
      row_Форма_DB_ЗП = rowByValue(ThisWorkbook.Name, "DB_ЗП", "Форма DB_ЗП", 100, 100)
      Call clearСontents2(ThisWorkbook.Name, "DB_ЗП", "A" + CStr(row_Форма_DB_ЗП + 3), "S" + CStr(row_Форма_DB_ЗП + 3 + 5))
      
      ' Очищаем таблицу "Форма Сотр_прод"
      row_Форма_Сотр_прод = rowByValue(ThisWorkbook.Name, "DB_ЗП", "Форма Сотр_прод", 100, 100)
      Call clearСontents2(ThisWorkbook.Name, "DB_ЗП", "A" + CStr(row_Форма_Сотр_прод + 3), "M" + CStr(row_Форма_Сотр_прод + 3 + 5))
      
      ' Открываем сводную таблицу "ТП (новое привлечение)"
      Workbooks(ReportName_String).Sheets("ТП (новое привлечение)").PivotTables("SASApp:TEMP.ZP_V3_TPL").PivotFields("Филиал").PivotItems("Тюменский ОО1").ShowDetail = True
      
      ' Определяем столбцы на листе "ТП (новое привлечение)"
      Application.StatusBar = "Определение столбцов в ТП (новое привлечение)..."
      row_Филиал_Офис = rowByValue(ReportName_String, "ТП (новое привлечение)", "Филиал - Офис", 50, 50)
      column_Карты_18_План = ColumnByName(ReportName_String, "ТП (новое привлечение)", row_Филиал_Офис, "Карты 18+ План")
      column_Карты_18К_Факт = ColumnByName(ReportName_String, "ТП (новое привлечение)", row_Филиал_Офис, "Карты 18К+ Факт")
      column_Выполнение_плана_Карты_18К = ColumnByName(ReportName_String, "ТП (новое привлечение)", row_Филиал_Офис, "Выполнение плана (Карты 18К+)")
      column_План_по_проникновению_ИБ = ColumnByName(ReportName_String, "ТП (новое привлечение)", row_Филиал_Офис, " План по  проникновению")
      column_Проникновение_ИБ_к_продажам_18 = ColumnByName(ReportName_String, "ТП (новое привлечение)", row_Филиал_Офис, " Проникновение ИБ к продажам 18+")
      column_Выполнение_план_по_проникновению_ИБ = ColumnByName(ReportName_String, "ТП (новое привлечение)", row_Филиал_Офис, " Выполнение план по проникновению ИБ")
      Application.StatusBar = ""

      ' Переходим к обработке листа "Детализация потенциала"
      ' SALORGCD
      column_Детпот_SALORGCD = ColumnByName(ReportName_String, "Детализация потенциала", 1, "SALORGCD")
      ' NAMEORG
      column_Детпот_NAMEORG = ColumnByName(ReportName_String, "Детализация потенциала", 1, "NAMEORG")
      ' INNSORG
      column_Детпот_INNSORG = ColumnByName(ReportName_String, "Детализация потенциала", 1, "INNSORG")
      ' ZCONTRN
      column_Детпот_ZCONTRN = ColumnByName(ReportName_String, "Детализация потенциала", 1, "ZCONTRN")
      ' CNTR_YM_FROM
      column_Детпот_CNTR_YM_FROM = ColumnByName(ReportName_String, "Детализация потенциала", 1, "CNTR_YM_FROM")
      ' Потенциал (НВ)
      column_Детпот_Потенциал_НВ = ColumnByName(ReportName_String, "Детализация потенциала", 1, "Потенциал (НВ)")
      ' Потенциал (ДВ)
      column_Детпот_Потенциал_ДВ = ColumnByName(ReportName_String, "Детализация потенциала", 1, "Потенциал (ДВ)")
      ' Дата выдачи
      column_Детпот_Дата_выдачи = ColumnByName(ReportName_String, "Детализация потенциала", 1, "Дата выдачи")
      ' Филиал
      column_Детпот_Филиал = ColumnByName(ReportName_String, "Детализация потенциала", 1, "Филиал")
      ' ДО
      column_Детпот_ДО = ColumnByName(ReportName_String, "Детализация потенциала", 1, "ДО")
      ' TB_CONTR
      column_Детпот_TB_CONTR = ColumnByName(ReportName_String, "Детализация потенциала", 1, "TB_CONTR")

      ' Установка фильтра на "Детализация потенциала" Внимание! не работает ускоренный поиск
      Workbooks(ReportName_String).Sheets("Детализация потенциала").ListObjects("Таблица2").Range.AutoFilter Field:=column_Детпот_Филиал, Criteria1:="Тюменский ОО1"

      ' Создание исходящей таблицы для карт с потенциалом
      OutBookName = ThisWorkbook.Path + "\Out\Cards_pot_" + Mid(ThisWorkbook.Sheets("DB_ЗП").Range("B5").Value, 38, 10) + ".xlsx"
      
      ' Проверяем - если файл есть, то удаляем его
      Call deleteFile(OutBookName)

      Call createBook_out_DB_ЗП(OutBookName)

      ThisWorkbook.Sheets("DB_ЗП").Range("T3").Value = OutBookName

      ' Обнуление переменных
      Итого_Карты_18_Факт = 0
      Итого_Карты_18_План = 0
      Итого_Портфель_ЗП18_Факт = 0
      Итого_Портфель_ЗП18_План = 0
      Итого_КК_ЗП_План = 0
      Итого_КК_ЗП_Факт = 0

      Итого_РОО_Потенциал_Выд_НВ = 0
      Итого_РОО_Потенциал_Выд_ДВ = 0
      Итого_РОО_Потенциал_Зачисл = 0
      
      ' Обрабатываем отчет
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

        ' Статус
        ' Application.StatusBar = officeNameInReport + "..."

        ThisWorkbook.Sheets("DB_ЗП").Activate

        ' 1) "ТП (новое привлечение)"
        count_Тюменский_OO1 = 0

        rowCount = row_Филиал_Офис + 1
        Do While InStr(Workbooks(ReportName_String).Sheets("ТП (новое привлечение)").Cells(rowCount, 1).Value, "Общий итог") = 0
       
          ' Должен быть второй Тюменский OO1
          If InStr(Workbooks(ReportName_String).Sheets("ТП (новое привлечение)").Cells(rowCount, 1).Value, "Тюменский ОО1") <> 0 And (officeNameInReport = "Тюменский") Then
            count_Тюменский_OO1 = count_Тюменский_OO1 + 1
          End If

          ' Если это текущий офис
          If ((InStr(Workbooks(ReportName_String).Sheets("ТП (новое привлечение)").Cells(rowCount, 1).Value, officeNameInReport) <> 0) And (officeNameInReport <> "Тюменский")) Or ((InStr(Workbooks(ReportName_String).Sheets("ТП (новое привлечение)").Cells(rowCount, 1).Value, "Тюменский") <> 0) And (count_Тюменский_OO1 = 2)) Then
            
            
            ' Выводим данные в Таблицу "Форма DB_ЗП"
            ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_DB_ЗП + 2 + i, 1).Value = CStr(i) + "."
            ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_DB_ЗП + 2 + i, 1).NumberFormat = "@"
            ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_DB_ЗП + 2 + i, 1).HorizontalAlignment = xlCenter
            
            ' Офис
            ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_DB_ЗП + 2 + i, 2).Value = getNameOfficeByNumber(i)
            ' Карты 18+ План
            ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_DB_ЗП + 2 + i, 3).Value = Workbooks(ReportName_String).Sheets("ТП (новое привлечение)").Cells(rowCount, column_Карты_18_План).Value
            ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_DB_ЗП + 2 + i, 3).NumberFormat = "#,##0"
            ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_DB_ЗП + 2 + i, 3).HorizontalAlignment = xlRight
            
            ' Проверяем план из этого отчета и план по ЗП картам на Лист8
            План_офис_Зарплатные_карты_18_Лист8 = ThisWorkbook.Sheets("Лист8").Cells(getRowFromSheet8(getNameOfficeByNumber(i), "Зарплатные карты 18+"), 5).Value
            '
            If План_офис_Зарплатные_карты_18_Лист8 <> ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_DB_ЗП + 2 + i, 3).Value Then
              ' Запрос
              If MsgBox("План ЗП18 Лист8 " + CStr(План_офис_Зарплатные_карты_18_Лист8) + " шт. План DB_ЗП " + CStr(ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_DB_ЗП + 2 + i, 3).Value) + " шт. Внести План с Лист8?", vbYesNo) = vbYes Then
                ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_DB_ЗП + 2 + i, 3).Value = План_офис_Зарплатные_карты_18_Лист8
              End If

            End If
            
            
            ' Сумма плана
            Итого_Карты_18_План = Итого_Карты_18_План + ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_DB_ЗП + 2 + i, 3).Value
            
            ' Карты 18+ Факт
            ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_DB_ЗП + 2 + i, 4).Value = Workbooks(ReportName_String).Sheets("ТП (новое привлечение)").Cells(rowCount, column_Карты_18К_Факт).Value
            ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_DB_ЗП + 2 + i, 4).NumberFormat = "#,##0"
            ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_DB_ЗП + 2 + i, 4).HorizontalAlignment = xlRight
            Итого_Карты_18_Факт = Итого_Карты_18_Факт + ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_DB_ЗП + 2 + i, 4).Value
            
            ' Карты 18+ Исп.
            ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_DB_ЗП + 2 + i, 5).Value = РассчетДоли(ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_DB_ЗП + 2 + i, 3).Value, ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_DB_ЗП + 2 + i, 4).Value, 3) ' Workbooks(ReportName_String).Sheets("ТП (новое привлечение)").Cells(rowCount, column_Выполнение_плана_Карты_18К).Value
            ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_DB_ЗП + 2 + i, 5).NumberFormat = "0%"
            ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_DB_ЗП + 2 + i, 5).HorizontalAlignment = xlRight
            
            ' Карты 18+ Прогноз
            ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_DB_ЗП + 2 + i, 6).Value = Прогноз_квартала_проц(dateDB_DB_ЗП, ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_DB_ЗП + 2 + i, 3).Value, ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_DB_ЗП + 2 + i, 4).Value, 5, 0)
            ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_DB_ЗП + 2 + i, 6).NumberFormat = "0%"
            ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_DB_ЗП + 2 + i, 6).HorizontalAlignment = xlRight
            ' Окраска ячейки СФЕТОФОР ' 70. Заливка ячейки цветом "светофор" - значение In_Value в %, In_Target в %
            Call Full_Color_RangeII("DB_ЗП", row_Форма_DB_ЗП + 2 + i, 6, ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_DB_ЗП + 2 + i, 6).Value, 1)
            
            ' Портфель ЗП План берем из Лист8 (стар. Проникновение ИБ План)
            row_Лист8_officeNameInReport_Портфель_ЗП = getRowFromSheet8(getNameOfficeByNumber(i), "Портфель ЗП 18+, шт._Квартал ")
            ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_DB_ЗП + 2 + i, 7).Value = ThisWorkbook.Sheets("Лист8").Cells(row_Лист8_officeNameInReport_Портфель_ЗП, 5).Value
            ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_DB_ЗП + 2 + i, 7).NumberFormat = "#,##0"
            ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_DB_ЗП + 2 + i, 7).HorizontalAlignment = xlRight
            Итого_Портфель_ЗП18_План = Итого_Портфель_ЗП18_План + ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_DB_ЗП + 2 + i, 7).Value
            
            ' Портфель ЗП Факт (стар. Проникновение ИБ Факт)
            ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_DB_ЗП + 2 + i, 8).Value = ThisWorkbook.Sheets("Лист8").Cells(row_Лист8_officeNameInReport_Портфель_ЗП, 6).Value
            ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_DB_ЗП + 2 + i, 8).NumberFormat = "#,##0"
            ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_DB_ЗП + 2 + i, 8).HorizontalAlignment = xlRight
            Итого_Портфель_ЗП18_Факт = Итого_Портфель_ЗП18_Факт + ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_DB_ЗП + 2 + i, 8).Value
            
            ' Портфель ЗП Исполнение (стар. Проникновение ИБ Исп)
            ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_DB_ЗП + 2 + i, 9).Value = РассчетДоли(ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_DB_ЗП + 2 + i, 7).Value, ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_DB_ЗП + 2 + i, 8).Value, 3)
            ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_DB_ЗП + 2 + i, 9).NumberFormat = "0%"
            ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_DB_ЗП + 2 + i, 9).HorizontalAlignment = xlRight
            ' Окраска ячейки СФЕТОФОР ' 70. Заливка ячейки цветом "светофор" - значение In_Value в %, In_Target в %
            Call Full_Color_RangeII("DB_ЗП", row_Форма_DB_ЗП + 2 + i, 9, ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_DB_ЗП + 2 + i, 9).Value, 1)
            
            ' Сплиты КК к ЗП берем из Лист8
            row_Лист8_officeNameInReport_КК_ЗП = getRowFromSheet8(getNameOfficeByNumber(i), "           КК к ЗП")
                       
            ' Сплиты КК к ЗП план
            ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_DB_ЗП + 2 + i, 10).Value = ThisWorkbook.Sheets("Лист8").Cells(row_Лист8_officeNameInReport_КК_ЗП, 5).Value
            ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_DB_ЗП + 2 + i, 10).NumberFormat = "#,##0"
            ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_DB_ЗП + 2 + i, 10).HorizontalAlignment = xlRight
            Итого_КК_ЗП_План = Итого_КК_ЗП_План + ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_DB_ЗП + 2 + i, 10).Value
            
            ' Сплиты КК к ЗП факт
            ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_DB_ЗП + 2 + i, 11).Value = ThisWorkbook.Sheets("Лист8").Cells(row_Лист8_officeNameInReport_КК_ЗП, 6).Value
            ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_DB_ЗП + 2 + i, 11).NumberFormat = "#,##0"
            ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_DB_ЗП + 2 + i, 11).HorizontalAlignment = xlRight
            Итого_КК_ЗП_Факт = Итого_КК_ЗП_Факт + ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_DB_ЗП + 2 + i, 11).Value
            
            ' Сплиты КК к ЗП исп.
            ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_DB_ЗП + 2 + i, 12).Value = РассчетДоли(ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_DB_ЗП + 2 + i, 10).Value, ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_DB_ЗП + 2 + i, 11).Value, 3)
            ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_DB_ЗП + 2 + i, 12).NumberFormat = "0%"
            ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_DB_ЗП + 2 + i, 12).HorizontalAlignment = xlRight

            ' Сплиты КК к ЗП прогноз.
            ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_DB_ЗП + 2 + i, 13).Value = Прогноз_квартала_проц(dateDB_DB_ЗП, ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_DB_ЗП + 2 + i, 10).Value, ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_DB_ЗП + 2 + i, 11).Value, 5, 0)
            ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_DB_ЗП + 2 + i, 13).NumberFormat = "0%"
            ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_DB_ЗП + 2 + i, 13).HorizontalAlignment = xlRight
            ' Окраска ячейки СФЕТОФОР ' 70. Заливка ячейки цветом "светофор" - значение In_Value в %, In_Target в %
            Call Full_Color_RangeII("DB_ЗП", row_Форма_DB_ЗП + 2 + i, 13, ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_DB_ЗП + 2 + i, 13).Value, 1)

            ' =======================================================================================================================================
            ' Переходим к обработке листа "Детализация потенциала"
            
            ' Обнуление переменных
            Итого_Потенциал_НВ = 0
            Итого_Потенциал_ДВ = 0
            ' Дата выдачи
            Итого_выданных_без_зачислений = 0

            
            rowCount_Детпот = 2
            Do While Workbooks(ReportName_String).Sheets("Детализация потенциала").Cells(rowCount_Детпот, 1).Value <> ""
       
              ' Если это текущий офис
              If (InStr(Workbooks(ReportName_String).Sheets("Детализация потенциала").Cells(rowCount_Детпот, column_Детпот_ДО).Value, officeNameInReport) <> 0) Then
                
                
                ' Дата выдачи "00000000" (текст)
                If Workbooks(ReportName_String).Sheets("Детализация потенциала").Cells(rowCount_Детпот, column_Детпот_Дата_выдачи).Value = "00000000" Then
                  
                  ' Потенциал НВ выпуска
                  If Workbooks(ReportName_String).Sheets("Детализация потенциала").Cells(rowCount_Детпот, column_Детпот_Потенциал_НВ).Value = "1" Then
                    Итого_Потенциал_НВ = Итого_Потенциал_НВ + 1
                    Итого_РОО_Потенциал_Выд_НВ = Итого_РОО_Потенциал_Выд_НВ + 1
                    Status_Var = "Не выдана (НВ)"
                  End If
                  
                  ' Потенциал ДВ выпуска
                  If Workbooks(ReportName_String).Sheets("Детализация потенциала").Cells(rowCount_Детпот, column_Детпот_Потенциал_ДВ).Value = "1" Then
                    Итого_Потенциал_ДВ = Итого_Потенциал_ДВ + 1
                    Итого_РОО_Потенциал_Выд_ДВ = Итого_РОО_Потенциал_Выд_ДВ + 1
                    Status_Var = "Не выдана (ДВ)"
                  End If
                  
                Else
                  ' Если дата выдачи не "00000000" (текст)
                  Итого_выданных_без_зачислений = Итого_выданных_без_зачислений + 1
                  Итого_РОО_Потенциал_Зачисл = Итого_РОО_Потенциал_Зачисл + 1
                  Status_Var = "Выдана без зачислений"
                End If
                
                ' Вставляем карту в OutBookName
                ' Вносим данные в BASE\Sales_Office по ПК.
                Call InsertRecordInBook(Dir(OutBookName), "Лист1", "TB_CONTR", Workbooks(ReportName_String).Sheets("Детализация потенциала").Cells(rowCount_Детпот, column_Детпот_TB_CONTR).Value, _
                                            "TB_CONTR", Workbooks(ReportName_String).Sheets("Детализация потенциала").Cells(rowCount_Детпот, column_Детпот_TB_CONTR).Value, _
                                              "SALORGCD", Workbooks(ReportName_String).Sheets("Детализация потенциала").Cells(rowCount_Детпот, column_Детпот_SALORGCD).Value, _
                                                "NAMEORG", Workbooks(ReportName_String).Sheets("Детализация потенциала").Cells(rowCount_Детпот, column_Детпот_NAMEORG).Value, _
                                                  "INNSORG", Workbooks(ReportName_String).Sheets("Детализация потенциала").Cells(rowCount_Детпот, column_Детпот_INNSORG).Value, _
                                                    "ZCONTRN", Workbooks(ReportName_String).Sheets("Детализация потенциала").Cells(rowCount_Детпот, column_Детпот_ZCONTRN).Value, _
                                                      "CNTR_YM_FROM", Workbooks(ReportName_String).Sheets("Детализация потенциала").Cells(rowCount_Детпот, column_Детпот_CNTR_YM_FROM).Value, _
                                                        "Status", Status_Var, _
                                                          "Дата выдачи", Workbooks(ReportName_String).Sheets("Детализация потенциала").Cells(rowCount_Детпот, column_Детпот_Дата_выдачи).Value, _
                                                            "Офис", Workbooks(ReportName_String).Sheets("Детализация потенциала").Cells(rowCount_Детпот, column_Детпот_ДО).Value, _
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
              rowCount_Детпот = rowCount_Детпот + 1
              Application.StatusBar = officeNameInReport + ": Детализация потенциала " + CStr(rowCount_Детпот) + "..."
              DoEventsInterval (rowCount_Детпот)
            Loop
            
            ' Выводим итоги по листу "Детализация потенциала"
            ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_DB_ЗП + 2 + i, 17).Value = Итого_Потенциал_НВ
            ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_DB_ЗП + 2 + i, 17).NumberFormat = "#,##0"
            ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_DB_ЗП + 2 + i, 17).HorizontalAlignment = xlRight
            ' Итого_КК_ЗП_Факт = Итого_КК_ЗП_Факт + ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_DB_ЗП + 2 + i, 11).Value

            ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_DB_ЗП + 2 + i, 18).Value = Итого_Потенциал_ДВ
            ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_DB_ЗП + 2 + i, 18).NumberFormat = "#,##0"
            ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_DB_ЗП + 2 + i, 18).HorizontalAlignment = xlRight
            ' Итого_КК_ЗП_Факт = Итого_КК_ЗП_Факт + ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_DB_ЗП + 2 + i, 11).Value

            ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_DB_ЗП + 2 + i, 19).Value = Итого_выданных_без_зачислений
            ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_DB_ЗП + 2 + i, 19).NumberFormat = "#,##0"
            ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_DB_ЗП + 2 + i, 19).HorizontalAlignment = xlRight
            ' Итого_КК_ЗП_Факт = Итого_КК_ЗП_Факт + ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_DB_ЗП + 2 + i, 11).Value

            
            ' =======================================================================================================================================
            
            
          End If ' Если это текущий офис
        
        
          ' Следующая запись
          rowCount = rowCount + 1
          Application.StatusBar = officeNameInReport + ": " + CStr(rowCount) + "..."
          DoEventsInterval (rowCount)
        Loop
   
        ' Выводим данные по офису
      
      Next i ' Следующий офис
      
      ' Выводим итоги обработки
      ' ----------------------------------------------------------------------------------------------------------------------------------
      ' Чертим горизонтальную линию 2 (указываем предидущее значение строки + 1)
      Call gorizontalLineII(ThisWorkbook.Name, "DB_ЗП", row_Форма_DB_ЗП + 2 + i, 2, 19)
      
      ' Итого Карты 18+ План
      ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_DB_ЗП + 2 + i, 3).Value = Итого_Карты_18_План
      ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_DB_ЗП + 2 + i, 3).NumberFormat = "#,##0"
      ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_DB_ЗП + 2 + i, 3).HorizontalAlignment = xlRight
      Call Полужирный_текст(ThisWorkbook.Name, "DB_ЗП", row_Форма_DB_ЗП + 2 + i, 3)
      
      ' Итого Карты 18+ Факт
      ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_DB_ЗП + 2 + i, 4).Value = Итого_Карты_18_Факт
      ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_DB_ЗП + 2 + i, 4).NumberFormat = "#,##0"
      ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_DB_ЗП + 2 + i, 4).HorizontalAlignment = xlRight
      Call Полужирный_текст(ThisWorkbook.Name, "DB_ЗП", row_Форма_DB_ЗП + 2 + i, 4)
      
      ' Итого Карты 18+ исп.
      ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_DB_ЗП + 2 + i, 5).Value = РассчетДоли(Итого_Карты_18_План, Итого_Карты_18_Факт, 3)
      ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_DB_ЗП + 2 + i, 5).NumberFormat = "0%"
      ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_DB_ЗП + 2 + i, 5).HorizontalAlignment = xlRight
      Call Полужирный_текст(ThisWorkbook.Name, "DB_ЗП", row_Форма_DB_ЗП + 2 + i, 5)
      
      ' Итого Карты 18+ прогноз.
      ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_DB_ЗП + 2 + i, 6).Value = Прогноз_квартала_проц(dateDB_DB_ЗП, Итого_Карты_18_План, Итого_Карты_18_Факт, 5, 0)
      ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_DB_ЗП + 2 + i, 6).NumberFormat = "0%"
      ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_DB_ЗП + 2 + i, 6).HorizontalAlignment = xlRight
      Call Полужирный_текст(ThisWorkbook.Name, "DB_ЗП", row_Форма_DB_ЗП + 2 + i, 6)
      
      ' Портфель ЗП План
      ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_DB_ЗП + 2 + i, 7).Value = Итого_Портфель_ЗП18_План
      ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_DB_ЗП + 2 + i, 7).NumberFormat = "#,##0"
      ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_DB_ЗП + 2 + i, 7).HorizontalAlignment = xlRight
      Call Полужирный_текст(ThisWorkbook.Name, "DB_ЗП", row_Форма_DB_ЗП + 2 + i, 7)
      
      ' Портфель ЗП Факт
      ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_DB_ЗП + 2 + i, 8).Value = Итого_Портфель_ЗП18_Факт
      ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_DB_ЗП + 2 + i, 8).NumberFormat = "#,##0"
      ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_DB_ЗП + 2 + i, 8).HorizontalAlignment = xlRight
      Call Полужирный_текст(ThisWorkbook.Name, "DB_ЗП", row_Форма_DB_ЗП + 2 + i, 8)
      
      ' Портфель ЗП Исполнение
      ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_DB_ЗП + 2 + i, 9).Value = РассчетДоли(Итого_Портфель_ЗП18_План, Итого_Портфель_ЗП18_Факт, 3)
      ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_DB_ЗП + 2 + i, 9).NumberFormat = "0%"
      ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_DB_ЗП + 2 + i, 9).HorizontalAlignment = xlRight
      Call Полужирный_текст(ThisWorkbook.Name, "DB_ЗП", row_Форма_DB_ЗП + 2 + i, 5)
      
      ' КК к ЗП План
      ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_DB_ЗП + 2 + i, 10).Value = Итого_КК_ЗП_План
      ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_DB_ЗП + 2 + i, 10).NumberFormat = "#,##0"
      ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_DB_ЗП + 2 + i, 10).HorizontalAlignment = xlRight
      Call Полужирный_текст(ThisWorkbook.Name, "DB_ЗП", row_Форма_DB_ЗП + 2 + i, 10)
                  
      ' КК к ЗП Факт
      ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_DB_ЗП + 2 + i, 11).Value = Итого_КК_ЗП_Факт
      ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_DB_ЗП + 2 + i, 11).NumberFormat = "#,##0"
      ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_DB_ЗП + 2 + i, 11).HorizontalAlignment = xlRight
      Call Полужирный_текст(ThisWorkbook.Name, "DB_ЗП", row_Форма_DB_ЗП + 2 + i, 11)
      
      ' КК к ЗП исп.
      ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_DB_ЗП + 2 + i, 12).Value = РассчетДоли(Итого_КК_ЗП_План, Итого_КК_ЗП_Факт, 3)
      ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_DB_ЗП + 2 + i, 12).NumberFormat = "0%"
      ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_DB_ЗП + 2 + i, 12).HorizontalAlignment = xlRight
      Call Полужирный_текст(ThisWorkbook.Name, "DB_ЗП", row_Форма_DB_ЗП + 2 + i, 12)

      ' КК к ЗП прогноз.
      ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_DB_ЗП + 2 + i, 13).Value = Прогноз_квартала_проц(dateDB_DB_ЗП, Итого_КК_ЗП_План, Итого_КК_ЗП_Факт, 5, 0)
      ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_DB_ЗП + 2 + i, 13).NumberFormat = "0%"
      ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_DB_ЗП + 2 + i, 13).HorizontalAlignment = xlRight
      Call Полужирный_текст(ThisWorkbook.Name, "DB_ЗП", row_Форма_DB_ЗП + 2 + i, 13)

      ' 17
      ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_DB_ЗП + 2 + i, 17).Value = Итого_РОО_Потенциал_Выд_НВ
      ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_DB_ЗП + 2 + i, 17).NumberFormat = "#,##0"
      ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_DB_ЗП + 2 + i, 17).HorizontalAlignment = xlRight
      Call Полужирный_текст(ThisWorkbook.Name, "DB_ЗП", row_Форма_DB_ЗП + 2 + i, 17)
      
      ' 18
      ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_DB_ЗП + 2 + i, 18).Value = Итого_РОО_Потенциал_Выд_ДВ
      ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_DB_ЗП + 2 + i, 18).NumberFormat = "#,##0"
      ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_DB_ЗП + 2 + i, 18).HorizontalAlignment = xlRight
      Call Полужирный_текст(ThisWorkbook.Name, "DB_ЗП", row_Форма_DB_ЗП + 2 + i, 18)
            
      ' 19
      ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_DB_ЗП + 2 + i, 19).Value = Итого_РОО_Потенциал_Зачисл
      ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_DB_ЗП + 2 + i, 19).NumberFormat = "#,##0"
      ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_DB_ЗП + 2 + i, 19).HorizontalAlignment = xlRight
      Call Полужирный_текст(ThisWorkbook.Name, "DB_ЗП", row_Форма_DB_ЗП + 2 + i, 19)
      

      ' 2. Обработка листа "Сотр. прод."
      ' Определяем столбцы на листе "Сотр. прод."
      Application.StatusBar = "Определение столбцов в " + кавычки() + "Сотр. прод." + кавычки() + "..."
      row_Сотрпрод_Сотрудник = rowByValue(ReportName_String, "Сотр. прод.", "Сотрудник", 50, 50)
      column_Сотрпрод_Сотрудник = ColumnByName(ReportName_String, "Сотр. прод.", row_Сотрпрод_Сотрудник, "Сотрудник")
      column_Сотрпрод_Филиал = ColumnByName(ReportName_String, "Сотр. прод.", row_Сотрпрод_Сотрудник, "Филиал")
      column_Сотрпрод_Офис = ColumnByName(ReportName_String, "Сотр. прод.", row_Сотрпрод_Сотрудник, "Офис")
      column_Сотрпрод_План = ColumnByName(ReportName_String, "Сотр. прод.", row_Сотрпрод_Сотрудник, " План")
      column_Сотрпрод_Факт_18 = ColumnByName(ReportName_String, "Сотр. прод.", row_Сотрпрод_Сотрудник, " Факт 18+")
      column_Сотрпрод_Выполнение_плана_18 = ColumnByName(ReportName_String, "Сотр. прод.", row_Сотрпрод_Сотрудник, "Выполнение плана 18+")
      ' column_Сотрпрод_ = ColumnByName(ReportName_String, "Сотр. прод.", row_Сотрпрод_Сотрудник, "")
      Application.StatusBar = ""

      ' Открываем сводную таблицу - в Лист1 должен открыться
      row_Сотрпрод_Общий_итог = rowByValue(ReportName_String, "Сотр. прод.", "Общий итог", 1000, 50)
      Workbooks(ReportName_String).Sheets("Сотр. прод.").Cells(row_Сотрпрод_Общий_итог, column_Сотрпрод_Факт_18).ShowDetail = True

      ThisWorkbook.Sheets("DB_ЗП").Activate

      ' Определяем столбцы на листе "Сотр. прод. Лист1"
      Application.StatusBar = "Определение столбцов в " + кавычки() + "Сотр. прод. Лист1" + кавычки() + "..."
      column_Сотрпрод_Лист1_Филиал = ColumnByName(ReportName_String, "Лист1", 1, "Филиал")
      column_Сотрпрод_Лист1_Сотрудник = ColumnByName(ReportName_String, "Лист1", 1, "Сотрудник")
      column_Сотрпрод_Лист1_План = ColumnByName(ReportName_String, "Лист1", 1, "План")
      column_Сотрпрод_Лист1_15_НВ = ColumnByName(ReportName_String, "Лист1", 1, "15+ НВ")
      column_Сотрпрод_Лист1_15_ДВ = ColumnByName(ReportName_String, "Лист1", 1, "15+ ДВ")
      column_Сотрпрод_Лист1_15_КДВ = ColumnByName(ReportName_String, "Лист1", 1, "15+ КДВ!")
      
      column_Сотрпрод_Лист1_План_актив_ИБ = ColumnByName(ReportName_String, "Лист1", 1, "План актив ИБ")
      column_Сотрпрод_Лист1_IB_ACT = ColumnByName(ReportName_String, "Лист1", 1, "IB_ACT")
      
      ' column_Сотрпрод_Лист1_ = ColumnByName(ReportName_String, "Лист1", 1, "")
      
      ' column_Сотрпрод_Лист1 = ColumnByName(ReportName_String, "Лист1", 1, "")
      Application.StatusBar = ""
      
      ' По РРКК
      Номер_пункта = 0
      Итого_Карты_18_РРКК_План = 0
      Итого_Карты_18_РРКК_Факт = 0
      Итого_IB_РРКК_План = 0
      Итого_IB_РРКК_Факт = 0
      
      rowCount = 2
      Do While Workbooks(ReportName_String).Sheets("Лист1").Cells(rowCount, 1).Value <> ""
        
        ' Если это Тюменский ОО1, то выводим в таблицу на листе "DB_ЗП" данные по сотруднику
        If InStr(Workbooks(ReportName_String).Sheets("Лист1").Cells(rowCount, column_Сотрпрод_Лист1_Филиал).Value, "Тюменский ОО1") <> 0 Then
          
          ' №
          Номер_пункта = Номер_пункта + 1
          ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_Сотр_прод + 2 + Номер_пункта, 1).Value = CStr(Номер_пункта) + "."
          ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_Сотр_прод + 2 + Номер_пункта, 1).NumberFormat = "@"
          ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_Сотр_прод + 2 + Номер_пункта, 1).HorizontalAlignment = xlCenter

          ' Сотрудник
          ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_Сотр_прод + 2 + Номер_пункта, 2).Value = Фамилия_и_Имя(Workbooks(ReportName_String).Sheets("Лист1").Cells(rowCount, column_Сотрпрод_Лист1_Сотрудник).Value, 3)
          
          ' Карты 18+ План РРКК
          ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_Сотр_прод + 2 + Номер_пункта, 3).Value = Workbooks(ReportName_String).Sheets("Лист1").Cells(rowCount, column_Сотрпрод_Лист1_План).Value
          ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_Сотр_прод + 2 + Номер_пункта, 3).NumberFormat = "#,##0"
          ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_Сотр_прод + 2 + Номер_пункта, 3).HorizontalAlignment = xlRight
          Итого_Карты_18_РРКК_План = Итого_Карты_18_РРКК_План + ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_Сотр_прод + 2 + Номер_пункта, 3).Value

          ' Карты 18+ Факт РРКК
          ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_Сотр_прод + 2 + Номер_пункта, 4).Value = Workbooks(ReportName_String).Sheets("Лист1").Cells(rowCount, column_Сотрпрод_Лист1_15_НВ).Value + _
                                                                                                  Workbooks(ReportName_String).Sheets("Лист1").Cells(rowCount, column_Сотрпрод_Лист1_15_ДВ).Value + _
                                                                                                    Workbooks(ReportName_String).Sheets("Лист1").Cells(rowCount, column_Сотрпрод_Лист1_15_КДВ).Value
          ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_Сотр_прод + 2 + Номер_пункта, 4).NumberFormat = "#,##0"
          ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_Сотр_прод + 2 + Номер_пункта, 4).HorizontalAlignment = xlRight
          Итого_Карты_18_РРКК_Факт = Итого_Карты_18_РРКК_Факт + ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_Сотр_прод + 2 + Номер_пункта, 4).Value
          
          ' Карты 18+ Исп. РРКК
          ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_Сотр_прод + 2 + Номер_пункта, 5).Value = РассчетДоли(ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_Сотр_прод + 2 + Номер_пункта, 3).Value, ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_Сотр_прод + 2 + Номер_пункта, 4).Value, 3)
          ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_Сотр_прод + 2 + Номер_пункта, 5).NumberFormat = "0%"
          ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_Сотр_прод + 2 + Номер_пункта, 5).HorizontalAlignment = xlRight

          ' Карты 18+ Прогноз. РРКК
          ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_Сотр_прод + 2 + Номер_пункта, 6).Value = Прогноз_квартала_проц(dateDB_DB_ЗП, ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_Сотр_прод + 2 + Номер_пункта, 3).Value, ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_Сотр_прод + 2 + Номер_пункта, 4).Value, 5, 0)
          ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_Сотр_прод + 2 + Номер_пункта, 6).NumberFormat = "0%"
          ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_Сотр_прод + 2 + Номер_пункта, 6).HorizontalAlignment = xlRight
          ' Окраска ячейки СФЕТОФОР ' 70. Заливка ячейки цветом "светофор" - значение In_Value в %, In_Target в %
          Call Full_Color_RangeII("DB_ЗП", row_Форма_Сотр_прод + 2 + Номер_пункта, 6, ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_Сотр_прод + 2 + Номер_пункта, 6).Value, 1)


          ' План IB РРКК
          ' 90% от выданных карт должно быть с ИБ
          ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_Сотр_прод + 2 + Номер_пункта, 7).Value = (ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_Сотр_прод + 2 + Номер_пункта, 4).Value / 100) * 90 ' Workbooks(ReportName_String).Sheets("Лист1").Cells(rowCount, column_Сотрпрод_Лист1_План_актив_ИБ).Value
          ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_Сотр_прод + 2 + Номер_пункта, 7).NumberFormat = "#,##0"
          ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_Сотр_прод + 2 + Номер_пункта, 7).HorizontalAlignment = xlRight
          Итого_IB_РРКК_План = Итого_IB_РРКК_План + ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_Сотр_прод + 2 + Номер_пункта, 7).Value

          ' Факт IB РРКК
          ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_Сотр_прод + 2 + Номер_пункта, 8).Value = Workbooks(ReportName_String).Sheets("Лист1").Cells(rowCount, column_Сотрпрод_Лист1_IB_ACT).Value
          ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_Сотр_прод + 2 + Номер_пункта, 8).NumberFormat = "#,##0"
          ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_Сотр_прод + 2 + Номер_пункта, 8).HorizontalAlignment = xlRight
          Итого_IB_РРКК_Факт = Итого_IB_РРКК_Факт + ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_Сотр_прод + 2 + Номер_пункта, 8).Value

          ' Исп IB РРКК
          ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_Сотр_прод + 2 + Номер_пункта, 9).Value = РассчетДоли(ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_Сотр_прод + 2 + Номер_пункта, 7).Value, ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_Сотр_прод + 2 + Номер_пункта, 8).Value, 3)
          ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_Сотр_прод + 2 + Номер_пункта, 9).NumberFormat = "0%"
          ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_Сотр_прод + 2 + Номер_пункта, 9).HorizontalAlignment = xlRight
          
        End If
        
        ' Следующая запись
        rowCount = rowCount + 1
        Application.StatusBar = "Сотр. прод.: " + CStr(rowCount) + "..."
        DoEventsInterval (rowCount)
      Loop

      ' Выводим итоги обработки
      ' ----------------------------------------------------------------------------------------------------------------------------------
      ' Чертим горизонтальную линию 2 (указываем предидущее значение строки + 1)
      Call gorizontalLineII(ThisWorkbook.Name, "DB_ЗП", row_Форма_Сотр_прод + 2 + Номер_пункта + 1, 2, 13)

      ' Итого Карты 18+ План
      ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_Сотр_прод + 2 + Номер_пункта + 1, 3).Value = Итого_Карты_18_РРКК_План
      ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_Сотр_прод + 2 + Номер_пункта + 1, 3).NumberFormat = "#,##0"
      ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_Сотр_прод + 2 + Номер_пункта + 1, 3).HorizontalAlignment = xlRight
      Call Полужирный_текст(ThisWorkbook.Name, "DB_ЗП", row_Форма_Сотр_прод + 2 + Номер_пункта + 1, 3)

      ' Итого Карты 18+ Факт
      ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_Сотр_прод + 2 + Номер_пункта + 1, 4).Value = Итого_Карты_18_РРКК_Факт
      ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_Сотр_прод + 2 + Номер_пункта + 1, 4).NumberFormat = "#,##0"
      ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_Сотр_прод + 2 + Номер_пункта + 1, 4).HorizontalAlignment = xlRight
      Call Полужирный_текст(ThisWorkbook.Name, "DB_ЗП", row_Форма_Сотр_прод + 2 + Номер_пункта + 1, 4)

      ' Итого Карты 18+ Исп
      ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_Сотр_прод + 2 + Номер_пункта + 1, 5).Value = РассчетДоли(Итого_Карты_18_РРКК_План, Итого_Карты_18_РРКК_Факт, 3)
      ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_Сотр_прод + 2 + Номер_пункта + 1, 5).NumberFormat = "0%"
      ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_Сотр_прод + 2 + Номер_пункта + 1, 5).HorizontalAlignment = xlRight
      Call Полужирный_текст(ThisWorkbook.Name, "DB_ЗП", row_Форма_Сотр_прод + 2 + Номер_пункта + 1, 5)

      ' Итого Карты 18+ Прогноз
      ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_Сотр_прод + 2 + Номер_пункта + 1, 6).Value = Прогноз_квартала_проц(dateDB_DB_ЗП, Итого_Карты_18_РРКК_План, Итого_Карты_18_РРКК_Факт, 5, 0)
      ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_Сотр_прод + 2 + Номер_пункта + 1, 6).NumberFormat = "0%"
      ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_Сотр_прод + 2 + Номер_пункта + 1, 6).HorizontalAlignment = xlRight
      Call Полужирный_текст(ThisWorkbook.Name, "DB_ЗП", row_Форма_Сотр_прод + 2 + Номер_пункта + 1, 6)

      ' Итого ИБ РРКК План
      ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_Сотр_прод + 2 + Номер_пункта + 1, 7).Value = Итого_IB_РРКК_План
      ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_Сотр_прод + 2 + Номер_пункта + 1, 7).NumberFormat = "#,##0"
      ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_Сотр_прод + 2 + Номер_пункта + 1, 7).HorizontalAlignment = xlRight
      Call Полужирный_текст(ThisWorkbook.Name, "DB_ЗП", row_Форма_Сотр_прод + 2 + Номер_пункта + 1, 7)

      ' Итого ИБ РРКК Факт
      ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_Сотр_прод + 2 + Номер_пункта + 1, 8).Value = Итого_IB_РРКК_Факт
      ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_Сотр_прод + 2 + Номер_пункта + 1, 8).NumberFormat = "#,##0"
      ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_Сотр_прод + 2 + Номер_пункта + 1, 8).HorizontalAlignment = xlRight
      Call Полужирный_текст(ThisWorkbook.Name, "DB_ЗП", row_Форма_Сотр_прод + 2 + Номер_пункта + 1, 8)

      ' Итого ИБ РРКК Исп
      ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_Сотр_прод + 2 + Номер_пункта + 1, 9).Value = РассчетДоли(Итого_IB_РРКК_План, Итого_IB_РРКК_Факт, 3)
      ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_Сотр_прод + 2 + Номер_пункта + 1, 9).NumberFormat = "0%"
      ThisWorkbook.Sheets("DB_ЗП").Cells(row_Форма_Сотр_прод + 2 + Номер_пункта + 1, 9).HorizontalAlignment = xlRight
      Call Полужирный_текст(ThisWorkbook.Name, "DB_ЗП", row_Форма_Сотр_прод + 2 + Номер_пункта + 1, 9)
 
      ' ================================================================================
      ' 3. Выгрузка потенциала по выдаче и зачислениям из листа "Детализация потенциала"
      ' Определяем столбцы "Филиал"
      ' Определяем столбцы "ДО"
      ' Определяем столбцы "INNSORG"
      ' Определяем столбцы "NAMEORG"
      ' Определяем столбцы "ZCONTRN" - ФИО клиента
      ' - Потенциал (НВ)
      ' - Потенциал (ДВ)
      ' - Дата выдачи
      
      
      
             
      ' ================================================================================
 
      ' Сохранение изменений
      ThisWorkbook.Save
    
      ' Копирование в таблицу для отправки
      ' Call copy_DB_ЗП_ToSend
    
      ' Отправка шаблона письма
      ' Call Отправка_Lotus_Notes_DB_ЗП
    
      ' Закрываем выходную книгу с выгрузкой
      Workbooks(Dir(OutBookName)).Close SaveChanges:=True
    
      ' Переменная завершения обработки
      finishProcess = True
    Else
      ' Сообщение о неверном формате отчета или даты
      MsgBox ("Проверьте отчет: " + CheckFormatReportResult + "!")
    End If ' Проверка формы отчета
    

    ' Закрываем файл с отчетом без сохранения изменений (параметр SaveChanges:=False)
    Workbooks(Dir(FileName)).Close SaveChanges:=False
    
    ' Переходим в ячейку M2
    ThisWorkbook.Sheets("DB_ЗП").Range("A1").Select

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


' Создание файла для отправки в офисы
Sub copy_DB_ЗП_ToSend()
Dim TemplatesFile As String

  Application.StatusBar = "Копирование..."

  ' Открываем шаблон "Отчет по ЗП_Nкв_YYYY_DD-MM-YYYY.xlsx"
  If Dir(ThisWorkbook.Path + "\Templates\" + "Отчет по ЗП_Nкв_YYYY_DD-MM-YYYY.xlsx") <> "" Then
    ' Открываем шаблон Templates\Ежедневный отчет по продажам
    TemplatesFileName = "Отчет по ЗП_Nкв_YYYY_DD-MM-YYYY"
  End If
              
  ' Открываем шаблон Протокола из C:\Users\...\Documents\#VBA\DB_Result\Templates
  Workbooks.Open (ThisWorkbook.Path + "\Templates\" + TemplatesFileName + ".xlsx")
           
  ' Переходим на окно DB
  ThisWorkbook.Sheets("DB_ЗП").Activate

  ' Обновляем список получателей
  ThisWorkbook.Sheets("DB_ЗП").Cells(rowByValue(ThisWorkbook.Name, "DB_ЗП", "Список получателей:", 100, 100), ColumnByValue(ThisWorkbook.Name, "DB_ЗП", "Список получателей:", 100, 100) + 2).Value = _
    getFromAddrBook("УДО2,УДО3,УДО4,УДО5,НОКП,РРКК,МПП", 2)

  ' Дата отчета в "B5" из имени файла "Отчет по ЗП_3кв_2021_29.07.2021.xlsb"
  dateDB_DB_ЗП = CDate(Mid(ThisWorkbook.Sheets("DB_ЗП").Range("B5").Value, 38, 10))

  ' Имя нового файла
  FileDBName = "Отчет по ЗП_" + quarterName3(dateDB_DB_ЗП) + " от " + strДД_MM_YYYY(dateDB_DB_ЗП) + ".xlsx"
  
  ' Проверяем - если файл есть, то удаляем его
  Call deleteFile(ThisWorkbook.Path + "\Out\" + FileDBName)
  
  Workbooks(TemplatesFileName).SaveAs FileName:=ThisWorkbook.Path + "\Out\" + FileDBName, FileFormat:=xlOpenXMLWorkbook, createBackUp:=False
  ThisWorkbook.Sheets("DB_ЗП").Range("R3").Value = ThisWorkbook.Path + "\Out\" + FileDBName
            
  ' *** Копирование данных ***
 
  ' Заголовок
  ThisWorkbook.Sheets("DB_ЗП").Cells(5, 2).Copy Destination:=Workbooks(FileDBName).Sheets("Лист1").Cells(5, 2)
 
  ' Шапки таблиц (1)
  ' 1
  Workbooks(FileDBName).Sheets("Лист1").Cells(7, 1).Value = ThisWorkbook.Sheets("DB_ЗП").Cells(7, 1).Value
  ' 2
  Workbooks(FileDBName).Sheets("Лист1").Cells(7, 2).Value = ThisWorkbook.Sheets("DB_ЗП").Cells(7, 2).Value
  ' 3
  Workbooks(FileDBName).Sheets("Лист1").Cells(7, 3).Value = ThisWorkbook.Sheets("DB_ЗП").Cells(7, 3).Value
  ' 7
  Workbooks(FileDBName).Sheets("Лист1").Cells(7, 7).Value = ThisWorkbook.Sheets("DB_ЗП").Cells(7, 7).Value
  ' 10
  Workbooks(FileDBName).Sheets("Лист1").Cells(7, 10).Value = ThisWorkbook.Sheets("DB_ЗП").Cells(7, 10).Value
  ' 14
  Workbooks(FileDBName).Sheets("Лист1").Cells(7, 14).Value = ThisWorkbook.Sheets("DB_ЗП").Cells(7, 14).Value
  
  ' Шапки таблиц (2)
  ' 1
  Workbooks(FileDBName).Sheets("Лист1").Cells(19, 1).Value = ThisWorkbook.Sheets("DB_ЗП").Cells(19, 1).Value
  ' 2
  Workbooks(FileDBName).Sheets("Лист1").Cells(19, 2).Value = ThisWorkbook.Sheets("DB_ЗП").Cells(19, 2).Value
  ' 3
  Workbooks(FileDBName).Sheets("Лист1").Cells(19, 3).Value = ThisWorkbook.Sheets("DB_ЗП").Cells(19, 3).Value
  ' 7
  Workbooks(FileDBName).Sheets("Лист1").Cells(19, 7).Value = ThisWorkbook.Sheets("DB_ЗП").Cells(19, 7).Value
  ' 10
  Workbooks(FileDBName).Sheets("Лист1").Cells(19, 10).Value = ThisWorkbook.Sheets("DB_ЗП").Cells(19, 10).Value
  

  ' Офисы
  For i = 9 To 14
    
    For j = 1 To 16
      ThisWorkbook.Sheets("DB_ЗП").Cells(i, j).Copy Destination:=Workbooks(FileDBName).Sheets("Лист1").Cells(i, j)
      Application.StatusBar = "Копирование 1: " + CStr(i) + "-" + CStr(j) + "..."
    Next j
  
  Next i
  
  ' РРКК
  For i = 21 To 24
    
    For j = 1 To 13
      ThisWorkbook.Sheets("DB_ЗП").Cells(i, j).Copy Destination:=Workbooks(FileDBName).Sheets("Лист1").Cells(i, j)
      Application.StatusBar = "Копирование 2: " + CStr(i) + "-" + CStr(j) + "..."
    Next j
  
  Next i
  
  
  ' ***
                    
  ' Закрытие файла
  Workbooks(FileDBName).Close SaveChanges:=True

  ' Копирование завершено
  Application.StatusBar = "Скопировано!"
  Application.StatusBar = ""

End Sub


' Отправка письма: отправляю шаблон самому себе для последующей отправки в сеть письма на его основе:
Sub Отправка_Lotus_Notes_DB_ЗП()
Dim темаПисьма, текстПисьма, hashTag, attachmentFile As String
Dim i As Byte
  
  ' Строка статуса
  Application.StatusBar = "Отправка письма ..."
  
  ' Запрос
  ' If MsgBox("Отправить себе Шаблон письма с фокусами контроля '" + ПериодКонтроля + "'?", vbYesNo) = vbYes Then
    
    ' Тема письма - Тема:
    ' темаПисьма = ThisWorkbook.Sheets("DB_ЗП").Cells(RowByValue(ThisWorkbook.Name, "DB_ЗП", "Тема:", 100, 100), ColumnByValue(ThisWorkbook.Name, "DB_ЗП", "Тема:", 100, 100) + 1).Value
    темаПисьма = subjectFromSheet("DB_ЗП")

    ' hashTag - Хэштэг:
    ' hashTag = ThisWorkbook.Sheets("DB_ЗП").Cells(RowByValue(ThisWorkbook.Name, "DB_ЗП", "Хэштэг:", 100, 100), ColumnByValue(ThisWorkbook.Name, "DB_ЗП", "Хэштэг:", 100, 100) + 1).Value
    ' hashTag - Хэштэг:
    hashTag = hashTagFromSheet("DB_ЗП")

    ' Файл-вложение (!!!)
    attachmentFile = ThisWorkbook.Sheets("DB_ЗП").Range("R3").Value
    
    ' Текст письма
    текстПисьма = "" + Chr(13)
    текстПисьма = текстПисьма + "" + ThisWorkbook.Sheets("DB_ЗП").Cells(rowByValue(ThisWorkbook.Name, "DB_ЗП", "Список получателей:", 100, 100), ColumnByValue(ThisWorkbook.Name, "DB_ЗП", "Список получателей:", 100, 100) + 2).Value + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "" + getFromAddrBook("РД", 2) + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "Уважаемые руководители," + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "Отчет по привлечению ЗП 18+ офисами и ОКП (файл во вложении)." + Chr(13)
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
     
    ' Строка статуса
    Application.StatusBar = ""
     
  ' End If
  
End Sub

' Обработка Pipe ЗП
Sub Обработка_Pipe_ЗП()

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
    ThisWorkbook.Sheets("DB_ЗП").Activate

    Sheets_Name = "{D24520B3-6725-EA11-B826-02BFAC"

    ' Проверка формы отчета
    CheckFormatReportResult = CheckFormatReport(ReportName_String, Sheets_Name, 21, Date)
    If CheckFormatReportResult = "OK" Then
      
      column_Регион = ColumnByValue(ReportName_String, Sheets_Name, "Регион", 100, 300)
      column_Офис_обслуживания = ColumnByValue(ReportName_String, Sheets_Name, "Офис обслуживания", 100, 300)
      column_Планируемая_выдача_в_этом_квартале = ColumnByValue(ReportName_String, Sheets_Name, "Планируемая выдача в этом квартале", 100, 300)
      column_Общий_потенциал_выдачи = ColumnByValue(ReportName_String, Sheets_Name, "Общий потенциал выдачи", 100, 300)
      '
      column_Потенциал = ColumnByValue(ThisWorkbook.Name, "DB_ЗП", "Потенциал", 100, 300)
      column_Выдачи_Q = ColumnByValue(ThisWorkbook.Name, "DB_ЗП", "Выдачи Q", 100, 300)
      column_Карты_18_шт = ColumnByValue(ThisWorkbook.Name, "DB_ЗП", "Карты 18+, шт.", 100, 300)
      column_процент_плана = ColumnByValue(ThisWorkbook.Name, "DB_ЗП", "% плана", 100, 300)
      
      ' Обнуление переменных по РОО
      Итого_РОО_Планируемая_выдача_в_этом_квартале = 0
      Итого_РОО_Общий_потенциал_выдачи = 0
      
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

        ' Обнуление переменных
        Итого_Планируемая_выдача_в_этом_квартале = 0
        Итого_Общий_потенциал_выдачи = 0

        rowCount = 2
        Do While Workbooks(ReportName_String).Sheets(Sheets_Name).Cells(rowCount, column_Регион).Value = "Тюменский ОО1"
        
          ' Если это текущий офис
          If InStr(Workbooks(ReportName_String).Sheets(Sheets_Name).Cells(rowCount, column_Офис_обслуживания).Value, officeNameInReport) <> 0 Then
            
            Итого_Планируемая_выдача_в_этом_квартале = Итого_Планируемая_выдача_в_этом_квартале + Workbooks(ReportName_String).Sheets(Sheets_Name).Cells(rowCount, column_Планируемая_выдача_в_этом_квартале).Value
            If column_Общий_потенциал_выдачи <> 0 Then
              Итого_Общий_потенциал_выдачи = Итого_Общий_потенциал_выдачи + Workbooks(ReportName_String).Sheets(Sheets_Name).Cells(rowCount, column_Общий_потенциал_выдачи).Value
            End If
            '
            Итого_РОО_Планируемая_выдача_в_этом_квартале = Итого_РОО_Планируемая_выдача_в_этом_квартале + Workbooks(ReportName_String).Sheets(Sheets_Name).Cells(rowCount, column_Планируемая_выдача_в_этом_квартале).Value
            
            If column_Общий_потенциал_выдачи <> 0 Then
              Итого_РОО_Общий_потенциал_выдачи = Итого_РОО_Общий_потенциал_выдачи + Workbooks(ReportName_String).Sheets(Sheets_Name).Cells(rowCount, column_Общий_потенциал_выдачи).Value
            End If
            
          End If
        
        
          ' Следующая запись
          rowCount = rowCount + 1
          Application.StatusBar = officeNameInReport + ": " + CStr(rowCount) + "..."
          DoEventsInterval (rowCount)
        Loop
   
        ' Выводим данные по офису
        row_officeNameInReport = rowByValue(ThisWorkbook.Name, "DB_ЗП", getNameOfficeByNumber(i), 100, 100)
        ' Потенциал
        ThisWorkbook.Sheets("DB_ЗП").Cells(row_officeNameInReport, column_Потенциал).Value = Итого_Общий_потенциал_выдачи
        ThisWorkbook.Sheets("DB_ЗП").Cells(row_officeNameInReport, column_Потенциал).NumberFormat = "#,##0"
        ThisWorkbook.Sheets("DB_ЗП").Cells(row_officeNameInReport, column_Потенциал).HorizontalAlignment = xlRight
        ' Выдача в квартале
        ThisWorkbook.Sheets("DB_ЗП").Cells(row_officeNameInReport, column_Выдачи_Q).Value = Итого_Планируемая_выдача_в_этом_квартале
        ThisWorkbook.Sheets("DB_ЗП").Cells(row_officeNameInReport, column_Выдачи_Q).NumberFormat = "#,##0"
        ThisWorkbook.Sheets("DB_ЗП").Cells(row_officeNameInReport, column_Выдачи_Q).HorizontalAlignment = xlRight
        ' Процент от плана
        ThisWorkbook.Sheets("DB_ЗП").Cells(row_officeNameInReport, column_процент_плана).Value = РассчетДоли(ThisWorkbook.Sheets("DB_ЗП").Cells(row_officeNameInReport, column_Карты_18_шт).Value, Итого_Планируемая_выдача_в_этом_квартале, 3)
        ThisWorkbook.Sheets("DB_ЗП").Cells(row_officeNameInReport, column_процент_плана).NumberFormat = "0%"
        ThisWorkbook.Sheets("DB_ЗП").Cells(row_officeNameInReport, column_процент_плана).HorizontalAlignment = xlRight
        ' Окраска ячейки СФЕТОФОР ' 70. Заливка ячейки цветом "светофор" - значение In_Value в %, In_Target в %
        Call Full_Color_RangeII("DB_ЗП", row_officeNameInReport, column_процент_плана, ThisWorkbook.Sheets("DB_ЗП").Cells(row_officeNameInReport, column_процент_плана).Value, 1)

      Next i ' Следующий офис
      
      ' Выводим итоги обработки
      ' Итого Потенциал
      ThisWorkbook.Sheets("DB_ЗП").Cells(row_officeNameInReport + 1, column_Потенциал).Value = Итого_РОО_Общий_потенциал_выдачи
      ThisWorkbook.Sheets("DB_ЗП").Cells(row_officeNameInReport + 1, column_Потенциал).NumberFormat = "#,##0"
      ThisWorkbook.Sheets("DB_ЗП").Cells(row_officeNameInReport + 1, column_Потенциал).HorizontalAlignment = xlRight
      Call Полужирный_текст(ThisWorkbook.Name, "DB_ЗП", row_officeNameInReport + 1, column_Потенциал)
      
      ' Итого Выдачи Q
      ThisWorkbook.Sheets("DB_ЗП").Cells(row_officeNameInReport + 1, column_Выдачи_Q).Value = Итого_РОО_Планируемая_выдача_в_этом_квартале
      ThisWorkbook.Sheets("DB_ЗП").Cells(row_officeNameInReport + 1, column_Выдачи_Q).NumberFormat = "#,##0"
      ThisWorkbook.Sheets("DB_ЗП").Cells(row_officeNameInReport + 1, column_Выдачи_Q).HorizontalAlignment = xlRight
      Call Полужирный_текст(ThisWorkbook.Name, "DB_ЗП", row_officeNameInReport + 1, column_Выдачи_Q)
      
      ' Итого % плана
      ThisWorkbook.Sheets("DB_ЗП").Cells(row_officeNameInReport + 1, column_процент_плана).Value = РассчетДоли(ThisWorkbook.Sheets("DB_ЗП").Cells(row_officeNameInReport + 1, column_Карты_18_шт).Value, Итого_РОО_Планируемая_выдача_в_этом_квартале, 3)
      ThisWorkbook.Sheets("DB_ЗП").Cells(row_officeNameInReport + 1, column_процент_плана).NumberFormat = "0%"
      ThisWorkbook.Sheets("DB_ЗП").Cells(row_officeNameInReport + 1, column_процент_плана).HorizontalAlignment = xlRight
      Call Полужирный_текст(ThisWorkbook.Name, "DB_ЗП", row_officeNameInReport + 1, column_процент_плана)
      
      ' Сохранение изменений
      ThisWorkbook.Save
    
      ' Копирование в таблицу для отправки
      Call copy_DB_ЗП_ToSend
    
      ' Отправка шаблона письма
      Call Отправка_Lotus_Notes_DB_ЗП
    
      ' Переменная завершения обработки
      finishProcess = True
    Else
      ' Сообщение о неверном формате отчета или даты
      MsgBox ("Проверьте отчет: " + CheckFormatReportResult + "!")
    End If ' Проверка формы отчета

    ' Закрываем файл с отчетом без сохранения изменений (параметр SaveChanges:=False)
    Workbooks(Dir(FileName)).Close SaveChanges:=False
    
    ' Переходим в ячейку M2
    ThisWorkbook.Sheets("DB_ЗП").Range("A1").Select

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


' Создание книги потенциалом карт НВ, ДВ, зачисление
Sub createBook_out_DB_ЗП(In_OutBookName)

    Workbooks.Add
    ActiveWorkbook.SaveAs FileName:=In_OutBookName
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Activate
    
    ' TB_CONTR
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 1).Value = "TB_CONTR"
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns("A:A").EntireColumn.ColumnWidth = 17
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 1).HorizontalAlignment = xlCenter
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 1).NumberFormat = "@"
    
    ' SALORGCD
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 2).Value = "SALORGCD"
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns("B:B").EntireColumn.ColumnWidth = 16
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 2).HorizontalAlignment = xlCenter
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 2).NumberFormat = "@"
    
    ' NAMEORG
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 3).Value = "NAMEORG"
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns("C:C").EntireColumn.ColumnWidth = 42
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 3).HorizontalAlignment = xlCenter
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 3).NumberFormat = "@"
    
    ' INNSORG
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 4).Value = "INNSORG"
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns("D:D").EntireColumn.ColumnWidth = 17
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 4).HorizontalAlignment = xlCenter
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 4).NumberFormat = "@"
    
    ' ZCONTRN
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 5).Value = "ZCONTRN"
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns("E:E").EntireColumn.ColumnWidth = 30
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 5).HorizontalAlignment = xlCenter
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 5).NumberFormat = "@"
    
    ' CNTR_YM_FROM
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 6).Value = "CNTR_YM_FROM"
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns("F:F").EntireColumn.ColumnWidth = 15
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 6).HorizontalAlignment = xlCenter
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 6).NumberFormat = "@"
    
    ' Потенциал (НВ) / Потенциал (ДВ) / Выдана без зачислений
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 7).Value = "Status"
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns("G:G").EntireColumn.ColumnWidth = 22
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 7).HorizontalAlignment = xlCenter
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 7).NumberFormat = "@"
    
    ' Дата выдачи
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 8).Value = "Дата выдачи"
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns("H:H").EntireColumn.ColumnWidth = 15
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 8).HorizontalAlignment = xlCenter
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 8).NumberFormat = "@"
    
    ' ДО
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 9).Value = "Офис"
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns("I:I").EntireColumn.ColumnWidth = 25
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 9).HorizontalAlignment = xlCenter
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 9).NumberFormat = "@"
    
    ' ActiveCell.Offset(0, -4).Columns("A:A").EntireColumn.Select
    ' Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Range("C:C").Select
    ' Числовой
    ' Selection.NumberFormat = "0"
    ' Текстовый
    ' Selection.NumberFormat = "@"

    ' Установка фильтров
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Range("A1:I1").Select
    Selection.AutoFilter
    
End Sub

