Attribute VB_Name = "Module_Лист9"
' Обработать отчет "Воронка"
Sub Обработать_Воронку()
Attribute Обработать_Воронку.VB_ProcData.VB_Invoke_Func = " \n14"
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
    ThisWorkbook.Sheets("Лист9").Activate

    ' Проверка формы отчета
    CheckFormatReportResult = CheckFormatReport(ReportName_String, "ПК", 15, Date)
    
    If CheckFormatReportResult = "OK" Then
      
      ' Открываем BASE\Credits
      OpenBookInBase ("Credits")
      
      ' Переходим на окно DB
      ThisWorkbook.Sheets("Лист9").Activate
      
      ' Файл с вложением в Q3
      ThisWorkbook.Sheets("Лист9").Range("Q3").Value = FileName
             
      ' Обновляем список получателей
      ThisWorkbook.Sheets("Лист9").Cells(rowByValue(ThisWorkbook.Name, "Лист8", "Список получателей:", 100, 100), ColumnByValue(ThisWorkbook.Name, "Лист8", "Список получателей:", 100, 100) + 2).Value = _
          getFromAddrBook("УДО2,УДО3,УДО4,УДО5,НОРПиКО1,НОРПиКО2,НОРПиКО3,НОРПиКО4,НОРПиКО5,НОКП,РРКК", 2)

      ' Очищаем таблицы на Листе9
      
      Application.StatusBar = "Определение столбцов..."

      ' Наименование офиса
      Column_Заголовок_столбца_офисы = ColumnByValue(ReportName_String, "ПК", "Наименование дополнительного офи", 1000, 1000)
      ' Заявка
      Column_CLAIM_ID = ColumnByValue(ReportName_String, "ПК", "CLAIM_ID", 1000, 1000) ' поле Заявка_ID
      Column_Дата_созд_заявки = ColumnByValue(ReportName_String, "ПК", "Дата созд заявки", 1000, 1000) ' поле Заявка_дата
      Column_Запрашив_сумма = ColumnByValue(ReportName_String, "ПК", "Запрашив сумма", 1000, 1000) ' поле Заявка_сумма
      Column_Создатель_заявки = ColumnByValue(ReportName_String, "ПК", "Создатель заявки", 1000, 1000) ' Заявка_создал
      Column_Источник_поступ = ColumnByValue(ReportName_String, "ПК", "Источник поступ", 1000, 1000) ' Заявка_источник
      ' Решение
      Column_Дата_решения = ColumnByValue(ReportName_String, "ПК", "Дата решения", 1000, 1000) ' поле Решение_дата
      Column_Статус = ColumnByValue(ReportName_String, "ПК", "Статус", 1000, 1000) ' поле Заявка_статус
      Column_Статус2 = ColumnByValue(ReportName_String, "ПК", "STATUS_CLAIM_DEMOGR", 1000, 1000) ' Заявка_статус2
      Column_Статус3 = ColumnByValue(ReportName_String, "ПК", "Детализац статуса", 1000, 1000) ' Заявка_статус3
      Column_Статус4 = ColumnByValue(ReportName_String, "ПК", "STATUS_DRR_VAR", 1000, 1000) ' STATUS_DRR_VAR -> Заявка_статус4
      Column_Статус5 = ColumnByValue(ReportName_String, "ПК", "DECL_REASON_DRR_VAR", 1000, 1000) ' DECL_REASON_DRR_VAR -> Заявка_статус5
      Column_Мин_сумма_одобрения = ColumnByValue(ReportName_String, "ПК", "Мин_сумма одобрения", 1000, 1000) ' Решение_сумма_min
      Column_Средняя_сумма_одобрения = ColumnByValue(ReportName_String, "ПК", "Средняя_сумма одобрения", 1000, 1000) ' Решение_сумма_mid
      Column_Мах_сумма_одобрения = ColumnByValue(ReportName_String, "ПК", "Мах_сумма одобрения", 1000, 1000) ' Решение_сумма_max
      Column_МинСтавка_Решения = ColumnByValue(ReportName_String, "ПК", "МинСтавка Решения", 1000, 1000) ' Решение_МинСтавка
      Column_Скор_бал = ColumnByValue(ReportName_String, "ПК", "Скор  бал", 1000, 1000) ' Решение_СкорБал
      Column_Программа_кредитования_итог = ColumnByValue(ReportName_String, "ПК", "Программа кредитования итог", 1000, 1000)
      ' Клиент
      Column_НК = ColumnByValue(ReportName_String, "ПК", "Идентиф клиента", 1000, 1000) ' Клиент_НК
      Column_Форма_справки_о_доходах = ColumnByValue(ReportName_String, "ПК", "Форма справки о доходах", 1000, 1000) ' Форма справки о доходах -> Справка_доход
      Column_BIRTH_DATE = ColumnByValue(ReportName_String, "ПК", "BIRTH_DATE", 1000, 1000) ' Клиент_ДР
      Column_FIO = ColumnByValue(ReportName_String, "ПК", "FIO", 1000, 1000) ' Клиент_FIO
      ' Место работы
      Column_Название_комп = ColumnByValue(ReportName_String, "ПК", "Название комп", 1000, 1000) ' Компания_наим
      Column_ИНН_компании = ColumnByValue(ReportName_String, "ПК", "ИНН компании", 1000, 1000) ' Компания_ИНН
      Column_Сегмент_клиента = ColumnByValue(ReportName_String, "ПК", "Сегмент клиента", 1000, 1000) ' Сегмент
      Column_Госслужащие = ColumnByValue(ReportName_String, "ПК", "Госслужащие - да", 1000, 1000)
      Column_Аккредитованная_компания = ColumnByValue(ReportName_String, "ПК", "Признак аккредит", 1000, 1000)
      Column_Зеленая_компания = ColumnByValue(ReportName_String, "ПК", Chr(34) + "Зеленая" + Chr(34) + " компания", 1000, 1000)
      ' Сплит к ПК
      Column_Предложен_Split = ColumnByValue(ReportName_String, "ПК", "Предложен split", 1000, 1000) ' Split
      Column_Взят_Split = ColumnByValue(ReportName_String, "ПК", "Взят split-sell", 1000, 1000)
      ' Выдача
      column_Дата_выдачи = ColumnByValue(ReportName_String, "ПК", "Дата выдачи", 1000, 1000) ' Выдача_дата
      Column_Номер_кредит_дог = ColumnByValue(ReportName_String, "ПК", "Номер кредит дог", 1000, 1000) ' Номер_КД
      Column_выдачи_руб = ColumnByValue(ReportName_String, "ПК", "Выдачи руб", 1000, 1000) ' Выдача_сумма
      Column_Ставка_выдачи = ColumnByValue(ReportName_String, "ПК", "Ставка выдачи", 1000, 1000) ' Выдача_ставка
      Column_Срок_кредита = ColumnByValue(ReportName_String, "ПК", "Срок кредита (мес) (2)", 1000, 1000) ' Выдача_срок
      Column_FIO_MPP_OTHER = ColumnByValue(ReportName_String, "ПК", "FIO_MPP_OTHER", 1000, 1000) ' Выдача_сотрудник
      Column_Flag_Chanel_IB = ColumnByValue(ReportName_String, "ПК", "Flag_Chanel_IB", 1000, 1000) ' Канал ИБ
      Column_Источник_заявки = ColumnByValue(ReportName_String, "ПК", "Источник поступ", 1000, 1000) ' Источник поступ
       
      Application.StatusBar = "Определение периода в Воронке..."
       
      ' Определяем через 1 проход максимальную и минимальную дату заявок в файле по всем офисам и регионам
      ' 1. Определение Максимальной дата заявки в Воронке
      ' Дата начала
      dateBegin = CDate("01.01.2022")
      ' Дата окончания
      dateEnd = CDate("01.01.2000")
      '
      rowCount = 2
      Do While Not IsEmpty(Workbooks(ReportName_String).Sheets("ПК").Cells(rowCount, 1).Value)
        
        ' Если это один из офисов
        If (InStr(Workbooks(ReportName_String).Sheets("ПК").Cells(rowCount, Column_Заголовок_столбца_офисы).Value, "Тюменский") <> 0) Or (InStr(Workbooks(ReportName_String).Sheets("ПК").Cells(rowCount, Column_Заголовок_столбца_офисы).Value, "Сургутский") <> 0) Or (InStr(Workbooks(ReportName_String).Sheets("ПК").Cells(rowCount, Column_Заголовок_столбца_офисы).Value, "Нижневартовский") <> 0) Or (InStr(Workbooks(ReportName_String).Sheets("ПК").Cells(rowCount, Column_Заголовок_столбца_офисы).Value, "Новоуренгойский") <> 0) Or (InStr(Workbooks(ReportName_String).Sheets("ПК").Cells(rowCount, Column_Заголовок_столбца_офисы).Value, "Тарко-Сале") <> 0) Then
        
          ' Если дата больше dateEnd
          If Workbooks(ReportName_String).Sheets("ПК").Cells(rowCount, Column_Дата_созд_заявки).Value > dateEnd Then
            dateEnd = CDate(Workbooks(ReportName_String).Sheets("ПК").Cells(rowCount, Column_Дата_созд_заявки).Value)
          End If
        
          ' Если дата меньше dateBegin
          If Workbooks(ReportName_String).Sheets("ПК").Cells(rowCount, Column_Дата_созд_заявки).Value < dateBegin Then
            dateBegin = CDate(Workbooks(ReportName_String).Sheets("ПК").Cells(rowCount, Column_Дата_созд_заявки).Value)
          End If
        
        End If
        
        ' Следующая запись
        rowCount = rowCount + 1
        Application.StatusBar = "Определение периода всей Воронки: " + CStr(rowCount) + "..."
        DoEventsInterval (rowCount)
      Loop

      ' Записываем значения дат в Тему
      ThisWorkbook.Sheets("Лист9").Range("P2").Value = "Воронка по ПК и КК " + CStr(dateBegin) + "-" + CStr(dateEnd)
       
      ' Обрабатываем отчет
      ' Цикл по 5-ти офисам
      ' Обработка отчета
      For i = 1 To 5
        
        ' Переходим на окно DB
        ThisWorkbook.Sheets("Лист9").Activate
        
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


        rowCount = 2
        Do While Not IsEmpty(Workbooks(ReportName_String).Sheets("ПК").Cells(rowCount, 1).Value)
        
          ' Если это текущий офис
          If InStr(Workbooks(ReportName_String).Sheets("ПК").Cells(rowCount, Column_Заголовок_столбца_офисы).Value, officeNameInReport) <> 0 Then
            
            ' Переменные (для сокращения строки передачи параметров в процедуру)
            Заявка_ID_Var = Workbooks(ReportName_String).Sheets("ПК").Cells(rowCount, Column_CLAIM_ID).Value
            Заявка_дата_Var = Workbooks(ReportName_String).Sheets("ПК").Cells(rowCount, Column_Дата_созд_заявки).Value
            Заявка_сумма_Var = Workbooks(ReportName_String).Sheets("ПК").Cells(rowCount, Column_Запрашив_сумма).Value
            Решение_дата_Var = Workbooks(ReportName_String).Sheets("ПК").Cells(rowCount, Column_Дата_решения).Value
            Заявка_статус_Var = Workbooks(ReportName_String).Sheets("ПК").Cells(rowCount, Column_Статус).Value
            Заявка_статус2_Var = Workbooks(ReportName_String).Sheets("ПК").Cells(rowCount, Column_Статус2).Value
            Заявка_статус3_Var = Workbooks(ReportName_String).Sheets("ПК").Cells(rowCount, Column_Статус3).Value
            Решение_сумма_min_Var = Workbooks(ReportName_String).Sheets("ПК").Cells(rowCount, Column_Мин_сумма_одобрения).Value
            Решение_сумма_mid_Var = Workbooks(ReportName_String).Sheets("ПК").Cells(rowCount, Column_Средняя_сумма_одобрения).Value
            Решение_сумма_max_Var = Workbooks(ReportName_String).Sheets("ПК").Cells(rowCount, Column_Мах_сумма_одобрения).Value
            Решение_МинСтавка_Var = Workbooks(ReportName_String).Sheets("ПК").Cells(rowCount, Column_МинСтавка_Решения).Value
            Решение_СкорБал_Var = Workbooks(ReportName_String).Sheets("ПК").Cells(rowCount, Column_Скор_бал).Value
            Клиент_НК_Var = Workbooks(ReportName_String).Sheets("ПК").Cells(rowCount, Column_НК).Value
            Компания_ИНН_Var = Workbooks(ReportName_String).Sheets("ПК").Cells(rowCount, Column_ИНН_компании).Value
            Компания_наим_Var = Workbooks(ReportName_String).Sheets("ПК").Cells(rowCount, Column_Название_комп).Value
            Сегмент_Var = Workbooks(ReportName_String).Sheets("ПК").Cells(rowCount, Column_Сегмент_клиента).Value
            Выдача_дата_Var = Workbooks(ReportName_String).Sheets("ПК").Cells(rowCount, column_Дата_выдачи).Value
            Выдача_сумма_Var = Workbooks(ReportName_String).Sheets("ПК").Cells(rowCount, Column_выдачи_руб).Value
            Канал_ИБ_Var = Workbooks(ReportName_String).Sheets("ПК").Cells(rowCount, Column_Flag_Chanel_IB).Value
            Воронка_файл_Var = Dir(ReportName_String)
          
            ' Добавляем 4 поля
            Заявка_статус4_Var = Workbooks(ReportName_String).Sheets("ПК").Cells(rowCount, Column_Статус4).Value
            Заявка_статус5_Var = Workbooks(ReportName_String).Sheets("ПК").Cells(rowCount, Column_Статус5).Value
            Программа_кредитования_Var = Workbooks(ReportName_String).Sheets("ПК").Cells(rowCount, Column_Программа_кредитования_итог).Value
            Справка_доход_Var = Workbooks(ReportName_String).Sheets("ПК").Cells(rowCount, Column_Форма_справки_о_доходах).Value
            Источник_заявки_Var = Workbooks(ReportName_String).Sheets("ПК").Cells(rowCount, Column_Источник_заявки).Value
            Создатель_заявки_Var = Workbooks(ReportName_String).Sheets("ПК").Cells(rowCount, Column_Создатель_заявки).Value

            ' Заносим в BASE\Credits
            Call InsertRecordInBook2("Credits", "Лист1", "Заявка_ID", Заявка_ID_Var, _
                                            "Заявка_ID", Заявка_ID_Var, _
                                              "Офис", officeNameInReport, _
                                                "Заявка_дата", Заявка_дата_Var, _
                                                  "Заявка_сумма", Заявка_сумма_Var, _
                                                    "Решение_дата", Решение_дата_Var, _
                                                      "Заявка_статус", Заявка_статус_Var, _
                                                       "Заявка_статус2", Заявка_статус2_Var, _
                                                         "Заявка_статус3", Заявка_статус3_Var, _
                                                            "Решение_сумма_min", Решение_сумма_min_Var, _
                                                              "Решение_сумма_mid", Решение_сумма_mid_Var, _
                                                                "Решение_сумма_max", Решение_сумма_max_Var, _
                                                                  "Решение_МинСтавка", Решение_МинСтавка_Var, _
                                                                    "Решение_СкорБал", Решение_СкорБал_Var, _
                                                                      "Клиент_НК", Клиент_НК_Var, _
                                                                        "Компания_ИНН", Компания_ИНН_Var, _
                                                                          "Компания_наим", Компания_наим_Var, _
                                                                            "Сегмент", Сегмент_Var, _
                                                                              "Выдача_дата", Выдача_дата_Var, _
                                                                                "Выдача_сумма", Выдача_сумма_Var, _
                                                                                  "Канал_ИБ", Канал_ИБ_Var, _
                                                                                    "Заявка_статус4", Заявка_статус4_Var, _
                                                                                      "Заявка_статус5", Заявка_статус5_Var, _
                                                                                        "Программа_кредитования", Программа_кредитования_Var, "Справка_доход", Справка_доход_Var, "Воронка_файл", Воронка_файл_Var, "Заявка_источник", Источник_заявки_Var, "Заявка_создал", Создатель_заявки_Var, "", "")
                                                                                    
 
            
            ' Открываем BASE\FUNNEL файл с ИНН и заносим данные по заявке (если нет, то создаем файл)
            Call InsertRecordInBASEFUNNELBook("Credits", "Лист1", "Заявка_ID", Workbooks(ReportName_String).Sheets("ПК").Cells(rowCount, Column_CLAIM_ID).Value, _
                                                "Заявка_ID", Workbooks(ReportName_String).Sheets("ПК").Cells(rowCount, Column_CLAIM_ID).Value, _
                                                  "Офис", officeNameInReport, _
                                                    "Заявка_дата", Workbooks(ReportName_String).Sheets("ПК").Cells(rowCount, Column_Дата_созд_заявки).Value, _
                                                      "Заявка_сумма", Workbooks(ReportName_String).Sheets("ПК").Cells(rowCount, Column_Запрашив_сумма).Value, _
                                                        "Решение_дата", Workbooks(ReportName_String).Sheets("ПК").Cells(rowCount, Column_Дата_решения).Value, _
                                                          "Заявка_статус", Workbooks(ReportName_String).Sheets("ПК").Cells(rowCount, Column_Статус).Value, _
                                                            "Заявка_статус2", Workbooks(ReportName_String).Sheets("ПК").Cells(rowCount, Column_Статус2).Value, _
                                                              "Заявка_статус3", Workbooks(ReportName_String).Sheets("ПК").Cells(rowCount, Column_Статус3).Value, _
                                                                "Решение_сумма_min", Workbooks(ReportName_String).Sheets("ПК").Cells(rowCount, Column_Мин_сумма_одобрения).Value, _
                                                                  "Решение_сумма_mid", Workbooks(ReportName_String).Sheets("ПК").Cells(rowCount, Column_Средняя_сумма_одобрения).Value, _
                                                                    "Решение_сумма_max", Workbooks(ReportName_String).Sheets("ПК").Cells(rowCount, Column_Мах_сумма_одобрения).Value, _
                                                                      "Решение_МинСтавка", Workbooks(ReportName_String).Sheets("ПК").Cells(rowCount, Column_МинСтавка_Решения).Value, _
                                                                        "Решение_СкорБал", Workbooks(ReportName_String).Sheets("ПК").Cells(rowCount, Column_Скор_бал).Value, _
                                                                          "Клиент_НК", Workbooks(ReportName_String).Sheets("ПК").Cells(rowCount, Column_НК).Value, _
                                                                            "Компания_ИНН", Workbooks(ReportName_String).Sheets("ПК").Cells(rowCount, Column_ИНН_компании).Value, _
                                                                              "Компания_наим", Workbooks(ReportName_String).Sheets("ПК").Cells(rowCount, Column_Название_комп).Value, _
                                                                                "Сегмент", Workbooks(ReportName_String).Sheets("ПК").Cells(rowCount, Column_Сегмент_клиента).Value, _
                                                                                  "Выдача_дата", Workbooks(ReportName_String).Sheets("ПК").Cells(rowCount, column_Дата_выдачи).Value, _
                                                                                    "Выдача_сумма", Workbooks(ReportName_String).Sheets("ПК").Cells(rowCount, Column_выдачи_руб).Value, _
                                                                                      "Канал_ИБ", Workbooks(ReportName_String).Sheets("ПК").Cells(rowCount, Column_Flag_Chanel_IB).Value, _
                                                                                        "Заявка_статус4", Заявка_статус4_Var, _
                                                                                          "Заявка_статус5", Заявка_статус5_Var, _
                                                                                            "Программа_кредитования", Программа_кредитования_Var, _
                                                                                             "Справка_доход", Справка_доход_Var, "Заявка_источник", Источник_заявки_Var, "Заявка_создал", Создатель_заявки_Var, "", "", "", "")
          End If
        
        
          ' Следующая запись
          rowCount = rowCount + 1
          Application.StatusBar = officeNameInReport + ": " + CStr(rowCount) + "..."
          DoEventsInterval (rowCount)
        Loop
   
        ' Выводим данные по офису
      
      Next i ' Следующий офис
      
      Application.StatusBar = "Обработка результатов..."
      
      ' Выводим итоги обработки на Лист8 из BASE\Credits
      Call Вывод_данных_по_заявкам_Лист9
      
      ' Сохранение изменений
      ThisWorkbook.Save
    
      ' Закрываем BASE\Credits
      CloseBook ("Credits")

      ' Закрываем файл с отчетом без сохранения изменений (параметр SaveChanges:=False)
      Workbooks(Dir(FileName)).Close SaveChanges:=False

      Application.StatusBar = ""

      ' Отправка письма с Воронкой в офисы
      Call Отправка_Lotus_Notes_Лист9
    
      ' Переменная завершения обработки
      finishProcess = True
    
    Else
      ' Сообщение о неверном формате отчета или даты
      MsgBox ("Проверьте отчет: " + CheckFormatReportResult + "!")
    End If ' Проверка формы отчета

    ' Закрываем BASE\Credits
    ' CloseBook ("Credits")

    ' Закрываем файл с отчетом без сохранения изменений (параметр SaveChanges:=False)
    ' Workbooks(Dir(FileName)).Close SaveChanges:=False
    
    ' Переходим в ячейку M2
    ThisWorkbook.Sheets("Лист9").Range("A1").Select

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

                                
' Открываем BASE\FUNNEL файл с ИНН и заносим данные по заявке (если нет, то создаем файл)
Sub InsertRecordInBASEFUNNELBook(In_BookName, In_Sheet, In_FieldKeyName, In_FieldKeyValue, In_FieldName1, In_FieldValue1, In_FieldName2, In_FieldValue2, In_FieldName3, In_FieldValue3, In_FieldName4, In_FieldValue4, In_FieldName5, In_FieldValue5, In_FieldName6, In_FieldValue6, In_FieldName7, In_FieldValue7, In_FieldName8, In_FieldValue8, In_FieldName9, In_FieldValue9, In_FieldName10, In_FieldValue10, In_FieldName11, In_FieldValue11, In_FieldName12, In_FieldValue12, In_FieldName13, In_FieldValue13, In_FieldName14, In_FieldValue14, In_FieldName15, In_FieldValue15, In_FieldName16, In_FieldValue16, In_FieldName17, In_FieldValue17, In_FieldName18, In_FieldValue18, In_FieldName19, In_FieldValue19, In_FieldName20, In_FieldValue20, In_FieldName21, In_FieldValue21, In_FieldName22, In_FieldValue22, In_FieldName23, In_FieldValue23, In_FieldName24, In_FieldValue24, In_FieldName25, In_FieldValue25, In_FieldName26, In_FieldValue26, In_FieldName27, In_FieldValue27, In_FieldName28, In_FieldValue28)
    
  ' ИНН это In_FieldValue15
  If In_FieldValue15 = "" Then
    In_FieldValue15 = "emptyInn"
  End If
  
  If Dir(ThisWorkbook.Path + "\BASE\FUNNEL\" + In_FieldValue15 + ".xlsx") = "" Then
    ' Если файла нет, то открываем шаблон и сохраняем в \BASE\FUNNEL
    Workbooks.Open (ThisWorkbook.Path + "\Templates\Воронка.xlsx")
    Workbooks("Воронка").SaveAs FileName:=ThisWorkbook.Path + "\BASE\FUNNEL\" + In_FieldValue15 + ".xlsx", createBackUp:=False
    ' Переходим на окно DB
    ThisWorkbook.Sheets("Лист9").Activate
  Else
    ' Если файл существует
    Workbooks.Open (ThisWorkbook.Path + "\BASE\FUNNEL\" + In_FieldValue15 + ".xlsx")
    ThisWorkbook.Sheets("Лист9").Activate
  End If
  
  ' Заносим данные в открытый файл
  Call InsertRecordInBook2(In_FieldValue15, In_Sheet, In_FieldKeyName, In_FieldKeyValue, In_FieldName1, In_FieldValue1, In_FieldName2, In_FieldValue2, In_FieldName3, In_FieldValue3, In_FieldName4, In_FieldValue4, In_FieldName5, In_FieldValue5, In_FieldName6, In_FieldValue6, In_FieldName7, In_FieldValue7, In_FieldName8, In_FieldValue8, In_FieldName9, In_FieldValue9, In_FieldName10, In_FieldValue10, In_FieldName11, In_FieldValue11, In_FieldName12, In_FieldValue12, In_FieldName13, In_FieldValue13, In_FieldName14, In_FieldValue14, In_FieldName15, In_FieldValue15, In_FieldName16, In_FieldValue16, In_FieldName17, In_FieldValue17, In_FieldName18, In_FieldValue18, In_FieldName19, In_FieldValue19, In_FieldName20, In_FieldValue20, In_FieldName21, In_FieldValue21, In_FieldName22, In_FieldValue22, In_FieldName23, In_FieldValue23, In_FieldName24, In_FieldValue24, In_FieldName25, In_FieldValue25, In_FieldName26, In_FieldValue26, In_FieldName27, In_FieldValue27, In_FieldName28, In_FieldValue28)
  
  ' Закрываем файл
  Workbooks(In_FieldValue15).Close SaveChanges:=True
  
End Sub
 
' Выводим одобрения и отказы на Лист8
Sub Вывод_данных_по_заявкам_Лист9()
  
  ' Очищаем значения в таблице на Лист9
  For i = 31 To 35
    For j = 3 To 16
      ThisWorkbook.Sheets("Лист9").Cells(i, j).Value = 0
    Next j
  Next i
  
  ' 1. Определение Максимальной дата заявки
  dateEnd = CDate("01.01.2000")
  
  ' Номера столбцов в BASE\Credits
  Column_Заявка_дата = ColumnByValue("Credits", "Лист1", "Заявка_дата", 1000, 1000)
  column_офис = ColumnByValue("Credits", "Лист1", "Офис", 1000, 1000)
  Column_Сегмент = ColumnByValue("Credits", "Лист1", "Сегмент", 1000, 1000)
  Column_Заявка_статус = ColumnByValue("Credits", "Лист1", "Заявка_статус", 1000, 1000)
  Column_Решение_сумма_min = ColumnByValue("Credits", "Лист1", "Решение_сумма_min", 1000, 1000)
  
  rowCount = 2
  Do While Not IsEmpty(Workbooks("Credits").Sheets("Лист1").Cells(rowCount, 1).Value)
        
    ' Если дата больше
    If Workbooks("Credits").Sheets("Лист1").Cells(rowCount, Column_Заявка_дата).Value > dateEnd Then
      dateEnd = Workbooks("Credits").Sheets("Лист1").Cells(rowCount, Column_Заявка_дата).Value
    End If
        
    ' Следующая запись
    rowCount = rowCount + 1
    Application.StatusBar = "Определение периода: " + CStr(rowCount) + "..."
    DoEventsInterval (rowCount)
  Loop

  ' Определяем дату начала
  ' Дата начала
  dateBegin = Date_begin_day_month(dateEnd)
  
  ' Вывод заголовка Таблицы на Лист9
  ThisWorkbook.Sheets("Лист9").Range("B28").Value = "Одобрено в сегментах с " + CStr(dateBegin) + " по " + CStr(dateEnd)
  
  ' 2. Выборка заявок по Сегментам за период
  rowCount = 2
  Do While Not IsEmpty(Workbooks("Credits").Sheets("Лист1").Cells(rowCount, 1).Value)
        
    ' Если даты в диапазоне
    If (Workbooks("Credits").Sheets("Лист1").Cells(rowCount, Column_Заявка_дата).Value >= dateBegin) And (Workbooks("Credits").Sheets("Лист1").Cells(rowCount, Column_Заявка_дата).Value <= dateEnd) Then
        
      ' Определяем строку и столбец на Листе9
      rowЛист9 = officeInЛист9(Workbooks("Credits").Sheets("Лист1").Cells(rowCount, column_офис).Value)
      columnЛист9 = segmentInЛист9(Workbooks("Credits").Sheets("Лист1").Cells(rowCount, Column_Сегмент).Value)
      
      ' Суммируем если Одобрено - в поле сумма (9) не пусто
      If Not IsEmpty(Workbooks("Credits").Sheets("Лист1").Cells(rowCount, Column_Решение_сумма_min).Value) Then
        ThisWorkbook.Sheets("Лист9").Cells(rowЛист9, columnЛист9).Value = ThisWorkbook.Sheets("Лист9").Cells(rowЛист9, columnЛист9).Value + 1
      End If
      
      ' Суммируем если Отказ
      If InStr(Workbooks("Credits").Sheets("Лист1").Cells(rowCount, Column_Заявка_статус).Value, "Отказ") <> 0 Then
        ThisWorkbook.Sheets("Лист9").Cells(rowЛист9, columnЛист9 + 1).Value = ThisWorkbook.Sheets("Лист9").Cells(rowЛист9, columnЛист9 + 1).Value + 1
      End If
      
    End If
        
    ' Следующая запись
    rowCount = rowCount + 1
    Application.StatusBar = "Выборка по Сегментам: " + CStr(rowCount) + "..."
    DoEventsInterval (rowCount)
  Loop
  
  Application.StatusBar = ""

End Sub

' rowЛист9 = officeInЛист9(Workbooks("Credits").Sheets("Лист1").Cells(rowCount, 2).Value)
Function officeInЛист9(In_Office) As Integer
        
        ' Номера офисов от 1 до 5
        Select Case In_Office
          Case "Тюменский" ' ОО «Тюменский»
            officeInЛист9 = 31
          Case "Сургутский" ' ОО «Сургутский»
            officeInЛист9 = 32
          Case "Нижневартовский" ' ОО «Нижневартовский»
            officeInЛист9 = 33
          Case "Новоуренгойский" ' ОО «Новоуренгойский»
            officeInЛист9 = 34
          Case "Тарко-Сале" ' ОО «Тарко-Сале»
            officeInЛист9 = 35
        End Select
  
End Function

' columnЛист9 = segmentInЛист9(Workbooks("Credits").Sheets("Лист1").Cells(rowCount, 14).Value)
Function segmentInЛист9(In_Segment) As Integer
  
  сегмент_определен = False
  
  ' ОПК
  If (сегмент_определен = False) And (InStr(In_Segment, "ОПК") <> 0) Then
    segmentInЛист9 = 3
    сегмент_определен = True
  End If
    
  ' Спецкомпании
  If (сегмент_определен = False) And (InStr(In_Segment, "СПЕЦ.") <> 0) Then
    segmentInЛист9 = 5
    сегмент_определен = True
  End If
    
  ' Госслужащие
  If (сегмент_определен = False) And (InStr(In_Segment, "ГОССЛУЖАЩИЕ") <> 0) Then
    segmentInЛист9 = 7
    сегмент_определен = True
  End If
    
  ' Зеленые
  If (сегмент_определен = False) And (InStr(In_Segment, "ЗЕЛЁНЫХ") <> 0) Then
    segmentInЛист9 = 9
    сегмент_определен = True
  End If
    
  ' Зарплатный
  If (сегмент_определен = False) And (InStr(In_Segment, "ЗАРПЛАТНЫЙ") <> 0) Then
    segmentInЛист9 = 11
    сегмент_определен = True
  End If
    
  ' Открытый  рынок
  If (сегмент_определен = False) And (InStr(In_Segment, "ОТКРЫТЫЙ РЫНОК") <> 0) Then
    segmentInЛист9 = 13
    сегмент_определен = True
  End If
    
  ' Прочие
  If сегмент_определен = False Then
    segmentInЛист9 = 15
  End If
  
  
End Function


' Отправка письма: отправляю шаблон самому себе для последующей отправки в сеть письма на его основе:
Sub Отправка_Lotus_Notes_Лист9()
Dim темаПисьма, текстПисьма, hashTag, attachmentFile As String
Dim i As Byte
  
  ' Запрос
  If MsgBox("Отправить себе Шаблон письма с итогами обработки Воронки?", vbYesNo) = vbYes Then
    
    ' Тема письма - Тема:
    ' темаПисьма = ThisWorkbook.Sheets("Лист9").Cells(RowByValue(ThisWorkbook.Name, "Лист9", "Тема:", 100, 100), ColumnByValue(ThisWorkbook.Name, "Лист9", "Тема:", 100, 100) + 1).Value
    темаПисьма = subjectFromSheet("Лист9")

    ' hashTag - Хэштэг:
    ' hashTag = ThisWorkbook.Sheets("Лист9").Cells(RowByValue(ThisWorkbook.Name, "Лист9", "Хэштэг:", 100, 100), ColumnByValue(ThisWorkbook.Name, "Лист9", "Хэштэг:", 100, 100) + 1).Value
    ' hashTag - Хэштэг:
    hashTag = hashTagFromSheet("Лист9")

    ' Файл-вложение (!!!)
    attachmentFile = ThisWorkbook.Sheets("Лист9").Range("Q3").Value
    
    ' Текст письма
    текстПисьма = "" + Chr(13)
    текстПисьма = текстПисьма + "" + ThisWorkbook.Sheets("Лист9").Cells(rowByValue(ThisWorkbook.Name, "Лист9", "Список получателей:", 100, 100), ColumnByValue(ThisWorkbook.Name, "Лист9", "Список получателей:", 100, 100) + 2).Value + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "" + getFromAddrBook("РД", 2) + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "Уважаемые руководители," + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "1. Воронка по потреб кредитам и кредитным картам" + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "2. Сегмент Бюджетники:" + Chr(13)
    текстПисьма = текстПисьма + "- принято заявок за <месяц>: ___ шт."
    текстПисьма = текстПисьма + "- одобрено: ___ шт. (__%)"
    текстПисьма = текстПисьма + "- в т.ч. решений со ставками:" + Chr(13)
    текстПисьма = текстПисьма + "___% - __ шт. (__%)" + Chr(13)
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




