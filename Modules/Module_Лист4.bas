Attribute VB_Name = "Module_Лист4"
' Лист 4 Выход вкладов на Дату

' Выход вкладов по дням
Sub Выход_вкладов()
  
Dim Дата_следующего_воскресенья, Дата_первого_понедельника_за_след_воскресеньем, Дата_конца_месяца As Date
Dim Выход_вкладов_неделя_шт, Выход_вкладов_через_неделю_и_до_конца_месяца_шт As Integer
Dim Выход_вкладов_неделя_руб, Выход_вкладов_через_неделю_и_до_конца_месяца_руб As Double
Dim rowCount, RecCount_In_OutBookName As Integer
Dim dateBeginWeek, dateEndWeek As Date
' Индиктор вызова процедур setК_порInЕСУП, currentК_порInЕСУП
Dim call_setК_порInЕСУП, call_currentК_порInЕСУП As Boolean
       
  ' Выход депозитов
  ' Открыть файл с отчетом
  FileName = Application.GetOpenFilename("Excel Files (*.xls), *.xls", , "Открытие файла с отчетом")

  ' Если файл был выбран
  If (Len(FileName) > 5) Then
  
    ' Строка статуса
    Application.StatusBar = "Обработка отчета"

    ' Выводим для инфо данные об имени файла
    DBstrName_String = Dir(FileName)
  
    ' Файл-отчет: в "S3"
    ThisWorkbook.Sheets("Лист4").Range("S3").Value = FileName
  
    ' Открываем выбранную книгу (UpdateLinks:=0)
    Workbooks.Open FileName, 0
        
    ' Открываем BASE\Tasks
    OpenBookInBase ("Tasks")

    ' Переходим текущую форму
    ThisWorkbook.Sheets("Лист4").Activate
               
    ' Обработка отчета
    ' Выход вкладов на неделе с 10.02 по 16.02 (с сегодняшнего дня и до следующего воскресенья)
    
    ' Дата начала недели
    dateBeginWeek = weekStartDate(Date)
    
    ' Дата конца недели
    dateEndWeek = weekEndDate(Date)

    ' Неделя
    ThisWorkbook.Sheets("Лист4").Cells(2, 10).Value = CStr(WeekNumber(Date))

    ' Индиктор вызова процедур setК_порInЕСУП, currentК_порInЕСУП
    call_setК_порInЕСУП = False
    call_currentК_порInЕСУП = False

    ' Заносим даты в заголовки
    ThisWorkbook.Sheets("Лист4").Cells(2, 2).Value = "Выход вкладов на " + CStr(Date) + " г."
    Дата_следующего_воскресенья = Next_sunday_date(Date)
    ThisWorkbook.Sheets("Лист4").Cells(4, 3).Value = "Выход вкладов на неделе (" + CStr(WeekNumber(Date)) + ")                           с " + Mid(CStr(Date), 1, 5) + " по " + Mid(CStr(Дата_следующего_воскресенья), 1, 5)
    Дата_первого_понедельника_за_след_воскресеньем = Дата_следующего_воскресенья + 1
    
    ' Дата_конца_месяца = Date_last_day_month(Date)
    Дата_конца_месяца = Date_last_day_month(Дата_первого_понедельника_за_след_воскресеньем)
    ThisWorkbook.Sheets("Лист4").Cells(4, 5).Value = "Выход вкладов                             с " + Mid(CStr(Дата_первого_понедельника_за_след_воскресеньем), 1, 5) + " по " + Mid(CStr(Дата_конца_месяца), 1, 5)
    
    ' Заголовок столбца ЕСУП (неделя)
    If НеделяНаЛистеN("ЕСУП") = WeekNumber(Date) Then
      ' по ИСЖ переносим на Лист 4 (перенесено сюда с листа ЕСУП)
      ThisWorkbook.Sheets("ЕСУП").Cells(rowByValue(ThisWorkbook.Name, "ЕСУП", "Поручения периода:", 100, 100) + 2, ColumnByValue(ThisWorkbook.Name, "ЕСУП", "Поручения периода:", 100, 100) + 3).Value = "План нед.(" + CStr(WeekNumber(Date)) + ")"
      ' Факт
      ThisWorkbook.Sheets("ЕСУП").Cells(6, 6).Value = "Факт нед.(" + CStr(WeekNumber(Date)) + ")"
    End If
    
    ' Тема (исходящего письма) (!!!)
    ThisWorkbook.Sheets("Лист4").Cells(2, 16).Value = "Потенциал для отработки ИСЖ/НСЖ/Инвесты c " + Mid(CStr(Date), 1, 5) + " по " + Mid(CStr(Дата_следующего_воскресенья), 1, 5)
    
    ' Создаем выходную книгу для выгрузки вкладов
    OutBookName = ThisWorkbook.Path + "\Out\DepositsFinish_" + Mid(CStr(Date), 1, 5) + "_" + Mid(CStr(Дата_следующего_воскресенья), 1, 5) + ".xlsx"
    Call createBook_out_DepositsFinish(OutBookName)
    
    ' Файл-вложение (!!!)
    ThisWorkbook.Sheets("Лист4").Cells(3, 17).Value = OutBookName
    
    ' Счетчик записей в файле OutBookName
    RecCount_In_OutBookName = 1
    
    ' Переходим текущую форму
    ThisWorkbook.Sheets("Лист4").Activate
        
    ' Обработка отчета
    For i = 1 To 5
      ' Номера офисов от 1 до 5
      Select Case i
        Case 1
          ' ОО «Тюменский»
          officeNameInReport = "Тюменский"
        Case 2
          ' ОО «Сургутский»
          officeNameInReport = "Сургутский"
        Case 3
        ' ОО «Нижневартовский»
          officeNameInReport = "Нижневартовский"
        Case 4
        ' ОО «Новоуренгойский»
          officeNameInReport = "Новоуренгойский"
        Case 5
        ' ОО «Тарко-Сале»
          officeNameInReport = "Тарко-Сале"
      End Select
      ' Обнуление переменных
      Выход_вкладов_неделя_руб = 0
      Выход_вкладов_неделя_шт = 0
      Выход_вкладов_через_неделю_и_до_конца_месяца_руб = 0
      Выход_вкладов_через_неделю_и_до_конца_месяца_шт = 0
      ' Обработка отчета
      ' 1 - ИТОГО в нац.эквиваленте:
      ' ---------
      ' 2 - "Доп.офис:"
      ' 7 - Тюменский/Сургутский/...
      ' ---------
      ' 1 - номер договора
      ' 5 - фио
      ' 9 - счет
      ' 10 - вид вклада
      ' 14 - дата
      ' 15 - сумма

      Считаем_по_офису = False
      rowCount = 1
      Do While (InStr(Workbooks(DBstrName_String).Sheets("Page 1").Cells(rowCount, 1).Value, "ИТОГО в нац.эквиваленте:") = 0)
        
        ' Если найден текущий офис
        If (InStr(Workbooks(DBstrName_String).Sheets("Page 1").Cells(rowCount, 2).Value, "Доп.офис:") <> 0) And (InStr(Workbooks(DBstrName_String).Sheets("Page 1").Cells(rowCount, 7).Value, officeNameInReport) <> 0) Then
          ' Взводим переменную
          Считаем_по_офису = True
        End If
        ' Если идет другой офис
        If (InStr(Workbooks(DBstrName_String).Sheets("Page 1").Cells(rowCount, 2).Value, "Доп.офис:") <> 0) And (InStr(Workbooks(DBstrName_String).Sheets("Page 1").Cells(rowCount, 7).Value, officeNameInReport) = 0) Then
          ' Взводим переменную
          Считаем_по_офису = False
        End If
        
        ' Обработка записи - если в 3-м столбце есть счет 423*
        If (InStr(Workbooks(DBstrName_String).Sheets("Page 1").Cells(rowCount, 9).Value, "423") <> 0) Then
          
          ' Если вклад закрывается на неделе
          If (Считаем_по_офису = True) And (CDate(Workbooks(DBstrName_String).Sheets("Page 1").Cells(rowCount, 14).Value) >= Date) And (CDate(Workbooks(DBstrName_String).Sheets("Page 1").Cells(rowCount, 14).Value) <= Дата_следующего_воскресенья) Then
            Выход_вкладов_неделя_руб = Выход_вкладов_неделя_руб + CDbl(Replace(Workbooks(DBstrName_String).Sheets("Page 1").Cells(rowCount, 15).Value, ".", ","))
            Выход_вкладов_неделя_шт = Выход_вкладов_неделя_шт + 1
            RecCount_In_OutBookName = RecCount_In_OutBookName + 1
            ' Добавляем вклад в Книгу исходящих депозитов для офисов
            Call insert_to_DepositsFinish(OutBookName, _
                                            RecCount_In_OutBookName, _
                                              CDate(Workbooks(DBstrName_String).Sheets("Page 1").Cells(rowCount, 14).Value), _
                                                Workbooks(DBstrName_String).Sheets("Page 1").Cells(rowCount, 1).Value, _
                                                  Workbooks(DBstrName_String).Sheets("Page 1").Cells(rowCount, 9).Value, _
                                                    Workbooks(DBstrName_String).Sheets("Page 1").Cells(rowCount, 5).Value, _
                                                      CDbl(Replace(Workbooks(DBstrName_String).Sheets("Page 1").Cells(rowCount, 15).Value, ".", ",")), _
                                                        officeNameInReport)
          End If
          
          ' Если вклад открывается через неделю и до конца месяца
          If (Считаем_по_офису = True) And (CDate(Workbooks(DBstrName_String).Sheets("Page 1").Cells(rowCount, 14).Value) >= Дата_первого_понедельника_за_след_воскресеньем) And (CDate(Workbooks(DBstrName_String).Sheets("Page 1").Cells(rowCount, 14).Value) <= Дата_конца_месяца) Then
            Выход_вкладов_через_неделю_и_до_конца_месяца_руб = Выход_вкладов_через_неделю_и_до_конца_месяца_руб + CDbl(Replace(Workbooks(DBstrName_String).Sheets("Page 1").Cells(rowCount, 15).Value, ".", ","))
            Выход_вкладов_через_неделю_и_до_конца_месяца_шт = Выход_вкладов_через_неделю_и_до_конца_месяца_шт + 1
          End If
          
        End If
        
        ' Следующая запись
        rowCount = rowCount + 1
      Loop ' While
      
      ' Выводим итоги
      ThisWorkbook.Sheets("Лист4").Cells(5 + i, 3).Value = Выход_вкладов_неделя_шт
      ThisWorkbook.Sheets("Лист4").Cells(5 + i, 4).Value = Round(Выход_вкладов_неделя_руб / 1000, 2)
      '
      ThisWorkbook.Sheets("Лист4").Cells(5 + i, 5).Value = Выход_вкладов_через_неделю_и_до_конца_месяца_шт
      ThisWorkbook.Sheets("Лист4").Cells(5 + i, 6).Value = Round(Выход_вкладов_через_неделю_и_до_конца_месяца_руб / 1000, 2)
      
      ' ЕСУП - план на неделю
      If НеделяНаЛистеN("ЕСУП") = WeekNumber(Date) Then
        ' План ПК неделя
        ThisWorkbook.Sheets("ЕСУП").Cells(6 + i, 5).Value = Round((ThisWorkbook.Sheets("Лист4").Cells(5 + i, 4).Value / 100) * 18, 0)
        ' Поручение офису ВКЛИСЖi
        Call setК_порInЕСУП(ThisWorkbook.Name, "ЕСУП", "ВКЛИСЖ" + CStr(i), ThisWorkbook.Sheets("ЕСУП").Cells(6 + i, 5).Value, dateBeginWeek, "тыс.руб.", "ИСЖ к выходящим вкладам на неделю 18% (всего " + CStr(Выход_вкладов_неделя_шт) + " шт. на сумму " + CStr(Round(ThisWorkbook.Sheets("Лист4").Cells(5 + i, 4).Value, 0)) + " тыс.руб.)")
        ' Переменная, что процедуру setК_порInЕСУП вызывали
        call_setК_порInЕСУП = True
      End If
      
    Next i ' Следующий офис
    
    ' Выводим итоги
    
    ' Закрываем файл с отчетом без сохранения изменений (параметр SaveChanges:=False)
    Workbooks(Dir(FileName)).Close SaveChanges:=False
            
    ' Закрываем выходную книгу с выгрузкой
    Call sort_OutBookName_ByDate(OutBookName)
    Workbooks(Dir(OutBookName)).Close SaveChanges:=True
                        
    ' Закрываем базу BASE\Tasks
    CloseBook ("Tasks")
                                               
    ' Формируем список для отправки (в "Список получателей:"):
    ThisWorkbook.Sheets("Лист4").Cells(rowByValue(ThisWorkbook.Name, "Лист4", "Список получателей:", 100, 100), ColumnByValue(ThisWorkbook.Name, "Лист4", "Список получателей:", 100, 100) + 2).Value = getFromAddrBook("УДО2,УДО3,УДО4,УДО5,НОРПиКО1,НОРПиКО2,НОРПиКО3,НОРПиКО4,НОРПиКО5,МРК1", 2)
                                                                                              
    ' Переходим в ячейку M2
    ThisWorkbook.Sheets("Лист4").Cells(2, 13).Select

    ' Строка статуса
    Application.StatusBar = ""
    
    ' Зачеркиваем пункт меню на стартовой страницы
    ' Call ЗачеркиваемТекстВячейке("Лист0", "D5")
    ' Отчет по выходу вкладчиков на неделю
    Call ЗачеркиваемТекстВячейке("Лист0", RangeByValue(ThisWorkbook.Name, "Лист0", "Отчет по выходу вкладчиков на неделю", 100, 100))
    
    MsgBox ("Обработка " + DBstrName_String + " завершена!")

    ' Запрос на отправку шаблона письма
    If MsgBox("Отправить шаблон письма по вкладчикам?", vbYesNo) = vbYes Then
      
      Call Отправка_Lotus_Notes_Лист4_Вклады
      
    End If


  End If ' Если файл был выбран

End Sub



' Создание книги с выходящими вкладами
Sub createBook_out_DepositsFinish(In_OutBookName)

    Workbooks.Add
    ActiveWorkbook.SaveAs FileName:=In_OutBookName
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Activate
    
    ' Форматирование полей
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 1).Value = "Дата_окончания"
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns("A:A").EntireColumn.ColumnWidth = 20
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 1).HorizontalAlignment = xlCenter
    '
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 2).Value = "Номер_договора"
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns("B:B").EntireColumn.ColumnWidth = 21
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 2).HorizontalAlignment = xlCenter
    '
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 3).Value = "Номер_счета"
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns("C:C").EntireColumn.ColumnWidth = 22
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 3).HorizontalAlignment = xlCenter
    '
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 4).Value = "Клиент"
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns("D:D").EntireColumn.ColumnWidth = 50
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 4).HorizontalAlignment = xlCenter
    '
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 5).Value = "Сумма"
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns("E:E").EntireColumn.ColumnWidth = 20
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 5).HorizontalAlignment = xlCenter
    '
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 6).Value = "Офис"
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns("F:F").EntireColumn.ColumnWidth = 18
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 6).HorizontalAlignment = xlCenter
    '
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 7).Value = "МРК"
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns("G:G").EntireColumn.ColumnWidth = 15
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 7).HorizontalAlignment = xlCenter
    '
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 8).Value = "Продажа: 1(да)/ 0(нет)"
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns("H:H").EntireColumn.ColumnWidth = 20
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 8).HorizontalAlignment = xlCenter
    '
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 9).Value = "НК клиента"
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns("I:I").EntireColumn.ColumnWidth = 10
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 9).HorizontalAlignment = xlCenter
    '
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 10).Value = "Ссылка в CRM"
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns("J:J").EntireColumn.ColumnWidth = 50
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, 10).HorizontalAlignment = xlCenter

    ' ActiveCell.Offset(0, -4).Columns("A:A").EntireColumn.Select
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Range("C:C").Select
    ' Числовой
    ' Selection.NumberFormat = "0"
    ' Текстовый
    Selection.NumberFormat = "@"

    ' Установка фильтров
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Range("A1:J1").Select
    Selection.AutoFilter
    
End Sub

' Добавление записи в книгу с выходящими вкладами
Sub insert_to_DepositsFinish(In_OutBookName, In_RecCount, In_DateEnd, In_NumberDoc, In_NumberAcc, In_Client, In_Sum, In_Office)
  
  Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(In_RecCount, 1).Value = In_DateEnd
  Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(In_RecCount, 2).Value = In_NumberDoc
  Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(In_RecCount, 3).Value = In_NumberAcc
  Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(In_RecCount, 4).Value = In_Client
  Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(In_RecCount, 5).Value = In_Sum
  Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(In_RecCount, 6).Value = In_Office
    
End Sub

' Сортировка данных по дате
Sub sort_OutBookName_ByDate(In_OutBookName)

    ' Переходим на книгу с выгрузкой
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Activate

    ' ActiveCell.Offset(0, -7).Columns("A:A").EntireColumn.Select
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns("A:A").Select
    
    ' ActiveWorkbook.Worksheets("Лист1").Sort.SortFields.Clear
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Sort.SortFields.Clear
    
    ' ActiveWorkbook.Worksheets("Лист1").Sort.SortFields.Add Key:=ActiveCell.Offset(-1, 0).Range("A1"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Sort.SortFields.Add Key:=Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Range("A1"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Sort
        .SetRange ActiveCell.Range("A2:F53")
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Range("A1").Select

    ' Переходим обратно в главную книгу
    ThisWorkbook.Sheets("Лист4").Activate

End Sub

' Отправка письма: отправляю шаблон самому себе для последующей отправки в сеть письма на его основе:
Sub Отправка_Lotus_Notes_Лист4_Вклады()
Dim темаПисьма, текстПисьма, hashTag, attachmentFile As String
Dim i As Byte
  
  ' If MsgBox("Отправить себе Шаблон письма?", vbYesNo) = vbYes Then
    
    ' Тема письма - Тема:
    ' темаПисьма = ThisWorkbook.Sheets("Лист4").Cells(RowByValue(ThisWorkbook.Name, "Лист4", "Тема:", 100, 100), ColumnByValue(ThisWorkbook.Name, "Лист4", "Тема:", 100, 100) + 1).Value
    темаПисьма = subjectFromSheet("Лист4")

    ' hashTag - Хэштэг:
    ' hashTag = ThisWorkbook.Sheets("Лист4").Cells(RowByValue(ThisWorkbook.Name, "Лист4", "Хэштэг:", 100, 100), ColumnByValue(ThisWorkbook.Name, "Лист4", "Хэштэг:", 100, 100) + 1).Value
    ' hashTag - Хэштэг:
    hashTag = hashTagFromSheet("Лист4")

    ' Файл-вложение (!!!)
    attachmentFile = ThisWorkbook.Sheets("Лист4").Cells(3, 17).Value
    
    ' Текст письма
    текстПисьма = "" + Chr(13)
    текстПисьма = текстПисьма + "" + ThisWorkbook.Sheets("Лист4").Cells(rowByValue(ThisWorkbook.Name, "Лист4", "Список получателей:", 100, 100), ColumnByValue(ThisWorkbook.Name, "Лист4", "Список получателей:", 100, 100) + 2).Value + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "" + getFromAddrBook("РД", 2) + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "Уважаемые руководители," + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "Потенциал выхода вкладов до конца недели. Прошу в срок до " + CStr(Next_sunday_date(Date)) + " организовать работу по привлечению в ИСЖ/НСЖ/Инвесты с нормативом конверсии не менее 18%. Список вкладчиков во вложении." + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    ' Визитка (подпись С Ув., )
    текстПисьма = текстПисьма + ПодписьВПисьме()
    ' Хэштег
    текстПисьма = текстПисьма + createBlankStr(20) + hashTag
    ' Вызов
    Call send_Lotus_Notes(темаПисьма, "Sergey Fedorovich Proschaev/Tyumen/PSBank/Ru", "Sergey Fedorovich Proschaev/Tyumen/PSBank/Ru", текстПисьма, attachmentFile)
  
    ' Сообщение
    ' MsgBox ("Письмо отправлено!")
     
  ' End If
  
End Sub

' Отчет План/Факт за 18.03.2020 по продуктам ИСЖ_НСЖ
Sub Отчет_План_Факт_по_продуктам_ИСЖ_НСЖ()

' Описание переменных
Dim ReportName_String, officeNameInReport, CheckFormatReportResult, ИСЖ_Лист4_Range_str As String
Dim i, rowCount As Integer
Dim finishProcess, officeWasFound As Boolean
Dim column_Регион, column_ДО, column_План, column_Факт, ИСЖ_Лист4_Range_Row, ИСЖ_Лист4_Range_Column As Byte
Dim date_Отчет_План_Факт_по_продуктам_ИСЖ_НСЖ As Date

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
    ThisWorkbook.Sheets("Лист4").Activate

    ' Проверка формы отчета
    CheckFormatReportResult = CheckFormatReport(ReportName_String, "Отчет", 8, Date)
    
    If CheckFormatReportResult = "OK" Then
      
      ' Открываем BASE\Tasks
      OpenBookInBase ("Tasks")
    
      ' Открываем сводную таблицу по ИСЖ (Лист1) и НСЖ (Лист2)
      openPivotTables_Отчет_План_Факт_по_продуктам_ИСЖ_НСЖ (ReportName_String)
      
      ' Находим строку и столбец "ИСЖ, тыс. руб." на Лист4
      ИСЖ_Лист4_Range_str = RangeByValue(ThisWorkbook.Name, "Лист4", "ИСЖ, тыс. руб.", 100, 100)
      ИСЖ_Лист4_Range_Row = Workbooks(ThisWorkbook.Name).Sheets("Лист4").Range(ИСЖ_Лист4_Range_str).Row
      ИСЖ_Лист4_Range_Column = Workbooks(ThisWorkbook.Name).Sheets("Лист4").Range(ИСЖ_Лист4_Range_str).Column

      ' Дата отчета: "Отчет обновлен на 20.03.2020 за 20.03.2020 (16.30)" или "Отчет обновлен на 23.03.2020 за 20-22.03.2020 (полный день)"
      ' date_Отчет_План_Факт_по_продуктам_ИСЖ_НСЖ = CDate(Mid(Workbooks(ReportName_String).Sheets("Отчет").Range("E2").Value, 33, 10))
      date_Отчет_План_Факт_по_продуктам_ИСЖ_НСЖ = getDate_Отчет_План_Факт_по_продуктам_ИСЖ_НСЖ(Workbooks(ReportName_String).Sheets("Отчет").Range("E2").Value)
      
      ThisWorkbook.Sheets("Лист4").Cells(ИСЖ_Лист4_Range_Row - 2, ИСЖ_Лист4_Range_Column - 1).Value = "Оперативная бизнес-справка по продуктам ИСЖ, НСЖ на " + CStr(date_Отчет_План_Факт_по_продуктам_ИСЖ_НСЖ) + " г."
      
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

        ' 1-ый проход по сводной таблице: Обработка План_Факт ИСЖ на Лист1
        ' Находим номера столбцов:
        column_Регион = ColumnByName(ReportName_String, "Лист1", 1, "Регион")
        column_ДО = ColumnByName(ReportName_String, "Лист1", 1, "ДО")
        column_План = ColumnByName(ReportName_String, "Лист1", 1, "План")
        column_Факт = ColumnByName(ReportName_String, "Лист1", 1, "Факт")
        '
        rowCount = 2
        officeWasFound = False
        Do While (Not IsEmpty(Workbooks(ReportName_String).Sheets("Лист1").Cells(rowCount, 1).Value)) And (officeWasFound = False)
        
          ' Проверяем строку
          If InStr(Workbooks(ReportName_String).Sheets("Лист1").Cells(rowCount, column_ДО).Value, officeNameInReport) <> 0 Then
            
            ' Выводим данные на Лист4
            
            ' ИСЖ План
            ThisWorkbook.Sheets("Лист4").Cells(ИСЖ_Лист4_Range_Row + 1 + i, ИСЖ_Лист4_Range_Column).Value = Workbooks(ReportName_String).Sheets("Лист1").Cells(rowCount, column_План).Value
            ' ИСЖ Факт
            ThisWorkbook.Sheets("Лист4").Cells(ИСЖ_Лист4_Range_Row + 1 + i, ИСЖ_Лист4_Range_Column + 1).Value = Workbooks(ReportName_String).Sheets("Лист1").Cells(rowCount, column_Факт).Value
            ' % Исполнения
            ThisWorkbook.Sheets("Лист4").Cells(ИСЖ_Лист4_Range_Row + 1 + i, ИСЖ_Лист4_Range_Column + 2).Value = РассчетДоли(ThisWorkbook.Sheets("Лист4").Cells(ИСЖ_Лист4_Range_Row + 1 + i, ИСЖ_Лист4_Range_Column).Value, ThisWorkbook.Sheets("Лист4").Cells(ИСЖ_Лист4_Range_Row + 1 + i, ИСЖ_Лист4_Range_Column + 1).Value, 3)
            ThisWorkbook.Sheets("Лист4").Cells(ИСЖ_Лист4_Range_Row + 1 + i, ИСЖ_Лист4_Range_Column + 2).NumberFormat = cellsNumberFormat(ThisWorkbook.Sheets("Лист4").Cells(ИСЖ_Лист4_Range_Row + 1 + i, ИСЖ_Лист4_Range_Column + 2).Value)

            ' Заносим факт исполнения ЕСУП
            If НеделяНаЛистеN("ЕСУП") = WeekNumber(date_Отчет_План_Факт_по_продуктам_ИСЖ_НСЖ) Then
              Call currentК_порInЕСУП(ThisWorkbook.Name, "ЕСУП", "ВКЛИСЖ" + CStr(i), date_Отчет_План_Факт_по_продуктам_ИСЖ_НСЖ, ThisWorkbook.Sheets("Лист4").Cells(ИСЖ_Лист4_Range_Row + 1 + i, ИСЖ_Лист4_Range_Column + 1).Value, "тыс.руб.")
            End If

            ' офис был найден
            officeWasFound = True
          End If
        
          ' Следующая запись
          rowCount = rowCount + 1
          Application.StatusBar = officeNameInReport + ": " + CStr(rowCount) + "..."
          DoEventsInterval (rowCount)
        
        Loop ' по ИСЖ
         
        ' 2-ый проход по сводной таблице: Обработка План_Факт НСЖ на Лист2
        ' Находим номера столбцов:
        column_Регион = ColumnByName(ReportName_String, "Лист2", 1, "Регион")
        column_ДО = ColumnByName(ReportName_String, "Лист2", 1, "ДО")
        column_План = ColumnByName(ReportName_String, "Лист2", 1, "План")
        column_Факт = ColumnByName(ReportName_String, "Лист2", 1, "Факт")
        '
        rowCount = 2
        officeWasFound = False
        Do While (Not IsEmpty(Workbooks(ReportName_String).Sheets("Лист2").Cells(rowCount, 1).Value)) And (officeWasFound = False)
        
          ' Проверяем строку
          If InStr(Workbooks(ReportName_String).Sheets("Лист2").Cells(rowCount, column_ДО).Value, officeNameInReport) <> 0 Then
            
            ' Выводим данные на Лист4
            
            ' НСЖ План
            ThisWorkbook.Sheets("Лист4").Cells(ИСЖ_Лист4_Range_Row + 1 + i, ИСЖ_Лист4_Range_Column + 3).Value = Workbooks(ReportName_String).Sheets("Лист2").Cells(rowCount, column_План).Value
            ' НСЖ Факт
            ThisWorkbook.Sheets("Лист4").Cells(ИСЖ_Лист4_Range_Row + 1 + i, ИСЖ_Лист4_Range_Column + 4).Value = Workbooks(ReportName_String).Sheets("Лист2").Cells(rowCount, column_Факт).Value
            ' % Исполнения
            ThisWorkbook.Sheets("Лист4").Cells(ИСЖ_Лист4_Range_Row + 1 + i, ИСЖ_Лист4_Range_Column + 5).Value = РассчетДоли(ThisWorkbook.Sheets("Лист4").Cells(ИСЖ_Лист4_Range_Row + 1 + i, ИСЖ_Лист4_Range_Column + 3).Value, ThisWorkbook.Sheets("Лист4").Cells(ИСЖ_Лист4_Range_Row + 1 + i, ИСЖ_Лист4_Range_Column + 4).Value, 3)
            ThisWorkbook.Sheets("Лист4").Cells(ИСЖ_Лист4_Range_Row + 1 + i, ИСЖ_Лист4_Range_Column + 5).NumberFormat = cellsNumberFormat(ThisWorkbook.Sheets("Лист4").Cells(ИСЖ_Лист4_Range_Row + 1 + i, ИСЖ_Лист4_Range_Column + 5).Value)

            ' офис был найден
            officeWasFound = True
          End If
        
          ' Следующая запись
          rowCount = rowCount + 1
          Application.StatusBar = officeNameInReport + ": " + CStr(rowCount) + "..."
          DoEventsInterval (rowCount)
        
        Loop ' по НСЖ
      
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
    
    ' Закрываем базу BASE\Tasks
    CloseBook ("Tasks")
    
    ' Переходим в ячейку M2
    ThisWorkbook.Sheets("Лист4").Range("L1").Select

    ' Строка статуса
    Application.StatusBar = ""

    ' Зачеркиваем пункт меню на стартовой страницы
    Call ЗачеркиваемТекстВячейке("Лист0", RangeByValue(ThisWorkbook.Name, "Лист0", "Отчет План-Факт по продуктам ИСЖ_НСЖ", 100, 100))
    
    ' Итоговое сообщение
    If finishProcess = True Then
      MsgBox ("Обработка " + Dir(ReportName_String) + " завершена!")
    Else
      MsgBox ("Обработка отчета была прервана!")
    End If

  End If ' Если файл был выбран

End Sub

' Открытие сводных таблиц в "Отчет План/Факт за ДД.ММ.ГГГГ по продуктам ИСЖ_НСЖ"
Sub openPivotTables_Отчет_План_Факт_по_продуктам_ИСЖ_НСЖ(In_ReportName_String)
Dim rowCount As Integer
Dim список_открыт As Boolean
Dim Range_str, Продукт As String
Dim Range_Row, Range_Column, i As Byte

  ' Находим на листе "Отчет" ячейку "ИСЖ" (затем ячейку "НСЖ")
  For i = 1 To 2
        ' Вид продукта
        Select Case i
          Case 1 ' ИСЖ
            Продукт = "ИСЖ"
          Case 2 ' НСЖ
            Продукт = "НСЖ"
  End Select
  
  ' Находим на Листе Продукт
  Range_str = RangeByValue(In_ReportName_String, "Отчет", Продукт, 100, 100)
  Range_Row = Workbooks(In_ReportName_String).Sheets("Отчет").Range(Range_str).Row
  Range_Column = Workbooks(In_ReportName_String).Sheets("Отчет").Range(Range_str).Column

                
          ' Открываем все ячейки с "Валяев Сергей Николаевич" в столбце A (1)
          rowCount = Range_Row + 3
          список_открыт = False
          
          Do While (Workbooks(In_ReportName_String).Sheets("Отчет").Cells(rowCount, Range_Column - 1).Value <> "Общий итог") And (список_открыт = False)
            
            ' Проверяем ячейку
            If Trim(Workbooks(In_ReportName_String).Sheets("Отчет").Cells(rowCount, Range_Column - 1).Value) = "Валяев Сергей Николаевич" Then
              
              ' Раскрываем сводную таблицу
              Workbooks(In_ReportName_String).Sheets("Отчет").Cells(rowCount, Range_Column).ShowDetail = True
              
              ' Переменная открытия списка
              список_открыт = True
              
            End If
            
            ' Следующая запись
            rowCount = rowCount + 1
        
          Loop

  Next i


  ' Переходим на окно DB
  ThisWorkbook.Sheets("Лист4").Activate

End Sub
