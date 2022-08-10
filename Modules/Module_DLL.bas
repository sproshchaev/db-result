Attribute VB_Name = "Module_DLL"
' ----------------------------------------------------------------------------------------------------------------
' Здесь хранятся все пользовательские процедуры и функции, которые затем можно перенести в надстройку *
'
' * Создание надстроек: https://excelpedia.ru/bez-rubriki/kak-sozdat-polzovatelskuyu-funkciyu-v-excel-ispolzuya-vba
'
' ----------------------------------------------------------------------------------------------------------------

' 1. Определение даты начала недели
Function weekStartDate(In_Date) As Date
  ' Иницииализация результата
  weekStartDate = In_Date
  ' От текущей даты In_Date отнимаем один день, пока не получим понедельник
  Do While Weekday(weekStartDate, vbMonday) <> 1
    weekStartDate = weekStartDate - 1
  Loop
End Function

' 2. Дата конца недели
Function weekEndDate(In_Date) As Date
  ' Иницииализация результата
  weekEndDate = In_Date
  ' От текущей даты In_Date отнимаем один день, пока не получим понедельник
  Do While Weekday(weekEndDate, vbMonday) <> 7
    weekEndDate = weekEndDate + 1
  Loop
  ' Из цикла выходим на субботе прибавляем +1 день
  ' weekEndDate = weekEndDate + 1
End Function

' 3. Подстрока ДД.ММ из Даты
Function strDDMM(In_Date) As String
  strDDMM = Mid(CStr(In_Date), 1, 5)
End Function

' 4. Подстрока ДДММГГГГ из Даты
Function strDDMMYYYY(In_Date) As String
  strDDMMYYYY = Mid(CStr(In_Date), 1, 2) + Mid(CStr(In_Date), 4, 2) + Mid(CStr(In_Date), 7, 4)
End Function

' 4. Подстрока ДДММГГ из Даты
Function strDDMMYY(In_Date) As String
  strDDMMYY = Mid(CStr(In_Date), 1, 2) + Mid(CStr(In_Date), 4, 2) + Mid(CStr(In_Date), 9, 2)
End Function

' 4.1 Подстрока ММГГ из Даты
Function strMMYY(In_Date) As String
  strMMYY = Mid(CStr(In_Date), 4, 2) + Mid(CStr(In_Date), 9, 2)
End Function

' 4.1 Подстрока ГГ из Даты
Function strYY(In_Date) As String
  strYY = Mid(CStr(In_Date), 9, 2)
End Function

' 4.1.1 Дата начала месяца из Подстроки ММГГ
Function dateBeginFromStrMMYY(In_ММГГ) As Date
  dateBeginFromStrMMYY = CDate("01." + Mid(CStr(In_ММГГ), 1, 2) + ".20" + Mid(CStr(In_ММГГ), 3, 2))
End Function

' 4.1.2 Дата конца месяца из Подстроки ММГГ
Function dateEndFromStrMMYY(In_ММГГ) As Date
  dateEndFromStrMMYY = Date_last_day_month(CDate("01." + Mid(CStr(In_ММГГ), 1, 2) + ".20" + Mid(CStr(In_ММГГ), 3, 2)))
End Function

' 4.2. Функция firstMonthYear_strMMYY возвращает из 06.08.2020 строку "0120"
Function firstMonthYear_strMMYY(In_Date) As String
  firstMonthYear_strMMYY = "01" + Mid(CStr(In_Date), 9, 2)
End Function

' 4.3 Если сейчас 06.08.2020, то функция вернет 0720
Function pastMonth_strMMYY(In_Date) As String
  Месяц = Month(In_Date)
  ' Номер предидущего месяца
  Месяц = Месяц - 1
  ' Cтавим "0" перед месяцем, если месяц от 1 до 9
  If Месяц < 10 Then
    Месяц_str = "0" + CStr(Месяц)
  Else
    Месяц_str = CStr(Месяц)
  End If
  pastMonth_strMMYY = Месяц_str + Mid(CStr(In_Date), 9, 2)
End Function

' 4.4 Подстрока ДД-ММ-ГГ из Даты
Function strДД_MM_YY(In_Date) As String
  strДД_MM_YY = Mid(CStr(In_Date), 1, 2) + "-" + Mid(CStr(In_Date), 4, 2) + "-" + Mid(CStr(In_Date), 9, 2)
End Function

' 4.4.1 Подстрока ДД.ММ.ГГ из Даты
Function strДД_MM_YY2(In_Date) As String
  strДД_MM_YY2 = Mid(CStr(In_Date), 1, 2) + "." + Mid(CStr(In_Date), 4, 2) + "." + Mid(CStr(In_Date), 9, 2)
End Function

' 4.5 Подстрока ДД-ММ-ГГГГ из Даты
Function strДД_MM_YYYY(In_Date) As String
  strДД_MM_YYYY = Mid(CStr(In_Date), 1, 2) + "-" + Mid(CStr(In_Date), 4, 2) + "-" + Mid(CStr(In_Date), 7, 4)
End Function

' 5. Определение числа рабочих дней в полном месяца
Function Working_days_in_the_FullMonth(In_Date, In_working_days_in_the_week) As Integer
  ' Декодируем дату In_Date
  ' Месяц
  Месяц = Month(In_Date)
  ' Год
  Год = Year(In_Date)
  ' Первый день следующего месяца
  If Месяц = 12 Then
    Месяц = 0
    Год = Год + 1
  End If
  
  Первый_день_следующего_месяца = CDate("01." + CStr(Месяц + 1) + "." + CStr(Год))
  ' Дата начала месяца
  Текущая_дата_рассчета = CDate("01." + Mid(CStr(In_Date), 4, 7))
  
  ' Делаем рассчет по датам
  Working_days_in_the_FullMonth = 0
  Do While Текущая_дата_рассчета < Первый_день_следующего_месяца
    ' Если Текущая_дата_рассчета не суббота
    If In_working_days_in_the_week = 5 Then
      ' Если пятидневка
      If (Weekday(Текущая_дата_рассчета, vbMonday) <> 6) And (Weekday(Текущая_дата_рассчета, vbMonday) <> 7) Then
        Working_days_in_the_FullMonth = Working_days_in_the_FullMonth + 1
      End If
    Else
      ' Если шестидневка - In_working_days_in_the_week = 6
      If (Weekday(Текущая_дата_рассчета, vbMonday) <> 7) Then
        Working_days_in_the_FullMonth = Working_days_in_the_FullMonth + 1
      End If
    End If
    ' Следующая дата
    Текущая_дата_рассчета = Текущая_дата_рассчета + 1
  Loop ' Следующая дата
End Function

' 6. Число рабочих дней между двумя датами
Function Working_days_between_dates(In_DateStart, In_DateEnd, In_working_days_in_the_week) As Integer
  ' Инициализация счетчика числа рабочих дней
  Working_days_between_dates = 0
  
  ' Дата начала месяца
  Текущая_дата_рассчета = In_DateStart
  
  ' Делаем рассчет по датам
  Do While Текущая_дата_рассчета <= In_DateEnd
  
    ' Если Текущая_дата_рассчета не суббота
    If In_working_days_in_the_week = 5 Then
      ' Если пятидневка
      If (Weekday(Текущая_дата_рассчета, vbMonday) <> 6) And (Weekday(Текущая_дата_рассчета, vbMonday) <> 7) Then
        Working_days_between_dates = Working_days_between_dates + 1
      End If
    Else
      ' Если шестидневка - In_working_days_in_the_week = 6
      If (Weekday(Текущая_дата_рассчета, vbMonday) <> 7) Then
        Working_days_between_dates = Working_days_between_dates + 1
      End If
    End If
    
    ' Следующая дата
    Текущая_дата_рассчета = Текущая_дата_рассчета + 1
  
  Loop ' Следующая дата
  
End Function

' 6.1. Число рабочих дней между двумя датами с учетом праздников
Function Working_days_between_datesII(In_DateStart, In_DateEnd, In_working_days_in_the_week) As Integer
  
  ' Инициализация счетчика числа рабочих дней
  Working_days_between_datesII = 0
  
  ' Открываем таблицу нерабочих дней NonWorkingDays
  ' Открываем BASE\NonWorkingDays
  OpenBookInBase ("NonWorkingDays")

  ' Убираем фильтр, иначе поиск не по всей таблице
  If Workbooks("NonWorkingDays").Sheets("Лист1").AutoFilterMode = True Then
    ' Выключаем Автофильтр
    Workbooks("NonWorkingDays").Sheets("Лист1").Cells(1, 1).AutoFilter
  End If


  ' Дата начала месяца
  Текущая_дата_рассчета = In_DateStart
  
  ' Делаем рассчет по датам
  Do While Текущая_дата_рассчета <= In_DateEnd
  
    ' Выполняем поиск - Текущая_дата_рассчета есть в BASE\NonWorkingDays?
    Set searchResults = Workbooks("NonWorkingDays").Sheets("Лист1").Columns("A:A").Find(Текущая_дата_рассчета, LookAt:=xlWhole)
  
    ' Проверяем - есть ли такая дата, если нет, то добавляем
    If searchResults Is Nothing Then
      ' Если не найдена - вставляем
      Праздничный_день = False
    Else
      ' Если найдена
      Праздничный_день = True
    End If

    ' Если Текущая_дата_рассчета не суббота
    If In_working_days_in_the_week = 5 Then
      
      ' Если пятидневка
      If (Weekday(Текущая_дата_рассчета, vbMonday) <> 6) And (Weekday(Текущая_дата_рассчета, vbMonday) <> 7) Then
        
        If Праздничный_день = False Then
          Working_days_between_datesII = Working_days_between_datesII + 1
        End If
      
      End If
    
    Else
      
      ' Если шестидневка - In_working_days_in_the_week = 6
      If (Weekday(Текущая_дата_рассчета, vbMonday) <> 7) Then
        
        If Праздничный_день = False Then
          Working_days_between_datesII = Working_days_between_datesII + 1
        End If
        
      End If
    
    End If
    
    ' Следующая дата
    Текущая_дата_рассчета = Текущая_дата_рассчета + 1
  
  Loop ' Следующая дата
  
  ' Закрываем BASE\NonWorkingDays
  CloseBook ("NonWorkingDays")
  
End Function


' 7. Дата последнего дня месяца. Последний день месяца
Function Date_last_day_month(In_Date) As Date
Dim Первый_день_следующего_месяца As Date
  ' Декодируем дату In_DateNow
  ' Месяц
  Месяц = Month(In_Date)
  ' Год
  Год = Year(In_Date)
  ' Первый день следующего месяца
  If Месяц = 12 Then
    Месяц = 0
    Год = Год + 1
  End If
  Первый_день_следующего_месяца = CDate("01." + CStr(Месяц + 1) + "." + CStr(Год))
  Date_last_day_month = Первый_день_следующего_месяца - 1
End Function

' 7.1 Дата первого дня месяца (первый день месяца)
Function Date_begin_day_month(In_Date) As Date
Dim Первый_день_следующего_месяца As Date
  ' Декодируем дату In_DateNow
  ' Месяц
  Месяц = Month(In_Date)
  ' Год
  Год = Year(In_Date)
  ' Генерируем дату первого дня месяца
  Date_begin_day_month = CDate("01." + CStr(Месяц) + "." + CStr(Год))
End Function


' 7.2.0 Дата первого дня квартала
Function Date_begin_day_quarter(In_Date) As Date
  
  ' Декодируем дату In_Date
  ' Месяц
  Месяц = Month(In_Date)
  ' Год
  Год = Year(In_Date)
  
  ' Месяц преобразуем в первый месяц квартала
  Select Case Месяц
        ' 1 кв. - 01.01.YYYY
        Case 1, 2, 3
          Месяц_str = "01"
        ' 2 кв. - 01.04.YYYY
        Case 4, 5, 6
          Месяц_str = "04"
        ' 3 кв. - 01.07.YYYY
        Case 7, 8, 9
          Месяц_str = "07"
        ' 4 кв. - 01.10.YYYY
        Case 10, 11, 12
          Месяц_str = "10"
  End Select
  
  Date_begin_day_quarter = CDate("01." + Месяц_str + "." + CStr(Год))
  
End Function

' 7.2.1 Дата последнего дня квартала
Function Date_last_day_quarter(In_Date) As Date
Dim Первый_день_следующего_месяца As Date
  ' Декодируем дату In_Date
  ' Месяц
  Месяц = Month(In_Date)
  ' Год
  Год = Year(In_Date)
  
  ' Месяц преобразуем в последний месяц квартала
  Select Case Месяц
        ' 1 кв. - 01.01.YYYY
        Case 1, 2, 3
          Месяц = 3
        ' 2 кв. - 01.04.YYYY
        Case 4, 5, 6
          Месяц = 6
        ' 3 кв. - 01.07.YYYY
        Case 7, 8, 9
          Месяц = 9
        ' 4 кв. - 01.10.YYYY
        Case 10, 11, 12
          Месяц = 12
  End Select

  ' Первый день следующего месяца
  If Месяц = 12 Then
    Месяц = 0
    Год = Год + 1
  End If
  Первый_день_следующего_месяца = CDate("01." + CStr(Месяц + 1) + "." + CStr(Год))
  Date_last_day_quarter = Первый_день_следующего_месяца - 1
End Function

' 7.3 Преобразование номера месяца квартала в строку: Номер месяца в квартале: 1-"", 2-"2", 3-"3"
Function Nom_mes_quarter_str(In_Date) As String
  
  ' Месяц
  Месяц = Month(In_Date)
  
  ' Месяц преобразуем в последний месяц квартала
  Select Case Месяц
        ' 1-ый месяц в квартале
        Case 1, 4, 7, 10
          Nom_mes_quarter_str = ""
        ' 2-ый месяц в квартале
        Case 2, 5, 8, 11
          Nom_mes_quarter_str = "2"
        ' 3-ый месяц в квартале
        Case 3, 6, 9, 12
          Nom_mes_quarter_str = "3"
  
  End Select

End Function

' 7.4 Книга открыта?
Function BookIsOpen(In_BookName) As Boolean
Dim wbBook As Workbook

  ' Книга уже открыта?
  BookIsOpen = False
  
  ' Поиск по окнам - есть ли среди открытых?
  For Each wbBook In Workbooks
    If Windows(wbBook.Name).Visible Then
      
      ' t = wbBook.Name
      
      ' If wbBook.Name = In_BookName Then BookIsOpen = True: Exit For
      
      If InStr(wbBook.Name, In_BookName + ".") <> 0 Then BookIsOpen = True: Exit For
      
    End If
    
  Next wbBook

End Function

' 8. Открытие Книги из каталога BASE\
Sub OpenBookInBase(In_BookName)
Dim wbBook As Workbook
  
  ' Книга уже открыта?
  Книга_открыта = False
  
  ' Поиск по окнам - есть ли среди открытых?
  For Each wbBook In Workbooks
    If Windows(wbBook.Name).Visible Then
      
      ' If wbBook.Name = wbName Then Книга_открыта = True: Exit For
      If InStr(wbBook.Name, In_BookName + ".") <> 0 Then Книга_открыта = True: Exit For
      
    End If
  Next wbBook
  
  
  ' Если не открыта, то открываем Книгу
  If Книга_открыта = False Then
    
    ' Открываем базу BASE\Indicators.xlsx
    ' Workbooks.Open (ThisWorkbook.Path + "\Base\" + In_BookName + ".xlsx")
    
    ' Открываем базу BASE\Indicators.xlsx (UpdateLinks:=0)
    Workbooks.Open (ThisWorkbook.Path + "\Base\" + In_BookName + ".xlsx"), 0
    
    ' Открытие Книги в фоновом режиме - работает, но при ошибках придется закрывать весь Excel, чтобы переоткрыть заново или же проверять - открыта такая Книга или нет!
    ' Windows(In_BookName).Visible = True
    ThisWorkbook.Activate
  
    ' Если в открываемой Книге есть "Лист1", то убираем на нем Автофильтр (Например в Пр. календарь 2021 такого листа нет!)
    SheetName_String = FindNameSheet(In_BookName, "Лист1")
    If SheetName_String <> "" Then

      ' Убираем фильтр, иначе поиск не по всей таблице
      If Workbooks(In_BookName).Sheets("Лист1").AutoFilterMode = True Then
        ' Выключаем Автофильтр
        Workbooks(In_BookName).Sheets("Лист1").Cells(1, 1).AutoFilter
      End If
  
    End If
  
  End If
  
End Sub

' 9. Закрытие Книги в каталоге BASE\
Sub CloseBook(In_BookName)
Dim wbBook As Workbook

  ' Проверить - открыта ли Книга?
  
  ' Книга открыта?
  Книга_открыта = False
  
  ' Поиск по окнам - есть ли среди открытых?
  For Each wbBook In Workbooks
    If Windows(wbBook.Name).Visible Then
      
      ' If wbBook.Name = wbName Then Книга_открыта = True: Exit For
      If InStr(wbBook.Name, In_BookName + ".") <> 0 Then Книга_открыта = True: Exit For
      
    End If
  Next wbBook
  
  
  ' Если не открыта, то открываем Книгу
  If Книга_открыта = True Then
  
    ' Закрытие Книги
    Workbooks(In_BookName).Close SaveChanges:=True
    
  End If
  
End Sub

' 10. Вставка записи в открытую книгу до 20 полей
Sub InsertRecordInBook(In_BookName, In_Sheet, In_FieldKeyName, In_FieldKeyValue, In_FieldName1, In_FieldValue1, In_FieldName2, In_FieldValue2, In_FieldName3, In_FieldValue3, In_FieldName4, In_FieldValue4, In_FieldName5, In_FieldValue5, In_FieldName6, In_FieldValue6, In_FieldName7, In_FieldValue7, In_FieldName8, In_FieldValue8, In_FieldName9, In_FieldValue9, In_FieldName10, In_FieldValue10, In_FieldName11, In_FieldValue11, In_FieldName12, In_FieldValue12, In_FieldName13, In_FieldValue13, In_FieldName14, In_FieldValue14, In_FieldName15, In_FieldValue15, In_FieldName16, In_FieldValue16, In_FieldName17, In_FieldValue17, In_FieldName18, In_FieldValue18, In_FieldName19, In_FieldValue19, In_FieldName20, In_FieldValue20)
Dim rowCount As Integer
  
  ' Убираем фильтр, иначе поиск не по всей таблице
  If Workbooks(In_BookName).Sheets(In_Sheet).AutoFilterMode = True Then
    ' Выключаем Автофильтр
    Workbooks(In_BookName).Sheets(In_Sheet).Cells(1, 1).AutoFilter
  End If
 
  ' Проверяем наличие записи In_FieldKeyName - In_FieldKeyValue
  Литера_столбца = ConvertToLetter(ColumnByName(In_BookName, In_Sheet, 1, In_FieldKeyName))
  Set searchResults = Workbooks(In_BookName).Sheets(In_Sheet).Columns(Литера_столбца + ":" + Литера_столбца).Find(In_FieldKeyValue, LookAt:=xlWhole)
  
  ' Проверяем - есть ли такая дата, если нет, то добавляем
  If searchResults Is Nothing Then
    ' Если не найдена - вставляем
    rowCount = 2
    ' t = ColumnByName(In_BookName, In_Sheet, 1, In_FieldKeyName).Value
    Do While Not IsEmpty(Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, 1))
      ' Следующая запись
      rowCount = rowCount + 1
    Loop
  Else
    ' Если найдена, то апдейтим
    rowCount = searchResults.Row
  End If

  ' Заносим данные - In_FieldName1
  If In_FieldName1 <> "" Then
    Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, ColumnByName(In_BookName, In_Sheet, 1, In_FieldName1)).Value = In_FieldValue1
    Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, ColumnByName(In_BookName, In_Sheet, 1, In_FieldName1)).WrapText = False
  End If
  If In_FieldName2 <> "" Then
    Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, ColumnByName(In_BookName, In_Sheet, 1, In_FieldName2)).Value = In_FieldValue2
    Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, ColumnByName(In_BookName, In_Sheet, 1, In_FieldName2)).WrapText = False
  End If
  If In_FieldName3 <> "" Then
    Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, ColumnByName(In_BookName, In_Sheet, 1, In_FieldName3)).Value = In_FieldValue3
    Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, ColumnByName(In_BookName, In_Sheet, 1, In_FieldName3)).WrapText = False
  End If
  If In_FieldName4 <> "" Then
    Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, ColumnByName(In_BookName, In_Sheet, 1, In_FieldName4)).Value = In_FieldValue4
    Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, ColumnByName(In_BookName, In_Sheet, 1, In_FieldName4)).WrapText = False
  End If
  If In_FieldName5 <> "" Then
    Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, ColumnByName(In_BookName, In_Sheet, 1, In_FieldName5)).Value = In_FieldValue5
    Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, ColumnByName(In_BookName, In_Sheet, 1, In_FieldName5)).WrapText = False
  End If
  If In_FieldName6 <> "" Then
    Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, ColumnByName(In_BookName, In_Sheet, 1, In_FieldName6)).Value = In_FieldValue6
    Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, ColumnByName(In_BookName, In_Sheet, 1, In_FieldName6)).WrapText = False
  End If
  If In_FieldName7 <> "" Then
    Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, ColumnByName(In_BookName, In_Sheet, 1, In_FieldName7)).Value = In_FieldValue7
    Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, ColumnByName(In_BookName, In_Sheet, 1, In_FieldName7)).WrapText = False
  End If
  If In_FieldName8 <> "" Then
    Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, ColumnByName(In_BookName, In_Sheet, 1, In_FieldName8)).Value = In_FieldValue8
    Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, ColumnByName(In_BookName, In_Sheet, 1, In_FieldName8)).WrapText = False
  End If
  If In_FieldName9 <> "" Then
    Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, ColumnByName(In_BookName, In_Sheet, 1, In_FieldName9)).Value = In_FieldValue9
    Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, ColumnByName(In_BookName, In_Sheet, 1, In_FieldName9)).WrapText = False
  End If
  If In_FieldName10 <> "" Then
    Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, ColumnByName(In_BookName, In_Sheet, 1, In_FieldName10)).Value = In_FieldValue10
    Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, ColumnByName(In_BookName, In_Sheet, 1, In_FieldName10)).WrapText = False
  End If
  If In_FieldName11 <> "" Then
    Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, ColumnByName(In_BookName, In_Sheet, 1, In_FieldName11)).Value = In_FieldValue11
    Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, ColumnByName(In_BookName, In_Sheet, 1, In_FieldName11)).WrapText = False
  End If
  If In_FieldName12 <> "" Then
    Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, ColumnByName(In_BookName, In_Sheet, 1, In_FieldName12)).Value = In_FieldValue12
    Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, ColumnByName(In_BookName, In_Sheet, 1, In_FieldName12)).WrapText = False
  End If
  If In_FieldName13 <> "" Then
    Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, ColumnByName(In_BookName, In_Sheet, 1, In_FieldName13)).Value = In_FieldValue13
    Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, ColumnByName(In_BookName, In_Sheet, 1, In_FieldName13)).WrapText = False
  End If
  If In_FieldName14 <> "" Then
    Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, ColumnByName(In_BookName, In_Sheet, 1, In_FieldName14)).Value = In_FieldValue14
    Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, ColumnByName(In_BookName, In_Sheet, 1, In_FieldName14)).WrapText = False
  End If
  If In_FieldName15 <> "" Then
    Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, ColumnByName(In_BookName, In_Sheet, 1, In_FieldName15)).Value = In_FieldValue15
    Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, ColumnByName(In_BookName, In_Sheet, 1, In_FieldName15)).WrapText = False
  End If
  If In_FieldName16 <> "" Then
    Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, ColumnByName(In_BookName, In_Sheet, 1, In_FieldName16)).Value = In_FieldValue16
    Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, ColumnByName(In_BookName, In_Sheet, 1, In_FieldName16)).WrapText = False
  End If
  If In_FieldName17 <> "" Then
    Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, ColumnByName(In_BookName, In_Sheet, 1, In_FieldName17)).Value = In_FieldValue17
    Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, ColumnByName(In_BookName, In_Sheet, 1, In_FieldName17)).WrapText = False
  End If
  If In_FieldName18 <> "" Then
    Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, ColumnByName(In_BookName, In_Sheet, 1, In_FieldName18)).Value = In_FieldValue18
    Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, ColumnByName(In_BookName, In_Sheet, 1, In_FieldName18)).WrapText = False
  End If
  If In_FieldName19 <> "" Then
    Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, ColumnByName(In_BookName, In_Sheet, 1, In_FieldName19)).Value = In_FieldValue19
    Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, ColumnByName(In_BookName, In_Sheet, 1, In_FieldName19)).WrapText = False
  End If
  If In_FieldName20 <> "" Then
    Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, ColumnByName(In_BookName, In_Sheet, 1, In_FieldName20)).Value = In_FieldValue20
    Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, ColumnByName(In_BookName, In_Sheet, 1, In_FieldName20)).WrapText = False
  End If
  
End Sub

' 10.1 Вставка записи в открытую книгу до 28 полей
Sub InsertRecordInBook2(In_BookName, In_Sheet, In_FieldKeyName, In_FieldKeyValue, In_FieldName1, In_FieldValue1, In_FieldName2, In_FieldValue2, In_FieldName3, In_FieldValue3, In_FieldName4, In_FieldValue4, In_FieldName5, In_FieldValue5, In_FieldName6, In_FieldValue6, In_FieldName7, In_FieldValue7, In_FieldName8, In_FieldValue8, In_FieldName9, In_FieldValue9, In_FieldName10, In_FieldValue10, In_FieldName11, In_FieldValue11, In_FieldName12, In_FieldValue12, In_FieldName13, In_FieldValue13, In_FieldName14, In_FieldValue14, In_FieldName15, In_FieldValue15, In_FieldName16, In_FieldValue16, In_FieldName17, In_FieldValue17, In_FieldName18, In_FieldValue18, In_FieldName19, In_FieldValue19, In_FieldName20, In_FieldValue20, In_FieldName21, In_FieldValue21, In_FieldName22, In_FieldValue22, In_FieldName23, In_FieldValue23, In_FieldName24, In_FieldValue24, In_FieldName25, In_FieldValue25, In_FieldName26, In_FieldValue26, In_FieldName27, In_FieldValue27, In_FieldName28, In_FieldValue28)
Dim rowCount As Integer
  
  ' Убираем фильтр, иначе поиск не по всей таблице
  If Workbooks(In_BookName).Sheets(In_Sheet).AutoFilterMode = True Then
    ' Выключаем Автофильтр
    Workbooks(In_BookName).Sheets(In_Sheet).Cells(1, 1).AutoFilter
  End If
  
  ' Проверяем наличие записи In_FieldKeyName - In_FieldKeyValue
  Литера_столбца = ConvertToLetter(ColumnByName(In_BookName, In_Sheet, 1, In_FieldKeyName))
  Set searchResults = Workbooks(In_BookName).Sheets(In_Sheet).Columns(Литера_столбца + ":" + Литера_столбца).Find(In_FieldKeyValue, LookAt:=xlWhole)
  
  ' Проверяем - есть ли такая дата, если нет, то добавляем
  If searchResults Is Nothing Then
    ' Если не найдена - вставляем
    rowCount = 2
    ' t = ColumnByName(In_BookName, In_Sheet, 1, In_FieldKeyName).Value
    Do While Not IsEmpty(Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, 1))
      ' Следующая запись
      rowCount = rowCount + 1
    Loop
  Else
    ' Если найдена, то апдейтим
    rowCount = searchResults.Row
  End If

  ' Заносим данные - In_FieldName1
  If In_FieldName1 <> "" Then
    Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, ColumnByName(In_BookName, In_Sheet, 1, In_FieldName1)).Value = In_FieldValue1
    Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, ColumnByName(In_BookName, In_Sheet, 1, In_FieldName1)).WrapText = False
  End If
  If In_FieldName2 <> "" Then
    Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, ColumnByName(In_BookName, In_Sheet, 1, In_FieldName2)).Value = In_FieldValue2
    Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, ColumnByName(In_BookName, In_Sheet, 1, In_FieldName2)).WrapText = False
  End If
  If In_FieldName3 <> "" Then
    Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, ColumnByName(In_BookName, In_Sheet, 1, In_FieldName3)).Value = In_FieldValue3
    Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, ColumnByName(In_BookName, In_Sheet, 1, In_FieldName3)).WrapText = False
  End If
  If In_FieldName4 <> "" Then
    Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, ColumnByName(In_BookName, In_Sheet, 1, In_FieldName4)).Value = In_FieldValue4
    Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, ColumnByName(In_BookName, In_Sheet, 1, In_FieldName4)).WrapText = False
  End If
  If In_FieldName5 <> "" Then
    Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, ColumnByName(In_BookName, In_Sheet, 1, In_FieldName5)).Value = In_FieldValue5
    Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, ColumnByName(In_BookName, In_Sheet, 1, In_FieldName5)).WrapText = False
  End If
  If In_FieldName6 <> "" Then
    Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, ColumnByName(In_BookName, In_Sheet, 1, In_FieldName6)).Value = In_FieldValue6
    Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, ColumnByName(In_BookName, In_Sheet, 1, In_FieldName6)).WrapText = False
  End If
  If In_FieldName7 <> "" Then
    Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, ColumnByName(In_BookName, In_Sheet, 1, In_FieldName7)).Value = In_FieldValue7
    Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, ColumnByName(In_BookName, In_Sheet, 1, In_FieldName7)).WrapText = False
  End If
  If In_FieldName8 <> "" Then
    Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, ColumnByName(In_BookName, In_Sheet, 1, In_FieldName8)).Value = In_FieldValue8
    Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, ColumnByName(In_BookName, In_Sheet, 1, In_FieldName8)).WrapText = False
  End If
  If In_FieldName9 <> "" Then
    Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, ColumnByName(In_BookName, In_Sheet, 1, In_FieldName9)).Value = In_FieldValue9
    Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, ColumnByName(In_BookName, In_Sheet, 1, In_FieldName9)).WrapText = False
  End If
  If In_FieldName10 <> "" Then
    Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, ColumnByName(In_BookName, In_Sheet, 1, In_FieldName10)).Value = In_FieldValue10
    Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, ColumnByName(In_BookName, In_Sheet, 1, In_FieldName10)).WrapText = False
  End If
  If In_FieldName11 <> "" Then
    Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, ColumnByName(In_BookName, In_Sheet, 1, In_FieldName11)).Value = In_FieldValue11
    Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, ColumnByName(In_BookName, In_Sheet, 1, In_FieldName11)).WrapText = False
  End If
  If In_FieldName12 <> "" Then
    Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, ColumnByName(In_BookName, In_Sheet, 1, In_FieldName12)).Value = In_FieldValue12
    Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, ColumnByName(In_BookName, In_Sheet, 1, In_FieldName12)).WrapText = False
  End If
  If In_FieldName13 <> "" Then
    Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, ColumnByName(In_BookName, In_Sheet, 1, In_FieldName13)).Value = In_FieldValue13
    Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, ColumnByName(In_BookName, In_Sheet, 1, In_FieldName13)).WrapText = False
  End If
  If In_FieldName14 <> "" Then
    Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, ColumnByName(In_BookName, In_Sheet, 1, In_FieldName14)).Value = In_FieldValue14
    Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, ColumnByName(In_BookName, In_Sheet, 1, In_FieldName14)).WrapText = False
  End If
  If In_FieldName15 <> "" Then
    Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, ColumnByName(In_BookName, In_Sheet, 1, In_FieldName15)).Value = In_FieldValue15
    Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, ColumnByName(In_BookName, In_Sheet, 1, In_FieldName15)).WrapText = False
  End If
  If In_FieldName16 <> "" Then
    Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, ColumnByName(In_BookName, In_Sheet, 1, In_FieldName16)).Value = In_FieldValue16
    Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, ColumnByName(In_BookName, In_Sheet, 1, In_FieldName16)).WrapText = False
  End If
  If In_FieldName17 <> "" Then
    Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, ColumnByName(In_BookName, In_Sheet, 1, In_FieldName17)).Value = In_FieldValue17
    Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, ColumnByName(In_BookName, In_Sheet, 1, In_FieldName17)).WrapText = False
  End If
  If In_FieldName18 <> "" Then
    Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, ColumnByName(In_BookName, In_Sheet, 1, In_FieldName18)).Value = In_FieldValue18
    Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, ColumnByName(In_BookName, In_Sheet, 1, In_FieldName18)).WrapText = False
  End If
  If In_FieldName19 <> "" Then
    Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, ColumnByName(In_BookName, In_Sheet, 1, In_FieldName19)).Value = In_FieldValue19
    Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, ColumnByName(In_BookName, In_Sheet, 1, In_FieldName19)).WrapText = False
  End If
  If In_FieldName20 <> "" Then
    Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, ColumnByName(In_BookName, In_Sheet, 1, In_FieldName20)).Value = In_FieldValue20
    Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, ColumnByName(In_BookName, In_Sheet, 1, In_FieldName20)).WrapText = False
  End If
  If In_FieldName21 <> "" Then
    Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, ColumnByName(In_BookName, In_Sheet, 1, In_FieldName21)).Value = In_FieldValue21
    Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, ColumnByName(In_BookName, In_Sheet, 1, In_FieldName21)).WrapText = False
  End If
  If In_FieldName22 <> "" Then
    Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, ColumnByName(In_BookName, In_Sheet, 1, In_FieldName22)).Value = In_FieldValue22
    Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, ColumnByName(In_BookName, In_Sheet, 1, In_FieldName22)).WrapText = False
  End If
  If In_FieldName23 <> "" Then
    Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, ColumnByName(In_BookName, In_Sheet, 1, In_FieldName23)).Value = In_FieldValue23
    Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, ColumnByName(In_BookName, In_Sheet, 1, In_FieldName23)).WrapText = False
  End If
  If In_FieldName24 <> "" Then
    Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, ColumnByName(In_BookName, In_Sheet, 1, In_FieldName24)).Value = In_FieldValue24
    Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, ColumnByName(In_BookName, In_Sheet, 1, In_FieldName24)).WrapText = False
  End If
  If In_FieldName25 <> "" Then
    Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, ColumnByName(In_BookName, In_Sheet, 1, In_FieldName25)).Value = In_FieldValue25
    Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, ColumnByName(In_BookName, In_Sheet, 1, In_FieldName25)).WrapText = False
  End If
  If In_FieldName26 <> "" Then
    Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, ColumnByName(In_BookName, In_Sheet, 1, In_FieldName26)).Value = In_FieldValue26
    Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, ColumnByName(In_BookName, In_Sheet, 1, In_FieldName26)).WrapText = False
  End If
  If In_FieldName27 <> "" Then
    Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, ColumnByName(In_BookName, In_Sheet, 1, In_FieldName27)).Value = In_FieldValue27
    Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, ColumnByName(In_BookName, In_Sheet, 1, In_FieldName27)).WrapText = False
  End If
  If In_FieldName28 <> "" Then
    Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, ColumnByName(In_BookName, In_Sheet, 1, In_FieldName28)).Value = In_FieldValue28
    Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, ColumnByName(In_BookName, In_Sheet, 1, In_FieldName28)).WrapText = False
  End If
  
End Sub


' 11. Номер недели в году по Дате (по календарю Лотуса совпадает, по календарям в Интернете нужно от номера недели отнимать 1)
Function WeekNumber(In_Date) As Integer
Dim CurrDate As Date
  ' Дата начала года - есть первая неделя
  CurrDate = CDate("01.01." + CStr(Year(In_Date)))
  
  ' В календаре Лотуса первая неделя не зависимо от того с какого дня она начинается - считается, как первая! В календарях Интернета первая неделя считается с первого понедельника
  WeekNumber = 1
  
  Do While CurrDate <= In_Date
    ' Если текущая дата CurrDate это понедельник, то считаем плюс неделю
    If Weekday(CurrDate, vbMonday) = 1 Then
      WeekNumber = WeekNumber + 1
    End If
    ' Следующая дата
    CurrDate = CurrDate + 1
  Loop
End Function

' 11.0 Псевдонимы функции на русском
Function Номер_недели(In_Date) As Integer
  Номер_недели = WeekNumber(In_Date)
End Function

' 11.1 Дата понедельника по номеру недели
Function MondayDateByWeekNumber(In_WeekNumber, In_Year) As Date
Dim CurrDate As Date
Dim WeekNumber As Byte
  
  ' Берем 01 января текущего года - это первая неделя
  CurrDate = CDate("01.01." + CStr(In_Year))
  MondayDateByWeekNumber = CurrDate
  
  WeekNumber = 1
  
  Do While WeekNumber <= In_WeekNumber
    
    ' Если текущая дата CurrDate это понедельник, то считаем плюс неделю
    If Weekday(CurrDate, vbMonday) = 1 Then
      WeekNumber = WeekNumber + 1
      '
      If WeekNumber = In_WeekNumber Then
        MondayDateByWeekNumber = CurrDate
      End If
      
    End If
    
    ' Следующая дата
    CurrDate = CurrDate + 1
    
  Loop
    
End Function


' 12. Наименование дня недели по дате
Function ДеньНедели(In_Date) As String
  ' День недели
  Select Case Weekday(In_Date, vbMonday)
    Case 1
          ДеньНедели = "понедельник"
        Case 2
          ДеньНедели = "вторник"
        Case 3
          ДеньНедели = "среда"
        Case 4
          ДеньНедели = "четверг"
        Case 5
          ДеньНедели = "пятница"
        Case 6
          ДеньНедели = "суббота"
        Case 7
          ДеньНедели = "воскресенье"
      End Select
End Function

' 12.1 Наименование дня недели по дате
Function День_Недели(In_Date) As String
  
  День_Недели = ДеньНедели(In_Date)
  
End Function


' 12.1 Остаток рабочих дней недели - строкой
Function remDayWorkWeek(In_Date) As String
  
  remDayWorkWeek = ""
  
  ' День недели
  Select Case Weekday(In_Date, vbMonday)
    Case 1
          remDayWorkWeek = "неделю"
        Case 2
          remDayWorkWeek = "4 дня"
        Case 3
          remDayWorkWeek = "3 дня"
        Case 4
          remDayWorkWeek = "2 оставшихся дня"
        Case 5
          remDayWorkWeek = "сегодня"
        Case 6
          remDayWorkWeek = ""
        Case 7
          remDayWorkWeek = ""
      End Select
      
End Function


' 13. Наименование месяца по дате
Function ИмяМесяца(In_Date) As String
  ' Месяц
  Select Case Month(In_Date)
    Case 1
          ИмяМесяца = "январь"
        Case 2
          ИмяМесяца = "февраль"
        Case 3
          ИмяМесяца = "март"
        Case 4
          ИмяМесяца = "апрель"
        Case 5
          ИмяМесяца = "май"
        Case 6
          ИмяМесяца = "июнь"
        Case 7
          ИмяМесяца = "июль"
        Case 8
          ИмяМесяца = "август"
        Case 9
          ИмяМесяца = "сентябрь"
        Case 10
          ИмяМесяца = "октябрь"
        Case 11
          ИмяМесяца = "ноябрь"
        Case 12
          ИмяМесяца = "декабрь"
      End Select
End Function

' 13.2 Наименование месяца по дате
Function ИмяМесяца2(In_Date) As String
  ' Месяц
  Select Case Month(In_Date)
    Case 1
          ИмяМесяца2 = "января"
        Case 2
          ИмяМесяца2 = "февраля"
        Case 3
          ИмяМесяца2 = "марта"
        Case 4
          ИмяМесяца2 = "апреля"
        Case 5
          ИмяМесяца2 = "мая"
        Case 6
          ИмяМесяца2 = "июня"
        Case 7
          ИмяМесяца2 = "июля"
        Case 8
          ИмяМесяца2 = "августа"
        Case 9
          ИмяМесяца2 = "сентября"
        Case 10
          ИмяМесяца2 = "октября"
        Case 11
          ИмяМесяца2 = "ноября"
        Case 12
          ИмяМесяца2 = "декабря"
      End Select
End Function

' 13.3 Наименование месяца по дате
Function ИмяМесяца3(In_Date) As String
  ' Месяц
  Select Case Month(In_Date)
    Case 1
          ИмяМесяца3 = "Январь"
        Case 2
          ИмяМесяца3 = "Февраль"
        Case 3
          ИмяМесяца3 = "Март"
        Case 4
          ИмяМесяца3 = "Апрель"
        Case 5
          ИмяМесяца3 = "Май"
        Case 6
          ИмяМесяца3 = "Июнь"
        Case 7
          ИмяМесяца3 = "Июль"
        Case 8
          ИмяМесяца3 = "Август"
        Case 9
          ИмяМесяца3 = "Сентябрь"
        Case 10
          ИмяМесяца3 = "Октябрь"
        Case 11
          ИмяМесяца3 = "Ноябрь"
        Case 12
          ИмяМесяца3 = "Декабрь"
      End Select
End Function


' 13.2 Наименование месяца и год по дате
Function ИмяМесяцаГод(In_Date) As String
  ИмяМесяцаГод = ИмяМесяца(In_Date) + " " + CStr(Year(In_Date)) + " г."
End Function

' 13.3 Наименование месяца и год по дате
Function ДеньМесяцГод(In_Date) As String
  ДеньМесяцГод = CStr(Day(In_Date)) + " " + ИмяМесяца2(In_Date) + " " + CStr(Year(In_Date)) + " г."
End Function


' 14. DoEvents
Function DoEventsInterval(In_Value)
  ' If InStr(CStr(In_Value / 100), ",") = 0 Then
  ' Получаем остаток от деления и если он равен нулю, запускаем DoEvents
  If x Mod 100 = 0 Then
    DoEvents
  End If
End Function

' 15. Зачеркиваем текст в ячейке
Sub ЗачеркиваемТекстВячейке(In_Sheets, In_Range)
    ' Зачеркиваем пункт меню на стартовой страницы
    ThisWorkbook.Sheets(In_Sheets).Range(In_Range).Font.Name = "Calibri"
    ThisWorkbook.Sheets(In_Sheets).Range(In_Range).Font.FontStyle = "полужирный"
    ThisWorkbook.Sheets(In_Sheets).Range(In_Range).Font.Size = 12
    ThisWorkbook.Sheets(In_Sheets).Range(In_Range).Font.Strikethrough = True
    ThisWorkbook.Sheets(In_Sheets).Range(In_Range).Font.Superscript = False
    ThisWorkbook.Sheets(In_Sheets).Range(In_Range).Font.Subscript = False
    ThisWorkbook.Sheets(In_Sheets).Range(In_Range).Font.OutlineFont = False
    ThisWorkbook.Sheets(In_Sheets).Range(In_Range).Font.Shadow = False
    ThisWorkbook.Sheets(In_Sheets).Range(In_Range).Font.Underline = xlUnderlineStyleNone
    ThisWorkbook.Sheets(In_Sheets).Range(In_Range).Font.ThemeColor = xlThemeColorLight1
    ThisWorkbook.Sheets(In_Sheets).Range(In_Range).Font.TintAndShade = 0
    ThisWorkbook.Sheets(In_Sheets).Range(In_Range).Font.ThemeFont = xlThemeFontMinor
    ThisWorkbook.Sheets(In_Sheets).Range(In_Range).Font.Bold = False
End Sub

' 16. Выделение жирным текста в ячейке
Sub ВыделениеЖирнымТекстаВячейке(In_Sheets, In_Range)
  ThisWorkbook.Sheets(In_Sheets).Range(In_Range).Font.Name = "Calibri"
  ThisWorkbook.Sheets(In_Sheets).Range(In_Range).Font.FontStyle = "обычный"
  ThisWorkbook.Sheets(In_Sheets).Range(In_Range).Font.Size = 12
  ThisWorkbook.Sheets(In_Sheets).Range(In_Range).Font.Strikethrough = False
  ThisWorkbook.Sheets(In_Sheets).Range(In_Range).Font.Superscript = False
  ThisWorkbook.Sheets(In_Sheets).Range(In_Range).Font.Subscript = False
  ThisWorkbook.Sheets(In_Sheets).Range(In_Range).Font.OutlineFont = False
  ThisWorkbook.Sheets(In_Sheets).Range(In_Range).Font.Shadow = False
  ThisWorkbook.Sheets(In_Sheets).Range(In_Range).Font.Underline = xlUnderlineStyleNone
  ThisWorkbook.Sheets(In_Sheets).Range(In_Range).Font.ThemeColor = xlThemeColorLight1
  ThisWorkbook.Sheets(In_Sheets).Range(In_Range).Font.TintAndShade = 0
  ThisWorkbook.Sheets(In_Sheets).Range(In_Range).Font.ThemeFont = xlThemeFontMinor
  ThisWorkbook.Sheets(In_Sheets).Range(In_Range).Font.Bold = True
End Sub

' 17. Выделение бледным текста в ячейки
Sub ВыделениеБледнымТекстаВячейке(In_Sheets, In_Range)
    ThisWorkbook.Sheets(In_Sheets).Range(In_Range).Font.Bold = False
    ThisWorkbook.Sheets(In_Sheets).Range(In_Range).Font.ThemeColor = xlThemeColorDark1
    ThisWorkbook.Sheets(In_Sheets).Range(In_Range).Font.TintAndShade = -4.99893185216834E-02
End Sub

' 18. Отправка почтой
Sub send_Lotus_Notes(In_Subject, In_Address, In_AddressCopy, In_body, strpath As String)

  Dim Session As Object
  Dim Dir As Object
  Dim Doc As Object
  Dim Workspace As Object
  Dim EditDoc As Object
  Dim AttachME As Object
  Dim UserName As String
  Dim MailDbName As String

  Set Workspace = CreateObject("Notes.NotesUIWorkspace")
  Set Session = CreateObject("notes.NOTESSESSION")
  UserName = Session.UserName
  MailDbName = Left$(UserName, 1) & Right$(UserName, (Len(UserName) - InStr(1, UserName, " "))) & ".nsf"
  Set Dir = Session.CurrentDatabase
  If Dir.IsOpen = False Then Call Dir.OpenMail
  Set Doc = Dir.CREATEDOCUMENT

  ' Тема
  Doc.Subject = In_Subject ' "Тест"
  ' Адрес
  Doc.SendTo = In_Address ' "Sergey Fedorovich Proschaev/Tyumen/PSBank/Ru"
  Doc.CopyTo = In_AddressCopy
  ' BlindCopyTo - это взял из примера, проверить!
  ' Doc.BlindCopyTo = ""
  
  ' Письмо
  Doc.body = In_body ' "Добрый день!"

  Attachment = strpath
  If Attachment <> "" Then
      Set AttachME = Doc.CREATERICHTEXTITEM("Attachment" & i)
      Set EmbedObj = AttachME.EmbedObject(1454, "", Attachment, "Attachment")
  End If

  Doc.SAVEMESSAGEONSEND = SaveIt
  
  ' Если ремарим это - будет отправка?
  Doc.send 0, In_Address ' "Sergey Fedorovich Proschaev/Tyumen/PSBank/Ru"

  Set Session = Nothing
  Set Dir = Nothing
  Set Doc = Nothing
  Set Workspace = Nothing
  Set EditDoc = Nothing

End Sub

' 18.1 Отправка почтой
Sub send_Lotus_Notes2(In_Subject, In_Address, In_AddressCopy, In_AddressBlind, In_body, strpath As String)

  Dim Session As Object
  Dim Dir As Object
  Dim Doc As Object
  Dim Workspace As Object
  Dim EditDoc As Object
  Dim AttachME As Object
  Dim UserName As String
  Dim MailDbName As String

  Set Workspace = CreateObject("Notes.NotesUIWorkspace")
  Set Session = CreateObject("notes.NOTESSESSION")
  UserName = Session.UserName
  MailDbName = Left$(UserName, 1) & Right$(UserName, (Len(UserName) - InStr(1, UserName, " "))) & ".nsf"
  Set Dir = Session.CurrentDatabase
  If Dir.IsOpen = False Then Call Dir.OpenMail
  Set Doc = Dir.CREATEDOCUMENT

  ' Тема
  Doc.Subject = In_Subject ' "Тест"
  ' Адрес
  Doc.SendTo = In_Address ' "Sergey Fedorovich Proschaev/Tyumen/PSBank/Ru"
  Doc.CopyTo = In_AddressCopy
  ' BlindCopyTo - это взял из примера, проверить!
  Doc.BlindCopyTo = In_AddressBlind
  
  ' Письмо
  Doc.body = In_body ' "Добрый день!"

  Attachment = strpath
  If Attachment <> "" Then
      Set AttachME = Doc.CREATERICHTEXTITEM("Attachment" & i)
      Set EmbedObj = AttachME.EmbedObject(1454, "", Attachment, "Attachment")
  End If

  Doc.SAVEMESSAGEONSEND = SaveIt
  
  ' Если ремарим это - будет отправка?
  Doc.send 0, In_Address ' "Sergey Fedorovich Proschaev/Tyumen/PSBank/Ru"

  Set Session = Nothing
  Set Dir = Nothing
  Set Doc = Nothing
  Set Workspace = Nothing
  Set EditDoc = Nothing

End Sub


' 19 Подпись для письма
Function ПодписьВПисьме() As String
  ' Визитка
  ПодписьВПисьме = ПодписьВПисьме + "" + Chr(13)
  ПодписьВПисьме = ПодписьВПисьме + "С уважением, Прощаев Сергей Федорович" + Chr(13)
  ПодписьВПисьме = ПодписьВПисьме + "Заместитель регионального директора" + Chr(13)
  ПодписьВПисьме = ПодписьВПисьме + "Региональный Операционный Офис «Тюменский»" + Chr(13)
  ПодписьВПисьме = ПодписьВПисьме + "ПАО «Промсвязьбанк»" + Chr(13)
  ПодписьВПисьме = ПодписьВПисьме + "e-mail: proschaevsf@ tyumen.psbank.ru" + Chr(13)
  ПодписьВПисьме = ПодписьВПисьме + "тел.: вн. 71-5913" + Chr(13)
  ПодписьВПисьме = ПодписьВПисьме + "моб.: +7 (922) 00-88-253" + Chr(13)
  ' ПодписьВПисьме = ПодписьВПисьме + "" + Chr(13)
End Function

' 20. Номер столбца в таблице по "Названию"
Function ColumnByName(In_Workbooks, In_Sheets, In_Row, In_ColumnName) As Integer
Dim i As Integer
  ColumnByName = 0
  ' Выполняем поиск номера столбца
  i = 1
  Найден_столбец = False
  Do While (Not IsEmpty(Workbooks(In_Workbooks).Sheets(In_Sheets).Cells(In_Row, i).Value)) And (Найден_столбец = False)
    ' Проверяем - значение столбца и имя которое ищем
    If Trim(Workbooks(In_Workbooks).Sheets(In_Sheets).Cells(In_Row, i).Value) = Trim(In_ColumnName) Then
      Найден_столбец = True
      ColumnByName = i
    End If
    ' Следующий столбец
    i = i + 1
  Loop
End Function

' 20.1 Номер столбца в таблице по "Названию" и по номеру такого названия с начала: План, План, План, План
Function ColumnByNameAndNumber(In_Workbooks, In_Sheets, In_Row, In_ColumnName, In_ColumnNameCount, In_maxColumnInSheet) As Integer
Dim i As Integer
  ColumnByNameAndNumber = 0
  ' Выполняем поиск номера столбца
  i = 1
  Найден_столбец = False
  ' Число найденых столбцов с таким именем (In_ColumnName)
  ЧислоНайденыхСтолбцов = 0
  Do While (i < In_maxColumnInSheet) And (Найден_столбец = False)
    ' Проверяем - значение столбца и имя которое ищем
    If Trim(Workbooks(In_Workbooks).Sheets(In_Sheets).Cells(In_Row, i).Value) = Trim(In_ColumnName) Then
      ЧислоНайденыхСтолбцов = ЧислоНайденыхСтолбцов + 1
      ' Число найденых таких столбцов
      If ЧислоНайденыхСтолбцов = In_ColumnNameCount Then
        ' Заканчиваем поиск
        Найден_столбец = True
      End If
      ColumnByNameAndNumber = i
    End If
    ' Следующий столбец
    i = i + 1
  Loop
End Function

' 21. Определение строки ячейки в которой находится заданное значение (текст)
Function rowByValue(In_Workbooks, In_Sheets, In_Value, In_maxRowInSheet, In_maxColumnInSheet) As Integer
  rowByValue = 0
  ColumnCount = 1
  Найдено_значение = False
  
  ' Двигаемся сначала по столбцу, потом по строке
  ' 23.12.2020 Do While (ColumnCount < In_maxColumnInSheet) And (Найдено_значение = False)
  Do While (ColumnCount <= In_maxColumnInSheet) And (Найдено_значение = False)
  
    ' По срокам
    rowCount = 1
    
    ' 23.12.2020 Do While (rowCount < In_maxRowInSheet) And (Найдено_значение = False)
    Do While (rowCount <= In_maxRowInSheet) And (Найдено_значение = False)
      ' t = Workbooks(In_Workbooks).Sheets(In_Sheets).Cells(rowCount, ColumnCount).Value
      ' t2 = InStr(CStr(t), "Error")
      If Trim(CStr(Workbooks(In_Workbooks).Sheets(In_Sheets).Cells(rowCount, ColumnCount).Value)) = Trim(In_Value) Then
      ' If Trim(t) = Trim(In_Value) Then
        rowByValue = rowCount
        Найдено_значение = True
      End If
      ' Следующая строка
      rowCount = rowCount + 1
    Loop
    ' Следующий столбец
    ColumnCount = ColumnCount + 1
  Loop
End Function

' 21.1 Определение строки ячейки в которой находится заданное значение (текст)
Function rowByValue2(In_Workbooks, In_Sheets, In_Value, In_maxRowInSheet, In_maxColumnInSheet, In_Count) As Integer
  rowByValue2 = 0
  ColumnCount = 1
  Найдено_значение = False
  Число_найденных = 0
  
  ' Двигаемся сначала по столбцу, потом по строке
  ' 23.12.2020 Do While (ColumnCount < In_maxColumnInSheet) And (Найдено_значение = False)
  Do While (ColumnCount <= In_maxColumnInSheet) And (Найдено_значение = False)
  
    ' По срокам
    rowCount = 1
    
    ' 23.12.2020 Do While (rowCount < In_maxRowInSheet) And (Найдено_значение = False)
    Do While (rowCount <= In_maxRowInSheet) And (Найдено_значение = False)
    
      If Trim(Workbooks(In_Workbooks).Sheets(In_Sheets).Cells(rowCount, ColumnCount).Value) = Trim(In_Value) Then
        rowByValue2 = rowCount
        Число_найденных = Число_найденных + 1
        If Число_найденных = In_Count Then
          Найдено_значение = True
        End If
      End If
      ' Следующая строка
      rowCount = rowCount + 1
    Loop
    ' Следующий столбец
    ColumnCount = ColumnCount + 1
  Loop
End Function


' 22. Определение колонки ячейки в которой находится заданное значение (текст)
Function ColumnByValue(In_Workbooks, In_Sheets, In_Value, In_maxRowInSheet, In_maxColumnInSheet) As Integer
  ColumnByValue = 0
  ColumnCount = 1
  Найдено_значение = False
  ' Двигаемся сначала по столбцу, потом по строке
  Do While (ColumnCount <= In_maxColumnInSheet) And (Найдено_значение = False)
    ' По срокам
    rowCount = 1
    Do While (rowCount <= In_maxRowInSheet) And (Найдено_значение = False)
      ' If Trim(Workbooks(In_Workbooks).Sheets(In_Sheets).Cells(rowCount, ColumnCount).Value) = Trim(In_Value) Then
      If Trim(CStr(Workbooks(In_Workbooks).Sheets(In_Sheets).Cells(rowCount, ColumnCount).Value)) = Trim(CStr(In_Value)) Then
        ColumnByValue = ColumnCount
        Найдено_значение = True
      End If
      ' Следующая строка
      rowCount = rowCount + 1
    Loop
    ' Следующий столбец
    ColumnCount = ColumnCount + 1
  Loop
End Function

' 22.1 Определение колонки ячейки в которой находится заданное значение (текст) и на листе может их быть несколько In_Count = 1-ый, 2-ой и т.д.
Function ColumnByValue2(In_Workbooks, In_Sheets, In_Value, In_maxRowInSheet, In_maxColumnInSheet, In_Count) As Integer
  ColumnByValue2 = 0
  ColumnCount = 1
  Найдено_значение = False
  Число_найденных = 0
  ' Двигаемся сначала по столбцу, потом по строке
  Do While (ColumnCount <= In_maxColumnInSheet) And (Найдено_значение = False)
    ' По срокам
    rowCount = 1
    Do While (rowCount <= In_maxRowInSheet) And (Найдено_значение = False)
      
      ' If Trim(Workbooks(In_Workbooks).Sheets(In_Sheets).Cells(rowCount, ColumnCount).Value) = Trim(In_Value) Then
      If Trim(CStr(Workbooks(In_Workbooks).Sheets(In_Sheets).Cells(rowCount, ColumnCount).Value)) = Trim(CStr(In_Value)) Then
        ColumnByValue2 = ColumnCount
        Число_найденных = Число_найденных + 1
        
        ' Проверяем по счету чсло которые нашли на Листе
        If Число_найденных = In_Count Then
          Найдено_значение = True
        End If
        
      End If
      
      ' Следующая строка
      rowCount = rowCount + 1
    Loop
    ' Следующий столбец
    ColumnCount = ColumnCount + 1
  Loop
End Function

' 22.2 Определение колонки ячейки в которой находится заданное значение (текст) - входящие и исходящие пробелы не удаляем!
Function ColumnByValue3(In_Workbooks, In_Sheets, In_Value, In_maxRowInSheet, In_maxColumnInSheet) As Integer
  ColumnByValue3 = 0
  ColumnCount = 1
  Найдено_значение = False
  ' Двигаемся сначала по столбцу, потом по строке
  Do While (ColumnCount <= In_maxColumnInSheet) And (Найдено_значение = False)
    ' По срокам
    rowCount = 1
    Do While (rowCount <= In_maxRowInSheet) And (Найдено_значение = False)
      ' If Trim(CStr(Workbooks(In_Workbooks).Sheets(In_Sheets).Cells(rowCount, ColumnCount).Value)) = Trim(CStr(In_Value)) Then
      If (CStr(Workbooks(In_Workbooks).Sheets(In_Sheets).Cells(rowCount, ColumnCount).Value)) = (CStr(In_Value)) Then
        ColumnByValue3 = ColumnCount
        Найдено_значение = True
      End If
      ' Следующая строка
      rowCount = rowCount + 1
    Loop
    ' Следующий столбец
    ColumnCount = ColumnCount + 1
  Loop
End Function


' 23. Получение Буквы по номеру столбца (взято из Интернет)
Function ConvertToLetter(iCol) As String
   Dim iAlpha As Integer
   Dim iRemainder As Integer
   iAlpha = Int(iCol / 27)
   iRemainder = iCol - (iAlpha * 26)
   If iAlpha > 0 Then
      ConvertToLetter = Chr(iAlpha + 64)
   End If
   If iRemainder > 0 Then
      ConvertToLetter = ConvertToLetter & Chr(iRemainder + 64)
   End If
End Function

' 24. Определение ячейки в которой находится заданное значение (текст), например G11
Function RangeByValue(In_Workbooks, In_Sheets, In_Value, In_maxRowInSheet, In_maxColumnInSheet) As String
  RangeByValue = ""
  ColumnCount = 1
  Найдено_значение = False
  ' Двигаемся сначала по столбцу, потом по строке
  Do While (ColumnCount < In_maxColumnInSheet) And (Найдено_значение = False)
    ' По срокам
    rowCount = 1
    Do While (rowCount < In_maxRowInSheet) And (Найдено_значение = False)
      
      ' If Trim(Workbooks(In_Workbooks).Sheets(In_Sheets).Cells(RowCount, ColumnCount).Value) = Trim(In_Value) Then
      If Trim(CStr(Workbooks(In_Workbooks).Sheets(In_Sheets).Cells(rowCount, ColumnCount).Value)) = Trim(In_Value) Then
        RangeByValue = ConvertToLetter(ColumnCount) + CStr(rowCount)
        Найдено_значение = True
      End If
      
      ' Следующая строка
      rowCount = rowCount + 1
    Loop
    ' Следующий столбец
    ColumnCount = ColumnCount + 1
  Loop
End Function

' 24.1 Определение ячейки в которой находится заданное значение (текст), например G11. Возвращает переменную "row-column"
Function cellByValue(In_Workbooks, In_Sheets, In_Value, In_maxRowInSheet, In_maxColumnInSheet) As String
  cellByValue = ""
  ColumnCount = 1
  Найдено_значение = False
  ' Двигаемся сначала по столбцу, потом по строке
  Do While (ColumnCount < In_maxColumnInSheet) And (Найдено_значение = False)
    ' По срокам
    rowCount = 1
    Do While (rowCount < In_maxRowInSheet) And (Найдено_значение = False)
      
      ' If Trim(Workbooks(In_Workbooks).Sheets(In_Sheets).Cells(RowCount, ColumnCount).Value) = Trim(In_Value) Then
      If Trim(CStr(Workbooks(In_Workbooks).Sheets(In_Sheets).Cells(rowCount, ColumnCount).Value)) = Trim(In_Value) Then
        cellByValue = CStr(rowCount) + "-" + CStr(ColumnCount)
        Найдено_значение = True
      End If
      
      ' Следующая строка
      rowCount = rowCount + 1
    Loop
    ' Следующий столбец
    ColumnCount = ColumnCount + 1
  Loop
End Function

' 24.2 Из переменной "row-column" получаем row
Function row_cellByValue(In_row_column) As Integer
  row_cellByValue = CInt(Mid(In_row_column, 1, InStr(In_row_column, "-") - 1))
End Function

' 24.3 Из переменной "row-column" получаем column
Function column_cellByValue(In_row_column) As Integer
  column_cellByValue = CInt(Mid(In_row_column, InStr(In_row_column, "-") + 1, Len(In_row_column) - InStr(In_row_column, "-")))
End Function


' 25. Расчет плана на неделю
Function ПланНаНеделю(In_ПланМесяца, In_ФактМесяца, In_DateNow, In_WorkingDayInWeek)
Dim dateBeginWeek, dateEndWeek As Date
  ' Дата начала недели
  dateBeginWeek = weekStartDate(In_DateNow)
  ' Дата конца недели
  dateEndWeek = weekEndDate(In_DateNow)
  ' Число рабочих дней в месяце
  ' workingDaysMonth = Working_days_in_the_FullMonth(максимальная_Дата_выдачи_кредита, 6)
  ' Делаем расчет
  ПланНаНеделю = (In_ПланМесяца - In_ФактМесяца) / Working_days_between_dates(dateBeginWeek, Date_last_day_month(In_DateNow), In_WorkingDayInWeek) * Working_days_between_dates(dateBeginWeek, dateEndWeek, In_WorkingDayInWeek)
  ' Если план на неделю меньше нуля, то выполнять план на этой неделе не надо больше
  If ПланНаНеделю < 0 Then
    ПланНаНеделю = 0
  End If
End Function

' 25.1 Расчет плана на неделю из плана квартала
Function ПланНаНеделю_Q(In_ПланКвартал, In_ФактКвартал, In_DateNow, In_WorkingDayInWeek)
Dim dateBeginWeek, dateEndWeek As Date
  
  ' Дата начала недели
  dateBeginWeek = weekStartDate(In_DateNow)
  ' Дата конца недели
  dateEndWeek = weekEndDate(In_DateNow)
  
  ' Число рабочих дней в месяце
  ' workingDaysMonth = Working_days_in_the_FullMonth(максимальная_Дата_выдачи_кредита, 6)
  
  ' Делаем расчет
  ПланНаНеделю_Q = (In_ПланКвартал - In_ФактКвартал) / Working_days_between_dates(dateBeginWeek, Date_last_day_quarter(In_DateNow), In_WorkingDayInWeek) * Working_days_between_dates(dateBeginWeek, dateEndWeek, In_WorkingDayInWeek)
  
  ' ***
  t0 = Working_days_between_dates(dateBeginWeek, Date_last_day_quarter(In_DateNow), In_WorkingDayInWeek)
  t01 = Date_last_day_quarter(In_DateNow)
  t02 = Working_days_between_dates(dateBeginWeek, dateEndWeek, In_WorkingDayInWeek)
  ' ***
  
  ' Если план на неделю меньше нуля, то выполнять план на этой неделе не надо больше
  If ПланНаНеделю_Q < 0 Then
    ПланНаНеделю_Q = 0
  End If
  
End Function


' 26. Преобразование подстроки Ватрушкина Ираида Семеновна (НК: 12345678) => Ватрушкина И.С. (НК: 12345678)
Function ПреобразованиеФИОиНК(In_ФИО_и_НК) As String
  ПреобразованиеФИОиНК = ""
  Позиция_первой_скобки = InStr(In_ФИО_и_НК, "(")
  Подстрока_1 = Mid(In_ФИО_и_НК, 1, Позиция_первой_скобки)
  Подстрока_2 = Mid(In_ФИО_и_НК, Позиция_первой_скобки, Len(In_ФИО_и_НК) - Позиция_первой_скобки + 1)
  ПреобразованиеФИОиНК = Фамилия_и_Имя(Подстрока_1, 3) + " " + Подстрока_2
End Function

' 26.1 Преобразование подстроки Ватрушкина Ираида Семеновна (НК: 12345678) => 12345678
Function ПреобразованиеФИОиНК2(In_ФИО_и_НК) As String
Dim Позиция_первой_скобки, Позиция_второй_скобки As Byte
  ПреобразованиеФИОиНК2 = ""
  Позиция_первой_скобки = InStr(In_ФИО_и_НК, "(")
  Позиция_второй_скобки = InStr(In_ФИО_и_НК, ")")
  ПреобразованиеФИОиНК2 = Trim(Mid(In_ФИО_и_НК, Позиция_первой_скобки + 4, Позиция_второй_скобки - Позиция_первой_скобки - 4))
End Function

' 26.2 Преобразование подстроки Ватрушкина Ираида Семеновна (НК: 12345678) => Ватрушкина
Function ПреобразованиеФИОиНК3(In_ФИО_и_НК) As String

  ПреобразованиеФИОиНК3 = Trim(Mid(In_ФИО_и_НК, 1, InStr(In_ФИО_и_НК, " ") - 1))
  
End Function

' 26.3 Преобразование подстроки Ватрушкина Ираида Семеновна (НК: 12345678) => Ираида
Function ПреобразованиеФИОиНК4(In_ФИО_и_НК) As String

  Позиция_первго_пробела = InStr(In_ФИО_и_НК, " ")
  Имя_Отчество = Mid(In_ФИО_и_НК, Позиция_первго_пробела + 1, Len(In_ФИО_и_НК) - Позиция_первго_пробела)
  ПреобразованиеФИОиНК4 = Trim(Mid(Имя_Отчество, 1, InStr(Имя_Отчество, " ")))
  
End Function


' 26.4 Преобразование подстроки Ватрушкина Ираида Семеновна (НК: 12345678) => Семеновна
Function ПреобразованиеФИОиНК5(In_ФИО_и_НК) As String
Dim Позиция_первой_скобки, Позиция_второй_скобки As Byte
  
  Позиция_первго_пробела = InStr(In_ФИО_и_НК, " ")
  Имя_Отчество = Mid(In_ФИО_и_НК, Позиция_первго_пробела + 1, Len(In_ФИО_и_НК) - Позиция_первго_пробела)
  Позиция_второго_пробела = Позиция_первго_пробела + InStr(Имя_Отчество, " ")
  Позиция_первой_скобки = InStr(In_ФИО_и_НК, "(")
  Позиция_второй_скобки = InStr(In_ФИО_и_НК, ")")
  ПреобразованиеФИОиНК5 = Trim(Mid(In_ФИО_и_НК, Позиция_второго_пробела + 1, Позиция_первой_скобки - Позиция_второго_пробела - 1))
  
End Function



' 27. Запись К_пор на лист ЕСУП
Sub setК_порInЕСУП(In_Workbooks, In_Sheets, In_К_пор, In_К_порValue, In_К_порDate, In_Size, In_К_порDetailing)
Dim RangeК_пор, Id_TaskVar, PersonVar As String
Dim RangeК_пор_Row, RangeК_пор_Column, OfficeNumberVar As Byte
  
  ' На странице ЕСУП заносим в первую таблицу "Поручения недели:"
  ' ?
  
  ' Находим ячейку (например G41), в которой записано значение ЗДК1
  RangeК_пор = RangeByValue(In_Workbooks, In_Sheets, In_К_пор, 100, 100)
  RangeК_пор_Row = Workbooks(In_Workbooks).Sheets(In_Sheets).Range(RangeК_пор).Row
  RangeК_пор_Column = Workbooks(In_Workbooks).Sheets(In_Sheets).Range(RangeК_пор).Column
  If In_К_порDetailing <> "" Then
      Workbooks(In_Workbooks).Sheets(In_Sheets).Cells(RangeК_пор_Row, RangeК_пор_Column - 5).Value = In_К_порDetailing
  End If
  Workbooks(In_Workbooks).Sheets(In_Sheets).Cells(RangeК_пор_Row, RangeК_пор_Column + 1).Value = In_К_порDate
  Workbooks(In_Workbooks).Sheets(In_Sheets).Cells(RangeК_пор_Row, RangeК_пор_Column + 2).Value = In_К_порValue
  Workbooks(In_Workbooks).Sheets(In_Sheets).Cells(RangeК_пор_Row, RangeК_пор_Column + 3).Value = In_Size
  
  ' Записываем поручение в BASE\Tasks, где поле Id_Task = Неделя-год-ЗДКi (10-2020-ЗДК1) - новая запись
  ' Внимание! Должна быть открыта таблица BASE\Tasks
  ' через InsertRecordInBook
  Id_TaskVar = CStr(WeekNumber(In_К_порDate)) + "-" + CStr(Year(In_К_порDate)) + "-" + In_К_пор
  ' Номер офиса (1..5)
  OfficeNumberVar = CInt(Mid(In_К_пор, Len(In_К_пор), 1))
  
  ' Ответственный за поручение - если это Офисы с 1 по 5
  If OfficeNumberVar = 1 Then
    PersonVar = getFromAddrBook("НОРПиКО1", 3)
  Else
    PersonVar = getFromAddrBook("УДО" + CStr(OfficeNumberVar), 3)
  End If
  ' Ответственный за поручение - если это ИЦ: ВИК1, ВИК2
  If Mid(In_К_пор, 1, 3) = "ВИК" Then
    PersonVar = getFromAddrBook("РИЦ", 3)
  End If
  
  ' Вызов InsertRecordInBook
  Call InsertRecordInBook("Tasks", "Лист1", "Id_Task", Id_TaskVar, _
                                            "Date", In_К_порDate, _
                                              "Protocol", CStr(WeekNumber(In_К_порDate)) + "-" + strDDMMYYYY(In_К_порDate), _
                                                "Id_Task", Id_TaskVar, _
                                                  "Division", getNameOfficeByNumber(OfficeNumberVar), _
                                                    "Person", PersonVar, _
                                                      "Description_task", Workbooks(In_Workbooks).Sheets(In_Sheets).Cells(RangeК_пор_Row, RangeК_пор_Column - 5).Value, _
                                                        "К_пор", In_К_пор, _
                                                          "Value", In_К_порValue, _
                                                            "Unit", In_Size, _
                                                              "Date_finish", weekEndDate(In_К_порDate), _
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

  
End Sub

' 28. Запись Факт исполнения К_пор на лист ЕСУП
Sub currentК_порInЕСУП(In_Workbooks, In_Sheets, In_К_пор, In_К_порDate, In_К_порValue, In_Size)
Dim RangeК_пор As String
Dim RangeК_пор_Row, RangeК_пор_Column As Byte
  
  ' Находим ячейку (например G41), в которой записано значение ЗДК1
  RangeК_пор = RangeByValue(In_Workbooks, In_Sheets, In_К_пор, 100, 100)
  RangeК_пор_Row = Workbooks(In_Workbooks).Sheets(In_Sheets).Range(RangeК_пор).Row
  RangeК_пор_Column = Workbooks(In_Workbooks).Sheets(In_Sheets).Range(RangeК_пор).Column
  
  Workbooks(In_Workbooks).Sheets(In_Sheets).Cells(RangeК_пор_Row, RangeК_пор_Column + 4).Value = In_К_порDate
  Workbooks(In_Workbooks).Sheets(In_Sheets).Cells(RangeК_пор_Row, RangeК_пор_Column + 5).Value = In_К_порValue
  Workbooks(In_Workbooks).Sheets(In_Sheets).Cells(RangeК_пор_Row, RangeК_пор_Column + 6).Value = In_Size
  
  ' ПроцентВыполнения
  ' t = Round(ПроцентВыполнения(Workbooks(In_Workbooks).Sheets(In_Sheets).Cells(RangeК_пор_Row, RangeК_пор_Column + 2).Value, In_К_порValue), 1)
  ' t2 = Round(ПроцентВыполнения(Workbooks(In_Workbooks).Sheets(In_Sheets).Cells(RangeК_пор_Row, RangeК_пор_Column + 2).Value, In_К_порValue), 2)
  ' t3 = ПроцентВыполнения(Workbooks(In_Workbooks).Sheets(In_Sheets).Cells(RangeК_пор_Row, RangeК_пор_Column + 2).Value, In_К_порValue)
  
  Workbooks(In_Workbooks).Sheets(In_Sheets).Cells(RangeК_пор_Row, RangeК_пор_Column + 7).Value = Round(ПроцентВыполнения(Workbooks(In_Workbooks).Sheets(In_Sheets).Cells(RangeК_пор_Row, RangeК_пор_Column + 2).Value, In_К_порValue), 2)
  
  ' Апдейт значений для поручения в BASE\Tasks, где поле Id_Task = Неделя-год-ЗДКi (10-2020-ЗДК1): Last_Date, Last_Value, Status_persent, Status
  ' Внимание! Должна быть открыта таблица BASE\Tasks
  ' через InsertRecordInBook апдейтим по ключу Id_Task
  Id_TaskVar = CStr(WeekNumber(In_К_порDate)) + "-" + CStr(Year(In_К_порDate)) + "-" + In_К_пор
  Call InsertRecordInBook("Tasks", "Лист1", "Id_Task", Id_TaskVar, _
                                            "Last_Date", In_К_порDate, _
                                              "Last_Value", In_К_порValue, _
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
                                                                                "", "", _
                                                                                  "", "")
  
End Sub

' 29. Прирост в % между двумя числами
Function Прирост_в_процентах(In_Start, In_End) As Double
  ' Прирост в процентах составляет % = (B-A)/A*100 A = Исходное значение B = Конечное значение
  If In_Start <> 0 Then
    Прирост_в_процентах = ((In_End - In_Start) / In_Start * 100) / 100
  Else
    Прирост_в_процентах = 1
  End If
End Function

' 30. Процент выполнения от Плана и Факта (https://calc.by/math-calculators/percent-calculator.html)
Function ПроцентВыполнения(In_План, In_Факт) As String
  If In_План > 0 Then
    ПроцентВыполнения = (In_Факт / In_План)
  Else
    ПроцентВыполнения = 1
  End If
End Function

' 31. Неделя на ЛистN
Function НеделяНаЛистеN(In_Sheet) As Integer
Dim Range_str As String
Dim Range_Row, Range_Column As Byte
  ' Находим ячейку в которой на листе ЕСУП записана "Неделя:"
  Range_str = RangeByValue(ThisWorkbook.Name, In_Sheet, "Неделя:", 100, 100)
  Range_Row = Workbooks(ThisWorkbook.Name).Sheets(In_Sheet).Range(Range_str).Row
  Range_Column = Workbooks(ThisWorkbook.Name).Sheets(In_Sheet).Range(Range_str).Column
  '
  НеделяНаЛистеN = 0
  НеделяНаЛистеN = CInt(ThisWorkbook.Sheets(In_Sheet).Cells(Range_Row, Range_Column + 1).Value)
End Function

' 32. Открываем ini-файл и считываем значение переменной
Function param_from_ini(In_iniFile, In_ParamName) As String
   param_from_ini = ""
   f = FreeFile
   Open In_iniFile For Input As #f
   Do While Not EOF(f)
     Line Input #f, S
     ' Если в строке есть знак "="
     If InStr(S, "=") <> 0 Then
       ' Определяем позицию знака "="
       Позиция_равно = InStr(S, "=")
       ' Сравниваем строки
       If Trim(Mid(S, 1, Позиция_равно - 1)) = In_ParamName Then
         ' Позиция_равно = InStr(s, "=")
         param_from_ini = Trim(Mid(S, Позиция_равно + 1, Len(S) - Позиция_равно))
       End If
     End If ' Если в строке есть знак "="
   Loop
   Close f
End Function

' 33. Отправка письма (тест)
Sub Отправка_Lotus_Notes()
Dim текстПисьма As String
  ' Текст письма
  текстПисьма = "Уважаемые сотрудники," & Chr(13)
  текстПисьма = текстПисьма + "" + Chr(13)
  текстПисьма = текстПисьма + "Объем кредитного портфеля с изменениями за период." + Chr(13)
  текстПисьма = текстПисьма + "" + Chr(13)
  текстПисьма = текстПисьма + "" + Chr(13)
  ' Визитка (подпись С Ув., )
  текстПисьма = текстПисьма + ПодписьВПисьме()
  ' Хэштег
  текстПисьма = текстПисьма + ThisWorkbook.Sheets("Лист3").Cells(2, 13).Value + Chr(13)
  ' Вызов
  Call send_Lotus_Notes(ThisWorkbook.Sheets("Лист3").Cells(2, 16).Value, "Sergey Fedorovich Proschaev/Tyumen/PSBank/Ru", "Sergey Fedorovich Proschaev/Tyumen/PSBank/Ru", текстПисьма, "C:\Users\proschaevsf\Documents\#DB_Result\Out\Тестовый_файл_для_отправки.xlsx")
End Sub

' 34. Проверка формата открываемых отчетов
' Вынесен в отдельный модуль "Module_CheckFormatReport"
    
' 35. Проверка наличия листа с заданным именем в Книге (моя версия)
Function Sheets_Exist2(In_Workbooks, In_Sheets) As Boolean
Dim wsSh As Worksheet
  On Error Resume Next
  Set wsSh = Workbooks(In_Workbooks).Sheets(In_Sheets)
  Sheets_Exist2 = Not wsSh Is Nothing
End Function

' 36. Список присутствующих на собрании в виде строки с Листа ЕСУП, In_Status = 1 - присутствующие, In_Status = 0 - отсутствующие
Function Присутствовавшие_на_Собрании(In_Status) As String

Dim Присутств_на_Собрании_Range, Пригл_на_Собрание_Range, Должность_и_ФИО, список_приглашенных_str As String
Dim rowCount, Присутств_на_Собрании_Row, Присутств_на_Собрании_Column, Число_присуствующих, Число_приглашенных, Пригл_на_Собрание_Row, Пригл_на_Собрание_Column As Byte
  
  ' Инициализация строки
  Присутствовавшие_на_Собрании = ""
  ' На Листе "ЕСУП" находим ячейку с данными "Присутств_на_Собрании"
  Присутств_на_Собрании_Range = RangeByValue(ThisWorkbook.Name, "ЕСУП", "Присутств_на_Собрании", 100, 100)
  Присутств_на_Собрании_Row = ThisWorkbook.Sheets("ЕСУП").Range(Присутств_на_Собрании_Range).Row
  Присутств_на_Собрании_Column = ThisWorkbook.Sheets("ЕСУП").Range(Присутств_на_Собрании_Range).Column
  
  ' Обрабатываем список
  Число_присуствующих = 0
  rowCount = Присутств_на_Собрании_Row + 1
  Do While ThisWorkbook.Sheets("ЕСУП").Cells(rowCount, Присутств_на_Собрании_Column).Value <> "Пригл_на_Собрание"
    
    ' Проверяем статус 1 - присутствует, 0 - отсутствует
    If ThisWorkbook.Sheets("ЕСУП").Cells(rowCount, Присутств_на_Собрании_Column).Value = In_Status Then
       
      ' ФИО <> 0
      If ThisWorkbook.Sheets("ЕСУП").Cells(rowCount, Присутств_на_Собрании_Column + 1).Value <> 0 Then
      
        Должность_и_ФИО = ThisWorkbook.Sheets("ЕСУП").Cells(rowCount, Присутств_на_Собрании_Column + 4).Value + " " + Фамилия_и_Имя(ThisWorkbook.Sheets("ЕСУП").Cells(rowCount, Присутств_на_Собрании_Column + 1).Value, 3)
      
        If Число_присуствующих <> 0 Then
          Присутствовавшие_на_Собрании = Присутствовавшие_на_Собрании + ", " + Должность_и_ФИО
        Else
          Присутствовавшие_на_Собрании = Должность_и_ФИО
        End If
        Число_присуствующих = Число_присуствующих + 1
        
      End If
      
    End If
    ' Следующая запись
    rowCount = rowCount + 1
  Loop
  
  ' Если генерируем список присутствующих, то добавляем и Приглашенных (если есть)
  If In_Status = 1 Then
    
    ' Список приглашенных участников:
    список_приглашенных_str = ""
    Число_приглашенных = 0
    
    Пригл_на_Собрание_Range = RangeByValue(ThisWorkbook.Name, "ЕСУП", "Пригл_на_Собрание", 100, 100)
    Пригл_на_Собрание_Row = ThisWorkbook.Sheets("ЕСУП").Range(Пригл_на_Собрание_Range).Row
    Пригл_на_Собрание_Column = ThisWorkbook.Sheets("ЕСУП").Range(Пригл_на_Собрание_Range).Column
  
    rowCount = Пригл_на_Собрание_Row + 1
  
    Do While ThisWorkbook.Sheets("ЕСУП").Cells(rowCount, Пригл_на_Собрание_Column).Value <> ""
      
      ' Проверяем статус 1 - присутствует, 0 - отсутствует
      If ThisWorkbook.Sheets("ЕСУП").Cells(rowCount, Пригл_на_Собрание_Column).Value = 1 Then
        
        Должность_и_ФИО = ThisWorkbook.Sheets("ЕСУП").Cells(rowCount, Пригл_на_Собрание_Column + 4).Value + " " + Фамилия_и_Имя(ThisWorkbook.Sheets("ЕСУП").Cells(rowCount, Пригл_на_Собрание_Column + 1).Value, 3)
        
        If Число_приглашенных <> 0 Then
          список_приглашенных_str = список_приглашенных_str + ", " + Должность_и_ФИО
        Else
          список_приглашенных_str = Должность_и_ФИО
        End If
        Число_приглашенных = Число_приглашенных + 1
      End If
      
      ' Следующая запись
      rowCount = rowCount + 1
    Loop
  
    ' Если список приглашенных не пустой - то добавляем к списку участников:
    If список_приглашенных_str <> "" Then
      Присутствовавшие_на_Собрании = Присутствовавшие_на_Собрании + ". Приглашенные участники: " + список_приглашенных_str
    End If
  End If
  
  ' Если список пустой - возвращаем "-"
  If Присутствовавшие_на_Собрании = "" Then
    Присутствовавшие_на_Собрании = "-"
  End If
End Function

' 37. Фамилия и Имя из ФИО (Параметр: 1-Иванов Иван, 2-Иванов И., 3-Иванов И.И., 4-Иванов И (без точки), 5-Иванов )
Function Фамилия_и_Имя(In_Фамилия_Имя_Отчество, In_Type As Byte) As String
Dim Первый_пробел, Второй_пробел As Byte
  
  If In_Фамилия_Имя_Отчество <> "" Then
    
    ' Исполняемый код Функции:
    Фамилия_и_Имя = "_<ФИО>_"
    
    ' Первый пробел
    Первый_пробел = InStr(In_Фамилия_Имя_Отчество, " ")
  
    ' Здесь ищем с позиции, следующей за первым пробелом
    Второй_пробел = InStr(Первый_пробел + 1, In_Фамилия_Имя_Отчество, " ")
    
    ' Если In_Type=1 выводим Иванов Иван
    If In_Type = 1 Then
       Фамилия_и_Имя = Mid(In_Фамилия_Имя_Отчество, 1, Первый_пробел - 1) + " " + Mid(In_Фамилия_Имя_Отчество, Первый_пробел + 1, Второй_пробел - Первый_пробел - 1)
    End If
    ' Если In_Type=2 выводим Иванов И.
    If In_Type = 2 Then
       Фамилия_и_Имя = Mid(In_Фамилия_Имя_Отчество, 1, Первый_пробел - 1) + " " + Mid(In_Фамилия_Имя_Отчество, Первый_пробел + 1, 1) + "."
    End If
    ' Если In_Type=3 выводим Иванов И.И.
    If In_Type = 3 Then
       Фамилия_и_Имя = Mid(In_Фамилия_Имя_Отчество, 1, Первый_пробел - 1) + " " + Mid(In_Фамилия_Имя_Отчество, Первый_пробел + 1, 1) + "." + Mid(In_Фамилия_Имя_Отчество, Второй_пробел + 1, 1) + "."
    End If
    ' Если In_Type=4 выводим Иванов И (без точки)
    If In_Type = 4 Then
       Фамилия_и_Имя = Mid(In_Фамилия_Имя_Отчество, 1, Первый_пробел - 1) + " " + Mid(In_Фамилия_Имя_Отчество, Первый_пробел + 1, 1) ' + "."
    End If
    ' Если In_Type=5 выводим только фамилию Иванов
    If In_Type = 5 Then
       Фамилия_и_Имя = Mid(In_Фамилия_Имя_Отчество, 1, Первый_пробел - 1) ' + " " + Mid(In_Фамилия_Имя_Отчество, Первый_пробел + 1, 1) ' + "."
    End If
  Else
    Фамилия_и_Имя = ""
  End If
  
End Function

' 38. Расчет высоты строки по длинне текста
Function lineHeight(In_Str, In_lineHeight, In_numberOfCharacters)
Dim Расчетная_высота As Integer
  
  ' 1-ый вариант: Целочисленное деление отличается от деления с плавающей точкой тем, что его результатом всегда есть целое число без дробной части. VBA отбрасывает (но не округляет!) любой дробный остаток результата выражения целочисленного деления. Например, выражения 22\5 и 24\5 будут иметь один и тот же результат = 4.
  ' Расчетная_высота = (Len(Trim(In_Str)) \ In_numberOfCharacters) * In_lineHeight
  
  ' 2-ой вариант:
  Расчетная_высота = (Len(Trim(In_Str)) \ In_numberOfCharacters) * In_lineHeight
  
  If Расчетная_высота < In_lineHeight Then
    lineHeight = In_lineHeight
  Else
    ' lineHeight = Расчетная_высота
    lineHeight = Расчетная_высота + In_lineHeight
  End If
  
End Function

' 39. Получение данных из адрессной книги
Function getFromAddrBook(In_К_дол, In_TypeData) As String
Dim К_дол_Range As String
Dim К_дол_Row, К_дол_Column, rowCount As Byte
  
  getFromAddrBook = ""
  
  ' Находим In_К_дол на Листе
  ' К_дол_Range = RangeByValue(ThisWorkbook.Name, "Addr.Book", In_К_дол, 100, 100)
  
  ' If К_дол_Range <> "" Then
    
    ' In_TypeData = 1: Должность (сокр) + ФИО (сокр)
    If In_TypeData = 1 Then
      ' Выполняем поиск In_К_дол
      К_дол_Range = RangeByValue(ThisWorkbook.Name, "Addr.Book", In_К_дол, 100, 100)
      К_дол_Row = ThisWorkbook.Sheets("Addr.Book").Range(К_дол_Range).Row
      К_дол_Column = ThisWorkbook.Sheets("Addr.Book").Range(К_дол_Range).Column
      getFromAddrBook = ThisWorkbook.Sheets("Addr.Book").Cells(К_дол_Row, К_дол_Column + 6).Value + " " + Фамилия_и_Имя(ThisWorkbook.Sheets("Addr.Book").Cells(К_дол_Row, К_дол_Column - 1).Value, 3)
    End If
  
    ' In_TypeData = 2: Вся розница для рассылки, которая есть в подстроке УДО2,УДО3,УДО4,УДО5,НОРПиКО1,НОРПиКО2,НОРПиКО3,НОРПиКО4,НОРПиКО5,ПМ,МРК1,МРК2,МРК3,МРК4,МРК5,НОКП,РРКК,МПП,РРИЦ,РИЦ,СотрИЦ
    If In_TypeData = 2 Then
      ' К_долStr1 = "УДО2,УДО3,УДО4,УДО5,НОРПиКО1,НОРПиКО2,НОРПиКО3,НОРПиКО4,НОРПиКО5,ПМ,МРК1,МРК2,МРК3,МРК4,МРК5,НОКП,РРКК,МПП,РРИЦ,РИЦ,СотрИЦ"
      rowCount = 7
      Do While ThisWorkbook.Sheets("Addr.Book").Cells(rowCount, 3).Value <> ""
        If InStr(In_К_дол, ThisWorkbook.Sheets("Addr.Book").Cells(rowCount, 3).Value) <> 0 Then
          If getFromAddrBook = "" Then
            getFromAddrBook = ThisWorkbook.Sheets("Addr.Book").Cells(rowCount, 10).Value
          Else
            getFromAddrBook = getFromAddrBook + ", " + ThisWorkbook.Sheets("Addr.Book").Cells(rowCount, 10).Value
          End If
        End If
        ' Следующая запись
        rowCount = rowCount + 1
      Loop
    End If
  
    ' In_TypeData = 3: ФИО (сокр)
    If In_TypeData = 3 Then
      ' Выполняем поиск In_К_дол
      К_дол_Range = RangeByValue(ThisWorkbook.Name, "Addr.Book", In_К_дол, 100, 100)
      If К_дол_Range <> "" Then
        К_дол_Row = ThisWorkbook.Sheets("Addr.Book").Range(К_дол_Range).Row
        К_дол_Column = ThisWorkbook.Sheets("Addr.Book").Range(К_дол_Range).Column
        getFromAddrBook = Фамилия_и_Имя(ThisWorkbook.Sheets("Addr.Book").Cells(К_дол_Row, К_дол_Column - 1).Value, 3)
      End If
    End If

    ' In_TypeData = 4: Фамилия Имя Отчество
    If In_TypeData = 4 Then
      ' Выполняем поиск In_К_дол
      К_дол_Range = RangeByValue(ThisWorkbook.Name, "Addr.Book", In_К_дол, 100, 100)
      If К_дол_Range <> "" Then
        К_дол_Row = ThisWorkbook.Sheets("Addr.Book").Range(К_дол_Range).Row
        К_дол_Column = ThisWorkbook.Sheets("Addr.Book").Range(К_дол_Range).Column
        getFromAddrBook = ThisWorkbook.Sheets("Addr.Book").Cells(К_дол_Row, К_дол_Column - 1).Value
      End If
    End If
    
    ' In_TypeData = 5: адрес электронной почты (10-ый столбец)
    If In_TypeData = 5 Then
      ' Выполняем поиск In_К_дол
      К_дол_Range = RangeByValue(ThisWorkbook.Name, "Addr.Book", In_К_дол, 100, 100)
      If К_дол_Range <> "" Then
        К_дол_Row = ThisWorkbook.Sheets("Addr.Book").Range(К_дол_Range).Row
        К_дол_Column = ThisWorkbook.Sheets("Addr.Book").Range(К_дол_Range).Column
        getFromAddrBook = ThisWorkbook.Sheets("Addr.Book").Cells(К_дол_Row, К_дол_Column + 7).Value
      End If
    End If
    
    ' 6: Имя Обращение - форма 1 (Имя)
    If In_TypeData = 6 Then
      ' Выполняем поиск In_К_дол
      К_дол_Range = RangeByValue(ThisWorkbook.Name, "Addr.Book", In_К_дол, 100, 100)
      If К_дол_Range <> "" Then
        К_дол_Row = ThisWorkbook.Sheets("Addr.Book").Range(К_дол_Range).Row
        К_дол_Column = ThisWorkbook.Sheets("Addr.Book").Range(К_дол_Range).Column
        getFromAddrBook = ThisWorkbook.Sheets("Addr.Book").Cells(К_дол_Row, К_дол_Column + 12).Value
      End If
    End If
 
    ' 7: Имя Отчество Обращение - форма 2 (Имя)
    If In_TypeData = 7 Then
      ' Выполняем поиск In_К_дол
      К_дол_Range = RangeByValue(ThisWorkbook.Name, "Addr.Book", In_К_дол, 100, 100)
      If К_дол_Range <> "" Then
        К_дол_Row = ThisWorkbook.Sheets("Addr.Book").Range(К_дол_Range).Row
        К_дол_Column = ThisWorkbook.Sheets("Addr.Book").Range(К_дол_Range).Column
        getFromAddrBook = ThisWorkbook.Sheets("Addr.Book").Cells(К_дол_Row, К_дол_Column + 13).Value
      End If
    End If
 
  ' End If ' К_дол_Range <> ""

End Function

' 39.1 Получение данных из адрессной книги по табельному номеру
Function getFromAddrBook2(In_Табномер, In_TypeData) As String
Dim К_дол_Range As String
Dim К_дол_Row, К_дол_Column, rowCount As Byte
  
  getFromAddrBook2 = ""
  
  ' In_TypeData = 1: Должность (сокр) + ФИО (сокр)
  If In_TypeData = 1 Then
    ' Выполняем поиск In_Табномер
    К_дол_Range = RangeByValue(ThisWorkbook.Name, "Addr.Book", In_Табномер, 100, 100)
    К_дол_Row = ThisWorkbook.Sheets("Addr.Book").Range(К_дол_Range).Row
    К_дол_Column = ThisWorkbook.Sheets("Addr.Book").Range(К_дол_Range).Column
    getFromAddrBook2 = ThisWorkbook.Sheets("Addr.Book").Cells(К_дол_Row, К_дол_Column + 3).Value + " " + Фамилия_и_Имя(ThisWorkbook.Sheets("Addr.Book").Cells(К_дол_Row, К_дол_Column - 4).Value, 3)
  End If
    
  ' In_TypeData = 2: ФИО (сокр)
  If In_TypeData = 2 Then
    ' Выполняем поиск In_Табномер
    К_дол_Range = RangeByValue(ThisWorkbook.Name, "Addr.Book", In_Табномер, 100, 100)
    К_дол_Row = ThisWorkbook.Sheets("Addr.Book").Range(К_дол_Range).Row
    К_дол_Column = ThisWorkbook.Sheets("Addr.Book").Range(К_дол_Range).Column
    getFromAddrBook2 = Фамилия_и_Имя(ThisWorkbook.Sheets("Addr.Book").Cells(К_дол_Row, К_дол_Column - 4).Value, 3)
  End If

  ' In_TypeData = 3: Должность полная + ФИО (сокр)
  If In_TypeData = 3 Then
    ' Выполняем поиск In_Табномер
    К_дол_Range = RangeByValue(ThisWorkbook.Name, "Addr.Book", In_Табномер, 100, 100)
    If К_дол_Range <> "" Then
      К_дол_Row = ThisWorkbook.Sheets("Addr.Book").Range(К_дол_Range).Row
      К_дол_Column = ThisWorkbook.Sheets("Addr.Book").Range(К_дол_Range).Column
      getFromAddrBook2 = ThisWorkbook.Sheets("Addr.Book").Cells(К_дол_Row, К_дол_Column + 2).Value + " " + Фамилия_и_Имя(ThisWorkbook.Sheets("Addr.Book").Cells(К_дол_Row, К_дол_Column - 4).Value, 3)
    End If
  End If

  ' In_TypeData = 4: Должность (сокр) + ФИО (сокр)
  If In_TypeData = 4 Then
    ' Выполняем поиск In_Табномер
    К_дол_Range = RangeByValue(ThisWorkbook.Name, "Addr.Book", In_Табномер, 100, 100)
    К_дол_Row = ThisWorkbook.Sheets("Addr.Book").Range(К_дол_Range).Row
    К_дол_Column = ThisWorkbook.Sheets("Addr.Book").Range(К_дол_Range).Column
    getFromAddrBook2 = ThisWorkbook.Sheets("Addr.Book").Cells(К_дол_Row, К_дол_Column + 3).Value + " " + Фамилия_и_Имя(ThisWorkbook.Sheets("Addr.Book").Cells(К_дол_Row, К_дол_Column - 4).Value, 3)
  End If

  ' In_TypeData = 5: Должность полная
  If In_TypeData = 5 Then
    ' Выполняем поиск In_Табномер
    К_дол_Range = RangeByValue(ThisWorkbook.Name, "Addr.Book", In_Табномер, 100, 100)
    If К_дол_Range <> "" Then
      К_дол_Row = ThisWorkbook.Sheets("Addr.Book").Range(К_дол_Range).Row
      К_дол_Column = ThisWorkbook.Sheets("Addr.Book").Range(К_дол_Range).Column
      getFromAddrBook2 = ThisWorkbook.Sheets("Addr.Book").Cells(К_дол_Row, К_дол_Column + 2).Value
    End If
  End If


End Function

' 39.2 Получение адреса почты из адрессной книги по полному ФИО Иванов Иван Иванович
Function getFromAddrBook3(In_ФИО) As String
  
  getFromAddrBook3 = "Адрес LotusNotes не определен (поиск по ФИО)!"
  
  ' Находим строку с ФИО в Addr.Book
  row_In_AddrBook = rowByValue(ThisWorkbook.Name, "Addr.Book", In_ФИО, 100, 100)
  
  ' Если запись была найдена
  If row_In_AddrBook <> 0 Then
    getFromAddrBook3 = ThisWorkbook.Sheets("Addr.Book").Cells(row_In_AddrBook, 10).Value
  End If
  
End Function


' 40. Получение наименование офиса по моему порядковому номеру 1 - Тюмень, 2 - Сургут, 3 - Нижневартовск, 4 - Новый Уренгой, 5 - Тарко-Сале
Function getNameOfficeByNumber(In_Number) As String
  getNameOfficeByNumber = "Офис не определен!"
  Select Case In_Number
    Case 1 ' ОО «Тюменский»
      getNameOfficeByNumber = "ОО «Тюменский»"
    Case 2 ' ОО «Сургутский»
      getNameOfficeByNumber = "ОО «Сургутский»"
    Case 3 ' ОО «Нижневартовский»
      getNameOfficeByNumber = "ОО «Нижневартовский»"
    Case 4 ' ОО «Новоуренгойский»
      getNameOfficeByNumber = "ОО «Новоуренгойский»"
    Case 5 ' ОО «Тарко-Сале»
      getNameOfficeByNumber = "ОО «Тарко-Сале»"
    Case 0 ' Итого по РОО «Тюменский»
      getNameOfficeByNumber = "Итого по РОО «Тюменский»"
      
  End Select
End Function

' 40. Получение наименование офиса по моему порядковому номеру 1 - Тюмень, 2 - Сургут, 3 - Нижневартовск, 4 - Новый Уренгой, 5 - Тарко-Сале
Function getNameOfficeByNumber2(In_Number) As String
  getNameOfficeByNumber2 = "Офис не определен!"
  Select Case In_Number
    Case 1 ' ОО «Тюменский»
      getNameOfficeByNumber2 = "Тюменский"
    Case 2 ' ОО «Сургутский»
      getNameOfficeByNumber2 = "Сургутский"
    Case 3 ' ОО «Нижневартовский»
      getNameOfficeByNumber2 = "Нижневартовский"
    Case 4 ' ОО «Новоуренгойский»
      getNameOfficeByNumber2 = "Новоуренгойский"
    Case 5 ' ОО «Тарко-Сале»
      getNameOfficeByNumber2 = "Тарко-Сале"
  End Select
End Function


' 40.1 Получение наименование офиса по городу Тюмень, 2 - Сургут, 3 - Нижневартовск, 4 - Новый Уренгой, 5 - Тарко-Сале
Function getNameOfficeByCity(In_City) As String
  getNameOfficeByCity = "Офис не определен!"
  Select Case In_Number
    Case "Тюмень" ' ОО «Тюменский»
      getNameOfficeByCity = "ОО «Тюменский»"
    Case "Сургут" ' ОО «Сургутский»
      getNameOfficeByCity = "ОО «Сургутский»"
    Case "Нижневартовский" ' ОО «Нижневартовский»
      getNameOfficeByCity = "ОО «Нижневартовский»"
    Case "Новый Уренгой" ' ОО «Новоуренгойский»
      getNameOfficeByCity = "ОО «Новоуренгойский»"
    Case "Тарко-Сале" ' ОО «Тарко-Сале»
      getNameOfficeByCity = "ОО «Тарко-Сале»"
  End Select
End Function

' 40.2 Получение номера офиса по наименованию 1 - Тюменский, 2 - Сургутский, 3 - Нижневартовский, 4 - Новоуренгойский, 5 - Тарко-Сале
Function getNumberOfficeByName(In_Office) As Byte
  getNumberOfficeByName = 0
  Select Case In_Office
    Case "Тюменский"
      getNumberOfficeByName = 1
    Case "Сургутский"
      getNumberOfficeByName = 2
    Case "Нижневартовский"
      getNumberOfficeByName = 3
    Case "Новоуренгойский"
      getNumberOfficeByName = 4
    Case "Тарко-Сале"
      getNumberOfficeByName = 5
  End Select
End Function

' 40.3 Получение номера офиса по наименованию 1 - Тюменский, 2 - Сургутский, 3 - Нижневартовский, 4 - Новоуренгойский, 5 - Тарко-Сале
' Функция getNumberOfficeByName2 обрабатывает ОО2 "Тарко-Сале"
Function getNumberOfficeByName2(In_Office) As Byte
  
  getNumberOfficeByName2 = 0
  
  If InStr(In_Office, "Тюменский") <> 0 Then
    getNumberOfficeByName2 = 1
  End If
      
  If InStr(In_Office, "Сургутский") <> 0 Then
    getNumberOfficeByName2 = 2
  End If
      
  If InStr(In_Office, "Нижневартовский") <> 0 Then
    getNumberOfficeByName2 = 3
  End If
      
  If InStr(In_Office, "Новоуренгойский") <> 0 Then
    getNumberOfficeByName2 = 4
  End If
      
  If InStr(In_Office, "Тарко-Сале") <> 0 Then
    getNumberOfficeByName2 = 5
  End If

End Function

' 40.3 Получение наименование офиса: Тюменский, Сургутский, Нижневартовский, Новоуренгойский, Тарко-Сале
' Функция getShortNameOfficeByName обрабатывает ОО2 "Тарко-Сале"
Function getShortNameOfficeByName(In_Office) As String
  
  getShortNameOfficeByName = 0
  
  If InStr(In_Office, "Тюменский") <> 0 Then
    getShortNameOfficeByName = "Тюменский"
  End If
      
  If InStr(In_Office, "Сургутский") <> 0 Then
    getShortNameOfficeByName = "Сургутский"
  End If
      
  If InStr(In_Office, "Нижневартовский") <> 0 Then
    getShortNameOfficeByName = "Нижневартовский"
  End If
      
  If InStr(In_Office, "Новоуренгойский") <> 0 Then
    getShortNameOfficeByName = "Новоуренгойский"
  End If
      
  If InStr(In_Office, "Тарко-Сале") <> 0 Then
    getShortNameOfficeByName = "Тарко-Сале"
  End If

End Function

' 40.4 Получение наименование офиса: ОО «Тюменский», ОО «Сургутский», ОО «Нижневартовский», ОО «Новоуренгойский», ОО «Тарко-Сале»
Function updateNameOfficeByName(In_Office) As String
  
  updateNameOfficeByName = ""
  
  If InStr(In_Office, "Тюменский") <> 0 Then
    updateNameOfficeByName = "ОО «Тюменский»"
  End If
      
  If InStr(In_Office, "Сургутский") <> 0 Then
    updateNameOfficeByName = "ОО «Сургутский»"
  End If
      
  If InStr(In_Office, "Нижневартовский") <> 0 Then
    updateNameOfficeByName = "ОО «Нижневартовский»"
  End If
      
  If InStr(In_Office, "Новоуренгойский") <> 0 Then
    updateNameOfficeByName = "ОО «Новоуренгойский»"
  End If
      
  If InStr(In_Office, "Тарко-Сале") <> 0 Then
    updateNameOfficeByName = "ОО «Тарко-Сале»"
  End If

End Function


' 41. Дата протокола (строка) из имени файла с протоколом - берем из G2 "10-02032020"
Function dateProtocol(In_ProtocolNumber) As Date
Dim позицияДефис As Byte
Dim str_dateProtocol
  ' Находим подстроку от "-"
  позицияДефис = InStr(In_ProtocolNumber, "-")
  str_dateProtocol = Mid(In_ProtocolNumber, позицияДефис + 1, Len(In_ProtocolNumber) - позицияДефис)
  ' Разделяем дату, месяц и год точкой
  dateProtocol = CDate(Mid(str_dateProtocol, 1, 2) + "." + Mid(str_dateProtocol, 3, 2) + "." + Mid(str_dateProtocol, 5, 4))
End Function

' 42. Создать Хэштэг в Листе0 P13 формат: 08.03.2020 10:06 => 832106
Function createHashTag(In_Letter)
Dim Число, Месяц, Год, Часы, Минуты As String
  ' Число
  Число = Mid(CStr(Date), 1, 2)
  If CInt(Число) < 10 Then
    Число = Replace(Число, "0", "")
  End If
  ' Месяц
  Месяц = Mid(CStr(Date), 4, 2)
  If CInt(Месяц) < 10 Then
    Месяц = Replace(Месяц, "0", "")
  End If
  ' Год 2020 -> "2", 2021 -> "21", 2022 -> "22"
  Год = Replace(Mid(Year(Date), 3, 2), "0", "")
  ' Часы = Replace(Mid(CStr(Time), 1, 2), "0", "")
  Часы = Mid(CStr(Time), 1, 2)
  If CInt(Часы) < 10 Then
    Часы = Replace(Часы, "0", "")
  End If
  ' Минуты
  Минуты = Mid(CStr(Time), 4, 2)
  If CInt(Минуты) < 10 Then
    Минуты = Replace(Минуты, "0", "")
  End If
  ' Генерация Хэштега
  createHashTag = "#" + In_Letter + Число + Месяц + Год + Часы + Минуты
End Function

' 43. Дата следующего воскресенья (исп во вкладах)
Function Next_sunday_date(In_Date) As Date
Dim Текущая_дата_рассчета As Date
  ' Берем переданную дату. Если In_Date - воскресенье, то считаем с понедельника
  If Weekday(In_Date, vbMonday) = 7 Then
    Текущая_дата_рассчета = In_Date + 1
  Else
    Текущая_дата_рассчета = In_Date
  End If
  ' Считаем дни до первого воскресенья
  Do While Weekday(Текущая_дата_рассчета, vbMonday) < 6
    ' Следующий день
    Текущая_дата_рассчета = Текущая_дата_рассчета + 1
  Loop
  Next_sunday_date = Текущая_дата_рассчета + 1
End Function

' 44. Определяем число рабочих дней с текущего дня до конца месяца (используется по вкладам)
Function Working_days_in_the_month(In_DateNow, In_working_days_in_the_week, In_считаем_сегодня) As Integer
  ' Декодируем дату In_DateNow
  ' Месяц
  Месяц = Month(In_DateNow)
  ' Год
  Год = Year(In_DateNow)
  ' Первый день следующего месяца
  If Месяц = 12 Then
    Месяц = 0
    Год = Год + 1
  End If
  Первый_день_следующего_месяца = CDate("01." + CStr(Месяц + 1) + "." + CStr(Год))
  ' Если считаем сегодняшний день
  If In_считаем_сегодня = True Then
    Текущая_дата_рассчета = In_DateNow
  Else
    Текущая_дата_рассчета = In_DateNow + 1
  End If
  ' Делаем рассчет по датам
  Working_days_in_the_month = 0
  Do While Текущая_дата_рассчета < Первый_день_следующего_месяца
    ' Если Текущая_дата_рассчета не суббота
    If In_working_days_in_the_week = 5 Then
      ' Если пятидневка
      If (Weekday(Текущая_дата_рассчета, vbMonday) <> 6) And (Weekday(Текущая_дата_рассчета, vbMonday) <> 7) Then
        Working_days_in_the_month = Working_days_in_the_month + 1
      End If
    Else
      ' Если шестидневка - In_working_days_in_the_week = 6
      If (Weekday(Текущая_дата_рассчета, vbMonday) <> 7) Then
        Working_days_in_the_month = Working_days_in_the_month + 1
      End If
    End If
    ' Следующая дата
    Текущая_дата_рассчета = Текущая_дата_рассчета + 1
  Loop ' Следующая дата
End Function

' 45. Установка % выполнения по задачам в BASE\Tasks
Sub setStatusInTasks(In_BookName, In_Sheet, In_Last_Date, In_К_пор, In_Protocol)

Dim rowCount As Integer
Dim column_К_пор, column_Last_Date, column_Value, column_Last_Value, column_Status_persent, column_Status, column_Description_status, column_Date_finish, column_Protocol As Byte

  ' Номер столбца Last_Date
  column_Last_Date = ColumnByName(In_BookName, In_Sheet, 1, "Last_Date")
  ' Номер столбца К_пор
  column_К_пор = ColumnByName(In_BookName, In_Sheet, 1, "К_пор")
  ' Номер столбца Value
  column_Value = ColumnByName(In_BookName, In_Sheet, 1, "Value")
  ' Номер столбца Last_Value
  column_Last_Value = ColumnByName(In_BookName, In_Sheet, 1, "Last_Value")
  ' Номер столбца Status_persent
  column_Status_persent = ColumnByName(In_BookName, In_Sheet, 1, "Status_persent")
  ' Номер столбца Status
  column_Status = ColumnByName(In_BookName, In_Sheet, 1, "Status")
    ' Номер столбца Description_status
  column_Description_status = ColumnByName(In_BookName, In_Sheet, 1, "Description_status")
  ' Номер столбца Date_finish
  column_Date_finish = ColumnByName(In_BookName, In_Sheet, 1, "Date_finish")
  ' Номер столбца Protocol
  column_Protocol = ColumnByName(In_BookName, In_Sheet, 1, "Protocol")
  
  ' Когда не указан Протокол - считаем, что текущее обновление
  If In_Protocol = "" Then
  
    rowCount = 2
    Do While Not IsEmpty(Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, 1))
    
      ' Если это запись с In_Last_Date и In_К_пор
      If (Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, column_Last_Date).Value = In_Last_Date) And (InStr(Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, column_К_пор).Value, In_К_пор) <> 0) Then
      
        ' Делаем расчет Status_persent в %: Last_Value к Value
        If Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, column_Value).Value > 0 Then
          Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, column_Status_persent).Value = Round(Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, column_Last_Value).Value / Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, column_Value).Value, 2)
        Else
          Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, column_Status_persent).Value = 1
        End If
        
        ' Статус 0-в работе 1-исполнено досрочно 10 - не исполнено 11 - исполнено
        If Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, column_Status_persent).Value < 1 Then
          Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, column_Status).Value = 0
          Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, column_Description_status).Value = "В работе"
        Else
          Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, column_Status).Value = 1
          Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, column_Description_status).Value = "Исполнено"
        End If
      
      End If
    
      ' Следующая запись
      rowCount = rowCount + 1
    Loop
  Else
    
    ' Если Протокол указан, то закрываем его
    rowCount = 2
    Do While Not IsEmpty(Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, 1))
    
      ' Если это запись с Protocol и In_К_пор
      If (Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, column_Protocol).Value = In_Protocol) And (InStr(Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, column_К_пор).Value, In_К_пор) <> 0) Then
      
        ' Делаем расчет Status_persent в %: Last_Value к Value
        If Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, column_Value).Value > 0 Then
          Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, column_Status_persent).Value = Round(Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, column_Last_Value).Value / Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, column_Value).Value, 2)
        Else
          ' План на неделю был 0, то ставим 100%
          Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, column_Status_persent).Value = 1
        End If
        
       
        ' Статус 0/1
        If Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, column_Status_persent).Value < 1 Then
          Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, column_Status).Value = 10
          Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, column_Description_status).Value = "Не исполнено"
        Else
          Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, column_Status).Value = 11
          Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, column_Description_status).Value = "Исполнено"
        End If
      
      End If
    
      ' Следующая запись
      rowCount = rowCount + 1
    Loop
    
  End If
  
End Sub

' 46. Значение между разделителями
Function Значение_между_разделителями(In_Строка, In_Разделитель, In_Начало, In_Конец) As String
Dim i, Позиция_начала, Позиция_конца As Integer
Dim Счетчик_найденных_разделителей As Byte

  Значение_между_разделителями = In_Строка
  ' Россия; Тюменская область; г. Тюмень; ул. Линейная; Дом:7;
  If Len(In_Строка) <> 0 Then
    Значение_между_разделителями = ""
    Позиция_начала = 0
    Позиция_конца = 0
    Счетчик_найденных_разделителей = 0
    ' Находим
    For i = 1 To Len(In_Строка)
      If Mid(In_Строка, i, 1) = In_Разделитель Then
        Счетчик_найденных_разделителей = Счетчик_найденных_разделителей + 1
        ' Если это начало
        If Счетчик_найденных_разделителей = In_Начало Then
          Позиция_начала = i
        End If
        ' Если это конец
        If Счетчик_найденных_разделителей = In_Конец Then
          Позиция_конца = i
        End If
      End If
    Next i
    ' Если Позиция начала и конца не нулевые
    If (Позиция_начала <> 0) And (Позиция_конца <> 0) Then
      Значение_между_разделителями = Mid(In_Строка, (Позиция_начала + 1), (Позиция_конца - Позиция_начала - 1))
    End If
  End If
  
End Function

' 47. Удаление ненужных символов: 10 и еще перевод строки 13
Function delSym(In_Str) As String
  delSym = Replace(In_Str, Chr(10), "")
  ' Добавил еще
  delSym = Replace(delSym, Chr(13), "")
End Function

' 48. Добавление символа перевода строки
Function createBlankStr(In_Count) As String
  ' Инициализация
  createBlankStr = ""
  ' Добавляем перевод строк
  For i = 1 To In_Count
    createBlankStr = createBlankStr + Chr(13)
  Next i
End Function

' 49. Получение номера недели из формата строки: Для Capacity 11.03.2020 г.: 10.02.03.20-08.03.20, 11.09.03.20-15.03.20 (один дефиз и 5 точек)
Function rowOfWeekPeriod(In_CellWeekPeriod) As Byte
Dim Позиция_точки_1, Позиция_точки_2, Позиция_точки_3, Позиция_точки_4, Позиция_точки_5, Позиция_дефис As Byte
  ' Инициализация
  rowOfWeekPeriod = 0
  '
  Позиция_точки_1 = InStr(In_CellWeekPeriod, ".")
  If Позиция_точки_1 <> 0 Then
    Позиция_точки_2 = InStr(Mid(In_CellWeekPeriod, Позиция_точки_1 + 1, Len(In_CellWeekPeriod) - Позиция_точки_1), ".")
    Позиция_точки_3 = InStr(Mid(In_CellWeekPeriod, Позиция_точки_2 + 1, Len(In_CellWeekPeriod) - Позиция_точки_2), ".")
    Позиция_точки_4 = InStr(Mid(In_CellWeekPeriod, Позиция_точки_3 + 1, Len(In_CellWeekPeriod) - Позиция_точки_3), ".")
    Позиция_точки_5 = InStr(Mid(In_CellWeekPeriod, Позиция_точки_4 + 1, Len(In_CellWeekPeriod) - Позиция_точки_4), ".")
    Позиция_дефис = InStr(In_CellWeekPeriod, "-")
    ' Финальная проверка
    If (Позиция_точки_1 <> 0) And (Позиция_точки_2 <> 0) And (Позиция_точки_3 <> 0) And (Позиция_точки_4 <> 0) And (Позиция_точки_5 <> 0) And (Позиция_дефис <> 0) Then
      ' Берем номер недели
      rowOfWeekPeriod = CInt(Mid(In_CellWeekPeriod, 1, Позиция_точки_1 - 1))
    End If
  End If
  
End Function

' 50. Очищаем ячейки отчета - медленный метод, есть версия clearСontents2
Sub clearСontents(In_BookName, In_Sheet, In_StartRange, In_EndRange)
Dim rowCount, ColumnCount, startRow, endRow, startColumn, endColumn As Byte
  '
  startRow = Workbooks(In_BookName).Sheets(In_Sheet).Range(In_StartRange).Row
  endRow = Workbooks(In_BookName).Sheets(In_Sheet).Range(In_EndRange).Row
  startColumn = Workbooks(In_BookName).Sheets(In_Sheet).Range(In_StartRange).Column
  endColumn = Workbooks(In_BookName).Sheets(In_Sheet).Range(In_EndRange).Column
    
  ' Двигаемся сначала по столбцу, потом по строке
  ColumnCount = startColumn
  Do While (ColumnCount <= endColumn)
    
    ' По срокам
    rowCount = startRow
    Do While (rowCount <= endRow)
      Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, ColumnCount).Value = ""
      Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, ColumnCount).HorizontalAlignment = xlLeft
      ' Убираем заливку
      Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, ColumnCount).Interior.Pattern = xlNone
      Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, ColumnCount).Interior.TintAndShade = 0
      Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, ColumnCount).Interior.PatternTintAndShade = 0
      ' Следующая строка
      rowCount = rowCount + 1
    Loop
    ' Следующий столбец
    ColumnCount = ColumnCount + 1
  Loop
  
End Sub

' 51. Очищаем ячейки отчета - скоростной метод
Sub clearСontents2(In_BookName, In_Sheet, In_StartRange, In_EndRange)
  '
  Workbooks(In_BookName).Sheets(In_Sheet).Range(In_StartRange + ":" + In_EndRange).Value = ""
  Workbooks(In_BookName).Sheets(In_Sheet).Range(In_StartRange + ":" + In_EndRange).HorizontalAlignment = xlLeft
  ' Убираем заливку
  Workbooks(In_BookName).Sheets(In_Sheet).Range(In_StartRange + ":" + In_EndRange).Interior.Pattern = xlNone
  Workbooks(In_BookName).Sheets(In_Sheet).Range(In_StartRange + ":" + In_EndRange).Interior.TintAndShade = 0
  Workbooks(In_BookName).Sheets(In_Sheet).Range(In_StartRange + ":" + In_EndRange).Interior.PatternTintAndShade = 0
  ' Убираем линии на границу
  Workbooks(In_BookName).Sheets(In_Sheet).Range(In_StartRange + ":" + In_EndRange).Borders(xlDiagonalDown).LineStyle = xlNone
  Workbooks(In_BookName).Sheets(In_Sheet).Range(In_StartRange + ":" + In_EndRange).Borders(xlDiagonalUp).LineStyle = xlNone
  Workbooks(In_BookName).Sheets(In_Sheet).Range(In_StartRange + ":" + In_EndRange).Borders(xlEdgeLeft).LineStyle = xlNone
  Workbooks(In_BookName).Sheets(In_Sheet).Range(In_StartRange + ":" + In_EndRange).Borders(xlEdgeTop).LineStyle = xlNone
  Workbooks(In_BookName).Sheets(In_Sheet).Range(In_StartRange + ":" + In_EndRange).Borders(xlEdgeBottom).LineStyle = xlNone
  Workbooks(In_BookName).Sheets(In_Sheet).Range(In_StartRange + ":" + In_EndRange).Borders(xlEdgeRight).LineStyle = xlNone
  Workbooks(In_BookName).Sheets(In_Sheet).Range(In_StartRange + ":" + In_EndRange).Borders(xlInsideVertical).LineStyle = xlNone
  Workbooks(In_BookName).Sheets(In_Sheet).Range(In_StartRange + ":" + In_EndRange).Borders(xlInsideHorizontal).LineStyle = xlNone

End Sub

' 51.2 Очищаем ячейки отчета - скоростной метод
Sub clearСontents3(In_BookName, In_Sheet, In_StartRow, In_StartColumn, In_EndRow, In_EndColumn)
' Dim In_StartRange, In_EndRange As Range
  '
  ' In_StartRange = Workbooks(In_BookName).Cells(In_StartRow, In_StartColumn).Range
  ' In_EndRange = Workbooks(In_BookName).Cells(In_EndRow, In_EndColumn).Range
  '
  Workbooks(In_BookName).Sheets(In_Sheet).Range(Cells(In_StartRow, In_StartColumn), Cells(In_EndRow, In_EndColumn)).Value = ""
  ' Workbooks(In_BookName).Sheets(In_Sheet).Range(Cells(In_StartRow, In_StartColumn), Cells(In_EndRow, In_EndColumn)).HorizontalAlignment = xlLeft
  ' Убираем заливку
  Workbooks(In_BookName).Sheets(In_Sheet).Range(Cells(In_StartRow, In_StartColumn), Cells(In_EndRow, In_EndColumn)).Interior.Pattern = xlNone
  Workbooks(In_BookName).Sheets(In_Sheet).Range(Cells(In_StartRow, In_StartColumn), Cells(In_EndRow, In_EndColumn)).Interior.TintAndShade = 0
  Workbooks(In_BookName).Sheets(In_Sheet).Range(Cells(In_StartRow, In_StartColumn), Cells(In_EndRow, In_EndColumn)).Interior.PatternTintAndShade = 0
End Sub


' 52. Заливка ячейки цветом:
Sub setColorCells(In_BookName, In_Sheet, In_RowStart, In_ColumnStart, In_RowEnd, In_ColumnEnd)
Dim RangeStart, RangeEnd As String
  ' Определение диапазона
  RangeStart = ConvertToLetter(In_ColumnStart) + CStr(In_RowStart)
  RangeEnd = ConvertToLetter(In_ColumnEnd) + CStr(In_RowEnd)
  ' Заливка диапазона
  Workbooks(In_BookName).Sheets(In_Sheet).Range(RangeStart + ":" + RangeEnd).Interior.Pattern = xlSolid
  Workbooks(In_BookName).Sheets(In_Sheet).Range(RangeStart + ":" + RangeEnd).Interior.PatternColorIndex = xlAutomatic
  Workbooks(In_BookName).Sheets(In_Sheet).Range(RangeStart + ":" + RangeEnd).Interior.ThemeColor = xlThemeColorAccent5
  Workbooks(In_BookName).Sheets(In_Sheet).Range(RangeStart + ":" + RangeEnd).Interior.TintAndShade = 0.599963377788629
  Workbooks(In_BookName).Sheets(In_Sheet).Range(RangeStart + ":" + RangeEnd).Interior.PatternTintAndShade = 0
End Sub

' 53. Доля (лучше в округлении ставить 3)
Function РассчетДоли(In_План, In_Факт, In_Dec) As Double
  If In_План > 0 Then
    РассчетДоли = Round((In_Факт / In_План), In_Dec)
  Else
    РассчетДоли = 0
  End If
End Function

' 54. Формат процентов
Function cellsNumberFormat(In_Value) As String
Dim ValueVar As Double
  cellsNumberFormat = "0.0%"
  
  ' Умножаем на 100 и анализируем
  ValueVar = In_Value * 100

  ' Если 0%
  If In_Value = 0 Then
    cellsNumberFormat = "0%"
  End If
  
  ' Если целое 5,0, если 1,1-1,0, то
  If ((ValueVar) - Int(ValueVar)) = 0 Then
    cellsNumberFormat = "0%"
  End If
  
End Function

' 55. Установка периода для формирования отчета из Ритейла
Sub setPriodReport(In_Sheet, In_Date)
Dim Range_str As String
Dim Range_Row, Range_Column As Byte
    
  ' Если это Активы (Лист3) или Банкострахование (Лист10)
  If (In_Sheet = "Лист3") Or (In_Sheet = "Лист10") Then
    ' Находим ячейку "Период:"
    Range_str = RangeByValue(ThisWorkbook.Name, In_Sheet, "Период:", 100, 100)
    Range_Row = Workbooks(ThisWorkbook.Name).Sheets(In_Sheet).Range(Range_str).Row
    Range_Column = Workbooks(ThisWorkbook.Name).Sheets(In_Sheet).Range(Range_str).Column
    ' Установка даты с ___ по ____
    ThisWorkbook.Sheets(In_Sheet).Cells(Range_Row, Range_Column + 1).Value = YearStartDate(In_Date)
    ThisWorkbook.Sheets(In_Sheet).Cells(Range_Row, Range_Column + 3).Value = In_Date - 1
  End If
  
  ' Если это карты
  If In_Sheet = "Лист5" Then
    ' Находим ячейку "Период:"
    Range_str = RangeByValue(ThisWorkbook.Name, In_Sheet, "Период:", 100, 100)
    Range_Row = Workbooks(ThisWorkbook.Name).Sheets(In_Sheet).Range(Range_str).Row
    Range_Column = Workbooks(ThisWorkbook.Name).Sheets(In_Sheet).Range(Range_str).Column
    ' Установка даты с ___ по ____
    ThisWorkbook.Sheets(In_Sheet).Cells(Range_Row, Range_Column + 1).Value = YearStartDate(In_Date)
    ThisWorkbook.Sheets(In_Sheet).Cells(Range_Row, Range_Column + 3).Value = In_Date
  End If
  
End Sub

' 56. Определение даты начала года по текущей дате
Function YearStartDate(In_Date) As Date
Dim Год As Integer
  ' Год
  Год = Year(In_Date)
  YearStartDate = CDate("01.01." + CStr(Год))
End Function

' 58. "Период по" на листе
' Период: 01.01.2020  по  06.07.2020 periodFromSheet -> возвращает 06.07.2020
Function periodFromSheet(In_Sheet) As Date
Dim Range_str As String
Dim Range_Row, Range_Column As Byte
    
    ' Находим ячейку "Период:"
    Range_str = RangeByValue(ThisWorkbook.Name, In_Sheet, "Период:", 100, 100)
    Range_Row = Workbooks(ThisWorkbook.Name).Sheets(In_Sheet).Range(Range_str).Row
    Range_Column = Workbooks(ThisWorkbook.Name).Sheets(In_Sheet).Range(Range_str).Column
    '
    periodFromSheet = CDate(ThisWorkbook.Sheets(In_Sheet).Cells(Range_Row, Range_Column + 3).Value)
    
End Function

' Период: 01.01.2020  по  06.07.2020  -> periodFromSheet2("Лист3", 1) возвращает 01.01.2020. periodFromSheet2("Лист3", 2) возвращает 06.07.2020
Function periodFromSheet2(In_Sheet, In_Period) As Date
Dim Range_str As String
Dim Range_Row, Range_Column, смещение As Byte
    
  ' Смещение:
  смещение = 3
  
  Select Case Weekday(In_Date, vbMonday)
    Case 1
      смещение = 1
    Case 2
      смещение = 3
  End Select

    ' Находим ячейку "Период:"
    Range_str = RangeByValue(ThisWorkbook.Name, In_Sheet, "Период:", 100, 100)
    Range_Row = Workbooks(ThisWorkbook.Name).Sheets(In_Sheet).Range(Range_str).Row
    Range_Column = Workbooks(ThisWorkbook.Name).Sheets(In_Sheet).Range(Range_str).Column
    '
    periodFromSheet2 = CDate(ThisWorkbook.Sheets(In_Sheet).Cells(Range_Row, Range_Column + смещение).Value)
    
End Function


' 59. "Хэштег" на листе
Function hashTagFromSheet(In_Sheet) As String
Dim Range_str As String
Dim Range_Row, Range_Column As Byte
    
    ' Находим ячейку "Хэштэг:"
    Range_str = RangeByValue(ThisWorkbook.Name, In_Sheet, "Хэштэг:", 100, 100)
    Range_Row = Workbooks(ThisWorkbook.Name).Sheets(In_Sheet).Range(Range_str).Row
    Range_Column = Workbooks(ThisWorkbook.Name).Sheets(In_Sheet).Range(Range_str).Column
    '
    hashTagFromSheet = ThisWorkbook.Sheets(In_Sheet).Cells(Range_Row, Range_Column + 1).Value
      
End Function

' 59.1 "Хэштег2" на листе
Function hashTagFromSheetII(In_Sheet, In_Number) As String
Dim Range_str As String
Dim Range_Row, Range_Column As Byte
    
    ' Находим ячейку "Хэштэг:"
    Range_str = RangeByValue(ThisWorkbook.Name, In_Sheet, "Хэштэг" + CStr(In_Number) + ":", 100, 100)
    Range_Row = Workbooks(ThisWorkbook.Name).Sheets(In_Sheet).Range(Range_str).Row
    Range_Column = Workbooks(ThisWorkbook.Name).Sheets(In_Sheet).Range(Range_str).Column
    '
    hashTagFromSheetII = ThisWorkbook.Sheets(In_Sheet).Cells(Range_Row, Range_Column + 1).Value
      
End Function


' 60. Тема: на листе
Function subjectFromSheet(In_Sheet) As String
Dim Range_str As String
Dim Range_Row, Range_Column As Byte
    
    ' Находим ячейку "Тема:"
    Range_str = RangeByValue(ThisWorkbook.Name, In_Sheet, "Тема:", 100, 100)
    Range_Row = Workbooks(ThisWorkbook.Name).Sheets(In_Sheet).Range(Range_str).Row
    Range_Column = Workbooks(ThisWorkbook.Name).Sheets(In_Sheet).Range(Range_str).Column
    '
    subjectFromSheet = ThisWorkbook.Sheets(In_Sheet).Cells(Range_Row, Range_Column + 1).Value
      
End Function

' 60.1 Тема2: на листе
Function subjectFromSheetII(In_Sheet, In_Number) As String
Dim Range_str As String
Dim Range_Row, Range_Column As Byte
    
    ' Находим ячейку "Тема:"
    Range_str = RangeByValue(ThisWorkbook.Name, In_Sheet, "Тема" + CStr(In_Number) + ":", 100, 100)
    Range_Row = Workbooks(ThisWorkbook.Name).Sheets(In_Sheet).Range(Range_str).Row
    Range_Column = Workbooks(ThisWorkbook.Name).Sheets(In_Sheet).Range(Range_str).Column
    '
    subjectFromSheetII = ThisWorkbook.Sheets(In_Sheet).Cells(Range_Row, Range_Column + 1).Value
      
End Function


' 61. Дата отчета из имени файла, например: Отчетность ЕСУП_итог 15.03.2020. Результат - как String, чтобы обрабатывать если даты нет в имени!
Function getDateReportFromFileName(In_ReportName_String) As String
Dim поз_точки As Byte
  
  ' Инициализация
  getDateReportFromFileName = ""
  
  ' Находим позицию первой точки
  поз_точки = InStr(In_ReportName_String, ".")
  
  getDateReportFromFileName = Mid(In_ReportName_String, поз_точки - 2, 10)
  
End Function

' 62. Дата начала месяца dateBeginMonth (имя функции меняем на monthStartDate, так как ранее уже использовалось имя переменной dateBeginMonth (на Лист3))
Function monthStartDate(In_Date) As Date
    ' Дата начала месяца
    monthStartDate = CDate("01." + Mid(CStr(In_Date), 4, 7))
End Function

' 63. Получение даты из Отчета: "Отчет обновлен на 20.03.2020 за 20.03.2020 (16.30)" или "Отчет обновлен на 23.03.2020 за 20-22.03.2020 (полный день)"
Function getDate_Отчет_План_Факт_по_продуктам_ИСЖ_НСЖ(In_StringWithDate) As Date
Dim позиция_скобки As Byte
  позиция_скобки = InStr(In_StringWithDate, "(")
  getDate_Отчет_План_Факт_по_продуктам_ИСЖ_НСЖ = CDate(Mid(In_StringWithDate, позиция_скобки - 11, 10))
End Function

' 64. "Хэштег2" на листе
Function hashTagFromSheet2(In_Sheet) As String
Dim Range_str As String
Dim Range_Row, Range_Column As Byte
    ' Находим ячейку "Хэштэг:"
    Range_str = RangeByValue(ThisWorkbook.Name, In_Sheet, "Хэштэг2:", 100, 100)
    Range_Row = Workbooks(ThisWorkbook.Name).Sheets(In_Sheet).Range(Range_str).Row
    Range_Column = Workbooks(ThisWorkbook.Name).Sheets(In_Sheet).Range(Range_str).Column
    '
    hashTagFromSheet2 = ThisWorkbook.Sheets(In_Sheet).Cells(Range_Row, Range_Column + 1).Value
End Function

' 65. Тема2: на листе
Function subjectFromSheet2(In_Sheet) As String
Dim Range_str As String
Dim Range_Row, Range_Column As Byte
    
    ' Находим ячейку "Тема:"
    Range_str = RangeByValue(ThisWorkbook.Name, In_Sheet, "Тема2:", 100, 100)
    Range_Row = Workbooks(ThisWorkbook.Name).Sheets(In_Sheet).Range(Range_str).Row
    Range_Column = Workbooks(ThisWorkbook.Name).Sheets(In_Sheet).Range(Range_str).Column
    '
    subjectFromSheet2 = ThisWorkbook.Sheets(In_Sheet).Cells(Range_Row, Range_Column + 1).Value
      
End Function

' 66. Список получателей: на листе
Function recipientList(In_Sheet) As String
Dim Range_str As String
Dim Range_Row, Range_Column As Byte
    
    ' Находим ячейку "Список получателей:"
    Range_str = RangeByValue(ThisWorkbook.Name, In_Sheet, "Список получателей:", 100, 100)
    Range_Row = Workbooks(ThisWorkbook.Name).Sheets(In_Sheet).Range(Range_str).Row
    Range_Column = Workbooks(ThisWorkbook.Name).Sheets(In_Sheet).Range(Range_str).Column
    '
    recipientList = ThisWorkbook.Sheets(In_Sheet).Cells(Range_Row, Range_Column + 2).Value
      
End Function

' 66.2 Список получателей2: на листе
Function recipientList2(In_Sheet) As String
Dim Range_str As String
Dim Range_Row, Range_Column As Byte
    
    ' Находим ячейку "Список получателей2:"
    Range_str = RangeByValue(ThisWorkbook.Name, In_Sheet, "Список получателей2:", 100, 100)
    Range_Row = Workbooks(ThisWorkbook.Name).Sheets(In_Sheet).Range(Range_str).Row
    Range_Column = Workbooks(ThisWorkbook.Name).Sheets(In_Sheet).Range(Range_str).Column
    '
    recipientList2 = ThisWorkbook.Sheets(In_Sheet).Cells(Range_Row, Range_Column + 2).Value
      
End Function

' 67. Город по наименованию офиса
' Function cityOfficeName(In_Office2_Name As String) As String
Function cityOfficeName(In_Office2_Name) As String

  cityOfficeName = In_Office2_Name
  ' Тюмень
  If InStr(In_Office2_Name, "Тюменский") <> 0 Then
    cityOfficeName = "Тюмень"
  End If
  ' Сургут
  If InStr(In_Office2_Name, "Сургутский") <> 0 Then
    cityOfficeName = "Сургут"
  End If
  ' Нижневартовск
  If InStr(In_Office2_Name, "Нижневартовский") <> 0 Then
    cityOfficeName = "Н-Вартовск"
  End If
  ' Новый Уренгой
  If InStr(In_Office2_Name, "Новоуренгойский") <> 0 Then
    cityOfficeName = "Н-Уренгой"
  End If
  ' Тарко-Сале
  If InStr(In_Office2_Name, "Тарко-Сале") <> 0 Then
    cityOfficeName = "Тарко-Сале"
  End If
End Function

' 67.2 Город по наименованию офиса
Function cityOfficeNameByNumber(In_NumberOffice) As String

          Select Case In_NumberOffice
          Case 1 ' ОО «Тюменский»
            cityOfficeNameByNumber = "Тюмень"
          Case 2 ' ОО «Сургутский»
            cityOfficeNameByNumber = "Сургут"
          Case 3 ' ОО «Нижневартовский»
            cityOfficeNameByNumber = "Н-Вартовск"
          Case 4 ' ОО «Новоуренгойский»
            cityOfficeNameByNumber = "Н-Уренгой"
          Case 5 ' ОО «Тарко-Сале»
            cityOfficeNameByNumber = "Тарко-Сале"
        End Select


End Function


' 68. Дата начала квартала (для рассчета накопительным по кварталу)
Function quarterStartDate(In_Date) As Date

  ' Месяц
  Select Case Month(In_Date)
        ' 1 кв. - 01.01.YYYY
        Case 1, 2, 3
          quarterStartDate = CDate("01.01." + CStr(Year(In_Date)))
        ' 2 кв. - 01.04.YYYY
        Case 4, 5, 6
          quarterStartDate = CDate("01.04." + CStr(Year(In_Date)))
        ' 3 кв. - 01.07.YYYY
        Case 7, 8, 9
          quarterStartDate = CDate("01.07." + CStr(Year(In_Date)))
        ' 4 кв. - 01.10.YYYY
        Case 10, 11, 12
          quarterStartDate = CDate("01.10." + CStr(Year(In_Date)))
      End Select
  
End Function

' 68.1 Дата начала второго месяца квартала
Function quarterSecondMonthStartDate(In_Date) As Date

  ' Месяц
  Select Case Month(In_Date)
        ' 1 кв. - 01.01.YYYY
        Case 1, 2, 3
          quarterSecondMonthStartDate = CDate("01.02." + CStr(Year(In_Date)))
        ' 2 кв. - 01.04.YYYY
        Case 4, 5, 6
          quarterSecondMonthStartDate = CDate("01.05." + CStr(Year(In_Date)))
        ' 3 кв. - 01.07.YYYY
        Case 7, 8, 9
          quarterSecondMonthStartDate = CDate("01.08." + CStr(Year(In_Date)))
        ' 4 кв. - 01.10.YYYY
        Case 10, 11, 12
          quarterSecondMonthStartDate = CDate("01.11." + CStr(Year(In_Date)))
      End Select
  
End Function


' 69. Наименование квартала (по дате)
Function quarterName(In_Date) As String

  ' Месяц
  Select Case Month(In_Date)
        ' 1 кв. - 01.01.YYYY
        Case 1, 2, 3
          quarterName = "1 кв. " + CStr(Year(In_Date)) + " г."
        ' 2 кв. - 01.04.YYYY
        Case 4, 5, 6
          quarterName = "2 кв. " + CStr(Year(In_Date)) + " г."
        ' 3 кв. - 01.07.YYYY
        Case 7, 8, 9
          quarterName = "3 кв. " + CStr(Year(In_Date)) + " г."
        ' 4 кв. - 01.10.YYYY
        Case 10, 11, 12
          quarterName = "4 кв. " + CStr(Year(In_Date)) + " г."
      End Select
  
End Function

' 69.1 Наименование квартала (по дате)
Function quarterName2(In_Date) As String

  ' Месяц
  Select Case Month(In_Date)
        ' 1 кв. - 01.01.YYYY
        Case 1, 2, 3
          quarterName2 = "1Q " + CStr(Year(In_Date)) + " г."
        ' 2 кв. - 01.04.YYYY
        Case 4, 5, 6
          quarterName2 = "2Q " + CStr(Year(In_Date)) + " г."
        ' 3 кв. - 01.07.YYYY
        Case 7, 8, 9
          quarterName2 = "3Q " + CStr(Year(In_Date)) + " г."
        ' 4 кв. - 01.10.YYYY
        Case 10, 11, 12
          quarterName2 = "4Q " + CStr(Year(In_Date)) + " г."
      End Select
  
End Function

' 69.2 Наименование квартала (по дате)
Function quarterName3(In_Date) As String

  ' Месяц
  Select Case Month(In_Date)
        ' 1 кв. - 01.01.YYYY
        Case 1, 2, 3
          quarterName3 = "1Q " + CStr(Year(In_Date))
        ' 2 кв. - 01.04.YYYY
        Case 4, 5, 6
          quarterName3 = "2Q " + CStr(Year(In_Date))
        ' 3 кв. - 01.07.YYYY
        Case 7, 8, 9
          quarterName3 = "3Q " + CStr(Year(In_Date))
        ' 4 кв. - 01.10.YYYY
        Case 10, 11, 12
          quarterName3 = "4Q " + CStr(Year(In_Date))
      End Select
  
End Function


' 69.3 Подстрока Квартал + Год "1Q20" из Даты
Function strNQYY(In_Date) As String
  
  ' Месяц
  Select Case Month(In_Date)
        ' 1 кв. - 01.01.YYYY
        Case 1, 2, 3
          strNQYY = "1Q" + Mid(CStr(Year(In_Date)), 3, 2)
        ' 2 кв. - 01.04.YYYY
        Case 4, 5, 6
          strNQYY = "2Q" + Mid(CStr(Year(In_Date)), 3, 2)
        ' 3 кв. - 01.07.YYYY
        Case 7, 8, 9
          strNQYY = "3Q" + Mid(CStr(Year(In_Date)), 3, 2)
        ' 4 кв. - 01.10.YYYY
        Case 10, 11, 12
          strNQYY = "4Q" + Mid(CStr(Year(In_Date)), 3, 2)
      End Select
  
  
End Function


' 70. Заливка ячейки цветом "светофор" - значение In_Value в %, In_Target в %
Sub Full_Color_RangeII(In_list, In_Row, In_Column, In_Value, In_Target)
  
  ' In_Value = In_Value * 100
  
  In_Value_tmp = (In_Value / In_Target) * 100
  
  ' Если до этого ячейка была цветная - сбрасываем цвет
  ThisWorkbook.Sheets(In_list).Cells(In_Row, In_Column).Interior.Color = xlNone
  ' Цвет текста - черный
  ThisWorkbook.Sheets(In_list).Cells(In_Row, In_Column).Font.Color = vbBlack
  ' От 100% и выше - Зеленый
  If (In_Value_tmp >= 100) Then
    ThisWorkbook.Sheets(In_list).Cells(In_Row, In_Column).Interior.Color = vbGreen
  End If
  ' От 90%-100% - Желтый
  If (In_Value_tmp >= 90) And (In_Value_tmp < 100) Then
    ThisWorkbook.Sheets(In_list).Cells(In_Row, In_Column).Interior.Color = vbYellow
  End If
  ' От 0% - 90% - Красный
  If (In_Value_tmp < 90) Then
    ThisWorkbook.Sheets(In_list).Cells(In_Row, In_Column).Interior.Color = vbRed
  End If
  
End Sub

' 70.1 Заливка ячейки цветом "светофор" - значение In_Value в %, In_Target в %. In_Yelow - уровень желтой зоны: 85% или 90%, значение передавать: 85 или 90
Sub Full_Color_RangeIII(In_list, In_Row, In_Column, In_Value, In_Target, In_Yelow)
  
  ' In_Value = In_Value * 100
  
  In_Value_tmp = (In_Value / In_Target) * 100
  
  ' Если до этого ячейка была цветная - сбрасываем цвет
  ThisWorkbook.Sheets(In_list).Cells(In_Row, In_Column).Interior.Color = xlNone
  ' Цвет текста - черный
  ThisWorkbook.Sheets(In_list).Cells(In_Row, In_Column).Font.Color = vbBlack
  
  ' От 100% и выше - Зеленый
  If (In_Value_tmp >= 100) Then
    ThisWorkbook.Sheets(In_list).Cells(In_Row, In_Column).Interior.Color = vbGreen
  End If
  
  ' От In_Yelow%-100% - Желтый
  If (In_Value_tmp >= In_Yelow) And (In_Value_tmp < 100) Then
    ThisWorkbook.Sheets(In_list).Cells(In_Row, In_Column).Interior.Color = vbYellow
  End If
  ' От 0% - In_Yelow% - Красный
  If (In_Value_tmp < In_Yelow) Then
    ThisWorkbook.Sheets(In_list).Cells(In_Row, In_Column).Interior.Color = vbRed
  End If
  
End Sub

' 70.2 Заливка ячейки цветом "светофор" - значение In_Value в %, In_Target в %. In_Yelow - уровень желтой зоны: 85% или 90%, значение передавать: 85 или 90
Sub Full_Color_RangeIV(In_Book, In_list, In_Row, In_Column, In_Value, In_Target, In_Yelow)
  
  ' In_Value = In_Value * 100
  
  ' План не ноль (иначе возникает ошибка при делении на нуль)
  If In_Target <> 0 Then
  
    In_Value_tmp = (In_Value / In_Target) * 100
  
    ' Если до этого ячейка была цветная - сбрасываем цвет
    Workbooks(In_Book).Sheets(In_list).Cells(In_Row, In_Column).Interior.Color = xlNone
    ' Цвет текста - черный
    Workbooks(In_Book).Sheets(In_list).Cells(In_Row, In_Column).Font.Color = vbBlack
  
    ' От 100% и выше - Зеленый
    If (In_Value_tmp >= 100) Then
      Workbooks(In_Book).Sheets(In_list).Cells(In_Row, In_Column).Interior.Color = vbGreen
    End If
  
    ' От In_Yelow%-100% - Желтый
    If (In_Value_tmp >= In_Yelow) And (In_Value_tmp < 100) Then
      Workbooks(In_Book).Sheets(In_list).Cells(In_Row, In_Column).Interior.Color = vbYellow
    End If
    
    ' От 0% - In_Yelow% - Красный
    If (In_Value_tmp < In_Yelow) Then
      Workbooks(In_Book).Sheets(In_list).Cells(In_Row, In_Column).Interior.Color = vbRed
    End If
  
  End If ' План не ноль (иначе возникает ошибка)
  
End Sub

' 70.3 Заливка ячейки цветом "светофор" - значение In_Value в %, In_Target в %
Sub Full_Color_RangeV(In_Book, In_list, In_Row, In_Column, In_Value, In_Target)
  
  ' In_Value = In_Value * 100
  
  In_Value_tmp = (In_Value / In_Target) * 100
  
  ' Если до этого ячейка была цветная - сбрасываем цвет
  Workbooks(In_Book).Sheets(In_list).Cells(In_Row, In_Column).Interior.Color = xlNone
  ' Цвет текста - черный
  Workbooks(In_Book).Sheets(In_list).Cells(In_Row, In_Column).Font.Color = vbBlack
  ' От 100% и выше - Зеленый
  If (In_Value_tmp >= 100) Then
    Workbooks(In_Book).Sheets(In_list).Cells(In_Row, In_Column).Interior.Color = vbGreen
  End If
  ' От 90%-100% - Желтый
  If (In_Value_tmp >= 90) And (In_Value_tmp < 100) Then
    Workbooks(In_Book).Sheets(In_list).Cells(In_Row, In_Column).Interior.Color = vbYellow
  End If
  ' От 0% - 90% - Красный
  If (In_Value_tmp < 90) Then
    Workbooks(In_Book).Sheets(In_list).Cells(In_Row, In_Column).Interior.Color = vbRed
  End If
  
End Sub


' 71. Установка факта План по офису на листе "План": Месяц (1-12), Офис (1-5)
Sub setResultValue(In_Month, In_Office, In_Value)
  
  ' Строка с офисом
  row_N = In_Office + 5
  
  ' Месяц - к номеру месяца прибавляем 3
  ' column_n = (2 * In_Month) + 2
  Select Case In_Month
      Case 1
        column_n = 4
      Case 2
        column_n = 6
      Case 3
        column_n = 8
      Case 4
        column_n = 12
      Case 5
        column_n = 14
      Case 6
        column_n = 16
      Case 7
        column_n = 20
      Case 8
        column_n = 22
      Case 9
        column_n = 24
      Case 10
        column_n = 28
      Case 11
        column_n = 30
      Case 12
        column_n = 32
  End Select
  
  ' Выводим значение
  ThisWorkbook.Sheets("План").Cells(row_N, column_n).Value = In_Value
    
End Sub

' 72. Получение информации по действующим сотрудникам из Таблицы BASE\ActiveStaff getInfoFromActiveStaff
Function getInfoFromActiveStaff(In_TabNumber) As String
Dim rowCount As Integer

  ' Убираем фильтр, иначе поиск не по всей таблице
  If Workbooks("ActiveStaff").Sheets("Лист1").AutoFilterMode = True Then
    ' Выключаем Автофильтр
    Workbooks("ActiveStaff").Sheets("Лист1").Cells(1, 1).AutoFilter
  End If

  ' Иницииализация результата
  getInfoFromActiveStaff = ""
  
  ' Проверяем наличие записи
  Set searchResults = Workbooks("ActiveStaff").Sheets("Лист1").Columns("A:A").Find(In_TabNumber, LookAt:=xlWhole)
  
  ' Проверяем - есть ли такая дата, если нет, то добавляем
  If searchResults Is Nothing Then
    
    ' Если не найдена - вставляем. Не работает!
    ' MsgBox (Workbooks("ActiveStaff").Sheets("Лист1").CountA(Columns(1))) - хотел найти число заполненных записей через CountA
    
  Else
    
    ' Если найдена:
    ' Дата увольнения
    If Workbooks("ActiveStaff").Sheets("Лист1").Cells(searchResults.Row, 5).Value <> "" Then
      getInfoFromActiveStaff = "Уволен " + CStr(Workbooks("ActiveStaff").Sheets("Лист1").Cells(searchResults.Row, 5).Value) + " г."
    End If
     
  End If
  
  
End Function

' 73. Преобразовать "MMYY" в номер месяца (Int)
Function decodeMMYY(In_MMYY) As Byte
  decodeMMYY = CInt(Mid(In_MMYY, 1, 2))
End Function

' 74. "MMYY" в номер месяца в январь 2020
Function decodeMMYY2(In_MMYY) As String
  decodeMMYY2 = ИмяМесяца(CDate("01." + Mid(In_MMYY, 1, 2) + ".2020")) + " 20" + Mid(In_MMYY, 3, 2)
End Function

' 75. Чтение из буфера обмена (взято из Интернет http://www.excelworld.ru/forum/10-4692-1)
' Внимание - выдает ошибку, если в буффере обмена ничего нет
Function ClipboardText()
    With GetObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
        .GetFromClipboard
        ClipboardText = .GetText
    End With
End Function

' 76. Запись в буфер обмена (взято из Интернет http://www.excelworld.ru/forum/10-4692-1)
Sub SetClipboardText(ByVal txt$)
    With GetObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
        .SetText txt$
        .PutInClipboard
    End With
End Sub

' 77. Получение имени Листа из DB
Function FindNameSheet(In_Workbooks, In_StringInSheet) As String
Dim i As Integer

  FindNameSheet = ""
    
  For i = 1 To Workbooks(In_Workbooks).Sheets.Count
       
    ' В Dashboard_new_РБ_09.09.2021 "неверная ссылка вперед или ссылка на неоткомпилированный тип" на листе с индексом 5
    If i <> 5 Then
    
      t = Workbooks(In_Workbooks).Sheets.Count
      t2 = Workbooks(In_Workbooks).Sheets(i).Visible
      t3 = InStr(Workbooks(In_Workbooks).Sheets(i).Name, In_StringInSheet)
      t4 = InStr(Workbooks(In_Workbooks).Sheets(i).Name, ".")
      t5 = Workbooks(In_Workbooks).Sheets(i).Name
    
      ' Ищем в имени листа подстроку + добавил 06.11 поиск ".", потому как был "Комм доход данные"
      If (InStr(Workbooks(In_Workbooks).Sheets(i).Name, In_StringInSheet) <> 0) And ((InStr(Workbooks(In_Workbooks).Sheets(i).Name, ".") <> 0)) Then
      
        ' Проверяем  - видимый ли этот Лист? В DB очень много скрытых листов с одинаковыми названиями
        If Workbooks(In_Workbooks).Sheets(i).Visible = xlSheetVisible Then
          FindNameSheet = Workbooks(In_Workbooks).Sheets(i).Name
        End If
      
      End If
    
    End If
    
  Next
End Function

' 78. Округление до большего при значении <1
Function ОкруглениеБольше(In_Value) As Integer
  ' Текущее округление
  ОкруглениеБольше = Round(In_Value, 0)
  ' Если значение <1, то округляем до 1
  If (In_Value > 0) And (In_Value < 1) Then
    ОкруглениеБольше = 1
  End If
  ' Если значение <0, то выдаем всегда 0
  If (In_Value <= 0) Then
    ОкруглениеБольше = 0
  End If
End Function

' 79. Прогноз месяца с учетом и без учета нерабочих дней: In_Date, In_Plan, In_Fact, In_working_days_in_the_week (5-ти/6-ти дневка), In_NonWorkingDays = 1/0 (учитывать нерабочие дни из BASE\NonWorkingDays)
Function Прогноз_месяца(In_Date, In_Plan, In_Fact, In_working_days_in_the_week, In_NonWorkingDays) As Double
  
  Дата_начала_Месяца = Date_begin_day_month(In_Date)
  
  Дата_конца_месяца = Date_last_day_month(In_Date)
  
  Число_прошедших_раб_дней = Working_days_between_dates(Дата_начала_Месяца, In_Date, 5)
  
  If In_NonWorkingDays = 1 Then
    Число_раб_дней_месяц = Working_days_between_datesII(Дата_начала_Месяца, Дата_конца_месяца, 5)
  Else
    Число_раб_дней_месяц = Working_days_between_dates(Дата_начала_Месяца, Дата_конца_месяца, 5)
  End If
  
  Прогноз_месяца = (In_Fact / Число_прошедших_раб_дней) * Число_раб_дней_месяц

End Function

' 79.1 Прогноз квартала с учетом и без учета нерабочих дней: In_Date, In_Plan, In_Fact, In_working_days_in_the_week (5-ти/6-ти дневка), In_NonWorkingDays = 1/0 (учитывать нерабочие дни из BASE\NonWorkingDays)
Function Прогноз_квартала(In_Date, In_Plan, In_Fact, In_working_days_in_the_week, In_NonWorkingDays) As Double
  
  Дата_начала_квартала = quarterStartDate(In_Date)
  
  Дата_конца_квартала = Date_last_day_quarter(In_Date)
  
  Число_прошедших_раб_дней = Working_days_between_dates(Дата_начала_квартала, In_Date, 5)
  
  If In_NonWorkingDays = 1 Then
    Число_раб_дней_квартал = Working_days_between_datesII(Дата_начала_квартала, Дата_конца_квартала, 5)
  Else
    Число_раб_дней_квартал = Working_days_between_dates(Дата_начала_квартала, Дата_конца_квартала, 5)
  End If
  
  Прогноз_квартала = (In_Fact / Число_прошедших_раб_дней) * Число_раб_дней_квартал

End Function

' 79.2 Прогноз квартала с учетом и без учета нерабочих дней: In_Date, In_Plan, In_Fact, In_working_days_in_the_week (5-ти/6-ти дневка), In_NonWorkingDays = 1/0 (учитывать нерабочие дни из BASE\NonWorkingDays)
Function Прогноз_квартала_проц(In_Date, In_Plan, In_Fact, In_working_days_in_the_week, In_NonWorkingDays) As Double
  
  Дата_начала_квартала = quarterStartDate(In_Date)
  
  Дата_конца_квартала = Date_last_day_quarter(In_Date)
  
  Число_прошедших_раб_дней = Working_days_between_dates(Дата_начала_квартала, In_Date, 5)
  
  If In_NonWorkingDays = 1 Then
    Число_раб_дней_квартал = Working_days_between_datesII(Дата_начала_квартала, Дата_конца_квартала, 5)
  Else
    Число_раб_дней_квартал = Working_days_between_dates(Дата_начала_квартала, Дата_конца_квартала, 5)
  End If
  
  Прогноз_квартала_объем = (In_Fact / Число_прошедших_раб_дней) * Число_раб_дней_квартал
  
  Прогноз_квартала_проц = РассчетДоли(In_Plan, Прогноз_квартала_объем, 3)

End Function

' 79.2.1 Прогноз месяца с учетом и без учета нерабочих дней: In_Date, In_Plan, In_Fact, In_working_days_in_the_week (5-ти/6-ти дневка), In_NonWorkingDays = 1/0 (учитывать нерабочие дни из BASE\NonWorkingDays)
Function Прогноз_месяца_проц(In_Date, In_Plan, In_Fact, In_working_days_in_the_week, In_NonWorkingDays) As Double
  
  Дата_начала_Месяца = Date_begin_day_month(In_Date)
  
  Дата_конца_месяца = Date_last_day_month(In_Date)
  
  Число_прошедших_раб_дней = Working_days_between_dates(Дата_начала_Месяца, In_Date, 5)
  
  If In_NonWorkingDays = 1 Then
    Число_раб_дней_месяц = Working_days_between_datesII(Дата_начала_Месяца, Дата_конца_месяца, 5)
  Else
    Число_раб_дней_месяц = Working_days_between_dates(Дата_начала_Месяца, Дата_конца_месяца, 5)
  End If
  
  Прогноз_месяца_объем = (In_Fact / Число_прошедших_раб_дней) * Число_раб_дней_месяц
  
  Прогноз_месяца_проц = РассчетДоли(In_Plan, Прогноз_месяца_объем, 3)

End Function


' 79.3 Факт на дату, чтобы достигнуть прогноза Квартала
Function Факт_на_дату_для_прогноза_квартала(In_Date, In_Plan, In_прогноза_квартала_проц, In_working_days_in_the_week, In_NonWorkingDays) As Double
  
  Дата_начала_квартала = quarterStartDate(In_Date)
  
  Дата_конца_квартала = Date_last_day_quarter(In_Date)
  
  Число_прошедших_раб_дней = Working_days_between_dates(Дата_начала_квартала, In_Date, 5)
  
  If In_NonWorkingDays = 1 Then
    Число_раб_дней_квартал = Working_days_between_datesII(Дата_начала_квартала, Дата_конца_квартала, 5)
  Else
    Число_раб_дней_квартал = Working_days_between_dates(Дата_начала_квартала, Дата_конца_квартала, 5)
  End If
  
  ' Прогноз_квартала_объем = (In_Fact / Число_прошедших_раб_дней) * Число_раб_дней_квартал

  Факт_на_дату_для_прогноза_квартала = (In_прогноза_квартала_проц * Число_прошедших_раб_дней * In_Plan) / Число_раб_дней_квартал

End Function


' 80. Внесение данных на лист книги
Sub setValueInBookSheet(In_Workbooks, In_Sheets, In_RowKey, In_Column, In_Value, In_maxRowInSheet, In_maxColumnInSheet)
  
End Sub

' Переход в браузере по ссылке на странице, в конце ставить "/" для перехода в категорию
Sub goToURL()
  
  SheetsVar = ThisWorkbook.ActiveSheet.Name
  rowVar = rowByValue(ThisWorkbook.Name, SheetsVar, "Ссылка:", 100, 100)
  columnVar = ColumnByValue(ThisWorkbook.Name, SheetsVar, "Ссылка:", 100, 100) + 1
  
  ' ThisWorkbook.FollowHyperlink ("http://isrb.psbnk.msk.ru/inf/6601/6622/ejednevnii_otchet_po_prodajam/")
  ThisWorkbook.FollowHyperlink (ThisWorkbook.Sheets(SheetsVar).Cells(rowVar, columnVar).Value)
  
End Sub

' Переход в браузере по ссылке на странице, в конце ставить "/" для перехода в категорию
Sub goToURL2()
  
  SheetsVar = ThisWorkbook.ActiveSheet.Name
  rowVar = rowByValue(ThisWorkbook.Name, SheetsVar, "Ссылка2:", 100, 100)
  columnVar = ColumnByValue(ThisWorkbook.Name, SheetsVar, "Ссылка2:", 100, 100) + 1
  
  ' ThisWorkbook.FollowHyperlink ("http://isrb.psbnk.msk.ru/inf/6601/6622/ejednevnii_otchet_po_prodajam/")
  ThisWorkbook.FollowHyperlink (ThisWorkbook.Sheets(SheetsVar).Cells(rowVar, columnVar).Value)
  
End Sub


' 81. Приветствие с 6 до 12 часов — утро; с 12 до 18 часов — день; с 18 до 24 часов — вечер. Time() — возвращает текущее системное время
Function Добрый_утро_день_вечер(In_Time, In_Д_д) As String
  
  Добрый_утро_день_вечер = ""

  If (In_Time >= "00:00:00") And ((In_Time <= "12:00:00")) Then
    Добрый_утро_день_вечер = In_Д_д + "оброе утро"
  End If

  If (In_Time >= "12:00:01") And ((In_Time <= "18:00:00")) Then
    Добрый_утро_день_вечер = In_Д_д + "обрый день"
  End If

  If (In_Time >= "18:00:01") And ((In_Time <= "23:59:59")) Then
    Добрый_утро_день_вечер = In_Д_д + "обрый вечер"
  End If


End Function

' 82. Дата DB с Лист7
Function dateDB_Лист_7() As Date
  
  dateDB_Лист_7 = CDate(Mid(ThisWorkbook.Sheets("Лист7").Range("B5").Value, 40, 10))

End Function


' 83. Дата DB с Лист8
Function dateDB_Лист_8() As Date
  
  dateDB_Лист_8 = CDate(Mid(ThisWorkbook.Sheets("Лист8").Range("B5").Value, 52, 10))

End Function

' 84. Удаляем файл если он есть
Sub deleteFile(In_fileName)

  If Dir(In_fileName) <> "" Then
    ' Удаляем старый отчет на диске
    Kill In_fileName
  End If

End Sub

' 85. Кодировщик наимнования продукта в короткий код
Function Product_Name_to_Product_Code(In_Product_Name) As String
  
  Product_Name_to_Product_Code = ""
  
  ' 20.06.2021 в BASE\Products создана таблица в которую будут заноситься все значения ProductName, ProductCode
  
  ' Проверяем - если BASE\Products не открыта, то открываем
  If BookIsOpen("Products") = True Then
    Книга_была_открыта = True
  Else
    Книга_была_открыта = False
    ' Открываем BASE\Products
    OpenBookInBase ("Products")
  End If
  
  ' Убираем фильтр, иначе поиск не по всей таблице
  If Workbooks("Products").Sheets("Лист1").AutoFilterMode = True Then
    ' Выключаем Автофильтр
    Workbooks("Products").Sheets("Лист1").Cells(1, 1).AutoFilter
  End If
  
  ' Выполняем поиск
  Set searchResults = Workbooks("Products").Sheets("Лист1").Columns("A:A").Find(In_Product_Name, LookAt:=xlWhole)
  
  ' Проверяем - есть ли такая дата, если нет, то добавляем
  If searchResults Is Nothing Then
    ' Если не найдена
    Product_Name_to_Product_Code = ""
  Else
    ' Если найдена
    Product_Name_to_Product_Code = Workbooks("Products").Sheets("Лист1").Cells(searchResults.Row, 2).Value
  End If

  
  ' Select Case In_Product_Name
  '       Case "Зарплатные карты 18+"
  '         Product_Name_to_Product_Code = "ЗП"
  '       Case "Портфель ЗП 18+, шт._Квартал "
  '         Product_Name_to_Product_Code = "Портфель_ЗП"
  '       Case "Потребительские кредиты"
  '         Product_Name_to_Product_Code = "ПК"
  '       Case "Кредитные карты (актив.)"
  '         Product_Name_to_Product_Code = "КК"
  '       Case "Комиссионный доход"
  '         Product_Name_to_Product_Code = "КД"
  '       Case "Пассивы"
  '         Product_Name_to_Product_Code = "Пассивы"
  '       Case "Orange Premium Club"
  '         Product_Name_to_Product_Code = "OPC"
  '       Case "Инвест"
  '         Product_Name_to_Product_Code = "ИНВ"
  '       Case "Инвест Брокер обслуж"
  '         Product_Name_to_Product_Code = "ИНВ_БО"
  '
  ' End Select
  
  ' Если BASE\Products была открыта до начала работы Product_Name_to_Product_Code, то не закрываем ее
  If Книга_была_открыта = False Then
    ' Закрываем BASE\Products
    CloseBook ("Products")
  End If
  
  ' Сообщение
  If Product_Name_to_Product_Code = "" Then
    MsgBox ("В Product_Name_to_Product_Code не найден " + In_Product_Name + "!")
  End If

End Function

' Получение единицы измерения по продукту
Function Product_Name_to_Unit(In_Product_Name) As String
  
  Product_Name_to_Unit = ""
  
  ' 20.06.2021 в BASE\Products создана таблица в которую будут заноситься все значения ProductName, ProductCode
  
  ' Проверяем - если BASE\Products не открыта, то открываем
  If BookIsOpen("Products") = True Then
    Книга_была_открыта = True
  Else
    Книга_была_открыта = False
    ' Открываем BASE\Products
    OpenBookInBase ("Products")
  End If
  
  ' Убираем фильтр, иначе поиск не по всей таблице
  If Workbooks("Products").Sheets("Лист1").AutoFilterMode = True Then
    ' Выключаем Автофильтр
    Workbooks("Products").Sheets("Лист1").Cells(1, 1).AutoFilter
  End If
  
  ' Выполняем поиск
  Set searchResults = Workbooks("Products").Sheets("Лист1").Columns("A:A").Find(In_Product_Name, LookAt:=xlWhole)
  
  ' Проверяем - есть ли такая дата, если нет, то добавляем
  If searchResults Is Nothing Then
    ' Если не найдена
    Product_Name_to_Unit = ""
  Else
    ' Если найдена
    Product_Name_to_Unit = Workbooks("Products").Sheets("Лист1").Cells(searchResults.Row, 3).Value
  End If

  ' Если BASE\Products была открыта до начала работы Product_Name_to_Unit, то не закрываем ее
  If Книга_была_открыта = False Then
    ' Закрываем BASE\Products
    CloseBook ("Products")
  End If
  
  ' Сообщение
  If Product_Name_to_Unit = "" Then
    MsgBox ("В Product_Name_to_Unit не найден " + In_Product_Name + "!")
  End If

End Function


' 85. ДеКодировщик короткий код в наимнования продукта
Function Product_Code_to_Product_Name(In_Product_Code) As String
  
  Product_Code_to_Product_Name = ""
  
  ' 20.06.2021 в BASE\Products создана таблица в которую будут заноситься все значения ProductName, ProductCode
  
  ' Проверяем - если BASE\Products не открыта, то открываем
  If BookIsOpen("Products") = True Then
    Книга_была_открыта = True
  Else
    Книга_была_открыта = False
    ' Открываем BASE\Products
    OpenBookInBase ("Products")
  End If
  
  ' Убираем фильтр, иначе поиск не по всей таблице
  If Workbooks("Products").Sheets("Лист1").AutoFilterMode = True Then
    ' Выключаем Автофильтр
    Workbooks("Products").Sheets("Лист1").Cells(1, 1).AutoFilter
  End If
  
  ' Выполняем поиск
  Set searchResults = Workbooks("Products").Sheets("Лист1").Columns("B:B").Find(In_Product_Code, LookAt:=xlWhole)
  
  ' Проверяем - есть ли такая дата, если нет, то добавляем
  If searchResults Is Nothing Then
    ' Если не найдена
    Product_Code_to_Product_Name = ""
  Else
    ' Если найдена
    Product_Code_to_Product_Name = Workbooks("Products").Sheets("Лист1").Cells(searchResults.Row, 1).Value
  End If

  
  ' Select Case In_Product_Name
  '       Case "Зарплатные карты 18+"
  '         Product_Code_to_Product_Name = "ЗП"
  '       Case "Портфель ЗП 18+, шт._Квартал "
  '         Product_Code_to_Product_Name = "Портфель_ЗП"
  '       Case "Потребительские кредиты"
  '         Product_Code_to_Product_Name = "ПК"
  '       Case "Кредитные карты (актив.)"
  '         Product_Code_to_Product_Name = "КК"
  '       Case "Комиссионный доход"
  '         Product_Code_to_Product_Name = "КД"
  '       Case "Пассивы"
  '         Product_Code_to_Product_Name = "Пассивы"
  '       Case "Orange Premium Club"
  '         Product_Code_to_Product_Name = "OPC"
  '       Case "Инвест"
  '         Product_Code_to_Product_Name = "ИНВ"
  '       Case "Инвест Брокер обслуж"
  '         Product_Code_to_Product_Name = "ИНВ_БО"
  '
  ' End Select
  
  ' Если BASE\Products была открыта до начала работы Product_Code_to_Product_Name, то не закрываем ее
  If Книга_была_открыта = False Then
    ' Закрываем BASE\Products
    CloseBook ("Products")
  End If
  
  ' Сообщение
  If Product_Code_to_Product_Name = "" Then
    MsgBox ("В Product_Code_to_Product_Name не найден " + In_Product_Name + "!")
  End If

End Function


' 86. Убрать заливку цветом ячейки
Sub Убрать_заливку_цветом(In_Workbooks, In_Sheets, In_Row, In_Col)
    
  Workbooks(In_Workbooks).Sheets(In_Sheets).Cells(In_Row, In_Col).Interior.Pattern = xlNone
  Workbooks(In_Workbooks).Sheets(In_Sheets).Cells(In_Row, In_Col).Interior.TintAndShade = 0
  Workbooks(In_Workbooks).Sheets(In_Sheets).Cells(In_Row, In_Col).Interior.PatternTintAndShade = 0

End Sub

' 87. Получение значения записи из таблицы каталога BASE\"Workbook"
Function getDataFrom_BASE_Workbook(In_BookName, In_Sheet, In_FieldKeyName, In_FieldKeyValue, In_FieldName1, In_Открывать_Закрывать_BookName)
    
  getDataFrom_BASE_Workbook = "not found"
  
  ' Если In_Открывать_Закрывать_BookName = 1 то перед работой функции getDataFrom_BASE_Workbook открываем таблицу
  If In_Открывать_Закрывать_BookName = 1 Then
    ' Открываем BASE\Sales
    OpenBookInBase (In_BookName)
  End If
  
  ' Убираем фильтр, иначе поиск не по всей таблице
  If Workbooks(In_BookName).Sheets(In_Sheet).AutoFilterMode = True Then
    ' Выключаем Автофильтр
    Workbooks(In_BookName).Sheets(In_Sheet).Cells(1, 1).AutoFilter
  End If
  
  ' Выполняем поиск по ключу
  ' Проверяем наличие записи In_FieldKeyName - In_FieldKeyValue
  Литера_столбца = ConvertToLetter(ColumnByName(In_BookName, In_Sheet, 1, In_FieldKeyName))
  Set searchResults = Workbooks(In_BookName).Sheets(In_Sheet).Columns(Литера_столбца + ":" + Литера_столбца).Find(In_FieldKeyValue, LookAt:=xlWhole)
  
  ' Проверяем - есть ли такая дата, если нет, то добавляем
  If searchResults Is Nothing Then
    ' Если не найдена
    getDataFrom_BASE_Workbook = "not found"
  Else
    ' Если найдена, то делаем поиск значения записи
    rowCount = searchResults.Row
    column_In_FieldName1 = ColumnByValue(In_BookName, In_Sheet, In_FieldName1, 1, 500)
    getDataFrom_BASE_Workbook = Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, column_In_FieldName1).Value
  End If
  
  ' Если In_Открывать_Закрывать_BookName = 1 то перед работой функции getDataFrom_BASE_Workbook открываем таблицу
  If In_Открывать_Закрывать_BookName = 1 Then
    ' Закрываем BASE\Sales
    CloseBook (In_BookName)
  End If
End Function

' 88. Получение значения записи из таблицы каталога BASE\"Workbook" с проверкой открытия/закрытия. Параметр In_Открывать_Закрывать_BookName - не используется!
Function getDataFrom_BASE_Workbook2(In_BookName, In_Sheet, In_FieldKeyName, In_FieldKeyValue, In_FieldName1, In_Открывать_Закрывать_BookName)
    
  getDataFrom_BASE_Workbook2 = "not found"
  
  ' Проверяем - если BASE\... не открыта, то открываем
  If BookIsOpen(In_BookName) = True Then
    Книга_была_открыта = True
  Else
    Книга_была_открыта = False
    ' Открываем BASE\Products
    OpenBookInBase (In_BookName)
  End If
    
  ' Убираем фильтр, иначе поиск не по всей таблице
  If Workbooks(In_BookName).Sheets(In_Sheet).AutoFilterMode = True Then
    ' Выключаем Автофильтр
    Workbooks(In_BookName).Sheets(In_Sheet).Cells(1, 1).AutoFilter
  End If
    
  ' Выполняем поиск по ключу
  ' Проверяем наличие записи In_FieldKeyName - In_FieldKeyValue
  Литера_столбца = ConvertToLetter(ColumnByName(In_BookName, In_Sheet, 1, In_FieldKeyName))
  Set searchResults = Workbooks(In_BookName).Sheets(In_Sheet).Columns(Литера_столбца + ":" + Литера_столбца).Find(In_FieldKeyValue, LookAt:=xlWhole)
  
  ' Проверяем - есть ли такая дата, если нет, то добавляем
  If searchResults Is Nothing Then
    ' Если не найдена
    getDataFrom_BASE_Workbook2 = "not found"
  Else
    ' Если найдена, то делаем поиск значения записи
    rowCount = searchResults.Row
    column_In_FieldName1 = ColumnByValue(In_BookName, In_Sheet, In_FieldName1, 1, 500)
    getDataFrom_BASE_Workbook2 = Workbooks(In_BookName).Sheets(In_Sheet).Cells(rowCount, column_In_FieldName1).Value
  End If
  
  
  
  ' Если BASE\Products была открыта до начала работы Функции, то не закрываем ее
  If Книга_была_открыта = False Then
    ' Закрываем BASE\Products
    CloseBook (In_BookName)
  End If
  
End Function


' 88. Получение Факта за квартал по продукту из таблицы каталога BASE\"Sales_Office". Если факта нет в Date_DD, то возвращаем 0!
Function Факт_Q_на_дату(In_OfficeNumber, In_Product_Code, In_Date) As Double

  Месяц_даты = Month(In_Date)
  Год_даты = Year(In_Date)
  
  ' Идентификатор ID_Rec для In_Date:
  ID_RecVar = CStr(CStr(In_OfficeNumber) + "-" + strMMYY(In_Date) + "-" + In_Product_Code)

  ' Текущие значения в месяце: Date_01 (N), Date_02 (O), Date_03 (P), Date_04 (Q), Date_05 Date_06 Date_07 Date_08 Date_09 Date_10 Date_11 Date_12 Date_13 Date_14 Date_15 Date_16 Date_17 Date_18 Date_19 Date_20 Date_21 Date_22 Date_23 Date_24 Date_25 Date_26 Date_27 Date_28 Date_29 Date_30 Date_31
  Column_Date_DD = "Date_" + Mid(In_Date, 1, 2)
  
  ' Записываем результат в переменную для проверки
  result_Месяц_getDataFrom_BASE_Workbook2 = getDataFrom_BASE_Workbook2("Sales_Office", "Лист1", "ID_Rec", ID_RecVar, Column_Date_DD, 0)
  
  ' Если результат не пустой - возникает, когда добавляется новый продукт и по нему не возвращает "not found", а приходит Empty. Было 12.08.2021 при добавлении Портфеля АУМ из DB от 10.08.2021
  If result_Месяц_getDataFrom_BASE_Workbook2 <> "" Then
  
    ' Проверяем результат поиска записи в BASE\Sales_Office по месяцу
    If result_Месяц_getDataFrom_BASE_Workbook2 <> "not found" Then

      ' Берем факт из Date_ДД
      ' Факт_Q_на_дату = getDataFrom_BASE_Workbook2("Sales_Office", "Лист1", "ID_Rec", ID_RecVar, Column_Date_DD, 0)
      Факт_Q_на_дату = result_Месяц_getDataFrom_BASE_Workbook2

      ' Если полученный факт на Дату <>0
      ' If Факт_Q_на_дату <> 0 Then

        ' Дата может относится:
        ' - 1-му месяцу Q
        ' If (Месяц_даты = 1) Or (Месяц_даты = 4) Or (Месяц_даты = 7) Then
          ' Берем факт из Date_ДД
          ' Факт_Q_на_дату = getDataFrom_BASE_Workbook2("Sales_Office", "Лист1", "ID_Rec", ID_RecVar, Column_Date_DD, 0)
        ' End If
  
        ' Если In_Datr относится к 2-му месяцу Q
        If (Месяц_даты = 2) Or (Месяц_даты = 5) Or (Месяц_даты = 8) Or (Месяц_даты = 11) Then
          ' Берем факт из Date_ДД
          '
          ' Идентификатор ID_Rec для 1 месяца Q:
          ID_RecVar = CStr(CStr(In_OfficeNumber) + "-" + strMMYY(Date_begin_day_quarter(In_Date)) + "-" + In_Product_Code)
          
          ' Прибавляем факт 1 месяца
          result_getDataFrom_BASE_Workbook2_Fact = getDataFrom_BASE_Workbook2("Sales_Office", "Лист1", "ID_Rec", ID_RecVar, "Fact", 0)
          If result_getDataFrom_BASE_Workbook2_Fact <> "not found" Then
            Факт_Q_на_дату = Факт_Q_на_дату + result_getDataFrom_BASE_Workbook2_Fact ' getDataFrom_BASE_Workbook2("Sales_Office", "Лист1", "ID_Rec", ID_RecVar, "Fact", 0)
          End If
          
        End If
  
        ' Если In_Datr относится к 3-му месяцу Q
        If (Месяц_даты = 3) Or (Месяц_даты = 6) Or (Месяц_даты = 9) Or (Месяц_даты = 12) Then
    
          ' Берем факт из Date_ДД
          '
          ' Идентификатор ID_Rec для 1 месяца Q:
          ID_RecVar = CStr(CStr(In_OfficeNumber) + "-" + strMMYY(Date_begin_day_quarter(In_Date)) + "-" + In_Product_Code)
          ' Прибавляем факт 1 месяца
          resilt_getDataFrom_BASE_Workbook2_Fact = getDataFrom_BASE_Workbook2("Sales_Office", "Лист1", "ID_Rec", ID_RecVar, "Fact", 0)
          If resilt_getDataFrom_BASE_Workbook2_Fact <> "not found" Then
            Факт_Q_на_дату = Факт_Q_на_дату + resilt_getDataFrom_BASE_Workbook2_Fact ' getDataFrom_BASE_Workbook2("Sales_Office", "Лист1", "ID_Rec", ID_RecVar, "Fact", 0)
          End If
    
          ' Идентификатор ID_Rec для 2 месяца Q:
          ID_RecVar = CStr(CStr(In_OfficeNumber) + "-" + strMMYY(CDate("01." + CStr(Месяц_даты - 1) + "." + CStr(Год_даты))) + "-" + In_Product_Code)
          ' Прибавляем факт 2 месяца
          result_getDataFrom_BASE_Workbook2_Fact = getDataFrom_BASE_Workbook2("Sales_Office", "Лист1", "ID_Rec", ID_RecVar, "Fact", 0)
          If result_getDataFrom_BASE_Workbook2_Fact <> "not found" Then
            Факт_Q_на_дату = Факт_Q_на_дату + result_getDataFrom_BASE_Workbook2_Fact ' getDataFrom_BASE_Workbook2("Sales_Office", "Лист1", "ID_Rec", ID_RecVar, "Fact", 0)
          End If
          
        End If
  
      ' End If ' Если полученный факт на Дату <>0

    Else
      Факт_Q_на_дату = 0
    End If ' Проверяем результат на возможность обработки по записям месяца
 
    ' Проверяем результат квартала
    If result_Месяц_getDataFrom_BASE_Workbook2 = "not found" Then
      
      '  Идентификатор ID_Rec для квартальных планов:
      ID_RecVar = CStr(CStr(In_OfficeNumber) + "-" + strNQYY(In_Date) + "-" + In_Product_Code)
                        
      ' Текущие значения в месяце: Date_01 (N), Date_02 (O), Date_03 (P), Date_04 (Q), Date_05 Date_06 Date_07 Date_08 Date_09 Date_10 Date_11 Date_12 Date_13 Date_14 Date_15 Date_16 Date_17 Date_18 Date_19 Date_20 Date_21 Date_22 Date_23 Date_24 Date_25 Date_26 Date_27 Date_28 Date_29 Date_30 Date_31
      ' Номер месяца в квартале: 1-"", 2-"2", 3-"3"
      M_num = Nom_mes_quarter_str(In_Date)
      Column_DateN_DD = "Date" + M_num + "_" + Mid(In_Date, 1, 2)
      
      result_getDataFrom_BASE_Workbook2_Column_DateN_DD = getDataFrom_BASE_Workbook2("Sales_Office", "Лист1", "ID_Rec", ID_RecVar, Column_DateN_DD, 0)
      If result_getDataFrom_BASE_Workbook2_Column_DateN_DD <> "not found" Then
        Факт_Q_на_дату = result_getDataFrom_BASE_Workbook2_Column_DateN_DD ' getDataFrom_BASE_Workbook2("Sales_Office", "Лист1", "ID_Rec", ID_RecVar, Column_DateN_DD, 0)
      End If
    End If

  Else
    Факт_Q_на_дату = 0
  End If ' Если результат не пустой - возникает, когда добавляется новый продукт и по нему не возвращает "not found", а приходит Empty. Было 12.08.2021 при добавлении Портфеля АУМ из DB от 10.08.2021
  


End Function

' 89. Первый_понедельник_от_даты(Date)
Function Первый_понедельник_от_даты(In_Date) As Date
  
  ' Берем текущую дату
  текущая_дата = In_Date + 1
  
  ' Переменная контроля
  понедельник_найден = False
  
  Do While понедельник_найден = False
    
    ' Проверяем день недели
    If Weekday(текущая_дата, vbMonday) = 1 Then
      Первый_понедельник_от_даты = текущая_дата
      понедельник_найден = True
    Else
      ' Иначе берем следующий день
      текущая_дата = текущая_дата + 1
    End If
    
  Loop
  
End Function

' 90. Кавычки
Function кавычки() As String
  кавычки = Chr(34)
End Function

' 91. Обычный текст в ячейке
Sub Обычный_текст(In_Workbooks, In_Sheets, In_Row, In_Col)
  
  Workbooks(In_Workbooks).Sheets(In_Sheets).Cells(In_Row, In_Col).Font.Bold = False
  ' Workbooks(In_Workbooks).Sheets(In_Sheets).Cells(In_Row, In_Col).Borders(xlDiagonalDown).LineStyle = xlNone
  ' Workbooks(In_Workbooks).Sheets(In_Sheets).Cells(In_Row, In_Col).Borders(xlDiagonalUp).LineStyle = xlNone
  ' Workbooks(In_Workbooks).Sheets(In_Sheets).Cells(In_Row, In_Col).Borders(xlEdgeLeft).LineStyle = xlNone
  ' Workbooks(In_Workbooks).Sheets(In_Sheets).Cells(In_Row, In_Col).Borders(xlEdgeTop).LineStyle = xlNone
  ' Workbooks(In_Workbooks).Sheets(In_Sheets).Cells(In_Row, In_Col).Borders(xlEdgeBottom).LineStyle = xlNone
  ' Workbooks(In_Workbooks).Sheets(In_Sheets).Cells(In_Row, In_Col).Borders(xlEdgeRight).LineStyle = xlNone
  ' Workbooks(In_Workbooks).Sheets(In_Sheets).Cells(In_Row, In_Col).Borders(xlInsideVertical).LineStyle = xlNone
  ' Workbooks(In_Workbooks).Sheets(In_Sheets).Cells(In_Row, In_Col).Borders(xlInsideHorizontal).LineStyle = xlNone

End Sub


' 92. Полужирный текст в ячейке
Sub Полужирный_текст(In_Workbooks, In_Sheets, In_Row, In_Col)

  Workbooks(In_Workbooks).Sheets(In_Sheets).Cells(In_Row, In_Col).Font.Bold = True
  ' Workbooks(In_Workbooks).Sheets(In_Sheets).Cells(In_Row, In_Col).Borders(xlDiagonalDown).LineStyle = xlNone
  ' Workbooks(In_Workbooks).Sheets(In_Sheets).Cells(In_Row, In_Col).Borders(xlDiagonalUp).LineStyle = xlNone
  ' Workbooks(In_Workbooks).Sheets(In_Sheets).Cells(In_Row, In_Col).Borders(xlEdgeLeft).LineStyle = xlNone
  ' Workbooks(In_Workbooks).Sheets(In_Sheets).Cells(In_Row, In_Col).Borders(xlEdgeTop).LineStyle = xlNone
  ' Workbooks(In_Workbooks).Sheets(In_Sheets).Cells(In_Row, In_Col).Borders(xlEdgeBottom).LineStyle = xlNone
  ' Workbooks(In_Workbooks).Sheets(In_Sheets).Cells(In_Row, In_Col).Borders(xlEdgeRight).LineStyle = xlNone
  ' Workbooks(In_Workbooks).Sheets(In_Sheets).Cells(In_Row, In_Col).Borders(xlInsideVertical).LineStyle = xlNone
  ' Workbooks(In_Workbooks).Sheets(In_Sheets).Cells(In_Row, In_Col).Borders(xlInsideHorizontal).LineStyle = xlNone
  

End Sub

' 93. Курсив текст в ячейке
Sub Курсив_текст(In_Workbooks, In_Sheets, In_Row, In_Col)
  
    
  
End Sub


' 94. Убрать рамки
Sub Убрать_рамку(In_Workbooks, In_Sheets, In_Row, In_Col)
  
  ' Workbooks(In_Workbooks).Sheets(In_Sheets).Cells(In_Row, In_Col).Font.Bold = False
  Workbooks(In_Workbooks).Sheets(In_Sheets).Cells(In_Row, In_Col).Borders(xlDiagonalDown).LineStyle = xlNone
  Workbooks(In_Workbooks).Sheets(In_Sheets).Cells(In_Row, In_Col).Borders(xlDiagonalUp).LineStyle = xlNone
  Workbooks(In_Workbooks).Sheets(In_Sheets).Cells(In_Row, In_Col).Borders(xlEdgeLeft).LineStyle = xlNone
  Workbooks(In_Workbooks).Sheets(In_Sheets).Cells(In_Row, In_Col).Borders(xlEdgeTop).LineStyle = xlNone
  Workbooks(In_Workbooks).Sheets(In_Sheets).Cells(In_Row, In_Col).Borders(xlEdgeBottom).LineStyle = xlNone
  Workbooks(In_Workbooks).Sheets(In_Sheets).Cells(In_Row, In_Col).Borders(xlEdgeRight).LineStyle = xlNone
  Workbooks(In_Workbooks).Sheets(In_Sheets).Cells(In_Row, In_Col).Borders(xlInsideVertical).LineStyle = xlNone
  Workbooks(In_Workbooks).Sheets(In_Sheets).Cells(In_Row, In_Col).Borders(xlInsideHorizontal).LineStyle = xlNone

End Sub

' 95. Факт месяца по продукту из поля "Fact", из In_Date определяем только номер месяца и год
Function Факт_М(In_Date, In_OfficeNumber, In_Product_Code) As Double
  
  Факт_М = 0
    
  ' Определяем ID_Rec
  If In_Product_Code <> "ЗП" Then
    ' Идентификатор ID_Rec в формате 5-0119-КД
    ID_RecVar = CStr(In_OfficeNumber) + "-" + strMMYY(In_Date) + "-" + In_Product_Code
  Else
    ' для 3, 6, 9, 12 месяцев
    If (Month(In_Date) = 3) Or (Month(In_Date) = 6) Or (Month(In_Date) = 9) Or (Month(In_Date) = 12) Then
      ' Идентификатор ID_Rec в формате 1-1Q19-ЗП для Product_Code=ЗП
      ID_RecVar = CStr(In_OfficeNumber) + "-" + strNQYY(In_Date) + "-" + In_Product_Code
    Else
      ' Формируем как для месяца, результат будет 0
      ' Идентификатор ID_Rec в формате 5-0119-КД
      ID_RecVar = CStr(In_OfficeNumber) + "-" + strMMYY(In_Date) + "-" + In_Product_Code
    End If
  End If
  
  ' Выполняем поиск
  result_getDataFrom_BASE_Workbook2 = getDataFrom_BASE_Workbook2("Sales_Office", "Лист1", "ID_Rec", ID_RecVar, "Fact", 1)
  
  ' Анализируем результат поиска
  If result_getDataFrom_BASE_Workbook2 <> "not found" Then
    Факт_М = result_getDataFrom_BASE_Workbook2
  Else
    Факт_М = 0
  End If
  
End Function

' 96. План месяца по продукту из поля "Plan", из In_Date определяем только номер месяца и год
Function План_М(In_Date, In_OfficeNumber, In_Product_Code) As Double
  
  План_М = 0
  
  ' Идентификатор ID_Rec в формате 5-0119-КД
  ' ID_RecVar = CStr(In_OfficeNumber) + "-" + strMMYY(In_Date) + "-" + In_Product_Code
  
  ' Определяем ID_Rec
  If In_Product_Code <> "ЗП" Then
    ' Идентификатор ID_Rec в формате 5-0119-КД
    ID_RecVar = CStr(In_OfficeNumber) + "-" + strMMYY(In_Date) + "-" + In_Product_Code
  Else
    ' для 3, 6, 9, 12 месяцев
    If (Month(In_Date) = 3) Or (Month(In_Date) = 6) Or (Month(In_Date) = 9) Or (Month(In_Date) = 12) Then
      ' Идентификатор ID_Rec в формате 1-1Q19-ЗП для Product_Code=ЗП
      ID_RecVar = CStr(In_OfficeNumber) + "-" + strNQYY(In_Date) + "-" + In_Product_Code
    Else
      ' Формируем как для месяца, результат будет 0
      ' Идентификатор ID_Rec в формате 5-0119-КД
      ID_RecVar = CStr(In_OfficeNumber) + "-" + strMMYY(In_Date) + "-" + In_Product_Code
    End If
  End If
  
  
  ' Выполняем поиск
  result_getDataFrom_BASE_Workbook2 = getDataFrom_BASE_Workbook2("Sales_Office", "Лист1", "ID_Rec", ID_RecVar, "Plan", 1)
  
  ' Анализируем результат поиска
  If result_getDataFrom_BASE_Workbook2 <> "not found" Then
    План_М = result_getDataFrom_BASE_Workbook2
  Else
    План_М = 0
  End If
  
End Function

' 97. Проверка наличия данных
Function CheckData(In_Value) As Double

  ' Проверяем, если не пусто
  If (In_Value <> "") And (In_Value <> "not found") Then
    ' Результат = вх.параметру
    CheckData = In_Value
  Else
    ' Иначе 0
    CheckData = 0
  End If
  
End Function

' 98. Получение продаж за период - на основе Function Факт_Q_на_дату(In_OfficeNumber, In_Product_Code, In_Date) As Double - Факт за квартал по продукту из таблицы каталога BASE\"Sales_Office". Если факта нет в Date_DD, то возвращаем 0!
Function Продажи_Q_за_период(In_OfficeNumber, In_Product_Code, In_DateStart, In_DateEnd) As Double
  
  Факт_Q_на_дату_DateEnd = CheckData(Факт_Q_на_дату(In_OfficeNumber, In_Product_Code, In_DateEnd))
  
  Факт_Q_на_дату_DateStart = CheckData(Факт_Q_на_дату(In_OfficeNumber, In_Product_Code, In_DateStart))
  
  Продажи_Q_за_период = Факт_Q_на_дату_DateEnd - Факт_Q_на_дату_DateStart
  
End Function

' 99. Дата из "12.08.2021-19.08.2021"
Function Дата1(In_DateStr) As Date
  
  Дата1 = CDate(Mid(In_DateStr, 1, 10))
  
End Function

' 99. Дата из "12.08.2021-19.08.2021"
Function Дата2(In_DateStr) As Date
  
  Дата2 = CDate(Mid(In_DateStr, 12, 10))
  
End Function

