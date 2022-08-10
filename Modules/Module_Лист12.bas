Attribute VB_Name = "Module_Лист12"
' Лист 12

' Обработка ежедневного отчета от офиса
Sub getDataFromOffice_Лист12_4()
  
    
  ' Открываем файл
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
    ThisWorkbook.Sheets("Лист12").Activate

    ' Проверка формы отчета
    ' CheckFormatReportResult = CheckFormatReport(ReportName_String, "___", 6, Date)
    ' If CheckFormatReportResult = "OK" Then
    If True Then
      
      ' Открываем таблицу BASE\Custflow
      OpenBookInBase ("Custflow")

      rowCount = 1
      Do While InStr(Workbooks(ReportName_String).Sheets("Ежедневный отчет").Cells(rowCount, 1).Value, "Итого по РОО") = 0

        ' Продажи за:
        If InStr(Workbooks(ReportName_String).Sheets("Ежедневный отчет").Cells(rowCount, 1).Value, "Продажи за:") <> 0 Then
           Продажи_за = CDate(Mid(Workbooks(ReportName_String).Sheets("Ежедневный отчет").Cells(rowCount, 1).Value, 13, 10))
        End If
   
        ' Если это "ОО «" и не пусто в строке клиентов
        If (InStr(Workbooks(ReportName_String).Sheets("Ежедневный отчет").Cells(rowCount, 1).Value, "ОО «") <> 0) And (Len(Trim(Workbooks(ReportName_String).Sheets("Ежедневный отчет").Cells(rowCount, 2).Value)) <> 0) Then
          
          ' ID_Rec
          ID_RecVar = strDDMMYYYY(Продажи_за) + "-" + cityOfficeName(Workbooks(ReportName_String).Sheets("Ежедневный отчет").Cells(rowCount, 1).Value)
          
          ' Заносим данные в таблицу BASE\Custflow: ID_Rec (Date-Офис), Date, Офис, Клиентопоток, Заявки_ДК, Заявления_ПФР, Подключенные_НС, Консультации_НСЖ, Продажи_НСЖ ИБ_новые_реактивация
          Call InsertRecordInBook("Custflow", "Лист1", "ID_Rec", ID_RecVar, _
                                          "ID_Rec", ID_RecVar, _
                                            "Date", Продажи_за, _
                                              "Офис", Workbooks(ReportName_String).Sheets("Ежедневный отчет").Cells(rowCount, 1).Value, _
                                                "Клиентопоток", Workbooks(ReportName_String).Sheets("Ежедневный отчет").Cells(rowCount, 2).Value, _
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

        End If
   
        ' Обрабатываем присланный отчет
        ' Следующая запись
        rowCount = rowCount + 1
        ' Application.StatusBar = officeNameInReport + ": " + CStr(rowCount) + "..."
        ' DoEventsInterval (rowCount)
          
      Loop


      ' Закрываем базу BASE\Tasks
      CloseBook ("Custflow")
    
      ' Переменная завершения обработки
      finishProcess = True
    Else
      ' Сообщение о неверном формате отчета или даты
      MsgBox ("Проверьте отчет: " + CheckFormatReportResult + "!")
    End If ' Проверка формы отчета

    ' Закрываем файл с отчетом без сохранения изменений (параметр SaveChanges:=False)
    Workbooks(Dir(FileName)).Close SaveChanges:=False
    
    ' Переходим в ячейку M2
    ' ThisWorkbook.Sheets("Лист12").Range("L78").Select

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


' Напоминание: прислать отчет по отработке вкладов
Sub toSend_Отправить_отчет_по_клиентопотоку()

Dim Range_str, Адрес_Ln_кому, Адрес_Ln_копия As String
Dim Range_Row, Range_Column, i As Byte
Dim темаПисьма, текстПисьма, hashTag As String

  
    ' Запрос
    If MsgBox("Отправить себе шаблон отчета по клиентопотоку?", vbYesNo) = vbYes Then
    
      If MsgBox("Дата отчета " + CStr(ThisWorkbook.Sheets("Лист12").Range("H2").Value) + "?", vbYesNo) = vbYes Then
        
      End If
    
      ' Отправка сообщения
      ' Адрес_Ln_кому = getFromAddrBook("УДО2,УДО3,УДО4,УДО5,НОРПиКО1,НОРПиКО2,НОРПиКО3,НОРПиКО4,НОРПиКО5", 2)
      ' Адрес_Ln_копия = getFromAddrBook("РД", 2)
      
      ' Тема письма - Тема:
      темаПисьма = "Отчет по клиентопотоку"
      ' hashTag - Хэштэг:
      hashTag = "#клиентопоток"
      ' Текст письма
      текстПисьма = "" + Chr(13)
      текстПисьма = текстПисьма + "Тюмень " + strDDMM(ThisWorkbook.Sheets("Лист12").Range("H2").Value) + ":" + Chr(13)
      текстПисьма = текстПисьма + "Клиентопоток - " + CStr(ThisWorkbook.Sheets("Лист12").Cells(11, 3).Value) + Chr(13)
      текстПисьма = текстПисьма + "Заявки ДК - " + CStr(ThisWorkbook.Sheets("Лист12").Cells(11, 4).Value) + Chr(13)
      текстПисьма = текстПисьма + "Заявления ПФР - " + CStr(ThisWorkbook.Sheets("Лист12").Cells(11, 5).Value) + Chr(13)
      текстПисьма = текстПисьма + "Подключенные НС - " + CStr(ThisWorkbook.Sheets("Лист12").Cells(11, 6).Value) + Chr(13)
      текстПисьма = текстПисьма + "Консультации НСЖ - " + CStr(ThisWorkbook.Sheets("Лист12").Cells(11, 7).Value) + Chr(13)
      текстПисьма = текстПисьма + "Продажи НСЖ - " + CStr(ThisWorkbook.Sheets("Лист12").Cells(11, 8).Value) + Chr(13)
      текстПисьма = текстПисьма + "ИБ (новые + реакт.) - " + CStr(ThisWorkbook.Sheets("Лист12").Cells(11, 9).Value) + Chr(13)
      текстПисьма = текстПисьма + "Коробки - " + CStr(ThisWorkbook.Sheets("Лист12").Cells(11, 10).Value) + Chr(13)
      ' Если за день есть Продажа ИСЖ, то дописываем
      If ThisWorkbook.Sheets("Лист12").Cells(11, 11).Value <> 0 Then
        текстПисьма = текстПисьма + "Продажи ИСЖ - " + CStr(ThisWorkbook.Sheets("Лист12").Cells(11, 11).Value) + Chr(13)
      End If
      '
      текстПисьма = текстПисьма + "" + Chr(13)
      текстПисьма = текстПисьма + "" + Chr(13)
      ' Визитка (подпись С Ув., )
      текстПисьма = текстПисьма + ПодписьВПисьме()
      ' Хэштег
      текстПисьма = текстПисьма + createBlankStr(35) + hashTag
      
      ' Вызов
      ' Call send_Lotus_Notes(темаПисьма, Адрес_Ln_кому, Адрес_Ln_копия, текстПисьма, "")
      
      ' Вызов с отправкой себе в скрытой копии
      Call send_Lotus_Notes2(темаПисьма, "Sergey Fedorovich Proschaev/Tyumen/PSBank/Ru", "", "", текстПисьма, "")
  
      ' Сообщение
      MsgBox ("Письмо отправлено!")
      
    End If
  
End Sub

' Очистить форму с даными за день
Sub clearGrid_Sheet12()
Dim i, j, new_j As Byte

    ' Запрос
    If MsgBox("Очистить форму?", vbYesNo) = vbYes Then
    
      For i = 17 To 21
        
        For j = 3 To 11
          
          ' 1. Очистка - Форма 12.2
          ThisWorkbook.Sheets("Лист12").Cells(i, j).Value = 0
          
          ' 2. Очистка - Форма 12.3
          ' Смещения по Y
          Select Case j
            Case 3 ' Клиентопоток
              new_j = 3
            Case 4 ' Заявки ДК
              new_j = 4
            Case 5 ' Заявления ПФР
              new_j = 8
            Case 6 ' Подключенные НС
              new_j = 6
            Case 7 ' Консультации НСЖ
              new_j = 9
            Case 8 ' Продажи НСЖ
              new_j = 10
            Case 9 ' ИБ (новые + реакт.)
              new_j = 11
            Case 10 ' Коробки
              new_j = 12
            Case 11 ' Продажи ИСЖ
              new_j = 13
          
          End Select

          ThisWorkbook.Sheets("Лист12").Cells(i + 11, new_j).Value = 0
        Next j
        
      Next i
      
      ' Обновление итоговой таблицы
      Call reFreshDataFromDay_Sheet12
      
      ' Сообщение
      MsgBox ("Форма очищена!")
  
    End If
    
End Sub

' Внести данные за день
Sub insertDataFromDay_Sheet12()
Dim i, j As Byte

    ' Проходим циклом и суммируем нарастающими итогами
    
    ' Отправить письмо?
    Call toSend_Отправить_отчет_по_клиентопотоку
    
    ' Запрос
    If MsgBox("Внести данные за " + CStr(ThisWorkbook.Sheets("Лист12").Range("H2").Value) + "?", vbYesNo) = vbYes Then
    
      ' Открываем таблицу BASE\Custflow
      OpenBookInBase ("Custflow")
    
      ' Итерация №1: прибавляем к данным месяца данные дня
      For i = 6 To 10
        
        ' Суммируем день с месяцем
        For j = 3 To 11
          
          ' Заносим данные
          ThisWorkbook.Sheets("Лист12").Cells(i + 11, j).Value = ThisWorkbook.Sheets("Лист12").Cells(i + 11, j).Value + ThisWorkbook.Sheets("Лист12").Cells(i, j).Value
                    
        Next j
        
        ' Заносим данные в таблицу BASE\Custflow: ID_Rec (Date-Офис), Date, Офис, Клиентопоток, Заявки_ДК, Заявления_ПФР, Подключенные_НС, Консультации_НСЖ, Продажи_НСЖ ИБ_новые_реактивация
        Call InsertRecordInBook("Custflow", "Лист1", "ID_Rec", strDDMMYYYY(ThisWorkbook.Sheets("Лист12").Range("H2").Value) + "-" + cityOfficeName(ThisWorkbook.Sheets("Лист12").Cells(i, 2).Value), _
                                          "ID_Rec", strDDMMYYYY(ThisWorkbook.Sheets("Лист12").Range("H2").Value) + "-" + cityOfficeName(ThisWorkbook.Sheets("Лист12").Cells(i, 2).Value), _
                                            "Date", ThisWorkbook.Sheets("Лист12").Range("H2").Value, _
                                              "Офис", ThisWorkbook.Sheets("Лист12").Cells(i, 2).Value, _
                                                "Клиентопоток", ThisWorkbook.Sheets("Лист12").Cells(i, 3).Value, _
                                                  "Заявки_ДК", ThisWorkbook.Sheets("Лист12").Cells(i, 4).Value, _
                                                    "Заявления_ПФР", ThisWorkbook.Sheets("Лист12").Cells(i, 5).Value, _
                                                      "Подключенные_НС", ThisWorkbook.Sheets("Лист12").Cells(i, 6).Value, _
                                                        "Консультации_НСЖ", ThisWorkbook.Sheets("Лист12").Cells(i, 7).Value, _
                                                          "Продажи_НСЖ", ThisWorkbook.Sheets("Лист12").Cells(i, 8).Value, _
                                                            "ИБ_новые_реактивация", ThisWorkbook.Sheets("Лист12").Cells(i, 9).Value, _
                                                               "Коробки", ThisWorkbook.Sheets("Лист12").Cells(i, 10).Value, _
                                                                 "Продажи_ИСЖ", ThisWorkbook.Sheets("Лист12").Cells(i, 11).Value, _
                                                                   "", "", _
                                                                     "", "", _
                                                                       "", "", _
                                                                         "", "", _
                                                                           "", "", _
                                                                             "", "", _
                                                                               "", "", _
                                                                                 "", "")
                                                                                  

        
      Next i
      
      ' Итерация №2: очищаем форму
      For i = 6 To 10
        
        For j = 3 To 11
          
          ' Обнуляем ячейку ввода - на период теста не обнуляем
          ThisWorkbook.Sheets("Лист12").Cells(i, j).Value = 0
          
        Next j
        
      Next i
      
      ' Закрываем базу BASE\Tasks
      CloseBook ("Custflow")

      ' Обновляем накопительным за месяц
      Call reFreshDataFromDay_Sheet12
      
      ' Зачеркнуть на листе 0 "Клиентопоток и продажи"
      Call ЗачеркиваемТекстВячейке("Лист0", RangeByValue(ThisWorkbook.Name, "Лист0", "Клиентопоток и продажи", 100, 100))

      ' Сообщение
      MsgBox ("Данные за " + CStr(ThisWorkbook.Sheets("Лист12").Range("H2").Value) + " успешно внесены!")
  
    End If
  
End Sub

' Обновить данные
Sub reFreshDataFromDay_Sheet12()
Dim i, j, new_j As Byte

  ' Запрос
  If MsgBox("Обновить данные?", vbYesNo) = vbYes Then
    
    ' Очищаем таблицу с C27 до K32
    Call clearСontents2(ThisWorkbook.Name, "Лист12", "C28", "M32")
    
    ' в B24: "Оперативная бизнес-справка клиентопоток и кросс-продажи за май 2020 г. (на 26.05.2020)"
    ThisWorkbook.Sheets("Лист12").Cells(24, 2).Value = "Оперативная бизнес-справка клиентопоток и кросс-продажи за " + ИмяМесяца(ThisWorkbook.Sheets("Лист12").Cells(2, 8).Value) + " (на " + CStr(ThisWorkbook.Sheets("Лист12").Cells(2, 8).Value) + ")"
    
    ' Заносим данные
      ' Итерация №1: прибавляем к данным месяца данные дня
      For i = 28 To 32
        
        ' Суммируем день с месяцем
        For j = 3 To 11
          
          ' Смещения по Y
          Select Case j
            Case 3 ' Клиентопоток
              new_j = 3
            Case 4 ' Заявки ДК
              new_j = 4
            Case 5 ' Заявления ПФР
              new_j = 8
            Case 6 ' Подключенные НС
              new_j = 6
            Case 7 ' Консультации НСЖ
              new_j = 9
            Case 8 ' Продажи НСЖ
              new_j = 10
            Case 9 ' ИБ (новые + реакт.)
              new_j = 11
            Case 10 ' Коробки
              new_j = 12
            Case 11 ' Продажи ИСЖ
              new_j = 13
              
          End Select

          ' Заносим данные
          ThisWorkbook.Sheets("Лист12").Cells(i, new_j).Value = ThisWorkbook.Sheets("Лист12").Cells(i - 11, j).Value
          ThisWorkbook.Sheets("Лист12").Cells(i, new_j).HorizontalAlignment = xlRight
                    
        Next j
      
      Next i
         
    ' Делаем расчет метрик
    For i = 28 To 32
      
      ' Заявки ДК
      ThisWorkbook.Sheets("Лист12").Cells(i, 5).Value = РассчетДоли(ThisWorkbook.Sheets("Лист12").Cells(i, 3).Value, ThisWorkbook.Sheets("Лист12").Cells(i, 4).Value, 3)
      ThisWorkbook.Sheets("Лист12").Cells(i, 5).NumberFormat = cellsNumberFormat(ThisWorkbook.Sheets("Лист12").Cells(i, 5).Value)
      ThisWorkbook.Sheets("Лист12").Cells(i, 5).HorizontalAlignment = xlRight
      ' Окраска ячейки СФЕТОФОР: если
      Call Full_Color_RangeII("Лист12", i, 5, (РассчетДоли(ThisWorkbook.Sheets("Лист12").Cells(i, 3).Value, ThisWorkbook.Sheets("Лист12").Cells(i, 4).Value, 3) * 100), 15)
      
      ' НС
      ThisWorkbook.Sheets("Лист12").Cells(i, 7).Value = РассчетДоли(ThisWorkbook.Sheets("Лист12").Cells(i, 3).Value, ThisWorkbook.Sheets("Лист12").Cells(i, 6).Value, 3)
      ThisWorkbook.Sheets("Лист12").Cells(i, 7).NumberFormat = cellsNumberFormat(ThisWorkbook.Sheets("Лист12").Cells(i, 7).Value)
      ThisWorkbook.Sheets("Лист12").Cells(i, 7).HorizontalAlignment = xlRight
      ' Окраска ячейки СФЕТОФОР: если
      Call Full_Color_RangeII("Лист12", i, 7, (РассчетДоли(ThisWorkbook.Sheets("Лист12").Cells(i, 3).Value, ThisWorkbook.Sheets("Лист12").Cells(i, 6).Value, 3) * 100), 15)
      
      ' ПФР нули
      ' Окраска ячейки СФЕТОФОР: если
      If ThisWorkbook.Sheets("Лист12").Cells(i, 8).Value = 0 Then
        Call Full_Color_RangeII("Лист12", i, 8, 0, 100)
      End If
      
      ' Коробки нули
      ' Окраска ячейки СФЕТОФОР: если
      If ThisWorkbook.Sheets("Лист12").Cells(i, 12).Value = 0 Then
        Call Full_Color_RangeII("Лист12", i, 12, 0, 100)
      End If
      
    Next i
         
  End If

End Sub

' Офистить форму Форма 12.4
Sub clearForm2_Лист12_4()
  ' Запрос
  If MsgBox("Очистить форму?", vbYesNo) = vbYes Then
    
    Call clearСontents2(ThisWorkbook.Name, "Лист12", "C44", "M48")
    
    ' Сообщение
    MsgBox ("Форма очищена!")
    
  End If
End Sub

' Офистить форму Форма 12.5
Sub clearForm2_Лист12_5()
  ' Запрос
  ' If MsgBox("Очистить форму?", vbYesNo) = vbYes Then
    
    Call clearСontents2(ThisWorkbook.Name, "Лист12", "C55", "M59")
    
    ' Сообщение
    ' MsgBox ("Форма очищена!")
    
  ' End If
End Sub



' Отправить отчет
Sub sendReport_Лист12()
  ' Запрос
  If MsgBox("Отправить отчет?", vbYesNo) = vbYes Then
        
    
    ' Сообщение
    MsgBox ("Отчет отправлен!")
        
  End If
End Sub

