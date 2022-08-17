Attribute VB_Name = "Module_ДР"
' *** Лист ДР ***

' *** Глобальные переменные ***
' Public numStr_Лист8 As Integer
' ***                       ***

' Отправка уведомлений о ДР клиентов из BASE\Birthdays
Sub Отправка_уведомлений_ДР()

  ' Строка статуса
  Application.StatusBar = "ДР: отправка уведомлений..."

  ' Очистить таблицу на Листе "ДР"
  Call clearСontents2(ThisWorkbook.Name, "ДР", "A9", "I28")

  ' Открыть BASE\Birthdays
  OpenBookInBase ("Birthdays")
  
  ' "ФИО"
  column_ДР_ФИО = ColumnByValue("Birthdays", "Лист1", "ФИО", 100, 100)

  ' "Дата рождения"
  column_ДР_Дата_рождения = ColumnByValue("Birthdays", "Лист1", "Дата рождения", 100, 100)
  
  ' "Организация"
  column_ДР_Организация = ColumnByValue("Birthdays", "Лист1", "Организация", 100, 100)
  
  ' "Должность"
  column_ДР_Должность = ColumnByValue("Birthdays", "Лист1", "Должность", 100, 100)
  
  ' "Примечание"
  column_ДР_Примечание = ColumnByValue("Birthdays", "Лист1", "Примечание", 100, 100)
  
  ' "ССП"
  column_ДР_ССП = ColumnByValue("Birthdays", "Лист1", "ССП", 100, 100)
  
  ' Категория "Категория поздравления"
  column_ДР_Категория_поздравления = ColumnByValue("Birthdays", "Лист1", "Категория поздравления", 100, 100)
  
  ' "1я - поздравление лично X.Y. Xyz (самые ценные для партнерства)"
  ' "2я - поздравление руководителем подразделения"
  ' "3я - подравление путем направления открытки (электронной/бумажной)"
  
  ' "Дата уведомления"
  column_ДР_Дата_уведомления = ColumnByValue("Birthdays", "Лист1", "Дата уведомления", 100, 100)
  
  
  ' Строка №
  row_N = rowByValue("Birthdays", "Лист1", "№", 100, 100)
    
  ' Отправить уведомление по клиентам, у которых ДР сегодня
  Номер_строки_Лист_ДР = 8
  Номер_позиции_Лист_ДР = 0
  rowCount = row_N + 2
  Do While Workbooks("Birthdays").Sheets("Лист1").Cells(rowCount, column_ДР_ФИО).Value <> ""
 
    Отправляем_уведомление = False
 
    ' Проверяем Дату ДР - она должна быть не пустая
    If Workbooks("Birthdays").Sheets("Лист1").Cells(rowCount, column_ДР_Дата_рождения).Value <> "" Then
    
      ' Проверяем Дату ДР - она должна быть или сегодня или не позднее -2 дня
      If Проверка_Даты_рождения_для_уведомления(CDate(Workbooks("Birthdays").Sheets("Лист1").Cells(rowCount, column_ДР_Дата_рождения).Value)) = True Then
        
        
        ' Проверяем - дату уведомления стоит?
        If Workbooks("Birthdays").Sheets("Лист1").Cells(rowCount, column_ДР_Дата_уведомления).Value <> "" Then
        
          ' Проверяем - отправлялось ли уведомление уже в этом году?
          If Year(Workbooks("Birthdays").Sheets("Лист1").Cells(rowCount, column_ДР_Дата_уведомления).Value) <> Year(Date) Then
            Отправляем_уведомление = True
          End If
        
        Else
          
          Отправляем_уведомление = True
          
        End If ' Проверяем - дату уведомления стоит?
        
      End If ' Проверяем Дату ДР
                
      ' Отправляем уведомление?
      If Отправляем_уведомление = True Then
        
            Номер_строки_Лист_ДР = Номер_строки_Лист_ДР + 1
        
            ' 1) Заносим Клиента на лист ДР
            ' N
            Номер_позиции_Лист_ДР = Номер_позиции_Лист_ДР + 1
            ThisWorkbook.Sheets("ДР").Cells(Номер_строки_Лист_ДР, 1).Value = Номер_позиции_Лист_ДР
            ' ФИО
            ThisWorkbook.Sheets("ДР").Cells(Номер_строки_Лист_ДР, 2).Value = Workbooks("Birthdays").Sheets("Лист1").Cells(rowCount, column_ДР_ФИО).Value
            ' ДР
            ThisWorkbook.Sheets("ДР").Cells(Номер_строки_Лист_ДР, 3).Value = Workbooks("Birthdays").Sheets("Лист1").Cells(rowCount, column_ДР_Дата_рождения).Value
            ' Организация
            ThisWorkbook.Sheets("ДР").Cells(Номер_строки_Лист_ДР, 4).Value = Workbooks("Birthdays").Sheets("Лист1").Cells(rowCount, column_ДР_Организация).Value
            ' Должность
            ThisWorkbook.Sheets("ДР").Cells(Номер_строки_Лист_ДР, 5).Value = Workbooks("Birthdays").Sheets("Лист1").Cells(rowCount, column_ДР_Должность).Value
            ' Примечание
            ThisWorkbook.Sheets("ДР").Cells(Номер_строки_Лист_ДР, 6).Value = Workbooks("Birthdays").Sheets("Лист1").Cells(rowCount, column_ДР_Примечание).Value
            ' ССП
            ThisWorkbook.Sheets("ДР").Cells(Номер_строки_Лист_ДР, 7).Value = Workbooks("Birthdays").Sheets("Лист1").Cells(rowCount, column_ДР_ССП).Value
            ' Категория
            ThisWorkbook.Sheets("ДР").Cells(Номер_строки_Лист_ДР, 8).Value = Workbooks("Birthdays").Sheets("Лист1").Cells(rowCount, column_ДР_Категория_поздравления).Value
            ' Дата уведомления
            ThisWorkbook.Sheets("ДР").Cells(Номер_строки_Лист_ДР, 9).Value = Date
        
            ' 2) Отправляем уведомление
            Call Отправка_Лист_ДР("Birthdays", "Лист1", rowCount)
        
            ' 3) Делаем отметку в "Дата уведомления"
            ' column_ДР_Дата_уведомления = ColumnByValue("Birthdays", "Лист1", "Дата уведомления", 100, 100)
            Workbooks("Birthdays").Sheets("Лист1").Cells(rowCount, column_ДР_Дата_уведомления).Value = Date

        
      End If ' Отправляем уведомление?
        
    End If
    
    ' Следующая запись
    Application.StatusBar = "ДР: отправка уведомлений " + CStr(rowCount) + "..."
    rowCount = rowCount + 1
    DoEventsInterval (rowCount)
    
  Loop
    
  ' Закрыть BASE\Birthdays
  CloseBook ("Birthdays")

  ' Строка статуса
  Application.StatusBar = ""


End Sub


' Отправка письма: отправляю шаблон самому себе для последующей отправки в сеть письма на его основе:
Sub Отправка_Лист_ДР(In_Workbooks, In_Sheets, In_Row)
Dim темаПисьма, текстПисьма, hashTag, attachmentFile As String
Dim i As Byte
  
  ' Строка статуса
  Application.StatusBar = "Отправка письма..."
  
    ' Workbooks("Birthdays").Sheets("Лист1").Cells(rowCount, column_ДР_Дата_уведомления).Value = Date
  
    
    ' dateDB_Лист8 = CDate(Mid(ThisWorkbook.Sheets("Лист8").Range("B5").Value, 52, 10))
   
    ' Тема письма - Тема:
    ' темаПисьма = ThisWorkbook.Sheets("Лист8").Cells(RowByValue(ThisWorkbook.Name, "Лист8", "Тема:", 100, 100), ColumnByValue(ThisWorkbook.Name, "Лист8", "Тема:", 100, 100) + 1).Value
    темаПисьма = "Дни рождения клиентов " + Фамилия_и_Имя(Workbooks(In_Workbooks).Sheets(In_Sheets).Cells(In_Row, 2).Value, 3)

    ' hashTag - Хэштэг:
    hashTag = "#др #дни_рождения #t227211055"

    ' Файл-вложение (!!!)
    attachmentFile = "" ' ThisWorkbook.Sheets("Лист8").Range("S3").Value
    
    ' Адресат письма
    Select Case Workbooks(In_Workbooks).Sheets(In_Sheets).Cells(In_Row, 7).Value
      
      Case "ОКП" ' ОО «Тюменский»
        Адресат_письма = getFromAddrBook("НОКП, РРКК", 2)
      
      Case "ОРПиКО Тюмень" ' ОО «Тюменский»
        Адресат_письма = getFromAddrBook("НОРПиКО1, ПМ", 2)
        
      Case "ИЦ" ' ОО «Тюменский»
        Адресат_письма = getFromAddrBook("РИЦ", 2)
        
      Case "Сургут" ' ОО «Сургутский»
        Аддресат_письма = getFromAddrBook("УДО2, НОРПиКО2", 2)
        
      Case "Нижневартовск" ' ОО «Нижневартовский»
        Аддресат_письма = getFromAddrBook("УДО3, НОРПиКО3", 2)
        
      Case "Новый Уренгой" ' ОО «Новоуренгойский»
        Аддресат_письма = getFromAddrBook("УДО4, НОРПиКО4", 2)
        
      Case "Тарко-Сале" ' ОО «Тарко-Сале»
        Аддресат_письма = getFromAddrBook("УДО5, НОРПиКО5", 2)
            
    End Select

    
    ' Текст письма
    текстПисьма = "" + Chr(13)
    текстПисьма = текстПисьма + "" + Адресат_письма + Chr(13) + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "" + getFromAddrBook("РД", 2) + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "Уважаемые руководители," + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "Направляю уведомление о Дне рождения клиента" + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + ">> " '  + Chr(13)
    
     
    ' ФИО
    текстПисьма = текстПисьма + Workbooks(In_Workbooks).Sheets(In_Sheets).Cells(In_Row, 2).Value + ", "
    ' ДР
    текстПисьма = текстПисьма + CStr(Workbooks(In_Workbooks).Sheets(In_Sheets).Cells(In_Row, 3).Value) + ", "
    ' Организация
    текстПисьма = текстПисьма + Workbooks(In_Workbooks).Sheets(In_Sheets).Cells(In_Row, 4).Value + ", "
    ' Должность
    текстПисьма = текстПисьма + Workbooks(In_Workbooks).Sheets(In_Sheets).Cells(In_Row, 5).Value + ", "
    ' Примечание
    текстПисьма = текстПисьма + Workbooks(In_Workbooks).Sheets(In_Sheets).Cells(In_Row, 6).Value + ", "
    ' ССП
    текстПисьма = текстПисьма + "Подразделение: " + Workbooks(In_Workbooks).Sheets(In_Sheets).Cells(In_Row, 7).Value + ", "
    ' Категория
    текстПисьма = текстПисьма + "Категория: " + Workbooks(In_Workbooks).Sheets(In_Sheets).Cells(In_Row, 8).Value + " "

    текстПисьма = текстПисьма + "" + Chr(13)
    ' Визитка (подпись С Ув., )
    текстПисьма = текстПисьма + ПодписьВПисьме()
    ' Хэштег
    текстПисьма = текстПисьма + createBlankStr(30) + hashTag
    ' Вызов
    Call send_Lotus_Notes(темаПисьма, "Sergey Fedorovich Proschaev/Tyumen/PSBank/Ru", "Sergey Fedorovich Proschaev/Tyumen/PSBank/Ru", текстПисьма, attachmentFile)
  
    ' Сообщение
    ' MsgBox ("Письмо отправлено!")
     
    ' Строка статуса
    Application.StatusBar = ""
     
  
End Sub

' Проверяем Дату ДР - она должна быть или сегодня или не позднее -2 дня
Function Проверка_Даты_рождения_для_уведомления(In_Date) As Boolean

  Проверка_Даты_рождения_для_уведомления = False
  
  ' Декомпозируем In_Date
  ' Месяц
  Месяц_ДР = Month(In_Date)
  ' Год
  Год_ДР = Year(In_Date)
  ' Число
  Число_ДР = CByte(Mid(CStr(In_Date), 1, 2))

  ' Собираем в дату текущего года
  Дата_ДР_в_этом_году = CDate(CStr(Число_ДР) + "." + CStr(Месяц_ДР) + "." + CStr(Year(Date)))
    
  ' Проверяем
  If (Дата_ДР_в_этом_году = Date) Or (Дата_ДР_в_этом_году >= (Date - 2)) And (Дата_ДР_в_этом_году < Date) Then
    
      Проверка_Даты_рождения_для_уведомления = True
   
  End If
  
End Function