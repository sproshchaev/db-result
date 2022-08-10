Attribute VB_Name = "Module_СМОТ"
' *** Лист 17 (СМОТ) ***

' *** Глобальные переменные ***
' Public numStr_Лист17 As Integer
' ***                       ***

' Разделение файла СМОТ на несколько файлов для отправки на согласование
Sub Разделение_файла_СМОТ()
' Описание переменных
Dim ReportName_String, officeNameInReport, CheckFormatReportResult As String
Dim i, rowCount As Integer
Dim finishProcess As Boolean
    
  ' Открыть файл с отчетом
  FileName = Application.GetOpenFilename("Excel Files (*.xls), *.xls", , "Открытие файла с отчетом")

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
    ThisWorkbook.Sheets("СМОТ").Activate

    ' Проверка формы отчета
    ' CheckFormatReportResult = CheckFormatReport(ReportName_String, "___", 6, Date)
    ' If CheckFormatReportResult = "OK" Then
    If True Then
      
      ' Очистка таблицы
      Call clearСontents2(ThisWorkbook.Name, "СМОТ", "A7", "I50")
      записьНаЛистСМОТ = 6
      номер = 0

      ' Имя файла без расширения
      ИмяФайлаБезРасширения = Mid(ReportName_String, 1, InStr(ReportName_String, ".") - 1)

      ' Тема
      ThisWorkbook.Sheets("СМОТ").Range("P2").Value = "СМОТ на согласование " + ИмяФайлаБезРасширения

      ' Находим столбец "Симв."
      Column_Симв = ColumnByValue(ReportName_String, "Лист1", "Симв.", 100, 100)
      ' Может
      Column_Фамилия_Имя_Отчество = ColumnByValue(ReportName_String, "Лист1", "Фамилия Имя Отчество", 100, 100)
      If Column_Фамилия_Имя_Отчество = 0 Then
        Column_Фамилия_Имя_Отчество = ColumnByValue(ReportName_String, "Лист1", "ФИО", 100, 100) ' ФИО
      End If

      rowCount = 1
      Do While Not IsEmpty(Workbooks(ReportName_String).Sheets("Лист1").Cells(rowCount, 1).Value)
        
        ' Если это текущий сотрудник
        ' If InStr(Workbooks(ReportName_String).Sheets("Лист1").Cells(rowCount, Column_Симв).Value, "Китог") <> 0 Then
        If (InStr(Workbooks(ReportName_String).Sheets("Лист1").Cells(rowCount, Column_Симв).Value, "Китог") <> 0) Or ((InStr(Workbooks(ReportName_String).Sheets("Лист1").Cells(rowCount, Column_Симв).Value, "Итог") <> 0)) Then
            
          ' ФИО сотрудника
          nameStaff = Workbooks(ReportName_String).Sheets("Лист1").Cells(rowCount, Column_Фамилия_Имя_Отчество).Value
          fullNameStaff = (Workbooks(ReportName_String).Sheets("Лист1").Cells(rowCount, Column_Фамилия_Имя_Отчество).Value)
          ' Нумерация на листе СМОТ
          записьНаЛистСМОТ = записьНаЛистСМОТ + 1
          номер = номер + 1
          ' Вывод данных на Лист
          ThisWorkbook.Sheets("СМОТ").Cells(записьНаЛистСМОТ, 1).Value = CStr(номер)
          ThisWorkbook.Sheets("СМОТ").Cells(записьНаЛистСМОТ, 2).Value = Фамилия_и_Имя(nameStaff, 3)
            
          ' Статус обработки
          Application.StatusBar = Фамилия_и_Имя(nameStaff, 3) + "..."
            
          ' Сохраняем новый файл
          FileNewVar = ThisWorkbook.Path + "\Out\" + ИмяФайлаБезРасширения + "_" + Фамилия_и_Имя(Workbooks(ReportName_String).Sheets("Лист1").Cells(rowCount, Column_Фамилия_Имя_Отчество).Value, 3) + ".xls"
          Workbooks(ReportName_String).SaveCopyAs FileName:=FileNewVar
          shortFileNewVar = Dir(FileNewVar)
          ' Открываем этот файл (UpdateLinks:=0)
          Workbooks.Open FileNewVar, 0

          ' Строка с именами файлов для архивирования
          ' strFileNewVar_Office = strFileNewVar_Office + ThisWorkbook.Path + "\Out\" + FileNewVar + " "
          
          ' Обработка файла
          ' Проходим по файлу с самого начала (со 2-ой) и удаляем
          rowCount2 = 2
          Do While Not IsEmpty(Workbooks(shortFileNewVar).Sheets("Лист1").Cells(rowCount2, 1).Value)
            
            ' Если это не ФИО текущего сотрудника - то удаляем запись
            If InStr(Workbooks(shortFileNewVar).Sheets("Лист1").Cells(rowCount2, Column_Фамилия_Имя_Отчество).Value, nameStaff) = 0 Then
              
              ' Удаляем текущую запись в файле - это работает, но окно нельзя закрывать
              ' Workbooks(shortFileNewVar).Sheets("Лист1").Rows(CStr(rowCount2) + ":" + CStr(rowCount2)).Select
              ' Selection.Delete Shift:=xlUp
              
              ' Вариант 2
              ' Workbooks(shortFileNewVar).Sheets("Лист1").Rows(CStr(rowCount2) + ":" + CStr(rowCount2)).Select
              Workbooks(shortFileNewVar).Sheets("Лист1").Rows(CStr(rowCount2) + ":" + CStr(rowCount2)).Delete Shift:=xlUp
              
              ' Если удадили запись, то стоим на ней же
              rowCount2 = rowCount2 - 1
              
            End If
            
            
            ' Следующая запись
            rowCount2 = rowCount2 + 1
            DoEventsInterval (rowCount)
          Loop
          
          ' Устанавливаем курсор на первую запись
          ' Workbooks(shortFileNewVar).Sheets("Лист1").Range("C1").Select
          
          ' Закрываем файл
          Workbooks(shortFileNewVar).Close SaveChanges:=True
          
          ' Отправка файла в почте отдельными файлами если это УДО, НОРПиКО, НОКП, РРКК
          Call Отправка_Lotus_Notes_ЛистСМОТ(fullNameStaff, FileNewVar)
          
          ' Переходим на окно СМОТ
          ThisWorkbook.Sheets("СМОТ").Activate
          
                       
        End If
        
        
        ' Следующая запись
        rowCount = rowCount + 1
        ' Application.StatusBar = officeNameInReport + ": " + CStr(rowCount) + "..."
        DoEventsInterval (rowCount)
      Loop
   
      ' Статус обработки
      Application.StatusBar = ""
      
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
    
    ' Переходим в ячейку M2
    ThisWorkbook.Sheets("СМОТ").Range("A1").Select

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

' Отправка письма: отправляю шаблон самому себе для последующей отправки в сеть письма на его основе:
Sub Отправка_Lotus_Notes_ЛистСМОТ(In_ФИО, In_СМОТFileName)
Dim темаПисьма, текстПисьма, hashTag, attachmentFile As String
Dim i As Byte
  
  ' Запрос
  ' If MsgBox("Отправить себе Шаблон письма с фокусами контроля '" + ПериодКонтроля + "'?", vbYesNo) = vbYes Then
    
    ' Адрес_получателя = ThisWorkbook.Sheets("Лист8").Cells(rowByValue(ThisWorkbook.Name, "СМОТ", "Список получателей:", 100, 100), ColumnByValue(ThisWorkbook.Name, "СМОТ", "Список получателей:", 100, 100) + 2).Value
    Адрес_получателя = getFromAddrBook3(In_ФИО)
    
    ' Тема письма - Тема:
    ' темаПисьма = ThisWorkbook.Sheets("Лист8").Cells(RowByValue(ThisWorkbook.Name, "Лист8", "Тема:", 100, 100), ColumnByValue(ThisWorkbook.Name, "Лист8", "Тема:", 100, 100) + 1).Value
    темаПисьма = subjectFromSheet("СМОТ")

    ' hashTag - Хэштэг:
    ' hashTag = ThisWorkbook.Sheets("Лист8").Cells(RowByValue(ThisWorkbook.Name, "Лист8", "Хэштэг:", 100, 100), ColumnByValue(ThisWorkbook.Name, "Лист8", "Хэштэг:", 100, 100) + 1).Value
    ' hashTag - Хэштэг:
    hashTag = hashTagFromSheet("СМОТ")

    ' Файл-вложение (!!!)
    attachmentFile = In_СМОТFileName
    
    ' Текст письма
    текстПисьма = "" + Chr(13)
    текстПисьма = текстПисьма + "" + Адрес_получателя + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "Добрый день!" + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "Направляю расчет СМОТ для согласования" + Chr(13)
    ' текстПисьма = текстПисьма + "" + Chr(13)
    ' текстПисьма = текстПисьма + "" + Chr(13)
    ' текстПисьма = текстПисьма + "" + Chr(13)
    ' текстПисьма = текстПисьма + "" + Chr(13)
    ' текстПисьма = текстПисьма + "" + Chr(13)
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

