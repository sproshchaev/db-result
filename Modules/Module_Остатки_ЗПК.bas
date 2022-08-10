Attribute VB_Name = "Module_Остатки_ЗПК"
' Лист "Остатки ЗПК"

' Удаление из выгруженных файлов информацию о сотрудниках ПСБ
Sub Подготовка_файлов_БК_из_Way4()

' Описание переменных
Dim ReportName_String, officeNameInReport, CheckFormatReportResult As String
' Dim i, rowCount As Integer
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
    ThisWorkbook.Sheets("Остатки ЗПК").Activate

    ' Проверка формы отчета
    ' CheckFormatReportResult = CheckFormatReport(ReportName_String, "___", 6, Date)
    ' If CheckFormatReportResult = "OK" Then
    If True Then
      
      ' Очистка таблицы
      Call clearСontents2(ThisWorkbook.Name, "Остатки ЗПК", "A7", "I50")
      записьНаЛистСМОТ = 6
      номер = 0

      ' Находим столбец "Код ЗП-организации"
      Column_Код_ЗП_организации = ColumnByValue(ReportName_String, "BranchClientReport", "Код ЗП-организации", 100, 100)
      ' Строка "Финансовый институт"
      rowCount = rowByValue(ReportName_String, "BranchClientReport", "Финансовый институт", 100, 100) + 1
      
      ' Сохраняем файл как
      ' Имя нового файла
      ' "Данные за квартал по Физики с даты 01.10.2020 00:00:00 ( report_id = 381 )"
      typeDownload = ""
      If InStr(Workbooks(ReportName_String).Sheets("BranchClientReport").Range("A1").Value, "Данные за квартал по Физики") <> 0 Then
        typeDownload = "физлица"
        dateReport = CDate(Mid(Workbooks(ReportName_String).Sheets("BranchClientReport").Range("A1").Value, 36, 10))
        ' Тема
        ThisWorkbook.Sheets("Остатки ЗПК").Range("P2").Value = "Остатки на картах - " + Workbooks(ReportName_String).Sheets("BranchClientReport").Range("A1").Value
        ' Тема
        ThisWorkbook.Sheets("Остатки ЗПК").Range("P2").Value = "Остатки на картах - Физлица"
      End If
      
      ' "Данные за квартал по Зарплатники с даты 01.10.2020 00:00:00 ( report_id = 382 )"
      If InStr(Workbooks(ReportName_String).Sheets("BranchClientReport").Range("A1").Value, "Данные за квартал по Зарплатники") <> 0 Then
        typeDownload = "ЗП-проекты"
        dateReport = CDate(Mid(Workbooks(ReportName_String).Sheets("BranchClientReport").Range("A1").Value, 41, 10))
        ' Тема
        ThisWorkbook.Sheets("Остатки ЗПК").Range("P2").Value = "Остатки на картах - ЗП проекты"
      End If
      
      ' Имя нового файла
      FileNewVar = "Остатки_на_картах_" + typeDownload + "_" + CStr(Date_last_day_quarter(dateReport)) + ".xls"
      Workbooks(ReportName_String).SaveAs FileName:=ThisWorkbook.Path + "\Out\" + FileNewVar, createBackUp:=False
      ' Имя файла без расширения
      ИмяФайлаБезРасширения = "Остатки_на_картах_" + typeDownload + "_" + CStr(Date_last_day_quarter(dateReport))
      ' Имя файла для архивации
      strFilesNameForArch = ThisWorkbook.Path + "\Out\" + FileNewVar
      ' Вложение Q3
      ThisWorkbook.Sheets("Остатки ЗПК").Range("Q3").Value = strFilesNameForArch
      
      ' Вывод в таблицу
      ' №
      ThisWorkbook.Sheets("Остатки ЗПК").Cells(7, 1).Value = "1"
      ' Имя файла
      ThisWorkbook.Sheets("Остатки ЗПК").Cells(7, 2).Value = FileNewVar
      
      ' Переменные
      Всего_записей = 0
      Удалено_записей = 0
      Выгружено_записей = 0
      
      ' Обработка
      Do While Not IsEmpty(Workbooks(FileNewVar).Sheets("BranchClientReport").Cells(rowCount, 1).Value)
        
        ' Переменные
        Всего_записей = Всего_записей + 1
        Выгружено_записей = Выгружено_записей + 1
        
        ' Если это Код организации - Удалить зарплатные карты ПСБ Код организации 00128
        If InStr(Workbooks(FileNewVar).Sheets("BranchClientReport").Cells(rowCount, Column_Код_ЗП_организации).Value, "00128") <> 0 Then
          
          ' Удаляем текущую запись в файле - это работает, но окно нельзя закрывать
          Workbooks(FileNewVar).Sheets("BranchClientReport").Rows(CStr(rowCount) + ":" + CStr(rowCount)).Delete Shift:=xlUp
              
          ' Если удадили запись, то стоим на ней же
          rowCount = rowCount - 1
        
          ' Переменные
          Удалено_записей = Удалено_записей + 1
          Выгружено_записей = Выгружено_записей - 1
        
        End If
        
        
        ' Следующая запись
        rowCount = rowCount + 1
        Application.StatusBar = "Обработка: " + CStr(rowCount) + "..."
        DoEventsInterval (rowCount)
      Loop
   
      ' Статус обработки
      Application.StatusBar = ""
      
      ' Выводим итоги обработки
      ' Вывод в таблицу
      ' Записей
      ThisWorkbook.Sheets("Остатки ЗПК").Cells(7, 3).Value = Всего_записей
      ' Удалено
      ThisWorkbook.Sheets("Остатки ЗПК").Cells(7, 4).Value = Удалено_записей
      ' Выгружено
      ThisWorkbook.Sheets("Остатки ЗПК").Cells(7, 5).Value = Выгружено_записей

      ' Сохранение изменений
      ThisWorkbook.Save
    
     ' Формируем список для отправки (в "Список получателей:"):
     ThisWorkbook.Sheets("Остатки ЗПК").Cells(rowByValue(ThisWorkbook.Name, "Остатки ЗПК", "Список получателей:", 100, 100), ColumnByValue(ThisWorkbook.Name, "Остатки ЗПК", "Список получателей:", 100, 100) + 2).Value _
       = getFromAddrBook("УДО2,УДО3,УДО4,УДО5,НОРПиКО1,НОРПиКО2,НОРПиКО3,НОРПиКО4,НОРПиКО5,НОКП", 2)
    
      ' Переменная завершения обработки
      finishProcess = True
    Else
      ' Сообщение о неверном формате отчета или даты
      MsgBox ("Проверьте отчет: " + CheckFormatReportResult + "!")
    End If ' Проверка формы отчета

    ' Закрываем файл с отчетом без сохранения изменений (параметр SaveChanges:=False)
    Workbooks(FileNewVar).Close SaveChanges:=True
    
    ' Переходим в ячейку M2
    ' ThisWorkbook.Sheets("Остатки ЗПК").Range("A1").Select

    ' Архивируем файл
    Application.StatusBar = "Создание архива..."

    ' Запускаем архиватор этого файла, Справка https://www.dmosk.ru/miniinstruktions.php?mini=7zip-cmd
    ' -sdel Удалить файлы после создания архива
    ' Имя файла архива
    File7zipName = ИмяФайлаБезРасширения + ".zip"
    Shell ("C:\Program Files\7-Zip\7z a -tzip -ssw -mx9 C:\Users\PROSCHAEVSF\Documents\#DB_Result\Out\" + File7zipName + " " + strFilesNameForArch)
    ' Вложение Q3
    ThisWorkbook.Sheets("Остатки ЗПК").Range("Q3").Value = "C:\Users\PROSCHAEVSF\Documents\#DB_Result\Out\" + File7zipName
    ' Сообщение
    Application.StatusBar = "Архив создан!"

    ' Строка статуса
    Application.StatusBar = ""

    ' Зачеркиваем пункт меню на стартовой страницы
    ' Call ЗачеркиваемТекстВячейке("Лист0", "D9")
    ' Call ЗачеркиваемТекстВячейке("Лист0", RangeByValue(ThisWorkbook.Name, "Лист0", "Оперативная справка по _________________", 100, 100))
    
    ' Итоговое сообщение
    If finishProcess = True Then
      MsgBox ("Обработка " + Dir(ReportName_String) + " завершена!")
      Call Отправка_Lotus_Notes_Остатки_ЗПК2
    Else
      MsgBox ("Обработка отчета была прервана!")
    End If

  End If ' Если файл был выбран
    
End Sub

' Отправка письма: отправляю шаблон самому себе для последующей отправки в сеть письма на его основе
Sub Отправка_Lotus_Notes_Остатки_ЗПК2()
Dim темаПисьма, текстПисьма, hashTag, attachmentFile As String
Dim i As Byte
  
  If MsgBox("Отправить себе Шаблон письма?", vbYesNo) = vbYes Then
    
    ' Тема письма - Тема:
    темаПисьма = ThisWorkbook.Sheets("Остатки ЗПК").Range("P2").Value

    ' hashTag - Хэштэг:
    hashTag = hashTagFromSheet("Остатки ЗПК")
    
    ' Файл-вложение (!!!)
    attachmentFile = ThisWorkbook.Sheets("Остатки ЗПК").Range("Q3").Value
    
    ' Список получателей
    Список_получателей = recipientList("Остатки ЗПК")
    
    ' Текст письма
    текстПисьма = "" + Chr(13)
    текстПисьма = текстПисьма + "" + Список_получателей + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "" + getFromAddrBook("РД", 2) + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "Уважаемые руководители," + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "Направляю выгрузку по банковским картам " + кавычки + ThisWorkbook.Sheets("Остатки ЗПК").Range("B7").Value + кавычки + " для отработки Инвестов/ИСЖ/НСЖ" + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    ' Визитка (подпись С Ув., )
    текстПисьма = текстПисьма + ПодписьВПисьме()
    ' Хэштег
    текстПисьма = текстПисьма + createBlankStr(25) + hashTag
    ' Вызов
    Call send_Lotus_Notes(темаПисьма, "Sergey Fedorovich Proschaev/Tyumen/PSBank/Ru", "Sergey Fedorovich Proschaev/Tyumen/PSBank/Ru", текстПисьма, attachmentFile)
  
    ' Сообщение
    MsgBox ("Письмо отправлено!")
     
  End If
  
End Sub



' Это уже не работает - файлы формируем самостоятельно в Way4
' Отправка письма: отправляю шаблон самому себе для последующей отправки в сеть письма на его основе
Sub Отправка_Lotus_Notes_Остатки_ЗПК()
Dim темаПисьма, текстПисьма, hashTag As String
Dim i As Byte
  
  If MsgBox("Отправить себе Шаблон письма?", vbYesNo) = vbYes Then
    
    ' Тема письма - Тема:
    темаПисьма = "Тюменский РОО - заказ выгрузки карт DPK"

    ' hashTag - Хэштэг:
    hashTag = "#ОстаткиЗПК"
    
    ' Текст письма
    текстПисьма = "" + Chr(13)
    текстПисьма = текстПисьма + "Elena Alexandrovna Belova/CardCentre/PSBank/Ru" + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "Елена, добрый день!" + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "Прошу по зарплатным банковским картам Тюменского РОО сформировать выгрузку данных DPK." + Chr(13)
    ' текстПисьма = текстПисьма + "" + Chr(13)
    ' текстПисьма = текстПисьма + "" + Chr(13)
    ' Визитка (подпись С Ув., )
    текстПисьма = текстПисьма + ПодписьВПисьме()
    ' Хэштег
    текстПисьма = текстПисьма + createBlankStr(20) + hashTag
    ' Вызов
    Call send_Lotus_Notes2(темаПисьма, "Sergey Fedorovich Proschaev/Tyumen/PSBank/Ru", "Sergey Fedorovich Proschaev/Tyumen/PSBank/Ru", "", текстПисьма, "")
  
    ' Сообщение
    MsgBox ("Письмо отправлено!")
     
  End If
  
End Sub

