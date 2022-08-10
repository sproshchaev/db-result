Attribute VB_Name = "Module_Лист16"
' Просрочки в CRM - обработка отчета "Продукты по статусам"
Sub Обработка_Продукты_по_статусам()
  
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
  
    ' В "S3" вносим наименование отчета:
    ThisWorkbook.Sheets("Лист16").Range("S3").Value = FileName
  
    ' Переменная начала обработки
    finishProcess = False

    ' Выводим для инфо данные об имени файла
    ReportName_String = Dir(FileName)
  
    ' Открываем выбранную книгу (UpdateLinks:=0)
    Workbooks.Open FileName, 0
      
    ' Переходим на окно DB
    ThisWorkbook.Sheets("Лист16").Activate

    ' Проверка формы отчета
    ' CheckFormatReportResult = CheckFormatReport(ReportName_String, "___", 6, Date)
    ' If CheckFormatReportResult = "OK" Then
    If True Then
            
      ' Файл с отчетом
      ThisWorkbook.Sheets("Лист16").Range("Q3").Value = FileName
      
      ' Обновляем список получателей
      ThisWorkbook.Sheets("Лист16").Cells(rowByValue(ThisWorkbook.Name, "Лист16", "Список получателей:", 100, 100), ColumnByValue(ThisWorkbook.Name, "Лист16", "Список получателей:", 100, 100) + 2).Value = _
         getFromAddrBook("УДО2,УДО3,УДО4,УДО5,НОРПиКО1,НОРПиКО2,НОРПиКО3,НОРПиКО4,НОРПиКО5", 2)

      ' Тема
      ThisWorkbook.Sheets("Лист16").Range("P2").Value = "CRM Dynamics 365 - Продукты по статусам " + CStr(Date)
                  
      ' Очистка таблицы
      Call clearСontents2(ThisWorkbook.Name, "Лист16", "C7", "H11")
            
      ' Получаем имя Листа в файле "Продукты по статусам"
      Sheet_Name_In_Report = Workbooks(ReportName_String).Sheets(1).Name

      ' Столбцы в отчете
      Column_Допофис = ColumnByValue(ReportName_String, Sheet_Name_In_Report, "Доп. офис", 100, 100)
      Column_Продукт_оформлен = ColumnByValue(ReportName_String, Sheet_Name_In_Report, "Продукт оформлен", 100, 100)
      Column_Менеджер = ColumnByValue(ReportName_String, Sheet_Name_In_Report, "Менеджер", 100, 100)
      Column_Встреча_просрочена = ColumnByValue(ReportName_String, Sheet_Name_In_Report, "Встреча просрочена", 100, 100)
      Column_Думает_после_встречи = ColumnByValue(ReportName_String, Sheet_Name_In_Report, "Думает после встречи", 100, 100)
      Column_Думает_после_звонка = ColumnByValue(ReportName_String, Sheet_Name_In_Report, "Думает после звонка", 100, 100)
      Column_Менеджер_назначен_нет_активностей = ColumnByValue(ReportName_String, Sheet_Name_In_Report, "Менеджер назначен, нет активностей", 100, 100)
      
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

        rowCount = 2
        ThisOffice = False
        Do While Not IsEmpty(Workbooks(ReportName_String).Sheets(Sheet_Name_In_Report).Cells(rowCount, Column_Продукт_оформлен + 2).Value)
        
          ' Если это текущий офис - взводим тригер
          If InStr(Workbooks(ReportName_String).Sheets(Sheet_Name_In_Report).Cells(rowCount, Column_Допофис).Value, officeNameInReport) <> 0 Then
            
            ThisOffice = True
                
          End If
        
          ' Если это текущий Офис ThisOffice = True и ячейки В и С пустые - берем значения для этого офиса из строки
          If (ThisOffice = True) And (IsEmpty(Workbooks(ReportName_String).Sheets(Sheet_Name_In_Report).Cells(rowCount, Column_Допофис).Value)) And (IsEmpty(Workbooks(ReportName_String).Sheets(Sheet_Name_In_Report).Cells(rowCount, Column_Менеджер).Value)) Then
            
            ' Выводим данные:
            ' В очереди - Встреча просрочена. Красим в красный цвет, если не ноль!
            Call Write_Лист16(ReportName_String, Sheet_Name_In_Report, _
                                rowCount, _
                                  Column_Встреча_просрочена, _
                                    6 + i, _
                                      3, _
                                        1)
                                      
            
            ' В работе - Думает после встречи
            ' ThisWorkbook.Sheets("Лист16").Cells(6 + i, 4).Value = Workbooks(ReportName_String).Sheets(Sheet_Name_In_Report).Cells(rowCount, 8).Value
            Call Write_Лист16(ReportName_String, Sheet_Name_In_Report, _
                                rowCount, _
                                  Column_Думает_после_встречи, _
                                    6 + i, _
                                      4, _
                                        0)
            
            
            ' В работе - Думает после звонка
            ' ThisWorkbook.Sheets("Лист16").Cells(6 + i, 5).Value = Workbooks(ReportName_String).Sheets(Sheet_Name_In_Report).Cells(rowCount, 9).Value
            Call Write_Лист16(ReportName_String, Sheet_Name_In_Report, _
                                rowCount, _
                                  Column_Думает_после_звонка, _
                                    6 + i, _
                                      5, _
                                        0)

            
            ' В работе - Менеджер назначен, нет активностей. Красим в красный цвет, если не ноль!
            ' ThisWorkbook.Sheets("Лист16").Cells(6 + i, 6).Value = Workbooks(ReportName_String).Sheets(Sheet_Name_In_Report).Cells(rowCount, 11).Value
            Call Write_Лист16(ReportName_String, Sheet_Name_In_Report, _
                                rowCount, _
                                  Column_Менеджер_назначен_нет_активностей, _
                                    6 + i, _
                                      6, _
                                        1)

            ' Всего (Продукт оформлен + 2)
            ' ThisWorkbook.Sheets("Лист16").Cells(6 + i, 7).Value = Workbooks(ReportName_String).Sheets(Sheet_Name_In_Report).Cells(Column_Допофис1).Value
            Call Write_Лист16(ReportName_String, Sheet_Name_In_Report, _
                                rowCount, _
                                  Column_Продукт_оформлен + 2, _
                                    6 + i, _
                                      7, _
                                       0)
            
            
            ' Продукт оформлен
            ' ThisWorkbook.Sheets("Лист16").Cells(6 + i, 8).Value = Workbooks(ReportName_String).Sheets(Sheet_Name_In_Report).Cells(rowCount, 19).Value
            Call Write_Лист16(ReportName_String, Sheet_Name_In_Report, _
                                rowCount, _
                                  Column_Продукт_оформлен, _
                                    6 + i, _
                                      8, _
                                        0)

            
            ' Сбрасываем тригер
            ThisOffice = False
                
          End If
          
          
        
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
      Call gorizontalLineII(ThisWorkbook.Name, "Лист16", 12, 2, 9)
      ' ----------------------------------------------------------------------------------------------------------------------------------
      
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
    ThisWorkbook.Sheets("Лист16").Range("A1").Select

    ' Строка статуса
    Application.StatusBar = ""

    ' Зачеркиваем пункт меню на стартовой страницы
    Call ЗачеркиваемТекстВячейке("Лист0", RangeByValue(ThisWorkbook.Name, "Лист0", "Просрочки CRM", 100, 100))
    
    ' Итоговое сообщение
    If finishProcess = True Then
      MsgBox ("Обработка " + Dir(ReportName_String) + " завершена!")
    Else
      MsgBox ("Обработка отчета была прервана!")
    End If

  End If ' Если файл был выбран

End Sub


' Отправка письма: отправляю шаблон самому себе для последующей отправки в сеть письма на его основе:
Sub Отправка_Lotus_Notes_Лист16()
Dim темаПисьма, текстПисьма, hashTag, attachmentFile As String
Dim i As Byte
  
  If MsgBox("Отправить себе Шаблон письма?", vbYesNo) = vbYes Then
    
    ' Тема письма - Тема:
    ' темаПисьма = ThisWorkbook.Sheets("Лист16").Cells(RowByValue(ThisWorkbook.Name, "Лист16", "Тема:", 100, 100), ColumnByValue(ThisWorkbook.Name, "Лист16", "Тема:", 100, 100) + 1).Value
    темаПисьма = subjectFromSheet("Лист16")

    ' hashTag - Хэштэг:
    ' hashTag = ThisWorkbook.Sheets("Лист16").Cells(RowByValue(ThisWorkbook.Name, "Лист16", "Хэштэг:", 100, 100), ColumnByValue(ThisWorkbook.Name, "Лист16", "Хэштэг:", 100, 100) + 1).Value
    ' hashTag - Хэштэг:
    hashTag = hashTagFromSheet("Лист16")

    ' Файл-вложение (!!!)
    attachmentFile = ThisWorkbook.Sheets("Лист16").Cells(3, 17).Value
    
    ' Текст письма
    текстПисьма = "" + Chr(13)
    текстПисьма = текстПисьма + "" + ThisWorkbook.Sheets("Лист16").Cells(rowByValue(ThisWorkbook.Name, "Лист16", "Список получателей:", 100, 100), ColumnByValue(ThisWorkbook.Name, "Лист16", "Список получателей:", 100, 100) + 2).Value + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "" + getFromAddrBook("РД", 2) + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "Уважаемые руководители," + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "Прошу отработать в офисах контакты с просроченными встречами и отсутствием активности!" + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "Нормативы отработки маркетинговых кампаний: доля думающих - 40%, доля отказов- 70%, доля выдач - 30% , доля просроченных - 20%." + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
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


            
' Выводим данные
Sub Write_Лист16(In_ReportName_String, In_Sheet_Name_In_Report, In_rowCount, In_columnCount, In_rowCount_Лист16, In_columnCount_Лист16, In_Color) ' 6 + i, 3)
           
            If IsEmpty(Workbooks(In_ReportName_String).Sheets(In_Sheet_Name_In_Report).Cells(In_rowCount, In_columnCount).Value) = False Then
              
              ' В очереди - Встреча просрочена
              ThisWorkbook.Sheets("Лист16").Cells(In_rowCount_Лист16, In_columnCount_Лист16).Value = Workbooks(In_ReportName_String).Sheets(In_Sheet_Name_In_Report).Cells(In_rowCount, In_columnCount).Value
              
              ' Если переменная In_Color = 1, то красим в желтый если ячейка не нулевая!
              If In_Color = 1 Then
                ThisWorkbook.Sheets("Лист16").Cells(In_rowCount_Лист16, In_columnCount_Лист16).Interior.Color = vbYellow
              End If
              
            Else
              
              ' Если ячейка была пустая, то записываем 0 в нее
              ThisWorkbook.Sheets("Лист16").Cells(In_rowCount_Лист16, In_columnCount_Лист16).Value = 0
              
            End If
            
            ' Центруем
            ThisWorkbook.Sheets("Лист16").Cells(In_rowCount_Лист16, In_columnCount_Лист16).HorizontalAlignment = xlCenter

End Sub


' Переход в браузере по ссылке на странице, в конце ставить "/" для перехода в категорию
Sub goToURL_Лист16()
  
  ' SheetsVar = ThisWorkbook.ActiveSheet.Name
  ' rowVar = rowByValue(ThisWorkbook.Name, SheetsVar, "Ссылка:", 100, 100)
  ' columnVar = ColumnByValue(ThisWorkbook.Name, SheetsVar, "Ссылка:", 100, 100) + 1
  
  ' ThisWorkbook.FollowHyperlink ("http://isrb.psbnk.msk.ru/inf/6601/6622/ejednevnii_otchet_po_prodajam/")
  ThisWorkbook.FollowHyperlink (ThisWorkbook.Sheets("Лист16").Range("T32").Value)
  
End Sub


' Активности по сотруднику (звонки)
Sub Обработка_Активности_по_сотруднику()
Dim ReportName_String, officeNameInReport, CheckFormatReportResult As String
Dim i, rowCount As Integer
Dim finishProcess As Boolean
Dim Дата_отчета As Date
    
  ' Открыть файл с отчетом
  FileName = Application.GetOpenFilename("Excel Files (*.xlsx), *.xlsx", , "Открытие файла с отчетом")

  ' Если файл был выбран
  If (Len(FileName) > 5) Then
  
    ' Строка статуса
    Application.StatusBar = "Обработка отчета..."
  
    ' В "S3" вносим наименование отчета:
    ThisWorkbook.Sheets("Лист16").Range("S17").Value = FileName
  
    ' Переменная начала обработки
    finishProcess = False

    ' Выводим для инфо данные об имени файла
    ReportName_String = Dir(FileName)
  
    ' Открываем выбранную книгу (UpdateLinks:=0)
    Workbooks.Open FileName, 0
      
    ' Переходим на окно DB
    ThisWorkbook.Sheets("Лист16").Activate

    ' Проверка формы отчета
    ' CheckFormatReportResult = CheckFormatReport(ReportName_String, "___", 6, Date)
    ' If CheckFormatReportResult = "OK" Then
    If True Then
            
      ' Файл с отчетом
      ThisWorkbook.Sheets("Лист16").Range("Q17").Value = FileName
      
      ' Обновляем список получателей
      ThisWorkbook.Sheets("Лист16").Cells(rowByValue(ThisWorkbook.Name, "Лист16", "Список получателей2:", 100, 100), ColumnByValue(ThisWorkbook.Name, "Лист16", "Список получателей2:", 100, 100) + 2).Value = _
         getFromAddrBook("УДО2,УДО3,УДО4,УДО5,НОРПиКО1,НОРПиКО2,НОРПиКО3,НОРПиКО4,НОРПиКО5", 2)

      ' Тема
      ThisWorkbook.Sheets("Лист16").Range("P16").Value = "CRM Dynamics 365 - Активности по сотрудникам " + CStr(Date)
                  
      ' Очистка таблицы
      Call clearСontents2(ThisWorkbook.Name, "Лист16", "D20", "D24")
            
      ' Получаем имя Листа в файле "Продукты по статусам"
      Sheet_Name_In_Report = Workbooks(ReportName_String).Sheets(1).Name

      ' Столбцы в отчете
      Column_Допофис = ColumnByValue(ReportName_String, Sheet_Name_In_Report, "Доп. офис", 100, 100)
      Column_Менеджер = ColumnByValue(ReportName_String, Sheet_Name_In_Report, "Менеджер", 100, 100)
      Column_Активность = ColumnByValue(ReportName_String, Sheet_Name_In_Report, "Активность", 100, 100)
      Column_Категория_продукта = ColumnByValue(ReportName_String, Sheet_Name_In_Report, "Категория продукта", 100, 100)
      Column_Продукт = ColumnByValue(ReportName_String, Sheet_Name_In_Report, "Продукт", 100, 100)
      Column_Активность_состоялась = ColumnByValue(ReportName_String, Sheet_Name_In_Report, "Активность состоялась", 100, 100)
      Column_Продажа = ColumnByValue(ReportName_String, Sheet_Name_In_Report, "Продажа", 100, 100)
      Column_Дата = ColumnByValue(ReportName_String, Sheet_Name_In_Report, "Дата", 100, 100)
      
      ' Строки на "Лист16"
      row_Форма_16_2 = rowByValue(ThisWorkbook.Name, "Лист16", "Форма 16.2", 100, 100)
      
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

        ' Ставим 0
        ThisWorkbook.Sheets("Лист16").Cells(row_Форма_16_2 + 2 + i, 4).Value = 0
        ThisWorkbook.Sheets("Лист16").Cells(row_Форма_16_2 + 2 + i, 4).NumberFormat = "#,##0"
        ThisWorkbook.Sheets("Лист16").Cells(row_Форма_16_2 + 2 + i, 4).HorizontalAlignment = xlCenter

        rowCount = 2
        ThisOffice = False
        ' Do While rowCount < 300 ' "Всего:" в 6-ом столбце
        Do While InStr(Workbooks(ReportName_String).Sheets(Sheet_Name_In_Report).Cells(rowCount, Column_Продукт).Value, "Всего:") = 0
        
          ' Из Column_Дата берем дату
          If Workbooks(ReportName_String).Sheets(Sheet_Name_In_Report).Cells(rowCount, Column_Дата).Value <> "" Then
            Дата_отчета = Workbooks(ReportName_String).Sheets(Sheet_Name_In_Report).Cells(rowCount, Column_Дата).Value
          End If
        
          ' Если это текущий офис - взводим тригер
          If (InStr(Workbooks(ReportName_String).Sheets(Sheet_Name_In_Report).Cells(rowCount, Column_Допофис).Value, officeNameInReport) <> 0) And (InStr(Workbooks(ReportName_String).Sheets(Sheet_Name_In_Report).Cells(rowCount, Column_Допофис).Value, "ИЦ") = 0) Then
            
            ' t0 = Workbooks(ReportName_String).Sheets(Sheet_Name_In_Report).Cells(rowCount, Column_Допофис).Value
            
            Итого_звонков_по_офису = 0
            
            ThisOffice = True
                
          End If
        
          ' Если новая секция с офисом B, С, D, E, F, то сбрасываем. Можно еще по цветам - комбинация столбец-цвет, столбец2-цвет2 и т.п.
          ' If (ThisOffice = True) And (IsEmpty(Workbooks(ReportName_String).Sheets(Sheet_Name_In_Report).Cells(rowCount, Column_Допофис).Value) = False) And (IsEmpty(Workbooks(ReportName_String).Sheets(Sheet_Name_In_Report).Cells(rowCount, Column_Допофис + 1).Value) = False) And (IsEmpty(Workbooks(ReportName_String).Sheets(Sheet_Name_In_Report).Cells(rowCount, Column_Допофис + 2).Value) = False) And (IsEmpty(Workbooks(ReportName_String).Sheets(Sheet_Name_In_Report).Cells(rowCount, Column_Допофис + 3).Value) = False) And (IsEmpty(Workbooks(ReportName_String).Sheets(Sheet_Name_In_Report).Cells(rowCount, Column_Допофис + 4).Value) = False) And (InStr(Workbooks(ReportName_String).Sheets(Sheet_Name_In_Report).Cells(rowCount, Column_Допофис).Value, officeNameInReport) = 0) And (InStr(Workbooks(ReportName_String).Sheets(Sheet_Name_In_Report).Cells(rowCount, Column_Допофис).Value, "ИЦ") = 0) Then
          If (ThisOffice = True) And _
               (Workbooks(ReportName_String).Sheets(Sheet_Name_In_Report).Cells(rowCount, Column_Допофис + 1).Interior.Color = 15128749) And _
                 (Workbooks(ReportName_String).Sheets(Sheet_Name_In_Report).Cells(rowCount, Column_Допофис + 2).Interior.Color = 15128749) And _
                   (Workbooks(ReportName_String).Sheets(Sheet_Name_In_Report).Cells(rowCount, Column_Допофис + 3).Interior.Color = 15128749) And _
                     (Workbooks(ReportName_String).Sheets(Sheet_Name_In_Report).Cells(rowCount, Column_Допофис + 4).Interior.Color = 16777215) Then
            
            t0 = Workbooks(ReportName_String).Sheets(Sheet_Name_In_Report).Cells(rowCount, Column_Допофис).Value
            
            ThisOffice = False
                
            ' Заносим число сотрудников
            ThisWorkbook.Sheets("Лист16").Cells(row_Форма_16_2 + 2 + i, 3).Value = Число_сотрудников_в_офисе_Лист7(i)
            ThisWorkbook.Sheets("Лист16").Cells(row_Форма_16_2 + 2 + i, 3).NumberFormat = "#,##0"
            ThisWorkbook.Sheets("Лист16").Cells(row_Форма_16_2 + 2 + i, 3).HorizontalAlignment = xlCenter
                
            ' Заносим итоги
            ThisWorkbook.Sheets("Лист16").Cells(row_Форма_16_2 + 2 + i, 4).Value = Итого_звонков_по_офису
            ThisWorkbook.Sheets("Лист16").Cells(row_Форма_16_2 + 2 + i, 4).NumberFormat = "#,##0"
            ThisWorkbook.Sheets("Лист16").Cells(row_Форма_16_2 + 2 + i, 4).HorizontalAlignment = xlCenter
                
            ' Выполнение норматива дня
            ThisWorkbook.Sheets("Лист16").Cells(row_Форма_16_2 + 2 + i, 5).Value = (ThisWorkbook.Sheets("Лист16").Cells(row_Форма_16_2 + 2 + i, 4).Value) / (ThisWorkbook.Sheets("Лист16").Cells(row_Форма_16_2 + 2 + i, 3).Value * 20)
            ThisWorkbook.Sheets("Лист16").Cells(row_Форма_16_2 + 2 + i, 5).NumberFormat = "0%"
            ThisWorkbook.Sheets("Лист16").Cells(row_Форма_16_2 + 2 + i, 5).HorizontalAlignment = xlCenter
                
          End If
          
        
          ' Если это текущий Офис ThisOffice = True и ячейки D, E, F пустые - берем значения для этого офиса из строки
          If (ThisOffice = True) And _
               (Workbooks(ReportName_String).Sheets(Sheet_Name_In_Report).Cells(rowCount, Column_Допофис + 1).Interior.Color = 15128749) And _
                 (Workbooks(ReportName_String).Sheets(Sheet_Name_In_Report).Cells(rowCount, Column_Допофис + 2).Interior.Color = 15658671) And _
                   (Workbooks(ReportName_String).Sheets(Sheet_Name_In_Report).Cells(rowCount, Column_Допофис + 3).Interior.Color = 15658671) And _
                     (Workbooks(ReportName_String).Sheets(Sheet_Name_In_Report).Cells(rowCount, Column_Допофис + 4).Interior.Color = 16777215) Then
          
          ' If (ThisOffice = True) And (IsEmpty(Workbooks(ReportName_String).Sheets(Sheet_Name_In_Report).Cells(rowCount, Column_Допофис + 2).Value)) And (IsEmpty(Workbooks(ReportName_String).Sheets(Sheet_Name_In_Report).Cells(rowCount, Column_Допофис + 3).Value)) And (IsEmpty(Workbooks(ReportName_String).Sheets(Sheet_Name_In_Report).Cells(rowCount, Column_Допофис + 4).Value)) And (Workbooks(ReportName_String).Sheets(Sheet_Name_In_Report).Cells(rowCount, Column_Допофис + 4).Interior.Color = 16777215) Then
          ' If (InStr(Workbooks(ReportName_String).Sheets(Sheet_Name_In_Report).Cells(rowCount, Column_Допофис).Value, officeNameInReport) <> 0) And (InStr(Workbooks(ReportName_String).Sheets(Sheet_Name_In_Report).Cells(rowCount, Column_Допофис).Value, "ИЦ") = 0) And (IsEmpty(Workbooks(ReportName_String).Sheets(Sheet_Name_In_Report).Cells(rowCount, Column_Допофис + 2).Value)) And (IsEmpty(Workbooks(ReportName_String).Sheets(Sheet_Name_In_Report).Cells(rowCount, Column_Допофис + 3).Value)) And (IsEmpty(Workbooks(ReportName_String).Sheets(Sheet_Name_In_Report).Cells(rowCount, Column_Допофис + 4).Value)) Then
            
            ' Итого_звонков_по_офису
            Итого_звонков_по_офису = Итого_звонков_по_офису + CheckData(Workbooks(ReportName_String).Sheets(Sheet_Name_In_Report).Cells(rowCount, Column_Активность_состоялась).Value)
            
            t1 = rowCount
            
            t2 = Workbooks(ReportName_String).Sheets(Sheet_Name_In_Report).Cells(rowCount, Column_Допофис + 2).Value
            t3 = Workbooks(ReportName_String).Sheets(Sheet_Name_In_Report).Cells(rowCount, Column_Допофис + 3).Value
            t3 = Workbooks(ReportName_String).Sheets(Sheet_Name_In_Report).Cells(rowCount, Column_Допофис + 4).Value
            t4 = Workbooks(ReportName_String).Sheets(Sheet_Name_In_Report).Cells(rowCount, Column_Допофис + 4).Interior.Color
            t5 = Workbooks(ReportName_String).Sheets(Sheet_Name_In_Report).Cells(rowCount, Column_Допофис + 4).Interior.Pattern
        
            ' .Pattern = xlSolid ' .PatternColorIndex = xlAutomatic ' .Color = 65535 ' .TintAndShade = 0 ' .PatternTintAndShade = 0
            
            ' Итоги по сотруднику (цвета ячеек в строке)
            t_3 = Workbooks(ReportName_String).Sheets(Sheet_Name_In_Report).Cells(35, 3).Interior.Color
            t_4 = Workbooks(ReportName_String).Sheets(Sheet_Name_In_Report).Cells(35, 4).Interior.Color
            t_5 = Workbooks(ReportName_String).Sheets(Sheet_Name_In_Report).Cells(35, 5).Interior.Color
            t_6 = Workbooks(ReportName_String).Sheets(Sheet_Name_In_Report).Cells(35, 6).Interior.Color
            
            ' Итоги по офису (цвета ячеек в строке)
            t__3 = Workbooks(ReportName_String).Sheets(Sheet_Name_In_Report).Cells(36, 3).Interior.Color
            t__4 = Workbooks(ReportName_String).Sheets(Sheet_Name_In_Report).Cells(36, 4).Interior.Color
            t__5 = Workbooks(ReportName_String).Sheets(Sheet_Name_In_Report).Cells(36, 5).Interior.Color
            t__6 = Workbooks(ReportName_String).Sheets(Sheet_Name_In_Report).Cells(36, 6).Interior.Color
            
            t4 = 0
            
        
            ' Выводим данные:
            ' В очереди - Встреча просрочена. Красим в красный цвет, если не ноль!
            ' Call Write_Лист16(ReportName_String, Sheet_Name_In_Report, _
            '                    rowCount, _
            '                      Column_Встреча_просрочена, _
            '                        6 + i, _
            '                          3, _
            '                            1)
                                      
            
            ' Сбрасываем тригер
            ' ThisOffice = False
                
          End If
          
          
        
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
      Call gorizontalLineII(ThisWorkbook.Name, "Лист16", 25, 2, 5)
      ' ----------------------------------------------------------------------------------------------------------------------------------
      
      ' Выполнение норматива дня
      ThisWorkbook.Sheets("Лист16").Cells(row_Форма_16_2 + 2 + i, 5).Value = (ThisWorkbook.Sheets("Лист16").Cells(row_Форма_16_2 + 2 + i, 4).Value) / (ThisWorkbook.Sheets("Лист16").Cells(row_Форма_16_2 + 2 + i, 3).Value * 20)
      ThisWorkbook.Sheets("Лист16").Cells(row_Форма_16_2 + 2 + i, 5).NumberFormat = "0%"
      ThisWorkbook.Sheets("Лист16").Cells(row_Форма_16_2 + 2 + i, 5).HorizontalAlignment = xlCenter
      
      ' Дата_отчета
      Дата_отчета = CDate(Mid(CStr(Дата_отчета), 1, 10))
      ThisWorkbook.Sheets("Лист16").Cells(row_Форма_16_2 - 1, 2).Value = "CRM Dynamics 365 Активности по сотруднику за " + CStr(Дата_отчета) + " г."
      ThisWorkbook.Sheets("Лист16").Range("P16").Value = "CRM Dynamics 365 - Активности по сотрудникам (звонки) " + CStr(Дата_отчета)
            
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
    ThisWorkbook.Sheets("Лист16").Range("A1").Select

    ' Строка статуса
    Application.StatusBar = ""

    ' Зачеркиваем пункт меню на стартовой страницы
    Call ЗачеркиваемТекстВячейке("Лист0", RangeByValue(ThisWorkbook.Name, "Лист0", "Просрочки CRM", 100, 100))
    
    ' Итоговое сообщение
    If finishProcess = True Then
      MsgBox ("Обработка " + Dir(ReportName_String) + " завершена!")
    Else
      MsgBox ("Обработка отчета была прервана!")
    End If

  End If ' Если файл был выбран

End Sub


' Отправка письма: отправляю шаблон самому себе для последующей отправки в сеть письма на его основе:
Sub Отправка_Lotus_Notes_Лист16_Активность_по_звонкам()
Dim темаПисьма, текстПисьма, hashTag, attachmentFile As String
Dim i As Byte
  
  If MsgBox("Отправить себе Шаблон письма?", vbYesNo) = vbYes Then
    
    ' Тема письма - Тема:
    ' темаПисьма = ThisWorkbook.Sheets("Лист16").Cells(RowByValue(ThisWorkbook.Name, "Лист16", "Тема:", 100, 100), ColumnByValue(ThisWorkbook.Name, "Лист16", "Тема:", 100, 100) + 1).Value
    темаПисьма = subjectFromSheetII("Лист16", 2)

    ' hashTag - Хэштэг:
    ' hashTag = ThisWorkbook.Sheets("Лист16").Cells(RowByValue(ThisWorkbook.Name, "Лист16", "Хэштэг:", 100, 100), ColumnByValue(ThisWorkbook.Name, "Лист16", "Хэштэг:", 100, 100) + 1).Value
    ' hashTag - Хэштэг:
    hashTag = hashTagFromSheetII("Лист16", 2)

    ' Файл-вложение (!!!)
    attachmentFile = ThisWorkbook.Sheets("Лист16").Range("Q17").Value
    
    ' Текст письма
    текстПисьма = "" + Chr(13)
    текстПисьма = текстПисьма + "" + ThisWorkbook.Sheets("Лист16").Cells(rowByValue(ThisWorkbook.Name, "Лист16", "Список получателей2:", 100, 100), ColumnByValue(ThisWorkbook.Name, "Лист16", "Список получателей2:", 100, 100) + 2).Value + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "" + getFromAddrBook("РД", 2) + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "Уважаемые руководители," + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "Выполнение норматива по исходящим звонкам за " + CStr(Дата_отчета_Форма_16_2()) + " г." + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
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

' Дата
Function Дата_отчета_Форма_16_2() As Date

  ' Строки на "Лист16"
  row_Форма_16_2 = rowByValue(ThisWorkbook.Name, "Лист16", "Форма 16.2", 100, 100)

  ' Дата_отчета "CRM Dynamics 365 Активности по сотруднику за " + CStr(Дата_отчета) + " г."
  ' t = Mid(ThisWorkbook.Sheets("Лист16").Cells(row_Форма_16_2 - 1, 2).Value, 46, 10)
  Дата_отчета_Форма_16_2 = CDate(Mid(ThisWorkbook.Sheets("Лист16").Cells(row_Форма_16_2 - 1, 2).Value, 46, 10))
  
End Function
