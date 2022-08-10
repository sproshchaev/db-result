Attribute VB_Name = "Module_План"
' План

' Шаблон_обработки_5_офисов
Sub Декомпозиция_планов_продаж_на_квартал()
Dim ReportName_String, officeNameInReport, CheckFormatReportResult As String
Dim i, rowCount As Integer
Dim finishProcess As Boolean
Dim monthInQuarter1, monthInQuarter2, monthInQuarter3 As String
    
  ' Открыть файл с отчетом
  FileName = Application.GetOpenFilename("Excel Files (*.xlsx), *.xlsx", , "Открытие файла с Декомпозицией")

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
    ThisWorkbook.Sheets("План").Activate

    ' Проверка формы отчета
    CheckFormatReportResult = CheckFormatReport(ReportName_String, "ПК", 10, Date)
    
    If CheckFormatReportResult = "OK" Then
      
      ' 1. ПК Заполняем переменные названия месяцев: monthInQuarter1, monthInQuarter2, monthInQuarter3 - Апрель  Май Июнь
      monthInQuarter1 = Workbooks(ReportName_String).Sheets("ПК").Cells(2, 5).Value
      monthInQuarter2 = Workbooks(ReportName_String).Sheets("ПК").Cells(2, 6).Value
      monthInQuarter3 = Workbooks(ReportName_String).Sheets("ПК").Cells(2, 7).Value
      
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

        rowCount = 3
        Do While Not IsEmpty(Workbooks(ReportName_String).Sheets("ПК").Cells(rowCount, 4).Value)
        
          ' Если это текущий офис
          If InStr(Workbooks(ReportName_String).Sheets("ПК").Cells(rowCount, 4).Value, officeNameInReport) <> 0 Then
            
            ' План 1-го месяца
            ThisWorkbook.Sheets("План").Cells(5 + i, ColumnByNameAndNumber(ThisWorkbook.Name, "План", 4, monthInQuarter1, 1, 32)).Value = Workbooks(ReportName_String).Sheets("ПК").Cells(rowCount, 5).Value
            
            ' План 2-го месяца
            ThisWorkbook.Sheets("План").Cells(5 + i, ColumnByNameAndNumber(ThisWorkbook.Name, "План", 4, monthInQuarter2, 1, 32)).Value = Workbooks(ReportName_String).Sheets("ПК").Cells(rowCount, 6).Value
            
            ' План 3-го месяца
            ThisWorkbook.Sheets("План").Cells(5 + i, ColumnByNameAndNumber(ThisWorkbook.Name, "План", 4, monthInQuarter3, 1, 32)).Value = Workbooks(ReportName_String).Sheets("ПК").Cells(rowCount, 7).Value
            
          End If
        
          ' Следующая запись
          rowCount = rowCount + 1
          Application.StatusBar = officeNameInReport + ": " + CStr(rowCount) + "..."
          DoEventsInterval (rowCount)
        Loop
   
        ' Выводим данные по офису
      
      Next i ' Следующий офис
      
      ' 2. Ипотека
      rowCount = 3
      Do While Not IsEmpty(Workbooks(ReportName_String).Sheets("Ипотека").Cells(rowCount, 3).Value)
        
        ' Если это текущий офис ОО "Тюменский"
        If (InStr(Workbooks(ReportName_String).Sheets("Ипотека").Cells(rowCount, 3).Value, "Тюменский") <> 0) Then
            
          ' План 1-го месяца
          ThisWorkbook.Sheets("План").Cells(15, ColumnByNameAndNumber(ThisWorkbook.Name, "План", 4, monthInQuarter1, 1, 32)).Value = Workbooks(ReportName_String).Sheets("Ипотека").Cells(rowCount, 4).Value
            
          ' План 2-го месяца
          ThisWorkbook.Sheets("План").Cells(15, ColumnByNameAndNumber(ThisWorkbook.Name, "План", 4, monthInQuarter2, 1, 32)).Value = Workbooks(ReportName_String).Sheets("Ипотека").Cells(rowCount, 5).Value
            
          ' План 3-го месяца
          ThisWorkbook.Sheets("План").Cells(15, ColumnByNameAndNumber(ThisWorkbook.Name, "План", 4, monthInQuarter3, 1, 32)).Value = Workbooks(ReportName_String).Sheets("Ипотека").Cells(rowCount, 6).Value
            
        End If
        
        ' Если это текущий офис ОО2"Сургутский"
        If (InStr(Workbooks(ReportName_String).Sheets("Ипотека").Cells(rowCount, 3).Value, "Сургутский") <> 0) Then
            
          ' План 1-го месяца
          ThisWorkbook.Sheets("План").Cells(16, ColumnByNameAndNumber(ThisWorkbook.Name, "План", 4, monthInQuarter1, 1, 32)).Value = Workbooks(ReportName_String).Sheets("Ипотека").Cells(rowCount, 4).Value
            
          ' План 2-го месяца
          ThisWorkbook.Sheets("План").Cells(16, ColumnByNameAndNumber(ThisWorkbook.Name, "План", 4, monthInQuarter2, 1, 32)).Value = Workbooks(ReportName_String).Sheets("Ипотека").Cells(rowCount, 5).Value
            
          ' План 3-го месяца
          ThisWorkbook.Sheets("План").Cells(16, ColumnByNameAndNumber(ThisWorkbook.Name, "План", 4, monthInQuarter3, 1, 32)).Value = Workbooks(ReportName_String).Sheets("Ипотека").Cells(rowCount, 6).Value
            
        End If
        
        
        ' Следующая запись
        rowCount = rowCount + 1
        Application.StatusBar = "ИЦ: " + CStr(rowCount) + "..."
        DoEventsInterval (rowCount)
      
      Loop
      
    
      ' Выводим итоги обработки
      
      ' Переменная завершения обработки
      finishProcess = True
    Else
      ' Сообщение о неверном формате отчета или даты
      MsgBox ("Проверьте отчет: " + CheckFormatReportResult + "!")
    End If ' Проверка формы отчета

    ' Закрываем файл с отчетом без сохранения изменений (параметр SaveChanges:=False)
    Workbooks(Dir(FileName)).Close SaveChanges:=False
    
    ' Переходим в ячейку M2
    ThisWorkbook.Sheets("План").Range("A1").Select

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

