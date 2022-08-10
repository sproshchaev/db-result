Attribute VB_Name = "Module_Templates"
' Шаблон

' Шаблон_обработки_5_офисов
Sub Шаблон_обработки_5_офисов()
  
' Описание переменных
Dim ReportName_String, officeNameInReport, CheckFormatReportResult As String
Dim i, rowCount As Integer
Dim finishProcess As Boolean
    
  ' Открыть файл с отчетом
  FileName = Application.GetOpenFilename("Excel Files (*.xml), *.xml", , "Открытие файла с отчетом")

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
    ThisWorkbook.Sheets("Лист3").Activate

    ' Проверка формы отчета
    CheckFormatReportResult = CheckFormatReport(ReportName_String, "___", 6, Date)
    If CheckFormatReportResult = "OK" Then
      
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

        rowCount = 1
        Do While Not IsEmpty(Workbooks(ReportName_String).Sheets("Список").Cells(rowCount, 1).Value)
        
          ' Если это текущий офис
          If InStr(Workbooks(ReportName_String).Sheets("Карты в сейфах").Cells(rowCount, 2).Value, officeNameInReport) <> 0 Then
            
            
                
          End If
        
        
          ' Следующая запись
          rowCount = rowCount + 1
          Application.StatusBar = officeNameInReport + ": " + CStr(rowCount) + "..."
          DoEventsInterval (rowCount)
        Loop
   
        ' Выводим данные по офису
      
      Next i ' Следующий офис
      
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
    ThisWorkbook.Sheets("Лист3").Range("L78").Select

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


' Шаблон вставки записи в BASE\Имя_книги на Лист 1, ключевое поле
Sub Шаблон_вставки_зависи_в_BASE()
    
    Call InsertRecordInBook("Имя_книги", "Лист1", "Ключевое_поле", "Значение_ключа", _
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
                                                                              "", "", _
                                                                                "", "", _
                                                                                  "", "")

End Sub


' Нахождение ячейки на Листе
Sub Шаблон_фиксации_ячейки_на_листе()
Dim Range_str As String
Dim Range_Row, Range_Column As Byte

  ' Находим ячейку (например G41), в которой записано значение In_К_пор
  Range_str = RangeByValue(In_Workbooks, In_Sheets, In_К_пор, 100, 100)
  Range_Row = Workbooks(In_Workbooks).Sheets(In_Sheets).Range(Range_str).Row
  Range_Column = Workbooks(In_Workbooks).Sheets(In_Sheets).Range(Range_str).Column

End Sub
