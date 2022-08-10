Attribute VB_Name = "Module_Лист10"
' Банкострахование (Лист10)
Sub Банкострахование()

' Описание переменных
Dim ReportName_String, officeNameInReport, CheckFormatReportResult As String
Dim i, rowCount As Integer
Dim finishProcess As Boolean
Dim Выдано_ПК_руб, Страхование_ЖиЗ_руб, Выдано_ПК_Турбоденьги_руб As Double
Dim Выдано_ПК_шт, Страхование_ЖиЗ_шт, Выдано_ПК_Турбоденьги_шт As Integer
    
    
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
    ThisWorkbook.Sheets("Лист10").Activate

    ' Проверка формы отчета
    CheckFormatReportResult = CheckFormatReport(ReportName_String, "Кредиты", 11, periodFromSheet("Лист10"))
    
    If CheckFormatReportResult = "OK" Then
          
    ' Дата окончания отчета в текущем месяце
    dateEndInMonth = periodFromSheet2("Лист10", 2)
      
    ' Дата начала месяца
    dateBeginMonth = Date_begin_day_month(dateEndInMonth)
            
    ' Заголовок отчета
    ThisWorkbook.Sheets("Лист10").Cells(2, 2).Value = "Банкострахование на " + ДеньМесяцГод(dateEndInMonth)
                        
    ' Неделя
    ThisWorkbook.Sheets("Лист10").Cells(2, 12).Value = CStr(WeekNumber(dateEndInMonth))
                  
    ' Тема письма
    ThisWorkbook.Sheets("Лист10").Cells(2, 19).Value = "Оперативная бизнес-справка (банкострахование) на  " + CStr(dateEndInMonth) + " г."
                                
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

      ' Обнуление переменных
      Выдано_ПК_руб = 0
      Выдано_ПК_шт = 0
      Страхование_ЖиЗ_руб = 0
      Страхование_ЖиЗ_шт = 0
      Выдано_ПК_Турбоденьги_руб = 0
      Выдано_ПК_Турбоденьги_шт = 0

      rowCount = 3
      Do While Not IsEmpty(Workbooks(ReportName_String).Sheets("Кредиты").Cells(rowCount, 1).Value)
        
          ' Если это кредит текущего офиса
          If InStr(Workbooks(ReportName_String).Sheets("Кредиты").Cells(rowCount, 8).Value, officeNameInReport) <> 0 Then
          
            ' Проверяем дату
            If (CDate(Workbooks(ReportName_String).Sheets("Кредиты").Cells(rowCount, 3).Value) >= dateBeginMonth) And (CDate(Workbooks(ReportName_String).Sheets("Кредиты").Cells(rowCount, 3).Value) <= dateEndInMonth) Then
              
              ' Выдано_ПК_руб
              Выдано_ПК_руб = Выдано_ПК_руб + CDbl(Replace(Workbooks(ReportName_String).Sheets("Кредиты").Cells(rowCount, 6).Value, ".", ","))
              
              ' Выдано_ПК_шт
              Выдано_ПК_шт = Выдано_ПК_шт + 1
              
              ' Считаем Турбоденьги
              If InStr(Workbooks(ReportName_String).Sheets("Кредиты").Cells(rowCount, 9).Value, "Турбоденьги") <> 0 Then
              
                ' Турбоденьги, руб.
                Выдано_ПК_Турбоденьги_руб = Выдано_ПК_Турбоденьги_руб + CDbl(Replace(Workbooks(ReportName_String).Sheets("Кредиты").Cells(rowCount, 6).Value, ".", ","))
              
                ' Турбоденьги, шт
                Выдано_ПК_Турбоденьги_шт = Выдано_ПК_Турбоденьги_шт + 1
              
              End If
              
              
              ' Страхование_ЖиЗ
              If InStr(Workbooks(ReportName_String).Sheets("Кредиты").Cells(rowCount, 10).Value, "да") <> 0 Then
              
                ' Страхование_ЖиЗ_руб
                Страхование_ЖиЗ_руб = Страхование_ЖиЗ_руб + CDbl(Replace(Workbooks(ReportName_String).Sheets("Кредиты").Cells(rowCount, 6).Value, ".", ","))
              
                ' Страхование_ЖиЗ_шт
                Страхование_ЖиЗ_шт = Страхование_ЖиЗ_шт + 1
              
              End If
              
            End If ' Проверяем дату
            
          End If ' Если это кредит текущего офиса
        
          ' Следующая запись
          rowCount = rowCount + 1
          Application.StatusBar = officeNameInReport + ": " + CStr(rowCount) + "..."
          DoEventsInterval (rowCount)
      Loop
   
        ' Выводим данные по офису
        ThisWorkbook.Sheets("Лист10").Cells(5 + i, 3).Value = Round(Выдано_ПК_руб / 1000, 0)
        ThisWorkbook.Sheets("Лист10").Cells(5 + i, 4).Value = Выдано_ПК_шт
        ThisWorkbook.Sheets("Лист10").Cells(5 + i, 5).Value = Round(Страхование_ЖиЗ_руб / 1000, 0)
        ThisWorkbook.Sheets("Лист10").Cells(5 + i, 6).Value = Страхование_ЖиЗ_шт
        ThisWorkbook.Sheets("Лист10").Cells(5 + i, 7).Value = РассчетДоли(Выдано_ПК_шт, Страхование_ЖиЗ_шт, 2)
        ' Выдано без Турбоденег
        ThisWorkbook.Sheets("Лист10").Cells(5 + i, 8).Value = Round((Выдано_ПК_руб - Выдано_ПК_Турбоденьги_руб) / 1000, 0)
        ThisWorkbook.Sheets("Лист10").Cells(5 + i, 9).Value = Выдано_ПК_шт - Выдано_ПК_Турбоденьги_шт
        ThisWorkbook.Sheets("Лист10").Cells(5 + i, 10).Value = Round(Страхование_ЖиЗ_руб / 1000, 0)
        ThisWorkbook.Sheets("Лист10").Cells(5 + i, 11).Value = Страхование_ЖиЗ_шт
        
        ' Доля по штукам
        ' ThisWorkbook.Sheets("Лист10").Cells(5 + i, 12).Value = РассчетДоли((Выдано_ПК_шт - Выдано_ПК_Турбоденьги_шт), Страхование_ЖиЗ_шт, 2)
        ' Окраска ячейки СФЕТОФОР: если
        ' Call Full_Color_RangeII("Лист10", 5 + i, 12, (РассчетДоли((Выдано_ПК_шт - Выдано_ПК_Турбоденьги_шт), Страхование_ЖиЗ_шт, 2) * 100), 80)

        ' Доля по объему
        ThisWorkbook.Sheets("Лист10").Cells(5 + i, 12).Value = РассчетДоли((Выдано_ПК_руб - Выдано_ПК_Турбоденьги_руб), Страхование_ЖиЗ_руб, 2)
        ' Окраска ячейки СФЕТОФОР: если
        Call Full_Color_RangeII("Лист10", 5 + i, 12, (РассчетДоли((Выдано_ПК_руб - Выдано_ПК_Турбоденьги_руб), Страхование_ЖиЗ_руб, 2) * 100), 80)


    Next i ' Следующий офис
      
      ' Выводим итоги обработки - в штуках
      ' ThisWorkbook.Sheets("Лист10").Cells(11, 7).Value = РассчетДоли(ThisWorkbook.Sheets("Лист10").Cells(11, 4).Value, ThisWorkbook.Sheets("Лист10").Cells(11, 6).Value, 2)
      ' ThisWorkbook.Sheets("Лист10").Cells(11, 12).Value = РассчетДоли(ThisWorkbook.Sheets("Лист10").Cells(11, 9).Value, ThisWorkbook.Sheets("Лист10").Cells(11, 11).Value, 2)
      
      ' В обьеме
      ThisWorkbook.Sheets("Лист10").Cells(11, 7).Value = РассчетДоли(ThisWorkbook.Sheets("Лист10").Cells(11, 3).Value, ThisWorkbook.Sheets("Лист10").Cells(11, 5).Value, 2)
      ThisWorkbook.Sheets("Лист10").Cells(11, 12).Value = РассчетДоли(ThisWorkbook.Sheets("Лист10").Cells(11, 8).Value, ThisWorkbook.Sheets("Лист10").Cells(11, 10).Value, 2)
      
      ' Формируем список для отправки (в "Список получателей:"):
      ThisWorkbook.Sheets("Лист10").Cells(rowByValue(ThisWorkbook.Name, "Лист10", "Список получателей:", 100, 100), ColumnByValue(ThisWorkbook.Name, "Лист10", "Список получателей:", 100, 100) + 2).Value = getFromAddrBook("УДО2,УДО3,УДО4,УДО5,НОРПиКО1,НОРПиКО2,НОРПиКО3,НОРПиКО4,НОРПиКО5,ПМ,МРК1,МРК2,МРК3,МРК4,МРК5", 2)
      
      ' Переменная завершения обработки
      finishProcess = True
      
      ' Зачеркиваем пункт меню на стартовой страницы
      Call ЗачеркиваемТекстВячейке("Лист0", RangeByValue(ThisWorkbook.Name, "Лист0", "Банкострахование", 100, 100))

    Else
      ' Сообщение о неверном формате отчета или даты
      MsgBox ("Проверьте отчет: " + CheckFormatReportResult + "!")
    End If ' Проверка формы отчета

    ' Закрываем файл с отчетом без сохранения изменений (параметр SaveChanges:=False)
    Workbooks(Dir(FileName)).Close SaveChanges:=False
    
    ' Переходим в ячейку M2
    ThisWorkbook.Sheets("Лист10").Range("O1").Select

    ' Строка статуса
    Application.StatusBar = ""
    
    ' Итоговое сообщение
    If finishProcess = True Then
      MsgBox ("Обработка " + Dir(ReportName_String) + " завершена!")
    Else
      MsgBox ("Обработка отчета была прервана!")
    End If

  End If ' Если файл был выбран

End Sub

' Отправка письма: отправляю шаблон самому себе для последующей отправки в сеть письма на его основе:
Sub Отправка_Lotus_Notes_Лист10_Банкострахование()
Dim темаПисьма, текстПисьма, hashTag As String
Dim i As Byte
  
  If MsgBox("Отправить себе Шаблон письма?", vbYesNo) = vbYes Then
    
    ' Тема письма - Тема:
    темаПисьма = subjectFromSheet("Лист10")

    ' hashTag - Хэштэг:
    hashTag = hashTagFromSheet("Лист10")
    
    ' Текст письма
    текстПисьма = "" + Chr(13)
    текстПисьма = текстПисьма + "" + ThisWorkbook.Sheets("Лист10").Cells(rowByValue(ThisWorkbook.Name, "Лист10", "Список получателей:", 100, 100), ColumnByValue(ThisWorkbook.Name, "Лист10", "Список получателей:", 100, 100) + 2).Value + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "" + getFromAddrBook("РД", 2) + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "Уважаемые сотрудники," + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "Проникновение страхования ЖиЗ в выданные потребкредиты." + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
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

