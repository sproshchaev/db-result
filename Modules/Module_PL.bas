Attribute VB_Name = "Module_PL"
' Отчетность от Котельниковой
Sub PL_в_разрезе_ДО_по_бизнесам()

Dim ReportName_String, officeNameInReport, CheckFormatReportResult As String
Dim i, rowCount As Integer
Dim finishProcess As Boolean
    
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
    ThisWorkbook.Sheets("PL").Activate

    ' Проверка формы отчета
    CheckFormatReportResult = CheckFormatReport(ReportName_String, "Оглавление ", 16, Date)
    If CheckFormatReportResult = "OK" Then
      
      
      ' Очищаем таблицу
      Call Очистить_PL1_PL2_ЛистPl
      
      ' Обрабатываем отчет
      ' Цикл по 5-ти офисам
      ' Обработка отчета
      For i = 1 To 5
        ' Номера офисов от 1 до 5
        Select Case i
          Case 1 ' ОО «Тюменский»
            officeNameInReport = "ГО " + Chr(34) + "Тюменский" + Chr(34)
          Case 2 ' ОО «Сургутский»
            officeNameInReport = "ДО " + Chr(34) + "Сургутский" + Chr(34)
          Case 3 ' ОО «Нижневартовский»
            officeNameInReport = "ДО " + Chr(34) + "Нижневартовский" + Chr(34)
          Case 4 ' ОО «Новоуренгойский»
            officeNameInReport = "ДО " + Chr(34) + "Новоуренгойский" + Chr(34)
          Case 5 ' ОО «Тарко-Сале»
            officeNameInReport = "ДО " + Chr(34) + "Тарко-Сале" + Chr(34)
        End Select


        ' Сокращенное наименование Листа (меньше текста)
        SheetsNameVar = "PL в разрезе ДО по бизнесам"

        ' *** Первый цикл обработки ***
        
         
        ' Столбцы
        Column_Названия_строк = ColumnByValue(ReportName_String, SheetsNameVar, "Названия строк", 100, 11)
        Column_Названия_столбцов = ColumnByValue(ReportName_String, SheetsNameVar, "Названия столбцов", 100, 11)
        Column_ГГГГ_Итог = Column_Названия_столбцов + 4
        
        ' Бизнеса:
        For j = 1 To 6
        Select Case j
          Case 1 ' MB RM (253)
            currBusiness = "MB RM (253)"
          Case 2 ' MASS (254)
            currBusiness = "MASS (254)"
          Case 3 ' Средний бизнес (27)
            currBusiness = "Средний бизнес (27)"
          Case 4 ' Корпоративные клиенты (1)
            currBusiness = "Корпоративные клиенты (1)"
          Case 5 ' VIP-клиенты (22)
            currBusiness = "VIP-клиенты (22)"
          Case 6 ' РБ (202)
            currBusiness = "РБ (202)"
        End Select
        
        ' Переменные обработки
        Сейчас_текущий_офис = False
        Сейчас_currBusiness = False
        
        ' Ячейка "Названия строк"
        rowCount = rowByValue(ReportName_String, SheetsNameVar, "Названия строк", 1000, 10)
        Do While Not IsEmpty(Workbooks(ReportName_String).Sheets(SheetsNameVar).Cells(rowCount, Column_Названия_строк).Value)
        
          ' --- Секция Офис ---
          ' Если находим в ячейке есть 'ДО "' или 'Рег.Сеть Связь-Банк', то значит следующий офис пошел
          If (Сейчас_текущий_офис = True) And ((InStr(Workbooks(ReportName_String).Sheets(SheetsNameVar).Cells(rowCount, Column_Названия_строк).Value, ("ДО " + Chr(34))) <> 0) Or (InStr(Workbooks(ReportName_String).Sheets(SheetsNameVar).Cells(rowCount, Column_Названия_строк).Value, ("Рег.Сеть Связь-Банк")) <> 0)) Then
            Сейчас_текущий_офис = False
          End If
          
          ' Если это текущий офис
          If InStr(Workbooks(ReportName_String).Sheets(SheetsNameVar).Cells(rowCount, Column_Названия_строк).Value, officeNameInReport) <> 0 Then
            Сейчас_текущий_офис = True
          End If
          ' --- Секция Офис ---
        
          ' --- Секция текущего бизнеса  ---
          If (Сейчас_текущий_офис = True) And InStr(Workbooks(ReportName_String).Sheets(SheetsNameVar).Cells(rowCount, Column_Названия_строк).Value, currBusiness) <> 0 Then
            
            Сейчас_currBusiness = True
            
            ' Обнуляем переменные
            ' Операционный_результат_ГГГГ_Итог = 0
            
            ' Резервы есть не по всем бизнесам
            ' Резервы = 0
            ' Расчитаны_резервы_по_бизнесу = False
            '
            ' Операционные_расходы_ГГГГ_Итог = 0
            ' Операционные_расходы_Кв_минус_1 = 0
          
          End If
          ' --- Секция текущего бизнеса ---
          
          ' --- Обработка строк  в текущем офисе и секции РБ ---
          ' Если это строка "Операционный результат (1)" для офиса и РБ
          If (Сейчас_текущий_офис = True) And (Сейчас_currBusiness = True) And (InStr(Workbooks(ReportName_String).Sheets(SheetsNameVar).Cells(rowCount, Column_Названия_строк).Value, "Операционный результат (1)") <> 0) Then
            
            ' Берем из столбца 6 "202X Итог"
            Операционный_результат_ГГГГ_Итог = Round(Workbooks(ReportName_String).Sheets(SheetsNameVar).Cells(rowCount, Column_ГГГГ_Итог).Value / 1000, 0)
            
            ' Выводим в таблицу на Листе PL
            Call setDataInSheetPL(officeNameInReport, "Опер.рез", currBusiness, Операционный_результат_ГГГГ_Итог)
            
          End If
          
          ' Если это строка Резервы (4398)
          If (Сейчас_текущий_офис = True) And (Сейчас_currBusiness = True) And (InStr(Workbooks(ReportName_String).Sheets(SheetsNameVar).Cells(rowCount, Column_Названия_строк).Value, "Резервы (4398)") <> 0) Then
            
            ' Берем из столбца 6 "202X Итог"
            Резервы = Round(Workbooks(ReportName_String).Sheets(SheetsNameVar).Cells(rowCount, Column_ГГГГ_Итог).Value / 1000, 0)
           
            ' Выводим в таблицу на Листе PL
            Call setDataInSheetPL(officeNameInReport, "Резервы", currBusiness, Резервы)
            
          End If
          
          ' Если это строка Операционные расходы (2105)
          If (Сейчас_текущий_офис = True) And (Сейчас_currBusiness = True) And (InStr(Workbooks(ReportName_String).Sheets(SheetsNameVar).Cells(rowCount, Column_Названия_строк).Value, "Операционные расходы (2105)") <> 0) Then
            
            ' Берем из столбца 6 "202X Итог"
            Операционные_расходы_ГГГГ_Итог = Round(Workbooks(ReportName_String).Sheets(SheetsNameVar).Cells(rowCount, Column_ГГГГ_Итог).Value / 1000, 0)
            
            ' Выводим в таблицу на Листе PL
            Call setDataInSheetPL(officeNameInReport, "Год_ОР", currBusiness, Операционные_расходы_ГГГГ_Итог)
            
            ' Берем теперь "Кв.-1_ОР". Нам нужно от Column_ГГГГ_Итог влево найти первый не пустой квартал
            Операционные_расходы_Кв_минус_1_определены = False
            column_count = Column_ГГГГ_Итог - 1
            Do While (column_count >= 2) And (Операционные_расходы_Кв_минус_1_определены = False)
            
              ' Проверяем - если в ячейке есть значение, то берем его
              If Not IsEmpty(Workbooks(ReportName_String).Sheets(SheetsNameVar).Cells(rowCount, column_count).Value) Then
                Операционные_расходы_Кв_минус_1 = Round(Workbooks(ReportName_String).Sheets(SheetsNameVar).Cells(rowCount, column_count).Value / 1000, 0)
                Операционные_расходы_Кв_минус_1_определены = True
              End If
              
              ' Уменьшаем столбец на 1 (двигаемся влево)
              column_count = column_count - 1
            Loop
            
            ' Выводим в таблицу на Листе PL
            Call setDataInSheetPL(officeNameInReport, "Кв.-1_ОР", currBusiness, Операционные_расходы_Кв_минус_1)
          
          End If
          
          ' Если это строка Чистая прибыль (1409), то это завершение секции по бизнесу, берем Инфраструктурные расходы
          If (Сейчас_текущий_офис = True) And (Сейчас_currBusiness = True) And (InStr(Workbooks(ReportName_String).Sheets(SheetsNameVar).Cells(rowCount, Column_Названия_строк).Value, "Чистая прибыль (1409)") <> 0) Then
            ' Если Чистая прибыль не справочна, то берем ее
            
            ' Вызываем процедуру расчета Инфраструктурных расходов
            Call Расчет_инфраструктурных_расходов(ReportName_String, SheetsNameVar, officeNameInReport, currBusiness)
            
            ' Завершаем секцию бизнеса
            Сейчас_currBusiness = False
            
          End If ' Если это строка Чистая прибыль (1409)
          
          ' --- Обработка строк ---
        
          ' Следующая запись
          rowCount = rowCount + 1
          Application.StatusBar = officeNameInReport + " " + currBusiness + ": " + CStr(rowCount) + "..."
          DoEventsInterval (rowCount)
        Loop
            
        ' *************************************************************************
        ' *** Форула  для самостоятельного расчета Расчетной ЧП = (Опер.рез.+Резервы+(Опер.расходы- Расходы на инфраструктуру))*(80%);       Для 4кв2020 также следует исключить 2/3   от суммы опер.расходов 3кв2020 за вычетом расходов на инфраструктуру 3кв2020. ***
        ' *************************************************************************
         
        ' >> Елена, Добрый день!
        ' Еще вопрос - данная формула будет действовать и в следующих кварталах, т.е. нужно только номера кварталов менять, а множители 2/3 и 80% остаются без изменений? -
        ' для всех периодов остается принцип ЧП = (Опер.рез.+Резервы+Опер.расходы за исключением инфраструктурных расходов)*80%
        ' Но есть отличия по Опер.расходам (они же АХР, они же ФСА)  - в каждом периоде алгоритм различается на расчет квартальных/месячных дат,
        ' т.е. за законченный квартал  в расчет включаются АХР именно этого квартала;
        ' за первый месяц незаконченного квартала = 1/3 от предыдущего квартала;
        ' за второй месяц незаконченного квартала = 2/3 от предыдущего квартала.
        ' Это связано с тем, что подгрузка в  PL   данных  из  модели ФСА  производится не ежемесячно, а  ежеквартально.
        ' >> Формула  для самостоятельного расчета Расчетной ЧП = (Опер.рез.+Резервы+(Опер.расходы- Расходы на инфраструктуру))*(80%);       Для 4кв2020 также следует исключить 2/3   от суммы опер.расходов 3кв2020 за вычетом расходов на инфраструктуру 3кв2020.
         
        ' Из I2 берем номер месяца в котором сделан расчет и определяем Коэф. по алгоритму: за первый месяц незаконченного квартала = 1/3 от предыдущего квартала; за второй месяц незаконченного квартала = 2/3 от предыдущего квартала
        Коэффициент_включения_АХР = Определение_Коэффициент_включения_АХР(ThisWorkbook.Sheets("PL").Range("I2").Value) ' для ноября 0.67 по отчетности за 11 мес.
        
        ОперРез = getDataInSheetPL(officeNameInReport, currBusiness, "Опер.рез")
        Резервы = getDataInSheetPL(officeNameInReport, currBusiness, "Резервы")
          
        ОперРасходы_Год = getDataInSheetPL(officeNameInReport, currBusiness, "Год_ОР")
        ОперРасходы_ПредКварт = getDataInSheetPL(officeNameInReport, currBusiness, "Кв.-1_ОР")
          
        РасходыИнфраструкт_Год = getDataInSheetPL(officeNameInReport, currBusiness, "Год_ИР")
        РасходыИнфраструкт_ПредКварт = getDataInSheetPL(officeNameInReport, currBusiness, "Кв.-1_ИР")
          
        Расчетная_ЧП = ОперРез + Резервы + (ОперРасходы_Год + Коэффициент_включения_АХР * ОперРасходы_ПредКварт) - ((РасходыИнфраструкт_Год + Коэффициент_включения_АХР * РасходыИнфраструкт_ПредКварт))
          
        Расчетная_ЧП = Расчетная_ЧП * 0.8
          
        ' Расчетная_ЧП = (getDataInSheetPL(officeNameInReport, "Опер.рез") + getDataInSheetPL(officeNameInReport, "Резервы") + (getDataInSheetPL(officeNameInReport, "Год_ОР") - getDataInSheetPL(officeNameInReport, "Кв.-1_ОР") / 3 * 2 - (getDataInSheetPL(officeNameInReport, "Год_ИР") - getDataInSheetPL(officeNameInReport, "Кв.-1_ИР") / 3 * 2))) * 0.8

        ' Выводим в таблицу на Листе PL
        Call setDataInSheetPL(officeNameInReport, "Расч.ЧП", currBusiness, Расчетная_ЧП)
       
        Next j ' Следующий Бизнес
      
      Next i ' Следующий офис
      
      ' Выводим итоги обработки
      
      ' Строка статуса
      Application.StatusBar = "Копирование итогов..."
      
      ' Копируем итоговый отчет в Книгу для отправки
      Call copyPLToSend
      
      Application.StatusBar = "Завершение..."
      
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
    ThisWorkbook.Sheets("PL").Range("A7").Select

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
Sub Отправка_Lotus_Notes_ЛистPL()
Dim темаПисьма, текстПисьма, hashTag, attachmentFile As String
Dim i As Byte
  
  ' Запрос
  If MsgBox("Отправить себе Шаблон письма с итогами обработки Воронки?", vbYesNo) = vbYes Then
    
    ' Тема письма - Тема:
    ' темаПисьма = ThisWorkbook.Sheets("Лист9").Cells(RowByValue(ThisWorkbook.Name, "Лист9", "Тема:", 100, 100), ColumnByValue(ThisWorkbook.Name, "Лист9", "Тема:", 100, 100) + 1).Value
    темаПисьма = subjectFromSheet("Лист9")

    ' hashTag - Хэштэг:
    ' hashTag = ThisWorkbook.Sheets("Лист9").Cells(RowByValue(ThisWorkbook.Name, "Лист9", "Хэштэг:", 100, 100), ColumnByValue(ThisWorkbook.Name, "Лист9", "Хэштэг:", 100, 100) + 1).Value
    ' hashTag - Хэштэг:
    hashTag = hashTagFromSheet("Лист9")

    ' Файл-вложение (!!!)
    attachmentFile = ThisWorkbook.Sheets("Лист9").Range("Q3").Value
    
    ' Текст письма
    текстПисьма = "" + Chr(13)
    текстПисьма = текстПисьма + "" + ThisWorkbook.Sheets("Лист9").Cells(rowByValue(ThisWorkbook.Name, "Лист9", "Список получателей:", 100, 100), ColumnByValue(ThisWorkbook.Name, "Лист9", "Список получателей:", 100, 100) + 2).Value + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "" + getFromAddrBook("РД", 2) + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "Уважаемые руководители," + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "Воронка по потреб кредитам и кредитным картам" + Chr(13)
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

' Выводим в таблицу на Листе PL
Sub setDataInSheetPL(In_officeNameInReport, In_Param, In_currBusiness, In_Value)
  
  If InStr(In_officeNameInReport, "Тюменский") Then
    In_officeNameInReport_ЛистPL = "ОО «Тюменский»"
  End If
  
  If InStr(In_officeNameInReport, "Сургутский") Then
    In_officeNameInReport_ЛистPL = "ОО «Сургутский»"
  End If
  
  If InStr(In_officeNameInReport, "Нижневартовский") Then
    In_officeNameInReport_ЛистPL = "ОО «Нижневартовский»"
  End If
  
  If InStr(In_officeNameInReport, "Новоуренгойский") Then
    In_officeNameInReport_ЛистPL = "ОО «Новоуренгойский»"
  End If
  
  If InStr(In_officeNameInReport, "Тарко-Сале") Then
    In_officeNameInReport_ЛистPL = "ОО «Тарко-Сале»"
  End If
  
  ' Столбцы на Листе PL
  Column_Param = ColumnByValue(ThisWorkbook.Name, "PL", In_Param, 5, 10)
  
  ' Форма PL1 (только РБ)
  If In_currBusiness = "РБ (202)" Then
    Row_officeName = rowByValue(ThisWorkbook.Name, "PL", In_officeNameInReport_ЛистPL, 10, 2)
    ThisWorkbook.Sheets("PL").Cells(Row_officeName, Column_Param).Value = In_Value
  End If
  
  ' Форма PL2 Все бизнесы
  код_офиса_бизнеса = CStr(getNumberOfficeByName2(In_officeNameInReport_ЛистPL)) + "_" + In_currBusiness
  Row_officeName_currBusiness = rowByValue(ThisWorkbook.Name, "PL", код_офиса_бизнеса, 51, 2)
  ThisWorkbook.Sheets("PL").Cells(Row_officeName_currBusiness, Column_Param).Value = In_Value
    
  
End Sub

' Получаем из таблицы данные на Листе PL
Function getDataInSheetPL(In_officeNameInReport, In_currBusiness, In_Param)
  
  If InStr(In_officeNameInReport, "Тюменский") Then
    In_officeNameInReport_ЛистPL = "ОО «Тюменский»"
  End If
  
  If InStr(In_officeNameInReport, "Сургутский") Then
    In_officeNameInReport_ЛистPL = "ОО «Сургутский»"
  End If
  
  If InStr(In_officeNameInReport, "Нижневартовский") Then
    In_officeNameInReport_ЛистPL = "ОО «Нижневартовский»"
  End If
  
  If InStr(In_officeNameInReport, "Новоуренгойский") Then
    In_officeNameInReport_ЛистPL = "ОО «Новоуренгойский»"
  End If
  
  If InStr(In_officeNameInReport, "Тарко-Сале") Then
    In_officeNameInReport_ЛистPL = "ОО «Тарко-Сале»"
  End If
  
  Column_Param = ColumnByValue(ThisWorkbook.Name, "PL", In_Param, 5, 10)
  ' Row_officeName = rowByValue(ThisWorkbook.Name, "PL", In_officeNameInReport_ЛистPL, 10, 2)
  
  ' Форма PL2 Все бизнесы
  код_офиса_бизнеса = CStr(getNumberOfficeByName2(In_officeNameInReport_ЛистPL)) + "_" + In_currBusiness
  Row_officeName_currBusiness = rowByValue(ThisWorkbook.Name, "PL", код_офиса_бизнеса, 51, 2)

  getDataInSheetPL = ThisWorkbook.Sheets("PL").Cells(Row_officeName_currBusiness, Column_Param).Value
  
End Function

' Из I2 берем номер месяца в котором сделан расчет и определяем Коэф. по алгоритму: за первый месяц незаконченного квартала = 1/3 от предыдущего квартала; за второй месяц незаконченного квартала = 2/3 от предыдущего квартала
Function Определение_Коэффициент_включения_АХР(In_Номер_месяца_отчетности) ' для ноября 0.67 по отчетности за 11 мес.
                    
          Select Case In_Номер_месяца_отчетности
          Case 1, 4, 7, 10
            Определение_Коэффициент_включения_АХР = 0.3333333
          Case 2, 5, 8, 11
            Определение_Коэффициент_включения_АХР = 0.6666666
        End Select

End Function

' Вызываем процедуру расчета Инфраструктурных расходов
Sub Расчет_инфраструктурных_расходов(In_ReportName_String, In_SheetsNameVar, In_officeNameInReport, In_currBusiness)
    
        ' *************************************************************************
        ' *** Второй цикл обработки этого же листа по Инфраструктурным расходам ***
        ' *************************************************************************
        ' Переменные обработки
        Сейчас_текущий_офис = False
        Сейчас_РБ_202 = False
         
        ' Столбцы
        Column_Названия_строк = ColumnByValue2(In_ReportName_String, In_SheetsNameVar, "Названия строк", 100, 100, 2)
        Column_Названия_столбцов = ColumnByValue2(In_ReportName_String, In_SheetsNameVar, "Названия столбцов", 100, 100, 2)
        Column_ГГГГ_Итог = Column_Названия_столбцов + 4
        
        ' Ячейка "Названия строк"
        rowCount = rowByValue2(In_ReportName_String, In_SheetsNameVar, "Названия строк", 1000, 100, 2)
        Do While Not IsEmpty(Workbooks(In_ReportName_String).Sheets(In_SheetsNameVar).Cells(rowCount, Column_Названия_строк).Value)
        
          ' --- Секция Офис ---
          ' Если находим в ячейке есть 'ДО "' или 'Рег.Сеть Связь-Банк', то значит следующий офис пошел
          If (Сейчас_текущий_офис = True) And ((InStr(Workbooks(In_ReportName_String).Sheets(In_SheetsNameVar).Cells(rowCount, Column_Названия_строк).Value, ("ДО " + Chr(34))) <> 0) Or (InStr(Workbooks(In_ReportName_String).Sheets(In_SheetsNameVar).Cells(rowCount, Column_Названия_строк).Value, ("Рег.Сеть Связь-Банк")) <> 0)) Then
            Сейчас_текущий_офис = False
          End If
          
          ' Если это текущий офис
          If InStr(Workbooks(In_ReportName_String).Sheets(In_SheetsNameVar).Cells(rowCount, Column_Названия_строк).Value, In_officeNameInReport) <> 0 Then
            Сейчас_текущий_офис = True
          End If
          ' --- Секция Офис ---
        
          ' --- Обработка строк ---
          ' Если это строка "Операционный результат (1)" для офиса и РБ
          ' If (Сейчас_текущий_офис = True) And (InStr(Workbooks(In_ReportName_String).Sheets(In_SheetsNameVar).Cells(rowCount, Column_Названия_строк).Value, "РБ (202)") <> 0) Then
          If (Сейчас_текущий_офис = True) And (InStr(Workbooks(In_ReportName_String).Sheets(In_SheetsNameVar).Cells(rowCount, Column_Названия_строк).Value, In_currBusiness) <> 0) Then
            
            ' Берем из столбца "202X Итог"
            Расходы_на_инфраструктуру_ГГГГ_Итог = Round(Workbooks(In_ReportName_String).Sheets(In_SheetsNameVar).Cells(rowCount, Column_ГГГГ_Итог).Value / 1000, 0)
            ' Выводим в таблицу на Листе PL
            Call setDataInSheetPL(In_officeNameInReport, "Год_ИР", In_currBusiness, Расходы_на_инфраструктуру_ГГГГ_Итог)
            
            ' Берем теперь "Кв.-1_ОР". Нам нужно от Column_ГГГГ_Итог влево найти первый не пустой квартал
            Расходы_на_инфраструктуру_Кв_минус_1_определены = False
            column_count = Column_ГГГГ_Итог - 1
            Do While (column_count >= 2) And (Расходы_на_инфраструктуру_Кв_минус_1_определены = False)
            
              ' Проверяем - если в ячейке есть значение, то берем его
              If Not IsEmpty(Workbooks(In_ReportName_String).Sheets(In_SheetsNameVar).Cells(rowCount, column_count).Value) Then
                Расходы_на_инфраструктуру_Кв_минус_1 = Round(Workbooks(In_ReportName_String).Sheets(In_SheetsNameVar).Cells(rowCount, column_count).Value / 1000, 0)
                Расходы_на_инфраструктуру_Кв_минус_1_определены = True
              End If
              
              ' Уменьшаем столбец на 1 (двигаемся влево)
              column_count = column_count - 1
            Loop
            ' Выводим в таблицу на Листе PL
            Call setDataInSheetPL(In_officeNameInReport, "Кв.-1_ИР", In_currBusiness, Расходы_на_инфраструктуру_Кв_минус_1)
                    
          End If
          
          ' --- Обработка строк ---
        
          ' Следующая запись
          rowCount = rowCount + 1
          Application.StatusBar = "Расходы на инфраструктуту " + In_officeNameInReport + " " + In_currBusiness + ": " + CStr(rowCount) + "..."
          DoEventsInterval (rowCount)
        Loop
    
    
End Sub

' Очищаем таблицу
Sub Очистить_PL1_PL2_ЛистPl()
      ' PL 1
      For i = 6 To 10
        For j = 3 To 9
          ThisWorkbook.Sheets("PL").Cells(i, j).Value = 0
        Next j
      Next i
      
      ' PL 2
      For o = 1 To 5
        For i = 18 To 23
          For j = 3 To 9
            ThisWorkbook.Sheets("PL").Cells(i + ((o - 1) * 7), j).Value = 0
          Next j
        Next i
      Next o
      
End Sub

' Делаем копию листа (здесь не используем)
Sub SaveToXlsFrom_ЛистPl()

  ' Копируем Лист2
  ThisWorkbook.Sheets("PL").Copy

  ' Workbooks("Книга1").Sheets("Лист1").Paste

End Sub

' Копируем итоговый отчет в Книгу для отправки
Sub copyPLToSend()
  Dim TemplatesFile As String

  Application.StatusBar = "Копирование..."

  ' Есть 2 варианта: "Ежедневный отчет по продажам.xlsx" и "Ежедневный отчет.xlsx"
  If Dir(ThisWorkbook.Path + "\Templates\" + "PL-файл.xlsx") <> "" Then
    ' Открываем шаблон Templates\Ежедневный отчет по продажам
    TemplatesFileName = "PL-файл"
  End If
              
  ' Открываем шаблон Протокола из C:\Users\...\Documents\#VBA\DB_Result\Templates
  Workbooks.Open (ThisWorkbook.Path + "\Templates\" + TemplatesFileName + ".xlsx")
           
  ' Переходим на PL
  ThisWorkbook.Sheets("PL").Activate

  ' Обновляем список получателей
  ThisWorkbook.Sheets("PL").Cells(rowByValue(ThisWorkbook.Name, "PL", "Список получателей:", 100, 100), ColumnByValue(ThisWorkbook.Name, "PL", "Список получателей:", 100, 100) + 2).Value = _
    getFromAddrBook("УДО2,УДО3,УДО4,УДО5,НОРПиКО1,НОРПиКО2,НОРПиКО3,НОРПиКО4,НОРПиКО5", 2)

  ' Имя нового файла
  FilePLName = "PL месяц " + CStr(ThisWorkbook.Sheets("PL").Range("I2").Value) + ".xlsx"
  Workbooks(TemplatesFileName).SaveAs FileName:=ThisWorkbook.Path + "\Out\" + FilePLName, FileFormat:=xlOpenXMLWorkbook, createBackUp:=False
  ThisWorkbook.Sheets("PL").Range("Q3").Value = ThisWorkbook.Path + "\Out\" + FilePLName
            
  ' *** Копирование данных ***
 
  ' Заголовок
  i = 2
    For j = 1 To 9
      ThisWorkbook.Sheets("PL").Cells(i, j).Copy Destination:=Workbooks(FilePLName).Sheets("Лист1").Cells(i, j)
    Next j
  
  ' Форма PL1
  For i = 6 To 11
    For j = 1 To 9
      ThisWorkbook.Sheets("PL").Cells(i, j).Copy Destination:=Workbooks(FilePLName).Sheets("Лист1").Cells(i, j)
    Next j
  Next i

  ' Форма PL2
  For i = 17 To 52
    For j = 1 To 9
      ThisWorkbook.Sheets("PL").Cells(i, j).Copy Destination:=Workbooks(FilePLName).Sheets("Лист1").Cells(i, j)
    Next j
  Next i
  
  ' ***
                    
  ' Закрытие файла
  Workbooks(FilePLName).Close SaveChanges:=True

  ' Копирование завершено
  Application.StatusBar = "Скопировано!"
  Application.StatusBar = ""


End Sub


' Отчетность от Котельниковой (по запросу) "PL Розница в разрезе ДО Тюмень X кв 20ГГ" - файл запрашивается у Котельниковой: "Расшифровку опер.рез. по статьям  и резервов в разрезе ДО, клиентов и продуктов, а также расш.ФСА по ДО"
Sub PL_Розница_в_разрезе_ДО()
  
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
    ThisWorkbook.Sheets("PL").Activate

    ' Проверка формы отчета
    ' CheckFormatReportResult = CheckFormatReport(ReportName_String, "Оглавление ", 16, Date)
    ' If CheckFormatReportResult = "OK" Then
    If True Then
      
      ' Очищаем таблицу
      Call Очистить_PL3_ЛистPl
      
      ' Обработка Листа открытой книги PL
      ' Ячейка "Названия строк"
      SheetsNameVar = "PL резервы ФизЛиц"
      Обработать_до_строки = "Рег.Сеть Связь-Банк"
      ' Обработать_до_строки = "Тюменский ОперОфис1" ' ДО "Нижневартовский" Тюменский ОперОфис1
      
      rowCount = rowByValue(ReportName_String, SheetsNameVar, "РБ (202)", 1000, 10) + 3
      Do While InStr(Workbooks(ReportName_String).Sheets(SheetsNameVar).Cells(rowCount, 1).Value, Обработать_до_строки) = 0 ' "Рег.Сеть Связь-Банк"
      
          ' --- Обработка строк ---
          If (InStr(Workbooks(ReportName_String).Sheets(SheetsNameVar).Cells(rowCount, 1).Value, "Резервы (4398)") = 0) And _
               (InStr(Workbooks(ReportName_String).Sheets(SheetsNameVar).Cells(rowCount, 1).Value, "ОперОфис1") = 0) And _
                 (InStr(Workbooks(ReportName_String).Sheets(SheetsNameVar).Cells(rowCount, 1).Value, "(") <> 0) Then
            
            ' Обрабатываем строку, в которой есть "("
            Учтено = False
            
            ' If False Then
            
            ' Ипотечное кредитование (интегрируемые банки) (2514)
            Строка_резерва = "Ипотечное кредитование (интегрируемые банки) (2514)"
            If (InStr(Workbooks(ReportName_String).Sheets(SheetsNameVar).Cells(rowCount, 1).Value, Строка_резерва) <> 0) And (Учтено = False) Then
              ' Заносим на Лист "PL" в форму PL3
              Call SetDataIn_PL3_ЛистPl(ReportName_String, SheetsNameVar, Строка_резерва, rowCount)
              Учтено = True
            End If
            
            ' Военная ипотека (интегрируемые банки) (4444)
            Строка_резерва = "Военная ипотека (интегрируемые банки) (4444)"
            If (InStr(Workbooks(ReportName_String).Sheets(SheetsNameVar).Cells(rowCount, 1).Value, Строка_резерва) <> 0) And (Учтено = False) Then
              ' Заносим на Лист "PL" в форму PL3
              Call SetDataIn_PL3_ЛистPl(ReportName_String, SheetsNameVar, Строка_резерва, rowCount)
              Учтено = True
            End If
            
            ' Кредиты ФЛ (интегрируемые банки) (2505)
            Строка_резерва = "Кредиты ФЛ (интегрируемые банки) (2505)"
            If (InStr(Workbooks(ReportName_String).Sheets(SheetsNameVar).Cells(rowCount, 1).Value, Строка_резерва) <> 0) And (Учтено = False) Then
              ' Заносим на Лист "PL" в форму PL3
              Call SetDataIn_PL3_ЛистPl(ReportName_String, SheetsNameVar, Строка_резерва, rowCount)
              Учтено = True
            End If
            
            ' Автокредитование (интегрируемые банки) (2991)
            Строка_резерва = "Автокредитование (интегрируемые банки) (2991)"
            If (InStr(Workbooks(ReportName_String).Sheets(SheetsNameVar).Cells(rowCount, 1).Value, Строка_резерва) <> 0) And (Учтено = False) Then
              ' Заносим на Лист "PL" в форму PL3
              Call SetDataIn_PL3_ЛистPl(ReportName_String, SheetsNameVar, Строка_резерва, rowCount)
              Учтено = True
            End If
            
            ' Прочее (интегрируемые банки)
            Строка_резерва = "интегрируемые банки"
            If (InStr(Workbooks(ReportName_String).Sheets(SheetsNameVar).Cells(rowCount, 1).Value, Строка_резерва) <> 0) And (Учтено = False) Then
              ' Заносим на Лист "PL" в форму PL3
              Call SetDataIn_PL3_ЛистPl(ReportName_String, SheetsNameVar, "Прочее (интегрируемые банки)", rowCount)
              Учтено = True
            End If
            
            ' Военная ипотека (вторичный рынок) (3855)
            Строка_резерва = "Военная ипотека (вторичный рынок) (3855)"
            If (InStr(Workbooks(ReportName_String).Sheets(SheetsNameVar).Cells(rowCount, 1).Value, Строка_резерва) <> 0) And (Учтено = False) Then
              ' Заносим на Лист "PL" в форму PL3
              Call SetDataIn_PL3_ЛистPl(ReportName_String, SheetsNameVar, Строка_резерва, rowCount)
              Учтено = True
            End If
            
            ' Гражданская Ипотека (Гос программа 2020) (4445)
            Строка_резерва = "Гражданская Ипотека (Гос программа 2020) (4445)"
            If (InStr(Workbooks(ReportName_String).Sheets(SheetsNameVar).Cells(rowCount, 1).Value, Строка_резерва) <> 0) And (Учтено = False) Then
              ' Заносим на Лист "PL" в форму PL3
              Call SetDataIn_PL3_ЛистPl(ReportName_String, SheetsNameVar, Строка_резерва, rowCount)
              Учтено = True
            End If
            
            ' Семейная ипотека (3856)
            Строка_резерва = "Семейная ипотека (3856)"
            If (InStr(Workbooks(ReportName_String).Sheets(SheetsNameVar).Cells(rowCount, 1).Value, Строка_резерва) <> 0) And (Учтено = False) Then
              ' Заносим на Лист "PL" в форму PL3
              Call SetDataIn_PL3_ЛистPl(ReportName_String, SheetsNameVar, Строка_резерва, rowCount)
              Учтено = True
            End If
            
            ' Рефинансирование военной ипотеки (3753)
            Строка_резерва = "Рефинансирование военной ипотеки (3753)"
            If (InStr(Workbooks(ReportName_String).Sheets(SheetsNameVar).Cells(rowCount, 1).Value, Строка_резерва) <> 0) And (Учтено = False) Then
              ' Заносим на Лист "PL" в форму PL3
              Call SetDataIn_PL3_ЛистPl(ReportName_String, SheetsNameVar, Строка_резерва, rowCount)
              Учтено = True
            End If
            
            ' Ипотека (классическая) (61)
            Строка_резерва = "Ипотека (классическая) (61)"
            If (InStr(Workbooks(ReportName_String).Sheets(SheetsNameVar).Cells(rowCount, 1).Value, Строка_резерва) <> 0) And (Учтено = False) Then
              ' Заносим на Лист "PL" в форму PL3
              Call SetDataIn_PL3_ЛистPl(ReportName_String, SheetsNameVar, Строка_резерва, rowCount)
              Учтено = True
            End If
            
            ' Новостройка (1418)
            Строка_резерва = "Новостройка (1418)"
            If (InStr(Workbooks(ReportName_String).Sheets(SheetsNameVar).Cells(rowCount, 1).Value, Строка_резерва) <> 0) And (Учтено = False) Then
              ' Заносим на Лист "PL" в форму PL3
              Call SetDataIn_PL3_ЛистPl(ReportName_String, SheetsNameVar, Строка_резерва, rowCount)
              Учтено = True
            End If
            
            ' Вторичный рынок (1308)
            Строка_резерва = "Вторичный рынок (1308)"
            If (InStr(Workbooks(ReportName_String).Sheets(SheetsNameVar).Cells(rowCount, 1).Value, Строка_резерва) <> 0) And (Учтено = False) Then
              ' Заносим на Лист "PL" в форму PL3
              Call SetDataIn_PL3_ЛистPl(ReportName_String, SheetsNameVar, Строка_резерва, rowCount)
              Учтено = True
            End If
            
            ' Зарплатная карта (все программы поиск подстроки)
            Строка_резерва = "Зарплатная карта"
            If (InStr(Workbooks(ReportName_String).Sheets(SheetsNameVar).Cells(rowCount, 1).Value, Строка_резерва) <> 0) And (Учтено = False) Then
              ' Заносим на Лист "PL" в форму PL3
              Call SetDataIn_PL3_ЛистPl(ReportName_String, SheetsNameVar, "Зарплатная карта (все программы поиск подстроки)", rowCount)
              Учтено = True
            End If
            
            ' Дебетовая карта (все программы поиск подстроки)
            Строка_резерва = "Дебетовая карта"
            If (InStr(Workbooks(ReportName_String).Sheets(SheetsNameVar).Cells(rowCount, 1).Value, Строка_резерва) <> 0) And (Учтено = False) Then
              ' Заносим на Лист "PL" в форму PL3
              Call SetDataIn_PL3_ЛистPl(ReportName_String, SheetsNameVar, "Дебетовая карта (все программы поиск подстроки)", rowCount)
              Учтено = True
            End If
            
            ' Платежная карта  (все программы поиск подстроки)
            Строка_резерва = "Платежная карта"
            If (InStr(Workbooks(ReportName_String).Sheets(SheetsNameVar).Cells(rowCount, 1).Value, Строка_резерва) <> 0) And (Учтено = False) Then
              ' Заносим на Лист "PL" в форму PL3
              Call SetDataIn_PL3_ЛистPl(ReportName_String, SheetsNameVar, "Платежная карта  (все программы поиск подстроки)", rowCount)
              Учтено = True
            End If
            
            ' Кредитная карта (все программы поиск подстроки)
            Строка_резерва = "Кредитная карта"
            If (InStr(Workbooks(ReportName_String).Sheets(SheetsNameVar).Cells(rowCount, 1).Value, Строка_резерва) <> 0) And (Учтено = False) Then
              ' Заносим на Лист "PL" в форму PL3
              Call SetDataIn_PL3_ЛистPl(ReportName_String, SheetsNameVar, "Кредитная карта (все программы поиск подстроки)", rowCount)
              Учтено = True
            End If
            
            ' ТурбоДеньги (3080)
            Строка_резерва = "ТурбоДеньги (3080)"
            If (InStr(Workbooks(ReportName_String).Sheets(SheetsNameVar).Cells(rowCount, 1).Value, Строка_резерва) <> 0) And (Учтено = False) Then
              ' Заносим на Лист "PL" в форму PL3
              Call SetDataIn_PL3_ЛистPl(ReportName_String, SheetsNameVar, Строка_резерва, rowCount)
              Учтено = True
            End If
            
            ' Pre_approved (2222)
            Строка_резерва = "Pre_approved (2222)"
            If (InStr(Workbooks(ReportName_String).Sheets(SheetsNameVar).Cells(rowCount, 1).Value, Строка_резерва) <> 0) And (Учтено = False) Then
              ' Заносим на Лист "PL" в форму PL3
              Call SetDataIn_PL3_ЛистPl(ReportName_String, SheetsNameVar, Строка_резерва, rowCount)
              Учтено = True
            End If
            
            ' Зеленые компании (2476)
            Строка_резерва = "Зеленые компании (2476)"
            If (InStr(Workbooks(ReportName_String).Sheets(SheetsNameVar).Cells(rowCount, 1).Value, Строка_резерва) <> 0) And (Учтено = False) Then
              ' Заносим на Лист "PL" в форму PL3
              Call SetDataIn_PL3_ЛистPl(ReportName_String, SheetsNameVar, Строка_резерва, rowCount)
              Учтено = True
            End If
            
            ' Открытый рынок (3486)
            Строка_резерва = "Открытый рынок (3486)"
            If (InStr(Workbooks(ReportName_String).Sheets(SheetsNameVar).Cells(rowCount, 1).Value, Строка_резерва) <> 0) And (Учтено = False) Then
              ' Заносим на Лист "PL" в форму PL3
              Call SetDataIn_PL3_ЛистPl(ReportName_String, SheetsNameVar, Строка_резерва, rowCount)
              Учтено = True
            End If
            
            ' End If
            
            ' Прочее (все остальное)
            Строка_резерва = ""
            If (Учтено = False) Then
              ' Заносим на Лист "PL" в форму PL3
              Call SetDataIn_PL3_ЛистPl(ReportName_String, SheetsNameVar, "Прочее (все остальное)", rowCount)
              Учтено = True
            End If
            
          End If
        
          ' Следующая запись
          rowCount = rowCount + 1
          Application.StatusBar = CStr(rowCount) + "..."
          DoEventsInterval (rowCount)
      Loop
      
      ' Заносим переменные
      ThisWorkbook.Sheets("PL").Range("B83").Value = "Обработано до строки №" + CStr(rowCount)
      
      Application.StatusBar = "Завершение..."
      
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
    ThisWorkbook.Sheets("PL").Range("A55").Select

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


' Очищаем таблицу
Sub Очистить_PL3_ЛистPl()
      
      ' PL 3
        For i = 58 To 80
        
          If ThisWorkbook.Sheets("PL").Cells(i, 2).Value <> "" Then
            
            ' Создание
            ThisWorkbook.Sheets("PL").Cells(i, 6).Value = 0
            ' Формат ячейки
            ThisWorkbook.Sheets("PL").Cells(i, 6).NumberFormat = "#,##0"
            ThisWorkbook.Sheets("PL").Cells(i, 6).Font.Bold = True
                
            ' Восстановление
            ThisWorkbook.Sheets("PL").Cells(i, 8).Value = 0
            ' Формат ячейки
            ThisWorkbook.Sheets("PL").Cells(i, 8).NumberFormat = "#,##0"
            ThisWorkbook.Sheets("PL").Cells(i, 8).Font.Bold = True

          End If
          
        Next i
      
End Sub

' Заносим на Лист "PL" в форму PL3
Sub SetDataIn_PL3_ЛистPl(In_ReportName_String, In_SheetsNameVar, In_Name_PL3, In_rowCount_report_PL)
  
  ' In_rowCount_report_PL - это номер записи на листе отчета от Котельниковой
  ' In_Name_PL3 - наименование статьи резерва на Листе PL в DB_Result
  
  ' Начало блока
  row_begin_ЛистPL = 58
  ' Конец блока
  row_end_ЛистPL = 82
  
  запись_найдена = False
  rowCount = row_begin_ЛистPL
  Do While (rowCount <= row_end_ЛистPL) And (запись_найдена = False) '
      
          
      
          ' --- Обработка строк ---
          If (InStr(ThisWorkbook.Sheets("PL").Cells(rowCount, 2).Value, In_Name_PL3) <> 0) Then
            
            ' t = Workbooks(In_ReportName_String).Sheets(In_SheetsNameVar).Cells(In_rowCount_report_PL, 6).Value
            
            ' Если строка со знаком "-" в столбце 6 "2020 Итог" - создание резерва
            If Workbooks(In_ReportName_String).Sheets(In_SheetsNameVar).Cells(In_rowCount_report_PL, 6).Value < 0 Then
              ThisWorkbook.Sheets("PL").Cells(rowCount, 6).Value = ThisWorkbook.Sheets("PL").Cells(rowCount, 6).Value + (Workbooks(In_ReportName_String).Sheets(In_SheetsNameVar).Cells(In_rowCount_report_PL, 6).Value / 1000)
            End If
            
            ' Если строка со знаком "+" - восстановление
            If Workbooks(In_ReportName_String).Sheets(In_SheetsNameVar).Cells(In_rowCount_report_PL, 6).Value > 0 Then
              ThisWorkbook.Sheets("PL").Cells(rowCount, 8).Value = ThisWorkbook.Sheets("PL").Cells(rowCount, 8).Value + (Workbooks(In_ReportName_String).Sheets(In_SheetsNameVar).Cells(In_rowCount_report_PL, 6).Value / 1000)
            End If
            
            запись_найдена = True
          End If
        
    ' Следующая запись
    rowCount = rowCount + 1
    ' Application.StatusBar = CStr(rowCount) + "..."
    ' DoEventsInterval (rowCount)
  Loop

End Sub
