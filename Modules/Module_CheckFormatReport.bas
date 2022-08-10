Attribute VB_Name = "Module_CheckFormatReport"
' Проверка формата открываемых файлов (перенесено сюда из Module_DLL)

' 34. Проверка формата открываемых отчетов
Function CheckFormatReport(In_Workbooks, In_Sheets, In_TypeReport, In_Date) As String

Dim find_office_1, find_office_2, find_office_3, find_office_4, find_office_5, stop_process As Boolean

  CheckFormatReport = ""

  ' 1 - DB: в на листе "Оглавление" в A1="Отчет по состоянию на 04.08.2020"
  If In_TypeReport = 1 Then
            
    ' Если на листе F1 в "A1" находится "Тип карт", то считаем, что формат верный
    If Sheets_Exist2(In_Workbooks, In_Sheets) = True Then
      '
      If (InStr(Workbooks(In_Workbooks).Sheets(In_Sheets).Range("A1").Value, "Отчет по состоянию на") <> 0) Then
        
        ' Проверяем дату отчета
        If True Then ' CDate(Workbooks(In_Workbooks).Sheets(In_Sheets).Range("C1").Value) = In_Date
          ' Значит отчет верный и правильная дата
          CheckFormatReport = "OK"
        Else
          CheckFormatReport = "неверная дата в отчёте"
        End If
      Else
        ' Если в книге нет в заданной ячейки заданного значения
        CheckFormatReport = "выбран неверный формат"
      End If
    Else
        ' Если в Книге нет Листа с таким именем
        CheckFormatReport = "выбран неверный формат"
    End If
  End If

  
  ' 2 - ML
  
  ' 3 - Объем кредитного портфеля с изменениями за период
  If In_TypeReport = 3 Then
    
    ' Если в A1 "Объем кредитного портфеля с учетом изменений за период c 01.01.2020 по 17.03.2020"
    If Sheets_Exist2(In_Workbooks, In_Sheets) = True Then
    
      If InStr(Workbooks(In_Workbooks).Sheets(In_Sheets).Range("A1").Value, "Объем кредитного портфеля с учетом изменений за период") <> 0 Then
      
        ' Проверяем дату отчета - в A2  "За период с 01.01.2020 по 12.03.2020"
        If (CDate(Mid(Workbooks(In_Workbooks).Sheets(In_Sheets).Range("A1").Value, 58, 10)) = YearStartDate(In_Date)) And (CDate(Mid(Workbooks(In_Workbooks).Sheets(In_Sheets).Range("A1").Value, 72, 10)) = In_Date) Then
          
          
          find_office_1 = False
          find_office_2 = False
          find_office_3 = False
          find_office_4 = False
          find_office_5 = False
          stop_process = False
          
          RecCount = 8
          Do While (Workbooks(In_Workbooks).Sheets(In_Sheets).Cells(RecCount, 2).Value <> "") And (stop_process = False)
            
            ' Проверяем офис
            If InStr(Workbooks(In_Workbooks).Sheets(In_Sheets).Cells(RecCount, 2).Value, "Тюменский") <> 0 Then
              find_office_1 = True
            End If
            ' Проверяем офис 2
            If InStr(Workbooks(In_Workbooks).Sheets(In_Sheets).Cells(RecCount, 2).Value, "Сургутский") <> 0 Then
              find_office_2 = True
            End If
            ' Проверяем офис 3
            If InStr(Workbooks(In_Workbooks).Sheets(In_Sheets).Cells(RecCount, 2).Value, "Нижневартовский") <> 0 Then
              find_office_3 = True
            End If
            ' Проверяем офис 4
            If InStr(Workbooks(In_Workbooks).Sheets(In_Sheets).Cells(RecCount, 2).Value, "Новоуренгойский") <> 0 Then
              find_office_4 = True
            End If
            ' Проверяем офис 5
            If InStr(Workbooks(In_Workbooks).Sheets(In_Sheets).Cells(RecCount, 2).Value, "Тарко-Сале") <> 0 Then
              find_office_5 = True
            End If
            
            ' Если нашли все 5, то останавливаем процесс
            If (find_office_1 = True) And (find_office_2 = True) And (find_office_3 = True) And (find_office_4 = True) And (find_office_5 = True) Then
              stop_process = True
            End If
            
            RecCount = RecCount + 1
          Loop
          
          ' Здесь проверяем наличе 5-ти офисов
          If (find_office_1 = True) And (find_office_2 = True) And (find_office_3 = True) And (find_office_4 = True) And (find_office_5 = True) Then
              ' Значит отчет верный и правильная дата
              CheckFormatReport = "OK"
            Else
              
              ' Значит в отчете нет всех офисов
              CheckFormatReport = "Не все офисы в отчете!"
            
          End If
          
        Else
          
          CheckFormatReport = "неверная дата в отчёте"
        
        End If
      Else
        ' Если в книге нет в заданной ячейки заданного значения
        CheckFormatReport = "выбран неверный формат"
      End If
    Else
        ' Если в Книге нет Листа с таким именем
        CheckFormatReport = "выбран неверный формат"
    End If
      
  End If
  

  ' 4 - Договора к закрытию в течение 30 дней
 
  ' 5 - Отчет об эмиссии банковских карт доп. офисами филиала
  If In_TypeReport = 5 Then
    
    ' Если в A1 "Отчет об эмиссии банковских карт доп. офисами филиала"
    If Sheets_Exist2(In_Workbooks, In_Sheets) = True Then
      If (Workbooks(In_Workbooks).Sheets(In_Sheets).Range("A1").Value = "Отчет об эмиссии банковских карт доп. офисами филиала") Then
        ' Проверяем дату отчета - в A2  "За период с 01.01.2020 по 12.03.2020"
        If (CDate(Mid(Workbooks(In_Workbooks).Sheets(In_Sheets).Range("A2").Value, 28, 10)) = In_Date) And (Mid(Workbooks(In_Workbooks).Sheets(In_Sheets).Range("A2").Value, 14, 10) = "01.01." + Mid(Workbooks(In_Workbooks).Sheets(In_Sheets).Range("A2").Value, 34, 4)) Then
          ' Значит отчет верный и правильная дата
          CheckFormatReport = "OK"
        Else
          CheckFormatReport = "неверная дата в отчёте"
        End If
      Else
        ' Если в книге нет в заданной ячейки заданного значения
        CheckFormatReport = "выбран неверный формат"
      End If
    Else
        ' Если в Книге нет Листа с таким именем
        CheckFormatReport = "выбран неверный формат"
    End If
      
  End If
  
  ' 6 - Кредитный портфель в аналитике для физ.лиц (по доп. офисам)
  If In_TypeReport = 6 Then
    ' Если в A1 "Кредитный портфель в аналитике на "
    If Sheets_Exist2(In_Workbooks, In_Sheets) = True Then
      If (Workbooks(In_Workbooks).Sheets(In_Sheets).Range("A1").Value = "Кредитный портфель в аналитике на ") Then
        ' Проверяем дату отчета
        If True Then ' CDate(Workbooks(In_Workbooks).Sheets(In_Sheets).Range("C1").Value) = In_Date Then
          ' Значит отчет верный и правильная дата
          CheckFormatReport = "OK"
        Else
          CheckFormatReport = "неверная дата в отчёте"
        End If
      Else
        ' Если в книге нет в заданной ячейки заданного значения
        CheckFormatReport = "выбран неверный формат"
      End If
    Else
        ' Если в Книге нет Листа с таким именем
        CheckFormatReport = "выбран неверный формат"
    End If
  End If
  ' --- 6 - Кредитный портфель в аналитике для физ.лиц (по доп. офисам)
  
  ' 7 - Отчет Capacity
  If In_TypeReport = 7 Then
    ' Если в A1 "Кредитный портфель в аналитике на "
    If Sheets_Exist2(In_Workbooks, In_Sheets) = True Then
      
      ' Иногда бывает пустая строка
      If (Workbooks(In_Workbooks).Sheets(In_Sheets).Range("B6").Value = "Кол-во клиентов") Or (Workbooks(In_Workbooks).Sheets(In_Sheets).Range("B7").Value = "Кол-во клиентов") Then
        ' Проверяем дату отчета
        If True Then ' CDate(Workbooks(In_Workbooks).Sheets(In_Sheets).Range("C1").Value) = In_Date
          ' Значит отчет верный и правильная дата
          CheckFormatReport = "OK"
        Else
          CheckFormatReport = "неверная дата в отчёте"
        End If
      Else
        ' Если в книге нет в заданной ячейки заданного значения
        CheckFormatReport = "выбран неверный формат"
      End If
    Else
        ' Если в Книге нет Листа с таким именем
        CheckFormatReport = "выбран неверный формат"
    End If
  End If
  
  ' 8 - Отчет План/Факт за ДД.ММ.ГГГГ по продуктам ИСЖ_НСЖ
  If In_TypeReport = 8 Then
    
    ' Если в E2 "Отчет обновлен на 19.03.2020 за 18.03.2020 (полный день)"
    If Sheets_Exist2(In_Workbooks, In_Sheets) = True Then
      '
      If InStr(Workbooks(In_Workbooks).Sheets(In_Sheets).Range("E2").Value, "Отчет обновлен на ") <> 0 Then
        
        ' Проверяем дату отчета
        If True Then ' CDate(Workbooks(In_Workbooks).Sheets(In_Sheets).Range("C1").Value) = In_Date
          ' Значит отчет верный и правильная дата
          CheckFormatReport = "OK"
        Else
          CheckFormatReport = "неверная дата в отчёте"
        End If
      Else
        ' Если в книге нет в заданной ячейки заданного значения
        CheckFormatReport = "выбран неверный формат"
      End If
    Else
        ' Если в Книге нет Листа с таким именем
        CheckFormatReport = "выбран неверный формат"
    End If
  End If
  
  ' 9 - Отчетность ЕСУП_итог 15.03.2020 (от Карэна)
  If In_TypeReport = 9 Then
    
    ' Если в C2 находится текст: Данный отчет позволяет проанализировать количество выкладываемых документов  в папки по Регулярным управленческим процедурам  в "Файловом хранилище руководителя"
    If Sheets_Exist2(In_Workbooks, In_Sheets) = True Then
      '
      If InStr(Workbooks(In_Workbooks).Sheets(In_Sheets).Range("C2").Value, "Данный отчет позволяет проанализировать количество выкладываемых документов") <> 0 Then
        
        ' Проверяем дату отчета
        If True Then ' CDate(Workbooks(In_Workbooks).Sheets(In_Sheets).Range("C1").Value) = In_Date
          ' Значит отчет верный и правильная дата
          CheckFormatReport = "OK"
        Else
          CheckFormatReport = "неверная дата в отчёте"
        End If
      Else
        ' Если в книге нет в заданной ячейки заданного значения
        CheckFormatReport = "выбран неверный формат"
      End If
    Else
        ' Если в Книге нет Листа с таким именем
        CheckFormatReport = "выбран неверный формат"
    End If
  End If
  
  
  ' 10 - "Декомпозиция планов продаж_Xкв.20YY" + " возможные дополнения"
  If In_TypeReport = 10 Then
    
    ' Если в "B1" находится "ПК МРК/МК, тыс. руб."
    If Sheets_Exist2(In_Workbooks, In_Sheets) = True Then
      '
      ' If InStr(Workbooks(In_Workbooks).Sheets(In_Sheets).Range("B1").Value, "ПК МРК/МК, тыс. руб.") <> 0 Then
      ' C 2021 года "ПК МРК/МК (Офис+ИБ), тыс. руб."
      If InStr(Workbooks(In_Workbooks).Sheets(In_Sheets).Range("B1").Value, "ПК МРК/МК (Офис+ИБ), тыс. руб.") <> 0 Then
      
        ' Проверяем дату отчета
        If True Then ' CDate(Workbooks(In_Workbooks).Sheets(In_Sheets).Range("C1").Value) = In_Date
          ' Значит отчет верный и правильная дата
          CheckFormatReport = "OK"
        Else
          CheckFormatReport = "неверная дата в отчёте"
        End If
      Else
        ' Если в книге нет в заданной ячейки заданного значения
        CheckFormatReport = "выбран неверный формат"
      End If
    Else
        ' Если в Книге нет Листа с таким именем
        CheckFormatReport = "выбран неверный формат"
    End If
  End If
  
  ' 11 - "Выданные кредиты по прескриннингу"
  If In_TypeReport = 11 Then
    
    ' Дата начала периода отчета
    ' 01.01.2020 => "01 января 2020"
    dateInStringFormat1 = "01 января " + CStr(Year(In_Date)) + " г."
    
    ' 01.01.2020 => "01 января 2020"
    dateInStringFormat2 = CStr(Day(In_Date)) + " " + ИмяМесяца2(In_Date) + " " + CStr(Year(In_Date)) + " г."
        
    ' Если в "A1" находится "Выданные кредиты по прескриннингу за период с 01 января 2020 г. по 06 июля 2020 г."
    If Sheets_Exist2(In_Workbooks, In_Sheets) = True Then
      '
      If (InStr(Workbooks(In_Workbooks).Sheets(In_Sheets).Range("A1").Value, dateInStringFormat1) <> 0) And ((InStr(Workbooks(In_Workbooks).Sheets(In_Sheets).Range("A1").Value, dateInStringFormat2) <> 0)) Then
        
        ' Проверяем дату отчета
        If True Then ' CDate(Workbooks(In_Workbooks).Sheets(In_Sheets).Range("C1").Value) = In_Date
          ' Значит отчет верный и правильная дата
          CheckFormatReport = "OK"
        Else
          CheckFormatReport = "неверная дата в отчёте"
        End If
      Else
        ' Если в книге нет в заданной ячейки заданного значения
        CheckFormatReport = "выбран неверный формат"
      End If
    Else
        ' Если в Книге нет Листа с таким именем
        CheckFormatReport = "выбран неверный формат"
    End If
  End If
  
  ' 12 - Cards_emisssion_ДД_ММ_ГГ_(2019)_2
  If In_TypeReport = 12 Then
            
    ' Если на листе F1 в "A1" находится "Тип карт", то считаем, что формат верный
    If Sheets_Exist2(In_Workbooks, In_Sheets) = True Then
      '
      If (InStr(Workbooks(In_Workbooks).Sheets(In_Sheets).Range("A1").Value, "Тип карт") <> 0) Then
        
        ' Проверяем дату отчета
        If True Then ' CDate(Workbooks(In_Workbooks).Sheets(In_Sheets).Range("C1").Value) = In_Date
          ' Значит отчет верный и правильная дата
          CheckFormatReport = "OK"
        Else
          CheckFormatReport = "неверная дата в отчёте"
        End If
      Else
        ' Если в книге нет в заданной ячейки заданного значения
        CheckFormatReport = "выбран неверный формат"
      End If
    Else
        ' Если в Книге нет Листа с таким именем
        CheckFormatReport = "выбран неверный формат"
    End If
  End If
  
  ' 14 - Выписка из Протоколов (Исполнено/Не исполнено)
  If In_TypeReport = 14 Then
            
    ' Если на листе F1 в "A1" находится "Тип карт", то считаем, что формат верный
    If Sheets_Exist2(In_Workbooks, In_Sheets) = True Then
      '
      If (InStr(Workbooks(In_Workbooks).Sheets(In_Sheets).Range("C1").Value, "Выписка из Протоколов") <> 0) Then
        
        ' Проверяем дату отчета
        If True Then ' CDate(Workbooks(In_Workbooks).Sheets(In_Sheets).Range("C1").Value) = In_Date
          ' Значит отчет верный и правильная дата
          CheckFormatReport = "OK"
        Else
          CheckFormatReport = "неверная дата в отчёте"
        End If
      Else
        ' Если в книге нет в заданной ячейки заданного значения
        CheckFormatReport = "выбран неверный формат"
      End If
    Else
        ' Если в Книге нет Листа с таким именем
        CheckFormatReport = "выбран неверный формат"
    End If
  End If
  
  ' 15 - Воронка
  If In_TypeReport = 15 Then
            
    ' Если на листе F1 в "A1" находится "Тип карт", то считаем, что формат верный
    If Sheets_Exist2(In_Workbooks, In_Sheets) = True Then
      '
      If (InStr(Workbooks(In_Workbooks).Sheets(In_Sheets).Range("A1").Value, "Flag_action") <> 0) Then
        
        ' Проверяем дату отчета
        If True Then ' CDate(Workbooks(In_Workbooks).Sheets(In_Sheets).Range("C1").Value) = In_Date
          ' Значит отчет верный и правильная дата
          CheckFormatReport = "OK"
        Else
          CheckFormatReport = "неверная дата в отчёте"
        End If
      Else
        ' Если в книге нет в заданной ячейки заданного значения
        CheckFormatReport = "выбран неверный формат"
      End If
    Else
        ' Если в Книге нет Листа с таким именем
        CheckFormatReport = "выбран неверный формат"
    End If
  End If
  
  ' 16 - PL
  If In_TypeReport = 16 Then
            
    ' Если на листе F1 в "A1" находится "Тип карт", то считаем, что формат верный
    If Sheets_Exist2(In_Workbooks, In_Sheets) = True Then
      '
      If (InStr(Workbooks(In_Workbooks).Sheets(In_Sheets).Range("B2").Value, "Оглавление") <> 0) Then
        
        ' Проверяем дату отчета
        If True Then ' CDate(Workbooks(In_Workbooks).Sheets(In_Sheets).Range("C1").Value) = In_Date
          ' Значит отчет верный и правильная дата
          CheckFormatReport = "OK"
        Else
          CheckFormatReport = "неверная дата в отчёте"
        End If
      Else
        ' Если в книге нет в заданной ячейки заданного значения
        CheckFormatReport = "выбран неверный формат"
      End If
    Else
        ' Если в Книге нет Листа с таким именем
        CheckFormatReport = "выбран неверный формат"
    End If
  End If
  
  ' 17 - Обработка отчета http://isrb.psbnk.msk.ru/inf/6601/6622/otchet_zp_org/
  If In_TypeReport = 17 Then
            
    ' Если на листе "сводная по ТП" в "A4" находится "бакет по численности орг", то считаем, что формат верный
    If Sheets_Exist2(In_Workbooks, In_Sheets) = True Then
      '
      ' If (InStr(Workbooks(In_Workbooks).Sheets(In_Sheets).Range("A4").Value, "бакет по численности орг") <> 0) Then
      
      If (InStr(Workbooks(In_Workbooks).Sheets(In_Sheets).Range("A4").Value, "категория") <> 0) Then
        
        ' Проверяем дату отчета
        If True Then ' CDate(Workbooks(In_Workbooks).Sheets(In_Sheets).Range("C1").Value) = In_Date
          ' Значит отчет верный и правильная дата
          CheckFormatReport = "OK"
        Else
          CheckFormatReport = "неверная дата в отчёте"
        End If
      Else
        ' Если в книге нет в заданной ячейки заданного значения
        CheckFormatReport = "выбран неверный формат"
      End If
    Else
        ' Если в Книге нет Листа с таким именем
        CheckFormatReport = "выбран неверный формат"
    End If
  End If
    
  ' 18 - Обработка отчета "Отчётность по входящему потоку с PA" http://isrb.psbnk.msk.ru/inf/6601/6622/ochet_PA/
  If In_TypeReport = 18 Then
            
    ' Если на листе "Вх. поток с РА-ПК (мес)" в "A6" находится "Таргет", то считаем, что формат верный
    If Sheets_Exist2(In_Workbooks, In_Sheets) = True Then
      '
      If (InStr(Workbooks(In_Workbooks).Sheets(In_Sheets).Range("A6").Value, "Таргет") <> 0) Then
        
        ' Проверяем дату отчета
        If True Then ' CDate(Workbooks(In_Workbooks).Sheets(In_Sheets).Range("C1").Value) = In_Date
          ' Значит отчет верный и правильная дата
          CheckFormatReport = "OK"
        Else
          CheckFormatReport = "неверная дата в отчёте"
        End If
      Else
        ' Если в книге нет в заданной ячейки заданного значения
        CheckFormatReport = "выбран неверный формат"
      End If
    Else
        ' Если в Книге нет Листа с таким именем
        CheckFormatReport = "выбран неверный формат"
    End If
  End If
    
  ' 19 - Новый Capacity (Capacity_new)
  If In_TypeReport = 19 Then
            
    ' Если на листе "Вх. поток с РА-ПК (мес)" в "A6" находится "Таргет", то считаем, что формат верный
    If Sheets_Exist2(In_Workbooks, In_Sheets) = True Then
      '
      If (InStr(Workbooks(In_Workbooks).Sheets(In_Sheets).Range("A1").Value, "Гр-КПС") <> 0) Then
        
        ' Проверяем дату отчета
        If True Then ' CDate(Workbooks(In_Workbooks).Sheets(In_Sheets).Range("C1").Value) = In_Date
          ' Значит отчет верный и правильная дата
          CheckFormatReport = "OK"
        Else
          CheckFormatReport = "неверная дата в отчёте"
        End If
      Else
        ' Если в книге нет в заданной ячейки заданного значения
        CheckFormatReport = "выбран неверный формат"
      End If
    Else
        ' Если в Книге нет Листа с таким именем
        CheckFormatReport = "выбран неверный формат"
    End If
  End If
    
  ' 20 - Отчет по ЗП_Nкв_YYYY_DD.MM.YYYY_v1 (1)
  If In_TypeReport = 20 Then
            
    ' Если на листе "Сотр. прод." в "A" находится "Квартал Год", то считаем, что формат верный
    If Sheets_Exist2(In_Workbooks, In_Sheets) = True Then
      '
      If (InStr(Workbooks(In_Workbooks).Sheets(In_Sheets).Range("A1").Value, "Квартал Год") <> 0) Then
        
        ' Проверяем дату отчета
        If True Then ' CDate(Workbooks(In_Workbooks).Sheets(In_Sheets).Range("C1").Value) = In_Date
          ' Значит отчет верный и правильная дата
          CheckFormatReport = "OK"
        Else
          CheckFormatReport = "неверная дата в отчёте"
        End If
      Else
        ' Если в книге нет в заданной ячейки заданного значения
        CheckFormatReport = "выбран неверный формат"
      End If
    Else
        ' Если в Книге нет Листа с таким именем
        CheckFormatReport = "выбран неверный формат"
    End If
  End If
    
  ' 21 - Pipe ЗП
  If In_TypeReport = 21 Then
            
    ' Если на листе "Сотр. прод." в "A" находится "Квартал Год", то считаем, что формат верный
    If Sheets_Exist2(In_Workbooks, In_Sheets) = True Then
      '
      If (InStr(Workbooks(In_Workbooks).Sheets(In_Sheets).Range("A1").Value, "Регион") <> 0) Then
        
        ' Проверяем дату отчета
        If True Then ' CDate(Workbooks(In_Workbooks).Sheets(In_Sheets).Range("C1").Value) = In_Date
          ' Значит отчет верный и правильная дата
          CheckFormatReport = "OK"
        Else
          CheckFormatReport = "неверная дата в отчёте"
        End If
      Else
        ' Если в книге нет в заданной ячейки заданного значения
        CheckFormatReport = "выбран неверный формат"
      End If
    Else
        ' Если в Книге нет Листа с таким именем
        CheckFormatReport = "выбран неверный формат"
    End If
  End If
    
    
End Function

