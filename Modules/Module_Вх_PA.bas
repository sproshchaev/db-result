Attribute VB_Name = "Module_Вх_PA"
' *** Отчётность по входящему потоку с PA (Вх_PA) ***

' *** Глобальные переменные ***
Public dateReport As Date
Public countRowNewLine_Вх_PA As Integer ' Счетчик вывода новых строк на лист "Вх_PA"

Public countКлиенты_с_PA_ПК As Integer  ' Счетчики
Public count_Выдача_РА_ПК As Integer    ' Счетчики

Public countРОО_Клиенты_с_PA_ПК As Integer  ' Счетчики
Public countРОО_Выдача_РА_ПК As Integer     ' Счетчики

Public countКлиенты_с_PA_КК As Integer     ' Счетчики
Public count_Заказ_РА_КК As Integer        ' Счетчики

Public countРОО_Клиенты_с_PA_КК As Integer ' Счетчики
Public countРОО_Заказ_РА_КК As Integer     ' Счетчики

Public Объем_КП_с_изм_за_период_НК As Long
Public Объем_КП_с_изм_за_период_Число_кредитов As Integer
Public Объем_КП_с_изм_за_период_Сумма_кредитов As Double
Public Объем_КП_с_изм_за_период_Виды_кредитов As String
Public ФИО_из_Объем_КП_с_изм_за_период As String



' ***                       ***
  
' Обработка отчета
Sub Отчётность_по_входящему_потоку_с_PA()

' Описание переменных
Dim ReportName_String, officeNameInReport, CheckFormatReportResult As String
Dim i, rowCount As Integer
Dim finishProcess As Boolean
    
  ' Сообщение о необходимости обновления отчета по активам Лист3 N1
  MsgBox ("Перед запуском обработки необходимо обновить отчет: " + ThisWorkbook.Sheets("Лист3").Range("N1").Value + " текущая версия " + Dir(ThisWorkbook.Sheets("Лист3").Range("Q3").Value) + "!")

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
      
    ' Открываем BASE\Clients
    OpenBookInBase ("Clients")
      
    ' Переходим на окно DB
    ThisWorkbook.Sheets("Вх_PA").Activate

    ' Проверка формы отчета
    CheckFormatReportResult = CheckFormatReport(ReportName_String, "Вх. поток с РА-ПК (мес)", 18, Date)
    If CheckFormatReportResult = "OK" Then
    
      ' Строка статуса
      Application.StatusBar = "Подготовка поля отчета..."
    
      ' В O2 устанавливаем Дату
      ThisWorkbook.Sheets("Вх_PA").Range("O2").Value = "Отчетность по входящему потоку с PA " + CStr(Workbooks(ReportName_String).Sheets("Вх. поток с РА-ПК (мес)").Range("B1").Value)
    
      ' Открываем "Файл-отчет:" с Листа3 - для проверки наличия кредитов
      ' Открываем выбранную книгу (UpdateLinks:=0)
      FileName2 = ThisWorkbook.Sheets("Лист3").Range("Q3").Value
      Workbooks.Open FileName2, 0
    
      ' Очистка поля отчета на Листе
      Call clearСontents_Вх_PA

      ' Установка настроек сводной таблицы на листе "Вх. поток с РА-ПК (мес)":
      
      ' 1) Здесь открывается Лист1 где есть расширенный статус PA ПК, но нет МРК
      Call setFilter_Вх_поток_с_РА_ПК_мес(ReportName_String, "Вх. поток с РА-ПК (мес)", "Срез_typeSeg1")
      
      ' 2) Здесь открывается Лист2 где есть усеченный статус PA ПК, но есть МРК
      Call setFilter_Вх_поток_с_РА_ПК_мес(ReportName_String, "CR РА-ПК (мес) Менеджер", "Срез_typeSeg4")
      
      ' 3) Здесь открывается Лист3 где PA КК
      Call setFilter_Вх_поток_с_РА_КК_мес(ReportName_String, "CR РА-KК (мес) Менеджер", "Срез_typeSeg41")
      
      ' Создаем выходную книгу для выгрузки PA ПК
      OutBookName = ThisWorkbook.Path + "\Out\Pre-App_ПК_" + strDDMMYYYY(dateReport) + ".xlsx"
      ThisWorkbook.Sheets("Вх_PA").Range("P3").Value = OutBookName ' Записываем имя файла во вложение
      ' Создать файл
      Call createBook_out_PA2(OutBookName)

      ' Создаем выходную книгу для выгрузки PA КК
      OutBookName2 = ThisWorkbook.Path + "\Out\Pre-App_КК_" + strDDMMYYYY(dateReport) + ".xlsx"
      ThisWorkbook.Sheets("Вх_PA").Range("R3").Value = OutBookName2 ' Записываем имя файла во вложение
      ' Создать файл
      Call createBook_out_PA2(OutBookName2)

      ' Переход на лист
      ThisWorkbook.Sheets("Вх_PA").Activate
      
      ' Строка статуса
      Application.StatusBar = "Определение столбцов на Лист1/Лист2/Лист3..."
      
      ' Определяем поля на Лист1
      column_Лист1_НК = ColumnByValue(ReportName_String, "Лист1", "NK", 100, 100)
      column_Лист1_Офис = ColumnByValue(ReportName_String, "Лист1", "DP4_отчет", 100, 100)
      column_Лист1_Сегмент_детал = ColumnByValue(ReportName_String, "Лист1", "typeSegDetal", 100, 100) ' typeSegDetal
      
      ' Определяем столбцы на Лист2
      column_Лист2_Офис = ColumnByValue(ReportName_String, "Лист2", "DP4_отчет", 100, 100) ' Q
      column_Лист2_Дата_обновления = ColumnByValue(ReportName_String, "Лист2", "load_date", 100, 100) ' Дата обновления
      column_Лист2_Месяц = ColumnByValue(ReportName_String, "Лист2", "month", 100, 100) '
      column_Лист2_НК = ColumnByValue(ReportName_String, "Лист2", "NK", 100, 100) ' НК ритейл
      column_Лист2_CRM_НК = ColumnByValue(ReportName_String, "Лист2", "ybpideal", 100, 100) ' Идентификатор клиента CRM
      column_Лист2_ФИО_МРК = ColumnByValue(ReportName_String, "Лист2", "Manager", 100, 100) ' МРК
      column_Лист2_ТН_МРК = ColumnByValue(ReportName_String, "Лист2", "Табельный_номер", 100, 100) ' табельный номер
      column_Лист2_BIC_RFOFICID = ColumnByValue(ReportName_String, "Лист2", "/BIC/RFOFICID", 100, 100) ' Офис операции
      column_Лист2_Есть_ПК_КК = ColumnByValue(ReportName_String, "Лист2", "type", 100, 100) ' type = 1/0 Наличие типа продукта РА (ПК / КК) в момент операции в офисе
      column_Лист2_Сегмент = ColumnByValue(ReportName_String, "Лист2", "typeSeg", 100, 100) ' тип сегмента
      column_Лист2_potok = ColumnByValue(ReportName_String, "Лист2", "potok", 100, 100) ' potok=1- клиент в офисе с операцией и имеет РА
      column_Лист2_Заявка_PA_KK = ColumnByValue(ReportName_String, "Лист2", "applicCC", 100, 100) ' applicCC=1 - заявка РА-КК
      column_Лист2_Выдача_PA_ПК = ColumnByValue(ReportName_String, "Лист2", "issuedLN", 100, 100) ' issuedLN=1 - выдача РА-ПК
      column_Лист2_Сегмент3 = ColumnByValue(ReportName_String, "Лист2", "typeSeg_MO", 100, 100) ' typeSeg_MO
      column_Лист2_cntProduct = ColumnByValue(ReportName_String, "Лист2", "cntProduct", 100, 100) ' cntProduct

      ' Определяем столбцы на Лист3
      column_Лист3_Офис = ColumnByValue(ReportName_String, "Лист3", "DP4_отчет", 100, 100) ' Q
      column_Лист3_Дата_обновления = ColumnByValue(ReportName_String, "Лист3", "load_date", 100, 100) ' Дата обновления
      column_Лист3_Месяц = ColumnByValue(ReportName_String, "Лист3", "month", 100, 100) '
      column_Лист3_НК = ColumnByValue(ReportName_String, "Лист3", "NK", 100, 100) ' НК ритейл
      column_Лист3_CRM_НК = ColumnByValue(ReportName_String, "Лист3", "ybpideal", 100, 100) ' Идентификатор клиента CRM
      column_Лист3_ФИО_МРК = ColumnByValue(ReportName_String, "Лист3", "Manager", 100, 100) ' МРК
      column_Лист3_ТН_МРК = ColumnByValue(ReportName_String, "Лист3", "Табельный_номер", 100, 100) ' табельный номер
      column_Лист3_BIC_RFOFICID = ColumnByValue(ReportName_String, "Лист3", "/BIC/RFOFICID", 100, 100) ' Офис операции
      column_Лист3_Есть_ПК_КК = ColumnByValue(ReportName_String, "Лист3", "type", 100, 100) ' type = 1/0 Наличие типа продукта РА (ПК / КК) в момент операции в офисе
      column_Лист3_Сегмент = ColumnByValue(ReportName_String, "Лист3", "typeSeg", 100, 100) ' тип сегмента
      column_Лист3_potok = ColumnByValue(ReportName_String, "Лист3", "potok", 100, 100) ' potok=1- клиент в офисе с операцией и имеет РА
      column_Лист3_Заявка_PA_KK = ColumnByValue(ReportName_String, "Лист3", "applicCC", 100, 100) ' applicCC=1 - заявка РА-КК
      column_Лист3_Выдача_PA_ПК = ColumnByValue(ReportName_String, "Лист3", "issuedLN", 100, 100) ' issuedLN=1 - выдача РА-ПК
      column_Лист3_Сегмент3 = ColumnByValue(ReportName_String, "Лист3", "typeSeg_MO", 100, 100) ' typeSeg_MO
      column_Лист3_cntProduct = ColumnByValue(ReportName_String, "Лист3", "cntProduct", 100, 100) ' cntProduct

      ' Счетчик вывода новых строк на лист "Вх_PA"
      countRowNewLine_Вх_PA = 5

      ' Счетчики
      countКлиенты_с_PA_ПК = 0
      count_Выдача_РА_ПК = 0
      
      countРОО_Клиенты_с_PA_ПК = 0
      countРОО_Выдача_РА_ПК = 0

      countКлиенты_с_PA_КК = 0
      count_Заказ_РА_КК = 0

      countРОО_Клиенты_с_PA_КК = 0
      countРОО_Заказ_РА_КК = 0


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

        ' Рисуем название офиса и синюю полоску
        Call writeOffice_Вх_PA(officeNameInReport, i)
        
        ' Обработка Лист2 *** Потребительские кредиты ***
        rowCount = 1
        Do While Not IsEmpty(Workbooks(ReportName_String).Sheets("Лист2").Cells(rowCount, 1).Value)
        
          ' Если это текущий офис
          If InStr(Workbooks(ReportName_String).Sheets("Лист2").Cells(rowCount, column_Лист2_Офис).Value, officeNameInReport) <> 0 Then
            
            ' Вносим МРК на лист отчета (если его нет) и  суммируем на нем данные
            Call writeМРК_Вх_PA(Workbooks(ReportName_String).Sheets("Лист2").Cells(rowCount, column_Лист2_ТН_МРК).Value, _
                                  Workbooks(ReportName_String).Sheets("Лист2").Cells(rowCount, column_Лист2_ФИО_МРК).Value, _
                                    Workbooks(ReportName_String).Sheets("Лист2").Cells(rowCount, column_Лист2_Есть_ПК_КК).Value, _
                                      Workbooks(ReportName_String).Sheets("Лист2").Cells(rowCount, column_Лист2_Выдача_PA_ПК).Value)
            
            ' НК клиента
            НК_RetailVar = Workbooks(ReportName_String).Sheets("Лист2").Cells(rowCount, column_Лист2_НК).Value
            
            ' Проверяем по НК наличие кредита у клиента из отчета по активам Лист3, если есть, то вносим в поле "Статус_отработки" (офис проставляет: вкладчик, есть кредит, сотрудник и т.п.)
            ' НК_RetailVar
            Call Проверка_действующего_кредита(Dir(FileName2), "Объем КП", НК_RetailVar) ' Результат записывается в глобальные переменные Объем_КП_с_изм_за_период_НК, Объем_КП_с_изм_за_период_Число_кредитов, Объем_КП_с_изм_за_период_Сумма_кредитов
            
            If Объем_КП_с_изм_за_период_Число_кредитов <> 0 Then
              Комментарий_Var = "Действующие кредиты " + CStr(Объем_КП_с_изм_за_период_Число_кредитов) + " шт., на сумму " + CStr(Round(Объем_КП_с_изм_за_период_Сумма_кредитов / 1000, 0)) + " тыс. руб. (" + Объем_КП_с_изм_за_период_Виды_кредитов + ")"
            Else
              Комментарий_Var = " "
            End If
            
            ' Вносим PA ПК
            Сегмент2Var = getDataFrom_Лист1(ReportName_String, "Лист1", column_Лист1_НК, Workbooks(ReportName_String).Sheets("Лист2").Cells(rowCount, column_Лист2_НК).Value, column_Лист1_Сегмент_детал)
            Call InsertRecordInBook(Dir(OutBookName), "Лист1", "НК_Retail", НК_RetailVar, _
                                              "Дата_загрузки", Workbooks(ReportName_String).Sheets("Лист2").Cells(rowCount, column_Лист2_Дата_обновления).Value, _
                                                "НК_Retail", НК_RetailVar, _
                                                  "ID_CRM", Workbooks(ReportName_String).Sheets("Лист2").Cells(rowCount, column_Лист2_CRM_НК).Value, _
                                                    "ФИО_клиента", ФИО_из_Объем_КП_с_изм_за_период, _
                                                      "Сегмент", Workbooks(ReportName_String).Sheets("Лист2").Cells(rowCount, column_Лист2_Сегмент).Value, _
                                                        "Сегмент2", Сегмент2Var, _
                                                          "Сегмент3", Workbooks(ReportName_String).Sheets("Лист2").Cells(rowCount, column_Лист2_Сегмент3).Value, _
                                                            "cntProduct", Workbooks(ReportName_String).Sheets("Лист2").Cells(rowCount, column_Лист2_cntProduct).Value, _
                                                              "PA_ПК", Workbooks(ReportName_String).Sheets("Лист2").Cells(rowCount, column_Лист2_Есть_ПК_КК).Value, _
                                                                "Выдача_PA_ПК", Workbooks(ReportName_String).Sheets("Лист2").Cells(rowCount, column_Лист2_Выдача_PA_ПК).Value, _
                                                                  "PA_KK", "", _
                                                                    "Заявка_РА_КК", "", _
                                                                      "ФИО_МРК", Фамилия_и_Имя(Workbooks(ReportName_String).Sheets("Лист2").Cells(rowCount, column_Лист2_ФИО_МРК).Value, 3), _
                                                                        "ДопОфис", officeNameInReport, _
                                                                          "Статус_обработки", " ", _
                                                                            "Комментарий", Комментарий_Var, _
                                                                              "", "", _
                                                                                "", "", _
                                                                                  "", "", _
                                                                                    "", "")
                                                                                    
                                                                                                
            ' Вносим в BASE\Clients поля Номер_клиента, Офис, PA_ПК, PA_KK, Сегмент, Сегмент2, Сегмент3, Дата_загрузки
            Call InsertRecordInBook("Clients", "Лист1", "Номер_клиента", НК_RetailVar, _
                                            "Номер_клиента", НК_RetailVar, _
                                              "Офис", cityOfficeName(officeNameInReport), _
                                                "PA_ПК", Workbooks(ReportName_String).Sheets("Лист2").Cells(rowCount, column_Лист2_Есть_ПК_КК).Value, _
                                                  "Сегмент", Workbooks(ReportName_String).Sheets("Лист2").Cells(rowCount, column_Лист2_Сегмент).Value, _
                                                    "Сегмент2", Сегмент2Var, _
                                                      "Сегмент3", Workbooks(ReportName_String).Sheets("Лист2").Cells(rowCount, column_Лист2_Сегмент3).Value, _
                                                        "Дата_загрузки", Workbooks(ReportName_String).Sheets("Лист2").Cells(rowCount, column_Лист2_Дата_обновления).Value, _
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
        
        
          ' Следующая запись
          rowCount = rowCount + 1
          Application.StatusBar = "Обработка " + officeNameInReport + ": " + CStr(rowCount) + "..."
          DoEventsInterval (rowCount)
        Loop
   
        ' Обработка Лист3 *** Кредитные карты ***
        rowCount = 1
        Do While Not IsEmpty(Workbooks(ReportName_String).Sheets("Лист3").Cells(rowCount, 1).Value)
        
          ' Если это текущий офис
          If InStr(Workbooks(ReportName_String).Sheets("Лист3").Cells(rowCount, column_Лист3_Офис).Value, officeNameInReport) <> 0 Then
            
            ' Вносим МРК на лист отчета (если его нет) и  суммируем на нем данные
            Call writeМРК_Вх_PA_KK(Workbooks(ReportName_String).Sheets("Лист3").Cells(rowCount, column_Лист3_ТН_МРК).Value, _
                                    Workbooks(ReportName_String).Sheets("Лист3").Cells(rowCount, column_Лист3_ФИО_МРК).Value, _
                                      Workbooks(ReportName_String).Sheets("Лист3").Cells(rowCount, column_Лист3_Есть_ПК_КК).Value, _
                                        Workbooks(ReportName_String).Sheets("Лист3").Cells(rowCount, column_Лист3_Заявка_PA_KK).Value)
            ' Вносим PA КК
            НК_RetailVar = Workbooks(ReportName_String).Sheets("Лист3").Cells(rowCount, column_Лист3_НК).Value
            Сегмент2Var = getDataFrom_Лист1(ReportName_String, "Лист1", column_Лист1_НК, Workbooks(ReportName_String).Sheets("Лист3").Cells(rowCount, column_Лист3_НК).Value, column_Лист1_Сегмент_детал)
            Call InsertRecordInBook(Dir(OutBookName2), "Лист1", "НК_Retail", НК_RetailVar, _
                                              "Дата_загрузки", Workbooks(ReportName_String).Sheets("Лист3").Cells(rowCount, column_Лист3_Дата_обновления).Value, _
                                                "НК_Retail", НК_RetailVar, _
                                                  "ID_CRM", Workbooks(ReportName_String).Sheets("Лист3").Cells(rowCount, column_Лист3_CRM_НК).Value, _
                                                    "ФИО_клиента", "", _
                                                      "Сегмент", Workbooks(ReportName_String).Sheets("Лист3").Cells(rowCount, column_Лист3_Сегмент).Value, _
                                                        "Сегмент2", Сегмент2Var, _
                                                          "Сегмент3", Workbooks(ReportName_String).Sheets("Лист3").Cells(rowCount, column_Лист3_Сегмент3).Value, _
                                                            "cntProduct", Workbooks(ReportName_String).Sheets("Лист3").Cells(rowCount, column_Лист3_cntProduct).Value, _
                                                              "PA_ПК", "", _
                                                                "Выдача_PA_ПК", "", _
                                                                  "PA_KK", Workbooks(ReportName_String).Sheets("Лист3").Cells(rowCount, column_Лист3_Есть_ПК_КК).Value, _
                                                                    "Заявка_РА_КК", Workbooks(ReportName_String).Sheets("Лист3").Cells(rowCount, column_Лист3_Заявка_PA_KK).Value, _
                                                                      "ФИО_МРК", Фамилия_и_Имя(Workbooks(ReportName_String).Sheets("Лист3").Cells(rowCount, column_Лист3_ФИО_МРК).Value, 3), _
                                                                        "ДопОфис", officeNameInReport, _
                                                                          "Статус_обработки", " ", _
                                                                            "Комментарий", " ", _
                                                                              "", "", _
                                                                                "", "", _
                                                                                  "", "", _
                                                                                    "", "")
            
            ' Вносим в BASE\Clients поля Номер_клиента, Офис, PA_ПК, PA_KK, Сегмент, Сегмент2, Сегмент3, Дата_загрузки
            Call InsertRecordInBook("Clients", "Лист1", "Номер_клиента", НК_RetailVar, _
                                            "Номер_клиента", НК_RetailVar, _
                                              "Офис", cityOfficeName(officeNameInReport), _
                                                "PA_KK", Workbooks(ReportName_String).Sheets("Лист3").Cells(rowCount, column_Лист2_Есть_ПК_КК).Value, _
                                                  "Сегмент", Workbooks(ReportName_String).Sheets("Лист3").Cells(rowCount, column_Лист2_Сегмент).Value, _
                                                    "Сегмент2", Сегмент2Var, _
                                                      "Сегмент3", Workbooks(ReportName_String).Sheets("Лист3").Cells(rowCount, column_Лист2_Сегмент3).Value, _
                                                        "Дата_загрузки", Workbooks(ReportName_String).Sheets("Лист3").Cells(rowCount, column_Лист2_Дата_обновления).Value, _
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
        
        
          ' Следующая запись
          rowCount = rowCount + 1
          Application.StatusBar = "Обработка " + officeNameInReport + ": " + CStr(rowCount) + "..."
          DoEventsInterval (rowCount)
        Loop
   
   
        ' Выводим данные по офису
      
      Next i ' Следующий офис
      
      ' Подводим итоги по прошлому офису i-1
      Call Подводим_итоги_по_прошлому_офису(i)
      
      ' Выводим итоги обработки
      Call Лист_Вх_PA_Итоги_РОО
      
      ' Закрываем выходную книгу с выгрузкой PA ПК
      Workbooks(Dir(OutBookName)).Close SaveChanges:=True
      
      ' Закрываем выходную книгу с выгрузкой PA КК
      Workbooks(Dir(OutBookName2)).Close SaveChanges:=True
      
      ' Закрываем базу BASE\Clients
      CloseBook ("Clients")
      
      ' Закрываем файл с отчетом без сохранения изменений (параметр SaveChanges:=False)
      Workbooks(Dir(FileName2)).Close SaveChanges:=False
      
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
    ThisWorkbook.Sheets("Вх_PA").Activate
    ThisWorkbook.Sheets("Вх_PA").Range("A1").Select

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

' Установка настроек сводной таблицы на листе "Вх. поток с РА-ПК (мес)"
Sub setFilter_Вх_поток_с_РА_ПК_мес(In_ReportName_String, In_Sheets, In_Срез_typeSeg)

  ' Строка статуса
  Application.StatusBar = "Открытие таблиц " + In_Sheets + "..."

  ' Переход на вкладку "Вх. поток с РА-ПК (мес)"/"CR РА-ПК (мес) Менеджер"
  Workbooks(In_ReportName_String).Sheets(In_Sheets).Activate
  
  ' Определяем столбец, в котором есть значение "вх.поток" на Листе
  ' Если это Лист "Вх. поток с РА-ПК (мес)", то ищем второй столбец по счету, содержащий "вх.поток": номер_вх_поток_на_Листе = 2
  If In_Sheets = "Вх. поток с РА-ПК (мес)" Then
    номер_вх_поток_на_Листе = 2
  End If
  ' Если это Лист "CR РА-ПК (мес) Менеджер", то ищем первый столбец по счету, содержащий "вх.поток": номер_вх_поток_на_Листе = 1
  If In_Sheets = "CR РА-ПК (мес) Менеджер" Then
    номер_вх_поток_на_Листе = 1
  End If
  
  ' Выполняем поиск столбца
  column_вх_поток = ColumnByValue2(In_ReportName_String, In_Sheets, "вх.поток", 1000, 1000, номер_вх_поток_на_Листе)
  
  ' Строка "Сумма по полю CR"
  ' row_Сумма_по_полю_CR = rowByValue(In_ReportName_String, In_Sheets, "Сумма по полю CR", 1000, 1000)
  ' Выполняем поиск столбца
  ' column_вх_поток = ColumnByNameAndNumber(In_ReportName_String, In_Sheets, row_Сумма_по_полю_CR, "вх.поток", номер_вх_поток_на_Листе, 100)

  ' Вкладка "Вх. поток с РА-ПК (мес)"
  ' Срез_typeSeg = "Срез_typeSeg1"
  
  ' Используем эту вкладку - тут есть МРК!
  ' Вкладка "CR РА-ПК (мес) Менеджер"
  ' Срез_typeSeg = "Срез_typeSeg4"

  ' Строка "Тюменский ОО1"
  row_Тюменский_ОО1 = rowByValue(In_ReportName_String, In_Sheets, "Тюменский ОО1", 1000, 1000)

  Workbooks(In_ReportName_String).SlicerCaches(In_Срез_typeSeg).SlicerItems("Вкладчик").Selected = True
  Workbooks(In_ReportName_String).SlicerCaches(In_Срез_typeSeg).SlicerItems("ЗП").Selected = True
  Workbooks(In_ReportName_String).SlicerCaches(In_Срез_typeSeg).SlicerItems("Дебетовщик").Selected = True
  Workbooks(In_ReportName_String).SlicerCaches(In_Срез_typeSeg).SlicerItems("Другой").Selected = True
  Workbooks(In_ReportName_String).SlicerCaches(In_Срез_typeSeg).SlicerItems("Заемщик").Selected = True
  Workbooks(In_ReportName_String).SlicerCaches(In_Срез_typeSeg).SlicerItems("(пусто)").Selected = False
  Workbooks(In_ReportName_String).SlicerCaches(In_Срез_typeSeg).SlicerItems("").Selected = False
  Workbooks(In_ReportName_String).ShowPivotTableFieldList = False
  ' Открываем новый ЛистX
  Workbooks(In_ReportName_String).Sheets(In_Sheets).Cells(row_Тюменский_ОО1, column_вх_поток - 1).Select
  Workbooks(In_ReportName_String).Sheets(In_Sheets).Cells(row_Тюменский_ОО1, column_вх_поток - 1).ShowDetail = True

  ' Открывается вкладка Лист1
  
  ' Определяем дату отчета
  column_ЛистX_Дата_обновления = ColumnByValue(In_ReportName_String, "Лист1", "load_date", 100, 100) ' Дата обновления
  dateReport = CDate(Workbooks(In_ReportName_String).Sheets("Лист1").Cells(2, column_ЛистX_Дата_обновления).Value)
 
' Вкладка "Вх. поток с РА-ПК (мес)"
'     With ActiveWorkbook.SlicerCaches("Срез_typeSeg1")
'        .SlicerItems("Вкладчик").Selected = True
'        .SlicerItems("ЗП").Selected = True
'        .SlicerItems("").Selected = False
'        .SlicerItems("Дебетовщик").Selected = False
'        .SlicerItems("Другой").Selected = False
'        .SlicerItems("Заемщик").Selected = False
'        .SlicerItems("(пусто)").Selected = False
'    End With
'    With ActiveWorkbook.SlicerCaches("Срез_typeSeg1")
'        .SlicerItems("Вкладчик").Selected = True
'        .SlicerItems("Дебетовщик").Selected = True
'        .SlicerItems("ЗП").Selected = True
'        .SlicerItems("").Selected = False
'        .SlicerItems("Другой").Selected = False
'        .SlicerItems("Заемщик").Selected = False
'        .SlicerItems("(пусто)").Selected = False
'    End With
'    With ActiveWorkbook.SlicerCaches("Срез_typeSeg1")
'        .SlicerItems("Вкладчик").Selected = True
'        .SlicerItems("Дебетовщик").Selected = True
'        .SlicerItems("Другой").Selected = True
'        .SlicerItems("ЗП").Selected = True
'        .SlicerItems("").Selected = False
'        .SlicerItems("Заемщик").Selected = False
'        .SlicerItems("(пусто)").Selected = False
'    End With
'    With ActiveWorkbook.SlicerCaches("Срез_typeSeg1")
'        .SlicerItems("Вкладчик").Selected = True
'        .SlicerItems("Дебетовщик").Selected = True
'        .SlicerItems("Другой").Selected = True
'        .SlicerItems("Заемщик").Selected = True
'        .SlicerItems("ЗП").Selected = True
'        .SlicerItems("").Selected = False
'        .SlicerItems("(пусто)").Selected = False
'    End With
'    ActiveWorkbook.ShowPivotTableFieldList = False
'    Range("T124").Select
'    Selection.ShowDetail = True
'

' Открывается Лист1

' Используем эту вкладку - тут есть МРК!
' Вкладка "CR РА-ПК (мес) Менеджер"
'    With ActiveWorkbook.SlicerCaches("Срез_typeSeg4")
'        .SlicerItems("Дебетовщик").Selected = True
'        .SlicerItems("Заемщик").Selected = False
'        .SlicerItems("ЗП").Selected = False
'        .SlicerItems("").Selected = False
'        .SlicerItems("Вкладчик").Selected = False
'        .SlicerItems("Другой").Selected = False
'        .SlicerItems("(пусто)").Selected = False
'    End With
'    With ActiveWorkbook.SlicerCaches("Срез_typeSeg4")
'        .SlicerItems("Дебетовщик").Selected = True
'        .SlicerItems("Заемщик").Selected = True
'        .SlicerItems("ЗП").Selected = False
'        .SlicerItems("").Selected = False
'        .SlicerItems("Вкладчик").Selected = False
'        .SlicerItems("Другой").Selected = False
'        .SlicerItems("(пусто)").Selected = False
'    End With
'    With ActiveWorkbook.SlicerCaches("Срез_typeSeg4")
'        .SlicerItems("Дебетовщик").Selected = True
'        .SlicerItems("Заемщик").Selected = True
'        .SlicerItems("ЗП").Selected = True
'        .SlicerItems("").Selected = False
'        .SlicerItems("Вкладчик").Selected = False
'        .SlicerItems("Другой").Selected = False
'        .SlicerItems("(пусто)").Selected = False
'    End With
'    With ActiveWorkbook.SlicerCaches("Срез_type4")
'        .SlicerItems("KK").Selected = True
'        .SlicerItems("PK").Selected = True
'        .SlicerItems("(пусто)").Selected = False
'    End With
'    With ActiveWorkbook.SlicerCaches("Срез_typeSeg4")
'        .SlicerItems("Вкладчик").Selected = True
'        .SlicerItems("Дебетовщик").Selected = True
'        .SlicerItems("Заемщик").Selected = True
'        .SlicerItems("ЗП").Selected = True
'        .SlicerItems("").Selected = False
'        .SlicerItems("Другой").Selected = False
'        .SlicerItems("(пусто)").Selected = False
'    End With
'    With ActiveWorkbook.SlicerCaches("Срез_typeSeg4")
'        .SlicerItems("Вкладчик").Selected = True
'        .SlicerItems("Дебетовщик").Selected = True
'        .SlicerItems("Другой").Selected = True
'        .SlicerItems("Заемщик").Selected = True
'        .SlicerItems("ЗП").Selected = True
'        .SlicerItems("").Selected = False
'        .SlicerItems("(пусто)").Selected = False
'    End With
'    ActiveWindow.SmallScroll Down:=33
'    Range("O76").Select
'    Selection.ShowDetail = True
' Открывается вкладка Лист1

  ' Строка статуса
  Application.StatusBar = ""

End Sub

' Создание книги с PA для обработки "Отчётность по входящему потоку с PA"
Sub createBook_out_PA2(In_OutBookName)

    ' Поля:
    
    Workbooks.Add
    ActiveWorkbook.SaveAs FileName:=In_OutBookName
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Activate
    
    ' Форматирование полей
    field_Number = 0
    
    ' Дата_загрузки
    field_Number = field_Number + 1
    field_Letter = ConvertToLetter(field_Number)
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, field_Number).Value = "Дата_загрузки"
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns(field_Letter + ":" + field_Letter).EntireColumn.ColumnWidth = 13
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, field_Number).HorizontalAlignment = xlLeft
    
    ' НК_Retail
    field_Number = field_Number + 1
    field_Letter = ConvertToLetter(field_Number)
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, field_Number).Value = "НК_Retail"
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns(field_Letter + ":" + field_Letter).EntireColumn.ColumnWidth = 11
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, field_Number).HorizontalAlignment = xlLeft
    
    ' ID_CRM
    field_Number = field_Number + 1
    field_Letter = ConvertToLetter(field_Number)
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, field_Number).Value = "ID_CRM"
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns(field_Letter + ":" + field_Letter).EntireColumn.ColumnWidth = 11
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, field_Number).HorizontalAlignment = xlLeft
    
    ' ФИО_клиента
    field_Number = field_Number + 1
    field_Letter = ConvertToLetter(field_Number)
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, field_Number).Value = "ФИО_клиента"
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns(field_Letter + ":" + field_Letter).EntireColumn.ColumnWidth = 15
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, field_Number).HorizontalAlignment = xlLeft
    
    ' Сегмент (c Листа2)
    field_Number = field_Number + 1
    field_Letter = ConvertToLetter(field_Number)
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, field_Number).Value = "Сегмент"
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns(field_Letter + ":" + field_Letter).EntireColumn.ColumnWidth = 10
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, field_Number).HorizontalAlignment = xlLeft
    
    ' Сегмент 2 (с Листа1)
    field_Number = field_Number + 1
    field_Letter = ConvertToLetter(field_Number)
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, field_Number).Value = "Сегмент2"
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns(field_Letter + ":" + field_Letter).EntireColumn.ColumnWidth = 10
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, field_Number).HorizontalAlignment = xlLeft
    
    ' Сегмент 3 (с Листа2 - гражданка/ОПК)
    field_Number = field_Number + 1
    field_Letter = ConvertToLetter(field_Number)
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, field_Number).Value = "Сегмент3"
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns(field_Letter + ":" + field_Letter).EntireColumn.ColumnWidth = 10
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, field_Number).HorizontalAlignment = xlLeft
    
    ' cntProduct - вероятно число продуктов
    field_Number = field_Number + 1
    field_Letter = ConvertToLetter(field_Number)
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, field_Number).Value = "cntProduct"
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns(field_Letter + ":" + field_Letter).EntireColumn.ColumnWidth = 9
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, field_Number).HorizontalAlignment = xlLeft
    
    ' PA_ПК
    field_Number = field_Number + 1
    field_Letter = ConvertToLetter(field_Number)
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, field_Number).Value = "PA_ПК"
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns(field_Letter + ":" + field_Letter).EntireColumn.ColumnWidth = 8.29
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, field_Number).HorizontalAlignment = xlLeft
 
    ' Выдача_PA_ПК
    field_Number = field_Number + 1
    field_Letter = ConvertToLetter(field_Number)
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, field_Number).Value = "Выдача_PA_ПК"
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns(field_Letter + ":" + field_Letter).EntireColumn.ColumnWidth = 15.86
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, field_Number).HorizontalAlignment = xlLeft
    
    ' PA_KK
    field_Number = field_Number + 1
    field_Letter = ConvertToLetter(field_Number)
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, field_Number).Value = "PA_KK"
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns(field_Letter + ":" + field_Letter).EntireColumn.ColumnWidth = 8.14
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, field_Number).HorizontalAlignment = xlLeft
    
    ' Заявка_РА_КК
    field_Number = field_Number + 1
    field_Letter = ConvertToLetter(field_Number)
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, field_Number).Value = "Заявка_РА_КК"
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns(field_Letter + ":" + field_Letter).EntireColumn.ColumnWidth = 15
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, field_Number).HorizontalAlignment = xlLeft
    
    ' ФИО_МРК
    field_Number = field_Number + 1
    field_Letter = ConvertToLetter(field_Number)
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, field_Number).Value = "ФИО_МРК"
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns(field_Letter + ":" + field_Letter).EntireColumn.ColumnWidth = 11
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, field_Number).HorizontalAlignment = xlLeft
    
    ' ДопОфис
    field_Number = field_Number + 1
    field_Letter = ConvertToLetter(field_Number)
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, field_Number).Value = "ДопОфис"
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns(field_Letter + ":" + field_Letter).EntireColumn.ColumnWidth = 11
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, field_Number).HorizontalAlignment = xlLeft
    
    ' Статус_отработки (офис проставляет: вкладчик, есть кредит, сотрудник и т.п.)
    field_Number = field_Number + 1
    field_Letter = ConvertToLetter(field_Number)
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, field_Number).Value = "Статус_обработки"
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns(field_Letter + ":" + field_Letter).EntireColumn.ColumnWidth = 18
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, field_Number).HorizontalAlignment = xlLeft

    ' Комментарий (вбивает комментарий)
    field_Number = field_Number + 1
    field_Letter = ConvertToLetter(field_Number)
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, field_Number).Value = "Комментарий"
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Columns(field_Letter + ":" + field_Letter).EntireColumn.ColumnWidth = 100
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Cells(1, field_Number).HorizontalAlignment = xlLeft

    ' Установка фильтров
    Workbooks(Dir(In_OutBookName)).Sheets("Лист1").Range("A1:" + field_Letter + "1").Select
    Selection.AutoFilter

End Sub


' Рисуем название офиса и синюю полоску
Sub writeOffice_Вх_PA(In_officeNameInReport, In_i)
  
  ' Подводим итоги по прошлому офису i-1
  Call Подводим_итоги_по_прошлому_офису(In_i)
  
  ' Номер строки на Листе "Вх PA"
  countRowNewLine_Вх_PA = countRowNewLine_Вх_PA + 1
  
  ' Офис i
  ThisWorkbook.Sheets("Вх_PA").Cells(countRowNewLine_Вх_PA, 1).Value = CStr(In_i)
  
  ' Офис наименование
  ThisWorkbook.Sheets("Вх_PA").Cells(countRowNewLine_Вх_PA, 2).Value = getNameOfficeByNumber(In_i)
  ThisWorkbook.Sheets("Вх_PA").Cells(countRowNewLine_Вх_PA, 2).HorizontalAlignment = xlLeft
  
  ' Цвет всей строки
  Call setColorCells(ThisWorkbook.Name, "Вх_PA", countRowNewLine_Вх_PA, 2, countRowNewLine_Вх_PA, 9)

  ' Счетчики сбрасываем
  countКлиенты_с_PA_ПК = 0
  count_Выдача_РА_ПК = 0
  
  countКлиенты_с_PA_КК = 0
  count_Заказ_РА_КК = 0

End Sub
        

' Вносим МРК на лист отчета (если его нет) и  суммируем на нем данные
Sub writeМРК_Вх_PA(In_МРК_ТабНом, In_МРК_ФИО, In_Клиент_с_PA_ПК, In_Выдан_PA_ПК)
    
  ' Выполняем поиск данного МРК на Лист "Вх_PA"
  row_МРК = rowByValue(ThisWorkbook.Name, "Вх_PA", "#" + In_МРК_ТабНом, 100, 100)
  
  ' Если МРК не найден
  If row_МРК = 0 Then
    ' Добавляем МРК
    countRowNewLine_Вх_PA = countRowNewLine_Вх_PA + 1
    row_МРК = countRowNewLine_Вх_PA
    ' Табномер
    ThisWorkbook.Sheets("Вх_PA").Cells(countRowNewLine_Вх_PA, 1).Value = "#" + In_МРК_ТабНом
    ThisWorkbook.Sheets("Вх_PA").Cells(countRowNewLine_Вх_PA, 1).NumberFormat = "@"
    ' Делаем текст в ячейке невидимым
    ThisWorkbook.Sheets("Вх_PA").Cells(countRowNewLine_Вх_PA, 1).Font.ThemeColor = xlThemeColorDark1
    ThisWorkbook.Sheets("Вх_PA").Cells(countRowNewLine_Вх_PA, 1).Font.TintAndShade = 0
    ' ФИО
    ThisWorkbook.Sheets("Вх_PA").Cells(countRowNewLine_Вх_PA, 2).Value = Фамилия_и_Имя(In_МРК_ФИО, 3)
    ThisWorkbook.Sheets("Вх_PA").Cells(countRowNewLine_Вх_PA, 2).NumberFormat = "@"
    ThisWorkbook.Sheets("Вх_PA").Cells(countRowNewLine_Вх_PA, 2).HorizontalAlignment = xlRight
    ' Клиенты
    ' ThisWorkbook.Sheets("Вх_PA").Cells(countRowNewLine_Вх_PA, 3).Value =
    ThisWorkbook.Sheets("Вх_PA").Cells(countRowNewLine_Вх_PA, 3).HorizontalAlignment = xlRight
    ' Клиенты с PA ПК
    ThisWorkbook.Sheets("Вх_PA").Cells(countRowNewLine_Вх_PA, 4).Value = 0
    ThisWorkbook.Sheets("Вх_PA").Cells(countRowNewLine_Вх_PA, 4).NumberFormat = "#,##0"
    ThisWorkbook.Sheets("Вх_PA").Cells(countRowNewLine_Вх_PA, 4).HorizontalAlignment = xlRight
    ' Выдача РА-ПК
    ThisWorkbook.Sheets("Вх_PA").Cells(countRowNewLine_Вх_PA, 5).Value = 0
    ThisWorkbook.Sheets("Вх_PA").Cells(countRowNewLine_Вх_PA, 5).NumberFormat = "#,##0"
    ThisWorkbook.Sheets("Вх_PA").Cells(countRowNewLine_Вх_PA, 5).HorizontalAlignment = xlRight
    ' Конверсия
    ThisWorkbook.Sheets("Вх_PA").Cells(countRowNewLine_Вх_PA, 6).Value = 0
    ' ThisWorkbook.Sheets("Вх_PA").Cells(countRowNewLine_Вх_PA, 6).NumberFormat = "0.0%"
    ThisWorkbook.Sheets("Вх_PA").Cells(countRowNewLine_Вх_PA, 6).NumberFormat = "0%"
    ThisWorkbook.Sheets("Вх_PA").Cells(countRowNewLine_Вх_PA, 6).HorizontalAlignment = xlRight
    
    ' *** КК ***
    ' Клиенты с PA КК
    ThisWorkbook.Sheets("Вх_PA").Cells(countRowNewLine_Вх_PA, 7).Value = 0
    ThisWorkbook.Sheets("Вх_PA").Cells(countRowNewLine_Вх_PA, 7).NumberFormat = "#,##0"
    ThisWorkbook.Sheets("Вх_PA").Cells(countRowNewLine_Вх_PA, 7).HorizontalAlignment = xlRight
    ' Заказ РА КК
    ThisWorkbook.Sheets("Вх_PA").Cells(countRowNewLine_Вх_PA, 8).Value = 0
    ThisWorkbook.Sheets("Вх_PA").Cells(countRowNewLine_Вх_PA, 8).NumberFormat = "#,##0"
    ThisWorkbook.Sheets("Вх_PA").Cells(countRowNewLine_Вх_PA, 8).HorizontalAlignment = xlRight
    ' Конверсия
    ThisWorkbook.Sheets("Вх_PA").Cells(countRowNewLine_Вх_PA, 9).Value = 0
    ' ThisWorkbook.Sheets("Вх_PA").Cells(countRowNewLine_Вх_PA, 9).NumberFormat = "0.0%"
    ThisWorkbook.Sheets("Вх_PA").Cells(countRowNewLine_Вх_PA, 9).NumberFormat = "0%"
    ThisWorkbook.Sheets("Вх_PA").Cells(countRowNewLine_Вх_PA, 9).HorizontalAlignment = xlRight

  End If
    
  ' Апгрейдим на нем цифры
  ' Клиенты
  ' ThisWorkbook.Sheets("Вх_PA").Cells(countRowNewLine_Вх_PA, 3).Value =
  
  ' Клиенты с PA ПК
  If In_Клиент_с_PA_ПК = "PK" Then
    ThisWorkbook.Sheets("Вх_PA").Cells(countRowNewLine_Вх_PA, 4).Value = ThisWorkbook.Sheets("Вх_PA").Cells(countRowNewLine_Вх_PA, 4).Value + 1
    countКлиенты_с_PA_ПК = countКлиенты_с_PA_ПК + 1
    countРОО_Клиенты_с_PA_ПК = countРОО_Клиенты_с_PA_ПК + 1
  End If
  
  ' Выдача РА-ПК
  If In_Выдан_PA_ПК = 1 Then
    ThisWorkbook.Sheets("Вх_PA").Cells(countRowNewLine_Вх_PA, 5).Value = ThisWorkbook.Sheets("Вх_PA").Cells(countRowNewLine_Вх_PA, 5).Value + 1
    count_Выдача_РА_ПК = count_Выдача_РА_ПК + 1
    countРОО_Выдача_РА_ПК = countРОО_Выдача_РА_ПК + 1
  End If
    
  ' Конверсия
  ThisWorkbook.Sheets("Вх_PA").Cells(countRowNewLine_Вх_PA, 6).Value = РассчетДоли(ThisWorkbook.Sheets("Вх_PA").Cells(countRowNewLine_Вх_PA, 4).Value, ThisWorkbook.Sheets("Вх_PA").Cells(countRowNewLine_Вх_PA, 5).Value, 3)
    
End Sub

' Вносим МРК на лист отчета (если его нет) и  суммируем на нем данные
Sub writeМРК_Вх_PA_KK(In_МРК_ТабНом, In_МРК_ФИО, In_Клиент_с_PA_КК, In_Заказан_PA_КК)
    
  ' Выполняем поиск данного МРК на Лист "Вх_PA"
  row_МРК = rowByValue(ThisWorkbook.Name, "Вх_PA", "#" + In_МРК_ТабНом, 100, 100)
  
  If row_МРК <> 0 Then
  
    ' Клиенты с PA КК
    If In_Клиент_с_PA_КК = "KK" Then
      ThisWorkbook.Sheets("Вх_PA").Cells(row_МРК, 7).Value = ThisWorkbook.Sheets("Вх_PA").Cells(row_МРК, 7).Value + 1
      countКлиенты_с_PA_КК = countКлиенты_с_PA_КК + 1
      countРОО_Клиенты_с_PA_КК = countРОО_Клиенты_с_PA_КК + 1
    End If
  
    ' Заказ РА КК
    If In_Заказан_PA_КК = 1 Then
      ThisWorkbook.Sheets("Вх_PA").Cells(row_МРК, 8).Value = ThisWorkbook.Sheets("Вх_PA").Cells(row_МРК, 8).Value + 1
      count_Заказ_РА_КК = count_Заказ_РА_КК + 1
      countРОО_Заказ_РА_КК = countРОО_Заказ_РА_КК + 1
    End If
    
    ' Конверсия
    ThisWorkbook.Sheets("Вх_PA").Cells(row_МРК, 9).Value = РассчетДоли(ThisWorkbook.Sheets("Вх_PA").Cells(row_МРК, 7).Value, ThisWorkbook.Sheets("Вх_PA").Cells(countRowNewLine_Вх_PA, 8).Value, 3)
    
  End If
    
End Sub

' Очистка поля отчета на Листе
Sub clearСontents_Вх_PA()

  rowBegin = rowByValue(ThisWorkbook.Name, "Вх_PA", "ОО «Тюменский»", 100, 100)
      
  If rowBegin = 0 Then
    rowBegin = 6
  End If
      
  rowEnd = rowByValue(ThisWorkbook.Name, "Вх_PA", "Итого по РОО", 100, 100)
        
  If rowEnd = 0 Then
    rowEnd = 30
  End If
      
  Call clearСontents2(ThisWorkbook.Name, "Вх_PA", "A" + CStr(rowBegin), "I" + CStr(rowEnd))
      
End Sub

' Подводим итоги по прошлому офису i-1
Sub Подводим_итоги_по_прошлому_офису(In_i)

  If (In_i - 1) > 0 Then
    
    rowPreviousOffice = rowByValue(ThisWorkbook.Name, "Вх_PA", getNameOfficeByNumber(In_i - 1), 100, 100)
    ' Клиенты
    ' ThisWorkbook.Sheets("Вх_PA").Cells(rowPreviousOffice, 3).Value =
    ' ThisWorkbook.Sheets("Вх_PA").Cells(rowPreviousOffice, 3).HorizontalAlignment = xlRight
    ' Клиенты с PA ПК
    ThisWorkbook.Sheets("Вх_PA").Cells(rowPreviousOffice, 4).Value = countКлиенты_с_PA_ПК
    ThisWorkbook.Sheets("Вх_PA").Cells(rowPreviousOffice, 4).NumberFormat = "#,##0"
    ThisWorkbook.Sheets("Вх_PA").Cells(rowPreviousOffice, 4).HorizontalAlignment = xlRight
    ' Выдача РА-ПК
    ThisWorkbook.Sheets("Вх_PA").Cells(rowPreviousOffice, 5).Value = count_Выдача_РА_ПК
    ThisWorkbook.Sheets("Вх_PA").Cells(rowPreviousOffice, 5).NumberFormat = "#,##0"
    ThisWorkbook.Sheets("Вх_PA").Cells(rowPreviousOffice, 5).HorizontalAlignment = xlRight
    ' Конверсия
    ThisWorkbook.Sheets("Вх_PA").Cells(rowPreviousOffice, 6).Value = РассчетДоли(ThisWorkbook.Sheets("Вх_PA").Cells(rowPreviousOffice, 4).Value, ThisWorkbook.Sheets("Вх_PA").Cells(rowPreviousOffice, 5).Value, 3)
    ' ThisWorkbook.Sheets("Вх_PA").Cells(rowPreviousOffice, 6).NumberFormat = "0.0%"
    ThisWorkbook.Sheets("Вх_PA").Cells(rowPreviousOffice, 6).NumberFormat = "0%"
    ThisWorkbook.Sheets("Вх_PA").Cells(rowPreviousOffice, 6).HorizontalAlignment = xlRight
    
    ' КК
    ' Клиенты с КК
    ThisWorkbook.Sheets("Вх_PA").Cells(rowPreviousOffice, 7).Value = countКлиенты_с_PA_КК
    ThisWorkbook.Sheets("Вх_PA").Cells(rowPreviousOffice, 7).NumberFormat = "#,##0"
    ThisWorkbook.Sheets("Вх_PA").Cells(rowPreviousOffice, 7).HorizontalAlignment = xlRight
    ' Заказ РА-КК
    ThisWorkbook.Sheets("Вх_PA").Cells(rowPreviousOffice, 8).Value = count_Заказ_РА_КК
    ThisWorkbook.Sheets("Вх_PA").Cells(rowPreviousOffice, 8).NumberFormat = "#,##0"
    ThisWorkbook.Sheets("Вх_PA").Cells(rowPreviousOffice, 8).HorizontalAlignment = xlRight
    ' Конверсия
    ThisWorkbook.Sheets("Вх_PA").Cells(rowPreviousOffice, 9).Value = РассчетДоли(ThisWorkbook.Sheets("Вх_PA").Cells(rowPreviousOffice, 7).Value, ThisWorkbook.Sheets("Вх_PA").Cells(rowPreviousOffice, 8).Value, 3)
    ' ThisWorkbook.Sheets("Вх_PA").Cells(rowPreviousOffice, 9).NumberFormat = "0.0%"
    ThisWorkbook.Sheets("Вх_PA").Cells(rowPreviousOffice, 9).NumberFormat = "0%"
    ThisWorkbook.Sheets("Вх_PA").Cells(rowPreviousOffice, 9).HorizontalAlignment = xlRight


  End If

End Sub

' Подведение итогов РОО
Sub Лист_Вх_PA_Итоги_РОО()
    
  countRowNewLine_Вх_PA = countRowNewLine_Вх_PA + 1
    
  ' Чертим горизонтальную линию 2 (указываем предидущее значение строки)
  Call gorizontalLineII(ThisWorkbook.Name, "Вх_PA", countRowNewLine_Вх_PA, 2, 9)
    
  ' Итого по РОО
  ThisWorkbook.Sheets("Вх_PA").Cells(countRowNewLine_Вх_PA, 2).Value = "Итого по РОО"
  ThisWorkbook.Sheets("Вх_PA").Cells(countRowNewLine_Вх_PA, 2).HorizontalAlignment = xlLeft
  ' Клиенты
  ' ThisWorkbook.Sheets("Вх_PA").Cells(countRowNewLine_Вх_PA, 3).Value =
  ' ThisWorkbook.Sheets("Вх_PA").Cells(countRowNewLine_Вх_PA, 3).HorizontalAlignment = xlRight
  ' Клиенты с PA ПК
  ThisWorkbook.Sheets("Вх_PA").Cells(countRowNewLine_Вх_PA, 4).Value = countРОО_Клиенты_с_PA_ПК
  ThisWorkbook.Sheets("Вх_PA").Cells(countRowNewLine_Вх_PA, 4).NumberFormat = "#,##0"
  ThisWorkbook.Sheets("Вх_PA").Cells(countRowNewLine_Вх_PA, 4).HorizontalAlignment = xlRight
  ' Выдача РА-ПК
  ThisWorkbook.Sheets("Вх_PA").Cells(countRowNewLine_Вх_PA, 5).Value = countРОО_Выдача_РА_ПК
  ThisWorkbook.Sheets("Вх_PA").Cells(countRowNewLine_Вх_PA, 5).NumberFormat = "#,##0"
  ThisWorkbook.Sheets("Вх_PA").Cells(countRowNewLine_Вх_PA, 5).HorizontalAlignment = xlRight
  ' Конверсия
  ThisWorkbook.Sheets("Вх_PA").Cells(countRowNewLine_Вх_PA, 6).Value = РассчетДоли(ThisWorkbook.Sheets("Вх_PA").Cells(countRowNewLine_Вх_PA, 4).Value, ThisWorkbook.Sheets("Вх_PA").Cells(countRowNewLine_Вх_PA, 5).Value, 3)
  ' ThisWorkbook.Sheets("Вх_PA").Cells(countRowNewLine_Вх_PA, 6).NumberFormat = "0.0%"
  ThisWorkbook.Sheets("Вх_PA").Cells(countRowNewLine_Вх_PA, 6).NumberFormat = "0%"
  ThisWorkbook.Sheets("Вх_PA").Cells(countRowNewLine_Вх_PA, 6).HorizontalAlignment = xlRight
  
  ' *** КК ***
  ' Клиенты с PA ПК
  ThisWorkbook.Sheets("Вх_PA").Cells(countRowNewLine_Вх_PA, 7).Value = countРОО_Клиенты_с_PA_КК
  ThisWorkbook.Sheets("Вх_PA").Cells(countRowNewLine_Вх_PA, 7).NumberFormat = "#,##0"
  ThisWorkbook.Sheets("Вх_PA").Cells(countRowNewLine_Вх_PA, 7).HorizontalAlignment = xlRight
  ' Выдача РА-ПК
  ThisWorkbook.Sheets("Вх_PA").Cells(countRowNewLine_Вх_PA, 8).Value = countРОО_Заказ_РА_КК
  ThisWorkbook.Sheets("Вх_PA").Cells(countRowNewLine_Вх_PA, 8).NumberFormat = "#,##0"
  ThisWorkbook.Sheets("Вх_PA").Cells(countRowNewLine_Вх_PA, 8).HorizontalAlignment = xlRight
  ' Конверсия
  ThisWorkbook.Sheets("Вх_PA").Cells(countRowNewLine_Вх_PA, 9).Value = РассчетДоли(ThisWorkbook.Sheets("Вх_PA").Cells(countRowNewLine_Вх_PA, 7).Value, ThisWorkbook.Sheets("Вх_PA").Cells(countRowNewLine_Вх_PA, 8).Value, 3)
  ' ThisWorkbook.Sheets("Вх_PA").Cells(countRowNewLine_Вх_PA, 9).NumberFormat = "0.0%"
  ThisWorkbook.Sheets("Вх_PA").Cells(countRowNewLine_Вх_PA, 9).NumberFormat = "0%"
  ThisWorkbook.Sheets("Вх_PA").Cells(countRowNewLine_Вх_PA, 9).HorizontalAlignment = xlRight
  
  
End Sub


' Получить typeSegDetal с Листа1 по НК клиента
Function getDataFrom_Лист1(In_ReportName_String, In_Sheets, In_column_NK, In_НК_Retail, In_columnNumber) As String
  
    getDataFrom_Лист1 = ""
  
    Литера_столбца = ConvertToLetter(In_column_NK)
  
    ' Выполняем поиск
    Set searchResults = Workbooks(In_ReportName_String).Sheets(In_Sheets).Columns(Литера_столбца + ":" + Литера_столбца).Find(In_НК_Retail, LookAt:=xlWhole)
  
    ' Проверяем - есть ли такая дата, если нет, то добавляем
    If searchResults Is Nothing Then
      ' Если не найдена
      
    Else
      ' Если найдена
      getDataFrom_Лист1 = Workbooks(In_ReportName_String).Sheets(In_Sheets).Cells(searchResults.Row, In_columnNumber).Value
      
    End If

  
End Function


' Установка фильтров для РА_КК
Sub setFilter_Вх_поток_с_РА_КК_мес(In_ReportName_String, In_Sheets, In_Срез_typeSeg)

  ' Строка статуса
  Application.StatusBar = "Открытие таблиц " + In_Sheets + "..."

  ' Переход на вкладку "CR РА-KК (мес) Менеджер"
  Workbooks(In_ReportName_String).Sheets(In_Sheets).Activate
  
  ' Выполняем поиск столбца
  column_вх_поток = ColumnByValue2(In_ReportName_String, In_Sheets, "вх.поток", 1000, 1000, 1)
  
  ' Строка "Тюменский ОО1"
  row_Тюменский_ОО1 = rowByValue(In_ReportName_String, In_Sheets, "Тюменский ОО1", 1000, 1000)

  Workbooks(In_ReportName_String).SlicerCaches(In_Срез_typeSeg).SlicerItems("Вкладчик").Selected = True
  Workbooks(In_ReportName_String).SlicerCaches(In_Срез_typeSeg).SlicerItems("ЗП").Selected = True
  Workbooks(In_ReportName_String).SlicerCaches(In_Срез_typeSeg).SlicerItems("Дебетовщик").Selected = True
  Workbooks(In_ReportName_String).SlicerCaches(In_Срез_typeSeg).SlicerItems("Другой").Selected = True
  Workbooks(In_ReportName_String).SlicerCaches(In_Срез_typeSeg).SlicerItems("Заемщик").Selected = True
  Workbooks(In_ReportName_String).SlicerCaches(In_Срез_typeSeg).SlicerItems("(пусто)").Selected = False
  Workbooks(In_ReportName_String).SlicerCaches(In_Срез_typeSeg).SlicerItems("").Selected = False
  Workbooks(In_ReportName_String).ShowPivotTableFieldList = False
  ' Открываем новый ЛистX
  Workbooks(In_ReportName_String).Sheets(In_Sheets).Cells(row_Тюменский_ОО1, column_вх_поток - 1).ShowDetail = True

  ' Открывается вкладка Лист1
    
  '  With ActiveWorkbook.SlicerCaches("Срез_typeSeg41")
  '      .SlicerItems("Вкладчик").Selected = True
  '      .SlicerItems("Заемщик").Selected = True
  '      .SlicerItems("ЗП").Selected = True
  '      .SlicerItems("").Selected = False
  '      .SlicerItems("Дебетовщик").Selected = False
  '      .SlicerItems("Другой").Selected = False
  '      .SlicerItems("(пусто)").Selected = False
  '  End With
  '  With ActiveWorkbook.SlicerCaches("Срез_typeSeg41")
  '      .SlicerItems("Вкладчик").Selected = True
  '      .SlicerItems("Дебетовщик").Selected = True
  '      .SlicerItems("Заемщик").Selected = True
  '      .SlicerItems("ЗП").Selected = True
  '      .SlicerItems("").Selected = False
  '      .SlicerItems("Другой").Selected = False
  '      .SlicerItems("(пусто)").Selected = False
  '  End With
  '  With ActiveWorkbook.SlicerCaches("Срез_typeSeg41")
  '      .SlicerItems("Вкладчик").Selected = True
  '      .SlicerItems("Дебетовщик").Selected = True
  '      .SlicerItems("Другой").Selected = True
  '      .SlicerItems("Заемщик").Selected = True
  '      .SlicerItems("ЗП").Selected = True
  '      .SlicerItems("").Selected = False
  '      .SlicerItems("(пусто)").Selected = False
  '  End With
  '  Range("O76").Select
  '  Selection.ShowDetail = True

End Sub

' Отправка_Lotus_Notes_Лист6_Pre_Approved Отправка письма: отправляю шаблон самому себе для последующей отправки в сеть письма на его основе:
Sub Отправка_Lotus_Notes_ЛистВХ_PA_Pre_Approved()
Dim темаПисьма, текстПисьма, hashTag, attachmentFile As String
Dim i As Byte
  
  ' Подтвержение
  If MsgBox("Отправить себе Шаблон письма с вложением Pre-Approved?", vbYesNo) = vbYes Then
    
    ' Формируем список для отправки (в "Список получателей:"):
    ThisWorkbook.Sheets("Вх_PA").Cells(rowByValue(ThisWorkbook.Name, "Вх_PA", "Список получателей:", 100, 100), ColumnByValue(ThisWorkbook.Name, "Вх_PA", "Список получателей:", 100, 100) + 2).Value _
                     = getFromAddrBook("УДО2,УДО3,УДО4,УДО5,НОРПиКО1,НОРПиКО2,НОРПиКО3,НОРПиКО4,НОРПиКО5,НОКП,РРКК", 2)
   
    ' Тема письма - Тема:
    темаПисьма = "Клиенты с Pre-Approved на " + Mid(ThisWorkbook.Sheets("Вх_PA").Range("O2").Value, 37, 10)

    ' hashTag - Хэштэг:
    hashTag = hashTagFromSheet("Вх_PA") + " #Pre-Approved"
    
    ' Файл-вложение из "Вложение2"
    attachmentFile = ThisWorkbook.Sheets("Вх_PA").Range("AO3").Value
 
    ' Текст письма
    текстПисьма = "" + Chr(13)
    текстПисьма = текстПисьма + "" + ThisWorkbook.Sheets("Вх_PA").Cells(rowByValue(ThisWorkbook.Name, "Вх_PA", "Список получателей:", 100, 100), ColumnByValue(ThisWorkbook.Name, "Вх_PA", "Список получателей:", 100, 100) + 2).Value + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "" + getFromAddrBook("РД", 2) + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "Уважаемые руководители," + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "Направляю список клиентов с упущенными готовыми решениями по потребкредитам и КК." + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "Прошу организовать отработку в срок до " + CStr(weekEndDate(Date) - 2) + " с конверсией не менее 20% " + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    ' Визитка (подпись С Ув., )
    текстПисьма = текстПисьма + ПодписьВПисьме()
    ' Хэштег
    текстПисьма = текстПисьма + createBlankStr(20) + hashTag
    
    ' Вызов
    Call send_Lotus_Notes(темаПисьма, "Sergey Fedorovich Proschaev/Tyumen/PSBank/Ru", "Sergey Fedorovich Proschaev/Tyumen/PSBank/Ru", текстПисьма, attachmentFile)

    ' Сообщение
    MsgBox ("Письмо отправлено!")
          
  End If
  
End Sub

            
' Проверяем по НК наличие кредита у клиента из отчета по активам Лист3, если есть, то вносим в поле "Статус_отработки" (офис проставляет: вкладчик, есть кредит, сотрудник и т.п.)
' НК_RetailVar
Sub Проверка_действующего_кредита(In_ReportName_String, In_Sheets, In_НК_Retail)
Dim НК_Retail_Int As Long

  ' Столбец "T" - "Номер клиента"
  ' Столбец "Q" - "Исходящий остаток"
  
  ' Сбразываем в глобальные переменные
  Объем_КП_с_изм_за_период_НК = 0
  Объем_КП_с_изм_за_период_Число_кредитов = 0
  Объем_КП_с_изм_за_период_Сумма_кредитов = 0
  Объем_КП_с_изм_за_период_Виды_кредитов = ""
  ФИО_из_Объем_КП_с_изм_за_период = ""
  
  ' Число_кредитов
  Число_кредитов = 0
  ' Сумма_кредитов
  Сумма_кредитов = 0

  ' Убираем нули
  НК_Retail_Int = CLng(In_НК_Retail)

  ' Обработка
  rowCount = 8 ' Данные с 8-ой строки
  Do While Not IsEmpty(Workbooks(In_ReportName_String).Sheets(In_Sheets).Cells(rowCount, 20).Value)
  
    ' Если это текущий клиент
    If Workbooks(In_ReportName_String).Sheets(In_Sheets).Cells(rowCount, 20).Value = CStr(НК_Retail_Int) Then
      
      ' Если Исх. остаток (17) > 0
      If Workbooks(In_ReportName_String).Sheets(In_Sheets).Cells(rowCount, 17).Value > 0 Then
        Число_кредитов = Число_кредитов + 1
        Сумма_кредитов = Сумма_кредитов + Workbooks(In_ReportName_String).Sheets(In_Sheets).Cells(rowCount, 17).Value
        Объем_КП_с_изм_за_период_Виды_кредитов = Объем_КП_с_изм_за_период_Виды_кредитов + Workbooks(In_ReportName_String).Sheets(In_Sheets).Cells(rowCount, 22).Value + " "
        
        ' Берем ФИО клиента из 21-го столбца "Ф.И.О."
        ФИО_из_Объем_КП_с_изм_за_период = Workbooks(In_ReportName_String).Sheets(In_Sheets).Cells(rowCount, 21).Value
        
      End If
      
    End If
  
    ' Следующая запись
    rowCount = rowCount + 1
    Application.StatusBar = "Проверка наличия кредитов: " + CStr(rowCount) + "..."
    DoEventsInterval (rowCount)
  Loop
  
  ' Записываем в глобальные переменные
  If Число_кредитов <> 0 Then
    Объем_КП_с_изм_за_период_НК = In_НК_Retail
    Объем_КП_с_изм_за_период_Число_кредитов = Число_кредитов
    Объем_КП_с_изм_за_период_Сумма_кредитов = Сумма_кредитов
  End If
  
End Sub