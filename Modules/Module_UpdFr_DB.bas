Attribute VB_Name = "Module_UpdFr_DB"
' *** Лист UpdFr_DB ***

' *** Глобальные переменные ***
Public In_Row_UpdFr_DB As Integer

' ***                       ***

' Добавить данные из DB
Sub UpdateFrom_DB()
Attribute UpdateFrom_DB.VB_ProcData.VB_Invoke_Func = " \n14"

  ' Открыть файл с отчетом
  FileName = Application.GetOpenFilename("Excel Files (*.xlsm), *.xlsm", , "Открытие файла с отчетом")
    
  ' Выводим для инфо данные об имени файла
  ReportName_String = Dir(FileName)
  
  ' Открываем выбранную книгу (UpdateLinks:=0)
  Workbooks.Open FileName, 0
      
  ' Переходим на окно DB
  ThisWorkbook.Sheets("UpdFr_DB").Activate

  ' Статус
  ThisWorkbook.Sheets("UpdFr_DB").Range("C6").Value = ""

  ' Очистить таблицу на Листе "UpdFr_DB"
  Call clearСontents2(ThisWorkbook.Name, "UpdFr_DB", "A9", "L14")

  ' Дата DB
  dateDB_UpdFr_DB = CDate(Mid(Workbooks(ReportName_String).Sheets("Оглавление").Cells(1, 1).Value, 23, 10))
  ThisWorkbook.Sheets("UpdFr_DB").Range("C7").Value = CStr(dateDB_UpdFr_DB)
  ' ThisWorkbook.Sheets("UpdFr_DB").Range("D8").Value = "Факт на " + strDDMMYY(dateDB_UpdFr_DB)

  ' Определяем столбец #Значение_переменной
  column_Значение_переменной = ColumnByValue(ThisWorkbook.Name, "UpdFr_DB", "#Значение_переменной", 100, 100)

  ' Инициализация переменных
  In_ReportName_String_Var = ReportName_String ' ThisWorkbook.Sheets("UpdFr_DB").Cells(8, 9).Value
  
  SheetName_String_Var = ThisWorkbook.Sheets("UpdFr_DB").Cells(rowByValue(ThisWorkbook.Name, "UpdFr_DB", "SheetName_String:", 100, 100), column_Значение_переменной).Value
  
  In_Product_Name_Var = ThisWorkbook.Sheets("UpdFr_DB").Cells(rowByValue(ThisWorkbook.Name, "UpdFr_DB", "In_Product_Name:", 100, 100), column_Значение_переменной).Value ' "Выдачи ПК"
  
  In_Product_Code_Var = Product_Name_to_Product_Code(In_Product_Name_Var) ' "Выдачи_ПК_шт"
  
  In_Unit_Var = Product_Name_to_Unit(In_Product_Name_Var)

  In_ColumnNameMonth = ThisWorkbook.Sheets("UpdFr_DB").Cells(rowByValue(ThisWorkbook.Name, "UpdFr_DB", "In_ColumnNameMonth:", 100, 100), column_Значение_переменной).Value
  
  In_ColumnNameQuarter = ThisWorkbook.Sheets("UpdFr_DB").Cells(rowByValue(ThisWorkbook.Name, "UpdFr_DB", "In_ColumnNameQuarter:", 100, 100), column_Значение_переменной).Value

  In_DeltaPrediction = ThisWorkbook.Sheets("UpdFr_DB").Cells(rowByValue(ThisWorkbook.Name, "UpdFr_DB", "In_DeltaPrediction:", 100, 100), column_Значение_переменной).Value

  In_Заголовок_столбца_офисы = ThisWorkbook.Sheets("UpdFr_DB").Cells(rowByValue(ThisWorkbook.Name, "UpdFr_DB", "In_Заголовок_столбца_офисы:", 100, 100), column_Значение_переменной).Value

  ' Проверка наличия Листа в DB
  StringInSheet = SheetName_String_Var
  SheetName_String = FindNameSheet(ReportName_String, StringInSheet) ' "3.11 Зарплатные карты"
  If SheetName_String <> "" Then

    ' Переходим в DB на нужный Лист
    Workbooks(ReportName_String).Sheets(SheetName_String_Var).Activate

    ' Переходим на окно DB
    ThisWorkbook.Sheets("UpdFr_DB").Activate

    ' Заголовки
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
            
      ' Номер строки для вывода на Листе UpdFr_DB
      In_Row_UpdFr_DB = i + 8
        
      ' Находим номер строки с наименованием офиса
      officeNameInReport_Var = officeNameInReport ' ThisWorkbook.Sheets("UpdFr_DB").Cells(8, 9).Value
        
      ' Поля: ID_Rec, Оffice_Number, Product_Name, Оffice, MMYY, Update_Date, Product_Code, Plan, Unit, Fact, Percent_Completion

      ' ПК Выдачи, шт.
      Call DB_UniversalSheetInDB_UpdFr_DB(In_ReportName_String_Var, _
                                           SheetName_String_Var, _
                                             officeNameInReport_Var, _
                                               0, _
                                                 0, _
                                                   In_Product_Name_Var, _
                                                     In_Product_Code_Var, _
                                                       In_Unit_Var, _
                                                         0, _
                                                           In_ColumnNameMonth, _
                                                             In_ColumnNameQuarter, _
                                                               In_DeltaPrediction, _
                                                                 In_Заголовок_столбца_офисы, _
                                                                   0, _
                                                                     0, _
                                                                       0, _
                                                                         0, 1, 1)

    Next i

    ' Статус
    ThisWorkbook.Sheets("UpdFr_DB").Range("C6").Value = "Статус: Данные за " + CStr(ThisWorkbook.Sheets("UpdFr_DB").Range("C7").Value) + " извлечены, проверьте итоги!"

  Else
    
    ' Сообщение
    MsgBox ("В DB не найден Лист " + SheetName_String_Var + "!")
    
  End If

    ' Закрываем файл с отчетом без сохранения изменений (параметр SaveChanges:=False)
    If MsgBox("Закрыть DB?", vbYesNo) = vbYes Then

      Workbooks(Dir(FileName)).Close SaveChanges:=False ' тестирование
      ' Переходим в ячейку M2
      ThisWorkbook.Sheets("UpdFr_DB").Range("A1").Select
    
    Else
    
      
    
    End If
    

    ' Сообщение
    MsgBox ("Обработка завершена!")
 
    Application.StatusBar = ""

End Sub



' Показатель из вкладки DB
Sub DB_UniversalSheetInDB_UpdFr_DB(In_ReportName_String, In_Sheets, In_officeNameInReport, In_Row_Лист8, In_N, In_Product_Name, In_Product_Code, In_Unit, In_Weight, In_ColumnNameMonth, In_ColumnNameQuarter, In_DeltaPrediction, In_Заголовок_столбца_офисы, In_ColumnNameMonth_смещение_План, In_ColumnNameQuarter_смещение_План, In_PlanMonth, In_PlanQuarter, In_Fact_Plan_displacement_Month, In_Fact_Plan_displacement_Quarter)
Dim dateDB As Date
    
  ' ***
  ' In_ColumnNameMonth - наименование столбца с планом месяца, например "Премия, тыс.руб._Месяц" для "3.6 ИСЖ_МАСС". Если планов на месяц нет, то In_ColumnNameMonth=""
  ' In_ColumnNameQuarter - наименование столбца с планом квартала, например "Премия, тыс.руб._Квартал" для "3.6 ИСЖ_МАСС"
  ' In_DeltaPrediction - + число столбцов от столбца План (месяца или квартала) в котором находится прогноз выполнения в %, например для "3.6 ИСЖ_МАСС" In_DeltaPrediction=3 ("План", "Факт" (+1), "% Вып-е" (+2), "% Вып-е_Прог" (+3) ). Если столбца "Прогноз" нет, то In_DeltaPrediction = 0
  ' In_Заголовок_столбца_офисы - наименование заголовка на листе, под которым идут филиалы: Алтайский ОО1, Архангельский ОО1, Астраханский ОО1 ...
  ' In_ColumnNameMonth_смещение_План - смещение относительно столбца In_ColumnNameMonth через которое выходим на "План месяца", например для "3.6 ИСЖ_МАСС" это смещение = 0, а для "3.5.1 ДВС" при In_ColumnNameMonth="Портфель, тыс.руб._Месяц" чтобы выйти на "ДВС_Итого-План" нужно In_ColumnNameMonth_смещение_План=12
  ' In_ColumnNameQuarter_смещение_План - смещение относительно столбца In_ColumnNameQuarter через которое выходим на "План квартала", например для, например для "3.6 ИСЖ_МАСС" это смещение = 0, а для "3.5.1 ДВС" при In_ColumnNameMonth="Портфель, тыс.руб._Квартал" чтобы выйти на "ДВС_Итого-План" нужно In_ColumnNameMonth_смещение_План=12
  ' In_PlanMonth - значение плана месяц цифрой, например 80% проникновения в страховки. Если 0, то берем из DB. Примечание - смещение In_ColumnNameMonth_смещение_План тогда = -1
  ' In_PlanQuarter - значение плана квартала цифрой, например 80% проникновения в страховки. Если 0, то берем из DB. Примечание - смещение In_ColumnNameQuarter_смещение_План = -1
  ' In_Fact_Plan_displacement_Month - смещение Факта относительно плана по Месяцу. По умолчанию = 1
  ' In_Fact_Plan_displacement_Quarter - смещение Факта относительно плана по Кварталу. По умолчанию = 1
  ' ***
    
  ' Дата DB
  dateDB = CDate(Mid(Workbooks(In_ReportName_String).Sheets("Оглавление").Cells(1, 1).Value, 23, 10))
  ' Дата DB с Лист8 (должны совпадать)
  ' dateDB_Лист8 = CDate(Mid(ThisWorkbook.Sheets("Лист8").Range("B5").Value, 52, 10))

  ' Апдейтим таблицу BASE\Products
  ' Call Update_BASE_Products(In_Product_Name, In_Product_Code, In_Unit)
  
  ' Вкладка In_Sheets
  ' 42
  Row_Заголовок_столбца_офисы = rowByValue(In_ReportName_String, In_Sheets, In_Заголовок_столбца_офисы, 300, 300) ' было 1000 1000
  ' 2
  Column_Заголовок_столбца_офисы = ColumnByValue(In_ReportName_String, In_Sheets, In_Заголовок_столбца_офисы, 300, 300)
  
  ' Выдачи_тыс_руб_Месяц - столбец "Выдачи, тыс.руб._Месяц" (в строке "Показатель")
  If In_ColumnNameMonth <> "" Then
    
    ' План (BK) 63
    Column_Продажи_Месяц_План = ColumnByValue(In_ReportName_String, In_Sheets, In_ColumnNameMonth, 500, 500) + In_ColumnNameMonth_смещение_План  ' "Выдачи, тыс.руб._Месяц" было 1000 1000
    ' Функция ColumnByValue3 - без удаления пробелов в строке поиска. Попробовал - не работает на ОФЗ! Вернул
    ' Column_Продажи_Месяц_План = ColumnByValue3(In_ReportName_String, In_Sheets, In_ColumnNameMonth, 500, 500) + In_ColumnNameMonth_смещение_План  ' "Выдачи, тыс.руб._Месяц" было 1000 1000
    
    ' Если столбец не найден - выдаем сообщение:
    If Column_Продажи_Месяц_План = 0 Then
      
      ' Заносим StringInSheet в переменную Строка_нет_листа_в_DB
      If InStr(Строка_нет_столбца_на_листе_в_DB, In_ColumnNameMonth) = 0 Then
    
        Строка_нет_столбца_на_листе_в_DB = Строка_нет_столбца_на_листе_в_DB + In_ColumnNameMonth + ", "
        ' Выводим сообщение
        MsgBox ("Внимание! По " + In_Product_Name + " не найден " + In_ColumnNameMonth + "!")

      End If
    
    End If
    
    ' Факт (BL) 64
    ' Column_Продажи_Месяц_Факт = Column_Продажи_Месяц_План + 1
    Column_Продажи_Месяц_Факт = Column_Продажи_Месяц_План + In_Fact_Plan_displacement_Month
    
    ' Прогноз (BO) 67
    If In_DeltaPrediction <> 0 Then
      Column_Продажи_Месяц_Прогноз = Column_Продажи_Месяц_План + In_DeltaPrediction ' (+ 4) параметр In_DeltaPrediction - это через сколько столбец с прогнозом в %
    End If
    
  End If
  
  ' Выдачи_тыс_руб_Квартал - столбец "Выдачи, тыс.руб._Квартал" (в строке "Показатель")
  ' План (CP) 94
  Column_Продажи_Квартал_План = ColumnByValue(In_ReportName_String, In_Sheets, In_ColumnNameQuarter, 500, 500) + In_ColumnNameQuarter_смещение_План ' "Выдачи, тыс.руб._Квартал" было 1000 1000
  ' Без удаления пробелов в поиске - ColumnByValue3. Не работает на ОФЗ, вернул!
  ' Column_Продажи_Квартал_План = ColumnByValue3(In_ReportName_String, In_Sheets, In_ColumnNameQuarter, 500, 500) + In_ColumnNameQuarter_смещение_План ' "Выдачи, тыс.руб._Квартал" было 1000 1000
  
  
  ' Факт (CQ) 95
  ' Column_Продажи_Квартал_Факт = Column_Продажи_Квартал_План + 1
  Column_Продажи_Квартал_Факт = Column_Продажи_Квартал_План + In_Fact_Plan_displacement_Quarter
   
  ' Прогноз (CT) 98
  If In_DeltaPrediction <> 0 Then
    Column_Продажи_Квартал_Прогноз = Column_Продажи_Квартал_План + In_DeltaPrediction ' (+ 4) параметр In_DeltaPrediction - это через сколько столбец с прогнозом в %
  End If
  
  ' Заносим наименование продукта на Лист8
  ThisWorkbook.Sheets("UpdFr_DB").Cells(In_Row_UpdFr_DB, 1).NumberFormat = "@"
  ThisWorkbook.Sheets("UpdFr_DB").Cells(In_Row_UpdFr_DB, 1).Value = In_Row_UpdFr_DB - 8 'In_N
  ThisWorkbook.Sheets("UpdFr_DB").Cells(In_Row_UpdFr_DB, 1).HorizontalAlignment = xlCenter
  ' Офис
  ThisWorkbook.Sheets("UpdFr_DB").Cells(In_Row_UpdFr_DB, 2).Value = In_officeNameInReport ' In_Product_Name
  ThisWorkbook.Sheets("UpdFr_DB").Cells(In_Row_UpdFr_DB, 2).HorizontalAlignment = xlLeft
  

  ' Контрольный показатель
  Офис_найден = False

  ' Находим в с столбце "Тюменский ОО1"
  rowCount = Row_Заголовок_столбца_офисы + 1
  Do While (InStr(Workbooks(In_ReportName_String).Sheets(In_Sheets).Cells(rowCount, Column_Заголовок_столбца_офисы).Value, "Общий итог") = 0) And (Not IsEmpty(Workbooks(In_ReportName_String).Sheets(In_Sheets).Cells(rowCount, Column_Заголовок_столбца_офисы).Value))
    
    ' Если это "Тюменский ОО1" - Раскрываем список
    If InStr(Workbooks(In_ReportName_String).Sheets(In_Sheets).Cells(rowCount, Column_Заголовок_столбца_офисы).Value, "Тюменский ОО1") <> 0 Then
      Workbooks(In_ReportName_String).Sheets(In_Sheets).Cells(rowCount, Column_Заголовок_столбца_офисы).ShowDetail = True
      Офис_найден = True
    End If
              
    ' Если это текущий офис
    If (InStr(Workbooks(In_ReportName_String).Sheets(In_Sheets).Cells(rowCount, Column_Заголовок_столбца_офисы).Value, In_officeNameInReport) <> 0) And (InStr(Workbooks(In_ReportName_String).Sheets(In_Sheets).Cells(rowCount, Column_Заголовок_столбца_офисы).Value, "ОО1") = 0) Then
      
      ' Берем из этой строки данные и копируем на Лист8
      
      ' Квартал:
      ' If (In_ColumnNameQuarter <> "") Then
      If (In_ColumnNameQuarter <> "") And (Column_Продажи_Квартал_План <> 0) Then ' 21.09 для обработки прошлых DB
        
        ' Квартал - план
        If In_PlanQuarter = 0 Then
          ' ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 5).Value = Workbooks(In_ReportName_String).Sheets(In_Sheets).Cells(rowCount, Column_Продажи_Квартал_План).Value
        Else
          ' ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 5).Value = In_PlanQuarter
        End If
        ' ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 5).NumberFormat = "#,##0"
        ' ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 5).HorizontalAlignment = xlRight
        

        ' Квартал - факт
        ' Если измерение в %
        If In_Unit <> "%" Then
          
          ' Квартал факт
          ' ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 6).Value = Workbooks(In_ReportName_String).Sheets(In_Sheets).Cells(rowCount, Column_Продажи_Квартал_Факт).Value
          
        Else
          ' Если это %, то умножаем на 100
          ' ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 6).Value = (Workbooks(In_ReportName_String).Sheets(In_Sheets).Cells(rowCount, Column_Продажи_Квартал_Факт).Value * 100)
        End If
        ' ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 6).NumberFormat = "#,##0"
        ' ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 6).HorizontalAlignment = xlRight
        

        ' Квартал - исполнение (в %)
        ' ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 7).Value = РассчетДоли(ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 5).Value, ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 6).Value, 3)
        ' ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 7).NumberFormat = "0%"
        ' ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 7).HorizontalAlignment = xlRight
        
        ' Если столбца "Прогноз" нет (In_DeltaPrediction = 0), то Окраска ячейки СФЕТОФОР ' 70. Заливка ячейки цветом "светофор" - значение In_Value в %, In_Target в %
        ' If (In_DeltaPrediction = 0) And (ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 5).Value <> 0) Then
          ' Окраска ячейки СФЕТОФОР ' 70. Заливка ячейки цветом "светофор" - значение In_Value в %, In_Target в %
        '   Call Full_Color_RangeII("Лист8", In_Row_Лист8, 7, ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 7).Value, 1)
        ' End If
      
        ' Квартал - прогноз
        ' If (In_DeltaPrediction <> 0) And (Workbooks(In_ReportName_String).Sheets(In_Sheets).Cells(rowCount, Column_Продажи_Квартал_План).Value <> 0) Then
        '   ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 8).Value = Workbooks(In_ReportName_String).Sheets(In_Sheets).Cells(rowCount, Column_Продажи_Квартал_Прогноз).Value
        '   ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 8).NumberFormat = "0%"
        '   ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 8).HorizontalAlignment = xlRight
          ' Окраска ячейки СФЕТОФОР ' 70. Заливка ячейки цветом "светофор" - значение In_Value в %, In_Target в %
        '   Call Full_Color_RangeII("Лист8", In_Row_Лист8, 8, ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 8).Value, 1)
        ' End If
        
        ' ***
        ' Тестирование Функции "Прогноз_квартала" по всем позициям, если измерение не в %
        ' If ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 4).Value <> "%" Then
        '   ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 20).Value = Прогноз_квартала_проц(dateDB, ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 5).Value, ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 6).Value, 5, 0)
        '   ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 20).NumberFormat = "0%"
        '   ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 20).HorizontalAlignment = xlRight
        ' End If
        ' ***
        
        ' Если по продукту есть квартальный показатель, критерий: (In_ColumnNameMonth = "") AND (In_ColumnNameQuarter <>"")
        ' If (In_ColumnNameMonth = "") And (In_ColumnNameQuarter <> "") Then
        
          ' Заносим в Sales_Office
          '  Идентификатор ID_Rec:
          ' ID_RecVar = CStr(CStr(getNumberOfficeByName(In_officeNameInReport)) + "-" + strNQYY(dateDB) + "-" + In_Product_Code)
                        
          ' Текущие значения в месяце: Date_01 (N), Date_02 (O), Date_03 (P), Date_04 (Q), Date_05 Date_06 Date_07 Date_08 Date_09 Date_10 Date_11 Date_12 Date_13 Date_14 Date_15 Date_16 Date_17 Date_18 Date_19 Date_20 Date_21 Date_22 Date_23 Date_24 Date_25 Date_26 Date_27 Date_28 Date_29 Date_30 Date_31
          ' Номер месяца в квартале: 1-"", 2-"2", 3-"3"
          ' M_num = Nom_mes_quarter_str(dateDB)
          ' curr_Day_Month_Q = "Date" + M_num + "_" + Mid(dateDB, 1, 2)
                                      
          ' Вносим данные в BASE\Sales_Office по ПК.
          ' Call InsertRecordInBook("Sales_Office", "Лист1", "ID_Rec", ID_RecVar, _
          '                                   "ID_Rec", ID_RecVar, _
          '                                     "Оffice_Number", getNumberOfficeByName(In_officeNameInReport), _
          '                                       "Product_Name", In_Product_Name, _
          '                                         "Оffice", In_officeNameInReport, _
          '                                           "MMYY", strNQYY(dateDB), _
          '                                             "Update_Date", dateDB, _
          '                                              "Product_Code", In_Product_Code, _
          '                                                "Plan", ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 5).Value, _
          '                                                   "Unit", In_Unit, _
          '                                                     "Fact", ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 6).Value, _
          '                                                       "Percent_Completion", ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 7).Value, _
          '                                                         curr_Day_Month_Q, ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 6).Value, _
          '                                                          "", "", _
          '                                                            "", "", _
          '                                                              "", "", _
          '                                                                "", "", _
          '                                                                  "", "", _
          '                                                                    "", "", _
          '                                                                      "", "", _
          '                                                                        "", "")

        
        ' End If
        
        
        
      End If
                  
      ' Месяц:
      If (In_ColumnNameMonth <> "") And (Column_Продажи_Месяц_План <> 0) Then ' 21.09 для обработки прошлых DB
  
        ' Месяц - план
        If In_PlanMonth = 0 Then
          ThisWorkbook.Sheets("UpdFr_DB").Cells(In_Row_UpdFr_DB, 3).Value = Workbooks(In_ReportName_String).Sheets(In_Sheets).Cells(rowCount, Column_Продажи_Месяц_План).Value
        Else
          ThisWorkbook.Sheets("UpdFr_DB").Cells(In_Row_UpdFr_DB, 3).Value = In_PlanMonth
        End If
        ThisWorkbook.Sheets("UpdFr_DB").Cells(In_Row_UpdFr_DB, 3).NumberFormat = "#,##0"
        ThisWorkbook.Sheets("UpdFr_DB").Cells(In_Row_UpdFr_DB, 3).HorizontalAlignment = xlRight
        
        
        ' Месяц - факт
        ' Если измерение в %
        If In_Unit <> "%" Then
          ThisWorkbook.Sheets("UpdFr_DB").Cells(In_Row_UpdFr_DB, 4).Value = Workbooks(In_ReportName_String).Sheets(In_Sheets).Cells(rowCount, Column_Продажи_Месяц_Факт).Value
        Else
          ' Если это %, то умножаем на 100
          ThisWorkbook.Sheets("UpdFr_DB").Cells(In_Row_UpdFr_DB, 4).Value = (Workbooks(In_ReportName_String).Sheets(In_Sheets).Cells(rowCount, Column_Продажи_Месяц_Факт).Value * 100)
        End If
        ThisWorkbook.Sheets("UpdFr_DB").Cells(In_Row_UpdFr_DB, 4).NumberFormat = "#,##0"
        ThisWorkbook.Sheets("UpdFr_DB").Cells(In_Row_UpdFr_DB, 4).HorizontalAlignment = xlRight
            
        ' Месяц - исполнение
        ThisWorkbook.Sheets("UpdFr_DB").Cells(In_Row_UpdFr_DB, 5).Value = РассчетДоли(ThisWorkbook.Sheets("UpdFr_DB").Cells(In_Row_UpdFr_DB, 3).Value, ThisWorkbook.Sheets("UpdFr_DB").Cells(In_Row_UpdFr_DB, 4).Value, 3)
        ThisWorkbook.Sheets("UpdFr_DB").Cells(In_Row_UpdFr_DB, 5).NumberFormat = "0%"
        ThisWorkbook.Sheets("UpdFr_DB").Cells(In_Row_UpdFr_DB, 5).HorizontalAlignment = xlRight
        ' Если столбца "Прогноз" нет (In_DeltaPrediction = 0), то Окраска ячейки СФЕТОФОР ' 70. Заливка ячейки цветом "светофор" - значение In_Value в %, In_Target в %
        If (In_DeltaPrediction = 0) And (ThisWorkbook.Sheets("UpdFr_DB").Cells(In_Row_UpdFr_DB, 4).Value <> 0) Then
          ' Окраска ячейки СФЕТОФОР ' 70. Заливка ячейки цветом "светофор" - значение In_Value в %, In_Target в %
          Call Full_Color_RangeII("UpdFr_DB", In_Row_UpdFr_DB, 5, ThisWorkbook.Sheets("UpdFr_DB").Cells(In_Row_UpdFr_DB, 5).Value, 1)
        End If

        ' Месяц - прогноз (штуки, тыс.руб и т.п.) делаем расчет
        If (In_DeltaPrediction <> 0) And (Workbooks(In_ReportName_String).Sheets(In_Sheets).Cells(rowCount, Column_Продажи_Месяц_План).Value <> 0) Then
      
          PredictionVar = (ThisWorkbook.Sheets("UpdFr_DB").Cells(In_Row_UpdFr_DB, 3).Value) * Workbooks(In_ReportName_String).Sheets(In_Sheets).Cells(rowCount, Column_Продажи_Месяц_Прогноз).Value
                
          ' Месяц - прогноз, %
          ThisWorkbook.Sheets("UpdFr_DB").Cells(In_Row_UpdFr_DB, 6).Value = Workbooks(In_ReportName_String).Sheets(In_Sheets).Cells(rowCount, Column_Продажи_Месяц_Прогноз).Value
          ThisWorkbook.Sheets("UpdFr_DB").Cells(In_Row_UpdFr_DB, 6).NumberFormat = "0%"
          ThisWorkbook.Sheets("UpdFr_DB").Cells(In_Row_UpdFr_DB, 6).HorizontalAlignment = xlRight
          PredictionPercent = ThisWorkbook.Sheets("UpdFr_DB").Cells(In_Row_UpdFr_DB, 6).Value
          ' Окраска ячейки СФЕТОФОР ' 70. Заливка ячейки цветом "светофор" - значение In_Value в %, In_Target в %
          Call Full_Color_RangeII("UpdFr_DB", In_Row_UpdFr_DB, 6, ThisWorkbook.Sheets("UpdFr_DB").Cells(In_Row_UpdFr_DB, 6).Value, 1)
        Else
          
          ' Если прогноза нет по продукту в DB
          PredictionVar = 0
          PredictionPercent = 0
        
        End If
      
        '  Идентификатор ID_Rec:
        ID_RecVar = CStr(CStr(getNumberOfficeByName(In_officeNameInReport)) + "-" + strMMYY(dateDB) + "-" + In_Product_Code)
      
        ' ID_Rec
        ThisWorkbook.Sheets("UpdFr_DB").Cells(In_Row_UpdFr_DB, 7).Value = ID_RecVar
        ThisWorkbook.Sheets("UpdFr_DB").Cells(In_Row_UpdFr_DB, 7).NumberFormat = "@"
        ThisWorkbook.Sheets("UpdFr_DB").Cells(In_Row_UpdFr_DB, 7).HorizontalAlignment = xlLeft
        
        ' Product_Name
        ThisWorkbook.Sheets("UpdFr_DB").Cells(In_Row_UpdFr_DB, 8).Value = In_Product_Name
        ThisWorkbook.Sheets("UpdFr_DB").Cells(In_Row_UpdFr_DB, 8).NumberFormat = "@"
        ThisWorkbook.Sheets("UpdFr_DB").Cells(In_Row_UpdFr_DB, 8).HorizontalAlignment = xlLeft
        
        ' Оffice
        ThisWorkbook.Sheets("UpdFr_DB").Cells(In_Row_UpdFr_DB, 9).Value = In_officeNameInReport
        ThisWorkbook.Sheets("UpdFr_DB").Cells(In_Row_UpdFr_DB, 9).NumberFormat = "@"
        ThisWorkbook.Sheets("UpdFr_DB").Cells(In_Row_UpdFr_DB, 9).HorizontalAlignment = xlLeft
        
        ' MMYY
        ' t = strMMYY(dateDB)
        ThisWorkbook.Sheets("UpdFr_DB").Cells(In_Row_UpdFr_DB, 10).Value = strMMYY(dateDB)
        ThisWorkbook.Sheets("UpdFr_DB").Cells(In_Row_UpdFr_DB, 10).NumberFormat = "@"
        ThisWorkbook.Sheets("UpdFr_DB").Cells(In_Row_UpdFr_DB, 10).HorizontalAlignment = xlCenter
        
        ' Update_Date
        ThisWorkbook.Sheets("UpdFr_DB").Cells(In_Row_UpdFr_DB, 11).Value = dateDB
        ThisWorkbook.Sheets("UpdFr_DB").Cells(In_Row_UpdFr_DB, 11).NumberFormat = "m/d/yyyy"
        ThisWorkbook.Sheets("UpdFr_DB").Cells(In_Row_UpdFr_DB, 11).HorizontalAlignment = xlLeft
        
        ' Product_Code
        ThisWorkbook.Sheets("UpdFr_DB").Cells(In_Row_UpdFr_DB, 12).Value = In_Product_Code
        ThisWorkbook.Sheets("UpdFr_DB").Cells(In_Row_UpdFr_DB, 12).NumberFormat = "@"
        ThisWorkbook.Sheets("UpdFr_DB").Cells(In_Row_UpdFr_DB, 12).HorizontalAlignment = xlLeft
      
        ' Заносим в Sales_Office
            
        ' Текущие значения в месяце: Date_01 (N), Date_02 (O), Date_03 (P), Date_04 (Q), Date_05 Date_06 Date_07 Date_08 Date_09 Date_10 Date_11 Date_12 Date_13 Date_14 Date_15 Date_16 Date_17 Date_18 Date_19 Date_20 Date_21 Date_22 Date_23 Date_24 Date_25 Date_26 Date_27 Date_28 Date_29 Date_30 Date_31
        curr_Day_Month = "Date_" + Mid(dateDB, 1, 2)
            
        ' Вносим данные в BASE\Sales_Office по ПК.
        ' Call InsertRecordInBook("Sales_Office", "Лист1", "ID_Rec", ID_RecVar, _
        '                                     "ID_Rec", ID_RecVar, _
        '                                       "Оffice_Number", getNumberOfficeByName(In_officeNameInReport), _
        '                                         "Product_Name", In_Product_Name, _
        '                                          "Оffice", In_officeNameInReport, _
        '                                             "MMYY", strMMYY(dateDB), _
        '                                               "Update_Date", dateDB, _
        '                                                "Product_Code", In_Product_Code, _
        '                                                  "Plan", ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 9).Value, _
        '                                                     "Unit", In_Unit, _
        '                                                       "Fact", ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 10).Value, _
        '                                                         "Percent_Completion", ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 11).Value, _
        '                                                           "Prediction", PredictionVar, _
        '                                                             "Percent_Prediction", PredictionPercent, _
        '                                                               curr_Day_Month, ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 10).Value, _
        '                                                                 "", "", _
        '                                                                   "", "", _
        '                                                                     "", "", _
        '                                                                       "", "", _
        '                                                                         "", "", _
        '                                                                           "", "")

      End If ' If In_ColumnNameMonth <> "" Then
      
    End If
    
    ' Следующая запись
    Application.StatusBar = In_Product_Code + " " + In_officeNameInReport + ": " + CStr(rowCount) + "..."
    rowCount = rowCount + 1
    DoEventsInterval (rowCount)
    
  Loop
  
  ' Контрольный показатель - если все 4 равны нулю, то данные из DB взяты не корректно
  If (Офис_найден = False) Then
    
    ' Если в DB Лист не найден
    MsgBox ("Внимание! По " + In_Product_Name + " не найдены Офисы!")

  End If

  
End Sub

' Показатель из вкладки DB (для Ипотеки)
Sub DB_UniversalSheetInDB_UpdFr_DB2(In_ReportName_String, In_Sheets, In_officeNameInReport, In_Row_Лист8, In_N, In_Product_Name, In_Product_Code, In_Unit, In_Weight, In_ColumnNameMonth, In_ColumnNameQuarter, In_DeltaPrediction, In_Заголовок_столбца_офисы, In_ColumnNameMonth_смещение_План, In_ColumnNameQuarter_смещение_План, In_PlanMonth, In_PlanQuarter, In_Fact_Plan_displacement_Month, In_Fact_Plan_displacement_Quarter)
Dim dateDB As Date
    
  ' ***
  ' In_ColumnNameMonth - наименование столбца с планом месяца, например "Премия, тыс.руб._Месяц" для "3.6 ИСЖ_МАСС". Если планов на месяц нет, то In_ColumnNameMonth=""
  ' In_ColumnNameQuarter - наименование столбца с планом квартала, например "Премия, тыс.руб._Квартал" для "3.6 ИСЖ_МАСС"
  ' In_DeltaPrediction - + число столбцов от столбца План (месяца или квартала) в котором находится прогноз выполнения в %, например для "3.6 ИСЖ_МАСС" In_DeltaPrediction=3 ("План", "Факт" (+1), "% Вып-е" (+2), "% Вып-е_Прог" (+3) ). Если столбца "Прогноз" нет, то In_DeltaPrediction = 0
  ' In_Заголовок_столбца_офисы - наименование заголовка на листе, под которым идут филиалы: Алтайский ОО1, Архангельский ОО1, Астраханский ОО1 ...
  ' In_ColumnNameMonth_смещение_План - смещение относительно столбца In_ColumnNameMonth через которое выходим на "План месяца", например для "3.6 ИСЖ_МАСС" это смещение = 0, а для "3.5.1 ДВС" при In_ColumnNameMonth="Портфель, тыс.руб._Месяц" чтобы выйти на "ДВС_Итого-План" нужно In_ColumnNameMonth_смещение_План=12
  ' In_ColumnNameQuarter_смещение_План - смещение относительно столбца In_ColumnNameQuarter через которое выходим на "План квартала", например для, например для "3.6 ИСЖ_МАСС" это смещение = 0, а для "3.5.1 ДВС" при In_ColumnNameMonth="Портфель, тыс.руб._Квартал" чтобы выйти на "ДВС_Итого-План" нужно In_ColumnNameMonth_смещение_План=12
  ' In_PlanMonth - значение плана месяц цифрой, например 80% проникновения в страховки. Если 0, то берем из DB. Примечание - смещение In_ColumnNameMonth_смещение_План тогда = -1
  ' In_PlanQuarter - значение плана квартала цифрой, например 80% проникновения в страховки. Если 0, то берем из DB. Примечание - смещение In_ColumnNameQuarter_смещение_План = -1
  ' In_Fact_Plan_displacement_Month - смещение Факта относительно плана по Месяцу. По умолчанию = 1
  ' In_Fact_Plan_displacement_Quarter - смещение Факта относительно плана по Кварталу. По умолчанию = 1
  ' ***
    
  ' Дата DB
  dateDB = CDate(Mid(Workbooks(In_ReportName_String).Sheets("Оглавление").Cells(1, 1).Value, 23, 10))
  ' Дата DB с Лист8 (должны совпадать)
  ' dateDB_Лист8 = CDate(Mid(ThisWorkbook.Sheets("Лист8").Range("B5").Value, 52, 10))

  ' Апдейтим таблицу BASE\Products
  ' Call Update_BASE_Products(In_Product_Name, In_Product_Code, In_Unit)
  
  ' Вкладка In_Sheets
  ' 42
  Row_Заголовок_столбца_офисы = rowByValue(In_ReportName_String, In_Sheets, In_Заголовок_столбца_офисы, 300, 300) ' было 1000 1000
  ' 2
  Column_Заголовок_столбца_офисы = ColumnByValue(In_ReportName_String, In_Sheets, In_Заголовок_столбца_офисы, 300, 300)
  
  ' Выдачи_тыс_руб_Месяц - столбец "Выдачи, тыс.руб._Месяц" (в строке "Показатель")
  If In_ColumnNameMonth <> "" Then
    
    ' План (BK) 63
    Column_Продажи_Месяц_План = ColumnByValue(In_ReportName_String, In_Sheets, In_ColumnNameMonth, 500, 500) + In_ColumnNameMonth_смещение_План  ' "Выдачи, тыс.руб._Месяц" было 1000 1000
    ' Функция ColumnByValue3 - без удаления пробелов в строке поиска. Попробовал - не работает на ОФЗ! Вернул
    ' Column_Продажи_Месяц_План = ColumnByValue3(In_ReportName_String, In_Sheets, In_ColumnNameMonth, 500, 500) + In_ColumnNameMonth_смещение_План  ' "Выдачи, тыс.руб._Месяц" было 1000 1000
    
    ' Если столбец не найден - выдаем сообщение:
    If Column_Продажи_Месяц_План = 0 Then
      
      ' Заносим StringInSheet в переменную Строка_нет_листа_в_DB
      If InStr(Строка_нет_столбца_на_листе_в_DB, In_ColumnNameMonth) = 0 Then
    
        Строка_нет_столбца_на_листе_в_DB = Строка_нет_столбца_на_листе_в_DB + In_ColumnNameMonth + ", "
        ' Выводим сообщение
        MsgBox ("Внимание! По " + In_Product_Name + " не найден " + In_ColumnNameMonth + "!")

      End If
    
    End If
    
    ' Факт (BL) 64
    ' Column_Продажи_Месяц_Факт = Column_Продажи_Месяц_План + 1
    Column_Продажи_Месяц_Факт = Column_Продажи_Месяц_План + In_Fact_Plan_displacement_Month
    
    ' Прогноз (BO) 67
    If In_DeltaPrediction <> 0 Then
      Column_Продажи_Месяц_Прогноз = Column_Продажи_Месяц_План + In_DeltaPrediction ' (+ 4) параметр In_DeltaPrediction - это через сколько столбец с прогнозом в %
    End If
    
  End If
  
  ' Выдачи_тыс_руб_Квартал - столбец "Выдачи, тыс.руб._Квартал" (в строке "Показатель")
  ' План (CP) 94
  Column_Продажи_Квартал_План = ColumnByValue(In_ReportName_String, In_Sheets, In_ColumnNameQuarter, 500, 500) + In_ColumnNameQuarter_смещение_План ' "Выдачи, тыс.руб._Квартал" было 1000 1000
  ' Без удаления пробелов в поиске - ColumnByValue3. Не работает на ОФЗ, вернул!
  ' Column_Продажи_Квартал_План = ColumnByValue3(In_ReportName_String, In_Sheets, In_ColumnNameQuarter, 500, 500) + In_ColumnNameQuarter_смещение_План ' "Выдачи, тыс.руб._Квартал" было 1000 1000
  
  
  ' Факт (CQ) 95
  ' Column_Продажи_Квартал_Факт = Column_Продажи_Квартал_План + 1
  Column_Продажи_Квартал_Факт = Column_Продажи_Квартал_План + In_Fact_Plan_displacement_Quarter
   
  ' Прогноз (CT) 98
  If In_DeltaPrediction <> 0 Then
    Column_Продажи_Квартал_Прогноз = Column_Продажи_Квартал_План + In_DeltaPrediction ' (+ 4) параметр In_DeltaPrediction - это через сколько столбец с прогнозом в %
  End If
  
  ' Заносим наименование продукта на Лист8
  ThisWorkbook.Sheets("UpdFr_DB").Cells(In_Row_UpdFr_DB, 1).NumberFormat = "@"
  ThisWorkbook.Sheets("UpdFr_DB").Cells(In_Row_UpdFr_DB, 1).Value = In_Row_UpdFr_DB - 8 'In_N
  ThisWorkbook.Sheets("UpdFr_DB").Cells(In_Row_UpdFr_DB, 1).HorizontalAlignment = xlCenter
  ' Офис
  ThisWorkbook.Sheets("UpdFr_DB").Cells(In_Row_UpdFr_DB, 2).Value = In_officeNameInReport ' In_Product_Name
  ThisWorkbook.Sheets("UpdFr_DB").Cells(In_Row_UpdFr_DB, 2).HorizontalAlignment = xlLeft
  

  ' Контрольный показатель
  Офис_найден = False

  ' Находим в с столбце "Тюменский ОО1"
  rowCount = Row_Заголовок_столбца_офисы + 1
  Do While (InStr(Workbooks(In_ReportName_String).Sheets(In_Sheets).Cells(rowCount, Column_Заголовок_столбца_офисы).Value, "Общий итог") = 0) And (Not IsEmpty(Workbooks(In_ReportName_String).Sheets(In_Sheets).Cells(rowCount, Column_Заголовок_столбца_офисы).Value))
    
    ' Если это "Тюменский ОО1" - Раскрываем список
    ' If InStr(Workbooks(In_ReportName_String).Sheets(In_Sheets).Cells(rowCount, Column_Заголовок_столбца_офисы).Value, "Тюменский ОО1") <> 0 Then
    If InStr(Workbooks(In_ReportName_String).Sheets(In_Sheets).Cells(rowCount, Column_Заголовок_столбца_офисы).Value, In_officeNameInReport) <> 0 Then
    '   Workbooks(In_ReportName_String).Sheets(In_Sheets).Cells(rowCount, Column_Заголовок_столбца_офисы).ShowDetail = True
      Офис_найден = True
    End If
              
    ' Если это текущий офис
    ' If (InStr(Workbooks(In_ReportName_String).Sheets(In_Sheets).Cells(rowCount, Column_Заголовок_столбца_офисы).Value, In_officeNameInReport) <> 0) And (InStr(Workbooks(In_ReportName_String).Sheets(In_Sheets).Cells(rowCount, Column_Заголовок_столбца_офисы).Value, "ОО1") = 0) Then
    If (InStr(Workbooks(In_ReportName_String).Sheets(In_Sheets).Cells(rowCount, Column_Заголовок_столбца_офисы).Value, In_officeNameInReport) <> 0) Then
      ' Берем из этой строки данные и копируем на Лист8
      
      ' Квартал:
      ' If (In_ColumnNameQuarter <> "") Then
      If (In_ColumnNameQuarter <> "") And (Column_Продажи_Квартал_План <> 0) Then ' 21.09 для обработки прошлых DB
        
        ' Квартал - план
        If In_PlanQuarter = 0 Then
          ' ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 5).Value = Workbooks(In_ReportName_String).Sheets(In_Sheets).Cells(rowCount, Column_Продажи_Квартал_План).Value
        Else
          ' ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 5).Value = In_PlanQuarter
        End If
        ' ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 5).NumberFormat = "#,##0"
        ' ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 5).HorizontalAlignment = xlRight
        

        ' Квартал - факт
        ' Если измерение в %
        If In_Unit <> "%" Then
          
          ' Квартал факт
          ' ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 6).Value = Workbooks(In_ReportName_String).Sheets(In_Sheets).Cells(rowCount, Column_Продажи_Квартал_Факт).Value
          
        Else
          ' Если это %, то умножаем на 100
          ' ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 6).Value = (Workbooks(In_ReportName_String).Sheets(In_Sheets).Cells(rowCount, Column_Продажи_Квартал_Факт).Value * 100)
        End If
        ' ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 6).NumberFormat = "#,##0"
        ' ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 6).HorizontalAlignment = xlRight
        

        ' Квартал - исполнение (в %)
        ' ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 7).Value = РассчетДоли(ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 5).Value, ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 6).Value, 3)
        ' ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 7).NumberFormat = "0%"
        ' ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 7).HorizontalAlignment = xlRight
        
        ' Если столбца "Прогноз" нет (In_DeltaPrediction = 0), то Окраска ячейки СФЕТОФОР ' 70. Заливка ячейки цветом "светофор" - значение In_Value в %, In_Target в %
        ' If (In_DeltaPrediction = 0) And (ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 5).Value <> 0) Then
          ' Окраска ячейки СФЕТОФОР ' 70. Заливка ячейки цветом "светофор" - значение In_Value в %, In_Target в %
        '   Call Full_Color_RangeII("Лист8", In_Row_Лист8, 7, ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 7).Value, 1)
        ' End If
      
        ' Квартал - прогноз
        ' If (In_DeltaPrediction <> 0) And (Workbooks(In_ReportName_String).Sheets(In_Sheets).Cells(rowCount, Column_Продажи_Квартал_План).Value <> 0) Then
        '   ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 8).Value = Workbooks(In_ReportName_String).Sheets(In_Sheets).Cells(rowCount, Column_Продажи_Квартал_Прогноз).Value
        '   ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 8).NumberFormat = "0%"
        '   ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 8).HorizontalAlignment = xlRight
          ' Окраска ячейки СФЕТОФОР ' 70. Заливка ячейки цветом "светофор" - значение In_Value в %, In_Target в %
        '   Call Full_Color_RangeII("Лист8", In_Row_Лист8, 8, ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 8).Value, 1)
        ' End If
        
        ' ***
        ' Тестирование Функции "Прогноз_квартала" по всем позициям, если измерение не в %
        ' If ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 4).Value <> "%" Then
        '   ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 20).Value = Прогноз_квартала_проц(dateDB, ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 5).Value, ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 6).Value, 5, 0)
        '   ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 20).NumberFormat = "0%"
        '   ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 20).HorizontalAlignment = xlRight
        ' End If
        ' ***
        
        ' Если по продукту есть квартальный показатель, критерий: (In_ColumnNameMonth = "") AND (In_ColumnNameQuarter <>"")
        ' If (In_ColumnNameMonth = "") And (In_ColumnNameQuarter <> "") Then
        
          ' Заносим в Sales_Office
          '  Идентификатор ID_Rec:
          ' ID_RecVar = CStr(CStr(getNumberOfficeByName(In_officeNameInReport)) + "-" + strNQYY(dateDB) + "-" + In_Product_Code)
                        
          ' Текущие значения в месяце: Date_01 (N), Date_02 (O), Date_03 (P), Date_04 (Q), Date_05 Date_06 Date_07 Date_08 Date_09 Date_10 Date_11 Date_12 Date_13 Date_14 Date_15 Date_16 Date_17 Date_18 Date_19 Date_20 Date_21 Date_22 Date_23 Date_24 Date_25 Date_26 Date_27 Date_28 Date_29 Date_30 Date_31
          ' Номер месяца в квартале: 1-"", 2-"2", 3-"3"
          ' M_num = Nom_mes_quarter_str(dateDB)
          ' curr_Day_Month_Q = "Date" + M_num + "_" + Mid(dateDB, 1, 2)
                                      
          ' Вносим данные в BASE\Sales_Office по ПК.
          ' Call InsertRecordInBook("Sales_Office", "Лист1", "ID_Rec", ID_RecVar, _
          '                                   "ID_Rec", ID_RecVar, _
          '                                     "Оffice_Number", getNumberOfficeByName(In_officeNameInReport), _
          '                                       "Product_Name", In_Product_Name, _
          '                                         "Оffice", In_officeNameInReport, _
          '                                           "MMYY", strNQYY(dateDB), _
          '                                             "Update_Date", dateDB, _
          '                                              "Product_Code", In_Product_Code, _
          '                                                "Plan", ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 5).Value, _
          '                                                   "Unit", In_Unit, _
          '                                                     "Fact", ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 6).Value, _
          '                                                       "Percent_Completion", ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 7).Value, _
          '                                                         curr_Day_Month_Q, ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 6).Value, _
          '                                                          "", "", _
          '                                                            "", "", _
          '                                                              "", "", _
          '                                                                "", "", _
          '                                                                  "", "", _
          '                                                                    "", "", _
          '                                                                      "", "", _
          '                                                                        "", "")

        
        ' End If
        
        
        
      End If
                  
      ' Месяц:
      If (In_ColumnNameMonth <> "") And (Column_Продажи_Месяц_План <> 0) Then ' 21.09 для обработки прошлых DB
  
        ' Месяц - план
        If In_PlanMonth = 0 Then
          ThisWorkbook.Sheets("UpdFr_DB").Cells(In_Row_UpdFr_DB, 3).Value = Workbooks(In_ReportName_String).Sheets(In_Sheets).Cells(rowCount, Column_Продажи_Месяц_План).Value
        Else
          ThisWorkbook.Sheets("UpdFr_DB").Cells(In_Row_UpdFr_DB, 3).Value = In_PlanMonth
        End If
        ThisWorkbook.Sheets("UpdFr_DB").Cells(In_Row_UpdFr_DB, 3).NumberFormat = "#,##0"
        ThisWorkbook.Sheets("UpdFr_DB").Cells(In_Row_UpdFr_DB, 3).HorizontalAlignment = xlRight
        
        
        ' Месяц - факт
        ' Если измерение в %
        If In_Unit <> "%" Then
          ThisWorkbook.Sheets("UpdFr_DB").Cells(In_Row_UpdFr_DB, 4).Value = Workbooks(In_ReportName_String).Sheets(In_Sheets).Cells(rowCount, Column_Продажи_Месяц_Факт).Value
        Else
          ' Если это %, то умножаем на 100
          ThisWorkbook.Sheets("UpdFr_DB").Cells(In_Row_UpdFr_DB, 4).Value = (Workbooks(In_ReportName_String).Sheets(In_Sheets).Cells(rowCount, Column_Продажи_Месяц_Факт).Value * 100)
        End If
        ThisWorkbook.Sheets("UpdFr_DB").Cells(In_Row_UpdFr_DB, 4).NumberFormat = "#,##0"
        ThisWorkbook.Sheets("UpdFr_DB").Cells(In_Row_UpdFr_DB, 4).HorizontalAlignment = xlRight
            
        ' Месяц - исполнение
        ThisWorkbook.Sheets("UpdFr_DB").Cells(In_Row_UpdFr_DB, 5).Value = РассчетДоли(ThisWorkbook.Sheets("UpdFr_DB").Cells(In_Row_UpdFr_DB, 3).Value, ThisWorkbook.Sheets("UpdFr_DB").Cells(In_Row_UpdFr_DB, 4).Value, 3)
        ThisWorkbook.Sheets("UpdFr_DB").Cells(In_Row_UpdFr_DB, 5).NumberFormat = "0%"
        ThisWorkbook.Sheets("UpdFr_DB").Cells(In_Row_UpdFr_DB, 5).HorizontalAlignment = xlRight
        ' Если столбца "Прогноз" нет (In_DeltaPrediction = 0), то Окраска ячейки СФЕТОФОР ' 70. Заливка ячейки цветом "светофор" - значение In_Value в %, In_Target в %
        If (In_DeltaPrediction = 0) And (ThisWorkbook.Sheets("UpdFr_DB").Cells(In_Row_UpdFr_DB, 4).Value <> 0) Then
          ' Окраска ячейки СФЕТОФОР ' 70. Заливка ячейки цветом "светофор" - значение In_Value в %, In_Target в %
          Call Full_Color_RangeII("UpdFr_DB", In_Row_UpdFr_DB, 5, ThisWorkbook.Sheets("UpdFr_DB").Cells(In_Row_UpdFr_DB, 5).Value, 1)
        End If

        ' Месяц - прогноз (штуки, тыс.руб и т.п.) делаем расчет
        If (In_DeltaPrediction <> 0) And (Workbooks(In_ReportName_String).Sheets(In_Sheets).Cells(rowCount, Column_Продажи_Месяц_План).Value <> 0) Then
      
          PredictionVar = (ThisWorkbook.Sheets("UpdFr_DB").Cells(In_Row_UpdFr_DB, 3).Value) * Workbooks(In_ReportName_String).Sheets(In_Sheets).Cells(rowCount, Column_Продажи_Месяц_Прогноз).Value
                
          ' Месяц - прогноз, %
          ThisWorkbook.Sheets("UpdFr_DB").Cells(In_Row_UpdFr_DB, 6).Value = Workbooks(In_ReportName_String).Sheets(In_Sheets).Cells(rowCount, Column_Продажи_Месяц_Прогноз).Value
          ThisWorkbook.Sheets("UpdFr_DB").Cells(In_Row_UpdFr_DB, 6).NumberFormat = "0%"
          ThisWorkbook.Sheets("UpdFr_DB").Cells(In_Row_UpdFr_DB, 6).HorizontalAlignment = xlRight
          PredictionPercent = ThisWorkbook.Sheets("UpdFr_DB").Cells(In_Row_UpdFr_DB, 6).Value
          ' Окраска ячейки СФЕТОФОР ' 70. Заливка ячейки цветом "светофор" - значение In_Value в %, In_Target в %
          Call Full_Color_RangeII("UpdFr_DB", In_Row_UpdFr_DB, 6, ThisWorkbook.Sheets("UpdFr_DB").Cells(In_Row_UpdFr_DB, 6).Value, 1)
        Else
          
          ' Если прогноза нет по продукту в DB
          PredictionVar = 0
          PredictionPercent = 0
        
        End If
      
        '  Идентификатор ID_Rec:
        ID_RecVar = CStr(CStr(getNumberOfficeByName(In_officeNameInReport)) + "-" + strMMYY(dateDB) + "-" + In_Product_Code)
      
        ' ID_Rec
        ThisWorkbook.Sheets("UpdFr_DB").Cells(In_Row_UpdFr_DB, 7).Value = ID_RecVar
        ThisWorkbook.Sheets("UpdFr_DB").Cells(In_Row_UpdFr_DB, 7).NumberFormat = "@"
        ThisWorkbook.Sheets("UpdFr_DB").Cells(In_Row_UpdFr_DB, 7).HorizontalAlignment = xlLeft
        
        ' Product_Name
        ThisWorkbook.Sheets("UpdFr_DB").Cells(In_Row_UpdFr_DB, 8).Value = In_Product_Name
        ThisWorkbook.Sheets("UpdFr_DB").Cells(In_Row_UpdFr_DB, 8).NumberFormat = "@"
        ThisWorkbook.Sheets("UpdFr_DB").Cells(In_Row_UpdFr_DB, 8).HorizontalAlignment = xlLeft
        
        ' Оffice
        ThisWorkbook.Sheets("UpdFr_DB").Cells(In_Row_UpdFr_DB, 9).Value = In_officeNameInReport
        ThisWorkbook.Sheets("UpdFr_DB").Cells(In_Row_UpdFr_DB, 9).NumberFormat = "@"
        ThisWorkbook.Sheets("UpdFr_DB").Cells(In_Row_UpdFr_DB, 9).HorizontalAlignment = xlLeft
        
        ' MMYY
        ' t = strMMYY(dateDB)
        ThisWorkbook.Sheets("UpdFr_DB").Cells(In_Row_UpdFr_DB, 10).Value = strMMYY(dateDB)
        ThisWorkbook.Sheets("UpdFr_DB").Cells(In_Row_UpdFr_DB, 10).NumberFormat = "@"
        ThisWorkbook.Sheets("UpdFr_DB").Cells(In_Row_UpdFr_DB, 10).HorizontalAlignment = xlCenter
        
        ' Update_Date
        ThisWorkbook.Sheets("UpdFr_DB").Cells(In_Row_UpdFr_DB, 11).Value = dateDB
        ThisWorkbook.Sheets("UpdFr_DB").Cells(In_Row_UpdFr_DB, 11).NumberFormat = "m/d/yyyy"
        ThisWorkbook.Sheets("UpdFr_DB").Cells(In_Row_UpdFr_DB, 11).HorizontalAlignment = xlLeft
        
        ' Product_Code
        ThisWorkbook.Sheets("UpdFr_DB").Cells(In_Row_UpdFr_DB, 12).Value = In_Product_Code
        ThisWorkbook.Sheets("UpdFr_DB").Cells(In_Row_UpdFr_DB, 12).NumberFormat = "@"
        ThisWorkbook.Sheets("UpdFr_DB").Cells(In_Row_UpdFr_DB, 12).HorizontalAlignment = xlLeft
      
        ' Заносим в Sales_Office
            
        ' Текущие значения в месяце: Date_01 (N), Date_02 (O), Date_03 (P), Date_04 (Q), Date_05 Date_06 Date_07 Date_08 Date_09 Date_10 Date_11 Date_12 Date_13 Date_14 Date_15 Date_16 Date_17 Date_18 Date_19 Date_20 Date_21 Date_22 Date_23 Date_24 Date_25 Date_26 Date_27 Date_28 Date_29 Date_30 Date_31
        curr_Day_Month = "Date_" + Mid(dateDB, 1, 2)
            
        ' Вносим данные в BASE\Sales_Office по ПК.
        ' Call InsertRecordInBook("Sales_Office", "Лист1", "ID_Rec", ID_RecVar, _
        '                                     "ID_Rec", ID_RecVar, _
        '                                       "Оffice_Number", getNumberOfficeByName(In_officeNameInReport), _
        '                                         "Product_Name", In_Product_Name, _
        '                                          "Оffice", In_officeNameInReport, _
        '                                             "MMYY", strMMYY(dateDB), _
        '                                               "Update_Date", dateDB, _
        '                                                "Product_Code", In_Product_Code, _
        '                                                  "Plan", ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 9).Value, _
        '                                                     "Unit", In_Unit, _
        '                                                       "Fact", ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 10).Value, _
        '                                                         "Percent_Completion", ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 11).Value, _
        '                                                           "Prediction", PredictionVar, _
        '                                                             "Percent_Prediction", PredictionPercent, _
        '                                                               curr_Day_Month, ThisWorkbook.Sheets("Лист8").Cells(In_Row_Лист8, 10).Value, _
        '                                                                 "", "", _
        '                                                                   "", "", _
        '                                                                     "", "", _
        '                                                                       "", "", _
        '                                                                         "", "", _
        '                                                                           "", "")

      End If ' If In_ColumnNameMonth <> "" Then
      
    End If
    
    ' Следующая запись
    Application.StatusBar = In_Product_Code + " " + In_officeNameInReport + ": " + CStr(rowCount) + "..."
    rowCount = rowCount + 1
    DoEventsInterval (rowCount)
    
  Loop
  
  ' Контрольный показатель - если все 4 равны нулю, то данные из DB взяты не корректно
  If (Офис_найден = False) Then
    
    ' Если в DB Лист не найден
    MsgBox ("Внимание! По " + In_Product_Name + " не найдены Офисы!")

  End If

  
End Sub



' Добавить данные из DB
Sub Добавить_данные_с_UpdFr_DB_в_Sales_Office()

  ' Запрос
  If MsgBox("Добавить данные с листа UpdFr_DB в BASE\Sales_Office?", vbYesNo) = vbYes Then
    
    ' Статус
    ThisWorkbook.Sheets("UpdFr_DB").Range("C6").Value = ""
    
    ' Открываем BASE\Sales
    OpenBookInBase ("Sales_Office")
    
    ' Открываем BASE\Products
    OpenBookInBase ("Products")

    
    ' Определяем столбцы: #ID_Rec #Product_Name   #Оffice #MMYY   #Update_Date    #Product_Code
    column_ID_Rec = ColumnByValue(ThisWorkbook.Name, "UpdFr_DB", "#ID_Rec", 100, 100)
    column_Product_Name = ColumnByValue(ThisWorkbook.Name, "UpdFr_DB", "#Product_Name", 100, 100)
    column_Оffice = ColumnByValue(ThisWorkbook.Name, "UpdFr_DB", "#Оffice", 100, 100)
    column_MMYY = ColumnByValue(ThisWorkbook.Name, "UpdFr_DB", "#MMYY", 100, 100)
    column_Update_Date = ColumnByValue(ThisWorkbook.Name, "UpdFr_DB", "#Update_Date", 100, 100)
    column_Product_Code = ColumnByValue(ThisWorkbook.Name, "UpdFr_DB", "#Product_Code", 100, 100)
    
    ' #План   #Факт   #Исп.   #Прогноз '     column_ = ColumnByValue(ThisWorkbook.Name, "UpdFr_DB", "", 100, 100)
    column_План = ColumnByValue(ThisWorkbook.Name, "UpdFr_DB", "#План", 100, 100)
    column_Факт = ColumnByValue(ThisWorkbook.Name, "UpdFr_DB", "#Факт", 100, 100)
    column_Исп = ColumnByValue(ThisWorkbook.Name, "UpdFr_DB", "#Исп.", 100, 100)
    column_Прогноз = ColumnByValue(ThisWorkbook.Name, "UpdFr_DB", "#Прогноз", 100, 100)

    
    ' Находим строку Форма UpdFr_DB_1
    rowCount = rowByValue(ThisWorkbook.Name, "UpdFr_DB", "Форма UpdFr_DB_1", 100, 100) + 3
    Do While (ThisWorkbook.Sheets("UpdFr_DB").Cells(rowCount, 2).Value <> "")
    
       
    
      ' Вставляем данные
      ' Вносим данные в BASE\Sales_Office по ПК.
      Call InsertRecordInBook("Sales_Office", "Лист1", "ID_Rec", ThisWorkbook.Sheets("UpdFr_DB").Cells(rowCount, column_ID_Rec).Value, _
                                           "ID_Rec", ThisWorkbook.Sheets("UpdFr_DB").Cells(rowCount, column_ID_Rec).Value, _
                                             "Оffice_Number", getNumberOfficeByName(ThisWorkbook.Sheets("UpdFr_DB").Cells(rowCount, column_Оffice).Value), _
                                               "Product_Name", Product_Code_to_Product_Name(ThisWorkbook.Sheets("UpdFr_DB").Cells(rowCount, column_Product_Code).Value), _
                                                "Оffice", ThisWorkbook.Sheets("UpdFr_DB").Cells(rowCount, column_Оffice).Value, _
                                                   "MMYY", ThisWorkbook.Sheets("UpdFr_DB").Cells(rowCount, column_MMYY).Value, _
                                                     "Update_Date", ThisWorkbook.Sheets("UpdFr_DB").Cells(rowCount, column_Update_Date).Value, _
                                                      "Product_Code", ThisWorkbook.Sheets("UpdFr_DB").Cells(rowCount, column_Product_Code).Value, _
                                                        "Plan", ThisWorkbook.Sheets("UpdFr_DB").Cells(rowCount, column_План).Value, _
                                                           "Unit", Product_Name_to_Unit(Product_Code_to_Product_Name(ThisWorkbook.Sheets("UpdFr_DB").Cells(rowCount, column_Product_Code).Value)), _
                                                             "Fact", ThisWorkbook.Sheets("UpdFr_DB").Cells(rowCount, column_Факт).Value, _
                                                               "Percent_Completion", ThisWorkbook.Sheets("UpdFr_DB").Cells(rowCount, column_Исп).Value, _
                                                                 "Prediction", "", _
                                                                   "Percent_Prediction", "", _
                                                                     "", "", _
                                                                       "", "", _
                                                                         "", "", _
                                                                           "", "", _
                                                                             "", "", _
                                                                               "", "", _
                                                                                 "", "")
    
      ' Следующая запись
      Application.StatusBar = "Обработано: " + ThisWorkbook.Sheets("UpdFr_DB").Cells(rowCount, 2).Value + "..."
      rowCount = rowCount + 1
      DoEventsInterval (rowCount)
    
    Loop
      
    ' Статус
    ThisWorkbook.Sheets("UpdFr_DB").Range("C6").Value = "Статус: Данные за " + CStr(ThisWorkbook.Sheets("UpdFr_DB").Range("C7").Value) + " добавлены в Sales_Office"
      
    ' Закрываем BASE\Products
    CloseBook ("Products")
   
    ' Закрываем BASE\Sales
    CloseBook ("Sales_Office")
    
    Application.StatusBar = ""
    
    MsgBox ("Данные добавлены!")
  
  End If



End Sub

' Апдейт поля Исп. по Плану и Факт
Sub Update_UpdFr_DB_Исп()

  ' #План   #Факт   #Исп.   #Прогноз '     column_ = ColumnByValue(ThisWorkbook.Name, "UpdFr_DB", "", 100, 100)
  column_План = ColumnByValue(ThisWorkbook.Name, "UpdFr_DB", "#План", 100, 100)
  column_Факт = ColumnByValue(ThisWorkbook.Name, "UpdFr_DB", "#Факт", 100, 100)
  column_Исп = ColumnByValue(ThisWorkbook.Name, "UpdFr_DB", "#Исп.", 100, 100)

  ' Находим строку Форма UpdFr_DB_1
  rowCount = rowByValue(ThisWorkbook.Name, "UpdFr_DB", "Форма UpdFr_DB_1", 100, 100) + 3
  Do While (ThisWorkbook.Sheets("UpdFr_DB").Cells(rowCount, 2).Value <> "")
  
    ThisWorkbook.Sheets("UpdFr_DB").Cells(rowCount, 5).Value = ThisWorkbook.Sheets("UpdFr_DB").Cells(rowCount, 4).Value / ThisWorkbook.Sheets("UpdFr_DB").Cells(rowCount, 3).Value
    
    ' Окраска ячейки СФЕТОФОР ' 70. Заливка ячейки цветом "светофор" - значение In_Value в %, In_Target в %
    Call Full_Color_RangeII("UpdFr_DB", rowCount, 5, ThisWorkbook.Sheets("UpdFr_DB").Cells(rowCount, 5).Value, 1)
  
    ' Следующая запись
    Application.StatusBar = "Обработано: " + ThisWorkbook.Sheets("UpdFr_DB").Cells(rowCount, 2).Value + "..."
    rowCount = rowCount + 1
    DoEventsInterval (rowCount)
    
  Loop

  Application.StatusBar = ""

End Sub

' Апдейт поля Исп. по Плану и Факт
Sub Update_UpdFr_DB_Ипотека_Исп()

  ' #План   #Факт   #Исп.   #Прогноз '     column_ = ColumnByValue(ThisWorkbook.Name, "UpdFr_DB", "", 100, 100)
  column_План = ColumnByValue(ThisWorkbook.Name, "UpdFr_DB", "#План2", 100, 100)
  column_Факт = ColumnByValue(ThisWorkbook.Name, "UpdFr_DB", "#Факт2", 100, 100)
  column_Исп = ColumnByValue(ThisWorkbook.Name, "UpdFr_DB", "#Исп2", 100, 100)
  column_Прог = ColumnByValue(ThisWorkbook.Name, "UpdFr_DB", "#Прог2", 100, 100)

  ' Находим строку Форма UpdFr_DB_1
  rowCount = rowByValue(ThisWorkbook.Name, "UpdFr_DB", "Форма UpdFr_DB_2", 100, 100) + 3
  Do While (ThisWorkbook.Sheets("UpdFr_DB").Cells(rowCount, 2).Value <> "")
  
    ThisWorkbook.Sheets("UpdFr_DB").Cells(rowCount, 5).Value = ThisWorkbook.Sheets("UpdFr_DB").Cells(rowCount, 4).Value / ThisWorkbook.Sheets("UpdFr_DB").Cells(rowCount, 3).Value
    ThisWorkbook.Sheets("UpdFr_DB").Cells(rowCount, 6).Value = ThisWorkbook.Sheets("UpdFr_DB").Cells(rowCount, 4).Value / ThisWorkbook.Sheets("UpdFr_DB").Cells(rowCount, 3).Value
    
    ' Окраска ячейки СФЕТОФОР ' 70. Заливка ячейки цветом "светофор" - значение In_Value в %, In_Target в %
    Call Full_Color_RangeII("UpdFr_DB", rowCount, 6, ThisWorkbook.Sheets("UpdFr_DB").Cells(rowCount, 6).Value, 1)
  
    ' Следующая запись
    Application.StatusBar = "Обработано: " + ThisWorkbook.Sheets("UpdFr_DB").Cells(rowCount, 2).Value + "..."
    rowCount = rowCount + 1
    DoEventsInterval (rowCount)
    
  Loop

  Application.StatusBar = ""

End Sub



' Добавить данные из DB по ипотеке
Sub UpdateFrom_DB_Ипотека()

  ' Открыть файл с отчетом
  FileName = Application.GetOpenFilename("Excel Files (*.xlsm), *.xlsm", , "Открытие файла с отчетом")
    
  ' Выводим для инфо данные об имени файла
  ReportName_String = Dir(FileName)
  
  ' Открываем выбранную книгу (UpdateLinks:=0)
  Workbooks.Open FileName, 0
      
  ' Переходим на окно DB
  ThisWorkbook.Sheets("UpdFr_DB").Activate

  row_Форма_UpdFr_DB_2 = rowByValue(ThisWorkbook.Name, "UpdFr_DB", "Форма UpdFr_DB_2", 100, 100)

  ' Статус
  ThisWorkbook.Sheets("UpdFr_DB").Range("C" + CStr(row_Форма_UpdFr_DB_2)).Value = ""

  ' Очистить таблицу на Листе "UpdFr_DB"
  Call clearСontents2(ThisWorkbook.Name, "UpdFr_DB", "A" + CStr(row_Форма_UpdFr_DB_2 + 3), "L14" + CStr(row_Форма_UpdFr_DB_2 + 3))

  ' Дата DB
  dateDB_UpdFr_DB = CDate(Mid(Workbooks(ReportName_String).Sheets("Оглавление").Cells(1, 1).Value, 23, 10))
  ThisWorkbook.Sheets("UpdFr_DB").Range("C" + CStr(row_Форма_UpdFr_DB_2 + 1)).Value = CStr(dateDB_UpdFr_DB)

  ' Определяем столбец #Значение_переменной
  column_Значение_переменной = ColumnByValue(ThisWorkbook.Name, "UpdFr_DB", "#Значение_переменной2", 100, 100)

  ' Инициализация переменных
  In_ReportName_String_Var = ReportName_String ' ThisWorkbook.Sheets("UpdFr_DB").Cells(8, 9).Value
  
  SheetName_String_Var = ThisWorkbook.Sheets("UpdFr_DB").Cells(rowByValue(ThisWorkbook.Name, "UpdFr_DB", "SheetName_String2:", 100, 100), column_Значение_переменной).Value
  
  In_officeNameInReport = ThisWorkbook.Sheets("UpdFr_DB").Cells(rowByValue(ThisWorkbook.Name, "UpdFr_DB", "In_officeNameInReport2:", 100, 100), column_Значение_переменной).Value
  
  In_Product_Name_Var = ThisWorkbook.Sheets("UpdFr_DB").Cells(rowByValue(ThisWorkbook.Name, "UpdFr_DB", "In_Product_Name2:", 100, 100), column_Значение_переменной).Value ' "Выдачи ПК"
  
  In_Product_Code_Var = Product_Name_to_Product_Code(In_Product_Name_Var) ' "Выдачи_ПК_шт"
  
  In_Unit_Var = Product_Name_to_Unit(In_Product_Name_Var)

  In_ColumnNameMonth = ThisWorkbook.Sheets("UpdFr_DB").Cells(rowByValue(ThisWorkbook.Name, "UpdFr_DB", "In_ColumnNameMonth2:", 100, 100), column_Значение_переменной).Value
  
  In_ColumnNameQuarter = ThisWorkbook.Sheets("UpdFr_DB").Cells(rowByValue(ThisWorkbook.Name, "UpdFr_DB", "In_ColumnNameQuarter2:", 100, 100), column_Значение_переменной).Value

  In_DeltaPrediction = ThisWorkbook.Sheets("UpdFr_DB").Cells(rowByValue(ThisWorkbook.Name, "UpdFr_DB", "In_DeltaPrediction2:", 100, 100), column_Значение_переменной).Value

  In_Заголовок_столбца_офисы = ThisWorkbook.Sheets("UpdFr_DB").Cells(rowByValue(ThisWorkbook.Name, "UpdFr_DB", "In_Заголовок_столбца_офисы2:", 100, 100), column_Значение_переменной).Value

  ' Проверка наличия Листа в DB
  StringInSheet = SheetName_String_Var
  SheetName_String = FindNameSheet(ReportName_String, StringInSheet) ' "3.11 Зарплатные карты"
  If SheetName_String <> "" Then

    ' Переходим в DB на нужный Лист
    Workbooks(ReportName_String).Sheets(SheetName_String_Var).Activate

    ' Переходим на окно DB
    ThisWorkbook.Sheets("UpdFr_DB").Activate

            
      ' Номер строки для вывода на Листе UpdFr_DB
      In_Row_UpdFr_DB = row_Форма_UpdFr_DB_2 + 3
        
      ' Находим номер строки с наименованием офиса
      officeNameInReport_Var = In_officeNameInReport ' "Тюменский" ' officeNameInReport ' ThisWorkbook.Sheets("UpdFr_DB").Cells(8, 9).Value
        
      ' Поля: ID_Rec, Оffice_Number, Product_Name, Оffice, MMYY, Update_Date, Product_Code, Plan, Unit, Fact, Percent_Completion

      ' ПК Выдачи, шт.
      Call DB_UniversalSheetInDB_UpdFr_DB2(In_ReportName_String_Var, _
                                             SheetName_String_Var, _
                                               officeNameInReport_Var, _
                                                 0, _
                                                   0, _
                                                     In_Product_Name_Var, _
                                                       In_Product_Code_Var, _
                                                         In_Unit_Var, _
                                                           0, _
                                                             In_ColumnNameMonth, _
                                                               In_ColumnNameQuarter, _
                                                                 In_DeltaPrediction, _
                                                                   In_Заголовок_столбца_офисы, _
                                                                     0, _
                                                                       0, _
                                                                         0, _
                                                                           0, 1, 1)


    ' Статус
    ThisWorkbook.Sheets("UpdFr_DB").Range("C6").Value = "Статус: Данные за " + CStr(ThisWorkbook.Sheets("UpdFr_DB").Range("C7").Value) + " извлечены, проверьте итоги!"

  Else
    
    ' Сообщение
    MsgBox ("В DB не найден Лист " + SheetName_String_Var + "!")
    
  End If

    ' Закрываем файл с отчетом без сохранения изменений (параметр SaveChanges:=False)
    If MsgBox("Закрыть DB?", vbYesNo) = vbYes Then

      Workbooks(Dir(FileName)).Close SaveChanges:=False ' тестирование
      ' Переходим в ячейку M2
      ThisWorkbook.Sheets("UpdFr_DB").Range("A1").Select
    
    Else
    
      
    
    End If
    

    ' Сообщение
    MsgBox ("Обработка завершена!")
 
    Application.StatusBar = ""
  

End Sub

' Добавить данные из DB
Sub Добавить_данные_с_UpdFr_DB_Ипотека_в_Sales_Office()

  ' Запрос
  If MsgBox("Добавить данные с листа UpdFr_DB (Ипотека) в BASE\Sales_Office?", vbYesNo) = vbYes Then
    
    ' Статус
    row_Форма_UpdFr_DB_2 = rowByValue(ThisWorkbook.Name, "UpdFr_DB", "Форма UpdFr_DB_2", 100, 100)
    ThisWorkbook.Sheets("UpdFr_DB").Range("C" + CStr(row_Форма_UpdFr_DB_2)).Value = ""
    
    ' Открываем BASE\Sales
    OpenBookInBase ("Sales_Office")
    
    ' Открываем BASE\Products
    OpenBookInBase ("Products")

    
    ' Определяем столбцы: #ID_Rec #Product_Name   #Оffice #MMYY   #Update_Date    #Product_Code
    column_ID_Rec = ColumnByValue(ThisWorkbook.Name, "UpdFr_DB", "#ID_Rec2", 100, 100)
    column_Product_Name = ColumnByValue(ThisWorkbook.Name, "UpdFr_DB", "#Product_Name2", 100, 100)
    column_Оffice = ColumnByValue(ThisWorkbook.Name, "UpdFr_DB", "#Оffice2", 100, 100)
    column_MMYY = ColumnByValue(ThisWorkbook.Name, "UpdFr_DB", "#MMYY2", 100, 100)
    column_Update_Date = ColumnByValue(ThisWorkbook.Name, "UpdFr_DB", "#Update_Date2", 100, 100)
    column_Product_Code = ColumnByValue(ThisWorkbook.Name, "UpdFr_DB", "#Product_Code2", 100, 100)
    
    ' #План   #Факт   #Исп.   #Прогноз '     column_ = ColumnByValue(ThisWorkbook.Name, "UpdFr_DB", "", 100, 100)
    column_План = ColumnByValue(ThisWorkbook.Name, "UpdFr_DB", "#План2", 100, 100)
    column_Факт = ColumnByValue(ThisWorkbook.Name, "UpdFr_DB", "#Факт2", 100, 100)
    column_Исп = ColumnByValue(ThisWorkbook.Name, "UpdFr_DB", "#Исп2", 100, 100)
    column_Прогноз = ColumnByValue(ThisWorkbook.Name, "UpdFr_DB", "#Прог2", 100, 100)

    
    ' Находим строку Форма UpdFr_DB_1
    rowCount = rowByValue(ThisWorkbook.Name, "UpdFr_DB", "Форма UpdFr_DB_2", 100, 100) + 3
    Do While (ThisWorkbook.Sheets("UpdFr_DB").Cells(rowCount, 2).Value <> "")
    
       
    
      ' Вставляем данные
      ' Вносим данные в BASE\Sales_Office по ПК.
      Call InsertRecordInBook("Sales_Office", "Лист1", "ID_Rec", ThisWorkbook.Sheets("UpdFr_DB").Cells(rowCount, column_ID_Rec).Value, _
                                           "ID_Rec", ThisWorkbook.Sheets("UpdFr_DB").Cells(rowCount, column_ID_Rec).Value, _
                                             "Оffice_Number", getNumberOfficeByName(ThisWorkbook.Sheets("UpdFr_DB").Cells(rowCount, column_Оffice).Value), _
                                               "Product_Name", Product_Code_to_Product_Name(ThisWorkbook.Sheets("UpdFr_DB").Cells(rowCount, column_Product_Code).Value), _
                                                "Оffice", ThisWorkbook.Sheets("UpdFr_DB").Cells(rowCount, column_Оffice).Value, _
                                                   "MMYY", ThisWorkbook.Sheets("UpdFr_DB").Cells(rowCount, column_MMYY).Value, _
                                                     "Update_Date", ThisWorkbook.Sheets("UpdFr_DB").Cells(rowCount, column_Update_Date).Value, _
                                                      "Product_Code", ThisWorkbook.Sheets("UpdFr_DB").Cells(rowCount, column_Product_Code).Value, _
                                                        "Plan", ThisWorkbook.Sheets("UpdFr_DB").Cells(rowCount, column_План).Value, _
                                                           "Unit", Product_Name_to_Unit(Product_Code_to_Product_Name(ThisWorkbook.Sheets("UpdFr_DB").Cells(rowCount, column_Product_Code).Value)), _
                                                             "Fact", ThisWorkbook.Sheets("UpdFr_DB").Cells(rowCount, column_Факт).Value, _
                                                               "Percent_Completion", ThisWorkbook.Sheets("UpdFr_DB").Cells(rowCount, column_Исп).Value, _
                                                                 "Prediction", "", _
                                                                   "Percent_Prediction", "", _
                                                                     "", "", _
                                                                       "", "", _
                                                                         "", "", _
                                                                           "", "", _
                                                                             "", "", _
                                                                               "", "", _
                                                                                 "", "")
    
      ' Следующая запись
      Application.StatusBar = "Обработано: " + ThisWorkbook.Sheets("UpdFr_DB").Cells(rowCount, 2).Value + "..."
      rowCount = rowCount + 1
      DoEventsInterval (rowCount)
    
    Loop
      
    ' Статус
    ThisWorkbook.Sheets("UpdFr_DB").Range("C" + CStr(row_Форма_UpdFr_DB_2)).Value = "Статус: Данные за " + CStr(ThisWorkbook.Sheets("UpdFr_DB").Range("C" + CStr(row_Форма_UpdFr_DB_2 + 1)).Value) + " добавлены в Sales_Office"
      
    ' Закрываем BASE\Products
    CloseBook ("Products")
   
    ' Закрываем BASE\Sales
    CloseBook ("Sales_Office")
    
    Application.StatusBar = ""
    
    MsgBox ("Данные добавлены!")
  
  End If
  

End Sub

