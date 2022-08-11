Attribute VB_Name = "Module_Архив"
' *** Архивные копии процедур и функций ***

' Выгрузить файл с дневным планом продаж по форме обновленной форме (на основе формы Данилова)
' Templates\Ежедневная форма отчёта (куратор) 2.xlsx
Sub Архив_Выгрузить_план_дневных_продаж2()
Dim FileNewVar As String

  ' Формируем цели на день по форме Данилова
  
  ' Запрос на формирование
  If MsgBox("Сформировать поручения на день для офисов?", vbYesNo) = vbYes Then
      
    ' Открываем шаблон C:\Users\proschaevsf\Documents\#DB_Result\Templates\Ежедневная форма отчёта (куратор) 2.xlsx
    fileTemplatesName = "Ежедневная форма отчёта (куратор) 2.xlsx"
    Workbooks.Open (ThisWorkbook.Path + "\Templates\" + fileTemplatesName)
           
    ' Переходим на окно DB
    ThisWorkbook.Sheets("Лист8").Activate

    ' Дата формирования - если сегодня понедельник, то формируем за пятницу
    ' Если текущая дата это понедельник, то формируем отчет за пятницу
    If Weekday(CurrDate, vbMonday) = 1 Then
      dateReport = Date - 3
    Else
      dateReport = Date
    End If

    ' Имя нового файла
    FileNewVar = "Ежедневная_форма_отчёта_" + strДД_MM_YYYY(dateReport) + ".xlsx"
    Workbooks(fileTemplatesName).SaveAs FileName:=ThisWorkbook.Path + "\Out\" + FileNewVar, FileFormat:=xlOpenXMLWorkbook, createBackUp:=False
    
    ' Наименование листа в файле (TemplateSheets)
    TS = "Ежедневный отчет"
    
    ' Остаток рабочих дней определяем число рабочих дней с понеделника до конца месяца Working_days_between_dateReports(In_dateReportStart, In_dateReportEnd, In_working_days_in_the_week) As Integer
    Остаток_рабочих_дней = Working_days_between_dates(dateReport - 1, Date_last_day_month(dateReport), 5)

    ' Заголовок
    Workbooks(FileNewVar).Sheets(TS).Range("A1").Value = "Продажи за: " + CStr(dateReport) + " (ост.дней " + CStr(Остаток_рабочих_дней) + ")"

    ' Строка с именами файлов для архивирования
    strFileNewVar_Office = ""

    ' Проходим по Листу8 и заполняем планы:
    For i = 1 To 5
        ' Номера офисов от 1 до 5
        Select Case i
          Case 1 ' ОО «Тюменский»
            officeNameInReport = "ОО «Тюменский»"
          Case 2 ' ОО «Сургутский»
            officeNameInReport = "ОО «Сургутский»"
          Case 3 ' ОО «Нижневартовский»
            officeNameInReport = "ОО «Нижневартовский»"
          Case 4 ' ОО «Новоуренгойский»
            officeNameInReport = "ОО «Новоуренгойский»"
          Case 5 ' ОО «Тарко-Сале»
            officeNameInReport = "ОО «Тарко-Сале»"
        End Select
        
        ' Сообщение
        Application.StatusBar = "Формирование по " + officeNameInReport
                
        ' Текущая строка офиса в отчете
        row_TS = rowByValue(FileNewVar, TS, officeNameInReport, 100, 100)
        
        ' Обрабатываем столбцы в fileTemplatesName в горизонтальном направлении
        ColumnCount = 1
        Do While (ColumnCount <= 100)
          
          ' Если находим # в ячейке
          If InStr(Workbooks(FileNewVar).Sheets(TS).Cells(1, ColumnCount).Value, "#") <> 0 Then
            
            ' Текущий продукт в форме отчета
            currProductName = Mid(Workbooks(FileNewVar).Sheets(TS).Cells(1, ColumnCount).Value, 2)
            
            ' Находим Текущий продукт на Лист8 для текущего офиса
            Row_Лист8 = getRowFromSheet8(officeNameInReport, currProductName)
            
            ' Расчет плана дня
            If Round(((ThisWorkbook.Sheets("Лист8").Cells(Row_Лист8, 9).Value - ThisWorkbook.Sheets("Лист8").Cells(Row_Лист8, 10).Value) / Остаток_рабочих_дней), 0) > 0 Then
              Workbooks(FileNewVar).Sheets(TS).Cells(row_TS, ColumnCount).Value = Round(((ThisWorkbook.Sheets("Лист8").Cells(Row_Лист8, 9).Value - ThisWorkbook.Sheets("Лист8").Cells(Row_Лист8, 10).Value) / Остаток_рабочих_дней), 0)
            Else
              Workbooks(FileNewVar).Sheets(TS).Cells(row_TS, ColumnCount).Value = 0
            End If

            ' Формат ячейки плана
            Workbooks(FileNewVar).Sheets(TS).Cells(row_TS, ColumnCount).NumberFormat = "#,##0"
            
          End If ' Если находим # в ячейке
          
          ' Следующий столбец
          ' Application.StatusBar = In_Product_Code + " " + In_officeNameInReport + ": " + CStr(rowCount) + "..."
          ColumnCount = ColumnCount + 1
          DoEventsInterval (ColumnCount)
        Loop
                
        ' Строка статуса
        Application.StatusBar = "Сохранение " + officeNameInReport + "..."
                
        ' Сохранение изменений
        ' Workbooks(FileNewVar).Save
        
        ' Офис отработан - нарезаем в отдельный файл
        FileNewVar_Office = ThisWorkbook.Path + "\Out\Ежедневный_отчёт_" + cityOfficeNameByNumber(i) + "_" + strДД_MM_YYYY(dateReport) + ".xlsx"
        Workbooks(FileNewVar).SaveCopyAs FileName:=FileNewVar_Office

        ' Строка с именами файлов для архивирования
        strFileNewVar_Office = strFileNewVar_Office + FileNewVar_Office + " "

        ' Переходим на окно DB
        ThisWorkbook.Sheets("Лист8").Activate

        ' Строка статуса
        Application.StatusBar = ""

        
    Next i
    
    ' Закрываем файл
    Workbooks(FileNewVar).Close SaveChanges:=True

    Application.StatusBar = "Сформирован файл " + ThisWorkbook.Path + "\Out\" + FileNewVar

    Application.StatusBar = "Создание архива"

    ' Запускаем архиватор этого файла
    ' Работает Shell ("C:\Program Files\7-Zip\7z a -tzip -ssw -mx0 C:\Users\PROSCHAEVSF\Documents\#DB_Result\Out\Отчет.zip C:\Users\PROSCHAEVSF\Documents\#DB_Result\OUT\Ежедневный_отчёт_Тюмень_07-02-2021.xlsx C:\Users\PROSCHAEVSF\Documents\#DB_Result\Out\Отчет.zip C:\Users\PROSCHAEVSF\Documents\#DB_Result\OUT\Ежедневный_отчёт_Тарко-Сале_07-02-2021.xlsx")
    Shell ("C:\Program Files\7-Zip\7z a -tzip -ssw -mx9 C:\Users\PROSCHAEVSF\Documents\#DB_Result\Out\Ежедневная_форма_отчёта_" + strДД_MM_YYYY(dateReport) + ".zip " + strFileNewVar_Office)
    ' Имя файла архива
    File7zipName = "Ежедневная_форма_отчёта_" + strДД_MM_YYYY(dateReport) + ".zip"

    Application.StatusBar = "Архив создан!"

    MsgBox ("Сформирован файл " + ThisWorkbook.Path + "\Out\" + FileNewVar + "!")

    ' Отправка в почте в офисы
    ' Call Отправка_Lotus_Notes_Выгр_день_Лист8(ThisWorkbook.Path + "\Out\" + FileNewVar, DateReport)
    Call Отправка_Lotus_Notes_Выгр_день_Лист8(ThisWorkbook.Path + "\Out\" + File7zipName, dateReport)
      
    ' Строка статуса
    Application.StatusBar = ""
      
  End If
  
End Sub