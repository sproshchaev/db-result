Attribute VB_Name = "Module_ЕСУП"
' Выгрузка протокола собрания ЕСУП
Sub Протокол_Собрания()
  
Dim FileProtocolName, str_Присутствовавшие_на_Собрании, str_Отсутствовавшие_на_Собрании, str_копияВадрес, К_пор_Range, str_Поручениеi, Присутств_на_Собрании_Range, range_Список_получателей As String
Dim НомерСтроки_Повестка_дня, Номер_вопроса, rowCount, К_пор_Row, К_пор_Column, Номер_поручения, текущаяСтрокаПротокола, i, Присутств_на_Собрании_Row, Присутств_на_Собрании_Column As Byte
Dim row_column_Список_получателей, column_Список_получателей As Byte

  ' Закрытие Протокола (из макроса)
  ' Workbooks.Open FileName:="C:\Users\Сергей\Documents\#VBA\DB_Result\Templates\Приложение 1. Протокол.xlsx"
  ' ChDir "C:\Users\Сергей\Documents\#VBA\DB_Result\Out"
  ' ActiveWorkbook.SaveAs FileName:="C:\Users\Сергей\Documents\#VBA\DB_Result\Out\Протокол_собрания.xlsx", FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
  ' ActiveWindow.Close

  ' Запрос на формирование протокола Собрания
  If MsgBox("Сформировать протокол Собрания?", vbYesNo) = vbYes Then
    
    ' Открываем шаблон Протокола из C:\Users\...\Documents\#VBA\DB_Result\Templates
    Workbooks.Open (ThisWorkbook.Path + "\Templates\Приложение 1. Протокол.xlsx")
         
    ' Имя файла с протоколом - берем из G2 "10-02032020"
    FileProtocolName = "Протокол _ РОО Тюменский_" + CStr(dateProtocol(ThisWorkbook.Sheets("ЕСУП").Range("G2").Value)) + ".xlsx"
    Workbooks("Приложение 1. Протокол.xlsx").SaveAs FileName:=ThisWorkbook.Path + "\Out\" + FileProtocolName, FileFormat:=xlOpenXMLWorkbook, createBackUp:=False
    
    ' Номер протокола
    Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Range("C1:E2").Value = "Протокол Собрания №" + ThisWorkbook.Sheets("ЕСУП").Range("G2").Value
    Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Range("C1:G2").MergeCells = True
    ' Увеличиваем ширину предпоследнего столбца
    Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Columns("H:H").ColumnWidth = 20.43  ' 20.43, 21.64-предел
    Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Columns("I:I").ColumnWidth = 3
    ' Тема
    Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Range("A4:C4").HorizontalAlignment = xlCenter
    Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Range("A4:C4").VerticalAlignment = xlCenter
    ' Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Range("D4:H4").Value = "Еженедельный конференц-колл с Управляющими офисов и НОРПиКО Тюменского РОО по подведению итогов работы офисов за предыдущую неделю и постановке бизнес-целей на текущую неделю"
    Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Range("D4:H4").Value = "Еженедельный конференц-колл с Управляющими офисов и НОРПиКО Тюменского РОО по подведению итогов работы офисов за предыдущую неделю и постановке бизнес-целей на период с " + strDDMM(weekStartDate(dateProtocol(ThisWorkbook.Sheets("ЕСУП").Range("G2").Value))) + " по " + CStr(weekEndDate(dateProtocol(ThisWorkbook.Sheets("ЕСУП").Range("G2").Value))) + " г."
    ' Дата проведения
    Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Range("A5:C5").HorizontalAlignment = xlCenter
    Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Range("A5:C5").VerticalAlignment = xlCenter
    Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Range("D5:H5").Value = CStr(dateProtocol(ThisWorkbook.Sheets("ЕСУП").Range("G2").Value)) + " г."
    ' Место проведения
    Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Range("A6:C6").HorizontalAlignment = xlCenter
    Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Range("A6:C6").VerticalAlignment = xlCenter
    Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Range("D6:H6").Value = "г.Тюмень, ул.Советская 51/1"
    ' Участники присутствовали
    Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Range("A7:B8").HorizontalAlignment = xlCenter
    Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Range("A7:B8").VerticalAlignment = xlCenter
    Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Range("C7").HorizontalAlignment = xlCenter
    Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Range("C7").VerticalAlignment = xlCenter
    str_Присутствовавшие_на_Собрании = Присутствовавшие_на_Собрании(1)
    Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Range("D7:H7").Value = str_Присутствовавшие_на_Собрании
    Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Range("D7:H7").WrapText = True
    Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Range("7:7").RowHeight = lineHeight(str_Присутствовавшие_на_Собрании, 15, 60) ' было 50 - норм. 60 - реальная ширина
        
    ' Участники отсутствовали
    Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Range("C8").HorizontalAlignment = xlCenter
    Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Range("C8").VerticalAlignment = xlCenter
    str_Отсутствовавшие_на_Собрании = Присутствовавшие_на_Собрании(0)
    Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Range("D8:H8").Value = str_Отсутствовавшие_на_Собрании
    Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Range("D8:H8").WrapText = True
    Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Range("8:8").RowHeight = lineHeight(str_Отсутствовавшие_на_Собрании, 15, 60) ' было 40 - норм
    ' Копия в адрес:
    Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Range("A9:C9").HorizontalAlignment = xlCenter
    Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Range("A9:C9").VerticalAlignment = xlCenter
    str_копияВадрес = getFromAddrBook("РД", 1) + ", " + getFromAddrBook("КОП", 1) + ", " + getFromAddrBook("КИп", 1) + ", " + getFromAddrBook("ККП", 1) + ", " + getFromAddrBook("Кaf", 1) + ", " + getFromAddrBook("КПВО", 1)
    Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Range("D9:H9").Value = str_копияВадрес
    Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Range("D9:H9").WrapText = True
    Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Range("9:9").RowHeight = lineHeight(str_копияВадрес, 15, 60) ' было 30 - норм
    
    ' Повестка дня:
    НомерСтроки_Повестка_дня = ThisWorkbook.Sheets("ЕСУП").Range(RangeByValue(ThisWorkbook.Name, "ЕСУП", "Повестка_дня", 100, 100)).Row
    rowCount = 2
    Номер_вопроса = 0
    Do While ThisWorkbook.Sheets("ЕСУП").Cells(НомерСтроки_Повестка_дня + rowCount, 1).Value <> ""
      ' Если у вопроса стоит отметка "1", то вносим его в протокол собрания
      If ThisWorkbook.Sheets("ЕСУП").Cells(НомерСтроки_Повестка_дня + rowCount, 14).Value = "1" Then
        
        Номер_вопроса = Номер_вопроса + 1
        
        ' Если номер вопроса более 6-ти, то вставляем строку
        If Номер_вопроса > 6 Then
          ' Вставляем строку
          Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Range(CStr(12 + Номер_вопроса) + ":" + CStr(12 + Номер_вопроса)).Select
          Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
          ' Нумерация "7." возможна только если формат преобразовать к текстовому ("@")
          Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Cells(12 + Номер_вопроса, 1).NumberFormat = "@"
          Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Cells(12 + Номер_вопроса, 1).Value = CStr(Номер_вопроса) + "."
          Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Cells(12 + Номер_вопроса, 1).HorizontalAlignment = xlLeft
          Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Range("B" + CStr(12 + Номер_вопроса) + ":H" + CStr(12 + Номер_вопроса)).MergeCells = True
          ' Рамка
          Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Range("A" + CStr(12 + Номер_вопроса) + ":H" + CStr(12 + Номер_вопроса)).Select
          With Selection.Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ThemeColor = 3
            .TintAndShade = -0.749992370372631
            .Weight = xlThin
          End With
          With Selection.Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ThemeColor = 3
            .TintAndShade = -0.749992370372631
            .Weight = xlThin
          End With
          With Selection.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ThemeColor = 3
            .TintAndShade = -0.749992370372631
            .Weight = xlThin
          End With
          With Selection.Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ThemeColor = 3
            .TintAndShade = -0.749992370372631
            .Weight = xlThin
          End With
          With Selection.Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .ThemeColor = 3
            .TintAndShade = -0.749992370372631
            .Weight = xlThin
          End With
        
        End If
        
        ' Формат номера пункта Повестки Дня
        Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Cells(12 + Номер_вопроса, 1).HorizontalAlignment = xlCenter
        Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Cells(12 + Номер_вопроса, 1).VerticalAlignment = xlCenter
        ' Вносим в Повестку Дня
        If ThisWorkbook.Sheets("ЕСУП").Cells(НомерСтроки_Повестка_дня + rowCount, 13).Value = "" Then
          Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Cells(12 + Номер_вопроса, 2).Value = ThisWorkbook.Sheets("ЕСУП").Cells(НомерСтроки_Повестка_дня + rowCount, 2).Value + ": " + ThisWorkbook.Sheets("ЕСУП").Cells(НомерСтроки_Повестка_дня + rowCount, 3).Value
        Else
          ' Если есть Хэштэг
          Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Cells(12 + Номер_вопроса, 2).Value = ThisWorkbook.Sheets("ЕСУП").Cells(НомерСтроки_Повестка_дня + rowCount, 2).Value + ": " + ThisWorkbook.Sheets("ЕСУП").Cells(НомерСтроки_Повестка_дня + rowCount, 3).Value + " (" + ThisWorkbook.Sheets("ЕСУП").Cells(НомерСтроки_Повестка_дня + rowCount, 13).Value + ")"
        End If
        Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Cells(12 + Номер_вопроса, 2).HorizontalAlignment = xlLeft
        Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Cells(12 + Номер_вопроса, 2).VerticalAlignment = xlTop
        Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Cells(12 + Номер_вопроса, 2).WrapText = True
        Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Cells(12 + Номер_вопроса, 2).RowHeight = lineHeight(delSym(Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Cells(12 + Номер_вопроса, 2).Value), 15, 90) ' было 15, 65
      End If
      ' Следующая строка
      rowCount = rowCount + 1
    Loop
    
    ' Корректируем номер вопроса, если он менее 6 для того, чтобы корректно рассчитывать число строк в Поручениях
    If Номер_вопроса < 6 Then
      Номер_вопроса = 6
    End If
       
    ' Поручения участникам
    Номер_поручения = 0
    For i = 1 To 5
      ' Находим столбец К_порi
      К_пор_Range = RangeByValue(ThisWorkbook.Name, "ЕСУП", "К_пор" + CStr(i), 100, 100)
      К_пор_Row = ThisWorkbook.Sheets("ЕСУП").Range(К_пор_Range).Row
      К_пор_Column = ThisWorkbook.Sheets("ЕСУП").Range(К_пор_Range).Column
      ' Обрабатываем Поручения по офису, где стоят даты
      rowCount = К_пор_Row + 1
      Do While ThisWorkbook.Sheets("ЕСУП").Cells(rowCount, К_пор_Column - 6).Value <> ""
        
        ' Если в поле Дата не пусто (+ если План не 0), то выводим в поручение
        If (ThisWorkbook.Sheets("ЕСУП").Cells(rowCount, К_пор_Column + 1).Value <> "") And (ThisWorkbook.Sheets("ЕСУП").Cells(rowCount, К_пор_Column + 2).Value <> 0) Then
          ' Номер поручения
          Номер_поручения = Номер_поручения + 1
          ' Если номер поручения более 6-ти, то вставляем строку
          If Номер_поручения > 6 Then
            
            ' Вставляем пустую строку в блок "Поручения"
            Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Range(CStr(12 + Номер_вопроса + 4 + Номер_поручения) + ":" + CStr(12 + Номер_вопроса + 4 + Номер_поручения)).Select
            Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
            ' Нумерация "7." возможна только если формат преобразовать к текстовому ("@")
            Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Cells(12 + Номер_вопроса + 4 + Номер_поручения, 1).NumberFormat = "@"
            Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Cells(12 + Номер_вопроса + 4 + Номер_поручения, 1).Value = CStr(Номер_поручения) + "."
            ' Объединяем B, С, D
            Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Range("B" + CStr(12 + Номер_вопроса + 4 + Номер_поручения) + ":D" + CStr(12 + Номер_вопроса + 4 + Номер_поручения)).MergeCells = True
            ' Объединяем G, Н
            Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Range("G" + CStr(12 + Номер_вопроса + 4 + Номер_поручения) + ":H" + CStr(12 + Номер_вопроса + 4 + Номер_поручения)).MergeCells = True
            ' Рамка
            Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Range("A" + CStr(12 + Номер_вопроса + 4 + Номер_поручения) + ":H" + CStr(12 + Номер_вопроса + 4 + Номер_поручения)).Select
            With Selection.Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .ThemeColor = 3
                .TintAndShade = -0.749992370372631
                .Weight = xlThin
            End With
            With Selection.Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .ThemeColor = 3
                .TintAndShade = -0.749992370372631
                .Weight = xlThin
            End With
            With Selection.Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .ThemeColor = 3
                .TintAndShade = -0.749992370372631
                .Weight = xlThin
            End With
            With Selection.Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .ThemeColor = 3
                .TintAndShade = -0.749992370372631
                .Weight = xlThin
            End With
            With Selection.Borders(xlInsideVertical)
                .LineStyle = xlContinuous
                .ThemeColor = 3
                .TintAndShade = -0.749992370372631
                .Weight = xlThin
            End With

          End If ' Вставляем новую строку Поручения и нумеруем
          
          ' Номер Поручения (№ п/п) - выравнивание по центру и по вертикали
          Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Cells(12 + Номер_вопроса + 4 + Номер_поручения, 1).HorizontalAlignment = xlCenter
          Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Cells(12 + Номер_вопроса + 4 + Номер_поручения, 1).VerticalAlignment = xlCenter

          ' Поручение
          str_Поручениеi = getNameOfficeByNumber(i) + ": " + ThisWorkbook.Sheets("ЕСУП").Cells(rowCount, К_пор_Column - 5).Value + " " + CStr(ThisWorkbook.Sheets("ЕСУП").Cells(rowCount, К_пор_Column + 2).Value) + " " + ThisWorkbook.Sheets("ЕСУП").Cells(rowCount, К_пор_Column + 3).Value
          
          ' Поручение - "Переносить по словам"
          Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Range("B" + CStr(12 + Номер_вопроса + 4 + Номер_поручения) + ":D" + CStr(12 + Номер_вопроса + 4 + Номер_поручения)).WrapText = True
          ' Поручение - выравнивание по центру и по вертикали
          Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Cells(12 + Номер_вопроса + 4 + Номер_поручения, 2).HorizontalAlignment = xlLeft
          Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Cells(12 + Номер_вопроса + 4 + Номер_поручения, 2).VerticalAlignment = xlTop
          ' Поручение - высота строки
          Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Range(CStr(12 + Номер_вопроса + 4 + Номер_поручения) + ":" + CStr(12 + Номер_вопроса + 4 + Номер_поручения)).RowHeight = lineHeight(str_Поручениеi, 15, 37) ' 20 - норм
          ' Поручение - Запись в протокол
          Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Cells(12 + Номер_вопроса + 4 + Номер_поручения, 2).Value = str_Поручениеi

          ' Ответственный
          ' Вариант 1 - Должность и ФИО
          ' Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Cells(12 + Номер_вопроса + 4 + Номер_поручения, 5).Value = ThisWorkbook.Sheets("ЕСУП").Cells(К_пор_Row - 1, К_пор_Column - 4).Value + " " + ThisWorkbook.Sheets("ЕСУП").Cells(К_пор_Row - 1, К_пор_Column - 3).Value
          ' Вариант 2 - ФИО
          Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Cells(12 + Номер_вопроса + 4 + Номер_поручения, 5).Value = ThisWorkbook.Sheets("ЕСУП").Cells(К_пор_Row - 1, К_пор_Column - 3).Value
          
          ' Ответственный - "Переносить по словам"
          Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Range("E" + CStr(12 + Номер_вопроса + 4 + Номер_поручения) + ":E" + CStr(12 + Номер_вопроса + 4 + Номер_поручения)).WrapText = True
          ' Ответственный - выравнивание по вертикали и горизонтали
          Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Cells(12 + Номер_вопроса + 4 + Номер_поручения, 5).VerticalAlignment = xlTop
          Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Cells(12 + Номер_вопроса + 4 + Номер_поручения, 5).HorizontalAlignment = xlCenter
          
          ' Срок исполнения
          Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Cells(12 + Номер_вопроса + 4 + Номер_поручения, 6).Value = CStr(weekEndDate(CDate(ThisWorkbook.Sheets("ЕСУП").Cells(rowCount, К_пор_Column + 1).Value)))
          ' Срок исполнения - выравнивание по центру и по вертикали
          Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Cells(12 + Номер_вопроса + 4 + Номер_поручения, 6).HorizontalAlignment = xlCenter
          Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Cells(12 + Номер_вопроса + 4 + Номер_поручения, 6).VerticalAlignment = xlTop
          
        End If
        ' Следующая строка
        DoEvents
        rowCount = rowCount + 1
      Loop ' Do While
      
    Next i ' Следующий офис
    
    ' Вывести результаты исполнения предидущих поручений из BASE\
    
    ' Моя Подпись под протоколом
    текущаяСтрокаПротокола = (12 + Номер_вопроса + 4 + Номер_поручения) + 2
    Call InsertRow_InProtocol(FileProtocolName, "Протокол_Собрания", текущаяСтрокаПротокола)
    Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Cells(текущаяСтрокаПротокола, 2).Value = "Заместитель директора по развитию розничного бизнеса"
    Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Cells(текущаяСтрокаПротокола, 7).Value = "Прощаев С.Ф."
    текущаяСтрокаПротокола = текущаяСтрокаПротокола + 1
    Call InsertRow_InProtocol(FileProtocolName, "Протокол_Собрания", текущаяСтрокаПротокола)
    текущаяСтрокаПротокола = текущаяСтрокаПротокола + 1
    Call InsertRow_InProtocol(FileProtocolName, "Протокол_Собрания", текущаяСтрокаПротокола)
    
    ' С протоколом ознакомлены: (по электронной почте) - направляем присутствующим и отсутствующим
    Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Cells(текущаяСтрокаПротокола, 2).Value = "C протоколом ознакомлены (по электронной почте):"
    ' Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Cells(текущаяСтрокаПротокола, 2).Font.Underline = xlUnderlineStyleSingle
    текущаяСтрокаПротокола = текущаяСтрокаПротокола + 1
    Call InsertRow_InProtocol(FileProtocolName, "Протокол_Собрания", текущаяСтрокаПротокола)
    '
    Присутств_на_Собрании_Range = RangeByValue(ThisWorkbook.Name, "ЕСУП", "Присутств_на_Собрании", 100, 100)
    Присутств_на_Собрании_Row = ThisWorkbook.Sheets("ЕСУП").Range(Присутств_на_Собрании_Range).Row
    Присутств_на_Собрании_Column = ThisWorkbook.Sheets("ЕСУП").Range(Присутств_на_Собрании_Range).Column
    
    rowCount = Присутств_на_Собрании_Row + 1
    Do While ThisWorkbook.Sheets("ЕСУП").Cells(rowCount, Присутств_на_Собрании_Column).Value <> "Пригл_на_Собрание"
      
      ' Если ФИО <>0
      If ThisWorkbook.Sheets("ЕСУП").Cells(rowCount, Присутств_на_Собрании_Column + 1).Value <> 0 Then
        
        ' Должность_и_ФИО = ThisWorkbook.Sheets("ЕСУП").Cells(RowCount, Присутств_на_Собрании_Column + 4).Value + " " + Фамилия_и_Имя(ThisWorkbook.Sheets("ЕСУП").Cells(RowCount, Присутств_на_Собрании_Column + 1).Value, 3)
        Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Cells(текущаяСтрокаПротокола, 2).Value = ThisWorkbook.Sheets("ЕСУП").Cells(rowCount, Присутств_на_Собрании_Column + 4).Value
        Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Cells(текущаяСтрокаПротокола, 6).Value = Фамилия_и_Имя(ThisWorkbook.Sheets("ЕСУП").Cells(rowCount, Присутств_на_Собрании_Column + 1).Value, 3)
        текущаяСтрокаПротокола = текущаяСтрокаПротокола + 1
        Call InsertRow_InProtocol(FileProtocolName, "Протокол_Собрания", текущаяСтрокаПротокола)
      
      End If
      
      ' Следующая запись
      rowCount = rowCount + 1
    Loop
 
    ' Формируем список для отправки (в "Список получателей:"):
    range_Список_получателей = RangeByValue(ThisWorkbook.Name, "ЕСУП", "Список получателей:", 100, 100)
    row_Список_получателей = ThisWorkbook.Sheets("ЕСУП").Range(range_Список_получателей).Row
    column_Список_получателей = ThisWorkbook.Sheets("ЕСУП").Range(range_Список_получателей).Column
    '
    ThisWorkbook.Sheets("ЕСУП").Cells(row_Список_получателей, column_Список_получателей + 2).Value = getFromAddrBook("УДО2,УДО3,УДО4,УДО5,НОРПиКО1,НОРПиКО2,НОРПиКО3,НОРПиКО4,НОРПиКО5,ПМ,МРК1,МРК2,МРК3,МРК4,МРК5,НОКП,РРКК,МПП", 2)
    ThisWorkbook.Sheets("ЕСУП").Cells(row_Список_получателей, column_Список_получателей + 3).Value = " "
    
    ' Перемещаем в ячейку A1
    Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Range("A1").Select
    
    ' Закрытие файла с Протоколом Собрания
    Workbooks(FileProtocolName).Close SaveChanges:=True
    
    ' Редактируем тему письма
    ThisWorkbook.Sheets("ЕСУП").Cells(1, 17).Value = "Тюменский РОО: Протокол конференц-колла с офисами №" + ThisWorkbook.Sheets("ЕСУП").Cells(2, 7).Value + " от " + CStr(dateProtocol(ThisWorkbook.Sheets("ЕСУП").Range("G2").Value)) + " г."
    
    ' Редактируем поля с Тэгами 1 и 2 на листе Есуп
    ThisWorkbook.Sheets("ЕСУП").Cells(1, 15).Value = "#protocol"
    ThisWorkbook.Sheets("ЕСУП").Cells(3, 15).Value = "#protocol_" + ThisWorkbook.Sheets("ЕСУП").Cells(2, 7).Value
   
    MsgBox ("Протокол сформирован!")
    
  End If ' Запрос на формирование
  
End Sub

' Новый бланк Протокола собрания (Перенести в архив повестку)
Sub Новый_бланк_Протокола_Собрания()
Dim НомерСтроки_Повестка_дня, rowCount, Номер_вопроса, row_Повестка_дня, column_Повестка_дня As Byte

  ' Запрос на формирование бланка протокола Собрания
  If MsgBox("Перенести Повестку и Поручения в Архив?", vbYesNo) = vbYes Then
    
    ' 1. Копирование полей протокола: переносим все в BASE\?
    ' OpenBookInBase InsertRecordInBook CloseBook
    ' 1.1. Копирование в BASE\Protocols
    Application.StatusBar = "Копирование данных: Protocols ..."
    ' Открываем BASE\Protocols
    OpenBookInBase ("Protocols")
    ThisWorkbook.Activate
    
    ' Вставляем протокол
    Call InsertRecordInBook("Protocols", "Лист1", "Protocol", ThisWorkbook.Sheets("ЕСУП").Range("G2").Value, _
                                            "Date", dateProtocol(ThisWorkbook.Sheets("ЕСУП").Range("G2").Value), _
                                              "Protocol", ThisWorkbook.Sheets("ЕСУП").Range("G2").Value, _
                                                "Theme", "Еженедельный конференц-колл с Управляющими офисов и НОРПиКО Тюменского РОО по подведению итогов работы офисов за предыдущую неделю и постановке бизнес-целей на текущую неделю", _
                                                  "Place", "г.Тюмень, ул.Советская 51/1", _
                                                    "Participants", Присутствовавшие_на_Собрании(1), _
                                                      "Lack", Присутствовавшие_на_Собрании(0), _
                                                        "Copy_to", getFromAddrBook("КОП", 1), _
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

    
    ' 1.2. Копирование в BASE\Themes
    Application.StatusBar = "Копирование данных: Themes ..."
    ' Открываем BASE\Themes
    OpenBookInBase ("Themes")
    ThisWorkbook.Activate
        
    НомерСтроки_Повестка_дня = ThisWorkbook.Sheets("ЕСУП").Range(RangeByValue(ThisWorkbook.Name, "ЕСУП", "Повестка_дня", 100, 100)).Row
    rowCount = 2
    Do While ThisWorkbook.Sheets("ЕСУП").Cells(НомерСтроки_Повестка_дня + rowCount, 1).Value <> ""
      
      ' Если у вопроса стоит отметка "1", то значит мы его осветили и переносим в Архив. Если стоит 0, то оставляем для следующего Собрания
      If ThisWorkbook.Sheets("ЕСУП").Cells(НомерСтроки_Повестка_дня + rowCount, 14).Value = "1" Then
        
        ' Вставляем Темы выступлений
        Call InsertRecordInBook("Themes", "Лист1", "Number_Theme", ThisWorkbook.Sheets("ЕСУП").Range("G2").Value + "-" + CStr(ThisWorkbook.Sheets("ЕСУП").Cells(НомерСтроки_Повестка_дня + rowCount, 1).Value), _
                                            "Date", dateProtocol(ThisWorkbook.Sheets("ЕСУП").Range("G2").Value), _
                                              "Protocol", ThisWorkbook.Sheets("ЕСУП").Range("G2").Value, _
                                                "Number_Theme", ThisWorkbook.Sheets("ЕСУП").Range("G2").Value + "-" + CStr(ThisWorkbook.Sheets("ЕСУП").Cells(НомерСтроки_Повестка_дня + rowCount, 1).Value), _
                                                  "Speker", ThisWorkbook.Sheets("ЕСУП").Cells(НомерСтроки_Повестка_дня + rowCount, 2).Value, _
                                                    "Theme", ThisWorkbook.Sheets("ЕСУП").Cells(НомерСтроки_Повестка_дня + rowCount, 3).Value, _
                                                      "HashTag", ThisWorkbook.Sheets("ЕСУП").Cells(НомерСтроки_Повестка_дня + rowCount, 13).Value, _
                                                        "Action", ThisWorkbook.Sheets("ЕСУП").Cells(НомерСтроки_Повестка_дня + rowCount, 14).Value, _
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

      End If ' Если у вопроса стоит 1
      
      ' Следующая строка
      rowCount = rowCount + 1
    Loop
        
    ' 1.3. Копирование в BASE\Tasks
    ' Application.StatusBar = "Копирование данных: Tasks ..."
    ' Открываем BASE\Tasks
    ' OpenBookInBase ("Tasks")
        
    ' Здесь копировать нечего
        
    ' Закрываем 3 таблицы
    CloseBook ("Protocols")
    CloseBook ("Themes")
    ' CloseBook ("Tasks")
                
    ' 2. Новый номер протокола в G2 Протокол:   10-02032020
    ' 3. Новый номер недели в J2 Неделя: 10
    ' 4. Присутств_на_Собрании - поставить всем 0
    ' 5. Пригл_на_Собрание - проставить всем 0
    
    ' 6. Повестка_дня - очистить все строки, которые были перенесены в архив с признаком 1, строки с 0 оставляем без изменений
    Application.StatusBar = "Очистка списка вопросов..."
    
    row_Повестка_дня = ThisWorkbook.Sheets("ЕСУП").Range(RangeByValue(ThisWorkbook.Name, "ЕСУП", "Повестка_дня", 100, 100)).Row
    column_Повестка_дня = ThisWorkbook.Sheets("ЕСУП").Range(RangeByValue(ThisWorkbook.Name, "ЕСУП", "Повестка_дня", 100, 100)).Column
    rowCount = 2
    Do While ThisWorkbook.Sheets("ЕСУП").Cells(row_Повестка_дня + rowCount, column_Повестка_дня).Value <> ""
    
      ' Если отметка = "1"
      If ThisWorkbook.Sheets("ЕСУП").Cells(row_Повестка_дня + rowCount, column_Повестка_дня + 12).Value = "1" Then
        
        ' Номер
        ThisWorkbook.Sheets("ЕСУП").Cells(row_Повестка_дня + rowCount, column_Повестка_дня - 1).Value = ""
        ' Выступающий:
        ThisWorkbook.Sheets("ЕСУП").Cells(row_Повестка_дня + rowCount, column_Повестка_дня).Value = ""
        ' Тема:
        ThisWorkbook.Sheets("ЕСУП").Cells(row_Повестка_дня + rowCount, column_Повестка_дня + 1).Value = ""
        ' HashTag
        ThisWorkbook.Sheets("ЕСУП").Cells(row_Повестка_дня + rowCount, column_Повестка_дня + 11).Value = ""
        ' Отметка
        ThisWorkbook.Sheets("ЕСУП").Cells(row_Повестка_дня + rowCount, column_Повестка_дня + 12).Value = ""
      
      End If
      
      ' Следующая строка
      rowCount = rowCount + 1
    Loop
    
    Application.StatusBar = ""
    
    ' 7. Очистить Поручения_участникам
    Application.StatusBar = "Очистка списка Поручений участникам..."
    
    row_Поручения_участникам = ThisWorkbook.Sheets("ЕСУП").Range(RangeByValue(ThisWorkbook.Name, "ЕСУП", "Поручения_участникам", 100, 100)).Row
    column_Поручения_участникам = ThisWorkbook.Sheets("ЕСУП").Range(RangeByValue(ThisWorkbook.Name, "ЕСУП", "Поручения_участникам", 100, 100)).Column
    
    ' Счетчик ошибок
    CountError = 0

    rowCount = 2
    Do While ThisWorkbook.Sheets("ЕСУП").Cells(row_Поручения_участникам + rowCount, column_Поручения_участникам + 2).Value <> ""
    
      ' Если отметка = "1"
      If ThisWorkbook.Sheets("ЕСУП").Cells(row_Поручения_участникам + rowCount, column_Поручения_участникам + 13).Value = "1" Then
        
        ' Номер
        ThisWorkbook.Sheets("ЕСУП").Cells(row_Поручения_участникам + rowCount, column_Поручения_участникам - 1).Value = ""
        ' Ответственный:
        ThisWorkbook.Sheets("ЕСУП").Cells(row_Поручения_участникам + rowCount, column_Поручения_участникам).Value = ""
        ' Срок:
        ThisWorkbook.Sheets("ЕСУП").Cells(row_Поручения_участникам + rowCount, column_Поручения_участникам + 1).Value = ""
        ' Содержание поручения:
        ThisWorkbook.Sheets("ЕСУП").Cells(row_Поручения_участникам + rowCount, column_Поручения_участникам + 2).Value = ""
        ' Отметка перенесено "В To-Do"
        ThisWorkbook.Sheets("ЕСУП").Cells(row_Поручения_участникам + rowCount, column_Поручения_участникам + 13).Value = ""
      
      End If
      
      ' Если есть отметка = "0"
      If ThisWorkbook.Sheets("ЕСУП").Cells(row_Поручения_участникам + rowCount, column_Поручения_участникам + 13).Value = "0" Then
        ' Счетчик ошибок
        CountError = CountError + 1
      End If
      
      ' Следующая строка
      rowCount = rowCount + 1
    Loop
    
    ' Проверяем результат на наличие ошибок
    If CountError <> 0 Then
      MsgBox ("Внимание! " + CStr(CountError) + " Поручений не перенесены в To-Do!")
    End If
    
    Application.StatusBar = ""
    
    ' Зачеркиваем: "Перенести в Архив Повестку Собрания"
    Call ЗачеркиваемТекстВячейке("Лист0", RangeByValue(ThisWorkbook.Name, "Лист0", "Перенести в Архив Повестку Собрания", 100, 100))
    
    ' Строка статуса
    Application.StatusBar = ""
    
    MsgBox ("Повестка и Поручения перенесены в Архив!")
  
  End If
End Sub

' Отправить Протокол собрания
Sub sendProtocol()
  ' Запрос на формирование бланка протокола Собрания
  If MsgBox("Отправить Протокол Собрания?", vbYesNo) = vbYes Then
    
    ' Отправить в Lotus Notes
    ' Строка статуса
    Application.StatusBar = "Отправка Протокола участникам в Lotus Notes ..."
    
    ' Строка статуса
    Application.StatusBar = ""
    
    MsgBox ("Протокол отправлен в Lotus Notes!")
    ' Скопировать в каталог ЕСУП
    ' Строка статуса
    Application.StatusBar = "Копирование Протокола в каталог ЕСУП ..."
    
    ' Строка статуса
    Application.StatusBar = ""
    MsgBox ("Протокол скопирован в каталог ЕСУП!")
  End If
End Sub


' Добавление строки
Sub InsertRow_InProtocol(In_FileProtocolName, In_Sheets, In_RowNumber)
  
  ' Вставляем пустую строку в блок "Поручения"
  Workbooks(In_FileProtocolName).Sheets(In_Sheets).Range(CStr(In_RowNumber) + ":" + CStr(In_RowNumber)).Select
  Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
              
End Sub

' ЕСУП: Нумеровать список вопросов на Листе в разделе повестка от 1 до ...
Sub createNumberingThemes()
Dim НомерСтроки_Повестка_дня, rowCount As Byte

  НомерСтроки_Повестка_дня = ThisWorkbook.Sheets("ЕСУП").Range(RangeByValue(ThisWorkbook.Name, "ЕСУП", "Повестка_дня", 100, 100)).Row
  rowCount = 2
  Do While ThisWorkbook.Sheets("ЕСУП").Cells(НомерСтроки_Повестка_дня + rowCount, 3).Value <> ""
    
    ' Нумерация
    ThisWorkbook.Sheets("ЕСУП").Cells(НомерСтроки_Повестка_дня + rowCount, 1).Value = (rowCount - 1)
    
    ' Снимаем установку "Переносить по словам"
    ' ThisWorkbook.Sheets("ЕСУП").Cells(НомерСтроки_Повестка_дня + rowCount, 3).HorizontalAlignment = xlGeneral
    ' ThisWorkbook.Sheets("ЕСУП").Cells(НомерСтроки_Повестка_дня + rowCount, 3).VerticalAlignment = xlBottom
    ThisWorkbook.Sheets("ЕСУП").Cells(НомерСтроки_Повестка_дня + rowCount, 3).WrapText = False
    ' ThisWorkbook.Sheets("ЕСУП").Cells(НомерСтроки_Повестка_дня + rowCount, 3).Orientation = 0
    ' ThisWorkbook.Sheets("ЕСУП").Cells(НомерСтроки_Повестка_дня + rowCount, 3).AddIndent = False
    ' ThisWorkbook.Sheets("ЕСУП").Cells(НомерСтроки_Повестка_дня + rowCount, 3).IndentLevel = 0
    ' ThisWorkbook.Sheets("ЕСУП").Cells(НомерСтроки_Повестка_дня + rowCount, 3).ShrinkToFit = False
    ' ThisWorkbook.Sheets("ЕСУП").Cells(НомерСтроки_Повестка_дня + rowCount, 3).ReadingOrder = xlContext
    ' ThisWorkbook.Sheets("ЕСУП").Cells(НомерСтроки_Повестка_дня + rowCount, 3).MergeCells = False
    
    ' Следующая строка
    rowCount = rowCount + 1
  Loop

End Sub

' ЕСУП: Нумеровать список вопросов на Листе в разделе Поручения от 1 до ...
Sub createNumberingTask()
Dim НомерСтроки_Поручения_участникам, rowCount As Byte

  НомерСтроки_Поручения_участникам = ThisWorkbook.Sheets("ЕСУП").Range(RangeByValue(ThisWorkbook.Name, "ЕСУП", "Поручения_участникам", 100, 100)).Row
  rowCount = 2
  Do While ThisWorkbook.Sheets("ЕСУП").Cells(НомерСтроки_Поручения_участникам + rowCount, 19).Value <> ""
    ThisWorkbook.Sheets("ЕСУП").Cells(НомерСтроки_Поручения_участникам + rowCount, 18).Value = (rowCount - 1)
    ' Следующая строка
    rowCount = rowCount + 1
  Loop

End Sub


' Перевести в ЕСУП номер следующей недели в CSMP (common standards for management procedures)
Sub changeWeekNumberInCSMP()
Dim currentOperDate As Date
Dim i, Range_Row, Range_Row2, Range_Column, Range_Column2, row_Присутств_на_Собрании, column_Присутств_на_Собрании As Byte
Dim ПланНаНеделю_Активы, ПланНаНеделю_ИК, ПланНаНеделю_ИК_Сургут As Double
Dim rowCount, ПланНаНеделю_ДК, ПланНаНеделю_КК As Integer
Dim Номер_протокола_прошлой_недели, Range_str, Range_str2 As String

  ' Текущая операционная дата на новой неделе - первый понедельник на неделя + 1
  ' currentOperDate = Date
  currentOperDate = MondayDateByWeekNumber(НеделяНаЛистеN("ЕСУП") + 1, Year(Date))
  
  ' Проверяем неделю на листе и текущую неделю
  If НеделяНаЛистеN("ЕСУП") < WeekNumber(currentOperDate) Then
  
    ' Запрос на формирование бланка протокола Собрания
    If MsgBox("Открыть новую неделю?", vbYesNo) = vbYes Then
    
      ' Строка статуса
      Application.StatusBar = "Открытие новой недели ..."
    
      ' Протокол прошлой недели
      Номер_протокола_прошлой_недели = ThisWorkbook.Sheets("ЕСУП").Cells(rowByValue(ThisWorkbook.Name, "ЕСУП", "Протокол:", 100, 100), ColumnByValue(ThisWorkbook.Name, "ЕСУП", "Протокол:", 100, 100) + 1).Value
      
      ' Открываем BASE\Tasks
      OpenBookInBase ("Tasks")

      ' Период: с 01.01 по 07.01.2020 г. (ищем "Единая Система Управленческих Процедур")
      ThisWorkbook.Sheets("ЕСУП").Cells(rowByValue(ThisWorkbook.Name, "ЕСУП", "Единая Система Управленческих Процедур", 100, 100) + 1, ColumnByValue(ThisWorkbook.Name, "ЕСУП", "Единая Система Управленческих Процедур", 100, 100)).Value = "Период: с " + CStr(weekStartDate(currentOperDate)) + " по " + CStr(weekEndDate(currentOperDate)) + " г."
    
      ' в G2 Протокол:   10-02032020
      ThisWorkbook.Sheets("ЕСУП").Cells(rowByValue(ThisWorkbook.Name, "ЕСУП", "Протокол:", 100, 100), ColumnByValue(ThisWorkbook.Name, "ЕСУП", "Протокол:", 100, 100) + 1).Value = CStr(WeekNumber(currentOperDate)) + "-" + strDDMMYYYY(currentOperDate)

      ' в J2 Неделя: 10. Устанавливаем номер недели как предидущая плюс 1!
      ThisWorkbook.Sheets("ЕСУП").Cells(rowByValue(ThisWorkbook.Name, "ЕСУП", "Неделя:", 100, 100), ColumnByValue(ThisWorkbook.Name, "ЕСУП", "Неделя:", 100, 100) + 1).Value = WeekNumber(currentOperDate)
      
      ' в ячейку C6 (ПК), E6 (ИСЖ), G6 (ДК), I6 (КК) занести номер недели "План нед.(11)"
      ' ПК
      ThisWorkbook.Sheets("ЕСУП").Cells(rowByValue(ThisWorkbook.Name, "ЕСУП", "Поручения периода:", 100, 100) + 2, ColumnByValue(ThisWorkbook.Name, "ЕСУП", "Поручения периода:", 100, 100) + 1).Value = "План нед.(" + CStr(WeekNumber(currentOperDate)) + ")"
      ' по ИСЖ переносим на Лист 4
      ' ThisWorkbook.Sheets("ЕСУП").Cells(RowByValue(ThisWorkbook.Name, "ЕСУП", "Поручения периода:", 100, 100) + 2, ColumnByValue(ThisWorkbook.Name, "ЕСУП", "Поручения периода:", 100, 100) + 3).Value = "План нед.(" + CStr(WeekNumber(currentOperDate)) + ")"
      ' ДК
      ThisWorkbook.Sheets("ЕСУП").Cells(rowByValue(ThisWorkbook.Name, "ЕСУП", "Поручения периода:", 100, 100) + 2, ColumnByValue(ThisWorkbook.Name, "ЕСУП", "Поручения периода:", 100, 100) + 5).Value = "План нед.(" + CStr(WeekNumber(currentOperDate)) + ")"
      ' КК
      ThisWorkbook.Sheets("ЕСУП").Cells(rowByValue(ThisWorkbook.Name, "ЕСУП", "Поручения периода:", 100, 100) + 2, ColumnByValue(ThisWorkbook.Name, "ЕСУП", "Поручения периода:", 100, 100) + 7).Value = "План нед.(" + CStr(WeekNumber(currentOperDate)) + ")"
      ' ИЦ
      ThisWorkbook.Sheets("ЕСУП").Cells(rowByValue(ThisWorkbook.Name, "ЕСУП", "Поручения периода:", 100, 100) + 13, ColumnByValue(ThisWorkbook.Name, "ЕСУП", "Поручения периода:", 100, 100) + 1).Value = "План нед.(" + CStr(WeekNumber(currentOperDate)) + ")"
      
    
      ' Делаем расчет планов на неделю по активам - берем план и факт с Лист3
      
      ' Активы: План-Поручение офису ВПКi
      ' Range_str = RangeByValue(ThisWorkbook.Name, "Лист3", "Форма 2", 100, 100)
      ' Форма 2.1
      Range_str = RangeByValue(ThisWorkbook.Name, "Лист3", "Форма 2.1", 100, 100)
      Range_Row = Workbooks(ThisWorkbook.Name).Sheets("Лист3").Range(Range_str).Row
      Range_Column = Workbooks(ThisWorkbook.Name).Sheets("Лист3").Range(Range_str).Column
      
      ' ДК и КК
      Range_str2 = RangeByValue(ThisWorkbook.Name, "Лист5", "Заявки на карточные продукты", 100, 100)
      Range_Row2 = Workbooks(ThisWorkbook.Name).Sheets("Лист5").Range(Range_str2).Row
      Range_Column2 = Workbooks(ThisWorkbook.Name).Sheets("Лист5").Range(Range_str2).Column
      
      
      ' Цикл по 5-ти офисам
      For i = 1 To 6
        
        ' Офисы - номера от 1 до 5
        If i <= 5 Then
        
          Application.StatusBar = "Расчет планов " + getNameOfficeByNumber(i) + "..."
        
          ' Расчет плана ПК
          ПланНаНеделю_Активы = Round(ПланНаНеделю(ThisWorkbook.Sheets("Лист3").Cells(Range_Row + 2 + i, Range_Column + 3).Value, ThisWorkbook.Sheets("Лист3").Cells(Range_Row + 2 + i, Range_Column + 4).Value, currentOperDate, 6), 0)
          ' Запись плана ПК
          Call setК_порInЕСУП(ThisWorkbook.Name, "ЕСУП", "ВПК" + CStr(i), ПланНаНеделю_Активы, weekStartDate(currentOperDate), "тыс.руб.", "")
          ' Установка исполнения плана ПК прошлой недели Номер_протокола_прошлой_недели
          Call setStatusInTasks("Tasks", "Лист1", currentOperDate, "ВПК" + CStr(i), Номер_протокола_прошлой_недели)
      
          ' Расчет плана ДК
          ПланНаНеделю_ДК = Round(ПланНаНеделю(ThisWorkbook.Sheets("Лист5").Cells(Range_Row2 + 2 + i, Range_Column2 + 1).Value, ThisWorkbook.Sheets("Лист5").Cells(Range_Row2 + 2 + i, Range_Column2 + 2).Value, currentOperDate, 6), 0)
          ' Запись плана ДК
          Call setК_порInЕСУП(ThisWorkbook.Name, "ЕСУП", "ЗДК" + CStr(i), ПланНаНеделю_ДК, weekStartDate(currentOperDate), "шт.", "")
          ' Установка исполнения плана ДК прошлой недели Номер_протокола_прошлой_недели
          Call setStatusInTasks("Tasks", "Лист1", currentOperDate, "ЗДК" + CStr(i), Номер_протокола_прошлой_недели)
       
          ' Расчет плана КК
          ПланНаНеделю_КК = Round(ПланНаНеделю(ThisWorkbook.Sheets("Лист5").Cells(Range_Row2 + 2 + i, Range_Column2 + 3).Value, ThisWorkbook.Sheets("Лист5").Cells(Range_Row2 + 2 + i, Range_Column2 + 4).Value, currentOperDate, 6), 0)
          ' Запись плана КК
          Call setК_порInЕСУП(ThisWorkbook.Name, "ЕСУП", "ЗКК" + CStr(i), ПланНаНеделю_КК, weekStartDate(currentOperDate), "шт.", "")
          ' Установка исполнения плана КК прошлой недели Номер_протокола_прошлой_недели
          Call setStatusInTasks("Tasks", "Лист1", currentOperDate, "ЗКК" + CStr(i), Номер_протокола_прошлой_недели)
                    
        Else
          ' Ипотека ВИК1
          Application.StatusBar = "Расчет планов ИЦ ОО «Тюменский» ..."
          ' Расчет плана ИЦ
          ПланНаНеделю_ИК = Round(ПланНаНеделю(ThisWorkbook.Sheets("Лист3").Cells(Range_Row + 16, Range_Column + 3).Value, ThisWorkbook.Sheets("Лист3").Cells(Range_Row + 16, Range_Column + 4).Value, currentOperDate, 5), 0)
          ' Запись плана ИЦ
          Call setК_порInЕСУП(ThisWorkbook.Name, "ЕСУП", "ВИК1", ПланНаНеделю_ИК, weekStartDate(currentOperDate), "тыс.руб.", "")
          ' Установка исполнения плана ИЦ прошлой недели Номер_протокола_прошлой_недели
          Call setStatusInTasks("Tasks", "Лист1", currentOperDate, "ВИК1", Номер_протокола_прошлой_недели)
          
          ' и в т.ч. по ОО «Сургутский» ВИК2
          ' Расчет плана ИЦ Сургут
          ПланНаНеделю_ИК_Сургут = Round(ПланНаНеделю(ThisWorkbook.Sheets("Лист3").Cells(Range_Row + 17, Range_Column + 3).Value, ThisWorkbook.Sheets("Лист3").Cells(Range_Row + 17, Range_Column + 4).Value, currentOperDate, 5), 0)
          ' Запись плана ИЦ Сургут
          Call setК_порInЕСУП(ThisWorkbook.Name, "ЕСУП", "ВИК2", ПланНаНеделю_ИК_Сургут, weekStartDate(currentOperDate), "тыс.руб.", "")
          ' Установка исполнения плана ИЦ Сургут прошлой недели Номер_протокола_прошлой_недели
          Call setStatusInTasks("Tasks", "Лист1", currentOperDate, "ВИК2", Номер_протокола_прошлой_недели)
          
        End If ' Офисы - номера от 1 до 5
      
      Next i
  
      Application.StatusBar = "Очистка списка Присутств_на_Собрании ..." ' Application.StatusBar = "Очистка списка Пригл_на_Собрание ..."
      row_Присутств_на_Собрании = ThisWorkbook.Sheets("ЕСУП").Range(RangeByValue(ThisWorkbook.Name, "ЕСУП", "Присутств_на_Собрании", 100, 100)).Row
      column_Присутств_на_Собрании = ThisWorkbook.Sheets("ЕСУП").Range(RangeByValue(ThisWorkbook.Name, "ЕСУП", "Присутств_на_Собрании", 100, 100)).Column
      rowCount = 1
      Do While ThisWorkbook.Sheets("ЕСУП").Cells(row_Присутств_на_Собрании + rowCount, column_Присутств_на_Собрании).Value <> ""
        If ThisWorkbook.Sheets("ЕСУП").Cells(row_Присутств_на_Собрании + rowCount, column_Присутств_на_Собрании).Value = "1" Then
          ThisWorkbook.Sheets("ЕСУП").Cells(row_Присутств_на_Собрании + rowCount, column_Присутств_на_Собрании).Value = "0"
        End If
        ' Следующая строка
        rowCount = rowCount + 1
      Loop

      ' Строка статуса
      Application.StatusBar = ""
                    
      ' Закрываем базу BASE\Tasks
      CloseBook ("Tasks")
    
      ' Зачеркнуть на Лист0 "Открыть новую неделю в ЕСУП"
      Call ЗачеркиваемТекстВячейке("Лист0", RangeByValue(ThisWorkbook.Name, "Лист0", "Открыть новую неделю в ЕСУП", 100, 100))

      MsgBox ("Неделя открыта!")
    End If
  Else
    MsgBox ("Уже установлена неделя: " + CStr(WeekNumber(currentOperDate)) + "!")
  End If
  
End Sub

' Перемещение пункта в повестке собрания вверх
Sub moveInListUp()
Dim Ячейка_Повестка_дня, Текущий_номер, Текущий_выступающий, Текущий_тема, Текущий_HashTag, Текущий_Отметка, Цель_номер, Цель_выступающий, Цель_тема, Цель_HashTag, Цель_Отметка As String
Dim НомерСтроки_Повестка_дня, НомерСтолбца_Повестка_дня, Текущий_Row, Текущий_Column As Byte

  ' Определяем, где находится текущая ячейка. Должен быть диапазон A62:N90 (в относительных от "Повестка_дня" координатах)
  Ячейка_Повестка_дня = RangeByValue(ThisWorkbook.Name, "ЕСУП", "Повестка_дня", 100, 100)
  НомерСтроки_Повестка_дня = ThisWorkbook.Sheets("ЕСУП").Range(Ячейка_Повестка_дня).Row
  НомерСтолбца_Повестка_дня = ThisWorkbook.Sheets("ЕСУП").Range(Ячейка_Повестка_дня).Column
  '
  If (ActiveCell.Row >= НомерСтроки_Повестка_дня + 2) And (ActiveCell.Row <= НомерСтроки_Повестка_дня + 20) And (ActiveCell.Column >= НомерСтолбца_Повестка_дня - 1) And ((ActiveCell.Column <= НомерСтолбца_Повестка_дня + 12)) Then
      ' Координаты
      Текущий_Row = ActiveCell.Row
      Текущий_Column = ActiveCell.Column
      ' Запоминаем текущий
      Текущий_номер = ThisWorkbook.Sheets("ЕСУП").Cells(Текущий_Row, НомерСтолбца_Повестка_дня - 1).Value
      Текущий_выступающий = ThisWorkbook.Sheets("ЕСУП").Cells(Текущий_Row, НомерСтолбца_Повестка_дня).Value
      Текущий_тема = ThisWorkbook.Sheets("ЕСУП").Cells(Текущий_Row, НомерСтолбца_Повестка_дня + 1).Value
      Текущий_HashTag = ThisWorkbook.Sheets("ЕСУП").Cells(Текущий_Row, НомерСтолбца_Повестка_дня + 11).Value
      Текущий_Отметка = ThisWorkbook.Sheets("ЕСУП").Cells(Текущий_Row, НомерСтолбца_Повестка_дня + 12).Value
      ' Запоминаем цель
      ' Цель_Row = Текущий_Row + 1
      Цель_номер = ThisWorkbook.Sheets("ЕСУП").Cells(Текущий_Row - 1, НомерСтолбца_Повестка_дня - 1).Value
      Цель_выступающий = ThisWorkbook.Sheets("ЕСУП").Cells(Текущий_Row - 1, НомерСтолбца_Повестка_дня).Value
      Цель_тема = ThisWorkbook.Sheets("ЕСУП").Cells(Текущий_Row - 1, НомерСтолбца_Повестка_дня + 1).Value
      Цель_HashTag = ThisWorkbook.Sheets("ЕСУП").Cells(Текущий_Row - 1, НомерСтолбца_Повестка_дня + 11).Value
      Цель_Отметка = ThisWorkbook.Sheets("ЕСУП").Cells(Текущий_Row - 1, НомерСтолбца_Повестка_дня + 12).Value
      ' Меняем местами:
      ' Текущий ставим в Цель:
      ThisWorkbook.Sheets("ЕСУП").Cells(Текущий_Row - 1, НомерСтолбца_Повестка_дня - 1).Value = Текущий_номер
      ThisWorkbook.Sheets("ЕСУП").Cells(Текущий_Row - 1, НомерСтолбца_Повестка_дня).Value = Текущий_выступающий
      ThisWorkbook.Sheets("ЕСУП").Cells(Текущий_Row - 1, НомерСтолбца_Повестка_дня + 1).Value = Текущий_тема
      ThisWorkbook.Sheets("ЕСУП").Cells(Текущий_Row - 1, НомерСтолбца_Повестка_дня + 11).Value = Текущий_HashTag
      ThisWorkbook.Sheets("ЕСУП").Cells(Текущий_Row - 1, НомерСтолбца_Повестка_дня + 12).Value = Текущий_Отметка
      ' Цель ставим в Текущий:
      ThisWorkbook.Sheets("ЕСУП").Cells(Текущий_Row, НомерСтолбца_Повестка_дня - 1).Value = Цель_номер
      ThisWorkbook.Sheets("ЕСУП").Cells(Текущий_Row, НомерСтолбца_Повестка_дня).Value = Цель_выступающий
      ThisWorkbook.Sheets("ЕСУП").Cells(Текущий_Row, НомерСтолбца_Повестка_дня + 1).Value = Цель_тема
      ThisWorkbook.Sheets("ЕСУП").Cells(Текущий_Row, НомерСтолбца_Повестка_дня + 11).Value = Цель_HashTag
      ThisWorkbook.Sheets("ЕСУП").Cells(Текущий_Row, НомерСтолбца_Повестка_дня + 12).Value = Цель_Отметка
      ' Устанавливаем на строку выше
      ThisWorkbook.Sheets("ЕСУП").Cells(Текущий_Row - 1, Текущий_Column).Select
      ' Производим перенумерацию списка
      Call createNumberingThemes
    Else
      MsgBox ("Укажите ячейку в диапазоне Повестка_дня!")
  End If
End Sub

' Перемещение пункта в повестке собрания вниз
Sub moveInListDown()
Dim Ячейка_Повестка_дня, Текущий_номер, Текущий_выступающий, Текущий_тема, Текущий_HashTag, Текущий_Отметка, Цель_номер, Цель_выступающий, Цель_тема, Цель_HashTag, Цель_Отметка As String
Dim НомерСтроки_Повестка_дня, НомерСтолбца_Повестка_дня, Текущий_Row, Текущий_Column As Byte

  ' Определяем, где находится текущая ячейка. Должен быть диапазон A62:N90 (в относительных от "Повестка_дня" координатах)
  Ячейка_Повестка_дня = RangeByValue(ThisWorkbook.Name, "ЕСУП", "Повестка_дня", 100, 100)
  НомерСтроки_Повестка_дня = ThisWorkbook.Sheets("ЕСУП").Range(Ячейка_Повестка_дня).Row
  НомерСтолбца_Повестка_дня = ThisWorkbook.Sheets("ЕСУП").Range(Ячейка_Повестка_дня).Column
  '
  If (ActiveCell.Row >= НомерСтроки_Повестка_дня + 2) And (ActiveCell.Row <= НомерСтроки_Повестка_дня + 20) And (ActiveCell.Column >= НомерСтолбца_Повестка_дня - 1) And ((ActiveCell.Column <= НомерСтолбца_Повестка_дня + 12)) Then
      ' Координаты
      Текущий_Row = ActiveCell.Row
      Текущий_Column = ActiveCell.Column
      ' Запоминаем текущий
      Текущий_номер = ThisWorkbook.Sheets("ЕСУП").Cells(Текущий_Row, НомерСтолбца_Повестка_дня - 1).Value
      Текущий_выступающий = ThisWorkbook.Sheets("ЕСУП").Cells(Текущий_Row, НомерСтолбца_Повестка_дня).Value
      Текущий_тема = ThisWorkbook.Sheets("ЕСУП").Cells(Текущий_Row, НомерСтолбца_Повестка_дня + 1).Value
      Текущий_HashTag = ThisWorkbook.Sheets("ЕСУП").Cells(Текущий_Row, НомерСтолбца_Повестка_дня + 11).Value
      Текущий_Отметка = ThisWorkbook.Sheets("ЕСУП").Cells(Текущий_Row, НомерСтолбца_Повестка_дня + 12).Value
      ' Запоминаем цель
      ' Цель_Row = Текущий_Row + 1
      Цель_номер = ThisWorkbook.Sheets("ЕСУП").Cells(Текущий_Row + 1, НомерСтолбца_Повестка_дня - 1).Value
      Цель_выступающий = ThisWorkbook.Sheets("ЕСУП").Cells(Текущий_Row + 1, НомерСтолбца_Повестка_дня).Value
      Цель_тема = ThisWorkbook.Sheets("ЕСУП").Cells(Текущий_Row + 1, НомерСтолбца_Повестка_дня + 1).Value
      Цель_HashTag = ThisWorkbook.Sheets("ЕСУП").Cells(Текущий_Row + 1, НомерСтолбца_Повестка_дня + 11).Value
      Цель_Отметка = ThisWorkbook.Sheets("ЕСУП").Cells(Текущий_Row + 1, НомерСтолбца_Повестка_дня + 12).Value
      ' Меняем местами:
      ' Текущий ставим в Цель:
      ThisWorkbook.Sheets("ЕСУП").Cells(Текущий_Row + 1, НомерСтолбца_Повестка_дня - 1).Value = Текущий_номер
      ThisWorkbook.Sheets("ЕСУП").Cells(Текущий_Row + 1, НомерСтолбца_Повестка_дня).Value = Текущий_выступающий
      ThisWorkbook.Sheets("ЕСУП").Cells(Текущий_Row + 1, НомерСтолбца_Повестка_дня + 1).Value = Текущий_тема
      ThisWorkbook.Sheets("ЕСУП").Cells(Текущий_Row + 1, НомерСтолбца_Повестка_дня + 11).Value = Текущий_HashTag
      ThisWorkbook.Sheets("ЕСУП").Cells(Текущий_Row + 1, НомерСтолбца_Повестка_дня + 12).Value = Текущий_Отметка
      ' Цель ставим в Текущий:
      ThisWorkbook.Sheets("ЕСУП").Cells(Текущий_Row, НомерСтолбца_Повестка_дня - 1).Value = Цель_номер
      ThisWorkbook.Sheets("ЕСУП").Cells(Текущий_Row, НомерСтолбца_Повестка_дня).Value = Цель_выступающий
      ThisWorkbook.Sheets("ЕСУП").Cells(Текущий_Row, НомерСтолбца_Повестка_дня + 1).Value = Цель_тема
      ThisWorkbook.Sheets("ЕСУП").Cells(Текущий_Row, НомерСтолбца_Повестка_дня + 11).Value = Цель_HashTag
      ThisWorkbook.Sheets("ЕСУП").Cells(Текущий_Row, НомерСтолбца_Повестка_дня + 12).Value = Цель_Отметка
      ' Устанавливаем на строку выше
      ThisWorkbook.Sheets("ЕСУП").Cells(Текущий_Row + 1, Текущий_Column).Select
      ' Производим перенумерацию списка
      Call createNumberingThemes
    Else
      MsgBox ("Укажите ячейку в диапазоне Повестка_дня!")
  End If

End Sub

' Удаление пункта из повестки
Sub deleteFromList()
Dim Ячейка_Повестка_дня As String
Dim НомерСтроки_Повестка_дня, НомерСтолбца_Повестка_дня As Byte

  ' Определяем, где находится текущая ячейка. Должен быть диапазон A62:N90 (в относительных от "Повестка_дня" координатах)
  Ячейка_Повестка_дня = RangeByValue(ThisWorkbook.Name, "ЕСУП", "Повестка_дня", 100, 100)
  НомерСтроки_Повестка_дня = ThisWorkbook.Sheets("ЕСУП").Range(Ячейка_Повестка_дня).Row
  НомерСтолбца_Повестка_дня = ThisWorkbook.Sheets("ЕСУП").Range(Ячейка_Повестка_дня).Column
  '
  If (ActiveCell.Row >= НомерСтроки_Повестка_дня + 2) And (ActiveCell.Row <= НомерСтроки_Повестка_дня + 20) And (ActiveCell.Column >= НомерСтолбца_Повестка_дня - 1) And ((ActiveCell.Column <= НомерСтолбца_Повестка_дня + 12)) And (ThisWorkbook.Sheets("ЕСУП").Cells(ActiveCell.Row, НомерСтолбца_Повестка_дня + 1).Value <> "") Then
    '
    If MsgBox("Удалить вопрос №" + CStr(ThisWorkbook.Sheets("ЕСУП").Cells(ActiveCell.Row, НомерСтолбца_Повестка_дня - 1).Value) + " из повестки?", vbYesNo) = vbYes Then
      ' Удаляем
      ThisWorkbook.Sheets("ЕСУП").Cells(ActiveCell.Row, НомерСтолбца_Повестка_дня - 1).Value = ""
      ThisWorkbook.Sheets("ЕСУП").Cells(ActiveCell.Row, НомерСтолбца_Повестка_дня).Value = ""
      ThisWorkbook.Sheets("ЕСУП").Cells(ActiveCell.Row, НомерСтолбца_Повестка_дня + 1).Value = ""
      ThisWorkbook.Sheets("ЕСУП").Cells(ActiveCell.Row, НомерСтолбца_Повестка_дня + 11).Value = ""
      ThisWorkbook.Sheets("ЕСУП").Cells(ActiveCell.Row, НомерСтолбца_Повестка_дня + 12).Value = ""
    End If
  Else
    MsgBox ("Укажите ячейку в диапазоне Повестка_дня!")
  End If
  
End Sub


' Открыть протокол Собрания
Sub Открыть_Протокол_Собрания()
  
  ' Открываем сформированный протокол по имени
           
  ' Имя файла с протоколом - берем из G2 "10-02032020"
  FileProtocolName = "Протокол _ РОО Тюменский_" + CStr(dateProtocol(ThisWorkbook.Sheets("ЕСУП").Range("G2").Value)) + ".xlsx"

  ' Проверка наличия файла
  If Dir(ThisWorkbook.Path + "\Out\" + FileProtocolName) <> "" Then
    ' Открываем шаблон Протокола из C:\Users\...\Documents\#VBA\DB_Result\Templates
    Workbooks.Open (ThisWorkbook.Path + "\Out\" + FileProtocolName)
  Else
    ' Сообщение, что файл не найден
    MsgBox ("Файл " + FileProtocolName + " не найден!")
  End If
  
End Sub

' Отправка письма: отправляю шаблон самому себе для последующей отправки в сеть письма на его основе:
Sub Отправка_Lotus_Notes_Лист_ЕСУП()
Dim темаПисьма, текстПисьма, hashTag, attachmentFile As String
Dim i As Byte
Dim Отправка_Lotus_Notes, FileCopy_To_ЕСУП As Boolean
  
  
  Отправка_Lotus_Notes = False
  FileCopy_To_ЕСУП = False
  
  If MsgBox("Отправить себе Шаблон письма?", vbYesNo) = vbYes Then
    
    ' Строка статуса
    Application.StatusBar = "Отправка письма в Lotus Notes ..."
    
    ' Тема письма - Тема:
    ' темаПисьма = "Тюменский РОО: Протокол конференц-колла с офисами от " + CStr(dateProtocol(ThisWorkbook.Sheets("ЕСУП").Range("G2").Value)) + " г."
    темаПисьма = ThisWorkbook.Sheets("ЕСУП").Cells(1, 17).Value
    
    ' hashTag - Хэштэг:
    hashTag = hashTagFromSheet("ЕСУП")
    
    ' Файл-вложение (!!!)
    attachmentFile = ThisWorkbook.Path + "\Out\Протокол _ РОО Тюменский_" + CStr(dateProtocol(ThisWorkbook.Sheets("ЕСУП").Range("G2").Value)) + ".xlsx"
    
    ' Текст письма
    текстПисьма = "" + Chr(13)
    текстПисьма = текстПисьма + "" + ThisWorkbook.Sheets("ЕСУП").Cells(rowByValue(ThisWorkbook.Name, "ЕСУП", "Список получателей:", 100, 100), ColumnByValue(ThisWorkbook.Name, "ЕСУП", "Список получателей:", 100, 100) + 2).Value + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    ' Адреса для копии
    
    '     str_копияВадрес = getFromAddrBook("РД", 1) + ", " + getFromAddrBook("КОП", 1) + ", " + getFromAddrBook("КИп", 1) + ", " + getFromAddrBook("ККП", 1) + ", " + getFromAddrBook("Кaf", 1)
    текстПисьма = текстПисьма + "" + getFromAddrBook("РД", 2) + ", " + getFromAddrBook("КОП", 2) + ", " + getFromAddrBook("КИп", 2) + ", " + getFromAddrBook("ККП", 2) + ", " + getFromAddrBook("Кaf", 2) + ", " + getFromAddrBook("КПВО", 1) + Chr(13)
    
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "Уважаемые сотрудники," + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "Направляю протокол конференции. Прошу принять поручения к исполнению." + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    ' Визитка (подпись С Ув., )
    текстПисьма = текстПисьма + ПодписьВПисьме()
    ' Хэштег
    текстПисьма = текстПисьма + createBlankStr(25) + hashTag + " " + hashTagFromSheetII("ЕСУП", 2)
    
    ' Вызов
    Call send_Lotus_Notes2(темаПисьма, "Sergey Fedorovich Proschaev/Tyumen/PSBank/Ru", "", "", текстПисьма, attachmentFile)
        
    Отправка_Lotus_Notes = True
        
    ' Строка статуса
    Application.StatusBar = ""
    
    ' Сообщение
    MsgBox ("Письмо отправлено!")
     
  End If
  
  ' Копирование поручений в To-Do
  Call copyTaskInToDo
  
  ' Создание задачи в ToDO с контролем исполнения в пятницу
  Application.StatusBar = "Создание задачи контроля в To-Do..."
  
  ' Открываем таблицу BASE\ToDo
  OpenBookInBase ("ToDo")

  ' Переходим на окно DB
  ThisWorkbook.Sheets("ЕСУП").Activate

  ' Вносим данные в BASE\ToDo
  hashTag = createHashTag("t")
  ' Id_Task
  Id_TaskVar = Replace(hashTag, "#t", "")

  Call InsertRecordInBook("ToDo", "Лист1", "Id_Task", Id_TaskVar, _
                                            "Date_Create", Date, _
                                              "Id_Task", Id_TaskVar, _
                                                "Task", "Запросить исполнение поручений по протоколу " + ThisWorkbook.Sheets("ЕСУП").Range("G2").Value, _
                                                  "Lotus_subject", subjectFromSheet("ЕСУП"), _
                                                    "Responsible", "УДО/НОРПиКО/НОКП", _
                                                      "Lotus_hashtag", hashTagFromSheetII("ЕСУП", 2), _
                                                        "Task_Status", 1, _
                                                          "Date_Control", weekEndDate(Date) - 2, _
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

  ' Закрываем таблицу BASE\ToDo
  CloseBook ("ToDo")
  
  ' Строка статуса
  Application.StatusBar = ""
  
  
  ' Перенести файл протокола в каталог ЕСУП? - https://www.excel-vba.ru/chto-umeet-excel/kak-sredstvami-vba-pereimenovatperemestitskopirovat-fajl/
  If MsgBox("Перенести файл Протокола Собрания в каталог ЕСУП?", vbYesNo) = vbYes Then
  
    ' Строка статуса
    Application.StatusBar = "Копирование в каталог ЕСУП ..."
    
    FileCopy attachmentFile, "\\probank\DavWWWRoot\drp\DocLib1\Тюменский ОО1\Управленческие процедуры\Собрания\Протокол _ РОО Тюменский_" + CStr(dateProtocol(ThisWorkbook.Sheets("ЕСУП").Range("G2").Value)) + ".xlsx"
  
    FileCopy_To_ЕСУП = True
  
    ' Строка статуса
    Application.StatusBar = ""

    ' Сообщение
    MsgBox ("Файл перенесен в каталог ЕСУП!")
    
  End If
  
  ' Если оба пункта выполнены - зачеркиваем: "Отправить Протокол Собрания в почте и в каталог ЕСУП"
  If (Отправка_Lotus_Notes = True) And (FileCopy_To_ЕСУП = True) Then
    Call ЗачеркиваемТекстВячейке("Лист0", RangeByValue(ThisWorkbook.Name, "Лист0", "Отправить Протокол Собрания в почте и в каталог ЕСУП", 100, 100))
  End If
  
End Sub

' Отчетность ЕСУП_итог ДД.ММ.ГГГГ
Sub Отчетность_ЕСУП_итог()

' Описание переменных
Dim ячейка_ДО_Лист_ДО_Рег_Сеть_Range_str, ячейка_Форма_Отчетность_ЕСУП_итог_str, ячейка_месяц_str, ячейка_Тюменский_ОО1_str, ReportName_String, officeNameInReport, CheckFormatReportResult, Тема2_Range_str, Хэштэг2_Range_str, поручениеЕСУПдо85процентов As String
Dim i, rowCount As Integer
Dim finishProcess As Boolean
Dim ячейка_ДО_Лист_ДО_Рег_Сеть_Range_Row, ячейка_ДО_Лист_ДО_Рег_Сеть_Range_Column, ячейка_Форма_Отчетность_ЕСУП_итог_row, ячейка_Форма_Отчетность_ЕСУП_итог_Column, ячейка_Тюменский_ОО1_Row, ячейка_Тюменский_ОО1_Column, Row_Включить_в_Собрание, Column_Включить_в_Собрание, Тема2_Range_Row, Тема2_Range_Column, Хэштэг2_Range_Row, Хэштэг2_Range_Column As Byte
Dim dateReportFrom_ReportName_String As Date
    
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
    ThisWorkbook.Sheets("ЕСУП").Activate

    ' Проверка формы отчета
    CheckFormatReportResult = CheckFormatReport(ReportName_String, "инструкция", 9, Date)
    If CheckFormatReportResult = "OK" Then
      
      ' Дата отчета из имени файла
      dateReportFrom_ReportName_String = CDate(getDateReportFromFileName(ReportName_String))
                
      ' Находим ячейку "Форма_Отчетность_ЕСУП_итог" на Листе "ЕСУП"
      ячейка_Форма_Отчетность_ЕСУП_итог_str = RangeByValue(ThisWorkbook.Name, "ЕСУП", "Форма «Отчетность ЕСУП_итог»", 100, 100)
      ячейка_Форма_Отчетность_ЕСУП_итог_row = Workbooks(ThisWorkbook.Name).Sheets("ЕСУП").Range(ячейка_Форма_Отчетность_ЕСУП_итог_str).Row
      ячейка_Форма_Отчетность_ЕСУП_итог_Column = Workbooks(ThisWorkbook.Name).Sheets("ЕСУП").Range(ячейка_Форма_Отчетность_ЕСУП_итог_str).Column

      ' Отчетность ЕСУП на ___
      ThisWorkbook.Sheets("ЕСУП").Cells(ячейка_Форма_Отчетность_ЕСУП_итог_row - 1, ячейка_Форма_Отчетность_ЕСУП_итог_Column).Value = "Отчетность ЕСУП на " + CStr(dateReportFrom_ReportName_String) + " г."
                      
      ' Тема2:
      Тема2_Range_str = RangeByValue(ThisWorkbook.Name, "ЕСУП", "Тема2:", 100, 100)
      Тема2_Range_Row = Workbooks(ThisWorkbook.Name).Sheets("ЕСУП").Range(Тема2_Range_str).Row
      Тема2_Range_Column = Workbooks(ThisWorkbook.Name).Sheets("ЕСУП").Range(Тема2_Range_str).Column
      ThisWorkbook.Sheets("ЕСУП").Cells(Тема2_Range_Row, Тема2_Range_Column + 1).Value = "Отчетность ЕСУП на " + CStr(dateReportFrom_ReportName_String) + " г."
    
      ' Хэштэг2: "#есуп"+strDDMMYYYY
      Хэштэг2_Range_str = RangeByValue(ThisWorkbook.Name, "ЕСУП", "Хэштэг2:", 100, 100)
      Хэштэг2_Range_Row = Workbooks(ThisWorkbook.Name).Sheets("ЕСУП").Range(Хэштэг2_Range_str).Row
      Хэштэг2_Range_Column = Workbooks(ThisWorkbook.Name).Sheets("ЕСУП").Range(Хэштэг2_Range_str).Column
      ThisWorkbook.Sheets("ЕСУП").Cells(Хэштэг2_Range_Row, Хэштэг2_Range_Column + 1).Value = "#есуп_" + strDDMMYYYY(dateReportFrom_ReportName_String)
                                
      ' Находим столбец "ДО" на листе "ДО_Рег.Сеть"
      ячейка_ДО_Лист_ДО_Рег_Сеть_Range_str = RangeByValue(ReportName_String, "ДО_Рег.Сеть", "ДО", 100, 100)
      ячейка_ДО_Лист_ДО_Рег_Сеть_Range_Row = Workbooks(ReportName_String).Sheets("ДО_Рег.Сеть").Range(ячейка_ДО_Лист_ДО_Рег_Сеть_Range_str).Row
      ячейка_ДО_Лист_ДО_Рег_Сеть_Range_Column = Workbooks(ReportName_String).Sheets("ДО_Рег.Сеть").Range(ячейка_ДО_Лист_ДО_Рег_Сеть_Range_str).Column

      ' Находим в первой строке начальную дату месяца отчета monthStartDate (Здесь на 160-ом столбце возникает ошибка при использовании функции из интернет ConvertToLetter)
      ' ячейка_месяц_str = RangeByValue(ReportName_String, "ДО_Рег.Сеть", CStr(monthStartDate(dateReportFrom_ReportName_String)), 100, 10000)
      ' ячейка_месяц_Row = Workbooks(ReportName_String).Sheets("ДО_Рег.Сеть").Range(ячейка_месяц_str).Row
      ' ячейка_месяц_Column = Workbooks(ReportName_String).Sheets("ДО_Рег.Сеть").Range(ячейка_месяц_str).Column

      ячейка_месяц_str = cellByValue(ReportName_String, "ДО_Рег.Сеть", CStr(monthStartDate(dateReportFrom_ReportName_String)), 100, 10000)
      ячейка_месяц_Row = row_cellByValue(ячейка_месяц_str)
      ячейка_месяц_Column = column_cellByValue(ячейка_месяц_str)

      ' Очищаем ячейки отчета
      ' Call clearСontents2(ThisWorkbook.Name, "ЕСУП", "X7", "AJ12")
      Call clearСontents3(ThisWorkbook.Name, "ЕСУП", ячейка_Форма_Отчетность_ЕСУП_итог_row + 3, ячейка_Форма_Отчетность_ЕСУП_итог_Column + 3, ячейка_Форма_Отчетность_ЕСУП_итог_row + 8, ячейка_Форма_Отчетность_ЕСУП_итог_Column + 12)

      ' Список офисов с исполнением ЕСУП менее 85%
      Список_офисов_с_исполнением_ЕСУП_менее_85 = ""
 
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

        rowCount = 4
        office_was_found = False
        Do While (Not IsEmpty(Workbooks(ReportName_String).Sheets("ДО_Рег.Сеть").Cells(rowCount, ячейка_ДО_Лист_ДО_Рег_Сеть_Range_Column).Value)) And (office_was_found = False)
          
          ' Проверка офиса
          If InStr(Workbooks(ReportName_String).Sheets("ДО_Рег.Сеть").Cells(rowCount, ячейка_ДО_Лист_ДО_Рег_Сеть_Range_Column).Value, officeNameInReport) <> 0 Then
            
            ' Записываем данные
            ' Индивидуальные встречи: План
            ThisWorkbook.Sheets("ЕСУП").Cells(ячейка_Форма_Отчетность_ЕСУП_итог_row + 2 + i, ячейка_Форма_Отчетность_ЕСУП_итог_Column + 3).Value = Workbooks(ReportName_String).Sheets("ДО_Рег.Сеть").Cells(rowCount, ячейка_месяц_Column).Value
            ' Индивидуальные встречи: Факт
            ThisWorkbook.Sheets("ЕСУП").Cells(ячейка_Форма_Отчетность_ЕСУП_итог_row + 2 + i, ячейка_Форма_Отчетность_ЕСУП_итог_Column + 4).Value = Workbooks(ReportName_String).Sheets("ДО_Рег.Сеть").Cells(rowCount, ячейка_месяц_Column + 1).Value
            ' Индивидуальные встречи:  % вып.
            ThisWorkbook.Sheets("ЕСУП").Cells(ячейка_Форма_Отчетность_ЕСУП_итог_row + 2 + i, ячейка_Форма_Отчетность_ЕСУП_итог_Column + 5).Value = Workbooks(ReportName_String).Sheets("ДО_Рег.Сеть").Cells(rowCount, ячейка_месяц_Column + 2).Value
            ThisWorkbook.Sheets("ЕСУП").Cells(ячейка_Форма_Отчетность_ЕСУП_итог_row + 2 + i, ячейка_Форма_Отчетность_ЕСУП_итог_Column + 5).NumberFormat = cellsNumberFormat(ThisWorkbook.Sheets("ЕСУП").Cells(ячейка_Форма_Отчетность_ЕСУП_итог_row + 2 + i, ячейка_Форма_Отчетность_ЕСУП_итог_Column + 5).Value) ' "0.0%"
            
            ' Наблюдения за работой сотр.: План
            ThisWorkbook.Sheets("ЕСУП").Cells(ячейка_Форма_Отчетность_ЕСУП_итог_row + 2 + i, ячейка_Форма_Отчетность_ЕСУП_итог_Column + 6).Value = Workbooks(ReportName_String).Sheets("ДО_Рег.Сеть").Cells(rowCount, ячейка_месяц_Column + 6).Value
            ' Наблюдения за работой сотр.: Факт
            ThisWorkbook.Sheets("ЕСУП").Cells(ячейка_Форма_Отчетность_ЕСУП_итог_row + 2 + i, ячейка_Форма_Отчетность_ЕСУП_итог_Column + 7).Value = Workbooks(ReportName_String).Sheets("ДО_Рег.Сеть").Cells(rowCount, ячейка_месяц_Column + 7).Value
            ' Наблюдения за работой сотр.:  % вып.
            ThisWorkbook.Sheets("ЕСУП").Cells(ячейка_Форма_Отчетность_ЕСУП_итог_row + 2 + i, ячейка_Форма_Отчетность_ЕСУП_итог_Column + 8).Value = Workbooks(ReportName_String).Sheets("ДО_Рег.Сеть").Cells(rowCount, ячейка_месяц_Column + 8).Value
            ThisWorkbook.Sheets("ЕСУП").Cells(ячейка_Форма_Отчетность_ЕСУП_итог_row + 2 + i, ячейка_Форма_Отчетность_ЕСУП_итог_Column + 8).NumberFormat = cellsNumberFormat(ThisWorkbook.Sheets("ЕСУП").Cells(ячейка_Форма_Отчетность_ЕСУП_итог_row + 2 + i, ячейка_Форма_Отчетность_ЕСУП_итог_Column + 8).Value) ' "0.0%"
            
            ' Собрания: План
            ThisWorkbook.Sheets("ЕСУП").Cells(ячейка_Форма_Отчетность_ЕСУП_итог_row + 2 + i, ячейка_Форма_Отчетность_ЕСУП_итог_Column + 9).Value = Workbooks(ReportName_String).Sheets("ДО_Рег.Сеть").Cells(rowCount, ячейка_месяц_Column + 9).Value
            ' Собрания: Факт
            ThisWorkbook.Sheets("ЕСУП").Cells(ячейка_Форма_Отчетность_ЕСУП_итог_row + 2 + i, ячейка_Форма_Отчетность_ЕСУП_итог_Column + 10).Value = Workbooks(ReportName_String).Sheets("ДО_Рег.Сеть").Cells(rowCount, ячейка_месяц_Column + 10).Value
            ' Собрания: % вып.
            ThisWorkbook.Sheets("ЕСУП").Cells(ячейка_Форма_Отчетность_ЕСУП_итог_row + 2 + i, ячейка_Форма_Отчетность_ЕСУП_итог_Column + 11).Value = Workbooks(ReportName_String).Sheets("ДО_Рег.Сеть").Cells(rowCount, ячейка_месяц_Column + 11).Value
            ThisWorkbook.Sheets("ЕСУП").Cells(ячейка_Форма_Отчетность_ЕСУП_итог_row + 2 + i, ячейка_Форма_Отчетность_ЕСУП_итог_Column + 11).NumberFormat = cellsNumberFormat(ThisWorkbook.Sheets("ЕСУП").Cells(ячейка_Форма_Отчетность_ЕСУП_итог_row + 2 + i, ячейка_Форма_Отчетность_ЕСУП_итог_Column + 11).Value) ' "0.0%"
            
            ' Итоговый рейтинг
            ThisWorkbook.Sheets("ЕСУП").Cells(ячейка_Форма_Отчетность_ЕСУП_итог_row + 2 + i, ячейка_Форма_Отчетность_ЕСУП_итог_Column + 12).Value = Workbooks(ReportName_String).Sheets("ДО_Рег.Сеть").Cells(rowCount, ячейка_месяц_Column + 12).Value
            ThisWorkbook.Sheets("ЕСУП").Cells(ячейка_Форма_Отчетность_ЕСУП_итог_row + 2 + i, ячейка_Форма_Отчетность_ЕСУП_итог_Column + 12).NumberFormat = cellsNumberFormat(ThisWorkbook.Sheets("ЕСУП").Cells(ячейка_Форма_Отчетность_ЕСУП_итог_row + 2 + i, ячейка_Форма_Отчетность_ЕСУП_итог_Column + 12).Value) ' "0.0%"
            
            ' Список офисов с исполнением ЕСУП менее 85%
            If ThisWorkbook.Sheets("ЕСУП").Cells(ячейка_Форма_Отчетность_ЕСУП_итог_row + 2 + i, ячейка_Форма_Отчетность_ЕСУП_итог_Column + 12).Value < 0.85 Then
              If Список_офисов_с_исполнением_ЕСУП_менее_85 = "" Then
                Список_офисов_с_исполнением_ЕСУП_менее_85 = getNameOfficeByNumber(i)
              Else
                Список_офисов_с_исполнением_ЕСУП_менее_85 = Список_офисов_с_исполнением_ЕСУП_менее_85 + ", " + getNameOfficeByNumber(i)
              End If
            End If
            
            ' Условное форматирование
            Call cellFormatConditions(ThisWorkbook.Name, "ЕСУП", ячейка_Форма_Отчетность_ЕСУП_итог_row + 2 + i, ячейка_Форма_Отчетность_ЕСУП_итог_Column + 12)
                        
            ' Переменная - Офис был найден
            office_was_found = True
          End If
        
          ' Следующая запись
          rowCount = rowCount + 1
          Application.StatusBar = officeNameInReport + ": " + CStr(rowCount) + "..."
          DoEventsInterval (rowCount)
        Loop
      
        ' Выводим данные по офису
      
      Next i ' Следующий офис
      
      ' ЕСУП по Ипотеке
        
      ' Находим столбец "Тюменский ОО1" на листе "ИПОТЕКА"
      ячейка_Тюменский_ОО1_str = RangeByValue(ReportName_String, "ИПОТЕКА", "Тюменский ОО1", 100, 100)
      ячейка_Тюменский_ОО1_Row = Workbooks(ReportName_String).Sheets("ДО_Рег.Сеть").Range(ячейка_Тюменский_ОО1_str).Row
      ячейка_Тюменский_ОО1_Column = Workbooks(ReportName_String).Sheets("ДО_Рег.Сеть").Range(ячейка_Тюменский_ОО1_str).Column

      ' Находим в первой строке начальную дату месяца отчета monthStartDate
      ' ячейка_месяц_str = RangeByValue(ReportName_String, "ИПОТЕКА", CStr(monthStartDate(dateReportFrom_ReportName_String)), 100, 10000)
      ' ячейка_месяц_Row = Workbooks(ReportName_String).Sheets("ИПОТЕКА").Range(ячейка_месяц_str).Row
      ' ячейка_месяц_Column = Workbooks(ReportName_String).Sheets("ИПОТЕКА").Range(ячейка_месяц_str).Column
          
      ячейка_месяц_Row = rowByValue(ReportName_String, "ИПОТЕКА", CStr(monthStartDate(dateReportFrom_ReportName_String)), 100, 10000)
      ячейка_месяц_Column = ColumnByValue(ReportName_String, "ИПОТЕКА", CStr(monthStartDate(dateReportFrom_ReportName_String)), 100, 10000)
      
            ' Записываем данные
            ' ИПОТЕКА Индивидуальные встречи: План
            ThisWorkbook.Sheets("ЕСУП").Cells(ячейка_Форма_Отчетность_ЕСУП_итог_row + 2 + i, ячейка_Форма_Отчетность_ЕСУП_итог_Column + 3).Value = Workbooks(ReportName_String).Sheets("ИПОТЕКА").Cells(ячейка_Тюменский_ОО1_Row, ячейка_месяц_Column + 6).Value
            ' ИПОТЕКА Индивидуальные встречи: Факт
            ThisWorkbook.Sheets("ЕСУП").Cells(ячейка_Форма_Отчетность_ЕСУП_итог_row + 2 + i, ячейка_Форма_Отчетность_ЕСУП_итог_Column + 4).Value = Workbooks(ReportName_String).Sheets("ИПОТЕКА").Cells(ячейка_Тюменский_ОО1_Row, ячейка_месяц_Column + 7).Value
            ' ИПОТЕКА Индивидуальные встречи:  % вып.
            ThisWorkbook.Sheets("ЕСУП").Cells(ячейка_Форма_Отчетность_ЕСУП_итог_row + 2 + i, ячейка_Форма_Отчетность_ЕСУП_итог_Column + 5).Value = Workbooks(ReportName_String).Sheets("ИПОТЕКА").Cells(ячейка_Тюменский_ОО1_Row, ячейка_месяц_Column + 8).Value
            ThisWorkbook.Sheets("ЕСУП").Cells(ячейка_Форма_Отчетность_ЕСУП_итог_row + 2 + i, ячейка_Форма_Отчетность_ЕСУП_итог_Column + 5).NumberFormat = cellsNumberFormat(ThisWorkbook.Sheets("ЕСУП").Cells(ячейка_Форма_Отчетность_ЕСУП_итог_row + 2 + i, ячейка_Форма_Отчетность_ЕСУП_итог_Column + 5).Value) ' "0.0%"
            
            ' ИПОТЕКА Наблюдения за работой сотр.: План
            ThisWorkbook.Sheets("ЕСУП").Cells(ячейка_Форма_Отчетность_ЕСУП_итог_row + 2 + i, ячейка_Форма_Отчетность_ЕСУП_итог_Column + 6).Value = Workbooks(ReportName_String).Sheets("ИПОТЕКА").Cells(ячейка_Тюменский_ОО1_Row, ячейка_месяц_Column + 3).Value
            ' ИПОТЕКА Наблюдения за работой сотр.: Факт
            ThisWorkbook.Sheets("ЕСУП").Cells(ячейка_Форма_Отчетность_ЕСУП_итог_row + 2 + i, ячейка_Форма_Отчетность_ЕСУП_итог_Column + 7).Value = Workbooks(ReportName_String).Sheets("ИПОТЕКА").Cells(ячейка_Тюменский_ОО1_Row, ячейка_месяц_Column + 4).Value
            ' ИПОТЕКА Наблюдения за работой сотр.:  % вып.
            ThisWorkbook.Sheets("ЕСУП").Cells(ячейка_Форма_Отчетность_ЕСУП_итог_row + 2 + i, ячейка_Форма_Отчетность_ЕСУП_итог_Column + 8).Value = Workbooks(ReportName_String).Sheets("ИПОТЕКА").Cells(ячейка_Тюменский_ОО1_Row, ячейка_месяц_Column + 5).Value
            ThisWorkbook.Sheets("ЕСУП").Cells(ячейка_Форма_Отчетность_ЕСУП_итог_row + 2 + i, ячейка_Форма_Отчетность_ЕСУП_итог_Column + 8).NumberFormat = cellsNumberFormat(ThisWorkbook.Sheets("ЕСУП").Cells(ячейка_Форма_Отчетность_ЕСУП_итог_row + 2 + i, ячейка_Форма_Отчетность_ЕСУП_итог_Column + 8).Value) ' "0.0%"
            
            ' ИПОТЕКА Собрания: План
            ThisWorkbook.Sheets("ЕСУП").Cells(ячейка_Форма_Отчетность_ЕСУП_итог_row + 2 + i, ячейка_Форма_Отчетность_ЕСУП_итог_Column + 9).Value = Workbooks(ReportName_String).Sheets("ИПОТЕКА").Cells(ячейка_Тюменский_ОО1_Row, ячейка_месяц_Column + 0).Value
            ' ИПОТЕКА Собрания: Факт
            ThisWorkbook.Sheets("ЕСУП").Cells(ячейка_Форма_Отчетность_ЕСУП_итог_row + 2 + i, ячейка_Форма_Отчетность_ЕСУП_итог_Column + 10).Value = Workbooks(ReportName_String).Sheets("ИПОТЕКА").Cells(ячейка_Тюменский_ОО1_Row, ячейка_месяц_Column + 1).Value
            ' ИПОТЕКА Собрания: % вып.
            ThisWorkbook.Sheets("ЕСУП").Cells(ячейка_Форма_Отчетность_ЕСУП_итог_row + 2 + i, ячейка_Форма_Отчетность_ЕСУП_итог_Column + 11).Value = Workbooks(ReportName_String).Sheets("ИПОТЕКА").Cells(ячейка_Тюменский_ОО1_Row, ячейка_месяц_Column + 2).Value
            ThisWorkbook.Sheets("ЕСУП").Cells(ячейка_Форма_Отчетность_ЕСУП_итог_row + 2 + i, ячейка_Форма_Отчетность_ЕСУП_итог_Column + 11).NumberFormat = cellsNumberFormat(ThisWorkbook.Sheets("ЕСУП").Cells(ячейка_Форма_Отчетность_ЕСУП_итог_row + 2 + i, ячейка_Форма_Отчетность_ЕСУП_итог_Column + 11).Value) ' "0.0%"
            
            ' ИПОТЕКА Итоговый рейтинг
            ThisWorkbook.Sheets("ЕСУП").Cells(ячейка_Форма_Отчетность_ЕСУП_итог_row + 2 + i, ячейка_Форма_Отчетность_ЕСУП_итог_Column + 12).Value = Workbooks(ReportName_String).Sheets("ИПОТЕКА").Cells(ячейка_Тюменский_ОО1_Row, ячейка_месяц_Column + 9).Value
            ThisWorkbook.Sheets("ЕСУП").Cells(ячейка_Форма_Отчетность_ЕСУП_итог_row + 2 + i, ячейка_Форма_Отчетность_ЕСУП_итог_Column + 12).NumberFormat = cellsNumberFormat(ThisWorkbook.Sheets("ЕСУП").Cells(ячейка_Форма_Отчетность_ЕСУП_итог_row + 2 + i, ячейка_Форма_Отчетность_ЕСУП_итог_Column + 12).Value) ' "0.0%"
    
            ' Условное форматирование
            Call cellFormatConditions(ThisWorkbook.Name, "ЕСУП", ячейка_Форма_Отчетность_ЕСУП_итог_row + 2 + i, ячейка_Форма_Отчетность_ЕСУП_итог_Column + 12)

            ' Список офисов с исполнением ЕСУП менее 85%
            If ThisWorkbook.Sheets("ЕСУП").Cells(ячейка_Форма_Отчетность_ЕСУП_итог_row + 2 + i, ячейка_Форма_Отчетность_ЕСУП_итог_Column + 12).Value < 0.85 Then
              If Список_офисов_с_исполнением_ЕСУП_менее_85 = "" Then
                Список_офисов_с_исполнением_ЕСУП_менее_85 = "ИЦ"
              Else
                Список_офисов_с_исполнением_ЕСУП_менее_85 = Список_офисов_с_исполнением_ЕСУП_менее_85 + ", ИЦ"
              End If
            End If



      ' Переменная завершения обработки
      finishProcess = True
    Else
      ' Сообщение о неверном формате отчета или даты
      MsgBox ("Проверьте отчет: " + CheckFormatReportResult + "!")
    End If ' Проверка формы отчета

    ' Закрываем файл с отчетом без сохранения изменений (параметр SaveChanges:=False)
    Workbooks(Dir(FileName)).Close SaveChanges:=False
    
    ' Формируем информацию в протокол Собрания через Включить_в_Собрание_Повестка_дня()
    ' 1) Заносим в ячейку "Включить в Собрание "Повестка_дня":"
    Row_Включить_в_Собрание = rowByValue(ThisWorkbook.Name, "Лист0", "Включить в Собрание " + Chr(34) + "Повестка_дня" + Chr(34) + ":", 100, 100)
    Column_Включить_в_Собрание = ColumnByValue(ThisWorkbook.Name, "Лист0", "Включить в Собрание " + Chr(34) + "Повестка_дня" + Chr(34) + ":", 100, 100)
    
    ' Сформировать поручение для офисов у которых менее 85%
    If Список_офисов_с_исполнением_ЕСУП_менее_85 <> "" Then
      поручениеЕСУПдо85процентов = "Руководителям " + Список_офисов_с_исполнением_ЕСУП_менее_85 + " провести в подразделениях Собрания, Индивидуальные встречи, Наблюдения за работой сотрудников для достижения уровня Итогового рейтинга за " + ИмяМесяцаГод(dateReportFrom_ReportName_String) + " - не менее 85%."
    Else
      поручениеЕСУПдо85процентов = ""
    End If
    
    ' Тема:
    ThisWorkbook.Sheets("Лист0").Cells(Row_Включить_в_Собрание + 2, Column_Включить_в_Собрание).Value = ThisWorkbook.Sheets("ЕСУП").Cells(ячейка_Форма_Отчетность_ЕСУП_итог_row - 1, ячейка_Форма_Отчетность_ЕСУП_итог_Column).Value + " Исполнение: " _
                                                                                                        + "ОО «Тюменский» " + CStr(Round(ThisWorkbook.Sheets("ЕСУП").Cells(ячейка_Форма_Отчетность_ЕСУП_итог_row + 2 + 1, ячейка_Форма_Отчетность_ЕСУП_итог_Column + 12).Value * 100, 0)) + "%, " _
                                                                                                          + "ОО «Сургутский» " + CStr(Round(ThisWorkbook.Sheets("ЕСУП").Cells(ячейка_Форма_Отчетность_ЕСУП_итог_row + 2 + 2, ячейка_Форма_Отчетность_ЕСУП_итог_Column + 12).Value * 100, 0)) + "%, " _
                                                                                                            + "ОО «Нижневартовский» " + CStr(Round(ThisWorkbook.Sheets("ЕСУП").Cells(ячейка_Форма_Отчетность_ЕСУП_итог_row + 2 + 3, ячейка_Форма_Отчетность_ЕСУП_итог_Column + 12).Value * 100, 0)) + "%, " _
                                                                                                              + "ОО «Новоуренгойский» " + CStr(Round(ThisWorkbook.Sheets("ЕСУП").Cells(ячейка_Форма_Отчетность_ЕСУП_итог_row + 2 + 4, ячейка_Форма_Отчетность_ЕСУП_итог_Column + 12).Value * 100, 0)) + "%, " _
                                                                                                                + "ОО «Тарко-Сале» " + CStr(Round(ThisWorkbook.Sheets("ЕСУП").Cells(ячейка_Форма_Отчетность_ЕСУП_итог_row + 2 + 5, ячейка_Форма_Отчетность_ЕСУП_итог_Column + 12).Value * 100, 0)) + "%, " _
                                                                                                                  + "Ипотечный центр " + CStr(Round(ThisWorkbook.Sheets("ЕСУП").Cells(ячейка_Форма_Отчетность_ЕСУП_итог_row + 2 + 6, ячейка_Форма_Отчетность_ЕСУП_итог_Column + 12).Value * 100, 0)) + "%. " _
                                                                                                                    + поручениеЕСУПдо85процентов
        
    ' HashTag на Лист0 из Хэштэг2:
    ThisWorkbook.Sheets("Лист0").Cells(Row_Включить_в_Собрание + 2, Column_Включить_в_Собрание + 14).Value = ThisWorkbook.Sheets("ЕСУП").Cells(Хэштэг2_Range_Row, Хэштэг2_Range_Column + 1).Value

    ' 2) Включить_в_Собрание_Повестка_дня()
    Call Включить_в_Собрание_Повестка_дня
    
    ' Формируем список для отправки (в "Список получателей2:"):
    range_Список_получателей = RangeByValue(ThisWorkbook.Name, "ЕСУП", "Список получателей2:", 100, 100)
    row_Список_получателей = ThisWorkbook.Sheets("ЕСУП").Range(range_Список_получателей).Row
    column_Список_получателей = ThisWorkbook.Sheets("ЕСУП").Range(range_Список_получателей).Column
    '
    ThisWorkbook.Sheets("ЕСУП").Cells(row_Список_получателей, column_Список_получателей + 2).Value = getFromAddrBook("УДО2,УДО3,УДО4,УДО5,НОРПиКО1,НОРПиКО2,НОРПиКО3,НОРПиКО4,НОРПиКО5,РИЦ", 2)
    ThisWorkbook.Sheets("ЕСУП").Cells(row_Список_получателей, column_Список_получателей + 3).Value = " "

    ' Переходим в ячейку M2
    ThisWorkbook.Sheets("ЕСУП").Range("AL4").Select

    ' Строка статуса
    Application.StatusBar = ""

    ' Зачеркиваем пункт меню на стартовой страницы
    ' Call ЗачеркиваемТекстВячейке("Лист0", RangeByValue(ThisWorkbook.Name, "Лист0", "Оперативная справка по _________________", 100, 100))
    
    ' Итоговое сообщение
    If finishProcess = True Then
      MsgBox ("Обработка " + Dir(ReportName_String) + " завершена!")
    Else
      MsgBox ("Обработка отчета была прервана!")
    End If

  End If ' Если файл был выбран


End Sub

' Отправка письма: отправляю шаблон самому себе для последующей отправки в сеть письма на его основе
Sub Отправка_Lotus_Notes_Лист_ЕСУП_ОО_и_ИЦ()
'
Dim темаПисьма, текстПисьма, hashTag As String
Dim i As Byte
  
  If MsgBox("Отправить себе Шаблон письма?", vbYesNo) = vbYes Then
    
    ' Тема письма - Тема:
    темаПисьма = subjectFromSheet2("ЕСУП")

    ' hashTag - Хэштэг:
    hashTag = hashTagFromSheet2("ЕСУП")
    
    ' Текст письма
    текстПисьма = "" + Chr(13)
    текстПисьма = текстПисьма + "" + recipientList2("ЕСУП") + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "" + getFromAddrBook("РД", 2) + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "Уважаемые руководители," + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "Информация по текущему исполнению отчетности. Норматив итогового рейтинга - 85%." + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    ' Визитка (подпись С Ув., )
    текстПисьма = текстПисьма + ПодписьВПисьме()
    ' Хэштег
    текстПисьма = текстПисьма + createBlankStr(20) + hashTag
    ' Вызов
    Call send_Lotus_Notes2(темаПисьма, "Sergey Fedorovich Proschaev/Tyumen/PSBank/Ru", "", "", текстПисьма, "")
  
    ' Сообщение
    MsgBox ("Письмо отправлено!")
     
  End If


End Sub

' Условное форматирование
Sub cellFormatConditions(In_BookName, In_Sheet, In_Row, In_Column)
                       
            Workbooks(In_BookName).Sheets(In_Sheet).Cells(In_Row, In_Column).FormatConditions.AddIconSetCondition
            Workbooks(In_BookName).Sheets(In_Sheet).Cells(In_Row, In_Column).FormatConditions(Workbooks(In_BookName).Sheets(In_Sheet).Cells(In_Row, In_Column).FormatConditions.Count).SetFirstPriority
            
            With Workbooks(In_BookName).Sheets(In_Sheet).Cells(In_Row, In_Column).FormatConditions(1)
              .ReverseOrder = False
              .ShowIconOnly = False
              .IconSet = ActiveWorkbook.IconSets(xl3TrafficLights1)
            End With
            
            With Workbooks(In_BookName).Sheets(In_Sheet).Cells(In_Row, In_Column).FormatConditions(1).IconCriteria(2)
              .Type = xlConditionValueNumber
              .Value = 0.5
              .Operator = 7
            End With
            
            With Workbooks(In_BookName).Sheets(In_Sheet).Cells(In_Row, In_Column).FormatConditions(1).IconCriteria(3)
              .Type = xlConditionValueNumber
              .Value = 0.85
              .Operator = 7
            End With

End Sub


' Выгрузка протокола собрания ЕСУП 2
Sub Протокол_Собрания2()
  
Dim FileProtocolName, str_Присутствовавшие_на_Собрании, str_Отсутствовавшие_на_Собрании, str_копияВадрес, К_пор_Range, str_Поручениеi, Присутств_на_Собрании_Range, range_Список_получателей As String
Dim НомерСтроки_Повестка_дня, Номер_вопроса, rowCount, К_пор_Row, К_пор_Column, Номер_поручения, текущаяСтрокаПротокола, i, Присутств_на_Собрании_Row, Присутств_на_Собрании_Column As Byte
Dim row_column_Список_получателей, column_Список_получателей As Byte

  ' Закрытие Протокола (из макроса)
  ' Workbooks.Open FileName:="C:\Users\Сергей\Documents\#VBA\DB_Result\Templates\Приложение 1. Протокол.xlsx"
  ' ChDir "C:\Users\Сергей\Documents\#VBA\DB_Result\Out"
  ' ActiveWorkbook.SaveAs FileName:="C:\Users\Сергей\Documents\#VBA\DB_Result\Out\Протокол_собрания.xlsx", FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
  ' ActiveWindow.Close

  ' Запрос на формирование протокола Собрания
  If MsgBox("Сформировать протокол Собрания?", vbYesNo) = vbYes Then
    
    ' Открываем таблицу BASE\ToDo
    OpenBookInBase ("ToDo")
    ' Переходим на окно DB
    ThisWorkbook.Sheets("ЕСУП").Activate

    
    ' Открываем шаблон Протокола из C:\Users\...\Documents\#VBA\DB_Result\Templates
    Workbooks.Open (ThisWorkbook.Path + "\Templates\Приложение 1. Протокол.xlsx")
         
    ' Имя файла с протоколом - берем из G2 "10-02032020"
    FileProtocolName = "Протокол _ РОО Тюменский_" + CStr(dateProtocol(ThisWorkbook.Sheets("ЕСУП").Range("G2").Value)) + ".xlsx"
    
    ' Проверяем - если файл есть, то удаляем его
    Call deleteFile(ThisWorkbook.Path + "\Out\" + FileProtocolName)
    
    ' Сохраняем файл с протоколом
    Workbooks("Приложение 1. Протокол.xlsx").SaveAs FileName:=ThisWorkbook.Path + "\Out\" + FileProtocolName, FileFormat:=xlOpenXMLWorkbook, createBackUp:=False
    
    ' Номер протокола
    Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Range("C1:E2").Value = "Протокол Собрания №" + ThisWorkbook.Sheets("ЕСУП").Range("G2").Value
    Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Range("C1:G2").MergeCells = True
    ' Увеличиваем ширину предпоследнего столбца
    Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Columns("H:H").ColumnWidth = 20.43  ' 20.43, 21.64-предел
    Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Columns("I:I").ColumnWidth = 3
    ' Тема
    Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Range("A4:C4").HorizontalAlignment = xlCenter
    Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Range("A4:C4").VerticalAlignment = xlCenter
    ' Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Range("D4:H4").Value = "Еженедельный конференц-колл с Управляющими офисов и НОРПиКО Тюменского РОО по подведению итогов работы офисов за предыдущую неделю и постановке бизнес-целей на текущую неделю"
    Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Range("D4:H4").Value = "Еженедельный конференц-колл с Управляющими офисов и НОРПиКО Тюменского РОО по подведению итогов работы офисов за предыдущую неделю и постановке бизнес-целей на период с " + strDDMM(weekStartDate(dateProtocol(ThisWorkbook.Sheets("ЕСУП").Range("G2").Value))) + " по " + CStr(weekEndDate(dateProtocol(ThisWorkbook.Sheets("ЕСУП").Range("G2").Value))) + " г."
    ' Дата проведения
    Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Range("A5:C5").HorizontalAlignment = xlCenter
    Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Range("A5:C5").VerticalAlignment = xlCenter
    Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Range("D5:H5").Value = CStr(dateProtocol(ThisWorkbook.Sheets("ЕСУП").Range("G2").Value)) + " г."
    ' Место проведения
    Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Range("A6:C6").HorizontalAlignment = xlCenter
    Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Range("A6:C6").VerticalAlignment = xlCenter
    Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Range("D6:H6").Value = "г.Тюмень, ул.Советская 51/1"
    ' Участники присутствовали
    Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Range("A7:B8").HorizontalAlignment = xlCenter
    Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Range("A7:B8").VerticalAlignment = xlCenter
    Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Range("C7").HorizontalAlignment = xlCenter
    Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Range("C7").VerticalAlignment = xlCenter
    str_Присутствовавшие_на_Собрании = Присутствовавшие_на_Собрании(1)
    Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Range("D7:H7").Value = str_Присутствовавшие_на_Собрании
    Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Range("D7:H7").WrapText = True
    Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Range("7:7").RowHeight = lineHeight(str_Присутствовавшие_на_Собрании, 15, 60) ' было 50 - норм. 60 - реальная ширина
        
    ' Участники отсутствовали
    Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Range("C8").HorizontalAlignment = xlCenter
    Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Range("C8").VerticalAlignment = xlCenter
    str_Отсутствовавшие_на_Собрании = Присутствовавшие_на_Собрании(0)
    Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Range("D8:H8").Value = str_Отсутствовавшие_на_Собрании
    Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Range("D8:H8").WrapText = True
    Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Range("8:8").RowHeight = lineHeight(str_Отсутствовавшие_на_Собрании, 15, 60) ' было 40 - норм
    ' Копия в адрес:
    Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Range("A9:C9").HorizontalAlignment = xlCenter
    Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Range("A9:C9").VerticalAlignment = xlCenter
    ' str_копияВадрес = getFromAddrBook("КОП", 1) ' + ", " + getFromAddrBook("ККП", 1) - пока без Воронцова
    str_копияВадрес = getFromAddrBook("РД", 1) + ", " + getFromAddrBook("КОП", 1) + ", " + getFromAddrBook("КИп", 1) + ", " + getFromAddrBook("ККП", 1) + ", " + getFromAddrBook("Кaf", 1) + ", " + getFromAddrBook("КПВО", 1)
    Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Range("D9:H9").Value = str_копияВадрес
    Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Range("D9:H9").WrapText = True
    Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Range("9:9").RowHeight = lineHeight(str_копияВадрес, 15, 60) ' было 30 - норм
    
    ' Повестка дня:
    НомерСтроки_Повестка_дня = ThisWorkbook.Sheets("ЕСУП").Range(RangeByValue(ThisWorkbook.Name, "ЕСУП", "Повестка_дня", 100, 100)).Row
    rowCount = 2
    Номер_вопроса = 0
    Do While ThisWorkbook.Sheets("ЕСУП").Cells(НомерСтроки_Повестка_дня + rowCount, 1).Value <> ""
      ' Если у вопроса стоит отметка "1", то вносим его в протокол собрания
      If ThisWorkbook.Sheets("ЕСУП").Cells(НомерСтроки_Повестка_дня + rowCount, 14).Value = "1" Then
        
        Номер_вопроса = Номер_вопроса + 1
        
        ' Если номер вопроса более 6-ти, то вставляем строку
        If Номер_вопроса > 6 Then
          ' Вставляем строку
          Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Range(CStr(12 + Номер_вопроса) + ":" + CStr(12 + Номер_вопроса)).Select
          Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
          ' Нумерация "7." возможна только если формат преобразовать к текстовому ("@")
          Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Cells(12 + Номер_вопроса, 1).NumberFormat = "@"
          Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Cells(12 + Номер_вопроса, 1).Value = CStr(Номер_вопроса) + "."
          Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Cells(12 + Номер_вопроса, 1).HorizontalAlignment = xlLeft
          Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Range("B" + CStr(12 + Номер_вопроса) + ":H" + CStr(12 + Номер_вопроса)).MergeCells = True
          ' Рамка
          Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Range("A" + CStr(12 + Номер_вопроса) + ":H" + CStr(12 + Номер_вопроса)).Select
          With Selection.Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ThemeColor = 3
            .TintAndShade = -0.749992370372631
            .Weight = xlThin
          End With
          With Selection.Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ThemeColor = 3
            .TintAndShade = -0.749992370372631
            .Weight = xlThin
          End With
          With Selection.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ThemeColor = 3
            .TintAndShade = -0.749992370372631
            .Weight = xlThin
          End With
          With Selection.Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ThemeColor = 3
            .TintAndShade = -0.749992370372631
            .Weight = xlThin
          End With
          With Selection.Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .ThemeColor = 3
            .TintAndShade = -0.749992370372631
            .Weight = xlThin
          End With
        
        End If
        
        ' Формат номера пункта Повестки Дня
        Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Cells(12 + Номер_вопроса, 1).HorizontalAlignment = xlCenter
        Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Cells(12 + Номер_вопроса, 1).VerticalAlignment = xlCenter
        ' Вносим в Повестку Дня
        If ThisWorkbook.Sheets("ЕСУП").Cells(НомерСтроки_Повестка_дня + rowCount, 13).Value = "" Then
          Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Cells(12 + Номер_вопроса, 2).Value = ThisWorkbook.Sheets("ЕСУП").Cells(НомерСтроки_Повестка_дня + rowCount, 2).Value + ": " + ThisWorkbook.Sheets("ЕСУП").Cells(НомерСтроки_Повестка_дня + rowCount, 3).Value
        Else
          ' Если есть Хэштэг
          Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Cells(12 + Номер_вопроса, 2).Value = ThisWorkbook.Sheets("ЕСУП").Cells(НомерСтроки_Повестка_дня + rowCount, 2).Value + ": " + ThisWorkbook.Sheets("ЕСУП").Cells(НомерСтроки_Повестка_дня + rowCount, 3).Value + " (" + ThisWorkbook.Sheets("ЕСУП").Cells(НомерСтроки_Повестка_дня + rowCount, 13).Value + ")"
        End If
        Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Cells(12 + Номер_вопроса, 2).HorizontalAlignment = xlLeft
        Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Cells(12 + Номер_вопроса, 2).VerticalAlignment = xlTop
        Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Cells(12 + Номер_вопроса, 2).WrapText = True
        Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Cells(12 + Номер_вопроса, 2).RowHeight = lineHeight(delSym(Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Cells(12 + Номер_вопроса, 2).Value), 15, 90) ' было 15, 65
      End If
      ' Следующая строка
      rowCount = rowCount + 1
    Loop
    
    ' Корректируем номер вопроса, если он менее 6 для того, чтобы корректно рассчитывать число строк в Поручениях
    If Номер_вопроса < 6 Then
      Номер_вопроса = 6
    End If
       
    ' Поручения участникам:
      
      ' Номер поручения
      Номер_поручения = 0
      
      ' Обрабатываем Поручения по офису, где стоят даты
      rowCount = 62
      Do While Trim(ThisWorkbook.Sheets("ЕСУП").Cells(rowCount, 19).Value) <> ""
        
          ' Номер поручения
          Номер_поручения = Номер_поручения + 1
          ' Если номер поручения более 6-ти, то вставляем строку
          If Номер_поручения > 6 Then
            
            ' Вставляем пустую строку в блок "Поручения"
            Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Range(CStr(12 + Номер_вопроса + 4 + Номер_поручения) + ":" + CStr(12 + Номер_вопроса + 4 + Номер_поручения)).Select
            Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
            
            ' Нумерация "7." возможна только если формат преобразовать к текстовому ("@")
            Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Cells(12 + Номер_вопроса + 4 + Номер_поручения, 1).NumberFormat = "@"
            Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Cells(12 + Номер_вопроса + 4 + Номер_поручения, 1).Value = CStr(Номер_поручения) + "."
            
            ' Объединяем B, С, D
            Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Range("B" + CStr(12 + Номер_вопроса + 4 + Номер_поручения) + ":D" + CStr(12 + Номер_вопроса + 4 + Номер_поручения)).MergeCells = True
            
            ' Объединяем G, Н
            Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Range("G" + CStr(12 + Номер_вопроса + 4 + Номер_поручения) + ":H" + CStr(12 + Номер_вопроса + 4 + Номер_поручения)).MergeCells = True
            
            ' Рамка
            Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Range("A" + CStr(12 + Номер_вопроса + 4 + Номер_поручения) + ":H" + CStr(12 + Номер_вопроса + 4 + Номер_поручения)).Select
            With Selection.Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .ThemeColor = 3
                .TintAndShade = -0.749992370372631
                .Weight = xlThin
            End With
            With Selection.Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .ThemeColor = 3
                .TintAndShade = -0.749992370372631
                .Weight = xlThin
            End With
            With Selection.Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .ThemeColor = 3
                .TintAndShade = -0.749992370372631
                .Weight = xlThin
            End With
            With Selection.Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .ThemeColor = 3
                .TintAndShade = -0.749992370372631
                .Weight = xlThin
            End With
            With Selection.Borders(xlInsideVertical)
                .LineStyle = xlContinuous
                .ThemeColor = 3
                .TintAndShade = -0.749992370372631
                .Weight = xlThin
            End With

          End If ' Вставляем новую строку Поручения и нумеруем
          
          ' Номер Поручения (№ п/п) - выравнивание по центру и по вертикали
          Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Cells(12 + Номер_вопроса + 4 + Номер_поручения, 1).HorizontalAlignment = xlCenter
          Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Cells(12 + Номер_вопроса + 4 + Номер_поручения, 1).VerticalAlignment = xlCenter

          ' Поручение
          str_Поручениеi = ThisWorkbook.Sheets("ЕСУП").Cells(rowCount, 21).Value
          
          ' Поручение - "Переносить по словам"
          Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Range("B" + CStr(12 + Номер_вопроса + 4 + Номер_поручения) + ":D" + CStr(12 + Номер_вопроса + 4 + Номер_поручения)).WrapText = True
          ' Поручение - выравнивание по центру и по вертикали
          Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Cells(12 + Номер_вопроса + 4 + Номер_поручения, 2).HorizontalAlignment = xlLeft
          Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Cells(12 + Номер_вопроса + 4 + Номер_поручения, 2).VerticalAlignment = xlTop
          ' Поручение - высота строки
          Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Range(CStr(12 + Номер_вопроса + 4 + Номер_поручения) + ":" + CStr(12 + Номер_вопроса + 4 + Номер_поручения)).RowHeight = lineHeight(str_Поручениеi, 15, 37) ' 20 - норм
          ' Поручение - Запись в протокол
          Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Cells(12 + Номер_вопроса + 4 + Номер_поручения, 2).Value = str_Поручениеi

          ' Ответственный
          ' Вариант 1 - Должность и ФИО
          ' Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Cells(12 + Номер_вопроса + 4 + Номер_поручения, 5).Value = ThisWorkbook.Sheets("ЕСУП").Cells(К_пор_Row - 1, К_пор_Column - 4).Value + " " + ThisWorkbook.Sheets("ЕСУП").Cells(К_пор_Row - 1, К_пор_Column - 3).Value
          ' Вариант 2 - ФИО
          Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Cells(12 + Номер_вопроса + 4 + Номер_поручения, 5).Value = ThisWorkbook.Sheets("ЕСУП").Cells(rowCount, 19).Value
          
          ' Ответственный - "Переносить по словам"
          Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Range("E" + CStr(12 + Номер_вопроса + 4 + Номер_поручения) + ":E" + CStr(12 + Номер_вопроса + 4 + Номер_поручения)).WrapText = True
          ' Ответственный - выравнивание по вертикали и горизонтали
          Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Cells(12 + Номер_вопроса + 4 + Номер_поручения, 5).VerticalAlignment = xlTop
          Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Cells(12 + Номер_вопроса + 4 + Номер_поручения, 5).HorizontalAlignment = xlCenter
          
          ' Срок исполнения
          Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Cells(12 + Номер_вопроса + 4 + Номер_поручения, 6).Value = CStr(ThisWorkbook.Sheets("ЕСУП").Cells(rowCount, 20).Value)
          ' Срок исполнения - выравнивание по центру и по вертикали
          Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Cells(12 + Номер_вопроса + 4 + Номер_поручения, 6).HorizontalAlignment = xlCenter
          Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Cells(12 + Номер_вопроса + 4 + Номер_поручения, 6).VerticalAlignment = xlTop
          
        
        ' Следующая строка
        DoEvents
        rowCount = rowCount + 1
      Loop ' Do While
      
    
    ' Вывести результаты исполнения предидущих поручений из BASE\To-Do у которых Protocol_Number<>"" и Protocol_Number2=""
    ' Обрабатываем Поручения по офису, где стоят даты
    rowCount = 2
    Do While Trim(Workbooks("ToDo").Sheets("Лист1").Cells(rowCount, 1).Value) <> ""
        
      ' Если (Protocol_Number<>"" и Protocol_Number2="") ИЛИ (Protocol_Number<>"" и Protocol_Number2=Текущему протоколу)
      ' If ((Workbooks("ToDo").Sheets("Лист1").Cells(rowCount, 10).Value <> "") And (Workbooks("ToDo").Sheets("Лист1").Cells(rowCount, 13).Value = "")) Or (Workbooks("ToDo").Sheets("Лист1").Cells(rowCount, 10).Value <> "") And (Workbooks("ToDo").Sheets("Лист1").Cells(rowCount, 13).Value = ThisWorkbook.Sheets("ЕСУП").Range("G2").Value) Then
      
      ' Вариант2 (выводились поручения с листа ЕСУП и дублировались они же из To-DO)
      ' Если (Protocol_Number<>"" и Protocol_Number<>Текущему протоколу и Protocol_Number2="") ИЛИ (Protocol_Number<>"" и Protocol_Number2=Текущему протоколу)
      If ((Workbooks("ToDo").Sheets("Лист1").Cells(rowCount, 10).Value <> "") And (Workbooks("ToDo").Sheets("Лист1").Cells(rowCount, 10).Value <> ThisWorkbook.Sheets("ЕСУП").Range("G2").Value) And (Workbooks("ToDo").Sheets("Лист1").Cells(rowCount, 13).Value = "")) Or (Workbooks("ToDo").Sheets("Лист1").Cells(rowCount, 10).Value <> "") And (Workbooks("ToDo").Sheets("Лист1").Cells(rowCount, 13).Value = ThisWorkbook.Sheets("ЕСУП").Range("G2").Value) Then
      
      
          ' Номер поручения
          Номер_поручения = Номер_поручения + 1
          ' Если номер поручения более 6-ти, то вставляем строку
          If Номер_поручения > 6 Then
            
            ' Вставляем пустую строку в блок "Поручения"
            Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Range(CStr(12 + Номер_вопроса + 4 + Номер_поручения) + ":" + CStr(12 + Номер_вопроса + 4 + Номер_поручения)).Select
            Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
            
            ' Нумерация "7." возможна только если формат преобразовать к текстовому ("@")
            Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Cells(12 + Номер_вопроса + 4 + Номер_поручения, 1).NumberFormat = "@"
            Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Cells(12 + Номер_вопроса + 4 + Номер_поручения, 1).Value = CStr(Номер_поручения) + "."
            
            ' Объединяем B, С, D
            Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Range("B" + CStr(12 + Номер_вопроса + 4 + Номер_поручения) + ":D" + CStr(12 + Номер_вопроса + 4 + Номер_поручения)).MergeCells = True
            
            ' Объединяем G, Н
            Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Range("G" + CStr(12 + Номер_вопроса + 4 + Номер_поручения) + ":H" + CStr(12 + Номер_вопроса + 4 + Номер_поручения)).MergeCells = True
            
            ' Рамка
            Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Range("A" + CStr(12 + Номер_вопроса + 4 + Номер_поручения) + ":H" + CStr(12 + Номер_вопроса + 4 + Номер_поручения)).Select
            With Selection.Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .ThemeColor = 3
                .TintAndShade = -0.749992370372631
                .Weight = xlThin
            End With
            With Selection.Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .ThemeColor = 3
                .TintAndShade = -0.749992370372631
                .Weight = xlThin
            End With
            With Selection.Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .ThemeColor = 3
                .TintAndShade = -0.749992370372631
                .Weight = xlThin
            End With
            With Selection.Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .ThemeColor = 3
                .TintAndShade = -0.749992370372631
                .Weight = xlThin
            End With
            With Selection.Borders(xlInsideVertical)
                .LineStyle = xlContinuous
                .ThemeColor = 3
                .TintAndShade = -0.749992370372631
                .Weight = xlThin
            End With

          End If ' Вставляем новую строку Поручения и нумеруем
          
          ' Номер Поручения (№ п/п) - выравнивание по центру и по вертикали
          Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Cells(12 + Номер_вопроса + 4 + Номер_поручения, 1).HorizontalAlignment = xlCenter
          Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Cells(12 + Номер_вопроса + 4 + Номер_поручения, 1).VerticalAlignment = xlCenter

          ' Поручение
          str_Поручениеi = Workbooks("ToDo").Sheets("Лист1").Cells(rowCount, 3).Value
          
          ' Поручение - "Переносить по словам"
          Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Range("B" + CStr(12 + Номер_вопроса + 4 + Номер_поручения) + ":D" + CStr(12 + Номер_вопроса + 4 + Номер_поручения)).WrapText = True
          ' Поручение - выравнивание по центру и по вертикали
          Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Cells(12 + Номер_вопроса + 4 + Номер_поручения, 2).HorizontalAlignment = xlLeft
          Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Cells(12 + Номер_вопроса + 4 + Номер_поручения, 2).VerticalAlignment = xlTop
          ' Поручение - высота строки
          Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Range(CStr(12 + Номер_вопроса + 4 + Номер_поручения) + ":" + CStr(12 + Номер_вопроса + 4 + Номер_поручения)).RowHeight = lineHeight(str_Поручениеi, 15, 37) ' 20 - норм
          ' Поручение - Запись в протокол
          Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Cells(12 + Номер_вопроса + 4 + Номер_поручения, 2).Value = str_Поручениеi

          ' Ответственный
          ' Вариант 1 - Должность и ФИО
          ' Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Cells(12 + Номер_вопроса + 4 + Номер_поручения, 5).Value = ThisWorkbook.Sheets("ЕСУП").Cells(К_пор_Row - 1, К_пор_Column - 4).Value + " " + ThisWorkbook.Sheets("ЕСУП").Cells(К_пор_Row - 1, К_пор_Column - 3).Value
          ' Вариант 2 - ФИО
          Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Cells(12 + Номер_вопроса + 4 + Номер_поручения, 5).Value = Workbooks("ToDo").Sheets("Лист1").Cells(rowCount, 5).Value
          
          ' Ответственный - "Переносить по словам"
          Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Range("E" + CStr(12 + Номер_вопроса + 4 + Номер_поручения) + ":E" + CStr(12 + Номер_вопроса + 4 + Номер_поручения)).WrapText = True
          ' Ответственный - выравнивание по вертикали и горизонтали
          Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Cells(12 + Номер_вопроса + 4 + Номер_поручения, 5).VerticalAlignment = xlTop
          Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Cells(12 + Номер_вопроса + 4 + Номер_поручения, 5).HorizontalAlignment = xlCenter
          
          ' Срок исполнения
          Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Cells(12 + Номер_вопроса + 4 + Номер_поручения, 6).Value = CStr(Workbooks("ToDo").Sheets("Лист1").Cells(rowCount, 8).Value)
          ' Срок исполнения - выравнивание по центру и по вертикали
          Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Cells(12 + Номер_вопроса + 4 + Номер_поручения, 6).HorizontalAlignment = xlCenter
          Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Cells(12 + Номер_вопроса + 4 + Номер_поручения, 6).VerticalAlignment = xlTop
          
          ' Статус выполнения/комментарии
          Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Cells(12 + Номер_вопроса + 4 + Номер_поручения, 7).Value = Workbooks("ToDo").Sheets("Лист1").Cells(rowCount, 9).Value + " (п. " + CStr(Workbooks("ToDo").Sheets("Лист1").Cells(rowCount, 12).Value) + " прот.№" + Workbooks("ToDo").Sheets("Лист1").Cells(rowCount, 10).Value + ")"
          ' Ответственный - "Переносить по словам" для "G39:H39"
          Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Range("G" + CStr(12 + Номер_вопроса + 4 + Номер_поручения) + ":H" + CStr(12 + Номер_вопроса + 4 + Номер_поручения)).WrapText = True
          ' Срок исполнения - выравнивание по центру и по вертикали
          Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Cells(12 + Номер_вопроса + 4 + Номер_поручения, 7).HorizontalAlignment = xlCenter
          Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Cells(12 + Номер_вопроса + 4 + Номер_поручения, 7).VerticalAlignment = xlCenter
          
          ' Если статус поручения в ToDo.Task_Status = 0, т.е. задача уже не активна, то вносим в ToDo.Protocol_Number2, ToDo.Protocol_Date2, ToDo.Protocol_Question_number2
          If Workbooks("ToDo").Sheets("Лист1").Cells(rowCount, 7).Value = 0 Then
            
             ' ToDo.Protocol_Number2
             Workbooks("ToDo").Sheets("Лист1").Cells(rowCount, 13).Value = ThisWorkbook.Sheets("ЕСУП").Range("G2").Value
             
             ' ToDo.Protocol_Date2
             Workbooks("ToDo").Sheets("Лист1").Cells(rowCount, 14).Value = getDateFromProtocolNumber(ThisWorkbook.Sheets("ЕСУП").Range("G2").Value)
             
             ' ToDo.Protocol_Question_number2 = Номер_поручения
             Workbooks("ToDo").Sheets("Лист1").Cells(rowCount, 15).Value = Номер_поручения
             
          End If
          
 
          
        End If ' Если Protocol_Number<>"" и Protocol_Number2=""
        
        ' Следующая строка
        DoEvents
        rowCount = rowCount + 1
      Loop ' Do While
    
    
    ' Здесь контролируем номер строки для вывода текущаяСтрокаПротокола
    If Номер_поручения < 6 Then
      ' Номер поручения
      Номер_поручения = 6
      
      ' Убираем нумерацию строк в поручениях с
      
    End If
    
    ' Моя Подпись под протоколом
    текущаяСтрокаПротокола = (12 + Номер_вопроса + 4 + Номер_поручения) + 2
    Call InsertRow_InProtocol(FileProtocolName, "Протокол_Собрания", текущаяСтрокаПротокола)
    Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Cells(текущаяСтрокаПротокола, 2).Value = "Заместитель директора по развитию розничного бизнеса"
    Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Cells(текущаяСтрокаПротокола, 7).Value = "Прощаев С.Ф."
    текущаяСтрокаПротокола = текущаяСтрокаПротокола + 1
    Call InsertRow_InProtocol(FileProtocolName, "Протокол_Собрания", текущаяСтрокаПротокола)
    текущаяСтрокаПротокола = текущаяСтрокаПротокола + 1
    Call InsertRow_InProtocol(FileProtocolName, "Протокол_Собрания", текущаяСтрокаПротокола)
    
    ' С протоколом ознакомлены: (по электронной почте) - направляем присутствующим и отсутствующим
    Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Cells(текущаяСтрокаПротокола, 2).Value = "C протоколом ознакомлены (по электронной почте):"
    ' Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Cells(текущаяСтрокаПротокола, 2).Font.Underline = xlUnderlineStyleSingle
    текущаяСтрокаПротокола = текущаяСтрокаПротокола + 1
    Call InsertRow_InProtocol(FileProtocolName, "Протокол_Собрания", текущаяСтрокаПротокола)
    '
    Присутств_на_Собрании_Range = RangeByValue(ThisWorkbook.Name, "ЕСУП", "Присутств_на_Собрании", 100, 100)
    Присутств_на_Собрании_Row = ThisWorkbook.Sheets("ЕСУП").Range(Присутств_на_Собрании_Range).Row
    Присутств_на_Собрании_Column = ThisWorkbook.Sheets("ЕСУП").Range(Присутств_на_Собрании_Range).Column
    rowCount = Присутств_на_Собрании_Row + 1
    Do While ThisWorkbook.Sheets("ЕСУП").Cells(rowCount, Присутств_на_Собрании_Column).Value <> "Пригл_на_Собрание"
      
      ' Если ФИО <>0
      If ThisWorkbook.Sheets("ЕСУП").Cells(rowCount, Присутств_на_Собрании_Column + 1).Value <> 0 Then

        ' Должность_и_ФИО = ThisWorkbook.Sheets("ЕСУП").Cells(RowCount, Присутств_на_Собрании_Column + 4).Value + " " + Фамилия_и_Имя(ThisWorkbook.Sheets("ЕСУП").Cells(RowCount, Присутств_на_Собрании_Column + 1).Value, 3)
        Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Cells(текущаяСтрокаПротокола, 2).Value = ThisWorkbook.Sheets("ЕСУП").Cells(rowCount, Присутств_на_Собрании_Column + 4).Value
        Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Cells(текущаяСтрокаПротокола, 6).Value = Фамилия_и_Имя(ThisWorkbook.Sheets("ЕСУП").Cells(rowCount, Присутств_на_Собрании_Column + 1).Value, 3)
        текущаяСтрокаПротокола = текущаяСтрокаПротокола + 1
        Call InsertRow_InProtocol(FileProtocolName, "Протокол_Собрания", текущаяСтрокаПротокола)
      
      End If
         
      ' Следующая запись
      rowCount = rowCount + 1
    Loop
 
    ' Формируем список для отправки (в "Список получателей:"):
    range_Список_получателей = RangeByValue(ThisWorkbook.Name, "ЕСУП", "Список получателей:", 100, 100)
    row_Список_получателей = ThisWorkbook.Sheets("ЕСУП").Range(range_Список_получателей).Row
    column_Список_получателей = ThisWorkbook.Sheets("ЕСУП").Range(range_Список_получателей).Column
    '
    ThisWorkbook.Sheets("ЕСУП").Cells(row_Список_получателей, column_Список_получателей + 2).Value = getFromAddrBook("УДО2,УДО3,УДО4,УДО5,НОРПиКО1,НОРПиКО2,НОРПиКО3,НОРПиКО4,НОРПиКО5,ПМ,МРК1,МРК2,МРК3,МРК4,МРК5,НОКП,РРКК,МПП", 2)
    ThisWorkbook.Sheets("ЕСУП").Cells(row_Список_получателей, column_Список_получателей + 3).Value = " "
    
    ' Перемещаем в ячейку A1
    Workbooks(FileProtocolName).Sheets("Протокол_Собрания").Range("A1").Select
    
    ' Закрытие файла с Протоколом Собрания
    Workbooks(FileProtocolName).Close SaveChanges:=True
    
    ' Редактируем тему письма
    ThisWorkbook.Sheets("ЕСУП").Cells(1, 17).Value = "Тюменский РОО: Протокол конференц-колла с офисами №" + ThisWorkbook.Sheets("ЕСУП").Cells(2, 7).Value + " от " + CStr(dateProtocol(ThisWorkbook.Sheets("ЕСУП").Range("G2").Value)) + " г."
    
    ' Редактируем поля с Тэгами 1 и 2 на листе Есуп
    ThisWorkbook.Sheets("ЕСУП").Cells(1, 15).Value = "#protocol"
    ThisWorkbook.Sheets("ЕСУП").Cells(3, 15).Value = "#protocol_" + ThisWorkbook.Sheets("ЕСУП").Cells(2, 7).Value
    
    ' Закрываем таблицу BASE\ToDo
    CloseBook ("ToDo")
 
    ' Сохранение изменений
    ThisWorkbook.Save

    MsgBox ("Протокол сформирован!")
    
  End If ' Запрос на формирование
  
End Sub


' Добавить красные зоны по месяцу в повестку
Sub add_Focus()
Dim rowCount As Integer
  
  ' Запрос на действие из Лист8 "B5" берем месяц
  If MsgBox("Добавить в повестку красные зоны " + ИмяМесяца2(DashboardDate()) + "?", vbYesNo) = vbYes Then
    
    Application.StatusBar = "Поиск красных зон..."
    
    ' Открываем таблицу MBO
    ' Открываем BASE\Sales
    OpenBookInBase ("MBO")
    ' ThisWorkbook.Sheets("Лист8").Activate

    
    ' Обработка DB
    ' Фокусы_контроля_строка = ""
    Строка_к_выводу = ""

    ' Начало блока офиса
    Начало_блока_офиса = False

    ' Добавлено вопросов в повестку
    Добавлено_вопросов = 0

    ' 1. ИЦ
    продукты_контроля_ИЦ_РОО = "           КК к Ипотеке,            КЛА премия ИЦ"
    
    rowCount = rowByValue(ThisWorkbook.Name, "Лист8", "Итого по РОО «Тюменский»", 500, 500)
    Do While ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value <> ""
    
      ' Если начинается раздел офиса, то Строка_к_выводу обнуляем
      If (InStr(ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value, "Итого по РОО «Тюменский»") <> 0) Then
        
        ' Начало блока офиса
        Начало_блока_офиса = True
        
        ' Имя офиса
        Имя_офиса = "ИЦ"
        Имя_офиса_в_Строке_к_выводу = "ИЦ"

        ' Счетчик красных зон офиса
        Счетчик_красных_зон_офиса = 0
        
      End If
    
      ' ИЦ: Если прогноз по кварталу есть
      If (Начало_блока_офиса = True) And (ThisWorkbook.Sheets("Лист8").Cells(rowCount, 8).Value < 1) And (InStr(продукты_контроля_ИЦ_РОО, ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value) <> 0) Then
      
        Строка_к_выводу = Строка_к_выводу + Сокр(ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value) + " (" + CStr(Round(ThisWorkbook.Sheets("Лист8").Cells(rowCount, 8).Value * 100, 0)) + "%), "
        Счетчик_красных_зон_офиса = Счетчик_красных_зон_офиса + 1
      
        ' Строка поручений
        Строка_поручений = Строка_поручений + Сокр(ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value) + " (+" + CStr(Round(ThisWorkbook.Sheets("Лист8").Cells(rowCount, 13).Value, 0)) + ")___, "
      
      
      End If
    
      ' ИЦ Если прогноза по кварталу нет - берем из 20-ой колонки Лист8 прогноз
      If (Начало_блока_офиса = True) And (ThisWorkbook.Sheets("Лист8").Cells(rowCount, 20).Value < 1) And (InStr(продукты_контроля_ИЦ_РОО, ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value) <> 0) And (ThisWorkbook.Sheets("Лист8").Cells(rowCount, 8).Value = "") Then
        
        Строка_к_выводу = Строка_к_выводу + Сокр(ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value) + " (" + CStr(Round(ThisWorkbook.Sheets("Лист8").Cells(rowCount, 20).Value * 100, 0)) + "%), "
        Счетчик_красных_зон_офиса = Счетчик_красных_зон_офиса + 1
        
        ' Строка поручений
        Строка_поручений = Строка_поручений + Сокр(ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value) + " (+" + CStr(Round(ThisWorkbook.Sheets("Лист8").Cells(rowCount, 13).Value, 0)) + ")___, "
      
      End If
    
    
      ' Если блок офиса закончился
      If (Начало_блока_офиса = True) And (ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value = "Ипотека") Then
        
        ' Вставить строку в Повестку собрания
        If Счетчик_красных_зон_офиса <> 0 Then
          
          ' Вставляем вопрос
          Call Вставка_строки_в_Повестку(Имя_офиса, _
                                           Имя_офиса_в_Строке_к_выводу + " прогноз исп. БП " + quarterName2(dateDB_Лист_8) + ": " + Строка_к_выводу + " Проект поручения: Выдачи ___, " + Строка_поручений, _
                                             "")
          ' Добавлено вопросов в повестку
          Добавлено_вопросов = Добавлено_вопросов + 1
          
        End If
        
        ' Начало блока офиса
        Начало_блока_офиса = False
      
        ' Обнуляем строку
        Строка_к_выводу = ""
        Строка_поручений = ""
      
      End If
    
      ' Следующая запись
      Application.StatusBar = "Анализ прогнозов " + CStr(rowCount) + "..."
      rowCount = rowCount + 1
      DoEventsInterval (rowCount)
  
    Loop

    
    ' 2. ОКП =========================================================================================================
    '  Показатели ОКП
    продукты_контроля_ОКП_РОО = "Зарплатные карты 18+, Портфель ЗП 18+, шт._Квартал ,            КК к ЗП"
    
    rowCount = rowByValue(ThisWorkbook.Name, "Лист8", "Итого по РОО «Тюменский»", 500, 500)
    Do While ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value <> ""
    
      ' Если начинается раздел офиса, то Строка_к_выводу обнуляем
      If (InStr(ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value, "Итого по РОО «Тюменский»") <> 0) Then
        
        ' Начало блока офиса
        Начало_блока_офиса = True
        
        ' Имя офиса
        Имя_офиса = "ОКП"
        Имя_офиса_в_Строке_к_выводу = "ОКП"

        ' Счетчик красных зон офиса
        Счетчик_красных_зон_офиса = 0
        
      End If
    
      ' ОКП: Если прогноз по кварталу есть
      If (Начало_блока_офиса = True) And (ThisWorkbook.Sheets("Лист8").Cells(rowCount, 8).Value < 1) And (InStr(продукты_контроля_ОКП_РОО, ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value) <> 0) Then
      
        Строка_к_выводу = Строка_к_выводу + Сокр(ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value) + " (" + CStr(Round(ThisWorkbook.Sheets("Лист8").Cells(rowCount, 8).Value * 100, 0)) + "%), "
        Счетчик_красных_зон_офиса = Счетчик_красных_зон_офиса + 1
      
      End If
    
      ' ОКП Если прогноза по кварталу нет - берем из 20-ой колонки Лист8 прогноз
      If (Начало_блока_офиса = True) And (ThisWorkbook.Sheets("Лист8").Cells(rowCount, 20).Value < 1) And (InStr(продукты_контроля_ОКП_РОО, ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value) <> 0) And (ThisWorkbook.Sheets("Лист8").Cells(rowCount, 8).Value = "") Then
        
        Строка_к_выводу = Строка_к_выводу + Сокр(ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value) + " (" + CStr(Round(ThisWorkbook.Sheets("Лист8").Cells(rowCount, 20).Value * 100, 0)) + "%), "
        Счетчик_красных_зон_офиса = Счетчик_красных_зон_офиса + 1
      
      End If
    
    
      ' Если блок офиса закончился
      If (Начало_блока_офиса = True) And (ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value = "Ипотека") Then
        
        ' Вставить строку в Повестку собрания
        If Счетчик_красных_зон_офиса <> 0 Then
          
          ' Вставляем вопрос
          Call Вставка_строки_в_Повестку(Имя_офиса, _
                                           Имя_офиса_в_Строке_к_выводу + " прогноз исп. БП " + quarterName2(dateDB_Лист_8) + ": " + Строка_к_выводу + " Проект поручения: Заключ.договоров___, Выпуск карт___, Выдача карт___, КК к ЗП___", _
                                             "")
          ' Добавлено вопросов в повестку
          Добавлено_вопросов = Добавлено_вопросов + 1
          
        End If
        
        ' Начало блока офиса
        Начало_блока_офиса = False
      
        ' Обнуляем строку
        Строка_к_выводу = ""
        Строка_поручений = ""

      End If
    
      ' Следующая запись
      Application.StatusBar = "Анализ прогнозов " + CStr(rowCount) + "..."
      rowCount = rowCount + 1
      DoEventsInterval (rowCount)
  
    Loop
    
    
    ' 3. ПВО =================================================================================================
    '  Показатели ПВО
    продукты_контроля_ПВО_РОО = "в т.ч. ПК DSA,            КК DSA"
    
    rowCount = rowByValue(ThisWorkbook.Name, "Лист8", "Итого по РОО «Тюменский»", 500, 500)
    Do While ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value <> ""
    
      ' Если начинается раздел офиса, то Строка_к_выводу обнуляем
      If (InStr(ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value, "Итого по РОО «Тюменский»") <> 0) Then
        
        ' Начало блока офиса
        Начало_блока_офиса = True
        
        ' Имя офиса
        Имя_офиса = "ПВО"
        Имя_офиса_в_Строке_к_выводу = "ПВО"

        ' Счетчик красных зон офиса
        Счетчик_красных_зон_офиса = 0
        
      End If
    
      ' ПВО: Если прогноз по кварталу есть
      If (Начало_блока_офиса = True) And (ThisWorkbook.Sheets("Лист8").Cells(rowCount, 8).Value < 1) And (InStr(продукты_контроля_ПВО_РОО, ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value) <> 0) Then
      
        Строка_к_выводу = Строка_к_выводу + Сокр(ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value) + " (" + CStr(Round(ThisWorkbook.Sheets("Лист8").Cells(rowCount, 8).Value * 100, 0)) + "%), "
        Счетчик_красных_зон_офиса = Счетчик_красных_зон_офиса + 1
      
        ' Строка поручений
        Строка_поручений = Строка_поручений + Сокр(ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value) + " (+" + CStr(Round(ThisWorkbook.Sheets("Лист8").Cells(rowCount, 13).Value, 0)) + ")___, "
      
      End If
    
      ' ПВО Если прогноза по кварталу нет - берем из 20-ой колонки Лист8 прогноз
      If (Начало_блока_офиса = True) And (ThisWorkbook.Sheets("Лист8").Cells(rowCount, 20).Value < 1) And (InStr(продукты_контроля_ПВО_РОО, ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value) <> 0) And (ThisWorkbook.Sheets("Лист8").Cells(rowCount, 8).Value = "") Then
        
        Строка_к_выводу = Строка_к_выводу + Сокр(ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value) + " (" + CStr(Round(ThisWorkbook.Sheets("Лист8").Cells(rowCount, 20).Value * 100, 0)) + "%), "
        Счетчик_красных_зон_офиса = Счетчик_красных_зон_офиса + 1
      
        ' Строка поручений
        Строка_поручений = Строка_поручений + Сокр(ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value) + " (+" + CStr(Round(ThisWorkbook.Sheets("Лист8").Cells(rowCount, 13).Value, 0)) + ")___, "
      
      End If
    
    
      ' Если блок офиса закончился
      If (Начало_блока_офиса = True) And (ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value = "Ипотека") Then
        
        ' Вставить строку в Повестку собрания
        If Счетчик_красных_зон_офиса <> 0 Then
          
          ' Вставляем вопрос
          Call Вставка_строки_в_Повестку(Имя_офиса, _
                                           Имя_офиса_в_Строке_к_выводу + " прогноз исп. БП " + quarterName2(dateDB_Лист_8) + ": " + Строка_к_выводу + " Проект поручения: " + Строка_поручений, _
                                             "")
          ' Добавлено вопросов в повестку
          Добавлено_вопросов = Добавлено_вопросов + 1
          
        End If
        
        ' Начало блока офиса
        Начало_блока_офиса = False
      
        ' Обнуляем строку
        Строка_к_выводу = ""
        Строка_поручений = ""

      End If
    
      ' Следующая запись
      Application.StatusBar = "Анализ прогнозов " + CStr(rowCount) + "..."
      rowCount = rowCount + 1
      DoEventsInterval (rowCount)
  
    Loop


    ' 4. ОО Тюменский (офис) ============================================================
    
    ' Показатели офисного канала Тюмени
    продукты_контроля_офис_Тюмень = "Потребительские кредиты, в т.ч. КК сеть,            КК OPC, Orange Premium Club, Комиссионный доход, Пассивы, Инвест, Инвест OPC, Инвест Брокер обслуж, Инвест Брокер обслуж OPC"
    
    rowCount = rowByValue(ThisWorkbook.Name, "Лист8", "Тюменский РОО", 100, 100) + 2
    Do While (InStr(ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value, "Интегральный рейтинг по офисам") = 0)
    
      ' Если начинается раздел офиса, то Строка_к_выводу обнуляем
      If (InStr(ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value, "Тюменский") <> 0) Then
        
        ' Начало блока офиса
        Начало_блока_офиса = True
        
        ' Имя офиса
        Имя_офиса = cityOfficeName(ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value)
        Имя_офиса_в_Строке_к_выводу = ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value + " (офис)"

        ' Счетчик красных зон офиса
        Счетчик_красных_зон_офиса = 0
        
      End If
    
      ' Тюмень (офис): Если прогноз по кварталу есть
      If (Начало_блока_офиса = True) And (ThisWorkbook.Sheets("Лист8").Cells(rowCount, 8).Value < 1) And (InStr(продукты_контроля_офис_Тюмень, ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value) <> 0) Then
      
        Строка_к_выводу = Строка_к_выводу + Сокр(ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value) + " (" + CStr(Round(ThisWorkbook.Sheets("Лист8").Cells(rowCount, 8).Value * 100, 0)) + "%), "
        Счетчик_красных_зон_офиса = Счетчик_красных_зон_офиса + 1
      
      End If
    
      ' Тюмень (офис): Если прогноза по кварталу нет (портфели)
      If (Начало_блока_офиса = True) And (InStr(продукты_контроля_офис_Тюмень, ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value) <> 0) And (ThisWorkbook.Sheets("Лист8").Cells(rowCount, 8).Value = "") And (ThisWorkbook.Sheets("Лист8").Cells(rowCount, 7).Value < 1) Then
        
        Строка_к_выводу = Строка_к_выводу + Сокр(ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value) + " (" + CStr(Round(ThisWorkbook.Sheets("Лист8").Cells(rowCount, 7).Value * 100, 0)) + "%), "
        Счетчик_красных_зон_офиса = Счетчик_красных_зон_офиса + 1
      
      End If
    
    
      ' Если блок офиса закончился
      If (Начало_блока_офиса = True) And (ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value = "Ипотека") Then
        
        ' Вставить строку в Повестку собрания
        If Счетчик_красных_зон_офиса <> 0 Then
          
          ' Вставляем вопрос
          Call Вставка_строки_в_Повестку(Имя_офиса, _
                                           Имя_офиса_в_Строке_к_выводу + " прогноз исп. БП " + quarterName2(dateDB_Лист_8) + ": " + Строка_к_выводу, _
                                             "")
          ' Добавлено вопросов в повестку
          Добавлено_вопросов = Добавлено_вопросов + 1
          
        End If
        
        ' Начало блока офиса
        Начало_блока_офиса = False
      
        ' Обнуляем строку
        Строка_к_выводу = ""
        Строка_поручений = ""

      End If
    
      ' Следующая запись
      Application.StatusBar = "Анализ прогнозов " + CStr(rowCount) + "..."
      rowCount = rowCount + 1
      DoEventsInterval (rowCount)
  
    Loop

    ' 5. Иногородние офисы (Сургут, Нижневартовск, Новый Уренгой, Тарко-Сале) все - Обрабатываем прогнозы по месяцу с нулем на Лист8 ================================================================================================
    rowCount = rowByValue(ThisWorkbook.Name, "Лист8", "ОО «Сургутский»", 100, 100) - 1
    Do While (InStr(ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value, "Интегральный рейтинг по офисам") = 0)
    
      ' Если начинается раздел офиса, то Строка_к_выводу обнуляем
      ' If (InStr(ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value, "Тюменский") <> 0) Or (InStr(ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value, "Сургутский") <> 0) Or (InStr(ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value, "Нижневартовский") Or (InStr(ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value, "Новоуренгойский")) Or (InStr(ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value, "Тарко-Сале") <> 0)) Then
      If (InStr(ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value, "Сургутский") <> 0) Or (InStr(ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value, "Нижневартовский") Or (InStr(ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value, "Новоуренгойский")) Or (InStr(ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value, "Тарко-Сале") <> 0)) Then
        
        ' Начало блока офиса
        Начало_блока_офиса = True
        
        ' Имя офиса
        Имя_офиса = cityOfficeName(ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value)
        Имя_офиса_в_Строке_к_выводу = ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value

        ' По ОО "Тюменский" добавляем "(офис)"
        ' If (InStr(ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value, "Тюменский") <> 0) Then
        '   Имя_офиса_в_Строке_к_выводу = Имя_офиса_в_Строке_к_выводу + " (офис)"
        ' End If

        ' Счетчик красных зон офиса
        Счетчик_красных_зон_офиса = 0
        
      End If
    
      ' Если прогноз по кварталу есть
      If (Начало_блока_офиса = True) And (ThisWorkbook.Sheets("Лист8").Cells(rowCount, 8).Value < 1) And ((ThisWorkbook.Sheets("Лист8").Cells(rowCount, 3).Value <> "") Or (Показатель_MBO(ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value) = True)) Then
        Строка_к_выводу = Строка_к_выводу + Сокр(ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value) + " (" + CStr(Round(ThisWorkbook.Sheets("Лист8").Cells(rowCount, 8).Value * 100, 0)) + "%), "
        Счетчик_красных_зон_офиса = Счетчик_красных_зон_офиса + 1
        
        ' Строка поручений
        Строка_поручений = Строка_поручений + Сокр(ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value) + " (+" + CStr(Round(ThisWorkbook.Sheets("Лист8").Cells(rowCount, 13).Value, 0)) + ")___, "
        
      End If
    
      ' Если прогноза по кварталу нет (портфели)
      If (Начало_блока_офиса = True) And ((ThisWorkbook.Sheets("Лист8").Cells(rowCount, 3).Value <> "") Or (Показатель_MBO(ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value) = True)) And (ThisWorkbook.Sheets("Лист8").Cells(rowCount, 8).Value = "") And (ThisWorkbook.Sheets("Лист8").Cells(rowCount, 7).Value < 1) Then
        Строка_к_выводу = Строка_к_выводу + Сокр(ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value) + " (" + CStr(Round(ThisWorkbook.Sheets("Лист8").Cells(rowCount, 7).Value * 100, 0)) + "%), "
        Счетчик_красных_зон_офиса = Счетчик_красных_зон_офиса + 1
          
        ' Строка поручений
        If ThisWorkbook.Sheets("Лист8").Cells(rowCount, 13).Value <> "" Then
          Строка_поручений = Строка_поручений + Сокр(ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value) + " (+" + CStr(Round(ThisWorkbook.Sheets("Лист8").Cells(rowCount, 13).Value, 0)) + ")___, "
        End If
  
      End If
    
    
      ' Если блок офиса закончился
      If (Начало_блока_офиса = True) And (ThisWorkbook.Sheets("Лист8").Cells(rowCount, 2).Value = "Ипотека") Then
        
        ' Вставить строку в Повестку собрания
        If Счетчик_красных_зон_офиса <> 0 Then
          
          ' Вставляем вопрос
          Call Вставка_строки_в_Повестку(Имя_офиса + " (" + CStr(Счетчик_красных_зон_офиса) + ")", _
                                           Имя_офиса_в_Строке_к_выводу + " прогноз исп. БП " + quarterName2(dateDB_Лист_8) + ": " + Строка_к_выводу + " Проект поручения: " + Строка_поручений, _
                                             "")
          ' Добавлено вопросов в повестку
          Добавлено_вопросов = Добавлено_вопросов + 1
          
        End If
        
        ' Начало блока офиса
        Начало_блока_офиса = False
      
        ' Обнуляем строку
        Строка_к_выводу = ""
        Строка_поручений = ""
      
      End If
    
      ' Следующая запись
      Application.StatusBar = "Анализ прогнозов " + CStr(rowCount) + "..."
      rowCount = rowCount + 1
      DoEventsInterval (rowCount)
  
    Loop
  
    ' Итоги
    ' Фокусы_контроля_строка = Фокусы_контроля_строка + Строка_к_выводу + Chr(13)
  
    ' Закрываем таблицу MBO
    ' Закрываем BASE\Sales
    CloseBook ("MBO")
    ' ThisWorkbook.Sheets("Лист8").Activate
  
    Application.StatusBar = ""
    
    MsgBox ("Добавлено " + CStr(Добавлено_вопросов) + " вопросов!")
    
  End If
  
End Sub

' Перемещение вверх по поручениям
Sub moveInListUp2()
Dim Ячейка_Поручения_участникам, Текущий_номер, Текущий_ответственный, Текущий_поручение, Текущий_Срок, Текущий_В_To_Do, Цель_номер, Цель_ответственный, Цель_поручение, Цель_Срок, Цель_В_ToDo As String
Dim НомерСтроки_Повестка_дня, НомерСтолбца_Повестка_дня, Текущий_Row, Текущий_Column As Byte

  ' Определяем, где находится текущая ячейка
  Ячейка_Поручения_участникам = RangeByValue(ThisWorkbook.Name, "ЕСУП", "Поручения_участникам", 100, 100)
  НомерСтроки_Поручения_участникам = ThisWorkbook.Sheets("ЕСУП").Range(Ячейка_Поручения_участникам).Row
  НомерСтолбца_Поручения_участникам = ThisWorkbook.Sheets("ЕСУП").Range(Ячейка_Поручения_участникам).Column
  '
  If (ActiveCell.Row >= НомерСтроки_Поручения_участникам + 2) And (ActiveCell.Row <= НомерСтроки_Поручения_участникам + 52) And (ActiveCell.Column >= НомерСтолбца_Поручения_участникам - 1) And ((ActiveCell.Column <= НомерСтолбца_Поручения_участникам + 13)) Then
      ' Координаты
      Текущий_Row = ActiveCell.Row
      Текущий_Column = ActiveCell.Column
      ' Запоминаем текущий
      Текущий_номер = ThisWorkbook.Sheets("ЕСУП").Cells(Текущий_Row, НомерСтолбца_Поручения_участникам - 1).Value
      Текущий_ответственный = ThisWorkbook.Sheets("ЕСУП").Cells(Текущий_Row, НомерСтолбца_Поручения_участникам).Value
      Текущий_поручение = ThisWorkbook.Sheets("ЕСУП").Cells(Текущий_Row, НомерСтолбца_Поручения_участникам + 2).Value
      Текущий_Срок = ThisWorkbook.Sheets("ЕСУП").Cells(Текущий_Row, НомерСтолбца_Поручения_участникам + 1).Value
      Текущий_В_To_Do = ThisWorkbook.Sheets("ЕСУП").Cells(Текущий_Row, НомерСтолбца_Поручения_участникам + 13).Value
      ' Запоминаем цель
      ' Цель_Row = Текущий_Row + 1
      Цель_номер = ThisWorkbook.Sheets("ЕСУП").Cells(Текущий_Row - 1, НомерСтолбца_Поручения_участникам - 1).Value
      Цель_ответственный = ThisWorkbook.Sheets("ЕСУП").Cells(Текущий_Row - 1, НомерСтолбца_Поручения_участникам).Value
      Цель_поручение = ThisWorkbook.Sheets("ЕСУП").Cells(Текущий_Row - 1, НомерСтолбца_Поручения_участникам + 2).Value
      Цель_Срок = ThisWorkbook.Sheets("ЕСУП").Cells(Текущий_Row - 1, НомерСтолбца_Поручения_участникам + 1).Value
      Цель_В_ToDo = ThisWorkbook.Sheets("ЕСУП").Cells(Текущий_Row - 1, НомерСтолбца_Поручения_участникам + 13).Value
      ' Меняем местами:
      ' Текущий ставим в Цель:
      ThisWorkbook.Sheets("ЕСУП").Cells(Текущий_Row - 1, НомерСтолбца_Поручения_участникам - 1).Value = Текущий_номер
      ThisWorkbook.Sheets("ЕСУП").Cells(Текущий_Row - 1, НомерСтолбца_Поручения_участникам).Value = Текущий_ответственный
      ThisWorkbook.Sheets("ЕСУП").Cells(Текущий_Row - 1, НомерСтолбца_Поручения_участникам + 2).Value = Текущий_поручение
      ThisWorkbook.Sheets("ЕСУП").Cells(Текущий_Row - 1, НомерСтолбца_Поручения_участникам + 1).Value = Текущий_Срок
      ThisWorkbook.Sheets("ЕСУП").Cells(Текущий_Row - 1, НомерСтолбца_Поручения_участникам + 13).Value = Текущий_В_To_Do
      ' Цель ставим в Текущий:
      ThisWorkbook.Sheets("ЕСУП").Cells(Текущий_Row, НомерСтолбца_Поручения_участникам - 1).Value = Цель_номер
      ThisWorkbook.Sheets("ЕСУП").Cells(Текущий_Row, НомерСтолбца_Поручения_участникам).Value = Цель_ответственный
      ThisWorkbook.Sheets("ЕСУП").Cells(Текущий_Row, НомерСтолбца_Поручения_участникам + 2).Value = Цель_поручение
      ThisWorkbook.Sheets("ЕСУП").Cells(Текущий_Row, НомерСтолбца_Поручения_участникам + 1).Value = Цель_Срок
      ThisWorkbook.Sheets("ЕСУП").Cells(Текущий_Row, НомерСтолбца_Поручения_участникам + 13).Value = Цель_В_ToDo
      ' Устанавливаем на строку выше
      ThisWorkbook.Sheets("ЕСУП").Cells(Текущий_Row - 1, Текущий_Column).Select
      ' Производим перенумерацию списка
      Call createNumberingTask
    Else
      MsgBox ("Укажите ячейку в диапазоне Поручения_участникам!")
  End If
End Sub

' Перемещение пункта в повестке собрания вниз
Sub moveInListDown2()
Dim Ячейка_Поручения_участникам, Текущий_номер, Текущий_ответственный, Текущий_поручение, Текущий_Срок, Текущий_в_ToDo, Цель_номер, Цель_выступающий, Цель_тема, Цель_HashTag, Цель_Отметка As String
Dim НомерСтроки_Поручения_участникам, НомерСтолбца_Поручения_участникам, Текущий_Row, Текущий_Column As Byte

  ' Определяем, где находится текущая ячейка. Должен быть диапазон A62:N90 (в относительных от "Повестка_дня" координатах)
  Ячейка_Поручения_участникам = RangeByValue(ThisWorkbook.Name, "ЕСУП", "Поручения_участникам", 100, 100)
  НомерСтроки_Поручения_участникам = ThisWorkbook.Sheets("ЕСУП").Range(Ячейка_Поручения_участникам).Row
  НомерСтолбца_Поручения_участникам = ThisWorkbook.Sheets("ЕСУП").Range(Ячейка_Поручения_участникам).Column
  '
  If (ActiveCell.Row >= НомерСтроки_Поручения_участникам + 2) And (ActiveCell.Row <= НомерСтроки_Поручения_участникам + 52) And (ActiveCell.Column >= НомерСтолбца_Поручения_участникам - 1) And ((ActiveCell.Column <= НомерСтолбца_Поручения_участникам + 13)) Then
      ' Координаты
      Текущий_Row = ActiveCell.Row
      Текущий_Column = ActiveCell.Column
      ' Запоминаем текущий
      Текущий_номер = ThisWorkbook.Sheets("ЕСУП").Cells(Текущий_Row, НомерСтолбца_Поручения_участникам - 1).Value
      Текущий_ответственный = ThisWorkbook.Sheets("ЕСУП").Cells(Текущий_Row, НомерСтолбца_Поручения_участникам).Value
      Текущий_поручение = ThisWorkbook.Sheets("ЕСУП").Cells(Текущий_Row, НомерСтолбца_Поручения_участникам + 2).Value
      Текущий_Срок = ThisWorkbook.Sheets("ЕСУП").Cells(Текущий_Row, НомерСтолбца_Поручения_участникам + 1).Value
      Текущий_в_ToDo = ThisWorkbook.Sheets("ЕСУП").Cells(Текущий_Row, НомерСтолбца_Поручения_участникам + 13).Value
      ' Запоминаем цель
      ' Цель_Row = Текущий_Row + 1
      Цель_номер = ThisWorkbook.Sheets("ЕСУП").Cells(Текущий_Row + 1, НомерСтолбца_Поручения_участникам - 1).Value
      Цель_выступающий = ThisWorkbook.Sheets("ЕСУП").Cells(Текущий_Row + 1, НомерСтолбца_Поручения_участникам).Value
      Цель_тема = ThisWorkbook.Sheets("ЕСУП").Cells(Текущий_Row + 1, НомерСтолбца_Поручения_участникам + 2).Value
      Цель_HashTag = ThisWorkbook.Sheets("ЕСУП").Cells(Текущий_Row + 1, НомерСтолбца_Поручения_участникам + 1).Value
      Цель_Отметка = ThisWorkbook.Sheets("ЕСУП").Cells(Текущий_Row + 1, НомерСтолбца_Поручения_участникам + 13).Value
      ' Меняем местами:
      ' Текущий ставим в Цель:
      ThisWorkbook.Sheets("ЕСУП").Cells(Текущий_Row + 1, НомерСтолбца_Поручения_участникам - 1).Value = Текущий_номер
      ThisWorkbook.Sheets("ЕСУП").Cells(Текущий_Row + 1, НомерСтолбца_Поручения_участникам).Value = Текущий_ответственный
      ThisWorkbook.Sheets("ЕСУП").Cells(Текущий_Row + 1, НомерСтолбца_Поручения_участникам + 2).Value = Текущий_поручение
      ThisWorkbook.Sheets("ЕСУП").Cells(Текущий_Row + 1, НомерСтолбца_Поручения_участникам + 1).Value = Текущий_Срок
      ThisWorkbook.Sheets("ЕСУП").Cells(Текущий_Row + 1, НомерСтолбца_Поручения_участникам + 13).Value = Текущий_в_ToDo
      ' Цель ставим в Текущий:
      ThisWorkbook.Sheets("ЕСУП").Cells(Текущий_Row, НомерСтолбца_Поручения_участникам - 1).Value = Цель_номер
      ThisWorkbook.Sheets("ЕСУП").Cells(Текущий_Row, НомерСтолбца_Поручения_участникам).Value = Цель_выступающий
      ThisWorkbook.Sheets("ЕСУП").Cells(Текущий_Row, НомерСтолбца_Поручения_участникам + 2).Value = Цель_тема
      ThisWorkbook.Sheets("ЕСУП").Cells(Текущий_Row, НомерСтолбца_Поручения_участникам + 1).Value = Цель_HashTag
      ThisWorkbook.Sheets("ЕСУП").Cells(Текущий_Row, НомерСтолбца_Поручения_участникам + 13).Value = Цель_Отметка
      ' Устанавливаем на строку выше
      ThisWorkbook.Sheets("ЕСУП").Cells(Текущий_Row + 1, Текущий_Column).Select
      ' Производим перенумерацию списка
      Call createNumberingTask
    Else
      MsgBox ("Укажите ячейку в диапазоне Поручения_участникам!")
  End If

End Sub

' Удаление пункта из поручения участников
Sub deleteFromList2()
Dim Ячейка_Поручения_участникам As String
Dim НомерСтроки_Поручения_участникам, НомерСтолбца_Поручения_участникам As Byte

  ' Определяем, где находится текущая ячейка. Должен быть диапазон A62:N90 (в относительных от "Повестка_дня" координатах)
  Ячейка_Поручения_участникам = RangeByValue(ThisWorkbook.Name, "ЕСУП", "Поручения_участникам", 100, 100)
  НомерСтроки_Поручения_участникам = ThisWorkbook.Sheets("ЕСУП").Range(Ячейка_Поручения_участникам).Row
  НомерСтолбца_Поручения_участникам = ThisWorkbook.Sheets("ЕСУП").Range(Ячейка_Поручения_участникам).Column
  '
  If (ActiveCell.Row >= НомерСтроки_Поручения_участникам + 2) And (ActiveCell.Row <= НомерСтроки_Поручения_участникам + 52) And (ActiveCell.Column >= НомерСтолбца_Поручения_участникам - 1) And ((ActiveCell.Column <= НомерСтолбца_Поручения_участникам + 13)) And (ThisWorkbook.Sheets("ЕСУП").Cells(ActiveCell.Row, НомерСтолбца_Поручения_участникам + 1).Value <> "") Then
    '
    If MsgBox("Удалить вопрос №" + CStr(ThisWorkbook.Sheets("ЕСУП").Cells(ActiveCell.Row, НомерСтолбца_Поручения_участникам - 1).Value) + " из повестки?", vbYesNo) = vbYes Then
      ' Удаляем
      ThisWorkbook.Sheets("ЕСУП").Cells(ActiveCell.Row, НомерСтолбца_Поручения_участникам - 1).Value = ""
      ThisWorkbook.Sheets("ЕСУП").Cells(ActiveCell.Row, НомерСтолбца_Поручения_участникам).Value = ""
      ThisWorkbook.Sheets("ЕСУП").Cells(ActiveCell.Row, НомерСтолбца_Поручения_участникам + 2).Value = ""
      ThisWorkbook.Sheets("ЕСУП").Cells(ActiveCell.Row, НомерСтолбца_Поручения_участникам + 1).Value = ""
      ThisWorkbook.Sheets("ЕСУП").Cells(ActiveCell.Row, НомерСтолбца_Поручения_участникам + 13).Value = ""
    End If
  Else
    MsgBox ("Укажите ячейку в диапазоне Поручения_участникам!")
  End If
  
End Sub

' Получение Даты из номера протокола NN-ДДММГГГГ
Function getDateFromProtocolNumber(In_ProtocolNumber) As Date
  
  Позиция_тире = InStr(In_ProtocolNumber, "-")

  getDateFromProtocolNumber = CDate(Mid(In_ProtocolNumber, Позиция_тире + 1, 2) + "." + Mid(In_ProtocolNumber, Позиция_тире + 3, 2) + "." + Mid(In_ProtocolNumber, Позиция_тире + 5, 4))
  
End Function


' Копирование поручений в To-Do
Sub copyTaskInToDo()
  
  ' Запрос
  If MsgBox("Скопировать поручения в To-Do?", vbYesNo) = vbYes Then
    
    ' Открываем таблицу BASE\ToDo
    OpenBookInBase ("ToDo")

    ' Переходим на окно DB
    ThisWorkbook.Sheets("ЕСУП").Activate

    ' Строка статуса
    Application.StatusBar = "Копирование..."

    ' На всякий случай нумеруем список
    ' Call createNumberingTask

    НомерСтроки_Поручения_участникам = ThisWorkbook.Sheets("ЕСУП").Range(RangeByValue(ThisWorkbook.Name, "ЕСУП", "Поручения_участникам", 100, 100)).Row
    НомерСтолбца_Поручения_участникам = ThisWorkbook.Sheets("ЕСУП").Range(RangeByValue(ThisWorkbook.Name, "ЕСУП", "Поручения_участникам", 100, 100)).Column
    
    rowCount = 2
    Do While ThisWorkbook.Sheets("ЕСУП").Cells(НомерСтроки_Поручения_участникам + rowCount, 19).Value <> ""
      
      ' Id_Task из G2 + номер вопроса
      Id_TaskVar = Replace(ThisWorkbook.Sheets("ЕСУП").Range("G2").Value, "-", "") + CStr(ThisWorkbook.Sheets("ЕСУП").Cells(НомерСтроки_Поручения_участникам + rowCount, НомерСтолбца_Поручения_участникам - 1).Value)
      
      ' Вносим данные в BASE\Sales по ПК.
      Call InsertRecordInBook("ToDo", "Лист1", "Id_Task", Id_TaskVar, _
                                            "Date_Create", getDateFromProtocolNumber(ThisWorkbook.Sheets("ЕСУП").Range("G2").Value), _
                                              "Id_Task", Id_TaskVar, _
                                                "Task", ThisWorkbook.Sheets("ЕСУП").Cells(НомерСтроки_Поручения_участникам + rowCount, НомерСтолбца_Поручения_участникам + 2).Value, _
                                                  "Lotus_subject", ThisWorkbook.Sheets("ЕСУП").Cells(1, 17).Value, _
                                                    "Responsible", ThisWorkbook.Sheets("ЕСУП").Cells(НомерСтроки_Поручения_участникам + rowCount, НомерСтолбца_Поручения_участникам).Value, _
                                                      "Lotus_hashtag", "#" + Id_TaskVar, _
                                                        "Task_Status", 1, _
                                                          "Date_Control", ThisWorkbook.Sheets("ЕСУП").Cells(НомерСтроки_Поручения_участникам + rowCount, НомерСтолбца_Поручения_участникам + 1).Value, _
                                                            "Comment", "В работе", _
                                                              "Protocol_Number", ThisWorkbook.Sheets("ЕСУП").Range("G2").Value, _
                                                                "Protocol_Date", getDateFromProtocolNumber(ThisWorkbook.Sheets("ЕСУП").Range("G2").Value), _
                                                                  "Protocol_Question_number", ThisWorkbook.Sheets("ЕСУП").Cells(НомерСтроки_Поручения_участникам + rowCount, НомерСтолбца_Поручения_участникам - 1).Value, _
                                                                    "", "", _
                                                                      "", "", _
                                                                        "", "", _
                                                                          "", "", _
                                                                            "", "", _
                                                                              "", "", _
                                                                                "", "", _
                                                                                  "", "")
      ' Отмечаем
      ThisWorkbook.Sheets("ЕСУП").Cells(НомерСтроки_Поручения_участникам + rowCount, НомерСтолбца_Поручения_участникам + 13).Value = 1
      
      ' Следующая строка
      rowCount = rowCount + 1
    Loop


    
    Application.StatusBar = "Завершение..."
    
    ' Закрываем таблицу BASE\ToDo
    CloseBook ("ToDo")
 
    ' Строка статуса
    Application.StatusBar = ""

    MsgBox ("Перенесено " + CStr(rowCount - 2) + " вопросов!")
    
  End If
  
End Sub

' Добавить ответственного
Sub addResponsibleSheet8(In_К_дол)
  ' Определяем, где находится текущая ячейка. Должен быть диапазон A62:N90 (в относительных от "Повестка_дня" координатах)
  Ячейка_Поручения_участникам = RangeByValue(ThisWorkbook.Name, "ЕСУП", "Поручения_участникам", 100, 100)
  НомерСтроки_Поручения_участникам = ThisWorkbook.Sheets("ЕСУП").Range(Ячейка_Поручения_участникам).Row
  НомерСтолбца_Поручения_участникам = ThisWorkbook.Sheets("ЕСУП").Range(Ячейка_Поручения_участникам).Column
  ' Проверяем активную ячейку
  If (ActiveCell.Row >= НомерСтроки_Поручения_участникам + 2) And (ActiveCell.Row <= НомерСтроки_Поручения_участникам + 52) And (ActiveCell.Column >= НомерСтолбца_Поручения_участникам - 1) And ((ActiveCell.Column <= НомерСтолбца_Поручения_участникам + 13)) Then
    ' Ответственный:
    ThisWorkbook.Sheets("ЕСУП").Cells(ActiveCell.Row, НомерСтолбца_Поручения_участникам).Value = getFromAddrBook(In_К_дол, 3)
  Else
    MsgBox ("Укажите ячейку в диапазоне Поручения_участникам!")
  End If
End Sub

' Добавить выступающего
Sub addSpeakerSheet8(In_К_дол)
  ' Определяем, где находится текущая ячейка. Должен быть диапазон A62:N90 (в относительных от "Повестка_дня" координатах)
  Ячейка_Повестка_дня = RangeByValue(ThisWorkbook.Name, "ЕСУП", "Повестка_дня", 100, 100)
  НомерСтроки_Повестка_дня = ThisWorkbook.Sheets("ЕСУП").Range(Ячейка_Повестка_дня).Row
  НомерСтолбца_Повестка_дня = ThisWorkbook.Sheets("ЕСУП").Range(Ячейка_Повестка_дня).Column
  
  ' Проверяем активную ячейку
  If (ActiveCell.Row >= НомерСтроки_Повестка_дня + 2) And (ActiveCell.Row <= НомерСтроки_Повестка_дня + 52) And (ActiveCell.Column >= НомерСтолбца_Повестка_дня - 1) And ((ActiveCell.Column <= НомерСтолбца_Повестка_дня + 13)) Then
    ' Ответственный:
    ThisWorkbook.Sheets("ЕСУП").Cells(ActiveCell.Row, НомерСтолбца_Повестка_дня).Value = getFromAddrBook(In_К_дол, 3)
  Else
    MsgBox ("Укажите ячейку в диапазоне Повестка_дня!")
  End If
  
End Sub


' Добавить ответственного РИЦ
Sub addResponsibleSheet8_РИЦ()
  Call addResponsibleSheet8("РИЦ")
End Sub


' Добавить ответственного НОКП
Sub addResponsibleSheet8_НОКП()
  Call addResponsibleSheet8("НОКП")
End Sub

' Добавить ответственного НОРПиКО1
Sub addResponsibleSheet8_НОРПиКО1()
  Call addResponsibleSheet8("НОРПиКО1")
End Sub

' Добавить ответственного УДО2
Sub addResponsibleSheet8_УДО2()
  Call addResponsibleSheet8("УДО2")
End Sub

' Добавить ответственного НОРПиКО2
Sub addResponsibleSheet8_НОРПиКО2()
  Call addResponsibleSheet8("НОРПиКО2")
End Sub

' Добавить ответственного УДО3
Sub addResponsibleSheet8_УДО3()
  Call addResponsibleSheet8("УДО3")
End Sub

' Добавить ответственного НОРПиКО3
Sub addResponsibleSheet8_НОРПиКО3()
  Call addResponsibleSheet8("НОРПиКО3")
End Sub

' Добавить ответственного УДО4
Sub addResponsibleSheet8_УДО4()
  Call addResponsibleSheet8("УДО4")
End Sub

' Добавить ответственного НОРПиКО4
Sub addResponsibleSheet8_НОРПиКО4()
  Call addResponsibleSheet8("НОРПиКО4")
End Sub

' Добавить ответственного УДО5
Sub addResponsibleSheet8_УДО5()
  Call addResponsibleSheet8("УДО5")
End Sub

' Добавить ответственного НОРПиКО5
Sub addResponsibleSheet8_НОРПиКО5()
  Call addResponsibleSheet8("НОРПиКО5")
End Sub

' Добавить срок исполнения
Sub addDateEndSheet8(In_DateEnd)
  ' Определяем, где находится текущая ячейка. Должен быть диапазон A62:N90 (в относительных от "Повестка_дня" координатах)
  Ячейка_Поручения_участникам = RangeByValue(ThisWorkbook.Name, "ЕСУП", "Поручения_участникам", 100, 100)
  НомерСтроки_Поручения_участникам = ThisWorkbook.Sheets("ЕСУП").Range(Ячейка_Поручения_участникам).Row
  НомерСтолбца_Поручения_участникам = ThisWorkbook.Sheets("ЕСУП").Range(Ячейка_Поручения_участникам).Column
  ' Проверяем активную ячейку
  If (ActiveCell.Row >= НомерСтроки_Поручения_участникам + 2) And (ActiveCell.Row <= НомерСтроки_Поручения_участникам + 52) And (ActiveCell.Column >= НомерСтолбца_Поручения_участникам - 1) And ((ActiveCell.Column <= НомерСтолбца_Поручения_участникам + 13)) Then
    ' Ответственный:
    ThisWorkbook.Sheets("ЕСУП").Cells(ActiveCell.Row, НомерСтолбца_Поручения_участникам + 1).Value = In_DateEnd
  Else
    MsgBox ("Укажите ячейку в диапазоне Поручения_участникам!")
  End If
End Sub

' Добавить срок исполнения - до пятницы
Sub addDateEndSheet8_Пятница()
  Call addDateEndSheet8(weekEndDate(Date) - 2)
End Sub

' Добавить срок исполнения - до конца месяца
Sub addDateEndSheet8_КонецМесяца()
  Call addDateEndSheet8(Date_last_day_month(Date))
End Sub

' Добавить срок исполнения - до конца квартала
Sub addDateEndSheet8_КонецКвартала()
  Call addDateEndSheet8(Date_last_day_quarter(Date))
End Sub

' Вставить строку в Поручения_участникам
Sub Вставка_строки_в_Поручения_участникам(In_FIO, In_DateEnd, In_Task)
Dim НомерСтрокиЛист_ЕСУП, НомерСтолбцаЛист_ЕСУП As Byte

    
    НомерСтрокиЛист_ЕСУП = rowByValue(ThisWorkbook.Name, "ЕСУП", "Поручения_участникам", 100, 100)
    НомерСтолбцаЛист_ЕСУП = ColumnByValue(ThisWorkbook.Name, "ЕСУП", "Поручения_участникам", 100, 100)
    
    ' Заносим на Лист "ЕСУП"
    i = НомерСтрокиЛист_ЕСУП + 2
    Номер_поручения = 0
    Do While ThisWorkbook.Sheets("ЕСУП").Cells(i, НомерСтолбцаЛист_ЕСУП + 2).Value <> ""
      Номер_поручения = Номер_поручения + 1
      i = i + 1
    Loop
    
    Номер_поручения = Номер_поручения + 1
    
    ' № поручения
    ThisWorkbook.Sheets("ЕСУП").Cells(i, НомерСтолбцаЛист_ЕСУП - 1).Value = Номер_поручения
    ' Ответственный
    ThisWorkbook.Sheets("ЕСУП").Cells(i, НомерСтолбцаЛист_ЕСУП).Value = In_FIO
    ' Дата контроля
    ThisWorkbook.Sheets("ЕСУП").Cells(i, НомерСтолбцаЛист_ЕСУП + 1).Value = In_DateEnd
    ' Поручение
    ThisWorkbook.Sheets("ЕСУП").Cells(i, НомерСтолбцаЛист_ЕСУП + 2).Value = In_Task
    ' Скопировано в To-Do 0/1
    ThisWorkbook.Sheets("ЕСУП").Cells(i, НомерСтолбцаЛист_ЕСУП + 13).Value = 0

End Sub


' Добавить выступающего НОКП
Sub addSpeakerSheet8_НОКП()
  Call addSpeakerSheet8("НОКП")
End Sub

' Добавить выступающего РИЦ
Sub addSpeakerSheet8_РИЦ()
  Call addSpeakerSheet8("РИЦ")
End Sub

' Добавить выступающего НОРПиКО1
Sub addSpeakerSheet8_НОРПиКО1()
  Call addSpeakerSheet8("НОРПиКО1")
End Sub

' Добавить выступающего УДО2
Sub addSpeakerSheet8_УДО2()
  Call addSpeakerSheet8("УДО2")
End Sub

' Добавить выступающего НОРПиКО2
Sub addSpeakerSheet8_НОРПиКО2()
  Call addSpeakerSheet8("НОРПиКО2")
End Sub

' Добавить выступающего УДО3
Sub addSpeakerSheet8_УДО3()
  Call addSpeakerSheet8("УДО3")
End Sub

' Добавить выступающего НОРПиКО3
Sub addSpeakerSheet8_НОРПиКО3()
  Call addSpeakerSheet8("НОРПиКО3")
End Sub

' Добавить выступающего УДО4
Sub addSpeakerSheet8_УДО4()
  Call addSpeakerSheet8("УДО4")
End Sub

' Добавить выступающего НОРПиКО4
Sub addSpeakerSheet8_НОРПиКО4()
  Call addSpeakerSheet8("НОРПиКО4")
End Sub

' Добавить выступающего УДО5
Sub addSpeakerSheet8_УДО5()
  Call addSpeakerSheet8("УДО5")
End Sub

' Добавить выступающего НОРПиКО5
Sub addSpeakerSheet8_НОРПиКО5()
  Call addSpeakerSheet8("НОРПиКО5")
End Sub