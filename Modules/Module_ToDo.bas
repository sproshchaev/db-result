Attribute VB_Name = "Module_ToDo"
' Лист To-Do

' Очистить лист и из таблицы ToDo
Sub ToDo_refresh()
Attribute ToDo_refresh.VB_ProcData.VB_Invoke_Func = " \n14"

  ' 1. Очищаем диапазон ячеек на листе
  Call clearСontents2(ThisWorkbook.Name, "To-Do", "A6", "L1000")

  ' 2. Открываем таблицу BASE\ToDo
  OpenBookInBase ("ToDo")

  ' Переходим на окно DB
  ThisWorkbook.Sheets("To-Do").Activate

  ' Строка статуса
  Application.StatusBar = "Обработка..."

  ' Строка поиска
  ThisWorkbook.Sheets("To-Do").Range("E1").Value = ""
  
  ' Заголовок
  ThisWorkbook.Sheets("To-Do").Range("B2").Value = "Поручения и вопросы на контроле на " + CStr(Date)

  ' Номер по порядку
  НомерПоПорядку = 0

  ' Выбираем все задачи со сроком, который наступил - Date_Control
  rowCount = 2
  rowCount2 = 5
  
  Do While Not IsEmpty(Workbooks("ToDo").Sheets("Лист1").Cells(rowCount, 1).Value)
        
    ' Если наступила дата и есть поручения со статусом = 1 (в работе)
    If (CDate(Workbooks("ToDo").Sheets("Лист1").Cells(rowCount, 8).Value) <= Date) And (Workbooks("ToDo").Sheets("Лист1").Cells(rowCount, 7).Value = 1) Then
              
      ' Инициализация переменной
      Выводить_запись = False
      
      ' Если это поручение из протокола и стоит опция Протоколы = 1
      If (Workbooks("ToDo").Sheets("Лист1").Cells(rowCount, 10).Value <> "") And (ThisWorkbook.Sheets("To-Do").Range("I1").Value = 1) Then
        Выводить_запись = True
      End If
      
      ' Если это поручение из протокола и стоит опция Протоколы = 0
      If (Workbooks("ToDo").Sheets("Лист1").Cells(rowCount, 10).Value <> "") And (ThisWorkbook.Sheets("To-Do").Range("I1").Value = 0) Then
        Выводить_запись = False
      End If
      
      ' Если это поручение не из протокола
      If (Workbooks("ToDo").Sheets("Лист1").Cells(rowCount, 10).Value = "") Then
        Выводить_запись = True
      End If
      
      ' Если это поручение не из протокола
      If Выводить_запись = True Then
      
        ' Счетчик выводимых строк на листе "To-Do"
        rowCount2 = rowCount2 + 1
      
        ' Номер по порядку
        НомерПоПорядку = НомерПоПорядку + 1
      
        ' Выводим на Лист "To-Do"
        ' №
        ThisWorkbook.Sheets("To-Do").Cells(rowCount2, 1).Value = CStr(НомерПоПорядку) + "."
        ThisWorkbook.Sheets("To-Do").Cells(rowCount2, 1).HorizontalAlignment = xlCenter
      
        ' Дата создания
        ThisWorkbook.Sheets("To-Do").Cells(rowCount2, 2).Value = Workbooks("ToDo").Sheets("Лист1").Cells(rowCount, 1).Value
        ThisWorkbook.Sheets("To-Do").Cells(rowCount2, 2).HorizontalAlignment = xlCenter
      
        ' Id_Task
        ThisWorkbook.Sheets("To-Do").Cells(rowCount2, 3).Value = Workbooks("ToDo").Sheets("Лист1").Cells(rowCount, 2).Value
        ThisWorkbook.Sheets("To-Do").Cells(rowCount2, 3).HorizontalAlignment = xlCenter
            
        ' Task
        ThisWorkbook.Sheets("To-Do").Cells(rowCount2, 4).Value = Workbooks("ToDo").Sheets("Лист1").Cells(rowCount, 3).Value
      
        ' Тема лотус
        ThisWorkbook.Sheets("To-Do").Cells(rowCount2, 5).Value = Workbooks("ToDo").Sheets("Лист1").Cells(rowCount, 4).Value
      
        ' 6 HashTag (если не заполнено, то ставим "")
        If IsEmpty(Workbooks("ToDo").Sheets("Лист1").Cells(rowCount, 6).Value) Then
          ThisWorkbook.Sheets("To-Do").Cells(rowCount2, 6).Value = " "
        Else
          ThisWorkbook.Sheets("To-Do").Cells(rowCount2, 6).Value = Workbooks("ToDo").Sheets("Лист1").Cells(rowCount, 6).Value
        End If
        
        ' 7 Статус задачи
        If Workbooks("ToDo").Sheets("Лист1").Cells(rowCount, 7).Value = 1 Then
          ThisWorkbook.Sheets("To-Do").Cells(rowCount2, 7).Value = "В работе"
        End If
        '
        If Workbooks("ToDo").Sheets("Лист1").Cells(rowCount, 7).Value = 0 Then
          ThisWorkbook.Sheets("To-Do").Cells(rowCount2, 7).Value = "Закрыта"
        End If

        ThisWorkbook.Sheets("To-Do").Cells(rowCount2, 7).HorizontalAlignment = xlCenter
        
        ' 8 Ответственный
        ThisWorkbook.Sheets("To-Do").Cells(rowCount2, 8).Value = Workbooks("ToDo").Sheets("Лист1").Cells(rowCount, 5).Value
        ThisWorkbook.Sheets("To-Do").Cells(rowCount2, 8).HorizontalAlignment = xlCenter
      
        ' 9 Дата контроля
        ThisWorkbook.Sheets("To-Do").Cells(rowCount2, 9).Value = Workbooks("ToDo").Sheets("Лист1").Cells(rowCount, 8).Value
        ThisWorkbook.Sheets("To-Do").Cells(rowCount2, 9).HorizontalAlignment = xlCenter
      
        ' 10 Комментарий
        ThisWorkbook.Sheets("To-Do").Cells(rowCount2, 10).Value = Workbooks("ToDo").Sheets("Лист1").Cells(rowCount, 9).Value
      
        ' 11 Протокол Собрания (Protocol_Number)
        ThisWorkbook.Sheets("To-Do").Cells(rowCount2, 11).Value = Workbooks("ToDo").Sheets("Лист1").Cells(rowCount, 10).Value
      
        ' 12 Вставить в повестку Placed_Agenda
        ThisWorkbook.Sheets("To-Do").Cells(rowCount2, 12).Value = Workbooks("ToDo").Sheets("Лист1").Cells(rowCount, 16).Value
      
      End If ' Выводить поручения из протоколов
      
    End If
                  
    ' Следующая запись
    rowCount = rowCount + 1
    ' Application.StatusBar = officeNameInReport + ": " + CStr(rowCount) + "..."
    ' DoEventsInterval (rowCount)
          
  Loop
  
  
  ' Закрываем таблицу BASE\ToDo
  CloseBook ("ToDo")
 
  ' Сохранение изменений
  ThisWorkbook.Save

  ' Строка статуса
  Application.StatusBar = ""
 
 
End Sub

' Сохранить все данные в таблицу ToDo
Sub ToDo_save()
  
  ' Открываем таблицу BASE\ToDo
  OpenBookInBase ("ToDo")

  ' Переходим на окно DB
  ThisWorkbook.Sheets("To-Do").Activate

  ' Строка статуса
  Application.StatusBar = "Сохранение..."

  ' Выбираем все задачи со сроком, который наступил - Date_Control
  rowCount = 6
  Номер_строки = 0
    
  ' Do While Not IsEmpty(ThisWorkbook.Sheets("To-Do").Cells(rowCount, 1).Value)
  Do While ThisWorkbook.Sheets("To-Do").Cells(rowCount, 1).Value <> ""
  
    ' Преобразование статуса задачи: "В работе" -> 1
    If ThisWorkbook.Sheets("To-Do").Cells(rowCount, 7).Value = "В работе" Then
      Task_Status_Var = 1
    Else
      Task_Status_Var = 0
    End If
    
    ' Внести в повестку
    If ThisWorkbook.Sheets("To-Do").Cells(rowCount, 12).Value = 1 Then
                
                
      ' Вставляем вопрос
      Call Вставка_строки_в_Повестку("Прощаев С.Ф.", _
                                       ThisWorkbook.Sheets("To-Do").Cells(rowCount, 4).Value, _
                                         ThisWorkbook.Sheets("To-Do").Cells(rowCount, 6).Value)

      ' Изменяем статус с 1 (вставить в повестку) на 2 (внесено в повестку)
      ThisWorkbook.Sheets("To-Do").Cells(rowCount, 12).Value = 2

      ' Сообщение о внесении в повестку
      MsgBox ("Внесено в повестку предстоящего собрания!(" + ThisWorkbook.Sheets("To-Do").Cells(rowCount, 6).Value + ")")

    End If
    
    ' Вносим данные в BASE\ToDo
    Call InsertRecordInBook("ToDo", "Лист1", "Id_Task", ThisWorkbook.Sheets("To-Do").Cells(rowCount, 3).Value, _
                                            "Date_Create", ThisWorkbook.Sheets("To-Do").Cells(rowCount, 2).Value, _
                                              "Id_Task", ThisWorkbook.Sheets("To-Do").Cells(rowCount, 3).Value, _
                                                "Task", ThisWorkbook.Sheets("To-Do").Cells(rowCount, 4).Value, _
                                                  "Lotus_subject", ThisWorkbook.Sheets("To-Do").Cells(rowCount, 5).Value, _
                                                    "Responsible", ThisWorkbook.Sheets("To-Do").Cells(rowCount, 8).Value, _
                                                      "Lotus_hashtag", ThisWorkbook.Sheets("To-Do").Cells(rowCount, 6).Value, _
                                                        "Task_Status", Task_Status_Var, _
                                                          "Date_Control", ThisWorkbook.Sheets("To-Do").Cells(rowCount, 9).Value, _
                                                            "Comment", ThisWorkbook.Sheets("To-Do").Cells(rowCount, 10).Value, _
                                                              "Protocol_Number", ThisWorkbook.Sheets("To-Do").Cells(rowCount, 11).Value, _
                                                                "Placed_Agenda", ThisWorkbook.Sheets("To-Do").Cells(rowCount, 12).Value, _
                                                                  "", "", _
                                                                    "", "", _
                                                                      "", "", _
                                                                        "", "", _
                                                                          "", "", _
                                                                            "", "", _
                                                                              "", "", _
                                                                                "", "", _
                                                                                  "", "")
          
    ' Следующая запись
    Application.StatusBar = "Сохранение (" + ThisWorkbook.Sheets("To-Do").Cells(rowCount, 1).Value + ") ..."
    rowCount = rowCount + 1
    ' DoEventsInterval (rowCount)
          
  Loop
  
  Application.StatusBar = "Завершение..."
    
  ' Закрываем таблицу BASE\ToDo
  CloseBook ("ToDo")
 
  ' Строка статуса
  Application.StatusBar = ""
  
  ' Запускаем обновление (и в ToDo_refresh есть сохранение)
  Call ToDo_refresh
  
End Sub

' Добавить задачу в таблицу на Листе To-DO
Sub add_Task_ToDo()
  
  ' Проходим до конца таблицы и добавляем новую запись
  rowCount = 6
  Номер_по_списку = 0
  
  ' Do While Not IsEmpty(ThisWorkbook.Sheets("To-Do").Cells(rowCount, 1).Value)
  Do While Len(ThisWorkbook.Sheets("To-Do").Cells(rowCount, 1).Value) <> 0
  
    Номер_по_списку = Номер_по_списку + 1
  
    ' Следующая запись
    Application.StatusBar = CStr(rowCount) + "..."
    rowCount = rowCount + 1
    ' DoEventsInterval (rowCount)
          
  Loop
  
  ' Добавляем новую задачу в конец
  hashTag = createHashTag("t")
  Номер_по_списку = Номер_по_списку + 1
  ThisWorkbook.Sheets("To-Do").Cells(rowCount, 1).Value = CStr(Номер_по_списку) + "."
  ThisWorkbook.Sheets("To-Do").Cells(rowCount, 1).HorizontalAlignment = xlCenter
  
  ' Дата создания
  ThisWorkbook.Sheets("To-Do").Cells(rowCount, 2).Value = Date
  ThisWorkbook.Sheets("To-Do").Cells(rowCount, 2).HorizontalAlignment = xlCenter
      
  ' Id_Task
  ThisWorkbook.Sheets("To-Do").Cells(rowCount, 3).Value = Replace(hashTag, "#t", "")
  ThisWorkbook.Sheets("To-Do").Cells(rowCount, 3).HorizontalAlignment = xlCenter
            
  ' Task
  ThisWorkbook.Sheets("To-Do").Cells(rowCount, 4).Value = ""
      
  ' Тема лотус
  ThisWorkbook.Sheets("To-Do").Cells(rowCount, 5).Value = ""
      
  ' 6 HashTag (если не заполнено, то ставим "")
  ThisWorkbook.Sheets("To-Do").Cells(rowCount, 6).Value = hashTag
  ThisWorkbook.Sheets("To-Do").Cells(rowCount, 6).HorizontalAlignment = xlCenter
  
  ' 7 Статус задачи
  ThisWorkbook.Sheets("To-Do").Cells(rowCount, 7).Value = "В работе"
  ThisWorkbook.Sheets("To-Do").Cells(rowCount, 7).HorizontalAlignment = xlCenter
        
  ' 8 Ответственный
  ThisWorkbook.Sheets("To-Do").Cells(rowCount, 8).Value = ""
  ThisWorkbook.Sheets("To-Do").Cells(rowCount, 8).HorizontalAlignment = xlCenter
      
  ' 9 Дата контроля
  ThisWorkbook.Sheets("To-Do").Cells(rowCount, 9).Value = Date
  ThisWorkbook.Sheets("To-Do").Cells(rowCount, 9).HorizontalAlignment = xlCenter
      
  ' 10 Комментарий
  ThisWorkbook.Sheets("To-Do").Cells(rowCount, 10).Value = ""
 
  ' Сохранение изменений
  ThisWorkbook.Save

End Sub

' Закрыть задачу в таблицу на Листе To-DO
Sub Task_ToClose()
        
  ' Строка
  ThisWorkbook.Sheets("To-Do").Cells(ActiveCell.Row, 7).Value = "Закрыта"

  ' Сохранение изменений
  ThisWorkbook.Save

  ' Запрос на запуск процедуры ToDo_save() сохранения в BASE\To-Do
  If MsgBox("Сохранить изменения в BASE\ToDo?", vbYesNo) = vbYes Then
    Call ToDo_save
  End If


End Sub

' Увеличить срок на In_Day день в текущей задаче в таблицу на Листе To-DO
Sub Task_AddDay(In_Day)
          
  ' Проверяем - стоит ли курсор на строке с задачей
  If Len(ThisWorkbook.Sheets("To-Do").Cells(ActiveCell.Row, 1).Value) <> 0 Then
        
    ' Строка
    ThisWorkbook.Sheets("To-Do").Cells(ActiveCell.Row, 9).Value = CDate(ThisWorkbook.Sheets("To-Do").Cells(ActiveCell.Row, 9).Value) + In_Day

    ' Сохранение изменений
    ThisWorkbook.Save
  Else
    
    ' Сообщение
    MsgBox ("Выберите задачу!")

  End If

End Sub

' Увеличить срок на 1 день в текущей задаче в таблицу на Листе To-DO
Sub Task_Add_1_Day()
          
  Call Task_AddDay(1)

End Sub

' Увеличить срок на 7 дней (неделя) в текущей задаче в таблицу на Листе To-DO
Sub Task_Add_7_Day()
          
  Call Task_AddDay(7)

End Sub

' Увеличить срок на 30 дней (месяц) в текущей задаче в таблицу на Листе To-DO
Sub Task_Add_30_Day()
          
  Call Task_AddDay(30)

End Sub


' Вставить из буфера обмена в столбец Тема данные и скопировать в буффер хэштег
Sub ToDo_InsertTheme()
  
  ' Проверяем - стоит ли курсор на строке с задачей
  If Len(ThisWorkbook.Sheets("To-Do").Cells(ActiveCell.Row, 1).Value) <> 0 Then
        
            
    ' Проверяем наличие информации в буффере обмена! Внимание - выдает ошибку, если в буффере обмена ничего нет
    If Len(ClipboardText()) <> 0 Then
      
      ' Вставляем Тему Lotus Notes из буффера обмена
      ThisWorkbook.Sheets("To-Do").Cells(ActiveCell.Row, 5).Select

      ActiveSheet.Paste
    
      ' Сохранение изменений
      ThisWorkbook.Save
    
      ' Копируем Хэштег в буффер обмена
      ThisWorkbook.Sheets("To-Do").Cells(ActiveCell.Row, 6).Copy

    Else
    
      ' Сообщение
      MsgBox ("Буффер обмена пуст!")
      
    End If
        
    
  Else
    
    ' Сообщение
    MsgBox ("Выберите задачу!")

  End If
  
End Sub

' Скопировать в буфер обмена в столбец Хэштег
Sub ToDo_CopyTag()
  
  ' Проверяем - стоит ли курсор на строке с задачей
  If Len(ThisWorkbook.Sheets("To-Do").Cells(ActiveCell.Row, 1).Value) <> 0 Then
                           
      ' Копируем Хэштег в буффер обмена
      ThisWorkbook.Sheets("To-Do").Cells(ActiveCell.Row, 6).Copy
            
  Else
    
    ' Сообщение
    MsgBox ("Выберите задачу!")

  End If
  
End Sub

' Скопировать в буфер обмена в столбец Тема
Sub ToDo_CopyTheme()
  
  ' Проверяем - стоит ли курсор на строке с задачей
  If Len(ThisWorkbook.Sheets("To-Do").Cells(ActiveCell.Row, 1).Value) <> 0 Then
                           
      ' Копируем Хэштег в буффер обмена
      ThisWorkbook.Sheets("To-Do").Cells(ActiveCell.Row, 5).Copy
            
  Else
    
    ' Сообщение
    MsgBox ("Выберите задачу!")

  End If
  
End Sub

' Добавить комментарий
Sub ToDo_Add_Upd()
  
  ' Проверяем - стоит ли курсор на строке с задачей
  If Len(ThisWorkbook.Sheets("To-Do").Cells(ActiveCell.Row, 1).Value) <> 0 Then
                           
      ' Добавляем конструкцию  =(Upd.ДДММ) в комментарий
      ThisWorkbook.Sheets("To-Do").Cells(ActiveCell.Row, 10).Select
      ThisWorkbook.Sheets("To-Do").Cells(ActiveCell.Row, 10).Value = ThisWorkbook.Sheets("To-Do").Cells(ActiveCell.Row, 10).Value + " =(Upd." + strDDMM(Date) + ") "
            
      ' Сохранение изменений
      ThisWorkbook.Save
      
  Else
    
    ' Сообщение
    MsgBox ("Выберите задачу!")

  End If
  
End Sub

' Найти задачи из BASE\To-Do
Sub ToDo_Find_Tasks()
Dim SearchString As String
Dim find_in_Task, find_in_Lotus_subject, find_in_Responsible, find_in_Lotus_hashtag, find_in_Comment, find_in_Protocol_Number As Byte

  ' 1. Очищаем диапазон ячеек на листе
  Call clearСontents2(ThisWorkbook.Name, "To-Do", "A6", "L100")

  ' 2. Открываем таблицу BASE\ToDo
  OpenBookInBase ("ToDo")

  ' Переходим на окно DB
  ThisWorkbook.Sheets("To-Do").Activate

  ' Строка статуса
  Application.StatusBar = "Обработка..."

  ' Строка поиска
  SearchString = ThisWorkbook.Sheets("To-Do").Range("E1").Value

  ' Заголовок
  ThisWorkbook.Sheets("To-Do").Range("B2").Value = "Поиск: " + SearchString

  ' Номер по порядку
  НомерПоПорядку = 0

  ' Выбираем задачи по SearchString
  rowCount = 2
  rowCount2 = 5
  
  Do While Not IsEmpty(Workbooks("ToDo").Sheets("Лист1").Cells(rowCount, 1).Value)
        
    ' Поиск в Task (3) (Прим.: 1 - поиск без учета регистра символов)
    find_in_Task = InStr(1, Workbooks("ToDo").Sheets("Лист1").Cells(rowCount, 3).Value, SearchString, 1)
    ' Поиск в Lotus_subject (4)
    find_in_Lotus_subject = InStr(1, Workbooks("ToDo").Sheets("Лист1").Cells(rowCount, 4).Value, SearchString, 1)
    ' Поиск в Responsible (5)
    find_in_Responsible = InStr(1, Workbooks("ToDo").Sheets("Лист1").Cells(rowCount, 5).Value, SearchString, 1)
    ' Поиск в Lotus_hashtag (6)
    find_in_Lotus_hashtag = InStr(1, Workbooks("ToDo").Sheets("Лист1").Cells(rowCount, 6).Value, SearchString, 1)
    ' Поиск в Comment (9)
    find_in_Comment = InStr(1, Workbooks("ToDo").Sheets("Лист1").Cells(rowCount, 9).Value, SearchString, 1)
    ' Поиск в Protocol_Number (10)
    find_in_Protocol_Number = InStr(1, Workbooks("ToDo").Sheets("Лист1").Cells(rowCount, 10).Value, SearchString, 1)
        
    ' Если выполняется условие
    If (find_in_Task <> 0) Or (find_in_Lotus_subject <> 0) Or (find_in_Responsible <> 0) Or (find_in_Lotus_hashtag <> 0) Or (find_in_Comment <> 0) Or (find_in_Protocol_Number <> 0) Then
      
      ' Счетчик выводимых строк на листе "To-Do"
      rowCount2 = rowCount2 + 1
      
      ' Номер по порядку
      НомерПоПорядку = НомерПоПорядку + 1
      
      ' Выводим на Лист "To-Do"
      ' №
      ThisWorkbook.Sheets("To-Do").Cells(rowCount2, 1).Value = CStr(НомерПоПорядку) + "."
      ThisWorkbook.Sheets("To-Do").Cells(rowCount2, 1).HorizontalAlignment = xlCenter
      
      ' Дата создания
      ThisWorkbook.Sheets("To-Do").Cells(rowCount2, 2).Value = Workbooks("ToDo").Sheets("Лист1").Cells(rowCount, 1).Value
      ThisWorkbook.Sheets("To-Do").Cells(rowCount2, 2).HorizontalAlignment = xlCenter
      
      ' Id_Task
      ThisWorkbook.Sheets("To-Do").Cells(rowCount2, 3).Value = Workbooks("ToDo").Sheets("Лист1").Cells(rowCount, 2).Value
      ThisWorkbook.Sheets("To-Do").Cells(rowCount2, 3).HorizontalAlignment = xlCenter
            
      ' Task
      ThisWorkbook.Sheets("To-Do").Cells(rowCount2, 4).Value = Workbooks("ToDo").Sheets("Лист1").Cells(rowCount, 3).Value
      
      ' Тема лотус
      ThisWorkbook.Sheets("To-Do").Cells(rowCount2, 5).Value = Workbooks("ToDo").Sheets("Лист1").Cells(rowCount, 4).Value
      
      ' 6 HashTag (если не заполнено, то ставим "")
      If IsEmpty(Workbooks("ToDo").Sheets("Лист1").Cells(rowCount, 6).Value) Then
        ThisWorkbook.Sheets("To-Do").Cells(rowCount2, 6).Value = " "
      Else
        ThisWorkbook.Sheets("To-Do").Cells(rowCount2, 6).Value = Workbooks("ToDo").Sheets("Лист1").Cells(rowCount, 6).Value
      End If
        
      ' 7 Статус задачи
      If Workbooks("ToDo").Sheets("Лист1").Cells(rowCount, 7).Value = 1 Then
        ThisWorkbook.Sheets("To-Do").Cells(rowCount2, 7).Value = "В работе"
      End If
      '
      If Workbooks("ToDo").Sheets("Лист1").Cells(rowCount, 7).Value = 0 Then
        ThisWorkbook.Sheets("To-Do").Cells(rowCount2, 7).Value = "Закрыта"
      End If
      
      ThisWorkbook.Sheets("To-Do").Cells(rowCount2, 7).HorizontalAlignment = xlCenter
        
      ' 8 Ответственный
      ThisWorkbook.Sheets("To-Do").Cells(rowCount2, 8).Value = Workbooks("ToDo").Sheets("Лист1").Cells(rowCount, 5).Value
      ThisWorkbook.Sheets("To-Do").Cells(rowCount2, 8).HorizontalAlignment = xlCenter
      
      ' 9 Дата контроля
      ThisWorkbook.Sheets("To-Do").Cells(rowCount2, 9).Value = Workbooks("ToDo").Sheets("Лист1").Cells(rowCount, 8).Value
      ThisWorkbook.Sheets("To-Do").Cells(rowCount2, 9).HorizontalAlignment = xlCenter
      
      ' 10 Комментарий
      ThisWorkbook.Sheets("To-Do").Cells(rowCount2, 10).Value = Workbooks("ToDo").Sheets("Лист1").Cells(rowCount, 9).Value
      
      ' 11 Протокол
      ThisWorkbook.Sheets("To-Do").Cells(rowCount2, 11).Value = Workbooks("ToDo").Sheets("Лист1").Cells(rowCount, 10).Value
      
      ' 12 В повестку
      ThisWorkbook.Sheets("To-Do").Cells(rowCount2, 12).Value = Workbooks("ToDo").Sheets("Лист1").Cells(rowCount, 16).Value
      
    End If
                  
    ' Следующая запись
    rowCount = rowCount + 1
    ' Application.StatusBar = officeNameInReport + ": " + CStr(rowCount) + "..."
    ' DoEventsInterval (rowCount)
          
  Loop
  
  
  ' Закрываем таблицу BASE\ToDo
  CloseBook ("ToDo")
 
  ' Сохранение изменений
  ThisWorkbook.Save

  ' Строка статуса
  Application.StatusBar = ""
    

End Sub

' Сформировать выписки из протоколов по поручениям на текущую дату
Sub Выписки_из_протокола()
Dim Текущий_исполнитель_должность, Текущий_исполнитель_ФИО, Дата_первого_поручения_str, Список_получателей_выписок As String
Dim Есть_поручения As Boolean
Dim Число_файлов_выписок As Byte

  ' Запрос на формирование
  If MsgBox("Сформировать выписки по исполнителям?", vbYesNo) = vbYes Then


    ' Открываем таблицу To-Do и выбираем все действующие поручения по Офису/Сотруднику
    

    ' Открываем таблицу BASE\ToDo
    OpenBookInBase ("ToDo")

    ' Переходим на окно DB
    ThisWorkbook.Sheets("To-Do").Activate

    ' Строка статуса
    Application.StatusBar = "Обработка..."

    ' Сколько файлов сформировано
    Число_файлов_выписок = 0

    ' Список получателей выписок
    Список_получателей_выписок = ""

    ' Выгружаем в Шаблон "Сводная выписка из протоколов"
    For i = 1 To 11
    
      ' Исполнители 9 шт.: НОРПиКО1, НОРПиКО2, НОРПиКО3, НОРПиКО4, НОРПиКО5, УДО2, УДО3, УДО4, УДО5
      Текущий_исполнитель_должность = ""
      Select Case i
        Case 1
          Текущий_исполнитель_должность = "НОРПиКО1"
        Case 2
          Текущий_исполнитель_должность = "НОРПиКО2"
        Case 3
          Текущий_исполнитель_должность = "НОРПиКО3"
        Case 4
          Текущий_исполнитель_должность = "НОРПиКО4"
        Case 5
          Текущий_исполнитель_должность = "НОРПиКО5"
        Case 6
          Текущий_исполнитель_должность = "УДО2"
        Case 7
          Текущий_исполнитель_должность = "УДО3"
        Case 8
          Текущий_исполнитель_должность = "УДО4"
        Case 9
          Текущий_исполнитель_должность = "УДО5"
        Case 10
          Текущий_исполнитель_должность = "НОКП"
        Case 11
          Текущий_исполнитель_должность = "РИЦ"
          
      End Select

      ' Определяем Текущий_исполнитель_ФИО
      Текущий_исполнитель_ФИО = Фамилия_и_Имя(getFromAddrBook(Текущий_исполнитель_должность, 4), 3)
      ' Есть поручения по этому исполнителю?
      Есть_поручения = False
      Дата_первого_поручения_str = ""
    
      ' Если текущий исполнитель<>""
      If Текущий_исполнитель_ФИО <> "" Then
      
        ' Выполняем поиск - есть ли по данному исполнителю поручения, попадающие в фильтр выборки?
        rowCount = 2
        Do While (Trim(Workbooks("ToDo").Sheets("Лист1").Cells(rowCount, 1).Value) <> "") And (Есть_поручения = False)
        
            ' Если (Protocol_Number<>"" и Protocol_Number2="") ИЛИ (Protocol_Number<>"" и Protocol_Number2=Текущему протоколу)
            If (Workbooks("ToDo").Sheets("Лист1").Cells(rowCount, 10).Value <> "") And (Workbooks("ToDo").Sheets("Лист1").Cells(rowCount, 13).Value = "") And (InStr(Workbooks("ToDo").Sheets("Лист1").Cells(rowCount, 5).Value, Текущий_исполнитель_ФИО)) Then
              Есть_поручения = True
              Дата_первого_поручения_str = CStr(Workbooks("ToDo").Sheets("Лист1").Cells(rowCount, 11).Value)
            End If
          
            ' Следующая строка
            DoEvents
            rowCount = rowCount + 1
        
          Loop ' Do While
      
      End If
      
      ' Если найдены поручения
      If Есть_поручения = True Then
      
        ' Число файлов-выписок
        Число_файлов_выписок = Число_файлов_выписок + 1
    
        ' Список получателей выписок
        If Список_получателей_выписок <> "" Then
          Список_получателей_выписок = Список_получателей_выписок + ", " + getFromAddrBook(Текущий_исполнитель_должность, 5)
        Else
          Список_получателей_выписок = getFromAddrBook(Текущий_исполнитель_должность, 5)
        End If

        ' Открываем шаблон "Выписка из Протоколов.xlsx" Открываем шаблон Протокола из C:\Users\...\Documents\#VBA\DB_Result\Templates
        Workbooks.Open (ThisWorkbook.Path + "\Templates\Выписка из Протоколов.xlsx")
         
        ' Имя файла с протоколом - берем из G2 "10-02032020"
        FileProtocolName = "Выписка из протоколов " + Replace(Текущий_исполнитель_ФИО, ".", "") + " (" + strДД_MM_YY(Date) + ").xlsx"
        Workbooks("Выписка из Протоколов.xlsx").SaveAs FileName:=ThisWorkbook.Path + "\Out\" + FileProtocolName, FileFormat:=xlOpenXMLWorkbook, createBackUp:=False

        ' Заполняем заголовок и ответственного
        Workbooks(FileProtocolName).Sheets("Действующие поручения").Cells(1, 3).Value = "Выписка из Протоколов на " + ДеньМесяцГод(Date)
        ' Ответственный
        Workbooks(FileProtocolName).Sheets("Действующие поручения").Cells(4, 4).Value = getFromAddrBook(Текущий_исполнитель_должность, 1)
        Workbooks(FileProtocolName).Sheets("Архив поручений").Cells(4, 4).Value = Workbooks(FileProtocolName).Sheets("Действующие поручения").Cells(4, 4).Value
       
        ' Обрабатываем Поручения по офису, где стоят даты
        Номер_поручения = 0
        Номер_поручения_Архив = 0
        ' Рейтинг
        Всего_поручений_закрытых = 0
        Всего_поручений_исполнено = 0
        Всего_поручений_не_исполнено = 0
        
        rowCount = 2
        Do While Trim(Workbooks("ToDo").Sheets("Лист1").Cells(rowCount, 1).Value) <> ""
        
            ' Если по исполнителю (Protocol_Number<>"" и Protocol_Number2="") ИЛИ (Protocol_Number<>"")
            If (Workbooks("ToDo").Sheets("Лист1").Cells(rowCount, 10).Value <> "") And (Workbooks("ToDo").Sheets("Лист1").Cells(rowCount, 13).Value = "") And (InStr(Workbooks("ToDo").Sheets("Лист1").Cells(rowCount, 5).Value, Текущий_исполнитель_ФИО)) Then
            
              Номер_поручения = Номер_поручения + 1
              
              ' Рейтинг
              ' Исполнено
              If InStr(Workbooks("ToDo").Sheets("Лист1").Cells(rowCount, 9).Value, "Исполнено") <> 0 Then
                Всего_поручений_исполнено = Всего_поручений_исполнено + 1
                Всего_поручений_закрытых = Всего_поручений_закрытых + 1
              End If
                  
              ' Не исполнено
              If InStr(Workbooks("ToDo").Sheets("Лист1").Cells(rowCount, 9).Value, "Не исполнено") <> 0 Then
                Всего_поручений_не_исполнено = Всего_поручений_не_исполнено + 1
                Всего_поручений_закрытых = Всего_поручений_закрытых + 1
              End If
            
            
              ' *** Выводим в Выписку ***
            
              If Номер_поручения > 1 Then
            
                ' Вставляем пустую строку в блок "Поручения"
                Workbooks(FileProtocolName).Sheets("Действующие поручения").Activate
                Workbooks(FileProtocolName).Sheets("Действующие поручения").Range(CStr(10 + Номер_поручения) + ":" + CStr(10 + Номер_поручения)).Select
                Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
            
                ' Нумерация "7." возможна только если формат преобразовать к текстовому ("@")
                Workbooks(FileProtocolName).Sheets("Действующие поручения").Cells(10 + Номер_поручения, 1).NumberFormat = "@"
                
                ' Объединяем B, С, D
                Workbooks(FileProtocolName).Sheets("Действующие поручения").Range("B" + CStr(10 + Номер_поручения) + ":D" + CStr(10 + Номер_поручения)).MergeCells = True
            
                ' Столбец F
                Workbooks(FileProtocolName).Sheets("Действующие поручения").Range("F" + CStr(10 + Номер_поручения) + ":F" + CStr(10 + Номер_поручения)).Select
                With Selection.Validation
                    .Delete
                    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="=Лист1!$A$1:$A$2"
                    .IgnoreBlank = True
                    .InCellDropdown = True
                    .InputTitle = ""
                    .ErrorTitle = ""
                    .InputMessage = ""
                    .ErrorMessage = ""
                    .ShowInput = True
                    .ShowError = True
                End With

            
                ' Объединяем G, Н
                Workbooks(FileProtocolName).Sheets("Действующие поручения").Range("G" + CStr(10 + Номер_поручения) + ":H" + CStr(10 + Номер_поручения)).MergeCells = True
            
                ' Рамка
                Call Рамка_в_строке_выписки_протокола(FileProtocolName, "Действующие поручения", Номер_поручения)

            End If ' Вставляем новую строку Поручения и нумеруем
          
            ' Номер протокола и номер вопроса
            Workbooks(FileProtocolName).Sheets("Действующие поручения").Cells(10 + Номер_поручения, 1).Value = "п. " + CStr(Workbooks("ToDo").Sheets("Лист1").Cells(rowCount, 12).Value) + " прот.№" + Workbooks("ToDo").Sheets("Лист1").Cells(rowCount, 10).Value
            Workbooks(FileProtocolName).Sheets("Действующие поручения").Range("A" + CStr(10 + Номер_поручения) + ":A" + CStr(10 + Номер_поручения)).WrapText = True
            Workbooks(FileProtocolName).Sheets("Действующие поручения").Cells(10 + Номер_поручения, 1).HorizontalAlignment = xlCenter
            Workbooks(FileProtocolName).Sheets("Действующие поручения").Cells(10 + Номер_поручения, 1).VerticalAlignment = xlCenter

            ' Поручение
            str_Поручениеi = Workbooks("ToDo").Sheets("Лист1").Cells(rowCount, 3).Value
          
            ' Поручение - "Переносить по словам"
            Workbooks(FileProtocolName).Sheets("Действующие поручения").Range("B" + CStr(10 + Номер_поручения) + ":D" + CStr(10 + Номер_поручения)).WrapText = True
            Workbooks(FileProtocolName).Sheets("Действующие поручения").Cells(10 + Номер_поручения, 2).HorizontalAlignment = xlLeft
            Workbooks(FileProtocolName).Sheets("Действующие поручения").Cells(10 + Номер_поручения, 2).VerticalAlignment = xlTop
            ' Поручение - высота строки
            Workbooks(FileProtocolName).Sheets("Действующие поручения").Range(CStr(10 + Номер_поручения) + ":" + CStr(10 + Номер_поручения)).RowHeight = lineHeight(str_Поручениеi, 15, 37) ' 20 - норм
            ' Поручение - Запись в выписку
            Workbooks(FileProtocolName).Sheets("Действующие поручения").Cells(10 + Номер_поручения, 2).Value = str_Поручениеi

            ' Срок исполнения
            Workbooks(FileProtocolName).Sheets("Действующие поручения").Cells(10 + Номер_поручения, 5).Value = CStr(Workbooks("ToDo").Sheets("Лист1").Cells(rowCount, 8).Value)
            Workbooks(FileProtocolName).Sheets("Действующие поручения").Range("E" + CStr(10 + Номер_поручения) + ":E" + CStr(10 + Номер_поручения)).WrapText = True
            Workbooks(FileProtocolName).Sheets("Действующие поручения").Cells(10 + Номер_поручения, 5).VerticalAlignment = xlTop
            Workbooks(FileProtocolName).Sheets("Действующие поручения").Cells(10 + Номер_поручения, 5).HorizontalAlignment = xlCenter
          
            ' Статус
            Workbooks(FileProtocolName).Sheets("Действующие поручения").Cells(10 + Номер_поручения, 6).Value = Workbooks("ToDo").Sheets("Лист1").Cells(rowCount, 9).Value
            Workbooks(FileProtocolName).Sheets("Действующие поручения").Cells(10 + Номер_поручения, 6).HorizontalAlignment = xlCenter
            Workbooks(FileProtocolName).Sheets("Действующие поручения").Cells(10 + Номер_поручения, 6).VerticalAlignment = xlTop

            ' Комментарий (выравнивание)
            Workbooks(FileProtocolName).Sheets("Действующие поручения").Cells(10 + Номер_поручения, 7).HorizontalAlignment = xlLeft
            Workbooks(FileProtocolName).Sheets("Действующие поручения").Cells(10 + Номер_поручения, 7).VerticalAlignment = xlTop


            ' *** Выводим в Выписку (конец) ***
            
            
          End If ' Если Protocol_Number<>"" и Protocol_Number2=""
        
        
          ' *** АРХИВ Поручений ***
          
                ' Если по исполнителю (Protocol_Number<>"" и Protocol_Number2<>"")
                If (Workbooks("ToDo").Sheets("Лист1").Cells(rowCount, 10).Value <> "") And (Workbooks("ToDo").Sheets("Лист1").Cells(rowCount, 13).Value <> "") And (InStr(Workbooks("ToDo").Sheets("Лист1").Cells(rowCount, 5).Value, Текущий_исполнитель_ФИО)) Then
                
                  Номер_поручения_Архив = Номер_поручения_Архив + 1
                
                  ' Рейтинг
                  ' Исполнено
                  If InStr(Workbooks("ToDo").Sheets("Лист1").Cells(rowCount, 9).Value, "Исполнено") <> 0 Then
                    Всего_поручений_исполнено = Всего_поручений_исполнено + 1
                    Всего_поручений_закрытых = Всего_поручений_закрытых + 1
                  End If
                  
                  ' Не исполнено
                  If InStr(Workbooks("ToDo").Sheets("Лист1").Cells(rowCount, 9).Value, "Не исполнено") <> 0 Then
                    Всего_поручений_не_исполнено = Всего_поручений_не_исполнено + 1
                    Всего_поручений_закрытых = Всего_поручений_закрытых + 1
                  End If
                  

                  ' *** Выводим в Выписку ***
                
                  If Номер_поручения_Архив > 1 Then
                
                    ' Вставляем пустую строку в блок "Поручения"
                    Workbooks(FileProtocolName).Sheets("Архив поручений").Activate
                    Workbooks(FileProtocolName).Sheets("Архив поручений").Range(CStr(10 + Номер_поручения_Архив) + ":" + CStr(10 + Номер_поручения_Архив)).Select
                    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
                
                    ' Нумерация "7." возможна только если формат преобразовать к текстовому ("@")
                    Workbooks(FileProtocolName).Sheets("Архив поручений").Cells(10 + Номер_поручения_Архив, 1).NumberFormat = "@"
                    
                    ' Объединяем B, С, D
                    Workbooks(FileProtocolName).Sheets("Архив поручений").Range("B" + CStr(10 + Номер_поручения_Архив) + ":D" + CStr(10 + Номер_поручения_Архив)).MergeCells = True
                
                    ' Объединяем G, Н
                    Workbooks(FileProtocolName).Sheets("Архив поручений").Range("G" + CStr(10 + Номер_поручения_Архив) + ":H" + CStr(10 + Номер_поручения_Архив)).MergeCells = True
                
                    ' Рамка
                    Call Рамка_в_строке_выписки_протокола(FileProtocolName, "Архив поручений", Номер_поручения_Архив)
    
                End If ' Вставляем новую строку Поручения и нумеруем
              
                ' Номер протокола и номер вопроса
                Workbooks(FileProtocolName).Sheets("Архив поручений").Cells(10 + Номер_поручения_Архив, 1).Value = "п. " + CStr(Workbooks("ToDo").Sheets("Лист1").Cells(rowCount, 12).Value) + " прот.№" + Workbooks("ToDo").Sheets("Лист1").Cells(rowCount, 10).Value
                Workbooks(FileProtocolName).Sheets("Архив поручений").Range("A" + CStr(10 + Номер_поручения_Архив) + ":A" + CStr(10 + Номер_поручения_Архив)).WrapText = True
                Workbooks(FileProtocolName).Sheets("Архив поручений").Cells(10 + Номер_поручения_Архив, 1).HorizontalAlignment = xlCenter
                Workbooks(FileProtocolName).Sheets("Архив поручений").Cells(10 + Номер_поручения_Архив, 1).VerticalAlignment = xlCenter
    
                ' Поручение
                str_Поручениеi = Workbooks("ToDo").Sheets("Лист1").Cells(rowCount, 3).Value
              
                ' Поручение - "Переносить по словам"
                Workbooks(FileProtocolName).Sheets("Архив поручений").Range("B" + CStr(10 + Номер_поручения_Архив) + ":D" + CStr(10 + Номер_поручения_Архив)).WrapText = True
                Workbooks(FileProtocolName).Sheets("Архив поручений").Cells(10 + Номер_поручения_Архив, 2).HorizontalAlignment = xlLeft
                Workbooks(FileProtocolName).Sheets("Архив поручений").Cells(10 + Номер_поручения_Архив, 2).VerticalAlignment = xlTop
                ' Поручение - высота строки
                Workbooks(FileProtocolName).Sheets("Архив поручений").Range(CStr(10 + Номер_поручения_Архив) + ":" + CStr(10 + Номер_поручения_Архив)).RowHeight = lineHeight(str_Поручениеi, 15, 37) ' 20 - норм
                ' Поручение - Запись в выписку
                Workbooks(FileProtocolName).Sheets("Архив поручений").Cells(10 + Номер_поручения_Архив, 2).Value = str_Поручениеi
    
                ' Срок исполнения
                Workbooks(FileProtocolName).Sheets("Архив поручений").Cells(10 + Номер_поручения_Архив, 5).Value = CStr(Workbooks("ToDo").Sheets("Лист1").Cells(rowCount, 8).Value)
                Workbooks(FileProtocolName).Sheets("Архив поручений").Range("E" + CStr(10 + Номер_поручения_Архив) + ":E" + CStr(10 + Номер_поручения_Архив)).WrapText = True
                Workbooks(FileProtocolName).Sheets("Архив поручений").Cells(10 + Номер_поручения_Архив, 5).VerticalAlignment = xlTop
                Workbooks(FileProtocolName).Sheets("Архив поручений").Cells(10 + Номер_поручения_Архив, 5).HorizontalAlignment = xlCenter
              
                ' Статус
                Workbooks(FileProtocolName).Sheets("Архив поручений").Cells(10 + Номер_поручения_Архив, 6).Value = Workbooks("ToDo").Sheets("Лист1").Cells(rowCount, 9).Value
                Workbooks(FileProtocolName).Sheets("Архив поручений").Cells(10 + Номер_поручения_Архив, 6).HorizontalAlignment = xlCenter
                Workbooks(FileProtocolName).Sheets("Архив поручений").Cells(10 + Номер_поручения_Архив, 6).VerticalAlignment = xlTop
    
                ' Комментарий
                Workbooks(FileProtocolName).Sheets("Архив поручений").Cells(10 + Номер_поручения_Архив, 7).Value = Workbooks("ToDo").Sheets("Лист1").Cells(rowCount, 17).Value
                ' Комментарий - файл с отчетом
                If Workbooks("ToDo").Sheets("Лист1").Cells(rowCount, 18).Value <> "" Then
                  Workbooks(FileProtocolName).Sheets("Архив поручений").Cells(10 + Номер_поручения_Архив, 7).Value = CStr(Workbooks(FileProtocolName).Sheets("Архив поручений").Cells(10 + Номер_поручения_Архив, 7).Value) + " Отчет: " + CStr(Workbooks("ToDo").Sheets("Лист1").Cells(rowCount, 18).Value)
                End If
                Workbooks(FileProtocolName).Sheets("Архив поручений").Cells(10 + Номер_поручения_Архив, 7).HorizontalAlignment = xlLeft
                Workbooks(FileProtocolName).Sheets("Архив поручений").Cells(10 + Номер_поручения_Архив, 7).VerticalAlignment = xlTop
    
                ' *** Выводим в Выписку (конец) ***
                
                
              End If ' Если Protocol_Number<>"" и Protocol_Number2<>""

          
          ' *** конец вывода в АРХИВ Поручений
        
          ' Индикация
          Application.StatusBar = "Обработка " + CStr(i) + "..."
        
          ' Следующая строка
          DoEvents
          rowCount = rowCount + 1
        Loop ' Do While
    
        ' Период
        Workbooks(FileProtocolName).Sheets("Действующие поручения").Cells(5, 4).Value = "с " + Дата_первого_поручения_str + " по " + CStr(Date)
        Workbooks(FileProtocolName).Sheets("Архив поручений").Cells(5, 4).Value = "с " + Дата_первого_поручения_str + " по " + CStr(Date)

        ' Доля исполненных поручений за период, %
        Workbooks(FileProtocolName).Sheets("Действующие поручения").Cells(6, 4).Value = CStr(РассчетДоли(Всего_поручений_закрытых, Всего_поручений_исполнено, 2) * 100) + "%"
        Workbooks(FileProtocolName).Sheets("Архив поручений").Cells(6, 1).Value = "Доля исполненных поручений за период, шт."
        Workbooks(FileProtocolName).Sheets("Архив поручений").Cells(6, 4).Value = "Всего поручений: " + CStr(Всего_поручений_закрытых) + " шт., в т.ч. со статусом Исполнено: " + CStr(Всего_поручений_исполнено) + " шт."
    
        ' Закрытие файла с Протоколом Собрания
        Workbooks(FileProtocolName).Sheets("Действующие поручения").Activate
        Workbooks(FileProtocolName).Close SaveChanges:=True
      
      End If ' Если поручения не найдены по текущему исполнителю
      
    Next i ' Следующий офис
    
    ' Формируем шаблон письма для отправки по схеме 1 выписка = 1 письмо
  
    ' Закрываем таблицу BASE\ToDo
    CloseBook ("ToDo")
 
    ' Строка статуса
    Application.StatusBar = ""
      
    ' Сообщение
    MsgBox ("Сформировано " + CStr(Число_файлов_выписок) + " выписок!")

    ' Запрос на отправку Шаблона сообщения Хэштег2 берем с листа ЕСУП
    If MsgBox("Отправить шаблон письма с Выписками в работу?", vbYesNo) = vbYes Then
          
      ' Формирование отправки письма
      ' Тема письма - Тема:
      ' темаПисьма = ThisWorkbook.Sheets("Лист8").Cells(RowByValue(ThisWorkbook.Name, "Лист8", "Тема:", 100, 100), ColumnByValue(ThisWorkbook.Name, "Лист8", "Тема:", 100, 100) + 1).Value
      темаПисьма = ">>: " + ThisWorkbook.Sheets("ЕСУП").Range("Q1").Value ' + subjectFromSheet("ЕСУП")

      ' hashTag - Хэштэг:
      hashTag = "#protocol" ' hashTagFromSheetII("ЕСУП", 1)

      ' Файл-вложение (!!!)
      ' attachmentFile = ThisWorkbook.Sheets("Лист8").Cells(3, 17).Value
    
      ' Текст письма
      текстПисьма = "" + Chr(13)
      текстПисьма = текстПисьма + "" + Список_получателей_выписок + Chr(13) ' ThisWorkbook.Sheets("ЕСУП").Cells(rowByValue(ThisWorkbook.Name, "ЕСУП", "Список получателей:", 100, 100), ColumnByValue(ThisWorkbook.Name, "ЕСУП", "Список получателей:", 100, 100) + 2).Value + Chr(13)
      текстПисьма = текстПисьма + "" + Chr(13)
      текстПисьма = текстПисьма + "" + getFromAddrBook("РД", 2) + Chr(13)
      текстПисьма = текстПисьма + "" + Chr(13)
      текстПисьма = текстПисьма + "Уважаемые руководители," + Chr(13)
      текстПисьма = текстПисьма + "" + Chr(13)
      текстПисьма = текстПисьма + "Направляю индивидуальные выписки из протокола на " + CStr(Date) + " г." + Chr(13)
      текстПисьма = текстПисьма + "" + Chr(13)
      текстПисьма = текстПисьма + "Прошу в срок до " + strDDMM(Первый_понедельник_от_даты(Date)) + " по итогам исполнения задачи проставить отметку (Исполнено/Не исполнено), при необходимости заполнить комментарий и направить файл в мой адрес для внесения в Протокол." + Chr(13)
      ' Визитка (подпись С Ув., )
      текстПисьма = текстПисьма + ПодписьВПисьме()
      ' Хэштег
      текстПисьма = текстПисьма + createBlankStr(27) + hashTag
      ' Вызов
      Call send_Lotus_Notes(темаПисьма, "Sergey Fedorovich Proschaev/Tyumen/PSBank/Ru", "Sergey Fedorovich Proschaev/Tyumen/PSBank/Ru", текстПисьма, "")
  
      ' Сообщение
      MsgBox ("Письмо отправлено!")
    
    End If

    ' Сохранение изменений
    ' ThisWorkbook.Save
  
  End If
  
End Sub


' Рамка
Sub Рамка_в_строке_выписки_протокола(In_FileProtocolName, In_Sheets, In_Номер_поручения)

                Workbooks(In_FileProtocolName).Sheets(In_Sheets).Range("A" + CStr(10 + In_Номер_поручения) + ":H" + CStr(10 + In_Номер_поручения)).Select
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

End Sub

' Обработать исполнение поручений
Sub Обработка_исполнения_поручений()
Dim ReportName_String, officeNameInReport, CheckFormatReportResult, Id_TaskVar As String
Dim i, rowCount, rowCount_searchResults As Integer
Dim finishProcess As Boolean
    
  ' Открыть файл с отчетом
  FileName = Application.GetOpenFilename("Excel Files (*.xlsx), *.xlsx", , "Открытие файла с отчетом")

  ' Если файл был выбран
  If (Len(FileName) > 5) Then
  
    ' Строка статуса
    Application.StatusBar = "Обработка отчетов..."
  
    ' Переменная начала обработки
    finishProcess = False

    ' Выводим для инфо данные об имени файла
    ReportName_String = Dir(FileName)
  
    ' Открываем выбранную книгу (UpdateLinks:=0)
    Workbooks.Open FileName, 0
      
    ' Открываем таблицу BASE\ToDo
    OpenBookInBase ("ToDo")
  
    ' Переходим на окно DB
    ThisWorkbook.Sheets("To-Do").Activate

    ' Проверка формы отчета
    CheckFormatReportResult = CheckFormatReport(ReportName_String, "Действующие поручения", 14, Date)
    If CheckFormatReportResult = "OK" Then
      
        ' Счетчик ошибок (проставления статуса исполнения)
        Count_error = 0
        ' Обработано поручений
        ОбработаноПоручений = 0
        
        ' Проверяем - был ли ранее обработан
        If Workbooks(ReportName_String).Sheets("Действующие поручения").Range("H1").Value <> "" Then
          MsgBox ("Внимание файл " + Workbooks(ReportName_String).Sheets("Действующие поручения").Range("H1").Value + "!")
        End If
        
        ' Находим позицию "№ прот., поруч."
        rowCount = rowByValue(ReportName_String, "Действующие поручения", "№ прот., поруч.", 100, 100) + 1
        Do While Not IsEmpty(Workbooks(ReportName_String).Sheets("Действующие поручения").Cells(rowCount, 1).Value)
        
       
                 
          ' Если это текущий офис
          If (Workbooks(ReportName_String).Sheets("Действующие поручения").Cells(rowCount, 6).Value = "Исполнено") Or (Workbooks(ReportName_String).Sheets("Действующие поручения").Cells(rowCount, 6).Value = "Не исполнено") Then
            
            ' Статус поручения проставлен верно - апдейтим таблицу To-Do для этой задачи
            
            ' Id_Task = Число "431910202019" из "п. 19 прот.№43-19102020"
            Id_TaskVar = genId_Task(Workbooks(ReportName_String).Sheets("Действующие поручения").Cells(rowCount, 1).Value)
            
            ' Выполняем поиск в To-DO
            Set searchResults = Workbooks("ToDo").Sheets("Лист1").Columns("B:B").Find(Id_TaskVar, LookAt:=xlWhole)
  
            ' Если найдена, то апдейтим
            If searchResults Is Nothing Then
              ' Иначе - сообщение об ошибке, что не найдена такая задача!
              ' Счетчик ошибок
              Count_error = Count_error + 1
              ' Сообщение о неверном формате отчета или даты
              MsgBox ("Не найдена задача: " + CStr(Id_TaskVar) + "!")
            Else
              
              ' Если найдена, то апдейтим в строке searchResults.Row
              rowCount_searchResults = searchResults.Row
              
              ' 1) Столбец I (9)  "Comment" - устанавливаем "Исполнено/Не исполнено"
              Workbooks("ToDo").Sheets("Лист1").Cells(rowCount_searchResults, 9).Value = Workbooks(ReportName_String).Sheets("Действующие поручения").Cells(rowCount, 6).Value
              
              ' 2) Столбец Q (17) "OfficeComment_protocolReport"  - Комментарий офиса
              Workbooks("ToDo").Sheets("Лист1").Cells(rowCount_searchResults, 17).Value = Workbooks(ReportName_String).Sheets("Действующие поручения").Cells(rowCount, 7).Value
              
              ' 3) Столбец R (18) "FileName_protocolReport"  - Имя файла с отчетом об исполнении поручений из которого загружена информация ReportName_String
              Workbooks("ToDo").Sheets("Лист1").Cells(rowCount_searchResults, 18).Value = ReportName_String
            
              ' 4) Столбец G (7) "Task_Status" = 0
              Workbooks("ToDo").Sheets("Лист1").Cells(rowCount_searchResults, 7).Value = 0
            
              ' Обработано поручений
              ОбработаноПоручений = ОбработаноПоручений + 1
          
            
            End If

            
          Else
            ' Счетчик ошибок
            Count_error = Count_error + 1
            ' Сообщение о неверном формате отчета или даты
            MsgBox ("Не верно проставлено исполнение: " + Workbooks(ReportName_String).Sheets("Действующие поручения").Cells(rowCount, 1).Value + " Статус - " + Workbooks(ReportName_String).Sheets("Действующие поручения").Cells(rowCount, 6).Value + "!")
          End If
        
          ' Следующая запись
          rowCount = rowCount + 1
          Application.StatusBar = CStr(rowCount) + "..."
          DoEventsInterval (rowCount)
        Loop
   
        ' Выводим данные по офису
      
      ' Выводим итоги обработки
      
      ' Строка статуса
      Application.StatusBar = "Завершение..."
      ' Сохранение изменений
      ThisWorkbook.Save
    
      ' Переменная завершения обработки
      finishProcess = True
    Else
      ' Сообщение о неверном формате отчета или даты
      MsgBox ("Проверьте отчет: " + CheckFormatReportResult + "!")
    End If ' Проверка формы отчета

    ' Если ошибок нет, то сохраняем файл в In\Completed
    If Count_error = 0 Then
      ' В "H1" ставим статус обработано Дата и Время
      Workbooks(ReportName_String).Sheets("Действующие поручения").Range("H1").Value = "Обработано " + CStr(Date) + " " + CStr(Time)
      
      ' Шрифт
      Workbooks(ReportName_String).Sheets("Действующие поручения").Range("H1").Font.Name = "Calibri"
      Workbooks(ReportName_String).Sheets("Действующие поручения").Range("H1").Font.Size = 8
      Workbooks(ReportName_String).Sheets("Действующие поручения").Range("H1").Font.Strikethrough = False
      Workbooks(ReportName_String).Sheets("Действующие поручения").Range("H1").Font.Superscript = False
      Workbooks(ReportName_String).Sheets("Действующие поручения").Range("H1").Font.Subscript = False
      Workbooks(ReportName_String).Sheets("Действующие поручения").Range("H1").Font.OutlineFont = False
      Workbooks(ReportName_String).Sheets("Действующие поручения").Range("H1").Font.Shadow = False
      Workbooks(ReportName_String).Sheets("Действующие поручения").Range("H1").Font.Underline = xlUnderlineStyleNone
      Workbooks(ReportName_String).Sheets("Действующие поручения").Range("H1").Font.ThemeColor = xlThemeColorLight1
      Workbooks(ReportName_String).Sheets("Действующие поручения").Range("H1").Font.TintAndShade = 0
      Workbooks(ReportName_String).Sheets("Действующие поручения").Range("H1").Font.ThemeFont = xlThemeFontMinor
      Workbooks(ReportName_String).Sheets("Действующие поручения").Range("H1").Font.Italic = True

      ' Cохраняем файл в In\Completed
      Workbooks(Dir(FileName)).SaveAs FileName:=ThisWorkbook.Path + "\In\Completed\" + Dir(FileName), FileFormat:=xlOpenXMLWorkbook, createBackUp:=False
      Workbooks(Dir(FileName)).Close SaveChanges:=True
    Else
      ' Закрываем файл с отчетом без сохранения изменений (параметр SaveChanges:=False)
      Workbooks(Dir(FileName)).Close SaveChanges:=False
    End If
    
    ' Закрываем таблицу BASE\ToDo
    CloseBook ("ToDo")

    ' Переходим в ячейку M2
    ThisWorkbook.Sheets("To-Do").Activate

    ' Строка статуса
    Application.StatusBar = ""
    
    ' Итоговое сообщение
    If finishProcess = True Then
      ' Есть ошибки?
      If Count_error = 0 Then
        MsgBox ("Обработка " + Dir(ReportName_String) + " завершена. Обработано успешно " + CStr(ОбработаноПоручений) + " поручений!")
      Else
        ' Если есть ошибки
        MsgBox ("Обработка " + Dir(ReportName_String) + " завершена! Ошибок: " + CStr(Count_error))
      
        ' Формируем запрос - отправить сообщение в LN исполнителю?
        If MsgBox("Отправить уведомление в адрес исполнителя?", vbYesNo) = vbYes Then
          
          ' Формирование отправки письма
          
          
          ' Сообщение
          MsgBox ("Шаблон письма отправлен!")
        End If
      
      End If
    Else
      MsgBox ("Обработка отчета была прервана!")
    End If

  End If ' Если файл был выбран

End Sub

' Генерация Id_Task в To-DO
Function genId_Task(In_NumberStr) As String

  ' Номер протокола 43-19102020 из "п. 19 прот.№43-19102020"
  Номер_протокола = Mid(In_NumberStr, InStr(In_NumberStr, "№") + 1, Len(In_NumberStr) - InStr(In_NumberStr, "№"))
  Номер_протокола = Replace(Номер_протокола, "-", "")
  
  ' Номер пункта 19 из "п. 19 прот.№43-19102020"
  ' t1 = InStr(In_NumberStr, "п.") + 3
  ' t2 = InStr(In_NumberStr, "прот.") - 1
  ' t3 = InStr(In_NumberStr, "прот.") - InStr(In_NumberStr, "п.") - 4
  
  ' Номер_пункта = Mid(In_NumberStr, )
  Номер_пункта = Mid(In_NumberStr, InStr(In_NumberStr, "п.") + 3, InStr(In_NumberStr, "прот.") - InStr(In_NumberStr, "п.") - 4)
  
  ' Id_Task = Число "431910202019" из "п. 19 прот.№43-19102020"
  t = Номер_протокола + Номер_пункта
  genId_Task = t
  
End Function

' Закрыть Поручение в таблице на Листе To-DO из Протокола
Sub Task_From_Protocol_ToClose()
        
  ' Строка
  ThisWorkbook.Sheets("To-Do").Cells(ActiveCell.Row, 7).Value = "Закрыта"

  ' Комментарий = Исполнено
  ThisWorkbook.Sheets("To-Do").Cells(ActiveCell.Row, 10).Value = "Исполнено"

  ' Сохранение изменений
  ThisWorkbook.Save

  ' Запрос на запуск процедуры ToDo_save() сохранения в BASE\To-Do
  If MsgBox("Сохранить изменения в BASE\ToDo?", vbYesNo) = vbYes Then
    Call ToDo_save
  End If


End Sub

' Закрыть Поручение в таблице на Листе To-DO из Протокола - статус "Не исполнено"
Sub Task_From_Protocol_ToClose_2()
        
  ' Строка
  ThisWorkbook.Sheets("To-Do").Cells(ActiveCell.Row, 7).Value = "Закрыта"

  ' Комментарий = Исполнено
  ThisWorkbook.Sheets("To-Do").Cells(ActiveCell.Row, 10).Value = "Не исполнено"

  ' Сохранение изменений
  ThisWorkbook.Save

  ' Запрос на запуск процедуры ToDo_save() сохранения в BASE\To-Do
  If MsgBox("Сохранить изменения в BASE\ToDo?", vbYesNo) = vbYes Then
    Call ToDo_save
  End If


End Sub


