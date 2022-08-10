Attribute VB_Name = "Module_Лист0"
' Лист 0 и навигации между листами

' *** Глобальные переменные ***
Public backUpFileName As String ' Имя Бэкап файла
' ***                       ***


' Установка операционной даты
Sub setDateNow()

Dim currentOperDate As Date
Dim firstDayOfWeek As Boolean

  ' Установка даты
  ' currentOperDate = Date ' + 1  ' (для тестирования день +1)
  
  ' Если стоит инкрементное увеличение даты, т.е. дата = дата + 1 "Инкрементное увеличение даты:"
  If CStr(ThisWorkbook.Sheets("Лист0").Cells(rowByValue(ThisWorkbook.Name, "Лист0", "Инкрементное увеличение даты:", 100, 100), ColumnByValue(ThisWorkbook.Name, "Лист0", "Инкрементное увеличение даты:", 100, 100) + 3).Value) = "1" Then
    
    currentOperDate = CDate(ThisWorkbook.Sheets("Лист0").Cells(rowByValue(ThisWorkbook.Name, "Лист0", "Операционная дата:", 100, 100), ColumnByValue(ThisWorkbook.Name, "Лист0", "Операционная дата:", 100, 100) + 3).Value) + 1
    
  Else
    
    ' Иначе используем системную дату
    currentOperDate = Date
    
  End If
  
  ' Проверка - первый день недели
  firstDayOfWeek = False
  If CStr(ThisWorkbook.Sheets("Лист0").Cells(rowByValue(ThisWorkbook.Name, "Лист0", "Первый день недели:", 100, 100), ColumnByValue(ThisWorkbook.Name, "Лист0", "Первый день недели:", 100, 100) + 2).Value) = "1" Then
    ' Обработать как первый день недели?
    If MsgBox("Сформировать задачи первого дня недели?", vbYesNo) = vbYes Then
      firstDayOfWeek = True
    Else
      firstDayOfWeek = False
    End If
  End If

  ' Номер недели:
  ThisWorkbook.Sheets("Лист0").Cells(rowByValue(ThisWorkbook.Name, "Лист0", "Неделя:", 100, 100), ColumnByValue(ThisWorkbook.Name, "Лист0", "Неделя:", 100, 100) + 1).Value = WeekNumber(currentOperDate)
  ' Дата
  ThisWorkbook.Sheets("Лист0").Cells(rowByValue(ThisWorkbook.Name, "Лист0", "Операционная дата:", 100, 100), ColumnByValue(ThisWorkbook.Name, "Лист0", "Операционная дата:", 100, 100) + 3).Value = currentOperDate
  ' Установка (дня недели)
  ThisWorkbook.Sheets("Лист0").Cells(rowByValue(ThisWorkbook.Name, "Лист0", "Операционная дата:", 100, 100), ColumnByValue(ThisWorkbook.Name, "Лист0", "Операционная дата:", 100, 100) + 4).Value = "(" + ДеньНедели(currentOperDate) + ")"
  
  ' Если сегодня понедельник - то прогнозы
  If (Weekday(currentOperDate, vbMonday) = 1) Or (firstDayOfWeek = True) Then
     
     ' Если понедельник (или первый день недели)
     ' Блок напоминаний:
     ' Вкладчики
     Call ВыделениеЖирнымТекстаВячейке("Лист0", RangeByValue(ThisWorkbook.Name, "Лист0", "Напоминание: прислать отчет по отработке вкладов", 100, 100))
     ' Прогнозы на неделю
     Call ВыделениеЖирнымТекстаВячейке("Лист0", RangeByValue(ThisWorkbook.Name, "Лист0", "Напоминание: прислать прогнозы на неделю", 100, 100))
     ' Запрос долгов у Сизиковой
     Call ВыделениеЖирнымТекстаВячейке("Лист0", RangeByValue(ThisWorkbook.Name, "Лист0", "Задолженность по списанию карт и КД", 100, 100))
     ' Оперативная справка по активам за неделю
     Call ВыделениеЖирнымТекстаВячейке("Лист0", RangeByValue(ThisWorkbook.Name, "Лист0", "Оперативная справка по активам за неделю", 100, 100))
     ' Оперативная справка по заявкам на карты за неделю
     Call ВыделениеЖирнымТекстаВячейке("Лист0", RangeByValue(ThisWorkbook.Name, "Лист0", "Оперативная справка по заявкам на карты за неделю", 100, 100))
     '
     Call ВыделениеЖирнымТекстаВячейке("Лист0", RangeByValue(ThisWorkbook.Name, "Лист0", "Открыть новую неделю в ЕСУП", 100, 100))
     Call ВыделениеЖирнымТекстаВячейке("Лист0", RangeByValue(ThisWorkbook.Name, "Лист0", "Отчет по выходу вкладчиков на неделю", 100, 100))
     '
     Call ВыделениеБледнымТекстаВячейке("Лист0", RangeByValue(ThisWorkbook.Name, "Лист0", "Оперативная справка по активам", 100, 100))
     Call ВыделениеБледнымТекстаВячейке("Лист0", RangeByValue(ThisWorkbook.Name, "Лист0", "Оперативная справка по заявкам на карты", 100, 100))
     
     ' Отправить Протокол Собрания в почте и в каталог ЕСУП
     Call ВыделениеБледнымТекстаВячейке("Лист0", RangeByValue(ThisWorkbook.Name, "Лист0", "Отправить Протокол Собрания в почте и в каталог ЕСУП", 100, 100))
     
     ' Перенести в Архив Повестку Собрания
     Call ВыделениеБледнымТекстаВячейке("Лист0", RangeByValue(ThisWorkbook.Name, "Лист0", "Перенести в Архив Повестку Собрания", 100, 100))
          
     Call ВыделениеБледнымТекстаВячейке("Лист0", RangeByValue(ThisWorkbook.Name, "Лист0", "Отчет кросс-продажи Capacity Model", 100, 100))
     
     ' Убираем отметки "Отпр.:" (для контроля прямой отправки сообщений пользователям)
     ThisWorkbook.Sheets("Лист0").Cells(rowByValue(ThisWorkbook.Name, "Лист0", "Напоминание: прислать отчет по отработке вкладов", 100, 100), ColumnByValue(ThisWorkbook.Name, "Лист0", "Напоминание: прислать отчет по отработке вкладов", 100, 100) + 6).Value = ""
     ThisWorkbook.Sheets("Лист0").Cells(rowByValue(ThisWorkbook.Name, "Лист0", "Задолженность по списанию карт и КД", 100, 100), ColumnByValue(ThisWorkbook.Name, "Лист0", "Задолженность по списанию карт и КД", 100, 100) + 6).Value = ""
  
     ' Клиентопоток и продажи
     Call ВыделениеЖирнымТекстаВячейке("Лист0", RangeByValue(ThisWorkbook.Name, "Лист0", "Клиентопоток и продажи", 100, 100))
       
     ' Банкострахование
     Call ВыделениеЖирнымТекстаВячейке("Лист0", RangeByValue(ThisWorkbook.Name, "Лист0", "Банкострахование", 100, 100))
  
       
  Else
     
     ' Если не понедельник, то заклеиваем:
     ' Напоминания
     Call ВыделениеБледнымТекстаВячейке("Лист0", RangeByValue(ThisWorkbook.Name, "Лист0", "Напоминание: прислать отчет по отработке вкладов", 100, 100))
     Call ВыделениеБледнымТекстаВячейке("Лист0", RangeByValue(ThisWorkbook.Name, "Лист0", "Напоминание: прислать прогнозы на неделю", 100, 100))
     Call ВыделениеБледнымТекстаВячейке("Лист0", RangeByValue(ThisWorkbook.Name, "Лист0", "Задолженность по списанию карт и КД", 100, 100))
     '
     Call ВыделениеБледнымТекстаВячейке("Лист0", RangeByValue(ThisWorkbook.Name, "Лист0", "Оперативная справка по активам за неделю", 100, 100))
     Call ВыделениеБледнымТекстаВячейке("Лист0", RangeByValue(ThisWorkbook.Name, "Лист0", "Оперативная справка по заявкам на карты за неделю", 100, 100))
     Call ВыделениеБледнымТекстаВячейке("Лист0", RangeByValue(ThisWorkbook.Name, "Лист0", "Отправить Протокол Собрания в почте и в каталог ЕСУП", 100, 100))
     Call ВыделениеБледнымТекстаВячейке("Лист0", RangeByValue(ThisWorkbook.Name, "Лист0", "Перенести в Архив Повестку Собрания", 100, 100))
     Call ВыделениеБледнымТекстаВячейке("Лист0", RangeByValue(ThisWorkbook.Name, "Лист0", "Открыть новую неделю в ЕСУП", 100, 100))
     Call ВыделениеБледнымТекстаВячейке("Лист0", RangeByValue(ThisWorkbook.Name, "Лист0", "Отчет по выходу вкладчиков на неделю", 100, 100))
     
     ' Выделение жирным
     Call ВыделениеЖирнымТекстаВячейке("Лист0", RangeByValue(ThisWorkbook.Name, "Лист0", "DashBoard (при наличии)", 100, 100))
     
     ' Активы:
     Call ВыделениеЖирнымТекстаВячейке("Лист0", RangeByValue(ThisWorkbook.Name, "Лист0", "Оперативная справка по активам", 100, 100))
     
     ' Карты:
     Call ВыделениеЖирнымТекстаВячейке("Лист0", RangeByValue(ThisWorkbook.Name, "Лист0", "Оперативная справка по заявкам на карты", 100, 100))
    
     ' Отчет кросс-продажи Capacity Model
     Call ВыделениеЖирнымТекстаВячейке("Лист0", RangeByValue(ThisWorkbook.Name, "Лист0", "Отчет кросс-продажи Capacity Model", 100, 100))
     
     ' Просрочки CRM (старый Отчет План-Факт по продуктам ИСЖ_НСЖ)
     Call ВыделениеЖирнымТекстаВячейке("Лист0", RangeByValue(ThisWorkbook.Name, "Лист0", "Просрочки CRM", 100, 100))
     
     ' Клиентопоток и продажи
     Call ВыделениеЖирнымТекстаВячейке("Лист0", RangeByValue(ThisWorkbook.Name, "Лист0", "Клиентопоток и продажи", 100, 100))
     
     ' Банкострахование
     Call ВыделениеЖирнымТекстаВячейке("Лист0", RangeByValue(ThisWorkbook.Name, "Лист0", "Банкострахование", 100, 100))
    
     
     ' Активы - дата отчета:
     ' Call setPriodReport("Лист3", currentOperDate)
          
     ' Карты - дата отчета:
     ' Call setPriodReport("Лист5", currentOperDate)
                         
  End If

  

  ' Установка периодов отчетов:
  ' Активы - дата отчета
  Call setPriodReport("Лист3", currentOperDate)
          
  ' Карты - дата отчета
  Call setPriodReport("Лист5", currentOperDate)
  
  ' Банкострахование
  Call setPriodReport("Лист10", currentOperDate)
  
  ' Клиентопоток - дата нового дня
  ThisWorkbook.Sheets("Лист12").Range("H2").Value = currentOperDate

  ' Запускаем обновление To-Do, ставим опцию выводить Протоколы = 1
  ThisWorkbook.Sheets("To-Do").Range("I1").Value = 1

  ' Запускаем обновление To-Do
  Call ToDo_refresh

  ' Сбрасываем
  ThisWorkbook.Sheets("To-Do").Range("I1").Value = 0
  
  ' D7, D8 и D9
  ' ThisWorkbook.Sheets("Лист0").Cells(7, 4).Value = "1) DashBoard (при наличии)"
  ' Call ВыделениеЖирнымТекстаВячейке("Лист0", "D7")
  ' Call ВыделениеЖирнымТекстаВячейке("Лист0", "D8")
  ' Call ВыделениеЖирнымТекстаВячейке("Лист0", "D9")
  
  ' Call ВыделениеЖирнымТекстаВячейке("Лист0", RangeByValue(ThisWorkbook.Name, "Лист0", "DashBoard (при наличии)", 100, 100))
  ' Call ВыделениеЖирнымТекстаВячейке("Лист0", RangeByValue(ThisWorkbook.Name, "Лист0", "Оперативная справка по активам", 100, 100))
  ' Call ВыделениеЖирнымТекстаВячейке("Лист0", RangeByValue(ThisWorkbook.Name, "Лист0", "Оперативная справка по заявкам на карты", 100, 100))
  
  ' Установка первого дня недели
  If firstDayOfWeek = True Then
    ThisWorkbook.Sheets("Лист0").Cells(rowByValue(ThisWorkbook.Name, "Лист0", "Первый день недели:", 100, 100), ColumnByValue(ThisWorkbook.Name, "Лист0", "Первый день недели:", 100, 100) + 1).Value = "0"
  End If

  ' Отправка уведомлений о ДР клиентов из BASE\Birthdays
  Call Отправка_уведомлений_ДР

  ' Создание BackUp
  Call createBackUp

  ' Переход на Лист0
  ThisWorkbook.Sheets("Лист0").Activate

  ' Сохранение изменений
  ThisWorkbook.Save

End Sub

' Переход на Лист0
Sub goToSheet0()
  ThisWorkbook.Sheets("Лист0").Select
End Sub

' Переход на Лист1 (DB)
Sub goToSheet1()
  ThisWorkbook.Sheets("Лист1").Select
End Sub

' Переход на Лист3 (активы)
Sub goToSheet3()
  ThisWorkbook.Sheets("Лист3").Select
  ThisWorkbook.Sheets("Лист3").Range("A1").Select
End Sub

' Переход на Лист4 (вклады)
Sub goToSheet4()
  ThisWorkbook.Sheets("Лист4").Select
  ThisWorkbook.Sheets("Лист4").Range("A1").Select
End Sub

' Переход на Лист4 (ИСЖ/НСЖ)
Sub goToSheet4_ИСЖ_НСЖ()
  ThisWorkbook.Sheets("Лист4").Select
  ThisWorkbook.Sheets("Лист4").Range("L14").Select
  '
  ActiveWindow.SmallScroll Down:=11
End Sub

' Переход на Лист5 (карты)
Sub goToSheet5()
  ThisWorkbook.Sheets("Лист5").Select
  ThisWorkbook.Sheets("Лист5").Range("A1").Select
End Sub

' Переход на Capacity (старый Капасити на Лист6)
Sub goToSheet6()
  ThisWorkbook.Sheets("Capacity").Select
  ThisWorkbook.Sheets("Capacity").Range("A1").Select
End Sub

' UpdFr_DB
Sub goToSheet_UpdFr_DB()
  ThisWorkbook.Sheets("UpdFr_DB").Select
  ThisWorkbook.Sheets("UpdFr_DB").Range("A1").Select
End Sub

' Динамика
Sub goToSheet_Динамика()
  ThisWorkbook.Sheets("Динамика").Select
  ThisWorkbook.Sheets("Динамика").Range("A1").Select
End Sub


' Переход на Лист7 (Интегральный рейтинг по сотрудникам)
Sub goToSheet7()
  ThisWorkbook.Sheets("Лист7").Select
  ThisWorkbook.Sheets("Лист7").Range("A1").Select
End Sub

' Переход на Лист12
Sub goToSheet12()
  ' ThisWorkbook.Sheets("Лист12").Select
  ' ThisWorkbook.Sheets("Лист12").Range("A1").Select
  ' ThisWorkbook.Sheets("Лист12").Range("A72").Select
  ' ThisWorkbook.Sheets("Лист12").Range("A40").Select

  ThisWorkbook.Sheets("Лист12").Select
  ThisWorkbook.Sheets("Лист12").Range("A1").Select
  row_Форма_12_4 = rowByValue(ThisWorkbook.Name, "Лист12", "Форма 12.4", 100, 100)
  ' Перемещаемся
  ActiveWindow.SmallScroll Down:=row_Форма_12_4 - 2


End Sub

' Переход на Лист13
Sub goToSheet13()
  ThisWorkbook.Sheets("Лист13").Select
End Sub

' Переход на Лист КроссЗП (Проникновение в ЗП)
Sub goToSheetКроссЗП()
  ThisWorkbook.Sheets("КроссЗП").Select
  ThisWorkbook.Sheets("КроссЗП").Range("A1").Select
End Sub

' Переход на Лист16 (Просрочки CRM)
Sub goToSheet16()
  ThisWorkbook.Sheets("Лист16").Select
  ThisWorkbook.Sheets("Лист16").Range("A1").Select
End Sub

' Переход на Лист ЕСУП
Sub goToSheetЕСУП()
  ThisWorkbook.Sheets("ЕСУП").Select
  ThisWorkbook.Sheets("ЕСУП").Range("A1").Select
End Sub

' Переход на Лист ЗПК
Sub goToSheetЗПК()
  ThisWorkbook.Sheets("ЗПК").Select
End Sub

' Переход на Лист Остатки ЗПК
Sub goToSheetОстаткиЗПК()
  ThisWorkbook.Sheets("Остатки ЗПК").Select
End Sub

' Переход на Лист PL
Sub goToSheetPL()
  ThisWorkbook.Sheets("PL").Select
End Sub

' Переход на Лист СМОТ
Sub goToSheetСМОТ()
  ThisWorkbook.Sheets("СМОТ").Select
End Sub

' Переход на Лист ' Отчётность по входящему потоку с PA
Sub goToSheetВх_PA()
  ThisWorkbook.Sheets("Вх_PA").Select
End Sub

' Переход на Лист Ипотека
Sub goToSheetИпотека()
  ThisWorkbook.Sheets("Ипотека").Select
End Sub


' Переход на Лист Подотчет
Sub goToSheetПодотчет()
  ThisWorkbook.Sheets("Подотчет").Select
End Sub

' Переход на Лист План
Sub goToSheetПлан()
  ThisWorkbook.Sheets("План").Select
End Sub

' Переход на Лист10 (Банкострахование)
Sub goToSheetЛист10()
  ThisWorkbook.Sheets("Лист10").Select
End Sub

' Переход на Лист9 (Воронка)
Sub goToSheetЛист9()
  ThisWorkbook.Sheets("Лист9").Select
End Sub

' Переход на КК_ЗП
Sub goToSheetКК_ЗП()
  ThisWorkbook.Sheets("КК_ЗП").Select
End Sub

' Переход на Обучения
Sub goToSheetОбучения()
  ThisWorkbook.Sheets("Обучения").Select
End Sub

' Переход на Addr.Book
Sub goToSheetAddrBook()
  ThisWorkbook.Sheets("Addr.Book").Select
End Sub

' Переход на To-Do
Sub goToSheetToDo()
  ThisWorkbook.Sheets("To-Do").Select
End Sub

' Переход на Лист8 (ИПЗ Управляющих)
Sub goToSheetЛист8()
  ThisWorkbook.Sheets("Лист8").Select
End Sub

' Переход
Sub ЕСУП_к_началу()
  ' ЕСУП_к_началу Макрос
  Range("A1").Select
End Sub

Sub Переход_на_Тюменский()
  
  ' Переход_на_Тюменский Макрос
  Range("A33").Select
  ActiveWindow.SmallScroll Down:=30
  ' ActiveWindow.SmallScroll ToRight:=8

End Sub
Sub Переход_на_Сургутский()
Attribute Переход_на_Сургутский.VB_ProcData.VB_Invoke_Func = " \n14"
  
  ' Переход_на_Сургутский Макрос
  Range("R33").Select
  ActiveWindow.SmallScroll Down:=30 ' 9 ' было 15
  ActiveWindow.SmallScroll ToRight:=17 ' 8 ' было 15

End Sub
Sub Переход_на_Нижневартовский()
Attribute Переход_на_Нижневартовский.VB_ProcData.VB_Invoke_Func = " \n14"
  
  ' Переход_на_Нижневартовский Макрос
  Range("AI33").Select
  ActiveWindow.SmallScroll Down:=30 ' 9 ' было 15
  ActiveWindow.SmallScroll ToRight:=11

End Sub
Sub Переход_на_Новоуренгойский()
Attribute Переход_на_Новоуренгойский.VB_ProcData.VB_Invoke_Func = " \n14"

  ' Переход_на_Новоуренгойский Макрос
  Range("AZ33").Select
  ActiveWindow.SmallScroll Down:=30 ' 9 ' было 15
  ActiveWindow.SmallScroll ToRight:=11

End Sub

Sub Переход_на_ТаркоСале()
Attribute Переход_на_ТаркоСале.VB_ProcData.VB_Invoke_Func = " \n14"

  ' Переход_на_ТаркоСале Макрос
  Range("BQ33").Select
  ActiveWindow.SmallScroll Down:=30 ' 9 ' было 15
  ActiveWindow.SmallScroll ToRight:=11

End Sub

Sub Переход_на_ИЦ()

  ' Переход_на_ИЦ Макрос
  Range("CH33").Select
  ActiveWindow.SmallScroll Down:=30 '9 ' было 15
  ActiveWindow.SmallScroll ToRight:=12

End Sub

Sub Переход_на_ОКП()

  ' Переход_на_ОКП Макрос
  Range("CY33").Select
  ActiveWindow.SmallScroll Down:=30 ' было 15
  ActiveWindow.SmallScroll ToRight:=12
  
End Sub


Sub Переход_на_ЕСУП_подразделений()
  
  ' Переход на ЕСУП подразделений
  Range("A33").Select
  ActiveWindow.SmallScroll Down:=0
  ActiveWindow.SmallScroll ToRight:=22

End Sub


' Поручения офису на неделю: ЗКК1, ЗДК1
Sub ПорученияНаНеделюi(In_Office, In_К_пор, In_Value)
  ' Номера офисов: 1- Тюмень, 2 - Сургут, 3 - Нижневартовск, 4 - Новый Уренгой 5 - Тарко-Сале
  ' ЗКК1
  
  ' ЗДК1
  
End Sub

' Если нет обновленного DB
Sub notActualDB()
  ThisWorkbook.Sheets("Лист0").Cells(7, 4).Value = "1) На " + CStr(Date) + " нет обновленного DashBoard"
  ' Зачеркиваем пункт меню на стартовой страницы
  ' Call ЗачеркиваемТекстВячейке("Лист0", "D7")
  Call ЗачеркиваемТекстВячейке("Лист0", RangeByValue(ThisWorkbook.Name, "Лист0", "DashBoard (при наличии)", 100, 100))
  
  Call goToSheet0
End Sub

Sub УбратьФильтрНаАктивномЛисте()
    ActiveSheet.ShowAllData
    Range("A1").Select
End Sub

Sub Переход_на_Повестка_дня()
  ThisWorkbook.Sheets("ЕСУП").Select
  Range("A60").Select
  ActiveWindow.SmallScroll Down:=17
End Sub

Sub Переход_DB_ЗП()
  ThisWorkbook.Sheets("DB_ЗП").Select
  Range("A1").Select
  ' ActiveWindow.SmallScroll Down:=17
End Sub

Sub Переход_ДР()
  ThisWorkbook.Sheets("ДР").Select
  Range("A1").Select
  ' ActiveWindow.SmallScroll Down:=17
End Sub


' Создание хэштега для новостей на Листе0
Sub createHashTag_n()
Dim Row_Включить_в_Собрание, Column_Включить_в_Собрание As Byte

  ' "Включить в Собрание "Повестка_дня":"
  Row_Включить_в_Собрание = rowByValue(ThisWorkbook.Name, "Лист0", "Включить в Собрание " + Chr(34) + "Повестка_дня" + Chr(34) + ":", 100, 100)
  Column_Включить_в_Собрание = ColumnByValue(ThisWorkbook.Name, "Лист0", "Включить в Собрание " + Chr(34) + "Повестка_дня" + Chr(34) + ":", 100, 100)
  
  ' Вносим Хэштэг
  ThisWorkbook.Sheets("Лист0").Cells(Row_Включить_в_Собрание + 2, Column_Включить_в_Собрание + 14).Value = createHashTag("n")

End Sub

' ЕСУП - включить вопрос в повестку собрания
Sub Включить_в_Собрание_Повестка_дня()
Dim НомерСтроки, сдвиг_в_право, Row_Включить_в_Собрание, Column_Включить_в_Собрание As Byte

  ' Для того, чтобы тернироваться на создании протокола
  ' сдвиг_в_право = 15

  ' "Включить в Собрание "Повестка_дня":"
  Row_Включить_в_Собрание = rowByValue(ThisWorkbook.Name, "Лист0", "Включить в Собрание " + Chr(34) + "Повестка_дня" + Chr(34) + ":", 100, 100)
  Column_Включить_в_Собрание = ColumnByValue(ThisWorkbook.Name, "Лист0", "Включить в Собрание " + Chr(34) + "Повестка_дня" + Chr(34) + ":", 100, 100)
  
  If Trim(ThisWorkbook.Sheets("Лист0").Cells(Row_Включить_в_Собрание + 2, Column_Включить_в_Собрание).Value) <> "" Then
  
    ' Генерируем свежий Хэштэг
    If ThisWorkbook.Sheets("Лист0").Cells(Row_Включить_в_Собрание + 2, Column_Включить_в_Собрание + 14).Value = "" Then
      If MsgBox("Сгенерировать Хэштэг?", vbYesNo) = vbYes Then
        Call createHashTag_n
      End If
    End If
    
    ' Вставить строку в Повестку собрания
    Call Вставка_строки_в_Повестку("Прощаев С.Ф.", _
                                     delSym(ThisWorkbook.Sheets("Лист0").Cells(Row_Включить_в_Собрание + 2, Column_Включить_в_Собрание).Value), _
                                       ThisWorkbook.Sheets("Лист0").Cells(Row_Включить_в_Собрание + 2, Column_Включить_в_Собрание + 14).Value)

    ' НомерСтроки = ThisWorkbook.Sheets("ЕСУП").Range(RangeByValue(ThisWorkbook.Name, "ЕСУП", "Повестка_дня", 100, 100)).Row
    ' Заносим на Лист "ЕСУП"
    ' i = 2
    ' Номер_вопроса = 0
    ' Do While ThisWorkbook.Sheets("ЕСУП").Cells(НомерСтроки + i, 1 + сдвиг_в_право).Value <> ""
    '   Номер_вопроса = Номер_вопроса + 1
    '   i = i + 1
    ' Loop
    ' Номер_вопроса = Номер_вопроса + 1
    ' ThisWorkbook.Sheets("ЕСУП").Cells(НомерСтроки + i, 1 + сдвиг_в_право).Value = Номер_вопроса
    ' ThisWorkbook.Sheets("ЕСУП").Cells(НомерСтроки + i, 2 + сдвиг_в_право).Value = "Прощаев С.Ф."
    ' ThisWorkbook.Sheets("ЕСУП").Cells(НомерСтроки + i, 3 + сдвиг_в_право).Value = delSym(ThisWorkbook.Sheets("Лист0").Cells(Row_Включить_в_Собрание + 2, Column_Включить_в_Собрание).Value)
    ' ThisWorkbook.Sheets("ЕСУП").Cells(НомерСтроки + i, 13 + сдвиг_в_право).Value = ThisWorkbook.Sheets("Лист0").Cells(Row_Включить_в_Собрание + 2, Column_Включить_в_Собрание + 14).Value
    ' ThisWorkbook.Sheets("ЕСУП").Cells(НомерСтроки + i, 14 + сдвиг_в_право).Value = 0
  
  
  
    ' Очищаем поле ввода на Листе0 B13
    ThisWorkbook.Sheets("Лист0").Cells(Row_Включить_в_Собрание + 2, Column_Включить_в_Собрание).Value = ""
    ' Очищаем Хэштэг
    ThisWorkbook.Sheets("Лист0").Cells(Row_Включить_в_Собрание + 2, Column_Включить_в_Собрание + 14).Value = ""
    
  End If
  
End Sub

' Вставить строку в Повестку собрания
Sub Вставка_строки_в_Повестку(In_FIO, In_Question, In_Tag)
Dim НомерСтроки As Byte


    НомерСтроки = ThisWorkbook.Sheets("ЕСУП").Range(RangeByValue(ThisWorkbook.Name, "ЕСУП", "Повестка_дня", 100, 100)).Row
    ' Заносим на Лист "ЕСУП"
    i = 2
    Номер_вопроса = 0
    Do While ThisWorkbook.Sheets("ЕСУП").Cells(НомерСтроки + i, 1 + сдвиг_в_право).Value <> ""
      Номер_вопроса = Номер_вопроса + 1
      i = i + 1
    Loop
    
    Номер_вопроса = Номер_вопроса + 1
    ThisWorkbook.Sheets("ЕСУП").Cells(НомерСтроки + i, 1 + сдвиг_в_право).Value = Номер_вопроса
    ' Выступающий
    ThisWorkbook.Sheets("ЕСУП").Cells(НомерСтроки + i, 2 + сдвиг_в_право).Value = In_FIO ' "Прощаев С.Ф."
    ' Вопрос
    ThisWorkbook.Sheets("ЕСУП").Cells(НомерСтроки + i, 3 + сдвиг_в_право).Value = In_Question ' delSym(ThisWorkbook.Sheets("Лист0").Cells(Row_Включить_в_Собрание + 2, Column_Включить_в_Собрание).Value)
    ' Хэштег
    ThisWorkbook.Sheets("ЕСУП").Cells(НомерСтроки + i, 13 + сдвиг_в_право).Value = In_Tag ' ThisWorkbook.Sheets("Лист0").Cells(Row_Включить_в_Собрание + 2, Column_Включить_в_Собрание + 14).Value
    ' Выступление состоялось 0/1
    ThisWorkbook.Sheets("ЕСУП").Cells(НомерСтроки + i, 14 + сдвиг_в_право).Value = 0


End Sub


' Перенумеровать список в Addr.Book
Sub Перенумеровать_список_AddrBook()
  
End Sub

' Напоминание: прислать отчет по отработке вкладов
Sub toSend_Прислать_отчет_по_отработке_вкладов()
Dim Range_str, Адрес_Ln_кому, Адрес_Ln_копия As String
Dim Range_Row, Range_Column, i As Byte
Dim темаПисьма, текстПисьма, hashTag As String

  ' Находим ячейку
  Range_str = RangeByValue(ThisWorkbook.Name, "Лист0", "Напоминание: прислать отчет по отработке вкладов", 100, 100)
  Range_Row = Workbooks(ThisWorkbook.Name).Sheets("Лист0").Range(Range_str).Row
  Range_Column = Workbooks(ThisWorkbook.Name).Sheets("Лист0").Range(Range_str).Column
  
  ' Проверка - отправлялось или нет
  If ThisWorkbook.Sheets("Лист0").Cells(Range_Row, Range_Column + 6).Value = "" Then
  
    ' Запрос
    If MsgBox("Отправить напоминание: прислать себе шаблон запроса отчета по отработке вкладов?", vbYesNo) = vbYes Then
    
      ' Отправка сообщения
      ' Адрес_Ln_кому = getFromAddrBook("УДО2,УДО3,УДО4,УДО5,НОРПиКО1,НОРПиКО2,НОРПиКО3,НОРПиКО4,НОРПиКО5", 2)
      ' Адрес_Ln_копия = getFromAddrBook("РД", 2)
      Адрес_Ln_кому = "Sergey Fedorovich Proschaev/Tyumen/PSBank/Ru"
      Адрес_Ln_копия = ""
      
      ' Тема письма - Тема:
      темаПисьма = "Отчет по отработке вкладов за прошлую неделю"
      ' hashTag - Хэштэг:
      hashTag = "#depositsfinish"
      ' Текст письма
      текстПисьма = "" + Chr(13)
      текстПисьма = текстПисьма + "Адреса получателей: " + getFromAddrBook("УДО2,УДО3,УДО4,УДО5,НОРПиКО1,НОРПиКО2,НОРПиКО3,НОРПиКО4,НОРПиКО5", 2) + Chr(13)
      текстПисьма = текстПисьма + "Копия: " + getFromAddrBook("РД", 2) + Chr(13)
      текстПисьма = текстПисьма + "" + Chr(13)
      текстПисьма = текстПисьма + "Уважаемые руководители," + Chr(13)
      текстПисьма = текстПисьма + "" + Chr(13)
      текстПисьма = текстПисьма + "Прошу направить информацию по отработке вкладчиков за прошлую неделю." + Chr(13)
      текстПисьма = текстПисьма + "" + Chr(13)
      текстПисьма = текстПисьма + "" + Chr(13)
      ' Визитка (подпись С Ув., )
      текстПисьма = текстПисьма + ПодписьВПисьме()
      ' Хэштег
      текстПисьма = текстПисьма + createBlankStr(35) + hashTag
      
      ' Вызов
      ' Call send_Lotus_Notes(темаПисьма, Адрес_Ln_кому, Адрес_Ln_копия, текстПисьма, "")
      
      ' Вызов с отправкой себе в скрытой копии
      Call send_Lotus_Notes2(темаПисьма, Адрес_Ln_кому, Адрес_Ln_копия, "Sergey Fedorovich Proschaev/Tyumen/PSBank/Ru", текстПисьма, "")
  
      ' Сообщение
      MsgBox ("Письмо отправлено!")

      ' Вставляем дату и вермя в ячейку "Напоминание: прислать отчет по отработке вкладов" + 6
      ThisWorkbook.Sheets("Лист0").Cells(Range_Row, Range_Column + 6).Value = Now
      
      ' Call ЗачеркиваемТекстВячейке("Лист0", "D9")
      Call ЗачеркиваемТекстВячейке("Лист0", RangeByValue(ThisWorkbook.Name, "Лист0", "Напоминание: прислать отчет по отработке вкладов", 100, 100))

    End If
  Else
    MsgBox ("Внимание! Сообщение уже было отправлено в " + CStr(ThisWorkbook.Sheets("Лист0").Cells(Range_Row, Range_Column + 6).Value) + "!")
  End If
  
End Sub

' Напоминание: прислать долги по списанию карт и досье
Sub toSend_Прислать_долги_по_картам_и_досье()
  
Dim Range_str, Адрес_Ln_кому, Адрес_Ln_копия As String
Dim Range_Row, Range_Column, i As Byte
Dim темаПисьма, текстПисьма, hashTag As String

  ' Находим ячейку
  Range_str = RangeByValue(ThisWorkbook.Name, "Лист0", "Задолженность по списанию карт и КД", 100, 100)
  Range_Row = Workbooks(ThisWorkbook.Name).Sheets("Лист0").Range(Range_str).Row
  Range_Column = Workbooks(ThisWorkbook.Name).Sheets("Лист0").Range(Range_str).Column
  
  ' Проверка - отправлялось или нет
  If ThisWorkbook.Sheets("Лист0").Cells(Range_Row, Range_Column + 6).Value = "" Then
  
    ' Запрос
    If MsgBox("Отправить себе шаблон: прислать долги по списанию карт и КД?", vbYesNo) = vbYes Then
    
      ' Отправка сообщения
      ' Адрес_Ln_кому = "Vera Sizikova/Tyumen/PSBank/Ru"
      ' Адрес_Ln_копия = "Alla Cherneckaya/Tyumen/PSBank/Ru"
      
      Адрес_Ln_кому = "Sergey Fedorovich Proschaev/Tyumen/PSBank/Ru"
      Адрес_Ln_копия = ""
      
      
      ' Тема письма - Тема:
      темаПисьма = "Долги офисов на " + CStr(Date) + " г."
      ' hashTag - Хэштэг:
      hashTag = "#долгиофисов"
      ' Текст письма
      текстПисьма = "" + Chr(13)
      текстПисьма = текстПисьма + "Адреса получателей: " + getFromAddrBook("УДО2,УДО3,УДО4,УДО5,НОРПиКО1,НОРПиКО2,НОРПиКО3,НОРПиКО4,НОРПиКО5", 2) + Chr(13)
      текстПисьма = текстПисьма + "Копия: " + getFromAddrBook("РД", 2) + Chr(13)
      текстПисьма = текстПисьма + "" + Chr(13)
      текстПисьма = текстПисьма + "Добрый день!" + Chr(13)
      текстПисьма = текстПисьма + "" + Chr(13)
      текстПисьма = текстПисьма + "Прошу направить информацию просроченным долгам офисов по списанию карт и исправлению ошибок в КД на " + CStr(Date) + " г. для постановки на контроль." + Chr(13)
      текстПисьма = текстПисьма + "" + Chr(13)
      текстПисьма = текстПисьма + "" + Chr(13)
      ' Визитка (подпись С Ув., )
      текстПисьма = текстПисьма + ПодписьВПисьме()
      ' Хэштег
      текстПисьма = текстПисьма + createBlankStr(35) + hashTag
      
      ' Вызов с отправкой себе в скрытой копии
      Call send_Lotus_Notes2(темаПисьма, Адрес_Ln_кому, Адрес_Ln_копия, "Sergey Fedorovich Proschaev/Tyumen/PSBank/Ru", текстПисьма, "")
  
      ' Сообщение
      MsgBox ("Письмо отправлено!")

      ' Вставляем дату и вермя в ячейку "Напоминание: прислать отчет по отработке вкладов" + 6
      ThisWorkbook.Sheets("Лист0").Cells(Range_Row, Range_Column + 6).Value = Now
      
      ' Call ЗачеркиваемТекстВячейке("Лист0", "D9")
      Call ЗачеркиваемТекстВячейке("Лист0", RangeByValue(ThisWorkbook.Name, "Лист0", "Задолженность по списанию карт и КД", 100, 100))

    End If
  Else
    MsgBox ("Внимание! Сообщение уже было отправлено в " + CStr(ThisWorkbook.Sheets("Лист0").Cells(Range_Row, Range_Column + 6).Value) + "!")
  End If
  
  
End Sub

' Напоминание: прислать прогнозы на неделю
Sub toSend_Прислать_Прогнозы_на_неделю()
'
Dim Range_str, Адрес_Ln_кому, Адрес_Ln_копия As String
Dim Range_Row, Range_Column, i As Byte
Dim темаПисьма, текстПисьма, hashTag As String

  ' Находим ячейку
  Range_str = RangeByValue(ThisWorkbook.Name, "Лист0", "Напоминание: прислать прогнозы на неделю", 100, 100)
  Range_Row = Workbooks(ThisWorkbook.Name).Sheets("Лист0").Range(Range_str).Row
  Range_Column = Workbooks(ThisWorkbook.Name).Sheets("Лист0").Range(Range_str).Column
  
  ' Проверка - отправлялось или нет
  If ThisWorkbook.Sheets("Лист0").Cells(Range_Row, Range_Column + 6).Value = "" Then
  
    ' Запрос
    If MsgBox("Отправить себе шаблон: Напоминание: прислать прогнозы на неделю?", vbYesNo) = vbYes Then
    
      ' Отправка сообщения
      ' Адрес_Ln_кому = getFromAddrBook("УДО2,УДО3,УДО4,УДО5,НОРПиКО1,НОРПиКО2,НОРПиКО3,НОРПиКО4,НОРПиКО5", 2)
      ' Адрес_Ln_копия = getFromAddrBook("РД", 2)
      
      Адрес_Ln_кому = "Sergey Fedorovich Proschaev/Tyumen/PSBank/Ru"
      Адрес_Ln_копия = ""
      
      
      ' Тема письма - Тема:
      темаПисьма = "Прогнозы продаж"
      ' hashTag - Хэштэг:
      hashTag = "#прогнозы"
      ' Текст письма
      текстПисьма = "" + Chr(13)
      текстПисьма = текстПисьма + "Адреса получателей: " + getFromAddrBook("УДО2,УДО3,УДО4,УДО5,НОРПиКО1,НОРПиКО2,НОРПиКО3,НОРПиКО4,НОРПиКО5", 2) + Chr(13)
      текстПисьма = текстПисьма + "Копия: " + getFromAddrBook("РД", 2) + Chr(13)
      текстПисьма = текстПисьма + "" + Chr(13)
      текстПисьма = текстПисьма + "Уважаемые руководители," + Chr(13)
      текстПисьма = текстПисьма + "" + Chr(13)
      текстПисьма = текстПисьма + "Прошу сегодня до 13:00 направить прогнозы продаж в офисах в период с " + CStr(strDDMM(weekStartDate(Date))) + " по " + CStr(weekEndDate(Date)) + " г." + Chr(13)
      текстПисьма = текстПисьма + "" + Chr(13)
      текстПисьма = текстПисьма + "" + Chr(13)
      ' Визитка (подпись С Ув., )
      текстПисьма = текстПисьма + ПодписьВПисьме()
      ' Хэштег
      текстПисьма = текстПисьма + createBlankStr(35) + hashTag
      
      ' Вызов с отправкой себе в скрытой копии
      Call send_Lotus_Notes2(темаПисьма, Адрес_Ln_кому, Адрес_Ln_копия, "Sergey Fedorovich Proschaev/Tyumen/PSBank/Ru", текстПисьма, "")
  
      ' Сообщение
      MsgBox ("Письмо отправлено!")

      ' Вставляем дату и вермя в ячейку "Напоминание: прислать отчет по отработке вкладов" + 6
      ThisWorkbook.Sheets("Лист0").Cells(Range_Row, Range_Column + 6).Value = Now
      
      ' Call ЗачеркиваемТекстВячейке("Лист0", "D9")
      Call ЗачеркиваемТекстВячейке("Лист0", RangeByValue(ThisWorkbook.Name, "Лист0", "Напоминание: прислать прогнозы на неделю", 100, 100))

    End If
  
  Else
    MsgBox ("Внимание! Сообщение уже было отправлено в " + CStr(ThisWorkbook.Sheets("Лист0").Cells(Range_Row, Range_Column + 6).Value) + "!")
  End If
  
  
End Sub

' Создать копию в Лотусе
Sub createBackUp_withMsgBox()
    ' Запрос
    If MsgBox("Создать BackUp и отправить в почте?", vbYesNo) = vbYes Then
      
      ' Вызов процедуры
      Call createBackUp
      
      ' Сообщение
      MsgBox ("Письмо с копией " + Dir(backUpFileName) + " отправлено!")

    End If

End Sub

' Создать копию в Лотусе
Sub createBackUp()
Dim attachmentFile As String
      
      ' Сохранение изменений
      ThisWorkbook.Save
    
      ' Имя BackUp файла
      backUpFileName = ThisWorkbook.Path + "\BackUp\DB_Result_" + strДД_MM_YYYY(Date) + ".xlsm"
    
      ' Строка статуса
      Application.StatusBar = "BackUp: Создание копии " + Dir(backUpFileName) + "..."
    
      ' Сохраняем текущие изменения в книге
      ThisWorkbook.SaveCopyAs FileName:=backUpFileName

      ' Запускаем архиватор этого файла
      ' Shell ("C:\Program Files\7-Zip\7z a -tzip -ssw -mx9 C:\Users\PROSCHAEVSF\Documents\#DB_Result\db_result.zip C:\Users\PROSCHAEVSF\Documents\#DB_Result")

      ' Строка статуса
      Application.StatusBar = "BackUp: Копия " + Dir(backUpFileName) + "создана!"

      ' Создание почтового сообщения
      ' Строка статуса
      Application.StatusBar = "BackUp: Подготовка сообщения LotusNotes..."
      ' Файл-вложение
      attachmentFile = backUpFileName
      ' Параметры сообщения
      Адрес_Ln_кому = "Sergey Fedorovich Proschaev/Tyumen/PSBank/Ru"
      Адрес_Ln_копия = ""
      ' Тема письма - Тема:
      темаПисьма = "BackUp DB_Result от " + strДД_MM_YYYY(Date) + " г."
      ' hashTag - Хэштэг:
      hashTag = "#BackUp #BackUp_DB_Result_" + strDDMMYYYY(Date) + " #формаотчета"
      ' Текст письма
      текстПисьма = "" + Chr(13)
      текстПисьма = текстПисьма + "Копия DB_Result на " + CStr(Date) + " г." + Chr(13)
      текстПисьма = текстПисьма + "" + Chr(13)
      текстПисьма = текстПисьма + "" + Chr(13)
      ' Визитка (подпись С Ув., )
      текстПисьма = текстПисьма + ПодписьВПисьме()
      ' Хэштег
      текстПисьма = текстПисьма + createBlankStr(35) + hashTag
      ' Строка статуса
      Application.StatusBar = "BackUp: Отправка сообщения в LotusNotes..."
      Call send_Lotus_Notes(темаПисьма, "Sergey Fedorovich Proschaev/Tyumen/PSBank/Ru", "Sergey Fedorovich Proschaev/Tyumen/PSBank/Ru", текстПисьма, attachmentFile)
      ' Строка статуса
      Application.StatusBar = "BackUp: Сообщение в LotusNotes отправлено!"
                
      ' Строка статуса
      Application.StatusBar = "BackUp: Создание копии завершено"

    
End Sub
