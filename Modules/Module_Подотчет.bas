Attribute VB_Name = "Module_Подотчет"
' Лист "Подотчет"
Sub Сформировать_Авансовый_отчет()

Dim Ячейка_Представительские_расходы, expenseReportName, Строка_участников_мероприятия As String
Dim НомерСтроки_Представительские_расходы, НомерСтолбца_Представительские_расходы, i As Byte
Dim expenseReportDate As Date
Dim Строка_Подотчета_row As Integer

  ' Определяем, где находится текущая ячейка. Должен быть диапазон A62:N90 (в относительных от "Повестка_дня" координатах)
  Ячейка_Представительские_расходы = RangeByValue(ThisWorkbook.Name, "Подотчет", "Представительские расходы", 100, 100)
  НомерСтроки_Представительские_расходы = ThisWorkbook.Sheets("Подотчет").Range(Ячейка_Представительские_расходы).Row
  НомерСтолбца_Представительские_расходы = ThisWorkbook.Sheets("Подотчет").Range(Ячейка_Представительские_расходы).Column
  
  ' Проверка диапазона
  If (ActiveCell.Row >= НомерСтроки_Представительские_расходы + 4) And (ActiveCell.Row <= НомерСтроки_Представительские_расходы + 20) And (ActiveCell.Column >= НомерСтолбца_Представительские_расходы - 1) And ((ActiveCell.Column <= НомерСтолбца_Представительские_расходы + 30)) And (ThisWorkbook.Sheets("Подотчет").Cells(ActiveCell.Row, НомерСтолбца_Представительские_расходы + 1).Value <> "") Then
    
    ' Запрос на формирование
    If MsgBox("Сформировать Авансовый отчет за " + CStr(ThisWorkbook.Sheets("Подотчет").Cells(ActiveCell.Row, НомерСтолбца_Представительские_расходы).Value) + "?", vbYesNo) = vbYes Then
      
      ' Строка
      Строка_Подотчета_row = ActiveCell.Row
      
      ' Дата авансового отчета
      expenseReportDate = CDate(ThisWorkbook.Sheets("Подотчет").Cells(ActiveCell.Row, НомерСтолбца_Представительские_расходы).Value)
      
      ' Открываем шаблон Протокола из C:\Users\...\Documents\#VBA\DB_Result\Templates
      Workbooks.Open (ThisWorkbook.Path + "\Templates\Авансовый отчет.xlsx")
      
      ' Имя файла
      expenseReportName = "Авансовый отчет Акт Смета за " + ИмяМесяцаГод(expenseReportDate) + " (представительские расходы ДРПКК ст.1637)" + ".xlsx"
      
      ' Сохраняем Файл с Авансовым отчетом
      Workbooks("Авансовый отчет.xlsx").SaveAs FileName:=ThisWorkbook.Path + "\Out\" + expenseReportName, FileFormat:=xlOpenXMLWorkbook, createBackUp:=False
      
      ' Заполняем поля
    
      ' Руководитель
      Workbooks(expenseReportName).Sheets("Авансовый отчет").Range("AV13").Value = "Региональный директор"
      ' Workbooks(expenseReportName).Sheets("Авансовый отчет").Range("AV13").Value = "Заместитель регионального директора по развитию розничного бизнеса"
      Workbooks(expenseReportName).Sheets("Акт").Range("G4").Value = Workbooks(expenseReportName).Sheets("Авансовый отчет").Range("AV13").Value
      Workbooks(expenseReportName).Sheets("Смета").Range("G4").Value = Workbooks(expenseReportName).Sheets("Авансовый отчет").Range("AV13").Value
      
      ' ФИО Руководителя
      Workbooks(expenseReportName).Sheets("Авансовый отчет").Range("AX15").Value = "Шевелев А.Ю."
      ' Workbooks(expenseReportName).Sheets("Авансовый отчет").Range("AX15").Value = "Прощаев С.Ф."
      Workbooks(expenseReportName).Sheets("Акт").Range("I6").Value = Workbooks(expenseReportName).Sheets("Авансовый отчет").Range("AX15").Value
      Workbooks(expenseReportName).Sheets("Смета").Range("I6").Value = Workbooks(expenseReportName).Sheets("Авансовый отчет").Range("AX15").Value
      
      ' Дата
      Workbooks(expenseReportName).Sheets("Акт").Range("D10").Value = "от " + CStr(expenseReportDate) + " г."
      Workbooks(expenseReportName).Sheets("Акт").Range("B16").Value = CStr(expenseReportDate) + " г."
      Workbooks(expenseReportName).Sheets("Смета").Range("D14").Value = CStr(expenseReportDate) + " г."
      
      ' Месяц
      Workbooks(expenseReportName).Sheets("Авансовый отчет").Range("AS17").Value = ИмяМесяца2(expenseReportDate)
      
      ' Год
      Workbooks(expenseReportName).Sheets("Авансовый отчет").Range("BC17").Value = CStr(Year(expenseReportDate))
      
      ' Назначение аванса
      Workbooks(expenseReportName).Sheets("Авансовый отчет").Range("AR24").Value = "представительские расходы"
      
      ' Итого получено
      Workbooks(expenseReportName).Sheets("Авансовый отчет").Range("AH33").Value = ThisWorkbook.Sheets("Подотчет").Cells(Строка_Подотчета_row, НомерСтолбца_Представительские_расходы + 9).Value
      Workbooks(expenseReportName).Sheets("Авансовый отчет").Range("AH35").Value = Workbooks(expenseReportName).Sheets("Авансовый отчет").Range("AH33").Value
      
      ' Израсходовано
      Workbooks(expenseReportName).Sheets("Авансовый отчет").Range("AH36").Value = ThisWorkbook.Sheets("Подотчет").Cells(Строка_Подотчета_row, НомерСтолбца_Представительские_расходы + 31).Value
      
      ' Остаток AH37 или перерасход AH38
      If (Workbooks(expenseReportName).Sheets("Авансовый отчет").Range("AH33").Value - Workbooks(expenseReportName).Sheets("Авансовый отчет").Range("AH36").Value) >= 0 Then
        ' остаток AH37
        Workbooks(expenseReportName).Sheets("Авансовый отчет").Range("AH37").Value = (Workbooks(expenseReportName).Sheets("Авансовый отчет").Range("AH33").Value - Workbooks(expenseReportName).Sheets("Авансовый отчет").Range("AH36").Value)
        
      Else
        ' перерасход AH38
        Workbooks(expenseReportName).Sheets("Авансовый отчет").Range("AH38").Value = Abs((Workbooks(expenseReportName).Sheets("Авансовый отчет").Range("AH33").Value - Workbooks(expenseReportName).Sheets("Авансовый отчет").Range("AH36").Value))
        
      End If
            
      ' Лист "Авансовый отчет (оборот)"
      ' Обработка отчета
      For i = 1 To 5

        If ThisWorkbook.Sheets("Подотчет").Cells(Строка_Подотчета_row, НомерСтолбца_Представительские_расходы + 18 + (3 * (i - 1))).Value <> 0 Then

          ' Номер строки в Авансовом отчете
          Workbooks(expenseReportName).Sheets("Авансовый отчет (оборот)").Cells(6 + i, 1).Value = CStr(i)
        
          ' Документы, подтверждающие расходы: Дата1. с 16
          Workbooks(expenseReportName).Sheets("Авансовый отчет (оборот)").Cells(6 + i, 4).Value = ThisWorkbook.Sheets("Подотчет").Cells(Строка_Подотчета_row, НомерСтолбца_Представительские_расходы + 16 + (3 * (i - 1))).Value
      
          ' Документы, подтверждающие расходы: Номер1. с 17
          Workbooks(expenseReportName).Sheets("Авансовый отчет (оборот)").Cells(6 + i, 11).Value = ThisWorkbook.Sheets("Подотчет").Cells(Строка_Подотчета_row, НомерСтолбца_Представительские_расходы + 17 + (3 * (i - 1))).Value
            
          ' Документы, подтверждающие расходы: Сумма1. с 18
          Workbooks(expenseReportName).Sheets("Авансовый отчет (оборот)").Cells(6 + i, 29).Value = ThisWorkbook.Sheets("Подотчет").Cells(Строка_Подотчета_row, НомерСтолбца_Представительские_расходы + 18 + (3 * (i - 1))).Value
          Workbooks(expenseReportName).Sheets("Авансовый отчет (оборот)").Cells(6 + i, 43).Value = Workbooks(expenseReportName).Sheets("Авансовый отчет (оборот)").Cells(6 + i, 29).Value
        
          ' Документы, подтверждающие расходы: Наименование документа R7
          Workbooks(expenseReportName).Sheets("Авансовый отчет (оборот)").Cells(6 + i, 18).Value = "Кассовый чек"
        
        End If ' Если сумма чека <>0
        
      Next i
            
      ' --- Акт ---
      ' Тема D16
      Workbooks(expenseReportName).Sheets("Акт").Range("D16").Value = ThisWorkbook.Sheets("Подотчет").Cells(Строка_Подотчета_row, НомерСтолбца_Представительские_расходы + 1).Value
      
      ' Участники H16
      Строка_участников_мероприятия = ""
      For i = 1 To 6
        
        If Строка_участников_мероприятия = "" Then
          Строка_участников_мероприятия = ThisWorkbook.Sheets("Подотчет").Cells(Строка_Подотчета_row, НомерСтолбца_Представительские_расходы + 9 + i).Value
        Else
          Строка_участников_мероприятия = Строка_участников_мероприятия + ", " + ThisWorkbook.Sheets("Подотчет").Cells(Строка_Подотчета_row, НомерСтолбца_Представительские_расходы + 9 + i).Value
        End If
        
      Next i
      
      ' Заносим участников
      Workbooks(expenseReportName).Sheets("Акт").Range("H16").Value = Строка_участников_мероприятия
      Workbooks(expenseReportName).Sheets("Акт").Range("16:16").RowHeight = lineHeight(Строка_участников_мероприятия, 15, 40)
      
      ' Вывод:
      Workbooks(expenseReportName).Sheets("Акт").Range("B19").Value = "В результате проверки предоставленных Прощаевым С.Ф. документов установлено, что представительские расходы на мероприятие " + CStr(expenseReportDate) + " г. составили сумму " + CStr(Workbooks(expenseReportName).Sheets("Авансовый отчет").Range("AH36").Value) + " руб."
      ' Заключение:
      Workbooks(expenseReportName).Sheets("Акт").Range("B25").Value = "1. Признать обоснованными представительские расходы, произведенные " + CStr(expenseReportDate) + " г. в сумме " + CStr(Workbooks(expenseReportName).Sheets("Авансовый отчет").Range("AH36").Value) + " руб."
      
      ' Участники
      Workbooks(expenseReportName).Sheets("Акт").Range("B35").Value = ThisWorkbook.Sheets("Подотчет").Cells(Строка_Подотчета_row, НомерСтолбца_Представительские_расходы + 13).Value
      Workbooks(expenseReportName).Sheets("Акт").Range("B39").Value = ThisWorkbook.Sheets("Подотчет").Cells(Строка_Подотчета_row, НомерСтолбца_Представительские_расходы + 14).Value
      Workbooks(expenseReportName).Sheets("Акт").Range("B41").Value = ThisWorkbook.Sheets("Подотчет").Cells(Строка_Подотчета_row, НомерСтолбца_Представительские_расходы + 15).Value
      ' "___________"
      Workbooks(expenseReportName).Sheets("Акт").Range("H35").Value = "___________"
      Workbooks(expenseReportName).Sheets("Акт").Range("H39").Value = "___________"
      Workbooks(expenseReportName).Sheets("Акт").Range("H41").Value = "___________"
             
      ' Отчет сформирован
      
      ' --- Смета ---
      ' Место:
      Workbooks(expenseReportName).Sheets("Смета").Range("D16").Value = "Тюмень, ул.Советская 51/1"
      ' Предполагаемое количество участников:
      Workbooks(expenseReportName).Sheets("Смета").Range("F18").Value = "6"
      ' Сумма - Приобретение продуктов питания
      Workbooks(expenseReportName).Sheets("Смета").Range("E25").Value = Workbooks(expenseReportName).Sheets("Авансовый отчет").Range("AH36").Value
      Workbooks(expenseReportName).Sheets("Смета").Range("H25").Value = Workbooks(expenseReportName).Sheets("Авансовый отчет").Range("AH36").Value
      ' Итого протокольное обслуживание:
      Workbooks(expenseReportName).Sheets("Смета").Range("E29").Value = Workbooks(expenseReportName).Sheets("Авансовый отчет").Range("AH36").Value
      Workbooks(expenseReportName).Sheets("Смета").Range("H29").Value = Workbooks(expenseReportName).Sheets("Авансовый отчет").Range("AH36").Value
      ' Итого
      Workbooks(expenseReportName).Sheets("Смета").Range("E38").Value = Workbooks(expenseReportName).Sheets("Авансовый отчет").Range("AH36").Value
      Workbooks(expenseReportName).Sheets("Смета").Range("H38").Value = Workbooks(expenseReportName).Sheets("Авансовый отчет").Range("AH36").Value
      
      ' Закрытие файла с Авансовым отчетом
      Workbooks(expenseReportName).Close SaveChanges:=True
     
      ' Формирование
      ' Авансовый Отчет

      ' ФИО и подразделение «Прощаев Сергей Федорович» 

      ' Сумма «9899,82»

      ' Число чеков (обязательно в рамках 223-ФЗ)  «1»

      ' ПФМ/ССП «3004600100»

      ' ФинПозиция/Статья «1637»

      ' Проект «000026»

      ' Код функциональной сферы* «»

      ' Наименование/Код проблемного клиента** «»

      ' Примечание:

      ' !!!Обязательно вложить:

      ' • Заполненный Авансовый отчет в формате файла Excel 

      ' • Копии первичных документов

      ' • Акт (Отчет) о представительских расходах - для Авансовых отчетов по представительским расходам

      ' *Заполняется обязательно для автотранспортных расходов

      ' ВАЖНО! Оригинал А/О распечатывается после дозаполнения его со стороны сотрудников бухгалтерии и вместе со всеми потдверждающими документами должен быть передан в подразделение ответственное за формирование документов дня
    
      ' Сообщение
      MsgBox ("Авансовый отчет " + ThisWorkbook.Path + "\Out\" + expenseReportName + " сформирован!")

    End If
  Else
    MsgBox ("Укажите ячейку в диапазоне Таблицы!")
  End If


End Sub

' Открыть Файл с отчетом: "O:\DirectSales\01_ЗП Проекты\Представительские расходы\Учет представительских расходов.xlsx"
Sub Подотчет_Открыть_Файл_с_отчетом()
    
    ' Открыть файл с отчетом
    FileName = ThisWorkbook.Sheets("Подотчет").Range("H2").Value

    ' Открываем выбранную книгу (UpdateLinks:=0)
    Workbooks.Open FileName, 0

End Sub


' Отправка письма: отправляю шаблон самому себе для последующей отправки в сеть письма на его основе:
Sub Отправка_Lotus_Notes_Подотчет()
Dim темаПисьма, текстПисьма, hashTag, attachmentFile As String
Dim i As Byte
  
  ' Запрос
  If MsgBox("Отправить себе Шаблон письма?", vbYesNo) = vbYes Then
    
    ' Тема письма - Тема:
    темаПисьма = "Проведение клиентского мероприятия"
    ' темаПисьма = subjectFromSheet("Лист8")

    ' hashTag - Хэштэг:
    ' hashTag = ThisWorkbook.Sheets("Лист8").Cells(RowByValue(ThisWorkbook.Name, "Лист8", "Хэштэг:", 100, 100), ColumnByValue(ThisWorkbook.Name, "Лист8", "Хэштэг:", 100, 100) + 1).Value
    ' hashTag - Хэштэг:
    hashTag = "#подотчет #представительские"

    ' Файл-вложение (!!!)
    attachmentFile = "" ' ThisWorkbook.Sheets("Лист8").Cells(3, 17).Value
    
    ' Текст письма
    текстПисьма = "" + Chr(13)
    текстПисьма = текстПисьма + "" + getFromAddrBook("НОКП, РРКК", 2) + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "" + getFromAddrBook("РД", 2) + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "Уважаемые сотрудники," + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "По итогам проведения клиентского мероприятия прошу каждого внести в форму внутренного учета проведенные активности с указанием: " + Chr(13)
    текстПисьма = текстПисьма + " - ФИО" + Chr(13)
    текстПисьма = текстПисьма + " - Должности" + Chr(13)
    текстПисьма = текстПисьма + " - Наименования компании и ИНН" + Chr(13)
    текстПисьма = текстПисьма + " - Что было передано" + Chr(13)
    текстПисьма = текстПисьма + "" + Chr(13)
    текстПисьма = текстПисьма + "Файл с формой внутреннего учета: " + ThisWorkbook.Sheets("Подотчет").Range("H2").Value + Chr(13)
        
    ' Визитка (подпись С Ув., )
    текстПисьма = текстПисьма + ПодписьВПисьме()
    ' Хэштег
    текстПисьма = текстПисьма + createBlankStr(27) + hashTag
    ' Вызов
    Call send_Lotus_Notes(темаПисьма, "Sergey Fedorovich Proschaev/Tyumen/PSBank/Ru", "Sergey Fedorovich Proschaev/Tyumen/PSBank/Ru", текстПисьма, attachmentFile)
  
    ' Зачеркнуть
    Call ЗачеркиваемТекстВячейке("Лист0", RangeByValue(ThisWorkbook.Name, "Лист0", "DashBoard (при наличии)", 100, 100))
  
    ' Сообщение
    MsgBox ("Письмо отправлено!")
     
  End If
  
End Sub

