Attribute VB_Name = "Module_Лист1"
' Вставляем запись в Таблицу BASE\Mortgage
Function Insert_To_Table_Mortgage(In_Номер_кредитного_договора, In_Date_iss, In_Ставка_выдачи, In_Срок_кредита, In_CREDIT_PROGRAMM_OTHER, In_Первоначальный_взнос, In_Стоимость, In_Адрес_предмета_залога, In_Сумма_выдачи_для_срвз, In_CLIENT_ID, In_FIO, In_Доход_по_осн_месту_работы, In_Форма_справки_о_доходах, In_Наименование_компании, In_ИНН, In_Название_партнера, In_Тип_партнера, In_Группа_компаний, In_Выдано, In_Филиал, In_Офис, In_Обновить_по_найденным)
  
  ' Переходим в Книгу Mortgage
  Workbooks("Mortgage").Activate
  Sheets("Лист1").Select
    
  ' Выполняем поиск
  Set Поиск_номера_договора = Columns("A:A").Find(In_Номер_кредитного_договора, LookAt:=xlWhole)
  ' Проверяем - есть ли такой ипотечный кредит в Mortgage, если нет, то добавляем
  If Поиск_номера_договора Is Nothing Then
    ' Если не найден
    Workbooks("Mortgage").Sheets("Лист1").Rows("2:2").Select
    Selection.Insert Shift:=xlDown
    RowNumber = 2
  Else
    ' Если найден, то записываем данные
    RowNumber = Поиск_номера_договора.Row
  End If
  
  ' В добавленную строку записываем данные если договор не был найден или обновляем поля если стоит опция одновить по найденным
  If (Поиск_номера_договора Is Nothing) Or (In_Обновить_по_найденным = True) Then
    ' № кредитного договора
    Range("A" + CStr(RowNumber)).Value = In_Номер_кредитного_договора
    ' Date_iss
    Range("B" + CStr(RowNumber)).Value = In_Date_iss
    ' Ставка выдачи
    Range("C" + CStr(RowNumber)).Value = In_Ставка_выдачи
    ' Срок кредита (мес)
    Range("D" + CStr(RowNumber)).Value = In_Срок_кредита
    ' CREDIT_PROGRAMM_OTHER
    Range("E" + CStr(RowNumber)).Value = In_CREDIT_PROGRAMM_OTHER
    ' Первоначальный взнос
    ' Range("F" + CStr(RowNumber)).Value = Round(In_Первоначальный_взнос / 1000, 3)
    Range("F" + CStr(RowNumber)).Value = In_Первоначальный_взнос
    ' Стоимость (для ПВ)
    ' Range("G" + CStr(RowNumber)).Value = Round(In_Стоимость / 1000, 3)
    Range("G" + CStr(RowNumber)).Value = In_Стоимость
    ' Адрес предмета залога
    Range("H" + CStr(RowNumber)).Value = In_Адрес_предмета_залога
    ' Сумма выдачи_для срвз
    ' Range("I" + CStr(RowNumber)).Value = Round(In_Сумма_выдачи_для_срвз, 3)
    Range("I" + CStr(RowNumber)).Value = In_Сумма_выдачи_для_срвз
    ' CLIENT_ID
    Range("J" + CStr(RowNumber)).Value = In_CLIENT_ID
    ' FIO
    Range("K" + CStr(RowNumber)).Value = In_FIO
    ' Доход по осн. месту работы
    Range("L" + CStr(RowNumber)).Value = In_Доход_по_осн_месту_работы
    ' Форма справки о доходах
    Range("M" + CStr(RowNumber)).Value = In_Форма_справки_о_доходах
    ' ИНН
    Range("N" + CStr(RowNumber)).Value = In_ИНН
    ' Наименование компании
    Range("O" + CStr(RowNumber)).Value = In_Наименование_компании
    ' Название партнера
    Range("P" + CStr(RowNumber)).Value = In_Название_партнера
    ' Тип партнера
    Range("Q" + CStr(RowNumber)).Value = In_Тип_партнера
    ' Группа компаний
    Range("R" + CStr(RowNumber)).Value = In_Группа_компаний
    ' Выдано
    Range("S" + CStr(RowNumber)).Value = In_Выдано
    ' Филиал
    Range("T" + CStr(RowNumber)).Value = In_Филиал
    ' Офис
    Range("U" + CStr(RowNumber)).Value = In_Офис

  End If
  
  ' Возврат в Книгу DB_Result
  ThisWorkbook.Activate
  ThisWorkbook.Sheets("Лист2").Select
  
End Function

' В обработке ML вывод итога по организации (числу ипотек) - выводим все в одну таблицу
Function Вывод_в_отчет_итогов_по_Организации_2(In_MLName, In_Текущий_ИНН, In_count_Текущий_ИНН, In_счетчик_строк, In_Текущая_организация_наименование, In_Обновлять_данные)
  
  ' Обновляем существующие записи/добавляем новые значения
  If In_Обновлять_данные = False Then
    
    ' Если добавляем данные
    ' Проверяем - если последняя запись была выведена в отчет?
    If In_Текущий_ИНН <> "" Then
        ThisWorkbook.Sheets("Лист2").Range("B" + CStr(In_счетчик_строк)).Value = In_Текущий_ИНН
      Else
        ThisWorkbook.Sheets("Лист2").Range("B" + CStr(In_счетчик_строк)).Value = "Нет ИНН"
    End If
  
    ' Проверяем наименование организации
    If In_Текущая_организация_наименование <> "" Then
        ThisWorkbook.Sheets("Лист2").Range("C" + CStr(In_счетчик_строк)).Value = In_Текущая_организация_наименование
      Else
        ThisWorkbook.Sheets("Лист2").Range("C" + CStr(In_счетчик_строк)).Value = "Не определена по ИНН"
    End If
  
    ThisWorkbook.Sheets("Лист2").Range("D" + CStr(In_счетчик_строк)).Value = CStr(In_count_Текущий_ИНН)
        
  Else
    ' Если данные обновлять
        
  End If
        
End Function


' В обработке ML вывод итога по организации (числу ипотек)
Function Вывод_в_отчет_итогов_по_Организации(In_MLName, In_Текущий_ИНН, In_count_Текущий_ИНН, In_счетчик_строк, In_Текущая_организация_наименование)

   ' Функция работает, только надо с синхронизацией вставки строк разобраться
   
   ' Переходим в окно активной книги DB_Result
   ThisWorkbook.Activate
   ThisWorkbook.Sheets("Лист2").Select

  ' Группируем по числу ипотек на организацию:
  ' Одиночные - Строка 10
  If In_count_Текущий_ИНН = 1 Then
    ' Добавление строки
    ThisWorkbook.Sheets("Лист2").Rows("10:10").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    In_счетчик_строк = 10
  End If
  
  ' 2 и более - Строка 5
  If In_count_Текущий_ИНН > 1 Then
    ' Добавление строки
    ThisWorkbook.Sheets("Лист2").Rows("5:5").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    In_счетчик_строк = 5
  End If
  
  ' Проверяем - если последняя запись была выведена в отчет?
  If In_Текущий_ИНН <> "" Then
      ThisWorkbook.Sheets("Лист2").Range("B" + CStr(In_счетчик_строк)).Value = In_Текущий_ИНН
    Else
      ThisWorkbook.Sheets("Лист2").Range("B" + CStr(In_счетчик_строк)).Value = "Нет ИНН"
  End If
  
  ThisWorkbook.Sheets("Лист2").Range("C" + CStr(In_счетчик_строк)).Value = In_Текущая_организация_наименование
  ThisWorkbook.Sheets("Лист2").Range("D" + CStr(In_счетчик_строк)).Value = CStr(In_count_Текущий_ИНН)
    
  ' Возвращаемся опять в книгу ML на Лист 1 - Переходим в ML-файл
  Workbooks(In_MLName).Activate
  Sheets("Лист1").Select
      
End Function


' Имя столбца из адреса $BK$25
Function Наименование_столбца(In_Адрес_ячейки As String) As String
  ' Наименование_столбца = In_Адрес_ячейки
  Позиция_первого_знака_$ = InStr(In_Адрес_ячейки, "$")
  Позиция_второго_знака_$ = Позиция_первого_знака_$ + InStr(Mid(In_Адрес_ячейки, Позиция_первого_знака_$ + 1, Len(In_Адрес_ячейки) - Позиция_первого_знака_$), "$")
  Наименование_столбца = Mid(In_Адрес_ячейки, Позиция_первого_знака_$ + 1, Позиция_второго_знака_$ - Позиция_первого_знака_$ - 1)
End Function


' Проверка - есть ли такая сводная таблица в Книге?
Function PivotExist(Name As String) As Boolean
  Dim sh As Worksheet
  Dim pt As PivotTable
  PivotExist = False
  For Each sh In ActiveWorkbook.Worksheets
      For Each pt In sh.PivotTables
          If pt.Name = Name Then
              PivotExist = True
              Exit For
          End If
      Next
  Next
  
   ' Этим скриптом выводим все поля Таблицы
   ' If PivotExist = True Then
   
   ' MyFile1 = "PivotTables.txt"
   ' Open MyFile1 For Output As #1
   ' Print #1, "--- Таблица " + Name + "---"
   ' For Each pvtField In ActiveSheet.PivotTables(Name).PivotFields
   '   Print #1, "Имя поля: " + pvtField.Name
   ' Next pvtField
   ' Close #1
   '
   ' End If
     
End Function

' Проверка наличия листа с заданным именем в Книге (версия из Интернет, используется на Листе1)
Function Sheets_Exist(wb As Workbook, sName As String) As Boolean
    Dim wsSh As Worksheet
    On Error Resume Next
    Set wsSh = wb.Sheets(sName)
    Sheets_Exist = Not wsSh Is Nothing
End Function

' Имя файла без расширения и пути
Function FileName_WithOutExt(In_fileName As String) As String
  In_fileName = Dir(In_fileName)
  FileName_WithOutExt = Mid(In_fileName, InStr(In_fileName, ".") + 1, Len(In_fileName) - InStr(In_fileName, "."))
End Function


' Предидущая буква английского алфавита
Function Предидущая_буква(In_Буква As String) As String
  Предидущая_буква = ""
  ' In_Буква
  If In_Буква = "A" Then
    Предидущая_буква = "Z"
  End If
  If In_Буква = "B" Then
    Предидущая_буква = "A"
  End If
  If In_Буква = "C" Then
    Предидущая_буква = "B"
  End If
  If In_Буква = "D" Then
    Предидущая_буква = "C"
  End If
  If In_Буква = "E" Then
    Предидущая_буква = "D"
  End If
  If In_Буква = "F" Then
    Предидущая_буква = "E"
  End If
  If In_Буква = "G" Then
    Предидущая_буква = "F"
  End If
  If In_Буква = "H" Then
    Предидущая_буква = "G"
  End If
  If In_Буква = "I" Then
    Предидущая_буква = "H"
  End If
  If In_Буква = "J" Then
    Предидущая_буква = "I"
  End If
  If In_Буква = "K" Then
    Предидущая_буква = "J"
  End If
  If In_Буква = "L" Then
    Предидущая_буква = "K"
  End If
  If In_Буква = "M" Then
    Предидущая_буква = "L"
  End If
  If In_Буква = "N" Then
    Предидущая_буква = "M"
  End If
  If In_Буква = "O" Then
    Предидущая_буква = "N"
  End If
  If In_Буква = "P" Then
    Предидущая_буква = "O"
  End If
  If In_Буква = "Q" Then
    Предидущая_буква = "P"
  End If
  If In_Буква = "R" Then
    Предидущая_буква = "Q"
  End If
  If In_Буква = "S" Then
    Предидущая_буква = "R"
  End If
    
End Function

' Получение имени файла без пути c:\1.txt => 1.txt
Function getFName(pf As String) As String
  If InStrRev(pf, "\") <> 0 Then
    getFName = Mid(pf, InStrRev(pf, "\") + 1)
  Else
    getFName = pf
  End If
End Function

' Заливка ячейки цветом "светофор"
Sub Full_Color_Range(In_list, In_Range, In_Value)
  In_Value = In_Value * 100
  ' Если до этого ячейка была цветная - сбрасываем цвет
  ThisWorkbook.Sheets(In_list).Range(In_Range).Interior.Color = xlNone
  ' Цвет текста - черный
  ThisWorkbook.Sheets(In_list).Range(In_Range).Font.Color = vbBlack
  ' От 100% и выше - Зеленый
  If (In_Value >= 100) Then
    ThisWorkbook.Sheets(In_list).Range(In_Range).Interior.Color = vbGreen
  End If
  ' От 90%-100% - Желтый
  If (In_Value >= 90) And (In_Value < 100) Then
    ThisWorkbook.Sheets(In_list).Range(In_Range).Interior.Color = vbYellow
  End If
  ' От 0% - 90% - Красный
  If (In_Value < 90) Then
    ThisWorkbook.Sheets(In_list).Range(In_Range).Interior.Color = vbRed
  End If
End Sub

' Заливка ячейки цветом "светофор" для Интегрального рейтинга - до 50% красный до 100% желтый, св 100% зеленый
Sub Full_Color_Range_For_Int_Rating(In_list, In_Range, In_Value)
  In_Value = In_Value * 100
  ' Если до этого ячейка была цветная - сбрасываем цвет
  ThisWorkbook.Sheets(In_list).Range(In_Range).Interior.Color = xlNone
  ' Цвет текста - черный
  ThisWorkbook.Sheets(In_list).Range(In_Range).Font.Color = vbBlack
  ' От 100% и выше - Зеленый
  If (In_Value >= 100) Then
    ThisWorkbook.Sheets(In_list).Range(In_Range).Interior.Color = vbGreen
  End If
  ' От 90%-100% - Желтый
  If (In_Value >= 50) And (In_Value < 100) Then
    ThisWorkbook.Sheets(In_list).Range(In_Range).Interior.Color = vbYellow
  End If
  ' От 0% - 90% - Красный
  If (In_Value < 50) Then
    ThisWorkbook.Sheets(In_list).Range(In_Range).Interior.Color = vbRed
  End If
End Sub


' Заливка текста ячейки цветом "светофор"
Sub Full_Color_Text(In_list, In_Range, In_Value)
  In_Value = In_Value * 100
  ' Если до этого ячейка была цветная - сбрасываем цвет
  ThisWorkbook.Sheets(In_list).Range(In_Range).Interior.Color = xlNone
  ' От 100% и выше - Зеленый
  If (In_Value >= 100) Then
    ThisWorkbook.Sheets(In_list).Range(In_Range).Font.Color = vbGreen
  End If
  ' От 90%-100% - Желтый
  If (In_Value >= 90) And (In_Value < 100) Then
    ThisWorkbook.Sheets(In_list).Range(In_Range).Font.Color = vbYellow
  End If
  ' От 0% - 90% - Красный
  If (In_Value < 90) Then
    ThisWorkbook.Sheets(In_list).Range(In_Range).Font.Color = vbRed
  End If
End Sub

' Установка кружка "светофор"
Sub Set_Color_circle(In_list, In_Range)
    ' Ячейка DB из которой копируем формат
    Range("C17").Select
    ' Само копирование
    Selection.Copy
    ' Выбираем куда копировать
    Range("D10").Select
    ' Сама вставка
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
End Sub



' Добавление в строку Лидеры продаж
Function Добавление_Список_МРК_Лидеры_продаж(In_Вып_Факт_Процент, In_FIO_Name, In_Список_МРК_Лидеры_продаж) As String
  If In_Вып_Факт_Процент >= 1 Then
        If In_Список_МРК_Лидеры_продаж <> "" Then
          In_Список_МРК_Лидеры_продаж = In_Список_МРК_Лидеры_продаж + ", " + In_FIO_Name
        Else
          In_Список_МРК_Лидеры_продаж = In_FIO_Name
        End If
  End If
  ' Результат
  Добавление_Список_МРК_Лидеры_продаж = In_Список_МРК_Лидеры_продаж
End Function

' Добавление в строку списка МРК с отсутствием продаж
Function Добавление_Список_МРК_с_Отсутствием_продаж(In_Вып_Факт_Процент, In_FIO_Name, In_Список_МРК_с_Отсутствием_продаж) As String
      
      If In_Вып_Факт_Процент = 0 Then
         If In_Список_МРК_с_Отсутствием_продаж <> "" Then
           In_Список_МРК_с_Отсутствием_продаж = In_Список_МРК_с_Отсутствием_продаж + ", " + In_FIO_Name
        Else
          In_Список_МРК_с_Отсутствием_продаж = In_FIO_Name
        End If
      End If
      
      ' Результат
      Добавление_Список_МРК_с_Отсутствием_продаж = In_Список_МРК_с_Отсутствием_продаж
      
End Function


' Импортирование данных с Листа Интегральный_рейтинг_по_сотрудникам
Sub Интегральный_рейтинг_по_сотрудникам(In_DBstrName As String)

' Переменные
Dim KuratorVar, FIO_Name, Office2_Name As String
' Факт выполнения показателей
Dim Вып_ПК_Факт_Процент, Вып_Страховки_к_ПК_Факт_Процент, Вып_КК_Факт_Процент, Вып_ДК_Факт_Процент, Вып_ИБ_Факт_Процент, Вып_НС_Факт_Процент, Вып_ИСЖ_Факт_Процент, Вып_НСЖ_Факт_Процент, Вып_КС_Факт_Процент, Вып_ЛА_Факт_Процент, Вып_СМС_Факт_Процент, Интегральный_рейтинг_Процент As Double
' Лидеры продаж
Dim Список_МРК_Лидеры_продаж_ПК, Список_МРК_Лидеры_продаж_СтраховокПК, Список_МРК_Лидеры_продаж_КК, Список_МРК_Лидеры_продаж_ДК, Список_МРК_Лидеры_продаж_ИСЖ, Список_МРК_Лидеры_продаж_НСЖ  As String
' Отсутствие продаж
Dim Список_МРК_с_Отсутствием_продаж_ПК, Список_МРК_с_Отсутствием_продаж_СтраховокПК, Список_МРК_с_Отсутствием_продаж_KK, Список_МРК_с_Отсутствием_продаж_ДK, Список_МРК_с_Отсутствием_продаж_ИСЖ, Список_МРК_с_Отсутствием_продаж_НСЖ As String
' Строка вывода
Dim RowToPrint As Byte
' Логирование в текстовый файл чтения DB
Dim Логирование_в_текстовые_файлы As Boolean
' Колонка_с_Интегральным_рейтингом
Dim Столбец_ФИО_откуда_берем_данные, Столбец_РОО_DP3_отчет_откуда_берем_данные, Столбец_ОО2_DP4_отчет_откуда_берем_данные, Столбец_с_Интегральным_рейтингом_откуда_берем_данные, Столбец_ПК_Факт_откуда_берем_данные, Столбец_Вып_Страховки_к_ПК_Факт_откуда_берем_данные, Столбец_Вып_КК_Факт_откуда_берем_данные, Столбец_Вып_ДК_Факт_откуда_берем_данные, Столбец_Вып_ИБ_Факт_откуда_берем_данные, Столбец_Вып_НС_Факт_откуда_берем_данные, Столбец_Вып_ИСЖ_Факт_откуда_берем_данные, Столбец_Вып_НСЖ_Факт_откуда_берем_данные, Столбец_Вып_КС_Факт_откуда_берем_данные, Столбец_Вып_ЛА_Факт_откуда_берем_данные, Столбец_Вып_СМС_Факт_откуда_берем_данные As String

  Application.StatusBar = "Интегральный_рейтинг_по_сотрудникам ..."
  
  ' Логирование в текстовый файл чтения DB: True - логируем, False - не логируем
  Логирование_в_текстовые_файлы = False

  ' Перейти на "4. Интегральный рей-г по сотруд"
  Sheets("4. Интегральный рей-г по сотруд").Select

  ' Логирование в текстовые файлы
  If Логирование_в_текстовые_файлы = True Then

    ' В файл выводим все:
    MyFile1 = In_DBstrName & "_Инт_рейт_сотр_1_log.txt"
    Open MyFile1 For Output As #1

    ' Второй вариант имени файла - вывод в конкретный каталог
    MyFile2 = In_DBstrName & "_Инт_рейт_сотр_2_log.txt"
    ' Открыли для записи
    Open MyFile2 For Output As #2

  End If
  
  ' Инициализация переменных
  ' Это строка 072?
  ThiStringIs072 = False
  ' Счетчик МРК
  CountMRK = 0
  ' Наименование офиса
  Office2_Name = ""
  ' Фио сотрудника
  FIO_Name = ""
  ' Вывод в мой файл строки
  Вывод_Номера_и_ФИО_сотрудника = False
  
  ' Обнуление списков
  Список_МРК_Лидеры_продаж_ПК = ""
  Список_МРК_с_Отсутствием_продаж_ПК = ""
  Список_МРК_Лидеры_продаж_СтраховокПК = ""
  Список_МРК_с_Отсутствием_продаж_СтраховокПК = ""
  Список_МРК_Лидеры_продаж_КК = ""
  Список_МРК_с_Отсутствием_продаж_КК = ""
  Список_МРК_Лидеры_продаж_ДК = ""
  Список_МРК_с_Отсутствием_продаж_ДК = ""
  Список_МРК_Лидеры_продаж_ИСЖ = ""
  Список_МРК_с_Отсутствием_продаж_ИСЖ = ""
  Список_МРК_Лидеры_продаж_НСЖ = ""
  Список_МРК_с_Отсутствием_продаж_НСЖ = ""

  ' Колонки откуда берем данные
  Столбец_с_Интегральным_рейтингом_откуда_берем_данные = ""
  Столбец_ФИО_откуда_берем_данные = ""
  Столбец_РОО_DP3_отчет_откуда_берем_данные = ""
  Столбец_ОО2_DP4_отчет_откуда_берем_данные = ""
  Столбец_ПК_Факт_откуда_берем_данные = ""
  Столбец_Вып_Страховки_к_ПК_Факт_откуда_берем_данные = ""
  Столбец_Вып_КК_Факт_откуда_берем_данные = ""
  Столбец_Вып_ДК_Факт_откуда_берем_данные = ""
  Столбец_Вып_ИБ_Факт_откуда_берем_данные = ""
  Столбец_Вып_НС_Факт_откуда_берем_данные = ""
  Столбец_Вып_ИСЖ_Факт_откуда_берем_данные = ""
  Столбец_Вып_НСЖ_Факт_откуда_берем_данные = ""
  Столбец_Вып_КС_Факт_откуда_берем_данные = ""
  Столбец_Вып_ЛА_Факт_откуда_берем_данные = ""
  Столбец_Вып_СМС_Факт_откуда_берем_данные = ""
  
  ' Цикл обработки Листа Excel
  For Each Cell In ActiveSheet.UsedRange
  
      ' Если ячейка не пустая
      If Not IsEmpty(Cell) Then
    
        ' Выводим все данные в строку
        If Логирование_в_текстовые_файлы = True Then
          Print #1, Cell.Address, ":" + Cell.Formula
        End If
    
        ' Выводим в текущую книгу информацию из ячейки
        ' ThisWorkbook.Sheets("Лист1").Range("A2").Value = Workbooks("Dashboard_new_РБ_25.09.2019.xlsm").Sheets("1. Интегральный рейтинг").Range("A1")
     
        ' Вывод строки и столбца
        ' Номер столбца
        intC = Cell.Column
        ' Номер строки
        intR = Cell.Row
      
        ' Если новый столбец А, то считаем, что строка не "Тюменский ОО1"
        If Mid(Cell.Address, 1, 3) = "$A$" Then
            ' Офис не Тюменский OO1
            ThiStringIs072 = False
            ' Вывод_Номера_и_ФИО_сотрудника
            Вывод_Номера_и_ФИО_сотрудника = False
        End If
            
        ' Если адрес ячейки $B$XXX - значит это ФИО, записываем в переменную (для любого офиса). Наименование столбца "ФИО"
        ' If Mid(cell.Address, 1, 3) = "$B$" Then
        ' If Mid(cell.Address, 1, 3) = Столбец_ФИО_откуда_берем_данные Then
        If Наименование_столбца(Cell.Address) = Столбец_ФИО_откуда_берем_данные Then
          FIO_Name = CStr(Cells(intR, intC).Value)
        End If
                        
        ' Если это столбец D ("Тюменский ОО1"). Наименование столбца "DP3_отчет". Столбец_РОО_DP3_отчет_откуда_берем_данные
        ' If (Mid(cell.Address, 1, 3) = "$D$") Then
        ' If (Mid(cell.Address, 1, 3) = Столбец_РОО_DP3_отчет_откуда_берем_данные) Then
        If Наименование_столбца(Cell.Address) = Столбец_РОО_DP3_отчет_откуда_берем_данные Then
          If (CStr(Cells(intR, intC).Value) = "Тюменский ОО1") Then
            ' Если офис Тюменский OO1
            ThiStringIs072 = True
            ' Счетчик
            CountMRK = CountMRK + 1
            ' Вывод_Номера_и_ФИО_сотрудника
            Вывод_Номера_и_ФИО_сотрудника = True
          Else
            ' Если офис не Тюменский OO1
            ThiStringIs072 = False
          End If
        End If
    
        ' Если адрес ячейки $E$X - значит это Офис второго уровня. Наименование столбца "DP4_отчет". Столбец_ОО2_DP4_отчет_откуда_берем_данные
        ' If (Mid(cell.Address, 1, 3) = "$E$") And (ThiStringIs072 = True) Then
        ' If (Mid(cell.Address, 1, 3) = Столбец_ОО2_DP4_отчет_откуда_берем_данные) And (ThiStringIs072 = True) Then
        If (Наименование_столбца(Cell.Address) = Столбец_ОО2_DP4_отчет_откуда_берем_данные) And (ThiStringIs072 = True) Then
          Office2_Name = CStr(Cells(intR, intC).Value)
          ThisWorkbook.Sheets("Лист1").Range("O" + CStr(9 + CountMRK)).Value = cityOfficeName(Office2_Name)
        End If
        
        ' === Блок записи показателей из DB в мою книгу ===
    
        ' Выводим в Мою книгу № и ФИО сотрудника
        If (Вывод_Номера_и_ФИО_сотрудника = True) And (ThiStringIs072 = True) Then
          ThisWorkbook.Sheets("Лист1").Range("A" + CStr(9 + CountMRK)).Value = CountMRK
          ThisWorkbook.Sheets("Лист1").Range("B" + CStr(9 + CountMRK)).Value = Фамилия_и_Имя(FIO_Name, 3)
          Вывод_Номера_и_ФИО_сотрудника = False
        End If
        
    
        ' 1 показатель: Если адрес ячейки $J$8 - значит это ПК_Факт. Наименование столбца: "Факт_ПК, тыс. руб.". Столбец_ПК_Факт_откуда_берем_данные
        ' If (Mid(cell.Address, 1, 3) = "$J$") And (ThiStringIs072 = True) Then
        ' If (Mid(cell.Address, 1, 3) = Столбец_ПК_Факт_откуда_берем_данные) And (ThiStringIs072 = True) Then
        If (Наименование_столбца(Cell.Address) = Столбец_ПК_Факт_откуда_берем_данные) And (ThiStringIs072 = True) Then
          ' Устанавливаем формат ячейки - процентный "0.00%"
          ThisWorkbook.Sheets("Лист1").Range("C" + CStr(9 + CountMRK)).NumberFormat = "0%"
          ' Копируем формат ячейки
          Cells(intR, intC).Copy Destination:=ThisWorkbook.Sheets("Лист1").Range("C" + CStr(9 + CountMRK))
          ' Берем Факт по ПК
          Вып_ПК_Факт_Процент = Round((Cells(intR, intC).Value), 2)
          ' Заносим данные в мою книгу
          ThisWorkbook.Sheets("Лист1").Range("C" + CStr(9 + CountMRK)).Value = Вып_ПК_Факт_Процент
          ' Список_МРК_Лидеры_продаж_ПК = ""
          Список_МРК_Лидеры_продаж_ПК = Добавление_Список_МРК_Лидеры_продаж(Вып_ПК_Факт_Процент, Фамилия_и_Имя(FIO_Name, 1), Список_МРК_Лидеры_продаж_ПК)
          ' Список_МРК_с_Отсутствием_продаж
          Список_МРК_с_Отсутствием_продаж_ПК = Добавление_Список_МРК_с_Отсутствием_продаж(Вып_ПК_Факт_Процент, Фамилия_и_Имя(FIO_Name, 2), Список_МРК_с_Отсутствием_продаж_ПК)
        End If

        ' 2 показатель: Если адрес ячейки $O$8 - значит Вып_Страховки_к_ПК_Факт_Процент. Наименование столбца: "% Вып_Страховки к ПК_Факт". Столбец_Вып_Страховки_к_ПК_Факт_откуда_берем_данные
        ' If (Mid(cell.Address, 1, 3) = "$O$") And (ThiStringIs072 = True) Then
        ' If (Mid(cell.Address, 1, 3) = Столбец_Вып_Страховки_к_ПК_Факт_откуда_берем_данные) And (ThiStringIs072 = True) Then
        If (Наименование_столбца(Cell.Address) = Столбец_Вып_Страховки_к_ПК_Факт_откуда_берем_данные) And (ThiStringIs072 = True) Then
          ' Устанавливаем формат ячейки - процентный "0.00%"
          ThisWorkbook.Sheets("Лист1").Range("D" + CStr(9 + CountMRK)).NumberFormat = "0%"
          ' Копируем формат ячейки
          Cells(intR, intC).Copy Destination:=ThisWorkbook.Sheets("Лист1").Range("D" + CStr(9 + CountMRK))
          ' Берем Вып_Страховки_к_ПК_Факт_Процент
          Вып_Страховки_к_ПК_Факт_Процент = Round((Cells(intR, intC).Value), 2)
          ' Заносим данные в мою книгу
          ThisWorkbook.Sheets("Лист1").Range("D" + CStr(9 + CountMRK)).Value = Вып_Страховки_к_ПК_Факт_Процент
          ' Красим цвет текста в ячейке в Светофора
          ' Call Full_Color_Text("Лист1", "D" + CStr(9 + CountMRK), Вып_Страховки_к_ПК_Факт_Процент)
          ' Список_МРК_Лидеры_продаж_СтраховокПК = ""
          Список_МРК_Лидеры_продаж_СтраховокПК = Добавление_Список_МРК_Лидеры_продаж(Вып_Страховки_к_ПК_Факт_Процент, Фамилия_и_Имя(FIO_Name, 1), Список_МРК_Лидеры_продаж_СтраховокПК)
          ' Список_МРК_с_Отсутствием_продаж
          Список_МРК_с_Отсутствием_продаж_СтраховокПК = Добавление_Список_МРК_с_Отсутствием_продаж(Вып_Страховки_к_ПК_Факт_Процент, Фамилия_и_Имя(FIO_Name, 2), Список_МРК_с_Отсутствием_продаж_СтраховокПК)
        End If

        ' 3 показатель: Если адрес ячейки $T$8 - значит Вып_КК_Факт_Процент. Наименование столбца: "% Вып_КК _Факт". Столбец_Вып_КК_Факт_откуда_берем_данные
        ' If (Mid(cell.Address, 1, 3) = "$T$") And (ThiStringIs072 = True) Then
        ' If (Mid(cell.Address, 1, 3) = Столбец_Вып_КК_Факт_откуда_берем_данные) And (ThiStringIs072 = True) Then
        If (Наименование_столбца(Cell.Address) = Столбец_Вып_КК_Факт_откуда_берем_данные) And (ThiStringIs072 = True) Then
          ' Устанавливаем формат ячейки - процентный "0.00%"
          ThisWorkbook.Sheets("Лист1").Range("E" + CStr(9 + CountMRK)).NumberFormat = "0%"
          ' Копируем формат ячейки
          Cells(intR, intC).Copy Destination:=ThisWorkbook.Sheets("Лист1").Range("E" + CStr(9 + CountMRK))
          ' Берем Вып_КК_Факт_Процент
          Вып_КК_Факт_Процент = Round((Cells(intR, intC).Value), 2)
          ' Заносим данные в мою книгу
          ThisWorkbook.Sheets("Лист1").Range("E" + CStr(9 + CountMRK)).Value = Вып_КК_Факт_Процент
          ' Красим цвет текста в ячейке в Светофора
          ' Call Full_Color_Text("Лист1", "E" + CStr(9 + CountMRK), Вып_КК_Факт_Процент)
          ' Список_МРК_Лидеры_продаж_ПК = ""
          Список_МРК_Лидеры_продаж_КК = Добавление_Список_МРК_Лидеры_продаж(Вып_КК_Факт_Процент, Фамилия_и_Имя(FIO_Name, 1), Список_МРК_Лидеры_продаж_КК)
          ' Список_МРК_с_Отсутствием_продаж
          Список_МРК_с_Отсутствием_продаж_KK = Добавление_Список_МРК_с_Отсутствием_продаж(Вып_КК_Факт_Процент, Фамилия_и_Имя(FIO_Name, 2), Список_МРК_с_Отсутствием_продаж_KK)
        End If
 
        ' 4 показатель: Если адрес ячейки $Y$8 - значит Вып_ДК_Факт_Процент. Наименование столбца: "% Вып_ДК _Факт". Столбец_Вып_ДК_Факт_откуда_берем_данные
        ' If (Mid(cell.Address, 1, 3) = "$Y$") And (ThiStringIs072 = True) Then
        ' If (Mid(cell.Address, 1, 3) = Столбец_Вып_ДК_Факт_откуда_берем_данные) And (ThiStringIs072 = True) Then
        If (Наименование_столбца(Cell.Address) = Столбец_Вып_ДК_Факт_откуда_берем_данные) And (ThiStringIs072 = True) Then
          ' Устанавливаем формат ячейки - процентный "0.00%"
          ThisWorkbook.Sheets("Лист1").Range("F" + CStr(9 + CountMRK)).NumberFormat = "0%"
          ' Копируем формат ячейки
          Cells(intR, intC).Copy Destination:=ThisWorkbook.Sheets("Лист1").Range("F" + CStr(9 + CountMRK))
          ' Берем Вып_ДК _Факт_Процент
          Вып_ДК_Факт_Процент = Round((Cells(intR, intC).Value), 2)
          ' Заносим данные в мою книгу
          ThisWorkbook.Sheets("Лист1").Range("F" + CStr(9 + CountMRK)).Value = Вып_ДК_Факт_Процент
          ' Красим цвет текста в ячейке в Светофора
          ' Call Full_Color_Text("Лист1", "F" + CStr(9 + CountMRK), Вып_ДК_Факт_Процент)
          ' Список_МРК_Лидеры_продаж = ""
          Список_МРК_Лидеры_продаж_ДК = Добавление_Список_МРК_Лидеры_продаж(Вып_ДК_Факт_Процент, Фамилия_и_Имя(FIO_Name, 1), Список_МРК_Лидеры_продаж_ДК)
          ' Список_МРК_с_Отсутствием_продаж
          Список_МРК_с_Отсутствием_продаж_ДK = Добавление_Список_МРК_с_Отсутствием_продаж(Вып_ДК_Факт_Процент, Фамилия_и_Имя(FIO_Name, 2), Список_МРК_с_Отсутствием_продаж_ДK)
         End If

        ' 5 показатель: Если адрес ячейки $AD$8 - значит Вып_ИБ_Факт_Процент. Наименование столбца: "% Вып_ИБ _Факт". Столбец_Вып_ИБ_Факт_откуда_берем_данные
        ' If (Mid(cell.Address, 1, 4) = "$AD$") And (ThiStringIs072 = True) Then
        ' If (Mid(cell.Address, 1, 4) = Столбец_Вып_ИБ_Факт_откуда_берем_данные) And (ThiStringIs072 = True) Then
        If (Наименование_столбца(Cell.Address) = Столбец_Вып_ИБ_Факт_откуда_берем_данные) And (ThiStringIs072 = True) Then
          ' Устанавливаем формат ячейки - процентный "0.00%"
          ThisWorkbook.Sheets("Лист1").Range("G" + CStr(9 + CountMRK)).NumberFormat = "0%"
          ' Копируем формат ячейки
          Cells(intR, intC).Copy Destination:=ThisWorkbook.Sheets("Лист1").Range("G" + CStr(9 + CountMRK))
          ' Берем Вып_ИБ_Факт_Процент
          Вып_ИБ_Факт_Процент = Round((Cells(intR, intC).Value), 2)
          ' Заносим данные в мою книгу
          ThisWorkbook.Sheets("Лист1").Range("G" + CStr(9 + CountMRK)).Value = Вып_ИБ_Факт_Процент
          ' Красим цвет текста в ячейке в Светофора
          ' Call Full_Color_Text("Лист1", "G" + CStr(9 + CountMRK), Вып_ИБ_Факт_Процент)
        End If

        ' 6 показатель: Если адрес ячейки $AI$8 - значит Вып_НС_Факт_Процент. Наименование столбца: "% Вып_НС _Факт". Столбец_Вып_НС_Факт_откуда_берем_данные
        ' If (Mid(cell.Address, 1, 4) = "$AI$") And (ThiStringIs072 = True) Then
        ' If (Mid(cell.Address, 1, 4) = Столбец_Вып_НС_Факт_откуда_берем_данные) And (ThiStringIs072 = True) Then
        If (Наименование_столбца(Cell.Address) = Столбец_Вып_НС_Факт_откуда_берем_данные) And (ThiStringIs072 = True) Then
          ' Устанавливаем формат ячейки - процентный "0.00%"
          ThisWorkbook.Sheets("Лист1").Range("H" + CStr(9 + CountMRK)).NumberFormat = "0%"
          ' Копируем формат ячейки
          Cells(intR, intC).Copy Destination:=ThisWorkbook.Sheets("Лист1").Range("H" + CStr(9 + CountMRK))
          ' Берем Вып_НС_Факт_Процент
          Вып_НС_Факт_Процент = Round((Cells(intR, intC).Value), 2)
          ' Заносим данные в мою книгу
          ThisWorkbook.Sheets("Лист1").Range("H" + CStr(9 + CountMRK)).Value = Вып_НС_Факт_Процент
          ' Красим цвет текста в ячейке в Светофора
          ' Call Full_Color_Text("Лист1", "H" + CStr(9 + CountMRK), Вып_НС_Факт_Процент)
        End If

        ' 7 показатель: Если адрес ячейки $AN$8 - значит Вып_ИСЖ_Факт_Процент. Наименование столбца: "% Вып_ ИСЖ_Факт". Столбец_Вып_ИСЖ_Факт_откуда_берем_данные
        ' If (Mid(cell.Address, 1, 4) = "$AN$") And (ThiStringIs072 = True) Then
        ' If (Mid(cell.Address, 1, 4) = Столбец_Вып_ИСЖ_Факт_откуда_берем_данные) And (ThiStringIs072 = True) Then
        If (Наименование_столбца(Cell.Address) = Столбец_Вып_ИСЖ_Факт_откуда_берем_данные) And (ThiStringIs072 = True) Then
          ' Устанавливаем формат ячейки - процентный "0.00%"
          ThisWorkbook.Sheets("Лист1").Range("I" + CStr(9 + CountMRK)).NumberFormat = "0%"
          ' Копируем формат ячейки
          Cells(intR, intC).Copy Destination:=ThisWorkbook.Sheets("Лист1").Range("I" + CStr(9 + CountMRK))
          ' Берем Вып_ИСЖ_Факт_Процент
          Вып_ИСЖ_Факт_Процент = Round((Cells(intR, intC).Value), 2)
          ' Заносим данные в мою книгу
          ThisWorkbook.Sheets("Лист1").Range("I" + CStr(9 + CountMRK)).Value = Вып_ИСЖ_Факт_Процент
          ' Красим цвет текста в ячейке в Светофора
          ' Call Full_Color_Text("Лист1", "I" + CStr(9 + CountMRK), Вып_ИСЖ_Факт_Процент)
          ' Список_МРК_Лидеры_продаж = ""
          Список_МРК_Лидеры_продаж_ИСЖ = Добавление_Список_МРК_Лидеры_продаж(Вып_ИСЖ_Факт_Процент, Фамилия_и_Имя(FIO_Name, 1), Список_МРК_Лидеры_продаж_ИСЖ)
          ' Список_МРК_с_Отсутствием_продаж
          Список_МРК_с_Отсутствием_продаж_ИСЖ = Добавление_Список_МРК_с_Отсутствием_продаж(Вып_ИСЖ_Факт_Процент, Фамилия_и_Имя(FIO_Name, 2), Список_МРК_с_Отсутствием_продаж_ИСЖ)
        End If

        ' 8 показатель: Если адрес ячейки $AS$8 - значит Вып_НСЖ_Факт_Процент. Наименование столбца: "% Вып_НСЖ_Факт". Столбец_Вып_НСЖ_Факт_откуда_берем_данные
        ' If (Mid(cell.Address, 1, 4) = "$AS$") And (ThiStringIs072 = True) Then
        ' If (Mid(cell.Address, 1, 4) = Столбец_Вып_НСЖ_Факт_откуда_берем_данные) And (ThiStringIs072 = True) Then
        If (Наименование_столбца(Cell.Address) = Столбец_Вып_НСЖ_Факт_откуда_берем_данные) And (ThiStringIs072 = True) Then
          ' Устанавливаем формат ячейки - процентный "0.00%"
          ThisWorkbook.Sheets("Лист1").Range("J" + CStr(9 + CountMRK)).NumberFormat = "0%"
          ' Копируем формат ячейки
          Cells(intR, intC).Copy Destination:=ThisWorkbook.Sheets("Лист1").Range("J" + CStr(9 + CountMRK))
          ' Берем Вып_НСЖ_Факт_Процент
          Вып_НСЖ_Факт_Процент = Round((Cells(intR, intC).Value), 2)
          ' Заносим данные в мою книгу
          ThisWorkbook.Sheets("Лист1").Range("J" + CStr(9 + CountMRK)).Value = Вып_НСЖ_Факт_Процент
          ' Красим цвет текста в ячейке в Светофора
          ' Call Full_Color_Text("Лист1", "J" + CStr(9 + CountMRK), Вып_НСЖ_Факт_Процент)
          ' Список_МРК_Лидеры_продаж = ""
          Список_МРК_Лидеры_продаж_НСЖ = Добавление_Список_МРК_Лидеры_продаж(Вып_НСЖ_Факт_Процент, Фамилия_и_Имя(FIO_Name, 1), Список_МРК_Лидеры_продаж_НСЖ)
          ' Список_МРК_с_Отсутствием_продаж
          Список_МРК_с_Отсутствием_продаж_НСЖ = Добавление_Список_МРК_с_Отсутствием_продаж(Вып_НСЖ_Факт_Процент, Фамилия_и_Имя(FIO_Name, 2), Список_МРК_с_Отсутствием_продаж_НСЖ)
        End If

        ' 9 показатель: Если адрес ячейки $AX$8 - значит Вып_КС_Факт_Процент. Наименование столбца: "% Вып_КС _Факт". Столбец_Вып_КС_Факт_откуда_берем_данные
        ' If (Mid(cell.Address, 1, 4) = "$AX$") And (ThiStringIs072 = True) Then
        ' If (Mid(cell.Address, 1, 4) = Столбец_Вып_КС_Факт_откуда_берем_данные) And (ThiStringIs072 = True) Then
        If (Наименование_столбца(Cell.Address) = Столбец_Вып_КС_Факт_откуда_берем_данные) And (ThiStringIs072 = True) Then
          ' Устанавливаем формат ячейки - процентный "0.00%"
          ThisWorkbook.Sheets("Лист1").Range("K" + CStr(9 + CountMRK)).NumberFormat = "0%"
          ' Копируем формат ячейки
          Cells(intR, intC).Copy Destination:=ThisWorkbook.Sheets("Лист1").Range("K" + CStr(9 + CountMRK))
          ' Берем Вып_КС_Факт_Процент
          Вып_КС_Факт_Процент = Round((Cells(intR, intC).Value), 2)
          ' Заносим данные в мою книгу
          ThisWorkbook.Sheets("Лист1").Range("K" + CStr(9 + CountMRK)).Value = Вып_КС_Факт_Процент
          ' Красим цвет текста в ячейке в Светофора
          ' Call Full_Color_Text("Лист1", "K" + CStr(9 + CountMRK), Вып_КС_Факт_Процент)
        End If

        ' 10 показатель: Если адрес ячейки $BC$8 - значит Вып_ЛА_Факт_Процент. Наименование столбца: "% Вып_ ЛА". Столбец_Вып_ЛА_Факт_откуда_берем_данные
        ' If (Mid(cell.Address, 1, 4) = "$BC$") And (ThiStringIs072 = True) Then
        ' If (Mid(cell.Address, 1, 4) = Столбец_Вып_ЛА_Факт_откуда_берем_данные) And (ThiStringIs072 = True) Then
        If (Наименование_столбца(Cell.Address) = Столбец_Вып_ЛА_Факт_откуда_берем_данные) And (ThiStringIs072 = True) Then
          ' Устанавливаем формат ячейки - процентный "0.00%"
          ThisWorkbook.Sheets("Лист1").Range("L" + CStr(9 + CountMRK)).NumberFormat = "0%"
          ' Копируем формат ячейки
          Cells(intR, intC).Copy Destination:=ThisWorkbook.Sheets("Лист1").Range("L" + CStr(9 + CountMRK))
          ' Берем Вып_ЛА_Факт_Процент
          Вып_ЛА_Факт_Процент = Round((Cells(intR, intC).Value), 2)
          ' Заносим данные в мою книгу
          ThisWorkbook.Sheets("Лист1").Range("L" + CStr(9 + CountMRK)).Value = Вып_ЛА_Факт_Процент
          ' Красим цвет текста в ячейке в Светофора
          ' Call Full_Color_Text("Лист1", "L" + CStr(9 + CountMRK), Вып_ЛА_Факт_Процент)
        End If

        ' 11 показатель: Если адрес ячейки $BH$8 - значит Вып_СМС_Факт_Процент. Наименование столбца: "% Вып_ СМС". Столбец_Вып_СМС_Факт_откуда_берем_данные
        ' If (Mid(cell.Address, 1, 4) = "$BH$") And (ThiStringIs072 = True) Then
        ' If (Mid(cell.Address, 1, 4) = Столбец_Вып_СМС_Факт_откуда_берем_данные) And (ThiStringIs072 = True) Then
        If (Наименование_столбца(Cell.Address) = Столбец_Вып_СМС_Факт_откуда_берем_данные) And (ThiStringIs072 = True) Then
          ' Устанавливаем формат ячейки - процентный "0.00%"
          ThisWorkbook.Sheets("Лист1").Range("M" + CStr(9 + CountMRK)).NumberFormat = "0%"
          ' Копируем формат ячейки
          Cells(intR, intC).Copy Destination:=ThisWorkbook.Sheets("Лист1").Range("M" + CStr(9 + CountMRK))
          ' Берем Вып_СМС_Факт_Процент
          Вып_СМС_Факт_Процент = Round((Cells(intR, intC).Value), 2)
          ' Заносим данные в мою книгу
          ThisWorkbook.Sheets("Лист1").Range("M" + CStr(9 + CountMRK)).Value = Вып_СМС_Факт_Процент
          ' Красим цвет текста в ячейке в Светофора
          ' Call Full_Color_Text("Лист1", "M" + CStr(9 + CountMRK), Вып_СМС_Факт_Процент)
        End If

        ' 12 показатель: Если адрес ячейки $BK$8 - значит Интегральный_рейтинг_Процент
        ' If (Mid(cell.Address, 1, 4) = "$BK$") And (ThiStringIs072 = True) Then
        ' If (Mid(cell.Address, 1, 4) = Столбец_с_Интегральным_рейтингом_откуда_берем_данные) And (ThiStringIs072 = True) Then
        If (Наименование_столбца(Cell.Address) = Столбец_с_Интегральным_рейтингом_откуда_берем_данные) And (ThiStringIs072 = True) Then
          ' Устанавливаем формат ячейки - процентный "0.00%"
          ThisWorkbook.Sheets("Лист1").Range("N" + CStr(9 + CountMRK)).NumberFormat = "0%"
          ' Копируем формат ячейки
          ' Cells(intR, intC).Copy Destination:=ThisWorkbook.Sheets("Лист1").Range("N" + CStr(9 + CountMRK))
            
          ' Берем Интегральный_рейтинг_Процент
          Интегральный_рейтинг_Процент = Round((Cells(intR, intC).Value), 2)
          ' Заносим данные в мою книгу
          ThisWorkbook.Sheets("Лист1").Range("N" + CStr(9 + CountMRK)).Value = Интегральный_рейтинг_Процент
            
          ' Красим цвет текста в ячейке в Светофора
          ' Call Full_Color_Text("Лист1", "N" + CStr(9 + CountMRK), Интегральный_рейтинг_Процент)
          Call Full_Color_Range_For_Int_Rating("Лист1", "N" + CStr(9 + CountMRK), Интегральный_рейтинг_Процент)
      
        End If
        
        ' --- Секция разбора столбцов откуда брать данные ---
        ' Находим столбец "Интегральный рейтинг"
        If (CStr(Cells(intR, intC).Value) = "Интегральный рейтинг") And (Столбец_с_Интегральным_рейтингом_откуда_берем_данные = "") Then
          If Логирование_в_текстовые_файлы = True Then
            Print #1, "Найден Столбец_с_Интегральным_рейтингом_откуда_берем_данные"
          End If
          ' Берем колонку с месяцем
          ' Столбец_с_Интегральным_рейтингом_откуда_берем_данные = Mid(cell.Address, 1, 4)
          Столбец_с_Интегральным_рейтингом_откуда_берем_данные = Наименование_столбца(Cell.Address)
          If Логирование_в_текстовые_файлы = True Then
            Print #1, "Столбец_с_Интегральным_рейтингом_откуда_берем_данные: " + Столбец_с_Интегральным_рейтингом_откуда_берем_данные
          End If
        End If

        ' Столбец_ФИО_откуда_берем_данные If Mid(cell.Address, 1, 3) = "$B$" Then
        If (CStr(Cells(intR, intC).Value) = "ФИО") And (Столбец_ФИО_откуда_берем_данные = "") Then
          If Логирование_в_текстовые_файлы = True Then
            Print #1, "Найден Столбец_ФИО_откуда_берем_данные"
          End If
          ' Берем колонку с месяцем
          
          ' Столбец_ФИО_откуда_берем_данные = Mid(cell.Address, 1, 3)
          Столбец_ФИО_откуда_берем_данные = Наименование_столбца(Cell.Address)
        
          If Логирование_в_текстовые_файлы = True Then
            Print #1, "Столбец_ФИО_откуда_берем_данные: " + Столбец_ФИО_откуда_берем_данные
          End If
        End If
        
        ' Столбец_РОО_DP3_отчет_откуда_берем_данные If (Mid(cell.Address, 1, 3) = "$D$") Then
        If (CStr(Cells(intR, intC).Value) = "DP3_отчет") And (Столбец_РОО_DP3_отчет_откуда_берем_данные = "") Then
          If Логирование_в_текстовые_файлы = True Then
            Print #1, "Столбец_РОО_DP3_отчет_откуда_берем_данные"
          End If
          ' Берем колонку с месяцем
          ' Столбец_РОО_DP3_отчет_откуда_берем_данные = Mid(cell.Address, 1, 3)
          Столбец_РОО_DP3_отчет_откуда_берем_данные = Наименование_столбца(Cell.Address)
          
          If Логирование_в_текстовые_файлы = True Then
            Print #1, "Столбец_РОО_DP3_отчет_откуда_берем_данные: " + Столбец_РОО_DP3_отчет_откуда_берем_данные
          End If
        End If
                
        ' Столбец_ОО2_DP4_отчет_откуда_берем_данные If (Mid(cell.Address, 1, 3) = "$E$") And (ThiStringIs072 = True) Then
        If (CStr(Cells(intR, intC).Value) = "DP4_отчет") And (Столбец_ОО2_DP4_отчет_откуда_берем_данные = "") Then
          If Логирование_в_текстовые_файлы = True Then
            Print #1, "Столбец_ОО2_DP4_отчет_откуда_берем_данные"
          End If
          ' Берем колонку с месяцем
          ' Столбец_ОО2_DP4_отчет_откуда_берем_данные = Mid(cell.Address, 1, 3)
          Столбец_ОО2_DP4_отчет_откуда_берем_данные = Наименование_столбца(Cell.Address)
          
          If Логирование_в_текстовые_файлы = True Then
            Print #1, "Столбец_ОО2_DP4_отчет_откуда_берем_данные: " + Столбец_ОО2_DP4_отчет_откуда_берем_данные
          End If
        End If

        ' Столбец_ПК_Факт_откуда_берем_данные If (Mid(cell.Address, 1, 3) = "$J$") And (ThiStringIs072 = True) Then
        If (Trim(CStr(Cells(intR, intC).Value)) = "% Вып_ПК_Факт") And (Столбец_ПК_Факт_откуда_берем_данные = "") Then
          If Логирование_в_текстовые_файлы = True Then
            Print #1, "Столбец_ПК_Факт_откуда_берем_данные"
          End If
          ' Берем колонку с месяцем
          ' Столбец_ПК_Факт_откуда_берем_данные = Mid(cell.Address, 1, 3)
          Столбец_ПК_Факт_откуда_берем_данные = Наименование_столбца(Cell.Address)
          
          If Логирование_в_текстовые_файлы = True Then
            Print #1, "Столбец_ПК_Факт_откуда_берем_данные: " + Столбец_ПК_Факт_откуда_берем_данные
          End If
        End If
        
        ' Столбец_Вып_Страховки_к_ПК_Факт_откуда_берем_данные If (Mid(cell.Address, 1, 3) = "$O$") And (ThiStringIs072 = True) Then
        If (Trim(CStr(Cells(intR, intC).Value)) = "% Вып_Страховки к ПК_Факт") And (Столбец_Вып_Страховки_к_ПК_Факт_откуда_берем_данные = "") Then
          If Логирование_в_текстовые_файлы = True Then
            Print #1, "Столбец_Вып_Страховки_к_ПК_Факт_откуда_берем_данные"
          End If
          ' Берем колонку с месяцем
          ' Столбец_Вып_Страховки_к_ПК_Факт_откуда_берем_данные = Mid(cell.Address, 1, 3)
          Столбец_Вып_Страховки_к_ПК_Факт_откуда_берем_данные = Наименование_столбца(Cell.Address)
          If Логирование_в_текстовые_файлы = True Then
            Print #1, "Столбец_Вып_Страховки_к_ПК_Факт_откуда_берем_данные: " + Столбец_Вып_Страховки_к_ПК_Факт_откуда_берем_данные
          End If
        End If

        ' Столбец_Вып_КК_Факт_откуда_берем_данные If (Mid(cell.Address, 1, 3) = "$T$") And (ThiStringIs072 = True) Then
        If (Trim(CStr(Cells(intR, intC).Value)) = "% Вып_КК _Факт") And (Столбец_Вып_КК_Факт_откуда_берем_данные = "") Then
          If Логирование_в_текстовые_файлы = True Then
            Print #1, "Столбец_Вып_КК_Факт_откуда_берем_данные"
          End If
          ' Берем колонку с месяцем
          ' Столбец_Вып_КК_Факт_откуда_берем_данные = Mid(cell.Address, 1, 3)
          Столбец_Вып_КК_Факт_откуда_берем_данные = Наименование_столбца(Cell.Address)
          If Логирование_в_текстовые_файлы = True Then
            Print #1, "Столбец_Вып_КК_Факт_откуда_берем_данные: " + Столбец_Вып_КК_Факт_откуда_берем_данные
          End If
        End If

        ' Столбец_Вып_ДК_Факт_откуда_берем_данные If (Mid(cell.Address, 1, 3) = "$Y$") And (ThiStringIs072 = True) Then
        If (Trim(CStr(Cells(intR, intC).Value)) = "% Вып_ДК _Факт") And (Столбец_Вып_ДК_Факт_откуда_берем_данные = "") Then
          If Логирование_в_текстовые_файлы = True Then
            Print #1, "Столбец_Вып_ДК_Факт_откуда_берем_данные"
          End If
          ' Берем колонку с месяцем
          ' Столбец_Вып_ДК_Факт_откуда_берем_данные = Mid(cell.Address, 1, 3)
          Столбец_Вып_ДК_Факт_откуда_берем_данные = Наименование_столбца(Cell.Address)
          If Логирование_в_текстовые_файлы = True Then
            Print #1, "Столбец_Вып_ДК_Факт_откуда_берем_данные: " + Столбец_Вып_ДК_Факт_откуда_берем_данные
          End If
        End If
        
        ' Столбец_Вып_ИБ_Факт_откуда_берем_данные If (Mid(cell.Address, 1, 4) = "$AD$") And (ThiStringIs072 = True) Then
        If (Trim(CStr(Cells(intR, intC).Value)) = "% Вып_ИБ _Факт") And (Столбец_Вып_ИБ_Факт_откуда_берем_данные = "") Then
          If Логирование_в_текстовые_файлы = True Then
            Print #1, "Столбец_Вып_ИБ_Факт_откуда_берем_данные"
          End If
          ' Берем колонку с месяцем
          ' Столбец_Вып_ИБ_Факт_откуда_берем_данные = Mid(cell.Address, 1, 4)
          Столбец_Вып_ИБ_Факт_откуда_берем_данные = Наименование_столбца(Cell.Address)
          If Логирование_в_текстовые_файлы = True Then
            Print #1, "Столбец_Вып_ИБ_Факт_откуда_берем_данные: " + Столбец_Вып_ИБ_Факт_откуда_берем_данные
          End If
        End If
        
        ' Столбец_Вып_НС_Факт_откуда_берем_данные If (Mid(cell.Address, 1, 4) = "$AI$") And (ThiStringIs072 = True) Then
        If (Trim(CStr(Cells(intR, intC).Value)) = "% Вып_НС _Факт") And (Столбец_Вып_НС_Факт_откуда_берем_данные = "") Then
          If Логирование_в_текстовые_файлы = True Then
            Print #1, "Столбец_Вып_НС_Факт_откуда_берем_данные"
          End If
          ' Берем колонку с месяцем
          ' Столбец_Вып_НС_Факт_откуда_берем_данные = Mid(cell.Address, 1, 4)
          Столбец_Вып_НС_Факт_откуда_берем_данные = Наименование_столбца(Cell.Address)
          If Логирование_в_текстовые_файлы = True Then
            Print #1, "Столбец_Вып_НС_Факт_откуда_берем_данные: " + Столбец_Вып_НС_Факт_откуда_берем_данные
          End If
        End If
        
        ' Столбец_Вып_ИСЖ_Факт_откуда_берем_данные If (Mid(cell.Address, 1, 4) = "$AN$") And (ThiStringIs072 = True) Then
        If (Trim(CStr(Cells(intR, intC).Value)) = "% Вып_ ИСЖ_Факт") And (Столбец_Вып_ИСЖ_Факт_откуда_берем_данные = "") Then
          If Логирование_в_текстовые_файлы = True Then
            Print #1, "Столбец_Вып_ИСЖ_Факт_откуда_берем_данные"
          End If
          ' Берем колонку с месяцем
          ' Столбец_Вып_ИСЖ_Факт_откуда_берем_данные = Mid(cell.Address, 1, 4)
          Столбец_Вып_ИСЖ_Факт_откуда_берем_данные = Наименование_столбца(Cell.Address)
          If Логирование_в_текстовые_файлы = True Then
            Print #1, "Столбец_Вып_ИСЖ_Факт_откуда_берем_данные: " + Столбец_Вып_ИСЖ_Факт_откуда_берем_данные
          End If
        End If
        
        ' Столбец_Вып_НСЖ_Факт_откуда_берем_данные If (Mid(cell.Address, 1, 4) = "$AS$") And (ThiStringIs072 = True) Then
        If (Trim(CStr(Cells(intR, intC).Value)) = "% Вып_НСЖ_Факт") And (Столбец_Вып_НСЖ_Факт_откуда_берем_данные = "") Then
          If Логирование_в_текстовые_файлы = True Then
            Print #1, "Столбец_Вып_НСЖ_Факт_откуда_берем_данные"
          End If
          ' Берем колонку с месяцем
          ' Столбец_Вып_НСЖ_Факт_откуда_берем_данные = Mid(cell.Address, 1, 4)
          Столбец_Вып_НСЖ_Факт_откуда_берем_данные = Наименование_столбца(Cell.Address)
          If Логирование_в_текстовые_файлы = True Then
            Print #1, "Столбец_Вып_НСЖ_Факт_откуда_берем_данные: " + Столбец_Вып_НСЖ_Факт_откуда_берем_данные
          End If
        End If
        
        ' Столбец_Вып_КС_Факт_откуда_берем_данные If (Mid(cell.Address, 1, 4) = "$AX$") And (ThiStringIs072 = True) Then
        If (Trim(CStr(Cells(intR, intC).Value)) = "% Вып_КС _Факт") And (Столбец_Вып_КС_Факт_откуда_берем_данные = "") Then
          If Логирование_в_текстовые_файлы = True Then
            Print #1, "Столбец_Вып_КС_Факт_откуда_берем_данные"
          End If
          ' Берем колонку с месяцем
          ' Столбец_Вып_КС_Факт_откуда_берем_данные = Mid(cell.Address, 1, 4)
          Столбец_Вып_КС_Факт_откуда_берем_данные = Наименование_столбца(Cell.Address)
          If Логирование_в_текстовые_файлы = True Then
            Print #1, "Столбец_Вып_КС_Факт_откуда_берем_данные: " + Столбец_Вып_КС_Факт_откуда_берем_данные
          End If
        End If
        
        ' Столбец_Вып_ЛА_Факт_откуда_берем_данные If (Mid(cell.Address, 1, 4) = "$BC$") And (ThiStringIs072 = True) Then
        If (Trim(CStr(Cells(intR, intC).Value)) = "% Вып_ ЛА") And (Столбец_Вып_ЛА_Факт_откуда_берем_данные = "") Then
          If Логирование_в_текстовые_файлы = True Then
            Print #1, "Столбец_Вып_ЛА_Факт_откуда_берем_данные"
          End If
          ' Берем колонку с месяцем
          ' Столбец_Вып_ЛА_Факт_откуда_берем_данные = Mid(cell.Address, 1, 4)
          Столбец_Вып_ЛА_Факт_откуда_берем_данные = Наименование_столбца(Cell.Address)
          If Логирование_в_текстовые_файлы = True Then
            Print #1, "Столбец_Вып_ЛА_Факт_откуда_берем_данные: " + Столбец_Вып_ЛА_Факт_откуда_берем_данные
          End If
        End If
        
        ' Столбец_Вып_СМС_Факт_откуда_берем_данные If (Mid(cell.Address, 1, 4) = "$BH$") And (ThiStringIs072 = True) Then
        If (Trim(CStr(Cells(intR, intC).Value)) = "% Вып_ СМС") And (Столбец_Вып_СМС_Факт_откуда_берем_данные = "") Then
          If Логирование_в_текстовые_файлы = True Then
            Print #1, "Столбец_Вып_СМС_Факт_откуда_берем_данные"
          End If
          ' Берем колонку с месяцем
          ' Столбец_Вып_СМС_Факт_откуда_берем_данные = Mid(cell.Address, 1, 4)
          Столбец_Вып_СМС_Факт_откуда_берем_данные = Наименование_столбца(Cell.Address)
          If Логирование_в_текстовые_файлы = True Then
            Print #1, "Столбец_Вып_СМС_Факт_откуда_берем_данные: " + Столбец_Вып_СМС_Факт_откуда_берем_данные
          End If
        End If
        
        ' --- Конец Секция разбора столбцов откуда брать данные ---

        ' ==== Окончание обработки строки ====
    
        ' Если в столбце D был "Тюменский ОО1", то записываем данные из ячеек строк в переменные
        If (ThiStringIs072 = True) Then
      
          ' Логирование
          If Логирование_в_текстовые_файлы = True Then
            ' Выводим данные в строку
            Print #2, Cell.Address, CStr(CountMRK) + "." + FIO_Name + ":" + Cell.Formula
          End If
      
        End If
        
        ' Выход из цикла при нахождении ячейки "Общий итог"
        ' If CStr(Cells(intR, intC).Value) = "Общий итог" Then
        '   MsgBox ("Выход из цикла по ячейке Общий итог")
        '   Exit For
        ' End If
    
      End If ' Если ячейка не пустая
      
  Next

  ' Вывод итогов обработки Интегрального рейтинга сотрудников
  RowToPrint = 2
  
  ' Лидеры продаж ПК
  If Список_МРК_Лидеры_продаж_ПК <> "" Then
    ThisWorkbook.Sheets("Лист1").Range("B" + CStr(9 + CountMRK + RowToPrint)).Value = "Лидеры продаж Потребительских кредитов: " + Список_МРК_Лидеры_продаж_ПК
    RowToPrint = RowToPrint + 2
  End If
  
  ' Лидеры продаж Страховок ПК
  If Список_МРК_Лидеры_продаж_СтраховокПК <> "" Then
    ThisWorkbook.Sheets("Лист1").Range("B" + CStr(9 + CountMRK + RowToPrint)).Value = "Лидеры продаж Страховок к Потребительским кредитам: " + Список_МРК_Лидеры_продаж_СтраховокПК
    RowToPrint = RowToPrint + 2
  End If
  
  ' Лидеры продаж КК
  If Список_МРК_Лидеры_продаж_КК <> "" Then
    ThisWorkbook.Sheets("Лист1").Range("B" + CStr(9 + CountMRK + RowToPrint)).Value = "Лидеры продаж Кредитных карт: " + Список_МРК_Лидеры_продаж_КК
    RowToPrint = RowToPrint + 2
  End If
  
  ' Лидеры продаж ДК
  If Список_МРК_Лидеры_продаж_ДК <> "" Then
    ThisWorkbook.Sheets("Лист1").Range("B" + CStr(9 + CountMRK + RowToPrint)).Value = "Лидеры продаж Дебетовых карт: " + Список_МРК_Лидеры_продаж_ДК
    RowToPrint = RowToPrint + 2
  End If
  
  ' Лидеры продаж ИСЖ
  If Список_МРК_Лидеры_продаж_ИСЖ <> "" Then
    ThisWorkbook.Sheets("Лист1").Range("B" + CStr(9 + CountMRK + RowToPrint)).Value = "Лидеры продаж ИСЖ: " + Список_МРК_Лидеры_продаж_ИСЖ
    RowToPrint = RowToPrint + 2
  End If
  
  ' Лидеры продаж НСЖ
  If Список_МРК_Лидеры_продаж_НСЖ <> "" Then
    ThisWorkbook.Sheets("Лист1").Range("B" + CStr(9 + CountMRK + RowToPrint)).Value = "Лидеры продаж НСЖ: " + Список_МРК_Лидеры_продаж_НСЖ
    RowToPrint = RowToPrint + 2
  End If
  
  ' Отсутствие продаж ПК
  If Список_МРК_с_Отсутствием_продаж_ПК <> "" Then
    ThisWorkbook.Sheets("Лист1").Range("B" + CStr(9 + CountMRK + RowToPrint)).Value = "Сотрудники без продаж Потребительских кредитов: " + Список_МРК_с_Отсутствием_продаж_ПК
    RowToPrint = RowToPrint + 2
  End If
  
  ' Отсутствие продаж ПК
  If Список_МРК_с_Отсутствием_продаж_СтраховокПК <> "" Then
    ThisWorkbook.Sheets("Лист1").Range("B" + CStr(9 + CountMRK + RowToPrint)).Value = "Сотрудники без продаж Страховок к Потребительским кредитам: " + Список_МРК_с_Отсутствием_продаж_СтраховокПК
    RowToPrint = RowToPrint + 2
  End If
  ' Отсутствие продаж КК
  If Список_МРК_с_Отсутствием_продаж_KK <> "" Then
    ThisWorkbook.Sheets("Лист1").Range("B" + CStr(9 + CountMRK + RowToPrint)).Value = "Сотрудники без продаж Кредитных карт: " + Список_МРК_с_Отсутствием_продаж_KK
    RowToPrint = RowToPrint + 2
  End If
  ' Отсутствие продаж ДК
  If Список_МРК_с_Отсутствием_продаж_ДK <> "" Then
    ThisWorkbook.Sheets("Лист1").Range("B" + CStr(9 + CountMRK + RowToPrint)).Value = "Сотрудники без продаж Дебетовых карт: " + Список_МРК_с_Отсутствием_продаж_ДK
    RowToPrint = RowToPrint + 2
  End If
  ' Отсутствие продаж ИСЖ
  If Список_МРК_с_Отсутствием_продаж_ИСЖ <> "" Then
    ThisWorkbook.Sheets("Лист1").Range("B" + CStr(9 + CountMRK + RowToPrint)).Value = "Сотрудники без продаж ИСЖ: " + Список_МРК_с_Отсутствием_продаж_ИСЖ
    RowToPrint = RowToPrint + 2
  End If
  ' Отсутствие продаж НСЖ
  If Список_МРК_с_Отсутствием_продаж_НСЖ <> "" Then
    ThisWorkbook.Sheets("Лист1").Range("B" + CStr(9 + CountMRK + RowToPrint)).Value = "Сотрудники без продаж НСЖ: " + Список_МРК_с_Отсутствием_продаж_НСЖ
    RowToPrint = RowToPrint + 2
  End If

  ' Логирование
  If Логирование_в_текстовые_файлы = True Then
    ' Закрываем файлы
    Close #1
    Close #2
  End If

End Sub

' В передаваемой ячейке убираем рамви - верхняя, нижняя, левая, правая
Sub Убираем_рамки_в_ячейке(In_Sheet As String, In_Cell As String)
  ThisWorkbook.Sheets(In_Sheet).Range(In_Cell).Borders(xlDiagonalDown).LineStyle = xlNone
  ThisWorkbook.Sheets(In_Sheet).Range(In_Cell).Borders(xlDiagonalUp).LineStyle = xlNone
  ThisWorkbook.Sheets(In_Sheet).Range(In_Cell).Borders(xlEdgeLeft).LineStyle = xlNone
  ThisWorkbook.Sheets(In_Sheet).Range(In_Cell).Borders(xlEdgeTop).LineStyle = xlNone
  ThisWorkbook.Sheets(In_Sheet).Range(In_Cell).Borders(xlEdgeBottom).LineStyle = xlNone
  ThisWorkbook.Sheets(In_Sheet).Range(In_Cell).Borders(xlEdgeRight).LineStyle = xlNone
  ThisWorkbook.Sheets(In_Sheet).Range(In_Cell).Borders(xlInsideVertical).LineStyle = xlNone
  ThisWorkbook.Sheets(In_Sheet).Range(In_Cell).Borders(xlInsideHorizontal).LineStyle = xlNone
End Sub

Sub Убираем_условное_форматирование_в_ячейке(In_Sheet As String, In_Cell As String)
  
  ThisWorkbook.Sheets(In_Sheet).Range(In_Cell).FormatConditions.Delete
  
End Sub

' Выборка данных с листа 1.1Интег-ый рейтинг  по офисам
Sub Интегральный_рейтинг_по_офисам(In_DBstrName As String)

Dim Office2_Name As String
' Логирование в текстовый файл чтения DB
Dim Логирование_в_текстовые_файлы As Boolean
' Факт выполнения показателей
Dim ПК_сеть, Ипотека, ИБ, НС, Комдоход, Офисные_продажи, ЗП_шт, OPC_шт, Интегр_рейтинг As Double
Dim Строка_перевыполнение_показателей_по, Строка_недовыполнения_по_офису_показателей_по, Строка_перевыполнение_по_офису_показателей, Строка_выполнение_менее_чем_на_половину As String
Dim Наименование_листа_ИнтРейтОфисов As String
Dim номерСтрокиЗаголовков, номерСтолбцаФилиал, номерСтолбцаРейтинг_офиса, номерСтолбцаНаименованиеОфиса, номерСтолбцаПК_сеть, номерСтолбцаИпотека, номерСтолбцаИБ, номерСтолбцаНС, номерСтолбцаКомдоход, номерСтолбцаОфисные_продажи, номерСтолбцаЗП_шт, номерСтолбцаOPC_шт, номерСтолбцаИнтегр_рейтинг As Byte

  Application.StatusBar = "Интегральный_рейтинг_по_офисам ..."

  ' Есть такой лист в Книге?
  If (Sheets_Exist(ActiveWorkbook, "1.1Интег-ый рейтинг  по офисам") = True) Or (Sheets_Exist(ActiveWorkbook, "1.1 Интег-ый рейтинг  по офисам") = True) Then

  ' Логирование в текстовый файл чтения DB: True - логируем, False - не логируем
  Логирование_в_текстовые_файлы = False

  ' Перейти на "1.1Интег-ый рейтинг  по офисам"
  If (Sheets_Exist(ActiveWorkbook, "1.1Интег-ый рейтинг  по офисам") = True) Then
    Sheets("1.1Интег-ый рейтинг  по офисам").Select
    Наименование_листа_ИнтРейтОфисов = "1.1Интег-ый рейтинг  по офисам"
  End If
  ' Второй вариант записи
  If (Sheets_Exist(ActiveWorkbook, "1.1 Интег-ый рейтинг  по офисам") = True) Then
    Sheets("1.1 Интег-ый рейтинг  по офисам").Select
    Наименование_листа_ИнтРейтОфисов = "1.1 Интег-ый рейтинг  по офисам"
  End If

  ' Логирование в текстовые файлы
  If Логирование_в_текстовые_файлы = True Then

    ' В файл выводим все:
    MyFile1 = In_DBstrName & "_Инт_рейт_офисы_1_log.txt"
    Open MyFile1 For Output As #1

    ' Второй вариант имени файла - вывод в конкретный каталог
    MyFile2 = In_DBstrName & "_Инт_рейт_офисы_2_log.txt"
    ' Открыли для записи
    Open MyFile2 For Output As #2

  End If

  Строка_перевыполнение_по_офису_показателей = ""
  Строка_выполнение_менее_чем_на_половину = ""

  ' Инициализация переменных
  ' Это строка 072?
  ThiStringIs072 = False
  ' Счетчик МРК
  CountOffice = 0
  ' Наименование офиса 2-го уровня
  Office2_Name = ""
  ' Вывод в мой файл строки
  Вывод_офиса_в_мой_файл = False

  ' Номер строки с заголовками: Дельта по сравнению с предыдущим кварталом  № п/п   ID_PROFIT   Офис    Филиал  % ПК_Сеть   % ПК_DSA     % Ипотека   % ДК    % КК    % ИБ    % Ком.доход    Интегральный рейтинг_Офисные продажи     % ЗП_шт.   % OPC_шт.   Интегральный рейтинг *  Интегральный рейтинг_Предыдущий квартал Динамика
  номерСтрокиЗаголовков = rowByValue(In_DBstrName, Наименование_листа_ИнтРейтОфисов, "Дельта по сравнению с предыдущим кварталом", 40, 250)
  ' Филиал
  номерСтолбцаФилиал = ColumnByNameAndNumber(In_DBstrName, Наименование_листа_ИнтРейтОфисов, номерСтрокиЗаголовков, "Филиал", 1, 60)
  ' Рейтинг_офиса
  номерСтолбцаРейтинг_офиса = ColumnByNameAndNumber(In_DBstrName, Наименование_листа_ИнтРейтОфисов, номерСтрокиЗаголовков, "№ п/п", 1, 60)
  ' номерСтолбцаНаименованиеОфиса
  номерСтолбцаНаименованиеОфиса = ColumnByNameAndNumber(In_DBstrName, Наименование_листа_ИнтРейтОфисов, номерСтрокиЗаголовков, "Офис", 1, 60)
  ' номерСтолбцаПК_сеть (в значениях стоят формулы, поэтому можно от филиала прибавлять число и получать столбец)
  номерСтолбцаПК_сеть = номерСтолбцаФилиал + 1
  ' номерСтолбцаИпотека
  номерСтолбцаИпотека = номерСтолбцаФилиал + 3
  ' номерСтолбцаИБ +6
  номерСтолбцаИБ = номерСтолбцаФилиал + 6
  ' номерСтолбцаНС
  номерСтолбцаНС = номерСтолбцаФилиал + 14
  ' номерСтолбцаКомдоход
  номерСтолбцаКомдоход = номерСтолбцаФилиал + 7
  ' номерСтолбцаОфисные_продажи
  номерСтолбцаОфисные_продажи = номерСтолбцаФилиал + 8
  ' номерСтолбцаЗП_шт
  номерСтолбцаЗП_шт = номерСтолбцаФилиал + 9
  ' номерСтолбцаOPC_шт
  номерСтолбцаOPC_шт = номерСтолбцаФилиал + 10
  ' номерСтолбцаИнтегр_рейтинг
  номерСтолбцаИнтегр_рейтинг = номерСтолбцаФилиал + 11

  ' Цикл обработки Листа Excel
  For Each Cell In ActiveSheet.UsedRange
  
      ' Если ячейка не пустая
      If Not IsEmpty(Cell) Then

        ' Выводим все данные в строку
        If Логирование_в_текстовые_файлы = True Then
          Print #1, Cell.Address, ":" + Cell.Formula
        End If

        ' Вывод строки и столбца
        ' Номер столбца
        intC = Cell.Column
        ' Номер строки
        intR = Cell.Row

        ' Если новый столбец $C$, то считаем, что строка не "Тюменский ОО1"
        ' If Mid(Cell.Address, 1, 3) = "$C$" Then
        ' If номерСтолбцаФилиал = intC Then
        If номерСтолбцаНаименованиеОфиса = intC Then
            ' Офис не Тюменский OO1
            ThiStringIs072 = False
            ' Вывод Офиса в мой файл
            Вывод_офиса_в_мой_файл = False
        End If

        ' Рейтинг_офиса - $C$
        ' If Mid(Cell.Address, 1, 3) = "$C$" Then
        If номерСтолбцаРейтинг_офиса = intC Then
          Рейтинг_офиса = CStr(Cells(intR, intC).Value)
        End If

        ' Если адрес ячейки $E$XXX - значит это Наименование офиса, записываем в переменную (для любого офиса)
        ' If Mid(Cell.Address, 1, 3) = "$E$" Then
        ' If номерСтолбцаНаименованиеОфиса = intC Then
        If номерСтолбцаФилиал = intC Then
          Office2_Name = CStr(Cells(intR, intC).Value)
        End If
                        
        ' Если это столбец F ("Тюменский ОО1")
        ' If (Mid(Cell.Address, 1, 3) = "$F$") Then
        ' If номерСтолбцаФилиал = intC Then
        If номерСтолбцаНаименованиеОфиса = intC Then
          If (CStr(Cells(intR, intC).Value) = "Тюменский ОО1") Then
            ' Если офис Тюменский OO1
            ThiStringIs072 = True
            ' Счетчик
            CountOffice = CountOffice + 1
            ' Вывод_Номера_и_ФИО_сотрудника
            Вывод_офиса_в_мой_файл = True
          Else
            ' Если офис не Тюменский OO1
            ThiStringIs072 = False
          End If
        End If

        ' Если это Тюменский РОО, то след столбец это наименование офиса
        If (CStr(Cells(intR, номерСтолбцаНаименованиеОфиса).Value) = "Тюменский ОО1") Then
          Office2_Name = CStr(Cells(intR, номерСтолбцаНаименованиеОфиса + 1).Value)
        End If

        ' Выводим в Мою книгу Рейтинг и Офис
        If (Вывод_офиса_в_мой_файл = True) And (ThiStringIs072 = True) Then
          ThisWorkbook.Sheets("Лист1").Range("A" + CStr(51 + CountOffice)).Value = Рейтинг_офиса
          ThisWorkbook.Sheets("Лист1").Range("B" + CStr(51 + CountOffice)).Value = Office2_Name
          Вывод_офиса_в_мой_файл = False
          ' Обнуляем подстроку выводов о работе офиса
          Строка_перевыполнение_показателей_по = ""
          Строка_недовыполнения_по_офису_показателей_по = ""
        End If


        ' 1 показатель: Если адрес ячейки $G$ - значит это ПК_сеть
        ' If (Mid(Cell.Address, 1, 3) = "$G$") And (ThiStringIs072 = True) Then
        If (номерСтолбцаПК_сеть = intC) And (ThiStringIs072 = True) Then
      
          ' Устанавливаем формат ячейки - процентный "0.00%"
          ' ThisWorkbook.Sheets("Лист1").Range("C" + CStr(9 + CountMRK)).NumberFormat = "0%"
          ' Копируем формат ячейки
          Cells(intR, intC).Copy Destination:=ThisWorkbook.Sheets("Лист1").Range("C" + CStr(51 + CountOffice))
          ' Берем ПК_сеть (Значение в 0,92)
          ПК_сеть = Round((Cells(intR, intC).Value), 2)
          ' Заносим данные в мою книгу
          ThisWorkbook.Sheets("Лист1").Range("C" + CStr(51 + CountOffice)).Value = ПК_сеть
           ' Убираем рамки
          Call Убираем_рамки_в_ячейке("Лист1", "C" + CStr(51 + CountOffice))
          ' Если показатель ПК выполняется, то
          If ПК_сеть >= 1 Then
            ' Если в строке что-то есть, то ставим ", "
            If Строка_перевыполнение_показателей_по <> "" Then
              Строка_перевыполнение_показателей_по = Строка_перевыполнение_показателей_по + ", "
            End If
            Строка_перевыполнение_показателей_по = Строка_перевыполнение_показателей_по + "Потребительскому кредитованию"
          End If
          ' Строка_выполнение_менее_чем_на_половину
          If ПК_сеть < 0.5 Then
            ' Если в строке что-то есть, то ставим ", "
            If Строка_недовыполнения_по_офису_показателей_по <> "" Then
              Строка_недовыполнения_по_офису_показателей_по = Строка_недовыполнения_по_офису_показателей_по + ", "
            End If
            Строка_недовыполнения_по_офису_показателей_по = Строка_недовыполнения_по_офису_показателей_по + "Потребительскому кредитованию"
          End If
        End If
      
        ' 2 показатель: Ипотека
        ' If (Mid(Cell.Address, 1, 3) = "$I$") And (ThiStringIs072 = True) Then
        If (номерСтолбцаИпотека = intC) And (ThiStringIs072 = True) Then
          ' Устанавливаем формат ячейки - процентный "0.00%"
          ' ThisWorkbook.Sheets("Лист1").Range("C" + CStr(9 + CountMRK)).NumberFormat = "0%"
          ' Копируем формат ячейки
          Cells(intR, intC).Copy Destination:=ThisWorkbook.Sheets("Лист1").Range("D" + CStr(51 + CountOffice))
          ' Берем Ипотека
          Ипотека = Round((Cells(intR, intC).Value), 2)
          ' Заносим данные в мою книгу
          ThisWorkbook.Sheets("Лист1").Range("D" + CStr(51 + CountOffice)).Value = Ипотека
           ' Убираем рамки
          Call Убираем_рамки_в_ячейке("Лист1", "D" + CStr(51 + CountOffice))
        End If
        
        ' 3 показатель ИБ
        ' If (Mid(Cell.Address, 1, 3) = "$J$") And (ThiStringIs072 = True) Then
        If (номерСтолбцаИБ = intC) And (ThiStringIs072 = True) Then
          
          ' Устанавливаем формат ячейки - процентный "0.00%"
          ' ThisWorkbook.Sheets("Лист1").Range("C" + CStr(9 + CountMRK)).NumberFormat = "0%"
          ' Копируем формат ячейки
          Cells(intR, intC).Copy Destination:=ThisWorkbook.Sheets("Лист1").Range("E" + CStr(51 + CountOffice))
          ' Берем ИБ
          ИБ = Round((Cells(intR, intC).Value), 2)
          ' Заносим данные в мою книгу
          ThisWorkbook.Sheets("Лист1").Range("E" + CStr(51 + CountOffice)).Value = ИБ
           ' Убираем рамки
          Call Убираем_рамки_в_ячейке("Лист1", "E" + CStr(51 + CountOffice))
          ' Если показатель ПК выполняется, то
          If ИБ >= 1 Then
            ' Если в строке что-то есть, то ставим ", "
            If Строка_перевыполнение_показателей_по <> "" Then
              Строка_перевыполнение_показателей_по = Строка_перевыполнение_показателей_по + ", "
            End If
            Строка_перевыполнение_показателей_по = Строка_перевыполнение_показателей_по + "Интернет-Банку"
          End If
          ' Строка_выполнение_менее_чем_на_половину
          If ИБ < 0.5 Then
            ' Если в строке что-то есть, то ставим ", "
            If Строка_недовыполнения_по_офису_показателей_по <> "" Then
              Строка_недовыполнения_по_офису_показателей_по = Строка_недовыполнения_по_офису_показателей_по + ", "
            End If
            Строка_недовыполнения_по_офису_показателей_по = Строка_недовыполнения_по_офису_показателей_по + "Интернет-Банку"
          End If
        End If
        
        ' 4 показатель НС
        ' If (Mid(Cell.Address, 1, 3) = "$K$") And (ThiStringIs072 = True) Then
        If (номерСтолбцаНС = intC) And (ThiStringIs072 = True) Then
          
          ' Устанавливаем формат ячейки - процентный "0.00%"
          ' ThisWorkbook.Sheets("Лист1").Range("C" + CStr(9 + CountMRK)).NumberFormat = "0%"
          ' Копируем формат ячейки
          Cells(intR, intC).Copy Destination:=ThisWorkbook.Sheets("Лист1").Range("F" + CStr(51 + CountOffice))
          ' Берем НС
          НС = Round((Cells(intR, intC).Value), 2)
          ' Заносим данные в мою книгу
          ThisWorkbook.Sheets("Лист1").Range("F" + CStr(51 + CountOffice)).Value = НС
           ' Убираем рамки
          Call Убираем_рамки_в_ячейке("Лист1", "F" + CStr(51 + CountOffice))
          ' Если показатель ПК выполняется, то
          If НС >= 1 Then
            ' Если в строке что-то есть, то ставим ", "
            If Строка_перевыполнение_показателей_по <> "" Then
              Строка_перевыполнение_показателей_по = Строка_перевыполнение_показателей_по + ", "
            End If
            Строка_перевыполнение_показателей_по = Строка_перевыполнение_показателей_по + "Накопительным счетам"
          End If
          ' Строка_выполнение_менее_чем_на_половину
          If НС < 0.5 Then
            ' Если в строке что-то есть, то ставим ", "
            If Строка_недовыполнения_по_офису_показателей_по <> "" Then
              Строка_недовыполнения_по_офису_показателей_по = Строка_недовыполнения_по_офису_показателей_по + ", "
            End If
            Строка_недовыполнения_по_офису_показателей_по = Строка_недовыполнения_по_офису_показателей_по + "Накопительным счетам"
          End If
        End If
        
        ' 5 показатель Комдоход
        ' If (Mid(Cell.Address, 1, 3) = "$L$") And (ThiStringIs072 = True) Then
        If (номерСтолбцаКомдоход = intC) And (ThiStringIs072 = True) Then
          
          ' Устанавливаем формат ячейки - процентный "0.00%"
          ' ThisWorkbook.Sheets("Лист1").Range("C" + CStr(9 + CountMRK)).NumberFormat = "0%"
          ' Копируем формат ячейки
          Cells(intR, intC).Copy Destination:=ThisWorkbook.Sheets("Лист1").Range("G" + CStr(51 + CountOffice))
          ' Берем Комдоход
          Комдоход = Round((Cells(intR, intC).Value), 2)
          ' Заносим данные в мою книгу
          ThisWorkbook.Sheets("Лист1").Range("G" + CStr(51 + CountOffice)).Value = Комдоход
           ' Убираем рамки
          Call Убираем_рамки_в_ячейке("Лист1", "G" + CStr(51 + CountOffice))
          ' Если показатель ПК выполняется, то
          If Комдоход >= 1 Then
            ' Если в строке что-то есть, то ставим ", "
            If Строка_перевыполнение_показателей_по <> "" Then
              Строка_перевыполнение_показателей_по = Строка_перевыполнение_показателей_по + ", "
            End If
            Строка_перевыполнение_показателей_по = Строка_перевыполнение_показателей_по + "Комиссионному доходу"
          End If
          ' Строка_выполнение_менее_чем_на_половину
          If Комдоход < 0.5 Then
            ' Если в строке что-то есть, то ставим ", "
            If Строка_недовыполнения_по_офису_показателей_по <> "" Then
              Строка_недовыполнения_по_офису_показателей_по = Строка_недовыполнения_по_офису_показателей_по + ", "
            End If
            Строка_недовыполнения_по_офису_показателей_по = Строка_недовыполнения_по_офису_показателей_по + "Комиссионному доходу"
          End If

        End If
        
        ' 6 показатель Офисные_продажи
        ' If (Mid(Cell.Address, 1, 3) = "$M$") And (ThiStringIs072 = True) Then
        If (номерСтолбцаОфисные_продажи = intC) And (ThiStringIs072 = True) Then
          ' Устанавливаем формат ячейки - процентный "0.00%"
          ' ThisWorkbook.Sheets("Лист1").Range("C" + CStr(9 + CountMRK)).NumberFormat = "0%"
          ' Копируем формат ячейки
          Cells(intR, intC).Copy Destination:=ThisWorkbook.Sheets("Лист1").Range("H" + CStr(51 + CountOffice))
          ' Берем Офисные_продажи
          Офисные_продажи = Round((Cells(intR, intC).Value), 2)
          ' Заносим данные в мою книгу
          ThisWorkbook.Sheets("Лист1").Range("H" + CStr(51 + CountOffice)).Value = Офисные_продажи
          ' Убираем рамки
          Call Убираем_рамки_в_ячейке("Лист1", "H" + CStr(51 + CountOffice))
          ' Если показатель ПК выполняется, то
          If Офисные_продажи >= 1 Then
            ' Если в строке что-то есть, то ставим ", "
            If Строка_перевыполнение_показателей_по <> "" Then
              Строка_перевыполнение_показателей_по = Строка_перевыполнение_показателей_по + ", "
            End If
            Строка_перевыполнение_показателей_по = Строка_перевыполнение_показателей_по + "Офисным продажам"
          End If
          ' Строка_выполнение_менее_чем_на_половину
          If Офисные_продажи < 0.5 Then
            ' Если в строке что-то есть, то ставим ", "
            If Строка_недовыполнения_по_офису_показателей_по <> "" Then
              Строка_недовыполнения_по_офису_показателей_по = Строка_недовыполнения_по_офису_показателей_по + ", "
            End If
            Строка_недовыполнения_по_офису_показателей_по = Строка_недовыполнения_по_офису_показателей_по + "Офисным продажам"
          End If

        End If
        
        ' 7 показатель ЗП_шт
        ' If (Mid(Cell.Address, 1, 3) = "$N$") And (ThiStringIs072 = True) Then
        If (номерСтолбцаЗП_шт = intC) And (ThiStringIs072 = True) Then
          ' Устанавливаем формат ячейки - процентный "0.00%"
          ' ThisWorkbook.Sheets("Лист1").Range("C" + CStr(9 + CountMRK)).NumberFormat = "0%"
          ' Копируем формат ячейки
          Cells(intR, intC).Copy Destination:=ThisWorkbook.Sheets("Лист1").Range("I" + CStr(51 + CountOffice))
          ' Берем ЗП_шт
          ЗП_шт = Round((Cells(intR, intC).Value), 2)
          ' Заносим данные в мою книгу
          ThisWorkbook.Sheets("Лист1").Range("I" + CStr(51 + CountOffice)).Value = ЗП_шт
           ' Убираем рамки
          Call Убираем_рамки_в_ячейке("Лист1", "I" + CStr(51 + CountOffice))
          ' Если показатель ПК выполняется, то
          If ЗП_шт >= 1 Then
            ' Если в строке что-то есть, то ставим ", "
            If Строка_перевыполнение_показателей_по <> "" Then
              Строка_перевыполнение_показателей_по = Строка_перевыполнение_показателей_по + ", "
            End If
            Строка_перевыполнение_показателей_по = Строка_перевыполнение_показателей_по + "Зарплатным картам"
          End If
          ' Строка_выполнение_менее_чем_на_половину
          If ЗП_шт < 0.5 Then
            ' Если в строке что-то есть, то ставим ", "
            If Строка_недовыполнения_по_офису_показателей_по <> "" Then
              Строка_недовыполнения_по_офису_показателей_по = Строка_недовыполнения_по_офису_показателей_по + ", "
            End If
            Строка_недовыполнения_по_офису_показателей_по = Строка_недовыполнения_по_офису_показателей_по + "Зарплатным картам"
          End If

        End If
        
        ' 8 показатель OPC_шт
        ' If (Mid(Cell.Address, 1, 3) = "$O$") And (ThiStringIs072 = True) Then
        If (номерСтолбцаOPC_шт = intC) And (ThiStringIs072 = True) Then
          ' Устанавливаем формат ячейки - процентный "0.00%"
          ' ThisWorkbook.Sheets("Лист1").Range("C" + CStr(9 + CountMRK)).NumberFormat = "0%"
          ' Копируем формат ячейки
          Cells(intR, intC).Copy Destination:=ThisWorkbook.Sheets("Лист1").Range("J" + CStr(51 + CountOffice))
          ' Берем OPC_шт
          OPC_шт = Round((Cells(intR, intC).Value), 2)
          ' Заносим данные в мою книгу
          ThisWorkbook.Sheets("Лист1").Range("J" + CStr(51 + CountOffice)).Value = OPC_шт
           ' Убираем рамки
          Call Убираем_рамки_в_ячейке("Лист1", "J" + CStr(51 + CountOffice))
          ' Если показатель ПК выполняется, то
          If OPC_шт >= 1 Then
            ' Если в строке что-то есть, то ставим ", "
            If Строка_перевыполнение_показателей_по <> "" Then
              Строка_перевыполнение_показателей_по = Строка_перевыполнение_показателей_по + ", "
            End If
            Строка_перевыполнение_показателей_по = Строка_перевыполнение_показателей_по + "OPC"
          End If
          ' Строка_выполнение_менее_чем_на_половину
          If OPC_шт < 0.5 Then
            ' Если в строке что-то есть, то ставим ", "
            If Строка_недовыполнения_по_офису_показателей_по <> "" Then
              Строка_недовыполнения_по_офису_показателей_по = Строка_недовыполнения_по_офису_показателей_по + ", "
            End If
            Строка_недовыполнения_по_офису_показателей_по = Строка_недовыполнения_по_офису_показателей_по + "OPC"
          End If

        End If
        
        ' 9 показатель Интегр_рейтинг
        ' If (Mid(Cell.Address, 1, 3) = "$P$") And (ThiStringIs072 = True) Then
        If (номерСтолбцаИнтегр_рейтинг = intC) And (ThiStringIs072 = True) Then
          ' Устанавливаем формат ячейки - процентный "0.00%"
          ' ThisWorkbook.Sheets("Лист1").Range("C" + CStr(9 + CountMRK)).NumberFormat = "0%"
          ' Копируем формат ячейки
          Cells(intR, intC).Copy Destination:=ThisWorkbook.Sheets("Лист1").Range("K" + CStr(51 + CountOffice))
          ' Берем Интегр_рейтинг
          Интегр_рейтинг = Round((Cells(intR, intC).Value), 2)
          ' Заносим данные в мою книгу
          ThisWorkbook.Sheets("Лист1").Range("K" + CStr(51 + CountOffice)).Value = Интегр_рейтинг
          ' Убираем рамки
          Call Убираем_рамки_в_ячейке("Лист1", "K" + CStr(51 + CountOffice))
          
          ' Подводим итоги по перевыполнению
          If Строка_перевыполнение_показателей_по <> "" Then
            Строка_перевыполнение_показателей_по = cityOfficeName(Office2_Name) + " по " + Строка_перевыполнение_показателей_по + ". "
            Строка_перевыполнение_по_офису_показателей = Строка_перевыполнение_по_офису_показателей + Строка_перевыполнение_показателей_по
          End If
          ' Подводим итоги по недовыполнению
          If Строка_недовыполнения_по_офису_показателей_по <> "" Then
            Строка_недовыполнения_по_офису_показателей_по = cityOfficeName(Office2_Name) + " по " + Строка_недовыполнения_по_офису_показателей_по + ". "
            Строка_выполнение_менее_чем_на_половину = Строка_выполнение_менее_чем_на_половину + Строка_недовыполнения_по_офису_показателей_по
          End If
          
        End If
        
      End If ' Если ячейка не пустая
      
  Next

  ' Выводим итоги
  If Строка_перевыполнение_по_офису_показателей <> "" Then
    ThisWorkbook.Sheets("Лист1").Range("B58").Value = "Перевыполнение планов в прогнозе: " + Строка_перевыполнение_по_офису_показателей
    Else: ThisWorkbook.Sheets("Лист1").Range("B58").Value = "Ни один из офисов не выполняет плановые показатели!"
  End If
  
  ' Строка_выполнение_менее_чем_на_половину
  If Строка_выполнение_менее_чем_на_половину <> "" Then
    ThisWorkbook.Sheets("Лист1").Range("B60").Value = "Прогноз выполнения планов менее 50%: " + Строка_выполнение_менее_чем_на_половину
    Else: ThisWorkbook.Sheets("Лист1").Range("B58").Value = "Все офисы выполняют показатели более чем на 50%"
  End If
  
  
  ' Логирование
  If Логирование_в_текстовые_файлы = True Then
    ' Закрываем файлы
    Close #1
    Close #2
  End If
  
  ' Есть такой лист в Книге?
  End If

End Sub

' Выполнение плана по ПК: 3.1 Потребительские  кредиты
Sub Потребительские_кредиты(In_DBstrName As String, In_Число_Офисов As Byte)

' Логирование в текстовый файл чтения DB
Dim Логирование_в_текстовые_файлы As Boolean
' Факт выполнения показателей
Dim Офис_Выдачи_Месяц_План, Офис_Выдачи_Месяц_Факт, Офис_Выдачи_Месяц_Выполнение_процент, Офис_Месяц_Проникновение_Страховок_процент, Офис_Выдачи_Месяц_Прогноз_процент As Double
Dim Офис_Выдачи_Квартал_План, Офис_Выдачи_Квартал_Факт, Офис_Выдачи_Квартал_Выполнение_процент, Офис_Квартал_Проникновение_Страховок_процент, Офис_Выдачи_Квартал_Прогноз_процент As Double
' Начало и конец блока офисов
Dim startOffice2Row, endOffice2Row As Byte
' Для определения имени Листа с ПК
Dim Текущий_Лист As Worksheet
Dim Лист_Потребительские_кредиты As String
Dim Список_открыли As Boolean

  Application.StatusBar = "Потребительские_кредиты ..."

  ' Логирование в текстовый файл чтения DB: True - логируем, False - не логируем
  Логирование_в_текстовые_файлы = False

  ' Примечание: ранее лист назывался "2.1 Потребительские  кредиты"
  ' Перейти на "3.1 Потребительские  кредиты"
  For Each Текущий_Лист In Worksheets
    If InStr(Текущий_Лист.Name, "Потребительские  кредиты") <> 0 Then Лист_Потребительские_кредиты = Текущий_Лист.Name
  Next
  
  ' Sheets("3.1 Потребительские  кредиты").Select
  Sheets(Лист_Потребительские_кредиты).Select

  ' Логирование в текстовые файлы
  If Логирование_в_текстовые_файлы = True Then
    ' В файл выводим все:
    MyFile1 = In_DBstrName & "_ПК_1_log.txt"
    Open MyFile1 For Output As #1
    ' Второй вариант имени файла - вывод в конкретный каталог
    MyFile2 = In_DBstrName & "_ПК_2_log.txt"
    ' Открыли для записи
    Open MyFile2 For Output As #2
  End If

  ' Инициализация переменных
  ' Счетчик офисов
  CountOffice = 0
  ' Наименование офиса 2-го уровня
  Office2_Name = ""
  ' Начало и конец блока офисов из сводной таблицы
  startOffice2Row = 0
  endOffice2Row = 0
  Офис_Выдачи_Месяц_План = 0
  Офис_Выдачи_Месяц_Факт = 0
  Офис_Выдачи_Месяц_Выполнение_процент = 0
  Офис_Месяц_Проникновение_Страховок_процент = 0
  Офис_Выдачи_Месяц_Прогноз_процент = 0
  Офис_Выдачи_Квартал_План = 0
  Офис_Выдачи_Квартал_Факт = 0
  Офис_Выдачи_Квартал_Выполнение_процент = 0
  Офис_Квартал_Проникновение_Страховок_процент = 0
  Офис_Выдачи_Квартал_Прогноз_процент = 0
  
  ' Максимальный объем выданных кредитов
  Строка_1 = ""
  ' Выполнение плана
  Строка_2 = ""
  ' Проникновение страховок в ПК > 75%
  Строка_3 = ""

  ' Цикл обработки Листа Excel
  For Each Cell In ActiveSheet.UsedRange
  
      ' Если ячейка не пустая
      If Not IsEmpty(Cell) Then

        ' Выводим все данные в строку
        If Логирование_в_текстовые_файлы = True Then
          Print #1, Cell.Address, ":" + Cell.Formula
        End If

        ' Вывод строки и столбца
        ' Номер столбца
        intC = Cell.Column
        ' Номер строки
        intR = Cell.Row
                        
        ' Если это столбец B ("Тюменский ОО1")
        If (Mid(Cell.Address, 1, 3) = "$B$") Then
          ' Если адрес ячейки $B$XXX - значит это Наименование офиса, записываем в переменную (для любого офиса)
          Office2_Name = CStr(Cells(intR, intC).Value)
          ' Если это РОО Тюменский
          If (CStr(Cells(intR, intC).Value) = "Тюменский ОО1") Then
            
            ' Открытие
            ' ActiveSheet.PivotTables("Сводная таблица2").PivotFields("DP3_отчет").PivotItems("Тюменский ОО1").ShowDetail = True
            
            ' Открытие списка
            Список_открыли = False
            If (PivotExist("Сводная таблица2") = True) And (Список_открыли = False) Then
              
              On Error Resume Next
              ActiveSheet.PivotTables("Сводная таблица2").PivotFields("DP3_отчет").PivotItems("Тюменский ОО1").ShowDetail = True
              ActiveSheet.PivotTables("Сводная таблица4").PivotFields("DP3_отчет").PivotItems("Тюменский ОО1").ShowDetail = True
              
              Список_открыли = True
            End If
            ' В старых версиях таблица называлась "СводнаяТаблица2" (без пробела)
            If (PivotExist("СводнаяТаблица2") = True) And (Список_открыли = False) Then
              ActiveSheet.PivotTables("СводнаяТаблица2").PivotFields("DP3_отчет").PivotItems("Тюменский ОО1").ShowDetail = True
              Список_открыли = True
            End If


            ' Начало и конец блока офисов
            startOffice2Row = intR + 1
            endOffice2Row = intR + In_Число_Офисов
          End If
          
          
        End If
        
        ' Выводим в Мою книгу Порядковый номер и наименование офиса
        If (Mid(Cell.Address, 1, 3) = "$B$") And (intR >= startOffice2Row) And (intR <= endOffice2Row) Then
          ' Счетчик офисов
          CountOffice = CountOffice + 1
          ThisWorkbook.Sheets("Лист1").Range("A" + CStr(70 + CountOffice)).Value = CountOffice
          ThisWorkbook.Sheets("Лист1").Range("B" + CStr(70 + CountOffice)).Value = Office2_Name
                  
        End If

        ' 1 показатель: Если адрес ячейки $BG$ - значит это Офис_Выдачи_Месяц_План
        If (Mid(Cell.Address, 1, 4) = "$BG$") And (intR >= startOffice2Row) And (intR <= endOffice2Row) Then
          ' Устанавливаем формат ячейки - процентный "0.00%"
          ' ThisWorkbook.Sheets("Лист1").Range("C" + CStr(9 + CountMRK)).NumberFormat = "0%"
          ' Копируем формат ячейки
          Cells(intR, intC).Copy Destination:=ThisWorkbook.Sheets("Лист1").Range("C" + CStr(70 + CountOffice))
          ' Берем Офис_Выдачи_Месяц_План
          Офис_Выдачи_Месяц_План = Round((Cells(intR, intC).Value), 2)
          ' Заносим данные в мою книгу
          ThisWorkbook.Sheets("Лист1").Range("C" + CStr(70 + CountOffice)).Value = Офис_Выдачи_Месяц_План
          ' Убираем рамки
          Call Убираем_рамки_в_ячейке("Лист1", "C" + CStr(70 + CountOffice))
        End If
            
        ' 2 показатель: Если адрес ячейки $BH$ - значит это Офис_Выдачи_Месяц_Факт
        If (Mid(Cell.Address, 1, 4) = "$BH$") And (intR >= startOffice2Row) And (intR <= endOffice2Row) Then
          ' Устанавливаем формат ячейки - процентный "0.00%"
          ' ThisWorkbook.Sheets("Лист1").Range("C" + CStr(9 + CountMRK)).NumberFormat = "0%"
          ' Копируем формат ячейки
          Cells(intR, intC).Copy Destination:=ThisWorkbook.Sheets("Лист1").Range("D" + CStr(70 + CountOffice))
          ' Берем Офис_Выдачи_Месяц_Факт
          Офис_Выдачи_Месяц_Факт = Round((Cells(intR, intC).Value), 2)
          ' Заносим данные в мою книгу
          ThisWorkbook.Sheets("Лист1").Range("D" + CStr(70 + CountOffice)).Value = Офис_Выдачи_Месяц_Факт
          ' Убираем рамки
          Call Убираем_рамки_в_ячейке("Лист1", "D" + CStr(70 + CountOffice))
          Call Убираем_условное_форматирование_в_ячейке("Лист1", "D" + CStr(70 + CountOffice))
        End If
            
        ' 3 показатель: Если адрес ячейки $BJ$ - значит это Офис_Выдачи_Месяц_Выполнение_процент
        If (Mid(Cell.Address, 1, 4) = "$BJ$") And (intR >= startOffice2Row) And (intR <= endOffice2Row) Then
          ' Устанавливаем формат ячейки - процентный "0.00%"
          ' ThisWorkbook.Sheets("Лист1").Range("C" + CStr(9 + CountMRK)).NumberFormat = "0%"
          ' Копируем формат ячейки
          Cells(intR, intC).Copy Destination:=ThisWorkbook.Sheets("Лист1").Range("E" + CStr(70 + CountOffice))
          ' Офис_Выдачи_Месяц_Выполнение_процент
          Офис_Выдачи_Месяц_Выполнение_процент = Round((Cells(intR, intC).Value), 2)
          ' Заносим данные в мою книгу
          ThisWorkbook.Sheets("Лист1").Range("E" + CStr(70 + CountOffice)).Value = Офис_Выдачи_Месяц_Выполнение_процент
          ' Убираем рамки
          Call Убираем_рамки_в_ячейке("Лист1", "E" + CStr(70 + CountOffice))
        End If
            
        ' 4 показатель: Если адрес ячейки $CQ$ - значит это Офис_Месяц_Проникновение_Страховок_процент
        If (Mid(Cell.Address, 1, 4) = "$CQ$") And (intR >= startOffice2Row) And (intR <= endOffice2Row) Then
          ' Устанавливаем формат ячейки - процентный "0.00%"
          ' ThisWorkbook.Sheets("Лист1").Range("C" + CStr(9 + CountMRK)).NumberFormat = "0%"
          ' Копируем формат ячейки
          Cells(intR, intC).Copy Destination:=ThisWorkbook.Sheets("Лист1").Range("F" + CStr(70 + CountOffice))
          ' Офис_Месяц_Проникновение_Страховок_процент
          Офис_Месяц_Проникновение_Страховок_процент = Round((Cells(intR, intC).Value), 2)
          ' Заносим данные в мою книгу
          ThisWorkbook.Sheets("Лист1").Range("F" + CStr(70 + CountOffice)).Value = Офис_Месяц_Проникновение_Страховок_процент
          ' Убираем рамки
          Call Убираем_рамки_в_ячейке("Лист1", "F" + CStr(70 + CountOffice))
        End If
            
        ' 5 показатель: Если адрес ячейки $BK$ - значит это Офис_Выдачи_Месяц_Прогноз_процент
        If (Mid(Cell.Address, 1, 4) = "$BK$") And (intR >= startOffice2Row) And (intR <= endOffice2Row) Then
          ' Устанавливаем формат ячейки - процентный "0.00%"
          ' ThisWorkbook.Sheets("Лист1").Range("C" + CStr(9 + CountMRK)).NumberFormat = "0%"
          ' Копируем формат ячейки
          Cells(intR, intC).Copy Destination:=ThisWorkbook.Sheets("Лист1").Range("G" + CStr(70 + CountOffice))
          ' Офис_Выдачи_Месяц_Прогноз_процент
          Офис_Выдачи_Месяц_Прогноз_процент = Round((Cells(intR, intC).Value), 2)
          ' Заносим данные в мою книгу
          ThisWorkbook.Sheets("Лист1").Range("G" + CStr(70 + CountOffice)).Value = Офис_Выдачи_Месяц_Прогноз_процент
          ' Убираем рамки
          Call Убираем_рамки_в_ячейке("Лист1", "G" + CStr(70 + CountOffice))
        End If
            
        ' == квартал ==
        
        ' 6 (1) показатель: Если адрес ячейки $BX$ - значит это Офис_Выдачи_Квартал_План
        If (Mid(Cell.Address, 1, 4) = "$BX$") And (intR >= startOffice2Row) And (intR <= endOffice2Row) Then
          ' Устанавливаем формат ячейки - процентный "0.00%"
          ' ThisWorkbook.Sheets("Лист1").Range("C" + CStr(9 + CountMRK)).NumberFormat = "0%"
          ' Копируем формат ячейки
          Cells(intR, intC).Copy Destination:=ThisWorkbook.Sheets("Лист1").Range("H" + CStr(70 + CountOffice))
          ' Берем Офис_Выдачи_Квартал_План
          Офис_Выдачи_Квартал_План = Round((Cells(intR, intC).Value), 2)
          ' Заносим данные в мою книгу
          ThisWorkbook.Sheets("Лист1").Range("H" + CStr(70 + CountOffice)).Value = Офис_Выдачи_Квартал_План
          ' Убираем рамки
          Call Убираем_рамки_в_ячейке("Лист1", "H" + CStr(70 + CountOffice))
        End If
            
        ' 7 (2) показатель: Если адрес ячейки $BY$ - значит это Офис_Выдачи_Квартал_Факт
        If (Mid(Cell.Address, 1, 4) = "$BY$") And (intR >= startOffice2Row) And (intR <= endOffice2Row) Then
          ' Устанавливаем формат ячейки - процентный "0.00%"
          ' ThisWorkbook.Sheets("Лист1").Range("C" + CStr(9 + CountMRK)).NumberFormat = "0%"
          ' Копируем формат ячейки
          Cells(intR, intC).Copy Destination:=ThisWorkbook.Sheets("Лист1").Range("I" + CStr(70 + CountOffice))
          ' Берем Офис_Выдачи_Квартал_Факт
          Офис_Выдачи_Квартал_Факт = Round((Cells(intR, intC).Value), 2)
          ' Заносим данные в мою книгу
          ThisWorkbook.Sheets("Лист1").Range("I" + CStr(70 + CountOffice)).Value = Офис_Выдачи_Квартал_Факт
          ' Убираем рамки
          Call Убираем_рамки_в_ячейке("Лист1", "I" + CStr(70 + CountOffice))
        End If
            
        ' 8 (3) показатель: Если адрес ячейки $CA$ - значит это Офис_Выдачи_Квартал_Выполнение_процент
        If (Mid(Cell.Address, 1, 4) = "$CA$") And (intR >= startOffice2Row) And (intR <= endOffice2Row) Then
          ' Устанавливаем формат ячейки - процентный "0.00%"
          ' ThisWorkbook.Sheets("Лист1").Range("C" + CStr(9 + CountMRK)).NumberFormat = "0%"
          ' Копируем формат ячейки
          Cells(intR, intC).Copy Destination:=ThisWorkbook.Sheets("Лист1").Range("J" + CStr(70 + CountOffice))
          ' Берем Офис_Выдачи_Квартал_Выполнение_процент
          Офис_Выдачи_Квартал_Выполнение_процент = Round((Cells(intR, intC).Value), 2)
          ' Заносим данные в мою книгу
          ThisWorkbook.Sheets("Лист1").Range("J" + CStr(70 + CountOffice)).Value = Офис_Выдачи_Квартал_Выполнение_процент
          ' Убираем рамки
          Call Убираем_рамки_в_ячейке("Лист1", "J" + CStr(70 + CountOffice))
        End If
            
        ' 9 (4) показатель: Если адрес ячейки $DK$ - значит это Офис_Квартал_Проникновение_Страховок_процент
        If (Mid(Cell.Address, 1, 4) = "$DK$") And (intR >= startOffice2Row) And (intR <= endOffice2Row) Then
          ' Устанавливаем формат ячейки - процентный "0.00%"
          ' ThisWorkbook.Sheets("Лист1").Range("C" + CStr(9 + CountMRK)).NumberFormat = "0%"
          ' Копируем формат ячейки
          Cells(intR, intC).Copy Destination:=ThisWorkbook.Sheets("Лист1").Range("K" + CStr(70 + CountOffice))
          ' Берем Офис_Квартал_Проникновение_Страховок_процент
          Офис_Квартал_Проникновение_Страховок_процент = Round((Cells(intR, intC).Value), 2)
          ' Заносим данные в мою книгу
          ThisWorkbook.Sheets("Лист1").Range("K" + CStr(70 + CountOffice)).Value = Офис_Квартал_Проникновение_Страховок_процент
          ' Убираем рамки
          Call Убираем_рамки_в_ячейке("Лист1", "K" + CStr(70 + CountOffice))
        End If
            
        ' 10 (5) показатель: Если адрес ячейки $CB$ - значит это Офис_Выдачи_Квартал_Прогноз_процент
        If (Mid(Cell.Address, 1, 4) = "$CB$") And (intR >= startOffice2Row) And (intR <= endOffice2Row) Then
          ' Устанавливаем формат ячейки - процентный "0.00%"
          ' ThisWorkbook.Sheets("Лист1").Range("C" + CStr(9 + CountMRK)).NumberFormat = "0%"
          ' Копируем формат ячейки
          Cells(intR, intC).Copy Destination:=ThisWorkbook.Sheets("Лист1").Range("L" + CStr(70 + CountOffice))
          ' Берем Офис_Выдачи_Квартал_Прогноз_процент
          Офис_Выдачи_Квартал_Прогноз_процент = Round((Cells(intR, intC).Value), 2)
          ' Заносим данные в мою книгу
          ThisWorkbook.Sheets("Лист1").Range("L" + CStr(70 + CountOffice)).Value = Офис_Выдачи_Квартал_Прогноз_процент
          ' Убираем рамки
          Call Убираем_рамки_в_ячейке("Лист1", "L" + CStr(70 + CountOffice))
        End If
                    
      End If ' Если ячейка не пустая
  
  Next
 
  ' Занести План по офисам текущего месяца на Лист3 в Оперативную бизнес-справку (активы)
  ' Значение_между_разделителями
  If False Then ' Пока не обновляем Лист3
    
    rowCount = 71
    ' Здесь идем по столбцу "Date"
    Do While Not IsEmpty(ThisWorkbook.Sheets("Лист1").Cells(rowCount, 2).Value)
      Наименование_офиса = Значение_между_разделителями(ThisWorkbook.Sheets("Лист1").Cells(rowCount, 2).Value, Chr(34), 1, 2)
    
      ' Выполняем поиск Офиса на Листе3 с позиции B6
      rowCount2 = 6
      Офис_найден = False
      Do While (Not IsEmpty(ThisWorkbook.Sheets("Лист3").Cells(rowCount2, 2).Value)) And (Офис_найден = False)
        ' Выполняем поиск подстроки
        If InStr(ThisWorkbook.Sheets("Лист3").Cells(rowCount2, 2).Value, Наименование_офиса) <> 0 Then
          ' Если находим - заносим план на офис
          ThisWorkbook.Sheets("Лист3").Cells(rowCount2, 5).Value = ThisWorkbook.Sheets("Лист1").Cells(rowCount, 3).Value
          ' Форма 2
          ThisWorkbook.Sheets("Лист3").Cells(rowCount2 + 32, 5).Value = ThisWorkbook.Sheets("Лист1").Cells(rowCount, 3).Value
          Офис_найден = True
        End If
        ' Следующая запись
        rowCount2 = rowCount2 + 1
      Loop ' Обработка строк в цикле на Листе3
    
      ' Следующая запись
      rowCount = rowCount + 1
    Loop ' Обработка строк в цикле на Листе1
  
  End If
 
  ' Логирование
  If Логирование_в_текстовые_файлы = True Then
    ' Закрываем файлы
    Close #1
    Close #2
  End If


End Sub

' 4. Выполнение плана по КК: 3.2 Кредитные карты. Число офисов=5
Sub Кредитные_карты(In_DBstrName As String, In_Число_Офисов As Byte)

' Логирование в текстовый файл чтения DB
Dim Логирование_в_текстовые_файлы As Boolean
Dim Выдача_сплитов_в_ПК_Месяц_процент, Активировано_выданых_сплитов_к_ПК_Месяц_процент, Заявки_КК_Месяц_Выполнение_Процент As Double
' Факт выполнения показателей
Dim Кредитные_карты_в_сейфе, Заявки_КК_Месяц_Выполнение As Byte
Dim Активированные_карты_Месяц_План, Активированные_карты_Месяц_Факт, Активированные_карты_Месяц_Выполнение_Процент, Активированные_карты_Месяц_Выполнение_прогноз_Процент As Double
Dim Активированные_карты_Квартал_План, Активированные_карты_Квартал_Факт, Активированные_карты_Квартал_Выполнение_Процент As Double
' Начало и конец блока офисов
Dim startOffice2Row, endOffice2Row As Byte
Dim Лист_Кредитные_карты As String
Dim Список_открыли As Boolean
Dim номерСтрокиЗаголовков, номерСтолбцаАктивированные_карты_Месяц_План, номерСтолбцаАктивированные_карты_Месяц_Факт, номерСтолбцаАктивированные_карты_Месяц_Выполнение_Процент, номерСтолбцаАктивированные_карты_Месяц_Выполнение_прогноз_Процент, номерСтолбцаАктивированные_карты_Квартал_План, номерСтолбцаАктивированные_карты_Квартал_Факт, номерСтолбцаАктивированные_карты_Квартал_Выполнение_Процент, номерСтолбцаАктивированные_карты_Квартал_Выполнение_прогноз_Процент, номерСтолбцаКредитные_карты_в_сейфе, номерСтолбцаВыдача_сплитов_в_ПК_Месяц_процент, номерСтолбцаАктивировано_выданых_сплитов_к_ПК_Месяц_процент, номерСтолбцаЗаявки_КК_Месяц_Выполнение_Процент, номерСтолбцаЗаявки_КК_Месяц_Выполнение As Byte

  Application.StatusBar = "Кредитные_карты ..."

  ' Логирование в текстовый файл чтения DB: True - логируем, False - не логируем
  Логирование_в_текстовые_файлы = False

  ' Перейти на "... Кредитные карты"
  For Each Текущий_Лист In Worksheets
    If InStr(Текущий_Лист.Name, "Кредитные карты") <> 0 Then Лист_Кредитные_карты = Текущий_Лист.Name
  Next
  
  ' Перейти на "3.2 Кредитные карты"
  ' Sheets("3.2 Кредитные карты").Select
  Sheets(Лист_Кредитные_карты).Select

  ' Логирование в текстовые файлы
  If Логирование_в_текстовые_файлы = True Then
    ' В файл выводим все:
    MyFile1 = In_DBstrName & "_КК_1_log.txt"
    Open MyFile1 For Output As #1
    ' Второй вариант имени файла - вывод в конкретный каталог
    MyFile2 = In_DBstrName & "_КК_2_log.txt"
    ' Открыли для записи
    Open MyFile2 For Output As #2
  End If

  ' Инициализация переменных
  ' Счетчик офисов
  CountOffice = 0
  ' Наименование офиса 2-го уровня
  Office2_Name = ""
  ' Начало и конец блока офисов из сводной таблицы
  startOffice2Row = 0
  endOffice2Row = 0
  ' Текущие показатели (месяц)
  Кредитные_карты_в_сейфе = 0
  Выдача_сплитов_в_ПК_Месяц_процент = 0
  Активировано_выданых_сплитов_к_ПК_Месяц_процент = 0
  Заявки_КК_Месяц_Выполнение = 0
  Заявки_КК_Месяц_Выполнение_Процент = 0
  ' Месяц
  Активированные_карты_Месяц_План = 0
  Активированные_карты_Месяц_Факт = 0
  Активированные_карты_Месяц_Выполнение_Процент = 0
  Активированные_карты_Месяц_Выполнение_прогноз_Процент = 0
  ' Квартал
  Активированные_карты_Квартал_План = 0
  Активированные_карты_Квартал_Факт = 0
  Активированные_карты_Квартал_Выполнение_Процент = 0
  
  ' Номер строки с заголовками: Филиал  План            Факт_день   Факт_день-1             Динамика                                  Динамика, %                          % Вып-е                      План                    Факт_день       Факт_день-1         Динамика, шт.        Динамика, %        % Вып-е         План    Факт    Прогноз % Вып-е % Вып-е_Прогноз План    Факт    Прогноз         % Вып-е     % Вып-е_Прог    Продуктивность сотр-ка за мес. (прогноз), шт.   Кол-во выданных ПК со сплитом   Кол-во выданных сплитов Доля выданных сплитов от выданнх ПК со сплитами, %  Кол-во активированных сплитов   Доля актив-х сплитов от выданных, % % Вып-е     % Вып-е _Прогноз
  номерСтрокиЗаголовков = rowByValue(In_DBstrName, "3.2 Кредитные карты", "Филиал", 200, 200)
  ' Номер столбца Активированные_карты_Месяц_План
  номерСтолбцаАктивированные_карты_Месяц_План = ColumnByNameAndNumber(In_DBstrName, "3.2 Кредитные карты", номерСтрокиЗаголовков, "План", 4, 60)
  ' Номер столбца Активированные_карты_Месяц_Факт
  номерСтолбцаАктивированные_карты_Месяц_Факт = ColumnByNameAndNumber(In_DBstrName, "3.2 Кредитные карты", номерСтрокиЗаголовков, "Факт", 2, 60)
  ' Активированные_карты_Месяц_Выполнение_Процент
  номерСтолбцаАктивированные_карты_Месяц_Выполнение_Процент = ColumnByNameAndNumber(In_DBstrName, "3.2 Кредитные карты", номерСтрокиЗаголовков, "% Вып-е", 4, 60)
  ' Активированные_карты_Месяц_Выполнение_прогноз_Процент
  номерСтолбцаАктивированные_карты_Месяц_Выполнение_прогноз_Процент = ColumnByNameAndNumber(In_DBstrName, "3.2 Кредитные карты", номерСтрокиЗаголовков, "% Вып-е_Прог", 1, 60)
  ' Активированные_карты_Квартал_План
  номерСтолбцаАктивированные_карты_Квартал_План = ColumnByNameAndNumber(In_DBstrName, "3.2 Кредитные карты", номерСтрокиЗаголовков, "План", 6, 60)
  ' Активированные_карты_Квартал_Факт
  номерСтолбцаАктивированные_карты_Квартал_Факт = ColumnByNameAndNumber(In_DBstrName, "3.2 Кредитные карты", номерСтрокиЗаголовков, "Факт", 4, 60)
  ' Активированные_карты_Квартал_Выполнение_Процент
  номерСтолбцаАктивированные_карты_Квартал_Выполнение_Процент = ColumnByNameAndNumber(In_DBstrName, "3.2 Кредитные карты", номерСтрокиЗаголовков, "% Вып-е", 6, 60)
  ' Активированные_карты_Квартал_Выполнение_прогноз_Процент
  номерСтолбцаАктивированные_карты_Квартал_Выполнение_прогноз_Процент = ColumnByNameAndNumber(In_DBstrName, "3.2 Кредитные карты", номерСтрокиЗаголовков, "% Вып-е_Прог", 3, 60)
  ' Кредитные_карты_в_сейфе
  номерСтолбцаКредитные_карты_в_сейфе = ColumnByNameAndNumber(In_DBstrName, "3.2 Кредитные карты", номерСтрокиЗаголовков, "Кол-во карт в сейфе", 1, 60)
  ' Выдача_сплитов_в_ПК_Месяц_процент
  номерСтолбцаВыдача_сплитов_в_ПК_Месяц_процент = ColumnByNameAndNumber(In_DBstrName, "3.2 Кредитные карты", номерСтрокиЗаголовков, "Доля выданных сплитов от выданнх ПК со сплитами, %", 1, 60)
  ' Активировано_выданых_сплитов_к_ПК_Месяц_процент
  номерСтолбцаАктивировано_выданых_сплитов_к_ПК_Месяц_процент = ColumnByNameAndNumber(In_DBstrName, "3.2 Кредитные карты", номерСтрокиЗаголовков, "Доля актив-х сплитов от выданных, %", 1, 60)
  ' Заявки_КК_Месяц_Выполнение_Процент
  номерСтолбцаЗаявки_КК_Месяц_Выполнение_Процент = ColumnByNameAndNumber(In_DBstrName, "3.2 Кредитные карты", номерСтрокиЗаголовков, "% Вып-е", 3, 60)
  ' Заявки_КК_Месяц_Выполнение
  номерСтолбцаЗаявки_КК_Месяц_Выполнение = ColumnByNameAndNumber(In_DBstrName, "3.2 Кредитные карты", номерСтрокиЗаголовков, "Факт", 1, 60)
  
  ' Цикл обработки Листа Excel
  For Each Cell In ActiveSheet.UsedRange
  
      ' Если ячейка не пустая
      If Not IsEmpty(Cell) Then

        ' Выводим все данные в строку
        If Логирование_в_текстовые_файлы = True Then
          Print #1, Cell.Address, ":" + Cell.Formula
        End If

        ' Вывод строки и столбца
        ' Номер столбца
        intC = Cell.Column
        ' Номер строки
        intR = Cell.Row
                        
        ' Если это столбец B ("Тюменский ОО1")
        If (Mid(Cell.Address, 1, 3) = "$B$") Then
          ' Если адрес ячейки $B$XXX - значит это Наименование офиса, записываем в переменную (для любого офиса)
          Office2_Name = CStr(Cells(intR, intC).Value)
          ' Если это РОО Тюменский
          If (CStr(Cells(intR, intC).Value) = "Тюменский ОО1") Then
            
            ' Открытие
            ' ActiveSheet.PivotTables("Сводная таблица2").PivotFields("DP3_отчет").PivotItems("Тюменский ОО1").ShowDetail = True
            
            ' Останавливаемся на варианте перечисления сводных таблиц при открытии неправильной будет возникать ошибка - работает!
            On Error Resume Next
            ActiveSheet.PivotTables("Сводная таблица2").PivotFields("DP3_отчет").PivotItems("Тюменский ОО1").ShowDetail = True
            ActiveSheet.PivotTables("СводнаяТаблица2").PivotFields("DP3_отчет").PivotItems("Тюменский ОО1").ShowDetail = True
            ' В феврале 2020 г. c Dashboard_new_РБ_11.02.2020.xlsm
            ActiveSheet.PivotTables("Сводная таблица4").PivotFields("DP3_отчет").PivotItems("Тюменский ОО1").ShowDetail = True
            
            
            ' Открытие списка
            ' Список_открыли = False
            ' If (PivotExist("Сводная таблица2") = True) And (Список_открыли = False) Then
            '   ActiveSheet.PivotTables("Сводная таблица2").PivotFields("DP3_отчет").PivotItems("Тюменский ОО1").ShowDetail = True
            '   Список_открыли = True
            ' End If
            ' В старых версиях таблица называлась "СводнаяТаблица2" (без пробела)
            ' If (PivotExist("СводнаяТаблица2") = True) And (Список_открыли = False) Then
            '   ActiveSheet.PivotTables("СводнаяТаблица2").PivotFields("DP3_отчет").PivotItems("Тюменский ОО1").ShowDetail = True
            '   Список_открыли = True
            ' End If
            
            ' Начало и конец блока офисов
            startOffice2Row = intR + 1
            endOffice2Row = intR + In_Число_Офисов
          End If
                    
        End If
        
        ' Выводим в Мою книгу Порядковый номер и наименование офиса
        If (Mid(Cell.Address, 1, 3) = "$B$") And (intR >= startOffice2Row) And (intR <= endOffice2Row) Then
          ' Счетчик офисов
          CountOffice = CountOffice + 1
          ThisWorkbook.Sheets("Лист1").Range("A" + CStr(89 + CountOffice)).Value = CountOffice
          ThisWorkbook.Sheets("Лист1").Range("B" + CStr(89 + CountOffice)).Value = Office2_Name
                  
        End If

        ' 1 показатель: Если адрес ячейки $R$ - значит это Активированные_карты_Месяц_План
        ' If (Mid(Cell.Address, 1, 3) = "$R$") And (intR >= startOffice2Row) And (intR <= endOffice2Row) Then
        ' Если мы находимся в столбце "План" по счету 3
        If (номерСтолбцаАктивированные_карты_Месяц_План = intC) And (intR >= startOffice2Row) And (intR <= endOffice2Row) Then
          ' Устанавливаем формат ячейки - процентный "0.00%"
          ' Копируем формат ячейки
          Cells(intR, intC).Copy Destination:=ThisWorkbook.Sheets("Лист1").Range("H" + CStr(89 + CountOffice))
          ' Активированные_карты_Месяц_План
          Активированные_карты_Месяц_План = Round((Cells(intR, intC).Value), 2)
          ' Заносим данные в мою книгу
          ThisWorkbook.Sheets("Лист1").Range("H" + CStr(89 + CountOffice)).Value = Активированные_карты_Месяц_План
           ' Убираем рамки
          Call Убираем_рамки_в_ячейке("Лист1", "H" + CStr(89 + CountOffice))
          
        End If

        ' 2 показатель: Если адрес ячейки $S$ - значит это Активированные_карты_Месяц_Факт
        ' If (Mid(Cell.Address, 1, 3) = "$S$") And (intR >= startOffice2Row) And (intR <= endOffice2Row) Then
        If (номерСтолбцаАктивированные_карты_Месяц_Факт = intC) And (intR >= startOffice2Row) And (intR <= endOffice2Row) Then
          ' Устанавливаем формат ячейки - процентный "0.00%"
          ' Копируем формат ячейки
          Cells(intR, intC).Copy Destination:=ThisWorkbook.Sheets("Лист1").Range("I" + CStr(89 + CountOffice))
          ' Активированные_карты_Месяц_Факт
          Активированные_карты_Месяц_Факт = Round((Cells(intR, intC).Value), 2)
          ' Заносим данные в мою книгу
          ThisWorkbook.Sheets("Лист1").Range("I" + CStr(89 + CountOffice)).Value = Активированные_карты_Месяц_Факт
           ' Убираем рамки
          Call Убираем_рамки_в_ячейке("Лист1", "I" + CStr(89 + CountOffice))
        End If

        ' 3 показатель: Если адрес ячейки $U$ - значит это Активированные_карты_Месяц_Выполнение_Процент
        ' If (Mid(Cell.Address, 1, 3) = "$U$") And (intR >= startOffice2Row) And (intR <= endOffice2Row) Then
        If (номерСтолбцаАктивированные_карты_Месяц_Выполнение_Процент = intC) And (intR >= startOffice2Row) And (intR <= endOffice2Row) Then
          ' Устанавливаем формат ячейки - процентный "0.00%"
          ' ThisWorkbook.Sheets("Лист1").Range("C" + CStr(9 + CountMRK)).NumberFormat = "0%"
          ' Копируем формат ячейки
          Cells(intR, intC).Copy Destination:=ThisWorkbook.Sheets("Лист1").Range("J" + CStr(89 + CountOffice))
          ' Активированные_карты_Месяц_Выполнение_Процент
          Активированные_карты_Месяц_Выполнение_Процент = Round((Cells(intR, intC).Value), 2)
          ' Заносим данные в мою книгу
          ThisWorkbook.Sheets("Лист1").Range("J" + CStr(89 + CountOffice)).Value = Активированные_карты_Месяц_Выполнение_Процент
           ' Убираем рамки
          Call Убираем_рамки_в_ячейке("Лист1", "J" + CStr(89 + CountOffice))
        End If

        ' 4 показатель: Если адрес ячейки $V$ - значит это Активированные_карты_Месяц_Выполнение_прогноз_Процент
        ' If (Mid(Cell.Address, 1, 3) = "$V$") And (intR >= startOffice2Row) And (intR <= endOffice2Row) Then
        If (номерСтолбцаАктивированные_карты_Месяц_Выполнение_прогноз_Процент = intC) And (intR >= startOffice2Row) And (intR <= endOffice2Row) Then
          ' Устанавливаем формат ячейки - процентный "0.00%"
          ' ThisWorkbook.Sheets("Лист1").Range("C" + CStr(9 + CountMRK)).NumberFormat = "0%"
          ' Копируем формат ячейки
          Cells(intR, intC).Copy Destination:=ThisWorkbook.Sheets("Лист1").Range("K" + CStr(89 + CountOffice))
          ' Активированные_карты_Месяц_Выполнение_прогноз_Процент
          Активированные_карты_Месяц_Выполнение_прогноз_Процент = Round((Cells(intR, intC).Value), 2)
          ' Заносим данные в мою книгу
          ThisWorkbook.Sheets("Лист1").Range("K" + CStr(89 + CountOffice)).Value = Активированные_карты_Месяц_Выполнение_прогноз_Процент
           ' Убираем рамки
          Call Убираем_рамки_в_ячейке("Лист1", "K" + CStr(89 + CountOffice))
        End If

        ' 5 показатель: Если адрес ячейки $AF$ - значит это Активированные_карты_Квартал_План
        ' If (Mid(Cell.Address, 1, 4) = "$AF$") And (intR >= startOffice2Row) And (intR <= endOffice2Row) Then
        If (номерСтолбцаАктивированные_карты_Квартал_План = intC) And (intR >= startOffice2Row) And (intR <= endOffice2Row) Then
          ' Устанавливаем формат ячейки - процентный "0.00%"
          ' ThisWorkbook.Sheets("Лист1").Range("C" + CStr(9 + CountMRK)).NumberFormat = "0%"
          ' Копируем формат ячейки
          Cells(intR, intC).Copy Destination:=ThisWorkbook.Sheets("Лист1").Range("L" + CStr(89 + CountOffice))
          ' Активированные_карты_Квартал_План
          Активированные_карты_Квартал_План = Round((Cells(intR, intC).Value), 2)
          ' Заносим данные в мою книгу
          ThisWorkbook.Sheets("Лист1").Range("L" + CStr(89 + CountOffice)).Value = Активированные_карты_Квартал_План
           ' Убираем рамки
          Call Убираем_рамки_в_ячейке("Лист1", "L" + CStr(89 + CountOffice))
        End If

        ' 6 показатель: Если адрес ячейки $AG$ - значит это Активированные_карты_Квартал_Факт
        ' If (Mid(Cell.Address, 1, 4) = "$AG$") And (intR >= startOffice2Row) And (intR <= endOffice2Row) Then
        If (номерСтолбцаАктивированные_карты_Квартал_Факт = intC) And (intR >= startOffice2Row) And (intR <= endOffice2Row) Then
          ' Устанавливаем формат ячейки - процентный "0.00%"
          ' ThisWorkbook.Sheets("Лист1").Range("C" + CStr(9 + CountMRK)).NumberFormat = "0%"
          ' Копируем формат ячейки
          Cells(intR, intC).Copy Destination:=ThisWorkbook.Sheets("Лист1").Range("M" + CStr(89 + CountOffice))
          ' Активированные_карты_Квартал_Факт
          Активированные_карты_Квартал_Факт = Round((Cells(intR, intC).Value), 2)
          ' Заносим данные в мою книгу
          ThisWorkbook.Sheets("Лист1").Range("M" + CStr(89 + CountOffice)).Value = Активированные_карты_Квартал_Факт
           ' Убираем рамки
          Call Убираем_рамки_в_ячейке("Лист1", "M" + CStr(89 + CountOffice))
        End If

        ' 7 показатель: Если адрес ячейки $AI$ - значит это Активированные_карты_Квартал_Выполнение_Процент
        ' If (Mid(Cell.Address, 1, 4) = "$AI$") And (intR >= startOffice2Row) And (intR <= endOffice2Row) Then
        If (номерСтолбцаАктивированные_карты_Квартал_Выполнение_Процент = intC) And (intR >= startOffice2Row) And (intR <= endOffice2Row) Then
          ' Устанавливаем формат ячейки - процентный "0.00%"
          ' ThisWorkbook.Sheets("Лист1").Range("C" + CStr(9 + CountMRK)).NumberFormat = "0%"
          ' Копируем формат ячейки
          Cells(intR, intC).Copy Destination:=ThisWorkbook.Sheets("Лист1").Range("N" + CStr(89 + CountOffice))
          ' Активированные_карты_Квартал_Выполнение_Процент
          Активированные_карты_Квартал_Выполнение_Процент = Round((Cells(intR, intC).Value), 2)
          ' Заносим данные в мою книгу
          ThisWorkbook.Sheets("Лист1").Range("N" + CStr(89 + CountOffice)).Value = Активированные_карты_Квартал_Выполнение_Процент
           ' Убираем рамки
          Call Убираем_рамки_в_ячейке("Лист1", "N" + CStr(89 + CountOffice))
        End If

        ' 8 показатель: Если адрес ячейки $AJ$ - значит это Активированные_карты_Квартал_Выполнение_прогноз_Процент
        ' If (Mid(Cell.Address, 1, 4) = "$AJ$") And (intR >= startOffice2Row) And (intR <= endOffice2Row) Then
        If (номерСтолбцаАктивированные_карты_Квартал_Выполнение_прогноз_Процент = intC) And (intR >= startOffice2Row) And (intR <= endOffice2Row) Then
          ' Устанавливаем формат ячейки - процентный "0.00%"
          ' ThisWorkbook.Sheets("Лист1").Range("C" + CStr(9 + CountMRK)).NumberFormat = "0%"
          ' Копируем формат ячейки
          Cells(intR, intC).Copy Destination:=ThisWorkbook.Sheets("Лист1").Range("O" + CStr(89 + CountOffice))
          ' Активированные_карты_Квартал_Выполнение_прогноз_Процент
          Активированные_карты_Квартал_Выполнение_прогноз_Процент = Round((Cells(intR, intC).Value), 2)
          ' Заносим данные в мою книгу
          ThisWorkbook.Sheets("Лист1").Range("O" + CStr(89 + CountOffice)).Value = Активированные_карты_Квартал_Выполнение_прогноз_Процент
           ' Убираем рамки
          Call Убираем_рамки_в_ячейке("Лист1", "O" + CStr(89 + CountOffice))
        End If

        ' ===========
        ' 9 показатель: Если адрес ячейки $AQ$ - значит это Кредитные_карты_в_сейфе
        ' If (Mid(Cell.Address, 1, 4) = "$AQ$") And (intR >= startOffice2Row) And (intR <= endOffice2Row) Then
        If (номерСтолбцаКредитные_карты_в_сейфе = intC) And (intR >= startOffice2Row) And (intR <= endOffice2Row) Then
          ' Устанавливаем формат ячейки - процентный "0.00%"
          ' ThisWorkbook.Sheets("Лист1").Range("C" + CStr(9 + CountMRK)).NumberFormat = "0%"
          ' Копируем формат ячейки
          Cells(intR, intC).Copy Destination:=ThisWorkbook.Sheets("Лист1").Range("C" + CStr(89 + CountOffice))
          ' Кредитные_карты_в_сейфе
          Кредитные_карты_в_сейфе = Round((Cells(intR, intC).Value), 2)
          ' Заносим данные в мою книгу
          ThisWorkbook.Sheets("Лист1").Range("C" + CStr(89 + CountOffice)).Value = Кредитные_карты_в_сейфе
           ' Убираем рамки
          Call Убираем_рамки_в_ячейке("Лист1", "C" + CStr(89 + CountOffice))
        End If

        ' 10 показатель: Если адрес ячейки $Z$ - значит это Выдача_сплитов_в_ПК_Месяц_процент
        ' If (Mid(Cell.Address, 1, 3) = "$Z$") And (intR >= startOffice2Row) And (intR <= endOffice2Row) Then
        If (номерСтолбцаВыдача_сплитов_в_ПК_Месяц_процент = intC) And (intR >= startOffice2Row) And (intR <= endOffice2Row) Then
          ' Устанавливаем формат ячейки - процентный "0.00%"
          ' ThisWorkbook.Sheets("Лист1").Range("C" + CStr(9 + CountMRK)).NumberFormat = "0%"
          ' Копируем формат ячейки
          Cells(intR, intC).Copy Destination:=ThisWorkbook.Sheets("Лист1").Range("D" + CStr(89 + CountOffice))
          ' Выдача_сплитов_в_ПК_Месяц_процент
          Выдача_сплитов_в_ПК_Месяц_процент = Round((Cells(intR, intC).Value), 2)
          ' Заносим данные в мою книгу
          ThisWorkbook.Sheets("Лист1").Range("D" + CStr(89 + CountOffice)).Value = Выдача_сплитов_в_ПК_Месяц_процент
           ' Убираем рамки
          Call Убираем_рамки_в_ячейке("Лист1", "D" + CStr(89 + CountOffice))
        End If

        ' 11 показатель: Если адрес ячейки $AB$ - значит это Активировано_выданых_сплитов_к_ПК_Месяц_процент
        ' If (Mid(Cell.Address, 1, 4) = "$AB$") And (intR >= startOffice2Row) And (intR <= endOffice2Row) Then
        If (номерСтолбцаАктивировано_выданых_сплитов_к_ПК_Месяц_процент = intC) And (intR >= startOffice2Row) And (intR <= endOffice2Row) Then
          ' Устанавливаем формат ячейки - процентный "0.00%"
          ' ThisWorkbook.Sheets("Лист1").Range("C" + CStr(9 + CountMRK)).NumberFormat = "0%"
          ' Копируем формат ячейки
          Cells(intR, intC).Copy Destination:=ThisWorkbook.Sheets("Лист1").Range("E" + CStr(89 + CountOffice))
          ' Активировано_выданых_сплитов_к_ПК_Месяц_процент
          Активировано_выданых_сплитов_к_ПК_Месяц_процент = Round((Cells(intR, intC).Value), 2)
          ' Заносим данные в мою книгу
          ThisWorkbook.Sheets("Лист1").Range("E" + CStr(89 + CountOffice)).Value = Активировано_выданых_сплитов_к_ПК_Месяц_процент
           ' Убираем рамки
          Call Убираем_рамки_в_ячейке("Лист1", "E" + CStr(89 + CountOffice))
        End If

        ' 12 показатель: Если адрес ячейки $Q$ - значит это Заявки_КК_Месяц_Выполнение_Процент
        ' If (Mid(Cell.Address, 1, 3) = "$Q$") And (intR >= startOffice2Row) And (intR <= endOffice2Row) Then
        If (номерСтолбцаЗаявки_КК_Месяц_Выполнение_Процент = intC) And (intR >= startOffice2Row) And (intR <= endOffice2Row) Then
          ' Устанавливаем формат ячейки - процентный "0.00%"
          ' ThisWorkbook.Sheets("Лист1").Range("C" + CStr(9 + CountMRK)).NumberFormat = "0%"
          ' Копируем формат ячейки
          Cells(intR, intC).Copy Destination:=ThisWorkbook.Sheets("Лист1").Range("G" + CStr(89 + CountOffice))
          ' Заявки_КК_Месяц_Выполнение_Процент
          Заявки_КК_Месяц_Выполнение_Процент = Round((Cells(intR, intC).Value), 2)
          ' Заносим данные в мою книгу
          ThisWorkbook.Sheets("Лист1").Range("G" + CStr(89 + CountOffice)).Value = Заявки_КК_Месяц_Выполнение_Процент
           ' Убираем рамки
          Call Убираем_рамки_в_ячейке("Лист1", "G" + CStr(89 + CountOffice))
        End If

        ' 13 показатель: Если адрес ячейки $P$ - значит это Заявки_КК_Месяц_Выполнение
        ' If (Mid(Cell.Address, 1, 3) = "$P$") And (intR >= startOffice2Row) And (intR <= endOffice2Row) Then
        If (номерСтолбцаЗаявки_КК_Месяц_Выполнение = intC) And (intR >= startOffice2Row) And (intR <= endOffice2Row) Then
          ' Устанавливаем формат ячейки - процентный "0.00%"
          ' ThisWorkbook.Sheets("Лист1").Range("C" + CStr(9 + CountMRK)).NumberFormat = "0%"
          ' Копируем формат ячейки
          Cells(intR, intC).Copy Destination:=ThisWorkbook.Sheets("Лист1").Range("F" + CStr(89 + CountOffice))
          ' Заявки_КК_Месяц_Выполнение
          Заявки_КК_Месяц_Выполнение = Round((Cells(intR, intC).Value), 2)
          ' Заносим данные в мою книгу
          ThisWorkbook.Sheets("Лист1").Range("F" + CStr(89 + CountOffice)).Value = Заявки_КК_Месяц_Выполнение
           ' Убираем рамки
          Call Убираем_рамки_в_ячейке("Лист1", "F" + CStr(89 + CountOffice))
        End If


      End If ' Если ячейка не пустая
  
  Next

  ' Логирование
  If Логирование_в_текстовые_файлы = True Then
    ' Закрываем файлы
    Close #1
    Close #2
  End If


End Sub

' 5. Выполнение плана по Комдоходу: 2. Ком доход. Число офисов=5
Sub Ком_доход(In_DBstrName As String, In_Число_Офисов As Byte)

' Логирование в текстовый файл чтения DB
Dim Логирование_в_текстовые_файлы As Boolean
' Переменные
Dim Страховки_к_ПК_месяц_План, Страховки_к_ПК_месяц_Факт, Страховки_к_ПК_месяц_Выполнение_Процент, Страховки_к_ПК_месяц_Выполнение_Прогноз_Процент As Double
Dim ИСЖ_месяц_План, ИСЖ_месяц_Факт, ИСЖ_месяц_Выполнение_Процент, ИСЖ_месяц_Выполнение_Прогноз_Процент As Double
Dim НСЖ_месяц_План, НСЖ_месяц_Факт, НСЖ_месяц_Выполнение_Процент, НСЖ_месяц_Выполнение_Прогноз_Процент As Double
' Начало и конец блока офисов
Dim startOffice2Row, endOffice2Row As Byte
Dim Лист_Ком_доход As String
Dim Список_открыли As Boolean

  Application.StatusBar = "Ком_доход ..."

  ' Логирование в текстовый файл чтения DB: True - логируем, False - не логируем
  Логирование_в_текстовые_файлы = False

  ' Перейти на "... Ком доход"
  For Each Текущий_Лист In Worksheets
    If InStr(Текущий_Лист.Name, "Ком доход") <> 0 Then Лист_Ком_доход = Текущий_Лист.Name
  Next

  ' Перейти на "... Ком доход"
  ' Sheets("2. Ком доход").Select
  Sheets(Лист_Ком_доход).Select

  ' Логирование в текстовые файлы
  If Логирование_в_текстовые_файлы = True Then
    ' В файл выводим все:
    MyFile1 = In_DBstrName & "_КД_1_log.txt"
    Open MyFile1 For Output As #1
    ' Второй вариант имени файла - вывод в конкретный каталог
    MyFile2 = In_DBstrName & "_КД_2_log.txt"
    ' Открыли для записи
    Open MyFile2 For Output As #2
  End If

  ' Инициализация переменных
  ' Счетчик офисов
  CountOffice = 0
  ' Наименование офиса 2-го уровня
  Office2_Name = ""
  ' Начало и конец блока офисов из сводной таблицы
  startOffice2Row = 0
  endOffice2Row = 0
  ' Текущие показатели (месяц)
  Страховки_к_ПК_месяц_План = 0
  Страховки_к_ПК_месяц_Факт = 0
  Страховки_к_ПК_месяц_Выполнение_Процент = 0
  Страховки_к_ПК_месяц_Выполнение_Прогноз_Процент = 0
  ИСЖ_месяц_План = 0
  ИСЖ_месяц_Факт = 0
  ИСЖ_месяц_Выполнение_Процент = 0
  ИСЖ_месяц_Выполнение_Прогноз_Процент = 0
  НСЖ_месяц_План = 0
  НСЖ_месяц_Факт = 0
  НСЖ_месяц_Выполнение_Процент = 0
  НСЖ_месяц_Выполнение_Прогноз_Процент = 0
    
  ' Цикл обработки Листа Excel
  For Each Cell In ActiveSheet.UsedRange
  
      ' Если ячейка не пустая
      If Not IsEmpty(Cell) Then

        ' Выводим все данные в строку
        If Логирование_в_текстовые_файлы = True Then
          Print #1, Cell.Address, ":" + Cell.Formula
        End If

        ' Вывод строки и столбца
        ' Номер столбца
        intC = Cell.Column
        ' Номер строки
        intR = Cell.Row
                        
        ' Если это столбец B ("Тюменский ОО1")
        If (Mid(Cell.Address, 1, 3) = "$B$") Then
          ' Если адрес ячейки $B$XXX - значит это Наименование офиса, записываем в переменную (для любого офиса)
          Office2_Name = CStr(Cells(intR, intC).Value)
          ' Если это РОО Тюменский
          If (CStr(Cells(intR, intC).Value) = "Тюменский ОО1") Then
            
            ' Открытие списка
            Список_открыли = False
            If (PivotExist("Сводная таблица2") = True) And (Список_открыли = False) Then
              ActiveSheet.PivotTables("Сводная таблица2").PivotFields("DP3_отчет").PivotItems("Тюменский ОО1").ShowDetail = True
              Список_открыли = True
            End If
            ' В старых версиях таблица называлась "СводнаяТаблица2" (без пробела)
            If (PivotExist("СводнаяТаблица2") = True) And (Список_открыли = False) Then
              ActiveSheet.PivotTables("СводнаяТаблица2").PivotFields("DP3_отчет").PivotItems("Тюменский ОО1").ShowDetail = True
              Список_открыли = True
            End If

            ' Начало и конец блока офисов
            startOffice2Row = intR + 1
            endOffice2Row = intR + In_Число_Офисов
          End If
                    
        End If
        
        ' Выводим в Мою книгу Порядковый номер и наименование офиса
        If (Mid(Cell.Address, 1, 3) = "$B$") And (intR >= startOffice2Row) And (intR <= endOffice2Row) Then
          ' Счетчик офисов
          CountOffice = CountOffice + 1
          ThisWorkbook.Sheets("Лист1").Range("A" + CStr(108 + CountOffice)).Value = CountOffice
          ThisWorkbook.Sheets("Лист1").Range("B" + CStr(108 + CountOffice)).Value = Office2_Name
                  
        End If

        ' 1 показатель: Если адрес ячейки $C$ - значит это Страховки_к_ПК_месяц_План
        If (Mid(Cell.Address, 1, 3) = "$C$") And (intR >= startOffice2Row) And (intR <= endOffice2Row) Then
          ' Устанавливаем формат ячейки - процентный "0.00%"
          ' ThisWorkbook.Sheets("Лист1").Range("C" + CStr(9 + CountMRK)).NumberFormat = "0%"
          ' Копируем формат ячейки
          Cells(intR, intC).Copy Destination:=ThisWorkbook.Sheets("Лист1").Range("C" + CStr(108 + CountOffice))
          ' Страховки_к_ПК_месяц_План
          Страховки_к_ПК_месяц_План = Round((Cells(intR, intC).Value), 2)
          ' Заносим данные в мою книгу
          ThisWorkbook.Sheets("Лист1").Range("C" + CStr(108 + CountOffice)).Value = Страховки_к_ПК_месяц_План
           ' Убираем рамки
          Call Убираем_рамки_в_ячейке("Лист1", "C" + CStr(108 + CountOffice))
          
          ' *** Апдейт для старой версии отчетов, где был офис Солнечный. Если Страховки_к_ПК_месяц_План=0, то прибавляем на единицу endOffice2Row
          If Страховки_к_ПК_месяц_План = 0 Then
            endOffice2Row = endOffice2Row + 1
          End If
          
        End If

        ' 2 показатель: Если адрес ячейки $D$ - значит это Страховки_к_ПК_месяц_Факт
        If (Mid(Cell.Address, 1, 3) = "$D$") And (intR >= startOffice2Row) And (intR <= endOffice2Row) Then
          ' Устанавливаем формат ячейки - процентный "0.00%"
          ' ThisWorkbook.Sheets("Лист1").Range("C" + CStr(9 + CountMRK)).NumberFormat = "0%"
          ' Копируем формат ячейки
          Cells(intR, intC).Copy Destination:=ThisWorkbook.Sheets("Лист1").Range("D" + CStr(108 + CountOffice))
          ' Страховки_к_ПК_месяц_Факт
          Страховки_к_ПК_месяц_Факт = Round((Cells(intR, intC).Value), 2)
          ' Заносим данные в мою книгу
          ThisWorkbook.Sheets("Лист1").Range("D" + CStr(108 + CountOffice)).Value = Страховки_к_ПК_месяц_Факт
           ' Убираем рамки
          Call Убираем_рамки_в_ячейке("Лист1", "D" + CStr(108 + CountOffice))
        End If

        ' 3 показатель: Если адрес ячейки $E$ - значит это Страховки_к_ПК_месяц_Выполнение_Процент
        If (Mid(Cell.Address, 1, 3) = "$E$") And (intR >= startOffice2Row) And (intR <= endOffice2Row) Then
          ' Устанавливаем формат ячейки - процентный "0.00%"
          ' ThisWorkbook.Sheets("Лист1").Range("C" + CStr(9 + CountMRK)).NumberFormat = "0%"
          ' Копируем формат ячейки
          Cells(intR, intC).Copy Destination:=ThisWorkbook.Sheets("Лист1").Range("E" + CStr(108 + CountOffice))
          ' Страховки_к_ПК_месяц_Выполнение_Процент
          Страховки_к_ПК_месяц_Выполнение_Процент = Round((Cells(intR, intC).Value), 2)
          ' Заносим данные в мою книгу
          ThisWorkbook.Sheets("Лист1").Range("E" + CStr(108 + CountOffice)).Value = Страховки_к_ПК_месяц_Выполнение_Процент
           ' Убираем рамки
          Call Убираем_рамки_в_ячейке("Лист1", "E" + CStr(108 + CountOffice))
        End If

        ' 4 показатель: Если адрес ячейки $F$ - значит это Страховки_к_ПК_месяц_Выполнение_Прогноз_Процент
        If (Mid(Cell.Address, 1, 3) = "$F$") And (intR >= startOffice2Row) And (intR <= endOffice2Row) Then
          ' Устанавливаем формат ячейки - процентный "0.00%"
          ' ThisWorkbook.Sheets("Лист1").Range("C" + CStr(9 + CountMRK)).NumberFormat = "0%"
          ' Копируем формат ячейки
          Cells(intR, intC).Copy Destination:=ThisWorkbook.Sheets("Лист1").Range("F" + CStr(108 + CountOffice))
          ' Страховки_к_ПК_месяц_Выполнение_Прогноз_Процент
          Страховки_к_ПК_месяц_Выполнение_Прогноз_Процент = Round((Cells(intR, intC).Value), 2)
          ' Заносим данные в мою книгу
          ThisWorkbook.Sheets("Лист1").Range("F" + CStr(108 + CountOffice)).Value = Страховки_к_ПК_месяц_Выполнение_Прогноз_Процент
           ' Убираем рамки
          Call Убираем_рамки_в_ячейке("Лист1", "F" + CStr(108 + CountOffice))
        End If

        ' 5 показатель: Если адрес ячейки $G$ - значит это ИСЖ_месяц_План
        If (Mid(Cell.Address, 1, 3) = "$G$") And (intR >= startOffice2Row) And (intR <= endOffice2Row) Then
          ' Устанавливаем формат ячейки - процентный "0.00%"
          ' ThisWorkbook.Sheets("Лист1").Range("C" + CStr(9 + CountMRK)).NumberFormat = "0%"
          ' Копируем формат ячейки
          Cells(intR, intC).Copy Destination:=ThisWorkbook.Sheets("Лист1").Range("G" + CStr(108 + CountOffice))
          ' ИСЖ_месяц_План
          ИСЖ_месяц_План = Round((Cells(intR, intC).Value), 2)
          ' Заносим данные в мою книгу
          ThisWorkbook.Sheets("Лист1").Range("G" + CStr(108 + CountOffice)).Value = ИСЖ_месяц_План
           ' Убираем рамки
          Call Убираем_рамки_в_ячейке("Лист1", "G" + CStr(108 + CountOffice))
        End If

        ' 6 показатель: Если адрес ячейки $H$ - значит это ИСЖ_месяц_Факт
        If (Mid(Cell.Address, 1, 3) = "$H$") And (intR >= startOffice2Row) And (intR <= endOffice2Row) Then
          ' Устанавливаем формат ячейки - процентный "0.00%"
          ' ThisWorkbook.Sheets("Лист1").Range("C" + CStr(9 + CountMRK)).NumberFormat = "0%"
          ' Копируем формат ячейки
          Cells(intR, intC).Copy Destination:=ThisWorkbook.Sheets("Лист1").Range("H" + CStr(108 + CountOffice))
          ' ИСЖ_месяц_Факт
          ИСЖ_месяц_Факт = Round((Cells(intR, intC).Value), 2)
          ' Заносим данные в мою книгу
          ThisWorkbook.Sheets("Лист1").Range("H" + CStr(108 + CountOffice)).Value = ИСЖ_месяц_Факт
           ' Убираем рамки
          Call Убираем_рамки_в_ячейке("Лист1", "H" + CStr(108 + CountOffice))
        End If

        ' 7 показатель: Если адрес ячейки $I$ - значит это ИСЖ_месяц_Выполнение_Процент
        If (Mid(Cell.Address, 1, 3) = "$I$") And (intR >= startOffice2Row) And (intR <= endOffice2Row) Then
          ' Устанавливаем формат ячейки - процентный "0.00%"
          ' ThisWorkbook.Sheets("Лист1").Range("C" + CStr(9 + CountMRK)).NumberFormat = "0%"
          ' Копируем формат ячейки
          Cells(intR, intC).Copy Destination:=ThisWorkbook.Sheets("Лист1").Range("I" + CStr(108 + CountOffice))
          ' ИСЖ_месяц_Выполнение_Процент
          ИСЖ_месяц_Выполнение_Процент = Round((Cells(intR, intC).Value), 2)
          ' Заносим данные в мою книгу
          ThisWorkbook.Sheets("Лист1").Range("I" + CStr(108 + CountOffice)).Value = ИСЖ_месяц_Выполнение_Процент
           ' Убираем рамки
          Call Убираем_рамки_в_ячейке("Лист1", "I" + CStr(108 + CountOffice))
        End If

        ' 8 показатель: Если адрес ячейки $J$ - значит это ИСЖ_месяц_Выполнение_Прогноз_Процент
        If (Mid(Cell.Address, 1, 3) = "$J$") And (intR >= startOffice2Row) And (intR <= endOffice2Row) Then
          ' Устанавливаем формат ячейки - процентный "0.00%"
          ' ThisWorkbook.Sheets("Лист1").Range("C" + CStr(9 + CountMRK)).NumberFormat = "0%"
          ' Копируем формат ячейки
          Cells(intR, intC).Copy Destination:=ThisWorkbook.Sheets("Лист1").Range("J" + CStr(108 + CountOffice))
          ' ИСЖ_месяц_Выполнение_Прогноз_Процент
          ИСЖ_месяц_Выполнение_Прогноз_Процент = Round((Cells(intR, intC).Value), 2)
          ' Заносим данные в мою книгу
          ThisWorkbook.Sheets("Лист1").Range("J" + CStr(108 + CountOffice)).Value = ИСЖ_месяц_Выполнение_Прогноз_Процент
           ' Убираем рамки
          Call Убираем_рамки_в_ячейке("Лист1", "J" + CStr(108 + CountOffice))
        End If

        ' 9 показатель: Если адрес ячейки $K$ - значит это НСЖ_месяц_План
        If (Mid(Cell.Address, 1, 3) = "$K$") And (intR >= startOffice2Row) And (intR <= endOffice2Row) Then
          ' Устанавливаем формат ячейки - процентный "0.00%"
          ' ThisWorkbook.Sheets("Лист1").Range("C" + CStr(9 + CountMRK)).NumberFormat = "0%"
          ' Копируем формат ячейки
          Cells(intR, intC).Copy Destination:=ThisWorkbook.Sheets("Лист1").Range("K" + CStr(108 + CountOffice))
          ' НСЖ_месяц_План
          НСЖ_месяц_План = Round((Cells(intR, intC).Value), 2)
          ' Заносим данные в мою книгу
          ThisWorkbook.Sheets("Лист1").Range("K" + CStr(108 + CountOffice)).Value = НСЖ_месяц_План
           ' Убираем рамки
          Call Убираем_рамки_в_ячейке("Лист1", "K" + CStr(108 + CountOffice))
        End If
       
        ' 10 показатель: Если адрес ячейки $L$ - значит это НСЖ_месяц_Факт
        If (Mid(Cell.Address, 1, 3) = "$L$") And (intR >= startOffice2Row) And (intR <= endOffice2Row) Then
          ' Устанавливаем формат ячейки - процентный "0.00%"
          ' ThisWorkbook.Sheets("Лист1").Range("C" + CStr(9 + CountMRK)).NumberFormat = "0%"
          ' Копируем формат ячейки
          Cells(intR, intC).Copy Destination:=ThisWorkbook.Sheets("Лист1").Range("L" + CStr(108 + CountOffice))
          ' НСЖ_месяц_Факт
          НСЖ_месяц_Факт = Round((Cells(intR, intC).Value), 2)
          ' Заносим данные в мою книгу
          ThisWorkbook.Sheets("Лист1").Range("L" + CStr(108 + CountOffice)).Value = НСЖ_месяц_Факт
           ' Убираем рамки
          Call Убираем_рамки_в_ячейке("Лист1", "L" + CStr(108 + CountOffice))
        End If
       
        ' 11 показатель: Если адрес ячейки $M$ - значит это НСЖ_месяц_Выполнение_Процент
        If (Mid(Cell.Address, 1, 3) = "$M$") And (intR >= startOffice2Row) And (intR <= endOffice2Row) Then
          ' Устанавливаем формат ячейки - процентный "0.00%"
          ' ThisWorkbook.Sheets("Лист1").Range("C" + CStr(9 + CountMRK)).NumberFormat = "0%"
          ' Копируем формат ячейки
          Cells(intR, intC).Copy Destination:=ThisWorkbook.Sheets("Лист1").Range("M" + CStr(108 + CountOffice))
          ' НСЖ_месяц_Выполнение_Процент
          НСЖ_месяц_Выполнение_Процент = Round((Cells(intR, intC).Value), 2)
          ' Заносим данные в мою книгу
          ThisWorkbook.Sheets("Лист1").Range("M" + CStr(108 + CountOffice)).Value = НСЖ_месяц_Выполнение_Процент
           ' Убираем рамки
          Call Убираем_рамки_в_ячейке("Лист1", "M" + CStr(108 + CountOffice))
        End If

        ' 12 показатель: Если адрес ячейки $N$ - значит это НСЖ_месяц_Выполнение_Прогноз_Процент
        If (Mid(Cell.Address, 1, 3) = "$N$") And (intR >= startOffice2Row) And (intR <= endOffice2Row) Then
          ' Устанавливаем формат ячейки - процентный "0.00%"
          ' ThisWorkbook.Sheets("Лист1").Range("C" + CStr(9 + CountMRK)).NumberFormat = "0%"
          ' Копируем формат ячейки
          Cells(intR, intC).Copy Destination:=ThisWorkbook.Sheets("Лист1").Range("N" + CStr(108 + CountOffice))
          ' НСЖ_месяц_Выполнение_Прогноз_Процент
          НСЖ_месяц_Выполнение_Прогноз_Процент = Round((Cells(intR, intC).Value), 2)
          ' Заносим данные в мою книгу
          ThisWorkbook.Sheets("Лист1").Range("N" + CStr(108 + CountOffice)).Value = НСЖ_месяц_Выполнение_Прогноз_Процент
           ' Убираем рамки
          Call Убираем_рамки_в_ячейке("Лист1", "N" + CStr(108 + CountOffice))
        End If


      End If ' Если ячейка не пустая
  
  Next

  ' Логирование
  If Логирование_в_текстовые_файлы = True Then
    ' Закрываем файлы
    Close #1
    Close #2
  End If



End Sub

' 6. Выполнение плана по ОРС: 3.10 OPC. Число офисов=5
Sub OPC(In_DBstrName As String, In_Число_Офисов As Byte)

' Логирование в текстовый файл чтения DB
Dim Логирование_в_текстовые_файлы As Boolean
' Переменные
Dim OPC_месяц_План, OPC_месяц_Факт, OPC_месяц_Выполнение_Процент, OPC_месяц_Прогноз_Процент As Double
Dim OPC_квартал_План, OPC_квартал_Факт, OPC_квартал_Выполнение_Процент, OPC_квартал_Прогноз_Процент As Double
' Начало и конец блока офисов
Dim startOffice2Row, endOffice2Row As Byte
Dim Лист_OPC As String

  Application.StatusBar = "OPC ..."

  ' Логирование в текстовый файл чтения DB: True - логируем, False - не логируем
  Логирование_в_текстовые_файлы = False

  ' Перейти на "... OPC"
  For Each Текущий_Лист In Worksheets
    If InStr(Текущий_Лист.Name, "OPC") <> 0 Then Лист_OPC = Текущий_Лист.Name
  Next

  ' Перейти на "3.1 Потребительские  кредиты" Лист_OPC
  ' Sheets("3.10 OPC").Select
  Sheets(Лист_OPC).Select

  ' Логирование в текстовые файлы
  If Логирование_в_текстовые_файлы = True Then
    ' В файл выводим все:
    MyFile1 = In_DBstrName & "_OPC_1_log.txt"
    Open MyFile1 For Output As #1
    ' Второй вариант имени файла - вывод в конкретный каталог
    MyFile2 = In_DBstrName & "_OPC_2_log.txt"
    ' Открыли для записи
    Open MyFile2 For Output As #2
  End If

  ' Инициализация переменных
  ' Счетчик офисов
  CountOffice = 0
  ' Наименование офиса 2-го уровня
  Office2_Name = ""
  ' Начало и конец блока офисов из сводной таблицы
  startOffice2Row = 0
  endOffice2Row = 0
  ' Текущие показатели (месяц)
  OPC_месяц_План = 0
  OPC_месяц_Факт = 0
  OPC_месяц_Выполнение_Процент = 0
  OPC_месяц_Прогноз_Процент = 0
  OPC_квартал_План = 0
  OPC_квартал_Факт = 0
  OPC_квартал_Выполнение_Процент = 0
  OPC_квартал_Прогноз_Процент = 0
    
  ' Цикл обработки Листа Excel
  For Each Cell In ActiveSheet.UsedRange
  
      ' Если ячейка не пустая
      If Not IsEmpty(Cell) Then

        ' Выводим все данные в строку
        If Логирование_в_текстовые_файлы = True Then
          Print #1, Cell.Address, ":" + Cell.Formula
        End If

        ' Вывод строки и столбца
        ' Номер столбца
        intC = Cell.Column
        ' Номер строки
        intR = Cell.Row
                        
        ' Если это столбец B ("Тюменский ОО1")
        If (Mid(Cell.Address, 1, 3) = "$B$") Then
          ' Если адрес ячейки $B$XXX - значит это Наименование офиса, записываем в переменную (для любого офиса)
          Office2_Name = CStr(Cells(intR, intC).Value)
          ' Если это РОО Тюменский
          If (CStr(Cells(intR, intC).Value) = "Тюменский ОО1") Then
            
            ' Открытие
                        
            ' ActiveSheet.PivotTables("СводнаяТаблица26").PivotFields("DP3_отчет").PivotItems("Тюменский ОО1").ShowDetail = True
            ' Для 25.12.2019
            ' ActiveSheet.PivotTables("Сводная таблица2").PivotFields("DP3_отчет").PivotItems("Тюменский ОО1").ShowDetail = True

            ' Открытие списка
            ' Список_открыли = False
            ' If (PivotExist("Сводная таблица2") = True) And (Список_открыли = False) Then
            '   ActiveSheet.PivotTables("Сводная таблица2").PivotFields("DP3_отчет").PivotItems("Тюменский ОО1").ShowDetail = True
            '   Список_открыли = True
            ' End If
            
            ' Открытие "СводнаяТаблица26"
            ' If (PivotExist("СводнаяТаблица26") = True) And (Список_открыли = False) Then
            '   ActiveSheet.PivotTables("СводнаяТаблица26").PivotFields("DP3_отчет").PivotItems("Тюменский ОО1").ShowDetail = True
            '   Список_открыли = True
            ' End If

            ' Этим скриптом выводим все поля Таблицы
            ' rw = 0
            ' For Each pvtField In ActiveSheet.PivotTables("Сводная таблица2").PivotFields
            ' rw = rw + 1
            ' ActiveSheet.Cells(rw, 12).Value = pvtField.Name
            ' Next pvtField

            ' Останавливаемся на варианте перечисления сводных таблиц при открытии неправильной будет возникать ошибка - работает!
            On Error Resume Next
            ActiveSheet.PivotTables("Сводная таблица2").PivotFields("DP3_отчет").PivotItems("Тюменский ОО1").ShowDetail = True
            ActiveSheet.PivotTables("СводнаяТаблица26").PivotFields("DP3_отчет").PivotItems("Тюменский ОО1").ShowDetail = True

            ' Начало и конец блока офисов
            startOffice2Row = intR + 1
            endOffice2Row = intR + In_Число_Офисов
          End If
                    
        End If
        
        ' Выводим в Мою книгу Порядковый номер и наименование офиса
        If (Mid(Cell.Address, 1, 3) = "$B$") And (intR >= startOffice2Row) And (intR <= endOffice2Row) Then
          ' Счетчик офисов
          CountOffice = CountOffice + 1
          ThisWorkbook.Sheets("Лист1").Range("A" + CStr(127 + CountOffice)).Value = CountOffice
          ThisWorkbook.Sheets("Лист1").Range("B" + CStr(127 + CountOffice)).Value = Office2_Name
        End If

        ' 1 показатель: Если адрес ячейки $C$ - значит это OPC_месяц_План
        If (Mid(Cell.Address, 1, 3) = "$C$") And (intR >= startOffice2Row) And (intR <= endOffice2Row) Then
          ' Устанавливаем формат ячейки - процентный "0.00%"
          ' ThisWorkbook.Sheets("Лист1").Range("C" + CStr(9 + CountMRK)).NumberFormat = "0%"
          ' Копируем формат ячейки
          Cells(intR, intC).Copy Destination:=ThisWorkbook.Sheets("Лист1").Range("C" + CStr(127 + CountOffice))
          ' OPC_месяц_План
          OPC_месяц_План = Round((Cells(intR, intC).Value), 2)
          ' Заносим данные в мою книгу
          ThisWorkbook.Sheets("Лист1").Range("C" + CStr(127 + CountOffice)).Value = OPC_месяц_План
           ' Убираем рамки
          Call Убираем_рамки_в_ячейке("Лист1", "C" + CStr(127 + CountOffice))
        End If

        ' 2 показатель: Если адрес ячейки $D$ - значит это OPC_месяц_Факт
        If (Mid(Cell.Address, 1, 3) = "$D$") And (intR >= startOffice2Row) And (intR <= endOffice2Row) Then
          ' Устанавливаем формат ячейки - процентный "0.00%"
          ' ThisWorkbook.Sheets("Лист1").Range("C" + CStr(9 + CountMRK)).NumberFormat = "0%"
          ' Копируем формат ячейки
          Cells(intR, intC).Copy Destination:=ThisWorkbook.Sheets("Лист1").Range("D" + CStr(127 + CountOffice))
          ' OPC_месяц_Факт
          OPC_месяц_Факт = Round((Cells(intR, intC).Value), 2)
          ' Заносим данные в мою книгу
          ThisWorkbook.Sheets("Лист1").Range("D" + CStr(127 + CountOffice)).Value = OPC_месяц_Факт
           ' Убираем рамки
          Call Убираем_рамки_в_ячейке("Лист1", "D" + CStr(127 + CountOffice))
        End If

        ' 3 показатель: Если адрес ячейки $F$ - значит это OPC_месяц_Выполнение_Процент
        If (Mid(Cell.Address, 1, 3) = "$F$") And (intR >= startOffice2Row) And (intR <= endOffice2Row) Then
          ' Устанавливаем формат ячейки - процентный "0.00%"
          ' ThisWorkbook.Sheets("Лист1").Range("C" + CStr(9 + CountMRK)).NumberFormat = "0%"
          ' Копируем формат ячейки
          Cells(intR, intC).Copy Destination:=ThisWorkbook.Sheets("Лист1").Range("E" + CStr(127 + CountOffice))
          ' OPC_месяц_Выполнение_Процент
          OPC_месяц_Выполнение_Процент = Round((Cells(intR, intC).Value), 2)
          ' Заносим данные в мою книгу
          ThisWorkbook.Sheets("Лист1").Range("E" + CStr(127 + CountOffice)).Value = OPC_месяц_Выполнение_Процент
           ' Убираем рамки
          Call Убираем_рамки_в_ячейке("Лист1", "E" + CStr(127 + CountOffice))
        End If

        ' 4 показатель: Если адрес ячейки $G$ - значит это OPC_месяц_Прогноз_Процент
        If (Mid(Cell.Address, 1, 3) = "$G$") And (intR >= startOffice2Row) And (intR <= endOffice2Row) Then
          ' Устанавливаем формат ячейки - процентный "0.00%"
          ' ThisWorkbook.Sheets("Лист1").Range("C" + CStr(9 + CountMRK)).NumberFormat = "0%"
          ' Копируем формат ячейки
          Cells(intR, intC).Copy Destination:=ThisWorkbook.Sheets("Лист1").Range("F" + CStr(127 + CountOffice))
          ' OPC_месяц_Прогноз_Процент
          OPC_месяц_Прогноз_Процент = Round((Cells(intR, intC).Value), 2)
          ' Заносим данные в мою книгу
          ThisWorkbook.Sheets("Лист1").Range("F" + CStr(127 + CountOffice)).Value = OPC_месяц_Прогноз_Процент
           ' Убираем рамки
          Call Убираем_рамки_в_ячейке("Лист1", "F" + CStr(127 + CountOffice))
        End If

        ' 5 показатель: Если адрес ячейки $H$ - значит это OPC_квартал_План
        If (Mid(Cell.Address, 1, 3) = "$H$") And (intR >= startOffice2Row) And (intR <= endOffice2Row) Then
          ' Устанавливаем формат ячейки - процентный "0.00%"
          ' ThisWorkbook.Sheets("Лист1").Range("C" + CStr(9 + CountMRK)).NumberFormat = "0%"
          ' Копируем формат ячейки
          Cells(intR, intC).Copy Destination:=ThisWorkbook.Sheets("Лист1").Range("G" + CStr(127 + CountOffice))
          ' OPC_квартал_План
          OPC_квартал_План = Round((Cells(intR, intC).Value), 2)
          ' Заносим данные в мою книгу
          ThisWorkbook.Sheets("Лист1").Range("G" + CStr(127 + CountOffice)).Value = OPC_квартал_План
           ' Убираем рамки
          Call Убираем_рамки_в_ячейке("Лист1", "G" + CStr(127 + CountOffice))
        End If

        ' 6 показатель: Если адрес ячейки $I$ - значит это OPC_квартал_Факт
        If (Mid(Cell.Address, 1, 3) = "$I$") And (intR >= startOffice2Row) And (intR <= endOffice2Row) Then
          ' Устанавливаем формат ячейки - процентный "0.00%"
          ' ThisWorkbook.Sheets("Лист1").Range("C" + CStr(9 + CountMRK)).NumberFormat = "0%"
          ' Копируем формат ячейки
          Cells(intR, intC).Copy Destination:=ThisWorkbook.Sheets("Лист1").Range("H" + CStr(127 + CountOffice))
          ' OPC_квартал_Факт
          OPC_квартал_Факт = Round((Cells(intR, intC).Value), 2)
          ' Заносим данные в мою книгу
          ThisWorkbook.Sheets("Лист1").Range("H" + CStr(127 + CountOffice)).Value = OPC_квартал_Факт
           ' Убираем рамки
          Call Убираем_рамки_в_ячейке("Лист1", "H" + CStr(127 + CountOffice))
        End If

        ' 7 показатель: Если адрес ячейки $K$ - значит это OPC_квартал_Выполнение_Процент
        If (Mid(Cell.Address, 1, 3) = "$K$") And (intR >= startOffice2Row) And (intR <= endOffice2Row) Then
          ' Устанавливаем формат ячейки - процентный "0.00%"
          ' ThisWorkbook.Sheets("Лист1").Range("C" + CStr(9 + CountMRK)).NumberFormat = "0%"
          ' Копируем формат ячейки
          Cells(intR, intC).Copy Destination:=ThisWorkbook.Sheets("Лист1").Range("I" + CStr(127 + CountOffice))
          ' OPC_квартал_Выполнение_Процент
          OPC_квартал_Выполнение_Процент = Round((Cells(intR, intC).Value), 2)
          ' Заносим данные в мою книгу
          ThisWorkbook.Sheets("Лист1").Range("I" + CStr(127 + CountOffice)).Value = OPC_квартал_Выполнение_Процент
           ' Убираем рамки
          Call Убираем_рамки_в_ячейке("Лист1", "I" + CStr(127 + CountOffice))
        End If

        ' 7 показатель: Если адрес ячейки $L$ - значит это OPC_квартал_Прогноз_Процент
        If (Mid(Cell.Address, 1, 3) = "$L$") And (intR >= startOffice2Row) And (intR <= endOffice2Row) Then
          ' Устанавливаем формат ячейки - процентный "0.00%"
          ' ThisWorkbook.Sheets("Лист1").Range("C" + CStr(9 + CountMRK)).NumberFormat = "0%"
          ' Копируем формат ячейки
          Cells(intR, intC).Copy Destination:=ThisWorkbook.Sheets("Лист1").Range("J" + CStr(127 + CountOffice))
          ' OPC_квартал_Прогноз_Процент
          OPC_квартал_Прогноз_Процент = Round((Cells(intR, intC).Value), 2)
          ' Заносим данные в мою книгу
          ThisWorkbook.Sheets("Лист1").Range("J" + CStr(127 + CountOffice)).Value = OPC_квартал_Прогноз_Процент
           ' Убираем рамки
          Call Убираем_рамки_в_ячейке("Лист1", "J" + CStr(127 + CountOffice))
        End If


      End If ' Если ячейка не пустая
  
  Next

  ' Логирование
  If Логирование_в_текстовые_файлы = True Then
    ' Закрываем файлы
    Close #1
    Close #2
  End If

End Sub

' 7. Выполнение плана по ОРС: 3.15 Ипотека
Sub Ипотека(In_DBstrName As String)

' Логирование в текстовый файл чтения DB
Dim Логирование_в_текстовые_файлы As Boolean
' Переменные
Dim Ипотека_План_тыс_руб, Ипотека_Факт_тыс_руб, Ипотека_выполнение_Процент, Ипотека_Прогноз_Процент, Прирост_доли_рынка As Double
Dim Ипотека_План_шт, Ипотека_Факт_шт As Byte
Dim Отчет_по_состоянию_на As String
' Начало и конец блока офисов
Dim startOffice2Row, endOffice2Row As Byte
Dim Лист_Ипотека As String

  Application.StatusBar = "Ипотека ..."

  ' Логирование в текстовый файл чтения DB: True - логируем, False - не логируем
  Логирование_в_текстовые_файлы = False

  ' Перейти на "... Ипотека"
  For Each Текущий_Лист In Worksheets
    If InStr(Текущий_Лист.Name, "Ипотека") <> 0 Then Лист_Ипотека = Текущий_Лист.Name
  Next

  ' Перейти на "... Ипотека"
  ' Sheets("3.15 Ипотека").Select
  Sheets(Лист_Ипотека).Select

  ' Логирование в текстовые файлы
  If Логирование_в_текстовые_файлы = True Then
    ' В файл выводим все:
    MyFile1 = In_DBstrName & "_Ипотека_1_log.txt"
    Open MyFile1 For Output As #1
    ' Второй вариант имени файла - вывод в конкретный каталог
    MyFile2 = In_DBstrName & "_Ипотека_2_log.txt"
    ' Открыли для записи
    Open MyFile2 For Output As #2
  End If

  ' Инициализация переменных
  ' Счетчик офисов
  CountOffice = 0
  ' Наименование офиса 2-го уровня
  Office2_Name = ""
  ' Начало и конец блока офисов из сводной таблицы
  startOffice2Row = 0
  endOffice2Row = 0
  ' Текущие показатели (месяц)
  Ипотека_План_тыс_руб = 0
  Ипотека_Факт_тыс_руб = 0
  Ипотека_План_шт = 0
  Ипотека_Факт_шт = 0
  Ипотека_выполнение_Процент = 0
  Ипотека_Прогноз_Процент = 0
  Прирост_доли_рынка = 0
    
  ' Пишем, за какой период сформированы данные по ипотеке
  ThisWorkbook.Sheets("Лист1").Range("B144").Value = "7. Ипотека (" + Workbooks(In_DBstrName).Sheets(Лист_Ипотека).Cells(1, 3) + ")"
    
  ' Цикл обработки Листа Excel
  For Each Cell In ActiveSheet.UsedRange
  
      ' Если ячейка не пустая
      If Not IsEmpty(Cell) Then

        ' Выводим все данные в строку
        If Логирование_в_текстовые_файлы = True Then
          Print #1, Cell.Address, ":" + Cell.Formula
        End If

        ' Вывод строки и столбца
        ' Номер столбца
        intC = Cell.Column
        ' Номер строки
        intR = Cell.Row
                        
        ' Если это столбец С ("ОО "Тюменский")
        If (Mid(Cell.Address, 1, 3) = "$C$") Then
          ' Если адрес ячейки $B$XXX - значит это Наименование офиса, записываем в переменную (для любого офиса)
          Office2_Name = CStr(Cells(intR, intC).Value)
          ' Если это РОО Тюменский
          If CStr(Cells(intR, intC).Value) = ("ОО " & Chr(34) & "Тюменский" & Chr(34)) Or CStr(Cells(intR, intC).Value) = "Тюменский ОО1" Then
            ' Начало и конец блока офисов
            startOffice2Row = intR
            endOffice2Row = intR
          End If
                    
        End If
        
        ' Выводим в Мою книгу Порядковый номер и наименование офиса
        If (Mid(Cell.Address, 1, 3) = "$C$") And (intR >= startOffice2Row) And (intR <= endOffice2Row) Then
          ' Счетчик офисов
          CountOffice = CountOffice + 1
          ThisWorkbook.Sheets("Лист1").Range("A147").Value = CountOffice
          ThisWorkbook.Sheets("Лист1").Range("B147").Value = Office2_Name
        End If

        ' 1 показатель: Если адрес ячейки $F$ - значит это Ипотека_План_тыс_руб
        If (Mid(Cell.Address, 1, 3) = "$F$") And (intR >= startOffice2Row) And (intR <= endOffice2Row) Then
          ' Устанавливаем формат ячейки - процентный "0.00%"
          ' ThisWorkbook.Sheets("Лист1").Range("C" + CStr(9 + CountMRK)).NumberFormat = "0%"
          ' Копируем формат ячейки
          Cells(intR, intC).Copy Destination:=ThisWorkbook.Sheets("Лист1").Range("C147")
          ' Ипотека_План_тыс_руб
          Ипотека_План_тыс_руб = Round((Cells(intR, intC).Value), 0)
          ' Заносим данные в мою книгу
          ThisWorkbook.Sheets("Лист1").Range("C147").Value = Ипотека_План_тыс_руб
          
          ' Заносим план месяца в оперативную бизнес-справку
          If False Then
            ThisWorkbook.Sheets("Лист3").Range("E19").Value = Ипотека_План_тыс_руб
            ThisWorkbook.Sheets("Лист3").Range("E51").Value = Ипотека_План_тыс_руб
          End If
           
           ' Убираем рамки
          Call Убираем_рамки_в_ячейке("Лист1", "C147")
        End If

        ' 2 показатель: Если адрес ячейки $G$ - значит это Ипотека_Факт_тыс_руб
        If (Mid(Cell.Address, 1, 3) = "$G$") And (intR >= startOffice2Row) And (intR <= endOffice2Row) Then
          ' Устанавливаем формат ячейки - процентный "0.00%"
          ' ThisWorkbook.Sheets("Лист1").Range("C" + CStr(9 + CountMRK)).NumberFormat = "0%"
          ' Копируем формат ячейки
          Cells(intR, intC).Copy Destination:=ThisWorkbook.Sheets("Лист1").Range("D147")
          ' Ипотека_Факт_тыс_руб
          Ипотека_Факт_тыс_руб = Round((Cells(intR, intC).Value), 0)
          ' Заносим данные в мою книгу
          ThisWorkbook.Sheets("Лист1").Range("D147").Value = Ипотека_Факт_тыс_руб
           ' Убираем рамки
          Call Убираем_рамки_в_ячейке("Лист1", "D147")
          ' Для ячейки устанавливаем формат без знаков после запятой
          ThisWorkbook.Sheets("Лист1").Range("D147").NumberFormat = "#,##0"
        End If

        ' 3 показатель: Если адрес ячейки $K$ - значит это Ипотека_План_шт
        If (Mid(Cell.Address, 1, 3) = "$K$") And (intR >= startOffice2Row) And (intR <= endOffice2Row) Then
          ' Устанавливаем формат ячейки - процентный "0.00%"
          ' ThisWorkbook.Sheets("Лист1").Range("C" + CStr(9 + CountMRK)).NumberFormat = "0%"
          ' Копируем формат ячейки
          Cells(intR, intC).Copy Destination:=ThisWorkbook.Sheets("Лист1").Range("E147")
          ' Ипотека_План_шт
          Ипотека_План_шт = Round((Cells(intR, intC).Value), 0)
          ' Заносим данные в мою книгу
          ThisWorkbook.Sheets("Лист1").Range("E147").Value = Ипотека_План_шт
           ' Убираем рамки
          Call Убираем_рамки_в_ячейке("Лист1", "E147")
        End If

        ' 4 показатель: Если адрес ячейки $L$ - значит это Ипотека_Факт_шт
        If (Mid(Cell.Address, 1, 3) = "$L$") And (intR >= startOffice2Row) And (intR <= endOffice2Row) Then
          ' Устанавливаем формат ячейки - процентный "0.00%"
          ' ThisWorkbook.Sheets("Лист1").Range("C" + CStr(9 + CountMRK)).NumberFormat = "0%"
          ' Копируем формат ячейки
          Cells(intR, intC).Copy Destination:=ThisWorkbook.Sheets("Лист1").Range("F147")
          ' Ипотека_Факт_шт
          Ипотека_Факт_шт = Round((Cells(intR, intC).Value), 0)
          ' Заносим данные в мою книгу
          ThisWorkbook.Sheets("Лист1").Range("F147").Value = Ипотека_Факт_шт
           ' Убираем рамки
          Call Убираем_рамки_в_ячейке("Лист1", "F147")
        End If

        ' 5 показатель: Если адрес ячейки $I$ - значит это Ипотека_выполнение_Процент
        If (Mid(Cell.Address, 1, 3) = "$I$") And (intR >= startOffice2Row) And (intR <= endOffice2Row) Then
          ' Устанавливаем формат ячейки - процентный "0.00%"
          ' ThisWorkbook.Sheets("Лист1").Range("C" + CStr(9 + CountMRK)).NumberFormat = "0%"
          ' Копируем формат ячейки
          Cells(intR, intC).Copy Destination:=ThisWorkbook.Sheets("Лист1").Range("G147")
          ' Ипотека_выполнение_Процент
          Ипотека_выполнение_Процент = Round((Cells(intR, intC).Value), 2)
          ' Заносим данные в мою книгу
          ThisWorkbook.Sheets("Лист1").Range("G147").Value = Ипотека_выполнение_Процент
           ' Убираем рамки
          Call Убираем_рамки_в_ячейке("Лист1", "G147")
        End If

        ' 6 показатель: Если адрес ячейки $N$ - значит это Ипотека_Прогноз_Процент
        If (Mid(Cell.Address, 1, 3) = "$N$") And (intR >= startOffice2Row) And (intR <= endOffice2Row) Then
          ' Устанавливаем формат ячейки - процентный "0.00%"
          ' ThisWorkbook.Sheets("Лист1").Range("C" + CStr(9 + CountMRK)).NumberFormat = "0%"
          ' Копируем формат ячейки
          Cells(intR, intC).Copy Destination:=ThisWorkbook.Sheets("Лист1").Range("H147")
          ' Ипотека_Прогноз_Процент
          Ипотека_Прогноз_Процент = Round((Cells(intR, intC).Value), 2)
          ' Заносим данные в мою книгу
          ThisWorkbook.Sheets("Лист1").Range("H147").Value = Ипотека_Прогноз_Процент
           ' Убираем рамки
          Call Убираем_рамки_в_ячейке("Лист1", "H147")
        End If

        ' 7 показатель: Если адрес ячейки $E$ - значит это Прирост_доли_рынка
        If (Mid(Cell.Address, 1, 3) = "$E$") And (intR >= startOffice2Row) And (intR <= endOffice2Row) Then
          ' Устанавливаем формат ячейки - процентный "0.00%"
          ' ThisWorkbook.Sheets("Лист1").Range("C" + CStr(9 + CountMRK)).NumberFormat = "0%"
          ' Копируем формат ячейки
          Cells(intR, intC).Copy Destination:=ThisWorkbook.Sheets("Лист1").Range("I147")
          ' Прирост_доли_рынка
          Прирост_доли_рынка = Round((Cells(intR, intC).Value), 2)
          ' Заносим данные в мою книгу
          ' ThisWorkbook.Sheets("Лист1").Range("H" + CStr(146 + CountOffice)).Value = Прирост_доли_рынка
           ' Убираем рамки
          Call Убираем_рамки_в_ячейке("Лист1", "I147")
        End If

        ' 8 показатель: Если адрес ячейки $P$ - значит это План_Заявки_шт
        If (Mid(Cell.Address, 1, 3) = "$P$") And (intR >= startOffice2Row) And (intR <= endOffice2Row) Then
          ' Устанавливаем формат ячейки - процентный "0.00%"
          ' ThisWorkbook.Sheets("Лист1").Range("C" + CStr(9 + CountMRK)).NumberFormat = "0%"
          ' Копируем формат ячейки
          Cells(intR, intC).Copy Destination:=ThisWorkbook.Sheets("Лист1").Range("C159")
           ' Убираем рамки
          Call Убираем_рамки_в_ячейке("Лист1", "C159")
          ' Форматируем ячейку
          ThisWorkbook.Sheets("Лист1").Range("C159").NumberFormat = "0"
          ThisWorkbook.Sheets("Лист1").Range("C159").Font.Bold = True

        End If

        ' 8 Отчет по состоянию на ___
        ' If (Mid(cell.Address, 1, 4) = "$C$1") Then
          ' Устанавливаем формат ячейки - процентный "0.00%"
          ' ThisWorkbook.Sheets("Лист1").Range("C" + CStr(9 + CountMRK)).NumberFormat = "0%"
          ' Копируем формат ячейки
        '  Cells(intR, intC).Copy Destination:=ThisWorkbook.Sheets("Лист1").Range(С144)
          ' Прирост_доли_рынка
        '  Отчет_по_состоянию_на = Cells(intR, intC).Value
          ' Заносим данные в мою книгу
        '  ThisWorkbook.Sheets("Лист1").Range(С144).Value = "(" + Отчет_по_состоянию_на + ")"
          ' Убираем рамки
        '  Call Убираем_рамки_в_ячейке("Лист1", "С144")
        ' End If
        
      End If ' Если ячейка не пустая
  
  Next
    
  ' Логирование
  If Логирование_в_текстовые_файлы = True Then
    ' Закрываем файлы
    Close #1
    Close #2
  End If


End Sub

' 8. Ипотека из ML (с 2021)
Sub Ипотека_ML_new(In_MLName)

' Логирование в текстовый файл чтения DB
Dim Логирование_в_текстовые_файлы As Boolean
Dim Колонка_с_месяцем_откуда_берем_данные, Строка_всего_заявок, Строка_Одобрено_вкл_ОУК, Строка_Выдано As String
Dim Выдачи_все_сумма, Выдачи_Военная_ипотека_сумма, Выдачи_Семейная_Ипотека_сумма, Выдачи_Новостройка_сумма, Выдачи_Вторичный_рынок_сумма, Выдачи_прочие_программы_сумма As Double
Dim Выдачи_все_шт, Выдачи_Военная_ипотека_шт, Выдачи_Семейная_Ипотека_шт, Выдачи_Новостройка_шт, Выдачи_Вторичный_рынок_шт, Выдачи_прочие_программы_шт As Integer
Dim Учитано_в_категории_программ As Boolean
Dim Наименование_столбца_CREDIT_PROGRAMM_OTHER, Наименование_столбца_Сумма_выдачи, Наименование_столбца_STATUS_RETAIL, Наименование_столбца_ИНН, Наименование_столбца_Наименование_компании, Наименование_столбца_Филиал, Наименование_столбца_Адрес_предмета_залога, Наименование_столбца_FIO, Наименование_столбца_Тип_партнера, Наименование_столбца_Название_партнера, Наименование_столбца_Продукт, Наименование_столбца_Офис, Наименование_столбца_Адрес_регистрации, Наименование_столбца_Дата_выдачи, Наименование_столбца_Выдано, Наименование_столбца_Доход_по_осн_месту_работы, Наименование_столбца_Номер_кредитного_договора As Integer
Dim Наименование_столбца_Date_iss, Наименование_столбца_Ставка_выдачи, Наименование_столбца_Срок_кредита, Наименование_столбца_Первоначальный_взнос, Наименование_столбца_Стоимость, Наименование_столбца_Сумма_выдачи_для_срвз, Наименование_столбца_CLIENT_ID, Наименование_столбца_Форма_справки_о_доходах, Наименование_столбца_Группа_компаний As Integer
Dim Текущий_ИНН, Текущая_организация_наименование As String
Dim count_Текущий_ИНН, счетчик_строк_в_DB_Result_Лист2 As Integer
Dim Текущий_месяц_год_по_выдачам, Текущий_месяц_по_выдачам, Текущий_год_по_выдачам As String

  ' Если был выбран ML-файл
  Application.StatusBar = "Ипотека_ML ..."
  
  ' Логирование в текстовый файл чтения DB: True - логируем, False - не логируем
  Логирование_в_текстовые_файлы = False

  ' Логирование в текстовые файлы
  If Логирование_в_текстовые_файлы = True Then
    ' В файл выводим все:
    MyFile1 = Dir(In_MLName) & "_Ипотека_ML_1_log.txt"
    Open MyFile1 For Output As #1
    ' Второй вариант имени файла - вывод в конкретный каталог
    MyFile2 = Dir(In_MLName) & "_Ипотека_ML_2_log.txt"
    ' Открыли для записи
    Open MyFile2 For Output As #2
  End If
        
  ' Выводим все данные в строку
  If Логирование_в_текстовые_файлы = True Then
    Print #1, "Обработка ML-файла:"
  End If
    
  ' Открываем таблицу BASE\Mortgage.xlsx
  Workbooks.Open (ThisWorkbook.Path + "\Base\Mortgage.xlsx")
        
  ' Переходим в ML-файл
  Workbooks(In_MLName).Activate
  
  ' Перейти на вкладку "Заявки_Выдачи"
  Sheets("Заявки_Выдачи").Select
  
  ' Устанавливаем фильтры Тюменский РОО и все виды заявок
  ActiveSheet.PivotTables("SASApp:CARDS.ALL_IPOT").PivotFields("Филиал").ClearAllFilters
  ActiveSheet.PivotTables("SASApp:CARDS.ALL_IPOT").PivotFields("Филиал").CurrentPage = "Тюменский ОО1"
  
  ' Находим столбец "Общий итог"
  column_Общий_итог = ColumnByValue(In_MLName, "Заявки_Выдачи", "Общий итог", 100, 100)
  
  ' Находим строку " Всего заявок"
  row_Всего_заявок = rowByValue(In_MLName, "Заявки_Выдачи", " Всего заявок", 100, 100)

  ' Открываем сводную таблицу
  Workbooks(In_MLName).Sheets("Заявки_Выдачи").Cells(row_Всего_заявок, column_Общий_итог).ShowDetail = True

  ' Перейти на открывшуюся вкладку "Лист3"
  Sheets("Лист3").Select
  
  ' Обработка строк в цикле
  rowCount = 2
  ' Выдачи
  Выдачи_все_сумма = 0
  Выдачи_все_шт = 0
  Выдачи_Военная_ипотека_шт = 0
  Выдачи_Военная_ипотека_сумма = 0
  Выдачи_Семейная_Ипотека_шт = 0
  Выдачи_Семейная_Ипотека_сумма = 0
  Выдачи_Новостройка_шт = 0
  Выдачи_Новостройка_сумма = 0
  Выдачи_Вторичный_рынок_шт = 0
  Выдачи_Вторичный_рынок_сумма = 0
  Выдачи_прочие_программы_шт = 0
  Выдачи_прочие_программы_сумма = 0
  ' Ищем нужные столбцы в строке 1
  Наименование_столбца_CREDIT_PROGRAMM_OTHER = 0
  Наименование_столбца_Сумма_выдачи = 0
  Наименование_столбца_STATUS_RETAIL = 0
  Наименование_столбца_ИНН = 0
  Наименование_столбца_Наименование_компании = 0
  Наименование_столбца_Филиал = 0
  Наименование_столбца_Адрес_предмета_залога = 0
  Наименование_столбца_FIO = 0
  Наименование_столбца_Тип_партнера = 0
  Наименование_столбца_Название_партнера = 0
  Наименование_столбца_Продукт = 0
  Наименование_столбца_Офис = 0
  Наименование_столбца_Адрес_регистрации = 0
  Наименование_столбца_Дата_выдачи = 0
  Наименование_столбца_Выдано = 0
  Наименование_столбца_Номер_кредитного_договора = 0
  Наименование_столбца_Date_iss = 0
  Наименование_столбца_Ставка_выдачи = 0
  Наименование_столбца_Срок_кредита = 0
  Наименование_столбца_Первоначальный_взнос = 0
  Наименование_столбца_Стоимость = 0
  Наименование_столбца_Сумма_выдачи_для_срвз = 0
  Наименование_столбца_CLIENT_ID = 0
  Наименование_столбца_Форма_справки_о_доходах = 0
  Наименование_столбца_Группа_компаний = 0

  ' Добавить: Одобренная сумма, Сумма выдачи из ML, потому как если это Реф, то сумма только в Сумма выдачи! (31-01-2020) и плюс в базе с ипотекой поставить человеческую дату выдачи, дату остатка и текущий остаток по кредиту

  ColumnCount = 1
  Do While (IsEmpty(Cells(1, ColumnCount)) = False)
    
    ' Ищем название столбца CREDIT_PROGRAMM_OTHER
    If Cells(1, ColumnCount).Value = "CREDIT_PROGRAMM_OTHER" Then
      Наименование_столбца_CREDIT_PROGRAMM_OTHER = ColumnCount
    End If
    ' Ищем название столбца Сумма_выдачи
    If Cells(1, ColumnCount).Value = "Сумма выдачи" Then
      Наименование_столбца_Сумма_выдачи = ColumnCount
    End If
    ' Ищем название столбца STATUS_RETAIL
    If Cells(1, ColumnCount).Value = "STATUS_RETAIL" Then
      Наименование_столбца_STATUS_RETAIL = ColumnCount
    End If
    ' Ищем название столбца ИНН
    If Cells(1, ColumnCount).Value = "ИНН" Then
      Наименование_столбца_ИНН = ColumnCount
    End If
    ' Ищем название столбца Наименование компании
    If Cells(1, ColumnCount).Value = "Наименование компании" Then
      Наименование_столбца_Наименование_компании = ColumnCount
    End If
    ' Ищем название столбца Филиал
    If Cells(1, ColumnCount).Value = "Филиал" Then
      Наименование_столбца_Филиал = ColumnCount
    End If
    ' Ищем название столбца Адрес_предмета_залога
    If Cells(1, ColumnCount).Value = "Адрес предмета залога" Then
      Наименование_столбца_Адрес_предмета_залога = ColumnCount
    End If
    ' Ищем название столбца FIO
    If Cells(1, ColumnCount).Value = "FIO" Then
      Наименование_столбца_FIO = ColumnCount
    End If
    ' Ищем название столбца Тип_партнера
    If Cells(1, ColumnCount).Value = "Тип партнера" Then
      Наименование_столбца_Тип_партнера = ColumnCount
    End If
    ' Ищем название столбца Название_партнера
    If Cells(1, ColumnCount).Value = "Название партнера" Then
      Наименование_столбца_Название_партнера = ColumnCount
    End If
    ' Ищем название столбца Продукт
    If Cells(1, ColumnCount).Value = "Продукт" Then
      Наименование_столбца_Продукт = ColumnCount
    End If
    ' Ищем название столбца Офис
    If Cells(1, ColumnCount).Value = "Офис" Then
      Наименование_столбца_Офис = ColumnCount
    End If
   ' Ищем название столбца Адрес регистрации
    If Cells(1, ColumnCount).Value = "Адрес регистрации" Then
      Наименование_столбца_Адрес_регистрации = ColumnCount
    End If
    ' Ищем название столбца Дата выдачи
    If Cells(1, ColumnCount).Value = "Дата выдачи" Then
      Наименование_столбца_Дата_выдачи = ColumnCount
    End If
    ' Ищем название столбца Выдано
    If Cells(1, ColumnCount).Value = "Выдано" Then
      Наименование_столбца_Выдано = ColumnCount
    End If
    ' Доход по осн. месту работы
    If Cells(1, ColumnCount).Value = "Доход по осн. месту работы" Then
      Наименование_столбца_Доход_по_осн_месту_работы = ColumnCount
    End If
    ' Ищем Номер договора
    If Cells(1, ColumnCount).Value = "№ кредитного договора" Then
      Наименование_столбца_Номер_кредитного_договора = ColumnCount
    End If
    ' Ищем Date_iss
    If Cells(1, ColumnCount).Value = "Date_iss" Then
      Наименование_столбца_Date_iss = ColumnCount
    End If
    ' Ищем Ставка выдачи
    If Cells(1, ColumnCount).Value = "Ставка выдачи" Then
      Наименование_столбца_Ставка_выдачи = ColumnCount
    End If
    ' Ищем Срок_кредита
    If Cells(1, ColumnCount).Value = "Срок кредита (мес)" Then
      Наименование_столбца_Срок_кредита = ColumnCount
    End If
    ' Ищем Первоначальный_взнос
    If Cells(1, ColumnCount).Value = "Первоначальный взнос" Then
      Наименование_столбца_Первоначальный_взнос = ColumnCount
    End If
    ' Ищем Стоимость
    If Cells(1, ColumnCount).Value = "Стоимость (для ПВ)" Then
      Наименование_столбца_Стоимость = ColumnCount
    End If
    ' Ищем Сумма выдачи_для срвз
    If Cells(1, ColumnCount).Value = "Сумма выдачи_для срвз" Then
      Наименование_столбца_Сумма_выдачи_для_срвз = ColumnCount
    End If
    ' Ищем CLIENT_ID
    If Cells(1, ColumnCount).Value = "CLIENT_ID" Then
      Наименование_столбца_CLIENT_ID = ColumnCount
    End If
    ' Ищем Форма справки о доходах
    If Cells(1, ColumnCount).Value = "Форма справки о доходах" Then
      Наименование_столбца_Форма_справки_о_доходах = ColumnCount
    End If
    ' Ищем Группа_компаний
    If Cells(1, ColumnCount).Value = "Группа компаний" Then
      Наименование_столбца_Группа_компаний = ColumnCount
    End If
    
    ' Следующая запись
    ColumnCount = ColumnCount + 1
  
  Loop ' Обработка строк в цикле
  
  ' Выводим все записи по столбец Филиал = "Тюменский ОО1"
  Do While (Trim(Cells(rowCount, Наименование_столбца_Филиал).Value) = "Тюменский ОО1")
   
    ' Обработка строки если "Кредит выдан (Кредит выдан)"
    ' If Cells(RowCount, Наименование_столбца_STATUS_RETAIL).Value = "Кредит выдан (Кредит выдан)" Then
    
    ' В старых версиях ML-файла нет столбца STATUS_RETAIL, поэтому переходим на "Выдано" = "1"
    If Cells(rowCount, Наименование_столбца_Выдано).Value = "1" Then
    
     ' Строка "12" Текущий_месяц_по_выдачам, Строка "2019" Текущий_год_по_выдачам
     If (Mid(CStr(CDate(Cells(rowCount, Наименование_столбца_Дата_выдачи).Value)), 4, 2) = Текущий_месяц_по_выдачам) And (Mid(CStr(CDate(Cells(rowCount, Наименование_столбца_Дата_выдачи).Value)), 7, 4) = Текущий_год_по_выдачам) Then

      ' Выдачи по программам
      Выдачи_все_сумма = Выдачи_все_сумма + Cells(rowCount, Наименование_столбца_Сумма_выдачи).Value
      Выдачи_все_шт = Выдачи_все_шт + 1
      ' Переменная Учтено
      Учитано_в_категории_программ = False
      ' Военная_ипотека
      If Cells(rowCount, Наименование_столбца_CREDIT_PROGRAMM_OTHER).Value = "ВОЕННАЯ ИПОТЕКА/РЕФИНАНСИРОВАНИЕ ВОЕННОЙ ИПОТЕКИ/ЗАЛОГ ПРИОБРЕТАЕМОЙ НЕДВИЖИМОСТИ/КВАРТИРА/ВТОРИЧНЫЙ РЫНОК" Then
        Выдачи_Военная_ипотека_шт = Выдачи_Военная_ипотека_шт + 1
        Выдачи_Военная_ипотека_сумма = Выдачи_Военная_ипотека_сумма + Cells(rowCount, Наименование_столбца_Сумма_выдачи).Value
        Учитано_в_категории_программ = True
      End If
      ' Выдачи_Семейная_Ипотека
      If (Cells(rowCount, Наименование_столбца_CREDIT_PROGRAMM_OTHER).Value = "ГОС.ПРОГРАММА/СЕМЕЙНАЯ ИПОТЕКА/ЗАЛОГ ПРИОБРЕТАЕМОЙ НЕДВИЖИМОСТИ/КВАРТИРА/ВТОРИЧНЫЙ РЫНОК") Or (Cells(rowCount, "AW").Value = "ГОС.ПРОГРАММА/СЕМЕЙНАЯ ИПОТЕКА/ЗАЛОГ ПРИОБРЕТАЕМОЙ НЕДВИЖИМОСТИ/КВАРТИРА/ПЕРВИЧНЫЙ РЫНОК") Then
        Выдачи_Семейная_Ипотека_шт = Выдачи_Семейная_Ипотека_шт + 1
        Выдачи_Семейная_Ипотека_сумма = Выдачи_Семейная_Ипотека_сумма + Cells(rowCount, Наименование_столбца_Сумма_выдачи).Value
        Учитано_в_категории_программ = True
      End If
      ' Выдачи_Новостройка
      If Cells(rowCount, Наименование_столбца_CREDIT_PROGRAMM_OTHER).Value = "ЗАЛОГ ПРИОБРЕТАЕМОЙ НЕДВИЖИМОСТИ(ПРАВ)/КВАРТИРА/ ПЕРВИЧНЫЙ РЫНОК" Then
        Выдачи_Новостройка_шт = Выдачи_Новостройка_шт + 1
        Выдачи_Новостройка_сумма = Выдачи_Новостройка_сумма + Cells(rowCount, Наименование_столбца_Сумма_выдачи).Value
        Учитано_в_категории_программ = True
      End If
      ' Выдачи_Вторичный_рынок
      If Cells(rowCount, Наименование_столбца_CREDIT_PROGRAMM_OTHER).Value = "ЗАЛОГ ПРИОБРЕТАЕМОЙ НЕДВИЖИМОСТИ(ПРАВ)/КВАРТИРА/ ВТОРИЧНЫЙ РЫНОК" Then
        Выдачи_Вторичный_рынок_шт = Выдачи_Вторичный_рынок_шт + 1
        Выдачи_Вторичный_рынок_сумма = Выдачи_Вторичный_рынок_сумма + Cells(rowCount, Наименование_столбца_Сумма_выдачи).Value
        Учитано_в_категории_программ = True
      End If
      ' Прочие программы
      If Учитано_в_категории_программ = False Then
        Выдачи_прочие_программы_шт = Выдачи_прочие_программы_шт + 1
        Выдачи_прочие_программы_сумма = Выдачи_прочие_программы_сумма + Cells(rowCount, Наименование_столбца_Сумма_выдачи).Value
        Учитано_в_категории_программ = True
      End If

      End If ' Текущий месяц и год
      
    End If ' Кредит выдан
      
    ' Следующая запись
    rowCount = rowCount + 1
  Loop ' Обработка строк в цикле
  
  ' Заносим данные в мою таблицу
  
  ' Выдачи Военная ипотека
  ThisWorkbook.Sheets("Лист1").Range("D148").Value = Выдачи_Военная_ипотека_сумма / 1000
  ThisWorkbook.Sheets("Лист1").Range("F148").Value = Выдачи_Военная_ипотека_шт
  ' Выдачи_Семейная_Ипотека
  ThisWorkbook.Sheets("Лист1").Range("D149").Value = Выдачи_Семейная_Ипотека_сумма / 1000
  ThisWorkbook.Sheets("Лист1").Range("F149").Value = Выдачи_Семейная_Ипотека_шт
  ' Выдачи_Новостройка
  ThisWorkbook.Sheets("Лист1").Range("D150").Value = Выдачи_Новостройка_сумма / 1000
  ThisWorkbook.Sheets("Лист1").Range("F150").Value = Выдачи_Новостройка_шт
  ' Выдачи_Вторичный_рынок
  ThisWorkbook.Sheets("Лист1").Range("D151").Value = Выдачи_Вторичный_рынок_сумма / 1000
  ThisWorkbook.Sheets("Лист1").Range("F151").Value = Выдачи_Вторичный_рынок_шт
  ' Прочие программы
  ThisWorkbook.Sheets("Лист1").Range("D152").Value = Выдачи_прочие_программы_сумма / 1000
  ThisWorkbook.Sheets("Лист1").Range("F152").Value = Выдачи_прочие_программы_шт
  ' Выполнение_сумма_факт
  ThisWorkbook.Sheets("Лист1").Range("D147").Value = Round(Выдачи_все_сумма / 1000, 0)
  ' Для ячейки устанавливаем формат без знаков после запятой
  ThisWorkbook.Sheets("Лист1").Range("D147").NumberFormat = "#,##0"
  ' Выполнение_шт_факт
  ThisWorkbook.Sheets("Лист1").Range("F147").Value = Выдачи_все_шт
  ' Вставляем формулу с % выполнения
  ThisWorkbook.Sheets("Лист1").Range("G147").FormulaR1C1 = "=IF(RC[-3]>0,((RC[-3]*100)/RC[-4])/100,0)"
  ' Проставляем доли и целевые ориентиры
  ' Доля СЖ (строящегося жилья) не менее 50-60%
  ' Формула =ЕСЛИ(D147>0;((Факт*100)/План)/100;0)
  ' ThisWorkbook.Sheets("Лист1").Range("G150").Value = "Доля " + CStr(Round(((ThisWorkbook.Sheets("Лист1").Range("F150").Value * 100) / ThisWorkbook.Sheets("Лист1").Range("F147").Value), 1)) + "%"
  ' ThisWorkbook.Sheets("Лист1").Range("H150").Value = "" ' "(Норматив: 50%)"
  ' Доля по Семейной ипотеке не менее 20%
  ' ThisWorkbook.Sheets("Лист1").Range("G149").Value = "Доля " + CStr(Round(((ThisWorkbook.Sheets("Лист1").Range("F149").Value * 100) / ThisWorkbook.Sheets("Лист1").Range("F147").Value), 1)) + "%"
  ' ThisWorkbook.Sheets("Лист1").Range("H149").Value = "" ' "(Норматив: 20%)"
  
  ' TR по СЖ не менее 70%
  ' TR по ГЖ не менее 35%
       
  ' Цели на 2020
  ' Добрый день!
  ' Целевые ориентиры перед всеми ИЦ стоят:
  ' по  выдачам,
  ' КСП,
  ' доле ВИ/ГИ -40/60,
  ' доле аккредитивов,
  ' доле SRG отчетов,
  ' направления дорожной карты по выполнению поставленного плана.
  ' Также РИЦ ставят себе 3  KPI на месяц:
  ' Валентина поставила задачу на январь:
  ' план по заявкам от партнеров 460 шт.
  ' по  изменению воронки продаж с 15/85 на 35/65 (СЖ/ГЖ)
  ' по увеличению общего показателя  TR с 18% (2019 г.) до 35%, на горизонте квартала - до 50%.
       
  ' --- Конец Анализируем выдачи ---

  ' --- Выборка потенциальных зарплатников ---
        
  ' Сортировка Листа 1 по столбцу ИНН Организации клиента, получившего ипотеку
  
  ' Обновить данные в таблице BASE\Mortgage по найденным в ней договорам из ML
  If param_from_ini(ThisWorkbook.Path + "\DB_Result.ini", "Обновить_по_найденным") = "1" Then
    Обновить_по_найденным = True
  Else
    Обновить_по_найденным = False
  End If
  
  ' Выполняем цикл по всем записям у которых в первом столбце либо 0 либо 1
  Текущий_ИНН = ""
  Текущая_организация_наименование = ""
  count_Текущий_ИНН = 0
  счетчик_строк_в_DB_Result_Лист2 = 4
  ' Счетчик - всего выведено Организаций в мой Лист 2  (счет идет с нуля)
  всего_выведено_организаций = 0
  ' Начинаем со второй записи
  rowCount = 2
  ' Здесь идем по столбцу "Филиал", в котором должно быть "Тюменский ОО1"
  Do While (Cells(rowCount, Наименование_столбца_Филиал) = ("Тюменский ОО1"))
      
    ' Если у записи кредит выдан?
    ' If Cells(RowCount, Наименование_столбца_STATUS_RETAIL).Value = "Кредит выдан (Кредит выдан)" Then В старых версиях ML-файла нет столбца STATUS_RETAIL, поэтому переходим на "Выдано" = "1"
    If Cells(rowCount, Наименование_столбца_Выдано).Value = "1" Then
               
      ' Если текущий ИНН не равен предыдущему, то это начало
      If Текущий_ИНН <> Cells(rowCount, Наименование_столбца_ИНН).Value Then
                
        ' Вставляем в лист запись с наименованием организации и числом кредитов, если счетчик не равен 0
        If count_Текущий_ИНН > 0 Then
          Call Вывод_в_отчет_итогов_по_Организации_2(In_MLName, Текущий_ИНН, count_Текущий_ИНН, счетчик_строк_в_DB_Result_Лист2, Текущая_организация_наименование, False)
          ' Счет идет не с нуля а с номер строки вывода данных в Листе 1
          счетчик_строк_в_DB_Result_Лист2 = счетчик_строк_в_DB_Result_Лист2 + 1
          ' Счетчик - всего выведено Организаций в мой Лист 2  (счет идет с нуля)
          всего_выведено_организаций = всего_выведено_организаций + 1
        End If
        
        ' Обнуляем счетчик
        count_Текущий_ИНН = 0
        
        ' Присваиваем значение ИНН
        Текущий_ИНН = Cells(rowCount, Наименование_столбца_ИНН).Value
        ' Наименование компании
        Текущая_организация_наименование = Cells(rowCount, Наименование_столбца_Наименование_компании).Value
        
        ' Данные по ипотечной заявке
        ' Адрес предмета залога
        ' FIO
        ' Тип партнера
        ' Название партнера
        ' Продукт
        ' Офис (Тюменский или Сургутский)

      End If
      
      ' Считаем выданные кредиты на организации
      count_Текущий_ИНН = count_Текущий_ИНН + 1
      
    End If
    
    ' Следующая запись
    rowCount = rowCount + 1
  Loop ' Обработка строк в цикле
  
  ' Проверяем - если последняя запись была выведена в отчет?
  If count_Текущий_ИНН > 0 Then
    Call Вывод_в_отчет_итогов_по_Организации_2(In_MLName, Текущий_ИНН, count_Текущий_ИНН, счетчик_строк_в_DB_Result_Лист2, Текущая_организация_наименование, False)
    счетчик_строк_в_DB_Result_Лист2 = счетчик_строк_в_DB_Result_Лист2 + 1
    ' Счетчик - всего выведено Организаций в мой Лист 2 (счет идет с нуля)
    всего_выведено_организаций = всего_выведено_организаций + 1
  End If
    
  ' Перейти на мой Лист 2 и отсортировать
  ThisWorkbook.Activate
  ThisWorkbook.Sheets("Лист2").Select
  Range("D4").Select
  ThisWorkbook.Worksheets("Лист2").Sort.SortFields.Clear
  ThisWorkbook.Worksheets("Лист2").Sort.SortFields.Add Key:=Range("D4"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ThisWorkbook.Worksheets("Лист2").Sort
        .SetRange Range("B3:D" + CStr(всего_выведено_организаций))
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

  ' Теперь вносим в отсортированную форму на Листе 2 ФИО клиентов
  If True Then
    
    ' Счетчик строк
    rowCount = 2
  
    ' Здесь идем по столбцу "Филиал", в котором должно быть "Тюменский ОО1"
    Do While (Workbooks(In_MLName).Worksheets("Лист1").Cells(rowCount, Наименование_столбца_Филиал) = "Тюменский ОО1")
      
      ' Если у записи кредит выдан?
      If Workbooks(In_MLName).Worksheets("Лист1").Cells(rowCount, Наименование_столбца_Выдано).Value = "1" Then
          
        ' Присваиваем значение ИНН
        If Workbooks(In_MLName).Worksheets("Лист1").Cells(rowCount, Наименование_столбца_ИНН).Value <> "" Then
          Текущий_ИНН = Trim(Workbooks(In_MLName).Worksheets("Лист1").Cells(rowCount, Наименование_столбца_ИНН).Value)
        Else
          Текущий_ИНН = "Нет ИНН"
        End If
                
        ' Выдано
        Текущий_Выдано = Workbooks(In_MLName).Worksheets("Лист1").Cells(rowCount, Наименование_столбца_Выдано).Value
        ' Дата выдачи - CDate(Cells(RowCount, Наименование_столбца_Дата_выдачи
        Текущий_Дата_выдачи = CDate(Workbooks(In_MLName).Worksheets("Лист1").Cells(rowCount, Наименование_столбца_Дата_выдачи).Value)
        ' Наименование компании
        Текущая_организация_наименование = Workbooks(In_MLName).Worksheets("Лист1").Cells(rowCount, Наименование_столбца_Наименование_компании).Value
        ' Адрес предмета залога
        Текущий_Адрес_предмета_залога = Workbooks(In_MLName).Worksheets("Лист1").Cells(rowCount, Наименование_столбца_Адрес_предмета_залога).Value
        ' FIO
        Текущий_FIO = Workbooks(In_MLName).Worksheets("Лист1").Cells(rowCount, Наименование_столбца_FIO).Value
        ' Тип партнера
        ' Название партнера
        Текущий_Название_партнера = Workbooks(In_MLName).Worksheets("Лист1").Cells(rowCount, Наименование_столбца_Название_партнера).Value
        ' Продукт
        ' Филиал
        Текущий_Филиал = Workbooks(In_MLName).Worksheets("Лист1").Cells(rowCount, Наименование_столбца_Филиал).Value
        ' Офис (Тюменский или Сургутский)
        Текущий_Офис = Workbooks(In_MLName).Worksheets("Лист1").Cells(rowCount, Наименование_столбца_Офис).Value
        ' Доход по осн месту работы
        Текущий_Доход_по_осн_месту_работы = Round(Workbooks(In_MLName).Worksheets("Лист1").Cells(rowCount, Наименование_столбца_Доход_по_осн_месту_работы).Value / 1000, 1)
        ' Номер договора
        Текущий_Номер_кредитного_договора = Workbooks(In_MLName).Worksheets("Лист1").Cells(rowCount, Наименование_столбца_Номер_кредитного_договора).Value
        ' Текущий_Date_iss
        Текущий_Date_iss = Workbooks(In_MLName).Worksheets("Лист1").Cells(rowCount, Наименование_столбца_Date_iss).Value
        ' Текущий_Ставка_выдачи
        Текущий_Ставка_выдачи = Workbooks(In_MLName).Worksheets("Лист1").Cells(rowCount, Наименование_столбца_Ставка_выдачи).Value
        ' Текущий_Срок_кредита
        Текущий_Срок_кредита = Workbooks(In_MLName).Worksheets("Лист1").Cells(rowCount, Наименование_столбца_Срок_кредита).Value
        ' + Текущий_CREDIT_PROGRAMM_OTHER
        Текущий_CREDIT_PROGRAMM_OTHER = Workbooks(In_MLName).Worksheets("Лист1").Cells(rowCount, Наименование_столбца_CREDIT_PROGRAMM_OTHER).Value
        ' Текущий_Первоначальный_взнос
        Текущий_Первоначальный_взнос = Workbooks(In_MLName).Worksheets("Лист1").Cells(rowCount, Наименование_столбца_Первоначальный_взнос).Value
        ' Текущий_Стоимость
        Текущий_Стоимость = Workbooks(In_MLName).Worksheets("Лист1").Cells(rowCount, Наименование_столбца_Стоимость).Value
        ' + Текущий_Адрес_предмета_залога
        ' Текущий_Сумма_выдачи_для_срвз
        Текущий_Сумма_выдачи_для_срвз = Workbooks(In_MLName).Worksheets("Лист1").Cells(rowCount, Наименование_столбца_Сумма_выдачи_для_срвз).Value
        ' Текущий_CLIENT_ID
        Текущий_CLIENT_ID = Workbooks(In_MLName).Worksheets("Лист1").Cells(rowCount, Наименование_столбца_CLIENT_ID).Value
        ' + Текущий_FIO
        ' + Текущий_Доход_по_осн_месту_работы
        ' Текущий_Форма_справки_о_доходах
        Текущий_Форма_справки_о_доходах = Workbooks(In_MLName).Worksheets("Лист1").Cells(rowCount, Наименование_столбца_Форма_справки_о_доходах).Value
        ' + Текущий_Наименование_компании
        ' + Текущий_ИНН
        ' + Текущий_Название_партнера
        ' Текущий_Тип_партнера
        Текущий_Тип_партнера = Workbooks(In_MLName).Worksheets("Лист1").Cells(rowCount, Наименование_столбца_Тип_партнера).Value
        ' Текущий_Группа_компаний
        Текущий_Группа_компаний = Workbooks(In_MLName).Worksheets("Лист1").Cells(rowCount, Наименование_столбца_Группа_компаний).Value
        ' В DB_Result на Лист 2 по столбцу "B" ищем ИНН
        ' Если в текущем ИНН первый нуль, то его удаляем
        If Mid(Текущий_ИНН, 1, 1) = "0" Then
          Текущий_ИНН = Mid(Текущий_ИНН, 2, Len(Текущий_ИНН) - 1)
        End If
        
        ' Выполняем поиск
        Set fcell = Columns("B:B").Find(Текущий_ИНН, LookAt:=xlWhole)
        If Not fcell Is Nothing Then
          
          ' MsgBox "Нашел в строке: " + CStr(fcell.Row)
          
          ' Выделяем данную запись, следующую за найденной Rows("6:6").Select
          Rows(CStr(fcell.Row + 1) + ":" + CStr(fcell.Row + 1)).Select
          
          ' Вставляем новую строку
          Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
          
          ' Вносим данные в новую строку
          ThisWorkbook.Sheets("Лист2").Range("C" + CStr(fcell.Row + 1)).Value = Текущий_FIO
          ' Подтвержденный доход - "Доход по осн. месту работы"
          ' Для сотрудников ПСБ с ИНН 7744000912 обнуляем доход
          If Текущий_ИНН = "7744000912" Then
            Текущий_Доход_по_осн_месту_работы = 0
          End If
          ' Доход_по_осн_месту_работы
          ThisWorkbook.Sheets("Лист2").Range("D" + CStr(fcell.Row + 1)).Value = Текущий_Доход_по_осн_месту_работы
          ' Адрес объекта
          ThisWorkbook.Sheets("Лист2").Range("E" + CStr(fcell.Row + 1)).Value = Значение_между_разделителями(Текущий_Адрес_предмета_залога, ";", 2, 3)
          ' Партнер
          ThisWorkbook.Sheets("Лист2").Range("F" + CStr(fcell.Row + 1)).Value = Текущий_Название_партнера
          ' Дата сделки
          ThisWorkbook.Sheets("Лист2").Range("G" + CStr(fcell.Row + 1)).Value = CStr(Текущий_Дата_выдачи)
          ' Ставка
          ThisWorkbook.Sheets("Лист2").Range("H" + CStr(fcell.Row + 1)).Value = Текущий_Ставка_выдачи
          ' В столбец H вносим данные по наименованию организации, если нет ИНН
          If Текущий_ИНН = "Нет ИНН" Then
            ThisWorkbook.Sheets("Лист2").Range("I" + CStr(fcell.Row + 1)).Value = "Организация: " + Текущая_организация_наименование
          End If
          ' В столбец J вносим Client ID
          ThisWorkbook.Sheets("Лист2").Range("J" + CStr(fcell.Row + 1)).Value = Текущий_CLIENT_ID
                    
        End If
                    
        ' Вставляем запись в Таблицу BASE\Mortgage
        Call Insert_To_Table_Mortgage(Текущий_Номер_кредитного_договора, Текущий_Date_iss, Текущий_Ставка_выдачи, Текущий_Срок_кредита, Текущий_CREDIT_PROGRAMM_OTHER, Текущий_Первоначальный_взнос, Текущий_Стоимость, Текущий_Адрес_предмета_залога, Текущий_Сумма_выдачи_для_срвз, Текущий_CLIENT_ID, Текущий_FIO, Текущий_Доход_по_осн_месту_работы, Текущий_Форма_справки_о_доходах, Текущая_организация_наименование, Текущий_ИНН, Текущий_Название_партнера, Текущий_Тип_партнера, Текущий_Группа_компаний, Текущий_Выдано, Текущий_Филиал, Текущий_Офис, Обновить_по_найденным)
                                                           
      End If
    
      ' Следующая запись
      rowCount = rowCount + 1
    
    Loop ' Обработка строк в цикле

    ' Вторая итерация - группировка клиентов по ИНН
      
    ' Счетчик строк - начинаем с первой записи
    rowCount = 4
    Начало_блока_группировки_записей = 0
    Блок_Начат = False
    ' Пока в "ИНН" и "Наименование организации" не будет пусто
    Do While (ThisWorkbook.Sheets("Лист2").Cells(rowCount, 2).Value <> "") Or (ThisWorkbook.Sheets("Лист2").Cells(rowCount, 3).Value <> "")
      
      ' Начало блока группировки
      If (ThisWorkbook.Sheets("Лист2").Cells(rowCount, 2).Value = "") And (ThisWorkbook.Sheets("Лист2").Cells(rowCount, 3).Value <> "") And (Блок_Начат = False) Then
        ' Начало блока группировки записей
        Начало_блока_группировки_записей = rowCount
        Блок_Начат = True
      End If

      ' Конец блока группировки
      If (ThisWorkbook.Sheets("Лист2").Cells(rowCount, 2).Value <> "") And (ThisWorkbook.Sheets("Лист2").Cells(rowCount, 3).Value <> "") And (Начало_блока_группировки_записей <> 0) Then
        ' Группируем с начального блока до -1
        ThisWorkbook.Sheets("Лист2").Range("C" + CStr(Начало_блока_группировки_записей) + ":C" + CStr(rowCount - 1)).Select
        Selection.Rows.Group
        Блок_Начат = False
      End If
      
      ' Следующая запись
      rowCount = rowCount + 1
    
    Loop ' Обработка строк в цикле
        
    Columns("E:I").Select
    Selection.Columns.Group
                
    ' Закрываем список
    ActiveSheet.Outline.ShowLevels RowLevels:=1, ColumnLevels:=1

    ' Переходим в первую ячейку
    Range("A1").Select
    
  End If
  
  ' Закрываем таблицу BASE\Mortgage.xlsx с сохранением внесенных изменений
  Workbooks("Mortgage.xlsx").Close SaveChanges:=True
  
  ' --- Конец Выборка потенциальных зарплатников ---

  ' Логирование
  If Логирование_в_текстовые_файлы = True Then
    ' Закрываем файлы
    Close #1
    Close #2
  End If

End Sub


' 8. Ипотека из ML (до 2021)
Sub Ипотека_ML(In_MLName)

' Логирование в текстовый файл чтения DB
Dim Логирование_в_текстовые_файлы As Boolean
Dim Колонка_с_месяцем_откуда_берем_данные, Строка_всего_заявок, Строка_Одобрено_вкл_ОУК, Строка_Выдано As String
Dim Выдачи_все_сумма, Выдачи_Военная_ипотека_сумма, Выдачи_Семейная_Ипотека_сумма, Выдачи_Новостройка_сумма, Выдачи_Вторичный_рынок_сумма, Выдачи_прочие_программы_сумма As Double
Dim Выдачи_все_шт, Выдачи_Военная_ипотека_шт, Выдачи_Семейная_Ипотека_шт, Выдачи_Новостройка_шт, Выдачи_Вторичный_рынок_шт, Выдачи_прочие_программы_шт As Integer
Dim Учитано_в_категории_программ As Boolean
Dim Наименование_столбца_CREDIT_PROGRAMM_OTHER, Наименование_столбца_Сумма_выдачи, Наименование_столбца_STATUS_RETAIL, Наименование_столбца_ИНН, Наименование_столбца_Наименование_компании, Наименование_столбца_Филиал, Наименование_столбца_Адрес_предмета_залога, Наименование_столбца_FIO, Наименование_столбца_Тип_партнера, Наименование_столбца_Название_партнера, Наименование_столбца_Продукт, Наименование_столбца_Офис, Наименование_столбца_Адрес_регистрации, Наименование_столбца_Дата_выдачи, Наименование_столбца_Выдано, Наименование_столбца_Доход_по_осн_месту_работы, Наименование_столбца_Номер_кредитного_договора As Integer
Dim Наименование_столбца_Date_iss, Наименование_столбца_Ставка_выдачи, Наименование_столбца_Срок_кредита, Наименование_столбца_Первоначальный_взнос, Наименование_столбца_Стоимость, Наименование_столбца_Сумма_выдачи_для_срвз, Наименование_столбца_CLIENT_ID, Наименование_столбца_Форма_справки_о_доходах, Наименование_столбца_Группа_компаний As Integer
Dim Текущий_ИНН, Текущая_организация_наименование As String
Dim count_Текущий_ИНН, счетчик_строк_в_DB_Result_Лист2 As Integer
Dim Текущий_месяц_год_по_выдачам, Текущий_месяц_по_выдачам, Текущий_год_по_выдачам As String

  ' Если был выбран ML-файл
  Application.StatusBar = "Ипотека_ML ..."
  
  ' Логирование в текстовый файл чтения DB: True - логируем, False - не логируем
  Логирование_в_текстовые_файлы = False

  ' Логирование в текстовые файлы
  If Логирование_в_текстовые_файлы = True Then
    ' В файл выводим все:
    MyFile1 = Dir(In_MLName) & "_Ипотека_ML_1_log.txt"
    Open MyFile1 For Output As #1
    ' Второй вариант имени файла - вывод в конкретный каталог
    MyFile2 = Dir(In_MLName) & "_Ипотека_ML_2_log.txt"
    ' Открыли для записи
    Open MyFile2 For Output As #2
  End If
        
  ' Выводим все данные в строку
  If Логирование_в_текстовые_файлы = True Then
    Print #1, "Обработка ML-файла:"
  End If
    
  ' Открываем таблицу BASE\Mortgage.xlsx
  Workbooks.Open (ThisWorkbook.Path + "\Base\Mortgage.xlsx")
        
  ' Переходим в ML-файл
  Workbooks(In_MLName).Activate
  
  ' Перейти на вкладку "Заявки_Выдачи"
  Sheets("Заявки_Выдачи").Select
  
  ' --- Выбор фильтра ML-файла ---
  ActiveSheet.PivotTables("SASApp:CARDS.ALL_IPOT").PivotFields("Филиал").ClearAllFilters
  ActiveSheet.PivotTables("SASApp:CARDS.ALL_IPOT").PivotFields("Филиал").CurrentPage = "Тюменский ОО1"
     
     
  ' *** Старая версия ML
     
     
  ' 1. Берем данные по Новостройке
  ' Установка фильтра
  ActiveSheet.PivotTables("SASApp:CARDS.ALL_IPOT").PivotFields("Продукт").ClearAllFilters
  ActiveSheet.PivotTables("SASApp:CARDS.ALL_IPOT").PivotFields("Продукт").CurrentPage = "Новостройка"
          
  ' Подготовка переменных - она меняется на каждом продукте
  Колонка_с_месяцем_откуда_берем_данные = ""
  Строка_всего_заявок = ""
  Строка_Одобрено_вкл_ОУК = ""
  Строка_Выдано = ""
  Строка_AR = ""
  
  ' Цикл обработки Листа Excel - определяем столбец текущего месяца из отчета $P$12 :Общий итог
  For Each Cell In ActiveSheet.UsedRange
    
    ' Если ячейка не пустая
    If Not IsEmpty(Cell) Then
      
      ' Вывод строки и столбца
      ' Номер столбца
      intC = Cell.Column
      ' Номер строки
      intR = Cell.Row

      ' Выводим все данные в строку
      If Логирование_в_текстовые_файлы = True Then
        Print #1, Cell.Address, ":" + Cell.Formula
      End If
        
      ' --- Обработка ячейки ML-файла ---
      
      ' Находим столбец откуда брать данные
      If (CStr(Cells(intR, intC).Value) = "Общий итог") And (Колонка_с_месяцем_откуда_берем_данные = "") Then
      
        If Логирование_в_текстовые_файлы = True Then
          Print #1, "Найден Общий итог"
        End If
        ' Берем колонку с месяцем
        Колонка_с_месяцем_откуда_берем_данные = Предидущая_буква(Mid(Cell.Address, 2, 1))
        If Логирование_в_текстовые_файлы = True Then
          Print #1, "Колонка с месяцем: " + Колонка_с_месяцем_откуда_берем_данные
        End If
      End If
            
      ' Находим строку "Всего заявок"
      If Trim(CStr(Cells(intR, intC).Value)) = "Всего заявок" Then
        If Логирование_в_текстовые_файлы = True Then
          Print #1, "Найдена строка Всего заявок"
        End If
        ' Берем строку
        Строка_всего_заявок = Mid(Cell.Address, 4, 2)
        If Логирование_в_текстовые_файлы = True Then
          Print #1, "Строка Всего заявок: " + Строка_всего_заявок
        End If
      End If
      
      ' Находим строку "Одобрено (вкл ОУК)"
      If Trim(CStr(Cells(intR, intC).Value)) = "Одобрено (вкл ОУК)" Then
        If Логирование_в_текстовые_файлы = True Then
          Print #1, "Найдена строка Одобрено (вкл ОУК)"
        End If
        ' Берем строку
        Строка_Одобрено_вкл_ОУК = Mid(Cell.Address, 4, 2)
        If Логирование_в_текстовые_файлы = True Then
          Print #1, "Строка_Одобрено_вкл_ОУК: " + Строка_Одобрено_вкл_ОУК
        End If
      End If
      
      ' Находим строку "Выдано"
      If Trim(CStr(Cells(intR, intC).Value)) = "Выдано" Then
        If Логирование_в_текстовые_файлы = True Then
          Print #1, "Найдена строка Строка_Выдано"
        End If
        ' Берем строку
        Строка_Выдано = Mid(Cell.Address, 4, 2)
        If Логирование_в_текстовые_файлы = True Then
          Print #1, "Строка_Выдано: " + Строка_Выдано
        End If
      End If
            
      ' Находим строку AR
      If Trim(CStr(Cells(intR, intC).Value)) = "AR (качественно), %" Then
        If Логирование_в_текстовые_файлы = True Then
          Print #1, "Найдена строка Строка_AR"
        End If
        ' Берем строку
        Строка_AR = Mid(Cell.Address, 4, 2)
        If Логирование_в_текстовые_файлы = True Then
          Print #1, "Строка_AR: " + Строка_AR
        End If
      End If
                        
      ' --- Конец обработки ячейки ML-файла ---
      
    End If ' Если ячейка не пустая
        
  Next
  
  ' Копирование данных
  Range(Колонка_с_месяцем_откуда_берем_данные + Строка_всего_заявок).Copy Destination:=ThisWorkbook.Sheets("Лист1").Range("D155")
  Range(Колонка_с_месяцем_откуда_берем_данные + Строка_Одобрено_вкл_ОУК).Copy Destination:=ThisWorkbook.Sheets("Лист1").Range("E155")
  Range(Колонка_с_месяцем_откуда_берем_данные + Строка_Выдано).Copy Destination:=ThisWorkbook.Sheets("Лист1").Range("G155")
  ' AR
  Range(Колонка_с_месяцем_откуда_берем_данные + Строка_AR).Copy Destination:=ThisWorkbook.Sheets("Лист1").Range("F155")
  
  ' 2. Берем данные по Вторичный рынок
  ' Установка фильтра
  ActiveSheet.PivotTables("SASApp:CARDS.ALL_IPOT").PivotFields("Продукт").ClearAllFilters
  ActiveSheet.PivotTables("SASApp:CARDS.ALL_IPOT").PivotFields("Продукт").CurrentPage = "Вторичный рынок"
            
  ' Подготовка переменных - она меняется на каждом продукте
  Колонка_с_месяцем_откуда_берем_данные = ""
  Строка_всего_заявок = ""
  Строка_Одобрено_вкл_ОУК = ""
  Строка_Выдано = ""
  Строка_AR = ""
  
  ' Цикл обработки Листа Excel - определяем столбец текущего месяца из отчета $P$12 :Общий итог
  For Each Cell In ActiveSheet.UsedRange
    
    ' Если ячейка не пустая
    If Not IsEmpty(Cell) Then
      
      ' Вывод строки и столбца
      ' Номер столбца
      intC = Cell.Column
      ' Номер строки
      intR = Cell.Row

      ' Выводим все данные в строку
      If Логирование_в_текстовые_файлы = True Then
        Print #1, Cell.Address, ":" + Cell.Formula
      End If
        
      ' --- Обработка ячейки ML-файла ---
      
      ' Находим столбец откуда брать данные
      If (CStr(Cells(intR, intC).Value) = "Общий итог") And (Колонка_с_месяцем_откуда_берем_данные = "") Then
        If Логирование_в_текстовые_файлы = True Then
          Print #1, "Найден Общий итог"
        End If
        ' Берем колонку с месяцем
        Колонка_с_месяцем_откуда_берем_данные = Предидущая_буква(Mid(Cell.Address, 2, 1))
        If Логирование_в_текстовые_файлы = True Then
          Print #1, "Колонка с месяцем: " + Колонка_с_месяцем_откуда_берем_данные
        End If
      End If
      
      ' Находим строку "Всего заявок"
      If Trim(CStr(Cells(intR, intC).Value)) = "Всего заявок" Then
        If Логирование_в_текстовые_файлы = True Then
          Print #1, "Найдена строка Всего заявок"
        End If
        ' Берем строку
        Строка_всего_заявок = Mid(Cell.Address, 4, 2)
        If Логирование_в_текстовые_файлы = True Then
          Print #1, "Строка Всего заявок: " + Строка_всего_заявок
        End If
      End If
      
      ' Находим строку "Одобрено (вкл ОУК)"
      If Trim(CStr(Cells(intR, intC).Value)) = "Одобрено (вкл ОУК)" Then
        If Логирование_в_текстовые_файлы = True Then
          Print #1, "Найдена строка Одобрено (вкл ОУК)"
        End If
        ' Берем строку
        Строка_Одобрено_вкл_ОУК = Mid(Cell.Address, 4, 2)
        If Логирование_в_текстовые_файлы = True Then
          Print #1, "Строка_Одобрено_вкл_ОУК: " + Строка_Одобрено_вкл_ОУК
        End If
      End If
      
      ' Находим строку "Выдано"
      If Trim(CStr(Cells(intR, intC).Value)) = "Выдано" Then
        If Логирование_в_текстовые_файлы = True Then
          Print #1, "Найдена строка Строка_Выдано"
        End If
        ' Берем строку
        Строка_Выдано = Mid(Cell.Address, 4, 2)
        If Логирование_в_текстовые_файлы = True Then
          Print #1, "Строка_Выдано: " + Строка_Выдано
        End If
      End If
      
      ' Находим строку AR
      If Trim(CStr(Cells(intR, intC).Value)) = "AR (качественно), %" Then
        If Логирование_в_текстовые_файлы = True Then
          Print #1, "Найдена строка Строка_AR"
        End If
        ' Берем строку
        Строка_AR = Mid(Cell.Address, 4, 2)
        If Логирование_в_текстовые_файлы = True Then
          Print #1, "Строка_AR: " + Строка_AR
        End If
      End If

      
      ' --- Конец обработки ячейки ML-файла ---
      
    End If ' Если ячейка не пустая
        
  Next
  
  ' Копирование данных
  Range(Колонка_с_месяцем_откуда_берем_данные + Строка_всего_заявок).Copy Destination:=ThisWorkbook.Sheets("Лист1").Range("D156")
  Range(Колонка_с_месяцем_откуда_берем_данные + Строка_Одобрено_вкл_ОУК).Copy Destination:=ThisWorkbook.Sheets("Лист1").Range("E156")
  Range(Колонка_с_месяцем_откуда_берем_данные + Строка_Выдано).Copy Destination:=ThisWorkbook.Sheets("Лист1").Range("G156")
  ' AR
  Range(Колонка_с_месяцем_откуда_берем_данные + Строка_AR).Copy Destination:=ThisWorkbook.Sheets("Лист1").Range("F156")
 
  ' 3. Берем данные по Рефинансирование
  
  ' Установка фильтра
  ActiveSheet.PivotTables("SASApp:CARDS.ALL_IPOT").PivotFields("Продукт").ClearAllFilters
  ActiveSheet.PivotTables("SASApp:CARDS.ALL_IPOT").PivotFields("Продукт").CurrentPage = "Рефинансирование"
  
  ' Подготовка переменных - она меняется на каждом продукте
  Колонка_с_месяцем_откуда_берем_данные = ""
  Строка_всего_заявок = ""
  Строка_Одобрено_вкл_ОУК = ""
  Строка_Выдано = ""
  Строка_AR = ""
  
  ' Цикл обработки Листа Excel - определяем столбец текущего месяца из отчета $P$12 :Общий итог
  For Each Cell In ActiveSheet.UsedRange
    
    ' Если ячейка не пустая
    If Not IsEmpty(Cell) Then
      
      ' Вывод строки и столбца
      ' Номер столбца
      intC = Cell.Column
      ' Номер строки
      intR = Cell.Row

      ' Выводим все данные в строку
      If Логирование_в_текстовые_файлы = True Then
        Print #1, Cell.Address, ":" + Cell.Formula
      End If
        
      ' --- Обработка ячейки ML-файла ---
      
      ' Находим столбец откуда брать данные
      If (CStr(Cells(intR, intC).Value) = "Общий итог") And (Колонка_с_месяцем_откуда_берем_данные = "") Then
        If Логирование_в_текстовые_файлы = True Then
          Print #1, "Найден Общий итог"
        End If
        ' Берем колонку с месяцем
        Колонка_с_месяцем_откуда_берем_данные = Предидущая_буква(Mid(Cell.Address, 2, 1))
        If Логирование_в_текстовые_файлы = True Then
          Print #1, "Колонка с месяцем: " + Колонка_с_месяцем_откуда_берем_данные
        End If
      End If
      
      ' Находим строку "Всего заявок"
      If Trim(CStr(Cells(intR, intC).Value)) = "Всего заявок" Then
        If Логирование_в_текстовые_файлы = True Then
          Print #1, "Найдена строка Всего заявок"
        End If
        ' Берем строку
        Строка_всего_заявок = Mid(Cell.Address, 4, 2)
        If Логирование_в_текстовые_файлы = True Then
          Print #1, "Строка Всего заявок: " + Строка_всего_заявок
        End If
      End If
      
      ' Находим строку "Одобрено (вкл ОУК)"
      If Trim(CStr(Cells(intR, intC).Value)) = "Одобрено (вкл ОУК)" Then
        If Логирование_в_текстовые_файлы = True Then
          Print #1, "Найдена строка Одобрено (вкл ОУК)"
        End If
        ' Берем строку
        Строка_Одобрено_вкл_ОУК = Mid(Cell.Address, 4, 2)
        If Логирование_в_текстовые_файлы = True Then
          Print #1, "Строка_Одобрено_вкл_ОУК: " + Строка_Одобрено_вкл_ОУК
        End If
      End If
      
      ' Находим строку "Выдано"
      If Trim(CStr(Cells(intR, intC).Value)) = "Выдано" Then
        If Логирование_в_текстовые_файлы = True Then
          Print #1, "Найдена строка Строка_Выдано"
        End If
        ' Берем строку
        Строка_Выдано = Mid(Cell.Address, 4, 2)
        If Логирование_в_текстовые_файлы = True Then
          Print #1, "Строка_Выдано: " + Строка_Выдано
        End If
      End If
      
      ' Находим строку AR
      If Trim(CStr(Cells(intR, intC).Value)) = "AR (качественно), %" Then
        If Логирование_в_текстовые_файлы = True Then
          Print #1, "Найдена строка Строка_AR"
        End If
        ' Берем строку
        Строка_AR = Mid(Cell.Address, 4, 2)
        If Логирование_в_текстовые_файлы = True Then
          Print #1, "Строка_AR: " + Строка_AR
        End If
      End If
      
      ' --- Конец обработки ячейки ML-файла ---
      
    End If ' Если ячейка не пустая
        
  Next
    
  ' Копирование данных
  Range(Колонка_с_месяцем_откуда_берем_данные + Строка_всего_заявок).Copy Destination:=ThisWorkbook.Sheets("Лист1").Range("D157")
  Range(Колонка_с_месяцем_откуда_берем_данные + Строка_Одобрено_вкл_ОУК).Copy Destination:=ThisWorkbook.Sheets("Лист1").Range("E157")
  Range(Колонка_с_месяцем_откуда_берем_данные + Строка_Выдано).Copy Destination:=ThisWorkbook.Sheets("Лист1").Range("G157")
  ' AR
  Range(Колонка_с_месяцем_откуда_берем_данные + Строка_AR).Copy Destination:=ThisWorkbook.Sheets("Лист1").Range("F157")
  
  

  ' 4. Берем данные по Залоговый целевой
  ' Установка фильтра
  ActiveSheet.PivotTables("SASApp:CARDS.ALL_IPOT").PivotFields("Продукт").ClearAllFilters
  ActiveSheet.PivotTables("SASApp:CARDS.ALL_IPOT").PivotFields("Продукт").CurrentPage = "Залоговый целевой"
  
  ' Подготовка переменных - она меняется на каждом продукте
  Колонка_с_месяцем_откуда_берем_данные = ""
  Строка_всего_заявок = ""
  Строка_Одобрено_вкл_ОУК = ""
  Строка_Выдано = ""
  Строка_AR = ""
  
  ' Цикл обработки Листа Excel - определяем столбец текущего месяца из отчета $P$12 :Общий итог
  For Each Cell In ActiveSheet.UsedRange
    
    ' Если ячейка не пустая
    If Not IsEmpty(Cell) Then
      
      ' Вывод строки и столбца
      ' Номер столбца
      intC = Cell.Column
      ' Номер строки
      intR = Cell.Row

      ' Выводим все данные в строку
      If Логирование_в_текстовые_файлы = True Then
        Print #1, Cell.Address, ":" + Cell.Formula
      End If
        
      ' --- Обработка ячейки ML-файла ---
      
      ' Находим столбец откуда брать данные
      If (CStr(Cells(intR, intC).Value) = "Общий итог") And (Колонка_с_месяцем_откуда_берем_данные = "") Then
        If Логирование_в_текстовые_файлы = True Then
          Print #1, "Найден Общий итог"
        End If
        ' Берем колонку с месяцем
        Колонка_с_месяцем_откуда_берем_данные = Предидущая_буква(Mid(Cell.Address, 2, 1))
        If Логирование_в_текстовые_файлы = True Then
          Print #1, "Колонка с месяцем: " + Колонка_с_месяцем_откуда_берем_данные
        End If
      End If
      
      ' Находим строку "Всего заявок"
      If Trim(CStr(Cells(intR, intC).Value)) = "Всего заявок" Then
        If Логирование_в_текстовые_файлы = True Then
          Print #1, "Найдена строка Всего заявок"
        End If
        ' Берем строку
        Строка_всего_заявок = Mid(Cell.Address, 4, 2)
        If Логирование_в_текстовые_файлы = True Then
          Print #1, "Строка Всего заявок: " + Строка_всего_заявок
        End If
      End If
      
      ' Находим строку "Одобрено (вкл ОУК)"
      If Trim(CStr(Cells(intR, intC).Value)) = "Одобрено (вкл ОУК)" Then
        If Логирование_в_текстовые_файлы = True Then
          Print #1, "Найдена строка Одобрено (вкл ОУК)"
        End If
        ' Берем строку
        Строка_Одобрено_вкл_ОУК = Mid(Cell.Address, 4, 2)
        If Логирование_в_текстовые_файлы = True Then
          Print #1, "Строка_Одобрено_вкл_ОУК: " + Строка_Одобрено_вкл_ОУК
        End If
      End If
      
      ' Находим строку "Выдано"
      If Trim(CStr(Cells(intR, intC).Value)) = "Выдано" Then
        If Логирование_в_текстовые_файлы = True Then
          Print #1, "Найдена строка Строка_Выдано"
        End If
        ' Берем строку
        Строка_Выдано = Mid(Cell.Address, 4, 2)
        If Логирование_в_текстовые_файлы = True Then
          Print #1, "Строка_Выдано: " + Строка_Выдано
        End If
      End If
            
      ' Находим строку AR
      If Trim(CStr(Cells(intR, intC).Value)) = "AR (качественно), %" Then
        If Логирование_в_текстовые_файлы = True Then
          Print #1, "Найдена строка Строка_AR"
        End If
        ' Берем строку
        Строка_AR = Mid(Cell.Address, 4, 2)
        If Логирование_в_текстовые_файлы = True Then
          Print #1, "Строка_AR: " + Строка_AR
        End If
      End If
            
      ' --- Конец обработки ячейки ML-файла ---
      
    End If ' Если ячейка не пустая
        
  Next
  
  ' Копирование данных
  Range(Колонка_с_месяцем_откуда_берем_данные + Строка_всего_заявок).Copy Destination:=ThisWorkbook.Sheets("Лист1").Range("D158")
  Range(Колонка_с_месяцем_откуда_берем_данные + Строка_Одобрено_вкл_ОУК).Copy Destination:=ThisWorkbook.Sheets("Лист1").Range("E158")
  Range(Колонка_с_месяцем_откуда_берем_данные + Строка_Выдано).Copy Destination:=ThisWorkbook.Sheets("Лист1").Range("G158")
  ' AR
  Range(Колонка_с_месяцем_откуда_берем_данные + Строка_AR).Copy Destination:=ThisWorkbook.Sheets("Лист1").Range("F158")

  ' --- Анализируем выдачи ---
  ' Находим столбец и строку для фильтров
  Наименование_столбца_месяца_выдачи = 0
  Наименование_строки_Общий_итог_по_выдачам = 0
  Наименование_столбца_Сумма_выдачи = 0
  
  ' Строки месяц, год выдачи
  Текущий_месяц_год_по_выдачам = ""
  Текущий_месяц_по_выдачам = ""
  Текущий_год_по_выдачам = ""
  
  ' Ищем наименования нужных столбцов и строк по выдачам
  For Each Cell In ActiveSheet.UsedRange
    
    ' Если ячейка не пустая
    If Not IsEmpty(Cell) Then
      
      ' Вывод строки и столбца
      ' Номер столбца
      intC = Cell.Column
      ' Номер строки
      intR = Cell.Row
      
      ' Находим столбец откуда брать данные. В этом столбце есть ячейка "Названия строк" - столбец "S"
      If Trim(CStr(Cells(intR, intC).Value)) = "Названия строк" Then
        Наименование_столбца_месяца_выдачи = intC
      End If
      
      ' Находим строку в секторе выдачи, где занесено "Общий итог"
      If (Trim(CStr(Cells(intR, intC).Value)) = "Общий итог") And (Наименование_столбца_месяца_выдачи <> 0) Then
        Наименование_строки_Общий_итог_по_выдачам = intR
      End If
      
      ' Находим столбец Сумма выдачи
      If Trim(CStr(Cells(intR, intC).Value)) = "Сумма выдачи" Then
        Наименование_столбца_Сумма_выдачи = intC
      End If
      
      
    End If
  Next
  
  ' Текущий месяц по выдачам - берем из предидущей строки, от "Общий итог". Месяц в формате "201912"
  Текущий_месяц_год_по_выдачам = Range(ConvertToLetter(Наименование_столбца_месяца_выдачи) + CStr(Наименование_строки_Общий_итог_по_выдачам - 1))
  ' Строка "12"
  Текущий_месяц_по_выдачам = Mid(Текущий_месяц_год_по_выдачам, 5, 2)
  ' Строка "2019"
  Текущий_год_по_выдачам = Mid(Текущий_месяц_год_по_выдачам, 1, 4)
          
  ' Фильтруем данные - происходит создание Листа 1
  ActiveSheet.PivotTables("СводнаяТаблица3").PivotFields("Филиал").ClearAllFilters
  ActiveSheet.PivotTables("СводнаяТаблица3").PivotFields("Филиал").CurrentPage = "Тюменский ОО1"
  ' Range("U24").Select ' Только за декабрь
  Range(ConvertToLetter(Наименование_столбца_Сумма_выдачи) + CStr(Наименование_строки_Общий_итог_по_выдачам)).Select ' За весь год
  Selection.ShowDetail = True


  ' *** Обработка старого ML ***







  ' Перейти на открывшуюся вкладку "Лист1"
  Sheets("Лист1").Select
  
  ' Обработка строк в цикле
  rowCount = 2
  ' Выдачи
  Выдачи_все_сумма = 0
  Выдачи_все_шт = 0
  Выдачи_Военная_ипотека_шт = 0
  Выдачи_Военная_ипотека_сумма = 0
  Выдачи_Семейная_Ипотека_шт = 0
  Выдачи_Семейная_Ипотека_сумма = 0
  Выдачи_Новостройка_шт = 0
  Выдачи_Новостройка_сумма = 0
  Выдачи_Вторичный_рынок_шт = 0
  Выдачи_Вторичный_рынок_сумма = 0
  Выдачи_прочие_программы_шт = 0
  Выдачи_прочие_программы_сумма = 0
  ' Ищем нужные столбцы в строке 1
  Наименование_столбца_CREDIT_PROGRAMM_OTHER = 0
  Наименование_столбца_Сумма_выдачи = 0
  Наименование_столбца_STATUS_RETAIL = 0
  Наименование_столбца_ИНН = 0
  Наименование_столбца_Наименование_компании = 0
  Наименование_столбца_Филиал = 0
  Наименование_столбца_Адрес_предмета_залога = 0
  Наименование_столбца_FIO = 0
  Наименование_столбца_Тип_партнера = 0
  Наименование_столбца_Название_партнера = 0
  Наименование_столбца_Продукт = 0
  Наименование_столбца_Офис = 0
  Наименование_столбца_Адрес_регистрации = 0
  Наименование_столбца_Дата_выдачи = 0
  Наименование_столбца_Выдано = 0
  Наименование_столбца_Номер_кредитного_договора = 0
  Наименование_столбца_Date_iss = 0
  Наименование_столбца_Ставка_выдачи = 0
  Наименование_столбца_Срок_кредита = 0
  Наименование_столбца_Первоначальный_взнос = 0
  Наименование_столбца_Стоимость = 0
  Наименование_столбца_Сумма_выдачи_для_срвз = 0
  Наименование_столбца_CLIENT_ID = 0
  Наименование_столбца_Форма_справки_о_доходах = 0
  Наименование_столбца_Группа_компаний = 0

  ' Добавить: Одобренная сумма, Сумма выдачи из ML, потому как если это Реф, то сумма только в Сумма выдачи! (31-01-2020) и плюс в базе с ипотекой поставить человеческую дату выдачи, дату остатка и текущий остаток по кредиту

  ColumnCount = 1
  Do While (IsEmpty(Cells(1, ColumnCount)) = False)
    
    ' Ищем название столбца CREDIT_PROGRAMM_OTHER
    If Cells(1, ColumnCount).Value = "CREDIT_PROGRAMM_OTHER" Then
      Наименование_столбца_CREDIT_PROGRAMM_OTHER = ColumnCount
    End If
    ' Ищем название столбца Сумма_выдачи
    If Cells(1, ColumnCount).Value = "Сумма выдачи" Then
      Наименование_столбца_Сумма_выдачи = ColumnCount
    End If
    ' Ищем название столбца STATUS_RETAIL
    If Cells(1, ColumnCount).Value = "STATUS_RETAIL" Then
      Наименование_столбца_STATUS_RETAIL = ColumnCount
    End If
    ' Ищем название столбца ИНН
    If Cells(1, ColumnCount).Value = "ИНН" Then
      Наименование_столбца_ИНН = ColumnCount
    End If
    ' Ищем название столбца Наименование компании
    If Cells(1, ColumnCount).Value = "Наименование компании" Then
      Наименование_столбца_Наименование_компании = ColumnCount
    End If
    ' Ищем название столбца Филиал
    If Cells(1, ColumnCount).Value = "Филиал" Then
      Наименование_столбца_Филиал = ColumnCount
    End If
    ' Ищем название столбца Адрес_предмета_залога
    If Cells(1, ColumnCount).Value = "Адрес предмета залога" Then
      Наименование_столбца_Адрес_предмета_залога = ColumnCount
    End If
    ' Ищем название столбца FIO
    If Cells(1, ColumnCount).Value = "FIO" Then
      Наименование_столбца_FIO = ColumnCount
    End If
    ' Ищем название столбца Тип_партнера
    If Cells(1, ColumnCount).Value = "Тип партнера" Then
      Наименование_столбца_Тип_партнера = ColumnCount
    End If
    ' Ищем название столбца Название_партнера
    If Cells(1, ColumnCount).Value = "Название партнера" Then
      Наименование_столбца_Название_партнера = ColumnCount
    End If
    ' Ищем название столбца Продукт
    If Cells(1, ColumnCount).Value = "Продукт" Then
      Наименование_столбца_Продукт = ColumnCount
    End If
    ' Ищем название столбца Офис
    If Cells(1, ColumnCount).Value = "Офис" Then
      Наименование_столбца_Офис = ColumnCount
    End If
   ' Ищем название столбца Адрес регистрации
    If Cells(1, ColumnCount).Value = "Адрес регистрации" Then
      Наименование_столбца_Адрес_регистрации = ColumnCount
    End If
    ' Ищем название столбца Дата выдачи
    If Cells(1, ColumnCount).Value = "Дата выдачи" Then
      Наименование_столбца_Дата_выдачи = ColumnCount
    End If
    ' Ищем название столбца Выдано
    If Cells(1, ColumnCount).Value = "Выдано" Then
      Наименование_столбца_Выдано = ColumnCount
    End If
    ' Доход по осн. месту работы
    If Cells(1, ColumnCount).Value = "Доход по осн. месту работы" Then
      Наименование_столбца_Доход_по_осн_месту_работы = ColumnCount
    End If
    ' Ищем Номер договора
    If Cells(1, ColumnCount).Value = "№ кредитного договора" Then
      Наименование_столбца_Номер_кредитного_договора = ColumnCount
    End If
    ' Ищем Date_iss
    If Cells(1, ColumnCount).Value = "Date_iss" Then
      Наименование_столбца_Date_iss = ColumnCount
    End If
    ' Ищем Ставка выдачи
    If Cells(1, ColumnCount).Value = "Ставка выдачи" Then
      Наименование_столбца_Ставка_выдачи = ColumnCount
    End If
    ' Ищем Срок_кредита
    If Cells(1, ColumnCount).Value = "Срок кредита (мес)" Then
      Наименование_столбца_Срок_кредита = ColumnCount
    End If
    ' Ищем Первоначальный_взнос
    If Cells(1, ColumnCount).Value = "Первоначальный взнос" Then
      Наименование_столбца_Первоначальный_взнос = ColumnCount
    End If
    ' Ищем Стоимость
    If Cells(1, ColumnCount).Value = "Стоимость (для ПВ)" Then
      Наименование_столбца_Стоимость = ColumnCount
    End If
    ' Ищем Сумма выдачи_для срвз
    If Cells(1, ColumnCount).Value = "Сумма выдачи_для срвз" Then
      Наименование_столбца_Сумма_выдачи_для_срвз = ColumnCount
    End If
    ' Ищем CLIENT_ID
    If Cells(1, ColumnCount).Value = "CLIENT_ID" Then
      Наименование_столбца_CLIENT_ID = ColumnCount
    End If
    ' Ищем Форма справки о доходах
    If Cells(1, ColumnCount).Value = "Форма справки о доходах" Then
      Наименование_столбца_Форма_справки_о_доходах = ColumnCount
    End If
    ' Ищем Группа_компаний
    If Cells(1, ColumnCount).Value = "Группа компаний" Then
      Наименование_столбца_Группа_компаний = ColumnCount
    End If
    
    ' Следующая запись
    ColumnCount = ColumnCount + 1
  
  Loop ' Обработка строк в цикле
  
  ' Выводим все записи по столбец Филиал = "Тюменский ОО1"
  Do While (Trim(Cells(rowCount, Наименование_столбца_Филиал).Value) = "Тюменский ОО1")
   
    ' Обработка строки если "Кредит выдан (Кредит выдан)"
    ' If Cells(RowCount, Наименование_столбца_STATUS_RETAIL).Value = "Кредит выдан (Кредит выдан)" Then
    
    ' В старых версиях ML-файла нет столбца STATUS_RETAIL, поэтому переходим на "Выдано" = "1"
    If Cells(rowCount, Наименование_столбца_Выдано).Value = "1" Then
    
     ' Строка "12" Текущий_месяц_по_выдачам, Строка "2019" Текущий_год_по_выдачам
     If (Mid(CStr(CDate(Cells(rowCount, Наименование_столбца_Дата_выдачи).Value)), 4, 2) = Текущий_месяц_по_выдачам) And (Mid(CStr(CDate(Cells(rowCount, Наименование_столбца_Дата_выдачи).Value)), 7, 4) = Текущий_год_по_выдачам) Then

      ' Выдачи по программам
      Выдачи_все_сумма = Выдачи_все_сумма + Cells(rowCount, Наименование_столбца_Сумма_выдачи).Value
      Выдачи_все_шт = Выдачи_все_шт + 1
      ' Переменная Учтено
      Учитано_в_категории_программ = False
      ' Военная_ипотека
      If Cells(rowCount, Наименование_столбца_CREDIT_PROGRAMM_OTHER).Value = "ВОЕННАЯ ИПОТЕКА/РЕФИНАНСИРОВАНИЕ ВОЕННОЙ ИПОТЕКИ/ЗАЛОГ ПРИОБРЕТАЕМОЙ НЕДВИЖИМОСТИ/КВАРТИРА/ВТОРИЧНЫЙ РЫНОК" Then
        Выдачи_Военная_ипотека_шт = Выдачи_Военная_ипотека_шт + 1
        Выдачи_Военная_ипотека_сумма = Выдачи_Военная_ипотека_сумма + Cells(rowCount, Наименование_столбца_Сумма_выдачи).Value
        Учитано_в_категории_программ = True
      End If
      ' Выдачи_Семейная_Ипотека
      If (Cells(rowCount, Наименование_столбца_CREDIT_PROGRAMM_OTHER).Value = "ГОС.ПРОГРАММА/СЕМЕЙНАЯ ИПОТЕКА/ЗАЛОГ ПРИОБРЕТАЕМОЙ НЕДВИЖИМОСТИ/КВАРТИРА/ВТОРИЧНЫЙ РЫНОК") Or (Cells(rowCount, "AW").Value = "ГОС.ПРОГРАММА/СЕМЕЙНАЯ ИПОТЕКА/ЗАЛОГ ПРИОБРЕТАЕМОЙ НЕДВИЖИМОСТИ/КВАРТИРА/ПЕРВИЧНЫЙ РЫНОК") Then
        Выдачи_Семейная_Ипотека_шт = Выдачи_Семейная_Ипотека_шт + 1
        Выдачи_Семейная_Ипотека_сумма = Выдачи_Семейная_Ипотека_сумма + Cells(rowCount, Наименование_столбца_Сумма_выдачи).Value
        Учитано_в_категории_программ = True
      End If
      ' Выдачи_Новостройка
      If Cells(rowCount, Наименование_столбца_CREDIT_PROGRAMM_OTHER).Value = "ЗАЛОГ ПРИОБРЕТАЕМОЙ НЕДВИЖИМОСТИ(ПРАВ)/КВАРТИРА/ ПЕРВИЧНЫЙ РЫНОК" Then
        Выдачи_Новостройка_шт = Выдачи_Новостройка_шт + 1
        Выдачи_Новостройка_сумма = Выдачи_Новостройка_сумма + Cells(rowCount, Наименование_столбца_Сумма_выдачи).Value
        Учитано_в_категории_программ = True
      End If
      ' Выдачи_Вторичный_рынок
      If Cells(rowCount, Наименование_столбца_CREDIT_PROGRAMM_OTHER).Value = "ЗАЛОГ ПРИОБРЕТАЕМОЙ НЕДВИЖИМОСТИ(ПРАВ)/КВАРТИРА/ ВТОРИЧНЫЙ РЫНОК" Then
        Выдачи_Вторичный_рынок_шт = Выдачи_Вторичный_рынок_шт + 1
        Выдачи_Вторичный_рынок_сумма = Выдачи_Вторичный_рынок_сумма + Cells(rowCount, Наименование_столбца_Сумма_выдачи).Value
        Учитано_в_категории_программ = True
      End If
      ' Прочие программы
      If Учитано_в_категории_программ = False Then
        Выдачи_прочие_программы_шт = Выдачи_прочие_программы_шт + 1
        Выдачи_прочие_программы_сумма = Выдачи_прочие_программы_сумма + Cells(rowCount, Наименование_столбца_Сумма_выдачи).Value
        Учитано_в_категории_программ = True
      End If

      End If ' Текущий месяц и год
      
    End If ' Кредит выдан
      
    ' Следующая запись
    rowCount = rowCount + 1
  Loop ' Обработка строк в цикле
  
  ' Заносим данные в мою таблицу
  
  ' Выдачи Военная ипотека
  ThisWorkbook.Sheets("Лист1").Range("D148").Value = Выдачи_Военная_ипотека_сумма / 1000
  ThisWorkbook.Sheets("Лист1").Range("F148").Value = Выдачи_Военная_ипотека_шт
  ' Выдачи_Семейная_Ипотека
  ThisWorkbook.Sheets("Лист1").Range("D149").Value = Выдачи_Семейная_Ипотека_сумма / 1000
  ThisWorkbook.Sheets("Лист1").Range("F149").Value = Выдачи_Семейная_Ипотека_шт
  ' Выдачи_Новостройка
  ThisWorkbook.Sheets("Лист1").Range("D150").Value = Выдачи_Новостройка_сумма / 1000
  ThisWorkbook.Sheets("Лист1").Range("F150").Value = Выдачи_Новостройка_шт
  ' Выдачи_Вторичный_рынок
  ThisWorkbook.Sheets("Лист1").Range("D151").Value = Выдачи_Вторичный_рынок_сумма / 1000
  ThisWorkbook.Sheets("Лист1").Range("F151").Value = Выдачи_Вторичный_рынок_шт
  ' Прочие программы
  ThisWorkbook.Sheets("Лист1").Range("D152").Value = Выдачи_прочие_программы_сумма / 1000
  ThisWorkbook.Sheets("Лист1").Range("F152").Value = Выдачи_прочие_программы_шт
  ' Выполнение_сумма_факт
  ThisWorkbook.Sheets("Лист1").Range("D147").Value = Round(Выдачи_все_сумма / 1000, 0)
  ' Для ячейки устанавливаем формат без знаков после запятой
  ThisWorkbook.Sheets("Лист1").Range("D147").NumberFormat = "#,##0"
  ' Выполнение_шт_факт
  ThisWorkbook.Sheets("Лист1").Range("F147").Value = Выдачи_все_шт
  ' Вставляем формулу с % выполнения
  ThisWorkbook.Sheets("Лист1").Range("G147").FormulaR1C1 = "=IF(RC[-3]>0,((RC[-3]*100)/RC[-4])/100,0)"
  ' Проставляем доли и целевые ориентиры
  ' Доля СЖ (строящегося жилья) не менее 50-60%
  ' Формула =ЕСЛИ(D147>0;((Факт*100)/План)/100;0)
  ThisWorkbook.Sheets("Лист1").Range("G150").Value = "Доля " + CStr(Round(((ThisWorkbook.Sheets("Лист1").Range("F150").Value * 100) / ThisWorkbook.Sheets("Лист1").Range("F147").Value), 1)) + "%"
  ThisWorkbook.Sheets("Лист1").Range("H150").Value = "" ' "(Норматив: 50%)"
  ' Доля по Семейной ипотеке не менее 20%
  ThisWorkbook.Sheets("Лист1").Range("G149").Value = "Доля " + CStr(Round(((ThisWorkbook.Sheets("Лист1").Range("F149").Value * 100) / ThisWorkbook.Sheets("Лист1").Range("F147").Value), 1)) + "%"
  ThisWorkbook.Sheets("Лист1").Range("H149").Value = "" ' "(Норматив: 20%)"
  
  ' TR по СЖ не менее 70%
  ' TR по ГЖ не менее 35%
       
  ' Цели на 2020
  ' Добрый день!
  ' Целевые ориентиры перед всеми ИЦ стоят:
  ' по  выдачам,
  ' КСП,
  ' доле ВИ/ГИ -40/60,
  ' доле аккредитивов,
  ' доле SRG отчетов,
  ' направления дорожной карты по выполнению поставленного плана.
  ' Также РИЦ ставят себе 3  KPI на месяц:
  ' Валентина поставила задачу на январь:
  ' план по заявкам от партнеров 460 шт.
  ' по  изменению воронки продаж с 15/85 на 35/65 (СЖ/ГЖ)
  ' по увеличению общего показателя  TR с 18% (2019 г.) до 35%, на горизонте квартала - до 50%.
       
  ' --- Конец Анализируем выдачи ---

  ' --- Выборка потенциальных зарплатников ---
        
  ' Сортировка Листа 1 по столбцу ИНН Организации клиента, получившего ипотеку
  ActiveWorkbook.Worksheets("Лист1").Range(ConvertToLetter(Наименование_столбца_ИНН) + "2").Select
  ActiveWorkbook.Worksheets("Лист1").ListObjects("Таблица1").Sort.SortFields.Clear
  ActiveWorkbook.Worksheets("Лист1").ListObjects("Таблица1").Sort.SortFields.Add Key:=Range(ConvertToLetter(Наименование_столбца_ИНН) + "2"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    
  ' Сама сортировка (взято из записанного макроса)
  With ActiveWorkbook.Worksheets("Лист1").ListObjects("Таблица1").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
  End With
  
  ' Обновить данные в таблице BASE\Mortgage по найденным в ней договорам из ML
  If param_from_ini(ThisWorkbook.Path + "\DB_Result.ini", "Обновить_по_найденным") = "1" Then
    Обновить_по_найденным = True
  Else
    Обновить_по_найденным = False
  End If
  
  ' Выполняем цикл по всем записям у которых в первом столбце либо 0 либо 1
  Текущий_ИНН = ""
  Текущая_организация_наименование = ""
  count_Текущий_ИНН = 0
  счетчик_строк_в_DB_Result_Лист2 = 4
  ' Счетчик - всего выведено Организаций в мой Лист 2  (счет идет с нуля)
  всего_выведено_организаций = 0
  ' Начинаем со второй записи
  rowCount = 2
  ' Здесь идем по столбцу "Филиал", в котором должно быть "Тюменский ОО1"
  Do While (Cells(rowCount, Наименование_столбца_Филиал) = ("Тюменский ОО1"))
      
    ' Если у записи кредит выдан?
    ' If Cells(RowCount, Наименование_столбца_STATUS_RETAIL).Value = "Кредит выдан (Кредит выдан)" Then В старых версиях ML-файла нет столбца STATUS_RETAIL, поэтому переходим на "Выдано" = "1"
    If Cells(rowCount, Наименование_столбца_Выдано).Value = "1" Then
               
      ' Если текущий ИНН не равен предыдущему, то это начало
      If Текущий_ИНН <> Cells(rowCount, Наименование_столбца_ИНН).Value Then
                
        ' Вставляем в лист запись с наименованием организации и числом кредитов, если счетчик не равен 0
        If count_Текущий_ИНН > 0 Then
          Call Вывод_в_отчет_итогов_по_Организации_2(In_MLName, Текущий_ИНН, count_Текущий_ИНН, счетчик_строк_в_DB_Result_Лист2, Текущая_организация_наименование, False)
          ' Счет идет не с нуля а с номер строки вывода данных в Листе 1
          счетчик_строк_в_DB_Result_Лист2 = счетчик_строк_в_DB_Result_Лист2 + 1
          ' Счетчик - всего выведено Организаций в мой Лист 2  (счет идет с нуля)
          всего_выведено_организаций = всего_выведено_организаций + 1
        End If
        
        ' Обнуляем счетчик
        count_Текущий_ИНН = 0
        
        ' Присваиваем значение ИНН
        Текущий_ИНН = Cells(rowCount, Наименование_столбца_ИНН).Value
        ' Наименование компании
        Текущая_организация_наименование = Cells(rowCount, Наименование_столбца_Наименование_компании).Value
        
        ' Данные по ипотечной заявке
        ' Адрес предмета залога
        ' FIO
        ' Тип партнера
        ' Название партнера
        ' Продукт
        ' Офис (Тюменский или Сургутский)

      End If
      
      ' Считаем выданные кредиты на организации
      count_Текущий_ИНН = count_Текущий_ИНН + 1
      
    End If
    
    ' Следующая запись
    rowCount = rowCount + 1
  Loop ' Обработка строк в цикле
  
  ' Проверяем - если последняя запись была выведена в отчет?
  If count_Текущий_ИНН > 0 Then
    Call Вывод_в_отчет_итогов_по_Организации_2(In_MLName, Текущий_ИНН, count_Текущий_ИНН, счетчик_строк_в_DB_Result_Лист2, Текущая_организация_наименование, False)
    счетчик_строк_в_DB_Result_Лист2 = счетчик_строк_в_DB_Result_Лист2 + 1
    ' Счетчик - всего выведено Организаций в мой Лист 2 (счет идет с нуля)
    всего_выведено_организаций = всего_выведено_организаций + 1
  End If
    
  ' Перейти на мой Лист 2 и отсортировать
  ThisWorkbook.Activate
  ThisWorkbook.Sheets("Лист2").Select
  Range("D4").Select
  ThisWorkbook.Worksheets("Лист2").Sort.SortFields.Clear
  ThisWorkbook.Worksheets("Лист2").Sort.SortFields.Add Key:=Range("D4"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ThisWorkbook.Worksheets("Лист2").Sort
        .SetRange Range("B3:D" + CStr(всего_выведено_организаций))
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

  ' Теперь вносим в отсортированную форму на Листе 2 ФИО клиентов
  If True Then
    
    ' Счетчик строк
    rowCount = 2
  
    ' Здесь идем по столбцу "Филиал", в котором должно быть "Тюменский ОО1"
    Do While (Workbooks(In_MLName).Worksheets("Лист1").Cells(rowCount, Наименование_столбца_Филиал) = "Тюменский ОО1")
      
      ' Если у записи кредит выдан?
      If Workbooks(In_MLName).Worksheets("Лист1").Cells(rowCount, Наименование_столбца_Выдано).Value = "1" Then
          
        ' Присваиваем значение ИНН
        If Workbooks(In_MLName).Worksheets("Лист1").Cells(rowCount, Наименование_столбца_ИНН).Value <> "" Then
          Текущий_ИНН = Trim(Workbooks(In_MLName).Worksheets("Лист1").Cells(rowCount, Наименование_столбца_ИНН).Value)
        Else
          Текущий_ИНН = "Нет ИНН"
        End If
                
        ' Выдано
        Текущий_Выдано = Workbooks(In_MLName).Worksheets("Лист1").Cells(rowCount, Наименование_столбца_Выдано).Value
        ' Дата выдачи - CDate(Cells(RowCount, Наименование_столбца_Дата_выдачи
        Текущий_Дата_выдачи = CDate(Workbooks(In_MLName).Worksheets("Лист1").Cells(rowCount, Наименование_столбца_Дата_выдачи).Value)
        ' Наименование компании
        Текущая_организация_наименование = Workbooks(In_MLName).Worksheets("Лист1").Cells(rowCount, Наименование_столбца_Наименование_компании).Value
        ' Адрес предмета залога
        Текущий_Адрес_предмета_залога = Workbooks(In_MLName).Worksheets("Лист1").Cells(rowCount, Наименование_столбца_Адрес_предмета_залога).Value
        ' FIO
        Текущий_FIO = Workbooks(In_MLName).Worksheets("Лист1").Cells(rowCount, Наименование_столбца_FIO).Value
        ' Тип партнера
        ' Название партнера
        Текущий_Название_партнера = Workbooks(In_MLName).Worksheets("Лист1").Cells(rowCount, Наименование_столбца_Название_партнера).Value
        ' Продукт
        ' Филиал
        Текущий_Филиал = Workbooks(In_MLName).Worksheets("Лист1").Cells(rowCount, Наименование_столбца_Филиал).Value
        ' Офис (Тюменский или Сургутский)
        Текущий_Офис = Workbooks(In_MLName).Worksheets("Лист1").Cells(rowCount, Наименование_столбца_Офис).Value
        ' Доход по осн месту работы
        Текущий_Доход_по_осн_месту_работы = Round(Workbooks(In_MLName).Worksheets("Лист1").Cells(rowCount, Наименование_столбца_Доход_по_осн_месту_работы).Value / 1000, 1)
        ' Номер договора
        Текущий_Номер_кредитного_договора = Workbooks(In_MLName).Worksheets("Лист1").Cells(rowCount, Наименование_столбца_Номер_кредитного_договора).Value
        ' Текущий_Date_iss
        Текущий_Date_iss = Workbooks(In_MLName).Worksheets("Лист1").Cells(rowCount, Наименование_столбца_Date_iss).Value
        ' Текущий_Ставка_выдачи
        Текущий_Ставка_выдачи = Workbooks(In_MLName).Worksheets("Лист1").Cells(rowCount, Наименование_столбца_Ставка_выдачи).Value
        ' Текущий_Срок_кредита
        Текущий_Срок_кредита = Workbooks(In_MLName).Worksheets("Лист1").Cells(rowCount, Наименование_столбца_Срок_кредита).Value
        ' + Текущий_CREDIT_PROGRAMM_OTHER
        Текущий_CREDIT_PROGRAMM_OTHER = Workbooks(In_MLName).Worksheets("Лист1").Cells(rowCount, Наименование_столбца_CREDIT_PROGRAMM_OTHER).Value
        ' Текущий_Первоначальный_взнос
        Текущий_Первоначальный_взнос = Workbooks(In_MLName).Worksheets("Лист1").Cells(rowCount, Наименование_столбца_Первоначальный_взнос).Value
        ' Текущий_Стоимость
        Текущий_Стоимость = Workbooks(In_MLName).Worksheets("Лист1").Cells(rowCount, Наименование_столбца_Стоимость).Value
        ' + Текущий_Адрес_предмета_залога
        ' Текущий_Сумма_выдачи_для_срвз
        Текущий_Сумма_выдачи_для_срвз = Workbooks(In_MLName).Worksheets("Лист1").Cells(rowCount, Наименование_столбца_Сумма_выдачи_для_срвз).Value
        ' Текущий_CLIENT_ID
        Текущий_CLIENT_ID = Workbooks(In_MLName).Worksheets("Лист1").Cells(rowCount, Наименование_столбца_CLIENT_ID).Value
        ' + Текущий_FIO
        ' + Текущий_Доход_по_осн_месту_работы
        ' Текущий_Форма_справки_о_доходах
        Текущий_Форма_справки_о_доходах = Workbooks(In_MLName).Worksheets("Лист1").Cells(rowCount, Наименование_столбца_Форма_справки_о_доходах).Value
        ' + Текущий_Наименование_компании
        ' + Текущий_ИНН
        ' + Текущий_Название_партнера
        ' Текущий_Тип_партнера
        Текущий_Тип_партнера = Workbooks(In_MLName).Worksheets("Лист1").Cells(rowCount, Наименование_столбца_Тип_партнера).Value
        ' Текущий_Группа_компаний
        Текущий_Группа_компаний = Workbooks(In_MLName).Worksheets("Лист1").Cells(rowCount, Наименование_столбца_Группа_компаний).Value
        ' В DB_Result на Лист 2 по столбцу "B" ищем ИНН
        ' Если в текущем ИНН первый нуль, то его удаляем
        If Mid(Текущий_ИНН, 1, 1) = "0" Then
          Текущий_ИНН = Mid(Текущий_ИНН, 2, Len(Текущий_ИНН) - 1)
        End If
        
        ' Выполняем поиск
        Set fcell = Columns("B:B").Find(Текущий_ИНН, LookAt:=xlWhole)
        If Not fcell Is Nothing Then
          
          ' MsgBox "Нашел в строке: " + CStr(fcell.Row)
          
          ' Выделяем данную запись, следующую за найденной Rows("6:6").Select
          Rows(CStr(fcell.Row + 1) + ":" + CStr(fcell.Row + 1)).Select
          
          ' Вставляем новую строку
          Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
          
          ' Вносим данные в новую строку
          ThisWorkbook.Sheets("Лист2").Range("C" + CStr(fcell.Row + 1)).Value = Текущий_FIO
          ' Подтвержденный доход - "Доход по осн. месту работы"
          ' Для сотрудников ПСБ с ИНН 7744000912 обнуляем доход
          If Текущий_ИНН = "7744000912" Then
            Текущий_Доход_по_осн_месту_работы = 0
          End If
          ' Доход_по_осн_месту_работы
          ThisWorkbook.Sheets("Лист2").Range("D" + CStr(fcell.Row + 1)).Value = Текущий_Доход_по_осн_месту_работы
          ' Адрес объекта
          ThisWorkbook.Sheets("Лист2").Range("E" + CStr(fcell.Row + 1)).Value = Значение_между_разделителями(Текущий_Адрес_предмета_залога, ";", 2, 3)
          ' Партнер
          ThisWorkbook.Sheets("Лист2").Range("F" + CStr(fcell.Row + 1)).Value = Текущий_Название_партнера
          ' Дата сделки
          ThisWorkbook.Sheets("Лист2").Range("G" + CStr(fcell.Row + 1)).Value = CStr(Текущий_Дата_выдачи)
          ' Ставка
          ThisWorkbook.Sheets("Лист2").Range("H" + CStr(fcell.Row + 1)).Value = Текущий_Ставка_выдачи
          ' В столбец H вносим данные по наименованию организации, если нет ИНН
          If Текущий_ИНН = "Нет ИНН" Then
            ThisWorkbook.Sheets("Лист2").Range("I" + CStr(fcell.Row + 1)).Value = "Организация: " + Текущая_организация_наименование
          End If
          ' В столбец J вносим Client ID
          ThisWorkbook.Sheets("Лист2").Range("J" + CStr(fcell.Row + 1)).Value = Текущий_CLIENT_ID
                    
        End If
                    
        ' Вставляем запись в Таблицу BASE\Mortgage
        Call Insert_To_Table_Mortgage(Текущий_Номер_кредитного_договора, Текущий_Date_iss, Текущий_Ставка_выдачи, Текущий_Срок_кредита, Текущий_CREDIT_PROGRAMM_OTHER, Текущий_Первоначальный_взнос, Текущий_Стоимость, Текущий_Адрес_предмета_залога, Текущий_Сумма_выдачи_для_срвз, Текущий_CLIENT_ID, Текущий_FIO, Текущий_Доход_по_осн_месту_работы, Текущий_Форма_справки_о_доходах, Текущая_организация_наименование, Текущий_ИНН, Текущий_Название_партнера, Текущий_Тип_партнера, Текущий_Группа_компаний, Текущий_Выдано, Текущий_Филиал, Текущий_Офис, Обновить_по_найденным)
                                                           
      End If
    
      ' Следующая запись
      rowCount = rowCount + 1
    
    Loop ' Обработка строк в цикле

    ' Вторая итерация - группировка клиентов по ИНН
      
    ' Счетчик строк - начинаем с первой записи
    rowCount = 4
    Начало_блока_группировки_записей = 0
    Блок_Начат = False
    ' Пока в "ИНН" и "Наименование организации" не будет пусто
    Do While (ThisWorkbook.Sheets("Лист2").Cells(rowCount, 2).Value <> "") Or (ThisWorkbook.Sheets("Лист2").Cells(rowCount, 3).Value <> "")
      
      ' Начало блока группировки
      If (ThisWorkbook.Sheets("Лист2").Cells(rowCount, 2).Value = "") And (ThisWorkbook.Sheets("Лист2").Cells(rowCount, 3).Value <> "") And (Блок_Начат = False) Then
        ' Начало блока группировки записей
        Начало_блока_группировки_записей = rowCount
        Блок_Начат = True
      End If

      ' Конец блока группировки
      If (ThisWorkbook.Sheets("Лист2").Cells(rowCount, 2).Value <> "") And (ThisWorkbook.Sheets("Лист2").Cells(rowCount, 3).Value <> "") And (Начало_блока_группировки_записей <> 0) Then
        ' Группируем с начального блока до -1
        ThisWorkbook.Sheets("Лист2").Range("C" + CStr(Начало_блока_группировки_записей) + ":C" + CStr(rowCount - 1)).Select
        Selection.Rows.Group
        Блок_Начат = False
      End If
      
      ' Следующая запись
      rowCount = rowCount + 1
    
    Loop ' Обработка строк в цикле
        
    Columns("E:I").Select
    Selection.Columns.Group
                
    ' Закрываем список
    ActiveSheet.Outline.ShowLevels RowLevels:=1, ColumnLevels:=1

    ' Переходим в первую ячейку
    Range("A1").Select
    
  End If
  
  ' Закрываем таблицу BASE\Mortgage.xlsx с сохранением внесенных изменений
  Workbooks("Mortgage.xlsx").Close SaveChanges:=True
  
  ' --- Конец Выборка потенциальных зарплатников ---

  ' Логирование
  If Логирование_в_текстовые_файлы = True Then
    ' Закрываем файлы
    Close #1
    Close #2
  End If

End Sub

' ===================================
' === Выборка_данных_по_DB Макрос ===
' ===================================
Sub Выборка_данных_по_Db()
Attribute Выборка_данных_по_Db.VB_ProcData.VB_Invoke_Func = " \n14"

' Переменные
Dim FileName As String
Dim DBstrName_String As String
Dim Обработка_завершена, Открываем_Capacity_Model, Открываем_CRM_ML, Открываем_реестр_Ипотека As Boolean


' Запуск Диалога окрытия файла DB Dashboard_new_РБ_ДД.ММ.ГГГГ.xlsm - выходит ошибка типов при использовании аргументов
' FileName = Application.GetOpenFilename("Excel files(*.xlsm), *.xlsm, All files(*.*), *.*", 1, "Открытие Dashboard", , True)
FileName = Application.GetOpenFilename("Excel Files (*.xlsm), *.xlsm", , "Открытие файла Dashboard")

' DBstrName = getFName(Application.GetOpenFilename())
Обработка_завершена = False

' Проводим очистку Листа с итогами
Call Очистка_данных_на_Листе1

' Запросы на открытие дополнительных Книг:
Открываем_Capacity_Model = False
Открываем_CRM_ML = False
Открываем_реестр_Ипотека = False

' Открыть книгу Capacity Model для получения данных
' If (Len(FileName) > 5) Then
'  If MsgBox("Открыть Capacity Model?", vbYesNo) = vbYes Then
'    CapacityModelName = getFName(Application.GetOpenFilename())
'    If CapacityModelName <> False Then
'      Workbooks.Open (CapacityModelName)
'      Открываем_Capacity_Model = True
'    End If
'  End If
' End If

' Открыть книгу CRM_ML для получения данных по заявкам
If param_from_ini(ThisWorkbook.Path + "\DB_Result.ini", "Call_Ипотека_ML") = "1" Then
  ' Открытие
  If (Len(FileName) > 5) Then
    If MsgBox("Открыть ML-файл?", vbYesNo) = vbYes Then
      ' .xlsx
      MLName = getFName(Application.GetOpenFilename("Excel Files (*.xlsx), *.xlsx", , "Открытие файла ML"))
      If MLName <> False Then
        ' Workbooks.Open (MLName) - Открытие переносим ниже
        Открываем_CRM_ML = True
      End If
    End If
  End If
End If

' Вопрос - нужен ли он или брать все из ML?
' Открыть книгу Реестр Ипотека для информации по заявкам:
' If (Len(FileName) > 5) Then
'   If MsgBox("Открыть Реестр Ипотека?", vbYesNo) = vbYes Then
'     ReestrIpotekaName = getFName(Application.GetOpenFilename())
'     If ReestrIpotekaName <> False Then
'       Workbooks.Open (ReestrIpotekaName)
'       Открываем_реестр_Ипотека = True
'     End If
'   End If
' End If

' Проверка - выбрана ли Книга?
If (Len(FileName) > 5) Then

  ' Проводим очистку Листа с итогами
  ' Call Очистка_данных_на_Листе1
  
  ' Выводим для инфо данные об имени файла
  DBstrName_String = Dir(FileName)
  Range("A1") = "Имя файла импорта данных: " + DBstrName_String
  
  ' Открываем выбранную книгу (UpdateLinks:=0)
  Workbooks.Open FileName, 0
  
  ' 1. Импортирование данных с Листа: 4. Интегральный рей-г по сотруд
  If param_from_ini(ThisWorkbook.Path + "\DB_Result.ini", "Call_Интегральный_рейтинг_по_сотрудникам") = "1" Then
    Call Интегральный_рейтинг_по_сотрудникам(DBstrName_String)
    DoEventsInterval (100)
  End If

  ' 2. Рейтинг по офисам: 1.1Интег-ый рейтинг  по офисам Примечание: до августа 2019 г. в DB Листа "1.1Интег-ый рейтинг  по офисам Примечание" не было
  If param_from_ini(ThisWorkbook.Path + "\DB_Result.ini", "Call_Интегральный_рейтинг_по_офисам") = "1" Then
    Call Интегральный_рейтинг_по_офисам(DBstrName_String)
    DoEventsInterval (100)
  End If

  ' 3. Выполнение плана по ПК: 3.1 Потребительские  кредиты. Число офисов=5. Примечание: ранее лист назывался "2.1 Потребительские  кредиты"
  If param_from_ini(ThisWorkbook.Path + "\DB_Result.ini", "Call_Потребительские_кредиты") = "1" Then
    Call Потребительские_кредиты(DBstrName_String, 5)
    DoEventsInterval (100)
  End If

  ' 4. Выполнение плана по КК: 3.2 Кредитные карты. Число офисов=5 Примечание: ранее лист назывался "2.2 Кредитные карты"
  If param_from_ini(ThisWorkbook.Path + "\DB_Result.ini", "Call_Кредитные_карты") = "1" Then
    Call Кредитные_карты(DBstrName_String, 5)
    DoEventsInterval (100)
  End If

  ' 5. Выполнение плана по Комдоходу: 2. Ком доход. Число офисов=5 Примечание: ранее лист назывался "3. Ком доход"
  If param_from_ini(ThisWorkbook.Path + "\DB_Result.ini", "Call_Ком_доход") = "1" Then
    Call Ком_доход(DBstrName_String, 5)
    DoEventsInterval (100)
  End If
  
  ' 6. Выполнение плана по ОРС: 3.10 OPC. Число офисов=5 Примечание: ранее лист назывался "2.11 OPC"
  If param_from_ini(ThisWorkbook.Path + "\DB_Result.ini", "Call_OPC") = "1" Then
    Call OPC(DBstrName_String, 5)
    DoEventsInterval (100)
  End If
  
  ' 7. Выполнение плана по Ипотеке: 3.15 Ипотека Примечание: ранее лист назывался "2.14 Ипотека"
  If param_from_ini(ThisWorkbook.Path + "\DB_Result.ini", "Call_Ипотека") = "1" Then
    Call Ипотека(DBstrName_String)
    DoEventsInterval (100)
  End If
  
  ' 8. Ипотека из ML-файла
  If param_from_ini(ThisWorkbook.Path + "\DB_Result.ini", "Call_Ипотека_ML") = "1" Then
  
    If Открываем_CRM_ML = True Then
      Workbooks.Open (MLName)
      Call Ипотека_ML(MLName)
      DoEventsInterval (100)
    End If
  
  End If
  
  ' Закрываем книгу откуда мы скопировали данные без сохранения изменений (параметр SaveChanges:=False)
  Workbooks(Dir(FileName)).Close SaveChanges:=False
  
  Обработка_завершена = True

' Проверка - выбрана ли Книга?
End If

' Если были открыты остальные книги - закрываем
' Открыть книгу Capacity Model для получения данных
If Открываем_Capacity_Model = True Then
  Workbooks(CapacityModelName).Close SaveChanges:=False
End If

' Открыть книгу CRM_ML для получения данных по заявкам
If Открываем_CRM_ML = True Then
  Workbooks(MLName).Close SaveChanges:=False
End If

' Открыть книгу Реестр Ипотека для информации по заявкам:
If Открываем_реестр_Ипотека = True Then
  Workbooks(ReestrIpotekaName).Close SaveChanges:=False
End If

' Сброс StatusBar и передача управления этой строкой в Excel
Application.StatusBar = False


' Сообщение Обработка_завершена
If Обработка_завершена = True Then
  
  ' Тема для письма
  ThisWorkbook.Sheets("Лист1").Range("P5").Value = " на " + Mid(DBstrName_String, 18, 10) + " г."

  ' Зачеркиваем пункт меню на стартовой страницы
  ' ThisWorkbook.Sheets("Лист0").Cells(7, 4).Value = "1) DashBoard за " + Mid(DBstrName_String, 18, 10) + " обработан"
  
  ' Call ЗачеркиваемТекстВячейке("Лист0", "D7")
  Call ЗачеркиваемТекстВячейке("Лист0", RangeByValue(ThisWorkbook.Name, "Лист0", "DashBoard (при наличии)", 100, 100))
  
  ' Переходим на Лист 1
  ThisWorkbook.Sheets("Лист1").Select
  Range("A1").Select
  MsgBox ("Обработка " + DBstrName_String + " завершена!")
  
End If

End Sub



' Где-то использовалась ...
Sub Форматирование_ячейки()

' Форматирование_ячейки Макрос
    Range("C17").Select
    Selection.Copy
    Range("D10").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
End Sub

Sub Очистка_данных_на_Листе1()
'
' Очистка_данных_на_Листе1 Макрос

    Range("A10:O48").Select
    Selection.ClearContents
    Range("O10").Select
    Selection.Copy
    Range("N10:N27").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
        
    Range("B29:B47").Select
    Selection.ClearContents
        
    Range("A52:K56").Select
    Selection.ClearContents
    
    Range("B57:B67").Select
    Selection.ClearContents
    
    Range("A71:L75").Select
    Selection.ClearContents
    
    Range("B76:B86").Select
    Selection.ClearContents
    
    Range("A90:O94").Select
    Selection.ClearContents
    
    Range("B95:B105").Select
    Selection.ClearContents
    Range("I90").Select
    Selection.Copy
    Range("I94").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    
    Range("A109:N114").Select
    Selection.ClearContents
    
    Range("B114:B124").Select
    Selection.ClearContents
    
    Range("A128:J132").Select
    Selection.ClearContents
    
    Range("B133:B143").Select
    Selection.ClearContents
    
    ' Ипотека
    Range("A147:I147").Select
    Selection.ClearContents
    
    Range("C148:I152").Select
    Selection.ClearContents
    
    Range("C155:G158").Select
    Selection.ClearContents
    
    Range("C159").Select
    ActiveCell.FormulaR1C1 = "0"
        
    ' Раскрыть все скрытые столбцы и строки
    Columns("A:O").Select
    Selection.EntireColumn.Hidden = False
    Rows("1:174").Select
    Selection.EntireRow.Hidden = False
    
    Range("A1").Select
    
    ' Очистка на Листе 2
    Sheets("Лист2").Select
    ' Удаляем группировку строк на листе
    Selection.ClearOutline
    
    ' Range("B4:J1000").Select
    Range("B4:J50000").Select
    Selection.ClearContents
    Range("A1").Select
    
    ' Бизнес-справка - не чистим!
    ' ThisWorkbook.Sheets("Лист3").Cells(2, 2).Value = "Оперативная бизнес-справка (активы) на <...>"
    ' Заносим даты в заголовки C5, Факт на <...>
    ' ThisWorkbook.Sheets("Лист3").Cells(5, 3).Value = "Факт на <...>"
    ' Заносим даты в заголовки I4, <...>, тыс. руб.
    ' ThisWorkbook.Sheets("Лист3").Cells(4, 9).Value = "Выдачи за <...>"
    ' Ипотека: Заносим даты в заголовки C5, Факт на <...>
    ' ThisWorkbook.Sheets("Лист3").Cells(18, 3).Value = "Факт на <...>"
    ' Ипотека: Заносим даты в заголовки I4, <...>, тыс. руб.
    ' ThisWorkbook.Sheets("Лист3").Cells(17, 9).Value = "Выдачи за <...>"

    ' Возврат на исходный Лист 1
    Sheets("Лист1").Select
    
End Sub

