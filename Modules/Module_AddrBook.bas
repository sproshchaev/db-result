Attribute VB_Name = "Module_AddrBook"
' Addr.Book

' Создание списка адресов УДО+НОРПиКО
Sub list_creation_1()
'
Dim Список_получателей_Range_str As String
Dim Список_получателей_Range_Row, Список_получателей_Range_Column As Byte

  ' Находим ячейку (например G41), в которой записано значение In_К_пор
  Список_получателей_Range_str = RangeByValue(ThisWorkbook.Name, "Addr.Book", "Список получателей:", 100, 100)
  Список_получателей_Range_Row = Workbooks(ThisWorkbook.Name).Sheets("Addr.Book").Range(Список_получателей_Range_str).Row
  Список_получателей_Range_Column = Workbooks(ThisWorkbook.Name).Sheets("Addr.Book").Range(Список_получателей_Range_str).Column

  '
  ThisWorkbook.Sheets("Addr.Book").Cells(Список_получателей_Range_Row, Список_получателей_Range_Column + 1).Value = getFromAddrBook("УДО2,УДО3,УДО4,УДО5,НОРПиКО1,НОРПиКО2,НОРПиКО3,НОРПиКО4,НОРПиКО5", 2)
  ThisWorkbook.Sheets("Addr.Book").Cells(Список_получателей_Range_Row, Список_получателей_Range_Column + 1).Copy

End Sub

' Создание списка адресов УДО+НОРПиКО+МРК+ПМ
Sub list_creation_2()
'
Dim Список_получателей_Range_str As String
Dim Список_получателей_Range_Row, Список_получателей_Range_Column As Byte

  ' Находим ячейку (например G41), в которой записано значение In_К_пор
  Список_получателей_Range_str = RangeByValue(ThisWorkbook.Name, "Addr.Book", "Список получателей:", 100, 100)
  Список_получателей_Range_Row = Workbooks(ThisWorkbook.Name).Sheets("Addr.Book").Range(Список_получателей_Range_str).Row
  Список_получателей_Range_Column = Workbooks(ThisWorkbook.Name).Sheets("Addr.Book").Range(Список_получателей_Range_str).Column

  '
  ThisWorkbook.Sheets("Addr.Book").Cells(Список_получателей_Range_Row, Список_получателей_Range_Column + 1).Value = getFromAddrBook("УДО2,УДО3,УДО4,УДО5,НОРПиКО1,НОРПиКО2,НОРПиКО3,НОРПиКО4,НОРПиКО5,ПМ,МРК1,МРК2,МРК3,МРК4,МРК5", 2)
  ThisWorkbook.Sheets("Addr.Book").Cells(Список_получателей_Range_Row, Список_получателей_Range_Column + 1).Copy

End Sub

' Создание списка адресов УДО+НОРПиКО+МРК+ПМ+ОКП
Sub list_creation_3()
'
Dim Список_получателей_Range_str As String
Dim Список_получателей_Range_Row, Список_получателей_Range_Column As Byte

  ' Находим ячейку (например G41), в которой записано значение In_К_пор
  Список_получателей_Range_str = RangeByValue(ThisWorkbook.Name, "Addr.Book", "Список получателей:", 100, 100)
  Список_получателей_Range_Row = Workbooks(ThisWorkbook.Name).Sheets("Addr.Book").Range(Список_получателей_Range_str).Row
  Список_получателей_Range_Column = Workbooks(ThisWorkbook.Name).Sheets("Addr.Book").Range(Список_получателей_Range_str).Column

  '
  ThisWorkbook.Sheets("Addr.Book").Cells(Список_получателей_Range_Row, Список_получателей_Range_Column + 1).Value = getFromAddrBook("УДО2,УДО3,УДО4,УДО5,НОРПиКО1,НОРПиКО2,НОРПиКО3,НОРПиКО4,НОРПиКО5,ПМ,МРК1,МРК2,МРК3,МРК4,МРК5,НОКП,РРКК,МПП", 2)
  ThisWorkbook.Sheets("Addr.Book").Cells(Список_получателей_Range_Row, Список_получателей_Range_Column + 1).Copy

End Sub

' Создание списка адресов ОКП
Sub list_creation_ОКП()
'
Dim Список_получателей_Range_str As String
Dim Список_получателей_Range_Row, Список_получателей_Range_Column As Byte

  ' Находим ячейку (например G41), в которой записано значение In_К_пор
  Список_получателей_Range_str = RangeByValue(ThisWorkbook.Name, "Addr.Book", "Список получателей:", 100, 100)
  Список_получателей_Range_Row = Workbooks(ThisWorkbook.Name).Sheets("Addr.Book").Range(Список_получателей_Range_str).Row
  Список_получателей_Range_Column = Workbooks(ThisWorkbook.Name).Sheets("Addr.Book").Range(Список_получателей_Range_str).Column

  '
  ThisWorkbook.Sheets("Addr.Book").Cells(Список_получателей_Range_Row, Список_получателей_Range_Column + 1).Value = getFromAddrBook("НОКП,РРКК,МПП", 2)
  ThisWorkbook.Sheets("Addr.Book").Cells(Список_получателей_Range_Row, Список_получателей_Range_Column + 1).Copy

End Sub

' Создание списка адресов ИЦ
Sub list_creation_ИЦ()
'
Dim Список_получателей_Range_str As String
Dim Список_получателей_Range_Row, Список_получателей_Range_Column As Byte

  ' Находим ячейку (например G41), в которой записано значение In_К_пор
  Список_получателей_Range_str = RangeByValue(ThisWorkbook.Name, "Addr.Book", "Список получателей:", 100, 100)
  Список_получателей_Range_Row = Workbooks(ThisWorkbook.Name).Sheets("Addr.Book").Range(Список_получателей_Range_str).Row
  Список_получателей_Range_Column = Workbooks(ThisWorkbook.Name).Sheets("Addr.Book").Range(Список_получателей_Range_str).Column

  '
  ThisWorkbook.Sheets("Addr.Book").Cells(Список_получателей_Range_Row, Список_получателей_Range_Column + 1).Value = getFromAddrBook("РРИЦ,РИЦ,СотрИЦ", 2)
  ThisWorkbook.Sheets("Addr.Book").Cells(Список_получателей_Range_Row, Список_получателей_Range_Column + 1).Copy

End Sub

' Создание списка адресов Весь РБ
Sub list_creation_Весь_РБ()
'
Dim Список_получателей_Range_str As String
Dim Список_получателей_Range_Row, Список_получателей_Range_Column As Byte

  ' Находим ячейку (например G41), в которой записано значение In_К_пор
  Список_получателей_Range_str = RangeByValue(ThisWorkbook.Name, "Addr.Book", "Список получателей:", 100, 100)
  Список_получателей_Range_Row = Workbooks(ThisWorkbook.Name).Sheets("Addr.Book").Range(Список_получателей_Range_str).Row
  Список_получателей_Range_Column = Workbooks(ThisWorkbook.Name).Sheets("Addr.Book").Range(Список_получателей_Range_str).Column
  '
  ThisWorkbook.Sheets("Addr.Book").Cells(Список_получателей_Range_Row, Список_получателей_Range_Column + 1).Value = getFromAddrBook("УДО2,УДО3,УДО4,УДО5,НОРПиКО1,НОРПиКО2,НОРПиКО3,НОРПиКО4,НОРПиКО5,ПМ,МРК1,МРК2,МРК3,МРК4,МРК5,НОКП,РРКК,МПП,РРИЦ,РИЦ,СотрИЦ", 2)
  ThisWorkbook.Sheets("Addr.Book").Cells(Список_получателей_Range_Row, Список_получателей_Range_Column + 1).Copy

End Sub

' Адресная книга: Нумеровать список
Sub createNumberingAddrBook()
Dim НомерСтроки_Адресная_книга, rowCount As Integer

  НомерСтроки_Адресная_книга = ThisWorkbook.Sheets("Addr.Book").Range(RangeByValue(ThisWorkbook.Name, "Addr.Book", "Адресная книга", 100, 100)).Row
  rowCount = 1
  Do While ThisWorkbook.Sheets("Addr.Book").Cells(НомерСтроки_Адресная_книга + 4 + rowCount, 2).Value <> ""
    
    ThisWorkbook.Sheets("Addr.Book").Cells(НомерСтроки_Адресная_книга + 4 + rowCount, 1).Value = rowCount
    
    ' Следующая строка
    rowCount = rowCount + 1
  
  Loop

End Sub

' Создание списка Кураторы РГС
Sub list_creation_Кураторы_РГС()
'
Dim Список_получателей_Range_str As String
Dim Список_получателей_Range_Row, Список_получателей_Range_Column As Byte

  ' Находим ячейку (например G41), в которой записано значение In_К_пор
  Список_получателей_Range_str = RangeByValue(ThisWorkbook.Name, "Addr.Book", "Список получателей:", 100, 100)
  Список_получателей_Range_Row = Workbooks(ThisWorkbook.Name).Sheets("Addr.Book").Range(Список_получателей_Range_str).Row
  Список_получателей_Range_Column = Workbooks(ThisWorkbook.Name).Sheets("Addr.Book").Range(Список_получателей_Range_str).Column

  '
  ThisWorkbook.Sheets("Addr.Book").Cells(Список_получателей_Range_Row, Список_получателей_Range_Column + 1).Value = getFromAddrBook("КРГС", 2)
  ThisWorkbook.Sheets("Addr.Book").Cells(Список_получателей_Range_Row, Список_получателей_Range_Column + 1).Copy

End Sub


' Копировать адрес сотрудника в буфер
Sub Копировать_адрес_сотрудника()
      
      
  ' Копируем Хэштег в буффер обмена
  ' ThisWorkbook.Sheets("To-Do").Cells(ActiveCell.Row, 5).Copy

  ' Находим ячейку (например G41), в которой записано значение In_К_пор
  Список_получателей_Range_str = RangeByValue(ThisWorkbook.Name, "Addr.Book", "Список получателей:", 100, 100)
  Список_получателей_Range_Row = Workbooks(ThisWorkbook.Name).Sheets("Addr.Book").Range(Список_получателей_Range_str).Row
  Список_получателей_Range_Column = Workbooks(ThisWorkbook.Name).Sheets("Addr.Book").Range(Список_получателей_Range_str).Column
  '
  ThisWorkbook.Sheets("Addr.Book").Cells(Список_получателей_Range_Row, Список_получателей_Range_Column + 1).Value = ThisWorkbook.Sheets("Addr.Book").Cells(ActiveCell.Row, 10).Value
  ThisWorkbook.Sheets("Addr.Book").Cells(Список_получателей_Range_Row, Список_получателей_Range_Column + 1).Copy


End Sub

