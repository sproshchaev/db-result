Attribute VB_Name = "Module_Test"
' ***           Лист Test             ***
' *** Тестирование процедур и функций ***

' 1. Тестирование функции getDataFrom_BASE_Workbook
Sub test_getDataFrom_BASE_Workbook()
  
  ThisWorkbook.Sheets("Test").Range("C5").Value = ""
  ThisWorkbook.Sheets("Test").Range("C5").Value = getDataFrom_BASE_Workbook(ThisWorkbook.Sheets("Test").Range("C6").Value, _
                                                                                ThisWorkbook.Sheets("Test").Range("E6").Value, _
                                                                                  ThisWorkbook.Sheets("Test").Range("G6").Value, _
                                                                                    ThisWorkbook.Sheets("Test").Range("I6").Value, _
                                                                                      ThisWorkbook.Sheets("Test").Range("K6").Value, _
                                                                                        1)
End Sub


' 2. Тестирование функции Product_Name_to_Product_Code(In_Product_Name) As String
Sub test_Product_Name_to_Product_Code()
  
  ThisWorkbook.Sheets("Test").Range("C9").Value = ""
  ThisWorkbook.Sheets("Test").Range("C9").Value = Product_Name_to_Product_Code(ThisWorkbook.Sheets("Test").Range("C10").Value)
  
End Sub

' 3. Тестирование функции ДеКодировщик короткий код в наимнования продукта Function Product_Code_to_Product_Name(In_Product_Code) As String
Sub test_Product_Code_to_Product_Name()
  
  ThisWorkbook.Sheets("Test").Range("C13").Value = ""
  ThisWorkbook.Sheets("Test").Range("C13").Value = Product_Code_to_Product_Name(ThisWorkbook.Sheets("Test").Range("C14").Value)
  
End Sub

' 4. Тестирование функции getDataFrom_BASE_Workbook2
Sub test_getDataFrom_BASE_Workbook2()
  
  ThisWorkbook.Sheets("Test").Range("C17").Value = ""
  ThisWorkbook.Sheets("Test").Range("C17").Value = getDataFrom_BASE_Workbook2(ThisWorkbook.Sheets("Test").Range("C18").Value, _
                                                                                ThisWorkbook.Sheets("Test").Range("E18").Value, _
                                                                                  ThisWorkbook.Sheets("Test").Range("G18").Value, _
                                                                                    ThisWorkbook.Sheets("Test").Range("I18").Value, _
                                                                                      ThisWorkbook.Sheets("Test").Range("K18").Value, _
                                                                                        1)
End Sub


' 5. Тестирование функции Function Факт_Q_на_дату(In_OfficeNumber, In_Product_Code, In_Date) As Double
Sub test_Факт_Q_на_дату()
  
  ThisWorkbook.Sheets("Test").Range("C21").Value = ""
  ThisWorkbook.Sheets("Test").Range("C21").Value = Факт_Q_на_дату(ThisWorkbook.Sheets("Test").Range("C22").Value, _
                                                                                ThisWorkbook.Sheets("Test").Range("E22").Value, _
                                                                                  ThisWorkbook.Sheets("Test").Range("G22").Value)
End Sub

' 6. Тестирование функции Function Первый_понедельник_от_даты(In_Date) As Date
Sub test_Первый_понедельник_от_даты()
  ThisWorkbook.Sheets("Test").Range("C25").Value = ""
  ThisWorkbook.Sheets("Test").Range("C25").Value = Первый_понедельник_от_даты(ThisWorkbook.Sheets("Test").Range("C26").Value)
End Sub

' 7. Тестирование функции Function Факт_на_дату_для_прогноза_квартала(In_Date, In_Plan, In_прогноза_квартала_проц, In_working_days_in_the_week, In_NonWorkingDays) As Double
Sub test_Факт_на_дату_для_прогноза_квартала()
  
  ThisWorkbook.Sheets("Test").Range("C29").Value = ""
  ThisWorkbook.Sheets("Test").Range("C29").Value = Факт_на_дату_для_прогноза_квартала(ThisWorkbook.Sheets("Test").Range("C30").Value, _
                                                                                        ThisWorkbook.Sheets("Test").Range("E30").Value, _
                                                                                          ThisWorkbook.Sheets("Test").Range("G30").Value, _
                                                                                            ThisWorkbook.Sheets("Test").Range("I30").Value, _
                                                                                              ThisWorkbook.Sheets("Test").Range("K30").Value)
End Sub


' 8. Прогноз квартала с учетом и без учета нерабочих дней: In_Date, In_Plan, In_Fact, In_working_days_in_the_week (5-ти/6-ти дневка), In_NonWorkingDays = 1/0 (учитывать нерабочие дни из BASE\NonWorkingDays) Function Прогноз_квартала(In_Date, In_Plan, In_Fact, In_working_days_in_the_week, In_NonWorkingDays) As Double
Sub test_Прогноз_квартала()
  
  ThisWorkbook.Sheets("Test").Range("C33").Value = ""
  ThisWorkbook.Sheets("Test").Range("C33").Value = Прогноз_квартала(ThisWorkbook.Sheets("Test").Range("C34").Value, _
                                                                                        ThisWorkbook.Sheets("Test").Range("E34").Value, _
                                                                                          ThisWorkbook.Sheets("Test").Range("G34").Value, _
                                                                                            ThisWorkbook.Sheets("Test").Range("I34").Value, _
                                                                                              ThisWorkbook.Sheets("Test").Range("K34").Value)
End Sub



' 9. Прогноз квартала с учетом и без учета нерабочих дней: In_Date, In_Plan, In_Fact, In_working_days_in_the_week (5-ти/6-ти дневка), In_NonWorkingDays = 1/0 (учитывать нерабочие дни из BASE\NonWorkingDays) Function Прогноз_квартала_проц(In_Date, In_Plan, In_Fact, In_working_days_in_the_week, In_NonWorkingDays) As Double
Sub test_Прогноз_квартала_проц()
  
  ThisWorkbook.Sheets("Test").Range("C37").Value = ""
  ThisWorkbook.Sheets("Test").Range("C37").Value = Прогноз_квартала_проц(ThisWorkbook.Sheets("Test").Range("C38").Value, _
                                                                                        ThisWorkbook.Sheets("Test").Range("E38").Value, _
                                                                                          ThisWorkbook.Sheets("Test").Range("G38").Value, _
                                                                                            ThisWorkbook.Sheets("Test").Range("I38").Value, _
                                                                                              ThisWorkbook.Sheets("Test").Range("K38").Value)
End Sub

' 10. Выставляем - Цель "На неделю:" в "M9" Sub Цель_на_неделю_Лист8()
Sub test_Цель_на_неделю_Лист8()
  
  Call Цель_на_неделю_Лист8
  
End Sub

' 11. Факт месяца по продукту из поля "Fact", из In_Date определяем только номер месяца и год
Sub test_Факт_М() ' Function Факт_М(In_Date, In_OfficeNumber, In_Product_Code) As Date
  
  ThisWorkbook.Sheets("Test").Range("C45").Value = ""
  ThisWorkbook.Sheets("Test").Range("C45").Value = Факт_М(ThisWorkbook.Sheets("Test").Range("C46").Value, _
                                                                                        ThisWorkbook.Sheets("Test").Range("E46").Value, _
                                                                                          ThisWorkbook.Sheets("Test").Range("G46").Value)
  
  
End Sub

' 12. Function Продажи_Q_за_период(In_OfficeNumber, In_Product_Code, In_DateStart, In_DateEnd) As Double
Sub test_Продажи_Q_за_период()
  
  ThisWorkbook.Sheets("Test").Range("C49").Value = ""
  ThisWorkbook.Sheets("Test").Range("C49").Value = Продажи_Q_за_период(ThisWorkbook.Sheets("Test").Range("C50").Value, _
                                                                                        ThisWorkbook.Sheets("Test").Range("E50").Value, _
                                                                                          ThisWorkbook.Sheets("Test").Range("G50").Value, _
                                                                                            ThisWorkbook.Sheets("Test").Range("I50").Value)
  
  
End Sub



