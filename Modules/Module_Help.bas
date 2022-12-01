Attribute VB_Name = "Module_Help"
' *** Лист "Справочник" ***

' Справка InStr https://docs.microsoft.com/ru-ru/office/vba/language/reference/user-interface-help/instr-function
' InStr([ начало ], строка1_где_ищем, строка2_что_ищем, [ сравнение ])

' В общем, чтобы нарисовать любой символ:
' 1) Запускаешь Эксель, в пустой ячейке любым образом рисуешь нужный символ. Например через Character Map.
' 2) В консоли VBA даешь команду ? AscW(ActiveCell.Text) получишь код этого символа как его понимает VBA (для двойной горизонтальной линии в шрифте Arial это будет 9552)
' Потом уже его можно печатать ActiveCell.Value = String(10, ChrW(9552))
'    текстПисьма = текстПисьма + "" + Chr(13)
'    текстПисьма = текстПисьма + "" + ChrW(9484) + ChrW(9472) + ChrW(9472) + ChrW(9516) + Chr(13)
'    текстПисьма = текстПисьма + "" + ChrW(9474) + "1. " + ChrW(9474) + Chr(13)

' Function кавычки() As String => "