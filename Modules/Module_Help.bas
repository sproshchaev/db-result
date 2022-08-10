Attribute VB_Name = "Module_Help"
' *** Ћист "—правочник" ***

' —правка InStr https://docs.microsoft.com/ru-ru/office/vba/language/reference/user-interface-help/instr-function
' InStr([ начало ], строка1_где_ищем, строка2_что_ищем, [ сравнение ])

' ¬ общем, чтобы нарисовать любой символ:
' 1) «апускаешь Ёксель, в пустой €чейке любым образом рисуешь нужный символ. Ќапример через Character Map.
' 2) ¬ консоли VBA даешь команду ? AscW(ActiveCell.Text) получишь код этого символа как его понимает VBA (дл€ двойной горизонтальной линии в шрифте Arial это будет 9552)
' ѕотом уже его можно печатать ActiveCell.Value = String(10, ChrW(9552))
'    текстѕисьма = текстѕисьма + "" + Chr(13)
'    текстѕисьма = текстѕисьма + "" + ChrW(9484) + ChrW(9472) + ChrW(9472) + ChrW(9516) + Chr(13)
'    текстѕисьма = текстѕисьма + "" + ChrW(9474) + "1. " + ChrW(9474) + Chr(13)

' Function кавычки() As String => "

