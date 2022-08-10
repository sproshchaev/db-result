Attribute VB_Name = "Module_Charts"
' *** Лист Charts (Графики) ***

' *** Глобальные переменные ***
' Public numStr_Лист8 As Integer


' ***                       ***

' Создать график по данным из таблицы
Sub Создать_график()
Attribute Создать_график.VB_ProcData.VB_Invoke_Func = " \n14"

    ' Выбор диапазона с данныйми по Y Range("C8:I9").Select
    ThisWorkbook.Sheets("Charts").Range("C9:I9").Select
    
    ' Добавление графика
    ActiveSheet.Shapes.AddChart2(332, xlLineMarkers, 1000, 150).Select
    
    ' ActiveChart.SetSourceData Source:=Range("Графики!$C$8:$I$9")
    ActiveChart.SetSourceData Source:=ThisWorkbook.Sheets("Charts").Range("Charts!$C$8:$I$9")
    
    ActiveChart.ChartTitle.Select
    ActiveChart.Axes(xlValue).MajorGridlines.Select
    
    ' Надпись наименования графика
    ActiveChart.ChartTitle.Text = "Выдача ПК"
    ActiveChart.ChartTitle.Format.TextFrame2.TextRange.Characters.Text = "Выдача ПК"
    
    ' Наименование первого ряда
    ActiveChart.FullSeriesCollection(1).Name = "=""Ряд_1"""
    
    ' Добавление второго ряда
    ActiveChart.PlotArea.Select
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.FullSeriesCollection(2).Name = "=""Ряд_2"""
    ActiveChart.FullSeriesCollection(2).Values = "=Charts!$C$10:$I$10"
    
    ' Добавление третьего ряда - индекс ряда нужно увеличить FullSeriesCollection(2, затем 3 и т.д.)
    ActiveChart.PlotArea.Select
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.FullSeriesCollection(3).Name = "=""Ряд_3"""
    ActiveChart.FullSeriesCollection(3).Values = "=Charts!$C$11:$I$11"
    
    ' Добавление легенды
    ' ActiveSheet.ChartObjects("Диаграмма 23").Activate
    ActiveChart.SetElement (msoElementLegendRight)
    ' ActiveSheet.ChartObjects("Диаграмма 23").Activate
    ' ActiveChart.Legend.Select
    ' ActiveChart.Legend.LegendEntries(1).Select
    ' Application.CommandBars("Format Object").Visible = False

    
End Sub
Sub Макрос4()
Attribute Макрос4.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Макрос4 Макрос
'

'
    ActiveSheet.ChartObjects("Диаграмма 9").Activate
    ActiveChart.PlotArea.Select
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.FullSeriesCollection(2).Name = "=""Ряд2"""
    ActiveChart.FullSeriesCollection(2).Values = "=Charts!$C$10:$I$10"
End Sub
Sub Макрос5()
Attribute Макрос5.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Макрос5 Макрос
'

'
    ActiveSheet.ChartObjects("Диаграмма 23").Activate
    ActiveChart.SetElement (msoElementLegendRight)
    ActiveSheet.ChartObjects("Диаграмма 23").Activate
    ActiveChart.Legend.Select
    ActiveChart.Legend.LegendEntries(1).Select
    Application.CommandBars("Format Object").Visible = False
End Sub
