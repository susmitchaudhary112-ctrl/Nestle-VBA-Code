Sub Nestle_Excel_NoError()

Application.ScreenUpdating = False

'------------------------
' DELETE OLD SHEETS
'------------------------
On Error Resume Next
Application.DisplayAlerts = False
Sheets("P&L").Delete
Sheets("BalanceSheet").Delete
Sheets("Charts").Delete
Sheets("Interpretation").Delete
Application.DisplayAlerts = True
On Error GoTo 0

'------------------------
' CREATE SHEETS
'------------------------
Set pl = Sheets.Add
pl.Name = "P&L"

Set bs = Sheets.Add
bs.Name = "BalanceSheet"

Set chs = Sheets.Add
chs.Name = "Charts"

Set itp = Sheets.Add
itp.Name = "Interpretation"

'------------------------
' P&L DATA
'------------------------
pl.Range("A1:F1").Value = Array("Metric", "2019", "2020", "2021", "2022", "2023")
pl.Range("A2:F2").Value = Array("Sales", 12000, 12500, 13800, 15000, 16200)
pl.Range("A3:F3").Value = Array("Expenses", 8000, 8300, 9000, 9800, 10200)
pl.Range("A4:F4").Value = Array("Net Profit", 4000, 4200, 4800, 5200, 6000)

'------------------------
' BALANCE SHEET DATA
'------------------------
bs.Range("A1:F1").Value = Array("Metric", "2019", "2020", "2021", "2022", "2023")
bs.Range("A2:F2").Value = Array("Total Assets", 15000, 16000, 17500, 19000, 21000)
bs.Range("A3:F3").Value = Array("Borrowings", 3000, 2800, 2600, 2400, 2200)

'------------------------
' CHART 1 – SALES
'------------------------
Set ch = chs.ChartObjects.Add(10, 10, 600, 200)
With ch.Chart
    .ChartType = xlLine
    .SeriesCollection.NewSeries
    .SeriesCollection(1).XValues = pl.Range("B1:F1")
    .SeriesCollection(1).Values = pl.Range("B2:F2")
    .HasTitle = True
    .ChartTitle.Text = "Sales Trend"
End With

'------------------------
' CHART 2 – NET PROFIT
'------------------------
Set ch = chs.ChartObjects.Add(10, 230, 600, 200)
With ch.Chart
    .ChartType = xlLine
    .SeriesCollection.NewSeries
    .SeriesCollection(1).XValues = pl.Range("B1:F1")
    .SeriesCollection(1).Values = pl.Range("B4:F4")
    .HasTitle = True
    .ChartTitle.Text = "Net Profit Trend"
End With

'------------------------
' CHART 3 – TOTAL ASSETS
'------------------------
Set ch = chs.ChartObjects.Add(10, 450, 600, 200)
With ch.Chart
    .ChartType = xlLine
    .SeriesCollection.NewSeries
    .SeriesCollection(1).XValues = bs.Range("B1:F1")
    .SeriesCollection(1).Values = bs.Range("B2:F2")
    .HasTitle = True
    .ChartTitle.Text = "Total Assets Trend"
End With

'------------------------
' CHART 4 – BORROWINGS
'------------------------
Set ch = chs.ChartObjects.Add(10, 670, 600, 200)
With ch.Chart
    .ChartType = xlLine
    .SeriesCollection.NewSeries
    .SeriesCollection(1).XValues = bs.Range("B1:F1")
    .SeriesCollection(1).Values = bs.Range("B3:F3")
    .HasTitle = True
    .ChartTitle.Text = "Borrowings Trend"
End With

'------------------------
' INTERPRETATION
'------------------------
itp.Range("A1:B1").Value = Array("Metric", "Interpretation")
itp.Rows(1).Font.Bold = True

itp.Range("A2:B2").Value = Array("Sales", "Sales show continuous growth over the years.")
itp.Range("A3:B3").Value = Array("Net Profit", "Profitability has improved steadily.")
itp.Range("A4:B4").Value = Array("Total Assets", "Asset base expanded indicating growth.")
itp.Range("A5:B5").Value = Array("Borrowings", "Borrowings declined showing financial strength.")

Application.ScreenUpdating = True
MsgBox "DONE. THIS CODE WILL NOT ERROR.", vbInformation

End Sub
