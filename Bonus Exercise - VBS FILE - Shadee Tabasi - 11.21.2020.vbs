Sub BonusScript()

'Loop Through All Sheets
For Each ws In Worksheets

'Create row headers across workbooks
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    
'Create row headers across workbooks
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"

'Find Max PERCENT INCREASE , Max Ticker and Insert into Summary Table
MaxNum = WorksheetFunction.Max(ws.Columns("K"))
MaxTicker = WorksheetFunction.Match(MaxNum, ws.Columns("K"), 0)
ws.Range("P2").Value = ws.Cells(MaxTicker + 1, 9)
ws.Range("Q2").Value = MaxNum
ws.Range("Q2").NumberFormat = "0.00%"

'Find Min PERCENT DECREASE, Min Ticker and Insert into Summary Table
MinNum = WorksheetFunction.Min(ws.Columns("K"))
MinTicker = WorksheetFunction.Match(MinNum, ws.Columns("K"), 0)
ws.Range("P3").Value = ws.Cells(MinTicker + 1, 9)
ws.Range("Q3").Value = MinNum
ws.Range("Q3").NumberFormat = "0.00%"

'Find MAX STOCK VOLUME Value , Min Ticker and Insert into Summary Table
MaxStock = WorksheetFunction.Max(ws.Columns("L"))
MaxStockTicker = WorksheetFunction.Match(MaxStock, ws.Columns("L"), 0)
ws.Range("P4").Value = ws.Cells(MaxStockTicker + 1, 9)
ws.Range("Q4").Value = MaxStock

Next ws

End Sub




