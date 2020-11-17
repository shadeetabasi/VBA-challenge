Sub StockMarketHomework():

'Loop through all sheets
For Each ws In Worksheets

    'Create column headers across workbooks
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    'Declare Variable to Hold Ticker Name
    Dim Ticker As String
    
    'Keep Track of the location for each Ticker in the summary table
    Dim TickerRow As Long
    TickerRow = 1
    
    'Declare Yearly Change Variables
    Dim OpenNum As Double
    Dim CloseNum As Double
    Dim PriceChange As Double
    Dim StockVolume As Double
    
    'Define last row
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'Loop Until Last Row
    For i = 2 To LastRow
    
        'If row is the earliest date for ticker calculate and save the opening value
        If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
             OpenNum = ws.Cells(i, 3).Value
             
             'Set Stock Volume to 0 for new ticker
             StockVolume = 0
             
        End If
        
        'Calculate Total Stock Volume
        StockVolume = ws.Cells(i, 7).Value + StockVolume
        
        'If row is the last date for a ticker, do the following steps:
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
            'Input Ticker in Column
            TickerRow = TickerRow + 1
            Ticker = ws.Cells(i, 1).Value
            ws.Cells(TickerRow, "I").Value = Ticker
    
            'Define Yearly Change Values
            CloseNum = ws.Cells(i, 6).Value
            PriceChange = CloseNum - OpenNum
    
            'Attempt to calculate Price Change
            ws.Cells(TickerRow, "J").Value = PriceChange
             
             'Calculate Percent Change
            If OpenNum <> 0 Then
                ws.Cells(TickerRow, "K").Value = PriceChange / OpenNum
                ws.Cells(TickerRow, "K").NumberFormat = "0.00%"
                
            End If
            
            'Calculate Total Stock Volume
                ws.Cells(TickerRow, "L").Value = StockVolume

       End If
        
    Next i

Next ws

End Sub


