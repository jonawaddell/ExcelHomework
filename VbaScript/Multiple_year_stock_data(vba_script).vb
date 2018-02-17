Sub Button1_Click()
'Set a variable to read all tickers
Dim tickerRow As Long
'Set a variable to print all tickers
Dim tickerSumm As Long
'Set variable to sum all values per ticker
Dim totalStockVol As Double
'Set variable to select all workshees
Dim ws As Worksheet

'----------------------------------------------------
'Run macro in each active worksheet
For Each ws In Sheets
    ws.Activate
    
    'Create header to summurize the volume for each ticker
    Cells(1, 10).Value = "Ticker"
    Cells(1, 11).Value = "Total Stock Volume"

    'Set ticker summary column to begin at row 2
    tickerSumm = 2

    'Loop through all tickers
    For tickerRow = 2 To 797711
        'check if we are still within the same stcok ticker, if not...
        If Cells(tickerRow, 1) <> Cells(tickerRow + 1, 1) Then
            'calculate the last ticker in the prompted ticker group
            totalStockVol = totalStockVol + Cells(tickerRow, 7)
            'copy and paste the ticker in the summary column
            Cells(tickerSumm, 10) = Cells(tickerRow, 1)
            'copy and paste the total volume of the prompted ticker in volume summary column
            Cells(tickerSumm, 11) = totalStockVol
            'increase the ticker summary value to begin print into the next row
            tickerSumm = tickerSumm + 1
            'reset the total to zero
            totalStockVol = 0
        Else
            'if the we are within the same ticker
            totalStockVol = totalStockVol + Cells(tickerRow, 7)
        End If
    
    Next

Next ws

End Sub
