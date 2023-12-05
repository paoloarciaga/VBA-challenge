Sub MultipleYearStockMarketData()

    Dim ws As Worksheet
    Dim ticker As String
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    Dim lastRow As Long
    Dim outputRow As Long
    
    'Set ws = ThisWorkbook.Sheets("A")
    For Each ws In Worksheets
    
    'Set header names
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Volume"
    
    'Set additional header names for ticker and value
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    
    'Set headers for greatest % increase, decrease and total volume
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    
    'Find the last row in worksheet
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    'Set variables for the first row
    openingPrice = ws.Cells(2, 3).Value
    ticker = ws.Cells(2, 1).Value
    totalVolume = 0
    outputRow = 2 'Starting row for output
    
    'Loop through each row of data
    For i = 2 To lastRow
    
    'Update total volume
    totalVolume = totalVolume + ws.Cells(i, 7).Value
    
    'Check if ticker has changed
    If ws.Cells(i, 1).Value <> ticker Then
    
    'Update closing price
    closingPrice = ws.Cells(i, 6).Value
        
    'Calculate yearly change and percent change
    yearlyChange = closingPrice - openingPrice
    percentChange = (closingPrice - openingPrice) / openingPrice
    
    'Output results for the previous ticker starting from outputRow
    ws.Cells(outputRow, 9).Value = ticker
    ws.Cells(outputRow, 10).Value = yearlyChange
    ws.Cells(outputRow, 11).Value = IIf(openingPrice = 0, 0, percentChange)
    ws.Cells(outputRow, 12).Value = totalVolume
    ws.Cells(outputRow, 11).NumberFormat = "0.00%"
    
    If ws.Cells(outputRow, 10).Value < 0 Then
    ws.Cells(outputRow, 10).Interior.Color = RGB(255, 0, 0)
    
    End If
    
    If ws.Cells(outputRow, 10).Value > 0 Then
    ws.Cells(outputRow, 10).Interior.Color = RGB(0, 255, 0)
    
    End If
        
    'OutputRow for the next set of results
    outputRow = outputRow + 1
    
    'Reset variables for the new ticker
    openingPrice = ws.Cells(i, 3).Value
    ticker = ws.Cells(i, 1).Value
    totalVolume = 0
    
    End If
    
        
    'Update closing price
    closingPrice = ws.Cells(i, 6).Value
        
    'Calculate yearly change and percent change
    yearlyChange = closingPrice - openingPrice
    percentChange = (closingPrice - openingPrice) / openingPrice * 100
    
    
    Next i
    
    'Find the last row in worksheet summary
    lastRow2 = ws.Cells(ws.Rows.Count, "I").End(xlUp).Row
    
    'Create variables for greatest percent increase, decrease, total volume
    GreatestIncrease = -Infinity
    GreatestDecrease = Infinity
    GreatestTotalVolume = -Infinity
    
    For i = 2 To lastRow2
       
    currentTicker = ws.Cells(i, "i").Value
    CurrentPercent = ws.Cells(i, "k").Value
    CurrentVolume = ws.Cells(i, "l").Value

    'Check if this percentage increase is the greatest so far
    If CurrentPercent > GreatestIncrease Then
    GreatestIncrease = CurrentPercent
    ws.Cells(2, "Q").Value = GreatestIncrease
    ws.Cells(2, "Q").NumberFormat = "0.00%"
    ws.Cells(2, "P").Value = currentTicker
        
    End If
        
    If CurrentPercent < GreatestDecrease Then
    GreatestDecrease = CurrentPercent
    ws.Cells(3, "Q").Value = GreatestDecrease
    ws.Cells(3, "Q").NumberFormat = "0.00%"
    ws.Cells(3, "P").Value = currentTicker
            
    End If
        
    If CurrentVolume > GreatestTotalVolume Then
    GreatestTotalVolume = CurrentVolume
    ws.Cells(4, "Q").Value = GreatestTotalVolume
    ws.Cells(4, "P").Value = currentTicker
    
    End If

    Next i
        
    Next

End Sub


