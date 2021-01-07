Sub stock_market()

' Declare worksheet variable
Dim ws As Worksheet

' Iterate through all worksheets
For Each ws In Worksheets

    ' Counts the number of rows
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
 
    ' Loop through each row
    For i = 2 To lastrow
        ' Copy the ticker name
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(i, 9).Value = ws.Cells(i, 1).Value
        
        ' Calculate the Yearly Change in stock value
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(i, 10).Value = ws.Cells(i, 6).Value - ws.Cells(i, 3).Value
                
            ' Adjust the Conditional Formatting
            If ws.Cells(i, 10).Value <= 0 Then
                ws.Cells(i, 10).Interior.ColorIndex = 3
            Else
               ws.Cells(i, 10).Interior.ColorIndex = 4
            End If
        
        ' Calculate the Percent Change in the Stock Value
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(i, 11).Value = (ws.Cells(i, 6).Value - ws.Cells(i, 3).Value) / ws.Cells(i, 3).Value
        ws.Cells(i, 11).NumberFormat = "0.00%"
        
        ' Calculate the Total Stock Volume
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(i, 12).Value = ws.Cells(i, 7).Value '* ws.Cells(i, 10).Value
    Next i
    
    ' Set up the Bonus Material
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    
    ' Find the ticker and maximum value of the percent change
    max_perc_increase = WorksheetFunction.Max(ws.Range("K:K").Value)
    ws.Range("Q2").Value = max_perc_increase
    ws.Cells(2, 17).NumberFormat = "0.00%"
    'Dim index As Double
    index_increase = WorksheetFunction.Match(max_perc_increase, ws.Range("K:K").Value, 0)
    ticker_increase = ws.Cells(index_increase, 1).Value
    ws.Range("P2").Value = ticker_increase
    
    ' Find the ticker and minimum value of the percent chage
    min_perc_decrease = WorksheetFunction.Min(ws.Range("K:K").Value)
    ws.Range("Q3").Value = min_perc_decrease
    ws.Cells(3, 17).NumberFormat = "0.00%"
    index_decrease = WorksheetFunction.Match(min_perc_decrease, ws.Range("K:K").Value, 0)
    ticker_decrease = ws.Cells(index_decrease, 1).Value
    ws.Range("P3").Value = ticker_decrease
    
    ' Find the ticker and maximum value of the total volume
    max_volume = WorksheetFunction.Max(ws.Range("L:L").Value)
    ws.Range("Q4").Value = max_volume
    index_volume = WorksheetFunction.Match(max_volume, ws.Range("L:L").Value, 0)
    ticker_volume = ws.Cells(index_volume, 1).Value
    ws.Range("P4").Value = ticker_volume
 
Next ws
 
End Sub
