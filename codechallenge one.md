# VBA-challenge

Sub CalculateStockDataYEAR()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim tickerSymbol As String
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    Dim outputColumn As Long
    Dim outputRow As Long
    Dim maxPercentIncrease As Double
    Dim maxPercentDecrease As Double
    Dim maxTotalVolume As Double
    Dim tickerMaxPercentIncrease As String
    Dim tickerMaxPercentDecrease As String
    Dim tickerMaxTotalVolume As String
    
    Dim wsNames As Variant
    wsNames = Array("2018", "2019", "2020") ' Add or remove worksheet names as needed
    
    For Each wsName In wsNames
        ' Specify the worksheet where the stock data is located
        Set ws = ThisWorkbook.Sheets(wsName)
        
        ' Find the last row in column A
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        
        ' Set initial values for output
        outputColumn = 13 ' Start from column M for output
        outputRow = 2 ' Start from row 2 for output
        
        ' Set headers for output columns
        ws.Cells(1, outputColumn).Value = "Ticker Symbol"
        ws.Cells(1, outputColumn + 1).Value = "Yearly Change"
        ws.Cells(1, outputColumn + 2).Value = "Percent Change"
        ws.Cells(1, outputColumn + 3).Value = "Total Volume"
        
        ' Loop through each row of stock data
        For i = 2 To lastRow
            ' Check if the ticker symbol changes
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                ' Get the current ticker symbol
                tickerSymbol = ws.Cells(i, 1).Value
                
                ' Get the opening and closing prices for the year
                openingPrice = ws.Cells(i, 3).Value
                closingPrice = ws.Cells(i, 6).Value
                
                ' Calculate the yearly change and percent change
                yearlyChange = closingPrice - openingPrice
                percentChange = yearlyChange / openingPrice
                
                ' Get the total volume for the ticker symbol
                totalVolume = Application.WorksheetFunction.Sum(ws.Range("G" & outputRow & ":G" & i))
                
                ' Output the results
                ws.Cells(outputRow, outputColumn).Value = tickerSymbol
                ws.Cells(outputRow, outputColumn + 1).Value = yearlyChange
                ws.Cells(outputRow, outputColumn + 2).Value = percentChange
                ws.Cells(outputRow, outputColumn + 3).Value = totalVolume
                
                ' Format the percent change as percentage
                ws.Cells(outputRow, outputColumn + 2).NumberFormat = "0.00%"
                
                ' Color the yearly change values
                If yearlyChange < 0 Then
                    ws.Cells(outputRow, outputColumn + 1).Interior.Color = RGB(255, 0, 0) ' Red
                ElseIf yearlyChange > 0 Then
                    ws.Cells(outputRow, outputColumn + 1).Interior.Color = RGB(0, 255, 0) ' Green
                End If
                
                ' Find the ticker with the greatest % increase, % decrease, and total volume
                If percentChange > maxPercentIncrease Then
                    maxPercentIncrease = percentChange
                    tickerMaxPercentIncrease = tickerSymbol
                ElseIf percentChange < maxPercentDecrease Then
                    maxPercentDecrease = percentChange
                    tickerMaxPercentDecrease = tickerSymbol
                End If
                
                If totalVolume > maxTotalVolume Then
                    maxTotalVolume = totalVolume
                    tickerMaxTotalVolume = tickerSymbol
                End If
                
                ' Move to the next row for output
                outputRow = outputRow + 1
            End If
        Next i
        
        ' Output the ticker with the greatest % increase, % decrease, and total volume
        ws.Cells(2, outputColumn + 6).Value = "Greatest % Increase"
        ws.Cells(3, outputColumn + 6).Value = "Greatest % Decrease"
        ws.Cells(4, outputColumn + 6).Value = "Greatest Total Volume"
        
        ws.Cells(2, outputColumn + 7).Value = tickerMaxPercentIncrease
        ws.Cells(3, outputColumn + 7).Value = tickerMaxPercentDecrease
        ws.Cells(4, outputColumn + 7).Value = tickerMaxTotalVolume
        
        ws.Cells(2, outputColumn + 8).Value = maxPercentIncrease
        ws.Cells(3, outputColumn + 8).Value = maxPercentDecrease
        ws.Cells(4, outputColumn + 8).Value = maxTotalVolume
        
        ' Format the percent increase and decrease as percentage
        ws.Cells(2, outputColumn + 8).NumberFormat = "0.00%"
        ws.Cells(3, outputColumn + 8).NumberFormat = "0.00%"
        
        ' Color the values in the "Greatest % Increase" and "Greatest % Decrease" columns
        ws.Cells(2, outputColumn + 8).Interior.Color = RGB(0, 255, 0) ' Green
        ws.Cells(3, outputColumn + 8).Interior.Color = RGB(255, 0, 0) ' Red
    Next wsName
End Sub
