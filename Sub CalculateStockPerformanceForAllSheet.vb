Sub CalculateStockPerformanceForAllSheets()
    Dim wb As Workbook
    Set wb = Workbooks("Multiple_year_stock_data.xlsm")
    
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim ticker As String, startOpenPrice As Double, endClosePrice As Double, totalVolume As Double
    Dim volumeCol As Long, openPriceCol As Long, closePriceCol As Long
    Dim outputRow As Long
    
    ' Variables for tracking the greatest values
    Dim greatestIncrease As Double: greatestIncrease = 0
    Dim greatestDecrease As Double: greatestDecrease = 0
    Dim greatestVolume As Double: greatestVolume = 0
    Dim tickerGreatestIncrease As String
    Dim tickerGreatestDecrease As String
    Dim tickerGreatestVolume As String
    
    ' Initialize column indices based on  data structure
    volumeCol = 7 ' Adjust this based on the actual volume column
    openPriceCol = 3 ' Adjust this based on the actual open price column
    closePriceCol = 6 ' Adjust this based on the actual close price column
    
    ' Loop through each worksheet in the workbook
    For Each ws In wb.Worksheets
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        outputRow = 2 ' Start output from row 2
        
        ' Reset greatest values for each sheet
        greatestIncrease = 0
        greatestDecrease = 0
        greatestVolume = 0
        
        ' Prepare headers for the output
        With ws
            .Cells(1, 9).Value = "Ticker"
            .Cells(1, 10).Value = "Yearly Change"
            .Cells(1, 11).Value = "Percent Change"
            .Cells(1, 12).Value = "Total Stock Volume"
            ' Headers for greatest values
            .Cells(2, 14).Value = "Greatest % Increase"
            .Cells(3, 14).Value = "Greatest % Decrease"
            .Cells(4, 14).Value = "Greatest Total Volume"
        End With
        
        totalVolume = 0
        If lastRow >= 2 Then
            startOpenPrice = ws.Cells(2, openPriceCol).Value ' Initialize startOpenPrice
            For i = 2 To lastRow
                ticker = ws.Cells(i, 1).Value
                totalVolume = totalVolume + ws.Cells(i, volumeCol).Value
                
                ' Clast row in the dataset
                If i = lastRow Or ws.Cells(i + 1, 1).Value <> ticker Then
                    endClosePrice = ws.Cells(i, closePriceCol).Value
                    Dim yearlyChange As Double: yearlyChange = endClosePrice - startOpenPrice
                    Dim percentChange As Double
                    If startOpenPrice <> 0 Then
                        percentChange = (yearlyChange / startOpenPrice) * 100
                    Else
                        percentChange = 0
                    End If
                    
                    ' Update greatest values
                    If percentChange > greatestIncrease Then
                        greatestIncrease = percentChange
                        tickerGreatestIncrease = ticker
                    End If
                    If percentChange < greatestDecrease Then
                        greatestDecrease = percentChange
                        tickerGreatestDecrease = ticker
                    End If
                    If totalVolume > greatestVolume Then
                        greatestVolume = totalVolume
                        tickerGreatestVolume = ticker
                    End If
                    
                    ' Output the data for the current ticker
                    With ws
                        .Cells(outputRow, 9).Value = ticker
                        .Cells(outputRow, 10).Value = yearlyChange
                        .Cells(outputRow, 11).Value = Format(percentChange, "0.00") & "%"
                        .Cells(outputRow, 12).Value = totalVolume
                        outputRow = outputRow + 1 ' Prepare for next ticker
                    End With
                    
                    totalVolume = 0 ' Reset totalVolume for the next ticker
                    If i <> lastRow Then
                        startOpenPrice = ws.Cells(i + 1, openPriceCol).Value ' Reset startOpenPrice for the next ticker
                    End If
                End If
            Next i
        End If
        
        ' Output the greatest values at the end of processing each sheet
        With ws
            .Cells(2, 15).Value = "Ticker with Greatest % Increase"
            .Cells(2, 16).Value = tickerGreatestIncrease
            .Cells(2, 17).Value = greatestIncrease & "%"
            
            .Cells(3, 15).Value = "Ticker with Greatest % Decrease"
            .Cells(3, 16).Value = tickerGreatestDecrease
            .Cells(3, 17).Value = greatestDecrease & "%"
            
            .Cells(4, 15).Value = "Ticker with Greatest Total Volume"
            .Cells(4, 16).Value = tickerGreatestVolume
            .Cells(4, 17).Value = greatestVolume
        End With
    Next ws
End Sub