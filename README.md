Sub StockAnalysis()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim ticker As String
    Dim openPrice As Double
    Dim closePrice As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    Dim summaryRow As Long
    
    ' Loop through each worksheet in the workbook
    For Each ws In ThisWorkbook.Worksheets
        ' Initialize variables
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        summaryRow = 2 ' Start summary output from row 2
        
        ' Set headers for summary output
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        ' Loop through each row of data
        For i = 2 To lastRow
            ' Check if it's a new ticker symbol
            If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
                ' Record the ticker symbol and open price
                ticker = ws.Cells(i, 1).Value
                openPrice = ws.Cells(i, 3).Value
            End If
            
            ' Accumulate the total volume for the ticker
            totalVolume = totalVolume + ws.Cells(i, 7).Value
            
            ' Check if it's the last row for the ticker
            If ws.Cells(i + 1, 1).Value <> ticker Then
                ' Record the close price
                closePrice = ws.Cells(i, 6).Value
                
                ' Calculate yearly change and percent change
                yearlyChange = closePrice - openPrice
                If openPrice <> 0 Then
                    percentChange = yearlyChange / openPrice * 100
                Else
                    percentChange = 0
                End If
                
                ' Output summary information
                ws.Cells(summaryRow, 9).Value = ticker
                ws.Cells(summaryRow, 10).Value = yearlyChange
                ws.Cells(summaryRow, 11).Value = percentChange
                ws.Cells(summaryRow, 12).Value = totalVolume
                
                ' Reset variables for the next ticker
                totalVolume = 0
                summaryRow = summaryRow + 1
            End If
        Next i
    Next ws
End Sub
