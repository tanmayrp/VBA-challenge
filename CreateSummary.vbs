Sub CreateSummary()

    Dim TickerSymbol As String
    Dim YearlyChangeBeginning As Double
    Dim YearlyChangeEnding As Double
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim TotalStockVolume As Double
    Dim SummaryTableRow As Integer
    
    
    'Variables For Bonus
    Dim GreatestPercentIncTicker As String
    Dim GreatestPercentIncrease As Double
    GreatestPercentIncrease = 0
    
    Dim GreatestPercentDecTicker As String
    Dim GreatestPercentDecrease As Double
    GreatestPercentDecrease = 0
    
    Dim GreatestTotalVolTicker As String
    Dim GreatestTotalVolume As Double
    GreatestTotalVolume = 0
    
    
    'Bonus
    For Each ws In Worksheets
    
        YearlyChangeBeginning = 0
        YearlyChangeEnding = 0
        YearlyChange = 0
        PercentChange = 0
        TotalStockVolume = 0
        SummaryTableRow = 2
        
        GreatestPercentIncrease = 0
        GreatestPercentDecrease = 0
        GreatestTotalVolume = 0
        
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"

        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        For i = 2 To lastRow
    
            'The current cell is different than the previous cell
            'We're at the beginning row of a given ticker - calculate Yearly Change Beginning
            If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                Ticker = ws.Cells(i, 1).Value
                YearlyChangeBeginning = ws.Cells(i, 3).Value
            
            End If
            
            'The current cell is different than the next cell
            'We're at the last row of a given ticker
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                Ticker = ws.Cells(i, 1).Value
                
                'Calculate summarized data
                YearlyChangeEnding = ws.Cells(i, 6).Value
                YearlyChange = YearlyChangeEnding - YearlyChangeBeginning
                
                'Ensure YearlyChangeBeginning is not 0, so we aren't dividing by 0 causing
                'an Overflow exception to be thrown
                If YearlyChangeBeginning = 0 Then
                    PercentChange = 0
                Else
                    PercentChange = YearlyChange / YearlyChangeBeginning
                End If
                
                TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
                
                'Bonus Calculations
                If PercentChange > GreatestPercentIncrease Then
                    GreatestPercentIncrease = PercentChange
                    GreatestPercentIncTicker = Ticker
                End If
                
                If PercentChange < GreatestPercentDecrease Then
                    GreatestPercentDecrease = PercentChange
                    GreatestPercentDecTicker = Ticker
                End If
                
                If TotalStockVolume > GreatestTotalVolume Then
                    GreatestTotalVolume = TotalStockVolume
                    GreatestTotalVolTicker = Ticker
                End If
                
                'Print Values
                ws.Range("I" & SummaryTableRow).Value = Ticker
                ws.Range("J" & SummaryTableRow).Value = YearlyChange
                ws.Range("K" & SummaryTableRow).Value = PercentChange
                ws.Range("K" & SummaryTableRow).NumberFormat = "0.00%"
                ws.Range("L" & SummaryTableRow).Value = TotalStockVolume
                
                'Format Printed Values
                If YearlyChange < 0 Then
                    ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 3
                Else
                    ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 4
                End If
                
                'set variables back to default
                SummaryTableRow = SummaryTableRow + 1
                TotalStockVolume = 0
                YearlyChangeBeginning = 0
                YearlyChangeEnding = 0
                PercentChange = 0
                
                
                
            Else
                TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
               
            End If
        Next i

        'Print Bonus
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        ws.Cells(2, 16).Value = GreatestPercentIncTicker
        ws.Cells(2, 17).Value = GreatestPercentIncrease
        ws.Cells(2, 17).NumberFormat = "0.00%"
        
        ws.Cells(3, 16).Value = GreatestPercentDecTicker
        ws.Cells(3, 17).Value = GreatestPercentDecrease
        ws.Cells(3, 17).NumberFormat = "0.00%"
        
        ws.Cells(4, 16).Value = GreatestTotalVolTicker
        ws.Cells(4, 17).Value = GreatestTotalVolume
        
        
                
    Next ws
    
   
    
End Sub

