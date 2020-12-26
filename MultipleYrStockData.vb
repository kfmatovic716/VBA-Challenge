Sub SummaryStockAnalysis()
       
   For Each ws In Worksheets
  
        
        Dim CurrentTicker As String
        Dim RowSumm As Integer
        Dim OpenPrice_boy As Double
        Dim ClosePrice_eoy As Double
        Dim PctChg As Double
        Dim PriceDiff As Double
        Dim YearlyChg As Double
        Dim RngPct As Range
        Dim RngVol As Range
        Dim MaxPct As Double
        Dim MinPct As Double
        
       
        RowSumm = 2     'Summary row starts at row 2
        TotalVol = 0    'Records the total volume for each ticker in the summary table
        RowBegin = 2   'Open price starts at row 2
        
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row 'Looks for the last row in the dataset
        FinalRow_Summary = ws.Cells(Rows.Count, 11).End(xlUp).Row
        
        'Create header titles for each worksheet
        ws.Cells(1, 9).Value = "Tickers"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        
        'Set ColumnWidth of columns O and Q to display the entire row title & values (total volume)
        ws.Range("J:J").ColumnWidth = 12.25
        ws.Range("K:K").ColumnWidth = 13.25
        ws.Range("L:L").ColumnWidth = 16.25 
        ws.Range("O:O").ColumnWidth = 20.25
        ws.Range("Q:Q").ColumnWidth = 12.25
        
        'Row labels for the min & max table
         ws.Cells(2, 15).Value = "Greatest % Increase"
         ws.Cells(3, 15).Value = "Greatest % Decrease"
         ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        'Loop through each ticker row
        For i = 2 To LastRow

            'Add current value to total
            TotalVol = TotalVol + ws.Cells(i, 7).Value
            
            'Grab current tickers and open price
            CurrentTicker = ws.Cells(i, 1).Value
            OpenPrice_boy = ws.Cells(RowBegin, 3).Value
            
            'if next ticker is not the same as the current ticker
            If CurrentTicker <> ws.Cells(i + 1, 1).Value Then
                
                'Grab the closing price & calculate price difference, which = yearly change
                ClosePrice_eoy = ws.Cells(i, 6).Value
                YearlyChg = ClosePrice_eoy - OpenPrice_boy
                
                'If Yearly change is negative, color cell with Red otherwise Green
                If YearlyChg < 0 Then
                    ws.Cells(RowSumm, 10).Interior.Color = RGB(255, 0, 0)
                Else
                    ws.Cells(RowSumm, 10).Interior.Color = RGB(0, 255, 0)
                End If
                
                'Record the current ticker, total volume and yearly change in the summary
                ws.Cells(RowSumm, 9).Value = CurrentTicker
                ws.Cells(RowSumm, 12).Value = TotalVol
                ws.Cells(RowSumm, 10).Value = YearlyChg
                
                'Calculate Percent Change
                'Conditional for cases w/ OpenPrice_boy = 0 (invalid - div by 0)
                If OpenPrice_boy <> 0 Then
                    PctChg = Round((YearlyChg / OpenPrice_boy), 2)
                Else
                    PctChg = 0
                End If
                
                'Record the result of Percent Change in the summary
                ws.Cells(RowSumm, 11).Value = PctChg
                
                'Move to the next row in summary
                RowSumm = RowSumm + 1
                
                'Reset all variables: total vol, tickers, PctChg, open/close price & price difference
                TotalVol = 0
                CurrentTicker = ws.Cells(i + 1, 1).Value
                OpenPrice_boy = 0
                ClosePrice_eoy = 0
                PriceDiff = 0
                PctChg = 0
                RowBegin = i + 1
            End If
        Next i
    
        
       'Set the range where to find the min/max values & total volume
        Set RngPct = ws.Range("K:K")
        Set RngVol = ws.Range("L:L")
       
        'Find min and max values using the given ranges
        MaxPct = Application.WorksheetFunction.Max(RngPct)
        MinPct = Application.WorksheetFunction.Min(RngPct)
        MaxVol = Application.WorksheetFunction.Max(RngVol)
        
        'Record the min and max values
        ws.Range("Q2").Value = MaxPct & "%"
        ws.Range("Q3").Value = MinPct & "%"
        ws.Range("Q4").Value = MaxVol
        
        'Lookup row position of the min/max values
        MaxPos = Application.WorksheetFunction.Match(MaxPct, RngPct, 0)
        MinPos = Application.WorksheetFunction.Match(MinPct, RngPct, 0)
        MaxVolPos = Application.WorksheetFunction.Match(MaxVol, RngVol, 0)
        
        'After finding the row position, specify the column number to match it to and
        'record the ticker
        ws.Range("P2") = ws.Cells(MaxPos, 9)
        ws.Range("P3") = ws.Cells(MinPos, 9)
        ws.Range("P4") = ws.Cells(MaxVolPos, 9)
       
    Next ws
End Sub





