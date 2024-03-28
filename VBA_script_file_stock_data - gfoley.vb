Sub Main()
    Dim ws As Worksheet
    
    For Each ws In Worksheets
        'Dim worksheet_name As String
        'worksheet_name = ws.Name
        'Debug.Print (ws.Name)
        
        'Last row variable
        Dim lastRow As Long
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'Formatting
        ws.Range("J1").Value = "Ticker"
        ws.Range("K1").Value = "Yearly Change"
        ws.Range("L1").Value = "Percent Change"
        ws.Range("M1").Value = "Total Stock Volume"
        ws.Range("J1:M1").Font.Bold = True
        ws.Range("P1:R1").Font.Bold = True
        ws.Range("J1:R1").Columns.AutoFit
        ws.Range("R1").ColumnWidth = "8"
        ws.Range("P2").Value = "Greatest % Increase"
        ws.Range("P3").Value = "Greatest % Decrease"
        ws.Range("P4").Value = "Greatest Total Volume"
        ws.Range("Q1").Value = "Ticker"
        ws.Range("R1").Value = "Value"
        ws.Range("P2:P4").Columns.AutoFit
        ws.Range("R2:R3").NumberFormat = "0.00%"
        ws.Range("K2:K" & lastRow).NumberFormat = "0.00"
        ws.Range("L2:L" & lastRow).NumberFormat = "0.00%"
                
        'Ticker cell variable
        Dim ticker As String
        'Creating variable for column J as "Ticker" to hold deduplicated ticker values
        Dim ticker_column As Range
        Set ticker_column = ws.Range("J2:J" & lastRow)
        'Setting open price and close price variables.
        Dim opening_price As Double
        Dim closing_price As Double
        'Setting variables for yearly difference from open to close, percent change from open to close
        'Creating variable for column K as "Yearly Change" and column L as "Percent Change"
        Dim yearly_diff As Double
        Dim yearly_column As Range
        Set yearly_column = ws.Range("K2:K" & lastRow)
        Dim percent_change As Double
        Dim percent_column As Range
        Set percent_column = ws.Range("L2:L" & lastRow)
        
        Dim total_volume As Double
        Dim volume_column As Range
        Set volume_column = ws.Range("M2:M" & lastRow)
        
        'For loop through each ticker value
        'If statement compares ticker values of the next row, if the current row does not match the previous row
        'If statement collects ticker value and opening price, then moves to the next row
        'If statement collects closing price when the current ticker value is the same as the previous ticker row
        
        
        j = 0   'j provides row level context for tickers, opening price, closing price
        total_volume = 0    'counter for total volume column
        
        For i = 2 To lastRow
        
            ticker = ws.Cells(i, 1).Value
            'Debug.Print (ticker)
            
            If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
                j = j + 1
                ticker_column.Cells(j, 1).Value = ticker
                opening_price = ws.Cells(i, 3).Value
                
                ElseIf ws.Cells(i, 1).Value = ws.Cells(i - 1, 1).Value Then
                closing_price = ws.Cells(i, 6).Value
                
            End If
            
            'calcuating yearly diff by subtracting closing price from opening price
            yearly_diff = closing_price - opening_price
            yearly_column.Cells(j, 1).Value = yearly_diff
            'Debug.Print (yearly_diff)
            'calculating percent_change by taking yearly diff divided by opening price, multiply by 100 to get percentage
            percent_change = yearly_diff / opening_price
            percent_column.Cells(j, 1).Value = percent_change
            'Debug.Print (percent_change)
            
            
            'If statement compares ticker values of the next row, if the current row matches the previous row
            'If statement sums or adds vol column to total volume variable, then moves to the next row
            'When the ticker values are different, the total volume counter starts over with the next value in vol column
            
            If ws.Cells(i, 1).Value = ws.Cells(i - 1, 1).Value Then
                total_volume = total_volume + ws.Cells(i, 7).Value
                volume_column.Cells(j, 1).Value = total_volume
                'total_volume counter is placed in column M matching with applicable ticker
                
                ElseIf ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
                total_volume = ws.Cells(i, 7).Value
                
            End If
            
        Next i
        
        For i = 2 To lastRow
            
            'Conditional formatting applied to Column L based on Percent Change increase or decrease
            If ws.Cells(i, 11) >= 0 And ws.Cells(i, 11) <> "" Then
                ws.Cells(i, 11).Interior.ColorIndex = 4
    
                ElseIf ws.Cells(i, 11) < 0 And ws.Cells(i, 11) <> "" Then
                ws.Cells(i, 11).Interior.ColorIndex = 3
     
            End If

        Next i
        
        'Setting variables for greatest increase, greatest decrease and greatest total volume values and tickers
        Dim increase As Double
        Dim decrease As Double
        Dim total As Double
        Dim ticker_increase As String
        Dim ticker_decrease As String
        Dim ticker_total As String
                
        increase = 0
        decrease = 0
        total = 0
            
        'For loop through Percent Change and Total Stock Volume
        For i = 2 To lastRow
                
            'If statement to review percent change increase and decrease
            'pulls out the greatest increase and greatest decrease to column Q, column R
            If ws.Cells(i, 12).Value > increase Then
                increase = ws.Cells(i, 12).Value
                ticker_increase = ws.Cells(i, 10).Value
                
            End If
                
            If ws.Cells(i, 12).Value < decrease Then
                decrease = ws.Cells(i, 12).Value
                ticker_decrease = ws.Cells(i, 10).Value
                
            End If
                
            'If statement to review total volume
            'pulls out the greatest total volume to column Q, column R
            If ws.Cells(i, 13).Value > total Then
                total = ws.Cells(i, 13).Value
                ticker_total = ws.Cells(i, 10).Value
                
            End If
            
        Next i
        
        'Populating all cell values with ticker and value variables for increase, decrease and total
        'Additional formatting applied
        ws.Cells(2, 17).Value = ticker_increase
        ws.Cells(3, 17).Value = ticker_decrease
        ws.Cells(4, 17).Value = ticker_total
        ws.Cells(2, 18).Value = increase
        ws.Cells(3, 18).Value = decrease
        ws.Cells(4, 18).Value = total

    Next ws
    
End Sub