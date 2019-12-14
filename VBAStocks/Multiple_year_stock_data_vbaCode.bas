Sub stockdata()
    
    'chanllenges 2 for VBA to run on every worksheet
    '--------------
    'loop through all worksheets
    Dim ws As Worksheet
    For Each ws In Worksheets
        'determine the last row of the worksheet
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        'add header for the summary table
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        'add header for bonus summary table
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        
        '---------------
        'instructions for stock data analysis
        '---------------
        'declare variables for data analysis and set the initial value for the variables
        Dim tickerName As String
        tickerName = ""
        Dim currTicker As String
        Dim nextTicker As String
        Dim openPrice As Double
        Dim closePrice As Double
        Dim yearly_change As Double
        Dim percent_change As Double
        Dim volume As Double
        
        volume = 0
        Dim i As Long
        
        'declare the summary table row and set the initial value
        Dim summary_table_row As Integer
        summary_table_row = 2
        

        
        '--------------------
        'calculate yearly price change for the stocks
        '--------------------
        'set initial open price for the first stock
        openPrice = ws.Cells(2, 3).Value
        'loop through all the stocks
        For i = 2 To lastRow
        'set initial value for current ticker and next ticker
        currTicker = ws.Cells(i, 1).Value
        nextTicker = ws.Cells(i + 1, 1).Value
        'add to tiakcer total value
        volume = volume + ws.Cells(i, 7).Value
            If currTicker <> nextTicker Then
            'MsgBox ("Find a new ticker!")
            'insert the ticker name to summary table
            tickerName = ws.Cells(i, 1).Value
            ws.Cells(summary_table_row, 9).Value = tickerName
            'print total volume to summary table
            ws.Cells(summary_table_row, 12).Value = volume
            'reset total volume for next tricker
            volume = 0
            'set the value for close price
            closePrice = ws.Cells(i, 6).Value
            
            'calculate yearly price change and write it to summary table
            yearly_change = closePrice - openPrice
            ws.Cells(summary_table_row, 10).Value = yearly_change
            '--------------------
            'calculate the percent change
            '--------------------
                If openPrice = 0 Then
                percent_change = 0
                Else
                percent_change = yearly_change / openPrice
                'adjust the format and print the percent change data to summary table
                ws.Cells(summary_table_row, 11).NumberFormat = "0.00%"
                ws.Cells(summary_table_row, 11).Value = percent_change
                End If
                '-----------------------
                'conditional formating for yearly price change, highlight positive to color green and negative to red
                '----------------------
                If ws.Cells(summary_table_row, 10).Value >= 0 Then
                    ws.Cells(summary_table_row, 10).Interior.ColorIndex = 4
                    Else
                    ws.Cells(summary_table_row, 10).Interior.ColorIndex = 3
                    End If
                
            'move to next ticker
            currTicker = nextTicker
            summary_table_row = summary_table_row + 1
            'set stock price to 0
            yearly_change = 0
            'set the open price to next stock
            openPrice = ws.Cells(i + 1, 3).Value
           
            End If
           
        Next i
        
        '------------------------
        'challenges 1 to calculate the greatest increase %, greatest decrease % and greatest total volume
        '------------------------
        'loop through the percent change
        For i = 2 To lastRow
        'use max and min function to get the greatest values and pring to summary table
         
            If ws.Cells(i, 11).Value = WorksheetFunction.Max(ws.Range("K:K")) Then
                ws.Cells(2, 17).Value = ws.Cells(i, 11).Value
                ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
                ws.Cells(2, 17).NumberFormat = "0.00%"
                ElseIf ws.Cells(i, 11).Value = WorksheetFunction.Min(ws.Range("K:K")) Then
                ws.Cells(3, 17).Value = ws.Cells(i, 11).Value
                ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
                ws.Cells(3, 17).NumberFormat = "0.00%"
                ElseIf ws.Cells(i, 12).Value = WorksheetFunction.Max(ws.Range("L:L")) Then
                ws.Cells(4, 17).Value = ws.Cells(i, 12).Value
                ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
            
            End If
            
        Next i
        
    'MsgBox ws.Name
    Next ws
    
End Sub