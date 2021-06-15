Sub vba_of_wallstreet():
    
    'Declaring variables
    Dim yearly_change As Double
    Dim percent_change As Double
    Dim total_stock_volume As Double
    Dim i As Long
    Dim lastrow As Long
    Dim year_open As Double
    Dim year_close As Double
    Dim row_count As Integer
    Dim sum_lastrow As Integer
    Dim j  As Range
    Dim k As Range
    Dim max As Double
    Dim min As Double
    Dim grt_total_vol As Double
    Dim ws As Worksheet
    
    'Looping through all the worksheets in the workbook
    For Each ws In Worksheets
    
        'Displaying headers for summary table
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        lastrow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        row_count = 2
        year_open = ws.Range("C2").Value
          
            'For loop to print summary table
            For i = 2 To lastrow
                
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                    
                    year_close = ws.Cells(i, 6).Value
                    yearly_change = year_close - year_open
                        
                        'If condition to avoid error not divisible by 0
                        If year_open <> 0 Then
                            'Calculating percent change
                            percent_change = ((year_close - year_open) / year_open)
                        Else
                        End If
                        
                    'Calculating total stock volume for a ticker in a year
                    total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value
                    
                    'Displaying values in summary table
                    ws.Range("I" & row_count).Value = ws.Cells(i, 1).Value
                    ws.Range("J" & row_count).Value = yearly_change
                    ws.Range("K" & row_count).Value = FormatPercent(percent_change, 2)
                    ws.Range("L" & row_count).Value = total_stock_volume
                    
                    'Setting year's open price for the next ticker
                    year_open = ws.Cells(i + 1, 3).Value
                    
                    'Setting row count in summary table
                    row_count = row_count + 1
                    
                    'resetting total stock volume for next ticker
                    total_stock_volume = 0
                Else
                    total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value
                End If
            Next i
        
        'lastrow value of summary table
        sum_lastrow = ws.Cells(ws.Rows.Count, "J").End(xlUp).Row
        
            'Conditional formatting
            For i = 2 To sum_lastrow
                If ws.Cells(i, 10).Value > 0 Then
                    'Formatting cell background with green color for positive change
                    ws.Cells(i, 10).Interior.ColorIndex = 4
                Else
                    'Formatting cell background with red color for negative change
                    ws.Cells(i, 10).Interior.ColorIndex = 3
                End If
            Next i
        
        'Displaying header and title for bonus table
        ws.Range("O2").Value = "Greatest % increase"
        ws.Range("O3").Value = "Greatest % decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        
        Set j = ws.Range("K2:K" & sum_lastrow)
        Set k = ws.Range("L2:L" & sum_lastrow)
        
        'Finding greatest % increase
        max = ws.Application.WorksheetFunction.max(j)
        
        'Finding greatest % decrease
        min = ws.Application.WorksheetFunction.min(j)
        
        'Finding greatest total volume
        grt_total_vol = ws.Application.WorksheetFunction.max(k)
        
        'Displaying the values
        ws.Range("Q2").Value = FormatPercent(max)
        ws.Range("Q3").Value = FormatPercent(min)
        ws.Range("Q4").Value = grt_total_vol
        
            'For loop to find ticker name that corresponds to greatest % increase
            For i = 2 To sum_lastrow
                If ws.Cells(i, 11).Value = max Then
                    ws.Range("P2").Value = ws.Cells(i, 9).Value
                End If
            Next i
            
            'For loop to find ticker name that corresponds to greatest % decrease
            For i = 2 To sum_lastrow
                If ws.Cells(i, 11).Value = min Then
                    ws.Range("P3").Value = ws.Cells(i, 9).Value
                End If
            Next i
            
            'For loop to find stock that corresponds to greatest total volume
            For i = 2 To sum_lastrow
                If ws.Cells(i, 12).Value = grt_total_vol Then
                    ws.Range("P4").Value = ws.Cells(i, 9).Value
                End If
            Next i
        
        'Autofit summary table
        ws.Columns("I:Q").EntireColumn.AutoFit
    
    Next ws

End Sub
