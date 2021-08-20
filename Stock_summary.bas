Attribute VB_Name = "Stock_summary"
Sub Stock_summary()
    
    'declace variables for summary table
    Dim total_stock_vol As LongLong
    Dim percent_change As Double
    Dim yearly_change As Double
    Dim close_val As Double
    Dim start_val As Double
    Dim current_ticker As String
    Dim next_ticker As String
    Dim last_ticker As String
    
    'declare variables for greatest table
    Dim greatest_inc As Double
    Dim greatest_dec As Double
    Dim greatest_total_vol As LongLong
    Dim dec_ticker As String
    Dim inc_ticker As String
    Dim greatest_stock_vol_ticker As String
    

    'row and column variables
    Dim last_row As Long
    Dim sum_table_row As Integer
    
    'declare worksheet variable
    Dim ws As Worksheet
    
    
    'cycle through each worksheet in the excel workbook
    For Each ws In Worksheets
    
    
        'track summary table length
        sum_table_row = 2
        
        'Find the last non-blank cell in the first column
        last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
            
            
        ' Create new headers for summary table
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        'set total stock volume to 0
        total_stock_vol = 0
    
        For I = 2 To last_row
            
            current_ticker = ws.Cells(I, 1).Value
            next_ticker = ws.Cells(I + 1, 1).Value
            last_ticker = ws.Cells(I - 1, 1).Value
            
            'add the daily volume to the total stock volume
            total_stock_vol = total_stock_vol + ws.Cells(I, 7).Value
        
            'check if current row is last row in the year for the current stock
            If current_ticker <> next_ticker Then
                
                'get final close value
                close_val = ws.Cells(I, 6).Value
                
                'calc yearly change ; calc percent change
                yearly_change = close_val - start_val
                
                'check if stock has non-zero open; percent_change returns Div 0 error if start val is 0
                If start_val = 0 Then
                    percent_change = 0
                    
                Else
                    percent_change = yearly_change / start_val
                    
                End If
                
                'update summary table values by adding current stock info
                ws.Range("I" & sum_table_row).Value = current_ticker
                ws.Range("J" & sum_table_row).Value = yearly_change
                ws.Range("K" & sum_table_row).Value = percent_change
                ws.Range("L" & sum_table_row).Value = total_stock_vol
                
                'format sum table: yearly change conditional formatting and percentage
                'red for negative percentage, green for positive, and yellow if invalid data (all zeros over time)
                If yearly_change < 0 Then
                    ws.Range("J" & sum_table_row).Interior.ColorIndex = 3
                    
                ElseIf yearly_change = 0 And percent_change = 0 And total_stock_vol = 0 Then
                    ws.Range("J" & sum_table_row).Interior.ColorIndex = 6
                    
                Else
                    ws.Range("J" & sum_table_row).Interior.ColorIndex = 4
                    
                End If
                
                ws.Range("K" & sum_table_row).NumberFormat = "0.00%"
                
                'iterate sum table
                sum_table_row = sum_table_row + 1
                
                'reset tsv
                total_stock_vol = 0
                
            
            'check if current row is the first row for a new ticker
            ElseIf current_ticker <> last_ticker Then
                
                'get open value
                start_val = ws.Cells(I, 3).Value
    
            End If
            
            Next I
                          
        'declare greatest values
        greatest_total_vol = 0
        greatest_inc = 0
        greatest_dec = 0
        
        'create greatest table by iterating through the summary table to find max/min percent change and total stock volume
        For j = 2 To sum_table_row
            
            'set total stock volume and percent change to current row
            total_stock_vol = ws.Cells(j, 12).Value
            percent_change = ws.Cells(j, 11).Value
            
            'check if current value less than minimum value
            If percent_change < greatest_dec Then
                
                'set greatest decrease value to current value and store ticker value
                greatest_dec = percent_change
                dec_ticker = ws.Cells(j, 9).Value
            
            'check if current value greater than max value
            ElseIf percent_change > greatest_inc Then
                
                'set greatest inc value to current value and store ticker value
                greatest_inc = percent_change
                inc_ticker = ws.Cells(j, 9).Value
                
            End If
            
            'check if current tsv is greater than stored greatest value
            If total_stock_vol > greatest_total_vol Then
                
                'set greatest total stock volume to current value and store ticker value
                greatest_total_vol = total_stock_vol
                greatest_stock_vol_ticker = ws.Cells(j, 9).Value
                
            End If
                   
        Next j
        
        'Input values into greatest table
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("P2").Value = inc_ticker
        ws.Range("Q2").Value = greatest_inc
        ws.Range("Q2").NumberFormat = "0.00%"
        
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("P3").Value = dec_ticker
        ws.Range("Q3").Value = greatest_dec
        ws.Range("Q3").NumberFormat = "0.00%"
          
        ws.Range("O4").Value = "Greatest Total Stock Volume"
        ws.Range("P4").Value = greatest_stock_vol_ticker
        ws.Range("Q4").Value = greatest_total_vol
        ws.Range("Q4").NumberFormat = "General"
    
    Next ws
    
End Sub

































