Attribute VB_Name = "Stock_summary"
Sub Stock_summary()
    
    'declace variables for new summary table
    Dim total_stock_vol As LongLong
    Dim percent_change As Double
    Dim yearly_change As Double
    Dim close_val As Double
    Dim start_val As Double
    Dim current_ticker As String
    Dim next_ticker As String
    Dim last_ticker As String


    'row and column variables
    Dim last_row As Long
    Dim last_col As Long
    Dim sum_table_row As Integer
    
    
    'track summary table length
    sum_table_row = 2
    
    'Find the last non-blank cell in column A(1)
    last_row = Cells(Rows.Count, 1).End(xlUp).Row
    
    'Find the last non-blank cell in row 1
    last_col = Cells(1, Columns.Count).End(xlToLeft).Column
    
    
    ' Create new headers for summary table
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    
    total_stock_vol = 0
    
    For I = 2 To last_row
        
        current_ticker = Cells(I, 1).Value
        next_ticker = Cells(I + 1, 1).Value
        last_ticker = Cells(I - 1, 1).Value
        
        'do stuff universal to each case:
        total_stock_vol = total_stock_vol + Cells(I, 7).Value
    
        'check if current row is last row in the year for the stock
        If current_ticker <> next_ticker Then
            
            'get final close value
            close_val = Cells(I, 6).Value
            
            'calc yearly change ; calc percent change
            yearly_change = close_val - start_val
            percent_change = (close_val - start_val) / start_val
            
            'update summary table values
            Range("I" & sum_table_row).Value = current_ticker
            Range("J" & sum_table_row).Value = yearly_change
            Range("K" & sum_table_row).Value = percent_change
            Range("L" & sum_table_row).Value = total_stock_vol
            
            'format sum table: yearly change conditional formatting and percentage
            If yearly_change < 0 Then
            
                Range("J" & sum_table_row).Interior.ColorIndex = 3
                
            Else
            
                Range("J" & sum_table_row).Interior.ColorIndex = 4
                
            End If
            
            Range("K" & sum_table_row).NumberFormat = "0.00%"
            
            'iterate sum table
            sum_table_row = sum_table_row + 1
            
            'reset tsv
            total_stock_vol = 0
            
        
        'check if current row is the first row for a new ticker
        ElseIf current_ticker <> last_ticker Then
            
            'get open value
            start_val = Cells(I, 3).Value

        End If
        
        Next I
    
    
    
End Sub

































