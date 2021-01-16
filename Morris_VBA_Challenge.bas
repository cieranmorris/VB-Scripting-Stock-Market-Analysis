Attribute VB_Name = "Module1"
Sub stock_market_data():

    'Variable for representing ticker symbol and total stock volume in table
    Dim ticker As String
    
    Dim total_stock_volume As Double
    
    Dim summary_table_row As Double
    
    
    Dim lastrow As Double
    
    Dim opening_price_row As Double
    
    
    
    'specify lastrow command and for loop through ticker symbol column
    
    
    
    total_stock_volume = 0
    
    Dim opening_price As Double
    Dim closing_price As Double
    Dim percent_change As Double
    Dim yearly_change As Double
    
    For Each ws In Worksheets
            ws.Activate
            
            lastrow = Cells(Rows.Count, 1).End(xlUp).Row
            
            summary_table_row = 2
            
            opening_price_row = 2
            
            percent_change_row = 2
            
            
            For i = 2 To lastrow
                
                
                'See if cell values do not match
               If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
               
                        ticker = Cells(i, 1).Value
                
                        total_stock_volume = Cells(i, 7).Value + total_stock_volume
                
                        opening_price = Cells(opening_price_row, 3).Value
                        closing_price = Cells(i, 6).Value
                        yearly_change = closing_price - opening_price
                
                
                If opening_price <> 0 Then
                    
                
                        percent_change = (yearly_change / opening_price) * 100
                        
                Else
                
                        percent_change = yearly_change * 100
                
                End If
            
                
                        'assign to summary table
                
                        Range("I" & summary_table_row).Value = ticker
                
                        Range("J" & summary_table_row).Value = yearly_change
                
                        Range("L" & summary_table_row).Value = total_stock_volume
                
                        Range("K" & summary_table_row).Value = percent_change
                
                
                If yearly_change > 0 Then
    
                        Cells(summary_table_row, 10).Interior.ColorIndex = 4
    
                ElseIf yearly_change <= 0 Then
                    
                        Cells(summary_table_row, 10).Interior.ColorIndex = 3
    
                End If
                
                        summary_table_row = summary_table_row + 1
                
                        total_stock_volume = 0
                
                        closing_price_row = i + 1
            
                Else
            
                        total_stock_volume = Cells(i, 7).Value + total_stock_volume
                
                
                End If
                
                
            
            Next i
        Next ws
        
End Sub
