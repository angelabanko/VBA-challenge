Attribute VB_Name = "Module1"
Sub stocks():

Dim ticker As String
Dim number_tickers As Integer
Dim lastrowstate As Long
Dim opening_price As Double
Dim closing_price As Double
Dim yearly_change As Double
Dim percent_change As Double
Dim total_stock_volume As Double
Dim greatest_percent_increase As Double
Dim greatest_percent_increase_ticker As String
Dim great_percent_decrease_ticker As String
Dim greatest_stock_volume_ticker As String

For Each ws In Worksheets
        ws.Activate
        
        'find the last row
        lastrowstate = ws.Cells(Rows.Count, "A").End(xlUp).Row
        
        'header for the columns on each worksheet
        ws.Cells(1, 9).Value = "ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Yearly Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        'set a value for the columns value
        
        number_tickers = 0
        ticker = ""
        yearly_change = 0
        opening_price = 0
        percent_change = 0
        total_stock_volume = 0
    
    'search through the ticker and opening price
    
    For i = 2 To lastrowstate
        ticker = Cells(i, 1).Value
        
        If opening_price = 0 Then
            opening_price = Cells(i, 3).Value
        End If
        
        total_stock_volume = total_stock_volume + Cells(i, 7).Value
        
        'find the ticker values for closing price, opening price and the yearly change
        
        If Cells(i + 1, 1).Value <> ticker Then
            
            number_tickers = number_tickers + 1
            Cells(number_tickers + 1, 9) = ticker
            
            closing_price = Cells(i, 6)
            
            yearly_change = closing_price - opening_price
            
            Cells(number_tickers + 1, 10).Value = yearly_change
            
            'coloring the box based on the value in the box
            
            If yearly_change > 0 Then
                Cells(number_tickers + 1, 10).Interior.ColorIndex = 4
            ElseIf yearly_change < 0 Then
                Cells(number_tickers + 1, 10).Interior.ColorIndex = 3
            Else
                Cells(number_tickers + 1, 10).Interior.ColorIndex = 6
            End If
            
            'getting the % change for yearly/opening
            
            If opening_price = 0 Then
                percent_change = 0
            Else
                percent_change = (yearly_change / opening_price)
            End If
            
            Cells(number_tickers + 1, 11).Value = Format(percent_change, "Percent")
            opening_price = 0
            
            Cells(number_tickers + 1, 12).Value = total_stock_volume
            
            total_stock_volume = 0
            
            End If
            
        Next i
    'assigning the text to cells
    
     ws.Cells(2, 15).Value = "Greatest % increase"
     ws.Cells(3, 15).Value = "Greatest % decrease"
     ws.Cells(4, 15).Value = "Greatest Total Volume"
     ws.Cells(1, 16).Value = "Ticker"
     ws.Cells(1, 17).Value = "Value"
     
     lastrowstate = ws.Cells(Rows.Count, "I").End(xlUp).Row
    
    'finding the greatest increase/decrease in the new columns
    
    greatest_percent_increase = Cells(2, 11).Value
    greatest_percent_increase_ticker = Cells(2, 9).Value
    greatest_percent_decrease = Cells(2, 11).Value
    greatest_percent_decrease_ticker = Cells(2, 9).Value
    greatest_stock_volume = Cells(2, 12).Value
    greatest_stock_volume_ticker = Cells(2, 9).Value
    
    
    'loop through the ticker values and find greatest increase/decrease values
    
    For i = 2 To lastrowstate
    
        If Cells(i, 11).Value > greatest_percent_increase Then
            greatest_percent_increase = Cells(i, 11).Value
            greatest_percent_increase_ticker = Cells(i, 9).Value
        End If
        
        
        If Cells(i, 11).Value < greatest_percent_decrease Then
            greatest_percent_decrease = Cells(i, 11).Value
            greatest_percent_decrease_ticker = Cells(i, 9).Value
        End If
        
        If Cells(i, 12).Value > greatest_stock_volume Then
            greatest_stock_volume = Cells(i, 12).Value
            greatest_stock_volume_ticker = Cells(i, 9).Value
        End If
        
    Next i
    
  'format the cells
    Cells(2, 16).Value = Format(greatest_percent_increase_ticker, "Percent")
    Cells(2, 17).Value = Format(greatest_percent_increase, "Percent")
    Cells(3, 16).Value = Format(greatest_percent_decrease_ticker, "Percent")
    Cells(3, 17).Value = Format(greatest_percent_decrease, "Percent")
    Cells(4, 16).Value = greatest_stock_volume_ticker
    Cells(4, 17).Value = greatest_stock_volume
        
Next ws



End Sub

