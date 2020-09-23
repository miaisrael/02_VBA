Sub stocks()
'Create a script that will loop through all the stocks for one year and output the following information.
    
    'Set variable to loop through all worksheets
    Dim ws As Worksheet
    
    'Loop through all worksheets
    For Each ws In Worksheets

        'Define variables
        Dim Ticker_Name As String
        Dim Yearly_Change As Double
        Dim Percent_Change As Double
        Dim Total_Volume As Double
        Dim Summary_Table_Row As Long
        Dim Lastrow As Long
        Dim Open_Price As Double
        Dim Close_Price As Double
        
        'Assign variables
        Open_Price = ws.Cells(2, 3).Value
        Percent_Change = 0
        Total_Volume = 0
        Summary_Table_Row = 2
        
        'Set headers for all worksheets
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
      
        'Determine the last row
        Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'Loop through all rows/tickers
        For i = 2 To Lastrow
        
            'Compute the total stock volume of the stock.
            Total_Volume = ws.Cells(i, 7).Value + Total_Volume
            
            'Check for same Ticker, if not
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                'Set the ticker symbol.
                Ticker_Name = ws.Cells(i, 1).Value
            
                'Determine yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
                Close_Price = ws.Cells(i, 6).Value
                Yearly_Change = Close_Price - Open_Price
            
                'The percent change from opening price at the beginning of a given year to the closing price at the end of that year. Check divide by zero first.
                If Open_Price <> 0 Then
                    
                    Percent_Change = (Yearly_Change / Open_Price) * 100
                    
                Else
                    
                    Percent_Change = 0
                
                End If
                
                'Print Ticker Name
                ws.Range("I" & Summary_Table_Row).Value = Ticker_Name
                
                'Print Yearly Change
                ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
                
                'Fill with green if positive change, red if negative change
                If (Yearly_Change > 0) Then
                
                    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                    
                ElseIf (Yearly_Change <= 0) Then
                    
                    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                
                End If
                
                'Print Percent Change
                ws.Range("K" & Summary_Table_Row).Value = (CStr(Percent_Change) & "%")
                
                'Print Total Volume
                ws.Range("L" & Summary_Table_Row).Value = Total_Volume
                 
                'Add 1 to Summary Table Row count
                Summary_Table_Row = Summary_Table_Row + 1
                
                'Reset changes
                Yearly_Change = 0
                Close_Price = 0
                Open_Price = ws.Cells((i + 1), 3).Value
                Total_Volume = 0
            
            End If
        
        Next
        
        'Make combined summary table with headers
        ws.Cells(2, 14).Value = "Greatest % Increase"
        ws.Cells(3, 14).Value = "Greatest % Decrease"
        ws.Cells(4, 14).Value = "Greatest Total Volume"
        ws.Cells(1, 15).Value = "Ticker"
        ws.Cells(1, 16).Value = "Value"
        
        'Define variables
        Dim Max_Ticker As String
        Dim Min_Ticker As String
        Dim Total_Ticker As String
        Dim Max_Value As Double
        Dim Min_Value As Double
        Dim Total_Value As Double
        
        'Assign variables
        Max_Ticker = " "
        Min_Ticker = " "
        Total_Ticker = ""
        Max_Value = ws.Cells(2, 11).Value
        Min_Value = ws.Cells(2, 11).Value
        Total_Value = ws.Cells(2, 12).Value
        
        For j = 2 To Lastrow
        
            'Collect values for combined summary table, greatest percent increase
            If (ws.Cells(j, 11).Value > Max_Value) Then
                Max_Value = ws.Cells(j, 11).Value
                Max_Ticker = ws.Cells(j, 9).Value
                
            End If
            
            'Collect values for combined summary table, greatest percent decrease
            If (ws.Cells(j, 11).Value < Min_Value) Then
                Min_Value = ws.Cells(j, 11).Value
                Min_Ticker = ws.Cells(j, 9).Value
                    
            End If
                    
            'Collect totals for combined summary table for last row, greatest total volume
            If (ws.Cells(j, 12) > Total_Value) Then
                Total_Value = ws.Cells(j, 12).Value
                Total_Ticker = ws.Cells(j, 9).Value
                    
            End If
                    
        Next
        
     'Back to original first sheet A, prints P values to combined summary table
        ws.Cells(2, 15).Value = Max_Ticker
        ws.Cells(3, 15).Value = Min_Ticker
        ws.Cells(2, 16).Value = Max_Value
        ws.Cells(2, 16).NumberFormat = "0.00%"
        ws.Cells(3, 16).Value = Min_Value
        ws.Cells(3, 16).NumberFormat = "0.00%"
        ws.Cells(4, 15).Value = Total_Ticker
        ws.Cells(4, 16).Value = Total_Value
       
   'End loop through all worksheets
   Next
   
End Sub
