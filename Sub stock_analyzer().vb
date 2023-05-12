Sub stock_analyzer()

    Dim Ticker_Name As String
    Dim Increase_Percent_Ticker As String
    Dim Decrease_Percent_Ticker As String
    Dim Greatest_Volume_Ticker As String
    
    Dim Total_Volume As Double
    Total_Volume = 0
    
    Dim Opening_Value As Double
    Opening_Value = 0
    
    Dim Closing_Value As Double
    
    Dim Yearly_Change As Double
    
    Dim Percent_Change As Double
    
    Dim Greatest_Percent_Increase As Double
    Greatest_Percent_Increase = 0
    
    Dim Greatest_Percent_Decrease As Double
    Greatest_Percent_Decrease = 0
    
    Dim Greatest_Total_Volume As Double
    Greatest_Total_Volume = 0
    
    
    'Keep track of where to put the information'
    Dim Table_Summary_Row As Integer
    Table_Summary_Row = 2
    
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    
    
    For Each ws In Worksheets
    
        'Create headers for new Columns'
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Volume"
        
    
            
       'Loop through all of the rows'
        For i = 2 To lastrow
                If Opening_Value = 0 Then
                Opening_Value = ws.Cells(i, 3).Value
                
            
            End If
                
            
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                    
                    'Set ticker name'
                Ticker_Name = ws.Cells(i, 1).Value
                    
                    'Add to the total volume'
                
                Total_Volume = Total_Volume + ws.Cells(i, 7).Value
                    
                    'Collect closing value of the stock'
                Closing_Value = ws.Cells(i, 6).Value
                    
                    
                    'Print Ticker name into Results'
                ws.Range("I" & Table_Summary_Row).Value = Ticker_Name
                    
                    'Print Total volume into the results'
                ws.Range("L" & Table_Summary_Row).Value = Total_Volume
                    
                    'Get Yearly change and add it to table'
                Yearly_Change = Closing_Value - Opening_Value
                    
                ws.Range("J" & Table_Summary_Row).Value = Yearly_Change
                    
                    
                    'Get Percent change and add it to table'
                Percent_Change = Yearly_Change / Opening_Value
                
            
                    
                ws.Range("K" & Table_Summary_Row).Value = Percent_Change
                
                
                
                
                    
                Table_Summary_Row = Table_Summary_Row + 1
                    
                Total_Volume = 0
                    
                Opening_Value = 0
                
             Else
                
                Total_Volume = Total_Volume + ws.Cells(i, 7).Value
                    
             End If
             
             Next i
             
             'loop throught to find the greatest numbers'
             
             For i = 2 To lastrow
             
             
                If ws.Cells(i, 11).Value > Greatest_Percent_Increase Then
                
                Greatest_Percent_Increase = ws.Cells(i, 11).Value
                
                Increase_Percent_Ticker = ws.Cells(i, 9).Value
                
                End If
                
                
                If ws.Cells(i, 11).Value < Greatest_Percent_Decrease Then
                
                Greatest_Percent_Decrease = ws.Cells(i, 11).Value
                
                Decrease_Percent_Ticker = ws.Cells(i, 9).Value
                
                End If
                
                
                If ws.Cells(i, 12).Value > Greatest_Total_Volume Then
                
                Greatest_Total_Volume = ws.Cells(i, 12).Value
                
                Greatest_Volume_Ticker = ws.Cells(i, 9).Value
                
                End If
                
                If ws.Cells(i, 10).Value < 0 Then
                              
                ws.Cells(i, 10).Interior.ColorIndex = 3
                            
                ElseIf ws.Cells(i, 10).Value > 0 Then
                            
                ws.Cells(i, 10).Interior.ColorIndex = 4
                
                Else
                
                ws.Cells(i, 10).Interior.ColorIndex = 0
            
                End If
                    
                Next i
        
        ws.Cells(1, 17).Value = "Ticker"
        ws.Cells(1, 18).Value = "Total"
        
        ws.Cells(2, 18).Value = Greatest_Percent_Increase
        ws.Cells(2, 17).Value = Increase_Percent_Ticker
        ws.Cells(2, 16).Value = "Greatest % Increase"
            
        ws.Cells(3, 18).Value = Greatest_Percent_Decrease
        ws.Cells(3, 17).Value = Decrease_Percent_Ticker
        ws.Cells(3, 16).Value = "Greatest % Decrease"
            
        ws.Cells(4, 18).Value = Greatest_Total_Volume
        ws.Cells(4, 17).Value = Greatest_Volume_Ticker
        ws.Cells(4, 16).Value = "Greatest Total Volume"
        
       
                 
    Next ws
        
End Sub
