Sub Stock_Analyzer()

LastRow = Cells(Rows.Count, 1).End(xlUp).Row

For Each ws In Worksheets

    Dim Stock_Ticker As String
    Dim Total_Volume As Double
    Total_Volume = 0
    
    Dim Stock_Open As Double
    Dim Stock_Close As Double
    Dim Yearly_Change As Double
    Dim Percent_Change As Double
    
    Dim Table_Summary_Row As Integer
    Table_Summary_Row = 2
    
    Dim IsOpeningValue As Boolean
    IsOpeningValue = False
    'Create headers for the columns'
    
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Volume"

    Dim Increase_Percent_Ticker As String
    Dim Decrease_Percent_Ticker As String
    Dim Greatest_Volume_Ticker As String
    Dim Greatest_Percent_Increase As Double
    Greatest_Percent_Increase = 0
    Dim Greatest_Percent_Decrease As Double
    Greatest_Percent_Decrease = 0
    Dim Greatest_Total_Volume As Double
    Greatest_Total_Volume = 0
    
    'loop through to find summary column numbers'
    For i = 2 To LastRow
    
        'Grabs the opening price for a ticker'
        If IsOpeningValue = False Then
            
            Stock_Open = ws.Cells(i, 3).Value
            
            IsOpeningValue = True
            
        End If
        
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
        'Add to total volume'
        Total_Volume = Total_Volume + ws.Cells(i, 7).Value
        
        'find ticker name for summary row'
        Stock_Ticker = ws.Cells(i, 1).Value
        ws.Range("I" & Table_Summary_Row).Value = Stock_Ticker
        
        'find yearly change'
        Stock_Close = ws.Cells(i, 6).Value
        Yearly_Change = Stock_Close - Stock_Open
        ws.Range("J" & Table_Summary_Row).Value = Yearly_Change
        
        'find percent change'
        Percent_Change = Yearly_Change / Stock_Open
        ws.Range("K" & Table_Summary_Row).Value = Percent_Change
        
        'Print total volume'
        ws.Range("L" & Table_Summary_Row).Value = Total_Volume
        
        'Reset for next stock'
        IsOpeningValue = False
        Total_Volume = 0
        Table_Summary_Row = Table_Summary_Row + 1
        
        Else
        
        'Add up total volume'
        Total_Volume = Total_Volume + ws.Cells(i, 7).Value
        
        End If
        
        Next i
            
        'Format Yearly Change'
        For i = 2 To LastRow
            
            If ws.Cells(i, 10).Value < 0 Then
                    
            ws.Cells(i, 10).Interior.ColorIndex = 3
                
            Else
                
            ws.Cells(i, 10).Interior.ColorIndex = 4
                
            End If
                
        Next i
        
        'Format Percent Change'
        
        For i = 2 To LastRow
        
            ws.Cells(i, 11).NumberFormat = "0.00%"
            
            Next i
            
           For i = 2 To LastRow
           
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
        Next i
            
            'Add headers to greatest summary table'
            
            ws.Range("O2").Value = "Greatest % Increase"
            ws.Range("O3").Value = "Greatest % Decrease"
            ws.Range("O4").Value = "Greatest Total Volume"
            ws.Range("P1").Value = "Ticker"
            ws.Range("Q1").Value = "Totals"
            
            
            'Add results to summary table
            ws.Range("P2").Value = Increase_Percent_Ticker
            ws.Range("P3").Value = Decrease_Percent_Ticker
            ws.Range("P4").Value = Greatest_Volume_Ticker
            ws.Range("Q2").Value = Greatest_Percent_Increase
            ws.Range("Q3").Value = Greatest_Percent_Decrease
            ws.Range("Q4").Value = Greatest_Total_Volume
            'Add percentage formatting to summary table
            ws.Range("Q2").NumberFormat = "0.00%"
            ws.Range("Q3").NumberFormat = "0.00%"
            
    
            
        
        

Next ws

End Sub
