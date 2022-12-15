Sub Ticker()

    'Loop through all Worksheets
    For Each ws In Worksheets
        
        ' Set an initial variable for holding the Stock Ticker name
        Dim Ticker_Name As String
    
        ' Set an initial variable for holding the Opening Price per Stock Ticker
        Dim Opening_Price As Double
        
        ' Set an initial variable for holding the Closing Price per Stock Ticker
        Dim Closing_Price As Double
        
        ' Set an initial variable for holding the Yearly Change per Stock Ticker
        Dim Yearly_Change As Double
        
        ' Set an initial variable for holding the Percent Change per Stock Ticker
        Dim Percent_Change As Double
        
        ' Set an initial variable for holding the Stock Volume Total per Stock Ticker
        Dim Volume_Total As Double
        Volume_Total = 0
    
        ' Keep track of the location for each Stock Ticker in the summary table
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
      
        'Create a variable for LastRow
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).row
        
        'Create column headers for the summary table
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
    
        ' Part 1: Loop through all Stock Tickers to get Unique Ticker Name and Total Stock Volume
        For i = 2 To LastRow
      
            ' Check if we are still within the same Stock Ticker, if it is not...
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
            ' Set the Stock Ticker name
            Ticker_Name = ws.Cells(i, 1).Value
          
            ' Add to the Volume Total
            Volume_Total = Volume_Total + ws.Cells(i, 7).Value
    
            ' Print the Stock Ticker name in the Summary Table
            ws.Range("I" & Summary_Table_Row).Value = Ticker_Name
    
            ' Print the Volume Total to the Summary Table
            ws.Range("L" & Summary_Table_Row).Value = Volume_Total
    
            ' Add one to the summary table row
            Summary_Table_Row = Summary_Table_Row + 1
          
            ' Reset the Volume Total
            Volume_Total = 0
    
            ' If the cell immediately following a row is the same Ticker...
            Else
    
            ' Add to the Volume Total
            Volume_Total = Volume_Total + ws.Cells(i, 7).Value
    
            End If
    
        Next i
              
        'Autofit Volume Total Column
        ws.Range("L1").EntireColumn.AutoFit
        
        'Part 2: Now loop through the data again to get Yearly Change and Percent Change
        Summary_Table_Row = 2
        For i = 2 To LastRow
      
            ' Check if we are still within the same Stock Ticker, if it is not...
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
            ' Set the Closing Price
            Closing_Price = ws.Cells(i, 6).Value
          
            'Check if Stock Ticker is the first of its range
            ElseIf ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
        
            'Set the Opening Price
            Opening_Price = ws.Cells(i, 3).Value
            
            End If
        
            ' Check if we are still within the same Stock Ticker, if it is not...
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
            'Calculate Yearly Change
            Yearly_Change = Closing_Price - Opening_Price
        
            'Calculate Percent_Change
            Percent_Change = Yearly_Change / Opening_Price
        
            'Print the Yearly Change to the Summary Table
            ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
                
                'Set conditional formatting for Yearly Change
                If Yearly_Change > 0 Then
                ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                Else
                ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                End If
        
            'Print the Percent Change to the Summary Table, and format as %
            ws.Range("K" & Summary_Table_Row).Value = FormatPercent(Percent_Change)
            
            ' Add one to the summary table row
            Summary_Table_Row = Summary_Table_Row + 1
            
            End If
      
        Next i
    
        'Autofit Yearly Change Column
        ws.Range("J1").EntireColumn.AutoFit
        
        'Autofit Percent Change Column
        ws.Range("K1").EntireColumn.AutoFit
        
    Next ws
    
End Sub

