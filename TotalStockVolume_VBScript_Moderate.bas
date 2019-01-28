Sub Moderate()

 
    ' Set an initial variable for each of the following
    Dim TotalStockVolume As Double
    TotalStockVolume = 0
    
    Dim ticker As String
    
    Dim year_open As Double
    Dim year_close As Double

    Dim Summary_Table_Row As Integer
    'Locate where the ticker will be in the summary table
    Summary_Table_Row = 2

    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "TotalStockVolume"
    
    'Loop through all ticker options
    For i = 2 To 760193
    
        'Use year open to find yearly change, starting at 0..
        If year_open = 0 Then
            
            'Identify year open values to set up next condition..
            year_open = Cells(i, 3).Value
            
        End If
    
        'Check if the ticker symbol is still the same, if not, move to the next one..
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value And Cells(i - 1, 1) = Cells(i, 1) Then
        
        'Set the Ticker, identifying the next one as the symbol changes
        ticker = Cells(i, 1).Value
        
        'Identify Year_close values
         year_close = Cells(i, 6).Value
         
        'Calculate Yearly Change
        YearlyChange = Cells(i, 3).Value - Cells(i, 6).Value
        
        'Calculate Percent Change
        PercentChange = (Cells(i, 6).Value - Cells(i, 3).Value) / Cells(i, 3).Value
        
        'Set the TotalStockVolume
        TotalStockVolume = TotalStockVolume + Cells(i, 7).Value
        
        'Print the Ticker in the Summary Table
        Range("I" & Summary_Table_Row).Value = ticker
        
        'Pring the Yearly Change in the Summary Table
        Range("J" & Summary_Table_Row).Value = YearlyChange
        
        'Print the TotalStockVolume in the Summary Table
        Range("K" & Summary_Table_Row).Value = PercentChange
        
        'Print the TotalStockVolume in the Summary Table
        Range("L" & Summary_Table_Row).Value = TotalStockVolume
        
        'Ensure that the summary table adds a new row...
        Summary_Table_Row = Summary_Table_Row + 1
        
        'Reset the Ticker
        TotalStockVolume = 0
        
        ' If the cell immediately following a row is the same brand...
    Else

      'Change the Ticker Symbol
      TotalStockVolume = TotalStockVolume + Cells(i, 7).Value
        
        End If
    Next i
    
End Sub

