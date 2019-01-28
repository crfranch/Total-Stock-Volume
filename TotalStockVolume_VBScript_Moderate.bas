Attribute VB_Name = "Module1"
Sub Easy()
 
    ' Set an initial variable for each of the following
    Dim TotalStockVolume As Double
    TotalStockVolume = 0
    
    Dim ticker As String

    Dim Summary_Table_Row As Integer
    'Locate where the ticker will be in the summary table
    Summary_Table_Row = 2

    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "TotalStockVolume"
    
    'Loop through all ticker options
    For i = 2 To 760193
    
        'Check if the ticker symbol is still the same, if not, move to the next one..
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
        'Set the Ticker
        ticker = Cells(i, 1).Value
        
        'Set the TotalStockVolume
        TotalStockVolume = TotalStockVolume + Cells(i, 7).Value
        
        'Print the ticker in the Summary Table
        Range("I" & Summary_Table_Row).Value = ticker
        
        'Print the TotalStockVolume in the Summary Table
        Range("J" & Summary_Table_Row).Value = TotalStockVolume
        
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
