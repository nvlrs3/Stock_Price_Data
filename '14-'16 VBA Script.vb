Sub Stock_Price_Calculator()
    
    'Create Ticker variable
    Dim Ticker As String
     
    'Create variable for volume
    Volume = 0
    
    'Keep track of each ticker in summary table
    Summary_Table_Row = 2

    'Find last row
    lastrow = Range("A1").End(xlDown).Row

    'Set up Beginning_Price and Ending_Price
    Beginning_Price = 0
    Ending_Price = 0
                     
    'Step through each row to calculate required fields
    For i = 2 To lastrow
    
        'Ignore the zero rows in 2014
        On Error Resume Next
    
        'Save the beginning price and volume
        If Cells(i, 1).Value <> Cells(i - 1, 1).Value Then
        Beginning_Price = Cells(i, 3).Value
        Volume = Volume + Cells(i, 7).Value
             
        'Check if we are still in the same ticker symbol. For the last ticker of each group:
        ElseIf Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
        
        'Save each ticker symbol
        Ticker = Cells(i, 1).Value
                     
        'Save the Ending Price
        Ending_Price = Cells(i, 6).Value
        
        'Add to Volume
        Volume = Volume + Cells(i, 7).Value

        'Output each Ticker in the summary table
        Range("I" & Summary_Table_Row).Value = Ticker
        
        'Output the yearly price change
        Range("J" & Summary_Table_Row).Value = Ending_Price - Beginning_Price

        'Output the percentage price change
        Percent_Price_Change = (Ending_Price / Beginning_Price - 1)
        Range("K" & Summary_Table_Row).Value = Percent_Price_Change
        Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
        
        'Output Volume in the summary table
        Range("L" & Summary_Table_Row).Value = Volume

        'Add one to the summary table row
        Summary_Table_Row = Summary_Table_Row + 1
        
        'Reset Volume
        Volume = 0
    
        'If the row represents the same ticker:
        Else
        
        'Add to the Volume total
        Volume = Volume + Cells(i, 7).Value
        
        End If
    Next i
End Sub
