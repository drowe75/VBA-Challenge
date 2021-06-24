Attribute VB_Name = "Module1"
 Sub Stocks()

'Set an initial variable to hold the Ticker
Dim Ticker As String

'Set an initial variable to hold the Total Stock Volume
Dim Total_Stock_Volume As Double
Total_Stock_Volume = 0

'Set an initial variable to hold the Opening Price
Dim Opening_Price As Long
Opening_Price = 2

'Set an initial variable to hold the Yearly Change in Price
Dim Yearly_Change As Double

'Set an initial variable to hold the Percent Change in Price
Dim Percent_Change

'Keep track of the location for each Stock in the summary table
Dim Summary_Table_Row As Long
Summary_Table_Row = 2
'Loop through all stocks traded

RowCount = Cells(Rows.Count, "A").End(xlUp).Row
  For i = 2 To RowCount
  
    'Check if it is still the same stock, if it is not...
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
        'set the Ticker
        Ticker = Cells(i, 1).Value
              
        'Subtract closing price at the end of the year from opening price at the begining of the year
        Yearly_Change = Cells(i, 6).Value - Cells(Opening_Price, 3)
        
                
        If Cells(Opening_Price, 3).Value <> 0 Then
                                              
        'Divide Yearly change by Opening Price
        Percent_Change = Round(((Cells(i, 6).Value - Cells(Opening_Price, 3)) / Cells(Opening_Price, 3)) * 100, 2)
            
        Else
          
        Percent_Change = 0
          
        End If
                                         
        If Cells(i + 1, 3).Value >= 0 Then
        Opening_Price = i + 1
        
        End If
                                         
        ' Add to the Total Stock Volume
        Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
        
        'Print the Ticker Name in the Summary Table
        Range("I" & Summary_Table_Row).Value = Ticker
        
        'Print the yearly Change in the Summary Table
        Range("J" & Summary_Table_Row).Value = Yearly_Change
        
        'Print the Percent Change in the Summary Table
        Range("K" & Summary_Table_Row).Value = Percent_Change
        
        'Print the Stock Amount to the Summary Table
        Range("L" & Summary_Table_Row).Value = Total_Stock_Volume
        
        'Add one to the summary table row
        Summary_Table_Row = Summary_Table_Row + 1
        
        'Reset the Stock Total
        Total_Stock_Volume = 0

    ' If the cell immediately following a row is the same Ticker...
    Else

      ' Add to the Stock Total
      Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value

    End If

Next i
                
      
End Sub

