# VBA_Challenge
Sub Stock_Market_Analysis()

'Sub Stock_Market_Analysis()

' Set variables

Dim Ticker As String
Dim Stock_Open As Double
Stock_Open = 0
Dim Stock_Close As Double
Stock_Close = 0
Dim Total_Stock_Volume As LongLong
Total_Stock_Volume = 0
Dim Yearly_Change As Double
Yearly_Change = 0
Dim Percent_Change As Double
Percent_Change = 0

Dim Summary_Table_Row As LongLong
Summary_Table_Row = 2
  
Dim Lastrow As Long
Lastrow = Cells(Rows.Count, 1).End(xlUp).Row

Stock_Open = Cells(2, 3).Value
        
  For i = 2 To Lastrow
          If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
          ' set ticker name
          Ticker = Cells(i, 1).Value
          'calculations of additional values
          Stock_Close = Cells(i, 6).Value
          Yearly_Change = Stock_Close - Stock_Open
          If Stock_Open <> 0 Then
          Perent_Change = (Yearly_Change / Stock_Open) * 100
      
      Else
      
      Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
      
      End If
      'Print values in summary table
      Range("I" & Summary_Table_Row).Value = Ticker
      Range("J" & Summary_Table_Row).Value = Yearly_Change
      
       If (Yearly_Change > 0) Then
                    'Fill column with GREEN color - good
                    Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                ElseIf (Yearly_Change <= 0) Then
                    'Fill column with RED color - bad
                   Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                   
        End If
                
      Range("K" & Summary_Table_Row).Value = Percent_Change
      Range("L" & Summary_Table_Row).Value = Total_Stock_Volume
    
    Summary_Table_Row = Summary_Table_Row + 1
    
    Total_Stock_Volume = 0
    Stock_Close = 0
    Stock_Open = 0
    
    Stock_Open = Cells(i + 1, 3).Value

End If
Next i
    
End Sub
