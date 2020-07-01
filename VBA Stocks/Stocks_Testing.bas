Attribute VB_Name = "Module2"
Sub stocks_testing():
    Dim ticker As String
    Dim year_open As Double
    Dim year_close As Double
    Dim yearly_change As Double
    Dim percent_change As Double
    Dim Table As Integer
    Dim Vol As Double
    Dim Last_Row As Long
    
    
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"
     
    Table = 2
    Vol = 0
    year_open = Cells(2, 3).Value
    Last_Row = Cells(Rows.Count, "A").End(xlUp).Row

   
        
    For i = 2 To Last_Row
        Vol = Vol + Cells(i, 7).Value
        
      'This is where stock ticker is changing
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
            ticker = Cells(i, 1).Value
    
            year_close = Cells(i, 6).Value
            yearly_change = year_close - year_open
            percent_change = 100 * yearly_change / year_open
            year_open = Cells(i + 1, 3).Value
            
            Cells(Table, 9).Value = ticker
            Cells(Table, 10).Value = yearly_change
            Cells(Table, 11).Value = percent_change
            Cells(Table, 12).Value = Vol
            Table = Table + 1
            
            Vol = 0
            
        End If
        
    Next i
    
    Dim cell As Range
    
    For Each cell In Range("Yearly_Change")
    
        If cell.Value > 0 Then
            cell.Interior.Color = "Green"
            
        Else
            cell.Interior.Color = "Red"
        
    End If
    
    Next cell
            
End Sub

