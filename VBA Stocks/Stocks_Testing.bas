Attribute VB_Name = "Module1"
Sub stocks_testing():
    Dim ticker As String
    Dim yearly_close As Double
    Dim yearly_change As Double
    Dim percent_change As Double
    Dim Table As Integer
    Dim Vol As Integer
    
    
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"
        
    For i = 2 To LastColumn
        
        If Cells(i + 1, Column).Value <> Cells(i, Column).Value Then
        
            ticker = Cells(i, 1).Value
            Vol = Cells(i, 7).Value
            year_open = Cells(i, 6).Value
            yearly_change = year_close - year_open
            
            Cells(Table, 9).Value = ticker
            Cells(Table, 10).Value = yearly_change
            Cells(Table, 11).Value = percent_change
            Cells(Table, 12).Value = Vol
            Table = Table + 1
            
            Vol = 0
            
        End If
        
    Next i
            
End Sub
