Attribute VB_Name = "Module3"
Sub Ticker2014()


Dim TickerName As String
Dim Ticker_Column As Integer
Dim total As Double, i As Long, per_change, year_change As Single, pop1, pop2 As Long
Dim sheet1, sheet2, sheet3 As Worksheet

'Set sheet1 = Worksheets("Year 2016")
'Set sheet2 = Worksheets("2015")
Set sheet3 = Worksheets("Year 2014")

Ticker_Column = 2
total = 0

lastrow = Cells(Rows.Count, 1).End(xlUp).Row


' Loop through all Tickernames

For i = 2 To lastrow

    If sheet3.Cells(i + 1, 1).Value <> sheet3.Cells(i, 1).Value Then
    
       ' Set the Ticker name
      TickerName = sheet3.Cells(i, 1).Value
      
      ' Add to the Total
      total = total + sheet3.Cells(i, 7).Value
      
      ' Print the TickerName in Column H
      
      sheet3.Range("H" & Ticker_Column).Value = TickerName
      
      ' Print the Total Volume in Column K
      sheet3.Range("K" & Ticker_Column).Value = total
      
      ' Add one to the summary table row
      Ticker_Column = Ticker_Column + 1
      
       ' Reset the Total
      total = 0
    
    'If the cell immediately following a row is the same ticker value.
    Else
    
        'Add to the Total
        
      total = total + sheet3.Cells(i, 7).Value
    
    End If
    
Next i

End Sub
