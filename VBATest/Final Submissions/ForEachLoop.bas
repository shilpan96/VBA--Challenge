Attribute VB_Name = "Module4"
Sub ForEachAlpha()
'Define Variables
Dim TickerName As String
Dim Yearchange, Per_change, TotalVolume, Openprice, Closeprice, Summary_Table_Row As Double
Dim sh As Worksheet
Yearchange = 0
Per_change = 0
TotalVolume = 0
Openprice = 2
Summary_Table_Row = 2
Dim Days As Integer
Days = 0

For Each ws In Worksheets

lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
newlastrow = ws.Cells(Rows.Count, 9).End(xlUp).Row

    For i = 2 To lastrow
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            TickerName = ws.Cells(i, 1).Value
            Openprice = ws.Cells(i - Days, 3).Value
            Closeprice = ws.Cells(i, 6).Value
            Yearchange = Closeprice - Openprice
    'To calculate Percent Change
        If Openprice = 0 Then
            Per_change = 0
        Else
            Per_change = (Closeprice - Openprice) / Openprice
        End If
        'To calculate Total
        TotalVolume = TotalVolume + ws.Cells(i, 7).Value
        ' Print the Ticker,Yearly,Percent and Volume
      ws.Range("H" & Summary_Table_Row).Value = TickerName
      ws.Range("I" & Summary_Table_Row).Value = Yearchange
      ws.Range("J" & Summary_Table_Row).Value = Format(Per_change, "Percent") 'Change per_change to Percent
      ws.Range("K" & Summary_Table_Row).Value = TotalVolume
      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
       ' Reset the TotalVolume,Days,PercentChange
      TotalVolume = 0
      Days = 0
      Per_change = 0
    'If the cell immediately following a row is the same ticker value.
    Else
    'Add to the TotalVolume,Days
      TotalVolume = TotalVolume + ws.Cells(i, 7).Value
      Days = Days + 1
      
    End If
    
Next i

 'To highlight positive change in green (4) and negative change in red(3).
 For j = 2 To newlastrow
 
    If ws.Cells(j, 9) > 0 Then
        ws.Cells(j, 9).Interior.ColorIndex = 4
    Else
        ws.Cells(j, 9).Interior.ColorIndex = 3
    End If
    
Next j

' Greatest % Increase
max_value = WorksheetFunction.Max(ws.Range("J:J"))
max_ticker = WorksheetFunction.Match(max_value, ws.Range("J:J"), 0)
ws.Cells(4, 15).Value = ws.Cells(max_ticker + 1, 1)
ws.Cells(4, 16).Value = max_value
'Greatest % Decrease
min_value = WorksheetFunction.Min(ws.Range("J:J"))
min_ticker = WorksheetFunction.Match(min_value, ws.Range("J:J"), 0)
ws.Cells(5, 15).Value = ws.Cells(min_ticker + 1, 1)
ws.Cells(5, 16).Value = min_value
'Greastest Total Volume
max_value = WorksheetFunction.Max(ws.Range("K:K"))
max_ticker = WorksheetFunction.Match(max_value, ws.Range("K:K"), 0)
ws.Cells(6, 15).Value = ws.Cells(max_ticker + 1, 1)
ws.Cells(6, 16).Value = max_value

Next ws

End Sub
