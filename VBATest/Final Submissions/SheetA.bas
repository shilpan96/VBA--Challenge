Attribute VB_Name = "Module1"
Sub AlphabestFinal()

'Define Variables

Dim TickerName As String
Dim Yearchange, Per_change, TotalVolume, Openprice, Closeprice, Summary_Table_Row As Double
Dim sheet1, sheet2, sheet3 As Worksheet

Yearchange = 0
Per_change = 0
TotalVolume = 0
Openprice = 2
Summary_Table_Row = 2

Dim Days As Integer
Days = 0


Set sheet1 = Worksheets("A")
'Set sheet2 = Worksheets("B")
'Set sheet3 = Worksheets("C")

lastrow = Cells(Rows.Count, 1).End(xlUp).Row
newlastrow = Cells(Rows.Count, 9).End(xlUp).Row


For i = 2 To lastrow
    If sheet1.Cells(i + 1, 1).Value <> sheet1.Cells(i, 1).Value Then
        TickerName = sheet1.Cells(i, 1).Value
    
        Openprice = sheet1.Cells(i - Days, 3).Value
        Closeprice = sheet1.Cells(i, 6).Value
        Yearchange = Closeprice - Openprice
        
            'If Openprice = 0 Then
                  'Per_change = 0
            'Else
                Per_change = (Closeprice - Openprice) / Openprice
        
  
                      
    
        TotalVolume = TotalVolume + sheet1.Cells(i, 7).Value
    
        ' Print the Ticker,Yearly,Percent and Volume
    
      
      sheet1.Range("H" & Summary_Table_Row).Value = TickerName
      sheet1.Range("I" & Summary_Table_Row).Value = Yearchange
      sheet1.Range("J" & Summary_Table_Row).Value = Format(Per_change, "Percent") 'Change per_change to Percent
      sheet1.Range("K" & Summary_Table_Row).Value = TotalVolume
      
         
      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
       ' Reset the TotalVolume,Days,PercentChange
      TotalVolume = 0
      Days = 0
      Per_change = 0
    
    'If the cell immediately following a row is the same ticker value.
    Else
    
    'If sheet1.Cells(i + 1, 1).Value = sheet1.Cells(i, 1).Value Then
        'Add to the TotalVolume,Days
        
      TotalVolume = TotalVolume + sheet1.Cells(i, 7).Value
      Days = Days + 1
    
    End If

Next i

 'To highlight positive change in green (4) and negative change in red(3).
 
       For j = 2 To newlastrow

    If sheet1.Cells(j, 9) > 0 Then
        sheet1.Cells(j, 9).Interior.ColorIndex = 4
    Else
        sheet1.Cells(j, 9).Interior.ColorIndex = 3
        
    End If
    


If sheet1.Cells(j, 10) > Max_Percentage Then
Max_Percentage = sheet1.Cells(j, 10).Value
TickerValue = sheet1.Cells(j, 8).Value
End If

Next j

sheet1.Cells(4, 15).Value = Max_Percentage
sheet1.Cells(4, 16).Value = TickerValue
sheet1.Cells(3, 15).Value = "Ticker"
sheet1.Cells(3, 16).Value = "Value"

End Sub
     
