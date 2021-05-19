 ' ------------------------------------------------------------------------------------------------------
 ' Bonus Question - Greatest and Lowest
 '-------------------------------------------------------------------------------------------------------
 Sub VBA_Neena_Bonus()

  Dim max_percent As Double
  Dim min_percent As Double
  Dim ticker_max As String
  Dim ticker_min As String
  Dim ticker_stock As String
  Dim max_stock As Double
  Dim j As Long
  Dim ws As Worksheet
  Dim LastRow As Long
  
  For Each ws In Worksheets
  
  LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
  
  ws.Range("O2").Value = "Greatest % Increase"
  ws.Range("O3").Value = "Greatest % Decrease"
  ws.Range("O4").Value = "Greatest Total Volume"
  ws.Range("P1").Value = "Ticker"
  ws.Range("Q1").Value = "Value"
  
  max_percent = ws.Cells(2, 11).Value
  min_percent = ws.Cells(2, 11).Value
  max_stock = ws.Cells(2, 12).Value
  ticker_max = ws.Cells(2, 9).Value
  ticker_min = ws.Cells(2, 9).Value
  ticker_stock = ws.Cells(2, 9).Value
  
  For j = 2 To LastRow
  
    If ws.Cells(j, 11).Value >= 0 Then
         If ws.Cells(j, 11).Value > max_percent Then
            max_percent = ws.Cells(j, 11).Value
            ticker_max = ws.Cells(j, 9).Value
         End If
    Else
         If ws.Cells(j, 11).Value < min_percent Then
            min_percent = ws.Cells(j, 11).Value
            ticker_min = ws.Cells(j, 9).Value
         End If
     End If
     
     If ws.Cells(j, 12).Value > max_stock Then
        max_stock = ws.Cells(j, 12).Value
        ticker_stock = ws.Cells(j, 9).Value
     End If
     
   Next j
   
   ws.Range("Q2") = FormatPercent(max_percent, 2)
   ws.Range("Q3") = FormatPercent(min_percent, 2)
   ws.Range("Q4") = max_stock
   
   ws.Range("P2") = ticker_max
   ws.Range("P3") = ticker_min
   ws.Range("P4") = ticker_stock
   
   ws.Columns("O:Q").AutoFit
   
 Next ws
   
End Sub