Sub Wall_Street_VBA_Exercise()
    Dim ticker As String
    Dim LastRow As Long
    Dim i As Long
    Dim ticker_row As Long
    Dim stock_volume As Double
    Dim year_change_row As Long
    Dim yearly_change As Double
    Dim opening_amount As Double
    Dim closing_amount As Double
    Dim percent_amount As Double
    Dim ws As Worksheet
    
' Looping thorugh each worksheet

For Each ws In Worksheets
    ticker_row = 2
    stock_volume = 0
    year_change_row = 2
    
    ' Determine the Last Row
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    ' Adding Headings for the Table
      ws.Range("I1") = "Ticker"
      ws.Range("J1") = "Yearly Change"
      ws.Range("K1") = "Percent Change"
      ws.Range("L1") = "Total Stock Volume"
      ws.Columns("I:L").AutoFit
    
    For i = 2 To LastRow
       If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
       
       '  Calculating Total Stock Volume
          stock_volume = stock_volume + ws.Cells(i, 7)
          ws.Range("L" & ticker_row).Value = stock_volume
          
       '  Populating Ticker
          ws.Range("I" & ticker_row).Value = ws.Cells(i, 1).Value
          
       '  Calculating Yearly Change and Percent Chane
          opening_amount = ws.Cells(year_change_row, 3).Value
          closing_amount = ws.Cells(i, 6).Value
          yearly_change = closing_amount - opening_amount
          
       '  For Records where Opening and Closing amounts are zeroes
           If opening_amount = 0 Then
            ws.Range("J" & ticker_row).Value = yearly_change
            ws.Range("K" & ticker_row).Value = 0
           Else
            percent_amount = yearly_change / opening_amount
            ws.Range("J" & ticker_row).Value = yearly_change
            ws.Range("K" & ticker_row).Value = FormatPercent(percent_amount, 2)
           End If
          
       '  Conditional Formatting for Positive and Negative Yearly Change
       
          If yearly_change >= 0 Then
             ws.Range("J" & ticker_row).Interior.ColorIndex = 4
          Else
             ws.Range("J" & ticker_row).Interior.ColorIndex = 3
          End If
          
          stock_volume = 0
          ticker_row = ticker_row + 1
          year_change_row = i + 1
          
       Else
       ' Calculating Total Stock Volume within the same ticker range
          stock_volume = stock_volume + ws.Cells(i, 7)
          
       End If
       
    Next i
    
  Next ws
        
End Sub