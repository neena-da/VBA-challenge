# VBA Homework - The VBA of Wall Street

## Task

* Create a script that will loop through all the stocks for one year and output the following information.

  * The ticker symbol.

  * Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.

  * The percent change from opening price at the beginning of a given year to the closing price at the end of that year.

  * The total stock volume of the stock.

* You should also have conditional formatting that will highlight positive change in green and negative change in red.

### Solution - VBA Code Wall_Street_VBA_Exercise

1. Define the variables as required and initialise variables as necessary
```
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
```

2. * Populate the headings of the table
    
   * Using a For loop, populate the ticker symbol and stock volume for each new ticker symbol

   * Calculate the yearly change by finding the difference between the opening stock and closing stock for each ticker in a year

   * Calculate the percent amount by dividing the yearly change with the opening amount 

   ** If the opening stock is 0, populate the Percent change as 0 to avoid stack overflow

```
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
```

3. Perform conditional formatting on the Yearly change field so that all positive changes are highlighted in Green and Negative    changes are highlighted in Red.

```
          If yearly_change >= 0 Then
             ws.Range("J" & ticker_row).Interior.ColorIndex = 4
          Else
             ws.Range("J" & ticker_row).Interior.ColorIndex = 3
          End If
```	

4. Loop the code through each worksheet so that it runs for each year in the workbook. This is done in the beginning of the code

```
          For Each ws In Worksheets
```

### Bonus Question - used the for loop to check for the greatest increase and decrease in the Percent change and also the highest stock volume

```
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
```


