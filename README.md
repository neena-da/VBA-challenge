# VBA Homework - The VBA of Wall Street

## Task

* Create a script that will loop through all the stocks for one year and output the following information.

  * The ticker symbol.

  * Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.

  * The percent change from opening price at the beginning of a given year to the closing price at the end of that year.

  * The total stock volume of the stock.

* You should also have conditional formatting that will highlight positive change in green and negative change in red.

### Solution - VBA Code Wall_Street_VBA_Exercise

1. Define the variables as required
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


2. Using a For loop, populate the ticker symbol and check for each new ticker symbol and 

### Files

* [Test Data](Resources/alphabetical_testing.xlsx) - Use this while developing your scripts.

* [Stock Data](Resources/Multiple_year_stock_data.xlsx) - Run your scripts on this data to generate the final homework report.

### Stock market analyst

![stock Market](Images/stockmarket.jpg)

## Instructions

* Create a script that will loop through all the stocks for one year and output the following information.

  * The ticker symbol.

  * Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.

  * The percent change from opening price at the beginning of a given year to the closing price at the end of that year.

  * The total stock volume of the stock.

* You should also have conditional formatting that will highlight positive change in green and negative change in red.

* The result should look as follows.

![moderate_solution](Images/moderate_solution.png)

## BONUS

* Your solution will also be able to return the stock with the "Greatest % increase", "Greatest % decrease" and "Greatest total volume". The solution will look as follows:

![hard_solution](Images/hard_solution.png)

* Make the appropriate adjustments to your VBA script that will allow it to run on every worksheet, i.e., every year, just by running the VBA script once.

## Other Considerations

* Use the sheet `alphabetical_testing.xlsx` while developing your code. This data set is smaller and will allow you to test faster. Your code should run on this file in less than 3-5 minutes.

* Make sure that the script acts the same on each sheet. The joy of VBA is to take the tediousness out of repetitive task and run over and over again with a click of the button.

* Some assignments, like this one, contain a bonus. It is possible to achieve mastery on this assignment without completing the bonus. The bonus adds an opportunity to further develop you skills and be rewarded extra points for doing so.

## Submission

* To submit please upload the following to Github:

  * A screen shot for each year of your results on the Multi Year Stock Data.

  * VBA Scripts as separate files.

* Ensure you commit regularly to your repository and it contains a README.md file.

* After everything has been saved, create a sharable link and submit that to <https://bootcampspot-v2.com/>.

## Rubric

[Unit 2 Rubric - VBA Homework - The VBA of Wall Street](https://docs.google.com/document/d/1OjDM3nyioVQ6nJkqeYlUK7SxQ3WZQvvV3T9MHCbnoWk/edit?usp=sharing)

- - -

© 2021 Trilogy Education Services, LLC, a 2U, Inc. brand. Confidential and Proprietary. All Rights Reserved.
