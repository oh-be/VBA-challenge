# Scripting in Excel (Developer Mode) with Visual Basic for Applications

## Background

Using VBA scripting, I analyzed real stock market data.

### Datasets

* [Test Data](Resources/alphabetical_testing.xlsx) - Dataset used to develop scripts.

* [Stock Data](Resources/Multiple_year_stock_data.xlsx) - Dataset used to test and run the scripts developed using the test data. This dataset is where I generated my final report.

### Stock Market Analysis

![stock Market](Images/stockmarket.jpg)

* I started by creating a script that looped through the entire test dataset and output the following information:

  * The ticker symbol.

  * Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.

  * The percent change from opening price at the beginning of a given year to the closing price at the end of that year.

  * The total stock volume of the stock.

* I used conditional formatting to highlight any positive change in green and any negative change in red (see below).

![moderate_solution](Images/moderate_solution.png)

* I created a small table that returns the stock with the "Greatest % increase", "Greatest % decrease" and "Greatest total volume" (see below).

![hard_solution](Images/hard_solution.png)

* I made the appropriate adjustments to my vba code that allowed a cycle through every worksheet, i.e., every year, for the full stock dataset. I made sure this could be accomplished by running the VBA script only once.
