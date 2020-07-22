# VBA - The VBA of Wall Street

## Background

In this project, I will use VBA scripting to analyze real stock market data.

### Before I began thus project

1. Created a new repository for this project called `VBA-challenge`.

2. Created a directory for both of the VBA Challenges. Using the folder name to correspond to the challenge: **VBAStocks**.

4. Inside **VBAStocks**, all VBA files were stored. These are the main scripts to run for each analysis.

5. Push the above changes to GitHub.

### Files

* [Test Data](Resources/alphabetical_testing.xlsx) - used  while developing my scripts.

* [Stock Data](Resources/Multiple_year_stock_data.xlsx) - all scripts were run on this data to generate the final report.

### Stock market analyst

![stock Market](Images/stockmarket.jpg)

## Method

* Created a script that will loop through all the stocks for one year and output the following information.

  * The ticker symbol.

  * Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.

  * The percent change from opening price at the beginning of a given year to the closing price at the end of that year.

  * The total stock volume of the stock.

* Includes conditional formatting that will highlight positive change in green and negative change in red.

* The results output as follows:

![moderate_solution](Images/moderate_solution.png)