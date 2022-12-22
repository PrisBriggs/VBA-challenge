# VBA-challenge
VBA-challenge (Homework for Module 2 - Priscila Briggs)

The VBA scripting presented in this challenge was created to analyze generated stock market data.
The code was created by Priscila Menezes Briggs in December 2022.

The source table with raw data was provided by GA Tech / edX Boot Camps LLC. The data is presented in an Excel file (.xlsx) containing 7 columns with data about different tickers during the years 2018 to 2020. The columns show information for ticker symbols, date, and for this date: opening price, highest and lowest price, closing price and volume of stocks.

The calculated data shows information about any given ticker symbol during each year. The data calculated was: yearly change, percent change and total stock volume, and besides this, the tickers with greatest percentage increase and greatest percentage decrease were found as well as the ticker with greatest total volume. All these values were calculated and are shown in the Results Panel table.

The VBA scripting includes one sub-routine which follows the steps below to calculate the data:

1.INITIAL PROCEDURES
  1.1. Setting loop through all worksheets
  1.2. Setting variables 
  1.3. Setting tables and rows headers
  
2. CALCULATING THE VALUES FOR THE RESULTS PANEL
  2.1. Procedures to find opening price and closing prices for each ticker 
  2.2. Calculations of the total stock volume for each ticker symbol during each year, the yearly change and the percent change for each ticker. 
  
3. CALCULATING THE MAXIMUM AND MINIMUM VALUES
  3.1. Calculations of the greatest percentage increase, greatest percentage decrease and greatest total stock volume between all the tickers on each year. 

4. FORMATTING
  4.1. Yearly Changes column formatted to show positive (including zero) numbers in green and negative numbers in red. 
  4.2. Both the Yearly Changes column and the Percent Changes Column formatted to show two decimal places. 
  4.3. Fortmatting of the Percentage Changes Column to show values in %. 

5. FINAL PROCEDURES
  5.1. Message Box to inform the user about the end of the calculations by Excel.
  5.2. Final procedure to ensure looping through all the worksheets.
  
The following files are attached to this assignment:
  - Word document with screenshots of each worksheet with all the results
  - .vbs file with code with comments
  - .vbs file with code without comments
  - this Read.me file with explanation about the assignment

The following websites were researched to perform some calculations and formatting:

Formula for finding the maximum value
  website (visited in December 2022)  https://www.educba.com/vba-max/
  
Formula for formatting numbers to be displayed in percentage and with two decimal places
  website (visited in December 2022)  https://excelforever.com.br/formato-de-numeros-por-codigos-no-vba/
  
All the remaining procedures of this script were based in the activities learned in class.


