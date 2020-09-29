# Stocks Anaysis

## Project overview
Utilizing Visual Basic for Applications (VBA) to analyze performance of 10 stocks.
The VBA macro is used on Excel sheet with daily stocks performance. The macro compiles daily data of each stocks and calculates yearly starting price, ending price and yearly volume sold. If the stock return is positive the stock is marked as "green"

## Macro details
The code loops over all rows and checks if the value in A column corresponds to the ticker - outer loop
If the value corresponds to the ticker, the inner loop calculates total volume, starting and ending price. If the stock return is positive the stock is marked as "green"
Once the ticker does no match value in cell of column A, the outer loop is executed that changes the ticker
I refined the code by removing extra loop from formatting results. I added starting and ending price.


## Resources

* Excel sheet with embedded macro: green_stocks_TO.xlsm
* VBA macro: Stocks_analysis.bas 




