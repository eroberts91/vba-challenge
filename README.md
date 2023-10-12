# vba-challenge
VBA script to iterate through stock ticker data and return certain parameters

given an entire year of stock data, gives a summary of data for each stock ticker
- Ticker
- Yearly Change
- Percent yearly change
- total stock volume

Does conditional formatting to show all positive changes in green and negative changes in red

Also iterates through the summary data in order to find the stocks with the
- greatest percent increase
- greatest percent decrease
- greatest total volume


Works by using a for loop to find opening and closing price for each stock, as well as summing total volume for each entry to give total. The summary data is then run through another for loop. 
