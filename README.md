# Stock-Analysis-in-Visual-Basic-for-Applications

> Use VBA to analyze generated stock market data. Output the changes that occurred in a year and the greatest increases, decreases, and total volume along with their associated tickers.

## Background - Requirements

> This project uses VBA script to analyze stock market data. In developing the script, the [Test Data](./VBA_Script_Test_Sample/alphabetical_testing.xlsm) was used for its smaller sample size.
once the script was developed, it was applied to the [Stock Data](./VBA_Stock_Data/Multiple_year_stock_data.xlsm) set. The script is meant to loop through all the data and output two sets of information.

| Data Set | Description |
|--------------|---------------|
| Summary of the Individual Stock | The script outputs a line for each stock. In this line, it will summarize the stock with its Ticker, Yearly Change (+/-), Percent Change(+/-), and Total Stock Volume |
| Summary of the Greatest Changes & Values | The script outputs a line for 3 stocks; the greatest % increase, the greatest % decrease, and greatest total volume along with the ticker for each stock referenced respectively |

<object data="./Images/2018Snapshot_of_Results.pdf" type="application/pdf" width ="100%">
</object>