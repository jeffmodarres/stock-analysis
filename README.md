# Module 2 Challenge: Stock Analysis

## Overview of Project
Steve is trying to analyze his stock market data for his parents. He wants to summarize total daily volume and one year return in a separate sheet and in this way he evaluates the stock performance.
### Purpose
Goal of this project is to refractor VBA code and measure its performance
## Analysis and Challenges
### Analysis of Stock performance between 2017 and 2018
To analyze stock performance, *volumes, starting and ending prices* for each stock must be found.
A loop was formed to go through all the rows.
*Volume* was found by summing all the tarded stocks over each ticker. 

**TickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value** 

To avoid any mistake, TickerVolumes was zeroed everytime a new ticker was analyzed.

Start and end prices were found by comparing ticker symbol with previous and next rows. 
here is the code snippet for finding Start prices which compares the current row ticker with the previous row. if is different, then it is the first row for the new ticker.

**If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
         End If**

finally, *Return* was calculated by dividing ending price by starting price minus 1 to be shown in percentage.

**tickerEndingPrices(i) / tickerStartingPrices(i) - 1**

In 2017, selected stocks performed extremly well with high return except one as shown below:

![2017_stock_performance](/Resources/2017_stock_performance.png)

**Fig. 1 - Stock performance in 2017**



In contrast, most of the stocks had negative return meaning lost their values as shown in Fig 2. 
"TERP" stock was the only stock that lost it value two years in a row. 

![2018_stock_performance](/Resources/2018_stock_performance.png)

**Fig. 2 - Stock performance in 2018**


### Analysis of the speed
Using "Timer" command, run time for the code was evaluated. The elapsed time for each run is shown below.

![VBA_CHALLENGE_2017](/Resources/VBA_Challenge_2017.png) ![VBA_CHALLENGE_2018](/Resources/VBA_Challenge_2018.png)


## Summary
Refactoring a code is restructuring and existing code without changing its external behaviour. 
In summary, the refactored code runs faster. Instead of looping through all the rows x times (x being the number of unique tickers) as was done in the original code, the refactored code only scans through the rows only once and finds/calculates all the required values. 

Both original and refactored code can furhter be improved by finding the unique values in the ticker column instead of defining them like <tickers(0) = "AY"> 
