# Project Overview

## Background

Steve's parents have invested a lot of their money into the DQ stock. Steve wants to analyze the performance of the DQ stock, along with similar stocks to track the health of his parent's investment. Steve will need the help of a VBA macro to create this analysis quickly!

## Purpose

While happy with the results of the workbook provided, Steve wants to do more research for his parents and include the entire stock market over the last few years. The original code must be refactored to perform the same function, but over any number of stocks.

# Results

Refactoring the code resulted in a much faster macro for Steve's workbook. By refactoring the macro to only loop through all the data one time, we have cut down the total running time down to about 21% (.12 / .56). Below are comparisons between the original and refactored analyses for 2017:

![2017 Original](https://github.com/juberr/VBA-Stock-Analysis/blob/master/Challenge/Resources/Original%202017%20Time.png?raw=truue) 
![2017 Refactored](https://github.com/juberr/VBA-Stock-Analysis/blob/master/Challenge/Resources/Refactored%202017%20Time.png?raw=true)

This comparison demonstrates the significant decrease in processing time that came as a result of refactoring the code (from .56 to .12 seconds). While it may not seem like a huge decrease for the analysis of 12 stocks, the decrease in time will pay dividends as the number of stocks Steve wants to analyze increases.

The power of this refactoring comes in the use of arrays and for loops for the entire given data. By creating a variable that indexes which Ticker it is on (tickerIndex) and how many Tickers there are (numTickers), this code can scale dynamically to the array declared before it. Only having to loop through the data once saved a lot of processing time!

The code below is a look at the for loops used in this refactoring. Using variables such as RowCount, numTickers, and tickerIndex allow for the code to dynamically scale to how much data there is.
```vb
For i = 0 To numTickers
        tickerVolumes(i) = 0
        
        Next i
        
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
            'If the previous row's ticker doesn't match, grab the tickerStartingPrice for the current tickerIndex
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
                
                tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
            
            End If
        
        '3c) check if the current row is the last row with the selected ticker:
            'If the next row’s ticker doesn’t match , grab the tickerEndingPrice for the current tickerIndex and then increase the tickerIndex.
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
                
                tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            

            '3d Increase the tickerIndex.
                tickerIndex = tickerIndex + 1
            
        End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To numTickers
        
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1

        
    Next i
```
 

# Summary

## Refactoring in General

Refactoring code is a great practice for creating efficiencies in functional programs. As programs grow in size, resources such as time and memory become greater concerns. A potential issue with refactoring is that it could introduce new bugs into an otherwise functional program. Also, if you're on a team it will take additional time to ensure everyone on the team understands the refactored code.


## Refactoring this Code
While Steve's program is only taking fractions of a second, it is only analyzing 12 stocks. For scale, the TSX has a total of 1,500 stocks. Cutting the running time of the code down to 21% will save Steve time in the order of magnitudes! The current program is taking about `.01` seconds per stock, which means we can estimate that this program would take `15` seconds to analyze the entire TSX (assuming the same amount of data is collected for each stock). For comparison, the original VBA script took about `.04` seconds per stock, which means we can estimate that it would take about `62` seconds to analyze the entire TSX (broke the minute barrier!).




