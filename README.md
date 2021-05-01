# stock-analysis

## Overview of Project
   The dataset contains stock prices for 12 different stocks and Steve knows how to analyze the stock performance for his parents by a click of a button. He likes the code that    was written for him. But the code can be written in a much cleaner way which is easy to understand and which is faster to execute. 	

### Purpose
   The purpose of this project is to refactor the Microsoft Excel VBA code that was written to collect stock performance for 12 stocks for the years 2017 and 2018 and make
   the execution time faster than the original code.

### The Data
   The data is being analyzed based on the stock name (tickers), stock volume (tickersVolume) which is the total volume of a stock on a given date, stock starting price           (tickerStartingPrices) and stock ending prices (tickerEndingPrices) to arrive at returns denoted in percentages

### Analysis
The existing written code was analyzing stock information in the 0.1953125 seconds for year 2017 and 0.1875 seconds for year 2018.

![VBA_Challenge_2017 (Original)](https://github.com/yashodhan1202/stock-analysis/blob/main/VBA_Challenge_2017%20(Original).png)

![VBA_Challenge_2018 (Original)](https://github.com/yashodhan1202/stock-analysis/blob/main/VBA_Challenge_2018%20(Original).png)

The approach taken to refactor the code was as follows :-
    
    '1a) Create a ticker Index
        tickerIndex = 0
        
    '1b) Create three output arrays
        ReDim tickerVolumes(12) As Long
        ReDim tickerStartingPrices(12) As Single
        ReDim tickerEndingPrices(12) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
        For i = 0 To 11
        tickerVolumes(i) = 0
        tickerStartingPrices(i) = 0
        tickerEndingPrices(i) = 0
        Next i
        
    ''2b) Loop over all the rows in the spreadsheet.
        
        For i = 2 To RowCount
           
        '3a) Increase volume for current ticker
        
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
             
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
           If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
           tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
           End If
        
        '3c) check if the current row is the last row with the selected ticker
        
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            End If
            
        '3d) Increase the tickerIndex.
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerIndex = tickerIndex + 1
            End If
                        
        Next i
                
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
        
        For i = 0 To 11
    
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
        Next i

## Results

After refactoring the code the runtime was 0.1679688 seconds for year 2017 and 0.1796875 seconds for year 2018. Below are the screenshots of runtime after refactoring of the code :-

![VBA_Challenge_2017](https://github.com/yashodhan1202/stock-analysis/blob/main/VBA_Challenge_2017.png)

![VBA_Challenge_2018](https://github.com/yashodhan1202/stock-analysis/blob/main/VBA_Challenge_2018.png)

The reduction in runtime was 0.0273437 and 0.0078125 for the years 2017 and 2018 respectively.The stock performance results were matched and rechecked thoroughly with the original code.

![Stock_Performance_2017](https://github.com/yashodhan1202/stock-analysis/blob/main/Stock_Performance_2017.png)

![Stock_Performance_2018](https://github.com/yashodhan1202/stock-analysis/blob/main/Stock_Performance_2018.png)

An extra step was taken to analyze the refactored written by putting a break at the "Next i" before point 4 in the above code and executing the code 251 times manually in the locals window to see if the ending price is being populated for ticker(0) when "i" is 252. Here is the screen shot of the locals window :-

![Ticker_num_screenshot1](https://github.com/yashodhan1202/stock-analysis/blob/main/Ticker_num_screenshot1.png)
![Ticker_num_screenshot2](https://github.com/yashodhan1202/stock-analysis/blob/main/Ticker_num_screenshot2.png)

## Summary

### Advantages of refactoring
 - The code can be made efficient by refactoring an existing code. This is evident in the lesser number of seconds taken to execute the code and give the results.
 - The code can be made cleaner by adding comments which VBA ignores in executing the codes. The comments act like stones from the Hansel and Gretel story incase we lose our way    and forget after a period of time and have to look back

### Disadvantages of refactoring 
 - Refactoring is using someone else's code. If not done properly the code can be inefficient and can render a perfectly running code inefficient and difficult to understand.
 - Refactoring can be time consuming as you first have to understand the logic of the code written and then analyze whether any improvements can be made or not

### Advantages of refactored code
 - The biggest advantage of the refactored code was reduction in execution time.
 - Since arrays were used they could be used in the code instead of calculating through a code in itself.

### Disadvantages of refactored code
 - There were no disadvantages noticed in the refactored code as the results were exactly the same as before and the runtime was faster than the previous one. Since this was the    first time I was refactoring the code, it was a bit time consuming to write an efficient code which would give the same results as before.
