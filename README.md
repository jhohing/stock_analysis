# VBA CHallenge

## Overview of Project

### Purpose
This project was meant to refactor the existing VBA script to loop through all the data and collect the same data as before but to make sure it runs faster.

## Results

The starter code for the refactoring was provided. I copied any code what was needed from the original code and then strted following the steps to refactor the code. Below is the code with the steps commented out above each step I completed.

  '1a) Create a ticker Index
    tickerIndex = 0

    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
        
        tickerVolumes(tickerIndex) = 0
        
    Next i
    
        
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then

            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value

        End If
            
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then

            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value


            '3d Increase the tickerIndex.
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
## Summary

###Pros and Cons of Refactoring Code
Refactoringis useful as it helps make our code look cleaner and more organized. Other advantages of refactoring include design and software improvement, debugging and faster programming. Refactoring can also benefit other users that may work on your code or use it for reference especially if it is properly commented. A disadvantage of refactoring is it can be more difficult to refactor another user's code pending how the code was maintained and if there is proper commenting. Without proper commenting, it is diffuclt to understand different parts of the code.

###The Advantages of Refactoring Stock Analysis
The most noticable take away from refactoring the Stock Analysis script is the decreased run time. The original script took nearly half a second to run, whereas the refactored script took a fifth of the time (0.10 seconds) to run. Another advantage is make this code ready to be run on a larger dataset in the future. Below are the screenshots of the run times for the original script and the refactored script.

Original Script:

![](/Resources/VBA_Challenge_2018.png)


Refactored Script:

![](/Resources/VBA_Challenge_2018_refactored.png)

