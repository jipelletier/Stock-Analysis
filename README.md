# Stock-Analysis

## Overview of Project
### Purpose
  The purpose of this project is to assist recent Finance graduate, Steve, by looking into DQ stocks for his parents. He has extracted data from a few green energy stocks in order to possibly diversify their portfolio. We are using VBA to automate tasks in order to supply Steve with code that will be able to be reused with any stock. We will complete this task by refractoring a VBA code to collect specific stock data in the years 2017 and 2018 in order to increase efficiency. We will use this data to determine whether or not it is worth it for Steve's parents to invest.

## Results
  There is a stark difference between the stock performance in 2017 and 2018. 10 of the 12 companies returned less than 1% for the year 2018 while 10 of the 11 companies returned a percentage above 1. The code below was used to determine the return and total daily volume of the "Tickers". The images listed below detail how we were able to increase the overall run time by refactoring the script.
  
### Total Daily Volume and Return of Tickers Code

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
        
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, i).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        
        End If
            
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
            
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            
        End If
        
            '3d Increase the tickerIndex.
        
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
   
 ### Execution Times
 
 ![VBA-Challenge_2018.png](https://github.com/jipelletier/Stock-Analysis/blob/main/Resources/VBA-Challenge_2018.png)
 ![VBA_Challenge_2017.png](https://github.com/jipelletier/Stock-Analysis/blob/main/Resources/VBA_Challenge_2017.png)

## Summary
### Advantages of Refactoring
  Advantages of refactoring include that the code is easier to read and understand. It helps to find any bugs and it provides a clearer and concise code to the user with the potential of it being reusable.
### Disadvantages of Refactoring
  A few disadvantages may prevent users from being able to refactor code such as if there are no reliable test cases. Refactoring can also be time consuming especially if the user is without the allotted time to be precise. This could lead to errors and the process of debugging the code.
### Advantages and Disadvantages of the Original and Refactored VBA Script
  The decrease in macro run time proved to be one of the most signicant advantages to refactoring the code. Refactoring the original VBA script proved that the new analysis took less time to run due to the refactoring of the code. The disadvantages to the refactored code included the process of debugging when trying to determine the appropriate code. The advantage of the original code was that is clearly defined each code and its purpose thus making it easier to determine the source of errors. The disadvantage of the original was that it was not as legible as the refactored. It took more time to dissect and determine where the code for different instructions were.
