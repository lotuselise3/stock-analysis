# VBA Challenge
## Module 2 Assignment ~ VBA of Wall Street
---
## Overview of Project
- The goal of this project is to help Steve and his parents analyze a much bigger (from a dozen to thousands) dataset that includes the entire stock market over the last few years as well as execute it even faster. Utilizing Visual Basic Application in Excel, we will refactor the code we started working on in this module to find the stock's total daily volume and annual return.  Our findings will help Steve and his parents determine what is the best option, particularly in 2017 and 2018.
### Purpose
- In this challenge, we will refactor (or edit) the Module 2 solution code and hopefully make it more efficient. We will test and see if the refractor code runs cleaner and faster than the code we started writing when learning how to work in VBA. 
---
### Results:  Refactor VBA code and measure performance
- First, I created a new variable called the tickerIndex, which is important to access the index across 4 different arrays. This variable allowed me to assign the output arrays (tickerVolumes, tickerStartingPrices, and tickerEndingPrices) to each ticker symbol before interating through the data set. This makes the analysis to be completed much faster than using the nested for loop for earlier.
- To make my code more efficient, I needed to a different way of nesting order my for loops. 

***Refactored Code***

    'Initialize array of all tickers
    Dim tickers(12) As String
    
    tickers(0) = "AY"
    tickers(1) = "CSIQ"
    tickers(2) = "DQ"
    tickers(3) = "ENPH"
    tickers(4) = "FSLR"
    tickers(5) = "HASI"
    tickers(6) = "JKS"
    tickers(7) = "RUN"
    tickers(8) = "SEDG"
    tickers(9) = "SPWR"
    tickers(10) = "TERP"
    tickers(11) = "VSLR"
    
    'Activate data worksheet
    Worksheets(yearValue).Activate
    
    'Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '1a) Create a ticker Index
    Dim tickerIndex As Single
    tickerIndex = 0

    '1b) Create three output arrays
    ReDim tickerVolumes(12) As Long
    ReDim tickerStartingPrices(12) As Single
    ReDim tickerEndingPrices(12) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
        tickerVolumes(i) = 0
    Next i
        
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        If Cells(i, 1).Value = ticker Then
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        End If
        
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
        tickerIndex = i
        Cells(i + 4, 1).Value = tickers(tickerIndex)
        Cells(i + 4, 2).Value = tickerVolumes(tickerIndex)
        Cells(i + 4, 3).Value = tickerEndingPrices(tickerIndex) / tickerStartingPrices(tickerIndex) - 1
        
    Next i
    
#### Run-time for Refactored Code in 2017 & 2018
![VBA_Challenge_2017](https://user-images.githubusercontent.com/68654746/177700633-bea2c1fa-3426-4c2c-83b7-1484e3bb0af4.png)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/68654746/177700642-963024b5-6f5f-4bb8-b038-55a14d1f631c.png)
#### Run-time for Original Code in 2017 & 2018
![green_stocks_2017](https://user-images.githubusercontent.com/68654746/177701649-2760e952-dfc6-47c3-9346-9058a90cf60b.png)
![green_stocks_2018](https://user-images.githubusercontent.com/68654746/177701860-daeff6a7-6757-400d-a94c-895872dd2e05.png)
- It appears that the run-times for the refactored code are nearly 1 second faster than those from the original code! YAY
- We can also conclude from the images that stock performances in 2017 are significantly better than in 2018.

---

### Summary
1. What are the advantages or disadvantages of refactoring code?
*Advantages...* 
- Refactoring code not only makes the code more efficient, but it also is a key part of any coding process. Code is constantly improved with new functionality as future coders use old code and make older code more readable. It takes fewer steps, uses less memory, and improves the logic overall.
*Disadvantages...*
- It may be disadvantageous to work on code that already works. And, sometimes if you run into errors, it could make the code just unusuable. 
2. How do these pros and cons apply to refactoring the original VBA script?
- Weighing the pros and cons of refactoring code, it challenges you to have a strong understanding of syntax and logic. It may also remind you to value saving often and saving the original code in case you need to either reference it or put it side by side with new code.
