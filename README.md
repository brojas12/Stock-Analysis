# Module 2 Challenge

## Overview of Project
**#Purpose**

    The purpose of this project was to gather stock information for the year 2017 and 2018 to determine
    which stocks were worth investing. Furthermore, the goal of this assignment was to 
    improve the efficiency of the original code.

**#The Data**

    The information presented includes two tables with stock information on 12 different stocks. 
    The stock information contains a ticker value, the date the stock was issued, the opening, 
    closing and adjusted closing price, the highest and lowest price, and the volume of the stock. 
    The objective is to retrieve the ticker, the total daily volume, and the return on each stock.

# Results

**##Analysis**

    According to the data gathered from the 2017 Stock Analysis the worst stock to invest in is 
    TERP with a return rate of -7.2%. Furthermore, based on the data gathered from the 2018 Stocks, 
    the best stocks to invest in with an 81.9% and 84% Return are ENPH and RUN respectively. 
    I’ve included some of the code down below with comments describing what each line of code does. 
    As well as the resulting grapshs.
    '1b) Create three output arrays
            Dim tickerVolumes(12) As Long
            Dim tickerStartingPrices(12) As Single
            Dim tickerEndingPrices(12) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
        For i = 0 To 11
        tickerVolumes(i) = 0
        tickerStartingPrices(i) = 0
        tickerEndingPrices(i) = 0
        Next i
        
    ''2b) Loop over all the rows in the spreadsheet.
        Worksheets(yearValue).Activate
    
        For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
                
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next rowâ€™s ticker doesnâ€™t match, increase the tickerIndex.
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
        
        
    Next i
    
    ![This is a alt text.](/Resources/VBA_Challenge_2017.png "VBA_Challenge_2017")

# Summary

**#Advantages and Disadvantages of Refactoring Code**

    The method known as refactoring creates cleaner and more organized code. The benefits include 
    improvements to the design and software, debugging, and quicker programming. The code itself 
    becomes more user friendly in the sense that it’s easier to read and understand. Some disadvantages 
    include having applications that are too large and not having the proper test cases for the 
    existing codes,  all of which pose some risk when attempting to refactor code.

**#Advantages of Refactoring Stock Analysis**

    The most noticeable change as a result of refactoring our code was a decrease in macro run time. 
    The module analysis took approximately one second to run, whereas the new analysis took 
    approximately 0.22 seconds to run. I’ve attached screenshots that indicate the run time 
    for the new analysis.

