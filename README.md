# An Analysis of Green Stocks in Excel VBA

## Overview of Project

This green stock analysis is for the client, Steve, to determine which stocks in which his parents should invest. The data was analyzed in Microsoft Excel VBA, initially with code that looped through the dataset 12 times but refactored to run through the data a single time in order to make it more efficient and allow the client to expand this code to the entire stock market in the future.

## Data

The data consists the ticker for each stock, date of trade, daily volume, and opening, closing, and adjucted closing price of 12 green stocks for 2017 and 2018. Data for 2017 and 2018 are stored in two sheets in Excel. 

## Model

The analysis used the ticker, starting value, ending value, and volume to determine the total volume and rate of return for each stock in both 2017 and 2018. The original code relied on a [nested loop](https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/fornext-statement) to determine the preformance of each stock as shown below. 


    For i = 0 To 11
        ticker = tickers(i)
        totalVolume = 0

        Sheets(yearValue).Activate
        For j = 2 To RowCount
        
            If Cells(j, 1).Value = ticker Then
                totalVolume = totalVolume + Cells(j, 8).Value
            End If
            
            If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                startingPrice = Cells(j, 6).Value
            End If
            
            If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                endingPrice = Cells(j, 6).Value
            End If
            
        Next j


I refactored the code to run in fewer steps by creating the variable tickerIndex and set it to zero prior to running through the sheets of data. The index was then used to access the correct stock across the four arrays: tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices. 

## Results

### Stock Analysis
The 12 green stocks performed well in 2017. All but one stock had positive returns, with four stocks having over 100% returns. However, in 2018, 10 of the 12 stocks had negative returns. ENPH had the most consistent positive returns at 129.5% and 81.9% in 2017 and 2018, respectively. 

| 2017                      | 2018                      |
|:-------------------------:|:-------------------------:|
|![Screen Shot 2022-06-17 at 5 25 18 PM](https://user-images.githubusercontent.com/106405775/174408528-2f566a15-58ad-41ee-8772-387fc3b75948.png)|![Screen Shot 2022-06-17 at 5 25 49 PM](https://user-images.githubusercontent.com/106405775/174408569-869fd6ce-27ff-49b8-93f5-6ada41e8428e.png)|

### Code Efficiency 
Instead of using a nested loop as shown below, the refactored code used the tickerIndex to analyze the data as shown below. 

 
      For i = 2 To RowCount
    
        If Cells(i, 1).Value = tickers(tickerIndex) Then
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        End If
        
        If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        End If
        
        If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
        End If

        If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
            tickerIndex = tickerIndex + 1
        End If
    
    Next i


The refactored code ran faster that the original, running in .2109 seconds for both 2017 and 2018 as compared to .5644 seconds.

|Refactored Code Speed in 2017  | Refactored Code Speed in 2018 |
|:-----------------------------:|:-----------------------------:|
|![VBA_Challenge_2017](/Resources/VBA_Challenge_2017.png)|![VBA_Challenge_2018](/Resources/VBA_Challenge_2018.png)|

## Summary

### Advantages and Disadvantages to Refractoring Code

Refactoring can lead to better, more effective code. The process can remove redundancies to run more effeciently, improve legibility for other programmers, and restructure to improve software. While there are few disadvantages, refactoring code could introduce new bugs or errors.

### Advantages and Disadvantages to Refactoring Stock Analysis Code in VBA 

The refactored code for the green stock analysis decreased the run time from .5644 seconds to .2109 second. While the dataset only contained two sheets with 3,112 lines in each, this code would work better for analyzing the entire stock market as the client intended.
