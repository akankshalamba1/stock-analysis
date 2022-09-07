# stock-analysis

## Background: 
Steven Green energy corporation analysis is an analysis of 12 stocks in the stock market of Green Energy Corporations. Our client Steven, who is a finance graduate, his parents are considereing to invest their whole investment portfolio in the DAQO New Energy Corporation, a company that makes silicon wafers for solar panels. They want Steven to handle there portfolio, whereas Steven want to invest the funds in a diversified portfolio, so he wants us to analyze several green energy stocks, in addition to DAQO stock. 
Steve loves the analysis workbook we prepared for him. At the click of a button, he was able to analyze an entire dataset. Now, to do a little more research for his parents, he wants to expand the dataset to include the entire stock market over the last few years. Our code works well for a dozen stocks, it might not work as well for thousands of stocks. So to resolve this issue we refactored our code to make it run quicker as the orignial code might take a long time to run if the data include thousands of values.

## Overview of Project: 
This project is based on the above mentioned analysis of the stocks. Our dataset consists of 3013 rows and 8 columns of 12 different stocks of green energy corporations namely **AY, CSIQ, DQ, ENPH, FSLR, HASI, JKS, RUN, SEDG, SPWR, TERP and VSLR** for the year 2017 and 2018. In our first analysis we used 1 Array "tickers" to run the entire code which provide us effective result but took a longer time as compared to our refactor code. For the 2017 data it took system 1.324219 seconds to run the code and for 2018 analysis it took the system 1.613281 seconds to run the code. As shown below

![Test Image](/Resources/allstockanalysisoutcome2017.png)

2017 Analysis and run time using All stock analysis technique(1 array)


![Test Image](/Resources/allstockanalysisoutcome2018.png)

2018 Analysis and run time using All stock analysis technique(1 array)


To resolve this issue we refactored this code by the use of 4 arrays "tickers, tickerVolumes, tickerEndingPrice, tickerStartingPrices" with the help of arrays and for loops the system was able to look through the dataset in a much quicker and effective manner.

### Steps involved in refactoring code:
1. Initalizing three new arrays: We used three new arrays in the refactor code by initializing tickerVolumes as Long, tickerStartingPrices and tickerEndingPrices as Single.

>   Dim tickerVolumes(11) As Long
    Dim tickerStartingPrices(11) As Single
    Dim tickerEndingPrices(11) As Single

Then we assign all of these arrays = 0 before entering into the For Loop to loop over all the rows in the spreadsheet
    
2. If-Then statements: In this code we used tickerIndex in thr arrays to store the value of tickerVolumes, tickerStartingPrices and tickerEnding prices for each tickers.

> tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value

> If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
  tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
  End If

>  If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
   tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
   End If



that it was able to provide the results as faster as 0.1914063 seconds for the year of 2017 and 0.2851563 seconds for the year 2018 as depicted in the attached screenshots below:

### Refactor code outcome for 2017
![Test Image](/Resources/VBA_Challenge_2017.png)

### Refactor code outcome for 2018
![Test Image](/Resources/VBA_Challenge_2018.png)

## Results:

## Summary:
### What are the advantages ro disadvantages of refactoring code?
### How do these pros and cons apply to refactoring the orignal VBA script?



