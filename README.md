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
1. Initalizing three new arrays: 
We used three new arrays in the refactor code by initializing tickerVolumes as Long, tickerStartingPrices and tickerEndingPrices as Single.

>    Dim tickerVolumes(11) As Long 
---
>    Dim tickerStartingPrices(11) As Single
---
>    Dim tickerEndingPrices(11) As Single
---
Then we assign all of these arrays = 0 before entering into the For Loop to loop over all the rows in the spreadsheet
    
2. If-Then statements: 
In this code we used tickerIndex in the arrays to store the value of tickerVolumes, tickerStartingPrices and tickerEnding prices for each tickers.

To calculate tickerVolumes for each stocks

> tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value

To calculate the tickerStartingPrices for each stock:
> If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then 
  tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
  End If
  
To calculate the tickerEndingPrices for each stock:
>  If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
   tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
   End If

By using tickerIndex we are able to get all the values for each stock. In order to get tickerIndex to increase and run throughout all the 3013 rows of the dataset we need to increase the tickerIndex so it changes the ticker and helps us accumilate the currect value for each stock.

Following code is used to increase the tickerIndex
> If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
   tickerIndex = tickerIndex + 1
   End If
   
All this is done in the same For loop **For i = 2 to 3013**

## Results:
By refactoring the code as mentioned above we were able to provide the results as faster as 0.1914063 seconds for the year of 2017 and 0.2851563 seconds for the year 2018 as depicted in the attached screenshots below:

### Refactor code outcome for 2017
![Test Image](/Resources/VBA_Challenge_2017.png)

### Refactor code outcome for 2018
![Test Image](/Resources/VBA_Challenge_2018.png)

#### Analytical outcome: 
- **Based on 2017 data**: As per the analysis of 2017 stocks of the 12 green energy corporations it is observed that all the stocks provided good returns except from the TERP stock which provide negative (-7.2%) return. Based on this analysis stock SPWR has the highest Total Daily Volume of $782,187,000 and 184.5% return whereas the highest percentage return is on DQ stock with 199.4%, but its Total daily volume is only $35,796,200 which is the least amount other stocks. Based on this analysis SPWR stock is the more benificial stock if Steven parents are looking at investing all their money in one stock. But we also need to see 2018 results to make any recommendations on stock performance.

- **Based on 2018 data**: As per the analysis of 2018 stock market of all the green energy stock corporations experenced a low returns on their shares. The returns of DQ and SPWR fall to negetive (-62.6%) and negetive (-44.6%). The only stock that had a boom in its growth as compared to 2017 data is RUN, with $502,757100 Total Daily Volume and 84.0% return in 2018 as compared to $267,681,300 Total Daily Volume and 5.5% return. 

## Summary:
Code refactoring is a technique used to rearrange or restructure the existing code as called Editing of the code. It is a most common practice used by Data professionals or developers. The process takes by executing factoring without changing the external behavior of the code or the output of the code. The reason one uses code refactoring is to improve the nonfunctional characters of the code. It helps in removing the complecation of the code and enhancing the reliability of the code. Code refactoring also helps in eliminating vulnerabilities of the system and also removes bugs. It is a cycle of continuous enhancements of the code by using different methods to make it better and more time efficient.
### What are the advantages to disadvantages of refactoring code?

#### Advantages of Refactoring

- **Increase code flexibilty**: It gives an opportunity to the developers to always implement new techniques and functions to the code in order to have a better performance as well as to make the code flexible by using standardized functions and thus increasing the capability of the code. Programmers can refactor while they also move forward with deployment processes.

- **Maintanability**: By using refactoring technique the developer keeps refreshing the code by adding or changing the code to provide outcome in a less timely manner with more accurace and making it easy to understand or read. As well as keep the code upto date

#### Disadvantages of Refactoring

- **Chances of introducing bugs**: In case if it went wrong, you will have to waste much more time in solving the problem and there are probable chances that it may go wrong due to complexity of the code and end up stuck with the debugging of the code. 

- **Time consuming and expensive**: In real life scenerios it could be a time consuming and expensive method as instead of creating a new code we are refactoring an old code. It will include finding ways to edit and modify the code in a way that it becomes less complex without effecting the outcome, it could be tricky sometimes and might include spending hours in editing the code. Rather than starting a new one.

### How do these pros and cons apply to refactoring the orignal VBA script?
The orignal VBA script provided in this challeneg is missing few steps like creating of the output arrays, creating for loops

