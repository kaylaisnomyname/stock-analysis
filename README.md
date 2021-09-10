# VBA Challenge
## Overview of Project 
During this module we created a VBA script to analysis the performance of a dozen stocks in 2017 and 2018. Now we want to refactor the original code, so that the enhanced version can run faster.   

## Results

The analysis results are shown in the following images:
- Image 1: Overall performance for 2017
- Image 2: Overall performance for 2018  
- Image 3: Execution time of refactored code for 2017
- Image 4: Execution time of refactored code for 2018  

### Stock Performance  
As shown in Image 1 and 2, the 12 target stocks performed better in 2017 comparing to 2018. In 2017, 11 out of this 12 stocks had positive returns; only one stock, TERP, had a drop of 7.2%. The best performed stock in 2017 was DQ, with a return of 199%. However in 2018, only two out of these stocks had  positive returns, while the rest of them all had negative returns. There are two stocks remain raising in the market: ENPH and RUN, both had gains in 2017 and more than 80% gains in 2018. Thus, based on the given stock data, the competitive stocks would be ENPH and RUN.  


##### Image 1
![Execution time for 2017, old code](https://github.com/kaylaisnomyname/stock-analysis/blob/main/executionTime2017.png?raw=true)

#### Image 2
![Execution time for 2018, old code](https://github.com/kaylaisnomyname/stock-analysis/blob/main/executionTime2018.png?raw=true)  


### Execution Performance
1. Execution time  
The execution time of the original code for 2017 is 1.4219s and for 2018 is 1.8281s, as shown in Image 1 and 2. The execution time of the refactored code for 2017 is 0.2891s and for 2018 is 0.3281s, as shown in Image 3 and 4. In short, the refactored code runs much faster than the original code. 

#### Image 3 
![Execution time for 2017, refactored code](https://github.com/kaylaisnomyname/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png?raw=true)

#### Image 4
![Execution time for 2018](https://github.com/kaylaisnomyname/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png?raw=true)

 
2. Code Structure  
The main differences between the original code and the refactored script are their code structures. The original code has bigger chunk of functional block, using nested For loops to iterate through both ticker index and row number to analysis data and also to output to result. Section of original code shown below(some codes to get data are omitted in the quote as they're in too details that is not relavent to this statement. Same as that for refactored code quote). 
``` 
  '4) loop through the tickers
    For i = 0 To 11              ' 1. 1st loop to loop through the ticker array
        ticker = tickers(i)
        totalVolume = 0
        ' 5) loop through rows in data
        Worksheets(yearValue).Activate
        For j = 2 To RowCount   ' 2. nested 2nd loop through the rows to get volumes, starting price and ending price
        
           ... 'omit actions to get results

         Next j
            
          ' 6) output data for current ticker   '3.  output result during looping
            Worksheets("All Stocks Analysis").Activate
            Cells(4 + i, 1).Value = ticker
            Cells(4 + i, 2).Value = totalVolume
            Cells(4 + i, 3).Value = endingPrice / startingPrice - 1
           
        Next i
```


2. Instead, the refactored code breaks this big chunk of functional process into small blocks. First it initializes arrays to hold results. Secondly, it uses a for loop iterating through rows to aquire analysis data and store the results in the result arrays. After all the results are loaded, it outputs the result in another loop.  Snapshot of the code is shown below:
```
    tickerIndex = 0                   '1. initializes tickerIndex as zero, this index is for one of the four output arrays
    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
      
    ...         ' omit some actions here 
    
    ''2b) Loop over all the rows 
    For i = 2 To RowCount              '2. single loop to iterate rows and get results
    
        ....    ' omit actions to aquire all analysis data
        tickerIndex = tickerIndex + 1   ' here it imcrements the tickerIndex
    Next i
    '4) Output result
    For i = 0 To 11
        Worksheets("Sheet2").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
    Next i    
```


## Summary 

- Refactoring is to reconstruct the original code, breaking the original large chunk of rundown into smaller functioning blocks, aiming to enhance efficiency. The most emplicit advantage of refactoring is to make the code run faster. Yet its limitation is obvious too. Since refactoring only restructuring the original code, it does not introduce new behaviour to the original. In other words, the refactored code will get the same analysis results as the original code.  
- For instant, the VBA script used in this module to analysis the target 12 stocks in given years. The original code takes about 1.5s to finish execution. The refactored code takes only about 1/5 of the original time and gets the same results. Now if the task were to analysis a larger stock dataset instead of the target 12 stocks, because the refactored code, same as the original one, only analyzes the given 12 stocks, a new VBA script, or a new version of it, would be needed. 
