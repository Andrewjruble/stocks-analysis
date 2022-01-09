# **Stock Analysis using VBA/Excel**
## Overview
A code in Virtual Basic for Applications (VBA)/Excel was written for a friend, Steve to easily use when analyzing stocks for his parents. The code was effective for the small amount of stocks Steve was trying to investigate. However, it may not work or be too time consuming when looking at larger data sets or several more stocks. It was decided to refactor the code to make it more efficient, making it more versatile and able to handle bigger investigations. While improving the time it takes to run the macro was the main goal, one could also refactor code to consume less memory or make the logic easier to follow for future users.
## Results
To make the code more efficient, the original 'nested for loops' were replaced with arrays that recognize the ticker symbol of desired stocks. In doing so, the macro would no longer look through the data several times. This in theory would make it run more quickly. Comparing the original code to the refactored version, one can see the changes that were made: 

### Original Code

```   '4) Loop through tickers
       For i = 0 To 11
       ticker = tickers(i)
       totalVolume = 0
       
       '5) loop through rows in the data
          Sheets(yearValue).Activate
       
       For j = 2 To RowCount
       
         '5a) Get total volume for current ticker
           If Cells(j, 1).Value = ticker Then
           totalVolume = totalVolume + Cells(j, 8).Value
         End If
           
         '5b) get starting price for current ticker
           If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
           startingPrice = Cells(j, 6).Value
         End If

         '5c) get ending price for current ticker
            If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
            endingPrice = Cells(j, 6).Value
          End If
       Next j 
```
The original code runs through the entire data set for each individual variable. Essentially the code is getting ran 12 times. The ticker is assigned, the macro runs through the data to find the cells containing the ticker. The initial total volume is created when the first ticker is found. The for loop then creates a running total by adding the quantity of the next row containing the ticker. This continues until the ticker changes to a different variable. Starting price and ending price are found by discovering the first and last ticker in the sequence also using for loops. Once the macro determines there is a change in the stock symbol, it can establish the starting and ending prices. This worked well for the size of the data set and the amount of tickers Steve wanted to look into, but eliminating the macro having to repeat itself over and over would speed up the process on bigger projects. 

### Refactored Code
```

'1a) Create a ticker Index. Set it to zero before iterating over the rows
        tickerindex = 0
            
    '1b) Create three output arrays
        Dim tickerVolumes(11) As Long
        Dim tickerStartingPrices(11) As Single
        Dim tickerEndingPrices(11) As Single
    
       
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
        For tickerindex = 0 To 11
            tickerVolumes(tickerindex) = 0
    
        
    ''2b) Create a for loop that will loop over all the rows in the spreadsheet.
        For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
            tickerVolumes(tickerindex) = tickerVolumes(tickerindex) + Cells(i, 8).Value
        
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
            If Cells(i, 1).Value = tickers(tickerindex) And Cells(i - 1, 1).Value <> tickers(tickerindex) Then
            tickerStartingPrices(tickerindex) = Cells(i, 6).Value
            
            End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
                If Cells(i, 1).Value = tickers(tickerindex) And Cells(i + 1, 1).Value <> tickers(tickerindex) Then
                tickerEndingPrices(tickerindex) = Cells(i, 6).Value
                End If
            
            
        '3d Increase the tickerIndex.
            If Cells(i, 1).Value = tickers(tickerindex) And Cells(i + 1, 1).Value <> tickers(tickerindex) Then
            tickerindex = tickerindex + 1
            End If
            
            Next i
            
        '4)Use a for through your arrays to output the “Ticker,” “Total Daily Volume,” and “Return” columns
            For i = 0 To 11
            
            Worksheets("All Stocks Analysis").Activate
            Cells(4 + i, 1).Value = tickers(i)
            Cells(4 + i, 2).Value = tickerVolumes(i)
            Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
            
            Next i
            Next tickerindex
```
By creating arrays on the refactored version, the volumes, starting prices and ending prices could be established for all applicable tickers in one run. Eliminating the having to run through the data one time each individual ticker. Arrays allow for the ticker to change once the the last variable in the data set is found, instead of having to completely run through everything before changing. The variable `tickerindex' is able to able change as soon as the last applicable row is found. 

### Runtimes of codes before and after refactoring

#### Original Version

![Original 2017](https://github.com/Andrewjruble/stocks-analysis/blob/main/Resources/VBA_Challenge_Orig2017.PNG)
![Original 2018](https://github.com/Andrewjruble/stocks-analysis/blob/main/Resources/VBA_Challenge_Orig2018.PNG)

#### Refactored 

![Refactored 2017](https://github.com/Andrewjruble/stocks-analysis/blob/main/Resources/VBA_Challenge_2017.PNG)
![Refactored 2018](https://github.com/Andrewjruble/stocks-analysis/blob/main/Resources/VBA_Challenge_2018.PNG)

Looking at the times of the original vs the refactored version, both years ran about a half second faster on the latter. 

## Summary

There can be many advantages to refactoring code, simplyfying it or improving the logic can:

       - Make the code run faster
       - Make it more efficient and take up less memory
       - Make it easier to follow to future users
 
 However, refactoring can have disadvantages and may not always be worth it do. The code could already be very simple, useful and easy to understand. There's a chance hours could be wasted trying to improve things only end up with minor or no improvements.
 
 When looking at our code, it could be argued the value is subjective. It was proven the newer version runs faster and could be theorized that it takes up less memory. This would an advantage if this was applied to a larger data set, or if several tickers needed to be evaulated. In that instance, getting results may crash the program or take too long using the original code. 
       



       



