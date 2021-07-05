# VBA_Challenge
Challenge Module2 - VSB_Challenge

## Overview of Project

### Background

>Steve wants to analyze an entire dataset for doing a little more research for his parents, he wants to expand the dataset to include the entire stock market over the last few years at the click of a button.

### Purpose

Via refacoring the Code to loop through all the data one time in order to collect the information that Steeve needs to compare the stock performance betweeen 2017 and 2018 at a click of a button.

There were measured two differents executions:

- *`Original Script:`* it's supposed to take a long time to execte.
- *`Refactored Script:`* it takes fewer steps and make a code more efficient.


## Results

For both, 2017 & 2018 anaysis, it was developed a visual basic coding for the data set, where the desired outcomes are: (**Ticker**, **Total Daily Volume** and **Return**)

Arrays creation sintaxis was looked at: [Visual Basic](https://docs.microsoft.com/en-us/dotnet/visual-basic/programming-guide/language-features/arrays/)

```
'1a) Create a ticker Index
    tickerIndex = 0

    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
            
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For x = 0 To 11
        tickerVolumes(x) = 0
    Next x
        
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
        ticker = tickers(tickerIndex)
        '3a) Increase volume for current ticker
            
        If Cells(i, 1).Value = ticker Then
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        End If
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
        If Cells(i - 1, 1) <> ticker And Cells(i, 1).Value = ticker Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        End If
                    
        'End If
        
        '3c) check if the current row is the last row with the selected ticker
        
        If Cells(i + 1, 1) <> ticker And Cells(i, 1) = ticker Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            '3d Increase the tickerIndex, to reuse code decided to increment tickerindex in here
            tickerIndex = tickerIndex + 1
        End If
    
    Next i
```

### Original vs Refactored Code Time Elapsed

|           | 2017  | 2018  |
| :------------ |:---------------:| :-----:|
| Original Code      | ![Original Code 2017](https://user-images.githubusercontent.com/86028032/124408605-2d0b6680-dd0c-11eb-9e7b-fcf882898437.PNG) | ![Original Code 2018](https://user-images.githubusercontent.com/86028032/124408626-385e9200-dd0c-11eb-9d0c-d069ebcae332.PNG) |
| Refactored Code      | ![Refactored Code 2017](https://user-images.githubusercontent.com/86028032/124408642-3eed0980-dd0c-11eb-97da-f0f3265650ae.PNG)        |   ![Refactored Code 2018](https://user-images.githubusercontent.com/86028032/124408647-444a5400-dd0c-11eb-8aa1-d260c66ff13b.PNG) |




### - 2017 Stock performance analysis

2017 Stock performance, had a great beahiour in almost all Tickers, since must of them had return values over 0%. These percentage values indicate that the ending price for each ticker was above the initial values.

![2017 Stock Analysis](https://user-images.githubusercontent.com/86028032/124404981-cfbee780-dd02-11eb-9cad-aa2272722273.PNG)


### - 2018 Stock performance analysis
In the other hand, 2018 Stock performance wasn't good, since the stocks had lost their values.

![2018 Stock Analysis](https://user-images.githubusercontent.com/86028032/124404993-de0d0380-dd02-11eb-8d86-2ac0c71d2fea.PNG)


## Summary

- Avantages and disadvantages of refactoring code
  - To refactored a code requires to understand what is needed in the project.
  - There are a lot of codes in the internet that could be helpful, but it doesn't mean that is useful for what we are looking. So it may take a lot of time to find a code that is useful for the application that is required.
  - A code that can deploy the desired data could be found, but refactoring it may take more time than writing a code from zero.
  - Refactoring code supposed to be more efficient, if know how to.

- Advantages and disadvantages of the original and refactored VBA script
  - The refactored VBA script takes less time to run
  - The refactored VBA script uses less memory to run
  - The refactored VBA script has less coding steps, so is more efficient
