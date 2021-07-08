# VBA of Wall Street

## Overview of Project

#### This project had two purposes: 1. To gain experience with VBA and Excel and 2. To analyse the Wall Street data to give Steve some insight on some green stocks for his parents investments. The new VBA experience showed how to refactor code to make it more efficient and automated.

## Results

#### The data for these analyses are a list of Stock Tickers with daily values for Open, High, Low, Close, and Volume for 2017 and 2018. Steve wants to get some aggregated data for all stocks to compare the performance. 

#### The data for these analyses are lists of Stock Tickers with daily values for Open, High, Low, Close, and Volume for 2017 and 2018. Steve wants to get some aggregated data for all stocks to compare the performance of various green stocks. The analyses used refactored code from the Module 2 steps. The performance time between years and code versions is also compared.

### The macro AllStocksAnalysisRefactored() utilized Cell assignments to create the table for the results.

```VB
    Worksheets("All Stocks Analysis").Activate
    
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
```
### Arrays were used to hold all the ticker stock abbreviations as well as the total volume and starting and ending prices. They were declared with the appropriate dat type.

```VB
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
    
    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
```

### Nested For loops were used to cycle through the rows of Tickers and the columns of daily stack metrics. For the refactored code, the variable "tickerIndex" was used instead of a simple single letter.

```VB
 '2a) Create a for loop to initialize the tickerVolumes to zero.
    For tickerIndex = 0 To 11
        'Use tickerIndex to reference arrays
        ticker = tickers(tickerIndex)
        tickerVolumes(tickerIndex) = 0
```

### The values in the arrays were updated within the nested loops. The Starting and Ending prices were determined by comparing the ticker value with the previous or next value. Text replacement was also used to make the Cells references in the loop code more readable. 

```VB
    '2b) Loop over all the rows in the spreadsheet.
        For j = 2 To RowCount
    
            'use these instead of Cells references
            num = Cells(j, 1).Value
            numprev = Cells(j - 1, 1).Value
            numnext = Cells(j + 1, 1).Value
            
    '3a) Increase volume for current ticker
            If num = ticker Then
                tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(j, 8).Value
            End If
            
    '3b) Check if the current row is the first row with the selected tickerIndex.
            If num = ticker And numprev <> ticker Then
                tickerStartingPrices(tickerIndex) = Cells(j, 6).Value
            End If
        
    '3c) check if the current row is the last row with the selected ticker
            If num = ticker And numnext <> ticker Then
                tickerEndingPrices(tickerIndex) = Cells(j, 6).Value
            End If
        Next j
```
### I used tickerIndex as the looping variable for the outer loop (see above) so in the instructions where it said to increase the tickerIndex, the "Next tickerIndex" code line does that automatically.

```VB
 For tickerIndex = 0 To 11
.
.
.
.
Next tickerIndex
```
### The final values for volume, start and end price were entered into their arrays within a loop.

```VB
    For i = 0 To 11

        Worksheets("All Stocks Analysis").Activate

        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = (tickerEndingPrices(i) / tickerStartingPrices(i)) - 1

    Next i
```
### The data were formatted into the Excel spreadsheet with color indicators for percent increase (Green) or decrease (Red). A process time was opened at the beginning of the code and closed at the end of the code and a message popup box with the process time for each year was displayed. The biggest difference in the refactored code was the use of arrays to store the data instead of displaying the data immediately to the final spreadsheet.