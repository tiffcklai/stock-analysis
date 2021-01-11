# Stock Analysis Challenge with Excel and VBA

## Overview of Project
Data analysis on green energy stocks with respect to their Total Daily Volume and Return. Refactoring the code to be more efficient when expanding the dataset to analyze the entire stock market.

## Results 
### Analysis
**Preparing the data for the code**

To expand the dataset to allow the entire stock market to be analyzed, a plan was mapped out prior to any coding. The pre-existing code was evaluated to determine what needed to be kept and what was required to be changed or added to provide the deliverable. The instructions for the refactoring process was then entered as comments into the code to keep structure and provide a data trail for the additions to the code. The updated code and instructions is seen below. 

``` VBA

Sub AllStocksAnalysisRefactored()

    Dim startTime As Single
    Dim endTime  As Single
    
    yearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
    
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

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
        Dim tickerindex As Integer
        tickerindex = 0       

    '1b) Create three output arrays
    
        Dim tickerVolumes(12) As Long
        
        Dim tickerStartingPrice(12) As Single
        
        Dim tickerEndingPrice(12) As Single  
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
        
        For i = 0 To 11
            
            tickerVolumes(i) = 0
            tickerStartingPrice(i) = 0
            tickerEndingPrice(i) = 0
        
        Next i
        
    ''2b) Loop over all the rows in the spreadsheet.
         
         For i = 2 To RowCount  
    
        '3a) Increase volume for current ticker
            tickerVolumes(tickerindex) = tickerVolumes(tickerindex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
            If Cells(i, 1).Value = tickers(tickerindex) And Cells(i - 1, 1).Value <> tickers(tickerindex) Then
                tickerStartingPrice(tickerindex) = Cells(i, 6).Value
            End If
            
        'End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
            If Cells(i, 1).Value = tickers(tickerindex) And Cells(i + 1, 1).Value <> tickers(tickerindex) Then
                tickerEndingPrice(tickerindex) = Cells(i, 6).Value
            End If
            
            '3d Increase the tickerIndex.
            If Cells(i + 1, 1).Value <> tickers(tickerindex) And Cells(i, 1).Value = tickers(tickerindex) Then
                tickerindex = tickerindex + 1            
            
        'End If
            End If
        Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrice(i) / tickerStartingPrice(i) - 1
        
    Next i
    
    'Formatting
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit

    dataRowStart = 4
    dataRowEnd = 15

    For i = dataRowStart To dataRowEnd
        
        If Cells(i, 3) > 0 Then
            
            Cells(i, 3).Interior.Color = vbGreen
            
        Else
        
            Cells(i, 3).Interior.Color = vbRed
            
        End If
        
    Next i
 
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub

```
## Results from data analysis

From the analysis, 11 of the 12 green energy stocks had a positive return percentage in 2017 and only 2 of 12 green energy stocks had a positive return percentage in 2018. This can be shown based on the number of green (positive return) and red (negative return) cells in the Return column. The performance of the stocks in 2017 were significantly better than the stock performance in 2018. 

Refactoring the code has shown improved code performance. In the original code, the time in seconds for analysis was 0.64 for 2017 and 2018 datasets. In the refactored code, the analysis time was 0.22 for 2017 and 0.13 in 2018. Thus, proving that the refactored code is more efficient.

**All Stocks Performance**

![All Stocks 2017 Performance](https://github.com/tiffcklai/stock-analysis/blob/main/All%20Stocks%202017%20Performance.png?raw=true)

![All Stocks 2018 Performance](https://github.com/tiffcklai/stock-analysis/blob/main/All%20Stocks%202018%20Performance.png?raw=true)

**Original Code Performance**

![Original 2017 Code Performance](https://github.com/tiffcklai/stock-analysis/blob/main/Original%202017%20Code%20Performance.png?raw=true)

![Original 2018 Code Performance](https://github.com/tiffcklai/stock-analysis/blob/main/Original%202018%20Code%20Performance.png?raw=true)

**Refactored Code Performance**

![Refactored 2017 Code Performance](https://github.com/tiffcklai/stock-analysis/blob/main/Refactored%202017%20Code%20Performance.png?raw=true)

![Refactored 2018 Code Performance](https://github.com/tiffcklai/stock-analysis/blob/main/Refactored%202018%20Code%20Performance.png?raw=true)

## Summary 
### Advantages and Disadvantages of Refactoring Code

The advantages of refactoring code is to provide increased readability and decrease the complexity of the code. Thus, optimizing memory and speed of code performance. The disadvantages of refactoring code is the increased probability of introducing bugs into the pre-existing code. 

### Advantages and Disadvantages of Original and Refactored Code

The benefits of refactoring the original code was the shortened code performance as shown above when comparing the time it takes for the code to run. Additionally, the refactored code allows for an expanded dataset to be analyzed while running more efficiently as well. The disadvantages of the refactored code include the difficulty of troubleshooting the code and having to ensure that any code added needs to work with the pre-existing code.
