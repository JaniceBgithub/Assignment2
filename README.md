# Assignment2

## Overview of the Project

This project involved analysis of stock data from two years, 2017 and 2018.  Twelve stocks were grouped together and sorted by date in an excel spreadsheet.  VBA code was written in two ways:
1. Loop through the entire file 12x for each stock ticker and determine the starting price, ending price and total stock volumes over the time period.  The return was calculate by dividing ending price by starting price. 
2. Similar to above, but the file was refactored to only go through the file one time for all the data. 

## Results

### Loop through the code 12x for each stock ticker

The first analysis went through the code 12x, once for each stock ticker

The first step was to create a header row for the output worksheet: 

    Cells(3, 1).Value = "Year"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
    
Tickers and variables were declared and initialized:
    
     '2 initialize tickers
    
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
    
    '3a initliaze variables for starting and ending price
    
    Dim startingPrice As Single
    Dim endingPrice As Single
    
We then used a line of code to determine the end of the data: 

    RowCount = Cells(Rows.Count, "A").End(xlUp).Row  
    
We then looped over all the rows for each stock ticker and added up total volumes and grabbed starting and ending price.  Data was also output into a worksheet: 
   
     For i = 0 To 11
        ticker = tickers(i)
        totalVolume = 0
        
        '5 loop through rows in data
        Worksheets(yearValue).Activate
        For j = 2 To RowCount
        
        'total volume is in column 8 in 2018 sheet
        'Cells function is row then column
            If Cells(j, 1).Value = ticker Then
                totalVolume = totalVolume + Cells(j, 8).Value
            End If
        
        'start price for current ticker
            If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                startingPrice = Cells(j, 6).Value
            End If
        
        'end price for current ticker
            If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                endingPrice = Cells(j, 6).Value
            End If
                   
        Next j
   
        'Output data for current ticker
    
        Worksheets("AllStocksAnalysis").Activate
        Cells(4 + i, 1).Value = ticker
        Cells(4 + i, 2).Value = totalVolume
        Cells(4 + i, 3).Value = (endingPrice / startingPrice - 1) * 100
    
     Next i

A timer was also added to determine run speed. 

An additional macro was also created to format output.  The cells were shaded red or green depending on positive or negative returns.

  Sub formatAllStocksAnalysisTable()

    Worksheets("AllStocksAnalysis").Activate
    
    'Formatting
    Range("A3:C3").Font.Bold = True
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("A4:A15").Font.Color = vbRed
    Range("A3:C3").Font.Size = 12
    Range("B4:B15").NumberFormat = "#,##0"
    Columns("B").AutoFit
    
    dataRowStart = 4
    dataRowEnd = 15
    For i = dataRowStart To dataRowEnd
    
        If Cells(i, 3) > 0 Then
            Cells(i, 3).Interior.Color = vbGreen
    
        ElseIf Cells(i, 3) < 0 Then
            Cells(i, 3).Interior.Color = vbRed
        
        Else
            Cells(i, 3).Interior.Color = xlNone
        End If
        
    Next i
    
    
End Sub


### Loop through the data only one time

This macro was similar to the above.  The main change was to use a ticker index to put results into result arrays for volumes, start price and end price. 

The main code change was in this section.  Note the ticker index increase at the end to increment the arrays by 1 value as new stock tickers are encountered in the spreadsheet:

    For i = 2 To RowCount
        
        'increment stock volumes as long as the tickers in column 1 match with previous row
   
        If Cells(i + 1, 1).Value = Cells(i, 1).Value Then
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        End If
        
        'Grab the value for starting stock price when the current row does not match with the previous row
        'ie a new stock ticker
        
        If Cells(i - 1, 1).Value <> Cells(i, 1) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        End If
        
        'Grab the value for ending stock price when the next ticker value does match current row
        'increment ticker index when go onto a new ticker
        
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            tickerIndex = tickerIndex + 1
        End If
                     
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
        
    'there are twelve total stock tickers
    For i = 0 To 11
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        
        'this the return
        
        Cells(4 + i, 3).Value = (tickerEndingPrices(i) / tickerStartingPrices(i) - 1)
                
    Next i


### Code Speed

Running through the code 12x is a bit slow - on my computer it took approximately 0.5 seconds per run.  After the code was refactored to only run through the data once, the speed to run was about 0.1 seconds.

![2017 Refactored run speed](https://github.com/JaniceBgithub/Assignment2/blob/master/VBA_Challenge_2017.png)

## Summary

Advantages of refactoring code include: 
- The main advantage is time savings 
- An additional advantage is re-use of code that we already know works. This should cut down on time for testing. 

Disadvantages of refactoring code include: 
- The program may have been poorly written to start with or poorly documented.  It may just be faster to start new code for a smaller project. 
- The old code needs to be verified for accuracy before using which may be time consuming. 

## General code improvements

The following could be done to make the code even better: 
- The manual input of the stock tickers into the code is not that good - time consuming and error prone and also does not allow for changes in the excel sheet.  This section could be modified to look for new stock tickers as the code progresses through the spreadsheet.  
- The code requires that the excel file be grouped by stock ticker and in ordered date.  The code could be modified so that neither of these are required. 
