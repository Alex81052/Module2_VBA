# Module2_VBA

## Overview of Project

I was asked to help out Steve by providing him different Stock Ticker Analysis in order to show his parents alternative investment options and simplifying the review 
by automating tasks and formatting the data, so it shows the key metrics requested in a visually pleasing manner. 

### Purpose

The purpose of this challenge was to understand for loops, adding buttons, and running different arrays in order to automate the stock analysis for Steve, 
so he could show his parents how different stocks performed based upon their returns from a specific year, and the total volume of shares traded. 

## Analysis

The first view we ran was to look into the Stock Ticker Performance from 2017 and to create a timer to show how long the code took to run which was .87seconds, 
the worst performing stock was TERP because it has -7% Return and DQ was the best because it had a 199% Return.    

![Image 1](https://github.com/Alex81052/Module2_VBA/blob/main/Resources/VBA_Challenge_2017.png)

The second view we ran was to look into the Stock Ticker Performance from 2018 and to create a timer to show how long the code took to run which was .87seconds, 
the worst performing stock was DQ because it had a -63% Return and RUN was the best because it had a 84% Return. 

![Image 1](https://github.com/Alex81052/Module2_VBA/blob/main/Resources/VBA_Challenge_2018.png)

## Results

Sub AllStockAnalysisRefactored()


Worksheets("All Stocks Analysis").Activate

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
    tickerIndex = 0
    

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
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
        If Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
            
        End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
        If Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            
        

            '3d Increase the tickerIndex.
            
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

The advantages of refactoring the code definitely simplified the assignment because I did not have start from scratch. Minor additions for 1a, 1b, 2a, 2b, and 3a helped enhance the data. 

The disadvantage was ensuring that End if were placed correctly in the code in order for the code to run the way the assignment wanted it to be run.   

I believe that the pro of refactoring the code added another building block on the Stock Analysis by adding different
variables to loop for, so there was more information provided to Steve's parents without adding much time to run extra analysis. 

I believe that a con of refactoring the code is that you must ensure you 