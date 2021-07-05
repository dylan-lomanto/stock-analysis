# Stock Analysis

### Overview

The client, Steve, has requested a refactored version of the originally presented code, that is efficient enough to analyze the entire stock market in a reasonable amount of time.  The original code examines tables for 2017 and 2018 for 12 green stocks that displays daily trade counts and a variety of daily pricing categories. For our purposes the relevant fields are ticker, closing price, and volume. The original code outputs a table for either year with the ticker, total daily volume, and annual return percentage.  The total daily volume is acquired by summing the daily volumes for each ticker. The return percentage is calculated by dividing the final daily closing value by the first daily closing value for each ticker.  The original code uses nested for loops to return the data. The refactored code eliminates the need for nested loops buy pulling all of the relevant information on the first pass through the data set.

### Results

#### 2017 to 2018 Comparison
 
Green stocks returned a significantly higher percentage in 2017 vs 2018.  In 2017 11 of the 12 stocks returned a positive value, while only two did so in 2018.  ENPH and RUN were the only tickers with a positive return both years. ENPH had higher total daily volumes and average returns over the course of 2017 and 2018, making it the most attractive stock of those examined.

![2017_Table](https://user-images.githubusercontent.com/86164867/124499088-46df9480-dd72-11eb-81e4-fcbab7b5fc43.PNG)

![2018_Table](https://user-images.githubusercontent.com/86164867/124499243-8d34f380-dd72-11eb-9680-e40eed19104b.PNG)

#### Refactored Code vs Original Code Execution Times

The refactored code is significantly more efficient than the original code, and thus executes much faster. The original code uses nested for loops to draw the information from the dataset. This is inefficient because the code only pulls one piece of information each time it loops over the dataset. The refactored code is able to pull all information from the first loop.

##### Original Code

    Dim startTime As Single
    Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")
    
        startTime = Timer

    Worksheets("All Stocks Analysis").Activate

    Range("A1").Value = "All Stocks (" + yearValue + ")"

    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
    
    'Create ticker array
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


    Dim startingPrice As Single
    Dim endingPrice As Single
    


    Worksheets(yearValue).Activate
    

    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
   

    For i = 0 To 11
        ticker = tickers(i)
        totalVolume = 0
        
        
        Worksheets(yearValue).Activate
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

    Worksheets("All Stocks Analysis").Activate
    Cells(4 + i, 1).Value = ticker
    Cells(4 + i, 2).Value = totalVolume
    Cells(4 + i, 3).Value = endingPrice / startingPrice - 1
    
    Next i

    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)


    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.Bold = True
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.00%"
    Columns("B").AutoFit

    dataRowStart = 4
    dataRowEnd = 15
    For i = dataRowStart To dataRowEnd

        If Cells(i, 3) > 0 Then

        'Change cell color to green
        Cells(i, 3).Interior.Color = vbGreen
        
        ElseIf Cells(i, 3) < 0 Then

        'Change cell color to red
        Cells(i, 3).Interior.Color = vbRed
        
        Else

        'Clear the cell color
        Cells(i, 3).Interior.Color = xlNone
    
        End If

    Next i

    Worksheets("2018").Activate
    Range("C2:F3013").NumberFormat = "$0.00"
   


