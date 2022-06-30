# stock-analysis
Module 2 - VBA Scripting
# The written analysis contains the following structure, organization, and formatting:

- There is a title, and there are multiple paragraphs (2 pt).
- Each paragraph has a heading (2 pt).
- There are subheadings to break up text (2 pt).
- Links are working, and images are formatted and displayed where appropriate (2 pt).
# Analysis Requirements (12 points)
- The written analysis has the following:

## Overview of Project
- The purpose and background are well defined (2 pt).

Explain the purpose of this analysis

After completing a workbook that includes a VBA script which analyses two different worksheets containing stock data from the years 2017 and 2018 respectively, Steve (a client) wants the VBA script refactored. He specificies that the code does not run quickly enough which could be problematic for datasets involving thousands of stocks. The purpose of refactoring the code is to figure out a way to loop through all of the data one time rather than looping through the same dataset multiple times which is how the original VBA script stands.

## Results
- The analysis is well described with screenshots and code (4 pt).

Using images and examples of your code, compare the stock performance between 2017 and 2018, as well as the execution times of the origina script and the refacotred script.

The results of the VBA script times for 2017 and 2018 for both the original and refactored code is as follows:

Original - 2017

![VBA_Challenge_2017_Original](https://user-images.githubusercontent.com/107309793/176657849-85c66ae3-a49b-4252-a427-10f2d35f738a.png)

Original - 2018

![VBA_Challenge_2018_Original](https://user-images.githubusercontent.com/107309793/176657884-07a7c720-0eeb-4519-9cf2-aaa39befae71.png)

Refactored - 2017

![VBA_Challenge_2017](https://user-images.githubusercontent.com/107309793/176657940-432535c7-d718-447f-ae33-6c01ac279d78.png)

Refactored - 2018

![VBA_Challenge_2018](https://user-images.githubusercontent.com/107309793/176657990-79a1e26d-691c-4d47-98b3-14151b864e2a.png)

To quickly summarize this numerically; Original - 2017 > Original 2018 > Refactored 2017 > Refactored 2018.

In both cases, the 2017 runtime was slower than the 2018 runtime. The 2017 and 2018 worksheets do not consist of datasets of significantly differing size (they both consist of 3013 rows and 8 columns of data) so this somewhat negligible change in runtime may be simply due to computer resources. In reference to page 2.5.3: Measure Code Performance in Module 2, "The first time you run a macro, the elapsed time may be longer than subsequent runs because computer resources need to be allocated to run the macro. Once allocated, these resources are ready for subsequent runs." I was running the code for 2017 first before 2018 in each case which could indicate that my computer successfully allocated resources differently between runs.

When comparing the original code and refactored code runtimes, it's clear that the refactored code runs quicker overall for both the 2017 and 2018 dataset. This is simply due to how the code is structured.

In the original script, the code utilizes a nested loop. It directs the computer to loop through every row of data 12 different times, collecting the variables we care about (totalVolume, startingPrice, endingPrice) and then outputting the value of these variables onto a new worksheet in between each of these runs. To walk through some of the most relevant code:

An array called "tickers" is initialized to categorize each different stock ticker:

```
Dim tickers(11) As String

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
```
The startingPrice and ending Price variables are initialized (and later defined in the code).

```
Dim startingPrice As Single
Dim endingPrice As Single
```

The respective year (2017 or 2018) worksheet is activated depending on the user's input prior to running the code. "yearValue" is a variable defined by the year input within an InputBox.

`Sheets(yearValue).Activate`

The number of rows to loop over is determined and defined in the variable "RowCount".

`RowCount = Cells(Rows.Count, "A").End(xlUp).Row`

A nested loop cycles through each iterator (0 to 11) one at a time which each involve another loop which cycles through every row (2 to RowCount). The code starts with the iterator values indicated by "i" which defines the variable "ticker" and initializes the totalVolume variable to zero:

```
For i = 0 To 11
    ticker = tickers(i)
    totalVolume = 0
```
This for loop will later be completed.

The inner for loop cycles through all of the rows (2 through 3013) and calculates the totalVolume, locates the startingPrice, and locates the endingPrice for the respective "ticker" with the following code:

```
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
```
The code then outputs the ticker, totalVolume, and utilizes the startingPrice and endingPrice to calculate "return". This output is printed in a newly activated worksheet called "All Stocks Analysis" each time a new iterator is cycled through. This output is completed 12 times (0 to 11).

```
    Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = ticker
        Cells(4 + i, 2).Value = totalVolume
        Cells(4 + i, 3).Value = endingPrice / startingPrice - 1
Next i
```
Note: This walkthrough of the original script does not include the separate subroutine which formats the table of data outputted into the "All Stocks Analysis" worksheet.

Moving on to the refactored script - this script does not utilize a nested for loop. The code utilizes 4 separate for loops; one that initialize each variable to a value of 0, another to cycle through all of the rows and to calculate tickerVolumes, tickerStartingPrices, and tickerEndingPrices, another loop to output data into the "All Stocks Analysis" worksheet, and lasly a for loop built to format the data outputted into the "All Stocks Analysis" worksheet. The loop which collects the data for tickerVolumes, tickerStartingPrices, and tickerEndingPrices only has to loop through rows 2 to 3013 ONE TIME. This significantly reduces the runtime.

To break down what the code is doing:

The "tickers" array is defined in the same way as outlined above.
The corresponding dataset is activated in same fashion (with an InputBox).
The RowCount variable is the same.

A new variable defined as "TickerIndex" is set to zero. This variable is key for references the tickers array in subsequent code without having to check unnecessary rows not pertaining to the ticker of interest.

`TickerIndex = 0`

The volumes, starting/ending prices variables are now defined as arrays also involving 12 indices:

```
Dim tickerVolumes(12) As Long  
Dim tickerStartingPrices(12) As Single   
Dim tickerEndingPrices(12) As Single
```
All of these arrays are initialized to a value of zero using a for loop (0 to 11)

```
For i = 0 To 11
    tickerVolumes(i) = 0
    tickerStartingPrices(i) = 0
    tickerEndingPrices(i) = 0
Next i
```
The code loops over all of the rows in the spreadsheet once, increasing the TickerIndex by 1 once the last ticker of that index is identified (i.e. the data collected for tickers(0) is stored after moving through its row, then the data for tickers(1) is stored after going through its rows, etc.):

```
For i = 2 To RowCount
    tickerVolumes(TickerIndex) = tickerVolumes(TickerIndex) + Cells(i, 8).Value
    
    If Cells(i - 1, 1).Value <> tickers(TickerIndex) Then
        tickerStartingPrices(TickerIndex) = Cells(i, 6).Value
    End If
    
    If Cells(i + 1, 1).Value <> tickers(TickerIndex) Then
        tickerEndingPrices(TickerIndex) = Cells(i, 6).Value
        TickerIndex = TickerIndex + 1
    End If           
Next i

The above block of code starts by summing the tickerVolumes found in the dataset and captures this value for a TickerIndex of 0. The tickerStartingPrices and tickerEndingPrices are then located for TickerIndex = 0, the TickerIndex then increases to TickerIndex = 1 once the final ticker in that group is identified. This process is repeated as it cycles through the rest of the rows.

The "All Stocks Analysis" worksheet is activated only once and all of the data collected from the previous for loop is outputted at once using the following code:

```
For i = 0 To 11
    Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
Next i
```
Lastly, this code includes a formatting block which is included in the overall runtime for the script.

```
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
```

## Summary
- There is a detailed statement on the advantages and disadvantages of refactoring code in general (3 pt).

What are the advantages or disadvantages of refactoring code?

- There is a detailed statement on the advantages and disadvantages of the original and refactored VBA script (3 pt).

How do these pros and cons apply to refactoring the original VBA script?
