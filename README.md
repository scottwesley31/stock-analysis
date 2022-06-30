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

When comparing the original code and refactored code runtimes, it's clear that the refactored code runs quicker overall for both the 2017 and 2018 dataset. This is simply due to how the code is structured. In the original script, the code utilizes a nested loop. It directs the computer to loop through every row of data 12 different times, collecting the variables we care about (totalVolume, startingPrice, endingPrice) and then outputting the value of these variables onto a new worksheet in between each of these runs. To walk through some of the most relevant code:

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
The code then outputs the ticker, totalVolume, and utizes the startingPrice and endingPrice to calculate "return". This output is printed in a newly activated worksheet called "All Stocks Analysis" each time a new iterator is cycled through. This output is completed 12 times (0 to 11).

```
    Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = ticker
        Cells(4 + i, 2).Value = totalVolume
        Cells(4 + i, 3).Value = endingPrice / startingPrice - 1
Next i
```

## Summary
- There is a detailed statement on the advantages and disadvantages of refactoring code in general (3 pt).

What are the advantages or disadvantages of refactoring code?

- There is a detailed statement on the advantages and disadvantages of the original and refactored VBA script (3 pt).

How do these pros and cons apply to refactoring the original VBA script?
