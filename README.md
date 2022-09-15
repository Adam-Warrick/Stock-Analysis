# **Stock-Analysis**

## **Overview of the Project**
 Our client would like to run an analysis on stock tickers to determine which stock has the best return for his parents. To complete this project I used VBA Macros to run the analysis. The first set of script worked well, however, we refactored the code to make it run more efficient and faster. Client would also like buttons in order to run the code quickly. 
 
## **Results**
The first version of code gave me the answers I was looking forever, however, it was a little slow. When I refactored the code, it made a staggering improvement on time to answer. We needed to refactor in this case as we want to expand our data search. If we did not refactor and added more tickers this would be a very slow macro which creates inefficiencies. Nested loops seem to slow the code down as there is more for the computer to read/do. Using arrays really helped speed things up when we added that to the refactored coding. Module 1 is original code and Module 2 is refactored code within the Worksheet> VB. As you can see below, the first script took over 5 seconds while the refactored code took a fraction of a second!

**Original Spreadsheet with Time Clock**

![2018 Analysis- Original Script](https://github.com/Adam-Warrick/Stock-Analysis/blob/main/2018%20Analysis%20-%20Original%20Script.png)

**Original Code**

![2018 Orginal Dim Code](https://github.com/Adam-Warrick/Stock-Analysis/blob/main/2018%20Original%20Dim%20Code.png)

**Refactored Spreadsheet with Time Clock**

![2018 Analysis - Refactored Script](https://github.com/Adam-Warrick/Stock-Analysis/blob/main/2018%20Analysis%20-%20Refactored%20Script.png)

**Refactored Code**

![2018 Recap - Arrays](https://github.com/Adam-Warrick/Stock-Analysis/blob/main/2018%20Recap%20-%20Arrays.png)

## **Summary**
Nested Loops and not using arrays are the culprits to slower code. In this case, it is proven that Refactoring is the way to go in order to be more efficient.

###### Advantages
Our refactored code ran quicker and more efficient then our previous original code. This is due to not having Nested Loops within the code which made the macro run slower initially. 

###### Disadvantages
Some disadvantages for the original code was that it was slow, clunky, and inefficient. The refactored code seemed slightly harder to build as a deeper understanding of the variables was needed. The hard part around writing efficient code is that your first pass may not always be the best result...with this, you need to dig deeper and think outside the box (Refactoring) to gain a better result. 

# **All Coding Below**

## **Original Code**

Sub AllStocksAnalysis()

'1) Format the output sheet on the "All Stocks Analysis" worksheet.

    Sheets("All Stocks Analysis").Activate
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
    'Start Timer
    Dim startTime As Single
    Dim endTime  As Single
    startTime = Timer
    
    'Ask question of user
    yearValue = InputBox("What year would you like to run the analysis on?")
    
    'Create Header Row
    Cells(3, 1).Value = "Year"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
    
'2) Intialize an array of all tickers.
    
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
  
'3) Initialize variables for the starting price and ending price.

    Dim startingPrice As Double
    Dim endingPrice As Double
    
'3A) Activate the data worksheet.

    Worksheets(yearValue).Activate
       
    
'3B) Find the number of rows to loop over.

    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
'4) Loop through the tickers

        For i = 0 To 11
            ticker = tickers(i)
            totalVolume = 0

'5) Loop Through rows in the data
    
    Worksheets(yearValue).Activate
    
        For j = 2 To RowCount

'5A) Find total volume for the current ticker.
    
        If Cells(j, 1).Value = ticker Then
            
            totalVolume = totalVolume + Cells(j, 8).Value
            
        End If
    
'5B) Find starting price for the current ticker.
        
        If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
        
            startingPrice = Cells(j, 6).Value
            
        End If
    
'5C) Find ending price for the current ticker.
    
        If Cells(j + 1, 1).Value <> ticker And Cells(j, 1) = ticker Then
            
            endingPrice = Cells(j, 6).Value
            
        End If
        
    Next j
    
'6) Output the data for the current ticker.
    
    Sheets("All Stocks Analysis").Activate
    
    Cells(4 + i, 1).Value = ticker
    Cells(4 + i, 2).Value = totalVolume
    Cells(4 + i, 3).Value = endingPrice / startingPrice - 1

Next i

'Formatting

Worksheets("All Stocks Analysis").Activate

Range("A3:C3").Font.Bold = True

Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous

Range("B4:B15").NumberFormat = "$#,##0"

Range("C4:C15").NumberFormat = "0.00%"

Columns("B").AutoFit

Columns("A").AutoFit

Columns("C").AutoFit

'If then Else Statements. Conditional formatting. And then Looping it together

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

'Added timer to Macro

endTime = Timer
MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

Range("A1").Value = "All Stocks (" + yearValue + ")"

End Sub

## **Refactored Code**

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
    
    tickerIndex = 0

    '1b) Create three output arrays
        
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    
    For i = 0 To 11
        
        tickerVolumes(i) = 0
       
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
            
        'End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next rowâ€™s ticker doesnâ€™t match, increase the tickerIndex.
        'If  Then
            
        If Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            tickerIndex = tickerIndex + 1
        End If

            
        'End If
    
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
