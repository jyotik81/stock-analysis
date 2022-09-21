# stock-analysis

Deliverable 2: A written analysis of your results (README.md)

1. Overview of Project: Explain the purpose of this analysis.

This project has two purpose. 

- find stock returns to specific year 
- To refactor the code already return to reduce the macro run time to increase efficiency of the code.

The purpose of this analysis is to find out the stock with best positive return out of all stocks.
For that we have to create code where it will give out total daily volumes of each stock along with the returns of all stocks for comparison. 
The stock with positive outcome were marked with green highlight. 

2. Results: Using images and examples of your code, compare the stock performance between 2017 and 2018, as well as the execution times of the original script and the refactored script.

Before refactoring code, I copied the code that was needed to create the input box, chart headers, ticker array, and to activate the appropriate worksheet. 
The steps were then listed out in order to set the structure for the refactoring. Below is the instruction and code as written in the file.


1a) Create a ticker Index
    
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
        'If Then
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
               tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            End If
            
            
        'End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
            
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
               tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            End If

            '3d Increase the tickerIndex.
            
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
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
    

Once this refactoring code is entered where I used arrays, loops to repeat data and display which increasing efficiency and taking less runtime for macro.
1) Summary: In a summary statement, address the following questions.
a) What are the advantages or disadvantages of refactoring code?
i) Refactoring code makes our code more easy, clean, organized. 
ii) Due to cleaner code helps to design and software improvement, debugging, programming speed improvement. Helps other programmers to read it easily. 
iii) Some times refactoring may be challenging, if the code or application is large and no proper test case for existing code.
b) How do these pros and cons apply to refactoring the original VBA script?
 
The biggest benefit that occurred as a result of the refactoring was an decrease in macro run time.
 The original analysis took approximately one second to run, whereas our new analysis only took about a four of the time (approximately 0.25 seconds) to run.
 Attached in resources folder are the screenshots that indicate the run time for our new analysis.






