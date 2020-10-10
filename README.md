# stock-analysis
**Overview of the Project**


This project is about the performance of 12 industries in 2017 and 2018 depending on their stock prices.
This project analyzed and refactored the VBA code to improve the performance and we found out  the best return of investment at the end of each year. 
Using array made some changes to the original code and measured the run time. With array we could improve the performance  and run time was much lesser than that of the original code.


Results:

Comparison between stock performances in the year 2017 and 2018


We did an analysis on 12 technology  industries stock performances.
In 2017 almost all of the stocks had positive returns except “TERP”.
“DQ had the highest return 199.4% and “SEDG” also performed Weill with a positive return of 184.5%.
“RUN got the least positive return of 5.5%. But when we compare the total volume “SPWR” and”FSLR”, got a total daily volume of 782,187,000 and 684,181,400. 
They were the most traded stocks in the year 2017.

But in 2018 only two stocks got positive returns and all the other stocks didn’t perform well.
2018 was not a good year for all of the technology stocks. “ENPH” and “RUN” got a high daily volumes.

Original code Run times

2017 Run time

<img width="758" alt="2017Runtime" src="https://user-images.githubusercontent.com/71113701/95643880-c7899280-0a77-11eb-8a40-eb1aaae69c8e.png">




2018 Run time

￼

  




     Original code took .4335938 seconds for the year 2017 and .4296875 seconds for 2018.
      But when I refactored the code using array it reduced the run time a lot.




'1a) Create a ticker Index
    
          tickerIndex = 0
  
  
    '1b) Create three output arrays
    
    
    Dim tickervolumes(12) As Long
    Dim tickerstartingprices(12) As Single
    Dim tickerendingprices(12)  As Single
    
    
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    ' If the next row’s ticker doesn’t match, increase the tickerIndex.
    
    
    For tickerIndex = 0 To 11
    
               tickervolumes(tickerIndex) = 0
    
        
    ''2b) Loop over all the rows in the spreadsheet.
    
         For i = 2 To RowCount
    
               If Cells(i, 1).Value = tickers(tickerIndex) Then
  
      '3a) Increase volume for current ticker
        
                tickervolumes(tickerIndex) = tickervolumes(tickerIndex) + Cells(i, 8).Value
        
                End If






Refactored code Run times

2017 improved run times


￼
<img width="613" alt="VBA_Challenge_2017 png " src="https://user-images.githubusercontent.com/71113701/95642714-dd935500-0a6f-11eb-8a08-1b959cc14846.png">


2018 improved run times


<img width="659" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/71113701/95642766-2a772b80-0a70-11eb-8ef8-ccec9305b057.png">


￼



Summary

   Advantages of  refactoring code is that it will take less time compared to the original code that is good for the developer and the system.
 It makes software easy to understand and finding bugs also easy and improves the design of software.

Disadvantages of refactoring code is that the refactoring often isn't done by the same person as the original designer. Therefore, he  doesn't have the same background in the system and the decisions that went behind the original design. You have  the risk that bugs avoided in the original design may come into the new design. Another thing is that it is time consuming, as it is a rework.
