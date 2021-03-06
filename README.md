# Refactor VBA Code and Measure Performance

## Overview of Project
### Purpose of Project
In this challenge, I refactored code to loop through all the data one time in order to collect the same information from the original code. Then I determined how I was able to refactor the code successfully to make the VBA script run faster. 

**Refactor**: editing exisiting code to make it more efficient—by taking fewer steps, using less memory, or improving the logic of the code to make it easier for future users to read

### Background of Project
The original code **AllStocksAnalysis()** compares stock market stocks. It works well for a dozen stocks so I refactored code **AllStocksAnalysisRefactored()** to expand the dataset to include the entire stock market over the last few years and improved the code so that it runs faster than the original. 

---
## Results
There are three key differences between the original and refactored code. 
1. The refactored code contains three output arrays: *Dim tickerVolumes(12) As Long, Dim tickerStartingPrices(12) As Single,* and *Dim tickerEndingPrices(12) As Single*. 
2. The refactored code contains a formatting loop while the original code runs that as a separate subroutine. 
3. The original code contains a nested loop while the refactored code contains three loops. 
        
        * Original Code: A loop to increase volume over all the rows in the spreadsheet nested inside a loop that initializes the tickerVolumes to zero.
        
        * Refactored Code: A loop to initialize the tickerVolumes to zero. A loop to increase volume over all the rows in the spreadsheet. A loop to format the spreadsheet.     

As a result, the codes produce the same stock performances, but the refactored code executes faster. 

### Stock Performance  
Of the twelve stocks this code compared, the best stocks were ENPH and RUN because they are the only stocks with an increase in return in both 2017 and 2018. Of the two, ENPH had significant double digit returns both years with 129.5% in 2017 and 81.0% in 2018 while RUN had returns of 5.5% and 84% respectively. The worst stock was TERP because it was the only stock with negative returns in both years.  

### Execution Time 
Each subroutine contains code that measures and reports on the execution time - *MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)*. 

On average, the refactored code runs half a second faster. For the year 2017 and 2018, AllStocksAnalysisRefractored() runs approximately .57 and .53 seconds faster respectively than AllStocksAnalysis(). The images below show the message box results. 

**2017 ORIGINAL:**![VBA_Challenge_2017_b4refractor](VBA_Challenge_2017_b4refractor.png) 

**2017 REFRACTORED:**![VBA_Challenge_2017](VBA_Challenge_2017.png) 

**2018 ORIGINAL:**![VBA_Challenge_2018_b4refractor](VBA_Challenge_2018_b4refractor.png) 

**2018 REFRACTORED:**![VBA_Challenge_2018](VBA_Challenge_2018.png)        


## Summary 
The codes are very similar. In terms of output, they produce the same results with the only exception being that the refactored code includes the formating that the original code runs as a separate subscript. Since the results are the same, the primary advantage to the user of the refactored code is the reduced execution time. For the coder, the advantage of the refactored code is enhanced readability. Since the code exisits in single loops, the code is more simple, meaning that understanding each line is more straightforward. 


While there are some disadvantages to refactoring code, the advantages far outweigh the disadvantages. First, coders should make sure they understand how the original code works and that it is bug free. From there, the main challenge is correctly naming arrays. In this challenge, I encountered a bug because I used the array *ticketVolume* instead of *tickerVolume*. If coders successfully navigate these challenges, then their focus can shift away from creating something new from scratch to creating a more efficient piece of code. The intended consequences being to build on exisiting resources to save time and redirect resources into other tasks.  
