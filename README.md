# Refactor VBA Code and Measure Performance

## Overview of Project
### Purpose of Project
In this challenge, I refactored code to loop through all the data one time in order to collect the same information from the original code. Then I determined how I was able to refactor the code successfully to make the VBA script run faster. 

**Refactor**: editing exisiting code to make it more efficient—by taking fewer steps, using less memory, or improving the logic of the code to make it easier for future users to read

### Background of Project
The original code **AllStocksAnalysis()** compares stock market stocks. It works well for a dozen stocks so I refractored code **AllStocksAnalysisRefactored()** to expand the dataset to include the entire stock market over the last few years and improved the code so that it runs faster than the original code. 

---
## Results
There are three key differences between the original and refractored code. 
1. The refractored code contains three output arrays: *Dim tickerVolumes(12) As Long, Dim tickerStartingPrices(12) As Single,* and *Dim tickerEndingPrices(12) As Single*. 
2. The refracted code contains a formatting loop while the original code runs that as a seperate subroutine. 
3. The original code contains a nested loop while the refractored code contains three loops. 
        
        * Original Code: A loop to increase volume over all the rows in the spreadsheet nested inside a loop that initializes the tickerVolumes to zero.
        
        * Refractored Code: A loop to initialize the tickerVolumes to zero. A loop to increase volume over all the rows in the spreadsheet. A loop to format the spreadsheet.     

As a result, the codes produce the same stock performances, but the refractored code executes the code faster. 

### Stock Performance  

### Execution Time 

![VBA_Challenge_2017_b4refractor](VBA_Challenge_2017_b4refractor.png) 

![VBA_Challenge_2017](VBA_Challenge_2017.png) 

![VBA_Challenge_2018_b4refractor](VBA_Challenge_2018_b4refractor.png) 

![VBA_Challenge_2018](VBA_Challenge_2018.png)        


## Summary 
