# Refactor VBA Code and Measure Performance

## Overview of Project
### Purpose of Project
In this challenge, you’ll edit, or refactor, the Module 2 solution code to loop through all the data one time in order to collect the same information that you did in this module. Then, you’ll determine whether refactoring your code successfully made the VBA script run faster. Finally, you’ll present a written analysis that explains your findings.

Refactoring is a key part of the coding process. When refactoring code, you aren’t adding new functionality; you just want to make the code more efficient—by taking fewer steps, using less memory, or improving the logic of the code to make it easier for future users to read. Refactoring is common on the job because first attempts at code won’t always be the best way to accomplish a task. Sometimes, refactoring someone else’s code will be your entry point to working with the existing code at a job.
### Background of Project
Steve loves the workbook you prepared for him. At the click of a button, he can analyze an entire dataset. Now, to do a little more research for his parents, he wants to expand the dataset to include the entire stock market over the last few years. Although your code works well for a dozen stocks, it might not work as well for thousands of stocks. And if it does, it may take a long time to execute.

---
## Results
I organized the data to see if the month the kickstarter launched could impact its success. First, I made the data more detailed by splitting the Category and Subcategory column into two distinct columns. This allowed me to view the wider category of theater and later the narrower subcategory plays. Next, I converted Unix timestamps to identify the launch date. For example, I turned the cell **1434931811** into **06/22/15** using the code =(((J2/60)/60)/24)+DATE(1970,1,1). Finally, I created a pivot table that filtered based on "Parent Category" and "Years." From that pivot table I created the line graph shown below. 

![VBA_Challenge_2017](VBA_Challenge_2017.png)

Based on the line graph above, I had the following takeways about theather campaigns:
* May, June, and July respectively had the highest number of successful campaigns. 
* October and May respectively had the highest number of failed campaigns.
* December had about as many successful campaigns as failed ones.
* In every month there were more successful than failed campagins. Generally, the successful and failed trend lines create the same shape and are similar to each other with the exception of May, October, and December. In May, the larger than average gap between successful and failed campaigns suggests a better chance for success. In October, the smaller than average gap between successful and failed campaigns suggests a better chance for failure. 


Next, I organized the data to see if the funding goal could impact the kickstarter's success. First, I needed to count the number of successful, failed, and canceled plays by goal. To do this, I used the COUNTIFS formula. For example, to count the number of successful plays with a goal between 1,000 and 4,999 I wrote the formula =COUNTIFS('Raw Data'!$O:$O, "plays",'Raw Data'!$D:$D, ">=1000", 'Raw Data'!$D:$D, "<5000", 'Raw Data'!$F:$F, "successful"). Then, I converted the number of successful, failed, and cancled plays to a percentage to more accurately compare them. Using this information, I created the line graph shown below. 

![VBA_Challenge_2018](VBA_Challenge_2018.png)

Based on the line graph above, I had the following takeways about play campaigns:
* Because there are zero cancled plays, successful and failed campaigns are symetric about the 50% line. 
* Generally, the percentage of successfully funded plays decreases as the goal increases with the exception of 29,999 to 49,999. This exception is likely due to a smaller sample size of 11 or less. 
* Plays with a goal of less than 5,000 have a success rate 20 points higher than plays with a goal between 5,000 and 20,000. 


### Challenges 
One challenge I encountered in writing my countifs formula was ensuring that I only counted the *plays*. Initially, my formula =COUNTIFS('Raw Data'!$D:$D, ">=1000", 'Raw Data'!$D:$D, "<5000", 'Raw Data'!$F:$F, "successful") counted every subcategory so I had to add an additional criteria and changed my formula to =COUNTIFS('Raw Data'!$O:$O, "plays",'Raw Data'!$D:$D, ">=1000", 'Raw Data'!$D:$D, "<5000", 'Raw Data'!$F:$F, "successful"). 

---
## Summary 
