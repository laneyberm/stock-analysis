# Does Refactoring VBA Code to make a VBA script for Stock Analysis in 2017 and 2018 run faster?

## Overview of Project

### Our client, Steve, wants to find the total daily volume and yearly return for each stock for an investment for his parents. Daily volume is the total number of shares traded throughout the day; it measures how actively a stock is traded. The yearly return is the percentage difference in price from the beginning of the year to the end of the year. Steve's parents are interested in Daqo's stock and want to know how actively DQ was traded in 2018. They believe that if a stock is traded often, then the price will accurately reflect the value of the stock. Steve wants to know how DQ performed in 2018 and depending on it's performance, he wants to offer some better stocks to his parents. We have analyzed multiple stocks to find some better choices for them. 

### Refactoring is a key part of coding process. When refactoring code, you aren’t adding new functionality; you just want to make the code more efficient—by taking fewer steps, using less memory, or improving the logic of the code to make it easier for future users to read. Refactoring is common on the job because first attempts at code won’t always be the best way to accomplish a task. Sometimes, refactoring someone else’s code will be your entry point to working with the existing code at a job. 

### In this project, we will refactor or edit a Stock Analysis solution code to loop through all the data one time in order to collect the same information from previous codes. Then, we’ll determine whether refactoring code successfully made the VBA script run faster. 

## Results

### Results for Steve's Investments

![Stock Analysis for 2017 and Refactoring time](/VBA_Challenge_2017.png)
#### After running the analysis for 2017, we find that Ticker:DQ has the highest return on investment at 199.4%. Additionally, all but Ticker: VSLR netted a positive return. 

![Stock Analysis for 2018 and Refactoring time](/VBA_Challenge_2018.png)
#### After running the analysis for 2018, we find that Ticker:DQ has a decrease of 63%. Additionally, we find that Ticker:ENPH increased 82% and Ticker:RUN increased return to 84%. After reviewing this data, it would be advisable to not invest in Ticker:DQ, but look to Ticker:ENPH and Ticker:RUN. 

#### We were able to compare stock performance between 2017 and 2018 by summing up all of the daily volume for the individual stocks. By doing this, we'll have the yearly volume and a rough idea of how often it gets traded. If a stock is traded often, then the price will accurately reflect the value of the stock. One way to measure this is to calculate the yearly return for DQ. The yearly return is the percentage increase or decrease in price from the beginning of the year to the end of the year. We used the following code in VBA to find the stock return for a given year.

![VBA Code for finding return](.png)

### Results in Refactored Script Times versus the Original Script Times

#### 

## Summary: In a summary statement, address the following questions.

- What are the advantages or disadvantages of refactoring code?

- How do these pros and cons apply to refactoring the original VBA script?

