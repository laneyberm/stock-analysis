# Does Refactoring VBA Code to make a VBA script run faster?

## Overview of Project

### Our client, Steve, wants to find the total daily volume and yearly return for each stock. Daily volume is the total number of shares traded throughout the day; it measures how actively a stock is traded. The yearly return is the percentage difference in price from the beginning of the year to the end of the year. Steve's parents are interested in Daqo's stock, so we'll start with DQ. Steve's parents want to know how actively DQ was traded in 2018. They believe that if a stock is traded often, then the price will accurately reflect the value of the stock. If we sum up all of the daily volume for DQ, we'll have the yearly volume and a rough idea of how often it gets traded. 

###Steve wants to know how DQ performed in 2018. One way to measure this is to calculate the yearly return for DQ. The yearly return is the percentage increase or decrease in price from the beginning of the year to the end of the year. In other words, if you invested in DQ at the beginning of the year and never sold, the yearly return is how much your investment grew or shrunk by the end of the year. Daqo dropped over 63% in 2018—yikes! Steve will definitely want to offer some better stocks to his parents. Since Daqo might not be the best option for Steve's parents to invest in, let's analyze multiple stocks to find some better choices for them. A lot of the work we've already done to analyze DQ can be repurposed to analyze any stock. With a little more code, we can analyze a whole list of stocks.

### Refactoring is a key part of coding process. When refactoring code, you aren’t adding new functionality; you just want to make the code more efficient—by taking fewer steps, using less memory, or improving the logic of the code to make it easier for future users to read. Refactoring is common on the job because first attempts at code won’t always be the best way to accomplish a task. Sometimes, refactoring someone else’s code will be your entry point to working with the existing code at a job. 

### In this project, I will refactor or edit a Stock Analysis solution code to loop through all the data one time in order to collect the same information from previous codes. Then, I’ll determine whether refactoring your code successfully made the VBA script run faster. The following are my results.

## Results

![Theater Outcomes vs Launch Dates](/VBA_Challenge_2017.png)

### Results: Using images and examples of your code, compare the stock performance between 2017 and 2018, as well as the execution times of the original script and the refactored script. 

### Also from the above graph, the number of failed events gradually increased from January until May. Between May and August, the number of events remained similiar. There is a spike in the number of failed events in October. The amount of canceled events remained very low throughout the year.

![Outcomes vs Goals](/VBA_Challenge_2018.png)

### The above graph provides percent information about play outcomes based upon is pledged goal raised amount in dollars. The percent of successful theater plays changed over time. There is a negative trend from less that $1000 between $24,999. There was a positive trend between $25000 and $44,999. After $45,000, the percent of successful events drop rapidly. The percent of failed events had the inverse information from the successful events.     

### How the graph is currently depicted, it is hard to compare positive and negative events without switching the data and graph to percentages when looking month to month. This information gives trend analysis of events that took place in a month. 

## Summary: In a summary statement, address the following questions.

- What are the advantages or disadvantages of refactoring code?

- How do these pros and cons apply to refactoring the original VBA script?

