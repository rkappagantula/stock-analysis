# Stock Analysis

## Overview
This report will analyze the data for 2018 and 2017 of 12 renewable energy stocks. The goal is to automate the formatting and analysis using VBA in order to output the yearly return and total volume for each stock. 

## Analysis 
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;*Refer to VBA_Challenge.xlsm in the repo specifically to the All Stocks Analysis worksheet. If you want to take a look to the entire code, open VBA and select the module called VBA_Challenge.

I first created an array for all the stocks, so that each one of them could be addressed with an index. 


To perform automated analysis of the stocks a nested for loop was used for reading data and storing the necessary information into arrays (total volume, first closing price, and last closing price) which was initiated at the beginning of the macro. 

since the output data was already saved into arrays, a for loop was created, and formatted the data to show highlight in green color if the return was greater than 0 and red if smaller than 0 to allow for a better interpretation of the findings.


## Results

Based on the analysis between 2017 and 2018. The performance of stocks is better in 2017 as compared to 2018. Ticker ENPH and RUN are the only two stocks which have performed well in both years. For better returns buy and hold ENPH and RUN stock in 2017 and hold it through 2018 as it had 81.9%  and 84% return, respectively. 
 

## Summary

The advantage with the orginal code was it would poit where the code was incorrect and it would only impact a specific macro and  the respective function. while other macros are not impacted. Meanwhile, the disadvantages  of the orginal code was it was all-over-the-place script. On the other hand. The advantage of the refactored code 
was that it was more efficient and organized by using multiple for loops vs using nested loop has optimized the performance of the VBscript.