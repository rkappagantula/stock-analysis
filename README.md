# Stock Analysis

## Overview
This report will analyze the data for 2018 and 2017 of 12 renewable energy stocks. The goal is to automate the formatting and analysis using VBA in order to output the yearly return and total volume for each stock. 

## Analysis 
Refer to VBA_Challenge.xlsm specifically to the All Stocks Analysis worksheet.Open VBA and select the module called VBA_Challenge to review the code.

Created an array for all the stocks, so that each one of them could be addressed with an index.
![image](https://user-images.githubusercontent.com/96051648/152624507-adc97b46-b2c1-47c3-8564-963ecf18dc7f.png)

To perform automated analysis of the stocks a nested for loop was used for reading data and storing the necessary information into arrays (total volume, first closing price, and last closing price) which was initiated at the beginning of the macro. Since the output data was already saved into arrays, a for loop was created, and formatted the data to show highlight in green color if the return was greater than 0 and red if smaller than 0 to allow for a better interpretation of the findings.


## Results

Based on the analysis between 2017 and 2018. The performance of stocks is better in 2017 as compared to 2018. Ticker ENPH and RUN are the only two stocks which have performed well in both years. For better returns buy and hold ENPH and RUN stock in 2017 and hold it through 2018 as it had 81.9%  and 84% return, respectively. 
2017 Results:
![image](https://user-images.githubusercontent.com/96051648/152624604-9b408eb5-2c12-485f-bdec-92209b0ed359.png)
 
2018 Results: 
![image](https://user-images.githubusercontent.com/96051648/152624616-79ea9554-3fd9-4ae3-a11d-712da041a77c.png)

## Summary

The advantage with the orginal code was it would poit where the code was incorrect and it would only impact a specific macro and  the respective function. while other macros are not impacted. The disadvantages  of the orginal code was it was all-over-the-place script. On the other hand. 

The advantage of the refactored code was that it was more efficient and organized by using multiple for loops vs using nested loop has optimized the performance of the VBscript.
