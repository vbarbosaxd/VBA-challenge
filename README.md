# VBA-challenge

## Introduction
VBA scripts were created to analyze multi year stock market data to make the large data set appear more easily understandable and legible.

![Stock_Market_Data_2016](https://user-images.githubusercontent.com/92385042/139357528-043a0061-a2ab-4c72-a553-7614a10c0bc5.png)

## Explanation of VBA scripts used
Two sets of VBA scripts were used to analyze the stock market data in Excel 2016: one for the year 2016 and another for the years 2014 and 2015.

This is because there was a different number of data pieces per ticker in 2014 and 2015 compared to the 2016 data, so the outputs using certain scripts were not accurate.

Conditional formatting was used to highlight positive change in green and negative change in red.

## 2014 and 2015 VBA Scripts:
Four individual scripts were created to analyze different portions:

*Ticker - used to find the different tickers in the data set

*Yearly_Change_2014_2015 - used to find the change between the end of the year and the beginning of the year for each ticker

*Percent_Change_2014_2015 - shows the previous data in percent form

*Total_Stock_Volume - sum of all stock volume for each ticker

And then they were all condensed into one script that can run through the entire worksheet:

*Stock_Data_2014_2015

## 2016 VBA Scripts:
A similar set was created for this year:

*Ticker - same one used for 2014 and 2015

*Yearly_Change2016

*Percent_Change_2016

*Total_Stock_Volume - same script used for 2014 and 2015 data

And then they were all condensed into one script that can run through the entire worksheet:

*Stock_Data_2016

