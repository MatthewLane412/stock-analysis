# Beating the Stock Market

## Overview of Project
The purpose of this analysis is to help Steve with an analysis of enery stocks, specifically "DAQO New Energy Corp (DQ). Steve's parent invested all there money into DQ and Steve is worried about them not diversifying, and how the stock has performed historically. To start this project Steve provided a spreadsheet of clean energy stocks to dig into. 

Key metrics Steve is looking for inclucde: 
- Total Daily Volumne
- Return


## Results

I have used VBA within Excel to complete this project. With VBA, we have created an easily repeatable deliverable. 

Some of the techniques used were:
- Setting variables (e.g. ```Dim startTime As Single```)
- Adding values to cells (e.g. ```Cells(3, 1).Value = "Ticker"```)
- Loading Arrays (e.g. ```tickers(0) = "AY"```)
- Looping through repetitive task with For Loops (e.g. ``` For tickerIndex = 0 To 11 ... Next tickerIndex```)



### Stock Analysis

![2017 Stock Analysis](/Resources/All_Stocks_(2017).png)
![2018 Stock Analysis](/Resources/All_Stocks_(2018).png)

Looking at the performance of energy stocks between 2017 and 2018, we can tell 2017 was a much better year for energy stocks based on our sample size. DQ in particular had a tough year in 2018. On the other hand, ENPH ticker looks like a great buy. Two great return years in a row. ENPH also has a higher total daily volume than any other stock in the sample size for 2018. 

### Before Refactored code

![2017 Before Refactored Run Time](/Resources/VBA_Challenge_2017_Before_Refactoring.PNG)
![2018 Before Refactored Run Time](/Resources/VBA_Challenge_2018_Before_Refactoring.PNG)

### Refactored Run Times

![2017 Refactored Run Time](/Resources/VBA_Challenge_2017.PNG)
![2018 Refactored Run Time](/Resources/VBA_Challenge_2018.PNG)


## Summary

## What are the advantages or disadvantages of refactoring code?
Advantages of refactoring code could include Enhancement of code, bug fixing, 


## How do these pros and cons apply to refactoring the original VBA script?


