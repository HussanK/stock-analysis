# Stock Analysis

## Overview of Project

### Purpose
	
The purpose of this challenge was to get acquainted with  excel and VBA and to gain more confidence when using the coding language and platform. Learning VBA will help not only in 
the understanding of Excel but also teaches us important fundamentals in coding that can be applied to other languages such as Python. Again these skills are invaluable in a data 
analysis environment and will assist in future data-analytic work. The data given to us is a data set containing infomation about certain stocks and their prices. This is 
something that could easily translate into real world work. 

## Results

The results of the project were that in 2017 the stock prices given to us mostly went up we calculate this data using the formuala for return on investments as seen here in the code: 
```
Cells(i + 4, 3) = tickerEndingPrice(i) / tickerStartingPrice(i) - 1
```
On the otherhand the stocks for 2018 went down quite condsiderably in  by comparison. During this challenge we also got to compare the time it takes for our code to run. 

![2018 time results](https://github.com/HussanK/stock-analysis/blob/main/VBA_Challenge_2018.png)

Here is how long is took to complete the 2018 results. They actually were a bit slower than the 2017 result, which means there was more data in 2018 for the code to parse through.

## Summary

This project involved us refactoring code given to us to fit the needed purpose. The advantage of this is that, with basic understanding of the code I can manipulate it into a 
form that best suits the project. Also for ease of access, I can ask someone to refactor my code instead of having to teach them about all of the  intricacies of my codes. This 
leads into the disadvantages, primarily that by refactoring code instead of writing new code, you miss out on gaining a deeper understanding of the code itself. It a professional 
enviroment refactoring too much code could lead to the problem of errors occuring and the person who refactored the code not knowing how to fix it.

Specically in this code. We were code to refactor. We took code that was using a for loop to iterate through the data and another for loop to iterate through the cells. I changed 
the first for loop to be and internal counter variable called "tickerIndex." This made the code run faster but ultimately the refactored code is less clear. The unfactor code is 
very easy to understand compared to the new code. 

