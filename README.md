### 1. OVERVIEW OF PROJECT : 
The purpose of this challenge is to refactor the VBA code from Module 2 solution and to measure its performance by analyzing the script run-times of the original    script and the refactored script for the different stocks of the year 2017 and 2018.       

The dataset includes two spreadsheets with information on 12 different stocks. The stock information contains a ticker value, the date the stock was issued, the opening, closing and adjusted closing price, the highest and lowest price, and the volume of the stock. The goal is to retrieve the ticker, the total daily volume, and the return on each stock.

### 2. HOW ANALYSIS WAS PERFORMED:
The starter code provided in the challenge was executed to deliver the following results in Excel for the year 2017 and 2018. 
<img width="486" alt="Screen Shot 2021-10-19 at 3 19 15 PM" src="https://user-images.githubusercontent.com/91294352/137976687-822c82dc-cf00-43c5-830b-5654949612f2.png">

Comparing the 2017 and the 2018 stocks, the difference in the total daily volume between the two years that resulted in less than a $100,000,000 in increased volume was not enough to generate a positive 2018 return percentage. The tickers ENPH and RUN had would have been considered good investments due to the positive returns in 2018, and both of the tickers had increases greater than $200,000,000 over the 2017 total daily volumes.

### POP-UPS WITH SCRIPT RUN TIME:
The starter code contains a nested for-loop. When a loop is nested inside another loop, the inner loop runs many times inside the outer loop. 



refactored code should run faster than it did in this module.
