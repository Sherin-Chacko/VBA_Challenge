### 1. OVERVIEW OF PROJECT : 
The purpose of this challenge is to refactor the VBA code from Module 2 solution and to measure its performance by analyzing the script run-times of the original    script and the refactored script for the different stocks of the year 2017 and 2018.       

The dataset includes two spreadsheets with information on 12 different stocks. The stock information contains a ticker value, the date the stock was issued, the opening, closing and adjusted closing price, the highest and lowest price, and the volume of the stock. The goal is to retrieve the ticker, the total daily volume, and the return on each stock.

### 2. A) COMPARISON OF STOCK PERFORMANCE FOR 2017 AND 2018:
The starter code provided in the challenge was executed to deliver the following results in Excel for the year 2017 and 2018.

<img width="486" alt="Screen Shot 2021-10-19 at 3 19 15 PM" src="https://user-images.githubusercontent.com/91294352/137976687-822c82dc-cf00-43c5-830b-5654949612f2.png">

* The tickers ENPH and RUN would have been good investments due to the positive returns in 2018, and both of the tickers had increases greater than         $200,000,000 over the 2017 total daily volumes.
* Comparing the 2017 and the 2018 stocks, the difference in the total daily volume between the two years that resulted in less than a $100,000,000 in increased       volume was not enough to generate a positive 2018 return percentage. 

### B) POP-UPS WITH SCRIPT RUN TIME:
The starter code contains a nested for-loop. When a loop is nested inside another loop, the inner loop runs many times inside the outer loop. The execution times of the original script is as follows:

<img width="265" alt="Screen Shot 2021-10-20 at 12 04 42 PM" src="https://user-images.githubusercontent.com/91294352/138129929-3d27ab5c-c16d-41db-a78c-2ef7f5bb8ba2.png">      <img width="267" alt="Screen Shot 2021-10-20 at 12 05 09 PM" src="https://user-images.githubusercontent.com/91294352/138129960-6749dd9b-c96e-4c92-aa53-ec9b0f0aa01d.png">

The refactored code ran faster than it did in the original script. This is because before iterating through the dataset, the variable enabled me to assign the ticker volumes, ticker starting price and the ticker ending price to each ticker symbol. The execution times of the refactored script is as follows:
