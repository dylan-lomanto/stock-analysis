# Stock Analysis

### Overview

The client, Steve, has requested a refactored version of the originally presented code, that is efficient enough to analyze the entire stock market in a reasonable amount of time.  The original code examines tables for 2017 and 2018 for 12 green stocks that displays daily trade counts and a variety of daily pricing categories. For our purposes the relevant fields are ticker, closing price, and volume. The original code outputs a table for either year with the ticker, total daily volume, and annual return percentage.  The total daily volume is acquired by summing the daily volumes for each ticker. The return percentage is calculated by dividing the final daily closing value by the first daily closing value for each ticker.  The original code uses nested for loops to return the data. The refactored code eliminates the need for nested loops buy pulling all of the relevant information on the first pass through the data set.

### Results

##### 2017 to 2018 Comparison
 
Green stocks returned a significantly higher percentage in 2017 vs 2018.  In 2017 11 of the 12 stocks returned a positive value, while only two did so in 2018.  ENPH and RUN were the only tickers with a positive return both years. ENPH had higher total daily volumes and average returns over the course of 2017 and 2018, making it the most attractive stock of those examined.

![2017_Table](https://user-images.githubusercontent.com/86164867/124499088-46df9480-dd72-11eb-81e4-fcbab7b5fc43.PNG)

![2018_Table](https://user-images.githubusercontent.com/86164867/124499243-8d34f380-dd72-11eb-9680-e40eed19104b.PNG)

