# stock-analysis
VBA project
## Overview of Project

In this project, we help Steve to find the best stock to invest in for his parents. His parents are passionate about green energy so they decided to invest into Daqo New Energy Corporation. Steve created an Excel file containing the stock data. We are going to use VBA to help him automate analyses by calculating “Total Daily Volume” and “Yearly Return”of Daqo stock to know how actively DQ was traded and performed over the last few years. 
Our review demonstrates that DQ is not a preferable purchase for this family as we see a significant drop in the second year so we are going to use flexible macro to produce the “Total Daily Volume” and “Return” for every single stock and based on that find the best stock for Steve’s parents to buy.  Refactoring our VBA code, we produce a more effective and efficient code that takes much less time to execute thousands of stocks.


## Results and Analysis

First, in the original VBA script, we need to calculate the yearly volume of each stock by adding up the daily volumes. For reaching this goal, we created an array of all 12 tickers with a loop through all rows of data in Year worksheet to extract “Total Volume” for each stock by using if-statements to distinguish “startingprice” and “endingprice” for each stock. Secondly, to know how each stock performed in 2017 and 2018, we should calculate the yearly return to find the percentage increase/decrease in price from the beginning of the year to the end of the year.

Return= endingPrice / startingPrice – 1


Reviewing all stocks in 2017 and 2018, we discovered that all stocks but “TERP” had positive return percentage in 2017 and “DQ” had the highest return rate of 199.4%. 

![All Stocks Analysis 2017.png](https://github.com/tjavaheripour/stock-analysis/blob/main/Resources/All%20Stocks%20Analysis%202017.PNG)


Furthermore, In 2018 , only two stocks “ENPH” and “RUN” had positive returns of 81.9% and 84.0% respectively and other stocks had negative rates.

![All Stocks Analysis 2018.png](https://github.com/tjavaheripour/stock-analysis/blob/main/Resources/All%20Stocks%20Analysis%202018.PNG)

So in conclusion, while DQ stock outraced all other stocks with the outstanding increase of 199.4% in 2017 but because it dropped significantly (-62.6%) in 2018 it is not the best stock to purchase.  Meanwhile RUN and ENPH rose about 130% and 6% respectively in 2017 and could maintain this positive rate of grow in 2018 so are performing much better than DQ over two years.


#### Compare VBA Stock execution times before and after refactoring code

By refactoring the code and taking fewer steps and using less memory, you just make the code more efficient running faster. You can see the difference between performance execution times before and after refactoring data as shown below. 

##### Original VBA Code Performance in 2017 & 2018

![year value analysis 2017.png](https://github.com/tjavaheripour/stock-analysis/blob/main/Resources/year%20value%20analysis%202017.png)![year value analysis 2018.png](https://github.com/tjavaheripour/stock-analysis/blob/main/Resources/year%20value%20analysis%202018.png)

As you can see from the below image, the original code contains a nested for loop that checks all rows of data for each stock ticker and gathers data before cycling to the next stock ticker. Because of that the code takes a long time to execute:

![Original VBA script.png](https://github.com/tjavaheripour/stock-analysis/blob/main/Resources/Original%20VBA%20script.PNG)

##### Refactored VBA Code Performance in 2017 & 2018

![VBA_Challenge_2017.png](https://github.com/tjavaheripour/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png)![VBA_Challenge_2018.png](https://github.com/tjavaheripour/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png)

As you can see, the refactored VBA script gathers data with only looping once through all the rows in the spreadsheet.:

![Refactored loop.png](https://github.com/tjavaheripour/stock-analysis/blob/main/Resources/Refactored%20loop.PNG)

## Summary



#### 1. What are the advantages or disadvantages of refactoring code?

Code Refactoring is a way of restructuring and optimizing existing code without changing its behavior. It is a way to improve the code quality. Sometimes bad patterns like tight coupling, duplicate code, long methods, large classes, etc. are detected in the code so the code should be refactored in this case. Refactoring code has lots of positive benefits, it helps in increasing the flexibility of the code, making it easier to understand and organize. Negative side of refactoring code is increasing the chance of mistake when it is long. you will have to spend much more time in solving the problem and the chances of making mistake increases due to complexity of the code.


#### 2. How do these pros and cons apply to refactoring the original VBA script?
In this case, the original VBA script may work as well to automate analysis of few stocks but executing the code to evaluate thousands of stocks takes a long time. The refactored VBA code script will loop through all the data by taking fewer steps in less time and use less memory for processing data. Moreover, clean and organized code is easier to read and maintain. 
