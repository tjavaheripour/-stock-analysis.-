# stock-analysis
VBA project
## Overview of Project

In this project, we help Steve to find the best stock to invest in for his parents. His parents are passionate about green energy so they decided to invest into Daqo New Energy Corporation. Steve created an Excel file containing the stock data. We are going to use VBA to help him automate analyses by calculating “Total Daily Volume” and “Yearly Return”of Daqo stock to know how actively DQ was traded and performed over the last few years. 
Our review demonstrates that DQ is not a preferable purchase for this family as we see a significant drop in the second year so we are going to use flexible macro to produce the “Total Daily Volume” and “Return” for every single stock and based on that find the best stock for Steve’s parents to buy.  Refactoring our VBA code, we produce a more effective and efficient code that takes much less time to execute thousands of stocks.


## Results and Analysis

Reviewing all stocks in 2017 and 2018, we discovered that all stocks but “TERP” had positive return percentage in 2017 and “DQ” had the highest return rate of 199.4%. 

![All Stocks Analysis 2017.png](https://github.com/tjavaheripour/stock-analysis/blob/main/Resources/All%20Stocks%20Analysis%202017.PNG)


Furthermore, In 2018 , only two stocks “ENPH” and “RUN” had positive returns of 81.9% and 84.0% respectively and other stocks had negative rates.

![All Stocks Analysis 2018.png](https://github.com/tjavaheripour/stock-analysis/blob/main/Resources/All%20Stocks%20Analysis%202018.PNG)

So in conclusion, while DQ stock outraced all other stocks with the outstanding increase of 199.4% in 2017 but because it dropped significantly (-62.6%) in 2018 it is not the best stock to purchase.  Meanwhile RUN and ENPH rose about 130% and 6% respectively in 2017 and could maintain this positive rate of grow in 2018 so are performing much better than DQ over two years.


#### Compare VBA Stock execution times before and after refactoring code

##### Original VBA Code Performance in 2017 & 2018
![year value analysis 2017.png](https://github.com/tjavaheripour/stock-analysis/blob/main/Resources/year%20value%20analysis%202017.png)
##### Refactored VBA Code Performance in 2017 & 2018
![VBA_Challenge_2017.png](https://github.com/tjavaheripour/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png)
## Summary



#### 1. What are the advantages or disadvantages of refactoring code?


#### 2. How do these pros and cons apply to refactoring the original VBA script?
