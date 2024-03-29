# VBA of Wall Street

## Overview of Project

### Background
In order to diversify his parent's investment portfolio, we helped Steve analyze stock data from 2017 to 2018 for twelve green-energy companies. Specifically, we created multiple Microsoft Excel VBA scripts, including one to analyze a single stock, another to analyze all twelve stocks from one year, then finally, a script to analyze all twelve stocks from one year of the user's choosing. The results of our stock analysis our displayed below. 

![Stock Analysis 2017](https://github.com/dharlerjr/stock-analysis/blob/main/Stock_data_2017.PNG)
![Stock Analysis 2018](https://github.com/dharlerjr/stock-analysis/blob/main/Stock_data_2018.PNG)

### Purpose
After helping him analyze the stock data provided in the worksheet, Steve was elated with the results. However, our script was not very efficient and would likely perform poorly if given many more stocks to analyze. Thus, to reduce our script's runtime, we refactored our code to loop through all of the stock data only once, rather than looping the all of the data once for each different stock.

## Results

### Stock Performance
According to our analysis, in 2017, each of the twelve stocks had a positive return, except for TERP. Notably, DQ and SEDG had the highest returns of 199.4% and 184.5% respectively! However, 2018 was not as fruitful. In fact, only two out of the twelve stocks had a positive return this year: ENPH (81.9%) and RUN (84.0%). 

### Execution Times (seconds)
Original Script (see images below)

![Original Script Runtime 2017](https://github.com/dharlerjr/stock-analysis/blob/main/green_stocks_2017.PNG)
![Original Script Runtime 2018](https://github.com/dharlerjr/stock-analysis/blob/main/green_stocks_2018.PNG)

Refactored Script (see images below)

![Refactored Script Runtime 2017](https://github.com/dharlerjr/stock-analysis/blob/main/VBA_Challenge_2017.PNG)
![Refactored Script Runtime 2018](https://github.com/dharlerjr/stock-analysis/blob/main/VBA_Challenge_2018.PNG)

As you can see, our refactored script runs in one-tenth of the time as the original script.

## Summary

What are the advantages or disadvantages of refactoring code?
- The main advantage of refactoring our code is to improve efficiency. Originally, we used a nested for loop to analyze all of the stock data. While this solution is logically correct and easy to understand, it is not the most efficient method. In our refactored script, we utilize multiple for loops, but none of them are nested, which is why our refactored code has a shorter runtime. Furthermore, our refactored solution utilizes a ticker index which keeps track of which ticker symbols we've already analyzed, thus allowing us to only the loop the entire dataset one time, rather than looping through the entire dataset one time for each stock. Another benefit of refactoring our code is to make it more readable by adding comments and whitespace. This way, if another developer takes a look at our code, he/she will have an easier time understanding our script. In my opinion, the only disadvantage of refactoring code occurs when your original code is disorganized to begin with. As a result, when you go to refactor your solution, you may have trouble...
  - understanding your original solution
  - finding the specific line of code you'd like to improve or remove, AND/OR
  - seeing how your original code can be improved.

How do these pros and cons apply to refactoring the original VBA script?
- Luckily, by following along with the module, my original code was neat and readable, with comments reminding me of the logic behind our original solution. Combined with the step-by-step solution provided to us in the challenge, I had no issues refactoring my code and in the end, the refactored solution significantly improvee the script's efficiency.
