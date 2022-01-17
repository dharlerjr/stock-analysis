# VBA of Wall Street

## Overview of Project

### Background
In this module, we successfully helped Steve analyze the "DQ" stock data, then all stock data from the year 2018, and lastly, stock data from either year of our choosing. We also learned how to format the spreadsheet using only VBA scripts. 

### Purpose
After helping him analyze the stock data provided in the worksheet, Steve was elated with the results. However, our code may not be the most efficient method, especially if Steve decides to analyze thousands of stocks instead of just a few. Thus, we'll refactor our code to loop through all the data only once, which will reduce our script's runtime. 

## Results

### Stock Performance
According to our analysis, in 2017, each of the twelve stocks that we're working with had a positive return, except for TERP. Notably, DQ and SEDG had the highest returns of 199.4% and 184.5% respectively! However, 2018 was not as fruitful. In fact, only two out of the twelve stocks had a positive return this year: ENPH (81.9%) and RUN (84.0%). 

### Execution Times (seconds)
Original Script:
    2017: 2.25
    2018: 2.27

Refactored Script:
    2017: 0.27
    2018: 0.27

Thus, the refacotred code is more efficient.

## Summary

- What are the advantages or disadvantages of refactoring code?
The main advantage of refactoring code is to improve efficiency. In the module, we used a nested for loop to analyze all the stock data. While this solution is logically correct and easy to understand, it is not the most efficient method. In our refactored code, we may have multiple for loops, but none of them are nested, which is why our refactored code has a shorter runtime and thus, our refactored code is more efficient.

Another benefit of refactoring code is to simply clean it up, organize it, and make it look nice by potentially adding comments and white spaces in between blocks of code. 

In my opinion, the only disadvantage of refactoring code occurs when your original code is disorganized to begin with. As a result, when you go to refactor your solution, you may have trouble...
1. understanding your original solution
2. finding the specific line of code you'd like to improve or remove, AND/OR
3. seeing how your original code can be improved.
Thus, as a programmer, it's important to be organized.

- How do these pros and cons apply to refactoring the original VBA script?
Luckily, by following along with the module, my original code was neat and readable, with comments reminding me how the solution works. Combined with the step-by-step solution provided to us in the challenge, I had no issues refactoring my code and in the end, the refactored solution did in fact improve the script's efficiency.