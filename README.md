# Stock-Analysis 
# Overview
1.	Create a VBA macro that can trigger pop-ups and inputs, read and change cell values, and format cells.
2.	Use for loops and conditionals to direct logic flow.
3.	Use nested for loops.
4.	Apply coding skills such as syntax recollection, pattern recognition, problem decomposition, and debugging.
Purpose
We performed an initial stock analysis for Steve. The original code was then refactored to loop only once. The purpose was to determine if the refactored changes made an impact on the run time. Then, we’ll determine whether refactoring your code successfully made the VBA script run faster. 

# Results
# Comparing 2017 and 2018 Stock Analysis
After comparing the 2017 and the 2018 stocks, the difference in the total daily volume between the two years resulted in less than a $100,000,000 in increased volume. This was not enough to generate a positive 2018 return percentage. The tickers ENPH and RUN would have been considered good investments due to the positive returns in 2018. 
![image](https://user-images.githubusercontent.com/95143562/149677784-aa4e82d8-9b83-4261-9471-c578bd907e95.png)
# Comparing the Original Run Times to the Refactored Run Times
Run times for the original code took around .4 seconds
![image](https://user-images.githubusercontent.com/95143562/149677837-02a776e4-c25e-4bb0-b771-1e6e733bd792.png)
![image](https://user-images.githubusercontent.com/95143562/149677843-cb44dd9c-8c0b-444e-85fa-4cc14d5dd110.png)

Run times for the refactored code took around .08 seconds, which is displayed in scientific notation: 8 E -02.
![image](https://user-images.githubusercontent.com/95143562/149677856-c1ec73dc-a9f5-447c-acce-23855e94e6a8.png)
![image](https://user-images.githubusercontent.com/95143562/149677863-262a3037-05c6-4df9-bdae-5acfad8a4755.png)

# Summary
# What are the advantages or disadvantages of refactoring code?
# Advantages
Finding the root cause of a potential bug can be done with refactoring. A programmer can catch duplicated subroutines, unnecessary loops, redundant statements, or code that was used to run down an error but was accidently left in the script.
# Disadvantages
When programming it can have multiple approaches to a solve a single problem. In those differences, programmers may have alternative steps, which will require testing to see how those differences play out in the script. Refracting a stable code to apply a different set of logic could be costly or introduce new bugs into the system
# How do these pros and cons apply to refactoring the original VBA script?
Reducing the number of loops decreases the memory needed for processing the data, which reduces the run time and optimizes the performance of the script. 
