# **Stock-Analysis**

## **Overview of the Project**
 Our client would like to run an analysis on stock tickers to determine which stock has the best return for his parents. To complete this project I used VBA Macros to run the analysis. The first set of script worked well, however, we refactored the code to make it run more efficient and faster. Client would also like buttons in order to run the code quickly. 
 
## **Results**
The first version of code gave me the answers I was looking forever, however, it was a little slow. When I refactored the code, it made a staggering improvement on time to answer. We needed to refactor in this case as we want to expand our data search. If we did not refactor and added more tickers this would be a very slow macro which creates inefficiencies. Nested loops seem to slow the code down as there is more for the computer to read/do. Using arrays really helped speed things up when we added that to the refactored coding. Module 1 is original code and Module 2 is refactored code within the Worksheet> VB. As you can see below, the first script took over 5 seconds while the refactored code took a fraction of a second!

**Original Spreadsheet with Time Clock**

![2018 Analysis- Original Script](https://github.com/Adam-Warrick/Stock-Analysis/blob/main/2018%20Analysis%20-%20Original%20Script.png)

**Original Code**

![2018 Orginal Dim Code](https://github.com/Adam-Warrick/Stock-Analysis/blob/main/2018%20Original%20Dim%20Code.png)

**Refactored Spreadsheet with Time Clock**

![2018 Analysis - Refactored Script](https://github.com/Adam-Warrick/Stock-Analysis/blob/main/2018%20Analysis%20-%20Refactored%20Script.png)

**Refactored Code**

![2018 Recap - Arrays](https://github.com/Adam-Warrick/Stock-Analysis/blob/main/2018%20Recap%20-%20Arrays.png)

## **Summary**
Nested Loops and not using arrays are the culprits to slower code. In this case, it is proven that Refactoring is the way to go in order to be more efficient.

###### Advantages
Our refactored code ran quicker and more efficient then our previous original code. This is due to not having Nested Loops within the code which made the macro run slower initially. 

###### Disadvantages
Some disadvantages for the original code was that it was slow, clunky, and inefficient. The refactored code seemed slightly harder to build as a deeper understanding of the variables was needed. The hard part around writing efficient code is that your first pass may not always be the best result...with this, you need to dig deeper and think outside the box (Refactoring) to gain a better result. 
