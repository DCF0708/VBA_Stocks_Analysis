# VBA Challenge: Stocks Analysis (MarkdownMonster Format)

## Project Overview

  Steve is a recent financial graduate who has identified a need to easily and quickly analyze stock data in order to properly advise his parents (and hopefully future clients) on their investments. We have already created the Module 2 workbook for Steve that analyzes *DAQO* (New Energy Corp), as well as comparing a handful of other stocks in the green energy field to help with diversification advice. The goal of the challenge is to refactor this existing code to loop through the data **one** time to collect the same information. After refactoring, a conclusion will be drawn on the speed of the refactored code vs the original Module 2 workbook. 
 
## Refactor Results

  Module 2 workbook met all the requirements for what the user needed. It analyzed the data and returned values that held useful information about the market. However, if Steve plans to grow his client base he will inevitably come across individuals with portfolios much more diverse than his parents'. Advising clients with diverse portfolios involves looking at a much more vast swath of the market and with this comes thousands more data points to analyze. While Module 2 fits the need for Steve's parents, it needs to be optimized in order to be able to handle more data. 
  
### Deep Dive
  
  A quick look at the *Module 2* code reveals multiple areas that are sub-optimal. Primarily, there is an embedded for loop that combs painstakingly through the data 12 times. This can be solved by creating arrays for storage of **total daily volume** by ticker, **ticker starting prices**, and **ticker ending prices** to calculate return. Modifying the 'for loop' to loop through the code only once and store all the values in arrays rather than iterating through the entire data set 12 times allows huge data sets to be analyzed in the most efficient way possible. Plugging a large data set into the *Module 2* code would result in monstrous wait time compared to the refactored *Challenge code*. This is exemplified in the screenshots of the run times collected at the end of the analyses of the *Module 2* code vs the *Challenge code*. **Note: the data set used constitutes a relatively small example of what could possibly be required by a client.**
  
#### 
  ![](../../Project%20Resources/VBA_Original_2017.png)
  ![](../../Project%20Resources/VBA_Original_2018.png)
  
#### **Above: Results and run times for *Module 2* (2017 top / 2018 bottom)**
#### 
  ![](../../Project%20Resources/VBA_Challenge_2017.png)
  ![](../../Project%20Resources/VBA_Challenge_2018.png)
#### 
#### **Above: Results and run times for *Challenge* (2017 top / 2018 bottom)**
#### 
Screenshots reveal that the refactored *Challenge code* runs approximately ***8.7 times faster*** than the *Module 2* code on the same data set. As the data set grows, so will the difference between the times. 

## Closing Statements
Refactoring code is the best way to assure the codes' efficiency and longevity. The programmer becomes more acquainted with the code which in turn allows for more time to discover optimizations resulting in a clean-running easier to read code. Refactoring means Steve can now approach new clients backed with effective data and confidence in the programs' ability to handle vast data sets in a timely manor. *Module 2* code was a means to accomplish the short term goals set by the user, this acted as the road map to accomplish the long term goal of handling bigger client's data. The only downside to refactoring applies to this project as well as coding in general. Refactoring is a time consuming process that can many times lack a definite indicator of just how long it may take. This could present an issue for extremely time-sensitive projects, however proper exercise of time management and scope analysis can do a great deal to shed light on what needs to be done. 
