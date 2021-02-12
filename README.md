# Stock-Analysis Using VBA

## 1. Overview
	Steve wanted to expand the capabilities of the spread sheet to analyze the entire data set. To achieve this, VBA code was used and a script to loop through the entire data sets for 2017 and 2018. Only 12 ticker symbols were provided in the data sets and the closing prices were used to determine the returns.
## 2. Results
	Overall, for these 12 stocks, 2017 saw a better performance compared with 2018. Image 1 and image 2 shows the table of results through the VBA analysis. In image 1 it was observed that all but 1 stock had positive returns except TERP. 
	![Image1 2017 Stocks Table](VBA_Challenge_2017.PNG) 
	![Image2 2018 Stocks Table](VBA_Challenge_2018.PNG) 
While in image 2 all but 2 stocks had negative returns. Only ENPH and RUN had positive returns at the end of 2018. To achieve this analysis a VBA module was used to iterate through the data sets to produce a summarized version of the data. Firstly an input function was implemented into the code to allow the user to select between 2017 and 2018 data sets. While a for loop was used to iterate through the data sets and the calculations were bound by conditional statements to product the total volume for each ticker symbol. The return was calculated using the end price divided by the starting price minus 1. This provided the percentage of return from each symbol. An example of the code can by found below. 
	![Image3 Example Refactored VBA Code\(Code_example.PNG)
	The 
## 3. Summary 
	3.1. What are the advantages or disadvantages of refactoring code?
		The advantage of refactoring code is it saves time when trying to achieve the end result. Overall calculations are already set in place and general logic guidelines based on the requirements have already been roughly outlined. However,
	3.2. How do these pros and cons apply to refactoring the original VBA script?


