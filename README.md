# Overview Of Project
We need to analyze the yearly change in stock prices, calculate percent changes from starting price and end of the year closing price. 
We also need to identify greatest increase stock, greatest decrease stock and the greatest traded volume. 
This analysis needs to be done on the given dataset on Multiple Year Stock Data, and analysis results should be added in the given excel sheet itself.

# Aim or Purpose
The purpose of this project is to analyze a dataset provided, and identify the change in stock price over the year. We will also identify the top increased and decreased stock in the years time and repeat this analysis for 3 years.

# Step by Step analysis details for identifying Stock variations
Created a script that loops through all the stocks for one year and outputs the following information:
- The ticker symbol in column I
- Yearly change from the opening price at the beginning of a given year to the closing price at the end of that year, in column J
- The percentage change from the opening price at the beginning of a given year to the closing price at the end of that year, in column K
- The total stock volume of the stock in column L
- We also applied conditional formatting on column J, if Yearly change is positive then green, and if its negative then red. This color coding helps in analyzing the result quickly. 
- applied conditional formatting for percent change using Graded color scale between Red to Green based on the values in the column K.

# Finding greatest increase and decrease in yearly change in stocks
- Wrote a script to find the greatest increase value and the ticker, and wrote in cell P2 and Q2
- Wrote a script to find the greatest decrease value and ticker, and wrote in cell P3 and Q3
- wrote a script to find the greatest traded stock and its traded volume, and wrote in cell P4 and Q4

# Performed the analysis for 2018, 2019 and 2020.
- Wrote a script to run the analysis for all sheet of data available in all 3 sheets in the workbook
- Traversed through all the available sheets and executed the program for each available sheet.

# Artifacts
- tickerAnalysis.vbs, contains code of the macro which is run on the dataset provided in workbook.
- MultiyearStockAnalysis_2018.png, MultiyearStockAnalysis_2019.png, MultiyearStockAnalysis_2020.png images are screen shots of results for quick lookup of results.
- Multiple_year_stock_data.xlsx is the dataset file, where the real data and analysis results reside. Macro enabled sheets are with macro. 