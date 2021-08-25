# Automating Stock Performance Analysis with Excel VBA

## Overview of Project

Our friend Steve has asked us to help him analyze a group of alternative energy stocks for his parents to invest in. We will use our newly acquired skills in VBA to create a script in an excel workbook that can sweep through stock data for 12 different companies and report on their traded volume and stock performance for 2017 and 2018.

### Methodology

We have been provided with a data table containing historicals for the 12 stocks, with columns corresponding to ticker symbol, trading date, open, hi, low, close, adjusted close and trading volume. The dataset was divided into two tables, one for 2017 and one for 2018. Each table was incorporated into an excel workbook as individual worksheets. Using VBA and additional worksheets to automate the analysis, we created different scripts to perform calculation and formatting of reported results. We also added buttons to make the process intuitive and convenient. 

Once we had our basic code running and producing the correct results, we focused our effort in optimizing the code through refactoring in order to improve run time. To do that, we added three new arrays to the code to capture the information in one sweep of the data. For each ticker, these arrays included traded volume, starting price, and ending price. The arrays were dimensioned according to their corresponding data types. 

![Created and dimensioned new arrays to hold the volume, starting price and ending price for each ticker.](https://github.com/IJG-DR/stock-analysis/blob/7fa92aa74855c72432d8d699f3fed5335582549c/Resources/Dimensioned_Arrays.png)

We also added code to initialize the volume data for each ticker to 0.

![Created a loop to set values in the Volume array to cero.](https://github.com/IJG-DR/stock-analysis/blob/7fa92aa74855c72432d8d699f3fed5335582549c/Resources/Created_Loop_to_Set_Volumes_to_Zero.png)

Additionally, we added code to the loop that sweeps over all the data by rows in order to accumulate volume data for each ticker, capture starting price and ending price for each ticker, and move to the next ticker once data for the previous ticker had been captured.

![Created a loop code to go over all rows and record volumes, starting price and ending price for each ticker, as well as moving to the next ticker when done.](https://github.com/IJG-DR/stock-analysis/blob/7fa92aa74855c72432d8d699f3fed5335582549c/Resources/Loop_Through_All_Rows.png)

Finally, we introduced code in order to generate the output for each ticker by looping through the ticker array and reporting the corresponding volume and return data.

![Created a loop to read from the data arrays and report the volume and return for each ticker.](https://github.com/IJG-DR/stock-analysis/blob/7fa92aa74855c72432d8d699f3fed5335582549c/Resources/Loop_Through_Arrays_to_Report_Results.png)

## Results

Running our stock analysis code on the data available for the 12 stocks produced the following results:

![Summary Table of Stock Returns for 2017](https://github.com/IJG-DR/stock-analysis/blob/7fa92aa74855c72432d8d699f3fed5335582549c/Resources/Stock_Performance_2017.png)

![Summary Table of Stock Returns for 2018](https://github.com/IJG-DR/stock-analysis/blob/7fa92aa74855c72432d8d699f3fed5335582549c/Resources/Stock_Performance_2018.png)

Of all the stocks analyzed for the period covering 2017 through 2018, ENPH had the best returns, being the only stock to report the highest positive returns for both years combined. Only one other stock had positive returns on both years, RUN. The rest had positive returns in 2017 (except TERP), and negative returns in 2018. TERP was the only stock in the group with negative returns in 2017 as well as in 2018, although SPWR had the worst cumulative return. Aside from ENPH and RUN, SEDG was also a stock that had a very high return in 2017 and a minimal loss in 2018, which would also make it a stock to recommend as it seems 2018 in general was a bear market and the stock fared better than most in the group.

## Summary

With the refactored code, we have provided Steve with an analysis tool that can compare stock performance conveniently and quickly for his parents.

The refactored code produced a substantial improvement in the runtimes of the routines as evidenced by the message box posted results. Compared with the original code which had a runtime of 1.179688 seconds for 2017, the refactored code ran in 0.1835938 seconds, nearly 6.5 times faster!

![Run time image for 2017](https://github.com/IJG-DR/stock-analysis/blob/4e24c7c675ea229823560fa8e42a0ef4e1bede8a/Resources/VBA_Challenge_2017.png)

As for the 2018 runtimes, the original code ran for 1.113281 seconds, while the refactored code ran for only 0.1757812 seconds, or 6.3 times faster.

![Run time image for 2018](https://github.com/IJG-DR/stock-analysis/blob/4e24c7c675ea229823560fa8e42a0ef4e1bede8a/Resources/VBA_Challenge_2018.png)

As a result of our analysis, we are recommending Steve's parents to focus their interest in ENPH, RUN and SEDG.

### General Advantages and Disadvantages of Refactoring Code

In general terms, refactoring produces more efficient scripts that are able to run faster, and generally have fewer lines of code. They also tend to be more complex and may be more difficult to interpret as you may have more moving parts occurring at the same time. It may also require the use of more variables and complex structures such as arrays, which makes it more important to keep track of what each variable is doing when writing the code. This may require testing small sections of the code to ensure that the parts are working correctly. Also, with more variables, we noticed that checking the spelling of the different variables was necessary in many places to ensure that the code was updating each variable correctly. Finally, since we were reutilizing code from our earlier test scripts, it was important to update previous variable names to the new ones.

### Advantages and Disadvantages of the Original and Refactored VBA Script

The refactored code produced a notable improvement in run times, as a result of a more efficient read through the data table by capturing all the required information in a single sweep. The code is also more compact and only needs one for-next loop to run over the data, as opposed to nested for-next loops. One minor disadvantage in refactoring is that, in the absence of appropriate commenting, it would be harder to interpret at first glance as opposed to the original code. There are a few other improvements we would consider making in the future. The script assumes that the data has been previously sorted by ticker and date, which might not always be the case. We could improve on this by adding lines to the code that would sort the data accordingly to ensure it is properly organized to work with our script. Also, the line of code that counts the number of rows in the data table used to establish the ending row parameter in the for-next loop may provide the wrong results if the data table had other unrelated data bellow the stock data we were analyzing (separated by blank rows). By going to the bottom of the sheet, and then finding the first row from the bottom with data, it would have returned a row number corresponding to the end of this unrelated data. It is therefore important when using this code in the future to ensure that there is no unrelated data below the data we want to analyze.