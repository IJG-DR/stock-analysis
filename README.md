# stock-analysis
Module 2 work repository

First commit uploaded file in macro-enabled excel type, after enabling VBA/Macros and creating subroutine MacroCheck() to test VBA was working

Second commit uploaded new version of the macro-enabled excel file after adding a new DQ Analysis worksheet tab and VBA subroutines to fill out headers in cells A1, A3, B3 and C3 on the new worksheet

Thrid commit uploaded a new version with revised macro that calculates total volume traded for stock ticker DQ during 2018 using for-next loop and conditional if-else statements

Fourth commit uploaded adds code to calculate row total and creates variables to track starting price and ending price for stock DQ in order to calculate yearly return.

Fifth commit adds a worksheet called "All Stocks Analysis" and a new macro called AllStocksAnalysis. For the moment, the macro only creates labels, dimensions an array to hold ticker symbols and runs some nested loops to: (1) fill first ten rows and columns with the value 1, (2) then fills the first ten rows and columns with the value of the sum of the row and column numbers and (3) clears the contents of the first ten rows and columns.

Sixth commit uses for-next loops and conditional statements to generate traded volume and return of all stocks in the dataset for 2018

Seventh commit added font and border formatting both for numbers and headers, as well as conditional formatting for returns: green for positive returns, red for negative and clear for 0 return. Also added a macro to create an 8x8 checkerboard.

Eighth commit adds buttons to run the analysis and clear the results, as well as code to input the desired year.
