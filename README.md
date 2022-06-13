# Stock-Analysis using VBA and Excel
## Overview
* In this project and analyisis, we’ll edit, or refactor, the Stock Market Dataset with VBA solution code to loop through all the data one time in order to collect an entire dataser. Then, we’ll determine whether refactoring your code successfully made the VBA script run faster. Finally, we just want to make the code more efficient—by taking fewer steps, using less memory, or improving the logic of the code to make it easier for future users to read. 
 >"Use your knowledge of VBA and the starter code provided in this Project to refactor the VBA Script dataset so we loop through the data one time and collect all of the information."

## Deliverables:
**1. The `tickerIndex` is set equal to zero before looping over the rows.**
> Created a `tickerIndex` variable and set it equal to zero before iterating over all the rows. Will use this `tickerIndex` to access the correct index across the four different arrays on VBA Code: the tickers array and the three output arrays created on next requierement.

**2. Arrays are created for `tickers`, `tickerVolumes`, `tickerStartingPrices`, and `tickerEndingPrices`.**

**3. The `tickerIndex` is used to access the stock ticker index for the `tickers`, `tickerVolumes`, `tickerStartingPrices`, and `tickerEndingPrices` arrays.**

**4. The script loops through stock data, reading and storing all of the following values from each row: `tickers`, `tickerVolumes`, `tickerStartingPrices`, and `tickerEndingPrices`.**

**5. Code for formatting the cells in the spreadsheet is working.**

**6. There are comments to explain the purpose of the code.**

**7. The outputs for the 2017 and 2018 stock analyses in the `VBA_Challenge.xlsm` workbook match the outputs from the AllStockAnalysis in the module**


## Results:

> Final VBA Analysis
***Final VBA Analysis 2017***


![image](https://github.com/SLCunningham21/Stock-Analysis/blob/main/Resources/VBA_Challenge_2017.PNG)


***Final VBA Analysis 2018***

![image](https://github.com/SLCunningham21/Stock-Analysis/blob/main/Resources/VBA_Challenge_2018.PNG)


***Time on VBA_Challenge_2017.PNG***

![image](https://github.com/SLCunningham21/Stock-Analysis/blob/main/Resources/Time%20for%202017%20analysis.PNG)

***Time on VBA_Challenge_2018.PNG***

![image](https://github.com/SLCunningham21/Stock-Analysis/blob/main/Resources/Time%20for%202018%20analysis.PNG)

## Summary:

**1. What are the advantages or disadvantages of refactoring code?**
**Disadvantages:**

 - Possibility for duplicate lines with a long procedure code
 - Logical structure is best moved to a new function and called from the other functions.
 - Refactoring process can affect the testing outcomes. 


**Advantages:**
 - Logical errors easily appear in well structure code that contains nested conditionals and loops. 
 - In our case, using Excel flow displays program logic in a more comprehensible manner, not tied to the order that the underlying code is written.
 - VBA interpretation (Excel) of code can reveal patterns that are not easy to see in the source.


**2. How do these pros and cons apply to refactoring the original VBA script?**

- A more logical structure to the code, without multiple duplicates, makes it easier to restructure code in the future and also easier to be understood by someone else reading the code.





