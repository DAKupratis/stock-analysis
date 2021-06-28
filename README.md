# Module 2 Challenge - VBA of Wall Street

### PURPOSE: The purpose of this challenge was to make the code we had prepared during the module to analyze volume and returns of 12 stocks more efficient, so that it would be able to be more efficiently utilized for analysis of a much larger number of stocks. We tested this by using the Timer function to measure how long the code took the run according to both methods. While there are a tenths-of-a-millisecond differences in the time it takes to run both sets of code, that could add up very quick if large amounts of real-time or historic stock market data is being interpreted, let alone across multiple stocks. Thus, cutting that time down might be imperative to a trading strategy!


### RESULTS: The first analysis of the stock data for AllStocksAnalyis() **"yearValueAnalysis()" in my VBA script** sets out a For loop structure only utilizing arrays for the stock tickers, unlike in AllStocksAnalysisRefactored() which matches the number of tickers in the tickers array with the with the arrays for volume, starting price and ending price, one for one. We only introduce variables for AllStocksAnalysis(). 

OLD
https://github.com/DAKupratis/stock-analysis/blob/main/Resources/allstocks1.png

NEW
https://github.com/DAKupratis/stock-analysis/blob/main/Resources/refactor1.png

In AllStocksAnalysisRefactored(), we then use a "blank" variable, tickerIndex, to project each iteration of a For loop onto when grabbing data per-ticker. This helps to align the data belonging to 1 of the 12 tickers(i) variables with the corresponding 1 of the 12 ultimate tickerVolumes/tickerStartingPrice/tickerEndingPrice variables with the rest. The first For loop then opens to set the volume to zero, so that it can be reset every time the loop runs and those number don't commingle between stocks. 

OLD
https://github.com/DAKupratis/stock-analysis/blob/main/Resources/allstocks2.png

NEW
https://github.com/DAKupratis/stock-analysis/blob/main/Resources/refactor2.png

In the AllStocksAnalysis() sub, the loop is structured so that it runs again per each row, and not as one continuous loop that passes over all of the data belonging to a ticker in one operation per iteration as AllStocksAnalyisRefactor() does. This functionality is added by step 3d, which increases the ticker index by one, allowing it continue going over all the rows belonging to a stock in one loop before moving onto the next loop. Otherwise, formatting, timers, column headers and the like retain their spots.


Having to run less loops overall for the same data grab appears to save a very considerable amount of time for AllStocksRefactored() - roughly 1/3rd compared to the original code!

2017
OLD
https://github.com/DAKupratis/stock-analysis/blob/main/Resources/2017_Old_Sub_Speed.png

NEW
https://github.com/DAKupratis/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png


2018 
OLD
https://github.com/DAKupratis/stock-analysis/blob/main/Resources/2018_Old_Sub_Speed.png

NEW
https://github.com/DAKupratis/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png

PS: I put buttons on the spreadsheet so you can compare the old vs. new macros in real time and clear them when you're done. 


###SUMMARY: Refactoring code is a great way to make code run more efficiently and dynamically, potentially taking up less time for the same outputs and making the code more adaptable. This streamlined code also helps to remove more potential points of failure/errors ("bugs") when writing and maintaining the code should edits need to be made, which altogether make the code easier to read and understand. This up-front investment of time (and maybe money!) could go a long way to reducing the overall operating costs when working on a real-life project. Some disadvantages of refactoring code might be that you might not always succeed, and might have wasted some precious time. This may be due to your code being at its most optimal, or because errors were made in the refactoring. This is likely exemplified as the amount of code being refactored increases. 


The upside of this when working on our VBA script was that we did ultimately get the code to run faster, and learned a way to reduce computing time and energy when creating loops by making them be able to pull more data per iteration. The tough part is that it took me a decent while longer than to write the original code, though the only way I had ever learned to write the loops prior was via the original method, so that might due to me being a beginner - this took numerous attempts, all for the same result. It would depend, as mentioned before, how often these loops need to run for the time differences to add up and make an impact. 
