VBA STOCK MARKET ANALYSIS 

This macro runs in every sheet (2018, 2019, 2020.) calcularing and printing the following information: 
In the first summary table (I:L), you see 4 different calculations: a. Ticker symbols b. Yearly change C. Percent Change D. Total stock value. 
In the Second summary table (O:Q) you see the calculation of the a.Greatest % Increase b.Greatest % Decrease c. Greatest Total Volume and the corresponding Ticker Symbol for each calculation. 


The vba macro includes 3 submacros:
1. Ticker (Calls the 2 other macros Conditional Formatting and Summary Table 2)
2. Conditional Formatting 
3. Summary Table 2


- Sub "TickerOnAllSheets": Loops through all the sheets in the workbook to run the "Ticker" macro for each sheet. 
- Sub "Ticker": Calculates yearly changes, and populates a summary table 
- Sub "Conditional_Formatting": Applies conditional formatting to highlight positive and negative changes in the summary table specifically in the Yearly Change Calculation (Green:positive - Red:Negative)
- Sub "Summary2": Identifies the greatest percentage increase, greatest percentage decrease, and greatest total volume, and shows the corresponding ticker symbols.