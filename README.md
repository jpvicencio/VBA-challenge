# VBA_challenge
VBA exercise
	1. Define variables
		Set Q1 as target worksheet
	2. Create a summary table with Ticker, Quarterly Change, Percentage Change, Total Stock Volume
	  Assign headers for the summary table
	  Define variables and range to compute for the values in the summary table
	  Loop through all rows
	  Define the formula to compute the values
		  Quarterly change - close price of the last date for the quarter minus open price of the first date of the quarter  
	  Apply conditional formatting
		  Make Quarterly Change green if the percentage change is positive number, red if negative number and no color if 0. 
	  Write the results in the summary table
	3. Get the greatest % increase, greatest % decrease and greatest total volume  from the summary table.
	  Add new variables
	  Loop through the data
	  Define formula to get the greatest % increase, greatest % decrease and greatest total volume
	  Write the results with the headers: Greatest % Increase, Greatest % Decrease, Greatest Total Volume, Ticker and   Value and their corresponding values
	4. Run the script to see if it works in Q1.
		Copy code to sheets Q2, Q3 and Q4, updating the worksheet name.
		Run to see if working for each worksheet.
	5. Once it works for all worksheets, put the code in 'ThisWorkbook'.
	  Make the necessary edits after variables, to loop through each worksheet
		  Read worksheets from left with 'Q'

	Sources: Learning Assistant, ChatGPT
