# stock-analysis

# Background
In this homework assignment, I used VBA scripting to analyze generated stock market data.

# Process
1. Created a script that loops through all the stocks for one year and outputs the following information:
  - The ticker symbol.
  - Yearly change from the opening price at the beginning of a given year to the closing price at the end of that year.
  - The percentage change from the opening price at the beginning of a given year to the closing price at the end of that year.
  - The total stock volume of the stock.

2. Added functionality to the script to return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume".

3. Made the appropriate adjustments to the VBA script to enable it to run on every worksheet (that is, every year) at once.

# References
This submisson contains refferences to code principles found in class and searched from web. See below for code location and its reference

1. line 22 - "For Each ws" loop concept is borrowed from solved class activity from Week 2/Class 3
2. line 28 - "lastrowcount" concept is taken from solved class actvity
3. line 69 - "NumberFormat" as percentage concept is taken from ExcelHowTo: https://www.excelhowto.com/macros/formatting-a-range-of-cells-in-excel-vba/
4. line 86 - "AutoFit" concept is taken from here:https://learn.microsoft.com/en-us/office/vba/api/excel.range.autofit
