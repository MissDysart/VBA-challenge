# VBA-challenge
Module 2 Challenge

Objectives:
1. Check each row of the <ticker> column and put name of ticker in a summary table in column (or Range) "K"
2. For each <ticker> find the daily change (close minus open) and add together to get Yearly Change. Print total in column "L"
3. For each <ticker> calculate the percent of change (open figure divided by the closing figure). Print to column "M"
4. For each <ticker> add the daily volume and print total to column "N"

Relied heavily on the credit card charge example and Census example from Module 2, day 3 (VBA). The "lastrow" formula, ColorIndex variables from Module 2, day 3.


Issues:
I couldn't figure out how to use the Match or Lookup Functions to get the Ticker that in Greatest % Increase, Greatest % Decrease, Greatest Total Volume. This is how far I got:

'match min, max, volume to Ticker
Dim lookupRange As Range
Dim resultRange As Range
Dim result As Variant
Dim lastrowK As Long
Dim lastrowI As Long
Dim lookupValue As Variant

lastrowK = Cells(Rows.Count, "K").End(xlUp).Row
lastrowI = Cells(Rows.Count, "I").End(xlUp).Row
lookupValue = Range("P2") ' The value you want to look up

Set lookupRange = Range("K2:K" & lastrowK) ' The range where you want to perform the lookup
Set resultRange = Range("I2:I" & lastrowI) ' The range from which you want to retrieve the result

result = WorksheetFunction.VLookup(lookupValue, lookupRange, resultRange.Column - lookupRange.Column + 1, False)