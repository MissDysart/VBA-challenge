Attribute VB_Name = "Module1"
Sub abcStocks()

'Declare variables
lastrow = Cells(Rows.Count, 1).End(xlUp).Row
Dim ticker As String
Dim volume_total As Double
volume_total = 0

'Create headers
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"


'To store findings in columns
Dim percent_change As Double
Dim sum_table_row As Integer
sum_table_row = 2


'Loop through all ticker information
For i = 2 To lastrow
    stock_open = Cells(2, 3).Value
    'Look for new values in ticker column
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
        'Set the ticker and closing price
        ticker = Cells(i, 1).Value
        stock_close = Cells(i, 6).Value
        
        'Calculate yearly change and percent change
        yearly_change = stock_close - stock_open
        'or (stock_close - stock_open) / stock_open
        percent_change = stock_close / stock_open - 1
        
        'Add to stock volume total
        volume_total = volume_total + Cells(i, 7).Value
        
        'Print ticker, total stock volume in summary
        Range("I" & sum_table_row).Value = ticker
        Range("J" & sum_table_row).Value = yearly_change
        Range("K" & sum_table_row).Value = percent_change
        Range("L" & sum_table_row).Value = volume_total
        
        'Formatting stuff
        If yearly_change > 0 Then
            Range("J" & sum_table_row).Interior.ColorIndex = 4
        Else
            Range("J" & sum_table_row).Interior.ColorIndex = 3
        End If
        Range("K" & sum_table_row).NumberFormat = "0.00%"
        
        'Go to the next row, reset stock volume and set open stock price
        sum_table_row = sum_table_row + 1
        volume_total = 0
        stock_open = Cells(i + 1, 3)
  
    'If the next cell is the same ticker
    Else
        
        'Add to the total stock volume
        volume_total = volume_total + Cells(i, 7).Value
  
    End If
  
Next i


'Create summary headers
Range("O2").Value = "Greatest % Increase"
Range("O3").Value = "Greatest % Decrease"
Range("O4").Value = "Greatest Total Volume"
Range("P1").Value = "Ticker"
Range("Q1").Value = "Value"

'Calculate min and max values
Range("Q2") = "%" & WorksheetFunction.Max(Range("K2:K" & sum_table_row)) * 100
Range("Q3") = "%" & WorksheetFunction.Min(Range("K2:K" & sum_table_row)) * 100
Range("Q4") = WorksheetFunction.Max(Range("L2:L" & sum_table_row)) * 100

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


' Autofit to display data
Columns("A:Q").AutoFit

End Sub

