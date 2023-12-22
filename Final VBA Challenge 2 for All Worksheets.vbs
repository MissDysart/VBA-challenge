Attribute VB_Name = "Module1"
Sub abcStocks()

For Each ws In Worksheets

'Create headers
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"

'Declare variables
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
Dim ticker As String
Dim volume_total As Double
volume_total = 0


'To store findings in columns
Dim percent_change As Double
Dim sum_table_row As Integer
sum_table_row = 2


'Loop through all ticker information
For i = 2 To lastrow
    stock_open = ws.Cells(2, 3).Value
    'Look for new values in ticker column
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
        'Set the ticker and closing price
        ticker = ws.Cells(i, 1).Value
        stock_close = ws.Cells(i, 6).Value
        
        'Calculate yearly change and percent change
        yearly_change = stock_close - stock_open
        percent_change = (stock_close - stock_open) / stock_open
        
        
        'Add to stock volume total
        volume_total = volume_total + ws.Cells(i, 7).Value
        
        'Print ticker, total stock volume in summary
        ws.Range("I" & sum_table_row).Value = ticker
        ws.Range("J" & sum_table_row).Value = yearly_change
        ws.Range("K" & sum_table_row).Value = percent_change
        ws.Range("L" & sum_table_row).Value = volume_total
        
        'Formatting stuff
        If yearly_change > 0 Then
            ws.Range("J" & sum_table_row).Interior.ColorIndex = 4
        Else
            ws.Range("J" & sum_table_row).Interior.ColorIndex = 3
        End If
        ws.Range("K" & sum_table_row).NumberFormat = "0.00%"
        
        'Go to the next row, reset stock volume and set open stock price
        sum_table_row = sum_table_row + 1
        volume_total = 0
        stock_open = ws.Cells(i + 1, 3)
  
    'If the next cell is the same ticker
    Else
        
        'Add to the total stock volume
        volume_total = volume_total + ws.Cells(i, 7).Value
  
    End If
  
Next i


'Create summary headers
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"

'Calculate min and max values
ws.Range("Q2") = "%" & WorksheetFunction.Max(Range("K2:K" & sum_table_row)) * 100
ws.Range("Q3") = "%" & WorksheetFunction.Min(Range("K2:K" & sum_table_row)) * 100
ws.Range("Q4") = WorksheetFunction.Max(Range("L2:L" & sum_table_row)) * 100

'match min and max to Ticker

' Autofit to display data
ws.Columns("A:Q").AutoFit

Next ws

End Sub

