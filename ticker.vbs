Sub volume()
'create a variable for worksheet
Dim ws As Worksheet

'loop through every worksheet
For Each ws In Worksheets

'create a variable for holding ticker name
Dim ticker As String

'create a variable for holding vol per ticker
Dim vol_total As Double
vol_total = 0

'track the location of each ticker in the summary table
Dim summary_table_row As Integer
summary_table_row = 2

'variable for yearly change and percentage change

Dim yearlychange As Double
Dim percentagechange As Double

'variable for first first row of each ticker

Dim ticker_first_row As Double

'Dim currentrow As Integer
'set a value for  first row of each ticker
ticker_first_row = 2

'determine the last row

LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'loop through the tickers

For currentrow = 2 To LastRow

'check if we are in the same ticker,if not then
If ws.Cells(currentrow + 1, 1).Value <> ws.Cells(currentrow, 1).Value Then

'select ticker
selectticker = ws.Cells(currentrow, 1).Value

'calculate yearlychange by finding difference between closing price of last row and opening price of first row for each ticker

yearlychange = ws.Cells(currentrow, 6).Value - ws.Cells(ticker_first_row, 3).Value

'if opening price in first row=0 then print 999999

If ws.Cells(ticker_first_row, 3).Value = 0 Then
    percentagechange = 1E+32
    
Else
    percentagechange = (yearlychange / ws.Cells(ticker_first_row, 3).Value) * 100
End If


'add the volume
vol_total = ws.Cells(currentrow, 7).Value + vol_total

'print ticker in summary table
ws.Range("I" & summary_table_row) = selectticker

'print yearly change
ws.Range("J" & summary_table_row) = yearlychange

'assign color index to yearly change

If yearlychange >= 0 Then
ws.Range("J" & summary_table_row).Interior.ColorIndex = 4

Else
ws.Range("J" & summary_table_row).Interior.ColorIndex = 3
End If


'print percentage change
ws.Range("K" & summary_table_row) = percentagechange

'remove color index from column(had assigned color index to column k and then removed)

ws.Range("K" & summary_table_row).Interior.Color = xlNone



'print volume in summary table
ws.Range("L" & summary_table_row) = vol_total

'ws.Range("M" & summary_table_row) = LastRow
'ws.Range("N" & summary_table_row) = ticker_first_row

'set vol_total as 0 to calculate total again for next ticker
vol_total = 0
'go to the next row
summary_table_row = summary_table_row + 1

'If ticker_first_row < LastRow Then

'go to the first row of the next ticker
ticker_first_row = currentrow + 1

'End If

Else

vol_total = vol_total + ws.Cells(currentrow, 7).Value

End If
Next currentrow

Next ws


End Sub

