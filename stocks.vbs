Sub stocks()
Dim stockSheet As Worksheet

'loop through each stock sheet
For Each stockSheet In Worksheets
    'declare variables
    Dim curr As Integer
    curr = 1
    Dim OpenT As Double
    Dim CloseT As Double
    Dim difT As Double
    Dim maxStock As Double
    Dim headerLine As Integer
    Dim incMax As Double
    Dim decMax As Double
    Dim vol As Double
    headerLine = 1
    'set header titles for stock analysis new column data
    stockSheet.Cells(headerLine, 9).Value = "Ticker"
    stockSheet.Cells(headerLine, 10).Value = "Yearly Change"
    stockSheet.Cells(headerLine, 11).Value = "Percent Change"
    stockSheet.Cells(headerLine, 12).Value = "Total Stock Volume"
    stockSheet.Cells(headerLine + 1, 14).Value = "Greatest % Increase"
    stockSheet.Cells(headerLine + 2, 14).Value = "Greatest % Decrease"
    stockSheet.Cells(headerLine + 3, 14).Value = "Greatest Total Volume"
    stockSheet.Cells(headerLine, 15).Value = "Ticker"
    stockSheet.Cells(headerLine, 16).Value = "Value"

    curr = curr + 1
    'loop through rows to get running totals for each stock based on tickers
    For ind = 2 To stockSheet.Range("A1").End(xlDown).Row
        maxStock = maxStock + stockSheet.Cells(ind, 7).Value
        If stockSheet.Cells(ind, 1).Value <> stockSheet.Cells(ind - 1, 1).Value Then
            OpenT = stockSheet.Cells(ind, 3).Value
        End If
        If stockSheet.Cells(ind, 1).Value <> stockSheet.Cells(ind + 1, 1).Value Then
            CloseT = stockSheet.Cells(ind, 6).Value
            difT = CloseT - OpenT
            stockSheet.Cells(curr, 9).Value = stockSheet.Cells(ind, 1).Value
            stockSheet.Cells(curr, 10).Value = difT
            'check for zero division error
            If OpenT = 0 Then
                stockSheet.Cells(curr, 11).Value = "N/A"
            Else
                stockSheet.Cells(curr, 11).Value = stockSheet.Cells(curr, 10) / OpenT
            End If
            stockSheet.Cells(curr, 12).Value = maxStock
            'increment counter
            curr = curr + 1
            maxStock = 0
        End If
    Next ind
    'find values for greatest percent increase, percent decrease, and greatest volume
    incMax = Application.WorksheetFunction.Max(stockSheet.Range("K:K"))
    stockSheet.Cells(2, 16) = incMax
    stockSheet.Cells(2, 16).NumberFormat = "0.00%"
    decMax = Application.WorksheetFunction.Min(stockSheet.Range("K:K"))
    stockSheet.Cells(3, 16) = decMax
    stockSheet.Cells(3, 16).NumberFormat = "0.00%"
    vol = Application.WorksheetFunction.Max(stockSheet.Range("L:L"))
    stockSheet.Cells(4, 16) = vol
    stockSheet.Cells(4, 16).NumberFormat = "##0.0E+0"
    'set values for bonus analysis in table
    For curRow = 2 To stockSheet.Range("I1").End(xlDown).Row
        If stockSheet.Cells(curRow, 11) = incMax Then
            stockSheet.Cells(2, 15).Value = stockSheet.Cells(curRow, 9).Value
        End If
        If stockSheet.Cells(curRow, 11) = decMax Then
            stockSheet.Cells(3, 15).Value = stockSheet.Cells(curRow, 9).Value
        End If
        If stockSheet.Cells(curRow, 12) = vol Then
            stockSheet.Cells(4, 15).Value = stockSheet.Cells(curRow, 9).Value
        End If
    Next curRow
    
    'color cells red or green for negative/positive conditionals
    For x = 2 To stockSheet.Range("A1").End(xlDown).Row
        If stockSheet.Cells(x, 10).Value > 0 Then
            stockSheet.Cells(x, 10).Interior.ColorIndex = 4
        ElseIf stockSheet.Cells(x, 10).Value < 0 Then
            stockSheet.Cells(x, 10).Interior.ColorIndex = 3
        End If

    Next x
    'format cells to appropriate percent/currency/alignment
    stockSheet.Columns("J").NumberFormat = "$0.00"
    stockSheet.Columns("K").NumberFormat = "0.00%"
    stockSheet.Cells.EntireColumn.AutoFit
    stockSheet.Range("Q2, Q3").NumberFormat = "0.00%"
Next stockSheet
End Sub
