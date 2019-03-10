'RBooth
'AZPHX20190219DATA
'9 March 2019
'####MODERATE Unit 2 | Assignment - The VBA of Wall Street
Sub StockAnalysis()
'Dimension variables
    Dim stockname As String
    Dim lastrow As Long
    Dim summary_table_row As Long
    Dim openprice As Double
    Dim closeprice As Double
    Dim yearlychange As Double

  'Sums Total Volume by Ticker for all sheets as well as calculating percent change in opening to closing price 
  'the yearly price change 
    For Each ws In Worksheets
        stocktotal = 0
        summary_table_row = 2
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        ws.Range("I1").Value = "Ticker"
        ws.Range("L1").Value = "Total stock volume"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        
        For i = 2 To lastrow
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                stockname = ws.Cells(i, 1).Value
                stocktotal = stocktotal + ws.Cells(i, 7).Value
                'To calculate yearly change and percent change
                closeprice = ws.Cells(i, 6).Value
                yearlychange = closeprice - openprice
                percentchange = yearlychange / openprice
                ws.Range("J" & summary_table_row).Value = yearlychange
                'Conditional formatting red/green
                If yearlychange > 0 Then
                    ws.Range("J" & summary_table_row).Interior.ColorIndex = 4
                    Else
                    ws.Range("J" & summary_table_row).Interior.ColorIndex = 3
                    End If
                ws.Range("I" & summary_table_row).Value = stockname
                ws.Range("L" & summary_table_row).Value = stocktotal
                ws.Range("K" & summary_table_row).Value = percentchange
                'Matching "percent" formatting to 2 decimals.
                  ws.Range("K" & summary_table_row).NumberFormat = "0.00%"
                summary_table_row = summary_table_row + 1
                stocktotal = 0
                openprice = 0
            Else
                stocktotal = stocktotal + ws.Cells(i, 7).Value
                If openprice = 0 Then
                    openprice = ws.Cells(i, 3).Value
                End If
            End If
        Next i

    Next ws

    'Resizes Columns to fit
    For Each ws In ActiveWorkbook.Worksheets
    On Error Resume Next
    ws.Activate
    lastCol = ws.Cells(1, Columns.Count).End(xlToLeft).Column
    
    ws.Columns.EntireColumn.AutoFit

  Next ws

Sheets(ws).Select

End Sub