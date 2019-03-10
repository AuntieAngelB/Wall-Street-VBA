

'RBooth
'AZPHX20190219DATA
'9 March 2019
'####EASY Unit 2 | Assignment - The VBA of Wall Street

Sub StockAnalysis()

'Sums Total Volume by Ticker for all sheets


'Defining Dimensions for Worksheet(in Worksheets) Ticker, Total Volume Sum(TSV), and Total Rows

    Dim ws As Worksheet
    Dim Ticker As String
    Dim TSV As Double
    Dim totalsrow As Integer

    For Each ws In Worksheets
    ws.Activate
    
'Set Total Volume Sum(TSV) to zero for loops

    TSV = 0

    ' Set starting point for rows to first value
    
    totalsrow = 2
    
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Total Stock Volume"
    
    ' Defines the Last Row (LastRow) in data
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
        For i = 2 To LastRow

            ' Check if we are still within the same Ticker symbol, if it is not...
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

                ' Set the Ticker Value
                Ticker = Cells(i, 1).Value

                ' Add to TSV
                TSV = TSV + Cells(i, 7).Value

                ' Print Ticker in the Summary Table
                ws.Range("I" & totalsrow).Value = Ticker

                ' Print the TSV to the Summary Table
                ws.Range("J" & totalsrow).Value = TSV

                ' Add one to the summary table row
                totalsrow = totalsrow + 1
                
                ' Reset TSV
                TSV = 0

            Else

                ' Add to the Total Stock Volume
                TSV = TSV + Cells(i, 7).Value

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
