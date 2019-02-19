Sub testHardHW(wsName As String)

    Dim x As Integer
    Dim lr As Long
    Dim holdTickerValue, holdDate As String
    Dim holdOpeningBid, holdClosingBid As Double

    'numer of rows in the sheet
    lr = Cells(Sheets(wsName).Rows.Count, 1).End(xlUp).Row

    'variable to define row where new data goes, will be increased as each new ticker is found
    x = 2

    'initialize holding variable of opening bid to first row data
    holdOpeningBid = Cells(2, 3).Value

    'Display New Column Headers and formatting
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Volume"
    Range("J2:J" & lr).NumberFormat = "0.00000"
    Range("K2:K" & lr).NumberFormat = "0.00%"

    'for each row, sum the volume by unique ticker, grab opening bid, grab closing bid and perform necessary calcs
    'assumption that first row of ticker data is opening and last row of ticker data is closing
    For i = 2 To lr

        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            holdTickerValue = Cells(i, 1).Value
        
            Cells(x, 9).Value = holdTickerValue
            Cells(x, 12).Value = CDbl(Cells(x, 12).Value) + CDbl(Cells(i, 7).Value)
        
            'grab closing bid for current Ticker
            holdClosingBid = Cells(i, 6).Value
            'calculate yearly change for current Ticker
            Cells(x, 10).Value = holdClosingBid - holdOpeningBid
            'calculate percent change and account for divide by zero errors
            If holdOpeningBid <> 0 Then
                Cells(x, 11).Value = (holdClosingBid - holdOpeningBid) / holdOpeningBid
            Else
                Cells(x, 11).Value = 0
            End If
                
            'hold opening bid for next Ticker
            holdOpeningBid = Cells(i + 1, 3).Value
            'initialize closing bid for next Ticker
            holdClosingBid = 0
        
            'increment next output row
            x = x + 1
        
        Else
            'keep a running total of the current tickers volume
            Cells(x, 12).Value = CDbl(Cells(x, 12).Value) + CDbl(Cells(i, 7).Value)
 
        End If
    Next i
    
    'call function to run the final summaries
    Call calcSummaryData

End Sub
