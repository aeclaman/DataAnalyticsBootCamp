Sub calcSummaryData()

    'for each row in the summary data, find greatest % increase, greatest % decrease and greatest volume
    Dim countSummaryRows, j As Integer
    Dim holdTickerIncr, holdTickerDecr, holdTickerVol As String
    Dim holdGreatestIncr, holdGreatestDecr, holdGreatestVol As Double
   
    countSummaryRows = Application.CountA(Range("J:J"))
   
    For j = 2 To countSummaryRows
        If Cells(j, 11).Value > holdGreatestIncr Then
            holdGreatestIncr = Cells(j, 11).Value
            holdTickerIncr = Cells(j, 9).Value
        ElseIf Cells(j, 11).Value < holdGreatestDecr Then
            holdGreatestDecr = Cells(j, 11).Value
            holdTickerDecr = Cells(j, 9).Value
        End If
        If Cells(j, 12).Value > holdGreatestVol Then
            holdGreatestVol = Cells(j, 12)
            holdTickerVol = Cells(j, 9).Value
        End If
    Next j
   
    'display final values on spreadsheet
    Cells(1, 15).Value = "Ticker"
    Cells(1, 16).Value = "Value"
    Cells(2, 14).Value = "Greatest % Increase"
    Cells(2, 15).Value = holdTickerIncr
    Cells(2, 16).Value = holdGreatestIncr
    Cells(3, 14).Value = "Greatest % Decrease"
    Cells(3, 15).Value = holdTickerDecr
    Cells(3, 16).Value = holdGreatestDecr
    Cells(4, 14).Value = "Greatest Total Volume"
    Cells(4, 15).Value = holdTickerVol
    Cells(4, 16).Value = holdGreatestVol
    Range("P2:P3").NumberFormat = "0.00%"
   
   End Sub
