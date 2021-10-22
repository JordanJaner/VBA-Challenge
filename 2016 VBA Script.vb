Sub StockTest_2016()

    Dim ticker As String
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    yearlyChange = 0
    percentChange = 0
    totalVolume = 0
    summaryTableRow = 2

    For i = 2 To 797711

        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            ticker = Cells(i, 1).Value
            yearlyChange = yearlyChange + (Cells(i, 6).Value - Cells(i, 3).Value)
            percentChange = percentChange + ((Cells(i, 6).Value - Cells(i, 3).Value) / Cells(i, 3).Value)
            totalVolume = totalVolume + Cells(i, 7).Value

            Range("I" & summaryTableRow).Value = ticker
            Range("J" & summaryTableRow).Value = yearlyChange
            Range("K" & summaryTableRow).Value = percentChange
            Range("L" & summaryTableRow).Value = totalVolume
            summaryTableRow = summaryTableRow + 1

            yearlyChange = 0
            percentChange = 0
            totalVolume = 0

        Else

            yearlyChange = yearlyChange + (Cells(i, 6).Value - Cells(i, 3).Value)
            percentChange = percentChange + ((Cells(i, 6).Value - Cells(i, 3).Value) / Cells(i, 3).Value)
            totalVolume = totalVolume + Cells(i, 7).Value

        End If

    Next i

End Sub
