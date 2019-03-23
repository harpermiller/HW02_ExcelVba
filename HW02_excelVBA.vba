Sub stockExchange()
Dim annualDifference As Single
Dim tickerSymbol As String
Dim resultRow As Integer
'Dim percentChange As

resultRow = 2
'annualDifference = 0
totalVolume = 0
percentChange = 0
lastRow = Cells(Rows.Count, 1).End(xlUp).Row

    For i = 2 To lastRow
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            tickerSymbol = Cells(i, 1).Value
            Cells(resultRow, 9).Value = (Cells(i, 1).Value)

            totalVolume = totalVolume + Cells(i, 7).Value
            Cells(resultRow, 10).Value = totalVolume

            'annualDifference = ((Cells(i, 6).Value) - (Cells(i, 3).Value))
            'Cells(resultRow, 10).Value = annualDifference

            'percentChange = ((Cells(i, 3).Value) / (Cells(i, 6).Value))
            'Cells(resultRow, 11).Value = percentChange

            resultRow = resultRow + 1

            totalVolume = 0
            annualDifference = 0
        Else
            Annual_Difference = (Annual_Difference) + ((Cells(i, 6).Value) - (Cells(i, 3).Value))
            totalVolume = totalVolume + (Cells(i, 7).Value)
        End If
' conditional that changes color to red if negative
        'If Cells(i, 10).Value < 0 Then
            'Cells(i, 10).Interior.Color = RGB(200, 0, 0)
' conditional that changes color to green if positive
        'ElseIf Cells(i, 10).Value > 0 Then
            'Cells(i, 10).Interior.Color = RGB(0, 128, 0)
' keeps color change only for fields with values in them
        'ElseIf Cells(i, 10).Value = Null Then
            'Cells(i, 10).Interior.Color = RGB(0, 0, 0)
        'End If

    Next i

End Sub
