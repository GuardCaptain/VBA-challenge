Attribute VB_Name = "Module1"
Sub ticker()

Range("I1") = "Ticker"

'keeps track of summary table row (vital)
Dim summarytablerow As Double
summarytablerow = 2

For i = 2 To 753000

If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then

Cells(summarytablerow, 9) = Cells(i, 1).Value
summarytablerow = summarytablerow + 1

End If
Next i
End Sub

Sub yearlychange()

Dim summaryrow As Double
Dim start1 As Double
Dim end1 As Double

summaryrow = 2
Range("J1") = "YearlyChange"

For i = 2 To 753000

If Cells(i, 1).Value = Cells(i + 1, 1).Value And Cells(i, 2).Value = 20180102 Then
start1 = Cells(i, 3).Value
ElseIf Cells(i, 1).Value = Cells(i + 1, 1).Value And Cells(i + 1, 2).Value = 20181231 Then
end1 = Cells(i + 1, 6).Value
Cells(summaryrow, 10) = (end1 - start1)
If Cells(summaryrow, 10).Value > 0 Then
Cells(summaryrow, 10).Interior.ColorIndex = 4
ElseIf Cells(summaryrow, 10).Value < 0 Then
Cells(summaryrow, 10).Interior.ColorIndex = 3
End If
summaryrow = summaryrow + 1
End If
Next i

End Sub

Sub Percentage()

Dim rowcount As Double
Range("K1") = "PercentChange"
rowcount = 2

For i = 2 To 753000
If Cells(i, 2) = 20180102 Then
Cells(rowcount, 11) = FormatPercent((Cells(rowcount, 10).Value / Cells(i, 3).Value), 2)
rowcount = rowcount + 1
End If
Next i

End Sub

Sub stockvolumn()

Dim rownumber As Double
Dim sumvol As Double
Range("L1") = "StockTotalVolume"
rownumber = 2
sumvol = 0
For i = 2 To 753000

If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
sumvol = sumvol + Cells(i, 7).Value
Cells(rownumber, 12) = sumvol

rownumber = rownumber + 1

sumvol = 0

Else
sumvol = sumvol + Cells(i, 7).Value
End If
Next i

End Sub

Sub greatestincrease()
Dim holder As Double
Dim tickersymb As String
holder = 0

Range("N2") = "Greatest % increase"
Range("O1") = "Ticker"
Range("P1") = "Value"

For i = 2 To 753000
If Cells(i, 11).Value > holder Then
holder = Cells(i, 11).Value
tickersymb = Cells(i, 9).Value

End If
Next i
Cells(2, 16) = FormatPercent(holder, 2)
Cells(2, 15) = tickersymb
End Sub

Sub greatestdecrease()
Dim holder2 As Double
Dim tickersymb2 As String
holder2 = 0
Range("N3") = "Greatest % decrease"

For i = 2 To 753000
If Cells(i, 11).Value < holder Then
holder = Cells(i, 11).Value
tickersymb2 = Cells(i, 9).Value
End If
Next i
Cells(3, 16) = FormatPercent(holder, 2)
Cells(3, 15) = tickersymb2
End Sub

Sub greatestvolume()
Dim holder3 As Double
Dim tickersymb3 As String
holder3 = 0
Range("N4") = "Greatest Total Volume"

For i = 2 To 753000
If Cells(i, 12).Value > holder3 Then
holder3 = Cells(i, 12).Value
tickersymb3 = Cells(i, 9).Value
End If
Next i
Cells(4, 15) = tickersymb3
Cells(4, 16) = holder3

End Sub
