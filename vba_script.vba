Sub StockDataCode()
Dim ws As Worksheet
Dim myTicker As String
Dim open1 As Double
Dim volume As Double
Dim close1 As Double
Dim yearlyChange As Double
Dim percentChange As Double
Dim startrow As Double
Dim endrow As Double
Dim lrow As Double



'Loop through each worksheet
For Each ws In Worksheets
ws.Activate

'Finds the last row that has data in it in column A
lrow = Cells(Rows.Count, "A").End(xlUp).Row

'Land headers for the data we will output with the code
Cells(1, "I").Value = "Ticker"
Cells(1, "J").Value = "Yearly Change"
Cells(1, "K").Value = "Percent Change"
Cells(1, "L").Value = "Volume"

'Identify the first instance of the ticker

For i = 2 To lrow
If Cells(i, 1).Value <> Cells(i - 1, 1).Value Then
'if it is the first instance get startrow, ticker and open1
open1 = Format(Cells(i, 3).Value, "#.00")
startrow = Cells(i, 1).Row
myTicker = Cells(i, 1).Value

Cells(Rows.Count, "I").End(xlUp).Offset(1, 0).Activate
ActiveCell.Value = myTicker

End If

'Identify the last instance of the ticker
If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
close1 = Format(Cells(i, 6).Value, "#.00")
endrow = Cells(i, 1).Row

'Sum up the volume for the current ticker
volume = Application.WorksheetFunction.Sum(Range(Cells(startrow, "G"), Cells(endrow, "G")))
'Get yearly change for current ticker
yearlyChange = close1 - open1
'Get percent change for current ticker
percentChange = (close1 - open1) / open1

'Land values in the first empty under headers
Cells(Rows.Count, "J").End(xlUp).Offset(1, 0).Activate
ActiveCell.Value = Format(yearlyChange, "#.00")
Cells(Rows.Count, "K").End(xlUp).Offset(1, 0).Activate
ActiveCell.Value = Format(percentChange, "0.00%")
Cells(Rows.Count, "L").End(xlUp).Offset(1, 0).Activate
ActiveCell.Value = volume

End If

Next i

'Autofit columns to be sure everything fits in the columns
Range("I:L").EntireColumn.AutoFit

'Color code based on yearly change value: Positive=Green, Negative=Red
'Find last row in column I for the loop stop point
lrow2 = Cells(Rows.Count, "I").End(xlUp).Row

For j = 2 To lrow2
If Cells(j, "J").Value > 0 Then
Cells(j, "J").Interior.Color = vbGreen
Else
Cells(j, "J").Interior.Color = vbRed
End If

Next j


'Land headers for % decrease, % increase etc.....
Range("N2").Value = "Greatest % Increase"
Range("N3").Value = "Greatest % Decrease"
Range("N4").Value = "Greatest Total Volume"

Range("O1").Value = "Ticker"
Range("P1").Value = "Value"

Range("P2").Value = Format(Application.WorksheetFunction.Max(Range(Cells(2, "K"), Cells(lrow2, "K"))), "0.00%")
Range("P3").Value = Format(Application.WorksheetFunction.Min(Range(Cells(2, "K"), Cells(lrow2, "K"))), "0.00%")
Range("P4").Value = Application.WorksheetFunction.Max(Range(Cells(2, "L"), Cells(lrow2, "L")))

'Find which ticker goes with each value in P2:P4
topIncrease = Format(Range("P2").Value, "0.00%")
Range("K:K").Find(topIncrease).Activate
ticker = ActiveCell.Offset(0, -2).Value
Range("O2").Value = ticker

topDecrease = Format(Range("P3").Value, "0.00%")
Range("K:K").Find(topDecrease).Activate
ticker = ActiveCell.Offset(0, -2).Value
Range("O3").Value = ticker

topVolume = Range("P4").Value
Range("L:L").Find(topVolume).Activate
ticker = ActiveCell.Offset(0, -3).Value
Range("O4").Value = ticker

Range("N:P").EntireColumn.AutoFit

Next ws


End Sub


