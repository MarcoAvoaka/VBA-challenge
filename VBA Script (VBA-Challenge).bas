Attribute VB_Name = "Module1"



Sub yearstock()

'Declaring the variable

Dim i As Long
Dim ticker As String
Dim YearlyChange As Double
Dim PercentChange As Double
Dim TotalStockvolume As Double
Dim TickerRow As Integer
Dim closing As Double
Dim opening As Double
Dim Current As Worksheet
For Each Current In Worksheets

Current.Cells(1, 9).Value = "Ticker"
Current.Cells(1, 10).Value = "Yearly change"
Current.Cells(1, 11).Value = "Percent change"
Current.Cells(1, 12).Value = "Total stock volume"
TickerRow = 1
TotalStockvolume = 0
openingprice = Current.Cells(2, 3).Value
closingprice = 0



'Loop through rows in the column

Maxrows = ActiveSheet.UsedRange.Rows.Count


For i = 2 To Maxrows

'conditions
If Current.Cells(i + 1, 1).Value <> Current.Cells(i, 1).Value Then
ticker = Current.Cells(i, 1).Value

TickerRow = TickerRow + 1
Current.Cells(TickerRow, 9).Value = ticker
Current.Cells(TickerRow, 12).Value = TotalStockvolume
closingprice = Current.Cells(i, 6).Value
YearlyChange = closingprice - openingprice
Current.Cells(TickerRow, 10).Value = YearlyChange
'Handling division by zero error
If openingprice = 0 Then
Current.Cells(TickerRow, 11).Value = Str(0) + "%"
Else
Current.Cells(TickerRow, 11).Value = Str(Round((YearlyChange / openingprice) * 100, 2)) + "%"
End If
'Resetting variables for next ticker

openingprice = Current.Cells(i + 1, 3).Value
TotalStockvolume = 0
If (YearlyChange > 0) Then
Current.Cells(TickerRow, 10).Interior.ColorIndex = 4
Else
Current.Cells(TickerRow, 10).Interior.ColorIndex = 3
End If

Else
TotalStockvolume = TotalStockvolume + Current.Cells(i, 7).Value
End If

Next i
Next
End Sub

