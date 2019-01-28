Attribute VB_Name = "Module2"
Sub Stockmagic()
Dim lastrow As Long
Dim rowcount As Double
Dim stock As Double
Dim summaryrow As Double
Dim begyear As Double
Dim endyear As Double
Dim greatest As Double
Dim least As Double
Dim greatcontender As Double
Dim leastcontenter As Double
Dim stockcontender As Double
Dim totalstock As Double
Dim ws As Worksheet

For Each ws In Worksheets
rowcount = Range("A2", Range("A2").End(xlDown)).Rows.Count + 1


stock = 0

greatest = 0
least = 700000
totalstock = 0

' create the ticker and value columns
ws.Range("Q1").Value = "Ticker"
ws.Range("R1").Value = "Value"
ws.Cells(2, 16).Value = "Greatest % Increase"
ws.Cells(3, 16).Value = "Greatest % decrease"
ws.Cells(4, 16).Value = "Greatest Total Volume"
' autofit
ws.Columns("P:P").EntireColumn.AutoFit

ws.Range("J1").Value = "Ticker"
ws.Range("K1").Value = "Yearly Change"
ws.Range("L1").Value = "Percentage Change"
ws.Range("M1").Value = "Total Volume"

ws.Range("A2", ws.Range("A3").End(xlDown)).AdvancedFilter xlFilterCopy, , ws.Range("J2"), True

ws.Range("J2", ws.Range("J2").End(xlDown)).RemoveDuplicates Columns:=1, Header:=xlNo


summaryrow = 2

For i = 2 To rowcount


If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Or IsEmpty(Cells(i + 1, 1).Value) Then

endyear = ws.Cells(i, 6).Value
' then we must do the math beg year - end year
ws.Cells(summaryrow, 11).Value = endyear - begyear
If begyear = 0 Then
ws.Cells(summaryrow, 12).Value = 0
Else
ws.Cells(summaryrow, 12).Value = (endyear - begyear) / begyear
End If
stock = stock + Cells(i, 7).Value
ws.Cells(summaryrow, 13).Value = stock
If ws.Cells(summaryrow, 11).Value < 0 Then
ws.Cells(summaryrow, 11).Interior.ColorIndex = 3
Else: ws.Cells(summaryrow, 11).Interior.ColorIndex = 4
End If
' insert hard here since we fill out our total stock, percent increase, and percent change here
' set contenders to current percentage changes
' dont forget summary row not i when trying to find those vlaues
greatcontender = ws.Cells(summaryrow, 12).Value
leastcontender = ws.Cells(summaryrow, 12).Value
stockcontender = ws.Cells(summaryrow, 13).Value
' if contender > greatest value then set greatest value to equal contender and zero out contender
If greatcontender > greatest Then
greatest = greatcontender

' set up greatest value and stock name into the appropriate cells
ws.Cells(2, 17).Value = ws.Cells(summaryrow, 10).Value
ws.Cells(2, 18).Value = greatest
' ditto for least but contender < least
ElseIf leastcontender < least Then
least = leastcontender
ws.Cells(3, 17).Value = Cells(summaryrow, 10).Value
ws.Cells(3, 18).Value = least
Else

End If


' ditto for stock volume
'
If stockcontender > totalstock Then
totalstock = stockcontender
ws.Cells(4, 17).Value = ws.Cells(summaryrow, 10).Value
ws.Cells(4, 18).Value = totalstock
Else
End If
' increment summary row by 1
summaryrow = summaryrow + 1
' reset the values
stock = 0
begyear = 0
endyear = 0

ElseIf ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
begyear = ws.Cells(i, 3).Value
' keep that first value
stock = stock + ws.Cells(i, 7).Value
Else
stock = stock + ws.Cells(i, 7).Value
End If




Next i



ws.Columns("L:L").Style = "Percent"
ws.Range("R2:R3").Style = "Percent"

  ws.Columns("K:K").EntireColumn.AutoFit
    ws.Columns("L:L").EntireColumn.AutoFit
    ws.Columns("M:M").EntireColumn.AutoFit
    ws.Columns("P:P").EntireColumn.AutoFit
    ws.Columns("Q:Q").EntireColumn.AutoFit
    ws.Columns("R:R").EntireColumn.AutoFit
Next ws

End Sub


