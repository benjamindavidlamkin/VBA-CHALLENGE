Attribute VB_Name = "Module1"

Sub multiYearStocks():
'Looping through multiple sheets
For Each ws In Worksheets
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    
'Declaring Variables-using the Long

Dim i As Long
Dim tickerName As String
Dim openYearly As Double
Dim totalVolume As Double
totalVolume = 0
Dim totalYearly As Double
totalYearly = 0
Dim percentChange As Double
Dim tickerRow As Long
tickerRow = 2
Dim lastRow As Long
lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Adding a loop
For i = 2 To lastRow
openYearly = ws.Cells(tickerRow, 3).Value

'Conditional for finding values
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        tickerName = ws.Cells(i, 1).Value
        ws.Range("I" & tickerRow).Value = tickerName
    
        totalYearly = totalYearly + (ws.Cells(i, 6).Value - openYearly)
        ws.Range("J" & tickerRow).Value = totalYearly
    
        percentChange = (totalYearly / openYearly)
        ws.Range("K" & tickerRow).Value = percentChange
        ws.Range("K" & tickerRow).Style = "Percent"
        
        totalVolume = totalVolume + ws.Cells(i, 7).Value
        ws.Range("L" & tickerRow).Value = totalVolume
        
        'Reset
        tickerRow = tickerRow + 1
        totalYearly = 0
        totalVolume = 0
        openYearly = ws.Cells(tickerRow, 3).Value
    Else
        totalVolume = totalVolume + ws.Cells(i, 7).Value
    End If
Next i


    
Next ws

End Sub
