Sub HWKSW2()
'Lopp over mulitple workshet
For Each ws In Worksheets
    ws.Cells(1, 9).Value = "Tricker"
    ws.Cells(1, 16).Value = "Tricker"
    ws.Cells(1, 10).Value = "Total Stock Volume"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = " Greates Total Volume"
'Declare Variables Calculate yearly total, percent change, and total volume by ticker
Dim I As Long
Dim TickerName As String
Dim openYearly As Double
Dim totalVolume As Double
totalVolume = 0
Dim tickerRow As Long
tickerRow = 2
Dim lastRow As Long
lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
'Add loop
For I = 2 To lastRow
openYarly = ws.Cells(tickerRow, 3).Value
'add conditional
    If ws.Cells(I + 1, 1).Value <> ws.Cells(I, 1).Value Then
        TickerName = ws.Cells(I, 1).Value
        ws.Range("I" & tickerRow).Value = TickerName
           
               
        totalVolume = totalVolume + ws.Cells(I, 7).Value
        ws.Range("J" & tickerRow).Value = totalVolume
        'Reset
        tickerRow = tickerRow + 1
        totalYearly = 0
        totalVolume = 0
        openYearly = ws.Cells(tickerRow, 3).Value
            Else
        totalVolume = totalVolume + ws.Cells(I, 7).Value
     End If
Next I
Next ws

End Sub
