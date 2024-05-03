Sub stockStats()
    
    
    Dim ws As Worksheet
    
    For Each ws In ThisWorkbook.Worksheets
        Dim row As Long
        Dim currentStock As String
        ' side table index
        Dim k As Long
        Dim openPrice As Double
        Dim closePrice As Double
        Dim qChange As Double
        Dim runTotal As LongLong
        openPrice = ws.Cells(2, 3).Value
        currentStock = ws.Cells(2, 1).Value
        ws.Cells(2, 9).Value = currentStock
        runTotal = 0
        k = 2
        
        
        Dim lastRow As Long
        lastRow = ActiveSheet.Cells(ActiveSheet.Rows.Count, 1).End(xlUp).row

        
        ' for loop 2 to last row
        For row = 2 To lastRow
            currentStock = ws.Cells(row, 1).Value
            runTotal = runTotal + ws.Cells(row, 7).Value
        ' check to see if row is different
            If (ws.Cells(row + 1, 1).Value <> currentStock) Then
             ' output ticker symbol to column I one per symbol
                ws.Cells(k + 1, 9).Value = ws.Cells(row + 1, 1).Value
                closePrice = ws.Cells(row, 6).Value
                qChange = closePrice - openPrice
                ws.Cells(k, 10).Value = qChange
                If qChange > 0 Then
                    ws.Cells(k, 10).Interior.ColorIndex = 4
                    ws.Cells(k, 11).Interior.ColorIndex = 4
                ElseIf qChange < 0 Then
                    ws.Cells(k, 10).Interior.ColorIndex = 3
                    ws.Cells(k, 11).Interior.ColorIndex = 3
                Else
                    ws.Cells(k, 10).Interior.ColorIndex = 2
                    ws.Cells(k, 11).Interior.ColorIndex = 2
                End If
                
                If openPrice = 0 Then
                    percentChange = 0
                Else
                    percentChange = qChange / openPrice
                End If
                
                ws.Cells(k, 11).NumberFormat = "0.00%"
                ws.Cells(k, 11).Value = percentChange
                openPrice = Cells(row + 1, 3).Value
                ws.Cells(k, 12).Value = runTotal
                runTotal = 0
                k = k + 1
            End If
        Next row
        
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        ' Store greatest % increase info
        Dim greatestInc_name As String
        Dim greatestInc_val As Double
        greatestInc_val = 0
        ' Store greatest % decrease info
        Dim greatestDec_name As String
        Dim greatestDec_val As Double
        greatestDec_val = 0
        ' Store greatest total volume
        Dim greatestVolume_name As String
        Dim greatestVolume_val As LongLong
        greatestVolume_val = 0
        
        Dim i As Long
        For i = 2 To k
            If ws.Cells(i, 11).Value > greatestInc_val Then
                greatestInc_val = ws.Cells(i, 11).Value
                greatestInc_name = ws.Cells(i, 9).Value
            End If
            
            If ws.Cells(i, 11).Value < greatestDec_val Then
                greatestDec_val = ws.Cells(i, 11).Value
                greatestDec_name = ws.Cells(i, 9).Value
            End If
            
            If ws.Cells(i, 12).Value > greatestVolume_val Then
                greatestVolume_val = ws.Cells(i, 12).Value
                greatestVolume_name = ws.Cells(i, 9).Value
            End If
        Next i
        
        ws.Cells(2, 16).Value = greatestInc_name
        ws.Cells(2, 17).Value = greatestInc_val
        ws.Cells(2, 17).NumberFormat = "0.00%"
        ws.Cells(3, 16).Value = greatestDec_name
        ws.Cells(3, 17).Value = greatestDec_val
        ws.Cells(3, 17).NumberFormat = "0.00%"
        ws.Cells(4, 16).Value = greatestVolume_name
        ws.Cells(4, 17).Value = greatestVolume_val
        
        ' Assign colunm headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        
        ws.Columns("I:Q").AutoFit
        
    Next ws
    
End Sub
