Sub StockData():

    For Each ws In Worksheets
    
'    declare variables i will use
        Dim WorksheetName As String
        Dim i As Long
        Dim j As Long
        Dim ticker As Long
        Dim lastRow1 As Long
        Dim lastRow2 As Long
        Dim percentChange As Double
        Dim greatestIncr As Double
        Dim greatestDecr As Double
        Dim greatestVol As Double
   
'   assign variables their values
        WorksheetName = ws.Name
        ticker = 2
        j = 2
        lastRow1 = ws.Cells(Rows.Count, 1).End(xlUp).Row
        lastRow2 = ws.Cells(Rows.Count, 9).End(xlUp).Row
        GreatVol = ws.Cells(2, 12).Value
        GreatIncr = ws.Cells(2, 11).Value
        GreatDecr = ws.Cells(2, 11).Value
        
        
'   name header cells
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
    
'  goes from first row at row 2 and scans for each value, prints to IJKL columns and color changes
    For i = 2 To lastRow1
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                    ws.Cells(ticker, 9).Value = ws.Cells(i, 1).Value
                    ws.Cells(ticker, 10).Value = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value
                
                If ws.Cells(ticker, 10).Value < 0 Then
                    ws.Cells(ticker, 10).Interior.ColorIndex = 3
                Else
                    ws.Cells(ticker, 10).Interior.ColorIndex = 4
                End If
                    
                If ws.Cells(j, 3).Value <> 0 Then
                    percentChange = ((ws.Cells(i, 6).Value - ws.Cells(j, 3).Value) / ws.Cells(j, 3).Value)
                    ws.Cells(ticker, 11).Value = Format(percentChange, "Percent")
                Else
                    ws.Cells(ticker, 11).Value = Format(0, "Percent")
                End If
                    
                ws.Cells(ticker, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(j, 7), ws.Cells(i, 7)))
                ticker = ticker + 1
                j = i + 1
                
                End If
            
            Next i
            
' goes from first row to last rowand scans for greatest changes and prints to OPQ
    For i = 2 To lastRow2
            If ws.Cells(i, 12).Value > greatestVol Then
                    greatestVol = ws.Cells(i, 12).Value
                    ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
                Else
                    greatestVol = greatestVol
                End If
        
                If ws.Cells(i, 11).Value > greatestIncr Then
                    greatestIncr = ws.Cells(i, 11).Value
                    ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
                Else
                    greatestIncr = greatestIncr
                End If
                
                If ws.Cells(i, 11).Value < greatestDecr Then
                    greatestDecr = ws.Cells(i, 11).Value
                    ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
                Else
                    greatestDecr = greatestDecr
                End If
                
                ws.Cells(2, 17).Value = Format(greatestIncr, "Percent")
                ws.Cells(3, 17).Value = Format(greatestDecr, "Percent")
                ws.Cells(4, 17).Value = Format(greatrestVol, "Scientific")
            
            Next i
            
    Next ws
        
End Sub