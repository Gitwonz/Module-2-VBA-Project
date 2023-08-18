# Module-2-VBA-Project
VBA stock market project

Create a script that loops through all the stocks for one year and outputs the following information:
The ticker symbol.
Yearly change from the opening price at the beginning of a given year to the closing price at the end of that year.
The percentage change from the opening price at the beginning of a given year to the closing price at the end of that year.
The total stock volume of the stock.


...vba
Sub stock_analysis():
    Dim i As Long
    Dim rowindex As Long
    Dim total As Double
    Dim change As Double
    Dim rowCount As Long
    Dim percentchange As Double
    Dim dailychange As Single
    Dim cellnumber As Integer
    Dim ws As Worksheet
    
    For Each ws In Worksheets
        total = 0
        cellnumber = 2
        Start = 2
        change = 0
        dailychange = 0
        
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        rowCount = ws.Cells(Rows.Count, "A").End(xlUp).Row
        
        For rowindex = 2 To rowCount
            If ws.Cells(rowindex + 1, 1).Value <> ws.Cells(rowindex, 1).Value Then
                total = total + ws.Cells(rowindex, 7).Value
                tick = ws.Cells(rowindex, 1).Value
                If ws.Cells(Start, 3).Value = 0 Then
                    percentchange = 0
                Else
                    change = ws.Cells(rowindex, 6).Value - ws.Cells(Start, 3).Value
                    percentchange = change / ws.Cells(Start, 3).Value
                End If
                
                ws.Range("I" & cellnumber).Value = tick
                ws.Range("J" & cellnumber).Value = change
                ws.Range("J" & cellnumber).NumberFormat = "0.00"
                ws.Range("K" & cellnumber).Value = percentchange
                ws.Range("K" & cellnumber).NumberFormat = "0.00%"
                ws.Range("L" & cellnumber).Value = total
                
                cellnumber = cellnumber + 1
                Start = rowindex + 1
                total = 0
            Else
                total = total + ws.Cells(rowindex, 7).Value
                
            End If
        Next rowindex

        rowNumber = ws.Cells(Rows.Count, "J").End(xlUp).Row
        For i = 2 To rowNumber
            If ws.Cells(i, 10).Value > 0 And ws.Cells(i, 11).Value > 0 Then
                ws.Cells(i, 10).Interior.ColorIndex = 4
                ws.Cells(i, 11).Interior.ColorIndex = 4
            Else
                ws.Cells(i, 10).Interior.ColorIndex = 3
                ws.Cells(i, 11).Interior.ColorIndex = 3
            End If
        Next i
        
        
        ws.Range("Q2") = "%" & WorksheetFunction.Max(ws.Range("K2:K" & rowCount)) * 100
        ws.Range("Q3") = "%" & WorksheetFunction.Min(ws.Range("K2:K" & rowCount)) * 100
        ws.Range("Q4") = WorksheetFunction.Max(ws.Range("L2:L" & rowCount))
        
        increasednum = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & rowCount)), ws.Range("K2:K" & rowCount), 0)
        decreasednum = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & rowCount)), ws.Range("K2:K" & rowCount), 0)
        volumenum = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & rowCount)), ws.Range("L2:L" & rowCount), 0)
        
        ws.Range("P2") = ws.Cells(increasednum + 1, 9)
        ws.Range("P3") = ws.Cells(decreasednum + 1, 9)
        ws.Range("P4") = ws.Cells(volumenum + 1, 9)
        
    Next ws
    
        
End Sub

...
