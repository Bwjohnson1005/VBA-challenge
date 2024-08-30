Attribute VB_Name = "Module1"
Sub QuarterlyStockAnalysisAllSheets_GlobalSummary()

    Dim ws As Worksheet
    Dim ticker As String
    Dim lastRow As Long
    Dim i As Long
    Dim quarterlyOpen As Double
    Dim quarterlyClose As Double
    Dim quarterlyVolume As Double
    Dim resultRow As Long
    Dim percentageChange As Double
    
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim greatestVolume As Double
    Dim tickerGreatestIncrease As String
    Dim tickerGreatestDecrease As String
    Dim tickerGreatestVolume As String
    
    greatestIncrease = -99999
    greatestDecrease = 99999
    greatestVolume = 0
    
    For Each ws In ThisWorkbook.Worksheets
        
        
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

        
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "Percentage Change"
        ws.Cells(1, 12).Value = "Total Volume"

        resultRow = 2

        
        ticker = ws.Cells(2, 1).Value
        quarterlyOpen = ws.Cells(2, 3).Value
        quarterlyVolume = 0

        For i = 2 To lastRow

            If ws.Cells(i, 1).Value <> ticker Then
                
                percentageChange = (quarterlyClose - quarterlyOpen) / quarterlyOpen * 100
                percentageChange = Round(percentageChange, 2)
                
                ws.Cells(resultRow, 9).Value = ticker
                ws.Cells(resultRow, 10).Value = quarterlyClose - quarterlyOpen
                ws.Cells(resultRow, 11).Value = percentageChange
                ws.Cells(resultRow, 12).Value = quarterlyVolume
                resultRow = resultRow + 1
                
                If percentageChange > greatestIncrease Then
                    greatestIncrease = percentageChange
                    tickerGreatestIncrease = ticker
                End If
                If percentageChange < greatestDecrease Then
                    greatestDecrease = percentageChange
                    tickerGreatestDecrease = ticker
                End If
                If quarterlyVolume > greatestVolume Then
                    greatestVolume = quarterlyVolume
                    tickerGreatestVolume = ticker
                End If

                ticker = ws.Cells(i, 1).Value
                quarterlyOpen = ws.Cells(i, 3).Value
                quarterlyVolume = 0
            End If
            
            quarterlyClose = ws.Cells(i, 6).Value
            quarterlyVolume = quarterlyVolume + ws.Cells(i, 7).Value
        Next i
        
        percentageChange = (quarterlyClose - quarterlyOpen) / quarterlyOpen * 100
        percentageChange = Round(percentageChange, 2)
        
        ws.Cells(resultRow, 9).Value = ticker
        ws.Cells(resultRow, 10).Value = quarterlyClose - quarterlyOpen
        ws.Cells(resultRow, 11).Value = percentageChange
        ws.Cells(resultRow, 12).Value = quarterlyVolume
        
        If percentageChange > greatestIncrease Then
            greatestIncrease = percentageChange
            tickerGreatestIncrease = ticker
        End If
        If percentageChange < greatestDecrease Then
            greatestDecrease = percentageChange
            tickerGreatestDecrease = ticker
        End If
        If quarterlyVolume > greatestVolume Then
            greatestVolume = quarterlyVolume
            tickerGreatestVolume = ticker
        End If
        
        
        With ws
            .Cells(1, 15).Value = "Summary"
            .Cells(2, 14).Value = "Greatest % Increase:"
            .Cells(2, 15).Value = tickerGreatestIncrease
            .Cells(2, 16).Value = greatestIncrease & "%"
            
            .Cells(3, 14).Value = "Greatest % Decrease:"
            .Cells(3, 15).Value = tickerGreatestDecrease
            .Cells(3, 16).Value = greatestDecrease & "%"
            
            .Cells(4, 14).Value = "Greatest Total Volume:"
            .Cells(4, 15).Value = tickerGreatestVolume
            .Cells(4, 16).Value = greatestVolume
        End With
        
    Next ws

    MsgBox "Quarterly stock analysis is complete!"

End Sub
