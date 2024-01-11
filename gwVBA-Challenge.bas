Attribute VB_Name = "Module1"
Sub TickerTally()

     'Declare Variables
    Dim ticker As String
    Dim openPrice As Double
    Dim closePrice As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    Dim analysisRow As Long
    Dim ws As Worksheet
    Dim colorRange As Range
    Dim percentRange As Range
    Dim maxVolume As Double
    Dim maxTicker As String
    
    maxVolume = 0
    
    
    'Direct to workbook and sheets within
    For Each ws In ThisWorkbook.Sheets(Array("2018", "2019", "2020"))
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        
        'Label headers for analysis
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        'Row to begin after headers
        analysisRow = 2
        
        'Added for Color formatting in Yearly Change summary
        Set colorRange = ws.Range(ws.Cells(analysisRow, 10), ws.Cells(lastRow, 10))
        'Added for Percentage Format in Percent Change summary
        Set percentRange = ws.Range(ws.Cells(analysisRow, 11), ws.Cells(lastRow, 11))
        
        
        For i = 2 To lastRow
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            ticker = ws.Cells(i, 1).Value
            closePrice = ws.Cells(i, 6).Value
                yearlyChange = closePrice - openPrice
                If openPrice <> 0 Then
                    percentChange = yearlyChange / openPrice
                Else
                    percentChange = 0
                End If
                
                totalVolume = totalVolume + ws.Cells(i, 7).Value
                'Summary values placement
                ws.Cells(analysisRow, 9).Value = ticker
                ws.Cells(analysisRow, 10).Value = yearlyChange
                ws.Cells(analysisRow, 11).Value = percentChange
                ws.Cells(analysisRow, 12).Value = totalVolume
                
                If yearlyChange > 0 Then
                    'If negative value color red
                    colorRange(analysisRow - 1).Interior.ColorIndex = 4
                ElseIf yearlyChange < 0 Then
                    'If positive value color green
                    colorRange(analysisRow - 1).Interior.ColorIndex = 3
                ElseIf yearlyChange = 0 Then
                    'If zero value color light blue
                    colorRange(analysisRow - 1).Interior.ColorIndex = 20
                
                End If
                
                'Formats Percent Change summary range to percentage
                percentRange(analysisRow - 1).NumberFormat = "0.00%"
                
                analysisRow = analysisRow + 1
                
                ' Reset values for the next ticker
                openPrice = 0
                closePrice = 0
                yearlyChange = 0
                percentChange = 0
                totalVolume = 0
            Else
                ' Set open price for the current ticker
                If openPrice = 0 Then
                    openPrice = ws.Cells(i, 3).Value ' Open price in column C
                End If
                ' Accumulate total volume
                totalVolume = totalVolume + ws.Cells(i, 7).Value
                
                'If totalVolume > maxVolume Then
                    'maxVolume = totalVolume
                    'maxTicker = ticker
                'End If
                
            End If
                Cells(1, 16).Value = "Ticker"
                Cells(1, 17).Value = "Value"
                Cells(2, 15).Value = "Greatest % Increase"
                Cells(3, 15).Value = "Greatest % Decrease"
                Cells(4, 15).Value = "Greatest Total Volume"
                
                Cells(4, 16).Value = maxTicker
                Cells(4, 17).Value = maxVolume
            
        Next i
            
                
            
    Next ws
    
End Sub
