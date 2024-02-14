Sub stock_m2()

    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        ws.Activate
        
        Dim i As Long
        Dim rowcount As Long
        Dim start As Long
        Dim change As Double
        Dim percentChange As Double
        Dim closePrice As Double
        Dim openPrice As Double
        Dim totalVolume As Double
        Dim j As Integer
        Dim GI As Double
        Dim GD As Double
        Dim GV As Double
        Dim GITicker As String
        Dim GDTicker As String
        Dim GVTicker As String
           
        ' Variables
        start = 2
        j = 0
        GI = 0
        GD = 0
        GV = 0
        
        ' Title rows
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Year Change"
        Range("K1").Value = "Percent Change"
        Range("L1").Value = "Total Stock Volume"
     
        ' Get the row number of the last cell with data
        rowcount = Cells(Rows.Count, "A").End(xlUp).Row
        
        ' Loop through each row
        For i = 2 To rowcount
            ' Acquire total volume for the current ticker
            totalVolume = totalVolume + Cells(i, 7).Value
            
            ' If ticker changes or if it's the last row
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                ' Set the closing price for the current ticker
                closePrice = Cells(i, 6).Value
                ' Get opening price for the ticker
                openPrice = Cells(start, 3).Value
                ' Calculate change from beginning to end
                change = closePrice - openPrice
                
                'Work out percent change
                If openPrice <> 0 Then
                    percentChange = (change / openPrice)
                Else
                    percentChange = 0
                End If
                
                ' Calculate on the greatest increase and decrease
                If percentChange > GI Then
                    GI = percentChange
                    GITicker = Cells(i, 1).Value
                ElseIf percentChange < GD Then
                    GD = percentChange
                    GDTicker = Cells(i, 1).Value
                End If
                
                ' Calculate on the greatest volume
                If totalVolume > GV Then
                    GV = totalVolume
                    GVTicker = Cells(i, 1).Value
                End If
                
                ' Output for ticker, year change, percent change, and total volume
                Range("I" & 2 + j).Value = Cells(i, 1).Value
                Range("J" & 2 + j).Value = change
                Range("J" & 2 + j).NumberFormat = "0.00"
                Range("K" & 2 + j).Value = percentChange
                Range("K" & 2 + j).NumberFormat = "0.00" & "%"
                Range("L" & 2 + j).Value = totalVolume
                
               ' Conditional formatting for percent change
                Select Case True
                    Case percentChange > 0
                        Range("K" & 2 + j).Interior.ColorIndex = 4
                    Case percentChange < 0
                        Range("K" & 2 + j).Interior.ColorIndex = 3
                    Case Else
                        Range("K" & 2 + j).Interior.ColorIndex = 0
                End Select

              ' Conditional formatting for year change
                Select Case True
                    Case change > 0
                        Range("J" & 2 + j).Interior.ColorIndex = 4
                    Case change < 0
                        Range("J" & 2 + j).Interior.ColorIndex = 3
                    Case Else
                        Range("J" & 2 + j).Interior.ColorIndex = 0
                    End Select

                ' Move on to the next ticker
                start = i + 1
                ' Move on to the next row for the next output
                j = j + 1
                ' Reset total volume for next ticker
                totalVolume = 0
                
            End If
        Next i
        
        ' After loop, output greatest increase, decrease, and volume
        Range("P2").Value = GITicker
        Range("Q2").Value = GI
        Range("P3").Value = GDTicker
        Range("Q3").Value = GD
        Range("P4").Value = GVTicker
        Range("Q4").Value = GV
        
        ' Additional titles
        Range("O2").Value = "Greatest % Increase"
        Range("O3").Value = "Greatest % Decrease"
        Range("O4").Value = "Greatest Total Volume"
        Range("P1").Value = "Ticker"
        Range("Q1").Value = "Value"
        
        ' Formating greatest % decrease and increase
        Range("Q3").NumberFormat = "0.00" & "%"
        Range("Q2").NumberFormat = "0.00" & "%"
        
    Next ws
    
End Sub

    
