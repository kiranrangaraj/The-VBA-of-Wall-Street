Attribute VB_Name = "Module1"
Sub alphabet_testing()

    Dim stockVolume As Double
    Dim i As Long
    Dim yearlyChange As Single
    Dim j As Integer
    Dim initialStockValue As Long
    Dim rowCount As Long
    Dim percentChange As Single
    Dim days As Integer
    Dim dailyChange As Single
    Dim averageChange As Single
    Dim ws As Worksheet
    
    For Each ws In Worksheets
        j = 0
        stockVolume = 0
        Change = 0
        initialStockValue = 2
        dailyChange = 0
   
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
    
        rowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
        For i = 2 To rowCount
    
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                stockVolume = stockVolume + ws.Cells(i, 7).Value
            
                If stockVolume = 0 Then
                    ws.Range("I" & 2 + j).Value = Cells(i, 1).Value
                    ws.Range("J" & 2 + j).Value = 0
                    ws.Range("K" & 2 + j).Value = "%" & 0
                    ws.Range("L" & 2 + j).Value = 0
                
                Else
                    If ws.Cells(initialStockValue, 3).Value = 0 Then
                        For findValue = initialStockValue To i
                            If ws.Cells(findValue, 3).Value <> 0 Then
                                initialStockValue = findValue
                                Exit For
                            End If
                        Next findValue
                    End If
                
                    yearlyChange = (ws.Cells(i, 6).Value - ws.Cells(initialStockValue, 3).Value)
                    percentChange = yearlyChange / ws.Cells(initialStockValue, 3).Value * 100
            
                    initialStockValue = i + 1
            
                    ws.Range("I" & 2 + j).Value = Cells(i, 1).Value
                    ws.Range("J" & 2 + j).Value = Round(yearlyChange, 2)
                    ws.Range("K" & 2 + j).Value = "%" & Round(percentChange, 2)
                    ws.Range("L" & 2 + j).Value = stockVolume
            
                    If ws.Range("J" & 2 + j).Value > 0 Then
                        ws.Range("J" & 2 + j).Interior.ColorIndex = 4
                    ElseIf ws.Range("J" & 2 + j).Value < 0 Then
                        ws.Range("J" & 2 + j).Interior.ColorIndex = 3
                    Else
                        ws.Range("J" & 2 + j).Interior.ColorIndex = 0
                
                    End If
            
                End If
        
                stockVolume = 0
                yearlyChange = 0
                j = j + 1
                days = 0
                dailyChange = 0
        
            Else
                stockVolume = stockVolume + ws.Cells(i, 7).Value
        
            End If
        
        Next i
    
        ws.Range("P2") = ws.Cells(WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & rowCount)), ws.Range("K2:K" & rowCount), 0) + 1, 9)
        ws.Range("P3") = ws.Cells(WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & rowCount)), ws.Range("K2:K" & rowCount), 0) + 1, 9)
        ws.Range("P4") = ws.Cells(WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & rowCount)), ws.Range("L2:L" & rowCount), 0) + 1, 9)
        
        ws.Range("Q2") = "%" & WorksheetFunction.Max(ws.Range("K2:K" & rowCount)) * 100
        ws.Range("Q3") = "%" & WorksheetFunction.Min(ws.Range("K2:K" & rowCount)) * 100
        ws.Range("Q4") = WorksheetFunction.Max(ws.Range("L2:L" & rowCount))
           
    Next ws

End Sub
