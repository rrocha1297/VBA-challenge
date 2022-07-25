Attribute VB_Name = "Module1"
Sub stockTotal()
    Dim last As Long
    Dim running As Long
    Dim ticker As String
    Dim openStock As Double
    Dim closeStock As Double
    Dim volumeStock As Double
    Dim maxInc As Double
    Dim maxDec As Double
    Dim maxVol As Double
        
    For Each ws In Worksheets
    
        last = ws.Cells(Rows.Count, "A").End(xlUp).Row

        running = 2
        maxInc = 0
        maxDec = 0
        maxVol = 0
    
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"

        openStock = ws.Cells(2, 3).Value
    
        For i = 2 To last
            ticker = ws.Cells(i, 1).Value
            volumeStock = volumeStock + ws.Cells(i, 7).Value
        
            If ws.Cells(i + 1, 1) <> ticker Then
                closeStock = ws.Cells(i, 6).Value
                ws.Cells(running, 9).Value = ticker
                ws.Cells(running, 10).Value = closeStock - openStock
                
                If openStock = 0 Then
                    If closeStock = 0 Then
                        ws.Cells(running, 11).Value = 0
                    Else
                        ws.Cells(running, 11).Value = closeStock / Abs(closeStock)
                    End If
                Else
                    ws.Cells(running, 11).Value = closeStock / openStock - 1
                End If

                ws.Cells(running, 12).Value = volumeStock
            
                If ws.Cells(running, 12).Value > maxVol Then
                    maxVol = ws.Cells(running, 12).Value
                    maxVolTck = ticker
                End If
                
                If ws.Cells(running, 11).Value > maxInc Then
                    maxInc = ws.Cells(running, 11).Value
                    maxIncTck = ticker
                    
                ElseIf ws.Cells(running, 11).Value < maxDec Then
                    maxDec = ws.Cells(running, 11).Value
                    maxDecTck = ticker
                End If
            
                ws.Cells(running, 11).NumberFormat = "0.00%"
                ws.Cells(running, 12).NumberFormat = "#,###"
            
                If ws.Cells(running, 10) < 0 Then
                    ws.Cells(running, 10).Interior.ColorIndex = 3
                Else
                    ws.Cells(running, 10).Interior.ColorIndex = 4
                End If
                
                running = running + 1
                openStock = ws.Cells(i + 1, 3).Value
                closeStock = 0
                volumeStock = 0
        
            End If
        
        Next i

    Next ws
    
End Sub
