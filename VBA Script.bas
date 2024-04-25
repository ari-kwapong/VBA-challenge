Attribute VB_Name = "Module1"
Sub stocks()
    Dim ws As Worksheet
    
    For Each ws In Worksheets
    
        ws.[I1] = "Ticker"
        ws.[J1] = "Yearly Change"
        ws.[K1] = "Percent Change"
        ws.[L1] = "Total Stock Volume"
        ws.[O2] = "Greatest % Increase"
        ws.[O3] = "Greatest % Decrease"
        ws.[O4] = "Greatest Total Volume"
        ws.[P1] = "Ticker"
        ws.[Q1] = "Value"
        
        Columns.AutoFit
        
        si = 2
        firstOpen = 0
        totalVol = 0
        greatestInc = 0
        greatestDec = 0
        
        lastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row
    
            ws.Activate
            
        For i = 2 To lastRow
            totalVol = totalVol + ws.Cells(i, "G")
            
            'This happens at the first row of the ticker
            If firstOpen = 0 Then
                firstOpen = ws.Cells(i, "C")
            End If
            
            'This only happens at the last ticker's row
            If ws.Cells(i, "A") <> ws.Cells(i + 1, "A") Then
                ws.Cells(si, "I") = ws.Cells(i, "A")
                
                yearlyCh = ws.Cells(i, "F") - firstOpen
                ws.Cells(si, "J") = yearlyCh
                
                If yearlyCh > 0 Then
                    ws.Cells(si, "J").Interior.ColorIndex = 4
                Else
                    ws.Cells(si, "J").Interior.ColorIndex = 3
                End If
                
                ws.Cells(si, "L") = totalVol
                
                pctCh = (yearlyCh / firstOpen) * 100
                ws.Cells(si, "K") = pctCh
                
                'Greatest Increase
                If pctCh > greatestInc Then
                    greatestInc = pctCh
                    ws.Range("P2") = ws.Cells(si, "I")
                    ws.Range("Q2") = greatestInc
                End If
                
                'Greatest Decrease
                If pctCh < greatestDec Then
                    greatestDec = pctCh
                    ws.Range("P3") = ws.Cells(si, "I")
                    ws.Range("Q3") = greatestDec
                End If
                
                'reset area
                si = si + 1
                firstOpen = 0
                totalVol = 0
            End If
        
        Next i
    Next ws
End Sub



