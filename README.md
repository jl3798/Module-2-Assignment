Sub stock_analysis()
    'Set dimensions
    Dim total As Double
    Dim rowindex As Long
    Dim change As Double
    Dim columnindex As Integer
    Dim start As Long
    Dim rowcount As Long
    Dim percentchange As Double
    Dim days As Integer
    Dim dailychange As Single
    Dim averagechange As Double
    Dim ws As Worksheet
    
    For Each ws In Worksheets
        columnindex = 0
        total = 0
        change = 0
        start = 2
        dailychange = 0
        
        'set title row
        
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O1").Value = "Greatest % increase"
        ws.Range("O2").Value = "Greatest % decrease"
        ws.Range("O3").Value = "Greatest % total volume"
        
        'get the row number of the last row with data
        rowcount = Cells(Rows.Count, "A").End(xlUp).Row
        
        For rowindex = 2 To rowcount
        
            'if ticker changes then print results
            If ws.Cells(rowindex + 1, 1).Value <> ws.Cells(rowindex, 1).Value Then
                
                'store results in variable
                total = total + ws.Cells(rowindex, 7).Value
                
                If total = 0 Then
                    'print the results
                    ws.Range("I" & 2 + columnindex).Value = Cells(rowindex, 1).Value
                    ws.Range("J" & 2 + columnindex).Value = 0
                    ws.Range("K" & 2 + columnindex).Value = "%" + 0
                    ws.Range("L" & 2 + columnindex).Value = 0
                Else
                    If ws.Cells(start, 3) = 0 Then
                        For find_values = start To rowindex
                            If ws.Cells(find_value, 3).Value <> 0 Then
                                start = find_value
                                Exit For
                        End If
                    Next
                 End If
            
            change = (ws.Cells(rowindex, 6) - ws.Cells(start, 3))
            percentchange = change / ws.Cells(start, 3)
            
            start = rowindex + 1
            
            ws.Range("I" & 2 + columnindex) = ws.Cells(rowindex, 1).Value
            ws.Range("J" & 2 + columnindex) = change
            ws.Range("J" & 2 + columnindex).NumberFormat = "0.00"
            ws.Range("K" & 2 + columnindex).Value = percentchange
            ws.Range("K" & 2 + columnindex).Numberformatt = "0.00%"
            ws.Range("L" & 2 + columnindex).Value = total
            
            Select Case change
                Case Is > 0
                    ws.Range("J" & 2 + columnindex).Interior.ColorIndex = 4
                Case Is < 0
                    ws.Range("J" & 2 + columnindex).Interior.ColorIndex = 3
                Case Else
                    ws.Range("J" & 2 + columnindex).Interior.ColorIndex = 0
            End Select
                
                
            End If
            
            total = 0
            change = 0
            columnindex = columnindex + 1
            days = 0
            daliychange = 0
            
        Else
            'if ticker is still the same add results
            total = total + ws.Cells(rwoindex, 7).Value
            
        End If
        
    Next rowindex
     
    'take the max and min and place them in a seperate part in the workesheet
    ws.Range("Q2") = "%" & WorksheetFunction.Max(ws.Range("K2:K" & rowcount)) * 100
    ws.Range("Q3") = "%" & WorksheetFunction.Min(ws.Range("K2:K" & rowcount)) * 100
    ws.Range("Q4") = WorksheetFunction.Max(ws.Range("L2:L" & rowcount))
    
    increase_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & rowcount)).ws.Range("K2:K" & rowcount), 0)
    decrease_number = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & rowcount)).ws.Range("K2:K" & rowcount), 0)
    volume_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & rowcount)).ws.Range("L2:L" & rowcount), 0)
    
    ws.Range("P2") = ws.Cells(increase_number + 1, 9)
    ws.Range("P3") = ws.Cells(decrease_number + 1, 9)
    ws.Range("P4") = ws.Cells(volume_number + 1, 9)
                
        
        
        
        Next ws
        
    
End Sub
