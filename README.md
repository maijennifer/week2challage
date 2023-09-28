# week2challage

Sub stock()
    
    [i1] = "Ticker"
    [o1] = "Ticker"
    [p1] = "Value"
    [j1] = "Yearly Change"
    [k1] = "Percent Change"
    [l1] = "Total Stock Volume"
    [n2] = "Greatest % Increase"
    [n3] = "Greatest % Decrease"
    [n4] = "Greatest Total Volume"
    
    Columns("I:P").AutoFit
    Columns("K").NumberFormat = "0.00%"
    
    lastRow = Cells(Rows.Count, "A").End(xlUp).Row
    si = 2
    firstOpen = 0
    TotalVo = 0
    greatestin = 0
    For i = 2 To lastRow
    
        TotalVo = TotalVo + Cells(i, "G")
        
        If firstOpen = 0 Then
            firstOpen = Cells(i, "C")
        End If
            
                        If Cells(i, "A") <> Cells(i + 1, "A") Then
                            Cells(si, "I") = Cells(i, "A")
                            
                            yearlyCh = Cells(i, "F") - firstOpen
                           Cells(si, "J") = yearlyCh
                            
                                                If yearlyCh > 0 Then
                                                    Cells(si, "J").Interior.ColorIndex = 4
                                                Else
                                                    Cells(si, "J").Interior.ColorIndex = 3
                                                End If
                                                
                                            Cells(si, "K") = yearlyCh / firstOpen
                                                
                                            Cells(si, "L") = TotalVo
                                
                            TotalVo = 0
                            firstOpen = 0
                            si = si + 1
                        End If
            
       
    
    Next i
    
End Sub

