Sub StockAnalysis():

    'Declare Variables
    Dim ws As Worksheet
    Dim start As Long
    Dim rowCount As Long
    Dim i As Long
    Dim totalVolume As Double
    Dim j As Integer
    Dim change As Double
    Dim percentChange As Double
    
    'counter variable here?
    
    'Loop through worksheets
    For Each ws In Worksheets


        'Set header row

        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"

        'Initalize Variables
        start = 2
        j = 0
        totalVolume = 0
        change = 0
        
        'get the row number of the last row of data
        rowCount = Cells(Rows.Count, "A").End(xlUp).Row
        
        'Begin loop through worksheet here
        For i = 2 To rowCount
            

            'If cells do not match
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                totalVolume = totalVolume + ws.Cells(i, 7).Value
                
                'zero total volume
                If totalVolume = 0 Then
                    
                    'print results
                    ws.Range("I" & 2 + j).Value = Cells(i, 1).Value
                    ws.Range("J" & 2 + j).Value = 0
                    ws.Range("K" & 2 + j).Value = "%" & 0
                    ws.Range("L" & 2 + j).Value = 0
                
                'Bulk of program
                
                Else
                    ' Find first non zero open price
                    If ws.Cells(start, 3) = 0 Then
                        For findValue = start To i
                            If ws.Cells(findValue, 3).Value <> 0 Then
                                start = findValue
                                Exit For
                            End If
                        Next findValue
                    End If
                    
                    'Calculate yearly change
                    change = (ws.Cells(i, 6) - ws.Cells(start, 3))
                    percentChange = Round((change / ws.Cells(start, 3) * 100), 2)
                    
                    start = i + 1
                
                
                    ws.Range("I" & 2 + j).Value = Cells(i, 1).Value
                    ws.Range("J" & 2 + j).Value = Round(change, 2)
                    ws.Range("K" & 2 + j).Value = "%" & percentChange
                    ws.Range("L" & 2 + j).Value = totalVolume
                    
                    'color code
                    Select Case change
                        Case Is > 0
                            ws.Range("J" & 2 + j).Interior.ColorIndex = 4
                        Case Is < 0
                            ws.Range("J" & 2 + j).Interior.ColorIndex = 3
                        Case Else
                            ws.Range("J" & 2 + j).Interior.ColorIndex = 0
                    End Select
                   
                    
                End If
                   'reset variables, ticker change
                totalVolume = 0
                j = j + 1
                change = 0
                  
                  
            'If ticket is the same, add the volume
            Else
                
                totalVolume = totalVolume + ws.Cells(i, 7).Value


            End If
            
        Next i

    Next ws


End Sub

