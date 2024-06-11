Sub moduleTwoChallenge():
    
    'Loop for each and every worksheet
    Dim ws As Worksheet
    For Each ws In Worksheets
    
    'Set all Dimensions and add values to variables
    'Assign variable to lastrow
        Dim tickerName As String
        Dim percentChange As Double
        Dim LastRowA As Double
            LastRowA = ws.Cells(Rows.Count, 1).End(xlUp).Row
        Dim LastRowB As Double
            LastRowB = ws.Cells(Rows.Count, 9).End(xlUp).Row
        Dim gIncrease As Double
        Dim gDecrease As Double
        Dim gVolume As Double
        Dim i As Double
        Dim k As Double
        Dim n As Double
        Dim oPrice As Double
        Dim cPrice As Double
        
        
        'To refer to all worksheets in workbook
        WorksheetName = ws.Name
    
        
        'Start of Data Extraction and Analysis
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        k = 2
        n = 2
        
            'Look up to see how to find and assign last row of a column
            'Start of Overall Loop
            For i = 2 To LastRowA
            oPrice = ws.Cells(k, 3).Value
            cPrice = ws.Cells(i, 6).Value
            
            'Apply Ticker Name to appropriate column
            ws.Cells(i, 9).Value = ws.Cells(i, 1).Value
            
                'Start of Nested loop1
                If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                'Quarterly Change = Close - Open
                ws.Cells(i, 10).Value = cPrice - oPrice
                
                'Insert Nested loop for color assignment to Quarterly Change
                    'Start of Nested loop2
                    If ws.Cells(i, 10).Value < 0 Then
                    'Red
                    ws.Cells(2, 10).Interior.ColorIndex = 3
                
                    Else
                    'Green
                    ws.Cells(2, 10).Interior.ColorIndex = 4
                
                    End If
                    'Make sure open doesn't = 0 because you cant divide by 0 (Will come back as an error!)
                    If ws.Cells(k, 3).Value <> 0 Then
                    'percentChange= ((earliest value - later value)/earliest value)*100
                    percentChange = ((oPrice - cPrice) / oPrice)
                    
                    'Look up how to change formating into "Percent" to negate having to *100 in formula
                    
                    ws.Cells(2, 11).Value = Format(percentChange, "Percent")
                    
                    Else
                    '0 means false in this case
                    ws.Cells(2, 11).Value = Format(0, "Percent")
                    'End of nested Loop2
                    End If
                    
                'Total Volume =
                ws.Cells(2, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(k, 7), ws.Cells(i, 7)))
                
                qChange = n + 1
                
                k = i + 1
                'End of nested loop1
                End If
            'End of overall loop
            Next i
        
        'Start of Summary Table Code
        'Assign headers for Summary Table
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        
        'Set Values to Cells
        gVolume = ws.Cells(2, 12).Value
        gIncrease = ws.Cells(2, 11).Value
        gDecrease = ws.Cells(2, 11).Value
        Ticker = ws.Cells(1, 16).Value
        Value = ws.Cells(1, 17).Value
        gIncName = ws.Cells(2, 16).Value
        gDecName = ws.Cells(3, 16).Value
        gVolName = ws.Cells(4, 16).Value
        
            For i = 2 To LastRowB
            
            
                If ws.Cells(i, 11).Value > gIncrease Then
                gIncrease = ws.Cells(i, 11).Value
                'Remember to have the tName to follow with the data
                gIncName = ws.Cells(i, 9).Value
                
                Else
                
                gIncrease = gIncrease
                
                End If
                
                If ws.Cells(i, 11).Value < gDecrease Then
                gDecrease = ws.Cells(i, 11).Value
                'tName to follow the data
                wDecName = ws.Cells(i, 9).Value
                
                Else
                
                gDecrease = gDecrease
                
                End If
                
                If ws.Cells(i, 12).Value > gVolume Then
                gVolume = ws.Cells(i, 12).Value
                't name to follow the data
                gVolName = ws.Cells(i, 9).Value
                
                Else
                
                gVolume = gVolume
                
                End If
                
                
            'Format Value
            ws.Cells(4, 17).Value = Format(gVolume, "Scientific")
            ws.Cells(2, 17).Value = Format(gIncrease, "Percent")
            ws.Cells(3, 17).Value = Format(gDecrease, "Percent")
            
            Next i
    Next ws
        
End Sub
