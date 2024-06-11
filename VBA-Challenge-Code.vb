Sub vbaChallenge()
    'Dim worksheet for loop
    Dim ws As Worksheet
    For Each ws In Worksheets
        'Identify the given values and what we're trying to find
        Dim tName As String
        Dim qChange As Double
        Dim pChange As Double
        Dim tVol As Double
        Dim LastRowA As Long
        LastRowA = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        Dim LastRowB As Long
        LastRowB = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
        'Dim what the tools we will use to do our calculation
        Dim i As Long
        Dim k As Long
        k = 2
        Dim oPrice As Double
        oPrice = ws.Cells(2, 3).Value
        Dim cPrice As Double
        'Dim for Summary Tabble
        Dim gVolume As Double
        Dim gInc As Double
        Dim gDec As Double
        'Create Header for Output Table
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        'Start of nestedLoop1 for Output Table
        'Ensure to use the LastRow Forumla
        For i = 2 To LastRowA
            'Make sure to combine all the Tickernames together
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                tName = ws.Cells(i, 1).Value
                tVol = tVol + ws.Cells(i, 7).Value
                cPrice = ws.Cells(i, 6).Value
                'Calc the quarter change = closePrice - openPrice
                qChange = cPrice - oPrice
                'Assign value to populate in cells
                ws.Cells(k, 10).Value = qChange
                ws.Cells(k, 9).Value = tName
                'Start of nested Loop2
                'Assign color index to show the quarterly changes that are <0 (Red) or >0 (Green)
                If qChange < 0 Then
                    ws.Cells(k, 10).Interior.ColorIndex = 3 'Red
                Else
                    ws.Cells(k, 10).Interior.ColorIndex = 4 'Green
                End If
                'Format cells into Percet
                If oPrice <> 0 Then
                    pChange = qChange / oPrice
                    ws.Cells(k, 11).Value = Format(pChange, "Percent")
                Else
                '0 in this case means false
                    ws.Cells(k, 11).Value = Format(0, "Percent")
                End If
                'End of nestedLoop2
                tVol = 0
                k = k + 1
            Else
                tVol = tVol + ws.Cells(i, 7).Value
                ws.Cells(k, 12).Value = tVol
            End If
            'End of nestedLoop1 for Output Table
        Next i

        'Summary Table Code
        'Assign Header cells
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        'Assign where to get values from
        gVolume = ws.Cells(2, 12).Value
        gInc = ws.Cells(2, 11).Value
        gDec = ws.Cells(2, 11).Value

        For i = 2 To LastRowB
            'Keep looping to check through the column and compare until they find the largest increase
            If ws.Cells(i, 11).Value > gInc Then
                gInc = ws.Cells(i, 11).Value
                ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
                ws.Cells(2, 17).Value = Format(gInc, "Percent")
            End If

            'Keep looping to check through the column and compare until they find the largest decrease
            If ws.Cells(i, 11).Value < gDec Then
                gDec = ws.Cells(i, 11).Value
                ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
                ws.Cells(3, 17).Value = Format(gDec, "Percent")
            End If

            'Keep looping to check through the column and compare until they find the largest volume
            If ws.Cells(i, 12).Value > gVolume Then
                gVolume = ws.Cells(i, 12).Value
                ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
                ws.Cells(4, 17).Value = Format(gVolume, "Scientific")
            End If
        Next i
    Next ws
End Sub
