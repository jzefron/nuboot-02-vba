Attribute VB_Name = "Module1"
Sub gather()
'loop through all open workbooks
For Each wb In Workbooks
    'look for the stock data
    If wb.Name Like "*_stock_data.xlsx" Or wb.Name = "alphabtical_testing.xlsx" Then

   For Each ws In wb.Worksheets
        With ws 'don't like typing wb all the time
        
        'wb.Activate
        
        

        Dim tableCounter As Integer
        Dim lastLine As Long
        Dim yOpen, yClose, yVolume As Double
        dateCounter = 2
        tableCounter = 2
        .Range("I1:L1") = Array("Ticker", "Year change", "percent change", "total volume")
        .Range("I2", .Range("L2").End(xlDown)).ClearContents
        
        lastLine = .Range("A2").End(xlDown).Row
        yOpen = .Cells(2, 3).Value
        For r = 2 To lastLine
            If .Cells(r, 1).Value <> .Cells(r + 1, 1).Value Then
                .Cells(tableCounter, 9).Value = .Cells(r, 1).Value
                yClose = .Cells(r, 6).Value
                .Cells(tableCounter, 10).Value = yClose - yOpen
                If yClose >= yOpen Then
                    .Cells(tableCounter, 10).Interior.ColorIndex = 4 ' green positive gain
                Else
                    .Cells(tableCounter, 10).Interior.ColorIndex = 3 ' red negative gain
                End If
                If yOpen <> 0 Then
                    .Cells(tableCounter, 11).Value = .Cells(tableCounter, 10).Value / yOpen
                    .Cells(tableCounter, 11).NumberFormat = "0.00%"
                End If
                ' this probably is faster way of summing up
                 .Cells(tableCounter, 12).Value = Application.Sum(Range(.Cells(dateCounter, 7), .Cells(r, 7)))
        
                yOpen = .Cells(r + 1, 3).Value
                tableCounter = tableCounter + 1
                dateCounter = r + 1
            End If
           
        Next r
        .Range("N1") = "Max positive change"
        .Range("n2") = "Max loss"
        .Range("n3") = "Greatest value"
        Dim maxGain, maxLoss As Double, maxVol As Double, gainT, lossT, volT As String
        
        maxGain = -9999#
        maxLoss = 9999#
        
        maxVol = 0
        For i = 2 To tableCounter
            If .Cells(i, 11).Value > maxGain Then
               maxGain = .Cells(i, 11).Value
               gainT = .Cells(i, 9).Value
            End If
            If .Cells(i, 11).Value < maxLoss Then
               maxLoss = .Cells(i, 11).Value
               lossT = .Cells(i, 9).Value
            End If
            If .Cells(i, 12).Value > maxVol Then
                maxVol = .Cells(i, 12).Value
                volT = .Cells(i, 9).Value
            End If
        Next i
       .Range("o1") = gainT
        .Range("o2") = lossT
        .Range("o3") = volT
        
        .Range("p1") = maxGain
        .Range("p2") = maxLoss
        .Range("p1:p2").NumberFormat = "0.00%"
        .Range("p3") = maxVol
    End With
    Next    'new work sheet
    End If ' end stock data
    
Next ' next work book
End Sub
