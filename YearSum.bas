Attribute VB_Name = "Module1"
Sub gather()
Dim tableCounter As Integer
Dim lastLine As Long
Dim yOpen, yClose, yVolume As Double
dateCounter = 2
tableCounter = 2
Range("I1:L1") = Array("Ticker", "Year change", "percent change", "total volume")
Range("I2", Range("L2").End(xlDown)).ClearContents

lastLine = Range("A2").End(xlDown).Row
yOpen = Cells(2, 3).Value
For r = 2 To lastLine
    If Cells(r, 1).Value <> Cells(r + 1, 1).Value Then
        Cells(tableCounter, 9).Value = Cells(r, 1).Value
        yClose = Cells(r, 6).Value
        Cells(tableCounter, 10).Value = yClose - yOpen
        Cells(tableCounter, 11).Value = Cells(tableCounter, 10).Value / yOpen
        Cells(tableCounter, 12).Value = Application.Sum(Range(Cells(dateCounter, 7), Cells(r, 7)))
     '   Range(Cells(tableCounter, 11)).NumberFormat = "Percent"
        yOpen = Cells(r + 1, 3).Value
        tableCounter = tableCounter + 1
        dateCounter = r + 1
    End If
   
Next r




End Sub
