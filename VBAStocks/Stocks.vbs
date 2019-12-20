Sub Stocks():
    Dim ticker As String
    Dim total As Double
    Dim totalStockVolume As Integer
    Dim start As Long
    Dim totalChange As Double
    Dim percentChange As Double
    
    
    total = 0
    totalStockVolume = 2
    start = 2
    totalChange = 0
    
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"

    
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To lastRow
        If (Cells(i + 1, 1).Value <> Cells(i, 1).Value) Then
            ticker = Cells(i, 1).Value
            total = total + Cells(i, 7).Value
            Cells(totalStockVolume, 9).Value = ticker
            Cells(totalStockVolume, 12).Value = total
            
            totalChange = Cells(i, 6) - Cells(start, 3)
            percentChange = Round((totalChange / Cells(start, 3) * 100), 2)
            start = i + 1
            Cells(totalStockVolume, 10).Value = totalChange
            Cells(totalStockVolume, 11).Value = percentChange
            
            totalStockVolume = totalStockVolume + 1
            total = 0
            totalChange = 0
        Else
            total = total + Cells(i, 7).Value
        End If
    Next i
       
    
End Sub

