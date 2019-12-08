Sub Stocks():
    Dim ticker As String
    Dim total As Double
    Dim totalStockVolume As Integer
    total = 0
    totalStockVolume = 2
    For i = 2 To 80000
        If (Cells(i + 1, 1).Value <> Cells(i, 1).Value) Then
            ticker = Cells(i, 1).Value
            total = total + Cells(i, 3).Value
            Cells(totalStockVolume, 9).Value = ticker
            Cells(totalStockVolume, 10).Value = total
            totalStockVolume = totalStockVolume + 1
            total = 0
        Else
            total = total + Cells(i, 3).Value
        End If
    Next i
End Sub