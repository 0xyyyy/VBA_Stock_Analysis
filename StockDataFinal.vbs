Sub StockOpen():
Dim Ticker As String
Dim lastrow As Double
Dim Output_row As Double
Dim VolumeTotal As LongLong
VolumeTotal = 0
Dim YearOp As Double
Dim YearCl As Double

For Each ws In Worksheets
    ws.Activate
    ws.Range("H1") = "Ticker"
    ws.Range("I1") = "Yearly Change"
    ws.Range("J1") = "Percent Change"
    ws.Range("K1") = "Total Volume Traded"
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    Output_row = 2
        For i = 2 To lastrow
        YearOp = Cells(2, 6)
            If ws.Cells(i, 1) <> ws.Cells(i + 1, 1) Then
                If YearOp <> 0 Then
                    YearCl = ws.Cells(i, 6)
                    YearlyChange = YearCl - YearOp
                    PercChange = (YearCl - YearOp) / (YearOp)
                    Range("I" & Output_row) = YearlyChange
                    Range("J" & Output_row) = PercChange
                    Range("J" & Output_row).Style = "Percent"
                        If Range("I" & Output_row) > 0 Then
                            Range("I" & Output_row).Interior.ColorIndex = 4
                            Range("J" & Output_row).Interior.ColorIndex = 4
                        Else
                            Range("I" & Output_row).Interior.ColorIndex = 3
                            Range("J" & Output_row).Interior.ColorIndex = 3
                        End If
                End If
                YearOp = ws.Cells(i + 1, 6)
                Ticker = Cells(i, 1).Value
                VolumeTotal = VolumeTotal + Cells(i, 7).Value
                Range("H" & Output_row) = Ticker
                Range("K" & Output_row) = VolumeTotal
                Output_row = Output_row + 1
                VolumeTotal = 0
            Else
                VolumeTotal = VolumeTotal + Cells(i, 7).Value
            End If
        Next i
Next ws
End Sub

