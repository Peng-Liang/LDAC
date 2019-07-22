Sub Homotest(ByVal n As Integer, zi(), se())
Dim i As Integer
Dim G As Single
Dim Weight, Prob As Single

    If n > 1 Then
        MeanValue = SumOfWeightDelta(n, 0, zi, se) / SumOfWeight(n, 0, zi, se)
        G = 0
        DF = n - 1
        For i = 1 To n
            Weight = 1 / (se(i) ^ 2)
            G = G + Weight * (zi(i) - MeanValue) ^ 2
        Next i
        Prob = Application.WorksheetFunction.ChiDist(G, DF)
        Sheet1.Cells(24, 29) = "P ( " + ChrW(967) + "2 ) = " + CStr(Format(Prob, "###0.0000")) + " | G = " + CStr(Round(G, 2))
    End If
End Sub