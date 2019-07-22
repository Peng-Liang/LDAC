Public Function SumOfWeight(ByVal n As Integer, ByVal OD As Double, zi(), se()) As Double
    If n > 0 Then
        Dim Weight() As Double
        Dim nTotal, i As Integer
        ReDim Weight(1 To n)
        nTotal = ValidXValues
        For i = 1 To n
            If se(i) <> 0 Then
            Weight(i) = 1 / (se(i) ^ 2 + OD ^ 2)
            Else
            Weight(i) = 0
            End If
        Next i
        SumOfWeight = Application.WorksheetFunction.Sum(Weight)
    ElseIf n = 0 Then
        If nTotal > 0 Then
            SumOfWeight = nTotal
        Else
            SumOfWeight = 1
        End If
    End If
End Function