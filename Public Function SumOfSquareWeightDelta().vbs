Public Function SumOfSquareWeightDelta(ByVal n As Integer, ByVal OD As Double, ByVal MeanValue As Double, zi(), se()) As Double
    Dim Weight(), SquareWeight(), sandarddeviationofweightdelta() As Double
    Dim i As Integer
    ReDim Weight(1 To n)
    ReDim SquareWeight(1 To n)
    ReDim sandarddeviationofweightdelta(1 To n)
    For i = 1 To n
        If se(i) <> 0 Then
            Weight(i) = 1 / (se(i) ^ 2 + OD ^ 2)
        Else
            Weight(i) = 0
        End If
        SquareWeight(i) = Weight(i) ^ 2
        sandarddeviationofweightdelta(i) = Weight(i) ^ 2 * (zi(i) - MeanValue) ^ 2
    Next i
    SumOfSquareWeightDelta = Application.WorksheetFunction.Sum(sandarddeviationofweightdelta) / Application.WorksheetFunction.Sum(Weight)
End Function