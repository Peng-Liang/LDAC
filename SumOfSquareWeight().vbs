Public Function SumOfSquareWeight(ByVal n As Integer, ByVal OD As Double, zi(), se()) As Double
    If n > 0 Then
        Dim i As Integer
        ReDim SquareWeight(1 To n) As Single
        For i = 1 To n
            If se(i) <> 0 Then
            SquareWeight(i) = (1 / (se(i) ^ 2 + OD ^ 2)) ^ 2
            End If
        Next i
        SumOfSquareWeight = Application.WorksheetFunction.Sum(SquareWeight)
    End If
End Function
