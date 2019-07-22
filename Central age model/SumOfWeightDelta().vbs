Public Function SumOfWeightDelta(ByVal n As Integer, ByVal OD As Double, zi(), se()) As Double
    If n > 0 Then
        Dim Weight(), weightdelta() As Double
        ReDim Weight(1 To n)
        ReDim weightdelta(1 To n)
        Dim MaxRowNumber As Integer
        MaxRowNumber = Application.WorksheetFunction.Max(Sheet1.Range("A3:A10003")) + 2
        For i = 1 To n
            If se(i) <> 0 Then
            Weight(i) = 1 / (se(i) ^ 2 + OD ^ 2)
            Else
            Weight(i) = 0
            End If
            weightdelta(i) = zi(i) * Weight(i)
        Next i
            SumOfWeightDelta = Application.WorksheetFunction.Sum(weightdelta)
    ElseIf n = 0 Then
        SumOfWeightDelta = Application.WorksheetFunction.Sum(Sheet1.Range(Sheet1.Range(Sheet1.Cells(3, 2), Sheet1.Cells(MaxRowNumber, 2))))
    End If
End Function
