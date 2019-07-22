Public Function LikeliCAM(ByVal n As Double, ByVal OverDispersionTerm As Double, ByVal MeanValue As Double, zi(), se()) As Double
On Error Resume Next
    Dim i As Integer
    Dim Weight(), llik() As Double
    Dim SD() As Double
    ReDim Weight(1 To n)
    ReDim llik(1 To n)
    
    For i = 1 To n
    If se(i) <> 0 Then
        Weight(i) = 1 / (se(i) ^ 2 + OverDispersionTerm ^ 2)
    End If
        llik(i) = Log(Weight(i)) - Weight(i) * (zi(i) - MeanValue) ^ 2
    Next i
    LikeliCAM = 0.5 * Application.WorksheetFunction.Sum(llik)
End Function