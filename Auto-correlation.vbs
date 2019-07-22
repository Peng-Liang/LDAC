Public Function ACF(ByVal lag As Integer, ByVal which As Integer, ByVal RowThin As Integer, ByRef prop() As Single, ByRef mamdose() As Single, ByRef sigmanote() As Single, ByRef mudose() As Single) As Single
    Dim n As Integer
    Dim S0, sk, MeanOfData As Single
    Dim dif1(), dif2(), Diff() As Single
    Dim i As Integer
    n = RowThin
    ReDim Para(0 To n) As Single

    If which = 1 Then
        Para = prop
    ElseIf which = 2 Then
        Para = mamdose
    ElseIf which = 3 Then
        Para = sigmanote
    ElseIf which = 4 Then
        Para = mudose
    End If
    
    MeanOfData = Application.WorksheetFunction.Average(Para)
    S0 = Application.WorksheetFunction.Var_P(Para)
    ReDim dif1(1 To n)
    ReDim dif2(1 To n)
    ReDim Diff(1 To n)
    For i = 1 To n - lag
        dif1(i) = Para(i) - MeanOfData
        dif2(i) = Para(i + lag) - MeanOfData
        Diff(i) = dif1(i) * dif2(i)
    Next i
    sk = 1 / n * Application.WorksheetFunction.Sum(Diff)
    ACF = sk / S0
End Function
