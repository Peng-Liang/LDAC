Sub LogProfile(ByVal n As Integer, ByVal OverDispersionTerm As Double, ByVal MeanValue As Double, ByVal DistributionIsLogNormal As Double, zi(), se())
    OptimizeVBA True
    Dim Lmax As Single
    Dim RelativeOD As Single
    Dim Step As Single
    Dim LoopCounter, LoopCounter1, LoopCounter2, LoopCounter3, LoopCounter4, FuHao As Integer
    Dim ODdown, ODup, LowerCI, UpperCI, OD As Double
    Dim Value, StDevOfOD As Single
    Dim Profile(1 To 200, 1 To 2) As Single
    
    Lmax = LikeliCAM(n, OverDispersionTerm, MeanValue, zi, se)
    Sheet10.Cells(2, 9) = Lmax
    If DistributionIsLogNormal = True Then
        RelativeOD = OverDispersionTerm
        Else
        RelativeOD = OverDispersionTerm / MeanValue
    End If
    
    If Round(RelativeOD, 3) <> 0 Then
        If n >= 10 Then
            If RelativeOD >= 0.03 Then
            Step = OverDispersionTerm * 0.05
            Else
            Step = OverDispersionTerm * 5
            End If
        Else
            If RelativeOD >= 0.03 Then
            Step = OverDispersionTerm * 0.5
            Else
            Step = OverDispersionTerm
            End If
        End If
        LoopCounter1 = 0
        LoopCounter2 = 0
        ODdown = OverDispersionTerm
        ODup = OverDispersionTerm
        
        'down search
        Do
        LoopCounter1 = LoopCounter1 + 1
        ODdown = ODdown - Step
        Value = SumOfWeightDelta(n, ODdown, zi, se) / SumOfWeight(n, ODdown, zi, se)
        Loop Until (LikeliCAM(n, ODdown, Value, zi, se) - Lmax <= -5 Or LoopCounter1 >= 200 Or ODdown <= 0)
        
        'up search
        Do
        LoopCounter2 = LoopCounter2 + 1
        ODup = ODup + Step
        Value = SumOfWeightDelta(n, ODup, zi, se) / SumOfWeight(n, ODup, zi, se)
        Loop Until (LikeliCAM(n, ODup, Value, zi, se) - Lmax <= -5 Or LoopCounter2 >= 200)
    Else
        ODdown = 0
        Do
        LoopCounter = LoopCounter + 1
            If DistributionIsLogNormal Then
            ODup = ODup + 0.2
            Else
            ODup = ODup + 20
            End If
        Value = SumOfWeightDelta(n, ODup, zi, se) / SumOfWeight(n, ODup, zi, se)
        Loop Until (LikeliCAM(n, ODup, Value, zi, se) - Lmax <= -100 Or LoopCounter > 50)
    End If
    
    If ODdown <= 0 Then
        ODdown = 0
    End If
    
    For i = 1 To 200
        OD = ODdown + (ODup - ODdown) / 199 * (i - 1)
        Value = SumOfWeightDelta(n, OD, zi, se) / SumOfWeight(n, OD, zi, se)
        Profile(i, 1) = OD
        Profile(i, 2) = LikeliCAM(n, OD, Value, zi, se) - Lmax
    Next i
    
    If Round(RelativeOD, 3) <> 0 Then
        LowerCI = OverDispersionTerm
        UpperCI = OverDispersionTerm
    
        'finner lower
        Tol = OverDispersionTerm
        LoopCounter1 = 0
        Do
            LoopCounter1 = LoopCounter1 + 1
            FuHao = Right(LoopCounter1, 1) Mod 2
            Tol = Tol * 0.1
            If FuHao = 0 Then
                LoopCounter2 = 0
                Do
                LoopCounter2 = LoopCounter2 + 1
                LowerCI = LowerCI + Tol
                MeanValue = SumOfWeightDelta(n, LowerCI, zi, se) / SumOfWeight(n, LowerCI, zi, se)
                Loop Until (LikeliCAM(n, LowerCI, MeanValue, zi, se) - Lmax >= -1.92 Or LowerCI <= 0 Or LoopCounter2 >= 1000)
            Else
                LoopCounter2 = 0
                Do
                LoopCounter2 = LoopCounter2 + 1
                LowerCI = LowerCI - Tol
                MeanValue = SumOfWeightDelta(n, LowerCI, zi, se) / SumOfWeight(n, LowerCI, zi, se)
                Loop Until (LikeliCAM(n, LowerCI, MeanValue, zi, se) - Lmax <= -1.92 Or LowerCI <= 0 Or LoopCounter2 >= 1000)
            End If
        Loop Until (Round(LikeliCAM(n, LowerCI, MeanValue, zi, se) - Lmax, 5) = -1.92 Or LoopCounter1 >= 50 Or LowerCI <= 0)
        If LowerCI <= 0 Then
        LowerCI = 0
        End If
        
        'finner upper
        Tol = OverDispersionTerm
        LoopCounter3 = 0
        Do
            LoopCounter3 = LoopCounter3 + 1
            Tol = Tol * 0.1
            FuHao = Right(LoopCounter3, 1) Mod 2
            If FuHao = 0 Then
                LoopCounter4 = 0
                Do
                LoopCounter4 = LoopCounter4 + 1
                UpperCI = UpperCI - Tol
                MeanValue = SumOfWeightDelta(n, UpperCI, zi, se) / SumOfWeight(n, UpperCI, zi, se)
                Loop Until (LikeliCAM(n, UpperCI, MeanValue, zi, se) - Lmax >= -1.92 Or LoopCounter4 >= 1000)
            Else
                LoopCounter4 = 0
                Do
                LoopCounter4 = LoopCounter4 + 1
                UpperCI = UpperCI + Tol
                MeanValue = SumOfWeightDelta(n, UpperCI, zi, se) / SumOfWeight(n, UpperCI, zi, se)
                Loop Until (LikeliCAM(n, UpperCI, MeanValue, zi, se) - Lmax <= -1.92 Or LoopCounter4 >= 1000)
            End If
        Loop Until (Round(LikeliCAM(n, UpperCI, MeanValue, zi, se) - Lmax, 5) = -1.92 Or LoopCounter3 >= 50)
     
        StDevOfOD = (UpperCI - LowerCI) / 3.92
        Sheet10.Cells(7, 43) = LowerCI
        Sheet10.Cells(7, 44) = LikeliCAM(n, LowerCI, SumOfWeightDelta(n, LowerCI, zi, se) / SumOfWeight(n, LowerCI, zi, se), zi, se) - Lmax
        Sheet10.Cells(8, 43) = UpperCI
        Sheet10.Cells(8, 44) = LikeliCAM(n, UpperCI, SumOfWeightDelta(n, UpperCI, zi, se) / SumOfWeight(n, UpperCI, zi, se), zi, se) - Lmax
        Sheet10.Cells(9, 43) = StDevOfOD
    End If
    Call DrawLikProfile(OverDispersionTerm, DistributionIsLogNormal, MeanValue, RelativeOD, Profile)
End Sub