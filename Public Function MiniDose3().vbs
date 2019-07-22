Sub MiniDose3()
OptimizeVBA True
    On Error Resume Next
    Dim sngStart As Single
        sngStart = Timer
    Dim zi(), si(), us() As Single
    Dim Row, Col As Integer
    Dim Normal, invert As Boolean
    Dim n, i, j, iteration, m, which As Integer
    Dim p, Gamma, Sigma As Single
    Dim Chains(), sigb() As Single
    Dim UpperSigma, LowerGamma, UpperGamma, XOffset, MaxValueZ, MinValueZ As Single
    Dim burn, thin As Integer
    Dim RowBurn, RowThin As Integer
    Dim Chainsburn(), Chainsthin() As Single
    Dim Sigmab, SigmabError, SigmaT, Sigmabb As Single
    Dim iCutOff As Single
    Dim TempArray() As Single

    'judge log,invert or not
    If (Sheet1.NormalStatisticsButton.Value = True) Then
        Normal = True
        If Sheet11.Max3Button.Value = True Then
            MsgBox ("The Maximum Age Model only supports the log-transformed data, please change the statistics button to 'Log-Normal'!")
            Sheet11.Max3Button.Value = False
            Exit Sub
        End If
    ElseIf (Sheet1.NormalStatisticsButton.Value = False) Then
        Normal = False
        If Sheet11.Max3Button.Value = True Then
            invert = True
        End If
    End If
    '==============================================================
     'copy the data to sheet("Temp")
    n = Sheet1.Cells(22, 29).Value2
    TempArray = GetData(n)
    ReDim zi(1 To n)
    ReDim si(1 To n)
    For i = 1 To n
        If invert = False Then
            zi(i) = TempArray(i, 1)
            ElseIf invert = True Then
            zi(i) = -1 * TempArray(i, 1)
        End If
        si(i) = TempArray(i, 2)
    Next i
    
    XOffset = Abs(Application.WorksheetFunction.Min(zi))
    
    For i = 1 To n
        If invert = True Then
            zi(i) = zi(i) + XOffset
        End If
    Next i
    '==============================================================
    
    ReDim us(1 To n)
    For i = 1 To n
        us(i) = (zi(i) - Application.WorksheetFunction.Average(zi)) ^ 2 / (n - 1)
    Next i
    
    'Set lower and upper boundaries for Gama.
    MinValueZ = Application.WorksheetFunction.Min(zi)
    MaxValueZ = Application.WorksheetFunction.Max(zi)
    
    If MinValueZ > 0 Then
        LowerGamma = MinValueZ * 0.999
    Else
        LowerGamma = MinValueZ * 1.001
    End If
    
    If MaxValueZ > 0 Then
        UpperGamma = MaxValueZ * 1.001
    Else
        UpperGamma = MaxValueZ * 0.999
    End If
    
    
    If Normal = True Then
        UpperSigma = Application.WorksheetFunction.Sum(us)
        Else
        UpperSigma = 10
    End If
    
    Sigmab = Sheet11.Cells(9, 3).Value2
    SigmabError = Sheet11.Cells(9, 4).Value2
    
    If (Sheet11.DefaultCheckButton.Value = True) Then
        p = 0.5
        Gamma = WorksheetFunction.Quartile(zi, 1)
        SigmaT = WorksheetFunction.Quartile(zi, 2)
        If Abs(SigmaT) < UpperSigma Then
            Sigma = Abs(SigmaT)
        Else
            Sigma = UpperSigma * 0.5
        End If
        
        With Sheet11
            .Cells(5, 3).Value2 = p
            .Cells(6, 3).Value2 = Gamma
            .Cells(7, 3).Value2 = "NA"
            .Cells(8, 3).Value2 = Sigma
        End With
        
        ElseIf (Sheet11.DefaultCheckButton.Value = False) Then
            p = Sheet11.Cells(5, 3).Value2
            Gamma = Sheet11.Cells(6, 3).Value2
            Sheet11.Cells(7, 3).Value2 = "NA"
            Sigma = Sheet11.Cells(8, 3).Value2
    End If
    
    
    If p < 0 Or p > 1 Or Gamma < LowerGamma Or Gamma > UpperGamma Or Sigma <= 0 Or Sigma > UpperSigma Then
        MsgBox ("Sorry, you give an illegal initial parameter!")
        Sheet11.MAM3Button.Value = False
        Exit Sub
        Else
        Sheet11.Range("C13:F19").ClearContents
    End If
    
    
    iteration = Sheet11.Cells(5, 9).Value
    burn = Sheet11.Cells(6, 9).Value
    thin = Sheet11.Cells(7, 9).Value
    
    If iteration <= 200 Then
        m = 10
        ElseIf iteration <= 1000 Then
        m = 5
        ElseIf iteration <= 2500 Then
        m = 2
        Else
        m = 1
    End If
    
    '===================================================================
    ReDim Chains(iteration, 3)
    ReDim sigb(1 To iteration)
    For i = 1 To iteration
        Application.StatusBar = "Monte-Carlo Simulation and Slice Sampling......" & Application.WorksheetFunction.MRound(i / iteration * 100, m) & "% Completed"
        
        Sigmabb = Sigmab + Application.WorksheetFunction.Norm_S_Inv(Rnd()) * SigmabError
        If Sigmabb >= 0 Then
        Sigmabb = Sigmabb
        Else
        Sigmabb = 0
        End If
        
        sigb(i) = Sigmabb
        
        'update p
        p = SliceMAM3(n, p, Gamma, Sigma, Sigmabb, 1, 0, 1, zi, si)
        Chains(i, 1) = p
        
        'Update Gamma
        Gamma = SliceMAM3(n, p, Gamma, Sigma, Sigmabb, 2, LowerGamma, UpperGamma, zi, si)
        Chains(i, 2) = Gamma
      
        'Update sigma
        Sigma = SliceMAM3(n, p, Gamma, Sigma, Sigmabb, 3, 0.01, UpperSigma, zi, si)
        Chains(i, 3) = Sigma
        
    Next i
    '=========================================================================
    Application.StatusBar = "Burning and Thining the final results..."
    
    RowBurn = iteration - burn
    RowThin = Round(RowBurn / thin, 0)
    
    ReDim Chainsburn(RowBurn, 3)
    ReDim Chainsthin(RowThin, 3)
    
    'burn-in
    For i = 1 To RowBurn
       For j = 1 To 3
       Chainsburn(i, j) = Chains(i + burn, j)
       Next j
    Next i
    
    'thining
    For i = 1 To RowThin
        For j = 1 To 3
        Chainsthin(i, j) = Chainsburn(thin * (i - 1) + 1, j)
        Next j
    Next i
    
    ReDim prop(1 To RowThin) As Single
    ReDim mamdose(1 To RowThin) As Single
    ReDim sigmanote(1 To RowThin) As Single
    
    For i = 1 To RowThin
        prop(i) = Chainsthin(i, 1)
    If Normal = True Then
        mamdose(i) = Chainsthin(i, 2)
        sigmanote(i) = Chainsthin(i, 3)
    Else
        If invert = True Then
            mamdose(i) = Exp((Chainsthin(i, 2) - XOffset) * -1)
            sigmanote(i) = Chainsthin(i, 3)
        Else
            mamdose(i) = Exp(Chainsthin(i, 2))
            sigmanote(i) = Chainsthin(i, 3)
        End If
    End If
    
    Next i
    
    Dim ACFCR(0 To 30, 1 To 3) As Single
    'Calculate auto - correlation
    For i = 0 To 30
        For which = 1 To 3
            ACFCR(i, which) = ACF(i, which, RowThin, prop, mamdose, sigmanote, prop)
        Next which
    Next i
    
    Application.StatusBar = "Reporting the MCMC results..."
    'calculate 95% CI
    iCutOff = RowThin * 0.025
    With Sheet11
        .Cells(13, 3) = Round(WorksheetFunction.Average(prop), 4)
        .Cells(13, 4) = Round(WorksheetFunction.StDev_S(prop), 4)
        .Cells(13, 5) = Round(WorksheetFunction.Small(prop, iCutOff), 4)
        .Cells(13, 6) = Round(WorksheetFunction.Large(prop, iCutOff), 4)
        .Cells(14, 3) = Round(WorksheetFunction.Average(mamdose), 4)
        .Cells(14, 4) = Round(WorksheetFunction.StDev_S(mamdose), 4)
        .Cells(14, 5) = Round(WorksheetFunction.Small(mamdose, iCutOff), 4)
        .Cells(14, 6) = Round(WorksheetFunction.Large(mamdose, iCutOff), 4)
        .Cells(16, 3) = Round(WorksheetFunction.Average(sigmanote), 4)
        .Cells(16, 4) = Round(WorksheetFunction.StDev_S(sigmanote), 4)
        .Cells(16, 5) = Round(WorksheetFunction.Small(sigmanote, iCutOff), 4)
        .Cells(16, 6) = Round(WorksheetFunction.Large(sigmanote, iCutOff), 4)
        .Cells(17, 3) = Round(WorksheetFunction.Average(sigb), 2) & " " + ChrW(177) + " " & Round(WorksheetFunction.StDev_S(sigb), 2)
        
        If Normal = True Then
            .Cells(18, 3) = Round(LikeliFunc3(n, .Cells(13, 3), .Cells(14, 3), .Cells(16, 3), WorksheetFunction.Average(sigb), zi, si), 4)
        Else
            If invert = True Then
                .Cells(18, 3) = Round(LikeliFunc3(n, .Cells(13, 3), XOffset - Log(.Cells(14, 3)), .Cells(16, 3), WorksheetFunction.Average(sigb), zi, si), 4)
            Else
                .Cells(18, 3) = Round(LikeliFunc3(n, .Cells(13, 3), Log(.Cells(14, 3)), .Cells(16, 3), WorksheetFunction.Average(sigb), zi, si), 4)
            End If
        End If
        .Range("C15:F15").Value = "NA"
        .Cells(19, 3) = CStr(.Cells(14, 3).Value) & " " + ChrW(177) + " " & CStr(Round((.Cells(14, 6).Value - .Cells(14, 5).Value) / 3.92, 3))

    End With
    
    With Sheet11
        .Shapes("SliceGammaPlot").Delete
        .Shapes("SlicePPlot").Delete
        .Shapes("TraceGammaPlot").Delete
        .Shapes("TracePlot").Delete
        .Shapes("ACFPlot").Delete
    End With
    
    With Sheet10
        .Range(Sheet10.Cells(2, 15), Sheet10.Cells(RowThin + 1, 15)).Value2 = Application.Transpose(prop)
        .Range(Sheet10.Cells(2, 16), Sheet10.Cells(RowThin + 1, 16)).Value2 = Application.Transpose(mamdose)
        .Range(Sheet10.Cells(2, 17), Sheet10.Cells(RowThin + 1, 17)).Value2 = Application.Transpose(sigmanote)
        .Range(Sheet10.Cells(2, 21), Sheet10.Cells(32, 23)).Value2 = ACFCR
    End With
    Application.StatusBar = "Finished"
    Sheet11.Cells(19, 8) = "Time used: " & CStr(Round(Timer - sngStart, 3)) & " sec"