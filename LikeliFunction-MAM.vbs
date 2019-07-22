Public Function LikeliFunc3(ByVal n As Integer, ByVal p As Single, ByVal Gamma As Single, ByVal Sigma As Single, ByVal Sigmab As Single, zi(), si()) As Double
On Error GoTo Errorhandler
    Dim s2(), sigma0(), mu0() As Single
    Dim f1() As Double
    Dim f2() As Double
    Dim f() As Double
    Dim likelihoodfunction() As Double
    Dim res1() As Double
    Dim i As Integer
    ReDim sigma0(1 To n)
    ReDim mu0(1 To n)
    ReDim f1(1 To n)
    ReDim f2(1 To n)
    ReDim f(1 To n)
    ReDim likelihoodfunction(1 To n)
    ReDim res1(1 To n)
    ReDim s2(1 To n)
    ReDim se(1 To n) As Single
    
    For i = 1 To n
        se(i) = Sqr(si(i) ^ 2 + Sigmab ^ 2)
    Next i
    
    For i = 1 To n
        s2(i) = Sigma ^ 2 + se(i) ^ 2
        sigma0(i) = 1 / Sqr(1 / Sigma ^ 2 + 1 / se(i) ^ 2)
        mu0(i) = (Gamma / Sigma ^ 2 + zi(i) / se(i) ^ 2) / (1 / Sigma ^ 2 + 1 / se(i) ^ 2)
        res1(i) = (Gamma - mu0(i)) / sigma0(i)
        f1(i) = 1 / (Sqr(2 * Pi * se(i) ^ 2)) * Exp(-(zi(i) - Gamma) ^ 2 / (2 * se(i) ^ 2))
        f2(i) = 1 / (Sqr(2 * Pi * s2(i))) * ((1 - Application.WorksheetFunction.Norm_S_Dist(res1(i), True)) / 0.5) * Exp(-(zi(i) - Gamma) ^ 2 / (2 * s2(i)))
        f(i) = p * f1(i) + (1 - p) * f2(i)
        If f(i) > 0 Then
            likelihoodfunction(i) = Log(f(i))
        Else
            likelihoodfunction(i) = 0
        End If
    Next i
    
    LikeliFunc3 = Application.WorksheetFunction.Sum(likelihoodfunction)
    Exit Function
    
Errorhandler:
    If MsgBox("Sorry, you get an error due to unfit data or wrong initial parameters, do you want to try it again", vbYesNo, "Warning") = vbYes Then
        Else
        Sheet11.MAM3Button.Value = False
        Sheet11.Max3Button.Value = False
        End
    End If
End Function

Public Function LikeliFunc4(ByVal n As Integer, ByVal p As Single, ByVal Gamma As Single, ByVal mu As Single, ByVal Sigma As Single, ByVal Sigmab As Single, zi(), si()) As Double
On Error GoTo Errorhandler
    Dim s2(), sigma0(), mu0(), se() As Single
    Dim f1() As Double
    Dim f2() As Double
    Dim f() As Double
    Dim likelihoodfunction() As Double
    Dim res1() As Double
    Dim res2 As Double
    Dim i As Integer
    Dim TempPhi As Double
    ReDim sigma0(1 To n)
    ReDim mu0(1 To n)
    ReDim f1(1 To n)
    ReDim f2(1 To n)
    ReDim f(1 To n)
    ReDim likelihoodfunction(1 To n)
    ReDim res1(1 To n)
    ReDim s2(1 To n)
    ReDim se(1 To n)

    For i = 1 To n
        se(i) = Sqr(si(i) ^ 2 + Sigmab ^ 2)
    Next i
    
    For i = 1 To n
        s2(i) = Sigma ^ 2 + se(i) ^ 2
        sigma0(i) = 1 / Sqr(1 / Sigma ^ 2 + 1 / se(i) ^ 2)
        mu0(i) = (mu / Sigma ^ 2 + zi(i) / se(i) ^ 2) / (1 / Sigma ^ 2 + 1 / se(i) ^ 2)
        res1(i) = (Gamma - mu0(i)) / sigma0(i)
        res2 = (Gamma - mu) / Sigma
        
        TempPhi = 1 - Application.WorksheetFunction.Norm_S_Dist(res2, True)
        If TempPhi = 0 Then
        TempPhi = 1 * 1E-50
        End If
        
        
        f1(i) = p / (Sqr(se(i) ^ 2 * 2 * Pi)) * Exp(-(zi(i) - Gamma) ^ 2 / (2 * se(i) ^ 2))
        f2(i) = (1 - p) / (Sqr(2 * Pi * s2(i))) * ((1 - Application.WorksheetFunction.Norm_S_Dist(res1(i), True)) / TempPhi) * Exp(-(zi(i) - mu) ^ 2 / (2 * s2(i)))
        f(i) = f1(i) + f2(i)
        If f(i) > 0 Then
            likelihoodfunction(i) = Log(f(i))
        Else
            likelihoodfunction(i) = 0
        End If
    Next i
    LikeliFunc4 = Application.WorksheetFunction.Sum(likelihoodfunction)
    Exit Function
    
Errorhandler:
    If MsgBox("Sorry, you get an error due to unfit data or wrong initial parameters, do you want to try it again?", vbYesNo, "Warning") = vbYes Then
    Else
        Sheet11.MAM4Button.Value = False
        Sheet11.Max4Button.Value = False
    End
    End If
    
End Function