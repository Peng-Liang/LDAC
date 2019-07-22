Sub DoseRateFindAge()
On Error Resume Next
    With Sheet6
        .Range("C8:C20").ClearContents
        .Range("E8:E15").ClearContents
        .Range("C23:C32").ClearContents
        .Range("E23:E32").ClearContents
    End With
    
    With Sheet18
        .Range("H7:H25").ClearContents
        .Range("J7:J25").ClearContents
        .Range("J20:J21").ClearContents
        .Range("M8:M9").ClearContents
        .Range("O8:O9").ClearContents
        .Range("L10:P10").ClearContents
        .Range("M13:M25").ClearContents
        .Range("O13:O25").ClearContents
    End With
    Sheet17.Cells(28, 5).ClearContents
    Sheet4.Range("N40:P40").ClearContents
    Sheet4.Range("N41:P41").ClearContents
    
    Dim iteration As Integer
    Dim i, j, k As Integer
    Dim Diff As Integer
    Dim DoseRate() As Single
    Dim Age() As Single
    Dim CosmicRate() As Single
    Dim Alpha1() As Single
    Dim Beta1() As Single
    Dim Gamma1() As Single
    Dim Alpha2() As Single
    Dim Beta2() As Single
    Dim Gamma2() As Single
    Dim Dc() As Single
    Dim De() As Single
    Dim H2O() As Single
    Dim Depth() As Variant
    Dim alpha_value() As Single
    Dim User_Internal_Dr() As Single
    Dim ER() As Single 'external radionuclide
    Dim IR() As Single 'internal radionuclide
    Dim IRContent() As Single
    Dim UContent() As Single
    Dim ThContent() As Single
    Dim KContent() As Single
    Dim RbContent() As Single
    Dim User_External_Dr() As Single 'alpha, beta, gamma
    Dim Dr1() As Single
    Dim Dr2() As Single
    Dim ExternalDr() As Single
    Dim InternalDr() As Single

    For i = 46 To 51
        If (IsEmpty(Sheet4.Cells(i, 11)) = False) Then
           Sheet6.Cells(8, 3) = Sheet4.Cells(i, 11).Value2
           Sheet6.Cells(8, 5) = Sheet4.Cells(i, 13).Value2
        End If
    Next i
    
    For k = 1 To 4
        If IsEmpty(Sheet4.Cells(k + 93, 5)) = False Then
            Sheet6.Cells(k + 23, 3) = Sheet4.Cells(k + 93, 5).Value2
            Sheet6.Cells(k + 23, 5) = Sheet4.Cells(k + 93, 9).Value2
        End If
    Next
    
    For j = 1 To 5
        If IsEmpty(Sheet4.Cells(j + 92, 14)) = False Then
            Sheet6.Cells(j + 27, 3) = Sheet4.Cells(j + 92, 14).Value2
            Sheet6.Cells(j + 27, 5) = Sheet4.Cells(j + 92, 16).Value2
        End If
    Next j
    
        If Sheet4.DatumBox.Value = False Then
            Diff = 0
            Sheet17.Cells(33, 8).ClearContents
        ElseIf Sheet4.DatumBox.Value = True Then
            Diff = Year(Sheet4.Cells(34, 5).Value2) - Sheet4.Cells(55, 16).Value2
            Sheet17.Cells(33, 8).Value = "Datum: " & Sheet4.Cells(55, 16).Value2 & " C.E."
        End If
    Sheet10.Cells(44, 12).Value = Diff

    If Sheet4.MonteCarloCheck = False Then
        Sheet17.Cells(28, 5) = "Quadrature"
        With Sheet6
            .Cells(9, 3) = Sheet4.Cells(17, 14).Value2
            .Cells(9, 5) = Sheet4.Cells(17, 16).Value2
            .Cells(10, 3) = Sheet4.Cells(18, 14).Value2
            .Cells(10, 5) = Sheet4.Cells(18, 16).Value2
            .Cells(11, 3) = Sheet4.Cells(19, 14).Value2
            .Cells(11, 5) = Sheet4.Cells(19, 16).Value2
            .Cells(12, 3) = Sheet4.Cells(20, 14).Value2
            .Cells(12, 5) = Sheet4.Cells(20, 16).Value2
            .Cells(13, 3) = Sheet4.Cells(21, 5).Value2
            .Cells(13, 5) = Sheet4.Cells(21, 9).Value2
            .Cells(14, 3) = Sheet4.Cells(25, 5).Value2
            .Cells(14, 5) = Sheet4.Cells(25, 9).Value2
            .Cells(15, 3) = Sheet4.Cells(27, 14).Value2
            .Cells(15, 5) = Sheet4.Cells(27, 16).Value2
            .Cells(16, 3) = Sheet4.Cells(22, 12).Value2
            .Cells(17, 3) = Sheet4.Cells(23, 12).Value2
            .Cells(18, 3) = Sheet4.Cells(24, 5).Value2
            .Cells(19, 3) = Sheet4.Cells(16, 5).Value2
            .Cells(20, 3) = Sheet4.Cells(16, 9).Value2
            .Cells(23, 3) = Sheet4.Cells(28, 14).Value2
            .Cells(23, 5) = Sheet4.Cells(28, 16).Value2
        End With
        
        Sheet4.Cells(41, 14) = Round(Sheet6.Cells(21, 8).Value2, 5) & " " + ChrW(177) + " " & Round(Sheet6.Cells(21, 10).Value2, 5) & " mGy/yr"
        Sheet4.Cells(40, 14) = Round(Sheet6.Cells(25, 8).Value2, 5) & " " + ChrW(177) + " " & Round(Sheet6.Cells(25, 10).Value2, 5) & " mGy/yr"
        Sheet4.Cells(59, 12) = Round(Sheet6.Cells(25, 8).Value2, 2) & " " + ChrW(177) + " " & Round(Sheet6.Cells(25, 10).Value2, 2)
        Sheet4.Cells(59, 14) = Round(Sheet6.Cells(21, 8).Value2, 2) & " " + ChrW(177) + " " & Round(Sheet6.Cells(21, 10).Value2, 2)
        If IsEmpty(Sheet4.Cells(59, 8)) Then
            MsgBox ("Please Calculate the equivalent dose first!")
            Application.StatusBar = "Finished"
            Exit Sub
        End If
        
        For i = 46 To 51
            If (IsEmpty(Sheet4.Cells(i, 11)) = False) Then
                If Sheet6.Cells(20, 8) >= 50000 Then
                Sheet4.Cells(45, 14) = "Age (ka)"
                Sheet4.Cells(i, 14) = Round(Sheet6.Cells(20, 8).Value2 / 1000, 2)
                Sheet4.Cells(i, 16) = Round(Sheet6.Cells(20, 10).Value2 / 1000, 2)
                Else
                Sheet4.Cells(45, 14) = "Age (year)"
                Sheet4.Cells(i, 14) = Round(Sheet6.Cells(20, 8).Value2, 2)
                Sheet4.Cells(i, 16) = Round(Sheet6.Cells(20, 10).Value2, 2)
                End If
            End If
        Next i

        If Sheet6.Cells(20, 8) >= 50000 Then
            Sheet4.Cells(59, 15) = Round((Sheet6.Cells(20, 8).Value2 - Diff) / 1000, 2) & " " + ChrW(177) + " " & Round(Sheet6.Cells(20, 10).Value2 / 1000, 2)
        Else
            Sheet4.Cells(59, 15) = Application.WorksheetFunction.MRound((Sheet6.Cells(20, 8).Value2 - Diff), 5) & " " + ChrW(177) + " " & Application.WorksheetFunction.MRound(Sheet6.Cells(20, 10).Value2, 5)
            If Application.WorksheetFunction.MRound(Sheet6.Cells(20, 10).Value2, 5) = 0 Then
           Sheet4.Cells(59, 15) = Application.WorksheetFunction.MRound((Sheet6.Cells(20, 8).Value2 - Diff), 5) & " " + ChrW(177) + " " & 5
            End If
        End If
        

        
        Application.StatusBar = "Plotting the Final results..."
        Call DrawNormalDistribution
        Application.StatusBar = "Finished"
    '===============================Monte Carlo simulation=========================================================================================================================
    ElseIf Sheet4.MonteCarloCheck = True Then
        Application.StatusBar = "Monte-Carlo Simulation started..."
        iteration = Sheet4.Cells(53, 16).Value
        Sheet17.Cells(28, 5) = "Monte-Carlo simulation (repeats:" & iteration & ")"
        'calculate the mean value use the original data of mean
        With Sheet6
            .Cells(9, 3) = Sheet4.Cells(17, 14).Value2
            .Cells(9, 5) = Sheet4.Cells(17, 16).Value2
            .Cells(10, 3) = Sheet4.Cells(18, 14).Value2
            .Cells(10, 5) = Sheet4.Cells(18, 16).Value2
            .Cells(11, 3) = Sheet4.Cells(19, 14).Value2
            .Cells(11, 5) = Sheet4.Cells(19, 16).Value2
            .Cells(12, 3) = Sheet4.Cells(20, 14).Value2
            .Cells(12, 5) = Sheet4.Cells(20, 16).Value2
            .Cells(13, 3) = Sheet4.Cells(21, 5).Value2
            .Cells(13, 5) = Sheet4.Cells(21, 9).Value2
            .Cells(14, 3) = Sheet4.Cells(25, 5).Value2
            .Cells(14, 5) = Sheet4.Cells(25, 9).Value2
            .Cells(15, 3) = Sheet4.Cells(27, 14).Value2
            .Cells(15, 5) = Sheet4.Cells(27, 16).Value2
            .Cells(16, 3) = Sheet4.Cells(22, 12).Value2
            .Cells(17, 3) = Sheet4.Cells(23, 12).Value2
            .Cells(18, 3) = Sheet4.Cells(24, 5).Value2
            .Cells(19, 3) = Sheet4.Cells(16, 5).Value2
            .Cells(20, 3) = Sheet4.Cells(16, 9).Value2
            .Cells(23, 3) = Sheet4.Cells(28, 14).Value2
            .Cells(23, 5) = Sheet4.Cells(28, 16).Value2
        End With
   
        'estimate the uncertainty using Monte-carlo simulation
        ReDim DoseRate(1 To iteration)
        ReDim Age(1 To iteration)
        ReDim CosmicRate(1 To iteration)
        ReDim Alpha1(1 To iteration)
        ReDim Beta1(1 To iteration)
        ReDim Gamma1(1 To iteration)
        ReDim Alpha2(1 To iteration)
        ReDim Beta2(1 To iteration)
        ReDim Gamma2(1 To iteration)
        ReDim Dc(1 To iteration)
        ReDim De(1 To iteration)
        ReDim H2O(1 To iteration)
        ReDim Depth(1 To iteration)
        ReDim alpha_value(1 To iteration)
        ReDim User_Internal_Dr(1 To iteration)
        ReDim ER(1 To 9, 1 To iteration)
        ReDim IR(1 To 6, 1 To iteration)
        ReDim IRContent(1 To 4, 1 To iteration)
        ReDim UContent(1 To iteration)
        ReDim ThContent(1 To iteration)
        ReDim KContent(1 To iteration)
        ReDim RbContent(1 To iteration)
        ReDim User_External_Dr(1 To 3, 1 To iteration)
        ReDim Dr1(1 To 12, 1 To iteration)
        ReDim Dr2(1 To 6, 1 To iteration)
        ReDim ExternalDr(1 To iteration)
        ReDim InternalDr(1 To iteration)
        
        
        For i = 1 To iteration
            Application.StatusBar = "Monte-Carlo Simulation for uncertainties estimation......" & Application.WorksheetFunction.MRound(i / iteration * 100, 10) & "% Completed"
            
            For k = 46 To 51
                If (IsEmpty(Sheet4.Cells(k, 11)) = False) Then
                De(i) = Sheet4.Cells(k, 11) + Sheet4.Cells(k, 13) * Application.WorksheetFunction.Norm_S_Inv(Rnd())
                End If
            Next k
        'external radionuclide content
                UContent(i) = Sheet4.Cells(17, 14) + Application.WorksheetFunction.Norm_S_Inv(Rnd()) * Sheet4.Cells(17, 16)
                ThContent(i) = Sheet4.Cells(18, 14).Value2 + Application.WorksheetFunction.Norm_S_Inv(Rnd()) * Sheet4.Cells(18, 16).Value2
                KContent(i) = Sheet4.Cells(19, 14).Value2 + Application.WorksheetFunction.Norm_S_Inv(Rnd()) * Sheet4.Cells(19, 16).Value2
                RbContent(i) = Sheet4.Cells(20, 14).Value2 + Application.WorksheetFunction.Norm_S_Inv(Rnd()) * Sheet4.Cells(20, 16).Value2
                ER(1, i) = UContent(i)
                ER(3, i) = UContent(i)
                ER(7, i) = UContent(i)
                ER(2, i) = ThContent(i)
                ER(4, i) = ThContent(i)
                ER(8, i) = ThContent(i)
                ER(5, i) = KContent(i)
                ER(9, i) = KContent(i)
                ER(6, i) = RbContent(i)
                H2O(i) = Sheet4.Cells(21, 5).Value2 + Application.WorksheetFunction.Norm_S_Inv(Rnd()) * Sheet4.Cells(21, 9).Value2
                Depth(i) = Sheet4.Cells(25, 5).Value2 + Application.WorksheetFunction.Norm_S_Inv(Rnd()) * Sheet4.Cells(25, 9).Value2
                alpha_value(i) = Sheet4.Cells(28, 14).Value2 + Application.WorksheetFunction.Norm_S_Inv(Rnd()) * Sheet4.Cells(28, 16).Value2
            
            'optional radionuclide contents
            For k = 1 To 4
                If IsEmpty(Sheet4.Cells(k + 93, 5)) = False Then
                IRContent(k, i) = Sheet4.Cells(k + 93, 5).Value2 + Sheet4.Cells(k + 93, 9).Value2 * Application.WorksheetFunction.Norm_S_Inv(Rnd())
                Else
                IRContent(k, i) = Null
                End If
            Next
                IR(1, i) = IRContent(1, i)
                IR(3, i) = IRContent(1, i)
                IR(2, i) = IRContent(2, i)
                IR(4, i) = IRContent(2, i)
                IR(5, i) = IRContent(3, i)
                IR(6, i) = IRContent(4, i)

            For j = 1 To 3
                If IsEmpty(Sheet4.Cells(j + 92, 14)) = False Then
                User_External_Dr(j, i) = Sheet4.Cells(j + 92, 14).Value2 + Sheet4.Cells(j + 92, 16).Value2 * Application.WorksheetFunction.Norm_S_Inv(Rnd())
                Else
                User_External_Dr(j, i) = Null
                End If
            Next j
            
            If IsEmpty(Sheet4.Cells(96, 14)) = False Then
                Dc(i) = Sheet4.Cells(96, 14).Value2 + Sheet4.Cells(96, 16).Value2 * Application.WorksheetFunction.Norm_S_Inv(Rnd())
            Else
                Dc(i) = Null
            End If
 
    '=================================================
    'Calculation of the dose rate (mGy/yr)
    'Calculation of the infinite matrix dose rates
 
    For j = 1 To 9
        If Sheet4.Cells(27, 5) = "Adamiec1998" Then
            Dr1(j, i) = ER(j, i) * (Sheet7.Cells(j + 3, 3).Value + Sheet7.Cells(j + 3, 3).Value * (Sheet7.Cells(j + 3, 6).Value / Sheet7.Cells(j + 3, 5).Value) * Application.WorksheetFunction.Norm_S_Inv(Rnd()))
        ElseIf Sheet4.Cells(27, 5) = "Liritzis2013" Then
            Dr1(j, i) = ER(j, i) * (Sheet7.Cells(j + 3, 5).Value + Sheet7.Cells(j + 3, 5).Value) * Application.WorksheetFunction.Norm_S_Inv(Rnd())
        Else
            Dr1(j, i) = ER(j, i) * (Sheet7.Cells(j + 3, 4).Value + Sheet7.Cells(j + 3, 4).Value * (Sheet7.Cells(j + 3, 6).Value / Sheet7.Cells(j + 3, 5).Value) * Application.WorksheetFunction.Norm_S_Inv(Rnd()))
        End If
    Next j
    
    For j = 10 To 12
            Dr1(j, i) = Sheet4.Cells(j + 83, 14).Value2 + Sheet4.Cells(j + 83, 16).Value2 * Application.WorksheetFunction.Norm_S_Inv(Rnd())
    Next j
    
    For j = 1 To 6
        If Sheet4.Cells(27, 5) = "Adamiec1998" Then
            Dr2(j, i) = IR(j, i) * (Sheet7.Cells(j + 3, 3).Value + Sheet7.Cells(j + 3, 3).Value * (Sheet7.Cells(j + 3, 6).Value / Sheet7.Cells(j + 3, 5).Value) * Application.WorksheetFunction.Norm_S_Inv(Rnd()))
        ElseIf Sheet4.Cells(27, 5) = "Liritzis2013" Then
            Dr2(j, i) = IR(j, i) * (Sheet7.Cells(j + 3, 5).Value + Sheet7.Cells(j + 3, 5).Value) * Application.WorksheetFunction.Norm_S_Inv(Rnd())
        Else
            Dr2(j, i) = IR(j, i) * (Sheet7.Cells(j + 3, 4).Value + Sheet7.Cells(j + 3, 4).Value * (Sheet7.Cells(j + 3, 6).Value / Sheet7.Cells(j + 3, 5).Value) * Application.WorksheetFunction.Norm_S_Inv(Rnd()))
        End If
    Next j
    
       
    'Gamma dose scaling at shallow depths
    If IsEmpty(Sheet4.Cells(25, 5)) = False And Depth(i) <= 0.3 Then
        For j = 7 To 9
            Dr1(j, i) = Dr1(j, i) * Application.WorksheetFunction.VLookup(Application.WorksheetFunction.MRound(Depth(i), 0.005), Sheet14.Range("H6:L66"), j - 5)
        Next j
        Dr1(12, i) = Dr1(12, i) * Application.WorksheetFunction.VLookup(Application.WorksheetFunction.MRound(Depth(i), 0.005), Sheet14.Range("H6:L66"), 5)
    End If

    'alpha Grain size attenuation
    If Sheet4.Cells(29, 14).Value = Sheet8.Cells(47, 2).Value Then
        For j = 1 To 2
            Dr1(j, i) = Dr1(j, i) * (Sheet8.Cells(47, 2 * j + 2).Value + Sheet8.Cells(47, 2 * j + 3).Value * Application.WorksheetFunction.Norm_S_Inv(Rnd()))
            Dr2(j, i) = Dr2(j, i) * (Sheet8.Cells(51, 2 * j + 2).Value + Sheet8.Cells(51, 2 * j + 3).Value * Application.WorksheetFunction.Norm_S_Inv(Rnd()))
        Next j
            Dr1(10, i) = Dr1(10, i) * (Sheet8.Cells(47, 8).Value + Sheet8.Cells(47, 9).Value * Application.WorksheetFunction.Norm_S_Inv(Rnd()))
    Else
        For j = 1 To 2
            Dr1(j, i) = Dr1(j, i) * (Sheet8.Cells(48, 2 * j + 2).Value + Sheet8.Cells(48, 2 * j + 3).Value * Application.WorksheetFunction.Norm_S_Inv(Rnd()))
            Dr2(j, i) = Dr2(j, i) * (Sheet8.Cells(52, 2 * j + 2).Value + Sheet8.Cells(52, 2 * j + 3).Value * Application.WorksheetFunction.Norm_S_Inv(Rnd()))
        Next j
            Dr1(10, i) = Dr1(10, i) * (Sheet8.Cells(48, 8).Value + Sheet8.Cells(48, 9).Value * Application.WorksheetFunction.Norm_S_Inv(Rnd()))
    End If
    
    
    'beta Grain size attenuation
    For j = 3 To 5
        If Sheet4.Cells(28, 7).Value = Sheet9.Cells(55, 22).Value Then
            Dr1(j, i) = Dr1(j, i) * (Sheet9.Cells(55, 2 * j + 18).Value + Sheet9.Cells(55, 2 * j + 19).Value * Application.WorksheetFunction.Norm_S_Inv(Rnd()))
            Dr2(j, i) = Dr2(j, i) * (Sheet9.Cells(62, 2 * j + 18).Value + Sheet9.Cells(62, 2 * j + 19).Value * Application.WorksheetFunction.Norm_S_Inv(Rnd()))
        ElseIf Sheet4.Cells(28, 7).Value = Sheet9.Cells(56, 22).Value Then
            Dr1(j, i) = Dr1(j, i) * (Sheet9.Cells(56, 2 * j + 18).Value + Sheet9.Cells(56, 2 * j + 19).Value * Application.WorksheetFunction.Norm_S_Inv(Rnd()))
            Dr2(j, i) = Dr2(j, i) * (Sheet9.Cells(63, 2 * j + 18).Value + Sheet9.Cells(63, 2 * j + 19).Value * Application.WorksheetFunction.Norm_S_Inv(Rnd()))
        ElseIf Sheet4.Cells(28, 7).Value = Sheet9.Cells(57, 22).Value Then
            Dr1(j, i) = Dr1(j, i) * (Sheet9.Cells(57, 2 * j + 18).Value + Sheet9.Cells(57, 2 * j + 19).Value * Application.WorksheetFunction.Norm_S_Inv(Rnd()))
            Dr2(j, i) = Dr2(j, i) * (Sheet9.Cells(64, 2 * j + 18).Value + Sheet9.Cells(64, 2 * j + 19).Value * Application.WorksheetFunction.Norm_S_Inv(Rnd()))
        ElseIf Sheet4.Cells(28, 7).Value = Sheet9.Cells(58, 22).Value Then
            Dr1(j, i) = Dr1(j, i) * (Sheet9.Cells(58, 2 * j + 18).Value + Sheet9.Cells(58, 2 * j + 19).Value * Application.WorksheetFunction.Norm_S_Inv(Rnd()))
            Dr2(j, i) = Dr2(j, i) * (Sheet9.Cells(65, 2 * j + 18).Value + Sheet9.Cells(65, 2 * j + 19).Value * Application.WorksheetFunction.Norm_S_Inv(Rnd()))
        End If
    Next j
    'Readhead2002 for Rb Grain size attenuation
    Dr1(6, i) = Dr1(6, i) * (Sheet9.Cells(59, 32).Value + Sheet9.Cells(58, 33).Value * Application.WorksheetFunction.Norm_S_Inv(Rnd()))
    Dr2(6, i) = Dr2(6, i) * (Sheet9.Cells(66, 32).Value + Sheet9.Cells(66, 33).Value * Application.WorksheetFunction.Norm_S_Inv(Rnd()))
    
    'if combined
    If Sheet4.Cells(28, 7).Value = Sheet9.Cells(55, 22).Value Then
        Dr1(11, i) = Dr1(11, i) * (Sheet9.Cells(55, 30).Value + Sheet9.Cells(55, 31).Value * Application.WorksheetFunction.Norm_S_Inv(Rnd()))
    ElseIf Sheet4.Cells(28, 7).Value = Sheet9.Cells(56, 22).Value Then
        Dr1(11, i) = Dr1(11, i) * (Sheet9.Cells(56, 30).Value + Sheet9.Cells(56, 31).Value * Application.WorksheetFunction.Norm_S_Inv(Rnd()))
    ElseIf Sheet4.Cells(28, 7).Value = Sheet9.Cells(57, 22).Value Then
        Dr1(11, i) = Dr1(11, i) * (Sheet9.Cells(57, 30).Value + Sheet9.Cells(57, 31).Value * Application.WorksheetFunction.Norm_S_Inv(Rnd()))
    ElseIf Sheet4.Cells(28, 7).Value = Sheet9.Cells(58, 22).Value Then
        Dr1(11, i) = Dr1(11, i) * (Sheet9.Cells(58, 30).Value + Sheet9.Cells(58, 31).Value * Application.WorksheetFunction.Norm_S_Inv(Rnd()))
    End If
    
    'Etch depth attenuation
        'alpha
        For j = 1 To 2
            Dr1(j, i) = Dr1(j, i) * (Sheet12.Cells(42, 2 * j + 8).Value + Sheet12.Cells(42, 2 * j + 9).Value * Application.WorksheetFunction.Norm_S_Inv(Rnd()))
            Dr2(j, i) = Dr2(j, i) * (Sheet12.Cells(43, 2 * j + 8).Value + Sheet12.Cells(43, 2 * j + 9).Value * Application.WorksheetFunction.Norm_S_Inv(Rnd()))
        Next j
        Dr1(10, i) = Dr1(10, i) * (Sheet12.Cells(42, 14).Value + Sheet12.Cells(42, 15).Value * Application.WorksheetFunction.Norm_S_Inv(Rnd()))
        
        'beta
        For j = 3 To 5
            If Sheet4.Cells(29, 6).Value = Sheet13.Cells(33, 13).Value Then
                Dr1(j, i) = Dr1(j, i) * (Sheet13.Cells(33, 2 * j + 9).Value + Sheet13.Cells(33, 2 * j + 10).Value * Application.WorksheetFunction.Norm_S_Inv(Rnd()))
                Dr2(j, i) = Dr2(j, i) * (Sheet13.Cells(37, 2 * j + 9).Value + Sheet13.Cells(37, 2 * j + 10).Value * Application.WorksheetFunction.Norm_S_Inv(Rnd()))
            ElseIf Sheet4.Cells(29, 6).Value = Sheet13.Cells(34, 13).Value Then
                Dr1(j, i) = Dr1(j, i) * (Sheet13.Cells(34, 2 * j + 9).Value + Sheet13.Cells(34, 2 * j + 10).Value * Application.WorksheetFunction.Norm_S_Inv(Rnd()))
                Dr2(j, i) = Dr2(j, i) * (Sheet13.Cells(38, 2 * j + 9).Value + Sheet13.Cells(38, 2 * j + 10).Value * Application.WorksheetFunction.Norm_S_Inv(Rnd()))
            End If
        Next j
        
        If Sheet4.Cells(29, 6).Value = Sheet13.Cells(33, 13).Value Then
            Dr1(11, i) = Dr1(11, i) * (Sheet13.Cells(33, 21).Value + Sheet13.Cells(33, 22).Value * Application.WorksheetFunction.Norm_S_Inv(Rnd()))
        ElseIf Sheet4.Cells(29, 6).Value = Sheet13.Cells(34, 13).Value Then
            Dr1(11, i) = Dr1(11, i) * (Sheet13.Cells(34, 21).Value + Sheet13.Cells(34, 22).Value * Application.WorksheetFunction.Norm_S_Inv(Rnd()))
        End If

    'alpha efficifiency correction
            For j = 1 To 2
                Dr1(j, i) = Dr1(j, i) * alpha_value(i)
                Dr2(j, i) = Dr2(j, i) * alpha_value(i)
            Next j
            Dr1(10, i) = Dr1(10, i) * alpha_value(i)
    
    'Water content attenuation
            For j = 1 To 2
                Dr1(j, i) = Dr1(j, i) / (1 + 1.49 * (H2O(i) / 100))
            Next j
            Dr1(10, i) = Dr1(10, i) / (1 + 1.49 * (H2O(i) / 100))
            
            For j = 3 To 6
                Dr1(j, i) = Dr1(j, i) / (1 + 1.25 * (H2O(i) / 100))
            Next j
            Dr1(11, i) = Dr1(11, i) / (1 + 1.25 * (H2O(i) / 100))
            
            For j = 7 To 9
                Dr1(j, i) = Dr1(j, i) / (1 + 1.14 * (H2O(i) / 100))
            Next j
            Dr1(12, i) = Dr1(12, i) / (1 + 1.14 * (H2O(i) / 100))
    
    'external alpha, beta, gamma
    If IsEmpty(Sheet4.Cells(93, 14)) = True Then
        Alpha1(i) = Dr1(1, i) + Dr1(2, i)
    Else
        Alpha1(i) = Dr1(10, i)
    End If
    
    If IsEmpty(Sheet4.Cells(94, 14)) = True Then
        Beta1(i) = Dr1(3, i) + Dr1(4, i) + Dr1(5, i) + Dr1(6, i)
    Else
        Beta1(i) = Dr1(11, i)
    End If
    
    If IsEmpty(Sheet4.Cells(95, 14)) = True Then
        Gamma1(i) = Dr1(7, i) + Dr1(8, i) + Dr1(9, i)
    Else
        Gamma1(i) = Dr1(12, i)
    End If
    
    If IsEmpty(Sheet4.Cells(96, 14)) = True Then
            CosmicRate(i) = Sheet6.Cells(121, 18).Value + Sheet6.Cells(121, 20).Value * Application.WorksheetFunction.Norm_S_Inv(Rnd())
    Else
            CosmicRate(i) = Dc(i)
    End If
    
    ExternalDr(i) = Alpha1(i) + Beta1(i) + Gamma1(i) + CosmicRate(i)
    
    'internal alpha, beta, gamma
    If IsEmpty(Sheet4.Cells(97, 14)) = True Then
        Alpha2(i) = Dr2(1, i) + Dr2(2, i)
        Beta2(i) = Dr2(3, i) + Dr2(4, i) + Dr2(5, i) + Dr2(6, i)
        InternalDr(i) = Alpha2(i) + Beta2(i)
    Else
        InternalDr(i) = Sheet4.Cells(97, 14).Value2 + Sheet4.Cells(97, 16).Value2 * Application.WorksheetFunction.Norm_S_Inv(Rnd())
    End If

    DoseRate(i) = ExternalDr(i) + InternalDr(i)
    Age(i) = De(i) * 1000 / DoseRate(i)
Next i

Application.StatusBar = "Outputting the Monte-Carlo results..."
'MsgBox Application.Average(InternalDR)
    With Sheet18
    .Cells(7, 8) = Application.WorksheetFunction.Average(De)
    .Cells(7, 10) = Application.WorksheetFunction.StDev_S(De)
    .Cells(8, 8) = Application.WorksheetFunction.Average(UContent)
    .Cells(8, 10) = Application.WorksheetFunction.StDev_S(UContent)
    .Cells(9, 8) = Application.WorksheetFunction.Average(ThContent)
    .Cells(9, 10) = Application.WorksheetFunction.StDev_S(ThContent)
    .Cells(10, 8) = Application.WorksheetFunction.Average(KContent)
    .Cells(10, 10) = Application.WorksheetFunction.StDev_S(KContent)
    .Cells(11, 8) = Application.WorksheetFunction.Average(RbContent)
    .Cells(11, 10) = Application.WorksheetFunction.StDev_S(RbContent)
    .Cells(12, 8) = Application.WorksheetFunction.Average(H2O)
    .Cells(12, 10) = Application.WorksheetFunction.StDev_S(H2O)
     .Cells(13, 8) = Application.WorksheetFunction.Average(Depth)
    .Cells(13, 10) = Application.WorksheetFunction.StDev_S(Depth)
    
    For j = 7 To 13
        If .Cells(j, 8).Value2 = 0 Then .Cells(j, 8).Value2 = ""
        If .Cells(j, 10).Value2 = 0 Then .Cells(j, 10).Value2 = ""
    Next j

    .Cells(16, 8) = Application.WorksheetFunction.Average(alpha_value)
    .Cells(16, 10) = Application.WorksheetFunction.StDev_S(alpha_value)
    .Cells(17, 8) = Application.WorksheetFunction.Average(Application.WorksheetFunction.Index(IRContent, 1, 0))
    .Cells(17, 10) = Application.WorksheetFunction.StDev_S(Application.WorksheetFunction.Index(IRContent, 1, 0))
    .Cells(18, 8) = Application.WorksheetFunction.Average(Application.WorksheetFunction.Index(IRContent, 2, 0))
    .Cells(18, 10) = Application.WorksheetFunction.StDev_S(Application.WorksheetFunction.Index(IRContent, 2, 0))
    .Cells(19, 8) = Application.WorksheetFunction.Average(Application.WorksheetFunction.Index(IRContent, 3, 0))
    .Cells(19, 10) = Application.WorksheetFunction.StDev_S(Application.WorksheetFunction.Index(IRContent, 3, 0))
    .Cells(20, 8) = Application.WorksheetFunction.Average(Application.WorksheetFunction.Index(IRContent, 4, 0))
    .Cells(20, 10) = Application.WorksheetFunction.StDev_S(Application.WorksheetFunction.Index(IRContent, 4, 0))
    .Cells(21, 8) = Application.WorksheetFunction.Average(Application.WorksheetFunction.Index(User_External_Dr, 1, 0))
    .Cells(21, 10) = Application.WorksheetFunction.StDev_S(Application.WorksheetFunction.Index(User_External_Dr, 1, 0))
    .Cells(22, 8) = Application.WorksheetFunction.Average(Application.WorksheetFunction.Index(User_External_Dr, 2, 0))
    .Cells(22, 10) = Application.WorksheetFunction.StDev_S(Application.WorksheetFunction.Index(User_External_Dr, 2, 0))
    .Cells(23, 8) = Application.WorksheetFunction.Average(Application.WorksheetFunction.Index(User_External_Dr, 3, 0))
    .Cells(23, 10) = Application.WorksheetFunction.StDev_S(Application.WorksheetFunction.Index(User_External_Dr, 3, 0))
    .Cells(24, 8) = Application.WorksheetFunction.Average(Dc)
    .Cells(24, 10) = Application.WorksheetFunction.StDev_S(Dc)
    If IsEmpty(Sheet4.Cells(97, 14)) = False Then
        .Cells(25, 8) = Application.WorksheetFunction.Average(InternalDr)
        .Cells(25, 10) = Application.WorksheetFunction.StDev_S(InternalDr)
    End If
    
    For j = 17 To 24
        If .Cells(j, 8).Value2 = 0 Then .Cells(j, 8).Value2 = ""
        If .Cells(j, 10).Value2 = 0 Then .Cells(j, 10).Value2 = ""
    Next j
    
    .Cells(13, 13) = Application.WorksheetFunction.Average(Age)
    .Cells(13, 15) = Application.WorksheetFunction.StDev_S(Age)
    .Cells(14, 13) = Application.WorksheetFunction.Average(DoseRate)
    .Cells(14, 15) = Application.WorksheetFunction.StDev_S(DoseRate)
    
    .Cells(16, 13) = Application.WorksheetFunction.Average(ExternalDr)
    .Cells(16, 15) = Application.WorksheetFunction.StDev_S(ExternalDr)
    .Cells(17, 13) = Application.WorksheetFunction.Average(InternalDr)
    .Cells(17, 15) = Application.WorksheetFunction.StDev_S(InternalDr)
    
    If IsEmpty(Sheet4.Cells(96, 14)) = True Then
        .Cells(18, 13) = Sheet6.Cells(121, 18).Value
        .Cells(18, 15) = Sheet6.Cells(121, 18).Value * 0.1
    Else
        .Cells(18, 13) = Sheet4.Cells(96, 14).Value
        .Cells(18, 15) = Sheet4.Cells(96, 14).Value * 0.1
    End If
    
    .Cells(20, 13) = Application.WorksheetFunction.Average(Alpha1)
    .Cells(20, 15) = Application.WorksheetFunction.StDev_S(Alpha1)
    .Cells(21, 13) = Application.WorksheetFunction.Average(Beta1)
    .Cells(21, 15) = Application.WorksheetFunction.StDev_S(Beta1)
    .Cells(22, 13) = Application.WorksheetFunction.Average(Gamma1)
    .Cells(22, 15) = Application.WorksheetFunction.StDev_S(Gamma1)
    
    .Cells(24, 13) = Application.WorksheetFunction.Average(Alpha2)
    .Cells(24, 15) = Application.WorksheetFunction.StDev_S(Alpha2)
    .Cells(25, 13) = Application.WorksheetFunction.Average(Beta2)
    .Cells(25, 15) = Application.WorksheetFunction.StDev_S(Beta2)
    
    .Cells(8, 13) = Sheet6.Cells(198, 15).Value2
    .Cells(9, 13) = Sheet6.Cells(193, 15).Value2
    .Cells(8, 15) = Application.WorksheetFunction.StDev_S(Age)
    .Cells(9, 15) = Application.WorksheetFunction.StDev_S(DoseRate) '(Application.WorksheetFunction.Large(doserate, iteration * 0.1585) - Application.WorksheetFunction.Small(doserate, iteration * 0.1585)) / 2
    End With
    
        
        If Sheet18.Cells(8, 13) < 10 Then
            Sheet18.Cells(10, 12) = "Asymmetric Age (1" & ChrW(963) & "): [" & Round(Application.WorksheetFunction.Small(Age, iteration * 0.1585) - Diff, 0) & ", " & Round(Application.WorksheetFunction.Large(Age, iteration * 0.1585) - Diff, 0) & "]"
        ElseIf Sheet18.Cells(8, 13) < 50000 Then
            Sheet18.Cells(10, 12) = "Asymmetric Age (1" & ChrW(963) & "): [" & Application.WorksheetFunction.MRound(Application.WorksheetFunction.Small(Age, iteration * 0.1585) - Diff, 5) & ", " & Application.WorksheetFunction.MRound(Application.WorksheetFunction.Large(Age, iteration * 0.1585) - Diff, 5) & "]"
        Else
            Sheet18.Cells(10, 12) = "Asymmetric Age (1" & ChrW(963) & "): [" & Round((Application.WorksheetFunction.Small(Age, iteration * 0.1585) - Diff) / 1000, 2) & ", " & Round((Application.WorksheetFunction.Large(Age, iteration * 0.1585) - Diff) / 1000, 2) & "]"
        End If
    
        With Sheet4
            .Cells(41, 14) = Round(Sheet18.Cells(9, 13).Value2, 5) & " " + ChrW(177) + " " & Round(Sheet18.Cells(9, 15).Value2, 5) & " mGy/yr"
            .Cells(40, 14) = Round(Sheet18.Cells(18, 13).Value2, 5) & " " + ChrW(177) + " " & Round(Sheet18.Cells(18, 15).Value2, 5) & " mGy/yr"
            .Cells(59, 12) = Round(Sheet18.Cells(18, 3).Value2, 2) & " " + ChrW(177) + " " & Round(Sheet18.Cells(18, 5).Value2, 2)
            .Cells(59, 14) = Round(Sheet18.Cells(9, 13).Value2, 2) & " " + ChrW(177) + " " & Round(Sheet18.Cells(9, 15).Value2, 2)
        End With
         
        If IsEmpty(Sheet4.Cells(59, 8)) Then
                MsgBox ("Please Calculate the equivalent dose first!")
                Application.StatusBar = "Finished"
            Exit Sub
        End If
        For i = 46 To 51
            If (IsEmpty(Sheet4.Cells(i, 11)) = False) Then
              If Sheet18.Cells(8, 13) >= 50000 Then
                Sheet4.Cells(45, 14) = "Age (ka)"
                Sheet4.Cells(i, 14) = Round(Sheet18.Cells(8, 13).Value2 / 1000, 2)
                Sheet4.Cells(i, 16) = Round(Sheet18.Cells(8, 15).Value2 / 1000, 2)
                Else
                Sheet4.Cells(45, 14) = "Age (year)"
                Sheet4.Cells(i, 14) = Round(Sheet18.Cells(8, 13).Value2, 2)
                Sheet4.Cells(i, 16) = Round(Sheet18.Cells(8, 15).Value2, 2)
                End If
            End If
        Next i
            
        If Sheet18.Cells(8, 13) < 50000 Then
            Sheet4.Cells(59, 15) = Application.WorksheetFunction.MRound((Sheet18.Cells(8, 13).Value2 - Diff), 5) & " " + ChrW(177) + " " & Application.WorksheetFunction.MRound(Sheet18.Cells(8, 15).Value2, 5)
            If Application.WorksheetFunction.MRound(Sheet18.Cells(8, 15).Value2, 5) = 0 Then
               Sheet4.Cells(59, 15) = Application.WorksheetFunction.MRound((Sheet18.Cells(8, 13).Value2 - Diff), 5) & " " + ChrW(177) + " " & 5
            End If
        Else
            Sheet4.Cells(59, 15) = Round((Sheet18.Cells(8, 13).Value2 - Diff) / 1000, 2) & " " + ChrW(177) + " " & Round(Sheet18.Cells(8, 15).Value2 / 1000, 2)
        End If
    Application.StatusBar = "Plotting the Monte-Carlo results..."
    Call DrawMCMC(iteration, Age)
    Application.StatusBar = "Finished"
    End If
End Sub
