Sub CalculateCAM()
On Error Resume Next
    Dim DistributionIsLogNormal As Boolean
    Dim OverDispersionTerm As Single
    Dim WeightedSum As Single
    Dim MeanValue As Single
    Dim StandardDeviation As Single
    Dim DataIsValid As Boolean
    Dim Value As Single
    Dim Error As Single
    Dim n, nTotal As Integer
    Dim TempArray() As Single
    Dim i As Integer
    Dim MaxRowNumber As Integer
    MaxRowNumber = Application.WorksheetFunction.Max(Sheet1.Range("A3:A10003")) + 2
    
    n = Sheet1.Cells(22, 29).Value2
    nTotal = ValidXValues
    TempArray = GetData(n)
    
    ReDim zi(1 To n)
    ReDim se(1 To n)
    For i = 1 To n
        zi(i) = TempArray(i, 1)
        se(i) = TempArray(i, 2)
    Next i
   
    'update buttons and labels
    With Sheet1
        .OverdispersionCheckBox.Enabled = True
        If (Sheet1.NormalStatisticsButton.Value = True) Then
            DistributionIsLogNormal = False
            If (.OverdispersionCheckBox.Value = True) Then
                .ModelLabel.Caption = "Model Selected:  CAM (unlog)"
            Else
                .ModelLabel.Caption = "Model Selected:  Common Age (unlog)"
            End If
        Else
            DistributionIsLogNormal = True
            If (.OverdispersionCheckBox.Value = True) Then
                .ModelLabel.Caption = "Model Selected:  CAM (log)"
            Else
                .ModelLabel.Caption = "Model Selected:  Common Age (log)"
            End If
        End If
    
        '  determine statistical values
        If (.OverdispersionCheckBox.Value = True) Then
            OverDispersionTerm = OverDispersionError(DistributionIsLogNormal, zi, se)
            MeanValue = SumOfWeightDelta(n, OverDispersionTerm, zi, se) / SumOfWeight(n, OverDispersionTerm, zi, se)
            If Sheet1.LogPlotCheck.Value = True Then
                Call LogProfile(n, OverDispersionTerm, MeanValue, DistributionIsLogNormal, zi, se)
            End If
        Else
            OverDispersionTerm = 0
        End If
    End With
    
    WeightedSum = SumOfWeight(n, OverDispersionTerm, zi, se)
    MeanValue = SumOfWeightDelta(n, OverDispersionTerm, zi, se) / WeightedSum
    
    If n > 0 Then
        StandardError = (WeightedSum) ^ (-0.5)
        DataIsValid = False
        If (DistributionIsLogNormal = True) Then
            Value = Exp(MeanValue)
            If (MeanValue <> 0) Then
                Error = Abs(Value * StandardError)
                DataIsValid = True
            Else
                MsgBox ("Distribution cannot be log-normal: Mean Value = 0")
                Sheet1.NormalStatisticsButton.Value = True
            End If
        Else
            Value = MeanValue
            If (StandardError > 0) Then
                Error = StandardError
                DataIsValid = True
            End If
        End If
        StandardDeviation = Error * Sqr(n)
    Else
        StandardError = Application.WorksheetFunction.StDev_S(Sheet1.Range(Sheet1.Cells(3, 2), Sheet1.Cells(MaxRowNumber, 2))) / Sqr(nTotal)
        DataIsValid = False
        If (DistributionIsLogNormal = True) Then
            Sheet1.NormalStatisticsButton.Value = True
            Value = Application.WorksheetFunction.Average(Sheet1.Range(Sheet1.Range(Sheet1.Cells(3, 2), Sheet1.Cells(MaxRowNumber, 2))))
            Error = StandardError
            DataIsValid = True
        Else
            Value = MeanValue
            Error = StandardError
            DataIsValid = True
        End If
        StandardDeviation = Error * Sqr(nTotal)
    End If
    
    If (DataIsValid = True) Then
        Sheet1.Cells(27, 5).Value = Value
        If Sheet1.MeanErrorButton.Value = True Then
            Sheet1.Cells(27, 7).Value = Error
        Else
            Sheet1.Cells(27, 7).Value = StandardDeviation
        End If
        
        If (Sheet1.OverdispersionCheckBox.Value = True) Then
            If Round(OverDispersionTerm, 4) = 0 Then
                Sheet1.Range("E24:H24").Value = "Overdispersion: 0"
                If DistributionIsLogNormal = True Then
                  Sheet1.Cells(28, 29) = OverDispersionTerm * 100 & " %"
                Else
                  Sheet1.Cells(28, 29) = OverDispersionTerm & " s"
                End If
            Else
                If Sheet1.LogPlotCheck.Value = True Then
                    StDevOfOD = Sheet10.Cells(9, 43).Value2
                Else
                    StDevOfOD = 1 / Sqr(2 * OverDispersionTerm ^ 2 * SumOfSquareWeight(n, OverDispersionTerm, zi, se))
                End If
                    
                If DistributionIsLogNormal = True Then
                  Sheet1.Range("E24:H24").Value = "Overdispersion: " + CStr(Round(OverDispersionTerm * 100, 0)) + " " + ChrW(177) + " " + CStr(Round(StDevOfOD * 100, 0)) + " %"
                  Sheet1.Cells(28, 29) = OverDispersionTerm * 100 & " %"
                  Sheet1.Cells(30, 29) = StDevOfOD * 100 & " %"
                Else
                  Sheet1.Range("E24:H24").Value = "Overdispersion: " + CStr(Round(OverDispersionTerm / MeanValue * 100, 0)) + " " + ChrW(177) + " " + CStr(Round(OverDispersionTerm / MeanValue * Sqr((StDevOfOD / OverDispersionTerm) ^ 2 + (Error / MeanValue) ^ 2) * 100, 0)) + " %"
                  Sheet1.Cells(28, 29) = OverDispersionTerm & " s"
                  Sheet1.Cells(30, 29) = StDevOfOD & " s"
                End If
            End If
        End If
    End If
    
    If Sheet1.NormalStatisticsButton.Value = True Then
            If Sheet1.OverdispersionCheckBox.Value = True Then
                Sheet1.Cells(27, 8) = "CAM-ul"
            Else
                Sheet1.Cells(27, 8) = "COM-ul"
            End If
    Else
            If Sheet1.OverdispersionCheckBox.Value = True Then
                Sheet1.Cells(27, 8) = "CAM"
            Else
                Sheet1.Cells(27, 8) = "COM"
            End If
    End If
    
    Call Homotest(n, zi, se)
    Call Calculateskewness(n, DistributionIsLogNormal, zi, se)
End Sub