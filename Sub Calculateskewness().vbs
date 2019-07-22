Sub Calculateskewness(ByVal NumberofData As Integer, ByVal DistributionIsLogNormal As Double, zi(), se())
 Application.ScreenUpdating = False
    Dim MeanValue, SumOfWi, Wi, Value, DeDiff As Double
    Dim sk, sigmask As Single
    Dim Row As Integer
    
    If NumberofData > 1 Then
        MeanValue = Sheet1.Cells(27, 5).Value
        SD = Application.WorksheetFunction.StDev_S(zi)
        If DistributionIsLogNormal = True Then
            MeanValue = Log(MeanValue)
        End If
    
    Row = 2
    SumOfWi = 0
    Value = 0
    For i = 1 To NumberofData
        If DistributionIsLogNormal = True Then
            Wi = 1 / Abs(se(i))
        Else
            Wi = Abs(zi(i) / se(i))
        End If
        DeDiff = zi(i) - MeanValue
        Value = Value + Wi * (DeDiff / SD) ^ 3
        SumOfWi = SumOfWi + Wi
    Next i
    
    sk = Value * (1 / SumOfWi)
    sigmask = Sqr(6 / NumberofData)
    Sheet1.Cells(26, 31) = ChrW(177) & CStr(Round(sigmask, 2))
    If sk >= sigmask Then
    Sheet1.Cells(26, 29) = CStr(Round(sk, 2)) + " | Positive"
    ElseIf sk <= -sigmask Then
    Sheet1.Cells(26, 29) = CStr(Round(sk, 2)) + " | Negative"
    Else
    Sheet1.Cells(26, 29) = CStr(Round(sk, 2)) + " | Not significant"
    End If
    End If
End Sub