Sub DrawKDEFigure()
On Error Resume Next
    Dim Row, i, j, k As Integer
    Dim Kernel As Double
    Dim SumOfKernel As Double
    Dim A, Bw As Single
    Dim NumberofData, nTotal As Integer
    Dim KDE(1 To 1000, 1 To 2) As Double
    Dim InterQuartileRange, SD, MaxError, MaxValue, MinValue As Single
    Dim KDEPlotFrame As Shape
    Dim KDEPlotIndex As Integer
    Dim MaxLik As Single
    Dim MaxLikRow As Integer
    Dim MinimumX, MaximumX, MaximumY, MeanValue As Single
    Dim MaxRow, MinRow As Integer
    Dim MaximumNumber As Integer
    Dim KDETemp(1 To 1000) As Single
    Dim HorizontalIndent, VerticalIndent, HalfBoxSize, BoxSize As Single
    Dim XIndent, YIndent, XAxisLocation, XGraphScale, YAxisLocation, YGraphScale As Single
    Dim CenterLine As Shape
    Dim XAxis, YAxis As Shape
    Dim Step As Single
    Dim XAxisTickMarkLength, YAxisTickMarkLength As Single
    Dim Counter, NumberOfXAxisTickMarks, NumberOfYAxisTickMarks, NumberOfXAxisLabels, NumberOfYAxisLabels As Integer
    Dim XAxisLabelScalingFactor, YAxisLabelScalingFactor, Indent As Single
    Dim BetaStrength As Single
    Dim la, apptext As Single
    Dim applabel As String
    Dim Xi As Single
    Dim L As Single
    Dim XValue, YValue As Single
    Dim OvalSize, HalfOvalSize, De, LowerLimit, UpperLimit, LowerXValue, UpperXValue As Single
    Dim DeErrorBar, DeData As Shape
    Dim TempArray() As Single
    Dim LeftValue, RightValue As Single
    Dim StandardErrorArea As Shape
    Dim MaxRowNumber As Integer
    MaxRowNumber = Application.WorksheetFunction.Max(Sheet1.Range("A3:A10003")) + 2

    NumberofData = Sheet1.Cells(22, 29).Value2
    nTotal = Application.WorksheetFunction.Count(Sheet10.Range(Sheet10.Cells(2, 7), Sheet10.Cells(102, 7)))
    If NumberofData >= 2 Then
        ReDim TempArray(1 To NumberofData, 1 To 2)
        TempArray = GetData(NumberofData)
        ReDim Dose(1 To NumberofData) As Single
        ReDim se(1 To NumberofData) As Single
        For i = 1 To NumberofData
          Dose(i) = TempArray(i, 1)
          se(i) = TempArray(i, 2)
        Next i
    ElseIf NumberofData < 2 And nTotal >= 2 Then
        ReDim Dose(1 To nTotal) As Single
        ReDim se(1 To nTotal) As Single
        For i = 1 To nTotal
        If Sheet1.LogNormalStatisticsButton.Value = True And Sheet10.Cells(i + 1, 7).Value2 > 0 Then
            Dose(i) = Log(Sheet10.Cells(i + 1, 7).Value2)
            se(i) = 0
        ElseIf Sheet1.NormalStatisticsButton.Value = True Then
            Dose(i) = Sheet10.Cells(i + 1, 7).Value2
            se(i) = 0
        ElseIf Sheet1.LogNormalStatisticsButton.Value = True And Sheet10.Cells(i + 1, 7).Value2 <= 0 Then
            MsgBox "Non-calculable:  Inclusion of a non-positive De value violates the basic premise of the logarithm"
            Sheet1.KDE.Value = False
            Sheet1.NormalStatisticsButton.Value = True
            Exit Sub
        End If
        Next i
        NumberofData = nTotal
    End If
    
    If Sheet1.NormalStatisticsButton.Value = True Then
        MaxError = Application.WorksheetFunction.Max(se)
        MaxValue = Abs(Dose(NumberofData) + MaxError)
        MinValue = Dose(1) - MaxError
    Else
        ReDim z1(1 To NumberofData) As Single
        For i = 1 To NumberofData
            z1(i) = Exp(Dose(i)) * se(i)
        Next i
        MaxError = Application.WorksheetFunction.Max(z1)
        MaxValue = Abs(Exp(Dose(NumberofData)) + MaxError)
        MinValue = Exp(Dose(1)) - MaxError
    End If
    
    If MinValue > 0 Then
        MinValue = 0.5 * MinValue
    Else
        MinValue = MinValue * 2
    End If

    If Sheet1.NormalStatisticsButton.Value = True Then
        MaxValue = MaxValue * 1.5
        MinValue = MinValue
    Else
        MaxValue = Log(MaxValue) * 1.5
        If MinValue = 0 Then
        MinValue = 0
        ElseIf MinValue > -1 Then
        MinValue = Application.WorksheetFunction.Min(Log(Abs(MinValue)), MinValue)
        ElseIf MinValue <= -1 Then
        MinValue = Application.WorksheetFunction.Min(-Log(-MinValue), MinValue)
        End If
    End If

    For i = 1 To 1000
        KDE(i, 1) = MinValue + (i - 1) * (MaxValue - MinValue) / 999
    Next i
 ' -----------------------------------------------------Adaptive KDE----------------------------------------------------------------------------
If Sheet1.Cells(22, 21).Value = "Adaptive" Or Sheet1.Cells(22, 21).Value = "User-Adaptive" Then
    ReDim KDE1(1 To NumberofData) As Single
    ReDim adaptiveh(1 To NumberofData) As Single
    Dim sumlogKDE, G As Single
    If Sheet1.Cells(22, 21).Value = "Adaptive" Then
            InterQuartileRange = Application.WorksheetFunction.Quartile(Dose, 3) - Application.WorksheetFunction.Quartile(Dose, 1)
            SD = Application.WorksheetFunction.StDev_S(Dose)
            A = Application.WorksheetFunction.Min(InterQuartileRange / 1.34, SD)
            Bw = 0.9 * A * NumberofData ^ (-1 / 5)
    ElseIf Sheet1.Cells(22, 21).Value = "User-Adaptive" Then
            If IsNumeric(Sheet1.Cells(22, 24)) And Sheet1.Cells(22, 24).Value > 0 Then
            Bw = Sheet1.Cells(22, 24).Value2
            Else
            MsgBox "Please enter a numeric bandwidth value in cell on the right"
            Exit Sub
            End If
    End If
        
        sumlogKDE = 0
        For i = 1 To NumberofData
                SumOfKernel = 0
                For k = 1 To NumberofData
                        Kernel = 1 / Sqr(2 * Pi) * Exp(-0.5 * ((Dose(i) - Dose(k)) / Bw) ^ 2)
                        SumOfKernel = SumOfKernel + Kernel
                Next k
                KDE1(i) = SumOfKernel / (Bw * NumberofData)
                sumlogKDE = sumlogKDE + Log(KDE1(i))
        Next i
        
        G = Exp(sumlogKDE / NumberofData)
        
        For i = 1 To NumberofData
            adaptiveh(i) = Bw * Sqr(G / KDE1(i))
        Next i
        
        Sheet1.Cells(22, 24).Value = CStr(Format(Application.WorksheetFunction.Min(adaptiveh), "###0.0")) & "-" & CStr(Format(Application.WorksheetFunction.Max(adaptiveh), "###0.0"))
        
        For j = 1 To 1000
            SumOfKernel = 0
            For k = 1 To NumberofData
                Kernel = (1 / Sqr(2 * Pi) * Exp(-0.5 * ((KDE(j, 1) - Dose(k)) / adaptiveh(k)) ^ 2)) / (adaptiveh(k) * NumberofData)
                SumOfKernel = SumOfKernel + Kernel
            Next k
            KDE(j, 2) = SumOfKernel
        Next j
ElseIf Sheet1.Cells(22, 21).Value = "PDF Plot" Then

        If Sheet1.Cells(22, 29).Value2 >= 2 Then
                For j = 1 To 1000
                    SumOfKernel = 0
                    For k = 1 To NumberofData
                        Kernel = (1 / Sqr(2 * Pi) * Exp(-0.5 * ((KDE(j, 1) - Dose(k)) / se(k)) ^ 2)) / (se(k) * NumberofData)
                        SumOfKernel = SumOfKernel + Kernel
                    Next k
                    KDE(j, 2) = SumOfKernel
                Next j
                Sheet1.Cells(22, 24).Value2 = "De-Error"
        Else
            MsgBox "PDF cannot be created when the valid data points < 2. Try other KDE plots"
            Exit Sub
        End If
Else
'------------------------------------------------------------------------------fix bandwidth KDE-------------------------------------------------------------------------
    'calculate the approporate bandwidth following bw.nrd0 method in R (Silverman, 1998, Density Estmation, p48, equation (3.31))
        If Sheet1.Cells(22, 21).Value = "User-Defined" Then
            If IsNumeric(Sheet1.Cells(22, 24).Value) And Sheet1.Cells(22, 24).Value > 0 Then
                    Bw = Sheet1.Cells(22, 24).Value2
            Else
                    MsgBox "The bandwidth value is not valid"
                    Exit Sub
            End If
        ElseIf Sheet1.Cells(22, 21).Value = "Silverman331" Then
            InterQuartileRange = Application.WorksheetFunction.Quartile(Dose, 3) - Application.WorksheetFunction.Quartile(Dose, 1)
            SD = Application.WorksheetFunction.StDev_S(Dose)
            A = Application.WorksheetFunction.Min(InterQuartileRange / 1.34, SD)
            Bw = 0.9 * A * NumberofData ^ (-1 / 5)
        ElseIf Sheet1.Cells(22, 21).Value = "Scott1992" Then
            InterQuartileRange = Application.WorksheetFunction.Quartile(Dose, 3) - Application.WorksheetFunction.Quartile(Dose, 1)
            SD = Application.WorksheetFunction.StDev_S(Dose)
            A = Application.WorksheetFunction.Min(InterQuartileRange / 1.34, SD)
            Bw = 1.06 * A * NumberofData ^ (-1 / 5)
        End If
        
        Sheet1.Cells(22, 24).Value2 = Bw

        For j = 1 To 1000
            SumOfKernel = 0
            For k = 1 To NumberofData
            Kernel = 1 / Sqr(2 * Pi) * Exp(-0.5 * ((KDE(j, 1) - Dose(k)) / Bw) ^ 2)
            SumOfKernel = SumOfKernel + Kernel
            Next k
            KDE(j, 2) = SumOfKernel / (Bw * NumberofData)
        Next j
        
  End If
  
'================================================================Draw Figure===============================================
    Sheet1.Cells(24, 21).Select
    If Sheet1.NormalStatisticsButton.Value = True Then
        If Sheet1.Cells(22, 21) = "PDF Plot" Then
        Sheet1.Cells(26, 19) = "Normal Probability Density Functions Plot"
        Else
        Sheet1.Cells(26, 19) = "Normal Kernel Density Plot"
        End If
    Else
        If Sheet1.Cells(22, 21) = "PDF Plot" Then
        Sheet1.Cells(26, 19) = "Log-Normal Proability Density Functions Plot"
        Else
        Sheet1.Cells(26, 19) = "Log-Normal Kernel Density Plot"
        End If
    End If
    
    For i = 1 To 1000
        KDETemp(i) = KDE(i, 2)
    Next i
    
    MaxLik = Application.WorksheetFunction.Max(KDETemp)
    MaxLikRow = Application.WorksheetFunction.Match(MaxLik, KDETemp, 0)
    For i = MaxLikRow To 1000
        ReDim Temp(i To 1000) As Single
        For k = i To 1000
          Temp(k) = KDETemp(k)
        Next k
        
        If (Application.WorksheetFunction.Max(Temp) < 1 * 10 ^ (-13)) Then
            MaximumX = KDE(i, 1)
            MaxRow = i
            Exit For
        Else
            MaximumX = KDE(1000, 1)
            MaxRow = 1000
        End If
    Next i
    
    For j = 1 To MaxLikRow
        If KDETemp(j) > 1 * 10 ^ (-13) Then
            MinimumX = KDE(j, 1)
            MinRow = j
            Exit For
        End If
    Next j
    
    MaximumNumber = Round((MaxRow - MinRow) / 3 - 0.5, 0) * 3 + 1
    MaximumY = MaxLik
    
        HorizontalIndent = Sheet1.Range("S4").Left + 1.75
        VerticalIndent = Sheet1.Range("S4").Top + 4.75
        HalfBoxSize = 113
        BoxSize = 2 * HalfBoxSize
        

        
        Set KDEPlotFrame = Sheet1.Shapes.AddShape(msoShapeRectangle, HorizontalIndent, VerticalIndent - 0.5, BoxSize + 35, BoxSize - 4.5)
        With KDEPlotFrame
        .Fill.ForeColor.RGB = RGB(243, 242, 233)
        .Fill.Visible = msoFalse
        .Line.DashStyle = msoLineSolid
        .Line.ForeColor.RGB = RGB(0, 0, 0)
        .Line.Visible = msoFalse
        .ZOrder msoBringToFront
        End With
        KDEPlotIndex = Sheet1.Shapes.Count - 1
        
        'define axis locations
        XIndent = 50
        YIndent = 50
        XAxisLocation = VerticalIndent + BoxSize - YIndent
        XGraphScale = 190 / (MaximumX - MinimumX)
        YAxisLocation = HorizontalIndent + XIndent
        YGraphScale = 140 / MaximumY
        
        If IsEmpty(Sheet1.Cells(21, 5)) = False Then
            If Sheet1.NormalStatisticsButton = True Then
                MeanValue = Sheet1.Cells(21, 5) - MinimumX
                LeftValue = Sheet1.Cells(21, 5) - Sheet1.Cells(21, 7) - MinimumX
                RightValue = Sheet1.Cells(21, 5) + Sheet1.Cells(21, 7) - MinimumX
            ElseIf Sheet1.LogNormalStatisticsButton = True Then
                MeanValue = Log(Sheet1.Cells(21, 5)) - MinimumX
                LeftValue = Log(Sheet1.Cells(21, 5) - Sheet1.Cells(21, 7)) - MinimumX
                RightValue = Log(Sheet1.Cells(21, 5) + Sheet1.Cells(21, 7)) - MinimumX
            End If
        'standard error
        Set StandardErrorArea = Sheet1.Shapes.AddShape(msoShapeRectangle, YAxisLocation + LeftValue * XGraphScale + 5, XAxisLocation - 140, (RightValue - LeftValue) * XGraphScale, 145)
            With StandardErrorArea
                .Fill.ForeColor.RGB = RGB(255, 0, 0)
                .Fill.Visible = msoTrue
                .Fill.Transparency = 0.88
                .Line.DashStyle = msoLineSolid
                .Line.ForeColor.RGB = RGB(0, 0, 0)
                .Line.Visible = msoFalse
                .ZOrder msoBringToFront
            End With
        Sheet1.Shapes.Range(Array(KDEPlotIndex, KDEPlotIndex + 1)).Group
        
        'plot center line
        Set CenterLine = Sheet1.Shapes.AddLine(YAxisLocation + MeanValue * XGraphScale + 5, XAxisLocation + 5, YAxisLocation + MeanValue * XGraphScale + 5, XAxisLocation - 140)
            With CenterLine.Line
                .DashStyle = msoLineSolid
                .Weight = 0.6
                .ForeColor.RGB = RGB(255, 0, 0)
                .Visible = msoTrue
            End With
        CenterLine.ZOrder msoBringToFront
        Sheet1.Shapes.Range(Array(KDEPlotIndex, KDEPlotIndex + 1)).Group
        
        End If
        'draw x axis
        Set XAxis = Sheet1.Shapes.AddLine(YAxisLocation, XAxisLocation + 8, YAxisLocation + 200, XAxisLocation + 8)
        With XAxis.Line
        .DashStyle = msoLineSolid
        .Weight = 1.2
        .ForeColor.RGB = RGB(0, 0, 0)
        .Visible = msoTrue
        End With
        XAxis.ZOrder msoBringToFront
        Sheet1.Shapes.Range(Array(KDEPlotIndex, KDEPlotIndex + 1)).Group
     
        Step = (MaximumX - MinimumX) / 6
        
        XAxisTickMarkLength = 4
        NumberOfXAxisTickMarks = 6
        ReDim XAxisTickMark(NumberOfXAxisTickMarks + 1) As Shape
        For Counter = 0 To NumberOfXAxisTickMarks
            Set XAxisTickMark(Counter) = Sheet1.Shapes.AddLine(YAxisLocation + Counter * Step * XGraphScale + 5, XAxisLocation + 8, YAxisLocation + Counter * Step * XGraphScale + 5, XAxisLocation + 8 + XAxisTickMarkLength)
            With XAxisTickMark(Counter).Line
            .Weight = 1
            .DashStyle = msoLineSolid
            .ForeColor.RGB = RGB(0, 0, 0)
            .Visible = msoTrue
            End With
            XAxisTickMark(Counter).ZOrder msoBringToFront
            Sheet1.Shapes.Range(Array(KDEPlotIndex, KDEPlotIndex + 1)).Group
        Next Counter
    
        'X label
        XAxisLabelScalingFactor = 0.1
        ReDim XAxisLabel(NumberOfXAxisTickMarks + 2) As Shape
        For Counter = 0 To NumberOfXAxisTickMarks
           
            Indent = Len(CStr(Round(Counter * Step, 0)))
            Set XAxisLabel(Counter) = Sheet1.Shapes.AddLabel(msoTextOrientationHorizontal, YAxisLocation + Counter * Step * XGraphScale - 5 * Indent * XAxisLabelScalingFactor, XAxisLocation + XAxisTickMarkLength + 5, 10 * Indent * XAxisLabelScalingFactor, 10)
        With XAxisLabel(Counter)
            .TextFrame.AutoSize = True
            .Line.Visible = msoFalse
            .Fill.Transparency = 1
        With .TextFrame
            .VerticalAlignment = xlVAlignCenter
            .HorizontalAlignment = xlHAlignCenter
            .Characters.Font.Size = 10.5
            .Characters.Font.Name = "Arial"
          
        BetaStrength = Sheet4.Cells(41, 5).Value / 60 'Gy/min to Gy/s
        If Sheet1.NormalStatisticsButton = True Then
            If Cells(24, 21) = "Time" Then
               If Counter * Step + MinimumX < 10 Then
               .Characters.Text = CStr(Format(Counter * Step + MinimumX, "###0.00"))
                Else
                .Characters.Text = CStr(Format(Counter * Step + MinimumX, "###0.0"))
               End If
            Else
                If (Counter * Step + MinimumX) * BetaStrength < 10 Then
                .Characters.Text = CStr(Format((Counter * Step + MinimumX) * BetaStrength, "###0.00"))
                Else
                .Characters.Text = CStr(Format((Counter * Step + MinimumX) * BetaStrength, "###0.0"))
                End If
            End If
        Else
            If Cells(24, 21) = "Time" Then
                If Exp(Step * Counter + MinimumX) < 10 Then
                    .Characters.Text = CStr(Format(Exp(Counter * Step + MinimumX), "###0.00"))
                Else
                    .Characters.Text = CStr(Format(Exp(Counter * Step + MinimumX), "###0.0"))
                End If
            Else
                la = Exp(Step * Counter + MinimumX) * BetaStrength
                If la < 10 Then
                .Characters.Text = CStr(Format(Exp(Step * Counter + MinimumX) * BetaStrength, "###0.00"))
                Else
                .Characters.Text = CStr(Format(Exp(Step * Counter + MinimumX) * BetaStrength, "###0.0"))
                End If
            End If
        End If
        End With
            .ZOrder msoBringToFront
        End With
        Sheet1.Shapes.Range(Array(KDEPlotIndex, KDEPlotIndex + 1)).Group
        Next Counter
                
        Set XAxisLabel(NumberOfXAxisTickMarks + 1) = Sheet1.Shapes.AddLabel(msoTextOrientationHorizontal, YAxisLocation + (MaximumX - MinimumX) * XGraphScale / 2, XAxisLocation + XAxisTickMarkLength + 20, 75, 10)
        With XAxisLabel(NumberOfXAxisTickMarks + 1)
        .TextFrame.AutoSize = True
        .Line.Visible = msoFalse
        .Fill.Transparency = 1
        With .TextFrame
        .VerticalAlignment = xlVAlignCenter
        .HorizontalAlignment = xlHAlignCenter
        .Characters.Font.Name = "Arial"
        .Characters.Font.Size = 11
        If Cells(24, 21) = "Time" Then
            .Characters.Text = "Exposure time (sec)"
        Else
           .Characters.Text = "Equivalent dose (Gy)"
        End If
        End With
        .ZOrder msoBringToFront
        End With
        Sheet1.Shapes.Range(Array(KDEPlotIndex, KDEPlotIndex + 1)).Group
        
    
    'draw y axis
        Set YAxis = Sheet1.Shapes.AddLine(YAxisLocation, XAxisLocation + 8, YAxisLocation, XAxisLocation - 150)
        With YAxis.Line
        .DashStyle = msoLineSolid
        .Weight = 1.2
        .ForeColor.RGB = RGB(0, 0, 0)
        .Visible = msoTrue
        End With
        YAxis.ZOrder msoBringToFront
        Sheet1.Shapes.Range(Array(KDEPlotIndex, KDEPlotIndex + 1)).Group
    
        Step = MaximumY / 6
        
        YAxisTickMarkLength = 4
        NumberOfYAxisTickMarks = Round(MaximumY / Step - 0.5, 0)
        ReDim YAxisTickMark(NumberOfYAxisTickMarks + 1) As Shape
        For Counter = 0 To NumberOfYAxisTickMarks
            Set YAxisTickMark(Counter) = Sheet1.Shapes.AddLine(YAxisLocation, XAxisLocation - Counter * Step * YGraphScale, YAxisLocation - YAxisTickMarkLength, XAxisLocation - Counter * Step * YGraphScale)
            With YAxisTickMark(Counter).Line
            .Weight = 1
            .DashStyle = msoLineSolid
            .ForeColor.RGB = RGB(0, 0, 0)
            .Visible = msoTrue
            End With
            YAxisTickMark(Counter).ZOrder msoBringToFront
            Sheet1.Shapes.Range(Array(KDEPlotIndex, KDEPlotIndex + 1)).Group
        Next Counter
        
        If Sheet1.Cells(22, 21).Value = "PDF Plot" Then
            If (Step < 0.000001) Then
                apptext = Round(Step * 1000000, 2)
                applabel = "Probability (10-6)"
            ElseIf (Step < 0.00001) Then
                apptext = Round(Step * 100000, 2)
                applabel = "Probability (10-5)"
            ElseIf (Step < 0.0001) Then
                apptext = Round(Step * 10000, 2)
                applabel = "Probability (10-4)"
            ElseIf (Step < 0.001) Then
                apptext = Round(Step * 1000, 2)
                applabel = "Probability (10-3)"
            ElseIf (Step < 0.01) Then
                apptext = Round(Step * 100, 2)
                applabel = "Probability (10-2)"
            ElseIf (Step < 0.1) Then
                apptext = Round(Step * 10, 2)
                applabel = "Probability (10-1)"
            Else
                apptext = Round(Step, 2)
                applabel = "Probability"
            End If
        Else
            If (Step < 0.000001) Then
                apptext = Round(Step * 1000000, 2)
                applabel = "Kernel Density (10-6)"
            ElseIf (Step < 0.00001) Then
                apptext = Round(Step * 100000, 2)
                applabel = "Kernel Density (10-5)"
            ElseIf (Step < 0.0001) Then
                apptext = Round(Step * 10000, 2)
                applabel = "Kernel Density (10-4)"
            ElseIf (Step < 0.001) Then
                apptext = Round(Step * 1000, 2)
                applabel = "Kernel Density (10-3)"
            ElseIf (Step < 0.01) Then
                apptext = Round(Step * 100, 2)
                applabel = "Kernel Density (10-2)"
            ElseIf (Step < 0.1) Then
                apptext = Round(Step * 10, 2)
                applabel = "Kernel Density (10-1)"
            Else
                apptext = Round(Step, 2)
                applabel = "Kernel Density"
            End If
        End If
        YAxisLabelScalingFactor = 0.3
        NumberOfYAxisLabels = NumberOfYAxisTickMarks
        ReDim YAxisLabel(NumberOfYAxisLabels + 2) As Shape
        For Counter = 0 To NumberOfYAxisLabels
            Set YAxisLabel(Counter) = Sheet1.Shapes.AddLabel(msoTextOrientationHorizontal, _
             HorizontalIndent + XIndent - YAxisTickMarkLength - Len(CStr(apptext)) - 6, _
             XAxisLocation - (Counter * Step * YGraphScale) - 10, 5, 10)
             With YAxisLabel(Counter)
             .TextFrame.AutoSize = True
             .Line.Visible = msoFalse
             .Fill.Transparency = 1
             With .TextFrame
             .VerticalAlignment = xlVAlignCenter
             .HorizontalAlignment = xlHAlignRight
             .Characters.Font.Size = 10.5
             .Characters.Font.Name = "Arial"
             .Characters.Text = CStr(Format(Counter * apptext, "###0.00"))
             End With
             .ZOrder msoBringToFront
             End With
             Sheet1.Shapes.Range(Array(KDEPlotIndex, KDEPlotIndex + 1)).Group
        Next Counter
        
        Set YAxisLabel(NumberOfYAxisLabels + 1) = Sheet1.Shapes.AddLabel(msoTextOrientationUpward, HorizontalIndent + (XIndent - YAxisTickMarkLength) / 4 - 17, XAxisLocation - (MaximumY * YGraphScale / 2), 10, 10)
            With YAxisLabel(NumberOfYAxisLabels + 1)
            .Line.Visible = msoFalse
            .Fill.Transparency = 1
            With .TextFrame
            .AutoSize = True
            .VerticalAlignment = xlVAlignCenter
            .HorizontalAlignment = xlHAlignCenter
            .Characters.Font.Size = 11
            .Characters.Font.Name = "Arial"
            .Characters.Text = applabel
            If applabel <> "Kernel Density" Then
            .Characters(19, 1).Font.Superscript = True
            .Characters(20, 1).Font.Superscript = True
            End If
            End With
            .ZOrder msoBringToFront
            End With
        Sheet1.Shapes.Range(Array(KDEPlotIndex, KDEPlotIndex + 1)).Group
        
      'plot likelihood curve
        ReDim Pts(1 To MaximumNumber, 1 To 2) As Single
        For Row = MinRow To MaximumNumber + MinRow - 1
            Xi = KDE(Row, 1) - MinimumX
            L = KDE(Row, 2)
            XValue = YAxisLocation + Xi * XGraphScale + 5
            YValue = XAxisLocation - L * YGraphScale
            Pts(Row - MinRow + 1, 1) = XValue
            Pts(Row - MinRow + 1, 2) = YValue
        Next Row

        With Sheet1
        .Shapes.AddCurve SafeArrayOfPoints:=Pts
        .Shapes(KDEPlotIndex + 1).Line.ForeColor.RGB = RGB(Cells(30, 21), Cells(30, 22), Cells(30, 24))
        .Shapes(KDEPlotIndex + 1).Line.Weight = 1
        If KDE(MinRow, 2) < 1 * 10 ^ -8 Then
            .Shapes(KDEPlotIndex + 1).Fill.Transparency = 0.975
            .Shapes(KDEPlotIndex + 1).Fill.ForeColor.RGB = RGB(Cells(30, 21), Cells(30, 22), Cells(30, 24))
        Else
            .Shapes(KDEPlotIndex + 1).Fill.Visible = msoFalse
        End If
        .Shapes(KDEPlotIndex + 1).ZOrder msoBringToFront
        .Shapes.Range(Array(KDEPlotIndex, KDEPlotIndex + 1)).Group
       End With
    
    'plot data points
        HalfOvalSize = 1.5
        OvalSize = 2 * HalfOvalSize
    For Row = 1 To NumberofData
        De = Dose(Row) - MinimumX
        If Sheet1.NormalStatisticsButton.Value = True Then
            LowerLimit = De - se(Row)
            UpperLimit = De + se(Row)
        Else
            LowerLimit = Exp(Dose(Row)) * (1 - se(Row))
            If LowerLimit <= 0 Then
            LowerLimit = 0
            Else
            LowerLimit = Log(LowerLimit) - MinimumX
            End If
            UpperLimit = Log(Exp(Dose(Row)) * (1 + se(Row))) - MinimumX
        End If
        
        XValue = YAxisLocation + De * XGraphScale + 5
        LowerXValue = YAxisLocation + LowerLimit * XGraphScale + 5
        UpperXValue = YAxisLocation + UpperLimit * XGraphScale + 5
        YValue = XAxisLocation - Row * (140 / NumberofData)
        
        If se(Row) <> 0 Then
            Set DeErrorBar = Sheet1.Shapes.AddLine(LowerXValue, YValue + (140 / NumberofData) / 2, UpperXValue, YValue + (140 / NumberofData) / 2)
                With DeErrorBar.Line
                .Weight = 0.8
                .DashStyle = msoLineSolid
                .ForeColor.RGB = RGB(Cells(28, 21), Cells(28, 22), Cells(28, 24))
                .Transparency = 0.3
                .Visible = msoTrue
                End With
            DeErrorBar.ZOrder msoBringToFront
            Sheet1.Shapes.Range(Array(KDEPlotIndex, KDEPlotIndex + 1)).Group
            
            Set DeData = Sheet1.Shapes.AddShape(msoShapeOval, XValue - HalfOvalSize, YValue - HalfOvalSize + (140 / NumberofData) / 2, OvalSize, OvalSize)
                With DeData
                .Fill.Visible = msoTrue
                .Fill.ForeColor.RGB = RGB(Cells(28, 21), Cells(28, 22), Cells(28, 24))
                .Fill.Transparency = 0.3
                .Line.Visible = msoFalse
                .ZOrder msoBringToFront
                End With
            Sheet1.Shapes.Range(Array(KDEPlotIndex, KDEPlotIndex + 1)).Group
        End If
        
        'plot data at bottom
        Set DeData = Sheet1.Shapes.AddLine(XValue, XAxisLocation, XValue, XAxisLocation + 5)
            With DeData
            .Line.ForeColor.RGB = RGB(50, 50, 50)
            .Line.Transparency = 0.5
            .Line.Visible = msoTrue
            .Line.Weight = 1
            .ZOrder msoBringToFront
            End With
        Sheet1.Shapes.Range(Array(KDEPlotIndex, KDEPlotIndex + 1)).Group
    Next Row
    Sheet1.Shapes(KDEPlotIndex).Name = "KDEPlot"
    Sheet1.Shapes("KDEPlot").Locked = False
End Sub