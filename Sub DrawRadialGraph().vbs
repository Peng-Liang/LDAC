Public Sub DrawRadialGraph(ByVal PresentDistributionStandardDeviation As Boolean, ByVal n As Integer)
On Error Resume Next
    Dim WeightedSum As Double
    Dim MeanValue As Double
    Dim ErrorInTheMeanValue As Double
    Dim StandardDeviation As Double
    Dim MaximumXDeviation As Double
    Dim MaximumAngle As Double
    Dim MinimumAngle As Double
    
    Dim HorizontalIndent As Single
    Dim VerticalIndent As Single
    Dim HalfBoxSize As Single
    Dim BoxSize As Single
    Dim RadialPlotFrame As Shape
    Dim RadialPlotIndex As Integer
    
    Dim XIndent As Single
    Dim YIndent As Single
    Dim VerticalCenterLine As Single
    Dim XAxisLocation As Single
    Dim XGraphScale As Single
    Dim ScalingFactor As Single
    Dim YAxisLocation As Single
    Dim YGraphScale As Single
    Dim Radius As Double
    Dim ZAxisLocation As Single
    
    Dim XAxis As Shape
    Dim DecimalValue As Double
    Dim TestValue As Double
    Dim BinaryValue As Double
    Dim Step As Single
    Dim XAxisTickMarkLength As Single
    Dim NumberOfXAxisTickMarks As Integer
    Dim Counter As Integer
    Dim XAxisTickMark() As Shape
    Dim XAxisLabel() As Shape
    Dim XAxisLabelScalingFactor As Single
    Dim Indent As Single
    Dim XAxisOtherTickMark() As Shape
    Dim ExcelVersionIs2007 As Boolean
    Dim PrecisionScalingFactor As Single
    Dim XAxisOtherLabel() As Shape
   
    Dim YAxis As Shape
    Dim YAxisTickMarkLength As Single
    Dim NumberOfYAxisTickMarks As Integer
    Dim YAxisTickMark() As Shape
    Dim NumberOfYAxisLabels As Integer
    Dim YAxisLabelScalingFactor As Single
    Dim YAxisLabel() As Shape
    
    Dim ChangeInAngle As Double
    Dim ZAxis As FreeformBuilder
    Dim ZAxisShape As Shape
    Dim Angle As Double
    Dim ZAxisTickMarkLength As Single
    Dim MaximumZValue As Double
    Dim MinimumZValue As Double
    Dim Zvalue As Double
    Dim MinimumValue As Double
    Dim NumberOfZAxisTickMarks As Integer
    Dim LogarithmicZAxisTicks As Boolean
    Dim ZAxisTickMark() As Shape
    Dim PreviousLabelPosition As Single
    Dim NumberOfZAxisLabels As Integer
    Dim ZAxisLabel() As Shape
    Dim ApparentValueScalingFactor As Single
    
    Dim HalfOvalSize As Single
    Dim OvalSize As Single
    Dim ApparentDoseDataPoint() As Shape
    Dim Row As Integer
    Dim Deviation As Double
    Dim Weight As Double
    Dim XValue As Double
    Dim YValue As Double
    
    Dim TwoSigmaLine(2) As Shape
    
    Dim Value As Double
    Dim Error As Double
    Dim DistributionIsLogNormal As Boolean
    Dim OverDispersionTerm As Double
    Dim MaximumYDeviation As Double
    
    Dim OldAngle As Double
    Dim UpperZValue As Double
    Dim LowerZValue As Double
    Dim GrayBar As FreeformBuilder
    Dim UpperAngle As Double
    Dim LowerAngle As Double
    Dim GrayBarShape As Shape
    
    Dim ComponentNumber As Long
    Dim ComponentEDValue As Double
    Dim ComponentEDError As Double
    Dim NumberOfValidDataPoints As Long
    Dim NumberOfComponents As Long
    Dim LabelText As String
    Dim MarkLine1, MarkLine2, MarkLine3 As Shape
    Dim XAxisOtherLabelScalingFactor, ScaledStandardDeviation, StandardizedEstimateScalingFactor, MaximumDose, MinimumDose, SignificantFigures As Double
    Dim MaximumEquivalentDose, MinimumEquivalentDose, EquivalentDose, InitialEquivalentDose, StepMultiplier As Double
    Dim NumberOfNegativeZAxisTickMarks, NumberOfPositiveZAxisTickMarks, Digits As Integer
    Dim BetaStrength As Double
    Dim i As Integer
    Dim TempArray() As Single

    Sheet1.Cells(24, 14).Select
    TempArray = GetData(n)
    ReDim zi(1 To n)
    ReDim se(1 To n)
    For i = 1 To n
        zi(i) = TempArray(i, 1)
        se(i) = TempArray(i, 2)
    Next i
    DistributionIsLogNormal = Not Sheet1.NormalStatisticsButton.Value
    OverDispersionTerm = OverDispersionError(DistributionIsLogNormal, zi, se)
    WeightedSum = SumOfWeight(n, OverDispersionTerm, zi, se)
    Value = SumOfWeightDelta(n, OverDispersionTerm, zi, se) / WeightedSum
    Error = (WeightedSum) ^ (-0.5)
    ErrorInTheMeanValue = Error
    If (DistributionIsLogNormal = True) Then
        ErrorInTheMeanValue = ErrorInTheMeanValue / Value
    End If
    MeanValue = Value
    StandardDeviation = Error * Sqr(ValidXValues)
    OverDispersionTerm = 0
    MaximumXDeviation = GreatestRadialPlotXValue(OverDispersionTerm, n, zi, se)
    MaximumAngle = RadialPlotThetaValue(MeanValue, True, n, zi, se)
    MinimumAngle = RadialPlotThetaValue(MeanValue, False, n, zi, se)

    
    If ((IsEmpty(Sheet1.Cells(22, 14)) = False) And (IsNumeric(Sheet1.Cells(22, 14)) = True) And (Sheet1.Cells(22, 14).Value > 0)) Then
        ScalingFactor = 1 / Sheet1.Cells(22, 14).Value
    Else
        ScalingFactor = 1
    End If

    If (MaximumAngle > 0) Then
        '  draw diagram box
        HorizontalIndent = Sheet1.Range("L4").Left + 1.25
        VerticalIndent = Sheet1.Range("L4").Top - 21.25
        HalfBoxSize = 113
        BoxSize = 2 * HalfBoxSize
        Set RadialPlotFrame = Sheet1.Shapes.AddShape(msoShapeRectangle, HorizontalIndent, VerticalIndent + 25, BoxSize + 35, BoxSize - 5)
        With RadialPlotFrame
        .Fill.ForeColor.RGB = RGB(243, 242, 233)
        .Fill.Visible = msoFalse
        .Line.DashStyle = msoLineSolid
        .Line.ForeColor.RGB = RGB(0, 0, 0)
        .Line.Visible = msoFalse
        .ZOrder msoBringToFront
        End With
        RadialPlotIndex = Sheet1.Shapes.Count - 1

        '  define axis locations
        XIndent = 40
        YIndent = 15
        VerticalCenterLine = VerticalIndent + HalfBoxSize + 8
        XAxisLocation = VerticalIndent + BoxSize - YIndent
        XGraphScale = 250 * BoxSize / MaximumXDeviation / 390
        YAxisLocation = HorizontalIndent + XIndent
        If (Abs(MaximumAngle) > Abs(MinimumAngle)) Then
            YGraphScale = XGraphScale / (2 * Tan(Abs(MaximumAngle)))
        Else
            YGraphScale = XGraphScale / (2 * Tan(Abs(MinimumAngle)))
        End If
        
        If ((YGraphScale / ScalingFactor) > 50) Then
            ScalingFactor = YGraphScale / 50
            YGraphScale = 50
        Else
            YGraphScale = YGraphScale / ScalingFactor
        End If
        
        Radius = Sqr((XAxisLocation - VerticalCenterLine) ^ 2 + (MaximumXDeviation * XGraphScale) ^ 2)
        ZAxisLocation = YAxisLocation + Radius
        
        '  draw x axis
        Set XAxis = Sheet1.Shapes.AddLine(YAxisLocation, XAxisLocation, YAxisLocation + MaximumXDeviation * XGraphScale, XAxisLocation)
        With XAxis.Line
        .DashStyle = msoLineSolid
        .Weight = 1.2
        .ForeColor.RGB = RGB(0, 0, 0)
        .Visible = msoTrue
        End With
        XAxis.ZOrder msoBringToFront
        Sheet1.Shapes.Range(Array(RadialPlotIndex, RadialPlotIndex + 1)).Group

        
        DecimalValue = 1
        TestValue = MaximumXDeviation
        Do While (TestValue < 5)
            TestValue = 10 * TestValue
            DecimalValue = DecimalValue / 10
        Loop
        
        BinaryValue = 1
        Do While (TestValue > 10)
            TestValue = TestValue / 2
            BinaryValue = 2 * BinaryValue
        Loop
        Step = DecimalValue * BinaryValue

        XAxisTickMarkLength = 4
        NumberOfXAxisTickMarks = Round(MaximumXDeviation / Step - 0.5, 0)
        ReDim XAxisTickMark(NumberOfXAxisTickMarks + 1) As Shape
        For Counter = 0 To NumberOfXAxisTickMarks
            Set XAxisTickMark(Counter) = Sheet1.Shapes.AddLine(YAxisLocation + Counter * Step * XGraphScale, XAxisLocation, YAxisLocation + Counter * Step * XGraphScale, XAxisLocation + XAxisTickMarkLength)
            With XAxisTickMark(Counter)
            .Line.Weight = 1
            .Line.DashStyle = msoLineSolid
            .Line.ForeColor.RGB = RGB(0, 0, 0)
            .Line.Visible = msoTrue
            .ZOrder msoBringToFront
            End With
            Sheet1.Shapes.Range(Array(RadialPlotIndex, RadialPlotIndex + 1)).Group
        Next Counter
        
        If (ActiveWorkbook.Application.Version >= 12) Then
            ExcelVersionIs2007 = True
        Else
            ExcelVersionIs2007 = False
        End If

        XAxisLabelScalingFactor = 0.5
        ReDim XAxisLabel(NumberOfXAxisTickMarks + 2) As Shape
        For Counter = 0 To NumberOfXAxisTickMarks
            If (DecimalValue < 1) Then
                Indent = Len(CStr(Round(Counter * Step, Round(-Log(DecimalValue) / Log(10) + 0.1, 0))))
            Else
                Indent = Len(CStr(Round(Counter * Step, 0)))
            End If
            Set XAxisLabel(Counter) = Sheet1.Shapes.AddLabel(msoTextOrientationHorizontal, YAxisLocation + Counter * Step * XGraphScale - 5 * Indent * XAxisLabelScalingFactor - 2, XAxisLocation + XAxisTickMarkLength - 3, 10 * Indent * XAxisLabelScalingFactor, 10)
            
            With XAxisLabel(Counter)
                .Line.Visible = msoFalse
                .Fill.Transparency = 1
                With .TextFrame
                .AutoSize = True
                .VerticalAlignment = xlVAlignCenter
                .HorizontalAlignment = xlHAlignCenter
                .Characters.Font.Size = 10
                .Characters.Font.Name = "Arial"
            
                If (DecimalValue < 1) Then
                    .Characters.Text = CStr(Round(Counter * Step, Round(-Log(DecimalValue) / Log(10) + 0.1, 0)))
                Else
                   .Characters.Text = CStr(Round(Counter * Step, 0))
                End If
                End With
                .ZOrder msoBringToFront
            End With
            Sheet1.Shapes.Range(Array(RadialPlotIndex, RadialPlotIndex + 1)).Group
        Next Counter

        PrecisionScalingFactor = 0
        Set XAxisLabel(NumberOfXAxisTickMarks + 1) = Sheet1.Shapes.AddLabel(msoTextOrientationHorizontal, YAxisLocation + MaximumXDeviation * XGraphScale / 2 - 45 * PrecisionScalingFactor, XAxisLocation + XAxisTickMarkLength + 13, 70, 10)
        With XAxisLabel(NumberOfXAxisTickMarks + 1)
            .Line.Visible = msoFalse
            .Fill.Transparency = 1
            With .TextFrame
                .AutoSize = True
                .VerticalAlignment = xlVAlignCenter
                .HorizontalAlignment = xlHAlignCenter
                .Characters.Font.Name = "Arial"
                .Characters.Font.Size = 11
                .Characters.Text = "Precision"
            End With
            .ZOrder msoBringToFront
        End With
        Sheet1.Shapes.Range(Array(RadialPlotIndex, RadialPlotIndex + 1)).Group
        
        
   'mark the sample information
        Set MarkLine1 = Sheet1.Shapes.AddLabel(msoTextOrientationHorizontal, HorizontalIndent + 15, VerticalIndent + 25, 20, 10)
        With MarkLine1
            .TextFrame.AutoSize = True
            .Line.Visible = msoFalse
            .Fill.Transparency = 1
            .TextFrame.VerticalAlignment = xlVAlignCenter
            .TextFrame.HorizontalAlignment = xlHAlignLeft
            .TextFrame.Characters.Font.Name = "Arial"
            .TextFrame.Characters.Font.Size = 10
            .TextFrame.Characters.Text = CStr(Sheet4.Cells(7, 5).Value2) + " " + "(n = " + CStr(Sheet1.Cells(22, 29).Value2) + ")"
            .ZOrder msoBringToFront
        End With
        Sheet1.Shapes.Range(Array(RadialPlotIndex, RadialPlotIndex + 1)).Group
        
    If (IsEmpty(Sheet1.Cells(24, 5)) = False) Then
        Set MarkLine2 = Sheet1.Shapes.AddLabel(msoTextOrientationHorizontal, HorizontalIndent + 15, VerticalIndent + 40, 20, 10)
        With MarkLine2
            .TextFrame.AutoSize = True
            .Line.Visible = msoFalse
            .Fill.Transparency = 1
            .TextFrame.VerticalAlignment = xlVAlignCenter
            .TextFrame.HorizontalAlignment = xlHAlignLeft
            .TextFrame.Characters.Font.Name = "Arial"
            .TextFrame.Characters.Font.Size = 10
            .TextFrame.Characters.Text = CStr(Sheet1.Cells(24, 5).Value2)
            .ZOrder msoBringToFront
        End With
        Sheet1.Shapes.Range(Array(RadialPlotIndex, RadialPlotIndex + 1)).Group
    End If
        
    If (IsEmpty(Sheet4.Cells(59, 8)) = False) Then
        Set MarkLine3 = Sheet1.Shapes.AddLabel(msoTextOrientationHorizontal, HorizontalIndent + 15, VerticalIndent + 55, 20, 10)
            With MarkLine3
                .TextFrame.AutoSize = True
                .Line.Visible = msoFalse
                .Fill.Transparency = 1
                .TextFrame.VerticalAlignment = xlVAlignCenter
                .TextFrame.HorizontalAlignment = xlHAlignLeft
                .TextFrame.Characters.Font.Name = "Arial"
                .TextFrame.Characters.Font.Size = 10
                .TextFrame.Characters.Text = "De (" + Sheet1.Cells(21, 8) + ")" + " = " + CStr(Sheet4.Cells(59, 8).Value2) + " Gy"
                .ZOrder msoBringToFront
            End With
        Sheet1.Shapes.Range(Array(RadialPlotIndex, RadialPlotIndex + 1)).Group
        End If
        
        
        If (DistributionIsLogNormal = True) Then
            MinimumValue = Round(100 / NumberOfXAxisTickMarks / Step + 0.5, 0)
            Step = 2
            NumberOfXAxisTickMarks = Round(Log(100 / MinimumValue) / Log(2), 0)
            ReDim XAxisOtherTickMark(NumberOfXAxisTickMarks + 1) As Shape
            For Counter = 0 To NumberOfXAxisTickMarks
                Set XAxisOtherTickMark(Counter) = Sheet1.Shapes.AddLine(YAxisLocation + 100 / MinimumValue / Round(2 ^ Counter, 0) * XGraphScale, XAxisLocation, YAxisLocation + 100 / MinimumValue / Round(2 ^ Counter, 0) * XGraphScale, XAxisLocation - XAxisTickMarkLength)
                XAxisOtherTickMark(Counter).Line.Weight = 1
                XAxisOtherTickMark(Counter).Line.DashStyle = msoLineSolid
                XAxisOtherTickMark(Counter).Line.ForeColor.RGB = RGB(0, 0, 0)
                XAxisOtherTickMark(Counter).Line.Visible = msoTrue
                XAxisOtherTickMark(Counter).ZOrder msoBringToFront
                Sheet1.Shapes.Range(Array(RadialPlotIndex, RadialPlotIndex + 1)).Group
            Next Counter
    
            XAxisOtherLabelScalingFactor = 0.5
            PreviousLabelPosition = 0
            ReDim XAxisOtherLabel(NumberOfXAxisTickMarks + 2) As Shape
            For Counter = 0 To NumberOfXAxisTickMarks
                Indent = Len(CStr(Counter * Step))
                If (Abs(100 / MinimumValue / Round(2 ^ Counter, 0) - PreviousLabelPosition) * XGraphScale > 20) Then
                    Set XAxisOtherLabel(Counter) = Sheet1.Shapes.AddLabel(msoTextOrientationHorizontal, YAxisLocation + 100 / MinimumValue / Round(2 ^ Counter, 0) * XGraphScale - 5 * (Indent + 0.3), XAxisLocation - XAxisTickMarkLength - 15, 10 * (Indent + 1) * XAxisOtherLabelScalingFactor, 10)
                    XAxisOtherLabel(Counter).TextFrame.AutoSize = True
                    XAxisOtherLabel(Counter).Line.Visible = msoFalse
                    XAxisOtherLabel(Counter).Fill.Transparency = 1
                    XAxisOtherLabel(Counter).TextFrame.VerticalAlignment = xlVAlignCenter
                    XAxisOtherLabel(Counter).TextFrame.HorizontalAlignment = xlHAlignCenter
                    XAxisOtherLabel(Counter).TextFrame.Characters.Font.Size = 10
                    XAxisOtherLabel(Counter).TextFrame.Characters.Font.Name = "Arial"
                    XAxisOtherLabel(Counter).TextFrame.Characters.Text = CStr(MinimumValue * Round(2 ^ Counter, 0))
                    XAxisOtherLabel(Counter).ZOrder msoBringToFront
                    Sheet1.Shapes.Range(Array(RadialPlotIndex, RadialPlotIndex + 1)).Group
                    PreviousLabelPosition = 100 / MinimumValue / Round(2 ^ Counter, 0)
                End If
            Next Counter
    
            Set XAxisOtherLabel(NumberOfXAxisTickMarks + 1) = Sheet1.Shapes.AddLabel(msoTextOrientationHorizontal, YAxisLocation + MaximumXDeviation * XGraphScale / 2 - 45 * PrecisionScalingFactor - 7, XAxisLocation - 34, 70, 10)
                If (ExcelVersionIs2007 = True) Then
                    XAxisOtherLabel(NumberOfXAxisTickMarks + 1).TextFrame.AutoSize = True
                Else
                    XAxisOtherLabel(NumberOfXAxisTickMarks + 1).TextFrame.AutoMargins = True
                End If
                XAxisOtherLabel(NumberOfXAxisTickMarks + 1).Line.Visible = msoFalse
                XAxisOtherLabel(NumberOfXAxisTickMarks + 1).Fill.Transparency = 1
                XAxisOtherLabel(NumberOfXAxisTickMarks + 1).TextFrame.VerticalAlignment = xlVAlignCenter
                XAxisOtherLabel(NumberOfXAxisTickMarks + 1).TextFrame.HorizontalAlignment = xlHAlignCenter
                XAxisOtherLabel(NumberOfXAxisTickMarks + 1).TextFrame.Characters.Font.Size = 10
                XAxisOtherLabel(NumberOfXAxisTickMarks + 1).TextFrame.Characters.Font.Name = "Arial"
                XAxisOtherLabel(NumberOfXAxisTickMarks + 1).TextFrame.Characters.Text = "Relative standard error (%)"
                XAxisOtherLabel(NumberOfXAxisTickMarks + 1).ZOrder msoBringToFront
            Sheet1.Shapes.Range(Array(RadialPlotIndex, RadialPlotIndex + 1)).Group
        End If

        '  draw y axis
        ScaledStandardDeviation = YGraphScale
        Set YAxis = Sheet1.Shapes.AddLine(YAxisLocation, VerticalCenterLine - 2 * ScaledStandardDeviation, YAxisLocation, VerticalCenterLine + 2 * ScaledStandardDeviation)
            YAxis.Line.DashStyle = msoLineSolid
            YAxis.Line.Weight = 1.2
            YAxis.Line.ForeColor.RGB = RGB(0, 0, 0)
            YAxis.Line.Visible = msoTrue
            YAxis.ZOrder msoBringToFront
        Sheet1.Shapes.Range(Array(RadialPlotIndex, RadialPlotIndex + 1)).Group
        
        YAxisTickMarkLength = 4
        NumberOfYAxisTickMarks = Round(Exp(Round(Log(4 * ScaledStandardDeviation / 10) / Log(2) - 0.5, 0) * Log(2)))
        If (NumberOfYAxisTickMarks < 2) Then
            NumberOfYAxisTickMarks = 2
        End If
        Step = 4 / NumberOfYAxisTickMarks
        ReDim YAxisTickMark(NumberOfYAxisTickMarks + 1) As Shape
        For Counter = 0 To NumberOfYAxisTickMarks
            If (Counter = NumberOfYAxisTickMarks / 2) Then
                Set YAxisTickMark(Counter) = Sheet1.Shapes.AddLine(YAxisLocation - YAxisTickMarkLength, VerticalCenterLine, YAxisLocation + YAxisTickMarkLength, VerticalCenterLine)
                YAxisTickMark(Counter).Line.Weight = 1
            Else
                If (Counter / 2 = Round(Counter / 2)) Then
                    Set YAxisTickMark(Counter) = Sheet1.Shapes.AddLine(YAxisLocation, VerticalCenterLine - (2 - Counter * Step) * ScaledStandardDeviation, YAxisLocation - YAxisTickMarkLength, VerticalCenterLine - (2 - Counter * Step) * ScaledStandardDeviation)
                Else
                    Set YAxisTickMark(Counter) = Sheet1.Shapes.AddLine(YAxisLocation, VerticalCenterLine - (2 - Counter * Step) * ScaledStandardDeviation, YAxisLocation - YAxisTickMarkLength / 2, VerticalCenterLine - (2 - Counter * Step) * ScaledStandardDeviation)
                End If
                If ((Counter = 0) Or (Counter = NumberOfYAxisTickMarks)) Then
                    YAxisTickMark(Counter).Line.Weight = 1.2
                Else
                    YAxisTickMark(Counter).Line.Weight = 1.2
                End If
            End If
            YAxisTickMark(Counter).Line.DashStyle = msoLineSolid
            YAxisTickMark(Counter).Line.ForeColor.RGB = RGB(0, 0, 0)
            YAxisTickMark(Counter).Line.Visible = msoTrue
            YAxisTickMark(Counter).ZOrder msoBringToFront
            Sheet1.Shapes.Range(Array(RadialPlotIndex, RadialPlotIndex + 1)).Group
        Next Counter
        
        YAxisLabelScalingFactor = 0.5
        NumberOfYAxisLabels = NumberOfYAxisTickMarks / 2
        ReDim YAxisLabel(NumberOfYAxisLabels + 2) As Shape
        For Counter = 0 To NumberOfYAxisLabels
            Set YAxisLabel(Counter) = Sheet1.Shapes.AddLabel(msoTextOrientationHorizontal, HorizontalIndent + 2 * (XIndent - YAxisTickMarkLength - 5 * Len(CStr(2 - 2 * Counter * Step)) * YAxisLabelScalingFactor) / 3, VerticalCenterLine - (2 - 2 * Counter * Step) * ScaledStandardDeviation - 10, 10 * Len(CStr(2 - 2 * Counter * Step)) * YAxisLabelScalingFactor, 10)
            YAxisLabel(Counter).TextFrame.AutoSize = True
            YAxisLabel(Counter).Line.Visible = msoFalse
            YAxisLabel(Counter).Fill.Transparency = 1
            YAxisLabel(Counter).TextFrame.VerticalAlignment = xlVAlignCenter
            YAxisLabel(Counter).TextFrame.HorizontalAlignment = xlHAlignCenter
            YAxisLabel(Counter).TextFrame.Characters.Font.Size = 10
            YAxisLabel(Counter).TextFrame.Characters.Font.Name = "Arial"
            YAxisLabel(Counter).TextFrame.Characters.Text = CStr(2 - 2 * Counter * Step)
            YAxisLabel(Counter).ZOrder msoBringToFront
            Sheet1.Shapes.Range(Array(RadialPlotIndex, RadialPlotIndex + 1)).Group
        Next Counter
        
        StandardizedEstimateScalingFactor = 0
        Set YAxisLabel(NumberOfYAxisLabels + 1) = Sheet1.Shapes.AddLabel(msoTextOrientationUpward, HorizontalIndent + (XIndent - YAxisTickMarkLength) / 4 - 15, VerticalCenterLine - 100 * StandardizedEstimateScalingFactor / 2, 10, 10)
            YAxisLabel(NumberOfYAxisLabels + 1).TextFrame.AutoSize = True
            YAxisLabel(NumberOfYAxisLabels + 1).Line.Visible = msoFalse
            YAxisLabel(NumberOfYAxisLabels + 1).Fill.Transparency = 1
            YAxisLabel(NumberOfYAxisLabels + 1).TextFrame.VerticalAlignment = xlVAlignCenter
            YAxisLabel(NumberOfYAxisLabels + 1).TextFrame.HorizontalAlignment = xlHAlignCenter
            YAxisLabel(NumberOfYAxisLabels + 1).TextFrame.Characters.Font.Size = 11
            YAxisLabel(NumberOfYAxisLabels + 1).TextFrame.Characters.Font.Name = "Arial"
            YAxisLabel(NumberOfYAxisLabels + 1).TextFrame.Characters.Text = "Standardized Estimate"
            YAxisLabel(NumberOfYAxisLabels + 1).ZOrder msoBringToFront
        Sheet1.Shapes.Range(Array(RadialPlotIndex, RadialPlotIndex + 1)).Group
        
        '  draw z axis
        ChangeInAngle = Atn(0.5) / 5
        Set ZAxis = Sheet1.Shapes.BuildFreeform(msoEditingCorner, YAxisLocation + Radius * Cos(5 * ChangeInAngle), VerticalCenterLine - Radius * Sin(5 * ChangeInAngle))
        For Counter = 1 To 10
            Angle = ChangeInAngle * (5 - Counter)
            ZAxis.AddNodes msoSegmentCurve, msoEditingAuto, YAxisLocation + Radius * Cos(Angle), VerticalCenterLine - Radius * Sin(Angle)
        Next Counter
        Set ZAxisShape = ZAxis.ConvertToShape
            ZAxisShape.Fill.Visible = msoFalse
            ZAxisShape.Line.ForeColor.RGB = RGB(0, 0, 0)
            ZAxisShape.Line.Weight = 1.2
            ZAxisShape.ZOrder msoBringToFront
        Sheet1.Shapes.Range(Array(RadialPlotIndex, RadialPlotIndex + 1)).Group
        MaximumDose = MeanValue + Atn(MaximumAngle) * ScalingFactor
        MinimumDose = MeanValue + Atn(MinimumAngle) * ScalingFactor
        If (DistributionIsLogNormal = True) Then
            MaximumDose = Exp(MaximumDose)
            MinimumDose = Exp(MinimumDose)
        End If
        
        SignificantFigures = 0
        Do
            SignificantFigures = SignificantFigures + 1
            MaximumEquivalentDose = ApproximateNumber(MaximumDose, SignificantFigures, False)
            MinimumEquivalentDose = ApproximateNumber(MinimumDose, SignificantFigures, True)
        Loop Until (Abs(MaximumEquivalentDose - MinimumEquivalentDose) > 0)
         
        ZAxisTickMarkLength = 4
        If (DistributionIsLogNormal = True) Then
            Step = StandardDeviation * ScalingFactor
            EquivalentDose = MeanValue
            NumberOfNegativeZAxisTickMarks = 0
            Do
                EquivalentDose = EquivalentDose - Step
                NumberOfNegativeZAxisTickMarks = NumberOfNegativeZAxisTickMarks + 1
            Loop Until (EquivalentDose < Log(MinimumEquivalentDose))
            
            InitialEquivalentDose = EquivalentDose
            EquivalentDose = MeanValue
            NumberOfPositiveZAxisTickMarks = 0
            Do
                EquivalentDose = EquivalentDose + Step
                NumberOfPositiveZAxisTickMarks = NumberOfPositiveZAxisTickMarks + 1
            Loop Until (EquivalentDose > Log(MaximumEquivalentDose))
            EquivalentDose = InitialEquivalentDose
            
        Else
            Step = StandardDeviation * ScalingFactor
            EquivalentDose = MeanValue
            NumberOfNegativeZAxisTickMarks = 0
            Do
                EquivalentDose = EquivalentDose - Step
                NumberOfNegativeZAxisTickMarks = NumberOfNegativeZAxisTickMarks + 1
            Loop Until (EquivalentDose < MinimumEquivalentDose)
            InitialEquivalentDose = EquivalentDose
            EquivalentDose = MeanValue
            NumberOfPositiveZAxisTickMarks = 0
            Do
                EquivalentDose = EquivalentDose + Step
                NumberOfPositiveZAxisTickMarks = NumberOfPositiveZAxisTickMarks + 1
            Loop Until (EquivalentDose > MaximumEquivalentDose)
            EquivalentDose = InitialEquivalentDose
        End If
        
        
        If (NumberOfPositiveZAxisTickMarks < 2) Then
            NumberOfPositiveZAxisTickMarks = 2
        End If
        If (NumberOfNegativeZAxisTickMarks < 2) Then
            NumberOfNegativeZAxisTickMarks = 2
        End If
        If (NumberOfNegativeZAxisTickMarks > NumberOfPositiveZAxisTickMarks) Then
            NumberOfPositiveZAxisTickMarks = NumberOfNegativeZAxisTickMarks
        ElseIf (NumberOfNegativeZAxisTickMarks < NumberOfPositiveZAxisTickMarks) Then
            NumberOfNegativeZAxisTickMarks = NumberOfPositiveZAxisTickMarks
        Else
            '  do nothing
        End If
        NumberOfZAxisTickMarks = NumberOfNegativeZAxisTickMarks + NumberOfPositiveZAxisTickMarks
        
        OldAngle = 0
        ReDim ZAxisTickMark(NumberOfZAxisTickMarks + 1) As Shape
        StepMultiplier = -1
        For Counter = 0 To NumberOfPositiveZAxisTickMarks
            If (DistributionIsLogNormal = True) Then
                EquivalentDose = Exp(MeanValue + Counter * Step)
                Angle = Atn((Log(EquivalentDose) - MeanValue) * YGraphScale * ScalingFactor / XGraphScale) / ScalingFactor
            Else
                EquivalentDose = MeanValue + Counter * Step
                Angle = Atn((EquivalentDose - MeanValue) * YGraphScale * ScalingFactor / XGraphScale) / ScalingFactor
            End If
            If ((Angle >= -Atn(0.5)) And (Angle <= Atn(0.5)) And ((Angle = 0) Or (Angle >= OldAngle + Atn(0.5) / 10))) Then
                Set ZAxisTickMark(Counter) = Sheet1.Shapes.AddLine(YAxisLocation + Radius * Cos(Angle), VerticalCenterLine - Radius * Sin(Angle), YAxisLocation + (Radius + ZAxisTickMarkLength) * Cos(Angle), VerticalCenterLine - (Radius + ZAxisTickMarkLength) * Sin(Angle))
                ZAxisTickMark(Counter).Line.Weight = 1
                ZAxisTickMark(Counter).Line.DashStyle = msoLineSolid
                ZAxisTickMark(Counter).Line.ForeColor.RGB = RGB(0, 0, 0)
                ZAxisTickMark(Counter).Line.Visible = msoTrue
                ZAxisTickMark(Counter).ZOrder msoBringToFront
                Sheet1.Shapes.Range(Array(RadialPlotIndex, RadialPlotIndex + 1)).Group
                OldAngle = Angle
            End If
        Next Counter
        
        OldAngle = 0
        For Counter = (NumberOfPositiveZAxisTickMarks + 1) To NumberOfZAxisTickMarks
            If (DistributionIsLogNormal = True) Then
                EquivalentDose = Exp(MeanValue - (Counter - NumberOfPositiveZAxisTickMarks) * Step)
                Angle = Atn((Log(EquivalentDose) - MeanValue) * YGraphScale * ScalingFactor / XGraphScale) / ScalingFactor
            Else
                EquivalentDose = MeanValue - (Counter - NumberOfPositiveZAxisTickMarks) * Step
                Angle = Atn((EquivalentDose - MeanValue) * YGraphScale * ScalingFactor / XGraphScale) / ScalingFactor
            End If
            If ((Angle >= -Atn(0.5)) And (Angle <= Atn(0.5)) And ((Angle = 0) Or (Angle <= OldAngle - Atn(0.5) / 10))) Then
                Set ZAxisTickMark(Counter) = Sheet1.Shapes.AddLine(YAxisLocation + Radius * Cos(Angle), VerticalCenterLine - Radius * Sin(Angle), YAxisLocation + (Radius + ZAxisTickMarkLength) * Cos(Angle), VerticalCenterLine - (Radius + ZAxisTickMarkLength) * Sin(Angle))
                ZAxisTickMark(Counter).Line.Weight = 1
                ZAxisTickMark(Counter).Line.DashStyle = msoLineSolid
                ZAxisTickMark(Counter).Line.ForeColor.RGB = RGB(0, 0, 0)
                ZAxisTickMark(Counter).Line.Visible = msoTrue
                ZAxisTickMark(Counter).ZOrder msoBringToFront
                Sheet1.Shapes.Range(Array(RadialPlotIndex, RadialPlotIndex + 1)).Group
                OldAngle = Angle
            End If
        Next Counter

        If (DistributionIsLogNormal = True) Then
            EquivalentDose = Exp(InitialEquivalentDose)
        Else
            EquivalentDose = InitialEquivalentDose
        End If
        
        BetaStrength = Sheet4.Cells(41, 5).Value / 60
        
        OldAngle = 0
        PreviousLabelPosition = 0
        ReDim ZAxisLabel(NumberOfZAxisTickMarks + 2) As Shape
        For Counter = 0 To NumberOfPositiveZAxisTickMarks
            If (DistributionIsLogNormal = True) Then
                EquivalentDose = Exp(MeanValue + Counter * Step)
                Angle = Atn((Log(EquivalentDose) - MeanValue) * YGraphScale * ScalingFactor / XGraphScale) / ScalingFactor
            Else
                EquivalentDose = MeanValue + Counter * Step
                Angle = Atn((EquivalentDose - MeanValue) * YGraphScale * ScalingFactor / XGraphScale) / ScalingFactor
            End If
            If ((Angle >= -Atn(0.5)) And (Angle <= Atn(0.5)) And ((Angle = 0) Or (Angle >= OldAngle + Atn(0.5) / 10))) Then
                If ((Angle = 0) Or (Abs(Radius * (Angle - PreviousLabelPosition)) > 10)) Then
                    Digits = SignificantFigures
                    If (IsEmpty(Sheet1.Cells(26, 14)) = False) Then
                        If (IsNumeric(Sheet1.Cells(26, 14)) = True) Then
                            Digits = Round(Sheet1.Cells(26, 14).Value)
                        End If
                    End If

                    If (Sheet1.Cells(24, 14).Value = "Time") Then
                       If Digits = 1 Then
                           LabelText = CStr(Format(EquivalentDose, "###0.0"))
                       ElseIf Digits = 2 Then
                           LabelText = CStr(Format(EquivalentDose, "###0.00"))
                       Else
                           LabelText = CStr(Round(EquivalentDose, Digits))
                       End If
                    Else
                       If Digits = 1 Then
                           LabelText = CStr(Format(EquivalentDose * BetaStrength, "###0.0"))
                       ElseIf Digits = 2 Then
                           LabelText = CStr(Format(EquivalentDose * BetaStrength, "###0.00"))
                       Else
                           LabelText = CStr(Round(EquivalentDose * BetaStrength, Digits))
                       End If
                    End If
                    
                    Set ZAxisLabel(Counter) = Sheet1.Shapes.AddLabel(msoTextOrientationHorizontal, YAxisLocation + (Radius + ZAxisTickMarkLength + 3) * Cos(Angle) + 3 * Len(LabelText) - 10, VerticalCenterLine - (Radius + ZAxisTickMarkLength + 3) * Sin(Angle) - 10, 10, 10)
                    If (ExcelVersionIs2007 = True) Then
                        ZAxisLabel(Counter).TextFrame.AutoSize = True
                    Else
                        ZAxisLabel(Counter).TextFrame.AutoMargins = True
                    End If
                    ZAxisLabel(Counter).Line.Visible = msoFalse
                    ZAxisLabel(Counter).Fill.Transparency = 1
                    ZAxisLabel(Counter).TextFrame.VerticalAlignment = xlVAlignCenter
                    ZAxisLabel(Counter).TextFrame.HorizontalAlignment = xlHAlignCenter
                    ZAxisLabel(Counter).TextFrame.Characters.Text = LabelText
                    ZAxisLabel(Counter).TextFrame.Characters.Font.Size = 10
                    ZAxisLabel(Counter).TextFrame.Characters.Font.Name = "Arial"
                    ZAxisLabel(Counter).ZOrder msoBringToFront
                    Sheet1.Shapes.Range(Array(RadialPlotIndex, RadialPlotIndex + 1)).Group
                    PreviousLabelPosition = Angle
                    OldAngle = Angle
                End If
            End If
        Next Counter
        
        OldAngle = 0
        For Counter = (NumberOfPositiveZAxisTickMarks + 1) To NumberOfZAxisTickMarks
            If (DistributionIsLogNormal = True) Then
                EquivalentDose = Exp(MeanValue - (Counter - NumberOfPositiveZAxisTickMarks) * Step)
                Angle = Atn((Log(EquivalentDose) - MeanValue) * YGraphScale * ScalingFactor / XGraphScale) / ScalingFactor
            Else
                EquivalentDose = MeanValue - (Counter - NumberOfPositiveZAxisTickMarks) * Step
                Angle = Atn((EquivalentDose - MeanValue) * YGraphScale * ScalingFactor / XGraphScale) / ScalingFactor
            End If
            If ((Angle >= -Atn(0.5)) And (Angle <= Atn(0.5)) And ((Angle = 0) Or (Angle <= OldAngle - Atn(0.5) / 8))) Then
                If (Abs(Radius * (Angle - PreviousLabelPosition)) > 10) Then
                    If ((IsEmpty(Sheet1.Cells(26, 14)) = False) And (IsNumeric(Sheet1.Cells(26, 14)) = True)) Then
                        Digits = Round(Sheet1.Cells(26, 14).Value)
                          'If (IsEmpty(sheet1.Cells(27, 13)) = True) Then
                        If (Sheet1.Cells(24, 14).Value = "Time") Then
                            If Digits = 1 Then
                                LabelText = CStr(Format(EquivalentDose, "###0.0"))
                            ElseIf Digits = 2 Then
                                LabelText = CStr(Format(EquivalentDose, "###0.00"))
                            Else
                                LabelText = CStr(Round(EquivalentDose, Digits))
                            End If
                        Else
                            If Digits = 1 Then
                                LabelText = CStr(Format(EquivalentDose * BetaStrength, "###0.0"))
                            ElseIf Digits = 2 Then
                                LabelText = CStr(Format(EquivalentDose * BetaStrength, "###0.00"))
                            Else
                                LabelText = CStr(Round(EquivalentDose * BetaStrength, Digits))
                            End If
                        End If
                    Else
                        Digits = SignificantFigures
                        Do
                            LabelText = CStr(Round(EquivalentDose * BetaStrength, Digits))
                            Digits = Digits + 1
                        Loop Until (CDbl(LabelText) <> 0)
                    End If
                    Set ZAxisLabel(Counter) = Sheet1.Shapes.AddLabel(msoTextOrientationHorizontal, YAxisLocation + (Radius + ZAxisTickMarkLength + 3) * Cos(Angle) + 3 * Len(LabelText) - 10, VerticalCenterLine - (Radius + ZAxisTickMarkLength + 3) * Sin(Angle) - 10, 10, 10)
                    If (ExcelVersionIs2007 = True) Then
                        ZAxisLabel(Counter).TextFrame.AutoSize = True
                    Else
                        ZAxisLabel(Counter).TextFrame.AutoMargins = True
                    End If
                    ZAxisLabel(Counter).Line.Visible = msoFalse
                    ZAxisLabel(Counter).Fill.Transparency = 1
                    ZAxisLabel(Counter).TextFrame.VerticalAlignment = xlVAlignCenter
                    ZAxisLabel(Counter).TextFrame.HorizontalAlignment = xlHAlignCenter
                    ZAxisLabel(Counter).TextFrame.Characters.Text = LabelText
                    ZAxisLabel(Counter).TextFrame.Characters.Font.Size = 10
                    ZAxisLabel(Counter).TextFrame.Characters.Font.Name = "Arial"
                    ZAxisLabel(Counter).ZOrder msoBringToFront
                    Sheet1.Shapes.Range(Array(RadialPlotIndex, RadialPlotIndex + 1)).Group
                    PreviousLabelPosition = Angle
                    OldAngle = Angle
                End If
            End If
        Next Counter

        ApparentValueScalingFactor = 0
        Set ZAxisLabel(NumberOfZAxisTickMarks + 1) = Sheet1.Shapes.AddLabel(msoTextOrientationDownward, HorizontalIndent + (BoxSize + XIndent + Radius + ZAxisTickMarkLength) / 2 + 38, VerticalCenterLine - 70 * ApparentValueScalingFactor, 10, 140)
            ZAxisLabel(NumberOfZAxisTickMarks + 1).TextFrame.AutoSize = True
            ZAxisLabel(NumberOfZAxisTickMarks + 1).Line.Visible = msoFalse
            ZAxisLabel(NumberOfZAxisTickMarks + 1).Fill.Transparency = 1
            ZAxisLabel(NumberOfZAxisTickMarks + 1).TextFrame.VerticalAlignment = xlVAlignCenter
            ZAxisLabel(NumberOfZAxisTickMarks + 1).TextFrame.HorizontalAlignment = xlHAlignCenter
            ZAxisLabel(NumberOfZAxisTickMarks + 1).TextFrame.Characters.Font.Size = 11
            ZAxisLabel(NumberOfZAxisTickMarks + 1).TextFrame.Characters.Font.Name = "Arial"
        If (Sheet1.Cells(24, 14).Value = "Dose") Then
            ZAxisLabel(NumberOfZAxisTickMarks + 1).TextFrame.Characters.Text = "Equivalent dose (Gy)"
        Else
            ZAxisLabel(NumberOfZAxisTickMarks + 1).TextFrame.Characters.Text = "Exposure time (sec)"
        End If
        ZAxisLabel(NumberOfZAxisTickMarks + 1).ZOrder msoBringToFront
        Sheet1.Shapes.Range(Array(RadialPlotIndex, RadialPlotIndex + 1)).Group
     
        '  draw wedge
        If IsEmpty(Sheet1.Cells(21, 5)) = False Then
            ComponentNumber = -1
            Do
                ComponentNumber = ComponentNumber + 1
                If (Sheet1.FiniteMixtureModelButton.Value = True And Sheet1.FMMOptionButton4.Value = True) Then
                    ComponentEDValue = Sheet1.Cells(ComponentNumber + 35, 5).Value
                    ComponentEDError = Sheet1.Cells(ComponentNumber + 35, 7).Value
                    NumberOfValidDataPoints = 1
                Else
                    ComponentEDValue = Sheet1.Cells(21, 5).Value
                    ComponentEDError = Sheet1.Cells(21, 7).Value
                    NumberOfValidDataPoints = ValidXValues
                End If
                
            Dim Zvalue1(0 To 1) As Single
            Dim XValue1(0 To 1) As Single
            Dim YValue1(0 To 1) As Single
    
            For Row = 0 To 1
                    If (PresentDistributionStandardDeviation = True) Then
                        If (DistributionIsLogNormal = True) Then
                            If (Sheet1.MeanErrorButton.Value = True) Then
                                Zvalue1(Row) = Log(ComponentEDValue) - MeanValue - (4 * Row - 2) * ComponentEDError * Sqr(NumberOfValidDataPoints) / ComponentEDValue
                            Else
                                Zvalue1(Row) = Log(ComponentEDValue) - MeanValue - (4 * Row - 2) * ComponentEDError / ComponentEDValue
                            End If
                        Else
                            If (Sheet1.MeanErrorButton.Value = True) Then
                                Zvalue1(Row) = ComponentEDValue - MeanValue - (4 * Row - 2) * ComponentEDError * Sqr(NumberOfValidDataPoints)
                            Else
                                Zvalue1(Row) = ComponentEDValue - MeanValue - (4 * Row - 2) * ComponentEDError
                            End If
                        End If
                        Angle = Atn(YGraphScale * Zvalue1(Row) / XGraphScale)
                    Else
                        If (Sheet1.CAMOptionButton1.Value = True) Then
                            If (DistributionIsLogNormal = True) Then
                            Zvalue1(Row) = Log(ComponentEDValue) - MeanValue
                            Else
                            Zvalue1(Row) = ComponentEDValue - MeanValue
                            End If
                        Else
                            If (DistributionIsLogNormal = True) Then
                               Zvalue1(Row) = (Log(ComponentEDValue) - MeanValue)
                            Else
                               Zvalue1(Row) = ComponentEDValue - MeanValue
                            End If
                        End If
                        OldAngle = Atn(YGraphScale * Zvalue1(Row) / XGraphScale)
                        Angle = Atn(Tan(OldAngle) - (4 * Row - 2) * ScaledStandardDeviation / Sqr(4 * ScaledStandardDeviation * ScaledStandardDeviation + Radius * Radius - (8 * Row - 4) * ScaledStandardDeviation * Radius * Sin(OldAngle)) / Cos(OldAngle))
                        If (Row = 0) Then
                            LowerAngle = Angle
                        Else
                            UpperAngle = Angle
                        End If
                    End If
                    XValue1(Row) = YAxisLocation + Radius * Cos(Angle)
                    YValue1(Row) = VerticalCenterLine - Radius * Sin(Angle)
                    If (YValue1(Row) < VerticalIndent) Then
                        YValue1(Row) = VerticalIndent
                        XValue1(Row) = YAxisLocation + Abs(VerticalCenterLine - YValue1(Row)) / Tan(Angle)
                    End If
                    If (YValue1(Row) > XAxisLocation) Then
                        YValue1(Row) = XAxisLocation
                        XValue1(Row) = YAxisLocation - Abs(VerticalCenterLine - YValue1(Row)) / Tan(Angle)
                    End If
                   If (XValue1(0) >= YAxisLocation) Then
                        If (PresentDistributionStandardDeviation = True) Then
                                Set TwoSigmaLine(Row) = Sheet1.Shapes.AddLine(YAxisLocation, VerticalCenterLine, XValue1(Row), YValue1(Row))
                                TwoSigmaLine(Row).Line.DashStyle = msoLineDash
                                TwoSigmaLine(Row).Line.Weight = 1.2
                                TwoSigmaLine(Row).Line.ForeColor.RGB = RGB(0, 142, 192)
                                TwoSigmaLine(Row).Line.Visible = msoTrue
                                TwoSigmaLine(Row).ZOrder msoBringToFront
                                Sheet1.Shapes.Range(Array(RadialPlotIndex, RadialPlotIndex + 1)).Group
                        End If
                    End If
                Next Row
                
                If (XValue1(0) >= YAxisLocation) Then
                    If (PresentDistributionStandardDeviation = False) Then
                        Dim CenterLine As Shape
                        Set CenterLine = Sheet1.Shapes.AddLine(YAxisLocation, VerticalCenterLine, (XValue1(0) + XValue1(1)) / 2, (YValue1(0) + YValue1(1)) / 2)
                            With CenterLine
                                .Line.DashStyle = msoLineSolid
                                .Line.Weight = 1
                                .Line.ForeColor.RGB = RGB(150, 150, 150)
                                .Line.Visible = msoTrue
                                .ZOrder msoBringToFront
                            End With
                        Sheet1.Shapes.Range(Array(RadialPlotIndex, RadialPlotIndex + 1)).Group
                        Dim Pts(1 To 4, 1 To 2) As Single
                         Pts(1, 1) = YAxisLocation
                         Pts(1, 2) = VerticalCenterLine + 2 * ScaledStandardDeviation
                         Pts(2, 1) = YAxisLocation
                         Pts(2, 2) = VerticalCenterLine + -2 * ScaledStandardDeviation
                         Pts(3, 1) = XValue1(0)
                         Pts(3, 2) = YValue1(0)
                         Pts(4, 1) = XValue1(1)
                         Pts(4, 2) = YValue1(1)
                         With Sheet1
                            .Shapes.AddPolyline SafeArrayOfPoints:=Pts
                            .Shapes(RadialPlotIndex + 1).Line.Visible = msoFalse
                            .Shapes(RadialPlotIndex + 1).Fill.Transparency = 0.8
                            .Shapes(RadialPlotIndex + 1).Fill.ForeColor.RGB = RGB(0, 0, 255)
                            .Shapes(RadialPlotIndex + 1).ZOrder msoBringToFront
                            .Shapes.Range(Array(RadialPlotIndex, RadialPlotIndex + 1)).Group
                        End With
                    End If
                End If
                If (Sheet1.FiniteMixtureModelButton.Value = True And Sheet1.FMMOptionButton4.Value = True) Then
                    If ((IsEmpty(Sheet1.Cells(ComponentNumber + 36, 5)) = False) And (IsNumeric(Sheet1.Cells(ComponentNumber + 36, 5)) = True)) Then
                        NumberOfComponents = ComponentNumber + 1
                    Else
                        NumberOfComponents = ComponentNumber
                    End If
                Else
                    NumberOfComponents = 0
                End If
            Loop Until (ComponentNumber = NumberOfComponents)
        End If
        
        ' plotting data
        HalfOvalSize = Sheet1.Cells(28, 14)
        OvalSize = 2 * HalfOvalSize
        Counter = -1
        Row = 3
        For Row = 1 To n
                Weight = se(Row)
                If (Weight < 0) Then
                    Weight = -Weight
                End If
                Deviation = zi(Row) - MeanValue
                Weight = 1 / Sqr(1 / (Weight * Weight + OverDispersionTerm * OverDispersionTerm))
                XValue = YAxisLocation + XGraphScale / Weight
                YValue = VerticalCenterLine - YGraphScale * Deviation / Weight
                If ((YValue > (VerticalIndent + HalfOvalSize)) And (YValue < (XAxisLocation - 30))) Then
                    Counter = Counter + 1
                    ReDim Preserve ApparentDoseDataPoint(Counter)
                    Set ApparentDoseDataPoint(Counter) = Sheet1.Shapes.AddShape(msoShapeOval, XValue - HalfOvalSize, YValue - HalfOvalSize, OvalSize, OvalSize)
                    With ApparentDoseDataPoint(Counter)
                        .Fill.ForeColor.RGB = RGB(Cells(30, 14), Cells(30, 15), Cells(30, 16))
                        .Fill.Transparency = 0.7
                        .Line.ForeColor.RGB = RGB(0, 0, 0)
                        .Line.Weight = 0.8
                        .ZOrder msoBringToFront
                    End With
                    Sheet1.Shapes.Range(Array(RadialPlotIndex, RadialPlotIndex + 1)).Group
                End If
        Next Row
    Else
        Sheet1.RadialPlotCheckBox.Value = False
    End If
    Sheet1.Shapes(RadialPlotIndex).Name = "RadialPlot"
    Sheet1.Shapes("RadialPlot").Locked = False
End Sub