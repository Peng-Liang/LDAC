Public Function DefineW(ByVal which As Integer, ByVal lower As Single, ByVal upper As Single) As Single
    Dim w As Single
    If which = 1 Then
        w = 0.3
    ElseIf which = 2 Then
        w = (Abs(upper) + Abs(lower)) * 0.4
    ElseIf which = 3 Then
        w = Sqr(upper) * 0.2
    ElseIf which = 4 Then
        w = (Abs(upper) + Abs(lower)) * 0.6
    End If
    DefineW = w
End Function
