Public Function SliceMAM3(ByVal n As Integer, ByVal p As Single, ByVal Gamma As Single, ByVal Sigma As Single, ByVal Sigmab As Single, ByVal which As Integer, ByVal lower As Single, ByVal upper As Single, zi(), si()) As Single
    Dim gx0, gx1, gL, gR, logy As Single
    Dim U, L, R, w As Single
    Dim x1 As Single
    gx0 = LikeliFunc3(n, p, Gamma, Sigma, Sigmab, zi, si)
    logy = gx0 + Log(Rnd())
    'define an appropriate step size w
    w = DefineW(which, lower, upper)
    'Find the initial interval [L,R] to sample from.
    If (which = 1) Then
        U = Rnd() * w
        L = p - U
        R = p + (w - U)
    ElseIf (which = 2) Then
        U = Rnd() * w
        L = Gamma - U
        R = Gamma + (w - U)
    ElseIf (which = 3) Then
        U = Rnd() * w
        L = Sigma - U
        R = Sigma + (w - U)
    End If
    
    If (L <= lower) Then
        L = lower
    End If
    
    If (R >= upper) Then
        R = upper
    End If
    
    'Expand the interval until its ends are outside the slice, or until the limit on steps is reached.
    Do
        If (which = 1) Then
            gL = LikeliFunc3(n, L, Gamma, Sigma, Sigmab, zi, si)
        ElseIf (which = 2) Then
            gL = LikeliFunc3(n, p, L, Sigma, Sigmab, zi, si)
        ElseIf (which = 3) Then
            gL = LikeliFunc3(n, p, Gamma, L, Sigmab, zi, si)
        End If
        L = L - w
    Loop Until (L <= lower Or gL <= logy)
        
        'For the right side
    Do
        If (which = 1) Then
            gR = LikeliFunc3(n, R, Gamma, Sigma, Sigmab, zi, si)
        ElseIf (which = 2) Then
            gR = LikeliFunc3(n, p, R, Sigma, Sigmab, zi, si)
        ElseIf (which = 3) Then
            gR = LikeliFunc3(n, p, Gamma, R, Sigmab, zi, si)
        End If
        R = R + w
    Loop Until (R >= upper Or gR <= logy)
    
    
    'Shrink the interval
    If (L <= lower) Then
        L = lower
    End If
    
    If (R >= upper) Then
        R = upper
    End If
    
    'Sample from the interval (with shrinking).
    Do
        x1 = L + (R - L) * Rnd()
        If (which = 1) Then
            gx1 = LikeliFunc3(n, x1, Gamma, Sigma, Sigmab, zi, si)
        ElseIf (which = 2) Then
            gx1 = LikeliFunc3(n, p, x1, Sigma, Sigmab, zi, si)
        ElseIf (which = 3) Then
            gx1 = LikeliFunc3(n, p, Gamma, x1, Sigmab, zi, si)
        End If
        
        If (which = 1) Then
            If (x1 > p) Then
                R = x1
            Else
                L = x1
            End If
        ElseIf which = 2 Then
            If x1 > Gamma Then
                R = x1
            Else
                L = x1
            End If
        ElseIf which = 3 Then
            If x1 > Sigma Then
                R = x1
            Else
                L = x1
            End If
        End If
    Loop Until (gx1 >= logy)
    SliceMAM3 = x1
End Function

Public Function SliceMAM4(ByVal n As Integer, ByVal p As Single, ByVal Gamma As Single, ByVal mu As Single, ByVal Sigma As Single, ByVal Sigmab As Single, ByVal which As Integer, ByVal lower As Single, ByVal upper As Single, zi(), si()) As Single
    Dim gx0, gx1, gL, gR, logy As Single
    Dim U, L, R, w As Single
    Dim x1 As Single
    gx0 = LikeliFunc4(n, p, Gamma, mu, Sigma, Sigmab, zi, si)
    logy = gx0 + Log(Rnd())
    'define an appropriate step size w
    w = DefineW(which, lower, upper)
    'Find the initial interval [L,R] to sample from.
    If (which = 1) Then
        U = Rnd() * w
        L = p - U
        R = p + (w - U)
    ElseIf (which = 2) Then
        U = Rnd() * w
        L = Gamma - U
        R = Gamma + (w - U)
    ElseIf (which = 3) Then
        U = Rnd() * w
        L = Sigma - U
        R = Sigma + (w - U)
    ElseIf which = 4 Then
        U = Rnd() * w
        L = mu - U
        R = mu + (w - U)
    End If
    
    If (L <= lower) Then
        L = lower
    End If
    
    If (R >= upper) Then
        R = upper
    End If
    
'Expand the interval until its ends are outside the slice, or until the limit on steps is reached.
    Do
        If (which = 1) Then
            gL = LikeliFunc4(n, L, Gamma, mu, Sigma, Sigmab, zi, si)
        ElseIf (which = 2) Then
            gL = LikeliFunc4(n, p, L, mu, Sigma, Sigmab, zi, si)
        ElseIf (which = 3) Then
            gL = LikeliFunc4(n, p, Gamma, mu, L, Sigmab, zi, si)
        ElseIf which = 4 Then
            gL = LikeliFunc4(n, p, Gamma, L, Sigma, Sigmab, zi, si)
        End If
        
        L = L - w
    Loop Until (L <= lower Or gL <= logy)
    
'For the right side
    Do
        If (which = 1) Then
            gR = LikeliFunc4(n, R, Gamma, mu, Sigma, Sigmab, zi, si)
        ElseIf (which = 2) Then
            gR = LikeliFunc4(n, p, R, mu, Sigma, Sigmab, zi, si)
        ElseIf (which = 3) Then
            gR = LikeliFunc4(n, p, Gamma, mu, R, Sigmab, zi, si)
        ElseIf which = 4 Then
            gR = LikeliFunc4(n, p, Gamma, R, Sigma, Sigmab, zi, si)
        End If
        R = R + w
    Loop Until (R >= upper Or gR <= logy)
    
'Shrink the interval
    If (L <= lower) Then
        L = lower
    End If
    
    If (R >= upper) Then
        R = upper
    End If
    
'Sample from the interval (with shrinking).
    Do
        x1 = L + (R - L) * Rnd()
        If (which = 1) Then
            gx1 = LikeliFunc4(n, x1, Gamma, mu, Sigma, Sigmab, zi, si)
        ElseIf (which = 2) Then
            gx1 = LikeliFunc4(n, p, x1, mu, Sigma, Sigmab, zi, si)
        ElseIf (which = 3) Then
            gx1 = LikeliFunc4(n, p, Gamma, mu, x1, Sigmab, zi, si)
        ElseIf which = 4 Then
            gx1 = LikeliFunc4(n, p, Gamma, x1, Sigma, Sigmab, zi, si)
        End If
        
        If (which = 1) Then
            If (x1 > p) Then
                R = x1
            Else
                L = x1
            End If
        ElseIf which = 2 Then
            If x1 > Gamma Then
                R = x1
            Else
                L = x1
            End If
        ElseIf which = 3 Then
            If x1 > Sigma Then
                R = x1
            Else
                L = x1
            End If
        ElseIf which = 4 Then
            If x1 > mu Then
                R = x1
            Else
                L = x1
            End If
        End If
    Loop Until (gx1 >= logy)
    SliceMAM4 = x1
End Function