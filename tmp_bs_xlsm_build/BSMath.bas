Attribute VB_Name = "BSMath"
Option Explicit

Public Function BSMat(ByVal nRows As Long, ByVal nCols As Long) As Variant
    Dim a() As Double
    ReDim a(1 To nRows, 1 To nCols)
    BSMat = a
End Function

Public Function BSIdentity(ByVal n As Long) As Variant
    Dim a As Variant, i As Long
    a = BSMat(n, n)
    For i = 1 To n: a(i, i) = 1#: Next i
    BSIdentity = a
End Function

Public Function BSTranspose(ByVal a As Variant) As Variant
    Dim r As Long, c As Long, i As Long, j As Long, b As Variant
    r = UBound(a, 1): c = UBound(a, 2): b = BSMat(c, r)
    For i = 1 To r
        For j = 1 To c: b(j, i) = a(i, j): Next j
    Next i
    BSTranspose = b
End Function

Public Function BSMultiply(ByVal a As Variant, ByVal b As Variant) As Variant
    Dim ar As Long, ac As Long, bc As Long, i As Long, j As Long, k As Long
    Dim out As Variant, s As Double
    ar = UBound(a, 1): ac = UBound(a, 2): bc = UBound(b, 2)
    out = BSMat(ar, bc)
    For i = 1 To ar
        For j = 1 To bc
            s = 0#
            For k = 1 To ac: s = s + a(i, k) * b(k, j): Next k
            out(i, j) = s
        Next j
    Next i
    BSMultiply = out
End Function

Public Function BSSubtract(ByVal a As Variant, ByVal b As Variant) As Variant
    Dim i As Long, j As Long, out As Variant
    out = BSMat(UBound(a, 1), UBound(a, 2))
    For i = 1 To UBound(a, 1)
        For j = 1 To UBound(a, 2): out(i, j) = a(i, j) - b(i, j): Next j
    Next i
    BSSubtract = out
End Function

Public Function BSInverse(ByVal a As Variant) As Variant
    Dim n As Long, aug As Variant, out As Variant
    Dim i As Long, j As Long, k As Long, p As Long
    Dim best As Double, v As Double, t As Double
    n = UBound(a, 1): aug = BSMat(n, n * 2)
    For i = 1 To n
        For j = 1 To n: aug(i, j) = a(i, j): Next j
        aug(i, n + i) = 1#
    Next i
    For j = 1 To n
        p = j: best = Abs(aug(j, j))
        For i = j + 1 To n
            If Abs(aug(i, j)) > best Then p = i: best = Abs(aug(i, j))
        Next i
        If best < 0.000000000001 Then Err.Raise vbObjectError + 701, , "Singular matrix"
        If p <> j Then
            For k = 1 To n * 2
                t = aug(j, k): aug(j, k) = aug(p, k): aug(p, k) = t
            Next k
        End If
        v = aug(j, j)
        For k = 1 To n * 2: aug(j, k) = aug(j, k) / v: Next k
        For i = 1 To n
            If i <> j Then
                v = aug(i, j)
                For k = 1 To n * 2: aug(i, k) = aug(i, k) - v * aug(j, k): Next k
            End If
        Next i
    Next j
    out = BSMat(n, n)
    For i = 1 To n
        For j = 1 To n: out(i, j) = aug(i, n + j): Next j
    Next i
    BSInverse = out
End Function

Public Sub BSJacobiEigen(ByVal inputMatrix As Variant, ByRef eigenValues As Variant, ByRef eigenVectors As Variant)
    Dim a As Variant, v As Variant, n As Long, iter As Long, maxIter As Long
    Dim i As Long, j As Long, p As Long, q As Long
    Dim maxOff As Double, app As Double, aqq As Double, apq As Double
    Dim phi As Double, c As Double, s As Double, aip As Double, aiq As Double
    a = inputMatrix: n = UBound(a, 1): v = BSIdentity(n): maxIter = 100 * n * n
    For iter = 1 To maxIter
        maxOff = 0#: p = 1: q = IIf(n > 1, 2, 1)
        For i = 1 To n - 1
            For j = i + 1 To n
                If Abs(a(i, j)) > maxOff Then maxOff = Abs(a(i, j)): p = i: q = j
            Next j
        Next i
        If maxOff < 0.0000000000001 Then Exit For
        app = a(p, p): aqq = a(q, q): apq = a(p, q)
        phi = 0.5 * Atn2(2# * apq, aqq - app)
        c = Cos(phi): s = Sin(phi)
        For i = 1 To n
            If i <> p And i <> q Then
                aip = a(i, p): aiq = a(i, q)
                a(i, p) = c * aip - s * aiq: a(p, i) = a(i, p)
                a(i, q) = s * aip + c * aiq: a(q, i) = a(i, q)
            End If
        Next i
        a(p, p) = c * c * app - 2# * s * c * apq + s * s * aqq
        a(q, q) = s * s * app + 2# * s * c * apq + c * c * aqq
        a(p, q) = 0#: a(q, p) = 0#
        For i = 1 To n
            aip = v(i, p): aiq = v(i, q)
            v(i, p) = c * aip - s * aiq
            v(i, q) = s * aip + c * aiq
        Next i
    Next iter
    ReDim eigenValues(1 To n)
    For i = 1 To n: eigenValues(i) = a(i, i): Next i
    eigenVectors = v
End Sub

Private Function Atn2(ByVal y As Double, ByVal x As Double) As Double
    Const PI As Double = 3.14159265358979
    If x > 0# Then
        Atn2 = Atn(y / x)
    ElseIf x < 0# And y >= 0# Then
        Atn2 = Atn(y / x) + PI
    ElseIf x < 0# And y < 0# Then
        Atn2 = Atn(y / x) - PI
    ElseIf x = 0# And y > 0# Then
        Atn2 = PI / 2#
    ElseIf x = 0# And y < 0# Then
        Atn2 = -PI / 2#
    Else
        Atn2 = 0#
    End If
End Function

Public Sub BSSortEigenDesc(ByRef values As Variant, ByRef vectors As Variant)
    Dim n As Long, i As Long, j As Long, k As Long, t As Double
    n = UBound(values)
    For i = 1 To n - 1
        For j = i + 1 To n
            If values(j) > values(i) Then
                t = values(i): values(i) = values(j): values(j) = t
                For k = 1 To n
                    t = vectors(k, i): vectors(k, i) = vectors(k, j): vectors(k, j) = t
                Next k
            End If
        Next j
    Next i
End Sub

Public Function BSSymmetricPower(ByVal a As Variant, ByVal power As Double, Optional ByVal floorValue As Double = 0#) As Variant
    Dim vals As Variant, vecs As Variant, d As Variant, out As Variant
    Dim i As Long, n As Long, x As Double
    BSJacobiEigen a, vals, vecs: n = UBound(vals): d = BSMat(n, n)
    For i = 1 To n
        x = vals(i)
        If power < 0# Then
            If x > floorValue Then d(i, i) = x ^ power Else d(i, i) = 0#
        Else
            If x < 0# And Abs(x) < 0.0000000001 Then x = 0#
            If x >= 0# Then d(i, i) = x ^ power
        End If
    Next i
    out = BSMultiply(BSMultiply(vecs, d), BSTranspose(vecs))
    BSSymmetricPower = out
End Function

Public Function BSPolarOrthogonal(ByVal a As Variant) As Variant
    Dim gram As Variant
    gram = BSMultiply(BSTranspose(a), a)
    BSPolarOrthogonal = BSMultiply(a, BSSymmetricPower(gram, -0.5, 0.000000000001))
End Function

Public Function BSCorrelation(ByVal x As Variant) As Variant
    Dim n As Long, p As Long, i As Long, j As Long, k As Long
    Dim means() As Double, sd() As Double, out As Variant, s As Double
    n = UBound(x, 1): p = UBound(x, 2): ReDim means(1 To p): ReDim sd(1 To p)
    For j = 1 To p
        For i = 1 To n: means(j) = means(j) + x(i, j): Next i
        means(j) = means(j) / n
        For i = 1 To n: sd(j) = sd(j) + (x(i, j) - means(j)) ^ 2: Next i
        If n > 1 Then sd(j) = Sqr(sd(j) / (n - 1))
    Next j
    out = BSMat(p, p)
    For j = 1 To p
        out(j, j) = 1#
        For k = j + 1 To p
            s = 0#
            For i = 1 To n: s = s + (x(i, j) - means(j)) * (x(i, k) - means(k)): Next i
            If n > 1 And sd(j) > 0# And sd(k) > 0# Then s = s / ((n - 1) * sd(j) * sd(k)) Else s = 0#
            out(j, k) = s: out(k, j) = s
        Next k
    Next j
    BSCorrelation = out
End Function

Public Function BSStandardize(ByVal x As Variant) As Variant
    Dim n As Long, p As Long, i As Long, j As Long
    Dim means() As Double, sd As Double, out As Variant
    n = UBound(x, 1): p = UBound(x, 2): ReDim means(1 To p): out = BSMat(n, p)
    For j = 1 To p
        For i = 1 To n: means(j) = means(j) + x(i, j): Next i
        means(j) = means(j) / n: sd = 0#
        For i = 1 To n: sd = sd + (x(i, j) - means(j)) ^ 2: Next i
        sd = Sqr(sd / n)
        If sd <= 0# Then Err.Raise vbObjectError + 702, , "Constant analysis variable"
        For i = 1 To n: out(i, j) = (x(i, j) - means(j)) / sd: Next i
    Next j
    BSStandardize = out
End Function

Private Sub BSEquamaxObjective(ByVal l As Variant, ByRef gradient As Variant, ByRef criterion As Double)
    Dim p As Long, k As Long, i As Long, j As Long, h As Long
    Dim sq As Variant, g As Variant, sumOtherFactors As Double, sumOtherVars As Double
    Dim f1 As Double, f2 As Double
    p = UBound(l, 1): k = UBound(l, 2): sq = BSMat(p, k): g = BSMat(p, k)
    For i = 1 To p
        For j = 1 To k: sq(i, j) = l(i, j) ^ 2: Next j
    Next i
    For i = 1 To p
        For j = 1 To k
            sumOtherFactors = 0#: sumOtherVars = 0#
            For h = 1 To k
                If h <> j Then sumOtherFactors = sumOtherFactors + sq(i, h)
            Next h
            For h = 1 To p
                If h <> i Then sumOtherVars = sumOtherVars + sq(h, j)
            Next h
            f1 = f1 + sq(i, j) * sumOtherFactors / 8#
            f2 = f2 + sq(i, j) * sumOtherVars / 8#
            g(i, j) = 0.5 * l(i, j) * (sumOtherFactors + sumOtherVars)
        Next j
    Next i
    criterion = f1 + f2: gradient = g
End Sub

Public Function BSEquamax(ByVal loadings As Variant) As Variant
    Dim rotation As Variant, rotated As Variant, objGrad As Variant, gradient As Variant
    Dim projected As Variant, m As Variant, sym As Variant, candidate As Variant, newRotation As Variant
    Dim criterion As Double, newCriterion As Double, alpha As Double, ss As Double
    Dim i As Long, j As Long, h As Long, iter As Long, trial As Long, k As Long
    k = UBound(loadings, 2): rotation = BSIdentity(k): alpha = 1#
    rotated = BSMultiply(loadings, rotation): BSEquamaxObjective rotated, objGrad, criterion
    gradient = BSMultiply(BSTranspose(loadings), objGrad)
    For iter = 0 To 250
        m = BSMultiply(BSTranspose(rotation), gradient): sym = BSMat(k, k)
        For i = 1 To k
            For j = 1 To k: sym(i, j) = (m(i, j) + m(j, i)) / 2#: Next j
        Next i
        projected = BSSubtract(gradient, BSMultiply(rotation, sym)): ss = 0#
        For i = 1 To k
            For j = 1 To k: ss = ss + projected(i, j) ^ 2: Next j
        Next i
        ss = Sqr(ss): If ss < 0.00001 Then Exit For
        alpha = alpha * 2#
        For trial = 0 To 10
            candidate = BSMat(k, k)
            For i = 1 To k
                For j = 1 To k: candidate(i, j) = rotation(i, j) - alpha * projected(i, j): Next j
            Next i
            newRotation = BSPolarOrthogonal(candidate)
            rotated = BSMultiply(loadings, newRotation)
            BSEquamaxObjective rotated, objGrad, newCriterion
            If newCriterion < criterion - 0.5 * ss * ss * alpha Then Exit For
            alpha = alpha / 2#
        Next trial
        rotation = newRotation: criterion = newCriterion
        gradient = BSMultiply(BSTranspose(loadings), objGrad)
    Next iter
    BSEquamax = rotated
End Function

Public Function BSFactorRatios(ByVal x As Variant, ByVal y As Variant, ByVal safeMapping As Boolean, ByRef collisionText As String) As Variant
    Dim n As Long, p As Long, corr As Variant, vals As Variant, vecs As Variant, loads As Variant
    Dim rotated As Variant, ordered As Variant, ss() As Double, order() As Long
    Dim i As Long, j As Long, k As Long, t As Long, td As Double
    Dim primary() As Long, used As Object, collision As Boolean, mapFactorToVar() As Long
    Dim bestScore As Double, score As Double, a As Long, b As Long, c As Long, d As Long
    Dim invCorr As Variant, temp As Variant, invSqrt As Variant, coeff As Variant, z As Variant, scores As Variant
    Dim xc As Variant, xtx As Variant, xty As Variant, beta As Variant
    Dim meanY As Double, sdY As Double, means() As Double, sds() As Double, ratios() As Double, total As Double
    n = UBound(x, 1): p = 4: ReDim ratios(1 To 4)
    If n < 4 Then BSFactorRatios = ratios: Exit Function
    corr = BSCorrelation(x): BSJacobiEigen corr, vals, vecs: BSSortEigenDesc vals, vecs
    loads = BSMat(p, p)
    For i = 1 To p
        For j = 1 To p
            If vals(j) > 0# Then loads(i, j) = vecs(i, j) * Sqr(vals(j))
        Next j
    Next i
    rotated = BSEquamax(loads): ReDim ss(1 To p): ReDim order(1 To p)
    For j = 1 To p
        order(j) = j
        For i = 1 To p: ss(j) = ss(j) + rotated(i, j) ^ 2: Next i
    Next j
    For i = 1 To p - 1
        For j = i + 1 To p
            If ss(j) > ss(i) Then
                td = ss(i): ss(i) = ss(j): ss(j) = td
                t = order(i): order(i) = order(j): order(j) = t
            End If
        Next j
    Next i
    ordered = BSMat(p, p)
    For i = 1 To p
        For j = 1 To p: ordered(i, j) = rotated(i, order(j)): Next j
    Next i
    rotated = ordered: ReDim primary(1 To p): Set used = CreateObject("Scripting.Dictionary")
    For i = 1 To p
        primary(i) = 1
        For j = 2 To p
            If Abs(rotated(i, j)) > Abs(rotated(i, primary(i))) Then primary(i) = j
        Next j
        If used.Exists(CStr(primary(i))) Then collision = True Else used.Add CStr(primary(i)), True
    Next i
    ReDim mapFactorToVar(1 To p)
    If collision And safeMapping Then
        bestScore = -1#
        For a = 1 To 4
            For b = 1 To 4
                If b <> a Then
                    For c = 1 To 4
                        If c <> a And c <> b Then
                            For d = 1 To 4
                                If d <> a And d <> b And d <> c Then
                                    score = Abs(rotated(1, a)) + Abs(rotated(2, b)) + Abs(rotated(3, c)) + Abs(rotated(4, d))
                                    If score > bestScore Then
                                        bestScore = score
                                        mapFactorToVar(a) = 1: mapFactorToVar(b) = 2: mapFactorToVar(c) = 3: mapFactorToVar(d) = 4
                                    End If
                                End If
                            Next d
                        End If
                    Next c
                End If
            Next b
        Next a
        collisionText = "Factor collision corrected"
    Else
        For i = 1 To p: mapFactorToVar(primary(i)) = i: Next i
        If collision Then collisionText = "Factor collision detected"
    End If
    invCorr = BSInverse(corr)
    temp = BSMultiply(BSMultiply(BSTranspose(rotated), invCorr), rotated)
    invSqrt = BSSymmetricPower(temp, -0.5, 0.000000000001)
    coeff = BSMultiply(BSMultiply(invCorr, rotated), invSqrt)
    z = BSStandardize(x): scores = BSMultiply(z, coeff)
    ReDim means(1 To p): ReDim sds(1 To p): meanY = 0#
    For i = 1 To n: meanY = meanY + y(i, 1): Next i
    meanY = meanY / n
    For j = 1 To p
        For i = 1 To n: means(j) = means(j) + scores(i, j): Next i
        means(j) = means(j) / n
    Next j
    xc = BSMat(n, p): xty = BSMat(p, 1): sdY = 0#
    For i = 1 To n
        sdY = sdY + (y(i, 1) - meanY) ^ 2
        For j = 1 To p
            xc(i, j) = scores(i, j) - means(j)
            sds(j) = sds(j) + xc(i, j) ^ 2
            xty(j, 1) = xty(j, 1) + xc(i, j) * (y(i, 1) - meanY)
        Next j
    Next i
    sdY = Sqr(sdY / (n - 1))
    xtx = BSMultiply(BSTranspose(xc), xc): beta = BSMultiply(BSInverse(xtx), xty)
    For j = 1 To p
        sds(j) = Sqr(sds(j) / (n - 1))
        If sdY > 0# Then beta(j, 1) = beta(j, 1) * sds(j) / sdY Else beta(j, 1) = 0#
        If mapFactorToVar(j) > 0 Then ratios(mapFactorToVar(j)) = beta(j, 1)
    Next j
    For i = 1 To p: total = total + Abs(ratios(i)): Next i
    If total > 0# Then
        For i = 1 To p
            ratios(i) = Abs(ratios(i)) / total * 100#
        Next i
    End If
    BSFactorRatios = ratios
End Function
