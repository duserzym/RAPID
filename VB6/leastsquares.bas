Attribute VB_Name = "modLeastSquares"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Copyright (c) 2006-2007, Sergey Bochkanov (ALGLIB project).
'
'>>> SOURCE LICENSE >>>
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation (www.fsf.org); either version 2 of the
'License, or (at your option) any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'A copy of the GNU General Public License is available at
'http://www.fsf.org/licensing/licenses
'
'>>> END OF LICENSE >>>
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Constants
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'Machine Epsilon = machine precision for the double data type
Public Const MachineEpsilon As Double = 2 ^ (-53)


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Routines
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'Weighted approximation by arbitrary function basis in a space of arbitrary
'dimension using linear least squares method.
'
'Input parameters:
'    Y   -   array[0..N-1]
'            It contains a set  of  function  values  in  N  points.  Space
'            dimension  and  points  don't  matter.  Procedure  works  with
'            function values in these points and values of basis  functions
'            only.
'
'    W   -   array[0..N-1]
'            It contains weights corresponding  to  function  values.  Each
'            summand in square sum of approximation deviations  from  given
'            values is multiplied by the square of corresponding weight.
'
'    FMatrix-a table of basis functions values, array[0..N-1, 0..M-1].
'            FMatrix[I, J] - value of J-th basis function in I-th point.
'
'    N   -   number of points used. N>=1.
'    M   -   number of basis functions, M>=1.
'
'Output parameters:
'    C   -   decomposition coefficients.
'            Array of real numbers whose index goes from 0 to M-1.
'            C[j] - j-th basis function coefficient.
'
'  -- ALGLIB --
'     Copyright by Bochkanov Sergey
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub BuildGeneralLeastSquares(ByRef Y() As Double, _
         ByRef W() As Double, _
         ByRef FMatrix() As Double, _
         ByVal N As Long, _
         ByVal M As Long, _
         ByRef C() As Double)
    Dim i As Long
    Dim J As Long
    Dim A() As Double
    Dim Q() As Double
    Dim VT() As Double
    Dim B() As Double
    Dim Tau() As Double
    Dim b2() As Double
    Dim TauQ() As Double
    Dim TauP() As Double
    Dim D() As Double
    Dim E() As Double
    Dim IsUpperA As Boolean
    Dim MI As Long
    Dim NI As Long
    Dim V As Double
    Dim i_ As Long
    Dim i1_ As Long

    MI = N
    NI = M
    ReDim C(0# To NI - 1#)
    
    '
    ' Initialize design matrix.
    ' Here we are making MI>=NI.
    '
    ReDim A(1# To NI, 1# To MaxInt(MI, NI))
    ReDim B(1# To MaxInt(MI, NI))
    For i = 1# To MI Step 1
        B(i) = W(i - 1#) * Y(i - 1#)
    Next i
    For i = MI + 1# To NI Step 1
        B(i) = 0#
    Next i
    For J = 1# To NI Step 1
        i1_ = (0#) - (1#)
        For i_ = 1# To MI Step 1
            A(J, i_) = FMatrix(i_ + i1_, J - 1#)
        Next i_
    Next J
    For J = 1# To NI Step 1
        For i = MI + 1# To NI Step 1
            A(J, i) = 0#
        Next i
    Next J
    For J = 1# To NI Step 1
        For i = 1# To MI Step 1
            A(J, i) = A(J, i) * W(i - 1#)
        Next i
    Next J
    MI = MaxInt(MI, NI)
    
    '
    ' LQ-decomposition of A'
    ' B2 := Q*B
    '
    Call LQDecomposition(A, NI, MI, Tau)
    Call UnpackQFromLQ(A, NI, MI, Tau, NI, Q)
    ReDim b2(1# To 1#, 1# To NI)
    For J = 1# To NI Step 1
        b2(1#, J) = 0#
    Next J
    For i = 1# To NI Step 1
        V = 0#
        For i_ = 1# To MI Step 1
            V = V + B(i_) * Q(i, i_)
        Next i_
        b2(1#, i) = V
    Next i
    
    '
    ' Back from A' to A
    ' Making cols(A)=rows(A)
    '
    For i = 1# To NI - 1# Step 1
        For i_ = i + 1# To NI Step 1
            A(i, i_) = A(i_, i)
        Next i_
    Next i
    For i = 2# To NI Step 1
        For J = 1# To i - 1# Step 1
            A(i, J) = 0#
        Next J
    Next i
    
    '
    ' Bidiagonal decomposition of A
    ' A = Q * d2 * P'
    ' B2 := (Q'*B2')'
    '
    Call ToBidiagonal(A, NI, NI, TauQ, TauP)
    Call MultiplyByQFromBidiagonal(A, NI, NI, TauQ, b2, 1#, NI, True, False)
    Call UnpackPTFromBidiagonal(A, NI, NI, TauP, NI, VT)
    Call UnpackDiagonalsFromBidiagonal(A, NI, NI, IsUpperA, D, E)
    
    '
    ' Singular value decomposition of A
    ' A = U * d * V'
    ' B2 := (U'*B2')'
    '
    If Not BidiagonalSVDDecomposition(D, E, NI, IsUpperA, False, b2, 1#, Q, 0#, VT, NI) Then
        For i = 0# To NI - 1# Step 1
            C(i) = 0#
        Next i
        Exit Sub
    End If
    
    '
    ' B2 := (d^(-1) * B2')'
    '
    If D(1#) <> 0# Then
        For i = 1# To NI Step 1
            If D(i) > MachineEpsilon * 10# * Sqr(NI) * D(1#) Then
                b2(1#, i) = b2(1#, i) / D(i)
            Else
                b2(1#, i) = 0#
            End If
        Next i
    End If
    
    '
    ' B := (V * B2')'
    '
    For i = 1# To NI Step 1
        B(i) = 0#
    Next i
    For i = 1# To NI Step 1
        V = b2(1#, i)
        For i_ = 1# To NI Step 1
            B(i_) = B(i_) + V * VT(i, i_)
        Next i_
    Next i
    
    '
    ' Out
    '
    For i = 1# To NI Step 1
        C(i - 1#) = B(i)
    Next i
End Sub





''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Weighted cubic spline approximation using linear least squares
'
'Input parameters:
'    X   -   array[0..N-1], abscissas
'    Y   -   array[0..N-1], function values
'    W   -   array[0..N-1], weights.
'    A, B-   interval to build splines in.
'    N   -   number of points used. N>=1.
'    M   -   number of basic splines, M>=2.
'
'Output parameters:
'    CTbl-   coefficients table to be used by SplineInterpolation function.
'  -- ALGLIB --
'     Copyright by Bochkanov Sergey
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub BuildSplineLeastSquares(ByRef X() As Double, _
         ByRef Y() As Double, _
         ByRef W() As Double, _
         ByVal A As Double, _
         ByVal B As Double, _
         ByVal N As Long, _
         ByVal M As Long, _
         ByRef CTbl() As Double)
    Dim i As Long
    Dim J As Long
    Dim MA() As Double
    Dim Q() As Double
    Dim VT() As Double
    Dim MB() As Double
    Dim Tau() As Double
    Dim b2() As Double
    Dim TauQ() As Double
    Dim TauP() As Double
    Dim D() As Double
    Dim E() As Double
    Dim IsUpperA As Boolean
    Dim MI As Long
    Dim NI As Long
    Dim V As Double
    Dim SX() As Double
    Dim SY() As Double
    Dim i_ As Long

    MI = N
    NI = M
    ReDim SX(0# To NI - 1#)
    ReDim SY(0# To NI - 1#)
    
    '
    ' Initializing design matrix
    ' Here we are making MI>=NI
    '
    ReDim MA(1# To NI, 1# To MaxInt(MI, NI))
    ReDim MB(1# To MaxInt(MI, NI))
    For J = 0# To NI - 1# Step 1
        SX(J) = A + (B - A) * J / (NI - 1#)
    Next J
    For J = 0# To NI - 1# Step 1
        For i = 0# To NI - 1# Step 1
            SY(i) = 0#
        Next i
        SY(J) = 1#
        Call BuildCubicSpline(SX, SY, NI, 0#, 0#, 0#, 0#, CTbl)
        For i = 0# To MI - 1# Step 1
            MA(J + 1#, i + 1#) = W(i) * SplineInterpolation(CTbl, X(i))
        Next i
    Next J
    For J = 1# To NI Step 1
        For i = MI + 1# To NI Step 1
            MA(J, i) = 0#
        Next i
    Next J
    
    '
    ' Initializing right part
    '
    For i = 0# To MI - 1# Step 1
        MB(i + 1#) = W(i) * Y(i)
    Next i
    For i = MI + 1# To NI Step 1
        MB(i) = 0#
    Next i
    MI = MaxInt(MI, NI)
    
    '
    ' LQ-decomposition of A'
    ' B2 := Q*B
    '
    Call LQDecomposition(MA, NI, MI, Tau)
    Call UnpackQFromLQ(MA, NI, MI, Tau, NI, Q)
    ReDim b2(1# To 1#, 1# To NI)
    For J = 1# To NI Step 1
        b2(1#, J) = 0#
    Next J
    For i = 1# To NI Step 1
        V = 0#
        For i_ = 1# To MI Step 1
            V = V + MB(i_) * Q(i, i_)
        Next i_
        b2(1#, i) = V
    Next i
    
    '
    ' Back from A' to A
    ' Making cols(A)=rows(A)
    '
    For i = 1# To NI - 1# Step 1
        For i_ = i + 1# To NI Step 1
            MA(i, i_) = MA(i_, i)
        Next i_
    Next i
    For i = 2# To NI Step 1
        For J = 1# To i - 1# Step 1
            MA(i, J) = 0#
        Next J
    Next i
    
    '
    ' Bidiagonal decomposition of A
    ' A = Q * d2 * P'
    ' B2 := (Q'*B2')'
    '
    Call ToBidiagonal(MA, NI, NI, TauQ, TauP)
    Call MultiplyByQFromBidiagonal(MA, NI, NI, TauQ, b2, 1#, NI, True, False)
    Call UnpackPTFromBidiagonal(MA, NI, NI, TauP, NI, VT)
    Call UnpackDiagonalsFromBidiagonal(MA, NI, NI, IsUpperA, D, E)
    
    '
    ' Singular value decomposition of A
    ' A = U * d * V'
    ' B2 := (U'*B2')'
    '
    If Not BidiagonalSVDDecomposition(D, E, NI, IsUpperA, False, b2, 1#, Q, 0#, VT, NI) Then
        For i = 1# To NI Step 1
            D(i) = 0#
            b2(1#, i) = 0#
            For J = 1# To NI Step 1
                If i = J Then
                    VT(i, J) = 1#
                Else
                    VT(i, J) = 0#
                End If
            Next J
        Next i
        b2(1#, 1#) = 1#
    End If
    
    '
    ' B2 := (d^(-1) * B2')'
    '
    For i = 1# To NI Step 1
        If D(i) > MachineEpsilon * 10# * Sqr(NI) * D(1#) Then
            b2(1#, i) = b2(1#, i) / D(i)
        Else
            b2(1#, i) = 0#
        End If
    Next i
    
    '
    ' B := (V * B2')'
    '
    For i = 1# To NI Step 1
        MB(i) = 0#
    Next i
    For i = 1# To NI Step 1
        V = b2(1#, i)
        For i_ = 1# To NI Step 1
            MB(i_) = MB(i_) + V * VT(i, i_)
        Next i_
    Next i
    
    '
    ' Forming result spline
    '
    For i = 0# To NI - 1# Step 1
        SY(i) = MB(i + 1#)
    Next i
    Call BuildCubicSpline(SX, SY, NI, 0#, 0#, 0#, 0#, CTbl)
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Polynomial approximation using least squares method
'
'The subroutine calculates coefficients  of  the  polynomial  approximating
'given function. It is recommended to use this function only if you need to
'obtain coefficients of approximation polynomial. If you have to build  and
'calculate polynomial approximation (NOT coefficients), it's better to  use
'BuildChebyshevLeastSquares      subroutine     in     combination     with
'CalculateChebyshevLeastSquares   subroutine.   The   result  of  Chebyshev
'polynomial approximation is equivalent to the result obtained using powers
'of X, but has higher  accuracy  due  to  better  numerical  properties  of
'Chebyshev polynomials.
'
'Input parameters:
'    X   -   array[0..N-1], abscissas
'    Y   -   array[0..N-1], function values
'    N   -   number of points, N>=1
'    M   -   order of polynomial required, M>=0
'
'Output parameters:
'    C   -   approximating polynomial coefficients, array[0..M],
'            C[i] - coefficient at X^i.
'
'  -- ALGLIB --
'     Copyright by Bochkanov Sergey
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub BuildPolynomialLeastSquares(ByRef X() As Double, _
         ByRef Y() As Double, _
         ByVal N As Long, _
         ByVal M As Long, _
         ByRef C() As Double)
    Dim CTbl() As Double
    Dim W() As Double
    Dim C1() As Double
    Dim MaxX As Double
    Dim MinX As Double
    Dim i As Long
    Dim J As Long
    Dim k As Long
    Dim E As Double
    Dim D As Double
    Dim L1 As Double
    Dim L2 As Double
    Dim Z2() As Double
    Dim Z1() As Double

    
    '
    ' Initialize
    '
    MaxX = X(0#)
    MinX = X(0#)
    For i = 1# To N - 1# Step 1
        If X(i) > MaxX Then
            MaxX = X(i)
        End If
        If X(i) < MinX Then
            MinX = X(i)
        End If
    Next i
    If MinX = MaxX Then
        MinX = MinX - 0.5
        MaxX = MaxX + 0.5
    End If
    ReDim W(0# To N - 1#)
    For i = 0# To N - 1# Step 1
        W(i) = 1#
    Next i
    
    '
    ' Build Chebyshev approximation
    '
    Call BuildChebyshevLeastSquares(X, Y, W, MinX, MaxX, N, M, CTbl)
    
    '
    ' From Chebyshev to powers of X
    '
    ReDim C1(0# To M)
    For i = 0# To M Step 1
        C1(i) = 0#
    Next i
    D = 0#
    For i = 0# To M Step 1
        For k = i To M Step 1
            E = C1(k)
            C1(k) = 0#
            If i <= 1# And k = i Then
                C1(k) = 1#
            Else
                If i <> 0# Then
                    C1(k) = 2# * D
                End If
                If k > i + 1# Then
                    C1(k) = C1(k) - C1(k - 2#)
                End If
            End If
            D = E
        Next k
        D = C1(i)
        E = 0#
        k = i
        Do While k <= M
            E = E + C1(k) * CTbl(k)
            k = k + 2#
        Loop
        C1(i) = E
    Next i
    
    '
    ' Linear translation
    '
    L1 = 2# / (CTbl(M + 2#) - CTbl(M + 1#))
    L2 = -(2# * CTbl(M + 1#) / (CTbl(M + 2#) - CTbl(M + 1#))) - 1#
    ReDim C(0# To M)
    ReDim Z2(0# To M)
    ReDim Z1(0# To M)
    C(0#) = C1(0#)
    Z1(0#) = 1#
    Z2(0#) = 1#
    For i = 1# To M Step 1
        Z2(i) = 1#
        Z1(i) = L2 * Z1(i - 1#)
        C(0#) = C(0#) + C1(i) * Z1(i)
    Next i
    For J = 1# To M Step 1
        Z2(0#) = L1 * Z2(0#)
        C(J) = C1(J) * Z2(0#)
        For i = J + 1# To M Step 1
            k = i - J
            Z2(k) = L1 * Z2(k) + Z2(k - 1#)
            C(J) = C(J) + C1(i) * Z2(k) * Z1(k)
        Next i
    Next J
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Chebyshev polynomial approximation using least squares method.
'
'The algorithm reduces interval [A, B] to the interval [-1,1], then  builds
'least squares approximation using Chebyshev polynomials.
'
'Input parameters:
'    X   -   array[0..N-1], abscissas
'    Y   -   array[0..N-1], function values
'    W   -   array[0..N-1], weights
'    A, B-   interval to build approximating polynomials in.
'    N   -   number of points used. N>=1.
'    M   -   order of polynomial, M>=0. This parameter is passed into
'            CalculateChebyshevLeastSquares function.
'
'Output parameters:
'    CTbl - coefficient table. This parameter is passed into
'            CalculateChebyshevLeastSquares function.
'  -- ALGLIB --
'     Copyright by Bochkanov Sergey
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub BuildChebyshevLeastSquares(ByRef X() As Double, _
         ByRef Y() As Double, _
         ByRef W() As Double, _
         ByVal A As Double, _
         ByVal B As Double, _
         ByVal N As Long, _
         ByVal M As Long, _
         ByRef CTbl() As Double)
    Dim i As Long
    Dim J As Long
    Dim MA() As Double
    Dim Q() As Double
    Dim VT() As Double
    Dim MB() As Double
    Dim Tau() As Double
    Dim b2() As Double
    Dim TauQ() As Double
    Dim TauP() As Double
    Dim D() As Double
    Dim E() As Double
    Dim IsUpperA As Boolean
    Dim MI As Long
    Dim NI As Long
    Dim V As Double
    Dim i_ As Long

    MI = N
    NI = M + 1#
    
    '
    ' Initializing design matrix
    ' Here we are making MI>=NI
    '
    ReDim MA(1# To NI, 1# To MaxInt(MI, NI))
    ReDim MB(1# To MaxInt(MI, NI))
    For J = 1# To NI Step 1
        For i = 1# To MI Step 1
            V = 2# * (X(i - 1#) - A) / (B - A) - 1#
            If J = 1# Then
                MA(J, i) = 1#
            End If
            If J = 2# Then
                MA(J, i) = V
            End If
            If J > 2# Then
                MA(J, i) = 2# * V * MA(J - 1#, i) - MA(J - 2#, i)
            End If
        Next i
    Next J
    For J = 1# To NI Step 1
        For i = 1# To MI Step 1
            MA(J, i) = W(i - 1#) * MA(J, i)
        Next i
    Next J
    For J = 1# To NI Step 1
        For i = MI + 1# To NI Step 1
            MA(J, i) = 0#
        Next i
    Next J
    
    '
    ' Initializing right part
    '
    For i = 0# To MI - 1# Step 1
        MB(i + 1#) = W(i) * Y(i)
    Next i
    For i = MI + 1# To NI Step 1
        MB(i) = 0#
    Next i
    MI = MaxInt(MI, NI)
    
    '
    ' LQ-decomposition of A'
    ' B2 := Q*B
    '
    Call LQDecomposition(MA, NI, MI, Tau)
    Call UnpackQFromLQ(MA, NI, MI, Tau, NI, Q)
    ReDim b2(1# To 1#, 1# To NI)
    For J = 1# To NI Step 1
        b2(1#, J) = 0#
    Next J
    For i = 1# To NI Step 1
        V = 0#
        For i_ = 1# To MI Step 1
            V = V + MB(i_) * Q(i, i_)
        Next i_
        b2(1#, i) = V
    Next i
    
    '
    ' Back from A' to A
    ' Making cols(A)=rows(A)
    '
    For i = 1# To NI - 1# Step 1
        For i_ = i + 1# To NI Step 1
            MA(i, i_) = MA(i_, i)
        Next i_
    Next i
    For i = 2# To NI Step 1
        For J = 1# To i - 1# Step 1
            MA(i, J) = 0#
        Next J
    Next i
    
    '
    ' Bidiagonal decomposition of A
    ' A = Q * d2 * P'
    ' B2 := (Q'*B2')'
    '
    Call ToBidiagonal(MA, NI, NI, TauQ, TauP)
    Call MultiplyByQFromBidiagonal(MA, NI, NI, TauQ, b2, 1#, NI, True, False)
    Call UnpackPTFromBidiagonal(MA, NI, NI, TauP, NI, VT)
    Call UnpackDiagonalsFromBidiagonal(MA, NI, NI, IsUpperA, D, E)
    
    '
    ' Singular value decomposition of A
    ' A = U * d * V'
    ' B2 := (U'*B2')'
    '
    If Not BidiagonalSVDDecomposition(D, E, NI, IsUpperA, False, b2, 1#, Q, 0#, VT, NI) Then
        For i = 1# To NI Step 1
            D(i) = 0#
            b2(1#, i) = 0#
            For J = 1# To NI Step 1
                If i = J Then
                    VT(i, J) = 1#
                Else
                    VT(i, J) = 0#
                End If
            Next J
        Next i
        b2(1#, 1#) = 1#
    End If
    
    '
    ' B2 := (d^(-1) * B2')'
    '
    For i = 1# To NI Step 1
        If D(i) > MachineEpsilon * 10# * Sqr(NI) * D(1#) Then
            b2(1#, i) = b2(1#, i) / D(i)
        Else
            b2(1#, i) = 0#
        End If
    Next i
    
    '
    ' B := (V * B2')'
    '
    For i = 1# To NI Step 1
        MB(i) = 0#
    Next i
    For i = 1# To NI Step 1
        V = b2(1#, i)
        For i_ = 1# To NI Step 1
            MB(i_) = MB(i_) + V * VT(i, i_)
        Next i_
    Next i
    
    '
    ' Forming result
    '
    ReDim CTbl(0# To NI + 1#)
    For i = 1# To NI Step 1
        CTbl(i - 1#) = MB(i)
    Next i
    CTbl(NI) = A
    CTbl(NI + 1#) = B
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Weighted Chebyshev polynomial constrained least squares approximation.
'
'The algorithm reduces [A,B] to [-1,1] and builds the Chebyshev polynomials
'series by approximating a given function using the least squares method.
'
'Input parameters:
'    X   -   abscissas, array[0..N-1]
'    Y   -   function values, array[0..N-1]
'    W   -   weights, array[0..N-1].  Each  item  in  the  squared  sum  of
'            deviations from given values is  multiplied  by  a  square  of
'            corresponding weight.
'    A, B-   interval in which the approximating polynomials are built.
'    N   -   number of points, N>0.
'    XC, YC, DC-
'            constraints (see description below)., array[0..NC-1]
'    NC  -   number of constraints. 0 <= NC < M+1.
'    M   -   degree of polynomial, M>=0. This parameter is passed into  the
'            CalculateChebyshevLeastSquares subroutine.
'
'Output parameters:
'    CTbl-   coefficient  table.  This  parameter  is   passed   into   the
'            CalculateChebyshevLeastSquares subroutine.
'
'Result:
'    True, if the algorithm succeeded.
'    False, if the internal singular value decomposition subroutine  hasn't
'converged or the given constraints could not be met  simultaneously  (e.g.
'P(0)=0 è P(0)=1).
'
'Specifying constraints:
'    This subroutine can solve  the  problem  having  constrained  function
'values or its derivatives in several points. NC specifies  the  number  of
'constraints, DC - the type of constraints, XC and YC - constraints as such.
'Thus, for each i from 0 to NC-1 the following constraint is given:
'    P(xc[i]) = yc[i],       if DC[i]=0
'or
'    d/dx(P(xc[i])) = yc[i], if DC[i]=1
'(here P(x) is approximating polynomial).
'    This version of the subroutine supports only either polynomial or  its
'derivative value constraints.  If  DC[i]  is  not  equal  to  0 and 1, the
'subroutine will be aborted. The number of constraints should be less  than
'the number of degrees of freedom of approximating  polynomial  -  M+1  (at
'that, it could be equal to 0).
'
'  -- ALGLIB --
'     Copyright by Bochkanov Sergey
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function BuildChebyshevLeastSquaresConstrained(ByRef X() As Double, _
         ByRef Y() As Double, _
         ByRef W() As Double, _
         ByVal A As Double, _
         ByVal B As Double, _
         ByVal N As Long, _
         ByRef XC() As Double, _
         ByRef YC() As Double, _
         ByRef DC() As Long, _
         ByVal NC As Long, _
         ByVal M As Long, _
         ByRef CTbl() As Double) As Boolean
    Dim Result As Boolean
    Dim i As Long
    Dim J As Long
    Dim ReducedSize As Long
    Dim DesignMatrix() As Double
    Dim RightPart() As Double
    Dim CMatrix() As Double
    Dim C() As Double
    Dim U() As Double
    Dim VT() As Double
    Dim D() As Double
    Dim CR() As Double
    Dim WS() As Double
    Dim TJ() As Double
    Dim UJ() As Double
    Dim DTJ() As Double
    Dim Tmp() As Double
    Dim Tmp2() As Double
    Dim TmpMatrix() As Double
    Dim V As Double
    Dim i_ As Long

    Result = True
    
    '
    ' Initialize design matrix and right part.
    ' Add fictional rows if needed to ensure that N>=M+1.
    '
    ReDim DesignMatrix(1# To MaxInt(N, M + 1#), 1# To M + 1#)
    ReDim RightPart(1# To MaxInt(N, M + 1#))
    For i = 1# To N Step 1
        For J = 1# To M + 1# Step 1
            V = 2# * (X(i - 1#) - A) / (B - A) - 1#
            If J = 1# Then
                DesignMatrix(i, J) = 1#
            End If
            If J = 2# Then
                DesignMatrix(i, J) = V
            End If
            If J > 2# Then
                DesignMatrix(i, J) = 2# * V * DesignMatrix(i, J - 1#) - DesignMatrix(i, J - 2#)
            End If
        Next J
    Next i
    For i = 1# To N Step 1
        For J = 1# To M + 1# Step 1
            DesignMatrix(i, J) = W(i - 1#) * DesignMatrix(i, J)
        Next J
    Next i
    For i = N + 1# To M + 1# Step 1
        For J = 1# To M + 1# Step 1
            DesignMatrix(i, J) = 0#
        Next J
    Next i
    For i = 0# To N - 1# Step 1
        RightPart(i + 1#) = W(i) * Y(i)
    Next i
    For i = N + 1# To M + 1# Step 1
        RightPart(i) = 0#
    Next i
    N = MaxInt(N, M + 1#)
    
    '
    ' Now N>=M+1 and we are ready to the next stage.
    ' Handle constraints.
    ' Represent feasible set of coefficients as x = C*t + d
    '
    ReDim C(1# To M + 1#, 1# To M + 1#)
    ReDim D(1# To M + 1#)
    If NC = 0# Then
        
        '
        ' No constraints
        '
        For i = 1# To M + 1# Step 1
            For J = 1# To M + 1# Step 1
                C(i, J) = 0#
            Next J
            D(i) = 0#
        Next i
        For i = 1# To M + 1# Step 1
            C(i, i) = 1#
        Next i
        ReducedSize = M + 1#
    Else
        
        '
        ' Constraints are present.
        ' Fill constraints matrix CMatrix and solve CMatrix*x = cr.
        '
        ReDim CMatrix(1# To NC, 1# To M + 1#)
        ReDim CR(1# To NC)
        ReDim TJ(0# To M)
        ReDim UJ(0# To M)
        ReDim DTJ(0# To M)
        For i = 0# To NC - 1# Step 1
            V = 2# * (XC(i) - A) / (B - A) - 1#
            For J = 0# To M Step 1
                If J = 0# Then
                    TJ(J) = 1#
                    UJ(J) = 1#
                    DTJ(J) = 0#
                End If
                If J = 1# Then
                    TJ(J) = V
                    UJ(J) = 2# * V
                    DTJ(J) = 1#
                End If
                If J > 1# Then
                    TJ(J) = 2# * V * TJ(J - 1#) - TJ(J - 2#)
                    UJ(J) = 2# * V * UJ(J - 1#) - UJ(J - 2#)
                    DTJ(J) = J * UJ(J - 1#)
                End If
                If DC(i) = 0# Then
                    CMatrix(i + 1#, J + 1#) = TJ(J)
                End If
                If DC(i) = 1# Then
                    CMatrix(i + 1#, J + 1#) = DTJ(J)
                End If
            Next J
            CR(i + 1#) = YC(i)
        Next i
        
        '
        ' Solve CMatrix*x = cr.
        ' Fill C and d:
        ' 1. SVD: CMatrix = U * WS * V^T
        ' 2. C := V[1:M+1,NC+1:M+1]
        ' 3. tmp := WS^-1 * U^T * cr
        ' 4. d := V[1:M+1,1:NC] * tmp
        '
        If Not SVDDecomposition(CMatrix, NC, M + 1#, 2#, 2#, 2#, WS, U, VT) Then
            Result = False
            BuildChebyshevLeastSquaresConstrained = Result
            Exit Function
        End If
        If WS(1#) = 0# Or WS(NC) <= MachineEpsilon * 10# * Sqr(NC) * WS(1#) Then
            Result = False
            BuildChebyshevLeastSquaresConstrained = Result
            Exit Function
        End If
        ReDim C(1# To M + 1#, 1# To M + 1# - NC)
        ReDim D(1# To M + 1#)
        For i = 1# To M + 1# - NC Step 1
            For i_ = 1# To M + 1# Step 1
                C(i_, i) = VT(NC + i, i_)
            Next i_
        Next i
        ReDim Tmp(1# To NC)
        For i = 1# To NC Step 1
            V = 0#
            For i_ = 1# To NC Step 1
                V = V + U(i_, i) * CR(i_)
            Next i_
            Tmp(i) = V / WS(i)
        Next i
        For i = 1# To M + 1# Step 1
            D(i) = 0#
        Next i
        For i = 1# To NC Step 1
            V = Tmp(i)
            For i_ = 1# To M + 1# Step 1
                D(i_) = D(i_) + V * VT(i, i_)
            Next i_
        Next i
        
        '
        ' Reduce problem:
        ' 1. RightPart := RightPart - DesignMatrix*d
        ' 2. DesignMatrix := DesignMatrix*C
        '
        For i = 1# To N Step 1
            V = 0#
            For i_ = 1# To M + 1# Step 1
                V = V + DesignMatrix(i, i_) * D(i_)
            Next i_
            RightPart(i) = RightPart(i) - V
        Next i
        ReducedSize = M + 1# - NC
        ReDim TmpMatrix(1# To N, 1# To ReducedSize)
        ReDim Tmp(1# To N)
        Call MatrixMatrixMultiply(DesignMatrix, 1#, N, 1#, M + 1#, False, C, 1#, M + 1#, 1#, ReducedSize, False, 1#, TmpMatrix, 1#, N, 1#, ReducedSize, 0#, Tmp)
        Call CopyMatrix(TmpMatrix, 1#, N, 1#, ReducedSize, DesignMatrix, 1#, N, 1#, ReducedSize)
    End If
    
    '
    ' Solve reduced problem DesignMatrix*t = RightPart.
    '
    If Not SVDDecomposition(DesignMatrix, N, ReducedSize, 1#, 1#, 2#, WS, U, VT) Then
        Result = False
        BuildChebyshevLeastSquaresConstrained = Result
        Exit Function
    End If
    ReDim Tmp(1# To ReducedSize)
    ReDim Tmp2(1# To ReducedSize)
    For i = 1# To ReducedSize Step 1
        Tmp(i) = 0#
    Next i
    For i = 1# To N Step 1
        V = RightPart(i)
        For i_ = 1# To ReducedSize Step 1
            Tmp(i_) = Tmp(i_) + V * U(i, i_)
        Next i_
    Next i
    For i = 1# To ReducedSize Step 1
        If WS(i) <> 0# And WS(i) > MachineEpsilon * 10# * Sqr(NC) * WS(1#) Then
            Tmp(i) = Tmp(i) / WS(i)
        Else
            Tmp(i) = 0#
        End If
    Next i
    For i = 1# To ReducedSize Step 1
        Tmp2(i) = 0#
    Next i
    For i = 1# To ReducedSize Step 1
        V = Tmp(i)
        For i_ = 1# To ReducedSize Step 1
            Tmp2(i_) = Tmp2(i_) + V * VT(i, i_)
        Next i_
    Next i
    
    '
    ' Solution is in the tmp2.
    ' Transform it from t to x.
    '
    ReDim CTbl(0# To M + 2#)
    For i = 1# To M + 1# Step 1
        V = 0#
        For i_ = 1# To ReducedSize Step 1
            V = V + C(i, i_) * Tmp2(i_)
        Next i_
        CTbl(i - 1#) = V + D(i)
    Next i
    CTbl(M + 1#) = A
    CTbl(M + 2#) = B

    BuildChebyshevLeastSquaresConstrained = Result
End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Calculation of a Chebyshev  polynomial  obtained   during  least  squares
'approximaion at the given point.
'
'Input parameters:
'    M   -   order of polynomial (parameter of the
'            BuildChebyshevLeastSquares function).
'    A   -   coefficient table.
'            A[0..M] contains coefficients of the i-th Chebyshev polynomial.
'            A[M+1] contains left boundary of approximation interval.
'            A[M+2] contains right boundary of approximation interval.
'    X   -   point to perform calculations in.
'
'The result is the value at the given point.
'
'It should be noted that array A contains coefficients  of  the  Chebyshev
'polynomials defined on interval [-1,1].   Argument  is  reduced  to  this
'interval before calculating polynomial value.
'  -- ALGLIB --
'     Copyright by Bochkanov Sergey
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CalculateChebyshevLeastSquares(ByRef M As Long, _
         ByRef A() As Double, _
         ByVal X As Double) As Double
    Dim Result As Double
    Dim b1 As Double
    Dim b2 As Double
    Dim i As Long

    X = 2# * (X - A(M + 1#)) / (A(M + 2#) - A(M + 1#)) - 1#
    b1 = 0#
    b2 = 0#
    i = M
    Do
        Result = 2# * X * b1 - b2 + A(i)
        b2 = b1
        b1 = Result
        i = i - 1#
    Loop Until Not i >= 0#
    Result = Result - X * b2

    CalculateChebyshevLeastSquares = Result
End Function


