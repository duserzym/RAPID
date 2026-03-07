Attribute VB_Name = "Module1"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Copyright (c) 1992-2007 The University of Tennessee.  All rights reserved.
'
'Contributors:
'    * Sergey Bochkanov (ALGLIB project). Translation from FORTRAN to
'      pseudocode.
'
'See subroutines comments for additional copyrights.
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
'Routines
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'QR decomposition of a rectangular matrix of size MxN
'
'Input parameters:
'    A   -   matrix A whose indexes range within [0..M-1, 0..N-1].
'    M   -   number of rows in matrix A.
'    N   -   number of columns in matrix A.
'
'Output parameters:
'    A   -   matrices Q and R in compact form (see below).
'    Tau -   array of scalar factors which are used to form
'            matrix Q. Array whose index ranges within [0.. Min(M-1,N-1)].
'
'Matrix A is represented as A = QR, where Q is an orthogonal matrix of size
'MxM, R - upper triangular (or upper trapezoid) matrix of size M x N.
'
'The elements of matrix R are located on and above the main diagonal of
'matrix A. The elements which are located in Tau array and below the main
'diagonal of matrix A are used to form matrix Q as follows:
'
'Matrix Q is represented as a product of elementary reflections
'
'Q = H(0)*H(2)*...*H(k-1),
'
'where k = min(m,n), and each H(i) is in the form
'
'H(i) = 1 - tau * v * (v^T)
'
'where tau is a scalar stored in Tau[I]; v - real vector,
'so that v(0:i-1) = 0, v(i) = 1, v(i+1:m-1) stored in A(i+1:m-1,i).
'
'  -- LAPACK routine (version 3.0) --
'     Univ. of Tennessee, Univ. of California Berkeley, NAG Ltd.,
'     Courant Institute, Argonne National Lab, and Rice University
'     February 29, 1992.
'     Translation from FORTRAN to pseudocode (AlgoPascal)
'     by Sergey Bochkanov, ALGLIB project, 2005-2007.
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RMatrixQR(ByRef A() As Double, _
         ByVal M As Long, _
         ByVal N As Long, _
         ByRef Tau() As Double)
    Dim WORK() As Double
    Dim T() As Double
    Dim I As Long
    Dim K As Long
    Dim MinMN As Long
    Dim Tmp As Double
    Dim i_ As Long
    Dim i1_ As Long

    If M <= 0# Or N <= 0# Then
        Exit Sub
    End If
    MinMN = MinInt(M, N)
    ReDim WORK(0# To N - 1#)
    ReDim T(1# To M)
    ReDim Tau(0# To MinMN - 1#)
    
    '
    ' Test the input arguments
    '
    K = MinMN
    For I = 0# To K - 1# Step 1
        
        '
        ' Generate elementary reflector H(i) to annihilate A(i+1:m,i)
        '
        i1_ = (I) - (1#)
        For i_ = 1# To M - I Step 1
            T(i_) = A(i_ + i1_, I)
        Next i_
        Call GenerateReflection(T, M - I, Tmp)
        Tau(I) = Tmp
        i1_ = (1#) - (I)
        For i_ = I To M - 1# Step 1
            A(i_, I) = T(i_ + i1_)
        Next i_
        T(1#) = 1#
        If I < N Then
            
            '
            ' Apply H(i) to A(i:m-1,i+1:n-1) from the left
            '
            Call ApplyReflectionFromTheLeft(A, Tau(I), T, I, M - 1#, I + 1#, N - 1#, WORK)
        End If
    Next I
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Partial unpacking of matrix Q from the QR decomposition of a matrix A
'
'Input parameters:
'    A       -   matrices Q and R in compact form.
'                Output of RMatrixQR subroutine.
'    M       -   number of rows in given matrix A. M>=0.
'    N       -   number of columns in given matrix A. N>=0.
'    Tau     -   scalar factors which are used to form Q.
'                Output of the RMatrixQR subroutine.
'    QColumns -  required number of columns of matrix Q. M>=QColumns>=0.
'
'Output parameters:
'    Q       -   first QColumns columns of matrix Q.
'                Array whose indexes range within [0..M-1, 0..QColumns-1].
'                If QColumns=0, the array remains unchanged.
'
'  -- ALGLIB --
'     Copyright 2005 by Bochkanov Sergey
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RMatrixQRUnpackQ(ByRef A() As Double, _
         ByVal M As Long, _
         ByVal N As Long, _
         ByRef Tau() As Double, _
         ByVal QColumns As Long, _
         ByRef Q() As Double)
    Dim I As Long
    Dim J As Long
    Dim K As Long
    Dim MinMN As Long
    Dim V() As Double
    Dim WORK() As Double
    Dim i_ As Long
    Dim i1_ As Long

    If M <= 0# Or N <= 0# Or QColumns <= 0# Then
        Exit Sub
    End If
    
    '
    ' init
    '
    MinMN = MinInt(M, N)
    K = MinInt(MinMN, QColumns)
    ReDim Q(0# To M - 1#, 0# To QColumns - 1#)
    ReDim V(1# To M)
    ReDim WORK(0# To QColumns - 1#)
    For I = 0# To M - 1# Step 1
        For J = 0# To QColumns - 1# Step 1
            If I = J Then
                Q(I, J) = 1#
            Else
                Q(I, J) = 0#
            End If
        Next J
    Next I
    
    '
    ' unpack Q
    '
    For I = K - 1# To 0# Step -1
        
        '
        ' Apply H(i)
        '
        i1_ = (I) - (1#)
        For i_ = 1# To M - I Step 1
            V(i_) = A(i_ + i1_, I)
        Next i_
        V(1#) = 1#
        Call ApplyReflectionFromTheLeft(Q, Tau(I), V, I, M - 1#, 0#, QColumns - 1#, WORK)
    Next I
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Unpacking of matrix R from the QR decomposition of a matrix A
'
'Input parameters:
'    A       -   matrices Q and R in compact form.
'                Output of RMatrixQR subroutine.
'    M       -   number of rows in given matrix A. M>=0.
'    N       -   number of columns in given matrix A. N>=0.
'
'Output parameters:
'    R       -   matrix R, array[0..M-1, 0..N-1].
'
'  -- ALGLIB --
'     Copyright 2005 by Bochkanov Sergey
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RMatrixQRUnpackR(ByRef A() As Double, _
         ByVal M As Long, _
         ByVal N As Long, _
         ByRef R() As Double)
    Dim I As Long
    Dim K As Long
    Dim i_ As Long

    If M <= 0# Or N <= 0# Then
        Exit Sub
    End If
    K = MinInt(M, N)
    ReDim R(0# To M - 1#, 0# To N - 1#)
    For I = 0# To N - 1# Step 1
        R(0#, I) = 0#
    Next I
    For I = 1# To M - 1# Step 1
        For i_ = 0# To N - 1# Step 1
            R(I, i_) = R(0#, i_)
        Next i_
    Next I
    For I = 0# To K - 1# Step 1
        For i_ = I To N - 1# Step 1
            R(I, i_) = A(I, i_)
        Next i_
    Next I
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Obsolete 1-based subroutine. See RMatrixQR for 0-based replacement.
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub QRDecomposition(ByRef A() As Double, _
         ByVal M As Long, _
         ByVal N As Long, _
         ByRef Tau() As Double)
    Dim WORK() As Double
    Dim T() As Double
    Dim I As Long
    Dim K As Long
    Dim MMIP1 As Long
    Dim MinMN As Long
    Dim Tmp As Double
    Dim i_ As Long
    Dim i1_ As Long

    MinMN = MinInt(M, N)
    ReDim WORK(1# To N)
    ReDim T(1# To M)
    ReDim Tau(1# To MinMN)
    
    '
    ' Test the input arguments
    '
    K = MinInt(M, N)
    For I = 1# To K Step 1
        
        '
        ' Generate elementary reflector H(i) to annihilate A(i+1:m,i)
        '
        MMIP1 = M - I + 1#
        i1_ = (I) - (1#)
        For i_ = 1# To MMIP1 Step 1
            T(i_) = A(i_ + i1_, I)
        Next i_
        Call GenerateReflection(T, MMIP1, Tmp)
        Tau(I) = Tmp
        i1_ = (1#) - (I)
        For i_ = I To M Step 1
            A(i_, I) = T(i_ + i1_)
        Next i_
        T(1#) = 1#
        If I < N Then
            
            '
            ' Apply H(i) to A(i:m,i+1:n) from the left
            '
            Call ApplyReflectionFromTheLeft(A, Tau(I), T, I, M, I + 1#, N, WORK)
        End If
    Next I
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Obsolete 1-based subroutine. See RMatrixQRUnpackQ for 0-based replacement.
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub UnpackQFromQR(ByRef A() As Double, _
         ByVal M As Long, _
         ByVal N As Long, _
         ByRef Tau() As Double, _
         ByVal QColumns As Long, _
         ByRef Q() As Double)
    Dim I As Long
    Dim J As Long
    Dim K As Long
    Dim MinMN As Long
    Dim V() As Double
    Dim WORK() As Double
    Dim VM As Long
    Dim i_ As Long
    Dim i1_ As Long

    If M = 0# Or N = 0# Or QColumns = 0# Then
        Exit Sub
    End If
    
    '
    ' init
    '
    MinMN = MinInt(M, N)
    K = MinInt(MinMN, QColumns)
    ReDim Q(1# To M, 1# To QColumns)
    ReDim V(1# To M)
    ReDim WORK(1# To QColumns)
    For I = 1# To M Step 1
        For J = 1# To QColumns Step 1
            If I = J Then
                Q(I, J) = 1#
            Else
                Q(I, J) = 0#
            End If
        Next J
    Next I
    
    '
    ' unpack Q
    '
    For I = K To 1# Step -1
        
        '
        ' Apply H(i)
        '
        VM = M - I + 1#
        i1_ = (I) - (1#)
        For i_ = 1# To VM Step 1
            V(i_) = A(i_ + i1_, I)
        Next i_
        V(1#) = 1#
        Call ApplyReflectionFromTheLeft(Q, Tau(I), V, I, M, 1#, QColumns, WORK)
    Next I
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Obsolete 1-based subroutine. See RMatrixQR for 0-based replacement.
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub QRDecompositionUnpacked(ByRef A_() As Double, _
         ByVal M As Long, _
         ByVal N As Long, _
         ByRef Q() As Double, _
         ByRef R() As Double)
    Dim A() As Double
    Dim I As Long
    Dim K As Long
    Dim Tau() As Double
    Dim WORK() As Double
    Dim V() As Double
    Dim i_ As Long
    A = A_

    K = MinInt(M, N)
    If N <= 0# Then
        Exit Sub
    End If
    ReDim WORK(1# To M)
    ReDim V(1# To M)
    ReDim Q(1# To M, 1# To M)
    ReDim R(1# To M, 1# To N)
    
    '
    ' QRDecomposition
    '
    Call QRDecomposition(A, M, N, Tau)
    
    '
    ' R
    '
    For I = 1# To N Step 1
        R(1#, I) = 0#
    Next I
    For I = 2# To M Step 1
        For i_ = 1# To N Step 1
            R(I, i_) = R(1#, i_)
        Next i_
    Next I
    For I = 1# To K Step 1
        For i_ = I To N Step 1
            R(I, i_) = A(I, i_)
        Next i_
    Next I
    
    '
    ' Q
    '
    Call UnpackQFromQR(A, M, N, Tau, M, Q)
End Sub


