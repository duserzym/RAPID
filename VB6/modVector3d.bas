Attribute VB_Name = "modVector3d"
' this is the 3d Vector code module

Option Explicit

'Global Const Pi As Double = 3.1415926536
Global Const rad As Double = (Pi / 180#)
Global Const deg As Double = (180# / Pi)

' note the (north,east,down) to (x,y,z) mapping
' agrees with spacial orientation and
' coordinate conventions

Global Const c_N = 0     ' north
Global Const c_E = 1    ' east
Global Const c_D = 2    ' down

Global Const c_x = 0
Global Const c_y = 1
Global Const c_z = 2

Function acos(theta As Double) As Double
    acos = Atn(-theta / Sqr(-theta ^ 2 + 1)) + Pi / 2
End Function

Sub Angular3dToCartesian3d(g As Angular3D, c As Cartesian3D)
    Dim p As Double
    p = Cos(g.inc * rad)
    c.X = g.mag * p * Cos(g.dec * rad)
    c.Y = g.mag * p * Sin(g.dec * rad)
    c.Z = g.mag * Sin(g.inc * rad)
End Sub

Public Function arcos(ByVal ar As Double)
    ' if ar is between -1 AND 1, return the corresponding arccos
    ' in radians

    If ar >= 1 Then arcos = 0: Return
    If ar <= -1 Then arcos = Pi: Return
    arcos = Atan2(ar, Sqr(1 - ar ^ 2))
End Function

Function atan(X As Double, Y As Double) As Double
    If X > 0 Then
        If Y > 0 Then
            atan = Atn(Y / X)
        Else
            atan = 2 * Pi + Atn(Y / X)
        End If
    ElseIf X < 0 Then
        atan = Pi + Atn(Y / X)
    Else
        If Y > 0 Then
            atan = Pi
        ElseIf Y < 0 Then
            atan = -Pi
        Else
            atan = 0#
        End If
    End If
End Function

Public Function Atan2(ByVal xC As Double, ByVal yC As Double) As Double
    ' Returns the angle from x clockwise to y, in radians.
    ' "xc" is in the x direction, "yc" is in the y direction.

    If (xC = 0 And yC >= 0) Then
        Atan2 = Pi / 2
    ElseIf (xC = 0 And yC < 0) Then
        Atan2 = 3 * Pi / 2
    Else
        Atan2 = Atn(Abs(yC / xC))
        If (yC >= 0 And xC > 0) Then
        ElseIf (yC >= 0 And xC < 0) Then
            Atan2 = Pi - Atan2
        ElseIf (yC <= 0 And xC < 0) Then
            Atan2 = Pi + Atan2
        ElseIf (yC <= 0 And xC > 0) Then
            Atan2 = 2 * Pi - Atan2
        End If
    End If
End Function

Sub cartesian3d_average(v1 As Cartesian3D, v2 As Cartesian3D, v As Cartesian3D)
    v.X = (v1.X + v2.X) / 2#
    v.Y = (v1.Y + v2.Y) / 2#
    v.Z = (v1.Z + v2.Z) / 2#
End Sub

Function Cartesian3d_DiffAngle(v1 As Cartesian3D, v2 As Cartesian3D) As Double
    Cartesian3d_DiffAngle = deg * acos(Cartesian3d_DotProduct(v1, v2) / v1.mag / v2.mag)
End Function

Sub Cartesian3d_Difference(v1 As Cartesian3D, v2 As Cartesian3D, v As Cartesian3D)
    v.X = v1.X - v2.X
    v.Y = v1.Y - v2.Y
    v.Z = v1.Z - v2.Z
End Sub

Function Cartesian3d_DotProduct(v1 As Cartesian3D, v2 As Cartesian3D) As Double
    Cartesian3d_DotProduct = (v1.X * v2.X) + (v1.Y * v2.Y) + (v1.Z * v2.Z)
End Function

' v1 : the vector to rotate
' axis : the axis to rotate around
' theta : the angle to rotate in radians
' v : the resultant vector
Sub Cartesian3d_Rotate(v1 As Cartesian3D, axis As Integer, theta As Double, v As Cartesian3D)
    Dim a0 As Integer, A1 As Integer   ' the axis that will be changed
    Dim c0 As Double, C1 As Double    ' temp values for the coordinates
    Dim ct As Double, st As Double    ' sin and cos of theta

    'Select Case axis
    'Case c_x
    '    a0 = c_y: a1 = c_z  ' rotate yz plane around x
    'Case c_y
    '    a0 = c_z: a1 = c_x  ' rotate zx plane around y
    'Case c_z
    '    a0 = c_x: a1 = c_y  ' rotate xy plane around z
    'End Select
    
    'ct = Cos(theta * rad): st = Sin(theta * rad)
    ' store temp values just in case v1=v
    'c0 = v1.C(a0): c1 = v1.C(a1)
    ' this is just standard trig
    'v.C(axis) = v1.C(axis)
    'v.C(a0) = c0 * ct + c1 * st
    'v.C(a1) = c1 * ct - c0 * st
End Sub

Function Cartesian3d_ScalarDiv(v1 As Cartesian3D, r As Double) As Cartesian3D
    Set Cartesian3d_ScalarDiv = New Cartesian3D

    Cartesian3d_ScalarDiv.X = v1.X / r
    Cartesian3d_ScalarDiv.Y = v1.Y / r
    Cartesian3d_ScalarDiv.Z = v1.Z / r
End Function

Sub Cartesian3d_ScalarMult(v1 As Cartesian3D, r As Double, v As Cartesian3D)
    v.X = r * v1.X
    v.Y = r * v1.Y
    v.Z = r * v1.Z
End Sub

Sub Cartesian3d_Square(v1 As Cartesian3D, v As Cartesian3D)
    v.X = v1.X ^ 2
    v.Y = v1.Y ^ 2
    v.Z = v1.Z ^ 2
End Sub

Sub Cartesian3d_SquareRoot(v1 As Cartesian3D, v As Cartesian3D)
    v.X = Sqr(v1.X)
    v.Y = Sqr(v1.Y)
    v.Z = Sqr(v1.Z)
End Sub

Function Cartesian3d_Sum(v1 As Cartesian3D, v2 As Cartesian3D) As Cartesian3D
    Set Cartesian3d_Sum = New Cartesian3D
    Cartesian3d_Sum.X = v1.X + v2.X
    Cartesian3d_Sum.Y = v1.Y + v2.Y
    Cartesian3d_Sum.Z = v1.Z + v2.Z
End Function

Sub Cartesian3d_Zero(v As Cartesian3D)
    v.X = 0
    v.Y = 0
    v.Z = 0
End Sub

Sub Cartesian3dToAngular3d(c As Cartesian3D, g As Angular3D)
    Dim S As Double
    S = c.X ^ 2 + c.Y ^ 2
    g.mag = Sqr(S + c.Z ^ 2)
    g.inc = deg * atan(Sqr(S), c.Z)
    If g.inc > 180 Then g.inc = g.inc - 360
    g.dec = deg * atan(c.X, c.Y)
End Sub

Function DegToRad(ang As Double) As Double
    ' This function returns an angle in radians from some degree value
    DegToRad = ang * rad
End Function

Function RadToDeg(ang As Double, Optional chVal As Boolean = False) As Double
    ' This function returns an angle in degrees from some radian value
    RadToDeg = ang * deg
    If chVal Then
        ' We always want to return the declination between 0 and 360
        While (RadToDeg > 360#)
            RadToDeg = RadToDeg - 360
        Wend
        While (RadToDeg < 0)
            RadToDeg = RadToDeg + 360
        Wend
    End If
End Function

