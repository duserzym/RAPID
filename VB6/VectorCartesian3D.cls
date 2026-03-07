VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Cartesian3D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Public X As Double
Public Y As Double
Public Z As Double

Const Pi As Double = 3.141592653589
Const rad As Double = (Pi / 180#)
Const deg As Double = (180# / Pi)

Function atan(ByVal X As Double, ByVal Y As Double) As Double
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

Public Property Get dec() As Double
    dec = atan(X, Y) * deg
End Property

Public Property Get inc() As Double
    Dim p As Double
    p = X ^ 2 + Y ^ 2
    inc = atan(Sqr(p), Z) * deg
End Property

Public Property Get mag() As Double
    mag = Sqr(X ^ 2 + Y ^ 2 + Z ^ 2)
End Property

Public Property Get UnitVectorX() As Double
    Dim M As Double
    M = mag
    
    If M = 0 Then
        UnitVectorX = 1
    Else
        UnitVectorX = X / M
    End If
End Property

Public Property Get UnitVectorY() As Double
    Dim M As Double
    M = mag
    
    If M = 0 Then
        UnitVectorY = 0
    Else
        UnitVectorY = Y / M
    End If
End Property

Public Property Get UnitVectorZ() As Double
    Dim M As Double
    M = mag
    
    If M = 0 Then
        UnitVectorZ = 0
    Else
        UnitVectorZ = Z / M
    End If
End Property

Public Sub ReadFromINI(sSection As String, sKey As String, sIniFile As String)
    Dim parseArray() As String
    parseArray = Split(Config_GetFromINI(sSection, sKey, "0,0,0", sIniFile), ",")
    X = val(parseArray(0))
    Y = val(parseArray(1))
    Z = val(parseArray(2))
End Sub

Public Sub ReadString(ByVal cartesianString As String)
    Dim parseArray() As String
    parseArray = Split(cartesianString, ",")
    X = val(parseArray(0))
    Y = val(parseArray(1))
    Z = val(parseArray(2))
End Sub

Public Function WriteString() As String
    Dim xStr As String, yStr As String, zStr As String
    
    xStr = Format$(X, "0.0000000E+")
    yStr = Format$(Y, "0.0000000E+")
    zStr = Format$(Z, "0.0000000E+")
    WriteString = xStr & "," & yStr & "," & zStr
End Function

Public Sub WriteToINI(sSection As String, sKey As String, sIniFile As String)
    Dim xStr As String, yStr As String, zStr As String
    
    xStr = Format$(X, "0.0000000E+")
    yStr = Format$(Y, "0.0000000E+")
    zStr = Format$(Z, "0.0000000E+")
    Config_AddToINI sSection, sKey, xStr & "," & yStr & "," & zStr, sIniFile
End Sub

