VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MeasurementBlock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Sample"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' This data type stores all the data fields measured in one
' pass of measurement in the magnetometer.
    
' Baseline measurements
Private mvarBaselines(1 To 2) As Cartesian3D
    
' Sample measurements
Private mvarSample(1 To 4) As Cartesian3D
Private mvarHolder(1 To 4) As Cartesian3D
Private mvarDirection As Integer

Public Property Get Average() As Cartesian3D
    Dim i As Integer
    Dim workingvector As Cartesian3D
    
    Set workingvector = Sum
    Set Average = New Cartesian3D
    
    
    With Average
        .X = workingvector.X / 4
        .Y = workingvector.Y / 4
        .Z = workingvector.Z / 4
    End With
    Set workingvector = Nothing
End Property

Public Property Get AverageMagnitude() As Double
    Dim workingvector As Cartesian3D
    
    Set workingvector = Average
    With workingvector
        AverageMagnitude = Sqr(.X ^ 2 + .Y ^ 2 + .Z ^ 2)
    End With
    Set workingvector = Nothing
End Property

Public Property Get BaselineAdjustedSample(ByVal i As Integer) As Cartesian3D
    Dim BaselineFactor(1 To 2) As Single
    
    BaselineFactor(1) = 1 - i / 5
    BaselineFactor(2) = i / 5

    If i < 1 Or i > 4 Then Exit Property
    
    Set BaselineAdjustedSample = New Cartesian3D
    With BaselineAdjustedSample
        .X = mvarSample(i).X - BaselineFactor(1) * mvarBaselines(1).X - BaselineFactor(2) * mvarBaselines(2).X - mvarHolder(i).X
        .Y = mvarSample(i).Y - BaselineFactor(1) * mvarBaselines(1).Y - BaselineFactor(2) * mvarBaselines(2).Y - mvarHolder(i).Y
        .Z = mvarSample(i).Z - BaselineFactor(1) * mvarBaselines(1).Z - BaselineFactor(2) * mvarBaselines(2).Z - mvarHolder(i).Z
    End With
End Property

Public Property Get Baselines(ByVal i As Integer) As Cartesian3D
    Set Baselines = mvarBaselines(i)
End Property

Public Property Get CorrectedSample(ByVal i As Integer) As Cartesian3D
    ' returns sample as corrected for angle and baseline adjusted
    Dim X As Double, Y As Double, Z As Double
    Dim workingvector As Cartesian3D
    
    Set workingvector = BaselineAdjustedSample(i)
    Set CorrectedSample = New Cartesian3D
    With CorrectedSample
        X = workingvector.X
        Y = workingvector.Y
        Z = workingvector.Z
    Select Case i
        Case 1
            .X = X
            .Y = -Y * mvarDirection
            .Z = Z * mvarDirection
        Case 2
            .X = Y
            .Y = X * mvarDirection
            .Z = Z * mvarDirection
        Case 3
            .X = -X
            .Y = Y * mvarDirection
            .Z = Z * mvarDirection
        Case 4
            .X = -Y
            .Y = -X * mvarDirection
            .Z = Z * mvarDirection
    End Select
    End With
    
    Set workingvector = Nothing
End Property

Public Property Get driftc() As Cartesian3D
    Set driftc = New Cartesian3D
    driftc.X = mvarBaselines(2).X - mvarBaselines(1).X
    driftc.Y = mvarBaselines(2).Y - mvarBaselines(1).Y
    driftc.Z = mvarBaselines(2).Z - mvarBaselines(1).Z
End Property

Public Property Get FischerSD() As Double
    ' returns theta-63
    ' this is circular SD from Fischer distribution
    
    FischerSD = 81 / Sqr(Kappa)
End Property

Public Property Get Holder(ByVal i As Integer) As Cartesian3D
    If i < 1 Or i > 4 Then Exit Property
    Set Holder = mvarHolder(i)
End Property

Public Property Get HolderFVal() As Double
    HolderFVal = AverageHolder.mag
End Property

Public Property Get induced() As Cartesian3D
    Dim i As Integer
    Dim workingvector As Cartesian3D
    Set induced = New Cartesian3D
    With induced
        For i = 1 To 4
            Set workingvector = BaselineAdjustedSample(i)
            .X = .X + workingvector.X / 4
            .Y = .Y + workingvector.Y / 4
            If i > 2 Then
                .Z = .Z - workingvector.Z / 4
            Else
                .Z = .Z + workingvector.Z / 4
            End If
            Set workingvector = Nothing
        Next i
    End With
End Property

Public Property Get isUp() As Boolean
    isUp = (mvarDirection = 1)
End Property

Public Property Get Kappa() As Double
    Dim N As Integer
    Dim r As Double
    
    N = 4
    r = SumUnitVectors.mag
    If N = r Then
        Kappa = 0.000000001
    Else
        Kappa = (N - 1) / (N - r)
    End If
End Property

Public Property Get Sample(ByVal i As Integer) As Cartesian3D
    Set Sample = mvarSample(i)
End Property

Public Property Get SigDrift() As Double
    ' SIGDRIFT  is sample moment/baseline drifts
    Dim DriftMag As Double
    
    DriftMag = driftc.mag
    If DriftMag = 0 Then DriftMag = 0.000000001
    SigDrift = AverageMagnitude / DriftMag
End Property

Public Property Get SigHolder() As Double
    ' SIGHOLDER is sample moment/holder moment
    Dim HolderMag As Double
    HolderMag = HolderFVal
    If HolderMag = 0 Then HolderMag = 0.000000001
    
    SigHolder = AverageMagnitude / HolderMag
End Property

Public Property Get SigInduced() As Double
    ' SIGINDUCED measures rotational asymmetry of X,Y, & Z
    Dim inducedMag As Double
    
    inducedMag = induced.mag
    If inducedMag = 0 Then inducedMag = 0.000000001
    SigInduced = AverageMagnitude / inducedMag
End Property

Public Property Get Sum() As Cartesian3D
    Dim i As Integer
    Dim workingvector As Cartesian3D
    Set Sum = New Cartesian3D
    
    With Sum
    For i = 1 To 4
        Set workingvector = CorrectedSample(i)
        .X = .X + workingvector.X
        .Y = .Y + workingvector.Y
        .Z = .Z + workingvector.Z
    Next i
    End With
    Set workingvector = Nothing

End Property

Public Property Get SumSqs() As Cartesian3D
    Dim i As Integer
    Dim workingvector As Cartesian3D
    Set SumSqs = New Cartesian3D
    
    With SumSqs
    For i = 1 To 4
        Set workingvector = CorrectedSample(i)
        .X = .X + workingvector.X ^ 2
        .Y = .Y + workingvector.Y ^ 2
        .Z = .Z + workingvector.Z ^ 2
    Next i
    End With
    Set workingvector = Nothing

End Property

Public Property Get SumUnitVectors() As Cartesian3D
    Dim i As Integer
    Dim workingvector As Cartesian3D
    Set SumUnitVectors = New Cartesian3D
    
    For i = 1 To 4
        Set workingvector = CorrectedSample(i)
        SumUnitVectors.X = SumUnitVectors.X + workingvector.UnitVectorX
        SumUnitVectors.Y = SumUnitVectors.Y + workingvector.UnitVectorY
        SumUnitVectors.Z = SumUnitVectors.Z + workingvector.UnitVectorZ
    Next i
    Set workingvector = Nothing

End Property

Public Property Let isUp(ByVal vData As Boolean)
    If vData Then
        mvarDirection = 1
    Else
        mvarDirection = -1
    End If
End Property

Private Function AverageHolder() As Cartesian3D
    Dim i As Integer
    Dim workingvector As Cartesian3D
    Set AverageHolder = New Cartesian3D
    
    With AverageHolder
    For i = 1 To 4
        Set workingvector = CorrectedHolder(i)
        .X = .X + workingvector.X / 4
        .Y = .Y + workingvector.Y / 4
        .Z = .Z + workingvector.Z / 4
    Next i
    End With
    Set workingvector = Nothing
End Function

Private Function CorrectedHolder(ByVal i As Integer) As Cartesian3D
    ' returns sample as corrected for angle and baseline adjusted
    Dim X As Double, Y As Double, Z As Double
    X = mvarHolder(i).X
    Y = mvarHolder(i).Y
    Z = mvarHolder(i).Z
    Set CorrectedHolder = New Cartesian3D
    With CorrectedHolder
    Select Case i
        Case 1
            .X = X
            .Y = -Y
            .Z = Z
        Case 2
            .X = Y
            .Y = X
            .Z = Z
        Case 3
            .X = -X
            .Y = Y
            .Z = Z
        Case 4
            .X = -Y
            .Y = -X
            .Z = Z
    End Select
    End With
End Function

Private Sub Class_Initialize()
    'creates the collection when this class is created
    Dim i As Integer
    For i = 1 To 4
        Set mvarSample(i) = New Cartesian3D
        Set mvarHolder(i) = New Cartesian3D
    Next i
    For i = 1 To 2
        Set mvarBaselines(i) = New Cartesian3D
    Next i
    mvarDirection = 1
End Sub

Private Sub Class_Terminate()
    'creates the collection when this class is created
    Dim i As Integer
    For i = 1 To 4
        Set mvarSample(i) = Nothing
        Set mvarHolder(i) = Nothing
    Next i
    For i = 1 To 2
        Set mvarBaselines(i) = Nothing
    Next i
End Sub

Public Sub SetBaseline(ByVal i As Integer, ByVal measurement As Cartesian3D)
    Set mvarBaselines(i) = measurement
End Sub

Public Sub SetHolder(ByVal i As Integer, ByVal vData As Cartesian3D)
    Set mvarHolder(i) = vData
End Sub

Public Sub SetSample(ByVal i As Integer, ByVal measurement As Cartesian3D)
    Set mvarSample(i) = measurement
End Sub

