VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MeasurementBlocks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"MeasurementBlock"
Attribute VB_Ext_KEY = "Member0" ,"MeasurementBlock"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable to hold collection
Private mCol As Collection

Public Property Get AverageBlock() As MeasurementBlock
    Dim i As Integer, j As Integer
    Dim cnt As Integer
    Dim sumMeasurements(1 To 4) As Cartesian3D
    Dim workingvector As Cartesian3D
    Set AverageBlock = New MeasurementBlock
    
    For j = 1 To 4
        Set sumMeasurements(j) = New Cartesian3D
    Next j
    
    cnt = Count
    If cnt = 0 Then Exit Property
    For i = 1 To cnt
            For j = 1 To 4
                Set workingvector = Item(i).BaselineAdjustedSample(j)
                With sumMeasurements(j)
                    .X = .X + workingvector.X
                    .Y = .Y + workingvector.Y
                    .Z = .Z + workingvector.Z
                End With
                Set workingvector = Nothing
            Next j
    Next i
    
    
    For i = 1 To 4
        With sumMeasurements(i)
            .X = .X / cnt
            .Y = .Y / cnt
            .Z = .Z / cnt
        End With
        AverageBlock.SetSample i, sumMeasurements(i)
    Next i
    
    For j = 1 To 4
        Set sumMeasurements(j) = Nothing
    Next j
End Property

Public Property Get Count() As Long
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    Count = mCol.Count
End Property

Public Property Get DirectionalSubset(isUp As Boolean) As MeasurementBlocks
    Dim i As Integer
    Dim cnt As Integer

    Set DirectionalSubset = New MeasurementBlocks
    cnt = Count
    If cnt = 0 Then Exit Property
    
    For i = 1 To cnt
        If Item(i).isUp = isUp Then
            DirectionalSubset.Add Item(i)
        End If
    Next i
End Property

Public Property Get ErrorHorizontal() As Double
    Dim dataUp As MeasurementBlocks, datadown As MeasurementBlocks
    Set dataUp = DirectionalSubset(True)
    Set datadown = DirectionalSubset(False)

    If datadown.Count = 0 Or dataUp.Count = 0 Then
        ErrorHorizontal = 0
    Else
        ErrorHorizontal = RadToDeg((Atan2(dataUp.VectAvg.X, dataUp.VectAvg.Y) - _
            Atan2(datadown.VectAvg.X, datadown.VectAvg.Y)))
    End If
End Property

Public Property Get FischerSD() As Double
    ' returns theta-63
    ' this is circular SD from Fischer distribution
    
    FischerSD = 81 / Sqr(Kappa)
End Property

Public Property Get HolderFVal() As Double
    Dim cnt As Integer, i As Integer
    cnt = Count
    If cnt = 0 Then Exit Property
    
    For i = 1 To cnt
        HolderFVal = HolderFVal + Item(i).HolderFVal / cnt
    Next i
End Property

Public Property Get Item(vntIndexKey As Variant) As MeasurementBlock
Attribute Item.VB_UserMemId = 0
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol(vntIndexKey)
End Property

Public Property Get Kappa() As Double
    Dim N As Integer
    Dim r As Double
    
    N = Count * 4
    r = ResultantVector.mag
    If N = r Then
        Kappa = 0.000000001
    Else
        Kappa = (N - 1) / (N - r)
    End If
End Property

Public Property Get Last() As MeasurementBlock
    Set Last = Item(Count)
End Property

Public Property Get Moment() As Double
    Moment = VectAvg.mag * RangeFact
End Property

Public Property Get MomentVector() As Cartesian3D
    Dim workingvector As Cartesian3D
    Set workingvector = VectAvg
    Set MomentVector = VectAvg
    With MomentVector
        .X = workingvector.X * RangeFact
        .Y = workingvector.Y * RangeFact
        .Z = workingvector.Z * RangeFact
    End With
    Set workingvector = Nothing
End Property

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Set NewEnum = mCol.[_NewEnum]
End Property

Public Property Get NumMeasurements() As Integer
    NumMeasurements = Count * 4
End Property

Public Property Get ResultantVector() As Cartesian3D
    Dim i As Integer
    Dim cnt As Integer
    Dim workingvector As Cartesian3D
    
    Set ResultantVector = New Cartesian3D
    
    cnt = Count
    If cnt = 0 Then Exit Property
    
    For i = 1 To cnt
        Set workingvector = Item(i).SumUnitVectors
        With ResultantVector
            .X = .X + workingvector.X
            .Y = .Y + workingvector.Y
            .Z = .Z + workingvector.Z
        End With
        Set workingvector = Nothing
    Next i
    
End Property

Public Property Get SigDrift() As Double
    Dim DriftMag As Double
    Dim cnt As Integer, i As Integer
    
    cnt = Count
    If cnt > 0 Then
        For i = 1 To cnt
            DriftMag = DriftMag + Item(i).driftc.mag / cnt
        Next i
    End If
    
    If DriftMag = 0 Then DriftMag = 0.000000001
    SigDrift = VectAvg.mag / DriftMag
End Property

Public Property Get SigHolder() As Double
    Dim HolderMag As Double
    
    HolderMag = HolderFVal
    If HolderMag = 0 Then HolderMag = 0.000000001
    SigHolder = VectAvg.mag / HolderMag
End Property

Public Property Get SigInduced() As Double
    Dim induced As Double
    
    induced = VectInd.mag
    If induced = 0 Then induced = 0.000000001
    SigInduced = VectAvg.mag / induced
End Property

Public Property Get SigNoise() As Double
    Dim noise As Double
    
    noise = VectSD.mag
    If noise = 0 Then noise = 0.000000001
    SigNoise = VectAvg.mag / noise
End Property

Public Property Get UpToDown() As Double
    Dim dataUp As MeasurementBlocks, datadown As MeasurementBlocks
    Set dataUp = DirectionalSubset(True)
    Set datadown = DirectionalSubset(False)

    Dim datadownmag As Double
    Dim dataupmag As Double
    
    datadownmag = datadown.VectAvg.mag
    dataupmag = dataUp.VectAvg.mag
    If datadownmag = 0 Then
        UpToDown = 0
        Exit Property
    End If
    UpToDown = dataUp.VectAvg.mag / datadownmag
End Property

Public Property Get VectAvg() As Cartesian3D
    Dim i As Integer
    Dim cnt As Integer
    Dim workingvector As Cartesian3D
    
    Set workingvector = VectSum
    Set VectAvg = New Cartesian3D
    
    cnt = Count * 4
    If cnt = 0 Then Exit Property
    With VectAvg
        .X = workingvector.X / cnt
        .Y = workingvector.Y / cnt
        .Z = workingvector.Z / cnt
    End With
    Set workingvector = Nothing
End Property

Public Property Get VectInd() As Cartesian3D
    Dim i As Integer
    Dim cnt As Integer
    Dim workingvector As Cartesian3D
    
    Set workingvector = VectSumInd
    Set VectInd = New Cartesian3D
    
    cnt = Count
    If cnt = 0 Then Exit Property
    With VectInd
        .X = workingvector.X / cnt
        .Y = workingvector.Y / cnt
        .Z = workingvector.Z / cnt
    End With
End Property

Public Property Get VectSD() As Cartesian3D
    ' the standard deviation around each X,Y, & Z components
    
    Dim cnt As Integer
    Dim sumsq As Cartesian3D, Sum As Cartesian3D
    Set VectSD = New Cartesian3D
    
    cnt = Count * 4
    If cnt = 0 Then Exit Property
    
    Set sumsq = VectSumSqs
    Set Sum = VectSum
    
    With VectSD
        .X = Sqr(Abs((sumsq.X * cnt - Sum.X ^ 2) _
                    / (cnt * (cnt - 1))))
        .Y = Sqr(Abs((sumsq.Y * cnt - Sum.Y ^ 2) _
                    / (cnt * (cnt - 1))))
        .Z = Sqr(Abs((sumsq.Z * cnt - Sum.Z ^ 2) _
                    / (cnt * (cnt - 1))))
    End With
    
    Set sumsq = Nothing
    Set Sum = Nothing
End Property

Public Property Get VectSum() As Cartesian3D
    Dim i As Integer
    Dim cnt As Integer
    Set VectSum = New Cartesian3D
    
    cnt = Count
    If cnt = 0 Then Exit Property
    With VectSum
        For i = 1 To cnt
            .X = .X + Item(i).Sum.X
            .Y = .Y + Item(i).Sum.Y
            .Z = .Z + Item(i).Sum.Z
        Next i
    End With
End Property

Public Property Get VectSumInd() As Cartesian3D
    Dim i As Integer
    Dim cnt As Integer
    Set VectSumInd = New Cartesian3D
    
    cnt = Count
    If cnt = 0 Then Exit Property
    With VectSumInd
        For i = 1 To cnt
            .X = .X + Item(i).induced.X
            .Y = .Y + Item(i).induced.Y
            .Z = .Z + Item(i).induced.Z
        Next i
    End With
End Property

Public Property Get VectSumSqs() As Cartesian3D
    Dim i As Integer
    Dim cnt As Integer
    Set VectSumSqs = New Cartesian3D
    
    cnt = Count
    If cnt = 0 Then Exit Property
    With VectSumSqs
        For i = 1 To cnt
            .X = .X + Item(i).SumSqs.X
            .Y = .Y + Item(i).SumSqs.Y
            .Z = .Z + Item(i).SumSqs.Z
        Next i
    End With
End Property

Public Function Add(ByVal block As MeasurementBlock, Optional sKey As String) As MeasurementBlock
    'create a new object
    
    Dim i As Integer
    Dim objNewMember As MeasurementBlock
    Set objNewMember = New MeasurementBlock
    
    With block
        For i = 1 To 4
            objNewMember.SetSample i, .Sample(i)
            objNewMember.SetHolder i, .Holder(i)
        Next i
        For i = 1 To 2
            objNewMember.SetBaseline i, .Baselines(i)
        Next i
        objNewMember.isUp = .isUp
    End With

    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If

    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing

End Function

Public Sub Assimilate(blocks As MeasurementBlocks)
    Dim workingBlock As MeasurementBlock
    For Each workingBlock In blocks
        Add workingBlock
    Next workingBlock
End Sub

Private Sub Class_Initialize()
    'creates the collection when this class is created
    Set mCol = New Collection
End Sub

Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Set mCol = Nothing
End Sub

Public Sub Remove(vntIndexKey As Variant)
    'used when removing an element from the collection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)


    mCol.Remove vntIndexKey
End Sub

