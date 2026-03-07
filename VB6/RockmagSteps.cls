VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RockmagSteps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"RockmagStep"
Attribute VB_Ext_KEY = "Member0" ,"RockmagStep"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable to hold collection
Private mCol As Collection
Private nextStepID As Integer
Private mvarCurrentStepIndex As Integer

Public Property Get Count() As Long
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    Count = mCol.Count
End Property

Public Property Get CurrentStep() As RockmagStep
    Dim blankRockmagStep As RockmagStep
    Set blankRockmagStep = New RockmagStep
    blankRockmagStep.StepType = vbNullString
    If mvarCurrentStepIndex > 0 Then
        Set CurrentStep = Item(mvarCurrentStepIndex)
    Else
        Set CurrentStep = blankRockmagStep
    End If
    Set blankRockmagStep = Nothing
End Property

Public Property Let CurrentStepIndex(mvar As Integer)
    mvarCurrentStepIndex = mvar
End Property

Public Property Get CurrentStepIndex() As Integer
    CurrentStepIndex = mvarCurrentStepIndex
End Property

Public Property Get DoesSusceptibility() As Boolean
    Dim i As Integer
    DoesSusceptibility = False
    If Count = 0 Then Exit Property
    For i = 1 To Count
        If Item(i).MeasureSusceptibility Then
            DoesSusceptibility = True
            Exit For
        End If
    Next i
End Property

Public Property Get Item(vntIndexKey As Variant) As RockmagStep
Attribute Item.VB_UserMemId = 0
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
    On Error GoTo oops
  Set Item = mCol(vntIndexKey)
  On Error GoTo 0
  Exit Sub
oops:
  Set Item = New RockmagStep
End Property

Public Property Get NeedsCoilAFDemag() As Boolean
    Dim i As Integer
    NeedsCoilAFDemag = False
    If Count = 0 Then Exit Sub
    For i = 1 To Count
        With Item(i)
            If .BiasField > 0 Or .DemagStepLabel = "IRM" Then
                NeedsCoilAFDemag = True
                Exit For
            End If
        End With
    Next i
End Property

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Set NewEnum = mCol.[_NewEnum]
End Property

Public Function Add(ByVal StepType As String, _
                    Optional ByVal Level As Double = 0, _
                    Optional ByVal BiasField As Double = 0, _
                    Optional ByVal SpinSpeed As Single = 0, _
                    Optional ByVal HoldTime As Single = 0, _
                    Optional ByVal Measure As Boolean = True, _
                    Optional ByVal MeasureSusceptibility = False, _
                    Optional ByVal Remarks As String, _
                    Optional ByVal BeforeStep As Long, _
                    Optional ByVal AfterStep As Long) As RockmagStep
                    
    'create a new object
    Dim objNewMember As RockmagStep
    Set objNewMember = New RockmagStep
    
    'set the properties passed into the method
    objNewMember.Level = Level
    objNewMember.BiasField = BiasField
    objNewMember.StepType = StepType
    objNewMember.SpinSpeed = SpinSpeed
    objNewMember.HoldTime = HoldTime
    objNewMember.Measure = Measure
    objNewMember.MeasureSusceptibility = MeasureSusceptibility
    objNewMember.Remarks = Remarks ' (November 2007 L Carporzen) Remarks column in RMG
    objNewMember.key = "S" & nextStepID
    If BeforeStep > 0 Then
        mCol.Add objNewMember, "S" & nextStepID, Before:=BeforeStep
    ElseIf AfterStep > 0 Then
        mCol.Add objNewMember, "S" & nextStepID, After:=AfterStep
    Else
        mCol.Add objNewMember, "S" & nextStepID
    End If
    nextStepID = nextStepID + 1
    
    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing
    
End Function

Public Sub AdvanceStep()
    mvarCurrentStepIndex = (mvarCurrentStepIndex + 1) Mod (Count + 1)
End Sub

Private Sub Class_Initialize()
    'creates the collection when this class is created
    Set mCol = New Collection
    nextStepID = 1
    mvarCurrentStepIndex = 1
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

