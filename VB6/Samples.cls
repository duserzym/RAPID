VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Samples"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"Sample"
Attribute VB_Ext_KEY = "Member0" ,"Sample"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' collection of samples
' used for storing samples contained in IndexFile (.SAM file)
'local variable to hold collection
Private mCol As Collection
Private mvarIndexFile As String

Public Property Get Count() As Long
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    Count = mCol.Count
End Property

Public Property Let IndexFile(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sampleName = 5
    mvarIndexFile = vData
End Property

Public Property Get IndexFile() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sampleName
    IndexFile = mvarIndexFile
End Property

Public Property Get Item(vntIndexKey As Variant) As Sample
Attribute Item.VB_UserMemId = 0
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
    On Error GoTo oops
    If LenB(vntIndexKey) > 0 Then Set Item = mCol(vntIndexKey)
    On Error GoTo 0
oops:
End Property

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Set NewEnum = mCol.[_NewEnum]
End Property

Public Function Add(Samplename As String, Optional sampleHole As Long = 0) As Sample
    'create a new object
    Dim objNewMember As Sample
    Set objNewMember = New Sample
    ' watch out for double adds
    On Error GoTo oops
    'set the properties passed into the method
    objNewMember.Samplename = Samplename
    objNewMember.sampleHole = sampleHole
    objNewMember.IndexFile = mvarIndexFile
    mCol.Add objNewMember, Samplename
    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing
    On Error GoTo 0
    Exit Function
oops:
    Select Case Err.number
        Case 457        ' key already used
            Set Add = objNewMember
            Item(Samplename).sampleHole = sampleHole
            'Set Item(sampleName) = objNewMember
        Case Else
            MsgBox "unknown error in samples"
    End Select
    Set objNewMember = Nothing
End Function

Private Sub Class_Initialize()
    'creates the collection when this class is created
    Set mCol = New Collection
End Sub

Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Set mCol = Nothing
End Sub

Public Sub Clear()
    Set mCol = Nothing
    Set mCol = New Collection
End Sub

Public Function IsValidSample(Samplename As String)
    ' is the sample Samplname in the collection?
    IsValidSample = False
    On Error GoTo fin
    IsValidSample = ((LenB(Item(Samplename).Samplename)) > 0)
fin:
    On Error GoTo 0
End Function

Public Sub Remove(vntIndexKey As Variant)
    'used when removing an element from the collection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)
    mCol.Remove vntIndexKey
End Sub

