VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Waves"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"Wave"
Attribute VB_Ext_KEY = "Member0" ,"Wave"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable to hold collection
Private mCol As Collection

Public Property Get Count() As Long
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    Count = mCol.Count
End Property

Public Property Get Item(vntIndexKey As Variant) As Wave
Attribute Item.VB_UserMemId = 0
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol(vntIndexKey)
End Property

Public Property Get NewIndex() As Long
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    
    If mCol.Count = 0 Then
    
        Err.Raise 616, "Boards.NewIndex()", "There are no entries loaded into this Boards collection!"
        
        End
        
    End If
        
    NewIndex = mCol.Count
        
End Property

Public Function Add(Optional NewWave As Wave, Optional ByVal sKey As String) As Wave
    
    If NewWave Is Nothing Then
    
        'Need to allocate new Wave object
        Set NewWave = New Wave
        
    End If
    
    If Len(sKey) = 0 Then
        mCol.Add NewWave
    Else
        mCol.Add NewWave, sKey
    End If
    
    Set Add = NewWave
    
End Function

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

