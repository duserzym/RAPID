VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "XYCup_Positions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"XYCup"
Attribute VB_Ext_KEY = "Member0" ,"XYCup"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable to hold collection
Private mCol As Collection

Public Property Get Count() As Long
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    Count = mCol.Count
End Property

Public Property Get Item(vntIndexKey As Variant) As XYCup
Attribute Item.VB_UserMemId = 0
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol(vntIndexKey)
End Property

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Set NewEnum = mCol.[_NewEnum]
End Property

Public Function Add(x_pos As Long, y_pos As Long, Cup_Number As Long, Optional sKey As String) As XYCup
    'create a new object
    Dim objNewMember As XYCup
    Set objNewMember = New XYCup


    'set the properties passed into the method
    objNewMember.x_pos = x_pos
    objNewMember.y_pos = y_pos
    objNewMember.Cup_Number = Cup_Number
    If Len(sKey) = 0 Then
        mCol.Add objNewMember, Trim(Str(Cup_Number))
    Else
        mCol.Add objNewMember, sKey
    End If


    'return the object created
    Set Add = objNewMember
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

    Dim cup As XYCup
    
    For Each cup In mCol
    
        Set cup = Nothing
    
    Next
    
    Set mCol = Nothing
    Set mCol = New Collection

End Sub

Public Function Contains(str_key As String) As Boolean

    

   On Error GoTo Contains_Error

        If Not mCol(str_key) Is Nothing Then
        
            Contains = True
            
        Else
        
            Contains = False
            
        End If

   On Error GoTo 0
   Exit Function

Contains_Error:
    Contains = False

End Function

Public Sub Remove(vntIndexKey As Variant)
    'used when removing an element from the collection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)


    mCol.Remove vntIndexKey
End Sub

