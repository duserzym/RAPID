VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Channels"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"Channel"
Attribute VB_Ext_KEY = "Member0" ,"Channel"
'local variable to hold collection
Private mCol As Collection

Public Property Get Count() As Long
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    Count = mCol.Count
End Property

Public Property Get Item(vntIndexKey As Variant) As Channel
Attribute Item.VB_UserMemId = 0
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol.Item(vntIndexKey)
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

Public Function Add(Optional Chan As Channel, Optional ByVal sKey As String) As Channel
    'return the object created
    
    If Chan Is Nothing Then
    
        'Chan object is not allocated yet
        Set Chan = New Channel
        
    End If
    
    If Len(sKey) = 0 Then
        mCol.Add Chan
    Else
        mCol.Add Chan, sKey
    End If
    
    Set Add = Chan

End Function

Public Function AddErrorCheck(Optional Chan As Channel, Optional ByVal sKey As String) As Boolean

    If Chan Is Nothing Then
    
        'Chan object is not allocated yet
        Set Chan = New Channel
        
    End If
    
    If Len(sKey) = 0 Then
        mCol.Add Chan
    Else
    
        'Turn on error handling
        On Error Resume Next
        
            'Attempt to add the channel
            mCol.Add Chan, sKey
            
            'look for error # 457
            If Err.number = 457 Then
            
                'We have a repeat key, return False
                'Add is unsuccessful due to repeat
                AddErrorCheck = False
                
                Exit Function
                
            ElseIf Err.number <> 0 Then
            
                'It's a different error
                Err.Raise Err.number, _
                          Err.Source, _
                          Err.Description, _
                          Err.HelpFile, _
                          Err.HelpContext
                          
            End If
            
        'Turn off error handling
        On Error GoTo 0
        
    End If
    
    AddErrorCheck = True

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

