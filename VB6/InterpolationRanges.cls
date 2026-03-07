VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "InterpolationRanges"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"InterpolationRange"
Attribute VB_Ext_KEY = "Member0" ,"InterpolationRange"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable to hold collection
Private mCol As Collection

Public Property Get Count() As Long
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    Count = mCol.Count
End Property

Public Property Get Item(vntIndexKey As Variant) As InterpolationRange
Attribute Item.VB_UserMemId = 0

    If IsNumeric(vntIndexKey) Then
        Dim Index As Long: Index = CLng(vntIndexKey) + 1
        
        Set Item = mCol(Index)
    Else
        Set Item = mCol(vntIndexKey)
    End If

  
End Property

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Set NewEnum = mCol.[_NewEnum]
End Property

Public Function Add(StartRow As Long, EndRow As Long, Optional sKey As String) As InterpolationRange
    'create a new object
    Dim objNewMember As InterpolationRange
    Set objNewMember = New InterpolationRange


    'set the properties passed into the method
    objNewMember.StartRow = StartRow
    objNewMember.EndRow = EndRow
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
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

    Dim range As InterpolationRange

    For Each range In mCol
    
        Set range = Nothing
            
    Next

    Set mCol = Nothing
    Set mCol = New Collection

End Sub

Public Function GetIndexByRow(ByVal row_num As Long) As Integer

    GetIndexByRow = -1

    If mCol Is Nothing Then Exit Function
    If mCol.Count = 0 Then Exit Function
    
    Dim i As Integer
    
    For i = 1 To mCol.Count
    
        If row_num = mCol(i).StartRow Or _
           row_num = mCol(i).EndRow Then
            GetIndexByRow = i - 1
            Exit Function
        End If
    Next i

End Function

Public Sub Remove(vntIndexKey As Variant)
    'used when removing an element from the collection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)
    If IsNumeric(vntIndexKey) Then
    
        Dim Index As Long: Index = CLng(vntIndexKey) + 1
        mCol.Remove Index
    
    Else
    
        mCol.Remove vntIndexKey
    
    End If

    
End Sub

