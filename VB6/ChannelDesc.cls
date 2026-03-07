VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ChannelDescs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvarChanDesc As Collection 'local copy

Public Property Get ChanDesc(ByVal vrntKeyIndex As Variant) As String

    'If Index > number of elements in the channel description
    'string array, then return "ERROR"

    ChanDesc = mvarChanDesc(vrntKeyIndex)

End Property

Public Property Get Count() As Long
    
    Count = mvarChanDesc.Count
    
End Property

Public Function AddDesc(ByVal sData As String, Optional ByVal vrntKeyIndex As Variant = "-1")

    Dim N As Long
    Dim NewKey As Variant
    
    'Check to see if the user entered a key
    If vrntKeyIndex = "-1" Then
    
        'No key entered
        'Get the number of elements in the collection
        N = mvarChanDesc.Count
        
        NewKey = Trim(Str(N + 1))
        
    Else
    
        NewKey = vrntKeyIndex
        
    End If

    mvarChanDesc.Add sData, NewKey

End Function

Private Sub Class_Initialize()
    'creates the collection when this class is created
    Set mvarChanDesc = New Collection
End Sub

Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Set mvarChanDesc = Nothing
End Sub

Public Sub Clear()

    'Deallocate the collection
    Set mvarChanDesc = Nothing
    
    'Reallocate a new collection
    Set mvarChanDesc = New Collection
    
End Sub

Public Function GetAll(Optional ByVal Delimeter As String = vbNewLine) As String

    Dim i As Long
    Dim N As Long
    Dim TempStr As Long
        
    'Initialize TempStr = empty string
    TempStr = ""
    
    N = mvarChanDesc.Count
    
    'Loop through the elements of the Chan Description string array
    'and concatenate the the elements together with a vbNewline character
    'in between each element
    For i = 1 To N
    
        TempStr = TempStr & mvarChanDesc(i) & Delimeter
        
    Next i
        
    'Now need to clip off the last vbnewline character
    TempStr = Mid(TempStr, 1, Len(TempStr) - 1)

    'Return TempStr as the concatenated string result
    GetAll = TempStr

End Function

Public Sub RemoveDesc(ByVal vrntKeyIndex As Variant)

    mvarChanDesc.Remove (vrntKeyIndex)
    
End Sub

