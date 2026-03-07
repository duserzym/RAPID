VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Boards"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"Board"
Attribute VB_Ext_KEY = "Member0" ,"Board"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable to hold collection
Private mCol As Collection

Public Property Get Count() As Long
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    Count = mCol.Count
    
End Property

Public Property Get Item(vntIndexKey As Variant) As Board
Attribute Item.VB_UserMemId = 0
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
    
    On Error GoTo BoardGetError:
    
        Set Item = mCol(vntIndexKey)
                
    On Error Resume Next
    
    Exit Property
            
BoardGetError:

    Set Item = Nothing
            
End Property

Public Property Get NewIndex() As Long
Attribute NewIndex.VB_UserMemId = -4
Attribute NewIndex.VB_MemberFlags = "40"
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    
    If mCol.Count = 0 Then
    
        Err.Raise 616, "Boards.NewIndex()", "There are no entries loaded into this Boards collection!"
        
        End
        
    End If
        
    NewIndex = mCol.Count
        
End Property

Public Function Add(Optional NewBoard As Board, Optional ByVal sKey As String) As Board
    
    If NewBoard Is Nothing Then
    
        'Need to setup a new board
        Set NewBoard = New Board
    
    End If
        
    If Len(sKey) = 0 Then
        mCol.Add NewBoard
    Else
        mCol.Add NewBoard, sKey
    End If
    
    'return the object created
    Set Add = NewBoard

End Function

Private Sub Class_Initialize()
    'creates the collection when this class is created
    Set mCol = New Collection
End Sub

Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Set mCol = Nothing
End Sub

Public Function GetMax_BoardININum() As Long

    Dim i As Long
    Dim N As Long
    Dim TempBoard As Board
    Dim TempL As Long
    Dim LastININum As Long
    Dim CurININum As Long
    
    N = mCol.Count
    
    Set TempBoard = Nothing
    
    'Set Last INI Num to a default of -1000
    LastININum = -1000
    
    If mCol.Count > 0 Then
    
        For i = 1 To N
        
            Set TempBoard = mCol.Item(i)
            
            CurININum = TempBoard.BoardININum
            
            'Find the max Board INI number, or
            'the first unused Board INI Number
            If TempL < CurININum Then
            
                TempL = CurININum
                
            End If
            
            'Check for gaps in the numbering
            If LastININum <> -1000 Then
            
                'This is the first iteration of the loop
                'set LastININum = CurININum
                LastININum = CurININum
                
            Else
            
                'This is the second or higher iteration of the loop
                If LastININum - CurININum >= 2 Then
                
                    'There's a gap in the numbering
                    'Set TempL = LastININum
                    TempL = LastININum
                    
                    'Set i > N to end the loop
                    i = N + 1
                    
                End If
                
                LastININum = CurININum
                
            End If
                
        Next i
        
    Else
    
        TempL = 0
        
    End If
    
    GetMax_BoardININum = TempL

End Function

Public Sub Remove(vntIndexKey As Variant)
    'used when removing an element from the collection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)

    mCol.Remove vntIndexKey
    
End Sub

