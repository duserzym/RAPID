VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Channel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Public key As Long
Private mvarChanDescs As ChannelDescs
Private mvarChanNum As Long
Private mvarChanName As String
Private mvarChanType As String
Private mvarBoardININum As Long
Private mvarBoardName As String

Public Property Let BoardININum(ByVal lData As Long)

    mvarBoardININum = lData

End Property

Public Property Get BoardININum() As Long

    BoardININum = mvarBoardININum
    
End Property

Public Property Let BoardName(ByVal sData As String)

    mvarBoardName = sData

End Property

Public Property Get BoardName() As String

    BoardName = mvarBoardName
    
End Property

Public Property Set ChanDescs(ByVal ChanDescs As ChannelDescs)
    
    Set mvarChanDescs = ChanDescs

End Property

Public Property Get ChanDescs() As ChannelDescs

    Set ChanDescs = mvarChanDescs

End Property

Public Property Let ChanName(ByVal vData As String)

    mvarChanName = vData

End Property

Public Property Get ChanName() As String

    ChanName = mvarChanName
    
End Property

Public Property Let ChanNum(ByVal vData As Long)

    mvarChanNum = vData
    
End Property

Public Property Get ChanNum() As Long

    ChanNum = mvarChanNum

End Property

Public Property Let ChanType(ByVal sData As String)

    mvarChanType = sData

End Property

Public Property Get ChanType() As String

    ChanType = mvarChanType
    
End Property

