VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Trigger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvarTrigType As Long 'local copy
Private mvarLowThreshold As Double
Private mvarHighThreshold As Double
Private mvarIsTriggered As Boolean

Public Property Let IsTriggered(ByVal vData As Boolean)

    mvarIsTriggered = vData
    
End Property

Public Property Get IsTriggered() As Boolean

    IsTriggered = mvarIsTriggered

End Property

Public Property Let TrigType(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TrigType = 5
    mvarTrigType = vData
End Property


Public Property Get TrigType() As Long
Attribute TrigType.VB_UserMemId = 0
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TrigType
    TrigType = mvarTrigType
End Property


Public Property Let LowThreshold(ByVal vData As Double)

    mvarLowThreshold = vData
    
End Property


Public Property Get LowThreshold() As Double

    LowThreshold = mvarLowThreshold

End Property


Public Property Let HighThreshold(ByVal vData As Double)

    mvarHighThreshold = vData
    
End Property


Public Property Get HighThreshold() As Double

    HighThreshold = mvarHighThreshold

End Property
