VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "InterpolationRange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvarStartRow As Long 'local copy
Private mvarEndRow As Long 'local copy

Public Property Let EndRow(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.EndRow = 5
    mvarEndRow = vData
End Property

Public Property Get EndRow() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.EndRow
    EndRow = mvarEndRow
End Property

Public Property Let StartRow(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.StartRow = 5
    mvarStartRow = vData
End Property

Public Property Get StartRow() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.StartRow
    StartRow = mvarStartRow
End Property

