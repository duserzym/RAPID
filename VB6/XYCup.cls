VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "XYCup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvarX_Pos As Long 'local copy
Private mvarY_Pos As Long 'local copy
Private mvarCup_Number As Long 'local copy

Public Property Let Cup_Number(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Cup_Number = 5
    mvarCup_Number = vData
End Property

Public Property Get Cup_Number() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Cup_Number
    Cup_Number = mvarCup_Number
End Property

Public Property Let x_pos(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.X_Pos = 5
    mvarX_Pos = vData
End Property

Public Property Get x_pos() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.X_Pos
    x_pos = mvarX_Pos
End Property

Public Property Let y_pos(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Y_Pos = 5
    mvarY_Pos = vData
End Property

Public Property Get y_pos() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Y_Pos
    y_pos = mvarY_Pos
End Property

