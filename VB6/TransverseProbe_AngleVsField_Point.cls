VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AngleVsField_Point"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvarangle As Double 'local copy
Private mvarpeak_field As Double 'local copy

Public Property Let angle(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.angle = 5
    mvarangle = vData
End Property

Public Property Get angle() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.angle
    angle = mvarangle
End Property

Public Property Let peak_field(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.peak_field = 5
    mvarpeak_field = vData
End Property

Public Property Get peak_field() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.peak_field
    peak_field = mvarpeak_field
End Property

