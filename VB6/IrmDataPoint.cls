VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IrmDataPoint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvartime As Long 'local copy
Private mvarread_voltage As Double 'local copy
Private mvardelta_rate As Double

Public Property Let delta_rate(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.read_voltage = 5
    mvardelta_rate = vData
End Property

Public Property Get delta_rate() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.read_voltage
    delta_rate = mvardelta_rate
End Property

Public Property Let read_voltage(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.read_voltage = 5
    mvarread_voltage = vData
End Property

Public Property Get read_voltage() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.read_voltage
    read_voltage = mvarread_voltage
End Property

Public Property Let time_stamp(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.time = 5
    mvartime = vData
End Property

Public Property Get time_stamp() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.time
    time_stamp = mvartime
End Property

