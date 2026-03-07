VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Plot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvaryStart As Long 'local copy
Private mvaryEnd As Long 'local copy
Private mvaryZero As Long 'local copy
Private mvaryTick As Long 'local copy
Private mvarxStart As Long 'local copy
Private mvarxEnd As Long 'local copy
Private mvarxTick As Long 'local copy
Private mvaryMax As Double 'local copy
Private mvaryMin As Double 'local copy
Private mvarxMax As Double 'local copy
Private mvarxMin As Double 'local copy
'local variable(s) to hold property value(s)
Private mvarControlIndex As Long 'local copy
Public Property Let NumPoints(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ControlIndex = 5
    mvarControlIndex = vData
End Property


Public Property Get NumPoints() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ControlIndex
     NumPoints = mvarControlIndex
End Property



Public Property Let xMin(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.xMin = 5
    mvarxMin = vData
End Property


Public Property Get xMin() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.xMin
     xMin = mvarxMin
End Property



Public Property Let xMax(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.xMax = 5
    mvarxMax = vData
End Property


Public Property Get xMax() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.xMax
     xMax = mvarxMax
End Property



Public Property Let yMin(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.yMin = 5
    mvaryMin = vData
End Property


Public Property Get yMin() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.yMin
    yMin = mvaryMin
End Property



Public Property Let yMax(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.yMax = 5
    mvaryMax = vData
End Property


Public Property Get yMax() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.yMax
    yMax = mvaryMax
End Property



Public Property Let xTick(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.xTick = 5
    mvarxTick = vData
End Property


Public Property Get xTick() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.xTick
    xTick = mvarxTick
End Property



Public Property Let xEnd(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.xEnd = 5
    mvarxEnd = vData
End Property


Public Property Get xEnd() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.xEnd
     xEnd = mvarxEnd
End Property



Public Property Let xStart(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.xStart = 5
    mvarxStart = vData
End Property


Public Property Get xStart() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.xStart
     xStart = mvarxStart
End Property



Public Property Let yTick(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.yTick = 5
    mvaryTick = vData
End Property


Public Property Get yTick() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.yTick
     yTick = mvaryTick
End Property



Public Property Let yZero(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.yZero = 5
    mvaryZero = vData
End Property


Public Property Get yZero() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.yZero
     yZero = mvaryZero
End Property



Public Property Let yEnd(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.yEnd = 5
    mvaryEnd = vData
End Property


Public Property Get yEnd() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.yEnd
    yEnd = mvaryEnd
End Property



Public Property Let yStart(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.yStart = 5
    mvaryStart = vData
End Property


Public Property Get yStart() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.yStart
    yStart = mvaryStart
End Property



