VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Angular3D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Public dec As Double
Public inc As Double 'local copy
Public mag As Double 'local copy

Const Pi As Double = 3.141592653589
Const rad As Double = (Pi / 180#)
Const deg As Double = (180# / Pi)

Public Property Get X() As Double
    Dim p As Double
    p = Cos(g.inc * rad)
    X = mag * p * Cos(dec * rad)
End Property

Public Property Get Y() As Double
    Dim p As Double
    p = Cos(g.inc * rad)
    Y = mag * p * Sin(dec * rad)

End Property

Public Property Get Z() As Double
    Z = mag * Sin(inc * rad)
End Property

