VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AdwinAfParameter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Enum AdwinParameterTypeEnum
    LongInteger = 0
    FloatingPoint = 1
End Enum

Private mvarParamType As AdwinParameterTypeEnum 'local copy
Private mvarParamName As String 'local copy
Private mvarParamNumber As Byte 'local copy
Private mvarParamSingle As Single 'local copy
Private mvarParamLong As Long 'local copy
Public Property Let ParamLong(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ParamLong = 5
    mvarParamLong = vData
End Property


Public Property Get ParamLong() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ParamLong
    ParamLong = mvarParamLong
End Property



Public Property Let ParamSingle(ByVal vData As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ParamSingle = 5
    mvarParamSingle = vData
End Property


Public Property Get ParamSingle() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ParamSingle
    ParamSingle = mvarParamSingle
End Property



Public Property Let ParamNumber(ByVal vData As Byte)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ParamNumber = 5
    mvarParamNumber = vData
End Property


Public Property Get ParamNumber() As Byte
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ParamNumber
    ParamNumber = mvarParamNumber
End Property



Public Property Let ParamName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ParamName = 5
    mvarParamName = vData
End Property


Public Property Get ParamName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ParamName
    ParamName = mvarParamName
End Property



Public Property Let ParamType(ByVal vData As AdwinParameterTypeEnum)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ParamType = 5
    mvarParamType = vData
End Property


Public Property Get ParamType() As AdwinParameterTypeEnum
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ParamType
    ParamType = mvarParamType
End Property

Public Function ToString() As String

    If mvarParamType = FloatingPoint Then
    
        Dim ret_val As String
    
        ret_val = Format$(mvarParamSingle, "#0.0###########")

        'Clip off dangling .0's
        
        If Right(ret_val, 2) = ".0" Then
        
            ret_val = Mid(ret_val, 1, Len(ret_val) - 2)
            
        End If
        
        ToString = ret_val
            
    Else
    
        ToString = Trim(CStr(mvarParamLong))
        
    End If

End Function

Public Function GetValue() As Variant
    
    If mvarParamType = FloatingPoint Then
    
        GetValue = mvarParamSingle
        
    Else
    
        GetValue = mvarParamLong
        
    End If
    
End Function

Public Function TrySetValue(ByVal value As Variant) As Boolean

    TrySetValue = True

    If mvarParamType = FloatingPoint Then
        
        On Error Resume Next
        mvarParamSingle = CSng(value)
        
        If Err.number <> 0 Then
            mvarParamSingle = 0
            TrySetValue = False
        End If
        
        On Error GoTo 0
        
    Else
    
        On Error Resume Next
        mvarParamLong = CLng(value)
        
        If Err.number <> 0 Then
            mvarParamLong = 0
            TrySetValue = False
        End If
        
        On Error GoTo 0
        
    End If
            
End Function

Public Function GetTypeString() As String

    If mvarParamType = FloatingPoint Then
    
        GetTypeString = "FPAR"
        
    ElseIf mvarParamType = LongInteger Then
    
        GetTypeString = "PAR"
        
    Else
    
        GetTypeString = "unknown"
        
    End If

End Function
