VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AdwinAfOutputParameters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarMax_Ramp_Voltage_Used As AdwinAfParameter 'local copy
Private mvarMeasured_Peak_Monitor_Voltage As AdwinAfParameter 'local copy
Private mvarTime_Step_Between_Points As AdwinAfParameter 'local copy
Private mvarNumber_Points_Per_Period As AdwinAfParameter 'local copy
Private mvarActual_Slope_Down_Used As AdwinAfParameter 'local copy
Private mvarTotal_Output_Points As AdwinAfParameter 'local copy
Private mvarTotal_Monitor_Points As AdwinAfParameter 'local copy
Private mvarRamp_Up_Last_Point As AdwinAfParameter 'local copy
Private mvarRamp_Down_First_Point As AdwinAfParameter 'local copy
Private mvarRamp_Coil As String

Public Property Get Coil() As String

    Coil = mvarRamp_Coil

End Property


Public Property Let Coil(ByVal value As String)

    mvarRamp_Coil = value

End Property

Public Property Set Ramp_Down_First_Point(ByVal vData As AdwinAfParameter)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Ramp_Down_First_Point = Form1
    Set mvarRamp_Down_First_Point = vData
End Property


Public Property Get Ramp_Down_First_Point() As AdwinAfParameter
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Ramp_Down_First_Point
    Set Ramp_Down_First_Point = mvarRamp_Down_First_Point
End Property



Public Property Set Ramp_Up_Last_Point(ByVal vData As AdwinAfParameter)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Ramp_Up_Last_Point = Form1
    Set mvarRamp_Up_Last_Point = vData
End Property


Public Property Get Ramp_Up_Last_Point() As AdwinAfParameter
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Ramp_Up_Last_Point
    Set Ramp_Up_Last_Point = mvarRamp_Up_Last_Point
End Property



Public Property Set Total_Monitor_Points(ByVal vData As AdwinAfParameter)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Total_Monitor_Points = Form1
    Set mvarTotal_Monitor_Points = vData
End Property


Public Property Get Total_Monitor_Points() As AdwinAfParameter
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Total_Monitor_Points
    Set Total_Monitor_Points = mvarTotal_Monitor_Points
End Property



Public Property Set Total_Output_Points(ByVal vData As AdwinAfParameter)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Total_Output_Points = Form1
    Set mvarTotal_Output_Points = vData
End Property


Public Property Get Total_Output_Points() As AdwinAfParameter
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Total_Output_Points
    Set Total_Output_Points = mvarTotal_Output_Points
End Property



Public Property Set Actual_Slope_Down_Used(ByVal vData As AdwinAfParameter)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Actual_Slope_Down_Used = Form1
    Set mvarActual_Slope_Down_Used = vData
End Property


Public Property Get Actual_Slope_Down_Used() As AdwinAfParameter
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Actual_Slope_Down_Used
    Set Actual_Slope_Down_Used = mvarActual_Slope_Down_Used
End Property



Public Property Set Number_Points_Per_Period(ByVal vData As AdwinAfParameter)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Number_Points_Per_Period = Form1
    Set mvarNumber_Points_Per_Period = vData
End Property


Public Property Get Number_Points_Per_Period() As AdwinAfParameter
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Number_Points_Per_Period
    Set Number_Points_Per_Period = mvarNumber_Points_Per_Period
End Property



Public Property Set Time_Step_Between_Points(ByVal vData As AdwinAfParameter)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Time_Step_Between_Points = Form1
    Set mvarTime_Step_Between_Points = vData
End Property


Public Property Get Time_Step_Between_Points() As AdwinAfParameter
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Time_Step_Between_Points
    Set Time_Step_Between_Points = mvarTime_Step_Between_Points
End Property



Public Property Set Measured_Peak_Monitor_Voltage(ByVal vData As AdwinAfParameter)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Measured_Peak_Monitor_Voltage = Form1
    Set mvarMeasured_Peak_Monitor_Voltage = vData
End Property


Public Property Get Measured_Peak_Monitor_Voltage() As AdwinAfParameter
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Measured_Peak_Monitor_Voltage
    Set Measured_Peak_Monitor_Voltage = mvarMeasured_Peak_Monitor_Voltage
End Property



Public Property Set Max_Ramp_Voltage_Used(ByVal vData As AdwinAfParameter)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Max_Ramp_Voltage_Used = Form1
    Set mvarMax_Ramp_Voltage_Used = vData
End Property


Public Property Get Max_Ramp_Voltage_Used() As AdwinAfParameter
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Max_Ramp_Voltage_Used
    Set Max_Ramp_Voltage_Used = mvarMax_Ramp_Voltage_Used
End Property

Public Function CreateNewFpar(ByVal name As String, _
                              ByVal number As Byte, _
                              ByVal value As Single) As AdwinAfParameter
                              
    Dim new_fpar As AdwinAfParameter
    Set new_fpar = New AdwinAfParameter
    
    new_fpar.ParamName = name
    new_fpar.ParamNumber = number
    new_fpar.ParamType = FloatingPoint
    new_fpar.ParamSingle = value
    
    Set CreateNewFpar = new_fpar
                              
End Function

Public Function CreateNewLpar(ByVal name As String, _
                              ByVal number As Byte, _
                              ByVal value As Long) As AdwinAfParameter
                              
    Dim new_lpar As AdwinAfParameter
    Set new_lpar = New AdwinAfParameter
    
    new_lpar.ParamName = name
    new_lpar.ParamNumber = number
    new_lpar.ParamType = LongInteger
    new_lpar.ParamLong = value
    
    Set CreateNewLpar = new_lpar
                              
End Function

Public Function GetRampDownDuration() As Single

    GetRampDownDuration = 0

    If mvarTime_Step_Between_Points Is Nothing Or _
       mvarRamp_Down_First_Point Is Nothing Or _
       mvarTotal_Monitor_Points Is Nothing Then
       
        Exit Function
        
    End If
    
    If mvarTime_Step_Between_Points.ParamSingle <= 0 Then
    
        Exit Function
        
    End If
    
    Dim num_points As Long
    
    num_points = mvarTotal_Monitor_Points.ParamLong - mvarRamp_Down_First_Point.ParamLong
    
    If num_points <= 0 Then num_points = 0
    
    GetRampDownDuration = CSng(num_points * mvarTime_Step_Between_Points.ParamSingle)
    
End Function

Public Function GetRampUpDuration() As Single

    GetRampUpDuration = 0

    If mvarTime_Step_Between_Points Is Nothing Or _
       mvarRamp_Up_Last_Point Is Nothing Then
       
        Exit Function
        
    End If
    
    If mvarTime_Step_Between_Points.ParamSingle <= 0 Then
    
        Exit Function
        
    End If
    
    Dim num_points As Long
    
    num_points = mvarRamp_Up_Last_Point.ParamLong
    
    If num_points <= 0 Then num_points = 0
    
    GetRampUpDuration = CSng(num_points * mvarTime_Step_Between_Points.ParamSingle)
    
End Function

Public Function GetTotalRampDuration() As Single

    GetTotalRampDuration = 0

    If mvarTime_Step_Between_Points Is Nothing Or _
       mvarTotal_Monitor_Points Is Nothing Then
       
        Exit Function
        
    End If
    
    If mvarTime_Step_Between_Points.ParamSingle <= 0 Then
    
        Exit Function
        
    End If
    
    Dim num_points As Long
    
    num_points = mvarTotal_Monitor_Points.ParamLong
    
    If num_points <= 0 Then num_points = 0
    
    GetTotalRampDuration = CSng(num_points * mvarTime_Step_Between_Points.ParamSingle)
    
End Function

Public Function SingleToString(ByVal value As Single) As String

    Dim ret_val As String
    
    ret_val = Format$(value, "#0.0##########")
    
    If Strings.Right(ret_val, 2) = ".0" Then
    
        ret_val = Mid(ret_val, 1, Len(ret_val) - 2)

    End If
    
    SingleToString = ret_val

End Function

Public Function GetPeakDuration() As Single

    GetPeakDuration = 0

    If mvarTime_Step_Between_Points Is Nothing Or _
       mvarRamp_Down_First_Point Is Nothing Or _
       mvarRamp_Up_Last_Point Is Nothing Then
       
        Exit Function
        
    End If
    
    If mvarTime_Step_Between_Points.ParamSingle <= 0 Then
    
        Exit Function
        
    End If
    
    Dim num_points As Long
    
    num_points = mvarRamp_Down_First_Point.ParamLong - mvarRamp_Up_Last_Point.ParamLong
    
    If num_points <= 0 Then num_points = 0
    
    GetPeakDuration = CSng(num_points * mvarTime_Step_Between_Points.ParamSingle)
    
End Function

Public Sub Class_Initialize()
    
    Set mvarMax_Ramp_Voltage_Used = CreateNewFpar("OutputMax", 5, 0)
    Set mvarMeasured_Peak_Monitor_Voltage = CreateNewFpar("MonitorMax", 4, 0)
    Set mvarTime_Step_Between_Points = CreateNewFpar("ACOUT_TIMESTEP", 6, 0)
    Set mvarNumber_Points_Per_Period = CreateNewFpar("NPointsPerPeriod", 7, 0)
    Set mvarActual_Slope_Down_Used = CreateNewFpar("ActualSlopeDown", 32, 0)
    
    Set mvarTotal_Output_Points = CreateNewLpar("OUTCOUNT", 5, 0)
    Set mvarTotal_Monitor_Points = CreateNewLpar("MonitorCount", 6, 0)
    Set mvarRamp_Up_Last_Point = CreateNewLpar("RAMPUPLASTPT", 7, 0)
    Set mvarRamp_Down_First_Point = CreateNewLpar("RAMPDOWNFIRSTPT", 8, 0)

End Sub

