VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AdwinAfInputParameters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarSlope_Up As AdwinAfParameter 'local copy
Private mvarSlope_Down As AdwinAfParameter 'local copy
Private mvarResonance_Freq As AdwinAfParameter 'local copy
Private mvarPeak_Monitor_Voltage As AdwinAfParameter 'local copy
Private mvarMax_Ramp_Voltage As AdwinAfParameter 'local copy
Private mvarMax_Monitor_Voltage As AdwinAfParameter 'local copy
Private mvarPeak_Ramp_Voltage As AdwinAfParameter 'local copy
Private mvarRamp_Mode As AdwinAfParameter 'local copy
Private mvarOutput_Port_Number As AdwinAfParameter 'local copy
Private mvarMonitor_Port_Number As AdwinAfParameter 'local copy
Private mvarProcess_Delay As AdwinAfParameter 'local copy
Private mvarNoise_Level As AdwinAfParameter 'local copy
Private mvarNumber_Periods_Hang_At_Peak As AdwinAfParameter 'local copy
Private mvarNumber_Periods_Ramp_Down As AdwinAfParameter 'local copy
Private mvarRamp_Down_Mode As AdwinAfParameter 'local copy
Private mvarRamp_Start_DateTime As Date
Private mvarRamp_End_DateTime As Date
Private mvarRamp_Coil As String

Public Property Get coil() As String

    coil = mvarRamp_Coil

End Property


Public Property Let coil(ByVal value As String)

    mvarRamp_Coil = value

End Property

Public Property Get Ramp_Start_DateTime() As Date

    Ramp_Start_DateTime = mvarRamp_Start_DateTime

End Property

Public Property Let Ramp_Start_DateTime(ByVal value As Date)

    mvarRamp_Start_DateTime = value

End Property

Public Property Get Ramp_End_DateTime() As Date

    Ramp_End_DateTime = mvarRamp_End_DateTime

End Property

Public Property Let Ramp_End_DateTime(ByVal value As Date)

    mvarRamp_End_DateTime = value

End Property

Public Property Set ramp_down_mode(ByVal vData As AdwinAfParameter)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Ramp_Down_Mode = Form1
    Set mvarRamp_Down_Mode = vData
End Property


Public Property Get ramp_down_mode() As AdwinAfParameter
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Ramp_Down_Mode
    Set ramp_down_mode = mvarRamp_Down_Mode
End Property

Public Property Set Number_Periods_Ramp_Down(ByVal vData As AdwinAfParameter)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Number_Periods_Ramp_Down = Form1
    Set mvarNumber_Periods_Ramp_Down = vData
End Property


Public Property Get Number_Periods_Ramp_Down() As AdwinAfParameter
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Number_Periods_Ramp_Down
    Set Number_Periods_Ramp_Down = mvarNumber_Periods_Ramp_Down
End Property



Public Property Set Number_Periods_Hang_At_Peak(ByVal vData As AdwinAfParameter)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Number_Periods_Hang_At_Peak = Form1
    Set mvarNumber_Periods_Hang_At_Peak = vData
End Property


Public Property Get Number_Periods_Hang_At_Peak() As AdwinAfParameter
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Number_Periods_Hang_At_Peak
    Set Number_Periods_Hang_At_Peak = mvarNumber_Periods_Hang_At_Peak
End Property



Public Property Set Noise_Level(ByVal vData As AdwinAfParameter)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Noise_Level = Form1
    Set mvarNoise_Level = vData
End Property


Public Property Get Noise_Level() As AdwinAfParameter
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Noise_Level
    Set Noise_Level = mvarNoise_Level
End Property



Public Property Set Process_Delay(ByVal vData As AdwinAfParameter)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Process_Delay = Form1
    Set mvarProcess_Delay = vData
End Property


Public Property Get Process_Delay() As AdwinAfParameter
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Process_Delay
    Set Process_Delay = mvarProcess_Delay
End Property



Public Property Set Monitor_Port_Number(ByVal vData As AdwinAfParameter)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Monitor_Port_Number = Form1
    Set mvarMonitor_Port_Number = vData
End Property


Public Property Get Monitor_Port_Number() As AdwinAfParameter
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Monitor_Port_Number
    Set Monitor_Port_Number = mvarMonitor_Port_Number
End Property



Public Property Set Output_Port_Number(ByVal vData As AdwinAfParameter)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Output_Port_Number = Form1
    Set mvarOutput_Port_Number = vData
End Property


Public Property Get Output_Port_Number() As AdwinAfParameter
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Output_Port_Number
    Set Output_Port_Number = mvarOutput_Port_Number
End Property



Public Property Set ramp_mode(ByVal vData As AdwinAfParameter)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Ramp_Mode = Form1
    Set mvarRamp_Mode = vData
End Property


Public Property Get ramp_mode() As AdwinAfParameter
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Ramp_Mode
    Set ramp_mode = mvarRamp_Mode
End Property



Public Property Set Peak_Ramp_Voltage(ByVal vData As AdwinAfParameter)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Peak_Ramp_Voltage = Form1
    Set mvarPeak_Ramp_Voltage = vData
End Property


Public Property Get Peak_Ramp_Voltage() As AdwinAfParameter
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Peak_Ramp_Voltage
    Set Peak_Ramp_Voltage = mvarPeak_Ramp_Voltage
End Property



Public Property Set Max_Monitor_Voltage(ByVal vData As AdwinAfParameter)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Max_Monitor_Voltage = Form1
    Set mvarMax_Monitor_Voltage = vData
End Property


Public Property Get Max_Monitor_Voltage() As AdwinAfParameter
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Max_Monitor_Voltage
    Set Max_Monitor_Voltage = mvarMax_Monitor_Voltage
End Property



Public Property Set Max_Ramp_Voltage(ByVal vData As AdwinAfParameter)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Max_Ramp_Voltage = Form1
    Set mvarMax_Ramp_Voltage = vData
End Property


Public Property Get Max_Ramp_Voltage() As AdwinAfParameter
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Max_Ramp_Voltage
    Set Max_Ramp_Voltage = mvarMax_Ramp_Voltage
End Property



Public Property Set Peak_Monitor_Voltage(ByVal vData As AdwinAfParameter)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Peak_Monitor_Voltage = Form1
    Set mvarPeak_Monitor_Voltage = vData
End Property


Public Property Get Peak_Monitor_Voltage() As AdwinAfParameter
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Peak_Monitor_Voltage
    Set Peak_Monitor_Voltage = mvarPeak_Monitor_Voltage
End Property



Public Property Set Resonance_Freq(ByVal vData As AdwinAfParameter)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Resonance_Freq = Form1
    Set mvarResonance_Freq = vData
End Property


Public Property Get Resonance_Freq() As AdwinAfParameter
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Resonance_Freq
    Set Resonance_Freq = mvarResonance_Freq
End Property



Public Property Set Slope_Down(ByVal vData As AdwinAfParameter)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Slope_Down = Form1
    Set mvarSlope_Down = vData
End Property


Public Property Get Slope_Down() As AdwinAfParameter
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Slope_Down
    Set Slope_Down = mvarSlope_Down
End Property



Public Property Set Slope_Up(ByVal vData As AdwinAfParameter)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Slope_Up = Form1
    Set mvarSlope_Up = vData
End Property


Public Property Get Slope_Up() As AdwinAfParameter
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Slope_Up
    Set Slope_Up = mvarSlope_Up
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

Public Function GetRampDownModeStringDescrip() As String

    GetRampDownModeStringDescrip = ""

    If mvarRamp_Down_Mode Is Nothing Then Exit Function
    
    Dim ramp_down_mode As Long
    ramp_down_mode = mvarRamp_Down_Mode.ParamLong
    
    Select Case ramp_down_mode
    
        Case 1:
            GetRampDownModeStringDescrip = "Use # Periods"
        Case 2:
            GetRampDownModeStringDescrip = "Use Volt/Second slope"
        Case Else:
            GetRampDownModeStringDescrip = "Unknown: " & Trim(CStr(ramp_down_mode))
    End Select

End Function

Public Function GetRampModeStringDescrip() As String

    GetRampModeStringDescrip = ""

    If mvarRamp_Mode Is Nothing Then Exit Function
    
    Dim ramp_mode As Long
    ramp_mode = mvarRamp_Mode.ParamLong
    
    Select Case ramp_mode
    
        Case 1:
            GetRampModeStringDescrip = "Monitored Ramp"
        Case 2:
            GetRampModeStringDescrip = "Get Data, Monitored Ramp"
        Case 3:
            GetRampModeStringDescrip = "Clip Test"
        Case 4:
            GetRampModeStringDescrip = "Af Frequency Tune (Not supported)"
        Case Else:
            GetRampModeStringDescrip = "Unknown: " & Trim(CStr(ramp_mode))
    End Select

End Function

Public Function GetShortRampDescrip() As String

    If mvarRamp_Mode Is Nothing Then Exit Function
    
    Dim ramp_mode As Long
    ramp_mode = mvarRamp_Mode.ParamLong
    
    Dim descrip As String
    descrip = "AF " & mvarRamp_Coil
    
    Select Case ramp_mode
    
        Case 1:
            descrip = descrip & " Ramp"
        Case 2:
            descrip = descrip & " Ramp"
        Case 3:
            descrip = descrip & " Clip Test"
        Case 4:
            descrip = descrip & " Freq Tune"
        Case Else:
            descrip = descrip & " Unknown(" & Trim(CStr(ramp_mode)) & ")"
    End Select
    
    GetShortRampDescrip = descrip

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

Public Sub Class_Initialize()
    
    Set mvarSlope_Up = CreateNewFpar("SLOPE_UP", 31, 0)
    Set mvarSlope_Down = CreateNewFpar("SLOPE_DOWN", 32, 0)
    Set mvarPeak_Monitor_Voltage = CreateNewFpar("PEAKVOLTAGE", 33, 0)
    Set mvarResonance_Freq = CreateNewFpar("FREQ", 34, 0)
    Set mvarPeak_Ramp_Voltage = CreateNewFpar("AC_AMPL_LIMIT", 35, 0)
    Set mvarMax_Ramp_Voltage = CreateNewFpar("MAX_RAMPVOLTAGE", 36, 0)
    Set mvarMax_Monitor_Voltage = CreateNewFpar("MAX_PEAKVOLTAGE", 37, 0)
        
    Set mvarRamp_Mode = CreateNewLpar("RAMPMODE", 31, 0)
    Set mvarOutput_Port_Number = CreateNewLpar("PORT_SINEOUT", 32, 0)
    Set mvarMonitor_Port_Number = CreateNewLpar("PORT_ACCUR", 33, 0)
    Set mvarProcess_Delay = CreateNewLpar("AFRAMP_PD", 34, 0)
    Set mvarNoise_Level = CreateNewLpar("NOISELEVEL", 35, 0)
    Set mvarNumber_Periods_Hang_At_Peak = CreateNewLpar("PEAKDELAY_PERIODS", 36, 0)
    Set mvarNumber_Periods_Ramp_Down = CreateNewLpar("NUMPERIODS", 37, 0)
    Set mvarRamp_Down_Mode = CreateNewLpar("RAMPDOWN_MODE", 38, 0)

End Sub




