VERSION 5.00
Begin VB.Form frmTempSensorSettings 
   Caption         =   "AF Coil Temp. Sensor Settings"
   ClientHeight    =   4335
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5175
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4335
   ScaleWidth      =   5175
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   3840
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4935
   End
End
Attribute VB_Name = "frmTempSensorSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim isChange As Boolean
Dim LastTUnits As String
Dim CurTunits As String

Private Function GetCurTUnits() As String

    'Read the currently selected item in the Temperature Sensor
    'Units combo-box and return the matching units string
    If cmbTunits.ListIndex = 0 Then
    
        GetCurTUnits = "C"
        
    ElseIf cmbTunits.ListIndex = 1 Then
    
        GetCurTUnits = "F"
        
    ElseIf cmbTunits.ListIndex = 2 Then
    
        GetCurTUnits = "K"
        
    Else

        'Default units to "C"
        'If some strange value gets into the units-combo box,
        'set it back to Celcius
        GetCurTUnits = "C"
        cmbTunits.ListIndex = 0
        
    End If

End Function

Private Function ConvertSlope(ByVal OldUnits As String, _
                                    ByVal NewUnits As String, _
                                    ByVal OldValue As Double) As Double
                                    
    Dim TempD As Double
    
    'Celcius to Farenheit
    If OldUnits = "C" And NewUnits = "F" Then
    
        TempD = 9 / 5 * OldValue
        
    ElseIf OldUnits = "F" And NewUnits = "C" Then
    
        TempD = OldValue * 5 / 9
        
    ElseIf OldUnits = "F" And NewUnits = "K" Then
    
        TempD = OldValue * 5 / 9
    
    ElseIf OldUnits = "K" And NewUnits = "F" Then
    
        TempD = OldValue * 9 / 5
        
   End If
                                    
   'Return the new slope
   ConvertSlope = TempD
                                    
End Function

Private Function ConvertTemperature(ByVal OldUnits As String, _
                                    ByVal NewUnits As String, _
                                    ByVal OldValue As Variant) As Variant
                                    
    Dim TempD As Double
    
    'Celcius to Farenheit
    If OldUnits = "C" And NewUnits = "F" Then
    
        TempD = 9 / 5 * CDbl(OldValue) + 32
        
    ElseIf OldUnits = "C" And NewUnits = "K" Then
    
        TempD = CDbl(OldValue) + 273.15
        
    ElseIf OldUnits = "F" And NewUnits = "C" Then
    
        TempD = (CDbl(OldValue) - 32) * 5 / 9
        
    ElseIf OldUnits = "F" And NewUnits = "K" Then
    
        TempD = (CDbl(OldValue) - 32) * 5 / 9 + 273.15
    
    ElseIf OldUnits = "K" And NewUnits = "F" Then
    
        TempD = (CDbl(OldValue) - 273.15) * 9 / 5 + 32
        
    ElseIf OldUnits = "K" And NewUnits = "C" Then
    
        TempD = CDbl(OldValue) - 273.15
        
    End If
                                    
    'If the old value was an integer, then output an integer,
    'otherwise, output a double
    If VarType(OldValue) = vbInteger Then
        
            ConvertTemperature = CInt(TempD)
            
        Else
        
            ConvertTemperature = TempD
            
    End If
                                    
End Function


Private Sub cmbTunits_Click()

    Dim CurTunits As String
    Dim LocalThot As Integer
    Dim LocalTmax As Integer
    Dim LocalTslope As Double
    Dim LocalToffset As Double
    
    CurTunits = GetCurTUnits

    'Check to see if this new units value is different from the current system units value
    If CurTunits <> LastTUnits Then
    
        'Load the four form values from the text-box controls on the form
        LocalThot = CInt(Me.txtThot)
        LocalTmax = CInt(Me.txtTmax)
        LocalTslope = val(Me.txtTslope)
        LocalToffset = val(Me.txtToffset)
    
        'Convert the values
        LocalThot = ConvertTemperature(LastTUnits, CurTunits, LocalThot)
        LocalTmax = ConvertTemperature(LastTUnits, CurTunits, LocalTmax)
        LocalToffset = ConvertTemperature(LastTUnits, CurTunits, LocalToffset)
        LocalTslope = ConvertSlope(LastTUnits, CurTunits, LocalTslope)
        
        'Now update the form text-box controls with the converted values
        Me.txtThot = Trim(Str(LocalThot))
        Me.txtTmax = Trim(Str(LocalTmax))
        Me.txtToffset = Trim(Str(LocalToffset))
        Me.txtTslope = Trim(Str(LocalTslope))
                                    
        'Now save the current units into the Last units variable
        LastTUnits = CurTunits
                            
    End If

End Sub

Private Sub cmdClose_Click()

    'default isChange to false
    isChange = False

    'Check for a change in the temperature sensor settings
    
    'Get the currently displayed temperature sensor units
    CurTunits = GetCurTUnits
    
    'Compare all of the diaplyed values to the system values
    'if there are any differences, set isChange = True
    If CurTunits <> modConfig.Tunits Or _
       CInt(Me.txtThot) <> modConfig.Thot Or _
       CInt(Me.txtTmax) <> modConfig.Tmax Or _
       val(Me.txtToffset) <> modConfig.Toffset Or _
       val(Me.txtTslope) <> modConfig.TSlope _
    Then
    
        isChange = True
        
    End If
    
    If isChange = True Then
    
        'Tell user that they need to save changes in the settings form
        MsgBox "Remember to save your Temp. Sensor settings changes by clicking " & _
               "the ""Apply"" or ""OK"" buttons in the main Settings window.", , _
               "Reminder"
        
    End If
    
    Me.Hide
    
End Sub

Private Sub Form_Load()

    'Store the Current value of the Temperature Sensor units into
    'the Last TUnits field
    LastTUnits = modConfig.Tunits
    
    'Clear and re-setup the items in the Temperature sensor units combo-box
    cmbTunits.Clear
    cmbTunits.AddItem "Celsius, °C", 0
    cmbTunits.AddItem "Farenheit, °F", 1
    cmbTunits.AddItem "Kelvin, K", 2
    
    'Load the Tunits into the units combo-box
    Select Case modConfig.Tunits
    
        Case "C"
        
            cmbTunits.ListIndex = 0
            
        Case "F"
        
            cmbTunits.ListIndex = 1
            
        Case "K"
        
            cmbTunits.ListIndex = 2
            
        Case Else
        
            'If weird value in units field, reset to Celcius
            cmbTunits.ListIndex = 0
            
    End Select
    
    'Change all of the units labels on the form
    Me.lblTHotUnits = "°" & modConfig.Tunits
    Me.lblTMaxUnits = "°" & modConfig.Tunits
    Me.lblTOffsetUnits = "°" & modConfig.Tunits
    Me.lblTSlopeUnits = "°" & modConfig.Tunits & " / V"
    
    'Now load the "Hot" temperature alarm value
    Me.txtThot = modConfig.Thot
    
    'Now load the Maximum allowed AF coil temperature before the AF ramp
    'is halted
    Me.txtTmax = modConfig.Tmax
    
    'Now load the Temperature offset for converted the thermocouple voltage
    Me.txtToffset = modConfig.Toffset
    
    'Now load the Temperature / Voltage slope for converted the thermocouple voltage
    Me.txtTslope = modConfig.TSlope

End Sub
