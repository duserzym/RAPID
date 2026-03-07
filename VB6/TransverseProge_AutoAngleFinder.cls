VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ProbeAngleOptimizer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private mvarAFCoilTemp1 As Double 'local copy
Private mvarAFCoilTemp2 As Double 'local copy
Private mvarAF_TargetMonitorVoltage As Double 'local copy
Private mvarAF_PeakHangTime As Double 'local copy
Private mvarCurrentAngle As Double 'local copy
Private mvarStartAngle As Double 'local copy
Private mvarEndAngle As Double 'local copy
Private mvarAngleStepSizeInDegrees As Double 'local copy
Private mvarangles_vs_fields As AngleVsFieldCollection 'local copy
Private mvarAngleScanStatus As AngleScanStatusEnum
Private mvarAngleScanCalibrationMode As AngleScanCalibrationModeEnum

Private mvarerror_str As String
Private manual_mode As Boolean
Private retry_last_step As Boolean

Public Event AngleScanStatusChange(ByVal angle_scan_status As AngleScanStatusEnum)
Public Event ProgressUpdate(ByVal current_angle As Double)

Public Enum AngleScanCalibrationModeEnum
    AF = 0
    IRM = 1
End Enum

Public Enum AngleScanStatusEnum
    IDLE = 0
    RUNNING = 1
    PAUSED = 2
    ABORTED = 3
End Enum

Public Property Let AF_PeakHangTime(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.AF_TargetMonitorVoltage = 5
    mvarAF_PeakHangTime = vData
End Property

Public Property Get AF_PeakHangTime() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.AF_TargetMonitorVoltage
    AF_PeakHangTime = mvarAF_PeakHangTime
End Property

Public Property Let AF_TargetMonitorVoltage(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.AF_TargetMonitorVoltage = 5
    mvarAF_TargetMonitorVoltage = vData
End Property

Public Property Get AF_TargetMonitorVoltage() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.AF_TargetMonitorVoltage
    AF_TargetMonitorVoltage = mvarAF_TargetMonitorVoltage
End Property

Public Property Let AFCoilTemp1(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.AFCoilTemp1 = 5
    mvarAFCoilTemp1 = vData
End Property

Public Property Get AFCoilTemp1() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.AFCoilTemp1
    AFCoilTemp1 = mvarAFCoilTemp1
End Property

Public Property Let AFCoilTemp2(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.AFCoilTemp2 = 5
    mvarAFCoilTemp2 = vData
End Property

Public Property Get AFCoilTemp2() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.AFCoilTemp2
    AFCoilTemp2 = mvarAFCoilTemp2
End Property

Public Property Set angles_vs_fields(ByVal vData As AngleVsFieldCollection)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.angles_vs_fields = Form1
    Set mvarangles_vs_fields = vData
End Property

Public Property Get angles_vs_fields() As AngleVsFieldCollection
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.angles_vs_fields
    Set angles_vs_fields = mvarangles_vs_fields
End Property

Public Property Let AngleScanCalibrationMode(ByVal vData As AngleScanCalibrationModeEnum)

    mvarAngleScanCalibrationMode = vData

End Property

Public Property Get AngleScanCalibrationMode() As AngleScanCalibrationModeEnum

    AngleScanCalibrationMode = mvarAngleScanCalibrationMode
    
End Property

Public Property Let AngleScanStatus(ByVal vData As AngleScanStatusEnum)

    mvarAngleScanStatus = vData

End Property

Public Property Get AngleScanStatus() As AngleScanStatusEnum

    AngleScanStatus = mvarAngleScanStatus
    
End Property

Public Property Let AngleStepSizeInDegrees(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.AngleStepSizeInDegrees = 5
    mvarAngleStepSizeInDegrees = vData
End Property

Public Property Get AngleStepSizeInDegrees() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.AngleStepSizeInDegrees
    AngleStepSizeInDegrees = mvarAngleStepSizeInDegrees
End Property

Public Property Let CurrentAngle(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CurrentAngle = 5
    mvarCurrentAngle = vData
End Property

Public Property Get CurrentAngle() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CurrentAngle
    CurrentAngle = mvarCurrentAngle
End Property

Public Property Let EndAngle(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.EndAngle = 5
    mvarEndAngle = vData
End Property

Public Property Get EndAngle() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.EndAngle
    EndAngle = mvarEndAngle
End Property

Public Property Get ErrorString() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.AngleStepSizeInDegrees
    ErrorString = mvarerror_str
End Property

Public Property Let StartAngle(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.StartAngle = 5
    mvarStartAngle = vData
End Property

Public Property Get StartAngle() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.StartAngle
    StartAngle = mvarStartAngle
End Property

Public Function AreValidAngles() As Boolean

    AreValidAngles = True

    If Me.AngleStepSizeInDegrees = 0 Then
    
        mvarerror_str = mvarerror_str & vbCrLf & "Angle Step Size cannot be zero."
        AreValidAngles = False
        
    End If
        
    If Me.StartAngle = Me.EndAngle Then
    
        mvarerror_str = mvarerror_str & vbCrLf & "Start Angle cannot be equal to End Angle."
        AreValidAngles = False

    End If

    If Me.AngleStepSizeInDegrees < 0 Then
    
        If Me.StartAngle < Me.EndAngle Then
            
            mvarerror_str = mvarerror_str & vbCrLf & "If the angle step size is negative, the start angle must be greater than the end angle."
            AreValidAngles = False
            
        End If
        
    Else
    
        If Me.StartAngle > Me.EndAngle Then
            
            mvarerror_str = mvarerror_str & vbCrLf & "If the angle step size is positive, the start angle must be less than the end angle."
            AreValidAngles = False
            
        End If
        
    End If
    
End Function

Public Function AreValidClassParameters() As Boolean

    Dim ret_val As Boolean: ret_val = True
    
    mvarerror_str = ""
    
    ret_val = ret_val And AreValidAngles
    ret_val = ret_val And IsValidMonitorVoltage
    ret_val = ret_val And IsValidPeakHangTime
    
    AreValidClassParameters = ret_val
    
End Function

Private Sub Class_Initialize()
    Me.AngleScanStatus = IDLE
    On Error Resume Next
    Me.CurrentAngle = frmDCMotors.TurningMotorAngle
    On Error GoTo 0
    Me.StartAngle = Me.CurrentAngle - 10
    Me.EndAngle = Me.CurrentAngle + 10
    Me.AngleStepSizeInDegrees = 5
    Me.AFCoilTemp1 = 0
    Me.AFCoilTemp2 = 0
    Set Me.angles_vs_fields = New AngleVsFieldCollection
    Me.AF_TargetMonitorVoltage = 0
    Me.AngleScanCalibrationMode = AF
    manual_mode = False
End Sub

Private Sub DoFieldGeneration()

    'Now start the Ramp - depending on the AF System being used
    If AFSystem = "ADWIN" And _
       Me.AngleScanCalibrationMode = AF _
    Then
        
        frmADWIN_AF.ExecuteRamp TransverseCoilSystem, _
                                Me.AF_TargetMonitorVoltage, _
                                , , , _
                                Me.AF_PeakHangTime, _
                                False, _
                                False, _
                                False
                
    ElseIf AFSystem = "2G" And _
           Me.AngleScanCalibrationMode = AF _
    Then
    
        'Execute a combo ramp
        'the uncalibrated amp and coil have already been set
        
        'Set RampFailCounter = 0
        RampFailCounter = 0
        
        'Loop until get good ramp, if has looped 5 times, then
        'send error and prompt user to cancel the calibration
        Do While frmAF_2G.ExecuteRamp("C") = False
        
            'Update the fail counter
            RampFailCounter = RampFailCounter + 1
            
            'Check to see if the ramp fail counter >=5
            If RampFailCounter >= 5 Then
            
                'Pop-up message box to the user
                MessageString = "2G AF Ramp has failed five times in a row." & _
                                vbNewLine & vbNewLine & _
                                "Would you like to continue trying to ramp?" & _
                                vbNewLine & vbNewLine & _
                                "If you answer 'No' the current auto-calibration session " & _
                                "will be aborted."
                                
                user_resp = frmDialog.DialogBox(MessageString, _
                                                   "AF Error!!", _
                                                   2, _
                                                   "Yes", _
                                                   "No")
                                                   
                If user_resp = vbNo Then
                
                    Me.AngleScanStatus = ABORTED
                    RaiseEvent AngleScanStatusChange(ABORTED)
                    
                    Exit Sub
                                        
                End If
                
            End If
                                    
        Loop
                
    ElseIf Me.AngleScanCalibrationMode = IRM Then
    
        'Click the uncalibrated volts IRM Fire button
        frmIRMARM.cmdIRMFire_Click
        
    End If

End Sub

Public Function FindBestTransverseProbeAngle() As Double

    'Clear AngleVsField collection
    Me.angles_vs_fields.Clear

    'Check Comm Mode & Prog_Halt
    If modProg.NOCOMM_MODE Or modFlow.Prog_halted Then Exit Function
    If Me.AngleScanStatus = ABORTED Then Exit Function
            
    'Verify class parameters
    If Not AreValidClassParameters Then
    
        Err.Raise -666, _
                  "ProbeAngleOptimizer.FindBestTransverseProbeAngle", _
                  mvarerror_str
        
        Exit Function
        
    End If
    
    'Connect to Gaussmeter
    'Load the gaussmeter form without showing it using special public subroutine
    Load frm908AGaussmeter
    
    Dim user_resp As VbMsgBoxResult
    
    'Prompt User to attach correct probe to the gaussmeter and turn it on
    MsgBox "While the power to the Gaussmeter is turned off and the USB-mini cable " & _
            "is NOT connected, connect the " & ProbeString & " probe." & vbNewLine & _
            "Then re-connect the USB-mini cable and WAIT for the Gaussmeter to switch." & _
            "back on.", , _
            "908A Gaussmeter Setup"
            
    'Wait and Prompt again to make sure the Gaussmeter is all the way on
    'Loop until
    Do
    
        PauseTill timeGetTime() + 500
    
        user_resp = MsgBox("Is the Gaussmeter all the way on and displaying data?", _
                                vbYesNoCancel, _
                                "908A Gaussmeter Setup")
                                
        UpdateScanStatus_BasedOnProgStatus
        WaitForAngleScanStatus_RunningOrAborted
        
                       
    Loop Until user_resp <> vbNo
    
    'If user has selected to cancel, then exit the sub-routine
    If user_resp = vbCancel Then
    
        Exit Function
        
    End If
    
    
    'Now connect the gaussmeter
    Dim TempL As Long
    TempL = frm908AGaussmeter.Connect
    
    If TempL <= 0 Then
    
        'Gaussmeter must not be connected
        'Prompt user to ask if they would like to continue with
        'calibration in Manual mode
        user_resp = MsgBox("Unable to communicate with the 908A Gaussmeter." & vbNewLine & _
                          vbNewLine & "Would you like to continue with the " & GetCoilString & _
                          " coil calibration in manual mode?", _
                          vbYesNo, _
                          "Whoops!")
        
        If user_resp = vbNo Then
        
            Me.AngleScanStatus = ABORTED
            RaiseEvent AngleScanStatusChange(ABORTED)
                    
            Exit Function
            
        ElseIf user_resp = vbYes Then
        
            manual_mode = True
            
        End If
                
    End If
    
    If manual_mode Then
    
        ManualSetupGaussmeter
        UpdateScanStatus_BasedOnProgStatus
        WaitForAngleScanStatus_RunningOrAborted
        
    End If
    
    ActiveCoilSystem = TransverseCoilSystem
    
    'Set relays for correct coil config
    If Me.AngleScanCalibrationMode = AF Then
    
        If modConfig.AFSystem = "2G" Then
        
            frmAF_2G.chkLocked.Value = Unchecked
            CoilsLocked = False
            
            frmAF_2G.SetActiveCoilSystem TransverseCoilSystem
            
            frmAF_2G.chkLocked.Value = Checked
            CoilsLocked = True
        
        ElseIf modConfig.AFSystem = "ADWIN" Then
        
            frmADWIN_AF.chkLockCoils.Value = Unchecked
            CoilsLocked = False
            
            frmADWIN_AF.optCoil(1).Value = True
            frmADWIN_AF.SetAFRelays
            frmADWIN_AF.chkLockCoils.Value = Checked
            CoilsLocked = True
        
        End If
        
    ElseIf Me.AngleScanCalibrationMode = IRM And _
           modConfig.EnableTransIRM _
    Then
    
       frmIRMARM.chkLockCoils.Value = Unchecked
       CoilsLocked = False
       frmIRMARM.optCoil(1).Value = True
       frmIRMARM.SetRelaysForIRM
       frmIRMARM.chkLockCoils.Value = Checked
       CoilsLocked = True
        
    End If
        
    
        
    'Verify X-verse probe is in the AF region (correct up/down motor position)
    Dim cur_updown_pos As Long
    cur_updown_pos = frmDCMotors.GetUpDownPos
    
    If Abs(cur_updown_pos) - modConfig.UpDownMotor1cm > Abs(modConfig.AFPos) Or _
       Abs(cur_updown_pos) + modConfig.UpDownMotor1cm < Abs(modConfig.AFPos) _
    Then
    
        'Message box user and tell them to move the Quartz glass tube to
        'the AF position
        user_resp = MsgBox("Up/Down rod is not in the AF region." & _
                   vbNewLine & vbNewLine & _
                   "Would you like to pause the AF calibration and move the " & _
                   "Up/Down rod into position?", _
                   vbYesNo, _
                   "Whoops!")
                   
        If user_resp = vbYes Then
        
            'Pause the Calibration
            Me.AngleScanStatus = PAUSED
            RaiseEvent AngleScanStatusChange(Me.AngleScanStatus)
            
        End If
                
        UpdateScanStatus_BasedOnProgStatus
        WaitForAngleScanStatus_RunningOrAborted
    
    End If
    
    'Loop through each test angle (start angle --> end angle using angle step size)
    '   1) Get Peak Field at this angle
    '   2) Save to AngleVsField collection
    Dim angle As Double
    For angle = Me.StartAngle To Me.EndAngle Step Me.AngleStepSizeInDegrees
    
        UpdateScanStatus_BasedOnProgStatus
        WaitForAngleScanStatus_RunningOrAborted
                
        frmDCMotors.TurningMotorRotate angle, True, True
        
        UpdateScanStatus_BasedOnProgStatus
        WaitForAngleScanStatus_RunningOrAborted
        
        Me.CurrentAngle = frmDCMotors.TurningMotorAngle
    
        UpdateScanStatus_BasedOnProgStatus
        WaitForAngleScanStatus_RunningOrAborted
    
        RaiseEvent ProgressUpdate(Me.CurrentAngle)
            
        Do
            Dim dc_peak_field As Double
            dc_peak_field = GetPeakFieldForCurrentTransverseProbeAngle()
            
        Loop Until Not retry_last_step
   
        'Store dc_peak_field + angle to angle vs field collection
        If Me.angles_vs_fields Is Nothing Then Set Me.angles_vs_fields = New AngleVsFieldCollection
        
        Dim angle_field As AngleVsField_Point
        Set angle_field = New AngleVsField_Point
        
        angle_field.angle = Me.CurrentAngle
        angle_field.peak_field = dc_peak_field
        
        Me.angles_vs_fields.Add angle_field.angle, angle_field.peak_field
        
        Set angle_field = Nothing
           
    Next angle
    
    'Get Max Field in AngleVSField Collection and return the matching angle
    FindBestTransverseProbeAngle = GetAngleWithMaxDCField()
    
    Me.AngleScanStatus = IDLE
    RaiseEvent AngleScanStatusChange(IDLE)

End Function

Public Function GetAngleWithMaxDCField() As Double

    If Me.angles_vs_fields Is Nothing Then Exit Function
    If Me.angles_vs_fields.Count <= 0 Then Exit Function
    
    Dim best_angle As Double
    Dim max_dc_field As Double: max_dc_field = 0
    Dim i As Integer
    
    For i = 1 To Me.angles_vs_fields.Count
    
        If max_dc_field < Abs(Me.angles_vs_fields(i).peak_field) Then
            max_dc_field = Abs(Me.angles_vs_fields(i).peak_field)
            best_angle = Me.angles_vs_fields(i).angle
        End If
        
    Next i
        
    GetAngleWithMaxDCField = best_angle
    
End Function

Private Function GetDCPeakFieldFrom908AGaussmeter() As Double

    'Get DC Peak Field from gaussmeter
    Dim dc_peak_field As Double
    Dim field_string As String
    
    If Not manual_mode Then
    
        'Now collect a data-point from the Gaussmeter
        frm908AGaussmeter.cmdSampleNow_Click
        
        'Now get the last data point converted to a string with
        'respect to the modconfig.afunits we're using
        frm908AGaussmeter.ConvertLastData field_string, modConfig.AFUnits
        
        dc_peak_field = val(field_string)
        
        'Now get rid of the last data point
        frm908AGaussmeter.cmdClearData_Click
        
        'Reset the gaussmeter DC-peak field
        frm908AGaussmeter.cmdResetPeak_Click
        
        'Wait 2000 ms
        PauseTill timeGetTime() + 2000
        
    Else
    
        Dim input_resp As String
        
        'Loop until the user clicks the OK or Re-Do button
        Do
        
            input_resp = InputBox("Please enter the resulting DC Peak Field from the Gaussmeter. " & vbCrLf & _
                                  "Please enter positive values only." & vbCrLf & _
                                 vbCrLf & _
                                 "If you need to redo the last " & _
                                 Trim(Str(IIf(Me.AngleScanCalibrationMode = AF, "AF Ramp Cycle", "IRM Pulse"))) & _
                                 "then enter REDO as the field value.", _
                                 "Peak Field Input", "0")
        
            DoEvents
            
            
            
            'Pause for 20 ms
            PauseTill timeGetTime() + 20
        
        Loop Until input_resp <> "0" And _
                   (input_resp = "REDO" Or _
                    (IsNumeric(input_resp) And val(input_resp) > 0))
        
                   
        If input_resp = "REDO" Then
            retry_last_step = True
            dc_peak_field = 0
        Else
        
            dc_peak_field = val(input_resp)
        
        End If
            
    End If
    
    GetDCPeakFieldFrom908AGaussmeter = dc_peak_field

End Function

Public Function GetPeakFieldForCurrentTransverseProbeAngle() As Double

    retry_last_step = False

    'Reset Gaussmeter DC Peak field
    ResetGaussmeter
    
    UpdateScanStatus_BasedOnProgStatus
    WaitForAngleScanStatus_RunningOrAborted
    
    'Pause 1 second (1000 ms)
    PauseTill timeGetTime() + 1000
    
    'Do IRM Pulse or AF Ramp Cycle
    DoFieldGeneration
                            
    UpdateScanStatus_BasedOnProgStatus
    WaitForAngleScanStatus_RunningOrAborted
                                
    'Wait 0.5 seconds
    PauseTill timeGetTime() + 500
                                
    GetPeakFieldForCurrentTransverseProbeAngle = GetDCPeakFieldFrom908AGaussmeter
        
End Function

Public Function GetPoint_WithMaxDCField() As AngleVsField_Point

    Set GetPoint_WithMaxDCField = Nothing

    If Me.angles_vs_fields Is Nothing Then Exit Function
    If Me.angles_vs_fields.Count <= 0 Then Exit Function
        
    Set GetPoint_WithMaxDCField = New AngleVsField_Point
    GetPoint_WithMaxDCField.angle = 0
    GetPoint_WithMaxDCField.peak_field = 0
    
    Dim i As Integer
    
    For i = 1 To Me.angles_vs_fields.Count
    
        If GetPoint_WithMaxDCField.peak_field < Abs(Me.angles_vs_fields(i).peak_field) Then
            GetPoint_WithMaxDCField.peak_field = Abs(Me.angles_vs_fields(i).peak_field)
            GetPoint_WithMaxDCField.angle = Me.angles_vs_fields(i).angle
        End If
        
    Next i
    
End Function

Public Function GetPoint_WithMinDCField() As AngleVsField_Point

    Set GetPoint_WithMinDCField = Nothing

    If Me.angles_vs_fields Is Nothing Then Exit Function
    If Me.angles_vs_fields.Count <= 0 Then Exit Function
        
    Set GetPoint_WithMinDCField = New AngleVsField_Point
    GetPoint_WithMinDCField.angle = 0
    GetPoint_WithMinDCField.peak_field = modProg.MAX_DOUBLE
    
    Dim i As Integer
    
    For i = 1 To Me.angles_vs_fields.Count
    
        If GetPoint_WithMinDCField.peak_field > Abs(Me.angles_vs_fields(i).peak_field) Then
            GetPoint_WithMinDCField.peak_field = Abs(Me.angles_vs_fields(i).peak_field)
            GetPoint_WithMinDCField.angle = Me.angles_vs_fields(i).angle
        End If
        
    Next i
    
End Function

Public Function IsValidMonitorVoltage() As Boolean

    IsValidMonitorVoltage = True

    If Me.AF_TargetMonitorVoltage <= 0 Then
    
        mvarerror_str = mvarerror_str & "Target AF Monitor voltage must be greater than zero."
        IsValidMonitorVoltage = False
        
    End If
    
    If Me.AF_TargetMonitorVoltage > modConfig.AfTransMonMax Then
    
        mvarerror_str = mvarerror_str & "Target AF Monitor voltage must be less than or equal " & _
                    "to the Transverse AF Max allowed Monitor voltage: " & _
                    Trim(Str(modConfig.AfTransMonMax))
        IsValidMonitorVoltage = False
        
    End If

End Function

Public Function IsValidPeakHangTime() As Boolean

    IsValidPeakHangTime = False
    
    If Me.AF_PeakHangTime <= 0 Then
    
        mvarerror_str = mvarerror_str & vbCrLf & "AF Ramp Peak Hang Time must be greater than zero"
        Exit Function
        
    End If
    
    'Set minimum value for peak hang time
    If Me.AF_PeakHangTime < 100 Then Me.AF_PeakHangTime = 100
        
    IsValidPeakHangTime = True
        
End Function

Private Sub ManualSetupGaussmeter()

    If Not manual_mode Then Exit Sub
    
    'We're in manual mode
    'Prompt User to attach correct probe to the gaussmeter and turn it on
    MsgBox "While the power to the Gaussmeter is turned off connect the " & _
           ProbeString & " probe." & vbNewLine & _
           "Then turn the Gaussmeter power on and wait for it to reload.", , _
            "908A Gaussmeter Setup"
    
    'Determine what the range-string should be
    If modConfig.AFUnits = "G" Then RangeString = "000.0 G"
    If modConfig.AFUnits = "Oe" Then RangeString = "000.0 Oe"
    If modConfig.AFUnits = "mT" Then RangeString = "00.00 mT"
    If modConfig.AFUnits = "kA/m" Then RangeString = "00.00 kA/m"
    
    MessageString = "When the Gaussmeter has finished loading:" & vbNewLine & vbNewLine & _
                     " 1) Place the tip of the " & ProbeString & " probe in the " & _
                     "center of the " & GetCoilString & " coil." & vbNewLine & _
                     " 2) Press the 'Menu' button on the probe." & vbNewLine & _
                     " 3) If the word ""Function"" does not appear in the lower left-hand corner " & _
                     "of the gaussmeter display, then press the 'Next' button until it does." & _
                     vbNewLine & _
                     " 4) With the word ""Function"" showing, click the 'Enter' button." & _
                     vbNewLine & _
                     " 5) Press 'Next' until the display in the lower-left corner reads " & _
                     """DC peak""." & vbNewLine & _
                     " 6) Press 'Enter' to set the gaussmeter to DC Peak Field mode." & _
                     vbNewLine & _
                     " 7) If the gaussmeter is not in " & modConfig.GetLongUnits & " units mode " & _
                     "then click 'Menu' and then click 'Next' until ""Units"" appears in the " & _
                     "lower left-hand corner." & vbNewLine & _
                     " 8) Click 'Enter' to access the ""Units"" menu.  Click next until you see " & _
                     modConfig.GetLongUnits & " displayed in the lower left-hand corner. " & _
                     " 9) Click 'Enter' again to set the units to " & modConfig.GetLongUnits & vbNewLine & _
                     "10) Check the Gaussmeter range. The display should read: """ & RangeString & """" & _
                     ". If not, then click the 'Range' button until the range changes to match: """ & _
                     RangeString & """" & vbNewLine
                     
    MessageString = MessageString & _
                    "11) Click the Reset button on the Gaussmeter and wait for the Displayed DC Peak value " & _
                    "to stabilize (about 1 - 10 seconds)" & vbNewLine & _
                    "12) Click the NULL button on the Gaussmeter and wait for the NULL process to complete" & _
                    vbNewLine & vbNewLine & _
                     "When you've completed all this, click the ""Continue"" button on this window."
                             
    user_resp = frmDialog.DialogBox(MessageString, _
                                       "908A Gaussmeter Setup", _
                                       3, _
                                       "Continue", _
                                       "Cancel")
                                      
    'Check for a negative response
    If user_resp = vbNo Then
   
        'User clicked cancel, end the calibration and exit the sub
        Me.AngleScanStatus = ABORTED
        RaiseEvent AngleScanStatusChange(ABORTED)
        
        Exit Sub
        
    End If

    

End Sub


Private Sub ResetGaussmeter()

    Dim user_resp As VbMsgBoxResult

    If manual_mode = False Then
        
        'Reset the gaussmeter
        frm908AGaussmeter.cmdResetPeak_Click
        
        'Wait 2 second ( 2000 ms)
        PauseTill timeGetTime() + 2000
            
    Else
                        
        MessageString = "Click the Reset button on the Gaussmeter and wait for the Displayed DC Peak value " & _
                        "to stabilize (about 1 - 10 seconds)" & vbNewLine & _
                         "When you've this, click the ""Continue"" button on this window."
                                 
        user_resp = frmDialog.DialogBox(MessageString, _
                                           "908A Gaussmeter Reset", _
                                           1, _
                                           "Continue")
              
    End If

End Sub

Public Sub UpdateScanStatus_BasedOnProgStatus()

    If modFlow.Prog_paused Then
        
        Me.AngleScanStatus = PAUSED
        RaiseEvent AngleScanStatusChange(PAUSED)
        
    End If
    
    If modFlow.Prog_halted Then
    
        Me.AngleScanStatus = ABORTED
        RaiseEvent AngleScanStatusChange(ABORTED)
    
    End If

End Sub

Private Sub WaitForAngleScanStatus_RunningOrAborted()

    DoEvents

    'Check to see which mode of calibration activity that we're in
    'running, paused, or end
    If Me.AngleScanStatus = PAUSED Then
    
        'Loop until Me.AngleScanStatus = RUNNING or END
        Do
        
            PauseTill timeGetTime() + 200
            
        Loop Until Me.AngleScanStatus <> PAUSED
        
    ElseIf Me.AngleScanStatus = ABORTED Then
    
        'This additional check is in here to catch if the user has clicked
        'end after clicking pause after the code prior to the pause code
        'was executed
    
        RaiseEvent AngleScanStatusChange(ABORTED)
    
        'Immediately end the subroutine
        Exit Sub
    
    End If

End Sub

