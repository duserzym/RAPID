VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RockmagStep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvarLevel As Double 'local copy
Private mvarBiasField As Double 'local copy
Private mvarStepType As String 'local copy
Private mvarSpinSpeed As Single 'local copy
Private mvarHoldTime As Single 'local copy
Public Measure As Boolean
Public MeasureSusceptibility As Boolean
Private mvarRemarks As String ' (November 2007 L Carporzen) Remarks column in RMG
Private mvarKey As String
Private Const RockmagStepAFmax = "AFmax"
Private Const RockmagStepAFz = "AFz"
Private Const RockmagStepAF = "AF"
Private Const RockmagStepUAFX1 = "UAFX1"
Private Const RockmagStepUAFX2 = "UAFX2"
Private Const RockmagStepUAFZ1 = "UAFZ1"
Private Const RockmagStepaTAFX = "aTAFX" ' (February 2010 L Carporzen) Measure the TAF and uncorrect them in sample file
Private Const RockmagStepaTAFY = "aTAFY"
Private Const RockmagStepaTAFZ = "aTAFZ"
Private Const RockmagStepARM = "ARM"
Private Const RockmagStepPulseIRMAxial = "IRMz"
Private Const RockmagStepPulseIRMTrans = "IRMx"
Private Const RockmagStepRRM = "RRM"
Private Const RockmagStepRRMz = "RRMz"

Public Property Let BiasField(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.BiasField = 5
        mvarBiasField = vData
End Property

Public Property Get BiasField() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.BiasField
    If Not EnableARM Then
        BiasField = 0
    ElseIf mvarBiasField > ARMMax Then
        BiasField = ARMMax
    ElseIf mvarBiasField < 0 Then
        BiasField = 0
    Else
        BiasField = mvarBiasField
    End If
End Property

Public Property Get DemagStepLabel() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DemagStepType
    Dim numLength As Integer
    Dim i As Integer
    Dim setLevel As Long
    DemagStepLabel = StepType
    If Len(DemagStepLabel) > DEMAGLEN Then DemagStepLabel = Left$(DemagStepLabel, DEMAGLEN)
    numLength = DEMAGLEN - Len(DemagStepLabel)
    If StepType = RockmagStepARM Then setLevel = BiasField Else setLevel = Level
    DemagStepLabel = DemagStepLabel & Right$(Format$(Format$(setLevel, "######"), "@@@@@@"), numLength)
    If Len(DemagStepLabel) < DEMAGLEN Then
        For i = Len(DemagStepLabel) + 1 To DEMAGLEN
            DemagStepLabel = DemagStepLabel & " "
        Next i
    End If
End Property

Public Property Get DemagStepLabelLong() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DemagStepType
    Dim setLevel As Long
    DemagStepLabelLong = StepType
    If StepType = RockmagStepARM Then setLevel = BiasField Else setLevel = Level
    DemagStepLabelLong = DemagStepLabelLong & " " & str$(setLevel)
    If StepType = RockmagStepRRM Then DemagStepLabelLong = DemagStepLabelLong & "/" & SpinSpeed & " rps"
End Property

Public Property Let HoldTime(ByVal vData As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.HoldTime = 5
    mvarHoldTime = vData
End Property

Public Property Get HoldTime() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.HoldTime
    If mvarHoldTime > 0 Then
        HoldTime = mvarHoldTime
    Else
        HoldTime = 0
    End If
End Property

Public Property Let key(ByVal vData As String)
    mvarKey = vData
End Property

Public Property Get key() As String
    key = mvarKey
End Property

Public Property Let Level(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Level = 5
    mvarLevel = vData
End Property

Public Property Get Level() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Level
    Dim maxAcceptableLevel As Long
    Dim minAcceptableLevel As Long
    Select Case mvarStepType
        Case RockmagStepAF
            If AfAxialMax > AfTransMax Then maxAcceptableLevel = AfTransMax Else maxAcceptableLevel = AfAxialMax
            If AfAxialMin < AfTransMin Then minAcceptableLevel = AfAxialMin Else minAcceptableLevel = AfTransMin
            If Not EnableAF Then maxAcceptableLevel = 0
        Case RockmagStepAFmax
            If AfAxialMax > AfTransMax Then maxAcceptableLevel = AfTransMax Else maxAcceptableLevel = AfAxialMax
            If AfAxialMin < AfTransMin Then minAcceptableLevel = AfAxialMin Else minAcceptableLevel = AfTransMin
            If Not EnableAF Then maxAcceptableLevel = 0
        Case RockmagStepAFz
            maxAcceptableLevel = AfAxialMax
            minAcceptableLevel = AfAxialMin
            If Not EnableAF Then maxAcceptableLevel = 0
        Case RockmagStepUAFX1
            maxAcceptableLevel = AfTransMax
            minAcceptableLevel = AfTransMin
            If Not EnableAF Then maxAcceptableLevel = 0
        Case RockmagStepUAFX2
            maxAcceptableLevel = AfTransMax
            minAcceptableLevel = AfTransMin
            If Not EnableAF Then maxAcceptableLevel = 0
        Case RockmagStepUAFZ1
            maxAcceptableLevel = AfAxialMax
            minAcceptableLevel = AfAxialMin
            If Not EnableAF Then maxAcceptableLevel = 0
        Case RockmagStepaTAFX
            maxAcceptableLevel = AfAxialMax
            minAcceptableLevel = AfAxialMin 'AfTransMax ' (February 2010 L Carporzen) TAF above max of transverse coil
            If Not EnableAF Then maxAcceptableLevel = 0
        Case RockmagStepaTAFY
            maxAcceptableLevel = AfAxialMax
            minAcceptableLevel = AfAxialMin 'AfTransMax ' (February 2010 L Carporzen) TAF above max of transverse coil
            If Not EnableAF Then maxAcceptableLevel = 0
        Case RockmagStepaTAFZ
            maxAcceptableLevel = AfAxialMax
            minAcceptableLevel = AfAxialMin 'AfTransMax ' (February 2010 L Carporzen) TAF above max of transverse coil
            If Not EnableAF Then maxAcceptableLevel = 0
        Case RockmagStepARM
            maxAcceptableLevel = AfAxialMax
            minAcceptableLevel = AfAxialMin
            If Not EnableARM Then maxAcceptableLevel = 0
        Case RockmagStepPulseIRMAxial
            If PulseTransMax > PulseAxialMax Then maxAcceptableLevel = PulseTransMax Else maxAcceptableLevel = PulseAxialMax
            If EnableIRMBackfield Then minAcceptableLevel = -PulseAxialMax Else minAcceptableLevel = PulseAxialMin
            
            If Not EnableAxialIRM Then
                maxAcceptableLevel = 0
                minAcceptableLevel = 0
            End If
        Case RockmagStepRRM
            maxAcceptableLevel = AfTransMax
            minAcceptableLevel = AfTransMin
            If Not EnableAF Then maxAcceptableLevel = 0
        Case Else
            maxAcceptableLevel = 999999
            minAcceptableLevel = -999999
    End Select
    If StepType = RockmagStepAFmax Then
        Level = maxAcceptableLevel
    ElseIf mvarLevel = 0 Then
        Level = 0
    ElseIf mvarLevel > maxAcceptableLevel Then
        Level = maxAcceptableLevel
    ElseIf mvarLevel < minAcceptableLevel Then
        Level = minAcceptableLevel
    Else
        Level = mvarLevel
    End If
End Property

Public Property Let Remarks(ByVal vData As String)
    mvarRemarks = vData ' (November 2007 L Carporzen) Remarks column in RMG
End Property

Public Property Get Remarks() As String
    Remarks = mvarRemarks ' (November 2007 L Carporzen) Remarks column in RMG
End Property

Public Property Let SpinSpeed(ByVal vData As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SpinSpeed = 5
    mvarSpinSpeed = vData
End Property

Public Property Get SpinSpeed() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SpinSpeed
    If mvarSpinSpeed > 40 Then
        SpinSpeed = 40
    ElseIf mvarSpinSpeed < -40 Then
        SpinSpeed = -40
    Else
        SpinSpeed = mvarSpinSpeed
    End If
End Property

Public Property Let StepType(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.StepType = "AF"
    mvarStepType = vData
End Property

Public Property Get StepType() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.StepType
    StepType = mvarStepType
End Property

'------------------------------------------------------------------------------------------------------------------'
'------------------------------------------------------------------------------------------------------------------'
'
'   August 11, 2010
'   Mod to Public Sub PerformStep
'
'   Isaac Hilburn
'
'   Summary:    Added If ... then ... elseif ... then statments to select the correct AFsystem to do
'               the treatment for the rock mag step.
'               Also, cleaned up the code appearance so that it's easier to read / understand.  Also
'               added more code documentation
'------------------------------------------------------------------------------------------------------------------'
'------------------------------------------------------------------------------------------------------------------'
Public Sub PerformStep(specimen As Sample)
    
    Dim SampleCenterRMPosition As Long
    Dim X As String, bOK As Boolean
    
    If Prog_halted Then Exit Sub
    
    If ARMMax > 0 Then frmIRMARM.SetBiasField 0
    
    If Not ((mvarStepType = RockmagStepAF) Or _
            (mvarStepType = RockmagStepARM) Or _
            (mvarStepType = RockmagStepPulseIRMAxial) Or _
            (mvarStepType = RockmagStepPulseIRMTrans) Or _
            (mvarStepType = RockmagStepRRM) Or _
            (mvarStepType = RockmagStepUAFX1) Or _
            (mvarStepType = RockmagStepUAFX2) Or _
            (mvarStepType = RockmagStepUAFZ1) Or _
            (mvarStepType = RockmagStepaTAFX) Or _
            (mvarStepType = RockmagStepaTAFY) Or _
            (mvarStepType = RockmagStepaTAFZ) Or _
            (mvarStepType = RockmagStepAFmax) Or _
            (mvarStepType = RockmagStepAFz)) _
    Then
        
        Exit Sub
        
    End If
    
    '(August 2010 - I Hilburn)
    'Added in logical conditions to check to see that the necessary module
    'for a rockmag step is enabled
    
    'Check AF module
    If (mvarStepType = RockmagStepAF Or _
        mvarStepType = RockmagStepAFz Or _
        mvarStepType = RockmagStepUAFX1 Or _
        mvarStepType = RockmagStepUAFX2 Or _
        mvarStepType = RockmagStepUAFZ1 Or _
        mvarStepType = RockmagStepaTAFX Or _
        mvarStepType = RockmagStepaTAFY Or _
        mvarStepType = RockmagStepaTAFZ Or _
        mvarStepType = RockmagStepAFmax Or _
        mvarStepType = RockmagStepARM) And _
       EnableAF = False _
    Then
    
        'Wha-oh, user is trying to do an AF step without the AF module switched on
        MsgBox "AF Module is currently disabled. AF demag cannot be performed.", , _
               "Whoops!"
               
        Exit Sub
        
    End If
    
    'Check IRM Axial module
    If mvarStepType = RockmagStepPulseIRMAxial And _
       EnableAxialIRM = False _
    Then
    
        'Wha-oh, user is trying to do an IRM axial step without the module switched on
        MsgBox "Axial IRM Module is currently disabled. IRM Axial pulse cannot be performed.", , _
               "Whoops!"
               
        Exit Sub
        
    End If
    
    'Check IRM Transverse module
    If mvarStepType = RockmagStepPulseIRMTrans And _
       EnableTransIRM = False _
    Then
    
        'Wha-oh, user is trying to do an IRM transverse step without the module switched on
        MsgBox "Transverse IRM Module is currently disabled. " & _
               "IRM Transverse pulse cannot be performed.", , _
               "Whoops!"
               
        Exit Sub
        
    End If
    
    'Check ARM
    If rockmagsteptype = RockmagStepARM And _
       EnableARM = False _
    Then
    
        'MsgBox the user - ARM module is disabled
        MsgBox "ARM Module is currently disabled. " & _
               "ARM Bias Voltage cannot be used right now, " & _
               "though the AF module is enabled.", , _
               "Whoops!"
               
        Exit Sub
        
    End If
    
    If specimen.Parent.doBoth And Not specimen.Parent.doUp Then Exit Sub
    
    ' (February 2010 L Carporzen) Measure the TAF and uncorrect them in sample file
    If (mvarStepType = RockmagStepaTAFX) Or _
       (mvarStepType = RockmagStepaTAFY) Or _
       (mvarStepType = RockmagStepaTAFZ) _
    Then
            
        If frmMeasure.lblDemag = specimen.Parent.measurementSteps.CurrentStep.DemagStepLabelLong Then
        
            ' We assume that we have the same height than before
        
        Else
        
        frmDCMotors.HomeToTop
        
        frmSendMail.MailNotification "2G Status Update", _
                                     "Sample is waiting for " & mvarStepType & "."
        
        Do
            X = InputBox("Doing " & mvarStepType & _
                         " axial demagnetization." & vbCr & _
                         "What is the height (in cm) of the sample?" & _
                         vbCr & " " & vbCr & "Orientation conventions:" & vbCr & _
                         "aTAFX vertical quartz disk with arrow toward the top and " & _
                         "sample on the clean lab side of the disk." & vbCr & _
                         "aTAFY vertical quartz disk with arrow toward the oven and " & _
                         "sample on the clean lab side of the disk.", _
                         "Important!", _
                         Format$(SampleHeight / UpDownMotor1cm, "##0.0"))
                         
            Select Case X
                
                Case Is = "False"
                    
                    'MsgBox "You don't want to press Cancel."
                    bOK = False
                
                Case Is = ""
                    
                    bOK = False
                    'MsgBox "You haven't entered anything!"
                
                Case Else
                    
                    If val(X) = 0 Then
                        
                        bOK = False
                        'MsgBox "That was only Spaces!"
                    
                    Else
                        
                        bOK = True
                    
                    End If
                
            End Select
            
        Loop Until bOK = True
        
        SampleHeight = UpDownMotor1cm * val(X)
        
        specimen.SampleHeight = SampleHeight
        
        End If
        
    End If
    
    'Position the center of the sample in the center of the
    'rock-mag coils
    SampleCenterRMPosition = Int(AFPos + SampleHeight / 2)
        
    If SampleCenterRMPosition / Abs(SampleCenterRMPosition) <> AFPos / Abs(AFPos) Then
        
        ' crap... our sample is too large to put in the AF coil!
        Exit Sub
        
    End If
    
    If mvarStepType = RockmagStepPulseIRMAxial Then
        
        frmDCMotors.UpDownMove 0, 1
        
    Else
        
        '  Move somewhat slowly into AF region
        MotorUpDn_Move SampleCenterRMPosition, 1
    
    End If
    
    frmDCMotors.TurningMotorRotate 0
    
    'Set the ARM Bias field if this rock-mag step requires it
    If EnableARM And mvarBiasField > 0 And _
       ((mvarStepType = RockmagStepARM) Or _
        (mvarStepType = RockmagStepRRM) Or _
        (mvarStepType = RockmagStepRRMz)) _
    Then
        
        frmIRMARM.SetBiasField BiasField
    
    End If
    
    Select Case mvarStepType
        
        Case RockmagStepAF
                                    
            If Level > 0 And EnableAF Then
                 
                'Axial AF Ramp
                If AFSystem = "2G" Then
                    
                    ' (August 2007 L Carporzen) Allow to wait between each ramp
                    If Not frmAF_2G.txtWaitingTime = 0 Then DelayTime (frmAF_2G.txtWaitingTime)
                    
                    frmAF_2G.CycleWithHold HoldTime, AxialCoilSystem, Level, AFRampRate
                
                    frmAF_2G.Disconnect
                
                ElseIf AFSystem = "ADWIN" Then
                
                    frmADWIN_AF.ExecuteRamp AxialCoilSystem, _
                                            Level, _
                                            , , , _
                                            , _
                                            True, _
                                            False, _
                                            (frmADWIN_AF.chkVerbose.value = Checked)
                                            
                End If
                                
                Dim three_seconds As Integer
                three_seconds = 3
                
                PauseBetweenUseCoils_InSeconds three_seconds
                 
                'Transverse AF Ramp - X diretion
                If AFSystem = "2G" Then
                    
                    frmAF_2G.Connect
                    frmAF_2G.CycleWithHold HoldTime, TransverseCoilSystem, Level, AFRampRate
                    
                ElseIf AFSystem = "ADWIN" Then
                
                    frmADWIN_AF.ExecuteRamp TransverseCoilSystem, _
                                            Level, _
                                            , , , _
                                            , _
                                            True, _
                                            False, _
                                            (frmADWIN_AF.chkVerbose.value = Checked)
                                            
                End If
                
                frmDCMotors.TurningMotorRotate 90
                                
                'Transverse AF Ramp - Y direction
                If AFSystem = "2G" Then
                
                    ' (August 2007 L Carporzen) Allow to wait between each ramp
                    If Not frmAF_2G.txtWaitingTime = 0 Then DelayTime (frmAF_2G.txtWaitingTime)
                    
                    frmAF_2G.CycleWithHold HoldTime, TransverseCoilSystem, Level, AFRampRate
                                        
                ElseIf AFSystem = "ADWIN" Then
                
                    frmADWIN_AF.ExecuteRamp TransverseCoilSystem, _
                                            Level, _
                                            , , , _
                                            , _
                                            True, _
                                            False, _
                                            (frmADWIN_AF.chkVerbose.value = Checked)
                                            
                End If
                
                frmDCMotors.TurningMotorRotate 360
                
                On Error GoTo ResetGaussmeter_DCPeak_Error
                
                frm908AGaussmeter.cmdResetPeak_Click
                
                On Error GoTo 0
                
ResetGaussmeter_DCPeak_Error:
                
                'Do nothing
                
            End If
        
        
        Case RockmagStepAFmax
            
            If Level > 0 And EnableAF Then
                
                If AFSystem = "2G" Then
                    
                    frmAF_2G.Connect
                    frmAF_2G.CycleWithHold HoldTime, TransverseCoilSystem, AfTransMax, AFRampRate
                    
                ElseIf AFSystem = "ADWIN" Then
                
                    frmADWIN_AF.ExecuteRamp TransverseCoilSystem, _
                                            AfTransMax, _
                                            , , , _
                                            , _
                                            True, _
                                            False, _
                                            (frmADWIN_AF.chkVerbose.value = Checked)
                                            
                End If
                
                frmDCMotors.TurningMotorRotate 90
                
                If AFSystem = "2G" Then
                    
                    ' (August 2007 L Carporzen) Allow to wait between each ramp
                    If Not frmAF_2G.txtWaitingTime = 0 Then DelayTime (frmAF_2G.txtWaitingTime)
                
                    frmAF_2G.CycleWithHold HoldTime, TransverseCoilSystem, AfTransMax, AFRampRate
                    
                ElseIf AFSystem = "ADWIN" Then
                
                    frmADWIN_AF.ExecuteRamp TransverseCoilSystem, _
                                            AfTransMax, _
                                            , , , _
                                            , _
                                            True, _
                                            False, _
                                            (frmADWIN_AF.chkVerbose.value = Checked)
                                            
                End If
                
                frmDCMotors.TurningMotorRotate 360
                
                If AFSystem = "2G" Then
                    
                    ' (August 2007 L Carporzen) Allow to wait between each ramp
                    If Not frmAF_2G.txtWaitingTime = 0 Then DelayTime (frmAF_2G.txtWaitingTime)
                
                    frmAF_2G.CycleWithHold HoldTime, AxialCoilSystem, AfAxialMax, AFRampRate
                    
                    frmAF_2G.Disconnect
                    
                ElseIf AFSystem = "ADWIN" Then
                
                    frmADWIN_AF.ExecuteRamp AxialCoilSystem, _
                                            AfAxialMax, _
                                            , , , _
                                            , _
                                            True, _
                                            False, _
                                            (frmADWIN_AF.chkVerbose.value = Checked)
                                            
                End If
                                
            End If
        
        
        Case RockmagStepAFz
            
            If Level > 0 And EnableAF Then
                
                If AFSystem = "2G" Then
                    
                    frmAF_2G.Connect
                    
                    frmAF_2G.CycleWithHold HoldTime, AxialCoilSystem, Level, AFRampRate
                    
                    frmAF_2G.Disconnect
                    
                ElseIf AFSystem = "ADWIN" Then
                
                    frmADWIN_AF.ExecuteRamp AxialCoilSystem, _
                                            Level, _
                                            , , , _
                                            , _
                                            True, _
                                            False, _
                                            (frmADWIN_AF.chkVerbose.value = Checked)
                                            
                End If
                
            End If
        
        
        Case RockmagStepUAFX1
            
            If Level > 0 And EnableAF Then
                
                If AFSystem = "2G" Then
                    
                    frmAF_2G.Connect
                    
                    frmAF_2G.CycleWithHold HoldTime, TransverseCoilSystem, Level, AFRampRate
                    
                    frmAF_2G.Disconnect
                    
                ElseIf AFSystem = "ADWIN" Then
                
                    frmADWIN_AF.ExecuteRamp TransverseCoilSystem, _
                                            Level, _
                                            , , , _
                                            , _
                                            True, _
                                            False, _
                                            (frmADWIN_AF.chkVerbose.value = Checked)
                                            
                End If
                
            End If
        
        
        Case RockmagStepUAFX2
            
            If Level > 0 And EnableAF Then
                
                frmDCMotors.TurningMotorRotate 90
                
                If AFSystem = "2G" Then
                    
                    frmAF_2G.Connect
                    
                    frmAF_2G.CycleWithHold HoldTime, TransverseCoilSystem, Level, AFRampRate
                    
                    frmAF_2G.Disconnect
                    
                ElseIf AFSystem = "ADWIN" Then
                
                    frmADWIN_AF.ExecuteRamp TransverseCoilSystem, _
                                            Level, _
                                            , , , _
                                            , _
                                            True, _
                                            False, _
                                            (frmADWIN_AF.chkVerbose.value = Checked)
                                            
                End If
                
                frmDCMotors.TurningMotorRotate 360
                
            End If
        
        
        Case RockmagStepUAFZ1
        
            If Level > 0 And EnableAF Then
                
                If AFSystem = "2G" Then
                    
                    frmAF_2G.Connect
                    
                    frmAF_2G.CycleWithHold HoldTime, AxialCoilSystem, Level, AFRampRate
                    
                    frmAF_2G.Disconnect
                    
                ElseIf AFSystem = "ADWIN" Then
                
                    frmADWIN_AF.ExecuteRamp AxialCoilSystem, _
                                            Level, _
                                            , , , _
                                            , _
                                            True, _
                                            False, _
                                            (frmADWIN_AF.chkVerbose.value = Checked)
                                            
                End If
                
            End If
            
            
        Case RockmagStepaTAFX
            
            If Level > 0 And EnableAF Then
                
                If AFSystem = "2G" Then
                    
                    frmAF_2G.Connect
                    
                    frmAF_2G.CycleWithHold HoldTime, AxialCoilSystem, Level, AFRampRate
                    
                    frmAF_2G.Disconnect
                    
                ElseIf AFSystem = "ADWIN" Then
                
                    frmADWIN_AF.ExecuteRamp AxialCoilSystem, _
                                            Level, _
                                            , , , _
                                            , _
                                            True, _
                                            False, _
                                            (frmADWIN_AF.chkVerbose.value = Checked)
                                            
                End If
                
            End If
        
        
        Case RockmagStepaTAFY
        
            If Level > 0 And EnableAF Then
                
                If AFSystem = "2G" Then
                    
                    frmAF_2G.Connect
                    
                    frmAF_2G.CycleWithHold HoldTime, AxialCoilSystem, Level, AFRampRate
                    
                    frmAF_2G.Disconnect
                    
                ElseIf AFSystem = "ADWIN" Then
                
                    frmADWIN_AF.ExecuteRamp AxialCoilSystem, _
                                            Level, _
                                            , , , _
                                            , _
                                            True, _
                                            False, _
                                            (frmADWIN_AF.chkVerbose.value = Checked)
                                            
                End If
                
            End If
            
            
        Case RockmagStepaTAFZ
        
            If Level > 0 And EnableAF Then
                
                If AFSystem = "2G" Then
                    
                    frmAF_2G.Connect
                    
                    frmAF_2G.CycleWithHold HoldTime, AxialCoilSystem, Level, AFRampRate
                    
                    frmAF_2G.Disconnect
                    
                ElseIf AFSystem = "ADWIN" Then
                
                    frmADWIN_AF.ExecuteRamp AxialCoilSystem, _
                                            Level, _
                                            , , , _
                                            , _
                                            True, _
                                            False, _
                                            (frmADWIN_AF.chkVerbose.value = Checked)
                                            
                End If
                
            End If
            
            
        Case RockmagStepARM
            
            If Level > 0 And EnableARM Then
                
                If AFSystem = "2G" Then
                    
                    frmAF_2G.Connect
                    
                    frmAF_2G.CycleWithHold HoldTime, AxialCoilSystem, Level, AFRampRate
                    
                    frmAF_2G.Disconnect
                    
                ElseIf AFSystem = "ADWIN" Then
                
                    frmADWIN_AF.ExecuteRamp AxialCoilSystem, _
                                            Level, _
                                            , , , _
                                            , _
                                            True, _
                                            False, _
                                            (frmADWIN_AF.chkVerbose.value = Checked)
                                            
                End If
                
            Else
            
                DelayTime HoldTime
                
            End If
            
            
        Case RockmagStepPulseIRMAxial
            
            If EnableAxialIRM Then
                
                SampleCenterRMPosition = Int(frmIRMARM.IRMCenteringPos(Level) + SampleHeight / 2)
                
                If SampleCenterRMPosition / Abs(SampleCenterRMPosition) <> AFPos / Abs(AFPos) Then
                    
                    ' crap... our sample is too large to put in the AF coil!
                    Exit Sub
                
                End If
                
                ' discharge with sample in load position
                MotorUpDn_Move 0, 1
                
                frmIRMARM.optCoil(0).value = True
                frmIRMARM.FireIRM 0
                
                DelayTime 1
                
                frmIRMARM.optCoil(0).value = True
                frmIRMARM.FireIRM 0
                
                MotorUpDn_Move SampleCenterRMPosition, 1
                
                If Abs(Level) > 0 Then
                    frmIRMARM.optCoil(0).value = True
                    frmIRMARM.FireIRMAtField Level
                End If
            
            End If
            
            
        Case RockmagStepRRM
            
            If EnableAF Then
                
                frmDCMotors.TurningMotorSpin SpinSpeed, 300 + HoldTime
                
                If Level > 0 Then
                    
                    If AFSystem = "2G" And COMPortAf > 0 Then
                            
                        frmAF_2G.Connect
                        
                        frmAF_2G.CycleWithHold HoldTime, TransverseCoilSystem, Level, AFRampRate
                        
                        frmAF_2G.Disconnect
                        
                            
                    ElseIf AFSystem = "ADWIN" Then
                    
                        frmADWIN_AF.ExecuteRamp TransverseCoilSystem, _
                                                Level, _
                                                , , , _
                                                , _
                                                True, _
                                                False, _
                                                (frmADWIN_AF.chkVerbose.value = Checked)
                    
                    End If
                
                End If
                
                frmDCMotors.TurningMotorSpin 0
            
            End If
            
            
        Case RockmagStepRRMz
            
            If EnableAF Then
                
                frmDCMotors.TurningMotorSpin SpinSpeed, 300 + HoldTime
                
                If Level > 0 Then
                    
                    If AFSystem = "2G" And COMPortAf > 0 Then
                    
                        frmAF_2G.Connect
                        
                        frmAF_2G.CycleWithHold HoldTime, AxialCoilSystem, Level, AFRampRate
                        
                        frmAF_2G.Disconnect
                    
                    ElseIf AFSystem = "ADWIN" Then
                    
                        frmADWIN_AF.ExecuteRamp AxialCoilSystem, _
                                                Level, _
                                                , , , _
                                                , _
                                                True, _
                                                False, _
                                                (frmADWIN_AF.chkVerbose.value = Checked)
                                                
                    End If
                        
                End If
                
                frmDCMotors.TurningMotorSpin 0
                
            End If
            
    End Select
    
    If mvarBiasField > 0 Then
    
        frmIRMARM.SetBiasField 0
        
    End If
    
End Sub

