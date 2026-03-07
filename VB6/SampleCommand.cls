VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SampleCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Sample"
'local variable(s) to hold property value(s)
Private mvarKey As String 'local copy
Private mvarcommandType As String 'local copy
Private mvarFileID As String
Private mvarSample As String
Private mvarHole As Long 'local copy
Private Const commandMeasure = "Meas"
Private Const commandInitUp = "InitUp"
Private Const commandHolder = "Holder"
Private Const commandFlip = "Flip"
Private Const commandFin = "Fin"
Private Const commandGoto = "Goto"

Public Property Let commandType(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.commandType = 5
    mvarcommandType = vData
End Property

Public Property Get commandType() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.commandType
    commandType = mvarcommandType
End Property

Public Property Let fileid(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Hole = 5
    mvarFileID = vbNullString
    If SampleIndexRegistry.IsValidFile(vData) Then _
        mvarFileID = SampleIndexRegistry(vData).filename
End Property

Public Property Get fileid() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Hole
    fileid = mvarFileID
End Property

Public Property Let hole(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Hole = 5
    If (vData < SlotMin) Or (vData > SlotMax) Then
    
        If UseXYTableAPS And vData = -1 Then
            mvarHole = -1
        Else
            mvarHole = 0
        End If
    Else
        mvarHole = vData
    End If
End Property

Public Property Get hole() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Hole
    hole = mvarHole
End Property

Public Property Let key(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Key = 5
    mvarKey = vData
End Property

Public Property Get key() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Key
    key = mvarKey
End Property

Public Property Let Sample(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Hole = 5
On Error GoTo oops
    mvarSample = SampleIndexRegistry(mvarFileID).sampleSet(vData).Samplename
    On Error GoTo 0
    Exit Property
oops:
    mvarSample = vbNullString
End Property

Public Property Get Sample() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Hole
    Sample = mvarSample
End Property

Public Sub Execute()
    Dim NotificationMessage As String
    Do Until Not Prog_paused
        If Prog_halted Then ' (September 2007 L Carporzen) New version of the Halt button
            Flow_Resume
            frmMeasure.updateFlowStatus
            Exit Sub
        End If
        DelayTime 0.5
    Loop
    frmMeasure.HideStats
    frmMeasure.clearData
    If DEBUG_MODE Then frmDebug.Msg "Executing from queue of " & SampleQueueCommands & _
        ": " & commandType & " hole " & hole
    Select Case mvarcommandType
        Case commandMeasure
            'frmMeasure.SetFieldsFromRegistry fileid
            MeasurementsSinceHolder = MeasurementsSinceHolder + 1
            modConfig.SampleHandlerCurrentHole = hole
            frmVacuum.MotorPower True
            frmMeasure.SetSample mvarSample
            With SampleIndexRegistry(mvarFileID)
                .measurementSteps.CurrentStepIndex = 1
                frmMeasure.SetFields .avgSteps, .curDemagLong, .doUp, .doBoth, mvarFileID
            End With
            frmMeasure.ZOrder
            frmMeasure.Show
            frmDCMotors.ChangerHole
            Changer_ProcessSample hole
        Case commandInitUp
            ' no longer serves a purpose except as marker for preprocessing
        Case commandHolder
            frmVacuum.MotorPower True
            'If MeasurementsSinceHolder = 0 Then Exit Sub
            MeasurementsSinceHolder = 0
            
''========================================================================================================
'            '(March 10, 2011 - I Hilburn)
'            'This code has been commented out as it is being applied even when the
'            'user has not selected for the susceptibility measurements to be performed
'            'New code has been added in SampleCommand.Execute and MeasureTreatAndRead in
'            'modMeasure to ensure that the susceptibility lagTime is set during the appropriate
'            'Holder measurements
''--------------------------------------------------------------------------------------------------------
'            If COMPortSusceptibility > 0 And EnableSusceptibility Then frmSusceptibilityMeter.LagTime
''========================================================================================================

            If Changer_ValidStart(hole) Then
                frmDCMotors.ChangerMotortoHole hole
                modConfig.SampleHandlerCurrentHole = hole
            Else
                Changer_NearestHole
            End If
            frmMeasure.SetFields 1, vbNullString, True, False, vbNullString
            frmMeasure.SetSample "Holder"
            frmMeasure.ZOrder
            frmMeasure.Show
            ' Process the current position (a hole)
            Changer_ProcessSample frmDCMotors.ChangerHole
        Case commandFlip
            If frmProgram.BackColor = &HFF00& Then
            If UseXYTableAPS Then
                frmDCMotors.MoveToCorner
                MeasurementsSinceHolder = 1000
            Else
                Changer_NearestHole
            End If
            With SampleIndexRegistry(mvarFileID)
                SetCodeLevel CodeOrange
                frmSendMail.MailNotification "Flip samples", "Please place the samples '" & .SampleCode & "' in the sample tray " & _
                    "with the arrows pointed down.", CodeOrange
            '    MsgBox "Please place the samples '" & .SampleCode & "' in the sample tray " & _
            '        "with the arrows pointed down."
                SetCodeLevel CodeGreen, True
                .doUp = Not .doUp
            End With
            Else
            frmVacuum.ValveConnect False
            frmVacuum.MotorPower False
            If UseXYTableAPS Then
                frmDCMotors.MoveToCorner
            Else
                Changer_NearestHole
            End If
            With SampleIndexRegistry(mvarFileID)
                SetCodeLevel CodeOrange
                frmSendMail.MailNotification "Flip samples", "Please place the samples '" & .SampleCode & "' in the sample tray " & _
                    "with the arrows pointed down.", CodeOrange
                MsgBox "Please place the samples '" & .SampleCode & "' in the sample tray " & _
                    "with the arrows pointed down."
                SetCodeLevel CodeGreen, True
                .doUp = Not .doUp
            End With
            frmVacuum.MotorPower True
            End If
        Case commandFin
            frmVacuum.ValveConnect False
            frmVacuum.MotorPower False
            frmMeasure.Hide
            If LenB(fileid) > 0 Then
                If SampleIndexRegistry(fileid).measurementSteps.NeedsCoilAFDemag Then
                    Changer_NearestHole
                    MotorUpDn_Move AFPos, 2   '  Move somewhat slowly into AF region
                    frmDCMotors.TurningMotorRotate 0
                    
'-----------------------------------------------------------------------------------------------------------------------'
'-----------------------------------------------------------------------------------------------------------------------'
'
'       New AF Code Mod
'       July 2010
'       Isaac Hilburn
'
'   Details:
'       Added in if ... then statement to take into account which AF system that the user has set.
'-----------------------------------------------------------------------------------------------------------------------'
                    
                    If AFSystem = "2G" Then
                        
                        frmAF_2G.Connect
                        frmAF_2G.ExecuteRamp "C", _
                                             TransverseCoilSystem, _
                                             AfTransMax, _
                                             AFDelay, _
                                             AFRampRate
                        
                    ElseIf AFSystem = "ADWIN" Then
                    
                        frmADWIN_AF.ExecuteRamp TransverseCoilSystem, _
                                                AfTransMax, _
                                                , , , 0, _
                                                True, False, _
                                                (frmADWIN_AF.chkVerbose.Value = Checked)
                                                
                    End If
                        
                    frmDCMotors.TurningMotorRotate 90
                    
                    If AFSystem = "2G" Then
                        
                        If Not frmAF_2G.txtWaitingTime = 0 Then DelayTime (frmAF_2G.txtWaitingTime)
                        
                        frmAF_2G.ExecuteRamp "C", _
                                             TransverseCoilSystem, _
                                             AfTransMax, _
                                             AFDelay, _
                                             AFRampRate
                        
                    ElseIf AFSystem = "ADWIN" Then
                    
                        frmADWIN_AF.ExecuteRamp TransverseCoilSystem, _
                                                AfTransMax, _
                                                , , , 0, _
                                                True, False, _
                                                (frmADWIN_AF.chkVerbose.Value = Checked)
                    
                    End If
                                                
                    frmDCMotors.TurningMotorRotate 360
                    
                    If AFSystem = "2G" Then
                    
                        If Not frmAF_2G.txtWaitingTime = 0 Then DelayTime (frmAF_2G.txtWaitingTime)
                        
                        frmAF_2G.ExecuteRamp "C", _
                                             AxialCoilSystem, _
                                             AfAxialMax, _
                                             AFDelay, _
                                             AFRampRate
                                             
                        frmAF_2G.Disconnect
                        
                    ElseIf AFSystem = "ADWIN" Then
                    
                        frmADWIN_AF.ExecuteRamp AxialCoilSystem, _
                                                AfAxialMax, _
                                                , , , 0, _
                                                True, False, _
                                                (frmADWIN_AF.chkVerbose.Value = Checked)
                    
                    End If
                    
                    MotorUpDn_Move 0, 2
                    
                End If
            End If
            ' reset SampleHolder step type to NRM, just in case
            SampleHolder.Parent.measurementSteps(1).StepType = "NRM"
            SampleHolder.Parent.measurementSteps(1).Level = 0
            
            If hole = 0 Then Changer_NearestHole
            
            NotificationMessage = "Measurements are done on " & SampleIndexRegistry(mvarFileID).SampleCode & "."
            frmSendMail.MailNotification "2G Status Update", NotificationMessage
            
            
            '(July 2011 - I Hilburn)
            'Added this in to prompt the user to turn off the air and
            'to turn off the power to the coil thermal sensors for a rockmag run
            If LenB(fileid) Then
                If SampleIndexRegistry(fileid).measurementSteps.NeedsCoilAFDemag Then
                    If EnableT1 Or EnableT2 Then
                    
            End If
            'automatically turn off the air
            If modConfig.DoDegausserCooling = True Then
            frmVacuum.DegausserCooler False
                        'Prompt user to turn off the air and the temperature sensor power
                        MsgBox "Please: " & vbNewLine & vbNewLine & _
                               " - Verify the air is off" & vbNewLine & _
                               " - Switch off the power to the Rockmag coil temperature sensors"
                    Else
                    MsgBox "Please: " & vbNewLine & vbNewLine & _
                               " - Turn off the air" & vbNewLine & _
                               " - Switch off the power to the Rockmag coil temperature sensors"

            End If
                                'automatically turn off the air
            If modConfig.DoDegausserCooling = True Then
            frmVacuum.DegausserCooler False
            MsgBox "Please verify the air is off."
            Else
                        'Prompt user to turn off the air
                        MsgBox "Please turn off the air."
                    End If
                End If
            End If
            
            'MsgBox (NotificationMessage)
        Case commandGoto
            
            If Usextableaps And hole = -1 Then
                frmDCMotors.MoveToCorner pauseOveride:=True
            
            ElseIf Changer_ValidStart(hole) Then
                modConfig.SampleHandlerCurrentHole = hole
                frmDCMotors.ChangerMotortoHole hole
                frmDCMotors.ChangerHole
            End If
    End Select
End Sub

