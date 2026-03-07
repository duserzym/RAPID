Attribute VB_Name = "modAF_2G"
' This module handles the interfacing between the
' DC motors and the AF system.
Option Explicit

Public Sub AF_Demagnetize(ByVal AFLevel As Double, Optional ByVal HoldTime As Double = 0)
    
    Dim SampleCenterAFPosition As Long
    
    If AFLevel <= 0 Then Exit Sub
    
    'Calculate and set the up/down motor pos to place the center of the sample into
    'the center of the AF coils
    SampleCenterAFPosition = Int(AFPos + SampleHeight / 2)
    
    If SampleCenterAFPosition / Abs(SampleCenterAFPosition) <> AFPos / Abs(AFPos) Then
        
        ' crap... our sample is too large to put in the AF coil!
        Exit Sub
    
    End If
    
    'Check to see if the AF module is enabled
    If EnableAF = False Then
    
        'Inform the user that the AF Module is off
        MsgBox "AF module is currently disabled.  Cannot perform desired AF demag at this time." & _
               vbNewLine & vbNewLine & "Sorry for the inconvenience.", , _
               "Whoops!"
    
        Exit Sub
        
    End If
    
    '  Move somewhat slowly into AF region
    MotorUpDn_Move SampleCenterAFPosition, 2
    
    frmDCMotors.TurningMotorRotate 0
    
'------------------------------------------------------------------------------------------------------------------'
'------------------------------------------------------------------------------------------------------------------'
'
'   Code Mod
'   8/11/2010
'   Isaac Hilburn
'
'   Summary:    Add in if ... then ... elseif ... then statements to select the correct AF system to use
'               Using the frmADWIN_AF.ExecuteRamp function to setup and run the AF ramp and write the data to
'               file if desired by the user
'
'------------------------------------------------------------------------------------------------------------------'
'------------------------------------------------------------------------------------------------------------------'
    
    If AFSystem = "2G" Then
        
        frmAF_2G.Connect
        
        frmAF_2G.CycleWithHold HoldTime, _
                               TransverseCoilSystem, _
                               AFLevel, _
                               AFRampRate
                               
    ElseIf AFSystem = "ADWIN" Then
    
        frmADWIN_AF.ExecuteRamp TransverseCoilSystem, _
                                AFLevel, _
                                , , , , _
                                True, _
                                False, _
                                (frmADWIN_AF.chkVerbose.Value = Checked)
                                
    End If
                                
                               
    frmDCMotors.TurningMotorRotate 90
    
    If AFSystem = "2G" Then
        
        ' (August 2007 L Carporzen) Allow to wait between each ramp
        If Not frmAF_2G.txtWaitingTime = 0 Then DelayTime (frmAF_2G.txtWaitingTime)
        
        frmAF_2G.CycleWithHold HoldTime, _
                               TransverseCoilSystem, _
                               AFLevel, _
                               AFRampRate
                               
    ElseIf AFSystem = "ADWIN" Then
    
        frmADWIN_AF.ExecuteRamp TransverseCoilSystem, _
                                AFLevel, _
                                , , , , _
                                True, _
                                False, _
                                (frmADWIN_AF.chkVerbose.Value = Checked)
                                
    End If
    frmDCMotors.TurningMotorRotate 360
    
    If AFSystem = "2G" Then
        
        ' (August 2007 L Carporzen) Allow to wait between each ramp
        If Not frmAF_2G.txtWaitingTime = 0 Then DelayTime (frmAF_2G.txtWaitingTime)
        
        frmAF_2G.CycleWithHold HoldTime, _
                               AxialCoilSystem, _
                               AFLevel, _
                               AFRampRate
                               
    ElseIf AFSystem = "ADWIN" Then
    
        frmADWIN_AF.ExecuteRamp AxialCoilSystem, _
                                AFLevel, _
                                , , , , _
                                True, _
                                False, _
                                (frmADWIN_AF.chkVerbose.Value = Checked)
                                
    End If
    frmAF_2G.Disconnect
End Sub

