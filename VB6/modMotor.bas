Attribute VB_Name = "modMotor"
' Motor Controller Driver
'
' This is the controlling software driver for the
' SilverMax Quicksilver DC servo motors. Modified from the
' CY550 Stepper Motor Controllers in the old Caltech system
'
' The only limit is that multiple motors can not be
' driven at the same time (obviously).

Option Explicit

Global Const MotorChanger As Integer = 1
Global Const MotorTurning As Integer = 2
Global Const MotorUpDown As Integer = 3
Global Const MotorChangerY As Integer = 4


Dim steps As Long

Public Sub MotorTurn_180()
    frmDCMotors.TurningMotorRotate 180
    'frmdcmotors.MotorSwitch "Off"
End Sub

Public Sub MotorTurn_270()
    frmDCMotors.TurningMotorRotate 270
    'frmdcmotors.MotorSwitch "Off"
End Sub

Public Sub MotorTurn_360()
    frmDCMotors.TurningMotorRotate 360
    frmDCMotors.SetTurningMotorAngle 0
    'frmdcmotors.MotorSwitch "Off"
End Sub

Public Sub MotorTurn_90()
    frmDCMotors.TurningMotorRotate 90
    'frmdcmotors.MotorSwitch "Off"
End Sub

'
'
Public Function MotorUPDN_home() As Long
    ' now actually move the motor
    frmDCMotors.SamplePickup
    
    MotorUPDN_home = frmDCMotors.UpDownHeight
    '
    ' now disconnect the motor
    'frmdcmotors.MotorSwitch "Off"
End Function

'  The following routines are the basic calls to James' routines for controlling the
'  SilverMax (QuickSilver) DC servo motors used on the new Caltech 200 sample changer
'

Public Sub MotorUpDn_Move(Position As Long, speed As Integer)
    frmDCMotors.UpDownMove Position, speed
End Sub

Public Sub MotorUPDN_TopReset()
 
    frmDCMotors.HomeToTop

End Sub

Public Sub MotorXYTable_CenterReset()

frmProgram.StatBarNew "Homing XY Table..."

Dim user_resp As VbMsgBoxResult
user_resp = frmDCMotors.PromptUser_DoHomeXYStage
        
If user_resp <> vbYes Then Exit Sub

Dim xPos As Long
Dim yPos As Long

'Load and show the shutdown msgform
    Load frmShutdownMsg
    frmXYHoming.ZOrder 0
    frmXYHoming.Show
    
    frmDCMotors.HomeToCenter xPos, yPos, pauseOveride:=True
    modConfig.XYTablePositions(0, 0) = xPos
    modConfig.XYTablePositions(0, 1) = yPos

'Hide and unload the shutdown msg form
    frmXYHoming.Hide
    Unload frmXYHoming
    
    frmProgram.StatBarNew ""
    
'
' This routine should do the soft home to the top of the run and set the position counter
' to zero.
'
'
End Sub

