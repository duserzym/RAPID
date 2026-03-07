Attribute VB_Name = "modFlow"
' Handles pausing of motors

Option Explicit

Private old_NOCOMMMODE      As Boolean
Public Prog_paused          As Boolean
Public Prog_halted          As Boolean

Public Sub Flow_Halt()
    Dim i As Integer
    frmDCMotors.MotorStop
    Prog_paused = True
    Prog_halted = True
    old_NOCOMMMODE = NOCOMM_MODE
    'NOCOMM_MODE = True
    frmProgram.updateFlowMenu
    frmMeasure.updateFlowStatus
    frmSettings.cmdFlowControl.Caption = "Resume Flow"      '(July 2010 - I Hilburn) Added to update new flow button on
                                                    'settings form
    frmDCMotors.lastMoveCommand = 0
End Sub

Public Sub Flow_Pause()
    Dim i As Integer
    If Not Prog_halted Then
        frmDCMotors.MotorStop
        Prog_halted = False
        Prog_paused = True
        old_NOCOMMMODE = NOCOMM_MODE
        'NOCOMM_MODE = True
        frmMeasure.updateFlowStatus
        frmProgram.updateFlowMenu
        frmSettings.cmdFlowControl.Caption = "Resume Flow"  '(July 2010 - I Hilburn) Added to update new flow button on
                                                    'settings form
    End If
End Sub

Public Sub Flow_Resume()
    Prog_paused = False
    Prog_halted = False
    NOCOMM_MODE = old_NOCOMMMODE
    If Prog_paused Or Prog_halted Then frmDCMotors.ResumeMove
    frmProgram.updateFlowMenu
    frmMeasure.updateFlowStatus
    frmSettings.cmdFlowControl.Caption = "Pause Flow"   '(July 2010 - I Hilburn) Added to update new flow button on
                                                'settings form
End Sub

Public Sub Flow_WaitForUnpaused()
    Do Until Not Prog_paused
        
        DoEvents    '(July 2010 - I Hilburn) Added so this function could be used as the default for
                    'waiting until the user un-pauses the code
    
        DelayTime 0.05
    Loop
End Sub

