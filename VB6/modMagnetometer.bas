Attribute VB_Name = "modMagnetometer"
' The PaleoMag Magnetometer

Option Explicit

Global SampleOrientationCurrent As Integer
Global SampleNameCurrent As String
Global SampleStepCurrent As String

Sub Magnetometer_Initialize()
    
    frmProgram.StatBarNew "Configuring SQUID..."
    frmSQUID.Configure "A"

    SampleNameCurrent = vbNullString
    SampleStepCurrent = vbNullString
    SampleOrientationCurrent = 0

    ' Move vertical motor to top position
    frmProgram.StatBarNew "Homing vertical motor..."
    MotorUPDN_TopReset
    
    ' Move XY Table To Center Position
    If UseXYTableAPS Then
    
        HasXYTableBeenHomed = False
    
        MotorXYTable_CenterReset
        
    End If
    
    FLAG_MagnetUse = False        ' Notify that we stopped
    
    ' Initialize Vacuum.
    
    frmProgram.StatBarNew "Initializing vacuum..."
    
    'frmVacuum.MotorPower True ' Vacuum Motor On
    If DoVacuumReset Then frmVacuum.Reset
    
    ' if EnableAxialIRM, then discharge
    If EnableARM Then frmIRMARM.SetBiasField 0
    
    If EnableAxialIRM Then
        frmProgram.StatBarNew "Discharging IRM..."
        frmIRMARM.optCoil(0).Value = True
        frmIRMARM.FireIRM 0
    End If
    
    FLAG_MagnetInit = True        ' We're done initializing
    frmProgram.StatBarNew vbNullString
    
End Sub

Sub Magnetometer_UnloadSample()
    If (SampleNameCurrent <> "Holder") Then
        SampleNameCurrent = vbNullString
        SampleStepCurrent = vbNullString
    End If
End Sub

