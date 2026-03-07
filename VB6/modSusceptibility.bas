Attribute VB_Name = "modSusceptibility"
Option Explicit

Public Function Susceptibility_Measure(processingSample As Sample, Optional ByVal IsHolder As Boolean = False) As Double
    Dim SCoilSampleCenterPos As Long
    Dim measured As Double
    
    Susceptibility_Measure = 0
    If COMPortSusceptibility < 1 Then Exit Function
    
    SCoilSampleCenterPos = Int(SCoilPos + processingSample.SampleHeight / 2)
    If SCoilSampleCenterPos / Abs(SCoilSampleCenterPos) <> SCoilPos / Abs(SCoilPos) Then
        ' crap... our sample is too large to put in the susceptibility coil!
        Exit Function
    End If
    
    If Abs(frmDCMotors.UpDownHeight) > 0.5 * Abs(SampleBottom) Then frmDCMotors.HomeToTop
    
    frmSusceptibilityMeter.Zero
    MotorUpDn_Move SCoilSampleCenterPos, 0 ' (December 2008 K Bradley) Slow down the rod when it goes to the X coil to prevent breaking the rod
    measured = frmSusceptibilityMeter.Measure
    
    If IsHolder Then Susceptibility_Measure = measured Else Susceptibility_Measure = Susceptibility_Standardize(measured)
    processingSample.Susceptibility = Susceptibility_Measure
End Function

Public Function Susceptibility_Standardize(uncalibrated As Double) As Double
    
    Susceptibility_Standardize = (uncalibrated - SampleHolder.Susceptibility) * SusceptibilityMomentFactorCGS
    
End Function

