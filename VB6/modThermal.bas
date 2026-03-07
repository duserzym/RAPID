Attribute VB_Name = "modThermal"
Option Explicit

Public Sub NotifySensorError(ByVal Temp1 As Double, _
                             ByVal Temp2 As Double)

    Dim ErrorMessage As String
    Dim UserResp As Long

    'Create Error message
    ErrorMessage = "AF Temperature sensors reading temperature below minimum limit." & vbNewLine & _
                   vbNewLine & "Sensor #1 Temp: " & Trim(Str(Temp1)) & " " & modConfig.Tunits & _
                   vbNewLine & "Sensor #2 Temp: " & Trim(Str(Temp2)) & " " & modConfig.Tunits & _
                   vbNewLine & "Execution has been paused. Please come in and check the temperature " & _
                   "sensor box.  The two 9V batteries may need to be replaced, or the switch may need to be turned on."

    Flow_Pause
                   
    'Send out email and put screen into code red
    frmSendMail.MailNotification "AF Temperature Sensor Error", ErrorMessage, CodeRed, True
    
    frmProgram.SetProgramCodeLevel CodeRed
    
    UserResp = MsgBox(ErrorMessage & vbNewLine & vbNewLine & _
                      "Would you like to resume the code execution? Choosing 'No' will leave the code paused.", _
                      vbYesNo, _
                      "AF Thermal Sensor Error")
                      
    If UserResp = vbYes Then
        
        Flow_Resume
        
    End If
    
    frmProgram.SetProgramCodeLevel StatusCodeColorLevelPrior

End Sub

Public Function ValidSensorTemp(ByVal Temp1 As Double, _
                                ByVal Temp2 As Double) As Boolean
                                
    Dim NoErrorT1 As Boolean
    Dim NoErrorT2 As Boolean
    
    'Default as having no NoError
    NoErrorT1 = True
    NoErrorT2 = True
                                
    'Test to see if Temp1 or Temp2 are within 20 degrees of the arithmetic
    'inverse of the their offset temperature setting
    If Abs(modConfig.Toffset + Temp1) < 20 Then NoErrorT1 = False
    If Abs(modConfig.Toffset + Temp2) < 20 Then NoErrorT2 = False
        
    ValidSensorTemp = NoErrorT1 And NoErrorT2
                                
End Function

