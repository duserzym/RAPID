Attribute VB_Name = "modStatusCode"
Public StatusCodeColorLevel  As String
Public StatusCodeColorLevelPrior As String

Global Const CodeRed = "Red"              'EMERGENCY!
Global Const CodeOrange = "Orange"          'Attention required
Global Const CodeYellow = "Yellow"          'Oops!
Global Const CodeGreen = "Green"            'Everything running
Global Const CodeBlue = "Blue"              'Magnetometer Free
Global Const CodeGrey = "Grey"              'No idea

Global Const ColorGrey = &H8000000C
Global Const ColorBlue = &HFFFF00
Global Const ColorGreen = &HFF00&
Global Const ColorYellow = &HFFFF&
Global Const ColorOrange = &H80FF&
Global Const ColorRed = &HFF&

Public Sub SetCodeLevel(newLevel As String, _
                        Optional withEmailUpdate = False)
    If newLevel = CodeRed Or _
       newLevel = CodeOrange Or _
       newLevel = CodeYellow Or _
       newLevel = CodeGreen Or _
       newLevel = CodeBlue Or _
       newLevel = CodeGrey _
    Then
    
        If StatusCodeColorLevel = CodeGreen Or _
           StatusCodeColorLevel = CodeBlue _
        Then
        
            StatusCodeColorLevelPrior = StatusCodeColorLevel
            
        ElseIf LenB(statuscolorlevelprior) = 0 Then
        
            StatusCodeColorLevelPrior = CodeBlue
            
        End If
        
        StatusCodeColorLevel = newLevel
        
        frmProgram.SetProgramCodeLevel newLevel
        
        If withEmailUpdate Then frmSendMail.StatusMonitorUpdate newLevel
        
    End If
End Sub

