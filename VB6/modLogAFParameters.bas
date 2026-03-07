Attribute VB_Name = "modLogAFParameters"
Option Explicit

Private af_log_folder_path As String
Private af_log_file_name As String

Private m_line_number As Long
Private Const max_allowed_number_of_lines_in_log_file As Long = 10000

Private m_fso As FileSystemObject
Private m_text_stream As TextStream

Private log_in_progress As Boolean

Private Function CheckFileSystemObjectStatus() As Boolean

    On Error GoTo CheckFileSystemObjectStatus_Error

    If m_fso Is Nothing Then
    
        Set m_fso = New FileSystemObject
        
    End If
    
    If m_fso Is Nothing Then
        CheckFileSystemObjectStatus = False
    Else
        CheckFileSystemObjectStatus = True
    End If
    
    On Error GoTo 0
    Exit Function
    
CheckFileSystemObjectStatus_Error:

    CheckFileSystemObjectStatus = False

End Function

Private Function CheckLogFolder() As Boolean

    If af_log_folder_path = "" Then
    
        af_log_folder_path = GetDefaultAfLogFolderPath
        
    End If

    If Not m_fso.FolderExists(af_log_folder_path) Then
        
        m_fso.CreateFolder (af_log_folder_path)
        
    End If
    
    CheckLogFolder = m_fso.FolderExists(af_log_folder_path)

End Function

Private Function CheckLogFileName() As Boolean

    If af_log_file_name = "" Then
    
        af_log_file_name = GetDefaultAfLogFileName
        
    End If
    
    CheckLogFileName = True
    
End Function

Private Sub AppendXMLHeaderToLogFile()

    If Not CheckTextStream Then Exit Sub
    If Not m_line_number <= 1 Then Exit Sub
    
    On Error Resume Next
    m_text_stream.Write GetAfLogFileHeaderXMLString(Now) & vbCrLf
    On Error GoTo 0

End Sub

Private Sub OpenLogFileForAppend()
    
    On Error GoTo OpenLogFileForAppend_Error
     
    If Not CheckFileSystemObjectStatus Then Exit Sub
    If Not CheckLogFolder Then Exit Sub
    If Not CheckLogFileName Then Exit Sub
    
    Dim file_path As String
    
    file_path = af_log_folder_path & af_log_file_name
    
    CloseTextStream
    Set m_text_stream = m_fso.OpenTextFile(file_path, ForAppending, True)
    m_line_number = m_text_stream.line
    
    If m_line_number <= 1 Then AppendXMLHeaderToLogFile
        
    On Error GoTo 0
    Exit Sub
    
OpenLogFileForAppend_Error:

    Set m_text_stream = Nothing
    Set m_fso = Nothing
    
End Sub

Private Function CheckTextStream() As Boolean

    On Error GoTo CheckTextStream_Error

    If m_text_stream Is Nothing Then
    
        CheckTextStream = False
        Exit Function
        
    End If
    
    Dim line_position As Long
    
    line_position = m_text_stream.line
    
    If line_position >= 0 Then
        
        CheckTextStream = True
        
    Else
    
        CheckTextStream = False
        
    End If
    
    On Error GoTo 0
    Exit Function
    
CheckTextStream_Error:

    CheckTextStream = False

End Function

Private Sub CloseTextStream()

    If Not m_text_stream Is Nothing Then
    
        On Error Resume Next
        m_text_stream.Close
        On Error GoTo 0
        
        Set m_text_stream = Nothing
        
    End If

End Sub

Public Sub CloseLogFile()
    If CheckTextStream Then
    
        m_text_stream.WriteLine "</logs>"

    End If

    CloseTextStream
    Set m_fso = Nothing
End Sub

Private Function SanitizeXMLValueString(ByVal value As String) As String

    Dim ret_val As String
    
    ret_val = Replace(value, "&", "&amp;")
    ret_val = Replace(ret_val, "<", "&lt;")
    ret_val = Replace(ret_val, ">", "&gt;")
    
    SanitizeXMLValueString = ret_val
    
End Function

Private Function SanitizeXMLAttributeString(ByVal value As String) As String

    Dim ret_val As String
    
    ret_val = Replace(value, "&", "&amp;")
    ret_val = Replace(ret_val, "<", "&lt;")
    ret_val = Replace(ret_val, ">", "&gt;")
    ret_val = Replace(ret_val, """", "&quot;")
    ret_val = Replace(ret_val, "'", "&apos;")
    
    SanitizeXMLAttributeString = ret_val
    
End Function

Private Function GetAfLogFileHeaderXMLString(ByVal date_time_stamp As Date) As String

    Dim ret_val As String
    
    
    ret_val = "<?xml version=""1.0""?>" & _
              vbCrLf & _
              "<logs file_created=""" & _
              Format$(date_time_stamp, "yyyy/mm/dd, hh:mm:ss") & """>"

              
    ret_val = ret_val & vbCrLf & _
              "<mail_from_name>" & _
              SanitizeXMLValueString(modConfig.MailFromName) & _
              "</mail_from_name>"
              
    ret_val = ret_val & vbCrLf & _
              "<log_file_path>" & _
              SanitizeXMLValueString(af_log_folder_path) & _
              "</log_file_path>"
              
    ret_val = ret_val & vbCrLf & _
              "<log_file_name>" & _
              SanitizeXMLValueString(af_log_file_name) & _
              "</log_file_name>"

    GetAfLogFileHeaderXMLString = ret_val

End Function

Private Function ParseAfOutputParametersToXMLString(ByRef outputs As AdwinAfOutputParameters) As String

    ParseAfOutputParametersToXMLString = ""
                                             
        On Error GoTo ParseAfOutputParametersToXMLString_Error
                                             
        If outputs Is Nothing Then Exit Function
        
        Dim ret_val As String
        
        With outputs
            
            ret_val = "<outputs coil=""" & _
                      SanitizeXMLAttributeString(Trim(outputs.Coil)) & """>"
                        
            ret_val = ret_val & _
                      ParseAdwinAfParameterToXMLTagString(.Measured_Peak_Monitor_Voltage)
                      
            ret_val = ret_val & _
                      ParseAdwinAfParameterToXMLTagString(.Max_Ramp_Voltage_Used)
                      
            ret_val = ret_val & _
                      ParseAdwinAfParameterToXMLTagString(.Time_Step_Between_Points)
                      
            ret_val = ret_val & _
                      ParseAdwinAfParameterToXMLTagString(.Number_Points_Per_Period)
                      
            ret_val = ret_val & _
                      ParseAdwinAfParameterToXMLTagString(.Total_Output_Points)
                      
            ret_val = ret_val & _
                      ParseAdwinAfParameterToXMLTagString(.Total_Monitor_Points)
            
            ret_val = ret_val & _
                      ParseAdwinAfParameterToXMLTagString(.Ramp_Up_Last_Point)
            
            ret_val = ret_val & _
                      ParseAdwinAfParameterToXMLTagString(.Ramp_Down_First_Point)
            
            ret_val = ret_val & _
                      ParseAdwinAfParameterToXMLTagString(.Actual_Slope_Down_Used)
                                  
            ret_val = ret_val & _
                      "</outputs>"
            
        End With
            
        ParseAfOutputParametersToXMLString = ret_val
                      
        On Error GoTo 0
        Exit Function
        
ParseAfOutputParametersToXMLString_Error:
                                    

End Function

Private Function ParseAfInputParametersToXMLString(ByRef inputs As AdwinAfInputParameters) As String
        
        ParseAfInputParametersToXMLString = ""
                                             
        On Error GoTo ParseAfInputParametersToXMLString_Error
                                             
        If inputs Is Nothing Then Exit Function
        
        Dim ret_val As String
        
        With inputs
            
            ret_val = "<inputs coil=""" & _
                      SanitizeXMLAttributeString(Trim(inputs.Coil)) & """>"
            
            ret_val = ret_val & _
                      ParseAdwinAfParameterToXMLTagString(.Slope_Up)
                      
            ret_val = ret_val & _
                      ParseAdwinAfParameterToXMLTagString(.Slope_Down)
                      
            ret_val = ret_val & _
                      ParseAdwinAfParameterToXMLTagString(.Peak_Monitor_Voltage)
                      
            ret_val = ret_val & _
                      ParseAdwinAfParameterToXMLTagString(.Resonance_Freq)
                      
            ret_val = ret_val & _
                      ParseAdwinAfParameterToXMLTagString(.Peak_Ramp_Voltage)
                      
            ret_val = ret_val & _
                      ParseAdwinAfParameterToXMLTagString(.Max_Monitor_Voltage)
                      
            ret_val = ret_val & _
                      ParseAdwinAfParameterToXMLTagString(.Max_Ramp_Voltage)
                      
            ret_val = ret_val & _
                      "<RampMode value=""" & _
                      .ramp_mode.ToString & """>" & _
                      SanitizeXMLValueString(.GetRampModeStringDescrip()) & _
                      "</RampMode>"
                      
            ret_val = ret_val & _
                      ParseAdwinAfParameterToXMLTagString(.Output_Port_Number)
                      
            ret_val = ret_val & _
                      ParseAdwinAfParameterToXMLTagString(.Monitor_Port_Number)
                      
            ret_val = ret_val & _
                      ParseAdwinAfParameterToXMLTagString(.Process_Delay)
                      
            ret_val = ret_val & _
                      ParseAdwinAfParameterToXMLTagString(.Noise_Level)
                      
            ret_val = ret_val & _
                      ParseAdwinAfParameterToXMLTagString(.Number_Periods_Hang_At_Peak)
                      
            ret_val = ret_val & _
                      ParseAdwinAfParameterToXMLTagString(.Number_Periods_Ramp_Down)
                      
            ret_val = ret_val & _
                      "<RampDownMode value=""" & _
                      .ramp_down_mode.ToString & """>" & _
                      SanitizeXMLValueString(.GetRampDownModeStringDescrip()) & _
                      "</RampDownMode>"
                      
            ret_val = ret_val & _
                      "</inputs>"
            
        End With
            
        ParseAfInputParametersToXMLString = ret_val
                      
        On Error GoTo 0
        Exit Function
        
ParseAfInputParametersToXMLString_Error:
                                             
End Function

Private Function ParseAdwinAfParameterToXMLTagString(ByVal parameter As AdwinAfParameter) As String

    ParseAdwinAfParameterToXMLTagString = ""

    On Error GoTo ParseAdwinAfParameterToXMLTagString_Error

    If parameter Is Nothing Then Exit Function
    
    Dim ret_val As String
    
    With parameter
    
        ret_val = "<" & .ParamName & " type=""" & _
                    SanitizeXMLAttributeString(.GetTypeString()) & """" & _
                    " number=""" & _
                    SanitizeXMLAttributeString(Trim(CStr(.ParamNumber))) & """>" & _
                    SanitizeXMLValueString(.ToString()) & _
                    "</" & .ParamName & ">"
                    
    End With
    
    ParseAdwinAfParameterToXMLTagString = ret_val
        
    On Error GoTo 0
    Exit Function
    
ParseAdwinAfParameterToXMLTagString_Error:
    
End Function

Public Sub LogAFRamp(ByVal inputs As AdwinAfInputParameters, _
                     ByVal outputs As AdwinAfOutputParameters, _
                     ByVal status As AdwinAfRampStatus)

    If log_in_progress Then Exit Sub
    
    log_in_progress = False

    If inputs Is Nothing Then Exit Sub
    If outputs Is Nothing Then Exit Sub
    If status Is Nothing Then Exit Sub

    log_in_progress = True

    On Error GoTo LogAFRamp_Error
        
        Dim msg As String
        
        msg = GetAfRampTagOpeningString(status, inputs, outputs)
        msg = msg & vbCrLf
        msg = msg & _
              ParseAfInputParametersToXMLString(inputs)
        msg = msg & vbCrLf
        msg = msg & _
              ParseAfOutputParametersToXMLString(outputs)
        msg = msg & vbCrLf
        msg = msg & GetAfRampTagClosingString()
        
        AppendStringToAfLogFile msg
        
        log_in_progress = False
        
    On Error GoTo 0
    Exit Sub
    
LogAFRamp_Error:

    log_in_progress = False

End Sub

Private Function GetAfRampTagClosingString() As String

    GetAfRampTagClosingString = "</af_ramp_log>" & vbCrLf

End Function

Private Function GetAfRampTagOpeningString(ByRef status As AdwinAfRampStatus, _
                                           ByRef inputs As AdwinAfInputParameters, _
                                           ByRef outputs As AdwinAfOutputParameters) As String
                                           
    GetAfRampTagOpeningString = ""
                                           
    If status Is Nothing Then Exit Function
    If inputs Is Nothing Then Exit Function
    If outputs Is Nothing Then Exit Function
    
    Dim ret_val As String
    
    ret_val = ""
    
    On Error GoTo GetAfRampTagOpeningString_Error
    
    
    ret_val = "<af_ramp_log coil=""" & _
              SanitizeXMLAttributeString(Trim(status.Coil)) & """" & _
              " successful=""" & _
              Trim(CStr(status.WasSuccessful)) & """" & _
              " ramp_started=""" & _
              Format$(status.Ramp_Start_Time, "yyyy/mm/dd, hh:mm:ss") & """" & _
              " ramp_ended=""" & _
              Format$(status.Ramp_End_Time, "yyyy/mm/dd, hh:mm:ss") & """" & _
              " program_duration_in_secs=""" & _
              Trim(CStr(status.GetProgramRampDurationInSeconds())) & """>" & _
              vbCrLf
              
    ret_val = ret_val & _
              "<peak_values field=""" & _
              SanitizeXMLAttributeString(status.TargetPeakField) & """" & _
              " target_voltage=""" & _
              inputs.Peak_Monitor_Voltage.ToString & " V""" & _
              " actual_voltage=""" & _
              outputs.Measured_Peak_Monitor_Voltage.ToString & " V""/>" & _
              vbCrLf
              
    ret_val = ret_val & _
              "<duration units=""msecs"">" & _
              "<total>" & _
              outputs.SingleToString(outputs.GetTotalRampDuration) & _
              "</total>" & _
              "<ramp_up>" & _
              outputs.SingleToString(outputs.GetRampUpDuration) & _
              "</ramp_up>" & _
              "<at_peak>" & _
              outputs.SingleToString(outputs.GetPeakDuration) & _
              "</at_peak>" & _
              "<ramp_down>" & _
              outputs.SingleToString(outputs.GetRampDownDuration) & _
              "</ramp_down>" & _
              "</duration>"
              
    GetAfRampTagOpeningString = ret_val
    
    On Error GoTo 0
    Exit Function
    
GetAfRampTagOpeningString_Error:
                                           
End Function

Private Sub AppendStringToAfLogFile(ByVal msg As String)
    
    On Error GoTo AppendStringToAfLogFile_Error
    
    'Filter for empty message
    If Trim(msg) = "" Then Exit Sub
    
    If Not CheckTextStream Then
        OpenLogFileForAppend
        If Not CheckTextStream Then
            Exit Sub
        End If
    End If
    
    m_text_stream.Write Trim(msg)
    
    If LogFileFull Then
        
        CloseTextStream
        
        PauseTill_NoEvents timeGetTime() + 1000
        
        af_log_file_name = GetDefaultAfLogFileName
            
        OpenLogFileForAppend
            
    End If
    
    On Error GoTo 0
    Exit Sub
    
AppendStringToAfLogFile_Error:
    
End Sub

Private Function LogFileFull() As Boolean

    If m_line_number >= max_allowed_number_of_lines_in_log_file Then
    
        LogFileFull = True
        
    Else
    
        LogFileFull = False
        
    End If

End Function

Private Function GetDefaultAfLogFolderPath() As String

    Dim folder_path As String
    
    folder_path = modConfig.ADWIN_AFDataLocalDir
    
    If Strings.Right(folder_path, 1) <> "\" Then folder_path = folder_path & "\"
    
    folder_path = folder_path & "Ramp Status Logs\"
    
    GetDefaultAfLogFolderPath = folder_path

End Function

Private Function GetDefaultAfLogFileName() As String

    GetDefaultAfLogFileName = Format$(Now, "yyyy-mm-dd_hhmmss") & ".xml"
    
End Function

