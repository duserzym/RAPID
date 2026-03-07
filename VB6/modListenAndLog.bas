Attribute VB_Name = "modListenAndLog"
Option Explicit

Private m_line_number As Long
Private Const max_allowed_number_of_lines_in_log_file As Long = 10000

Public IsLogFileOpenForAppend As Boolean

Private m_fso As FileSystemObject
Private m_text_stream As TextStream

Private Sub PrependDateStampToString(ByRef str As String)

    str = Format$(Now, "yyyy/mm/dd, hh:mm:ss") & "," & Trim(str)

End Sub

Private Sub PostpendMailFromNameToString(ByRef str As String)

    str = str & "," & EncapsulateStringWithDoubleQuotationMarks(Trim(modConfig.MailFromName))

End Sub

Private Function AddLineNumberToString(ByVal str As String) As String

    AddLineNumberToString = Trim(CStr(m_line_number)) & "," & Trim(str)

End Function

Private Sub IncrementLineNumber()

    m_line_number = m_line_number + 1

End Sub

Private Function EncapsulateStringWithDoubleQuotationMarks(ByVal str As String) As String

    EncapsulateStringWithDoubleQuotationMarks = """" & str & """"

End Function

Public Function CheckFileSystemObjectStatus() As Boolean

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

Public Function CheckLogFolder() As Boolean

    If modConfig.LogFolderPath = "" Then
    
        modConfig.LogFolderPath = GetDefaultCommLogFolderPath
        
    End If

    If Not m_fso.FolderExists(modConfig.LogFolderPath) Then
        
        m_fso.CreateFolder (modConfig.LogFolderPath)
        
    End If
    
    CheckLogFolder = m_fso.FolderExists(modConfig.LogFolderPath)

End Function

Public Function CheckLogFileName() As Boolean

    If modConfig.LogFileName = "" Then
    
        modConfig.LogFileName = GetDefaultCommLogFileName
        
    End If
    
    CheckLogFileName = True
    
End Function

Public Sub OpenLogFileForAppend()
    
    On Error GoTo OpenLogFileForAppend_Error
     
    If Not CheckFileSystemObjectStatus Then Exit Sub
    If Not CheckLogFolder Then Exit Sub
    If Not CheckLogFileName Then Exit Sub
    
    Dim file_path As String
    
    file_path = modConfig.LogFolderPath & modConfig.LogFileName
    
    CloseTextStream
    Set m_text_stream = m_fso.OpenTextFile(file_path, ForAppending, True)
    m_line_number = m_text_stream.line
    
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
    CloseTextStream
    Set m_fso = Nothing
End Sub

Public Sub AppendRS232MessageToLogFile(ByVal msg As String, ByVal comm_port_num As Integer, ByVal input_or_output_string As String)
    
    On Error GoTo AppendMessageToLogFile_Error
    
    'Filter for empty message
    If Trim(msg) = "" Then Exit Sub
    
    Dim line_str As String
    
    line_str = "COM-" & Trim(CStr(comm_port_num)) & "," & _
               input_or_output_string & "," & _
               EncapsulateStringWithDoubleQuotationMarks(msg)
        
    PrependDateStampToString line_str
    PostpendMailFromNameToString line_str
    
    If Not CheckTextStream Then
        OpenLogFileForAppend
        If Not CheckTextStream Then
            Exit Sub
        End If
    End If
    
    IncrementLineNumber
    
    line_str = AddLineNumberToString(line_str)
    
    m_text_stream.WriteLine line_str
    
    If LogFileFull Then
        
        CloseTextStream
        
        PauseTill_NoEvents timeGetTime() + 1000
        
        modConfig.LogFileName = GetDefaultCommLogFileName
            
        OpenLogFileForAppend
            
    End If
    
    On Error GoTo 0
    Exit Sub
    
AppendMessageToLogFile_Error:
    
End Sub

Public Function LogFileFull() As Boolean

    If m_line_number >= max_allowed_number_of_lines_in_log_file Then
    
        LogFileFull = True
        
    Else
    
        LogFileFull = False
        
    End If

End Function

Public Function GetDefaultCommLogFolderPath() As String

    Dim folder_path As String
    
    folder_path = App.path
    
    If Strings.Right(folder_path, 1) <> "\" Then folder_path = folder_path & "\"
    
    folder_path = folder_path & "Comm Logs\"
    
    GetDefaultCommLogFolderPath = folder_path

End Function

Public Function GetDefaultCommLogFileName() As String

    GetDefaultCommLogFileName = Format$(Now, "yyyy-mm-dd_hhmmss") & ".csv"
    
End Function
