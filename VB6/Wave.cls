VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Wave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"range"
Attribute VB_Ext_KEY = "Member1" ,"Range"
Attribute VB_Ext_KEY = "Member2" ,"Range"
Attribute VB_Ext_KEY = "Member3" ,"Range"
'local variable(s) to hold property value(s)
Private mvarIO As String 'local copy
Private mvarMemBuffer As Long 'local copy
Private mvarBufferAlloc As Boolean 'local copy
Private mvarDoDeallocate As Boolean
Private mvarNumPoints As Long 'local copy
Private mvarStartPoint As Long 'local copy
Private mvarCurrentPoint As Long
Private mvarIORate As Long 'local copy
Private mvarSineFreqMin As Double 'local copy
Private mvarSineFreqMax As Double
Private mvarDuration As Long 'local copy
Private mvarIOOptions As Long 'local copy
Private mvarPeakVoltage As Double 'local copy
Private mvarCurrentVoltage As Double
Private mvarMinVoltage As Double 'local copy
Private mvarSlope As Double
Private mvarBoardUsed As Board 'local copy
Private mvarPtsPerPeriod As Double 'local copy
Private mvarRange As range
Private mvarChan As Channel
Private mvarTimeStep As Double 'local copy
Private mvarWaveININum As Long
Private mvarWaveName As String
Private mvarWaveDesc As String

Public Property Set BoardUsed(ByVal vData As Board)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.BoardUsed = Form1

    Set mvarBoardUsed = vData
    
End Property

Public Property Get BoardUsed() As Board
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.BoardUsed

        Set BoardUsed = mvarBoardUsed

End Property

Public Property Let BufferAlloc(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.BufferAlloc = 5
    mvarBufferAlloc = vData
End Property

Public Property Get BufferAlloc() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.BufferAlloc
    BufferAlloc = mvarBufferAlloc
End Property

Public Property Set Chan(ByVal vData As Channel)
Attribute Chan.VB_UserMemId = 0
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Chan = Form1

    If vData Is Nothing Then
    
        Set vData = New Channel
        
    End If
    
    Set mvarChan = vData
End Property

Public Property Get Chan() As Channel
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Chan

       Set Chan = mvarChan
        
End Property

Public Property Let CurrentPoint(ByVal vData As Long)

    mvarCurrentPoint = vData

End Property

Public Property Get CurrentPoint() As Long

    CurrentPoint = mvarCurrentPoint

End Property

Public Property Let CurrentVoltage(ByVal vData As Double)

    mvarCurrentVoltage = vData

End Property

Public Property Get CurrentVoltage() As Double

    CurrentVoltage = mvarCurrentVoltage

End Property

Public Property Let DoDeallocate(ByVal vData As Boolean)

    mvarDoDeallocate = vData
    
End Property

Public Property Get DoDeallocate() As Boolean

    DoDeallocate = mvarDoDeallocate
    
End Property

Public Property Let Duration(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Duration = 5
    mvarDuration = vData
End Property

Public Property Get Duration() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Duration
    Duration = mvarDuration
End Property

Public Property Let IO(ByVal sData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.IO = 5
    mvarIO = sData
End Property

Public Property Get IO() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.IO
    IO = mvarIO
End Property

Public Property Let IOOptions(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.IOOptions = 5
    mvarIOOptions = vData
End Property

Public Property Get IOOptions() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.IOOptions
    IOOptions = mvarIOOptions
End Property

Public Property Let IORate(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.IORate = 5
    mvarIORate = vData
End Property

Public Property Get IORate() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.IORate
    IORate = mvarIORate
End Property

Public Property Let MemBuffer(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MemBuffer = 5
    mvarMemBuffer = vData
End Property

Public Property Get MemBuffer() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MemBuffer
    MemBuffer = mvarMemBuffer
End Property

Public Property Let MinVoltage(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MinVoltage = 5
    mvarMinVoltage = vData
End Property

Public Property Get MinVoltage() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MinVoltage
    MinVoltage = mvarMinVoltage
End Property

Public Property Let NumPoints(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.NumPoints = 5
    mvarNumPoints = vData
End Property

Public Property Get NumPoints() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.NumPoints
    NumPoints = mvarNumPoints
End Property

Public Property Let PeakVoltage(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PeakVoltage = 5
    mvarPeakVoltage = vData
End Property

Public Property Get PeakVoltage() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PeakVoltage
    PeakVoltage = mvarPeakVoltage
End Property

Public Property Let PtsPerPeriod(ByVal vData As Double)

    mvarPtsPerPeriod = vData

End Property

Public Property Get PtsPerPeriod() As Double

    PtsPerPeriod = mvarPtsPerPeriod

End Property

Public Property Set range(vData As range)

    Set mvarRange = vData

End Property

Public Property Get range() As range

    Set range = mvarRange

End Property

Public Property Set SineFitTextStream(ByVal vData As TextStream)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.SineFitTextStream = Form1
    Set mvarSineFitTextStream = vData
End Property

Public Property Get SineFitTextStream() As TextStream
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SineFitTextStream
    Set SineFitTextStream = mvarSineFitTextStream
End Property

Public Property Let SineFreqMax(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SineFreqMin = 5
    mvarSineFreqMax = vData
End Property

Public Property Get SineFreqMax() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SineFreqMin
    SineFreqMax = mvarSineFreqMax
End Property

Public Property Let SineFreqMin(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SineFreqMin = 5
    mvarSineFreqMin = vData
End Property

Public Property Get SineFreqMin() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SineFreqMin
    SineFreqMin = mvarSineFreqMin
End Property

Public Property Let Slope(ByVal vData As Double)

    mvarSlope = vData

End Property

Public Property Get Slope() As Double

    Slope = mvarSlope

End Property

Public Property Let StartPoint(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.StartPoint = 5
    mvarStartPoint = vData
End Property

Public Property Get StartPoint() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.StartPoint
    StartPoint = mvarStartPoint
End Property

Public Property Let TimeStep(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TimeStep = 5
    mvarTimeStep = vData
End Property

Public Property Get TimeStep() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TimeStep
    TimeStep = mvarTimeStep
End Property

Public Property Let WaveDesc(ByVal sData As String)

    mvarWaveDesc = sData

End Property

Public Property Get WaveDesc() As String

    WaveDesc = mvarWaveDesc

End Property

Public Property Let WaveININum(ByVal lData As Long)

    mvarWaveININum = lData

End Property

Public Property Get WaveININum() As Long

    WaveININum = mvarWaveININum
    
End Property

Public Property Let WaveName(ByVal sData As String)

    mvarWaveName = sData

End Property

Public Property Get WaveName() As String

    WaveName = mvarWaveName
    
End Property

Public Function GetBackgroundProcessStatus(ByVal IOFunction As Long, _
                                            ByRef DataArray() As Double, _
                                            ByRef ProcessDone As Boolean, _
                                            Optional ProcessName As String = "", _
                                            Optional GetPoints As Boolean = False, _
                                            Optional NumPts As Long = 0, _
                                            Optional NumInputChannels As Long = 1) As Boolean
                                            
    Dim ULStats As Long
    Dim Status As Integer
    Dim CurCount, CurIndex As Long
    Dim TempArray() As Integer
    Dim TempI As Integer
    Dim Temp2_I As Integer
    Dim TempSng As Single
    Dim TempD As Double
                               
    'Check for NOCOMM_MODE
    If NOCOMM_MODE = True Then
    
        GetBackgroundProcessStatus = False
                
        Exit Function
        
    End If
                               
    With Me
            
        'Need certain # of points from the Monitor memory buffer, going to try and get
        'them without checking if there are enough.
        ULStats = cbGetStatus(.BoardUsed.BoardNum, _
                                Status, _
                                CurCount, _
                                CurIndex, _
                                AIFUNCTION)
                            
        'Error Check
        If ULStats <> 0 Then
        
            'Send Email Error Msg - Code Yellow Status
            frmSendMail.MailNotification "Code Yellow", _
                                         "Could not get status of " & IOString & " " & ProcessName & _
                                         " background DAQ process on board """ & _
                                         WaveForms("ALTAFMONITOR").BoardUsed.BoardName & """." & _
                                         vbNewLine & ProcessName & " DAQ process will be aborted." & _
                                         vbNewLine & "Code exceution will continue.", _
                                         CodeYellow, _
                                         True
                    
            GetBackgroundProcessStatus = False
                    
            Exit Function
            
        End If
        
        'If the background process has finished (i.e. memory buffer is full),
        'then Set ProcessDone = True
        If Status = 0 Then
            
            ProcessDone = True
            
        Else
        
            ProcessDone = False
            
        End If
        
        
        .CurrentPoint = CurIndex
        
        If .CurrentPoint > (NumPts * NumInputChannels) And GetPoints Then
        
            'Size the two arrays needed to the correct value
            ReDim DataArray(NumPts * NumInputChannels)
            ReDim TempArray(NumPts * NumInputChannels)
        
            'Memory Buffer has necessary # of points in it, can get those points
            'and dump them as MCC integer counts into the TempArray()
            ULStats = cbWinBufToArray(.MemBuffer, _
                                        TempArray(0), _
                                        .CurrentPoint - (NumPts * NumInputChannels), _
                                        NumPts * NumInputChannels)
            
            'Error Check
            If ULStats <> 0 Then
            
                'Error Msg
                frmSendMail.MailNotification "Code Yellow", _
                                             "Could not get data for " & IOString & " " & ProcessName & _
                                             " DAQ process on board """ & .BoardUsed.BoardName & """." & _
                                             vbNewLine & ProcessName & " DAQ process will be aborted." & _
                                             vbNewLine & "Code execution is paused.", _
                                             CodeYellow, _
                                             True
                                        
                GetBackgroundProcessStatus = False
                        
                Exit Function
                        
            End If
            
            'Convert TempArray MCC integer count values to double-type voltage values
            For i = 0 To NumPts * NumInputChannels - 1 Step NumInputChannels
            
                TempI% = TempArray(i)
                
                cbToEngUnits .BoardUsed.BoardNum, _
                             .range.RangeType, _
                             TempI%, _
                             TempSng
                            
                DataArray(i) = CDbl(TempSng)
                
            Next i
            
        End If
                                       
    End With
    
    GetBackgroundProcessStatus = True
                                   
End Function

Public Function ManageBackgroundProcess(ByVal IOFunction As Long, _
                                        ByRef DataArray() As Double, _
                                        Optional ProcessName As String = "", _
                                        Optional KillProcess As Boolean = False, _
                                        Optional ReadFromBuffer As Boolean = False, _
                                        Optional NumInputChannels As Long = 1) As Boolean

    Dim ULStats As Long
    Dim IOString As String
    Dim TempSng As Single
    Dim TempD As Double
    Dim TempI As Integer
    Dim Temp2_I As Integer
    Dim TempArray() As Integer
    Dim gainArray(1) As Long
    Dim Status As Integer
    Dim CurCount, CurIndex As Long
    
    'Check for NOCOMM_MODE
    If NOCOMM_MODE = True Then
    
        ManageBackgroundProcess = False
        
        Exit Function
        
    End If
    

    If IOFunction = AOFUNCTION Then
    
        IOString = "Analog Output"
        
    Else
    
        IOString = "Analog Input"
        
    End If

    

    With Me

        'Don't stop background process before you read the memory buffer into a data-array
        'post-stop points default to MCC Count = 0, Voltage = -10
        If Not ReadFromBuffer Then
        
            'Reset the board being used by this wave
            ULStats = cbStopBackground(.BoardUsed.BoardNum, _
                                        IOFunction)
                                        
            'Error Check
            If ULStats <> 0 Then
            
                'Couldn't reset the Analog input monitor channel
                frmSendMail.MailNotification "CodeYellow", _
                                             "Unable to stop " & IOString & " background " & ProcessName & _
                                             " DAQ process on board """ & .BoardUsed.BoardName & """." & _
                                             vbNewLine & ProcessName & " DAQ process will be aborted." & _
                                             vbNewLine & "Code execution is paused.", _
                                             CodeYellow, _
                                             True
                                             
                ManageBackgroundProcess = False
                
                Exit Function
                
            End If
            
        End If

        If KillProcess Then
        
            'If the code needs to load the whole memory buffer
            'into a data array
            If ReadFromBuffer Then
                
                'First Get the current point in the windows memory buffer
                ULStats = cbGetStatus(.BoardUsed.BoardNum, _
                                      Status, _
                                      CurCount, _
                                      CurIndex, _
                                      IOFunction)
                                      
                'error check
                If ULStats <> 0 Then
                
                    frmSendMail.MailNotification _
                                "Code Yellow", _
                                "Could not get status of " & IOString & " " & ProcessName & _
                                " DAQ background process on board """ & .BoardUsed.BoardName & """." & _
                                vbNewLine & ProcessName & " process will be aborted." & vbNewLine & _
                                "AF Ramp Execution will continue.", _
                                CodeYellow, _
                                True
                                
                    ManageBackgroundProcess = False
                    
                    Exit Function
                    
                End If
                
                'Load the Current index into the wave storage field
                .CurrentPoint = CurIndex
                
                'Redimension the temp array and data array so that it can
                'receive the points from the memory buffer
                ReDim TempArray(.CurrentPoint - .StartPoint)
                ReDim DataArray(.CurrentPoint - .StartPoint)
                
                'Load the points into the temp array as MCC integer counts
                ULStats = cbWinBufToArray(.MemBuffer, _
                                            TempArray(0), _
                                            .StartPoint, _
                                            .CurrentPoint - .StartPoint)
                                            
                'Error Check
                If ULStats <> 0 Then
                
                    frmSendMail.MailNotification _
                                "Code Yellow", _
                                "Could not download data from windows memory buffer for the " & _
                                IOString & " " & ProcessName & " background DAQ process on " & _
                                "board """ & .BoardUsed.BoardName & """." & vbNewLine & _
                                ProcessName & " DAQ Process will be aborted." & vbNewLine & _
                                "Code execution is paused.", _
                                CodeYellow, _
                                True
                    
                    ManageBackgroundProcess = False
                    
                    Exit Function
                    
                End If
                
                'Now need to convert all those points to single data type voltages
                For i = 0 To .CurrentPoint - .StartPoint - 1
                
                    TempI = TempArray(i)
                    
                    'Range converter requires three inputs - 1 double (the voltage var.),
                    'and 2 integers (the MCC Counts value & the fixed counts value - see
                    'MCC_RangeConverter subroutine in the Range class module for a
                    'description of the differences between these two).
                    cbToEngUnits .BoardUsed.BoardNum, _
                                 mvarRange.RangeType, _
                                 TempI, _
                                 TempSng
                                                  
                    DataArray(i) = CDbl(TempSng)
                                              
                Next i
                
                'Now Stop the background process on the board
                ULStats = cbStopBackground(.BoardUsed.BoardNum, _
                                            IOFunction)
                                            
                'Error Check
                If ULStats <> 0 Then
                
                    'Couldn't reset the Analog input monitor channel
                    frmSendMail.MailNotification _
                            "Code Yellow", _
                            "Unable to Stop " & IOString & " background " & ProcessName & _
                            " DAQ process on board """ & .BoardUsed.BoardName & """." & _
                            vbNewLine & ProcessName & " DAQ Process will be aborted." & _
                            vbNewLine & "Code execution is paused.", _
                            CodeYellow, _
                            True
                           
                    ManageBackgroundProcess = True
                    
                    Exit Function
                    
                End If
                
            'Done processing data from the memory buffer
            End If

            'If ReadBuffer = false, then the background process on has already been stopped
            'and doesn't need to be called again, here
            
        
            'If memory buffer is allocated, deallocate it
            If .BufferAlloc = True Then
            
                'MCC windows memory buffer deallocate function call
                ULStats = cbWinBufFree(.MemBuffer)
                
                'Error check
                If ULStats <> 0 Then
                
                    frmSendMail.MailNotification _
                            "Code Yellow", _
                            "Could not deallocate " & ProcessName & " DAQ Process memory buffer for " & _
                            "board """ & .BoardUsed.BoardName & """." & vbNewLine & _
                            ProcessName & " DAQ process will be aborted." & vbNewLine & _
                            "Code execution is paused.", _
                            CodeYellow, _
                            True
                                                        
                    ManageBackgroundProcess = True
                    
                    Exit Function
                    
                End If
                
                'No error, so Memory buffer was successfully deallocated
                .BufferAlloc = False
                
            End If
            
            'If the process is an analog output process, zero the wave chan as a last act
            If IOFunction = AOFUNCTION Then
            
                'Write analog zero voltage to wave's output channel
                ULStats = .BoardUsed.AnalogOut(.Chan, _
                                               0)
                                                               
                'Error check
                If ULStats <> 0 Then

                    frmSendMail.MailNotification _
                                "Code Yellow", _
                                "Could not zero channel """ & .Chan.ChanName & """ used for the " & _
                                ProcessName & " background DAQ process on board """ & _
                                .BoardUsed.BoardName & """." & vbNewLine & ProcessName & _
                                " DAQ process will be aborted." & vbNewLine & _
                                "Code execution has been paused.", _
                                CodeYellow, _
                                True
                                
                    ManageBackgroundProcess = True
                                    
                    Exit Function
                                    
                End If
                
            End If
                  
        Else
        'User is trying to setup a background / IO process, not kill a process
        
            'Now Create an memory buffer for the background process
            'If Windows memory problems are going to happen, they're going to happen here!
            On Error Resume Next
            
                .MemBuffer = cbWinBufAlloc(.NumPoints * NumInputChannels)
                
                'Error Check
                If .MemBuffer = 0 Or Err.number <> 0 Then
                
                    'Crap, an error of some sort happened
                    frmSendMail.MailNotification _
                                "Code Yellow", _
                                "Could not allocate windows memory buffer for " & ProcessName & " DAQ process " & _
                                "on board """ & .BoardUsed.BoardName & """." & vbNewLine & ProcessName & _
                                " DAQ process will be aborted." & vbNewLine & _
                                "Code execution has been paused.", _
                                CodeYellow, _
                                True
                            
                    ManageBackgroundProcess = False
                    
                    Exit Function
                    
                End If
            
            On Error GoTo 0
            
            'Indicate that the wave memory buffer has been allocated
            .BufferAlloc = True
        
            If IOFunction = AIFUNCTION Then
        
                'Start the monitor wave background analog input process
                'Note - mutiply the Wave.NumPoints to scan in
                'value by the number of channels - will need to pull in
                'more points to get the same time duration of background process
                ULStats = cbAInScan(.BoardUsed.BoardNum, _
                                    .Chan.ChanNum, _
                                    .Chan.ChanNum + NumInputChannels - 1, _
                                    .NumPoints * NumInputChannels, _
                                    .IORate, _
                                    .range.RangeType, _
                                    .MemBuffer, _
                                    .IOOptions)
                
                'Error check
                If ULStats <> 0 Then
                    
                    frmSendMail.MailNotification _
                                "Code Yellow", _
                                "Could not start the " & IOString & " background " & _
                                ProcessName & " DAQ process on board """ & _
                                .BoardUsed.BoardName & """." & vbNewLine & ProcessName & _
                                " DAQ process will be aborted." & vbNewLine & _
                                "Code execution has been paused.", _
                                CodeYellow, _
                                True
                                
                    ManageBackgroundProcess = False
                    
                    Exit Function
                    
                End If
                
            Else
            
                'Redimension the temp array used to convert the output array
                'into the MCC integer counts
                ReDim TempArray(.NumPoints)
                
                'Convert double-type voltage values in the DataArray
                'into MCC Integer counts
                For i = 0 To .NumPoints - 1
                
                    TempSng = CSng(DataArray(i))
                
                    'Range converter requires three inputs - 1 double (the voltage var.),
                    'and 2 integers (the MCC Counts value & the fixed counts value - see
                    'MCC_RangeConverter subroutine in the Range class module for a
                    'description of the differences between these two).
                    cbFromEngUnits .BoardUsed.BoardNum, _
                                   mvarRange.RangeType, _
                                   TempSng, _
                                   TempI
                    
                    TempArray(i) = TempI
                
                Next i
                            
                'Load those integer counts into the memory buffer
                ULStats = cbWinArrayToBuf(TempArray(0), _
                                          .MemBuffer, _
                                          .StartPoint, _
                                          .NumPoints)
                
                'Error Check
                If ULStats <> 0 Then
                
                    frmSendMail.MailNotification _
                                "Code Yellow", _
                                "Could not upload output data array " & _
                                "for the " & IOString & " " & ProcessName & " background DAQ process " & _
                                "windows memory buffer on board """ & .BoardUsed.BoardName & """." & _
                                ProcessName & " DAQ process will be aborted." & vbNewLine & _
                                "Code execution has been paused.", _
                                CodeYellow, _
                                True
                                
                    ManageBackgroundProcess = False
                    
                    Exit Function
                    
                End If
                            
                            
                'Start the wave background analog output process
                ULStats = cbAOutScan(.BoardUsed.BoardNum, _
                                    .Chan.ChanNum, _
                                    .Chan.ChanNum, _
                                    .NumPoints, _
                                    .IORate, _
                                    .range.RangeType, _
                                    .MemBuffer, _
                                    .IOOptions)
                
                'Error check
                If ULStats <> 0 Then
                    
                
                    frmSendMail.MailNotification _
                                "Code Yellow", _
                                "Could not start the " & IOString & " background " & _
                                ProcessName & " DAQ process on board """ & .BoardUsed.BoardName & """." & _
                                vbNewLine & ProcessName & " DAQ process will be aborted." & vbNewLine & _
                                "Code execution has been paused.", _
                                CodeYellow, _
                                True
                                
                    ManageBackgroundProcess = False
                    
                    Exit Function
                    
                End If
                
            End If
            
        End If
        
    End With
    
    ManageBackgroundProcess = True
        
End Function

Public Sub setrange(ByVal vData As Long)

    If IsEmpty(mvarRange) Or mvarRange Is Nothing Then
    
        Set mvarRange = Nothing
        Set mvarRange = New range
        
    End If
    
    mvarRange.RangeType = vData
        
End Sub

