VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Board"
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
Private mvarBoardININum As Long
Private mvarBoardNum As Long 'local copy
Private mvarMaxAInRate As Long 'local copy
Private mvarMaxAOutRate As Long 'local copy
Private mvarBoardMode As Long 'local copy
Private mvarDIOConfigured As Boolean ' local copy
Private mvarAInChannels As Channels 'local copy
Private mvarAOutChannels As Channels 'local copy
Private mvarDInChannels As Channels 'local copy
Private mvarDOutChannels As Channels 'local copy
Private mvarRange As range 'local copy
Private mvarBoardName As String 'local copy
Private mvarCommProtocol As BoardProtocol 'local copy
Private mvarBoardFunction As String 'local copy
Private mvarDoutPortType As Long

Public Enum BoardMode_enum
    SINGLEMODE = 1
    DIFFERENTIALMODE = 0
End Enum

Public Enum BoardProtocol
    MCC_UL = 1
    ADWIN_COM = 2
End Enum

Public Property Set AInChannels(ByVal vData As Channels)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.AInChannels = Form1

    If vData Is Nothing Then
    
        Set vData = New Channels
        
    End If

    Set mvarAInChannels = vData
End Property

Public Property Get AInChannels() As Channels
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.AInChannels

    If mvarAInChannels Is Nothing Then
    
        Set mvarAInChannels = New Channels
        
    End If

    Set AInChannels = mvarAInChannels
        
End Property

Public Property Set AOutChannels(vData As Channels)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.AOutChannels = Form1

    If vData Is Nothing Then
    
        Set vData = New Channels
        
    End If

    Set mvarAOutChannels = vData
    
End Property

Public Property Get AOutChannels() As Channels
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.AOutChannels

    If mvarAOutChannels Is Nothing Then
    
        Set mvarAOutChannels = New Channels
        
    End If

    Set AOutChannels = mvarAOutChannels
        
End Property

Public Property Let BoardFunction(ByVal vData As String)
Attribute BoardFunction.VB_UserMemId = 0
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.BoardFunction = Form1
    mvarBoardFunction = vData
End Property

Public Property Get BoardFunction() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.BoardFunction
    BoardFunction = mvarBoardFunction
End Property

Public Property Let BoardININum(ByVal lData As Long)

    mvarBoardININum = lData

End Property

Public Property Get BoardININum() As Long

    BoardININum = mvarBoardININum

End Property

Public Property Let BoardMode(ByVal vData As BoardMode_enum)

    mvarBoardMode = vData

End Property

Public Property Get BoardMode() As BoardMode_enum

    BoardMode = mvarBoardMode

End Property

Public Property Let BoardName(ByVal vData As String)

    mvarBoardName = vData

End Property

Public Property Get BoardName() As String
    
    BoardName = mvarBoardName

End Property

Public Property Let BoardNum(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.BoardNum = 5
    mvarBoardNum = vData
End Property

Public Property Get BoardNum() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.BoardNum
    BoardNum = mvarBoardNum
End Property

'local variable(s) to hold property value(s)
Public Property Let CommProtocol(ByVal vData As BoardProtocol)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.CommProtocol = Form1
    mvarCommProtocol = vData
End Property

Public Property Get CommProtocol() As BoardProtocol
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CommProtocol
    CommProtocol = mvarCommProtocol
End Property

Public Property Set DInChannels(ByVal vData As Channels)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.DInChannels = Form1

    If vData Is Nothing Then
    
        Set vData = New Channels
        
    End If

    Set mvarDInChannels = vData
End Property

Public Property Get DInChannels() As Channels
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DInChannels
    
    If mvarDInChannels Is Nothing Then
    
        Set mvarDInChannels = New Channels
        
    End If
    
    Set DInChannels = mvarDInChannels
    
End Property

Public Property Let DIOConfigured(ByVal vData As Boolean)

    mvarDIOConfigured = vData

End Property

Public Property Get DIOConfigured() As Boolean

    DIOConfigured = mvarDIOConfigured

End Property

Public Property Set DOutChannels(ByVal vData As Channels)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.DOutChannels = Form1
    
    If vData Is Nothing Then
    
        Set vData = New Channels
        
    End If
    
    Set mvarDOutChannels = vData
End Property

Public Property Get DOutChannels() As Channels
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DOutChannels

    If mvarDOutChannels Is Nothing Then
    
        Set mvarDOutChannels = New Channels
        
    End If

    Set DOutChannels = mvarDOutChannels
    
End Property

Public Property Let DoutPortType(ByVal lData As Long)

    mvarDoutPortType = lData
    
End Property

Public Property Get DoutPortType() As Long

    DoutPortType = mvarDoutPortType
    
End Property

Public Property Let MaxAInRate(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MaxAInRate = 5
    mvarMaxAInRate = vData
End Property

Public Property Get MaxAInRate() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MaxAInRate
    MaxAInRate = mvarMaxAInRate
End Property

Public Property Let MaxAOutRate(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MaxAOutRate = 5
    mvarMaxAOutRate = vData
End Property

Public Property Get MaxAOutRate() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MaxAOutRate
    MaxAOutRate = mvarMaxAOutRate
End Property

Public Property Set range(ByVal rData As range)

    Set mvarRange = rData

End Property

Public Property Get range() As range

    Set range = mvarRange

End Property

Public Function AnalogIn(ByVal InputChannel As Channel) As Variant

    Dim InputValue As Integer
    Dim FixedCounts As Integer
    Dim voltage As Single
    Dim ULStats As Long
    
    'Check for NOCOMM_MODE
    If NOCOMM_MODE = True Then
    
        'Return -616 - custom Paleomag error code
        AnalogIn = -616
        
        Exit Function
        
    End If
    
    If mvarCommProtocol = MCC_UL Then
    
        'Get channel input
        ULStats = cbAIn(mvarBoardNum, _
                        InputChannel.ChanNum, _
                        mvarRange.RangeType, _
                        InputValue)
                        
        'Error Check
        If ULStats <> 0 Then
        
            'Put in code to halt the code and trigger a code Red and email
            'the error text to the user
            
            frmSendMail.MailNotification "Analog Input Error", _
                                         "Analog Input Error on board = """ & mvarBoardName & """" & _
                                         vbNewLine & vbNewLine & _
                                         "An error has been raised in the code, execution is halted. " & _
                                         "Please come and check the machine.", _
                                         CodeRed, _
                                         True
                                         
            SetCodeLevel CodeRed
            
            Flow_Pause
        
            'Raise Error
            Err.Raise ULStats, _
                      "Board->AnalogIn", _
                      "Unable to read analog input value on: " & vbNewLine & _
                      "Board = " & Trim(mvarBoardName) & " (" & _
                      Trim(Str(mvarBoardNum)) & ")" & vbNewLine & _
                      "Channel = " & Trim(InputChannel.ChanName) & " (" & _
                      Trim(Str(InputChannel.ChanNum)) & ")" & vbNewLine & vbNewLine & _
                      "Please Refer to Error Codes section of the " & _
                      "Measurement Computing Universal Library Help PDF Guide."
                      
            'For now, return "ERR: " & error-code
            AnalogIn = "ERR: " & Trim(Str(ULStats))
            
            'End the function
            Exit Function
            
        End If
                        
        'Convert MCC counts to voltage
        cbToEngUnits mvarBoardNum, _
                     mvarRange.RangeType, _
                     InputValue, _
                     voltage

        'Return the voltage
        AnalogIn = CDbl(voltage)
        
    ElseIf mvarCommProtocol = ADWIN_COM Then
    
        'Send error message explaining that this isn't possible right now
        MsgBox "Sorry, it is not yet possible for users to read a " & _
               "single Analog input value from the Analog input channels " & _
               "on the " & Trim(mvarBoardName) & " board.", , _
               "Unsupported Comm Error"
               
        AnalogIn = "ERR"
               
    End If
    
End Function

Public Function AnalogOut(ByRef OutputChannel As Channel, ByVal voltage As Double) As Long
Attribute AnalogOut.VB_Description = "This functions is used to implement the analog out board functions in a non-board specific manner.  The function will look at the board and channel object given to it, assess what type of board it is, and use the right function based on a board-list"

    Dim OutputValue As Integer
    Dim FixedCounts As Integer
    Dim TempSng As Single
    Dim ULStats As Long
    
    'Check for NOCOMM_MODE
    If NOCOMM_MODE = True Then
    
        'Return -616 - custom Paleomag error code
        AnalogOut = -616
        
        Exit Function
        
    End If
    
    'Start with error code as zero
    AnalogOut = 0
    
    If mvarCommProtocol = MCC_UL Then
            
        TempSng = CSng(voltage)
        
        'Convert volts to counts
        cbFromEngUnits mvarBoardNum, _
                       mvarRange.RangeType, _
                       TempSng, _
                       OutputValue
                                        
        'Output Counts through board now
        ULStats = cbAOut(mvarBoardNum, _
                         OutputChannel.ChanNum, _
                         mvarRange.RangeType, _
                         OutputValue)
                            
        'Send Error Status Back
        AnalogOut = ULStats

    ElseIf mvarCommProtocol = ADWIN_COM Then
    
        'Send error message explaining that this isn't possible right now
        MsgBox "Sorry, it is not yet possible for users to read a " & _
               "single Analog input value from the Analog input channels " & _
               "on the " & Trim(mvarBoardName) & " board.", , _
               "Unsupported Comm Error"
        
        'Return a -1 status value
        AnalogOut = -1
               
    End If

End Function

Public Function CalcADWINDigOutBit(ByRef DigOut_Chan As Channel, _
                                   ByVal SetHigh As Boolean, _
                                   Optional ByVal OneChanOn As Boolean = True) As Long
                                   
    Dim BitValue As Long
                                   
    'Need is the channel being turned on or off
    If SetHigh = True Then
    
        BitValue = 2 ^ (DigOut_Chan.ChanNum)
        
    Else
    
        BitValue = 0
        
    End If

    'Check to see if the user wants only one digital channel on at
    'once
    If OneChanOn = False Then
        'Other channels can stay open, so need to see what other
        'channels are open
    
        'Now need to see what other channels are on so that we
        'don't affect them when we change the desired channel
        If modAF_DAQ.ADWIN_LastDigOut_BitNum > -1 And _
           modAF_DAQ.ADWIN_LastDigOut_BitNum < 64 Then
               
            'Last Value is legit
            'Check to see what ports are open
            If modAF_DAQ.ADWIN_LastDigOut_BitNum Mod 2 = 1 And _
               DigOut_Chan.ChanNum <> 0 Then
                
                '0th Digital channel needs to stay open, add one
                BitValue = BitValue + 1
                
            End If
            
            If (modAF_DAQ.ADWIN_LastDigOut_BitNum \ 2) Mod 2 = 1 And _
               DigOut_Chan.ChanNum <> 1 Then
               
                '1st Digital channel needs to stay open, add two
                BitValue = BitValue + 2
               
            End If
            
            If (modAF_DAQ.ADWIN_LastDigOut_BitNum \ 4) Mod 2 = 1 And _
               DigOut_Chan.ChanNum <> 2 Then
            
                '2nd Digital channel needs to stay open, add four
                BitValue = BitValue + 4
            
            End If
            
            If (modAF_DAQ.ADWIN_LastDigOut_BitNum \ 8) Mod 2 = 1 And _
               DigOut_Chan.ChanNum <> 3 Then
               
                '3rd Digital channel needs to stay open, add eight
                BitValue = BitValue + 8
               
            End If
            
            If (modAF_DAQ.ADWIN_LastDigOut_BitNum \ 16) Mod 2 = 1 And _
               DigOut_Chan.ChanNum <> 4 Then
               
                '4th Digital channel needs to stay open, add sixteen
                BitValue = BitValue + 16
               
            End If
            
            If (modAF_DAQ.ADWIN_LastDigOut_BitNum \ 32) Mod 2 = 1 And _
               DigOut_Chan.ChanNum <> 5 Then
               
                '5th Digital channel needs to stay open, add thirty-two
                BitValue = BitValue + 32
               
            End If
         
        End If
        
    End If
    
    CalcADWINDigOutBit = BitValue
                                   
End Function

Private Sub Class_Terminate()
    Set mvarChannels = Nothing
End Sub

Public Function DigitalIn_ADWIN(InChannel As Channel) As Long

    Dim InputVal As Long
    
    'Turn off the error-pop-up in the Boot process
    ADWIN.Show_Errors (0)
    
    'Check to see that the ADWIN board has been booted
    If ADWIN.ADWIN_BootBoard(Me) = False Then
        
'            'Raise Error
'            Err.Raise -1, _
'                      "optCoil_Click->ADWIN.ADWIN_BootBoard", _
'                      "Was unable to boot the ADWIN board system." & _
'                      vbNewLine & "ADWIN Dev # = " & Trim(Str(TTLBoard.BoardNum))
                      
       'Return Error value = -616 indicating bad ADWIN board boot
        DigitalIn_ADWIN = -616
                      
        Exit Function
            
    End If
    
    'Get the Input Bit setting
    InputVal = ADWIN.Dig_In
    
    'Need to translate that input value to see whether or not the inputed channel
    'is set to high (1) or low (0)
    If InputVal >= 0 And InputVal <= 63 Then
        
        If (InputVal \ 2 ^ (InChannel.ChanNum - 1)) Mod 2 = 0 Then
        
            'This channel is off
            DigitalIn_ADWIN = 0
            
        Else
        
            'This channel is on
            DigitalIn_ADWIN = 1
            
        End If
        
    Else
    
        'Error, Return -1
        DigitalIn_ADWIN = -1
        
    End If
    
End Function

Public Function DigitalIn_MCC _
    (ByRef InChannel As Channel, _
     Optional ByVal PortType As Long = -1) As Long

    Dim ULStats As Long
    Dim InputValue As Integer
    
    'Check to see if a dPortType value was actually
    'given by the external calling function
    'if not, set it to the type required for this MCC board
    If PortType = -1 Then
    
        PortType = mvarDoutPortType
    
    End If
    
   'MCC comm protocol
   'Need to use PortType  and bitnum = InChannel.ChanNum
   ULStats = cbDConfigBit(mvarBoardNum, _
                          PortType, _
                          InChannel.ChanNum, _
                          DIGITALIN)
        
    'Check for Error
    If ULStats <> 0 Then
    
        'Return Error
        DigitalIn_MCC = ULStats
        
        Exit Function
        
    End If
    
    'Now use bit out command to write to port
    ULStats = cbDBitIn(mvarBoardNum, _
                       PortType, _
                       InChannel.ChanNum, _
                       InputValue)
    
    'Check for Error
    If ULStats <> 0 Then

        'Return Error
        DigitalIn_MCC = ULStats
        
        Exit Function
        
    End If
        
    'If InputValue = 1 need to change the sign to distinguish the return from an
    'MCC Error = 1, BADBOARD
    If InputValue = 1 Then InputValue = -1
    
    'Digital input is done, return the read in value
    DigitalIn_MCC = InputValue
    
End Function

Public Function DigitalInput _
    (ByRef InChan As Channel, _
     Optional ByVal dPortType As Long = -1) As Long
     
    Dim ReturnVal As Long
     
    'Default Digital Input = -2
    DigitalInput = -2
    
     'Check for NOCOMM_MODE - NOTE: can just exit because return value is set above
    If NOCOMM_MODE = True Then Exit Function
     
    If mvarCommProtocol = MCC_UL Then
    
        ReturnVal = DigitalIn_MCC(InChan, dPortType)
        
    ElseIf mvarCommProtocol = ADWIN_COM Then
    
        ReturnVal = DigitalIn_ADWIN(InChan)
        
    End If
    
    DigitalInput = ReturnVal
    
End Function

Public Function DigitalOut_ADWIN(ByVal DigOut_Bit As Long)

    If mvarCommProtocol = ADWIN_COM And DigOut_Bit >= 0 And DigOut_Bit <= 63 Then
    
        'Totally different Digital out setting system is used than for MCC
        
        'user should have inputed a long with a value of 0 - 63

        'Turn off the error-pop-up in the Boot process
        ADWIN.Show_Errors (0)

        'Check to make sure the ADWIN board is booted
        If ADWIN.ADWIN_BootBoard(Me) = False Then
        
'            'Raise Error
'            Err.Raise -1, _
'                      "optCoil_Click->ADWIN.ADWIN_BootBoard", _
'                      "Was unable to boot the ADWIN board system." & _
'                      vbNewLine & "ADWIN Dev # = " & Trim(Str(TTLBoard.BoardNum))
                      
            'Return Error value = -616
            DigitalOut_ADWIN = -616
                      
            Exit Function
            
        End If
        
        'Otherwise, the ADWIN board is booted and ready to change the digital ports
        'configuration
        ADWIN.Set_Digout DigOut_Bit
        
        modAF_DAQ.ADWIN_LastDigOut_BitNum = DigOut_Bit
        
        'Return Zero
        
        DigitalOut_ADWIN = 0
            
    Else
    
        'User has put in the wrong type of input, return a -616 to denote this
        'type of error
        DigitalOut_ADWIN = -616
        
    End If

End Function

Public Function DigitalOut_MCC(ByRef OutChannel As Channel, ByVal SetHigh As Boolean, Optional ByVal PortType As Long = AUXPORT) As Long

    Dim ULStats As Long
    Dim DataValue As Long
    
    If mvarCommProtocol = MCC_UL Then
    
        'Need to use PortType and portnum = OutChannel.
        ULStats = cbDConfigBit(mvarBoardNum, _
                               PortType, _
                               OutChannel.ChanNum, _
                               DIGITALOUT)
        
        'Error Check
        If ULStats <> 0 Then
            
            'Return Error
            DigitalOut_MCC = ULStats
            
            Exit Function
            
        End If
                
        'Determine from the input whether to set the
        'port to the high value (1) or the low value (0)
        If SetHigh = True Then

            DataValue = 1
            
        Else
        
            DataValue = 0
            
        End If
                
        'Now use bit out command to write to port
        ULStats = cbDBitOut(mvarBoardNum, _
                            PortType, _
                            OutChannel.ChanNum, _
                            DataValue)
        
        'Error check
        If ULStats <> 0 Then
            
            'Return error
            DigitalOut_MCC = ULStats
            
            Exit Function
            
        End If
        
        DigitalOut_MCC = 0
        
    Else
    
        'Bad input / wrong function, return -616
        DigitalOut_MCC = -616
        
    End If
    
End Function

Public Function DigitalOutput _
    (ByRef DigOut_Chan As Channel, _
     ByVal SetHigh As Boolean, _
     Optional ByVal OneChanOn As Boolean = True) As Long

    Dim ReturnVal As Long
    Dim BitValue As Long
    
    'Default Digital output to -2
    DigitalOutput = -2
    
     'Check for NOCOMM_MODE - NOTE: can just exit because return value is set above
    If NOCOMM_MODE = True Then Exit Function
    
    If mvarCommProtocol = MCC_UL Then
    
        'Can just call the MCC digital output function
        ReturnVal = Me.DigitalOut_MCC(DigOut_Chan, _
                                      SetHigh, _
                                      mvarDoutPortType)
                                      
        DigitalOutput = ReturnVal
        
    ElseIf mvarCommProtocol = ADWIN_COM Then
    
        BitValue = CalcADWINDigOutBit(DigOut_Chan, _
                                      SetHigh, _
                                      OneChanOn)
                           
        'Now, using this bit value, call the ADWIN digital output function
        ReturnVal = Me.DigitalOut_ADWIN(BitValue)
           
        'Return the return val to the larger calling function
        DigitalOutput = ReturnVal
                           
    End If
    
End Function

Public Function Get_DOutPortTypeStr() As String

    Dim TempStr As String

    Select Case mvarDoutPortType
    
        Case AUXPORT
        
            TempStr = "AUXPORT"
        
        
        Case FIRSTPORTA
        
            TempStr = "FIRSTPORTA"
        
        
        Case FIRSTPORTB
        
            TempStr = "FIRSTPORTB"
        
        
        Case FIRSTPORTCL
        
            TempStr = "FIRSTPORTCL"
        
        
        Case FIRSTPORTC
        
            TempStr = "FIRSTPORTC"
        
        
        Case FIRSTPORTCH
        
            TempStr = "FIRSTPORTCH"
        
        
        Case SECONDPORTA
        
            TempStr = "SECONDPORTA"
        
        
        Case SECONDPORTB
        
            TempStr = "SECONDPORTB"
        
        
        Case SECONDPORTCL
        
            TempStr = "SECONDPORTCL"
        
        
        Case SECONDPORTCH
        
            TempStr = "SECONDPORTCH"
        
        
        Case THIRDPORTA
        
            TempStr = "THIRDPORTA"
        
        
        Case THIRDPORTB
        
            TempStr = "THIRDPORTB"
        
        
        Case THIRDPORTCL
        
            TempStr = "THIRDPORTCL"
        
        
        Case THIRDPORTCH
        
            TempStr = "THIRDPORTCH"
        
        
        Case FOURTHPORTA
        
            TempStr = "FOURTHPORTA"
        
        
        Case FOURTHPORTB
        
            TempStr = "FOURTHPORTB"
        
        
        Case FOURTHPORTCL
        
            TempStr = "FOURTHPORTCL"
        
        
        Case FOURTHPORTCH
        
            TempStr = "FOURTHPORTCH"
        
        
        Case FIFTHPORTA
        
            TempStr = "FIFTHPORTA"
        
        
        Case FIFTHPORTB
        
            TempStr = "FIFTHPORTB"
        
        
        Case FIFTHPORTCL
        
            TempStr = "FIFTHPORTCL"
        
        
        Case FIFTHPORTCH
        
            TempStr = "FIFTHPORTCH"
        
        
        Case SIXTHPORTA
        
            TempStr = "SIXTHPORTA"
        
        
        Case SIXTHPORTB
        
            TempStr = "SIXTHPORTB"
        
        
        Case SIXTHPORTCL
        
            TempStr = "SIXTHPORTCL"
        
        
        Case SIXTHPORTCH
        
            TempStr = "SIXTHPORTCH"
        
        
        Case SEVENTHPORTA
        
            TempStr = "SEVENTHPORTA"
        
        
        Case SEVENTHPORTB
        
            TempStr = "SEVENTHPORTB"
        
        
        Case SEVENTHPORTCL
        
            TempStr = "SEVENTHPORTCL"
        
        
        Case SEVENTHPORTCH
        
            TempStr = "SEVENTHPORTCH"
        
        
        Case EIGHTHPORTA
        
            TempStr = "EIGHTHPORTA"
        
        
        Case EIGHTHPORTB
        
            TempStr = "EIGHTHPORTB"
        
        
        Case EIGHTHPORTCL
        
            TempStr = "EIGHTHPORTCL"
        
        
        Case EIGHTHPORTCH
        
            TempStr = "EIGHTHPORTCH"
            
        Case Else
        
            TempStr = "ERROR"
            
    End Select

    Get_DOutPortTypeStr = TempStr
    
End Function

