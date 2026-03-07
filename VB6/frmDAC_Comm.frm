VERSION 5.00
Begin VB.Form frmDAQ_Comm 
   Caption         =   "DAQ Boards Comm Control"
   ClientHeight    =   4515
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6135
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4515
   ScaleWidth      =   6135
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Default         =   -1  'True
      Height          =   372
      Left            =   240
      TabIndex        =   26
      Top             =   4080
      Width           =   2652
   End
   Begin VB.Frame Frame2 
      Caption         =   "Digital IO"
      Height          =   3852
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   2892
      Begin VB.CommandButton cmdDigitalIn 
         Caption         =   "Digital In"
         Height          =   372
         Left            =   1560
         TabIndex        =   25
         Top             =   3360
         Width           =   1092
      End
      Begin VB.CommandButton cmdDigitalOut 
         Caption         =   "Digital Out"
         Height          =   372
         Left            =   240
         TabIndex        =   24
         Top             =   3360
         Width           =   1092
      End
      Begin VB.ComboBox cmbBoardD 
         Height          =   288
         Left            =   120
         TabIndex        =   22
         Top             =   480
         Width           =   2532
      End
      Begin VB.OptionButton optHighLow 
         Caption         =   "High (1)"
         Height          =   252
         Index           =   1
         Left            =   960
         TabIndex        =   21
         Top             =   3000
         Width           =   1212
      End
      Begin VB.OptionButton optHighLow 
         Caption         =   "Low  (0)"
         Height          =   252
         Index           =   0
         Left            =   960
         TabIndex        =   20
         Top             =   2640
         Width           =   1212
      End
      Begin VB.ComboBox cmbChanDOut 
         Height          =   288
         Left            =   120
         TabIndex        =   10
         Top             =   1920
         Width           =   2052
      End
      Begin VB.ComboBox cmbChanDIn 
         Height          =   288
         Left            =   120
         TabIndex        =   9
         Top             =   1200
         Width           =   2052
      End
      Begin VB.Label Label4 
         Caption         =   "Digital Channel Value:"
         Height          =   252
         Left            =   360
         TabIndex        =   23
         Top             =   2400
         Width           =   1692
      End
      Begin VB.Label Label5 
         Caption         =   "Board:"
         Height          =   252
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   1212
      End
      Begin VB.Label Label8 
         Caption         =   "Input Channel:"
         Height          =   252
         Left            =   120
         TabIndex        =   12
         Top             =   960
         Width           =   1212
      End
      Begin VB.Label Label26 
         Caption         =   "Output Channel:"
         Height          =   252
         Left            =   120
         TabIndex        =   11
         Top             =   1680
         Width           =   1212
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Analog IO"
      Height          =   4332
      Left            =   3120
      TabIndex        =   0
      Top             =   120
      Width           =   2892
      Begin VB.CommandButton cmdAnalogIn 
         Caption         =   "Analog In"
         Height          =   372
         Left            =   1560
         TabIndex        =   19
         Top             =   3360
         Width           =   1092
      End
      Begin VB.CommandButton cmdAnalogOut 
         Caption         =   "Analog Out"
         Height          =   372
         Left            =   240
         TabIndex        =   18
         Top             =   3360
         Width           =   1092
      End
      Begin VB.TextBox txtRawA 
         Height          =   288
         Left            =   1320
         TabIndex        =   17
         Top             =   2880
         Width           =   1092
      End
      Begin VB.TextBox txtEngA 
         Height          =   288
         Left            =   1320
         TabIndex        =   16
         Top             =   2400
         Width           =   1092
      End
      Begin VB.ComboBox cmbChanAIn 
         Height          =   288
         Left            =   120
         TabIndex        =   15
         Top             =   1200
         Width           =   2052
      End
      Begin VB.ComboBox cmbBoardA 
         Height          =   288
         Left            =   120
         TabIndex        =   14
         Top             =   480
         Width           =   2532
      End
      Begin VB.ComboBox cmbChanAOut 
         Height          =   288
         Left            =   120
         TabIndex        =   1
         Top             =   1920
         Width           =   2052
      End
      Begin VB.Label lblBoard 
         Caption         =   "Board:"
         Height          =   252
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   972
      End
      Begin VB.Label Label3 
         Caption         =   "English Value:"
         Height          =   252
         Left            =   120
         TabIndex        =   6
         Top             =   2400
         Width           =   1092
      End
      Begin VB.Label Label2 
         Caption         =   "Board Value:"
         Height          =   252
         Left            =   120
         TabIndex        =   5
         Top             =   2880
         Width           =   972
      End
      Begin VB.Label Label1 
         Caption         =   "Input Channel:"
         Height          =   252
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   1212
      End
      Begin VB.Label Label25 
         Caption         =   "Output Channel:"
         Height          =   252
         Left            =   120
         TabIndex        =   3
         Top             =   1680
         Width           =   1212
      End
      Begin VB.Label Label29 
         Alignment       =   2  'Center
         Caption         =   "Outputs read English value."
         Height          =   252
         Left            =   360
         TabIndex        =   2
         Top             =   3960
         Width           =   2292
      End
   End
End
Attribute VB_Name = "frmDAQ_Comm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This status flag shows whether or not the DAQ Board comm form has been loaded
Dim FormStatus As Boolean
Dim DigitalBit As Boolean

Private Sub cmbBoardA_Click()

    Dim i As Long
    Dim N As Long
    
    'Change analog output and input channel combo boxes
    Me.cmbChanAIn.Clear
    Me.cmbChanAOut.Clear
    
    'Need to check the System Boards collection
    On Error Resume Next
    
        'if an error happens here, then need to load boards
        If SystemBoards.Count = 0 Then
        
            Exit Sub
            
        End If
        
        If Err.number <> 0 Then
        
            Exit Sub
            
        End If
        
    On Error GoTo 0
    
    'Use the item-data of the current cmbBoardA element as the
    'key to get the intended board
    With SystemBoards(cmbBoardA.List(cmbBoardA.ListIndex))
    
        'Load all of the Analog Output channels into the
        'Analog Output Channel selector combo box
        N = .AOutChannels.Count
        
        For i = 1 To N
        
            cmbChanAOut.AddItem .AOutChannels(i).ChanName
            
        Next i
        
        'Load all of the Analog Input channels into the
        'Analog Input Channel selector combo box
        N = .AInChannels.Count
        
        For i = 1 To N
        
            cmbChanAIn.AddItem .AInChannels(i).ChanName
            
        Next i
        
    End With

End Sub

Private Sub cmbBoardD_Click()

    Dim i As Long
    
    'Clear the digital input and output channel combo boxes
    Me.cmbChanDIn.Clear
    Me.cmbChanDOut.Clear
    
    'Need to check the System Boards collection
    On Error Resume Next
    
        'if an error happens here, then need to load boards
        If SystemBoards.Count = 0 Then
        
            Exit Sub
            
        End If
        
        If Err.number <> 0 Then
        
            Exit Sub
            
        End If
        
    On Error GoTo 0
    
    'Use the item-data of the current cmbBoardD element as the
    'key to get the intended board
    With SystemBoards(cmbBoardD.List(cmbBoardD.ListIndex))
    
        'Load all of the Digital Output channels into the
        'Digital Output Channel selector combo box
        For i = 1 To .DOutChannels.Count
        
            cmbChanDOut.AddItem (.DOutChannels(i).ChanName)
            
        Next
        
        'Load all of the Digital Input channels into the
        'Digital Input Channel selector combo box
        For i = 1 To .DInChannels.Count
        
            cmbChanDIn.AddItem (.DInChannels(i).ChanName)
            
        Next
        
    End With

End Sub

Private Sub cmdAnalogIn_Click()

    DoAnalogInput

End Sub

Private Sub cmdAnalogOut_Click()

    Dim TempBoard As Board
    Dim TempChan As Channel
    Dim OutVal As Single
    Dim MCC_Counts As Integer
    Dim FixedCounts As Integer
    
    Dim ReturnVal As Long
    
    On Error Resume Next
    
        'Use board name as key to find the board in the System Boards collection
        Set TempBoard = SystemBoards(Trim(cmbBoardA.List(cmbBoardA.ListIndex)))
    
        'error check
        If Err.number <> 0 Then
        
            'Wasn't able to find the correct board, exit function
            
            Exit Sub
            
        End If
        
    On Error GoTo 0
    
    'Read the Output Value
    OutVal = CSng(val(Me.txtEngA))
    
    If TempBoard.CommProtocol = MCC_UL Then
    
        'Calculate the raw MCC counts value
        cbFromEngUnits TempBoard.BoardNum, _
                       TempBoard.range.RangeType, _
                       OutVal, _
                       MCC_Counts
                       
        Me.txtRawA = Trim(Str(MCC_Counts))
        
    ElseIf TempBoard.CommProtocol = ADWIN_COM Then
    
        'Calculate the raw ADWIN counts value
        Me.txtRawA = Trim(Str(TempBoard.range.ADWIN_RangeConverter(CDbl(OutVal))))
                        
    End If
    
    'Now find the right analog output channel in the Board Analog output
    'channels collection
    Set TempChan = TempBoard.AOutChannels(cmbChanAOut.List(cmbChanAOut.ListIndex))
    
    'Output the point
    ReturnVal = TempBoard.AnalogOut(TempChan, OutVal)
    
    'Check for errors
    If ReturnVal = -1 Then
    
        'Set value in txtEngA = ERROR
        Me.txtEngA = "ERROR"
        
    ElseIf ReturnVal <> 0 Then
    
        'Change value in txtEngA
        Me.txtEngA = "ERR: " & Trim(Str(ReturnVal))
        
        frmSendMail.MailNotification "Analog Output Error", _
                                     "Analog Ouput error on board = """ & TempBoard.BoardName & _
                                     """" & vbNewLine & vbNewLine & _
                                     "An error exception has been raised and the code is halted.", _
                                     CodeRed, _
                                     True
    
        'Raise an actual error
        Err.Raise ReturnVal, _
                  "frmDAQ_Comm.cmdAnalogOut", _
                  "Analog Output failed on:" & vbNewLine & _
                  "DAQ Board = " & TempBoard.BoardName & vbNewLine & _
                  "Channel = " & TempChan.ChanName & " (" & _
                  Trim(Str(TempChan.ChanNum)) & ")"
                  
    End If

End Sub

Private Sub cmdClose_Click()

    'Unload the form
    Form_Unload 0
    
    'Hide the form
    Me.Hide

End Sub

Private Sub cmdDigitalIn_Click()

    Dim ReturnVal As Long
    Dim TempBoard As Board
    Dim TempChan As Channel
    
    On Error Resume Next
    
        'Use board name as key to find the board in the System Boards collection
        Set TempBoard = SystemBoards(cmbBoardA.List(cmbBoardA.ListIndex))
    
        'error check
        If Err.number <> 0 Then
        
            'Wasn't able to find the correct board, exit function
            
            Exit Sub
            
        End If
        
    On Error GoTo 0
    
    'Clear the optHighLow values
    optHighLow(0).Value = False
    optHighLow(1).Value = False
    
    'Now find the right analog output channel in the Board Analog output
    'channels collection
    Set TempChan = TempBoard.DInChannels(cmbChanAIn.List(cmbChanAIn.ListIndex))

    'Now call the digital input shell
    ReturnVal = CInt(TempBoard.DigitalInput(TempChan))
    
    'Error Check
    If ReturnVal > 1 Or ReturnVal < 0 Then
    
        'An Error occurred, raise an error
        Err.Raise ReturnVal, _
                  "frmDAQ_Comm.cmdDigitalIn", _
                  "Unable to read from Digital Input port:" & vbNewLine & _
                  "Board = " & TempBoard.BoardName & vbNewLine & _
                  "Channel = " & TempChan.ChanName & " (" & _
                  Trim(Str(TempChan.ChanNum)) & ")"
                  
        Exit Sub
        
    End If
    
    'Pause for 300 ms
    PauseTill timeGetTime() + 300
    
    'UpDate the optHighLow value display
    optHighLow_Click (ReturnVal)
    
End Sub

Private Sub cmdDigitalOut_Click()

    Dim TempBoard As Board
    Dim TempChan As Channel
    Dim ReturnVal As Long
    
    'Check to make sure the system boards collection has the desired board
    On Error Resume Next
    
        'Use board name as key to find the board in the System Boards collection
        Set TempBoard = SystemBoards(cmbBoardD.List(cmbBoardD.ListIndex))
    
        'error check
        If Err.number <> 0 Then
        
            'Wasn't able to find the correct board, exit function
            
            Exit Sub
            
        End If
        
    On Error GoTo 0
    
    'Need to get the user selected digital output channel now
    Set TempChan = TempBoard.DOutChannels(cmbChanDOut.List(cmbChanDOut.ListIndex))
    
    'Now do the actual digital outputing -
    'only changing one digital channel at a time
    ReturnVal = TempBoard.DigitalOutput(TempChan, DigitalBit, True)
    
End Sub

Private Sub DoAnalogInput()

    Dim TempBoard As Board
    Dim TempChan As Channel
    Dim ReturnVal As Variant
    Dim InputVal As Double
    Dim MCC_Counts As Integer
    Dim FixedCounts As Integer
    
    On Error Resume Next
    
        'Use board name as key to find the board in the System Boards collection
        Set TempBoard = SystemBoards(cmbBoardA.List(cmbBoardA.ListIndex))
    
        'error check
        If Err.number <> 0 Then
        
            'Wasn't able to find the correct board, exit function
            
            Exit Sub
            
        End If
        
    On Error GoTo 0
    
       
    'Now find the right analog output channel in the Board Analog output
    'channels collection
    Set TempChan = TempBoard.AInChannels(cmbChanAIn.List(cmbChanAIn.ListIndex))
    
    'Output the point
    ReturnVal = TempBoard.AnalogIn(TempChan)
    
    'Check to see if return value is a string
    If VarType(ReturnVal) = 8 Then
    
        'String return = error occured
        Me.txtEngA = CStr(ReturnVal)
        
        Exit Sub
        
    End If
    
    'ReturnVal must be a double
    InputVal = CDbl(ReturnVal)
    
    'Show in control on the form
    Me.txtEngA = Trim(Str(InputVal))
    
End Sub

Public Function DoDAQIO(ByRef ChanObj As Channel, _
                        Optional ByVal NumValue As Double = "-10000", _
                        Optional ByVal BoolValue As Boolean = False) As Variant

    Dim TempBoard As Board
    Dim ReturnVal As Long
    
    'If NOCOMM_MODE is on, then exit the function and return -616
    If NOCOMM_MODE = True Then
    
        DoDAQIO = -616
        
        Exit Function
        
    End If
    
    'Initialize TempBoard to Nothing
    Set TempBoard = Nothing
    
    If ChanObj Is Nothing Then
    
        Set DoDAQIO = Nothing
        
        Exit Function
        
    End If
        
    'Get the Board object corresponding to this channel
    
    'Turn on error handling
    On Error GoTo NoSysBoards:
    
        Set TempBoard = modConfig.SystemBoards(ChanObj.BoardName)
        
    'Resume normal error flow
    On Error GoTo 0
    
    'Error check a second time
    If TempBoard Is Nothing Then
    
        'Error has occurred when trying to get the board
        'Either the board name stored in the inputed channel object
        'is bad, or the corresponding board has been removed since
        'the channel object was allocated
        Err.Raise Err.number, _
                  "frmDAQ_Comm.DoDAQIO", _
                  "Unable to retrieve Board from System Boards collection." & _
                  vbNewLine & _
                  "Bad board name, or desired board object has gone missing." & _
                  vbNewLine & _
                  vbNewLine & "Board Name used = " & Trim(ChanObj.BoardName)
        
        Set DoDAQIO = Nothing
        
        Exit Function
        
    End If
        
    Select Case ChanObj.ChanType
    
        Case "AI"
        
'------------------------------------------------------------------------------------------------------------------------'
'------------------------------------------------------------------------------------------------------------------------'
'
'   July 2010
'   Isaac Hilburn
'   Code Problem Fix
'
'   DIO and AIO tasks take too long when called by DoDAQIO if we try to update the frmDAQ_Comm display
'   with each IO task.
'
'   Fix:   Call analog / digital IO functions in the Board object classes directly without updating
'          the displays on frmDAQ_Comm
'
'------------------------------------------------------------------------------------------------------------------------'
'
'       Old Code, commented out
'
'------------------------------------------------------------------------------------------------------------------------'
'
'            Set_Board Me.cmbBoardA, _
'                      TempBoard
'
'            cmbBoardA_Click
'
'            Set_Chan Me.cmbChanAIn, _
'                     ChanObj
'
'            cmdAnalogIn_Click
'
'            DoDAQIO = val(Me.txtEngA)
'
'------------------------------------------------------------------------------------------------------------------------'
'
'       New Code
'
'------------------------------------------------------------------------------------------------------------------------'

            DoDAQIO = TempBoard.AnalogIn(ChanObj)
            
            NumValue = DoDAQIO
                       
        Case "AO"
        
'------------------------------------------------------------------------------------------------------------------------'
'------------------------------------------------------------------------------------------------------------------------'
'
'   July 2010
'   Isaac Hilburn
'   Code Problem Fix
'
'   DIO and AIO tasks take too long when called by DoDAQIO if we try to update the frmDAQ_Comm display
'   with each IO task.
'
'   Fix:   Call analog / digital IO functions in the Board object classes directly without updating
'          the displays on frmDAQ_Comm
'
'------------------------------------------------------------------------------------------------------------------------'
'
'       Old Code, commented out
'
'------------------------------------------------------------------------------------------------------------------------'
'
'            Set_Board Me.cmbBoardA, _
'                      TempBoard
'
'            cmbBoardA_Click
'
'            Set_Chan Me.cmbChanAOut, _
'                     ChanObj
'
'            'If NumValue = default (-10000), then user forgot to input a real value
'            If NumValue = -10000 Then
'
'                'Tell the user in a message box that something is up
'                MsgBox "No Analog Voltage given for output over Channel """ & ChanObj.ChanName & _
'                       """ on DAQ Board """ & ChanObj.BoardName & """." & vbNewLine & vbNewLine & _
'                       "Cannot proceed with Analog / Digital output on this channel.", _
'                       vbCritical, _
'                       "Critical User Error!"
'
'                Set DoDAQIO = Nothing
'
'                Exit Function
'
'            End If
'
'            Me.txtEngA = Trim(Str(NumValue))
'
'            cmdAnalogOut_Click
'
'            DoDAQIO = val(Me.txtEngA)
'
'------------------------------------------------------------------------------------------------------------------------'
'
'       New Code
'
'------------------------------------------------------------------------------------------------------------------------'

            'If NumValue = default (-10000), then user forgot to input a real value
            If NumValue = -10000 Then
            
                'Tell the user in a message box that something is up
                Err.Raise -616, _
                          "frmDAQ_Comm.DoDAQIO", _
                          "No Analog Voltage given for output over Channel """ & ChanObj.ChanName & _
                          """ on DAQ Board """ & ChanObj.BoardName & """." & vbNewLine & vbNewLine & _
                          "Cannot proceed with Analog / Digital output on this channel."
                       
                Set DoDAQIO = Nothing
                       
                Exit Function
                
            End If

            'Otherwise, run the Analog Out board function
            DoDAQIO = TempBoard.AnalogOut(ChanObj, NumValue)
                        
        Case "DI"
        
'------------------------------------------------------------------------------------------------------------------------'
'------------------------------------------------------------------------------------------------------------------------'
'
'   July 2010
'   Isaac Hilburn
'   Code Problem Fix
'
'   DIO and AIO tasks take too long when called by DoDAQIO if we try to update the frmDAQ_Comm display
'   with each IO task.
'
'   Fix:   Call analog / digital IO functions in the Board object classes directly without updating
'          the displays on frmDAQ_Comm
'
'------------------------------------------------------------------------------------------------------------------------'
'
'       Old Code, commented out
'
'------------------------------------------------------------------------------------------------------------------------'
'
'            Set_Board Me.cmbBoardD, _
'                      TempBoard
'
'            cmbBoardA_Click
'
'            Set_Chan Me.cmbChanDIn, _
'                     ChanObj
'
'            cmdDigitalIn_Click
'
'            DoDAQIO = Me.optHighLow(1).Value
'
'------------------------------------------------------------------------------------------------------------------------'
'
'       New Code
'
'------------------------------------------------------------------------------------------------------------------------'
        
            ReturnVal = TempBoard.DigitalInput(ChanObj, TempBoard.DoutPortType)
            
            If ReturnVal <> 0 And ReturnVal <> -1 Then
            
                'An error Occurred in the Digital Input process that needs to be called
                Flow_Pause
                
                SetCodeLevel CodeRed
                
                frmSendMail.MailNotification "Digital Input has failed on board = """ & _
                                             Trim(Str(TempBoard.BoardName)) & """." & vbNewLine & _
                                             vbNewLine & _
                                             "Execution has been paused.  Please check the machine.", _
                                             CodeRed, _
                                             True
                                             
                MsgBox "Digital Input has failed on board = """ & _
                        Trim(Str(TempBoard.BoardName)) & """, " & _
                        "and channel = """ & ChanObj.ChanName & """ (" & _
                        Trim(Str(ChanObj.ChanNum)) & vbNewLine & vbNewLine & _
                        "MCC Error #: " & Trim(Str(ReturnVal))
                        
                SetCodeLevel modStatusCode.StatusCodeColorLevelPrior
                
                DoDAQIO = False
                
                Exit Function
                
            End If
            
            'No errors occurred, convert the return value into a boolean
            DoDAQIO = CBool(ReturnVal)
            
        Case "DO"
        
'------------------------------------------------------------------------------------------------------------------------'
'------------------------------------------------------------------------------------------------------------------------'
'
'   July 2010
'   Isaac Hilburn
'   Code Problem Fix
'
'   DIO and AIO tasks take too long when called by DoDAQIO if we try to update the frmDAQ_Comm display
'   with each IO task.
'
'   Fix:   Call analog / digital IO functions in the Board object classes directly without updating
'          the displays on frmDAQ_Comm
'
'------------------------------------------------------------------------------------------------------------------------'
'
'       Old Code, commented out
'
'------------------------------------------------------------------------------------------------------------------------'
'
'            Set_Board Me.cmbBoardD, _
'                      TempBoard
'
'            cmbBoardD_Click
'
'            Set_Chan Me.cmbChanDOut, _
'                     ChanObj
'
'            If BoolValue = True Then
'
'                Me.optHighLow(1).Value = True
'
'            Else
'
'                Me.optHighLow(0).Value = True
'
'            End If
'
'            cmdDigitalOut_Click
'
'            DoDAQIO = BoolValue
'
'------------------------------------------------------------------------------------------------------------------------'
'
'       New Code
'
'------------------------------------------------------------------------------------------------------------------------'
        
            DoDAQIO = TempBoard.DigitalOutput(ChanObj, BoolValue, True)
        
    End Select
        
    Exit Function
    
NoSysBoards:

    'Error Check
    If Err.number <> 0 Then
    
        'Error has occurred when trying to get the board
        'Either the board name stored in the inputed channel object
        'is bad, or the corresponding board has been removed since
        'the channel object was allocated
        Err.Raise Err.number, _
                  "frmDAQ_Comm.DoDAQIO", _
                  "Unable to retrieve Board from System Boards collection." & _
                  vbNewLine & _
                  "Bad board name, or desired board object has gone missing." & _
                  vbNewLine & _
                  vbNewLine & "Board Name used = " & Trim(ChanObj.BoardName)
                  
        Exit Function
        
    End If
        
End Function

Private Sub ErrorHandling()

NoBoardsLoaded:

    'Do nothing for now
    
End Sub

Public Sub Form_Load()

    Dim i As Long
    
    'Set form height and width
    Me.Height = 5025
    Me.Width = 6255
    
    '(March 11, 2011 - I Hilburn)
    'Update all the form's icons with the program Ico file
    SetFormIcon Me
    
    'Need to check the System Boards collection
    On Error Resume Next
    
        'if an error happens here, then need to load boards
        If SystemBoards.Count = 0 Then
            
            Exit Sub
        
        End If
        
        If Err.number <> 0 Then
        
            Exit Sub
            
        End If
        
    On Error GoTo 0
            
    'Initialize TempBoard as a Board
    Set TempBoard = New Board
            
    'Clear All the combo boxes of their prior content
    cmbBoardA.Clear
    cmbBoardD.Clear
    cmbChanAOut.Clear
    cmbChanAIn.Clear
    cmbChanDOut.Clear
    cmbChanDIn.Clear
            
    For i = 1 To SystemBoards.Count
    
        'Load Names of each Board object in the System Boards collection
        'into the item data fields of the each of the Board combo boxes
        cmbBoardA.AddItem SystemBoards(i).BoardName
        cmbBoardD.AddItem SystemBoards(i).BoardName
        
    Next
    
    'Select the first board in the two board combo boxes
    cmbBoardA.ListIndex = 0
    cmbBoardD.ListIndex = 0
    
    'Activate the Board Combo boxes' change events
    cmbBoardA_Click
    cmbBoardD_Click
    
    'Set form load status = True
    FormStatus = True
    
End Sub

Public Sub Form_Unload(Cancel As Integer)
    
    'Set form load status = False (unloaded)
    FormStatus = False

End Sub

Private Sub optHighLow_Click(Index As Integer)

    DigitalBit = CBool(Index)
    
    'Make sure only one coil is selected at a time
    If Index = 0 Then optHighLow(1) = Not optHighLow(0)
    If Index = 1 Then optHighLow(0) = Not optHighLow(1)
    
End Sub

Private Sub Set_Board(ByRef cmbBoard As ComboBox, _
                      ByRef BoardObj As Board)
                      
    Dim i As Long
    Dim N As Long
    
    N = cmbBoard.ListCount
    
    'Check to see if the cmbBoard count is zero
    If N < 1 Then
    
        'No boards have been loaded, crap
        'Raise an error
        Err.Raise -616, _
                  "frmDAQ_Comm.DoDAQIO->frmDAQ_Comm.Set_Board", _
                  "No boards have been loaded into the Board combo-box controls" & _
                  "On form frmSystemBoards." & vbNewLine & _
                  "This should not happen. Check INI file. Also, an error may have " & _
                  "occurred in global System Boards variable assignment."
                  
        Exit Sub
                  
    End If
    
    For i = 1 To N
    
        'Check to see if the item data in this item of the cmbBoard combo-box
        'has a board name = that of the BoardObj
        If BoardObj.BoardName = cmbBoard.List(i - 1) Then
        
            'Set the list-index to the item
            cmbBoard.ListIndex = i - 1
            
            'Set i > N so that this for loop ends
            i = N + 1
            
        End If
        
    Next i
                      
End Sub

Private Sub Set_Chan(ByRef cmbChan As ComboBox, _
                      ByRef ChanObj As Channel)
                      
    Dim i As Long
    Dim N As Long
    
    N = cmbChan.ListCount
    
    'Check to see if the cmbBoard count is zero
    If N < 1 Then
    
        'No boards have been loaded, crap
        'Raise an error
        Err.Raise -616, _
                  "frmDAQ_Comm.DoDAQIO->frmDAQ_Comm.Set_Chan", _
                  "No channels have been loaded into the channel combo-box controls" & _
                  "On form frmSystemBoards." & vbNewLine & _
                  "This should not happen. Check INI file. Also, an error may have " & _
                  "occurred in global System Boards variable assignment."
                  
        Exit Sub
                  
    End If
    
    For i = 1 To N
    
        'Check to see if the item data in this item of the cmbBoard combo-box
        'has a board name = that of the BoardObj
        If ChanObj.ChanName = cmbChan.List(i - 1) Then
        
            'Set the list-index to the item
            cmbChan.ListIndex = i - 1
            
            'Set i > N so that this for loop ends
            i = N + 1
            
        End If
        
    Next i
                      
End Sub

