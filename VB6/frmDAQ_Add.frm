VERSION 5.00
Begin VB.Form frmDAQ_Add 
   Caption         =   "DAQ Board Add/Edit Wizard"
   ClientHeight    =   7755
   ClientLeft      =   11325
   ClientTop       =   2325
   ClientWidth     =   9495
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7755
   ScaleWidth      =   9495
   Begin VB.CheckBox checkAddNew 
      Caption         =   "Add New?"
      Height          =   192
      Left            =   4200
      TabIndex        =   61
      Top             =   7440
      Visible         =   0   'False
      Width           =   1332
   End
   Begin VB.TextBox txtOldBoardName 
      Height          =   288
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   60
      Top             =   7320
      Visible         =   0   'False
      Width           =   612
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   372
      Left            =   6240
      TabIndex        =   27
      Top             =   7320
      Width           =   3132
   End
   Begin VB.CommandButton cmdAddEdit 
      Caption         =   "Add"
      Default         =   -1  'True
      Height          =   372
      Left            =   120
      TabIndex        =   26
      Top             =   7320
      Width           =   2652
   End
   Begin VB.Frame frameBoard 
      Caption         =   "DAQ Board Settings"
      Height          =   7092
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9252
      Begin VB.ComboBox cmbBoardName 
         Height          =   288
         Left            =   1560
         TabIndex        =   62
         Top             =   360
         Width           =   1932
      End
      Begin VB.CommandButton cmdLoadPCIDAS 
         Caption         =   "Load  PCI-DAS6030  board template"
         Height          =   372
         Left            =   4560
         TabIndex        =   25
         Top             =   6360
         Width           =   3612
      End
      Begin VB.CommandButton cmdLoadADWIN 
         Caption         =   "Load  ADWIN-light-16  board template"
         Height          =   372
         Left            =   4560
         TabIndex        =   24
         Top             =   5760
         Width           =   3612
      End
      Begin VB.Frame frameDigitalOutput 
         Caption         =   "Digital Output"
         Height          =   1212
         Left            =   3720
         TabIndex        =   55
         Top             =   4200
         Width           =   5292
         Begin VB.TextBox txtDOutChanNumInc 
            Height          =   288
            Left            =   4200
            TabIndex        =   66
            Top             =   720
            Width           =   855
         End
         Begin VB.TextBox txtDOutFirstChanNum 
            Height          =   288
            Left            =   3000
            TabIndex        =   23
            Top             =   720
            Width           =   732
         End
         Begin VB.TextBox txtDOutChanNamePrefix 
            Height          =   288
            Left            =   1320
            TabIndex        =   22
            Top             =   720
            Width           =   1332
         End
         Begin VB.TextBox txtDOutNumChans 
            Height          =   288
            Left            =   120
            TabIndex        =   21
            Top             =   720
            Width           =   732
         End
         Begin VB.Label Label28 
            Caption         =   "Chan. # Increment :"
            Height          =   492
            Left            =   4200
            TabIndex        =   59
            Top             =   240
            Width           =   852
         End
         Begin VB.Label Label27 
            Caption         =   "1st Chan. #:"
            Height          =   372
            Left            =   3000
            TabIndex        =   58
            Top             =   480
            Width           =   852
         End
         Begin VB.Label Label26 
            Caption         =   "Channel Name Prefix:"
            Height          =   492
            Left            =   1320
            TabIndex        =   57
            Top             =   240
            Width           =   1092
         End
         Begin VB.Label Label25 
            Caption         =   "# of Channels:"
            Height          =   372
            Left            =   120
            TabIndex        =   56
            Top             =   240
            Width           =   852
         End
      End
      Begin VB.Frame frameDigitalInput 
         Caption         =   "Digital Input"
         Height          =   1212
         Left            =   3720
         TabIndex        =   50
         Top             =   2880
         Width           =   5292
         Begin VB.TextBox txtDInChanNumInc 
            Height          =   288
            Left            =   4200
            TabIndex        =   65
            Top             =   720
            Width           =   855
         End
         Begin VB.TextBox txtDInFirstChanNum 
            Height          =   288
            Left            =   3000
            TabIndex        =   20
            Top             =   720
            Width           =   732
         End
         Begin VB.TextBox txtDInChanNamePrefix 
            Height          =   288
            Left            =   1320
            TabIndex        =   19
            Top             =   720
            Width           =   1332
         End
         Begin VB.TextBox txtDInNumChans 
            Height          =   288
            Left            =   120
            TabIndex        =   18
            Top             =   720
            Width           =   732
         End
         Begin VB.Label Label24 
            Caption         =   "Chan. # Increment :"
            Height          =   492
            Left            =   4200
            TabIndex        =   54
            Top             =   240
            Width           =   852
         End
         Begin VB.Label Label23 
            Caption         =   "1st Chan. #:"
            Height          =   372
            Left            =   3000
            TabIndex        =   53
            Top             =   480
            Width           =   852
         End
         Begin VB.Label Label22 
            Caption         =   "Channel Name Prefix:"
            Height          =   492
            Left            =   1320
            TabIndex        =   52
            Top             =   240
            Width           =   1092
         End
         Begin VB.Label Label21 
            Caption         =   "# of Channels:"
            Height          =   372
            Left            =   120
            TabIndex        =   51
            Top             =   240
            Width           =   852
         End
      End
      Begin VB.Frame frameAnalogOutput 
         Caption         =   "Analog Output"
         Height          =   1212
         Left            =   3720
         TabIndex        =   45
         Top             =   1560
         Width           =   5292
         Begin VB.TextBox txtAOutChanNumInc 
            Height          =   288
            Left            =   4200
            TabIndex        =   64
            Top             =   720
            Width           =   855
         End
         Begin VB.TextBox txtAOutNumChans 
            Height          =   288
            Left            =   120
            TabIndex        =   15
            Top             =   720
            Width           =   732
         End
         Begin VB.TextBox txtAOutChanNamePrefix 
            Height          =   288
            Left            =   1320
            TabIndex        =   16
            Top             =   720
            Width           =   1332
         End
         Begin VB.TextBox txtAOutFirstChanNum 
            Height          =   288
            Left            =   3000
            TabIndex        =   17
            Top             =   720
            Width           =   732
         End
         Begin VB.Label Label20 
            Caption         =   "# of Channels:"
            Height          =   372
            Left            =   120
            TabIndex        =   49
            Top             =   240
            Width           =   852
         End
         Begin VB.Label Label19 
            Caption         =   "Channel Name Prefix:"
            Height          =   492
            Left            =   1320
            TabIndex        =   48
            Top             =   240
            Width           =   1092
         End
         Begin VB.Label Label18 
            Caption         =   "1st Chan. #:"
            Height          =   372
            Left            =   3000
            TabIndex        =   47
            Top             =   480
            Width           =   852
         End
         Begin VB.Label Label17 
            Caption         =   "Chan. # Increment :"
            Height          =   492
            Left            =   4200
            TabIndex        =   46
            Top             =   240
            Width           =   852
         End
      End
      Begin VB.Frame frameAnalogInput 
         Caption         =   "Analog Input"
         Height          =   1212
         Left            =   3720
         TabIndex        =   40
         Top             =   240
         Width           =   5292
         Begin VB.TextBox txtAInChanNumInc 
            Height          =   288
            Left            =   4200
            TabIndex        =   63
            Top             =   720
            Width           =   855
         End
         Begin VB.TextBox txtAInFirstChanNum 
            Height          =   288
            Left            =   3000
            TabIndex        =   14
            Top             =   720
            Width           =   732
         End
         Begin VB.TextBox txtAInChanNamePrefix 
            Height          =   288
            Left            =   1320
            TabIndex        =   13
            Top             =   720
            Width           =   1332
         End
         Begin VB.TextBox txtAInNumChans 
            Height          =   288
            Left            =   120
            TabIndex        =   12
            Top             =   720
            Width           =   732
         End
         Begin VB.Label Label16 
            Caption         =   "Chan. # Increment :"
            Height          =   492
            Left            =   4200
            TabIndex        =   44
            Top             =   240
            Width           =   852
         End
         Begin VB.Label Label15 
            Caption         =   "1st Chan. #:"
            Height          =   372
            Left            =   3000
            TabIndex        =   43
            Top             =   480
            Width           =   852
         End
         Begin VB.Label Label14 
            Caption         =   "Channel Name Prefix:"
            Height          =   492
            Left            =   1320
            TabIndex        =   42
            Top             =   240
            Width           =   1092
         End
         Begin VB.Label Label13 
            Caption         =   "# of Channels:"
            Height          =   372
            Left            =   120
            TabIndex        =   41
            Top             =   240
            Width           =   852
         End
      End
      Begin VB.CheckBox checkDIOPreconfig 
         Caption         =   "Check1"
         Height          =   252
         Left            =   1560
         TabIndex        =   11
         Top             =   6600
         Width           =   252
      End
      Begin VB.ComboBox cmbDOutPortType 
         Height          =   288
         Left            =   1560
         TabIndex        =   10
         Top             =   6000
         Width           =   1332
      End
      Begin VB.TextBox txtRangeMin 
         Height          =   288
         Left            =   1560
         TabIndex        =   9
         Top             =   5400
         Width           =   612
      End
      Begin VB.TextBox txtRangeMax 
         Height          =   288
         Left            =   1560
         TabIndex        =   8
         Top             =   4920
         Width           =   612
      End
      Begin VB.ComboBox cmbRangeType 
         Height          =   288
         Left            =   1560
         TabIndex        =   7
         Top             =   4440
         Width           =   1332
      End
      Begin VB.TextBox txtMaxAnalogOutRate 
         Height          =   288
         Left            =   1560
         TabIndex        =   6
         Top             =   3960
         Width           =   1092
      End
      Begin VB.TextBox txtMaxAnalogInRate 
         Height          =   288
         Left            =   1560
         TabIndex        =   5
         Top             =   3360
         Width           =   1092
      End
      Begin VB.ComboBox cmbCommProtocol 
         Height          =   288
         Left            =   1560
         TabIndex        =   3
         Top             =   2160
         Width           =   1332
      End
      Begin VB.ComboBox cmbAInMode 
         Height          =   288
         Left            =   1560
         TabIndex        =   4
         Top             =   2760
         Width           =   1332
      End
      Begin VB.TextBox txtBoardFunction 
         Height          =   612
         Left            =   1560
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   1320
         Width           =   1812
      End
      Begin VB.TextBox txtBoardNum 
         Height          =   288
         Left            =   1560
         TabIndex        =   1
         Top             =   840
         Width           =   612
      End
      Begin VB.Label Label12 
         Caption         =   "Digital IO Preconfigured?:"
         Height          =   495
         Left            =   240
         TabIndex        =   39
         Top             =   6480
         Width           =   1215
      End
      Begin VB.Label Label11 
         Caption         =   "Digital Output Port Type:"
         Height          =   492
         Left            =   240
         TabIndex        =   38
         Top             =   5880
         Width           =   1212
      End
      Begin VB.Label Label10 
         Caption         =   "Range Min (V):"
         Height          =   252
         Left            =   240
         TabIndex        =   37
         Top             =   5400
         Width           =   1212
      End
      Begin VB.Label Label9 
         Caption         =   "Range Max (V):"
         Height          =   252
         Left            =   240
         TabIndex        =   36
         Top             =   4920
         Width           =   1212
      End
      Begin VB.Label Label8 
         Caption         =   "Range Type:"
         Height          =   252
         Left            =   240
         TabIndex        =   35
         Top             =   4440
         Width           =   1212
      End
      Begin VB.Label Label7 
         Caption         =   "Max. Analog Output Rate(Hz):"
         Height          =   492
         Left            =   240
         TabIndex        =   34
         Top             =   3840
         Width           =   1212
      End
      Begin VB.Label Label6 
         Caption         =   "Max. Analog Input Rate(Hz):"
         Height          =   492
         Left            =   240
         TabIndex        =   33
         Top             =   3240
         Width           =   1212
      End
      Begin VB.Label Label5 
         Caption         =   "Comm Protocol:"
         Height          =   252
         Left            =   240
         TabIndex        =   32
         Top             =   2160
         Width           =   1212
      End
      Begin VB.Label Label4 
         Caption         =   "Analog Input Channel Mode:"
         Height          =   372
         Left            =   240
         TabIndex        =   31
         Top             =   2640
         Width           =   1212
      End
      Begin VB.Label Label3 
         Caption         =   "Board Function:"
         Height          =   252
         Left            =   240
         TabIndex        =   30
         Top             =   1320
         Width           =   1212
      End
      Begin VB.Label Label2 
         Caption         =   "Board Num:"
         Height          =   252
         Left            =   240
         TabIndex        =   29
         Top             =   840
         Width           =   1212
      End
      Begin VB.Label Label1 
         Caption         =   "Board Name:"
         Height          =   252
         Left            =   240
         TabIndex        =   28
         Top             =   360
         Width           =   1212
      End
   End
End
Attribute VB_Name = "frmDAQ_Add"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbCommProtocol_Click()

    'If board is an MCC board, need to enable Range Type and DOut Port Type
    If cmbCommProtocol.ItemData(cmbCommProtocol.ListIndex) = "MCC" Then
    
        cmbRangeType.Enabled = True
        cmbDOutPortType.Enabled = True
        
    Else
    
        cmbRangeType.Enabled = False
        cmbDOutPortType.Enabled = False
        
    End If
    
End Sub

Private Sub cmbRangeType_Click()

    Dim RangeObj As Range
    
    Set RangeObj = New Range
    
    'Use the properties of the Range Object to convert the MCC
    'Range Type value into max and min range values
    RangeObj.RangeType = cmbRangeType.ItemData(cmbRangeType.ListIndex)
    
    txtRangeMax = Trim(Str(RangeObj.MaxValue))
    txtRangeMin = Trim(Str(RangeObj.MinValue))
    
    'Deallocate memory
    Set RangeObj = Nothing

End Sub

Private Function ExportToNewBoardObj(Optional ByVal OldININum As Long = -1) As Board

    Dim TempBoard As Board
    Dim TempStr As String
    Dim TempL As Long
    Dim i As Long
    Dim N As Long
    
    'Allocate the new board object
    Set TempBoard = New Board
    
    With TempBoard
        
        'Set the Board Name
        .BoardName = Trim(cmbBoardName.List(cmbBoardName.ListIndex))
        
        'Set the Board #
        .BoardNum = val(txtBoardNum)
        
        'Set the BoardININum
        'Check to see if the user has indicated an old INI number to use
        '(i.e. this new board is an updated copy of an older board object
        ' already in the System Boards collection)
        If OldININum < 0 Then
        
            'No valid new ini number has been inputed, search for the
            'first available new INI board number
            .BoardININum = LocalBoards.GetMax_BoardININum + 1
            
        Else
        
            'A Valid INI number was inputed
            .BoardININum = OldININum
        
        End If
        
        'Set the Board Function
        .BoardFunction = Trim(txtBoardFunction)
        
        'Set the Comm protocol
        .CommProtocol = cmbCommProtocol.ItemData(cmbCommProtocol.ListIndex)
        
        'Set the Board Analog Input mode
        .BoardMode = cmbAInMode.ItemData(cmbAInMode.ListIndex)
        
        'Set the Analog Max Input & Output Rates (in Hz)
        .MaxAInRate = val(txtMaxAnalogInRate)
        .MaxAOutRate = val(txtMaxAnalogOutRate)
        
        'Allocate the Range Object for the board
        Set .Range = New Range
        
        'If this is a MCC_UL board, then set the Range type, else just
        'input the Range Max and Min values
        If .CommProtocol = MCC_UL Then
        
            .Range.RangeType = cmbRangeType.ItemData(cmbRangeType.ListIndex)
            
        Else
        
            .Range.MaxValue = CLng(val(txtRangeMax))
            .Range.MinValue = CLng(val(txtRangeMin))
            
        End If
               
        'If this is a MCC_UL board, then need to set the Digital Output port type
        'Else, do nothing
        If .CommProtocol = MCC_UL Then
        
            .DoutPortType = cmbDOutPortType.ItemData(cmbDOutPortType.ListIndex)
            
        End If
        
        'Set the Digital I/O configuration status
        .DIOConfigured = (checkDIOPreconfig.Value = Checked)
        
        'Now need to create all the channel-type collections
        Set .AInChannels = New Channels
        Set .AOutChannels = New Channels
        Set .DInChannels = New Channels
        Set .DOutChannels = New Channels
        
        'Call function specifically for the allocation of the new channel objects
        ExportChannelsToNewBoardObj .AInChannels, _
                                    "AI", _
                                    CLng(val(txtAInNumChans)), _
                                    Trim(txtAInChanNamePrefix), _
                                    CLng(val(txtAInFirstChanNum)), _
                                    CLng(val(txtAInChanNumInc)), _
                                    TempBoard.BoardName, _
                                    TempBoard.BoardNum, _
                                    TempBoard.BoardININum
        
        ExportChannelsToNewBoardObj .AOutChannels, _
                                    "AO", _
                                    CLng(val(txtAOutNumChans)), _
                                    Trim(txtAOutChanNamePrefix), _
                                    CLng(val(txtAOutFirstChanNum)), _
                                    CLng(val(txtAOutChanNumInc)), _
                                    TempBoard.BoardName, _
                                    TempBoard.BoardNum, _
                                    TempBoard.BoardININum
        
        ExportChannelsToNewBoardObj .DInChannels, _
                                    "DI", _
                                    CLng(val(txtDInNumChans)), _
                                    Trim(txtDInChanNamePrefix), _
                                    CLng(val(txtDInFirstChanNum)), _
                                    CLng(val(txtDInChanNumInc)), _
                                    TempBoard.BoardName, _
                                    TempBoard.BoardNum, _
                                    TempBoard.BoardININum
                                    
        ExportChannelsToNewBoardObj .DOutChannels, _
                                    "DO", _
                                    CLng(val(txtDOutNumChans)), _
                                    Trim(txtDOutChanNamePrefix), _
                                    CLng(val(txtDOutFirstChanNum)), _
                                    CLng(val(txtDOutChanNumInc)), _
                                    TempBoard.BoardName, _
                                    TempBoard.BoardNum, _
                                    TempBoard.BoardININum
        


    End With

    'Return the New Board Object
    Set ExportToNewBoardObj = TempBoard
    
    'Deallocate TempBoard
    Set TempBoard = Nothing

End Function

Private Sub ExportChannelsToNewObj(ByRef ChanCol As Channels, _
                                   ByVal ChanTypeStr As String, _
                                   ByVal NumChannels As Long, _
                                   ByVal ChanPrefix As String, _
                                   ByVal FirstChanNum As Long, _
                                   ByVal ChanInc As Long, _
                                   ByVal BoardName As String, _
                                   ByVal BoardDevNo As Long, _
                                   ByVal BoardININum As Long)

    Dim i As Long
    Dim N As Long
    Dim TempStr As String
    Dim ChanNumb As Long
    Dim FinalNum As Long
    
    'Set N = number of channels to load into the channel collection
    N = NumChannels
    
    If N > 0 Then
        
        'Calculate the Number of the last Channel
        FinalNum = FirstChanNum + (N - 1) * ChanInc
    
        For i = 1 To N
        
            'Calculate the Channel Number starting at the 1st number
            ChanNumb = FirstChanNum + (i - 1) * ChanInc
        
            'Determine the new channels name
            If FinalNumb <= 99 Then
        
                TempStr = ChanPrefix & Format(ChanNumb, "00")
                
            ElseIf FinalNumb <= 999 Then
            
                TempStr = ChanPrefix & Format(ChanNumb, "000")
                
            Else
            
                TempStr = ChanPrefix & Trim(Str(ChanNumb))
                
            End If
                
            'Add the new channel, with the new channel name as the key
            ChanCol.add , TempStr
            
            'Now load up the channel object properties with the correct values
            With ChanCol(TempStr)
            
                .BoardININum = BoardININum
                .BoardName = BoardName
                .ChanName = TempStr
                .ChanNum = ChanNumb
                .ChanType = ChanTypeStr
                
            End With
            
        Next i
        
    Else
    
        'number of channels has been set to Zero!
        'Set the channel collection = Nothing
        Set ChanCol = Nothing
        
    End If

End Sub

Private Sub cmdAddEdit_Click()

    Dim UserResponse As Long
    Dim TempBoard As Board
    
    'Need to load values from this form window to the
    'System Boards Collection directly - this is the only way
    'to do it and preserve all the needed information
    
    'Need to assess which board this belongs to in the System Boards collection
    
    'If this is a new board, need to just add
    'a new boad object to the system boards collection
    If checkAddNew.Value = Checked Then
        
        'Send message to user's
        UserResponse = MsgBox("This will add a new DAQ Board named """ & TempBoard.BoardName & """" & _
                              vbNewLine & " to the Paleomag program's " & _
                              "current DAQ Board settings, but will not affect the .ini file." & _
                              vbNewLine & vbNewLine & _
                              "Are you sure you want to make these changes?", _
                              vbYesNo, _
                              "Warning!")
                          
        'Check for a yes response
        If UserResponse = vbYes Then
                    
            'Load Values on Screen into a new board object
            Set TempBoard = ExportToNewBoardObj
            
            'Turn on Error Handling
            On Error Resume Next
            
                'Attempt Adding board to the system boards collection
                LocalBoards.add TempBoard, TempBoard.BoardName
                
                'Check to see if this was successful
                If Err.number <> 0 Then
                
                    If Err.number = 457 Then
                    
                        'This is the error for adding an object with a key that has already been
                        'used by an object currently in the System Boards collection
                        
                        'Open Message Window to user telling them to change the Board Name
                        MsgBox "Board Name: """ & TempBoard.BoardName & """ is in use by another " & _
                               "DAQ Board already loaded in the System Boards Collection." & _
                               vbNewLine & vbNewLine & "Please use a different Board Name!", , _
                               "Warning!"
                               
                        'Deallocate TempBoard before existing the sub-routine
                        Set TempBoard = Nothing
                               
                        Exit Sub
                               
                    Else
                    
                        Err.Raise Err.number, _
                                  "frmDAQ_Add->cmdAddEdit_Click", _
                                  "Unknown Error when loading Board named """ & TempBoard.BoardName & _
                                  """ into the System Boards collection." & vbNewLine & _
                                  "System Error Message: " & vbNewLine & vbNewLine & _
                                  Err.Description
                                  
                        'Deallocate TempBoard before existing the sub-routine
                        Set TempBoard = Nothing
                                  
                        Exit Sub
                                  
                    End If
                    
                End If
                
            'Turn Off Error Handling
            On Error GoTo 0
            
            'Add was successful
            
        Else
        
            Exit Sub
            
        End If
            
    Else
        
        'This is an existing board
        
        'Check to see if the Board's name is being changed
        TempStr = Trim(cmbBoardName.List(cmbBoardName.ListIndex))
        
        If txtOldBoardName <> TempStr Then
        
            'Crap! Need to delete the old board object from System Boards
            'And add this new one in it's place.
            
            'Get the old board's INI Num
            TempL = LocalBoards(txtOldBoardName).BoardININum
            
            'Generate the new board object to use to replace the old
            Set TempBoard = ExportToNewBoardObj(TempL)
               
            'Check for Changes in the Channel Function Assignment dependencies
            doContinue = CheckBoardDependencies(LocalAssignedChannels, _
                                                LocalBoards(txtOldBoardName), _
                                                TempBoard)
                                                      
            If doContinue = False Then
            
                'User has selected not to go through with the board edit command
                'exit the subroutine
                Exit Sub
                
            End If
            
            'Resolve the Changes to the Board Dependencies
            ResolveBoardDependencies LocalAssignedChannels, _
                                     LocalBoards(txtOldBoardName), _
                                     TempBoard
            
            'Add the updated board, first
                
            'Turn on error handling
            On Error Resume Next
                
                'Attempt Adding board to the system boards collection
                LocalBoards.add TempBoard, TempBoard.BoardName
                
                'Check to see if this was successful
                If Err.number <> 0 Then
                
                    If Err.number = 457 Then
                    
                        'This is the error for adding an object with a key that has already been
                        'used by an object currently in the System Boards collection
                        
                        'Open Message Window to user telling them to change the Board Name
                        MsgBox "Board Name: """ & TempBoard.BoardName & """ is in use by another " & _
                               "DAQ Board already loaded in the System Boards Collection." & _
                               vbNewLine & vbNewLine & "Please use a different Board Name!", , _
                               "Warning!"
                               
                        'Deallocate TempBoard before existing the sub-routine
                        Set TempBoard = Nothing
                               
                        Exit Sub
                               
                    Else
                    
                        Err.Raise Err.number, _
                                  "frmDAQ_Add->cmdAddEdit_Click", _
                                  "Unknown Error when loading Board named """ & TempBoard.BoardName & _
                                  """ into the System Boards collection." & vbNewLine & _
                                  "System Error Message: " & vbNewLine & vbNewLine & _
                                  Err.Description
                                  
                        'Deallocate TempBoard before existing the sub-routine
                        Set TempBoard = Nothing
                                  
                        Exit Sub
                                  
                    End If
                    
                End If
            
            'Turn off error handling
            On Error GoTo 0
            
            'Updated board was added successfully, now delete the old instance
                        
            'Turn on Error Handling
            On Error Resume Next
                
                'Attempt to remove the old instance of the board that's being edited
                LocalBoards.Remove txtOldBoardName
                
                'Error check
                If Err.number <> 0 Then
                
                    'Check for Board not being in the System Boards collection
                    If Err.number = 5 Then
                    
                        'Tell the user what's going on
                        MsgBox "Unable to find and remove a DAQ Board in the System Board's collection with the " & _
                               "name """ & txtOldBoardName & """" & vbNewLine & vbNewLine & _
                               "System Board's collection may be corrupted." & vbNewLine & vbNewLine & _
                               "User Action Needed:" & vbNewLine & _
                               "1. Save the current DAQ Board configuration to the .ini file in the DAQ Board " & _
                               "Settings window. " & vbNewLine & _
                               "   (The old .ini file " & _
                               "will be automatically backed up before the new settings are saved.)" & vbNewLine & _
                               "2. Exit and Restart the Paleomag program." & vbNewLine & vbNewLine & _
                               "If the problem re-occurs, contact the RAPID Software Development Team.", , _
                               "Critical Error!!!"
                               
                        'Deallocate TempBoard before existing the sub-routine
                        Set TempBoard = Nothing
                               
                        Exit Sub
                        
                    Else
                    
                        'Raise the error
                        Err.Rause Err.number, _
                                  "frmDAQ_Add->cmdAddEdit_Click", _
                                  "Unable to remove the DAQ Board named """ & txtOldBoardName & _
                                  """ from the System Board's collection." & vbNewLine & vbNewLine & _
                                  "System Error Message: " & vbNewLine & _
                                  Err.Description
                           
                        'Deallocate TempBoard before existing the sub-routine
                        Set TempBoard = Nothing
                           
                        Exit Sub
                                  
                    End If
                    
                End If
                
            'Turn error handling off
            On Error GoTo 0
            
            'Editing has gone successfully!
            
        Else
        
            'Create new temporary board object with all the changes from the window
            Set TempBoard = ExportToNewBoardObj
            
            'Just need to Edit the existing Board object in the system Boards collection
            EditSystemBoard LocalAssignedChannels, _
                            LocalBoards(txtOldBoardName), _
                            TempBoard
                    
        End If
        
    End If
    
    'Deallocate TempBoard
    Set TempBoard = Nothing

    'Change Status Flag to Show that the Local Boards collection is now different
    'from the global System Boards Collection
    modConfig.LocalAndSystemDifferent = True

    'Now unload and then Load the DAQ Board Settings form
    Unload frmSystemBoardsettings
    Load frmSystemBoardsettings
    
    frmSystemBoardsettings.Show
    Me.Hide
    
End Sub

Private Sub cmdCancel_Click()

    Me.Hide
    Unload Me
    
End Sub

Private Sub cmdLoadADWIN_Click()

    'Load in values for the ADWIN board
    txtBoardName = "ADWIN-light-16"
    txtBoardNum = "1"
    txtBoardFunction = "AF Ramp, AF Monitor, AF/IRM TTL"
    cmbCommProtocol.ListIndex = 1
    
    'Apply the effects of the new comm protocol selection
    cmbCommProtocol_Click
    
    'Continue loading values
    '(Set Analog input mode to Differential Mode)
    cmbAInMode.ListIndex = 1
    
    'Set max analog input and output rates to 50 kHz
    txtMaxAnalogInRate = 50000
    txtMaxAnalogOutRate = 50000
    
    'Range Type is not set - not used for ADWIN boards
    
    'Set Range Min and Max to -10 & 10 Volts
    txtRangeMax = 10
    txtRangeMin = -10
    
    'Digital Output Port Type - not used for ADWIN boards
    
    'Digital IO is already preconfigured as output & input by port binding
    checkDIOPreconfig = Checked
    
    'Load Analog Input Channel preferences
    txtAInNumChans = "8"
    txtAInChanNamePrefix = "ADC-"
    txtAInFirstChanNum = "1"
    cmbAInChanInc.ListIndex = 1
    
    'Load Analog Output Channel preferences
    txtAOutNumChans = "2"
    txtAOutChanNamePrefix = "DAC-"
    txtAOutFirstChanNum = "1"
    cmbAOutChanInc.ListIndex = 0
    
    'Load Digital Input Channel preferences
    txtDInNumChans = "6"
    txtDInChanNamePrefix = "DIGIN-"
    txtDInFirstChanNum = "0"
    cmbDInChanInc.ListIndex = 0
    
    'Load Digital Output Channel preferences
    txtDOutNumChans = "6"
    txtDOutChanNamePrefix = "DIGOUT-"
    txtDOutFirstChanNum = "0"
    cmbDOutChanInc.ListIndex = 0
    
End Sub

Private Sub cmdLoadPCIDAS_Click()

    'Load in values for the MCC PCI-DAS6030 board
    txtBoardName = "PCI-DAS6030"
    txtBoardNum = "0"
    txtBoardFunction = "Alt AF Monitor, IRM Monitor, ARM/IRM Comm"
    cmbCommProtocol.ListIndex = 0
    
    'Apply the effects of the new comm protocol selection
    cmbCommProtocol_Click
    
    'Continue loading values
    '(Set Analog input mode to Single Mode)
    cmbAInMode.ListIndex = 0
    
    'Set max analog input and output rates to 100 kHz
    txtMaxAnalogInRate = 100000
    txtMaxAnalogOutRate = 100000
    
    'Range Type = BIP10VOLTS - only supported range for this MCC board
    cmbRangeType.ListIndex = 0
    
    'Set Range Min and Max by using the range type combo-box change subroutine
    cmbRangeType_Click
    
    'Set the Digital Output Port Type
    cmbDOutPortType = 0   '(Auxport)
    
    'Digital IO is Not preconfigured as output & input by port binding
    'for MCC boards
    checkDIOPreconfig = Unchecked
    
    'Load Analog Input Channel preferences
    txtAInNumChans = "16"
    txtAInChanNamePrefix = "AI-"
    txtAInFirstChanNum = "0"
    cmbAInChanInc.ListIndex = 0
    
    'Load Analog Output Channel preferences
    txtAOutNumChans = "2"
    txtAOutChanNamePrefix = "AO-"
    txtAOutFirstChanNum = "0"
    cmbAOutChanInc.ListIndex = 0
    
    'Load Digital Input Channel preferences
    txtDInNumChans = "8"
    txtDInChanNamePrefix = "DIO-"
    txtDInFirstChanNum = "0"
    cmbDInChanInc.ListIndex = 0
    
    'Load Digital Output Channel preferences
    txtDOutNumChans = "8"
    txtDOutChanNamePrefix = "DIO-"
    txtDOutFirstChanNum = "0"
    cmbDOutChanInc.ListIndex = 0

End Sub

Private Sub Form_Load()

'    'Clear all the combo-boxes
'    Me.cmbAInChanInc.Clear
'    cmbAInMode.Clear
'    cmbAOutChanInc.Clear
'    cmbCommProtocol.Clear
'    cmbDInChanInc.Clear
'    cmbDOutChanInc.Clear
'    cmbRangeType.Clear
'    cmbBoardName.Clear
    
    'Load the Board Name combo-box
    LoadBoardNameCmbBox
    
    'Change Caption on Add/Edit button to "Add"
    cmdAddEdit.Caption = "Add"
    
    'Load the Comm Protocol Combo Box
    cmbCommProtocol.AddItem "MCC", 0
    cmbCommProtocol.ItemData(cmbCommProtocol.NewIndex) = MCC_UL
    cmbCommProtocol.AddItem "ADWIN", 1
    cmbCommProtocol.ItemData(cmbCommProtocol.NewIndex) = ADWIN_COM
    cmbCommProtocol.AddItem "Other", 2
    cmbCommProtocol.ItemData(cmbCommProtocol.NewIndex) = -1
    
    'Load the Analog Input mode combo-box
    cmbAInMode.AddItem "SINGLEMODE", 0
    cmbAInMode.ItemData(cmbAInMode.NewIndex) = SINGLEMODE
    cmbAInMode.AddItem "DIFFERENTIALMODE", 1
    cmbAInMode.ItemData(cmbAInMode.NewIndex) = DIFFERENTIALMODE
        
    'Load the Range Type Combo-box
    cmbRangeType.AddItem "BIP10VOLTS"
    cmbRangeType.ItemData(cmbRangeType.NewIndex) = 1
    cmbRangeType.AddItem "UNI10VOLTS"
    cmbRangeType.ItemData(cmbRangeType.NewIndex) = 100
    cmbRangeType.AddItem "BIP60VOLTS"
    cmbRangeType.ItemData(cmbRangeType.NewIndex) = 20
    cmbRangeType.AddItem "BIP20VOLTS"
    cmbRangeType.ItemData(cmbRangeType.NewIndex) = 15
    cmbRangeType.AddItem "BIP15VOLTS"
    cmbRangeType.ItemData(cmbRangeType.NewIndex) = 21
    cmbRangeType.AddItem "BIP5VOLTS"
    cmbRangeType.ItemData(cmbRangeType.NewIndex) = 0
    cmbRangeType.AddItem "BIP4VOLTS"
    cmbRangeType.ItemData(cmbRangeType.NewIndex) = 16
    cmbRangeType.AddItem "BIP2PT5VOLTS"
    cmbRangeType.ItemData(cmbRangeType.NewIndex) = 2
    cmbRangeType.AddItem "BIP2VOLTS"
    cmbRangeType.ItemData(cmbRangeType.NewIndex) = 14
    cmbRangeType.AddItem "BIP1PT25VOLTS"
    cmbRangeType.ItemData(cmbRangeType.NewIndex) = 3
    cmbRangeType.AddItem "BIP1VOLTS"
    cmbRangeType.ItemData(cmbRangeType.NewIndex) = 4
    cmbRangeType.AddItem "BIPPT625VOLTS"
    cmbRangeType.ItemData(cmbRangeType.NewIndex) = 5
    cmbRangeType.AddItem "BIPPT5VOLTS"
    cmbRangeType.ItemData(cmbRangeType.NewIndex) = 6
    cmbRangeType.AddItem "BIPPT25VOLTS"
    cmbRangeType.ItemData(cmbRangeType.NewIndex) = 12
    cmbRangeType.AddItem "BIPPT2VOLTS"
    cmbRangeType.ItemData(cmbRangeType.NewIndex) = 13
    cmbRangeType.AddItem "BIPPT1VOLTS"
    cmbRangeType.ItemData(cmbRangeType.NewIndex) = 7
    cmbRangeType.AddItem "BIPPT05VOLTS"
    cmbRangeType.ItemData(cmbRangeType.NewIndex) = 8
    cmbRangeType.AddItem "BIPPT01VOLTS"
    cmbRangeType.ItemData(cmbRangeType.NewIndex) = 9
    cmbRangeType.AddItem "BIPPT005VOLTS"
    cmbRangeType.ItemData(cmbRangeType.NewIndex) = 10
    cmbRangeType.AddItem "BIP1PT67VOLTS"
    cmbRangeType.ItemData(cmbRangeType.NewIndex) = 11
    cmbRangeType.AddItem "BIPPT312VOLTS"
    cmbRangeType.ItemData(cmbRangeType.NewIndex) = 17
    cmbRangeType.AddItem "BIPPT156VOLTS"
    cmbRangeType.ItemData(cmbRangeType.NewIndex) = 18
    cmbRangeType.AddItem "BIPPT125VOLTS"
    cmbRangeType.ItemData(cmbRangeType.NewIndex) = 22
    cmbRangeType.AddItem "BIPPT078VOLTS"
    cmbRangeType.ItemData(cmbRangeType.NewIndex) = 19
    cmbRangeType.AddItem "UNI5VOLTS"
    cmbRangeType.ItemData(cmbRangeType.NewIndex) = 101
    cmbRangeType.AddItem "UNI4VOLTS"
    cmbRangeType.ItemData(cmbRangeType.NewIndex) = 114
    cmbRangeType.AddItem "UNI2PT5VOLTS"
    cmbRangeType.ItemData(cmbRangeType.NewIndex) = 102
    cmbRangeType.AddItem "UNI2VOLTS"
    cmbRangeType.ItemData(cmbRangeType.NewIndex) = 103
    cmbRangeType.AddItem "UNI1PT67VOLTS"
    cmbRangeType.ItemData(cmbRangeType.NewIndex) = 109
    cmbRangeType.AddItem "UNI1PT25VOLTS"
    cmbRangeType.ItemData(cmbRangeType.NewIndex) = 104
    cmbRangeType.AddItem "UNI1VOLTS"
    cmbRangeType.ItemData(cmbRangeType.NewIndex) = 105
    cmbRangeType.AddItem "UNIPT5VOLTS"
    cmbRangeType.ItemData(cmbRangeType.NewIndex) = 110
    cmbRangeType.AddItem "UNIPT25VOLTS"
    cmbRangeType.ItemData(cmbRangeType.NewIndex) = 111
    cmbRangeType.AddItem "UNIPT2VOLTS"
    cmbRangeType.ItemData(cmbRangeType.NewIndex) = 112
    cmbRangeType.AddItem "UNIPT1VOLTS"
    cmbRangeType.ItemData(cmbRangeType.NewIndex) = 106
    cmbRangeType.AddItem "UNIPT05VOLTS"
    cmbRangeType.ItemData(cmbRangeType.NewIndex) = 113
    cmbRangeType.AddItem "UNIPT02VOLTS"
    cmbRangeType.ItemData(cmbRangeType.NewIndex) = 108
    cmbRangeType.AddItem "UNIPT01VOLTS"
    cmbRangeType.ItemData(cmbRangeType.NewIndex) = 107
    
    'Load the Digital Output Port Type values to combo-box
    cmbDOutPortType.AddItem "AUXPORT"
    cmbDOutPortType.ItemData(cmbDOutPortType.NewIndex) = 1
    cmbDOutPortType.AddItem "FIRSTPORTA"
    cmbDOutPortType.ItemData(cmbDOutPortType.NewIndex) = 10
    cmbDOutPortType.AddItem "FIRSTPORTB"
    cmbDOutPortType.ItemData(cmbDOutPortType.NewIndex) = 11
    cmbDOutPortType.AddItem "FIRSTPORTCL"
    cmbDOutPortType.ItemData(cmbDOutPortType.NewIndex) = 12
    cmbDOutPortType.AddItem "FIRSTPORTC"
    cmbDOutPortType.ItemData(cmbDOutPortType.NewIndex) = 12
    cmbDOutPortType.AddItem "FIRSTPORTCH"
    cmbDOutPortType.ItemData(cmbDOutPortType.NewIndex) = 13
    cmbDOutPortType.AddItem "SECONDPORTA"
    cmbDOutPortType.ItemData(cmbDOutPortType.NewIndex) = 14
    cmbDOutPortType.AddItem "SECONDPORTB"
    cmbDOutPortType.ItemData(cmbDOutPortType.NewIndex) = 15
    cmbDOutPortType.AddItem "SECONDPORTCL"
    cmbDOutPortType.ItemData(cmbDOutPortType.NewIndex) = 16
    cmbDOutPortType.AddItem "SECONDPORTCH"
    cmbDOutPortType.ItemData(cmbDOutPortType.NewIndex) = 17
    cmbDOutPortType.AddItem "THIRDPORTA"
    cmbDOutPortType.ItemData(cmbDOutPortType.NewIndex) = 18
    cmbDOutPortType.AddItem "THIRDPORTB"
    cmbDOutPortType.ItemData(cmbDOutPortType.NewIndex) = 19
    cmbDOutPortType.AddItem "THIRDPORTCL"
    cmbDOutPortType.ItemData(cmbDOutPortType.NewIndex) = 20
    cmbDOutPortType.AddItem "THIRDPORTCH"
    cmbDOutPortType.ItemData(cmbDOutPortType.NewIndex) = 21
    cmbDOutPortType.AddItem "FOURTHPORTA"
    cmbDOutPortType.ItemData(cmbDOutPortType.NewIndex) = 22
    cmbDOutPortType.AddItem "FOURTHPORTB"
    cmbDOutPortType.ItemData(cmbDOutPortType.NewIndex) = 23
    cmbDOutPortType.AddItem "FOURTHPORTCL"
    cmbDOutPortType.ItemData(cmbDOutPortType.NewIndex) = 24
    cmbDOutPortType.AddItem "FOURTHPORTCH"
    cmbDOutPortType.ItemData(cmbDOutPortType.NewIndex) = 25
    cmbDOutPortType.AddItem "FIFTHPORTA"
    cmbDOutPortType.ItemData(cmbDOutPortType.NewIndex) = 26
    cmbDOutPortType.AddItem "FIFTHPORTB"
    cmbDOutPortType.ItemData(cmbDOutPortType.NewIndex) = 27
    cmbDOutPortType.AddItem "FIFTHPORTCL"
    cmbDOutPortType.ItemData(cmbDOutPortType.NewIndex) = 28
    cmbDOutPortType.AddItem "FIFTHPORTCH"
    cmbDOutPortType.ItemData(cmbDOutPortType.NewIndex) = 29
    cmbDOutPortType.AddItem "SIXTHPORTA"
    cmbDOutPortType.ItemData(cmbDOutPortType.NewIndex) = 30
    cmbDOutPortType.AddItem "SIXTHPORTB"
    cmbDOutPortType.ItemData(cmbDOutPortType.NewIndex) = 31
    cmbDOutPortType.AddItem "SIXTHPORTCL"
    cmbDOutPortType.ItemData(cmbDOutPortType.NewIndex) = 32
    cmbDOutPortType.AddItem "SIXTHPORTCH"
    cmbDOutPortType.ItemData(cmbDOutPortType.NewIndex) = 33
    cmbDOutPortType.AddItem "SEVENTHPORTA"
    cmbDOutPortType.ItemData(cmbDOutPortType.NewIndex) = 34
    cmbDOutPortType.AddItem "SEVENTHPORTB"
    cmbDOutPortType.ItemData(cmbDOutPortType.NewIndex) = 35
    cmbDOutPortType.AddItem "SEVENTHPORTCL"
    cmbDOutPortType.ItemData(cmbDOutPortType.NewIndex) = 36
    cmbDOutPortType.AddItem "SEVENTHPORTCH"
    cmbDOutPortType.ItemData(cmbDOutPortType.NewIndex) = 37
    cmbDOutPortType.AddItem "EIGHTHPORTA"
    cmbDOutPortType.ItemData(cmbDOutPortType.NewIndex) = 38
    cmbDOutPortType.AddItem "EIGHTHPORTB"
    cmbDOutPortType.ItemData(cmbDOutPortType.NewIndex) = 39
    cmbDOutPortType.AddItem "EIGHTHPORTCL"
    cmbDOutPortType.ItemData(cmbDOutPortType.NewIndex) = 40
    cmbDOutPortType.AddItem "EIGHTHPORTCH"
    cmbDOutPortType.ItemData(cmbDOutPortType.NewIndex) = 41
    
       
    'Now check to see if the Add New board option is checked
    If checkAddNew.Value = Unchecked Then
    
        Dim TempBoard As Board
    
        'Initialize temp board = nothing
        Set TempBoard = Nothing
    
        'User has clicked the "Edit Board" button
        'Need to use the board name in the old board name text box to
        'find the correct Board object from the System Boards collection
        
        'Set TempBoard = Board Name in old board name text box
        'Turn on error handling
        On Error Resume Next
        
            Set TempBoard = LocalBoards(txtOldBoardName)
    
            If Err.number <> 0 Then
            
                'Failed to find a matching board object
                'Pop-up an error message and exit this sub-routine
                MsgBox "Unable to find matching Board Object in the System Boards global collection." & _
                       vbNewLine & "Cannot proceed with the Board edit command." & vbNewLine & _
                       "Sorry.", , _
                       "Oops!"
    
                Me.Hide
    
                Exit Sub
                
            End If
            
        'Turn off error handling
        On Error GoTo 0
        
        'Error check a second time
        If TempBoard Is Nothing Then
            
            'Failed to find a matching board object
            'Pop-up an error message and exit this sub-routine
            MsgBox "Unable to find matching Board Object in the System Boards global collection." & _
                   vbNewLine & "Cannot proceed with the Board edit command." & vbNewLine & _
                   "Sorry.", , _
                   "Oops!"
    
            Me.Hide
    
            Exit Sub
            
        End If

        'Change Caption on Add/Edit button to "Edit"
        cmdAddEdit.Caption = "Save Changes"

        'Matching board has been found!
        'Load values from this board object into the fields
        'on this form
        DisplayBoardSettings TempBoard
        
        
    End If

End Sub
Public Function RegExpMatch_FirstInstance(ByVal InputString As String, _
                                     ByVal MatchString As String, _
                                     Optional ByVal isIgnoreCase As Boolean = True, _
                                     Optional ByVal isGlobal As Boolean = False) As Long
                       
    Dim RegularExprObj As New RegExp
    Dim Matches As MatchCollection
    Dim Ma As Match
    Dim TempL As Long
        
    RegularExprObj.Pattern = MatchString
    RegularExprObj.IgnoreCase = isIgnoreCase
    RegularExprObj.Global = isGlobal
    
    If RegularExprObj.test(InputString) = False Then
    
        RegExpMatch_FirstInstance = -1
        
        Exit Function
        
    End If
    
    Set Matches = RegularExprObj.Execute(InputString)
    
    TempL = Len(InputString) + 10
    
    For Each Ma In Matches
    
        If TempL > Ma.FirstIndex Then
        
            TempL = Ma.FirstIndex
            
        End If
        
    Next
    
    RegExpMatch_FirstInstance = TempL
    
    Set RegularExprObj = Nothing
    Set Matches = Nothing
    Set Ma = Nothing
                                                      
End Function
Private Sub DisplayBoardSettings(ByRef BoardObj As Board)

    Dim i As Long
    Dim N As Long
    Dim TempStr As String
    Dim TempL As Long
    Dim ItemFound As Boolean
         
        
    N = cmbBoardName.ListCount
    
    'Set ItemFound to default of false
    ItemFound = False
    
    'Set the Board Name combo-box to the name of the board
    'in the inputed object
    For i = 1 To N
    
        If BoardObj.BoardName = cmbBoardName.List(i - 1) Then
        
            'Make this index in the combo-box the selected index
            cmbBoardName.ListIndex = i - 1
            
            'Set i = N+1 to end the for loop
            i = N + 1
            
            'Set ItemFound = True
            ItemFound = True
            
        End If
        
    Next i
    
    'Check to see if the board name was found in the combo-box
    'if not, add a new element with the board name to the combo-box
    If ItemFound = False Then
    
        cmbBoardName.AddItem BoardObj.BoardName
        
        'Set new index as the selected index
        cmbBoardName.ListIndex = N
        
    End If
    
    'Display the Board Number
    txtBoardNum = Trim(Str(BoardObj.BoardNum))
    
    'Display the Board Function
    txtBoardFunction = Trim(Str(BoardObj.BoardFunction))
    
    'Choose the correct comm-protocol from the combo-box
    N = cmbCommProtocol.ListCount
    
    'Set ItemFound back to false
    ItemFound = False
    
    'Loop through the contents of the comm protocol combo box to find the
    'comm protocol matching that of the board object
    For i = 1 To N
    
        If BoardObj.CommProtocol = cmbCommProtocol.ItemData(i - 1) Then
        
            'Set this list-index as the selected list-index
            cmbCommProtocol.ListIndex = i - 1
            
            'Set i = N + 1 to end the for loop
            i = N + 1
            
            'Set ItemFound = True
            ItemFound = True
            
        End If
        
    Next i
    
    'Check to see if there was a match
    If ItemFound = False Then
    
        'Set the combo-box to the last entry, which will be "Other"
        cmbCommProtocol.ListIndex = N - 1
        
    End If
        
        
    'Choose the correct analog input channel mode from the combo-box
    N = cmbAInMode.ListCount
    
    'Set ItemFound back to false
    ItemFound = False
    
    'Loop through the contents of the Analog input mode combo box to find the
    'mode matching that of the board object
    For i = 1 To N
    
        If BoardObj.BoardMode = cmbAInMode.ItemData(i - 1) Then
        
            'Set this list-index as the selected list-index
            cmbAInMode.ListIndex = i - 1
            
            'Set i = N + 1 to end the for loop
            i = N + 1
            
            'Set ItemFound = True
            ItemFound = True
            
        End If
        
    Next i
    
    'Check to see if there was a match
    If ItemFound = False Then
    
        'Add a new element to the combo-box called "ERROR"
        cmbAInMode.AddItem "ERROR", N
        
        'Set the combo-box to the new "ERROR" entry
        cmbAInMode.ListIndex = N
        
    End If
    
    'Display the Max Analog Input and Output Rates
    txtMaxAnalogInRate = Trim(Str(BoardObj.MaxAInRate))
    txtMaxAnalogOutRate = Trim(Str(BoardObj.MaxAOutRate))
        
    'If Comm Protocol = MCC_UL, then need to select the RangeType of the board in the combo-box
    'Else, the range type combo-box should be disabled
    If BoardObj.CommProtocol = ADWIN_COM Then
    
        'Disable range type combo-box
        cmbRangeType.Enabled = False

    Else
    
        'Enable range type combo-box
        cmbRangeType.Enabled = True
        
        
        'Set ItemFound back to false
        ItemFound = False
        
        'Loop through the contents of the range type combo box to find the
        'range type matching that of the board object
        TempL = BoardObj.Range.RangeType
        
        For i = 1 To N
        
            If TempL = cmbRangeType.ItemData(i - 1) Then
            
                'Set this list-index as the selected list-index
                cmbRangeType.ListIndex = i - 1
                
                'Set i = N + 1 to end the for loop
                i = N + 1
                
                'Set ItemFound = True
                ItemFound = True
                
            End If
            
        Next i
        
        'Check to see if there was a match
        If ItemFound = False Then
        
            'Add a new element to the combo-box called "ERROR"
            cmbRangeType.AddItem "ERROR", N
            
            'Set the combo-box to the new "ERROR" entry
            cmbRangeType.ListIndex = N
            
        End If
        
    End If
    
    'Display the Board's Analog max and min range values
    txtRangeMax = Trim(Str(BoardObj.Range.MaxValue))
    txtRangeMin = Trim(Str(BoardObj.Range.MinValue))

    'Again, only for MCC_UL boards, need to set the digital output port type
    'If Comm Protocol = MCC_UL, then need to select the digital output port type of the board in the combo-box
    'Else, the digital output port type combo-box should be disabled
    If BoardObj.CommProtocol = ADWIN_COM Then
    
        'Disable the Digital output port type combo-box
        cmbDOutPortType.Enabled = False

    Else
    
        'Enable Digital output port type combo-box
        cmbDOutPortType.Enabled = True
                
        'Set ItemFound back to false
        ItemFound = False
        
        'Loop through the contents of the Digital Output port type combo box to find the
        'value matching that of the board object
        TempL = BoardObj.DoutPortType
        
        For i = 1 To N
        
            If TempL = cmbDOutPortType.ItemData(i - 1) Then
            
                'Set this list-index as the selected list-index
                cmbDOutPortType.ListIndex = i - 1
                
                'Set i = N + 1 to end the for loop
                i = N + 1
                
                'Set ItemFound = True
                ItemFound = True
                
            End If
            
        Next i
        
        'Check to see if there was a match
        If ItemFound = False Then
        
            'Add a new element to the combo-box called "ERROR"
            cmbDOutPortType.AddItem "ERROR", N
            
            'Set the combo-box to the new "ERROR" entry
            cmbDOutPortType.ListIndex = N
            
        End If
        
    End If
    
    'Display whether or not the Digital I/O channels are dedicated / pre-set to digital Input
    'or digital output (ADWIN = Yes, PCIDAS = No)
    If BoardObj.DIOConfigured = True Then
    
        checkDIOPreconfig.Value = Checked
        
    Else
    
        checkDIOPreconfig.Value = Unchecked

    End If

    'Display the Analog I/O, and Digital I/O channel numbers
    txtAInNumChans = Trim(Str(BoardObj.AInChannels.Count))
    txtAOutNumChans = Trim(Str(BoardObj.AOutChannels.Count))
    txtDInNumChans = Trim(Str(BoardObj.DInChannels.Count))
    txtDOutNumChans = Trim(Str(BoardObj.DOutChannels.Count))
    
    'Need to get the channel name prefix for the Analog Input channels
    '(the characters to add before the first digit of the channel number)
    TempStr = BoardObj.AInChannels.Item(1).ChanName
    
    'Now use regular expression match shell function to get the position
    'of the first digit in the channel name
    TempL = RegExpMatch_FirstInstance(TempStr, _
                                         "[0-9]")
    
    'If TempL = -1, then no digits are present in the channel name string
    If TempL = -1 Then
    
        'Set Channel Prefix string = "ERROR"
        txtAInChanNamePrefix = "ERROR"
        
    Else
    
        'Set Channel prefix = part of channel name up to the first digit in the string
        txtAInChanNamePrefix = Mid(TempStr, 1, TempL - 1)
        
    End If
    
    
    'Need to get the channel name prefix for the Analog Output channels
    '(the characters to add before the first digit of the channel number)
    TempStr = BoardObj.AOutChannels.Item(1).ChanName
    
    'Now use regular expression match shell function to get the position
    'of the first digit in the channel name
    TempL = RegExpMatch_FirstInstance(TempStr, _
                                         "[0-9]")
    
    'If TempL = -1, then no digits are present in the channel name string
    If TempL = -1 Then
    
        'Set Channel Prefix string = "ERROR"
        txtAOutChanNamePrefix = "ERROR"
        
    Else
    
        'Set Channel prefix = part of channel name up to the first digit in the string
        txtAOutChanNamePrefix = Mid(TempStr, 1, TempL - 1)
        
    End If
    
    
    'Need to get the channel name prefix for the Digital Input Channels
    '(the characters to add before the first digit of the channel number)
    TempStr = BoardObj.DInChannels.Item(1).ChanName
    
    'Now use regular expression match shell function to get the position
    'of the first digit in the channel name
    TempL = RegExpMatch_FirstInstance(TempStr, _
                                         "[0-9]")
    
    'If TempL = -1, then no digits are present in the channel name string
    If TempL = -1 Then
    
        'Set Channel Prefix string = "ERROR"
        txtDInChanNamePrefix = "ERROR"
        
    Else
    
        'Set Channel prefix = part of channel name up to the first digit in the string
        txtDInChanNamePrefix = Mid(TempStr, 1, TempL - 1)
        
    End If
    
    
    'Need to get the channel name prefix for the Digital Output channels
    '(the characters to add before the first digit of the channel number)
    TempStr = BoardObj.DOutChannels.Item(1).ChanName
    
    'Now use regular expression match shell function to get the position
    'of the first digit in the channel name
    TempL = RegExpMatch_FirstInstance(TempStr, _
                                         "[0-9]")
    
    'If TempL = -1, then no digits are present in the channel name string
    If TempL = -1 Then
    
        'Set Channel Prefix string = "ERROR"
        txtDOutChanNamePrefix = "ERROR"
        
    Else
    
        'Set Channel prefix = part of channel name up to the first digit in the string
        txtDOutChanNamePrefix = Mid(TempStr, 1, TempL - 1)
        
    End If
    
    'Now need to get the first channel number for each type of channel
    txtAInFirstChanNum = Trim(Str(BoardObj.AInChannels(1).ChanNum))
    txtAOutFirstChanNum = Trim(Str(BoardObj.AOutChannels(1).ChanNum))
    txtDInFirstChanNum = Trim(Str(BoardObj.DInChannels(1).ChanNum))
    txtDOutFirstChanNum = Trim(Str(BoardObj.DOutChannels(1).ChanNum))
    
    'Now need to get the increment between the channels
    
    'Check to see if there is more than one channel
    If BoardObj.AInChannels.Count > 1 Then
    
        'Display the increment between the first two channels
        'Hoping that this increment is representative of all the channels
        txtAInChanNumInc = Trim(Str(BoardObj.AInChannels(2).ChanNum - _
                                       BoardObj.AInChannels(1).ChanNum))
    
    Else
    
        txtAInChanNumInc = "1"
        
    End If
    
    
    'Check to see if there is more than one channel
    If BoardObj.AOutChannels.Count > 1 Then
    
        'Display the increment between the first two channels
        'Hoping that this increment is representative of all the channels
        txtAOutChanNumInc = Trim(Str(BoardObj.AOutChannels(2).ChanNum - _
                                       BoardObj.AOutChannels(1).ChanNum))
    
    Else
    
        txtAOutChanNumInc = "1"
        
    End If
    
    
    'Check to see if there is more than one channel
    If BoardObj.DInChannels.Count > 1 Then
    
        'Display the increment between the first two channels
        'Hoping that this increment is representative of all the channels
        txtDInChanNumInc = Trim(Str(BoardObj.DInChannels(2).ChanNum - _
                                       BoardObj.DInChannels(1).ChanNum))
    
    Else
    
        txtDInChanNumInc = "1"
        
    End If
    
    
    'Check to see if there is more than one channel
    If BoardObj.DOutChannels.Count > 1 Then
    
        'Display the increment between the first two channels
        'Hoping that this increment is representative of all the channels
        txtDOutChanNumInc = Trim(Str(BoardObj.DOutChannels(2).ChanNum - _
                                       BoardObj.DOutChannels(1).ChanNum))
    
    Else
    
        txtDOutChanNumInc = "1"
        
    End If
        
    'Yay!  The Board settings display process is done!
        
End Sub
Private Sub LoadBoardNameCmbBox()

    Dim i As Long
    Dim N As Long
    Dim ADWIN_found As Boolean
    Dim PCIDAS_found As Boolean
    
    'Navigate through the System Boards Collection
    'And snatch all of the existing board names and load them into the
    'board name combo-box
    On Error Resume Next
    
        'Get the number of board objects in the system boards collection
        N = LocalBoards.Count
        
        'Error check
        If Err.number <> 0 Then
        
            'System Boards has not been loaded yet,
            'make sure the Add New check box is selected
            'and load the standard board names into the combo-box
            
            checkAddNew = Checked
            cmbBoardNameAddItem "ADWIN-light-16", 0
            cmbBoardName.AddItem "PCI-DAS6030", 1
            cmbBoardName.AddItem "Type new name in...", 2
            
            Exit Sub
            
        End If
        
    On Error GoTo 0
        
        
    'Error check a second time
    If N < 1 Then
    
        'No System Boards have been loaded yet,
        'make sure the Add New check box is selected
        'and load the standard board names into the combo-box
        
        checkAddNew = Checked
        cmbBoardName.AddItem "ADWIN-light-16", 0
        cmbBoardName.AddItem "PCI-DAS6030", 1
        cmbBoardName.AddItem "Type new name in...", 2
            
        Exit Sub

    End If
    
    'Set status flags for whether the adwin or pci-das boards have been added yet to
    'false
    ADWIN_found = False
    PCIDAS_found = False
    
    'There are System Boards loaded
    'Add their names to the combo-box
    'and make sure the two default names are also added
    '(ADWIN-light-16 & PCI-DAS6030)
    For i = 1 To N
    
        cmbBoardName.AddItem LocalBoards(i).BoardName, i - 1
        
        If LocalBoards(i).BoardName = "ADWIN-light-16" Then
        
            ADWIN_found = True
            
        End If

        If LocalBoards(i).BoardName = "PCI-DAS6030" Then
        
            PCIDAS_found = True
            
        End If
        
    Next i

    'If the adwin-board name was not present in the LocalBoards collection
    'still need to add it to the combo-box as an option
    If ADWIN_found = False Then
    
        'Iterate number of board names added
        N = N + 1
        
        'Add ADwin board name
        cmbBoardName.AddItem "ADWIN-light-16", N - 1
        
    End If
    
    'Ditto for the PCI-DAS6030 board
    If PCIDAS_found = False Then
    
        'Iterate number of board names added
        N = N + 1
        
        'Add pci-das board name
        cmbBoardName.AddItem "PCI-DAS6030", N - 1
        
    End If
    
    'Now add last entry to combo-box letting the user type in the name
    'for an as yet unknown board
    cmbBoardName.AddItem "Type new name in...", N

End Sub

Private Sub txtAInChanNamePrefix_LostFocus()

    If txtAInChanNamePrefix = "" Then
    
        txtAInChanNamePrefix = "AI"
        
    End If

End Sub

Private Sub txtAInFirstChanNum_LostFocus()

    'Coerce value to be positive
    If val(txtAInFirstChanNum) < 0 Then
    
        txtAInFirstChanNum = Trim(Str(-1 * val(txtAInFirstChanNum)))
        
    End If
    
    'Coerce value to be an integer
    txtAInFirstChanNum = Trim(Str(Int(val(txtAInFirstChanNum))))
    
End Sub

Private Sub txtAOutFirstChanNum_LostFocus()

    'Coerce value to be positive
    If val(txtAOutFirstChanNum) < 0 Then
    
        txtAOutFirstChanNum = Trim(Str(-1 * val(txtAOutFirstChanNum)))
        
    End If
    
    'Coerce value to be an integer
    txtAOutFirstChanNum = Trim(Str(Int(val(txtAOutFirstChanNum))))
    
End Sub

Private Sub txtDInFirstChanNum_LostFocus()

    'Coerce value to be positive
    If val(txtDInFirstChanNum) < 0 Then
    
        txtDInFirstChanNum = Trim(Str(-1 * val(txtDInFirstChanNum)))
        
    End If
    
    'Coerce value to be an integer
    txtDInFirstChanNum = Trim(Str(Int(val(txtDInFirstChanNum))))
    
End Sub

Private Sub txtDOutFirstChanNum_LostFocus()

    'Coerce value to be positive
    If val(txtDOutFirstChanNum) < 0 Then
    
        txtDOutFirstChanNum = Trim(Str(-1 * val(txtDOutFirstChanNum)))
        
    End If
    
    'Coerce value to be an integer
    txtDOutFirstChanNum = Trim(Str(Int(val(txtDOutFirstChanNum))))
    
End Sub

Private Sub txtAOutChanNamePrefix_LostFocus()

    If txtAOutChanNamePrefix = "" Then
    
        txtAOutChanNamePrefix = "AO"
        
    End If

End Sub

Private Sub txtDInChanNamePrefix_LostFocus()

    If txtDInChanNamePrefix = "" Then
    
        txtDInChanNamePrefix = "DI"
        
    End If

End Sub

Private Sub txtDOutChanNamePrefix_LostFocus()

    If txtDOutChanNamePrefix = "" Then
    
        txtDOutChanNamePrefix = "DO"
        
    End If

End Sub

Private Sub txtAInChanNumInc_LostFocus()

    'Coerce value to be an integer
    txtAInChanNumInc = Trim(Str(Int(val(txtAInChanNumInc))))
    
End Sub

Private Sub txtAInNumChans_LostFocus()

    'Coerce Value to be positive
    If val(txtAInNumChans) < 0 Then
    
        txtAInNumChans = Trim(Str(-1 * val(txtAInNumChans)))
               
    End If

    'Coerce value to be an integer
    txtAInNumChans = Trim(Str(Int(val(txtAInNumChans))))

End Sub

Private Sub txtAOutNumChans_LostFocus()

    'Coerce Value to be positive
    If val(txtAOutNumChans) < 0 Then
    
        txtAOutNumChans = Trim(Str(-1 * val(txtAOutNumChans)))
               
    End If
    
    'Coerce value to be an integer
    txtAOutNumChans = Trim(Str(Int(val(txtAOutNumChans))))

End Sub

Private Sub txtDInNumChans_LostFocus()

    'Coerce Value to be positive
    If val(txtDInNumChans) < 0 Then
    
        txtDInNumChans = Trim(Str(-1 * val(txtDInNumChans)))
               
    End If

    'Coerce value to be an integer
    txtDInNumChans = Trim(Str(Int(val(txtDInNumChans))))

End Sub

Private Sub txtDOutNumChans_LostFocus()

    'Coerce Value to be positive
    If val(txtDOutNumChans) < 0 Then
    
        txtDOutNumChans = Trim(Str(-1 * val(txtDOutNumChans)))
               
    End If

    'Coerce value to be an integer
    txtDOutNumChans = Trim(Str(Int(val(txtDOutNumChans))))
    
End Sub

Private Sub txtAOutChanNumInc_LostFocus()

    'Coerce value to be an integer
    txtAOutChanNumInc = Trim(Str(Int(val(txtAOutChanNumInc))))
    
End Sub

Private Sub txtDInChanNumInc_LostFocus()

    'Coerce value to be an integer
    txtDInChanNumInc = Trim(Str(Int(val(txtDInChanNumInc))))
    
End Sub

Private Sub txtDOutChanNumInc_LostFocus()

    'Coerce value to be an integer
    txtDOutChanNumInc = Trim(Str(Int(val(txtDOutChanNumInc))))
    
End Sub

