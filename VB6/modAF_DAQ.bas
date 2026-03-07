Attribute VB_Name = "modAF_DAQ"
Global Const SIGNALGENERATOR = "SIGNALGENERATOR"
Global Const AFRamp = "AFRAMP"
Global Const AFRELAYCONTROL = "AFRELAYCONTROL"
Global Const MONITOR = "MONITOR"
Global Const Pi As Double = 3.14159265358979
Global Const TRIGOFF = -1
Global Const ARM = 16

'Controls the noise tolerance on the ADWIN AF system
'*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*'
'
' IMPORTANT!!!!
'
'If you change this value, you will need to recalibrate
'your AF coils!!!
'
'
'*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*'
Global Const NoiseLevel = 5

Public CoilString As String
Public ReRampThreshold As Double

Public isSystemCoilChange As Boolean


'Last Digital output bit-num assigned to the ADWIN board
Public ADWIN_LastDigOut_BitNum As Long

'Machine Epsilon = machine precision for the double data type
Global Const MachineEpsilon As Double = 2 ^ (-53)

Public Enum Coil_Type

    Axial = 0
    Transverse = 1
    IRMAxial = 2
    IRMTrans = 4
    
End Enum

Public MonitorWave As Wave
Public UpWave As Wave
Public DownWave As Wave

'Internal array for the ArrayInterpolateAndSplice function
Dim ElementsToSave() As Long

Public Declare Function timeBeginPeriod Lib "winmm.dll" (ByVal uPeriod As Long) As Long
Public Declare Function timeEndPeriod Lib "winmm.dll" (ByVal uPeriod As Long) As Long
Public Declare Function timeGetTime Lib "winmm.dll" () As Long

Public Function Atn2(ByVal Y As Double, ByVal X As Double) As Double
    If Abs(X) > CDbl(0.0000000001 * Abs(Y)) Then
        If X < 0 Then
            If Y = 0 Then
                Atn2 = Pi
            Else
                Atn2 = Atn(Y / X) + Pi * Sgn(Y)
            End If
        Else
            Atn2 = Atn(Y / X)
        End If
    Else
        Atn2 = Sgn(Y) * Pi / 2
    End If
End Function

                                  

Public Function CheckBoardDependencies(ByRef AssignedChanCol As Channels, _
                                       ByRef BoardObj As Board, _
                                       ByRef NewBoard As Board) As Boolean

    Dim DependecyList() As String
    Dim N As Long
    Dim i As Long
    
    'Default Dependency List to 1 element = "-1"
    ReDim DependencyList(1) As String
    DependecyList(0) = "-1"

    'Check to see if the user put in a empty board for the new board
    If NewBoard Is Nothing Then
    
        'This function is being called before a board deletion
        'Therefore, there is no new board for the old board to be compared to

        'Get the number of Channels in the Assigned Channel Collection
        N = AssignedChanCol.Count
        
        'Loop through the Assigned channels and search for matches
        'using the Board Name field
        For i = 1 To N
        
            With AssignedChanCol(i)
            
                If .BoardName = BoardObj.BoardName And _
                   .BoardININum = BoardObj.BoardININum _
                Then
                                
                    'Get the current size of the dependency array
                    N = UBound(DependencyList)
                    
                    'Resize the array to fit one more
                    ReDim Preserve DependecyList(N + 1)
                    
                    DependencyList(N) = .ChanDescs.GetAll(", ") & " Channel:" & .ChanName
                    
                End If
            
            End With
            
        Next i
        
        
        'A True Value = continue with delete
        'A False Value = Do NOT delete
        CheckBoardDependencies = DependencyMessage(DependencyList, _
                                                   BoardObj.BoardName)
    
    Else
    
        'The board is being edited, not erased
        'Need to compare old and new board to see which channels
        'have dependencies that won't be affected, or that can be transfered
        GetChannelColDependencies AssignedChanCol, _
                                  BoardObj.AInChannels, _
                                  NewBoard.AInChannels, _
                                  DependencyList
    
        GetChannelColDependencies AssignedChanCol, _
                                  BoardObj.AOutChannels, _
                                  NewBoard.AOutChannels, _
                                  DependencyList
                                  
        GetChannelColDependencies AssignedChanCol, _
                                  BoardObj.DInChannels, _
                                  NewBoard.DInChannels, _
                                  DependencyList
                                  
        GetChannelColDependencies AssignedChanCol, _
                                  BoardObj.DOutChannels, _
                                  NewBoard.DOutChannels, _
                                  DependencyList
                                  
        'A True Value = continue with edit
        'A False Value = Do NOT edit
        CheckBoardDependencies = DependsAndTransfersMessage(DependencyList, _
                                                            BoardObj.BoardName)
                                                   
        
    
    End If
    
End Function

Public Function CheckChannelDependencies(ByRef AssignedChanCol As Channels, _
                                         ByRef ChanObj As Channel, _
                                         ByRef NewChan As Channel) As Boolean
                                        
    Dim DependecyList() As String
    Dim N As Long
    Dim i As Long
    
    'Default Dependency List to 1 element = "-1"
    ReDim DependencyList(1) As String
    DependecyList(0) = "-1"
                                                                                               
    'Check the dependencies of the Assigned DAQ channel objects
    'Get the number of channels in the inputed Assigned Channels collection
    N = AssignedChanCol.Count
    
    'Loop through the assigned channels and check for matches
    'to the inputed channel object
    '(using the BoardName, ChanName, and ChanType fields)
    For i = 1 To N
    
        With AssignedChanCol(i)
        
            If .BoardName = ChanObj.BoardName And _
               .BoardININum = ChanObj.BoardININum And _
               .ChanName = ChanObj.ChanName And _
               .ChanType = ChanObj.ChanType _
            Then
            
                'Get current size the the dependency array
                N = UBound(DependencyList)
            
                'Resize the array to fit one more
                ReDim Preserve DependecyList(N + 1)
                
                DependencyList(0) = .ChanDescs & " Channel: " & .ChanName
                
            End If
            
        End With
        
    Next i
    
    'A True Value = continue with delete
    'A False Value = Do NOT delete
    Dim blah(10) As String
    CheckChannelDependencies = DependencyMessage(DependencyList, ChanObj.BoardName, ChanObj.ChanName)
End Function

Public Sub ConcatenateChanCol(ByRef BoardObj As Board, _
                              ByRef ChanCol As Channels)
                              
    Dim N As Long
    Dim i As Long
    
    'Swipe the ChanCol (Set it to Nothing)
    Set ChanCol = Nothing
    Set ChanCol = New Channels
    
    'Get the # of channels in the Analog Input Channel Collection
    N = BoardObj.AInChannels.Count
    
    'Add the Analog Input channels to the new Channel collection being created
    If N > 0 Then
        
        'Loop through and add each channel to ChanCol
        For i = 1 To N
        
            With BoardObj.AInChannels
            
                ChanCol.Add .Item(i), .Item(i).ChanName
                
            End With
            
        Next i
        
    End If
    
    
    'Get the # of channels in the Analog Output Channel Collection
    N = BoardObj.AOutChannels.Count
    
    'Add the Analog Output channels to the new Channel collection being created
    If N > 0 Then
        
        'Loop through and add each channel to ChanCol
        For i = 1 To N
        
            With BoardObj.AOutChannels
            
                ChanCol.Add .Item(i), .Item(i).ChanName
                
            End With
            
        Next i
        
    End If
                              
                              
    'Get the # of channels in the Digital Input Channel Collection
    N = BoardObj.DInChannels.Count
    
    'Add the Digital Input channels to the new Channel collection being created
    If N > 0 Then
        
        'Loop through and add each channel to ChanCol
        For i = 1 To N
        
            With BoardObj.DInChannels
            
                ChanCol.Add .Item(i), .Item(i).ChanName
                
            End With
            
        Next i
        
    End If
    
                              
    'Get the # of channels in the Digital Output Channel Collection
    N = BoardObj.DOutChannels.Count
    
    'Add the Digital Output channels to the new Channel collection being created
    If N > 0 Then
        
        'Loop through and add each channel to ChanCol
        For i = 1 To N
        
            With BoardObj.DOutChannels
            
                ChanCol.Add .Item(i), .Item(i).ChanName
                
            End With
            
        Next i
        
    End If
                              
End Sub

Public Function CustomDialog(ByVal MessageStr As String, _
                             ByVal PosBtnCaption As String, _
                             Optional ByVal NegBtnCaption As String = "-1", _
                             Optional ByVal CancelBtnCaption As String = "-1") As String

    Load frmDialog
    
    With frmDialog
    
        .txtMessage.text = TempStr
         
        If NegBtnCaption = "-1" And _
           CancelBtnCaption = "-1" _
        Then
         
            frmDialog.NumButtons = 1
            
        ElseIf NegBtnCaption <> "-1" And _
               CancelBtnCaption = "-1" _
        Then
            
            frmDialog.NumButtons = 2
            
        ElseIf NegBtnCaption <> "-1" And _
               CancelBtnCaption = "-1" _
        Then
        
            frmDialog.NumButtons = 3
        
        End If
        
        .cmdPositive.Caption = PosBtnCaption
        .cmdNegative.Caption = NegBtnCaption
        .cmdCancel.Caption = CancelBtnCaption
            
    End With
    
    frmDialog.ReSizeWindow
    
    frmDialog.Show
    
    'Go into a loop waiting for the user to click a button on frmDialog
    Do
    
        DoEvents
        
        PauseTill timeGetTime() + 100
        
    Loop Until frmDialog.UserResponse <> "-1"
        
    CustomDialog = frmDialog.UserResponse
    
    Unload frmDialog
    frmDialog.Hide
    
End Function

Public Function DeleteChanFromColl(ByRef AssignedChanCol As Channels, _
                                   ByRef BoardsCol As Boards, _
                                   ByVal BoardName As String, _
                                   ByVal ChanType As String, _
                                   ByVal ChanName As String) As Boolean

    Dim TempBoard As Board
    Dim TempChannels As Channels
    Dim doContinue As Boolean
    
    'Initialize TempBoard, and TempChannels to Nothing
    Set TempBoard = Nothing
    Set TempChannels = Nothing
    
    'Get TempBoard from BoardsCol
    Set TempBoard = BoardsCol(BoardName)

    'Get the corresponding channel objects
    Select Case ChanType
    
        Case "AI"
        
            'Channel is in the Analog Input channels collection
            'Snatch it!
            Set TempChannels = TempBoard.AInChannels
        
        Case "AO"
        
            'Channel is in the Analog Output channels collection
            'Snatch it!
            Set TempChannels = TempBoard.AOutChannels
        
        Case "DI"
        
            'Channel is in the Digital Input channels collection
            'Snatch it!
            Set TempChannels = TempBoard.DInChannels
        
        Case "DO"
        
            'Channel is in the Digital Output channels collection
            'Snatch it!
            Set TempChannels = TempBoard.DOutChannels
        
    End Select
    
    'Check for any channel assignments for this channel
    doContinue = CheckChannelDependencies(AssignedChanCol, TempChannels(ChanName), Nothing)
    
    If doContinue = False Then
    
        'User has chosen not to continue with the channel delete
        DeleteChanFromColl = False
        Exit Function
        
    End If
    
    'Resolve the Channel Dependency changes
    ResolveChannelDependencies AssignedChanCol, _
                               TempChannels(ChanName), _
                               Nothing
    
    'Remove Channel from the Temp Channels Collection
    TempChannels.Remove ChanName
    
    'Now Save TempChannels back to the correct local Board & local Chan collection
    'To change the Local Boards collection
    Select Case ChanType
    
        Case "AI"
        
            Set BoardsCol(BoardName).AInChannels = TempChannels
        
        Case "AO"
        
            Set BoardsCol(BoardName).AOutChannels = TempChannels
        
        Case "DI"
        
            Set BoardsCol(BoardName).DInChannels = TempChannels
        
        Case "DO"
    
            Set BoardsCol(BoardName).DOutChannels = TempChannels
    
    End Select
    
    'Change the status flag to indicate that the Local Boards Collection
    'is now different from the System Boards Collection
    modConfig.LocalAndSystemDifferent = True
    
    'Return True - channel has been deleted
    DeleteChanFromColl = True
    
    'Deallocate Temp Board and Channels objects
    Set TempBoard = Nothing
    Set TempChannels = Nothing

End Function

Public Sub DeleteCol(ByRef gridobj As MSHFlexGrid, _
                     ByRef CallingForm As Form, _
                     ByVal FirstCol As Long, _
                     Optional ByVal LastCol As Long = -1)
                          
    'Temp array for sorting the gridObj values
    Dim TempArray() As String
    Dim N As Long
    Dim M As Long
    Dim i As Long
    Dim j As Long
    Dim colMod As Long
    Dim FixedCols As Long
    Dim FixedRows As Long
    Dim NumDelcols As Long
    
    With gridobj
    
        'Store the number of cols and rows in the unaltered grid
        N = .Cols
        M = .Rows
        
        'Store the number of fixed cols and fixed rows
        FixedCols = .FixedCols
        FixedRows = .FixedRows
    
    End With
    
    'Load the entire unaltered grid into the array
    SaveGridToArray gridobj, _
                    TempArray(), _
                    0, _
                    N - 1, _
                    0, _
                    M - 1
                    
    'Now check to see if the last col = -1
    If LastCol <> -1 Then
    
        'There could be more than one col that needs to be deleted
        
        'if lastcol is smaller than the first col, exchange the
        'values of the two
        If LastCol < FirstCol Then
        
            i = FirstCol
            FirstCol = LastCol
            LastCol = i
        
        End If
        
        'Determine the number of cols to delete.  Remember that this
        'is an inclusive value
        NumDelcols = LastCol - FirstCol + 1
                                        
    Else
    
        'No last col inputed, therefore there is only one
        'col to delete
        NumDelcols = 1
        
    End If
                    
    With gridobj
    
        'Clear the contents and structure of the current Grid
        .Clear
        .ClearStructure
        
        'Rebuild the grid structure
        
        'Set the number of cols to the number that are remaining
        'after the desired cols are deleted
        .Cols = N - NumDelcols
        
        'Set the rows (were aren't deleting any rowumns right now
        .row = M
        
        'Start the col Index modifier out at zero.  This is
        'used so that we can access the values in the TempArray
        'even though we are skipping some of the cols in the matrix
        'stored in this array variable
        colMod = 0
        
        'Now iterate through the cols and rows and repopulate
        'the grid
        For i = 0 To N - 1
        
            'Only load data into the col if it is not
            'one of the cols being deleted
            If i < FirstCol Or _
               i > LastCol _
            Then
            
                'Load the corresponding rowumn elements
                'from the data array to the grid object cells
                For j = 0 To .Rows - 1
                
                  .TextMatrix(i + colMod, j) = TempArray(i, j)
                  
                Next j
                
            Else
            
                'This col has been deleted, col mod needs to be
                'decremented by one
                colMod = colMod - 1
                
            End If
            
        Next i
        
    End With
                    
End Sub

   
Public Sub DeleteRow(ByRef gridobj As MSHFlexGrid, _
                          ByRef CallingForm As Form, _
                          ByVal FirstRow As Long, _
                          Optional ByVal LastRow As Long = -1, _
                          Optional ByVal FirstColNumbered As Boolean = False, _
                          Optional ByVal doResizeGrid As Boolean = False)
                          
    'Temp array for sorting the gridObj values
    Dim TempArray() As String
    Dim N As Long
    Dim M As Long
    Dim i As Long
    Dim j As Long
    Dim RowMod As Long
    Dim FixedRows As Long
    Dim FixedCols As Long
    Dim NumDelRows As Long
    
    With gridobj
    
        'Store the number of Rows and Cols in the unaltered grid
        N = .Rows
        M = .Cols
        
        'Store the number of fixed Rows and fixed Cols
        FixedRows = .FixedRows
        FixedCols = .FixedCols
    
    End With
    
    'Load the entire unaltered grid into the array
    SaveGridToArray gridobj, _
                    TempArray(), _
                    0, _
                    N - 1, _
                    0, _
                    M - 1
                    
    'Now check to see if the last Row = -1
    If LastRow <> -1 Then
    
        'There could be more than one Row that needs to be deleted
        
        'if lastRow is smaller than the first Row, exchange the
        'values of the two
        If LastRow < FirstRow Then
        
            i = FirstRow
            FirstRow = LastRow
            LastRow = i
        
        End If
        
        'Determine the number of Rows to delete.  Remember that this
        'is an inclusive value
        NumDelRows = LastRow - FirstRow + 1
                                        
    Else
    
        'No last Row inputed, therefore there is only one
        'Row to delete
        NumDelRows = 1
        
    End If
                    
    With gridobj
    
        'Clear the contents and structure of the current Grid
        .Clear
        .ClearStructure
        
        'Rebuild the grid structure
        
        'Set the number of Rows to the number that are remaining
        'after the desired Rows are deleted
        .Rows = N - NumDelRows
        
        'Set the Rows (were aren't deleting any Rowumns right now
        .Cols = M
        
        'Start the Row Index modifier out at zero.  This is
        'used so that we can access the values in the TempArray
        'even though we are skipping some of the Rows in the matrix
        'stored in this array variable
        RowMod = 0
        
        'Now iterate through the Rows and Cols and repopulate
        'the grid
        For i = 0 To N - 1
        
            'Only load data into the Row if it is not
            'one of the Rows being deleted
            If i < FirstRow Or _
               i > LastRow _
            Then
            
                'Load the corresponding Rowumn elements
                'from the data array to the grid object cells
                For j = 0 To .Cols - 1
                
                  .TextMatrix(i + RowMod, j) = TempArray(i, j)
                  
                Next j
                
            Else
            
                'This Row has been deleted, Row mod needs to be
                'decremented by one
                RowMod = RowMod - 1
                
            End If
            
        Next i
            
        'Renumber first Column if so indicated
        If FirstColNumbered = True Then
        
            RenumberGridCol gridobj, _
                            CallingForm, _
                            0
                            
        End If
        
        'Resize the grid Columns if so inputed
        If doResizeGrid = True Then
        
            ResizeGrid gridobj, _
                       CallingForm
                       
        End If
            
    End With
                    
End Sub

Public Function DependencyMessage(ByRef DependList() As String, _
                                  ByVal BoardName As String, _
                                  Optional ByVal ChanName As String = "-1") As Boolean
                             
    Dim N As Long
    Dim StartIndex As Long
    Dim i As Long
    Dim TempStr As String
    Dim UserResp As Long
    
    'Get the length of the dependency string array
    N = UBound(DependList)
    
    'If the first element = "-1", set StartIndex = 1
    'Else, Set StartIndex = 0
    If DependList(0) = "-1" Then
    
        StartIndex = 1
        
    Else
    
        StartIndex = 0
        
    End If
    
    'Check to see if Start Index = N, (i.e. no actual dependencies found)
    If StartIndex = N Then
    
        DependencyMessage = True
        
        Exit Function
        
    End If
    
    'Default TempStr
    TempStr = ""
    
    'There are dependencies, check to see if the there
    'are for a board, or for a board/channel combination
    If ChanName = "-1" Then
    
        'This is for a board only
        TempStr = "Board: """ & BoardName & """ has " & Trim(str(N - StartIndex)) & _
                  " AF / Rock-mag channel assignments bound to it." & vbNewLine & vbNewLine & _
                  "Do you still want to delete this board?" & vbNewLine & vbNewLine & _
                  "List of Channel Assignments: "
                  
    Else
    
        'This is for a particular channel on a particular board
        TempStr = "Channel: """ & ChanName & """ on Board: """ & BoardName & """" & _
                  " has " & Trim(str(N - StartIndex)) & " AF / Rock-mag channel " & _
                  "assignments bound to it." & vbNewLine & vbNewLine & _
                  "Do you still want to delete this channel?" & vbNewLine & vbNewLine & _
                  "List of Channel Assignments: "
                  
    End If

    'need to concatenate dependency strings
    For i = StartIndex To N - 1
    
        'Concentate the string + a breakline and a line numbering
        TempStr = TempStr & vbNewLine & vbNewLine & vbTab & Trim(str(i)) & "." & vbTab & DependList(i)
        
    Next i
    
    'Custom Dialog creates a variable size message-box like dialog with up to three buttons
    'with customizable captions
    'RespStr is a string = caption of the button that was clicked
    RespStr = CustomDialog(TempStr, "Yes", "No")
    
    'Check to see if the user answered "Yes"
    If RespStr = "Yes" Then
    
        'User Does want to go through with the delete
        DependencyMessage = True
        
    Else
    
        DependencyMessage = False
        
    End If
    
End Function

Public Function DependsAndTransfersMessage(ByRef DependList() As String, _
                                           ByVal BoardName As String, _
                                           Optional ByVal ChanName As String = "-1") As Boolean
                             
    Dim N As Long
    Dim StartIndex As Long
    Dim i As Long
    Dim TempStr As String
    Dim TempStr2 As String
    Dim UserResp As Long
    Dim TempL As Long
    Dim TempL2 As Long
    
    'Get the length of the dependency string array
    N = UBound(DependList)
    
    'Check to see if Start Index = N, (i.e. no actual dependencies found)
    If StartIndex = 1 Then
    
        DependsAndTransfersMessage = True
        
        Exit Function
        
    End If
    
    'Default TempStr, TempStr2 = ""
    TempStr = ""
    TempStr2 = ""
    
    'There are dependencies, check to see if the there
    'are for a board, or for a board/channel combination
    If ChanName = "-1" Then
    
        'This is for a board only
        TempStr = "DAQ Board: """ & BoardName & """ has " & Trim(str(N - StartIndex)) & _
                  " AF / Rock-mag channel assignments bound to it that will become " & _
                  "channel-less in the DAQ Board edit process" & vbNewLine & vbNewLine & _
                  "Do you still want to edit this board?" & vbNewLine & vbNewLine & _
                  "List of Channel Assignments that Will be Lost: "
                  
    Else
    
        'This is for a particular channel on a particular board
        TempStr = "Channel: """ & ChanName & """ on DAQ Board: """ & BoardName & """" & _
                  " has " & Trim(str(N - StartIndex)) & " AF / Rock-mag channel " & _
                  "assignments bound to it that will be become channel-less in " & _
                  "the channel edit process." & vbNewLine & vbNewLine & _
                  "Do you still want to edit this channel?" & vbNewLine & vbNewLine & _
                  "List of Channel Assignments that Will be Lost: "
                  
    End If

    'need to concatenate dependency strings
    For i = StartIndex To N - 1
    
        'Get the position of the marker for the start
        'of info about a transferable dependency ("[{")
        'and also the length of the list entry - 2
        '(clip off two characters for the marker
        'for the end of transfers dependence info - "}]")
        TempL = InStrRev(DependList(i), "[{")
        TempL2 = Len(DependList(i)) - 2
        
        If TempL > 0 Then
        
            'This is an entry for transferable dependency
            'Need to parse List into it's two parts
            TempStr2 = TempStr2 & vbNewLine & vbNewLine & "  " & Trim(str(i)) & "." & vbTab & _
                       Mid(DependList(i), 1, TempL - 2) & vbNewLine & _
                       Mid(DependList(i), TempL + 1, TempL2 - TempL - 1)
    
        End If
    
        'Concatenate the string + a breakline and a line numbering
        TempStr = TempStr & vbNewLine & vbNewLine & "  " & Trim(str(i)) & "." & vbTab & DependList(i)
        
    Next i
    
    TempStr = TempStr & vbNewLine & vbNewLine & _
              "List of Channel Assignments that can be transfered to the new Board/Channel:" & _
              TempStr2
              
    
    'Custom Dialog creates a variable size message-box like dialog with up to three buttons
    'with customizable captions
    'RespStr is a string = caption of the button that was clicked
    RespStr = CustomDialog(TempStr, "Yes", "No")
    
    'Check to see if the user answered "Yes"
    If RespStr = "Yes" Then
    
        'User Does want to go through with the delete
        DependsAndTransfersMessage = True
        
    Else
    
        DependsAndTransfersMessage = False
        
    End If
    
End Function

Public Function EditChanInColl(ByRefBoardsCol As Boards, _
                               ByRef NewChan As Channel, _
                               ByVal OldChanName As String, _
                               ByVal OldChanType As String) As Boolean

End Function

Public Sub EditSystemBoard(ByRef AssignedChanCol As Channels, _
                           ByRef SysBoard As Board, _
                           ByRef NB As Board)
                            
                            
    'Leave the System Board Name and the Board INI Number alone,
    'otherwise, change all the other settings
    With SysBoard
    
        .BoardNum = NB.BoardNum
        .BoardFunction = NB.BoardFunction
        .CommProtocol = NB.CommProtocol
        .BoardMode = NB.BoardMode
        .MaxAInRate = NB.MaxAInRate
        .MaxAOutRate = NB.MaxAOutRate
        Set .range = NB.range
        .DIOConfigured = NB.DIOConfigured
        
        If .CommProtocol = MCC_UL Then
        
            .DoutPortType = NB.DoutPortType
            
        End If
        
        'Set the old channel collections to the new collections
        'Need to check to see if function assignments for a channel will be broken by a
        'Channel Name Prefix Change
        EditSystemChanCol AssignedChanCol, _
                          .AInChannels, _
                          NB.AInChannels
        
        EditSystemChanCol AssignedChanCol, _
                          .AOutChannels, _
                          NB.AOutChannels
        
        EditSystemChanCol AssignedChanCol, _
                          .DInChannels, _
                          NB.DInChannels
        
        EditSystemChanCol AssignedChanCol, _
                          .DOutChannels, _
                          NB.DOutChannels
                
    End With
                                    
End Sub

Public Sub EditSystemChanCol(ByRef AssignedChanCol As Channels, _
                             ByRef SysChanCol As Channels, _
                             ByRef NC As Channels)
                              
    Dim N As Long
    Dim M As Long
    Dim i As Long
    Dim j As Long
    
    'Set N = # channels in the SysChanCol
    N = SysChanCol.Count
    
    'Set M = # channels in the New Channel collection
    M = NC.Count
    
    If N <= 0 Then
    
        'No channel dependencies to worry about,
        'exit the sub-routine
        Exit Sub
        
    End If
    
    If N = M Then
    
        'The number of channels is the same
        'Make an educated guess that the channel dependencies
        'need to be transfered if the channel names differ
        
        'Loop through the channels,
        'if the channel types match
        'call the TransferChannelDependencies function
        'to change the Assigned Channels' info to the new assignments
        For i = 1 To N
        
            ResolveChannelDependencies AssignedChanCol, _
                                       SysChanCol, _
                                       NC
                                                   
        Next i

    Else
    
        'ARRRRRRGH!
        'Need to do a matching
        'With each channel in the SysChanCol being compared with
        'each channel in the New Channels collection
        
        'A Match is when both Board names, Channel names, and Channel types are identical
        'Where there is no match, call ResolveChannelDependencies to
        'resolve orphaned function assignments
        If M > 0 Then
            
            For i = 1 To N
                
                For j = 1 To M
                
                    If SysChanCol(i).BoardName = NC(j).BoardName And _
                       SysChanCol(i).ChanName = NC(j).ChanName And _
                       SysChanCol(i).ChanType = NC(j).ChanType _
                    Then
                
                        'Match found, end the inner for loop
                        'by setting j > M
                        j = M + 10
                        
                    End If
                    
                Next j
                
                'Check to see if a match was found
                If j <> M + 10 Then
                
                    'No matching Chan obj in the new board,
                    'so the old Chan obj will be effectively deleted
                    ResolveChannelDependencies AssignedChanCol, _
                                               SysChanCol(i), _
                                               Nothing
                    
                End If
                
            Next i
            
        Else
        
            'M = 0 - There are no channels in the new channel collection
            'need to resolve every channel's dependencies
            For i = 1 To N
            
                ResolveChannelDependencies AssignedChanCol, _
                                           SysChanCol, _
                                           Nothing
                                           
            Next i
            
        End If
            
    End If
    
    'Set The old collection = new collection
    Set SysChanCol = NC
        
End Sub

Public Function ElementAlreadySaved(ByVal NewOrOld As Long, _
                                    ByVal CheckElement As Long) As Boolean
    
    Dim i As Long
    Dim N As Long
    
    N = UBound(ElementsToSave, 1)
        
    For i = 0 To N - 1

        If ElementsToSave(i, NewOrOld) = CheckElement Then
            
            ElementAlreadySaved = True
            
            Exit Function

        End If
        
    Next i
    
    'Element not found
    ElementAlreadySaved = False
    
End Function

Public Function FindFieldFromVolts(ByVal MonVolt As Double, _
                                   Optional ByVal AFCoilSystem As Long = -128) As Double
                                   
    '(July 2010, I Hilburn)
    '
    'This function does the inverse of FindXCalibValue.
    'It takes a voltage or 2G counts input and returns a double value containing
    'the corresponding Field value.
    
    Dim i As Long
    Dim Slope As Double
    
    'If the user has not inputed a coil system, set the coil system
    'to the global active coil system
    If AFCoilSystem = -128 Then AFCoilSystem = ActiveCoilSystem
    
    'Check for MonVolt <= 0
    If MonVolt <= 0 Then
    
        'Return Zero
        FindFieldFromVolts = 0
        
        'Exit the function
        Exit Function
        
    End If
    
    If AFCoilSystem = AxialCoilSystem Then
    'Axial Coil
        
        'Check to make sure AFAxialCount > 1
        If AFAxialCount <= 1 Then
        
            'User hasn't entered in enough calibration values
            MsgBox "Only one AF Axial Coil calibration value has been set. " & _
                   "Paleomag Code will now end." & vbNewLine & _
                   "Please restart the code and go to the Settings Window " & _
                   "to add more calibration values.", , _
                   "AF ERROR!"
                   
            End
                   
        End If
            
        'Loop through the Axial coil calibration array
        'Note: this may loop like an Off-By-One error, but the calibration array
        '      is actually one row larger than AFAxialCount (the zeroth row
        '      contains zero, zero)
        For i = 1 To AFAxialCount
        
            'Check to see if MonVolt is equal to the X calibration value
            If MonVolt = AFAxial(i, 0) Then
            
                'Return the matching field value
                FindFieldFromVolts = AFAxial(i, 1)
                
                Exit For
                
            End If
            
            'Now, check to see if we're in between the current and prior calibration values
            If MonVolt < AFAxial(i, 0) And MonVolt > AFAxial(i - 1, 0) Then
            
                'User linear interpolation (Y = A*(MonVolt - X(i-1)) + Y(i-1)) to get the matching field value
                Slope = (AFAxial(i, 1) - AFAxial(i - 1, 1)) / (AFAxial(i, 0) - AFAxial(i - 1, 0))
                
                FindFieldFromVolts = AFAxial(i - 1, 1) + Slope * (MonVolt - AFAxial(i - 1, 0))
            
                Exit For
                
            End If
            
        Next i
        
        'Check to see if MonVolt is greater than the larger X value in the calibration array
        If MonVolt > AFAxial(AFAxialCount, 0) Then
        
            i = AFAxialCount
        
            'Need to interpolate upward using last two points of the calibration array
            Slope = (AFAxial(i, 1) - AFAxial(i - 1, 1)) / (AFAxial(i, 0) - AFAxial(i - 1, 0))
                
            FindFieldFromVolts = AFAxial(i, 1) + Slope * (MonVolt - AFAxial(i, 0))
            
        End If
        
    ElseIf AFCoilSystem = TransverseCoilSystem Then
    'Transverse Coil
    
        'Check to make sure AFTransCount > 1
        If AFTransCount <= 1 Then
        
            'User hasn't entered in enough calibration values
            MsgBox "Only one AF Trans Coil calibration value has been set. " & _
                   "Paleomag Code will now end." & vbNewLine & _
                   "Please restart the code and go to the Settings Window " & _
                   "to add more calibration values.", , _
                   "AF ERROR!"
                   
            End
                   
        End If
            
        'Loop through the Trans coil calibration array
        'Note: this may loop like an Off-By-One error, but the calibration array
        '      is actually one row larger than AFTransCount (the zeroth row
        '      contains zero, zero)
        For i = 1 To AFTransCount
        
            'Check to see if MonVolt is equal to the X calibration value
            If MonVolt = AFTrans(i, 0) Then
            
                'Return the matching field value
                FindFieldFromVolts = AFTrans(i, 1)
                
                Exit For
                
            End If
            
            'Now, check to see if we're in between the current and prior calibration values
            If MonVolt < AFTrans(i, 0) And MonVolt > AFTrans(i - 1, 0) Then
            
                'User linear interpolation (Y = A*(MonVolt - X(i-1)) + Y(i-1)) to get the matching field value
                Slope = (AFTrans(i, 1) - AFTrans(i - 1, 1)) / (AFTrans(i, 0) - AFTrans(i - 1, 0))
                
                FindFieldFromVolts = AFTrans(i - 1, 1) + Slope * (MonVolt - AFTrans(i - 1, 0))
            
                Exit For
                
            End If
            
        Next i
        
        'Check to see if MonVolt is greater than the larger X value in the calibration array
        If MonVolt > AFTrans(AFTransCount, 0) Then
        
            i = AFTransCount
        
            'Need to interpolate upward using last two points of the calibration array
            Slope = (AFTrans(i, 1) - AFTrans(i - 1, 1)) / (AFTrans(i, 0) - AFTrans(i - 1, 0))
                
            FindFieldFromVolts = AFTrans(i, 1) + Slope * (MonVolt - AFTrans(i, 0))
            
        End If
        
    End If
            
End Function

                               
Public Function GetChanFromColl(ByRef BoardsCol As Boards, _
                                ByVal BoardName As String, _
                                ByVal ChanType As String, _
                                ByVal ChanName As String) As Channel

    Dim TempBoard As Board
    Dim TempChannels As Channels
    
    'Initialize TempBoard, and TempChannels to Nothing
    Set TempBoard = Nothing
    Set TempChannels = Nothing
    
    'Get TempBoard from BoardsCol
    Set TempBoard = BoardsCol(BoardName)

    'Get the corresponding channel objects
    Select Case ChanType
    
        Case "AI"
        
            'Channel is in the Analog Input channels collection
            'Snatch it!
            Set TempChannels = TempBoard.AInChannels
        
        Case "AO"
        
            'Channel is in the Analog Output channels collection
            'Snatch it!
            Set TempChannels = TempBoard.AOutChannels
        
        Case "DI"
        
            'Channel is in the Digital Input channels collection
            'Snatch it!
            Set TempChannels = TempBoard.DInChannels
        
        Case "DO"
        
            'Channel is in the Digital Output channels collection
            'Snatch it!
            Set TempChannels = TempBoard.DOutChannels
        
    End Select
    
    'Set Return Channel value to Null first
    Set GetChanFromColl = Nothing
    
    'Turn on error handling
    On Error Resume Next
    
        'Attempt to get channel
        Set GetChanFromColl = TempChannels(ChanName)
        
        'Error Check
        If Err.number <> 0 Then
        
            'Wasn't able to find a matching channel
            Err.Raise Err.number, _
                      "modAF_DAQ->GetChanFromColl", _
                      "Unable to get DAQ Channel named """ & ChanName & _
                      """ from DAQ Boards Collection." & vbNewLine & _
                      "Parent Board: """ & BoardName & """" & vbNewLine & _
                      "Channel Type: """ & ChanType & """" & vbNewLine & vbNewLine & _
                      "System Error Message:" & vbNewLine & _
                      Err.Description
                      
            'Return a Nothing value
            Set GetChanFromCoil = Nothing
                  
            Exit Function
            
        End If
        
    'Turn Off error handling
    If GetChanFromColl Is Nothing Then
    
        'Wasn't able to find a matching channel
        Err.Raise 5, _
                  "modAF_DAQ->GetChanFromColl", _
                  "Unable to get DAQ Channel named """ & ChanName & _
                  """ from DAQ Boards Collection." & vbNewLine & _
                  "Parent Board: """ & BoardName & """" & vbNewLine & _
                  "Channel Type: """ & ChanType & """"
    
        'Return a nothing
        Set GetChanFromCoil = Nothing
    
        Exit Function
        
    End If
        
    
    'Deallocate Temp Board and Channels objects
    Set TempBoard = Nothing
    Set TempChannels = Nothing

End Function

Public Sub GetChannelColDependencies _
    (ByRef AssignedChanCol As Channels, _
     ByRef OldChanCol As Channels, _
     ByRef NewChanCol As Channels, _
     ByRef DependencyList() As String)
                                            
    Dim N As Long
    Dim M As Long
    Dim i As Long
    Dim j As Long
    Dim BeforeLen As Long
    Dim AfterLen As Long
    
    If NewChanCol Is Nothing Then
    
        'Only getting dependencies for one Channel collection
        '(with no comparison to a new channel collection)
        N = OldChanCol.Count
        
        If N <= 0 Then
        
            'There are no channels in this collection
            'There cannot be any function assignments for
            'this empty channel collection
            Exit Sub
            
        End If
        
        For i = 1 To N
        
            GetChannelDependencies AssignedChanCol, _
                                   OldChanCol(i), _
                                   DependencyList
                                   
        Next i
    
    Else
    
        'Check to see if the Two collection lists are the same length
        N = OldChanCol.Count
        M = NewChanCol.Count
        
        'Check to see if the two collections are the same length
        If N = M Then
        
            'Yay, they're the same length
            'If Chan Type is the same, then even if the Channel Name
            'and Board Name are different, match the two together
            'and do not add to the list of dependencies
            
            'Loop through the channels
            For i = 1 To N
            
                'If Chan types are different, then there is a conflict
                'And the dependencies can't be transfered
                If OldChanCol(i).ChanType <> NewChanCol(i).ChanType Then
                
                    'Need to get the channel dependencies for the old channel
                    GetChannelDependencies AssignedChanCol, _
                                           OldChanCol, _
                                           DependencyList
                                           
                ElseIf OldChanCol(i).ChanName <> NewChanCol(i).ChanName Or _
                       OldChanCol(i).BoardName <> NewChanCol(i).BoardName _
                Then
                
                    'Store the dimensions of DependencyList prior to getting
                    'the transferable dependencies
                    
                    'Turn on error handling
                    On Error Resume Next
                    
                        BeforeLen = UBound(DependencyList)
                        
                        'If DependencyList has not been dimensioned yet
                        'Set the BeforeLen = 1
                        If Err.number = 9 Then
                        
                            BeforeLen = 1
                            
                        End If
                        
                    'Turn off error handling
                    On Error GoTo 0
                
                    'Get the Dependencies, but add:
                    '" [{Old Board: BoardName, Old Channel: ChanName " & vbnewline &
                    '   "New Board: BoardName, New Channel: ChanName]"
                    'to the End of each assigned function dependency string
                    GetChannelDependencies AssignedChanCol, _
                                           OldChanCol, _
                                           DependencyList
                                           
                    'Get the new length after the dependency search
                    AfterLen = UBound(DependencyList)
                    
                    'If AfterLen > BeforeLen, then dependencies were found
                    If AfterLen > BeforeLen Then
                    
                        'Loop through the dependencies from Before Len to AfterLen - 1
                        For j = BeforeLen To AfterLen - 1
                        
                            TempStr = DependencyList(j)
                            DependencyList(j) = TempStr & " [{Old Board: " & _
                                                OldChanCol(i).BoardName & _
                                                ", Old Channel: " & _
                                                OldChanCol(i).ChanName & vbNewLine & _
                                                "New Board: " & _
                                                NewChanCol(i).BoardName & _
                                                ", New Channel: " & _
                                                NewChanCol(i).ChanName & "}]"
                                                
                        Next j
                                     
                    End If
                
                End If
                
            Next i
            
        Else
        
            'The two channel collections have different lengths
            'If there is a complete match (Board Name, Channel Name, Channel Type)
            'then don't add that channel's dependency to the list
            'Else, add the dependency
            
            'Need to compare each channel in each collection with each other
            'with no repeats (ugh.)
            'Setup for loops to be a lower or upper triangle as opposed to a full matrix comparison
            
            If M > 0 Then
            
                For i = 1 To N
                
                    For j = 1 To M
                        
                        With OldChanCol(i)
                            
                            If .BoardName = NewChanCol(j).BoardName And _
                               .ChanName = NewChanCol(j).ChanName And _
                               .ChanType = NewChanCol(j).ChanType _
                            Then
                                
                                'No need to do anything, this
                                'is an exact match
                                'end the inner for loop
                                j = M + 10
                                
                            End If
                            
                        End With
                        
                    Next j
                    
                    'Check to see if a match was found
                    If j <> M + 10 Then
                    
                        'No match was found
                        'Need to add dependencies to the list
                        GetChannelDependencies AssignedChanCol, _
                                               OldChanCol(i), _
                                               DependencyList
                    
                    
                    End If
                    
                Next i
                            
            Else
            
                'M = 0 - there are no new channels (old channels will be totally erased)
                'if the board edit proceeds
                'Add the dependencies of the channels in the old channel collection
                For i = 1 To N
                
                    GetChannelDependencies AssignedChanCol, _
                                           OldChanCol(i), _
                                           DependencyList
                                           
                Next i
            
            End If
    
        End If
        
    End If
        
End Sub

Public Sub GetChannelDependencies(ByRef AssignedChanCol As Channels, _
                                  ByRef ChanObj As Channel, _
                                  ByRef DependencyList() As String)
                                        
    Dim N As Long
    Dim M As Long
    Dim i As Long
    
    'This code enables the user to call multiple instances of GetChannelDependencies
    'and append the results onto an ever growing dependency list array
    
    'Turn on error handling
    On Error Resume Next
    
        'Try to get the dimensions of the Dependency list
        N = UBound(DependencyList)
        
        'If Error # = 9, then the Dependency list hasn't been dimensioned yet
        If Err.number = 9 Then
        
            'Default Dependency List to 1 element = "-1"
            ReDim DependencyList(1)
            DependencyList(0) = "-1"
            
        End If
        
    On Error GoTo 0
       
    'Check the dependencies of the Assigned DAQ channel objects
    
    'Get the number of channels in the inputed Assigned Channels collection
    N = AssignedChanCol.Count
    
    'Loop through the assigned channels and check for matches
    'to the inputed channel object (using the BoardName, ChanName, and ChanType fields)
    For i = 1 To N
    
        With AssignedChanCol(i)
        
            If .BoardName = ChanObj.BoardName And _
               .BoardININum = ChanObj.BoardININum And _
               .ChanName = ChanObj.ChanName And _
               .ChanType = ChanObj.ChanType _
            Then
            
                'Check to see the current size of dependency list
                N = UBound(DependencyList)
                
                'Resize the list to add a new element
                ReDim Preserve DependencyList(N + 1)
                
                DependencyList(0) = .ChanDescs.GetAll(", ") & " Channel: " & .ChanName
                
            End If
            
        End With
        
    Next i
                                                
End Sub

                                        


Public Sub GetFunctions(ByRef AssignedChanCol As Channels, _
                        ByRef FunctionArray() As String, _
                        ByRef ChanTypeArray() As String, _
                        Optional ByVal ChanType As String = "-1", _
                        Optional ByVal UnassignedOnly As Boolean = True)
                        
    Dim N As Long
    Dim i As Long
    
    'Set N = number of assigned channels
    N = AssignedChanCol.Count
    
    'Redimension the function and channel type arrays
    ReDim FunctionArray(1)
    ReDim ChanTypeArray(1)
    
    'Set first elements = "-1" by default
    FunctionArray(0) = "-1"
    ChanTypeArray(0) = "-1"
    
    'Loop through the channels and add there
    'descriptions to the FunctionArray()
    'Depending on the values of ChanType and UnassignedOnly
    For i = 1 To N
    
        With AssignedChanCol(i)
            
            'Check to make sure that the channels
            'complies with the imputed filters
            If (UnassignedOnly = False Or _
                .ChanName <> "ERROR") And _
               (ChanType = "-1" Or _
                ChanType = .ChanType) _
            Then
            
                'Get current size of the function array
                N = UBound(FunctionArray)
                
                'Add one to the sizes of both function and channel type arrays
                '(The two arrays should always be the same size)
                ReDim Preserve FunctionArray(N + 1)
                ReDim Preserve ChanTypeArray(N + 1)
                
                'Add the current channels description to the new element
                FunctionArray(N) = .ChanDescs.GetAll(", ")
                
                'Add the current channels description to the new element
                ChanTypeArray(N) = .ChanType
                
            End If
            
        End With
        
    Next i
                        
End Sub

Public Function HiMetricToPixels(lVal As Long) As Long

    HiMetricToPixels = frmAFTuner.ScaleX(lVal, vbHimetric, vbPixels)

End Function

    
Public Sub Initialize_Boards()
    
    Dim i As Long
    Set SystemBoards = Nothing
    Set SystemBoards = New Boards

    'Need to Inialize the two Boards
    'This will be done from the paleomag.ini file when this code is inserted
    'Into the over paleomag code module and will link to a Options tab
    'So that the User can edit this information
    
    'For inserting this code into Paleomag program
    'This initialization will happen at login\
    'And the information from the two boards will be read in
    'From the paleomag.ini file
    'Additionally, there will be an Options tab for the
    'User to setup new boards
    
    'Setup Monitor Board
    With SystemBoards.Add(, "PCI-DAS6030")
        
        .BoardName = "PCI-DAS6030"
        .CommProtocol = MCC_UL
        .BoardNum = 0
        .BoardFunction = MONITOR & "," & AFRamp
        .BoardMode = SINGLEMODE
        .MaxAInRate = 100000
        .MaxAOutRate = 100000
        Set .range = New range
        .range.RangeType = BIP10VOLTS
        .DoutPortType = AUXPORT
                
        For i = 0 To 15
        
            With .AInChannels.Add(, "AI-" & Trim(str(i)))
            
                .ChanName = "AI-" & Trim(str(i))
                .ChanNum = i
                
            End With
                            
            If i < 2 Then
                
                With .AOutChannels.Add(, "AO-" & Trim(str(i)))
                
                    .ChanName = "AO-" & Trim(str(i))
                    .ChanNum = i
                    
                End With
                               
            End If
            
        Next i
        
        With .DOutChannels
        
            For i = 0 To 7
            
                With .Add(, "DIO-" & Trim(str(i)))
                
                    .ChanName = "DIO-" & Trim(str(i))
                    .ChanNum = i
                    
                End With
                
            Next i
            
        End With
        
        With .DInChannels
        
            For i = 0 To 7
            
                With .Add(, "DIO-" & Trim(str(i)))
                
                    .ChanName = "DIO-" & Trim(str(i))
                    .ChanNum = i
                    
                End With
                
            Next i
                        
        End With
        
        .DIOConfigured = False
        
    End With
    
    
    'Setup SignalBoard
    With SystemBoards.Add(, "USB-1616HS-2")
            
        .BoardName = "USB-1616HS-2"
        .CommProtocol = MCC_UL
        .BoardNum = 1
        .BoardFunction = SIGNALGENERATOR & "," & AFRELAYCONTROL
        .BoardMode = SINGLEMODE
        .MaxAInRate = 1000000
        .MaxAOutRate = 1000000
        Set .range = New range
        .range.RangeType = BIP10VOLTS
        .DoutPortType = FIRSTPORTA
                
        For i = 0 To 15
        
            With .AInChannels.Add(, "AI-" & Trim(str(i)))
        
                .ChanName = "AI-" & Trim(str(i))
                .ChanNum = i
                
            End With
            
            If i < 2 Then
                
                With .AOutChannels.Add(, "AO-" & Trim(str(i)))
                
                    .ChanName = "AO-" & Trim(str(i))
                    .ChanNum = i
                    
                End With
                        
            End If
            
        Next i
        
        With .DOutChannels
        
            For i = 0 To 23
            
                With .Add(, "DIO-" & Trim(str(i)))
                
                    .ChanName = "DIO-" & Trim(str(i))
                    .ChanNum = i
                    
                End With
                
            Next i
            
        End With
        
        With .DInChannels
        
            For i = 0 To 23
            
                With .Add(, "DIO-" & Trim(str(i)))
                
                    .ChanName = "DIO-" & Trim(str(i))
                    .ChanNum = i
                    
                End With
                
            Next i
        
        End With
        
        .DIOConfigured = False
        
    End With
        
    'Setup ADWIN Board
    With SystemBoards.Add(, "ADWIN-light-16")
        
        .BoardName = "ADWIN-light-16"
        .CommProtocol = ADWIN_COM
        .BoardNum = 1
        .BoardFunction = "MONITOR,AFRAMP,SIGNALGENERATOR,AFRElAYCONTROL"
        .BoardMode = DIFFERENTIALMODE
        .MaxAInRate = 50000
        .MaxAOutRate = 50000
        Set .range = New range
        .range.MaxValue = 10
        .range.MinValue = -10
                
        With .AInChannels
                
            For i = 1 To 15 Step 2
            
                With .Add(, "ADC-" & Format(i, "00"))
                
                    .ChanName = "ADC-" & Format(i, "00")
                    .ChanNum = i
                    
                End With
                
            Next i
                
                                
        End With
                                
        With .AOutChannels
        
            For i = 1 To 2
            
                With .Add(, "DAC-" & Format(i, "00"))
                
                    .ChanName = "DAC-" & Format(i, "00")
                    .ChanNum = i
                    
                End With
                
            Next i
                        
        End With
                                   
            
        With .DOutChannels
            
            For i = 0 To 5
            
                With .Add(, "DIGOUT-" & Trim(str(i)))
                
                    .ChanName = "DIGOUT-" & Trim(str(i))
                    .ChanNum = i
                    
                End With
                
            Next i
                
        End With
            
        With .DInChannels
            
            For i = 0 To 5
            
                With .Add(, "DIGIN-" & Trim(str(i)))
                
                    .ChanName = "DIGIN-" & Trim(str(i))
                    .ChanNum = i
                    
                End With
                
            Next i
            
        End With
        
        .DIOConfigured = True
        
    End With
        
End Sub

Public Sub Initialize_Waves()

    Set WaveForms = Nothing
    Set WaveForms = New Waves
    
    'In Paleomag program, this initialization process
    'Will pull data from the paleomag.ini file
    'And will be called during the user login process
    
     
    With WaveForms.Add(, "SIGNALGENERATOR")
        
        Set .BoardUsed = SystemBoards.Item(2)
        .setrange BIP10VOLTS
        .BufferAlloc = False
        .DoDeallocate = True
        .MemBuffer = 0
        .Duration = 0
        .IO = IOOUTPUT
        .IOOptions = BACKGROUND + CONTINUOUS
        .MinVoltage = 0
        .PeakVoltage = 0

        .NumPoints = 0
        .SineFreqMin = 0
        .StartPoint = 0
        Set .Chan = SystemBoards.Item(2).AOutChannels.Item(1)
        
    End With
    
    
    With WaveForms.Add(, "AFRAMPUP")
        
        .WaveName = AFRAMPUP
        Set .BoardUsed = SystemBoards.Item("ADWIN-light-16")
        .setrange UNI10VOLTS
        .BufferAlloc = False
        .DoDeallocate = True
        .MemBuffer = 0
        .Duration = 0
        .IO = IOOUTPUT
        .IOOptions = BACKGROUND
        .MinVoltage = 0
        .PeakVoltage = 0
        .NumPoints = 0
        .SineFreqMin = -1
        .IORate = 0
        .StartPoint = 0
        Set .Chan = SystemBoards.Item("ADWIN-light-16").AOutChannels.Item(1)
        
    End With
    
    With WaveForms.Add(, "AFRAMPDOWN")
        
        .WaveName = AFRAMPDOWN
        Set .BoardUsed = SystemBoards.Item("ADWIN-light-16")
        .setrange UNI10VOLTS
        .BufferAlloc = False
        .DoDeallocate = True
        .MemBuffer = 0
        .Duration = 0
        .IO = IOOUTPUT
        .IOOptions = BACKGROUND
        .MinVoltage = 0
        .PeakVoltage = 0
        .NumPoints = 0
        .SineFreqMin = -1
        .IORate = 0
        .StartPoint = 0
        Set .Chan = SystemBoards.Item("ADWIN-light-16").AOutChannels.Item(1)
        
    End With
    
    With WaveForms.Add(, "AFMONITOR")
        
        .WaveName = AFMonitor
        Set .BoardUsed = SystemBoards.Item("ADWIN-light-16")
        .setrange BIP10VOLTS
        .BufferAlloc = False
        .DoDeallocate = True
        .MemBuffer = 0
        .Duration = 0
        .IO = IOINPUT
        .IOOptions = BACKGROUND
        .MinVoltage = 0
        .PeakVoltage = 0
        .NumPoints = 0
        .SineFreqMin = -1
        .IORate = 0
        .StartPoint = 0
        Set .Chan = SystemBoards.Item("ADWIN-light-16").AInChannels.Item(1)
        
    End With
    
    With WaveForms.Add(, "IRMMONITOR")
    
        .WaveName = IRMMonitor
        Set .BoardUsed = SystemBoards.Item("PCI-DAS6030")
        .setrange BIP10VOLTS
        .BufferAlloc = False
        .DoDeallocate = True
        .MemBuffer = 0
        .Duration = 0
        .IO = IOINPUT
        .IOOptions = BACKGROUND
        .MinVoltage = 0
        .PeakVoltage = 0
        .NumPoints = 0
        .SineFreqMin = -1
        .IORate = 0
        .StartPoint = 0
        Set .Chan = SystemBoards.Item("PCI-DAS6030").AInChannels.Item(1)
        
    End With
    
    With WaveForms.Add(, "BASELINE")
    
        .WaveName = Baseline
        Set .BoardUsed = SystemBoards.Item("PCI-DAS6030")
        .setrange BIP10VOLTS
        .BufferAlloc = False
        .DoDeallocate = True
        .MemBuffer = 0
        .Duration = 0
        .IO = IOINPUT
        .IOOptions = BACKGROUND
        .MinVoltage = 0
        .PeakVoltage = 0
        .NumPoints = 0
        .SineFreqMin = -1
        .IORate = 0
        .StartPoint = 0
        Set .Chan = SystemBoards.Item("PCI-DAS6030").AInChannels.Item(1)
        
    End With
            
End Sub

'Public Sub radix2_FFT(ByVal N As Integer)
'
'1000 'THE FAST FOURIER TRANSFORM
''copyright © 1997-1999 by California Technical Publishing
''published with  permission from Steven W Smith, www.dspguide.com
''GUI by logix4u , www.logix4u.net
''modified by logix4u, www.logix4.net
'1010 'Upon entry, N% contains the number of points in the DFT, Real[ ] and
'1020 'Imag[ ] contain the real and imaginary parts of the input. Upon return,
'1030 'Real[ ] and Imag[ ] contain the DFT output. All signals run from 0 to N%-1.
'1060 NM1% = N% - 1
'1070 ND2% = N% / 2
'1080 M% = CInt(Log(N%) / Log(2))
'1090 j% = ND2%
'1100 '
'1110 For i% = 1 To N% - 2 'Bit reversal sorting
'1120 If i% >= j% Then GoTo 1190
'1130 TR = Real(j%)
'1140 TI = Imag(j%)
'1150 Real(j%) = Real(i%)
'1160 Imag(j%) = Imag(i%)
'1170 Real(i%) = TR
'1180 Imag(i%) = TI
'1190 k% = ND2%
'1200 If k% > j% Then GoTo 1240
'1210 j% = j% - k%
'1220 k% = k% / 2
'1230 GoTo 1200
'1240 j% = j% + k%
'1250 Next i%
'1260 '
'1270 For L% = 1 To M% 'Loop for each stage
'1280 Le% = CInt(2 ^ L%)
'1290 Le2% = Le% / 2
'1300 UR = 1
'1310 UI = 0
'1320 SR = Cos(Pi / Le2%) 'Calculate sine & cosine values
'1330 SI = -Sin(Pi / Le2%)
'1340 For j% = 1 To Le2% 'Loop for each sub DFT
'1350 JM1% = j% - 1
'1360 For i% = JM1% To NM1% Step Le% 'Loop for each butterfly
'1370 IP% = i% + Le2%
'1380 TR = Real(IP%) * UR - Imag(IP%) * UI 'Butterfly calculation
'1390 TI = Real(IP%) * UI + Imag(IP%) * UR
'1400 Real(IP%) = Real(i%) - TR
'1410 Imag(IP%) = Imag(i%) - TI
'1420 Real(i%) = Real(i%) + TR
'1430 Imag(i%) = Imag(i%) + TI
'1440 Next i%
'1450 TR = UR
'1460 UR = TR * SR - UI * SI
'1470 UI = TR * SI + UI * SR
'1480 Next j%
'1490 Next L%
'1500 '
'End Sub

'Public Sub radix2_FFT_isaac(ByRef DataArray() As Double, ByVal N As Long)
'
'1000 'THE FAST FOURIER TRANSFORM
''copyright © 1997-1999 by California Technical Publishing
''published with  permission from Steven W Smith, www.dspguide.com
''GUI by logix4u , www.logix4u.net
''modified by logix4u, www.logix4.net
''modified by Isaac Hilburn, Caltech Paleomag Lab, Sep., 2009
''Note:
''   1)  Changing original FFT algorithm from Logix4u.net so that the Real & Imag
''       Arrays are two halves of the DataArray that's passed in
''       For real-valued input, the imag half will contain zero values
''   2)  Also, changed variables from integers to longs so could handle
''       FFT's with greater than 32678 real values (>65k total points in array)
''   3)  Added new variables to reduce the number of calculations performed in the
''       for loops
''   4)  Replaced If...then Goto... statements in bit reversing code with Do...while...loop or If...then... statemtns
'
'Dim NM1 As Long     'N - 1
'Dim NM2 As Long     'N - 2
'Dim Ndiv4 As Long   'N \ 4 - Note using integer division operator
'Dim Ndiv2 As Long   'N \ 2
'Dim M As Long       'N = 2^M
'Dim j As Long       'For loop counter
'Dim i As Long       'For loop counter
'Dim k As Long       'For loop counter
'Dim L As Long       'For loop counter for L-shaped butterfly (?)
'Dim Le As Long      'L raised to a power of 2
'Dim LeDiv2 As Long  'Le \ 2 - using integer division operator
'Dim JM1 As Long     ' j - 1
'
'Dim TempR As Double   'Temp variable for storing real var results
'Dim TempI As Double   'Temp variable for storing imag var results
'Dim CC As Double   'Memory space to store cosine of Le or LeDiv2 as needed for DFT for that butterfly step
'Dim SS As Double   'Memory space to store sine of Le or LeDiv2 as needed for DFT...
'
''Set initial values for relevant counter, end/start value holding variables
'Ndiv2 = N \ 2
'ND2M1 = Ndiv2 - 1
'ND2M2 = Ndiv2 - 2
''Using Ndiv4 instead of Ndiv2 as N = total number of Real + Imag data points
'Ndiv4 = N \ 4       'Note, using integer/long division operator, \
'M = CLng(Log(N) / Log(2)) 'Using double division operator, /
'
'j = Ndiv4
'
''----------------------Bit Reversal Sorting--------------------------
'For i = 1 To ND2M2
'
'    If i < j Then
'
'        TempR = DataArray(j)                        'Real data
'        TempI = DataArray(j + Ndiv2)                'Imag data
'        DataArray(j) = DataArray(i)                 'Bit swap real data
'        DataArray(j + Ndiv2) = DataArray(i + Ndiv2) 'Bit swap imag data
'        DataArray(i) = TempR
'        DataArray(i + Ndiv2) = TempI
'
'
'1010 'Upon entry, N% contains the number of points in the DFT, Real[ ] and
'1020 'Imag[ ] contain the real and imaginary parts of the input. Upon return,
'1030 'Real[ ] and Imag[ ] contain the DFT output. All signals run from 0 to N%-1.
'1060 NM1% = N% - 1
'1070 ND2% = N% / 2
'1080 M% = CInt(Log(N%) / Log(2))
'1090 j% = ND2%
'1100 '
'1110 For i% = 1 To N% - 2 'Bit reversal sorting
'1120 If i% >= j% Then GoTo 1190
'1130 TR = Real(j%)
'1140 TI = Imag(j%)
'1150 Real(j%) = Real(i%)
'1160 Imag(j%) = Imag(i%)
'1170 Real(i%) = TR
'1180 Imag(i%) = TI
'1190 k% = ND2%
'1200 If k% > j% Then GoTo 1240
'1210 j% = j% - k%
'1220 k% = k% / 2
'1230 GoTo 1200
'1240 j% = j% + k%
'1250 Next i%
'1260 '
'1270 For L% = 1 To M% 'Loop for each stage
'1280 Le% = CInt(2 ^ L%)
'1290 Le2% = Le% / 2
'1300 UR = 1
'1310 UI = 0
'1320 SR = Cos(Pi / Le2%) 'Calculate sine & cosine values
'1330 SI = -Sin(Pi / Le2%)
'1340 For j% = 1 To Le2% 'Loop for each sub DFT
'1350 JM1% = j% - 1
'1360 For i% = JM1% To NM1% Step Le% 'Loop for each butterfly
'1370 IP% = i% + Le2%
'1380 TR = Real(IP%) * UR - Imag(IP%) * UI 'Butterfly calculation
'1390 TI = Real(IP%) * UI + Imag(IP%) * UR
'1400 Real(IP%) = Real(i%) - TR
'1410 Imag(IP%) = Imag(i%) - TI
'1420 Real(i%) = Real(i%) + TR
'1430 Imag(i%) = Imag(i%) + TI
'1440 Next i%
'1450 TR = UR
'1460 UR = TR * SR - UI * SI
'1470 UI = TR * SI + UI * SR
'1480 Next j%
'1490 Next L%
'1500 '
'End Sub

Public Sub Invert3x3HermitianMatrix(ByRef Matrix() As Double)
                                    
    Dim Determinant As Double
    Dim m00, m01, m02 As Double
    Dim m11, m12 As Double
    Dim m22 As Double
    
    m00 = Matrix(0)
    m01 = Matrix(1)
    m02 = Matrix(2)
    m11 = Matrix(3)
    m12 = Matrix(4)
    m22 = Matrix(5)
           
    'Find the Determinant
    Determinant = m00 * m11 * m22 _
                    + 2 * m01 * m12 * m02 _
                        - m00 * m12 * m12 _
                        - m01 * m01 * m22 _
                        - m02 * m11 * m02
                    
    'If the determinant is non-zero, find the inverse
    If Determinant <> 0 Then
    
        Matrix(0) = (m11 * m22 - m12 * m12) / Determinant
        Matrix(1) = (m02 * m12 - m01 * m22) / Determinant
        Matrix(2) = (m01 * m12 - m02 * m11) / Determinant
        Matrix(3) = (m00 * m22 - m02 * m02) / Determinant
        Matrix(4) = (m02 * m01 - m00 * m12) / Determinant
        Matrix(5) = (m00 * m11 - m01 * m01) / Determinant
        
    Else
    
        'Determinant is zero, matrix is non-invertible
        Err.Raise 616, "Invert3x3HermitianMatrix", "Matrix is non-invertible."
        
    End If
                                    
End Sub

Public Sub Invert4x4HermitianMatrix(ByRef Matrix() As Double)
                                    
    Dim Determinant As Double
    Dim m00, m01, m02, m03 As Double
    Dim m11, m12, m13 As Double
    Dim m22, m23 As Double
    Dim m33 As Double
    
    Dim m01x01, m01x02, m01x03 As Double
    Dim m02x02, m03x03 As Double
    Dim m12x12, m13x13, m23x23 As Double
    
    m00 = Matrix(0)
    m01 = Matrix(1)
    m02 = Matrix(2)
    m03 = Matrix(3)
    m11 = Matrix(4)
    m12 = Matrix(5)
    m13 = Matrix(6)
    m22 = Matrix(7)
    m23 = Matrix(8)
    m33 = Matrix(9)
    
    m01x01 = m01 * m01
    m01x02 = m01 * m02
    m01x03 = m01 * m03
    m02x02 = m02 * m02
    m02x03 = m02 * m03
    m03x03 = m03 * m03
    m12x12 = m12 * m12
    m13x13 = m13 * m13
    m23x23 = m23 * m23
    
    'Calculate the determinant - some duplicate calculations are being made
    'here, but no need to waste time in optimization
    Determinant = m03x03 * m12x12 - _
                        m02x03 * m13 * m12 - _
                        m03x03 * m11 * m22 + _
                        m01x03 * m13 * m22 + _
                        m02x03 * m11 * m23 - _
                        m01x03 * m12 * m23 - _
                        m02x03 * m12 * m13 + _
                        m02x02 * m13x13 + _
                        m01x03 * m22 * m13 - _
                        m00 * m13x13 * m22 - _
                        m01x02 * m23 * m13 + _
                        m00 * m12 * m23 * m13 + _
                        m02x03 * m11 * m23 - _
                        m01x02 * m13 * m23 - _
                        m01x03 * m12 * m23 + _
                        m00 * m13 * m12 * m23 + _
                        m01x01 * m23x23 - _
                        m00 * m11 * m23x23 - _
                        m02x02 * m11 * m33 + _
                        m01x02 * m12 * m33 + _
                        m01x02 * m12 * m33 - _
                        m00 * m12x12 * m33 - _
                        m01x01 * m22 * m33 + _
                        m00 * m11 * m22 * m33


    If Determinant <> 0 Then

        Matrix(0) = (m12 * m23 * m13 - _
                            m13x13 * m22 + _
                            m13 * m12 * m23 - _
                            m11 * m23x23 - _
                            m12x12 * m33 + _
                            m11 * m22 * m33) / Determinant
    
        Matrix(1) = (m03 * m22 * m13 - _
                            m02 * m23 * m13 - _
                            m03 * m12 * m23 + _
                            m01 * m23x23 + _
                            m02 * m12 * m33 - _
                            m01 * m22 * m33) / Determinant
    
        Matrix(2) = (m02 * m13x13 - _
                             m03 * m12 * m13 + _
                             m03 * m11 * m23 - _
                             m01 * m13 * m23 - _
                             m02 * m11 * m33 + _
                             m01 * m12 * m33) / Determinant
        
        Matrix(3) = (m03 * m12x12 - _
                            m02 * m13 * m12 - _
                            m03 * m11 * m22 + _
                            m01 * m13 * m22 + _
                            m02 * m11 * m23 - _
                            m01 * m12 * m23) / Determinant
    
        Matrix(4) = (m02x03 * m23 - _
                            m03x03 * m22 + _
                            m02x03 * m23 - _
                            m00 * m23x23 - _
                            m02x02 * m33 + _
                            m00 * m22 * m33) / Determinant
    
        Matrix(5) = (m03x03 * m12 - _
                            m02x03 * m13 - _
                            m01x03 * m23 + _
                            m00 * m13 * m23 + _
                            m01x02 * m33 - _
                            m00 * m12 * m33) / Determinant
    
        Matrix(6) = (m02x02 * m13 - _
                            m02x03 * m12 + _
                            m01x03 * m22 - _
                            m00 * m13 * m22 - _
                            m01x02 * m23 + _
                            m00 * m12 * m23) / Determinant
    
        Matrix(7) = (m01x03 * m13 - _
                            m03x03 * m11 + _
                            m01x03 * m13 - _
                            m00 * m13x13 - _
                            m01x01 * m33 + _
                            m00 * m11 * m33) / Determinant
    
        Matrix(8) = (m02x03 * m11 - _
                            m01x02 * m13 - _
                            m01x03 * m12 + _
                            m00 * m13 * m12 + _
                            m01x01 * m23 - _
                            m00 * m11 * m23) / Determinant
    
        Matrix(9) = (m01x02 * m12 - _
                            m02x02 * m11 + _
                            m01x02 * m12 - _
                            m00 * m12x12 - _
                            m01x01 * m22 + _
                            m00 * m11 * m22) / Determinant
    
    Else
    
        Err.Raise 616, "Invert4x4HermitianMatrix", "Matrix Non-invertible."
    
    End If
                                    
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Linear approximation using least squares method
'
'The subroutine calculates coefficients of  the  line  approximating  given
'function.
'
'Input parameters:
'    X   -   array[0..N-1], it contains a set of abscissas.
'    Y   -   array[0..N-1], function values.
'    StartIndex - Point in the X & Y arrays to start looking at
'    N   -   number of points, N>=1
'
'Output parameters:
'    A, B-   coefficients of linear approximation y(x) = A+B*x
'    R_2 -   goodness-of-fit factor:
'            R_2 = 1 - Sum[ {Y(x) - (A+B*x)}^2 ] / Sum[ {Y(x) - Y_average }^2 ]
'            R_2 = 1 - Sum of the square of the residuals in Y / Sum of the variances in Y
'
'  -- ALGLIB --
'     Copyright by Bochkanov Sergey
'
'  Modified by RAPID, November 2009
'  Isaac Hilburn
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub LinearLeastSquares(ByRef X() As Double, _
         ByRef Y() As Double, _
         ByVal N As Long, _
         ByVal StartIndex As Long, _
         ByRef a As Double, _
         ByRef b As Double, _
         ByRef R_2 As Double)
         
    Dim NumPoints As Long
    Dim SumX_2 As Double
    Dim SumX As Double
    Dim SumY As Double
    Dim SumXY As Double
    Dim SumRes2 As Double
    Dim SumVar As Double
    Dim AvgY As Double
    Dim D1 As Double
    Dim D2 As Double
    Dim t1 As Double
    Dim T2 As Double
    Dim Phi As Double
    Dim c As Double
    Dim S As Double
    Dim M As Double
    Dim i As Long

    NumPoints = N - StartIndex
    SumX_2 = 0#
    SumX = 0#
    SumY = 0#
    SumXY = 0#
    For i = StartIndex To N - 1 Step 1
        SumX = SumX + X(i)
        SumX_2 = SumX_2 + X(i) ^ 2
        SumY = SumY + Y(i)
        SumXY = SumXY + X(i) * Y(i)
    Next i
    Phi = Atn2(2# * SumX, SumX_2 - NumPoints) / 2#
    c = Cos(Phi)
    S = Sin(Phi)
    D1 = c ^ 2 * NumPoints + S ^ 2 * SumX_2 - 2# * S * c * SumX
    D2 = S ^ 2 * NumPoints + c ^ 2 * SumX_2 + 2# * S * c * SumX
    If Abs(D1) > Abs(D2) Then
        M = Abs(D1)
    Else
        M = Abs(D2)
    End If
    t1 = c * SumY - S * SumXY
    T2 = S * SumY + c * SumXY
    If Abs(D1) > M * MachineEpsilon * 1000# Then
        t1 = t1 / D1
    Else
        t1 = 0#
    End If
    If Abs(D2) > M * MachineEpsilon * 1000# Then
        T2 = T2 / D2
    Else
        T2 = 0#
    End If
    a = c * t1 + S * T2
    b = -(S * t1) + c * T2
    
    'Now that have the slope and intercept for the best estimate line,
    'can get the R^2 factor.
    'SumY can be used to get the average Y value
    AvgY = SumY / NumPoints
    
    'Initialize sum of square of the residuals and the variance to 0
    SumRes2 = 0#
    SumVar = 0#
    
    'Now calculate the sum of the square of the residuals
    'r_i = Y(i) - (A + B*X(i))
    For i = StartIndex To N - 1 Step 1
    
        SumRes2 = SumRes2 + (Y(i) - a - b * X(i)) ^ 2
        SumVar = SumVar + (Y(i) - AvgY) ^ 2
    
    Next i
    
    'Set R_2 = 1 - sum of the squared residuals divided (scaled) by the sum of the
    'variances in Y
    R_2 = 1 - SumRes2 / SumVar
    
End Sub

Public Sub LoadArrayToGrid(ByRef gridobj As MSHFlexGrid, _
                           ByRef DataArray() As String, _
                           ByRef CallingForm As Form, _
                           ByVal FirstRow As Long, _
                           Optional ByVal FirstCol As Long = 0, _
                           Optional ByVal AddLineNumbers As Boolean = False)
                           
    'Take the values in a N x M array and store them in a
    'N x M size span of a grid object at a start row and start col
    'position
                      
    Dim i As Long
    Dim j As Long
    Dim N As Long
    Dim M As Long
    
    'If dataarray is nothing, exit the sub
    On Error Resume Next
    
        'Get the number of data rows in the array
        N = UBound(DataArray, 1)
        
        If Err.number <> 0 Then Exit Sub
        
    On Error GoTo 0
        
    'If N > 0, get the number of data cols in the array
    If N > 0 Then
    
        M = UBound(DataArray, 2)
        
    Else
    
        M = 0
        
    End If
    
    'If there are zero rows and cols, then exit the sub
    If M = 0 And N = 0 Then Exit Sub
        
    'If User has selected to add line numbers, then add an additonal col to the
    'size of the flex-grid
    If AddLineNumbers = True Then M = M + 1
        
    With gridobj
    
        'Now resize the grid object as needed to store the new text values
        .Rows = FirstRow + N
        .Cols = FirstCol + M
        
        'Now loop through the data array and load the values to
        'the grid object
        For i = 0 To N - 1
        
            'Set the row
            .row = i + FirstRow
            
            'Loop through the columns
            For j = 0 To M - 1
            
                'Set the col
                .Col = j + FirstCol
                
                'check to see if the user has chosen to add line numbers
                'and if this is the first column
                If AddLineNumbers = True And _
                   .Col = FirstCol _
                Then
                
                    'Add the line number to text
                    .text = Trim(str(i + FirstRow))
                    
                    'Resize the column width as needed
                    If .ColWidth(FirstCol) < CallingForm.TextWidth(.text) * 2 Then
                    
                        .ColWidth(FirstCol) = CallingForm.TextWidth(.text) * 2
                        
                    End If
                    
                Else
                
                    If FormatString = vbNullString Then
                        
                        'Set the text
                        .text = Trim(DataArray(i, j))
                        
                    Else
                    
                        'Set the text with formating
                        .text = Trim(DataArray(i, j))
                    
                    End If
                    
                    If val(.text) = 0 Then .text = ""
                    
                    'Set the colwidth, if needed
                    If .ColWidth(j + FirstCol) < CallingForm.TextWidth(.text) * 1.2 Then
                    
                        .ColWidth(j + FirstCol) = CallingForm.TextWidth(.text) * 1.2
                        
                    End If
                
                End If
                
            Next j
            
        Next i
                    
    End With
                           
End Sub

                     
'Adaptation of MedianThreeQuickSort
'
' Uses a median of three sorting algorithm to sort the rows of a 2D array
'
' Original Algorithm by: Niklaus Emil Wirth
'
' VB6 coded version by: Ellis Dee
' URL: http://www.vbforums.com/showpost.php?p=2909260&postcount=14
'
' July 2010
' Isaac Hilburn
' Code Modified to work with 2D arrays and allow user to specify
' which column to sort on
'
Public Sub MedianThreeQuickSort_DBL_2D(ByRef pvararray() As Double, _
                                    Optional SortCol As Long = 0, _
                                    Optional ByVal plngLeft As Long = 0, _
                                    Optional ByVal plngRight As Long = 0)

    Dim lngFirst As Long

    Dim lngLast As Long

    Dim varMid As Double

    Dim lngIndex As Long
    
    Dim ColsHi As Long
    Dim ColsLo As Long

    Dim varSwap As Double

    Dim a As Long

    Dim b As Long

    Dim c As Long

    Dim i As Long
   
    If plngRight = 0 Then

        plngLeft = LBound(pvararray, 1)

        plngRight = UBound(pvararray, 1)

    End If
    
    If plngRight > 0 Then
    
        ColsHi = UBound(pvararray, 2)
        ColsLo = UBound(pvararray, 2)

    End If
    
    
    '(July 2010 - I Hilburn)
    'Do a data check on the sort column
    If SortCol > ColsHi - 1 Or SortCol < ColsLo Then
    
        'Exit the function, the user has asked to sort on a
        'non-existent column
        Exit Sub
        
    End If

    lngFirst = plngLeft

    lngLast = plngRight

    lngIndex = plngRight - plngLeft + 1

    a = Int(lngIndex * Rnd) + plngLeft

    b = Int(lngIndex * Rnd) + plngLeft

    c = Int(lngIndex * Rnd) + plngLeft

    If pvararray(a, SortCol) <= pvararray(b, SortCol) And _
       pvararray(b, SortCol) <= pvararray(c, SortCol) _
    Then

        lngIndex = b

    Else

        If pvararray(b, SortCol) <= pvararray(a, SortCol) And _
           pvararray(a, SortCol) <= pvararray(c, SortCol) _
        Then

            lngIndex = a

        Else

            lngIndex = c

        End If

    End If

    varMid = pvararray(lngIndex, SortCol)

    Do

        Do While pvararray(lngFirst, SortCol) < varMid And _
                 lngFirst < plngRight

            lngFirst = lngFirst + 1

        Loop

        Do While varMid < pvararray(lngLast, SortCol) And _
                 lngLast > plngLeft

            lngLast = lngLast - 1

        Loop

        If lngFirst <= lngLast Then

            '(July 2010 - I Hilburn)
            'Added in this for loop to allow the transfer of 2D array col data
            For i = ColsLo To ColsHi - 1

                varSwap = pvararray(lngFirst, i)

                pvararray(lngFirst, i) = pvararray(lngLast, i)

                pvararray(lngLast, i) = varSwap
                
            Next i

            lngFirst = lngFirst + 1

            lngLast = lngLast - 1

        End If

    Loop Until lngFirst > lngLast

    If (lngLast - plngLeft) < (plngRight - lngFirst) Then
    
        If plngLeft < lngLast Then MedianThreeQuickSort_DBL_2D pvararray, plngLeft, lngLast

        If lngFirst < plngRight Then MedianThreeQuickSort_DBL_2D pvararray, lngFirst, plngRight

    Else

        If lngFirst < plngRight Then MedianThreeQuickSort_DBL_2D pvararray, lngFirst, plngRight

        If plngLeft < lngLast Then MedianThreeQuickSort_DBL_2D pvararray, plngLeft, lngLast

    End If

End Sub

                     
                     
'Adaptation of MedianThreeQuickSort
'
' Uses a median of three sorting algorithm to sort the rows of a 2D array
'
' Original Algorithm by: Niklaus Emil Wirth
'
' VB6 coded version by: Ellis Dee
' URL: http://www.vbforums.com/showpost.php?p=2909260&postcount=14
'
' July 2010
' Isaac Hilburn
' Code Modified to work with 2D arrays and allow user to specify
' which column to sort on
'
Public Sub MedianThreeQuickSort1_2D(ByRef pvararray() As String, _
                                    Optional SortCol As Long = 0, _
                                    Optional ByVal plngLeft As Long = 0, _
                                    Optional ByVal plngRight As Long = 0)

    Dim lngFirst As Long

    Dim lngLast As Long

    Dim varMid As Double

    Dim lngIndex As Long
    
    Dim ColsHi As Long
    Dim ColsLo As Long

    Dim varSwap As String

    Dim a As Long

    Dim b As Long

    Dim c As Long

    Dim i As Long
   
    If plngRight = 0 Then

        plngLeft = LBound(pvararray, 1)

        plngRight = UBound(pvararray, 1)

    End If
    
    If plngRight > 0 Then
    
        ColsHi = UBound(pvararray, 2)
        ColsLo = UBound(pvararray, 2)

    End If
    
    
    '(July 2010 - I Hilburn)
    'Do a data check on the sort column
    If SortCol > ColsHi - 1 Or SortCol < ColsLo Then
    
        'Exit the function, the user has asked to sort on a
        'non-existent column
        Exit Sub
        
    End If

    lngFirst = plngLeft

    lngLast = plngRight

    lngIndex = plngRight - plngLeft + 1

    a = Int(lngIndex * Rnd) + plngLeft

    b = Int(lngIndex * Rnd) + plngLeft

    c = Int(lngIndex * Rnd) + plngLeft

    If val(pvararray(a, SortCol)) <= val(pvararray(b, SortCol)) And _
       val(pvararray(b, SortCol)) <= val(pvararray(c, SortCol)) _
    Then

        lngIndex = b

    Else

        If val(pvararray(b, SortCol)) <= val(pvararray(a, SortCol)) And _
           val(pvararray(a, SortCol)) <= val(pvararray(c, SortCol)) _
        Then

            lngIndex = a

        Else

            lngIndex = c

        End If

    End If

    varMid = val(pvararray(lngIndex, SortCol))

    Do

        Do While val(pvararray(lngFirst, SortCol)) < varMid And _
                 lngFirst < plngRight

            lngFirst = lngFirst + 1

        Loop

        Do While varMid < val(pvararray(lngLast, SortCol)) And _
                 lngLast > plngLeft

            lngLast = lngLast - 1

        Loop

        If lngFirst <= lngLast Then

            '(July 2010 - I Hilburn)
            'Added in this for loop to allow the transfer of 2D array col data
            For i = ColsLo To ColsHi - 1

                varSwap = pvararray(lngFirst, i)

                pvararray(lngFirst, i) = pvararray(lngLast, i)

                pvararray(lngLast, i) = varSwap
                
            Next i

            lngFirst = lngFirst + 1

            lngLast = lngLast - 1

        End If

    Loop Until lngFirst > lngLast

    If (lngLast - plngLeft) < (plngRight - lngFirst) Then
    
        If plngLeft < lngLast Then MedianThreeQuickSort1_2D pvararray, plngLeft, lngLast

        If lngFirst < plngRight Then MedianThreeQuickSort1_2D pvararray, lngFirst, plngRight

    Else

        If lngFirst < plngRight Then MedianThreeQuickSort1_2D pvararray, lngFirst, plngRight

        If plngLeft < lngLast Then MedianThreeQuickSort1_2D pvararray, plngLeft, lngLast

    End If

End Sub

Public Sub NormalizeMatrix(ByRef Matrix() As Double)
'Takes an NxM real matrix and divides all of it's elements by the matrix norm

    Dim N As Long
    Dim M As Long
    Dim i As Long
    Dim j As Long
    Dim Max As Double
    Dim temp As Double
    
    N = UBound(Matrix, 1)
    M = UBound(Matrix, 2)
    
    'Initialize Max = 0
    Max = 0
    
    'Go through every element in the matrix and find the maximum value
    For i = 0 To N - 1
    
        For j = 0 To M - 1

            temp = Abs(Matrix(i, j))
            If temp > Max Then Max = temp
            
        Next j
        
    Next i
    
    'If Max doesn't equal zero, then divide all the elements of the Matrix
    'by max
    If Max <> 0 Then
        
        For i = 0 To N - 1
        
            For j = 0 To M - 1
        
                Matrix(i, j) = Matrix(i, j) / Max
                
            Next j
            
        Next i
        
    End If
        
End Sub

'----------------------------------------------------------------------------------
'   NumericalDeriv
'
'   A very basic numerical derivative subroutine that looks at the instantaneous
'   slope between one or more points.  For sets of points > 2, the function
'   uses a linear least squares algorithm to get the slope of that point set
'
'   Inputs:
'           RMS_Array() - an N x M array where the first dimension contains
'                         N data points to be numerically differentiated
'                         and the second dimension can be used for additional
'                         N - point data sets.
'           Windows()   - a M x 1 vector containing the "window" size = the number
'                         of points over which to determine the slope / derivative
'                         of the data.
'
'   Outputs:
'         Deriv_Array() - an N x M array with the numerical derivative results for
'                         the M N-point datasets inputed by the user in RMS_Array
'
'   Author:
'           Isaac Hilburn, Feb. 2010
'           RAPID Consortium
'           Caltech
'
'----------------------------------------------------------------------------------
Public Sub NumericalDeriv(ByRef RMS_array() As Double, _
                                ByRef Deriv_Array() As Double, _
                                ByRef Windows() As Long)
                                
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim N As Long
    Dim M As Long
    Dim X() As Double
    Dim Y() As Double
    Dim Slope As Double
    Dim Intercept As Double
    Dim RMS As Double
    Dim MaxWindow As Long
    
    N = UBound(RMS_array, 1)
    M = UBound(RMS_array, 2)
    
    'Change M if user has not inputed enough sliding windows for the derivation
    If UBound(Windows) < M Then M = UBound(Windows)
    
    'Redimension the Least Squares Array
    'For Each window need:
    '   X array of Length = Window
    '   Y array of Length = Window
    'Therefore first need the size of the largest Window
    MaxWindow = Windows(0)
    For i = 0 To M - 1
    
        'Need at least two points to get an instantaneous slope
        If Windows(i) < 2 Then Windows(i) = 2
    
        'Store the max window size
        If MaxWindows < Windows(i) Then MaxWindows = Windows(i)
        
    Next i
    
    'Now can dimension the least squares array
    ReDim X(MaxWindows)
    ReDim Y(MaxWindows)
    
    For i = 0 To M - 1
    
        
        For j = 0 To N - 1
    
            'If there are enough points to match this particular window
            If j >= Windows(i) - 1 Then
            
                If Windows(i) = 2 Then
                
                    Deriv_Array(j, i) = (RMS_array(j, i) - RMS_array(j - 1, i)) / Windows(i)
            
                Else
            
                    'Capture the past windows worth of data
                    For k = j - (Windows(i) - 1) To j
                        
                        X(k - (j - (Windows(i) - 1))) = RMS_array(k, i)
                        Y(k - (j - (Windows(i) - 1))) = k - (j - (Windows(i) - 1))
                        
                    Next k
                    
                    LinearLeastSquares X, _
                                        Y, _
                                        Windows(i), _
                                        0, _
                                        Slope, _
                                        Intercept, _
                                        RMS
                                        
                    Deriv_Array(j, i) = Slope / Windows(i)
                
                End If
                
            Else
            
                'Can't compute a slope for this point
                Deriv_Array(j, i) = 0
                
            End If
            
        Next j
            
    Next i
                       
End Sub

Public Sub OldInvert4x4HermitianMatrix(ByRef Matrix() As Double, _
                                    ByRef Inverse() As Double)

    'Whole Crap Ton of variables
    Dim A1, B1, C1, D1, E1, F1, G1, H1, I1, J1, K1, L1 As Double
    Dim A2, B2, C2, D2, E2, F2, G2, H2, I2, J2, K2, L2, M2, N2, O2, P2, Q2, R2, S2 As Double
    Dim A3, B3, C3, D3, E3, F3, G3, H3, I3, J3, K3, L3, M3 As Double
    Dim A4, B4, C4, D4, E4, F4, G4, H4, I4, J4 As Double
    Dim A5, B5, A6, B6, A7 As Double
    Dim Sm01, Sm02, Sm03m, Sm12, Sm13, Sm23 As Double
    Dim m01x02, m01x03, m02x03, m11x22, m11x23, m11x33 As Double
    Dim m12x23, m12x33, m13x22, m13x23, m22x33 As Double
    
    'Determinant
    Dim Determinant As Double

    'Squares of Off-Diagonal elements
    Sm01 = Matrix(0, 1) * Matrix(0, 1)
    Sm02 = Matrix(0, 2) * Matrix(0, 2)
    Sm03 = Matrix(0, 3) * Matrix(0, 3)
    Sm12 = Matrix(1, 2) * Matrix(1, 2)
    Sm13 = Matrix(1, 3) * Matrix(1, 3)
    Sm23 = Matrix(2, 3) * Matrix(2, 3)
    
    'Cross terms
    m01x02 = Matrix(0, 1) * Matrix(0, 2)
    m01x03 = Matrix(0, 1) * Matrix(0, 3)
    m02x03 = Matrix(0, 2) * Matrix(0, 3)
    m11x22 = Matrix(1, 1) * Matrix(2, 2)
    m11x23 = Matrix(1, 1) * Matrix(2, 3)
    m11x33 = Matrix(1, 1) * Matrix(3, 3)
    m12x23 = Matrix(1, 2) * Matrix(2, 3)
    m12x33 = Matrix(1, 2) * Matrix(3, 3)
    m13x22 = Matrix(1, 3) * Matrix(2, 2)
    m13x23 = Matrix(1, 3) * Matrix(2, 3)
    m22x33 = Matrix(2, 2) * Matrix(3, 3)

    'All possible non-repeating combinations of three Matrix() elements
    A1 = m00 * m11x22
    B1 = m00 * m11x23
    C1 = m00 * m11x33
    D1 = m00 * Sm12
    E1 = m00 * m12x13
    F1 = m00 * m12x23
    G1 = m00 * m12x33
    H1 = m00 * Sm13
    I1 = m00 * m13x22
    J1 = m00 * m13x23
    K1 = m00 * m22x33
    L1 = m00 * Sm23

    A2 = Sm01 * Matrix(2, 2)
    B2 = Sm01 * Matrix(2, 3)
    C2 = Sm01 * Matrix(3, 3)
    D2 = m01x02 * Matrix(1, 2)
    E2 = m01x02 * Matrix(1, 3)
    F2 = m01x02 * Matrix(2, 3)
    G2 = m01x02 * Matrix(3, 3)
    H2 = m01x03 * Matrix(1, 2)
    I2 = m01x03 * Matrix(1, 3)
    J2 = m01x03 * Matrix(2, 2)
    K2 = m01x03 * Matrix(2, 3)
    L2 = Matrix(0, 1) * Sm12
    M2 = Matrix(0, 1) * m12x23
    N2 = Matrix(0, 1) * m12x33
    O2 = Matrix(0, 1) * Sm13
    P2 = Matrix(0, 1) * m13x22
    Q2 = Matrix(0, 1) * m13x23
    R2 = Matrix(0, 1) * m22x33
    S2 = Matrix(0, 1) * Sm23

    A3 = Sm02 * Matrix(1, 1)
    B3 = Sm02 * Matrix(1, 3)
    C3 = Sm02 * Matrix(3, 3)
    D3 = m02x03 * Matrix(1, 1)
    E3 = m02x03 * Matrix(1, 2)
    F3 = m02x03 * Matrix(1, 3)
    G3 = m02x03 * Matrix(2, 3)
    H3 = Matrix(0, 2) * m11x23
    I3 = Matrix(0, 2) * m11x33
    J3 = Matrix(0, 2) * m12x13
    K3 = Matrix(0, 2) * m12x33
    L3 = Matrix(0, 2) * Sm13
    M3 = Matrix(0, 2) * m13x23
    
    A4 = Sm03 * Matrix(1, 1)
    B4 = Sm03 * Matrix(1, 2)
    C4 = Sm03 * Matrix(2, 2)
    D4 = Matrix(0, 3) * Matrix(1, 1) * Matrix(2, 2)
    E4 = Matrix(0, 3) * Matrix(1, 1) * Matrix(2, 3)
    F4 = Matrix(0, 3) * Sm12
    G4 = Matrix(0, 3) * m12x13
    H4 = Matrix(0, 3) * m12x23
    I4 = Matrix(0, 3) * m13x22
    J4 = Matrix(0, 3) * m22x23
    
    A5 = Matrix(1, 1) * Matrix(2, 2) * Matrix(3, 3)
    B5 = Matrix(1, 1) * Sm23
    
    A6 = Sm12 * Matrix(3, 3)
    B6 = Matrix(1, 2) * Matrix(1, 3) * Matrix(2, 3)
    
    A7 = Sm13 * Matrix(2, 2)
    
    'Determinant Calculation
    Determinant = Sm03 * Sm12 - 2 * m02x03 * m12x13 - Sm03 * m11x22 + 2 * m01x03 * m13x22 + _
        2 * m02x03 * m11x23 - 2 * m01x03 * m12x23 + Sm02 * Sm13 - m00 * A7 - _
        2 * m01x02 * m13x23 + 2 * m00 * B6 + Sm01 * Sm23 - m00 * B5 - Sm02 * m11x33 + _
        2 * m01x02 * m12x33 - m00 * A6 - Sm01 * m22x33 + m00 * A5
   
   'Inverse calculation is Determinant is non-zero
    If Determinant <> 0 Then
            
        'Inverse must exist
        Inverse(0, 0) = (B6 - A7 + B6 - B5 - A6 + A5) / Determinant
        Inverse(0, 1) = (I4 - M3 - H4 + S2 + K3 - R2) / Determinant
        Inverse(0, 2) = (L3 - G4 + E4 - Q2 - I3 + N2) / Determinant
        Inverse(0, 3) = (F4 - J3 - D4 + P2 + H3 - M2) / Determinant
        Inverse(1, 1) = (G3 - C4 + G3 - L1 - C3 + K1) / Determinant
        Inverse(1, 2) = (B4 - F3 - K2 + J1 + G2 - G1) / Determinant
        Inverse(1, 3) = (B3 - E3 + J2 - I1 - F2 + F1) / Determinant
        Inverse(2, 2) = (I2 - A4 + I2 - H1 - C2 + C1) / Determinant
        Inverse(2, 3) = (D3 - E2 - H2 + E1 + B2 - B1) / Determinant
        Inverse(3, 3) = (D2 - A3 + D2 - D1 - A2 + A1) / Determinant
        
    Else
    
        Err.Raise 616, "Invert4x4Matrix", "Matrix Is Non-invertible"

    End If

End Sub

'-----------------------------------------------------------------------------------------------------------------------------'
'-----------------------------------------------------------------------------------------------------------------------------'
'
'  July 23, 2010
'  Commented Out by: Isaac Hilburn
'
'  No longer using a look-up table to calibrate the ramp voltage vs the monitor voltage for the ADWIN AF ramp process
'  Instead, the ADBasic code continues to ramp the output voltage up if it looks like it's possible to reach
'  the monitor input voltage beofre the output voltage reaches the Af ramp max voltage for that coil
'
'  The frmADWIN_VoltageCalibration form has also been removed from the Paleomag project, so this code will break
'  if it is run.
'
'-----------------------------------------------------------------------------------------------------------------------------'
'
'Public Function FindRampVoltFromMonVolt(ByVal MonVolt As Double, _
'                                        Optional ByVal AFCoilSystem As Long = -128, _
'                                        Optional ByVal CallingFormName As String = "frmADWIN_AF") As Double
'
'    '(July 2010, I Hilburn)
'    '
'    'This function takes in a Peak monitor voltage value and returns the
'    'best Ramp voltage value to use to reach that peak value.  The array
'    'used to do this is different from that used for the AF Field calibration
'    'This array updates dynamically after every successful ramp cycle
'
'    Dim i As Long
'    Dim UserResponse As Long
'    Dim Slope As Double
'    Dim TempD As Double
'
'    'If the user has not inputed a coil system, set the coil system
'    'to the global active coil system
'    If AFCoilSystem = -128 Then AFCoilSystem = activecoilsystem
'
'    'Check for MonVolt <=0
'    If MonVolt <= 0 Then
'
'        'Return zero
'        FindRampVoltFromMonVolt = 0
'
'        'Exit the function
'        Exit Function
'
'    End If
'
'    'Check to see which AF coil is active
'    If activecoilsystem = axialcoilsystem Then
'
'        'Check to see if the Mon to Ramp calibration array has been created yet
'        If modConfig.AFRampAxialCount <= 1 Then
'
'            'MsgBox the User and see if they want to do the ADWIN voltage calibration right now
'            UserResponse = MsgBox("AF " & WaveForms("AFRAMPUP").BoardUsed.BoardName & " board has " & _
'                                  "not had it's Output vs Input voltages calibrated yet.  This is " & _
'                                  "independent of the Monitor Voltage vs Peak Field calibration and " & _
'                                  "must be done before the Peak Field calibration or any ADWIN AF ramp " & _
'                                  "can be performed." & vbNewLine & vbNewLine & _
'                                  "The process is fully automated, would you like to start it now?", _
'                                  vbYesNo, _
'                                  "Warning!")
'
'            'Check for a "No" response
'            If UserResponse = vbNo Then
'
'                'Return a -2 voltage
'                FindRampVoltFromMonVolt = -2
'
'                Exit Function
'
'            End If
'
'            'No monitor, ramp value pairs have been calculated yet
'            'Check to see what the source form is
'            If CallingFormName = "frmADWIN_AF" Then
'
'                'Ask user to run the ADWIN Ramp VS Monitor voltage calibration
'                Load frmADWIN_VoltageCalibration
'                frmADWIN_VoltageCalibration.Show
'
'                FindRampVoltFromMonVolt = -1
'
'                Exit Function
'
'            ElseIf CallingFormName = "frmCalibrateCoils" Then
'
'                'Ask user to run the ADWIN Ramp VS Monitor voltage calibration
'                Load frmADWIN_VoltageCalibration
'                frmADWIN_VoltageCalibration.Show
'
'                FindRampVoltFromMonVolt = -1
'
'                Exit Function
'
'            End If
'
'        End If
'
'        'Now we have at least two Mon to Ramp voltage calibration values,
'        'can interpolate between them
'
'        'Loop through the Axial Ramp voltage calibration array
'        'Note: this may loop like an Off-By-One error, but the calibration array
'        '      is actually one row larger than AFRampAxialCount (the zeroth row
'        '      contains zero, zero)
'        For i = 1 To AFRampAxialCount
'
'            'Check to see if MonVolt is equal to the X calibration value
'            If MonVolt = AFRampAxial(i, 0) Then
'
'                'Return the matching field value
'                FindRampVoltFromMonVolt = AFRampAxial(i, 1)
'
'                Exit Function
'
'            End If
'
'            'Now, check to see if we're in between the current and prior calibration values
'            If MonVolt < AFRampAxial(i, 0) And MonVolt > AFRampAxial(i - 1, 0) Then
'
'                'User linear interpolation (Y = A*(MonVolt - X(i-1)) + Y(i-1)) to get the matching field value
'                Slope = (AFRampAxial(i, 1) - AFRampAxial(i - 1, 1)) / (AFRampAxial(i, 0) - AFRampAxial(i - 1, 0))
'
'                FindRampVoltFromMonVolt = AFRampAxial(i - 1, 1) + Slope * (MonVolt - AFRampAxial(i - 1, 0))
'
'                Exit Function
'
'            End If
'
'        Next i
'
'        'Now, if no match was found, then need to interpolate upwards to the Ramp voltage max
'        'Check to see if MonVolt is greater than the larger X value in the calibration array
'        i = AFRampAxialCount
'
'        'Need to interpolate upward using last two points of the calibration array
'        Slope = (AFRampAxial(i, 1) - AFRampAxial(i - 1, 1)) / (AFRampAxial(i, 0) - AFRampAxial(i - 1, 0))
'
'        'If the resulting matching Ramp Voltage is greater than the max ramp voltage,
'        'then return the max ramp voltage
'        TempD = AFRampAxial(i, 1) + Slope * (MonVolt - AFRampAxial(i, 0))
'        If TempD > modConfig.AfAxialRampMax Then TempD = modConfig.AfAxialRampMax
'
'        FindRampVoltFromMonVolt = TempD
'
'    Else
'
'        'Check to see if the Mon to Ramp calibration array has been created yet
'        If modConfig.AFRampTransCount <= 1 Then
'
'            'No monitor, ramp value pairs have been calculated yet
'            'Return the value in the frmADWIN_AF.txtRampPeakVoltage
'            FindRampVoltFromMonVolt = val(frmADWIN_AF.txtRampPeakVoltage)
'
'            Exit Function
'
'        End If
'
'        'Now we have at least two Mon to Ramp voltage calibration values,
'        'can interpolate between them
'
'        'Loop through the Axial Ramp voltage calibration array
'        'Note: this may loop like an Off-By-One error, but the calibration array
'        '      is actually one row larger than AFRampTransCount (the zeroth row
'        '      contains zero, zero)
'        For i = 1 To AFRampTransCount
'
'            'Check to see if MonVolt is equal to the X calibration value
'            If MonVolt = AFRampTrans(i, 0) Then
'
'                'Return the matching field value
'                FindRampVoltFromMonVolt = AFRampTrans(i, 1)
'
'                Exit Function
'
'            End If
'
'            'Now, check to see if we're in between the current and prior calibration values
'            If MonVolt < AFRampTrans(i, 0) And MonVolt > AFRampTrans(i - 1, 0) Then
'
'                'User linear interpolation (Y = A*(MonVolt - X(i-1)) + Y(i-1)) to get the matching field value
'                Slope = (AFRampTrans(i, 1) - AFRampTrans(i - 1, 1)) / (AFRampTrans(i, 0) - AFRampTrans(i - 1, 0))
'
'                FindRampVoltFromMonVolt = AFRampTrans(i - 1, 1) + Slope * (MonVolt - AFRampTrans(i - 1, 0))
'
'                Exit Function
'
'            End If
'
'        Next i
'
'        'Now, if no match was found, then need to interpolate upwards to the Ramp voltage max
'        'Check to see if MonVolt is greater than the larger X value in the calibration array
'        i = AFRampTransCount
'
'        'Need to interpolate upward using last two points of the calibration array
'        Slope = (AFRampTrans(i, 1) - AFRampTrans(i - 1, 1)) / (AFRampTrans(i, 0) - AFRampTrans(i - 1, 0))
'
'        'If the resulting matching Ramp Voltage is greater than the max ramp voltage,
'        'then return the max ramp voltage
'        TempD = AFRampTrans(i, 1) + Slope * (MonVolt - AFRampTrans(i, 0))
'        If TempD > modConfig.AfTransRampMax Then TempD = modConfig.AfTransRampMax
'
'        FindRampVoltFromMonVolt = TempD
'
'    End If
'
'End Function
'
'-----------------------------------------------------------------------------------------------------------------------------'
'-----------------------------------------------------------------------------------------------------------------------------'

Public Sub PauseBetweenUseCoils_InSeconds(ByVal pause_seconds As Integer)

     Dim start_time As Long
     
     start_time = timeGetTime()
     
     Dim prewait_status As String
     
     prewait_status = frmProgram.sbStatusBar.Panels(3).text
     
     frmProgram.StatusBar "Waiting ... " + Format(pause_seconds, "0.0"), 3
     
     Dim elapsed_time As Double
     Dim time_left As Double
     
     elapsed_time = (timeGetTime() - start_time) / 1000
     
     Do While elapsed_time < pause_seconds And elapsed_time >= 0
     
        DoEvents
        
        time_left = pause_seconds - elapsed_time
        
        
        If ((time_left * 10) Mod 1) < 0.1 Then
            frmProgram.StatusBar "Waiting ... " + Format(time_left, "0.0"), 3
        End If
     
        elapsed_time = (timeGetTime() - start_time) / 1000
     Loop
     
     frmProgram.StatusBar prewait_status, 3

End Sub

Public Sub PauseTill(ByVal EndTime As Double)

    Dim doContinue As Boolean
    Dim CurTimer As Double
    
    'Set continue flag for the pause loop = true
    doContinue = True
    
'    'If End time of pause loop crosses over midnight, then
'    'remove one days worth of time from it so that the program
'    'won't get stuck in an unending loop
'    If EndTime > 86400 Then
'
'        EndTime = EndTime - 86400
'
'    End If
    Do While doContinue
    
        CurTimer = timeGetTime()
        
        'If current time is greater than or equal to the end time
        'then set the loop continue flag to false so that the pause
        'loop ends
        If CurTimer >= EndTime Then
        
            doContinue = False
            
        End If
        
        DoEvents
        
    Loop

End Sub

Public Sub PauseTill_NoEvents(ByVal EndTime As Double)

    Dim doContinue As Boolean
    Dim CurTimer As Double
    
    'Set continue flag for the pause loop = true
    doContinue = True
    
'    'If End time of pause loop crosses over midnight, then
'    'remove one days worth of time from it so that the program
'    'won't get stuck in an unending loop
'    If EndTime > 86400 Then
'
'        EndTime = EndTime - 86400
'
'    End If
    Do While doContinue
    
        CurTimer = timeGetTime()
        
        'If current time is greater than or equal to the end time
        'then set the loop continue flag to false so that the pause
        'loop ends
        If CurTimer >= EndTime Then
        
            doContinue = False
            
        End If
                        
    Loop

End Sub


Public Sub Predict_Sine_Wave(ByRef SineArray() As Single, _
                                ByVal IORate As Long, _
                                ByVal PeakVoltage As Double, _
                                ByRef Freq As Double, _
                                ByRef TimeToPeak As Double, _
                                ByRef HighestVoltage As Double, _
                                Optional ByVal BaselineBiasVoltage As Double = 0)
                                
    '--------------------------------------------------------------
    '   Important Note!!
    '
    '   Need to correct for bias on the
    '   Monitor Channel
    '
    '   A baseline scan was done before the start of the ramp
    '   on the Monitor Analog Input Channel
    '
    '   Before process the sine wave data from the Monitor
    '   memory buffer, need to remove the average bias from
    '   all the data points to re-zero them
    '-------------------------------------------------------------



    Dim N As Long               'Number of elements in SineArray
    Dim M As Long               'Number of elements in Max And Min arrays

    'Vars to store the zero corrected current and last sine values
    Dim CurSineValue As Double
    Dim LastSineValue As Double

    Dim ZeroPos() As Long
    Dim ZeroVal() As Single
    Dim MaxAndMinPos() As Double
    Dim MaxAndMinVal() As Double
    
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim StartIndex As Long
    
    Dim SumZeroSpacing As Long
    Dim SumVarZeroSpacing As Double
    Dim AvgZeroSpacing As Double
    
    'Boolean Flags
    Dim IsAboveMinThreshold As Boolean  'This flag is set to true if both the max and min values
                                        'For the monitor signal sine wave are above the
                                        'public var: MonitorMinVoltageThreshold
    
    
    'Variables to store the results of the least squares fit
    'to the absolute value of the max and min values
    Dim Slope As Double
    Dim Intercept As Double
    Dim R_2 As Double
    
    'Get number of points in the sine array
    N = UBound(SineArray)
    
    'Redimension arrays so that they can hold one data point.
    ReDim ZeroPos(1)
    ReDim ZeroVal(1)
    ReDim MaxAndMinPos(1)
    ReDim MaxAndMinVal(1)
    
    'Initialize Max and Min arrays and zero arrays to 0
    ZeroPos(0) = -1
    ZeroVal(0) = -1
    MaxAndMinVal(0) = -1
    MaxAndMinPos(0) = -1
    
    'Initialize Is Above Minimum Monitor Voltage boolean flag to false
    IsAboveMinThreshold = False
    
    'Set Index used to load and access data in the zero arrays to 0
    k = 0
    
    'Set Index for Max and min arrays to 0
    j = 0
    
'-----------Debug--------------------------------
'
'    For i = 0 To N - 1
'
'        Debug.Print Trim(Str(i)) & ", " & Trim(Str(SineArray(i)))
'
'    Next i
'
'------------------------------------------------
    
    'Cycle through Sine Array first time to correct for the baseline bias
    'voltage on the Monitor Analog input channel
    For i = 0 To N - 1
    
        SineArray(i) = SineArray(i) - BaselineBiasVoltage
        
    Next i
      
    'Cycle through the Sine array again and get the zeros and the max and min values
    For i = 0 To N - 1
    
        If i = 0 Then
        
            'First point in array, don't look at previous point, because
            'there isn't one
            
            'Now see if the first point is equal to the ZeroVoltage
            If SineArray(i) = 0 Then
            
                'That was easy, we've found a zero
                'Record the zero value and the position
                ZeroVal(k) = SineArray(i)
                ZeroPos(k) = i
                                
                'Iterate k forward
                k = k + 1
                        
                'Redimension the Zero Arrays
                ReDim Preserve ZeroVal(k + 1)
                ReDim Preserve ZeroPos(k + 1)
                
                'Initialize new Zero array indeces to -1
                ZeroVal(k) = -1
                ZeroPos(k) = -1
                
                'If Current Max And Min Val is -1, then do not iterate
                'the counter on the Max and Min Arrays, do not redimension
                If MaxAndMinVal(j) = -1 And MaxAndMinPos(j) = -1 Then
                
                    'Do nothing
                    
                Else
                
                    'Iterate j index
                    j = j + 1
                    
                    'Redimension the Max & Min arrays
                    ReDim Preserve MaxAndMinVal(j + 1)
                    ReDim Preserve MaxAndMinPos(j + 1)
                    
                    'Initialize new Max And Min array indeces to -1
                    MaxAndMinVal(j) = -1
                    MaxAndMinPos(j) = -1
                    
                End If
                
            End If
                
                
        Else

            'i > 0, we can look at the previous value and also
            'compare the MaxAndMinVal to the current value

            'Now see if absolute value of the current point is greater
            'than the current max/min value
            ' --and--
            'Check to see that the absolute value of the current point is
            'is greater than the minimum monitor voltage threshold
            'This screens out data that is too noisy for the prediction
            'code to work correctly
            
            If MaxAndMinVal(j) < Abs(SineArray(i)) _
                And Abs(SineArray(i)) > MonitorMinVoltageThreshold Then
                    
                'Change Boolean flag for Monitor threshold
                IsAboveMinThreshold = True
                            
             
                'Now need to see if the current value is just due to spikiness
                'or noise in the data, this'll also skip over inflection
                'points, but we don't need to worry about where those are
                If Abs(SineArray(i) - SineArray(i - 1)) _
                    / Abs(SineArray(i - 1)) > 0.05 Then
                
                    'Values changing way too fast to be at the true max or min
                    'of the sine array and not be noise
                    'Do not record the current absolute value to the max or min array
                    
                Else
               
                    'Values are changing slowly
                    'Now check to see if two zeros have been found yet
                    If k < 2 Then

                        'Only 1 or no zero has been found so far
                        'Record absolute value of current point as max/min

                        'Set max/min to the abs(sine array value)
                        MaxAndMinVal(j) = Abs(SineArray(i))

                        'Store position of this point
                        MaxAndMinPos(j) = i


                    Else

                        'At least two zeros have been found
                        'See if this max or min lies less than 20 points
                        'away from where we'd expect it to be based on the prior
                        'two zero positions
                        If Abs(Abs(i - ZeroPos(k - 1)) - Abs((ZeroPos(k - 1) - ZeroPos(k - 2)) / 2)) < 20 Then

                            'We're where the actual max or min of this period of the sine
                            'wave should be.  Record the max or min

                            'Set max/min to the abs(sine array value)
                            MaxAndMinVal(j) = Abs(SineArray(i))

                            'Store position of this point
                            MaxAndMinPos(j) = i

                        End If

                    End If
                    
                End If
                
            End If
            
            'Now see if the current point is zero
            If SineArray(i) = 0 Then
            
                'That was easy, we've found a zero
                'Record the zero value and the position
                ZeroVal(k) = SineArray(i)
                ZeroPos(k) = i
                
                'Iterate k forward
                k = k + 1
                        
                'Redimension the Zero arrays
                ReDim Preserve ZeroVal(k + 1)
                ReDim Preserve ZeroPos(k + 1)
                
                'Initialize new Zero array indeces to -1
                ZeroVal(k) = -1
                ZeroPos(k) = -1
                
                'If Current Max And Min Val is -1, then do not iterate
                'the counter on the Max and Min Arrays, do not redimension
                If MaxAndMinVal(j) = -1 And MaxAndMinPos(j) = -1 Then
                
                    'Do nothing
                    
                Else
                
                    'Iterate j index
                    j = j + 1
                    
                    'Redimension the Max & Min arrays
                    ReDim Preserve MaxAndMinVal(j + 1)
                    ReDim Preserve MaxAndMinPos(j + 1)
                    
                    'Initialize new Max And Min array indeces to -1
                    MaxAndMinVal(j) = -1
                    MaxAndMinPos(j) = -1
                    
                End If
                            
            'Current point is not zero
            Else
                        
                'Now see if prior point is zero
                If SineArray(i - 1) = 0 Then
                
                    'Prior point is zero, don't look for a zero
                    'this iteration, in fact, do nothing
                    
                Else
                
                    'Prior point and current point aren't zero,
                    'can divide by them
                    'and check to see if there's been a sign
                    'change between the current point and the
                    'prior point
                    If CInt(SineArray(i - 1) / Abs(SineArray(i - 1))) = _
                        -1 * CInt(SineArray(i) / Abs(SineArray(i))) Then
                        
                        'There's been a sign change
                        'Record the current pos and value as the zero
                        ZeroVal(k) = SineArray(i)
                        ZeroPos(k) = i
                        
                        'Iterate k forward
                        k = k + 1
                       
                        'Redimension the Zero arrays
                        ReDim Preserve ZeroVal(k + 1)
                        ReDim Preserve ZeroPos(k + 1)
                        
                        'Initialize new Zero array indeces to -1
                        ZeroVal(k) = -1
                        ZeroPos(k) = -1
                        
                        'If Current Max And Min Val is zero, then do not iterate
                        'the counter on the Max and Min Arrays, do not redimension
                        If MaxAndMinVal(j) = -1 And MaxAndMinPos(j) = -1 Then
                        
                            'Do nothing
                            
                        Else
                        
                            'Iterate j index
                            j = j + 1
                            
                            'Redimension the Max & Min arrays
                            ReDim Preserve MaxAndMinVal(j + 1)
                            ReDim Preserve MaxAndMinPos(j + 1)
                            
                            'Initialize new Max And Min array indeces to -1
                            MaxAndMinVal(j) = -1
                            MaxAndMinPos(j) = -1
                            
                        End If
                        
                    End If
                        
                End If
                
            End If
                    
        End If
        
    Next i
    
'----------Debug Code-------------------------

    'Open file for storing the zeros and max and min values
    Dim StrDate As String
    Dim DirPath As String
    
    DirPath = "C:\Documents and Settings\lab\Desktop\Test MCC Board 11-16-2009\"
    StrDate = Format(Now, "DD-MM-YYYY_HH_MM")
    
    Open DirPath & StrDate & ".csv" For Append As #2
    
    'Now, have zeros, and absolute values of all the min and max values
    'Can get an estimate of the frequency from the spacing between
    ' all of the zeros
    
    'Set the variable to store the sum of the # of points between each zero to 0
    SumZeroSpacing = 0
    
    'When running through the zeros, ignore the last point in the array
    'it may never have been filled with an actual zero
    'And start at the second point in the array; we want the distance
    'between the zeros, not the zero positions themselves
    M = UBound(ZeroPos)
    
    Print #2, "Zeros - Total # = " & Trim(str(M))
    Print #2, "0," & ZeroPos(0) & "," & ZeroVal(0)
    
    For i = 1 To M - 2
    
        Print #2, Trim(str(i)) & "," & ZeroPos(i) & "," & ZeroVal(i)
        
        SumZeroSpacing = SumZeroSpacing + ZeroPos(i) - ZeroPos(i - 1)
        
    Next i

    'Divide by the number of spaces which equals i
    AvgZeroSpacing = SumZeroSpacing / i
    
    'Divide the IORate by the 2 * average zero spacing (should be one full period)
    'and set the result to Freq.
    Freq = IORate / (AvgZeroSpacing * 2)
    
    'If zero positions aren't pretty darned regular, then the AF signal
    'is not symmteric about zero,  get the standard deviation and raise an
    'error if it's too big, and the error will be grabbed by the procedure
    'calling this sub and used to determine if the AF ramp should be ended.
    'If the sine wave signal is not zeroed, then we shouldn't be trying
    'to AF a real sample - that'd be bad
    
    'Set the Sum of the variances in the zero position to zero
    SumVarZeroPos = 0
    
    For i = 1 To M - 2
    
        SumVarZeroPos = SumVarZeroPos + _
                        ((ZeroPos(i) - ZeroPos(i - 1)) - AvgZeroSpacing) ^ 2
    
    Next i
    
    'If standard deviation is greater than 10 points, then raise the non-zero
    'error
'    If Sqr(SumVarZeroPos / i) > 10 And MaxAndMinPos(0) <> -1 Then
'
'        'Time to raise the non-zero error - #667 (rouding error....)
'        Err.Raise 667, "Predict_Sine_Wave", "AF Signal Not Symmetric About Zero"
'
'        Exit Sub
'
'    End If
        
    'Now, need to get a leastquares fit to the absolute value of
    'the max and min values
    
    'First, inspect the first max or min position and make sure it's the proper
    'distance from the first zero
    If Abs(Abs(MaxAndMinPos(0) - ZeroPos(0)) _
                     - AvgZeroPosition / 2) > 2 Then
                     
        'Spacing is wrong
        'Skip first point
        'Set Start Index = 1
        StartIndex = 1
        
    Else
    
        'Spacing is right, first max or min is in the expected spot
        'Include the first point
        StartIndex = 0
        
    End If
        
    'Store the number of Max and Min values that we snatched
    M = UBound(MaxAndMinVal)
    
    'Do Least Squares Fit, resulting line parameters y = Intercept + Slope * x
    'are referenced into the last two input parameters
    
    If UBound(MaxAndMinPos) = 1 And MaxAndMinPos(0) = -1 Then
    
        'Do nothing, voltage hasn't reach above the threshold to do prediction yet
        TimeToPeak = 100
        HighestVoltage = -1
        
        'End sub routine - can't make a real prediction yet
        Exit Sub
        
    End If
    
    LinearLeastSquares MaxAndMinPos, MaxAndMinVal, M, StartIndex, Intercept, Slope, R_2

    Print #2, " "
    Print #2, "Max & Min - Total # = " & Trim(str(M))
    Print #2, "intercept = " & Trim(str(Intercept))
    Print #2, "slope = " & Trim(str(Slope))
    Print #2, "R^2 = " & Trim(str(R_2))
    
    'Cycle through Max and Min Vals and print them to file with the least squares fit parameters
    For i = 0 To M - 1
    
        Print #2, Trim(str(i)) & "," & MaxAndMinPos(i) & "," & MaxAndMinVal(i)
    
    Next i
    
    Print #2, " "
    Close #2
    
    'Check R_2 value,
'    If R_2 < 0.5 Then
'
'        'We've got some distortion on the sine wave, or enough noise that the max and
'        'min values have been thrown off
'
'        'Raise an AF Wave Distortion error
'        Err.Raise 668, _
'                    "Predict_Sine_Wave", _
'                    "AF Signal has non-trivial noise obscurring the max and min " & _
'                    "values of the Sine wave signal." & vbNewLine & "R^2 = " & Trim(Str(R_2))
'
'    End If
                

    'Check slope of line
    'If slope is close to zero, then compare to peak voltage
    
    'Accuracy of the PCI-DAS-6030 board is about 1.2 mV
    'Therefore, if the total change in value
    'over the line is within 2 mV of zero, then the slope is
    'effectively zero.
    If Abs(Slope) * (M - StartIndex) < 0.002 Then
    
        'If Slope is ~ zero, then the Y-intercept is
        'approximately the highest voltage
        HighestVoltage = Intercept
        
        
    'Check if the slope is negative
    ElseIf Slope < 0 Then
    
        'Crap
        'Raise a diminishing slope error.
        Err.Raise 668, _
                    "Predict_Sine_Wave", _
                    "Slope of least squares best fit line to envelope of the AF Ramp " & _
                    "Up Signal is negative and non-zero."
                    
        Exit Sub
        
        
    'Slope is non-zero and positive
    Else
    
        'Get the highest modelled voltage through the least squares
        'fit slope and Y-intercept
        HighestVoltage = (M - StartIndex) * Slope + Intercept
        
    End If
    
    'y_peak = intercept + slope * x_peak, where x_peak = point from the start of
    'the AF monitor waveform point sample set at which the AF ramp
    'will reach the peak voltage based on this linear approximation to the AF signal
    'Therefore, Time to Peak Point = x_peak / IORate (peak voltage - intercept) / slope
    TimeToPeak = ((PeakVoltage - Intercept) / Slope) / IORate
    
    'Need to adjust TimeToPeak based on the number of points sampled from the AF monitor wave
    'The amount of time elapsed over the sample set of points = N / IORate
    'Note, this does not include the time delay of extracting these points from the
    'memory buffer, nor does it include the time delay of running all of this code
    'Which, hopefully, is only an itsy-bitsy amount of time
    TimeToPeak = TimeToPeak - N / IORate
    
    
End Sub

Public Sub RenumberGridCol(ByRef gridobj As MSHFlexGrid, _
                           ByRef CallingForm As Form, _
                           ByVal ColIndex As Long, _
                           Optional ByVal FirstNumb As Long = 1, _
                           Optional ByVal FirstRow As Long = 0, _
                           Optional ByVal LastRow As Long = -1, _
                           Optional ByVal doResizeGrid As Boolean = True)
                           
    Dim i As Long
    Dim TempL As Long
                               
    With gridobj
                               
        'Check to see if the inputed ColIndex is greater than
        'the last column in the grid object,
        'If so, exit the sub routine - can't number nonexistent columns
        If gridobj.Cols - 1 < ColIndex Then Exit Sub
        
        'Ditto if ColIndex < 0
        If ColIndex < 0 Then Exit Sub
        
        'We have a good ColIndex value
        
        'Check the values of the optional parameters
        If LastRow = -1 Then LastRow = gridobj.Rows - 1
        
        'Set the FirstNumb to a local variable
        TempL = FirstNumb
        
        'Check to see if FirstRow is smaller than the number of fixed rows in the grid
        If FirstRow < .FixedRows And _
           LastRow >= .FixedRows _
        Then
        
            FirstRow = .FixedRows
        
        End If
        
        'Now iterate through the rows in the grid renumbering each cell
        'in the indicated column
        For i = FirstRow To LastRow
    
            'Save the current number as a string to the current
            'cell in the column
            .TextMatrix(i, ColIndex) = Trim(str(TempL))
            
            'Increment the number to save to the column
            TempL = TempL + 1
            
        Next i
        
        'Resize the grid for this column
        If doResizeGrid = True Then
                
            ResizeGrid gridobj, _
                       CallingForm, _
                       FirstRow, _
                       LastRow, _
                       ColIndex, _
                       ColIndex
                       
        End If
                       
    End With

End Sub

Public Sub ResizeGrid(ByRef gridobj As MSHFlexGrid, _
                      ByRef CallingForm As Form, _
                      Optional ByVal FirstRow As Long = 0, _
                      Optional ByVal LastRow As Long = -1, _
                      Optional ByVal FirstCol As Long = 0, _
                      Optional ByVal LastCol As Long = -1, _
                      Optional ByVal Multiplier As Double = 1.2, _
                      Optional ByVal SkipMerged As Boolean = True, _
                      Optional ByVal FirstColDbl As Boolean = True)
                      
    Dim i As Long
    Dim j As Long
    Dim MaxWidth As Long
    Dim TempStr As String
       
    With gridobj
       
        'Check the optional parameter values for -1's
        If LastRow = -1 Then LastRow = .Rows - 1
        If LastCol = -1 Then LastCol = .Cols - 1
        
        'Iterate through every cell and resize it if it's larger than the
        'max prior size
        'Iterate first by col, and then by row (resize one column at a time)
        For j = FirstCol To LastCol
            
            'Start max width off at zero
            MaxWidth = 0
            
            'row iteration
            For i = FirstRow To LastRow
            
                'Save the text of this cell to a local var
                TempStr = .TextMatrix(i, j)
                
                If CallingForm.TextWidth(TempStr) > MaxWidth And _
                   Not (SkipMerged = True And _
                        .MergeRow(i) = True) _
                Then
                
                    MaxWidth = CallingForm.TextWidth(TempStr)
                    
                End If
                
            Next i
            
            'If it's col 0, increase the multiplier to 2
            If j = 0 And FirstColDbl = True Then
            
                .ColWidth(j) = MaxWidth * 2
            
            Else
                
                'Resize the column to the max width * multiplier
                .ColWidth(j) = MaxWidth * Multiplier
                          
            End If
                          
       Next j
       
    End With
                          
End Sub

Public Sub ResolveBoardDependencies(ByRef AssignedChanCol As Channels, _
                                    ByRef OldBoard As Board, _
                                    ByRef NewBoard As Board)

    Dim N As Long
    Dim i As Long

    'Check to see if the NewBoard is nothing
    If NewBoard Is Nothing Then
    
        'The old board is being deleted not edited
        'Just need to go through the put in "ERROR"
        'in the Channel Name and Board Name field of every assigned
        'channel that is dependent upon the old board
    
        'Get the number of assigned channels in the collection
        N = AssignedChanCol.Count
        
        'Iterate through the assigned channels
        For i = 1 To N
        
            With AssignedChanCol(i)
            
                If .BoardName = OldBoard.BoardName And _
                   .BoardININum = OldBoard.BoardININum _
                Then
                
                    'Need to Write "ERROR" into both the board and channel name fields
                    .ChanName = "ERROR"
                    .BoardName = "ERROR"
                    
                End If
                
            End With
            
        Next i

    Else
    
        'New Board needs to be compared versus the old board to see
        'what channels can be transfered
        ResolveChanCollDependencies AssignedChanCol, _
                                    OldBoard.AInChannels, _
                                    NewBoard.AInChannels
                                    
        ResolveChanCollDependencies AssignedChanCol, _
                                    OldBoard.AOutChannels, _
                                    NewBoard.AOutChannels
                                    
        ResolveChanCollDependencies AssignedChanCol, _
                                    OldBoard.DInChannels, _
                                    NewBoard.DInChannels
                                    
        ResolveChanCollDependencies AssignedChanCol, _
                                    OldBoard.DOutChannels, _
                                    NewBoard.DOutChannels
                                    
    End If
                                    

End Sub

Public Sub ResolveChanCollDependencies(ByRef AssignedChanCol As Channels, _
                                   ByRef OldChanCol As Channels, _
                                   ByRef NewChanCol As Channels)
                                   
    Dim N As Long
    Dim M As Long
    Dim i As Long
    Dim j As Long
    
    
    'Get the number of channels in the old channel collection
    N = OldChanCol.Count
    
    If NewChanCol Is Nothing Then
    
        'OldChanCol is being deleted
        'Run ResolveChannelDependencies on every channel in OldChanCol
        'in delete mode
        
        If N <= 0 Then
        
            'No channel dependencies to worry about,
            Exit Sub
            
        End If
        
        For i = 1 To N
        
            ResolveChannelDependencies AssignedChanCol, _
                                       OldChanCol(i), _
                                       Nothing
                                       
        Next i
        
    Else
    
        'Get the number of channels in the new channel collection
        M = NewChanCol.Count
        
        'If both collections are the same size,
        'Run ResolveChannelDependencies on all of the channel pairings in edit mode
        If M = N Then
        
            For i = 1 To N
            
                ResolveChannelDependencies AssignedChanCol, _
                                           OldChanCol(i), _
                                           NewChanCol(i)
                                           
            Next i
            
        Else
        
            'The Channel collections are two different sizes
            'Again, run ResolveChannelDependencies on every unique pairing
            'Using a double for loop
            For i = 1 To N
            
                For j = 1 To M
                
                    If OldChanCol(i).BoardName = NewChanCol(j).BoardName And _
                       OldChanCol(i).ChanName = NewChanCol(j).ChanName And _
                       OldChanCol(i).ChanType = NewChanCol(j).ChanType _
                    Then
                    
                        'We have a match, nothing needs to be done
                        'Just need to exit the for loop
                        j = M + 10
                        
                    Else
                    
                        'The old and new channels don't match, do nothing
                        
                    End If
                    
                    'Check to see if there was a match
                    If j <> M + 10 Then
                
                    'No match was found for OldChanCol(i)
                    'If will be deleted in the edit process
                    ResolveChannelDependencies AssignedChanCol, _
                                               OldChanCol(i), _
                                               Nothing
                    End If
                                               
                                               
                Next j
     
            Next i
                
        End If
            
    End If
                                       
End Sub

Public Sub ResolveChannelDependencies(ByRef AssignedChanCol As Channels, _
                                      ByRef OldChan As Channel, _
                                      ByRef NewChan As Channel)
                                          
    Dim N As Long
    Dim i As Long
                                          
    'Get the number of assigned channels in the collection
    N = AssignedChanCol.Count
                                          
    'Check to see if a new channel object was inputed
    If NewChan Is Nothing Then
    
        'No new channel, therefore need to mark the damage
        'directly in thein the Assigned Channels collection
        'that will done by the Channel deletion
        
        'Iterate through the Assigned Channels collection
        'and change the Channel Name and Parent Board Name
        'of any channel that matches the inputed old channel
        'that is going to be deleted to "ERROR"
        For i = 1 To N
        
            With AssignedChanCol(i)
            
                If .ChanName = OldChan.ChanName And _
                   .ChanType = OldChan.ChanType And _
                   .BoardName = OldChan.BoardName _
                Then
                
                    'We have a match
                    'Change the Channel Name and Board Name to "ERROR"
                    .ChanName = "ERROR"
                    .BoardName = "ERROR"
                    
                End If
                
            End With
            
        Next i
        
    Else
    
        'New Chan is not nothing,
        'need to check and see how similar the new and old channels are
            
        'Iterate through the assigned channels
        For i = 1 To N
        
            With AssignedChanCol(i)
            
                If .ChanName = OldChan.ChanName And _
                   .ChanType = OldChan.ChanType And _
                   .BoardName = OldChan.BoardName _
                Then
                
                    'If new and old channel share the same channel type,
                    'then can supplant the old channel name and board name
                    'with the new channel name and board name
                    'for the matching assigned channel object
                    If NewChan.ChanType = OldChan.ChanType Then
                    
                        .ChanName = NewChan.ChanName
                        .BoardName = NewChan.BoardName
                        
                    End If
                    
                End If
                
            End With
            
        Next i
        
    End If
                                      
End Sub

Public Sub RVFFT(ByRef DataArray() As Double, ByVal N As Long)

    'RVFFT - Real Values Split-Radix Fast Fourier Transform
    'Original algorythm publised in Appendix of:
    'Sorensen HV, Jones DL, Heideman MT, Burrus CS, "Real-Valued Fast Fourier
    'Transform Algorithms".  IEEE Transactions on Acoustics, Speech and Signal
    'Processing, ASSP-35(6), June 1987.
    '
    'Translated to Visual Basic 6 by:
    'I. Hilburn, California Insitute of Technology, September 2009.

    'Inputs:
    'DataArray = array with N elements, each element of double type (or single)
    'N = length of DataArray and N = 2^M, where M is a positive whole number
    'DataArray contains Real valued time data to be FFT'd
    'DataArray = (x(t0),x(t1),x(t2),...,x(tN-1))
    'where t0,...,tN-1 are the time values of the data points
    'and Delta-T is constant, i.e. N * (t_j - t_k) = tN-1 - t0, where j - k = 1, j <= N-1
    
    'Output, written over the inputed DataArray:
    'DataArray = (X-real(F0),X-real(F1),...,X-real(FN/2-1),X-imag(F0),X-imag(F1),...,X-imag(FN/2-1))
    'where F0,...,FN/2-1 are the freq values of the FFT's data
    'and X-real, and X-imag are the real and imag portions of the FFT results
    'also: FN/2-1 == Nyquist Frequency (I think...)

    'N = Number of elements in DataArray, power of 2
    Dim M As Long 'N = 2^M elements, M = log-base-2(N)
    
    M = CLng(Log(N) / Log(2))
    
    If 2 ^ M <> N Then
    
        'Number of elements contained in DataArray is not a power
        'of two.  This should never happen. Whole application will end now.
        MsgBox "Number of elements in array submitted to the sub-routine: " & _
                "RVFFT() is not equal to a power of 2. (i.e. Num elements <> 2^M, " & _
                "where M is a positive whole number)" & vbNewLine & _
                "Number of elements = " & str(N) & vbNewLine & _
                "Corresponding power of 2 = " & str(Log(N) / Log(2)) & vbNewLine & _
                vbNewLine & "The whole program will end now!"
        
        End
        
    End If
    
    'Three counters for nested For loops
    Dim i As Long 'For loop counter
    Dim j As Long 'For loop counter
    Dim k As Long 'For Loop Counter
    
    'For loop dynamic start and step value variables
    Dim iStart As Long 'Start value of For loop
    Dim IStep As Long 'Step increment value of For loop
    
    'Temp variable for storing an array element of DataArray for transfering
    'Values from one element to another (mostly for bit reverse sorting)
    Dim DataTemp As Double
    
    Dim i_minus1 As Long 'i - 1
    Dim j_minus1 As Long 'j - 1
    Dim k_minus1 As Long 'k - 1
    Dim Nminus1 As Long 'N - 1
    Dim Ndiv2 As Long 'N \ 2 (\ = integer division)
    
    Dim N2 As Long 'Dynamic doubling variable for executing 1st butterfly / split
    Dim N4 As Long 'Dynamic var - cut into fourths, for 2nd sized partition
    Dim N8 As Long 'Dynamic var - cut into eighths, for 3rd sized partition
    
    'Variables to store array element indices from eight different
    'elements in the array (in the prior butterfly/split section)
    Dim I1 As Long
    Dim I2 As Long
    Dim I3 As Long
    Dim I4 As Long
    Dim I5 As Long
    Dim I6 As Long
    Dim I7 As Long
    Dim I8 As Long
    
    'Temp variables for multiplying array elements referenced by I1 - I8
    'with cosine and sine functions and adding the results of those multiplications
    'together before writing the assembled results to the elements of the
    'DataArray belonging to the current butterfly / split section
    Dim t1 As Double
    Dim T2 As Double
    Dim T3 As Double
    Dim T4 As Double
    Dim T5 As Double
    Dim T6 As Double
    
    '2*PI divided by N2 - partitionons angle in sine and cosine function - breaks freq. space into
    'appropriately sized segments for each butterfly
    Dim E As Double  'Adjusted in outermost for loop
    Dim a As Double  'Dynamic var, based off of E and adjusted in inner level loops
    Dim A3 As Double 'A * 3
    
    'Sine and Cosine value holder vars
    Dim SS1 As Double 'Sine(A)
    Dim SS3 As Double 'Sine(A * 3)
    Dim CC1 As Double 'Cosine(A)
    Dim CC3 As Double 'Cosine(A * 3)
    
    
'NOTE: RVFFT routine originally written in code with an index of 1 = first index in
' an array.  This is not so in VB6.  However, to keep the simplicity of the
' power-of-2 based integer division that allows this algorythm to do it's magic
' we need to keep the counters on a 1 as first-element system.  Which means,
' when we access the array elements, we need to subtract 1, sometimes, from
' the for loop counters.  Hence the variables i_minus1, j_minus1, and k_minus1
' the definitions for I1 - I8 have also been adjusted to implement this off-by-one
' array index conversion

'--------------------------------Do Bit Reverse sorting----------------------
    
    j = 1
    Nminus1 = N - 1
    Ndiv2 = N \ 2   'Using \ integer division operator
    For i = 1 To Nminus1 Step 1
        
        i_minus1 = i - 1
        j_minus1 = j - 1
        
        If i < j Then
        
            DataTemp = DataArray(j_minus1)
            DataArray(j_minus1) = DataArray(i_minus1)
            DataArray(i_minus1) = DataTemp
            
        End If
        
        k = Ndiv2
        
        Do While k < j
        
            j = j - k
            k = k \ 2  'Again, using \ integer division operator instead of /
            
        Loop
        
        j = j + k

    Next i
    
'--------------------------------Length Two Butterflies----------------------
    
    'Note: for length two butterflies
    
    iStart = 1
    IStep = 4
    
    Do While IStep < N
        
        For i = iStart To N Step IStep
            
            i_minus1 = i - 1
            DataTemp = DataArray(i_minus1)
            DataArray(i_minus1) = DataTemp + DataArray(i)
            DataArray(i) = DataTemp - DataArray(i)
            
        Next i
                
        iStart = 2 * IStep - 1
        IStep = 4 * IStep
        
    Loop
    
'--------------------------------L Shaped Butterflies------------------------
    N2 = 2
    For k = 2 To M Step 1
    
        N2 = N2 * 2
        N4 = N2 \ 4     'Using \ instead of /
        N8 = N2 \ 8     'Using \ instead of /
        
        E = 2 * Pi / N2 'Using / floating division operator, E is a double
        
        iStart = 0      'Don't be fooled - the zero value here is intended in the
                        'original algorithm to be one less than the first array index
                        'still need to handle off-by-one conversion
        IStep = N2 * 2
            
        Do While IStep < N
        
            For i = iStart To Nminus1 Step IStep
                
                I1 = i          'In original code, I1 = i + 1
                I2 = I1 + N4
                I3 = I2 + N4
                I4 = I3 + N4
                
                t1 = DataArray(I4) + DataArray(I3)
                
                DataArray(I4) = DataArray(I4) - DataArray(I3)
                DataArray(I3) = DataArray(i) - t1
                DataArray(i) = DataArray(i) + t1
                
                If N4 <> 1 Then
                
                    I1 = I1 + N8
                    I2 = I2 + N8
                    I3 = I3 + N8
                    I4 = I4 + N8
                    
                    t1 = (DataArray(I3) + DataArray(I4)) / Sqr(2)
                    T2 = (DataArray(I3) - DataArray(I4)) / Sqr(2)
                    
                    DataArray(I4) = DataArray(I2) - t1
                    DataArray(I3) = -DataArray(I2) - t1
                    DataArray(I2) = DataArray(I1) - T2
                    DataArray(I1) = DataArray(I1) + T2
                    
                End If
                
            Next i
            
            iStart = 2 * IStep - N2
            IStep = 4 * IStep
        
        Loop
            
        a = E
            
        For j = 2 To N8 Step 1
        
            A3 = 3 * a
            
            CC1 = Cos(a)
            SS1 = Sin(a)
            CC3 = Cos(A3)
            SS3 = Sin(A3)
            
            a = j * E
            
            iStart = 0
            IStep = 2 * N2
            
            j_minus1 = j - 1
            
            Do While IStep < N
            
                For i = iStart To Nminus1 Step IStep
                
                    i_minus1 = i - 1
                    
                    I1 = i_minus1 + j_minus1  'In Original algorithm, I1 = i + j
                    I2 = I1 + N4
                    I3 = I2 + N4
                    I4 = I3 + N4
                    'In original algorithm, I5 = i + N4 - j + 2
                    I5 = i_minus1 + N4 - j_minus1 + 2
                    I6 = I5 + N4
                    I7 = I6 + N4
                    I8 = I7 + N4
                    
                    t1 = DataArray(I3) * CC1 + DataArray(I7) * SS1
                    T2 = DataArray(I7) * CC1 - DataArray(I3) * SS1
                    T3 = DataArray(I4) * CC3 + DataArray(I8) * SS3
                    T4 = DataArray(I8) * CC3 - DataArray(I4) * SS3
                    
                    T5 = t1 + T3
                    T6 = T2 + T4
                    T3 = t1 - T3
                    T4 = T2 - T4
                    
                    T2 = DataArray(I6) + T6
                    DataArray(I3) = T6 - DataArray(I6)
                    DataArray(I8) = T2
                    
                    T2 = DataArray(I2) - T3
                    DataArray(I7) = -DataArray(I2) - T3
                    DataArray(I4) = T2
                    
                    t1 = DataArray(I1) + T5
                    DataArray(I6) = DataArray(I1) - T5
                    DataArray(I1) = t1
                    
                    t1 = DataArray(I5) + T4
                    DataArray(I5) = DataArray(I5) - T4
                    DataArray(I2) = t1
                    
                Next i
                
                iStart = 2 * IStep - N2
                IStep = 4 * IStep
                
            Loop
        
        Next j
            
    Next k
        
End Sub

Public Sub SaveGridToArray(ByRef gridobj As MSHFlexGrid, _
                       ByRef DataArray() As String, _
                       ByVal FirstRow As Long, _
                       ByVal LastRow As Long, _
                       Optional ByVal FirstCol As Long = 0, _
                       Optional ByVal LastCol As Long = 0)

    'Take the text values stored in a particular row & col
    'span (N x M) in the MSHFlexGrid object and store them to a
    'correspondingly size (N x M) array

    Dim i As Long
    Dim j As Long
    Dim N As Long
    Dim M As Long
    Dim TempL As Long
    
    'Get the number of rows to save
    N = LastRow - FirstRow
    
    'If N < 0, flip the sign and swap first and last row
    If N < 0 Then
    
        N = -1 * N
        TempL = FirstRow
        FirstRow = LastRow
        LastRow = TempL
        
    End If
    
    'Get the number of columns to save
    M = LastCol - FirstCol
    
    'If M < 0, flip the sign and swap first and last col
    If M < 0 Then
    
        M = -1 * M
        TempL = FirstCol
        FirstCol = LastCol
        LastCol = TempL
        
    End If
    
    
    'Resize DataArray based to (N + 1) x (M + 1) dimensions
    ReDim DataArray(N + 1, M + 1)
    
    'need to access grid object now
    With gridobj
    
        'Loop from the first row to the last
        For i = FirstRow To LastRow
    
            'set the active row in the grid
            .row = i
            
            'now loop through the columns
            For j = FirstCol To LastCol
            
                'set the active col in the grid
                .Col = j
                
                'Save the value to the corresponding element in DataArray
                DataArray(i - FirstRow, j - FirstCol) = .text
                
            Next j
            
        Next i
        
    End With
            
End Sub

'This subroutine computes the sine curve fit for a vector of data values: Y_in(),
'and a given vector of time steps: Time()
'Inputs:
'         Y_in()    -   Input signal to be fitted (Y(0),...,Y(N-1)) where N = UBound(Y_in)
'         TimeStep  -   Time interval between each point in Y_in(), assumes time interval
'                       is constant for all points in Y_in()
'         FreqEst   -   Signal frequency (actual or estimate) in Hz
'
'Outputs:
'           NOTE: Outputs are returned by altering the variable references of the
'                 last four arguments to this function.  Therefore, these variables
'                 need to be created outside of the sine fit function and inputed into it
'
'     FitParams()   -   Array with four elements, containing:
'                       [ Y-value offset of sine fit,
'                         Amplitude of sine fit,
'                         Freq(Hz) of sine fit,
'                         Phase(Rad) of sine fit ]
'
'     Y_est()       -   Estimated Sinusoid:
'                       Y-offset + Amplitude * Sine(2 * PI * Freq + PhaseShift)
'
'     Y_res()       -   Y_in() - Y_est(), the residual to the sine curve fit
'
'     RMS           -   Root mean square of the residual, Y_res()
'
'   Adapted from MatLab code for an IEEE standard for Digitizing Waveform Recorders
'     (IEEE Std 1057)
'   By: Isaac Hilburn, Jan. 2010
Public Sub SineFit(ByRef Y_in() As Double, _
                    ByVal TimeStep As Double, _
                    ByVal FreqEst As Double, _
                    ByRef FitParams() As Double, _
                    ByRef Y_est() As Double, _
                    ByRef Y_res() As Double, _
                    ByRef RMS As Double) ',
                    'ByRef SineStream As TextStream)
                    
    Dim TOL As Double
    Dim MTOL As Double
    Dim Max_Func As Integer
    Dim Max_Iter As Integer
    Dim i As Long
    Dim N As Long
    Dim SumY_res2 As Double         'Sum of the square of each elements of Y_res
    Dim T() As Double
    Dim W As Double                 'Freq in rad / s units
    
    Dim iter As Integer             'Number of iterations taken to create the sine fit
                                    'for each call of 4-param sine fit function
    Dim iter_total As Integer       'Total number of iterations
    Dim func_iter As Integer        'Number of times 4-param sine fit function has been called
    
    'Convergence factors = 4 element array containing:
    '   A0, B0, Y-Offset and delta-w (freq step in 1/rad) for each sine fit
    '   where:
    '       Y_est() = Y-Offset + A0 * cos(w * T()) + B0 * sine(w * T())
    '       Phase = Atan2(-B0/A0)
    '
    Dim ConvFactors(4) As Double
    
    TOL = 0.00000022204 'Normalized initial fit tolerance
    MTOL = 10         'TOL relaxation factor, MTOL > 1 (or else code explodes)
    Max_Func = 2       'Maximum number of times the 4 parameter fit function can be called
                        'to complete the sine curve fit
    Max_Iter = 10       'Maximum number of fit iterations per each 4 parameter fit function call
    
    'Initialize iteration variables to 0
    iter = 0
    iter_total = 0
    func_iter = 0
    
    'Initialize the sum of the square of the residual to Zero
    SumY_res2 = 0
        
    N = UBound(Y_in)    'N = length of signal input vector
    
    'Redimension and Populate T() with the appropriate time values using the
    'input variable TimeStep
    ReDim T(N)          'T() now is the same length as Y_in
    ReDim Y_est(N)      'Make sure the Y estimate and residual arrays are also the same length
    ReDim Y_res(N)
    
    'Set w = 2 * Pi * Freq Estimate
    W = 2 * Pi * FreqEst
    
'    SineStream.WriteLine "Inside SineFit"
'    SineStream.WriteLine "N = " & Trim(Str(N))
'    SineStream.WriteLine "w = " & Trim(Str(w))
'    SineStream.WriteLine "Time Step = " & Trim(Str(TimeStep))
'    SineStream.WriteBlankLines (1)
'    SineStream.WriteLine "N,T(N),Data(N)"
    
    For i = 0 To N - 1
    
        T(i) = i * TimeStep
'        SineStream.WriteLine Trim(Str(i)) & "," & _
'                             Trim(Str(T(i))) & "," & _
'                             Trim(Str(Y_in(i)))
        
    Next i
   
    'Now call first run of 4-parameter sine fit function
    SineFit4Param Y_in(), _
                    T(), _
                    N, _
                    TimeStep, _
                    W, _
                    TOL, _
                    Max_Iter, _
                    ConvFactors(), _
                    iter ',
                    'SineStream
                    
    'Update iteration variables
    iter_total = iter_total + iter
    func_iter = func_iter + 1
                    
    'If the number of iterations used in SineFit4ParamDebug > Max_Iter, then
    'the first sine fit was unsuccessful
    If iter > Max_Iter Then
    
        'Need to run SineFit again
        Do While iter > Max_Iter And func_iter <= Max_Func
        
            'Increase the tolerance level of the function
            TOL = TOL * MTOL
            
           
            'Now call 4-parameter sine fit function again
            SineFit4Param Y_in(), _
                            T(), _
                            N, _
                            TimeStep, _
                            W, _
                            TOL, _
                            Max_Iter, _
                            ConvFactors(), _
                            iter ',
                            'SineStream
                            
                            
            'Update iteration variables
            iter_total = iter_total + iter
            func_iter = func_iter + 1
            
        Loop
        
        'Test to see if the 4 parameter sine fit ever actually converged on a solution
'        If iter > MaxIter Then
'
'            'No convergence, raise an Error
'            Err.Raise 616, _
'                        "SineFit", _
'                        "Fit algorithm not able to converge upon a sine wave" & _
'                        "function for the inputed signal data."
'
'            Exit Sub
'
'        End If
        
    End If
    
    'the 4-param sine fit has converged on a solution!
    'Or exhausted itself trying.
    'Load fit parameters into FitParams() array

    'First fit parameter is the Y-offset of the sine fit
    FitParams(0) = ConvFactors(2)
    
    'Second fit parameter is the amplitude of the sine fit
    FitParams(1) = Sqr((ConvFactors(0)) ^ 2 + (ConvFactors(1)) ^ 2)
    
    'Third fit parameter is the frequency in Hz of the sine fit
    FitParams(2) = ConvFactors(3) / (2 * Pi)

    'Fourth fit parameter is the phase in radians of the sine fit
    FitParams(3) = Atn2(-ConvFactors(0), ConvFactors(1))

    'Now load results into Y_est() variable and get residuals at the same time
    For i = 0 To N - 1

        'Calculate the i-th element of Y_est()
        Y_est(i) = CSng(ConvFactors(2) _
                        + ConvFactors(0) * Cos(ConvFactors(3) * T(i)) _
                        + ConvFactors(1) * Sin(ConvFactors(3) * T(i)))

        'Find the i-th element of Y_res from the difference of Y_in(i) and Y_est(i)
        Y_res(i) = CSng(Y_in(i) - Y_est(i))

        'Add the square of the i-th element of Y_res() to the Sum of each element squared
        SumY_res2 = SumY_res2 + (Y_res(i)) ^ 2

    Next i

    'Now calculate the RMS value
    RMS = Sqr(SumY_res2 / N)
                    
End Sub

'Three Parameter Sine Fit algorithm
'Fits sine wave with a known freq, but unknown offset, amplitude, and phase shift
'
'Adapted from MatLab code for IEEE Std 1057 sine fit algorithm
'By:    Isaac Hilburn, Jan. 2010
'
'Inputs:
'   Y_in()          - N x 1 vector of data values to be fit
'   T()             - N x 1 vector of corresponding time values for each element of Y_in()
'   N               - Number of data points, size of Y_in() and T()
'   w               - Frequency to be used for making the fit
'Output:
'   ConvFactors()   - 4 x 1 vector:
'                       {A0, B0, Y-Offset, Delta-w}
'                      4 elements that uniquely describe the sine fit to the data
'                     where Y_fit = Y-offset + A0 * cos( w * T ) + B0 * sine( w * T )
'                     The fourth element, Delta-w, is not used by the 3-parameter sine fit
'                     and that fourth element is unaffected by this algorithm

Public Sub SineFit3Param(ByRef Y_in() As Double, _
                        ByRef T() As Double, _
                        ByVal N As Long, _
                        ByVal W As Double, _
                        ByRef ConvFactors() As Double) ',
                        'ByRef SineStream As TextStream)
                
    Dim D0() As Double         'Main array for least squares inversion
    ReDim D0(N, 3)
    
    Dim D0_x_Yin(3) As Double    'D0 * Y_in (should be a 3 x 1 vector as a result)
    Dim D0T_x_D0(6) As Double    'Transpose(D0) * D0 - 6x1 vector representation of
                                ' the upper triangle of a 3 x 3 hermetian matrix
                                'that will be inverted to solve the system of
                                'linear equations for the approx. least squares
                                'solution to a sine wave with a KNOWN frequency
    Dim i As Long
    Dim j As Long
    Dim success As Boolean
    
'    SineStream.WriteBlankLines (1)
'    SineStream.WriteLine "In SineFit3Param"
'    SineStream.WriteLine "w = " & Trim(Str(w))
'    SineStream.WriteLine "N = " & Trim(Str(N))
'    SineStream.WriteBlankLines (1)
'    SineStream.WriteLine "D0(N,3)"
    
    'Set elements of D0_x_Tin and first three ConvFactors to Zero
    For i = 0 To 2
    
        D0_x_Yin(i) = 0
        ConvFactors(i) = 0
        
    Next i
    
    'Set elements of D0T_x_D0 to zero
    For i = 0 To 5
    
        D0T_x_D0(i) = 0
            
    Next i
    
    
    'Calculate and load values into D0 such that:
    '
    '           |   cos( w * T(0) )     sine( w * T(0) )    1   |
    '   D0  =   |       :                   :               :   |
    '           |       :                   :               :   |
    '           |   cos( w * T(N-1) )   sine( w * T(N-1) )  1   |
    
    For i = 0 To N - 1
    
        'Set elements of D0
        D0(i, 0) = Cos(W * T(i))
        D0(i, 1) = Sin(W * T(i))
        D0(i, 2) = 1
        
'        SineStream.WriteLine Trim(Str(i)) & "," & _
'                            Trim(Str(D0(i, 0))) & "," & _
'                            Trim(Str(D0(i, 1))) & "," & _
'                            Trim(Str(D0(i, 2)))
        
        For j = 0 To 2
        
            'Multiply and sum elements of transpose(D0) and Y_in
            D0_x_Yin(j) = D0_x_Yin(j) + D0(i, j) * Y_in(i)
            
        Next j
        
        'Now sum up three diagonal elements of D0T_x_D0
        D0T_x_D0(0) = D0T_x_D0(0) + (D0(i, 0)) ^ 2
        D0T_x_D0(3) = D0T_x_D0(3) + (D0(i, 1)) ^ 2
        D0T_x_D0(5) = D0T_x_D0(5) + (D0(i, 2)) ^ 2
        
        
        'Now sum up three unique non-diagonal elements in the upper-triangle\
        'Since the matrix is hermitian, these are the only elements that we need
        D0T_x_D0(1) = D0T_x_D0(1) + D0(i, 0) * D0(i, 1)
        D0T_x_D0(2) = D0T_x_D0(2) + D0(i, 0) * D0(i, 2)
        D0T_x_D0(4) = D0T_x_D0(4) + D0(i, 1) * D0(i, 2)
        
    Next i
    
'    SineStream.WriteBlankLines (1)
'    SineStream.WriteLine ("Hermitian Matrix")
'    SineStream.WriteLine Trim(Str(D0T_x_D0(0))) & "," & _
'                        Trim(Str(D0T_x_D0(1))) & "," & _
'                        Trim(Str(D0T_x_D0(2)))
'    SineStream.WriteLine "0," & _
'                        Trim(Str(D0T_x_D0(3))) & "," & _
'                        Trim(Str(D0T_x_D0(4)))
'    SineStream.WriteLine "0,0," & _
'                        Trim(Str(D0T_x_D0(5)))
        
    
    'D0T_x_D0 is now a Hermetian matrix (positive-real symmetric)
    'Can easily invert it
    Invert3x3HermitianMatrix D0T_x_D0()
    
'    SineStream.WriteBlankLines (1)
'    SineStream.WriteLine ("Inverse Hermitian Matrix")
'    SineStream.WriteLine Trim(Str(D0T_x_D0(0))) & "," & _
'                        Trim(Str(D0T_x_D0(1))) & "," & _
'                        Trim(Str(D0T_x_D0(2)))
'    SineStream.WriteLine "0," & _
'                        Trim(Str(D0T_x_D0(3))) & "," & _
'                        Trim(Str(D0T_x_D0(4)))
'    SineStream.WriteLine "0,0," & _
'                        Trim(Str(D0T_x_D0(4)))
    
    'Now need to multiply the inverse matrix of D0T_x_D0 with D0_x_Yin
    'Typing this out brute force to double check it
    'Also, flip hermitian elements around as the lower triangle of inverse
    'is all zero
    ConvFactors(0) = D0T_x_D0(0) * D0_x_Yin(0) + _
                    D0T_x_D0(1) * D0_x_Yin(1) + _
                    D0T_x_D0(2) * D0_x_Yin(2)
                    
    ConvFactors(1) = D0T_x_D0(1) * D0_x_Yin(0) + _
                    D0T_x_D0(3) * D0_x_Yin(1) + _
                    D0T_x_D0(4) * D0_x_Yin(2)
    
    ConvFactors(2) = D0T_x_D0(2) * D0_x_Yin(0) + _
                    D0T_x_D0(4) * D0_x_Yin(1) + _
                    D0T_x_D0(5) * D0_x_Yin(2)
    
'    SineStream.WriteBlankLines (1)
'    SineStream.WriteLine "Convergence Factors"
'    SineStream.WriteLine Trim(Str(ConvFactors(0))) & "," & _
'                        Trim(Str(ConvFactors(1))) & "," & _
'                        Trim(Str(ConvFactors(2))) & "," & _
'                        Trim(Str(ConvFactors(3)))
    
    
End Sub

'Four Parameter Sine Fit algorithm
'Fits sine wave with an unknown but guessed freq,
'and unknown offset, amplitude, and phase shift
'
'Adapted from MatLab code for IEEE Std 1057 sine fit algorithm
'By:    Isaac Hilburn, Jan. 2010
'
'Inputs:
'   Y_in()          - N x 1 vector of data values to be fit
'   T()             - N x 1 vector of corresponding time values for each element of Y_in()
'   N               - Number of data points, size of Y_in() and T()
'   TimeStep        - Assumes a constant step in time between each element in T()
'                       this time step is used later to scale the Delta-w of each iteration
'                       and determine if another fit iteration should be run.
'   w               - First estimate of the frequency of the sine data
'                       This first guess if it is very far off from the correct frequency
'                       (> 1 - 2 %) can cause the 4 parameter fit to be unstable
'   TOL             - If abs(Old Convergence Factors - New Convergence Factors) > TOL, then
'                     the 4-parameter fit updates w (w_i = w_i-1 + Delta-w) and does another
'                     iteration of the 4-parameter fit
'   Max_Iter        - Maximum number of fit iterations for each function call of SineFit4Param
'
'
'Output:
'   ConvFactors()   - 4 x 1 vector:
'                       {A0, B0, Y-Offset, Delta-w}
'                      4 elements that uniquely describe the sine fit to the data
'                     where Y_fit = Y-offset + A0 * cos( w * T ) + B0 * sine( w * T )
'                     The fourth element, Delta-w, is used by update the new estimate
'                       of w used in the next iteration of the model.  The prior
'                       iterations values of A0, B0 and Y-Offset are also used in the
'                       each current iteration of the fit algorithm.  This is
'                       why the 4-parameter fit calls the 3-parameter fit to
'                       generate a 0th iteration set of values for A0, B0 and Y-offset
'                       to be used in the 1st iteration of the 4-parameter fit
'   Iter            - Integer value that contains the number of iterations of the 4-parameter
'                     fit routine that were run during this call of the SineFit4Param function

Public Sub SineFit4Param(ByRef Y_in() As Double, _
                        ByRef T() As Double, _
                        ByVal N As Long, _
                        ByVal TimeStep As Double, _
                        ByVal W As Double, _
                        ByVal TOL As Double, _
                        ByVal Max_Iter, _
                        ByRef ConvFactors() As Double, _
                        ByRef iter As Integer) ',
                        'ByRef SineStream As TextStream)
                       
    Dim D0() As Double         'Main least squares solution matrix - will be a N x 4 matrix
    ReDim D0(N, 4)
    
    Dim D0T_x_D0(10) As Double    'Hermetian matrix made from transpose(D0) * D0, matrix that needs
                                'to be inverted - 10x1 vector that contains the upper triangle
                                ' of a 4 x 4 hermetian matrix
    Dim D0_x_Yin(4) As Double   'Transpose(D0) * Y_in() - multiplied with inverse(D0T_x_D0)
                                'to solve the system of linear equations for the least squares
                                'sine fit.  4 x 1 vector
    Dim OldConvFactors(4) As Double     'Storage array for old solution parameters to the
                                        'least squares fit - used to evaluate if current
                                        'freq for fit needs to be adjusted up or down
                                        
    Dim ResFactors(4) As Double     'Change in convergence factors between iterations
                                        'of freq in fit process, with the delta-freq factor
                                        'normalized by the time-step
    Dim Max_ResFactor As Double         'Maximum of the residuals of the factors, with the
                                        'delta-freq factor normalized by the time-step
                                
    Dim success As Boolean
    Dim isError As Boolean
    Dim i As Long
    Dim j As Long
                            
    'Run 3 parameter sine fit to get first guess
    'Results loaded into ConvFactors()
    SineFit3Param Y_in(), T(), N, W, ConvFactors() ', SineStream
    
    'Set last of the ConvFactors, the freq step = zero
    ConvFactors(3) = 0
    
    'Restart iter at Zero
    iter = 0
    
    'Set flag for successful convergence to false
    success = False
    
'    SineStream.WriteBlankLines (1)
'    SineStream.WriteLine "In SineFit4Param"
'    SineStream.WriteBlankLines (1)
'    SineStream.WriteLine "D0(N,4)"
    
    
    'Convergence loop - will run 4-parameter sine fit until convergence is reached within
    'current tolerance, or loop continues through Max_Iter number of attemps
    Do While Not success
    
        'Advance the iteration counter by one
        iter = iter + 1
    
        'Iterate Freq with delta-w
        W = W + ConvFactors(3)
    
        'Set elements of D0_x_Yin to zero
        For i = 0 To 3
        
            D0_x_Yin(i) = 0
            
        Next i
        
        'Set elements of D0T_x_D0 to zero
        For i = 0 To 9
        
            D0T_x_D0(i) = 0
            
        Next i
        
        For i = 0 To N - 1
                    
            'Create D0 matrix
            D0(i, 0) = Cos(W * T(i))
            D0(i, 1) = Sin(W * T(i))
            D0(i, 2) = 1
            D0(i, 3) = -ConvFactors(0) * T(i) * D0(i, 1) _
                            + ConvFactors(1) * T(i) * D0(i, 0)
            
'            SineStream.WriteLine Trim(Str(i)) & "," & _
'                            Trim(Str(D0(i, 0))) & "," & _
'                            Trim(Str(D0(i, 1))) & "," & _
'                            Trim(Str(D0(i, 2))) & "," & _
'                            Trim(Str(D0(i, 3)))
            
            For j = 0 To 3
        
                'Multiply and sum elements of transpose(D0) and Y_in
                D0_x_Yin(j) = D0_x_Yin(j) + D0(i, j) * Y_in(i)
                
            Next j
            
            'Now sum up four diagonal elements of D0T_x_D0
            D0T_x_D0(0) = D0T_x_D0(0) + (D0(i, 0)) ^ 2
            D0T_x_D0(4) = D0T_x_D0(4) + (D0(i, 1)) ^ 2
            D0T_x_D0(7) = D0T_x_D0(7) + (D0(i, 2)) ^ 2
            D0T_x_D0(9) = D0T_x_D0(9) + (D0(i, 3)) ^ 2
        
            'Now sum up six unique non-diagonal elements in the upper-triangle
            D0T_x_D0(1) = D0T_x_D0(1) + D0(i, 0) * D0(i, 1)
            D0T_x_D0(2) = D0T_x_D0(2) + D0(i, 0) * D0(i, 2)
            D0T_x_D0(3) = D0T_x_D0(3) + D0(i, 0) * D0(i, 3)
            D0T_x_D0(5) = D0T_x_D0(5) + D0(i, 1) * D0(i, 2)
            D0T_x_D0(6) = D0T_x_D0(6) + D0(i, 1) * D0(i, 3)
            D0T_x_D0(8) = D0T_x_D0(8) + D0(i, 2) * D0(i, 3)
    
        Next i
                
'        SineStream.WriteBlankLines (1)
'        SineStream.WriteLine ("Hermitian Matrix")
'        SineStream.WriteLine Trim(Str(D0T_x_D0(0))) & "," & _
'                            Trim(Str(D0T_x_D0(1))) & "," & _
'                            Trim(Str(D0T_x_D0(2))) & "," & _
'                            Trim(Str(D0T_x_D0(3)))
'        SineStream.WriteLine "0," & _
'                            Trim(Str(D0T_x_D0(4))) & "," & _
'                            Trim(Str(D0T_x_D0(5))) & "," & _
'                            Trim(Str(D0T_x_D0(6)))
'        SineStream.WriteLine "0,0," & _
'                            Trim(Str(D0T_x_D0(7))) & "," & _
'                            Trim(Str(D0T_x_D0(8)))
'        SineStream.WriteLine "0,0,0," & _
'                            Trim(Str(D0T_x_D0(9)))
                
                
        'D0T_x_D0 is now a vector representation of the upper trianlge of a Hermetian matrix
        'Can now invert it.  After this call, D0T_x_D0 will be overwritten with it's inverse
        Invert4x4HermitianMatrix D0T_x_D0()
        
'        SineStream.WriteBlankLines (1)
'        SineStream.WriteLine ("Inverse Hermitian Matrix")
'        SineStream.WriteLine Trim(Str(D0T_x_D0(0))) & "," & _
'                            Trim(Str(D0T_x_D0(1))) & "," & _
'                            Trim(Str(D0T_x_D0(2))) & "," & _
'                            Trim(Str(D0T_x_D0(3)))
'        SineStream.WriteLine "0," & _
'                            Trim(Str(D0T_x_D0(4))) & "," & _
'                            Trim(Str(D0T_x_D0(5))) & "," & _
'                            Trim(Str(D0T_x_D0(6)))
'        SineStream.WriteLine "0,0," & _
'                            Trim(Str(D0T_x_D0(7))) & "," & _
'                            Trim(Str(D0T_x_D0(8)))
'        SineStream.WriteLine "0,0,0," & _
'                            Trim(Str(D0T_x_D0(9)))
              
        'Set Max_ResFactor = Zero
        Max_ResFactor = 0
        
        'Now need to save the old convergence factors
        For i = 0 To 3
        
            'Save old conversion factors before setting them to new ones
            OldConvFactors(i) = ConvFactors(i)
            
        Next i
            
        'Now need to multiply the inverse matrix of D0T_x_D0 with D0_x_Yin
        'Typing this out brute force to double check it
        'Also, flip hermitian elements around as the lower triangle of inverse
        'is all zero
        ConvFactors(0) = D0T_x_D0(0) * D0_x_Yin(0) + _
                        D0T_x_D0(1) * D0_x_Yin(1) + _
                        D0T_x_D0(2) * D0_x_Yin(2) + _
                        D0T_x_D0(3) * D0_x_Yin(3)
                        
        ConvFactors(1) = D0T_x_D0(1) * D0_x_Yin(0) + _
                        D0T_x_D0(4) * D0_x_Yin(1) + _
                        D0T_x_D0(5) * D0_x_Yin(2) + _
                        D0T_x_D0(6) * D0_x_Yin(3)
                        
        ConvFactors(2) = D0T_x_D0(2) * D0_x_Yin(0) + _
                        D0T_x_D0(5) * D0_x_Yin(1) + _
                        D0T_x_D0(7) * D0_x_Yin(2) + _
                        D0T_x_D0(8) * D0_x_Yin(3)
                        
        ConvFactors(3) = D0T_x_D0(3) * D0_x_Yin(0) + _
                        D0T_x_D0(6) * D0_x_Yin(1) + _
                        D0T_x_D0(8) * D0_x_Yin(2) + _
                        D0T_x_D0(9) * D0_x_Yin(3)
            
'        SineStream.WriteBlankLines (1)
'        SineStream.WriteLine "Convergence Factors"
'        SineStream.WriteLine Trim(Str(ConvFactors(0))) & "," & _
'                            Trim(Str(ConvFactors(1))) & "," & _
'                            Trim(Str(ConvFactors(2))) & "," & _
'                            Trim(Str(ConvFactors(3)))
            
        'Now go through the factors again and get the residuals
        For i = 0 To 3
            
            'Calculate the absolute values of the residuals of the new vs old factors
            ResFactors(i) = Abs(OldConvFactors(i) - ConvFactors(i))
            
            'Normalize the 4th convergence factor, Delta-w, by the Time Step
            'to allow comparison with the other factors
            If i = 3 Then ResFactors(i) = ResFactors(i) * TimeStep
            
            'Compare current ResFactor element with the past maximum factor to store the
            'bigger value of the two
            If ResFactors(i) > Max_ResFactor Then Max_ResFactor = ResFactors(i)
                        
        Next i
        
        'Now Have the maximum of the Residuals of the new vs old convergence Factors
        'Check to see if it's smaller than the tolerance
        '   --or--
        'if iter > Max_Iter
        If Max_ResFactor < TOL Or iter > Max_Iter Then
        
            success = True
            
        End If
        
    Loop
    
    ConvFactors(3) = W
       
End Sub

'Version of SineFit4Param with additional code added in to record the Convergence factors
'and the elements of the Hermitian solution matrix and it's inverse TO FILE.  This
'code is slower than SineFit4Param as it also opens a file and writes data to it.
Public Sub SineFit4ParamDebug(ByRef Y_in() As Double, _
                        ByRef T() As Double, _
                        ByVal N As Long, _
                        ByRef TimeStep As Double, _
                        ByVal W As Double, _
                        ByVal TOL As Double, _
                        ByVal Max_Iter, _
                        ByRef ConvFactors() As Double, _
                        ByRef iter As Integer, _
                        ByRef TotalTime As Double, _
                        Optional ByVal DebugFileName As String = "C:/debug.txt")
                       
    Dim D0() As Double          'Main least squares solution matrix - will be a N x 4 matrix
    Dim D0T_x_D0() As Double    'Hermetian matrix made from transpose(D0) * D0, matrix that needs
                                'to be inverted - will be a 4 x 4 hermetian matrix
    Dim Inverse() As Double     'Inverse of the above matrix
    Dim D0_x_Yin() As Double   'Transpose(D0) * Y_in() - multiplied with inverse(D0T_x_D0)
                                'to solve the system of linear equations for the least squares
                                'sine fit.  4 x 1 vector
    Dim OldConvFactors(4) As Double     'Storage array for old solution parameters to the
                                        'least squares fit - used to evaluate if current
                                        'freq for fit needs to be adjusted up or down
                                        
    Dim ResFactors(4) As Double     'Change in convergence factors between iterations
                                        'of freq in fit process, with the delta-freq factor
                                        'normalized by the time-step
    Dim Max_ResFactor As Double         'Maximum of the residuals of the factors, with the
                                        'delta-freq factor normalized by the time-step
                                
    Dim success As Boolean
    Dim isError As Boolean
    Dim i As Long
    Dim j As Long
    
'---------------Debug----------------------------------

    Dim fso As New Scripting.FileSystemObject
    Dim DebugFile As File
    Dim DebugStream As TextStream
    Dim LineText As String
    Dim time As Double
    Dim ElapsedTime As Double
    
    Set DebugFile = fso.GetFile(DebugFileName)
    Set DebugStream = DebugFile.OpenAsTextStream(ForAppending)
    
'------------------------------------------------------
                            
    'Run 3 parameter sine fit to get first guess
    'Results loaded into ConvFactors()
    time = CDbl(Now)
    
    SineFit3Param Y_in(), T(), N, W, ConvFactors()
    
    ElapsedTime = CDbl(Now) - time
    TotalTime = TotalTime + ElapsedTime
    
    'Set last of the ConvFactors, the freq step = zero
    ConvFactors(3) = 0
    
'---------------Debug----------------------------------
       
    DebugStream.WriteLine ("After 3-parameter Sine Fit:")
    DebugStream.WriteLine ("Run-time = " & Trim(str(ElapsedTime)))
    DebugStream.WriteLine ("Total-time = " & Trim(str(TotalTime)))
    DebugStream.WriteLine ("A0 = " & Trim(str(ConvFactors(0))))
    DebugStream.WriteLine ("B0 = " & Trim(str(ConvFactors(1))))
    DebugStream.WriteLine ("Y-Offset = " & Trim(str(ConvFactors(2))))
    DebugStream.WriteLine ("Freq = " & Trim(str(W / (2 * Pi))))
    DebugStream.WriteLine ("Delta-Freq = " & Trim(str(ConvFactors(3) / (2 * Pi))))
    DebugStream.WriteBlankLines (1)
    
'------------------------------------------------------

    time = CDbl(Now)

    'Redim all N x 4 matrices
    ReDim D0(N, 4)
    
    'ReDim all 4 x 4 matrices
    ReDim D0T_x_D0(4, 4)
    ReDim Inverse(4, 4)
    
    'ReDim all 4 x 1 matrices
    ReDim D0_x_Yin(4)
    
    'Restart iter at Zero
    iter = 0
    
    'Set flag for successful convergence to false
    success = False
    
    'Convergence loop - will run 4-parameter sine fit until convergence is reached within
    'current tolerance, or loop continues through Max_Iter number of attemps
    Do While Not success
    
        'Advance the iteration counter by one
        iter = iter + 1
    
        'Iterate Freq with delta-w
        W = W + ConvFactors(3)
    
        'Set elements of D0_x_Tin and ConvFactors to Zero
        For i = 0 To 2
        
            D0_x_Yin(i) = 0
            
        Next i
        
        'Set elements of D0T_x_D0 to zero
        For i = 0 To 2
        
            For j = 0 To 2
            
                D0T_x_D0(i, j) = 0
                
            Next j
            
        Next i
        
        For i = 0 To N - 1
                    
            'Create D0 matrix
            D0(i, 0) = Cos(W * T(i))
            D0(i, 1) = Sin(W * T(i))
            D0(i, 2) = 1
            D0(i, 3) = -ConvFactors(0) * T(i) * D0(i, 1) _
                      + ConvFactors(1) * T(i) * D0(i, 0)
            
            For j = 0 To 3
        
                'Multiply and sum elements of transpose(D0) and Y_in
                D0_x_Yin(j) = D0_x_Yin(j) + D0(i, j) * Y_in(i)
                
                'Now sum up diagonal elements of D0T_x_D0
                D0T_x_D0(j, j) = D0T_x_D0(j, j) + (D0(i, j)) ^ 2
                
            Next j
            
            'Now sum up six unique non-diagonal elements in the upper-triangle
            D0T_x_D0(0, 1) = D0T_x_D0(0, 1) + D0(i, 0) * D0(i, 1)
            D0T_x_D0(0, 2) = D0T_x_D0(0, 2) + D0(i, 0) * D0(i, 2)
            D0T_x_D0(0, 3) = D0T_x_D0(0, 3) + D0(i, 0) * D0(i, 3)
            D0T_x_D0(1, 2) = D0T_x_D0(1, 2) + D0(i, 1) * D0(i, 2)
            D0T_x_D0(1, 3) = D0T_x_D0(1, 3) + D0(i, 1) * D0(i, 3)
            D0T_x_D0(2, 3) = D0T_x_D0(2, 3) + D0(i, 2) * D0(i, 3)
    
        Next i
        
        'D0T_x_D0 is now a Hermetian matrix
        'Can now invert it
        Inverse = DOT_x_DO
        Invert4x4HermitianMatrix Inverse()
        
'--------------Debug----------------------------------
'        DebugStream.WriteLine ("D'D Matrix")
'
'        For i = 0 To 3
'
'            LineText = ""
'
'            For J = 0 To 3
'
'                If i > J Then
'
'                    LineText = LineText & Trim(Str(D0T_x_D0(J, I)))
'
'                Else
'
'                    LineText = LineText & Trim(Str(D0T_x_D0(I, J)))
'
'                End If
'
'                If J < 3 Then
'
'                    LineText = LineText & ","
'
'                End If
'
'            Next J
'
'            DebugStream.WriteLine (LineText)
'
'        Next I
'
'        DebugStream.WriteBlankLines (1)
'        DebugStream.WriteLine ("Inverse Matrix")
'
'        For i = 0 To 3
'
'            LineText = ""
'
'            For J = 0 To 3
'
'                If i > J Then
'
'                    LineText = LineText & Trim(Str(Inverse(J, I)))
'
'                Else
'
'                    LineText = LineText & Trim(Str(Inverse(I, J)))
'
'                End If
'
'                If J < 3 Then
'
'                    LineText = LineText & ","
'
'                End If
'
'            Next J
'
'            DebugStream.WriteLine (LineText)
'
'        Next I
'
'        DebugStream.WriteBlankLines (1)
'
'        MatrixMatrixMultiply D0T_x_D0(), 0, 3, 0, 3, False, _
'                            Inverse(), 0, 3, 0, 3, False, 1, _
'                            Temp(), 0, 3, 0, 3, 1, _
'                            Work()
'
'
'        'Check the Inversion process
'        Debug.Print "|  " & Trim(Str(Temp(0, 0))) & "  " & Trim(Str(Temp(0, 1))) & "  " _
'                            & Trim(Str(Temp(0, 2))) & "  " & Trim(Str(Temp(0, 3))) & "|" & vbNewLine & _
'                    "|  " & Trim(Str(Temp(1, 0))) & "  " & Trim(Str(Temp(1, 1))) & "  " _
'                            & Trim(Str(Temp(1, 2))) & "  " & Trim(Str(Temp(1, 3))) & "|" & vbNewLine & _
'                    "|  " & Trim(Str(Temp(2, 0))) & "  " & Trim(Str(Temp(2, 1))) & "  " _
'                            & Trim(Str(Temp(2, 2))) & "  " & Trim(Str(Temp(2, 3))) & "|" & vbNewLine & _
'                    "|  " & Trim(Str(Temp(3, 0))) & "  " & Trim(Str(Temp(3, 1))) & "  " _
'                            & Trim(Str(Temp(3, 2))) & "  " & Trim(Str(Temp(3, 3))) & "|" & vbNewLine

        'Set Max_ResFactor = Zero
        Max_ResFactor = 0
        
        'Now need to save the old convergence factors
        For i = 0 To 3
        
            'Save old conversion factors before setting them to new ones
            OldConvFactors(i) = ConvFactors(i)
            
        Next i
            
        'Now need to multiply the inverse matrix of D0T_x_D0 with D0_x_Yin
        'Typing this out brute force to double check it
        'Also, flip hermitian elements around as the lower triangle of inverse
        'is all zero
        ConvFactors(0) = Inverse(0, 0) * D0_x_Yin(0) + _
                        Inverse(0, 1) * D0_x_Yin(1) + _
                        Inverse(0, 2) * D0_x_Yin(2) + _
                        Inverse(0, 3) * D0_x_Yin(3)
                        
        ConvFactors(1) = Inverse(0, 1) * D0_x_Yin(0) + _
                        Inverse(1, 1) * D0_x_Yin(1) + _
                        Inverse(1, 2) * D0_x_Yin(2) + _
                        Inverse(1, 3) * D0_x_Yin(3)
                        
        ConvFactors(2) = Inverse(0, 2) * D0_x_Yin(0) + _
                        Inverse(1, 2) * D0_x_Yin(1) + _
                        Inverse(2, 2) * D0_x_Yin(2) + _
                        Inverse(2, 3) * D0_x_Yin(3)
                        
        ConvFactors(3) = Inverse(0, 3) * D0_x_Yin(0) + _
                        Inverse(1, 3) * D0_x_Yin(1) + _
                        Inverse(2, 3) * D0_x_Yin(2) + _
                        Inverse(3, 3) * D0_x_Yin(3)
            
        'Now go through the factors again and get the residuals
        For i = 0 To 3
            
            'Calculate the absolute values of the residuals of the new vs old factors
            ResFactors(i) = Abs(OldConvFactors(i) - ConvFactors(i))
            
            'Normalize the 4th convergence factor, Delta-w, by the Time Step
            'to allow comparison with the other factors
            If i = 3 Then ResFactors(i) = ResFactors(i) * TimeStep
            
'            Debug.Print "ResFactor(" & Trim(Str(i)) & ")= " & Trim(Str(ResFactors(i)))
            
            'Compare current ResFactor element with the past maximum factor to store the
            'bigger value of the two
            If ResFactors(i) > Max_ResFactor Then Max_ResFactor = ResFactors(i)
                        
        Next i
        
        'Now Have the maximum of the Residuals of the new vs old convergence Factors
        'Check to see if it's smaller than the tolerance
        '   --or--
        'if iter > Max_Iter
        If Max_ResFactor < TOL Or iter > Max_Iter Then
        
            success = True
            
        End If
        
        ElapsedTime = CDbl(Now) - time
        TotalTime = TotalTime + ElapsedTime
        
        '---------------Debug----------------------------------
       
            DebugStream.WriteLine ("4-param fit Iteration # = " & Trim(str(iter)))
            DebugStream.WriteLine ("Run-time = " & Trim(str(ElapsedTime)))
            DebugStream.WriteLine ("Total-time = " & Trim(str(TotalTime)))
            DebugStream.WriteLine ("A0 = " & Trim(str(ConvFactors(0))))
            DebugStream.WriteLine ("B0 = " & Trim(str(ConvFactors(1))))
            DebugStream.WriteLine ("Y-Offset = " & Trim(str(ConvFactors(2))))
            DebugStream.WriteLine ("Freq = " & Trim(str(W / (2 * Pi))))
            DebugStream.WriteLine ("Delta-Freq = " & Trim(str(ConvFactors(3) / (2 * Pi))))
            DebugStream.WriteBlankLines (1)
            
        '------------------------------------------------------
        
        
    Loop
    
    ConvFactors(3) = W
       
End Sub

'Public Sub ArrayInterpolateAndSplice(ByRef PriorArray() As Double, _
'                                     ByRef NewArray() As Double, _
'                                     Optional ByVal PreserveNumber As Boolean = True, _
'                                     Optional ByVal SpliceDown As Boolean = True, _
'                                     Optional ByVal SpliceUp As Boolean = False)
'
'    'This function takes two arrays and interpolates that second array into the first
'    'If the SpliceDown flag is true, the interpolation pulls down the higher value(s)
'    'on a line with the lower value(s), if the SpliceUp flag is true (and the SpliceDown
'    'flag is false), then the interpolation pulls up the lower value(s) on a line with the
'    'higher value(s), if the SpliceUp and SpliceDown flags are both false, or conflict, then
'    'the interpolate just plops the new point between the surrounding existing ones.
'    '
'    'If the PreserveNumber flag is true, no new points are added to the prior array,
'    'the point values in the prior array can be modified
'    'If PreserveNumber = True and no up or down splice is called for, then the
'    'least squares fit line for the two old points and the one new point is used
'    'to reposition the y-value of the two old points.  Otherwise, the splice up or down
'    'is done and the new point is not saved.
'    'If PreserverNumber = False and either splice is on, the old value is just replaced /
'    ' modified based on the new.  No new points need to be added
'
'    'PriorArray is overwritten at the end of this function with the resulting
'    'interpolate and splice with the new array.
'
'    Dim i As Long
'    Dim j As Long
'    Dim k As Long
'    Dim N As Long
'    Dim M As Long
'
'    Dim TempD As Double
'    Dim Slope As Double
'
'    Dim TempB As Boolean
'
'    'Check for nothing values in place of the arrays
'
'    'Both arrays are nothing, exit the sub
'    If (UBound(PriorArray) < 0 And UBound(NewArray) < 0) Then Exit Sub
'    'If PriorArray Is Nothing And _
'     '  NewArray Is Nothing Then Exit Sub
'
'    'PriorArray is nothing
'    If (UBound(PriorArray) < 0) Then
'
'        'Set Prior Array, element by element = New array
'        M = UBound(NewArray, 1)
'
'        If M > 0 Then
'
'            'ReDim the PriorArray
'            ReDim PriorArray(M, 2)
'
'            For i = 0 To M - 1
'
'                For j = 0 To 1
'
'                    PriorArray(i, j) = NewArray(i, j)
'
'                Next j
'
'            Next i
'
'        End If
'
'        'Exit the sub
'        Exit Sub
'
'    End If
'
'    'New array is nothing - exit the sub, there's nothing to add
'    If (UBound(NewArray) < 0) Then Exit Sub
'
'    'Get the size of the prior array
'    N = UBound(PriorArray, 1)
'
'    'Get the size of the new array
'    M = UBound(NewArray, 1)
'
'    'Now check for zero arrays
'
'    'Both empty, exit the sub
'    If N = 0 And M = 0 Then Exit Sub
'
'    'Prior Array is empty
'    If N = 0 And M > 0 Then
'
'        'Set the prior array to the new array element by element
'        'ReDim the PriorArray
'        ReDim PriorArray(M, 2)
'
'        For i = 0 To M - 1
'
'            For j = 0 To 1
'
'                PriorArray(i, j) = NewArray(i, j)
'
'            Next j
'
'        Next i
'
'        Exit Sub
'
'    End If
'
'    'New array is empty, exit the sub, there's nothing to add
'    If M = 0 Then Exit Sub
'
'    'Set the dimensions of resulting new elements that need to be saved
'    ReDim ElementsToSave(N + M, 2)
'
'    'Default all the values in Elements to save to -1
'    For i = 0 To N + M - 1
'
'        ElementsToSave(i, 0) = -1
'
'    Next i
'
'    'Start the active element of ELements to Save at zero
'    k = 0
'
'    'Now, have two non-empty arrays to work with
'    'Loop through the new array
'    For i = 0 To M - 1
'
'        'Compare each point to each point in the prior array
'        For j = 0 To N - 1
'
'            'Do the Col 0 values in both arrays equal each other?
'            If PriorArray(j, 0) = NewArray(i, 0) Then
'
'                'Compare the Col 1 values
'                'If prior value needs to be pushed down, and
'                'the new value is lower, replace the prior value with it
'                'If prior value needs to be pushed up, and
'                'the new value is higher, replace the prior value with it
'                If (SpliceDown = True And _
'                    SpliceUp = False And _
'                    NewArray(i, 1) < PriorArray(j, 1)) Or _
'                   (SpliceDown = False And _
'                    SpliceUp = True And _
'                    NewArray(i, 1) > PriorArray(j, 1)) _
'                Then
'
'                    'Save the New element
'                    ' If saved in column zero, element is from the prior array
'                    ' If saved in column one, element is from the new array
'                    ElementsToSave(k, 1) = i
'
'                    'Increment k
'                    k = k + 1
'
'                ElseIf (SpliceDown = True And _
'                        SpliceUp = False And _
'                        NewArray(i, 1) > PriorArray(j, 1)) Or _
'                       (SpliceDown = False And _
'                        SpliceUp = True And _
'                        NewArray(i, 1) < PriorArray(j, 1)) _
'                Then
'
'                    'Need to keep the prior value instead of the new value
'                    ElementsToSave(k, 0) = j
'                    k = k + 1
'
'                ElseIf SpliceDown = False And _
'                       SpliceUp = False And _
'                       NewArray(i, 1) <> PriorArray(j, 1) _
'                Then
'
'                    TempD = (PriorArray(j, 1) + NewArray(i, 1)) / 2
'                    PriorArray(j, 1) = TempD
'
'                    'Record this element as one that needs saving
'                    ' 0 = from PriorArray
'                    ' i = index to save
'                    ElementsToSave(k, 0) = j
'
'                    'Increment k
'                    k = k + 1
'
'                End If
'
'                'Exit the interior For loop
'                Exit For
'
'            End If
'
'            'If still in the loop, is this the first point in the prior array,
'            'and is the prior value > new value
'            'then need to compare the interpolated new value with the prior value
'            If j = 0 And PriorArray(j, 0) > NewArray(i, 0) Then
'
'                'Check to see if there's only one point in PriorArray
'                If N = 1 Then
'
'                    'Even if PreserveNumber = true, need to save the new value
'                    'and prior value
'                    ElementsToSave(k, 1) = i
'                    k = k + 1
'
'                    TempB = ElementAlreadySaved(0, j)
'
'                    'Check to see if the prior has already been saved
'                    If TempB = True Then
'
'                        ElementsToSave(k, 0) = j
'                        k = k + 1
'
'                    End If
'
'                    'Exit the inner for loop
'                    Exit For
'
'                Else
'
'                    'There's more than one element in the prior array
'                    'Use the 1st and 2nd elements of the prior array
'                    'to draw a line and compare the projected value to that of
'                    'the new array.
'                    Slope = (PriorArray(j + 1, 1) - PriorArray(j, 1)) / _
'                            (PriorArray(j + 1, 0) - PriorArray(j, 0))
'
'                    'Set TempD into the projected downward (in Ramp ("X") coordinates) Col1 value
'                    TempD = PriorArray(j, 1) - Slope * (PriorArray(j, 0) - NewArray(i, 0))
'
'                    'Compare TempD with the value in NewArray, is Splice up or down is true
'                    'and TempD > or < New Array and PreserveNumber = True, then replace
'                    'the Prior Array point with the New array point
'                    'If PreserveNumber = False, simply add the new array point
'                    'If TempD = New Array value, save the new array point
'                    If (TempD > NewArray(i, 1) And SpliceDown = True) Or _
'                       (TempD < NewArray(i, 1) And SpliceUp = True) _
'                    Then
'
'                        If PreserveNumber = True Then
'
'                            'Save only new point
'                            ElementsToSave(k, 1) = i
'                            k = k + 1
'
'                        Else
'
'                            TempB = ElementAlreadySaved(0, j)
'
'                            'Check to see if the prior has already been saved
'                            If TempB = True Then
'
'                                ElementsToSave(k, 0) = j
'                                k = k + 1
'
'                            End If
'
'                            'Save the new point
'                            ElementsToSave(k, 1) = i
'                            k = k + 1
'
'                        End If
'
'                        'Exit the inner for loop
'                        Exit For
'
'                    ElseIf TempD = NewArray(i, 1) Then
'
'                        'Save new array point
'                        ElementsToSave(k, 1) = i
'                        k = k + 1
'
'                        'Exit the inner for loop
'                        Exit For
'
'                    End If
'
'                End If
'
'            End If
'
'            'If still in the loop, is this the last point in the prior array,
'            'and is the prior value < new value
'            'then need to compare the interpolated new value with the prior value
'            If j = N - 1 And PriorArray(j, 0) < NewArray(i, 0) Then
'
'                'Check to see if there's only one point in PriorArray
'                If N = 1 Then
'
'                    'Even if PreserveNumber = true, need to save the new value
'                    'and prior value
'                    ElementsToSave(k, 1) = i
'                    k = k + 1
'
'                    TempB = ElementAlreadySaved(0, j)
'
'                    'Check to see if the prior has already been saved
'                    If TempB = True Then
'
'                        ElementsToSave(k, 0) = j
'                        k = k + 1
'
'                    End If
'
'                    'Exit the inner for loop
'                    Exit For
'
'                Else
'
'                    'There's more than one element in the prior array
'                    'Use the 1st and 2nd elements of the prior array
'                    'to draw a line and compare the projected value to that of
'                    'the new array.
'                    Slope = (PriorArray(j + 1, 1) - PriorArray(j, 1)) / _
'                            (PriorArray(j + 1, 0) - PriorArray(j, 0))
'
'                    'Set TempD into the projected upward (in Ramp ("X") coordinates) Col1 value
'                    TempD = PriorArray(j, 1) + Slope * (PriorArray(j, 0) - NewArray(i, 0))
'
'                    'Compare TempD with the value in NewArray, is Splice up or down is true
'                    'and TempD > or < New Array and PreserveNumber = True, then replace
'                    'the Prior Array point with the New array point
'                    'If PreserveNumber = False, simply add the new array point
'                    'If TempD = New Array value, save the new array point
'                    If (TempD > NewArray(i, 1) And SpliceDown = True) Or _
'                       (TempD < NewArray(i, 1) And SpliceUp = True) _
'                    Then
'
'                        If PreserveNumber = True Then
'
'                            'Save only new point
'                            ElementsToSave(k, 1) = i
'                            k = k + 1
'
'                        Else
'
'                            TempB = ElementAlreadySaved(0, j)
'
'                            'Check to see if the prior has already been saved
'                            If TempB = True Then
'
'                                ElementsToSave(k, 0) = j
'                                k = k + 1
'
'                            End If
'
'                            'Save the new point
'                            ElementsToSave(k, 1) = i
'                            k = k + 1
'
'                        End If
'
'                        'Exit the inner for loop
'                        Exit For
'
'                    ElseIf TempD = NewArray(i, 1) Then
'
'                        'Save new array point
'                        ElementsToSave(k, 1) = i
'                        k = k + 1
'
'                        'Exit the inner for loop
'                        Exit For
'
'                    End If
'
'                End If
'
'            End If
'
'            'Now we're done with first and last point cases
'            'The only remaining matches have new points that lie in between
'            'old points
'            If NewArray(i, 0) > PriorArray(j - 1, 0) And _
'               NewArray(i, 0) < PriorArray(j, 0) _
'            Then
'
'                'Get Slope
'                Slope = (PriorArray(j, 1) - PriorArray(j - 1, 1)) / _
'                        (PriorArray(j, 0) - PriorArray(j - 1, 0))
'
'                'Interpolate Prior Col1 value at the Col0 position of the New array element
'                TempD = PriorArray(j, 1) - Slope * (PriorArray(j, 0) - NewArray(i, 0))
'
'                'If TempD = NewArray(i,1) then, just need to save two old points
'                'if they already haven't been saved
'                End If
'
'End Sub

'Version of SineFit subroutine used for analyzing fit results afterwards - writes to
'file log of convergence factors as the code iterates towards a solution
Public Sub SineFitDebug(ByRef Y_in() As Double, _
                    ByVal TimeStep As Double, _
                    ByVal FreqEst As Double, _
                    ByRef FitParams() As Double, _
                    ByRef Y_est() As Double, _
                    ByRef Y_res() As Double, _
                    ByRef RMS As Double)
                    
    Dim TOL As Double
    Dim MTOL As Double
    Dim Max_Func As Integer
    Dim Max_Iter As Integer
    Dim i As Long
    Dim N As Long
    Dim SumY_res2 As Double         'Sum of the square of each elements of Y_res
    Dim T() As Double
    Dim W As Double                 'Freq in rad / s units
    
    Dim iter As Integer             'Number of iterations taken to create the sine fit
                                    'for each call of 4-param sine fit function
    Dim iter_total As Integer       'Total number of iterations
    Dim func_iter As Integer        'Number of times 4-param sine fit function has been called
    
    'Convergence factors = 4 element array containing:
    '   A0, B0, Offset and delta-w (freq step in 1/rad) for each sine fit
    '   where:
    '       Y_est() = Offset + A0 * cos(w * T()) + B0 * sine(w * T())
    '       Phase = Atan2(-B0/A0)
    '
    '       Y_est() = Offset + B0 * sine(w * T() + Phase)
    Dim ConvFactors(4) As Double
    
    TOL = 2.2204E-16 'Normalized initial fit tolerance
    MTOL = 4         'TOL relaxation factor, MTOL > 1 (or else code explodes)
    Max_Func = 16       'Maximum number of times the 4 parameter fit function can be called
                        'to complete the sine curve fit
    Max_Iter = 32       'Maximum number of fit iterations per each 4 parameter fit function call
    
    'Initialize iteration variables to 0
    iter = 0
    iter_total = 0
    func_iter = 0
    
    'Initialize the sum of the square of the residual to Zero
    SumY_res2 = 0
        
    N = UBound(Y_in)    'N = length of signal input vector
    
    'Redimension and Populate T() with the appropriate time values using the
    'input variable TimeStep
    ReDim T(N)          'T() now is the same length as Y_in
    
    For i = 0 To N - 1
    
        T(i) = i * TimeStep
        
    Next i
    
    
    'Check to make sure that Y_est() and Y_res() are the same length as Y_in()
    'If not, re-dimension them
    If UBound(Y_est) <> N Then
    
        ReDim Y_est(N)
        
    End If
    
    If UBound(Y_res) <> N Then
    
        ReDim Y_res(N)
        
    End If
    
    
    'Set w = 2 * Pi * Freq Estimate
    W = 2 * Pi * FreqEst
    
'------------Debug----------------------
    Dim StrDate As String
    Dim fso As New Scripting.FileSystemObject
    Dim DebugFile As File
    Dim DebugStream As TextStream
    
    StrDate = Format(Now, "MM-DD-YYYY_HH_MM")
    StrDate = "C:\Documents and Settings\lab\Desktop\Test MCC Board 11-16-2009\" & _
                "Test MCC Board\Debug_Sine_Fit" & StrDate & ".txt"
                
    'Create debug file
    fso.CreateTextFile StrDate, True
    Set DebugFile = fso.GetFile(StrDate)
    Set DebugStream = DebugFile.OpenAsTextStream(ForWriting)
        
    DebugStream.WriteLine ("Func Call # = 1")
        
    DebugStream.Close
    
    Dim TotalTime As Double
    
    TotalTime = 0
    
'---------------------------------------
    
    
    'Now call first run of 4-parameter sine fit function
    SineFit4ParamDebug Y_in(), _
                    T(), _
                    N, _
                    TimeStep, _
                    W, _
                    TOL, _
                    Max_Iter, _
                    ConvFactors(), _
                    iter, _
                    TotalTime, _
                    StrDate
                    
                    
    'Update iteration variables
    iter_total = iter_total + iter
    func_iter = func_iter + 1
                    
    'If the number of iterations used in SineFit4ParamDebug > Max_Iter, then
    'the first sine fit was unsuccessful
    If iter > Max_Iter Then
    
        'Need to run SineFit again
        Do While iter > Max_Iter And func_iter <= Max_Func
        
            'Increase the tolerance level of the function
            TOL = TOL * MTOL
            
'---------------Debug-------------------------------
            Set DebugStream = DebugFile.OpenAsTextStream(ForAppending)
            DebugStream.WriteLine ("Func Call # = " & Trim(str(func_iter)))
            DebugStream.Close
'---------------------------------------------------
            
            'Now call 4-parameter sine fit function again
            SineFit4ParamDebug Y_in(), _
                            T(), _
                            N, _
                            TimeStep, _
                            W, _
                            TOL, _
                            Max_Iter, _
                            ConvFactors(), _
                            iter, _
                            TotalTime, _
                            StrDate
                            
                            
            'Update iteration variables
            iter_total = iter_total + iter
            func_iter = func_iter + 1
            
        Loop
        
        'Test to see if the 4 parameter sine fit ever actually converged on a solution
'        If iter > MaxIter Then
'
'            'No convergence, raise an Error
'            Err.Raise 616, _
'                        "SineFit", _
'                        "Fit algorithm not able to converge upon a sine wave" & _
'                        "function for the inputed signal data."
'
'            Exit Sub
'
'        End If
        
    End If
    
    'Otherwise, the 4-param sine fit has converged on a solution!
    'Load fit parameters into FitParams() array
    
    'First fit parameter is the Y-offset of the sine fit
    FitParams(0) = ConvFactors(2)
    
    'Second fit parameter is the amplitude of the sine fit
    FitParams(1) = Sqr((ConvFactors(0)) ^ 2 + (ConvFactors(1)) ^ 2)
    
    'Third fit parameter is the frequency in Hz of the sine fit
    FitParams(2) = ConvFactors(3) / (2 * Pi)
    
    'Fourth fit parameter is the phase in radians of the sine fit
    FitParams(3) = Atn2(-ConvFactors(0), ConvFactors(1))
    
    'Now load results into Y_est() variable and get residuals at the same time
    For i = 0 To N - 1
    
        'Calculate the i-th element of Y_est()
        Y_est(i) = ConvFactors(2) _
                    + ConvFactors(0) * Cos(ConvFactors(3) * T(i)) _
                    + ConvFactors(1) * Sin(ConvFactors(3) * T(i))
                    
        'Find the i-th element of Y_res from the difference of Y_in(i) and Y_est(i)
        Y_res(i) = Y_in(i) - Y_est(i)
        
        'Add the square of the i-th element of Y_res() to the Sum of each element squared
        SumY_res2 = SumY_res2 + (Y_res(i)) ^ 2
        
    Next i

    'Now calculate the RMS value
    RMS = Sqr(SumY_res2 / N)
                    
End Sub

Public Sub SortGrid(ByRef gridobj As MSHFlexGrid, _
                    ByRef CallingForm As Form, _
                    ByVal FirstRow As Long, _
                    ByVal LastRow As Long, _
                    Optional ByVal FirstCol As Long = 0, _
                    Optional ByVal LastCol As Long = 0, _
                    Optional ByVal FirstColNumbered As Boolean = False)
                                   
    'Temp array for sorting the gridObj values
    Dim TempArray() As String
    Dim TempL As Long
    
    'Load the grid into the array
    SaveGridToArray gridobj, _
                    TempArray(), _
                    FirstRow, _
                    LastRow, _
                    FirstCol, _
                    LastCol
                    
    'Now sort the array using a recursive function
    MedianThreeQuickSort1_2D TempArray(), 0
    
    'Now reload the temparray to the gridObj
    LoadArrayToGrid gridobj, _
                    TempArray(), _
                    CallingForm, _
                    FirstRow, _
                    FirstCol
                    
    If FirstColNumbered = True Then
    
        RenumberGridCol gridobj, _
                        CallingForm, _
                        0, _
                        1, _
                        FirstRow, _
                        LastRow, _
                        False
        
    End If
                    
End Sub

