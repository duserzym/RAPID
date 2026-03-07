VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmDAQBoardSettings 
   Caption         =   "DAQ Board Settings"
   ClientHeight    =   6495
   ClientLeft      =   1185
   ClientTop       =   2895
   ClientWidth     =   9255
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6495
   ScaleWidth      =   9255
   Begin VB.CommandButton cmdEditBoard 
      Caption         =   "Edit Board"
      Height          =   372
      Left            =   1560
      TabIndex        =   14
      Top             =   6000
      Width           =   972
   End
   Begin VB.Frame frameChannels 
      Caption         =   "Channels"
      Height          =   5772
      Left            =   4080
      TabIndex        =   6
      Top             =   120
      Width           =   5052
      Begin VB.CommandButton cmdEditChan 
         Caption         =   "Edit..."
         Height          =   372
         Left            =   2040
         TabIndex        =   15
         Top             =   5160
         Width           =   1092
      End
      Begin VB.CommandButton cmdAddChan 
         Caption         =   "Add..."
         Height          =   372
         Left            =   3480
         TabIndex        =   13
         Top             =   5160
         Width           =   1092
      End
      Begin VB.CommandButton cmdDeleteChan 
         Caption         =   "Delete..."
         Height          =   372
         Left            =   480
         TabIndex        =   12
         Top             =   5160
         Width           =   1212
      End
      Begin VB.ComboBox cmbChanType 
         Height          =   288
         Left            =   2880
         TabIndex        =   10
         Top             =   600
         Width           =   1932
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridDAQChannels 
         Height          =   3972
         Left            =   240
         TabIndex        =   9
         Top             =   1080
         Width           =   4572
         _ExtentX        =   8070
         _ExtentY        =   7011
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.ComboBox cmbDAQBoard 
         Height          =   288
         Left            =   240
         TabIndex        =   7
         Top             =   600
         Width           =   2172
      End
      Begin VB.Label Label2 
         Caption         =   "Channel Type:"
         Height          =   252
         Left            =   2880
         TabIndex        =   11
         Top             =   360
         Width           =   1452
      End
      Begin VB.Label Label1 
         Caption         =   "DAQ Board:"
         Height          =   252
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   1332
      End
   End
   Begin VB.CommandButton cmdAddBoard 
      Caption         =   "Add Board"
      Height          =   372
      Left            =   2880
      TabIndex        =   5
      Top             =   6000
      Width           =   1092
   End
   Begin VB.CommandButton cmdDeleteBoard 
      Caption         =   "Delete Board"
      Height          =   372
      Left            =   120
      TabIndex        =   4
      Top             =   6000
      Width           =   1092
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   372
      Left            =   6120
      TabIndex        =   3
      Top             =   6000
      Width           =   1092
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   372
      Left            =   4440
      TabIndex        =   2
      Top             =   6000
      Width           =   1092
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Default         =   -1  'True
      Height          =   372
      Left            =   7800
      TabIndex        =   1
      Top             =   6000
      Width           =   1092
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridDAQBoards 
      Height          =   5772
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3852
      _ExtentX        =   6800
      _ExtentY        =   10186
      _Version        =   393216
      SelectionMode   =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Menu mnuDAQBoardSettings 
      Caption         =   "DAQ Board Settings"
      Index           =   0
      Visible         =   0   'False
      Begin VB.Menu mnuDAQDelete 
         Caption         =   "&Delete"
         Index           =   1
      End
      Begin VB.Menu mnuDAQAdd 
         Caption         =   "&Add"
         Index           =   2
      End
      Begin VB.Menu mnuDAQEdit 
         Caption         =   "&Edit"
         Index           =   3
      End
   End
End
Attribute VB_Name = "frmDAQBoardSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim ActiveGrid As String

Private Sub cmbDAQBoard_Click()

    'When this board changes, all the info in the Channels
    'MSH flex-grid needs to change as well
    
    'Clear the MSH Channels flex-grid
    Me.gridDAQChannels.Clear
    Me.gridDAQChannels.ClearStructure
    
    'Reload with the currently selected Board Object passed from the
    'System Boards collection into the Load grid function
    LoadChannelsGrid SystemBoards(cmbDAQBoard.ItemData(cmbDAQBoard.ListIndex))

End Sub

Private Sub cmdAddBoard_Click()

    'Initialize the "Add new" check box on the Add Board form to 'checked'
    frmDAQ_Add.checkAddNew = Checked
    
    'Intialize the old board name = ""
    frmDAQ_Add.txtOldBoardName = ""

    'Open the Add/Edit Board Form
    Load frmDAQ_Add
    frmDAQ_Add.Show
        
End Sub

Private Sub cmdAddChan_Click()

    'Set the Add New check box to Checked
    frmDAQChannel_Add.checkAddNew.Value = Checked
    
    'Set the old chan name and old chan type text-boxes t0 ""
    frmDAQChannel_Add.txtOldChanName = ""
    frmDAQChannel_Add.txtOldChanType = ""
    
    'Set the Caption on the "Add/Edit" button on the DAQ Channel add form to "Change"
    frmDAQChannel_Add.cmdAddEditChan.Caption = "Add"
        
    'Load and Open the Add/Edit Channel form
    Load frmDAQChannel_Add
    frmDAQChannel_Add.Show

End Sub

Private Sub cmdApply_Click()

    Dim UserResp As Long

    'Open a confirmation message to the user
    UserResp = MsgBox("Applying these new DAQ Board settings could lead to fatal communication " & _
                      "problems between the Paleomag program and the AF, IRM, and ARM systems." & _
                      vbNewLine & vbNewLine & "Do you want to proceed with these changes?" & _
                      vbNewLine & vbNewLine & "This will not affect the .ini file", _
                      vbYesNo, _
                      "Caution!")
                      
    'If user clicks yes, then proceed
    If UserResp = vbYes Then
    
        'Overwrite the System Boards collection with the Local Boards collection
        Set SystemBoards = LocalBoards
        
        'Change status flag, Local and System Board collections are now identical
        modConfig.LocalAndSystemDifferent = False
        
    End If
            
End Sub



Private Sub cmdOK_Click()

    Dim UserResp As Long

    'Open a confirmation message to the user
    UserResp = MsgBox("Saving these new DAQ Board settings will alter the .ini file and could lead to " & _
                      "fatal communication " & _
                      "problems between the Paleomag program and the AF, IRM, and ARM systems." & _
                      vbNewLine & vbNewLine & "Do you want to proceed with these changes?", _
                      vbYesNo, _
                      "Caution!!!")
                      
    'If user clicks yes, then proceed
    If UserResp = vbYes Then
    
        'Overwrite the System Boards collection with the Local Boards collection
        Set SystemBoards = LocalBoards
        
        'Change status flag, Local and System Board collections are now identical
        modConfig.LocalAndSystemDifferent = False
        
        'Now Save the SystemBoards collection to the .ini file
        modConfig.Save_BoardsToINI
        
        'Now Reload the SystemBoards Collection from the .ini file
        
        'Turn on error handling
        On Error Resume Next
            
            modConfig.Get_BoardsFromIni
            
            'error check
            If Err.number <> 0 Then
            
                'Tell User a huge error happened
                MsgBox "Critical Error #" & Trim(Str(Err.number)) & _
                       " Occured During re-load of the Paleomag DAQ Board" & _
                       " Settings from the .ini file." & vbNewLine & vbNewLine & _
                       "The Paleomag code will now exit." & vbNewLine & _
                       "Please RESTORE The Pre-Change, Backed Up .INI file BEFORE re-starting the code!" & _
                       vbNewLine & vbNewLine & _
                       "System Error Source: " & Err.Source & vbNewLine & vbNewLine & _
                       "System Error Message:" & vbNewLine & _
                       Err.Description

                       
                'Kill the Code
                End
                
            End If
            
        'Turn off error handling
        On Error GoTo 0
        
        'Now Reload the System Channels from the .ini file
        '(This will re-load the channel descriptions and check
        ' for errors in the reload proces)
        
        'Turn on error handling
        On Error Resume Next
            
            modConfig.Get_ChannelsFromIni
            
            'error check
            If Err.number <> 0 Then
            
                'Tell User a huge error happened -
                'Conflict between old channel settings and new board settings
                MsgBox "Critical Error #" & Trim(Str(Err.number)) & _
                       " Occured During re-load of the Paleomag DAQ Channel" & _
                       " Assignments from the .ini file." & vbNewLine & _
                       "An essential DAQ Board was deleted from the DAQ Board settings; " & _
                       "DAQ Channel assignments point to a now missing DAQ Board." & vbNewLine & vbNewLine & _
                       "The Paleomag code will now exit." & vbNewLine & _
                       "Please RESTORE The Pre-Change, Backed Up .INI file BEFORE re-starting the code!" & _
                       vbNewLine & vbNewLine & _
                       "System Error Source: " & Err.Source & vbNewLine & vbNewLine & _
                       "System Error Message:" & vbNewLine & _
                       Err.Description
                       
                'Kill the Code
                End
                
            End If
            
        'Turn off error handling
        On Error GoTo 0
        
    End If
            
End Sub


Private Sub cmdDeleteChan_Click()

    Dim i As Long
    Dim N As Long
    Dim RowToDelete As Long
    Dim BoardName As String
    Dim ChanName As String
    Dim ChanType As String
    Dim doContinue As Boolean
    
    'This subroutine will only delete the channel from the grid-display
    'and will not affect the Channels Collections for the Board object
    'in the global System Boards Collection or the .INI file
    'To change those, the user would have to click the "Apply" or "OK" buttons

    With Me.gridSystemBoards
    
        'If user has selected the first row, then do nothing
        If .row = 0 Then
        
            'Tell user they are an idiot
            MsgBox "This row cannot be deleted. Please select another row.", , _
                   "Ooops!"
    
            Exit Sub
            
        End If
        
        'Make sure the user has actually selected a row
        If .row <> .RowSel Then
        
            'Tell user to actually select a @$%&!!! row
            MsgBox "No Channel selected for deletion. Please choose a Channel and then " & _
                   "re-click the ""Delete..."" button", , _
                   "Ooops!"
                   
            Exit Sub
            
        End If
        
        'Delete the selected Row
        RowToDelete = .row
        
        'Get the Board Name
        BoardName = Me.cmbDAQBoard.List(cmbDAQBoard.ListIndex)
        
        'Get the Channel Name
        .Col = 2
        ChanName = .text
        
        'Get the Channel Type
        .Col = 1
        ChanType = .text
        
        'Delete the channel from the Local Boards collection
        doContinue = DeleteChanFromColl(LocalBoards, _
                                        BoardName, _
                                        ChanType, _
                                        ChanName)
        
        If doContinue = False Then
        
            Exit Sub
            
        End If
        
        
        'Delete the Channel's row from the Channels flex grid display
        .RemoveItem RowToDelete
        
        'Set N = new number of rows
        N = .Rows
        
        'Now redo the numbering in the 1st column
        For i = 1 To N - 1
        
            .row = i
            .Col = 0
            .text = Trim(Str(i))
            .ColWidth = Me.TextWidth(.text)
            
        Next i
        
    End With
        
    'It's much easier to delete rows than to delete columns
    'Curse you Microsoft!

End Sub

Private Sub cmdEditBoard_Click()

    Dim TempName As String

    'Get the Board Name from the DAQ Board flex grid
    With Me.gridSystemBoards
    
        If .row = .RowSel And .Col = .ColSel Then
        
            'No row is selected in the DAQ Boards flex-grid, pop-up a quick message
            'to the user telling them to select a row
            MsgBox "No Board selected for editing. " & vbNewLine & _
                   "Please select a Board to edit, first!", , _
                   "Warning!"
                   
            'Leave this sub-routine
            Exit Sub
                   
        End If
        
        .row = .RowSel
        .Col = 0
        TempName = .text
        
    End With
    
    'Set the "Add New" checkbox to Unchecked
    frmDAQ_Add.checkAddNew = Unchecked
    
    'Set the old board name to the board name of the board that has been selected
    frmDAQ_Add.txtOldBoardName = TempName

    'Open the Add/Edit Board Form
    Load frmDAQ_Add
    frmDAQ_Add.Show
        
End Sub

Private Sub cmdEditChan_Click()

    Dim ChanName As String
    Dim ChanType As String
    Dim BoardName As String
    Dim TempChan As Channel
    
    'Get the channel object that the user has selected to edit
    With Me.gridDAQChannels
    
        'Check to see if there is a selected row
        If .row <> .RowSel Then
        
            'No row selected
            'Pop-up a message box telling the user to select a channel first
            MsgBox "No channel selected for editing.  Please select a channel in the " & _
                   "grid above and re-click the ""Edit..."" button", , _
                   "Ooops!"
        
        End If
        
        'Set the column get to the channel name:
        .Col = 2
        ChanName = .text
        
        'Set the column to get the channel type
        .Col = 3
        ChanType = .text
        
    End With
    
    'Get the board name from the DAQ Board combo-box
    BoardName = Me.cmbDAQBoard.List(Me.cmbDAQBoard.ListIndex)
    
    'Set fields in the Add/Edit Channel form
    frmDAQChannel_Add.txtOldChanName = ChanName
    frmDAQChannel_Add.txtOldChanType = ChanType
    frmDAQChannel_Add.txtBoardName = BoardName
    
    'Set the Add New check box to Unchecked
    frmDAQChannel_Add.checkAddNew.Value = Unchecked
    
    'Set the Caption on the "Add/Edit" button on the DAQ Channel add form to "Change"
    frmDAQChannel_Add.cmdAddEditChan.Caption = "Change"
        
    'Load frmDAQCHannel and show the form
    Load frmDAQChannel_Add
    frmDAQChannel_Add.Show
    
End Sub

Private Sub cmdDeleteBoard_Click()

    Dim ColToDelete As Long
    Dim i, j As Long
    Dim N, M As Long
    Dim doShift As Boolean
    Dim TempStr As String
    Dim doContinue As Boolean

    'This subroutine will only delete the board from the grid-display
    'and will not affect the global System Boards Collection or the .INI file
    'To change those, the user would have to click the "Apply" or "OK" buttons
    
    With Me.gridSystemBoards
    
        'Store the Col to delete
        ColToDelete = .Col
        
        'Check to see that the only board's column is selected
        If .ColSel <> ColToDelete Then
            
            'No column is selected.
            'Pop-up a "Please Select a column" message box
            MsgBox "No DAQ Board has been selected for deletion." & vbNewLine & _
                   "Please select a DAQ Board in the table above, first.", , _
                   "Ooops!"
        
            Exit Sub
                
        End If
            
        'If user has selected the first row, then do nothing
        If ColToDelete = 0 Then
        
            'Tell user they are an idiot
            MsgBox "This column cannot be deleted. Please select another column.", , _
                   "Ooops!"
    
            Exit Sub
            
        End If
            
        'Get the Board Name and delete it from the LocalBoards collection
        .row = 1
        
        'Check to see if the board fingered for deletion has Channels assigned
        'to it for Rock-mag communications and control
        doContinue = CheckBoardDependencies(LocalAssignedChannels, _
                                            LocalBoards(.text), _
                                            Nothing)
        
        If doContinue = False Then
        
            'User has selected not to proceed with the Board deletion
            Exit Sub
            
        End If
        
        'Resolve the changes to the Board / assigned function dependencies
        ResolveBoardDependencies LocalAssignedChannels, _
                                 LocalBoards(.text), _
                                 Nothing
        
        'Remove the selected board from the Local Boards Collection
        'Let's assume no errors will happen in which the board being deleted
        'can't be found in the LocalBoards collection
        'Save's on a heck of a lot of error checking code (yay!)
        LocalBoards.Remove .text
        
        'Change Status flag to show that the Local Boards collection is now different
        'from the System Boards collection
        modConfig.LocalAndSystemDifferent = True
        
        'Set N = number of rows
        N = .Rows
        
        'Set N = current number of cols
        N = .Cols
        
        'if .Cols = 2 then - only one board loaded to delete
        'Can do this by setting .Cols = 1
        If .Cols = 2 Then
        
            'Now change the columns to one to remove the deleted board from the display
            .Cols = 1
        
            'Exit the subroutine
            Exit Sub
            
        End If
        
        'Initialize doShift flag to false
        doShift = False
        
        'Have more than one board loaded
        'Loop through the data columns and if needed, shift columns
        'that are to the right of the deleted col to the left by one column
        For i = 1 To M - 1
        
            If i = ColToDelete Then
            
                doShift = True
                
            End If
            
            If doShift And i < M - 1 Then
            
                'Shift the entire contents of the col (all the rows worth)
                'one to the right into the current col
                
                'Iterate through each row
                For j = 1 To N - 1
                    
                    'Set the row
                    .row = j
                    
                    'Shift the contents of the columns
                    .Col = i + 1
                    TempStr = .text
                    .Col = i
                    .text = TempStr
                    
                    'Readjust the cell width of the current col
                    If .ColWidth < Me.TextWidth(.text) Then
                    
                        .ColWidth = Me.TextWidth(.text)
                        
                    End If
                    
                Next j
                
            End If
            
        Next i
            
        'Resize the number of colums less one
        .Cols = M - 1

        'Now renumber the column headers
        For i = 1 To M - 2
        
            .row = 0
            .text = "Board #" & Trim(Str(i))
        
        Next i
        
    End With

End Sub

Private Sub Form_Load()

    Dim i As Long
    Dim N As Long
    
    'if the local boards collection and the system boards collection
    'are not different from each other (i.e. by a board deletion, addition, or edit)
    'then set the Local Boards collection = System Boards collection
    ' and also set the Local Assigned channels = the System version of that collection
    If modConfig.LocalAndSystemDifferent = False Then
        
        Set LocalBoards = Nothing
        Set LocalBoards = SystemBoards
        
        Set LocalAssignedChannels = Nothing
        Set LocalAssignedChannels = SystemAssignedChannels
    
    End If
    
    'Check to see if Boards have been imported and if there are / were no INI boards
    If modConfig.NoINIBoards = True Or _
       modConfig.ImportBoardsDone = False _
    Then
        
        'Disable the "Apply", "OK", "Delete Board" buttons
        'and all the Channel controls
        'Set the "Add Board" button to default
        Me.cmdApply.Enabled = False
        Me.cmdDeleteBoard.Enabled = False
        Me.cmdOK.Enabled = False
        Me.cmdEditBoard.Enabled = False
        
        Me.cmdAddBoard.Enabled = True
        Me.cmdAddBoard.Default = True
        
        Me.frameChannels.Enabled = False
        Me.cmdAddChan.Enabled = False
        Me.cmdEditChan.Enabled = False
        Me.cmdClearChannels.Enabled = False
        Me.cmdDeleteChan.Enabled = False
        Me.cmbDAQBoard.Enabled = False
        Me.cmbChanType.Enabled = False
        Me.gridDAQChannels.Enabled = False
                
        'Exit the sub-routine -
        'Nothing more to do
        Exit Sub
                
    End If

    'There are boards to load
    On Error GoTo 0
    
        N = LocalBoards.Count
    
        If Err.number <> 0 Then
        
            'There aren't any boards loaded
            'in the System Boards global collection
            
            'Set no-INI boards flag = True
            modConfig.NoINIBoards = True
            modConfig.ImportBoardsDone = False
                    
            'Disable the "Apply", "OK", "Delete Board" buttons
            'Set the "Add Board" button to default
            Me.cmdApply.Enabled = False
            Me.cmdDeleteBoard.Enabled = False
            Me.cmdOK.Enabled = False
            Me.cmdAddBoard.Enabled = True
            Me.cmdAddBoard.Default = True
            Me.cmdEditBoard.Enabled = False
            Me.frameChannels.Enabled = False
            Me.cmdAddChan.Enabled = False
            Me.cmdEditChan.Enabled = False
            Me.cmdClearChannels.Enabled = False
            Me.cmdDeleteChan.Enabled = False
            Me.cmbDAQBoard.Enabled = False
            Me.cmbChanType.Enabled = False
            Me.gridDAQChannels.Enabled = False
            
            Exit Sub
            
        End If
    
    'Turn normal error flow back on
    On Error GoTo 0
    
    'Still aren't any boards loaded
    If N < 1 Then
    
        'Set no-INI boards flag = True
        modConfig.NoINIBoards = True
        modConfig.ImportBoardsDone = False
        
        'Disable the "Apply", "OK", "Delete Board" buttons
        'Set the "Add Board" button to default
        Me.cmdApply.Enabled = False
        Me.cmdDeleteBoard.Enabled = False
        Me.cmdOK.Enabled = False
        Me.cmdAddBoard.Enabled = True
        Me.cmdAddBoard.Default = True
        Me.cmdEditBoard.Enabled = False
        Me.frameChannels.Enabled = False
        Me.cmdAddChan.Enabled = False
        Me.cmdEditChan.Enabled = False
        Me.cmdClearChannels.Enabled = False
        Me.cmdDeleteChan.Enabled = False
        Me.cmbDAQBoard.Enabled = False
        Me.cmbChanType.Enabled = False
        Me.gridDAQChannels.Enabled = False
        
        'Exit the sub-routine
        Exit Sub
        
    End If
    
    'Otherwise, everything should be enabled
    'except the Channel add, delete & clear buttons
    Me.cmdEditBoard.Enabled = True
    Me.cmdApply.Enabled = True
    Me.cmdDeleteBoard.Enabled = True
    Me.cmdAddBoard.Enabled = True
    Me.cmdOK.Enabled = Ture
    Me.cmdCancel.Default = True
    Me.frameChannels.Enabled = True
    Me.gridDAQChannels.Enabled = True
    Me.cmdAddChan.Enabled = False
    Me.cmdEditChan.Enabled = False
    Me.cmdClearChannels.Enabled = False
    Me.cmdDeleteChan.Enabled = False
    Me.cmbDAQBoard.Enabled = True
    Me.cmbChanType.Enabled = True
    
    'Clear the two combo-boxes
    Me.cmbChanType.Clear
    Me.cmbDAQBoard.Clear
    
    'Load the values into the ChanType combo box
    Me.cmbChanType.AddItem "All", 0
    Me.cmbChanType.AddItem "Analog", 1
    Me.cmbChanType.AddItem "Digital", 2
    Me.cmbChanType.AddItem "Analog Input", 3
    Me.cmbChanType.AddItem "Analog Output", 4
    Me.cmbChanType.AddItem "Digital Input", 5
    Me.cmbChanType.AddItem "Digital Output", 6
    
    'Load the DAQ Boards into the MSH Board Flex-grid
    LoadBoardsGrid

End Sub
Public Sub LoadBoardsGrid()

    Dim N As Long
    Dim i As Long

    'Initialize the Column headers in the first fixed column
    'of the DAQ Boards MSH Flex-grid
    With Me.gridSystemBoards
    
        N = LocalBoards.Count
        
        'Enable the cmbDAQBoard combo box
        Me.cmbDAQBoard.Enabled = True
        
        'Set the rows and cols for the grid
        .Rows = 16
        .Cols = N + 1
        
        'Make sure a heavy rectangle
        'goes around the currently active cell
        .FocusRect = flexFocusHeavy
                    
        'Right Board Properties in the fixed col
        '(Col = 0) on the left-hand side of the Flex-grid
        .row = 1
        .Col = 0
        .text = "Board Name"
        
        .row = 2
        .text = "Board Device #"
        
        .row = 3
        .text = "Board Function"
        
        .row = 4
        .text = "Comm Protocol"
        
        .row = 5
        .text = "Analog In Channel Mode"
        
        .row = 6
        .text = "Max Analog In Rate (Hz)"
        
        .row = 7
        .text = "Max Analog Out Rate (Hz)"
        
        .row = 8
        .text = "Range Type (MCC only)"
        
        .row = 9
        .text = "Range Max (V)"
        
        .row = 10
        .text = "Range Min (V)"
        
        .row = 11
        .text = "Dig. IO Configured?"
        
        .row = 12
        .text = "Dig. Out Port Type?"
        
        .row = 13
        .text = "# of Analog In Channels"
        
        .row = 14
        .text = "# of Analog out Channels"
        
        .row = 15
        .text = "# of Digital In Channels"
        
        .row = 16
        .text = "# of Digital Out Channels"
    
        'Now, need to fill in values for each property
        'for each DAQ board currently loaded into the global
        'System Boards collection
    
        'N >=1, can do a for loop
        For i = 1 To N
        
            'Header in the fixed row (Row = 0)
            'This is so user can click this header
            'to select the entire column corresponding
            'to one board
            .row = 0
            .Col = i
            .text = "Board #" & Trim(Str(i))
    
            'Readjust the cell width of the current col
            If .ColWidth < Me.TextWidth(.text) Then
            
                .ColWidth = Me.TextWidth(.text)
                
            End If
    
    
            'Property values
            .row = 1
            .text = LocalBoards(i).BoardName
            
            'Readjust the cell width of the current col
            If .ColWidth < Me.TextWidth(.text) Then
            
                .ColWidth = Me.TextWidth(.text)
                
            End If
            
            
            .row = 2
            .text = Trim(Str(LocalBoards(i).BoardNum))
            
            'Readjust the cell width of the current col
            If .ColWidth < Me.TextWidth(.text) Then
            
                .ColWidth = Me.TextWidth(.text)
                
            End If
            
            
            .row = 3
            .text = LocalBoards(i).BoardFunction
            
            'Readjust the cell width of the current col
            If .ColWidth < Me.TextWidth(.text) Then
            
                .ColWidth = Me.TextWidth(.text)
                
            End If
            
            
            'Converting CommProtocol into a string
            'Instead of a mysterious number
            .row = 4
            Select Case LocalBoards(i).CommProtocol
            
                Case ADWIN_COM
            
                    .text = "ADWIN"
                    
                Case MCC_UL
                
                    .text = "MCC"
                    
                Else
                
                    .text = "Other?"
                    
            End Select
            
            'Readjust the cell width of the current col
            If .ColWidth < Me.TextWidth(.text) Then
            
                .ColWidth = Me.TextWidth(.text)
                
            End If
            
            
            'Ditto for the Analog Input Channel mode
            .row = 5
            If LocalBoards(i).BoardMode = DIFFERENTIALMODE Then
            
                .text = "Differential"
                
            Else
            
                .text = "Single"
                
            End If
                
             'Readjust the cell width of the current col
            If .ColWidth < Me.TextWidth(.text) Then
            
                .ColWidth = Me.TextWidth(.text)
                
            End If
               
               
            .row = 6
            .text = Trim(Str(LocalBoards(i).MaxAInRate))
            
            'Readjust the cell width of the current col
            If .ColWidth < Me.TextWidth(.text) Then
            
                .ColWidth = Me.TextWidth(.text)
                
            End If
            
            
            .row = 7
            .text = Trim(Str(LocalBoards(i).MaxAOutRate))
            
            'Readjust the cell width of the current col
            If .ColWidth < Me.TextWidth(.text) Then
            
                .ColWidth = Me.TextWidth(.text)
                
            End If
            
                        
            'If this is an MCC board, then it has a RangeType
            'value stored for it.  The RangeType is a long
            'corresponding to a global const declared in the MCC
            '.BAS module (modMCC).
            'The Range class module has in it a function that will
            'convert a RangeType long value to the corresponding
            'string descriptor = variable name used for the const
            'in modMCC
            .row = 8
            If LocalBoards(i).CommProtocol = MCC_UL Then
            
                .text = LocalBoards(i).Range.Get_RangeTypeStr()

            Else
            
                .text = "----"
                
            End If
            
            'Readjust the cell width of the current col
            If .ColWidth < Me.TextWidth(.text) Then
            
                .ColWidth = Me.TextWidth(.text)
                
            End If
            
            
            .row = 9
            .text = Trim(Str(LocalBoards(i).Range.MaxValue))
            
            'Readjust the cell width of the current col
            If .ColWidth < Me.TextWidth(.text) Then
            
                .ColWidth = Me.TextWidth(.text)
                
            End If
            
            
            .row = 10
            .text = Trim(Str(LocalBoards(i).Range.MinValue))
    
            'Readjust the cell width of the current col
            If .ColWidth < Me.TextWidth(.text) Then
            
                .ColWidth = Me.TextWidth(.text)
                
            End If
    
    
            .row = 11
            .text = Trim(Str(LocalBoards(i).DIOConfigured))
            
            'Readjust the cell width of the current col
            If .ColWidth < Me.TextWidth(.text) Then
            
                .ColWidth = Me.TextWidth(.text)
                
            End If
            
            
            'MCC boards support two different styles of
            'digital input and output port-specification
            'USB-HS1616-2 boards use the FIRSTPORTA, FIRSTPORTB, etc. system
            'PCI-DAS6030 boards use the AUXPORT convention
            'The Const values are named and stored as global constants
            'in the modMCC .BAS module
            .row = 12
            If LocalBoards(i).CommProtocol = MCC_UL Then
            
                'Handy-dandy function in the Board Class module
                'that converts the DOutPortType long value
                'into a string containing the const variable name
                'for the port type in the modMCC module library.
                .text = LocalBoards(i).Get_DOutPortTypeStr
            
            Else
            
                'ADWIN boards have a much different digital input / output
                'system - don't need the DOutPortType field
                .text = "----"
                
            End If
            
            'Readjust the cell width of the current col
            If .ColWidth < Me.TextWidth(.text) Then
            
                .ColWidth = Me.TextWidth(.text)
                
            End If
            
            
            .row = 13
            .text = Trim(Str(LocalBoards(i).AInChannels.Count))
            
            'Readjust the cell width of the current col
            If .ColWidth < Me.TextWidth(.text) Then
            
                .ColWidth = Me.TextWidth(.text)
                
            End If

            
            .row = 14
            .text = Trim(Str(LocalBoards(i).AOutChannels.Count))
            
            'Readjust the cell width of the current col
            If .ColWidth < Me.TextWidth(.text) Then
            
                .ColWidth = Me.TextWidth(.text)
                
            End If
            
            
            .row = 15
            .text = Trim(Str(LocalBoards(i).DInChannels.Count))
            
            'Readjust the cell width of the current col
            If .ColWidth < Me.TextWidth(.text) Then
            
                .ColWidth = Me.TextWidth(.text)
                
            End If
            
            
            .row = 16
            .text = Trim(Str(LocalBoards(i).DOutChannels.Count))
            
            'Readjust the cell width of the current col
            If .ColWidth < Me.TextWidth(.text) Then
            
                .ColWidth = Me.TextWidth(.text)
                
            End If
            
            
            'Load the board name of this board into the cmbDAQBoard combo box
            cmbDAQBoard.AddItem LocalBoards(i).BoardName, i - 1
            
        Next i
        
    End With
    
    'Select one of the Boards in the combo-box
    'then re-populate the channels grid based on that
    cmbDAQBoard.ListIndex = 0
    cmbDAQBoard_Click

End Sub
Public Sub LoadChannelsGrid(ByRef BoardObj As Board)

    Dim i As Long
    Dim N As Long
    
    With Me.gridDAQChannels
    
        'Fill in the Column Headers
        .row = 0
        .Col = 1
        .text = "Chan. Type"
        .ColWidth = Me.TextWidth(.text)
        
        .Col = 2
        .text = "Chan. Name"
        .ColWidth = Me.TextWidth(.text)
        
        .Col = 3
        .text = "Chan. #"
        .ColWidth = Me.TextWidth(.text)
        
        .Col = 4
        .text = "Channel Assignments"
        .ColWidth = Me.TextWidth(.text)
        
    End With
    
    'Determine how large to make the channels flex grid
    Select Case Me.cmbChanType.ItemData(Me.cmbChanType.ListIndex)
    
        Case "All"
        
            With BoardObj
        
                N = .AInChannels.Count + _
                    .AOutChannels.Count + _
                    .DInChannels.Count + _
                    .DOutChannels.Count
                
                'Set the size of the Channels Flex grid
                Me.gridDAQChannels.Cols = 5
                Me.gridDAQChannels.Rows = N + 1
            
                'Run through all the channels
                'Analog Input Channels
                LoadChannelCollection Me.gridDAQChannels, _
                                      .AInChannels, _
                                      1
                
                'Analog Output Channels
                LoadChannelCollection Me.gridDAQChannels, _
                                      .AOutChannels, _
                                      .AInChannels.Count
                    
                'Digital Input Channels
                LoadChannelCollection Me.gridDAQChannels, _
                                      .DInChannels, _
                                      .AInChannels.Count + _
                                      .AOutChannels.Count
                        
                'Digital Output Channels
                LoadChannelCollection Me.gridDAQChannels, _
                                      .DOutChannels, _
                                      .AInChannels.Count + _
                                      .AOutChannels.Count + _
                                      .DInChannels.Count
                                      
            End With
                        
        Case "Analog"
        
            With BoardObj
        
                N = .AInChannels.Count + _
                    .AOutChannels.Count
                
                'Set the size of the Channels Flex grid
                Me.gridDAQChannels.Cols = 5
                Me.gridDAQChannels.Rows = N + 1
            
                'Run through all the channels
                'Analog Input Channels
                LoadChannelCollection Me.gridDAQChannels, _
                                      .AInChannels, _
                                      1
                
                'Analog Output Channels
                LoadChannelCollection Me.gridDAQChannels, _
                                      .AOutChannels, _
                                      .AInChannels.Count
                                      
            End With
                
        Case "Digital"
        
            With BoardObj
        
                N = .DInChannels.Count + _
                    .DOutChannels.Count
                
                'Set the size of the Channels Flex grid
                Me.gridDAQChannels.Cols = 5
                Me.gridDAQChannels.Rows = N + 1
            
                'Run through all the channels
                'Digital Input Channels
                LoadChannelCollection Me.gridDAQChannels, _
                                      .DInChannels, _
                                      1
                        
                'Digital Output Channels
                LoadChannelCollection Me.gridDAQChannels, _
                                      .DOutChannels, _
                                      .DInChannels.Count
                                      
            End With
                
        Case "Analog Input"
        
            With BoardObj
            
                N = .AInChannels.Count
                
                'Set the size of the Channels Flex grid
                Me.gridDAQChannels.Cols = 5
                Me.gridDAQChannels.Rows = N + 1
            
                'Run through all the channels
                'Analog Input Channels
                LoadChannelCollection Me.gridDAQChannels, _
                                      .AInChannels, _
                                      1
                                 
            End With
            
        Case "Analog Output"
        
            With BoardObj
        
                N = .AOutChannels.Count
                
                'Set the size of the Channels Flex grid
                Me.gridDAQChannels.Cols = 5
                Me.gridDAQChannels.Rows = N + 1
            
                'Run through all the channels
                'Analog Output Channels
                LoadChannelCollection Me.gridDAQChannels, _
                                      .AOutChannels, _
                                      1
                                     
            End With
            
        Case "Digital Input"
        
            With BoardObj
        
                N = .DInChannels.Count
                
                'Set the size of the Channels Flex grid
                Me.gridDAQChannels.Cols = 5
                Me.gridDAQChannels.Rows = N + 1
            
                'Run through all the channels
                'Digital Input Channels
                LoadChannelCollection Me.gridDAQChannels, _
                                      .DInChannels, _
                                      1
                                                      
            End With
            
        Case "Digital Output"
        
            With BoardObj
        
                N = .DOutChannels.Count
                
                'Set the size of the Channels Flex grid
                Me.gridDAQChannels.Cols = 5
                Me.gridDAQChannels.Rows = N + 1
            
                'Run through all the channels
                'Digital Output Channels
                LoadChannelCollection Me.gridDAQChannels, _
                                      .DOutChannels, _
                                      1
                                      
            End With
            
    End Select
    
End Sub

Private Sub LoadChannelCollection(ByRef gridObj As MSHFlexGrid, _
                                  ByRef ChanCol As Channels, _
                                  Optional ByVal StrtRowIndex As Long = 1)

    Dim i As Long
    Dim N As Long

    'Get the number of items in the Channel Collection
    N = ChanCol.Count
    
    'Make sure that the flex grid has enough rows
    'to accomodate N new row entries
    If .Rows < StrtRowIndex + N Then
    
        .Rows = StrtRowIndex + N
        
    End If


    With gridObj
    
        'Run through all but the first row
        For i = StrtRowIndex To N - 1 + StrtRowIndex
    
            'Write in Channel number in 1st col
            .row = i
            .Col = 0
            .text = Trim(Str(i + 1))
            .ColWidth = Me.TextWidth(.text)
    
    
            'What type of channel are we writting in now?
            'Channel Type
            .Col = 1
            .text = ChanCol(i - StrtRowIndex).ChanType
            
            'Adjust column width if needed
            If .ColWidth < Me.TextWidth(.text) Then
            
                .ColWidth = Me.TextWidth(.text)
                
            End If
                
                
            'Channel Name
            .Col = 2
            .text = ChanCol(i - StrtRowIndex).ChanName
            
            'Adjust column width if needed
            If .ColWidth < Me.TextWidth(.text) Then
            
                .ColWidth = Me.TextWidth(.text)
                
            End If
            
            
            'Channel #
            .Col = 2
            .text = Trim(Str(ChanCol(i - StrtRowIndex).ChanNum))
            
            'Adjust column width if needed
            If .ColWidth < Me.TextWidth(.text) Then
            
                .ColWidth = Me.TextWidth(.text)
                
            End If
            
            
            'Channel Assignments
            .Col = 3
            .text = ChanCol(i - StrtRowIndex).ChanDescs.GetAll
            
            'Adjust column width if needed
            If .ColWidth < Me.TextWidth(.text) Then
            
                .ColWidth = Me.TextWidth(.text)
                
            End If
            
        Next i
        
    End With

End Sub
                                  

Private Sub gridDAQChannels_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    'This event handler only allows one row at a time to be selected in the flex grid

    'Is left mouse button held down?
    
    If Button = vbLeftButton Then
        
        'Move selected row
        gridDAQChannels.row = gridDAQChannels.MouseRow
    
    End If

End Sub

Private Sub gridSystemBoards_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    'This event handler only allows one row at a time to be selected in the flex grid

    'Is left mouse button held down?
    
    If Button = vbLeftButton Then
        
        'Move selected row
        gridSystemBoards.row = gridSystemBoards.MouseRow
    
    End If

End Sub

Private Sub gridDAQChannels_DblClick()

    Dim StartRow As Long
    Dim EndRow As Long
    
    'When user double clicks a cell in a row, highlight the whole row
    '(one channels-worth) of cells
    
    With Me.gridDAQChannels
    
        'Make sure only one row is selected
        .RowSel = .row
    
    End With
    
End Sub

Private Sub gridDAQChannels_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = vbRightButton Then
    
        'Highlight the whole row
        Me.gridDAQChannels_DblClick
        
        'Set Active Grid = Channel
        ActiveGrid = "Channel"
        
        PopupMenu Me.mnuSystemBoardsettings
        
    End If

End Sub


Private Sub mnuDAQAdd_Click(Index As Integer)

    'Determine which grid the popup menu came up for
    If ActiveGrid = "Board" Then
    
        Me.cmdAddBoard_Click
        
    Else
    
        Me.cmdAddChan_Click
        
    End If

End Sub

Private Sub mnuDAQDelete_Click(Index As Integer)

    'Determine which grid the popup menu came up for
    If ActiveGrid = "Board" Then
    
        Me.cmdDeleteBoard_Click
        
    Else
    
        Me.cmdDeleteChan_Click
        
    End If

End Sub

Private Sub mnuDAQEdit_Click(Index As Integer)

    'Determine which grid the popup menu came up for
    If ActiveGrid = "Board" Then
    
        Me.cmdEditBoard_Click
        
    Else
    
        Me.cmdEditChan_Click
        
    End If

End Sub
