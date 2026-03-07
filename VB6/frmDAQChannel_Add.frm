VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmDAQChannel_Add 
   Caption         =   "Add/Edit DAQ Channel"
   ClientHeight    =   5880
   ClientLeft      =   10950
   ClientTop       =   3270
   ClientWidth     =   5640
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5880
   ScaleWidth      =   5640
   Begin VB.PictureBox picAddFunction 
      BackColor       =   &H80000013&
      Height          =   2535
      Left            =   360
      ScaleHeight     =   2475
      ScaleWidth      =   4755
      TabIndex        =   18
      Top             =   360
      Visible         =   0   'False
      Width           =   4815
      Begin VB.CheckBox checkMatchingTypeOnly 
         BackColor       =   &H80000013&
         Caption         =   "Show Functions with matching Channel Type, Only"
         Height          =   375
         Left            =   240
         TabIndex        =   24
         Top             =   600
         Width           =   4095
      End
      Begin VB.CommandButton cmdCancelFunction 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   3000
         TabIndex        =   23
         Top             =   1920
         Width           =   1455
      End
      Begin VB.CommandButton cmdChooseFunction 
         Caption         =   "Choose this Function"
         Height          =   375
         Left            =   240
         TabIndex        =   22
         Top             =   1920
         Width           =   1935
      End
      Begin VB.ComboBox cmbFunctionViewer 
         Height          =   315
         Left            =   240
         TabIndex        =   21
         Top             =   1440
         Width           =   4215
      End
      Begin VB.CheckBox checkUnassignedOnly 
         BackColor       =   &H80000013&
         Caption         =   "Show Un-assigned Functions, Only"
         Height          =   375
         Left            =   240
         TabIndex        =   19
         Top             =   120
         Width           =   2895
      End
      Begin VB.Label Label6 
         BackColor       =   &H80000013&
         Caption         =   "AF / Rock Mag Functions:"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   1200
         Width           =   2655
      End
   End
   Begin VB.TextBox txtOldChanType 
      Height          =   285
      Left            =   2520
      TabIndex        =   16
      TabStop         =   0   'False
      Text            =   "Text1"
      Top             =   5520
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtOldChanName 
      Height          =   285
      Left            =   2520
      TabIndex        =   15
      TabStop         =   0   'False
      Text            =   "Text1"
      Top             =   5280
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Default         =   -1  'True
      Height          =   495
      Left            =   3240
      TabIndex        =   9
      Top             =   5280
      Width           =   2295
   End
   Begin VB.CommandButton cmdAddEditChan 
      Caption         =   "Add"
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   5280
      Width           =   2295
   End
   Begin VB.Frame Frame1 
      Caption         =   "DAQ Channel Settings"
      Height          =   5055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5415
      Begin VB.CheckBox checkAddNew 
         Caption         =   "Add New?"
         Height          =   615
         Left            =   4200
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   2160
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtBoardName 
         BackColor       =   &H80000016&
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   480
         Width           =   2295
      End
      Begin VB.CommandButton cmdAddFunc 
         BackColor       =   &H8000000A&
         Caption         =   "Add Function..."
         Height          =   435
         Left            =   2880
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   4440
         Width           =   1575
      End
      Begin VB.CommandButton cmdDeleteFunc 
         BackColor       =   &H8000000A&
         Caption         =   "Delete Function..."
         Height          =   435
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   4440
         Width           =   1575
      End
      Begin VB.TextBox txtChanNum 
         Height          =   315
         Left            =   1800
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   2280
         Width           =   975
      End
      Begin VB.TextBox txtChanName 
         Height          =   285
         Left            =   1800
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   1680
         Width           =   2295
      End
      Begin VB.ComboBox cmbChanType 
         Height          =   315
         Left            =   1800
         TabIndex        =   2
         Text            =   "Combo1"
         Top             =   1080
         Width           =   2295
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridChanAssignments 
         Height          =   1215
         Left            =   120
         TabIndex        =   5
         Top             =   3120
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   2143
         _Version        =   393216
         BackColorBkg    =   -2147483638
         SelectionMode   =   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label Label5 
         Caption         =   "Parent DAQ Board:"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "AF / Rockmag Functions Assigned to This Channel:"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   2880
         Width           =   4695
      End
      Begin VB.Label Label3 
         Caption         =   "Channel #:"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Channel Name:"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Channel Type:"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   1080
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmDAQChannel_Add"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim OldChan As Channel

Private Sub LoadFuncComboBox()

    Dim ChanType As String
    Dim UnassignedOnly As Boolean
    Dim FuncArray() As String
    Dim ChanTypeArray() As String
    Dim N As Long
    Dim i As Long
    Dim TempL As Long
    
    'Clear the Functions Combo box
    Me.cmbFunctionViewer.Clear
    
    'Get the Channel Type filter value
    If Me.checkMatchingTypeOnly.Value = Checked Then
    
        ChanType = GetFormChannelType
        
    Else
    
        ChanType = "-1"
        
    End If
    
    'Get the Unassigned only filter value
    UnassignedOnly = (Me.checkUnassignedOnly.Value = Checked)
       
    
    'Load the Function Array
    modAF_DAQ.GetFunctions LocalAssignedChannels, _
                           FuncArray, _
                           ChanTypeArray, _
                           ChanType, _
                           UnassignedOnly
                           
    'Get the size of FuncArray
    N = UBound(FuncArray)
    
    'Add results to combo-box
    If N <= 1 Then
    
        'No functions survived the filters,
        'Add "None" as the only element of the combo box
        Me.cmbFunctionViewer.AddItem "None", 0
        
    Else
    
        'Loop through the matches and add them one-by-one to the cmb-box
        'Translate the channel Types into integers:
        '(0 = "AI", 1 = "AO", 2 = "DI", 3 = "DO") to store in the
        'combo box ItemData fields
        For i = 1 To N
        
            Select Case ChanTypeArray(i)
            
                Case "AI"
                
                    TempL = 0
                    
                Case "AO"
                
                    TempL = 1
                    
                Case "DI"
                
                    TempL = 2
                    
                Case "DO"
                
                    TempL = 3
                    
            End Select
            
        Next i
        
        Me.cmbFunctionViewer.AddItem FuncArray(i), i - 1
        Me.cmbFunctionViewer.ItemData(i - 1) = TempL
        
    End If
    
End Sub


Private Sub checkMatchingTypeOnly_Click()

    'Filter has changed, need to reload the
    'function combo box

    'Run the LoadFuncComboBox function
    LoadFuncComboBox

End Sub

Private Sub checkUnassignedOnly_Click()

    'Filter has changed, need to reload the
    'function combo box

    'Run the LoadFuncComboBox function
    LoadFuncComboBox

End Sub

Private Sub cmdAddFunc_Click()

    'Load the Function Combo Box with function descriptions
    LoadFuncComboBox
    
    'Disable the main form controls
    Me.cmdAddEditChan.Enabled = False
    Me.cmdCancel.Enabled = False
    Me.cmdAddFunc.Enabled = False
    Me.cmdDeleteFunc.Enabled = False
    Me.gridChanAssignments.Enabled = False
    
    'Show the Add Function Picture box
    Me.picAddFunction.Visible = True
    
End Sub

Private Sub cmdCancel_Click()

    Unload Me
    Me.Hide

End Sub

Private Sub cmdCancelFunction_Click()

    'Hide the add function picture box and enabled all the main form buttons
    Me.picAddFunction.Visible = False
    Me.cmdAddEditChan.Enabled = True
    Me.cmdCancel.Enabled = True
    Me.cmdAddFunc.Enabled = True
    Me.cmdDeleteFunc.Enabled = True
    Me.gridChanAssignments.Enabled = True

End Sub

Private Sub cmdChooseFunction_Click()

    Dim TempL As Long
    Dim ChanType As String
    Dim ChanDesc As String

    'Turn on all the main form controls
    Me.cmdAddEditChan.Enabled = True
    Me.cmdCancel.Enabled = True
    Me.cmdAddFunc.Enabled = True
    Me.cmdDeleteFunc.Enabled = True
    Me.gridChanAssignments.Enabled = True
    
    'Get the channel type from the Function combo-box
    Select Case Me.cmbFunctionViewer.ItemData(Me.cmbFunctionViewer.ListIndex)
    
        Case 0
        
            ChanType = "AI"
            
        Case 1
        
            ChanType = "AO"
            
        Case 2
        
            ChanType = "DI"
            
        Case 3
        
            ChanType = "DO"
            
    End Select
    
    'Get the Channel Description as well
    ChanDesc = Me.cmbFunctionViewer.List(Me.cmbFunctionViewer.ListIndex)
       
    
    With Me.gridChanAssignments
    
        'Get the current # of rows in the grid
        TempL = .Rows
        
        'Add a new row
        .Rows = TempL + 1
        
        'Set the active row to the new row
        .row = .Rows - 1
        
        'Set the active col = 0
        'Write in the row #
        .Col = 0
        .text = Trim(Str(.row))
        .Width = Me.TextWidth(.text)
        
        'Set the active Col = 1
        .Col = 1
        
        'Write in the error status
        'Mis-matched channel type
        If ChanType <> GetFormChannelType Then
        
            .Col = 1
            .BackColor = QBColor(4)
            .ForeColor = QBColor(6)
            .text = "!"
            
            .Col = 3
            .BackColor = QBColor(12)
            
        End If
        
        'More than one function assignment (Gasp!)
        If .row > 2 Then
        
            .Col = 1
            .BackColor = QBColor(4)
            .ForeColor = QBColor(6)
            .text = "!"
            
            .Col = 2
            .BackColor = QBColor(12)
            
        End If
        
        'Write in the values now
        .Col = 2
        .text = ChanDesc
        .Col = 3
        .text = ChanType
        
    End With

End Sub

Private Sub cmdDeleteFunc_Click()

    'Just change the display grid, nothing else!
    'Check to see first that a row is actually selected
    If .row <> .RowSel Then
    
        'Pop-up a message box
        MsgBox "No Function selected to Delete.  Please selected a row and then re-click " & _
               "the button ""Delete Function""", , _
               "Ooops!"
               
        Exit Sub
        
    End If
    
    'If user has selected the header row (.Row = 0), then tell them that they're an idiot
    If .row = 0 Then
    
        MsgBox "This row cannot be deleted.  Please select another row!", , _
               "Ooops!"
    
        Exit Sub
        
    End If
        
    'Else, need to delete the selected row
    With Me.gridChanAssignments
    
        .RemoveItem .row
        
    End If
    
End Sub

Private Sub Form_Load()

    Dim ChanFunctions As String

    'Initialize OldChan with a Null value
    Set OldChan = Nothing

    'Set the Function Viewer box to not-visible
    Me.picAddFunction.Visible = False
    
    'Set the Function Viewer:
    'view only unassigned functions and
    'view only functions with matching channel type
    Me.checkUnassignedOnly.Value = Checked
    Me.checkMatchingTypeOnly.Value = Checked
        
    'Enable the main form buttons and flex-grid
    Me.cmdAddEditChan.Enabled = True
    Me.cmdCancel.Enabled = True
    Me.cmdAddFunc.Enabled = True
    Me.cmdDeleteFunc.Enabled = True
    Me.gridChanAssignments.Enabled = True
        
    'Clear the Channel Type Combo Box
    Me.cmbChanType.Clear
    
    'Reload the values for the Channel Type Combo Box
    Me.cmbChanType.AddItem "Analog Input", 0
    Me.cmbChanType.AddItem "Analog Output", 1
    Me.cmbChanType.AddItem "Digital Input", 2
    Me.cmbChanType.AddItem "Digital Output", 3
    
    'Clear the Channel function assignments grid
    Me.gridChanAssignments.Clear
    Me.gridChanAssignments.ClearStructure
    
    With Me.gridChanAssignments
        
        'Reload the headings for the grid
        .Rows = 2
        .Cols = 4
        .FixedCols = 2
                
        .row = 0
        .Col = 0
        .Width = Me.TextWidth("1")
        
        .row = 0
        .Col = 1
        .Width = Me.TextWidth("!!")
        
        .row = 0
        .Col = 2
        .text = "Function"
        .Width = Me.TextWidth(.text)
        
        .row = 0
        .Col = 3
        .text = "Req. Chan. Type"
        .Width = Me.TextWidth(.text)
    
    End With
    
    'Check to see if this is a Channel Add request, or a channel edit request
    If Me.checkAddNew.Value = Checked Then
    
        'We're adding a new channel
        
        'Blank the Channel Name and Channel Num
        Me.txtChanName = ""
        Me.txtChanNum = ""
        
    Else
    
        'We're editing an existing channel
        
        'Get the Channel object using the Local Collection Board Name, and ChanType, and Channel Name
        Set OldChan = modAF_DAQ.GetChanFromColl(LocalBoards, _
                                                Me.txtBoardName, _
                                                Me.txtOldChanName, _
                                                Me.txtOldChanType)
                                                
        'Set the display fields using the values in Old Chan
        '(leave the Board Name field alone)
        Me.txtChanName = OldChan.ChanName
        Me.txtChanNum = OldChan.ChanName
        
        SetChannelFunctions OldChan
        
    'Deallocate TempChan
    Set TempChan = Nothing
    
End Sub

Private Function GetFormChannelType() As String

    Dim TempStr As String

    'Get the channel type from the combo box
    TempStr = Me.cmbChanType.List(Me.cmbChanType.ListIndex)

    'Translate the channel type to the two-character string descriptor
    'used in the Channel object
    Select Case TempStr
    
        Case "Analog Input"
        
            GetFormChannelType = "AI"
    
        Case "Analog Output"
        
            GetFormChannelType = "AO"
            
        Case "Digital Input"
        
            GetFormChannelType = "DI"
            
        Case "Digital Output"
        
            GetFormChannelType = "DO"
            
        Else
        
            GetFormChannelType = "ERROR"
            
    End Select

End Function

Private Sub SetChannelFunctions(ByRef ChanObj As Channel)

    Dim FuncArray() As String
    Dim N As Long
    Dim M As Long
    Dim StartIndex As Long
    Dim i As Long
    Dim TempL As Long

    'Get the Dependencies string array for this channel
    modAF_DAQ.GetChannelDependencies LocalAssignedChannels, _
                                     ChanObj, _
                                     FuncArray()
                                       
    'Get the array size
    N = UBound(FuncArray, 1)
    
    'If N > 1, then we have acutal assigned channel functions
    If N <= 1 Then
    
        'There are no Channel function assignments
        Exit Sub
            
    Else
    
        'Loop through Channel function assignments / dependencies
        For i = 1 To N - 1
        
            'Now start editing the display grid
            With Me.gridChanAssignments
            
                'Add new row to the grid if needed
                If .Rows < i + 1 Then
                
                    TempL = .Rows
                    .Rows = TempL + 1
                    
                End If
                
                .row = i
                .Col = 0
                .text = Trim(Str(.row))
                .Width = Me.TextWidth(.text)
                
                'Check to see if the ChanObj channel type, and the Assignment channel type
                'Are the same.  If not, flag the channel assignment as bad
                If ChanObj.ChanType <> getformchantype Then
                
                    'Turn on Red backcoloring and Yellow text
                
                    .Col = 1
                    .BackColor = QBColor(4)
                    .ForeColor = QBColor(6)
                    .text = "!"
                    
                    .Col = 3
                    .BackColor = QBColor(12)
                                                                              
                Else
                
                    .Col = 3
                    .BackColor = &H80000005
                    
                End If
                
                If N >= 3 Then
                
                    'More than one function assigned to this channel
                    .Col = 1
                    .BackColor = QBColor(4)
                    .ForeColor = QBColor(6)
                    .text = "!"
                    
                    .Col = 2
                    .BackColor = QBColor(12)
                    
                Else
                
                    .Col = 2
                    .BackColor = &H80000005
                    
                End If
                
                If N <= 2 And ChanObj.ChanType = getformchantype Then
                
                    'No Error
                    'Turn off red back-coloring and Yellow text
                
                    .Col = 1
                    .BackColor = &H80000005
                    .ForeColor = &H80000008
                    .text = ""
                    
                End If
                
                'Now write in the values of the FuncArray into this row
                .Col = 2
                .text = FuncArray(i)
                                
                'Now check to see if there are any incompatible functions
                .Col = 3
                .text = ChanObj.ChanType
                
            End With
            
        Next i
        
    End If
    
End Sub

Private Sub gridChanAssignments_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    
    'Use the ToolTipText property of the Channel Function Assignments flex-grid
    'to pop-up a message when the user hovers the mouse over rows or cols with an error flagged
    '(i.e. rows or cols with red or light-red back-color)
    
    'Check to see if the row corresponding to the mouse row has it's 2nd or 3rd columns highlighted in light red
    With gridChanAssignments
    
        'Make sure only one row is selected at a time
        If .RowSel <> .row And Button = vbLeftButton And Shift = vbShiftMask Then
        
            .RowSel = .row
            
            'don't continue doing any thing else
            Exit Sub
            
        End If
    
        .row = .MouseRow
        
        If .MouseCol = 3 Then
        
            .Col = 3
            
            If .BackColor = QBColor(12) Then
            
                'This cell is highlighted with an error
                'Change the tool tips
                .ToolTipText = "DAQ Channel type is not compatible with this function assignment!"
                
            Else
            
                .ToolTipText = ""
                
            End If
            
        ElseIf .MouseCol = 2 Then
        
            .Col = 2
            
            If .BackColor = QBColor(12) Then
            
                'This cell is highlighted with an error
                'Change the tool tips
                .ToolTipText = "Multiple conflicting functions have been assigned to this DAQ Channel!"
                
            Else
            
                .ToolTipText = ""
                
        Else
        
            .Col = 1
            
            If .BackColor = QBColor(4) Then
            
                'There is an error flag highlighted for this row
                .ToolTipText = "One or more errors for this Function assignment. Move the mouse over " & _
                               "columns highlighted in light red for more information."
                               
            Else
            
                .ToolTipText = ""
                
            End If
            
        End If
        
    End With
        
End Sub
