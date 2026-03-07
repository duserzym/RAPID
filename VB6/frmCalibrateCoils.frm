VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmCalibrateCoils 
   Caption         =   "Dim LastTUnits As String"
   ClientHeight    =   7740
   ClientLeft      =   7380
   ClientTop       =   4395
   ClientWidth     =   8040
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7740
   ScaleWidth      =   8040
   Begin VB.TextBox txtMaxPauseBetweenRamps 
      Height          =   288
      Left            =   5040
      TabIndex        =   54
      Text            =   "30"
      Top             =   3240
      Width           =   855
   End
   Begin VB.TextBox txtMinPauseBetweenRamps 
      Height          =   288
      Left            =   3480
      TabIndex        =   53
      Text            =   "5"
      Top             =   3240
      Width           =   855
   End
   Begin VB.CommandButton cmdCalibrateTransverseProbeAngle 
      Caption         =   "Calibrate Transverse Probe Angle"
      Height          =   375
      Left            =   240
      TabIndex        =   49
      Top             =   2280
      Width           =   2535
   End
   Begin RichTextLib.RichTextBox rtfPrint 
      Height          =   8535
      Left            =   4200
      TabIndex        =   45
      Top             =   7680
      Visible         =   0   'False
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   15055
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmCalibrateCoils.frx":0000
   End
   Begin VB.CommandButton cmdSort 
      Caption         =   "Sort"
      Height          =   372
      Left            =   3360
      TabIndex        =   44
      Top             =   6600
      Width           =   972
   End
   Begin VB.TextBox txtCellEdit 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2040
      TabIndex        =   43
      Top             =   4320
      Width           =   855
   End
   Begin VB.PictureBox picGetCalPeak 
      BackColor       =   &H00FFFFC0&
      Height          =   1815
      Left            =   2760
      ScaleHeight     =   1755
      ScaleWidth      =   2955
      TabIndex        =   37
      Top             =   4320
      Width           =   3015
      Begin VB.CommandButton cmdManualRedo 
         Caption         =   "Re-do"
         Height          =   315
         Left            =   1560
         TabIndex        =   42
         Top             =   1320
         Width           =   1095
      End
      Begin VB.CommandButton cmdManualOK 
         Caption         =   "OK"
         Height          =   315
         Left            =   240
         TabIndex        =   41
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox txtPeakField 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   720
         TabIndex        =   39
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label lblFieldUnits 
         BackStyle       =   0  'Transparent
         Caption         =   "G"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2160
         TabIndex        =   40
         Top             =   740
         Width           =   495
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Enter the DC Peak Field Value currently on the Gaussmeter display:"
         Height          =   495
         Left            =   120
         TabIndex        =   38
         Top             =   120
         Width           =   2655
      End
   End
   Begin VB.CheckBox chkManualCalibration 
      Caption         =   "Manual Calibration?"
      Height          =   252
      Left            =   5400
      TabIndex        =   36
      Top             =   1440
      Width           =   1692
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Save to Settings (OK)"
      Height          =   372
      Left            =   1920
      TabIndex        =   35
      Top             =   7200
      Width           =   2295
   End
   Begin VB.CommandButton cmdPauseCalibration 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Pause Calibration"
      Height          =   372
      Left            =   6240
      MaskColor       =   &H8000000F&
      TabIndex        =   22
      Top             =   7200
      UseMaskColor    =   -1  'True
      Width           =   1575
   End
   Begin VB.CommandButton cmdStartCalibration 
      BackColor       =   &H0080FF80&
      Caption         =   "Start Calibration"
      Height          =   372
      Left            =   4680
      MaskColor       =   &H8000000F&
      TabIndex        =   21
      Top             =   7200
      UseMaskColor    =   -1  'True
      Width           =   1455
   End
   Begin VB.CommandButton cmdClearData 
      Caption         =   "Clear Data"
      Height          =   372
      Left            =   1200
      TabIndex        =   16
      Top             =   6600
      Width           =   972
   End
   Begin VB.CommandButton cmdLoadFromCSVFile 
      Caption         =   "Load from File"
      Height          =   372
      Left            =   6240
      TabIndex        =   19
      Top             =   6600
      Width           =   1332
   End
   Begin VB.CommandButton cmdSaveToCSVFile 
      Caption         =   "Save to file"
      Height          =   372
      Left            =   4680
      TabIndex        =   18
      Top             =   6600
      Width           =   1332
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   372
      Left            =   240
      TabIndex        =   15
      Top             =   6600
      Width           =   855
   End
   Begin VB.CheckBox chkVerbose 
      Caption         =   "Debug Mode?"
      Height          =   252
      Left            =   2880
      TabIndex        =   12
      Top             =   1440
      Width           =   1692
   End
   Begin VB.TextBox txtNumReplicateRamps 
      Height          =   288
      Left            =   3480
      TabIndex        =   11
      Top             =   2760
      Width           =   972
   End
   Begin VB.CommandButton cmdAddSteps 
      Caption         =   "Add"
      Height          =   372
      Left            =   7080
      TabIndex        =   10
      Top             =   1920
      Width           =   612
   End
   Begin VB.TextBox txtFromVolts 
      Height          =   288
      Left            =   4560
      TabIndex        =   8
      Top             =   1920
      Width           =   852
   End
   Begin VB.TextBox txtToVolts 
      Height          =   288
      Left            =   6000
      TabIndex        =   9
      Top             =   1920
      Width           =   852
   End
   Begin VB.Frame frameAxialMaxAndMin 
      Caption         =   "Axial Max / Min Voltages"
      Height          =   1215
      Left            =   2760
      TabIndex        =   27
      Top             =   120
      Width           =   2412
      Begin VB.TextBox txtAFAxialMinMonitorVoltage 
         Height          =   288
         Left            =   720
         TabIndex        =   3
         Top             =   720
         Width           =   1092
      End
      Begin VB.TextBox txtAFAxialMaxMonitorVoltage 
         Height          =   288
         Left            =   720
         TabIndex        =   2
         Top             =   360
         Width           =   1092
      End
      Begin VB.Label Label4 
         Caption         =   "Max:"
         Height          =   255
         Left            =   240
         TabIndex        =   29
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Min:"
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Top             =   720
         Width           =   375
      End
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear All"
      Height          =   372
      Left            =   2280
      TabIndex        =   17
      Top             =   6600
      Width           =   972
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   372
      Left            =   120
      TabIndex        =   20
      Top             =   7200
      Width           =   1215
   End
   Begin VB.Frame frameTransMaxAndMin 
      Caption         =   "Trans. Max / Min Voltages"
      Height          =   1215
      Left            =   5280
      TabIndex        =   24
      Top             =   120
      Width           =   2532
      Begin VB.TextBox txtAFTransMinMonitorVoltage 
         Height          =   288
         Left            =   840
         TabIndex        =   5
         Top             =   720
         Width           =   1092
      End
      Begin VB.TextBox txtAFTransMaxMonitorVoltage 
         Height          =   288
         Left            =   840
         TabIndex        =   4
         Top             =   360
         Width           =   1092
      End
      Begin VB.Label Label2 
         Caption         =   "Min:"
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   720
         Width           =   375
      End
      Begin VB.Label lblMaxTransVoltage 
         Caption         =   "Max:"
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.CheckBox chkLogScale 
      Caption         =   "Log Scale"
      Height          =   252
      Left            =   2760
      TabIndex        =   7
      Top             =   1920
      Width           =   1092
   End
   Begin VB.TextBox txtStepSize 
      Height          =   288
      Left            =   1680
      TabIndex        =   6
      Top             =   1920
      Width           =   852
   End
   Begin VB.Frame frameCoilSelection 
      Caption         =   "AF Coil Selection"
      Height          =   1575
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   2532
      Begin MSComDlg.CommonDialog dlgPrint 
         Left            =   1680
         Top             =   360
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CheckBox chkLockCoils 
         Caption         =   "Lock coil selection"
         Height          =   255
         Left            =   240
         TabIndex        =   34
         Top             =   1200
         Width           =   2175
      End
      Begin VB.OptionButton optCoil 
         Caption         =   "Transverse"
         Height          =   252
         Index           =   1
         Left            =   240
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   720
         Width           =   1212
      End
      Begin VB.OptionButton optCoil 
         Caption         =   "Axial"
         Height          =   252
         Index           =   0
         Left            =   240
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   360
         Width           =   852
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridCalibration 
      Height          =   2775
      Left            =   120
      TabIndex        =   13
      Top             =   3720
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   4895
      _Version        =   393216
      Rows            =   10
      FixedCols       =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000A&
      DragMode        =   1  'Automatic
      Height          =   3492
      Left            =   0
      TabIndex        =   33
      Top             =   3600
      Width           =   7935
   End
   Begin VB.TextBox txtSingleStep 
      Height          =   288
      Left            =   5280
      TabIndex        =   46
      Top             =   2400
      Width           =   852
   End
   Begin VB.CommandButton cmdAddSingle 
      Caption         =   "Add Single Step"
      Height          =   372
      Left            =   6360
      TabIndex        =   48
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Label Label9 
      Caption         =   "Max:"
      Height          =   255
      Left            =   4560
      TabIndex        =   52
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label Label8 
      Caption         =   "Min:"
      Height          =   255
      Left            =   3120
      TabIndex        =   51
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label Label7 
      Caption         =   "Pause Between Ramps (secs) "
      Height          =   255
      Left            =   600
      TabIndex        =   50
      Top             =   3240
      Width           =   2295
   End
   Begin VB.Label lblAFVoltStep 
      Alignment       =   1  'Right Justify
      Caption         =   "AF Volt Step:"
      Height          =   255
      Index           =   1
      Left            =   3720
      TabIndex        =   47
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label lblNumReplicates 
      Caption         =   "# of Replicate AF Ramps per Voltage Step:"
      Height          =   255
      Left            =   120
      TabIndex        =   32
      Top             =   2760
      Width           =   3255
   End
   Begin VB.Label Label6 
      Caption         =   "To:"
      Height          =   255
      Left            =   5640
      TabIndex        =   31
      Top             =   1920
      Width           =   495
   End
   Begin VB.Label Label5 
      Caption         =   "From:"
      Height          =   255
      Left            =   4080
      TabIndex        =   30
      Top             =   1920
      Width           =   495
   End
   Begin VB.Label lblAFVoltStep 
      Alignment       =   1  'Right Justify
      Caption         =   "AF Volt Step:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   23
      Top             =   1920
      Width           =   1455
   End
End
Attribute VB_Name = "frmCalibrateCoils"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------------------------------------------'
'--------------------------------------------------------------------------------------------------------------------'
'
'   July, 2010
'
'   Isaac Hilburn
'
'   Form for running the IRM & AF calibration either automatically through USB or Comm port connection to the
'   Hirst Instruments 908A gaussmeter, or semi-automatically through the user manually entering the peak DC-field
'   measurements from the gaussmeter display.
'
'--------------------------------------------------------------------------------------------------------------------'
'
'   IMPORTANT NOTE!
'   August 5, 2010
'
'   Isaac Hilburn
'
'   Do NOT(!!) set a default button for frmCalibrateCoils.  This will break the flex-grid cell editing routine that allows
'   the user to exit from editing a cell by clicking the 'Enter' key.  This routine uses the KeyDown event
'   for the wandering text-box control (txtCellEdit) that serves as a stealthy flex-grid cell editor.  KeyDown cannot
'   capture an 'Enter' key-down event if a default button is set for the object's parent form.
'
'--------------------------------------------------------------------------------------------------------------------'
'--------------------------------------------------------------------------------------------------------------------'

Public InAFMode As Boolean
Dim CoilString As String
Dim CurrentRow As Long
Dim LastRange As Long
Dim CurrentRange As Long
Dim isUserChange As Boolean
Dim AxialCalData() As String
Dim transCalData() As String
Dim UnsavedChanges As Boolean
Dim ManualMode As Boolean
Dim UserResponse As Long
Dim CurrentCellPos(2) As Single
Dim CurrentCell(2) As Long

Dim gaussmeter_log_file_name As String

Dim CellEditHasFocus As Boolean

Dim CalStatus As String

Private Type ControlPositionType
    Left As Single
    Top As Single
    Width As Single
    Height As Single
    FontSize As Single
End Type
Private m_ControlPositions() As ControlPositionType
Private m_FormWid As Single
Private m_FormHgt As Single

Private Function CheckRange(ByVal MagField As Double, _
                            Optional ByVal Units As String = vbNullString) As Boolean

    Dim Range3_Max As Double
    Dim Range2_Max As Double
    Dim Range1_Max As Double
    
    Dim RangeString As String
    
    If Units = vbNullString Then Units = modConfig.AFUnits
    
    If Right(Units, 1) = "G" Then
    
        Range3_Max = 29.9
        Range2_Max = 299.9
        
        If InAFMode = True Then
            
            Range1_Max = 10000
            
        Else
        
            Range1_Max = 2999
        
        End If
        
    ElseIf Right(Units, 2) = "Oe" Then
    
        Range3_Max = 29.99
        Range2_Max = 299.9
        Range1_Max = 2999
        
        If InAFMode = True Then
            
            Range1_Max = 10000
            
        Else
        
            Range1_Max = 3000
        
        End If
        
    ElseIf Right(Units, 1) = "T" Then
    
        Range3_Max = 0.002999
        Range2_Max = 0.02999
        
        If InAFMode = True Then
            
            Range1_Max = 1
            
        Else
        
            Range1_Max = 0.2999
        
        End If
        
        
        
    ElseIf Right(Units, 3) = "A/m" Then
    
        Range3_Max = 0.002999
        Range2_Max = 0.02999
        
        If InAFMode = True Then
            
            Range1_Max = 1
            
        Else
        
            Range1_Max = 0.2999
        
        End If
        
    End If

    'If in automatic mode, read values from the gaussmeter to determine if we're in the right range
    'Check to see if we're using the right range
    If ManualMode = False Then
        
        If val(MagField) > Range3_Max And _
           val(MagField) < Range2_Max And _
           frm908AGaussmeter.optRange(2).value <> True And _
           LastRange <> 2 _
        Then
    
            'Store the value of the last Gaussmeter range
            LastRange = frm908AGaussmeter.CurrentRange
            
            'Change the range
            frm908AGaussmeter.optRange(2).value = True
            
            'Wait 1 second ( 1000 ms)
            PauseTill timeGetTime() + 1000
            
            'Reset the DC Peak field value
            frm908AGaussmeter.cmdResetPeak_Click
            
            'Wait 1 second ( 1000 ms)
            PauseTill timeGetTime() + 1000
            
            'NULL the Gaussmeter
            frm908AGaussmeter.DoSilentNull
                        
            'Wait 3 seconds ( 3000 ms)
            PauseTill timeGetTime() + 3000
            
            'Set return value to indicate that the range was changed
            CheckRange = True
    
        ElseIf val(MagField) > Range2_Max And _
               val(MagField) < Range1_Max And _
               frm908AGaussmeter.optRange(1).value <> True And _
               LastRange <> 1 _
        Then
    
            'Store the value of the last Gaussmeter range
            LastRange = frm908AGaussmeter.CurrentRange
            
            'Change the range
            frm908AGaussmeter.optRange(1).value = True
            
            'Wait 1 second ( 1000 ms)
            PauseTill timeGetTime() + 1000
            
            'Reset the DC Peak field value
            frm908AGaussmeter.cmdResetPeak_Click
            
            'Wait 1 second ( 1000 ms)
            PauseTill timeGetTime() + 1000
            
            'NULL the Gaussmeter
            frm908AGaussmeter.DoSilentNull
                        
             'Wait 3 seconds ( 3000 ms)
            PauseTill timeGetTime() + 3000
            
            'Set return value to indicate that the range was changed
            CheckRange = True
            
        ElseIf val(MagField) > Range1_Max And _
               frm908AGaussmeter.optRange(0).value <> True _
        Then
    
            'Store the value of the last Gaussmeter range
            LastRange = frm908AGaussmeter.CurrentRange
            
            'Change the range
            frm908AGaussmeter.optRange(0).value = True
            
            'Wait 1 second ( 1000 ms)
            PauseTill timeGetTime() + 1000
            
            'Reset the DC Peak field value
            frm908AGaussmeter.cmdResetPeak_Click
            
            'Wait 1 second ( 1000 ms)
            PauseTill timeGetTime() + 1000
            
            'NULL the Gaussmeter
            frm908AGaussmeter.DoSilentNull
                        
             'Wait 3 seconds ( 3000 ms)
            PauseTill timeGetTime() + 3000
            
            'Set return value to indicate that the range was changed
            CheckRange = True
    
        Else
        
            CheckRange = False
            
        End If
        
    Else
    
        'We're in manual mode
        'If PeakField > Range3_max, need to have user
        'change range to Range_2
        
        If Abs(MaxField) > Range3_Max And _
           CurrentRange > 2 _
        Then
            
            RangeString = GetRangeString(modConfig.AFUnits, 2)
            LastRange = 3
            CurrentRange = 2
            
        ElseIf Abs(MaxField) > Range2_Max And _
               CurrentRange > 1 _
        Then
        
            RangeString = GetRangeString(modConfig.AFUnits, 1)
            LastRange = 2
            CurrentRange = 1
            
        ElseIf Abs(MaxField) > Range1_Max And _
               CurrentRange > 0 _
        Then
        
            RangeString = GetRangeString(modConfig.AFUnits, 0)
            LastRange = 1
            CurrentRange = 0
            
        Else
        
            'No range change needed,
            'return false and exit the function
            CheckRange = False
            Exit Function
            
        End If
            
        MsgBox "1) Please click the Gaussmeter 'Range' button until the display " & _
               "looks like: """ & RangeString & """" & vbNewLine & _
               "2) Then click the ""Reset"" button" & vbNewLine & _
               "3) Wait a few seconds for the new DC Peak Value to stabilize" & vbNewLine & _
               "4) Then Click the NULL button and wait a few seconds after the NULL process has finished." & _
               vbNewLine & vbNewLine & _
               "Click the ""OK"" button below when you have finished the steps above.", _
               vbOKOnly, _
               "Change Gaussmeter Range"
            
        CheckRange = True
        
    End If
        
End Function

Private Sub chkLockCoils_Click()

    If Me.chkLockCoils.value = Checked Then
    
        CoilsLocked = True
        optCoil(0).Enabled = False
        optCoil(1).Enabled = False
        
    ElseIf Me.chkLockCoils.value = Unchecked Then
    
        CoilsLocked = False
        optCoil(0).Enabled = True
        optCoil(1).Enabled = True
        
    End If

End Sub

Private Sub chkManualCalibration_Click()

    If Me.chkManualCalibration.value = Checked Then
    
        ManualMode = True
        
    ElseIf Me.chkManualCalibration.value = Unchecked Then
    
        ManualMode = False
        
    End If

End Sub

Private Sub ClearAndLoadHeaders()
                                
    With Me.gridCalibration
    
        'Clear the grid
        .Clear
        .ClearStructure
    
        .Rows = 2
        .Cols = 4
        .FixedRows = 1
        .FixedCols = 1
        
    End With
                                
    LoadHeaders
                                
End Sub

Public Function ClipHangingDecimal(ByVal NumbStr As String) As String

    Dim TempString As String
    
    If Right(NumbStr, 1) = "." Then
    
        TempString = Mid(NumbStr, 1, Len(NumbStr) - 1)
        ClipHangingDecimal = TempString
        
    Else
    
        ClipHangingDecimal = NumbStr
        
    End If

End Function

Private Sub cmdAddSingle_Click()

    Dim N As Long
    Dim TempStr As String
    
    With Me.gridCalibration
    
        'Store the single step text-box string to a local
        TempStr = Trim(Me.txtSingleStep.text)
    
        'Validate the text in the single step text-box
        If ValidateTargetText(TempStr) = False Then
        
            'Text is bad, need to tell user and blank the txtSingleStep textbox
            MsgBox "The value that you attempted to add to the table is invalid." & _
                   vbNewLine & vbNewLine & _
                   "Value = " & TempStr & vbNewLine & _
                   "Value must be a number, must be > 0, and must be <= to" & _
                   "the maximum allowed value (3999 for the 2G system, and 10 for the " & _
                   "ADWIN system)", , _
                   "Whoops!"
                   
            'Blank the contents of single step textbox
            Me.txtSingleStep.text = vbNullString
            
            Exit Sub
            
        End If
        
        'Otherwise, text is valid and can be added to the calibration grid
        N = .Rows
        
        If .TextMatrix(N - 1, 1) = "" Or _
           .TextMatrix(N - 1, 1) = vbNullString _
        Then
        
            'Row is empty, can add the new value to this last row
            .TextMatrix(N - 1, 1) = TempStr
            
        Else
        
            'Last row is not empty, need to add a new row
            .Rows = N + 1
            
            'Add the new value to this new row
            .TextMatrix(N, 1) = TempStr
            
        End If
        
        'Re-sort the calibration grid
        SortGrid Me.gridCalibration, _
                 Me, _
                 1, _
                 .Rows - 1, _
                 1, .Cols - 1, _
                 True
                 
        'Store the text in the first cell to the cell-edit textbox
        CurrentCell(0) = 1
        CurrentCell(1) = 1
        .row = 1
        .Col = 1
        Me.txtCellEdit.text = .TextMatrix(1, 1)
        
        'Set the top-most row to the last row
        .TopRow = .Rows - 1
        
    End With
        
End Sub

Private Sub cmdAddSteps_Click()

    Dim i As Long
    Dim Volts As Double
    Dim StepSize As Double
    Dim isLog As Boolean
    Dim FromVolts As Double
    Dim ToVolts As Double
    Dim MaxCoilVolts As Double
    Dim MinCoilVolts As Double
    Dim NumReplicates As Long
    Dim StartRow As Long
    Dim ForStep As Long
    
    'Determine step size based upon what mode this form has been called in
    'and the AF system being used
    If AFSystem = "ADWIN" Or _
       InAFMode = False _
    Then
    
        'IRM or ADWIN AF
        ForStep = 2
        
    Else
    
        '2G AF
        ForStep = 1
    
    End If
    
    'Hide the cell edit text-box
    Me.txtCellEdit.Visible = False
    
    'Read out values into local variables for from and to volts
    FromVolts = val(Me.txtFromVolts)
    ToVolts = val(Me.txtToVolts)
    
    If AFSystem = "ADWIN" Or _
       InAFMode = False _
    Then
    
        'Depending on which coil is selected
        'by the user, load that coil's max and min voltages into
        'the two local variables
        If ActiveCoilSystem = AxialCoilSystem Then
        
            If InAFMode = True Then
                    
                'Adding rows for an AF Axial calibration
                MaxCoilVolts = modConfig.AfAxialMonMax
                MinCoilVolts = 0
                CoilString = "AF Axial"
            
            Else
            
                'Adding rows for an IRM Axial calibration
                MaxCoilVolts = modConfig.IRMAxialVoltMax
                MinCoilVolts = 0
                CoilString = "IRM Axial"
            
            End If
            
        ElseIf ActiveCoilSystem = TransverseCoilSystem Then
            
            If InAFMode = True Then
            
                'Adding rows for an AF Transverse calibration
                MaxCoilVolts = modConfig.AfTransMonMax
                MinCoilVolts = 0
                CoilString = "AF Transverse"
                
            Else
            
                'Adding rows for an IRM transverse calibration
                MaxCoilVolts = modConfig.IRMTransVoltMax
                MinCoilVolts = 0
                CoilString = "IRM Transverse"
                
            End If
                
        End If
        
        'Now validate Max and Min coil voltages
        If MaxCoilVolts <= MinCoilVolts Then
        
            'Quick Message Box to user
            MsgBox "Max " & CoilString & " coil voltage must be larger than the Min voltage." & _
                    vbNewLine & vbNewLine & "Max Voltage = " & Trim(str(MaxCoilVolts)) & _
                    " Volts" & vbNewLine & "Min Voltage = " & Trim(str(MinCoilVolts)) & _
                    " Volts", , _
                    "Warning!"
                    
            Exit Sub
            
        End If
        
        'Make sure both max and min coil voltages are greater than zero
        'Note: We also never want the Max coil voltage to equal zero.
        If MaxCoilVolts <= 0 Or MinCoilVolts < 0 Then
        
            MsgBox "Max and/or Min " & CoilString & " coil voltages are less than zero." & _
                    vbNewLine & vbNewLine & "Max Voltage = " & Trim(str(MaxCoilVolts)) & _
                    " Volts" & vbNewLine & "Min Voltage = " & Trim(str(MinCoilVolts)) & _
                    " Volts", , _
                    "Warning!"
                    
            Exit Sub
            
        End If
        
        isLog = False
        
        If chkLogScale.value = Checked Then
        
            If val(txtStepSize) <= 0 Then
            
                MsgBox "Can't use log scale with a negative or zero voltage step size!" & _
                        vbNewLine & "Voltage step size = " & Me.txtStepSize.text & " Volts"
            
                Exit Sub
                
            End If
            
            isLog = True
            
        End If
        
        'Now coerce from and to voltages if they are wrong
        'if from voltage < MinCoilVolts, set it equal to the MinCoilVolts
        If FromVolts < MinCoilVolts Then
            
            FromVolts = MinCoilVolts
            txtFromVolts.text = Trim(str(FromVolts))
            
        End If
        
        If ToVolts > MaxCoilVolts Then
        
            ToVolts = MaxCoilVolts
            txtToVolts.text = Trim(str(ToVolts))
            
        End If
        
        'Set the StepSize local variable
        StepSize = val(Me.txtStepSize)
        
        If FromVolts > ToVolts And StepSize > 0 Then
        
            MsgBox "From Voltage must be smaller than To voltage with a positive volt step size." & _
                    vbNewLine & "From = " & Trim(str(FromVolts)) & " Volts" & vbNewLine & _
                    "To = " & Trim(str(ToVolts)) & " Volts" & vbNewLine & "Step Size = " & _
                    Trim(str(StepSize))
                    
            Exit Sub
    
        End If
        
        If FromVolts < ToVolts And StepSize < 0 Then
        
            MsgBox "From Voltage must be larger than To voltage with a negative volt step size." & _
                    vbNewLine & "From = " & Trim(str(FromVolts)) & " Volts" & vbNewLine & _
                    "To = " & Trim(str(ToVolts)) & " Volts" & vbNewLine & "Step Size = " & _
                    Trim(str(StepSize))
    
            Exit Sub
    
        End If
                
        'Check the number of replicates to run
        If val(Me.txtNumReplicateRamps) < 3 Then
        
            Me.txtNumReplicateRamps = "3"
            
        End If
        
        'Store to local variable the number of replicate times
        'to ramp up and measure the field
        NumReplicates = CLng(val(Me.txtNumReplicateRamps))
        
        'Set number of columns
        Me.gridCalibration.Cols = GetColCount(NumReplicates)
                
        For i = 4 To 4 + (NumReplicates * ForStep) - 1 Step ForStep
        
            With Me.gridCalibration
                
                .row = 0
                .Col = i
                .text = "Field #" & Trim(str((i - 3) \ ForStep + 1)) & " (" & modConfig.AFUnits & ")"
                                
                .row = 0
                .Col = i + 1
                .text = "Max Volts #" & Trim(str((i - 3) \ ForStep + 1)) & ""
                
            End With
            
        Next i
                    
    Else
    
        'This is a 2G calibration run
    
        'Recast from and to and step size values into the nearest integer values
        FromVolts = CInt(FromVolts)
        ToVolts = CInt(ToVolts)
        StepSize = CInt(StepSize)
            
        'Depending on which coil is selected
        'by the user, set the correct coil-string
        If ActiveCoilSystem = AxialCoilSystem Then
        
            CoilString = "AF Axial"
            
        ElseIf ActiveCoilSystem = TransverseCoilSystem Then
        
            CoilString = "AF Transverse"
            
        Else
        
            CoilString = ""
            
        End If
                
        isLog = False
        
        If chkLogScale.value = Checked Then
        
            If val(txtStepSize) <= 0 Then
            
                MsgBox "Can't use log scale with a negative or zero voltage step size!" & _
                        vbNewLine & "2G Counts step size = " & Me.txtStepSize.text
            
                Exit Sub
                
            End If
            
            isLog = True
            
        End If
        
        'Now coerce from and to voltages if they are wrong
        'if from voltage < 0 2G counts, set it equal to zero
        If FromVolts < 0 Then
            
            FromVolts = 0
            txtFromVolts.text = Trim(str(FromVolts))
            
        End If
        
        'If to voltage > 3999 2G counts, set it equal to 3999
        If ToVolts > 3999 Then
        
            ToVolts = 3999
            txtToVolts.text = Trim(str(ToVolts))
            
        End If
        
        'Set the StepSize local variable
        StepSize = val(Me.txtStepSize)
                
        If FromVolts > ToVolts And StepSize > 0 Then
        
            MsgBox "From 2G counts must be smaller than To 2G counts with a positive step size." & _
                    vbNewLine & "From = " & Trim(str(FromVolts)) & vbNewLine & _
                    "To = " & Trim(str(ToVolts)) & vbNewLine & "Step Size = " & _
                    Trim(str(StepSize))
                    
            Exit Sub
    
        End If
        
        If FromVolts < ToVolts And StepSize < 0 Then
        
            MsgBox "From 2G counts must be larger than To 2G counts with a negative step size." & _
                    vbNewLine & "From = " & Trim(str(FromVolts)) & vbNewLine & _
                    "To = " & Trim(str(ToVolts)) & "Step Size = " & _
                    Trim(str(StepSize))
    
            Exit Sub
    
        End If
                
        'Check the number of replicates to run
        If val(Me.txtNumReplicateRamps) < 3 Then
        
            Me.txtNumReplicateRamps = "3"
            
        End If
        
        'Store to local variable the number of replicate times
        'to ramp up and measure the field
        NumReplicates = CLng(val(Me.txtNumReplicateRamps))
        
        'Set number of columns
        Me.gridCalibration.Cols = GetColCount(NumReplicates)
                    
        For i = 4 To 4 + NumReplicates - 1
        
            With Me.gridCalibration
            
                .row = 0
                .Col = i
                .text = "Field #" & Trim(str((i - 3) \ 2 + 1)) & " (" & modConfig.AFUnits & ")"
                
            End With
                
        Next i
    
    End If
    
    'Number of rows to add
    If StepSize = 0 Then Exit Sub
    
    N = Round((ToVolts * StepSize / Abs(StepSize) - FromVolts * StepSize / Abs(StepSize)) _
                        / StepSize * StepSize / Abs(StepSize), _
                    0)
    
    If N <= 0 Then Exit Sub
    
    StartRow = CurrentRow
                
    If StartRow > 1 Then
    
        With Me.gridCalibration
        
            .row = StartRow - 1
            .Col = 1
            
            If Abs(val(.text) - FromVolts) < 0.0001 Then
            
                'We're repeating a step at the same voltage
                'as the last voltage of a previously added step set
                'Quick fix - reduce StartRow by 1
                StartRow = StartRow - 1
                
            End If
            
        End With
        
    End If
        
    Volts = FromVolts
    i = StartRow
        
    Do While Volts <= ToVolts And Volts >= FromVolts
    
        'If the coil selected is the axial coil, add
        'a new row to the axial grid
        With gridCalibration
    
            If i >= .Rows Then
            
                .Rows = i + 1
    
            End If
    
            .row = i
            .Col = 1
            
            If AFSystem = "2G" And _
               InAFMode = True _
            Then
                
                .text = Format(Volts, "0")
            
            Else
            
                .text = Format(Volts, "#0.0###")
            
            End If
            
            .RowExpanded = True
            If .RowHeight(i) = 0 Then .RowHeight(i) = 228
            
            .Col = 0
            .text = Trim(str(i))
                    
        End With
        
        i = i + 1
        
        If isLog Then
    
            Volts = Volts * (StepSize) ^ (i - StartRow)
        
        Else
        
            Volts = FromVolts + (i - StartRow) * StepSize
    
        End If
        
        DoEvents
        
    Loop
'-------------------------------------------------------------------------------------------------------------'
'-------------------------------------------------------------------------------------------------------------'
'
'   Commented Out
'   By: Isaac Hilburn
' When: Sept. 14, 2010
'
'   This With...End With block is causing the Add button to overshoot and add one row too many where the value
'   of the last row is greater than the value the user has entered in the To: field above the Calibration grid
'
'-------------------------------------------------------------------------------------------------------------'
'
'    With gridCalibration
'
'        If i >= .Rows Then
'
'            .Rows = i + 1
'
'        End If
'
'        .row = i
'        .Col = 1
'
'        If AFSystem = "2G" And _
'           InAFMode = True _
'        Then
'
'            .text = Format(Volts, "0")
'
'        Else
'
'            .text = Format(Volts, "#0.0###")
'
'        End If
'
'        .RowExpanded = True
'        If .RowHeight(i) = 0 Then .RowHeight(i) = 228
'
'        .Col = 0
'        .text = Trim(Str(i))
'        If .ColWidth(0) < Me.TextWidth(Trim(Str(i))) * 2 Then
'
'            .ColWidth(0) = Me.TextWidth(Trim(Str(i))) * 2
'
'        End If
'
'    End With
'
'-------------------------------------------------------------------------------------------------------------'
'-------------------------------------------------------------------------------------------------------------'
    
    'Resize the entire grid using a 1.5 multiplier
    ResizeGrid Me.gridCalibration, _
               Me, , , , , _
               1.5
    
    'Make sure that the Cell edit text box contains the value
    'in the currently active cell
    With Me.gridCalibration
        CurrentCell(0) = CurrentRow - 1
        CurrentCell(1) = 1
        .row = CurrentRow - 1
        .Col = 1
        Me.txtCellEdit.text = .text
    End With
    
    'Show the bottom-most row in the grid
    If i > 1 Then
        Me.gridCalibration.TopRow = i - 1
    End If
    CurrentRow = i
        
    'Set Unsaved Changes = True
    UnsavedChanges = True
    
End Sub

Private Sub cmdApply_Click()

    Dim TempStr

    'Hide the cell edit text-box
    Me.txtCellEdit.Visible = False
    
    'Get the coil description
    TempStr = GetCoilString
    
    'Prompt user to see if they actually want to overwrite the old calibration values
    UserResponse = MsgBox("Are you sure you want to replace the old " & TempStr & _
                      " coil calibration values with these new values?" & vbNewLine & _
                      vbNewLine & "This could break your system and lead to exceptionally " & _
                      "bad science.", _
                      vbYesNo, _
                      "Warning!")
                      
    'Check for a no response
    If UserResponse = vbNo Then Exit Sub

    'Need to transfer values from this Calibration form to the global array
    SaveGridData
    
    'Need to save the min and max field values for the axial and trans coils
    If InAFMode = True Then
        
        modConfig.AfAxialMin = Me.txtAFAxialMinMonitorVoltage
        modConfig.AfAxialMax = Me.txtAFAxialMaxMonitorVoltage
        
        modConfig.AfTransMin = Me.txtAFTransMinMonitorVoltage
        modConfig.AfTransMax = Me.txtAFTransMaxMonitorVoltage
        
    Else
        
        modConfig.PulseAxialMax = Me.txtAFAxialMaxMonitorVoltage
        modConfig.PulseAxialMin = Me.txtAFAxialMinMonitorVoltage
        
        modConfig.PulseTransMax = Me.txtAFTransMaxMonitorVoltage
        modConfig.PulseTransMin = Me.txtAFTransMinMonitorVoltage
        
    End If
    
    'Set unsaved changes = false
    UnsavedChanges = False
    
End Sub

Private Sub cmdCalibrateTransverseProbeAngle_Click()
    Load frmTransverseProbeAutoPosition
    frmTransverseProbeAutoPosition.Show
End Sub

Private Sub cmdClear_Click()

    If Me.cmdStartCalibration.Caption = "End Calibration" Then Exit Sub

    'Hide the cell edit text-box
    Me.txtCellEdit.Visible = False

    'Clear and reload the headers
    ClearAndLoadHeaders
    
    'Restore # of columns for replicates
    txtNumReplicateRamps_Change
    
    'Set unsaved changes = false
    UnsavedChanges = False
    
    'Reset the current row
    CurrentRow = 1
    
    CurrentCell(0) = 1
    CurrentCell(1) = 1
    With Me.gridCalibration
        .row = CurrentRow
        .Col = 1
        Me.txtCellEdit.text = .text
    End With
        
End Sub

Private Sub cmdClearData_Click()

    Dim NumRows As Long
    Dim NumCols As Long
    Dim i, j As Long
    
    If Me.cmdStartCalibration.Caption = "End Calibration" Then Exit Sub
    
    'Hide the cell edit text-box
    Me.txtCellEdit.Visible = False
    
    'Need to delete the text in cols 2 through N - 1 in
    'rows 1 through N - 1

    NumRows = Me.gridCalibration.Rows
    NumCols = Me.gridCalibration.Cols
    
    With Me.gridCalibration
    
        For i = 1 To NumRows - 1
        
            .row = i
        
            For j = 2 To NumCols - 1
            
                .Col = j
                
                .text = ""
                
            Next j
            
        Next i
        
    End With
    
    'if the number of replicate columns = 0, set them = 3
    If val(Me.txtNumReplicateRamps) = 0 Then Me.txtNumReplicateRamps = "3"
    
    'Reload the headers of each column
    LoadHeaders
    
    With Me.gridCalibration
        CurrentCell(0) = 1
        CurrentCell(1) = 1
        .row = 1
        .Col = 1
        Me.txtCellEdit.text = .text
    End With
    
End Sub

Private Sub cmdClose_Click()

    'Hide the cell edit text-box
    Me.txtCellEdit.Visible = False

    Unload Me

    Me.Hide

End Sub

Private Sub cmdDelete_Click()

    Dim RowStart As Long
    Dim i As Long
    Dim j As Long
    Dim RowEnd As Long
    Dim NumRows As Long

    'Can't delete anything while the calibration is running.
    If Me.cmdStartCalibration.Caption = "End Calibration" Then Exit Sub

    'Hide the cell edit text-box
    Me.txtCellEdit.Visible = False

    With Me.gridCalibration
        
        'Find bounds on the rows to eliminate
        If .row > .RowSel Then
        
            RowStart = .RowSel
            RowEnd = .row
            
        Else
        
            RowStart = .row
            RowEnd = .RowSel
        
        End If
        
        'Use the DeleteRow function
        'Renumber Col #0 and resize the grid
        DeleteRow Me.gridCalibration, _
                  Me, _
                  RowStart, _
                  RowEnd, _
                  True, _
                  True
                  
        'Set the current row to the last row in the grid
        CurrentRow = .Rows - 1
        
        'Reset the currently selected cell row
        CurrentCell(0) = 1
        CurrentCell(1) = 1
        
        'Reset the text in the grid box
        .row = 1
        .Col = 1
        Me.txtCellEdit.text = .text
                
    End With

End Sub

Private Sub cmdLoadFromCSVFile_Click()

    Dim wasLoadSuccessful As Boolean
    
    'Hide the cell edit text-box
    Me.txtCellEdit.Visible = False
    
    wasLoadSuccessful = frmFileSave.LoadAFCalibrationTable(Me.gridCalibration, _
                                                           modConfig.AFUnits)
                                                           
    'Check to see if the load-table operation was successful
    'If not, user needs to change the settings in the AF File
    'Save Settings window
    If wasLoadSuccessful = False Then
    
        Load frmFileSave
        frmFileSave.Show
        
    Else
    
        'Update the number of replicates on the form control
        Me.txtNumReplicateRamps = Trim(str(Me.gridCalibration.Cols - 4)) \ 2
        
        'Save the value in the active row to the txtCellEdit textbox
        With Me.gridCalibration
        
            CurrentCell(0) = 1
            CurrentCell(1) = 1
            .row = 1
            .Col = 1
            Me.txtCellEdit.text = .text
            
        End With
        
        CurrentRow = Me.gridCalibration.Rows
        
    End If
    
    'Set unsaved changes = false
    'Since this data was just loaded from a data file, no need to protect this
    'data from being accidentally erased by switching coils
    UnsavedChanges = False
    
    'Load the headers
    LoadHeaders

End Sub

Private Sub cmdManualOK_Click()

    'Hide the cell edit text-box
    Me.txtCellEdit.Visible = False

    'If the user hasn't entered a field value, then don't do anything
    If Me.txtPeakField = vbNullString Then Exit Sub
    
    'Else, hid the picture box and set UserResponse = vbOK
    Me.picGetCalPeak.Visible = False
    UserResponse = vbOK
    
End Sub

Private Sub cmdManualRedo_Click()

    'Hide the cell edit text-box
    Me.txtCellEdit.Visible = False

    'hid the picture box and set UserResponse = vbRetry
    Me.picGetCalPeak.Visible = False
    UserResponse = vbRetry

End Sub

Private Sub cmdOK_Click()

    Dim TempStr

    'Hide the cell edit text-box
    Me.txtCellEdit.Visible = False

    'Get the coil description
    TempStr = GetCoilString
    
    'Prompt user to see if they actually want to overwrite the old calibration values
    UserResponse = MsgBox("Are you sure you want to replace the old " & TempStr & _
                      " coil calibration values with these new values?" & vbNewLine & _
                      vbNewLine & "This could break your system and lead to exceptionally " & _
                      "bad science.", _
                      vbYesNo, _
                      "Warning!")
                      
    'Check for a no response
    If UserResponse = vbNo Then Exit Sub

    'First transfer values from this Calibration form to the global array
    SaveGridData
    
    'Need to save the min and max field values for the axial and trans coils
    If InAFMode = True Then
        
        modConfig.AfAxialMin = val(Me.txtAFAxialMinMonitorVoltage.text)
        modConfig.AfAxialMax = val(Me.txtAFAxialMaxMonitorVoltage.text)
        
        modConfig.AfTransMin = val(Me.txtAFTransMinMonitorVoltage.text)
        modConfig.AfTransMax = val(Me.txtAFTransMaxMonitorVoltage.text)
        
    Else
        
        modConfig.PulseAxialMax = val(Me.txtAFAxialMaxMonitorVoltage.text)
        modConfig.PulseAxialMin = val(Me.txtAFAxialMinMonitorVoltage.text)
        
        modConfig.PulseTransMax = val(Me.txtAFTransMaxMonitorVoltage.text)
        modConfig.PulseTransMin = val(Me.txtAFTransMinMonitorVoltage.text)
        
    End If
    
    'Second, write all the globals to the .INI file
    modConfig.Config_writeSettingstoINI
    
    'Set Unsaved changes = false
    UnsavedChanges = False

End Sub

Private Sub cmdPauseCalibration_Click()
'NOTE:  This sub does NOTHING if the user has clicked to end the run
'       or no run is currently ongoing (Status = "DONE")

    'Check to see if the CalStatus is "RUNNING"
    'If So, change it to "PAUSED"
    If CalStatus = "RUNNING" Then
    
        CalStatus = "PAUSED"
        
        'Check to see if in IRM mode, if so, trigger the IRM ARM form to re-read the input voltage
        If InAFMode = False Then
            frmIRMARM.IRMAverageVoltageIn
        End If
        
        'Change the caption and color of the button
        Me.cmdPauseCalibration.BackColor = &HFF7F&
        Me.cmdPauseCalibration.Caption = "Resume Calibration"
        
        'Exit the sub to avoid logic loops
        Exit Sub
        
    End If
    
    'Now, if CalStatus = "PAUSED" then need to change
    'CalStatus to "RUNNING" - user has clicked to resume the run
    If CalStatus = "PAUSED" Then
    
        CalStatus = "RUNNING"
        
        'Change the caption and color of the button
        Me.cmdPauseCalibration.BackColor = &HFFFFC0
        Me.cmdPauseCalibration.Caption = "Pause Calibration"
        
    End If
    
End Sub

Private Sub cmdSaveToCSVFile_Click()

    Dim wasSaveSuccessful As Boolean
    Dim CurTime
    
    'Hide the cell edit text-box
    Me.txtCellEdit.Visible = False
    
    CurTime = Now
    
    If InAFMode = True Then
        
        wasSaveSuccessful = frmFileSave.SaveAFCalibrationTable( _
                                            Me.gridCalibration, _
                                            CurTime, _
                                            modConfig.AFUnits)
                                            
    Else
    
        wasSaveSuccessful = frmFileSave.SaveIRMCalibrationTable(Me.gridCalibration, CurTime, modConfig.AFUnits)
                                        
    End If
    
    'Set unsaved changes to false if the save was successful
    If wasSaveSuccessful = True Then
    
        UnsavedChanges = False
        
    End If
                                           
End Sub

Private Sub cmdSort_Click()

    modAF_DAQ.SortGrid Me.gridCalibration, _
                       Me, _
                       1, _
                       Me.gridCalibration.Rows - 1, _
                       1, _
                       Me.gridCalibration.Cols - 1
                       
End Sub

Private Sub cmdStartCalibration_Click()

    Dim i As Long
    Dim j As Long
    Dim TempL As Long
    Dim RampFailCounter As Integer
    Dim StepSize As Long
        
    Dim MaxCoilVolts As Double
    Dim MinCoilVolts As Double
    Dim PeakRampVolt As Double
    Dim SumField As Double
    Dim SumVarField As Double
    Dim AvgField As Double
    Dim StdDevField As Double
    
    Dim PeakField As String
    Dim PeakVoltage As String
    Dim TargetVoltage As String
    Dim ProbeString As String
    Dim MessageString As String
    Dim RangeString As String
    
    Dim PriorAFAnalysis As Boolean
    Dim RangeChanged As Boolean
    
    'If CalStatus = "RUNNING", change to "ENDED"
    If CalStatus = "RUNNING" Then
    
        'Change the Button caption and color
        Me.cmdStartCalibration.Caption = "Start Calibration"
        Me.cmdStartCalibration.BackColor = &H80FF80
        Me.refresh
    
        'If coil selected is an IRM coil, then set IRMChargeInterrupted
        If InAFMode = False Then
        
            frmIRMARM.IRMInterruptCharge
            
        End If
    
        CalStatus = "ENDED"
        
        'Exit the subroutine
        Exit Sub
    
    End If
    
    'If CalStatus is not paused, change to running
    If CalStatus <> "PAUSED" Then
    
        CalStatus = "RUNNING"
    
        'Change the Button caption and color
        Me.cmdStartCalibration.Caption = "End Calibration"
        Me.cmdStartCalibration.BackColor = QBColor(4)
        Me.refresh
    
    End If
    
    'Check to see which mode of calibration activity that we're in
    'running, paused, or end
    If CalStatus = "PAUSED" Then
    
        'Loop until CalStatus = RUNNING or END
        Do
        
            PauseTill timeGetTime() + 200
            
        Loop Until CalStatus <> "PAUSED"
        
    ElseIf CalStatus = "ENDED" Then
    
        'This additional check is in here to catch if the user has clicked
        'end after clicking pause after the code prior to the pause code
        'was executed
    
        'Change the Button caption and color
        Me.cmdStartCalibration.Caption = "Start Calibration"
        Me.cmdStartCalibration.BackColor = &H80FF80
        Me.refresh
    
        'Immediately end the subroutine
        Exit Sub
    
    End If
    
    'If the code has made it this far, then need to lock the coil selection
    CoilsLocked = True
    Me.chkLockCoils.value = Checked
        
    'Store the Current state AF analysis enabled state to PriorAFAnalysis
    PriorAFAnalysis = modConfig.EnableAFAnalysis
    
    'Prompt User to see if they want to turn off AF Analysis mode
    'if it's on
    If modConfig.EnableAFAnalysis = True Then
    
        UserResponse = MsgBox("AF Analysis mode is on.  This will lengthen the time of the " & _
                          "AF calibration substantially." & vbNewLine & vbNewLine & _
                          "Would you like to run the calibration with AF Analysis mode off?", _
                          vbYesNo, _
                          "Warning!")
                          
        'If user answers yes, switch off the AF Analysis mode
        If UserResponse = vbYes Then
        
            modConfig.EnableAFAnalysis = False
            
        End If
        
    End If
    
    'Sort the contents of gridCalibrate prior to running the calibration
    'in case the user loads values out of order
    modAF_DAQ.SortGrid Me.gridCalibration, _
                       Me, _
                       1, _
                       Me.gridCalibration.Rows - 1, _
                       1, _
                       Me.gridCalibration.Cols - 1
        
        
    Dim CoilType As Coil_Type
        
    'Depending on which coil is selected
    'by the user, load that coil's max and min voltages into
    'the two local variables
    If InAFMode = True And _
       ActiveCoilSystem = AxialCoilSystem _
    Then
    
        MaxCoilVolts = modConfig.AfAxialMonMax
        MinCoilVolts = 0
        ProbeString = "Axial"
        CoilType = Coil_Type.Axial
        
    ElseIf InAFMode = True And _
           ActiveCoilSystem = TransverseCoilSystem _
    Then
        
        MaxCoilVolts = modConfig.AfTransMonMax
        MinCoilVolts = 0
        ProbeString = "Horizontal"
        
        CoilType = Coil_Type.Transverse
        
    ElseIf InAFMode = False And _
           ActiveCoilSystem = AxialCoilSystem _
    Then
    
        MaxCoilVolts = modConfig.IRMAxialVoltMax
        MinCoilVolts = 0
        ProbeString = "Axial"
        
        CoilType = Coil_Type.IRMAxial
        
    ElseIf InAFMode = False And _
           ActiveCoilSystem = TransverseCoilSystem _
    Then
    
        MaxCoilVolts = modConfig.IRMTransVoltMax
        MinCoilVolts = 0
        ProbeString = "Trans"
        
        CoilType = Coil_Type.IRMTrans
    
    End If
    
    'Now validate Max and Min coil voltages
    If MaxCoilVolts <= MinCoilVolts And _
       AFSystem = "ADWIN" _
    Then
    
        'Quick Message Box to user
        MsgBox "Max " & GetCoilString & " coil voltage must be larger than the Min voltage." & _
               vbNewLine & "Max Voltage = " & Trim(str(MaxCoilVolts)) & _
               " Volts" & vbNewLine & "Min Voltage = " & Trim(str(MinCoilVolts)) & _
               " Volts", , _
               "Warning!"
               
        'Click the start button again to end the calibration process
        cmdStartCalibration_Click
        
        'Exit this instance of this event subroutine
        Exit Sub
        
    End If
    
    'Make sure both max and min coil voltages are greater than zero
    'Note: We also never want the Max coil voltage to equal zero.
    If (MaxCoilVolts <= 0 Or _
       MinCoilVolts < 0) And _
       AFSystem = "ADWIN" _
    Then
    
        MsgBox "Max and/or Min " & GetCoilString & " coil voltages are less than zero." & _
               vbNewLine & vbNewLine & "Max Voltage = " & Trim(str(MaxCoilVolts)) & _
               " Volts" & vbNewLine & "Min Voltage = " & Trim(str(MinCoilVolts)) & _
               " Volts", , _
               "Warning!"
                
        Exit Sub
        
        'Click the start button again to end the calibration process
        cmdStartCalibration_Click
        
        'Exit this instance of this event subroutine
        Exit Sub
        
    End If
    
    'For good measure, set the AF relays again depending on the AF system being used
    'If this is an IRM calibration, then need to configure the radio buttons on
    'frmIRMARM for the coil selection
    
    Dim resFreq As Double
    
    'AF 2G
    If AFSystem = "2G" Then
    
        frmAF_2G.SetActiveCoilSystem ActiveCoilSystem
        
    'AF Axial + ADWIN
    ElseIf ActiveCoilSystem = AxialCoilSystem And _
           InAFMode = True And _
           AFSystem = "ADWIN" _
    Then
    
        SystemBoards(AFAxialRelay.BoardName).DigitalOutput _
                                             AFAxialRelay, _
                                             True, _
                                             True
                                             
        'store the AF coil res freq
        resFreq = modConfig.AfAxialResFreq
                                             
    'AF Transverse + ADWIN
    ElseIf ActiveCoilSystem = TransverseCoilSystem And _
           InAFMode = True And _
           AFSystem = "ADWIN" _
    Then
    
        SystemBoards(AFTransRelay.BoardName).DigitalOutput _
                                                AFTransRelay, _
                                                True, _
                                                True
                                                
        'store the AF coil res freq
        resFreq = modConfig.AfTransResFreq
                                                
    'IRM Axial
    ElseIf InAFMode = False And _
           ActiveCoilSystem = AxialCoilSystem _
    Then
    
        'Unlock the global coils lock variable
        CoilsLocked = False
        
        'Set the radio buttons on frmIRMARM
        frmIRMARM.optCoil(0).value = True
        
        'Lock the coils again
        CoilsLocked = True
        
    'IRM Transverse (only allowed with the ADWIN AF system
    ElseIf InAFMode = False And _
           ActiveCoilSystem = TransverseCoilSystem And _
           AFSystem = "ADWIN" _
    Then
    
        'Unlock the global coils lock variable
        CoilsLocked = False
        
        'Set the radio buttons on frmIRMARM
        frmIRMARM.optCoil(1).value = True
        
        'Lock the coils again
        CoilsLocked = True
        
    End If
        
    
    'If in automatic mode, attempt to connect to the 908A Gaussmeter
    If ManualMode = False Then
    
        'Load the gaussmeter form without showing it using special public subroutine
        Load frm908AGaussmeter
        
        gaussmeter_log_file_name = "GM908A_datalog_AF" & ProbeString & "Cal_" & Format(Now, "yyyymmdd_hhnnss") & ".csv"
        
        
        mod908AGaussmeter.StartLogDataFile (gaussmeter_log_file_name)
        
        'Prompt User to attach correct probe to the gaussmeter and turn it on
        MsgBox "While the power to the Gaussmeter is turned off and the USB-mini cable " & _
                "is NOT connected, connect the " & ProbeString & " probe." & vbNewLine & _
                "Then re-connect the USB-mini cable and WAIT for the Gaussmeter to switch." & _
                "back on.", , _
                "908A Gaussmeter Setup"
                
        'Wait 4 seconds and Prompt again to make sure the Gaussmeter is all the way on
        'Loop until
        Do
        
            PauseTill timeGetTime() + 500
        
            UserResponse = MsgBox("Is the Gaussmeter all the way on and displaying data?", _
                                    vbYesNoCancel, _
                                    "908A Gaussmeter Setup")
                                    
            'Check to see which mode of calibration activity that we're in
            'running, paused, or end
            If CalStatus = "PAUSED" Then
            
                'Loop until CalStatus = RUNNING or END
                Do
                
                    PauseTill timeGetTime() + 200
                    
                Loop Until CalStatus <> "PAUSED"
                
            ElseIf CalStatus = "ENDED" Then
            
                'Change the Button caption and color
                Me.cmdStartCalibration.Caption = "Start Calibration"
                Me.cmdStartCalibration.BackColor = &H80FF80
                Me.refresh
            
                'Immediately end the subroutine
                Exit Sub
            
            End If
                           
        Loop Until UserResponse <> vbNo
        
        'If user has selected to cancel, then exit the sub-routine
        If UserResponse = vbCancel Then
        
            Exit Sub
            
        End If
                                
        'Now connect the gaussmeter
        TempL = frm908AGaussmeter.Connect
        
        If TempL <= 0 Then
        
            'Gaussmeter must not be connected
            'Prompt user to ask if they would like to continue with
            'calibration in Manual mode
            UserResponse = MsgBox("Unable to communicate with the 908A Gaussmeter." & vbNewLine & _
                              vbNewLine & "Would you like to continue with the " & GetCoilString & _
                              " coil calibration in manual mode?", _
                              vbYesNo, _
                              "Whoops!")
            
            If UserResponse = vbNo Then
            
                'Click the Start Calibration button again to set the calibration
                'process status to "ENDED"
                cmdStartCalibration_Click
            
                'Exit calibration on a 'No' response
                Exit Sub
                
            ElseIf UserResponse = vbYes Then
            
                'Toggle manual calibration mode on
                Me.chkManualCalibration.value = Checked
                ManualMode = True
            
            End If
                    
        End If
        
    
        
                     
    End If
    
    
    'If this is an IRM calibration, turn off the IRM backfield
    'mode on frmIRMARM, then re-zero the IRM box voltage
    If InAFMode = False Then
    
        frmIRMARM.chkBackfield.value = Unchecked
                
        'Fire the IRM at 0 volts to clear it
        frmIRMARM.FireIRM 0
        
    End If
    
    'Need two separate if statements to check for automatic vs manual mode
    'in case the attempt above to connect to the gaussmeter fails and the user
    'selects to continue with the calibration in manual mode
    
    'Again, if in automatic mode, then do more gaussmeter settings config
    If ManualMode = False Then
        
        'Now Change the mode to DC-Peak
        frm908AGaussmeter.optFunction(1).value = True
        
        'Now Set the modconfig.afunits on the Gaussmeter
        frm908AGaussmeter.SetUnits modConfig.AFUnits
        
        'Now set the range to the 0-300 G range
        'Range2 - this range appears to be the best behaved of the three
        'gaussmeter ranges
        LastRange = 2
        frm908AGaussmeter.optRange(2).value = True
        
        'If the clear-data button is enabled on the gaussmeter form, click it
        If frm908AGaussmeter.cmdClearData.Enabled = True Then
        
            frm908AGaussmeter.cmdClearData_Click
            
        End If
        
        'Reset the gaussmeter
        frm908AGaussmeter.cmdResetPeak_Click
        
        PauseTill timeGetTime() + 2000
        
        'NULL the gaussmeter (set current offset to zero)
        frm908AGaussmeter.DoSilentNull
        
        PauseTill timeGetTime() + 3000
    
    Else
    
        'We're in manual mode
        'Prompt User to attach correct probe to the gaussmeter and turn it on
        MsgBox "While the power to the Gaussmeter is turned off connect the " & _
               ProbeString & " probe." & vbNewLine & _
               "Then turn the Gaussmeter power on and wait for it to reload.", , _
                "908A Gaussmeter Setup"
        
        'Determine what the range-string should be
        If modConfig.AFUnits = "G" Then RangeString = "000.0 G"
        If modConfig.AFUnits = "Oe" Then RangeString = "000.0 Oe"
        If modConfig.AFUnits = "mT" Then RangeString = "00.00 mT"
        If modConfig.AFUnits = "kA/m" Then RangeString = "00.00 kA/m"
        
        MessageString = "When the Gaussmeter has finished loading:" & vbNewLine & vbNewLine & _
                         " 1) Place the tip of the " & ProbeString & " probe in the " & _
                         "center of the " & GetCoilString & " coil." & vbNewLine & _
                         " 2) Press the 'Menu' button on the probe." & vbNewLine & _
                         " 3) If the word ""Function"" does not appear in the lower left-hand corner " & _
                         "of the gaussmeter display, then press the 'Next' button until it does." & _
                         vbNewLine & _
                         " 4) With the word ""Function"" showing, click the 'Enter' button." & _
                         vbNewLine & _
                         " 5) Press 'Next' until the display in the lower-left corner reads " & _
                         """DC peak""." & vbNewLine & _
                         " 6) Press 'Enter' to set the gaussmeter to DC Peak Field mode." & _
                         vbNewLine & _
                         " 7) If the gaussmeter is not in " & modConfig.GetLongUnits & " units mode " & _
                         "then click 'Menu' and then click 'Next' until ""Units"" appears in the " & _
                         "lower left-hand corner." & vbNewLine & _
                         " 8) Click 'Enter' to access the ""Units"" menu.  Click next until you see " & _
                         modConfig.GetLongUnits & " displayed in the lower left-hand corner. " & _
                         " 9) Click 'Enter' again to set the units to " & modConfig.GetLongUnits & vbNewLine & _
                         "10) Check the Gaussmeter range. The display should read: """ & RangeString & """" & _
                         ". If not, then click the 'Range' button until the range changes to match: """ & _
                         RangeString & """" & vbNewLine
                         
        MessageString = MessageString & _
                        "11) Click the Reset button on the Gaussmeter and wait for the Displayed DC Peak value " & _
                        "to stabilize (about 1 - 10 seconds)" & vbNewLine & _
                        "12) Click the NULL button on the Gaussmeter and wait for the NULL process to complete" & _
                        vbNewLine & vbNewLine & _
                         "When you've completed all this, click the ""Continue"" button on this window."
                                 
        UserResponse = frmDialog.DialogBox(MessageString, _
                                           "908A Gaussmeter Setup", _
                                           3, _
                                           "Continue", _
                                           "Cancel", _
                                           "Print...")
                                          
        'Check for a negative response
        If UserResponse = vbNo Then
       
            'User clicked cancel, end the calibration and exit the sub
            cmdStartCalibration_Click
            Exit Sub
            
        End If
        
        If UserResponse = vbCancel Then
        
            modFileSave.PrintRichText Me.dlgPrint, _
                                      Me.rtfPrint, _
                                      MessageString, _
                                      "Print AF Manual Calibration Instructions"
                                      
        End If
              
    End If
        
    'Pause 1 second (1000 ms)
    PauseTill timeGetTime() + 1000
    
    'Check to see if the up/down position of the quartz tube is within a cm
    'of the AF coil height
    TempL = frmDCMotors.GetUpDownPos
    
    If Abs(TempL) - modConfig.UpDownMotor1cm > Abs(modConfig.AFPos) Or _
       Abs(TempL) + modConfig.UpDownMotor1cm < Abs(modConfig.AFPos) _
    Then
    
        'Message box user and tell them to move the Quartz glass tube to
        'the AF position
        UserResponse = MsgBox("Up/Down rod is not in the AF region." & _
               vbNewLine & vbNewLine & _
               "Would you like to pause the AF calibration and move the " & _
               "Up/Down rod into position?", _
               vbYesNo, _
               "Whoops!")
               
        If UserResponse = vbYes Then
        
            'Click the pause button on the form
            cmdPauseCalibration_Click
            
        End If
        
    End If
    
    'Check to see which mode of calibration activity that we're in
    'running, paused, or end
    If CalStatus = "PAUSED" Then
    
        'Loop until CalStatus = RUNNING or END
        Do
        
            PauseTill timeGetTime() + 200
            
        Loop Until CalStatus <> "PAUSED"
        
    ElseIf CalStatus = "ENDED" Then
    
        'This additional check is in here to catch if the user has clicked
        'end after clicking pause after the code prior to the pause code
        'was executed
    
        'Change the Button caption and color
        Me.cmdStartCalibration.Caption = "Start Calibration"
        Me.cmdStartCalibration.BackColor = &H80FF80
        Me.refresh
    
        'Immediately end the subroutine
        Exit Sub
    
    End If
    
    'Determine the column step size to use for the grid
    If AFSystem = "ADWIN" Or _
       InAFMode = False _
    Then
    
        StepSize = 2
        
    Else
    
        StepSize = 1
        
    End If
    
    
    'Now go through each row in the selected coil grid sheet
    'and ramp up and down at each while recording the gaussmeter reading
    With Me.gridCalibration
        
        For i = 1 To .Rows - 1
                
            'Set all the sums (field + std dev field) to zero
            SumField = 0
            SumVarField = 0
        
            .Col = 1
            
            If i > .Rows - 1 Then Exit For
            
            .row = i
              
            'Validate the Voltage or 2G counts value
            If ValidateTargetText(.text) = False Then
            
                'Go to the next iteration of the for loop
                i = i + 1
                
                'Check to see if i > .Rows - 1,
                'If so, exit the for loop
                If i > .Rows - 1 Then Exit For
                
            End If
            
            TargetVoltage = .text
              
            'Set the target system value based on the AF system
            'and whether or not this is an AF or IRM calibration
            If AFSystem = "ADWIN" And _
               InAFMode = True _
            Then
                
                'Do Nothing for now
                            
            ElseIf AFSystem = "2G" And _
                   InAFMode = True _
            Then
                
                'This is a 2G ramp
                'Set PeakRampVolt = target 2G count value
                PeakRampVolt = CInt(.text)
                
                If PeakRampVolt > 3999 Then
                
                    PeakRampVolt = 3999
                    
                    'Update the display
                    .text = Trim(str(PeakRampVolt))
                    
                End If
                
                'Set the uncalibrated amplitude
                frmAF_2G.txtUncalAmplitude = Trim(str(PeakRampVolt))
                frmAF_2G.cmdSetUncalAmp_Click
                
                'Set the verbose / debug mode check box to the correct setting
                If Me.chkVerbose.value = Checked Then
                
                    frmAF_2G.chkVerbose.value = Checked
                    
                Else
                
                    frmAF_2G.chkVerbose.value = Unchecked
                    
                End If
                
                PeakVoltage = PeakRampVoltage
                
            ElseIf InAFMode = False Then
            
                'Check to make sure the current volts to fire the IRM pulse at
                'do not exceed the max IRM pulse voltage
                If val(.text) > modConfig.IRMAxialVoltMax Then
                
                    .text = Trim(str(modConfig.IRMAxialVoltMax))
                    
                End If
                
                'Set target voltage on the IRM Form
                frmIRMARM.txtPulseVolts = Trim(.text)
            
            End If
            
            'Now start doing the replicate ramps while getting the peak DC field
            'from the gaussmeter
            For j = 4 To .Cols - 1 Step StepSize
                
                'Allow other events to occur
                DoEvents
                
                'Check to see which mode of calibration activity that we're in
                'running, paused, or end
                If CalStatus = "PAUSED" Then
                
                    'Loop until CalStatus = RUNNING or END
                    Do
                    
                        PauseTill timeGetTime() + 200
                        
                    Loop Until CalStatus <> "PAUSED"
                    
                ElseIf CalStatus = "ENDED" Then
                
                    'Change the Button caption and color
                    Me.cmdStartCalibration.Caption = "Start Calibration"
                    Me.cmdStartCalibration.BackColor = &H80FF80
                    Me.refresh
                            
                    'Immediately end the subroutine
                    Exit Sub
                
                End If
                
                Me.refresh
                
                'Check to see if this calibration is in manual mode
                If ManualMode = True Then
                
                    'Loop through this as many times as needed
                    Do
                
                        'Need to tell user to reset the gaussmeter
                        MsgBox "Please reset the gaussmeter by clicking the 'Reset' button." & vbNewLine & _
                               "Click 'OK' once you have done this."
                        
                        'Confirm reset
                        UserResponse = MsgBox("Has the Gaussmeter fully reset and is it displaying data " & _
                                              "in DC peak field mode?", _
                                              vbYesNoCancel, _
                                              "Confirmation")
                        
                        'Check for a 'Cancel' click
                        If UserResponse = vbCancel Then
                        
                            'This will put calibration flow into the "Ended" state
                            cmdStartCalibration_Click
                            
                            'Exit sub routine
                            Exit Sub
                            
                        End If
                        
                    Loop Until UserResponse = vbYes
                    
                End If
                
                'Reset the Row & Col value to the current row & the col containing
                'the target voltages
                .row = i
                .Col = 1
                
                'Now start the Ramp - depending on the AF System being used
                If AFSystem = "ADWIN" And _
                   InAFMode = True _
                Then
                    
                    Dim PeakHangTime As Double
                    If resFreq <> 0 Then
                        PeakHangTime = Trim(str(1 / val(resFreq) * 10000))
                    Else
                        PeakHangTime = 500
                    End If
                    
                    frmADWIN_AF.ExecuteRamp ActiveCoilSystem, _
                                            val(TargetVoltage), , , _
                                            WaveForms("AFRAMPUP").IORate, _
                                            PeakHangTime, _
                                            False, _
                                            False, _
                                            (Me.chkVerbose.value = Checked)
                    
                    'Get the Peak monitor voltage from the last ramp
                    'Depending on which monitor wave for was used
                    PeakVoltage = Format(WaveForms("AFMONITOR").CurrentVoltage, "0.###")
                    
                ElseIf AFSystem = "2G" And _
                       InAFMode = True _
                Then
                
                    'Execute a combo ramp
                    'the uncalibrated amp and coil have already been set
                    
                    'Set RampFailCounter = 0
                    RampFailCounter = 0
                    
                    'Loop until get good ramp, if has looped 5 times, then
                    'send error and prompt user to cancel the calibration
                    Do While frmAF_2G.ExecuteRamp("C") = False
                    
                        'Update the fail counter
                        RampFailCounter = RampFailCounter + 1
                        
                        'Check to see if the ramp fail counter >=5
                        If RampFailCounter >= 5 Then
                        
                            'Pop-up message box to the user
                            MessageString = "2G AF Ramp has failed five times in a row." & _
                                            vbNewLine & vbNewLine & _
                                            "Would you like to continue trying to ramp?" & _
                                            vbNewLine & vbNewLine & _
                                            "If you answer 'No' the current auto-calibration session " & _
                                            "will be aborted."
                                            
                            UserResponse = frmDialog.DialogBox(MessageString, _
                                                               "AF Error!!", _
                                                               2, _
                                                               "Yes", _
                                                               "No")
                                                               
                            If UserResponse = vbNo Then
                            
                                'Click the start button again to set all the controls to the
                                'Calibration Ended status
                                cmdStartCalibration_Click
                                
                                'Exit this instance of the event handler
                                Exit Sub
                                
                            End If
                            
                        End If
                                                
                    Loop
                    
                    'Peak Voltage has already been set to
                    'the 2G Counts that were used prior to this for loop
                    
                ElseIf InAFMode = False Then
                
                    'Click the uncalibrated volts IRM Fire button
                    frmIRMARM.cmdIRMFire_Click
                    
                    PeakVoltage = frmIRMARM.IRMPeakVoltage
                
                End If
                    
                'Wait half a second
                
                PauseTill timeGetTime() + 500
                
                'Check to see if this calibration is being done in automatic
                'mode (908A gaussmeter hooked up to computer) or manual mode
                '(user enters in gaussmeter peak field values).
                If ManualMode = False Then
                    
                    'This is an automatic calibration
                        
                    'Now collect a data-point from the Gaussmeter
                    frm908AGaussmeter.cmdSampleNow_Click
                    
                    'Now get the last data point converted to a string with
                    'respect to the modconfig.afunits we're using
                    frm908AGaussmeter.ConvertLastData PeakField, modConfig.AFUnits
                    
                    mod908AGaussmeter.LogDataToFile PeakField, _
                                                    CStr(frm908AGaussmeter.CurrentRange), _
                                                    "DC Peak", _
                                                    ProbeString, _
                                                    Now, _
                                                    gaussmeter_log_file_name
                    
                    'Now get rid of the last data point
                    frm908AGaussmeter.cmdClearData_Click
                    
                    'Reset the gaussmeter DC-peak field
                    frm908AGaussmeter.cmdResetPeak_Click
                    
                Else
                
                    'This is a manual calibration
                    'Need to Prompt the user to write the calibration value
                    'into the picture box
                    Me.txtPeakField = vbNullString
                    Me.lblFieldUnits = modConfig.AFUnits
                    Me.picGetCalPeak.Visible = True
                    Me.txtPeakField.SetFocus
                    
                    'Set UserResponse = -1
                    UserResponse = -1
                    
                    'Loop until the user clicks the OK or Re-Do button
                    Do
                    
                        DoEvents
                        
                        'Pause for 20 ms
                        PauseTill timeGetTime() + 20
                    
                    Loop Until UserResponse = vbOK Or _
                               UserResponse = vbRetry
                               
                    PeakField = Me.txtPeakField.text
                        
                End If
                
                'Check for User Response = vbRetry during a manual calibration
                'Which means the user clicked the 'Re-do' button
                'and the ramp needs to be repeated
                If UserResponse = vbRetry And _
                   ManualMode = True _
                Then
                    
                        j = j - StepSize
                    
                Else
                
                    'Check to see if the Range needs to be changed
                    If CheckRange(PeakField) = True Then
                    
                        'The Range has been changed
                        'Redo the measurement
                        
                        .row = i
                        .Col = j
                        .text = Format(val(PeakField), "0.###")
                        
                        .text = ClipHangingDecimal(.text)
                        
                        'Subtract step-size from j
                        j = j - StepSize
                        
                    Else
                    
                        'Check to see if this is a 2G run or an IRM / ADWIN run
                        If AFSystem = "2G" And InAFMode = True Then
                                                    
                            'Add the peak field to the sum field
                            SumField = SumField + val(PeakField)
                            
                            'Write the Peak Field to the appropriate column in the grid-sheet
                            .row = i
                            .Col = j
                            .text = Format(val(PeakField), "0.###")
                            
                            .text = ClipHangingDecimal(.text)
                            
                        Else
                        
                            '----------------------------------------------------------------------'
                            '----------------------------------------------------------------------'
                            '(July 2, 2011 - I Hilburn)
                            'Commenting this out and using the normal statistics because
                            'I'm not certain if this is the correct approach to get reproducable
                            'results
                            '----------------------------------------------------------------------'
                            'Old Code
                            '----------------------------------------------------------------------'
'                            'Instead of averaging the peak fields,
'                            'average the ratio of the peak field to the peak voltage
'                            SumField = SumField + (val(PeakField) / val(PeakVoltage))
                            '----------------------------------------------------------------------'
                            'Replacement Code
                            '----------------------------------------------------------------------'
                            SumField = SumField + val(PeakField)
                            '----------------------------------------------------------------------'
                            '----------------------------------------------------------------------'
                            
                            'Write the Peak Field to the appropriate column in the grid-sheet
                            .row = i
                            .Col = j
                            .text = Format(val(PeakField), "0.###")
                            
                            .text = ClipHangingDecimal(.text)
                                            
                            'Write the Peak Voltage into the correct column
                            .row = i
                            .Col = j + 1
                            .text = Format(val(PeakVoltage), "0.####")
                            
                            .text = ClipHangingDecimal(.text)
                                                                            
                        End If
                        
                        'Change the active column back to Col #1 with the target voltage
                        .Col = 1
                        
                    End If
                    
                End If
                
                'Resize this column of the grid based upon values in rows
                'up to this row, using a 1.5 multiplier value
                ResizeGrid Me.gridCalibration, _
                           Me, , , _
                           j, _
                           j, _
                           1.5
                           
                'Wait User set amount of wait time between using the coils for AF or IRM
                PauseBetweenUseCoils_InSeconds Me.GetBetweenRampPauseTime(val(TargetVoltage), CoilType)
                                
            Next j
            
            'Now take the Sum of the field values and devide it
            'by the number of replicates
            If val(Me.txtNumReplicateRamps) = 0 Then
            
                Me.txtNumReplicateRamps = "3"
                
            End If
            
            AvgField = SumField / val(Me.txtNumReplicateRamps)
            
            'Check to see if this is a 2G AF run or one of the other calibration runs
            If AFSystem = "2G" And InAFMode = True Then
                
                'Now run through the table and save the sum of the
                'variance
                For j = 4 To .Cols - 1 Step StepSize
                    
                    .row = i
                    .Col = j
                    SumVarField = SumVarField + (AvgField - val(.text)) ^ 2
                
                Next j
                
                'Now calculate the standard deviation from the sum of the variances
                StdDevField = Sqr(SumVarField / val(Me.txtNumReplicateRamps))
                
            Else
                '----------------------------------------------------------------------'
                '----------------------------------------------------------------------'
                '(July 2, 2011 - I Hilburn)
                ' Not using the weight average any more, AvgField = Field
                '----------------------------------------------------------------------'
                'This is an IRM or AF ADWIN run, AvgField actually is the average
                'PeakField / PeakVoltage for each run,
                
                'To get the sum of the variance for this, we need to compare the
                'average PeakField / PeakVoltage * Target Voltage to the peak Fields
                'in the table
                '----------------------------------------------------------------------'
                For j = 4 To .Cols - 1 Step StepSize
                
                    .row = i
                    .Col = j
                    
                    'Get the peak field
                    PeakField = .text
                    
                '----------------------------------------------------------------------'
                '----------------------------------------------------------------------'
                '(July 2, 2011 - I Hilburn)
                ' Not using the weight average any more, AvgField = Field
                '----------------------------------------------------------------------'
'                   .Col = j + 1
'
'                    'Get the peak voltage
'                    PeakVoltage = .text
'
'
'                    'We need to compare the average peakfield / peakvoltage ratio with the peak field and peak
'                    'voltage ratios for each column, and THEN multiply the resulting difference by the target voltage
'                    'to get a true peak field variance
'                    SumVarField = SumVarField + ((AvgField - (val(PeakField) / val(PeakVoltage))) * val(TargetVoltage)) ^ 2
'
                '----------------------------------------------------------------------'
                    SumVarField = SumVarField + (AvgField - val(PeakField)) ^ 2
                '----------------------------------------------------------------------'
                '----------------------------------------------------------------------'
                Next j
                
                'Now calculate the standard deviation from the sum of the variances
                StdDevField = Sqr(SumVarField / val(Me.txtNumReplicateRamps))
                
                '----------------------------------------------------------------------'
                '----------------------------------------------------------------------'
                '(July 2, 2011 - I Hilburn)
                ' Not using the weight average any more, AvgField = Field
                '----------------------------------------------------------------------'
'                'To get the true average peak field
'                ' need to multiple this value by the target voltage
'                AvgField = AvgField * val(TargetVoltage)
                '----------------------------------------------------------------------'
                '----------------------------------------------------------------------'
                
            End If
                
                
            'Now write the average field value to the table
            .row = i
            .Col = 2
            .text = Format(AvgField, "0.###")
            
            .text = ClipHangingDecimal(.text)
                                    
            'Now write the standard deviation of the average
            'field value to the table
            .row = i
            .Col = 3
            .text = Format(StdDevField, "0.###")
            
            .text = ClipHangingDecimal(.text)
                
            'Resize 2nd and 3rd column
            'of 1.5
            ResizeGrid Me.gridCalibration, _
                       Me, , , _
                       2, _
                       3, _
                       1.5
                            
        Next i
            
    End With

    'Change the AF Analysis enabled state back to it's pre-calibration value
    modConfig.EnableAFAnalysis = PriorAFAnalysis

    'Change Calibration status to idle
    CalStatus = "DONE"
    
    'Send out an email that the calibration is done
    frmSendMail.MailNotification GetCoilString & " Calibration Done", _
                                 GetCoilString & " coil field calibration has finished. " & _
                                 "Please come and check the machine.", _
                                 CodeGreen, _
                                 True
    
    'Change the start calibration button caption back to "Start Calibration"
    Me.cmdStartCalibration.Caption = "Start Calibration"
    Me.cmdStartCalibration.BackColor = &H80FF80
    
    'Set unsaved changes = true
    UnsavedChanges = True
    
    'Unlock the coil selection
    CoilsLocked = False
    Me.chkLockCoils.value = Unchecked

    'Refresh the form
    Me.refresh
    
End Sub

Private Sub Form_Activate()

    'Make sure the rich text box control is hidden
    Me.rtfPrint.Visible = False

    SetControls

    If InAFMode = True And _
       EnableAF = False _
    Then
        
        'AF's not enabled, cannot calibrate the AF coils
        'Tell user that calibration is turned off, but
        'can still edit values
        MsgBox "The AF module is currently disabled.  AF coil calibration" & _
               " cannot be performed." & vbNewLine & _
               "However, you can edit the values below by hand.", , _
               "Whoops!"
               
        'Disable all the necessary buttons on the form
        Me.cmdAddSteps.Enabled = False
        Me.cmdStartCalibration.Enabled = False
        Me.cmdPauseCalibration.Enabled = False
        
    ElseIf InAFMode = True Then
    
        'Disable all the necessary buttons on the form
        Me.cmdAddSteps.Enabled = True
        Me.cmdStartCalibration.Enabled = True
        Me.cmdPauseCalibration.Enabled = True
        
    End If
    
    'Check to see if the IRM coils are enabled
    If InAFMode = False And _
       (EnableAxialIRM = False And _
        EnableTransIRM = False) _
    Then
        
        'IRM's not enabled, cannot calibrate the IRM coils
        'Tell user that calibration is turned off, but
        'can still edit values
        MsgBox "The IRM modules are currently disabled.  IRM coil calibration" & _
               " cannot be performed." & vbNewLine & _
               "However, you can edit the values below by hand.", , _
               "Whoops!"
               
        'Disable all the necessary buttons on the form
        Me.cmdAddSteps.Enabled = False
        Me.cmdStartCalibration.Enabled = False
        Me.cmdPauseCalibration.Enabled = False
        
    ElseIf InAFMode = False Then
    
        'Disable all the necessary buttons on the form
        Me.cmdAddSteps.Enabled = True
        Me.cmdStartCalibration.Enabled = True
        Me.cmdPauseCalibration.Enabled = True
        
        SetControlsToIRM
        
    End If

    'First propagate the locked coils state
    If CoilsLocked = True Then Me.chkLockCoils.value = Checked
    If CoilsLocked = False Then Me.chkLockCoils.value = Unchecked

    'If the window is activated, need to propagate
    'the current active coil settings to the radio buttons
    If ActiveCoilSystem = AxialCoilSystem Then
    
        optCoil(0).value = True
        
    ElseIf ActiveCoilSystem = TransverseCoilSystem Then
    
        optCoil(1).value = True
        
    Else
    
        optCoil(0).value = False
        optCoil(1).value = False
        
        ActiveCoilSystem = NoCoilSystem
        If AFSystem = "ADWIN" Then
        
            frmADWIN_AF.SetAFRelays
            
        End If
        
    End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    'Check to see if the user has pressed the 'ESC' key
    'If so, need to pause the auto-calibration if the calibration
    'is already running
    If KeyCode = 27 And _
       CalStatus = "RUNNING" _
    Then
    
        'Call the Pause Calibration button click method
        cmdPauseCalibration_Click
    
    End If

End Sub

Public Sub Form_Load()

    Dim i As Long
    Dim NumReplicates As Long
    
    'Set the form window size
    Me.Height = 8310
    Me.Width = 8175
    
    '(March 11, 2011 - I Hilburn)
    'Update all the form's icons with the program Ico file
    SetFormIcon Me
    
    'Make sure the rich text box control is hidden
    Me.rtfPrint.Visible = False
        
    'Make sure the picture box and txtCellEdit are not visible
    Me.picGetCalPeak.Visible = False
    Me.txtCellEdit.Visible = False
    
    'Check to see if the AF's are enabled
    If InAFMode = True And _
       EnableAF = False _
    Then
        
        'AF's not enabled, cannot calibrate the AF coils
               
        'Disable all the necessary buttons on the form
        Me.cmdAddSteps.Enabled = False
        Me.cmdStartCalibration.Enabled = False
        Me.cmdPauseCalibration.Enabled = False
        
        'Load in the min & max Axial and Transverse coil field values
        Me.txtAFAxialMaxMonitorVoltage.text = Trim(str(modConfig.AfAxialMax))
        Me.txtAFAxialMinMonitorVoltage.text = Trim(str(modConfig.AfAxialMin))
        Me.txtAFTransMaxMonitorVoltage.text = Trim(str(modConfig.AfTransMax))
        Me.txtAFTransMinMonitorVoltage.text = Trim(str(modConfig.AfTransMin))
        
    ElseIf InAFMode = True Then
    
        'enable all the necessary buttons on the form
        Me.cmdAddSteps.Enabled = True
        Me.cmdStartCalibration.Enabled = True
        Me.cmdPauseCalibration.Enabled = True
        
        'Load in the min & max Axial and Transverse coil field values
        Me.txtAFAxialMaxMonitorVoltage.text = Trim(str(modConfig.AfAxialMax))
        Me.txtAFAxialMinMonitorVoltage.text = Trim(str(modConfig.AfAxialMin))
        Me.txtAFTransMaxMonitorVoltage.text = Trim(str(modConfig.AfTransMax))
        Me.txtAFTransMinMonitorVoltage.text = Trim(str(modConfig.AfTransMin))
        
    End If
    
    'Check to see if the IRM coils are enabled
    If InAFMode = False And _
       (EnableAxialIRM = False And _
        EnableTransIRM = False) _
    Then
        
        'IRM's not enabled, cannot calibrate the IRM coils
               
        'Disable all the necessary buttons on the form
        Me.cmdAddSteps.Enabled = False
        Me.cmdStartCalibration.Enabled = False
        Me.cmdPauseCalibration.Enabled = False
        
        'Load in the min & max Axial and Transverse coil field values
        Me.txtAFAxialMaxMonitorVoltage.text = Trim(str(modConfig.PulseAxialMax))
        Me.txtAFAxialMinMonitorVoltage.text = Trim(str(modConfig.PulseAxialMin))
        Me.txtAFTransMaxMonitorVoltage.text = Trim(str(modConfig.PulseTransMax))
        Me.txtAFTransMinMonitorVoltage.text = Trim(str(modConfig.PulseTransMin))
        
    ElseIf InAFMode = False Then
    
        'Enable all the necessary buttons on the form
        Me.cmdAddSteps.Enabled = True
        Me.cmdStartCalibration.Enabled = True
        Me.cmdPauseCalibration.Enabled = True
        
        'Load in the min & max Axial and Transverse coil field values
        Me.txtAFAxialMaxMonitorVoltage.text = Trim(str(modConfig.PulseAxialMax))
        Me.txtAFAxialMinMonitorVoltage.text = Trim(str(modConfig.PulseAxialMin))
        Me.txtAFTransMaxMonitorVoltage.text = Trim(str(modConfig.PulseTransMax))
        Me.txtAFTransMinMonitorVoltage.text = Trim(str(modConfig.PulseTransMin))
        
    End If
    
    'First propagate the locked coils state
    If CoilsLocked = True Then Me.chkLockCoils.value = Checked
    If CoilsLocked = False Then Me.chkLockCoils.value = Unchecked
    
    'Set Radio buttons for correct coil system
    If ActiveCoilSystem = AxialCoilSystem Then
    
        optCoil(0).value = True
    
    ElseIf ActiveCoilSystem = TransverseCoilSystem Then
    
        optCoil(1).value = True
        
    Else
    
        'No coil selected
        optCoil(0).value = False
        optCoil(1).value = False
        
        ActiveCoilSystem = NoCoilSystem
        
    End If
            
    SetControls
   
    'Clear out the values in the voltage step add line
    Me.txtStepSize.text = ""
    Me.txtFromVolts.text = ""
    Me.txtToVolts.text = ""
    
    'Set # of replicates to zero
    Me.txtNumReplicateRamps.text = "0"
    NumReplicates = CLng(val(Me.txtNumReplicateRamps))
    txtNumReplicateRamps_Change
    
    'Clear grid and write in the Column Headers
    LoadHeaders
                                                 
    'Refresh the form display
    Me.refresh
    SaveSizes
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    'if X & Y are over the calibration grid, call that event
    If Y >= Me.gridCalibration.Top And _
       Y <= Me.gridCalibration.Top + Me.gridCalibration.Height And _
       X >= Me.gridCalibration.Left And _
       X <= Me.gridCalibration.Left + Me.gridCalibration.Width _
    Then
    
        'Mouse is over the calibration grid
        'activate the calibration grid mousedown event
        gridCalibration_MouseDown Button, Shift, X, Y
        
    Else
    
        'Mouse down is somewhere else in the form
        'Hide txtCellEdit
        Me.txtCellEdit.Visible = False
        
    End If
        
End Sub

Private Sub Form_Resize()
If m_FormWid <> 0 Then
    ResizeControls
End If
End Sub

Private Function GetCoilString() As String

    If ActiveCoilSystem = AxialCoilSystem And _
       InAFMode = True _
    Then
    
        GetCoilString = "AF Axial"
        
    ElseIf ActiveCoilSystem = TransverseCoilSystem And _
           InAFMode = True _
    Then
    
        GetCoilString = "AF Transverse"
    
    ElseIf ActiveCoilSystem = AxialCoilSystem And _
           InAFMode = False _
    Then
    
        GetCoilString = "IRM Axial"
        
    ElseIf ActiveCoilSystem = TransverseCoilSystem And _
           InAFMode = False _
    Then

        GetCoilString = "IRM Transverse"
        
    Else
    
        GetCoilString = ""
        
    End If

End Function

Private Function GetColCount(ByVal NumReplicates As Long) As Long

    If AFSystem = "ADWIN" Or _
       InAFMode = False _
    Then
    
        GetColCount = 4 + NumReplicates * 2
        
    Else
    
        GetColCount = 4 + NumReplicates
        
    End If
    
End Function

Public Function GetBetweenRampPauseTime(ByVal target_monitor_voltage As Double, ByVal target_coil As Coil_Type) As Integer

    Dim min_secs As Integer
    Dim max_secs As Integer
    
    On Error Resume Next
    min_secs = CInt(Me.txtMinPauseBetweenRamps.text)
    
    If Err.number <> 0 Then min_secs = 5
    On Error GoTo 0
    
    On Error Resume Next
    max_secs = CInt(Me.txtMaxPauseBetweenRamps.text)
    
    If Err.number <> 0 Then max_secs = 30
    On Error GoTo 0
    
    If (min_secs < 5) Then min_secs = 5
    If (max_secs > 300) Then max_secs = 300
    
    Me.txtMaxPauseBetweenRamps.text = Format(max_secs, "0")
    Me.txtMinPauseBetweenRamps.text = Format(min_secs, "0")
    
    'Minimum wait time for first 40% of range, then scales up TO MAX at 80% of range
    Dim max_mon_voltage As Double
    
    If target_coil = Coil_Type.Axial Then
        max_mon_voltage = modConfig.AfAxialMonMax
    ElseIf target_coil = Coil_Type.Transverse Then
        max_mon_voltage = modConfig.AfTransMonMax
    ElseIf target_coil = Coil_Type.IRMAxial Then
        max_mon_voltage = modConfig.IRMAxialVoltMax
    ElseIf target_coil = Coil_Type.IRMTrans Then
        max_mon_voltage = modConfig.IRMTransVoltMax
    Else
        'Set max to 10 - will lead to short wait time
        max_mon_voltage = 10
    End If
    
    If 0.4 * max_mon_voltage > target_monitor_voltage Then
    
        GetBetweenRampPauseTime = min_secs
        
    ElseIf (target_monitor_voltage >= 0.8 * max_mon_voltage) Then
    
        GetBetweenRampPauseTime = max_secs
        
    Else
        
        GetBetweenRampPauseTime = (max_secs - min_secs) * (target_monitor_voltage - 0.4 * max_mon_voltage) / (0.4 * max_mon_voltage) + min_secs
        
    End If
    
End Function

Private Function GetRangeString(ByVal Units As String, _
                           ByVal RangeMode As Long) As String
                           
    If Right(Units, 1) = "G" Then
    
        Select Case RangeMode
        
            Case 0
            
                GetRangeString = "00.00 kG"
                
            Case 1
            
                GetRangeString = "0.000 kG"
            
            Case 2
            
                GetRangeString = "000.0 G"
                
            Case 3
            
                GetRangeString = "00.00 G"
        End Select
                
    ElseIf Right(Units, 2) = "Oe" Then
    
        Select Case RangeMode
        
            Case 0
            
                GetRangeString = "00.00 kOe"
                
            Case 1
            
                GetRangeString = "0.000 kOe"
            
            Case 2
            
                GetRangeString = "000.0 Oe"
                
            Case 3
            
                GetRangeString = "00.00 Oe"
        End Select
                
    ElseIf Right(Units, 1) = "T" Then
    
        Select Case RangeMode
            
            Case 0
            
                GetRangeString = "0.000 T "
                
            Case 1
            
                GetRangeString = "000.0 mT"
            
            Case 2
            
                GetRangeString = "00.00 mT"
                
            Case 3
            
                GetRangeString = "000.0 mT"
                
        End Select
                
    ElseIf Right(Units, 3) = "A/m" Then
    
        Select Case RangeMode
            
            Case 0
            
                GetRangeString = "0000 kA/m "
                
            Case 1
            
                GetRangeString = "000.0 kA/m"
            
            Case 2
            
                GetRangeString = "00.00 kA/m"
                
            Case 3
            
                GetRangeString = "0.000 kA/m"
        End Select
                
    End If
                           
End Function


Private Sub gridCalibration_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    'When user clicks grid with the mouse, need to allow them to
    'edit the data cell they have clicked on (and nothing else)
    
    'If an active calibration is running right now, then don't allow the user to do anything
    If Me.cmdStartCalibration.Caption = "End Calibration" Then Exit Sub
       
    'User must left click to enter the cell
    If Button <> vbLeftButton Then Exit Sub
    
    'Now set the active cell to match the mouse-down cell
    With Me.gridCalibration
       
        CurrentCell(0) = .row
        CurrentCell(1) = .Col
       
        'if this is a fixed row or column, exit this sub
        If .Col = 0 Or .row = 0 Then Exit Sub
        
        'Now size and position the edit text box
        Me.txtCellEdit.text = .text
        
        Me.txtCellEdit.Top = .RowPos(CurrentCell(0)) + .Top
        Me.txtCellEdit.Left = .ColPos(CurrentCell(1)) + .Left + 10
        On Error Resume Next
            
            Me.txtCellEdit.Width = .ColWidth(CurrentCell(1))
            
            If Err.number <> 0 Then
            
                'User has selected to click on a cell that is not fully displayed
                'need to deactivate txtCellEdit
                Me.txtCellEdit.Visible = False
                
                Exit Sub
                
            End If
            
        On Error GoTo 0
        
        Me.txtCellEdit.Height = .RowHeight(CurrentCell(0))
        
        'Set current cell position
        CurrentCellPos(0) = CSng(txtCellEdit.Top + txtCellEdit.Height / 2)
        CurrentCellPos(1) = CSng(txtCellEdit.Left + txtCellEdit.Width / 2)
        
        'Show the cell-edit textbox
        Me.txtCellEdit.ZOrder 0
        Me.txtCellEdit.Visible = True
        Me.txtCellEdit.SetFocus
    
    End With

End Sub

Private Sub gridCalibration_Scroll()

    'If active calibration is running, short-circuit this event
    If Me.cmdStartCalibration.Caption = "End Calibration" Then Exit Sub

    'Save the value of the the Edit cell textbox
    With Me.gridCalibration
    
        .row = CurrentCell(0)
        .Col = CurrentCell(1)
        .text = Me.txtCellEdit
        
    End With
    
    'Hide the Edit cell textbox
    Me.txtCellEdit.Visible = False

End Sub

Private Function LoadCalibrationGrid(ByRef gridobj As MSHFlexGrid, _
                                     ByRef CalArray() As Double, _
                                     Optional ByVal CalCount As Long = -1) As Long
                                
    Dim i As Long
    
    With gridobj
    
        'Find the number of rows needed
        'Add one more row for the headers
        If CalCount = -1 Then
    
            .Rows = UBound(CalArray, 1)
        
        Else
        
            .Rows = CalCount + 1
            
        End If
        
        .Cols = 3
            
        For i = 1 To .Rows - 1
            
            'Row #
            .row = i
            .Col = 0
            .text = Trim(str(i))
                        
            'Voltage or 2G value
            .Col = 1
            
            'Format as Integer for 2G AF calibration
            If AFSystem = "2G" And _
               InAFMode = True _
            Then
            
                .text = Format(CalArray(i, 0), "###0")
                
            Else
            
                .text = Format(CalArray(i, 0), "#0.0##")
                
            End If
                    
            'Matching DC Peak Magnetic Field
            .Col = 2
            .text = Format(CalArray(i, 1), "#0.0##")
            
        Next i
        
    End With

    LoadHeaders
    LoadCalibrationGrid = i
    
    
    'If there are no calibration values, then do not resize col #0
    If UBound(CalArray, 1) = 1 Then
        
        ResizeGrid Me.gridCalibration, _
                   Me, , , _
                   1, _
                   gridCalibration.Cols - 1, _
                   1.5
    
    Else
    
        'There are calibration values
        'Resize the entire grid using a 1.5 multiplier
        ResizeGrid Me.gridCalibration, _
                   Me, , , , _
                   1.5
                   
    End If
    
    With Me.gridCalibration
        .row = .Rows - 1
        .Col = 1
        CurrentCell(0) = .row
        CurrentCell(1) = .Col
        Me.txtCellEdit.text = .text
    End With
    
    CurrentRow = i
    CurrentCell(0) = CurrentRow - 1
    CurrentCell(1) = 1
            
    'Set unsaved changes = false
    UnsavedChanges = False

End Function

Private Sub LoadGridData()

    'Select which Coil system's AF or IRM calibration data to display
    If ActiveCoilSystem = AxialCoilSystem And _
       InAFMode = True _
    Then
    
        CurrentRow = LoadCalibrationGrid(Me.gridCalibration, _
                                         AFAxial, _
                                         AFAxialCount)
                                         
    ElseIf ActiveCoilSystem = TransverseCoilSystem And _
           InAFMode = True _
    Then
            
        CurrentRow = LoadCalibrationGrid(Me.gridCalibration, _
                                         AFTrans, _
                                         AFTransCount)
            
    ElseIf ActiveCoilSystem = TransverseCoilSystem And _
           InAFMode = False _
    Then
            
        CurrentRow = LoadCalibrationGrid(Me.gridCalibration, _
                                         PulseAxial, _
                                         PulseAxialCount)
            
    ElseIf ActiveCoilSystem = TransverseCoilSystem And _
           InAFMode = False _
    Then
            
        CurrentRow = LoadCalibrationGrid(Me.gridCalibration, _
                                         PulseTrans, _
                                         PulseTransCount)
                                         
    End If

End Sub

Private Sub LoadHeaders()

    With Me.gridCalibration
                           
        .Cols = 4
        If .Rows = 1 Then .Rows = 2
        .FixedRows = 1
        .FixedCols = 1
                           
        .row = 0
        .Col = 1
        
        
        'Col 1 header is different for
        '2G AF calibrations from all other calibrations
        If AFSystem = "2G" And _
           InAFMode = True _
        Then
            
            .text = "2G Counts"
            
        Else
        
            .text = "Target Volts"
            
        End If
    
        .RowSizingMode = flexRowSizeIndividual
        .RowHeight(0) = 456
    
        .Col = 2
        .text = "Field (" & modConfig.AFUnits & ")"
        
        If .Cols < 4 Then Exit Sub
        
        .Col = 3
        .text = "StDev (" & modConfig.AFUnits & ")"
        
        'Resize columns 1 through 3 using a 1.5 multiplier
        ResizeGrid Me.gridCalibration, _
                   Me, , , _
                   1, _
                   3, _
                   1.5
        
    End With

End Sub


Private Sub optCoil_Click(Index As Integer)

    'Hide the cell edit text-box
    Me.txtCellEdit.Visible = False

    If CoilsLocked = True Then Exit Sub
    
    'First check to see if there is data that there are unsaved
    'changes that the user might lose
    If UnsavedChanges = True Then
    
        'MsgBox the User
        UserResponse = MsgBox("There are unsaved " & GetCoilString & " coil calibration " & _
                              "values in the calibration grid below.  Switching coils will " & _
                              "erase those values." & vbNewLine & vbNewLine & _
                              "Would you like to switch coils anyway?", _
                              vbYesNo, _
                              "Warning!")
                          
        'Check for a 'No' response
        If UserResponse = vbNo Then Exit Sub
        
    End If

    If Index = 0 And _
       optCoil(Index).value = True _
    Then
    'Axial Coil selected
        
        Me.cmdCalibrateTransverseProbeAngle.Enabled = False
        
        'Set the active coil system to Axial
        ActiveCoilSystem = AxialCoilSystem
        
        'Are we in AF calibration or IRM calibration mode?
        If InAFMode = True Then
        
            'If AF Axial calibration values exist
            'load them into gridCalibration
            LoadCalibrationGrid Me.gridCalibration, _
                                modConfig.AFAxial
                                
            'If this is a 2G af system, configure coils
            If AFSystem = "2G" Then
            
                frmAF_2G.SetActiveCoilSystem ActiveCoilSystem
                
            ElseIf AFSystem = "ADWIN" Then
            
                frmADWIN_AF.SetAFRelays
                                
            End If
                                
        Else
        
            'Check to see if this IRM Coil module is enabled
            If (ActiveCoilSystem = AxialCoilSystem And _
                EnableAxialIRM = False) _
            Then
            
                'Tell the user that the IRM for this coil is not
                'enabled
                MsgBox "The Axial IRM module is currently disabled." & vbNewLine & _
                       "IRM calibration for the Axial coil cannot be done, but you can still " & _
                       "edit the calibration values below by hand.", , _
                       "Whoops!"
                       
                'Disable the calibration buttons
                Me.cmdAddSteps.Enabled = False
                Me.cmdStartCalibration.Enabled = False
                Me.cmdPauseCalibration.Enabled = False
                
            Else
            
                'Disable the calibration buttons
                Me.cmdAddSteps.Enabled = True
                Me.cmdStartCalibration.Enabled = True
                Me.cmdPauseCalibration.Enabled = True
                
            End If
        
            'If IRM Axial calibration values exist
            'load them into gridCalibration
            LoadCalibrationGrid Me.gridCalibration, _
                                modConfig.PulseAxial
                       
        End If
            
    ElseIf Index = 1 And _
           optCoil(Index).value = True _
    Then
    'Transverse Coil selected
    
        Me.cmdCalibrateTransverseProbeAngle.Enabled = True
        
        'Set the active coil system to Transverse
        ActiveCoilSystem = TransverseCoilSystem
        
        'Are we in AF calibration or IRM calibration mode?
        If InAFMode = True Then
        
            'If AF Transverse calibration values exist
            'load them into gridCalibration
            LoadCalibrationGrid Me.gridCalibration, _
                                modConfig.AFTrans
                                
            'If this is a 2G af system, configure coils
            If AFSystem = "2G" Then
            
                frmAF_2G.SetActiveCoilSystem ActiveCoilSystem
                
            ElseIf AFSystem = "ADWIN" Then
            
                frmADWIN_AF.SetAFRelays
                                
            End If
                                
        Else
        
            'Check to see if this IRM Coil module is enabled
            If (ActiveCoilSystem = AxialCoilSystem And _
                EnableAxialIRM = False) _
            Then
            
                'Tell the user that the IRM for this coil is not
                'enabled
                MsgBox "The Axial IRM module is currently disabled." & vbNewLine & _
                       "IRM calibration for the Axial coil cannot be done, but you can still " & _
                       "edit the calibration values below by hand.", , _
                       "Whoops!"
                       
                'Disable the calibration buttons
                Me.cmdAddSteps.Enabled = False
                Me.cmdStartCalibration.Enabled = False
                Me.cmdPauseCalibration.Enabled = False
                
            Else
            
                'Disable the calibration buttons
                Me.cmdAddSteps.Enabled = True
                Me.cmdStartCalibration.Enabled = True
                Me.cmdPauseCalibration.Enabled = True
                
            End If
            
            'If IRM transverse calibration values exist
            'load them into gridCalibration
            LoadCalibrationGrid Me.gridCalibration, _
                                modConfig.PulseTrans

        End If
        
    Else
    
        ActiveCoilSystem = NoCoilSystem
        
    End If
            
End Sub

' Arrange the controls for the new size.
Private Sub ResizeControls()
Dim i As Integer
Dim ctl As Control
Dim x_scale As Single
Dim y_scale As Single

    ' Don't bother if we are minimized.
    If WindowState = vbMinimized Then Exit Sub

    ' Get the form's current scale factors.
    x_scale = ScaleWidth / m_FormWid
    y_scale = ScaleHeight / m_FormHgt

    ' Position the controls.
    i = 1
    For Each ctl In Controls
        With m_ControlPositions(i)
            If TypeOf ctl Is line Then
                ctl.X1 = x_scale * .Left
                ctl.Y1 = y_scale * .Top
                ctl.X2 = ctl.X1 + x_scale * .Width
                ctl.Y2 = ctl.Y1 + y_scale * .Height
            ElseIf TypeOf ctl Is CommonDialog Then
                'Do nothing
            Else
            
                ctl.Left = x_scale * .Left
                ctl.Top = y_scale * .Top
                ctl.Width = x_scale * .Width
                If Not (TypeOf ctl Is ComboBox) Then
                    ' Cannot change height of ComboBoxes.
                    ctl.Height = y_scale * .Height
                End If
                On Error Resume Next
                ctl.Font.size = y_scale * .FontSize
                On Error GoTo 0
            
            End If
        End With
        i = i + 1
    Next ctl
End Sub

Private Sub SaveCalibrationGrid(ByRef gridobj As MSHFlexGrid, _
                                     ByRef CalArray() As Double)
                                     
    Dim i As Long
    Dim j As Long
    
    With gridobj
    
        'Need to clear out the Calibration Array
        Erase CalArray()
        
        'Need to see how many valid rows there are in the calibration grid
        'and prune the bad rows out
        For i = 1 To .Rows - 1
                                     
            If i > .Rows - 1 Then Exit For
                                     
            .row = i
            .Col = 1
            
            If ValidateTargetText(.text) = False Then
            
                'Remove this row from the grid
                .RemoveItem .row
                
            End If
            
        Next i
        
        'Need to resize the Calibration array to the number of rows
        ReDim CalArray(.Rows, 2)
                                     
        'Set the initial elements of the calibration array to zero
        CalArray(0, 0) = 0
        CalArray(0, 1) = 0
                                             
        For i = 1 To .Rows - 1
                                     
            .row = i
            .Col = 1
            
            'Save the target monitor voltage/2G counts value
            CalArray(i, 0) = val(.text)
            
            .Col = 2
            
            'Save the resulting peak field
            CalArray(i, 1) = val(.text)
                                
        Next i
        
    End With

    modAF_DAQ.MedianThreeQuickSort_DBL_2D CalArray, 0
                                         
End Sub

Private Sub SaveGridData()

    If ActiveCoilSystem = AxialCoilSystem And _
       InAFMode = True _
    Then
    
        SaveCalibrationGrid Me.gridCalibration, _
                            modConfig.AFAxial
                            
        modConfig.AFAxialCount = UBound(modConfig.AFAxial, 1) - 1
        modConfig.AFAxialCalDone = True
                            
    ElseIf ActiveCoilSystem = TransverseCoilSystem And _
           InAFMode = True _
    Then
                            
        SaveCalibrationGrid Me.gridCalibration, _
                            modConfig.AFTrans
                            
        modConfig.AFTransCount = UBound(modConfig.AFTrans, 1) - 1
        modConfig.AFTransCalDone = True
        
    ElseIf ActiveCoilSystem = AxialCoilSystem And _
           InAFMode = False _
    Then
    
        SaveCalibrationGrid Me.gridCalibration, _
                            modConfig.PulseAxial

        modConfig.PulseAxialCount = UBound(modConfig.PulseAxial, 1) - 1
        modConfig.IRMAxialCalDone = True

    ElseIf ActiveCoilSystem = TransverseCoilSystem And _
           InAFMode = False _
    Then
    
        SaveCalibrationGrid Me.gridCalibration, _
                            modConfig.PulseTrans
                            
        modConfig.PulseTransCount = UBound(modConfig.PulseTrans, 1) - 1
        modConfig.IRMTransCalDone = True

    End If

End Sub

' Save the form's and controls' dimensions.
Private Sub SaveSizes()
Dim i As Integer
Dim ctl As Control

    ' Save the controls' positions and sizes.
    ReDim m_ControlPositions(1 To Controls.Count)
    i = 1
    For Each ctl In Controls
        With m_ControlPositions(i)
            If TypeOf ctl Is line Then
                .Left = ctl.X1
                .Top = ctl.Y1
                .Width = ctl.X2 - ctl.X1
                .Height = ctl.Y2 - ctl.Y1
            ElseIf TypeOf ctl Is CommonDialog Then
                'Skip
                
            Else
                           
                .Left = ctl.Left
                .Top = ctl.Top
                .Width = ctl.Width
                .Height = ctl.Height
                On Error Resume Next
                .FontSize = ctl.Font.size
                On Error GoTo 0
            End If
        End With
        i = i + 1
    Next ctl

    ' Save the form's size.
    m_FormWid = ScaleWidth
    m_FormHgt = ScaleHeight
End Sub

Private Sub SetControls()

    'Figure out which is the selected Coil (Axial, Transverse, IRM-LF, IRM-HF)
    If InAFMode = True Then
    
        SetControlsToAF
        frameCoilSelection.Visible = True
                        
    ElseIf InAFMode = False And _
           (EnableAxialIRM = True Or _
            EnableTransIRM = True Or _
            EnableIRMBackfield = True) _
    Then
        
        SetControlsToIRM
        frameCoilSelection.Visible = True
                              
    Else
        
        'If no IRM modules are activated, then do not load / show this form
        Me.Hide
        
        Exit Sub
        
    End If
    
End Sub

Private Sub SetControlsToAF()

        'Show the AF coil selector frame
    Me.frameCoilSelection.Caption = "AF Coil Selection"
    Me.Caption = "AF Calibration"
        
    'Change the captions on the Max & Min frames and on the form labels
    Me.frameAxialMaxAndMin.Caption = "AF Axial Max / Min Fields"
    Me.frameTransMaxAndMin.Caption = "AF Trans. Max / Min Fields"
              
    'Change Step labels based on AF system being used
    If modConfig.AFSystem = "ADWIN" Then
    
        Me.lblAFVoltStep(0).Caption = "AF Volt Step:"
        Me.lblAFVoltStep(1).Caption = "AF Volt Step:"
        Me.lblNumReplicates.Caption = "# of Replicate AF Ramps per Voltage Step:"
        
    ElseIf AFSystem = "2G" Then
    
        Me.lblAFVoltStep(0).Caption = "2G Counts Step:"
        Me.lblAFVoltStep(1).Caption = "2G Counts Step:"
        Me.lblNumReplicates.Caption = "# of Replicate AF Ramps per 2G value:"
        
    End If
        
End Sub

Private Sub SetControlsToIRM()

    'Change the Captions on the Coil selector frame
    Me.frameCoilSelection.Caption = "IRM Coil Selection"
    Me.Caption = "IRM Calibration"
            
    'Change the captions on the Max & Min frames and on the form labels
    Me.frameAxialMaxAndMin.Caption = "IRM Axial Max / Min Fields"
    Me.frameTransMaxAndMin.Caption = "IRM Trans. Max / Min Fields"
    Me.lblAFVoltStep(0).Caption = "IRM Volt Step:"
    Me.lblAFVoltStep(1).Caption = "IRM Volt Step:"
    Me.lblNumReplicates.Caption = "# of Replicate IRM Pulses per Voltage Step:"
    
    'Set Max and Min Axial & Transverse
    'if IRM calibrations are done, display the min & max IRM fields
    Me.txtAFAxialMaxMonitorVoltage = Trim(str(modConfig.PulseAxialMax))
    Me.txtAFAxialMinMonitorVoltage = Trim(str(modConfig.PulseAxialMin))
    Me.txtAFTransMaxMonitorVoltage = Trim(str(modConfig.PulseTransMax))
    Me.txtAFTransMinMonitorVoltage = Trim(str(modConfig.PulseTransMin))

End Sub

Private Sub Text1_Click()

    'Hide the cell edit text-box
    Me.txtCellEdit.Visible = False

End Sub

Private Sub txtCellEdit_Change()

    'Need to save the contents of the cell edit text-box as they are now to
    'the flex-grid cell that it's ghosting for
    With Me.gridCalibration
    
        .row = CurrentCell(0)
        .Col = CurrentCell(1)
        .text = Me.txtCellEdit.text
        
    End With

End Sub

Private Sub txtCellEdit_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Dim OldPos As Long
    
    With Me.gridCalibration

        If KeyCode = vbEnter Or _
           KeyCode = vbKeyDown _
        Then
        
            'User has selected to shift to the cell below
        
            'Save the current value of txtCellEdit to the flex-grid cell
            'that it's stealthily filling in for
            .row = CurrentCell(0)
            .Col = CurrentCell(1)
            .text = Me.txtCellEdit.text
            
            'Is this cell in the last row?
            'if so, exit the sub without doing anything
            If .row + 1 >= .Rows Then Exit Sub
            
            'Otherwise, there is a row beyond this
            'Activate the mouse-down event for the next row, same col
            'Advance to the next row
            .row = .row + 1
            CurrentCell(0) = .row
            
            gridCalibration_MouseDown vbLeftButton, _
                                      0, _
                                      CurrentCellPos(0), _
                                      CurrentCellPos(1) + .CellHeight
                                      
            Exit Sub
            
        ElseIf KeyCode = vbKeyUp Then
        
            'User has selected to shift to the cell above
        
            'Save the current value of txtCellEdit to the flex-grid cell
            'that it's stealthily filling in for
            .row = CurrentCell(0)
            .Col = CurrentCell(1)
            .text = Me.txtCellEdit.text
            
            'Is this cell in the first data row?
            'if so, exit the sub without doing anything
            If .row <= 1 Then Exit Sub
            
            'Retreat to the prior row
            .row = .row - 1
            CurrentCell(0) = .row
            
            'Activate the mouse-down event for the above row, same col
            gridCalibration_MouseDown vbLeftButton, _
                                      0, _
                                      CurrentCellPos(0), _
                                      CurrentCellPos(1) - .CellHeight
                                      
            Exit Sub
            
        ElseIf KeyCode = vbKeyRight And Shift = vbShiftMask Then
        
            'User has selected to shift to the cell above
        
            'Save the current value of txtCellEdit to the flex-grid cell
            'that it's stealthily filling in for
            .row = CurrentCell(0)
            .Col = CurrentCell(1)
            .text = Me.txtCellEdit.text
            
            'Is this cell in the first data col?
            'if so, exit the sub without doing anything
            If .Col + 1 >= .Cols Then Exit Sub
            
            'Move right one col
            .Col = .Col + 1
            CurrentCell(1) = .Col
            
            'Activate the mouse-down event for the same row, next col
            gridCalibration_MouseDown vbLeftButton, _
                                      0, _
                                      CurrentCellPos(0) - (.CellWidth + .ColWidth(.Col)) / 2, _
                                      CurrentCellPos(1)
                                      
            Exit Sub
            
        ElseIf KeyCode = vbKeyLeft And Shift = vbShiftMask Then
        
            'User has selected to shift to the next cell over
        
            'Save the current value of txtCellEdit to the flex-grid cell
            'that it's stealthily filling in for
            .row = CurrentCell(0)
            .Col = CurrentCell(1)
            .text = Me.txtCellEdit.text
            
            'Is this cell in the last col?
            'if so, exit the sub without doing anything
            If .Col <= 1 Then Exit Sub
            
            'Move left one col
            .Col = .Col - 1
            CurrentCell(1) = .Col
            
            'Activate the mouse-down event for the same row, prior col
            gridCalibration_MouseDown vbLeftButton, _
                                      0, _
                                      CurrentCellPos(0) + (.CellWidth + .ColWidth(.Col)) / 2, _
                                      CurrentCellPos(1)
                                      
            Exit Sub
            
        ElseIf KeyCode = vbKeyPageUp Then
        
            'User has selected to jump up ten cells
            'Save the current value of txtCellEdit to the flex-grid cell
            'that it's stealthily filling in for
            .row = CurrentCell(0)
            .Col = CurrentCell(1)
            .text = Me.txtCellEdit.text
            
            'Is this cell in the first data row?
            'if so, exit the sub without doing anything
            If .row <= 1 Then Exit Sub
            
            'Is this cell in the first ten data rows?
            'if so, only shift up to the first data row and no further
            If .row <= 10 Then
                
                'Store the old position
                OldPos = .row
                
                'Set the active row to row #1
                .row = 1
                CurrentCell(0) = .row
                
                'Activate the mouse-down event for the target row, same col
                gridCalibration_MouseDown vbLeftButton, _
                                          0, _
                                          CurrentCellPos(0), _
                                          CurrentCellPos(1) - (OldPos - 1) * .CellHeight
            
            Else
            
                'It's safe to jump a full 10 rows
                .row = .row - 10
                CurrentCell(0) = .row
                
                'Activate the mouse-down event for the target row, same col
                gridCalibration_MouseDown vbLeftButton, _
                                          0, _
                                          CurrentCellPos(0), _
                                          CurrentCellPos(1) - 10 * .CellHeight
            
            End If
            
        ElseIf KeyCode = vbKeyPageDown Then
        
            'User has selected to jump down ten cells
            'Save the current value of txtCellEdit to the flex-grid cell
            'that it's stealthily filling in for
            .row = CurrentCell(0)
            .Col = CurrentCell(1)
            .text = Me.txtCellEdit.text
            
            'Is this cell in the first data row?
            'if so, exit the sub without doing anything
            If .row + 1 >= .Rows Then Exit Sub
            
            'Is this cell in the first ten data rows?
            'if so, only shift up to the first data row and no further
            If .row >= .Rows - 10 Then
                
                'Store the old row
                OldPos = .Rows
                
                'Set the active row = last row
                .row = .Rows - 1
                CurrentCell(0) = .row
                
                'Activate the mouse-down event for the target row, same col
                gridCalibration_MouseDown vbLeftButton, _
                                          0, _
                                          CurrentCellPos(0), _
                                          CurrentCellPos(1) + (OldPos - 1) * .CellHeight
            
            Else
            
                'It's safe to jump a full 10 rows
                .row = .row + 10
                CurrentCell(0) = .row
                
                'Activate the mouse-down event for the target row, same col
                gridCalibration_MouseDown vbLeftButton, _
                                          0, _
                                          CurrentCellPos(0), _
                                          CurrentCellPos(1) + 10 * .CellHeight
            
            End If
            
        End If
            
    End With
        
End Sub

Private Sub txtFromVolts_Click()

    'Hide the cell edit text-box
    Me.txtCellEdit.Visible = False

End Sub

Private Sub txtNumReplicateRamps_Change()

    Dim i As Long
    Dim NumReplicates As Long

    On Error Resume Next

    NumReplicates = Round(val(txtNumReplicateRamps.text), 0)

    If Err.number <> 0 Then
    
        Exit Sub
        
    End If
    
    On Error GoTo 0
    
    Me.gridCalibration.Cols = GetColCount(NumReplicates)
    
    If NumReplicates = 0 Then
        
        Exit Sub
        
    End If
    
    If AFSystem = "ADWIN" Or _
       InAFMode = False _
    Then
        
        For i = 4 To 4 + (NumReplicates * 2) - 1 Step 2
    
            With Me.gridCalibration
            
                .row = 0
                .Col = i
                .text = "Field #" & Trim(str((i - 3) \ 2 + 1)) & " (" & modConfig.AFUnits & ")"
                
                .row = 0
                .Col = i + 1
                .text = "Max Volt. #" & Trim(str((i - 3) \ 2 + 1)) & " (V)"
                
            End With
        
        Next i
        
    Else
    
        For i = 4 To 4 + NumReplicates - 1
    
            With Me.gridCalibration
            
                .row = 0
                .Col = i
                .text = "Field #" & Trim(str((i - 3) \ 2 + 1)) & " (" & modConfig.AFUnits & ")"
                
            End With
        
        Next i
        
    End If
    
    'Resize grid for columns 4+, using a 1.5 multiplier
    ResizeGrid Me.gridCalibration, _
               Me, , , _
               4, , _
               1.5
        
End Sub

Private Sub txtPeakField_Click()

    'Hide the cell edit text-box
    Me.txtCellEdit.Visible = False

End Sub

Private Sub txtSingleStep_Change()

    On Error GoTo NotNumber:
    
        If val(Me.txtSingleStep.text) < 0 Then
        
            Me.txtSingleStep.text = vbNullString
            
        End If
        
    On Error GoTo 0
    
    'Exit the subroutine if no error has happened
    Exit Sub
    
NotNumber:

    'Blank the contents of the Single step box
    Me.txtSingleStep.text = vbNullString

End Sub

Private Sub txtStepSize_Click()

    'Hide the cell edit text-box
    Me.txtCellEdit.Visible = False

End Sub

Private Sub txtToVolts_Click()

    'Hide the cell edit text-box
    Me.txtCellEdit.Visible = False

End Sub

Private Function ValidateTargetText(ByVal TargetStr As String) As Boolean '

    Dim TempD As Double

    'Default return value = true, text is good
    ValidateTargetText = True

    'if the target text cannot be converted into a number, return false
    On Error Resume Next
    
        TempD = val(TargetStr)
        
        If Err <> 0 Then
        
            'An error occurred, the target text cannot be converted
            'into a double type number
            ValidateTargetText = False
            
            'Exit the function
            Exit Function
            
        End If

    'Resume normal error flow
    On Error GoTo 0

    'If the target text is negative or zero, return false
    If TempD <= 0 Then ValidateTargetText = False
    
    'If the target text exceeds allowed maximum values, then return false
    If (AFSystem = "2G" And _
        TempD > 3999) Or _
       (AFSystem = "ADWIN" And _
        InAFMode = True And _
        ((ActiveCoilSystem = AxialCoilSystem And _
          TempD > modConfig.AfAxialMonMax) Or _
         (ActiveCoilSystem = TransverseCoilSystem And _
          TempD > modConfig.AfTransMonMax))) Or _
       (InAFMode = False And _
        ((ActiveCoilSystem = AxialCoilSystem And _
          TempD > modConfig.IRMAxialVoltMax) Or _
         (ActiveCoilSystem = TransverseCoilSystem And _
          TempD > modConfig.IRMTransVoltMax))) _
    Then
        ValidateTargetText = False
    End If
End Function

