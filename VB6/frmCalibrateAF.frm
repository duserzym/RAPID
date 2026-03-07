VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmCalibrateCoils 
   Caption         =   "AF / IRM Calibration"
   ClientHeight    =   6465
   ClientLeft      =   7380
   ClientTop       =   4395
   ClientWidth     =   7965
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6465
   ScaleWidth      =   7965
   Begin VB.CommandButton cmdPauseCalibration 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Pause Calibration"
      Height          =   372
      Left            =   6120
      MaskColor       =   &H8000000F&
      TabIndex        =   23
      Top             =   6000
      UseMaskColor    =   -1  'True
      Width           =   1692
   End
   Begin VB.CommandButton cmdStartCalibration 
      BackColor       =   &H0080FF80&
      Caption         =   "Start Calibration"
      Height          =   372
      Left            =   4320
      MaskColor       =   &H8000000F&
      TabIndex        =   22
      Top             =   6000
      UseMaskColor    =   -1  'True
      Width           =   1572
   End
   Begin VB.CommandButton cmdClearData 
      Caption         =   "Clear Data"
      Height          =   372
      Left            =   1440
      TabIndex        =   16
      Top             =   5400
      Width           =   972
   End
   Begin VB.CommandButton cmdLoadFromCSVFile 
      Caption         =   "Load from File"
      Height          =   372
      Left            =   6240
      TabIndex        =   19
      Top             =   5400
      Width           =   1332
   End
   Begin VB.CommandButton cmdSaveToCSVFile 
      Caption         =   "Save to file"
      Height          =   372
      Left            =   4680
      TabIndex        =   18
      Top             =   5400
      Width           =   1332
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   372
      Left            =   240
      TabIndex        =   15
      Top             =   5400
      Width           =   972
   End
   Begin VB.CheckBox chkVerbose 
      Caption         =   "Debug Mode?"
      Height          =   252
      Left            =   5040
      TabIndex        =   12
      Top             =   2040
      Width           =   1692
   End
   Begin VB.TextBox txtNumReplicateRamps 
      Height          =   288
      Left            =   3600
      TabIndex        =   11
      Top             =   2040
      Width           =   972
   End
   Begin VB.CommandButton cmdAddSteps 
      Caption         =   "Add"
      Height          =   372
      Left            =   7080
      TabIndex        =   10
      Top             =   1560
      Width           =   612
   End
   Begin VB.TextBox txtFromVolts 
      Height          =   288
      Left            =   4320
      TabIndex        =   8
      Top             =   1560
      Width           =   852
   End
   Begin VB.TextBox txtToVolts 
      Height          =   288
      Left            =   6000
      TabIndex        =   9
      Top             =   1560
      Width           =   852
   End
   Begin VB.Frame frameAxialMaxAndMin 
      Caption         =   "Axial Max / Min Voltages"
      Height          =   1212
      Left            =   2760
      TabIndex        =   28
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
         Height          =   252
         Left            =   240
         TabIndex        =   30
         Top             =   360
         Width           =   492
      End
      Begin VB.Label Label3 
         Caption         =   "Min:"
         Height          =   252
         Left            =   240
         TabIndex        =   29
         Top             =   720
         Width           =   372
      End
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Save to Settings Window"
      Height          =   372
      Left            =   1800
      TabIndex        =   21
      Top             =   6000
      Width           =   2172
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear All"
      Height          =   372
      Left            =   2640
      TabIndex        =   17
      Top             =   5400
      Width           =   972
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Default         =   -1  'True
      Height          =   372
      Left            =   120
      TabIndex        =   20
      Top             =   6000
      Width           =   1332
   End
   Begin VB.Frame frameTransMaxAndMin 
      Caption         =   "Trans. Max / Min Voltages"
      Height          =   1212
      Left            =   5280
      TabIndex        =   25
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
         Height          =   252
         Left            =   240
         TabIndex        =   27
         Top             =   720
         Width           =   372
      End
      Begin VB.Label lblMaxTransVoltage 
         Caption         =   "Max:"
         Height          =   252
         Left            =   240
         TabIndex        =   26
         Top             =   360
         Width           =   492
      End
   End
   Begin VB.CheckBox chkLogScale 
      Caption         =   "Log Scale"
      Height          =   252
      Left            =   2400
      TabIndex        =   7
      Top             =   1560
      Width           =   1092
   End
   Begin VB.TextBox txtStepSize 
      Height          =   288
      Left            =   1320
      TabIndex        =   6
      Top             =   1560
      Width           =   852
   End
   Begin VB.Frame frameAFCoilSelection 
      Caption         =   "AF Coil Selection"
      Height          =   1212
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   2532
      Begin VB.CheckBox chkForceIRMCoil 
         Caption         =   "Lock coil selection"
         Height          =   495
         Left            =   1440
         TabIndex        =   35
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton optAFCoil 
         Caption         =   "Transverse"
         Height          =   252
         Index           =   1
         Left            =   240
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   720
         Width           =   1212
      End
      Begin VB.OptionButton optAFCoil 
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
      Top             =   2520
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   4895
      _Version        =   393216
      Rows            =   10
      FixedCols       =   0
      SelectionMode   =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000A&
      Height          =   3492
      Left            =   0
      TabIndex        =   34
      Top             =   2400
      Width           =   7932
   End
   Begin VB.Label lblNumReplicates 
      Caption         =   "# of Replicate AF Ramps per Voltage Step:"
      Height          =   252
      Left            =   120
      TabIndex        =   33
      Top             =   2040
      Width           =   3252
   End
   Begin VB.Label Label6 
      Caption         =   "To:"
      Height          =   252
      Left            =   5400
      TabIndex        =   32
      Top             =   1560
      Width           =   492
   End
   Begin VB.Label Label5 
      Caption         =   "From:"
      Height          =   252
      Left            =   3720
      TabIndex        =   31
      Top             =   1560
      Width           =   492
   End
   Begin VB.Label lblAFVoltStep 
      Caption         =   "AF Volt Step:"
      Height          =   252
      Left            =   120
      TabIndex        =   24
      Top             =   1560
      Width           =   1092
   End
End
Attribute VB_Name = "frmCalibrateCoils"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public InAFMode As Boolean
Dim CoilString As String
Dim AxialCurrentRow As Long
Dim TransCurrentRow As Long
Dim IRMAxialCurrentRow As Long
Dim IRMTransCurrentRow As Long
Dim LastRange As Long

Dim CalStatus As String

Public Function ClipHangingDecimal(ByVal NumbStr As String) As String

    Dim TempString As String
    
    If Right(NumbStr, 1) = "." Then
    
        TempString = Mid(NumbStr, 1, Len(NumbStr) - 1)
        ClipHangingDecimal = TempString
        
    Else
    
        ClipHangingDecimal = NumbStr
        
    End If

End Function

Private Function CheckRange(ByVal MagField As Double, _
                            Optional ByVal Units As String = vbNullString) As Boolean

    Dim Range3_Max As Double
    Dim Range2_Max As Double
    Dim Range1_Max As Double
    
    If Units = vbNullString Then Units = modConfig.AFUnits
    
    If Right(Units, 1) = "G" Then
    
        Range3_Max = 30
        Range2_Max = 300
        Range1_Max = 3000
        
    ElseIf Right(Units, 1) = "T" Then
    
        Range3_Max = 0.002999
        Range2_Max = 0.02999
        Range1_Max = 0.29999
        
    End If

    'Check to see if we're using the right range
    If val(MagField) > Range3_Max And _
       val(MagField) < Range2_Max And _
       frm908AGaussmeter.optRange(2).Value <> True And _
       LastRange <> 2 _
    Then

        'Store the value of the last Gaussmeter range
        LastRange = frm908AGaussmeter.CurrentRange
        
        'Change the range
        frm908AGaussmeter.optRange(2).Value = True
        
        'Set return value to indicate that the range was changed
        CheckRange = True

    ElseIf val(MagField) > Range2_Max And _
           val(MagField) < Range1_Max And _
           frm908AGaussmeter.optRange(1).Value <> True And _
           LastRange <> 1 _
    Then

        'Store the value of the last Gaussmeter range
        LastRange = frm908AGaussmeter.CurrentRange
        
        'Change the range
        frm908AGaussmeter.optRange(1).Value = True
        
        'Set return value to indicate that the range was changed
        CheckRange = True
        
    ElseIf val(MagField) > Range1_Max And _
           frm908AGaussmeter.optRange(0).Value <> True _
    Then

        'Store the value of the last Gaussmeter range
        LastRange = frm908AGaussmeter.CurrentRange
        
        'Change the range
        frm908AGaussmeter.optRange(0).Value = True
        
        'Set return value to indicate that the range was changed
        CheckRange = True

    Else
    
        CheckRange = False
        
    End If

End Function

Private Sub cmdApply_Click()

    'Show the Settings window
    frmSettings.Show

    'Need to transfer values from this Calibration form to frmSettings
    If ActiveAFCoilSystem = AxialAFCoilSystem And _
       InAFMode = True _
    Then
    
        With frmSettings.grdCalibAxial
        
            'Resize number of data rows in the
            'frmSettings window
            .Rows = Me.gridAFAxialCalibration.Rows
            .Cols = 3
            
            'Set the column headers
            .row = 0
            .Col = 1
            .text = "Volts"
            .ColWidth(1) = frmSettings.TextWidth(.text) * 1.2
            
            .Col = 2
            .text = "Field (" & modConfig.AFUnits & ")"
            .ColWidth(2) = frmSettings.TextWidth(.text) * 1.2
            
            For i = 1 To .Rows - 1
            
                
                .row = i
                
                'Make row number
                .Col = 0
                .text = Trim(Str(i))
                If .ColWidth(0) < frmSettings.TextWidth(.text) * 2 Then
                
                    .ColWidth(0) = frmSettings.TextWidth(.text) * 2
                    
                End If
                
                'Transfer target voltage
                .Col = 1
                Me.gridAFAxialCalibration.row = i
                Me.gridAFAxialCalibration.Col = 1
                .text = Trim(Me.gridAFAxialCalibration.text)
                                
                'Transfer field value
                .Col = 2
                Me.gridAFAxialCalibration.row = i
                Me.gridAFAxialCalibration.Col = 2
                .text = Trim(Me.gridAFAxialCalibration.text)
                
            Next i
    
        End With
        
    ElseIf ActiveAFCoilSystem = TransverseAFCoilSystem And _
           InAFMode = True _
    Then
    
        With frmSettings.grdCalibTrans
        
            'Resize number of data rows in the
            'frmSettings window
            .Rows = Me.gridAFTransverseCalibration.Rows
            .Cols = 3
            
            'Set the column headers
            .row = 0
            .Col = 1
            .text = "Volts"
            .ColWidth(1) = frmSettings.TextWidth(.text) * 1.2
            
            .Col = 2
            .text = "Field (" & modConfig.AFUnits & ")"
            .ColWidth(2) = frmSettings.TextWidth(.text) * 1.2
            
            For i = 1 To .Rows - 1
            
                .row = i
                
                'Add line number
                .Col = 0
                .text = Trim(Str(i))
                If .ColWidth(0) < frmSettings.TextWidth(.text) * 2 Then
                
                    .ColWidth(0) = frmSettings.TextWidth(.text) * 2
                    
                End If
                
                'Transfer target voltage
                .Col = 1
                Me.gridAFTransverseCalibration.row = i
                Me.gridAFTransverseCalibration.Col = 1
                .text = Trim(Me.gridAFTransverseCalibration.text)
                
                'Transfer field value
                .Col = 2
                Me.gridAFTransverseCalibration.row = i
                Me.gridAFTransverseCalibration.Col = 2
                .text = Trim(Me.gridAFTransverseCalibration.text)
                
            Next i
            
        End With
        
    ElseIf InAFMode = False And _
           ActiveAFCoilSystem = AxialAFCoilSystem _
    Then
    
        With frmSettings.grdCalibIRMAxial
        
            'Resize number of data rows in the
            'frmSettings window
            .Rows = Me.gridAFAxialCalibration.Rows
            .Cols = 3
            
            'Set the column headers
            .row = 0
            .Col = 1
            .text = "Volts"
            .ColWidth(1) = frmSettings.TextWidth(.text) * 1.2
            
            .Col = 2
            .text = "Field (" & modConfig.AFUnits & ")"
            .ColWidth(2) = frmSettings.TextWidth(.text) * 1.2
            
            For i = 1 To .Rows - 1
            
                .row = i
                
                'Add line number
                .Col = 0
                .text = Trim(Str(i))
                If .ColWidth(0) < frmSettings.TextWidth(.text) * 2 Then
                
                    .ColWidth(0) = frmSettings.TextWidth(.text) * 2
                    
                End If
                
                'Transfer target voltage
                .Col = 1
                Me.gridIRMAxial.row = i
                Me.gridIRMAxial.Col = 1
                .text = Trim(Me.gridIRMAxial.text)
                
                'Transfer field value
                .Col = 2
                Me.gridIRMAxial.row = i
                Me.gridIRMAxial.Col = 2
                .text = Trim(Me.gridIRMAxial.text)
                
            Next i
    
        End With
        
    ElseIf InAFMode = False And _
           ActiveAFCoilSystem = TransverseAFCoilSystem _
    Then
    
        With frmSettings.grdCalibIRMTrans
        
            'Resize number of data rows in the
            'frmSettings window
            .Rows = Me.gridIRMTrans.Rows
            .Cols = 3
            
            'Set the column headers
            .row = 0
            .Col = 1
            .text = "Volts"
            .ColWidth(1) = frmSettings.TextWidth(.text) * 1.2
            
            .Col = 2
            .text = "Field (" & modConfig.AFUnits & ")"
            .ColWidth(2) = frmSettings.TextWidth(.text) * 1.2
            
            For i = 1 To .Rows - 1
            
                
                .row = i
                
                'Add line number
                .Col = 0
                .text = Trim(Str(i))
                If .ColWidth(0) < frmSettings.TextWidth(.text) * 2 Then
                
                    .ColWidth(0) = frmSettings.TextWidth(.text) * 2
                    
                End If
                
                'Transfer target voltage
                .Col = 1
                Me.gridIRMTrans.row = i
                Me.gridIRMTrans.Col = 1
                .text = Trim(Me.gridIRMTrans.text)
                
                'Transfer field value
                .Col = 2
                Me.gridIRMTrans.row = i
                Me.gridIRMTrans.Col = 2
                .text = Trim(Me.gridIRMTrans.text)
                
            Next i
            
        End With
                
    End If

    'Select the appropriate tab
    If InAFMode = True Then
    
        'Set the settings form to the AF settings tab
        frmSettings.selectTab 4
        
    Else
    
        'Set the Settings form to the IRM Settings tab
        frmSettings.selectTab 6
        
    End If
    
End Sub

Private Sub cmdPauseCalibration_Click()
'NOTE:  This sub does NOTHING if the user has clicked to end the run
'       or no run is currently ongoing (Status = "DONE")

    'Check to see if the CalStatus is "RUNNING"
    'If So, change it to "PAUSED"
    If CalStatus = "RUNNING" Then
    
        CalStatus = "PAUSED"
        
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

Private Sub cmdStartCalibration_Click()

    Dim i As Long
    Dim j As Long
    Dim UserResponse As Long
    
    Dim MaxCoilVolts As Double
    Dim MinCoilVolts As Double
    Dim PeakRampVolt As Double
    Dim SumField As Double
    Dim SumVarField As Double
    Dim AvgField As Double
    Dim StdDevField As Double
    
    Dim PeakField As String
    Dim PeakVoltage As String
    Dim ProbeString As String
    
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
    
    'Store the Current state AF analysis enabled state to PriorAFAnalysis
    PriorAFAnalysis = modConfig.EnableAFAnalysis
    
    'Prompt User to see if they want to turn off AF Analysis mode
    'if it's on
    If modConfig.EnableAFAnalysis = True Then
    
        UserResp = MsgBox("AF Analysis mode is on.  This will lengthen the time of the " & _
                          "AF calibration substantially." & vbNewLine & vbNewLine & _
                          "Would you like to run the calibration with AF Analysis mode off?", _
                          vbYesNo, _
                          "Warning!")
                          
        'If user answers yes, switch off the AF Analysis mode
        If UserResp = vbYes Then
        
            modConfig.EnableAFAnalysis = False
            
        End If
        
    End If
        
    'Depending on which coil is selected
    'by the user, load that coil's max and min voltages into
    'the two local variables
    If InAFMode = True And _
       ActiveAFCoilSystem = AxialAFCoilSystem _
    Then
    
        MaxCoilVolts = modConfig.AfAxialMonMax
        MinCoilVolts = 0
        CoilString = "Axial"
        ProbeString = "Axial"
        
    ElseIf InAFMode = True And _
           ActiveAFCoilSystem = TransverseAFCoilSystem _
    Then
        
        MaxCoilVolts = modConfig.AfTransMonMax
        MinCoilVolts = 0
        CoilString = "Transverse"
        ProbeString = "Horizontal"
        
    ElseIf InAFMode = False And _
           ActiveAFCoilSystem = AxialAFCoilSystem _
    Then
    
        MaxCoilVolts = modConfig.IRMAxialVoltMax
        MinCoilVolts = 0
        CoilString = "IRM"
        ProbeString = "Axial"
        
    ElseIf InAFMode = False And _
           ActiveAFCoilSystem = TransverseAFCoilSystem _
    Then
    
        MaxCoilVolts = modConfig.IRMTransVoltMax
        MinCoilVolts = 0
        CoilString = "IRM"
        ProbeString = "Trans"
    
    End If
    
    'Now validate Max and Min coil voltages
    If MaxCoilVolts <= MinCoilVolts Then
    
        'Quick Message Box to user
        MsgBox "Max " & CoilString & " coil voltage must be larger than the Min voltage." & _
               vbNewLine & "Max Voltage = " & Trim(Str(MaxCoilVolts)) & _
               " Volts" & vbNewLine & "Min Voltage = " & Trim(Str(MinCoilVolts)) & _
               " Volts", , _
               "Warning!"
                
        Exit Sub
        
    End If
    
    'Make sure both max and min coil voltages are greater than zero
    'Note: We also never want the Max coil voltage to equal zero.
    If MaxCoilVolts <= 0 Or MinCoilVolts < 0 Then
    
        MsgBox "Max and/or Min " & CoilString & " coil voltages are less than zero." & _
               vbNewLine & vbNewLine & "Max Voltage = " & Trim(Str(MaxCoilVolts)) & _
               " Volts" & vbNewLine & "Min Voltage = " & Trim(Str(MinCoilVolts)) & _
               " Volts", , _
               "Warning!"
                
        Exit Sub
        
    End If
    
    'Here code would be inserted to grab the resonance frequencies from the values
    'read at the beginning of the p-mag code from the .ini file.
    'Instead, I'm going to temporarily hard-wire in values
    If CoilSelected = Axial Then
    
        Freq = modConfig.AfAxialResFreq
        
    ElseIf CoilSelected = Transverse Then
    
        Freq = modConfig.AfTransResFreq
        
    End If
    
    'Load necessary values into the frmADWIN_AF form text fields
    'for the ramp cycle
    With frmADWIN_AF

        .txtFreq = Trim(Str(Freq))
        .txtRampRate = WaveForms("AFRAMPUP").IORate
        
    End With

    'Load the gaussmeter form without showing it using special public subroutine
    frm908AGaussmeter.LoadForm
    
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
                       
    Loop Until UserResponse <> 7
    
    'If user has selected to cancel, then exit the sub-routine
    If UserResponse = 2 Then
    
        Exit Sub
        
    End If
                            
    'Now connect the gaussmeter
    frm908AGaussmeter.cmdConnectButton_Click
    
    'Now Change the mode to DC-Peak
    frm908AGaussmeter.optFunction(1).Value = True
    
    'Now Set the modconfig.afunits on the Gaussmeter
    frm908AGaussmeter.SetUnits modConfig.AFUnits
    
    'Now set the range to the second highest - this is the best calibrated range
    LastRange = 3
    frm908AGaussmeter.optRange(3).Value = True
    
    'If the clear-data button is enabled on the gaussmeter form, click it
    If frm908AGaussmeter.cmdClearData.Enabled = True Then
    
        frm908AGaussmeter.cmdClearData_Click
        
    End If
        
    'Pause 1 second (1000 ms)
    PauseTill timeGetTime() + 1000
        
    'Now go through each row in the selected coil grid sheet
    'and ramp up and down at each while recording the gaussmeter reading
    If InAFMode = True And _
       ActiveAFCoilSystem = AxialAFCoilSystem _
    Then
                
        With Me.gridAFAxialCalibration
            
            For i = 1 To .Rows - 1
                    
                'Set all the sums (field + std dev field) to zero
                SumField = 0
                SumVarField = 0
            
                .Col = 1
                .row = i
                  
                'Set the target system value based on the AF system
                If AFSystem = "ADWIN" Then
                    
                    frmADWIN_AF.txtMonitorTrigVolt = .text
                
                    'Calculate the Ramp voltage to use
                    If val(.text) > 0.5 Then
                        
                        PeakRampVolt = val(.text) / 10
                        
                    ElseIf val(.text) > 0.1 Then
                    
                        PeakRampVolt = val(.text) / 5
                        
                    Else
                    
                        PeakRampVolt = val(.text) / 3
                        
                    End If
                                        
                    If PeakRampVolt > modConfig.AfAxialRampMax Then
                    
                        PeakRampVolt = modConfig.AfAxialRampMax
                        
                    End If
                    
                    frmADWIN_AF.txtRampPeakVoltage = Trim(Str(PeakRampVolt))
                    
                    'Hange at peak for 10 periods
                    frmADWIN_AF.txtRampPeakDuration = 100000 * (1 / Freq) 'Remember, peak duration is in ms
                    
                    'Set the Ramp up & down slopes, calibrated for the frequency being used
                    frmADWIN_AF.txtRampUpSlope = PeakRampVolt / (modConfig.AfAxialResFreq / Freq)
                    frmADWIN_AF.txtRampDownSlope = PeakRampVolt / (modConfig.AfAxialResFreq / Freq)
                    
                    'If verbose is set on the Calibrate AF form, set it on the MCC Sine Wave form
                    If chkVerbose.Value = Checked Then
                    
                        frmADWIN_AF.checkVerbose.Value = Checked
                        
                    Else
                    
                        frmADWIN_AF.checkVerbose.Value = Unchecked
                                
                    End If
                                
                Else
                
                    'This is a 2G ramp
                    'Set PeakRampVolt = target 2G count value
                    PeakRampVolt = CInt(.text)
                    
                    If PeakRampVolt > 3999 Then
                    
                        PeakRampVolt = 3999
                        
                        'Update the display
                        .text = Trim(Str(PeakRampVolt))
                        
                    End If
                    
                    'Set the uncalibrated amplitude
                    frmAF_2G.txtUncalAmplitude = Trim(Str(PeakRampVolt))
                    frmAF_2G.cmdSetUncalAmp_Click
                    
                    'Set the verbose / debug mode check box to the correct setting
                    If Me.chkVerbose.Value = Checked Then
                    
                        frmAF_2G.chkVerbose.Value = Checked
                        
                    Else
                    
                        frmAF_2G.chkVerbose.Value = Unchecked
                        
                    End If
                    
                    PeakVoltage = PeakRampVoltage
                    
                End If
                
                'Now start doing the replicate ramps while getting the peak DC field
                'from the gaussmeter
                For j = 4 To .Cols - 1 Step 2
                    
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
                    
                    'Now start the Ramp - depending on the AF System being used
                    If AFSystem = "ADWIN" Then
                        
                        frmADWIN_AF.cmdStartRamp_Click
                        
                        'Get the Peak monitor voltage from the last ramp
                        'Depending on which monitor wave for was used
                        PeakVoltage = Format(WaveForms("AFMONITOR").CurrentVoltage, "0.###")
                        
                    ElseIf AFSystem = "2G" Then
                    
                        'Execute a combo ramp
                        'the uncalibrated amp and coil have already been set
                        frmAF_2G.ExecuteRamp "C"
                        
                        'Peak Voltage has already been set to
                        'the 2G Counts that were used prior to this for loop
                        
                    End If
                        
                    'Now collect a data-point from the Gaussmeter
                    frm908AGaussmeter.cmdSampleNow_Click
                    
                    'Now get the last data point converted to a string with
                    'respect to the modconfig.afunits we're using
                    frm908AGaussmeter.ConvertLastData PeakField, modConfig.AFUnits
                    
                    'Add the peak voltage to the sum field
                    SumField = SumField + val(PeakField)
                    
                    'Now get rid of the last data point
                    frm908AGaussmeter.cmdClearData_Click
                    
                    'Reset the gaussmeter DC-peak field
                    frm908AGaussmeter.cmdResetPeak_Click
                    
                    'Wait 200 ms
                    PauseTill timeGetTime() + 200
                    
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
                    
                    RangeChanged = CheckRange(PeakField, modConfig.AFUnits)
                    
                    If RangeChanged = True Then
                    
                        'Decrement J so measurement is repeated
                        j = j - 2
                        
                        'Remove PeakField value from the sum field
                        SumField = SumField - val(PeakField)
                        
                    End If
                    
                Next j
                
                'Now take the Sum of the field values and devide it
                'by the number of replicates
                AvgField = SumField / val(Me.txtNumReplicateRamps)
                
                'Now run through the table and save the sum of the
                'variance
                For j = 4 To .Cols - 1 Step 2
                
                    .row = i
                    .Col = j
                    SumVarField = SumVarField + (AvgField - val(.text)) ^ 2
                
                Next j
                
                'Now calculate the standard deviation from the sum of the variances
                StdDevField = Sqr(SumVarField / val(Me.txtNumReplicateRamps))
                
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
                    
            Next i
            
        End With
                    
    ElseIf InAFMode = False And _
           ActiveAFCoilSystem = AxialAFCoilSystem _
    Then
                
        With Me.gridIRMAxial
            
            For i = 1 To .Rows - 1
                    
                'Set all the sums (field + std dev field) to zero
                SumField = 0
                SumVarField = 0
            
                .Col = 1
                .row = i
                
                
                'Check to make sure the current volts to fire the IRM pulse at
                'do not exceed the max IRM pulse voltage
                If val(.text) > modConfig.IRMAxialVoltMax Then
                
                    .text = Trim(Str(modConfig.IRMAxialVoltMax))
                    
                End If
                
                'Set target voltage on the IRM Form
                frmIRMARM.txtPulseVolts = Trim(.text)
            
                'Set the HF checkbox to unchecked
                frmIRMARM.chkIRMHFCoil.Value = Unchecked
                
                'Set the LF checkbox to checked
                frmIRMARM.chkIRMLFCoil.Value = Checked
                
                'Set the backfield checkbox to unchecked
                frmIRMARM.chkBackfield = Unchecked
            
                'Now start doing the replicate IRM Pulses while measuring
                'the DC peak field on the gaussmeter
                For j = 4 To .Cols - 1 Step 2
                    
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
                
                    'Click the uncalibrated volts IRM Fire button
                    frmIRMARM.cmdIRMFire_Click
                    
                    'Wait 0.5 seconds
                    PauseTill timeGetTime() + 500
                    
                    'Now collect a data-point from the Gaussmeter
                    frm908AGaussmeter.cmdSampleNow_Click
                    
                    'Now get the last data point converted to a string with
                    'respect to the modconfig.afunits we're using
                    frm908AGaussmeter.ConvertLastData PeakField, modConfig.AFUnits
                    
                    'Add the peak voltage to the sum field
                    SumField = SumField + val(PeakField)
                    
                    'Now get rid of the last data point
                    frm908AGaussmeter.cmdClearData_Click
                    
                    'Reset the gaussmeter DC-peak field
                    frm908AGaussmeter.cmdResetPeak_Click
                    
                    'Wait 200 ms
                    PauseTill timeGetTime() + 200
                    
                    'Write the Peak Field to the appropriate column in the grid-sheet
                    .row = i
                    .Col = j
                    .text = Format(val(PeakField), "0.###")
                    
                    .text = ClipHangingDecimal(.text)
                                        
                    'Write the Peak Voltage into the correct column
                    .row = i
                    .Col = j + 1
                    .text = Format(val(frmIRMARM.txtPulseVolts), "0.####")
                    
                    .text = ClipHangingDecimal(.text)
                    
                    RangeChanged = CheckRange(PeakField, modConfig.AFUnits)
                    
                    If RangeChanged = True Then
                    
                        'Decrement J so measurement is repeated
                        j = j - 2
                        
                        'Remove PeakField value from the sum field
                        SumField = SumField - val(PeakField)
                        
                    End If
                    
                Next j
                               
                'Now take the Sum of the field values and devide it
                'by the number of replicates
                AvgField = SumField / val(Me.txtNumReplicateRamps)
                
                'Now run through the table and save the sum of the
                'variance
                For j = 4 To .Cols - 1 Step 2
                
                    .row = i
                    .Col = j
                    SumVarField = SumVarField + (AvgField - val(.text)) ^ 2
                
                Next j
                
                'Now calculate the standard deviation from the sum of the variances
                StdDevField = Sqr(SumVarField / val(Me.txtNumReplicateRamps))
                
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
                
            Next i
            
        End With
        
    ElseIf InAFMode = False And _
           ActiveAFCoilSystem = TransverseAFCoilSystem _
    Then
                
        With Me.gridIRMTrans
            
            For i = 1 To .Rows - 1
                    
                'Set all the sums (field + std dev field) to zero
                SumField = 0
                SumVarField = 0
            
                .Col = 1
                .row = i
                
                
                'Check to make sure the current volts to fire the IRM pulse at
                'do not exceed the max IRM pulse voltage
                If val(.text) > modConfig.IRMTransVoltMax Then
                
                    .text = Trim(Str(modConfig.IRMTransVoltMax))
                    
                End If
                
                'Set target voltage on the IRM Form
                frmIRMARM.txtPulseVolts = Trim(.text)
            
                'Set the HF checkbox to checked
                frmIRMARM.chkIRMHFCoil.Value = Checked
                
                'Set the LF checkbox to unchecked
                frmIRMARM.chkIRMLFCoil.Value = Unchecked
                
                'Set the backfield checkbox to unchecked
                frmIRMARM.chkBackfield = Unchecked
            
                'Now start doing the replicate IRM Pulses while measuring
                'the DC peak field on the gaussmeter
                For j = 4 To .Cols - 1 Step 2
                    
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
                
                    'Click the uncalibrated volts IRM Fire button
                    frmIRMARM.cmdIRMFire_Click
                    
                    'Wait 0.5 seconds
                    PauseTill timeGetTime() + 500
                    
                    'Now collect a data-point from the Gaussmeter
                    frm908AGaussmeter.cmdSampleNow_Click
                    
                    'Now get the last data point converted to a string with
                    'respect to the modconfig.afunits we're using
                    frm908AGaussmeter.ConvertLastData PeakField, modConfig.AFUnits
                    
                    'Add the peak voltage to the sum field
                    SumField = SumField + val(PeakField)
                    
                    'Now get rid of the last data point
                    frm908AGaussmeter.cmdClearData_Click
                    
                    'Reset the gaussmeter DC-peak field
                    frm908AGaussmeter.cmdResetPeak_Click
                    
                    'Wait 200 ms
                    PauseTill timeGetTime() + 200
                    
                    'Write the Peak Field to the appropriate column in the grid-sheet
                    .row = i
                    .Col = j
                    .text = Format(val(PeakField), "0.###")
                    
                    .text = ClipHangingDecimal(.text)
                                        
                    'Write the Peak Voltage into the correct column
                    .row = i
                    .Col = j + 1
                    .text = Format(val(frmIRMARM.txtPulseVolts), "0.####")
                    
                    .text = ClipHangingDecimal(.text)
                    
                    RangeChanged = CheckRange(PeakField, modConfig.AFUnits)
                    
                    If RangeChanged = True Then
                    
                        'Decrement J so measurement is repeated
                        j = j - 2
                        
                        'Remove PeakField value from the sum field
                        SumField = SumField - val(PeakField)
                        
                    End If
                    
                Next j
                               
                'Now take the Sum of the field values and devide it
                'by the number of replicates
                AvgField = SumField / val(Me.txtNumReplicateRamps)
                
                'Now run through the table and save the sum of the
                'variance
                For j = 4 To .Cols - 1 Step 2
                
                    .row = i
                    .Col = j
                    SumVarField = SumVarField + (AvgField - val(.text)) ^ 2
                
                Next j
                
                'Now calculate the standard deviation from the sum of the variances
                StdDevField = Sqr(SumVarField / val(Me.txtNumReplicateRamps))
                
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
                
            Next i
            
        End With
        
    ElseIf InAFMode = True And _
           ActiveAFCoilSystem = TransverseAFCoilSystem _
    Then
    
        With Me.gridAFTransverseCalibration
            
            For i = 1 To .Rows - 1
            
                'Set all the sums (field + std dev field) to zero
                SumField = 0
                SumVarField = 0
            
                .Col = 1
                .row = i
                
                'If this is an ADWIN calibration ramp, need to setup
                'the values in the ADWIN form
                If modConfig.AFSystem = "ADWIN" Then
                    
                    frmADWIN_AF.txtFreq = modConfig.AfTransResFreq
                    frmADWIN_AF.txtMonitorTrigVolt = .text
                    
                    'Calculate the Ramp voltage to use
                    If val(.text) > 0.5 Then
                        
                        PeakRampVolt = val(.text) / 10
                        
                    ElseIf val(.text) > 0.1 Then
                    
                        PeakRampVolt = val(.text) / 5
                        
                    Else
                    
                        PeakRampVolt = val(.text) / 3
                        
                    End If
                    
                    If PeakRampVolt > modConfig.AfTransRampMax Then
                    
                        PeakRampVolt = modConfig.AfTransRampMax
                        
                    End If
                    
                    frmADWIN_AF.txtRampPeakVoltage = Trim(Str(PeakRampVolt))
                    
                    'Hange at peak for 10 periods
                    frmADWIN_AF.txtRampPeakDuration = 100000 * (1 / Freq) 'Remember, peak duration is in ms
                    
                    'Set the Ramp up & down slopes, calibrated for the frequency being used
                    frmADWIN_AF.txtRampUpSlope = PeakRampVolt / (modConfig.AfAxialResFreq / Freq)
                    frmADWIN_AF.txtRampDownSlope = PeakRampVolt / (modConfig.AfAxialResFreq / Freq)
                    
                    'If verbose is set on the Calibrate AF form, set it on the MCC Sine Wave form
                    If chkVerbose.Value = Checked Then
                    
                        frmADWIN_AF.checkVerbose.Value = Checked
                        
                    Else
                    
                        frmADWIN_AF.checkVerbose.Value = Unchecked
                                
                    End If
                                
                Else
                
                    'This is a 2G ramp
                    'Set PeakRampVolt = target 2G count value
                    PeakRampVolt = CInt(.text)
                    
                    If PeakRampVolt > 3999 Then
                    
                        PeakRampVolt = 3999
                        
                        'Update the display
                        .text = Trim(Str(PeakRampVolt))
                        
                    End If
                    
                    'Set the uncalibrated amplitude
                    frmAF_2G.txtUncalAmplitude = Trim(Str(PeakRampVolt))
                    frmAF_2G.cmdSetUncalAmp_Click
                    
                    'Set the verbose / debug mode check box to the correct setting
                    If Me.chkVerbose.Value = Checked Then
                    
                        frmAF_2G.chkVerbose.Value = Checked
                        
                    Else
                    
                        frmAF_2G.chkVerbose.Value = Unchecked
                        
                    End If
                    
                    PeakVoltage = PeakRampVoltage
                    
                End If
                
                'Now start doing the replicate ramps while getting the peak DC field
                'from the gaussmeter
                'Now start doing the replicate ramps while getting the peak DC field
                'from the gaussmeter
                For j = 4 To .Cols - 1 Step 2
                    
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
                                        
                    'Now start the Ramp - depending on the AF System being used
                    If AFSystem = "ADWIN" Then
                        
                        frmADWIN_AF.cmdStartRamp_Click
                        
                        'Get the Peak monitor voltage from the last ramp
                        PeakVoltage = Format(WaveForms("AFMONITOR").CurrentVoltage, "0.###")
                        
                    ElseIf AFSystem = "2G" Then
                    
                        'Execute a combo ramp
                        'the uncalibrated amp and coil have already been set
                        frmAF_2G.ExecuteRamp "C"
                        
                        'Peak Voltage has already been set to
                        'the 2G Counts that were used prior to this for loop
                        
                    End If
                    
                    'Now collect a data-point from the Gaussmeter
                    frm908AGaussmeter.cmdSampleNow_Click
                    
                    'Now get the last data point converted to a string with
                    'respect to the modconfig.afunits we're using
                    frm908AGaussmeter.ConvertLastData PeakField, modConfig.AFUnits
                    
                    'Add the peak voltage to the sum field
                    SumField = SumField + val(PeakField)
                    
                    'Now get rid of the last data point
                    frm908AGaussmeter.cmdClearData_Click
                    
                    'Reset the gaussmeter by changing to DC mode then back
                    'to DC-Peak
                    frm908AGaussmeter.cmdResetPeak_Click
                    
                    'Wait 200 ms
                    PauseTill timeGetTime() + 200
                    
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
                    
'                    Debug.Print modConfig.AFUnits
                    RangeChanged = CheckRange(val(PeakField), modConfig.AFUnits)
                    
                    If RangeChanged = True Then
                    
                        'Decrement J so measurement is repeated
                        j = j - 2
                        
                        'Remove PeakField value from the sum field
                        SumField = SumField - val(PeakField)
                        
                    End If
                    
                Next j
                                
                'Now take the Sum of the field values and devide it
                'by the number of replicates
                AvgField = SumField / val(Me.txtNumReplicateRamps)
                
                'Now run through the table and save the sum of the
                'variance
                For j = 4 To .Cols - 1 Step 2
                
                    .row = i
                    .Col = j
                    SumVarField = SumVarField + (AvgField - val(.text)) ^ 2
                
                Next j
                
                'Now calculate the standard deviation from the sum of the variances
                StdDevField = Sqr(SumVarField / val(Me.txtNumReplicateRamps))
                
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
                
            Next i
            
        End With
        
    End If

    'Change the AF Analysis enabled state back to it's pre-calibration value
    modConfig.EnableAFAnalysis = PriorAFAnalysis

    'Change Calibration status to idle
    CalStatus = "DONE"
    
    'Change the start calibration button caption back to "Start Calibration"
    Me.cmdStartCalibration.Caption = "Start Calibration"
    Me.cmdStartCalibration.BackColor = &H80FF80
    Me.refresh

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
    
    'Read out values into local variables for from and to volts
    FromVolts = val(Me.txtFromVolts)
    ToVolts = val(Me.txtToVolts)
    
    If AFSystem = "ADWIN" Or _
       CoilSelected = IRMLF Or IRMHF _
    Then
    
        'Depending on which coil is selected
        'by the user, load that coil's max and min voltages into
        'the two local variables
        If CoilSelected = Axial Then
        
            MaxCoilVolts = modConfig.AfAxialMonMax
            MinCoilVolts = 0
            CoilString = "Axial"
            
        ElseIf CoilSelected = Transverse Then
            
            MaxCoilVolts = modConfig.AfTransMonMax
            MinCoilVolts = 0
            CoilString = "Transverse"
            
        ElseIf CoilSelected = IRMLF Then
        
            MaxCoilVolts = modConfig.IRMAxialVoltMax
            MinCoilVolts = 0
            CoilString = "IRM Low-field"
            
        ElseIf CoilSelected = IRMHF Then
        
            MaxCoilVolts = modConfig.IRMTransVoltMax
            MinCoilVolts = 0
            CoilString = "IRM Hi-Field"
            
        End If
        
        'Now validate Max and Min coil voltages
        If MaxCoilVolts <= MinCoilVolts Then
        
            'Quick Message Box to user
            MsgBox "Max " & CoilString & " coil voltage must be larger than the Min voltage." & _
                    vbNewLine & vbNewLine & "Max Voltage = " & Trim(Str(MaxCoilVolts)) & _
                    " Volts" & vbNewLine & "Min Voltage = " & Trim(Str(MinCoilVolts)) & _
                    " Volts", , _
                    "Warning!"
                    
            Exit Sub
            
        End If
        
        'Make sure both max and min coil voltages are greater than zero
        'Note: We also never want the Max coil voltage to equal zero.
        If MaxCoilVolts <= 0 Or MinCoilVolts < 0 Then
        
            MsgBox "Max and/or Min " & CoilString & " coil voltages are less than zero." & _
                    vbNewLine & vbNewLine & "Max Voltage = " & Trim(Str(MaxCoilVolts)) & _
                    " Volts" & vbNewLine & "Min Voltage = " & Trim(Str(MinCoilVolts)) & _
                    " Volts", , _
                    "Warning!"
                    
            Exit Sub
            
        End If
        
        isLog = False
        
        If chkLogScale.Value = Checked Then
        
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
            txtFromVolts.text = Trim(Str(FromVolts))
            
        End If
        
        If ToVolts > MaxCoilVolts Then
        
            ToVolts = MaxCoilVolts
            txtToVolts.text = Trim(Str(ToVolts))
            
        End If
        
        'Set the StepSize local variable
        StepSize = val(Me.txtStepSize)
        
        If FromVolts > ToVolts And StepSize > 0 Then
        
            MsgBox "From Voltage must be smaller than To voltage with a positive volt step size." & _
                    vbNewLine & "From = " & Trim(Str(FromVolts)) & " Volts" & vbNewLine & _
                    "To = " & Trim(Str(ToVolts)) & " Volts" & vbNewLine & "Step Size = " & _
                    Trim(Str(StepSize))
                    
            Exit Sub
    
        End If
        
        If FromVolts < ToVolts And StepSize < 0 Then
        
            MsgBox "From Voltage must be larger than To voltage with a negative volt step size." & _
                    vbNewLine & "From = " & Trim(Str(FromVolts)) & " Volts" & vbNewLine & _
                    "To = " & Trim(Str(ToVolts)) & " Volts" & vbNewLine & "Step Size = " & _
                    Trim(Str(StepSize))
    
            Exit Sub
    
        End If
                
        'Check the number of replicates to run
        If val(Me.txtNumReplicateRamps) < 3 Then
        
            Me.txtNumReplicateRamps = "3"
            
        End If
        
        'Store to local variable the number of replicate times
        'to ramp up and measure the field
        NumReplicates = CLng(val(Me.txtNumReplicateRamps))
        
        'change grid for selected coil
        If CoilSelected = Axial Then
            
            Me.gridAFAxialCalibration.Cols = 4 + NumReplicates * 2
        
        ElseIf CoilSelected = Transverse Then
        
            Me.gridAFTransverseCalibration.Cols = 4 + NumReplicates * 2
            
        ElseIf CoilSelected = IRMLF Then
        
            Me.gridIRMAxial.Cols = 4 + NumReplicates * 2
            
        ElseIf CoilSelected = IRMHF Then
        
            Me.gridIRMTrans.Cols = 4 + NumReplicates * 2
            
        End If
            
        For i = 4 To 4 + (NumReplicates * 2) - 1 Step 2
        
            If CoilSelected = Axial Then
            
                With Me.gridAFAxialCalibration
                
                    .row = 0
                    .Col = i
                    .text = "Field #" & Trim(Str((i - 3) \ 2 + 1)) & " (" & modConfig.AFUnits & ")"
                    If .ColWidth(i) < Me.TextWidth(.text) * 1.2 Then .ColWidth(i) = Me.TextWidth(.text) * 1.2
                    
                    .row = 0
                    .Col = i + 1
                    .text = "Max Volts #" & Trim(Str((i - 3) \ 2 + 1)) & ""
                    If .ColWidth(i) < Me.TextWidth(.text) * 1.2 Then .ColWidth(i) = Me.TextWidth(.text) * 1.2
                    
                End With
                
            ElseIf CoilSelected = Transverse Then
            
                With Me.gridAFTransverseCalibration
                
                    .row = 0
                    .Col = i
                    .text = "Field #" & Trim(Str((i - 3) \ 2 + 1)) & " (" & modConfig.AFUnits & ")"
                    If .ColWidth(i) < Me.TextWidth(.text) * 1.2 Then .ColWidth(i) = Me.TextWidth(.text) * 1.2
                    
                    .row = 0
                    .Col = i + 1
                    .text = "Max Volts #" & Trim(Str((i - 3) \ 2 + 1)) & ""
                    If .ColWidth(i) < Me.TextWidth(.text) * 1.2 Then .ColWidth(i) = Me.TextWidth(.text) * 1.2
                    
                End With
                
            ElseIf CoilSelected = IRMLF Then
            
                With Me.gridIRMAxial
                
                    .row = 0
                    .Col = i
                    .text = "Field #" & Trim(Str((i - 3) \ 2 + 1)) & " (" & modConfig.AFUnits & ")"
                    If .ColWidth(i) < Me.TextWidth(.text) * 1.2 Then .ColWidth(i) = Me.TextWidth(.text) * 1.2
                    
                    .row = 0
                    .Col = i + 1
                    .text = "Max Volts #" & Trim(Str((i - 3) \ 2 + 1)) & ""
                    If .ColWidth(i) < Me.TextWidth(.text) * 1.2 Then .ColWidth(i) = Me.TextWidth(.text) * 1.2
                    
                End With
                
            ElseIf CoilSelected = IRMHF Then
            
                With Me.gridIRMTrans
                
                    .row = 0
                    .Col = i
                    .text = "Field #" & Trim(Str((i - 3) \ 2 + 1)) & " (" & modConfig.AFUnits & ")"
                    If .ColWidth(i) < Me.TextWidth(.text) * 1.2 Then .ColWidth(i) = Me.TextWidth(.text) * 1.2
                    
                    .row = 0
                    .Col = i + 1
                    .text = "Max Volts #" & Trim(Str((i - 3) \ 2 + 1)) & ""
                    If .ColWidth(i) < Me.TextWidth(.text) * 1.2 Then .ColWidth(i) = Me.TextWidth(.text) * 1.2
                    
                End With
                
            End If
                
            
        Next i
    
    Else
    
        'This is a 2G calibration run
    
        'Recast from and to and step size values into the nearest integer values
        FromVolts = CInt(FromVolts)
        ToVolts = CInt(ToVolts)
        StepSize = CInt(StepSize)
            
        'Depending on which coil is selected
        'by the user, set the correct coil-string
        If CoilSelected = Axial Then
        
            CoilString = "Axial"
            
        Else
        
            CoilString = "Transverse"
            
        End If
                
        isLog = False
        
        If chkLogScale.Value = Checked Then
        
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
            txtFromVolts.text = Trim(Str(FromVolts))
            
        End If
        
        'If to voltage > 3999 2G counts, set it equal to 3999
        If ToVolts > 3999 Then
        
            ToVolts = 3999
            txtToVolts.text = Trim(Str(ToVolts))
            
        End If
        
        'Set the StepSize local variable
        StepSize = val(Me.txtStepSize)
                
        If FromVolts > ToVolts And StepSize > 0 Then
        
            MsgBox "From 2G counts must be smaller than To 2G counts with a positive step size." & _
                    vbNewLine & "From = " & Trim(Str(FromVolts)) & vbNewLine & _
                    "To = " & Trim(Str(ToVolts)) & vbNewLine & "Step Size = " & _
                    Trim(Str(StepSize))
                    
            Exit Sub
    
        End If
        
        If FromVolts < ToVolts And StepSize < 0 Then
        
            MsgBox "From 2G counts must be larger than To 2G counts with a negative step size." & _
                    vbNewLine & "From = " & Trim(Str(FromVolts)) & vbNewLine & _
                    "To = " & Trim(Str(ToVolts)) & "Step Size = " & _
                    Trim(Str(StepSize))
    
            Exit Sub
    
        End If
                
        'Check the number of replicates to run
        If val(Me.txtNumReplicateRamps) < 3 Then
        
            Me.txtNumReplicateRamps = "3"
            
        End If
        
        'Store to local variable the number of replicate times
        'to ramp up and measure the field
        NumReplicates = CLng(val(Me.txtNumReplicateRamps))
        
        'change grid for selected coil
        If CoilSelected = Axial Then
            
            Me.gridAFAxialCalibration.Cols = 4 + NumReplicates * 2
        
        Else
        
            Me.gridAFTransverseCalibration.Cols = 4 + NumReplicates * 2
            
        End If
            
        For i = 4 To 4 + (NumReplicates * 2) - 1 Step 2
        
            If CoilSelected = Axial Then
            
                With Me.gridAFAxialCalibration
                
                    .row = 0
                    .Col = i
                    .text = "Field #" & Trim(Str((i - 3) \ 2 + 1)) & " (" & modConfig.AFUnits & ")"
                    If .ColWidth(i) < Me.TextWidth(.text) * 1.2 Then .ColWidth(i) = Me.TextWidth(.text) * 1.2
                    
                    .row = 0
                    .Col = i + 1
                    .text = "2G Counts #" & Trim(Str((i - 3) \ 2 + 1)) & ""
                    If .ColWidth(i) < Me.TextWidth(.text) * 1.2 Then .ColWidth(i) = Me.TextWidth(.text) * 1.2
                
                End With
                
            ElseIf CoilSelected = Transverse Then
            
                With Me.gridAFTransverseCalibration
                
                    .row = 0
                    .Col = i
                    .text = "Field #" & Trim(Str((i - 3) \ 2 + 1)) & " (" & modConfig.AFUnits & ")"
                    If .ColWidth(i) < Me.TextWidth(.text) * 1.2 Then .ColWidth(i) = Me.TextWidth(.text) * 1.2
                    
                    .row = 0
                    .Col = i + 1
                    .text = "2G Counts #" & Trim(Str((i - 3) \ 2 + 1)) & ""
                    If .ColWidth(i) < Me.TextWidth(.text) * 1.2 Then .ColWidth(i) = Me.TextWidth(.text) * 1.2
                
                End With
                
            End If
                
            
        Next i
    
    End If
    
       
    'Number of rows to add
    N = Round((ToVolts * StepSize / Abs(StepSize) - FromVolts * StepSize / Abs(StepSize)) _
                        / StepSize * StepSize / Abs(StepSize), _
                    0)
    
    'Set the Start row = Currently calibration system's current row
    If InAFMode = True And _
       ActiveAFCoilSystem = AxialAFCoilSystem _
    Then

        StartRow = AxialCurrentRow
            
    ElseIf InAFMode = True And _
           ActiveAFCoilSystem = TransverseAFCoilSystem _
    Then
        
            StartRow = TransCurrentRow
            
    ElseIf InAFMode = False And _
           ActiveAFCoilSystem = AxialCoilSystem _
    Then
            StartRow = IRMAxialCurrentRow
            
    ElseIf InAFMode = False And _
           ActiveAFCoilSystem = TransverseAFCoilSystem _
    Then
        
            StartRow = IRMTransCurrentRow
            
    End If
                
    If StartRow > 1 Then
    
        If InAFMode = True And _
           ActiveAFCoilSystem = AxialAFCoilSystem _
        Then
        
            With Me.gridAFAxialCalibration
            
                .row = StartRow - 1
                .Col = 1
                
                If Abs(val(.text) - FromVolts) < 0.0001 Then
                
                    'We're repeating a step at the same voltage
                    'as the last voltage of a previously added step set
                    'Quick fix - reduce StartRow by 1
                    StartRow = StartRow - 1
                    
                End If
                
            End With
    
        ElseIf InAFMode = True And _
               ActiveAFCoilSystem = TransverseAFCoilSystem _
        Then
        
            With Me.gridAFTransverseCalibration
            
                .row = StartRow - 1
                .Col = 1
                
                If Abs(val(.text) - FromVolts) < 0.0001 Then
                
                    'We're repeating a step at the same voltage
                    'as the last voltage of a previously added step set
                    'Quick fix - reduce StartRow by 1
                    StartRow = StartRow - 1
                    
                End If
                
            End With
            
        ElseIf InAFMode = False And _
               ActiveAFCoilSystem = AxialAFCoilSystem _
        Then
        
            With Me.gridIRMAxial
            
                .row = StartRow - 1
                .Col = 1
                
                If Abs(val(.text) - FromVolts) < 0.0001 Then
                
                    'We're repeating a step at the same voltage
                    'as the last voltage of a previously added step set
                    'Quick fix - reduce StartRow by 1
                    StartRow = StartRow - 1
                    
                End If
                
            End With
            
        ElseIf InAFMode = False And _
               ActiveAFCoilSystem = TransverseAFCoilSystem _
        Then
        
            With Me.gridIRMTrans
            
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
        
    End If
        
    Volts = FromVolts
    i = StartRow
        
    Do While Volts <= ToVolts And Volts >= FromVolts
    
        'If the coil selected is the axial coil, add
        'a new row to the axial grid
        If InAFMode = True And _
           ActiveAFCoilSystem = AxialAFCoilSystem _
        Then
            
            With gridAFAxialCalibration
    
                If i >= .Rows Then
                
                    .Rows = i + 1
        
                End If
        
                .row = i
                .Col = 1
                
                If AFSystem = "ADWIN" Then
                    
                    .text = Format(Volts, "#0.0###")
                
                ElseIf AFSystem = "2G" Then
                
                    .text = Format(Volts, "0")
                
                End If
                
                .RowExpanded = True
                If .RowHeight(i) = 0 Then .RowHeight(i) = 228
                
                .Col = 0
                .text = Trim(Str(i))
                If .ColWidth(0) < Me.TextWidth(Trim(Str(i))) * 2 Then
                
                    .ColWidth(0) = Me.TextWidth(Trim(Str(i))) * 2
                    
                End If
                    
            End With
              
        ElseIf InAFMode = True And _
               ActiveAFCoilSystem = TransverseAFCoilSystem _
        Then
        
            With gridAFTransverseCalibration
    
                If i >= .Rows Then
                
                    .Rows = i + 1
        
                End If
        
                .row = i
                .Col = 1
                
                If AFSystem = "ADWIN" Then
                    
                    .text = Format(Volts, "#0.0###")
                
                ElseIf AFSystem = "2G" Then
                
                    .text = Format(Volts, "0")
                
                End If
    
                .RowExpanded = True
                If .RowHeight(i) = 0 Then .RowHeight(i) = 228
                
                .Col = 0
                .text = Trim(Str(i))
                If .ColWidth(0) < Me.TextWidth(Trim(Str(i))) * 2 Then
                
                    .ColWidth(0) = Me.TextWidth(Trim(Str(i))) * 2
                    
                End If
              
            End With
            
        ElseIf InAFMode = False And _
               ActiveAFCoilSystem = AxialAFCoilSystem _
        Then
            
            With gridIRMAxial
    
                If i >= .Rows Then
                
                    .Rows = i + 1
        
                End If
        
                .row = i
                .Col = 1
                .text = Format(Volts, "#0.0###")
                
                .RowExpanded = True
                If .RowHeight(i) = 0 Then .RowHeight(i) = 228
                
                .Col = 0
                .text = Trim(Str(i))
                If .ColWidth(0) < Me.TextWidth(Trim(Str(i))) * 2 Then
                
                    .ColWidth(0) = Me.TextWidth(Trim(Str(i))) * 2
                    
                End If
              
            End With
            
        ElseIf InAFMode = False And _
               ActiveAFCoilSystem = TransverseAFCoilSystem _
        Then
        
            With gridIRMTrans
    
                If i >= .Rows Then
                
                    .Rows = i + 1
        
                End If
        
                .row = i
                .Col = 1
                .text = Format(Volts, "#0.0###")
                
                .RowExpanded = True
                If .RowHeight(i) = 0 Then .RowHeight(i) = 228
                
                .Col = 0
                .text = Trim(Str(i))
                If .ColWidth(0) < Me.TextWidth(Trim(Str(i))) * 2 Then
                
                    .ColWidth(0) = Me.TextWidth(Trim(Str(i))) * 2
                    
                End If
              
            End With
                
        End If
        
        i = i + 1
        
        If isLog Then
    
            Volts = Volts * (StepSize) ^ (i - StartRow)
        
        Else
        
            Volts = FromVolts + (i - StartRow) * StepSize
    
        End If
        
        DoEvents
        
    Loop
    
    If InAFMode = True And _
       ActiveAFCoilSystem = AxialAFCoilSystem _
    Then

        With gridAFAxialCalibration

            If i >= .Rows Then

                .Rows = i + 1

            End If

            .row = i
            .Col = 1
            
            If AFSystem = "ADWIN" Then
                
                .text = Format(Volts, "#0.0###")
            
            ElseIf AFSystem = "2G" Then
            
                .text = Format(Volts, "0")
            
            End If
                
            .RowExpanded = True
            If .RowHeight(i) = 0 Then .RowHeight(i) = 228
            
            .Col = 0
            .text = Trim(Str(i))
            If .ColWidth(0) < Me.TextWidth(Trim(Str(i))) * 2 Then
                
                .ColWidth(0) = Me.TextWidth(Trim(Str(i))) * 2
                
            End If

        End With

    ElseIf InAFMode = True And _
           ActiveAFCoilSystem = TransverseAFCoilSystem _
    Then


        With gridAFTransverseCalibration

            If i >= .Rows Then

                .Rows = i + 1

            End If

            .row = i
            .Col = 1
            
            If AFSystem = "ADWIN" Then
                
                .text = Format(Volts, "#0.0###")
            
            ElseIf AFSystem = "2G" Then
            
                .text = Format(Volts, "0")
            
            End If
            
            .RowExpanded = True
            If .RowHeight(i) = 0 Then .RowHeight(i) = 228
            
            .Col = 0
            .text = Trim(Str(i))
            If .ColWidth(0) < Me.TextWidth(Trim(Str(i))) * 2 Then
                
                .ColWidth(0) = Me.TextWidth(Trim(Str(i))) * 2
                
            End If

        End With

    ElseIf InAFMode = False And _
           ActiveAFCoilSystem = AxialAFCoilSystem _
    Then
        With gridIRMAxial

            If i >= .Rows Then

                .Rows = i + 1

            End If

            .row = i
            .Col = 1
            .text = Format(Volts, "#0.0###")
            
            .RowExpanded = True
            If .RowHeight(i) = 0 Then .RowHeight(i) = 228
            
            .Col = 0
            .text = Trim(Str(i))
            If .ColWidth(0) < Me.TextWidth(Trim(Str(i))) * 2 Then
                
                .ColWidth(0) = Me.TextWidth(Trim(Str(i))) * 2
                
            End If

        End With

    ElseIf InAFMode = False And _
           ActiveAFCoilSystem = TransverseAFCoilSystem _
    Then

        With gridIRMTrans

            If i >= .Rows Then

                .Rows = i + 1

            End If

            .row = i
            .Col = 1
            .text = Format(Volts, "#0.0###")
            
            .RowExpanded = True
            If .RowHeight(i) = 0 Then .RowHeight(i) = 228
            
            .Col = 0
            .text = Trim(Str(i))
            If .ColWidth(0) < Me.TextWidth(Trim(Str(i))) * 2 Then
                
                .ColWidth(0) = Me.TextWidth(Trim(Str(i))) * 2
                
            End If

        End With

    End If

    i = i + 1
    
    If InAFMode = True And _
       ActiveAFCoilSystem = AxialAFCoilSystem _
    Then
        
        If i > 1 Then
            Me.gridAFAxialCalibration.TopRow = i - 1
        End If
        AxialCurrentRow = i
        
    ElseIf InAFMode = True And _
           ActiveAFCoilSystem = TransverseAFCoilSystem _
    Then
    
        If i > 1 Then
            Me.gridAFTransverseCalibration.TopRow = i - 1
        End If
        TransCurrentRow = i
    
    ElseIf InAFMode = False And _
           ActiveAFCoilSystem = AxialAFCoilSystem _
    Then
        
        If i > 1 Then
            Me.gridIRMAxial.TopRow = i - 1
        End If
        IRMAxialCurrentRow = i
        
    ElseIf InAFMode = False And _
           ActiveAFCoilSystem = TransverseAFCoilSystem _
    Then
    
        If i > 1 Then
            Me.gridIRMTrans.TopRow = i - 1
        End If
        IRMTransCurrentRow = i
    
    End If
    
End Sub

Private Sub cmdClear_Click()

    If InAFMode = True And _
       ActiveAFCoilSystem = AxialAFCoilSystem _
    Then

        With gridAFAxialCalibration

            .Clear
            .ClearStructure

            .Cols = 4
            .Rows = 2
            .FixedCols = 1
            .FixedRows = 1
            
            .Col = 1
            .row = 0
            
            If AFSystem = "ADWIN" Then
            
                .text = "Target Volts"
                
            Else
            
                .text = "2G Counts"
                
            End If
            
            If .ColWidth(1) < Me.TextWidth(.text) * 1.2 Then .ColWidth(1) = Me.TextWidth(.text) * 1.2
            
            .Col = 2
            .row = 0
            .text = "Field (" & modConfig.AFUnits & ")"
            If .ColWidth(2) < Me.TextWidth(.text) * 1.2 Then .ColWidth(2) = Me.TextWidth(.text) * 1.2
            
            .Col = 3
            .row = 0
            .text = "StDev (" & modConfig.AFUnits & ")"
            If .ColWidth(3) < Me.TextWidth(.text) * 1.2 Then .ColWidth(3) = Me.TextWidth(.text) * 1.2
            
            AxialCurrentRow = 1
            
        End With
    
    ElseIf InAFMode = True And _
           ActiveAFCoilSystem = TransverseAFCoilSystem _
    Then
    
        With gridAFTransverseCalibration

            .Clear
            .ClearStructure

            .Cols = 4
            .Rows = 2
            .FixedCols = 1
            .FixedRows = 1
            
            .Col = 1
            .row = 0
            
            If AFSystem = "ADWIN" Then
            
                .text = "Target Volts"
                
            Else
            
                .text = "2G Counts"
                
            End If
            
            If .ColWidth(1) < Me.TextWidth(.text) * 1.2 Then .ColWidth(1) = Me.TextWidth(.text) * 1.2
            
            .Col = 2
            .row = 0
            .text = "Field (" & modConfig.AFUnits & ")"
            If .ColWidth(2) < Me.TextWidth(.text) * 1.2 Then .ColWidth(2) = Me.TextWidth(.text) * 1.2
            
            .Col = 3
            .row = 0
            .text = "StDev (" & modConfig.AFUnits & ")"
            If .ColWidth(3) < Me.TextWidth(.text) * 1.2 Then .ColWidth(3) = Me.TextWidth(.text) * 1.2
                        
            TransCurrentRow = 1
            
        End With
        
    ElseIf InAFMode = False And _
           ActiveAFCoilSystem = AxialAFCoilSystem _
    Then
    
        With gridIRMAxial

            .Clear
            .ClearStructure

            .Cols = 4
            .Rows = 2
            .FixedCols = 1
            .FixedRows = 1
            
            .Col = 1
            .row = 0
            .text = "Volts"
            If .ColWidth(1) < Me.TextWidth(.text) * 1.2 Then .ColWidth(1) = Me.TextWidth(.text) * 1.2
                
            .Col = 2
            .row = 0
            .text = "Field (" & modConfig.AFUnits & ")"
            If .ColWidth(2) < Me.TextWidth(.text) * 1.2 Then .ColWidth(2) = Me.TextWidth(.text) * 1.2
            
            .Col = 3
            .row = 0
            .text = "StDev (" & modConfig.AFUnits & ")"
            If .ColWidth(3) < Me.TextWidth(.text) * 1.2 Then .ColWidth(3) = Me.TextWidth(.text) * 1.2
            
            IRMAxialCurrentRow = 1
            
        End With
        
    ElseIf InAFMode = False And _
           ActiveAFCoilSystem = TransverseAFCoilSystem _
    Then
    
        With gridIRMTrans

            .Clear
            .ClearStructure

            .Cols = 4
            .Rows = 2
            .FixedCols = 1
            .FixedRows = 1
            
            .Col = 1
            .row = 0
            .text = "Volts"
            If .ColWidth(1) < Me.TextWidth(.text) * 1.2 Then .ColWidth(1) = Me.TextWidth(.text) * 1.2
                
            .Col = 2
            .row = 0
            .text = "Field (" & modConfig.AFUnits & ")"
            If .ColWidth(2) < Me.TextWidth(.text) * 1.2 Then .ColWidth(2) = Me.TextWidth(.text) * 1.2
            
            .Col = 3
            .row = 0
            .text = "StDev (" & modConfig.AFUnits & ")"
            If .ColWidth(3) < Me.TextWidth(.text) * 1.2 Then .ColWidth(3) = Me.TextWidth(.text) * 1.2
            
            IRMTransCurrentRow = 1
            
        End With
        
    End If
    
    'Restore # of columns for replicates
    txtNumReplicateRamps_Change
    
End Sub

Private Sub cmdClose_Click()

    Me.Hide
    frmADWIN_AF.Show

End Sub

Private Sub cmdDelete_Click()

    Dim RowStart As Long
    Dim i As Long
    Dim j As Long
    Dim RowEnd As Long

    If InAFMode = True And _
       ActiveAFCoilSystem = AxialAFCoilSystem _
    Then
    
        With Me.gridAFAxialCalibration
        
            'Find bounds on the rows to eliminate
            If .row > .RowSel Then
            
                RowStart = .RowSel
                RowEnd = .row
                
            Else
            
                RowStart = .row
                RowEnd = .RowSel
            
            End If
            
            'Eliminate the rows selected
            For i = RowEnd To RowStart Step -1
            
                If .Rows = 2 Then
                
                    For j = 0 To .Cols - 1
                        .row = 1
                        .Col = j
                        .text = ""
                    Next j
                    AxialCurrentRow = 1
                    
                Else
                    .RemoveItem i
                    AxialCurrentRow = AxialCurrentRow - 1
                End If
                
            Next i
            
            'ReDo the numbering in Col 0
            For i = 1 To .Rows - 1
            
                .Col = 0
                .row = i
                .text = Trim(Str(i))
                If .ColWidth(0) < Me.TextWidth(.text) * 2 Then
                
                    .ColWidth(0) = Me.TextWidth(.text) * 2
                    
                End If

            Next i
            
        End With
        
    ElseIf InAFMode = True And _
           ActiveAFCoilSystem = TransverseAFCoilSystem _
    Then
    
        With Me.gridAFTransverseCalibration
        
            'Find bounds on the rows to eliminate
            If .row > .RowSel Then
            
                RowStart = .RowSel
                RowEnd = .row
                
            Else
            
                RowStart = .row
                RowEnd = .RowSel
            
            End If
            
            'Eliminate the rows selected
            For i = RowEnd To RowStart Step -1
            
                If .Rows = 2 Then
                
                    For j = 0 To .Cols - 1
                        .row = 1
                        .Col = j
                        .text = ""
                    Next j
                    TransCurrentRow = 1
                    
                Else
                    .RemoveItem i
                    TransCurrentRow = TransCurrentRow - 1
                End If
                
            Next i
            
            'ReDo the numbering in Col 0
            For i = 1 To .Rows - 1
            
                .Col = 0
                .row = i
                .text = Trim(Str(i))
                If .ColWidth(0) < Me.TextWidth(.text) * 2 Then
                
                    .ColWidth(0) = Me.TextWidth(.text) * 2
                    
                End If

            Next i
            
        End With
        
    ElseIf InAFMode = False And _
           ActiveAFCoilSystem = AxialAFCoilSystem _
    Then
        
        With Me.gridIRMAxial
        
            'Find bounds on the rows to eliminate
            If .row > .RowSel Then
            
                RowStart = .RowSel
                RowEnd = .row
                
            Else
            
                RowStart = .row
                RowEnd = .RowSel
            
            End If
            
            'Eliminate the rows selected
            For i = RowEnd To RowStart Step -1
            
                If .Rows = 2 Then
                
                    For j = 0 To .Cols - 1
                        .row = 1
                        .Col = j
                        .text = ""
                    Next j
                    IRMAxialCurrentRow = 1
                    
                Else
                    .RemoveItem i
                    IRMAxialCurrentRow = IRMAxialCurrentRow - 1
                End If
                
            Next i
            
            'ReDo the numbering in Col 0
            For i = 1 To .Rows - 1
            
                .Col = 0
                .row = i
                .text = Trim(Str(i))
                If .ColWidth(0) < Me.TextWidth(.text) * 2 Then
                
                    .ColWidth(0) = Me.TextWidth(.text) * 2
                    
                End If

            Next i
            
        End With
            
    ElseIf InAFMode = False And _
           ActiveAFCoilSystem = TransverseAFCoilSystem _
    Then
        
        With Me.gridIRMTrans
        
            'Find bounds on the rows to eliminate
            If .row > .RowSel Then
            
                RowStart = .RowSel
                RowEnd = .row
                
            Else
            
                RowStart = .row
                RowEnd = .RowSel
            
            End If
            
            'Eliminate the rows selected
            For i = RowEnd To RowStart Step -1
            
                If .Rows = 2 Then
                
                    For j = 0 To .Cols - 1
                        .row = 1
                        .Col = j
                        .text = ""
                    Next j
                    IRMTransCurrentRow = 1
                    
                Else
                    .RemoveItem i
                    IRMTransCurrentRow = IRMTransCurrentRow - 1
                End If
                
            Next i
            
            'ReDo the numbering in Col 0
            For i = 1 To .Rows - 1
            
                .Col = 0
                .row = i
                .text = Trim(Str(i))
                If .ColWidth(0) < Me.TextWidth(.text) * 2 Then
                
                    .ColWidth(0) = Me.TextWidth(.text) * 2
                    
                End If
                
            Next i
            
        End With
        
    End If

End Sub

Private Sub cmdLoadFromCSVFile_Click()

    Dim wasLoadSuccessful As Boolean
    
    If InAFMode = True And _
       ActiveAFCoilSystem = AxialAFCoilSystem _
    Then
    
        wasLoadSuccessful = frmFileSave.LoadAFCalibrationTable(Me.gridAFAxialCalibration, _
                                                               modConfig.AFUnits)
                                                               
        'Check to see if the load-table operation was successful
        'If not, user needs to change the settings in the AF File
        'Save Settings window
        If wasLoadSuccessful = False Then
        
            Load frmFileSave
            frmFileSave.Show
            
        Else
        
            'Update the number of replicates on the form control
            Me.txtNumReplicateRamps = Trim(Str(Me.gridAFAxialCalibration.Cols - 4)) \ 2
            
        End If
                                
    ElseIf InAFMode = True And _
           ActiveAFCoilSystem = TransverseAFCoilSystem _
    Then
    
        wasLoadSuccessful = frmFileSave.LoadAFCalibrationTable(Me.gridAFTransverseCalibration, _
                                                               modConfig.AFUnits)
                                
                'Check to see if the load-table operation was successful
        'If not, user needs to change the settings in the AF File
        'Save Settings window
        If wasLoadSuccessful = False Then
        
            Load frmFileSave
            frmFileSave.Show
            
        Else
        
            'Update the number of replicates on the form control
            Me.txtNumReplicateRamps = Trim(Str(Me.gridAFTransverseCalibration.Cols - 4)) \ 2
            
        End If
                                                    
    ElseIf InAFMode = False And _
           ActiveAFCoilSystem = AxialAFCoilSystem _
    Then
    
        wasLoadSuccessful = frmFileSave.LoadAFCalibrationTable(Me.gridIRMAxial, _
                                                               modConfig.AFUnits)
                                                               
        'Check to see if the load-table operation was successful
        'If not, user needs to change the settings in the AF File
        'Save Settings window
        If wasLoadSuccessful = False Then
        
            Load frmFileSave
            frmFileSave.Show
            
        Else
        
            'Update the number of replicates on the form control
            Me.txtNumReplicateRamps = Trim(Str(Me.gridIRMAxial.Cols - 4)) \ 2
            
        End If
                                          
    ElseIf InAFMode = False And _
           ActiveAFCoilSystem = TransverseAFCoilSystem _
    Then
    
        wasLoadSuccessful = frmFileSave.LoadAFCalibrationTable(Me.gridIRMTrans, _
                                                               modConfig.AFUnits)
                                                               
        'Check to see if the load-table operation was successful
        'If not, user needs to change the settings in the AF File
        'Save Settings window
        If wasLoadSuccessful = False Then
        
            Load frmFileSave
            frmFileSave.Show
            
        Else
        
            'Update the number of replicates on the form control
            Me.txtNumReplicateRamps = Trim(Str(Me.gridIRMTrans.Cols - 4)) \ 2
            
        End If
                                                    
    End If

End Sub

Private Sub cmdSaveToCSVFile_Click()

    Dim wasSaveSuccessful As Boolean
    Dim CurTime
    
    CurTime = Now
    
                 
    ElseIf InAFMode = True And _
           ActiveAFCoilSystem = AxialAFCoilSystem _
    Then
    
        wasSaveSuccessful = frmFileSave.SaveAFCalibrationTable( _
                                            Me.gridAFAxialCalibration, _
                                            CurTime, _
                                            modConfig.AFUnits)
                                           
    ElseIf InAFMode = True And _
           ActiveAFCoilSystem = TransverseAFCoilSystem _
    Then
    
        wasSaveSuccessful = frmFileSave.SaveAFCalibrationTable( _
                                            Me.gridAFTransverseCalibration, _
                                            CurTime, _
                                            modConfig.AFUnits)
                                           
    ElseIf InAFMode = False And _
           ActiveAFCoilSystem = AxialAFCoilSystem _
    Then
    
        wasSaveSuccessful = frmFileSave.SaveIRMCalibrationTable( _
                                            Me.gridIRMAxial, _
                                            CurTime, _
                                            modConfig.AFUnits)
                                           
    ElseIf InAFMode = False And _
           ActiveAFCoilSystem = TransverseAFCoilSystem _
    Then
    
        wasSaveSuccessful = frmFileSave.SaveIRMCalibrationTable( _
                                            Me.gridIRMTrans, _
                                            CurTime, _
                                            modConfig.AFUnits)
                                           
    End If
    
    'Check to see if the save was successful
    If wasSaveSuccessful = False Then
        
        'user need's to change the settings in the AF file save settings window
        'show that window
        Load frmFileSave
        frmFileSave.Show
                                           
    End If
                                           
End Sub



Private Sub cmdClearData_Click()

    Dim NumRows As Long
    Dim NumCols As Long
    Dim i, j As Long
    
    'Need to delete the text in cols 2 through N - 1 in
    'rows 1 through N - 1
    
    ElseIf InAFMode = True And _
           ActiveAFCoilSystem = AxialAFCoilSystem _
    Then
    
        NumRows = Me.gridAFAxialCalibration.Rows
        NumCols = Me.gridAFAxialCalibration.Cols
        
        With Me.gridAFAxialCalibration
        
            For i = 1 To NumRows - 1
            
                .row = i
            
                For j = 2 To NumCols - 1
                
                    .Col = j
                    
                    .text = ""
                    
                Next j
                
            Next i
            
        End With
        
    ElseIf InAFMode = True And _
           ActiveAFCoilSystem = TransverseAFCoilSystem _
    Then
    
        NumRows = Me.gridAFTransverseCalibration.Rows
        NumCols = Me.gridAFTransverseCalibration.Cols
    
        With Me.gridAFTransverseCalibration
        
            For i = 1 To NumRows - 1
            
                .row = i
            
                For j = 2 To NumCols - 1
                
                    .Col = j
                    
                    .text = ""
                    
                Next j
                
            Next i
            
        End With
        
    ElseIf InAFMode = False And _
           ActiveAFCoilSystem = AxialAFCoilSystem _
    Then
    
        NumRows = Me.gridIRMAxial.Rows
        NumCols = Me.gridIRMAxial.Cols
        
        With Me.gridIRMAxial
        
            For i = 1 To NumRows - 1
            
                .row = i
            
                For j = 2 To NumCols - 1
                
                    .Col = j
                    
                    .text = ""
                    
                Next j
                
            Next i
            
        End With
                
    ElseIf InAFMode = False And _
           ActiveAFCoilSystem = TransverseAFCoilSystem _
    Then
    
        NumRows = Me.gridIRMTrans.Rows
        NumCols = Me.gridIRMTrans.Cols
        
        With Me.gridIRMTrans
        
            For i = 1 To NumRows - 1
            
                .row = i
            
                For j = 2 To NumCols - 1
                
                    .Col = j
                    
                    .text = ""
                    
                Next j
                
            Next i
            
        End With
        
    End If
        
End Sub

Public Sub Form_Load()

    Dim i As Long
    Dim NumReplicates As Long
    
    'Set the form window size
    Me.Height = 6975
    Me.Width = 8085
    
    'Set Relays for correct coil system
    If ActiveAFCoilSystem = AxialAFCoilSystem Then
    
        If AFSystem = "2G" Then
        
            frmAF_2G.ConfigureCoil AfAxialCoord
    
        ElseIf AFSystem = "ADWIN" Then
        
            SetADWINCoils 0
            
        End If
    
    ElseIf ActiveAFCoilSystem = TransverseAFCoilSystem Then
    
        If AFSystem = "2G" Then
        
            frmAF_2G.ConfigureCoil AfTransCoord
    
        ElseIf AFSystem = "ADWIN" Then
        
            SetADWINCoils 1
            
        End If
        
    Else
    
        If AFSystem = "ADWIN" Then
        
            SetADWINCoils -10
            
        End If
        
    End If
            
    'Figure out which is the selected Coil (Axial, Transverse, IRM-LF, IRM-HF)
    If InAFMode = True Then
    
        'Show the AF coil selector frame
        Me.frameAFCoilSelection.Visible = True
        
        'Hide the IRM Coil selector frame
        Me.frameIRMCoil.Visible = False
    
        'Change the captions on the Max & Min frames and on the form labels
        Me.frameAxialMaxAndMin.Caption = "Axial Max / Min Fields"
        Me.frameTransMaxAndMin.Caption = "Trans. Max / Min Fields"
                  
        'Change Step labels based on AF system being used
        If modConfig.AFSystem = "ADWIN" Then
        
            Me.lblAFVoltStep.Caption = "AF Volt Step:"
            Me.lblNumReplicates.Caption = "# of Replicate AF Ramps per Voltage Step:"
            
        ElseIf AFSystem = "2G" Then
        
            Me.lblAFVoltStep.Caption = "2G Counts Step:"
            Me.lblNumReplicates.Caption = "# of Replicate AF Ramps per 2G value:"
            
        End If
            
        'Set Max and Min Axial & Transverse
        'If calibration for each coil is done, update the value of Min & Max field
        If modConfig.AFAxialCalDone = True Then
        
            Me.txtAFAxialMaxMonitorVoltage = Trim(Str(modConfig.AfAxialMax))
            Me.txtAFAxialMinMonitorVoltage = Trim(Str(modConfig.AfAxialMin))
            
        Else
        
            Me.txtAFAxialMaxMonitorVoltage = ""
            Me.txtAFAxialMinMonitorVoltage = ""
            
        End If
    
        If modConfig.AFTransCalDone = True Then
        
            Me.txtAFTransMaxMonitorVoltage = Trim(Str(modConfig.AfTransMax))
            Me.txtAFTransMinMonitorVoltage = Trim(Str(modConfig.AfTransMin))
        
        Else
        
            Me.txtAFTransMaxMonitorVoltage = ""
            Me.txtAFTransMinMonitorVoltage = ""
            
        End If
        
        
    ElseIf InAFMode = False And _
           (EnableAxialIRM = True Or _
            EnableTransIRM = True Or _
            EnableIRMBackfield = True) _
    Then
        
        'Change the Captions on the optCoil radio buttons
        
        
        'Change the captions on the Max & Min frames and on the form labels
        Me.frameAxialMaxAndMin.Caption = "IRM-LF Max / Min Fields"
        Me.frameTransMaxAndMin.Caption = "IRM-HF Max / Min Fields"
        Me.lblAFVoltStep.Caption = "IRM Volt Step:"
        Me.lblNumReplicates.Caption = "# of Replicate IRM Pulses per Voltage Step:"
        
        'Set Max and Min Axial & Transverse
        'if IRM calibrations are done, display the min & max IRM fields
        If modConfig.IRMLFCalDone = True Then
        
            Me.txtAFAxialMaxMonitorVoltage = Trim(Str(modConfig.PulseLFMax))
            Me.txtAFAxialMinMonitorVoltage = Trim(Str(modConfig.PulseLFMin))
            
        Else
        
            Me.txtAFAxialMaxMonitorVoltage = ""
            Me.txtAFAxialMinMonitorVoltage = ""
            
        End If
        
        If modConfig.IRMHFCalDone = True Then
        
            Me.txtAFTransMaxMonitorVoltage = Trim(Str(modConfig.PulseHFMax))
            Me.txtAFTransMinMonitorVoltage = Trim(Str(modConfig.PulseHFMin))
            
        Else
        
            Me.txtAFTransMaxMonitorVoltage = ""
            Me.txtAFTransMinMonitorVoltage = ""
            
        End If
        
        
        If CoilSelected = IRMLF And _
           (EnableAxialIRM = True Or _
            EnableIRMBackfield = True) _
        Then
        
            optIRMCoil(0).Value = True
            optIRMCoil_Click (0)
            CoilString = "IRM Low-Field"
            
        ElseIf EnableTransIRM = True Then
        
            optIRMCoil(1).Value = True
            optIRMCoil_Click (1)
            CoilString = "IRM High-Field"
            
        End If
        
    Else
        
        'If no IRM modules are activated, then do not load / show this form
        Me.Hide
        
        Exit Sub
        
    End If
    
    'Clear out the values in the voltage step add line
    Me.txtStepSize.text = ""
    Me.txtFromVolts.text = ""
    Me.txtToVolts.text = ""
    
    'Set # of replicates to zero
    Me.txtNumReplicateRamps.text = "0"
    NumReplicates = CLng(val(Me.txtNumReplicateRamps))
    txtNumReplicateRamps_Change
    
    'Swipe the grids clear
    Me.gridAFAxialCalibration.Clear
    Me.gridAFAxialCalibration.ClearStructure
    Me.gridAFTransverseCalibration.Clear
    Me.gridAFTransverseCalibration.ClearStructure
    
    'Set fixed rows and columns
    Me.gridAFAxialCalibration.FixedCols = 1
    Me.gridAFAxialCalibration.FixedRows = 1
    Me.gridAFTransverseCalibration.FixedCols = 1
    Me.gridAFTransverseCalibration.FixedRows = 1
    
    'Allow Word Wrap
    Me.gridAFAxialCalibration.WordWrap = True
    Me.gridAFTransverseCalibration.WordWrap = True
    Me.gridIRMAxial.WordWrap = True
    Me.gridIRMTrans.WordWrap = True
    
    'Set the number of Rows and Columns in the Axial and Transverse grids
    Me.gridAFAxialCalibration.Rows = 2
    Me.gridAFAxialCalibration.Cols = 4
    Me.gridAFTransverseCalibration.Rows = 2
    Me.gridAFTransverseCalibration.Cols = 4
    Me.gridIRMAxial.Rows = 2
    Me.gridIRMAxial.Cols = 4
    Me.gridIRMTrans.Rows = 2
    Me.gridIRMTrans.Cols = 4
    AxialCurrentRow = 1
    TransCurrentRow = 1
    IRMAxialCurrentRow = 1
    IRMTransCurrentRow = 1
    
    'Write in the Column Headers
    
    With Me.gridAFAxialCalibration
        
        .row = 0
        .Col = 1
        
        If AFSystem = "ADWIN" Then
            
            .text = "Target Voltage"
            
        ElseIf AFSystem = "2G" Then
        
            .text = "2G Counts"
            
        End If
        
        .ColWidth(1) = Me.TextWidth(.text) * 1.2
        
        .RowSizingMode = flexRowSizeIndividual
        .RowHeight(0) = 456
        
        .Col = 2
        .text = "Field (" & modConfig.AFUnits & ")"
        .ColWidth(2) = Me.TextWidth(.text) * 1.2
        
        .Col = 3
        .text = "StDev (" & modConfig.AFUnits & ")"
        .ColWidth(3) = Me.TextWidth(.text) * 1.2
        
    End With
        
    With Me.gridAFTransverseCalibration
    
        .row = 0
        .Col = 1
        
        If AFSystem = "ADWIN" Then
            
            .text = "Target Voltage"
            
        ElseIf AFSystem = "2G" Then
        
            .text = "2G Counts"
            
        End If
        
        .ColWidth(1) = Me.TextWidth(.text) * 1.2
        
        .RowSizingMode = flexRowSizeIndividual
        .RowHeight(0) = 456
        
        .Col = 2
        .text = "Field (" & modConfig.AFUnits & ")"
        .ColWidth(2) = Me.TextWidth(.text) * 1.2
        
        .Col = 3
        .text = "StDev (" & modConfig.AFUnits & ")"
        .ColWidth(3) = Me.TextWidth(.text) * 1.2
                
    End With
        
    With gridIRMAxial
        
        .row = 0
        .Col = 1
        .text = "Target Voltage"
        .RowSizingMode = flexRowSizeIndividual
        .RowHeight(0) = 456
        .ColWidth(1) = Me.TextWidth(.text) * 1.2
    
        .Col = 2
        .text = "Field (" & modConfig.AFUnits & ")"
        .ColWidth(2) = Me.TextWidth(.text) * 1.2
        
        .Col = 3
        .text = "StDev (" & modConfig.AFUnits & ")"
        .ColWidth(3) = Me.TextWidth(.text) * 1.2
            
    End With
    
    With gridIRMTrans
        
        .row = 0
        .Col = 1
        .text = "Target Voltage"
        .RowSizingMode = flexRowSizeIndividual
        .RowHeight(0) = 456
        .ColWidth(1) = Me.TextWidth(.text) * 1.2
    
        .Col = 2
        .text = "Field (" & modConfig.AFUnits & ")"
        .ColWidth(2) = Me.TextWidth(.text) * 1.2
        
        .Col = 3
        .text = "StDev (" & modConfig.AFUnits & ")"
        .ColWidth(3) = Me.TextWidth(.text) * 1.2
            
    End With
    
    'Load the values from frmSettings
    'from the prior AF/IRM calibrations into the
    'Axial and Transverse grids
    
    'Is this form in IRM or AF calibration mode?
    If (CoilSelected = Axial Or _
        CoilSelected = Transverse) _
    Then
    
        'We're in AF mode
        With Me.gridAFAxialCalibration
        
            .Rows = val(modConfig.AFAxialCount) + 1
            .Cols = 4
            
            For i = 1 To .Rows - 1
            
                .row = i
                .Col = 0
                .text = Trim(Str(i))
                If .ColWidth(0) < Me.TextWidth(Trim(Str(i))) * 2 Then
                
                    .ColWidth(0) = Me.TextWidth(Trim(Str(i))) * 2
                
                End If
                
                .Col = 1
                frmSettings.grdCalibAxial.row = i
                frmSettings.grdCalibAxial.Col = 1
                .text = Trim(Str(val(frmSettings.grdCalibAxial.text)))
                If .ColWidth(1) < Me.TextWidth(.text) * 1.2 Then .ColWidth(1) = Me.TextWidth(.text) * 1.2
                
                .Col = 2
                frmSettings.grdCalibAxial.row = i
                frmSettings.grdCalibAxial.Col = 2
                .text = Trim(Str(val(frmSettings.grdCalibAxial.text)))
                If .ColWidth(2) < Me.TextWidth(.text) * 1.2 Then .ColWidth(2) = Me.TextWidth(.text) * 1.2
    
            Next i
            
            AxialCurrentRow = i
            
        End With
    
        With Me.gridAFTransverseCalibration
        
            .Rows = val(modConfig.AFTransCount) + 1
            .Cols = 4
            
            For i = 1 To .Rows - 1
            
                .row = i
                .Col = 0
                .text = Trim(Str(i))
                If .ColWidth(0) < Me.TextWidth(.text) * 2 Then .ColWidth(0) = Me.TextWidth(.text) * 2
                
                .Col = 1
                frmSettings.grdCalibTrans.row = i
                frmSettings.grdCalibTrans.Col = 1
                .text = Trim(Str(val(frmSettings.grdCalibTrans.text)))
                If .ColWidth(1) < Me.TextWidth(.text) * 1.2 Then .ColWidth(1) = Me.TextWidth(.text) * 1.2
                
                .Col = 2
                frmSettings.grdCalibTrans.row = i
                frmSettings.grdCalibTrans.Col = 2
                .text = Trim(Str(val(frmSettings.grdCalibTrans.text)))
                If .ColWidth(2) < Me.TextWidth(.text) * 1.2 Then .ColWidth(2) = Me.TextWidth(.text) * 1.2
    
            Next i
            
            TransCurrentRow = i
            
        End With
        
    Else
    
        'In IRM Mode
        With Me.gridIRMAxial
        
            .Rows = val(modConfig.PulseLFCount) + 1
            .Cols = 4
            
            For i = 1 To .Rows - 1
            
                .row = i
                .Col = 0
                .text = Trim(Str(i))
                If .ColWidth(0) < Me.TextWidth(.text) * 2 Then .ColWidth(0) = Me.TextWidth(.text) * 2
                
                
                .Col = 1
                frmSettings.gridCalibIRMAxial.row = i
                frmSettings.gridCalibIRMAxial.Col = 1
                .text = Trim(Str(val(frmSettings.gridCalibIRMAxial.text)))
                If .ColWidth(1) < Me.TextWidth(.text) * 1.2 Then .ColWidth(1) = Me.TextWidth(.text) * 1.2
                
                .Col = 2
                frmSettings.gridCalibIRMAxial.row = i
                frmSettings.gridCalibIRMAxial.Col = 2
                .text = Trim(Str(val(frmSettings.gridCalibIRMAxial.text)))
                If .ColWidth(2) < Me.TextWidth(.text) * 1.2 Then .ColWidth(2) = Me.TextWidth(.text) * 1.2
    
            Next i
            
            IRMAxialCurrentRow = i
            
        End With
        
        With Me.gridIRMTrans
        
            .Rows = val(modConfig.PulseHFCount) + 1
            .Cols = 4
            
            For i = 1 To .Rows - 1
            
                .row = i
                .Col = 0
                .text = Trim(Str(i))
                If .ColWidth(0) < Me.TextWidth(.text) * 2 Then .ColWidth(0) = Me.TextWidth(.text) * 2
                
                .Col = 1
                frmSettings.gridCalibIRMTrans.row = i
                frmSettings.gridCalibIRMTrans.Col = 1
                .text = Trim(Str(val(frmSettings.gridCalibIRMTrans.text)))
                If .ColWidth(1) < Me.TextWidth(.text) * 1.2 Then .ColWidth(1) = Me.TextWidth(.text) * 1.2
                
                .Col = 2
                frmSettings.gridCalibIRMTrans.row = i
                frmSettings.gridCalibIRMTrans.Col = 2
                .text = Trim(Str(val(frmSettings.gridCalibIRMTrans.text)))
                If .ColWidth(2) < Me.TextWidth(.text) * 1.2 Then .ColWidth(2) = Me.TextWidth(.text) * 1.2
    
            Next i
            
            IRMTransCurrentRow = i
            
        End With
        
    End If
        
    'Refresh the form display
    Me.refresh
    
End Sub

Private Sub optAFCoil_Click(Index As Integer)

    If Index = 0 Then
    'Axial Coil selected
        
        CoilSelected = Axial
        
        'Show the Axial Grid spreadsheet
        'and hide the Transverse Grid
        Me.gridAFAxialCalibration.Visible = True
        Me.gridAFTransverseCalibration.Visible = False
        Me.gridIRMAxial.Visible = False
        Me.gridIRMTrans.Visible = False
        
        With Me.gridAFAxialCalibration
        
            If .Rows > 2 Then
            
                .TopRow = .Rows - 1
                
            End If
        
        End With
        
        'Now actually change the coil relays, depending
        'on the AF system in use
        If AFSystem = "ADWIN" Then
            
            SetADWINCoils 0
                        
        ElseIf AFSystem = "2G" Then
        
            frmAF_2G.ConfigureCoil modConfig.AfAxialCoord
            ActiveAFCoilSystem = AxialAFCoilSystem
            
        End If
            
    ElseIf Index = 1 Then
    
        CoilSelected = Transverse
    
        'Show the Transverse Grid spreadsheet
        'and hide the Axial Grid
        Me.gridAFTransverseCalibration.Visible = True
        Me.gridAFAxialCalibration.Visible = False
        Me.gridIRMAxial.Visible = False
        Me.gridIRMTrans.Visible = False
        
        With Me.gridAFTransverseCalibration
        
            If .Rows > 2 Then
            
                .TopRow = .Rows - 1
                
            End If
        
        End With
        
        'Now actually change the coil relays, depending
        'on the AF system in use
        If AFSystem = "ADWIN" Then
            
            SetADWINCoils 1
            
        ElseIf AFSystem = "2G" Then
        
            frmAF_2G.ConfigureCoil modConfig.AfTransCoord
            ActiveAFCoilSystem = TransverseAFCoilSystem
            
        End If
        
    Else
    
        If AFSystem = "ADWIN" Then
        
            'No coil selected
            SetADWINCoils -10
            
        End If
        
    End If
        
End Sub

Private Sub optIRMCoil_Click(Index As Integer)

    'Need to set the relay so the correct IRM coil is selected
    If Index = 0 Then
    
        'Now actually set the relays for the IRM Low-Field config
        'depending on the AF System being used
        frmIRMARM.SetIRMCoilSystem IRMLFCoilSystem
                                                      
        'Set the global variable Coil String
        CoilString = "IRM"
       
        'Change which grids are visible
        Me.gridAFTransverseCalibration.Visible = False
        Me.gridAFAxialCalibration.Visible = False
        Me.gridIRMAxial.Visible = True
        Me.gridIRMTrans.Visible = False
                                                      
    'Check to see if the IRM High Field module is on
    ElseIf Index = 1 Then
    
        'Configure the relays for the IRM Hi-field pulse
        'depending on the AF System being used
                                                  
                                                  
                                                  
        'Set the global variable Coil String
        CoilString = "IRM High-Field"
        
        'set the coil selected
        CoilSelected = IRMHF
        
        'Change which grids are visible
        Me.gridAFTransverseCalibration.Visible = False
        Me.gridAFAxialCalibration.Visible = False
        Me.gridIRMAxial.Visible = False
        Me.gridIRMTrans.Visible = True
        
    End If

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
    
    If CoilSelected = Axial Then
        
        Me.gridAFAxialCalibration.Cols = 4 + NumReplicates * 2
    
    ElseIf CoilSelected = Transverse Then
    
        Me.gridAFTransverseCalibration.Cols = 4 + NumReplicates * 2
        
    ElseIf CoilSelected = IRMLF Then
    
        Me.gridIRMAxial.Cols = 4 + NumReplicates * 2
    
    ElseIf CoilSelected = IRMHF Then
    
        Me.gridIRMTrans.Cols = 4 + NumReplicates * 2
    
    End If
    
    If NumReplicates = 0 Then
        
        Exit Sub
        
    End If
    
    For i = 4 To 4 + (NumReplicates * 2) - 1 Step 2

        If CoilSelected = Axial Then
        
            With Me.gridAFAxialCalibration
            
                .row = 0
                .Col = i
                .text = "Field #" & Trim(Str((i - 3) \ 2 + 1)) & " (" & modConfig.AFUnits & ")"
                If .ColWidth(i) < Me.TextWidth(.text) * 1.2 Then .ColWidth(i) = Me.TextWidth(.text) * 1.2
                
                .row = 0
                .Col = i + 1
                .text = "Max Volt. #" & Trim(Str((i - 3) \ 2 + 1)) & " (V)"
                If .ColWidth(i) < Me.TextWidth(.text) * 1.2 Then .ColWidth(i) = Me.TextWidth(.text) * 1.2
                
            End With

        ElseIf CoilSelected = Transverse Then

            With Me.gridAFTransverseCalibration
            
                .row = 0
                .Col = i
                .text = "Field #" & Trim(Str((i - 3) \ 2 + 1)) & " (" & modConfig.AFUnits & ")"
                If .ColWidth(i) < Me.TextWidth(.text) * 1.2 Then .ColWidth(i) = Me.TextWidth(.text) * 1.2
                
                .row = 0
                .Col = i + 1
                .text = "Max Volt. #" & Trim(Str((i - 3) \ 2 + 1)) & " (V)"
                If .ColWidth(i) < Me.TextWidth(.text) * 1.2 Then .ColWidth(i) = Me.TextWidth(.text) * 1.2
                
            End With
            
        ElseIf CoilSelected = IRMLF Then

            With Me.gridIRMAxial
            
                .row = 0
                .Col = i
                .text = "Field #" & Trim(Str((i - 3) \ 2 + 1)) & " (" & modConfig.AFUnits & ")"
                If .ColWidth(i) < Me.TextWidth(.text) * 1.2 Then .ColWidth(i) = Me.TextWidth(.text) * 1.2
                
                .row = 0
                .Col = i + 1
                .text = "Max Volt. #" & Trim(Str((i - 3) \ 2 + 1)) & " (V)"
                If .ColWidth(i) < Me.TextWidth(.text) * 1.2 Then .ColWidth(i) = Me.TextWidth(.text) * 1.2
                
            End With
            
        ElseIf CoilSelected = IRMHF Then

            With Me.gridIRMTrans
            
                .row = 0
                .Col = i
                .text = "Field #" & Trim(Str((i - 3) \ 2 + 1)) & " (" & modConfig.AFUnits & ")"
                If .ColWidth(i) < Me.TextWidth(.text) * 1.2 Then .ColWidth(i) = Me.TextWidth(.text) * 1.2
                
                .row = 0
                .Col = i + 1
                .text = "Max Volt. #" & Trim(Str((i - 3) \ 2 + 1)) & " (V)"
                If .ColWidth(i) < Me.TextWidth(.text) * 1.2 Then .ColWidth(i) = Me.TextWidth(.text) * 1.2
                
            End With
            
        End If

    Next i

End Sub
