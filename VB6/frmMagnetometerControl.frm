VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmMagnetometerControl 
   Caption         =   "Magnetometer Control"
   ClientHeight    =   3060
   ClientLeft      =   10650
   ClientTop       =   4620
   ClientWidth     =   4185
   Icon            =   "frmMagnetometerControl.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3060
   ScaleWidth      =   4185
   Visible         =   0   'False
   Begin VB.Frame frameControl 
      BorderStyle     =   0  'None
      Caption         =   "Automatic Data Collection"
      Height          =   975
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   3612
      Begin VB.CommandButton cmdChangerEdit 
         Caption         =   "&Modify"
         Enabled         =   0   'False
         Height          =   315
         Left            =   240
         TabIndex        =   1
         Top             =   600
         Width           =   735
      End
      Begin VB.CommandButton cmdChangerOK 
         Caption         =   "Start changer"
         Enabled         =   0   'False
         Height          =   315
         Left            =   1440
         TabIndex        =   2
         Top             =   600
         Width           =   1212
      End
   End
   Begin VB.Frame frameControl 
      BorderStyle     =   0  'None
      Caption         =   "Manual Data Collection"
      Height          =   1935
      Index           =   1
      Left            =   240
      TabIndex        =   3
      Top             =   480
      Width           =   3612
      Begin VB.ComboBox cmbManSample 
         Height          =   315
         Left            =   2160
         TabIndex        =   4
         Top             =   120
         Width           =   1332
      End
      Begin VB.TextBox txtSampleHeight 
         Height          =   285
         Left            =   2160
         TabIndex        =   5
         Top             =   600
         Width           =   1335
      End
      Begin VB.CommandButton cmdManHolder 
         Caption         =   "Measure &Holder"
         Enabled         =   0   'False
         Height          =   315
         Left            =   360
         TabIndex        =   6
         Top             =   1080
         Width           =   1332
      End
      Begin VB.CommandButton cmdManRun 
         Caption         =   "&Measure Sample"
         Enabled         =   0   'False
         Height          =   315
         Left            =   2040
         TabIndex        =   7
         Top             =   1080
         Width           =   1452
      End
      Begin VB.CheckBox chkVacuum 
         Caption         =   "Keep the vacuum on"
         Height          =   315
         Left            =   0
         TabIndex        =   8
         Top             =   1680
         Width           =   1935
      End
      Begin VB.CommandButton cmdOpenSampleFile 
         Caption         =   "Open Sample File"
         Enabled         =   0   'False
         Height          =   315
         Left            =   2040
         TabIndex        =   9
         Top             =   1560
         Width           =   1452
      End
      Begin VB.Label lblManSample 
         Caption         =   "Choose Sample:"
         Height          =   255
         Left            =   360
         TabIndex        =   10
         Top             =   120
         Width           =   1305
      End
      Begin VB.Label Label1 
         Caption         =   "Sample Height (cm):"
         Height          =   255
         Left            =   360
         TabIndex        =   11
         Top             =   600
         Width           =   1455
      End
   End
   Begin VB.CheckBox chkSlow 
      Caption         =   "Slow motions"
      Height          =   315
      Left            =   120
      TabIndex        =   12
      Top             =   2640
      Width           =   1335
   End
   Begin ComctlLib.TabStrip tbsControl 
      Height          =   2415
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   4260
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   2
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Automatic Data Collection"
            Key             =   "tabAutomatic"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Manual Data Collection"
            Key             =   "tabManual"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.ComboBox cmbSusceptibilityScaleFactor 
      Height          =   315
      Left            =   3240
      TabIndex        =   14
      Text            =   "Combo1"
      Top             =   2640
      Width           =   855
   End
   Begin VB.Label Label19 
      Caption         =   "Susceptibility scale:"
      Height          =   375
      Left            =   1680
      TabIndex        =   15
      Top             =   2640
      Width           =   1575
   End
End
Attribute VB_Name = "frmMagnetometerControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private SampleCode As String
Private DataFileDrv As String
Private DataFileDir As String
Private DataFileName As String
Private fileReadyToLoad As Boolean
Private initialized As Boolean

Private Sub chkSlow_Click()
    
    ' Up down speeds all slow (October 2009 L Carporzen)
    If chkSlow = Checked Then
        
        LiftSpeedSlow = val(Config_GetFromINI("SteppingMotor", "LiftSpeedSlow", "4000000", Prog_INIFile))
        LiftSpeedNormal = LiftSpeedSlow
        LiftSpeedFast = LiftSpeedSlow
        
        frmSettings.txtLiftSpeedSlow = LiftSpeedSlow
        frmSettings.txtLiftSpeedNormal = LiftSpeedNormal
        frmSettings.txtLiftSpeedFast = LiftSpeedFast
        frmSettings.txtLiftAcceleration = LiftAcceleration
    
    Else
        
        LiftSpeedSlow = val(Config_GetFromINI("SteppingMotor", "LiftSpeedSlow", "4000000", Prog_INIFile))
        LiftSpeedNormal = val(Config_GetFromINI("SteppingMotor", "LiftSpeedNormal", "25000000", Prog_INIFile))
        LiftSpeedFast = val(Config_GetFromINI("SteppingMotor", "LiftSpeedFast", "50000000", Prog_INIFile))
        LiftAcceleration = val(Config_GetFromINI("SteppingMotor", "LiftAcceleration", "90000", Prog_INIFile))
        
        frmSettings.txtLiftSpeedSlow = LiftSpeedSlow
        frmSettings.txtLiftSpeedNormal = LiftSpeedNormal
        frmSettings.txtLiftSpeedFast = LiftSpeedFast
        frmSettings.txtLiftAcceleration = LiftAcceleration
    
    End If
End Sub

Private Sub chkVacuum_Click()
    If cmdManHolder.Enabled = True Then
      
      If frmVacuum.VacuumConnectOn = True And chkVacuum = Unchecked Then
        
        MsgBox "Will switch off the vacuum."
        
        frmVacuum.ValveConnect False
        frmVacuum.MotorPower False
      
      End If
    
    End If

End Sub

Private Sub cmbManSample_Change()
    
    Dim specName As String, specParent As String
    Dim specimen As Sample
    
    If LenB(cmbManSample.text) = 0 Then Exit Sub
    
    EnableMagnetCmds
    specName = cmbManSample.text
    specParent = SampleIndexRegistry.SampleFileByIndex(cmbManSample.ListIndex + 1)
    
    If SampleIndexRegistry.IsValidSample(specParent, specName) Then
        
        cmdOpenSampleFile.Enabled = True
        
        Set specimen = SampleIndexRegistry(specParent).sampleSet(specName)
        
        If specimen.SampleHeight > 0 Then
            
            txtSampleHeight = Format$(specimen.SampleHeight / UpDownMotor1cm, "0.00")
        
        Else
            
            txtSampleHeight = Format$(SampleHeight / UpDownMotor1cm, "0.00")
        
        End If
        
        Set specimen = Nothing
    
    Else
        
        cmdOpenSampleFile.Enabled = False
    
    End If

End Sub

Private Sub cmbManSample_click()
    
    cmbManSample_Change

End Sub

Private Sub cmbSusceptibilityScaleFactor_Click()
    
    SusceptibilityScaleFactor = val(cmbSusceptibilityScaleFactor)
    
    Config_SaveSetting "SusceptibilityCalibration", "SusceptibilityScaleFactor", Str$(SusceptibilityScaleFactor)
    
    frmSusceptibilityMeter.LagTime

End Sub

Private Sub cmdChangerEdit_Click()
    
    If frmMeasure.buttonHalt.Enabled = False Then
        
        ' (September 2007 L Carporzen) Allow to restart a measurement after pressing Halt during the previous
        Flow_Resume
        
        frmMeasure.updateFlowStatus
        
    End If
    
    MainChanger.cmdSeq_Click
    
End Sub

Private Sub cmdChangerOK_Click()
         
         cmdChangerEdit.Enabled = False 'cmdChangerEdit.Enabled = True
         
         FLAG_MagnetUse = True         ' Notify that we're using magnetometer
         
         DisableMagnetCmds             ' Disable buttons that use magnetometer
         
         frmProgram.mnuViewMeasurement.Enabled = True ' Update menu bar
         
         SampQueue.Execute
            
         'Changer_ProcessSamplesToQueue repHolder
         If UseXYTableAPS And MainChanger.optLoadReturn(0).Value Then
            frmDCMotors.MoveToCorner (True)
         Else
            Changer_NearestHole           ' Always park at nearest hole
         End If
         
         FLAG_MagnetUse = False        ' Notify that we're done!
         
         EnableMagnetCmds
         
End Sub

Private Sub cmdManHolder_Click()
    
    'Check to see if the Sample Height is OK
    If SampleHeightCheck(True) = False Then Exit Sub
    
    If Not FLAG_MagnetUse Then
        
        FLAG_MagnetUse = True        ' Notify that we're using magnetometer
        
        DisableMagnetCmds            ' Disable buttons that use magnetometer
        
        Changer_NearestHole          ' Move sample changer to nearest hole
        
        SampleHolder.SampleHeight = val(txtSampleHeight) * UpDownMotor1cm
        
        frmProgram.StatBarNew "Measuring holder..."
        
        'DisplayStatus (4)            ' Measuring Holder...
        frmProgram.mnuViewMeasurement.Enabled = True
        frmProgram.mnuViewMeasurement.Checked = True
        
        Load frmMeasure
        Load frmStats
        
        frmMeasure.HideStats
        frmMeasure.clearStats
        frmMeasure.clearData
        
        frmMeasure.SetSample "Holder"
        frmMeasure.MomentX.Visible = False ' (October 2007 L Carporzen) Susceptibility versus demagnetization
        
        frmMeasure.framJumps.Top = 5040
        frmMeasure.framJumps.Left = 5400
        
        frmMeasure.InitEqualArea ' (August 2007 L Carporzen) Equal area plot
        
        frmMeasure.ZOrder
        frmMeasure.Show
        
''========================================================================================================
'            '(March 10, 2011 - I Hilburn)
'            'This code has been commented out as it is being applied even when the
'            'user has not selected for the susceptibility measurements to be performed
'            'New code has been added in MeasureTreatAndRead in
'            'modMeasure to ensure that the susceptibility lagTime is set during the appropriate
'            'Holder measurements
''--------------------------------------------------------------------------------------------------------
'            If COMPortSusceptibility > 0 And EnableSusceptibility Then frmSusceptibilityMeter.LagTime
''========================================================================================================

        
        ' reset SampleHolder step type to NRM, just in case
        SampleHolder.Parent.measurementSteps(1).StepType = "NRM"
        SampleHolder.Parent.measurementSteps(1).Level = 0

'        Motor_MoveLoadToZero         ' Lower sample to zero position
        'MotorUpDn_Move ZeroPos, 2
        
        Measure_TreatAndRead SampleHolder, False ' Read Holder

'        Motor_MoveZeroToLoad         ' Raise sample back to load position
        
        MotorUpDn_Move 0, 2
        HolderMeasured = True        ' Set the "holder measured" flag
        
        'DisplayStatus (5)            ' Waiting for motor to stop...
 '       Motor_WaitStop ("UPDOWN")              ' Wait for motor
        'DisplayStatus (-1)           ' Clear status bar
        
        frmProgram.StatBarNew vbNullString
        
        FLAG_MagnetUse = False       ' Notify that we stopped
        
        EnableMagnetCmds
        
        SampleHolder.SampleHeight = 0
    
    End If

End Sub

Private Sub cmdManRun_Click()
    Dim ret As VbMsgBoxResult
    Dim specName As String
    Dim specParent As String
    Dim specimen As Sample
    Dim doUpOriginal As Boolean
    If frmMeasure.buttonHalt.Enabled = False Then
        Flow_Resume ' (September 2007 L Carporzen) Allow to restart a measurement after pressing Halt during the previous
        frmMeasure.updateFlowStatus
    End If
    
    If SampleIndexRegistry.Count = 0 Then
        MsgBox "Please press Add to registry..." ' (September 2007 L Carporzen) Need to have something in the registry
        Exit Sub
    End If
    
    
    frmVacuum.MotorPower True
    
    '(March 2011, I Hilburn)
    'Do a sample height check:
    '1) Test to see if the sign of the sample height is correct for the sign of the Motor UpDown position
    '2) Test to see, if AF, ARM, IRM, or susceptibility steps are involved if the Sample Height is too large
    '   for the sample to be raised high enough to reach the needed motor up down position.
    If SampleHeightCheck = False Then Exit Sub
    
    If Not FLAG_MagnetUse Then
        If Not HolderMeasured Then
            ' We haven't measured the holder yet. Query user.
            ret = MsgBox("The holder has not been measured yet, " & _
                "measure it now?", vbOKCancel, "Continue?")
        If ret = vbOK Then
                cmdManHolder_Click
        Else
            Exit Sub
        End If
    End If
    
    If Prog_halted Then Exit Sub ' (September 2007 L Carporzen) New version of the Halt button
    
    FLAG_MagnetUse = True   ' Notify that we're using magnetometer
    DisableMagnetCmds       ' Disable buttons that use magnetometer
    frmProgram.mnuViewMeasurement.Enabled = True   ' Update menu bar
    specName = cmbManSample.text
    specParent = SampleIndexRegistry.SampleFileByIndex(cmbManSample.ListIndex + 1)
    Set specimen = SampleIndexRegistry(specParent).sampleSet(specName)
    SampleHeight = val(txtSampleHeight) * UpDownMotor1cm
    specimen.SampleHeight = SampleHeight
    Load frmMeasure
    Load frmStats
    frmMeasure.clearData
    frmMeasure.HideStats
    frmMeasure.MomentX.Visible = False ' (October 2007 L Carporzen) Susceptibility & demagnetization versus steps
    frmMeasure.framJumps.Top = 5040
    frmMeasure.framJumps.Left = 5400
    frmMeasure.InitEqualArea ' (August 2007 L Carporzen) Equal area plot
    frmMeasure.ZOrder
    frmMeasure.Show
    With specimen.Parent
        doUpOriginal = .doUp
        frmMeasure.SetFields .avgSteps, .curDemagLong, .doUp, .doBoth, .filename
    End With
    frmMeasure.SetSample cmbManSample.text
    frmProgram.StatBarNew "Measuring..."
    frmMeasure.ZOrder
    frmMeasure.Show
        ' if we're in rockmag mode and IRM is enabled, discharge the IRM coil
        ' before loading the sample
    If specimen.Parent.RockmagMode And EnableAxialIRM Then
        frmIRMARM.optCoil(0).Value = True
        frmIRMARM.FireIRM 0
    End If
    
    Measure_QueryLoad frmMeasure.GetSample, frmMeasure.getMeasDir
    frmProgram.mnuViewMeasurement.Checked = True
    With specimen.Parent
        MotorUpDn_Move Int(SCoilPos + AFPos / 2), 1
        Measure_TreatAndRead specimen, False    ' Read data starting from zero pos
        MotorUpDn_Move 0, 2
        If .doUp And .doBoth Then
            specimen.Parent.doUp = False
            Measure_QueryLoad frmMeasure.GetSample, Magnet_SampleOrientationDown
            frmMeasure.ZOrder
            frmMeasure.Show
            frmProgram.mnuViewMeasurement.Checked = True
            MotorUpDn_Move Int(ZeroPos + specimen.SampleHeight / 2), 2
            Measure_TreatAndRead specimen, False   ' Read data starting from zero pos
            MotorUpDn_Move 0, 2
            .doUp = doUpOriginal
        End If
    End With
        
    If Prog_halted Then ' (September 2007 L Carporzen) New version of the Halt button
        Flow_Resume
        frmMeasure.updateFlowStatus
        Exit Sub
    End If
    
    If specimen.Parent.measurementSteps.Count > 1 Then
        SetCodeLevel CodeOrange
        frmSendMail.MailNotification "Sample done", "Sample " & specName & " done. Please remove sample.", CodeOrange
    End If
        
    If chkVacuum.Value = Unchecked Then ' (August 2007 L Carporzen) Vacuum could stay after the sample done
        MsgBox "Sample " & specName & " done. Please remove sample."
    Else
        MsgBox "Sample " & specName & " done. The vacuum will stay on."
    End If
        
        
        
        '(July 2011 - I Hilburn)
        'Added this in to prompt the user to turn off the air and
        'to turn off the power to the coil thermal sensors
        If specimen.Parent.RockmagMode = True And _
           (EnableT1 Or EnableT2) _
        Then
        'Automatically turn off the air
            If modConfig.DoDegausserCooling = True Then
            frmVacuum.DegausserCooler False
        'Prompt user to turn off the air and the temperature sensor power
                MsgBox "Please: " & vbNewLine & vbNewLine & _
                   " - Verify the air is off" & vbNewLine & _
                   " - Switch off the power to the Rockmag coil temperature sensors"
            Else
            'Prompt user to turn off the air and the temperature sensor power
                MsgBox "Please: " & vbNewLine & vbNewLine & _
                   " - Turn off the air" & vbNewLine & _
                   " - Switch off the power to the Rockmag coil temperature sensors"
            End If
        ElseIf specimen.Parent.RockmagMode = True Then
        
            'Prompt user to turn off the air
            'Automatically turn off the air
            If modConfig.DoDegausserCooling = True Then
                frmVacuum.DegausserCooler False
                MsgBox "Please verify the air is off."
            Else
                MsgBox "Please turn off the air."
            End If 'If modConfig.DoDegausserCooling = True Then
        End If
            SetCodeLevel CodeBlue, True
            SampleNameCurrent = vbNullString
            If chkVacuum.Value = Unchecked Then frmVacuum.ValveConnect False ' (August 2007 L Carporzen) Vacuum could stay
            frmStats.Hide
            frmMeasure.Hide
            FLAG_MagnetUse = False      ' Notify that we're done
            EnableMagnetCmds
        End If  'If Not FLAG_MagnetUse Then
    Set specimen = Nothing
    If chkVacuum.Value = Unchecked Then frmVacuum.MotorPower False ' (August 2007 L Carporzen) Vacuum could stay
End Sub

Private Sub cmdOpenSampleFile_Click()
    
    Dim filename As String
    Dim specName As String, specParent As String
    Dim specimen As Sample
    
    If LenB(cmbManSample.text) = 0 Then Exit Sub
    
    specName = cmbManSample.text
    specParent = SampleIndexRegistry.SampleFileByIndex(cmbManSample.ListIndex + 1)
    
    If SampleIndexRegistry.IsValidSample(specParent, specName) Then
        
        Set specimen = SampleIndexRegistry(specParent).sampleSet(specName)
        filename = specimen.SpecFilePath
        Set specimen = Nothing
    
    End If
    
    DataAnalysis_SampleFile (filename)

End Sub

Public Sub DisableMagnetCmds()
    ' Disable all commands that use the magnetometer
    'tbsControl.Enabled = False
    cmdChangerEdit.Enabled = False
    cmbSusceptibilityScaleFactor.Enabled = False
    cmdChangerOK.Enabled = False
    cmdManHolder.Enabled = False
    cmdManRun.Enabled = False
End Sub

Public Sub EnableMagnetCmds()
    ' This procedure enables the commands that require use of the
    ' magnetometer once the magnetometer is initialized.  We know when
    ' it is initialized when the flag FLAG_MagnetInit is true.
    If FLAG_MagnetInit And Not FLAG_MagnetUse Then
        'tbsControl.Enabled = True
        cmdChangerEdit.Enabled = True
        'cmdChangerOK.Enabled = True ' (October 2007 L Carporzen) Allow to run later
        cmdManHolder.Enabled = True
        cmbSusceptibilityScaleFactor.Enabled = True
 '       frmsettings.cmdVac.Enabled = True
        EnableMagnetRun
    End If
End Sub

Public Sub EnableMagnetRun()
    ' When the sample selected is changed, and a valid sample is
    ' selected, then enable the cmdManRun button, so we can run the
    ' sample.
    If FLAG_MagnetInit And Not FLAG_MagnetUse _
        And (Not cmdManRun.Enabled) Then
        ' The button is not enabled and we are allowed to enable it
        If cmbManSample.ListIndex <> -1 Then
            ' We have selected a sample in the combo box
            cmdManRun.Enabled = True        ' Enable the Run button
        Else
            ' !! Make a way to just type in the same name
        End If
    End If
End Sub

Private Sub Form_Hide()
    'Close all sub forms
    If Me.WindowState <> vbMinimized Then
        Config_SaveSetting "Program", "MagnetometerControlWindowLeft", Str(Me.Left)
        Config_SaveSetting "Program", "MagnetometerControlWindowTop", Str(Me.Top)
    End If
End Sub

Private Sub Form_Load()
    
    Dim i As Integer
    
    '(March 11, 2011 - I Hilburn)
    'Update all the form's icons with the program Ico file
    SetFormIcon Me
    
    ' Initialize private variables
    initialized = False
    HolderMeasured = False
    
    ' Fill normal combo boxes
    cmbSusceptibilityScaleFactor.Clear
    cmbSusceptibilityScaleFactor.AddItem "1.0"
    cmbSusceptibilityScaleFactor.AddItem "0.1"
    cmbSusceptibilityScaleFactor = Format$(SusceptibilityScaleFactor, "0.0")
    cmdOpenSampleFile.Enabled = False
    
    selectTab (1)

End Sub

Private Sub Form_Resize()
    
    ' Dont let the user resize this...
    If Me.WindowState = vbNormal Then
        
        Me.Height = 3465
        Me.Width = 4305
    
    End If

End Sub

Private Sub form_show()

    Me.Left = val(Config_GetSetting("Program", "MagnetometerControlWindowLeft", "0"))
    Me.Top = val(Config_GetSetting("Program", "MagnetometerControlWindowTop", "0"))
    
End Sub

Public Sub RefreshManSampleList()
    ' Adds fields to the combobox
    ' cmbManSamp, so the user can manually select samples to view
    ' or measure.
    Dim i As Integer, j As Integer
    cmbManSample.Clear
On Error GoTo fin
    If SampleIndexRegistry.Count = 0 Then Exit Sub
    For i = 1 To SampleIndexRegistry.Count
        With SampleIndexRegistry(i).sampleSet
        If .Count > 0 Then
            For j = 1 To .Count
                cmbManSample.AddItem .Item(j).Samplename
            Next j
        End If
        End With
    Next i
    On Error GoTo 0
fin:
End Sub

'Pub Function SampleHeightCheck
'
'  Author: I Hilburn
' Created: March 7, 2011
'
' Summary: Function checks the sign and value of the sample height, returning a false
'          if the height does not check out in some way.
'
'          If the height has the wrong sign, the function will change the sign of the height
'          in txtSampleHeight, call cmdManRun_click again and return a false to the present
'          instance of the event handler to close that first instance.
'
'          If the height is too great for the sample to fit in the AF region or Susc. coil
'          region, a message box will pop-up and say the max allowed height
Public Function SampleHeightCheck(Optional ByVal IsHolder As Boolean = False) As Boolean

    Dim SampleHeight As Double
    Dim SampleMotorUnits As Long
    Dim specimen As Sample
    Dim specParent As String
    Dim specName As String
    Dim YesSusc As Boolean
    Dim YesRockmag As Boolean
    Dim i As Long
    
    'Store the sample height to local variable in CM's
    SampleHeight = val(Me.txtSampleHeight.text)
    
    'First check to see if the the SampleMotorUnits have the correct sign
    If Sgn(SampleHeight) = Sgn(modConfig.MeasPos) Then
    
        'Sample height has the wrong sign, it should have the opposite sign
        'as the Measurement Position. (i.e. positive measurement position = negative updownmotor1cm
        'conversion factor, which means to get a positive sample motor height, you need a negative
        'sample height in cm's first).
        'Change the sign of the sample height
        Me.txtSampleHeight.text = Trim(Str(-1 * SampleHeight))
        
        'Return a false
        SampleHeightCheck = False
        
        If IsHolder = False Then
            
            'Call the cmdManRun_Click event handler again
            cmdManRun_Click
            
        Else
        
            'Call the cmdManHolder_Click event handler again
            cmdManHolder_Click
            
        End If
        
        'Exit this function
        Exit Function
        
    End If
    
    'Now, the SampleHeight has the correct sign, can check to make sure
    'that the SampleMotorUnis is such that the sample will fit into the AF or susc. coil region
    
    'Convert that height / 2 to Sample Motor Units (because the changer system centers the sample
    'in the AF or susc. coil; therefore, only need to constrain 1/2 of the sample height.
    SampleMotorUnits = CLng(SampleHeight * modConfig.UpDownMotor1cm / 2)
    
    'Get the Sample Object for this specimen
    specName = cmbManSample.text
    specParent = SampleIndexRegistry.SampleFileByIndex(cmbManSample.ListIndex + 1)
    Set specimen = SampleIndexRegistry(specParent).sampleSet(specName)
    
    'Check to see if the specimen is nothing
    If specimen Is Nothing Then
    
        SampleHeightCheck = False
        
        'Raise a message box to let the user know that they need to select a real sample
        MsgBox "Could not find sample: """ & specName & """ in SAM file: " & vbNewLine & _
               specParent & vbNewLine & vbNewLine & _
               "Please check SAM file or select a different sample.", , _
               "Whoops!"
               
        Exit Function
        
    End If
    
    'Set YesSusc = False
    YesSusc = False
    
    'Check to see if susceptibility measurements are being performed on the sample
    With specimen.Parent
        
        For i = 1 To .measurementSteps.Count
        
            'Set the current step index = i
            .measurementSteps.CurrentStepIndex = i
            
            If .measurementSteps.CurrentStep.MeasureSusceptibility = True Then
            
                'Toggle the measuring susceptibility flag to true
                YesSusc = True
                
                'Exit the for loop
                Exit For
                
            End If
            
        Next i
        
        'Now also check to see if rockmag is being done on this sample
        YesRockmag = .RockmagMode
    
    End With
        
    'If YesSusc = True, and the measure susceptiblity module is enabled
    'then need to make sure the sample will be able to be placed within
    'the susceptibility coils
    If YesSusc = True And _
       modConfig.EnableSusceptibility = True _
    Then
    
        'Will the sample be able to be fit into the susceptibility coils?
        'If No, then return false and exit the function
        If Abs(SampleMotorUnits) > Abs(SCoilPos) Then
        
            'Sample height is too large
            'Pop-up a message box
            'Note: Max allowed height = distance in CM to the susc. coil - 0.5 cm for slop
            MsgBox "Sample Height is too large to raise the sample into the susceptibility coil." & _
                   vbNewLine & _
                   "Current Sample Height = " & Trim(Str(SampleHeight)) & " cm" & _
                   vbNewLine & _
                   "Max. Allowed Height = " & Format(-2 * (SCoilPos / UpDownMotor1cm) - 0.5, "0.00") & " cm", , _
                   "Bad Height!"
                   
            'Return false
            SampleHeightCheck = False
            
            Exit Function
            
        End If
        
    End If
    
    'Now check to see if the sample height is too large to do rockmag
    If YesRockmag And _
       Abs(SampleMotorUnits) > Abs(AFPos) _
    Then
    
        'Wha-oh! Sample height is too large to do rockmag = AF / ARM / IRM
        'Pop-up a message box
        'Note: Max allowed height = distance in cm to the AF coil center - 0.5 cm for slop
        MsgBox "Sample Height is too large to raise the sample into the AF / IRM / ARM coils." & _
               vbNewLine & _
               "Current Sample Height = " & Trim(Str(SampleHeight)) & " cm" & _
               vbNewLine & _
               "Max. Allowed Height = " & Format(-2 * (AFPos / UpDownMotor1cm) - 0.5, "0.00") & " cm", , _
               "Bad Height!"

        'Return False
        SampleHeightCheck = False
        
        Exit Function
        
    End If
    
    'Return True
    SampleHeightCheck = True
                   
End Function

Private Sub selectTab(tabtoselect As Integer)
    Dim i As Integer
    'show and enable the selected tab's controls
    'and hide and disable all others
    RefreshManSampleList
    For i = 0 To tbsControl.Tabs.Count - 1
        If i = tabtoselect - 1 Then
            frameControl(i).Visible = True
            frameControl(i).Enabled = True
            frameControl(i).ZOrder 0
        Else
            frameControl(i).Visible = False
            frameControl(i).Enabled = False
        End If
    Next
End Sub

Private Sub tbsControl_Click()
    selectTab tbsControl.SelectedItem.Index
End Sub

