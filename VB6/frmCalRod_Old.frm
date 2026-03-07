VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmCalRod 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "WriteCalRodData WriteFile, CurrentPosition, Int(frmDCMotors.UpDownHeight), currentData"
   ClientHeight    =   4965
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6315
   Icon            =   "frmCalRod.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   6315
   Begin VB.OptionButton optOrientation 
      Caption         =   "Down"
      Height          =   255
      Index           =   1
      Left            =   4080
      TabIndex        =   38
      Top             =   1800
      Width           =   975
   End
   Begin VB.OptionButton optOrientation 
      Caption         =   "Up"
      Height          =   255
      Index           =   0
      Left            =   4080
      TabIndex        =   37
      Top             =   1440
      Width           =   975
   End
   Begin VB.TextBox txtUpDownSampleCm 
      Height          =   285
      Left            =   960
      TabIndex        =   36
      Text            =   "Text1"
      Top             =   3600
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtData 
      Height          =   285
      Index           =   4
      Left            =   4080
      TabIndex        =   33
      Text            =   "Text1"
      Top             =   3600
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtData 
      Height          =   285
      Index           =   3
      Left            =   2640
      TabIndex        =   32
      Text            =   "Text1"
      Top             =   3600
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CheckBox chkSingleScan 
      Caption         =   "Single Scan"
      Height          =   255
      Left            =   120
      TabIndex        =   31
      Top             =   4680
      Width           =   1575
   End
   Begin MSComDlg.CommonDialog dlgOpenCreateFile 
      Left            =   4680
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdBrowseForFile 
      Caption         =   "..."
      Height          =   255
      Left            =   5640
      TabIndex        =   29
      Top             =   2280
      Width           =   495
   End
   Begin VB.TextBox txtCoreLength 
      Height          =   288
      Left            =   1920
      TabIndex        =   26
      Top             =   1800
      Width           =   972
   End
   Begin VB.TextBox txtStartPosition 
      Height          =   288
      Left            =   1920
      TabIndex        =   24
      Top             =   960
      Width           =   972
   End
   Begin VB.CheckBox chkUchannel 
      Caption         =   "U channel"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   4440
      Width           =   1095
   End
   Begin VB.TextBox txtStanHeight 
      Height          =   285
      Left            =   1920
      TabIndex        =   0
      Text            =   "2.415"
      Top             =   90
      Width           =   375
   End
   Begin VB.CheckBox chksusceptibility 
      Caption         =   "Kvol (x10-6 CGS):"
      Height          =   255
      Left            =   2880
      TabIndex        =   1
      Top             =   120
      Value           =   1  'Checked
      Width           =   1575
   End
   Begin VB.TextBox txtSusce 
      Height          =   285
      Left            =   4440
      TabIndex        =   2
      Text            =   "2473"
      Top             =   90
      Width           =   495
   End
   Begin VB.TextBox txtSpacing 
      Height          =   285
      Left            =   1920
      TabIndex        =   3
      Text            =   "0.3"
      Top             =   450
      Width           =   375
   End
   Begin VB.TextBox txtThreshold 
      Height          =   285
      Left            =   1200
      TabIndex        =   4
      Text            =   "10"
      Top             =   2250
      Width           =   375
   End
   Begin VB.TextBox txtFileName 
      Height          =   645
      Left            =   3120
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Text            =   "CalibRod"
      Top             =   2250
      Width           =   2415
   End
   Begin VB.TextBox txtUpDown 
      Height          =   285
      Left            =   960
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   3090
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtData 
      Height          =   285
      Index           =   0
      Left            =   2040
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   3090
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtData 
      Height          =   285
      Index           =   1
      Left            =   3120
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   3090
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtData 
      Height          =   285
      Index           =   2
      Left            =   4200
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   3090
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdStartStop 
      Caption         =   "Start"
      Default         =   -1  'True
      Height          =   435
      Left            =   2040
      TabIndex        =   10
      Top             =   4080
      Width           =   975
   End
   Begin VB.CommandButton cmdResume 
      Caption         =   "Resume"
      Height          =   435
      Left            =   3960
      TabIndex        =   11
      Top             =   4080
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblDec 
      Caption         =   "Dec"
      Height          =   255
      Left            =   3720
      TabIndex        =   35
      Top             =   3600
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblInc 
      Caption         =   "Inc"
      Height          =   255
      Left            =   2280
      TabIndex        =   34
      Top             =   3630
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblEndPosition 
      Caption         =   "End Position:"
      Height          =   255
      Left            =   120
      TabIndex        =   30
      Top             =   1440
      Width           =   3975
   End
   Begin VB.Label Label12 
      Caption         =   "cm"
      Height          =   255
      Left            =   3000
      TabIndex        =   28
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label11 
      Caption         =   "Core Piece Length:"
      Height          =   255
      Left            =   120
      TabIndex        =   27
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label Label9 
      Caption         =   "motor counts"
      Height          =   255
      Left            =   3000
      TabIndex        =   25
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label7 
      Caption         =   "Starting Up/Down Motor Position:"
      Height          =   375
      Left            =   120
      TabIndex        =   23
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Standard sample height:"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "cm"
      Height          =   255
      Left            =   2400
      TabIndex        =   13
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label4 
      Caption         =   "Finest spacing:"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "cm"
      Height          =   252
      Left            =   2400
      TabIndex        =   15
      Top             =   480
      Width           =   252
   End
   Begin VB.Label Label6 
      Caption         =   "Tolerance (%):"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Write to file:"
      Height          =   255
      Left            =   2160
      TabIndex        =   17
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label UpDown 
      Caption         =   "Up/Down"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   3120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label X 
      Caption         =   "X"
      Height          =   255
      Left            =   1800
      TabIndex        =   19
      Top             =   3120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Y 
      Caption         =   "Y"
      Height          =   255
      Left            =   2880
      TabIndex        =   20
      Top             =   3120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Z 
      Caption         =   "Z"
      Height          =   255
      Left            =   3960
      TabIndex        =   21
      Top             =   3120
      Visible         =   0   'False
      Width           =   255
   End
End
Attribute VB_Name = "frmCalRod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' (June 2007 L Carporzen) Form to adjust automatically the positions each time a new rod is installed.
' All we need is approximatly good (+-5 cm) positions in Paleomag.INI and the Bartington susceptibility standard (white cylinder placed on top of the MS2B box).
' NB that the susceptibility is very sensitive to the altitude in order to have reproducible measurement.
' The susceptibility position should be easy to define at less than 1 cm prior to running that sequence.
' (March 2008 L Carporzen) It can also correct the negative sample height that some systems are still using.
' (October 2009 - Feb 2010 I Hilburn) Added new functionality to the Long Core (UChannel) measurement subroutine
' Also added file browser to frmCalRod.
'      IMPORTANT NOTE in regards to the changes:  There are two values used in the RUChannel that will be
'      different for each sample changer system:
'                                       The maximum end position for the scan, and the position of the floor
'                                       relative to the up/down arm of the sample changer
'      These values are stored in the two new global variables:
'                                       UpDownMinPosition = -30500 for the Shoemaker sample changer at Caltech
'                                       UpDownFloorPosition = -43000 for the Shoemaker sample changer  "  "
'      If this is the first time you are running this code on your sample changer system, these
'      two values will need to be determined and changed below in the code at the place indicated
'      by my comments.
'
'      To determine the UpDownMinPosition, you will need to move the Up/Down arm using the DC Motors
'      window as far down the up/down arm as it will go without getting stuck.  This value should NOT
'      be the furthest down the up/down carriage can go (that value for Shoemaker is 31037) but rather,
'      the furthest down value that the up/down carriage can go without getting stuck.
'
'      Determining the floor position is a little harder.  With a meter/yard stick or a tape measure,
'      measure the distance between the vacuum tube plastic guide on the sample changer tray and
'      the level of the first obstruction beneath your magnetometer.  (For Shoemaker, the first obstruction
'      is a foam pad that overlies the floor of the Shoemaker shielded room).
'
'      Convert this distance into centimeters and then multiply it by the value UpDownMotor1CM.
'      For safety's sake, reduce this value by 1000 motor position units. (i.e. a value of -44000 would
'      be changed to -43000, and a value of 44000 would be changed to 43000).  This value then
'      is your UpDownFloorPosition
'
Dim CurrentlyRunning As Boolean
Dim Susce As Double
Dim StanHeight As Double
Dim Spacing As Double
Dim Threshold As Double
Dim WriteFile As String
Dim UpDownMinPosition As Double
Dim UpDownFloorPosition As Double

Private Sub cmdBrowseForFile_Click()

    dlgOpenCreateFile.Flags = cdlOFNHideReadOnly Or cdlOFNCreatePrompt
    dlgOpenCreateFile.DialogTitle = "Open or Create File for Long Core Measurement"
    dlgOpenCreateFile.ShowOpen
    
    txtFileName = dlgOpenCreateFile.filename

End Sub

Private Sub cmdResume_Click()
    If Prog_paused = True Then
        Flow_Resume
        frmMeasure.updateFlowStatus
    End If
End Sub

Private Sub cmdStartStop_Click()
    If CurrentlyRunning Then
        CurrentlyRunning = False
        cmdStartStop.Caption = "Start"
    Else
        CurrentlyRunning = True
        StanHeight = Abs(val(txtStanHeight))
        Susce = val(txtSusce) * 0.00001
        Spacing = val(txtSpacing)
        Threshold = val(txtThreshold)
        WriteFile = txtFileName
        cmdStartStop.Caption = "Stop"
        Run
        cmdStartStop.Caption = "Start"
        CurrentlyRunning = False
        frmCalRod.Hide
    End If
End Sub

Private Sub Run()
    Dim currentData As Cartesian3D
    Dim CurrentPosition As Long
    Dim Xavg As Cartesian3D
    Dim Max As Cartesian3D
    Dim j As Long
    If CurrentlyRunning And chkUchannel.value = Checked Then
    UpDown.Visible = True
    X.Visible = True
    Y.Visible = True
    Z.Visible = True
    txtUpDown = ""
    txtUpDown.Visible = True
    For j = 0 To 2
        txtData(j) = ""
        txtData(j).Visible = True
    Next j
     RunUchannel
    Else
    If CurrentlyRunning Then RunSampleChangerSeq
    DelayTime 5
    UpDown.Visible = True
    X.Visible = True
    Y.Visible = True
    Z.Visible = True
    txtUpDown = ""
    txtUpDown.Visible = True
    For j = 0 To 2
        txtData(j) = ""
        txtData(j).Visible = True
    Next j
    If CurrentlyRunning Then
        RunCalRodSeq Max, Xavg
    Else
        Set Max = New Cartesian3D
        Set Xavg = New Cartesian3D
    End If
    DelayTime 5
    If CurrentlyRunning And chksusceptibility.value = Checked Then RunSusceSeq
    cmdResume.Visible = True
    MotorUpDn_Move 0, 1
    frmDCMotors.ChangerMotortoHole 1
    frmDCMotors.SampleDropOff
    frmVacuum.ValveConnect (False)
    MotorUpDn_Move 0, 1
    frmDCMotors.ChangerMotortoHole 200
    cmdResume.Visible = False
    frmVacuum.MotorPower False
    UpDown.Visible = False
    X.Visible = False
    Y.Visible = False
    Z.Visible = False
    txtUpDown.Visible = False
    For j = 0 To 2
        txtData(j).Visible = False
    Next j
    If Not CurrentlyRunning Then End ' Complete end of the program
    Set currentData = New Cartesian3D
    currentData.X = 0
    currentData.Y = 0
    currentData.Z = 0
    j = ZeroPos
    WriteCalRodDataEnd WriteFile, "Zeroing position", j, currentData
    currentData.X = Max.X
    currentData.Y = Max.Y
    currentData.Z = Max.Z
    j = MeasPos
    WriteCalRodDataEnd WriteFile, "Measuring position", j, currentData
    WriteCalRodDataEnd WriteFile, "Three maximum", j, Xavg
    j = AFPos
    WriteCalRodDataEnd WriteFile, "AF position", j, currentData
    j = IRMPos
    WriteCalRodDataEnd WriteFile, "IRM position", j, currentData
    Set currentData = New Cartesian3D
    currentData.X = Susce * (1 - Threshold / 100)
    currentData.Y = Susce
    currentData.Z = Susce * (1 + Threshold / 100)
    j = SCoilPos
    WriteCalRodDataEnd WriteFile, "S Coil position", j, currentData
    j = SusceptibilityMomentFactorCGS / 0.000000001
    WriteCalRodDataEnd WriteFile, "S Coil correction factor x 10-9", j, currentData
    Set currentData = New Cartesian3D
    currentData.X = Format$((SampleTop - SampleBottom) / UpDownMotor1cm, "0.00")
    currentData.Y = SampleTop
    currentData.Z = SampleBottom
    j = SampleTop - SampleBottom
    WriteCalRodDataEnd WriteFile, "Default samp. height (in units)- (in cm)- top and bot.", j, currentData
    j = Threshold
    WriteCalRodDataEnd WriteFile, "Tolerance (%)", j, currentData
    If ((chksusceptibility.value = Checked And COMPortSusceptibility < 1) Or NOCOMM_MODE) Then
    MsgBox ("The new values are " & MeasPos & " for the measurement position, " & _
    SCoilPos & " for the S coil position, " & AFPos & " for the AF position." & _
    Chr(13) + "The file " & WriteFile & " recorded all the data." + Chr(13) + "You should verify the AF and IRM altitudes." & _
    Chr(13) + "Click OK in the settings window to save them." + Chr(13) + "Then, restart the program.")
    Else
    MsgBox ("The new values are " & MeasPos & " for the measurement position, " & _
    SCoilPos & " for the S coil position, " & AFPos & " for the AF position." & _
    Chr(13) + "The file " & WriteFile & " recorded all the data." + Chr(13) + "You should verify the S coil, Af and IRM altitudes." & _
    Chr(13) + "Click OK in the settings window to save them." + Chr(13) + "Then, restart the program.")
    End If
    End If
End Sub

Private Sub RunSampleChangerSeq()
    If NOCOMM_MODE Then currentPosInitialized = True
    If Not currentPosInitialized Then frmChanger.GetCurrentChangerPos
    FLAG_MagnetUse = True                                ' Notify that we're using magnetometer
    frmMagnetometerControl.DisableMagnetCmds             ' Disable buttons that use magnetometer
    frmProgram.mnuViewMeasurement.Enabled = True         ' Update menu bar
    cmdResume.Visible = True
    frmDCMotors.TurningMotorAngleOffset TrayOffsetAngle
    If frmDCMotors.UpDownHeight > 100 Then frmDCMotors.HomeToTop
    WriteCalRodDataHeaders WriteFile
    frmDCMotors.ChangerMotortoHole 1
    frmDCMotors.SamplePickup
    If Not NOCOMM_MODE Then SampleBottom = frmDCMotors.UpDownHeight + (frmDCMotors.UpDownHeight / Abs(frmDCMotors.UpDownHeight)) * StanHeight * UpDownMotor1cm * (UpDownMotor1cm / Abs(UpDownMotor1cm))
    frmDCMotors.HomeToTop
    frmDCMotors.ChangerMotortoHole 199
    frmDCMotors.UpDownMove SampleBottom, 0
    If Not NOCOMM_MODE Then SampleBottom = frmDCMotors.UpDownHeight
    frmDCMotors.HomeToTop
    frmVacuum.MotorPower True
    DelayTime 0.5
    frmDCMotors.ChangerMotortoHole 1
    frmDCMotors.SamplePickup
    If Not NOCOMM_MODE Then SampleTop = frmDCMotors.UpDownHeight
    frmVacuum.ValveConnect True
    DelayTime 0.3
    frmDCMotors.HomeToTop
    frmDCMotors.ChangerMotortoHole 200
    cmdResume.Visible = False
    If Abs(SampleTop - SampleBottom) > (1 + 0.5 * Threshold / 100) * StanHeight * Abs(UpDownMotor1cm) Or Abs(SampleTop - SampleBottom) < (1 - 0.5 * Threshold / 100) * StanHeight * Abs(UpDownMotor1cm) Then
        MsgBox ("The sample changer plate is too elastic (standard between " & Int(SampleBottom) & " and " & Int(SampleTop) & ", recovered standard height " & Format$((SampleTop - SampleBottom) / UpDownMotor1cm, "0.00") & " cm), you need to calibrate the default sample top and bottom manually.")
        SampleBottom = val(frmSettings.txtSampleBottom)
        SampleTop = val(frmSettings.txtSampleTop)
        If NOCOMM_MODE Then MsgBox ("The new values are " & Int(SampleBottom + ((SampleTop - SampleBottom) / Abs(SampleTop - SampleBottom)) * SampleHeight * (SampleHeight / Abs(SampleHeight))) & " and " & Int(SampleBottom) & " for the default sample top and bottom, recovered standard height " & Format$((SampleTop - SampleBottom) / UpDownMotor1cm, "0.00") & " cm")
    Else
        MsgBox ("The new values are " & Int(SampleBottom + ((SampleTop - SampleBottom) / Abs(SampleTop - SampleBottom)) * SampleHeight * (SampleHeight / Abs(SampleHeight))) & " and " & Int(SampleBottom) & " for the default sample top and bottom, recovered standard height " & Format$((SampleTop - SampleBottom) / UpDownMotor1cm, "0.00") & " cm")
        SampleTop = SampleBottom + ((SampleTop - SampleBottom) / Abs(SampleTop - SampleBottom)) * SampleHeight * (SampleHeight / Abs(SampleHeight))
    End If
    frmSettings.txtSampleBottom = SampleBottom
    frmSettings.txtSampleTop = SampleTop
End Sub

Private Sub RunUchannel()
    Dim Warning As Boolean
    Dim currentData As Cartesian3D
    Dim FirstData As Cartesian3D
    Dim SM As Cartesian3D
    Dim SXM As Cartesian3D
    Dim GaussianX As Variant
    Dim GaussianY As Variant
    Dim GaussianZ As Variant
    Dim RodPosition As Variant
    Dim CurrentPosition As Long
    Dim coarse As Long
    Dim fine As Long
    Dim i As Long
    Dim Delay As Double
    Dim StartPosition As Long
    Dim EndPosition As Long
    Dim SampleLength As Long
    Dim SampleMeasurement As MeasurementBlock
    
'*********************************************************************************************
'---------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------
'           IMPORTANT NOTE!!!!
'           Feb 25, 2010
'           I Hilburn
'
'           This code sets two global variables whose values are specific to the particular
'           sample changer system being used.
'
'           Therefore, their values need to be changed here in the code if you are using this
'           code for the first time.  The values stored here are for the Shoemaker sample changer
'           system at Caltech.
'
'           Future versions of this code will write and read these values from the
'           paleomag.ini config file.
'---------------------------------------------------------------------------------------------
'   'Set globals UpDownMinPosition and UpDownFloorPosition
    UpDownMinPosition = -35000
    UpDownFloorPosition = -43000
'---------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------
'*********************************************************************************************

    MsgBox ("Place the sample.")
    Warning = False
    If NOCOMM_MODE Then
    Delay = 0
    Threshold = 99
    currentPosInitialized = True
    Else
    Delay = 2
    End If
    i = 0
    cmdResume.Visible = False
    
    Set SampleMeasurement = Nothing
    Set SampleMeasurement = New MeasurementBlock
    
    For i = 1 To 2
    
        
        SampleMeasurement.Baselines(i).X = 0
        SampleMeasurement.Baselines(i).Y = 0
        SampleMeasurement.Baselines(i).Z = 0
        
    Next i
    
    For i = 1 To 4
    
        SampleMeasurement.Holder(i).X = 0
        SampleMeasurement.Holder(i).Y = 0
        SampleMeasurement.Holder(i).Z = 0
        
    Next i
    
    i = 0
    
    Set currentData = New Cartesian3D
    Set FirstData = New Cartesian3D
    frmSQUID.CLP "A"
    frmSQUID.ResetCount "A"
    frmProgram.StatusBar "Resetting...", 3
    DelayTime (Delay)
    Set currentData = frmSQUID.getData
    FirstData.X = currentData.X
    FirstData.Y = currentData.Y
    FirstData.Z = currentData.Z

    
    SampleMeasurement.isUp = optOrientation(0)
        
    txtUpDown = CurrentPosition
    
    txtData(0) = Str$(currentData.X - FirstData.X)
    txtData(1) = Str$(currentData.Y - FirstData.Y)
    txtData(2) = Str$(currentData.Z - FirstData.Z)
    txtData(3).Visible = True
    txtData(4).Visible = True
    txtData(3) = ""
    txtData(4) = ""
    
    
    txtUpDownSampleCm.Visible = True
    txtUpDownSampleCm = ""
    
    lblInc.Visible = True
    lblDec.Visible = True
    
    
    '-------------------------------------------------
    '   Added October 8, 2009 - I. Hilburn
    '-------------------------------------------------
    SampleLength = CLng(val(txtCoreLength) * UpDownMotor1cm)
        
    If SampleLength < 0 Or SampleLength > 30000 Then
    
        'User has entered a bad / invalid sample length, pop up an error and exit the sub-routine
        MsgBox "User entered a core sample length: " & vbNewLine & _
                txtCoreLength & " cm" & vbNewLine & _
                Trim(Str(SampleLength)) & " motor counts", , "Bad Core Length!"
                
        Exit Sub
        
    End If
        
   
    StartPosition = CLng(val(txtStartPosition))
    
    If Abs(StartPosition) < 0 Or Abs(StartPosition) > 30500 Then
    
        'User has selected a crappy start position, send an error and exit the subroutine.
        MsgBox "Bad UChannel scan start motor count position:" & vbNewLine & _
                txtStartPosition, , "Bad Start Position!"
    
        Exit Sub
        
    End If
    
    'Adjust starting position based on the sample length
    StartPosition = StartPosition - MeasPos / Abs(MeasPos) * SampleLength
    
    
    'Adjust End Position based on the sample length and already adjusted start position
    EndPosition = (UpDownMinPosition * UpDownMinPosition / Abs(UpDownMinPosition) + Abs(SampleLength)) * MeasPos / Abs(MeasPos)
    
    If Abs(EndPosition) > UpDownFloorPosition * UpDownFloorPosition / Abs(UpDownFloorPosition) Then  'UpDownFloorPosition will be different for other sample changer systems
    
        MsgBox "Sample is too long - it will smack into the floor." & vbNewLine & vbNewLine & _
                "You will need to measure the sample in the opposite orientation to get full" & _
                "coverage."
    
        EndPosition = UpDownFloorPosition * UpDownFloorPosition / Abs(UpDownFloorPosition)  'UpDownFloorPosition will be different for other sample changer systems

    End If
    
    EndPosition = EndPosition - Abs(SampleLength) * MeasPos / Abs(MeasPos)
    lblEndPosition.Caption = "End Position:     " & Trim(Str(EndPosition)) & "  motor counts"
    
    '--------------------------------------------
    
    CurrentPosition = CInt(StartPosition)     'To let user input start position of UChannel scan (Oct 2009, I. Hilburn)
    
    txtUpDownSampleCm = Trim(Str((Abs(CurrentPosition) - Abs(StartPosition)) / UpDownMotor1cm))
    
    Do While Not Warning
        If Not CurrentlyRunning Then Exit Sub
        CurrentPosition = CurrentPosition + (ZeroPos / Abs(ZeroPos)) * Spacing * UpDownMotor1cm * (UpDownMotor1cm / Abs(UpDownMotor1cm))
        If Abs(CurrentPosition) <= Abs(EndPosition) Then  '(Oct 2009, I. Hilburn) - End position set currently to the maximum travel
                                                        'possible for the Shoemaker up/down arm.  This value is machine specific and needs
                                                        'to be changed for other sample changer systems!!
            i = i + 1
            coarse = i
            
        Else
            Warning = True
        End If
    Loop
    Warning = False
    ReDim GaussianX(coarse + 1)
    ReDim GaussianY(coarse + 1)
    ReDim GaussianZ(coarse + 1)
    ReDim RodPosition(coarse + 1)
    Set SM = New Cartesian3D
    Set SXM = New Cartesian3D
    Set Xavg = New Cartesian3D
    Set Max = New Cartesian3D
    SM.X = 0
    SM.Y = 0
    SM.Z = 0
    SXM.X = 0
    SXM.Y = 0
    SXM.Z = 0
    Max.X = 0
    Max.Y = 0
    Max.Z = 0
    CurrentPosition = CInt(StartPosition)
    frmDCMotors.TurningMotorRotate 0, False
    For i = 1 To coarse
        If Not CurrentlyRunning Then Exit Sub
        CurrentPosition = CurrentPosition + (ZeroPos / Abs(ZeroPos)) * Spacing * UpDownMotor1cm * (UpDownMotor1cm / Abs(UpDownMotor1cm))
        frmDCMotors.UpDownMove CurrentPosition, 0
        DelayTime (Delay)
        Set currentData = frmSQUID.getData
        If NOCOMM_MODE Then
            currentData.X = 0.9 * Abs(1 / (MeasPos - CurrentPosition - (UpDownMotor1cm / Abs(UpDownMotor1cm)) * (CurrentPosition / Abs(CurrentPosition)) * StanHeight * UpDownMotor1cm / 2))
            currentData.Y = 0.95 * Abs(1 / (MeasPos - CurrentPosition - (UpDownMotor1cm / Abs(UpDownMotor1cm)) * (CurrentPosition / Abs(CurrentPosition)) * StanHeight * UpDownMotor1cm / 2))
            currentData.Z = Abs(1 / (MeasPos - CurrentPosition - (UpDownMotor1cm / Abs(UpDownMotor1cm)) * (CurrentPosition / Abs(CurrentPosition)) * StanHeight * UpDownMotor1cm / 2))
            GaussianX(i) = currentData.X
            GaussianY(i) = currentData.Y
            GaussianZ(i) = currentData.Z
            RodPosition(i) = CurrentPosition
        Else
            currentData.X = currentData.X - FirstData.X
            currentData.Y = currentData.Y - FirstData.Y
            currentData.Z = currentData.Z - FirstData.Z
            GaussianX(i) = currentData.X
            GaussianY(i) = currentData.Y
            GaussianZ(i) = currentData.Z
            RodPosition(i) = frmDCMotors.UpDownHeight
        End If
        txtUpDown = CurrentPosition
        txtUpDownSampleCm = Trim(Str((Abs(CurrentPosition) - Abs(StartPosition)) / UpDownMotor1cm))
    
        txtData(0) = Str$(currentData.X)
        txtData(1) = Str$(currentData.Y)
        txtData(2) = Str$(currentData.Z)
        
        WriteCalRodData WriteFile, CurrentPosition, Int(frmDCMotors.UpDownHeight), currentData
        
        For k = 1 To 4
        
            SampleMeasurement.Sample(k).X = currentData.X
            SampleMeasurement.Sample(k).Y = currentData.Y
            SampleMeasurement.Sample(k).Z = currentData.Z
        
        
        Next k
        
        Set currentData = SampleMeasurement.CorrectedSample(2)
        
        txtData(3) = Str$(currentData.inc)
        txtData(4) = Str$(currentData.dec)
        
        
        Set currentData = Nothing
                
    Next i
    
    'Home to top and end sub if signle scan is checked by user (10/23/2009, I Hilburn)
    If chkSingleScan.value = Checked Then
    
        
        frmDCMotors.HomeToTop
        
        'Notify User that the sample is done
        frmSendMail.MailNotification "2G Status Update", _
                                    "Long core sample done.  Sample saved to file: " & Me.txtFileName
        
        Exit Sub
        
    End If
    
    frmDCMotors.TurningMotorRotate 90, False
    
    CurrentPosition = Abs(EndPosition) * MeasPos / Abs(MeasPos)
    Warning = False
    
    
    Do While Not Warning
        If Not CurrentlyRunning Then Exit Sub
        CurrentPosition = CurrentPosition - (ZeroPos / Abs(ZeroPos)) * _
                            (UpDownMotor1cm / Abs(UpDownMotor1cm)) * Spacing * UpDownMotor1cm
        If Abs(CurrentPosition) >= Spacing * Abs(UpDownMotor1cm) Then 'To don't hit the switch...
            frmDCMotors.UpDownMove CurrentPosition, 1
            DelayTime (Delay * 0.1)
            Set currentData = frmSQUID.getData
            If NOCOMM_MODE Then
            currentData.X = 0.9 * Abs(1 / (MeasPos - CurrentPosition - (UpDownMotor1cm / Abs(UpDownMotor1cm)) _
                                * (CurrentPosition / Abs(CurrentPosition)) * StanHeight * UpDownMotor1cm / 2))
            currentData.Y = 0.95 * Abs(1 / (MeasPos - CurrentPosition - (UpDownMotor1cm / Abs(UpDownMotor1cm)) _
                                * (CurrentPosition / Abs(CurrentPosition)) * StanHeight * UpDownMotor1cm / 2))
            currentData.Z = Abs(1 / (MeasPos - CurrentPosition - (UpDownMotor1cm / Abs(UpDownMotor1cm)) * _
                            (CurrentPosition / Abs(CurrentPosition)) * StanHeight * UpDownMotor1cm / 2))
            Else
            currentData.X = currentData.X - FirstData.X
            currentData.Y = currentData.Y - FirstData.Y
            currentData.Z = currentData.Z - FirstData.Z
            End If
            txtUpDown = CurrentPosition
            txtUpDownSampleCm = Trim(Str((Abs(CurrentPosition) - Abs(StartPosition)) / UpDownMotor1cm))
        
            txtData(0) = Str$(currentData.X)
            txtData(1) = Str$(currentData.Y)
            txtData(2) = Str$(currentData.Z)
            
            WriteCalRodData WriteFile, CurrentPosition, Int(frmDCMotors.UpDownHeight), currentData
            
            For k = 1 To 4
        
                SampleMeasurement.Sample(k).X = currentData.X
                SampleMeasurement.Sample(k).Y = currentData.Y
                SampleMeasurement.Sample(k).Z = currentData.Z
            
            
            Next k
            
            Set currentData = SampleMeasurement.CorrectedSample(2)
                
            txtData(3) = Str$(currentData.inc)
            txtData(4) = Str$(currentData.dec)
        
            Set currentData = Nothing
        Else
            Warning = True
        End If
        
        If Abs(CurrentPosition) <= Abs(StartPosition) Then
        
            Warning = True
            
        End If
            
        
    Loop
    
    CurrentPosition = CInt(StartPosition)
    frmDCMotors.TurningMotorRotate 180, False
    For i = 1 To coarse
        If Not CurrentlyRunning Then Exit Sub
        CurrentPosition = CurrentPosition + (ZeroPos / Abs(ZeroPos)) * Spacing * UpDownMotor1cm * (UpDownMotor1cm / Abs(UpDownMotor1cm))
        frmDCMotors.UpDownMove CurrentPosition, 0
        DelayTime (Delay)
        Set currentData = frmSQUID.getData
        If NOCOMM_MODE Then
            currentData.X = 0.9 * Abs(1 / (MeasPos - CurrentPosition - (UpDownMotor1cm / Abs(UpDownMotor1cm)) * (CurrentPosition / Abs(CurrentPosition)) * StanHeight * UpDownMotor1cm / 2))
            currentData.Y = 0.95 * Abs(1 / (MeasPos - CurrentPosition - (UpDownMotor1cm / Abs(UpDownMotor1cm)) * (CurrentPosition / Abs(CurrentPosition)) * StanHeight * UpDownMotor1cm / 2))
            currentData.Z = Abs(1 / (MeasPos - CurrentPosition - (UpDownMotor1cm / Abs(UpDownMotor1cm)) * (CurrentPosition / Abs(CurrentPosition)) * StanHeight * UpDownMotor1cm / 2))
            GaussianX(i) = currentData.X
            GaussianY(i) = currentData.Y
            GaussianZ(i) = currentData.Z
            RodPosition(i) = CurrentPosition
        Else
            currentData.X = currentData.X - FirstData.X
            currentData.Y = currentData.Y - FirstData.Y
            currentData.Z = currentData.Z - FirstData.Z
            GaussianX(i) = currentData.X
            GaussianY(i) = currentData.Y
            GaussianZ(i) = currentData.Z
            RodPosition(i) = frmDCMotors.UpDownHeight
        End If
        txtUpDown = CurrentPosition
        txtUpDownSampleCm = Trim(Str((Abs(CurrentPosition) - Abs(StartPosition)) / UpDownMotor1cm))
    
        txtData(0) = Str$(currentData.X)
        txtData(1) = Str$(currentData.Y)
        txtData(2) = Str$(currentData.Z)
        
        WriteCalRodData WriteFile, CurrentPosition, Int(frmDCMotors.UpDownHeight), currentData
        
        For k = 1 To 4
        
            SampleMeasurement.Sample(k).X = currentData.X
            SampleMeasurement.Sample(k).Y = currentData.Y
            SampleMeasurement.Sample(k).Z = currentData.Z
        
        
        Next k
        
        Set currentData = SampleMeasurement.CorrectedSample(2)
        
        txtData(3) = Str$(currentData.inc)
        txtData(4) = Str$(currentData.dec)
        
        Set currentData = Nothing
                
    Next i
    
    frmDCMotors.TurningMotorRotate 270, False
    
    CurrentPosition = Abs(EndPosition) * MeasPos / Abs(MeasPos)
    Warning = False
    
    Do While Not Warning
        If Not CurrentlyRunning Then Exit Sub
        CurrentPosition = CurrentPosition - (ZeroPos / Abs(ZeroPos)) * _
                            (UpDownMotor1cm / Abs(UpDownMotor1cm)) * Spacing * UpDownMotor1cm
        If Abs(CurrentPosition) >= Spacing * Abs(UpDownMotor1cm) Then 'To don't hit the switch...
            frmDCMotors.UpDownMove CurrentPosition, 1
            DelayTime (Delay * 0.1)
            Set currentData = frmSQUID.getData
            If NOCOMM_MODE Then
            currentData.X = 0.9 * Abs(1 / (MeasPos - CurrentPosition - (UpDownMotor1cm / Abs(UpDownMotor1cm)) * (CurrentPosition / Abs(CurrentPosition)) * StanHeight * UpDownMotor1cm / 2))
            currentData.Y = 0.95 * Abs(1 / (MeasPos - CurrentPosition - (UpDownMotor1cm / Abs(UpDownMotor1cm)) * (CurrentPosition / Abs(CurrentPosition)) * StanHeight * UpDownMotor1cm / 2))
            currentData.Z = Abs(1 / (MeasPos - CurrentPosition - (UpDownMotor1cm / Abs(UpDownMotor1cm)) * (CurrentPosition / Abs(CurrentPosition)) * StanHeight * UpDownMotor1cm / 2))
            Else
            currentData.X = currentData.X - FirstData.X
            currentData.Y = currentData.Y - FirstData.Y
            currentData.Z = currentData.Z - FirstData.Z
            End If
            txtUpDown = CurrentPosition
            txtUpDownSampleCm = Trim(Str((Abs(CurrentPosition) - Abs(StartPosition)) / UpDownMotor1cm))
        
            txtData(0) = Str$(currentData.X)
            txtData(1) = Str$(currentData.Y)
            txtData(2) = Str$(currentData.Z)
            
            WriteCalRodData WriteFile, CurrentPosition, Int(frmDCMotors.UpDownHeight), currentData
            
            For k = 1 To 4
        
                SampleMeasurement.Sample(k).X = currentData.X
                SampleMeasurement.Sample(k).Y = currentData.Y
                SampleMeasurement.Sample(k).Z = currentData.Z
            
            
            Next k
            
            Set currentData = SampleMeasurement.CorrectedSample(2)
            
            txtData(3) = Str$(currentData.inc)
            txtData(4) = Str$(currentData.dec)
        
            Set currentData = Nothing
        Else
            Warning = True
        End If
        
        If Abs(CurrentPosition) <= Abs(StartPosition) Then
        
            Warning = True
            
        End If
            
        
    Loop
    
    Warning = False
    frmDCMotors.TurningMotorRotate 0, False
    frmDCMotors.HomeToTop
    
    'Notify User that the sample is done
    frmSendMail.MailNotification "2G Status Update", _
                        "Long core sample done.  Sample saved to file: " & Me.txtFileName
                            
    
End Sub

Private Sub RunCalRodSeq(Xavg As Cartesian3D, Max As Cartesian3D)
    Dim Warning As Boolean
    Dim currentData As Cartesian3D
    Dim FirstData As Cartesian3D
    Dim SM As Cartesian3D
    Dim SXM As Cartesian3D
    Dim GaussianX As Variant
    Dim GaussianY As Variant
    Dim GaussianZ As Variant
    Dim RodPosition As Variant
    Dim CurrentPosition As Long
    Dim coarse As Long
    Dim fine As Long
    Dim i As Long
    Dim Delay As Double
    Dim StartPosition As Long
    Dim EndPosition As Long
    
    Warning = False
    If NOCOMM_MODE Then
    Delay = 0
    Threshold = 99
    currentPosInitialized = True
    Else
    Delay = 2
    End If
    i = 0
    CurrentPosition = ZeroPos - (ZeroPos / Abs(ZeroPos)) * (UpDownMotor1cm / Abs(UpDownMotor1cm)) * StanHeight * UpDownMotor1cm / 2
    cmdResume.Visible = True
    frmDCMotors.UpDownMove CurrentPosition, 2
    cmdResume.Visible = False
    Set currentData = New Cartesian3D
    Set FirstData = New Cartesian3D
    DelayTime (Delay)
    Set currentData = frmSQUID.getData
    FirstData.X = currentData.X
    FirstData.Y = currentData.Y
    FirstData.Z = currentData.Z
    txtUpDown = CurrentPosition
    txtData(0) = Str$(currentData.X - FirstData.X)
    txtData(1) = Str$(currentData.Y - FirstData.Y)
    txtData(2) = Str$(currentData.Z - FirstData.Z)
    CurrentPosition = ZeroPos - (ZeroPos / Abs(ZeroPos)) * (UpDownMotor1cm / Abs(UpDownMotor1cm)) * StanHeight * UpDownMotor1cm / 2 - (ZeroPos / Abs(ZeroPos)) * 10 * Spacing * UpDownMotor1cm * (UpDownMotor1cm / Abs(UpDownMotor1cm))
    Do While Not Warning
        If Not CurrentlyRunning Then Exit Sub
        CurrentPosition = CurrentPosition + (ZeroPos / Abs(ZeroPos)) * 10 * Spacing * UpDownMotor1cm * (UpDownMotor1cm / Abs(UpDownMotor1cm))
        If Abs(CurrentPosition) <= 1.15 * Abs(MeasPos) Then 'Try to don't hit the bottom...
            i = i + 1
            coarse = i
        Else
            Warning = True
        End If
    Loop
    Warning = False
    fine = Int(5 + StanHeight / Spacing)
    If fine * Spacing < 4 Then fine = 4 / Spacing '4 cm amplitude as a minimum for the fine scan
    If fine < 15 Then fine = 15  'zoom arround the maximum with 7 measurements on each side + repeat the max
    ReDim GaussianX(coarse + fine + 1)
    ReDim GaussianY(coarse + fine + 1)
    ReDim GaussianZ(coarse + fine + 1)
    ReDim RodPosition(coarse + fine + 1)
    Set SM = New Cartesian3D
    Set SXM = New Cartesian3D
    Set Xavg = New Cartesian3D
    Set Max = New Cartesian3D
    SM.X = 0
    SM.Y = 0
    SM.Z = 0
    SXM.X = 0
    SXM.Y = 0
    SXM.Z = 0
    Max.X = 0
    Max.Y = 0
    Max.Z = 0
    CurrentPosition = ZeroPos - (ZeroPos / Abs(ZeroPos)) * (UpDownMotor1cm / Abs(UpDownMotor1cm)) * StanHeight * UpDownMotor1cm / 2 - (ZeroPos / Abs(ZeroPos)) * 10 * Spacing * UpDownMotor1cm * (UpDownMotor1cm / Abs(UpDownMotor1cm))
    For i = 1 To coarse
        If Not CurrentlyRunning Then Exit Sub
        CurrentPosition = CurrentPosition + (ZeroPos / Abs(ZeroPos)) * 10 * Spacing * UpDownMotor1cm * (UpDownMotor1cm / Abs(UpDownMotor1cm))
        cmdResume.Visible = True
        frmDCMotors.UpDownMove CurrentPosition, 0
        cmdResume.Visible = False
        DelayTime (Delay)
        Set currentData = frmSQUID.getData
        If NOCOMM_MODE Then
            currentData.X = 0.9 * Abs(1 / (MeasPos - CurrentPosition - (UpDownMotor1cm / Abs(UpDownMotor1cm)) * (CurrentPosition / Abs(CurrentPosition)) * StanHeight * UpDownMotor1cm / 2))
            currentData.Y = 0.95 * Abs(1 / (MeasPos - CurrentPosition - (UpDownMotor1cm / Abs(UpDownMotor1cm)) * (CurrentPosition / Abs(CurrentPosition)) * StanHeight * UpDownMotor1cm / 2))
            currentData.Z = Abs(1 / (MeasPos - CurrentPosition - (UpDownMotor1cm / Abs(UpDownMotor1cm)) * (CurrentPosition / Abs(CurrentPosition)) * StanHeight * UpDownMotor1cm / 2))
            GaussianX(i) = currentData.X
            GaussianY(i) = currentData.Y
            GaussianZ(i) = currentData.Z
            RodPosition(i) = CurrentPosition
        Else
            currentData.X = currentData.X - FirstData.X
            currentData.Y = currentData.Y - FirstData.Y
            currentData.Z = currentData.Z - FirstData.Z
            GaussianX(i) = currentData.X
            GaussianY(i) = currentData.Y
            GaussianZ(i) = currentData.Z
            RodPosition(i) = frmDCMotors.UpDownHeight
        End If
        txtUpDown = CurrentPosition
        txtData(0) = Str$(currentData.X)
        txtData(1) = Str$(currentData.Y)
        txtData(2) = Str$(currentData.Z)
        WriteCalRodData WriteFile, CurrentPosition, Int(frmDCMotors.UpDownHeight), currentData
        If Abs(currentData.X) >= Max.X Then Max.X = Abs(currentData.X)
        If Abs(currentData.Y) >= Max.Y Then Max.Y = Abs(currentData.Y)
        If Abs(currentData.Z) >= Max.Z Then Max.Z = Abs(currentData.Z)
        Set currentData = Nothing
    Next i
    For i = 1 To coarse
        If Not CurrentlyRunning Then Exit Sub
        If Abs(GaussianX(i)) >= Max.X * (1 - Threshold / 100) Then
            SM.X = SM.X + GaussianX(i)
            SXM.X = SXM.X + GaussianX(i) * RodPosition(i) * (RodPosition(i) / Abs(RodPosition(i)))
            Xavg.X = Abs(SXM.X / SM.X) * (RodPosition(i) / Abs(RodPosition(i)))
        End If
        If Abs(GaussianY(i)) >= Max.Y * (1 - Threshold / 100) Then
            SM.Y = SM.Y + GaussianY(i)
            SXM.Y = SXM.Y + GaussianY(i) * RodPosition(i) * (RodPosition(i) / Abs(RodPosition(i)))
            Xavg.Y = Abs(SXM.Y / SM.Y) * (RodPosition(i) / Abs(RodPosition(i)))
        End If
        If Abs(GaussianZ(i)) >= Max.Z * (1 - Threshold / 100) Then
            SM.Z = SM.Z + GaussianZ(i)
            SXM.Z = SXM.Z + GaussianZ(i) * RodPosition(i) * (RodPosition(i) / Abs(RodPosition(i)))
            Xavg.Z = Abs(SXM.Z / SM.Z) * (RodPosition(i) / Abs(RodPosition(i)))
        End If
    Next i
    'MsgBox (Xavg.Z & " " & CurrentPosition)
    ' We can probably compare the three maxima depending of the importance of the moment SM of each
    If (Abs(Max.X) >= Abs(Max.Y)) And (Abs(Max.X) >= Abs(Max.Z)) Then CurrentPosition = Xavg.X + (Xavg.X / Abs(Xavg.X)) * (UpDownMotor1cm / Abs(UpDownMotor1cm)) * Int((fine + 1) / 2) * Spacing * UpDownMotor1cm
    If (Abs(Max.Y) >= Abs(Max.X)) And (Abs(Max.Y) >= Abs(Max.Z)) Then CurrentPosition = Xavg.Y + (Xavg.Y / Abs(Xavg.Y)) * (UpDownMotor1cm / Abs(UpDownMotor1cm)) * Int((fine + 1) / 2) * Spacing * UpDownMotor1cm
    If (Abs(Max.Z) >= Abs(Max.X)) And (Abs(Max.Z) >= Abs(Max.Y)) Then CurrentPosition = Xavg.Z + (Xavg.Z / Abs(Xavg.Z)) * (UpDownMotor1cm / Abs(UpDownMotor1cm)) * Int((fine + 1) / 2) * Spacing * UpDownMotor1cm
    If ((Abs(Max.X) >= Abs(Max.Y)) And (Abs(Max.X) >= Abs(Max.Z)) And (Abs(GaussianX(1)) = Max.X Or Abs(GaussianX(coarse)) = Max.X)) Or ((Abs(Max.Y) >= Abs(Max.X)) And (Abs(Max.Y) >= Abs(Max.Z)) And (Abs(GaussianY(1)) = Max.Y Or Abs(GaussianY(coarse)) = Max.Y)) Or ((Abs(Max.Z) >= Abs(Max.X)) And (Abs(Max.Z) >= Abs(Max.Y)) And (Abs(GaussianZ(1)) = Max.Z Or Abs(GaussianZ(coarse)) = Max.Z)) Then
        MsgBox ("Problem, can't locate even the coarse gaussian. Please correct manually the initial positions which are too far. " & _
        Chr(13) + "The file " & WriteFile & " recorded all the data." + Chr(13) + "DO NOT click OK in the settings window and wait for the program to close!")
        CurrentlyRunning = False
    Else
    cmdResume.Visible = True
    frmDCMotors.UpDownMove CurrentPosition, 2
    cmdResume.Visible = False
    DelayTime (Delay)
    SM.X = 0
    SM.Y = 0
    SM.Z = 0
    SXM.X = 0
    SXM.Y = 0
    SXM.Z = 0
    Max.X = 0
    Max.Y = 0
    Max.Z = 0
    CurrentPosition = CurrentPosition + (CurrentPosition / Abs(CurrentPosition)) * (UpDownMotor1cm / Abs(UpDownMotor1cm)) * Spacing * UpDownMotor1cm
    For i = coarse + 1 To coarse + fine + 1
        If Not CurrentlyRunning Then Exit Sub
        CurrentPosition = CurrentPosition - (CurrentPosition / Abs(CurrentPosition)) * (UpDownMotor1cm / Abs(UpDownMotor1cm)) * Spacing * UpDownMotor1cm
        If Abs(CurrentPosition) <= 1.15 * Abs(MeasPos) Then 'Try to don't hit the bottom...
            cmdResume.Visible = True
            frmDCMotors.UpDownMove CurrentPosition, 0
            cmdResume.Visible = False
            DelayTime (Delay)
            Set currentData = frmSQUID.getData
            If NOCOMM_MODE Then
            currentData.X = 0.9 * Abs(1 / (MeasPos - CurrentPosition - (UpDownMotor1cm / Abs(UpDownMotor1cm)) * (CurrentPosition / Abs(CurrentPosition)) * StanHeight * UpDownMotor1cm / 2))
            currentData.Y = 0.95 * Abs(1 / (MeasPos - CurrentPosition - (UpDownMotor1cm / Abs(UpDownMotor1cm)) * (CurrentPosition / Abs(CurrentPosition)) * StanHeight * UpDownMotor1cm / 2))
            currentData.Z = Abs(1 / (MeasPos - CurrentPosition - (UpDownMotor1cm / Abs(UpDownMotor1cm)) * (CurrentPosition / Abs(CurrentPosition)) * StanHeight * UpDownMotor1cm / 2))
            GaussianX(i) = currentData.X
            GaussianY(i) = currentData.Y
            GaussianZ(i) = currentData.Z
            RodPosition(i) = CurrentPosition
            Else
            currentData.X = currentData.X - FirstData.X
            currentData.Y = currentData.Y - FirstData.Y
            currentData.Z = currentData.Z - FirstData.Z
            GaussianX(i) = currentData.X
            GaussianY(i) = currentData.Y
            GaussianZ(i) = currentData.Z
            RodPosition(i) = frmDCMotors.UpDownHeight
            End If
            txtUpDown = CurrentPosition
            txtData(0) = Str$(currentData.X)
            txtData(1) = Str$(currentData.Y)
            txtData(2) = Str$(currentData.Z)
            WriteCalRodData WriteFile, CurrentPosition, Int(frmDCMotors.UpDownHeight), currentData
            If Abs(currentData.X) >= Max.X Then Max.X = Abs(currentData.X)
            If Abs(currentData.Y) >= Max.Y Then Max.Y = Abs(currentData.Y)
            If Abs(currentData.Z) >= Max.Z Then Max.Z = Abs(currentData.Z)
            Set currentData = Nothing
        End If
    Next i
    If ((Abs(Max.X) >= Abs(Max.Y)) And (Abs(Max.X) >= Abs(Max.Z)) And (Abs(GaussianX(coarse + 1)) = Max.X Or Abs(GaussianX(coarse + fine + 1)) = Max.X)) Or ((Abs(Max.Y) >= Abs(Max.X)) And (Abs(Max.Y) >= Abs(Max.Z)) And (Abs(GaussianY(coarse + 1)) = Max.Y Or Abs(GaussianY(coarse + fine + 1)) = Max.Y)) Or ((Abs(Max.Z) >= Abs(Max.X)) And (Abs(Max.Z) >= Abs(Max.Y)) And (Abs(GaussianZ(coarse + 1)) = Max.Z Or Abs(GaussianZ(coarse + fine + 1)) = Max.Z)) Then
        MsgBox ("Problem, fine gaussian non centered, restart or do it manually. " & GaussianY(coarse + 1) & " " & GaussianY(coarse + fine + 1) & " " & Max.Y & _
        Chr(13) + "The file " & WriteFile & " recorded all the data." + Chr(13) + "DO NOT click OK in the settings window and wait for the program to close!")
        CurrentlyRunning = False
    Else
    For i = 1 To coarse + fine + 1
        If Not CurrentlyRunning Then Exit Sub
        If Abs(GaussianX(i)) >= Max.X * (1 - Threshold / 100) Then
            SM.X = SM.X + GaussianX(i)
            SXM.X = SXM.X + GaussianX(i) * RodPosition(i) * (RodPosition(i) / Abs(RodPosition(i)))
            Xavg.X = Abs(SXM.X / SM.X) * (RodPosition(i) / Abs(RodPosition(i))) + StanHeight * (RodPosition(i) / Abs(RodPosition(i))) * (UpDownMotor1cm / Abs(UpDownMotor1cm)) * UpDownMotor1cm / 2
        End If
        If Abs(GaussianY(i)) >= Max.Y * (1 - Threshold / 100) Then
            SM.Y = SM.Y + GaussianY(i)
            SXM.Y = SXM.Y + GaussianY(i) * RodPosition(i) * (RodPosition(i) / Abs(RodPosition(i)))
            Xavg.Y = Abs(SXM.Y / SM.Y) * (RodPosition(i) / Abs(RodPosition(i))) + StanHeight * (RodPosition(i) / Abs(RodPosition(i))) * (UpDownMotor1cm / Abs(UpDownMotor1cm)) * UpDownMotor1cm / 2
        End If
        If Abs(GaussianZ(i)) >= Max.Z * (1 - Threshold / 100) Then
            SM.Z = SM.Z + GaussianZ(i)
            SXM.Z = SXM.Z + GaussianZ(i) * RodPosition(i) * (RodPosition(i) / Abs(RodPosition(i)))
            Xavg.Z = Abs(SXM.Z / SM.Z) * (RodPosition(i) / Abs(RodPosition(i))) + StanHeight * (RodPosition(i) / Abs(RodPosition(i))) * (UpDownMotor1cm / Abs(UpDownMotor1cm)) * UpDownMotor1cm / 2
        End If
    Next i
    'MsgBox (Xavg.Z & " " & CurrentPosition)
    ' We can probably compare the three maxima depending of the importance of the moment SM of each
    CurrentPosition = ZeroPos - (ZeroPos / Abs(ZeroPos)) * (UpDownMotor1cm / Abs(UpDownMotor1cm)) * StanHeight * UpDownMotor1cm / 2
    cmdResume.Visible = True
    frmDCMotors.UpDownMove CurrentPosition, 2
    cmdResume.Visible = False
    DelayTime (Delay)
    Set currentData = frmSQUID.getData
    If NOCOMM_MODE Then
    currentData.X = 0.9 * Abs(1 / (MeasPos - CurrentPosition - (UpDownMotor1cm / Abs(UpDownMotor1cm)) * (CurrentPosition / Abs(CurrentPosition)) * StanHeight * UpDownMotor1cm / 2))
    currentData.Y = 0.95 * Abs(1 / (MeasPos - CurrentPosition - (UpDownMotor1cm / Abs(UpDownMotor1cm)) * (CurrentPosition / Abs(CurrentPosition)) * StanHeight * UpDownMotor1cm / 2))
    currentData.Z = Abs(1 / (MeasPos - CurrentPosition - (UpDownMotor1cm / Abs(UpDownMotor1cm)) * (CurrentPosition / Abs(CurrentPosition)) * StanHeight * UpDownMotor1cm / 2))
    Else
    currentData.X = currentData.X - FirstData.X
    currentData.Y = currentData.Y - FirstData.Y
    currentData.Z = currentData.Z - FirstData.Z
    End If
    txtUpDown = CurrentPosition
    txtData(0) = Str$(currentData.X - FirstData.X)
    txtData(1) = Str$(currentData.Y - FirstData.Y)
    txtData(2) = Str$(currentData.Z - FirstData.Z)
    WriteCalRodData WriteFile, CurrentPosition, Int(frmDCMotors.UpDownHeight), currentData
    If (Abs(currentData.X) Or Abs(currentData.Y) Or Abs(currentData.Z)) > JumpThreshold Then
    cmdResume.Visible = True
    MotorUpDn_Move 0, 1
    frmDCMotors.ChangerMotortoHole 1
    frmDCMotors.SampleDropOff
    frmVacuum.ValveConnect (False)
    MotorUpDn_Move 0, 1
    frmDCMotors.ChangerMotortoHole 200
    cmdResume.Visible = False
    frmVacuum.MotorPower False
    MsgBox ("SQUID important drift or jump, restart the procedure after, but now wait for the program to close!")
    CurrentlyRunning = False
    Else
    Do While Not Warning
        If Not CurrentlyRunning Then Exit Sub
        CurrentPosition = CurrentPosition - (CurrentPosition / Abs(CurrentPosition)) * (UpDownMotor1cm / Abs(UpDownMotor1cm)) * 20 * Spacing * UpDownMotor1cm
        If Abs(CurrentPosition) >= 20 * Spacing * Abs(UpDownMotor1cm) Then 'To don't hit the switch...
            frmDCMotors.UpDownMove CurrentPosition, 1
            DelayTime (Delay * 0.1)
            Set currentData = frmSQUID.getData
            If NOCOMM_MODE Then
            currentData.X = 0.9 * Abs(1 / (MeasPos - CurrentPosition - (UpDownMotor1cm / Abs(UpDownMotor1cm)) * (CurrentPosition / Abs(CurrentPosition)) * StanHeight * UpDownMotor1cm / 2))
            currentData.Y = 0.95 * Abs(1 / (MeasPos - CurrentPosition - (UpDownMotor1cm / Abs(UpDownMotor1cm)) * (CurrentPosition / Abs(CurrentPosition)) * StanHeight * UpDownMotor1cm / 2))
            currentData.Z = Abs(1 / (MeasPos - CurrentPosition - (UpDownMotor1cm / Abs(UpDownMotor1cm)) * (CurrentPosition / Abs(CurrentPosition)) * StanHeight * UpDownMotor1cm / 2))
            Else
            currentData.X = currentData.X - FirstData.X
            currentData.Y = currentData.Y - FirstData.Y
            currentData.Z = currentData.Z - FirstData.Z
            End If
            txtUpDown = CurrentPosition
            txtData(0) = Str$(currentData.X)
            txtData(1) = Str$(currentData.Y)
            txtData(2) = Str$(currentData.Z)
            WriteCalRodData WriteFile, CurrentPosition, Int(frmDCMotors.UpDownHeight), currentData
            Set currentData = Nothing
        Else
            Warning = True
        End If
    Loop
    Warning = False
    If (Abs(SM.X) >= Abs(SM.Y)) And (Abs(SM.X) >= Abs(SM.Z)) Then
        SCoilPos = Int(val(Xavg.X)) - (MeasPos - SCoilPos)
        AFPos = Int(val(Xavg.X)) - (MeasPos - AFPos)
        IRMPos = Int(val(Xavg.X)) - (MeasPos - IRMPos)
        MeasPos = Int(val(Xavg.X))
    End If
    If (Abs(SM.Y) >= Abs(SM.X)) And (Abs(SM.Y) >= Abs(SM.Z)) Then
        SCoilPos = Int(val(Xavg.Y)) - (MeasPos - SCoilPos)
        AFPos = Int(val(Xavg.Y)) - (MeasPos - AFPos)
        IRMPos = Int(val(Xavg.Y)) - (MeasPos - IRMPos)
        MeasPos = Int(val(Xavg.Y))
    End If
    If (Abs(SM.Z) >= Abs(SM.X)) And (Abs(SM.Z) >= Abs(SM.Y)) Then
        SCoilPos = Int(val(Xavg.Z)) - (MeasPos - SCoilPos)
        AFPos = Int(val(Xavg.Z)) - (MeasPos - AFPos)
        IRMPos = Int(val(Xavg.Z)) - (MeasPos - IRMPos)
        MeasPos = Int(val(Xavg.Z))
    End If
    frmSettings.txtMeasPos = MeasPos
    frmSettings.txtAFPos = AFPos
    frmSettings.txtIRMPos = IRMPos
    If (ZeroPos / Abs(ZeroPos)) * (MeasPos - ZeroPos) < (15 - 1.5) * Abs(UpDownMotor1cm) Then
        MsgBox ("The Zero position has been changed from " & ZeroPos & " to " & MeasPos - 15 * (UpDownMotor1cm / Abs(UpDownMotor1cm)) * (MeasPos / Abs(MeasPos)) * UpDownMotor1cm)
        ZeroPos = MeasPos - 15 * (UpDownMotor1cm / Abs(UpDownMotor1cm)) * (MeasPos / Abs(MeasPos)) * UpDownMotor1cm
        frmSettings.txtZeroPos = ZeroPos
    Else
        If (ZeroPos / Abs(ZeroPos)) * (MeasPos - ZeroPos) > (15 + 1.5) * Abs(UpDownMotor1cm) Then
            MsgBox ("The Zero position has been changed from " & ZeroPos & " to " & MeasPos - 15 * (UpDownMotor1cm / Abs(UpDownMotor1cm)) * (MeasPos / Abs(MeasPos)) * UpDownMotor1cm)
            ZeroPos = MeasPos - 15 * (UpDownMotor1cm / Abs(UpDownMotor1cm)) * (MeasPos / Abs(MeasPos)) * UpDownMotor1cm
            frmSettings.txtZeroPos = ZeroPos
        End If
    End If
    End If
    End If
    End If
End Sub

Private Sub RunSusceSeq()
    Dim currentData As Cartesian3D
    Dim CurrentPosition As Long
    Dim fine As Long
    Dim SMX As Double
    Dim SXMX As Double
    Dim XavgX As Double
    Dim MaxX As Double
    Dim RodPosition As Variant
    Dim Susceptibility As Variant
    fine = Int(5 + StanHeight / (Spacing / 2))
    If fine * Spacing / 2 < 4 Then fine = 4 / (Spacing / 2) '4 cm amplitude as a minimum for the scan
    If fine < 21 Then fine = 21  'zoom arround the maximum with 10 measurements on each side
    ReDim RodPosition(fine)
    ReDim Susceptibility(fine)
    i = 1
    CurrentPosition = Int(SCoilPos - (SCoilPos / Abs(SCoilPos)) * (UpDownMotor1cm / Abs(UpDownMotor1cm)) * StanHeight * UpDownMotor1cm / 2)
    cmdResume.Visible = True
    frmDCMotors.HomeToTop
    frmSusceptibilityMeter.Zero
    frmDCMotors.UpDownMove CurrentPosition, 1
    cmdResume.Visible = False
    Susceptibility(i) = frmSusceptibilityMeter.Measure * SusceptibilityMomentFactorCGS
    Set currentData = New Cartesian3D
    currentData.X = Susce * (1 - Threshold / 100)
    currentData.Y = Susceptibility(i)
    currentData.Z = Susce * (1 + Threshold / 100)
    txtUpDown = CurrentPosition
    txtData(0) = Str$(currentData.X)
    txtData(1) = Str$(currentData.Y)
    txtData(2) = Str$(currentData.Z)
    WriteCalRodData WriteFile, CurrentPosition, Int(frmDCMotors.UpDownHeight), currentData
    i = 2
    CurrentPosition = Int(val(frmSettings.txtSCoilPos) - (val(frmSettings.txtSCoilPos) / Abs(val(frmSettings.txtSCoilPos))) * (UpDownMotor1cm / Abs(UpDownMotor1cm)) * StanHeight * UpDownMotor1cm / 2)
    cmdResume.Visible = True
    frmDCMotors.HomeToTop
    frmSusceptibilityMeter.Zero
    frmDCMotors.UpDownMove CurrentPosition, 1
    cmdResume.Visible = False
    Susceptibility(i) = frmSusceptibilityMeter.Measure * SusceptibilityMomentFactorCGS
    Set currentData = New Cartesian3D
    currentData.X = Susce * (1 - Threshold / 100)
    currentData.Y = Susceptibility(i)
    currentData.Z = Susce * (1 + Threshold / 100)
    txtUpDown = CurrentPosition
    txtData(0) = Str$(currentData.X)
    txtData(1) = Str$(currentData.Y)
    txtData(2) = Str$(currentData.Z)
    WriteCalRodData WriteFile, CurrentPosition, Int(frmDCMotors.UpDownHeight), currentData
    ' need to scan around the best S Coil position to find the good position
    If Susceptibility(1) > Susceptibility(2) Then
        CurrentPosition = SCoilPos - (SCoilPos / Abs(SCoilPos)) * (UpDownMotor1cm / Abs(UpDownMotor1cm)) * StanHeight * UpDownMotor1cm / 2 - Int((fine - 1) / 2) * (Spacing / 2) * UpDownMotor1cm * (SCoilPos / Abs(SCoilPos)) * (UpDownMotor1cm / Abs(UpDownMotor1cm))
        frmSettings.txtSCoilPos = SCoilPos
    Else
        CurrentPosition = val(frmSettings.txtSCoilPos) - (val(frmSettings.txtSCoilPos) / Abs(val(frmSettings.txtSCoilPos))) * (UpDownMotor1cm / Abs(UpDownMotor1cm)) * StanHeight * UpDownMotor1cm / 2 - (val(frmSettings.txtSCoilPos) / Abs(val(frmSettings.txtSCoilPos))) * Int((fine - 1) / 2) * (Spacing / 2) * UpDownMotor1cm * (UpDownMotor1cm / Abs(UpDownMotor1cm))
        'If Not NOCOMM_MODE Then MsgBox ("The previous position of the S Coil was better than the new one, we will scan around the old position. Please check very carefully the AF and IRM new positions.")
        AFPos = val(frmSettings.txtAFPos) + val(frmSettings.txtSCoilPos) - SCoilPos
        IRMPos = val(frmSettings.txtIRMPos) + val(frmSettings.txtSCoilPos) - SCoilPos
        SCoilPos = val(frmSettings.txtSCoilPos)
        frmSettings.txtAFPos = AFPos
        frmSettings.txtIRMPos = IRMPos
    End If
    CurrentPosition = CurrentPosition - (CurrentPosition / Abs(CurrentPosition)) * (UpDownMotor1cm / Abs(UpDownMotor1cm)) * (Spacing / 2) * UpDownMotor1cm
    SMX = 0
    SXMX = 0
    MaxX = 0
    For i = 1 To fine
        If Not CurrentlyRunning Then Exit Sub
        CurrentPosition = CurrentPosition + (CurrentPosition / Abs(CurrentPosition)) * (UpDownMotor1cm / Abs(UpDownMotor1cm)) * (Spacing / 2) * UpDownMotor1cm
        If Abs(CurrentPosition) >= Spacing * Abs(UpDownMotor1cm) Then 'To don't hit the switch...
            cmdResume.Visible = True
            frmDCMotors.HomeToTop
            frmSusceptibilityMeter.Zero
            frmDCMotors.UpDownMove CurrentPosition, 1
            cmdResume.Visible = False
            If NOCOMM_MODE Then
            Susceptibility(i) = Abs(1 / (SCoilPos - CurrentPosition - (CurrentPosition / Abs(CurrentPosition)) * (UpDownMotor1cm / Abs(UpDownMotor1cm)) * StanHeight * UpDownMotor1cm / 2))
            RodPosition(i) = CurrentPosition
            Else
            Susceptibility(i) = frmSusceptibilityMeter.Measure * SusceptibilityMomentFactorCGS
            RodPosition(i) = frmDCMotors.UpDownHeight
            End If
            Set currentData = New Cartesian3D
            currentData.X = Susce * (1 - Threshold / 100)
            currentData.Y = Susceptibility(i)
            currentData.Z = Susce * (1 + Threshold / 100)
            txtUpDown = CurrentPosition
            txtData(0) = Str$(currentData.X)
            txtData(1) = Str$(currentData.Y)
            txtData(2) = Str$(currentData.Z)
            WriteCalRodData WriteFile, CurrentPosition, Int(frmDCMotors.UpDownHeight), currentData
            If Abs(Susceptibility(i)) >= MaxX Then MaxX = Abs(Susceptibility(i))
            Set currentData = Nothing
        End If
    Next i
    If Abs(Susceptibility(1)) = MaxX Or Abs(Susceptibility(fine)) = MaxX Then
        MsgBox ("Problem, can't locate the susceptibility gaussian. Please correct manually the initial S Coil position which is too far. " & _
        Chr(13) + "The file " & WriteFile & " recorded all the data." + Chr(13) + "We will interpolate the AF, IRM and S Coil positions from the measurement region.")
        Exit Sub
    Else
    For i = 1 To fine
        If Not CurrentlyRunning Then Exit Sub
        If Susceptibility(i) >= MaxX * (1 - Threshold / 100) Then
            SMX = SMX + Susceptibility(i)
            SXMX = SXMX + Susceptibility(i) * RodPosition(i) * (RodPosition(i) / Abs(RodPosition(i)))
            XavgX = Abs(SXMX / SMX) * (RodPosition(i) / Abs(RodPosition(i))) + StanHeight * (RodPosition(i) / Abs(RodPosition(i))) * (UpDownMotor1cm / Abs(UpDownMotor1cm)) * UpDownMotor1cm / 2
        End If
    Next i
    i = 1
    CurrentPosition = Int(XavgX - StanHeight * (RodPosition(i) / Abs(RodPosition(i))) * (UpDownMotor1cm / Abs(UpDownMotor1cm)) * UpDownMotor1cm / 2)
    cmdResume.Visible = True
    frmDCMotors.HomeToTop
    frmSusceptibilityMeter.Zero
    frmDCMotors.UpDownMove CurrentPosition, 1
    cmdResume.Visible = False
    If NOCOMM_MODE Then
        Susceptibility(i) = Abs(1 / (SCoilPos - CurrentPosition - (UpDownMotor1cm / Abs(UpDownMotor1cm)) * (CurrentPosition / Abs(CurrentPosition)) * StanHeight * UpDownMotor1cm / 2))
        RodPosition(i) = CurrentPosition
    Else
        Susceptibility(i) = frmSusceptibilityMeter.Measure * SusceptibilityMomentFactorCGS
        RodPosition(i) = frmDCMotors.UpDownHeight
    End If
    'MsgBox (XavgX & " " & CurrentPosition)
    Set currentData = New Cartesian3D
    currentData.X = Susce * (1 - Threshold / 100)
    currentData.Y = Susceptibility(i)
    currentData.Z = Susce * (1 + Threshold / 100)
    txtUpDown = CurrentPosition
    txtData(0) = Str$(currentData.X)
    txtData(1) = Str$(currentData.Y)
    txtData(2) = Str$(currentData.Z)
    WriteCalRodData WriteFile, CurrentPosition, Int(frmDCMotors.UpDownHeight), currentData
    If Susceptibility(i) < Susce * (1 + Threshold / 100) And Susceptibility(i) > Susce * (1 - Threshold / 100) Then
        AFPos = Int(val(XavgX)) + (AFPos - SCoilPos)
        IRMPos = Int(val(XavgX)) + (IRMPos - SCoilPos)
        SCoilPos = Int(val(XavgX))
    Else
        If Susceptibility(i) * 0.00001 / SusceptibilityMomentFactorCGS < Susce * (1 + Threshold / 100) And Susceptibility(i) * 0.00001 / SusceptibilityMomentFactorCGS > Susce * (1 - Threshold / 100) Then
            SusceptibilityMomentFactorCGS = 0.00001
        Else
            SusceptibilityMomentFactorCGS = 0.00001 * Susce / Susceptibility(i)
        End If
        AFPos = Int(val(XavgX)) + (AFPos - SCoilPos)
        IRMPos = Int(val(XavgX)) + (IRMPos - SCoilPos)
        SCoilPos = Int(val(XavgX))
        If Not NOCOMM_MODE Then MsgBox ("In the Magnetometer tab, we needed to change the Moment Susceptibility Calibration Factor (CGS) from " & val(frmSettings.txtSusceptibilityMomentFactorCGS) & " to " & SusceptibilityMomentFactorCGS)
        frmSettings.txtSusceptibilityMomentFactorCGS = SusceptibilityMomentFactorCGS
    End If
    End If
    frmSettings.txtAFPos = AFPos
    frmSettings.txtIRMPos = IRMPos
    frmSettings.txtSCoilPos = SCoilPos
End Sub

Private Sub WriteCalRodDataHeaders(filename As String)
    Dim filenum As Long
    filenum = FreeFile
    On Error GoTo oops
    Open filename For Append As #filenum
    Print #filenum, "To find the rod altitudes you need to add the half height of the standard ," & Format$((ZeroPos / Abs(ZeroPos)) * StanHeight * Abs(UpDownMotor1cm) / 2, "0.00") & ", to the altitudes below."
    Print #filenum, "Requested, Position, x, y, z"
    Close #filenum
    GoTo stillworking
oops:
    CurrentlyRunning = False
    MsgBox "Unable to write to " & filename & "! Stopping the calibration run."
stillworking:
End Sub

Private Sub WriteCalRodData(filename As String, reqpos As Long, realpos As Long, data As Cartesian3D)
    Dim filenum As Long
    filenum = FreeFile
    On Error GoTo oops
    Open filename For Append As #filenum
    With data
    Print #filenum, reqpos; ","; realpos; ","; .X; ","; .Y; ","; .Z
    End With
    Close #filenum
    GoTo stillworking
oops:
    CurrentlyRunning = False
    MsgBox "Unable to write to " & filename & "! Stopping the calibration run."
stillworking:
End Sub

Private Sub WriteCalRodDataEnd(filename As String, txt As String, pos As Long, data As Cartesian3D)
    Dim filenum As Long
    filenum = FreeFile
    On Error GoTo oops
    Open filename For Append As #filenum
    With data
    Print #filenum, txt; ","; pos; ","; .X; ","; .Y; ","; .Z
    End With
    Close #filenum
    GoTo stillworking
oops:
    CurrentlyRunning = False
    MsgBox "Unable to write to " & filename & "! Stopping the calibration run."
stillworking:
End Sub

Private Sub Form_Load()

    Dim i As Long

    txtStartPosition = ZeroPos - (ZeroPos / Abs(ZeroPos)) * 5000
    EndPosition = (30500 + CLng(Abs(val(txtCoreLength) * UpDownMotor1cm))) * MeasPos / Abs(MeasPos)
    txtCoreLength = "0"
    chkUchannel = Unchecked
    chkSingleScan = Unchecked
    
    
    txtUpDownSampleCm.Visible = False

    For i = 0 To txtData.Count - 1
    
        txtData(i).Visible = False
        
    Next i

    
    lblInc.Visible = False
    lblDec.Visible = False
    optOrientation(0) = True
    optOrientation(1) = False

End Sub

Private Sub optOrientation_Click(Index As Integer)

    If Index = 0 Then
    
        optOrientation(1) = Not optOrientation(0)
        
    Else
    
        optOrientation(0) = Not optOrientation(1)
        
    End If

End Sub

Private Sub txtCoreLength_Change()

    Dim EndPosition As Long
    Dim StartPosition As Long

    StartPosition = CLng(val(txtStartPosition))

    EndPosition = (30500 + CLng(Abs(val(txtCoreLength) * UpDownMotor1cm))) * MeasPos / Abs(MeasPos)
     
    If Abs(EndPosition) > 43000 Then
    
        EndPosition = MeasPos / Abs(MeasPos) * 43000

    End If
    
    
  
    lblEndPosition.Caption = "End Position:     " & Trim(Str(EndPosition)) & "  motor counts"

End Sub

