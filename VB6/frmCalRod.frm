VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmCalRod 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Long Scan / Rod Position Calibration"
   ClientHeight    =   7830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8055
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7830
   ScaleWidth      =   8055
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   435
      Left            =   3840
      TabIndex        =   69
      Top             =   7200
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtNumMeasTillNextZero 
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   60
      Top             =   5760
      Width           =   735
   End
   Begin VB.TextBox txtTimeRemaining 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   58
      Top             =   5280
      Width           =   1335
   End
   Begin VB.TextBox txtElapsedTime 
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   57
      Top             =   4800
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Caption         =   "EPR Tube Scan"
      Height          =   6015
      Left            =   4200
      TabIndex        =   39
      Top             =   600
      Width           =   3735
      Begin VB.TextBox txtNumberMeasurements 
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   68
         Top             =   5160
         Width           =   735
      End
      Begin VB.CheckBox chkDoZeroMeas 
         Caption         =   "Do Zero Measurements During scan"
         Height          =   195
         Left            =   240
         TabIndex        =   66
         ToolTipText     =   "If this is unclicked, Zero measurements will only be done at the start and end of the EPR scan."
         Top             =   4320
         Width           =   3135
      End
      Begin VB.TextBox txtEstimatedRunTime 
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   65
         Top             =   5520
         Width           =   1215
      End
      Begin VB.TextBox txtMeasurementsPerZero 
         Height          =   285
         Left            =   1440
         TabIndex        =   62
         Top             =   4680
         Width           =   735
      End
      Begin VB.CheckBox chkDoEPRTubeScan 
         Caption         =   "Do EPR Tube Scan"
         Height          =   255
         Left            =   240
         TabIndex        =   50
         Top             =   3240
         Width           =   1815
      End
      Begin VB.TextBox txtEPRBottomBuffer 
         Height          =   285
         Left            =   2640
         TabIndex        =   49
         Top             =   2760
         Width           =   855
      End
      Begin VB.TextBox txtEPRTopBuffer 
         Height          =   285
         Left            =   2640
         TabIndex        =   47
         Top             =   2280
         Width           =   855
      End
      Begin VB.TextBox txtEPRSampleBottom 
         Height          =   285
         Left            =   2640
         TabIndex        =   45
         Top             =   1680
         Width           =   855
      End
      Begin VB.TextBox txtEPRSampleTop 
         Height          =   285
         Left            =   2640
         TabIndex        =   43
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox txtEPRTubeLength 
         Height          =   285
         Left            =   2640
         TabIndex        =   41
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label16 
         Caption         =   "# of Measurements:"
         Height          =   255
         Left            =   240
         TabIndex        =   67
         Top             =   5160
         Width           =   1575
      End
      Begin VB.Label Label20 
         Caption         =   "Estimated Run Time:"
         Height          =   255
         Left            =   240
         TabIndex        =   64
         Top             =   5520
         Width           =   1575
      End
      Begin VB.Label Label19 
         Caption         =   "measurements"
         Height          =   255
         Left            =   2280
         TabIndex        =   63
         Top             =   4680
         Width           =   1095
      End
      Begin VB.Label Label18 
         Caption         =   "Do Zero Every: "
         Height          =   255
         Left            =   240
         TabIndex        =   61
         Top             =   4680
         Width           =   1215
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   3600
         Y1              =   4080
         Y2              =   4080
      End
      Begin VB.Label lblSampleLen 
         Caption         =   "Sample Length (cm):"
         Height          =   255
         Left            =   120
         TabIndex        =   54
         Top             =   3720
         Width           =   3495
      End
      Begin VB.Label Label15 
         Caption         =   "Scan zone below sample (cm):"
         Height          =   255
         Left            =   240
         TabIndex        =   48
         Top             =   2760
         Width           =   2295
      End
      Begin VB.Label Label14 
         Caption         =   "Scan zone above sample (cm):"
         Height          =   495
         Left            =   240
         TabIndex        =   46
         Top             =   2280
         Width           =   2295
      End
      Begin VB.Label Label13 
         Caption         =   "Bottom of Sample on EPR Tube (cm from tube bottom):"
         Height          =   495
         Left            =   240
         TabIndex        =   44
         Top             =   1560
         Width           =   2295
      End
      Begin VB.Label Label10 
         Caption         =   "Top of Sample on EPR Tube (cm from tube bottom):"
         Height          =   495
         Left            =   240
         TabIndex        =   42
         Top             =   960
         Width           =   2295
      End
      Begin VB.Label Label8 
         Caption         =   "Length of EPR tube beyond the Quartz glass rod (cm):"
         Height          =   495
         Left            =   240
         TabIndex        =   40
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Core Scan"
      Height          =   2535
      Left            =   120
      TabIndex        =   30
      Top             =   1200
      Width           =   3975
      Begin VB.OptionButton optOrientation 
         Caption         =   "Up"
         Height          =   255
         Index           =   0
         Left            =   2400
         TabIndex        =   53
         Top             =   1680
         Width           =   975
      End
      Begin VB.OptionButton optOrientation 
         Caption         =   "Down"
         Height          =   255
         Index           =   1
         Left            =   2400
         TabIndex        =   52
         Top             =   2040
         Width           =   975
      End
      Begin VB.CheckBox chkDoCoreScan 
         Caption         =   "Do Core Scan"
         Height          =   255
         Left            =   360
         TabIndex        =   38
         Top             =   1920
         Width           =   2415
      End
      Begin VB.TextBox txtStartPosition 
         Height          =   288
         Left            =   1680
         TabIndex        =   32
         Top             =   360
         Width           =   972
      End
      Begin VB.TextBox txtCoreLength 
         Height          =   288
         Left            =   1680
         TabIndex        =   31
         Text            =   "3"
         Top             =   1200
         Width           =   972
      End
      Begin VB.Label Label7 
         Caption         =   "Starting Up/Down Motor Position:"
         Height          =   375
         Left            =   240
         TabIndex        =   37
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label9 
         Caption         =   "motor counts"
         Height          =   255
         Left            =   2760
         TabIndex        =   36
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label11 
         Caption         =   "Core Piece Length:"
         Height          =   255
         Left            =   240
         TabIndex        =   35
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label12 
         Caption         =   "cm"
         Height          =   255
         Left            =   2760
         TabIndex        =   34
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label lblEndPosition 
         Caption         =   "End Position:"
         Height          =   255
         Left            =   240
         TabIndex        =   33
         Top             =   840
         Width           =   3255
      End
   End
   Begin VB.TextBox txtUpDownSampleCm 
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   29
      Top             =   6720
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtData 
      Height          =   285
      Index           =   4
      Left            =   7080
      TabIndex        =   26
      Top             =   7320
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtData 
      Height          =   285
      Index           =   3
      Left            =   5640
      TabIndex        =   25
      Top             =   7320
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CheckBox chkSingleScan 
      Caption         =   "Single Scan"
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   7440
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog dlgOpenCreateFile 
      Left            =   5880
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdBrowseForFile 
      Caption         =   "..."
      Height          =   255
      Left            =   3600
      TabIndex        =   23
      Top             =   3960
      Width           =   495
   End
   Begin VB.CheckBox chkUchannel 
      Caption         =   "U channel"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   7200
      Width           =   1095
   End
   Begin VB.TextBox txtStanHeight 
      Height          =   285
      Left            =   1920
      TabIndex        =   0
      Text            =   "2.415"
      Top             =   90
      Width           =   495
   End
   Begin VB.CheckBox chksusceptibility 
      Caption         =   "Kvol (x10-6 CGS):"
      Height          =   255
      Left            =   3120
      TabIndex        =   1
      Top             =   120
      Value           =   1  'Checked
      Width           =   1575
   End
   Begin VB.TextBox txtSusce 
      Height          =   285
      Left            =   4680
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
      Width           =   495
   End
   Begin VB.TextBox txtThreshold 
      Height          =   285
      Left            =   1920
      TabIndex        =   4
      Text            =   "10"
      Top             =   840
      Width           =   495
   End
   Begin VB.TextBox txtFileName 
      Height          =   645
      Left            =   1080
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Text            =   "CalibRod"
      Top             =   3960
      Width           =   2415
   End
   Begin VB.TextBox txtUpDown 
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   6330
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtData 
      Height          =   285
      Index           =   0
      Left            =   5040
      TabIndex        =   7
      Top             =   6810
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtData 
      Height          =   285
      Index           =   1
      Left            =   6120
      TabIndex        =   8
      Top             =   6810
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtData 
      Height          =   285
      Index           =   2
      Left            =   7200
      TabIndex        =   9
      Top             =   6810
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdStartStop 
      Caption         =   "Start"
      Default         =   -1  'True
      Height          =   435
      Left            =   1440
      TabIndex        =   10
      Top             =   7200
      Width           =   975
   End
   Begin VB.CommandButton cmdResume 
      Caption         =   "Resume"
      Height          =   435
      Left            =   2640
      TabIndex        =   11
      Top             =   7200
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   3360
      Y1              =   6180
      Y2              =   6180
   End
   Begin VB.Label lblNumMeasTillNextZero 
      Caption         =   "# Meas. Till Next Zero:"
      Height          =   255
      Left            =   120
      TabIndex        =   59
      Top             =   5760
      Width           =   1695
   End
   Begin VB.Label lblTimeRemaining 
      Caption         =   "Time Remaining:"
      Height          =   255
      Left            =   120
      TabIndex        =   56
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Label lblElapsedTime 
      Caption         =   "Elapsed Time:"
      Height          =   255
      Left            =   120
      TabIndex        =   55
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Label lblSamplePos 
      Caption         =   "Sample Pos (cm):"
      Height          =   375
      Left            =   120
      TabIndex        =   51
      Top             =   6720
      Width           =   1335
   End
   Begin VB.Label lblDec 
      Caption         =   "Dec"
      Height          =   255
      Left            =   6720
      TabIndex        =   28
      Top             =   7320
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblInc 
      Caption         =   "Inc"
      Height          =   255
      Left            =   5280
      TabIndex        =   27
      Top             =   7350
      Visible         =   0   'False
      Width           =   255
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
      Left            =   2520
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
      Height          =   255
      Left            =   2520
      TabIndex        =   15
      Top             =   480
      Width           =   255
   End
   Begin VB.Label Label6 
      Caption         =   "Tolerance (%):"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Write to file:"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   3960
      Width           =   975
   End
   Begin VB.Label UpDown 
      Caption         =   "Up/Down Motor Pos:"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   6360
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label X 
      Caption         =   "X"
      Height          =   255
      Left            =   4800
      TabIndex        =   19
      Top             =   6840
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Y 
      Caption         =   "Y"
      Height          =   255
      Left            =   5880
      TabIndex        =   20
      Top             =   6840
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Z 
      Caption         =   "Z"
      Height          =   255
      Left            =   6960
      TabIndex        =   21
      Top             =   6840
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
'      IMPORTANT NOTE in regards to the changes:  There are two values used in the RunUChannel that will be
'      different for each sample changer system:
'                                       The maximum end position for the scan, and the position of the floor
'                                       relative to the up/down arm of the sample changer
'      These values are stored in the two new global variables in modConfig:
'                                       MinUpDownPos = -30500 for the Shoemaker sample changer at Caltech
'                                       FloorPos = -43000 for the Shoemaker sample changer  "  "
'      If this is the first time you are running this code on your sample changer system, these
'      two values will need to be determined and added to your Paleomag.ini file BEFORE running the program
'
'      To determine the MinUpDownPos, you will need to move the Up/Down arm using the DC Motors
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
'      is your FloorPos
'
Dim CurrentlyRunning As Boolean
Dim Susce As Double
Dim StanHeight As Double
Dim Spacing As Double
Dim Threshold As Double
Dim WriteFile As String

'UChannel Module-Global Status Variables
Dim SampleLenCm As Double
Dim RunTimeSeconds As Long
Dim TimePerMeasurement As Double
Dim SetupTime As Double
Dim MeasStartTime As Long
Dim AvgMeasTime As Double
Dim DuringMeas As Boolean
Dim DuringMoveToZero As Boolean
Dim StartZeroMoveTime As Long
Dim AvgZeroMoveRate As Long
Dim MyZeroPos As Long
Dim StartPosition As Long

Private Sub chkDoZeroMeas_Click()

    txtSpacing_Change

End Sub

Private Sub chkSingleScan_Click()

    txtSpacing_Change

End Sub

Private Sub cmdBrowseForFile_Click()

    dlgOpenCreateFile.flags = cdlOFNHideReadOnly Or cdlOFNCreatePrompt
    dlgOpenCreateFile.DialogTitle = "Open or Create File for Long Core Measurement"
    dlgOpenCreateFile.ShowOpen
    
    txtFileName = dlgOpenCreateFile.filename

End Sub

Private Sub cmdClose_Click()

    'Don't allow the window to close if a calibration or Uchannel or EPR tube scan is running
    If cmdStartStop.Caption = "Stop" Then Exit Sub
    
    Me.Hide

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
        
        'No run running, can close the window now
        cmdClose.Enabled = True
    Else
        CurrentlyRunning = True
        StanHeight = Abs(val(txtStanHeight))
        Susce = val(txtSusce) * 0.00001
        Spacing = val(txtSpacing)
        Threshold = val(txtThreshold)
        WriteFile = txtFileName
        cmdStartStop.Caption = "Stop"
        
        'Do not allow the user to close the window during a calibration
        'or UChannel run
        cmdClose.Enabled = False
        
        Run
        cmdStartStop.Caption = "Start"
        CurrentlyRunning = False
        
        'Enable the Close button
        cmdClose.Enabled = True
        
    End If
End Sub

Private Sub DoZeroMeas(ByVal CurStep As Long, _
                       ByVal StartTime, _
                       ByRef SampleMeasurement As MeasurementBlock, _
                       Optional ByVal ReturnAfterZero As Boolean = False, _
                       Optional ByVal BeforeZeroPos As Long = -1)
                       
    Dim CurrentData As Cartesian3D
                       
    'Don't do a zero if the user wants to return to the before zero
    'position, but doesn't input a position to return to
    If BeforeZeroPos = -1 And ReturnAfterZero = True Then Exit Sub
                       
    'Update the time remaining
    UChannelUpdateTime DateDiff("s", StartTime, Now), _
                       CurStep, , _
                       True, _
                       0, _
                       1
    
    'Move to the zero position
    frmDCMotors.UpDownMove MyZeroPos, 0
    
    'Update time again to get the rate of movement captures
    UChannelUpdateTime DateDiff("s", StartTime, Now), _
                       CurStep, , _
                       True, _
                       0, _
                       1
    
    'Get initial zero
    Set CurrentData = frmSQUID.getData
        
    'Update the time remaining
    UChannelUpdateTime DateDiff("s", StartTime, Now), CurStep
        
    If NOCOMM_MODE Then
        CurrentData.X = 0.9 * Abs(1 / (MeasPos - MyZeroPos))
        CurrentData.Y = 0.95 * Abs(1 / (MeasPos - MyZeroPos))
        CurrentData.Z = Abs(1 / (MeasPos - MyZeroPos))
        
    Else
        CurrentData.X = CurrentData.X
        CurrentData.Y = CurrentData.Y
        CurrentData.Z = CurrentData.Z
        
    End If
    txtUpDown = Format(MyZeroPos, "#0.0##")
    txtUpDownSampleCm = Trim(Str((Abs(MyZeroPos) - Abs(StartPosition)) / UpDownMotor1cm))

    txtData(0) = Str$(CurrentData.X)
    txtData(1) = Str$(CurrentData.Y)
    txtData(2) = Str$(CurrentData.Z)
    
    WriteUChannelData WriteFile, _
                      MyZeroPos, _
                      DateDiff("s", StartTime, Now), _
                      Int(frmDCMotors.UpDownHeight), _
                      frmDCMotors.TurningMotorAngle, _
                      CurrentData
                      
    'Update the time remaining
    UChannelUpdateTime DateDiff("s", StartTime, Now), CurStep
    
    For k = 1 To 4
    
        SampleMeasurement.Sample(k).X = CurrentData.X
        SampleMeasurement.Sample(k).Y = CurrentData.Y
        SampleMeasurement.Sample(k).Z = CurrentData.Z
            
    Next k
    
    Set CurrentData = SampleMeasurement.CorrectedSample(2)
    
    txtData(3) = Str$(CurrentData.inc)
    txtData(4) = Str$(CurrentData.dec)
    
    If ReturnAfterZero = True Then
    
        'Update the time remaining
        UChannelUpdateTime DateDiff("s", StartTime, Now), _
                           CurStep, , _
                           True, _
                           0, _
                           1
    
        frmDCMotors.UpDownMove BeforeZeroPos, 0
                       
        'Update the time remaining
        'Again, determine the movement rate to adjust the average movement
        'rate
        UChannelUpdateTime DateDiff("s", StartTime, Now), _
                           CurStep, , _
                           True, _
                           0, _
                           1
                       
    End If
                       
End Sub

Private Sub Form_Load()

    Dim i As Long
    Dim temp As Integer
    

    'Show the close button
    Me.cmdClose.Visible = True

    '(March 11, 2011 - I Hilburn)
    'Update all the form's icons with the program Ico file
    SetFormIcon Me

    txtStartPosition = ZeroPos - (ZeroPos / Abs(ZeroPos)) * 5000
    temp = MeasPos / Abs(MeasPos)
    EndPosition = (30500 + CLng(Abs(val(txtCoreLength) * UpDownMotor1cm))) * temp 'MeasPos / Abs(MeasPos)
    txtCoreLength = "0"
    chkUchannel = Unchecked
    chkSingleScan = Unchecked
    
    'Uncheck the two check-boxes for doing an EPR or a Core scan
    Me.chkDoCoreScan.Value = Unchecked
    Me.chkDoEPRTubeScan.Value = Unchecked
    
    txtUpDownSampleCm.Visible = False
    Me.lblSamplePos.Visible = False

    For i = 0 To txtData.Count - 1
    
        txtData(i).Visible = False
        
    Next i

    'Set the initial value for TimePerMeasurement
    TimePerMeasurement = 6
    AvgMeasTime = -1
    MeasStartTime = -1
    DuringMeas = False
    DuringMoveToZero = False
    StartZeroMoveTime = -1
    
    txtSpacing_Change
    
    'Hide all of the Data & Time display labels and text-boxes
    lblInc.Visible = False
    lblDec.Visible = False
    X.Visible = False
    Y.Visible = False
    Z.Visible = False
    Me.UpDown.Visible = False
    Me.txtUpDown.Visible = False
    Me.lblSamplePos.Visible = False
    Me.txtUpDownSampleCm = False
    Me.lblElapsedTime.Visible = False
    Me.lblTimeRemaining.Visible = False
    Me.lblNumMeasTillNextZero.Visible = False
    Me.txtElapsedTime.Visible = False
    Me.txtTimeRemaining.Visible = False
    Me.txtNumMeasTillNextZero.Visible = False
    
    
    optOrientation(0) = True
    optOrientation(1) = False
    
End Sub

'Sub Form_QueryUnload
'
' Created: March 11, 2011
'  Author: I Hilburn
'
'  Reason:  To prevent the user from being able to close the Calibrate Rod
'           / UChannel run window while a run is in progress.
'
'  Inputs:
'   Cancel  -   Integer or Boolean indicating whether the Unload event
'               that was triggered should be canceled or not
' UnloadMode -  Manner in which the Unload event was triggered:
'                   1) By user
'                   2) By the code
'                   3) By the Windows Task Manager
'                   4) By the Paleomag program being closed
'
' Outputs:  None
'
' Effects:  If the Unload event was triggered by the user, the event handler
'           checks the caption of the Start/Stop button.  If the button caption
'           reads "Stop", then a run is in progress and Cancel is set to true,
'           preventing the Unload event from proceeding
'
'           If the Unload event was triggered in any other way, the event handler
'           sets Cancel = False, and the Unloading of frmCalRod will proceed without
'           interference.  This is to prevent issues when exiting the Paleomag program,
'           when restarting the computer with frmCalRod open, or trying to force quit
'           the Paleomag program from the task manager while frmCalRod is open.
'
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If UnloadMode = vbFormControlMenu And _
       Me.cmdStartStop.Caption = "Stop" _
    Then
    
        Cancel = True
        
        Exit Sub
        
    Else
    
        Cancel = False
        
    End If

End Sub

Private Function FormatSecondsToStr(ByVal TotalSeconds As Long) As String

    Dim NumDays As Long
    Dim NumHours As Long
    Dim NumMinutes As Long
    Dim NumSeconds As Long
    Dim Remainder As Long
    
    'Calculate how many days, hours, minutes, and seconds the total
    'time in seconds comes out to
    NumDays = Int((TotalSeconds / 3600) / 24)
    
    Remainder = TotalSeconds - (NumDays * 3600 * 24)
    
    NumHours = Int((Remainder / 3600))
    
    Remainder = Remainder - (NumHours * 3600)
    
    NumMinutes = Int(Remainder / 60)
    
    Remainder = Remainder - NumMinutes * 60
    
    NumSeconds = Remainder
    
    'Return the correctly formated time
    FormatSecondsToStr = Format(NumDays, "#00:") & _
                         Format(NumHours, "00:") & _
                         Format(NumMinutes, "00:") & _
                         Format(NumSeconds, "00")
    
End Function

Private Sub optOrientation_Click(Index As Integer)

    If Index = 0 Then
    
        optOrientation(1) = Not optOrientation(0)
        
    Else
    
        optOrientation(0) = Not optOrientation(1)
        
    End If

End Sub

Private Sub Run()
    Dim CurrentData As Cartesian3D
    Dim CurrentPosition As Long
    Dim Xavg As Cartesian3D
    Dim Max As Cartesian3D
    Dim j As Long
    Dim StartTime
    If CurrentlyRunning And chkUchannel.Value = Checked Then
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
    If CurrentlyRunning And chksusceptibility.Value = Checked Then RunSusceSeq
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
    Set CurrentData = New Cartesian3D
    CurrentData.X = 0
    CurrentData.Y = 0
    CurrentData.Z = 0
    j = ZeroPos
    WriteCalRodDataEnd WriteFile, "Zeroing position", j, CurrentData
    CurrentData.X = Max.X
    CurrentData.Y = Max.Y
    CurrentData.Z = Max.Z
    j = MeasPos
    WriteCalRodDataEnd WriteFile, "Measuring position", j, CurrentData
    WriteCalRodDataEnd WriteFile, "Three maximum", j, Xavg
    j = AFPos
    WriteCalRodDataEnd WriteFile, "AF position", j, CurrentData
    j = IRMPos
    WriteCalRodDataEnd WriteFile, "IRM position", j, CurrentData
    Set CurrentData = New Cartesian3D
    CurrentData.X = Susce * (1 - Threshold / 100)
    CurrentData.Y = Susce
    CurrentData.Z = Susce * (1 + Threshold / 100)
    j = SCoilPos
    WriteCalRodDataEnd WriteFile, "S Coil position", j, CurrentData
    j = SusceptibilityMomentFactorCGS / 0.000000001
    WriteCalRodDataEnd WriteFile, "S Coil correction factor x 10-9", j, CurrentData
    Set CurrentData = New Cartesian3D
    CurrentData.X = Format$((SampleTop - SampleBottom) / UpDownMotor1cm, "0.00")
    CurrentData.Y = SampleTop
    CurrentData.Z = SampleBottom
    j = SampleTop - SampleBottom
    WriteCalRodDataEnd WriteFile, "Default samp. height (in units)- (in cm)- top and bot.", j, CurrentData
    j = Threshold
    WriteCalRodDataEnd WriteFile, "Tolerance (%)", j, CurrentData
    If ((chksusceptibility.Value = Checked And COMPortSusceptibility < 1) Or NOCOMM_MODE) Then
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

Private Sub RunCalRodSeq(Xavg As Cartesian3D, Max As Cartesian3D)
    Dim Warning As Boolean
    Dim CurrentData As Cartesian3D
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
    Dim EndPosition As Long
    Dim StartTime
    
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
    Set CurrentData = New Cartesian3D
    Set FirstData = New Cartesian3D
    DelayTime (Delay)
    Set CurrentData = frmSQUID.getData
    FirstData.X = CurrentData.X
    FirstData.Y = CurrentData.Y
    FirstData.Z = CurrentData.Z
    txtUpDown = CurrentPosition
    txtData(0) = Str$(CurrentData.X - FirstData.X)
    txtData(1) = Str$(CurrentData.Y - FirstData.Y)
    txtData(2) = Str$(CurrentData.Z - FirstData.Z)
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
        Set CurrentData = frmSQUID.getData
        If NOCOMM_MODE Then
            CurrentData.X = 0.9 * Abs(1 / (MeasPos - CurrentPosition - (UpDownMotor1cm / Abs(UpDownMotor1cm)) * (CurrentPosition / Abs(CurrentPosition)) * StanHeight * UpDownMotor1cm / 2))
            CurrentData.Y = 0.95 * Abs(1 / (MeasPos - CurrentPosition - (UpDownMotor1cm / Abs(UpDownMotor1cm)) * (CurrentPosition / Abs(CurrentPosition)) * StanHeight * UpDownMotor1cm / 2))
            CurrentData.Z = Abs(1 / (MeasPos - CurrentPosition - (UpDownMotor1cm / Abs(UpDownMotor1cm)) * (CurrentPosition / Abs(CurrentPosition)) * StanHeight * UpDownMotor1cm / 2))
            GaussianX(i) = CurrentData.X
            GaussianY(i) = CurrentData.Y
            GaussianZ(i) = CurrentData.Z
            RodPosition(i) = CurrentPosition
        Else
            CurrentData.X = CurrentData.X - FirstData.X
            CurrentData.Y = CurrentData.Y - FirstData.Y
            CurrentData.Z = CurrentData.Z - FirstData.Z
            GaussianX(i) = CurrentData.X
            GaussianY(i) = CurrentData.Y
            GaussianZ(i) = CurrentData.Z
            RodPosition(i) = frmDCMotors.UpDownHeight
        End If
        txtUpDown = CurrentPosition
        txtData(0) = Str$(CurrentData.X)
        txtData(1) = Str$(CurrentData.Y)
        txtData(2) = Str$(CurrentData.Z)
        WriteCalRodData WriteFile, CurrentPosition, Int(frmDCMotors.UpDownHeight), CurrentData
        If Abs(CurrentData.X) >= Max.X Then Max.X = Abs(CurrentData.X)
        If Abs(CurrentData.Y) >= Max.Y Then Max.Y = Abs(CurrentData.Y)
        If Abs(CurrentData.Z) >= Max.Z Then Max.Z = Abs(CurrentData.Z)
        Set CurrentData = Nothing
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
            Set CurrentData = frmSQUID.getData
            If NOCOMM_MODE Then
            CurrentData.X = 0.9 * Abs(1 / (MeasPos - CurrentPosition - (UpDownMotor1cm / Abs(UpDownMotor1cm)) * (CurrentPosition / Abs(CurrentPosition)) * StanHeight * UpDownMotor1cm / 2))
            CurrentData.Y = 0.95 * Abs(1 / (MeasPos - CurrentPosition - (UpDownMotor1cm / Abs(UpDownMotor1cm)) * (CurrentPosition / Abs(CurrentPosition)) * StanHeight * UpDownMotor1cm / 2))
            CurrentData.Z = Abs(1 / (MeasPos - CurrentPosition - (UpDownMotor1cm / Abs(UpDownMotor1cm)) * (CurrentPosition / Abs(CurrentPosition)) * StanHeight * UpDownMotor1cm / 2))
            GaussianX(i) = CurrentData.X
            GaussianY(i) = CurrentData.Y
            GaussianZ(i) = CurrentData.Z
            RodPosition(i) = CurrentPosition
            Else
            CurrentData.X = CurrentData.X - FirstData.X
            CurrentData.Y = CurrentData.Y - FirstData.Y
            CurrentData.Z = CurrentData.Z - FirstData.Z
            GaussianX(i) = CurrentData.X
            GaussianY(i) = CurrentData.Y
            GaussianZ(i) = CurrentData.Z
            RodPosition(i) = frmDCMotors.UpDownHeight
            End If
            txtUpDown = CurrentPosition
            txtData(0) = Str$(CurrentData.X)
            txtData(1) = Str$(CurrentData.Y)
            txtData(2) = Str$(CurrentData.Z)
            WriteCalRodData WriteFile, CurrentPosition, Int(frmDCMotors.UpDownHeight), CurrentData
            If Abs(CurrentData.X) >= Max.X Then Max.X = Abs(CurrentData.X)
            If Abs(CurrentData.Y) >= Max.Y Then Max.Y = Abs(CurrentData.Y)
            If Abs(CurrentData.Z) >= Max.Z Then Max.Z = Abs(CurrentData.Z)
            Set CurrentData = Nothing
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
    Set CurrentData = frmSQUID.getData
    If NOCOMM_MODE Then
    CurrentData.X = 0.9 * Abs(1 / (MeasPos - CurrentPosition - (UpDownMotor1cm / Abs(UpDownMotor1cm)) * (CurrentPosition / Abs(CurrentPosition)) * StanHeight * UpDownMotor1cm / 2))
    CurrentData.Y = 0.95 * Abs(1 / (MeasPos - CurrentPosition - (UpDownMotor1cm / Abs(UpDownMotor1cm)) * (CurrentPosition / Abs(CurrentPosition)) * StanHeight * UpDownMotor1cm / 2))
    CurrentData.Z = Abs(1 / (MeasPos - CurrentPosition - (UpDownMotor1cm / Abs(UpDownMotor1cm)) * (CurrentPosition / Abs(CurrentPosition)) * StanHeight * UpDownMotor1cm / 2))
    Else
    CurrentData.X = CurrentData.X - FirstData.X
    CurrentData.Y = CurrentData.Y - FirstData.Y
    CurrentData.Z = CurrentData.Z - FirstData.Z
    End If
    txtUpDown = CurrentPosition
    txtData(0) = Str$(CurrentData.X - FirstData.X)
    txtData(1) = Str$(CurrentData.Y - FirstData.Y)
    txtData(2) = Str$(CurrentData.Z - FirstData.Z)
    WriteCalRodData WriteFile, CurrentPosition, Int(frmDCMotors.UpDownHeight), CurrentData
    If (Abs(CurrentData.X) Or Abs(CurrentData.Y) Or Abs(CurrentData.Z)) > JumpThreshold Then
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
            Set CurrentData = frmSQUID.getData
            If NOCOMM_MODE Then
            CurrentData.X = 0.9 * Abs(1 / (MeasPos - CurrentPosition - (UpDownMotor1cm / Abs(UpDownMotor1cm)) * (CurrentPosition / Abs(CurrentPosition)) * StanHeight * UpDownMotor1cm / 2))
            CurrentData.Y = 0.95 * Abs(1 / (MeasPos - CurrentPosition - (UpDownMotor1cm / Abs(UpDownMotor1cm)) * (CurrentPosition / Abs(CurrentPosition)) * StanHeight * UpDownMotor1cm / 2))
            CurrentData.Z = Abs(1 / (MeasPos - CurrentPosition - (UpDownMotor1cm / Abs(UpDownMotor1cm)) * (CurrentPosition / Abs(CurrentPosition)) * StanHeight * UpDownMotor1cm / 2))
            Else
            CurrentData.X = CurrentData.X - FirstData.X
            CurrentData.Y = CurrentData.Y - FirstData.Y
            CurrentData.Z = CurrentData.Z - FirstData.Z
            End If
            txtUpDown = CurrentPosition
            txtData(0) = Str$(CurrentData.X)
            txtData(1) = Str$(CurrentData.Y)
            txtData(2) = Str$(CurrentData.Z)
            WriteCalRodData WriteFile, CurrentPosition, Int(frmDCMotors.UpDownHeight), CurrentData
            Set CurrentData = Nothing
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
    If UseXYTableAPS Then
    frmDCMotors.ChangerMotortoHole 46
    Else
    frmDCMotors.ChangerMotortoHole 199
    End If
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
    If UseXYTableAPS Then
    frmDCMotors.ChangerMotortoHole 46
    Else
    frmDCMotors.ChangerMotortoHole 200
    End If
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

Private Sub RunSusceSeq()
    Dim CurrentData As Cartesian3D
    Dim CurrentPosition As Long
    Dim StartTime
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
    Set CurrentData = New Cartesian3D
    CurrentData.X = Susce * (1 - Threshold / 100)
    CurrentData.Y = Susceptibility(i)
    CurrentData.Z = Susce * (1 + Threshold / 100)
    txtUpDown = CurrentPosition
    txtData(0) = Str$(CurrentData.X)
    txtData(1) = Str$(CurrentData.Y)
    txtData(2) = Str$(CurrentData.Z)
    WriteCalRodData WriteFile, CurrentPosition, Int(frmDCMotors.UpDownHeight), CurrentData
    i = 2
    CurrentPosition = Int(val(frmSettings.txtSCoilPos) - (val(frmSettings.txtSCoilPos) / Abs(val(frmSettings.txtSCoilPos))) * (UpDownMotor1cm / Abs(UpDownMotor1cm)) * StanHeight * UpDownMotor1cm / 2)
    cmdResume.Visible = True
    frmDCMotors.HomeToTop
    frmSusceptibilityMeter.Zero
    frmDCMotors.UpDownMove CurrentPosition, 1
    cmdResume.Visible = False
    Susceptibility(i) = frmSusceptibilityMeter.Measure * SusceptibilityMomentFactorCGS
    Set CurrentData = New Cartesian3D
    CurrentData.X = Susce * (1 - Threshold / 100)
    CurrentData.Y = Susceptibility(i)
    CurrentData.Z = Susce * (1 + Threshold / 100)
    txtUpDown = CurrentPosition
    txtData(0) = Str$(CurrentData.X)
    txtData(1) = Str$(CurrentData.Y)
    txtData(2) = Str$(CurrentData.Z)
    WriteCalRodData WriteFile, CurrentPosition, Int(frmDCMotors.UpDownHeight), CurrentData
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
            Set CurrentData = New Cartesian3D
            CurrentData.X = Susce * (1 - Threshold / 100)
            CurrentData.Y = Susceptibility(i)
            CurrentData.Z = Susce * (1 + Threshold / 100)
            txtUpDown = CurrentPosition
            txtData(0) = Str$(CurrentData.X)
            txtData(1) = Str$(CurrentData.Y)
            txtData(2) = Str$(CurrentData.Z)
            WriteCalRodData WriteFile, CurrentPosition, Int(frmDCMotors.UpDownHeight), CurrentData
            If Abs(Susceptibility(i)) >= MaxX Then MaxX = Abs(Susceptibility(i))
            Set CurrentData = Nothing
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
    Set CurrentData = New Cartesian3D
    CurrentData.X = Susce * (1 - Threshold / 100)
    CurrentData.Y = Susceptibility(i)
    CurrentData.Z = Susce * (1 + Threshold / 100)
    txtUpDown = CurrentPosition
    txtData(0) = Str$(CurrentData.X)
    txtData(1) = Str$(CurrentData.Y)
    txtData(2) = Str$(CurrentData.Z)
    WriteCalRodData WriteFile, CurrentPosition, Int(frmDCMotors.UpDownHeight), CurrentData
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

Private Sub RunUchannel()
    Dim Warning As Boolean
    Dim CurrentData As Cartesian3D
    Dim FirstData As Cartesian3D
    Dim SM As Cartesian3D
    Dim SXM As Cartesian3D
    Dim GaussianX As Variant
    Dim GaussianY As Variant
    Dim GaussianZ As Variant
    Dim RodPosition As Variant
    Dim CurrentPosition As Long
    Dim StartPosInCm As Double
    Dim CurPosInCm As Double
    Dim coarse As Long
    Dim fine As Long
    Dim i As Long
    Dim Delay As Double
    Dim EndPosition As Long
    Dim SampleLength As Long
    Dim SampleMeasurement As MeasurementBlock
    Dim ZeroEvery As Long
    
    Dim EPRTubeLen As Long
    Dim ToSampleTopLen As Long
    Dim BufferAboveLen As Long
    Dim BufferBelowLen As Long
    
    Dim StartTime
    Dim ElapsedTime As Long
    
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
    
    'Set the initial values for recalculating TimePerMeasurement
    'during the scan
    TimePerMeasurement = 6
    AvgMeasTime = -1
    MeasStartTime = -1
    DuringMeas = False
    DuringMoveToZero = False
    StartZeroMoveTime = -1
    AvgZeroMoveRate = -1
    
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
    
    Set CurrentData = New Cartesian3D
    Set FirstData = New Cartesian3D
    frmSQUID.CLP "A"
    frmSQUID.ResetCount "A"
    frmProgram.StatusBar "Resetting...", 3
    DelayTime (Delay)
    Set CurrentData = frmSQUID.getData
    FirstData.X = CurrentData.X
    FirstData.Y = CurrentData.Y
    FirstData.Z = CurrentData.Z

    
    SampleMeasurement.isUp = optOrientation(0)
        
    txtUpDown.Visible = True
    txtUpDown = CurrentPosition
    Me.txtElapsedTime.Visible = True
    Me.txtElapsedTime = "00:00:00:00"
    Me.txtTimeRemaining.Visible = True
    Me.txtTimeRemaining = Me.txtEstimatedRunTime
        
    txtData(0) = Str$(CurrentData.X)
    txtData(1) = Str$(CurrentData.Y)
    txtData(2) = Str$(CurrentData.Z)
    txtData(0).Visible = True
    txtData(1).Visible = True
    txtData(2).Visible = True
    txtData(3).Visible = True
    txtData(4).Visible = True
    txtData(3) = ""
    txtData(4) = ""
    
    txtUpDownSampleCm.Visible = True
    txtUpDownSampleCm = ""
    
    lblInc.Visible = True
    lblDec.Visible = True
    UpDown.Visible = True
    X.Visible = True
    Y.Visible = True
    Z.Visible = True
    Me.lblSamplePos.Visible = True
    Me.lblElapsedTime.Visible = True
    Me.lblTimeRemaining.Visible = True
    
    
    '-------------------------------------------------
    '   Added August 2, 2010 - I. Hilburn
    '
    '   For doing EPR tube long-scans
    '-------------------------------------------------
    
    'Check to see if both checks are selected
    If Me.chkDoCoreScan.Value = Checked And _
       Me.chkDoEPRTubeScan.Value = Checked _
    Then
    
        'Can only have one of the two selected
        'pop-up a message and exit the sub
        MsgBox "Both Core Scan and EPR Tube scan have been checked." & vbNewLine & _
               "Only one of these two can be done at a time.  Please uncheck " & _
               "one of them.", , _
               "ERROR!"
               
        Exit Sub
        
    End If
    
    'check to see if EPR Tube scan is checked
    If Me.chkDoEPRTubeScan.Value = Checked Then
    
        'Need to turn on the vacuum no if it isn't already on
        If frmVacuum.VacuumMotorOn = False Then frmVacuum.MotorPower True
        If frmVacuum.VacuumConnectOn = False Then frmVacuum.ValveConnect True
    
        'Prompt the User to load the sample
        MsgBox "Please load the EPR tube sample."
    
        'Need to load necessary values (start & end positions)
        'for the EPR Tube scan
        EPRTubeLen = CLng(Abs(val(Me.txtEPRTubeLength) * UpDownMotor1cm))
        BufferAboveLen = CLng(Abs(val(Me.txtEPRTopBuffer) * UpDownMotor1cm))
        BufferBelowLen = CLng(Abs(val(Me.txtEPRBottomBuffer) * UpDownMotor1cm))
        
        'Get the sample length
        SampleLength = CLng(Abs((val(Me.txtEPRSampleTop) - val(Me.txtEPRSampleBottom)) * _
                                 UpDownMotor1cm))
        
        'Get the length of the EPR tube up to the sample top
        ToSampleTopLen = CLng(Abs((val(Me.txtEPRTubeLength) - val(Me.txtEPRSampleTop)) * _
                                  UpDownMotor1cm))
                                          
        'Store the number of measurements in between each zero
        ZeroEvery = val(Me.txtMeasurementsPerZero)
                                          
        'Write the data file headers
        WriteUChannelHeaders WriteFile, _
                             Round(Abs(SampleLength / UpDownMotor1cm), 1), _
                             Round(Abs((BufferAboveLen + BufferBelowLen + SampleLength) / UpDownMotor1cm), 1), _
                             val(Me.txtEPRSampleTop), _
                             val(Me.txtEPRSampleBottom), _
                             val(Me.txtEPRTubeLength)
                                          
        StartPosition = MeasPos - (Sgn(MeasPos) * _
                                   (ToSampleTopLen + SampleLength + BufferBelowLen))
                                   
        'Get the Zero Motor Position that the sample is going to be moved to
        MyZeroPos = MeasPos - (Sgn(MeasPos) * _
                               CLng((Abs((val(Me.txtEPRTubeLength) - val(Me.txtEPRSampleTop)) * _
                                    UpDownMotor1cm) + _
                               SampleLenCm * Abs(UpDownMotor1cm) + _
                               Abs((val(Me.txtEPRBottomBuffer) + 5) * UpDownMotor1cm))))
                                   
        Me.txtStartPosition = Trim(Str(StartPosition))
                   
        EndPosition = MeasPos - (Sgn(MeasPos) * _
                                 (ToSampleTopLen - BufferAboveLen))
                                 
        Me.lblEndPosition = "End Position:   " & Trim(Str(EndPosition))
                   
    ElseIf Me.chkDoCoreScan.Value = Checked Then
                    
        MsgBox ("Please load the Long-Core sample.")
                    
        '-------------------------------------------------
        '   Added October 8, 2009 - I. Hilburn
        '-------------------------------------------------
        SampleLength = CLng(val(txtCoreLength) * UpDownMotor1cm)
            
            
        If SampleLength < 0 Or SampleLength > Abs(MinUpDownPos) Then
        
            'User has entered a bad / invalid sample length, pop up an error and exit the sub-routine
            MsgBox "User entered a core sample length: " & vbNewLine & _
                    txtCoreLength & " cm" & vbNewLine & _
                    Trim(Str(SampleLength)) & " motor counts", , "Bad Core Length!"
                    
            Exit Sub
            
        End If
            
       
        StartPosition = CLng(val(txtStartPosition))
        
        If Abs(StartPosition) < 0 Or Abs(StartPosition) > Abs(MinUpDownPos) Then
        
            'User has selected a crappy start position, send an error and exit the subroutine.
            MsgBox "Bad UChannel scan start motor count position:" & vbNewLine & _
                    txtStartPosition, , "Bad Start Position!"
        
            Exit Sub
            
        End If
        
        'Adjust starting position based on the sample length
        StartPosition = StartPosition - Sgn(MeasPos) * SampleLength
        
        
        'Adjust End Position based on the sample length and already adjusted start position
        EndPosition = (MinUpDownPos * Sgn(MinUpDownPos) + Abs(SampleLength)) * Sgn(MeasPos)
        
        If Abs(EndPosition) > FloorPos * Sgn(FloorPos) Then  'FloorPos will be different for other sample changer systems
        
            MsgBox "Sample is too long - it will smack into the floor." & vbNewLine & vbNewLine & _
                    "You will need to measure the sample in the opposite orientation to get full" & _
                    "coverage."
        
            EndPosition = FloorPos * Sgn(FloorPos)  'FloorPos will be different for other sample changer systems
    
        End If
        
        EndPosition = EndPosition - Abs(SampleLength) * Sgn(MeasPos)
        lblEndPosition.Caption = "End Position:     " & Trim(Str(EndPosition)) & "  motor counts"
        
        'Write in Data file headers
        WriteUChannelHeaders WriteFile, _
                             Round(Abs(SampleLength / UpDownMotor1cm), 1), _
                             Round(Abs(EndPosition - StartPosition) / Abs(UpDownMotor1cm), 1)
                                     
    Else
    
        'Neither is checked!
        
        MsgBox "Neither Do Core Scan or Do EPR Tube scan is checked." & vbNewLine & _
               "One of these two scan types must be checked before a scan can be run.", , _
               "ERROR!"
               
        Exit Sub
        
    End If
        
    '--------------------------------------------
    CurrentPosition = 0
    
    txtUpDownSampleCm = Trim(Str((Abs(CurrentPosition) - Abs(StartPosition)) / UpDownMotor1cm))
    
    '(Mar 2010, I Hilburn) - changing it again so that the position is incremented
    'in centimeters and then translated for each movement of the updown arm from centimeters
    'to motor counts to prevent position drift due to cm to motor count conversion rounding errors
    StartPosInCm = StartPosition * Sgn(StartPosition) _
                    / UpDownMotor1cm * Sgn(UpDownMotor1cm)
        
    'Initialize i (movement counter), to zero
    i = 0
    
    'Before Scan starts, set startTime to the current system time
    StartTime = Now
            
    Do While Not Warning
        If Not CurrentlyRunning Then Exit Sub
                        
        '(Mar, 2010 I Hilburn) - Calculate current position in Centimeters from the start position
        CurPosInCm = StartPosInCm * Sgn(StartPosInCm) _
                        + i * Spacing * Sgn(Spacing)
        
        '(Mar, 2010 I Hilburn) Then translate this current position to a motor position
        'This eliminates fractional drift do to the fact that Spacing * UpDownMotor1cm is not an integer value
        CurrentPosition = (Sgn(ZeroPos)) * CurPosInCm * UpDownMotor1cm * (Sgn(UpDownMotor1cm))
        
        If Abs(CurrentPosition) <= Abs(EndPosition) Then  '(Mar 2010, I. Hilburn) - End Position now set to MinUpDownPos which is recorded in the Paleomag.ini file and can be changed in the frmSettings window
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
    
    frmDCMotors.TurningMotorRotate 0, False
    DelayTime (Delay)
    
    DoZeroMeas 0, StartTime, SampleMeasurement
    
    Set CurrentData = Nothing
        
    For i = 0 To coarse - 1
        If Not CurrentlyRunning Then Exit Sub
                
        '(Mar, 2010 I Hilburn) - Calculate current position in Centimeters from the start position
        CurPosInCm = StartPosInCm * Sgn(StartPosInCm) _
                        + i * Spacing * Sgn(Spacing)
        
        '(Mar, 2010 I Hilburn) Then translate this current position to a motor position
        'This eliminates fractional drift do to the fact that Spacing * UpDownMotor1cm is not an integer value
        CurrentPosition = (Sgn(ZeroPos)) * CurPosInCm * UpDownMotor1cm * (Sgn(UpDownMotor1cm))
        
        If Me.chkDoZeroMeas.Value = Checked Then
        
            'Check to see if need to do a zero measurement this step
            If i Mod ZeroEvery = 0 Then
        
                'Update the time remaining,
                'Don't add the time for this measurement to average
                UChannelUpdateTime DateDiff("s", StartTime, Now), _
                                   i + 1
                                   
                DoZeroMeas i + 1, StartTime, SampleMeasurement, True, CurrentPosition
                
                'Set DuringMeas = False, don't want to capture the time taken for this measurement
                DuringMeas = False
                                   
            Else
            
                'This measurement is normal, can be used to calculate the time
                'per measurement
                UChannelUpdateTime DateDiff("s", StartTime, Now), _
                                   i + 1, _
                                   True
                                   
            End If
            
        Else
        
            'This measurement is normal, can be used to calculate the time
            'per measurement
            UChannelUpdateTime DateDiff("s", StartTime, Now), _
                               i + 1, _
                               True
                               
        End If
        
        frmDCMotors.UpDownMove CurrentPosition, 0
        
        'Update the time remaining
        UChannelUpdateTime DateDiff("s", StartTime, Now), _
                           i + 1
                
        DelayTime (Delay)
        Set CurrentData = frmSQUID.getData
        
        'Update the time remaining
        UChannelUpdateTime DateDiff("s", StartTime, Now), _
                           i + 1
    
        If NOCOMM_MODE Then
            CurrentData.X = 0.9 * Abs(1 / (MeasPos - CurrentPosition))
            CurrentData.Y = 0.95 * Abs(1 / (MeasPos - CurrentPosition))
            CurrentData.Z = Abs(1 / (MeasPos - CurrentPosition))
            GaussianX(i) = CurrentData.X
            GaussianY(i) = CurrentData.Y
            GaussianZ(i) = CurrentData.Z
            RodPosition(i) = CurrentPosition
        Else
            CurrentData.X = CurrentData.X
            CurrentData.Y = CurrentData.Y
            CurrentData.Z = CurrentData.Z
            GaussianX(i) = CurrentData.X
            GaussianY(i) = CurrentData.Y
            GaussianZ(i) = CurrentData.Z
            RodPosition(i) = frmDCMotors.UpDownHeight
        End If
        txtUpDown = Format(CurrentPosition, "#0.0##")
        txtUpDownSampleCm = Trim(Str((Abs(CurrentPosition) - Abs(StartPosition)) / UpDownMotor1cm))
    
        txtData(0) = Str$(CurrentData.X)
        txtData(1) = Str$(CurrentData.Y)
        txtData(2) = Str$(CurrentData.Z)
        
        WriteUChannelData WriteFile, _
                          CurrentPosition, _
                              DateDiff("s", StartTime, Now), _
                          Int(frmDCMotors.UpDownHeight), _
                          frmDCMotors.TurningMotorAngle, _
                          CurrentData
        
        'Update the time remaining
        UChannelUpdateTime DateDiff("s", StartTime, Now), i + 1
        
        For k = 1 To 4
        
            SampleMeasurement.Sample(k).X = CurrentData.X
            SampleMeasurement.Sample(k).Y = CurrentData.Y
            SampleMeasurement.Sample(k).Z = CurrentData.Z
                
        Next k
        
        Set CurrentData = SampleMeasurement.CorrectedSample(2)
        
        txtData(3) = Str$(CurrentData.inc)
        txtData(4) = Str$(CurrentData.dec)
                
        Set CurrentData = Nothing
                
    Next i
    
    'Reset During Measurement flag to false
    DuringMeas = False
    
    'Home to top and end sub if signle scan is checked by user (10/23/2009, I Hilburn)
    If chkSingleScan.Value = Checked Then
            
        'Update the time remaining
        Me.txtElapsedTime = FormatSecondsToStr(DateDiff("s", StartTime, Now))
        Me.txtTimeRemaining = FormatSecondsToStr(SetupTime)
            
        frmDCMotors.HomeToTop
        
        'Set Remaining time to 0
        'Update the time remaining
        Me.txtElapsedTime = FormatSecondsToStr(DateDiff("s", StartTime, Now))
        Me.txtTimeRemaining = FormatSecondsToStr(0)
        
        'Set code level to orange
        SetCodeLevel CodeOrange
        
        'Notify User that the sample is done
        frmSendMail.MailNotification "2G Status Update", _
                                    "Long core sample done.  Sample saved to file: " & Me.txtFileName
                
        MsgBox "Please remove sample from the quartz glass tube." & _
               "The vacuum will turn OFF after you click ""OK"""
        
        'If the vacuum is on, turn it off
        If frmVacuum.VacuumConnectOn = True Then frmVacuum.ValveConnect False
        If frmVacuum.VacuumMotorOn = True Then frmVacuum.MotorPower False
            
        'Set the code level back to Code blue
        SetCodeLevel CodeBlue, True
        
        Exit Sub
        
    End If
    
    'Update the time remaining
    UChannelUpdateTime DateDiff("s", StartTime, Now), coarse
    
    frmDCMotors.TurningMotorRotate 90, False
    
    'Update the time remaining
    UChannelUpdateTime DateDiff("s", StartTime, Now), coarse
    
    'Do a zero measurement
    
    '(Mar 2010, I Hilburn) - changing it again so that the position is incremented
    'in centimeters and then translated for each movement of the updown arm from centimeters
    'to motor counts to prevent position drift due to cm to motor count conversion rounding errors
    StartPosInCm = EndPosition * Sgn(EndPosition) _
                    / UpDownMotor1cm * Sgn(UpDownMotor1cm)
    
    Warning = False
    
    
    For i = 0 To coarse - 1
        If Not CurrentlyRunning Then Exit Sub
        
        '(Mar, 2010 I Hilburn) - Calculate current position in Centimeters from the start position
        CurPosInCm = StartPosInCm * Sgn(StartPosInCm) _
                        - i * Spacing * Sgn(Spacing)
        
        '(Mar, 2010 I Hilburn) Then translate this current position to a motor position
        'This eliminates fractional drift do to the fact that Spacing * UpDownMotor1cm is not an integer value
        CurrentPosition = (Sgn(ZeroPos)) * CurPosInCm * UpDownMotor1cm * (Sgn(UpDownMotor1cm))
        
        If Me.chkDoZeroMeas.Value = Checked Then
        
            'Check to see if need to do a zero measurement this step
            If (i + coarse) Mod ZeroEvery = 0 Then
        
                'Update the time remaining,
                'Don't add the time for this measurement to average
                UChannelUpdateTime DateDiff("s", StartTime, Now), _
                                   i + coarse + 1
                                   
                DoZeroMeas i + coarse + 1, StartTime, SampleMeasurement, True, CurrentPosition
                
                'Set DuringMeas = False, don't want to capture the time taken for this measurement
                DuringMeas = False
                                   
            Else
            
                'This measurement is normal, can be used to calculate the time
                'per measurement
                UChannelUpdateTime DateDiff("s", StartTime, Now), _
                                   i + coarse + 1, _
                                   True
                                   
            End If
            
        Else
        
            'This measurement is normal, can be used to calculate the time
            'per measurement
            UChannelUpdateTime DateDiff("s", StartTime, Now), _
                               i + coarse + 1, _
                               True
                               
        End If
        
        
        If Abs(CurrentPosition) >= Spacing * Abs(UpDownMotor1cm) Then 'To don't hit the switch...
            
            frmDCMotors.UpDownMove CurrentPosition, 1
            
            'Update the time remaining
            UChannelUpdateTime DateDiff("s", StartTime, Now), _
                               i + coarse + 1
            
            
            DelayTime (Delay * 0.1)
            Set CurrentData = frmSQUID.getData
            
            'Update the time remaining
            UChannelUpdateTime DateDiff("s", StartTime, Now), _
                               i + coarse + 1
                    
            If NOCOMM_MODE Then
            CurrentData.X = 0.9 * Abs(1 / (MeasPos - CurrentPosition))
            CurrentData.Y = 0.95 * Abs(1 / (MeasPos - CurrentPosition))
            CurrentData.Z = Abs(1 / (MeasPos - CurrentPosition))
            Else
            CurrentData.X = CurrentData.X
            CurrentData.Y = CurrentData.Y
            CurrentData.Z = CurrentData.Z
            End If
            txtUpDown = CurrentPosition
            txtUpDownSampleCm = Trim(Str((Abs(CurrentPosition) - Abs(StartPosition)) / UpDownMotor1cm))
        
            txtData(0) = Str$(CurrentData.X)
            txtData(1) = Str$(CurrentData.Y)
            txtData(2) = Str$(CurrentData.Z)
            
            WriteUChannelData WriteFile, _
                              CurrentPosition, _
                              DateDiff("s", StartTime, Now), _
                              Int(frmDCMotors.UpDownHeight), _
                              frmDCMotors.TurningMotorAngle, _
                              CurrentData
            
            'Update the time remaining
            UChannelUpdateTime DateDiff("s", StartTime, Now), _
                               i + coarse + 1
                        
            For k = 1 To 4
        
                SampleMeasurement.Sample(k).X = CurrentData.X
                SampleMeasurement.Sample(k).Y = CurrentData.Y
                SampleMeasurement.Sample(k).Z = CurrentData.Z
            
            
            Next k
            
            Set CurrentData = SampleMeasurement.CorrectedSample(2)
                
            txtData(3) = Str$(CurrentData.inc)
            txtData(4) = Str$(CurrentData.dec)
        
            Set CurrentData = Nothing
        Else
            Warning = True
        End If
        
        If Abs(CurrentPosition) <= Abs(StartPosition) Then
        
            Warning = True
            
        End If
            
    Next i
    
    '(Mar 2010, I Hilburn) - changing it again so that the position is incremented
    'in centimeters and then translated for each movement of the updown arm from centimeters
    'to motor counts to prevent position drift due to cm to motor count conversion rounding errors
    StartPosInCm = StartPosition * Sgn(StartPosition) _
                     / UpDownMotor1cm * Sgn(UpDownMotor1cm)
                         
    'Update the time remaining
    UChannelUpdateTime DateDiff("s", StartTime, Now), _
                       2 * coarse
                                 
                         
    frmDCMotors.TurningMotorRotate 180, False
    
    'Update the time remaining
    UChannelUpdateTime DateDiff("s", StartTime, Now), 2 * coarse
    
    For i = 0 To coarse - 1
        If Not CurrentlyRunning Then Exit Sub
        
        '(Mar, 2010 I Hilburn) - Calculate current position in Centimeters from the start position
        CurPosInCm = StartPosInCm * Sgn(StartPosInCm) _
                        + i * Spacing * Sgn(Spacing)
        
        '(Mar, 2010 I Hilburn) Then translate this current position to a motor position
        'This eliminates fractional drift do to the fact that Spacing * UpDownMotor1cm is not an integer value
        CurrentPosition = (Sgn(ZeroPos)) * CurPosInCm * UpDownMotor1cm * (Sgn(UpDownMotor1cm))
                        
        If Me.chkDoZeroMeas.Value = Checked Then
        
            'Check to see if need to do a zero measurement this step
            If (i + 2 * coarse) Mod ZeroEvery = 0 Then
        
                'Update the time remaining,
                'Don't add the time for this measurement to average
                UChannelUpdateTime DateDiff("s", StartTime, Now), _
                                   i + 2 * coarse + 1
                                   
                DoZeroMeas i + 2 * coarse + 1, StartTime, SampleMeasurement, True, CurrentPosition
                
                'Set DuringMeas = False, don't want to capture the time taken for this measurement
                DuringMeas = False
                                   
            Else
            
                'This measurement is normal, can be used to calculate the time
                'per measurement
                UChannelUpdateTime DateDiff("s", StartTime, Now), _
                                   i + 2 * coarse + 1, _
                                   True
                                   
            End If
            
        Else
        
            'This measurement is normal, can be used to calculate the time
            'per measurement
            UChannelUpdateTime DateDiff("s", StartTime, Now), _
                               i + 2 * coarse + 1, _
                               True
                               
        End If
                        
        frmDCMotors.UpDownMove CurrentPosition, 0
        DelayTime (Delay)
        
        'Update the time remaining
        UChannelUpdateTime DateDiff("s", StartTime, Now), _
                               i + 2 * coarse + 1
                
        Set CurrentData = frmSQUID.getData
        
        'Update the time remaining
        UChannelUpdateTime DateDiff("s", StartTime, Now), _
                               i + 2 * coarse + 1
        
        If NOCOMM_MODE Then
            CurrentData.X = 0.9 * Abs(1 / (MeasPos - CurrentPosition))
            CurrentData.Y = 0.95 * Abs(1 / (MeasPos - CurrentPosition))
            CurrentData.Z = Abs(1 / (MeasPos - CurrentPosition))
            GaussianX(i) = CurrentData.X
            GaussianY(i) = CurrentData.Y
            GaussianZ(i) = CurrentData.Z
            RodPosition(i) = CurrentPosition
        Else
            CurrentData.X = CurrentData.X
            CurrentData.Y = CurrentData.Y
            CurrentData.Z = CurrentData.Z
            GaussianX(i) = CurrentData.X
            GaussianY(i) = CurrentData.Y
            GaussianZ(i) = CurrentData.Z
            RodPosition(i) = frmDCMotors.UpDownHeight
        End If
        txtUpDown = CurrentPosition
        txtUpDownSampleCm = Trim(Str((Abs(CurrentPosition) - Abs(StartPosition)) / UpDownMotor1cm))
    
        txtData(0) = Str$(CurrentData.X)
        txtData(1) = Str$(CurrentData.Y)
        txtData(2) = Str$(CurrentData.Z)
        
        WriteUChannelData WriteFile, _
                          CurrentPosition, _
                          DateDiff("s", StartTime, Now), _
                          Int(frmDCMotors.UpDownHeight), _
                          frmDCMotors.TurningMotorAngle, _
                          CurrentData
        
        'Update the time remaining
        UChannelUpdateTime DateDiff("s", StartTime, Now), _
                               i + 2 * coarse + 1
        
        For k = 1 To 4
        
            SampleMeasurement.Sample(k).X = CurrentData.X
            SampleMeasurement.Sample(k).Y = CurrentData.Y
            SampleMeasurement.Sample(k).Z = CurrentData.Z
        
        
        Next k
        
        Set CurrentData = SampleMeasurement.CorrectedSample(2)
        
        txtData(3) = Str$(CurrentData.inc)
        txtData(4) = Str$(CurrentData.dec)
        
        Set CurrentData = Nothing
                
    Next i
    
    frmDCMotors.TurningMotorRotate 270, False
    
    'Update the time remaining
    UChannelUpdateTime DateDiff("s", StartTime, Now), 3 * coarse
    
    '(Mar 2010, I Hilburn) - changing it again so that the position is incremented
    'in centimeters and then translated for each movement of the updown arm from centimeters
    'to motor counts to prevent position drift due to cm to motor count conversion rounding errors
    StartPosInCm = EndPosition * EndPosition / Abs(EndPosition) _
                    / UpDownMotor1cm * Sgn(UpDownMotor1cm)
        
    Warning = False
    
    For i = 0 To coarse - 1
        If Not CurrentlyRunning Then Exit Sub
        
        '(Mar, 2010 I Hilburn) - Calculate current position in Centimeters from the start position
        CurPosInCm = StartPosInCm * Sgn(StartPosInCm) _
                        - i * Spacing * Sgn(Spacing)
        
        '(Mar, 2010 I Hilburn) Then translate this current position to a motor position
        'This eliminates fractional drift do to the fact that Spacing * UpDownMotor1cm is not an integer value
        CurrentPosition = (Sgn(ZeroPos)) * CurPosInCm * UpDownMotor1cm * (Sgn(UpDownMotor1cm))
        
        If Me.chkDoZeroMeas.Value = Checked Then
        
            'Check to see if need to do a zero measurement this step
            If (i + 3 * coarse) Mod ZeroEvery = 0 Then
        
                'Update the time remaining,
                'Don't add the time for this measurement to average
                UChannelUpdateTime DateDiff("s", StartTime, Now), _
                                   i + 3 * coarse + 1
                                   
                DoZeroMeas i + 3 * coarse + 1, StartTime, SampleMeasurement, True, CurrentPosition
                
                'Set DuringMeas = False, don't want to capture the time taken for this measurement
                DuringMeas = False
                                   
            Else
            
                'This measurement is normal, can be used to calculate the time
                'per measurement
                UChannelUpdateTime DateDiff("s", StartTime, Now), _
                                   i + 3 * coarse + 1, _
                                   True
                                   
            End If
            
        Else
        
            'This measurement is normal, can be used to calculate the time
            'per measurement
            UChannelUpdateTime DateDiff("s", StartTime, Now), _
                               i + 3 * coarse + 1, _
                               True
                               
        End If
        
        If Abs(CurrentPosition) >= Spacing * Abs(UpDownMotor1cm) Then 'To don't hit the switch...
        
            'Update the time remaining
            UChannelUpdateTime DateDiff("s", StartTime, Now), _
                                   i + 3 * coarse + 1
        
            frmDCMotors.UpDownMove CurrentPosition, 1
            
            DelayTime (Delay * 0.1)
            Set CurrentData = frmSQUID.getData
            
            'Update the time remaining
            UChannelUpdateTime DateDiff("s", StartTime, Now), _
                                   i + 3 * coarse + 1
            
            If NOCOMM_MODE Then
            CurrentData.X = 0.9 * Abs(1 / (MeasPos - CurrentPosition))
            CurrentData.Y = 0.95 * Abs(1 / (MeasPos - CurrentPosition))
            CurrentData.Z = Abs(1 / (MeasPos - CurrentPosition))
            Else
            CurrentData.X = CurrentData.X
            CurrentData.Y = CurrentData.Y
            CurrentData.Z = CurrentData.Z
            End If
            txtUpDown = CurrentPosition
            txtUpDownSampleCm = Trim(Str((Abs(CurrentPosition) - Abs(StartPosition)) / UpDownMotor1cm))
        
            txtData(0) = Str$(CurrentData.X)
            txtData(1) = Str$(CurrentData.Y)
            txtData(2) = Str$(CurrentData.Z)
            
            WriteUChannelData WriteFile, _
                              CurrentPosition, _
                              DateDiff("s", StartTime, Now), _
                              Int(frmDCMotors.UpDownHeight), _
                              frmDCMotors.TurningMotorAngle, _
                              CurrentData
            
            'Update the time remaining
            UChannelUpdateTime DateDiff("s", StartTime, Now), _
                                   i + 3 * coarse + 1
            
            For k = 1 To 4
        
                SampleMeasurement.Sample(k).X = CurrentData.X
                SampleMeasurement.Sample(k).Y = CurrentData.Y
                SampleMeasurement.Sample(k).Z = CurrentData.Z
            
            
            Next k
            
            Set CurrentData = SampleMeasurement.CorrectedSample(2)
            
            txtData(3) = Str$(CurrentData.inc)
            txtData(4) = Str$(CurrentData.dec)
        
            Set CurrentData = Nothing
        Else
            Warning = True
        End If
        
        If Abs(CurrentPosition) <= Abs(StartPosition) Then
        
            Warning = True
            
        End If
            
    Next i
        
    'Update the time remaining
    UChannelUpdateTime DateDiff("s", StartTime, Now), 4 * coarse
    
    Warning = False
    frmDCMotors.TurningMotorRotate 0, False
    
    'Update the time remaining
    Me.txtElapsedTime = FormatSecondsToStr(DateDiff("s", StartTime, Now))
    Me.txtTimeRemaining = FormatSecondsToStr(SetupTime + TimePerMeasurement)
        
    frmDCMotors.HomeToTop
    
    'Update the time remaining
    Me.txtElapsedTime = FormatSecondsToStr(DateDiff("s", StartTime, Now))
    Me.txtTimeRemaining = FormatSecondsToStr(TimePerMeasurement)
   
    DelayTime (Delay)
    Set CurrentData = frmSQUID.getData
    If NOCOMM_MODE Then
        CurrentData.X = 0.9 * Abs(1 / (MeasPos - CurrentPosition))
        CurrentData.Y = 0.95 * Abs(1 / (MeasPos - CurrentPosition))
        CurrentData.Z = Abs(1 / (MeasPos - CurrentPosition))
        GaussianX(i) = CurrentData.X
        GaussianY(i) = CurrentData.Y
        GaussianZ(i) = CurrentData.Z
        RodPosition(i) = CurrentPosition
    Else
        CurrentData.X = CurrentData.X
        CurrentData.Y = CurrentData.Y
        CurrentData.Z = CurrentData.Z
        GaussianX(i) = CurrentData.X
        GaussianY(i) = CurrentData.Y
        GaussianZ(i) = CurrentData.Z
        RodPosition(i) = frmDCMotors.UpDownHeight
    End If
    txtUpDown = Format(CurrentPosition, "#0.0##")
    txtUpDownSampleCm = Trim(Str((Abs(CurrentPosition) - Abs(StartPosition)) / UpDownMotor1cm))

    txtData(0) = Str$(CurrentData.X)
    txtData(1) = Str$(CurrentData.Y)
    txtData(2) = Str$(CurrentData.Z)
    
    WriteUChannelData WriteFile, _
                      CurrentPosition, _
                              DateDiff("s", StartTime, Now), _
                      Int(frmDCMotors.UpDownHeight), _
                      frmDCMotors.TurningMotorAngle, _
                      CurrentData
    
    For k = 1 To 4
    
        SampleMeasurement.Sample(k).X = CurrentData.X
        SampleMeasurement.Sample(k).Y = CurrentData.Y
        SampleMeasurement.Sample(k).Z = CurrentData.Z
            
    Next k
    
    Set CurrentData = SampleMeasurement.CorrectedSample(2)
    
    txtData(3) = Str$(CurrentData.inc)
    txtData(4) = Str$(CurrentData.dec)
        
    Set CurrentData = Nothing
    
    SetCodeLevel CodeOrange
    
    'Set Remaining time to 0
    'Update the time remaining
    Me.txtElapsedTime = FormatSecondsToStr(DateDiff("s", StartTime, Now))
    Me.txtTimeRemaining = FormatSecondsToStr(0)
    
    
    'Notify User that the sample is done
    frmSendMail.MailNotification "2G Status Update", _
                        "Long core sample done.  Sample saved to file: " & Me.txtFileName, _
                        CodeOrange, _
                        True
                        
    MsgBox "Please remove sample from the quartz glass tube."
        
    'If the vacuum is on, turn it off
    If frmVacuum.VacuumConnectOn = True Then frmVacuum.ValveConnect False
    If frmVacuum.VacuumMotorOn = True Then frmVacuum.MotorPower False
        
    'Reset Module Global Measurement Time Calculation variables
    AvgMeasTime = -1
    MeasStartTime = -1
    DuringMeas = False
    DuringMoveToZero = False
    StartZeroMoveTime = -1
    AvgZeroMoveRate = -1
        
    'Set the code level back to Code blue
    SetCodeLevel CodeBlue, True
        
End Sub

Private Sub txtCoreLength_Change()

    Dim EndPosition As Long
    Dim temp As Integer

    StartPosition = CLng(val(txtStartPosition))
    
    temp = MeasPos / Abs(MeasPos)

    EndPosition = (30500 + CLng(Abs(val(txtCoreLength) * UpDownMotor1cm))) * temp 'MeasPos / Abs(MeasPos)
     
    If Abs(EndPosition) > 43000 Then
    
        EndPosition = MeasPos / Abs(MeasPos) * 43000

    End If
    
    
  
    lblEndPosition.Caption = "End Position:     " & Trim(Str(EndPosition)) & "  motor counts"

End Sub

Private Sub txtEPRBottomBuffer_Change()

    txtSpacing_Change

End Sub

Private Sub txtEPRSampleBottom_Change()

    txtEPRSampleTop_Change

End Sub

Private Sub txtEPRSampleTop_Change()

    Me.lblSampleLen.Caption = "Sample Length (cm): " & _
                              Format(val(Me.txtEPRSampleTop) - val(Me.txtEPRSampleBottom), "#0.0##")
                              
    'Update the estimated time & number of measurements
    txtSpacing_Change
                              
End Sub

Private Sub txtEPRTopBuffer_Change()

    txtSpacing_Change

End Sub

Private Sub txtEPRTubeLength_Change()

    txtSpacing_Change

End Sub

Private Sub txtMeasurementsPerZero_Change()

    txtSpacing_Change

End Sub

Private Sub txtSpacing_Change()

    Dim MeasLen As Double
    Dim NumSteps As Long
    Dim NumZeros As Long
    Dim TravelDist As Double
    Dim TravelTime As Long
    Dim i As Long

    If val(Me.txtSpacing) <= 0 Then Exit Sub

    'Need to recalculate the time & number of steps that the code will last
    
    'Get the sample length
    SampleLenCm = Abs(val(Me.txtEPRSampleTop) - val(Me.txtEPRSampleBottom))

    'Get the measurement length
    MeasLen = SampleLenCm + val(Me.txtEPRBottomBuffer) + val(Me.txtEPRTopBuffer)

    'Calculate the number of steps
    NumSteps = CLng(MeasLen / val(Me.txtSpacing)) + 1

    'If single-scan is NOT selected, times NumSteps by 4
    If Me.chkSingleScan.Value = Unchecked Then
    
        NumSteps = NumSteps * 4
        
    End If
    
    'Now display the number of steps
    Me.txtNumberMeasurements = Trim(Str(NumSteps))

    'Now estimate the amount of time the run will take
    RunTimeSeconds = 0
    
    RunTimeSeconds = CLng((NumSteps) * TimePerMeasurement)
        
    'Now estimate the time spent doing zeros during the scan
    If Me.chkDoZeroMeas.Value = Checked And val(Me.txtMeasurementsPerZero) <> 0 Then
    
        'Calculate the total movement time associated with doing zeros
        NumZeros = Int(NumSteps / val(Me.txtMeasurementsPerZero)) + 1
        
        'Initialize travel distance for doing all of the zero measurements to zero
        TravelDist = 0
        
        For i = 0 To NumZeros - 1
        
            'Travel Distance is in motor units
            TravelDist = TravelDist + _
                         (Abs(5 * UpDownMotor1cm) + _
                          i * val(Me.txtMeasurementsPerZero) * _
                              val(Me.txtSpacing) * _
                              Abs(UpDownMotor1cm)) * 2
                             
        Next i
        
        'Travel time assumes that the motor speed is the number
        'of motor counts moved per hour (McPH)
        TravelTime = CLng(TravelDist / modConfig.LiftSpeedSlow * 3600)
        
        RunTimeSeconds = RunTimeSeconds + _
                         TravelTime + _
                         (NumZeros * TimePerMeasurement)     'SQUID measurement time
                              
    End If
    
    'Calculate Setup time
    'Get the Start Motor Position that the sample is going to be moved to
    StartPosition = MeasPos - (Sgn(MeasPos) * _
                               CLng((Abs((val(Me.txtEPRTubeLength) - val(Me.txtEPRSampleTop)) * _
                                     UpDownMotor1cm) + _
                                    SampleLenCm * Abs(UpDownMotor1cm) + _
                                    Abs(val(Me.txtEPRBottomBuffer) * UpDownMotor1cm))))
                                    
    SetupTime = CLng(Abs(StartPosition) / modConfig.LiftSpeedSlow * 3600)
    
    'Add in Time For moving sample from home position to zero position,
    'for the first & last zero measurements, and returning the sample
    'to the home position
    RunTimeSeconds = RunTimeSeconds + SetupTime * 2
    
    'Now need to reformat RunTimeSeconds as a Day:HH:MM:SS string
    Me.txtEstimatedRunTime = FormatSecondsToStr(RunTimeSeconds)
    
End Sub

Private Sub UChannelUpdateTime(ByVal CurSecs As Long, _
                               ByVal CurStep As Long, _
                               Optional ByVal IsMeas As Boolean = False, _
                               Optional ByVal IsZero As Boolean = False, _
                               Optional ByVal BeforeZeroPos As Long = -1, _
                               Optional ByVal CurZero As Long = 0)
                               
    Dim MeasLen As Double
    Dim NumSteps As Long
    Dim NumZeros As Long
    Dim CurZeroMoveRate As Double
    Dim TravelDist As Double
    Dim TravelTime As Long
    Dim i As Long
    Dim j As Long

    If val(Me.txtSpacing) <= 0 Then Exit Sub

    'Get the sample length
    SampleLenCm = Abs(val(Me.txtEPRSampleTop) - val(Me.txtEPRSampleBottom))

    'Get the measurement length
    MeasLen = SampleLenCm + val(Me.txtEPRBottomBuffer) + val(Me.txtEPRTopBuffer)

    'Calculate the number of steps
    NumSteps = val(Me.txtNumberMeasurements)
    
    'Now, check to see if we can update TimePerMeasurement
    If IsMeas = True And DuringMeas = False Then
    
        'This is the start of the current measurement cycle
        MeasStartTime = CurSecs
        
        'Set the During measurement flag to True
        DuringMeas = True
        
    ElseIf IsMeas = True And DuringMeas = True Then
    
        'This is the end of the measurement cycle
        'UpDate the Average measurement time
        If AvgMeasTime = -1 Then
        
            AvgMeasTime = CurSecs - MeasStartTime
            
        Else
        
            AvgMeasTime = AvgMeasTime * ((CurStep - 1) / CurStep) + (CurSecs - MeasStartTime) / CurStep _
                          
        End If
        
        'Set the time per measurement to the new average
        TimePerMeasurement = AvgMeasTime
        
        'Set DuringMeas = False
        DuringMeas = False
        
    End If
    
    'Now estimate the amount of time the remaining run will take
    RunTimeSeconds = 0
    
    RunTimeSeconds = CLng((NumSteps - 3 - CurStep) * TimePerMeasurement)
        
    'Now estimate the time spent doing zeros during the scan
    If Me.chkDoZeroMeas.Value = Checked Then
    
        'Check to see if we can recalculate the time taken to move to
        'the zero position
        If IsZero = True And DuringMoveToZero = False Then
        
            'Were at the start of a movement of the motors to the zero position
            
            'Store the current step as the last zero step
            LastZeroStep = CurStep
            
            'Store the time
            StartZeroMoveTime = CurSecs
            
            DuringMoveToZero = True
            
        ElseIf IsZero = True And DuringMoveToZero = True Then
        
            'Were at the end of a movement of the motors to the zero position
            
            'Calculate the new zero-move motor steps / time
            CurZeroMoveRate = Abs(Abs(MyZeroPos) - Abs(BeforeZeroPos)) _
                              / (CurSecs - StartZeroMoveTime)
        
            'Update the Average Zero Move Rate
            If AvgZeroMoveRate = -1 Then
            
                AvgZeroMoveRate = CurZeroMoveRate
                        
            Else
            
                AvgZeroMoveRate = AvgZeroMoveRate * ((CurZero - 1) / CurZero) _
                                  + CurZeroMoveRate / CurZero
                        
            End If
            
        End If
                        
        'Initialize travel distance for doing all of the zero measurements to zero
        TravelDist = 0
        j = 0
        
        For i = (LastZeroStep + val(Me.txtMeasurementsPerZero)) _
                To NumSteps _
                Step val(Me.txtMeasurementsPerZero)
        
            'Travel Distance is in motor units
            TravelDist = TravelDist + _
                         (Abs(Abs(MyZeroPos) - Abs(StartPosition)) + _
                          (i * _
                           val(Me.txtSpacing) * _
                           Abs(UpDownMotor1cm))) _
                         * 2
                                     
            'Count the remaining number of zeros
            j = j + 1
                             
        Next i
        
        Debug.Print "Update = " & Trim(Str(TravelDist))
        
        'Travel time uses the average motor travel rate during a zero measurement
        If AvgZeroMoveRate = 0 Then
        
            TravelTime = CLng(TravelDist / modConfig.LiftSpeedSlow * 3600)
        
        Else
        
            TravelTime = CLng(TravelDist / AvgZeroMoveRate)
            
        End If
        
        'Calculate run time using the number of zeros remaining
        RunTimeSeconds = RunTimeSeconds + _
                         TravelTime + _
                         (j * TimePerMeasurement)     'SQUID measurement time
                              
    End If
    
    'Calculate Setup time
    If AvgZeroMoveRate = 0 Then
    
        SetupTime = Abs(MyZeroPos / modConfig.LiftSpeedSlow * 3600)
        
    Else
        
        SetupTime = Abs(MyZeroPos / AvgZeroMoveRate)
        
    End If
    
    'Add in Time For moving sample from zero position to home
    'for the last zero measurement
    RunTimeSeconds = RunTimeSeconds + CLng(SetupTime)
    
    'Now need to reformat RunTimeSeconds as a Day:HH:MM:SS string
    Me.txtTimeRemaining = FormatSecondsToStr(RunTimeSeconds)
    Me.txtElapsedTime = FormatSecondsToStr(CurSecs)
                               
End Sub

Private Sub WriteCalRodData(filename As String, ReqPos As Long, RealPos As Long, data As Cartesian3D)
    Dim filenum As Long
    filenum = FreeFile
    On Error GoTo oops
    Open filename For Append As #filenum
    With data
    Print #filenum, ReqPos; ","; RealPos; ","; .X; ","; .Y; ","; .Z
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

Private Sub WriteUChannelData(ByVal filename As String, _
                              ByVal ReqPos As Long, _
                              ByVal CurSecs As Long, _
                              ByVal RealPos As Long, _
                              ByVal CurAngle As Double, _
                              ByRef CurData As Cartesian3D)
                              
    Dim fso As FileSystemObject
    Dim TStream As TextStream
    Dim TempStr As String
                              
    Set fso = New FileSystemObject
    
    Set TStream = fso.OpenTextFile(filename, ForAppending)
    

    TempStr = Format(CurAngle, "#0.##")
    
    If Right(TempStr, 1) = "." Then
        
        TempStr = Mid(TempStr, 1, Len(TempStr) - 1)
        
    End If

    TStream.WriteLine Trim(Str(ReqPos)) & "," & _
                      Trim(Str(RealPos)) & "," & _
                      Trim(Str(CurSecs)) & "," & _
                      TempStr & "," & _
                      Trim(Str(CurData.X)) & "," & _
                      Trim(Str(CurData.Y)) & "," & _
                      Trim(Str(CurData.Z))
                          

                              
    TStream.Close
                              
End Sub

Private Sub WriteUChannelHeaders(ByRef filename As String, _
                                 ByVal SampleLength As Double, _
                                 Optional ByVal ScanLength As Double = -1, _
                                 Optional ByVal SampleTopPos As Double = -1, _
                                 Optional ByVal SampleBottomPos As Double = -1, _
                                 Optional ByVal EPRTubeLength As Double = -1)
                              
    Dim fso As FileSystemObject
    Dim TStream As TextStream
    Dim TempStr As String
    Dim TempIndex As Long
    Dim TempL As String
                              
    Set fso = New FileSystemObject
    
    If fso.FileExists(filename) = True Then
    
        TempL = InStrRev(filename, ".")
        
        If TempL = 0 Then
        
            TempStr = filename
            
        Else
        
            If IsNumeric(Mid(filename, TempL - 1, 1)) And _
               Mid(filename, TempL - 2, 1) = "_" _
            Then
            
                TempIndex = val(Mid(filename, TempL - 1, 1))
                TempStr = Mid(filename, 1, TempL - 3)
                
            Else
            
                TempIndex = 2
                TempStr = Mid(filename, 1, TempL - 1)
                
            End If
        
            TempStr = TempStr & "_" & Trim(Str(TempIndex)) & ".csv"
            
        End If
        
        filename = TempStr
        
    End If
    
    Set TStream = fso.OpenTextFile(filename, ForWriting, True)
    
    'Check to see which type of U-scan was run
    If Me.chkDoEPRTubeScan.Value = Checked Then
    
        TempStr = "EPR Tube"
        
    Else
    
        TempStr = "Long Core"
    
    End If
    
    'First write a general Intro
    TStream.WriteLine TempStr & " Scan"
    TStream.WriteLine Format(Now, "dddddd")
    TStream.WriteLine Format(Now, "ttttt")
    TStream.WriteLine "Sample Length (cm):," & Trim(Str(SampleLength))
    TStream.WriteLine "Scan Step size (cm):," & Trim(Me.txtSpacing)
    
    If ScanLength <> -1 Then
    
        TempStr = Format(ScanLength, "#0.##")
        If Right(TempStr, 1) = "." Then
        
            TempStr = Mid(TempStr, 1, Len(TempStr) - 1)
            
        End If
    
        TStream.WriteLine "Scan Length (cm):," & TempStr
    
    End If
    
    'If this is an EPR scan, then need to write in more info
    If Me.chkDoEPRTubeScan.Value = Checked Then
    
        TempStr = Format(SampleTopPos, "#0.##")
        If Right(TempStr, 1) = "." Then
        
            TempStr = Mid(TempStr, 1, Len(TempStr) - 1)
        
        End If
    
        TStream.WriteLine "Sample Top (cm above end of EPR Tube):," & TempStr
    
        TempStr = Format(SampleBottomPos, "#0.##")
        If Right(TempStr, 1) = "." Then
        
            TempStr = Mid(TempStr, 1, Len(TempStr) - 1)
        
        End If
        
        TStream.WriteLine "Sample Bottom (cm above end of tube):," & TempStr
    
        TempStr = Format(EPRTubeLength, "#0.##")
        If Right(TempStr, 1) = "." Then
        
            TempStr = Mid(TempStr, 1, Len(TempStr) - 1)
        
        End If
        
        TStream.WriteLine "EPR Tube Length (cm)::," & TempStr
        
    End If
    
    TStream.WriteBlankLines 1
    TStream.WriteLine "Requested Up/Down Pos," & _
                      "Actual Up/Down Pos," & _
                      "Time (sec)," & _
                      "Turn Angle," & _
                      "SQUID X," & _
                      "SQUID Y," & _
                      "SQUID Z"
                      
    TStream.Close
                              
End Sub

