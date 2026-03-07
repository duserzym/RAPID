VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmSettings 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Settings Dialog"
   ClientHeight    =   4560
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   6975
   ControlBox      =   0   'False
   Icon            =   "frmSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   6975
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   3360
      TabIndex        =   0
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4560
      TabIndex        =   1
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   375
      Left            =   5760
      TabIndex        =   2
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton cmdCalibNewRod 
      Caption         =   "Calibration of the rod"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   4080
      Width           =   1620
   End
   Begin VB.Frame frameOptions 
      BorderStyle     =   0  'None
      Caption         =   "DC Motor settings"
      Height          =   3372
      Index           =   0
      Left            =   240
      TabIndex        =   4
      Top             =   480
      Width           =   6492
      Begin VB.TextBox txtZeroPos 
         Height          =   288
         Left            =   1680
         TabIndex        =   5
         Top             =   120
         Width           =   1212
      End
      Begin VB.TextBox txtMeasPos 
         Height          =   288
         Left            =   1680
         TabIndex        =   6
         Top             =   480
         Width           =   1212
      End
      Begin VB.TextBox txtAFPos 
         Height          =   288
         Left            =   1680
         TabIndex        =   7
         Top             =   840
         Width           =   1212
      End
      Begin VB.TextBox txtIRMPos 
         Height          =   288
         Left            =   1680
         TabIndex        =   8
         Top             =   1200
         Width           =   1212
      End
      Begin VB.TextBox txtSCoilPos 
         Height          =   288
         Left            =   1680
         TabIndex        =   9
         Top             =   1560
         Width           =   1212
      End
      Begin VB.TextBox txtSampleTop 
         Height          =   288
         Left            =   1680
         TabIndex        =   10
         Top             =   1920
         Width           =   1212
      End
      Begin VB.TextBox txtSampleBottom 
         Height          =   288
         Left            =   1680
         TabIndex        =   11
         Top             =   2280
         Width           =   1212
      End
      Begin VB.TextBox txtTurningMotorFullRotation 
         Height          =   288
         Left            =   1680
         TabIndex        =   12
         Top             =   2640
         Width           =   1212
      End
      Begin VB.TextBox txtTrayOffsetAngle 
         Height          =   288
         Left            =   1680
         TabIndex        =   13
         Top             =   3000
         Width           =   1212
      End
      Begin VB.TextBox txtLiftSpeedSlow 
         Height          =   288
         Left            =   4800
         TabIndex        =   14
         Top             =   120
         Width           =   1212
      End
      Begin VB.TextBox txtLiftSpeedNormal 
         Height          =   288
         Left            =   4800
         TabIndex        =   15
         Top             =   480
         Width           =   1212
      End
      Begin VB.TextBox txtLiftSpeedFast 
         Height          =   288
         Left            =   4800
         TabIndex        =   16
         Top             =   840
         Width           =   1212
      End
      Begin VB.TextBox txtChangerSpeed 
         Height          =   288
         Left            =   4800
         TabIndex        =   17
         Top             =   1200
         Width           =   1212
      End
      Begin VB.TextBox txtTurnerSpeed 
         Height          =   288
         Left            =   4800
         TabIndex        =   18
         Top             =   1560
         Width           =   1212
      End
      Begin VB.TextBox txtSCurveFactor 
         Height          =   288
         Left            =   4800
         TabIndex        =   19
         Top             =   1920
         Width           =   1212
      End
      Begin VB.TextBox txtSampHoleAlignOffset 
         Height          =   288
         Left            =   4800
         TabIndex        =   20
         Top             =   2280
         Width           =   1212
      End
      Begin VB.TextBox txtTurningMotor1rps 
         Height          =   288
         Left            =   4800
         TabIndex        =   21
         Top             =   2640
         Width           =   1212
      End
      Begin VB.TextBox txtUpDownMotor1cm 
         Height          =   288
         Left            =   4800
         TabIndex        =   22
         Top             =   3000
         Width           =   1212
      End
      Begin VB.Label Label20 
         Caption         =   "Zeroing position:"
         Height          =   252
         Left            =   240
         TabIndex        =   23
         Top             =   120
         Width           =   1572
      End
      Begin VB.Label Label15 
         Caption         =   "Measuring position:"
         Height          =   252
         Left            =   240
         TabIndex        =   24
         Top             =   480
         Width           =   1452
      End
      Begin VB.Label Label14 
         Caption         =   "AF position:"
         Height          =   372
         Left            =   240
         TabIndex        =   25
         Top             =   840
         Width           =   1572
      End
      Begin VB.Label Label13 
         Caption         =   "IRM position:"
         Height          =   372
         Left            =   240
         TabIndex        =   26
         Top             =   1200
         Width           =   1572
      End
      Begin VB.Label Label23 
         Caption         =   "S Coil position:"
         Height          =   252
         Left            =   240
         TabIndex        =   27
         Top             =   1560
         Width           =   1572
      End
      Begin VB.Label Label22 
         Caption         =   "Default samp. top:"
         Height          =   252
         Left            =   240
         TabIndex        =   28
         Top             =   1920
         Width           =   1452
      End
      Begin VB.Label Label21 
         Caption         =   "Default samp. bot.:"
         Height          =   372
         Left            =   240
         TabIndex        =   29
         Top             =   2280
         Width           =   1572
      End
      Begin VB.Label Label6 
         Caption         =   "Turning full rotation:"
         Height          =   372
         Left            =   240
         TabIndex        =   30
         Top             =   2640
         Width           =   1572
      End
      Begin VB.Label Label17 
         Caption         =   "Tray offset (deg.):"
         Height          =   375
         Left            =   240
         TabIndex        =   31
         Top             =   3000
         Width           =   1575
      End
      Begin VB.Label Label27 
         Caption         =   "Lift speed slow:"
         Height          =   252
         Left            =   3360
         TabIndex        =   32
         Top             =   120
         Width           =   1572
      End
      Begin VB.Label Label26 
         Caption         =   "Lift speed medium:"
         Height          =   252
         Left            =   3360
         TabIndex        =   33
         Top             =   480
         Width           =   1452
      End
      Begin VB.Label Label25 
         Caption         =   "Lift speed fast:"
         Height          =   372
         Left            =   3360
         TabIndex        =   34
         Top             =   840
         Width           =   1572
      End
      Begin VB.Label Label24 
         Caption         =   "Changer speed:"
         Height          =   372
         Left            =   3360
         TabIndex        =   35
         Top             =   1200
         Width           =   1572
      End
      Begin VB.Label Label29 
         Caption         =   "Turning speed:"
         Height          =   252
         Left            =   3360
         TabIndex        =   36
         Top             =   1560
         Width           =   1572
      End
      Begin VB.Label Label28 
         Caption         =   "S-curve factor:"
         Height          =   252
         Left            =   3360
         TabIndex        =   37
         Top             =   1920
         Width           =   1452
      End
      Begin VB.Label Label1 
         Caption         =   "Frac. samp. hole alignment offset:"
         Height          =   495
         Left            =   3360
         TabIndex        =   38
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label Label10 
         Caption         =   "Turning motor 1 rps:"
         Height          =   375
         Left            =   3360
         TabIndex        =   39
         Top             =   2640
         Width           =   1575
      End
      Begin VB.Label Label11 
         Caption         =   "Up/down 1 cm:"
         Height          =   375
         Left            =   3360
         TabIndex        =   40
         Top             =   3000
         Width           =   1575
      End
   End
   Begin VB.Frame frameOptions 
      BorderStyle     =   0  'None
      Caption         =   "Advanced settings"
      Height          =   3372
      Index           =   1
      Left            =   240
      TabIndex        =   41
      Top             =   480
      Width           =   6492
      Begin VB.TextBox txtMotorIDTurning 
         Height          =   288
         Left            =   1680
         TabIndex        =   42
         Top             =   480
         Width           =   1212
      End
      Begin VB.TextBox txtMotorIDChanger 
         Height          =   288
         Left            =   1680
         TabIndex        =   43
         Top             =   840
         Width           =   1212
      End
      Begin VB.TextBox txtMotorIDUpDown 
         Height          =   288
         Left            =   1680
         TabIndex        =   44
         Top             =   1200
         Width           =   1212
      End
      Begin VB.TextBox txtCmdHometoTop 
         Height          =   288
         Left            =   5040
         TabIndex        =   45
         Top             =   480
         Width           =   1212
      End
      Begin VB.TextBox txtCmdSamplePickup 
         Height          =   288
         Left            =   5040
         TabIndex        =   46
         Top             =   840
         Width           =   1212
      End
      Begin VB.TextBox txtUpDownTorquefactor 
         Height          =   288
         Left            =   2400
         TabIndex        =   47
         Top             =   1800
         Width           =   1212
      End
      Begin VB.TextBox txtPickupTorqueThrottle 
         Height          =   288
         Left            =   2400
         TabIndex        =   48
         Top             =   2160
         Width           =   1212
      End
      Begin VB.Label Label54 
         Caption         =   "Turning motor ID:"
         Height          =   252
         Left            =   240
         TabIndex        =   49
         Top             =   480
         Width           =   1572
      End
      Begin VB.Label Label49 
         Caption         =   "Changer motor ID:"
         Height          =   252
         Left            =   240
         TabIndex        =   50
         Top             =   840
         Width           =   1452
      End
      Begin VB.Label Label48 
         Caption         =   "Up/Down motor ID:"
         Height          =   372
         Left            =   240
         TabIndex        =   51
         Top             =   1200
         Width           =   1572
      End
      Begin VB.Label Label57 
         Caption         =   "Home to top code:"
         Height          =   252
         Left            =   3360
         TabIndex        =   52
         Top             =   480
         Width           =   1572
      End
      Begin VB.Label Label56 
         Caption         =   "Sample pickup code:"
         Height          =   252
         Left            =   3360
         TabIndex        =   53
         Top             =   840
         Width           =   1572
      End
      Begin VB.Label Label55 
         Caption         =   "Up/Down Torque Factor:"
         Height          =   255
         Left            =   0
         TabIndex        =   54
         Top             =   1800
         Width           =   2295
      End
      Begin VB.Label Label60 
         Caption         =   "Sample pickup torque throttle:"
         Height          =   255
         Left            =   0
         TabIndex        =   55
         Top             =   2160
         Width           =   2295
      End
   End
   Begin VB.Frame frameOptions 
      BorderStyle     =   0  'None
      Caption         =   "SQUID Magnetometer Calibration factors"
      Height          =   3372
      Index           =   2
      Left            =   240
      TabIndex        =   56
      Top             =   480
      Width           =   6492
      Begin VB.TextBox txtXCal 
         Height          =   288
         Left            =   2040
         TabIndex        =   57
         Top             =   120
         Width           =   972
      End
      Begin VB.TextBox txtYCal 
         Height          =   288
         Left            =   2040
         TabIndex        =   58
         Top             =   600
         Width           =   972
      End
      Begin VB.TextBox txtZCal 
         Height          =   288
         Left            =   2040
         TabIndex        =   59
         Top             =   1080
         Width           =   972
      End
      Begin VB.TextBox txtRangeFact 
         Height          =   288
         Left            =   4920
         TabIndex        =   60
         Top             =   120
         Width           =   1092
      End
      Begin VB.TextBox txtReadDelay 
         Height          =   285
         Left            =   4920
         TabIndex        =   61
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox txtSusceptibilityMomentFactorCGS 
         Height          =   288
         Left            =   3840
         TabIndex        =   62
         Top             =   1680
         Width           =   972
      End
      Begin VB.TextBox txtSusceptibilitySettings 
         Height          =   288
         Left            =   3840
         TabIndex        =   63
         Top             =   2160
         Width           =   972
      End
      Begin VB.Label Label33 
         Caption         =   "X Calibration factor:"
         Height          =   255
         Left            =   240
         TabIndex        =   64
         Top             =   120
         Width           =   1575
      End
      Begin VB.Label Label34 
         Caption         =   "Y Calibration factor:"
         Height          =   255
         Left            =   240
         TabIndex        =   65
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label35 
         Caption         =   "Z Calibration factor:"
         Height          =   375
         Left            =   240
         TabIndex        =   66
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label30 
         Caption         =   "Range factor:"
         Height          =   255
         Left            =   3600
         TabIndex        =   67
         Top             =   120
         Width           =   1575
      End
      Begin VB.Label Label19 
         Caption         =   "Read delay"
         Height          =   255
         Left            =   3600
         TabIndex        =   68
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label18 
         Caption         =   "Moment Susceptibility Calibration factor (CGS):"
         Height          =   255
         Left            =   240
         TabIndex        =   69
         Top             =   1800
         Width           =   3975
      End
      Begin VB.Label Label42 
         Caption         =   "Susceptibility COM settings:"
         Height          =   255
         Left            =   240
         TabIndex        =   70
         Top             =   2160
         Width           =   3975
      End
   End
   Begin VB.Frame frameOptions 
      BorderStyle     =   0  'None
      Caption         =   "AF Demagnetization"
      Height          =   3372
      Index           =   3
      Left            =   240
      TabIndex        =   71
      Top             =   480
      Width           =   6492
      Begin VB.ComboBox cmbAFAxialCoord 
         Height          =   315
         Left            =   1680
         TabIndex        =   72
         Text            =   "Combo1"
         Top             =   120
         Width           =   612
      End
      Begin VB.TextBox txtAFAxialMax 
         Height          =   288
         Left            =   1680
         TabIndex        =   73
         Top             =   480
         Width           =   1212
      End
      Begin VB.TextBox txtAFAxialMin 
         Height          =   288
         Left            =   1680
         TabIndex        =   74
         Top             =   840
         Width           =   1212
      End
      Begin VB.ComboBox cmbAFDelay 
         Height          =   315
         Left            =   1680
         TabIndex        =   75
         Text            =   "Combo1"
         Top             =   1200
         Width           =   732
      End
      Begin VB.TextBox txtCalibAxialEdit 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   2640
         TabIndex        =   76
         Text            =   "Text1"
         Top             =   2880
         Visible         =   0   'False
         Width           =   495
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdCalibAxial 
         Height          =   1455
         Left            =   240
         TabIndex        =   77
         Top             =   1920
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   2566
         _Version        =   393216
         Rows            =   26
         FixedCols       =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin VB.ComboBox cmbAFTransCoord 
         Height          =   315
         Left            =   4800
         TabIndex        =   78
         Text            =   "Combo1"
         Top             =   120
         Width           =   612
      End
      Begin VB.TextBox txtAFTransMax 
         Height          =   288
         Left            =   4800
         TabIndex        =   79
         Top             =   480
         Width           =   1212
      End
      Begin VB.TextBox txtAFTransMin 
         Height          =   288
         Left            =   4800
         TabIndex        =   80
         Top             =   840
         Width           =   1212
      End
      Begin VB.ComboBox cmbAFRampRate 
         Height          =   315
         Left            =   4800
         TabIndex        =   81
         Top             =   1200
         Width           =   732
      End
      Begin VB.TextBox txtCalibTransEdit 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   5880
         TabIndex        =   82
         Text            =   "Text1"
         Top             =   3000
         Visible         =   0   'False
         Width           =   495
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdCalibTrans 
         Height          =   1455
         Left            =   3600
         TabIndex        =   83
         Top             =   1920
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   2566
         _Version        =   393216
         Rows            =   26
         FixedCols       =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label Label2 
         Caption         =   "Axial Coordinate:"
         Height          =   255
         Left            =   240
         TabIndex        =   84
         Top             =   120
         Width           =   1575
      End
      Begin VB.Label Label40 
         Caption         =   "Axial Max:"
         Height          =   255
         Left            =   240
         TabIndex        =   85
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label16 
         Caption         =   "Axial Min:"
         Height          =   255
         Left            =   240
         TabIndex        =   86
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "AF Delay:"
         Height          =   255
         Left            =   240
         TabIndex        =   87
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         Caption         =   "Axial Calibration Factors"
         Height          =   255
         Left            =   240
         TabIndex        =   88
         Top             =   1680
         Width           =   2295
      End
      Begin VB.Label Label3 
         Caption         =   "Trans. coordinate:"
         Height          =   255
         Left            =   3360
         TabIndex        =   89
         Top             =   120
         Width           =   1575
      End
      Begin VB.Label Label46 
         Caption         =   "Trans. Max:"
         Height          =   255
         Left            =   3360
         TabIndex        =   90
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label12 
         Caption         =   "Trans. Min:"
         Height          =   255
         Left            =   3360
         TabIndex        =   91
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label5 
         Caption         =   "AF Ramp Rate:"
         Height          =   255
         Left            =   3360
         TabIndex        =   92
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label32 
         Alignment       =   2  'Center
         Caption         =   "Trans. Calibration Factors"
         Height          =   255
         Left            =   3600
         TabIndex        =   93
         Top             =   1680
         Width           =   2295
      End
   End
   Begin VB.Frame frameOptions 
      BorderStyle     =   0  'None
      Caption         =   "IRM"
      Height          =   3372
      Index           =   4
      Left            =   240
      TabIndex        =   94
      Top             =   480
      Width           =   6492
      Begin VB.TextBox txtPulseMCCVoltConversion 
         Height          =   288
         Left            =   1680
         TabIndex        =   95
         Top             =   120
         Width           =   1212
      End
      Begin VB.TextBox txtPulseReturnMCCVoltConversion 
         Height          =   288
         Left            =   1680
         TabIndex        =   96
         Top             =   480
         Width           =   1212
      End
      Begin VB.TextBox txtPulseVoltMax 
         Height          =   288
         Left            =   1680
         TabIndex        =   97
         Top             =   840
         Width           =   1212
      End
      Begin VB.TextBox txtPulseLFMax 
         Height          =   288
         Left            =   1680
         TabIndex        =   98
         Top             =   1200
         Width           =   1212
      End
      Begin VB.TextBox txtPulseLFMin 
         Height          =   288
         Left            =   1680
         TabIndex        =   99
         Top             =   1560
         Width           =   1212
      End
      Begin VB.TextBox txtCalibIRMLFEdit 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   2760
         TabIndex        =   100
         Text            =   "Text1"
         Top             =   2880
         Visible         =   0   'False
         Width           =   495
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdCalibIRMLF 
         Height          =   1215
         Left            =   360
         TabIndex        =   101
         Top             =   2160
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   2143
         _Version        =   393216
         Rows            =   21
         FixedCols       =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin VB.ComboBox cmbIRMHFAxis 
         Height          =   315
         Left            =   4800
         TabIndex        =   102
         Text            =   "Combo1"
         Top             =   0
         Width           =   612
      End
      Begin VB.ComboBox cmbIRMLFAxis 
         Height          =   315
         Left            =   4800
         TabIndex        =   103
         Text            =   "Combo1"
         Top             =   360
         Width           =   612
      End
      Begin VB.ComboBox cmbIRMLFBackfieldAxis 
         Height          =   315
         Left            =   4800
         TabIndex        =   104
         Text            =   "Combo1"
         Top             =   720
         Width           =   612
      End
      Begin VB.TextBox txtPulseHFMax 
         Height          =   288
         Left            =   4800
         TabIndex        =   105
         Top             =   1200
         Width           =   1212
      End
      Begin VB.TextBox txtPulseHFMin 
         Height          =   288
         Left            =   4800
         TabIndex        =   106
         Top             =   1560
         Width           =   1212
      End
      Begin VB.TextBox txtCalibIRMHFEdit 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   5640
         TabIndex        =   107
         Text            =   "Text1"
         Top             =   2880
         Visible         =   0   'False
         Width           =   495
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdCalibIRMHF 
         Height          =   1215
         Left            =   3240
         TabIndex        =   108
         Top             =   2160
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   2143
         _Version        =   393216
         Rows            =   21
         FixedCols       =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin VB.Label Label44 
         Caption         =   "IRM Volts/MC Volts:"
         Height          =   375
         Left            =   120
         TabIndex        =   109
         Top             =   120
         Width           =   1575
      End
      Begin VB.Label Label50 
         Caption         =   "Ret. Volts/MC Volts:"
         Height          =   375
         Left            =   120
         TabIndex        =   110
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label51 
         Caption         =   "Pulse Volt Max:"
         Height          =   375
         Left            =   120
         TabIndex        =   111
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label52 
         Caption         =   "IRM LF Pulse Max:"
         Height          =   375
         Left            =   120
         TabIndex        =   112
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label53 
         Caption         =   "IRM LF Pulse Min:"
         Height          =   375
         Left            =   120
         TabIndex        =   113
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label Label38 
         Alignment       =   2  'Center
         Caption         =   "IRM LF Calibration Factors"
         Height          =   255
         Left            =   360
         TabIndex        =   114
         Top             =   1920
         Width           =   2295
      End
      Begin VB.Label Label36 
         Caption         =   "IRM HF Axis:"
         Height          =   375
         Left            =   3240
         TabIndex        =   115
         Top             =   0
         Width           =   1575
      End
      Begin VB.Label Label37 
         Caption         =   "IRM LF Axis:"
         Height          =   375
         Left            =   3240
         TabIndex        =   116
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label43 
         Caption         =   "IRM LF Backfield:"
         Height          =   375
         Left            =   3240
         TabIndex        =   117
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label58 
         Caption         =   "IRM HF Pulse Max:"
         Height          =   375
         Left            =   3240
         TabIndex        =   118
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label59 
         Caption         =   "IRM HF Pulse Min:"
         Height          =   375
         Left            =   3240
         TabIndex        =   119
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label Label41 
         Alignment       =   2  'Center
         Caption         =   "IRM HF Calibration Factors"
         Height          =   255
         Left            =   3240
         TabIndex        =   120
         Top             =   1920
         Width           =   2295
      End
   End
   Begin VB.Frame frameOptions 
      BorderStyle     =   0  'None
      Caption         =   "ARM"
      Height          =   3372
      Index           =   5
      Left            =   240
      TabIndex        =   121
      Top             =   480
      Width           =   6492
      Begin VB.TextBox txtARMMax 
         Height          =   288
         Left            =   1560
         TabIndex        =   122
         Top             =   120
         Width           =   1212
      End
      Begin VB.TextBox txtARMVoltGauss 
         Height          =   288
         Left            =   1560
         TabIndex        =   123
         Top             =   480
         Width           =   1212
      End
      Begin VB.TextBox txtARMVoltMax 
         Height          =   288
         Left            =   1560
         TabIndex        =   124
         Top             =   840
         Width           =   1212
      End
      Begin VB.Label Label39 
         Caption         =   "ARM Max:"
         Height          =   255
         Left            =   0
         TabIndex        =   125
         Top             =   120
         Width           =   1575
      End
      Begin VB.Label Label45 
         Caption         =   "ARM Volt/Gauss:"
         Height          =   255
         Left            =   0
         TabIndex        =   126
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label47 
         Caption         =   "ARM Volt Max:"
         Height          =   375
         Left            =   0
         TabIndex        =   127
         Top             =   840
         Width           =   1575
      End
   End
   Begin VB.Frame frameOptions 
      BorderStyle     =   0  'None
      Caption         =   "Vacuum"
      Height          =   3372
      Index           =   6
      Left            =   240
      TabIndex        =   128
      Top             =   480
      Width           =   6492
      Begin VB.TextBox txtMotorToggle 
         Height          =   288
         Left            =   4080
         TabIndex        =   129
         Top             =   375
         Width           =   255
      End
      Begin VB.TextBox txtVacuumToggleA 
         Height          =   288
         Left            =   4080
         TabIndex        =   130
         Top             =   735
         Width           =   255
      End
      Begin VB.CheckBox chkDoVacuumReset 
         Caption         =   "Reset vacuum on startup"
         Height          =   192
         Left            =   0
         TabIndex        =   131
         Top             =   240
         Width           =   2292
      End
      Begin VB.Label Label68 
         Caption         =   "DIO line assignments:"
         Height          =   255
         Left            =   4080
         TabIndex        =   132
         Top             =   120
         Width           =   1800
      End
      Begin VB.Label Label69 
         Caption         =   "Motor Toggle"
         Height          =   255
         Left            =   4440
         TabIndex        =   133
         Top             =   420
         Width           =   1695
      End
      Begin VB.Label Label70 
         Caption         =   "Vacuum Toggle A"
         Height          =   255
         Left            =   4440
         TabIndex        =   134
         Top             =   780
         Width           =   1695
      End
   End
   Begin VB.Frame frameOptions 
      BorderStyle     =   0  'None
      Caption         =   "SQUID Magnetometer Calibration factors"
      Height          =   3372
      Index           =   7
      Left            =   240
      TabIndex        =   135
      Top             =   480
      Width           =   6492
      Begin VB.TextBox txtAFWait 
         Height          =   285
         Left            =   2280
         TabIndex        =   175
         Top             =   120
         Width           =   375
      End
      Begin VB.TextBox txtTmax 
         Height          =   285
         Left            =   2160
         TabIndex        =   171
         Top             =   3120
         Width           =   375
      End
      Begin VB.TextBox txtThot 
         Height          =   285
         Left            =   1200
         TabIndex        =   170
         Top             =   3120
         Width           =   375
      End
      Begin VB.TextBox txtToffset 
         Height          =   285
         Left            =   2160
         TabIndex        =   169
         Top             =   2760
         Width           =   615
      End
      Begin VB.TextBox txtTSlope 
         Height          =   285
         Left            =   1440
         TabIndex        =   167
         Top             =   2760
         Width           =   495
      End
      Begin VB.TextBox txtAnalogT2 
         Height          =   285
         Left            =   3960
         TabIndex        =   164
         Top             =   1320
         Width           =   255
      End
      Begin VB.TextBox txtAnalogT1 
         Height          =   285
         Left            =   3480
         TabIndex        =   163
         Top             =   1320
         Width           =   255
      End
      Begin VB.CheckBox chkEnableT2 
         Caption         =   "T2"
         Height          =   255
         Left            =   240
         TabIndex        =   162
         Top             =   3120
         Width           =   495
      End
      Begin VB.CheckBox chkEnableT1 
         Caption         =   "T1"
         Height          =   255
         Left            =   240
         TabIndex        =   161
         Top             =   2760
         Width           =   495
      End
      Begin VB.CheckBox chkEnableIRM 
         Caption         =   "IRM pulse coil"
         Height          =   375
         Left            =   240
         TabIndex        =   136
         Top             =   240
         Width           =   2415
      End
      Begin VB.CheckBox chkEnableIRMHi 
         Caption         =   "High-field IRM pulse coil"
         Height          =   375
         Left            =   240
         TabIndex        =   137
         Top             =   600
         Width           =   2415
      End
      Begin VB.CheckBox chkEnableARM 
         Caption         =   "ARM biasing field coil"
         Height          =   375
         Left            =   240
         TabIndex        =   138
         Top             =   960
         Width           =   2415
      End
      Begin VB.CheckBox chkEnableAF 
         Caption         =   "AF degaussing coils"
         Height          =   375
         Left            =   240
         TabIndex        =   139
         Top             =   1320
         Width           =   2415
      End
      Begin VB.CheckBox chkEnableSusceptibility 
         Caption         =   "Susceptibility Coil"
         Height          =   375
         Left            =   240
         TabIndex        =   140
         Top             =   1680
         Width           =   2415
      End
      Begin VB.CheckBox chkEnableIRMBackfield 
         Caption         =   "IRM pulse coil backfield switch"
         Height          =   375
         Left            =   240
         TabIndex        =   141
         Top             =   2040
         Width           =   2775
      End
      Begin VB.CheckBox chkEnableIRMReturn 
         Caption         =   "IRM return voltage"
         Height          =   375
         Left            =   240
         TabIndex        =   142
         Top             =   2400
         Width           =   2775
      End
      Begin VB.TextBox txtARMVoltageOut 
         Height          =   288
         Left            =   4440
         TabIndex        =   143
         Top             =   360
         Width           =   255
      End
      Begin VB.TextBox txtIRMVoltageOut 
         Height          =   288
         Left            =   4440
         TabIndex        =   144
         Top             =   720
         Width           =   255
      End
      Begin VB.TextBox txtIRMCapacitorVoltageIn 
         Height          =   288
         Left            =   4440
         TabIndex        =   145
         Top             =   1320
         Width           =   255
      End
      Begin VB.TextBox txtARMSet 
         Height          =   288
         Left            =   4440
         TabIndex        =   146
         Top             =   1920
         Width           =   255
      End
      Begin VB.TextBox txtIRMFire 
         Height          =   288
         Left            =   4440
         TabIndex        =   147
         Top             =   2280
         Width           =   255
      End
      Begin VB.TextBox txtIRMTrim 
         Height          =   288
         Left            =   4440
         TabIndex        =   148
         Top             =   2640
         Width           =   255
      End
      Begin VB.TextBox txtIRMReady 
         Height          =   288
         Left            =   4440
         TabIndex        =   149
         Top             =   3000
         Width           =   255
      End
      Begin VB.Label Label77 
         Caption         =   "s before AF alert"
         Height          =   255
         Left            =   2640
         TabIndex        =   176
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label Label76 
         Caption         =   "Hot:"
         Height          =   255
         Left            =   840
         TabIndex        =   174
         Top             =   3120
         Width           =   375
      End
      Begin VB.Label Label75 
         Caption         =   "°C Max:"
         Height          =   255
         Left            =   1560
         TabIndex        =   173
         Top             =   3120
         Width           =   615
      End
      Begin VB.Label Label74 
         Caption         =   "°C"
         Height          =   255
         Left            =   2520
         TabIndex        =   172
         Top             =   3120
         Width           =   255
      End
      Begin VB.Label Label73 
         Caption         =   "Volts x              -"
         Height          =   255
         Left            =   960
         TabIndex        =   168
         Top             =   2760
         Width           =   1575
      End
      Begin VB.Label Label72 
         Caption         =   "T2"
         Height          =   255
         Left            =   4200
         TabIndex        =   166
         Top             =   1350
         Width           =   255
      End
      Begin VB.Label Label71 
         Caption         =   "T1"
         Height          =   255
         Left            =   3720
         TabIndex        =   165
         Top             =   1350
         Width           =   255
      End
      Begin VB.Label Label7 
         Caption         =   "Analog channel output:"
         Height          =   255
         Left            =   4440
         TabIndex        =   150
         Top             =   120
         Width           =   1795
      End
      Begin VB.Label Label8 
         Caption         =   "ARM Voltage Out"
         Height          =   255
         Left            =   4800
         TabIndex        =   151
         Top             =   405
         Width           =   1695
      End
      Begin VB.Label Label9 
         Caption         =   "IRM Voltage Out"
         Height          =   255
         Left            =   4800
         TabIndex        =   152
         Top             =   765
         Width           =   1695
      End
      Begin VB.Label Label61 
         Caption         =   "Analog input:"
         Height          =   255
         Left            =   4440
         TabIndex        =   153
         Top             =   1035
         Width           =   1795
      End
      Begin VB.Label Label62 
         Caption         =   "IRMCapacitorVoltageIn"
         Height          =   255
         Left            =   4800
         TabIndex        =   154
         Top             =   1350
         Width           =   1695
      End
      Begin VB.Label Label63 
         Caption         =   "DIO line assignments:"
         Height          =   255
         Left            =   4440
         TabIndex        =   155
         Top             =   1665
         Width           =   1795
      End
      Begin VB.Label Label64 
         Caption         =   "ARM Set"
         Height          =   255
         Left            =   4800
         TabIndex        =   156
         Top             =   1965
         Width           =   1695
      End
      Begin VB.Label Label65 
         Caption         =   "IRM Fire"
         Height          =   255
         Left            =   4800
         TabIndex        =   157
         Top             =   2325
         Width           =   1695
      End
      Begin VB.Label Label66 
         Caption         =   "IRM Trim"
         Height          =   255
         Left            =   4800
         TabIndex        =   158
         Top             =   2685
         Width           =   1695
      End
      Begin VB.Label Label67 
         Caption         =   "IRM Ready"
         Height          =   255
         Left            =   4800
         TabIndex        =   159
         Top             =   3045
         Width           =   1695
      End
   End
   Begin ComctlLib.TabStrip tbsOptions 
      Height          =   3852
      Left            =   120
      TabIndex        =   160
      Top             =   120
      Width           =   6732
      _ExtentX        =   11880
      _ExtentY        =   6800
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   8
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "DC &Motors (1)"
            Key             =   "DCMotors"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Motor settings"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "DC Motors (&2)"
            Key             =   "MotorCommands"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Motor command settings"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Ma&gnetometer"
            Key             =   "Squid"
            Object.Tag             =   ""
            Object.ToolTipText     =   "SQUID magnetometer and susceptibility bridge settings"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&AF Demag"
            Key             =   "AF"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Alternating field demagnetization settings"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&IRM"
            Key             =   "IRM"
            Object.Tag             =   ""
            Object.ToolTipText     =   "IRM settings"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "ARM"
            Key             =   "ARM"
            Object.Tag             =   ""
            Object.ToolTipText     =   "ARM settings"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab7 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Vacuum"
            Key             =   "Vacuum"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Vacuum settings"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab8 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Modules"
            Key             =   "Modules"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Modules"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim activeRow As Integer
Dim activeCol As Integer

Private Sub importSettings()
    Dim i As Integer
    txtZeroPos = ZeroPos
    txtMeasPos = MeasPos
    txtAFPos = AFPos
    txtIRMPos = IRMPos
    txtSCoilPos = SCoilPos
    txtSampleBottom = SampleBottom
    txtSampleTop = SampleTop
    txtSampHoleAlignOffset = SampleHoleAlignmentOffset
    txtLiftSpeedSlow = LiftSpeedSlow
    txtLiftSpeedNormal = LiftSpeedNormal
    txtLiftSpeedFast = LiftSpeedFast
    txtChangerSpeed = ChangerSpeed
    txtTurnerSpeed = TurnerSpeed
    txtSCurveFactor = SCurveFactor
    txtTurningMotorFullRotation = TurningMotorFullRotation
    txtTurningMotor1rps = TurningMotor1rps
    txtUpDownMotor1cm = UpDownMotor1cm
    txtTrayOffsetAngle = TrayOffsetAngle
    txtPickupTorqueThrottle = PickupTorqueThrottle
    txtUpDownTorquefactor = UpDownTorqueFactor
    txtZCal = ZCal
    txtXCal = XCal
    txtYCal = YCal
    txtRangeFact = RangeFact
    txtReadDelay = ReadDelay ' (March 2008 L Carporzen) Read delay
    txtSusceptibilityMomentFactorCGS = SusceptibilityMomentFactorCGS
    txtSusceptibilitySettings = SusceptibilitySettings
    cmbAFDelay = Str$(AFDelay)
    cmbAFRampRate = Str$(AFRampRate)
    txtAFWait = AFWait
    txtTSlope = TSlope
    txtToffset = Toffset
    txtThot = Thot
    txtTmax = Tmax
    cmbAFAxialCoord = AfAxialCoord
    txtAFAxialMax = AfAxialMax
    txtAFAxialMin = AfAxialMin
    cmbAFTransCoord = AfTransCoord
    txtAFTransMax = AfTransMax
    txtAFTransMin = AfTransMin
    cmbIRMHFAxis = IRMHFAxis
    cmbIRMLFAxis = IRMLFAxis
    cmbIRMLFBackfieldAxis = IRMLFBackfieldAxis
    txtPulseLFMax = PulseLFMax
    txtPulseLFMin = PulseLFMin
    txtPulseHFMax = PulseHFMax
    txtPulseHFMin = PulseHFMin
    txtPulseVoltMax = PulseVoltMax
    txtPulseMCCVoltConversion = PulseMCCVoltConversion
    txtPulseReturnMCCVoltConversion = PulseReturnMCCVoltConversion
    txtARMMax = ARMMax
    txtARMVoltGauss = ARMVoltGauss
    txtARMVoltMax = ARMVoltMax
    ' (March 2008 L Carporzen) Put in Settings the IRM/ARM channels (MIT acquisition board does not work on IRMTrim = 3
    ' Analog channel output
    txtARMVoltageOut = ARMVoltageOut
    txtIRMVoltageOut = IRMVoltageOut
    ' Analog input
    txtIRMCapacitorVoltageIn = IRMCapacitorVoltageIn
    txtAnalogT1 = AnalogT1
    txtAnalogT2 = AnalogT2
    ' DIO line assignments
    txtARMSet = ARMSet
    txtIRMFire = IRMFire
    txtIRMTrim = IRMTrim
    txtIRMReady = IRMReady
    txtMotorToggle = MotorToggle
    txtVacuumToggleA = VacuumToggleA
'    txtVacuumToggleB = VacuumToggleB
    txtCmdHometoTop = CmdHomeToTop
    txtCmdSamplePickup = CmdSamplePickup
    txtMotorIDTurning = MotorIDTurning
    txtMotorIDChanger = MotorIDChanger
    txtMotorIDUpDown = MotorIDUpDown
    If DoVacuumReset Then chkDoVacuumReset.value = 1 Else chkDoVacuumReset.value = 0
    If EnableIRM Then chkEnableIRM.value = 1 Else chkEnableIRM.value = 0
    If EnableIRMHi Then chkEnableIRMHi.value = 1 Else chkEnableIRMHi.value = 0
    If EnableIRMBackfield Then chkEnableIRMBackfield.value = 1 Else chkEnableIRMBackfield.value = 0
    If EnableIRMReturn Then chkEnableIRMReturn.value = 1 Else chkEnableIRMReturn.value = 0
    If EnableARM Then chkEnableARM.value = 1 Else chkEnableARM.value = 0
    If EnableAF Then chkEnableAF.value = 1 Else chkEnableAF.value = 0
    If EnableT1 Then chkEnableT1.value = 1 Else chkEnableT1.value = 0
    If EnableT2 Then chkEnableT2.value = 1 Else chkEnableT2.value = 0
    If EnableSusceptibility Then chkEnableSusceptibility.value = 1 Else chkEnableSusceptibility.value = 0
    For i = 1 To 25
        grdCalibTrans.TextMatrix(i, 0) = Format$(AFTransX(i), "0")
        grdCalibTrans.TextMatrix(i, 1) = Format$(AFTransY(i), "0")
        grdCalibAxial.TextMatrix(i, 0) = Format$(AFAxialX(i), "0")
        grdCalibAxial.TextMatrix(i, 1) = Format$(AFAxialY(i), "0")
    Next i
    For i = 1 To 20
        grdCalibIRMHF.TextMatrix(i, 0) = Format$(PulseHFX(i), "0")
        grdCalibIRMHF.TextMatrix(i, 1) = Format$(PulseHFY(i), "0")
        grdCalibIRMLF.TextMatrix(i, 0) = Format$(PulseLFX(i), "0")
        grdCalibIRMLF.TextMatrix(i, 1) = Format$(PulseLFY(i), "0")
    Next i
End Sub

Private Sub applySettings()
        Dim i As Integer
        ZeroPos = Int(val(txtZeroPos))
        MeasPos = Int(val(txtMeasPos))
        AFPos = Int(val(txtAFPos))
        IRMPos = Int(val(txtIRMPos))
        SCoilPos = Int(val(txtSCoilPos))
        SampleBottom = Int(val(txtSampleBottom))
        SampleTop = Int(val(txtSampleTop))
        SampleHoleAlignmentOffset = val(txtSampHoleAlignOffset)
        LiftSpeedSlow = Int(val(txtLiftSpeedSlow))
        LiftSpeedNormal = Int(val(txtLiftSpeedNormal))
        LiftSpeedFast = Int(val(txtLiftSpeedFast))
        ChangerSpeed = Int(val(txtChangerSpeed))
        TurnerSpeed = Int(val(txtTurnerSpeed))
        SCurveFactor = Int(val(txtSCurveFactor))
        TurningMotorFullRotation = Int(val(txtTurningMotorFullRotation))
        TurningMotor1rps = Int(val(txtTurningMotor1rps))
        UpDownMotor1cm = val(txtUpDownMotor1cm)
        TrayOffsetAngle = val(txtTrayOffsetAngle)
        PickupTorqueThrottle = val(txtPickupTorqueThrottle)
        UpDownTorqueFactor = val(txtUpDownTorquefactor)
        ZCal = val(txtZCal)
        XCal = val(txtXCal)
        YCal = val(txtYCal)
        RangeFact = val(txtRangeFact)
        ReadDelay = Int(val(txtReadDelay)) ' (March 2008 L Carporzen) Read delay
        SusceptibilitySettings = txtSusceptibilitySettings
        SusceptibilityMomentFactorCGS = val(txtSusceptibilityMomentFactorCGS)
        AFDelay = val(cmbAFDelay)
        AFRampRate = val(cmbAFRampRate)
        AFWait = val(txtAFWait)
        TSlope = val(txtTSlope)
        Toffset = val(txtToffset)
        Thot = val(txtThot)
        Tmax = val(txtTmax)
        AfAxialCoord = cmbAFAxialCoord
        AfAxialMax = val(txtAFAxialMax)
        AfAxialMin = val(txtAFAxialMin)
        AfTransCoord = cmbAFTransCoord
        AfTransMin = val(txtAFTransMin)
        AfTransMax = val(txtAFTransMax)
        IRMHFAxis = cmbIRMHFAxis
        IRMLFAxis = cmbIRMLFAxis
        IRMLFBackfieldAxis = cmbIRMLFBackfieldAxis
        PulseLFMax = val(txtPulseLFMax)
        PulseLFMin = val(txtPulseLFMin)
        PulseHFMax = val(txtPulseHFMax)
        PulseHFMin = val(txtPulseHFMin)
        PulseVoltMax = val(txtPulseVoltMax)
        PulseReturnMCCVoltConversion = val(txtPulseReturnMCCVoltConversion)
        PulseMCCVoltConversion = txtPulseMCCVoltConversion
        ARMMax = val(txtARMMax)
        ARMVoltGauss = val(txtARMVoltGauss)
        ARMVoltMax = val(txtARMVoltMax)
        ' (March 2008 L Carporzen) Put in Settings the IRM/ARM channels (MIT acquisition board does not work on IRMTrim = 3
        ' Analog channel output
        ARMVoltageOut = val(txtARMVoltageOut)
        IRMVoltageOut = val(txtIRMVoltageOut)
        ' Analog input
        IRMCapacitorVoltageIn = val(txtIRMCapacitorVoltageIn)
        AnalogT1 = val(txtAnalogT1)
        AnalogT2 = val(txtAnalogT2)
        ' DIO line assignments
        ARMSet = val(txtARMSet)
        IRMFire = val(txtIRMFire)
        IRMTrim = val(txtIRMTrim)
        IRMReady = val(txtIRMReady)
        MotorToggle = val(txtMotorToggle)
        VacuumToggleA = val(txtVacuumToggleA)
'        VacuumToggleB = val(txtVacuumToggleB)
        CmdHomeToTop = Int(val(txtCmdHometoTop))
        CmdSamplePickup = Int(val(txtCmdSamplePickup))
        MotorIDTurning = Int(val(txtMotorIDTurning))
        MotorIDChanger = Int(val(txtMotorIDChanger))
        MotorIDUpDown = Int(val(txtMotorIDUpDown))
        DoVacuumReset = chkDoVacuumReset.value
        EnableIRM = chkEnableIRM.value
        EnableIRMHi = chkEnableIRMHi.value
        EnableIRMBackfield = chkEnableIRMBackfield.value
        EnableIRMReturn = chkEnableIRMReturn.value
        EnableARM = chkEnableARM.value
        EnableAF = chkEnableAF.value
        EnableT1 = chkEnableT1.value
        EnableT2 = chkEnableT2.value
        EnableSusceptibility = chkEnableSusceptibility.value
        For i = 1 To 25
            AFTransX(i) = val(grdCalibTrans.TextMatrix(i, 0))
            AFTransY(i) = val(grdCalibTrans.TextMatrix(i, 1))
            AFAxialX(i) = val(grdCalibAxial.TextMatrix(i, 0))
            AFAxialY(i) = val(grdCalibAxial.TextMatrix(i, 1))
        Next i
        For i = 1 To 20
            PulseLFX(i) = val(grdCalibIRMLF.TextMatrix(i, 0))
            PulseLFY(i) = val(grdCalibIRMLF.TextMatrix(i, 1))
            PulseHFX(i) = val(grdCalibIRMHF.TextMatrix(i, 0))
            PulseHFY(i) = val(grdCalibIRMHF.TextMatrix(i, 1))
        Next i
End Sub

Private Sub cmdApply_Click()
    Dim response As VbMsgBoxResult
    response = MsgBox("Incorrect values can break the system!" & vbCrLf & "Are you sure you want to make changes?", vbYesNo, "Warning!")
    If response = vbYes Then
        applySettings
    End If
    importSettings
    Me.refresh
End Sub

Private Sub cmdcancel_click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    Dim response As VbMsgBoxResult
    response = MsgBox("Incorrect values can break the system!" & vbCrLf & "Are you sure you want to make changes?", vbYesNo, "Warning!")
    If response = vbYes Then
        applySettings
        importSettings
        Config_writeSettingstoINI
        Unload Me
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    'handle ctrl+tab to move to the next tab
    If Shift = vbCtrlMask And KeyCode = vbKeyTab Then
        i = tbsOptions.SelectedItem.Index
        If i = tbsOptions.Tabs.Count Then
            'last tab so we need to wrap to tab 1
            Set tbsOptions.SelectedItem = tbsOptions.Tabs(1)
        Else
            'increment the tab
            Set tbsOptions.SelectedItem = tbsOptions.Tabs(i + 1)
        End If
    End If
End Sub

Private Sub Form_Load()
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
    cmbAFRampRate.Clear
    cmbAFDelay.Clear
    cmbAFAxialCoord.Clear
    cmbAFTransCoord.Clear
    cmbAFRampRate.AddItem "3"
    cmbAFRampRate.AddItem "5"
    cmbAFRampRate.AddItem "7"
    cmbAFRampRate.AddItem "9"
    cmbAFDelay.AddItem "1"
    cmbAFDelay.AddItem "2"
    cmbAFDelay.AddItem "3"
    cmbAFDelay.AddItem "4"
    cmbAFDelay.AddItem "5"
    cmbAFDelay.AddItem "6"
    cmbAFDelay.AddItem "7"
    cmbAFDelay.AddItem "8"
    cmbAFDelay.AddItem "9"
    cmbAFAxialCoord.AddItem "X"
    cmbAFAxialCoord.AddItem "Y"
    cmbAFAxialCoord.AddItem "Z"
    cmbAFTransCoord.AddItem "X"
    cmbAFTransCoord.AddItem "Y"
    cmbAFTransCoord.AddItem "Z"
    cmbIRMHFAxis.AddItem "X"
    cmbIRMHFAxis.AddItem "Y"
    cmbIRMHFAxis.AddItem "Z"
    cmbIRMLFAxis.AddItem "X"
    cmbIRMLFAxis.AddItem "Y"
    cmbIRMLFAxis.AddItem "Z"
    cmbIRMLFBackfieldAxis.AddItem "X"
    cmbIRMLFBackfieldAxis.AddItem "Y"
    cmbIRMLFBackfieldAxis.AddItem "Z"
    grdCalibAxial.TextMatrix(0, 0) = "2G"
    grdCalibAxial.TextMatrix(0, 1) = "Gauss"
    grdCalibTrans.TextMatrix(0, 0) = "2G"
    grdCalibTrans.TextMatrix(0, 1) = "Gauss"
    grdCalibIRMLF.TextMatrix(0, 0) = "V"
    grdCalibIRMLF.TextMatrix(0, 1) = "Gauss"
    grdCalibIRMHF.TextMatrix(0, 0) = "V"
    grdCalibIRMHF.TextMatrix(0, 1) = "Gauss"
    importSettings
    selectTab 1
End Sub

Private Sub selectTab(tabtoselect As Integer)
    Dim i As Integer
    'show and enable the selected tab's controls
    'and hide and disable all others
    For i = 0 To tbsOptions.Tabs.Count - 1
        If i = tabtoselect - 1 Then
            'frameOptions(i).Left = 210
            'frameOptions(i).visible = true
            frameOptions(i).Enabled = True
            frameOptions(i).Visible = True
            frameOptions(i).ZOrder 0
        Else
            'frameOptions(i).Left = -20000
            frameOptions(i).Visible = False
            frameOptions(i).Enabled = False
        End If
    Next
End Sub

Private Sub tbsOptions_Click()
    selectTab tbsOptions.SelectedItem.Index
End Sub

Private Sub grdCalibIRMLF_Click()
    activeRow = grdCalibIRMLF.row
    activeCol = grdCalibIRMLF.Col
    txtCalibIRMLFEdit.Visible = True
    txtCalibIRMLFEdit.Width = grdCalibIRMLF.CellWidth - 2
    txtCalibIRMLFEdit.Height = grdCalibIRMLF.CellHeight - 2
    txtCalibIRMLFEdit.Top = grdCalibIRMLF.CellTop + grdCalibIRMLF.Top
    txtCalibIRMLFEdit.Left = grdCalibIRMLF.CellLeft + grdCalibIRMLF.Left
    txtCalibIRMLFEdit.Text = grdCalibIRMLF.Text
    txtCalibIRMLFEdit.SelStart = 0
    txtCalibIRMLFEdit.SelLength = Len(txtCalibIRMLFEdit.Text)
    txtCalibIRMLFEdit.ZOrder
    txtCalibIRMLFEdit.SetFocus
End Sub

Private Sub grdCalibIRMHF_Click()
    activeRow = grdCalibIRMHF.row
    activeCol = grdCalibIRMHF.Col
    txtCalibIRMHFEdit.Visible = True
    txtCalibIRMHFEdit.Width = grdCalibIRMHF.CellWidth - 2
    txtCalibIRMHFEdit.Height = grdCalibIRMHF.CellHeight - 2
    txtCalibIRMHFEdit.Top = grdCalibIRMHF.CellTop + grdCalibIRMHF.Top
    txtCalibIRMHFEdit.Left = grdCalibIRMHF.CellLeft + grdCalibIRMHF.Left
    txtCalibIRMHFEdit.Text = grdCalibIRMHF.Text
    txtCalibIRMHFEdit.SelStart = 0
    txtCalibIRMHFEdit.SelLength = Len(txtCalibIRMHFEdit.Text)
    txtCalibIRMHFEdit.ZOrder
    txtCalibIRMHFEdit.SetFocus
End Sub

Private Sub grdCalibTrans_Click()
    activeRow = grdCalibTrans.row
    activeCol = grdCalibTrans.Col
    txtCalibTransEdit.Visible = True
    txtCalibTransEdit.Width = grdCalibTrans.CellWidth - 2
    txtCalibTransEdit.Height = grdCalibTrans.CellHeight - 2
    txtCalibTransEdit.Top = grdCalibTrans.CellTop + grdCalibTrans.Top
    txtCalibTransEdit.Left = grdCalibTrans.CellLeft + grdCalibTrans.Left
    txtCalibTransEdit.Text = grdCalibTrans.Text
    txtCalibTransEdit.SelStart = 0
    txtCalibTransEdit.SelLength = Len(txtCalibTransEdit.Text)
    txtCalibTransEdit.ZOrder
    txtCalibTransEdit.SetFocus
    
End Sub

Private Sub grdCalibAxial_Click()
    activeRow = grdCalibAxial.row
    activeCol = grdCalibAxial.Col
    txtCalibAxialEdit.Visible = True
    txtCalibAxialEdit.Width = grdCalibAxial.CellWidth - 2
    txtCalibAxialEdit.Height = grdCalibAxial.CellHeight - 2
    txtCalibAxialEdit.Top = grdCalibAxial.CellTop + grdCalibAxial.Top
    txtCalibAxialEdit.Left = grdCalibAxial.CellLeft + grdCalibAxial.Left
    txtCalibAxialEdit.Text = grdCalibAxial.Text
    txtCalibAxialEdit.SelStart = 0
    txtCalibAxialEdit.SelLength = Len(txtCalibAxialEdit.Text)
    txtCalibAxialEdit.ZOrder
    txtCalibAxialEdit.SetFocus
End Sub

Private Sub txtCalibIRMLFEdit_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        grdCalibIRMLF.Text = txtCalibIRMLFEdit.Text
        If grdCalibIRMLF.row = grdCalibIRMLF.Rows - 1 Then
            grdCalibIRMLF.row = grdCalibIRMLF.row
        Else
            grdCalibIRMLF.row = grdCalibIRMLF.row + 1
        End If
        grdCalibIRMLF.SetFocus
        txtCalibIRMLFEdit.Visible = False
    End If
End Sub

Private Sub txtCalibIRMLFEdit_LostFocus()
    If txtCalibIRMLFEdit.Visible Then grdCalibIRMLF.TextMatrix(activeRow, activeCol) = txtCalibIRMLFEdit
    txtCalibIRMLFEdit.Visible = False
End Sub

Private Sub txtCalibIRMHFEdit_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        grdCalibIRMHF.Text = txtCalibIRMHFEdit.Text
        If grdCalibIRMHF.row = grdCalibIRMHF.Rows - 1 Then
            grdCalibIRMHF.row = grdCalibIRMHF.row
        Else
            grdCalibIRMHF.row = grdCalibIRMHF.row + 1
        End If
        grdCalibIRMHF.SetFocus
        txtCalibIRMHFEdit.Visible = False
    End If
End Sub

Private Sub txtCalibIRMHFEdit_LostFocus()
    If txtCalibIRMHFEdit.Visible Then grdCalibIRMHF.TextMatrix(activeRow, activeCol) = txtCalibIRMHFEdit
    txtCalibIRMHFEdit.Visible = False
End Sub

Private Sub txtCalibTransEdit_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        grdCalibTrans.Text = txtCalibTransEdit.Text
        If grdCalibTrans.row = grdCalibTrans.Rows - 1 Then
            grdCalibTrans.row = grdCalibTrans.row
        Else
            grdCalibTrans.row = grdCalibTrans.row + 1
        End If
        grdCalibTrans.SetFocus
        txtCalibTransEdit.Visible = False
    End If
End Sub

Private Sub txtCalibTransEdit_LostFocus()
    If txtCalibTransEdit.Visible Then grdCalibTrans.TextMatrix(activeRow, activeCol) = txtCalibTransEdit
    txtCalibTransEdit.Visible = False
End Sub

Private Sub txtCalibAxialEdit_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        grdCalibAxial.Text = txtCalibAxialEdit.Text
        If grdCalibAxial.row = grdCalibAxial.Rows - 1 Then
            grdCalibAxial.row = grdCalibAxial.row
        Else
            grdCalibAxial.row = grdCalibAxial.row + 1
        End If
        grdCalibAxial.SetFocus
        txtCalibAxialEdit.Visible = False
    End If
End Sub

Private Sub txtCalibAxialEdit_LostFocus()
    If txtCalibAxialEdit.Visible Then grdCalibAxial.TextMatrix(activeRow, activeCol) = txtCalibAxialEdit
    txtCalibAxialEdit.Visible = False
End Sub

Private Sub grdCalibAxial_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call grdCalibAxial_Click
    End If
    If KeyCode = vbKeyDelete Then
        grdCalibAxial.Text = "0"
    End If
End Sub

Private Sub grdCalibTrans_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call grdCalibTrans_Click
    End If
    If KeyCode = vbKeyDelete Then
        grdCalibTrans.Text = "0"
    End If
End Sub

Private Sub grdCalibIRMLF_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call grdCalibIRMLF_Click
    End If
    If KeyCode = vbKeyDelete Then
        grdCalibIRMLF.Text = "0"
    End If
End Sub

Private Sub grdCalibIRMHF_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call grdCalibIRMHF_Click
    End If
    If KeyCode = vbKeyDelete Then
        grdCalibIRMHF.Text = "0"
    End If
End Sub
' (June 2007 L Carporzen) Form to adjust automatically the positions each time a new rod is installed.
' All we need is approximatly good (+-5 cm) positions in Paleomag.INI and the Bartington susceptibility standard (white cylinder placed on top of the MS2B box).
' NB that the susceptibility is very sensitive to the altitude in order to have reproducible measurement.
' The susceptibility position should be easy to define at less than 1 cm prior to running that sequence.
' (March 2008 L Carporzen) It can also correct the negative sample height that some systems are still using.
Private Sub cmdCalibNewRod_Click()
    If UpDownMotor1cm < 0 And ZeroPos < 0 Or UpDownMotor1cm > 0 And ZeroPos > 0 Then
        UpDownMotor1cm = val(InputBox("You can change the sign of the Sample Height " & Format$((SampleTop - SampleBottom) / UpDownMotor1cm, "0.00") & " to positive values by changing the sign of the Up/Down scale " & UpDownMotor1cm & " to " & -UpDownMotor1cm & ". Are you agree?", "Important!", -UpDownMotor1cm))
        If UpDownMotor1cm = 0 Then
            UpDownMotor1cm = frmSettings.txtUpDownMotor1cm
        Else
            frmSettings.txtUpDownMotor1cm = UpDownMotor1cm
            MsgBox ("The default sign of the Sample Height is now " & Format$((SampleTop - SampleBottom) / UpDownMotor1cm, "0.00") & ". Click OK in the settings window to save it.")
        End If
    End If
    SampleHoleAlignmentOffset = 0
    MsgBox "Please load the Bartington susceptibility standard sample face up into the hole 1 of the sample changer, let free the hole 199." + Chr(13) + "The Bartington susceptibility meter should be in CGS units on the range 1.0." + Chr(13) + "We strongly recommand to backup the Paleomag.INI file..."
    frmCalRod.ZOrder
    frmCalRod.Show
    SampleHoleAlignmentOffset = val(txtSampHoleAlignOffset)
End Sub
