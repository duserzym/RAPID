VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmSettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Settings Dialog"
   ClientHeight    =   5985
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   13185
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   13185
   Visible         =   0   'False
   Begin VB.Frame frameOptions 
      BorderStyle     =   0  'None
      Caption         =   "Vacuum"
      Height          =   4695
      Index           =   7
      Left            =   480
      TabIndex        =   68
      Top             =   480
      Width           =   10215
      Begin VB.Frame frameAscIRMOutputVoltageBoostFactors 
         Caption         =   "Output Voltage Boost Factors"
         Height          =   1575
         Left            =   6600
         TabIndex        =   346
         Top             =   1680
         Width           =   3255
         Begin VB.TextBox txtMaxIrmVoltageOut_BoostPercentage 
            Height          =   288
            Left            =   1920
            TabIndex        =   348
            Top             =   960
            Width           =   1095
         End
         Begin VB.TextBox txtMinIrmVoltageOut_BoostPercentage 
            Height          =   288
            Left            =   1920
            TabIndex        =   347
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label lblSettings 
            Caption         =   "% Boost at Maximum Field:"
            Height          =   495
            Index           =   88
            Left            =   240
            TabIndex        =   350
            Top             =   960
            Width           =   1575
         End
         Begin VB.Label lblSettings 
            Caption         =   "% Boost at Minimum Field:"
            Height          =   615
            Index           =   87
            Left            =   240
            TabIndex        =   349
            Top             =   360
            Width           =   1575
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "IRM Trim Settings"
         Height          =   1335
         Left            =   3360
         TabIndex        =   286
         Top             =   3360
         Width           =   3135
         Begin VB.OptionButton optTrimOnFalse 
            Caption         =   "False"
            Height          =   375
            Left            =   1800
            TabIndex        =   289
            Top             =   720
            Width           =   1095
         End
         Begin VB.OptionButton optTrimOnTrue 
            Caption         =   "True"
            Height          =   375
            Left            =   1800
            TabIndex        =   288
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "IRM Trim On State ="
            Height          =   255
            Left            =   120
            TabIndex        =   287
            Top             =   600
            Width           =   1695
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "IRM Transverse"
         Height          =   1575
         Left            =   3360
         TabIndex        =   278
         Top             =   1680
         Width           =   3135
         Begin VB.TextBox txtTransMaxCapVoltage 
            Height          =   288
            Left            =   1800
            TabIndex        =   283
            Top             =   1080
            Width           =   1095
         End
         Begin VB.TextBox txtPulseTransMax 
            Height          =   288
            Left            =   1800
            TabIndex        =   280
            Top             =   360
            Width           =   1095
         End
         Begin VB.TextBox txtPulseTransMin 
            Height          =   288
            Left            =   1800
            TabIndex        =   279
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label lblTransMaxCapVoltage 
            Caption         =   "Max Capacitor Volt:"
            Height          =   255
            Left            =   240
            TabIndex        =   284
            Top             =   1080
            Width           =   1815
         End
         Begin VB.Label lblIRMTransPulseMin 
            Caption         =   "Min Field (G):"
            Height          =   255
            Left            =   240
            TabIndex        =   282
            Top             =   720
            Width           =   1575
         End
         Begin VB.Label lblIRMTransPulseMax 
            Caption         =   "Max Field (G):"
            Height          =   255
            Left            =   240
            TabIndex        =   281
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "IRM Axial"
         Height          =   1575
         Left            =   0
         TabIndex        =   271
         Top             =   1680
         Width           =   3255
         Begin VB.TextBox txtAxialMaxCapVoltage 
            Height          =   288
            Left            =   1920
            TabIndex        =   274
            Top             =   1080
            Width           =   1095
         End
         Begin VB.TextBox txtPulseAxialMax 
            Height          =   288
            Left            =   1920
            TabIndex        =   273
            Top             =   360
            Width           =   1095
         End
         Begin VB.TextBox txtPulseAxialMin 
            Height          =   288
            Left            =   1920
            TabIndex        =   272
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label lblIRMAxialPulseMin 
            Caption         =   "Min Field (G):"
            Height          =   255
            Left            =   240
            TabIndex        =   277
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label lblAxialMaxCapVoltage 
            Caption         =   "Max Capacitor Volt:"
            Height          =   255
            Left            =   240
            TabIndex        =   276
            Top             =   1080
            Width           =   1455
         End
         Begin VB.Label lblIRMAxialPulseMax 
            Caption         =   "Max Field (G):"
            Height          =   255
            Left            =   240
            TabIndex        =   275
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.CommandButton cmdOpenIRMARMForm 
         Caption         =   "Open IRM/ARM Control Window"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   0
         TabIndex        =   270
         Top             =   4320
         Width           =   3135
      End
      Begin VB.ComboBox cmbIRMBackfieldAxis 
         Height          =   315
         Left            =   5160
         TabIndex        =   74
         Top             =   1080
         Width           =   612
      End
      Begin VB.ComboBox cmbIRMAxis 
         Height          =   315
         Left            =   5160
         TabIndex        =   73
         Top             =   600
         Width           =   612
      End
      Begin VB.TextBox txtPulseVoltMax 
         Height          =   288
         Left            =   1920
         TabIndex        =   72
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox txtPulseReturnMCCVoltConversion 
         Height          =   288
         Left            =   1920
         TabIndex        =   71
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox txtPulseMCCVoltConversion 
         Height          =   288
         Left            =   1920
         TabIndex        =   70
         Top             =   600
         Width           =   1095
      End
      Begin VB.CommandButton cmdCalibrateIRMFields 
         Caption         =   "Calibrate IRM Fields"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   0
         TabIndex        =   75
         Top             =   3360
         Width           =   3135
      End
      Begin VB.ComboBox cmbIRMSystem 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1920
         TabIndex        =   69
         Top             =   120
         Width           =   1935
      End
      Begin VB.CommandButton cmdCalibrateIRMVoltages 
         Caption         =   "Calibrate IRM DAQ Voltages"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   0
         TabIndex        =   76
         Top             =   3840
         Width           =   3135
      End
      Begin VB.Label lblPulseVoltsMax 
         Caption         =   "DAQ Output Volt Max:"
         Height          =   375
         Left            =   240
         TabIndex        =   226
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label lblPulseVoltsPerMCCVols 
         Caption         =   "IRM Volts/MC Volts:"
         Height          =   255
         Left            =   240
         TabIndex        =   225
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label lblIRMBackfieldAxis 
         Caption         =   "IRM Backfield Axis:"
         Height          =   255
         Left            =   3600
         TabIndex        =   224
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label lblIRMAxis 
         Caption         =   "Axial IRM Axis:"
         Height          =   255
         Left            =   3600
         TabIndex        =   223
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label lblRetVoltsPerMCCVolts 
         Caption         =   "Ret. Volts/MC Volts:"
         Height          =   255
         Left            =   240
         TabIndex        =   222
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label lblIRMSystem 
         Caption         =   "IRM System:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   221
         Top             =   120
         Width           =   1455
      End
   End
   Begin VB.Frame frameOptions 
      BorderStyle     =   0  'None
      Caption         =   "DC Motor settings"
      Height          =   4455
      Index           =   0
      Left            =   480
      TabIndex        =   6
      Top             =   600
      Width           =   6492
      Begin VB.TextBox txtLiftAcceleration 
         Height          =   288
         Left            =   4800
         TabIndex        =   323
         Top             =   1200
         Width           =   1212
      End
      Begin VB.TextBox txtMinUpDownPos 
         Height          =   288
         Left            =   1680
         TabIndex        =   13
         Top             =   2280
         Width           =   1212
      End
      Begin VB.TextBox txtFloorPos 
         Height          =   288
         Left            =   1680
         TabIndex        =   12
         Top             =   1920
         Width           =   1212
      End
      Begin VB.TextBox txtZeroPos 
         Height          =   288
         Left            =   1680
         TabIndex        =   7
         Top             =   120
         Width           =   1212
      End
      Begin VB.TextBox txtMeasPos 
         Height          =   288
         Left            =   1680
         TabIndex        =   8
         Top             =   480
         Width           =   1212
      End
      Begin VB.TextBox txtAFPos 
         Height          =   288
         Left            =   1680
         TabIndex        =   9
         Top             =   840
         Width           =   1212
      End
      Begin VB.TextBox txtIRMPos 
         Height          =   288
         Left            =   1680
         TabIndex        =   10
         Top             =   1200
         Width           =   1212
      End
      Begin VB.TextBox txtSCoilPos 
         Height          =   288
         Left            =   1680
         TabIndex        =   11
         Top             =   1560
         Width           =   1212
      End
      Begin VB.TextBox txtSampleTop 
         Height          =   288
         Left            =   1680
         TabIndex        =   14
         Top             =   2640
         Width           =   1212
      End
      Begin VB.TextBox txtSampleBottom 
         Height          =   288
         Left            =   1680
         TabIndex        =   15
         Top             =   3000
         Width           =   1212
      End
      Begin VB.TextBox txtTurningMotorFullRotation 
         Height          =   288
         Left            =   1680
         TabIndex        =   16
         Top             =   3360
         Width           =   1212
      End
      Begin VB.TextBox txtTrayOffsetAngle 
         Height          =   288
         Left            =   1680
         TabIndex        =   17
         Top             =   3720
         Width           =   1212
      End
      Begin VB.TextBox txtLiftSpeedSlow 
         Height          =   288
         Left            =   4800
         TabIndex        =   18
         Top             =   120
         Width           =   1212
      End
      Begin VB.TextBox txtLiftSpeedNormal 
         Height          =   288
         Left            =   4800
         TabIndex        =   19
         Top             =   480
         Width           =   1212
      End
      Begin VB.TextBox txtLiftSpeedFast 
         Height          =   288
         Left            =   4800
         TabIndex        =   20
         Top             =   840
         Width           =   1212
      End
      Begin VB.TextBox txtChangerSpeed 
         Height          =   288
         Left            =   4800
         TabIndex        =   21
         Top             =   1680
         Width           =   1212
      End
      Begin VB.TextBox txtTurnerSpeed 
         Height          =   288
         Left            =   4800
         TabIndex        =   22
         Top             =   2040
         Width           =   1212
      End
      Begin VB.TextBox txtSCurveFactor 
         Height          =   288
         Left            =   4800
         TabIndex        =   23
         Top             =   2400
         Width           =   1212
      End
      Begin VB.TextBox txtSampHoleAlignOffset 
         Height          =   288
         Left            =   4800
         TabIndex        =   24
         Top             =   2880
         Width           =   1212
      End
      Begin VB.TextBox txtTurningMotor1rps 
         Height          =   288
         Left            =   4800
         TabIndex        =   25
         Top             =   3360
         Width           =   1212
      End
      Begin VB.TextBox txtUpDownMotor1cm 
         Height          =   288
         Left            =   4800
         TabIndex        =   26
         Top             =   3720
         Width           =   1212
      End
      Begin VB.Label lblSettings 
         Caption         =   "Lift Acceleration:"
         Height          =   375
         Index           =   84
         Left            =   3360
         TabIndex        =   324
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label lblSettings 
         Caption         =   "Min Up/Down pos:"
         Height          =   252
         Index           =   6
         Left            =   240
         TabIndex        =   162
         Top             =   2280
         Width           =   1572
      End
      Begin VB.Label lblSettings 
         Caption         =   "Floor position:"
         Height          =   252
         Index           =   5
         Left            =   240
         TabIndex        =   161
         Top             =   1920
         Width           =   1572
      End
      Begin VB.Label lblSettings 
         Caption         =   "Zeroing position:"
         Height          =   252
         Index           =   0
         Left            =   240
         TabIndex        =   126
         Top             =   120
         Width           =   1572
      End
      Begin VB.Label lblSettings 
         Caption         =   "Measuring position:"
         Height          =   252
         Index           =   1
         Left            =   240
         TabIndex        =   127
         Top             =   480
         Width           =   1452
      End
      Begin VB.Label lblSettings 
         Caption         =   "AF position:"
         Height          =   372
         Index           =   2
         Left            =   240
         TabIndex        =   128
         Top             =   840
         Width           =   1572
      End
      Begin VB.Label lblSettings 
         Caption         =   "IRM position:"
         Height          =   372
         Index           =   3
         Left            =   240
         TabIndex        =   129
         Top             =   1200
         Width           =   1572
      End
      Begin VB.Label lblSettings 
         Caption         =   "S Coil position:"
         Height          =   252
         Index           =   4
         Left            =   240
         TabIndex        =   130
         Top             =   1560
         Width           =   1572
      End
      Begin VB.Label lblSettings 
         Caption         =   "Default samp. top:"
         Height          =   252
         Index           =   7
         Left            =   240
         TabIndex        =   131
         Top             =   2640
         Width           =   1452
      End
      Begin VB.Label lblSettings 
         Caption         =   "Default samp. bot.:"
         Height          =   372
         Index           =   8
         Left            =   240
         TabIndex        =   132
         Top             =   3000
         Width           =   1572
      End
      Begin VB.Label lblSettings 
         Caption         =   "Turning full rotation:"
         Height          =   372
         Index           =   9
         Left            =   240
         TabIndex        =   133
         Top             =   3360
         Width           =   1572
      End
      Begin VB.Label lblSettings 
         Caption         =   "Tray offset (deg.):"
         Height          =   372
         Index           =   10
         Left            =   240
         TabIndex        =   134
         Top             =   3720
         Width           =   1572
      End
      Begin VB.Label lblSettings 
         Caption         =   "Lift speed slow:"
         Height          =   252
         Index           =   11
         Left            =   3360
         TabIndex        =   135
         Top             =   120
         Width           =   1572
      End
      Begin VB.Label lblSettings 
         Caption         =   "Lift speed medium:"
         Height          =   252
         Index           =   12
         Left            =   3360
         TabIndex        =   136
         Top             =   480
         Width           =   1452
      End
      Begin VB.Label lblSettings 
         Caption         =   "Lift speed fast:"
         Height          =   372
         Index           =   13
         Left            =   3360
         TabIndex        =   137
         Top             =   840
         Width           =   1572
      End
      Begin VB.Label lblSettings 
         Caption         =   "Changer speed:"
         Height          =   375
         Index           =   14
         Left            =   3360
         TabIndex        =   138
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label lblSettings 
         Caption         =   "Turning speed:"
         Height          =   255
         Index           =   15
         Left            =   3360
         TabIndex        =   139
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Label lblSettings 
         Caption         =   "S-curve factor:"
         Height          =   255
         Index           =   16
         Left            =   3360
         TabIndex        =   140
         Top             =   2400
         Width           =   1455
      End
      Begin VB.Label lblSettings 
         Caption         =   "Frac. samp. hole alignment offset:"
         Height          =   495
         Index           =   17
         Left            =   3360
         TabIndex        =   141
         Top             =   2760
         Width           =   1455
      End
      Begin VB.Label lblSettings 
         Caption         =   "Turning motor 1 rps:"
         Height          =   375
         Index           =   18
         Left            =   3360
         TabIndex        =   142
         Top             =   3360
         Width           =   1575
      End
      Begin VB.Label lblSettings 
         Caption         =   "Up/down 1 cm:"
         Height          =   375
         Index           =   19
         Left            =   3360
         TabIndex        =   143
         Top             =   3720
         Width           =   1575
      End
   End
   Begin VB.Frame frameOptions 
      BorderStyle     =   0  'None
      Caption         =   "Counts"
      Height          =   4695
      Index           =   2
      Left            =   360
      TabIndex        =   294
      Top             =   480
      Width           =   9735
      Begin VB.Frame frameCurrentXYMotorPosition 
         Caption         =   "Current Stage Position"
         Height          =   1815
         Left            =   3360
         TabIndex        =   339
         Top             =   1080
         Width           =   2175
         Begin VB.TextBox txtMotorPosition 
            Height          =   285
            Index           =   2
            Left            =   1200
            TabIndex        =   345
            Top             =   1320
            Width           =   855
         End
         Begin VB.TextBox txtMotorPosition 
            Height          =   285
            Index           =   1
            Left            =   840
            TabIndex        =   344
            Top             =   840
            Width           =   1215
         End
         Begin VB.TextBox txtMotorPosition 
            Height          =   285
            Index           =   0
            Left            =   840
            TabIndex        =   343
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label lblCurrentXYPosition 
            Caption         =   "Cup / Hole:"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   342
            Top             =   1320
            Width           =   975
         End
         Begin VB.Label lblCurrentXYPosition 
            Caption         =   "Y Motor:"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   341
            Top             =   840
            Width           =   735
         End
         Begin VB.Label lblCurrentXYPosition 
            Caption         =   "X Motor:"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   340
            Top             =   360
            Width           =   735
         End
      End
      Begin VB.TextBox txtEditGridCell 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3720
         TabIndex        =   338
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton ReadUDIO 
         Caption         =   "U/D"
         Height          =   255
         Left            =   2040
         TabIndex        =   318
         Top             =   3960
         Width           =   495
      End
      Begin VB.TextBox SlotMaxBox 
         Height          =   285
         Left            =   2400
         TabIndex        =   316
         Text            =   "Text1"
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox HomeHoleLocationBox 
         Height          =   285
         Left            =   2400
         TabIndex        =   314
         Text            =   "Text1"
         Top             =   600
         Width           =   615
      End
      Begin VB.CheckBox ckUseXYTableAPS 
         Caption         =   "Use XY Table"
         Height          =   375
         Left            =   0
         TabIndex        =   313
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton testAllCups 
         Caption         =   "Test All Cups"
         Height          =   375
         Left            =   1800
         TabIndex        =   312
         Top             =   2880
         Width           =   1335
      End
      Begin VB.CommandButton MoveHomeButton 
         Caption         =   "Move Home"
         Height          =   375
         Left            =   1800
         TabIndex        =   311
         Top             =   2400
         Width           =   1335
      End
      Begin VB.TextBox IOResult 
         Height          =   375
         Left            =   1920
         TabIndex        =   310
         Top             =   4320
         Width           =   735
      End
      Begin VB.TextBox ReadIOLine 
         Height          =   375
         Left            =   600
         TabIndex        =   309
         Text            =   "1"
         Top             =   4320
         Width           =   375
      End
      Begin VB.CommandButton ReadYIO 
         Caption         =   "Y"
         Height          =   255
         Left            =   1320
         TabIndex        =   308
         Top             =   3960
         Width           =   495
      End
      Begin VB.CommandButton ReadXIO 
         Caption         =   "X"
         Height          =   255
         Left            =   600
         TabIndex        =   307
         Top             =   3960
         Width           =   495
      End
      Begin VB.CommandButton MoveToCupButton 
         Caption         =   "Move To Cup"
         Height          =   375
         Left            =   0
         TabIndex        =   306
         Top             =   2880
         Width           =   1335
      End
      Begin VB.CommandButton SetCupButton 
         Caption         =   "Set Cup"
         Height          =   375
         Left            =   0
         TabIndex        =   305
         Top             =   2400
         Width           =   1335
      End
      Begin VB.CommandButton YNegButton 
         Caption         =   "-Y"
         Height          =   255
         Left            =   1080
         TabIndex        =   304
         Top             =   1560
         Width           =   495
      End
      Begin VB.CommandButton XNegButton 
         Caption         =   "-X"
         Height          =   255
         Left            =   240
         TabIndex        =   303
         Top             =   1200
         Width           =   495
      End
      Begin VB.CommandButton XPosButton 
         Caption         =   "+X"
         Height          =   255
         Left            =   1920
         TabIndex        =   302
         Top             =   1200
         Width           =   495
      End
      Begin VB.TextBox MoveSpeed 
         Height          =   285
         Left            =   2160
         TabIndex        =   299
         Text            =   "100000000"
         Top             =   2010
         Width           =   975
      End
      Begin VB.TextBox MoveCounts 
         Height          =   285
         Left            =   600
         TabIndex        =   298
         Text            =   "5000"
         Top             =   2010
         Width           =   735
      End
      Begin VB.CommandButton YPosButton 
         Caption         =   "+Y"
         Height          =   255
         Left            =   1080
         TabIndex        =   297
         Top             =   840
         Width           =   495
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid XYHolePositionsFlexGrid 
         Height          =   4455
         Left            =   5760
         TabIndex        =   295
         Top             =   120
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   7858
         _Version        =   393216
         Rows            =   102
         ScrollBars      =   2
         GridLineWidthFixed=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label Label10 
         Caption         =   "Result:"
         Height          =   255
         Left            =   1200
         TabIndex        =   327
         Top             =   4440
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "I/O:"
         Height          =   255
         Left            =   120
         TabIndex        =   326
         Top             =   4440
         Width           =   375
      End
      Begin VB.Label Label8 
         Caption         =   "Test limit Switches"
         Height          =   255
         Left            =   0
         TabIndex        =   325
         Top             =   3720
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "# Cups:"
         Height          =   255
         Left            =   1680
         TabIndex        =   317
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "Home Hole Location:"
         Height          =   255
         Left            =   720
         TabIndex        =   315
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Counts"
         Height          =   255
         Left            =   0
         TabIndex        =   301
         Top             =   2040
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Speed"
         Height          =   255
         Left            =   1560
         TabIndex        =   300
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Move Motors"
         Height          =   255
         Left            =   840
         TabIndex        =   296
         Top             =   1200
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdOKApplyCancel 
      Caption         =   "Save to Current Session, only"
      Height          =   375
      Index           =   1
      Left            =   9720
      TabIndex        =   3
      Top             =   5520
      Width           =   3375
   End
   Begin VB.CommandButton cmdOKApplyCancel 
      Caption         =   "Cancel"
      Height          =   375
      Index           =   2
      Left            =   8640
      TabIndex        =   2
      Top             =   5520
      Width           =   975
   End
   Begin VB.CommandButton cmdFlowControl 
      Caption         =   "Resume run"
      Height          =   375
      Left            =   2040
      TabIndex        =   4
      Top             =   5520
      Width           =   1335
   End
   Begin ComctlLib.TabStrip tbsOptions 
      Height          =   5295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   12975
      _ExtentX        =   22886
      _ExtentY        =   9340
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   12
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
            Caption         =   "DC Motors (&XY)"
            Key             =   "XYMotorCommands"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Control XY Table"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Ma&gnetometer"
            Key             =   "Squid"
            Object.Tag             =   ""
            Object.ToolTipText     =   "SQUID magnetometer and susceptibility bridge settings"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&AF Demag(1)"
            Key             =   "AF"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Alternating field demagnetization settings"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&AF Demag(2)"
            Key             =   "ADWIN_AF"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab7 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "AF/IRM &Channels"
            Key             =   "TTLRelays"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab8 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&IRM"
            Key             =   "IRM"
            Object.Tag             =   ""
            Object.ToolTipText     =   "IRM settings"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab9 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&ARM"
            Key             =   "ARM"
            Object.Tag             =   ""
            Object.ToolTipText     =   "ARM settings"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab10 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Vacuum"
            Key             =   "Vacuum"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Vacuum settings"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab11 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Mo&dules"
            Key             =   "Modules"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Modules"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab12 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "AF &Temp. Sensors"
            Key             =   "Temperature"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdCalibNewRod 
      Caption         =   "Calibration of the rod"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   5520
      Width           =   1620
   End
   Begin VB.CommandButton cmdOKApplyCancel 
      Caption         =   "Save To .INI File"
      Height          =   375
      Index           =   0
      Left            =   6120
      TabIndex        =   1
      Top             =   5520
      Width           =   2415
   End
   Begin VB.Frame frameOptions 
      BorderStyle     =   0  'None
      Caption         =   "Advanced settings"
      Height          =   3372
      Index           =   1
      Left            =   360
      TabIndex        =   27
      Top             =   720
      Width           =   6492
      Begin VB.TextBox txtUpDownMaxTorque 
         Height          =   288
         Left            =   2400
         TabIndex        =   320
         Top             =   2880
         Width           =   1212
      End
      Begin VB.TextBox txtMotorIDChangerY 
         Height          =   288
         Left            =   1680
         TabIndex        =   293
         Top             =   1200
         Width           =   1212
      End
      Begin VB.TextBox txtMotorIDTurning 
         Height          =   288
         Left            =   1680
         TabIndex        =   28
         Top             =   480
         Width           =   1212
      End
      Begin VB.TextBox txtMotorIDChanger 
         Height          =   288
         Left            =   1680
         TabIndex        =   29
         Top             =   840
         Width           =   1212
      End
      Begin VB.TextBox txtMotorIDUpDown 
         Height          =   288
         Left            =   1680
         TabIndex        =   30
         Top             =   1560
         Width           =   1212
      End
      Begin VB.TextBox txtCmdHometoTop 
         Height          =   288
         Left            =   5040
         TabIndex        =   31
         Top             =   480
         Width           =   1212
      End
      Begin VB.TextBox txtCmdSamplePickup 
         Height          =   288
         Left            =   5040
         TabIndex        =   32
         Top             =   840
         Width           =   1212
      End
      Begin VB.TextBox txtUpDownTorquefactor 
         Height          =   288
         Left            =   2400
         TabIndex        =   33
         Top             =   2160
         Width           =   1212
      End
      Begin VB.TextBox txtPickupTorqueThrottle 
         Height          =   288
         Left            =   2400
         TabIndex        =   34
         Top             =   2520
         Width           =   1212
      End
      Begin VB.Label lblSettings 
         Caption         =   "Up/Down Max Torque:"
         Height          =   255
         Index           =   83
         Left            =   0
         TabIndex        =   319
         Top             =   2880
         Width           =   2295
      End
      Begin VB.Label lblSettings 
         Caption         =   "ChangerY motor ID:"
         Height          =   255
         Index           =   78
         Left            =   240
         TabIndex        =   292
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label lblSettings 
         Caption         =   "Turning motor ID:"
         Height          =   252
         Index           =   20
         Left            =   240
         TabIndex        =   144
         Top             =   480
         Width           =   1572
      End
      Begin VB.Label lblSettings 
         Caption         =   "Changer motor ID:"
         Height          =   252
         Index           =   21
         Left            =   240
         TabIndex        =   145
         Top             =   840
         Width           =   1452
      End
      Begin VB.Label lblSettings 
         Caption         =   "Up/Down motor ID:"
         Height          =   375
         Index           =   22
         Left            =   240
         TabIndex        =   146
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label lblSettings 
         Caption         =   "Home to top code:"
         Height          =   252
         Index           =   23
         Left            =   3360
         TabIndex        =   147
         Top             =   480
         Width           =   1572
      End
      Begin VB.Label lblSettings 
         Caption         =   "Sample pickup code:"
         Height          =   252
         Index           =   24
         Left            =   3360
         TabIndex        =   148
         Top             =   840
         Width           =   1572
      End
      Begin VB.Label lblSettings 
         Caption         =   "Up/Down Torque Factor:"
         Height          =   255
         Index           =   25
         Left            =   0
         TabIndex        =   149
         Top             =   2160
         Width           =   2295
      End
      Begin VB.Label lblSettings 
         Caption         =   "Sample pickup torque throttle:"
         Height          =   255
         Index           =   26
         Left            =   0
         TabIndex        =   150
         Top             =   2520
         Width           =   2295
      End
   End
   Begin VB.Frame frameOptions 
      BorderStyle     =   0  'None
      Caption         =   "SQUID Magnetometer Calibration factors"
      Height          =   3372
      Index           =   3
      Left            =   360
      TabIndex        =   35
      Top             =   720
      Width           =   6492
      Begin VB.TextBox txtXCal 
         Height          =   288
         Left            =   2040
         TabIndex        =   36
         Top             =   120
         Width           =   972
      End
      Begin VB.TextBox txtYCal 
         Height          =   288
         Left            =   2040
         TabIndex        =   37
         Top             =   600
         Width           =   972
      End
      Begin VB.TextBox txtZCal 
         Height          =   288
         Left            =   2040
         TabIndex        =   38
         Top             =   1080
         Width           =   972
      End
      Begin VB.TextBox txtRangeFact 
         Height          =   288
         Left            =   4920
         TabIndex        =   39
         Top             =   120
         Width           =   1092
      End
      Begin VB.TextBox txtReadDelay 
         Height          =   285
         Left            =   4920
         TabIndex        =   40
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox txtSusceptibilityMomentFactorCGS 
         Height          =   288
         Left            =   3840
         TabIndex        =   41
         Top             =   1680
         Width           =   972
      End
      Begin VB.TextBox txtSusceptibilitySettings 
         Height          =   288
         Left            =   3840
         TabIndex        =   42
         Top             =   2160
         Width           =   972
      End
      Begin VB.Label lblSettings 
         Caption         =   "X Calibration factor:"
         Height          =   255
         Index           =   27
         Left            =   240
         TabIndex        =   151
         Top             =   120
         Width           =   1575
      End
      Begin VB.Label lblSettings 
         Caption         =   "Y Calibration factor:"
         Height          =   255
         Index           =   28
         Left            =   240
         TabIndex        =   152
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label lblSettings 
         Caption         =   "Z Calibration factor:"
         Height          =   375
         Index           =   29
         Left            =   240
         TabIndex        =   153
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label lblSettings 
         Caption         =   "Range factor:"
         Height          =   255
         Index           =   30
         Left            =   3600
         TabIndex        =   154
         Top             =   120
         Width           =   1575
      End
      Begin VB.Label lblSettings 
         Caption         =   "Read delay"
         Height          =   255
         Index           =   31
         Left            =   3600
         TabIndex        =   155
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label lblSettings 
         Caption         =   "Moment Susceptibility Calibration factor (CGS):"
         Height          =   255
         Index           =   32
         Left            =   240
         TabIndex        =   156
         Top             =   1680
         Width           =   3975
      End
      Begin VB.Label lblSettings 
         Caption         =   "Susceptibility COM settings:"
         Height          =   255
         Index           =   33
         Left            =   240
         TabIndex        =   157
         Top             =   2160
         Width           =   3975
      End
   End
   Begin VB.Frame frameOptions 
      BorderStyle     =   0  'None
      Caption         =   "Vacuum"
      Height          =   4695
      Index           =   4
      Left            =   240
      TabIndex        =   43
      Top             =   480
      Width           =   6495
      Begin VB.CommandButton cmdOpenAFForm 
         Caption         =   "Open AF Control Window"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         TabIndex        =   285
         Top             =   3840
         Width           =   3375
      End
      Begin VB.ComboBox cmbAFSystem 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1560
         TabIndex        =   44
         Top             =   240
         Width           =   1692
      End
      Begin VB.ComboBox cmbAFAxialCoord 
         Height          =   315
         Left            =   1800
         TabIndex        =   50
         Top             =   1440
         Width           =   612
      End
      Begin VB.TextBox txtAFAxialMax 
         Height          =   288
         Left            =   1800
         TabIndex        =   46
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox txtAFAxialMin 
         Height          =   288
         Left            =   1800
         TabIndex        =   47
         Top             =   1080
         Width           =   1215
      End
      Begin VB.ComboBox cmbAFDelay 
         Height          =   315
         Left            =   1800
         TabIndex        =   210
         Top             =   1800
         Width           =   732
      End
      Begin VB.ComboBox cmbAFTransCoord 
         Height          =   315
         Left            =   4920
         TabIndex        =   51
         Top             =   1440
         Width           =   612
      End
      Begin VB.TextBox txtAFTransMax 
         Height          =   288
         Left            =   4920
         TabIndex        =   48
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox txtAFTransMin 
         Height          =   288
         Left            =   4920
         TabIndex        =   49
         Top             =   1080
         Width           =   1215
      End
      Begin VB.ComboBox cmbAFRampRate 
         Height          =   315
         Left            =   4920
         TabIndex        =   209
         Top             =   1800
         Width           =   732
      End
      Begin VB.CommandButton cmdCalAFCoils 
         Caption         =   "Calibrate AF Coils"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         TabIndex        =   52
         Top             =   2400
         Width           =   3375
      End
      Begin VB.ComboBox cmbAFUnits 
         Height          =   315
         Left            =   4920
         TabIndex        =   45
         Top             =   240
         Width           =   732
      End
      Begin VB.CommandButton cmdAFFileSettings 
         Caption         =   "AF File Settings"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         TabIndex        =   54
         Top             =   3360
         Width           =   3375
      End
      Begin VB.CommandButton cmdTuneAF 
         Caption         =   "Tune AF Coils"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         TabIndex        =   53
         Top             =   2880
         Width           =   3375
      End
      Begin VB.Label lblAFSystem 
         Caption         =   "AF System:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   220
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblAxialCoord 
         Caption         =   "Axial Coordinate:"
         Height          =   255
         Left            =   360
         TabIndex        =   219
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label lblAxialMax 
         Caption         =   "Axial Max:"
         Height          =   255
         Left            =   360
         TabIndex        =   218
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label lblAFDelay 
         Caption         =   "AF Delay:"
         Height          =   255
         Left            =   360
         TabIndex        =   217
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label lblTransCoord 
         Caption         =   "Trans. coordinate:"
         Height          =   255
         Left            =   3480
         TabIndex        =   216
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label lblTransMax 
         Caption         =   "Trans. Max:"
         Height          =   255
         Left            =   3480
         TabIndex        =   215
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label lblAFRampRate 
         Caption         =   "AF Ramp Rate:"
         Height          =   255
         Left            =   3480
         TabIndex        =   214
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label lblAFUnits 
         Caption         =   "AF/IRM Field Units:"
         Height          =   255
         Left            =   3480
         TabIndex        =   213
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lblAxialMin 
         Caption         =   "Axial Min:"
         Height          =   255
         Left            =   360
         TabIndex        =   212
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label lblTransMin 
         Caption         =   "Trans. Min:"
         Height          =   255
         Left            =   3480
         TabIndex        =   211
         Top             =   1080
         Width           =   1335
      End
   End
   Begin VB.Frame frameOptions 
      BorderStyle     =   0  'None
      Caption         =   "Vacuum"
      Height          =   4695
      Index           =   5
      Left            =   240
      TabIndex        =   246
      Top             =   600
      Width           =   6495
      Begin VB.TextBox txtPeakPeriods 
         Height          =   285
         Left            =   2160
         TabIndex        =   336
         Top             =   4080
         Width           =   975
      End
      Begin VB.TextBox txtAFTransverseRampUpVoltsPerSec 
         Height          =   285
         Left            =   1920
         TabIndex        =   267
         Top             =   3480
         Width           =   1215
      End
      Begin VB.TextBox txtAFAxialRampUpVoltsPerSec 
         Height          =   285
         Left            =   1920
         TabIndex        =   266
         Top             =   2880
         Width           =   1215
      End
      Begin VB.TextBox txtAFMinRampDownNumPeriods 
         Height          =   285
         Left            =   5040
         TabIndex        =   251
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox txtAFMaxRampDownNumPeriods 
         Height          =   285
         Left            =   5040
         TabIndex        =   249
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox txtAFMaxRampUpTime 
         Height          =   285
         Left            =   1920
         TabIndex        =   250
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox txtAFMinRampUpTime 
         Height          =   285
         Left            =   1920
         TabIndex        =   252
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox txtSampleRampOutputVoltage 
         Height          =   285
         Left            =   5040
         TabIndex        =   247
         Top             =   2400
         Width           =   1215
      End
      Begin VB.TextBox txtAFRampDownPeriodsPerVolt 
         Height          =   285
         Left            =   5040
         TabIndex        =   248
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label11 
         Caption         =   "Duration in Sine Wave Periods at the Peak Field:"
         Height          =   735
         Left            =   240
         TabIndex        =   337
         Top             =   3960
         Width           =   2055
      End
      Begin VB.Label lblAFTransverseRampUpVoltsPerSec 
         Caption         =   "Transverse Ramp Up Volts Per Second:"
         Height          =   495
         Left            =   240
         TabIndex        =   269
         Top             =   3360
         Width           =   1575
      End
      Begin VB.Label lblAFAxialRampUpVoltsPerSec 
         Caption         =   "Axial Ramp Up Volts Per Second:"
         Height          =   495
         Left            =   240
         TabIndex        =   268
         Top             =   2760
         Width           =   1575
      End
      Begin VB.Label lblAFMaxRampDownNumPeriods 
         Caption         =   "AF Max # Ramp Down Periods:"
         Height          =   375
         Left            =   3600
         TabIndex        =   265
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label lblAFMinRampUpTime 
         Caption         =   "AF Min Ramp Up Time (ms):"
         Height          =   495
         Left            =   240
         TabIndex        =   264
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label lblAFMaxRampUpTime 
         Caption         =   "AF Max Ramp Up Time (ms):"
         Height          =   495
         Left            =   240
         TabIndex        =   263
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label lblAFMinRampDownNumPeriods 
         Caption         =   "AF Min # Ramp Down Periods:"
         Height          =   495
         Left            =   3600
         TabIndex        =   262
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label lblAFRampDownPeriodsPerVolt 
         Caption         =   "AF # of Periods / Ramp Output Volts:"
         Height          =   495
         Left            =   3600
         TabIndex        =   261
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label lblRampDownTimeHeader 
         Caption         =   "Ramp Down Time (ms):"
         Height          =   255
         Left            =   3600
         TabIndex        =   260
         Top             =   2880
         Width           =   1815
      End
      Begin VB.Label lblRampOutputMaxVoltHeader 
         Caption         =   "AF Ramp Ouput Max Voltages:"
         Height          =   255
         Left            =   240
         TabIndex        =   259
         Top             =   1560
         Width           =   2655
      End
      Begin VB.Label lblAFAxialRampMaxLabel 
         Caption         =   "Axial:"
         Height          =   255
         Left            =   240
         TabIndex        =   258
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label lblAFTransverseRampMaxLabel 
         Caption         =   "Transverse:"
         Height          =   255
         Left            =   240
         TabIndex        =   257
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label lblAFAxialRampMax 
         Caption         =   "Label3"
         Height          =   255
         Left            =   1320
         TabIndex        =   256
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label lblAFTransverseRampMax 
         Caption         =   "Label3"
         Height          =   255
         Left            =   1320
         TabIndex        =   255
         Top             =   2280
         Width           =   975
      End
      Begin VB.Label lblResultingRampDownTime 
         Caption         =   "Label9"
         Height          =   975
         Left            =   3720
         TabIndex        =   254
         Top             =   3120
         Width           =   2655
      End
      Begin VB.Label lblSampleRampOutputVoltHeader 
         Caption         =   "Sample Ramp Output Voltage:"
         Height          =   495
         Left            =   3600
         TabIndex        =   253
         Top             =   2280
         Width           =   1335
      End
   End
   Begin VB.Frame frameOptions 
      BorderStyle     =   0  'None
      Caption         =   "Frame5"
      Height          =   4572
      Index           =   6
      Left            =   360
      TabIndex        =   55
      Top             =   600
      Width           =   6492
      Begin VB.Frame frameAFRamp 
         Caption         =   "AF Ramp Output"
         Height          =   1455
         Left            =   3480
         TabIndex        =   233
         Top             =   3120
         Width           =   2535
         Begin VB.ComboBox cmbAFRampBoard 
            Height          =   315
            Left            =   240
            TabIndex        =   66
            Top             =   480
            Width           =   2055
         End
         Begin VB.ComboBox cmbAFRampChan 
            Height          =   315
            Left            =   960
            TabIndex        =   67
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label lblSettings 
            Caption         =   "DAQ Board:"
            Height          =   255
            Index           =   47
            Left            =   240
            TabIndex        =   235
            Top             =   240
            Width           =   1815
         End
         Begin VB.Label lblSettings 
            Caption         =   "Channel:"
            Height          =   255
            Index           =   48
            Left            =   240
            TabIndex        =   234
            Top             =   960
            Width           =   1455
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "AF Axial TTL Relay"
         Height          =   1452
         Left            =   480
         TabIndex        =   197
         Top             =   0
         Width           =   2532
         Begin VB.ComboBox cmbAxialRelayChan 
            Height          =   288
            Left            =   960
            TabIndex        =   57
            Top             =   960
            Width           =   1332
         End
         Begin VB.ComboBox cmbAxialRelayBoard 
            Height          =   288
            Left            =   240
            TabIndex        =   56
            Top             =   480
            Width           =   2052
         End
         Begin VB.Label lblSettings 
            Caption         =   "Channel:"
            Height          =   252
            Index           =   38
            Left            =   240
            TabIndex        =   199
            Top             =   960
            Width           =   732
         End
         Begin VB.Label lblSettings 
            Caption         =   "DAQ Board:"
            Height          =   252
            Index           =   37
            Left            =   240
            TabIndex        =   198
            Top             =   240
            Width           =   1092
         End
      End
      Begin VB.Frame frameAFMonitor 
         Caption         =   "AF Monitor Input"
         Height          =   1455
         Left            =   3480
         TabIndex        =   230
         Top             =   1560
         Width           =   2535
         Begin VB.ComboBox cmbAFMonitorChan 
            Height          =   315
            Left            =   960
            TabIndex        =   65
            Top             =   960
            Width           =   1335
         End
         Begin VB.ComboBox cmbAFMonitorBoard 
            Height          =   315
            Left            =   240
            TabIndex        =   64
            Top             =   480
            Width           =   2055
         End
         Begin VB.Label lblSettings 
            Caption         =   "Channel:"
            Height          =   255
            Index           =   46
            Left            =   240
            TabIndex        =   232
            Top             =   960
            Width           =   1455
         End
         Begin VB.Label lblSettings 
            Caption         =   "DAQ Board:"
            Height          =   255
            Index           =   45
            Left            =   240
            TabIndex        =   231
            Top             =   240
            Width           =   1695
         End
      End
      Begin VB.Frame frameAltAFMonitor 
         Caption         =   "Alternate AF Monitor Input"
         Height          =   1455
         Left            =   3480
         TabIndex        =   227
         Top             =   0
         Width           =   2535
         Begin VB.ComboBox cmbAltAFMonitorBoard 
            Height          =   315
            Left            =   240
            TabIndex        =   62
            Top             =   480
            Width           =   2055
         End
         Begin VB.ComboBox cmbAltAFMonitorChan 
            Height          =   315
            Left            =   960
            TabIndex        =   63
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label lblSettings 
            Caption         =   "DAQ Board:"
            Height          =   375
            Index           =   43
            Left            =   240
            TabIndex        =   229
            Top             =   240
            Width           =   1815
         End
         Begin VB.Label lblSettings 
            Caption         =   "Channel:"
            Height          =   255
            Index           =   44
            Left            =   240
            TabIndex        =   228
            Top             =   960
            Width           =   1455
         End
      End
      Begin VB.Frame frameIRMRelay 
         Caption         =   "IRM TTL Relay"
         Height          =   1452
         Left            =   480
         TabIndex        =   203
         Top             =   3120
         Width           =   2532
         Begin VB.ComboBox cmbIRMRelayChan 
            Height          =   288
            Left            =   960
            TabIndex        =   61
            Top             =   960
            Width           =   1332
         End
         Begin VB.ComboBox cmbIRMRelayBoard 
            Height          =   288
            Left            =   240
            TabIndex        =   60
            Top             =   480
            Width           =   2052
         End
         Begin VB.Label lblSettings 
            Caption         =   "Channel:"
            Height          =   252
            Index           =   42
            Left            =   240
            TabIndex        =   205
            Top             =   960
            Width           =   732
         End
         Begin VB.Label lblSettings 
            Caption         =   "DAQ Board:"
            Height          =   252
            Index           =   41
            Left            =   240
            TabIndex        =   204
            Top             =   240
            Width           =   1092
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "AF Transverse TTL Relay"
         Height          =   1452
         Left            =   480
         TabIndex        =   200
         Top             =   1560
         Width           =   2532
         Begin VB.ComboBox cmbTransRelayBoard 
            Height          =   288
            Left            =   240
            TabIndex        =   58
            Top             =   480
            Width           =   2052
         End
         Begin VB.ComboBox cmbTransRelayChan 
            Height          =   288
            Left            =   960
            TabIndex        =   59
            Top             =   960
            Width           =   1332
         End
         Begin VB.Label lblSettings 
            Caption         =   "DAQ Board:"
            Height          =   252
            Index           =   39
            Left            =   240
            TabIndex        =   202
            Top             =   240
            Width           =   1092
         End
         Begin VB.Label lblSettings 
            Caption         =   "Channel:"
            Height          =   252
            Index           =   40
            Left            =   240
            TabIndex        =   201
            Top             =   960
            Width           =   732
         End
      End
   End
   Begin VB.Frame frameOptions 
      BorderStyle     =   0  'None
      Caption         =   "ARM"
      Height          =   3372
      Index           =   8
      Left            =   360
      TabIndex        =   77
      Top             =   600
      Width           =   6492
      Begin VB.TextBox txtARMMax 
         Height          =   288
         Left            =   1800
         TabIndex        =   78
         Top             =   240
         Width           =   1212
      End
      Begin VB.TextBox txtARMVoltGauss 
         Height          =   288
         Left            =   1800
         TabIndex        =   79
         Top             =   720
         Width           =   1212
      End
      Begin VB.TextBox txtARMVoltMax 
         Height          =   288
         Left            =   1800
         TabIndex        =   80
         Top             =   1200
         Width           =   1212
      End
      Begin VB.Label lblSettings 
         Caption         =   "ARM Max:"
         Height          =   255
         Index           =   34
         Left            =   240
         TabIndex        =   158
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label lblSettings 
         Caption         =   "ARM Volt/Gauss:"
         Height          =   252
         Index           =   35
         Left            =   240
         TabIndex        =   159
         Top             =   720
         Width           =   1452
      End
      Begin VB.Label lblSettings 
         Caption         =   "ARM Volt Max:"
         Height          =   372
         Index           =   36
         Left            =   240
         TabIndex        =   160
         Top             =   1200
         Width           =   1572
      End
   End
   Begin VB.Frame frameOptions 
      BorderStyle     =   0  'None
      Caption         =   "Vacuum"
      Height          =   4575
      Index           =   9
      Left            =   240
      TabIndex        =   81
      Top             =   600
      Width           =   6495
      Begin VB.CheckBox chkDegausserAirCooler 
         Caption         =   "Run Air Cooler While Degaussing"
         Height          =   735
         Left            =   0
         TabIndex        =   334
         Top             =   2640
         Width           =   2175
      End
      Begin VB.Frame Frame10 
         Caption         =   "Air Cooler"
         Height          =   1812
         Left            =   2760
         TabIndex        =   328
         Top             =   2520
         Width           =   2532
         Begin VB.ComboBox cmbDegausserCoolerChan 
            Height          =   315
            Left            =   240
            TabIndex        =   330
            Top             =   1320
            Width           =   2052
         End
         Begin VB.ComboBox cmbDegausserCoolerBoard 
            Height          =   315
            Left            =   240
            TabIndex        =   329
            Top             =   600
            Width           =   2052
         End
         Begin VB.Label lblSettings 
            Caption         =   "Channel:"
            Height          =   252
            Index           =   85
            Left            =   240
            TabIndex        =   332
            Top             =   1080
            Width           =   732
         End
         Begin VB.Label lblSettings 
            Caption         =   "DAQ Board:"
            Height          =   252
            Index           =   86
            Left            =   240
            TabIndex        =   331
            Top             =   360
            Width           =   1092
         End
      End
      Begin VB.TextBox txtDropoffVacuumDelay 
         Height          =   375
         Left            =   4920
         TabIndex        =   322
         Text            =   "Text1"
         Top             =   120
         Width           =   1215
      End
      Begin VB.Frame Frame2 
         Caption         =   "Vacuum Toggle A"
         Height          =   1812
         Left            =   2760
         TabIndex        =   166
         Top             =   600
         Width           =   2532
         Begin VB.ComboBox cmbVacuumToggleABoard 
            Height          =   315
            Left            =   240
            TabIndex        =   85
            Top             =   600
            Width           =   2052
         End
         Begin VB.ComboBox cmbVacuumToggleAChan 
            Height          =   315
            Left            =   240
            TabIndex        =   86
            Top             =   1320
            Width           =   2052
         End
         Begin VB.Label lblSettings 
            Caption         =   "DAQ Board:"
            Height          =   252
            Index           =   52
            Left            =   240
            TabIndex        =   168
            Top             =   360
            Width           =   1092
         End
         Begin VB.Label lblSettings 
            Caption         =   "Channel:"
            Height          =   252
            Index           =   53
            Left            =   240
            TabIndex        =   167
            Top             =   1080
            Width           =   732
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Motor Toggle"
         Height          =   1812
         Left            =   120
         TabIndex        =   163
         Top             =   600
         Width           =   2532
         Begin VB.ComboBox cmbMotorToggleChan 
            Height          =   315
            Left            =   240
            TabIndex        =   84
            Top             =   1320
            Width           =   2052
         End
         Begin VB.ComboBox cmbMotorToggleBoard 
            Height          =   315
            Left            =   240
            TabIndex        =   83
            Top             =   600
            Width           =   2052
         End
         Begin VB.Label lblSettings 
            Caption         =   "Channel:"
            Height          =   252
            Index           =   51
            Left            =   240
            TabIndex        =   165
            Top             =   1080
            Width           =   732
         End
         Begin VB.Label lblSettings 
            Caption         =   "DAQ Board:"
            Height          =   252
            Index           =   50
            Left            =   240
            TabIndex        =   164
            Top             =   360
            Width           =   1092
         End
      End
      Begin VB.CheckBox chkDoVacuumReset 
         Caption         =   "Reset vacuum on startup"
         Height          =   192
         Left            =   240
         TabIndex        =   82
         Top             =   240
         Width           =   2292
      End
      Begin VB.Label lblSettings 
         Caption         =   "DIO line assignments:"
         Height          =   255
         Index           =   49
         Left            =   3480
         TabIndex        =   333
         Top             =   1920
         Width           =   1800
      End
      Begin VB.Label Label7 
         Caption         =   "Dropoff Vacuum Delay: "
         Height          =   255
         Left            =   3120
         TabIndex        =   321
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame frameOptions 
      BorderStyle     =   0  'None
      Caption         =   "2080"
      Height          =   4695
      Index           =   10
      Left            =   240
      TabIndex        =   87
      Top             =   480
      Width           =   6495
      Begin VB.CheckBox chkEnableDegausserCooler 
         Caption         =   "Enable Degausser Cooler"
         Height          =   375
         Left            =   120
         TabIndex        =   335
         Top             =   4080
         Width           =   1935
      End
      Begin VB.CheckBox chkEnableVacuum 
         Caption         =   "Enable Vacuum"
         Height          =   375
         Left            =   120
         TabIndex        =   290
         Top             =   3720
         Width           =   1935
      End
      Begin VB.CheckBox chkEnableAltAFMonitor 
         Caption         =   "Alternate AF Monitor"
         Height          =   375
         Left            =   120
         TabIndex        =   95
         Top             =   2760
         Width           =   2055
      End
      Begin VB.CheckBox chkEnableIRMTrans 
         Caption         =   "IRM Transverse Pulse"
         Height          =   375
         Left            =   120
         TabIndex        =   89
         Top             =   360
         Width           =   1935
      End
      Begin VB.CheckBox chkEnableAFAnalysis 
         Caption         =   "AF Analysis Mode"
         Height          =   375
         Left            =   120
         TabIndex        =   94
         Top             =   2400
         Width           =   2055
      End
      Begin VB.CheckBox chkEnableIRMMonitor 
         Caption         =   "IRM Monitor"
         Height          =   375
         Left            =   120
         TabIndex        =   91
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Frame frameIRMMonitor 
         Caption         =   "IRM Monitor"
         Height          =   1212
         Left            =   2400
         TabIndex        =   206
         Top             =   3360
         Width           =   3972
         Begin VB.ComboBox cmbIRMMonitorBoard 
            Height          =   315
            Left            =   1560
            TabIndex        =   106
            Top             =   360
            Width           =   2052
         End
         Begin VB.ComboBox cmbIRMMonitorChan 
            Height          =   315
            Left            =   1560
            TabIndex        =   107
            Top             =   720
            Width           =   2052
         End
         Begin VB.Label lblSettings 
            Caption         =   "DAQ Board:"
            Height          =   252
            Index           =   57
            Left            =   240
            TabIndex        =   208
            Top             =   360
            Width           =   1092
         End
         Begin VB.Label lblSettings 
            Caption         =   "Channel:"
            Height          =   252
            Index           =   58
            Left            =   480
            TabIndex        =   207
            Top             =   720
            Width           =   732
         End
      End
      Begin VB.Frame frameARMVoltageOut 
         Caption         =   "ARM Voltage Out"
         Height          =   1212
         Left            =   2400
         TabIndex        =   188
         Top             =   720
         Width           =   3972
         Begin VB.ComboBox cmbARMVoltageOutBoard 
            Height          =   315
            Left            =   1560
            TabIndex        =   98
            Top             =   360
            Width           =   2052
         End
         Begin VB.ComboBox cmbARMVoltageOutChan 
            Height          =   315
            Left            =   1560
            TabIndex        =   99
            Top             =   720
            Width           =   2052
         End
         Begin VB.Label lblSettings 
            Caption         =   "DAQ Board:"
            Height          =   252
            Index           =   55
            Left            =   240
            TabIndex        =   190
            Top             =   360
            Width           =   1092
         End
         Begin VB.Label lblSettings 
            Caption         =   "Channel:"
            Height          =   252
            Index           =   56
            Left            =   480
            TabIndex        =   189
            Top             =   720
            Width           =   732
         End
      End
      Begin VB.Frame frameIRMPowerAmpVoltageIn 
         Caption         =   "IRM Power Amp Voltage In"
         Height          =   1212
         Left            =   2400
         TabIndex        =   185
         Top             =   3360
         Width           =   3972
         Begin VB.ComboBox cmbIRMPowerAmpVoltageInBoard 
            Height          =   315
            Left            =   1560
            TabIndex        =   100
            Top             =   360
            Width           =   2052
         End
         Begin VB.ComboBox cmbIRMPowerAmpVoltageInChan 
            Height          =   315
            Left            =   1560
            TabIndex        =   101
            Top             =   720
            Width           =   2052
         End
         Begin VB.Label lblSettings 
            Caption         =   "DAQ Board:"
            Height          =   252
            Index           =   65
            Left            =   240
            TabIndex        =   187
            Top             =   360
            Width           =   1092
         End
         Begin VB.Label lblSettings 
            Caption         =   "Channel:"
            Height          =   252
            Index           =   66
            Left            =   480
            TabIndex        =   186
            Top             =   720
            Width           =   732
         End
      End
      Begin VB.Frame frameIRMTrim 
         Caption         =   "IRM Trim"
         Height          =   1212
         Left            =   2400
         TabIndex        =   182
         Top             =   2040
         Width           =   3972
         Begin VB.ComboBox cmbIRMTrimChan 
            Height          =   315
            Left            =   1560
            TabIndex        =   113
            Top             =   720
            Width           =   2052
         End
         Begin VB.ComboBox cmbIRMTrimBoard 
            Height          =   315
            Left            =   1560
            TabIndex        =   112
            Top             =   360
            Width           =   2052
         End
         Begin VB.Label lblSettings 
            Caption         =   "Channel:"
            Height          =   252
            Index           =   60
            Left            =   480
            TabIndex        =   184
            Top             =   720
            Width           =   732
         End
         Begin VB.Label lblSettings 
            Caption         =   "DAQ Board:"
            Height          =   252
            Index           =   59
            Left            =   240
            TabIndex        =   183
            Top             =   360
            Width           =   1092
         End
      End
      Begin VB.Frame frameIRMFire 
         Caption         =   "IRM Fire"
         Height          =   1212
         Left            =   2400
         TabIndex        =   179
         Top             =   720
         Width           =   3972
         Begin VB.ComboBox cmbIRMFireChan 
            Height          =   315
            Left            =   1560
            TabIndex        =   109
            Top             =   720
            Width           =   2052
         End
         Begin VB.ComboBox cmbIRMFireBoard 
            Height          =   315
            Left            =   1560
            TabIndex        =   108
            Top             =   360
            Width           =   2052
         End
         Begin VB.Label lblSettings 
            Caption         =   "Channel:"
            Height          =   252
            Index           =   62
            Left            =   480
            TabIndex        =   181
            Top             =   720
            Width           =   732
         End
         Begin VB.Label lblSettings 
            Caption         =   "DAQ Board:"
            Height          =   252
            Index           =   61
            Left            =   240
            TabIndex        =   180
            Top             =   360
            Width           =   1092
         End
      End
      Begin VB.Frame frameIRMVoltageOut 
         Caption         =   "IRM Voltage Out"
         Height          =   1212
         Left            =   2400
         TabIndex        =   176
         Top             =   720
         Width           =   3972
         Begin VB.ComboBox cmbIRMVoltageOutChan 
            Height          =   315
            Left            =   1560
            TabIndex        =   111
            Top             =   720
            Width           =   2052
         End
         Begin VB.ComboBox cmbIRMVoltageOutBoard 
            Height          =   315
            Left            =   1560
            TabIndex        =   110
            Top             =   360
            Width           =   2052
         End
         Begin VB.Label lblSettings 
            Caption         =   "Channel:"
            Height          =   252
            Index           =   64
            Left            =   480
            TabIndex        =   178
            Top             =   720
            Width           =   732
         End
         Begin VB.Label lblSettings 
            Caption         =   "DAQ Board:"
            Height          =   252
            Index           =   63
            Left            =   240
            TabIndex        =   177
            Top             =   360
            Width           =   1092
         End
      End
      Begin VB.Frame frameARMSet 
         Caption         =   "ARM Set"
         Height          =   1212
         Left            =   2400
         TabIndex        =   173
         Top             =   2040
         Width           =   3972
         Begin VB.ComboBox cmbARMSetBoard 
            Height          =   315
            Left            =   1560
            TabIndex        =   102
            Top             =   360
            Width           =   2052
         End
         Begin VB.ComboBox cmbARMSetChan 
            Height          =   315
            Left            =   1560
            TabIndex        =   103
            Top             =   720
            Width           =   2052
         End
         Begin VB.Label lblSettings 
            Caption         =   "DAQ Board:"
            Height          =   252
            Index           =   67
            Left            =   240
            TabIndex        =   175
            Top             =   360
            Width           =   1092
         End
         Begin VB.Label lblSettings 
            Caption         =   "Channel:"
            Height          =   255
            Index           =   68
            Left            =   480
            TabIndex        =   174
            Top             =   720
            Width           =   735
         End
      End
      Begin VB.Frame frameIRMCapacitorVoltageIn 
         Caption         =   "IRM Capacitor Voltage In"
         Height          =   1212
         Left            =   2400
         TabIndex        =   170
         Top             =   2040
         Width           =   3972
         Begin VB.ComboBox cmbIRMCapacitorVoltageInChan 
            Height          =   315
            Left            =   1560
            TabIndex        =   105
            Top             =   720
            Width           =   2052
         End
         Begin VB.ComboBox cmbIRMCapacitorVoltageInBoard 
            Height          =   315
            Left            =   1560
            TabIndex        =   104
            Top             =   360
            Width           =   2052
         End
         Begin VB.Label lblSettings 
            Caption         =   "Channel:"
            Height          =   252
            Index           =   70
            Left            =   480
            TabIndex        =   172
            Top             =   720
            Width           =   732
         End
         Begin VB.Label lblSettings 
            Caption         =   "DAQ Board:"
            Height          =   252
            Index           =   69
            Left            =   240
            TabIndex        =   171
            Top             =   360
            Width           =   1092
         End
      End
      Begin ComctlLib.TabStrip tbsARMIRMChannels 
         Height          =   4575
         Left            =   2280
         TabIndex        =   97
         Top             =   120
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   8070
         _Version        =   327682
         BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
            NumTabs         =   3
            BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "IRM (1)"
               Key             =   ""
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "IRM (2)"
               Key             =   ""
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "ARM"
               Key             =   ""
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.CheckBox chkEnableIRMAxial 
         Caption         =   "IRM Axial Pulse"
         Height          =   375
         Left            =   0
         TabIndex        =   88
         Top             =   0
         Width           =   1575
      End
      Begin VB.CheckBox chkEnableARM 
         Caption         =   "ARM Biasing Field Coil"
         Height          =   375
         Left            =   120
         TabIndex        =   92
         Top             =   1560
         Width           =   2415
      End
      Begin VB.CheckBox chkEnableAF 
         Caption         =   "AF Degaussing Coils"
         Height          =   375
         Left            =   120
         TabIndex        =   93
         Top             =   2040
         Width           =   1935
      End
      Begin VB.CheckBox chkEnableSusceptibility 
         Caption         =   "Susceptibility Coil"
         Height          =   375
         Left            =   120
         TabIndex        =   96
         Top             =   3240
         Width           =   2415
      End
      Begin VB.CheckBox chkEnableIRMBackfield 
         Caption         =   "IRM Backfield Pulse"
         Height          =   375
         Left            =   120
         TabIndex        =   90
         Top             =   720
         Width           =   2052
      End
      Begin VB.Line Line4 
         X1              =   0
         X2              =   2040
         Y1              =   3660
         Y2              =   3660
      End
      Begin VB.Line Line3 
         X1              =   0
         X2              =   2040
         Y1              =   3165
         Y2              =   3165
      End
      Begin VB.Line Line2 
         X1              =   0
         X2              =   2040
         Y1              =   1995
         Y2              =   1995
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   2040
         Y1              =   1485
         Y2              =   1485
      End
      Begin VB.Label lblSettings 
         Caption         =   "ARM/IRM  DAQ Board/Channel Settings"
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
         Index           =   54
         Left            =   2400
         TabIndex        =   169
         Top             =   120
         Width           =   3615
      End
   End
   Begin VB.Frame frameOptions 
      BorderStyle     =   0  'None
      Caption         =   "Vacuum"
      Height          =   4575
      Index           =   11
      Left            =   240
      TabIndex        =   114
      Top             =   600
      Width           =   6615
      Begin VB.Frame Frame4 
         Caption         =   "Temp. Sensor #1"
         Height          =   1455
         Left            =   3960
         TabIndex        =   194
         Top             =   600
         Width           =   2532
         Begin VB.ComboBox cmbAnalogT1Board 
            Height          =   315
            Left            =   240
            TabIndex        =   121
            Top             =   480
            Width           =   2052
         End
         Begin VB.ComboBox cmbAnalogT1Chan 
            Height          =   315
            Left            =   960
            TabIndex        =   122
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label lblSettings 
            Caption         =   "DAQ Board:"
            Height          =   255
            Index           =   79
            Left            =   240
            TabIndex        =   196
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label lblSettings 
            Caption         =   "Channel:"
            Height          =   255
            Index           =   80
            Left            =   240
            TabIndex        =   195
            Top             =   960
            Width           =   735
         End
      End
      Begin VB.TextBox txtToffset 
         Height          =   285
         Left            =   1800
         TabIndex        =   118
         Top             =   3480
         Width           =   1215
      End
      Begin VB.TextBox txtTslope 
         Height          =   285
         Left            =   1800
         TabIndex        =   119
         Top             =   3960
         Width           =   1215
      End
      Begin VB.ComboBox cmbTunits 
         Height          =   315
         Left            =   1680
         TabIndex        =   115
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox txtThot 
         Height          =   285
         Left            =   2160
         TabIndex        =   116
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox txtTmax 
         Height          =   285
         Left            =   2160
         TabIndex        =   117
         Top             =   1800
         Width           =   855
      End
      Begin VB.CheckBox checkAnalogT2 
         Caption         =   "Enable Transverse Temp. Sensor (T2)"
         Height          =   432
         Left            =   4080
         TabIndex        =   123
         Top             =   2400
         Width           =   1935
      End
      Begin VB.CheckBox checkAnalogT1 
         Caption         =   "Enable Axial Temp. Sensor (T1)"
         Height          =   432
         Left            =   4080
         TabIndex        =   120
         Top             =   240
         Width           =   1815
      End
      Begin VB.Frame Frame3 
         Caption         =   "Temp. Sensor #2"
         Height          =   1455
         Left            =   3960
         TabIndex        =   191
         Top             =   2880
         Width           =   2532
         Begin VB.ComboBox cmbAnalogT2Chan 
            Height          =   315
            Left            =   960
            TabIndex        =   125
            Top             =   960
            Width           =   1335
         End
         Begin VB.ComboBox cmbAnalogT2Board 
            Height          =   315
            Left            =   240
            TabIndex        =   124
            Top             =   480
            Width           =   2052
         End
         Begin VB.Label lblSettings 
            Caption         =   "Channel:"
            Height          =   255
            Index           =   82
            Left            =   240
            TabIndex        =   193
            Top             =   960
            Width           =   735
         End
         Begin VB.Label lblSettings 
            Caption         =   "DAQ Board:"
            Height          =   255
            Index           =   81
            Left            =   240
            TabIndex        =   192
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Label lblSettings 
         Caption         =   "Meas. T  =  T Offset   -   (Transducer Volt.)                                               x  Temp. / Volt."
         Height          =   615
         Index           =   74
         Left            =   120
         TabIndex        =   245
         Top             =   2880
         Width           =   3375
      End
      Begin VB.Label lblSettings 
         Caption         =   "Temperature Offset:"
         Height          =   255
         Index           =   75
         Left            =   120
         TabIndex        =   244
         Top             =   3480
         Width           =   1815
      End
      Begin VB.Label lblSettings 
         Caption         =   "Slope (Temp. / Volts):"
         Height          =   255
         Index           =   77
         Left            =   120
         TabIndex        =   243
         Top             =   3960
         Width           =   1575
      End
      Begin VB.Label lblToffsetUnits 
         Caption         =   "°C"
         Height          =   255
         Left            =   3120
         TabIndex        =   242
         Top             =   3480
         Width           =   255
      End
      Begin VB.Label lblTslopeUnits 
         Caption         =   "°C / V"
         Height          =   255
         Left            =   3120
         TabIndex        =   241
         Top             =   3960
         Width           =   495
      End
      Begin VB.Label lblSettings 
         Caption         =   "Temp. Sensor Units:"
         Height          =   255
         Index           =   71
         Left            =   120
         TabIndex        =   240
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label lblSettings 
         Caption         =   """Hot"" Alarm Temperature:"
         Height          =   255
         Index           =   72
         Left            =   120
         TabIndex        =   239
         Top             =   1320
         Width           =   2055
      End
      Begin VB.Label lblTHotUnits 
         Caption         =   "°C"
         Height          =   255
         Left            =   3120
         TabIndex        =   238
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label lblSettings 
         Caption         =   "Max Allowed Temperature:"
         Height          =   255
         Index           =   73
         Left            =   120
         TabIndex        =   237
         Top             =   1800
         Width           =   2055
      End
      Begin VB.Label lblTMaxUnits 
         Caption         =   "°C"
         Height          =   255
         Left            =   3120
         TabIndex        =   236
         Top             =   1800
         Width           =   255
      End
   End
   Begin VB.Label lblSettings 
      Height          =   255
      Index           =   76
      Left            =   360
      TabIndex        =   291
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Line Line5 
      X1              =   0
      X2              =   1920
      Y1              =   0
      Y2              =   0
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim activeRow As Integer
Dim ActiveCol As Integer
Dim SelectedGrid As String
Dim PriorListIndex As Long
Dim isUserChange As Boolean
Dim LastTUnits As String
Dim CurTunits As String

Dim isADwinRampSettings_dirty As Boolean

Dim TempMaxRampUpTime As Double
Dim TempMinRampUpTime As Double

Dim CurrentCell(2) As Integer
Dim CurrentCellPos(2) As Single

Dim interpolation_ranges As InterpolationRanges
Dim delta_positions As XYCup_Positions

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Sub ApplyCupPositionSettings()

    With Me.XYHolePositionsFlexGrid
    
        If .Rows <= 1 And .Cols <> 3 Then Exit Sub
    
        Dim i As Integer
    
        For i = 1 To .Rows - 1
        
            .row = i
            .Col = 1
            
            Dim pos_string As String
            Dim filter As RegExp
            
            Set filter = New RegExp
            filter.Pattern = "[^0-9-]"
            pos_string = filter.Replace(.text, "")
                    
            modConfig.XYTablePositions(i - 1, 0) = CLng(pos_string)
        
            .Col = 2
            pos_string = filter.Replace(.text, "")
            
            modConfig.XYTablePositions(i - 1, 1) = CLng(pos_string)
            
        Next i
        
    End With


End Sub

Private Sub applySettings()
        
    Dim i As Integer
    
    UseXYTableAPS = ckUseXYTableAPS.Value
    
    If UseXYTableAPS Then ApplyCupPositionSettings
    
    HoleSlotNum = Int(val(HomeHoleLocationBox.text))
    SlotMax = Int(val(SlotMaxBox.text))
    
    ZeroPos = Int(val(txtZeroPos))
    MeasPos = Int(val(txtMeasPos))
    AFPos = Int(val(txtAFPos))
    IRMPos = Int(val(txtIRMPos))
    SCoilPos = Int(val(txtSCoilPos))
    FloorPos = Int(val(txtFloorPos))
    MinUpDownPos = Int(val(txtMinUpDownPos))
    
    SampleBottom = Int(val(txtSampleBottom))
    SampleTop = Int(val(txtSampleTop))
    
    SampleHoleAlignmentOffset = val(txtSampHoleAlignOffset)
    LiftSpeedSlow = Int(val(txtLiftSpeedSlow))
    LiftSpeedNormal = Int(val(txtLiftSpeedNormal))
    LiftSpeedFast = Int(val(txtLiftSpeedFast))
    LiftAcceleration = Int(val(txtLiftAcceleration))
    ChangerSpeed = Int(val(txtChangerSpeed))
    TurnerSpeed = Int(val(txtTurnerSpeed))
    
    SCurveFactor = Int(val(txtSCurveFactor))
    TurningMotorFullRotation = Int(val(txtTurningMotorFullRotation))
    TurningMotor1rps = Int(val(txtTurningMotor1rps))
    
    UpDownMotor1cm = val(txtUpDownMotor1cm)
    TrayOffsetAngle = val(txtTrayOffsetAngle)
    
    PickupTorqueThrottle = val(txtPickupTorqueThrottle)
    UpDownTorqueFactor = val(txtUpDownTorquefactor)
    UpDownMaxTorque = val(txtUpDownMaxTorque)
    
    ZCal = val(txtZCal)
    XCal = val(txtXCal)
    YCal = val(txtYCal)
    RangeFact = val(txtRangeFact)
    ReadDelay = Int(val(txtReadDelay)) ' (March 2008 L Carporzen) Read delay
    
    SusceptibilitySettings = txtSusceptibilitySettings
    SusceptibilityMomentFactorCGS = val(txtSusceptibilityMomentFactorCGS)
    
    AFSystem = Me.cmbAFSystem.List(cmbAFSystem.ListIndex)
    AFUnits = Me.cmbAFUnits.List(cmbAFUnits.ListIndex)
    AFDelay = val(cmbAFDelay)
    AFRampRate = val(cmbAFRampRate)
    AfAxialCoord = cmbAFAxialCoord
    AfAxialMax = val(txtAFAxialMax)
    AfAxialMin = val(txtAFAxialMin)
    AfTransCoord = cmbAFTransCoord
    AfTransMin = val(txtAFTransMin)
    AfTransMax = val(txtAFTransMax)
    
    MinRampDown_NumPeriods = val(txtAFMinRampDownNumPeriods)
    MaxRampDown_NumPeriods = val(txtAFMaxRampDownNumPeriods)
    MinRampUpTime_ms = val(txtAFMinRampUpTime)
    MaxRampUpTime_ms = val(txtAFMaxRampUpTime)
    AxialRampUpVoltsPerSec = val(txtAFAxialRampUpVoltsPerSec)
    TransRampUpVoltsPerSec = val(txtAFTransverseRampUpVoltsPerSec)
    RampDownNumPeriodsPerVolt = val(txtAFRampDownPeriodsPerVolt)
    HoldAtPeakField_NumPeriods = CLng(txtPeakPeriods)
        
    isADwinRampSettings_dirty = False
        
    'Update temperature settings
    'Update the units
    If cmbTunits.ListIndex = 0 Then modConfig.Tunits = "C"
    If cmbTunits.ListIndex = 1 Then modConfig.Tunits = "F"
    If cmbTunits.ListIndex = 2 Then modConfig.Tunits = "K"
    
    'Update the "Hot" alarm temperature and the Max temperature
    modConfig.Thot = CInt(txtThot)
    modConfig.Tmax = CInt(txtTmax)
    
    'Update the Temperature slope and offset conversion factors
    modConfig.TSlope = val(txtTslope)
    modConfig.Toffset = val(txtToffset)
    
    'Save the IRM system
    'Load the IRM system
    If Me.cmbIRMSystem.ListIndex = 1 Then
    
        modConfig.IRMSystem = "Matsusada"
    
    ElseIf Me.cmbIRMSystem.ListIndex = 0 Then
        
        modConfig.IRMSystem = "Old"
        
    ElseIf cmbIRMSystem.ListIndex = 2 Then
    
        modConfig.IRMSystem = "ASC"
        
    End If
    
    modConfig.AscSetVoltageMaxBoostMultiplier = GetIRMBoostMultiplier_FromPercentage(val(Me.txtMaxIrmVoltageOut_BoostPercentage.text))
    modConfig.AscSetVoltageMinBoostMultiplier = GetIRMBoostMultiplier_FromPercentage(val(Me.txtMinIrmVoltageOut_BoostPercentage.text))
    
    modConfig.TrimOnTrue = Me.optTrimOnTrue.Value
    
    IRMAxis = cmbIRMAxis
    IRMBackfieldAxis = cmbIRMBackfieldAxis
    PulseAxialMax = val(txtPulseAxialMax)
    PulseAxialMin = val(txtPulseAxialMin)
    PulseTransMax = val(txtPulseTransMax)
    PulseTransMin = val(txtPulseTransMin)
    IRMAxialVoltMax = val(txtAxialMaxCapVoltage)
    
    If AxialTransMaxCapVoltsSame = False Then IRMTransVoltMax = val(txtTransMaxCapVoltage)
            
    PulseVoltMax = val(txtPulseVoltMax)
    PulseReturnMCCVoltConversion = val(txtPulseReturnMCCVoltConversion)
    PulseMCCVoltConversion = txtPulseMCCVoltConversion
    
    ARMMax = val(txtARMMax)
    ARMVoltGauss = val(txtARMVoltGauss)
    ARMVoltMax = val(txtARMVoltMax)
       
'----------------------------------------------------------------------------------------'
'----------------------------------------------------------------------------------------'
'       May, 2010
'       Isaac Hilburn
'
'       Modifications to the export of DAQ Channel settings from the frmSettings window
'       to the global channel object variables
'
'       See modConfig for more details
'----------------------------------------------------------------------------------------'

    'Need to export the Board and Channel combo-boxes to the appropriate
    'Channel objects (declared as stand-alone global variables, or declared
    'as channel objects attached to a wave object)
    
    'Analog Output
    
    'Save the AF ADWIN Ramp voltage Analog output channel
    Set modConfig.AFRampChan = SaveDAQSetting(Me.cmbAFRampBoard, _
                                              Me.cmbAFRampChan, _
                                              "AO", _
                                              "AF ADWIN Ramp Analog Output")
                                              
                                              
    'Set the Channel attached to the AFRAMPUP and AFRAMPDOWN waveform objects
    Set WaveForms("AFRAMPUP").Chan = modConfig.AFRampChan
    Set WaveForms("AFRAMPDOWN").Chan = modConfig.AFRampChan
    
    'Save the ARM Voltage Analog Output channel
    Set modConfig.ARMVoltageOut = SaveDAQSetting(Me.cmbARMVoltageOutBoard, _
                                                 Me.cmbARMVoltageOutChan, _
                                                 "AO", _
                                                 "ARM Voltage Analog Output")
                                                 
    'Save the IRM Voltage Analog Output channel
    Set modConfig.IRMVoltageOut = SaveDAQSetting(Me.cmbIRMVoltageOutBoard, _
                                                 Me.cmbIRMVoltageOutChan, _
                                                 "AO", _
                                                 "IRM Voltage Analog Output")
                                                 
    'Analog Input
    
    'Save the AF ADWIN Monitor voltage analog input channel
    Set modConfig.AFMonitorChan = SaveDAQSetting(Me.cmbAFMonitorBoard, _
                                                 Me.cmbAFMonitorChan, _
                                                 "AI", _
                                                 "AF ADWIN Monitor Analog Input")
                                                 
    'Set the Channel attached to the AFMONITOR waveform object
    Set WaveForms("AFMONITOR").Chan = modConfig.AFMonitorChan
                                                 
    'Save the Alternate AF monitor voltage analog input channel
    Set modConfig.AltAFMonitorChan = SaveDAQSetting(Me.cmbAltAFMonitorBoard, _
                                                    Me.cmbAltAFMonitorChan, _
                                                    "AI", _
                                                    "Alternate AF Monitor Analog Input")
                                                    
    'Set the Channel attached to the AFMONITOR waveform object
    Set WaveForms("ALTAFMONITOR").Chan = modConfig.AltAFMonitorChan
                                                        
    'Save the IRM Capacitor Voltage Analog Input channel
    Set modConfig.IRMCapacitorVoltageIn = _
                    SaveDAQSetting(Me.cmbIRMCapacitorVoltageInBoard, _
                                   Me.cmbIRMCapacitorVoltageInChan, _
                                   "AI", _
                                   "IRM Capacitor Voltage Analog Input")
                                   
    'Save the IRM Ready status Digital Input channel
    Set modConfig.IRMPowerAmpVoltageIn = SaveDAQSetting( _
                                            Me.cmbIRMPowerAmpVoltageInBoard, _
                                            Me.cmbIRMPowerAmpVoltageInChan, _
                                            "AI", _
                                            "IRM Power Amp Voltage Analog Input")
                                   
    'Save the IRM Monitor Analog Input channel
    Set modConfig.IRMMonitor = SaveDAQSetting(Me.cmbIRMMonitorBoard, _
                                              Me.cmbIRMMonitorChan, _
                                              "AI", _
                                             "IRM Monitor Analog Input")
                                             
    'Set the channel associated with the IRM Monitor waveform object
    Set WaveForms("IRMMONITOR").Chan = modConfig.IRMMonitor
    
    'Save the AF Coil Temperature Sensor #1 Analog Input channel
    Set modConfig.AnalogT1 = SaveDAQSetting(Me.cmbAnalogT1Board, _
                                            Me.cmbAnalogT1Chan, _
                                            "AI", _
                                            "AF Coil Temperature Sensor #1 Analog Input")
                                            
    'Save the AF Coil Temperature Sensor #2 Analog Input channel
    Set modConfig.AnalogT2 = SaveDAQSetting(Me.cmbAnalogT2Board, _
                                            Me.cmbAnalogT2Chan, _
                                            "AI", _
                                            "AF Coil Temperature Sensor #2 Analog Input")
                                            
    'Digital Output
    
    'Save the ARM Set TTL Digital Output channel
    Set modConfig.ARMSet = SaveDAQSetting(Me.cmbARMSetBoard, _
                                            Me.cmbARMSetChan, _
                                            "DO", _
                                            "ARM Set TTL Digital Output")
                                            
    'Save the IRM Fire TTL Digital Output channel
    Set modConfig.IRMFire = SaveDAQSetting(Me.cmbIRMFireBoard, _
                                            Me.cmbIRMFireChan, _
                                            "DO", _
                                            "IRM Fire TTL Digital Output")
                                                 
    'Save the IRM Trim TTL Digital Output channel
    Set modConfig.IRMTrim = SaveDAQSetting(Me.cmbIRMTrimBoard, _
                                            Me.cmbIRMTrimChan, _
                                            "DO", _
                                            "IRM Trim TTL Digital Output")

    'Save the IRM Low-Field Relay TTL Digital Output channel
    Set modConfig.IRMRelay = SaveDAQSetting(Me.cmbIRMRelayBoard, _
                                            Me.cmbIRMRelayChan, _
                                            "DO", _
                                            "IRM Low-Field Relay TTL Digital Output")
                                            
    'Save the AF Axial Relay TTL Digital Output channel
    Set modConfig.AFAxialRelay = SaveDAQSetting(Me.cmbAxialRelayBoard, _
                                                Me.cmbAxialRelayChan, _
                                                "DO", _
                                                "AF Axial Relay TTL Digital Output")

    'Save the AF Tranverse Relay TTL Digital Output channel
    Set modConfig.AFTransRelay = SaveDAQSetting(Me.cmbTransRelayBoard, _
                                                Me.cmbTransRelayChan, _
                                                "DO", _
                                                "AF Transverse Relay TTL Digital Output")

    'Save the Vacuum Motor Toggle TTL Digital Output channel
    Set modConfig.MotorToggle = SaveDAQSetting(Me.cmbMotorToggleBoard, _
                                               Me.cmbMotorToggleChan, _
                                               "DO", _
                                               "Vacuum Motor Toggle TTL Digital Output")
                                               
    'Save the Vacuum Motor Toggle TTL Digital Output channel
    Set modConfig.VacuumToggleA = SaveDAQSetting(Me.cmbVacuumToggleABoard, _
                                               Me.cmbVacuumToggleAChan, _
                                               "DO", _
                                               "Vacuum Toggle A TTL Digital Output")
                                               
    'Save the Degausser Cooler Digital Output channel
    Set modConfig.DegausserToggle = SaveDAQSetting(Me.cmbDegausserCoolerBoard, _
                                               Me.cmbDegausserCoolerChan, _
                                               "DO", _
                                               "Degausser Cooler TTL Digital Output")
                                               

    CmdHomeToTop = Int(val(txtCmdHometoTop))
    CmdSamplePickup = Int(val(txtCmdSamplePickup))
    MotorIDTurning = Int(val(txtMotorIDTurning))
    MotorIDChanger = Int(val(txtMotorIDChanger))
    MotorIDChangerY = Int(val(txtMotorIDChangerY))
    MotorIDUpDown = Int(val(txtMotorIDUpDown))
    DropoffVacuumDelay = val(txtDropoffVacuumDelay)
    DoVacuumReset = chkDoVacuumReset.Value
    DoDegausserCooling = chkDegausserAirCooler.Value
    EnableAxialIRM = chkEnableIRMAxial.Value
    EnableTransIRM = chkEnableIRMTrans.Value
    EnableIRMBackfield = chkEnableIRMBackfield.Value
    EnableARM = chkEnableARM.Value
    EnableAF = chkEnableAF.Value
    EnableAFAnalysis = chkEnableAFAnalysis.Value
    EnableSusceptibility = chkEnableSusceptibility.Value
    EnableVacuum = chkEnableVacuum.Value
    EnableDegausserCooler = chkEnableDegausserCooler.Value
    
    'Save whether the Analog temperature sensors are enabled
    EnableT1 = checkAnalogT1.Value
    EnableT2 = checkAnalogT2.Value
       
End Sub

'Private Sub MarkAsUnchecked(ByVal start_row As Long, ByVal end_row As Long)
'
'    If start_row <= 0 Or start_row > end_row Then Exit Sub
'
'    If start_row < 2 Then start_row = 2
'
'    With Me.XYHolePositionsFlexGrid
'
'        Dim i As Integer
'
'        If end_row > .Rows - 1 Then end_row = .Rows - 1
'
'        For i = start_row To end_row
'
'            .row = i
'            .Col = 3
'            Set .CellPicture = Me.picUnchecked.Picture
'
'        Next i
'
'    End With
'
'End Sub

Private Function Atan2(ByVal Y As Double, ByVal X As Double) As Double

If Y > 0 Then
    If X >= Y Then
        Atan2 = Atn(Y / X)
    ElseIf X <= -Y Then
        Atan2 = Atn(Y / X) + Pi
    Else
        Atan2 = Pi / 2 - Atn(X / Y)
    End If
Else
    If X >= -Y Then
        Atan2 = Atn(Y / X)
    ElseIf X <= Y Then
        Atan2 = Atn(Y / X) - Pi
    Else
        Atan2 = -Atn(X / Y) - Pi / 2
    End If
End If

End Function

Private Sub BubbleSortInterpolationRanges()

    Dim N As Integer, i As Integer, j As Integer

    N = interpolation_ranges.Count

    For i = (N - 1) To 1 Step -1

        For j = 0 To i - 1

            Dim max_j_value As Long
            Dim max_jplusone_value As Long

            max_j_value = MaxL(interpolation_ranges(j).StartRow, _
                               interpolation_ranges(j).EndRow)
                               
            max_jplusone_value = MaxL(interpolation_ranges(j + 1).StartRow, _
                                      interpolation_ranges(j + 1).EndRow)


            If max_j_value > max_jplusone_value Then
            
                Swap interpolation_ranges(j), interpolation_ranges(j + 1)
            
            End If

        Next j

    Next i

End Sub

Private Sub checkAnalogT1_Click()

    Dim isEnabled As Boolean

    If checkAnalogT1.Value = Checked Then
    
        isEnabled = True
        
    Else
    
        isEnabled = False
        
    End If
    
    Me.cmbAnalogT1Board.Enabled = isEnabled
    Me.cmbAnalogT1Chan.Enabled = isEnabled

End Sub

Private Sub checkAnalogT2_Click()

    Dim isEnabled As Boolean

    If checkAnalogT2.Value = Checked Then
    
        isEnabled = True
        
    Else
    
        isEnabled = False
        
    End If
    
    Me.cmbAnalogT2Board.Enabled = isEnabled
    Me.cmbAnalogT2Chan.Enabled = isEnabled

End Sub

Private Sub ClearBoardComboBoxes()

    'Clear the board combo-boxes
    Me.cmbAnalogT1Board.Clear
    Me.cmbAnalogT2Board.Clear
    Me.cmbARMSetBoard.Clear
    Me.cmbARMVoltageOutBoard.Clear
    Me.cmbAxialRelayBoard.Clear
    Me.cmbIRMCapacitorVoltageInBoard.Clear
    Me.cmbIRMFireBoard.Clear
    Me.cmbIRMRelayBoard.Clear
    Me.cmbIRMPowerAmpVoltageInBoard.Clear
    Me.cmbIRMTrimBoard.Clear
    Me.cmbIRMVoltageOutBoard.Clear
    Me.cmbMotorToggleBoard.Clear
    Me.cmbTransRelayBoard.Clear
    Me.cmbVacuumToggleABoard.Clear
    Me.cmbDegausserCoolerBoard.Clear
    Me.cmbIRMMonitorBoard.Clear
    Me.cmbAFMonitorBoard.Clear
    Me.cmbAltAFMonitorBoard.Clear
    Me.cmbAFRampBoard.Clear

End Sub

Private Sub ClearChannelComboBoxes()

    'Clear all the combo-boxes
    Me.cmbAnalogT1Chan.Clear
    Me.cmbAnalogT2Chan.Clear
    Me.cmbARMSetChan.Clear
    Me.cmbARMVoltageOutChan.Clear
    Me.cmbAxialRelayChan.Clear
    Me.cmbIRMCapacitorVoltageInChan.Clear
    Me.cmbIRMFireChan.Clear
    Me.cmbIRMRelayChan.Clear
    Me.cmbIRMPowerAmpVoltageInChan.Clear
    Me.cmbIRMTrimChan.Clear
    Me.cmbIRMVoltageOutChan.Clear
    Me.cmbMotorToggleChan.Clear
    Me.cmbTransRelayChan.Clear
    Me.cmbVacuumToggleAChan.Clear
    Me.cmbDegausserCoolerChan.Clear
    Me.cmbIRMMonitorChan.Clear
    Me.cmbAFMonitorChan.Clear
    Me.cmbAltAFMonitorChan.Clear
    Me.cmbAFRampChan.Clear

End Sub

Private Sub cmbAFMonitorBoard_Click()

    If cmbAFMonitorBoard.ListIndex <> PriorListIndex And _
       PriorListIndex <> -1 _
    Then

        LoadChannelComboBox cmbAFMonitorChan, _
                            AFMonitorChan, _
                            cmbAFMonitorBoard
                            
                            
        
                            
    ElseIf PriorListIndex = -1 Then
    
        LoadChannelComboBox cmbAFMonitorChan, _
                            AFMonitorChan
                            
    End If
        
    PriorListIndex = cmbAFMonitorBoard.ListIndex

End Sub

Private Sub cmbAFRampBoard_Click()

    If cmbAFRampBoard.ListIndex <> PriorListIndex And _
       PriorListIndex <> -1 _
    Then

        LoadChannelComboBox cmbAFRampChan, _
                            AFRampChan, _
                            cmbAFRampBoard
                            
                            
        
                            
    ElseIf PriorListIndex = -1 Then
    
        LoadChannelComboBox cmbAFRampChan, _
                            AFRampChan
                            
    End If
        
    PriorListIndex = cmbAFRampBoard.ListIndex

End Sub

Private Sub cmbAFSystem_Click()

    Dim UserResp As Long

    If isUserChange = False Then
    
        isUserChange = True
        Exit Sub
        
    End If

    If AFSystem = cmbAFSystem Then Exit Sub

    'Inform the user about the MASSIVE change to their system changing this field will cause
    UserResp = MsgBox("Changing the AF System will immediately cause MASSIVE changes " & _
                      "to your settings " & _
                      "and your Rock-magnetics system." & vbNewLine & vbNewLine & _
                      "Are you sure you want to do this?", _
                      vbYesNo, _
                      "AF System Change!")
                        
    If UserResp = vbYes Then
        
        'Deactivate the cmbAFSystem_Click() event
        isUserChange = False
        
        'User has selected to change the AF System value
        SetAFSystem cmbAFSystem.List(cmbAFSystem.ListIndex)
        
    Else
    
        'Deactivate the cmbAFSystem_Click() event
        isUserChange = False
    
        'User has selected to revert to the prior AF system value
        SetAFSystem AFSystem
        
    End If
    
    isUserChange = True

End Sub

Private Sub cmbAFUnits_Click()

    If cmbAFUnits.ListIndex <> PriorListIndex Then
    
        'Update the global units variable
        If modConfig.AFUnits <> cmbAFUnits.List(cmbAFUnits.ListIndex) Then
            
            modConfig.AFUnits = cmbAFUnits.List(cmbAFUnits.ListIndex)
            
            'Convert all of the units values in all the forms of this program
            ConvertFieldValues Me
            
        End If
        
        PriorListIndex = cmbAFUnits.ListIndex
        
    End If
    
End Sub

Private Sub cmbAFUnits_GotFocus()

    PriorListIndex = cmbAFUnits.ListIndex

End Sub

Private Sub cmbAltAFMonitorBoard_Click()

    If cmbAltAFMonitorBoard.ListIndex <> PriorListIndex And _
       PriorListIndex <> -1 _
    Then

        LoadChannelComboBox cmbAltAFMonitorChan, _
                            AltAFMonitorChan, _
                            cmbAltAFMonitorBoard
                            
                            
        
                            
    ElseIf PriorListIndex = -1 Then
    
        LoadChannelComboBox cmbAltAFMonitorChan, _
                            AltAFMonitorChan
                            
    End If
        
    PriorListIndex = cmbAltAFMonitorBoard.ListIndex

End Sub

Private Sub cmbAnalogT1Board_Click()

    If cmbAnalogT1Board.ListIndex <> PriorListIndex And _
       PriorListIndex <> -1 _
    Then

        LoadChannelComboBox cmbAnalogT1Chan, _
                            AnalogT1, _
                            cmbAnalogT1Board
                            
                            
        
                            
    ElseIf PriorListIndex = -1 Then
    
        LoadChannelComboBox cmbAnalogT1Chan, _
                            AnalogT1
                            
    End If
        
    PriorListIndex = cmbAnalogT1Board.ListIndex
    
End Sub

Private Sub cmbAnalogT1Board_GotFocus()

    PriorListIndex = cmbAnalogT1Board.ListIndex

End Sub

Private Sub cmbAnalogT2Board_Click()
    
    If cmbAnalogT2Board.ListIndex <> PriorListIndex And _
       PriorListIndex <> -1 _
    Then

        LoadChannelComboBox cmbAnalogT2Chan, _
                            AnalogT2, _
                            cmbAnalogT2Board
                            
                            
        
                            
    ElseIf PriorListIndex = -1 Then
    
        LoadChannelComboBox cmbAnalogT2Chan, _
                            AnalogT2
                            
    End If
        
    PriorListIndex = cmbAnalogT2Board.ListIndex

End Sub

Private Sub cmbAnalogT2Board_GotFocus()

    PriorListIndex = cmbAnalogT2Board.ListIndex

End Sub

Private Sub cmbARMSetBoard_Click()
    
    If cmbARMSetBoard.ListIndex <> PriorListIndex And _
       PriorListIndex <> -1 _
    Then

        LoadChannelComboBox cmbARMSetChan, _
                            ARMSet, _
                            cmbARMSetBoard
                            
                            
        
                            
    ElseIf PriorListIndex = -1 Then
    
        LoadChannelComboBox cmbARMSetChan, _
                            ARMSet
                            
    End If
        
    PriorListIndex = cmbARMSetBoard.ListIndex
    
End Sub

Private Sub cmbARMSetBoard_GotFocus()

    PriorListIndex = cmbARMSetBoard.ListIndex

End Sub

Private Sub cmbARMVoltageOutBoard_Click()

    If cmbARMVoltageOutBoard.ListIndex <> PriorListIndex And _
       PriorListIndex <> -1 _
    Then

        LoadChannelComboBox cmbARMVoltageOutChan, _
                            ARMVoltageOut, _
                            cmbARMVoltageOutBoard
                            
                            
        
                            
    ElseIf PriorListIndex = -1 Then
    
        LoadChannelComboBox cmbARMVoltageOutChan, _
                            ARMVoltageOut
                            
    End If
        
    PriorListIndex = cmbARMVoltageOutBoard.ListIndex
    
End Sub

Private Sub cmbARMVoltageOutBoard_GotFocus()

    PriorListIndex = cmbARMVoltageOutBoard.ListIndex

End Sub

Private Sub cmbAxialRelayBoard_Click()
    
    If cmbAxialRelayBoard.ListIndex <> PriorListIndex And _
       PriorListIndex <> -1 _
    Then

        LoadChannelComboBox cmbAxialRelayChan, _
                            AFAxialRelay, _
                            cmbAxialRelayBoard
                            
                            
        
                            
    ElseIf PriorListIndex = -1 Then
    
        LoadChannelComboBox cmbAxialRelayChan, _
                            AFAxialRelay
                            
    End If
        
    PriorListIndex = cmbAxialRelayBoard.ListIndex
    
End Sub

Private Sub cmbAxialRelayBoard_GotFocus()

    PriorListIndex = cmbAxialRelayBoard.ListIndex

End Sub

Private Sub cmbDegausserCoolerBoard_Click()

    If cmbDegausserCoolerBoard.ListIndex <> PriorListIndex And _
       PriorListIndex <> -1 _
    Then

        LoadChannelComboBox cmbDegausserCoolerChan, _
                            DegausserToggle, _
                            cmbDegausserCoolerBoard
                            
                            
        
                            
    ElseIf PriorListIndex = -1 Then
    
        LoadChannelComboBox cmbDegausserCoolerChan, _
                            DegausserToggle
                            
    End If
    
    PriorListIndex = cmbDegausserCoolerBoard.ListIndex

End Sub

Private Sub cmbDegausserCoolerBoard_GotFocus()

    PriorListIndex = cmbDegausserCoolerBoard.ListIndex

End Sub

Private Sub cmbIRMCapacitorVoltageInBoard_Click()

    If cmbIRMCapacitorVoltageInBoard.ListIndex <> PriorListIndex And _
       PriorListIndex <> -1 _
    Then

        LoadChannelComboBox cmbIRMCapacitorVoltageInChan, _
                            IRMCapacitorVoltageIn, _
                            cmbIRMCapacitorVoltageInBoard
                            
                            
        
                            
    ElseIf PriorListIndex = -1 Then
    
        LoadChannelComboBox cmbIRMCapacitorVoltageInChan, _
                            IRMCapacitorVoltageIn
                            
    End If
        
    PriorListIndex = cmbIRMCapacitorVoltageInBoard.ListIndex

End Sub

Private Sub cmbIRMCapacitorVoltageInBoard_GotFocus()

    PriorListIndex = cmbIRMCapacitorVoltageInBoard.ListIndex

End Sub

Private Sub cmbIRMFireBoard_Click()

    If cmbIRMFireBoard.ListIndex <> PriorListIndex And _
       PriorListIndex <> -1 _
    Then

        LoadChannelComboBox cmbIRMFireChan, _
                            IRMFire, _
                            cmbIRMFireBoard
                            
                            
        
                            
    ElseIf PriorListIndex = -1 Then
    
        LoadChannelComboBox cmbIRMFireChan, _
                            IRMFire
                            
    End If
        
    PriorListIndex = cmbIRMFireBoard.ListIndex

End Sub

Private Sub cmbIRMFireBoard_GotFocus()

    PriorListIndex = cmbIRMFireBoard.ListIndex

End Sub

Private Sub cmbIRMMonitorBoard_Click()

    If cmbIRMMonitorBoard.ListIndex <> PriorListIndex And _
       PriorListIndex <> -1 _
    Then

        LoadChannelComboBox cmbIRMMonitorChan, _
                            IRMMonitor, _
                            cmbIRMMonitorBoard
                            
                            
        
                            
    ElseIf PriorListIndex = -1 Then
    
        LoadChannelComboBox cmbIRMMonitorChan, _
                            IRMMonitor
                            
    End If
        
    PriorListIndex = cmbIRMMonitorBoard.ListIndex

End Sub

Private Sub cmbIRMMonitorBoard_GotFocus()

    PriorListIndex = cmbIRMMonitorBoard.ListIndex

End Sub

Private Sub cmbIRMPowerAmpVoltageInBoard_Click()

    If cmbIRMPowerAmpVoltageInBoard.ListIndex <> PriorListIndex And _
       PriorListIndex <> -1 _
    Then

        LoadChannelComboBox cmbIRMPowerAmpVoltageInChan, _
                            IRMPowerAmpVoltageIn, _
                            cmbIRMPowerAmpVoltageInBoard
                            
                            
        
                            
    ElseIf PriorListIndex = -1 Then
    
        LoadChannelComboBox cmbIRMPowerAmpVoltageInChan, _
                            IRMPowerAmpVoltageIn
                            
    End If
        
    PriorListIndex = cmbIRMPowerAmpVoltageInBoard.ListIndex
    
End Sub

Private Sub cmbIRMPowerAmpVoltageInBoard_GotFocus()

    PriorListIndex = cmbIRMPowerAmpVoltageInBoard.ListIndex
    
End Sub

Private Sub cmbIRMRelayBoard_Click()

    If cmbIRMRelayBoard.ListIndex <> PriorListIndex And _
       PriorListIndex <> -1 _
    Then

        LoadChannelComboBox cmbIRMRelayChan, _
                            IRMRelay, _
                            cmbIRMRelayBoard
                            
                            
        
                            
    ElseIf PriorListIndex = -1 Then
    
        LoadChannelComboBox cmbIRMRelayChan, _
                            IRMRelay
                            
    End If
        
    PriorListIndex = cmbIRMRelayBoard.ListIndex

End Sub

Private Sub cmbIRMRelayBoard_GotFocus()

    PriorListIndex = cmbIRMRelayBoard.ListIndex

End Sub

Private Sub cmbIRMSystem_Click()

    If Me.cmbIRMSystem.ListIndex = 2 Then
    
        EnableAscOnlyIRMControls
        
    Else
    
        DisableAscOnlyIRMControls
        
    End If

End Sub

Private Sub cmbIRMTrimBoard_Click()

    If cmbIRMTrimBoard.ListIndex <> PriorListIndex And _
       PriorListIndex <> -1 _
    Then

        LoadChannelComboBox cmbIRMTrimChan, _
                            IRMTrim, _
                            cmbIRMTrimBoard
                            
                            
        
                            
    ElseIf PriorListIndex = -1 Then
    
        LoadChannelComboBox cmbIRMTrimChan, _
                            IRMTrim
                            
    End If
        
    PriorListIndex = cmbIRMTrimBoard.ListIndex

End Sub

Private Sub cmbIRMTrimBoard_GotFocus()

    PriorListIndex = cmbIRMTrimBoard.ListIndex

End Sub

Private Sub cmbIRMVoltageOutBoard_Click()

    If cmbIRMVoltageOutBoard.ListIndex <> PriorListIndex And _
       PriorListIndex <> -1 _
    Then

        LoadChannelComboBox cmbIRMVoltageOutChan, _
                            IRMVoltageOut, _
                            cmbIRMVoltageOutBoard
                            
                            
        
                            
    ElseIf PriorListIndex = -1 Then
    
        LoadChannelComboBox cmbIRMVoltageOutChan, _
                            IRMVoltageOut
                            
    End If

    PriorListIndex = cmbIRMVoltageOutBoard.ListIndex

End Sub

Private Sub cmbIRMVoltageOutBoard_GotFocus()

    PriorListIndex = cmbIRMVoltageOutBoard.ListIndex

End Sub

Private Sub cmbMotorToggleBoard_Click()

    If cmbMotorToggleBoard.ListIndex <> PriorListIndex And _
       PriorListIndex <> -1 _
    Then

        LoadChannelComboBox cmbMotorToggleChan, _
                            MotorToggle, _
                            cmbMotorToggleBoard
                            
                            
        
                            
    ElseIf PriorListIndex = -1 Then
    
        LoadChannelComboBox cmbMotorToggleChan, _
                            MotorToggle
                            
    End If
    
    PriorListIndex = cmbMotorToggleBoard.ListIndex
    
End Sub

Private Sub cmbMotorToggleBoard_GotFocus()

    PriorListIndex = cmbMotorToggleBoard.ListIndex

End Sub

Private Sub cmbTransRelayBoard_Click()

    If cmbTransRelayBoard.ListIndex <> PriorListIndex And _
       PriorListIndex <> -1 _
    Then

        LoadChannelComboBox cmbTransRelayChan, _
                            AFTransRelay, _
                            cmbTransRelayBoard
                            
                            
        
                            
    ElseIf PriorListIndex = -1 Then
    
        LoadChannelComboBox cmbTransRelayChan, _
                            AFTransRelay
                            
    End If
    
    PriorListIndex = cmbTransRelayBoard.ListIndex
    
End Sub

Private Sub cmbTunits_Click()

    Dim CurTunits As String
    Dim LocalThot As Integer
    Dim LocalTmax As Integer
    Dim LocalTslope As Double
    Dim LocalToffset As Double
    
    CurTunits = GetCurTUnits

    'Check to see if this new units value is different from the current system units value
    If CurTunits <> LastTUnits Then
    
        'Load the four form values from the text-box controls on the form
        LocalThot = CInt(Me.txtThot)
        LocalTmax = CInt(Me.txtTmax)
        LocalTslope = val(Me.txtTslope)
        LocalToffset = val(Me.txtToffset)
    
        'Convert the values
        LocalThot = ConvertTemperature(LastTUnits, CurTunits, LocalThot)
        LocalTmax = ConvertTemperature(LastTUnits, CurTunits, LocalTmax)
        LocalToffset = ConvertTemperature(LastTUnits, CurTunits, LocalToffset)
        LocalTslope = ConvertSlope(LastTUnits, CurTunits, LocalTslope)
        
        'Now update the form text-box controls with the converted values
        Me.txtThot = Trim(Str(LocalThot))
        Me.txtTmax = Trim(Str(LocalTmax))
        Me.txtToffset = Trim(Str(LocalToffset))
        Me.txtTslope = Trim(Str(LocalTslope))
        
        'Change all of the units labels on the form
        'First check to see if curTunits is in Kelvin
        If CurTunits = "K" Then
            
            Me.lblTHotUnits = CurTunits
            Me.lblTMaxUnits = CurTunits
            Me.lblToffsetUnits = CurTunits
            Me.lblTslopeUnits = CurTunits & " / V"
                                    
        Else
        
            Me.lblTHotUnits = "°" & CurTunits
            Me.lblTMaxUnits = "°" & CurTunits
            Me.lblToffsetUnits = "°" & CurTunits
            Me.lblTslopeUnits = "°" & CurTunits & " / V"
        
        End If
                                
        'Now save the current units into the Last units variable
        LastTUnits = CurTunits
                            
    End If

End Sub

Private Sub cmbVacuumToggleABoard_Click()

    If cmbVacuumToggleABoard.ListIndex <> PriorListIndex And _
       PriorListIndex <> -1 _
    Then

        LoadChannelComboBox cmbVacuumToggleAChan, _
                            VacuumToggleA, _
                            cmbVacuumToggleABoard
                            
                            
        
                            
    ElseIf PriorListIndex = -1 Then
    
        LoadChannelComboBox cmbVacuumToggleAChan, _
                            VacuumToggleA
                            
    End If
    
    PriorListIndex = cmbVacuumToggleABoard.ListIndex

End Sub

Private Sub cmbVacuumToggleABoard_GotFocus()

    PriorListIndex = cmbVacuumToggleABoard.ListIndex

End Sub

Private Sub cmdAFFileSettings_Click()

    Load frmFileSave
    frmFileSave.Show

End Sub

Private Sub cmdCalAFCoils_Click()

    'Set the calibration form for AF mode
    'and show it
    frmCalibrateCoils.InAFMode = True
    Load frmCalibrateCoils
    frmCalibrateCoils.Show

End Sub

' (June 2007 L Carporzen) Form to adjust automatically the positions each time a new rod is installed.
' All we need is approximatly good (+-5 cm) positions in Paleomag.INI and the Bartington susceptibility standard (white cylinder placed on top of the MS2B box).
' NB that the susceptibility is very sensitive to the altitude in order to have reproducible measurement.
' The susceptibility position should be easy to define at less than 1 cm prior to running that sequence.
'
' (March 2008 L Carporzen) It can also correct the negative sample height that some systems are still using.
'
'(May 2010 I Hilburn) Added new documentation to the function and changed the Input
' box used for prompting the user to change the UpDownMotor1cm sign to a two button
' MsgBox
Public Sub cmdCalibNewRod_Click()

    Dim UserResponse As Long

    If UpDownMotor1cm < 0 And ZeroPos < 0 Or UpDownMotor1cm > 0 And ZeroPos > 0 Then
        
        '--------------------------------------------------------------------'
        '   Code Change
        '   5/7/2010
        '   Isaac Hilburn
        '
        '   Reason: Input box is not the appropriate dialog box to use
        '           for a sign change - instead use a Message Box
        '           in vbOKCancel mode and prompt the user to click "OK"
        '           to change the sign of the UpDown Motor conversion factor
        '           or "Cancel" to leave it alone.
        '--------------------------------------------------------------------'
        
        UserResponse = MsgBox("You can change the sign of the Sample Height " & _
                              Format$((SampleTop - SampleBottom) / _
                                      UpDownMotor1cm, "0.00") & _
                              " to positive values by changing the sign of the " & _
                              "Up/Down motor scale " & vbNewLine & _
                              "Current: " & vbTab & Trim(Str(UpDownMotor1cm)) & vbNewLine & _
                              "New: " & vbTab & Trim(Str(-UpDownMotor1cm)) & vbNewLine & _
                              "Click ""OK"" to agree to the change, or ""Cancel"" to keep " & _
                              "the current setting.", _
                              vbOKCancel, _
                              "Important!")
        
        'Check to see if the user clicked "OK"
        If UserResponse = vbOK Then
        
            'Update the global variable
            UpDownMotor1cm = -UpDownMotor1cm
            
        End If
        
        'Let the user know the change has been made
        'Comment: I Hilburn, 5/7/2010
        'I'd like to comment this out.
        MsgBox "The default sign of the Sample Height is now: " & _
               Format$((SampleTop - SampleBottom) / UpDownMotor1cm, "0.00") & "." & _
               vbNewLine & vbNewLine & "Click OK in the settings window to save it."
        
    End If
    
    'Comment: I Hilburn, 5/7/2010
    'This just seems dangerous.
    'I think this line should be commented out.  Why are we monkeying with the user's setting?
    SampleHoleAlignmentOffset = 0
    
    'Prompt user to load the Bartington susc. standard
    MsgBox "Please load the Bartington susceptibility standard sample face up " & _
           "into the Hole 1 of the Sample Changer belt. Leave free Hole 199!" & _
           vbNewLine & "The Bartington susceptibility meter should be in " & _
           "CGS units and set to range 1.0." & vbNewLine & _
           "We strongly recommand backing up the Paleomag.INI file..." & vbNewLine & vbNewLine & _
           "(It'd be a pity if your...um...INI file should have a...uh...tragic accident.)"
    
    'Bring to the front and Show the calibrate rod form
    frmCalRod.ZOrder
    frmCalRod.Show
    
    'Make sure the sample-hole alignment offset global variable has the current value
    SampleHoleAlignmentOffset = val(txtSampHoleAlignOffset)
    
End Sub

Private Sub cmdCalibrateIRMFields_Click()

    'Set the calibration form for IRM mode
    'and show it
    frmCalibrateCoils.InAFMode = False
    Load frmCalibrateCoils
    frmCalibrateCoils.Show
    frmCalibrateCoils.ZOrder 0

End Sub

Private Sub cmdCalibrateIRMVoltages_Click()
    
    'Set & load the IRM DAQ voltages calibration form
    'and show it
    Load frmIRM_VoltageCalibration
    frmIRM_VoltageCalibration.Show
    frmIRM_VoltageCalibration.ZOrder 0

End Sub

Private Sub cmdFlowControl_Click()

    If Prog_paused = True Then
    
        Flow_Resume
        
        cmdFlowControl.Caption = "Pause run"
        
    ElseIf Prog_paused = False Then
    
        Flow_Pause
    
        cmdFlowControl.Caption = "Resume run"
    
    End If

End Sub

'Private Sub cmdInterpolateBetweenCheckedPoints_Click()
'
'    If interpolation_ranges Is Nothing Then Exit Sub
'
'    If interpolation_ranges.Count = 0 Then Exit Sub
'
'    If interpolation_ranges(0).StartRow = -1 And _
'       interpolation_ranges(0).EndRow = -1 And _
'       interpolation_ranges.Count = 1 Then Exit Sub
'
'    Dim i As Integer
'    Dim pheta As Double
'
'    For i = 0 To interpolation_ranges.Count - 1
'
'        If interpolation_ranges(i).EndRow <> -1 And _
'           interpolation_ranges(i).StartRow <> -1 And _
'           interpolation_ranges(i).EndRow - interpolation_ranges(i).StartRow > 1 Then
'
'            Dim delta1 As XYCup
'            Dim delta2 As XYCup
'
'            Dim new1 As XYCup
'            Dim new2 As XYCup
'
'            Dim d As Double
'
'            Set delta1 = delta_positions(Trim(Str(interpolation_ranges(i).StartRow - 1)))
'            Set delta2 = delta_positions(Trim(Str(interpolation_ranges(i).EndRow - 1)))
'
'            With Me.XYHolePositionsFlexGrid
'
'                Set new1 = New XYCup
'                .row = interpolation_ranges(i).StartRow
'
'                .Col = 1
'                new1.x_pos = CLng(.text)
'
'                .Col = 2
'                new1.y_pos = CLng(.text)
'
'
'                Set new2 = New XYCup
'                .row = interpolation_ranges(i).EndRow
'
'                .Col = 1
'                new2.x_pos = CLng(.text)
'
'                .Col = 2
'                new2.y_pos = CLng(.text)
'
'            End With
'
'            d = Abs(new2.x_pos - new1.x_pos) + Abs(new2.y_pos - new1.y_pos)
'
'            If d <> 0 Then
'
'                Dim temp As Double
'
'                temp = Math.Sqr((delta1.x_pos) ^ 2 + (delta1.y_pos) ^ 2) / d
'
'                pheta = 2 * Atan2(temp, Math.Sqr(1 - temp ^ 2))
'
'            End If
'
'            If pheta <> 0 Then
'
'                'Need to go through all the rows between the start and end row
'                Dim j As Integer
'
'                For j = interpolation_ranges(i).StartRow + 1 To interpolation_ranges(i).EndRow - 1
'
'                    With Me.XYHolePositionsFlexGrid
'
'                        Dim X As Long
'                        Dim Y As Long
'
'                        .row = j
'                        .Col = 1
'                        X = CLng(.text)
'
'                        .Col = 2
'                        Y = CLng(.text)
'
'                        .Col = 1
'                        .text = Trim(Str(Round(X * Math.Cos(pheta) - Y * Math.Sin(pheta), 0)))
'
'                        .Col = 2
'                        .text = Trim(Str(Round(X * Math.Sin(pheta) + Y * Math.Cos(pheta), 0)))
'
'                    End With
'
'                Next j
'
'            End If
'
'        End If
'
'    Next i
'
'    'Clear the interpolation ranges
'    interpolation_ranges.Clear
'
'    'Clear the delta positions
'    delta_positions.Clear
'
'    'Reset all the pictures to unchecked check-boxes
'    MarkAsUnchecked 2, XYHolePositionsFlexGrid.Rows - 1
'
'End Sub

Public Sub cmdOKApplyCancel_Click(Index As Integer)
    
    Dim Response As VbMsgBoxResult
    
    'If user has clicked the 'OK' or 'Apply' buttons
    If Index < 2 Then
    
        Response = MsgBox("Incorrect values can break the system!" & vbCrLf & _
                          "Are you sure you want to make changes?", _
                          vbYesNo, _
                          "Warning!")
                          
        If Response = vbYes Then
            
            applySettings
            importSettings
            
            'If user clicked the 'OK' button, save the new global settings
            'to the .INI file
            If Index = 0 Then
                
                modConfig.Config_writeSettingstoINI
                
            End If
                
            Unload Me
            
        End If
        
    ElseIf Index = 2 Then
    
        'User has clicked cancel
        Unload Me
        
    End If
    
End Sub

Private Sub cmdOpenAFForm_Click()

    'Load & show the AF diagnostic / control form
    If AFSystem = "2G" Then
    
        Load frmAF_2G
        frmAF_2G.Show
        
    ElseIf AFSystem = "ADWIN" Then
    
        Load frmADWIN_AF
        frmADWIN_AF.Show
        
    End If

End Sub

Private Sub cmdOpenIRMARMForm_Click()

    'Load & show the IRM / ARM diagnostic / control form
    Load frmIRMARM
    frmIRMARM.Show

End Sub

Private Sub cmdTuneAF_Click()

    Load frmAFTuner
    frmAFTuner.Show

End Sub

Public Sub ConvertFieldValues(ByRef CallingForm As Form)

    

End Sub

Private Function ConvertSlope(ByVal OldUnits As String, _
                                    ByVal NewUnits As String, _
                                    ByVal OldValue As Double) As Double
                                    
    Dim TempD As Double
    
    'Celcius to Farenheit
    If OldUnits = "C" And NewUnits = "F" Then
    
        TempD = 9 / 5 * OldValue
        
    ElseIf OldUnits = "F" And NewUnits = "C" Then
    
        TempD = OldValue * 5 / 9
        
    ElseIf OldUnits = "F" And NewUnits = "K" Then
    
        TempD = OldValue * 5 / 9
    
    ElseIf OldUnits = "K" And NewUnits = "F" Then
    
        TempD = OldValue * 9 / 5
        
   End If
                                    
   'Return the new slope
   ConvertSlope = TempD
                                    
End Function

Private Function ConvertTemperature(ByVal OldUnits As String, _
                                    ByVal NewUnits As String, _
                                    ByVal OldValue As Variant) As Variant
                                    
    Dim TempD As Double
    
    'Celcius to Farenheit
    If OldUnits = "C" And NewUnits = "F" Then
    
        TempD = 9 / 5 * CDbl(OldValue) + 32
        
    ElseIf OldUnits = "C" And NewUnits = "K" Then
    
        TempD = CDbl(OldValue) + 273.15
        
    ElseIf OldUnits = "F" And NewUnits = "C" Then
    
        TempD = (CDbl(OldValue) - 32) * 5 / 9
        
    ElseIf OldUnits = "F" And NewUnits = "K" Then
    
        TempD = (CDbl(OldValue) - 32) * 5 / 9 + 273.15
    
    ElseIf OldUnits = "K" And NewUnits = "F" Then
    
        TempD = (CDbl(OldValue) - 273.15) * 9 / 5 + 32
        
    ElseIf OldUnits = "K" And NewUnits = "C" Then
    
        TempD = CDbl(OldValue) - 273.15
        
    End If
                                    
    'If the old value was an integer, then output an integer,
    'otherwise, output a double
    If VarType(OldValue) = vbInteger Then
        
            ConvertTemperature = CInt(TempD)
            
        Else
        
            ConvertTemperature = TempD
            
    End If
                                    
End Function

Private Sub DoApplyChangeToXForEntireRowOfCups(ByVal start_row_num As Integer, _
                                               ByVal delta_x_pos As Long)
                                               
    Dim old_x_pos As Long
    Dim i As Integer
    
    'Filter for bad values
    If start_row_num <= 1 Then Exit Sub
    
    If start_row_num >= XYHolePositionsFlexGrid.Rows - 1 Then Exit Sub
    
    '(Last cup in row was edited)
    If (start_row_num - 1) Mod 10 = 0 Then Exit Sub
    
    Dim from_row_num As Integer: from_row_num = start_row_num + 1
    Dim to_row_num As Integer: to_row_num = start_row_num + 10 - ((start_row_num - 1) Mod 10)
        
    For i = from_row_num To to_row_num
    
         XYHolePositionsFlexGrid.row = i
         XYHolePositionsFlexGrid.Col = 1
         old_x_pos = CLng(XYHolePositionsFlexGrid.text)
         XYHolePositionsFlexGrid.text = old_x_pos + delta_x_pos
         modConfig.XYTablePositions(i - 1, 0) = old_x_pos + delta_x_pos
             
    Next i
                                                   
    XYHolePositionsFlexGrid.row = start_row_num
                                                   
End Sub

Private Sub DoApplyChangeToYForEntireRowOfCups(ByVal start_row_num As Integer, _
                                               ByVal delta_y_pos As Long)
                                               
    Dim old_y_pos As Long
    Dim i As Integer
    
    'Filter for bad values
    If start_row_num <= 1 Then Exit Sub
    
    If start_row_num >= XYHolePositionsFlexGrid.Rows - 1 Then Exit Sub
    
    '(Last cup in row was edited)
    If (start_row_num - 1) Mod 10 = 0 Then Exit Sub
    
    Dim from_row_num As Integer: from_row_num = start_row_num + 1
    Dim to_row_num As Integer: to_row_num = start_row_num + 10 - ((start_row_num - 1) Mod 10)
        
    For i = from_row_num To to_row_num
    
         XYHolePositionsFlexGrid.row = i
         XYHolePositionsFlexGrid.Col = 2
         old_y_pos = CLng(XYHolePositionsFlexGrid.text)
         XYHolePositionsFlexGrid.text = old_y_pos + delta_y_pos
         modConfig.XYTablePositions(i - 1, 1) = old_y_pos + delta_y_pos
             
    Next i
    
    XYHolePositionsFlexGrid.row = start_row_num
                                                   
End Sub

Private Function FindActiveTab() As Integer

    Dim i As Integer
    Dim N As Integer
    
    'Set N = number of tabs
    N = Me.tbsOptions.Tabs.Count
    
    'Loop through them and find out which one is visible and enabled
    For i = 0 To N - 1
    
        If frameOptions(i).Enabled = True And _
           frameOptions(i).Visible = True _
        Then
        
            'This tab is active, return it's index
            FindActiveTab = i
            
            Exit Function
            
        End If
        
    Next i
    
    'Default, if no tab is selected, to the first tab
    FindActiveTab = 0

End Function

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

Public Sub Form_Load()
    
    '(March 11, 2011 - I Hilburn)
    'Update all the form's icons with the program Ico file
    SetFormIcon Me
    
    'Set the Flow control button caption
    If Prog_paused = True Then
    
        cmdFlowControl.Caption = "Resume run"
        
    ElseIf Prog_paused = False Then
    
        cmdFlowControl.Caption = "Pause run"
    
    End If
    
    Me.txtEditGridCell.Visible = False
    
    Set delta_positions = Nothing
    Set delta_positions = New XYCup_Positions
    
    'Center Settings form window in the Paleomag program screen
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
    
    'Set the tab for the ARM/IRM Channels tab strip
    tbsARMIRMChannels_Click
    
    'Load the options into the AF System selector combo-box
    cmbAFSystem.Clear
    cmbAFSystem.AddItem "2G", 0
    cmbAFSystem.AddItem "ADWIN", 1
    
    'Load the units options for the AF & IRM systems
    Me.cmbAFUnits.Clear
    Me.cmbAFUnits.AddItem "G", 0
        
    '(April 2010 - Isaac Hilburn, AF DAQ system mod)
    'If the AF system type is MCC or ADWIN, then hide the AFRampRate and the AFDelay
    'combo boxes.  Likewise, hide the AF axial coord and trans coord combo boxes
            
    'Load all the necessary values into the 2G variable combo-boxes
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
    
    'Load IRM system combo-box
    Me.cmbIRMSystem.Clear
    cmbIRMSystem.AddItem "Caltech Old", 0
    cmbIRMSystem.AddItem "Caltech - Matsusada", 1
    cmbIRMSystem.AddItem "ASC", 2
    
    'Load IRM Low-field & High-field calibration controls
    cmbIRMAxis.Clear
    cmbIRMAxis.AddItem "X"
    cmbIRMAxis.AddItem "Y"
    cmbIRMAxis.AddItem "Z"
    
    cmbIRMBackfieldAxis.Clear
    cmbIRMBackfieldAxis.AddItem "X"
    cmbIRMBackfieldAxis.AddItem "Y"
    cmbIRMBackfieldAxis.AddItem "Z"
    
    'Clear all the form DAQ Board & channel combo boxes
    ClearChannelComboBoxes
    ClearBoardComboBoxes
    
    'Check to see if the System DAQ boards have been imported from the INI file
    If ImportBoardsDone = True Then
    
        'Load all the possible Boards and channels into the combo boxes
        LoadBoardChanComboBoxes
        
    Else
    
        'I think it's time to pop-up an error
'        Err.Raise -616, _
'                  "frmSettings.Form_Load", _
'                  "Settings window unable to load and display the DAQ comm board. " & _
'                  vbNewLine & _
'                  "DAQ comm board settings have not been successfully imported from the " & _
'                  "Paleomag.ini file."


    End If
    
    'Now that the Board & Channel Combo boxes are loaded, can import the global settings
    'into the Settings form controls.
    importSettings
     
    'Now that settings are imported, can set the AF System
    SetAFSystem modConfig.AFSystem, False
     
    'Figure out which tab is active and make sure that that tabs
    'control is properly clicked in the display
    selectTab FindActiveTab + 1
     
    'refresh the form
    Me.refresh
    
    'Set isUserChange = True
    isUserChange = True
    
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    'if X & Y are over the calibration grid, call that event
    If Y >= Me.XYHolePositionsFlexGrid.Top And _
       Y <= Me.XYHolePositionsFlexGrid.Top + Me.XYHolePositionsFlexGrid.Height And _
       X >= Me.XYHolePositionsFlexGrid.Left And _
       X <= Me.XYHolePositionsFlexGrid.Left + Me.XYHolePositionsFlexGrid.Width And _
       Me.frameOptions(2).Visible = True _
    Then
    
        'Mouse is over the calibration grid
        'activate the calibration grid mousedown event
        XYHolePositionsFlexGrid_MouseDown Button, Shift, X, Y
        
    Else
    
        'Mouse down is somewhere else in the form
        'Hide txtEditGridCell
        Me.txtEditGridCell.Visible = False
        
    End If
        
End Sub

Private Sub Form_Resize()

    Me.Height = 6540
    Me.Width = 13320

End Sub

Private Function GetCurrentXMotorPosition() As Long
    GetCurrentXMotorPosition = frmDCMotors.ReadPosition(modMotor.MotorChanger)
End Function

Private Function GetCurrentXYCupPosition() As Long
    GetCurrentXYCupPosition = CLng(frmDCMotors.ChangerHole)
End Function

Private Sub GetCurrentXYPositions_AndSaveToXYCurrentMotorPositionControls()

    Dim x_pos As Long
    Dim y_pos As Long
    Dim cup As Long
    
    x_pos = GetCurrentXMotorPosition()
    y_pos = GetCurrentYMotorPosition()
    cup = GetCurrentXYCupPosition()
    
    UpdateXYCurrentMotorPositionControls x_pos, y_pos, cup

End Sub

Private Function GetCurrentYMotorPosition() As Long
    GetCurrentYMotorPosition = frmDCMotors.ReadPosition(modMotor.MotorChangerY)
End Function

Private Function GetCurTUnits() As String

    'Read the currently selected item in the Temperature Sensor
    'Units combo-box and return the matching units string
    If cmbTunits.ListIndex = 0 Then
    
        GetCurTUnits = "C"
        
    ElseIf cmbTunits.ListIndex = 1 Then
    
        GetCurTUnits = "F"
        
    ElseIf cmbTunits.ListIndex = 2 Then
    
        GetCurTUnits = "K"
        
    Else

        'Default units to "C"
        'If some strange value gets into the units-combo box,
        'set it back to Celcius
        GetCurTUnits = "C"
        cmbTunits.ListIndex = 0
        
    End If

End Function

Private Sub importSettings()
    Dim i As Integer
    txtZeroPos = ZeroPos
    txtMeasPos = MeasPos
    txtAFPos = AFPos
    txtIRMPos = IRMPos
    txtSCoilPos = SCoilPos
    txtFloorPos = FloorPos
    txtMinUpDownPos = MinUpDownPos
    txtSampleBottom = SampleBottom
    txtSampleTop = SampleTop
    txtSampHoleAlignOffset = SampleHoleAlignmentOffset
    txtLiftSpeedSlow = LiftSpeedSlow
    txtLiftSpeedNormal = LiftSpeedNormal
    txtLiftSpeedFast = LiftSpeedFast
    txtLiftAcceleration = LiftAcceleration
    txtChangerSpeed = ChangerSpeed
    txtTurnerSpeed = TurnerSpeed
    txtSCurveFactor = SCurveFactor
    txtTurningMotorFullRotation = TurningMotorFullRotation
    txtTurningMotor1rps = TurningMotor1rps
    txtUpDownMotor1cm = UpDownMotor1cm
    txtTrayOffsetAngle = TrayOffsetAngle
    txtPickupTorqueThrottle = PickupTorqueThrottle
    txtUpDownTorquefactor = UpDownTorqueFactor
    txtUpDownMaxTorque = UpDownMaxTorque
    txtZCal = ZCal
    txtXCal = XCal
    txtYCal = YCal
    txtRangeFact = RangeFact
    txtReadDelay = ReadDelay ' (March 2008 L Carporzen) Read delay
    txtSusceptibilityMomentFactorCGS = SusceptibilityMomentFactorCGS
    txtSusceptibilitySettings = SusceptibilitySettings
    
    'XY Table Settings
    If modConfig.UseXYTableAPS = True Then
        ckUseXYTableAPS.Value = vbChecked
        HomeHoleLocationBox.text = HoleSlotNum
        SlotMaxBox.text = SlotMax
    Else
        ckUseXYTableAPS.Value = vbUnchecked
    End If
   
    XYHolePositionsFlexGrid.Cols = 3
    XYHolePositionsFlexGrid.row = 0
    XYHolePositionsFlexGrid.Col = 0
    XYHolePositionsFlexGrid.text = "Cup"
    
    XYHolePositionsFlexGrid.Col = 1
    XYHolePositionsFlexGrid.text = "X Pos"
    XYHolePositionsFlexGrid.Col = 2
    XYHolePositionsFlexGrid.text = "Y Pos"
    
    XYHolePositionsFlexGrid.row = 1
    XYHolePositionsFlexGrid.Col = 0
    XYHolePositionsFlexGrid.text = "Home"
    XYHolePositionsFlexGrid.Col = 1
    XYHolePositionsFlexGrid.text = modConfig.XYTablePositions(0, 0)
    XYHolePositionsFlexGrid.Col = 2
    XYHolePositionsFlexGrid.text = modConfig.XYTablePositions(0, 1)
            
    For i = 2 To 101
        XYHolePositionsFlexGrid.row = i
        'Index
        XYHolePositionsFlexGrid.Col = 0
        XYHolePositionsFlexGrid.text = Str(i - 1)
        
        'X values
        XYHolePositionsFlexGrid.Col = 1
        XYHolePositionsFlexGrid.text = modConfig.XYTablePositions(i - 1, 0)
        
        'Y Values
        XYHolePositionsFlexGrid.Col = 2
        XYHolePositionsFlexGrid.text = modConfig.XYTablePositions(i - 1, 1)
        
    Next i
    
    'AF settings
    cmbAFDelay = Str$(AFDelay)
    cmbAFRampRate = Str$(AFRampRate)
    cmbAFAxialCoord = AfAxialCoord
    txtAFAxialMax = Trim(Str(AfAxialMax))
    txtAFAxialMin = Trim(Str(AfAxialMin))
    cmbAFTransCoord = AfTransCoord
    txtAFTransMax = Trim(Str(AfTransMax))
    txtAFTransMin = Trim(Str(AfTransMin))
    txtAFMinRampDownNumPeriods = Trim(Str(MinRampDown_NumPeriods))
    txtAFMaxRampDownNumPeriods = Trim(Str(MaxRampDown_NumPeriods))
    txtAFMinRampUpTime = Trim(Str(MinRampUpTime_ms))
    txtAFAxialRampUpVoltsPerSec = Trim(Str(modConfig.AxialRampUpVoltsPerSec))
    txtAFTransverseRampUpVoltsPerSec = Trim(Str(modConfig.TransRampUpVoltsPerSec))
    txtAFMaxRampUpTime = Trim(Str(MaxRampUpTime_ms))
    txtPeakPeriods = Trim(Str(HoldAtPeakField_NumPeriods))
    
    txtAFRampDownPeriodsPerVolt = Trim(Str(RampDownNumPeriodsPerVolt))
    lblAFAxialRampMax.Caption = Trim(Str(modConfig.AfAxialRampMax))
    lblAFTransverseRampMax.Caption = Trim(Str(modConfig.AfTransRampMax))
           
    isADwinRampSettings_dirty = False
           
    'Load the AF System combo box
    If AFSystem = "2G" Then
    
        isUserChange = False
        cmbAFSystem.ListIndex = 0
        
    ElseIf AFSystem = "ADWIN" Then
    
        isUserChange = False
        cmbAFSystem.ListIndex = 1
        
    End If
    
    isUserChange = True
        
    'Set the AF Units combo box
    If modConfig.AFUnits = "G" Then
    
        cmbAFUnits.ListIndex = 0
        
    ElseIf modConfig.AFUnits = "T" Then
    
        cmbAFUnits.ListIndex = 1
        
    End If
        
    'Load the temperature sensor settings
    
    'Store the Current value of the Temperature Sensor units into
    'the Last TUnits field
    LastTUnits = modConfig.Tunits
    
    'Clear and re-setup the items in the Temperature sensor units combo-box
    cmbTunits.Clear
    cmbTunits.AddItem "Celsius, °C", 0
    cmbTunits.AddItem "Farenheit, °F", 1
    cmbTunits.AddItem "Kelvin, K", 2
    
    'Load the Tunits into the units combo-box
    Select Case modConfig.Tunits
    
        Case "C"
        
            cmbTunits.ListIndex = 0
            
        Case "F"
        
            cmbTunits.ListIndex = 1
            
        Case "K"
        
            cmbTunits.ListIndex = 2
            
        Case Else
        
            'If weird value in units field, reset to Celcius
            cmbTunits.ListIndex = 0
            
    End Select
    
    'Change all of the units labels on the form
    'First check to see if modConfig.Tunits is inZ Kelvin
    If CurTunits = "K" Then
        
        lblTHotUnits.Caption = modConfig.Tunits
        lblTMaxUnits.Caption = modConfig.Tunits
        lblToffsetUnits.Caption = modConfig.Tunits
        lblTslopeUnits.Caption = modConfig.Tunits & " / V"
                                
    Else
    
        lblTHotUnits.Caption = "°" & modConfig.Tunits
        lblTMaxUnits.Caption = "°" & modConfig.Tunits
        lblToffsetUnits.Caption = "°" & modConfig.Tunits
        lblTslopeUnits.Caption = "°" & modConfig.Tunits & " / V"
    
    End If
    
    'Now load the "Hot" temperature alarm value
    txtThot = modConfig.Thot
    
    'Now load the Maximum allowed AF coil temperature before the AF ramp
    'is halted
    txtTmax = modConfig.Tmax
    
    'Now load the Temperature offset for converted the thermocouple voltage
    txtToffset = modConfig.Toffset
    
    'Now load the Temperature / Voltage slope for converted the thermocouple voltage
    txtTslope = modConfig.TSlope
    
    Me.txtMaxIrmVoltageOut_BoostPercentage.text = GetIRMBoostPercentage_FromMultiplier(modConfig.AscSetVoltageMaxBoostMultiplier)
    Me.txtMinIrmVoltageOut_BoostPercentage.text = GetIRMBoostPercentage_FromMultiplier(modConfig.AscSetVoltageMinBoostMultiplier)
   
    'Load the IRM system
    If modConfig.IRMSystem = "Matsusada" Then
    
        cmbIRMSystem.ListIndex = 1
        DisableAscOnlyIRMControls
        
    ElseIf modConfig.IRMSystem = "Old" Then
    
        cmbIRMSystem.ListIndex = 0
        DisableAscOnlyIRMControls
        
    ElseIf modConfig.IRMSystem = "ASC" Then
    
        cmbIRMSystem.ListIndex = 2
        EnableAscOnlyIRMControls
        
    Else
    
        DisableAscOnlyIRMControls
        
    End If
   
    'Read the AF/IRM Coil Digital Channels in
    cmbIRMAxis = modConfig.IRMAxis
    cmbIRMBackfieldAxis = modConfig.IRMBackfieldAxis
    
    'Set the IRM settings
    Me.optTrimOnTrue.Value = modConfig.TrimOnTrue
    Me.optTrimOnFalse.Value = Not modConfig.TrimOnTrue
    txtPulseAxialMax = PulseAxialMax
    txtPulseAxialMin = PulseAxialMin
    txtAxialMaxCapVoltage = IRMAxialVoltMax
    If AxialTransMaxCapVoltsSame = True Then
    
        txtTransMaxCapVoltage = txtAxialMaxCapVoltage
        
    Else
    
        txtTransMaxCapVoltage = IRMTransVoltMax
        
    End If
    
    txtPulseTransMax = PulseTransMax
    txtPulseTransMin = PulseTransMin
    txtPulseVoltMax = PulseVoltMax
    txtPulseMCCVoltConversion = PulseMCCVoltConversion
    txtPulseReturnMCCVoltConversion = PulseReturnMCCVoltConversion
    txtARMMax = ARMMax
    txtARMVoltGauss = ARMVoltGauss
    txtARMVoltMax = ARMVoltMax
      
'------------------------------------------------------------------------------------------------'
'------------------------------------------------------------------------------------------------'
'
'   Major Code Modification
'   Created: April 2010
'    Author: Isaac Hilburn
'
'   As described in modConfig, the ARM, IRM, Vacuum, and Analog temperature sensor
'   DAQ Board and Channel settings are now setup using the Board & Channel class object
'   libraries.
'
'   These channels & boards are stored in the .INI file and read in by special functions
'   in modConfig.
'
'   The User interface in frmSettings has been changed to display both the Board and the
'   channel on that board that has been selected/assigned for a particular use:
'
'   Analog Output Channels:
'
'       ARM Voltage Out
'       IRM Voltage Out
'       AF ADWIN Ramp Out
'
'   Analog Input Channels:
'
'       Axial Coil Temperature Sensor (AnalogT1)
'       Trans. Coil Temperature Sensor (AnalogT2)
'       IRM Capacitor Voltage In
'       IRM Monitor (different - used for real time record of IRM pulse through the
'                    axial coil)
'       AF ADWIN Monitor
'       Alternate AF Monitor (MCC board)
'
'   Digital Output Channels:
'
'       ARM Set
'       IRM Fire
'       IRM Trim
'       Vacuum Motor Toggle On/Off
'       Vacuum valve Toggle (Vacuum Toggle A)
'       Axial AF coil Relay
'       Transverse AF coil relay
'       IRM relay
'
'
'   Digital Input Channels:
'
'       IRM Ready
'
'   The channel class objects keep track of whether a particular board / channel combination
'   has already been assigned to something and will notify the user if they try to
'   assign more than one function to a particular channel on a particular board
'------------------------------------------------------------------------------------------------'

    'Analog Ouput for the AF ADWIN Ramp voltage output to the Crest Audio amplifier
    SetBoardAndChanComboBoxes cmbAFRampBoard, _
                              cmbAFRampChan, _
                              modConfig.AFRampChan, _
                              "AF ADWIN Ramp"

    'Analog Output for the ARM set voltage output
    SetBoardAndChanComboBoxes cmbARMVoltageOutBoard, _
                              cmbARMVoltageOutChan, _
                              modConfig.ARMVoltageOut, _
                              "ARM Bias Voltage"

    'Analog Output for the IRM set voltage output
    SetBoardAndChanComboBoxes cmbIRMVoltageOutBoard, _
                              cmbIRMVoltageOutChan, _
                              modConfig.IRMVoltageOut, _
                              "IRM Set Voltage"

    'Analog Input for the AF ADWIN Monitor voltage input from the AF LC circuit
    SetBoardAndChanComboBoxes cmbAFMonitorBoard, _
                              cmbAFMonitorChan, _
                              modConfig.AFMonitorChan, _
                              "AF ADWIN Monitor"
                         
    'Analog Input for the Alternate AF Monitor voltage input from the AF LC circuit
    '(Using the MCC DAQ Board - can be used with 2G or ADWIN ramp)
    SetBoardAndChanComboBoxes cmbAltAFMonitorBoard, _
                              cmbAltAFMonitorChan, _
                              modConfig.AltAFMonitorChan, _
                              "Alternate AF Monitor"

    'Analog Input for the 1st AF coil temperature sensor
    SetBoardAndChanComboBoxes cmbAnalogT1Board, _
                              cmbAnalogT1Chan, _
                              modConfig.AnalogT1, _
                              "AF Coil Temp. Sensor #1"
                              
    'Analog Input for the 2nd AF coil temperature sensor
    SetBoardAndChanComboBoxes cmbAnalogT2Board, _
                              cmbAnalogT2Chan, _
                              modConfig.AnalogT2, _
                              "AF Coil Temp. Sensor #2"
                 
    'Analog Input for the IRM capacitor voltage
    SetBoardAndChanComboBoxes cmbIRMCapacitorVoltageInBoard, _
                              cmbIRMCapacitorVoltageInChan, _
                              modConfig.IRMCapacitorVoltageIn, _
                              "IRM Capacitor Return Voltage"
                              
    'Analog Input for the IRM Monitor
    SetBoardAndChanComboBoxes cmbIRMMonitorBoard, _
                              cmbIRMMonitorChan, _
                              modConfig.IRMMonitor, _
                              "IRM Monitor Voltage"
                              
    'Digital Output for the ARM Set
    SetBoardAndChanComboBoxes cmbARMSetBoard, _
                              cmbARMSetChan, _
                              modConfig.ARMSet, _
                              "ARM Set Dig. Output"

    'Digital Output for the IRM Fire
    SetBoardAndChanComboBoxes cmbIRMFireBoard, _
                              cmbIRMFireChan, _
                              modConfig.IRMFire, _
                              "IRM Fire Dig. Output"
                              
    'Digital Output for the IRM Trim
    SetBoardAndChanComboBoxes cmbIRMTrimBoard, _
                              cmbIRMTrimChan, _
                              modConfig.IRMTrim, _
                              "IRM Trim Dig. Output"
                              
    'Digital Output for the Vacuum Motor On/Off toggle
    SetBoardAndChanComboBoxes cmbMotorToggleBoard, _
                              cmbMotorToggleChan, _
                              modConfig.MotorToggle, _
                              "Vacuum Motor On/Off Dig. Output"
                              
    'Digital Output for the Vacuum Motor On/Off toggle
    SetBoardAndChanComboBoxes cmbVacuumToggleABoard, _
                              cmbVacuumToggleAChan, _
                              modConfig.VacuumToggleA, _
                              "Vacuum Valve Toggle Dig. Output"
                              
    'Digital Output for the Degausser Air Cooler toggle
    SetBoardAndChanComboBoxes cmbDegausserCoolerBoard, _
                              cmbDegausserCoolerChan, _
                              modConfig.DegausserToggle, _
                              "Degausser Toggle On/Off Dig. Output"
                                 
    'Digital Output for the AF Axial Relay
    SetBoardAndChanComboBoxes cmbAxialRelayBoard, _
                              cmbAxialRelayChan, _
                              modConfig.AFAxialRelay, _
                              "AF Axial Relay"
                              
    'Digital Output for the AF Trans Relay
    SetBoardAndChanComboBoxes cmbTransRelayBoard, _
                              cmbTransRelayChan, _
                              modConfig.AFTransRelay, _
                              "AF Trans. Relay"

    'Digital Output for the IRM Pulse Relay
    SetBoardAndChanComboBoxes cmbIRMRelayBoard, _
                              cmbIRMRelayChan, _
                              modConfig.IRMRelay, _
                              "IRM Relay"
                              
    'Digital Input for IRM Ready status
    SetBoardAndChanComboBoxes cmbIRMPowerAmpVoltageInBoard, _
                              cmbIRMPowerAmpVoltageInChan, _
                              modConfig.IRMPowerAmpVoltageIn, _
                              "IRM Ready Status Dig. Input"
                              
'------------------------------------------------------------------------------------------------'
'
'   Old Legacy Code:
'------------------------------------------------------------------------------------------------'
'    ' (March 2008 L Carporzen) Put in Settings the IRM/ARM channels (MIT acquisition board does not work on IRMTrim = 3
'
'    ' Analog channel output
'    txtARMVoltageOut = ARMVoltageOut
'    txtIRMVoltageOut = IRMVoltageOut
'
'    ' Analog input
'    txtIRMCapacitorVoltageIn = IRMCapacitorVoltageIn
'
'    ' DIO line assignments
'    txtARMSet = ARMSet
'    txtIRMFire = IRMFire
'    txtIRMTrim = IRMTrim
'    txtIRMPowerAmpVoltageIn = IRMPowerAmpVoltageIn
'    txtMotorToggle = MotorToggle
'    txtVacuumToggleA = VacuumToggleA
''    txtVacuumToggleB = VacuumToggleB
'------------------------------------------------------------------------------------------------'
'------------------------------------------------------------------------------------------------'
        
    txtCmdHometoTop = CmdHomeToTop
    txtCmdSamplePickup = CmdSamplePickup
    txtMotorIDTurning = MotorIDTurning
    txtMotorIDChanger = MotorIDChanger
    txtMotorIDChangerY = MotorIDChangerY
    txtMotorIDUpDown = MotorIDUpDown
    txtDropoffVacuumDelay = DropoffVacuumDelay
    
    If DoVacuumReset = True Then
    
        chkDoVacuumReset.Value = Checked
        
    Else
    
        chkDoVacuumReset.Value = Unchecked
        
    End If
    
    If DoDegausserCooling = True Then
    
        chkDegausserAirCooler.Value = Checked
        
    Else
    
        chkDegausserAirCooler.Value = Unchecked
        
    End If
    
    If EnableAxialIRM = True Then
    
        chkEnableIRMAxial.Value = Checked
        
    Else
    
        chkEnableIRMAxial.Value = Unchecked
        
    End If
    
    If EnableTransIRM = True Then
    
        chkEnableIRMTrans.Value = Checked
        
    Else
    
        chkEnableIRMTrans.Value = Unchecked
        
    End If
    
    If EnableIRMBackfield = True Then
        
        chkEnableIRMBackfield.Value = Checked
        
    Else
    
        chkEnableIRMBackfield.Value = Unchecked
        
    End If
    
    
    If EnableIRMMonitor = True Then
        
        chkEnableIRMMonitor.Value = Checked
        
    Else
        
        chkEnableIRMMonitor.Value = Unchecked
    
    End If
    
    If EnableARM = True Then
    
        chkEnableARM.Value = Checked
    
    Else
    
        chkEnableARM.Value = Unchecked
        
    End If
    
    If EnableAF = True Then
    
        chkEnableAF.Value = Checked
    
    Else
    
        chkEnableAF.Value = Unchecked
    
    End If
    
    If EnableSusceptibility = True Then
    
        chkEnableSusceptibility.Value = Checked
    
    Else
    
        chkEnableSusceptibility.Value = Unchecked
    
    End If
    
    If EnableAFAnalysis = True Then
    
        Me.chkEnableAFAnalysis.Value = Checked
        
    Else
    
        Me.chkEnableAFAnalysis.Value = Unchecked
    
    End If
    
    If EnableVacuum = True Then
    
        Me.chkEnableVacuum.Value = Checked
        
    Else
    
        Me.chkEnableVacuum.Value = Unchecked
        
    End If
    
    If EnableDegausserCooler = True Then
    
        Me.chkEnableDegausserCooler.Value = Checked
        
    Else
    
        Me.chkEnableDegausserCooler.Value = Unchecked
        
    End If
    
    If EnableT1 = True Then
    
        Me.checkAnalogT1.Value = Checked
        
    Else
    
        Me.checkAnalogT1.Value = Unchecked
    
    End If
        
    If EnableT2 = True Then
    
        Me.checkAnalogT2.Value = Checked
        
    Else
    
        Me.checkAnalogT2.Value = Unchecked
    
    End If
        
End Sub


Private Function GetIRMBoostPercentage_FromMultiplier(ByVal boost_multiplier As Double) As Double

    If boost_multiplier > 0 Then

        GetIRMBoostPercentage_FromMultiplier = (boost_multiplier - 1) * 100
        
    Else
        
        GetIRMBoostPercentage_FromMultiplier = 0
    
    End If
    
        

End Function

Private Function GetIRMBoostMultiplier_FromPercentage(ByVal boost_percentage As Double) As Double

    GetIRMBoostMultiplier_FromPercentage = (boost_percentage / 100) + 1
    
    If GetIRMBoostMultiplier_FromPercentage < 0 Then GetIRMBoostMultiplier_FromPercentage = 0

End Function


Private Sub DisableAscOnlyIRMControls()

    Me.frameAscIRMOutputVoltageBoostFactors.Enabled = False
    Me.txtMaxIrmVoltageOut_BoostPercentage.Enabled = False
    Me.txtMinIrmVoltageOut_BoostPercentage.Enabled = False

End Sub

Private Sub EnableAscOnlyIRMControls()

    Me.frameAscIRMOutputVoltageBoostFactors.Enabled = True
    Me.txtMaxIrmVoltageOut_BoostPercentage.Enabled = True
    Me.txtMinIrmVoltageOut_BoostPercentage.Enabled = True

End Sub



Private Sub LoadBoardChanComboBoxes()

    Dim i As Long
    Dim N As Long
    
    'Get the number of boards in System Boards Collection
    N = SystemBoards.Count

    'Need to load the board combo boxes from the System DAQ Boards collection now
    'For each Board in the System Boards collection, add on new item into the combo-boxes
    For i = 1 To N
    
        With SystemBoards(i)
    
            Me.cmbAnalogT1Board.AddItem .BoardName
            Me.cmbAnalogT2Board.AddItem .BoardName
            Me.cmbARMSetBoard.AddItem .BoardName
            Me.cmbARMVoltageOutBoard.AddItem .BoardName
            Me.cmbAxialRelayBoard.AddItem .BoardName
            Me.cmbIRMCapacitorVoltageInBoard.AddItem .BoardName
            Me.cmbIRMFireBoard.AddItem .BoardName
            Me.cmbIRMRelayBoard.AddItem .BoardName
            Me.cmbIRMPowerAmpVoltageInBoard.AddItem .BoardName
            Me.cmbIRMTrimBoard.AddItem .BoardName
            Me.cmbIRMVoltageOutBoard.AddItem .BoardName
            Me.cmbMotorToggleBoard.AddItem .BoardName
            Me.cmbTransRelayBoard.AddItem .BoardName
            Me.cmbVacuumToggleABoard.AddItem .BoardName
            Me.cmbDegausserCoolerBoard.AddItem .BoardName
            Me.cmbIRMMonitorBoard.AddItem .BoardName
            Me.cmbAFMonitorBoard.AddItem .BoardName
            Me.cmbAltAFMonitorBoard.AddItem .BoardName
            Me.cmbAFRampBoard.AddItem .BoardName
            
        End With
        
    Next
        
    'Set the active list-index for all the board combo boxes
    Me.cmbAnalogT1Board.ListIndex = 0
    Me.cmbAnalogT2Board.ListIndex = 0
    Me.cmbARMSetBoard.ListIndex = 0
    Me.cmbARMVoltageOutBoard.ListIndex = 0
    Me.cmbAxialRelayBoard.ListIndex = 0
    Me.cmbIRMCapacitorVoltageInBoard.ListIndex = 0
    Me.cmbIRMFireBoard.ListIndex = 0
    Me.cmbIRMRelayBoard.ListIndex = 0
    Me.cmbIRMPowerAmpVoltageInBoard.ListIndex = 0
    Me.cmbIRMTrimBoard.ListIndex = 0
    Me.cmbIRMVoltageOutBoard.ListIndex = 0
    Me.cmbMotorToggleBoard.ListIndex = 0
    Me.cmbTransRelayBoard.ListIndex = 0
    Me.cmbVacuumToggleABoard.ListIndex = 0
    Me.cmbDegausserCoolerBoard.ListIndex = 0
    Me.cmbIRMMonitorBoard.ListIndex = 0
    Me.cmbAFMonitorBoard.ListIndex = 0
    Me.cmbAltAFMonitorBoard.ListIndex = 0
    Me.cmbAFRampBoard.ListIndex = 0
    
    'Make Sure all the board combo boxes are unlocked
    Me.cmbAnalogT1Board.Locked = False
    Me.cmbAnalogT2Board.Locked = False
    Me.cmbARMSetBoard.Locked = False
    Me.cmbARMVoltageOutBoard.Locked = False
    Me.cmbAxialRelayBoard.Locked = False
    Me.cmbIRMCapacitorVoltageInBoard.Locked = False
    Me.cmbIRMFireBoard.Locked = False
    Me.cmbIRMRelayBoard.Locked = False
    Me.cmbIRMPowerAmpVoltageInBoard.Locked = False
    Me.cmbIRMTrimBoard.Locked = False
    Me.cmbIRMVoltageOutBoard.Locked = False
    Me.cmbMotorToggleBoard.Locked = False
    Me.cmbTransRelayBoard.Locked = False
    Me.cmbVacuumToggleABoard.Locked = False
    Me.cmbDegausserCoolerBoard.Locked = False
    Me.cmbIRMMonitorBoard.Locked = False
    Me.cmbAFMonitorBoard.Locked = False
    Me.cmbAltAFMonitorBoard.Locked = False
    Me.cmbAFRampBoard.Locked = False
    
    'Make Sure all the channel combo boxes are also unlocked
    Me.cmbAnalogT1Chan.Locked = False
    Me.cmbAnalogT2Chan.Locked = False
    Me.cmbARMSetChan.Locked = False
    Me.cmbARMVoltageOutChan.Locked = False
    Me.cmbAxialRelayChan.Locked = False
    Me.cmbIRMCapacitorVoltageInChan.Locked = False
    Me.cmbIRMFireChan.Locked = False
    Me.cmbIRMRelayChan.Locked = False
    Me.cmbIRMPowerAmpVoltageInChan.Locked = False
    Me.cmbIRMTrimChan.Locked = False
    Me.cmbIRMVoltageOutChan.Locked = False
    Me.cmbMotorToggleChan.Locked = False
    Me.cmbTransRelayChan.Locked = False
    Me.cmbVacuumToggleAChan.Locked = False
    Me.cmbDegausserCoolerChan.Locked = False
    Me.cmbIRMMonitorChan.Locked = False
    Me.cmbAFMonitorChan.Locked = False
    Me.cmbAltAFMonitorChan.Locked = False
    Me.cmbAFRampChan.Locked = False
    
    'Now call the board combo-box change event for every board
    'These functions will load the possible channels into the
    'matching channel combo-boxes
    PriorListIndex = -1
    cmbAnalogT1Board_Click
    PriorListIndex = -1
    cmbAnalogT2Board_Click
    PriorListIndex = -1
    cmbARMSetBoard_Click
    PriorListIndex = -1
    cmbARMVoltageOutBoard_Click
    PriorListIndex = -1
    cmbAxialRelayBoard_Click
    PriorListIndex = -1
    cmbIRMCapacitorVoltageInBoard_Click
    PriorListIndex = -1
    cmbIRMFireBoard_Click
    PriorListIndex = -1
    cmbIRMRelayBoard_Click
    PriorListIndex = -1
    cmbIRMPowerAmpVoltageInBoard_Click
    PriorListIndex = -1
    cmbIRMTrimBoard_Click
    PriorListIndex = -1
    cmbIRMMonitorBoard_Click
    PriorListIndex = -1
    cmbIRMVoltageOutBoard_Click
    PriorListIndex = -1
    cmbMotorToggleBoard_Click
    PriorListIndex = -1
    cmbTransRelayBoard_Click
    PriorListIndex = -1
    cmbVacuumToggleABoard_Click
    PriorListIndex = -1
    cmbDegausserCoolerBoard_Click
    PriorListIndex = -1
    cmbAFMonitorBoard_Click
    PriorListIndex = -1
    cmbAltAFMonitorBoard_Click
    PriorListIndex = -1
    cmbAFRampBoard_Click
        
End Sub

'Subroutine LoadChannelComboBox
'
'Takes in pointer to a Channel selector combo-box and a channel object and
'uses the board-name and channel type stored in the channel object to
'load a list of possible channels from which to select
'
'   Inputs:
'
'   cmbChan     -   Reference to the channel combo-box control for a DAQ Channel
'
'   ChanObj     -   Reference to a channel object containing the desired comm settings
'                   to use to set the channel combo-box above
'
'   Output:
'
'   cmbChan     -   Function modifies the channel combo-box control so that
'                   it has a list of all the needed channels and sets the active
'                   ListIndex to the entry in the combo-box matching the channel name
'                   property stored in the inputed ChanObj
'
Sub LoadChannelComboBox(ByRef cmbChan As ComboBox, _
                        ByRef ChanObj As Channel, _
                        Optional ByRef cmbBoard As ComboBox = Nothing)
                        
    Dim i As Long
    Dim N As Long
    Dim TempBoard As Board
    Dim TempChannels As Channels
    Dim ChanFound As Boolean
    
    'Default TempBoard = Nothing
    Set TempBoard = Nothing
    
    'Clear the channel combo box
    cmbChan.Clear
    
    'Turn on error handling
    On Error Resume Next
    
        'Check to see if the user entered an alternate board-name to use
        'to get the desired board object
        If Not cmbBoard Is Nothing Then
    
            Set TempBoard = SystemBoards(cmbBoard.List(cmbBoard.ListIndex))
    
        Else
    
            Set TempBoard = SystemBoards(ChanObj.BoardName)
            
        End If
        
        'Error check
        If Err.number <> 0 Or TempBoard Is Nothing Then
        
            'Wasn't able to find a matching board
            'Add the channel "ERROR" as an item in the combo-box
            cmbChan.Clear
            cmbChan.AddItem "ERROR", 0
            
            'Also add an additional entry for the user to type in a channel name
            cmbChan.AddItem "Type in name...", 1
                
            'We're done for now
            Exit Sub
            
        End If
        
    'turn off error handling
    On Error GoTo 0
    
    'We have a board object to work with, need to select the right
    'Channels collection to use from that board
    
    'Now populate the cmbChan with the correct channel options
    Select Case ChanObj.ChanType
    
        Case "AI"
    
            Set TempChannels = TempBoard.AInChannels
            
        Case "AO"
        
            Set TempChannels = TempBoard.AOutChannels
            
        Case "DI"
        
            Set TempChannels = TempBoard.DInChannels
        
        Case "DO"
        
            Set TempChannels = TempBoard.DOutChannels
            
    End Select
    
    'Get the number of channels in the TempChannels collection
    'Turn on error handling
    On Error Resume Next
    
        N = TempChannels.Count
        
        'Check for error
        If Err.number <> 0 Then
        
            'This channel object is blank,
            'Do the same as was done for the no-matching board error above
            
            'Add the channel "ERROR" as an item in the combo-box
            cmbChan.Clear
            cmbChan.AddItem "ERROR", 0
            
            'Also add an additional entry for the user to type in a channel name
            cmbChan.AddItem "Type in name...", 1
                
            'We're done for now
            Exit Sub
            
        End If
        
    'Turn off error handling
    On Error GoTo 0
    
    'Check to see if N = 0
    If N <= 0 Then
    
        'No channels present in this channels collection
        'Do the same as was done for the no-matching board error above
        
        'Add the channel "ERROR" as an item in the combo-box
        cmbChan.Clear
        cmbChan.AddItem "ERROR", 0
        
        'Also add an additional entry for the user to type in a channel name
        cmbChan.AddItem "Type in name...", 1
            
        'We're done for now
        Exit Sub
        
    End If

    'Default Channel Found bool flag to false
    ChanFound = False

    'Now Load the channels in the Channel combo-box
    For i = 1 To N
    
        cmbChan.AddItem TempChannels(i).ChanName, i - 1
    
        'If new channel's name matches that of the inputed channel object
        'then set the list-index to that name
        If cmbChan.List(i - 1) = ChanObj.ChanName Then
        
            cmbChan.ListIndex = i - 1
            ChanFound = True
            
        End If
    
    Next i
    
    'If channel wasn't found, set the list-index to 0
    If ChanFound = False Then
    
        cmbChan.ListIndex = 0
        
    End If
    
    'Deallocate TempBoard and TempChannels
    Set TempBoard = Nothing
    Set TempChannels = Nothing
                        
End Sub

Public Function MaxL(ByVal val1 As Long, ByVal val2 As Long) As Long

    If val2 > val1 Then
    
        MaxL = val2
    
    Else
    
        MaxL = val1
        
    End If

End Function

Private Sub MoveHomeButton_Click()
Dim xPos As Long
Dim yPos As Long

    UpdateXYCurrentMotorPositionControls_Moving

    frmDCMotors.HomeToCenter xPos, yPos, pauseOveride:=True
    modConfig.XYTablePositions(0, 0) = xPos
    modConfig.XYTablePositions(0, 1) = yPos
    modConfig.XYTablePositions(46, 0) = xPos
    modConfig.XYTablePositions(46, 1) = yPos
    
    GetCurrentXYPositions_AndSaveToXYCurrentMotorPositionControls
End Sub

Private Sub MoveToCupButton_Click()
Dim xPos As Long
Dim yPos As Long

XYHolePositionsFlexGrid.Col = 1
xPos = CLng(XYHolePositionsFlexGrid.text)
XYHolePositionsFlexGrid.Col = 2
yPos = CLng(XYHolePositionsFlexGrid.text)

UpdateXYCurrentMotorPositionControls_Moving

frmDCMotors.MotorStop MotorChanger
frmDCMotors.MotorStop MotorChangerY
frmDCMotors.MoveMotorAbsoluteXY MotorChanger, xPos, val(LTrim$(MoveSpeed.text)), False, False
frmDCMotors.MoveMotorAbsoluteXY MotorChangerY, yPos, val(LTrim$(MoveSpeed.text)), False, False

GetCurrentXYPositions_AndSaveToXYCurrentMotorPositionControls

End Sub

Private Sub optTrimOnFalse_Click()

    If Me.optTrimOnFalse.Value = False Then Exit Sub

    'Set Trim On True radio button to false
    Me.optTrimOnTrue.Value = False
    
End Sub

Private Sub optTrimOnTrue_Click()

    If Me.optTrimOnTrue.Value = False Then Exit Sub

    'Set Trim On False radio button to false
    Me.optTrimOnFalse.Value = False
    
End Sub

Private Sub ReadUDIO_Click()
IOResult.text = frmDCMotors.CheckInternalStatus(MotorUpDown, val(ReadIOLine.text) + 3)

GetCurrentXYPositions_AndSaveToXYCurrentMotorPositionControls
End Sub

Private Sub ReadXIO_Click()
IOResult.text = frmDCMotors.CheckInternalStatus(MotorChanger, val(ReadIOLine.text) + 3)

GetCurrentXYPositions_AndSaveToXYCurrentMotorPositionControls
End Sub

Private Sub ReadYIO_Click()
IOResult.text = frmDCMotors.CheckInternalStatus(MotorChangerY, val(ReadIOLine.text) + 3)

GetCurrentXYPositions_AndSaveToXYCurrentMotorPositionControls
End Sub

Public Function SaveDAQSetting _
    (ByRef cmbBoard As ComboBox, _
     ByRef cmbChan As ComboBox, _
     ByVal ChanType As String, _
     Optional ByVal ChanDesc As String = "") As Channel
                               
    Dim TempChan As Channel
    Dim TempChannels As Channels
    Dim TempBoard As Board
    
    Dim ChanFullType As String
    
    Set TempBoard = Nothing
    Set TempChan = Nothing
    Set TempChannels = Nothing
    
    'Turn on Error handling
    On Error Resume Next
        
        'Get the System boards object
        Set TempBoard = SystemBoards(cmbBoard.List(cmbBoard.ListIndex))

        'Check for error
        If Err.number <> 0 Then
        
            'Bad Board name or Board is missing in the global systemBoards collection
            '(DOH!)
            Err.Raise Err.number, _
                      "frmSettings.SaveDAQSetting", _
                      "For the " & ChanDesc & " channel." & _
                      vbNewLine & _
                      "Bad Board Name used to find a Board object in the System " & _
                      "Boards collection, or a Board object is missing from the " & _
                      "System Boards collection."
              
            'Return a null channel
            Set SaveDAQSetting = Nothing
              
            Exit Function
                      
        End If
        
    On Error GoTo 0

    'Error check
    If TempBoard Is Nothing Then
    
        'No matching board found in the system boards collection
        'Bad Board name or Board is missing in the global systemBoards collection
        '(DOH!)
        Err.Raise -616, _
                  "frmSettings.SaveDAQSetting", _
                  "For the " & ChanDesc & " channel." & _
                  vbNewLine & _
                  "Bad Board Name used to find a Board object in the System " & _
                  "Boards collection, or a Board object is missing from the " & _
                  "System Boards collection."
                  
        'Return a null channel
        Set SaveDAQSetting = Nothing
                  
        Exit Function
                  
    End If
        
    'Board is good / legit
    'Now need to search the right Channels collection
    Select Case ChanType
    
        Case "AO"
        
            Set TempChannels = TempBoard.AOutChannels
            ChanFullType = "Analog Output"
            
        Case "AI"
        
            Set TempChannels = TempBoard.AInChannels
            ChanFullType = "Analog Input"
            
        Case "DO"
        
            Set TempChannels = TempBoard.DOutChannels
            ChanFullType = "Digital Output"
            
        Case "DI"
        
            Set TempChannels = TempBoard.DInChannels
            ChanFullType = "Digital Input"
         
    End Select
    
    'Turn on Error Handling
    On Error Resume Next
    
        Set TempChan = TempChannels(cmbChan.List(cmbChan.ListIndex))
        
        'Error Check
        If Err.number <> 0 Then
        
            'Bad channel name, or channel is missing from the channels
            'collections of this board
            Err.Raise Err.number, _
                      "frmSettings.SaveDAQSetting", _
                      "For " & ChanDesc & " channel," & vbNewLine & _
                      "Bad Channel name, or channel is missing from the " & _
                      ChanFullType & " channels collection for Board: " & _
                      TempBoard.BoardName
                      
            'Return a null value
            Set SaveDAQSetting = Nothing
            
            Exit Function
            
        End If
        
    On Error GoTo 0
    
    'Check for a Nothing value (error check #2)
    If TempChan Is Nothing Then
                     
        'Bad channel name, or channel is missing from the channels
        'collections of this board
        Err.Raise -616, _
                  "frmSettings.SaveDAQSetting", _
                  "For " & ChanDesc & " channel," & vbNewLine & _
                  "Bad Channel name, or channel is missing from the " & _
                  ChanFullType & " channels collection for Board: " & _
                  TempBoard.BoardName
                  
        'Return a null value
        Set SaveDAQSetting = Nothing
        
        Exit Function
        
    End If
    
    'Return TempChan as the desired channel object
    Set SaveDAQSetting = TempChan
    
    'Deallocate all the temp objects
    Set TempBoard = Nothing
    Set TempChannels = Nothing
    Set TempChan = Nothing
    
End Function

Public Sub selectTab(ByVal tabtoselect As Integer)
    Dim i As Integer
    'show and enable the selected tab's controls
    'and hide and disable all others
    For i = 0 To tbsOptions.Tabs.Count - 1
        If i = tabtoselect - 1 Then
            frameOptions(i).Enabled = True
            frameOptions(i).Visible = True
            frameOptions(i).ZOrder 0
        Else
            frameOptions(i).Visible = False
            frameOptions(i).Enabled = False
        End If
    Next
End Sub

Public Sub SetAFSystem _
    (ByVal NewAFSystem As String, _
     Optional ByVal UserDriven As Boolean = False)

    'Deactivate the AF system combo-box Change event
    'if so specified by the calling function
    isUserChange = UserDriven
    
    If isUserChange = False Then
    
        'Need to change the value in the combo-box
        Me.cmbAFSystem = NewAFSystem
        
    End If

    'Show/Hide the combo boxes for the 2G variables
    cmbAFRampRate.Visible = (cmbAFSystem = "2G")
    cmbAFDelay.Visible = (cmbAFSystem = "2G")
    cmbAFAxialCoord.Visible = (cmbAFSystem = "2G")
    cmbAFTransCoord.Visible = (cmbAFSystem = "2G")
    Me.cmbIRMAxis.Visible = (cmbAFSystem = "2G")
    Me.cmbIRMBackfieldAxis.Visible = (cmbAFSystem = "2G")
    Me.lblIRMBackfieldAxis.Visible = (cmbAFSystem = "2G")
    Me.lblIRMAxis.Visible = (cmbAFSystem = "2G")
        
    'Show/Hide the labels for those combo boxes
    Me.lblAFRampRate.Visible = (cmbAFSystem = "2G")
    Me.lblAFDelay.Visible = (cmbAFSystem = "2G")
    Me.lblAxialCoord.Visible = (cmbAFSystem = "2G")
    Me.lblTransCoord.Visible = (cmbAFSystem = "2G")
        
    'Enable the AF Calibration button
    Me.cmdCalAFCoils.Enabled = True
    
    'Enable/Disable the AF Tune button
    Me.cmdTuneAF.Enabled = (cmbAFSystem = "ADWIN")
    
    'Show/Hide the combo-boxes for changing the channel & board
    'for the AF / IRM relays
    Me.cmbAxialRelayBoard.Enabled = (cmbAFSystem = "ADWIN")
    Me.cmbAxialRelayChan.Enabled = (cmbAFSystem = "ADWIN")
    Me.cmbTransRelayBoard.Enabled = (cmbAFSystem = "ADWIN")
    Me.cmbTransRelayChan.Enabled = (cmbAFSystem = "ADWIN")
    Me.cmbIRMRelayBoard.Enabled = (cmbAFSystem = "ADWIN")
    Me.cmbIRMRelayChan.Enabled = (cmbAFSystem = "ADWIN")
    Me.cmbAFMonitorBoard.Enabled = (cmbAFSystem = "ADWIN")
    Me.cmbAFMonitorChan.Enabled = (cmbAFSystem = "ADWIN")
    Me.cmbAFRampBoard.Enabled = (cmbAFSystem = "ADWIN")
    Me.cmbAFRampChan.Enabled = (cmbAFSystem = "ADWIN")
    
    'Enable / Disable the ADWIN Advanced settings
    ShowHideADWIN_AFAdvancedSettings ((cmbAFSystem = "ADWIN"))
            
    'Enable the alternate AF monitor combo-boxes if that module has been
    'enabled by the user
    If Me.chkEnableAltAFMonitor.Value = Checked Then
        Me.cmbAltAFMonitorBoard.Enabled = True
        Me.cmbAltAFMonitorChan.Enabled = True
    End If
        
    If NewAFSystem = "2G" Then
    
        'Change the list index on the cmb box
        Me.cmbAFSystem.ListIndex = 0
                
    ElseIf NewAFSystem = "ADWIN" Then
    
        Me.cmbAFSystem.ListIndex = 1
        
    ElseIf NewAFSystem = "MCC" Then
    
        'What the hey!?! This setting
        'shouldn't even be possible
        
        'Set NewAFSystem = Old AFSystem
        NewAFSystem = AFSystem
        
    End If
    
    'Finally, change the global variable
    AFSystem = NewAFSystem
    
'(August 3, 2010 - I Hilburn) Commented this out.  Reloading the form could erase
'                             unsaved user settings changes
'    'Reload the Form
'    Form_Load

End Sub

'Subroutine SetBoardAndChanComboBoxes
'
' Takes in pointers to the two DAQ Board and Board chan combo-box controls
' and another pointer to the channel object that contains the board and channel
' that needs to be set as the active list-indices of the combo-boxes
'
'   Inputs:
'
'   cmbBoard    -   Reference to the combo-box control for the DAQ board related
'                   to a particular comm setting
'
'   cmbChan     -   Reference to the combo-box control for the DAQ channel related
'                   to that some comm settings
'
'   ChanObj     -   Reference to a channel object containing the needed comm settings
'
'
'   Outputs:
'
'   cmbBoard    -   Reference to the now modified DAQ Board combo-box control with
'                   the active list-item (listindex) set to the board-name specified
'                   in ChanObj.BoardName
'
'   cmbChan     -   Ditto, but for the DAQ Chan combo-box control
'
Public Sub SetBoardAndChanComboBoxes _
    (ByRef cmbBoard As ComboBox, _
     ByRef cmbChan As ComboBox, _
     ByRef ChanObj As Channel, _
     Optional ByVal ChanDesc As String = "")
     
    If ChanObj Is Nothing Then Exit Sub
     
    Dim i As Long
    Dim N As Long
    
    'Set N = number of items in the cmbBoard control
    N = cmbBoard.ListCount
    
    'Validate N
    If N < 1 Then
    
        'No boards loaded into the DAQ Board combo-box yet,
        'raise an error
'        Err.Raise -616, _
'                  "frmSettings.ImportSettings->SetBoardAndChanComboBoxes", _
'                  "No DAQ boards loaded into the " & ChanDesc & _
'                  " combo-box control." & vbNewLine & _
'                  "Cannot import DAQ Board settings for the " & _
'                  ChanDesc & " channel."
'
        Exit Sub
        
    End If
    
    'Run through all the items in the DAQ board combo box
    'until one is found that matches the BoardName property of the
    'Channel Object
    For i = 1 To N
    
        If cmbBoard.List(i - 1) = ChanObj.BoardName Then
        
            'Change the DAQ Board combo box to select
            'this list-index
            cmbBoard.ListIndex = i - 1
            
            'This will end the for loop
            i = N + 1
            
        End If
        
    Next i
    
End Sub

Private Sub SetCupButton_Click()
    Dim xPos As Long
    Dim yPos As Long
    
    Dim old_x_pos As Long
    Dim old_y_pos As Long
    
    
    xPos = frmDCMotors.ReadPosition(MotorChanger)
    yPos = frmDCMotors.ReadPosition(MotorChangerY)
    
    XYHolePositionsFlexGrid.Col = 1
    old_x_pos = CLng(Me.XYHolePositionsFlexGrid.text)
    XYHolePositionsFlexGrid.text = xPos
    modConfig.XYTablePositions(XYHolePositionsFlexGrid.row - 1, 0) = xPos
    XYHolePositionsFlexGrid.Col = 2
    old_y_pos = CLng(Me.XYHolePositionsFlexGrid.text)
    XYHolePositionsFlexGrid.text = yPos
    modConfig.XYTablePositions(XYHolePositionsFlexGrid.row - 1, 1) = yPos
   
    If XYHolePositionsFlexGrid.row = 1 Then
    'This is the home row
        Dim UserResp As Long
    
        'Prompt user that this change will recalculate all cups
        ' The following code is specific to the APS-designed XY table
        UserResp = MsgBox("Making changes to the Home Cup position will recalculate " & _
                          "the X and Y position of all other cups." & vbNewLine & vbNewLine & _
                          "All manually entered cup positions will be recalculated." & _
                          vbNewLine & vbNewLine & _
                          "Are you sure you want to recalibrate all cup positions to this home position?", _
                          vbYesNo, _
                          "Warning!")
                          
        If UserResp = vbYes Then
            frmDCMotors.RelabelPos MotorChanger, 0
            frmDCMotors.RelabelPos MotorChangerY, 0
            xPos = frmDCMotors.ReadPosition(MotorChanger)
            yPos = frmDCMotors.ReadPosition(MotorChangerY)
        
            XYHolePositionsFlexGrid.Col = 1
            XYHolePositionsFlexGrid.text = xPos
            modConfig.XYTablePositions(XYHolePositionsFlexGrid.row - 1, 0) = xPos
            XYHolePositionsFlexGrid.Col = 2
            XYHolePositionsFlexGrid.text = yPos
            modConfig.XYTablePositions(XYHolePositionsFlexGrid.row - 1, 1) = yPos
        End If
    
    End If
    
    GetCurrentXYPositions_AndSaveToXYCurrentMotorPositionControls
End Sub

Private Sub ShowHideADWIN_AFAdvancedSettings(ByVal isShow As Boolean)
        
    'Show/Hide the Advanced ADWIN AF settings
    Me.lblAFMaxRampDownNumPeriods.Enabled = (isShow = True)
    Me.lblAFMinRampDownNumPeriods.Enabled = (isShow = True)
    Me.lblAFMinRampUpTime.Enabled = (isShow = True)
    Me.lblAFMaxRampUpTime.Enabled = (isShow = True)
    Me.txtAFMaxRampDownNumPeriods.Enabled = (isShow = True)
    Me.txtAFMaxRampUpTime.Enabled = (isShow = True)
    Me.txtAFMinRampDownNumPeriods.Enabled = (isShow = True)
    Me.txtAFMinRampUpTime.Enabled = (isShow = True)
    Me.lblRampDownTimeHeader.Enabled = (isShow = True)
    Me.lblResultingRampDownTime.Enabled = (isShow = True)
    Me.lblSampleRampOutputVoltHeader.Enabled = (isShow = True)
    Me.txtSampleRampOutputVoltage.Enabled = (isShow = True)
    Me.lblRampOutputMaxVoltHeader.Enabled = (isShow = True)
    Me.lblAFAxialRampMax.Enabled = (isShow = True)
    Me.lblAFAxialRampMaxLabel.Enabled = (isShow = True)
    Me.lblAFTransverseRampMax.Enabled = (isShow = True)
    Me.lblAFTransverseRampMaxLabel.Enabled = (isShow = True)
    Me.lblAFRampDownPeriodsPerVolt.Enabled = (isShow = True)
    Me.txtAFRampDownPeriodsPerVolt.Enabled = (isShow = True)
    Me.txtAFAxialRampUpVoltsPerSec.Enabled = (isShow = True)
    Me.lblAFAxialRampUpVoltsPerSec.Enabled = (isShow = True)
    Me.txtAFTransverseRampUpVoltsPerSec.Enabled = (isShow = True)
    Me.lblAFTransverseRampUpVoltsPerSec.Enabled = (isShow = True)
         
End Sub

Private Sub SortInterpolationRanges()

    If interpolation_ranges Is Nothing Then Exit Sub
    If interpolation_ranges.Count <= 1 Then Exit Sub
    
    BubbleSortInterpolationRanges

End Sub

Private Sub Swap(ByRef range1 As InterpolationRange, ByRef range2 As InterpolationRange)

    Dim temp As InterpolationRange
    Set temp = New InterpolationRange

    temp.StartRow = range1.StartRow
    temp.EndRow = range1.EndRow
    
    range1.StartRow = range2.StartRow
    range1.EndRow = range2.EndRow

    range2.StartRow = temp.StartRow
    range2.EndRow = temp.EndRow
    
    Set temp = Nothing

End Sub


'Private Sub UpdateXYGridCheckboxesByRowAndCol(ByVal row_num As Long, _
'                                              ByVal col_num As Long)
'
'    If interpolation_ranges Is Nothing Then
'        Set interpolation_ranges = New InterpolationRanges
'    End If
'
'    If interpolation_ranges.Count = 0 Then
'        interpolation_ranges.Add -1, -1
'    End If
'
'    Dim i As Integer
'
'    With Me.XYHolePositionsFlexGrid
'
'        .row = row_num
'        .Col = col_num
'
'        Dim temp_range As InterpolationRange
'
'        If .CellPicture = Me.picUnchecked.Picture Then
'
'            Set .CellPicture = Me.picChecked.Picture
'
'            SortInterpolationRanges
'
'            For i = 0 To interpolation_ranges.Count - 1
'
'                With interpolation_ranges(i)
'
'                    If .StartRow = -1 And _
'                       .EndRow > row_num _
'                    Then
'
'                       .StartRow = row_num
'                       MarkAsDisabled .StartRow, .EndRow
'                       SortInterpolationRanges
'                       Exit Sub
'
'                    ElseIf .StartRow <> -1 And _
'                           .StartRow < row_num And _
'                           .EndRow = -1 _
'                    Then
'
'                        .EndRow = row_num
'                        MarkAsDisabled .StartRow, .EndRow
'                        SortInterpolationRanges
'                        Exit Sub
'
'                    ElseIf .StartRow = -1 And _
'                           .EndRow <> -1 And _
'                           .EndRow < row_num _
'                    Then
'
'                        .StartRow = .EndRow
'                        .EndRow = row_num
'                        MarkAsDisabled .StartRow, .EndRow
'                        SortInterpolationRanges
'                        Exit Sub
'
'                    ElseIf .StartRow = -1 And _
'                            .EndRow = -1 _
'                    Then
'
'                       .StartRow = row_num
'                       SortInterpolationRanges
'                       Exit Sub
'
'                    End If
'
'                End With
'
'            Next
'
'            'If made it this far, need to add additional array element for a new range
'            interpolation_ranges.Add row_num, -1
'
'        ElseIf .CellPicture = Me.picChecked.Picture Then
'
'            Set .CellPicture = Me.picUnchecked.Picture
'
'            'Find matching interpolation range in the array
'            For i = 0 To interpolation_ranges.Count - 1
'
'                If interpolation_ranges(i).StartRow = row_num Then
'
'                    interpolation_ranges(i).StartRow = -1
'
'                    If interpolation_ranges(i).EndRow = -1 And _
'                       interpolation_ranges.Count > 1 _
'                    Then
'
'                       interpolation_ranges.Remove (i)
'
'                    ElseIf interpolation_ranges(i).EndRow <> -1 Then
'
'                        UnmarkAsDisabled row_num, interpolation_ranges(i).EndRow
'
'                    End If
'
'                    SortInterpolationRanges
'                    Exit Sub
'
'                End If
'
'                If interpolation_ranges(i).EndRow = row_num Then
'
'                    interpolation_ranges(i).EndRow = -1
'
'                    If interpolation_ranges(i).StartRow = -1 And _
'                       interpolation_ranges.Count > 1 _
'                    Then
'
'                        interpolation_ranges.Remove (i)
'
'                    ElseIf interpolation_ranges(i).StartRow > 0 Then
'
'                        UnmarkAsDisabled interpolation_ranges(i).StartRow, row_num
'
'                    End If
'
'                    SortInterpolationRanges
'                    Exit Sub
'
'                End If
'
'            Next i
'
'        End If
'
'    End With
'
'    SortInterpolationRanges
'
'End Sub
'
'Private Sub MarkAsDisabled(ByVal before_row As Long, ByVal after_row As Long)
'
'    Dim i As Long
'
'    With Me.XYHolePositionsFlexGrid
'
'        Dim orig_row As Long: orig_row = .row
'        Dim orig_col As Long: orig_col = .Col
'
'        For i = before_row + 1 To after_row - 1
'
'            .row = i
'            .Col = 3
'
'            Dim range_index As Integer: range_index = interpolation_ranges.GetIndexByRow(i)
'
'            If range_index > -1 Then interpolation_ranges.Remove range_index
'            Set .CellPicture = Me.picDisabled.Picture
'
'        Next i
'
'        .row = orig_row
'        .Col = orig_col
'
'    End With
'
'End Sub
'
'Private Sub UnmarkAsDisabled(ByVal before_row As Long, ByVal after_row As Long)
'
'    Dim i As Long
'
'    With Me.XYHolePositionsFlexGrid
'
'        Dim orig_row As Long: orig_row = .row
'        Dim orig_col As Long: orig_col = .Col
'
'        For i = before_row + 1 To after_row - 1
'
'            .row = i
'            .Col = 3
'
'            Dim range_index As Integer: range_index = interpolation_ranges.GetIndexByRow(i)
'
'            If range_index > -1 Then interpolation_ranges.Remove range_index
'
'            Set .CellPicture = Me.picUnchecked.Picture
'
'        Next i
'
'        .row = orig_row
'        .Col = orig_col
'
'    End With
'
'End Sub

Private Sub tbsARMIRMChannels_Click()

    Dim i As Long
    
    'To start with, hide all of the board/channel frames
    Me.frameARMSet.Visible = False
    Me.frameARMVoltageOut.Visible = False
    Me.frameIRMCapacitorVoltageIn.Visible = False
    Me.frameIRMFire.Visible = False
    Me.frameIRMPowerAmpVoltageIn.Visible = False
    Me.frameIRMTrim.Visible = False
    Me.frameIRMVoltageOut.Visible = False
    Me.frameIRMMonitor.Visible = False
    
    'Get the index of the selected tab
    i = tbsARMIRMChannels.SelectedItem.Index

    'Now use Select statement to hide / show appropriate
    'Board / channel frames depending on the currently
    'selected tab
    Select Case i
    
        Case 1
        'IRM (1) channels are selected
        
            'Show the IRM Voltage out, IRM Capacitor In, and IRM Power Amp In
            'frames
            Me.frameIRMVoltageOut.Visible = True
            Me.frameIRMCapacitorVoltageIn.Visible = True
            Me.frameIRMPowerAmpVoltageIn.Visible = True
            
        Case 2
        'IRM (2) channels are selected
        
            'Show the IRM Fire, Trim, and Monitor frames
            Me.frameIRMFire.Visible = True
            Me.frameIRMTrim.Visible = True
            Me.frameIRMMonitor.Visible = True
            
        Case 3
        'Digital Output Channels is selected
        
            'Show the ARM Set, and ARM voltage out frames
            Me.frameARMVoltageOut.Visible = True
            Me.frameARMSet.Visible = True
                        
    End Select

End Sub

Public Sub tbsOptions_Click()
    selectTab tbsOptions.SelectedItem.Index
End Sub

Private Sub testAllCups_Click()
Dim xPos As Long
Dim yPos As Long
Dim i As Integer

If NOCOMM_MODE Then Exit Sub
If modFlow.Prog_halted Then Exit Sub

If Not modConfig.HasXYTableBeenHomed Then
        
    GetCurrentXYPositions_AndSaveToXYCurrentMotorPositionControls

    frmProgram.StatBarNew "Home to top..."
    modMotor.MotorUPDN_TopReset
    
    UpdateXYCurrentMotorPositionControls_Moving
    
    modMotor.MotorXYTable_CenterReset
    
End If

For i = 1 To SlotMax
    If Not NOCOMM_MODE Then
          
    
        GetCurrentXYPositions_AndSaveToXYCurrentMotorPositionControls
        
        'Move Home
        XYHolePositionsFlexGrid.row = 1
        XYHolePositionsFlexGrid.Col = 1
        xPos = XYHolePositionsFlexGrid.text
        XYHolePositionsFlexGrid.Col = 2
        yPos = XYHolePositionsFlexGrid.text
        
        UpdateXYCurrentMotorPositionControls_Moving
        
        frmDCMotors.MotorStop MotorChanger
        frmDCMotors.MotorStop MotorChangerY
        frmDCMotors.MoveMotorAbsoluteXY MotorChanger, xPos, ChangerSpeed, False, False
        frmDCMotors.MoveMotorAbsoluteXY MotorChangerY, yPos, ChangerSpeed, False, False
        
        GetCurrentXYPositions_AndSaveToXYCurrentMotorPositionControls
        
        Sleep 5000
        'Move To Cup i
        XYHolePositionsFlexGrid.row = i + 1
        XYHolePositionsFlexGrid.Col = 1
        xPos = XYHolePositionsFlexGrid.text
        XYHolePositionsFlexGrid.Col = 2
        yPos = XYHolePositionsFlexGrid.text
        
        UpdateXYCurrentMotorPositionControls_Moving
        
        frmDCMotors.MotorStop MotorChanger
        frmDCMotors.MotorStop MotorChangerY
        frmDCMotors.MoveMotorAbsoluteXY MotorChanger, xPos, ChangerSpeed, False, False
        frmDCMotors.MoveMotorAbsoluteXY MotorChangerY, yPos, ChangerSpeed, False, False
        
        GetCurrentXYPositions_AndSaveToXYCurrentMotorPositionControls
        
        Sleep 5000
    End If
Next i


End Sub

Private Sub txtAFAxialRampUpVoltsPerSec_Change()

    If val(Me.txtAFAxialRampUpVoltsPerSec.text) <> modConfig.AxialRampUpVoltsPerSec Then
        isADwinRampSettings_dirty = True
    End If

End Sub

Private Sub txtAFMaxRampDownNumPeriods_Change()

    If val(Me.txtAFMaxRampDownNumPeriods.text) <> modConfig.MaxRampDown_NumPeriods Then
        isADwinRampSettings_dirty = True
    End If


    txtSampleRampOutputVoltage_Change

End Sub

Private Sub txtAFMaxRampUpTime_Change()

    If val(Me.txtAFMaxRampUpTime.text) <> modConfig.MaxRampUpTime_ms Then
        isADwinRampSettings_dirty = True
    End If
    
End Sub

Private Sub txtAFMinRampDownNumPeriods_Change()

    If val(Me.txtAFMinRampDownNumPeriods.text) <> modConfig.MinRampDown_NumPeriods Then
        isADwinRampSettings_dirty = True
    End If

    txtSampleRampOutputVoltage_Change

End Sub

Private Sub txtAFMinRampUpTime_Change()

   If val(Me.txtAFMinRampUpTime.text) <> modConfig.MinRampUpTime_ms Then
        isADwinRampSettings_dirty = True
    End If

End Sub

Private Sub txtAFRampDownPeriodsPerVolt_Change()

    If val(Me.txtAFRampDownPeriodsPerVolt.text) <> modConfig.RampDownNumPeriodsPerVolt Then
        isADwinRampSettings_dirty = True
    End If

    txtSampleRampOutputVoltage_Change

End Sub

Private Sub txtAFTransverseRampUpVoltsPerSec_Change()

    If val(Me.txtAFTransverseRampUpVoltsPerSec.text) <> modConfig.TransRampUpVoltsPerSec Then
        isADwinRampSettings_dirty = True
    End If

End Sub

Private Sub txtEditGridCell_Change()

    'Need to save the contents of the cell edit text-box as they are now to
    'the flex-grid cell that it's ghosting for
   On Error GoTo txtEditGridCell_Change_Error

    With Me.XYHolePositionsFlexGrid
    
        .row = CurrentCell(0)
        .Col = CurrentCell(1)
        
        Dim old_value As String: old_value = Trim(.text)
        
        .text = Me.txtEditGridCell.text
                
    End With

   On Error GoTo 0
   Exit Sub

txtEditGridCell_Change_Error:

    MsgBox "Error " & Err.number & " (" & Err.Description & ") in procedure txtEditGridCell_Change of Form frmSettings"

End Sub

Private Sub txtEditGridCell_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Dim OldPos As Long
    
    With Me.XYHolePositionsFlexGrid

        If KeyCode = vbEnter Or _
           KeyCode = vbKeyDown _
        Then
        
            'User has selected to shift to the cell below
        
            'Save the current value of txtEditGridCell to the flex-grid cell
            'that it's stealthily filling in for
            .row = CurrentCell(0)
            .Col = CurrentCell(1)
            .text = Me.txtEditGridCell.text
            
            If .row >= .Rows Then Exit Sub
            
            'Otherwise, there is a row beyond this
            'Activate the mouse-down event for the next row, same col
            'Advance to the next row
            .row = .row + 1
            CurrentCell(0) = .row
            Me.txtEditGridCell.text = .text
            
           XYHolePositionsFlexGrid_MouseDown vbLeftButton, _
                                              0, _
                                              CurrentCellPos(0), _
                                              CurrentCellPos(1) + .CellHeight
                                      
            Exit Sub
            
        ElseIf KeyCode = vbKeyUp Then
        
            'User has selected to shift to the cell above
        
            'Save the current value of txtEditGridCell to the flex-grid cell
            'that it's stealthily filling in for
            .row = CurrentCell(0)
            .Col = CurrentCell(1)
            
            'Check for .Col = 0
            If .Col = 0 Then
                CurrentCell(1) = 1
                .Col = 1
            End If
            
            .text = Me.txtEditGridCell.text
            
            'Is this cell in the first data row?
            'if so, exit the sub without doing anything
            If .row < 1 Then Exit Sub
            
            'Retreat to the prior row
            .row = .row - 1
            CurrentCell(0) = .row
            Me.txtEditGridCell.text = .text
            
            'Activate the mouse-down event for the above row, same col
            XYHolePositionsFlexGrid_MouseDown vbLeftButton, _
                                                0, _
                                                CurrentCellPos(0), _
                                                CurrentCellPos(1) - .CellHeight
                                      
            Exit Sub
            
        ElseIf KeyCode = vbKeyRight And Shift = vbShiftMask Then
        
            'User has selected to shift to the cell above
        
            'Save the current value of txtEditGridCell to the flex-grid cell
            'that it's stealthily filling in for
            .row = CurrentCell(0)
            .Col = CurrentCell(1)
            .text = Me.txtEditGridCell.text
            
            If .Col + 2 >= .Cols Then Exit Sub
            
            'Move right one col
            .Col = .Col + 1
            CurrentCell(1) = .Col
            Me.txtEditGridCell.text = .text
            
            'Activate the mouse-down event for the same row, next col
            XYHolePositionsFlexGrid_MouseDown vbLeftButton, _
                                                0, _
                                                CurrentCellPos(0) - (.CellWidth + .ColWidth(.Col)) / 2, _
                                                CurrentCellPos(1)
                                      
            Exit Sub
            
        ElseIf KeyCode = vbKeyLeft And Shift = vbShiftMask Then
        
            'User has selected to shift to the next cell over
        
            'Save the current value of txtEditGridCell to the flex-grid cell
            'that it's stealthily filling in for
            .row = CurrentCell(0)
            .Col = CurrentCell(1)
            .text = Me.txtEditGridCell.text
            
            If .Col <= 1 Then Exit Sub
            
            'Move left one col
            .Col = .Col - 1
            CurrentCell(1) = .Col
            Me.txtEditGridCell.text = .text
            
            'Activate the mouse-down event for the same row, prior col
            XYHolePositionsFlexGrid_MouseDown vbLeftButton, _
                                                0, _
                                                CurrentCellPos(0) + (.CellWidth + .ColWidth(.Col)) / 2, _
                                                CurrentCellPos(1)
                                      
            Exit Sub
            
        ElseIf KeyCode = vbKeyPageUp Then
        
            'User has selected to jump up ten cells
            'Save the current value of txtEditGridCell to the flex-grid cell
            'that it's stealthily filling in for
            .row = CurrentCell(0)
            .Col = CurrentCell(1)
            .text = Me.txtEditGridCell.text
            
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
                Me.txtEditGridCell.text = .text
                
                'Activate the mouse-down event for the target row, same col
                XYHolePositionsFlexGrid_MouseDown vbLeftButton, _
                                                    0, _
                                                    CurrentCellPos(0), _
                                                    CurrentCellPos(1) - (OldPos - 1) * .CellHeight
            
            Else
            
                'It's safe to jump a full 10 rows
                .row = .row - 10
                CurrentCell(0) = .row
                Me.txtEditGridCell.text = .text
                
                'Activate the mouse-down event for the target row, same col
                XYHolePositionsFlexGrid_MouseDown vbLeftButton, _
                                                    0, _
                                                    CurrentCellPos(0), _
                                                    CurrentCellPos(1) - 10 * .CellHeight
            
            End If
            
        ElseIf KeyCode = vbKeyPageDown Then
        
            'User has selected to jump down ten cells
            'Save the current value of txtEditGridCell to the flex-grid cell
            'that it's stealthily filling in for
            .row = CurrentCell(0)
            .Col = CurrentCell(1)
            .text = Me.txtEditGridCell.text
            
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
                Me.txtEditGridCell.text = .text
                
                'Activate the mouse-down event for the target row, same col
                XYHolePositionsFlexGrid_MouseDown vbLeftButton, _
                                                    0, _
                                                    CurrentCellPos(0), _
                                                    CurrentCellPos(1) + (OldPos - 1) * .CellHeight
            
            Else
            
                'It's safe to jump a full 10 rows
                .row = .row + 10
                CurrentCell(0) = .row
                Me.txtEditGridCell.text = .text
                
                'Activate the mouse-down event for the target row, same col
                XYHolePositionsFlexGrid_MouseDown vbLeftButton, _
                                                    0, _
                                                    CurrentCellPos(0), _
                                                    CurrentCellPos(1) + 10 * .CellHeight
            
            End If
            
        End If
            
    End With
        
End Sub

Private Sub txtSampleRampOutputVoltage_Change()

    Dim TempL As Long

    'Need to update the Resulting Ramp Down number of periods
    
    'Calculate the number of ramp-down periods for this ramp voltage
    'using the Ramp Periods / Ramp Volts ratio
    TempL = CLng(val(Me.txtAFRampDownPeriodsPerVolt) * _
                 val(Me.txtSampleRampOutputVoltage))
                 
    'Check to see if the calculated value is within bounds
    If TempL < val(Me.txtAFMinRampDownNumPeriods) Then
    
        TempL = val(Me.txtAFMinRampDownNumPeriods)
        
    End If
    
    If TempL > val(Me.txtAFMaxRampDownNumPeriods) Then
    
        TempL = val(Me.txtAFMaxRampDownNumPeriods)
        
    End If
    
    Dim axial_ramp_down_time_ms As Double
    Dim transverse_ramp_down_time_ms As Double
    Dim caption_string As String
        
    If modConfig.AfAxialResFreq <> 0 Then
    
        axial_ramp_down_time_ms = TempL * (1 / modConfig.AfAxialResFreq) * 1000
        caption_string = "Axial: " & Format(axial_ramp_down_time_ms, "#0.00")
        
        
    Else
    
        caption_string = "ERROR: Axial Res. Freq. = 0"
        
    End If
        
        
            
    If modConfig.AfTransResFreq <> 0 Then
    
        transverse_ramp_down_time_ms = TempL * (1 / modConfig.AfTransResFreq) * 1000
        
        caption_string = caption_string & vbNewLine & "Transverse: " & Format(transverse_ramp_down_time_ms, "#0.00")
        
    Else
    
        caption_string = caption_string & vbNewLine & "ERROR: Trans. Res. Freq. = 0"
        
    End If
    
    lblResultingRampDownTime.Caption = caption_string

End Sub

Private Sub UpdateXYCurrentMotorPositionControls(ByVal x_position As Long, _
                                                 ByVal y_position As Long, _
                                                 ByVal xy_cup_position As Long)

    Me.txtMotorPosition(0).text = Format(x_position, "#0")
    Me.txtMotorPosition(1).text = Format(y_position, "#0")
    Me.txtMotorPosition(2).text = Format(xy_cup_position, "#0")
    
End Sub

Private Sub UpdateXYCurrentMotorPositionControls_Moving()

    Me.txtMotorPosition(0).text = "Moving..."
    Me.txtMotorPosition(1).text = "Moving..."
    Me.txtMotorPosition(2).text = "Moving..."

End Sub

Private Sub UpdateXYCurrentMotorPositionControls_Error()

    Me.txtMotorPosition(0).text = "Error!"
    Me.txtMotorPosition(1).text = "Error!"
    Me.txtMotorPosition(2).text = "Error!"

End Sub

Private Sub XNegButton_Click()
    frmDCMotors.MotorStop MotorChanger

    If Not modConfig.HasXYTableBeenHomed Then
    
        frmProgram.StatBarNew "Home to top..."
        modMotor.MotorUPDN_TopReset
        
        UpdateXYCurrentMotorPositionControls_Moving
        
        modMotor.MotorXYTable_CenterReset
        
        GetCurrentXYPositions_AndSaveToXYCurrentMotorPositionControls
    
    End If

    UpdateXYCurrentMotorPositionControls_Moving

    frmDCMotors.MoveMotorXY MotorChanger, val(LTrim$(MoveCounts.text)), val(LTrim$(MoveSpeed.text)), False, False
    
    GetCurrentXYPositions_AndSaveToXYCurrentMotorPositionControls
End Sub

Private Sub XPosButton_Click()
    frmDCMotors.MotorStop MotorChanger

    If Not modConfig.HasXYTableBeenHomed Then
    
        frmProgram.StatBarNew "Home to top..."
        modMotor.MotorUPDN_TopReset
        
        UpdateXYCurrentMotorPositionControls_Moving
        
        modMotor.MotorXYTable_CenterReset
        
        GetCurrentXYPositions_AndSaveToXYCurrentMotorPositionControls
    
    End If

    UpdateXYCurrentMotorPositionControls_Moving

    frmDCMotors.MoveMotorXY MotorChanger, -val(LTrim$(MoveCounts.text)), val(LTrim$(MoveSpeed.text)), False, False
    
    GetCurrentXYPositions_AndSaveToXYCurrentMotorPositionControls
End Sub

Private Sub XYHolePositionsFlexGrid_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    'When user clicks grid with the mouse, need to allow them to
    'edit the data cell they have clicked on (and nothing else)
    
    'User must left click to enter the cell
    If Button <> vbLeftButton Then Exit Sub
    
    'Now set the active cell to match the mouse-down cell
    With Me.XYHolePositionsFlexGrid
       
        CurrentCell(0) = .row
        CurrentCell(1) = .Col
       
        'if this is a fixed row or column, exit this sub
        If .Col = 0 Or .row = 0 Or .RowSel <> .row Or .ColSel <> .Col Then
        
            Me.txtEditGridCell.Visible = False
            Exit Sub
                        
        End If
        
        'Now size and position the edit text box
        Me.txtEditGridCell.text = .text
        
        Me.txtEditGridCell.Top = .RowPos(CurrentCell(0)) + .Top
        Me.txtEditGridCell.Left = .ColPos(CurrentCell(1)) + .Left + 10
        On Error Resume Next
            
            Me.txtEditGridCell.Width = .CellWidth + .GridLineWidth * 20
            
            If Err.number <> 0 Then
            
                'User has selected to click on a cell that is not fully displayed
                'need to deactivate txtEditGridCell
                Me.txtEditGridCell.Visible = False
                
                Exit Sub
                
            End If
            
        On Error GoTo 0
        
        Me.txtEditGridCell.Height = .CellHeight
        
        'Set current cell position
        CurrentCellPos(0) = CSng(txtEditGridCell.Top + txtEditGridCell.Height / 2)
        CurrentCellPos(1) = CSng(txtEditGridCell.Left + txtEditGridCell.Width / 2)
        
        'Show the cell-edit textbox
        Me.txtEditGridCell.ZOrder 0
        Me.txtEditGridCell.Visible = True
        Me.txtEditGridCell.SetFocus
    
    End With

End Sub

Private Sub XYHolePositionsFlexGrid_Scroll()

    'Save the value of the the Edit cell textbox
    With Me.XYHolePositionsFlexGrid
    
        .row = CurrentCell(0)
        .Col = CurrentCell(1)
        If Len(Trim(Me.txtEditGridCell.text)) > 0 Then .text = Me.txtEditGridCell
        
    End With
    
    'Hide the Edit cell textbox
    Me.txtEditGridCell.Visible = False

End Sub

Private Sub YNegButton_Click()
    frmDCMotors.MotorStop MotorChangerY

    If Not modConfig.HasXYTableBeenHomed Then
    
        frmProgram.StatBarNew "Home to top..."
        modMotor.MotorUPDN_TopReset
        
        UpdateXYCurrentMotorPositionControls_Moving
        
        modMotor.MotorXYTable_CenterReset
    
        GetCurrentXYPositions_AndSaveToXYCurrentMotorPositionControls
    
    End If
    
    UpdateXYCurrentMotorPositionControls_Moving

    frmDCMotors.MoveMotorXY MotorChangerY, -val(LTrim$(MoveCounts.text)), val(LTrim$(MoveSpeed.text)), False, False
    
    GetCurrentXYPositions_AndSaveToXYCurrentMotorPositionControls
End Sub

Private Sub YPosButton_Click()
    frmDCMotors.MotorStop MotorChangerY
    
    If Not modConfig.HasXYTableBeenHomed Then
    
        frmProgram.StatBarNew "Home to top..."
        modMotor.MotorUPDN_TopReset
        
        UpdateXYCurrentMotorPositionControls_Moving
        
        modMotor.MotorXYTable_CenterReset
        
        GetCurrentXYPositions_AndSaveToXYCurrentMotorPositionControls
    
    End If
    
    UpdateXYCurrentMotorPositionControls_Moving
    
    frmDCMotors.MoveMotorXY MotorChangerY, val(LTrim$(MoveCounts.text)), val(LTrim$(MoveSpeed.text)), False, False
     
     GetCurrentXYPositions_AndSaveToXYCurrentMotorPositionControls
End Sub

