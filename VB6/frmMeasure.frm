VERSION 5.00
Begin VB.Form frmMeasure 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Measurement Window"
   ClientHeight    =   8745
   ClientLeft      =   210
   ClientTop       =   315
   ClientWidth     =   15405
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8745
   ScaleWidth      =   15405
   Visible         =   0   'False
   Begin VB.PictureBox MomentX 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      Height          =   3660
      Left            =   5400
      ScaleHeight     =   1.14
      ScaleLeft       =   -0.16
      ScaleMode       =   0  'User
      ScaleTop        =   -0.07
      ScaleWidth      =   1.2
      TabIndex        =   111
      Top             =   5000
      Width           =   5000
      Begin VB.CheckBox ChkAllSteps 
         BackColor       =   &H80000005&
         Caption         =   "Display steps above current"
         Height          =   195
         Left            =   1440
         TabIndex        =   114
         Top             =   0
         Width           =   2295
      End
   End
   Begin VB.CheckBox ChkX 
      Caption         =   "Susceptibility"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   10440
      TabIndex        =   112
      Top             =   5880
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin VB.CheckBox ChkM 
      Caption         =   "Moment magnitude"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   10440
      TabIndex        =   113
      Top             =   5640
      Value           =   1  'Checked
      Width           =   1935
   End
   Begin VB.Frame framJumps 
      Height          =   1935
      Left            =   5400
      TabIndex        =   88
      Top             =   5040
      Width           =   4935
      Begin VB.Label lblRed 
         Caption         =   "Red = noise higher than 5 times the moment, I'm not redoing it..."
         Height          =   255
         Left            =   120
         TabIndex        =   104
         Top             =   1560
         Width           =   4695
      End
      Begin VB.Label lblOrange 
         Caption         =   "Orange = order of magnitude (1 to 5) of the moment, be attentive..."
         Height          =   255
         Left            =   120
         TabIndex        =   103
         Top             =   1320
         Width           =   4695
      End
      Begin VB.Label lblWarning 
         Alignment       =   2  'Center
         Caption         =   "It's up to you to retry manually that measurement !!!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   102
         Top             =   120
         Visible         =   0   'False
         Width           =   4695
      End
      Begin VB.Label lblDeltaX 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1680
         TabIndex        =   101
         Top             =   600
         Width           =   975
      End
      Begin VB.Label lblDeltaY 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2760
         TabIndex        =   100
         Top             =   600
         Width           =   975
      End
      Begin VB.Label lblDeltaZ 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   3840
         TabIndex        =   99
         Top             =   600
         Width           =   975
      End
      Begin VB.Label lblRatioX 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1680
         TabIndex        =   98
         Top             =   960
         Width           =   975
      End
      Begin VB.Label lblRatioY 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2760
         TabIndex        =   97
         Top             =   960
         Width           =   975
      End
      Begin VB.Label lblRatioZ 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   3840
         TabIndex        =   96
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label29 
         Alignment       =   2  'Center
         Caption         =   "X"
         Height          =   255
         Left            =   1680
         TabIndex        =   95
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         Caption         =   "Y"
         Height          =   255
         Left            =   2760
         TabIndex        =   94
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label32 
         Alignment       =   2  'Center
         Caption         =   "Z"
         Height          =   255
         Left            =   3840
         TabIndex        =   93
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label33 
         Caption         =   "D"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   8.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   92
         Top             =   600
         Width           =   135
      End
      Begin VB.Label Label34 
         Caption         =   "4 Positions (emu)"
         Height          =   255
         Left            =   240
         TabIndex        =   91
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label35 
         Caption         =   "D"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   8.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   90
         Top             =   960
         Width           =   135
      End
      Begin VB.Label Label36 
         Caption         =   "4 Positions/Moment"
         Height          =   255
         Left            =   240
         TabIndex        =   89
         Top             =   960
         Width           =   1425
      End
   End
   Begin VB.OptionButton optBedding 
      Caption         =   "Bedding coordinates"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   12960
      TabIndex        =   106
      Top             =   5760
      Width           =   2415
   End
   Begin VB.OptionButton optGeographic 
      Caption         =   "Geographic coordinates"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   12960
      TabIndex        =   87
      Top             =   5520
      Width           =   2415
   End
   Begin VB.OptionButton optCore 
      Caption         =   "Core coordinates"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   12960
      TabIndex        =   86
      Top             =   5280
      Value           =   -1  'True
      Width           =   1935
   End
   Begin VB.CommandButton buttonPause 
      Caption         =   "Pause run"
      Height          =   372
      Left            =   3840
      TabIndex        =   71
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton buttonHalt 
      Caption         =   "Halt run"
      Height          =   372
      Left            =   2520
      TabIndex        =   70
      Top             =   1680
      Width           =   1092
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   345
      Left            =   1200
      TabIndex        =   65
      Top             =   1680
      Width           =   858
   End
   Begin VB.Frame framStats 
      Height          =   2052
      Left            =   120
      TabIndex        =   44
      Top             =   4920
      Width           =   5055
      Begin VB.CommandButton cmdStats 
         Caption         =   "&Show Stats"
         Enabled         =   0   'False
         Height          =   372
         Left            =   3720
         TabIndex        =   66
         Top             =   240
         Width           =   1092
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "CSD"
         Height          =   255
         Left            =   3480
         TabIndex        =   73
         Top             =   840
         Width           =   495
      End
      Begin VB.Label lblCSD 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3360
         TabIndex        =   72
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label lblavgmag 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   240
         TabIndex        =   69
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Moment"
         Height          =   255
         Left            =   240
         TabIndex        =   68
         Top             =   840
         Width           =   855
      End
      Begin VB.Label lblAvgDec 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1710
         TabIndex        =   60
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label lblAvgInc 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2550
         TabIndex        =   59
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label30 
         Alignment       =   2  'Center
         Caption         =   "Inc"
         Height          =   255
         Left            =   2550
         TabIndex        =   58
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "Dec"
         Height          =   255
         Left            =   1710
         TabIndex        =   57
         Top             =   840
         Width           =   735
      End
      Begin VB.Label lblDSigInduced 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   252
         Left            =   3504
         TabIndex        =   56
         Top             =   1692
         Width           =   1320
      End
      Begin VB.Label lblDSigDrift 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   252
         Left            =   240
         TabIndex        =   55
         Top             =   1692
         Width           =   1320
      End
      Begin VB.Label lblDSigHolder 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   252
         Left            =   1884
         TabIndex        =   54
         Top             =   1692
         Width           =   1320
      End
      Begin VB.Label Label5 
         Caption         =   "Signal/Drift:"
         Height          =   252
         Left            =   240
         TabIndex        =   53
         Top             =   1452
         Width           =   1320
      End
      Begin VB.Label Label6 
         Caption         =   "Signal/Holder:"
         Height          =   252
         Left            =   1884
         TabIndex        =   52
         Top             =   1452
         Width           =   1320
      End
      Begin VB.Label Label7 
         Caption         =   "Signal/Induced:"
         Height          =   252
         Left            =   3504
         TabIndex        =   51
         Top             =   1452
         Width           =   1308
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "Avg. Z"
         Height          =   253
         Left            =   2167
         TabIndex        =   50
         Top             =   242
         Width           =   858
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "Avg. Y"
         Height          =   253
         Left            =   1210
         TabIndex        =   49
         Top             =   242
         Width           =   858
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Caption         =   "Avg. X"
         Height          =   253
         Left            =   242
         TabIndex        =   48
         Top             =   242
         Width           =   847
      End
      Begin VB.Label lblavgx 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   253
         Left            =   242
         TabIndex        =   47
         Top             =   484
         Width           =   858
      End
      Begin VB.Label lblavgy 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   253
         Left            =   1199
         TabIndex        =   46
         Top             =   484
         Width           =   858
      End
      Begin VB.Label lblavgz 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   253
         Left            =   2167
         TabIndex        =   45
         Top             =   484
         Width           =   858
      End
   End
   Begin VB.CommandButton cmdHide 
      Caption         =   "&Hide"
      Height          =   345
      Left            =   120
      TabIndex        =   43
      Top             =   1680
      Width           =   858
   End
   Begin VB.Frame Frame1 
      Height          =   2892
      Left            =   120
      TabIndex        =   8
      Top             =   2016
      Width           =   5055
      Begin VB.CommandButton cmdShowPlots 
         Caption         =   "Show plots"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3960
         TabIndex        =   84
         Top             =   200
         Width           =   975
      End
      Begin VB.Label lblZSQUID 
         Alignment       =   2  'Center
         Height          =   255
         Left            =   2160
         TabIndex        =   110
         Top             =   960
         Width           =   855
      End
      Begin VB.Label lblYSQUID 
         Alignment       =   2  'Center
         Height          =   255
         Left            =   1200
         TabIndex        =   109
         Top             =   960
         Width           =   855
      End
      Begin VB.Label lblXSQUID 
         Alignment       =   2  'Center
         Height          =   255
         Left            =   240
         TabIndex        =   108
         Top             =   960
         Width           =   855
      End
      Begin VB.Label lblRescan 
         Height          =   255
         Left            =   3240
         TabIndex        =   105
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label lblCalcInc 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   3
         Left            =   4080
         TabIndex        =   42
         Top             =   2520
         Width           =   735
      End
      Begin VB.Label lblCalcInc 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   1
         Left            =   4080
         TabIndex        =   41
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label lblCalcInc 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   2
         Left            =   4080
         TabIndex        =   40
         Top             =   2160
         Width           =   735
      End
      Begin VB.Label lblCalcInc 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   0
         Left            =   4080
         TabIndex        =   39
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label lblCalcDec 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   3
         Left            =   3240
         TabIndex        =   38
         Top             =   2520
         Width           =   735
      End
      Begin VB.Label lblCalcDec 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   1
         Left            =   3240
         TabIndex        =   37
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label lblCalcDec 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   2
         Left            =   3240
         TabIndex        =   36
         Top             =   2160
         Width           =   735
      End
      Begin VB.Label lblCalcDec 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   0
         Left            =   3240
         TabIndex        =   35
         Top             =   1455
         Width           =   735
      End
      Begin VB.Label lblMeasZ 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   3
         Left            =   2160
         TabIndex        =   34
         Top             =   2520
         Width           =   855
      End
      Begin VB.Label lblMeasZ 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   1
         Left            =   2160
         TabIndex        =   33
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label lblMeasZ 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   2
         Left            =   2160
         TabIndex        =   32
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label lblMeasZ 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   0
         Left            =   2160
         TabIndex        =   31
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label lblMeasY 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   3
         Left            =   1200
         TabIndex        =   30
         Top             =   2520
         Width           =   855
      End
      Begin VB.Label lblMeasY 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   1
         Left            =   1200
         TabIndex        =   29
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label lblMeasY 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   2
         Left            =   1200
         TabIndex        =   28
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label lblMeasY 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   0
         Left            =   1200
         TabIndex        =   27
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label lblMeasX 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   26
         Top             =   2520
         Width           =   855
      End
      Begin VB.Label lblMeasX 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   25
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label lblMeasX 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   24
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label lblMeasX 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   23
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label lblMeasZZero 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   1
         Left            =   2160
         TabIndex        =   22
         Top             =   720
         Width           =   855
      End
      Begin VB.Label lblMeasYZero 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   1
         Left            =   1200
         TabIndex        =   21
         Top             =   720
         Width           =   855
      End
      Begin VB.Label lblMeasXZero 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Height          =   252
         Index           =   1
         Left            =   240
         TabIndex        =   20
         Top             =   720
         Width           =   850
      End
      Begin VB.Label lblMeasZZero 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   0
         Left            =   2160
         TabIndex        =   19
         Top             =   480
         Width           =   855
      End
      Begin VB.Label lblMeasYZero 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   0
         Left            =   1200
         TabIndex        =   18
         Top             =   480
         Width           =   855
      End
      Begin VB.Label lblMeasXZero 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Height          =   252
         Index           =   0
         Left            =   240
         TabIndex        =   17
         Top             =   480
         Width           =   850
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Caption         =   "X"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         Caption         =   "Y"
         Height          =   255
         Left            =   1200
         TabIndex        =   15
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Caption         =   "Z"
         Height          =   255
         Left            =   2160
         TabIndex        =   14
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         Caption         =   "Dec"
         Height          =   255
         Left            =   3240
         TabIndex        =   13
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         Caption         =   "Inc"
         Height          =   255
         Left            =   4080
         TabIndex        =   12
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label16 
         Caption         =   "XZero:"
         Height          =   255
         Left            =   285
         TabIndex        =   11
         Top             =   240
         Width           =   780
      End
      Begin VB.Label Label17 
         Caption         =   "YZero:"
         Height          =   255
         Left            =   1245
         TabIndex        =   10
         Top             =   240
         Width           =   780
      End
      Begin VB.Label Label18 
         Caption         =   "ZZero:"
         Height          =   255
         Left            =   2205
         TabIndex        =   9
         Top             =   240
         Width           =   780
      End
   End
   Begin VB.TextBox lblSampleHeight 
      Alignment       =   2  'Center
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   4440
      TabIndex        =   76
      Top             =   968
      Width           =   495
   End
   Begin VB.PictureBox EqualArea 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      Height          =   5000
      Left            =   5400
      ScaleHeight     =   1
      ScaleMode       =   0  'User
      ScaleWidth      =   1
      TabIndex        =   78
      Top             =   0
      Width           =   5000
      Begin VB.CommandButton cmdHideEqu 
         BackColor       =   &H80000005&
         Caption         =   "Hide"
         Height          =   345
         Left            =   4320
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   79
         Top             =   4440
         Width           =   495
      End
   End
   Begin VB.TextBox txtZijLines 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   11205
      TabIndex        =   80
      Text            =   "15"
      Top             =   5325
      Width           =   615
   End
   Begin VB.PictureBox Zijderveld 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      Height          =   5000
      Left            =   10400
      ScaleHeight     =   1
      ScaleMode       =   0  'User
      ScaleWidth      =   1
      TabIndex        =   81
      Top             =   0
      Width           =   5000
   End
   Begin VB.CommandButton cmdHideZij 
      Caption         =   "Hide"
      Height          =   285
      Left            =   10440
      TabIndex        =   82
      Top             =   5325
      Width           =   495
   End
   Begin VB.Label lblCupNumber 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   960
      TabIndex        =   116
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label Label37 
      Caption         =   "Cup #:"
      Height          =   255
      Left            =   120
      TabIndex        =   115
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label lblOrientation 
      Height          =   615
      Left            =   10560
      TabIndex        =   107
      Top             =   6120
      Width           =   4695
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Zijderveld [1967] plot (N-S orthographic projection)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10440
      TabIndex        =   85
      Top             =   5040
      Width           =   4935
   End
   Begin VB.Label Label19 
      Caption         =   "cm"
      Height          =   255
      Left            =   5040
      TabIndex        =   77
      Top             =   975
      Width           =   255
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      Caption         =   "Dec"
      Height          =   252
      Left            =   480
      TabIndex        =   67
      Top             =   5280
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.Label Label21 
      Caption         =   "Directions:"
      Height          =   255
      Left            =   120
      TabIndex        =   64
      Top             =   975
      Width           =   855
   End
   Begin VB.Label Label22 
      Caption         =   "Measuring:"
      Height          =   255
      Left            =   1695
      TabIndex        =   63
      Top             =   975
      Width           =   855
   End
   Begin VB.Label Label23 
      Caption         =   "Nb:"
      Height          =   255
      Left            =   3285
      TabIndex        =   74
      Top             =   975
      Width           =   300
   End
   Begin VB.Label lblDirs 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   975
      TabIndex        =   62
      Top             =   960
      Width           =   495
   End
   Begin VB.Label lblMeasDir 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2535
      TabIndex        =   61
      Top             =   975
      Width           =   495
   End
   Begin VB.Label lblMeascount 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   3675
      TabIndex        =   75
      Top             =   975
      Width           =   495
   End
   Begin VB.Label lblDataFileName 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   960
      TabIndex        =   7
      Top             =   600
      Width           =   4185
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblDemag 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4320
      TabIndex        =   6
      Top             =   240
      Width           =   855
   End
   Begin VB.Label lblAvgCycles 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   3000
      TabIndex        =   5
      Top             =   240
      Width           =   495
   End
   Begin VB.Label lblSampName 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   960
      TabIndex        =   4
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label24 
      Caption         =   "File Path:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label25 
      Caption         =   "Demag:"
      Height          =   255
      Left            =   3600
      TabIndex        =   2
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label26 
      Caption         =   "Avg. Cycles"
      Height          =   255
      Index           =   1
      Left            =   2040
      TabIndex        =   1
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label27 
      Caption         =   "Sample:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label28 
      Caption         =   "previous steps"
      Height          =   255
      Left            =   11850
      TabIndex        =   83
      Top             =   5370
      Width           =   1095
   End
End
Attribute VB_Name = "frmMeasure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' arc cosine
' error if NUMBER is outside the range [-1,1]
Function acos(ByVal number As Double) As Double
    If Abs(number) <> 1 Then
        acos = 1.5707963267949 - Atn(number / Sqr(1 - number * number))
    ElseIf number = -1 Then
        acos = 3.14159265358979
    End If
    'elseif number=1 --> Acos=0 (implicit)
End Function

' arc cotangent
' error if NUMBER is zero
Function ACot(Value As Double) As Double
    ACot = Atn(1 / Value)
End Function

' arc cosecant
' error if value is inside the range [-1,1]
Function ACsc(Value As Double) As Double
    ' NOTE: the following lines can be replaced by a single call
    '            ACsc = ASin(1 / value)
    If Abs(Value) <> 1 Then
        ACsc = Atn((1 / Value) / Sqr(1 - 1 / (Value * Value)))
    Else
        ACsc = 1.5707963267949 * Sgn(Value)
    End If
End Function

Private Sub Actualize()
    If optCore.Value = True Then
      If lblOrientation.Visible = True Then
        InitEqualArea
        EqualArea.CurrentX = 0
        EqualArea.CurrentY = 0.92
        EqualArea.FontBold = True
        EqualArea.Print "Core" & vbCrLf & "coordinates"
        EqualArea.FontBold = False
        ImportZijRoutine lblSampName, _
            frmStats.lblCDec.Caption, frmStats.lblCInc.Caption, _
            lblavgmag.Caption, True
        AveragePlotEqualArea frmStats.lblCDec.Caption, frmStats.lblCInc.Caption, frmStats.lblErrAngle.Caption
      Else
        ImportZijRoutine lblSampName, _
            frmStats.lblCDec.Caption, frmStats.lblCInc.Caption, _
            lblavgmag.Caption, False
      End If
    ElseIf optGeographic.Value = True Then
        InitEqualArea
        EqualArea.CurrentX = 0
        EqualArea.CurrentY = 0.92
        EqualArea.FontBold = True
        EqualArea.Print "Geographic" & vbCrLf & "coordinates"
        EqualArea.FontBold = False
        ImportZijRoutine lblSampName, _
        frmStats.lblGDec.Caption, frmStats.lblGInc.Caption, _
        lblavgmag.Caption, True
        AveragePlotEqualArea frmStats.lblGDec.Caption, frmStats.lblGInc.Caption, frmStats.lblErrAngle.Caption
    ElseIf optBedding.Value = True Then
        InitEqualArea
        EqualArea.CurrentX = 0
        EqualArea.CurrentY = 0.92
        EqualArea.FontBold = True
        EqualArea.Print "Bedding" & vbCrLf & "coordinates"
        EqualArea.FontBold = False
        ImportZijRoutine lblSampName, _
        frmStats.lblBDec.Caption, frmStats.lblBInc.Caption, _
        lblavgmag.Caption, True
        AveragePlotEqualArea frmStats.lblBDec.Caption, frmStats.lblBInc.Caption, frmStats.lblErrAngle.Caption
    End If
    frmMeasure.EqualArea.Circle (0.8, 0.04), 0.01, RGB(255, 0, 0)
    frmMeasure.EqualArea.Line (0.8 - 0.01, 0.015 - 0.01)-(0.8 + 0.01, 0.015 + 0.01), 0.01, B
    frmMeasure.EqualArea.Circle (0.89, 0.04), 0.01, RGB(0, 0, 255)
    frmMeasure.EqualArea.Line (0.89 - 0.01, 0.015 - 0.01)-(0.89 + 0.01, 0.015 + 0.01), 0.01, BF
End Sub

' arc secant
' error if value is inside the range [-1,1]
Function ASec(Value As Double) As Double
    ' NOTE: the following lines can be replaced by a single call
    '            ASec = ACos(1 / value)
    If Abs(Value) <> 1 Then
        ASec = 1.5707963267949 - Atn((1 / Value) / Sqr(1 - 1 / (Value * Value)))
    Else
        ASec = 3.14159265358979 * Sgn(Value)
    End If
End Function

' arc sine
' error if value is outside the range [-1,1]
Function ASin(Value As Double) As Double
    If Abs(Value) <> 1 Then
        ASin = Atn(Value / Sqr(1 - Value * Value))
    Else
        ASin = 1.5707963267949 * Sgn(Value)
    End If
End Function

Public Sub AveragePlotEqualArea(ByVal dec As Double, ByVal inc As Double, ByVal CSD As Double)
    ' (August 2007 L Carporzen) Plot of the averaged measurement (holder substracted)
    Dim L0 As Double
    Dim L As Double
    Dim ax As Double
    Dim bx As Double
    Dim ay As Double
    Dim by As Double
    Dim X1 As Double
    Dim X2 As Double
    Dim Y1 As Double
    Dim Y2 As Double
    Dim i As Integer
    If CSD > 180 Then CSD = 0
    L0 = 1 / Sqr(Cos(inc * Pi / 180) * Cos(dec * Pi / 180) * Cos(inc * Pi / 180) * Cos(dec * Pi / 180) + Cos(inc * Pi / 180) * Sin(dec * Pi / 180) * Cos(inc * Pi / 180) * Sin(dec * Pi / 180))
    If inc >= 0 Then ' Down direction
        L = L0 * Sqr(1 - Sin(inc * Pi / 180))
        ' Plot of the average measurement as a black square
        EqualArea.Line ((Sin(dec * Pi / 180) * L / L0) / 2 + 0.5 - 0.01, -(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5 - 0.01)-((Sin(dec * Pi / 180) * L / L0) / 2 + 0.5 + 0.01, -(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5 + 0.01), 0.01, BF
        If CSD > 5 Then ' No calcul for small CSD
        If inc + CSD >= 90 Then
        ' The center of the equal area is include in the a95 which will be draw as a circle with the CSD as radius
        ax = (Sin(dec * Pi / 180) * L / L0) / 2 + 0.5
        ay = -(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5 + Sqr(1 - Sin((90 - CSD) * Pi / 180)) / 2
        bx = (Sin(dec * Pi / 180) * L / L0) / 2 + 0.5 + Sqr(1 - Sin((90 - CSD) * Pi / 180)) / 2
        by = -(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5
        Else
        ' Calcul of the coordinates of the axis of the ellipsoid
        ax = (Sin((dec + ASin(Sin((CSD) * Pi / 180) / Cos(inc * Pi / 180)) * 180 / Pi) * Pi / 180) * Sqr(1 - Sin((ASin(1 - (2 * ((-(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5) - 0.5) / (-(Cos(dec * Pi / 180))) * (2 * ((-(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5) - 0.5) / (-(Cos(dec * Pi / 180)))))) * 180 / Pi) * Pi / 180))) / 2 + 0.5
        ay = Abs(-(Cos((dec + ASin(Sin((CSD) * Pi / 180) / Cos(inc * Pi / 180)) * 180 / Pi) * Pi / 180) * Sqr(1 - Sin((ASin(1 - (2 * ((-(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5) - 0.5) / (-(Cos(dec * Pi / 180))) * (2 * ((-(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5) - 0.5) / (-(Cos(dec * Pi / 180)))))) * 180 / Pi) * Pi / 180))) / 2 + 0.5)
        bx = (Sin(dec * Pi / 180) * Sqr(1 - Sin((inc + CSD) * Pi / 180))) / 2 + 0.5
        by = Abs(-(Cos(dec * Pi / 180) * Sqr(1 - Sin((inc + CSD) * Pi / 180))) / 2 + 0.5)
        If ay > 1 Then ay = 1 - (ay - 1)
        If by > 1 Then by = 1 - (by - 1)
        End If
        ' Plot of the ellipsoid/circle by small segments (5 degrees)
        For i = 0 To 90
            Y1 = (Sqr(((Sin(dec * Pi / 180) * L / L0) / 2 + 0.5 - bx) * ((Sin(dec * Pi / 180) * L / L0) / 2 + 0.5 - bx) + (-(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5 - by) * (-(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5 - by))) * Sin(i * Pi / 180)
            X1 = Sqr((Sqr(((Sin(dec * Pi / 180) * L / L0) / 2 + 0.5 - ax) * ((Sin(dec * Pi / 180) * L / L0) / 2 + 0.5 - ax) + (-(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5 - ay) * (-(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5 - ay))) ^ 2 * (1 - Sin(i * Pi / 180) * Sin(i * Pi / 180)))
            Y2 = (Sqr(((Sin(dec * Pi / 180) * L / L0) / 2 + 0.5 - bx) * ((Sin(dec * Pi / 180) * L / L0) / 2 + 0.5 - bx) + (-(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5 - by) * (-(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5 - by))) * Sin((i + 1) * Pi / 180)
            X2 = Sqr((Sqr(((Sin(dec * Pi / 180) * L / L0) / 2 + 0.5 - ax) * ((Sin(dec * Pi / 180) * L / L0) / 2 + 0.5 - ax) + (-(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5 - ay) * (-(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5 - ay))) ^ 2 * (1 - Sin((i + 1) * Pi / 180) * Sin((i + 1) * Pi / 180)))
            ' Test to don't plot the parts of the ellipsoid/circle which are outside of the plane inc = 0
            If (Abs(-(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5 - X1 * Sin((-dec) * Pi / 180) + Y1 * Cos((-dec) * Pi / 180) - 0.5)) ^ 2 < Abs(0.5 ^ 2 - ((Sin(dec * Pi / 180) * L / L0) / 2 + 0.5 + X1 * Cos((-dec) * Pi / 180) + Y1 * Sin((-dec) * Pi / 180) - 0.5) ^ 2) Then EqualArea.Line ((Sin(dec * Pi / 180) * L / L0) / 2 + 0.5 + X1 * Cos((-dec) * Pi / 180) + Y1 * Sin((-dec) * Pi / 180), -(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5 - X1 * Sin((-dec) * Pi / 180) + Y1 * Cos((-dec) * Pi / 180))-((Sin(dec * Pi / 180) * L / L0) / 2 + 0.5 + X2 * Cos((-dec) * Pi / 180) + Y2 * Sin((-dec) * Pi / 180), -(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5 - X2 * Sin((-dec) * Pi / 180) + Y2 * Cos((-dec) * Pi / 180))
            If (Abs(-(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5 - X1 * Sin((-dec) * Pi / 180) - Y1 * Cos((-dec) * Pi / 180) - 0.5)) ^ 2 < Abs(0.5 ^ 2 - ((Sin(dec * Pi / 180) * L / L0) / 2 + 0.5 + X1 * Cos((-dec) * Pi / 180) - Y1 * Sin((-dec) * Pi / 180) - 0.5) ^ 2) Then EqualArea.Line ((Sin(dec * Pi / 180) * L / L0) / 2 + 0.5 + X1 * Cos((-dec) * Pi / 180) - Y1 * Sin((-dec) * Pi / 180), -(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5 - X1 * Sin((-dec) * Pi / 180) - Y1 * Cos((-dec) * Pi / 180))-((Sin(dec * Pi / 180) * L / L0) / 2 + 0.5 + X2 * Cos((-dec) * Pi / 180) - Y2 * Sin((-dec) * Pi / 180), -(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5 - X2 * Sin((-dec) * Pi / 180) - Y2 * Cos((-dec) * Pi / 180))
            If (Abs(-(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5 + X1 * Sin((-dec) * Pi / 180) + Y1 * Cos((-dec) * Pi / 180) - 0.5)) ^ 2 < Abs(0.5 ^ 2 - ((Sin(dec * Pi / 180) * L / L0) / 2 + 0.5 - X1 * Cos((-dec) * Pi / 180) + Y1 * Sin((-dec) * Pi / 180) - 0.5) ^ 2) Then EqualArea.Line ((Sin(dec * Pi / 180) * L / L0) / 2 + 0.5 - X1 * Cos((-dec) * Pi / 180) + Y1 * Sin((-dec) * Pi / 180), -(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5 + X1 * Sin((-dec) * Pi / 180) + Y1 * Cos((-dec) * Pi / 180))-((Sin(dec * Pi / 180) * L / L0) / 2 + 0.5 - X2 * Cos((-dec) * Pi / 180) + Y2 * Sin((-dec) * Pi / 180), -(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5 + X2 * Sin((-dec) * Pi / 180) + Y2 * Cos((-dec) * Pi / 180))
            If (Abs(-(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5 + X1 * Sin((-dec) * Pi / 180) - Y1 * Cos((-dec) * Pi / 180) - 0.5)) ^ 2 < Abs(0.5 ^ 2 - ((Sin(dec * Pi / 180) * L / L0) / 2 + 0.5 - X1 * Cos((-dec) * Pi / 180) - Y1 * Sin((-dec) * Pi / 180) - 0.5) ^ 2) Then EqualArea.Line ((Sin(dec * Pi / 180) * L / L0) / 2 + 0.5 - X1 * Cos((-dec) * Pi / 180) - Y1 * Sin((-dec) * Pi / 180), -(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5 + X1 * Sin((-dec) * Pi / 180) - Y1 * Cos((-dec) * Pi / 180))-((Sin(dec * Pi / 180) * L / L0) / 2 + 0.5 - X2 * Cos((-dec) * Pi / 180) - Y2 * Sin((-dec) * Pi / 180), -(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5 + X2 * Sin((-dec) * Pi / 180) - Y2 * Cos((-dec) * Pi / 180))
        Next i
        End If
    Else ' Up direction
        L = L0 * Sqr(1 + Sin(inc * Pi / 180))
        ' Plot of the average measurement as a white square
        EqualArea.Line ((Sin(dec * Pi / 180) * L / L0) / 2 + 0.5 - 0.01, -(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5 - 0.01)-((Sin(dec * Pi / 180) * L / L0) / 2 + 0.5 + 0.01, -(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5 + 0.01), 0.01, B
        If CSD > 5 Then ' No calcul for small CSD
        If Abs(inc) + CSD >= 90 Then
        ' The center of the equal area is include in the a95 which will be draw as a circle with the CSD as radius
        ax = (Sin(dec * Pi / 180) * L / L0) / 2 + 0.5
        ay = -(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5 + Sqr(1 - Sin((90 - CSD) * Pi / 180)) / 2
        bx = (Sin(dec * Pi / 180) * L / L0) / 2 + 0.5 + Sqr(1 - Sin((90 - CSD) * Pi / 180)) / 2
        by = -(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5
        Else
        ' Calcul of the coordinates of the axis of the ellipsoid
        ax = (Sin((dec + ASin(Sin((CSD) * Pi / 180) / Cos(Abs(inc) * Pi / 180)) * 180 / Pi) * Pi / 180) * Sqr(1 - Sin((ASin(1 - (2 * ((-(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5) - 0.5) / (-(Cos(dec * Pi / 180))) * (2 * ((-(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5) - 0.5) / (-(Cos(dec * Pi / 180)))))) * 180 / Pi) * Pi / 180))) / 2 + 0.5
        ay = Abs(-(Cos((dec + ASin(Sin((CSD) * Pi / 180) / Cos(Abs(inc) * Pi / 180)) * 180 / Pi) * Pi / 180) * Sqr(1 - Sin((ASin(1 - (2 * ((-(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5) - 0.5) / (-(Cos(dec * Pi / 180))) * (2 * ((-(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5) - 0.5) / (-(Cos(dec * Pi / 180)))))) * 180 / Pi) * Pi / 180))) / 2 + 0.5)
        bx = (Sin(dec * Pi / 180) * Sqr(1 - Sin((Abs(inc) + CSD) * Pi / 180))) / 2 + 0.5
        by = Abs(-(Cos(dec * Pi / 180) * Sqr(1 - Sin((Abs(inc) + CSD) * Pi / 180))) / 2 + 0.5)
        If ay > 1 Then ay = 1 - (ay - 1)
        If by > 1 Then by = 1 - (by - 1)
        End If
        ' Plot of the ellipsoid/circle by small segments (5 degrees)
        For i = 0 To 30
            ' The up ellipsoid/circle is a dash line
            If i = 1 Or i = 3 Or i = 5 Or i = 7 Or i = 9 Or i = 11 Or i = 13 Or i = 15 Or i = 17 Or i = 19 Or i = 21 Or i = 23 Or i = 25 Or i = 27 Or i = 29 Then
            Y1 = (Sqr(((Sin(dec * Pi / 180) * L / L0) / 2 + 0.5 - bx) * ((Sin(dec * Pi / 180) * L / L0) / 2 + 0.5 - bx) + (-(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5 - by) * (-(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5 - by))) * Sin(3 * i * Pi / 180)
            X1 = Sqr((Sqr(((Sin(dec * Pi / 180) * L / L0) / 2 + 0.5 - ax) * ((Sin(dec * Pi / 180) * L / L0) / 2 + 0.5 - ax) + (-(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5 - ay) * (-(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5 - ay))) ^ 2 * (1 - Sin(3 * i * Pi / 180) * Sin(3 * i * Pi / 180)))
            Y2 = (Sqr(((Sin(dec * Pi / 180) * L / L0) / 2 + 0.5 - bx) * ((Sin(dec * Pi / 180) * L / L0) / 2 + 0.5 - bx) + (-(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5 - by) * (-(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5 - by))) * Sin(3 * (i + 1) * Pi / 180)
            X2 = Sqr((Sqr(((Sin(dec * Pi / 180) * L / L0) / 2 + 0.5 - ax) * ((Sin(dec * Pi / 180) * L / L0) / 2 + 0.5 - ax) + (-(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5 - ay) * (-(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5 - ay))) ^ 2 * (1 - Sin(3 * (i + 1) * Pi / 180) * Sin(3 * (i + 1) * Pi / 180)))
            ' Test to don't plot the parts of the ellipsoid/circle which are outside of the plane inc = 0
            If (Abs(-(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5 - X1 * Sin((-dec) * Pi / 180) + Y1 * Cos((-dec) * Pi / 180) - 0.5)) ^ 2 < Abs(0.5 ^ 2 - ((Sin(dec * Pi / 180) * L / L0) / 2 + 0.5 + X1 * Cos((-dec) * Pi / 180) + Y1 * Sin((-dec) * Pi / 180) - 0.5) ^ 2) Then EqualArea.Line ((Sin(dec * Pi / 180) * L / L0) / 2 + 0.5 + X1 * Cos((-dec) * Pi / 180) + Y1 * Sin((-dec) * Pi / 180), -(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5 - X1 * Sin((-dec) * Pi / 180) + Y1 * Cos((-dec) * Pi / 180))-((Sin(dec * Pi / 180) * L / L0) / 2 + 0.5 + X2 * Cos((-dec) * Pi / 180) + Y2 * Sin((-dec) * Pi / 180), -(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5 - X2 * Sin((-dec) * Pi / 180) + Y2 * Cos((-dec) * Pi / 180))
            If (Abs(-(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5 - X1 * Sin((-dec) * Pi / 180) - Y1 * Cos((-dec) * Pi / 180) - 0.5)) ^ 2 < Abs(0.5 ^ 2 - ((Sin(dec * Pi / 180) * L / L0) / 2 + 0.5 + X1 * Cos((-dec) * Pi / 180) - Y1 * Sin((-dec) * Pi / 180) - 0.5) ^ 2) Then EqualArea.Line ((Sin(dec * Pi / 180) * L / L0) / 2 + 0.5 + X1 * Cos((-dec) * Pi / 180) - Y1 * Sin((-dec) * Pi / 180), -(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5 - X1 * Sin((-dec) * Pi / 180) - Y1 * Cos((-dec) * Pi / 180))-((Sin(dec * Pi / 180) * L / L0) / 2 + 0.5 + X2 * Cos((-dec) * Pi / 180) - Y2 * Sin((-dec) * Pi / 180), -(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5 - X2 * Sin((-dec) * Pi / 180) - Y2 * Cos((-dec) * Pi / 180))
            If (Abs(-(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5 + X1 * Sin((-dec) * Pi / 180) + Y1 * Cos((-dec) * Pi / 180) - 0.5)) ^ 2 < Abs(0.5 ^ 2 - ((Sin(dec * Pi / 180) * L / L0) / 2 + 0.5 - X1 * Cos((-dec) * Pi / 180) + Y1 * Sin((-dec) * Pi / 180) - 0.5) ^ 2) Then EqualArea.Line ((Sin(dec * Pi / 180) * L / L0) / 2 + 0.5 - X1 * Cos((-dec) * Pi / 180) + Y1 * Sin((-dec) * Pi / 180), -(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5 + X1 * Sin((-dec) * Pi / 180) + Y1 * Cos((-dec) * Pi / 180))-((Sin(dec * Pi / 180) * L / L0) / 2 + 0.5 - X2 * Cos((-dec) * Pi / 180) + Y2 * Sin((-dec) * Pi / 180), -(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5 + X2 * Sin((-dec) * Pi / 180) + Y2 * Cos((-dec) * Pi / 180))
            If (Abs(-(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5 + X1 * Sin((-dec) * Pi / 180) - Y1 * Cos((-dec) * Pi / 180) - 0.5)) ^ 2 < Abs(0.5 ^ 2 - ((Sin(dec * Pi / 180) * L / L0) / 2 + 0.5 - X1 * Cos((-dec) * Pi / 180) - Y1 * Sin((-dec) * Pi / 180) - 0.5) ^ 2) Then EqualArea.Line ((Sin(dec * Pi / 180) * L / L0) / 2 + 0.5 - X1 * Cos((-dec) * Pi / 180) - Y1 * Sin((-dec) * Pi / 180), -(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5 + X1 * Sin((-dec) * Pi / 180) - Y1 * Cos((-dec) * Pi / 180))-((Sin(dec * Pi / 180) * L / L0) / 2 + 0.5 - X2 * Cos((-dec) * Pi / 180) - Y2 * Sin((-dec) * Pi / 180), -(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5 + X2 * Sin((-dec) * Pi / 180) - Y2 * Cos((-dec) * Pi / 180))
            End If
        Next i
        End If
    End If
End Sub

Private Sub buttonHalt_Click()
    Flow_Pause
    updateFlowStatus
    Unload Me
    If frmProgram.mnuViewMeasurement.Checked Then frmProgram.mnuViewMeasurement.Checked = False
    ' (September 2007 L Carporzen) Logout after Halt
    frmDCMotors.TurningMotorRotate 0, False ' Reset the rod rotation if Halt
    frmDCMotors.HomeToTop ' Reset the rod position if Halt
    If frmVacuum.VacuumConnectOn = True Then
    MsgBox "You are logging out, remove the sample before the vacuum switch off." ' Allows to remove the sample if Halt
    End If
    frmProgram.StatBarNew vbNullString
    Set SampQueue = Nothing
    Set SampleIndexRegistry = Nothing
    Set SampleHolder = Nothing
    Set SusceptibilityStandard = Nothing
    Set MainChanger = Nothing
    frmMagnetometerControl.Hide
    frmSampleIndexRegistry.Hide
    Unload frmMagnetometerControl
    Unload frmSampleIndexRegistry
    frmMagnetometerControl.DisableMagnetCmds
    frmVacuum.ValveConnect False
    frmVacuum.MotorPower False
    DoEvents
    FLAG_MagnetInit = False
    FLAG_MagnetUse = False
    Set SampQueue = New SampleCommands
    Set SampleIndexRegistry = New SampleIndexRegistrations
    Set SampleHolder = SampleIndexRegistry("!Holder").sampleSet("Holder")
    Set SusceptibilityStandard = SampleIndexRegistry("!Holder").sampleSet("SusStd")
    Set MainChanger = New frmChanger
    MainChanger.IsMasterList = True
    Load MainChanger
    frmProgram.SignalReady
    frmMagnetometerControl.cmdManHolder.Enabled = True
    frmMagnetometerControl.cmdManRun.Enabled = True
    frmMagnetometerControl.cmdChangerEdit.Enabled = True
    'frmMagnetometerControl.cmdChangerOK.Enabled = True
    Flow_Halt
    updateFlowStatus
End Sub

Private Sub buttonPause_Click()
    If Prog_paused Then
        Flow_Resume
    Else
        Flow_Pause
    End If
    updateFlowStatus
End Sub

Private Sub ChkAllSteps_Click()
    Actualize
End Sub

Private Sub chkM_Click()
    Actualize
End Sub

Private Sub chkX_Click()
    Actualize
End Sub

Public Sub clearData()
    ' This function clears all the currently displayed data.
    Dim i As Integer
    cmdPrint.Enabled = True
    lblMeasXZero(0).Caption = vbNullString
    lblMeasYZero(0).Caption = vbNullString
    lblMeasZZero(0).Caption = vbNullString
    lblMeasXZero(1).Caption = vbNullString
    lblMeasYZero(1).Caption = vbNullString
    lblMeasZZero(1).Caption = vbNullString
    For i = 0 To 3
        lblMeasX(i).Caption = vbNullString
        lblMeasY(i).Caption = vbNullString
        lblMeasZ(i).Caption = vbNullString
        lblCalcInc(i).Caption = vbNullString
        lblCalcDec(i).Caption = vbNullString
    Next i
End Sub

Public Sub clearStats()
    ' This procedure clears data from the stat data boxes
    frmStats.cmdPrint.Enabled = True
    lblavgx.Caption = vbNullString
    lblavgy.Caption = vbNullString
    lblavgz.Caption = vbNullString
    lblDSigDrift.Caption = vbNullString
    lblDSigHolder.Caption = vbNullString
    lblDSigInduced.Caption = vbNullString
    lblavgmag.Caption = vbNullString
    lblCSD.Caption = vbNullString
End Sub

Private Sub cmdHide_Click()
    ' Update the main form and hide the window.
    frmProgram.mnuViewMeasurement.Checked = False
    Me.Hide
End Sub

Private Sub cmdHideEqu_Click()
    cmdShowPlots.Enabled = True
    If Me.Height = 5350 Then
        Me.Height = 5350
    Else
        Me.Height = 7440
        cmdHideZij.Enabled = False
    End If
    Me.Width = 5400
End Sub

Private Sub cmdHideZij_Click()
    MomentX.Visible = False
    Me.Height = 7440
    frmMeasure.framJumps.Top = 5040
    frmMeasure.framJumps.Left = 5400
    ChkM.Visible = False
    ChkX.Visible = False
    cmdHideZij.Enabled = False
    cmdShowPlots.Enabled = True
    Me.Width = 10490
End Sub

Private Sub cmdPrint_Click()
    On Error GoTo ErrorHandler
    Dim X As Integer
    Dim numlines As Integer    ' Number of lines printed here
    numlines = 21
    cmdPrint.Enabled = False
    If Print_LinesLeft < numlines Then
        Print_PageBreak
    End If
    Print_Line
    Print_Line "Measurement:"
    Print_Line "Sample - " & lblSampName.Caption
    Print_Line
    Print_Line vbTab & vbTab & _
               "XZero" & vbTab & vbTab & _
               "YZero" & vbTab & vbTab & _
               "ZZero"
    Print_Line "First:" & vbTab & vbTab & _
               lblMeasXZero(0).Caption & vbTab & _
               lblMeasYZero(0).Caption & vbTab & _
               lblMeasZZero(0).Caption
    Print_Line "Second:" & vbTab & vbTab & _
               lblMeasXZero(1).Caption & vbTab & _
               lblMeasYZero(1).Caption & vbTab & _
               lblMeasZZero(1).Caption
    Print_Line
    Print_Line vbTab & vbTab & _
               "X" & vbTab & vbTab & _
               "Y" & vbTab & vbTab & _
               "Z" & vbTab & vbTab & _
               "Dec:" & vbTab & vbTab & _
               "Inc:"
    For X = 0 To 3
        Print_Line vbTab & vbTab & _
                   lblMeasX(X).Caption & vbTab & _
                   lblMeasY(X).Caption & vbTab & _
                   lblMeasZ(X).Caption & vbTab & _
                   lblCalcDec(X).Caption & vbTab & vbTab & _
                   lblCalcInc(X).Caption
    Next X
    Print_Line
    Print_Line vbTab & vbTab & _
               "Avg. X" & vbTab & vbTab & _
               "Avg. Y" & vbTab & vbTab & _
               "Avg. Z" & vbTab & vbTab
    Print_Line vbTab & vbTab & _
               lblavgx.Caption & vbTab & _
               lblavgy.Caption & vbTab & _
               lblavgz.Caption & vbTab
    Print_Line
    Print_Line "Signal / Drift:  " & vbTab & lblDSigDrift.Caption
    Print_Line "Signal / Holder: " & vbTab & lblDSigHolder.Caption
    Print_Line "Signal / Induced:" & vbTab & lblDSigInduced.Caption
    Print_Line
    Exit Sub
ErrorHandler:
    MsgBox ("There was a problem printing to the printer.")
End Sub

Private Sub cmdShowPlots_click()
    If lblSampName = "" Then Exit Sub
    cmdShowPlots.Enabled = False
    If cmdHideZij.Enabled = False Then
        If lblSampName = "Holder" Then Exit Sub
        cmdHideZij.Enabled = True
        Me.Width = 15480
        If MomentX.Visible = True Then Me.Height = 9060
        Actualize
    Else
        Me.Width = 10490
    End If
End Sub

Private Sub cmdStats_Click()
    cmdStats.Enabled = False
    frmStats.ZOrder
    frmStats.Show
End Sub

Private Sub Form_Activate()
    If Me.Visible Then
        frmProgram.mnuViewMeasurement.Checked = True
    Else
        ' would this ever happen?
        frmProgram.mnuViewMeasurement.Checked = False
    End If
End Sub

Private Sub Form_Load()

    '(March 11, 2011 - I Hilburn)
    'Update all the form's icons with the program Ico file
    SetFormIcon Me

    lblDataFileName.Caption = vbNullString     ' Show file name
    frmMeasure.lblDataFileName = vbNullString
    HideStats
    updateFlowStatus
    
End Sub

Public Function getMeasDir() As Integer
    ' This function returns the current direction of the
    ' arrow on the sample that is being measured. (up/down)
    If lblMeasDir.Caption = "U" Then
        getMeasDir = Magnet_SampleOrientationUp
    Else
        getMeasDir = Magnet_SampleOrientationDown
    End If
End Function

Public Function GetSample() As String
    ' This function returns the name of the current sample
    ' that is being measured.  It assumes we have already
    ' told the form which sample we're measuring at some
    ' previous time.
    GetSample = lblSampName.Caption
End Function

Public Sub HideStats()
    ' This procedure hides the statistics frame and resizes
    ' the measurement window
    
    'Add in a three second pause here
    PauseTill timeGetTime + 3000
    
    frmStats.Hide
    framStats.Visible = False
    Me.Height = 5350
    Me.Width = 5400
End Sub

Public Sub ImportZijRoutine(ByVal FilePath As String, crdec As Double, crinc As Double, momentvol As Double, refresh As Boolean)
    ' (August 2007 L Carporzen) Zijderveld diagram near the measurement window & the equal area plot
    Dim filenum As Integer
    Dim whole_file As String ' To read the sample file for the Zijderveld diagram
    Dim lines As Variant
    Dim num_rows As Long
    Dim ZijLines As Long
    Dim r As Long
    Dim dec As Double
    Dim inc As Double
    Dim readMoment As Double
    Dim readcrdec As Double
    Dim readcrinc As Double
    Dim readcrdec2 As Double ' (June 2008 L Carporzen) Link the previous directions
    Dim readcrinc2 As Double ' (June 2008 L Carporzen) Link the previous directions
    Dim MaxZijX As Double
    Dim MaxZijY As Double
    Dim MinZijX As Double
    Dim MinZijY As Double
    Dim ZijScale As Double
    Dim ZijHoriOrig As Double
    Dim ZijVertOrig As Double
    Dim ZijX As Variant
    Dim ZijY As Variant
    Dim ZijZ As Variant
    Dim RMGlines As Variant ' To read the RMG file for the susceptibility versus demagnetization
    Dim RMGarray As Variant
    Dim numRMGrows As Long
    Dim SusceLines As Long
    Dim p As Long
    Dim Q As Long
    Dim MaxMoment As Double
    Dim MaxDemag As Double
    Dim MaxSusceptibility As Double
    Dim MinSusceptibility As Double
    Dim SusceScale As Double
    Dim SusceOrig As Double
    Dim DemagStep As Variant
    Dim Susceptibility As Variant
    Dim AF As Boolean
    Dim Thermal As Boolean
    Dim samplepath As String
    For r = 0 To frmSampleIndexRegistry.cmbSampCode.ListCount - 1
    If LenB(dir$(frmSampleIndexRegistry.txtDir & frmSampleIndexRegistry.cmbSampCode.List(r) & "\" & FilePath)) > 0 Then
        If FilePath = "" Then Exit Sub
        If lblDataFileName = frmSampleIndexRegistry.txtDir & frmSampleIndexRegistry.cmbSampCode.List(r) & "\" & frmSampleIndexRegistry.cmbSampCode.List(r) & ".sam" Then
            samplepath = frmSampleIndexRegistry.txtDir & frmSampleIndexRegistry.cmbSampCode.List(r) & "\" & FilePath
        End If
        If cmdShowPlots.Enabled = False Then
            ShowZijderveld
        Else
            Exit Sub
        End If
    End If
    Next r
    If Not LenB(samplepath) > 0 Then
        Exit Sub
    Else
        If FilePath = "" Then Exit Sub
        If cmdShowPlots.Enabled = False Then
            ShowZijderveld
        Else
            Exit Sub
        End If
    End If
    If Not LenB(dir$(samplepath & ".rmg")) > 0 Then
        ChkX.Visible = False
    Else
        ChkX.Visible = True ' Allow reading the RMG file for susceptibility versus demagnetization
    End If
    If txtZijLines = "" Then txtZijLines = 0
    If txtZijLines < 0 Then txtZijLines = 0
    ZijLines = txtZijLines ' Nb of previous steps plot for the comparison
    Zijderveld.Cls ' Clean the plot
    ChkM.Visible = True
    MomentX.Cls ' Clean the plot
    filenum = FreeFile ' Read the sample file
    Open samplepath For Input As #filenum
    whole_file = Input$(LOF(filenum), #filenum)
    Close #filenum
    lines = Split(whole_file, vbCrLf) ' Cut the file in lines
    whole_file = ""
    num_rows = UBound(lines)
    If num_rows < ZijLines + 2 Then ZijLines = num_rows - 2
    If ZijLines < 1 Then Exit Sub
    ReDim ZijX(ZijLines)
    ReDim ZijY(ZijLines)
    ReDim ZijZ(ZijLines)
    MaxZijX = 0.0000000001
    MaxZijY = 0.0000000001
    MinZijX = -0.0000000001
    MinZijY = -0.0000000001
    p = 1
    MaxMoment = 0.0000000001
    MaxDemag = 0
    MaxSusceptibility = 0.00001
    MinSusceptibility = 0
    If Abs(val(Right$(lblDemag.Caption, 4))) = 0 And Not Left$(lblDemag.Caption, 3) = "ARM" Then
        MomentX.Visible = False
        Me.Height = 7440
        frmMeasure.framJumps.Top = 5040
        frmMeasure.framJumps.Left = 5400
        ChkM.Visible = False
        ChkX.Visible = False
    Else
        MomentX.Visible = True
        Me.Height = 9060
        frmMeasure.framJumps.Top = 6720
        frmMeasure.framJumps.Left = 10440
    End If
    If Left$(lblDemag.Caption, 3) = "ARM" Then ChkAllSteps.Value = Checked
    If ChkM.Value = Unchecked And ChkX.Value = Unchecked Then
        MomentX.Visible = False
        Me.Height = 7440
        frmMeasure.framJumps.Top = 5040
        frmMeasure.framJumps.Left = 5400
    End If
    If MomentX.Visible = True And ChkX.Visible = True Then ' (October 2007 L Carporzen) Susceptibility versus demagnetization below the equal area plot
        filenum = FreeFile ' Read the RMG file
        Open samplepath & ".rmg" For Input As #filenum
        whole_file = Input$(LOF(filenum), #filenum)
        Close #filenum
        RMGlines = Split(whole_file, vbCrLf)
        whole_file = ""
        numRMGrows = UBound(RMGlines)
        ReDim RMGarray(3 * ZijLines)
        If numRMGrows > 3 * ZijLines Then
            SusceLines = 3 * ZijLines
        Else
            SusceLines = numRMGrows
        End If
        For r = 1 To SusceLines
            If RMGlines(numRMGrows - r) = "" Then Exit Sub
            RMGarray(r) = Split(RMGlines(numRMGrows - r), ",")
        Next r
        ReDim Susceptibility(ZijLines)
    End If
    ReDim DemagStep(ZijLines)
    For r = 1 To ZijLines
        readMoment = val(Mid$(lines(num_rows - r), 32, 8))
        If optCore.Value = True Then ' Core coordinates
            readcrdec = val(Mid$(lines(num_rows - r), 47, 5))
            readcrinc = val(Mid$(lines(num_rows - r), 53, 5))
            If r > 1 Then ' (June 2008 L Carporzen) Link the previous directions
            readcrdec2 = val(Mid$(lines(num_rows - r + 1), 47, 5))
            readcrinc2 = val(Mid$(lines(num_rows - r + 1), 53, 5))
            Else
            readcrdec2 = frmStats.lblCDec.Caption
            readcrinc2 = frmStats.lblCInc.Caption
            End If
        End If
        If optGeographic.Value = True Then ' Geographic coordinates
            readcrdec = val(Mid$(lines(num_rows - r), 8, 5))
            readcrinc = val(Mid$(lines(num_rows - r), 14, 5))
            If r > 1 Then ' (June 2008 L Carporzen) Link the previous directions
            readcrdec2 = val(Mid$(lines(num_rows - r + 1), 8, 5))
            readcrinc2 = val(Mid$(lines(num_rows - r + 1), 14, 5))
            Else
            readcrdec2 = frmStats.lblGDec.Caption
            readcrinc2 = frmStats.lblGInc.Caption
            End If
        End If
        If optBedding.Value = True Then ' Bedding coordinates
            readcrdec = val(Mid$(lines(num_rows - r), 20, 5))
            readcrinc = val(Mid$(lines(num_rows - r), 26, 5))
            If r > 1 Then ' (June 2008 L Carporzen) Link the previous directions
            readcrdec2 = val(Mid$(lines(num_rows - r + 1), 20, 5))
            readcrinc2 = val(Mid$(lines(num_rows - r + 1), 26, 5))
            Else
            readcrdec2 = frmStats.lblBDec.Caption
            readcrinc2 = frmStats.lblBInc.Caption
            End If
        End If
        If refresh = True Then
            lblOrientation.Caption = "Geographic and bedding coordinates are read in the sample file, it could be not valid if you changed the orientation parameters since the previous measurements."
            lblOrientation.Visible = True
            PlotHistory readcrdec, readcrinc, readcrdec2, readcrinc2 ' (June 2008 L Carporzen) Link the previous directions
        Else
            lblOrientation.Visible = False
            optCore.Value = True
            optGeographic.Value = False
            optBedding.Value = False
            readcrdec = val(Mid$(lines(num_rows - r), 47, 5))
            readcrinc = val(Mid$(lines(num_rows - r), 53, 5))
            If r > 1 Then ' (June 2008 L Carporzen) Link the previous directions
            readcrdec2 = val(Mid$(lines(num_rows - r + 1), 20, 5))
            readcrinc2 = val(Mid$(lines(num_rows - r + 1), 26, 5))
            Else
            readcrdec2 = frmStats.lblBDec.Caption
            readcrinc2 = frmStats.lblBInc.Caption
            End If
            PlotHistory val(Mid$(lines(num_rows - r), 20, 5)), val(Mid$(lines(num_rows - r), 26, 5)), readcrdec2, readcrinc2 ' (June 2008 L Carporzen) Link the previous directions
        End If
        ZijX(r) = readMoment * Cos(readcrinc * Pi / 180) * Cos(readcrdec * Pi / 180)
        ZijY(r) = readMoment * Cos(readcrinc * Pi / 180) * Sin(readcrdec * Pi / 180)
        ZijZ(r) = readMoment * Sin(readcrinc * Pi / 180)
        If -ZijX(r) > MaxZijX Then MaxZijX = -ZijX(r)
        If ZijY(r) > MaxZijY Then MaxZijY = ZijY(r)
        If ZijZ(r) > MaxZijX Then MaxZijX = ZijZ(r)
        If -ZijX(r) < MinZijX Then MinZijX = -ZijX(r)
        If ZijY(r) < MinZijY Then MinZijY = ZijY(r)
        If ZijZ(r) < MinZijX Then MinZijX = ZijZ(r)
        If MomentX.Visible = True Then
          DemagStep(r) = Abs(val(Mid$(lines(num_rows - r), 4, 3)))
          If ChkX.Visible = True Then
            For Q = p To SusceLines
                If Left$(lines(num_rows - r), 2) = RMGarray(Q)(0) Then ' Are the 2 letters of the step labels are equals (AF or TT)
                    If RMGarray(Q)(0) = "TT" Then
                        Thermal = True
                    Else
                        Thermal = False
                    End If
                    If RMGarray(Q)(0) = "AF" Then
                        AF = True
                    Else
                        AF = False
                    End If
                    If DemagStep(r) = Abs(val(Right$(RMGarray(Q)(1), 3))) Then  ' Are the 3 last digits of the step numbers are equals
                        DemagStep(r) = Abs(val(RMGarray(Q)(1)))
                        Susceptibility(r) = val(RMGarray(Q)(8))
                        p = Q + 1
                        Exit For
                    ElseIf DemagStep(r) = Abs(val(Right$(RMGarray(Q)(1), 4))) Then ' Are the 4 last digits of the step numbers are equals
                        DemagStep(r) = Abs(val(RMGarray(Q)(1)))
                        Susceptibility(r) = val(RMGarray(Q)(8))
                        p = Q + 1
                        Exit For
                    End If
                ElseIf Left$(lines(num_rows - r), 3) = RMGarray(Q)(0) Then ' Are the 3 letters of the step labels are equals
                    If RMGarray(Q)(0) = "IRM" Or RMGarray(Q)(0) = "ARM" Or RMGarray(Q)(0) = "AFz" Then
                        AF = True
                    Else
                        AF = False
                    End If
                    If DemagStep(r) = Abs(val(Right$(RMGarray(Q)(1), 3))) Then ' Are the 3 last digits of the step numbers are equals
                        DemagStep(r) = Abs(val(RMGarray(Q)(1)))
                        Susceptibility(r) = val(RMGarray(Q)(8))
                        p = Q + 1
                        Exit For
                    ElseIf DemagStep(r) = Abs(val(Right$(RMGarray(Q)(1), 4))) Then ' Are the 4 last digits of the step numbers are equals
                        DemagStep(r) = Abs(val(RMGarray(Q)(1)))
                        Susceptibility(r) = val(RMGarray(Q)(8))
                        p = Q + 1
                        Exit For
                    ElseIf Left$(lines(num_rows - r), 3) = "ARM" Then ' Is their only a real number only in the RMG (ARM)
                        DemagStep(r) = Abs(val(RMGarray(Q)(1)))
                        Susceptibility(r) = val(RMGarray(Q)(8))
                        p = Q + 1
                        Exit For
                    End If
                ElseIf Left$(lines(num_rows - r), 5) = RMGarray(Q)(0) Then ' Are the step labels are equals to AFmax
                    If RMGarray(Q)(0) = "AFmax" Then
                        AF = True
                    Else
                        AF = False
                    End If
                    DemagStep(r) = Abs(val(RMGarray(Q)(1)))
                    Susceptibility(r) = val(RMGarray(Q)(8))
                    p = Q + 1
                    Exit For
                End If
            Next Q
            If Susceptibility(r) = "" Then Susceptibility(r) = 0
          End If
          If ChkAllSteps.Value = Checked Or DemagStep(r) <= Abs(val(Right$(lblDemag.Caption, 4))) Then ' We don't want to plot the previous steps with higher demag numbers
            If readMoment > MaxMoment Then MaxMoment = readMoment
            If DemagStep(r) > MaxDemag Then MaxDemag = DemagStep(r)
            If ChkX.Visible = True And ChkX.Value = Checked Then
              If Susceptibility(r) > MaxSusceptibility Then MaxSusceptibility = Susceptibility(r)
              If Susceptibility(r) < MinSusceptibility Then MinSusceptibility = Susceptibility(r)
            End If
          End If
        End If
    Next r
    If -momentvol * Cos(crinc * Pi / 180) * Cos(crdec * Pi / 180) > MaxZijX Then MaxZijX = -momentvol * Cos(crinc * Pi / 180) * Cos(crdec * Pi / 180)
    If momentvol * Cos(crinc * Pi / 180) * Sin(crdec * Pi / 180) > MaxZijY Then MaxZijY = momentvol * Cos(crinc * Pi / 180) * Sin(crdec * Pi / 180)
    If momentvol * Sin(crinc * Pi / 180) > MaxZijX Then MaxZijX = momentvol * Sin(crinc * Pi / 180)
    If -momentvol * Cos(crinc * Pi / 180) * Cos(crdec * Pi / 180) < MinZijX Then MinZijX = -momentvol * Cos(crinc * Pi / 180) * Cos(crdec * Pi / 180)
    If momentvol * Cos(crinc * Pi / 180) * Sin(crdec * Pi / 180) < MinZijY Then MinZijY = momentvol * Cos(crinc * Pi / 180) * Sin(crdec * Pi / 180)
    If momentvol * Sin(crinc * Pi / 180) < MinZijX Then MinZijX = momentvol * Sin(crinc * Pi / 180)
    If MomentX.Visible = True Then
      If Abs(val(Right$(lblDemag.Caption, 4))) > MaxDemag Then MaxDemag = Abs(val(Right$(lblDemag.Caption, 4)))
      If ChkM.Value = Checked Or ChkX.Value = Checked Then
        MomentX.Line (0, 1)-(1, 1) ' horizontal axis
        MomentX.Line (1, 1 - 0.02)-(1, 1 + 0.02)
        MomentX.CurrentX = 0
        MomentX.CurrentY = 1.02
        MomentX.Print "0"
        MomentX.Line (0, 1 - 0.02)-(0, 1 + 0.02)
        MomentX.CurrentX = 1 - 0.06
        MomentX.CurrentY = 1.02
        MomentX.Print MaxDemag
        If Thermal = True Then
          MomentX.CurrentX = 0.5
          MomentX.CurrentY = 1.01
          MomentX.Print "°C"
        ElseIf AF = True Then
          MomentX.CurrentX = 0.5
          MomentX.CurrentY = 1.01
          MomentX.Print "Oe"
        Else
          MomentX.CurrentX = 0.4
          MomentX.CurrentY = 1.01
          MomentX.Print "Oe & °C"
        End If
      End If
      If momentvol > MaxMoment Then MaxMoment = momentvol
      If MaxMoment = 0.0000000001 Or MaxDemag = 0 Then
        ChkM.Visible = False
      ElseIf ChkM.Visible = True And ChkM.Value = Checked Then
        MomentX.Line (0, 0)-(0, 1), RGB(255, 0, 0) ' vertical axis
        MomentX.Line (-0.02, 1)-(0.02, 1), RGB(255, 0, 0)
        MomentX.Line (-0.02, 0)-(0.02, 0), RGB(255, 0, 0)
        MomentX.ForeColor = RGB(255, 0, 0)
        MomentX.CurrentX = -0.15
        MomentX.CurrentY = 0.5
        MomentX.Print "emu"
        MomentX.CurrentX = -0.15
        MomentX.CurrentY = -0.07
        MomentX.Print Format$(MaxMoment, "0.00E+")
        MomentX.CurrentX = -0.05
        MomentX.CurrentY = 1 - 0.05
        MomentX.Print "0"
        MomentX.ForeColor = RGB(0, 0, 0)
      End If
      If ChkX.Visible = True Or ChkX.Value = Checked Then
        If COMPortSusceptibility > 0 And EnableSusceptibility And Not val(frmSusceptibilityMeter.InputText) = -1 Then
            If val(frmSusceptibilityMeter.InputText) * SusceptibilityMomentFactorCGS > MaxSusceptibility Then MaxSusceptibility = val(frmSusceptibilityMeter.InputText) * SusceptibilityMomentFactorCGS
            If val(frmSusceptibilityMeter.InputText) * SusceptibilityMomentFactorCGS < MinSusceptibility Then MinSusceptibility = val(frmSusceptibilityMeter.InputText) * SusceptibilityMomentFactorCGS
        End If
        If MaxSusceptibility = 0.00001 And MinSusceptibility = 0 Or MaxDemag = 0 Then
            If ChkM.Value = Checked Then ChkX.Visible = False
            If ChkM.Visible = False Or ChkM.Value = Unchecked Then
                MomentX.Visible = False
                Me.Height = 7440
                frmMeasure.framJumps.Top = 5040
                frmMeasure.framJumps.Left = 5400
                If ChkM.Value = Checked Then ChkM.Visible = False
            End If
        End If
      End If
    End If
    If MomentX.Visible = True And ChkX.Visible = True And ChkX.Value = Checked Then
        SusceScale = Abs(MaxSusceptibility - MinSusceptibility)
        SusceOrig = Abs(MinSusceptibility / SusceScale)
        If ChkM.Visible = True And ChkM.Value = Checked Then
            MomentX.Line (0, 0)-(0, 1) ' vertical axis
            MomentX.Line (-0.02, 1)-(0.02, 1)
            MomentX.Line (-0.02, 0)-(0.02, 0)
        Else
            MomentX.Line (0, 0)-(0, 1), RGB(0, 0, 255) ' vertical axis
            MomentX.Line (-0.02, 1)-(0.02, 1), RGB(0, 0, 255)
            MomentX.Line (-0.02, 0)-(0.02, 0), RGB(0, 0, 255)
        End If
        MomentX.Line (0, 1 - SusceOrig)-(1, 1 - SusceOrig), RGB(0, 0, 255) ' horizontal axis
        MomentX.Line (1, 1 - SusceOrig - 0.02)-(1, 1 - SusceOrig + 0.02), RGB(0, 0, 255)
        MomentX.ForeColor = RGB(0, 0, 255)
        MomentX.CurrentX = -0.05
        MomentX.CurrentY = 1 - SusceOrig - 0.05
        MomentX.Print "0"
        MomentX.CurrentX = -0.15
        If 1 - SusceOrig > 0.5 Then
            MomentX.CurrentY = 1 - SusceOrig - 0.05 - 0.05
        Else
            MomentX.CurrentY = 1 - SusceOrig + 0.05 - 0.05
        End If
        MomentX.Print "emu/Oe"
        MomentX.CurrentX = -0.15
        MomentX.CurrentY = 0.01
        MomentX.Print Format$(MaxSusceptibility, "0.00E+")
        If Not MinSusceptibility = 0 Then
            MomentX.CurrentX = -0.15
            MomentX.CurrentY = 1.01
            MomentX.Print Format$(MinSusceptibility, "0.00E+")
        End If
        MomentX.ForeColor = RGB(0, 0, 0)
    End If
    If MaxZijX = MinZijX Or MaxZijY = MinZijY Then
        ' Do nothing, avoid bugs???
    Else ' We can plot
        ZijScale = Abs(MaxZijX - MinZijX)
        MaxZijX = MaxZijX + 0.05 * ZijScale
        MinZijX = MinZijX - 0.05 * ZijScale
        ZijScale = Abs(MaxZijY - MinZijY)
        MaxZijY = MaxZijY + 0.05 * ZijScale
        MinZijY = MinZijY - 0.05 * ZijScale
        If Abs(MaxZijX - MinZijX) > Abs(MaxZijY - MinZijY) Then
            ZijScale = Abs(MaxZijX - MinZijX) ' Same scale for both axis
            ZijHoriOrig = Abs(MinZijY / ZijScale) + (1 - Abs(MaxZijY - MinZijY) / ZijScale) / 2 'Center the plot in the page
            ZijVertOrig = Abs(MinZijX / ZijScale) 'The lowest and highest values are on the borders of the plot
        Else
            ZijScale = Abs(MaxZijY - MinZijY) ' Same scale for both axis
            ZijHoriOrig = Abs(MinZijY / ZijScale) 'The lowest and highest values are on the borders of the plot
            ZijVertOrig = Abs(MinZijX / ZijScale) + (1 - Abs(MaxZijX - MinZijX) / ZijScale) / 2 'Center the plot in the page
        End If
        Zijderveld.Line (ZijHoriOrig, 0)-(ZijHoriOrig, 1) ' vertical axis
        Zijderveld.Line (0, ZijVertOrig)-(1, ZijVertOrig) ' horizontal axis
        Zijderveld.CurrentX = 0
        Zijderveld.CurrentY = ZijVertOrig
        Zijderveld.Print "W"
        Zijderveld.CurrentX = 1 - 0.025
        Zijderveld.CurrentY = ZijVertOrig
        Zijderveld.Print "E"
        Zijderveld.CurrentX = ZijHoriOrig - 0.03
        Zijderveld.CurrentY = 0
        Zijderveld.Print "N  Up"
        Zijderveld.CurrentX = ZijHoriOrig - 0.03
        Zijderveld.CurrentY = 1 - 0.04
        Zijderveld.Print "S  Down"
        Zijderveld.Circle (ZijHoriOrig - 0.06, 0.02), 0.02, RGB(0, 0, 255)
        Zijderveld.Circle (ZijHoriOrig + 0.08, 0.02), 0.02, RGB(255, 0, 0)
        If ZijHoriOrig >= 0.5 Then
            Zijderveld.CurrentX = 0.04
        Else
            Zijderveld.CurrentX = 0.7
        End If
        Zijderveld.CurrentY = 0
        Zijderveld.Print "View: " & Format$(ZijScale, "0.00E+") & " emu"   ' scale
        If ZijHoriOrig >= 0.5 Then
            Zijderveld.CurrentX = 0.04
        Else
            Zijderveld.CurrentX = 0.7
        End If
        Zijderveld.CurrentY = 1 - 0.04
        Zijderveld.Print "Last: " & Format$(momentvol, "0.00E+") & " emu"   ' last magnetization
        For r = 1 To ZijLines 'Circles for each step
            lines(num_rows - r) = ""
            If MomentX.Visible = True Then
                If ChkAllSteps.Value = Checked Or DemagStep(r) <= Abs(val(Right$(lblDemag.Caption, 4))) Then
                    If ChkM.Visible = True And ChkM.Value = Checked And Not Sqr(ZijX(r) ^ 2 + ZijY(r) ^ 2 + ZijZ(r) ^ 2) = 0 Then MomentX.Circle (DemagStep(r) / MaxDemag, 1 - Sqr(ZijX(r) ^ 2 + ZijY(r) ^ 2 + ZijZ(r) ^ 2) / MaxMoment), 0.005, RGB(255, 0, 0)
                    If ChkX.Visible = True And ChkX.Value = Checked Then
                        If Not Susceptibility(r) = 0 Then MomentX.Circle (DemagStep(r) / MaxDemag, 1 - (Susceptibility(r) / SusceScale + SusceOrig)), 0.005, RGB(0, 0, 255)
                    End If
                End If
            End If
            Zijderveld.Circle (ZijY(r) / ZijScale + ZijHoriOrig, -ZijX(r) / ZijScale + ZijVertOrig), 0.005, RGB(0, 0, 255)
            Zijderveld.Circle (ZijY(r) / ZijScale + ZijHoriOrig, ZijZ(r) / ZijScale + ZijVertOrig), 0.005, RGB(255, 0, 0)
        Next r
        For r = 1 To ZijLines - 1 'Link each step by a line
            If MomentX.Visible = True Then
                If ChkAllSteps.Value = Checked Or DemagStep(r) <= Abs(val(Right$(lblDemag.Caption, 4))) And DemagStep(r + 1) <= Abs(val(Right$(lblDemag.Caption, 4))) Then
                    If ChkM.Visible = True And ChkM.Value = Checked And Not (Sqr(ZijX(r) ^ 2 + ZijY(r) ^ 2 + ZijZ(r) ^ 2) = 0 Or Sqr(ZijX(r + 1) ^ 2 + ZijY(r + 1) ^ 2 + ZijZ(r + 1) ^ 2) = 0) Then MomentX.Line (DemagStep(r) / MaxDemag, 1 - Sqr(ZijX(r) ^ 2 + ZijY(r) ^ 2 + ZijZ(r) ^ 2) / MaxMoment)-(DemagStep(r + 1) / MaxDemag, 1 - Sqr(ZijX(r + 1) ^ 2 + ZijY(r + 1) ^ 2 + ZijZ(r + 1) ^ 2) / MaxMoment), RGB(255, 0, 0)
                    If ChkX.Visible = True And ChkX.Value = Checked Then
                        If Not (Susceptibility(r) = 0 Or Susceptibility(r + 1) = 0) Then MomentX.Line (DemagStep(r) / MaxDemag, 1 - (Susceptibility(r) / SusceScale + SusceOrig))-(DemagStep(r + 1) / MaxDemag, 1 - (Susceptibility(r + 1) / SusceScale + SusceOrig)), RGB(0, 0, 255)
                    End If
                End If
            End If
            Zijderveld.Line (ZijY(r) / ZijScale + ZijHoriOrig, -ZijX(r) / ZijScale + ZijVertOrig)-(ZijY(r + 1) / ZijScale + ZijHoriOrig, -ZijX(r + 1) / ZijScale + ZijVertOrig), RGB(0, 0, 255)
            Zijderveld.Line (ZijY(r) / ZijScale + ZijHoriOrig, ZijZ(r) / ZijScale + ZijVertOrig)-(ZijY(r + 1) / ZijScale + ZijHoriOrig, ZijZ(r + 1) / ZijScale + ZijVertOrig), RGB(255, 0, 0)
        Next r
        ' Design a cross for the last step
        Zijderveld.Line (momentvol * Cos(crinc * Pi / 180) * Sin(crdec * Pi / 180) / ZijScale + ZijHoriOrig - 0.015, -momentvol * Cos(crinc * Pi / 180) * Cos(crdec * Pi / 180) / ZijScale + ZijVertOrig - 0.015)-(momentvol * Cos(crinc * Pi / 180) * Sin(crdec * Pi / 180) / ZijScale + ZijHoriOrig + 0.015, -momentvol * Cos(crinc * Pi / 180) * Cos(crdec * Pi / 180) / ZijScale + ZijVertOrig + 0.015), RGB(0, 0, 255)
        Zijderveld.Line (momentvol * Cos(crinc * Pi / 180) * Sin(crdec * Pi / 180) / ZijScale + ZijHoriOrig - 0.015, -momentvol * Cos(crinc * Pi / 180) * Cos(crdec * Pi / 180) / ZijScale + ZijVertOrig + 0.015)-(momentvol * Cos(crinc * Pi / 180) * Sin(crdec * Pi / 180) / ZijScale + ZijHoriOrig + 0.015, -momentvol * Cos(crinc * Pi / 180) * Cos(crdec * Pi / 180) / ZijScale + ZijVertOrig - 0.015), RGB(0, 0, 255)
        Zijderveld.Line (momentvol * Cos(crinc * Pi / 180) * Sin(crdec * Pi / 180) / ZijScale + ZijHoriOrig - 0.015, momentvol * Sin(crinc * Pi / 180) / ZijScale + ZijVertOrig - 0.015)-(momentvol * Cos(crinc * Pi / 180) * Sin(crdec * Pi / 180) / ZijScale + ZijHoriOrig + 0.015, momentvol * Sin(crinc * Pi / 180) / ZijScale + ZijVertOrig + 0.015), RGB(255, 0, 0)
        Zijderveld.Line (momentvol * Cos(crinc * Pi / 180) * Sin(crdec * Pi / 180) / ZijScale + ZijHoriOrig - 0.015, momentvol * Sin(crinc * Pi / 180) / ZijScale + ZijVertOrig + 0.015)-(momentvol * Cos(crinc * Pi / 180) * Sin(crdec * Pi / 180) / ZijScale + ZijHoriOrig + 0.015, momentvol * Sin(crinc * Pi / 180) / ZijScale + ZijVertOrig - 0.015), RGB(255, 0, 0)
        ' Design a big circle around the cross for the last step
        Zijderveld.Circle (momentvol * Cos(crinc * Pi / 180) * Sin(crdec * Pi / 180) / ZijScale + ZijHoriOrig, -momentvol * Cos(crinc * Pi / 180) * Cos(crdec * Pi / 180) / ZijScale + ZijVertOrig), 0.02, RGB(0, 0, 255)
        Zijderveld.Circle (momentvol * Cos(crinc * Pi / 180) * Sin(crdec * Pi / 180) / ZijScale + ZijHoriOrig, momentvol * Sin(crinc * Pi / 180) / ZijScale + ZijVertOrig), 0.02, RGB(255, 0, 0)
        If MomentX.Visible = True Then
            If ChkM.Visible = True And ChkM.Value = Checked Then
                MomentX.Line (Abs(val(Right$(lblDemag.Caption, 4))) / MaxDemag - 0.015, 1 - momentvol / MaxMoment - 0.015)-(Abs(val(Right$(lblDemag.Caption, 4))) / MaxDemag + 0.015, 1 - momentvol / MaxMoment + 0.015), RGB(255, 0, 0)
                MomentX.Line (Abs(val(Right$(lblDemag.Caption, 4))) / MaxDemag - 0.015, 1 - momentvol / MaxMoment + 0.015)-(Abs(val(Right$(lblDemag.Caption, 4))) / MaxDemag + 0.015, 1 - momentvol / MaxMoment - 0.015), RGB(255, 0, 0)
                MomentX.Circle (Abs(val(Right$(lblDemag.Caption, 4))) / MaxDemag, 1 - momentvol / MaxMoment), 0.02, RGB(255, 0, 0)
            End If
            If ChkX.Visible = True And ChkX.Value = Checked Then
                MomentX.Line (Abs(val(Right$(lblDemag.Caption, 4))) / MaxDemag - 0.015, 1 - (val(frmSusceptibilityMeter.InputText) * SusceptibilityMomentFactorCGS / SusceScale + SusceOrig) - 0.015)-(Abs(val(Right$(lblDemag.Caption, 4))) / MaxDemag + 0.015, 1 - (val(frmSusceptibilityMeter.InputText) * SusceptibilityMomentFactorCGS / SusceScale + SusceOrig) + 0.015), RGB(0, 0, 255)
                MomentX.Line (Abs(val(Right$(lblDemag.Caption, 4))) / MaxDemag - 0.015, 1 - (val(frmSusceptibilityMeter.InputText) * SusceptibilityMomentFactorCGS / SusceScale + SusceOrig) + 0.015)-(Abs(val(Right$(lblDemag.Caption, 4))) / MaxDemag + 0.015, 1 - (val(frmSusceptibilityMeter.InputText) * SusceptibilityMomentFactorCGS / SusceScale + SusceOrig) - 0.015), RGB(0, 0, 255)
                MomentX.Circle (Abs(val(Right$(lblDemag.Caption, 4))) / MaxDemag, 1 - (val(frmSusceptibilityMeter.InputText) * SusceptibilityMomentFactorCGS / SusceScale + SusceOrig)), 0.02, RGB(0, 0, 255)
            End If
        End If
        If ZijLines > 0 Then
        r = 1
        ' Link the last step to the previous ones
        Zijderveld.Line (ZijY(r) / ZijScale + ZijHoriOrig, -ZijX(r) / ZijScale + ZijVertOrig)-(momentvol * Cos(crinc * Pi / 180) * Sin(crdec * Pi / 180) / ZijScale + ZijHoriOrig, -momentvol * Cos(crinc * Pi / 180) * Cos(crdec * Pi / 180) / ZijScale + ZijVertOrig), RGB(0, 0, 255)
        Zijderveld.Line (ZijY(r) / ZijScale + ZijHoriOrig, ZijZ(r) / ZijScale + ZijVertOrig)-(momentvol * Cos(crinc * Pi / 180) * Sin(crdec * Pi / 180) / ZijScale + ZijHoriOrig, momentvol * Sin(crinc * Pi / 180) / ZijScale + ZijVertOrig), RGB(255, 0, 0)
        If MomentX.Visible = True Then
            If ChkM.Visible = True And ChkM.Value = Checked And (ChkAllSteps.Value = Checked Or DemagStep(r) <= Abs(val(Right$(lblDemag.Caption, 4)))) And Not Sqr(ZijX(r) ^ 2 + ZijY(r) ^ 2 + ZijZ(r) ^ 2) = 0 Then
                MomentX.Line (DemagStep(r) / MaxDemag, 1 - Sqr(ZijX(r) ^ 2 + ZijY(r) ^ 2 + ZijZ(r) ^ 2) / MaxMoment)-(Abs(val(Right$(lblDemag.Caption, 4))) / MaxDemag, 1 - momentvol / MaxMoment), RGB(255, 0, 0)
            End If
            If ChkX.Visible = True And ChkX.Value = Checked Then
               If (ChkAllSteps.Value = Checked Or DemagStep(r) <= Abs(val(Right$(lblDemag.Caption, 4)))) And Not Susceptibility(r) = 0 Then MomentX.Line (DemagStep(r) / MaxDemag, 1 - (Susceptibility(r) / SusceScale + SusceOrig))-(1, 1 - (val(frmSusceptibilityMeter.InputText) * SusceptibilityMomentFactorCGS / SusceScale + SusceOrig)), RGB(0, 0, 255)
            End If
        End If
        End If
    End If
End Sub

Public Sub InitEqualArea()
    ' (August 2007 L Carporzen) Equal area plot near the measurement window
    Dim i As Integer
    cmdHideZij.Enabled = True
    EqualArea.Cls ' Clean the plot
    EqualArea.CurrentX = 0
    EqualArea.CurrentY = 0
    EqualArea.FontBold = True
    EqualArea.Print "Equal area" & vbCrLf & "stereoplot"
    EqualArea.FontBold = False
    EqualArea.Circle (0.5, 0.5), 0.5 ' external circle
    EqualArea.Line (0.5, 0)-(0.5, 1) ' vertical axis
    EqualArea.Line (0, 0.5)-(1, 0.5) ' horizontal axis
    For i = 1 To 8
        EqualArea.Line (0.5 + Sqr(1 - Sin(10 * i * Pi / 180)) / 2, 0.49)-(0.5 + Sqr(1 - Sin(10 * i * Pi / 180)) / 2, 0.52) ' vertical ticks
        EqualArea.Line (0.5 - Sqr(1 - Sin(10 * i * Pi / 180)) / 2, 0.49)-(0.5 - Sqr(1 - Sin(10 * i * Pi / 180)) / 2, 0.52) ' vertical ticks
        EqualArea.Line (0.49, 0.5 + Sqr(1 - Sin(10 * i * Pi / 180)) / 2)-(0.52, 0.5 + Sqr(1 - Sin(10 * i * Pi / 180)) / 2) ' horizontal ticks
        EqualArea.Line (0.49, 0.5 - Sqr(1 - Sin(10 * i * Pi / 180)) / 2)-(0.52, 0.5 - Sqr(1 - Sin(10 * i * Pi / 180)) / 2) ' horizontal ticks
    Next i
    EqualArea.CurrentX = 0.5 + 0.01
    EqualArea.CurrentY = 0.01
    EqualArea.Print "N" ' Label
    EqualArea.CurrentX = 0.5 + 0.01
    EqualArea.CurrentY = 1 - 0.05
    EqualArea.Print "S" ' Label
    EqualArea.CurrentX = 1 - 0.03
    EqualArea.CurrentY = 0.5 + 0.01
    EqualArea.Print "E" ' Label
    EqualArea.CurrentX = 0.01
    EqualArea.CurrentY = 0.5 + 0.01
    EqualArea.Print "W" ' Label
    EqualArea.CurrentX = 0.82
    EqualArea.CurrentY = 0.01
    EqualArea.Print "Up"
    EqualArea.CurrentX = 0.91
    EqualArea.CurrentY = 0.01
    EqualArea.Print "Down"
End Sub

Private Sub optBedding_Click()
    optCore.Value = False
    optGeographic.Value = False
    optBedding.Value = True
    InitEqualArea
    Actualize
End Sub

Private Sub optCore_Click()
    optCore.Value = True
    optGeographic.Value = False
    optBedding.Value = False
    Actualize
End Sub

Private Sub optGeographic_Click()
    optCore.Value = False
    optGeographic.Value = True
    optBedding.Value = False
    Actualize
End Sub

Public Sub PlotEqualArea(ByVal dec As Double, ByVal inc As Double)
    ' (August 2007 L Carporzen) Plot of the current 4 measurements (holder not substracted)
    Dim L0 As Double
    Dim L As Double
    L0 = 1 / Sqr(Cos(inc * Pi / 180) * Cos(dec * Pi / 180) * Cos(inc * Pi / 180) * Cos(dec * Pi / 180) + Cos(inc * Pi / 180) * Sin(dec * Pi / 180) * Cos(inc * Pi / 180) * Sin(dec * Pi / 180))
    If inc >= 0 Then ' Down direction
        L = L0 * Sqr(1 - Sin(inc * Pi / 180))
        EqualArea.Line ((Sin(dec * Pi / 180) * L / L0) / 2 + 0.5 - 0.01, -(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5 - 0.01)-((Sin(dec * Pi / 180) * L / L0) / 2 + 0.5 + 0.01, -(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5 + 0.01), 0.01, BF
    Else ' Up direction
        L = L0 * Sqr(1 + Sin(inc * Pi / 180))
        EqualArea.Line ((Sin(dec * Pi / 180) * L / L0) / 2 + 0.5 - 0.01, -(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5 - 0.01)-((Sin(dec * Pi / 180) * L / L0) / 2 + 0.5 + 0.01, -(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5 + 0.01), 0.01, B
    End If
End Sub

Public Sub PlotHistory(ByVal dec As Double, ByVal inc As Double, ByVal dec2 As Double, ByVal inc2 As Double)
    ' (August 2007 L Carporzen) Plot of the previous directions
    ' (June 2008 L Carporzen) Link the previous directions
    Dim L0 As Double
    Dim L As Double
    Dim L02 As Double
    Dim L2 As Double
    L0 = 1 / Sqr(Cos(inc * Pi / 180) * Cos(dec * Pi / 180) * Cos(inc * Pi / 180) * Cos(dec * Pi / 180) + Cos(inc * Pi / 180) * Sin(dec * Pi / 180) * Cos(inc * Pi / 180) * Sin(dec * Pi / 180))
    L02 = 1 / Sqr(Cos(inc2 * Pi / 180) * Cos(dec2 * Pi / 180) * Cos(inc2 * Pi / 180) * Cos(dec2 * Pi / 180) + Cos(inc2 * Pi / 180) * Sin(dec2 * Pi / 180) * Cos(inc2 * Pi / 180) * Sin(dec2 * Pi / 180))
    If inc >= 0 Then ' Down direction
        L = L0 * Sqr(1 - Sin(inc * Pi / 180))
        L2 = L02 * Sqr(1 - Sin(inc2 * Pi / 180))
        EqualArea.Circle ((Sin(dec * Pi / 180) * L / L0) / 2 + 0.5, -(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5), 0.005, RGB(0, 0, 255)
        If Not inc2 = 0 And Not L02 = 0 And Not inc = inc2 Then
            If inc / inc2 > 0 Then EqualArea.Line ((Sin(dec * Pi / 180) * L / L0) / 2 + 0.5, -(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5)-((Sin(dec2 * Pi / 180) * L2 / L02) / 2 + 0.5, -(Cos(dec2 * Pi / 180) * L2 / L02) / 2 + 0.5), RGB(0, 0, 255)
        End If
    Else ' Up direction
        L = L0 * Sqr(1 + Sin(inc * Pi / 180))
        L2 = L02 * Sqr(1 + Sin(inc2 * Pi / 180))
        EqualArea.Circle ((Sin(dec * Pi / 180) * L / L0) / 2 + 0.5, -(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5), 0.005, RGB(255, 0, 0)
        If Not inc2 = 0 And Not L02 = 0 And Not inc = inc2 Then
            If inc / inc2 > 0 Then EqualArea.Line ((Sin(dec * Pi / 180) * L / L0) / 2 + 0.5, -(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5)-((Sin(dec2 * Pi / 180) * L2 / L02) / 2 + 0.5, -(Cos(dec2 * Pi / 180) * L2 / L02) / 2 + 0.5), RGB(255, 0, 0)
        End If
    End If
End Sub

' A random number in the range (low, high)
Function Rnd2(low As Single, high As Single) As Single
    Rnd2 = Rnd * (high - low) + low
End Function

Public Sub SetFields(ByVal steps As Integer, ByVal demag As String, _
    ByVal isUp As Boolean, ByVal isBoth As Boolean, ByVal DataFileName As String)
    ' This procedure sets some of the information in the
    ' fields on this form.  It is called by frmMagnetometerControl.
    lblAvgCycles.Caption = steps
    lblDemag.Caption = demag
    frmStats.lblAvgCycles.Caption = steps
    frmStats.lblDemag.Caption = demag
    If isUp Then lblMeasDir.Caption = "U" Else lblMeasDir.Caption = "D"
    If isBoth Then lblDirs.Caption = "U/D" _
        Else lblDirs.Caption = lblMeasDir.Caption
    frmStats.lblDirs = lblDirs.Caption
    lblDataFileName = DataFileName
    frmStats.lblDataFileName = DataFileName
    lblCupNumber.Caption = Str$(modConfig.SampleHandlerCurrentHole)
End Sub

Public Sub SetSample(smp As String)
    ' This procedure sets the name of the current sample
    ' that is being processed by the magnetometer.
    lblSampName.Caption = smp
    lblCupNumber.Caption = Str$(modConfig.SampleHandlerCurrentHole)
    frmStats.lblSampName.Caption = smp
    lblSampleHeight = vbNullString
    lblMeascount = vbNullString
    lblRescan.Caption = " "
    lblXSQUID.Caption = " "
    lblYSQUID.Caption = " "
    lblZSQUID.Caption = " "
End Sub

Public Sub ShowAngDat(ByVal dec As Double, ByVal inc As Double, _
    ByVal num As Integer)
    ' This function displays angular data in appropriate boxes
    lblCalcDec(num - 1).Caption = Format$(dec, "##0.0")
    lblCalcInc(num - 1).Caption = Format$(inc, "##0.0")
End Sub

Public Sub showData(ByVal datX As Double, ByVal datY As Double, _
    ByVal datZ As Double, num As Integer)
    ' This routine displays the data fields in dat in the
    ' appropriate fields in the form. 'num' designates which
    ' fields to display into
    Dim angdat As Measure_Unfolded
    ' Show the height of each sample, measured in automatic sample changer mode (June 2007 L Carporzen)
    If lblSampName = "Holder" Then
    lblSampleHeight = Format$(SampleHolder.SampleHeight / UpDownMotor1cm, "##0.0")
    Else
    lblSampleHeight = Format$(SampleHeight / UpDownMotor1cm, "##0.0")
    End If
    lblMeascount = Meascount
    Select Case num
        Case 0
            ' Show as First Zero data points
            lblMeasXZero(0).Caption = FormatNumber(datX)
            lblMeasYZero(0).Caption = FormatNumber(datY)
            lblMeasZZero(0).Caption = FormatNumber(datZ)
        Case 1 To 4
            ' Show as Data points
            lblMeasX(num - 1).Caption = FormatNumber(datX)
            lblMeasY(num - 1).Caption = FormatNumber(datY)
            lblMeasZ(num - 1).Caption = FormatNumber(datZ)
        Case 5
            ' Show as Last Zero data points
            lblMeasXZero(1).Caption = FormatNumber(datX)
            lblMeasYZero(1).Caption = FormatNumber(datY)
            lblMeasZZero(1).Caption = FormatNumber(datZ)
    End Select
End Sub

Public Sub ShowPlots()
    If cmdShowPlots.Enabled = False Then
        cmdHideZij.Enabled = True
        Me.Width = 10490
    End If
End Sub

Public Sub ShowStats(ByVal X As Double, ByVal Y As Double, ByVal Z As Double, _
    ByVal dec As Double, ByVal inc As Double, ByVal SigDrift As Double, _
    ByVal SigHold As Double, ByVal SigInd As Double, ByVal CSD As Double)
    ' This procedure displays statistical information gathered
    ' from a measurement cycle with the magnetometer.
    Dim unfolded As Measure_Unfolded
    Me.Height = 7430
    framStats.Visible = True
    lblavgx.Caption = FormatNumber(X)
    lblavgy.Caption = FormatNumber(Y)
    lblavgz.Caption = FormatNumber(Z)
    lblavgmag.Caption = Format$(RangeFact * Sqr(X ^ 2 + Y ^ 2 + Z ^ 2), "0.0000E+")
    lblAvgDec.Caption = Format$(dec, "##0.0")
    lblAvgInc.Caption = Format$(inc, "##0.0")
    lblDSigDrift.Caption = FormatNumber(SigDrift)
    lblDSigHolder.Caption = Format$(SigHold, "0.0000E+")
    lblDSigInduced.Caption = FormatNumber(SigInd)
    lblCSD.Caption = Format$(CSD, "##0.00")
End Sub

Public Sub ShowZijderveld()
    If cmdShowPlots.Enabled = False Then Me.Width = 15480
End Sub

Public Sub updateFlowStatus()
    If Prog_halted Then
        buttonHalt.Enabled = False
        buttonPause.Caption = "Resume run"
        buttonPause.FontBold = True
    Else
        If Prog_paused Then
            buttonPause.Caption = "Resume run"
            buttonPause.FontBold = True
            buttonHalt.Enabled = True
        Else
            buttonPause.Caption = "Pause run"
            buttonPause.FontBold = False
            buttonHalt.Enabled = True
        End If
    End If
End Sub

