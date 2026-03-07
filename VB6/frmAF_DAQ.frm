VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmAF_DAQ 
   Caption         =   "MCC AF Ramp"
   ClientHeight    =   7944
   ClientLeft      =   168
   ClientTop       =   552
   ClientWidth     =   10092
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7944
   ScaleWidth      =   10092
   Begin VB.CommandButton cmdChangeFileSaveSettings 
      Caption         =   "File Save Settings"
      Height          =   372
      Left            =   4800
      TabIndex        =   94
      Top             =   1080
      Width           =   1452
   End
   Begin VB.CommandButton cmdCalibrate 
      Caption         =   "Calibrate AF Coils"
      Height          =   372
      Left            =   4800
      TabIndex        =   93
      Top             =   600
      Width           =   1452
   End
   Begin VB.CommandButton cmdTestGaussMeter 
      Caption         =   "Test Gaussmeter"
      Height          =   372
      Left            =   4800
      TabIndex        =   92
      Top             =   120
      Width           =   1452
   End
   Begin VB.CommandButton cmdTestFreq 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Test Freq"
      Height          =   372
      Left            =   1680
      MaskColor       =   &H0000FF00&
      TabIndex        =   91
      Top             =   7440
      Width           =   1332
   End
   Begin VB.CommandButton cmdTestSineFit 
      Caption         =   "Test Sine Fit "
      Height          =   372
      Left            =   8280
      TabIndex        =   90
      Top             =   7440
      Width           =   1572
   End
   Begin VB.CommandButton cmdTestRange 
      Caption         =   "Test Range Conversion"
      Height          =   372
      Left            =   6360
      TabIndex        =   87
      Top             =   7440
      Width           =   1812
   End
   Begin VB.Frame Frame7 
      Caption         =   "Active Coil"
      Height          =   1212
      Left            =   3120
      TabIndex        =   73
      Top             =   120
      Width           =   1572
      Begin VB.OptionButton optCoil 
         Caption         =   "Transverse"
         Height          =   192
         Index           =   1
         Left            =   240
         TabIndex        =   15
         Top             =   720
         Width           =   1212
      End
      Begin VB.OptionButton optCoil 
         Caption         =   "Axial"
         Height          =   192
         Index           =   0
         Left            =   240
         TabIndex        =   14
         Top             =   360
         Width           =   732
      End
   End
   Begin VB.CommandButton cmdTestFFT 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Test FFT"
      Height          =   372
      Left            =   120
      MaskColor       =   &H0000FF00&
      TabIndex        =   39
      Top             =   7440
      Width           =   1332
   End
   Begin VB.Frame Frame4 
      Caption         =   "Ramp To Peak"
      Height          =   7212
      Left            =   6360
      TabIndex        =   57
      Top             =   120
      Width           =   3612
      Begin VB.CheckBox checkClippingTest 
         Caption         =   "Un-monitored"
         Height          =   252
         Left            =   2040
         TabIndex        =   95
         Top             =   5760
         Width           =   1332
      End
      Begin VB.CheckBox checkVerbose 
         Caption         =   "Verbose?"
         Height          =   252
         Left            =   480
         TabIndex        =   36
         Top             =   5760
         Width           =   1092
      End
      Begin MSComDlg.CommonDialog cdlgTestSineFit 
         Left            =   1920
         Top             =   120
         _ExtentX        =   677
         _ExtentY        =   677
         _Version        =   393216
      End
      Begin VB.TextBox txtMonitorTrigVolt 
         Height          =   285
         Left            =   2280
         TabIndex        =   30
         Top             =   3240
         Width           =   1092
      End
      Begin VB.CheckBox checkDoSineWave 
         Caption         =   "Check1"
         Height          =   252
         Left            =   2280
         TabIndex        =   35
         Top             =   5280
         Width           =   252
      End
      Begin VB.TextBox txtRampRate 
         Height          =   285
         Left            =   2280
         TabIndex        =   32
         Top             =   3960
         Width           =   1092
      End
      Begin VB.TextBox txtRampDownDuration 
         Height          =   285
         Left            =   2280
         TabIndex        =   34
         Top             =   4680
         Width           =   1092
      End
      Begin VB.TextBox txtRampUpDuration 
         Height          =   285
         Left            =   2280
         TabIndex        =   33
         Top             =   4320
         Width           =   1092
      End
      Begin VB.CommandButton cmdAbortRamp 
         BackColor       =   &H000000FF&
         Caption         =   "Abort Ramp!!"
         Height          =   372
         Left            =   600
         TabIndex        =   38
         Top             =   6720
         Width           =   2295
      End
      Begin VB.CommandButton cmdStartRamp 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Start Ramp"
         Height          =   372
         Left            =   600
         TabIndex        =   37
         Top             =   6240
         Width           =   2295
      End
      Begin VB.TextBox txtRampPeakDuration 
         Height          =   285
         Left            =   2280
         TabIndex        =   31
         Top             =   3600
         Width           =   1092
      End
      Begin VB.TextBox txtRampPeakVoltage 
         Height          =   285
         Left            =   2280
         TabIndex        =   29
         Top             =   2880
         Width           =   1092
      End
      Begin VB.Frame Frame6 
         Caption         =   "Output Comm"
         Height          =   1212
         Left            =   120
         TabIndex        =   61
         Top             =   1560
         Width           =   3372
         Begin VB.TextBox txtOutBoardRampName 
            Height          =   288
            Left            =   1800
            TabIndex        =   82
            Top             =   360
            Width           =   1452
         End
         Begin VB.ComboBox cmbOutBoardRamp 
            Height          =   288
            Left            =   840
            TabIndex        =   27
            Top             =   360
            Width           =   732
         End
         Begin VB.ComboBox cmbOutChanRamp 
            Height          =   288
            Left            =   840
            TabIndex        =   28
            Top             =   840
            Width           =   2412
         End
         Begin VB.Label Label31 
            Caption         =   "Board Name:"
            Height          =   252
            Left            =   1800
            TabIndex        =   83
            Top             =   120
            Width           =   1212
         End
         Begin VB.Label Label18 
            Caption         =   "Board #:"
            Height          =   252
            Left            =   120
            TabIndex        =   63
            Top             =   360
            Width           =   1452
         End
         Begin VB.Label Label17 
            Caption         =   "Channel:"
            Height          =   252
            Left            =   120
            TabIndex        =   62
            Top             =   840
            Width           =   732
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Input Comm"
         Height          =   1212
         Left            =   120
         TabIndex        =   58
         Top             =   240
         Width           =   3372
         Begin VB.TextBox txtInBoardRampName 
            Height          =   288
            Left            =   1800
            TabIndex        =   80
            TabStop         =   0   'False
            Top             =   360
            Width           =   1452
         End
         Begin VB.ComboBox cmbInChanRamp 
            Height          =   288
            Left            =   840
            TabIndex        =   26
            Top             =   840
            Width           =   2412
         End
         Begin VB.ComboBox cmbInBoardRamp 
            Height          =   288
            Left            =   840
            TabIndex        =   25
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label30 
            Caption         =   "Board Name:"
            Height          =   252
            Left            =   1800
            TabIndex        =   81
            Top             =   120
            Width           =   1212
         End
         Begin VB.Label Label16 
            Caption         =   "Channel:"
            Height          =   252
            Left            =   120
            TabIndex        =   60
            Top             =   840
            Width           =   732
         End
         Begin VB.Label Label11 
            Caption         =   "Board #:"
            Height          =   252
            Left            =   120
            TabIndex        =   59
            Top             =   360
            Width           =   1452
         End
      End
      Begin VB.Label Label33 
         Caption         =   "Monitor Trigger Voltage:"
         Height          =   252
         Left            =   240
         TabIndex        =   86
         Top             =   3240
         Width           =   1932
      End
      Begin VB.Label Label24 
         Caption         =   "Ramp With Along with Sine Wave Generator?"
         Height          =   372
         Left            =   240
         TabIndex        =   69
         Top             =   5160
         Width           =   1812
      End
      Begin VB.Label Label23 
         Caption         =   "Ramp Rate (Hz):"
         Height          =   252
         Left            =   240
         TabIndex        =   68
         Top             =   3960
         Width           =   1932
      End
      Begin VB.Label Label22 
         Caption         =   "Ramp Down Duration (ms):"
         Height          =   252
         Left            =   240
         TabIndex        =   67
         Top             =   4680
         Width           =   1932
      End
      Begin VB.Label Label21 
         Caption         =   "Ramp Up Duration (ms):"
         Height          =   252
         Left            =   240
         TabIndex        =   66
         Top             =   4320
         Width           =   2052
      End
      Begin VB.Label Label20 
         Caption         =   "Duration at Peak (ms):"
         Height          =   252
         Left            =   240
         TabIndex        =   65
         Top             =   3600
         Width           =   1932
      End
      Begin VB.Label Label19 
         Caption         =   "Amplitude Voltage (0-10 V):"
         Height          =   252
         Left            =   240
         TabIndex        =   64
         Top             =   2880
         Width           =   1932
      End
   End
   Begin VB.TextBox txtAmplitude 
      Height          =   285
      Left            =   5040
      TabIndex        =   17
      Top             =   2280
      Width           =   975
   End
   Begin VB.Frame Frame3 
      Caption         =   "Generate Sine Wave"
      Height          =   5892
      Left            =   3120
      TabIndex        =   50
      Top             =   1440
      Width           =   3132
      Begin VB.TextBox txtSineIORate 
         Height          =   285
         Left            =   1920
         TabIndex        =   88
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox txtBoardSineName 
         Height          =   288
         Left            =   1440
         TabIndex        =   84
         TabStop         =   0   'False
         Top             =   3240
         Width           =   1452
      End
      Begin VB.CommandButton cmdStartAFTuner 
         BackColor       =   &H00FF80FF&
         Caption         =   "Tune AF Coils"
         Height          =   372
         Left            =   240
         MaskColor       =   &H00C000C0&
         TabIndex        =   24
         Top             =   5280
         Width           =   2652
      End
      Begin VB.CommandButton cmdStopSineWave 
         BackColor       =   &H008080FF&
         Caption         =   "Stop Analog Output"
         Height          =   372
         Left            =   240
         MaskColor       =   &H000000C0&
         TabIndex        =   23
         Top             =   4800
         Width           =   2652
      End
      Begin VB.TextBox txtPtsPerPeriod 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1920
         TabIndex        =   19
         Top             =   2280
         Width           =   975
      End
      Begin VB.TextBox txtDuration 
         Height          =   285
         Left            =   1920
         TabIndex        =   18
         Top             =   1800
         Width           =   975
      End
      Begin VB.CommandButton cmdOutputSineA 
         BackColor       =   &H0080FF80&
         Caption         =   "Ouput Analog Sine Wave"
         Height          =   372
         Left            =   240
         MaskColor       =   &H0000C000&
         TabIndex        =   22
         Top             =   4320
         Width           =   2652
      End
      Begin VB.ComboBox cmbChanSine 
         Height          =   288
         Left            =   1920
         TabIndex        =   21
         Top             =   3720
         Width           =   735
      End
      Begin VB.ComboBox cmbBoardSine 
         Height          =   288
         Left            =   1920
         TabIndex        =   20
         Top             =   2760
         Width           =   735
      End
      Begin VB.TextBox txtFreq 
         Height          =   285
         Left            =   1920
         TabIndex        =   16
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label34 
         Caption         =   "Board IORate (Hz):"
         Height          =   252
         Left            =   240
         TabIndex        =   89
         Top             =   1320
         Width           =   1692
      End
      Begin VB.Label Label32 
         Caption         =   "Board Name:"
         Height          =   252
         Left            =   240
         TabIndex        =   85
         Top             =   3240
         Width           =   1212
      End
      Begin VB.Label Label15 
         Caption         =   "# Points / Period:"
         Height          =   252
         Left            =   240
         TabIndex        =   56
         Top             =   2280
         Width           =   1692
      End
      Begin VB.Label Label14 
         Caption         =   "Duration of Wave (ms):"
         Height          =   252
         Left            =   240
         TabIndex        =   55
         Top             =   1800
         Width           =   1692
      End
      Begin VB.Label Label13 
         Caption         =   "Channel/Pin:"
         Height          =   252
         Left            =   240
         TabIndex        =   54
         Top             =   3720
         Width           =   1452
      End
      Begin VB.Label Label12 
         Caption         =   "Board #:"
         Height          =   252
         Left            =   240
         TabIndex        =   53
         Top             =   2760
         Width           =   1452
      End
      Begin VB.Label Label10 
         Caption         =   "Amplitude (0 - 10 V):"
         Height          =   252
         Left            =   240
         TabIndex        =   52
         Top             =   840
         Width           =   1452
      End
      Begin VB.Label Label9 
         Caption         =   "Frequency (Hz):"
         Height          =   252
         Left            =   240
         TabIndex        =   51
         Top             =   360
         Width           =   1212
      End
   End
   Begin VB.TextBox txtRawD 
      Height          =   285
      Left            =   240
      TabIndex        =   10
      Top             =   6000
      Width           =   852
   End
   Begin VB.TextBox txtEngD 
      Height          =   285
      Left            =   1680
      TabIndex        =   11
      Top             =   6000
      Width           =   852
   End
   Begin VB.ComboBox cmbBoardD 
      Height          =   288
      Left            =   240
      TabIndex        =   7
      Top             =   3960
      Width           =   735
   End
   Begin VB.ComboBox cmbBoardA 
      Height          =   288
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   735
   End
   Begin VB.CommandButton cmdDigitalOutput 
      Caption         =   "Digital Output"
      Height          =   372
      Left            =   1680
      TabIndex        =   13
      Top             =   6480
      Width           =   1092
   End
   Begin VB.CommandButton cmdDigitalInput 
      Caption         =   "Digital Input"
      Height          =   372
      Left            =   240
      TabIndex        =   12
      Top             =   6480
      Width           =   972
   End
   Begin VB.CommandButton cmdAnalogOutput 
      Caption         =   "Analog Output"
      Height          =   372
      Left            =   1680
      TabIndex        =   6
      Top             =   2400
      Width           =   1212
   End
   Begin VB.CommandButton cmdAnalogInput 
      Caption         =   "Analog Input"
      Height          =   372
      Left            =   240
      TabIndex        =   5
      Top             =   2400
      Width           =   1092
   End
   Begin VB.TextBox txtEngA 
      Height          =   285
      Left            =   1680
      TabIndex        =   4
      Top             =   1920
      Width           =   852
   End
   Begin VB.TextBox txtRawA 
      Height          =   285
      Left            =   240
      TabIndex        =   3
      Top             =   1920
      Width           =   852
   End
   Begin VB.ComboBox cmbChanAIn 
      Height          =   288
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   1212
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H008080FF&
      Caption         =   "Close"
      Height          =   372
      Left            =   3240
      MaskColor       =   &H000000C0&
      TabIndex        =   41
      Top             =   7440
      Width           =   2892
   End
   Begin VB.Frame Frame1 
      Caption         =   "Analog IO"
      Height          =   3132
      Left            =   120
      TabIndex        =   40
      Top             =   120
      Width           =   2892
      Begin VB.TextBox txtBoardNameA 
         Height          =   288
         Left            =   1200
         TabIndex        =   78
         TabStop         =   0   'False
         Top             =   480
         Width           =   1452
      End
      Begin VB.ComboBox cmbChanAOut 
         Height          =   288
         Left            =   1560
         TabIndex        =   2
         Top             =   1080
         Width           =   1212
      End
      Begin VB.Label Label28 
         Caption         =   "Board Name:"
         Height          =   252
         Left            =   1200
         TabIndex        =   79
         Top             =   240
         Width           =   1212
      End
      Begin VB.Label Label29 
         Alignment       =   2  'Center
         Caption         =   "Outputs read English value."
         Height          =   252
         Left            =   360
         TabIndex        =   75
         Top             =   2760
         Width           =   2292
      End
      Begin VB.Label Label25 
         Caption         =   "Output Channel:"
         Height          =   252
         Left            =   1560
         TabIndex        =   70
         Top             =   840
         Width           =   1212
      End
      Begin VB.Label Label1 
         Caption         =   "Input Channel:"
         Height          =   252
         Left            =   120
         TabIndex        =   45
         Top             =   840
         Width           =   1212
      End
      Begin VB.Label Label2 
         Caption         =   "Raw Value:"
         Height          =   252
         Left            =   120
         TabIndex        =   44
         Top             =   1560
         Width           =   1212
      End
      Begin VB.Label Label3 
         Caption         =   "English Value:"
         Height          =   252
         Left            =   1560
         TabIndex        =   43
         Top             =   1560
         Width           =   1092
      End
      Begin VB.Label lblBoard 
         Caption         =   "Board #:"
         Height          =   252
         Left            =   120
         TabIndex        =   42
         Top             =   240
         Width           =   1212
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Digital IO"
      Height          =   3972
      Index           =   0
      Left            =   120
      TabIndex        =   46
      Top             =   3360
      Width           =   2892
      Begin VB.TextBox txtBoardNameD 
         Height          =   288
         Left            =   1200
         TabIndex        =   76
         TabStop         =   0   'False
         Top             =   600
         Width           =   1452
      End
      Begin VB.ComboBox cmbChanDIn 
         Height          =   288
         Left            =   120
         TabIndex        =   8
         Top             =   1320
         Width           =   2532
      End
      Begin VB.ComboBox cmbChanDOut 
         Height          =   288
         Left            =   120
         TabIndex        =   9
         Top             =   2040
         Width           =   2532
      End
      Begin VB.Label Label27 
         Caption         =   "Board Name:"
         Height          =   252
         Left            =   1200
         TabIndex        =   77
         Top             =   360
         Width           =   1212
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "Outputs read English value."
         Height          =   252
         Left            =   240
         TabIndex        =   74
         Top             =   3600
         Width           =   2292
      End
      Begin VB.Label Label26 
         Caption         =   "Output Channel:"
         Height          =   252
         Left            =   120
         TabIndex        =   72
         Top             =   1800
         Width           =   1212
      End
      Begin VB.Label Label8 
         Caption         =   "Input Channel:"
         Height          =   252
         Left            =   120
         TabIndex        =   71
         Top             =   1080
         Width           =   1212
      End
      Begin VB.Label Label7 
         Caption         =   "Raw Value:"
         Height          =   252
         Left            =   120
         TabIndex        =   49
         Top             =   2400
         Width           =   1212
      End
      Begin VB.Label Label6 
         Caption         =   "English Value:"
         Height          =   252
         Left            =   1560
         TabIndex        =   48
         Top             =   2400
         Width           =   1212
      End
      Begin VB.Label Label5 
         Caption         =   "Board #:"
         Height          =   252
         Left            =   120
         TabIndex        =   47
         Top             =   360
         Width           =   1212
      End
   End
   Begin VB.Menu mnuLines 
      Caption         =   "Lines"
      NegotiatePosition=   3  'Right
      Visible         =   0   'False
      Begin VB.Menu mnuDigiCheck 
         Caption         =   "DigiCheck"
      End
      Begin VB.Menu mnuClearBit 
         Caption         =   "Clear Bit"
      End
      Begin VB.Menu mnuSetBit 
         Caption         =   "Set Bit"
      End
   End
End
Attribute VB_Name = "frmAF_DAQ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' MCC Controller
'
' This is the driver for the MCC I/O Interface
'
Public NOCOMM_MODE As Boolean
Public Sub ZeroRampVoltage()

    Dim i As Long
    Dim ULStats As Long
    Dim RampWave As Wave
    
    Set RampWave = Nothing
    
    If isnothing(DAQBoards) Or isnothing(WaveForms) Then
    
        Initialize_Boards
        
        Initialize_Waves
        
    End If
    
    For i = 1 To WaveForms.count
    
        With WaveForms.Item(i)
        
            If .WaveType = AFRAMPUP Then
            
                Set RampWave = WaveForms.Item(i)
                
            End If
            
        End With
        
    Next i
    
    With RampWave
        
        'Zero Ramp voltage
        ULStats = .BoardUsed.AnalogOut(.Range, _
                                        .Chan, _
                                        0)
                            
        'Error Check
        If ULStats <> 0 Then
        
            Err.Raise ULStats, _
                        "ZeroRampVoltage->Board.AnalogOut", _
                        "Error setting Ramp Amplifier analog out voltage to zero." & _
                        vbNewLine & "Board = " & .BoardUsed.BoardName & " (" & _
                        Trim(Str(.BoardUsed.BoardNum)) & ")" & vbNewLine & _
                        "Channel = " & .Chan.ChanName & " (" & _
                        Trim(Str(.Chan.ChanNum)) & ")"
                            
'--------Debug ONLY! - break program here---------------------------------------
            End
'-------------------------------------------------------------------------------
            
        End If
    
    End With

End Sub

Private Sub cmbBoardA_Click()

    Dim i As Long
    Dim BoardNum As Long
        
    With cmbBoardA
        
        BoardNum = .ItemData(.ListIndex)
    
    End With
    
    cmbChanAIn.Clear
    cmbChanAOut.Clear
    
    If Not IsObject(DAQBoards) Then
    
        'DAQBoards is not initialized
        Initialize_Boards
        
    End If
    
    If BoardNum > DAQBoards.count Then
        
        If DAQBoards.count = 0 Then
        
            'No Boards Loaded!
            Initialize_Boards
            
        End If
        
        'The inputed Board Num does not correspond to a loaded board
        MsgBox "Board ID Error!" & vbNewLine & _
                "Board Number = " & Str(BoardNum - 1) & vbNewLine & _
                "Valid Board Numbers range from 0 to " & _
                Trim(Str(DAQBoards.count - 1))
    
        Exit Sub
        
    End If
    
    With DAQBoards.Item(BoardNum + 1)
        
        txtBoardNameA = .BoardName
        
        For i = 1 To .AInChannels.count
        
            cmbChanAIn.AddItem .AInChannels.Item(i).ChanName
            cmbChanAIn.ItemData(cmbChanAIn.NewIndex) = .AInChannels.Item(i).ChanNum
            
            If i = 1 Then
            
                cmbChanAIn.ListIndex = 0
                
            End If
            
        Next i
        
        For i = 1 To .AOutChannels.count
        
            cmbChanAOut.AddItem .AOutChannels.Item(i).ChanName
            cmbChanAOut.ItemData(cmbChanAOut.NewIndex) = .AOutChannels.Item(i).ChanNum
            
            If i = 1 Then
            
                cmbChanAOut.ListIndex = 0
                
            End If
            
        Next i
    
    End With
    
End Sub
Private Sub cmbBoardD_Click()

    Dim i As Long
    Dim BoardNum As Long
    
    With cmbBoardD
    
        BoardNum = .ItemData(.ListIndex)
    
    End With
    
    cmbChanDIn.Clear
    cmbChanDOut.Clear
    
    If Not IsObject(DAQBoards) Then
    
        'DAQBoards is not initialized
        Initialize_Boards
        
    End If
    
    If BoardNum > DAQBoards.count Then
        
        If DAQBoards.count = 0 Then
        
            'No Boards Loaded!
            Initialize_Boards
            
        End If
        
        'The inputed Board Num does not correspond to a loaded board
        MsgBox "Board ID Error!" & vbNewLine & _
                "Board Number = " & Str(BoardNum - 1) & vbNewLine & _
                "Valid Board Numbers range from 0 to " & _
                Trim(Str(DAQBoards.count - 1))
    
        Exit Sub
        
    End If
    
    With DAQBoards.Item(BoardNum + 1)
        
        txtBoardNameD = .BoardName
        
        For i = 1 To .DInChannels.count
        
            cmbChanDIn.AddItem .DInChannels.Item(i).ChanName
            cmbChanDIn.ItemData(cmbChanDIn.NewIndex) = .DInChannels.Item(i).ChanNum
            
            If i = 1 Then
            
                cmbChanDIn.ListIndex = 0
                
            End If
                        
        Next i
        
        For i = 1 To .DOutChannels.count
        
            cmbChanDOut.AddItem .DOutChannels.Item(i).ChanName
            cmbChanDOut.ItemData(cmbChanDOut.NewIndex) = .DOutChannels.Item(i).ChanNum
            
            If i = 1 Then
            
                cmbChanDOut.ListIndex = 0
                
            End If
            
        Next i
    
    End With
End Sub

Private Sub cmbBoardSine_Click()

    Dim i As Long

    Dim BoardNum As Long
    
    With cmbBoardSine
    
        BoardNum = .ItemData(.ListIndex) + 1
    
    End With
    
    cmbChanSine.Clear
    If Not IsObject(DAQBoards) Then
    
        'DAQBoards is not initialized
        Initialize_Boards
        
    End If
    
    If BoardNum > DAQBoards.count Then
        
        If DAQBoards.count = 0 Then
        
            'No Boards Loaded!
            Initialize_Boards
            
        End If
        
        'The inputed Board Num does not correspond to a loaded board
        MsgBox "Board ID Error!" & vbNewLine & _
                "Board Number = " & Str(BoardNum - 1) & vbNewLine & _
                "Valid Board Numbers range from 0 to " & _
                Trim(Str(DAQBoards.count - 1))
    
        Exit Sub
        
    End If
    
    With DAQBoards.Item(BoardNum)
    
        If InStr(1, .BoardFunction, AFRAMP) > 0 Then
            
                       
            'User Selected Board is NOT a Function generator board!
            MsgBox "Wrong Board Error!" & vbNewLine & _
                    "Board Num = " & Trim(Str(.BoardNum)) & vbNewLine & _
                    "Board Name = " & Trim(.BoardName)
            
            Exit Sub
            
        End If
        
        Me.txtBoardSineName = .BoardName
        
        For i = 1 To .AOutChannels.count
        
            cmbChanSine.AddItem .AOutChannels.Item(i).ChanName
            cmbChanSine.ItemData(cmbChanSine.NewIndex) = .AOutChannels.Item(i).ChanNum
            
            If i = 1 Then
            
                cmbChanSine.ListIndex = 0
                
            End If
            
        Next i
        
    End With
        
End Sub


Private Sub cmbInBoardRamp_Click()

    Dim i As Long

    Dim BoardNum As Long

    With cmbInBoardRamp
    
        BoardNum = .ItemData(.ListIndex) + 1
    
    End With
    
    cmbInChanRamp.Clear
        
    With DAQBoards.Item(BoardNum).AInChannels
        
        Me.txtInBoardRampName = DAQBoards.Item(BoardNum).BoardName
        
        For i = 1 To .count
            
            cmbInChanRamp.AddItem .Item(i).ChanName
            cmbInChanRamp.ItemData(cmbInChanRamp.NewIndex) = .Item(i).ChanNum
                
            If i = 1 Then
                
                cmbInChanRamp.ListIndex = 0
                    
            End If
                
        Next i
        
    End With

End Sub





Private Sub cmbOutBoardRamp_Click()

    Dim i As Long

    Dim BoardNum As Long

    With cmbOutBoardRamp

        BoardNum = .ItemData(.ListIndex) + 1
    
    End With
    
    cmbOutChanRamp.Clear
        
    With DAQBoards.Item(BoardNum).AOutChannels
        
        Me.txtOutBoardRampName = DAQBoards.Item(BoardNum).BoardName
        
        For i = 1 To .count
            
            cmbOutChanRamp.AddItem .Item(i).ChanName
            cmbOutChanRamp.ItemData(cmbOutChanRamp.NewIndex) = .Item(i).ChanNum
                
            If i = 1 Then
                
                cmbOutChanRamp.ListIndex = 0
                    
            End If
                
        Next i
    
    End With

End Sub



Private Sub cmdAbortRamp_Click()

    Dim i As Long
    Dim BoardNum As Long
    Dim Chan As Long
    Dim DataValue As Integer
    
    Dim FuncType As Integer
    Dim Status As Integer
    Dim CurCount As Long
    Dim CurIndex As Long
    
    BoardNum = cmbOutBoardRamp.ItemData(cmbOutBoardRamp.ListIndex)
    Chan = cmbOutChanRamp.ItemData(cmbOutChanRamp.ListIndex)
    
    If WaveForms Is Nothing Or WaveForms.count = 0 Then
    
        'CRAP!
        MsgBox "Bad Ramp Wave info!" & vbNewLine & _
                "Garbage values must have been dumped into the Ramp Wave object." & _
                vbNewLine & "Code will end right now!"
        
        End
        
    End If
    
    With WaveForms
    
        For i = 1 To .count
        
            With .Item(i)
            
                If .BoardUsed Is Nothing Then
                    
                    'CRAP!
                    MsgBox "Bad Ramp Board info!" & vbNewLine & _
                            "Garbage values must have been dumped into the Ramp Board object." & _
                            vbNewLine & "Code will end right now!"
                    End
                    
                End If
                
                                   
                'Check if Ramp Wave channel is a non-nothing channel object
                If .Chan Is Nothing Then
                
                    'CRAP!
                    MsgBox "Bad Ramp Channel info!" & vbNewLine & _
                            "Garbage values must have been dumped into the Ramp Channel object." & _
                            vbNewLine & "Code will end right now!"
                    End
                
                End If
                
                If .IO = IOINPUT And .Chan.ChanName Like "A*" Then
                
                    'This is an analog input channel
                    
                    FuncType = AIFUNCTION
                    
                End If
                
                If .IO = IOINPUT And .Chan.ChanName Like "D*" Then
                
                    'This is a digital input channel
                    FuncType = DIFUNCTION
                    
                End If
                
                If .IO = IOOUTPUT And .Chan.ChanName Like "A*" Then
                
                    'This is an analog output channel
                    FuncType = AOFUNCTION
                    
                End If
                
                If .IO = IOOUTPUT And .Chan.ChanName Like "D*" Then
                
                    'This is an analog output channel
                    FuncType = DOFUNCTION
                    
                End If
                
                'Stop Background process on this board
                ULStats = cbStopBackground(.BoardUsed.BoardNum, FuncType)
                
                If ULStats <> 0 Then
    
                    MsgBox "Ending Background Ramp Failed!" & vbNewLine & _
                            "Board Number = " & Str(.BoardUsed.BoardNum) & _
                            ", " & .BoardUsed.BoardName & vbNewLine & _
                            "Err Number: " & Str(ULStats)
                    
                    Exit Sub
        
                End If
                
                ULStats = cbFromEngUnits(.BoardUsed.BoardNum, _
                                            .Range.RangeType, _
                                            0, _
                                            DataValue)
            
                If ULStats <> 0 Then
                
                    MsgBox "Converting Zero Volts to count number failed!" & vbNewLine & _
                            "Board Number = " & Str(.BoardUsed.BoardNum) & _
                            ", " & .BoardUsed.BoardName & vbNewLine & _
                            "Err Number: " & Str(ULStats)
                    
                    Exit Sub
                    
                End If
            
                ULStats = cbAOut(.BoardUsed.BoardNum, _
                                    .Chan.ChanNum, _
                                    .Range.RangeType, _
                                    DataValue)
                
                If ULStats <> 0 Then
                
                    MsgBox "Zeroing at end of Abort Failed!" & vbNewLine & _
                            "Board Number = " & Str(.BoardUsed.BoardNum) & _
                            ", " & .BoardUsed.BoardName & vbNewLine & _
                            "Channel = " & .Chan.ChanName & " (" & _
                            Trim(Str(.Chan.ChanNum)) & ")" & vbNewLine & _
                            "Err Number: " & Str(ULStats)
                    
                    Exit Sub
                    
                End If
                                            
                'Check to see if the memory buffer for this wave is still allocated
                'and still needs to be freed and should be freed
                If .BufferAlloc And .DoDeallocate Then
                    
                    'Now clear out the wave memory buffer
                    ULStats = cbWinBufFree(.MemBuffer)
                    
                    If ULStats <> 0 Then
                    
                        MsgBox "Error deallocating Ramp memory buffer!" & vbNewLine & _
                                "Memory Buffer Reference # = " & Trim(Str(.MemBuffer)) & _
                                vbNewLine & "Err Number: " & Str(ULStats)
                                
                        Exit Sub
                        
                    End If
                    
                    .MemBuffer = 0
                    .BufferAlloc = False
                    
                End If
          
            End With
            
        Next i
        
    End With
    
    Me.cmdStartRamp.Enabled = True
    
End Sub

Private Sub cmdAnalogInput_Click()

    Dim DataValue As Integer
    Dim BoardNum As Long
    
    BoardNum = cmbBoardA.ItemData(cmbBoardA.ListIndex)

    AnalogInput BoardNum, _
                cmbChanAIn.ItemData(cmbChanAIn.ListIndex), _
                DataValue%, _
                DAQBoards(BoardNum + 1).Range
                
                
    
End Sub

Private Sub cmdAnalogOutput_Click()

    Dim BoardNum As Long
    
    BoardNum = cmbBoardA.ItemData(cmbBoardA.ListIndex)

    AnalogOutput BoardNum, _
                    cmbChanAOut.ItemData(cmbChanAOut.ListIndex), _
                    val(txtEngA), _
                    DAQBoards(BoardNum + 1).Range
                    
End Sub

Private Sub cmdCalibrate_Click()

    frmCalibrateAF.Show
        
End Sub

Private Sub cmdChangeFileSaveSettings_Click()

    frmFileSave.Show

End Sub

Private Sub cmdDigitalInput_Click()
    
    Dim DataValue As Integer

    DigitalInput cmbBoardD.ItemData(cmbBoardD.ListIndex), _
                    cmbChanDIn.ItemData(cmbChanDIn.ListIndex), _
                    DataValue
                    
End Sub

Private Sub cmdDigitalOutput_Click()
    DigitalOutput cmbBoardD.ItemData(cmbBoardD.ListIndex), _
                    cmbChanDOut.ItemData(cmbChanDOut.ListIndex), _
                    val(txtEngD)
End Sub

Private Sub cmdOutputSineA_Click()

    Dim BoardNum As Long
    Dim SineChan As Channel
    Dim i As Long
    Dim found As Boolean
    Dim waveIndex As Long
    
    Set SineChan = Nothing
    Set SineChan = New Channel
    
    With cmbBoardSine
    
        BoardNum = .ItemData(.ListIndex) + 1
        
    End With
    
    With DAQBoards.Item(BoardNum)
        
        If InStr(1, .BoardFunction, AFRAMP) > 0 Then
           
            'User Selected Board is NOT a Function generator board!
            MsgBox "Wrong Board Error!" & vbNewLine & _
                    "Board Num = " & Trim(Str(.BoardNum)) & vbNewLine & _
                    "Board Name = " & Trim(.BoardName)
        
            Exit Sub
        
        End If
    
        With cmbChanSine
        
            SineChan.ChanName = .Text
            SineChan.ChanNum = .ItemData(.ListIndex)

        End With

        found = False

        With .AOutChannels
        
            For i = 1 To .count
            
                If SineChan.ChanName = .Item(i).ChanName And _
                    SineChan.ChanNum = .Item(i).ChanNum _
                Then
                
                    found = True
                    
                    Exit For
                    
                End If
                
            Next i
                
        End With

        If Not found Then
            
            'No channel on this board corresponds to the user inputed channel
            'Taunt the user with their obvious stupidity
            MsgBox "User has selected a bad channel for this board!" & vbNewLine & _
                    "Board = " & .BoardName & " (" & Trim(Str(.BoardNum)) & ")" & _
                    vbNewLine & "Has possible Analog Output channels 0 - " & _
                    Trim(Str(.AOutChannels.count - 1)) & vbNewLine & _
                    "Channel number inputed was: " & Str(SineChan.ChanNum)
                    
            Exit Sub
            
        End If
        
    End With

    With WaveForms
    
        For i = 1 To .count
        
            If .Item(i).WaveType = SINEWAVE Then
            
                waveIndex = i
                Exit For
                
            End If
            
        Next i
        
        With .Item(waveIndex)
        
            Set .Chan = SineChan
            Set .BoardUsed = DAQBoards(BoardNum)
            .Duration = val(txtDuration.Text)
            .SineFreq = val(txtFreq.Text)
            .IORate = CLng(val(txtSineIORate))
            .PtsPerPeriod = val(txtPtsPerPeriod.Text)
            .PeakVoltage = val(txtAmplitude.Text)
            .Range.RangeType = BIP10VOLTS
                        
        End With
            
    End With
    
    generateWave WaveForms(waveIndex)
    
    Set SineChan = Nothing
    
End Sub
Public Function generateWave(SWave As Wave) As Boolean

    Dim Wave() As Integer
    Dim SineArray() As Double
    Dim ULStats As Long
    Dim NumPeriods As Double
    Dim DataValue As Integer
    Dim engUnits As Single
    
    Dim i As Long
    Dim Status As Integer
    Dim CurCount As Long
    Dim CurIndex As Long
    Dim Voltage As Double
    Dim MCCCounts As Integer
    Dim FixedCounts As Integer
    Dim Sum As Single
    Dim gainArray(1) As Long
    Dim DurationStr As String
    Dim NumZeros As Long
    
    Dim Residual As Double
    Dim doContinue As Boolean
    Dim NoError As Boolean

    'Now need to construct one phase of the Sine wave to repeat
    'Maximum output resolution of board is MAX_FREQ (usually 1 MHz),
    'so need to compare freq desired to 1Mhz and see how many points
    'in the wave are possible
    
    With SWave
        
        If Abs(.PeakVoltage) > 10 Or .PeakVoltage < 0 Then
        
            MsgBox "Amplitude is outside of allowed range (0 to 10 volts)"
            
            Exit Function
            
        End If
        
        If .IORate > .BoardUsed.MaxAOutRate Then
        
            MsgBox "User has chosen a sine wave frequency and # of points per period that requires " & _
                    "a faster analog output rate than the max rate for this board: " & _
                    vbNewLine & Str(.BoardUsed.MaxAOutRate)
                    
        End If
        
        
        'Find the number of zeros at the end of the value in .Duration
        
        DurationStr = Trim(Str(.Duration))
        
        For i = 0 To Len(DurationStr) - 1
        
            If Mid(DurationStr, Len(DurationStr) - 1 - i, 1) <> "0" Then
            
                'Exit the loop by altering duration string so that it's
                'length goes to zero
                NumZeros = i + 1
                i = Len(DurationStr) + 1
                
            End If
            
        Next i
        
        'Round .PtsPerPeriod so that it has the corresponding # of significant digits.
        .PtsPerPeriod = Round(.PtsPerPeriod, NumZeros)
        
        .NumPoints = CLng(.PtsPerPeriod * .Duration)
            
'-------Sine Frequency Error debug code, only-----------------------------------------
'       Code added to record to file NumPeriods, Duration, .SineFreq, and .NumPoints
'       and the spacing in between zero-crossings in the sine wave
'
'       November 21, 2009
'
'
'        'Create File system object
'        Dim fso
'        Dim TxtFile
'        Dim FileName As String
'
'        'Generate the filename for the log file
'        FileName = "C:\Documents and Settings\lab\Desktop\Test MCC Board 11-16-2009\" & _
'                    Trim(Str(.SineFreq)) & "Hz_" & Trim(Str(.Duration)) & "ms_" & _
'                    Trim(Str(.PtsPerPeriod)) & "pper.tsv"
'
'        'Set / allocate the file system and text stream object
'        Set fso = CreateObject("Scripting.FileSystemObject")
'        Set TxtFile = fso.CreateTextFile(FileName, True)
'
'        'Write the Sine Freq
'        TxtFile.WriteLine ("Sine Freq" & vbTab & Trim(Str(.SineFreq)) & vbTab & "Hz")
'
'        'Write the Duration
'        TxtFile.WriteLine ("Duration" & vbTab & Trim(Str(.Duration)) & vbTab & "msec")
'
'        'Write the number of periods
'        TxtFile.WriteLine ("Pts per Period" & vbTab & Trim(Str(.PtsPerPeriod)))
'
'        'Write the Number of periods
'        TxtFile.WriteLine ("Number Periods" & vbTab & Trim(Str(NumPeriods)))
'
'        'Write number of points
'        TxtFile.WriteLine ("Number Points" & vbTab & Trim(Str(.NumPoints)))
'
'        'Now write the column headers
'        TxtFile.WriteLine ("Point Position" & vbTab & "Zero value" & vbTab & "Sine Input")
'
'------------------------------------------------------------------------------------------
        
        'Create array to store sine function
        'Add 100 points to allow zero-capturing in the last periods worth of data
        ReDim SineArray(.NumPoints + 100)
                
        For i = 0 To .NumPoints - 1
                                
            'Find current needed voltage
            SineArray(i) = .PeakVoltage * Sin(i / .PtsPerPeriod * 2 * Pi)
            
        Next i
                
        
        'Now, setup and start the Sine Wave analog output process on the MCC board
        NoError = .ManageBackgroundProcess(AOFUNCTION, _
                                                SineArray, _
                                                "Sine Wave", _
                                                False, _
                                                False)
                                                
        'Propagate error status to whatever called generateWave
        generateWave = NoError
        
    End With
                   
End Function
Public Function StopWave(ByRef SWave As Wave) As Boolean

    Dim NoError As Boolean
    Dim IOFunction As Long
    Dim TempArray(1) As Double
    
    If SWave.IO = IOOUTPUT Then
    
        IOFunction = AOFUNCTION
        
    Else
    
        IOFunction = AIFUNCTION
        
    End If

    NoError = SWave.ManageBackgroundProcess(IOFunction, _
                                            TempArray, _
                                            "Sine Wave", _
                                            True, _
                                            False)

    StopWave = NoError

End Function


Public Function DoRamp(ByRef UpWave As Wave, _
                       ByRef DownWave As Wave, _
                       ByRef MonitorWave As Wave, _
                       ByRef BaselineWave As Wave, _
                       Optional ByVal HangeTime As Long = 0, _
                       Optional doSineWaveToo As Boolean = False, _
                       Optional Verbose As Boolean = False, _
                       Optional ClippingTest As Boolean = False) As Long

    Dim i As Long

    Dim RampStatus As Long
    Dim ULStats As Long
    Dim gainArray(1) As Long
            
    Dim Status As Integer
    Dim CurCount As Long
    Dim CurIndex As Long
    Dim DataValue As Integer
        
    Dim AFMonitorArray() As Double
    Dim BaselineArray() As Double
    Dim NumChannels As Long
    Dim TempArray() As Double
    Dim RMS_array() As Double
    Dim SineFitArray() As String
    Dim RampPoints As Long
            
    Dim NoError As Boolean
    Dim ProcessDone As Boolean
            
    Dim BaselineBuffer As Long
    Dim doContinue As Boolean
    Dim SumZero() As Double
    Dim BaselineAvgs() As Double
    
    Dim CurTime
    
    'Set number of channels to monitor analog inputs on
    NumChannels = 4
    
    'Initialize SineFitArray so VB doesn't send a bad/empty array-var error
    ReDim SineFitArray(1)
    SineFitArray(0) = ""

    'Double Check that UpWave & UpWave share the same board, output channel, and peak voltage
    If UpWave.BoardUsed.BoardNum <> DownWave.BoardUsed.BoardNum Then
    
        'Ramp Up and Ramp Down have been assigned to two separate boards!
        MsgBox "Ramp Up and Ramp Down have been assigned to two separate boards!" & _
                vbNewLine & "Ramp Up Board Num = " & Str(UpWave.BoardUsed.BoardNum) & _
                vbNewLine & "Ramp Down Board Num = " & Str(UpWave.BoardUsed.BoardNum)

        DoRamp = -1
        Exit Function
        
    End If
    
    If UpWave.Chan.ChanNum <> DownWave.Chan.ChanNum Then
    
        'Ramp Up and Ramp Down have been assigned to two output channels!
        MsgBox "Ramp Up and Ramp Down have been assigned to two separate boards!" & _
                vbNewLine & "Ramp Up Output Channel Num = " & Str(UpWave.Chan.ChanNum) & _
                vbNewLine & "Ramp Down Output Channel Num = " & Str(UpWave.Chan.ChanNum)

        DoRamp = -1
        Exit Function
        
    End If

    'Check Inputs against the output board parameters
    If Abs(UpWave.PeakVoltage) > 5 Or _
        UpWave.PeakVoltage < 0 Or _
        Abs(DownWave.PeakVoltage) > 5 Or _
        DownWave.PeakVoltage < 0 _
    Then
    
        MsgBox "Bad Peak Voltage setting!" & vbNewLine & _
                Trim(Str(UpWave.PeakVoltage)) & " volts is outside of the (0 - 10) volt peak range"
                
        DoRamp = 10
        Exit Function

    End If

    'Disable the Start Ramp button until the end of this program!
    cmdStartRamp.Enabled = False

    'Determine number of up and down points that there'll be
    UpWave.NumPoints = CLng(UpWave.IORate * UpWave.Duration / 1000)
    DownWave.NumPoints = CLng(DownWave.IORate * DownWave.Duration / 1000)
    
'------Board Specific code - this code is directly specific to the Measurement Computing-----
'------DAS-PCI-6030 board setup and may not be applicable to other board setups--------------
    
    'Need to make sure that the TTL switch for the DAS-Board output is set to
    'output the AF Ramp signal instead of the ARM signal
    'TTL switch is on the DAS-PCI board (UpWave board), on Chan Three, and needs
    'to be set to a zero value.
    DigitalOutput UpWave.BoardUsed.BoardNum, _
                    3, _
                    0

'--------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------
    
    'If the user has not set the Verbose option, then don't need to waste
    'time taking a baseline of the monitor channel
    If Verbose Then
    
        With BaselineWave
            
            'Redminesion the BaselineArray and TempArray arrays to hold 1000 points
            .NumPoints = 1000
            .StartPoint = 0
            .CurrentPoint = 0
            
            ReDim BaselineArray(.NumPoints)
            ReDim TempArray(.NumPoints)
            
            'Before starting the ramp, sample the voltage for ~1000 points on the
            'Monitor wave channel and get a baseline - an average minimum voltage
            'this will be used to re-zero the data when it gets written to file
                
            'Run Baseline scan in the background
            NoError = .ManageBackgroundProcess(AIFUNCTION, _
                                               BaselineArray, _
                                               "Baseline monitor channel scan", _
                                               False, _
                                               False, _
                                               NumChannels)
            
            ProcessDone = False
            
            'Loop until baseline scan is done
            Do
                
                'Capture status of baseline scan in ProcessDone variable
                NoError = .GetBackgroundProcessStatus(AIFUNCTION, _
                                                      TempArray, _
                                                      ProcessDone, _
                                                      "Baseline monitor channel scan")
                                                                  
                If Not NoError Then
                
                    DoRamp = -1
                    
                    Exit Function
                    
                End If
                
            Loop Until ProcessDone
                                                                    
            
            If Not NoError Then
            
                DoRamp = -1
            
                Exit Function
                
            End If
              
            'Kill the baseline process and load the data from the memory buffer
            'into the BaselineArray
            NoError = BaselineWave.ManageBackgroundProcess(AIFUNCTION, _
                                                           BaselineArray, _
                                                           "Baseline monitor channel scan", _
                                                           True, _
                                                           True, _
                                                           NumChannels)
            
            If Not NoError Then
            
                DoRamp = -1
                
                Exit Function
                
            End If
                    
            'ReDim the SumZero and BaselineAvg arrays
            ReDim SumZero(NumChannels)
            ReDim BaselineAvgs(NumChannels)
                    
            'Zero sums and averages over all channels
            For j = 0 To NumChannels - 1
            
                SumZero(j) = 0
                BaselineAvgs(j) = 0
                
            Next j
                                                
            'Now get average value for the minimum voltage
            For i = 0 To UBound(BaselineArray) - 1 Step NumChannels
            
                'Sum the baseline signals over all input channels, separately
                For j = 0 To NumChannels - 1
                    
                    SumZero(j) = SumZero(j) + BaselineArray(i + j)
                
                Next j
                
            Next i
        
            'Get average values by dividing the sums by the total number of
            'baseline points over all the channels
            For j = 0 To NumChannels - 1
            
                BaselineAvgs(j) = SumZero(j) / (UBound(BaselineArray) / NumChannels)
                                
            Next j
                
        End With
            
    End If
       
'---DEBUG ONLY!!--------------------------------------------------------------
'-----------------------------------------------------------------------------
'   (Mar 2010, I Hilburn)
'
'   Starting analog input monitoring of the LC circuit through the Ammeter
'   prior to the generation of the sine wave.  There's a 0.5 second pause in
'   between starting the input process and starting the sine wave
'   Also, this code increases the number of points (MonitorWave.NumPoints) in
'   the monitor wave memory buffer.
'-----------------------------------------------------------------------------

    '1st check to see if the user has selected the verbose option
    'if not, then there's no point in starting the monitor scan here
    If Verbose Then
    
        With MonitorWave
    
            'Increase the number of points for the monitor scan
            'Adding 2 seconds worth of points for the 1 second prior
            'to the ramp (0.5 before the sine wave, 0.5 after, 0.5 after
            'the ramp down has finish, and 0.5 after the sine wave has been
            'stopped
            .NumPoints = .NumPoints + 2 * .IORate
            
            'Note - Wave.ManageBackgroundProcess handles the task of redimensioning
            'AFMonitorArray internally, so we don't need to do that prior to the
            'function call
            NoError = .ManageBackgroundProcess(AIFUNCTION, _
                                               AFMonitorArray(), _
                                               "AF Ramp Monitor", , , _
                                               NumChannels)
                                         
            'Error Check
            If Not NoError Then
            
                'Crap, exit the ramp function
                DoRamp = -1
                
                Exit Function
                
            End If
            
            'Monitor analog input scan is now running in the background,
            'Wait 0.5 seconds (500 ms)
            PauseTill timeGetTime() + 500
            
        End With
        
    End If
            
'-----------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------
       
'-----------------------------------------------------------------------------
'-----------------------------------------------------------------------------
'   (Mar 2010, I Hilburn)
'
'   The sine wave signal generation used to be done
'   prior to the baseline channel scan.  In order to start the analog input
'   process prior to the turning on of the sine wave, I had to move this
'   chunk of code after the baseline scan.
'-----------------------------------------------------------------------------
      
    'If the user has selected that a sine wave be generated for the ramp
    If doSineWaveToo Then
    
        'Need to make sure that the Ramp reference / control voltage
        'is zero before the sine wave is generatred.  We want to make
        'sure that we don't accidentally send a full strength sine signal
        'through the Crest Audio amplifier into the coils.
        ULStats = UpWave.BoardUsed.AnalogOut(UpWave.Range, _
                                             UpWave.Chan, _
                                             0)
        
        'Error Checking
        If ULStats <> 0 Then
        
            MsgBox "Error Outputing Zero Voltage on Board #: " & UpWave.BoardUsed.BoardName & _
                    " (" & Str(UpWave.BoardUsed.BoardNum) & ")" & _
                    vbNewLine & "Channel #: " & Str(OutChan) & vbNewLine & _
                    "Err: " & Str(ULStats)
            
            DoRamp = ULStats
            Exit Function
            
        End If
        
        'Reset Ramp status
        RampStatus = 0
        
        'Generate sine-wave using user input information from the form
        cmdOutputSineA_Click
        
    End If
            
'----------------------------------------------------------------------------------
'----------------------------------------------------------------------------------
            
'    'Ready for the actual ramp cycle now!
    
    'If user has selected to run a clipping test, then need to dimension the
    'RMS_array to the number of Ramp Points - which is known and the same for
    'the Ramp Up and Ramp Down process.
    If ClippingTest Then
        
        'If we're running a clipping test, then will have a set number of Ramp Up points
        RampPoints = UpWave.NumPoints
        DownWave.PeakVoltage = UpWave.PeakVoltage
        DownWave.MinVoltage = UpWave.MinVoltage
        DownWave.NumPoints = UpWave.NumPoints
        
        'Redimension the RMS_array so that it has the correct number of entries
        ReDim RMS_array(RampPoints + 1, 2)
        
    End If
            
    
'-------Debug Only--------------------------------------------------
'   (Mar 2010, I Hilburn)
'   Wait 0.5 seconds (500 ms) prior to the start of the Ramp Up
'   but after the sine wave has been started on the function generator
'   board
'-------------------------------------------------------------------

    PauseTill timeGetTime() + 500
    
'-------------------------------------------------------------------
        
    'This loop is here in case the Ramp up doesn't reach the target
    'Monitor voltage.  In that case, the RampUp is run again with
    'a 50% increase in the maximum allowed Ramp reference / control
    'voltage sent to the EE shop -10 to +10 amplifier
    Do
    
        'Ramp the voltage up
        RampStatus = MonitoredRampUp(UpWave, _
                                     DownWave, _
                                     MonitorWave, _
                                     SineFitArray, _
                                     RMS_array, _
                                     ClippingTest, _
                                     Verbose, _
                                     NumChannels)
    
    'Error Code = 1234 means user has selected to retry the ramp up with a 50%
    'peak ramp voltage
    Loop Until RampStatus <> 1234
    
    
    'Check for Error
    If RampStatus <> 0 Then
    
        MsgBox "Ramp Up Error!"
        
        DoRamp = RampStatus
        Exit Function
        
    End If
    
       
    If HangeTime <> 0 Then

        'Now pause for HangeTime (in milliseconds)
        'While pausing, run sine fits and on the signal from the monitor channel
        
        'If the user has selected verbose - then monitor the peak voltage while hanging
        If Verbose Then
        
            NoError = MonitoredHangeAtPeak(HangeTime, _
                                           MonitorWave, _
                                           SineFitArray, _
                                           Verbose, _
                                           NumChannels)
                                           
        Else
        
            'PauseTill now works with millisecond values, not with second values
            PauseTill timeGetTime() + HangeTime
            
        End If

    End If
    
    'Need to set the Current Voltage in the DownWave object
    'to the last current voltage in the UpWave Object
    DownWave.CurrentVoltage = UpWave.CurrentVoltage
    DownWave.PeakVoltage = DownWave.CurrentVoltage
        
    'Now time to Ramp Down
    RampStatus = MonitoredRampDown(DownWave, _
                                   MonitorWave, _
                                   SineFitArray, _
                                   RMS_array, _
                                   ClippingTest, _
                                   Verbose)
    
    'Check for Error
    If RampStatus <> 0 Then
            
        MsgBox "Ramp Down Error!"
            
        DoRamp = RampStatus
        Exit Function
        
    End If
    
    'Pause 0.5 seconds before ending the sine wave output
    PauseTill timeGetTime() + 500

    'Start Ramp, and Start Sine Wave buttons will be re-enabled by this function call
    'Kill the sine wave function generator
    Me.cmdStopSineWave_Click
    
    'Pause another 0.5 seconds before stopping the background monitor process
    PauseTill timeGetTime() + 500

    'Terminate the Background Monitor Process for this AF Ramp
    'Pass verbose to the function to choose whether or not to read the
    'buffer into the AFMonitorArray
    NoError = MonitorWave.ManageBackgroundProcess(AIFUNCTION, _
                                                  AFMonitorArray, _
                                                  "monitor AF Ramp", _
                                                  True, _
                                                  Verbose, _
                                                  NumChannels)
                                
    'Error Check
    If Not NoError Then
    
        DoRamp = -1
        
        Exit Function
        
    End If
    
'    Next i
    
    'Check to see if Verbose is set - if so, record AFMonitorArray to data file(s)
    If Verbose Then
        
        CurTime = Now
        FolderName = Format(MonitorWave.PeakVoltage, "0.0##") & "V " & _
                     Format(CurTime, "MM-DD-YYYY HH-MM-SS") & "\"
        
        'Now write the AFMonitorArray to a file
        frmFileSave.MultiRampFileSave AFMonitorArray, _
                                      BaselineAvgs, _
                                      MonitorWave.TimeStep, _
                                      65000, _
                                      FolderName, _
                                      CurTime, _
                                      SineFitArray, _
                                      False, _
                                      True, _
                                      CInt(MonitorWave.IORate / MonitorWave.SineFreq) + 1
                                          
    End If

    
    
End Function
Private Function GhettoRampUp(ByRef UpWave As Wave, ByRef MonitorWave As Wave) As Long

    Dim StrIntervTime
    Dim EndIntervTime
    Dim TimeInterval
    
    Dim OutputMCCCounts As Integer
    Dim OutputStatus As Long
    Dim OutputIndex As Long
    Dim ErrorCode As Long
    Dim i As Long
    Dim NumPoints As Long
    Dim doContinue As Boolean
    Dim TempI As Integer
    Dim UserResp As Long
    
    'Start Output Index off at the first point
    OutputIndex = 1
    
    'Calculate the time interval between the ramp up output points
    TimeInterval = 1 / UpWave.IORate
    
    'Set start point to 1
    UpWave.StartPoint = 1
    
    'Calculate the number of points
    NumPoints = UpWave.NumPoints
           
    For i = 1 To NumPoints
        
        With UpWave
        
            'Calculate ramp voltage to send out through the DAQ board
            .CurrentVoltage = .PeakVoltage * i / NumPoints
            
            'Store current point
            .CurrentPoint = i
            
            'Set the start time for this interval
            StrIntervTime = Timer
            
            'Set the end time for this interval
            EndIntervTime = StrIntervTime + TimeInterval
            
            'Output the DAQ Board counts
            .BoardUsed.AnalogOut .Range, _
                                    .Chan, _
                                    .CurrentVoltage
            
            'Error Check
            If ErrorCode <> 0 Then
            
'---------------Debug Code only!----------------------------------------------------------------
                'Tell the user
                UserResp = MsgBox("Unable to output analog Ramp Up point." & _
                            "User Input required!" & vbNewLine & _
                            "Point #: " & Trim(Str(i)) & vbNewLine & _
                            "Voltage: " & Trim(Str(OutputVoltage)) & vbNewLine & vbNewLine & _
                            "Board: " & .BoardUsed.BoardName & " (" & _
                            Trim(Str(.BoardUsed.BoardNum)) & vbNewLine & _
                            "Channel: " & .Chan.ChanName & " (" & _
                            Trim(Str(.Chan.ChanNum)) & vbNewLine & vbNewLine & _
                            "Err: " & Trim(Str(ErrorCode)), vbAbortRetryIgnore, _
                            "Ramp Up Error")
                
                'Process user response to error message
                If UserResp = 3 Then
                
                    'User has selected to abort the code
                    
                    'Pass along the error code
                    GhettoRampUp = ErrorCode
                    
                    Exit Function
                    
                ElseIf UserResp = 4 Then
                
                    'User has selected to retry
                    
                    'decrement i by 1 so that this iteration of the for loop will repeat
                    i = i - 1
                    
                End If
                
'-----------------------------------------------------------------------------------------------
                                        
            End If
        
        End With
        
        'Pause until the end of the interval time between points
        PauseTill EndIntervTime
        
    Next i
    
    'We've gone all the way through the RampUp loop
    'We're done!
    
    GhettoRampUp = 0

End Function
Private Function MonitoredHangeAtPeak _
    (ByVal HangeTime As Long, _
     ByRef MonitorWave As Wave, _
     ByRef SineFitArray() As String, _
     Optional Verbose As Boolean = False, _
     Optional NumChannels As Long = 1) As Boolean
    
    Dim ii As Long
    Dim k As Long
    Dim NoError As Boolean
    Dim ProcessDone As Boolean
    
    Dim FitLength As Long
    Dim SineArray() As Double
    Dim TempArray() As Double
    Dim Sine_est() As Double
    Dim Sine_res() As Double
    Dim FitParams(4) As Double
    Dim SizeSineFitArray As Long
    
    Dim RMS As Double
 
    Dim startTime
    Dim IntervTime As Double
    Dim EndIntervTime As Double
    
    'Time at the start of the hange at peak interval, will be checked
    'against later calls of timegettime() to see if the desired hange time interval
    'has been used up
    startTime = timeGetTime()
    
    'Set interval time to 10 milliseconds - this is the time between each check of the signal
    'input monitor memory buffer and run of the sine-fit routine
    IntervTime = 10
    
    'Set the number of points from the monitor input buffer
    'to be fit using the SineFit sub-routine to two periods worth of signal input points
    'remember input rate is at 100 kHz
    FitLength = Round(2 * MonitorWave.PtsPerPeriod, 0)
    
    'Redimension the necessary arrays for the Sine Fit process
    ReDim SineArray(FitLength)
    ReDim TempArray(FitLength * NumChannels)
    ReDim Sine_est(FitLength)
    ReDim Sine_res(FitLength)
        
    Do While startTime + HangeTime < timeGetTime()
        
        'Set the End Interval time for when the loop will be allowed to flip arround to the
        'next loop instance
        EndIntervTime = timeGetTime() + TimeInterval
        
        'Get the necessary number of points from the Monitor Wave's analog input memory buffer
        NoError = MonitorWave.GetBackgroundProcessStatus(AIFUNCTION, _
                                                         TempArray, _
                                                         ProcessDone, _
                                                         "AF Monitor scan", _
                                                         True, _
                                                         FitLength, _
                                                         NumChannels)
        
        If NoError Then
        
            'Only do the sine fit if there was no error retrieving values from the monitor
            'memory buffer.  Otherwise, do nothing - just let the time run out on the hange time
            'and start the monitored Ramp Down function.
        
            'Counter for Sine Array element index
            ii = 0
            
            For k = 0 To FitLength * NumChannels - 1 Step NumChannels
            
                SineArray(ii) = TempArray(k)
            
                ii = ii + 1
            
            Next k

            With MonitorWave
            
                'Error trapping
                On Error GoTo BadSineFit:
                
                'Now have 100 points, can dump them into the Sine Fit program
                SineFit SineArray(), _
                        .TimeStep, _
                        .SineFreq, _
                        FitParams(), _
                        Sine_est(), _
                        Sine_res(), _
                        RMS
                        'SineStream
                
                'Turn off error handling
                On Error GoTo 0
                
                'If Verbose setting is true, record the fit parameters to file
                If Verbose Then
    
                    SizeSineFitArray = UBound(SineFitArray)
                    
                    SizeSineFitArray = SizeSineFitArray + 1
                    ReDim Preserve SineFitArray(SizeSineFitArray)
                                                                
                    SineFitArray(i) = Trim(Str(.CurrentPoint - FitLength)) & "," & _
                                      Format(.TimeStep * _
                                                (.CurrentPoint - FitLength), "0.0#####") & "," & _
                                      Trim(Str(FitParams(0))) & "," & _
                                      Trim(Str(FitParams(1))) & "," & _
                                      Trim(Str(.SineFreq)) & "," & _
                                      Trim(Str(FitParams(2))) & "," & _
                                      Trim(Str(FitParams(3))) & "," & _
                                      Trim(Str(RMS)) & "," & _
                                      Trim(Str(.IORate)) & "," & _
                                      Trim(Str(UpWave.CurrentVoltage)) & "," & _
                                      Trim(Str(CurRampCounts))
                                                        
                End If
                
            End With
            
BadSineFit:
                'This is after all the logging of values from above, just go into
                'the next instance of the monitor / sine-fit loop as if this
                'aborted sine fit never happened

        End If
        
        'Pause till the interval time has stopped
        PauseTill EndIntervTime
        
    Loop
    
    'Set return value = No error
    MonitoredHangeAtPeak = True
     
End Function
     

Private Function MonitoredRampUp _
    (ByRef UpWave As Wave, _
     ByRef DownWave As Wave, _
     ByRef MonitorWave As Wave, _
     ByRef SineFitArray() As String, _
     ByRef RMS_array() As Double, _
     Optional ClippingTest As Boolean = False, _
     Optional Verbose As Boolean = False, _
     Optional NumChannels As Long = 1) As Long

    Dim StrIntervTime
    Dim EndIntervTime
    Dim TimeInterval
    
    Dim ErrorCode As Long
    Dim i As Long
    Dim ii As Long
    Dim j As Long
    Dim k As Long
    Dim m As Long
    Dim CurRampCounts As Long
    Dim UserResp As Long
    Dim BadFit As Boolean
    Dim NoError As Boolean
    Dim ProcessDone As Boolean
    
    Dim FitLength As Long
    Dim SineArray() As Double
    Dim TempArray() As Double
    Dim TempArray2(1) As Double
    Dim Sine_est() As Double
    Dim Sine_res() As Double
    Dim FitParams(4) As Double
    Dim OldFitParams(4) As Double
    Dim SizeSineFitArray As Long
    
    Dim RMS As Double
    
    Dim Delta_V As Double
    Dim Lambda As Double
    Dim Threshold, Threshold2 As Double
    Dim RampSlowed As Boolean
    Dim FirstFit As Boolean
    Dim PeakReached As Boolean
    
    Dim time, ElapsedTime
    
    'Set the number of points from the monitor input buffer
    'to be fit using the SineFit sub-routine
    FitLength = Round(2 * MonitorWave.PtsPerPeriod, 0)
    
    'Redimension the necessary arrays for the Sine Fit process
    ReDim SineArray(FitLength)
    ReDim TempArray(FitLength * NumChannels)
    ReDim Sine_est(FitLength)
    ReDim Sine_res(FitLength)
    
    'Initialize FirstFit to True
    FirstFit = True
    
    'Initialize PeakReached to False
    PeakReached = False
    
    'Set flag indicating if the ramp up has been slowed to false - ramp rate will
    'start at full speed
    RampSlowed = False
    
    'We want to ramp up until the peak voltage has been reached
    'This loop may be obsolete now given the loop around the Monitored Ramp Up
    'function call in the DoRamp function.
    '
    'This loop may be removed in the near future (Mar 2010, I Hilburn)
    Do While Not PeakReached
        
        'If the monitor process's memory buffer has not yet been
        'allocated, then the background analog input scan for this
        'process has not been started yet.
        If MonitorWave.BufferAlloc = False Then
        
            'Start the background analog input process
            NoError = MonitorWave.ManageBackgroundProcess(AIFUNCTION, _
                                                          TempArray2(), _
                                                          "Monitor AF Ramp", , , _
                                                          NumChannels)
                                                            
            'Error check - if the start process failed, then
            'cancel the AF ramp and exit the function
            If Not NoError Then
            
                MonitoredRampUp = -1
                
                Exit Function
                
            End If
            
        End If
        
        'Set Delta_V - voltage increment value
        Delta_V = UpWave.PeakVoltage / UpWave.NumPoints
        
        'Initialize current voltage of UpWave to zero
        UpWave.CurrentVoltage = 0
        
        'Counter to measure the number of points
        i = 0
               
        'Set start point to i
        UpWave.StartPoint = i
               
        '(March 2010, I Hilburn)
        'Ramp Up to Low Scan Voltage (UpWave.MinVoltage)
        'prior to starting the monitor-feedback ramp process
        'Note - the background analog input monitor process
        'has already been started, we're just not reading any
        'points from the input buffer until after
        'the Ramp Up reference voltage has been raised up to
        'the Ramp Up minimum voltage.
        '
        'This functionality is used pretty much exclusively for
        'the clipping test, though it could also be used for
        'the Ramp up to high voltages if there's a problem with
        'apparent sine-wave distortion at the very beginning of
        'the Ramp Up
        
        'Need this if...then statement to stop the for loop
        'below from running once with Voltage = 0
        If UpWave.MinVoltage <> 0 Then
        
            For Voltage = 0 To UpWave.MinVoltage Step 0.002
            
                StrIntervTime = timeGetTime()
            
                'Output Voltage through UpWave board & chan
                UpWave.BoardUsed.AnalogOut UpWave.Range, _
                                           UpWave.Chan, _
                                           Voltage
                                        
                
                'Store current UpWave output voltage
                UpWave.CurrentVoltage = Voltage
                    
                'Store Current point of Ramp Up
                UpWave.CurrentPoint = i
                    
                'iterate i
                i = i + 1
                    
                'Wait 1 millisec before next ramp up increment
                PauseTill StrIntervTime + 1
                
            Next Voltage
          
        End If
          
        'Store start point of monitored Ramp up
        UpWave.StartPoint = i
        
        'Start Ramp Up Analog point output For loop
        Do While UpWave.CurrentVoltage <= UpWave.PeakVoltage
            
            With UpWave
            
                'Calculate ramp voltage to send out through the DAQ board
                .CurrentVoltage = .CurrentVoltage + Delta_V
                   
                'Store current point
                .CurrentPoint = i
                
                'Debug.Print Trim(Str(.CurrentPoint)) & ", " & Trim(Str(.CurrentVoltage)) & " Volts"
                        
                'Set the start time for this interval
                StrIntervTime = timeGetTime()
                
                'Set the end time for this interval
                EndIntervTime = StrIntervTime + UpWave.TimeStep * 1000
                
                'Output the DAQ Board counts
                ULStats = .BoardUsed.AnalogOut(.Range, _
                                        .Chan, _
                                        .CurrentVoltage)
                
                'Error Check
                If ULStats < 0 Then
                
    '---------------Debug Code only!----------------------------------------------------------------
                    'Tell the user
                    UserResp = MsgBox("Unable to output analog Ramp Up point." & _
                                "User Input required!" & vbNewLine & _
                                "Point #: " & Trim(Str(i)) & vbNewLine & _
                                "Voltage: " & Trim(Str(OutputVoltage)) & vbNewLine & vbNewLine & _
                                "Board: " & .BoardUsed.BoardName & " (" & _
                                Trim(Str(.BoardUsed.BoardNum)) & _
                                vbNewLine & vbNewLine & "MCC Err: " & Trim(Str(-1 * ULStat)), _
                                vbAbortRetryIgnore, _
                                "Ramp Up Error")
                    
                    'Process user response to error message
                    If UserResp = 3 Then
                    
                        'User has selected to abort the code
                        
                        'Pass along the error code
                        MonitoredRampUp = ErrorCode
                        
                        Exit Function
                        
                    ElseIf UserResp = 4 Then
                    
                        'User has selected to retry
                        
                        'decrement i by 1 so that this iteration of the for loop will repeat
                        i = i - 1
                        
                        'Also need to decrement current voltage by the delta-voltage
                        .CurrentVoltage = .CurrentVoltage - Delta_V
                        
                    End If
                    
                End If
             
            End With
            
            'Else, ULStats holds the MCC Counts sent through the Analog Output Ramp control
            'channel, this needs to be stored to another local variable so that it
            'can be written to file
            CurRampCounts = ULStats
                        
    '-----------------------------------------------------------------------------------------------
    '       Monitor Peak Voltage code - using the sine fit algorithm
    '-----------------------------------------------------------------------------------------------
            
            time = timeGetTime()
            
            'Get the necessary number of points (for the current number of channels
            'being monitored) to get FitLength points from the monitor input buffer
            NoError = MonitorWave.GetBackgroundProcessStatus(AIFUNCTION, _
                                                             TempArray(), _
                                                             ProcessDone, _
                                                             "Monitor AF Ramp", _
                                                             True, _
                                                             FitLength, _
                                                             NumChannels)
                                                   
            
            If Not NoError Then
            
                'Something bad has happened, an error message has already been sent to the
                'user, do a ghetto ramp down and end this function
                
                'Set peak voltage to ramp down from equal to current voltage of the
                'aborted ramp-up process
                DownWave.PeakVoltage = UpWave.CurrentVoltage
                DownWave.CurrentVoltage = DownWave.PeakVoltage
                
                'Do Ramp down
                ULStats = GhettoRampDown(DownWave)
                
                'error check
                If ULStats <> 0 Then
                
                    'Crap - now need to send another error message to the user to
                    'let them know that they need to crank down the power on the
                    'Audio Amplifier
                    Err.Raise ULStats, _
                              "GhettoRampDown", _
                              "Emergency Ramp Down during Ramp Cycle to a monitor " & _
                              "voltage of: " & Trim(Str(MonitorWave.PeakVoltage)) & _
                              " Volts, has failed due to an MCC comm error." & _
                              vbNewLine & vbNewLine & "Please turn down the gain on the" & _
                              "Crest Audio Amplifier manually!!"
                    
                End If
            
                'Whether or not the ramp down worked, need to exit this function
                MonitoredRampUp = -1
                    
                Exit Function
                
            End If
                    
            'Now load every other value from the Temp Array
            'which contains two channels worth of data into the Sine Array
            
            'Counter for Sine Array element index
            ii = 0

            For k = 0 To FitLength * NumChannels - 1 Step NumChannels

                SineArray(ii) = TempArray(k)

                ii = ii + 1

            Next k
                    
            With MonitorWave
            
                If Not FirstFit Then
                
                    'Not the first fit, set OldFitParams to last cycles fit parameters
                    OldFitParams(0) = FitParams(0)
                    OldFitParams(1) = FitParams(1)
                    OldFitParams(2) = FitParams(2)
                    OldFitParams(3) = FitParams(3)
                    
                End If
                
                'Error trapping
                On Error GoTo BadSineFit:
                
                'Now have 100 points, can dump them into the Sine Fit program
                SineFit SineArray(), _
                        .TimeStep, _
                        .SineFreq, _
                        FitParams(), _
                        Sine_est(), _
                        Sine_res(), _
                        RMS
                        'SineStream
                
                'If Verbose setting is true, record the fit parameters to file
                If Verbose Then
    
                    SizeSineFitArray = UBound(SineFitArray)
                    
                    If SizeSineFitArray = 1 And FirstFit Then
                    
                        'Do nothing
                                          
                    Else
                    
                        'Resize the array so that it's one size larger
                        SizeSineFitArray = SizeSineFitArray + 1
                        ReDim Preserve SineFitArray(SizeSineFitArray)
                    
                    End If
                        
                    SineFitArray(i) = Trim(Str(.CurrentPoint - FitLength)) & "," & _
                                      Format(.TimeStep * _
                                                (.CurrentPoint - FitLength), "0.0#####") & "," & _
                                      Trim(Str(FitParams(0))) & "," & _
                                      Trim(Str(FitParams(1))) & "," & _
                                      Trim(Str(.SineFreq)) & "," & _
                                      Trim(Str(FitParams(2))) & "," & _
                                      Trim(Str(FitParams(3))) & "," & _
                                      Trim(Str(RMS)) & "," & _
                                      Trim(Str(.IORate)) & "," & _
                                      Trim(Str(UpWave.CurrentVoltage)) & "," & _
                                      Trim(Str(CurRampCounts))
                                                        
                End If
                        
                If ClippingTest Then
                
                    RMS_array(i - UpWave.StartPoint, 0) = FitParams(1)
                    RMS_array(i - UpWave.StartPoint, 1) = RMS
                    
                    PeakReached = True
                    
                Else
                            
                    If FirstFit Then
                    
                        'If is the first fit, then set FirstFit = False
                        FirstFit = False
                        
                        'Also check for a massive initial distortion of the
                        'sine wave
                        If (Abs(FitParams(0)) > .PeakVoltage _
                                Or Abs(FitParams(1)) > .PeakVoltage) _
                            And Not Abs(FitParams(1)) < 0.01 _
                        Then
                            
                            'First fit is bad, set BadFit Flag
                            BadFit = True
                            
                            
                        Else
                        
                            BadFit = False
                        
                        End If
                        
                        'Record current voltage to the MonitorWave object
                        .CurrentVoltage = FitParams(1)
                                                
                    Else
                        
                        'If fit parameters are still bad, then abort ramp up
                        'ramp down, and then exit the function
                        If (Abs(FitParams(0)) > .PeakVoltage _
                                Or Abs(FitParams(2) - .SineFreq) / .SineFreq > 0.5 _
                                Or FitParams(1) > .PeakVoltage * 2) _
                            And Not FitParams(1) < 0.01 _
                        Then
                            
                            'Set flag that this is a bad sine fit
                            'and should not be used to monitor the sine signal voltage
                            BadFit = True
                            
                        Else
                        
                            BadFit = False
                        
                        End If
                        
                    End If
                    
                    'Store current monitor input signal sine fit amplitude to
                    'MonitorWave object
                    MonitorWave.CurrentVoltage = FitParams(1)
                                
        '                'Have current amplitude - can use to parameterize lambda constant
        '                'for a logarithmic ramp up
        '                Lambda = -1 / i * Log(1 - FitParams(1) / .PeakVoltage)
        '
        '                'Now can use lambda and the derivative of V(t) = 1 - e^(-lambda * t)
        '                'To get the value for Delta-V
        '                Delta_V = Lambda * UpWave.PeakVoltage * Exp(-Lambda * I)
                    
                            
                    'Set Thresholds for Ramp up slow down + end
                    'If delta_V is larger than 10% of the peak voltage, need to
                    'change threshold for slow down and stop of ramp up
                    If Delta_V > 0.1 * .PeakVoltage _
                        And Not RampSlowed _
                    Then
                    
                        Threshold = Delta_V
                        
                    ElseIf Not RampSlowed Then
                    
                        Threshold = 0.1 * .PeakVoltage
                        
                    End If
                    
                    If RampSlowed And Delta_V > 0.001 * .PeakVoltage Then
                    
                        Threshold2 = Delta_V
                        
                    ElseIf RampSlowed Then
                    
                        Threshold2 = 0.001 * .PeakVoltage
                        
                    End If
                    
                    'Rescale threshold for small peak voltages
                    If Threshold < 0.01 Then Threshold = 0.01
                    
                    'Rescale threhold2 for small peak voltages
                    If Threshold2 < 0.002 Then Threshold2 = 0.002
                    
                    If (.PeakVoltage - FitParams(1) <= Threshold) _
                        And .PeakVoltage - FitParams(1) > Threshold2 _
                        And Not RampSlowed _
                        And Not BadFit _
                    Then
        
                        'Indicate that the ramp has been slowed down
                        RampSlowed = True
        
                        'Decrease Delta_V by a factor of 10
                        Delta_V = Delta_V / 10
        
                    ElseIf Not BadFit Then
                    
                        If .PeakVoltage - FitParams(1) <= Threshold2 Then
                        
                            'This should end the do-while loop
                            UpWave.PeakVoltage = UpWave.CurrentVoltage - 100
                            
                            'This should end the larger loop
                            PeakReached = True
                            
                        End If
                        
                    End If
                    
                End If
              
            End With
              
AfterCompareAmplitude:
            
            ElapsedTime = timeGetTime() - time
'            Debug.Print ElapsedTime
                    
            'Iterate i
            i = i + 1
                    
            'Pause until the end of the interval time between points
            PauseTill EndIntervTime
            
        Loop
    
        'Tell user if the peak voltage wasn't reached, and this isn't a clipping test
        'then give user the option to repeat the ramp-up with a higher Ramp peak voltage
        If Not PeakReached And Not ClippingTest Then
        
            UserResp = MsgBox("Peak Voltage not reached during ramp-up." & vbNewLine & _
                    "Peak Voltage = " & Trim(Str(MonitorWave.PeakVoltage)) & _
                    vbNewLine & "Last Voltage = " & Trim(Str(FitParams(1))) & _
                    vbNewLine & vbNewLine & "Retry with a 50% increase in ramp voltage?", _
                    vbRetryCancel, _
                    "AF Ramp Error")
                    
            If UserResp = 2 Then
            
                'User has selected to cancel retrying the ramp up
                PeakReached = True
                
            Else
            
                'Ramp Down from current voltage
                DownWave.CurrentVoltage = UpWave.CurrentVoltage
                DownWave.PeakVoltage = DownWave.CurrentVoltage
                
                GhettoRampDown DownWave
                
                UpWave.PeakVoltage = UpWave.PeakVoltage * 1.5
                DownWave.PeakVoltage = UpWave.PeakVoltage
                
                'Kill Monitor Background process
                NoError = MonitorWave.ManageBackgroundProcess(AIFUNCTION, _
                                                             TempArray2(), _
                                                             "Monitor AF Ramp", _
                                                             True, _
                                                             False)
                                                             
                If Not NoError Then
                
                    MonitoredRampUp = -1
                    
                    Exit Function
                    
                End If
                    
            End If
        
        End If
    
    Loop
       
AfterLastFit:
        
    'If Verbose is true, then need to close the output file now
    If Verbose Then
    
        SineStream.Close
        
    End If
        
    MonitoredRampUp = 0
    
    Exit Function

BadSineFit:

    'Return program flow to normal
    Resume Next

    'Set the flag for a bad sine fit
    BadFit = True
    
    'Goto a line in the code after
    'the post-fit sine fit amplitude VS peak comparison
    GoTo AfterCompareAmplitude:

GetMemoryBuffer:
    
    If Verbose Then
    
        With MonitorWave
    
            SizeSineFitArray = UBound(SineFitArray)
            
            SizeSineFitArray = SizeSineFitArray + 1
            ReDim Preserve SineFitArray(SizeSineFitArray)
    
            SineFitArray(SizeSineFitArray - 1) = _
                        Trim(Str(.CurrentPoint - FitLength)) & "," & _
                        Format(.TimeStep * _
                                (.CurrentPoint - FitLength), "0.0#####") & "," & _
                        Trim(Str(FitParams(0))) & "," & _
                        Trim(Str(FitParams(1))) & "," & _
                        Trim(Str(.SineFreq)) & "," & _
                        Trim(Str(FitParams(2))) & "," & _
                        Trim(Str(FitParams(3))) & "," & _
                        Trim(Str(RMS)) & "," & _
                        Trim(Str(.IORate)) & "," & _
                        Trim(Str(UpWave.CurrentVoltage)) & "," & _
                        Trim(Str(CurRampCounts))
                        
        End With
                                                            
    End If
    
    MonitoredRampUp = 0

End Function
Public Function MonitoredRampDown(ByRef DownWave As Wave, _
                                    ByRef MonitorWave As Wave, _
                                    ByRef SineFitArray() As String, _
                                    ByRef RMS_array() As Double, _
                                    Optional ClippingTest As Boolean = False, _
                                    Optional Verbose As Boolean = False, _
                                    Optional NumChannels As Long = 1) As Long

    Dim StrIntervTime
    Dim EndIntervTime
    
    Dim ULStats As Long
    Dim i, ii As Long
    Dim j As Long
    Dim k As Long
    Dim m As Long
    Dim Voltage As Double
    
    Dim ProcessDone As Boolean
    
    Dim Status As Integer
    Dim CurCount As Long
    Dim CurIndex As Long
    Dim FitLength As Long
    Dim SineArray() As Double
    Dim TempArray() As Double
    Dim Sine_est() As Double
    Dim Sine_res() As Double
    Dim FitParams(4) As Double
    
    Dim RMS As Double
    Dim gainArray(1) As Long
    
    Dim time, ElapsedTime
    
    'Set the number of points from the monitor input buffer
    'to be fit using the SineFit sub-routine
    FitLength = Round(2 * MonitorWave.PtsPerPeriod, 0)
    
    'Redimension the necessary arrays for the Sine Fit process
    ReDim SineArray(FitLength)
    ReDim TempArray(FitLength * NumChannels)
    ReDim Sine_est(FitLength)
    ReDim Sine_res(FitLength)
    
    'Ramp Down from High Scan Voltage to Low Scan Voltage
    'Note - this loop is running backwards, j is decreasing
    'This enables a degree of sublime laziness in comparing
    'the RMS values between the up and down ramps
    'and enables me not to have to change most of the code
    'Debug.Print DownWave.NumPoints & "," & DownWave.Duration
    
    For j = DownWave.NumPoints - 1 To 0 Step -1
        
        With DownWave
        
            'Calculate ramp voltage to send out through the DAQ board
            .CurrentVoltage = .MinVoltage _
                                + (.PeakVoltage - .MinVoltage) * j / (.NumPoints - 1)
               
            'Store current point
            .CurrentPoint = i
            
            'Iterate i
            i = i + 1
            
            'Debug.Print Trim(Str(.CurrentPoint)) & ", " & Trim(Str(.CurrentVoltage)) & " Volts"
                    
            'Set the start time for this interval
            StrIntervTime = timeGetTime()
            
            'Set the end time for this interval
            EndIntervTime = StrIntervTime + DownWave.TimeStep * 1000
            
            'Output the Current Output point
            ULStats = .BoardUsed.AnalogOut(.Range, _
                                           .Chan, _
                                           .CurrentVoltage)
                                           
            'Error Check
            If ULStats < 0 Then
            
'---------------Debug Code only!----------------------------------------------------------------
                'Tell the user
                UserResp = MsgBox("Unable to output analog Ramp Down point." & _
                            "User Input required!" & vbNewLine & _
                            "Point #: " & Trim(Str(i)) & vbNewLine & _
                            "Voltage: " & Trim(Str(OutputVoltage)) & vbNewLine & vbNewLine & _
                            "Board: " & .BoardUsed.BoardName & " (" & _
                            Trim(Str(.BoardUsed.BoardNum)) & _
                            vbNewLine & vbNewLine & "MCC Err: " & Trim(Str(-1 * ULStat)), _
                            vbAbortRetryIgnore, _
                            "Ramp Down Error")
                
                'Process user response to error message
                If UserResp = 3 Then
                
                    'User has selected to abort the code
                    
                    'Pass along the error code
                    MonitoredRampDown = -1 * ULStats
                    
                    Exit Function
                    
                ElseIf UserResp = 4 Then
                
                    'User has selected to retry
                    
                    'decrement i by 1 so that this iteration of the for loop will repeat
                    i = i - 1
                    
                End If
                
            End If
            
                        
            'Else, ULStats holds the MCC Counts sent through the Analog Output Ramp control
            'channel, this needs to be stored to another local variable so that it
            'can be written to file
            CurRampCounts = ULStats
            
        End With
'-----------------------------------------------------------------------------------------------
'       Monitor RMS of the sine fit versus the input monitor signal
'-----------------------------------------------------------------------------------------------
        
        time = timeGetTime()
        
        NoError = MonitorWave.GetBackgroundProcessStatus(AIFUNCTION, _
                                               TempArray, _
                                               ProcessDone, _
                                               "Monitor AF Ramp down", _
                                               True, _
                                               FitLength, _
                                               NumChannels)
                                               
        'Start ii index at zero
        ii = 0
        
        'now extract one channel's worth of points from the temp array
        For k = 0 To FitLength * NumChannels - 1 Step NumChannels
        
            SineArray(ii) = TempArray(k)
            
            ii = ii + 1
            
        Next k
                
        If NoError Then
        
            With MonitorWave
                                                                            
                'Set Time Step
                MonitorWave.TimeStep = 1 / (.IORate)
                                            
                'Error trapping
                On Error GoTo SineFitError:
                    
                    'Now have fit-lengths worth of points, can dump them into the Sine Fit program
                    SineFit SineArray(), _
                            .TimeStep, _
                            .SineFreq, _
                            FitParams(), _
                            Sine_est(), _
                            Sine_res(), _
                            RMS
                            'SineStream
                
                On Error GoTo 0
                
                'If Verbose setting is true, record the fit parameters to file
                If Verbose Then

                    SizeSineFitArray = UBound(SineFitArray)
        
                    SizeSineFitArray = SizeSineFitArray + 1
                    ReDim Preserve SineFitArray(SizeSineFitArray)
    
                    SineFitArray(SizeSineFitArray - 1) = _
                            Trim(Str(.CurrentPoint - FitLength)) & "," & _
                            Format(.TimeStep * _
                                (.CurrentPoint - FitLength), "0.0#####") & "," & _
                            Trim(Str(FitParams(0))) & "," & _
                            Trim(Str(FitParams(1))) & "," & _
                            Trim(Str(.SineFreq)) & "," & _
                            Trim(Str(FitParams(2))) & "," & _
                            Trim(Str(FitParams(3))) & "," & _
                            Trim(Str(RMS)) & "," & _
                            Trim(Str(.IORate)) & "," & _
                            Trim(Str(DownWave.CurrentVoltage)) & "," & _
                            Trim(Str(CurRampCounts))
                            
                End If
            
            End With
                        
        End If
                    
        
        ElapsedTime = timeGetTime() - time
'        Debug.Print ElapsedTime
           
AfterSineFitError:
           
        'Now Compare the RMS values for this point between the Ramp Up
        'and Ramp Down processes.  The current voltage at this j during the Ramp Down
        'should be the same voltage as that at this value of j during the Ramp up
        
        'If the RMS value from the Ramp Up is -10
        'the Ramp down RMS is NOT -10, then keep only the good Ramp Down value
        If ClippingTest Then
'            Debug.Print UBound(RMS_array, 1)
            If RMS_array(j, 1) = -10 And RMS <> -10 Then
                
                'Store just this RMS value and this amplitude
                RMS_array(j, 0) = FitParams(1)
                RMS_array(j, 1) = RMS
                
            ElseIf RMS_array(j, 1) = -10 And RMS = -10 Then
            
                'Both RMS values are crap - at both points sinefit failed
                'Write -10 value to this element of the array.
                'Write -10 as the amplitude as well so that this point will
                'be ignored
                RMS_array(j, 1) = -10
                RMS_array(j, 0) = -10
                
            Else
            
                'Both RMS values are good, average them together
                RMS_array(j, 1) = (RMS_array(j, 1) + RMS) / 2
                RMS_array(j, 0) = (RMS_array(j, 0) + FitParams(1)) / 2
                
                
            End If
            
        End If
        'Pause until the end of the interval time between points
        PauseTill EndIntervTime
        
    Next j
    
    'We've gone all the way through the RampDown loop
    'We're done!
    
    With DownWave
                
        'Ramp Dowm from Low Scan Voltage (DownWave.MinVoltage)
        'prior to end the auto clip test process
        
        'Need this if...then statement to stop the for loop
        'below from running once with Voltage = 0
        If DownWave.MinVoltage <> 0 Then
        
            For Voltage = DownWave.MinVoltage To 0 Step -0.002
            
                StrIntervTime = timeGetTime()
            
                'Output Voltage through UpWave board & chan
                DownWave.BoardUsed.AnalogOut DownWave.Range, _
                                             DownWave.Chan, _
                                             Voltage
                                        
                
                'Store current UpWave output voltage
                DownWave.CurrentVoltage = Voltage
                    
                'Store Current point of Ramp Up
                DownWave.CurrentPoint = i
                    
                'iterate i
                i = i + 1
                    
                'Wait 1 millisec before next ramp up increment
                PauseTill StrIntervTime + 1
                
            Next Voltage
        
        End If
        
    End With
    
    MonitoredRampDown = 0
    
    Exit Function

SineFitError:

    'Return program flow to normal
    Resume Next

    'If Sine Fit algorithm crashes, set RMS value to -10
    'this will cause a noticable blip in the final plot
    RMS = -10

    GoTo AfterSineFitError:

End Function

Public Function GhettoRampDown(ByRef DownWave As Wave) As Long

    Dim StrIntervTime
    Dim EndIntervTime
    Dim TimeInterval
    
    Dim OutputMCCCounts As Integer
    Dim OutputStatus As Long
    Dim OutputIndex As Long
    Dim ErrorCode As Long
    Dim i As Long
    Dim NumPoints As Long
    Dim doContinue As Boolean
    Dim TempI As Integer
    Dim UserResp As Long
    
    'Start Output Index off at the first point
    OutputIndex = 1
    
    'Calculate the time interval between the ramp up output points
    TimeInterval = 1 / DownWave.IORate
    
    'Limit time interval to minimum time interval of 1 millisecond
    If TimeInterval < 1 Then TimeInterval = 1
    
    'Set Start Point to 1
    DownWave.StartPoint = 1
    
    'Calculate the number of points
    NumPoints = DownWave.NumPoints
    
    'Write current voltage on DownWave object (taken from last current voltage
    'on the UpWave object in DoRamp), and make that the Peak Voltage to start
    'ramping down from
    DownWave.PeakVoltage = DownWave.CurrentVoltage
           
    For i = 1 To NumPoints
        
        With DownWave
        
            'Calculate ramp voltage to send out through the DAQ board
            .CurrentVoltage = .PeakVoltage - (.PeakVoltage * i / NumPoints)
            
'            Debug.Print .CurrentVoltage
            
            'Set the start time for this interval
            StrIntervTime = timeGetTime()
            
            'Set the end time for this interval
            EndIntervTime = StrIntervTime + TimeInterval
            
            'Output the DAQ Board counts
            .BoardUsed.AnalogOut .Range, _
                                    .Chan, _
                                    .CurrentVoltage
            
            'Error Check
            If ErrorCode <> 0 Then
            
'---------------Debug Code only!----------------------------------------------------------------
                'Tell the user
                UserResp = MsgBox("Unable to output analog Ramp Down point." & _
                            "User Input required!" & vbNewLine & _
                            "Point #: " & Trim(Str(i)) & vbNewLine & _
                            "Voltage: " & Trim(Str(OutputVoltage)) & vbNewLine & vbNewLine & _
                            "Board: " & .BoardUsed.BoardName & " (" & _
                            Trim(Str(.BoardUsed.BoardNum)) & vbNewLine & _
                            "Channel: " & .Chan.ChanName & " (" & _
                            Trim(Str(.Chan.ChanNum)) & vbNewLine & vbNewLine & _
                            "Err: " & Trim(Str(ErrorCode)), vbAbortRetryIgnore, _
                            "Ramp Down Error")
                
                'Process user response to error message
                If UserResp = 3 Then
                
                    'User has selected to abort the code
                    
                    'Pass along the error code
                    GhettoRampDown = ErrorCode
                    
                    Exit Function
                    
                ElseIf UserResp = 4 Then
                
                    'User has selected to retry
                    
                    'decrement i by 1 so that this iteration of the for loop will repeat
                    i = i - 1
                    
                End If
                
'-----------------------------------------------------------------------------------------------
                                        
            End If
        
        End With
        
        'Pause until the end of the interval time between points
        PauseTill EndIntervTime
        
    Next i
    
    'We've gone all the way through the RampUp loop
    'We're done!
    
    GhettoRampDown = 0

End Function
Private Sub cmdStartAFTuner_Click()
    
    Me.Hide
    frmAFTuner.Show
            
End Sub

Private Sub cmdStartRamp_Click()

    Dim Chan(2) As Channel
    Dim BoardNum(2) As Long
    Dim i As Long
    
    Dim UpWave As Wave
    Dim DownWave As Wave
    Dim MonitorWave As Wave
    Dim BaselineWave As Wave
    Dim Verbose As Boolean
    Dim ClippingTest As Boolean
        
    Set OutChan = Nothing
    Set OutChan = New Channel
    Set UpWave = Nothing
    Set DownWave = Nothing
    Set MonitorWave = Nothing
    Set BaselineWave = Nothing
       
    With cmbOutBoardRamp
            
        BoardNum(0) = .ItemData(.ListIndex) + 1
    
    End With
        
    For i = 0 To 1
        
        Set Chan(i) = Nothing
        Set Chan(i) = New Channel
            
    Next i
    
    With cmbOutChanRamp
        
        Chan(0).ChanName = .Text
        Chan(0).ChanNum = .ItemData(.ListIndex)
        
    End With
    
    With cmbInBoardRamp
    
        BoardNum(1) = .ItemData(.ListIndex) + 1
        
    End With
    
    With cmbInChanRamp
    
        Chan(1).ChanName = .Text
        Chan(1).ChanNum = .ItemData(.ListIndex)
                
    End With
    
    If WaveForms Is Nothing Or WaveForms.count = 0 Then
    
        'CRAP!
        MsgBox "Bad Ramp Wave info!" & vbNewLine & _
                "Garbage values must have been dumped into the Ramp Wave object." & _
                vbNewLine & "Code will end right now!"
        
        End
    
    End If
    
    With WaveForms
    
        For i = 1 To .count
            
            If .Item(i).WaveType = AFRAMPUP Then
                
                Set UpWave = .Item(i)
                
            End If
            
            If .Item(i).WaveType = AFRAMPDOWN Then
            
                Set DownWave = .Item(i)
                
            End If
            
            If .Item(i).WaveType = AFMONITOR Then
            
                Set MonitorWave = .Item(i)
            
            End If
            
            If .Item(i).WaveType = Baseline Then
            
                Set BaselineWave = .Item(i)
                
            End If
            
        Next i
        
    
    End With
    
    If val(txtRampPeakDuration.Text) < 0 Then
    
        'Reset bad peak hold time to zero seconds
        txtRampPeakDuration.Text = "0"
        
    End If
    
    
    With UpWave
    
        Set .Chan = Chan(0)
        Set .BoardUsed = DAQBoards(BoardNum(0))
        .Duration = val(txtRampUpDuration.Text)
        .IORate = val(txtRampRate.Text)
        .TimeStep = 1 / .IORate
        .IOOptions = BACKGROUND
        .Range.RangeType = UNI10VOLTS
        .MinVoltage = 0
        .PeakVoltage = val(txtRampPeakVoltage.Text)
        
        If .PeakVoltage > .Range.MaxValue Then
        
            .PeakVoltage = .Range.MaxValue
            txtRampPeakVoltage.Text = Trim(Str(.PeakVoltage))
            
        End If
        
        'Set up wave trigger to off
        Set .Trig = Nothing
        Set .Trig = New Trigger
        .Trig.TrigType = TRIGOFF
        
        
    End With
    
    With DownWave
    
        Set .Chan = Chan(0)
        Set .BoardUsed = DAQBoards(BoardNum(0))
        .Duration = val(txtRampDownDuration.Text)
        .IORate = UpWave.IORate
        .TimeStep = UpWave.TimeStep
        
        .Range.RangeType = UNI10VOLTS
        .MinVoltage = 0
        .PeakVoltage = val(txtRampPeakVoltage.Text)
        
        If .PeakVoltage > .Range.MaxValue Then
        
            .PeakVoltage = .Range.MaxValue
            txtRampPeakVoltage.Text = Trim(Str(.PeakVoltage))
            
        End If
        
        .IOOptions = BACKGROUND
                
        'Set Trigger to TRIGOFF = -1
        Set .Trig = Nothing
        Set .Trig = New Trigger
        .Trig.TrigType = TRIGOFF
        
    End With
    
    With MonitorWave
        
        Set .Chan = Chan(1)
        Set .BoardUsed = DAQBoards(BoardNum(1))
        .Duration = (UpWave.Duration + DownWave.Duration) * 4 + val(txtRampPeakDuration.Text)
    
        'Record the sine freq that the user has selected on the
        'main page as the sine freq that the monitor is trying to match
        .SineFreq = val(txtFreq)
        
        'Now see if 100 pts per period resolution is below the max analog
        'in rate for the board associated with this wave
        .IORate = .BoardUsed.MaxAInRate \ 2
        .TimeStep = 1 / .IORate
        .PtsPerPeriod = .IORate / .SineFreq
        .NumPoints = .Duration / 1000 * .IORate
        
        .MinVoltage = 0
        .PeakVoltage = val(txtMonitorTrigVolt)
        .Range.RangeType = BIP10VOLTS
        If .PeakVoltage > .Range.MaxValue Then
        
            .PeakVoltage = .Range.MaxValue
            txtMonitorTrigVolt.Text = Trim(Str(.PeakVoltage))
            
        End If
        
        .IOOptions = BACKGROUND
        
        'Set tigger to a particular window range = +- Monitor trigger voltage
        Set .Trig = Nothing
        Set .Trig = New Trigger
        .Trig.TrigType = TRIGOFF        'GATEINWINDOW
'        .Trig.HighThreshold = Val(Me.txtMonitorTrigVolt)
'        .Trig.LowThreshold = -1# * Val(Me.txtMonitorTrigVolt)
        
    End With
    
    With BaselineWave
    
        Set .Chan = MonitorWave.Chan
        Set .BoardUsed = MonitorWave.BoardUsed
        .Duration = -1
        .IORate = MonitorWave.IORate
        .MinVoltage = 0
        .PeakVoltage = -1
        .IOOptions = BACKGROUND
        
        'Set tigger to a particular window range = +- Monitor trigger voltage
        Set .Trig = Nothing
        Set .Trig = New Trigger
        .Trig.TrigType = TRIGOFF
        
    End With
    
    'Now check to see if user has clicked the Verbose? checkbox
    If checkVerbose = Checked Then
    
        Verbose = True
        
    Else
    
        Verbose = False
        
    End If
    
    'Now see if this is an unmonitored ramp cycle for a clipping test
    If checkClippingTest.Value = Checked Then
    
        ClippingTest = True
        
    Else
    
        ClippingTest = False
        
    End If
    
    
    DoRamp UpWave, _
           DownWave, _
           MonitorWave, _
           BaselineWave, _
           val(txtRampPeakDuration.Text), _
           checkDoSineWave, _
           Verbose, _
           ClippingTest
    
    Set UpWave = Nothing
    Set DownWave = Nothing
    Set MonitorWave = Nothing
    Set BaselineWave = Nothing
    Set Chan(0) = Nothing
    Set Chan(1) = Nothing
    
End Sub

Public Sub cmdStopSineWave_Click()
    
    Dim i As Long
    Dim BoardNum As Long
    Dim Chan As Long
    Dim DataValue As Integer
    Dim NoError As Boolean
    
    Dim FuncType As Integer
    Dim Status As Integer
    Dim CurCount As Long
    Dim CurIndex As Long
    
    BoardNum = val(cmbBoardSine.Text)
    Chan = val(cmbChanSine.Text)
    
    If WaveForms Is Nothing Or WaveForms.count = 0 Then
    
        'CRAP!
        MsgBox "Bad Sine Wave info!" & vbNewLine & _
                "Garbage values must have been dumped into the Sine Wave object." & _
                vbNewLine & "Code will end right now!"
        
        End
        
    End If
    
    NoError = True
    
    With WaveForms
    
        For i = 1 To .count
        
            With .Item(i)
            
                If .WaveType = SINEWAVE Then
                
                    'This is the right wave form to pull information out of
                    'Now check that this Wave's board is an actual board object
                    If .BoardUsed Is Nothing Then
                    
                        'CRAP!
                        MsgBox "Bad Sine Board info!" & vbNewLine & _
                                "Garbage values must have been dumped into the Sine Board object." & _
                                vbNewLine & "Code will end right now!"
                        End
                        
                    End If
                    
                                       
                    'Check if Sine Wave channel is a non-nothing channel object
                    If .Chan Is Nothing Then
                    
                        'CRAP!
                        MsgBox "Bad Sine Channel info!" & vbNewLine & _
                                "Garbage values must have been dumped into the Sine Channel object." & _
                                vbNewLine & "Code will end right now!"
                        End
                    
                    End If
                    
                    If .IO = IOINPUT And .Chan.ChanName Like "A*" Then
                    
                        'This is an analog input channel
                        
                        FuncType = AIFUNCTION
                        
                    End If
                    
                    If .IO = IOINPUT And .Chan.ChanName Like "D*" Then
                    
                        'This is a digital input channel
                        FuncType = DIFUNCTION
                        
                    End If
                    
                    If .IO = IOOUTPUT And .Chan.ChanName Like "A*" Then
                    
                        'This is an analog output channel
                        FuncType = AOFUNCTION
                        
                    End If
                    
                    If .IO = IOOUTPUT And .Chan.ChanName Like "D*" Then
                    
                        'This is a digital output channel
                        FuncType = DOFUNCTION
                        
                    End If
                    
                    NoError = NoError And Me.StopWave(WaveForms(i))
                               
                End If
                                
            End With
            
        Next i
        
    End With
    
    If Not NoError Then
    
        'Disable Sine Wave generation button
        Me.cmdOutputSineA.Enabled = False
        
        'Disable Ramp button
        Me.cmdStartRamp.Enabled = False
        
        'Enable Stop Sine Wave button
        Me.cmdStopSineWave.Enabled = True
        
    Else
    
        'Disable Sine Wave generation button
        Me.cmdOutputSineA.Enabled = True
        
        'Disable Ramp button
        Me.cmdStartRamp.Enabled = True
        
        'Enable Stop Sine Wave button
        Me.cmdStopSineWave.Enabled = False
        
    End If
    
End Sub

Private Sub cmdTestFFT_Click()

    Me.Hide
    frm_testRVFFT.Show

End Sub

Private Sub cmdTestFreq_Click()

    frmTestActualFreq.Show

End Sub

Private Sub cmdTestGaussMeter_Click()

    frm908AGaussmeter.Show

End Sub

Private Sub cmdTestRange_Click()

    frmTestRangeConverter.Show
    
End Sub

Private Sub cmdTestSineFit_Click()

    'Need to see if file exists with sine-wave data in it
    Dim fso As New Scripting.FileSystemObject
    Dim SineFile As file
    Dim SineStream As TextStream
    Dim FileExist As Boolean
    Dim PERIOD As Long
    Dim N As Long
    Dim Shift As Long
    Dim i As Long
    Dim TimeStep As Double
    Dim Amplitude As Double
    Dim Freq As Double
    Dim FilePath As String
    Dim TextLine As String
    Dim StrDate As String
    
    Dim SineArray() As Single
    Dim SineEst() As Single
    Dim SineRes() As Single
    Dim SineParam(4) As Double
    Dim RMS As Double
    Dim UseFile As Boolean
    
    UseFile = True
    
    'Set Number of points and points per period
    N = val(InputBox("Input # of data points to fit."))
    PERIOD = 100
    
    'Set Amplitude = 100
    Amplitude = 100
        
    'Re-size the Sine Array to have the necessary number of elements
    ReDim SineArray(N)
    ReDim SineEst(N)
    ReDim SineRes(N)
    
    If UseFile Then
        
        cdlgTestSineFit.FILTER = "*.csv"
        cdlgTestSineFit.InitDir = "C:\Documents and Settings\lab\Desktop\Test MCC Board 11-16-2009\"
        cdlgTestSineFit.flags = cdlOFNFileMustExist
        cdlgTestSineFit.ShowOpen
        FilePath = cdlgTestSineFit.FileName
        
'        Debug.Print FilePath
        
        'Open one of the monitor input signal data files and run the signal
        Set SineFile = fso.GetFile(FilePath)
    
        Set SineStream = SineFile.OpenAsTextStream(ForReading)
        
        Shift = val(InputBox("Start Loading X points into data file." & vbNewLine & _
                                "X = ", _
                                "Data Window Shift", _
                                "0"))
        
        For i = 0 To Shift - 1
        
            SineStream.SkipLine
            
        Next i
        
        For i = 0 To N - 1
        
            TextLine = SineStream.ReadLine
            SineArray(i) = CSng(val(Mid(TextLine, InStr(1, TextLine, ",") + 1)))
            
        Next i
        
        Freq = 917.5
        
        'Set Time Step so IORate = 50 * Freq
        TimeStep = 1 / 100000
        
    Else
    
        'Set Time Step so IORate = 100 KHz
        TimeStep = 1 / 100000
    
        'Calculate sine values to load into SineArray elements
        For i = 0 To N - 1
        
            SineArray(i) = CSng(Amplitude * Sin(2 * Pi * i / PERIOD))
            
        Next i
        
        Freq = 1 / TimeStep / PERIOD
    
    End If
    
    'Initialize RMS and Sine Parameter array to zeros
    For i = 0 To 3
    
        SineParam(i) = 0
        
    Next i
    
    RMS = 0
    
    Dim startTime
    Dim ElapsedTime
    
    startTime = timeGetTime()
    
'----Debug-------------------------------------------------------------------------
'    StrDate = Format(Now, "_MM-DD-YYYY_HH-mm-SS")
'    FilePath = "SineArray" & StrDate & ".csv"
'
'    fso.CreateTextFile (FilePath)
'    Set SineFile = fso.GetFile(FilePath)
'    Set SineStream = SineFile.OpenAsTextStream(ForWriting)
'
'----------------------------------------------------------------------------------
    
    'Call Sine Fit function
    SineFit SineArray(), _
                    TimeStep, _
                    Freq, _
                    SineParam(), _
                    SineEst(), _
                    SineRes(), _
                    RMS
                    'SineStream
                            
    ElapsedTime = timeGetTime() - startTime
    
    MsgBox "Elapsed Time to run SineFit routine = " & Trim(Str(ElapsedTime))
            
    'Now Write Sine Est, and Sine Res to file with RMS on the first line

    'Create the File Name using current date and time
    StrDate = Format(Now, "DD-MM-YYYY_HH_MM")

    'Create Output File for Writing
    fso.CreateTextFile ("C:\Documents and Settings\lab\Desktop\" & _
                        "Test MCC Board 11-16-2009\Test MCC Board\" & _
                        "SineFit_" & StrDate & ".txt")

    Set SineFile = fso.GetFile("C:\Documents and Settings\lab\Desktop\" & _
                        "Test MCC Board 11-16-2009\Test MCC Board\" & _
                        "SineFit_" & StrDate & ".txt")

    Set SineStream = SineFile.OpenAsTextStream(ForWriting)

    'Write RMS and fit Params
    SineStream.WriteLine ("Y-value offset = " & Trim(Str(SineParam(0))))
    SineStream.WriteLine ("Amplitude = " & Trim(Str(SineParam(1))))
    SineStream.WriteLine ("Freq(Hz) = " & Trim(Str(SineParam(2))))
    SineStream.WriteLine ("Phase(Deg) = " & Trim(Str(SineParam(3) / (2 * Pi) * 180)))
    SineStream.WriteLine ("RMS = " & Trim(Str(RMS)))


    'Now Write data in format:
    'i,Sine Estimate(i),Sine Residual(i)
    For i = 0 To N - 1

        TextLine = Trim(Str(i)) & "," & Trim(Str(SineArray(i))) & "," & _
                    Trim(Str(SineEst(i))) & "," & Trim(Str(SineRes(i)))

        SineStream.WriteLine (TextLine)

    Next i

    SineStream.Close
        
End Sub

Private Sub form_resize()
    'Me.Height = 4260
    'Me.Width = 3105
End Sub

Private Sub cmdClose_Click()
    Me.Hide
    Form_Unload 1
    
    End
End Sub

Private Sub Form_Load()

    'In final version of this code added to the VB Paleomag program
    'These initialization functions will be called
    'In the login phase of the program and will pull their
    'Values from the paleomag.ini file

    Dim i As Integer
    Dim j As Integer
    Dim RampRange As Range
    Dim SineRange As Range
    Dim CurChannel As Channel
    Dim N As Integer
    
    Set RampRange = Nothing
    Set RampRange = New Range
    Set SineRange = Nothing
    Set SineRange = New Range
    RampRange.RangeType = UNI10VOLTS
    SineRange.RangeType = BIP10VOLTS
    
    Set CurChannel = Nothing
    Set CurChannel = New Channel

    AFSystem = "MCC"

    'Set acceptable zero level to 10 mV
    AcceptableZero = 0.01
    
    'Set Interval at which Timer operates to 1 micro-s
    timeBeginPeriod 1
    
    'Initialize Boards & Channels
    Initialize_Boards
        
    'Initialize Waves
    Initialize_Waves
    
    'Lock points per period box, it needs to be set by the board IO Rate
    txtPtsPerPeriod.Enabled = False
    
    'Now Load the Possible Board Values
    'To the Sine Wave Generator Form Window
    
    cmbBoardA.Clear
    cmbBoardD.Clear
    cmbBoardSine.Clear
    cmbInBoardRamp.Clear
    cmbOutBoardRamp.Clear
    
    'Set Un-monitored /clipping test check-box to unchecked.
    '(Clipping test = unmonitored ramp cycle)
    Me.checkClippingTest.Value = Unchecked
   
    'Lock Board Name text boxes - these are not for user editing
    txtBoardNameA.locked = True
    txtBoardNameD.locked = True
    
    With DAQBoards
    
        For i = 1 To .count
        
            With .Item(i)
        
                If .AOutChannels.count <> 0 Or .AInChannels.count <> 0 Then
                
                    cmbBoardA.AddItem (Trim(Str(.BoardNum)))
                    cmbBoardA.ItemData(cmbBoardA.NewIndex) = .BoardNum
                    
                End If
                
                If .DOutChannels.count <> 0 Or .DInChannels.count <> 0 Then
                
                    cmbBoardD.AddItem (Trim(Str(.BoardNum)))
                    cmbBoardD.ItemData(cmbBoardD.NewIndex) = .BoardNum
                    
                    
                End If

                If InStr(1, .BoardFunction, SIGNALGENERATOR) > 0 Then
                
                    'If SIGNALGENERATOR is one of board's functions, .BoardFunction will be odd
                    cmbBoardSine.AddItem (Trim(Str(.BoardNum)))
                    cmbBoardSine.ItemData(cmbBoardSine.NewIndex) = .BoardNum
                    
'------------------Board & Channel specific code-----------------------------------------
'
'   (March 2010, I Hilburn) - code put in to zero the two analog output ports
'   on the USB-16HSHS-2 board that is specific to our MCC AF System setup
'   Specific range used as well (BIP10VOLTS: -10 to +10 Volts)
'----------------------------------------------------------------------------------------
                    
                    For j = 1 To 2
                    
                        'Set the two analog out channels to 0
                        Set CurChannel = .AOutChannels.Item(j)
                    
                        'Output an analog zero voltage to the selected channel
                        'on the USB-16HSHS-2 board
                        ULStats = DAQBoards.Item(i).AnalogOut(SineRange, _
                                                              CurChannel, _
                                                              0)
                                                    
                        If ULStats <> 0 Then
                        
                            MsgBox "Error zeroing signal generator channel." & vbNewLine & _
                                    "Board = " & DAQBoards.Item(i).BoardName & " (" & _
                                    Trim(Str(DAQBoards.Item(i).BoardNum)) & ")" & vbNewLine & _
                                    "Channel = " & CurChannel.ChanName & " (" & _
                                    Trim(Str(CurChannel.ChanNum)) & ")" & vbNewLine & vbNewLine & _
                                    "MCC Err: " & Trim(Str(ULStats))
                                    
                        End If
                        
                    Next j
                                        
                End If
                
                If InStr(1, .BoardFunction, AFRAMP) > 0 Then
                
                    cmbOutBoardRamp.AddItem (Trim(Str(.BoardNum)))
                    cmbOutBoardRamp.ItemData(cmbOutBoardRamp.NewIndex) = .BoardNum
                    
                    'Set all analog out channels on the board to zero
                    For j = 1 To .AOutChannels.count
                
                        Set CurChannel = .AOutChannels.Item(i)
                
                        DAQBoards.Item(i).AnalogOut RampRange, _
                                                    CurChannel, _
                                                    0
                
                    Next j
                    
                End If
                
                If InStr(1, .BoardFunction, MONITOR) > 0 Then
                
                    'MONITOR is one of board's functions at .BoardFunction values greater than 8
                    'a .BoardFunction value > 14 is not valid
                    cmbInBoardRamp.AddItem (Trim(Str(.BoardNum)))
                    cmbInBoardRamp.ItemData(cmbInBoardRamp.NewIndex) = .BoardNum
                                        
                End If
                
            End With
        
        Next i
    
    End With

    'Set Board combo box values for making debug go faster
    cmbBoardSine.ListIndex = 0
    cmbInBoardRamp.ListIndex = 0
    cmbOutBoardRamp.ListIndex = 0

    cmbBoardSine_Click
    cmbInBoardRamp_Click
    cmbOutBoardRamp_Click
    

    'All the channels in each board have been loaded, so can now set switch on
    'TTL line so analog output from the PCI-DAS board is set to AF and not ARM
    For i = 1 To WaveForms.count
    
        With WaveForms
        
            If .Item(i).WaveType = ARMAFTTL Then
            
                DigitalOutput .Item(i).BoardUsed.BoardNum, _
                              .Item(i).Chan.ChanNum, _
                              0
                              
            End If
            
        End With
        
    Next i

    'Just for now, store channels in sine and ramp user controls so can test code faster
    cmbChanSine.ListIndex = 0
    cmbInChanRamp.ListIndex = 1
    cmbOutChanRamp.ListIndex = 1
    
    
    
    'Just for now, store values that are testing in user controls so can test code faster
    Me.txtFreq = 917.5
    Me.txtAmplitude = 10
    Me.txtDuration = 1000
    Me.txtSineIORate = 1000000
    'Update pts per period value
    txtSineIORate_Change
    Me.txtRampUpDuration = 1000
    Me.txtRampDownDuration = 1000
    Me.txtRampRate = 100
    Me.txtRampPeakVoltage = 0.2
    Me.txtMonitorTrigVolt = 1
    Me.checkDoSineWave = Checked
    
End Sub


Public Function AnalogInput(ByVal BoardNum As Long, ByVal Chan As Long, ByRef DataValue As Integer, ByVal gain As Long) As Double
    If NOCOMM_MODE Then Exit Function

    Dim ULStat As Long
    Dim engUnits As Single

    txtRawA = vbNullString
    txtEngA = vbNullString

    cmbBoardA.ListIndex = BoardNum
    cmbBoardA_Click
    
    For i = 0 To cmbChanAIn.ListCount - 1
    
        If cmbChanAIn.ItemData(i) = Chan Then
        
            cmbChanAIn.ListIndex = i
            
        End If
        
    Next i
    
    ' Collect the data with cbAIn%()

    '  Parameters:
    '    BoardNum    :the number used by CB.CFG to describe this board
    '    Chan       :the input channel number
    '    Gain       :the gain for the board.
    '    DataValue%  :the name for the value collected
    
    ULStat = cbAIn(BoardNum, Chan, gain, DataValue%)
    If ULStat = 30 Then MsgBox "Change the Gain argument to one supported by this board.", 0, "Unsupported Gain"
    If ULStat <> 0 Then
        txtEngA = "ERR: " & Trim(Str(ULStat))
        Exit Function
    End If
   
    cmbChanD = Str$(Chan)
    txtRawA = Str$(DataValue)
   
    ULStat = cbToEngUnits(BoardNum, gain, DataValue, engUnits)
    If ULStat <> 0 Then
        txtEngA = "ERR: " & Trim(Str(ULStat))
        Exit Function
    End If

    txtEngA = Str$(engUnits)

End Function

Public Sub AnalogOutput(ByVal BoardNum, ByVal Chan As Long, ByVal engUnits As Double, ByVal Range As Long)

    If NOCOMM_MODE Then Exit Sub

    Dim ULStat As Long
    Dim DataValue As Integer

    cmbBoardA.ListIndex = BoardNum
    cmbBoardA_Click
    
    For i = 0 To cmbChanAOut.ListCount - 1
    
        If cmbChanAOut.ItemData(i) = Chan Then
        
            cmbChanAOut.ListIndex = i
            
        End If
        
    Next i

     ' send the digital output value to D/A Chan with cbAOut%()

   ' Parameters:
   '   BoardNum    :the number used by CB.CFG to describe this board
   '   Chan%       :the D/A output channel
   '   Range%      :ignored if board does not have programmable range
   '   DataValue%  :the value to send to Chan%
   
   
   ULStat = cbFromEngUnits(BoardNum, Range, engUnits, DataValue%)
   If ULStat <> 0 Then
        txtEngA = "ERR: " & Trim(Str(ULStat))
        Exit Sub
   End If
         
   ULStat = cbAOut(BoardNum, Chan, Range, DataValue%)
   If ULStat <> 0 Then
        txtEngA = "ERR: " & Trim(Str(ULStat))
        MsgBox "Analog Output Error on Chan = " & Trim(Str(Chan)) & vbNewLine & _
                "Error Number = " & Trim(Str(ULStat))
        Exit Sub
   End If

    cmbChanD = Str$(Chan)
    txtRawA = Str$(DataValue)
    txtEngA = Str$(engUnits)

End Sub


Public Sub DigitalInput(ByVal BoardNum As Long, ByVal PortNum As Long, ByRef DataValue As Integer)

    Dim ULStat As Long

    If NOCOMM_MODE Then Exit Sub

    cmbBoardD.ListIndex = BoardNum
    cmbBoardD_Click
    
    For i = 0 To cmbChanDIn.ListCount - 1
    
        If cmbChanDIn.ItemData(i) = PortNum Then
        
            cmbChanDIn.ListIndex = i
            
        End If
        
    Next i
    
    With DAQBoards.Item(BoardNum + 1)
    
        If .BoardName = "PCI-DAS6030" Then
        
            'Need to use AUXPORT and portnum = bitnum
            ULStat = cbDConfigBit(BoardNum, AUXPORT, PortNum, DigitalIn)
            
            If ULStat <> 0 Then
                txtEngD = "ERR: " & Trim(Str(ULStat))
                Exit Sub
            End If
            
            'Now use bit out command to write to port
            ULStat = cbDBitIn(BoardNum, AUXPORT, PortNum, DataValue)
            
            If ULStat <> 0 Then
                txtEngD = "ERR: " & Trim(Str(ULStat))
                Exit Sub
            End If
            
            Exit Sub
            
            txtRawD = Str$(DataValue)
            txtEngD = Str$(DataValue)
            
        End If
        
    End With

    ULStat = cbDConfigPort(BoardNum, PortNum, DigitalIn)
    If ULStat <> 0 Then
            txtEngD = "ERR: " & Trim(Str(ULStat))
            Exit Sub
    End If

    txtRawD = vbNullString
    txtEngD = vbNullString
    
    ' read digital input and display
   
    ULStat = cbDIn(BoardNum, PortNum, DataValue%)
    If ULStat <> 0 Then
        txtEngD = "ERR: " & Trim(Str(ULStat))
        Exit Sub
    End If

    txtRawD = Str$(DataValue)
    txtEngD = Str$(DataValue)
    
End Sub

Public Sub DigitalOutput(ByVal BoardNum As Long, ByVal PortNum As Long, ByVal DataValue As Long)
'This function only works for the two MCC boards - DAS-PCI-6030 and 1616HS-2 boards.

    Dim ULStat As Long
   
    If NOCOMM_MODE Then Exit Sub
   
    txtRawD = txtEngD
   
    cmbBoardD.ListIndex = BoardNum
    cmbBoardD_Click
        
    For i = 0 To cmbChanDOut.ListCount - 1
    
        If cmbChanDOut.ItemData(i) = PortNum Then
        
            cmbChanDOut.ListIndex = i
            
        End If
        
    Next i
    
    With DAQBoards.Item(BoardNum + 1)
    
        If .BoardName = "PCI-DAS6030" Then
        
            'Need to use AUXPORT and portnum = bitnum
            ULStat = cbDConfigBit(BoardNum, AUXPORT, PortNum, DigitalOut)
            
            If ULStat <> 0 Then
                txtEngD = "ERR: " & Trim(Str(ULStat))
                Exit Sub
            End If
            
            'Now use bit out command to write to port
            ULStat = cbDBitOut(BoardNum, AUXPORT, PortNum, DataValue)
            
            If ULStat <> 0 Then
                txtEngD = "ERR: " & Trim(Str(ULStat))
                Exit Sub
            End If
            
            txtRawD = Str$(DataValue)
            txtEngD = Str$(DataValue)
            
            Exit Sub
            
        End If
        
    End With
    
    ULStat = cbDConfigPort(BoardNum, PortNum, DigitalOut)
    If ULStat <> 0 Then
        txtEngD = "ERR: " & Trim(Str(ULStat))
        Exit Sub
    End If
   
    ' write the value
  
    ULStat = cbDOut(BoardNum, PortNum, DataValue)
   
    If ULStat <> 0 Then
        txtEngD = "ERR: " & Trim(Str(ULStat))
    End If
    txtRawD = Str$(DataValue)
    txtEngD = Str$(DataValue)

End Sub
Private Function ULongValToInt(LongVal As Long) As Integer

   Select Case LongVal
      Case Is > 65535
         ULongValToInt = -1
      Case Is < 0
         ULongValToInt = 0
      Case Else
         ULongValToInt = (LongVal - 32768) Xor &H8000
   End Select

End Function

Private Sub Form_Unload(Cancel As Integer)

    Dim i As Long
    Dim j As Long

    'Stop the analog signal generator background process
    cmdStopSineWave_Click
    
    'Return Timer to normal ~15 ms accuracy
    timeEndPeriod 1
    
    'Turn all the relay switches closed so that their control relays switch off
    'and don't over heat
    For i = 1 To DAQBoards.count
        
        If InStr(1, DAQBoards(i).BoardFunction, AFRELAYCONTROL) > 0 Then
        
            With DAQBoards(i)
            
                cmbBoardD.ListIndex = .BoardNum
            
                'Close all of the relay switches so that they power off
                'and don't overheat
                DigitalOutput .BoardNum, FIRSTPORTA, 1
                DigitalOutput .BoardNum, FIRSTPORTB, 1
                DigitalOutput .BoardNum, FIRSTPORTC, 1
                
            End With
            
        End If
        
'        With DAQBoards(i)
'
'            'If board has analog output channels,
'            'Zero all the analog output channels for that board
'            If .AOutChannels.count > 0 Then
'
'                For j = 1 To .AOutChannels.count
'
'                    AnalogOutput .BoardNum, _
'                                    .AOutChannels.Item(j).ChanNum, _
'                                    0, _
'                                    .Range
'
'                Next j
'
'            End If
'
'        End With
            
    Next i

End Sub

Private Sub optCoil_Click(Index As Integer)
    
    Dim TTLBoard As Board
    Dim DataValue As Integer
    Dim ULStats As Long
    Dim i, N As Long
    Dim MemBuffer As Long
    Dim DataArray() As Double
    Dim TempArray() As Integer
    Dim Status As Integer
    Dim CurCount, CurIndex As Long
        
    Dim CRange As Range
    Dim TempD As Double
    Dim TempI As Integer
    Dim Temp2_I As Integer
    
    Set TTLBoard = Nothing
        
'------DEBUG-------------------------------------------------------------
'    'Start an analog input scan on the green donut feedback line
'
'    MemBuffer = cbWinBufAlloc(100000)
'
'    ULstats = cbAInScan(0, 1, 2, 100000, 30000, BIP10VOLTS, MemBuffer, BACKGROUND)
'
'    If ULstats <> 0 Then
'
'        Err.Raise ULstats, "cbAinScan"
'
'    End If
    
    'Search through boards until find the first AFRELAYCONTROL board
    For i = 1 To DAQBoards.count
    
        'If the board number matches the digital out
        If InStr(1, DAQBoards(i).BoardFunction, AFRELAYCONTROL) > 0 Then
        
            Set TTLBoard = DAQBoards(i)
            
            i = DAQBoards.count + 1
            
        End If
        
    Next i

    'Make sure only one coil is selected at a time
    If Index = 0 Then optCoil(1) = Not optCoil(0)
    If Index = 1 Then optCoil(0) = Not optCoil(1)
    
    DataValue = 1
    
    'If Axial coil is selected
    If optCoil(0) Then
    
        'Close Transverse Coil port, set FIRSTPORTB to zero
        DigitalOutput TTLBoard.BoardNum, FIRSTPORTB, 0
    
        'Open Axial Coil port, Set FIRSTPORTC to 1 volts
        DigitalOutput TTLBoard.BoardNum, FIRSTPORTC, 1
        
    End If
    
    'If Transverse coil is selected
    If optCoil(1) Then
    
        'Close Axial Coil port, FIRSTPORTC, to zero
        DigitalOutput TTLBoard.BoardNum, FIRSTPORTC, 0
                
        'Open Transverse Coil port, FIRSTPORTB, to one
        DigitalOutput TTLBoard.BoardNum, FIRSTPORTB, 1
        
    End If
    
    
'    ULstats = cbGetStatus(0, Status, CurCount, CurIndex, AIFUNCTION)
'
'    If ULstats <> 0 Then
'
'        Err.Raise ULstats, "cbGetStatus"
'
'    End If
'
'    ReDim DataArray(CurIndex)
'    ReDim TempArray(CurIndex)
'
'    ULstats = cbWinBufToArray(MemBuffer, TempArray(0), 0, CurIndex)
'
'    If ULstats <> 0 Then
'
'        Err.Raise ULstats, "cbWinBufToArray"
'
'    End If
'
'    ULstats = cbStopBackground(0, AIFUNCTION)
'
'    If ULstats <> 0 Then
'
'        Err.Raise ULstats, "cbStopBackground"
'
'        Exit Sub
'
'    End If
'
'    Set CRange = New Range
'    CRange.RangeType = BIP10VOLTS
'
'    For i = 0 To UBound(TempArray) - 1
'
'        TempI = TempArray(i)
'
'        CRange.MCC_RangeConverter TempD, _
'                                  TempI, _
'                                  Temp2_I, _
'                                  MCC_CountsToVolts
'
'        DataArray(i) = TempD
'
'    Next i
'
'    Dim fso As FileSystemObject
'    Dim DataStream As TextStream
'
'    Set fso = New FileSystemObject
'    Set DataStream = fso.CreateTextFile("C:\Documents and Settings\lab\Desktop\Test MCC Board 11-16-2009\CoilSwitch_" & Format(Now, "MM-DD-YY_HH-MM-SS") & ".csv")
'
'    If Index = 0 Then
'
'        DataStream.WriteLine "Switch to Axial Coil"
'
'    Else
'
'        DataStream.WriteLine "Switch to Transverse Coil"
'
'    End If
'
'    DataStream.WriteLine Format(Now, "long date") & "," & Format(Now, "long time")
'    DataStream.WriteBlankLines (1)
'    DataStream.WriteLine "Pt #,Ch1,Ch2"
'
'    N = UBound(DataArray)
'
'    For i = 0 To N - 1 Step 2
'
'        DataStream.WriteLine Trim(Str(i \ 2)) & "," & Trim(Str(DataArray(i))) & _
'                             "," & Trim(Str(DataArray(i + 1))) & "," & _
'                             Trim(Str(TempArray(i))) & "," & Trim(Str(TempArray(i + 1)))
'
'    Next i
'
'    DataStream.Close
'
'    ULstats = cbWinBufFree(MemBuffer)
    
End Sub

Private Sub txtFreq_Change()

    txtSineIORate_Change
    
End Sub

Private Sub txtSineIORate_Change()

    If val(txtFreq) <> 0 Then
        Me.txtPtsPerPeriod = Trim(Str((val(txtSineIORate) / val(txtFreq))))
        Me.txtRampPeakDuration = Trim(Str(1 / val(txtFreq) * 100000))
    End If

End Sub


