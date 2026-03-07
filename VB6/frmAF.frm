VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmAF 
   Caption         =   "AF Degausser"
   ClientHeight    =   4755
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7470
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4755
   ScaleWidth      =   7470
   Begin ComctlLib.ProgressBar progressBarAFFileSave 
      Height          =   255
      Left            =   4320
      TabIndex        =   34
      Top             =   3720
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.CheckBox chkVerbose 
      Caption         =   "Verbose?"
      Height          =   375
      Left            =   4680
      TabIndex        =   33
      Top             =   2880
      Width           =   1335
   End
   Begin MSCommLib.MSComm MSCommAF 
      Left            =   4680
      Top             =   720
      _ExtentX        =   794
      _ExtentY        =   794
      _Version        =   393216
      DTREnable       =   0   'False
      OutBufferSize   =   8
      BaudRate        =   1200
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Connect"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   0
      Width           =   1335
   End
   Begin VB.CommandButton cmdSendStatus 
      Caption         =   "Send Status"
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   0
      Width           =   1335
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   252
      Left            =   3960
      TabIndex        =   2
      Top             =   0
      Width           =   972
   End
   Begin VB.TextBox txtSampHeight 
      Height          =   285
      Left            =   6960
      TabIndex        =   3
      Text            =   "0"
      Top             =   0
      Width           =   495
   End
   Begin VB.TextBox OutputText 
      Height          =   285
      Left            =   720
      TabIndex        =   4
      Top             =   600
      Width           =   852
   End
   Begin VB.TextBox InputText 
      Height          =   288
      Left            =   2400
      TabIndex        =   5
      Top             =   600
      Width           =   1932
   End
   Begin VB.CheckBox chkLocked 
      Caption         =   "Locked"
      Height          =   252
      Left            =   2400
      TabIndex        =   6
      Top             =   960
      Width           =   1212
   End
   Begin VB.CommandButton cmdManualAxialAF 
      Caption         =   "Manual Axial AF"
      Height          =   375
      Left            =   5520
      TabIndex        =   7
      Top             =   360
      Width           =   1815
   End
   Begin VB.CommandButton cmdManualTransverseAF 
      Caption         =   "Manual Transverse AF"
      Height          =   375
      Left            =   5520
      TabIndex        =   8
      Top             =   840
      Width           =   1815
   End
   Begin VB.Frame frameActiveCoil 
      Caption         =   "Active Coil System"
      Height          =   852
      Left            =   240
      TabIndex        =   9
      Top             =   1320
      Width           =   1812
      Begin VB.OptionButton optActiveAxial 
         Caption         =   "Axial"
         Height          =   252
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1572
      End
      Begin VB.OptionButton optActiveTransverse 
         Caption         =   "Transverse"
         Height          =   252
         Left            =   120
         TabIndex        =   11
         Top             =   480
         Width           =   1572
      End
   End
   Begin VB.TextBox txtUncalAmplitude 
      Height          =   288
      Left            =   4440
      TabIndex        =   12
      Top             =   1560
      Width           =   852
   End
   Begin VB.CommandButton cmdSetUncalAmp 
      Caption         =   "Set Amplitude w/o Calibration"
      Height          =   492
      Left            =   5520
      TabIndex        =   13
      Top             =   1440
      Width           =   1692
   End
   Begin VB.CommandButton cmdCleanCoils 
      Caption         =   "Clean Coils"
      Height          =   372
      Left            =   4680
      TabIndex        =   14
      Top             =   2400
      Width           =   1452
   End
   Begin VB.TextBox txtAmplitude 
      Height          =   288
      Left            =   1200
      TabIndex        =   15
      Top             =   2760
      Width           =   852
   End
   Begin VB.CommandButton cmdConfigAmplitude 
      Caption         =   "Set Amplitude"
      Height          =   252
      Left            =   2160
      TabIndex        =   16
      Top             =   2760
      Width           =   1692
   End
   Begin VB.ComboBox cmbDelay 
      Height          =   315
      Left            =   1320
      TabIndex        =   17
      Text            =   "Combo1"
      Top             =   3240
      Width           =   732
   End
   Begin VB.CommandButton cmdConfigDelay 
      Caption         =   "Set Delay"
      Height          =   252
      Left            =   2160
      TabIndex        =   18
      Top             =   3240
      Width           =   1692
   End
   Begin VB.ComboBox cmbRampRate 
      Height          =   315
      Left            =   1320
      TabIndex        =   19
      Top             =   3720
      Width           =   732
   End
   Begin VB.CommandButton cmdConfigureRampRate 
      Caption         =   "Set Ramp"
      Height          =   252
      Left            =   2160
      TabIndex        =   20
      Top             =   3720
      Width           =   1692
   End
   Begin VB.ComboBox cmbRampMode 
      Height          =   315
      Left            =   1440
      TabIndex        =   21
      Text            =   "Combo1"
      Top             =   4200
      Width           =   612
   End
   Begin VB.CommandButton cmdExecuteRamp 
      Caption         =   "Execute Ramp"
      Height          =   252
      Left            =   2160
      TabIndex        =   22
      Top             =   4200
      Width           =   1692
   End
   Begin VB.TextBox txtWaitingTime 
      Height          =   285
      Left            =   4680
      TabIndex        =   23
      Text            =   "0"
      Top             =   4200
      Width           =   375
   End
   Begin VB.Label lblFileSaveProg 
      Caption         =   "File Save % Complete:"
      Height          =   255
      Left            =   4320
      TabIndex        =   35
      Top             =   3360
      Width           =   2535
   End
   Begin VB.Label Label9 
      Caption         =   "Sample Height (cm):"
      Height          =   255
      Left            =   5400
      TabIndex        =   24
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Output:"
      Height          =   255
      Left            =   120
      TabIndex        =   25
      Top             =   600
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Input:"
      Height          =   255
      Left            =   1920
      TabIndex        =   26
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label7 
      Caption         =   "Uncalibrated Amplitude:"
      Height          =   255
      Left            =   2640
      TabIndex        =   27
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Amplitude:"
      Height          =   255
      Left            =   240
      TabIndex        =   28
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Delay:"
      Height          =   255
      Left            =   240
      TabIndex        =   29
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Ramp:"
      Height          =   255
      Left            =   240
      TabIndex        =   30
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "Ramp Mode:"
      Height          =   255
      Left            =   240
      TabIndex        =   31
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Label Label8 
      Caption         =   "Waiting time (in s) between ramps"
      Height          =   375
      Left            =   5160
      TabIndex        =   32
      Top             =   4200
      Width           =   1335
   End
End
Attribute VB_Name = "frmAF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Alternating Field Demagnetization Driver
' 2G AFnetization unit Driver
'
' This is the driver for the 2G Demagnetization driver
' It provides the most common functions
' Note that the Keithley routines to check the status of the Af signal lines
' and to switch between biomag and paleomag are not used any more.
Option Explicit ' enforce variable declaration!
Dim locked As Boolean
Dim ActiveCoilSystem As Integer
Dim currentDelay As Integer
Dim currentRampRate As Integer
Dim currentUncalAmp As Double
Dim currentCalAmp As Double
Dim currentAxis As String
Dim Feedback As String ' (August 2007 L Carporzen) Allow to record in a file the AF communications

'-------------------------------
'---DEBUG!!---------------------
'   (Mar 2010, I Hilburn)
'   Variables to allow analog input background process on the PCI-DAS6030 board to monitor
'   and save the donut feedback voltage from the AF LC circuit.
Dim NumPoints As Long
Dim IORate As Long
Dim IOOptions As Long
Dim BoardNum As Long
Dim ChanNum As Long
Dim ULStats As Long
Dim Range As Long
Dim MemBuffer As Long
Dim NoError As Boolean
Dim MonitorArray() As Single
Dim gainArray(1) As Long
Dim Status As Integer
Dim CurCount As Long
Dim CurIndex As Long
'-------------------------------
'-------------------------------

Public Function FindXCalibValue(field As Double, Optional CoilSystem As Integer = -128)
    ' Find X (input to AF) from field
    Dim i As Integer
    Dim slope As Double
    FindXCalibValue = -1
    If CoilSystem = -128 Then CoilSystem = ActiveCoilSystem
        If CoilSystem = AxialCoilSystem Then
        If field > AfAxialMax Then
            field = AfAxialMax
        ElseIf (field < AfAxialMin) And (field <> 0) Then
            field = AfAxialMin
        End If
        For i = 1 To 25
            If AFAxialY(i) = field Then
                FindXCalibValue = AFAxialX(i)
                Exit Function
            ElseIf AFAxialY(i - 1) < field And AFAxialY(i) > field Then
                slope = (AFAxialX(i) - AFAxialX(i - 1)) / (AFAxialY(i) - AFAxialY(i - 1))
                FindXCalibValue = AFAxialX(i - 1) + slope * (field - AFAxialY(i - 1))
                Exit For
            End If
        Next i
    ElseIf CoilSystem = TransverseCoilSystem Then
        If field > AfTransMax Then
            field = AfTransMax
        ElseIf (field < AfTransMin) And (field <> 0) Then
            field = AfTransMin
        End If
        For i = 1 To 25
            If AFTransY(i) = field Then
                FindXCalibValue = AFTransX(i)
                Exit Function
            ElseIf AFTransY(i - 1) < field And AFTransY(i) > field Then
                slope = (AFTransX(i) - AFTransX(i - 1)) / (AFTransY(i) - AFTransY(i - 1))
                FindXCalibValue = AFTransX(i - 1) + slope * (field - AFTransY(i - 1))
                Exit For
            End If
        Next i
    End If
End Function

Private Sub form_resize()
    Me.Height = 5445
    Me.Width = 8025
End Sub

Private Sub LockAF(locking As Boolean)
    If locking Then
        locked = True
        cmdConfigAmplitude.Enabled = False
        cmdConfigDelay.Enabled = False
        cmdConnect.Enabled = False
        cmdSendStatus.Enabled = False
        cmdExecuteRamp.Enabled = False
        cmdSetUncalAmp.Enabled = False
        cmdConfigureRampRate.Enabled = False
        chkLocked.value = vbChecked
    Else
        locked = False
        cmdConfigAmplitude.Enabled = True
        cmdConfigDelay.Enabled = True
        cmdConnect.Enabled = True
        cmdSendStatus.Enabled = True
        cmdExecuteRamp.Enabled = True
        cmdSetUncalAmp.Enabled = True
        cmdConfigureRampRate.Enabled = True
        chkLocked.value = vbUnchecked
        SetForm
    End If
End Sub

Private Sub SetForm()
    txtAmplitude = currentCalAmp
    txtUncalAmplitude = currentUncalAmp
    cmbDelay = currentDelay
    cmbRampRate = currentRampRate
    If currentAxis = AfAxialCoord Then
        optActiveAxial.value = True
        optActiveTransverse.value = False
    ElseIf currentAxis = AfTransCoord Then
        optActiveAxial.value = False
        optActiveTransverse.value = True
    Else
        optActiveAxial.value = False
        optActiveTransverse.value = False
    End If
End Sub

Private Sub SetActiveCoilSystem(newActiveCoilSystem As Integer)
    If locked Then
        SetForm
        Exit Sub
    End If
    If newActiveCoilSystem = AxialCoilSystem Then
        ActiveCoilSystem = newActiveCoilSystem
        optActiveAxial.value = True
        ConfigureCoil AfAxialCoord
    End If
    If newActiveCoilSystem = TransverseCoilSystem Then
        ActiveCoilSystem = newActiveCoilSystem
        optActiveTransverse.value = True
        ConfigureCoil AfTransCoord
    End If
End Sub

Public Sub CleanCoils()
    ExecuteRamp "C", AxialCoilSystem, AfAxialMax, AFDelay, AFRampRate
    ExecuteRamp "C", TransverseCoilSystem, AfTransMax, AFDelay, AFRampRate
End Sub

Public Sub ManualAxialAF()
    If frmVacuum.VacuumConnectOn = True Then ' (February 2008 L Carporzen) Manual Axial demag
        If val(txtAmplitude) <= 0 Then txtAmplitude = val(InputBox("What is amplitude (in Oe) of the axial demagnetization you want?", "Important!", txtAmplitude))
        If val(txtAmplitude) > AfAxialMax Then txtAmplitude = AfAxialMax
        If val(txtAmplitude) <= 0 Then Exit Sub
        txtSampHeight = val(InputBox("Doing a " & Int(val(txtAmplitude)) & " Oe axial demagnetization" & vbCr & "What is the height (in cm) of the sample?", "Important!", txtSampHeight))
        frmDCMotors.UpDownMove (AFPos + txtSampHeight * UpDownMotor1cm / 2), 1
        ExecuteRamp "C", AxialCoilSystem, val(txtAmplitude), frmSettings.cmbAFDelay, frmSettings.cmbAFRampRate
        frmDCMotors.HomeToTop
    Else
        MsgBox "Aborted! Place a sample first..."
    End If
End Sub

Public Sub ManualTransverseAF()
    If frmVacuum.VacuumConnectOn = True Then ' (February 2008 L Carporzen) Manual Transverse demag
        If val(txtAmplitude) <= 0 Then txtAmplitude = val(InputBox("What is amplitude (in Oe) of the transverse demagnetization you want?", "Important!", txtAmplitude))
        If val(txtAmplitude) > AfTransMax Then txtAmplitude = AfTransMax
        If val(txtAmplitude) <= 0 Then Exit Sub
        txtSampHeight = val(InputBox("Doing a " & Int(val(txtAmplitude)) & " Oe transverse demagnetization" & vbCr & "What is the height (in cm) of the sample?", "Important!", txtSampHeight))
        frmDCMotors.UpDownMove (AFPos + txtSampHeight * UpDownMotor1cm / 2), 1
        ExecuteRamp "C", TransverseCoilSystem, val(txtAmplitude), frmSettings.cmbAFDelay, frmSettings.cmbAFRampRate
        frmDCMotors.HomeToTop
    Else
        MsgBox "Aborted! Place a sample first..."
    End If
End Sub

Private Sub ConfigureAmplitude(ByVal Amplitude As Double, Optional ByVal CoilSystem As Integer = -128)
    ' Calibrate and set amplitude.
    Dim AFLevel As Double
    If locked Then
        SetForm
        Exit Sub
    End If
    If CoilSystem = -128 Then CoilSystem = ActiveCoilSystem
        currentCalAmp = Amplitude
    SetAmplitude (FindXCalibValue(Amplitude, CoilSystem))
End Sub

Private Sub SetAmplitude(ByVal AFLevel As Double)
    If locked Then
        SetForm
        Exit Sub
    End If
    If AFLevel < 0 Then AFLevel = 0
    'If AFLevel > 3000 Then AFLevel = 3000
    'If currentUncalAmp = AFLevel Then Exit Sub
    txtUncalAmplitude = AFLevel
    LockAF True
    SendCommand ("DCA" + Format$(AFLevel, "0000"))
    Feedback = GetResponse
    'If InStr(Feedback, Format$(AFLevel, "0000")) Then ' (August 2007 L Carporzen) Allow to record in a file the AF communications
    currentUncalAmp = AFLevel
    'WriteAF "DCA" + Format$(AFLevel, "0000"), "AFLevel"
    'Else
    'WriteAF "DCA" + Format$(AFLevel, "0000"), "AFLevel"
    'WriteAF Feedback, "Wrong answer"
       ' MsgBox ("Amplitude " & AFLevel & " not received, the feedback is " & Feedback)
    'End If
    LockAF False
End Sub

Public Sub ConfigureCoil(axis As String)
    If locked Then
        SetForm
        Exit Sub
    End If
    If currentAxis = axis Then Exit Sub
    LockAF True
    SendCommand "DCC" + axis
    Feedback = GetResponse
    'If InStr(Feedback, axis) Then ' (August 2007 L Carporzen) Allow to record in a file the AF communications
    currentAxis = axis
    'WriteAF "DCC" + axis, "Axis"
    'Else
    'WriteAF "DCC" + axis, "Axis"
    'WriteAF Feedback, "Wrong answer"
       ' MsgBox ("Axis " & axis & " not received, the feedback is " & Feedback)
    'End If
    LockAF False
End Sub

Private Sub ConfigureDelay(ByVal Delay As Integer)
    If locked Then
        SetForm
        Exit Sub
    End If
    If Delay < 1 Then
        Delay = 1
    ElseIf Delay > 9 Then
        Delay = 9
    End If
    'If currentDelay = Delay Then Exit Sub
    LockAF True
    cmbDelay = Left$(Format$(Delay), 1)
    SendCommand "DCD" + Left$(Format$(Delay), 1)
    Feedback = GetResponse
    'If InStr(Feedback, Left$(Format$(Delay), 1)) Then ' (August 2007 L Carporzen) Allow to record in a file the AF communications
    currentDelay = Delay
    'WriteAF "DCD" + Left$(Format$(Delay), 1), "Delay"
    'Else
    'WriteAF "DCD" + Left$(Format$(Delay), 1), "Delay"
    'WriteAF Feedback, "Wrong answer"
       ' MsgBox ("Delay " & Delay & " not received, the feedback is " & Feedback)
    'End If
    LockAF False
End Sub

Private Sub ConfigureRampRate(ByVal ramp As Integer)
    If locked Then Exit Sub
    If Not (ramp = 3 Or ramp = 5 Or ramp = 7 Or ramp = 9) Then Exit Sub
    If currentRampRate = ramp Then Exit Sub
    LockAF True
    cmbRampRate = Left$(Format$(ramp), 1)
    SendCommand "DCR" + Left$(Format$(ramp), 1)
    Feedback = GetResponse
    'If InStr(Feedback, Left$(Format$(ramp), 1)) Then ' (August 2007 L Carporzen) Allow to record in a file the AF communications
    currentRampRate = ramp
    'WriteAF "DCR" + Left$(Format$(ramp), 1), "Ramp"
    'Else
    'WriteAF "DCR" + Left$(Format$(ramp), 1), "Ramp"
    'WriteAF Feedback, "Wrong answer"
       ' MsgBox ("Ramp rate " & ramp & " not received, the feedback is " & Feedback)
    'End If
    LockAF False
End Sub

Public Sub CycleWithHold(Optional ByVal HoldTime As Integer = 0, Optional CoilSystem As Integer = -128, _
    Optional ByVal Amplitude As Double = -1, _
    Optional RampRate As Integer = -1)
    Dim olddelay As Integer
    olddelay = currentDelay
    If HoldTime = 0 Then HoldTime = AFDelay
    ExecuteRamp "C", CoilSystem, Amplitude, HoldTime, RampRate
    If Not HoldTime = AFDelay Then ConfigureDelay olddelay
End Sub

Public Sub ExecuteRamp(ByVal Mode As String, Optional CoilSystem As Integer = -128, _
    Optional ByVal Amplitude As Double = -1, Optional Delay As Integer = -1, _
    Optional RampRate As Integer = -1)
    
    Dim reply As String
    Dim FolderName As String
    Dim WindowForFindingMax As Long
    Dim CurTime
    
    
    If Not MSCommAF.PortOpen And Not NOCOMM_MODE Then Connect
    If locked Then
        MsgBox "AF unit is use.  Ramp execution is not possible."
        Exit Sub
    End If
    If Not NOCOMM_MODE Then
        frmProgram.StatusBar "AF config", 2
        If (CoilSystem = AxialCoilSystem) Or (CoilSystem = TransverseCoilSystem) Then _
            SetActiveCoilSystem CoilSystem
        If Amplitude >= 0 Then ConfigureAmplitude Amplitude
        If Delay > 0 Then ConfigureDelay Delay
        If RampRate > 0 Then ConfigureRampRate RampRate
        If (CoilSystem = AxialCoilSystem) Or (CoilSystem = TransverseCoilSystem) Then _
            SetActiveCoilSystem CoilSystem
                If Mode <> "U" And Mode <> "D" And Mode <> "C" Then Exit Sub
        If Mode = "U" Then MsgBox "Ramping up without ramping down is dangerous..."
        LockAF True
        frmProgram.StatusBar "AF execute", 2
        cmbRampMode = Mode
        
'-------Debug Code-------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------
'   (Mar, 2010 - I Hilburn)
'   Recording AF Ramp using the 2G box through an Analog input port on the PCI-DAS6030 card to
'   compare to the AF Ramp using the MCC hardware and new AF code
'   Input Rate = 100 kHz
'   The values here, including the channel, ramp rate, IO Options, boardnum, etc., are hard
'   wired in this code!
        
        If Me.chkVerbose.value = Checked Then
        
            'User has selected a Verbose ramp, which means we need to allocate an MCC Windows Memory buffer
            
            'Set the Ramp Analog input process parameters
            'Set IORate = 100 kHz
            IORate = 100000
            
            'Set BoardNum = 0, which should be the Dev board # for the PCI-DAS6030 board on this machine
            BoardNum = 0
            
            'Set Channel Num = 1, Analog Input Channel #2
            ChanNum = 1
            
            'Set Range = + or - 10 Volts
            Range = BIP10VOLTS
            
            'IO Options = allow analog input process to run in the background
            IOOptions = BACKGROUND
        
            'NumPoints = IORate * 30 seconds we expect the ramp will run.
            NumPoints = IORate * 30
            
            'Set flag to indicate if no error has happened to True
            NoError = True
            
            'Allocate the windows memory buffer for the board to dump points into
            MemBuffer = cbWinBufAlloc(NumPoints)
            
            'Error Check
            If MemBuffer = 0 Then
            
                MsgBox "Memory Buffer allocation failed." & vbNewLine & _
                        "Buffer # = " & Trim(Str(MemBuffer)), , _
                        "Ramp Monitor Error"
                
                'Set the No Error flag to false
                NoError = False
                        
            End If
            
            'If no error has happened, start the analog input scan
            If NoError Then
            
                ULStats = cbAInScan(BoardNum, _
                                        ChanNum, _
                                        ChanNum, _
                                        NumPoints, _
                                        IORate, _
                                        Range, _
                                        MemBuffer, _
                                        IOOptions)
                                        
                Debug.Print "# pts = " & Trim(Str(NumPoints))
                Debug.Print "IORate = " & Trim(Str(IORate))
                                        
                'Error Check
                If ULStats <> 0 Then
                
                    MsgBox "Could not start Analog Input scan before 2G Ramp." & vbNewLine & vbNewLine & _
                            "Board # = " & Trim(Str(BoardNum)) & vbNewLine & _
                            "Channel # = " & Trim(Str(ChanNum)) & vbNewLine & _
                            "Err = " & Trim(Str(ULStats)), , _
                            "Ramp Monitor Error"
                            
                    NoError = False
                    
                End If
                
            End If
                        
        End If
        
'----------------------------------------------------------------------------------
'----------------------------------------------------------------------------------
        
        SendCommand "DER" + Mode
        'Feedback = GetResponse ' (August 2007 L Carporzen) Allow to record in a file the AF communications
        'WriteAF "DER" + Mode, "Mode"
        'WriteAF Feedback, "Answer"
        PollAFUnit
        frmProgram.StatusBar vbNullString, 2
        
        
'-------Debug Code-------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------
'   (Mar, 2010 - I Hilburn)
'   Recording AF Ramp using the 2G box through an Analog input port on the PCI-DAS6030 card to
'   compare to the AF Ramp using the MCC hardware and new AF code.
'
'   AF Ramp has finished, need to end the background analog input process on the Measurement computing
'   board if no error in setting up the input scan occurred
        
        If chkVerbose.value = Checked Then
        
            'Now, if no error has occurred, need to figure out how many points of data were loaded into the
            'Memory buffer
            If NoError Then
            
                ULStats = cbGetStatus(BoardNum, _
                                        Status, _
                                        CurCount, _
                                        CurIndex, _
                                        AIFUNCTION)
                                        
                Debug.Print "Cur Index = " & Trim(Str(CurIndex))
                                        
                'Error Check
                If ULStats <> 0 Then
                
                    MsgBox "Could not get status of Analog Input Scan after 2G Ramp." & vbNewLine & vbNewLine & _
                            "Board # = " & Trim(Str(BoardNum)) & vbNewLine & _
                            "Channel # = " & Trim(Str(ChanNum)) & vbNewLine & _
                            "Memory Buffer # = " & Trim(Str(MemBuffer)) & vbNewLine & _
                            "Err = " & Trim(Str(ULStats)), , _
                            "Ramp Monitor Error"
                            
                    NoError = False
                    
                End If
                
            End If
        
            If NoError Then
            
                'Now need to end the background process on the analog input channel
                ULStats = cbStopBackground(BoardNum, AIFUNCTION)
                
                'Error Check
                If ULStats <> 0 Then
                
                    MsgBox "Could not end background analog input process after 2G Ramp." & vbNewLine & vbNewLine & _
                            "Board # = " & Trim(Str(BoardNum)) & vbNewLine & _
                            "Channel # = " & Trim(Str(ChanNum)) & vbNewLine & _
                            "Err = " & Trim(Str(ULStats)), , _
                            "Ramp Monitor Error"
                            
                    NoError = False
                    
                End If
                
            End If
            
            'If no error, then need to redimension the MonitorArray by the number of points (CurIndex) in the monitor memory
            'buffer, then need to pull all the data points from the buffer into the array.  This could take a while.
            If NoError Then
            
                On Error Resume Next
                
                ReDim MonitorArray(CurIndex)
                
                If Err.number <> 0 Then
                
                    MsgBox "Could not redimension the Monitor Array." & vbNewLine & _
                            "Num Points in buffer = " & Trim(Str(CurIndex)), , _
                            "Ramp Monitor Error"
                            
                    NoError = False
                    
                End If
                
                On Error GoTo 0
                
            End If
            
            If NoError Then
            
                'Store the range in the gainArray
                gainArray(0) = Range
            
                'Pull the contents of the windows memory buffer into the Monitor Array
                ULStats = cbWinBufToEngUnits(BoardNum, _
                                                gainArray(0), _
                                                1, _
                                                MemBuffer, _
                                                MonitorArray(0), _
                                                0, _
                                                CurIndex)
                                                
                'Error Check
                If ULStats <> 0 Then
                
                    MsgBox "Could not load 2g Ramp data points in memory buffer to Data Array." & vbNewLine & _
                            "Memory Buffer # = " & Trim(Str(MemBuffer)) & vbNewLine & _
                            "# of Points = " & Trim(Str(CurIndex)) & vbNewLine & _
                            "Array Size = " & Trim(Str(UBound(MonitorArray))) & vbNewLine & vbNewLine & _
                            "Err = " & Trim(Str(ULStats)), , _
                            "Ramp Monitor Error"
                            
                    NoError = False
                    
                End If
                
            End If
            
            
            'If no errors have happened, then need to write the contents of the MonitorArray to file
            If NoError Then
                
                CurTime = Now
                
                'Set New AF Ramp Data Folder Name = combo of Peak Field and current Date-Time
                FolderName = "AF Ramp - " & Me.txtAmplitude & " G - " & Format(CurTime, "MM-DD-YY_HH-MM-SS") & "/"
                
                'Set the Window of points over which to search for a max value
                If optActiveAxial.value = True Then
                
                    WindowForFindingMax = 150
                    
                Else
                
                    WindowForFindingMax = 400
                    
                End If
                
                'Note - while this function may pause the code, it should never break the code,
                'So it is safe to wait to deallocate the windows memory buffer until
                'after this function has run
                MultiRampFileSave MonitorArray, _
                                    65000, _
                                    RampDataLocalPath, _
                                    FolderName, _
                                    CurTime, _
                                    RampDataBackupPath, _
                                    True, _
                                    WindowForFindingMax
                                    
                
                'Now Deallocate the Windows memory buffer
                ULStats = cbWinBufFree(MemBuffer)
                
                'Error Check
                If ULStats <> 0 Then
                
                    MsgBox "Could not deallocate windows memory buffer." & vbNewLine & _
                            "Memory Buffer # = " & Trim(Str(MemBuffer)) & vbNewLine & _
                            "Err = " & Trim(Str(ULStats)), , _
                            "Ramp Monitor Error"
                            
                End If

                                    
            End If
            
        End If
        
    End If
    LockAF False
    
End Sub

'Sub AF_Grab()
'    Keithley_SetPosition ChannelKeithDig, AFGrabPinPosition
'End Sub


'Sub AF_Release()
'    Keithley_ClearPosition ChannelKeithDig, AFGrabPinPosition
'End Sub

Private Function SendStatus() As String
    Dim alreadyLocked As Boolean
    alreadyLocked = False
    alreadyLocked = locked
    If Not alreadyLocked Then LockAF True
    SendCommand "DSS"
    SendStatus = GetResponse
    If Not alreadyLocked Then LockAF False
End Function

Private Sub chkLocked_Click()
    If locked And chkLocked.value = vbUnchecked Then
        LockAF False
    ElseIf Not locked And chkLocked.value = vbChecked Then
        LockAF True
    End If
End Sub

Private Sub cmdCleanCoils_Click()
    CleanCoils
End Sub

Private Sub cmdManualAxialAF_Click()
    ManualAxialAF
End Sub

Private Sub cmdManualTransverseAF_Click()
    ManualTransverseAF
End Sub

Private Sub cmdClose_Click()
    Me.Hide
End Sub

Private Sub cmdConfigAmplitude_Click()
    ConfigureAmplitude val(txtAmplitude.Text)
End Sub

Private Sub cmdConfigDelay_Click()
    ConfigureDelay Int(val(cmbDelay))
End Sub

Private Sub cmdConfigureRampRate_Click()
    ConfigureRampRate Int(val(cmbRampRate))
End Sub

Private Sub ConnectButton_Click()
    If MSCommAF.PortOpen Then
        Disconnect
    Else
        Connect
    End If
End Sub

Private Sub cmdConnect_Click()
    If MSCommAF.PortOpen Then
        Disconnect
    Else
        Connect
    End If
End Sub

Private Sub cmdExecuteRamp_Click()
    ExecuteRamp cmbRampMode
End Sub

Private Sub cmdSendStatus_Click()
    SendStatus
End Sub

Private Sub cmdSetUncalAmp_Click()
    SetAmplitude val(txtUncalAmplitude)
End Sub

Private Sub Form_Load()
    ActiveCoilSystem = AxialCoilSystem
    cmbRampMode.Clear
    cmbRampMode.AddItem "U"
    cmbRampMode.AddItem "D"
    cmbRampMode.AddItem "C"
    cmbRampRate.Clear
    cmbRampRate.AddItem "3"
    cmbRampRate.AddItem "5"
    cmbRampRate.AddItem "7"
    cmbRampRate.AddItem "9"
    cmbDelay.Clear
    cmbDelay.AddItem "1"
    cmbDelay.AddItem "2"
    cmbDelay.AddItem "3"
    cmbDelay.AddItem "4"
    cmbDelay.AddItem "5"
    cmbDelay.AddItem "6"
    cmbDelay.AddItem "7"
    cmbDelay.AddItem "8"
    cmbDelay.AddItem "9"
    cmbDelay.Text = "1"
    currentDelay = -1
    currentRampRate = -1
    currentUncalAmp = -1
    
    Me.progressBarAFFileSave.Visible = False
    Me.lblFileSaveProg.Visible = False
    Me.lblFileSaveProg.Caption = "File Save % Complete:"
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If MSCommAF.PortOpen = True Then
        MSCommAF.PortOpen = False
    End If
End Sub

Private Sub SendCommand(outstring As String)
    Dim i As Integer
    Dim outchar As String
    If Not EnableAF Then Exit Sub
    frmProgram.StatusBar outchar, 3
    If MSCommAF.PortOpen = True Then
        MSCommAF.RTSEnable = True
        MSCommAF.OutBufferCount = 0
        MSCommAF.InBufferCount = 0
        ' Because AF unit it stupid, we send out one character
        ' every 15 ms.
        MSCommAF.Output = Chr$(13)
        OutputText = vbNullString
        DelayTime 0.15
        For i = 1 To Len(outstring)
            outchar = Mid$(outstring, i, 1)
            MSCommAF.Output = outchar
            OutputText = OutputText & outchar
            DelayTime 0.15
        Next i
        MSCommAF.Output = Chr$(13)
        DelayTime 0.15
        'MSCommAF.InBufferCount = 0
        MSCommAF.OutBufferCount = 0
        If DEBUG_MODE Then frmDebug.Msg "COM " & Str$(MSCommAF.CommPort) & " out: " & outstring
    Else
        If Not NOCOMM_MODE Then MsgBox "AF Comm Port Not Open"
    End If
    frmProgram.StatusBar vbNullString, 3
End Sub

Private Sub PollAFUnit()
    Dim finished As Boolean
    Dim PollText As String
    Dim errormessage As String
    'Dim delay As Double
    Dim startTime As Double, totalsecs As Double   'new SWB
    'Dim startTime As Double, lag As Double
    Dim Status As String
    'startTime = Now
    'delay = Timer
    startTime = Timer   'Timer-starttime is seconds since start of polling
    totalsecs = 0   'use this to count total elapsed time for error msg
    finished = False
    Do While Not finished
        'If InStr(Feedback, "DON") > 0 Then ' (August 2007 L Carporzen) Execute ramp could have DONE in its Feedback
            'errormessage = "Strange DONE from AF box " & "on axis " & currentAxis & " at amplitude " & Format(currentCalAmp, 0) & ". " & PollText
            'frmSendMail.MailNotification "AF Alert", errormessage, CodeYellow
            'MsgBox errormessage
        '    finished = True
        'End If
        PollText = GetResponse
        'Normally, get DONE when AF cycle complete
        If InStr(PollText, "DON") > 0 Then
            finished = True
        'If a problem, should get ZERO ERROR or TRACK ERROR.  Look for ERROR in string
        ElseIf InStr(PollText, "ERROR") > 0 Then
            'YIKES - all engines stop
            errormessage = "The AF degaussing unit is experiencing an error on axis " & currentAxis & _
                " at amplitude " & Format(currentCalAmp, 0) & ":" & vbCrLf & vbCrLf & PollText & vbCrLf & _
                vbCrLf & "Execution has been paused and Ramp Down command sent. Please check machine."
            SendCommand "DERD"
            Flow_Pause
            SetCodeLevel CodeRed
            frmSendMail.MailNotification "AF Error", errormessage, CodeRed
            MsgBox errormessage
            SetCodeLevel StatusCodeColorLevelPrior, True
            'NOTE:  will exit loop by getting Z repsonse to DERD
        'Z sent back after DERD successful.  If ZERO ERROR, caught above
        ElseIf InStr(PollText, "Z") > 0 Then
            finished = True
        'T sent back after DERU successful.  If TRACK ERROR, caught above
        ElseIf InStr(PollText, "T") > 0 Then
            finished = True
        ' (August 2007 L Carporzen) We don't want to wait when the two lights are on, we need to switch them off quickly
        'ElseIf InStr(PollText, "") > 0 And InStr(PollText, "DERC") = 0 Then
        '    Status = SendStatus 'Need to verify that it is the error with a ? instead of Z
        '    If InStr(Status, "S ?") > 0 Then
        '        errormessage = "No DONE from AF box, we switch off the two lights after " & _
        '        Format$(totalsecs, "0.0") & " seconds; It appends " & "on axis " & currentAxis & _
        '        " at amplitude " & Format(currentCalAmp, 0) & " (2G = " & Format(currentUncalAmp, 0) & ")." & _
        '        " Target amplitude reported as zero, so unit appears to have reset. Execution will continue. " & Status
        '        SendCommand "DERD"
        '        frmSendMail.MailNotification "AF Alert", errormessage, CodeYellow
        '        LockAF False
        '        ActiveCoilSystem = 0
        '        currentDelay = -1
        '        currentRampRate = -1
        '        currentUncalAmp = -1
        '        currentCalAmp = -1
        '        currentAxis = vbNullString
        '        finished = True
        '    Else
        '    SetCodeLevel CodeGreen
        '    End If
        End If
        'Handle Timer rollover at midnight.  86400 secs/day
        If Timer < startTime Then startTime = startTime - 86400
        '====CODE BELOW CLIPPED OUT BY SWB======================================
        'If Timer < delay Then delay = delay - 86400
        'If Timer - delay > 9 Then
            'Status = SendStatus
            'If ((InStr(Status, "S ?") > 0)) Then
                'errormessage = "The AF degaussing unit reports status unknown." & _
                    vbCrLf & vbCrLf & PollText & vbCrLf & _
                    vbCrLf & "Execution has been paused and Ramp Down command sent. Please check machine."
                'SendCommand "DERD"
                'Flow_Pause
                'SetCodeLevel CodeRed
                'frmSendMail.MailNotification "AF Error", errormessage, CodeRed
                'MsgBox errormessage
                'SetCodeLevel CodeGreen, True
            'ElseIf (InStr(Status, "S Z") > 0) Or (InStr(Status, "S T")) Or (InStr(Status, "A    0") > 0) Then
                'finished = True
            'End If
            'delay = Timer
       'End If
        'lag = 1440 * (Now - startTime)
        'If lag > 2.5 Then
        '=======================================================================
        'Oxy system takes 41 s ramp up to 900, so this is plenty for full cycle.
        'If still looping after 90 sec, there is trouble....
        'SWB:  TEST CODE TO PREVENT HALTS WHEN AF OK BUT NO DONE MSG
        If (Timer - startTime) > 90 Then
            totalsecs = totalsecs + (Timer - startTime)
            'DO WE NEED TO PANIC?
            PollText = SendStatus
            If InStr(PollText, "A    0") > 0 Then
                ' unit has reset! No need to panic, but does reflect a bug with the unit.
                errormessage = "No DONE from AF box for " & Format$(totalsecs, "0.0") & _
                " seconds " & "on axis " & currentAxis & " at amplitude " & Format(currentCalAmp, 0) & "." & _
                " Target amplitude reported as zero, so unit appears to have reset. Execution will continue. " & PollText
                If DEBUG_MODE Then
                   frmDebug.Msg "From PollAF: " & errormessage
                End If
                frmSendMail.MailNotification "AF Alert", errormessage, CodeYellow
                LockAF False
                ActiveCoilSystem = 0
                currentDelay = -1
                currentRampRate = -1
                currentUncalAmp = -1
                currentCalAmp = -1
                currentAxis = vbNullString
               finished = True  'so that we exit PollAf
            ElseIf InStr(PollText, "S Z") > 0 Then
               'NO!
                errormessage = "No DONE from AF box for " & Format$(totalsecs, "0.0") & " seconds." & _
                "on axis " & currentAxis & " at amplitude " & Format(currentCalAmp, 0) & "." & _
                " But, AF status=zero. Execution will continue. " & PollText
               If DEBUG_MODE Then
                   frmDebug.Msg "From PollAF: " & errormessage
               End If
               frmSendMail.MailNotification "AF Alert", errormessage, CodeYellow
               finished = True  'so that we exit PollAf
            Else
                'YES!  CALL 911
                errormessage = "The AF degaussing coil has not responded for " & Format$(totalsecs, "0.0") & " seconds" & _
                "on axis " & currentAxis & " at amplitude " & Format(currentCalAmp, 0) & "." & vbCrLf & vbCrLf & _
                vbCrLf & "Execution has been paused and Ramp Down command sent. Please check machine. " & PollText
                SendCommand "DERD"
                Flow_Pause
                SetCodeLevel CodeRed
                frmSendMail.MailNotification "AF Error", errormessage, CodeRed
                MsgBox errormessage
                SetCodeLevel CodeGreen, True
                'reset clock, send error msg every 90 secs if really stuck
                startTime = Timer
                'NOTE: will exit loop by getting Z response to DERD.  Loop till then
            End If
        End If
    Loop
    'Status = SendStatus
    'WriteAF PollText, "PollText"
    'WriteAF Status, "Status"
End Sub

Private Sub WriteAF(txt As String, Label As String)
    ' Subroutine added by L Carporzen (August 2007) to record the communications with the 2G degausser.
    Dim filenum As Integer
    Dim FileName As String
    filenum = FreeFile
    FileName = Prog_DefaultPath & "\AFsequence.txt"
    On Error GoTo oops
    Open FileName For Append As #filenum
    Print #filenum, txt; ","; Label
    Close #filenum
    GoTo stillworking
oops:
    MsgBox "Unable to write to " & FileName & "!"
stillworking:
End Sub

Private Function GetResponse() As String
    Dim Delay As Double
    Dim inputchar As String
    Dim responsepoint As Integer
    Dim inputlen As Integer
    If Not EnableAF Then Exit Function
    Delay = Timer   ' Set delaystart time.
    inputchar = vbNullString
    Do While Not NOCOMM_MODE
        DoEvents
        If MSCommAF.InBufferCount > 0 Then
            'OLD CODE==============
            'inputchar = inputchar + MSCommAF.Input
            'inputlen = Len(inputchar)
            'If inputlen > 4 Then
            '    responsepoint = InStr(Mid$(inputchar, 3), vbCrLf) + 2
             '  If Right$(inputchar, 2) = vbCrLf And responsepoint < inputlen - 5 Then
             '       inputchar = Mid$(inputchar, responsepoint + 2)
             '   Exit Do
              '  End If
            'End If
            '========================
            'NEW CODE BY SWB 3/16/06
            'Empty the bugger, I mean buffer.
            'Little delay in loop to give AF time to send.
            'Delay upped from 0.01 to 0.05  3/20/06 SWB
            'Do While MSCommAF.InBufferCount > 0
             '    inputchar = inputchar + MSCommAF.Input
              '   DelayTime 0.05
            'Loop
            DelayTime 0.5  '0.3 needed to get all of status string
            Do While MSCommAF.InBufferCount > 0
                 inputchar = inputchar + MSCommAF.Input
            Loop
            'Now, handle DSS case.  Want substring after DSS(CR)
            responsepoint = InStr(inputchar, "DSS")
            If responsepoint > 0 Then
              inputchar = Mid$(inputchar, responsepoint + 4)
            End If
            Exit Do     'the usual way out, something in and read
            'note: cannot have the Msg box after 3 sec timeout
            'because this always happens while polling from PollAF.
            'Msg box interrupts flow.
            'But need to break out of GetResponse to go back to
            'timing loop in PollAF in case of AF problem
            '========================
        End If
        If Timer < Delay Then Delay = Delay - 86400
        If Timer - Delay > 3 Then
            Exit Do
            'MsgBox "Timeout in AF GetResponse routine"
        End If
    Loop
    MSCommAF.OutBufferCount = 0
    'line below commmented out by SWB  3/20/06
    'MSCommAF.InBufferCount = 0
    GetResponse = inputchar
    InputText = inputchar
    If DEBUG_MODE And Len(inputchar) > 0 Then frmDebug.Msg "COM " & Str$(MSCommAF.CommPort) & " in: " & inputchar
End Function

Public Sub Connect()
    If Not EnableAF Then Exit Sub
    If MSCommAF.PortOpen = False And Not NOCOMM_MODE Then
        On Error GoTo ErrorHandler  ' Enable error-handling routine.
        MSCommAF.CommPort = COMPortAf
        MSCommAF.Settings = "1200,n,8,1"
        MSCommAF.SThreshold = 1
        MSCommAF.RThreshold = 0
        MSCommAF.inputlen = 1
        MSCommAF.PortOpen = True
        On Error GoTo 0 ' Turn off error trapping.
        If MSCommAF.PortOpen = True Then
            cmdConnect.Caption = "Disconnect"
            ' disable the other connection buttons here until com is free
        End If
    End If
Exit Sub        ' Exit to avoid handler.
ErrorHandler:   ' Error-handling routine.
    Select Case Err.number  ' Evaluate error number.
        Case 8002
            MsgBox "Invalid Port Number"
        Case 8005
            MsgBox "Port already open" + Chr(13) + "(Already is use?)"
        Case 8010
            MsgBox "The hardware is not available (locked by another device)"
        Case 8012
            MsgBox "The device is not open"
        Case 8013
            MsgBox "The device is already open"
        Case Else
            MsgBox "Unknown error trying to Connect Comm Port"
    End Select
End Sub

Public Sub Disconnect()
    If MSCommAF.PortOpen = True Then
        MSCommAF.PortOpen = False
        cmdConnect.Caption = "Connect"
    End If
End Sub

Private Sub optActiveAxial_Click()
    SetActiveCoilSystem AxialCoilSystem
End Sub

Private Sub optActiveTransverse_Click()
    SetActiveCoilSystem TransverseCoilSystem
End Sub

Private Sub MultiRampFileSave(ByRef DataArray() As Double, _
                                ByVal PtsPerFile As Long, _
                                ByVal FolderPath As String, _
                                ByVal FolderName As String, _
                                ByVal CurTime, _
                                Optional ByVal BackupFolderPath As String = "", _
                                Optional SaveMaxAmp As Boolean = False, _
                                Optional PtsWindowForFindingMax As Long = 0)
                                
    Dim i, j, k, N As Long
    Dim NumBackupFiles As Long
    Dim Temp, MaxAmp As Double
    
    Dim fso As FileSystemObject
    Dim DataStream As TextStream
    Dim MaxAmpStream As TextStream
    Dim MaxAmpFileName As String
    Dim DataFileName As String
    Dim CoilString As String
    
    If Not fso.FolderExists(FolderPath) Then
    
        On Error GoTo BadLocalFolderPath:
        
            fso.CreateFolder (FolderPath)
            
        On Error GoTo 0
    
    End If
    
    If Not fso.FolderExists(FolderPath & FolderName) Then
    
        On Error GoTo BadLocalFolderName:
        
            fso.CreateFolder (FolderPath & FolderName)
            
        On Error GoTo 0
    
    End If
    
    If SaveMaxAmp = True Then
    
        MaxAmpFileName = "AFRamplitude_" & Format(CurTime, "MM-DD-YY_HH-MM-SS") & ".csv"
    
        On Error GoTo MaxAmpFileCreateError:
    
            Set MaxAmpStream = fso.CreateTextFile(FolderPath & FolderName & MaxAmpFileName, True)
            
        On Error GoTo 0
        
        AmpStream.WriteLine AFSystem & " AF Ramp Amplitudes on " & CoilString & " coil"
        AmpStream.WriteLine Format(CurTime, "long date") & ", " & Format(CurTime, "long time")
        AmpStream.WriteBlankLines (1)
        AmpStream.WriteLine "Sliding Window =" & "," & Trim(Str(PtsWindowForFindingMax))
        AmpStream.WriteLine "Point #,Donut Voltage"

    Else
    
        Set MaxAmpStream = Nothing

    End If
    
    N = UBound(datarray)
    
    'Create first data file
    DataFileName = "AFramp_pts0-" & Trim(Str(PtsPerFile)) & "_" & Format(CurTime, "MM-DD-YY_HH-MM-SS") & ".csv"
    
    On Error GoTo DataFileCreateError:
    
        Set DataStream = fso.CreateTextFile(FolderPath & FolderName & DataFileName, True)
        
    On Error GoTo 0
    
    'Make File Save Progress bar visible
    Me.progressBarAFFileSave.Min = 0
    Me.progressBarAFFileSave.Max = 32767
    Me.progressBarAFFileSave.value = 0
    Me.progressBarAFFileSave.Visible = True
    Me.lblFileSaveProg.Caption = "File Save % Complete:   0%"
    Me.lblFileSaveProg.Visible = True
    Me.refresh
    
    If Me.optActiveAxial.value = True Then
    
        CoilString = "Axial"
        
    Else
    
        CoilString = "Transverse"
        
    End If
    
    If Not fso.FolderExists(FolderPath) Then
        
        fso.CreateFolder FolderPath
        
    End If
    
    DataStream.WriteLine AFSystem & " AF Ramp on " & CoilString & " coil"
    DataStream.WriteLine Format(CurTime, "long date") & ", " & Format(CurTime, "long time")
    DataStream.WriteBlankLines (1)
    DataStream.WriteLine "From = ,0"
    DataStream.WriteLine "To = ," & Trim(Str(PtsPerFile))
    DataStream.WriteLine "Point #,Donut Voltage"
    
    j = PtsWindowForFindingMax
                    
    N = UBound(MonitorArray)
    
    'Initialize k = # pts per file + 1
    k = PtsPerFile + 1
    
    'Initialize max amp to the absolute value first point of the monitor array
    MaxAmp = Abs(DataArray(0))
    
    For i = 0 To N - 1
    
        If i Mod 5000 = 0 Then
        
            'Need to update progress
            Me.progressBarAFFileSave.value = CInt(32767 * i / (N - 1))
            PercComplete = Trim(Str(CInt(100 * i / (N - 1))))
            Do While Len(PercComplete) < 4
            
                PercComplete = " " & PercComplete
                
            Loop
            
            Me.lblFileSaveProg.Caption = "File Save % Complete:" & PercComplete & "%"
    
            Me.refresh
    
        End If
    
        Temp = DataArray(i)
    
        DataStream.WriteLine Trim(Str(i)) & "," & Trim(Str(Temp))
        
        If MaxAmp < Abs(Temp) And SaveMaxAmp = True Then MaxAmp = Abs(Temp)
        
        j = j - 1
        k = k - 1
        
        If j = 0 And SaveMaxAmp = True Then
        
            AmpStream.WriteLine Trim(Str(i)) & "," & Trim(Str(MaxAmp))
            
            j = PtsWindowForFindingMax
            
            MaxAmp = -100
            
        End If
        
        If k = 0 And Not i = N - 1 Then
        
            'Need to close the current text stream, create a new text file,
            'and open it up for writing
            DataStream.Close
            
            'Create next data file
            'Adjust name to refltect the current point we're on and adjust the
            'point to for the final file
            If i + PtsPerFile > N - 1 Then
                
                DataFileName = "AFramp_pts" & Trim(Str(i + 1)) & "-" & Trim(Str(N - 1)) & _
                                "_" & Format(CurTime, "MM-DD-YY_HH-MM-SS") & ".csv"
                                
            Else
            
                DataFileName = "AFramp_pts" & Trim(Str(i + 1)) & "-" & Trim(Str(i + PtsPerFile)) & _
                                "_" & Format(CurTime, "MM-DD-YY_HH-MM-SS") & ".csv"
                                
            End If
            
            On Error GoTo DataFileCreateError:
            
                Set DataStream = fso.CreateTextFile(FolderPath & FolderName & DataFileName, True)
                
            On Error GoTo 0
            
            DataStream.WriteLine AFSystem & " AF Ramp on " & CoilString & " coil"
            DataStream.WriteLine Format(CurTime, "long date") & ", " & Format(CurTime, "long time")
            DataStream.WriteBlankLines (1)
            DataStream.WriteLine "From = ," & Trim(Str(i + 1))
            DataStream.WriteLine "To = ," & Trim(Str(i + PtsPerFile))
            DataStream.WriteLine "Point #,Donut Voltage"
            
            k = PtsPerFile
            
        End If
        
    Next i
    
    'Close the final file string
    DataStream.Close
                
    If BackupFolderPath <> "" Then
    
        'Now need to copy the ramp files to the remote backup path
        
        
        'Change Value on Progress bar to 0
        Me.progressBarAFFileSave.value = 0
        
        'Determine the number of files to backup
        NumBackupFiles CInt(N / PtsPerFile) + 1
        
        If SaveMaxAmp = True Then NumBackupFiles = NumBackupFiles + 1
        
        'Change Caption in lblFileSaveprog
        Me.lblFileSaveProg.Caption = "Creating Backup Folder..."
        
        'See if there is a main backup ramp data folder yet,
        'if not, create it
        If Not fso.FolderExists(BackupFolderPath) Then
        
            On Error GoTo BadBackupFolderPath:
            
                fso.CreateFolder BackupFolderPath
                
            On Error GoTo 0
            
        End If
        
        'Create the folder for this AF ramp's worth of data files
        If Not fso.FolderExists(BackupFolderPath & FolderName) Then
        
            On Error GoTo CreateBackupFolderError:
            
                fso.CreateFolder BackupFolderPath & FolderName
        
            On Error GoTo 0
            
        End If
        
        'Change Caption on lblFileSaveProg
        Me.lblFileSaveProg.Caption = "Backing up Files: " & NumBackupFiles & " Files remaining..."
        
            
        
                
    'Remove Progress Bar and file save progress lable, and reset their values to 0% progress
    Me.lblFileSaveProg.Visible = False
    Me.progressBarAFFileSave.Visible = False
    Me.lblFileSaveProg.Caption = ""
    Me.progressBarAFFileSave.value = 0
    
    'Uncheck the Monitor & Save AF Ramp box
    chkVerbose.value = Unchecked

    Exit Sub
    
BadLocalFolderPath:

    errormessage = "Could not find/access AF Ramp main data folder. Code Execution paused." & vbNewLine & _
                    "Current path = " & FolderPath & vbNewLine & _
                    "Error Code = " & Trim(Str(Err.number))
    
    SetCodeLevel CodeYellow
    frmSendMail.MailNotification "Create Folder Error", errormessage, CodeYellow
    
    MsgBox errormessage
    
    SetCodeLevel StatusCodeColorLevelPrior, True
    
    Exit Sub

BadLocalFolderName:

    'This is an internal code error that should never happen
    'send MsgBox if this occurs
    errormessage = "Could not create data folder for this AF Ramp. Code Execution paused.  " & vbNewLine & _
                    "Current Folder Name: " & FolderPath & FolderName & vbNewLine & _
                    "Error Code = " & Trim(Str(Err.number))
    
    SetCodeLevel CodeYellow
    frmSendMail.MailNotification "Create Folder Error", errormessage, CodeYellow
    
    MsgBox errormessage
    
    SetCodeLevel StatusCodeColorLevelPrior, True
    
    Exit Sub
    
MaxAmpFileCreateError:

    errormessage = "Could not create AF Ramp amplitudes data file. Code Execution paused.  " & vbNewLine & _
                    "File Path = " & FolderPath & FolderName & MaxAmpFileName & vbNewLine & _
                    "Error Code = " & Trim(Str(Err.number))
                    
    SetCodeLevel CodeYellow
    frmSendMail.MailNotification "Create Folder Error", errormessage, CodeYellow
    
    MsgBox errormessage
    
    SetCodeLevel StatusCodeColorLevelPrior, True
    
    Exit Sub
    
DataFileCreateError:

    errormessage = "Could not create AF ramp Data File. Code Execution paused.  " & vbNewLine & _
                    "File Path = " & FolderPath & FolderName & DataFileName & vbNewLine & _
                    "Error Code = " & Trim(Str(Err.number))
                    
    SetCodeLevel CodeYellow
    frmSendMail.MailNotification "Create Folder Error", errormessage, CodeYellow
    
    MsgBox errormessage
    
    SetCodeLevel StatusCodeColorLevelPrior, True
    
    Exit Sub

BadBackupFolderPath:

    errormessage = "Could not access/create Main Backup AF Ramp Data Folder. Code Execution will continue." & vbNewLine & _
                    "Backup Folder Path = " & BackupFolderPath & vbNewLine & _
                    "Error Code = " & Trim(Str(Err.number))

    frmSendMail.MailNotification "Create Folder Error", errormessage, CodeYellow
        
    Exit Sub

CreateBackupFolderError:

    'This error should only happen if there are permissions issues with creating folders
    'on the backup drive
    errormessage = "Could not create backup folder for this AF Ramp's set of data files. Code Execution will continue.  " & vbNewLine & _
                    "Backup Folder Name: " & FolderPath & FolderName & vbNewLine & _
                    "Error Code = " & Trim(Str(Err.number))
    
    frmSendMail.MailNotification "Create Folder Error", errormessage, CodeYellow
    
    Exit Sub

End Sub

