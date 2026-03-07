VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmAF_2G 
   Caption         =   "2G AF Degausser"
   ClientHeight    =   5880
   ClientLeft      =   5880
   ClientTop       =   2145
   ClientWidth     =   7350
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5880
   ScaleWidth      =   7350
   Begin VB.PictureBox picAdBox 
      Height          =   495
      Left            =   120
      ScaleHeight     =   435
      ScaleWidth      =   7035
      TabIndex        =   42
      Top             =   5280
      Visible         =   0   'False
      Width           =   7095
   End
   Begin VB.CheckBox chkVerbose 
      Caption         =   "AF Debug Mode"
      Height          =   192
      Left            =   4080
      TabIndex        =   15
      Top             =   4080
      Width           =   1692
   End
   Begin VB.CommandButton cmdTemp 
      Caption         =   "Refresh T"
      Height          =   255
      Left            =   2040
      TabIndex        =   14
      Top             =   2760
      Width           =   1695
   End
   Begin VB.TextBox txtTemp2 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   720
      TabIndex        =   13
      Top             =   2640
      Width           =   735
   End
   Begin VB.TextBox txtTemp1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   720
      TabIndex        =   12
      Top             =   2280
      Width           =   735
   End
   Begin MSCommLib.MSComm MSCommAF 
      Left            =   4560
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
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   1335
   End
   Begin VB.CommandButton cmdSendStatus 
      Caption         =   "Send Status"
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   0
      Width           =   1335
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   3840
      TabIndex        =   2
      Top             =   0
      Width           =   972
   End
   Begin VB.TextBox txtSampHeight 
      Height          =   285
      Left            =   6720
      TabIndex        =   7
      Text            =   "0"
      Top             =   120
      Width           =   495
   End
   Begin VB.TextBox OutputText 
      Height          =   285
      Left            =   720
      TabIndex        =   3
      Top             =   600
      Width           =   852
   End
   Begin VB.TextBox InputText 
      Height          =   288
      Left            =   2280
      TabIndex        =   4
      Top             =   600
      Width           =   1932
   End
   Begin VB.CheckBox chkLocked 
      Caption         =   "Locked"
      Height          =   252
      Left            =   2280
      TabIndex        =   26
      Top             =   960
      Width           =   1212
   End
   Begin VB.CommandButton cmdManualAxialAF 
      Caption         =   "Manual Axial AF"
      Height          =   375
      Left            =   5400
      TabIndex        =   8
      Top             =   480
      Width           =   1815
   End
   Begin VB.CommandButton cmdManualTransverseAF 
      Caption         =   "Manual Transverse AF"
      Height          =   375
      Left            =   5400
      TabIndex        =   9
      Top             =   960
      Width           =   1815
   End
   Begin VB.Frame frameActiveCoil 
      Caption         =   "Active Coil System"
      Height          =   852
      Left            =   120
      TabIndex        =   27
      Top             =   1200
      Width           =   1812
      Begin VB.OptionButton optActiveAxial 
         Caption         =   "Axial"
         Height          =   252
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1572
      End
      Begin VB.OptionButton optActiveTransverse 
         Caption         =   "Transverse"
         Height          =   252
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   1572
      End
   End
   Begin VB.TextBox txtUncalAmplitude 
      Height          =   288
      Left            =   4320
      TabIndex        =   10
      Top             =   1560
      Width           =   852
   End
   Begin VB.CommandButton cmdSetUncalAmp 
      Caption         =   "Set Amplitude w/o Calibration"
      Height          =   492
      Left            =   5400
      TabIndex        =   11
      Top             =   1440
      Width           =   1692
   End
   Begin VB.CommandButton cmdCleanCoils 
      Caption         =   "Clean Coils"
      Height          =   372
      Left            =   4080
      TabIndex        =   24
      Top             =   3480
      Width           =   1452
   End
   Begin VB.TextBox txtAmplitude 
      Height          =   288
      Left            =   1200
      TabIndex        =   16
      Top             =   3360
      Width           =   852
   End
   Begin VB.CommandButton cmdConfigAmplitude 
      Caption         =   "Set Amplitude"
      Height          =   252
      Left            =   2160
      TabIndex        =   17
      Top             =   3360
      Width           =   1572
   End
   Begin VB.ComboBox cmbDelay 
      Height          =   315
      Left            =   1320
      TabIndex        =   18
      Top             =   3840
      Width           =   732
   End
   Begin VB.CommandButton cmdConfigDelay 
      Caption         =   "Set Delay"
      Height          =   252
      Left            =   2160
      TabIndex        =   21
      Top             =   3840
      Width           =   1572
   End
   Begin VB.ComboBox cmbRampRate 
      Height          =   315
      Left            =   1320
      TabIndex        =   19
      Top             =   4320
      Width           =   732
   End
   Begin VB.CommandButton cmdConfigureRampRate 
      Caption         =   "Set Ramp"
      Height          =   252
      Left            =   2160
      TabIndex        =   22
      Top             =   4320
      Width           =   1572
   End
   Begin VB.ComboBox cmbRampMode 
      Height          =   315
      Left            =   1440
      TabIndex        =   20
      Top             =   4800
      Width           =   612
   End
   Begin VB.CommandButton cmdExecuteRamp 
      Caption         =   "Execute Ramp"
      Height          =   252
      Left            =   2160
      TabIndex        =   23
      Top             =   4800
      Width           =   1572
   End
   Begin VB.TextBox txtWaitingTime 
      Height          =   285
      Left            =   4080
      TabIndex        =   25
      Text            =   "0"
      Top             =   4560
      Width           =   375
   End
   Begin VB.Line Line8 
      BorderColor     =   &H8000000C&
      X1              =   120
      X2              =   3840
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Line Line7 
      BorderColor     =   &H8000000C&
      X1              =   120
      X2              =   6120
      Y1              =   5160
      Y2              =   5160
   End
   Begin VB.Line Line6 
      BorderColor     =   &H8000000C&
      X1              =   6120
      X2              =   6120
      Y1              =   3240
      Y2              =   5160
   End
   Begin VB.Line Line5 
      BorderColor     =   &H8000000C&
      X1              =   120
      X2              =   120
      Y1              =   3240
      Y2              =   5160
   End
   Begin VB.Line Line4 
      BorderColor     =   &H8000000C&
      X1              =   120
      X2              =   120
      Y1              =   2160
      Y2              =   3120
   End
   Begin VB.Line Line3 
      BorderColor     =   &H8000000C&
      X1              =   120
      X2              =   6120
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Line Line2 
      BorderColor     =   &H8000000C&
      X1              =   3840
      X2              =   3840
      Y1              =   2160
      Y2              =   3120
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000C&
      X1              =   120
      X2              =   3840
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Label Label13 
      Caption         =   "Trans."
      Height          =   255
      Left            =   240
      TabIndex        =   41
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label Label12 
      Caption         =   "Axial"
      Height          =   255
      Left            =   240
      TabIndex        =   40
      Top             =   2280
      Width           =   495
   End
   Begin VB.Label lblAFtooHot 
      Caption         =   "The AF unit is too hot so let's pause a little bit..."
      Height          =   495
      Left            =   2040
      TabIndex        =   39
      Top             =   2280
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label11 
      Caption         =   "°C"
      Height          =   255
      Left            =   1560
      TabIndex        =   38
      Top             =   2640
      Width           =   255
   End
   Begin VB.Label Label10 
      Caption         =   "°C"
      Height          =   255
      Left            =   1560
      TabIndex        =   37
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Label9 
      Caption         =   "Sample Height (cm):"
      Height          =   255
      Left            =   5160
      TabIndex        =   28
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Output:"
      Height          =   255
      Left            =   120
      TabIndex        =   29
      Top             =   600
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Input:"
      Height          =   255
      Left            =   1800
      TabIndex        =   30
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label7 
      Caption         =   "Uncalibrated Amplitude:"
      Height          =   255
      Left            =   2400
      TabIndex        =   31
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Amplitude:"
      Height          =   255
      Left            =   240
      TabIndex        =   32
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Delay:"
      Height          =   255
      Left            =   240
      TabIndex        =   33
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Ramp:"
      Height          =   255
      Left            =   240
      TabIndex        =   34
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "Ramp Mode:"
      Height          =   255
      Left            =   240
      TabIndex        =   35
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Label Label8 
      Caption         =   "Waiting time (in s) between ramps"
      Height          =   375
      Left            =   4560
      TabIndex        =   36
      Top             =   4560
      Width           =   1335
   End
End
Attribute VB_Name = "frmAF_2G"
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
Dim CoilsLocked As Boolean
Dim currentDelay As Integer
Dim currentRampRate As Integer
Dim currentUncalAmp As Double
Dim currentCalAmp As Double
Dim currentAxis As String
Dim Feedback As String ' (August 2007 L Carporzen) Allow to record in a file the AF communications
    
'(February 2011, I Hilburn)
' Added in these two form locals to store the ramp type (calibrated or uncalibrated)
' and the Magnetic field target if the the ramp is calibrated
' this is to support the Alternate AF monitor module code, so that the new 2G ramp data
' folder that is created from the record of the AF ramp has the correct Field value
' or 2G count
Dim is2GCalRamp As Boolean
Dim Cal2GTarget As Double

Private Sub chkLocked_Click()
    
    If Me.chkLocked.value = Checked Then
    
        CoilsLocked = True
        Me.optActiveAxial.Enabled = False
        Me.optActiveTransverse.Enabled = False
        cmdConfigAmplitude.Enabled = False
        cmdConfigDelay.Enabled = False
        cmdConnect.Enabled = False
        cmdSendStatus.Enabled = False
        cmdExecuteRamp.Enabled = False
        cmdSetUncalAmp.Enabled = False
        cmdConfigureRampRate.Enabled = False
        
    End If
    
    If Me.chkLocked.value = Unchecked Then
    
        CoilsLocked = False
        Me.optActiveAxial.Enabled = True
        Me.optActiveTransverse.Enabled = True
        cmdConfigAmplitude.Enabled = True
        cmdConfigDelay.Enabled = True
        cmdConnect.Enabled = True
        cmdSendStatus.Enabled = True
        cmdExecuteRamp.Enabled = True
        cmdSetUncalAmp.Enabled = True
        cmdConfigureRampRate.Enabled = True
        
    End If
    
End Sub

Private Sub chkVerbose_Click()

    If chkVerbose.value = Checked Then
    
        'Tell user that this will cause LOTS of data to be saved
        MsgBox "Turning on AF Debug mode will result in the creation of a data folder " & _
                " with ~10-100 Mb worth of files recording the next 2G AF ramp." & _
                vbNewLine & vbNewLine & "Before you do a AF Debug ramp, make sure: " & _
                vbNewLine & "1) Your " & _
                "hardware is properly configured - with the voltage from the BNC terminal " & _
                "of the green donut ammeter in the grey capacitor box connected to both the " & _
                "2G AF Box signal monitor input -AND- the correct analog input port on the " & _
                AltAFMonitorChan.BoardName & " board in your computer." & _
                vbNewLine & vbNewLine & "2) Your " & _
                "Paleomag software settings are correct as well.  Make sure that in the frmSettings " & _
                "AF tab, that you have entered a data directory and a backup directory in which the " & _
                "2G AF Ramp data will be stored.  Additionally, in this tab, you must have " & _
                "selected the correct analog input port on the " & AltAFMonitorChan.BoardName & _
                " board for the computer " & _
                "to look for the AF green-donut ammeter signal." & vbNewLine & vbNewLine & _
                "If these setting and hardware changes have not been made, please Uncheck the AF " & _
                "Debug Mode checkbox.", , _
                "AF Debug Mode Configuration Check"
                
    End If

End Sub

Public Sub CleanCoils()
    ExecuteRamp "C", AxialCoilSystem, AfAxialMax, AFDelay, AFRampRate
    ExecuteRamp "C", TransverseCoilSystem, AfTransMax, AFDelay, AFRampRate
End Sub

Private Sub cmdCleanCoils_Click()
    CleanCoils
End Sub

Private Sub cmdClose_Click()
    Me.Hide
End Sub

Private Sub cmdConfigAmplitude_Click()
    
    'Set the status flag indicating that this is a calibrated ramp
    is2GCalRamp = True
    
    'Store the intended Gauss Target
    Cal2GTarget = val(txtAmplitude.text)
    
    'Change the Field value (double type) to 2G count value (integer type)
    ConfigureAmplitude val(txtAmplitude.text)
    
End Sub

Private Sub cmdConfigDelay_Click()
    ConfigureDelay Int(val(cmbDelay))
End Sub

Private Sub cmdConfigureRampRate_Click()
    ConfigureRampRate Int(val(cmbRampRate))
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

Private Sub cmdManualAxialAF_Click()
    ManualAxialAF
End Sub

Private Sub cmdManualTransverseAF_Click()
    ManualTransverseAF
End Sub

Private Sub cmdSendStatus_Click()
    SendStatus
End Sub

Public Sub cmdSetUncalAmp_Click()
    
    'This is an uncalibrated ramp, set the is2GCalRamp flag = False
    is2GCalRamp = False

    SetAmplitude val(txtUncalAmplitude)
End Sub

Private Sub cmdTemp_Click()
    Dim Temp1 As Double ' (February 2010 L Carporzen) Monitor temperature of the coils before executing AF
    Dim Temp2 As Double
    
    '(July 2010 - I Hilburn) Replaces old frmMCC with frmDAQ_Comm using the Channel object variables
    If EnableT1 Then
    
        frmDAQ_Comm.DoDAQIO AnalogT1, Temp1
        
        Temp1 = Temp1 * TSlope - Toffset
        
    End If
        
    txtTemp1 = Format$(Temp1, "##0.00")
    
    If EnableT2 Then
    
        frmDAQ_Comm.DoDAQIO AnalogT2, Temp2
        
        Temp2 = Temp2 * TSlope - Toffset
    
    End If
        
    txtTemp2 = Format$(Temp2, "##0.00")
    
End Sub

Private Sub ConfigureAmplitude(ByVal Amplitude As Double, Optional ByVal AFCoilSystem As Integer = -128)
    ' Calibrate and set amplitude.
    Dim AFLevel As Double
    If CoilsLocked Then
        SetForm
        Exit Sub
    End If
    If AFCoilSystem = -128 Then AFCoilSystem = ActiveCoilSystem
        currentCalAmp = Amplitude
    SetAmplitude CInt((FindXCalibValue(Amplitude, AFCoilSystem)))
End Sub

Public Sub ConfigureCoil(axis As String)
    If CoilsLocked Then
        SetForm
        Exit Sub
    End If
    
    If EnableAF = False Then Exit Sub
    
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
    If CoilsLocked Then
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
    If CoilsLocked Then Exit Sub
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
            MsgBox "The hardware is not available (coilslocked by another device)"
        Case 8012
            MsgBox "The device is not open"
        Case 8013
            MsgBox "The device is already open"
        Case Else
            MsgBox "Unknown error trying to Connect Comm Port"
    End Select
End Sub

Private Sub ConnectButton_Click()
    If MSCommAF.PortOpen Then
        Disconnect
    Else
        Connect
    End If
End Sub

Public Sub CycleWithHold(Optional ByVal HoldTime As Integer = 0, Optional AFCoilSystem As Integer = -128, _
    Optional ByVal Amplitude As Double = -1, _
    Optional RampRate As Integer = -1)
    Dim olddelay As Integer
    olddelay = currentDelay
    If HoldTime = 0 Then HoldTime = AFDelay
    ExecuteRamp "C", AFCoilSystem, Amplitude, HoldTime, RampRate
    If Not HoldTime = AFDelay Then ConfigureDelay olddelay
End Sub

Public Sub Disconnect()
        
    If EnableAF = False Then Exit Sub
    
    If MSCommAF.PortOpen = True Then
        MSCommAF.PortOpen = False
        cmdConnect.Caption = "Connect"
    End If
End Sub

Public Function ExecuteRamp(ByVal Mode As String, _
                            Optional AFCoilSystem As Integer = -128, _
                            Optional ByVal Amplitude As Double = -1, _
                            Optional Delay As Integer = -1, _
                            Optional RampRate As Integer = -1) As Boolean
                       
    Dim reply As String
    Dim ErrorMessage As String
    
    ' (February 2010 L Carporzen) Monitor temperature of the coils before executing AF
    Dim Temp1 As Double
    Dim Temp2 As Double
    Dim TWarning As Boolean
        
    ' (April 2010, I Hilburn) Additional code for monitoring the 2G Af Ramp
    Dim AFData() As Double
    Dim SineFit_Data() As Double
    Dim NoError As Boolean
    Dim FolderName As String
    Dim CurTime
    
    'Default return value of the function
    ExecuteRamp = False
    
    If Not MSCommAF.PortOpen And Not NOCOMM_MODE Then Connect
    
    If EnableAF = False Then
    
        'Whoops! Tell the user
        MsgBox "AF Module is currently disabled.  This AF ramp cycle will now abort.", , _
               "Whoops!"
        
               
        Exit Function
        
    End If
    
    If CoilsLocked Then
    
        MsgBox "AF unit is use.  Ramp execution is not possible."
        Exit Function
    End If
    
    If Not NOCOMM_MODE Then
        
        frmProgram.StatusBar "AF config", 2
        
        If (AFCoilSystem = AxialCoilSystem) Or _
           (AFCoilSystem = TransverseCoilSystem) _
        Then
            
            SetActiveCoilSystem AFCoilSystem
            
        End If
        
        If Amplitude >= 0 Then
        
            'User has input a Field amplitude, need to set
            'is2GCalRamp = True and store the target field value
            is2GCalRamp = True
            Cal2GTarget = Amplitude
            
            ConfigureAmplitude Amplitude
            
        End If
        
        If Delay > 0 Then ConfigureDelay Delay
        
        If RampRate > 0 Then ConfigureRampRate RampRate
        
        If (AFCoilSystem = AxialCoilSystem) Or _
           (AFCoilSystem = TransverseCoilSystem) Then
            
            SetActiveCoilSystem AFCoilSystem
                
        End If
                
        If Mode <> "U" And Mode <> "D" And Mode <> "C" Then Exit Function
        
        If Mode = "U" Then MsgBox "Ramping up without ramping down is dangerous..."
        
        ' (February 2010 L Carporzen) Monitor temperature of the coils before executing AF
        On Error GoTo oops
        
        TWarning = False
        
        '(July 2010 - I Hilburn) Replaces old frmMCC with frmDAQ_Comm using the Channel object variables
        If EnableT1 Then
        
            frmDAQ_Comm.DoDAQIO AnalogT1, Temp1
            
            Temp1 = Temp1 * TSlope - Toffset
        
        End If
        
        txtTemp1 = Format$(Temp1, "##0.00")
        
        '(July 2010 - I Hilburn) Replaces old frmMCC with frmDAQ_Comm using the Channel object variables
        If EnableT2 Then
        
            frmDAQ_Comm.DoDAQIO AnalogT2, Temp2
            
            Temp2 = Temp2 * TSlope - Toffset
        
        End If
        
        txtTemp2 = Format$(Temp2, "##0.00")
        
        'Check Temperature to see if it is not zeroed (gone within 20 deg of -1 * Toffset)
        If Not ValidSensorTemp(Temp1, Temp2) Then
        
            'Start code to tell user that the temp sensor values are bad
            NotifySensorError Temp1, Temp2
            
        End If
        
        If EnableT1 Or EnableT2 Then
        
        Do While Temp1 >= Thot Or Temp2 >= Thot
            
            frmAF_2G.ZOrder
            frmAF_2G.Show
            
            lblAFtooHot.Visible = True
            txtTemp1.BackColor = ColorOrange
            txtTemp2.BackColor = ColorOrange
            
            ErrorMessage = "The AF degaussing unit is above " & Thot & "°C: " & Format$(Temp1, "##0.00") & _
                "°C and " & Format$(Temp2, "##0.00") & "°C." & _
                vbCrLf & "Execution will restart soon."
            
            If TWarning = False Then frmSendMail.MailNotification "AF too hot", ErrorMessage, CodeYellow
            
            TWarning = True
            
            ' MsgBox "Pause... " & Temp1 & "°C " & Temp2 & "°C"
            ' Loop until the temperature which was above Thot decreases at least 5 degrees before restarting
            Do While Temp1 >= Thot - 5 Or Temp2 >= Thot - 5
                
                DelayTime (1)
                
                '(July 2010 - I Hilburn) Replaces old frmMCC with frmDAQ_Comm using the Channel object variables
                If EnableT1 Then
                
                    frmDAQ_Comm.DoDAQIO AnalogT1, Temp1
                    
                    Temp1 = Temp1 * TSlope - Toffset
                
                End If
                
                txtTemp1 = Format$(Temp1, "##0.00")
                
                '(July 2010 - I Hilburn) Replaces old frmMCC with frmDAQ_Comm using the Channel object variables
                If EnableT2 Then
                
                    frmDAQ_Comm.DoDAQIO AnalogT2, Temp2
                    
                    Temp2 = Temp2 * TSlope - Toffset
                
                End If
                
                txtTemp2 = Format$(Temp2, "##0.00")
                
                'Check Temperature to see if it is not zeroed (gone within 20 deg of -1 * Toffset)
                If Not ValidSensorTemp(Temp1, Temp2) Then
                
                    'Start code to tell user that the temp sensor values are bad
                    NotifySensorError Temp1, Temp2
                    
                End If
            
            Loop
        
        Loop
        
        End If
        
        txtTemp1.BackColor = RGB(255, 255, 255)
        txtTemp2.BackColor = RGB(255, 255, 255)
        
        lblAFtooHot.Visible = False
                
oops:
        LockAF True
        
'---------------------------------------------------------------------------------------------------------------------'
'---------------------------------------------------------------------------------------------------------------------'
'   (Mar, 2010 - I Hilburn)
'   Recording 2G AF Ramp through an Analog input port on the Alternate AF Monitor Board
'
'---------------------------------------------------------------------------------------------------------------------'

        If chkVerbose.value = Checked Then
        
            'Update AF status bar
            frmProgram.StatusBar "Config AF Monitor", 3
        
            'Start a background monitor process on the Alternate AF Monitor Board
            NoError = WaveForms("ALTAFMONITOR").ManageBackgroundProcess(AIFUNCTION, _
                                                                         AFData, _
                                                                         "2G AF monitor", _
                                                                         False, _
                                                                         True)
                                                                         
            'Update status of monitor attempt
            If NoError = True Then
            
                frmProgram.StatusBar "Config Successful", 3
        
            Else
            
                frmProgram.StatusBar "Config Failed", 3
        
            End If
            
            'Wait a half second for the user to register what was written
            PauseTill timeGetTime() + 500
        
        End If
               
        frmProgram.StatusBar "AF execute", 2
        
        'Now display in the status bar the level in Gauss that the coil is ramping to
        If is2GCalRamp = True Then
        
            frmProgram.StatusBar Trim(Me.txtAmplitude.text) & " " & modConfig.AFUnits, 3
            
        Else
        
            frmProgram.StatusBar Trim(Me.txtUncalAmplitude.text) & " 2G counts", 3
            
        End If
        
        cmbRampMode = Mode
       
        SendCommand "DER" + Mode, False
        'Feedback = GetResponse ' (August 2007 L Carporzen) Allow to record in a file the AF communications
        'WriteAF "DER" + Mode, "Mode"
        'WriteAF Feedback, "Answer"
        ExecuteRamp = PollAFUnit
        
        frmProgram.StatusBar vbNullString, 2
        frmProgram.StatusBar vbNullString, 3
        
        If NoError = True Then
           
            'The Background process has been started successfully
            'And the AF ramp has finished
            
            'Update Program form status
            frmProgram.StatusBar "Getting 2G Ramp Data...", 3
            
            'Terminate background process and get data
            NoError = WaveForms("ALTAFMONITOR").ManageBackgroundProcess(AIFUNCTION, _
                                                                        AFData(), _
                                                                        "2G AF Monitor", _
                                                                        True, _
                                                                        True, _
                                                                        1)
                                                                        
            'Update whether data retrieval was successful
            If NoError = True Then
            
                frmProgram.StatusBar "Data gotten!", 3
                
            Else
            
                frmProgram.StatusBar "Data Retrieval Failed!", 3
                                                              
            End If
            
            'Pause 0.5 seconds
            PauseTill timeGetTime() + 500
            
        End If
              
        'Now need to analyze the data
        If NoError = True Then
        
            'Sine fit the data
            frmADWIN_AF.DoSineFitAnalysis WaveForms("ALTAFMONITOR"), _
                                          AFData, _
                                          SineFit_Data, _
                                          Me.txtAmplitude, _
                                          5000
                                          
            FolderName = "2G "
                                          
            'Create a name for the new data folder
            If ActiveCoilSystem = AxialCoilSystem Then
            
                FolderName = FolderName & "Axial "
                
            Else
            
                FolderName = FolderName & "Trans "
                
            End If
            
            'Check to see if this is a calibrated ramp (trying to match a particular
            'magnetic field value) or an uncalibrate ramp to a particular 2G counts value
            If is2GCalRamp = True Then
            
                FolderName = FolderName & Trim(str(Cal2GTarget)) & "G - "
                
            Else
            
                FolderName = FolderName & Trim(str(Me.txtUncalAmplitude.text)) & _
                             " 2G Num - "
                             
            End If
                             
            'Store the current time
            CurTime = Now
                             
            'Now store a short date/Time string
            FolderName = FolderName & Format(CurTime, " YY-MM-DD HH-MM-SS") & "/"
                                                                                    
            'Save the data to file
            frmFileSave.MultiRampFileSave_MCC AFData, _
                                              CDbl(1 / WaveForms("ALTAFMONITOR").IORate), _
                                              1048000, _
                                              FolderName, _
                                              CurTime, _
                                              SineFit_Data, _
                                              (frmFileSave.chkBackupRampData.value = Checked), _
                                              True, _
                                              1000
                                          
        End If
       
    End If
        
    'If the AF ramp has failed, reset the connect to the 2G AF box
    If ExecuteRamp = False Then
        
        'Update status bar field
        frmProgram.StatusBar "AF Ramp Failed", 2
        frmProgram.StatusBar "Reseting connection", 3
        
        Disconnect
        DelayTime 3
        Connect
        
        frmProgram.StatusBar vbNullString, 2
        frmProgram.StatusBar vbNullString, 3
        
    End If
            
    LockAF False

End Function

Public Function FindXCalibValue(field As Double, _
                                Optional AFCoilSystem As Integer = -128) As Variant
    ' Find X (input to AF) from field

'-------------------------------------------------------------------------------------------------------------------------'
'-------------------------------------------------------------------------------------------------------------------------'
'
'   Code Mod
'   (July 2010, I Hilburn)
'
'   Changed code to match new array setup in modConfig and frmSettings used to store the AF and IRM calibration values
'   Instead of two arrays per coil (X & Y), there's just one N x 2 array where col 0 = X, and col 1 = Y.
'   Also, the array is now dynamic and can be larger than 25 elements (the auto-calibration routine makes it easy to generate
'   long AF calibration arrays).
'
'   This function has otherwise been preserved as it was originally written.  Since the calibration array will be
'   changed in frmSettings depending on the AF system, this one function can be used to convert field values to either
'   2G numbers or Voltage double values for both the 2G and ADWIN AF systems. (Yay!)
'
'   To make this function useable by either 2G or ADWIN AF system, the return value has been changed to a Variant
'
'-------------------------------------------------------------------------------------------------------------------------'
'-------------------------------------------------------------------------------------------------------------------------'

    Dim i As Integer
    Dim Slope As Double
    
    FindXCalibValue = -1
    
    If AFCoilSystem = -128 Then AFCoilSystem = ActiveCoilSystem
    
    If AFCoilSystem = AxialCoilSystem Then
        
        If field > AfAxialMax Then
            
            field = AfAxialMax
        
        ElseIf (field < AfAxialMin) And (field <> 0) Then
            
            field = AfAxialMin
        
        End If
        
        'Check to make sure AFAxialCount > 1
        If AFAxialCount <= 1 Then
        
            'User hasn't entered in enough calibration values
            MsgBox "Only one AF Axial Coil calibration value has been set. " & _
                   "Paleomag Code will now end." & vbNewLine & _
                   "Please restart the code and go to the Settings Window " & _
                   "to add more calibration values.", , _
                   "AF ERROR!"
                   
            End
                   
        End If
        
        For i = 1 To AFAxialCount
        
            If AFAxial(i, 1) = field Then
                
                FindXCalibValue = AFAxial(i, 0)
                
                Exit Function
                            
            ElseIf AFAxial(i - 1, 1) < field And AFAxial(i, 1) > field Then
                
                Slope = (AFAxial(i, 0) - AFAxial(i - 1, 0)) / (AFAxial(i, 1) - AFAxial(i - 1, 1))
                
                FindXCalibValue = AFAxial(i - 1, 0) + Slope * (field - AFAxial(i - 1, 1))
                
                Exit For
            
            ElseIf AFAxial(AFAxialCount, 1) < field Then
                
                'Field is larger than largest field in the calibration table
            
                Slope = (AFAxial(AFAxialCount, 0) - AFAxial(AFAxialCount - 1, 0)) / _
                        (AFAxial(AFAxialCount, 1) - AFAxial(AFAxialCount - 1, 1))
                        
                FindXCalibValue = AFAxial(AFAxialCount, 0) + _
                                  Slope * (field - AFAxial(AFAxialCount, 1))
            
            End If
        
        Next i
    
    ElseIf AFCoilSystem = TransverseCoilSystem Then
        
        If field > AfTransMax Then
            
            field = AfTransMax
        
        ElseIf (field < AfTransMin) And (field <> 0) Then
            
            field = AfTransMin
        
        End If
        
        'Check to make sure AFAxialCount > 1
        If AFTransCount <= 1 Then
        
            'User hasn't entered in enough calibration values
            MsgBox "Only one AF Trans Coil calibration value has been set. " & _
                   "Paleomag Code will now end." & vbNewLine & _
                   "Please restart the code and go to the Settings Window " & _
                   "to add more calibration values.", , _
                   "AF ERROR!"
                   
            End
                   
        End If
        
        For i = 1 To AFTransCount
            
            If AFTrans(i, 1) = field Then
                
                FindXCalibValue = AFTrans(i, 0)
                
                Exit Function
            
            ElseIf AFTrans(i - 1, 1) < field And AFTrans(i, 1) > field Then
                
                Slope = (AFTrans(i, 0) - AFTrans(i - 1, 0)) / (AFTrans(i, 1) - AFTrans(i - 1, 1))
                
                FindXCalibValue = AFTrans(i - 1, 0) + Slope * (field - AFTrans(i - 1, 1))
                
                Exit For
            
            ElseIf AFTrans(AFTransCount, 1) < field Then
                
                'Field is larger than largest field in the calibration table
            
                Slope = (AFTrans(AFTransCount, 0) - AFTrans(AFTransCount - 1, 0)) / _
                        (AFTrans(AFTransCount, 1) - AFTrans(AFTransCount - 1, 1))
                        
                FindXCalibValue = AFTrans(AFTransCount, 0) + _
                                  Slope * (field - AFTrans(AFTransCount, 1))
            
            End If
        
        Next i
    
    End If

End Function

Private Sub Form_Activate()

    If EnableAF = False Then
        
        'AF's not enabled, cannot Tune the AF coils
        'Tell user that calibration is turned off, but
        'can still edit values
        MsgBox "The AF module is currently disabled.  AF Ramp cycles " & _
               " cannot be performed now." & _
               "Whoops!"
               
        'Disable all the necessary buttons on the form
        Me.cmdSendStatus.Enabled = False
        Me.cmdCleanCoils.Enabled = False
        Me.cmdConfigAmplitude.Enabled = False
        Me.cmdConfigDelay.Enabled = False
        Me.cmdConfigureRampRate.Enabled = False
        Me.cmdConnect.Enabled = False
        Me.cmdExecuteRamp.Enabled = False
        Me.cmdManualAxialAF.Enabled = False
        Me.cmdManualTransverseAF.Enabled = False
        Me.cmdSetUncalAmp.Enabled = False
        Me.optActiveAxial.Enabled = False
        Me.optActiveTransverse.Enabled = False
        
    Else
    
        'Enable all the necessary buttons on the form
        Me.cmdSendStatus.Enabled = True
        Me.cmdCleanCoils.Enabled = True
        Me.cmdConfigAmplitude.Enabled = True
        Me.cmdConfigDelay.Enabled = True
        Me.cmdConfigureRampRate.Enabled = True
        Me.cmdConnect.Enabled = True
        Me.cmdExecuteRamp.Enabled = True
        Me.cmdManualAxialAF.Enabled = True
        Me.cmdManualTransverseAF.Enabled = True
        Me.cmdSetUncalAmp.Enabled = False
        Me.optActiveAxial.Enabled = True
        Me.optActiveTransverse.Enabled = True
        
    End If

    'Disable/Enable the Debug mode check-box
    Me.chkVerbose.Enabled = (modConfig.EnableAltAFMonitor = True)

    'First propagate the coilslocked coils state
    If CoilsLocked = True Then Me.chkLocked.value = Checked
    If CoilsLocked = False Then Me.chkLocked.value = Unchecked

    'If the window is loaded/activated, need to propagate
    'the current active coil settings to the radio buttons
    If ActiveCoilSystem = AxialCoilSystem Then
    
        Me.optActiveAxial.value = True
        
    ElseIf ActiveCoilSystem = TransverseCoilSystem Then
    
        Me.optActiveTransverse.value = True
        
    Else
    
        optActiveAxial.value = False
        optActiveTransverse.value = False
        
        ActiveCoilSystem = NoCoilSystem
        If AFSystem = "ADWIN" Then
        
            frmADWIN_AF.SetAFRelays
            
        End If
        
    End If

End Sub

Private Sub Form_Load()
        
    '(March 11, 2011 - I Hilburn)
    'Update all the form's icons with the program Ico file
    SetFormIcon Me
        
    If EnableAF = False Then
        
        'AF's not enabled, cannot use the AF coils
               
        'Disable all the necessary buttons on the form
        Me.cmdSendStatus.Enabled = False
        Me.cmdCleanCoils.Enabled = False
        Me.cmdConfigAmplitude.Enabled = False
        Me.cmdConfigDelay.Enabled = False
        Me.cmdConfigureRampRate.Enabled = False
        Me.cmdConnect.Enabled = False
        Me.cmdExecuteRamp.Enabled = False
        Me.cmdManualAxialAF.Enabled = False
        Me.cmdManualTransverseAF.Enabled = False
        Me.cmdSetUncalAmp.Enabled = False
        Me.optActiveAxial.Enabled = False
        Me.optActiveTransverse.Enabled = False
        
    Else
    
        'Enable all the necessary buttons on the form
        Me.cmdSendStatus.Enabled = True
        Me.cmdCleanCoils.Enabled = True
        Me.cmdConfigAmplitude.Enabled = True
        Me.cmdConfigDelay.Enabled = True
        Me.cmdConfigureRampRate.Enabled = True
        Me.cmdConnect.Enabled = True
        Me.cmdExecuteRamp.Enabled = True
        Me.cmdManualAxialAF.Enabled = True
        Me.cmdManualTransverseAF.Enabled = True
        Me.cmdSetUncalAmp.Enabled = False
        Me.optActiveAxial.Enabled = True
        Me.optActiveTransverse.Enabled = True
        
    End If
    
    'Disable/Enable the Debug mode check-box
    Me.chkVerbose.Enabled = (modConfig.EnableAltAFMonitor = True)
   
    'First propagate the coilslocked coils state
    If CoilsLocked = True Then Me.chkLocked.value = Checked
    If CoilsLocked = False Then Me.chkLocked.value = Unchecked

    'If the window is loaded/activated, need to propagate
    'the current active coil settings to the radio buttons
    If ActiveCoilSystem = AxialCoilSystem Then
    
        Me.optActiveAxial.value = True
        
    ElseIf ActiveCoilSystem = TransverseCoilSystem Then
    
        Me.optActiveTransverse.value = True
        
    Else
    
        optActiveAxial.value = False
        optActiveTransverse.value = False
        
        ActiveCoilSystem = NoCoilSystem
        If AFSystem = "ADWIN" Then
        
            frmADWIN_AF.SetAFRelays
            
        End If
        
    End If
        
'(July 2010 - I Hilburn) Commenting this line of code out.  It seems to be causing
'                        problems with the ADWIN AF code routine somehow.
'    activecoilsystem = axialcoilsystem
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
    cmbDelay.text = "1"
    currentDelay = -1
    currentRampRate = -1
    currentUncalAmp = -1
    
End Sub

Private Sub Form_Resize()
    Me.Height = 5880
    Me.Width = 8025
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If MSCommAF.PortOpen = True Then
        MSCommAF.PortOpen = False
    End If
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
    If DEBUG_MODE And Len(inputchar) > 0 Then frmDebug.msg "COM " & str$(MSCommAF.CommPort) & " in: " & inputchar
End Function

Private Sub LockAF(locking As Boolean)
    If locking Then
        CoilsLocked = True
        cmdConfigAmplitude.Enabled = False
        cmdConfigDelay.Enabled = False
        cmdConnect.Enabled = False
        cmdSendStatus.Enabled = False
        cmdExecuteRamp.Enabled = False
        cmdSetUncalAmp.Enabled = False
        cmdConfigureRampRate.Enabled = False
        chkLocked.value = vbChecked
    Else
        CoilsLocked = False
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

Public Sub ManualAxialAF()
    If frmVacuum.VacuumConnectOn = True Then ' (February 2008 L Carporzen) Manual Axial demag
        If val(txtAmplitude) <= 0 Then txtAmplitude = val(InputBox("What is amplitude (in Oe) of the axial demagnetization you want?", "Important!", txtAmplitude))
        If val(txtAmplitude) > AfAxialMax Then txtAmplitude = AfAxialMax
        If val(txtAmplitude) <= 0 Then Exit Sub
        txtSampHeight = val(InputBox("Doing a " & Int(val(txtAmplitude)) & _
        " Oe axial demagnetization" & vbCr & _
        "What is the height (in cm) of the sample?", "Important!", txtSampHeight))
        If val(txtSampHeight) = 0 Then Exit Sub ' (February 2010 L Carporzen) Do not run with 0 cm altitude (occurs when user press Cancel)
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
        txtSampHeight = val(InputBox("Doing a " & Int(val(txtAmplitude)) & _
        " Oe transverse demagnetization" & vbCr & _
        "What is the height (in cm) of the sample?", "Important!", txtSampHeight))
        If val(txtSampHeight) = 0 Then Exit Sub ' (February 2010 L Carporzen) Do not run with 0 cm altitude (occurs when user press Cancel)
        frmDCMotors.UpDownMove (AFPos + txtSampHeight * UpDownMotor1cm / 2), 1
        ExecuteRamp "C", TransverseCoilSystem, val(txtAmplitude), frmSettings.cmbAFDelay, frmSettings.cmbAFRampRate
        frmDCMotors.HomeToTop
    Else
        MsgBox "Aborted! Place a sample first..."
    End If
End Sub

Public Sub optActiveAxial_Click()
    SetActiveCoilSystem AxialCoilSystem
End Sub

Public Sub optActiveTransverse_Click()
    SetActiveCoilSystem TransverseCoilSystem
End Sub

Private Function PollAFUnit() As Boolean
    Dim finished As Boolean
    Dim PollText As String
    Dim ErrorMessage As String
    
    'Dim delay As Double
    Dim StartTime As Double, totalsecs As Double   'new SWB
    
    'Dim startTime As Double, lag As Double
    Dim status As String
    
    'startTime = Now
    'delay = Timer
    
    Dim Temp1 As Double ' (February 2010 L Carporzen) Monitor temperature of the coils while waiting for Done
    Dim Temp2 As Double
    
    StartTime = Timer   'Timer-starttime is seconds since start of polling
    totalsecs = 0   'use this to count total elapsed time for error msg
    
    'Default finished to stop, otherwise the function won't work
    'properly!
    finished = False
    
    'Default return value - AF run succeeded
    PollAFUnit = True
    
    Do While Not finished
        ' (February 2010 L Carporzen) Monitor temperature of the coils while waiting for Done
        On Error GoTo oops
        
        '(July 2010 - I Hilburn) Had to change to update due to differences in DoDAQIO function
        'from old AnalogInput function
        If EnableT1 Then
        
            'Get the Analog input temp value
            frmDAQ_Comm.DoDAQIO AnalogT1, Temp1
            
            'Convert value to actual temp
            Temp1 = Temp1 * TSlope - Toffset
            
        End If
            
        txtTemp1 = Format$(Temp1, "##0.00")
        
        '(July 2010 - I Hilburn) Had to change to update due to differences in
        'frmDAQ_Comm.DoDAQIO function from old frmMCC.AnalogInput function
        If EnableT2 Then
        
            'Get the Analog input temp value
            frmDAQ_Comm.DoDAQIO AnalogT2, Temp2
            
            'Convert value to actual temp
            Temp2 = Temp2 * TSlope - Toffset
            
        End If
        txtTemp2 = Format$(Temp2, "##0.00")
        If Temp1 > Tmax Or Temp2 > Tmax Then
            frmAF_2G.ZOrder
            frmAF_2G.Show
            lblAFtooHot.Visible = True
            txtTemp1.BackColor = ColorRed
            txtTemp2.BackColor = ColorRed
            ErrorMessage = "The AF degaussing unit is too hot: " & Format$(Temp1, "##0.00") & _
                "°C and " & Format$(Temp2, "##0.00") & "°C." & _
                vbCrLf & "Please check machine."
            PollText = SendStatus
            'SendCommand "DERD",False ' ramp down quickly
            'Flow_Pause
            'SetCodeLevel CodeRed
            frmSendMail.MailNotification "AF Burning", ErrorMessage, CodeRed
            'MsgBox "Panic!!! " & Temp1 & "°C " & Temp2 & "°C"
            ' Panic???
        End If
        lblAFtooHot.Visible = False
        txtTemp1.BackColor = RGB(255, 255, 255)
        txtTemp2.BackColor = RGB(255, 255, 255)
oops:
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
            ErrorMessage = "The AF degaussing unit is experiencing an error on axis " & currentAxis & _
                " at amplitude " & Format(currentCalAmp, 0) & ":" & vbCrLf & vbCrLf & PollText & vbCrLf & _
                vbCrLf & "Execution has been paused and Ramp Down command sent. Please check machine."
            SendCommand "DERD", False
            Flow_Pause
            SetCodeLevel CodeRed
            frmSendMail.MailNotification "AF Error", ErrorMessage, CodeRed
            MsgBox ErrorMessage
            SetCodeLevel StatusCodeColorLevelPrior, True
            
            'Return False - failed AF run
            PollAFUnit = False
            
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
        '        SendCommand "DERD",False
        '        frmSendMail.MailNotification "AF Alert", errormessage, CodeYellow
        '        LockAF False
        '        activecoilsystem = 0
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
        If Timer < StartTime Then StartTime = StartTime - 86400
        '====CODE BELOW CLIPPED OUT BY SWB======================================
        'If Timer < delay Then delay = delay - 86400
        'If Timer - delay > 9 Then
            'Status = SendStatus
            'If ((InStr(Status, "S ?") > 0)) Then
                'errormessage = "The AF degaussing unit reports status unknown." & _
                    vbCrLf & vbCrLf & PollText & vbCrLf & _
                    vbCrLf & "Execution has been paused and Ramp Down command sent. Please check machine."
                'SendCommand "DERD",False
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
        If (Timer - StartTime) > AFWait Then ' (February 2010 L Carporzen) change 90 to 45 for MIT
            totalsecs = totalsecs + (Timer - StartTime)
            'DO WE NEED TO PANIC?
            PollText = SendStatus
            If InStr(PollText, "A    0") > 0 Then
                ' unit has reset! No need to panic, but does reflect a bug with the unit.
                ErrorMessage = "No DONE from AF box for " & Format$(totalsecs, "0.0") & _
                " seconds " & "on axis " & currentAxis & " at amplitude " & Format(currentCalAmp, 0) & "." & _
                " Target amplitude reported as zero, so unit appears to have reset. Execution will continue. " & PollText
                If DEBUG_MODE Then
                   frmDebug.msg "From PollAF: " & ErrorMessage
                End If
                frmSendMail.MailNotification "AF Alert", ErrorMessage, CodeYellow
                LockAF False
                ActiveCoilSystem = 0
                currentDelay = -1
                currentRampRate = -1
                currentUncalAmp = -1
                currentCalAmp = -1
                currentAxis = vbNullString
                
                'Return False, AF ramp has failed
                PollAFUnit = False
                
               finished = True  'so that we exit PollAf
            ElseIf InStr(PollText, "S Z") > 0 Then
               'NO!
                ErrorMessage = "No DONE from AF box for " & Format$(totalsecs, "0.0") & " seconds." & _
                "on axis " & currentAxis & " at amplitude " & Format(currentCalAmp, 0) & "." & _
                " But, AF status=zero. Execution will continue. " & PollText
               If DEBUG_MODE Then
                   frmDebug.msg "From PollAF: " & ErrorMessage
               End If
               frmSendMail.MailNotification "AF Alert", ErrorMessage, CodeYellow
               
               'Return False, AF ramp has failed
               PollAFUnit = False
               
               finished = True  'so that we exit PollAf
            Else
                'YES!  CALL 911
                ErrorMessage = "The AF degaussing coil has not responded for " & Format$(totalsecs, "0.0") & " seconds" & _
                "on axis " & currentAxis & " at amplitude " & Format(currentCalAmp, 0) & "." & vbCrLf & vbCrLf & _
                vbCrLf & "Execution has been paused and Ramp Down command sent. Please check machine. " & PollText
                SendCommand "DERD", False
                Flow_Pause
                SetCodeLevel CodeRed
                frmSendMail.MailNotification "AF Error", ErrorMessage, CodeRed
                MsgBox ErrorMessage
                SetCodeLevel CodeGreen, True
                'reset clock, send error msg every 90 secs if really stuck
                StartTime = Timer
                'NOTE: will exit loop by getting Z response to DERD.  Loop till then
                
                'AF run has failed, Return false
                PollAFUnit = False
                
            End If
        End If
    Loop
    'Status = SendStatus
    'WriteAF PollText, "PollText"
    'WriteAF Status, "Status"
End Function

'(February 2011, I Hilburn)
'Added additional optional parameter to subroutine (allowCommReset)
'Sub Send Command
'   Summary:    Takes in a specially formatted string command to send character by
'               character over RS232 connection to the 2G AF box.  If the first attempt
'               to send the command generates an error, the code resets the connection to the
'               2G Box and trys again once more.  After that, an message box is raised to
'               tell the user that the 2G comm port comm is not working.
'   Parameters:
' outstring  -  Specially formated ASCII string containing 2G commands for the AF
'               demagnetizer box.  For info on different commands, please read AF Demagnetizer
'               manual or go to the applied physics website for documentation:
'               http://www.appliedphysics.com/
' allowCommReset -  Boolean status flag.
'                   True = allow code to reset comm to 2G box and
'                   try sending ASCII command again
'                   False = do not allow reset, proceed directly to error
'
' ISSUES / CONCERNS -   the comm reset option may allow the 2G to hang at dangerous
'                       times - ie. while tracking a Ramp cycle.  This could lead to
'                       damage to the AF coils.
'
'                       Therefore, the specific calls sending the Ramp Execute commands:
'                       "DERC","DERD",and "DERU" should always have the allowCommReset
'                       Parameter set to False.  The SendCommand function can only be
'                       call from within this form frmAF_2G.
'
Private Sub SendCommand(outstring As String, _
                        Optional ByVal allowCommReset As Boolean = True)
    Dim i As Integer
    Dim outchar As String
    If Not EnableAF Then Exit Sub
    
    '(Feb 2011, I Hilburn)
    'New private subroutine now used to set the status in panel three
    'to add the AF command to the status panel if the panel is not already vbnullstring
    UpdatePanel outstring, 3
      
    If MSCommAF.PortOpen = True Then
        MSCommAF.RTSEnable = True
        MSCommAF.OutBufferCount = 0
        MSCommAF.InBufferCount = 0
        ' Because AF unit is stupid, we send out one character
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
        If DEBUG_MODE Then frmDebug.msg "COM " & str$(MSCommAF.CommPort) & " out: " & outstring
    Else
        If Not NOCOMM_MODE Then
            
            '(February 2011, I Hilburn)
            'Bug Fix - AF Comm port is busy on the first attempt to connect to the
            '2G demag box using v2.4.0.  The problem is fixed by reseting the comm
            '(disconnect serial comm to 2G Box, wait for box to process the command,
            ' then reconnect and try sending the original command again)
            'The allowCommReset parameter prevents this reset from being used more than once
            '(otherwise the code could be stuck in an infinite loop if a user entered in the
            ' wrong comm port number in the settings for the 2G box.)
            If allowCommReset = True Then
                
                'Disconnect comm
                Disconnect
                
                'Pause three seconds
                DelayTime 3
                
                'Reconnect AF comm
                Connect
                
                'Recursively call SendCommand
                SendCommand outstring, False
                                
            Else
                    
                'We've already tried disconnecting and reconnecting from the 2G box
                'send error message now
                MsgBox "AF Comm Port Not Open"
                
            End If
            
        End If
            
    End If
    
    If frmProgram.sbStatusBar.Panels(3).text = outstring Then
        
        frmProgram.StatusBar vbNullString, 3
        
    End If
    
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
    alreadyLocked = CoilsLocked
    If Not alreadyLocked Then LockAF True
    SendCommand "DSS"
    SendStatus = GetResponse
    If Not alreadyLocked Then LockAF False
End Function

Public Sub SetActiveCoilSystem(newactivecoilsystem As Integer)
    If CoilsLocked Then
        SetForm
        Exit Sub
    End If
    If newactivecoilsystem = AxialCoilSystem Then
        ActiveCoilSystem = newactivecoilsystem
        optActiveAxial.value = True
        ConfigureCoil AfAxialCoord
    End If
    If newactivecoilsystem = TransverseCoilSystem Then
        ActiveCoilSystem = newactivecoilsystem
        optActiveTransverse.value = True
        ConfigureCoil AfTransCoord
    End If
End Sub

Private Sub SetAmplitude(ByVal AFLevel As Double)
    If CoilsLocked Then
        SetForm
        Exit Sub
    End If
        
    If EnableAF = False Then Exit Sub
    
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

'(February 2011, I Hilburn)
'
'Sub Update3rdPanel
'
' Summary:  Takes in the command string that will be sent to the 2G AF box
'           and, if the panel contains text already, don't add the command
Private Sub UpdatePanel(ByVal PanelStr As String, _
                        ByVal PanelIndex As Integer)

    Dim TempStr As String

    'We can't update the text in nonexistent panels
    If PanelIndex > frmProgram.sbStatusBar.Panels.Count Then Exit Sub

    'Locally store the text currently in the 3rd panel
    TempStr = frmProgram.sbStatusBar.Panels(PanelIndex).text

    'Check to see if the 3rd panel is empty
    If TempStr = vbNullString Then
    
        frmProgram.StatusBar PanelStr, PanelIndex
        
    End If

End Sub

Private Sub WriteAF(txt As String, Label As String)
    ' Subroutine added by L Carporzen (August 2007) to record the communications with the 2G degausser.
    Dim filenum As Integer
    Dim filename As String
    filenum = FreeFile
    filename = Prog_DefaultPath & "\AFsequence.txt"
    On Error GoTo oops
    Open filename For Append As #filenum
    Print #filenum, txt; ","; Label
    Close #filenum
    GoTo stillworking
oops:
    MsgBox "Unable to write to " & filename & "!"
stillworking:
End Sub

