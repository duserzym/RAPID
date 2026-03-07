VERSION 5.00
Begin VB.Form frmIRMARM 
   Caption         =   "ARM/IRM Controller"
   ClientHeight    =   5265
   ClientLeft      =   5880
   ClientTop       =   2145
   ClientWidth     =   7905
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5265
   ScaleWidth      =   7905
   Begin VB.Frame Frame3 
      Caption         =   "Coil Temperatures"
      Height          =   3375
      Left            =   120
      TabIndex        =   40
      Top             =   1800
      Width           =   2055
      Begin VB.TextBox txtTemp1 
         Height          =   285
         Left            =   360
         TabIndex        =   43
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox txtTemp2 
         Height          =   285
         Left            =   360
         TabIndex        =   42
         Top             =   1320
         Width           =   735
      End
      Begin VB.CommandButton cmdTemp 
         Caption         =   "Refresh T"
         Height          =   375
         Left            =   360
         TabIndex        =   41
         Top             =   2760
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Axial Coil:"
         Height          =   255
         Left            =   360
         TabIndex        =   48
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "Transver Coil:"
         Height          =   255
         Left            =   360
         TabIndex        =   47
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "°C"
         Height          =   255
         Left            =   1200
         TabIndex        =   46
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label8 
         Caption         =   "°C"
         Height          =   255
         Left            =   1200
         TabIndex        =   45
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label lblAFtooHot 
         Caption         =   "The AF unit is too hot so let's pause a little bit..."
         Height          =   615
         Left            =   360
         TabIndex        =   44
         Top             =   1920
         Visible         =   0   'False
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "IRM Calibration"
      Height          =   1695
      Left            =   2280
      TabIndex        =   34
      Top             =   3480
      Width           =   2775
      Begin VB.CommandButton cmdChangeIRMSettings 
         Caption         =   "Change IRM Settings..."
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
         Left            =   120
         TabIndex        =   39
         Top             =   1200
         Width           =   2535
      End
      Begin VB.CommandButton cmdCalibrateVoltages 
         Caption         =   "Calibrate IRM DAQ Volts"
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
         Left            =   120
         TabIndex        =   36
         Top             =   720
         Width           =   2535
      End
      Begin VB.CommandButton cmdCalibrateFields 
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
         Height          =   375
         Left            =   120
         TabIndex        =   35
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.Timer tmrARMWatch 
      Interval        =   10000
      Left            =   120
      Top             =   5160
   End
   Begin VB.Frame frameIRM 
      Caption         =   "IRM"
      Height          =   3255
      Left            =   2280
      TabIndex        =   4
      Top             =   120
      Width           =   2775
      Begin VB.OptionButton optCoil 
         Caption         =   "Transverse"
         Height          =   375
         Index           =   1
         Left            =   480
         TabIndex        =   38
         Top             =   1440
         Width           =   1215
      End
      Begin VB.OptionButton optCoil 
         Caption         =   "Axial"
         Height          =   375
         Index           =   0
         Left            =   480
         TabIndex        =   37
         Top             =   1080
         Width           =   975
      End
      Begin VB.CommandButton cmdInterruptCharge 
         Caption         =   "Interrupt"
         Height          =   375
         Left            =   1800
         TabIndex        =   32
         Top             =   1200
         Width           =   735
      End
      Begin VB.CommandButton cmdIRMAverageVoltageIn 
         Caption         =   "Read IRM Voltage:"
         Height          =   252
         Left            =   120
         TabIndex        =   31
         Top             =   2760
         Width           =   1575
      End
      Begin VB.TextBox txtIRMAverageVoltageIn 
         Height          =   288
         Left            =   1800
         TabIndex        =   30
         Top             =   2760
         Width           =   735
      End
      Begin VB.CheckBox chkBackfield 
         Caption         =   "Negative Polarity"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   2280
         Width           =   1815
      End
      Begin VB.CheckBox chkLockCoils 
         Caption         =   "Lock coil selection"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1800
         Width           =   1695
      End
      Begin VB.TextBox txtPulseField 
         Height          =   288
         Left            =   1200
         TabIndex        =   9
         Top             =   720
         Width           =   492
      End
      Begin VB.CommandButton cmdIRMFirebyGauss 
         Caption         =   "Fire"
         Height          =   372
         Left            =   1800
         TabIndex        =   8
         Top             =   720
         Width           =   732
      End
      Begin VB.CommandButton cmdIRMFire 
         Caption         =   "Fire"
         Height          =   372
         Left            =   1800
         TabIndex        =   7
         Top             =   240
         Width           =   732
      End
      Begin VB.TextBox txtPulseVolts 
         Height          =   288
         Left            =   1200
         TabIndex        =   5
         Top             =   240
         Width           =   492
      End
      Begin VB.Label lblIRMStatus 
         Alignment       =   2  'Center
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   2520
         Width           =   2415
      End
      Begin VB.Label Label2 
         Caption         =   "Peak field (G):"
         Height          =   252
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   1332
      End
      Begin VB.Label Label1 
         Caption         =   "Voltage:"
         Height          =   252
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1332
      End
   End
   Begin VB.Frame frameARM 
      Caption         =   "ARM"
      Height          =   1575
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   2055
      Begin VB.TextBox txtBiasField 
         Height          =   288
         Left            =   1320
         TabIndex        =   12
         Top             =   480
         Width           =   492
      End
      Begin VB.CommandButton cmdSetBiasField 
         Caption         =   "Set"
         Height          =   372
         Left            =   360
         TabIndex        =   11
         Top             =   960
         Width           =   732
      End
      Begin VB.Label Label3 
         Caption         =   "Bias field (G):"
         Height          =   252
         Left            =   120
         TabIndex        =   13
         Top             =   480
         Width           =   1332
      End
   End
   Begin VB.Frame frameMCC 
      Caption         =   "IRM/ARM DAQ Controller"
      Height          =   4575
      Left            =   5160
      TabIndex        =   1
      Top             =   120
      Width           =   2655
      Begin VB.TextBox txtMCCARMSet 
         Height          =   288
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   33
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox txtMCCIRMPowerAmpVoltageIn 
         BackColor       =   &H00FFFFFF&
         Height          =   288
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   4080
         Width           =   855
      End
      Begin VB.CommandButton cmdMCCIRMPowerAmpVoltageIn 
         Caption         =   "Read IRM Power Amp V in:"
         Height          =   495
         Left            =   120
         TabIndex        =   28
         Top             =   3960
         Width           =   1335
      End
      Begin VB.TextBox txtMCCIRMTrim 
         Height          =   288
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   3480
         Width           =   855
      End
      Begin VB.CommandButton cmdMCCIRMTrim 
         Caption         =   "IRM Trim:"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   3480
         Width           =   975
      End
      Begin VB.TextBox txtMCCIRMFire 
         Height          =   288
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   3000
         Width           =   855
      End
      Begin VB.CommandButton cmdMCCIRMFire 
         Caption         =   "IRM Fire:"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   3000
         Width           =   975
      End
      Begin VB.CommandButton cmdMCCARMSet 
         Caption         =   "ARM Set:"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox txtMCCIRMVin 
         Height          =   288
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   2520
         Width           =   855
      End
      Begin VB.CommandButton cmdMCCIRMVin 
         Caption         =   "Read IRM V in:"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   2520
         Width           =   1335
      End
      Begin VB.TextBox txtMCCIRMVout 
         Height          =   288
         Left            =   1560
         TabIndex        =   20
         Top             =   2040
         Width           =   855
      End
      Begin VB.TextBox txtMCCARMVout 
         Height          =   288
         Left            =   1560
         TabIndex        =   19
         Top             =   840
         Width           =   855
      End
      Begin VB.CommandButton cmdMCCIRMVout 
         Caption         =   "IRM V:"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   2040
         Width           =   735
      End
      Begin VB.CommandButton cmdMCCARMVout 
         Caption         =   "ARM V:"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   840
         Width           =   735
      End
      Begin VB.CommandButton cmdShowMCC 
         Caption         =   "Show"
         Height          =   255
         Left            =   960
         TabIndex        =   2
         Top             =   360
         Width           =   1095
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   2520
         Y1              =   1800
         Y2              =   1800
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Default         =   -1  'True
      Height          =   372
      Left            =   5160
      TabIndex        =   0
      Top             =   4800
      Width           =   2655
   End
End
Attribute VB_Name = "frmIRMARM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' This is for handling directives to DAQ Boards
' for control of IRM/ARM
Option Explicit

Dim isUserChange As Boolean
Dim IRMPastVolts(50) As Double
Public IRMPeakVoltage As Double

'Instantiate an instance of IRMData
Private irm_pulse_data As IRMData
Private stop_charging_change_threshold_in_percent As Double
Private stop_trimming_change_threshold_in_percent As Double

'-----------------------------------------------------------------------------------------------------------------------------'
'-----------------------------------------------------------------------------------------------------------------------------'
'
'       HUGE Code Modification
'       July, 30 2010
'       Isaac Hilburn
'
'   Stripped out all code related to the IRM Hi-field and replaced the IRM hi/lo field with IRM Axial / Transverse
'   Added Axial / Transverse radio buttons to frmIRMARM to allow switching between the two coils
'   for IRM.  Note, these radio buttons will do the inverse of the same radio buttons on other forms
'   (Axial IRM = "Down" position, whereas Axial AF = "Up" position for the Axial relay)
'-----------------------------------------------------------------------------------------------------------------------------'
'-----------------------------------------------------------------------------------------------------------------------------'

'-----------------------------------------------------------------------------------------------------------------------------'
'-----------------------------------------------------------------------------------------------------------------------------'
'
'       Major Code Modification
'       June 2010
'       Isaac Hilburn
'
'   Changed Analog & Digital input output to use the new form frmDAQ_Comm instead of frmMCC.  Replaced the calls:
'   frmMCC.AnalogInput, frmDAQ_Comm.DoDAQIO, frmMCC.DigitialInput, and frmDAQ_Comm.Dodaqio with the call:
'   frmDAQ_Comm.DoDAQIO - which handles all A/D, D/A and DIO functions.
'   The frmDAQComm functions support the new Channel / Board object implementation of the DAQ comm - in which the
'   IO action to perform is contained within the Channel Object, itself.
'
'   Form DAQ comm only supports
'   both the Measurement Computing and the ADWIN boards and can be used with Measurement Computing boards other than
'   the PCI-DAS6030 board.  (Different MCC boards have different Digital I/O configurations, and the old code
'   only supports the configuration present in the PCI-DAS6030 board.)
'
'   Also changed the display for the digital I/O controls (to "Off"/"On" or "Ready" / "NOT Ready"), with light green and
'   light red shading to emphasize the boolean state.
'
'-----------------------------------------------------------------------------------------------------------------------------'
'
'Old Form Comments:
'
' (March 2008 L Carporzen) Put in Settings the IRM/ARM channels (MIT acquisition board does not work on IRMTrim = 3
' Analog channel output
'Const ARMVoltageOut As Integer = 0
'Const IRMVoltageOut As Integer = 1
' Analog input
'Const IRMCapacitorVoltageIn As Integer = 0
' DIO line assignments
'Const ARMSet As Integer = 0
'Const IRMFire As Integer = 1
'Const IRMTrim As Integer = 7 '3 (August 2007 L Carporzen) Pin changed after Pin 3 fried on the acquisition box/card
'Const IRMPowerAmpVoltageIn As Integer = 4
'
'-----------------------------------------------------------------------------------------------------------------------------'
'-----------------------------------------------------------------------------------------------------------------------------'

Dim ARMStartTime As Date
Dim CurrentBiasField As Double
Dim IRMBackfieldMode As Boolean
Dim IRMInterrupt As Boolean

Private Function AtEndOfChargeCycle() As Boolean

    AtEndOfChargeCycle = False
    
    

    'Compare the average change over the past window to the average change over the entire ramp.
    If irm_pulse_data.average_change_over_window <> 0 And _
       irm_pulse_data.average_change_entire_charging_cycle <> 0 Then
       
       Dim change_ratio As Double
       change_ratio = Abs(irm_pulse_data.average_change_over_window / irm_pulse_data.average_change_entire_charging_cycle)
       
       If irm_pulse_data.average_change_over_window > 0.005 Then
       
            AtEndOfChargeCycle = False
            Exit Function
            
       End If
       
       If change_ratio * 100 < stop_charging_change_threshold_in_percent Then AtEndOfChargeCycle = True
       
    End If

End Function

Private Function AtEndOfTrimCycle() As Boolean

    AtEndOfTrimCycle = False
    
    If irm_pulse_data.average_change_over_window <> 0 And _
       irm_pulse_data.average_change_entire_charging_cycle <> 0 Then
       
       Dim change_ratio As Double
       change_ratio = Abs(irm_pulse_data.average_change_over_window / irm_pulse_data.average_change_entire_charging_cycle)
       
       'Stop on trim rate slowing
       If change_ratio * 100 < stop_trimming_change_threshold_in_percent Then AtEndOfTrimCycle = True
       
       'Stop on no trim happening at all
       If Abs(1 - change_ratio) * 100 < stop_trimming_change_threshold_in_percent Then AtEndOfTrimCycle = True
       
    End If

End Function

Public Function CheckIRMPast(ByRef PastArray() As Double, _
                             Optional ByVal TargetVoltage As Double = -1) As Boolean

    Dim AvgSlope As Double
    Dim AvgVal As Double
    Dim StdDevSlope As Double
    Dim N As Long
    Dim i As Long
    Dim SlopeTol As Double
            
    'Start the average slope out at 0
    AvgSlope = 0
    AvgVal = 0
    
    N = UBound(PastArray)
    
    'Get the sum of the slopes
    For i = 0 To N - 2
    
        AvgSlope = AvgSlope + (PastArray(i + 1) - PastArray(i))
        AvgVal = AvgVal + PastArray(i)
        
    Next i
    
    AvgVal = AvgVal + PastArray(N - 1)
    
    'Change the sum into an average
    AvgSlope = AvgSlope / (N - 1)
    AvgVal = AvgVal / N
    
'    Debug.Print "Average Slope = " & Trim(Str(AvgSlope))
    
    'If values are too low, may get a false plateau.
    'Also, definition of "too low" is different for high-voltage ramps
    'than it is for low-voltage ramps
    If TargetVoltage = -1 And AvgVal < 1 Then
    
        CheckIRMPast = False
        
        Exit Function
        
    ElseIf TargetVoltage <> -1 And AvgVal < TargetVoltage / 10 Then
    
        CheckIRMPast = False
        
        Exit Function
        
    End If

'****************************************************************************'
'  (Feb 2011, I Hilburn)
'   Commenting all of this code out.
'
'   Reason: This code leads to a false identification of a plateauing IRM capacitor
'           voltage at low voltages.  Also, it's unnecessarily complicated.
'
'   Replacement code:   Just use average slope
'    'Start the StdDevSlope out at 0
'    StdDevSlope = 0
'
'---------------------------------------------------------------------------'
'
'    'Get the sum of the variances of the slope
'    For i = 0 To N - 2
'
'        StdDevSlope = StdDevSlope + (AvgSlope - (PastArray(i + 1) - PastArray(i))) ^ 2
'
'    Next i
'
'    StdDevSlope = Sqr(StdDevSlope / (N - 1))
'
''    If StdDevSlope > 0.5 Then StdDevSlope = 0.5
'
'    If TargetVoltage / 20 < StdDevSlope Then StdDevSlope = TargetVoltage / 20
'
'    'Now, if the latest value is within 1 standard deviation of the slope
'    'of the average slope is a negative with an amplitude greater than 1 standard deviation
'    'of the slope, then the IRM has reached a plateau
'
'------------------------------------------------------------------------------'
    
    'Tolerance for the platuea = 1/100000 of the target voltage
    SlopeTol = 0.00001 * TargetVoltage
    
    'If the Tolerance for the slope change < 0.0005 then raise it to 0.0005
    If SlopeTol < 0.0005 Then SlopeTol = 0.0005
    
    If Abs(AvgSlope) < SlopeTol Then
    
        'We've hit a plateau!
        CheckIRMPast = True
        
'******************************************************************************'
    
'(September 27, 2010 - I Hilburn)
'This elseif statement may be causing problems with the high-voltage IRMs
'Identifying a false plateau during the early ramp up
'    ElseIf AvgSlope < 0 And Abs(AvgSlope) > StdDevSlope Then
'
'        'The slope is significantly negative
'        CheckIRMPast = True
        
    Else
    
        'We're not in a plateau
        CheckIRMPast = False
        
    End If
        
End Function

Private Sub chkBackfield_Click()
    IRMBackfieldMode = chkBackfield.value
End Sub

Private Sub chkLockCoils_Click()

    If Me.chkLockCoils.value = Checked Then
    
        CoilsLocked = True
        optCoil(0).Enabled = False
        optCoil(1).Enabled = False
        
    ElseIf Me.chkLockCoils.value = Unchecked Then
    
        CoilsLocked = False
        optCoil(0).Enabled = True
        optCoil(1).Enabled = True
        
    End If

End Sub

Private Sub cmdCalibrateFields_Click()

    'Set IRM mode, Load, and show the AF Calibration form
    frmCalibrateCoils.InAFMode = False
    Load frmCalibrateCoils
    frmCalibrateCoils.Show
    frmCalibrateCoils.ZOrder 0

End Sub

Private Sub cmdCalibrateVoltages_Click()

    'Load, and show the AF Calibration form
    Load frmIRM_VoltageCalibration
    frmIRM_VoltageCalibration.Show
    frmIRM_VoltageCalibration.ZOrder 0

End Sub

Private Sub cmdChangeIRMSettings_Click()

    'Load and show the settings form with the tab set to the IRM page
    Load frmSettings
    frmSettings.selectTab 7
    frmSettings.Show
    
End Sub

Private Sub cmdClose_Click()
    Me.Hide
End Sub

Private Sub cmdInterruptCharge_Click()
    IRMInterruptCharge
End Sub

Private Sub cmdIRMAverageVoltageIn_Click()
    IRMAverageVoltageIn
End Sub

Public Sub cmdIRMFire_Click()
    FireIRM val(txtPulseVolts)
End Sub

Private Sub cmdIRMFirebyGauss_Click()
    FireIRMAtField (val(txtPulseField))
End Sub

Private Sub cmdMCCARMSet_Click()
'(July 2010 - I Hilburn) Replaces old frmMCC with frmDAQ_Comm using the Channel object variables

    If Me.txtMCCARMSet = "Off" Then
    
        'Need to toggle TTL for ARM shut
        frmDAQ_Comm.DoDAQIO ARMSet, , False
        
        Me.txtMCCARMSet = "On"
        Me.txtMCCARMSet.BackColor = QBColor(10)
        
    ElseIf Me.txtMCCARMSet = "On" Then
    
        'Need to toggle TTL for ARM open
        frmDAQ_Comm.DoDAQIO ARMSet, , True
        
        Me.txtMCCARMSet = "Off"
        Me.txtMCCARMSet.BackColor = QBColor(12)
        
    End If
        
End Sub

Private Sub cmdMCCARMVout_Click()
'(July 2010 - I Hilburn) Replaces old frmMCC with frmDAQ_Comm using the Channel object variables
    frmDAQ_Comm.DoDAQIO ARMVoltageOut, val(txtMCCARMVout)
End Sub

Private Sub cmdMCCIRMFire_Click()
'(July 2010 - I Hilburn) Replaces old frmMCC with frmDAQ_Comm using the Channel object variables

    If Me.txtMCCIRMFire = "Off" Then
    
        'Need to toggle TTL for ARM shut
        frmDAQ_Comm.DoDAQIO IRMFire, , False
        
        Me.txtMCCIRMFire = "On!"
        Me.txtMCCIRMFire.BackColor = QBColor(10)
        
    ElseIf Me.txtMCCIRMFire = "On!" Then
    
        'Need to toggle TTL for ARM open
        frmDAQ_Comm.DoDAQIO IRMFire, , True
        
        Me.txtMCCIRMFire = "Off"
        Me.txtMCCIRMFire.BackColor = QBColor(12)
        
    End If

End Sub

Private Sub cmdMCCIRMPowerAmpVoltageIn_Click()

    Dim TempD As Double

'(July 2010 - I Hilburn) Replaces old frmMCC with frmDAQ_Comm using the Channel object variables
    TempD = CDbl(frmDAQ_Comm.DoDAQIO(IRMPowerAmpVoltageIn))
    
    Me.txtMCCIRMPowerAmpVoltageIn = Trim(Str(TempD))
    
End Sub

Private Sub cmdMCCIRMTrim_Click()
'(July 2010 - I Hilburn) Replaces old frmMCC with frmDAQ_Comm using the Channel object variables
    
    If txtMCCIRMTrim = "Off" Then
    
        'Need to toggle TTL for ARM shut
        frmDAQ_Comm.DoDAQIO IRMTrim, , TrimOnOff(True)
        
        txtMCCIRMTrim = "On"
        txtMCCIRMTrim.BackColor = QBColor(10)
        
    ElseIf txtMCCIRMTrim = "On" Then
    
        'Need to toggle TTL for ARM open
        frmDAQ_Comm.DoDAQIO IRMTrim, , TrimOnOff(False)
        
        txtMCCIRMTrim = "Off"
        txtMCCIRMTrim.BackColor = QBColor(12)
        
    End If
    
End Sub

Private Sub cmdMCCIRMVin_Click()

    '(July 2010 - I Hilburn) Replaces old frmMCC with frmDAQ_Comm using the Channel object variables
    txtMCCIRMVin = Format(CDbl(frmDAQ_Comm.DoDAQIO(IRMCapacitorVoltageIn)), "#0.0#####")
    
End Sub

Private Sub cmdMCCIRMVout_Click()
'(July 2010 - I Hilburn) Replaces old frmMCC with frmDAQ_Comm using the Channel object variables
    frmDAQ_Comm.DoDAQIO IRMVoltageOut, val(txtMCCIRMVout)
End Sub

Private Sub cmdSetBiasField_Click()
    SetBiasField val(txtBiasField)
End Sub

Private Sub cmdShowMCC_Click()
    frmDAQ_Comm.ZOrder
    frmDAQ_Comm.Show
End Sub

Private Sub cmdTemp_Click()

    '(July 2010 - I Hilburn)
    'Copied from Laurent's code in frmAF_2G. Only Changed textbox object references.

    Dim Temp1 As Double ' (February 2010 L Carporzen) Monitor temperature of the coils before executing AF
    Dim Temp2 As Double
    
    '(July 2010 - I Hilburn) Replaces old frmMCC with frmDAQ_Comm using the Channel object variables
    If EnableT1 Then
    
        Temp1 = frmDAQ_Comm.DoDAQIO(AnalogT1)
        
        Temp1 = Temp1 * TSlope - Toffset
        
    End If
        
    txtTemp1 = Format$(Temp1, "##0.00")
    
    If EnableT2 Then
    
        Temp2 = frmDAQ_Comm.DoDAQIO(AnalogT2)
        
        Temp2 = Temp2 * TSlope - Toffset
        
    End If
        
    txtTemp2 = Format$(Temp2, "##0.00")

End Sub

'Specific function for interpolating the Volts vs Field value table for the
'IRM low-field pulse to convert field values to pulse voltage values
Private Function ConvertGaussToPulseAxialVolts(field As Double) As Double

    Dim i As Integer
    Dim N As Integer
    Dim Slope As Double
    
    'If the field is less than zero, than invert it's sign
    If field < 0 Then field = -field
    
    'Coerce the field value in range with the max and min field values
    If field < PulseAxialMin Then field = PulseAxialMin
    If field > PulseAxialMax Then field = PulseAxialMax
    
    'Set the function return = -1
    ConvertGaussToPulseAxialVolts = -1
    
    'Get the Number of elements in the field calibration table
    N = UBound(PulseAxial, 1)
    
    'Loop through the table, ignoring the field (zero) entry
    'and find the pulse voltage that matches the inputed field value
    For i = 1 To N - 1
        
        If PulseAxial(i, 1) = field Then
        
            ConvertGaussToPulseAxialVolts = PulseAxial(i, 0)
            Exit Function
            
        ElseIf PulseAxial(i - 1, 1) < field And PulseAxial(i, 1) > field Then
        
            'Calculate the local slope (Voltage / Field)
            Slope = (PulseAxial(i, 0) - PulseAxial(i - 1, 0)) / (PulseAxial(i, 1) - PulseAxial(i - 1, 1))
            
            'Use X = Y(0) + Slope * Y to calculate the matching voltage value
            ConvertGaussToPulseAxialVolts = PulseAxial(i - 1, 0) + Slope * (field - PulseAxial(i - 1, 1))
            
            Exit For
            
        ElseIf PulseAxial(PulseAxialCount, 1) < field Then
        
            'Need to extrapolate out slope
            'Calculate the local slope (Voltage / Field)
            Slope = (PulseAxial(PulseAxialCount, 0) - PulseAxial(PulseAxialCount - 1, 0)) / _
                    (PulseAxial(PulseAxialCount, 1) - PulseAxial(PulseAxialCount - 1, 1))
            
            ConvertGaussToPulseAxialVolts = PulseAxial(PulseAxialCount, 0) + _
                                            Slope * (field - PulseAxial(PulseAxialCount, 1))
            
            Exit For
            
        End If
        
    Next i
    
    'ConvertGaussToPulseAxialVolts = (field - PulseAxialY) / PulseAxialSlope
End Function

'Specific function for interpolating the Volts vs Field value table for the
'IRM high-field pulse to convert field values to pulse voltage values
Private Function ConvertGaussToPulseTransVolts(field As Double) As Double

    Dim i As Integer
    Dim N As Integer
    Dim Slope As Double
    
    'If the field is less than zero, than invert it's sign
    If field < 0 Then field = -field
    
    'Coerce the field value in range with the max and min field values
    If field < PulseTransMin Then field = PulseTransMin
    If field > PulseTransMax Then field = PulseTransMax
    
    'Set the function return = -1
    ConvertGaussToPulseTransVolts = -1
    
    'Get the Number of elements in the field calibration table
    N = UBound(PulseTrans, 1)
    
    'Loop through the table, ignoring the field (zero) entry
    'and find the pulse voltage that matches the inputed field value
    For i = 1 To N - 1
        
        If PulseTrans(i, 1) = field Then
        
            ConvertGaussToPulseTransVolts = PulseTrans(i, 0)
            Exit Function
            
        ElseIf PulseTrans(i - 1, 1) < field And PulseTrans(i, 1) > field Then
        
            'Calculate the local slope (Voltage / Field)
            Slope = (PulseTrans(i, 0) - PulseTrans(i - 1, 0)) / (PulseTrans(i, 1) - PulseTrans(i - 1, 1))
            
            'Use X = Y(0) + Slope * Y to calculate the matching voltage value
            ConvertGaussToPulseTransVolts = PulseTrans(i - 1, 0) + Slope * (field - PulseTrans(i - 1, 1))
            
            Exit For
            
        ElseIf PulseTrans(PulseTransCount, 1) < field Then
        
            'Need to extrapolate out slope
            'Calculate the local slope (Voltage / Field)
            Slope = (PulseTrans(PulseTransCount, 0) - PulseTrans(PulseTransCount - 1, 0)) / _
                    (PulseTrans(PulseTransCount, 1) - PulseTrans(PulseTransCount - 1, 1))
            
            ConvertGaussToPulseTransVolts = PulseTrans(PulseTransCount, 0) + _
                                            Slope * (field - PulseTrans(PulseTransCount, 1))
            
            Exit For
            
        End If
        
    Next i
    
    'ConvertGaussToPulseTransVolts = (field - PulseTransY) / PulseTransSlope
End Function

'Convert a field value in gauss to a low or high field IRM Pulse
'using the IRM field calibration table
Private Function ConvertGaussToPulseVolts(field As Double) As Double
    
    'If the field is negative, flip the sign on it
    If field < 0 Then field = -field
    
    If optCoil(1).value = True Then
    
        ConvertGaussToPulseVolts = ConvertGaussToPulseTransVolts(field)
        
    Else
    
        ConvertGaussToPulseVolts = ConvertGaussToPulseAxialVolts(field)
        
    End If
    
End Function

Public Function ConvertMCCVoltsToPulseVolts(Volts As Double) As Double
    ConvertMCCVoltsToPulseVolts = Volts / modConfig.PulseReturnMCCVoltConversion
End Function

'Specific function for interpolating the Volts vs Field value table for the
'IRM low-field pulse to convert pulse volts to field values
Private Function ConvertPulseAxialVoltsToGauss(Volts As Double) As Double
    
    Dim i As Integer
    Dim N As Integer
    Dim Slope As Double
    
    'ConvertPulseAxialVoltsToGauss = PulseAxialY + PulseAxialSlope * VOLTS
    ConvertPulseAxialVoltsToGauss = 0
    
    If Volts < 0 Then Volts = -Volts
    
    If Volts = 0 Then
    
        ConvertPulseAxialVoltsToGauss = 0
        Exit Function
        
    End If
        
    'Get the number of calibration values in the IRM low-field pulse
    'calibration table
    N = UBound(PulseAxial, 1)
    
    'Loop through the values in the IRM Low-field pulse calibration table
    'and interpolate matching field value for inputed pulse voltage
    For i = 1 To N - 1
    
        If PulseAxial(i, 0) = Volts Then
        
            ConvertPulseAxialVoltsToGauss = PulseAxial(i, 1)
            
            Exit Function
            
        ElseIf PulseAxial(i - 1, 0) < Volts And PulseAxial(i, 0) > Volts Then
        
            'Calculate local slope of the IRM low-field calibration table
            '(Field / Voltage)
            Slope = (PulseAxial(i, 1) - PulseAxial(i - 1, 1)) / (PulseAxial(i, 0) - PulseAxial(i - 1, 0))
            
            'Use Y = X(0) + Slope * X to get the field value
            'Where Y = field value, and X = pulse value
            ConvertPulseAxialVoltsToGauss = PulseAxial(i - 1, 1) + Slope * (Volts - PulseAxial(i - 1, 0))
            
            Exit For
            
        ElseIf PulseAxial(PulseAxialCount, 0) < Volts Then
        
            'Calculate local slope of the IRM low-field calibration table
            '(Field / Voltage)
            Slope = (PulseAxial(PulseAxialCount, 1) - PulseAxial(PulseAxialCount - 1, 1)) / _
                    (PulseAxial(PulseAxialCount, 0) - PulseAxial(PulseAxialCount - 1, 0))
            
            'Extrapolate past voltage in (i,0)
            ConvertPulseAxialVoltsToGauss = PulseAxial(PulseAxialCount, 1) + _
                                            Slope * (Volts - PulseAxial(PulseAxialCount, 0))
            
            Exit For
            
        End If
        
    Next i
    
End Function

'Specific function for interpolating the Volts vs Field value table for the
'IRM high-field pulse to convert pulse volts to field values
Private Function ConvertPulseTransVoltsToGauss(Volts As Double) As Double
    
    Dim i As Integer
    Dim N As Integer
    Dim Slope As Double
    
    'ConvertPulseTransVoltsToGauss = PulseTransY + PulseTransSlope * VOLTS
    ConvertPulseTransVoltsToGauss = 0
    
    If Volts < 0 Then Volts = -Volts
    
    'Get the number of calibration values in the IRM high-field pulse
    'calibration table
    N = UBound(PulseTrans, 1)
    
    'Loop through the values in the IRM high-field pulse calibration table
    'and interpolate matching field value for inputed pulse voltage
    For i = 1 To N - 1
    
        If PulseTrans(i, 0) = Volts Then
        
            ConvertPulseTransVoltsToGauss = PulseTrans(i, 1)
            
            Exit Function
            
        ElseIf PulseTrans(i - 1, 0) < Volts And PulseTrans(i, 0) > Volts Then
        
            'Calculate local slope of the IRM low-field calibration table
            '(Field / Voltage)
            Slope = (PulseTrans(i, 1) - PulseTrans(i - 1, 1)) / (PulseTrans(i, 0) - PulseTrans(i - 1, 0))
            
            'Use Y = X(0) + Slope * X to get the field value
            'Where Y = field value, and X = pulse value
            ConvertPulseTransVoltsToGauss = PulseTrans(i - 1, 0) + Slope * (Volts - PulseTrans(i - 1, 0))
            
            Exit For
            
        ElseIf PulseTrans(PulseTransCount, 0) < Volts Then
        
            'Calculate local slope of the IRM low-field calibration table
            '(Field / Voltage)
            Slope = (PulseTrans(PulseTransCount, 1) - PulseTrans(PulseTransCount - 1, 1)) / _
                    (PulseTrans(PulseTransCount, 0) - PulseTrans(PulseTransCount - 1, 0))
            
            'Extrapolate past voltage in (i,0)
            ConvertPulseTransVoltsToGauss = PulseTrans(PulseTransCount, 1) + _
                                            Slope * (Volts - PulseTrans(PulseTransCount, 0))
            
            Exit For
            
        End If
        
    Next i
    
End Function

'Convert a IRM axial or transverse pulse voltage to a field value in gauss
'using the IRM field calibration table
Private Function ConvertPulseVoltsToGauss(Volts As Double)
    
    If Me.optCoil(1).value = True Then
    
        ConvertPulseVoltsToGauss = ConvertPulseTransVoltsToGauss(Volts)
        
    ElseIf Me.optCoil(0).value = True Then
    
        ConvertPulseVoltsToGauss = ConvertPulseAxialVoltsToGauss(Volts)
        
    End If
    
End Function

Public Function ConvertPulseVoltsToMCCVolts(Volts As Double) As Double
        ConvertPulseVoltsToMCCVolts = Volts * PulseMCCVoltConversion
End Function

Public Sub FireASC_IrmAtPulseVolts(ByVal pulse_volts As Double, _
                                   Optional ByVal CalibrationMode As Boolean = False)

    Dim TempD As Double
    Dim TempS As String
    Dim TempL As Long
    Dim Temp1 As Double
    Dim Temp2 As Double
    Dim TWarning As Boolean
    Dim ErrorMessage As String
        
    On Error GoTo oops:
        
    If pulse_volts > 0 Then
             
         'Before doing anything with the ADWIN board, get the AF coil temperatures
         '(July 2010 - I Hilburn) Replaces old frmMCC with frmDAQ_Comm using the Channel object variables
         If EnableT1 Then
         
             Temp1 = frmDAQ_Comm.DoDAQIO(AnalogT1)
             
             Temp1 = Temp1 * TSlope - Toffset
             
         End If
             
         'Update display on the IRM form
         txtTemp1 = Format$(Temp1, "##0.00")
             
         If EnableT2 Then
         
             Temp2 = frmDAQ_Comm.DoDAQIO(AnalogT2)
             
             Temp2 = Temp2 * TSlope - Toffset
         
         End If
         
         'Update display on the IRM Form
         txtTemp2 = Format$(Temp2, "##0.00")
         
         'Check Temperature to see if it is not zeroed (gone within 20 deg of -1 * Toffset)
         If Not ValidSensorTemp(Temp1, Temp2) Then
         
             'Start code to tell user that the temp sensor values are bad
             NotifySensorError Temp1, Temp2
             
         End If
             
         
         If EnableT1 Or EnableT2 Then
             
             Do While Temp1 >= Thot Or Temp2 >= Thot
                 
                 lblAFtooHot.Visible = True
                 txtTemp1.BackColor = ColorOrange
                 txtTemp2.BackColor = ColorOrange
                 
                 ErrorMessage = "The IRM Coil(s) temperature is above " & Thot & "°C: " & Format$(Temp1, "##0.00") & _
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
                     
                         Temp1 = frmDAQ_Comm.DoDAQIO(AnalogT1)
                         
                         Temp1 = Temp1 * TSlope - Toffset
                         
                     End If
                     
                     txtTemp1 = Format$(Temp1, "##0.00")
                     
                     If EnableT2 Then
                     
                         Temp2 = frmDAQ_Comm.DoDAQIO(AnalogT2)
                         
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
        
    End If
        
    On Error GoTo 0
   
oops:
        
    If Prog_paused = True Then
    
        'Set TempS equal to the current contents
        'of the 2nd status bar panel
        TempS = frmProgram.sbStatusBar.Panels(2).text
        
        'Update the program form status bar
        frmProgram.StatusBar "Paused...", 2
        
        'Loop and wait for status to change and DoEvents
        'to allow user to make changes
        Do
        
            'Pause 100 ms
            PauseTill timeGetTime() + 100
            
            DoEvents
            
        Loop Until Prog_paused = False
        
        'Return the old string to panel 2 of the status bar
        frmProgram.StatusBar TempS, 2
        
    End If



    Dim i As Integer
    Dim readVoltage As Double
    Dim TargetVoltage As Double
    
    Dim StartTime
    Dim ElapsedTime
    Dim ChargeTime As Double
        
    Dim PercentDone As String
    Dim TimeRemaining As String
    
    'First, lock the Coil selection
    CoilsLocked = True
    Me.chkLockCoils.value = Checked
    
    If (EnableAxialIRM = False And EnableTransIRM = False) _
       Or NOCOMM_MODE = True _
    Then Exit Sub
    
    If DEBUG_MODE Then frmDebug.msg "Fire IRM at " & pulse_volts & " V"
    IRMInterrupt = False
    
    'Update the program form status bar
    frmProgram.StatusBar "IRM Config", 2
    
    'Default IRM backfield mode to false
    SetIRMBackFieldMode False
    
    'If pulse_volts is negative, set the backfield mode to true, else
    'set it to false
    If pulse_volts < 0 And _
       EnableIRMBackfield = True _
    Then
    
        'Set the Backfield IRM mode = True
        SetIRMBackFieldMode True
        
        pulse_volts = -1 * pulse_volts
        
    ElseIf pulse_volts < 1 Then
    
        pulse_volts = 0
        SetIRMBackFieldMode False
            
    ElseIf EnableIRMBackfield = False Then
    
        SetIRMBackFieldMode False
        
    End If
    
    TargetVoltage = pulse_volts
    
    Dim DaqControlVoltage As Double
    
    'Convert capacitor pulse_volts into a DAQ Control pulse_volts
    DaqControlVoltage = ConvertPulseVoltsToMCCVolts(TargetVoltage)
    
    If DaqControlVoltage > modConfig.PulseVoltMax Then DaqControlVoltage = modConfig.PulseVoltMax
        
    'Convert back to Capacitor Volts for the displays on frmIRMARM
    TargetVoltage = ConvertMCCVoltsToPulseVolts(DaqControlVoltage)
    txtPulseVolts = Format$(TargetVoltage, "0.0")
    txtPulseField = Format$(ConvertPulseVoltsToGauss(TargetVoltage), "0.0")
    
    DaqControlVoltage = DaqControlVoltage * CalculateAscBoostMultiplier(TargetVoltage)
    
    If DaqControlVoltage > modConfig.PulseVoltMax Then DaqControlVoltage = modConfig.PulseVoltMax
       
    'Call function to set the relays into the correct configuration
    'If settings are not correct / relays were unable to set, Abort the IRM pulse fire process
    If SetRelaysForIRM = False Then
    
        ErrorMessage = "IRM system settings need to be changed.  Current IRM pulse has been aborted." & _
                       vbNewLine & _
                       "IRM Backfield Enabled = " & Trim(Str(modConfig.EnableIRMBackfield)) & _
                       vbNewLine & _
                       "Axial IRM Enabled = " & Trim(Str(modConfig.EnableAxialIRM)) & _
                       vbNewLine & _
                       "Transverse IRM Enabled = " & Trim(Str(modConfig.EnableTransIRM))

                
        SetCodeLevel CodeRed
        
        Flow_Pause
                
        'Send an email stating that the IRM settings are wrong
        frmSendMail.MailNotification "Bad IRM Settings", _
                                     ErrorMessage, _
                                     CodeRed, _
                                     True
                                     
        MsgBox ErrorMessage, , "Bad IRM Settings!"
        
        SetCodeLevel modStatusCode.StatusCodeColorLevelPrior
        
    End If
           
    ' clear pulse_volts
    '(July 2010 - I Hilburn) Replaces old frmMCC with frmDAQ_Comm using the Channel object variables
    frmDAQ_Comm.DoDAQIO IRMVoltageOut, 0
    txtMCCIRMVout = "0"
    
    'Wait for pulse_volts clear command to process through
    DelayTime 0.05
      
    'If we're in calibrate mode, pop-up picture-box in frmIRM_VoltageCalibration
    If CalibrationMode = True Then
    
        'Set the directions label in the picture box
        frmIRM_VoltageCalibration.lblDirections.Caption = "Waiting for IRM capacitor to charge"
        
        'Hide the text-box and highlighting
        frmIRM_VoltageCalibration.txtCapacitorVoltage.Visible = False
        frmIRM_VoltageCalibration.picHighlight.Visible = False
        
        'Hide the accept and redo buttons
        frmIRM_VoltageCalibration.cmdAccept.Visible = False
        frmIRM_VoltageCalibration.cmdRedo.Visible = False
        
        'Show the picture box
        frmIRM_VoltageCalibration.picGetCapacitorVoltage.Visible = True
        
    End If
    
    'Get the start time of the IRM charge process
    StartTime = timeGetTime()
    
    'Set Voltage to new IRM output pulse_volts target
    '(July 2010 - I Hilburn) Replaces old frmMCC with frmDAQ_Comm using the Channel object variables
    frmDAQ_Comm.DoDAQIO IRMVoltageOut, DaqControlVoltage
    txtMCCIRMVout = Str$(pulse_volts)
    
    'Update Status Panel 2
    frmProgram.StatusBar "IRM @ " & Trim(Str(Me.txtPulseVolts)) & " Volts", 2
    
    'Wait again for pulse_volts set command to process
    DelayTime 0.1
    
    frmDAQ_Comm.DoDAQIO IRMTrim, , TrimOnOff(False)
    Me.txtMCCIRMTrim = "Off"
    
    'Change the color on the IRMTrim Digital output user control
    Me.txtMCCIRMTrim.BackColor = QBColor(12)
    
    'Clear and reinstantiate irm_pulse_data collection
    Set irm_pulse_data = Nothing
    Set irm_pulse_data = New IRMData
    
    stop_charging_change_threshold_in_percent = 0.1
    stop_trimming_change_threshold_in_percent = 0.5
    
    'Update the IRM Charging status on the main form
    If TargetVoltage > 0 Then
    
        'Turn off trim / voltage bleed
        frmDAQ_Comm.DoDAQIO IRMTrim, , TrimOnOff(False)
        Me.txtMCCIRMTrim = "Off"
    
        lblIRMStatus.Caption = "Charging"
        
        'Update Program status Bar
        frmProgram.StatusBar "Charging...   0%", 3
                
        'Wait only 0.1 seconds
        DelayTime 0.1
                                        
        Dim still_charging As Boolean: still_charging = True
                
        Do While still_charging
                
            'Get the current IRM charge set pulse_volts
            readVoltage = IRMAverageVoltageIn
              
            If readVoltage >= TargetVoltage Then
                frmDAQ_Comm.DoDAQIO IRMVoltageOut, _
                                    DaqControlVoltage * (modConfig.PulseReturnMCCVoltConversion / _
                                                         modConfig.PulseMCCVoltConversion)
                Exit Do
                
            End If
              
            'Store pulse_volts and time
            irm_pulse_data.Add (timeGetTime() - StartTime), readVoltage
              
            If AtEndOfChargeCycle Then Exit Do
            
            PercentDone = Format(100 * readVoltage / TargetVoltage, "##0.0")
            PercentDone = PadLeft(PercentDone, 8) & "%"
            
            'Read in the poweramp pulse_volts
            Me.txtMCCIRMPowerAmpVoltageIn = Format(readVoltage, _
                                                   "#0.0#####")
            
            'Update the status display on the IRM form
            lblIRMStatus.Caption = "Charging... " & PercentDone
            
            'Update Program form
            frmProgram.StatusBar "Charging... " & PercentDone, 3
            
            DoEvents

            'If this IRM pulse is being run in calibration mode,
            'need to update the directions caption on frmIRM_VoltageCalibration
            If CalibrationMode = True Then

                'Format the time remaining
                TimeRemaining = Trim(Str(CLng((ChargeTime - ElapsedTime) / 1000)))

                TimeRemaining = PadLeft(TimeRemaining, 3)

                'Update the picture box
                frmIRM_VoltageCalibration.lblDirections.Caption = _
                    "Waiting for IRM Capacitor to charge." & vbNewLine & _
                    "Prepare to read IRM Box pulse_volts display in: " & _
                    TimeRemaining & " sec."

                If CLng(ChargeTime - ElapsedTime) < 3000 Then

                    'Show the text-box with a pink highlight
                    frmIRM_VoltageCalibration.txtCapacitorVoltage.Visible = True

                    frmIRM_VoltageCalibration.picHighlight.BackColor = QBColor(13)
                    frmIRM_VoltageCalibration.picHighlight.Visible = True

                End If

            End If
            
            'Same for either system
            If IRMInterrupt Then
                
                IRMInterrupt = False
                
                lblIRMStatus.Caption = "Charge Interrupted"
                                                               
                'Update program status bar
                frmProgram.StatusBar "Charging... Interrupted!", 3
                
                'Pause 2 seconds
                PauseTill timeGetTime() + 2000
                
                'Wipe the status bars clean
                frmProgram.StatusBar vbNullString, 2
                frmProgram.StatusBar vbNullString, 3
                
                Exit Sub
            
            End If
            
            
                    
        Loop
        
        'Update program status bar
        frmProgram.StatusBar "Charging... Done!", 3
            
    ElseIf TargetVoltage = 0 Then
    
        Dim still_trimming As Boolean: still_trimming = True
        
        'Turn on trim / voltage bleed
        frmDAQ_Comm.DoDAQIO IRMTrim, , TrimOnOff(True)
        Me.txtMCCIRMTrim = "On"
    
        'Loop until capacitor voltage is less than 3 V
        Do While still_trimming
            
            readVoltage = IRMAverageVoltageIn
            
            If modConfig.AscIrmMaxFireAtZeroGaussReadVoltage <= 0 Then modConfig.AscIrmMaxFireAtZeroGaussReadVoltage = 3
            If readVoltage <= modConfig.AscIrmMaxFireAtZeroGaussReadVoltage Then Exit Do
            
            'Store pulse_volts and time
            irm_pulse_data.Add (timeGetTime() - StartTime), readVoltage
              
            If AtEndOfTrimCycle Then Exit Do
            
            If readVoltage = 0 Then
                PercentDone = "100.00"
            Else
                PercentDone = Format(100 * Abs(readVoltage - modConfig.AscIrmMaxFireAtZeroGaussReadVoltage) / readVoltage, "##0.0")
                PercentDone = PadLeft(PercentDone, 8) & "%"
            End If
            
            'Read in the poweramp pulse_volts
            Me.txtMCCIRMPowerAmpVoltageIn = Format(readVoltage, _
                                                   "#0.0#####")
            
            'Update the status display on the IRM form
            lblIRMStatus.Caption = "Trimming... " & PercentDone
            
            'Update Program form
            frmProgram.StatusBar "Trimming... " & PercentDone, 3
            
            DoEvents
            
            'Same for either system
            If IRMInterrupt Then
                
                IRMInterrupt = False
                
                lblIRMStatus.Caption = "Trim Interrupted"
                                                               
                'Update program status bar
                frmProgram.StatusBar "Trimming... Interrupted!", 3
                
                'Pause 2 seconds
                PauseTill timeGetTime() + 2000
                
                'Wipe the status bars clean
                frmProgram.StatusBar vbNullString, 2
                frmProgram.StatusBar vbNullString, 3
                
                Exit Sub
            
            End If
            
            'Pause 50 ms
            PauseTill timeGetTime + 50
                
        Loop
                
        'We're ready to fire the IRM coils at charge voltage = 0
        'Turn off trim / voltage bleed
        frmDAQ_Comm.DoDAQIO IRMTrim, , TrimOnOff(False)
        Me.txtMCCIRMTrim = "Off"
        
    End If
        
    'Now, if this is a calibration run,
    'Write the Return voltage to frmIRM_VoltageCalibration grid
    If CalibrationMode = True Then
    
        With frmIRM_VoltageCalibration.gridVoltageCal
        
            .row = frmIRM_VoltageCalibration.CurrentRow
            
            'Store the Output DAQ Voltage to the IRM capacitor box
            .Col = 2
            .text = Format(pulse_volts, "#0.0#####")
                        
            'Store the Return Voltage from the IRM capacitor box
            .Col = 3
            .text = Format(readVoltage * modConfig.PulseReturnMCCVoltConversion, "#0.0#####")
    
            'Resize the 2nd and 3rd columns of the grid
            ResizeGrid frmIRM_VoltageCalibration.gridVoltageCal, _
                       frmIRM_VoltageCalibration, , , _
                       2, _
                       3
                           
        End With
    
        'Update picture box and tell user to write in the Calibration display voltage
        frmIRM_VoltageCalibration.lblDirections = "Write in the highest reached IRM capacitor box voltage."
        frmIRM_VoltageCalibration.txtCapacitorVoltage.Visible = True
        frmIRM_VoltageCalibration.picHighlight.BackColor = QBColor(4)
        frmIRM_VoltageCalibration.picHighlight.Visible = True
                    
    End If
    
    'IRM box is all charged up, can fire the IRM now
    ' fire IRM
        
    'Read in the current peak voltage (fewer read in's to lessen amount of time hanging here)
    readVoltage = IRMAverageVoltageIn(10)
    
    'Save readVoltage to peak voltage
    IRMPeakVoltage = readVoltage
    
    'Update program status bar and frmIRMARM display
    frmProgram.StatusBar "Firing!", 3
    lblIRMStatus.Caption = "Firing"
        
    '(July 2010 - I Hilburn) Replaces old frmMCC with frmDAQ_Comm using the Channel object variables
    'Close the TTL switch to connect the IRM circuit
    frmDAQ_Comm.DoDAQIO IRMFire, , False
    txtMCCIRMFire = "On!"
    Me.txtMCCIRMFire.BackColor = QBColor(10)
    
    'Pause while the IRM pulse
    'goes through the AF Coil
    'Check to see which coil is active to see if we
    'need to pause three times longer for the transverse coil IRM pulse
    Dim amplitude_dependent_wait_factor As Double
    amplitude_dependent_wait_factor = 0
    
'    If TargetVoltage >= modConfig.IRMAxialVoltMax Then
'
'        amplitude_dependent_wait_factor = 3
'
'    Else
'
'        amplitude_dependent_wait_factor = (modConfig.IRMAxialVoltMax - TargetVoltage) / modConfig.IRMAxialVoltMax
'
'    End If
        
    If ActiveCoilSystem = AxialCoilSystem Then
    
        'Pause for a second
        DelayTime 1 + Round(amplitude_dependent_wait_factor, 2)
                    
    ElseIf ActiveCoilSystem = TransverseCoilSystem Then
    
        'Pause for 3 seconds
        DelayTime 3 + Round(amplitude_dependent_wait_factor, 2)
        
    Else
    
        'Delay for 1 second
        DelayTime 1
        
    End If
        
    'Reset IRM fire status
    'Open the TTL switch to break the IRM circuit
    frmDAQ_Comm.DoDAQIO IRMFire, , True
    frmDAQ_Comm.DoDAQIO IRMVoltageOut, 0
    
    Me.txtMCCIRMFire = "Off"
    Me.txtMCCIRMFire.BackColor = QBColor(12)
    txtMCCIRMVout = "0"
        
    
    lblIRMStatus.Caption = vbNullString
    
    'Update program status bar - IRM Pulse done
    frmProgram.StatusBar vbNullString, 3
    frmProgram.StatusBar vbNullString, 2
            
    'Last, unlock the coil selection
    CoilsLocked = False
    Me.chkLockCoils.value = Unchecked
        
End Sub


Public Function CalculateAscBoostMultiplier(ByVal target_capacitor_voltage As Double) As Double

    CalculateAscBoostMultiplier = 1
    
    If modConfig.IRMAxialVoltMax <= 0 Then Exit Function
    
    Dim ratio_to_max As Double
    
    ratio_to_max = target_capacitor_voltage / modConfig.IRMAxialVoltMax
    
    If modConfig.AscSetVoltageMaxBoostMultiplier <= _
       modConfig.AscSetVoltageMinBoostMultiplier Then
       
       CalculateAscBoostMultiplier = Round(modConfig.AscSetVoltageMinBoostMultiplier - _
                                            (ratio_to_max) * (modConfig.AscSetVoltageMinBoostMultiplier - _
                                                              modConfig.AscSetVoltageMaxBoostMultiplier), 2)
                                                      
    Else
    
        CalculateAscBoostMultiplier = Round(modConfig.AscSetVoltageMinBoostMultiplier + _
                                                (ratio_to_max) * (modConfig.AscSetVoltageMaxBoostMultiplier - _
                                                                  modConfig.AscSetVoltageMinBoostMultiplier), 2)

    End If

End Function

Public Sub FireIRM(ByVal voltage As Double, _
                   Optional ByVal CalibrationMode As Boolean = False)

    'Default form to Axial IRM if Transverse IRM is disabled
    If modConfig.EnableAxialIRM And _
       Not modConfig.EnableTransIRM Then
       
       Me.optCoil(0).value = True
       ActiveCoilSystem = AxialCoilSystem
       
    End If

    If modConfig.IRMSystem = "ASC" Then
    
        FireASC_IrmAtPulseVolts voltage, CalibrationMode
        Exit Sub
        
    End If

'-----------------------------------------------------------------------------------------------------------------------------'
'-----------------------------------------------------------------------------------------------------------------------------'
'
'   Major Code Mod
'   June 2010, Isaac Hilburn
'
'   Added in if/then statements to accomodate the user's AF system settings.
'   With a 2G AF system, the AF Demag box controls the AF/IRM relays.
'   With the ADWIN AF system, the ADWIN board controls the AF/IRM relays.
'-----------------------------------------------------------------------------------------------------------------------------'
'-----------------------------------------------------------------------------------------------------------------------------'

'Do coil temperature check before running the IRM pulse
'------------------------------------------------------------------------------------------------------------------------------------'
'------------------------------------------------------------------------------------------------------------------------------------'
'
'   July 2011
'   Authors:
'   Laurent Corporozen
'   Isaac Hilburn
'
'   Copied temperature check code from frmAF_2G.ExecuteRamp to the IRM fire function
'   Code copied verbatim with minor changes to switch comm implementation from frmMCC to frmDAQ_Comm
'   frmDAQ_Comm takes Channel objects instead of channel port numbers
'------------------------------------------------------------------------------------------------------------------------------------'
        
    Dim TempD As Double
    Dim TempS As String
    Dim TempL As Long
    Dim Temp1 As Double
    Dim Temp2 As Double
    Dim TWarning As Boolean
    Dim ErrorMessage As String
        
    On Error GoTo oops:
        
    If voltage > 0 Then
             
         'Before doing anything with the ADWIN board, get the AF coil temperatures
         '(July 2010 - I Hilburn) Replaces old frmMCC with frmDAQ_Comm using the Channel object variables
         If EnableT1 Then
         
             Temp1 = frmDAQ_Comm.DoDAQIO(AnalogT1)
             
             Temp1 = Temp1 * TSlope - Toffset
             
         End If
             
         'Update display on the IRM form
         txtTemp1 = Format$(Temp1, "##0.00")
             
         If EnableT2 Then
         
             Temp2 = frmDAQ_Comm.DoDAQIO(AnalogT2)
             
             Temp2 = Temp2 * TSlope - Toffset
         
         End If
         
         'Update display on the IRM Form
         txtTemp2 = Format$(Temp2, "##0.00")
         
         'Check Temperature to see if it is not zeroed (gone within 20 deg of -1 * Toffset)
         If Not ValidSensorTemp(Temp1, Temp2) Then
         
             'Start code to tell user that the temp sensor values are bad
             NotifySensorError Temp1, Temp2
             
         End If
             
         
         If EnableT1 Or EnableT2 Then
             
             Do While Temp1 >= Thot Or Temp2 >= Thot
                 
                 lblAFtooHot.Visible = True
                 txtTemp1.BackColor = ColorOrange
                 txtTemp2.BackColor = ColorOrange
                 
                 ErrorMessage = "The IRM Coil(s) temperature is above " & Thot & "°C: " & Format$(Temp1, "##0.00") & _
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
                     
                         Temp1 = frmDAQ_Comm.DoDAQIO(AnalogT1)
                         
                         Temp1 = Temp1 * TSlope - Toffset
                         
                     End If
                     
                     txtTemp1 = Format$(Temp1, "##0.00")
                     
                     If EnableT2 Then
                     
                         Temp2 = frmDAQ_Comm.DoDAQIO(AnalogT2)
                         
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
        
    End If
        
    On Error GoTo 0
   
oops:
        
    '(April 30, 2011) I Hilburn
    'Added in check to see if the program flow is paused.  If so, the AF code will wait until the flow
    'is returned to Resume
    If Prog_paused = True Then
    
        'Set TempS equal to the current contents
        'of the 2nd status bar panel
        TempS = frmProgram.sbStatusBar.Panels(2).text
        
        'Update the program form status bar
        frmProgram.StatusBar "Paused...", 2
        
        'Loop and wait for status to change and DoEvents
        'to allow user to make changes
        Do
        
            'Pause 100 ms
            PauseTill timeGetTime() + 100
            
            DoEvents
            
        Loop Until Prog_paused = False
        
        'Return the old string to panel 2 of the status bar
        frmProgram.StatusBar TempS, 2
        
    End If



    Dim i As Integer
    Dim readySignals As Integer
    Dim readVoltage As Double
    Dim TargetVoltage As Double
    Dim fakeTargetV As Double       'This is a voltage > TargetVoltage used to get the target voltage
                                    'to the desired value at the lower voltages
    Dim deltaV As Double
    Dim preV As Double
    
    Dim StartTime
    Dim ElapsedTime
    Dim ChargeTime As Double
        
    Dim PercentDone As String
    Dim TimeRemaining As String
    
    'First, lock the Coil selection
    CoilsLocked = True
    Me.chkLockCoils.value = Checked
    
    '(September 25, 2010 - I Hilburn)
    'Altered this IRM exit condition
    'If both IRM axes are disabled, or we are in NOCOMM_MODE
    'Then exit the FireIRM subroutine
    'No longer allowing a NOCOMM_MODE IRM charge/fire cycle
    If (EnableAxialIRM = False And EnableTransIRM = False) _
       Or NOCOMM_MODE = True _
    Then Exit Sub
    
    If DEBUG_MODE Then frmDebug.msg "Fire IRM at " & voltage & " V"
    IRMInterrupt = False
    
    'Update the program form status bar
    frmProgram.StatusBar "IRM Config", 2
    
    'Default IRM backfield mode to false
    SetIRMBackFieldMode False
    
    'If voltage is negative, set the backfield mode to true, else
    'set it to false
    If voltage < 0 And _
       EnableIRMBackfield = True _
    Then
    
        'Set the Backfield IRM mode = True
        SetIRMBackFieldMode True
        
        voltage = -1 * voltage
        
    ElseIf voltage < 1 Then
    
        voltage = 0
        SetIRMBackFieldMode False
            
    ElseIf EnableIRMBackfield = False Then
    
        SetIRMBackFieldMode False
        
    End If
    
    TargetVoltage = voltage
    
    'For the Matsusada system, need to adjust the target voltage
    If modConfig.IRMSystem = "Matsusada" And voltage <> 0 Then
         
        'Now scale voltage up depending on it's value
        fakeTargetV = ScaleUp(voltage)
       
        'Convert capacitor voltage into a DAQ Control voltage
        voltage = ConvertPulseVoltsToMCCVolts(voltage)
        
        'Check to make sure that we don't exceed the max allowed DAQ control voltage
        If voltage > PulseVoltMax Then voltage = PulseVoltMax
        
        'Convert back to Capacitor Volts for the displays on frmIRMARM
        'Back scale down by the difference between the fake target (scaled up voltage) and
        'the actual desired target.  We don't want to confuse the user by displaying
        'a voltage greater than the one they requested.
        txtPulseVolts = Format$(ConvertMCCVoltsToPulseVolts(voltage) / fakeTargetV * TargetVoltage, "0.0")
        txtPulseField = Format$(ConvertPulseVoltsToGauss(TargetVoltage), "0.0")
    
    Else
    'For the "Old" IRM system, voltage scaling would only make things worse
    'so we're not going to use it.
        
        'Convert capacitor voltage into a DAQ Control voltage
        voltage = ConvertPulseVoltsToMCCVolts(voltage)
        
        'Check to make sure that we don't exceed the max allowed DAQ control voltage
        If voltage > PulseVoltMax Then voltage = PulseVoltMax
        
        'Convert back to Capacitor Volts for the displays on frmIRMARM
        txtPulseVolts = Format$(ConvertMCCVoltsToPulseVolts(voltage), "0.0")
        txtPulseField = Format$(ConvertPulseVoltsToGauss(TargetVoltage), "0.0")
        
    End If
       
    'Call function to set the relays into the correct configuration
    'If settings are not correct / relays were unable to set, Abort the IRM pulse fire process
    If SetRelaysForIRM = False Then
    
        ErrorMessage = "IRM system settings need to be changed.  Current IRM pulse has been aborted." & _
                       vbNewLine & _
                       "IRM Backfield Enabled = " & Trim(Str(modConfig.EnableIRMBackfield)) & _
                       vbNewLine & _
                       "Axial IRM Enabled = " & Trim(Str(modConfig.EnableAxialIRM)) & _
                       vbNewLine & _
                       "Transverse IRM Enabled = " & Trim(Str(modConfig.EnableTransIRM))

                
        SetCodeLevel CodeRed
        
        Flow_Pause
                
        'Send an email stating that the IRM settings are wrong
        frmSendMail.MailNotification "Bad IRM Settings", _
                                     ErrorMessage, _
                                     CodeRed, _
                                     True
                                     
        MsgBox ErrorMessage, , "Bad IRM Settings!"
        
        SetCodeLevel modStatusCode.StatusCodeColorLevelPrior
        
    End If
           
    ' clear voltage
    '(July 2010 - I Hilburn) Replaces old frmMCC with frmDAQ_Comm using the Channel object variables
    frmDAQ_Comm.DoDAQIO IRMVoltageOut, 0
    txtMCCIRMVout = "0"
    
    'Wait for voltage clear command to process through
    DelayTime 0.05
'------------------------------------------------------------------------------------------------------------'
'
'   Commented Out:  September 27, 2010
'              By:  Isaac Hilburn
'
'          Reason:  Code no longer needed with direct feedback voltage from the IRM capacitor box.
'                   Don't need to wait a set amount of time after the Matsusada box reaches the target voltage
'------------------------------------------------------------------------------------------------------------'
'    'If this is the Matsusada IRM system, determine how long the charge cycle will be
'    'It takes ~ 20 seconds to reach 400 volts, so, pause for
'    '5 - 20 seconds for 50 - 400 volts, and 5 seconds for everything lower than that
'    '(Note, charge time is in milliseconds)
'    If TargetVoltage > 250 Then
'
'        ChargeTime = 0.1 * TargetVoltage * 1000
'
'    ElseIf TargetVoltage > 100 Then
'
'        ChargeTime = 0.15 * TargetVoltage * 1000
'
'    ElseIf TargetVoltage > 50 Then
'
'        ChargeTime = 0.25 * TargetVoltage * 1000
'
'    Else
'
'        'It takes longer to charge up, proportionally, at lower voltages
'        'Lower IRM pulses need a full 15 seconds to charge before firing.
'        ChargeTime = 15000
'
'    End If
'
'    'Make sure have at least 15 seconds worth of charge time
'    If ChargeTime < 15000 Then ChargeTime = 15000
'
'------------------------------------------------------------------------------------------------------------'
    
    
    'If we're in calibrate mode, pop-up picture-box in frmIRM_VoltageCalibration
    If CalibrationMode = True Then
    
        'Set the directions label in the picture box
        frmIRM_VoltageCalibration.lblDirections.Caption = "Waiting for IRM capacitor to charge"
        
        'Hide the text-box and highlighting
        frmIRM_VoltageCalibration.txtCapacitorVoltage.Visible = False
        frmIRM_VoltageCalibration.picHighlight.Visible = False
        
        'Hide the accept and redo buttons
        frmIRM_VoltageCalibration.cmdAccept.Visible = False
        frmIRM_VoltageCalibration.cmdRedo.Visible = False
        
        'Show the picture box
        frmIRM_VoltageCalibration.picGetCapacitorVoltage.Visible = True
        
    End If
    
    'Get the start time of the IRM charge process
    StartTime = timeGetTime()
    
    'Set Voltage to new IRM output voltage target
    '(July 2010 - I Hilburn) Replaces old frmMCC with frmDAQ_Comm using the Channel object variables
    frmDAQ_Comm.DoDAQIO IRMVoltageOut, voltage
    txtMCCIRMVout = Str$(voltage)
    
    'Update Status Panel 2
    frmProgram.StatusBar "IRM @ " & Trim(Str(Me.txtPulseVolts)) & " Volts", 2
    
    'Wait again for voltage set command to process
    DelayTime 0.1
    
    'Turn off the IRM Trim (so the voltage doesn't bleed away)
    '(July 2010 - I Hilburn) Replaces old frmMCC with frmDAQ_Comm using the Channel object variables
    '(February 2011 - I Hilburn)
    'Bug: The IRMTrim - trim on setting is the logical reverse for the old system than for the
    '     new Matsusada system. (yay!)
    '     Old system: True = Trim On
    '                 False = Trim Off
    '     Matsusada system: False = Trim On
    '                       True = Trim Off
    '     Ugh.
    'Fix: Call newly created Functions TrimOnOff, True = Trim On, False = Trim On
    '     The new function will handle what to do for the various systems.
    '     In the future, if there's even more variation in this, may need to
    '     add settings to frmSettings for the user to set True = Trim On or False = Trim On
    '     Double Ugh.
    frmDAQ_Comm.DoDAQIO IRMTrim, , TrimOnOff(False)
    Me.txtMCCIRMTrim = "Off"
    
    'Change the color on the IRMTrim Digital output user control
    Me.txtMCCIRMTrim.BackColor = QBColor(12)
    
    'Update the IRM Charging status on the main form
    If TargetVoltage > 0 Then
    
        lblIRMStatus.Caption = "Charging"
        
        'Update Program status Bar
        frmProgram.StatusBar "Charging...   0%", 3
        
    
        ' wait for the capacitor voltage to be within 0.5% of target
        deltaV = TargetVoltage * 0.005
        
        readySignals = 0
        
        'Zero the IRM Past voltages array
        For i = 0 To UBound(IRMPastVolts) - 1
        
            IRMPastVolts(i) = 0
            
        Next i
                        
        'Check to see which system is being used
        If IRMSystem = "Old" Then
            
            'If DeltaV is less than 0.2 Volts, set it to 0.2 volts
            If deltaV < 0.5 Then deltaV = 0.5
            
            'If this is the old system, wait two seconds,
            DelayTime 2
            
            Do While readySignals < 3 And TargetVoltage > 0 And Not IRMInterrupt
        
                'Get the current IRM voltage
                readVoltage = IRMAverageVoltageIn
                
                'Update the charge level on the IRM/ARM control
                lblIRMStatus.Caption = "Charging: " & Format$(100 * (readVoltage / TargetVoltage), "##0.0") & "%"
                
                PercentDone = Format(100 * readVoltage / TargetVoltage, "##0.0")
                PercentDone = PadLeft(PercentDone, 8) & "%"
                                   
                'Update program status bar
                frmProgram.StatusBar "Charging... " & PercentDone, 3
                                    
                If Abs(readVoltage - TargetVoltage) < deltaV Then
                    
                    'Count up to three reads where the IRM charge voltage is within the target by 1%
                    readySignals = readySignals + 1
                    
                    'Zero the IRM output voltage to allow the voltage level to stay constant
                    '(July 2010 - I Hilburn) Replaces old frmMCC with frmDAQ_Comm using the Channel object variables
                    frmDAQ_Comm.DoDAQIO IRMVoltageOut, 0
                    txtMCCIRMVout = "0"
                    
                    'Wait before next status update
                    DelayTime 0.1
                    
                'IRM voltage is too high!
                ElseIf readVoltage > TargetVoltage Then
                
                    If CalibrationMode = True Then
                    
                        'Check for a plateau
                        'Refresh an array storing the past N read-in voltages
                        UpdateIRMPast IRMPastVolts, readVoltage
                        
                        'Now check to see if the read in voltages are flat-lining or decreasing
                        If CheckIRMPast(IRMPastVolts, TargetVoltage) = True Then
                        
                            readySignals = 3
                            
                        End If
                        
                        'Tell the IRM control to keep charging up the voltage
                        '(July 2010 - I Hilburn) Replaces old frmMCC with frmDAQ_Comm using the Channel object variables
                        frmDAQ_Comm.DoDAQIO IRMVoltageOut, voltage
                        txtMCCIRMVout = Format(voltage, "#0.0#####")
                            
                        
                    Else
                        
                        'Need to trim voltage, if drops below, then FIRE!
                        If IRMBleedBelowVoltage(voltage, deltaV, False, True) = False Then
                            readySignals = 4
                        End If
                                    
                    End If
                                    
                'IRM Voltage is too low.
                ElseIf readVoltage < TargetVoltage Then
                    
                    'Need to charge up more
                    readySignals = 0
                    
                    'Need code to see if we're hanging at a plateau
                    'Refresh an array storing the past 10 read in voltages
                    UpdateIRMPast IRMPastVolts, readVoltage
                    
                    'Now check to see if the read in voltages are flat-lining or decreasing
                    If CheckIRMPast(IRMPastVolts, TargetVoltage) = True Then
                    
                        readySignals = 3
                        
                    End If
                    
                    'Tell the IRM control to keep charging up the voltage
                    '(July 2010 - I Hilburn) Replaces old frmMCC with frmDAQ_Comm using the Channel object variables
                    frmDAQ_Comm.DoDAQIO IRMVoltageOut, voltage
                    txtMCCIRMVout = Format(voltage, "#0.0#####")
                    
                End If
                
                'Same for either system
                If IRMInterrupt Then
                    
                    IRMInterrupt = False
                    
                    lblIRMStatus.Caption = "Charge Interrupted"
                                                                   
                    'Update program status bar
                    frmProgram.StatusBar "Charging... Interrupted!", 3
                    
                    'Pause 2 seconds
                    PauseTill timeGetTime() + 2000
                    
                    'Wipe the status bars clean
                    frmProgram.StatusBar vbNullString, 2
                    frmProgram.StatusBar vbNullString, 3
                    
                    Exit Sub
                
                End If
                
                'Wait 0.05 seconds between loops
                DelayTime 0.05
                
            Loop
            
            'Update program status bar
            frmProgram.StatusBar "Charging... Done!", 3
            
        Else
        
            'New Matsusada system has a higher accuracy level -
            'can get charge to within 0.2 volts
            If deltaV > 0.2 Then deltaV = 0.2
            
            'Set preV = 0.1
            'This is to help prevent overshoots and all the trimming that that causes
            preV = 0.1
            
            'Check the ratio of preV to the target voltage
            If preV / TargetVoltage < 0.005 Then preV = TargetVoltage * 0.005
            
            'Set ready status = Charging
            readySignals = 0
            
            'Wait only 0.1 seconds
            DelayTime 0.1
                                            
            Do While readySignals < 3
                
                'When the
                
                'Get the current IRM charge set voltage
                readVoltage = IRMAverageVoltageIn
                  
                PercentDone = Format(100 * readVoltage / TargetVoltage, "##0.0")
                PercentDone = PadLeft(PercentDone, 8) & "%"
                
                'Read in the poweramp voltage
                Me.txtMCCIRMPowerAmpVoltageIn = Format(CDbl(frmDAQ_Comm.DoDAQIO(IRMPowerAmpVoltageIn)), _
                                                       "#0.0#####")
                
                'Update the status display on the IRM form
                lblIRMStatus.Caption = "Charging... " & PercentDone
                
                'Update Program form
                frmProgram.StatusBar "Charging... " & PercentDone, 3
                

                'If this IRM pulse is being run in calibration mode,
                'need to update the directions caption on frmIRM_VoltageCalibration
                If CalibrationMode = True Then

                    'Format the time remaining
                    TimeRemaining = Trim(Str(CLng((ChargeTime - ElapsedTime) / 1000)))

                    TimeRemaining = PadLeft(TimeRemaining, 3)

                    'Update the picture box
                    frmIRM_VoltageCalibration.lblDirections.Caption = _
                        "Waiting for IRM Capacitor to charge." & vbNewLine & _
                        "Prepare to read IRM Box voltage display in: " & _
                        TimeRemaining & " sec."

                    If CLng(ChargeTime - ElapsedTime) < 3000 Then

                        'Show the text-box with a pink highlight
                        frmIRM_VoltageCalibration.txtCapacitorVoltage.Visible = True

                        frmIRM_VoltageCalibration.picHighlight.BackColor = QBColor(13)
                        frmIRM_VoltageCalibration.picHighlight.Visible = True

                    End If

                End If
                
                If Abs(readVoltage - TargetVoltage) < deltaV + preV Then
                    
                    'Need to break loop instantly, or voltage will overshoot
                    readySignals = 4
                    
'------------------------------------------------------------------------------------------------------------'
                    '(July 2, 2011 - I Hilburn)
                    ' Uncommented this code.  We need to zero the IRM capacitor control
                    ' voltage because the target voltage that was actually sent to the
                    ' capacitor box is greater than the desired target value (the value
                    ' was scaled up, see function ScaleUp above)
                    
                    '(September 29, 2010 - I Hilburn)
                    ' Commented out the code to zero the driving voltage sent to the
                    ' Matsusada power amp during a non-calibration IRM charge.
                    ' New code lets IRM return voltage reach a plateau
                    ' before firing.  Yields a more consistent final charge
'------------------------------------------------------------------------------------------------------------'
                    If CalibrationMode = False Then

                        'Zero the IRM output voltage to allow the voltage level to stay constant
                        frmDAQ_Comm.DoDAQIO IRMVoltageOut, 0
                        txtMCCIRMVout = "0"

                    End If
'------------------------------------------------------------------------------------------------------------'
                    
                'IRM voltage is too high!
                ElseIf readVoltage > TargetVoltage Then
                    
                    
                    '(July 2, 2011 - I Hilburn)
                    ' Only Allow the IRM voltage to plateau if this is a voltage calibration run
                    ' otherwise, the voltage needs to be trimmed
                    '(September 29, 2010 - I Hilburn)
                    ' Commented out the code - want the capacitor voltage to reach plateau
                    ' for all types of IRM charging - both voltage calibration and normal
                    If CalibrationMode = True Then
                    
                        'Check for a plateau
                        'Refresh an array storing the past 10 read in voltages
                        UpdateIRMPast IRMPastVolts, readVoltage
                        
                        'Now check to see if the read in voltages are flat-lining or decreasing
                        If CheckIRMPast(IRMPastVolts, TargetVoltage) = True Then
                        
                            readySignals = 3
                            
                        End If
                    Else
                        
                        'We're not in Calibration mode, overshoots are bad!
                        
                        'Turn on the IRM trim
                        If IRMBleedBelowVoltage(TargetVoltage, deltaV, False, True) = False Then
                                readySignals = 4
                        End If
                        
                    End If
                        
                'IRM Voltage is too low.
                'However, if prior code has given the go ahead to fire, ignore this
                ElseIf readVoltage < TargetVoltage And readySignals < 3 Then
                    
                    'Need to charge up more
                    readySignals = 0
                    
                    'Need code to see if we're hanging at a plateau
                    'Refresh an array storing the past 10 read in voltages
                    UpdateIRMPast IRMPastVolts, readVoltage
                    
                    'Now check to see if the read in voltages are flat-lining or decreasing
                    If CheckIRMPast(IRMPastVolts, TargetVoltage) = True Then
                    
                        readySignals = 3
                        
                    End If
                    
                    'Tell the IRM control to keep charging up the voltage
                    '(July 2010 - I Hilburn) Replaces old frmMCC with frmDAQ_Comm using the Channel object variables
                    frmDAQ_Comm.DoDAQIO IRMVoltageOut, voltage
                    txtMCCIRMVout = Format(voltage, "#0.0#####")
                    
                End If
                
                'Same for either system
                If IRMInterrupt Then
                    
                    IRMInterrupt = False
                    
                    lblIRMStatus.Caption = "Charge Interrupted"
                                                                   
                    'Update program status bar
                    frmProgram.StatusBar "Charging... Interrupted!", 3
                    
                    'Pause 2 seconds
                    PauseTill timeGetTime() + 2000
                    
                    'Wipe the status bars clean
                    frmProgram.StatusBar vbNullString, 2
                    frmProgram.StatusBar vbNullString, 3
                    
                    Exit Sub
                
                End If
                
                'if still charge, then Pause 50 ms between each loop
                If readySignals < 3 Then PauseTill timeGetTime() + 50
                        
            Loop
            
            'Update program status bar
            frmProgram.StatusBar "Charging... Done!", 3
            
        End If
            
'------------------------------------------------------------------------------------------------------------'
'   Commented Out:  September 25, 2010
'              By:  Isaac Hilburn
'          Reason:  Code no longer needed - IRM Ready digital Input no longer setup.
'                   EnableIRMReturn has been removed from the code
'------------------------------------------------------------------------------------------------------------'
'        Else
'
'
'            'Don't have things setup to get a return voltage
'            'on the Old IRM system
'
'            'Set a general status of "Charging"
'            lblIRMStatus.Caption = "Charging"
'
'            'Update program status bar
'            frmProgram.StatusBar "Charging... ", 3
'
'            'Check which IRM system we're using
'            If IRMSystem = "Old" Then
'
'                'There's no feedback voltage from the IRM box
'                'Just wait for IRM Ready status to change
'                DelayTime 2
'                For i = 1 To 3
'                    Do While Not (IRMIsReady = 0 Or NOCOMM_MODE Or IRMInterrupt)
'                        lblIRMStatus.Caption = "Charging"
'                        DelayTime 0.3
'                    Loop
'                    If IRMInterrupt Then
'                        IRMInterrupt = False
'                        lblIRMStatus.Caption = "Charge Interrupted"
'                        Exit Sub
'                    End If
'                    If (IRMIsReady = 0 Or NOCOMM_MODE) Then lblIRMStatus.Caption = "Ready Signal " & Str$(i) Else i = 0
'                    DelayTime 0.3
'                Next i
'
'            End If
'
'            'Update program status bar
'            frmProgram.StatusBar "Charging... Done!", 3
'
'        End If
'------------------------------------------------------------------------------------------------------------'

        
    ElseIf TargetVoltage = 0 Then
    
        'set readysignals = 0
        readySignals = 0
    
        'Loop until capacitor voltage is less than 3 V
        Do While readySignals < 3
            
            'Need to wait for the current voltage on the IRM to trim down to 0 +-3
            readVoltage = IRMAverageVoltageIn
            
            If readVoltage > 10 Then
            
                'Decrement readysignals
                readySignals = readySignals - 1
                
                'Voltage is still too high, need to bleed below the target voltage
                If readySignals < 0 Then IRMBleedBelowVoltage 10, _
                                                              CalibrationMode, _
                                                              True
                                                                              
            ElseIf readVoltage <= 10 Then
            
                readySignals = 3
                
            End If
            
            'Pause 50 ms
            PauseTill timeGetTime + 50
                
        Loop
                
        'We're ready to fire the IRM coils at charge voltage = 0
        
    End If
        
    'Now, if this is a calibration run,
    'Write the Return voltage to frmIRM_VoltageCalibration grid
    If CalibrationMode = True Then
    
        With frmIRM_VoltageCalibration.gridVoltageCal
        
            .row = frmIRM_VoltageCalibration.CurrentRow
            
            'Store the Output DAQ Voltage to the IRM capacitor box
            .Col = 2
            .text = Format(voltage, "#0.0#####")
                        
            'Store the Return Voltage from the IRM capacitor box
            .Col = 3
            .text = Format(readVoltage * modConfig.PulseReturnMCCVoltConversion, "#0.0#####")
    
            'Resize the 2nd and 3rd columns of the grid
            ResizeGrid frmIRM_VoltageCalibration.gridVoltageCal, _
                       frmIRM_VoltageCalibration, , , _
                       2, _
                       3
                           
        End With
    
        'Update picture box and tell user to write in the Calibration display voltage
        frmIRM_VoltageCalibration.lblDirections = "Write in the highest reached IRM capacitor box voltage."
        frmIRM_VoltageCalibration.txtCapacitorVoltage.Visible = True
        frmIRM_VoltageCalibration.picHighlight.BackColor = QBColor(4)
        frmIRM_VoltageCalibration.picHighlight.Visible = True
                    
    End If
    
    'IRM box is all charged up, can fire the IRM now
    ' fire IRM
    
    'If this is the "Old" IRM system, wait 1 second before firing
    If modConfig.IRMSystem = "Old" Then DelayTime 1
    
    'Read in the current peak voltage (fewer read in's to lessen amount of time hanging here)
    readVoltage = IRMAverageVoltageIn(10)
    
    'Save readVoltage to peak voltage
    IRMPeakVoltage = readVoltage
    
    'Update program status bar and frmIRMARM display
    frmProgram.StatusBar "Firing!", 3
    lblIRMStatus.Caption = "Firing"
        
    '(July 2010 - I Hilburn) Replaces old frmMCC with frmDAQ_Comm using the Channel object variables
    'Close the TTL switch to connect the IRM circuit
    frmDAQ_Comm.DoDAQIO IRMFire, , False
    txtMCCIRMFire = "On!"
    Me.txtMCCIRMFire.BackColor = QBColor(10)
    
    'Pause while the IRM pulse
    'goes through the AF Coil
    'Check to see which coil is active to see if we
    'need to pause three times longer for the transverse coil IRM pulse
    If ActiveCoilSystem = AxialCoilSystem Then
    
        'Pause for a second
        DelayTime 1
                    
    ElseIf ActiveCoilSystem = TransverseCoilSystem Then
    
        'Pause for 3 seconds
        DelayTime 3
        
    Else
    
        'Delay for 1 second
        DelayTime 1
        
    End If
        
    'Before breaking the IRM circuit, set the IRM voltage to 0
    '(July 2010 - I Hilburn) Replaces old frmMCC with frmDAQ_Comm using the Channel object variables
    frmDAQ_Comm.DoDAQIO IRMVoltageOut, 0
    txtMCCIRMVout = "0"
            
    'Reset IRM fire status
    'Open the TTL switch to break the IRM circuit
    '(July 2010 - I Hilburn) Replaces old frmMCC with frmDAQ_Comm using the Channel object variables
    frmDAQ_Comm.DoDAQIO IRMFire, , True
    Me.txtMCCIRMFire = "Off"
    Me.txtMCCIRMFire.BackColor = QBColor(12)
    
    lblIRMStatus.Caption = vbNullString
    
    'Update program status bar - IRM Pulse done
    frmProgram.StatusBar vbNullString, 3
    frmProgram.StatusBar vbNullString, 2
            
    'Last, unlock the coil selection
    CoilsLocked = False
    Me.chkLockCoils.value = Unchecked
        
End Sub

Public Function FireIRMAtField(ByVal Gauss As Double) As Double
    Dim PulseVoltsOut As Double
    Dim MCCVoltsOut As Double
    
    'First, lock the Coil selection
    CoilsLocked = True
    Me.chkLockCoils.value = Checked
        
    'Update the program form status bar
    frmProgram.StatusBar "IRM Config", 2
    
    If Gauss < 0 And EnableIRMBackfield Then
        SetIRMBackFieldMode True
        Gauss = -Gauss
        
        PulseVoltsOut = ConvertGaussToPulseVolts(Gauss)
        MCCVoltsOut = ConvertPulseVoltsToMCCVolts(PulseVoltsOut)
        PulseVoltsOut = -1 * PulseVoltsOut
        
    Else
    
        SetIRMBackFieldMode False
        PulseVoltsOut = ConvertGaussToPulseVolts(Gauss)
        MCCVoltsOut = ConvertPulseVoltsToMCCVolts(PulseVoltsOut)
        
        
    End If
       
    
    txtPulseVolts = PulseVoltsOut
    If MCCVoltsOut < 0 Then
        MCCVoltsOut = 0
    ElseIf MCCVoltsOut > 10 Then
        MCCVoltsOut = 10
    End If
    FireIRM PulseVoltsOut
    FireIRMAtField = ConvertMCCVoltsToPulseVolts(MCCVoltsOut)
    ' return true field
    
    'Update the program form status bar
    frmProgram.StatusBar vbNullString, 2
    
    'Last, unlock the coil selection
    CoilsLocked = False
    Me.chkLockCoils.value = Unchecked
    
End Function

Private Sub Form_Activate()

    If EnableAxialIRM = False And _
       EnableTransIRM = False _
    Then
    
        'No IRM coil is enabled, disable all the IRM Buttons on this form
        'and msg-box the user
        MsgBox "The IRM Axial & Transverse modules are currently disabled." & vbNewLine & _
               "No IRM's can be performed right now until those settings are changed."
               
        'Disable all the IRM buttons on the form (except the calibration buttons)
        Me.cmdInterruptCharge.Enabled = False
        Me.cmdIRMAverageVoltageIn.Enabled = False
        Me.cmdIRMFire.Enabled = False
        Me.cmdIRMFirebyGauss.Enabled = False
        Me.cmdMCCIRMFire.Enabled = False
        Me.cmdMCCIRMPowerAmpVoltageIn.Enabled = False
        Me.cmdMCCIRMTrim.Enabled = False
        Me.cmdMCCIRMVin.Enabled = False
        Me.cmdMCCIRMVout.Enabled = False
        Me.txtIRMAverageVoltageIn.Enabled = False
        Me.txtMCCIRMFire.Enabled = False
        Me.txtMCCIRMPowerAmpVoltageIn.Enabled = False
        Me.txtMCCIRMTrim.Enabled = False
        Me.txtMCCIRMVin.Enabled = False
        Me.txtMCCIRMVout.Enabled = False
        
    Else
    
        'Enable all the IRM buttons on the form (except the calibration buttons)
        Me.cmdInterruptCharge.Enabled = True
        Me.cmdIRMAverageVoltageIn.Enabled = True
        Me.cmdIRMFire.Enabled = True
        Me.cmdIRMFirebyGauss.Enabled = True
        Me.cmdMCCIRMFire.Enabled = True
        Me.cmdMCCIRMPowerAmpVoltageIn.Enabled = True
        Me.cmdMCCIRMTrim.Enabled = True
        Me.cmdMCCIRMVin.Enabled = True
        Me.cmdMCCIRMVout.Enabled = True
        Me.txtIRMAverageVoltageIn.Enabled = True
        Me.txtMCCIRMFire.Enabled = True
        Me.txtMCCIRMPowerAmpVoltageIn.Enabled = True
        Me.txtMCCIRMTrim.Enabled = True
        Me.txtMCCIRMVin.Enabled = True
        Me.txtMCCIRMVout.Enabled = True
        
    End If
    
    'Disable / Enable the Axial / Transverse IRM coil radio buttons depending on the Module settings
    Me.optCoil(1).Enabled = EnableTransIRM
    Me.optCoil(0).Enabled = EnableAxialIRM
        
    'Now do the same for the ARM
    If EnableARM = False Then
    
        'MsgBox the user
        MsgBox "ARM module is currently disabled. ARM bias voltage set is unavailable " & _
               "right now."
               
        'Disable all the necessary buttons
        Me.cmdMCCARMSet.Enabled = False
        Me.cmdMCCARMVout.Enabled = False
        Me.txtMCCARMSet.Enabled = False
        Me.txtBiasField.Enabled = False
        Me.txtMCCARMVout.Enabled = False

    Else
    
        'Disable all the necessary buttons
        Me.cmdMCCARMSet.Enabled = True
        Me.cmdMCCARMVout.Enabled = True
        Me.txtMCCARMSet.Enabled = True
        Me.txtBiasField.Enabled = True
        Me.txtMCCARMVout.Enabled = True

    End If

    'Set the coils-locked status
    If CoilsLocked = True Then Me.chkLockCoils.value = Checked
    If CoilsLocked = False Then Me.chkLockCoils.value = Unchecked

    'Set the correct active coil based on the active coil global variable
    If ActiveCoilSystem = AxialCoilSystem Then
    
        optCoil(0).value = True
        
    ElseIf ActiveCoilSystem = TransverseCoilSystem Then
    
        optCoil(1).value = True
        
    Else
    
        'No coil selected
        optCoil(0).value = False
        optCoil(1).value = False
        
        ActiveCoilSystem = NoCoilSystem
        
    End If

End Sub

Public Sub Form_Load()
        
    'Set the Form height & width
    Me.Height = 5430
    Me.Width = 7935
    
    '(March 11, 2011 - I Hilburn)
    'Update all the form's icons with the program Ico file
    SetFormIcon Me
    
    Me.Caption = "IRM / ARM Control"
    
    If EnableAxialIRM = False And _
       EnableTransIRM = False _
    Then
               
        'Disable all the IRM buttons on the form (except the calibration buttons)
        Me.cmdInterruptCharge.Enabled = False
        Me.cmdIRMAverageVoltageIn.Enabled = False
        Me.cmdIRMFire.Enabled = False
        Me.cmdIRMFirebyGauss.Enabled = False
        Me.cmdMCCIRMFire.Enabled = False
        Me.cmdMCCIRMPowerAmpVoltageIn.Enabled = False
        Me.cmdMCCIRMTrim.Enabled = False
        Me.cmdMCCIRMVin.Enabled = False
        Me.cmdMCCIRMVout.Enabled = False
        Me.txtIRMAverageVoltageIn.Enabled = False
        Me.txtMCCIRMFire.Enabled = False
        Me.txtMCCIRMPowerAmpVoltageIn.Enabled = False
        Me.txtMCCIRMTrim.Enabled = False
        Me.txtMCCIRMVin.Enabled = False
        Me.txtMCCIRMVout.Enabled = False
               
    Else
    
        'Enable all the IRM buttons on the form (except the calibration buttons)
        Me.cmdInterruptCharge.Enabled = True
        Me.cmdIRMAverageVoltageIn.Enabled = True
        Me.cmdIRMFire.Enabled = True
        Me.cmdIRMFirebyGauss.Enabled = True
        Me.cmdMCCIRMFire.Enabled = True
        Me.cmdMCCIRMPowerAmpVoltageIn.Enabled = True
        Me.cmdMCCIRMTrim.Enabled = True
        Me.cmdMCCIRMVin.Enabled = True
        Me.cmdMCCIRMVout.Enabled = True
        Me.txtIRMAverageVoltageIn.Enabled = True
        Me.txtMCCIRMFire.Enabled = True
        Me.txtMCCIRMPowerAmpVoltageIn.Enabled = True
        Me.txtMCCIRMTrim.Enabled = True
        Me.txtMCCIRMVin.Enabled = True
        Me.txtMCCIRMVout.Enabled = True
        
    End If
        
    Me.optCoil(0).Enabled = EnableAxialIRM
    Me.optCoil(1).Enabled = EnableTransIRM
        
    'Now do the same for the ARM
    If EnableARM = False Then
                   
        'Disable all the necessary buttons
        Me.cmdMCCARMSet.Enabled = False
        Me.cmdMCCARMVout.Enabled = False
        Me.txtMCCARMSet.Enabled = False
        Me.txtBiasField.Enabled = False
        Me.txtMCCARMVout.Enabled = False

    Else
    
        'Disable all the necessary buttons
        Me.cmdMCCARMSet.Enabled = True
        Me.cmdMCCARMVout.Enabled = True
        Me.txtMCCARMSet.Enabled = True
        Me.txtBiasField.Enabled = True
        Me.txtMCCARMVout.Enabled = True

    End If

    
    'Set the coils-locked status
    If CoilsLocked = True Then Me.chkLockCoils.value = Checked
    If CoilsLocked = False Then Me.chkLockCoils.value = Unchecked
    
    'Set the correct active coil based on the active coil global variable
    If ActiveCoilSystem = AxialCoilSystem Then
    
        optCoil(0).value = True
        
    ElseIf ActiveCoilSystem = TransverseCoilSystem Then
    
        optCoil(1).value = True
        
    Else
    
        'No coil selected
        optCoil(0).value = False
        optCoil(1).value = False
        
        ActiveCoilSystem = NoCoilSystem
        
    End If
    
    'Load the DAQ Comm form
    Load frmDAQ_Comm
    
    '(July 2010 - I Hilburn) Replaces old frmMCC with frmDAQ_Comm using the Channel object variables
    frmDAQ_Comm.DoDAQIO ARMVoltageOut, 0
    txtMCCARMVout = "0"
    frmDAQ_Comm.DoDAQIO IRMVoltageOut, 0
    txtMCCIRMVout = "0"
    
    '(July 2010 - I Hilburn) Replaces old frmMCC with frmDAQ_Comm using the Channel object variables
    frmDAQ_Comm.DoDAQIO IRMTrim, , TrimOnOff(True)
    txtMCCIRMTrim = "On"
    Me.txtMCCIRMTrim.BackColor = QBColor(10)
    
    '(July 2010 - I Hilburn) Replaces old frmMCC with frmDAQ_Comm using the Channel object variables
    frmDAQ_Comm.DoDAQIO IRMFire, , True
    txtMCCIRMFire = "Off"
    Me.txtMCCIRMFire.BackColor = QBColor(12)
    
    'Wiat a half second
    DelayTime 0.5
    
    '(July 2010 - I Hilburn) Replaces old frmMCC with frmDAQ_Comm using the Channel object variables
    frmDAQ_Comm.DoDAQIO ARMSet, , True
    txtMCCARMSet = "Off"
    txtMCCARMSet.BackColor = QBColor(12)
   
    'Set IRM to false
    SetIRMBackFieldMode False
    
End Sub

Private Sub Form_Resize()
    Me.Height = 5670
    Me.Width = 7935
End Sub

Public Function IRMAverageVoltageIn(Optional ByVal Times As Integer = 200) As Double
    Dim i As Integer
    Dim Sum As Double
    Dim working As Double
    For i = 1 To Times
        working = IRMCapacitorVoltage
        Sum = Sum + working
        txtMCCIRMVin = Format(working, "#0.0#####")
    Next i
    IRMAverageVoltageIn = Sum / Times / PulseReturnMCCVoltConversion
    txtIRMAverageVoltageIn = Format(IRMAverageVoltageIn, "#0.0#")
End Function

'------------------------------------------------------------------------------------------------------------'
'------------------------------------------------------------------------------------------------------------'
'
'   Function IRMIsReady
'   Commented Out:  September 27, 2010
'              By:  Isaac Hilburn
'
'          Reason:  IRM hardware no longer returns a digital ready status.
'                   This code is now obsolete
'
'------------------------------------------------------------------------------------------------------------'
'Public Function IRMIsReady() As Integer
'
'    Dim TempB As Boolean
'
'    'If the Old IRM system is in use
'    'Get the IRM ready status
'    If IRMSystem = "Old" Then
'
'        '(July 2010 - I Hilburn) Replaces old frmMCC with frmDAQ_Comm using the Channel object variables
'        TempB = CBool(frmDAQ_Comm.DoDAQIO(IRMPowerAmpVoltageIn))
'
'        IRMIsReady = CInt(TempB)
'
'        'We do use the IRM ready status
'        Me.txtMCCIRMPowerAmpVoltageIn.Enabled = True
'
'        'Determine display setting for IRM ready text-box
'        If IRMIsReady = -1 Then
'
'            txtMCCIRMPowerAmpVoltageIn = "NOT Ready"
'            Me.txtMCCIRMPowerAmpVoltageIn.BackColor = QBColor(12)
'
'        Else
'
'            txtMCCIRMPowerAmpVoltageIn = "Ready"
'            Me.txtMCCIRMPowerAmpVoltageIn.BackColor = QBColor(10)
'
'        End If
'
'    Else
'
'        'We don't use the IRM ready status
'        IRMIsReady = -1
'
'        Me.txtMCCIRMPowerAmpVoltageIn.text = ""
'
'        Me.txtMCCIRMPowerAmpVoltageIn.Enabled = False
'
'    End If
'
'End Function
'------------------------------------------------------------------------------------------------------------'
'------------------------------------------------------------------------------------------------------------'

Public Function IRMBleedBelowVoltage(ByVal voltage As Double, _
                                     ByVal deltaV As Double, _
                                     Optional ByVal CalibrationMode As Boolean = False, _
                                     Optional ByVal doZero As Boolean = False) As Boolean
    
    Dim readySignals As Integer
    Dim readVoltage As Double
    Dim OvershootVoltage As Double
    Dim MCCVoltage As Double
    Dim NewVoltage As Double
    Dim PastArray(100) As Double
    Dim i As Integer
    
    Dim PercentDone As String
    
    If voltage < 0.1 Then Exit Function
    
    readVoltage = IRMAverageVoltageIn
    
    If readVoltage < voltage Then Exit Function
    
    'Set the pastarray to the default value = 0
    For i = 0 To UBound(PastArray) - 1
    
        PastArray(i) = 0
        
    Next i
        
    'Check to see which IRM system is being used
    If IRMSystem = "Old" Then
        
        'Set charge target voltage to zero
        '(July 2010 - I Hilburn) Replaces old frmMCC with frmDAQ_Comm using the Channel object variables
        frmDAQ_Comm.DoDAQIO IRMVoltageOut, 0
        txtMCCIRMVout = "0"
        
        'Trim down charge voltage
        '(July 2010 - I Hilburn) Replaces old frmMCC with frmDAQ_Comm using the Channel object variables
        frmDAQ_Comm.DoDAQIO IRMTrim, , TrimOnOff(True)
        txtMCCIRMTrim = "On"
        Me.txtMCCIRMTrim.BackColor = QBColor(10)
        
        'Wait for charge voltage to consistently be lower than the needed voltage
        Do While readySignals < 3
            
            readVoltage = IRMAverageVoltageIn
            lblIRMStatus.Caption = "Trimming: " & Format$(100 * (readVoltage / voltage), "##0.0") & "%"
            
            PercentDone = Trim(Str(CInt(100 * readVoltage / voltage)))
            PercentDone = PadLeft(PercentDone, 6) & "%"
            
            'Update Program form
            frmProgram.StatusBar "Trimming... " & PercentDone, 3
            
            If readVoltage < voltage Then readySignals = readySignals + 1
            DelayTime 0.02
            
            'Check for a plateau - a voltage that the IRM will not trim below
            UpdateIRMPast PastArray, readVoltage
                
            If CheckIRMPast(PastArray) = True Then
                
                readySignals = 3
                
                'Failed trim - couldn't reach target
                IRMBleedBelowVoltage = False
                
                'Turn Trim off
                '(July 2010 - I Hilburn) Replaces old frmMCC with frmDAQ_Comm using the Channel object variables
                frmDAQ_Comm.DoDAQIO IRMTrim, , TrimOnOff(False)
                txtMCCIRMTrim = "Off"
                Me.txtMCCIRMTrim.BackColor = QBColor(12)
                
                Exit Function
                
            End If
            
            'Same for either system
            If IRMInterrupt Then
                
                IRMInterrupt = False
                
                lblIRMStatus.Caption = "Trim Interrupted"
                                                               
                'Update program status bar
                frmProgram.StatusBar "Trimming... Interrupted!", 3
                
                'Pause 2 seconds
                PauseTill timeGetTime() + 2000
                
                'Wipe the status bars clean
                frmProgram.StatusBar vbNullString, 2
                frmProgram.StatusBar vbNullString, 3
                
                'failed trim, user interrupted it, need to fire after this
                'function exits
                IRMBleedBelowVoltage = False
                
                'Turn Trim off
                '(July 2010 - I Hilburn) Replaces old frmMCC with frmDAQ_Comm using the Channel object variables
                frmDAQ_Comm.DoDAQIO IRMTrim, , TrimOnOff(False)
                txtMCCIRMTrim = "Off"
                Me.txtMCCIRMTrim.BackColor = QBColor(12)
                
                Exit Function
            
            End If
            
        Loop
        
        'Turn Trim off
        '(July 2010 - I Hilburn) Replaces old frmMCC with frmDAQ_Comm using the Channel object variables
        frmDAQ_Comm.DoDAQIO IRMTrim, , TrimOnOff(False)
        txtMCCIRMTrim = "Off"
        Me.txtMCCIRMTrim.BackColor = QBColor(12)
    
    Else
    
        '(July 2, 2011 - I Hilburn)
        'Set charge target voltage to zero - to prevent the IRM system from continuing to charge
        'while trimming the voltage
        '(July 2010 - I Hilburn)
        'Replaces old frmMCC with frmDAQ_Comm using the Channel object variables
        frmDAQ_Comm.DoDAQIO IRMVoltageOut, 0
        txtMCCIRMVout = "0"
            
        'Trim down charge voltage
        '(July 2010 - I Hilburn) Replaces old frmMCC with frmDAQ_Comm using the Channel object variables
        frmDAQ_Comm.DoDAQIO IRMTrim, , TrimOnOff(True)
        txtMCCIRMTrim = "On"
        Me.txtMCCIRMTrim.BackColor = QBColor(10)
        
        'Set readysignal = 0
        readySignals = 0
        
        Do While readySignals < 3
            
            'Get the current charge voltage
            readVoltage = IRMAverageVoltageIn
            
            'Display the charge voltage to the IRM/ARM form
            lblIRMStatus.Caption = "Trimming: " & Format$(100 * (readVoltage / voltage), "##0.0") & "%"
            
            PercentDone = Format(100 * readVoltage / voltage, "##0.0")
            PercentDone = PadLeft(PercentDone, 8) & "%"
            
            'Read in the poweramp voltage
            Me.txtMCCIRMPowerAmpVoltageIn = Trim(Str(CDbl(frmDAQ_Comm.DoDAQIO(IRMPowerAmpVoltageIn))))
            
            'Update Program form
            frmProgram.StatusBar "Trimming... " & PercentDone, 3
            
            'If we've reached within 0.005 of the target voltage, then exit the trim immediately
            If Abs(readVoltage - voltage) < deltaV Or _
               readVoltage < voltage _
            Then
                readySignals = 4
            End If
            
            'Don't proceed with the rest of the code if the trim has been determined to be done
            If readySignals < 3 Then
                'Check for a plateau - a voltage that the IRM will not trim below
                UpdateIRMPast PastArray, readVoltage
                    
                If CheckIRMPast(PastArray) = True Then
                
                    readySignals = 3
                    
                    'Failed trim - couldn't reach target
                    IRMBleedBelowVoltage = False
                    
                    'Turn Trim off
                    '(July 2010 - I Hilburn) Replaces old frmMCC with frmDAQ_Comm using the Channel object variables
                    frmDAQ_Comm.DoDAQIO IRMTrim, , TrimOnOff(False)
                    txtMCCIRMTrim = "Off"
                    Me.txtMCCIRMTrim.BackColor = QBColor(12)
                    
                    Exit Function
                    
                End If
                
                'Pause 100 ms between loops
                DelayTime 0.01
                
            End If
                                           
            'Same for either system
            If IRMInterrupt Then
                
                IRMInterrupt = False
                
                lblIRMStatus.Caption = "Charge Interrupted"
                                                               
                'Update program status bar
                frmProgram.StatusBar "Charging... Interrupted!", 3
                
                'Pause 2 seconds
                PauseTill timeGetTime() + 2000
                
                'Wipe the status bars clean
                frmProgram.StatusBar vbNullString, 2
                frmProgram.StatusBar vbNullString, 3
                
                'Failed trim - couldn't reach target
                IRMBleedBelowVoltage = False
                
                'Turn Trim off
                '(July 2010 - I Hilburn) Replaces old frmMCC with frmDAQ_Comm using the Channel object variables
                frmDAQ_Comm.DoDAQIO IRMTrim, , TrimOnOff(False)
                txtMCCIRMTrim = "Off"
                Me.txtMCCIRMTrim.BackColor = QBColor(12)
                
                Exit Function
                
            End If

            
        Loop
    
        'Turn Trim off
        '(July 2010 - I Hilburn) Replaces old frmMCC with frmDAQ_Comm using the Channel object variables
        frmDAQ_Comm.DoDAQIO IRMTrim, , TrimOnOff(False)
        txtMCCIRMTrim = "Off"
        Me.txtMCCIRMTrim.BackColor = QBColor(12)
    
    End If
    
    IRMBleedBelowVoltage = True
    
    lblIRMStatus.Caption = vbNullString
    
End Function

Public Function IRMCapacitorVoltage() As Double
        
    '(July 2010 - I Hilburn) Replaces old frmMCC with frmDAQ_Comm using the Channel object variables
    IRMCapacitorVoltage = CDbl(frmDAQ_Comm.DoDAQIO(IRMCapacitorVoltageIn))
    
    If IRMCapacitorVoltage < 0 Then IRMCapacitorVoltage = 0
    
    txtMCCIRMVin = Format(IRMCapacitorVoltage, "#0.0#####")

End Function

Public Function IRMCenteringPos(Optional ByVal field As Double = 0) As Long
    
    IRMCenteringPos = IRMPos
    
End Function

Public Sub IRMInterruptCharge()
    IRMInterrupt = True
End Sub

Public Sub optCoil_Click(Index As Integer)

    'Setting this radio button does not toggle the AF/IRM relays
    'That is done within the IRM fire code, itself
    
    'This does set the active af coil system
    
    'This event is deactivated if CoilsLocked = false
    'or if this event was triggered by a code change of value
    'rather than a user click
    If CoilsLocked = True Then Exit Sub
    
    If optCoil(1).value = True And _
       EnableTransIRM = False And _
       EnableAxialIRM = True Then
       
       optCoil(1).value = False
       optCoil(0).value = True
       
    ElseIf optCoil(0).value = True And _
           EnableAxialIRM = False And _
           EnableTransIRM = True Then
           
        optCoil(0).value = False
        optCoil(1).value = True
        
    ElseIf EnableAxialIRM = False And _
           EnableTransIRM = False Then
           
        optCoil(0).value = False
        optCoil(1).value = False
       
    End If

    If Index = 0 And _
       optCoil(Index).value = True _
    Then
       
        ActiveCoilSystem = AxialCoilSystem
                
    ElseIf Index = 1 And _
       optCoil(Index).value = True _
    Then
      
        ActiveCoilSystem = TransverseCoilSystem
                
    Else
    
        ActiveCoilSystem = NoCoilSystem
        
    End If

End Sub

Private Function ScaleUp(ByRef voltage As Double) As Double
    
    'If voltage is between 0 - 20, then add 50% to the value
    If voltage <= 20 Then
        voltage = voltage * 1.5
    ElseIf voltage > 20 And voltage <= 50 Then
    'If voltage is between 20 and 50, then add 33% to the value
        voltage = voltage * 1.33
    ElseIf voltage > 50 And voltage <= 100 Then
    'If voltage is between 50 and 100, then add 20% to the value
        voltage = voltage * 1.2
    ElseIf voltage > 100 And voltage <= 200 Then
    'If voltage is between 100 and 200, then add 10% to the value
        voltage = voltage * 1.1
    ElseIf voltage > 200 And voltage <= 350 Then
    'if voltage is between 200 and 350, then add 2% to the value
        voltage = voltage * 1.02
    Else
    'if voltage > 350, add 1% to the value
        voltage = voltage * 1.01
    End If
    
    ScaleUp = voltage

End Function

Public Sub SetBiasField(ByVal Gauss As Double)
    
    Dim i As Integer
    Dim voltage As Double
    
    'If ARM is not enabled in the Modules settings, then don't
    'let the code set an ARM Bias voltage!
    If Not EnableARM Then Exit Sub
    
    'Log action in the debug form
    If DEBUG_MODE Then frmDebug.msg "Set bias field " & Str$(Gauss) & " G"
    
    'Get the current bias field level
    CurrentBiasField = Gauss
    
    'Convert that field value to voltage using the ARM calibration value
    voltage = ARMVoltGauss * Gauss
    
    'No negative voltages allowed
    If voltage < 0 Then voltage = 0
    
    'Coerce voltage to be less than the ARM max voltage
    If voltage > ARMVoltMax Then voltage = ARMVoltMax
    
    'Display the new bias voltage
    txtBiasField = voltage / ARMVoltGauss
    
    'Zero the ARM Bias voltage on the ARM Box
    '(July 2010 - I Hilburn) Replaces old frmMCC with frmDAQ_Comm using the Channel object variables
    frmDAQ_Comm.DoDAQIO ARMVoltageOut, 0
    txtMCCARMVout = "0"
    
    'Wait 0.5 seconds
    DelayTime 0.5
    
    '?????? - why is this If statement here?
    If Gauss > 0 Then
    
        'Close the TTL switch to connect the ARM box
        'to the ARM Coils
        '(July 2010 - I Hilburn) Replaces old frmMCC with frmDAQ_Comm using the Channel object variables
        frmDAQ_Comm.DoDAQIO ARMSet, , False
        txtMCCARMSet = "On"
        txtMCCARMSet.BackColor = QBColor(10)
                
        'Wait another half second
        DelayTime 0.5
        
        'Set the ARM bias voltage to the desired target voltage
        '(July 2010 - I Hilburn) Replaces old frmMCC with frmDAQ_Comm using the Channel object variables
        frmDAQ_Comm.DoDAQIO ARMVoltageOut, voltage
        txtMCCARMVout = Str$(voltage)
        
        'Start the clock on tracking how long the ARM bias voltage
        'has been on
        ARMStartTime = Now
        
        '(October 2010 - I Hilburn)
        'Wait 1 second to allow the ARM charge to settle
        DelayTime 1
        
    Else
        
        'Re-zero the ARM Bias voltage
        '(July 2010 - I Hilburn) Replaces old frmMCC with frmDAQ_Comm using the Channel object variables
        frmDAQ_Comm.DoDAQIO ARMVoltageOut, 0
        txtMCCARMVout = "0"
        
        'Wait 0.5 seconds
        DelayTime 0.5
        
        'Open the TTL switch to disconnect the ARM Box
        'fromt the ARM coils
        '(July 2010 - I Hilburn) Replaces old frmMCC with frmDAQ_Comm using the Channel object variables
        frmDAQ_Comm.DoDAQIO ARMSet, , True
        txtMCCARMSet = "Off"
        txtMCCARMSet.BackColor = QBColor(12)
        
    End If
    
End Sub

Public Sub SetIRMBackFieldMode(enabling As Boolean)
    If Not EnableIRMBackfield Then enabling = False
    If enabling Then
        IRMBackfieldMode = True
        chkBackfield.value = Checked
    Else
        IRMBackfieldMode = False
        chkBackfield.value = Unchecked
    End If
End Sub

Public Function SetRelaysForIRM() As Boolean
                                                      
    'Function changes the AF/IRM relays depending on the IRM system
    'that the user inputs and returns a success status
    'True = new coil set successfully
    'False = failure to set new coil or NoIRMCoilSystem passed
    'in as the IRM system set value
                                                    
                                                      
    'Check to see what AF System is in place
    If AFSystem = "2G" Then
    
        'Connect to the 2G box
        frmAF_2G.Connect
        
        'If Backfield mode is on, and Backfield is enabled
        'set IRM to backfield axis
        If IRMBackfieldMode = True And _
           modConfig.EnableIRMBackfield = True _
        Then
        
            frmAF_2G.ConfigureCoil modConfig.IRMBackfieldAxis
            
            SetRelaysForIRM = True
                        
        ElseIf IRMBackfieldMode = False And _
               (modConfig.EnableAxialIRM = True Or _
                modConfig.EnableTransIRM = True) _
        Then
        
            frmAF_2G.ConfigureCoil modConfig.IRMAxis
            
            SetRelaysForIRM = True
            
        Else
        
            'If the code gets to this
            'bit, then the user has disabled
            'the IRM system and no IRM ramp should be
            'run
            SetRelaysForIRM = False
            
        End If
                
    ElseIf AFSystem = "ADWIN" Then
    
        'Need to figure out which Relays need to be set
        'First, see if this is a backfield or normal IRM
        
        If modConfig.EnableIRMBackfield And _
           IRMBackfieldMode = True _
        Then
            
            'Need to switch IRM relay into
            'the "Down" position
            'and for good measure, put all the relays
            'into the down position
            With SystemBoards(IRMRelay.BoardName)
            
                .DigitalOut_ADWIN (0)
                
            End With
            
        ElseIf IRMBackfieldMode = False And _
               (modConfig.EnableAxialIRM = True Or _
                modConfig.EnableTransIRM = True) _
        Then
        
            'Need to set the IRM relay into the
            '"Up" position and set all the other
            'relays into the down position
            With SystemBoards(IRMRelay.BoardName)
            
                .DigitalOutput IRMRelay, _
                            True, _
                            True
                            
            End With
            
        Else
        
            'No IRM system is enabled
            'Do nothing, but return a false
            'so that the IRM ramp ends
            SetRelaysForIRM = False
            
            Exit Function
            
        End If
        
        'Now need to make sure the Axial
        'and Transverse coils are in the right position
        If IsIRMAxialCoilSelected And _
           modConfig.EnableAxialIRM = True _
        Then
        
            'Doing an Axial IRM
            'Config needed:
            '     Axial Relay - "Down"
            'Transverse Relay - "Up"
            
            'Axial and Trans relays are
            'already both in the down position
            'from the IRM relay code above
            
            'Just need to raise Trans relay to the "Up" position
            With SystemBoards(AFTransRelay.BoardName)
            
                'Raise the AF Trans relay without
                'effecting the other relays
                .DigitalOutput AFTransRelay, _
                            True, _
                            False
                            
            End With
        
            SetRelaysForIRM = True
            
        ElseIf IsIRMTransverseCoilSelected And _
               modConfig.EnableTransIRM = True _
        Then
        
            'Doing a Transverse IRM
            'Config needed:
            '     Axial Relay - "Up"
            'Transverse Relay - "Down"
            
            'Axial and Trans relays are
            'already both in the down position
            'from the IRM relay code above
            
            'Just need to raise Axial relay to the "Up" position
            With SystemBoards(AFAxialRelay.BoardName)
            
                'Raise the AF Trans relay without
                'effecting the other relays
                .DigitalOutput AFAxialRelay, _
                            True, _
                            False
                            
            End With
        
            SetRelaysForIRM = True
            
        Else
        
            'Needed IRM system is not enabled
            'return false
            SetRelaysForIRM = False
            
        End If
        
    End If
                       
End Function


Private Function IsIRMAxialCoilSelected() As Boolean

    Dim index_of_axial As Integer
    index_of_axial = 0

    If ActiveCoilSystem = AxialCoilSystem Then
        Me.optCoil(index_of_axial).value = True
    End If

    If Me.optCoil(index_of_axial).value = True Then
        IsIRMAxialCoilSelected = True
        ActiveCoilSystem = AxialCoilSystem
    Else
        IsIRMAxialCoilSelected = False
    End If

End Function

Private Function IsIRMTransverseCoilSelected() As Boolean

    Dim index_of_transverse As Integer
    index_of_transverse = 1

    If Me.optCoil(index_of_transverse).value = True Then
        IsIRMTransverseCoilSelected = True
    Else
        IsIRMTransverseCoilSelected = False
    End If

End Function


Private Sub tmrARMWatch_Timer()
    ' Check and see if Bias field has been on too long
    If ((Now - ARMStartTime) > ARMTimeMax) And (CurrentBiasField > 0) Then
        SetBiasField 0
    End If
End Sub

Public Function TrimOnOff(ByVal TrimOn As Boolean) As Boolean

    'If the user has set that the IRM trim is wired such that it is
    'turned on by passing a
    'logic high state to the DAQ comm board IRM Trim DO channel, then
    'return the TrimOn input variable as it is
    'However, if the user has set that the IRM Trim is turned on by
    'passing a logic low state to the IRM Trim channel, then need to
    'return the logical oposite of TrimOn
    If modConfig.TrimOnTrue = True Then TrimOnOff = TrimOn
    If modConfig.TrimOnTrue = False Then TrimOnOff = Not TrimOn
    'was TrimOnTrue

'(Feb 2011, I. Hilburn)
'This case statement is no longer necessary.  A TrimOnTrue setting
'has been added to the INI file and a control added to change this value
'in frmSettings.  For each system, the IRM Trim On logic state can be assigned
'to True (hi logic state) or False (low logic state)
'
'    Select Case modConfig.IRMSystem
'
'        Case "Old"
'
'            'For the Old system:
'            '   True = Trim On
'            '   False = Trim Off
'            TrimOnOff = TrimOn
'
'        Case "Matsusada"
'
'            'For the Matsusada system:
'            '   False = Trim On
'            '   True = Trim Off
'            TrimOnOff = Not TrimOn
'
'        Case Else
'
'            SetCodeLevel CodeRed
'
'            'Send error email
'            frmSendMail.MailNotification "IRM Settings Error", _
'                                         "Invalid IRM System setting!" & _
'                                         "  Current setting is: " & modConfig.IRMSystem & _
'                                         vbNewLine & vbNewLine & _
'                                         "Code execution has been paused", _
'                                         CodeRed, _
'                                         True
'
'            'Pop-up modal message box on the screen
'            MsgBox "Invalid IRM System setting!  INI settings file error!" & _
'                   vbNewLine & vbNewLine & "The code execution has been paused!", , _
'                   "IRM Settings Error"
'
'            End
'
'            'Pause the code
'            modFlow.Flow_Pause
'
'            'In the end, just assume need to return TrimOn as is
'            TrimOnOff = TrimOn
'
'    End Select

End Function

Public Sub UpdateIRMPast(ByRef PastArray() As Double, _
                         ByVal CurValue As Double)
                         
    Dim i As Long
    Dim N As Long
    
    'Get the array length
    N = UBound(PastArray)
    
    'Iterate through the array and set element i =
    'element i + 1.
    For i = 0 To N - 2
    
        PastArray(i) = PastArray(i + 1)
        
    Next i
    
    'Set the last element = CurValue
    PastArray(N - 1) = CurValue
                                                  
End Sub

