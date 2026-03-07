VERSION 5.00
Begin VB.Form frmAFTuner 
   Caption         =   "MCC AF Tuner"
   ClientHeight    =   8535
   ClientLeft      =   7200
   ClientTop       =   2325
   ClientWidth     =   10695
   DrawMode        =   1  'Blackness
   FillColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8535
   ScaleWidth      =   10695
   Begin VB.CheckBox chkDebugMode 
      Caption         =   "Debug Mode"
      Height          =   492
      Left            =   2760
      TabIndex        =   4
      Top             =   480
      Width           =   855
   End
   Begin VB.PictureBox picPickMaxVoltage 
      BackColor       =   &H00FFFFD0&
      Height          =   2000
      Left            =   4080
      ScaleHeight     =   2000
      ScaleMode       =   0  'User
      ScaleWidth      =   3000
      TabIndex        =   57
      Top             =   1200
      Width           =   3000
      Begin VB.CommandButton cmdAcceptMaxPick 
         Caption         =   "OK"
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
         Left            =   480
         TabIndex        =   59
         Top             =   1440
         Width           =   732
      End
      Begin VB.CommandButton cmdPickMaxCancel 
         Caption         =   "Cancel"
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
         Left            =   1680
         TabIndex        =   58
         Top             =   1440
         Width           =   852
      End
      Begin VB.Shape shapePickVoltageBorder 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   5
         Height          =   1956
         Left            =   0
         Top             =   0
         Width           =   2952
      End
      Begin VB.Label lblPickMaxResults 
         BackStyle       =   0  'Transparent
         Height          =   612
         Left            =   120
         TabIndex        =   61
         Top             =   840
         Width           =   2772
      End
      Begin VB.Label lblPickMaxVoltages 
         BackStyle       =   0  'Transparent
         Height          =   612
         Left            =   120
         TabIndex        =   60
         Top             =   120
         Width           =   2772
      End
   End
   Begin VB.Frame frameMaxCoilVoltages 
      Caption         =   "Axial Coil"
      Height          =   1572
      Left            =   3720
      TabIndex        =   43
      Top             =   4680
      Width           =   6855
      Begin VB.CommandButton cmdSaveAxialVolt 
         Caption         =   "Save Max Volt"
         Height          =   372
         Left            =   4200
         TabIndex        =   63
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton cmdSaveAxialFreq 
         Caption         =   "Save Freq"
         Height          =   372
         Left            =   4200
         TabIndex        =   62
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtOldMaxAxialMonitor 
         Height          =   288
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox txtNewMaxAxialMonitor 
         Height          =   288
         Left            =   1920
         TabIndex        =   22
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox txtOldAxialResFreq 
         Height          =   288
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox txtFitAxialResFreq 
         Height          =   288
         Left            =   3000
         TabIndex        =   24
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox txtFitMaxAxialRamp 
         Height          =   288
         Left            =   840
         TabIndex        =   20
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox txtOldMaxAxialRamp 
         Height          =   288
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Max Monitor           Volts"
         Height          =   375
         Left            =   1920
         TabIndex        =   54
         Top             =   240
         Width           =   975
      End
      Begin VB.Line Line16 
         BorderColor     =   &H8000000C&
         X1              =   3960
         X2              =   3960
         Y1              =   240
         Y2              =   1440
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   " Max Ramp         Volts"
         Height          =   372
         Left            =   840
         TabIndex        =   53
         Top             =   240
         Width           =   852
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   " Tuning Freq"
         Height          =   255
         Left            =   3000
         TabIndex        =   49
         Top             =   360
         Width           =   975
      End
      Begin VB.Line Line8 
         BorderColor     =   &H8000000C&
         X1              =   120
         X2              =   120
         Y1              =   600
         Y2              =   1440
      End
      Begin VB.Line Line7 
         BorderColor     =   &H8000000C&
         X1              =   720
         X2              =   3960
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Line Line6 
         BorderColor     =   &H8000000C&
         X1              =   120
         X2              =   3960
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Line Line5 
         BorderColor     =   &H8000000C&
         X1              =   120
         X2              =   3960
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Line Line4 
         BorderColor     =   &H8000000C&
         X1              =   720
         X2              =   720
         Y1              =   240
         Y2              =   1440
      End
      Begin VB.Line Line3 
         BorderColor     =   &H8000000C&
         X1              =   2880
         X2              =   2880
         Y1              =   240
         Y2              =   1440
      End
      Begin VB.Line Line2 
         BorderColor     =   &H8000000C&
         X1              =   1800
         X2              =   1800
         Y1              =   240
         Y2              =   1440
      End
      Begin VB.Label Label23 
         Caption         =   "New:"
         Height          =   252
         Left            =   240
         TabIndex        =   46
         Top             =   1080
         Width           =   492
      End
      Begin VB.Label Label24 
         Caption         =   "Old:"
         Height          =   252
         Left            =   240
         TabIndex        =   45
         Top             =   720
         Width           =   372
      End
   End
   Begin VB.Frame frameCoilResFreq 
      Caption         =   "Transverse Coil"
      Height          =   1572
      Left            =   3720
      TabIndex        =   38
      Top             =   6360
      Width           =   6855
      Begin VB.CommandButton cmdSaveTransVolt 
         Caption         =   "Save Max Volt"
         Height          =   372
         Left            =   4200
         TabIndex        =   65
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton cmdSaveTransFreq 
         Caption         =   "Save Freq"
         Height          =   372
         Left            =   4200
         TabIndex        =   64
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtOldMaxTransverseMonitor 
         Height          =   288
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox txtNewMaxTransverseMonitor 
         Height          =   288
         Left            =   1920
         TabIndex        =   28
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox txtFitMaxTransverseRamp 
         Height          =   288
         Left            =   840
         TabIndex        =   26
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox txtOldMaxTransverseRamp 
         Height          =   288
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox txtOldTransverseResFreq 
         Height          =   288
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox txtFitTransverseResFreq 
         Height          =   288
         Left            =   3000
         TabIndex        =   30
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Max Monitor           Volts"
         Height          =   375
         Left            =   1920
         TabIndex        =   56
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   " Max Ramp         Volts"
         Height          =   372
         Left            =   840
         TabIndex        =   55
         Top             =   240
         Width           =   852
      End
      Begin VB.Line Line17 
         BorderColor     =   &H8000000C&
         X1              =   3960
         X2              =   3960
         Y1              =   240
         Y2              =   1440
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   " Tuning Freq"
         Height          =   255
         Left            =   3000
         TabIndex        =   50
         Top             =   360
         Width           =   975
      End
      Begin VB.Line Line15 
         BorderColor     =   &H8000000C&
         X1              =   120
         X2              =   3960
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Line Line14 
         BorderColor     =   &H8000000C&
         X1              =   120
         X2              =   120
         Y1              =   600
         Y2              =   1440
      End
      Begin VB.Line Line13 
         BorderColor     =   &H8000000C&
         X1              =   720
         X2              =   3960
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Line Line12 
         BorderColor     =   &H8000000C&
         X1              =   120
         X2              =   3960
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Line Line11 
         BorderColor     =   &H8000000C&
         X1              =   720
         X2              =   720
         Y1              =   240
         Y2              =   1440
      End
      Begin VB.Line Line10 
         BorderColor     =   &H8000000C&
         X1              =   2880
         X2              =   2880
         Y1              =   240
         Y2              =   1440
      End
      Begin VB.Line Line9 
         BorderColor     =   &H8000000C&
         X1              =   1800
         X2              =   1800
         Y1              =   240
         Y2              =   1440
      End
      Begin VB.Label Label21 
         Caption         =   "New:"
         Height          =   252
         Left            =   240
         TabIndex        =   48
         Top             =   1080
         Width           =   492
      End
      Begin VB.Label Label22 
         Caption         =   "Old:"
         Height          =   252
         Left            =   240
         TabIndex        =   47
         Top             =   720
         Width           =   372
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Default         =   -1  'True
      Height          =   372
      Left            =   3720
      TabIndex        =   31
      Top             =   8040
      Width           =   6855
   End
   Begin VB.Frame frameClippingTest 
      Caption         =   "Auto Clipping Test"
      Height          =   3975
      Left            =   120
      TabIndex        =   35
      Top             =   4440
      Width           =   3495
      Begin VB.TextBox txtNumSineFits 
         Height          =   288
         Left            =   1800
         TabIndex        =   17
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox txtRampDownSlope 
         BackColor       =   &H8000000E&
         Height          =   288
         Left            =   1560
         TabIndex        =   16
         Top             =   3000
         Width           =   735
      End
      Begin VB.TextBox txtRampUpSlope 
         BackColor       =   &H8000000E&
         Height          =   288
         Left            =   1560
         TabIndex        =   15
         Top             =   2520
         Width           =   735
      End
      Begin VB.CommandButton cmdStartAutoClipTest 
         BackColor       =   &H0000FF00&
         Caption         =   "Start Clipping Auto-Test"
         Height          =   372
         Left            =   600
         MaskColor       =   &H00008000&
         TabIndex        =   18
         Top             =   3480
         Width           =   2172
      End
      Begin VB.TextBox txtMaxClipAmp 
         Height          =   288
         Left            =   1800
         TabIndex        =   14
         Top             =   1440
         Width           =   855
      End
      Begin VB.TextBox txtMinClippingAmp 
         Height          =   288
         Left            =   1800
         TabIndex        =   13
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox txtClippingSineFreq 
         Height          =   288
         Left            =   1800
         TabIndex        =   12
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "AF Output Ramp Voltage to Amplifier:"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   70
         Top             =   720
         Width           =   3015
      End
      Begin VB.Label lblRampDownDuration 
         Height          =   255
         Left            =   2520
         TabIndex        =   69
         Top             =   3000
         Width           =   735
      End
      Begin VB.Label lblRampUpDuration 
         Height          =   255
         Left            =   2520
         TabIndex        =   68
         Top             =   2565
         Width           =   735
      End
      Begin VB.Label Label10 
         Caption         =   "Duration (ms):"
         Height          =   255
         Left            =   2400
         TabIndex        =   67
         Top             =   2280
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Ramp Down Slope (V / sec):"
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   66
         Top             =   2880
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "# of Sine Fits:"
         Height          =   255
         Left            =   240
         TabIndex        =   52
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label25 
         Caption         =   "Ramp Up Slope (V / sec):"
         Height          =   375
         Left            =   240
         TabIndex        =   44
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label Label16 
         Caption         =   "To Voltage:"
         Height          =   375
         Left            =   600
         TabIndex        =   42
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label15 
         Caption         =   "From Voltage:"
         Height          =   375
         Left            =   600
         TabIndex        =   41
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label14 
         Caption         =   "Sine Freq (Hz):"
         Height          =   252
         Left            =   240
         TabIndex        =   40
         Top             =   360
         Width           =   1212
      End
   End
   Begin VB.PictureBox picDCResponse 
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000E&
      BeginProperty DataFormat 
         Type            =   2
         Format          =   "0.000E+00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   9
      EndProperty
      FontTransparent =   0   'False
      Height          =   4455
      Left            =   3720
      MousePointer    =   2  'Cross
      ScaleHeight     =   10000
      ScaleMode       =   0  'User
      ScaleWidth      =   14500
      TabIndex        =   32
      Top             =   120
      Width           =   6855
   End
   Begin VB.Frame frameAFAutoTune 
      Caption         =   "Auto-Tuner"
      Height          =   2892
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   3495
      Begin VB.TextBox txtFreqStepSize 
         Height          =   288
         Left            =   2280
         TabIndex        =   8
         Top             =   960
         Width           =   852
      End
      Begin VB.PictureBox picBluePixel 
         Height          =   12
         Left            =   2280
         Picture         =   "frmAFTuner.frx":0000
         ScaleHeight     =   0.027
         ScaleMode       =   0  'User
         ScaleWidth      =   0.027
         TabIndex        =   39
         Top             =   240
         Visible         =   0   'False
         Width           =   12
      End
      Begin VB.TextBox txtAmplitude 
         Height          =   288
         Left            =   2280
         TabIndex        =   10
         Top             =   1920
         Width           =   852
      End
      Begin VB.TextBox txtDuration 
         Height          =   288
         Left            =   2280
         TabIndex        =   9
         Top             =   1440
         Width           =   852
      End
      Begin VB.CommandButton cmdAutoTuneAF 
         Caption         =   "Start Auto-Tune"
         Height          =   372
         Left            =   600
         TabIndex        =   11
         Top             =   2400
         Width           =   2292
      End
      Begin VB.TextBox txtHighFreq 
         Height          =   288
         Left            =   2040
         TabIndex        =   7
         Top             =   480
         Width           =   1092
      End
      Begin VB.TextBox txtLowFreq 
         Height          =   288
         Left            =   360
         TabIndex        =   6
         Top             =   480
         Width           =   1092
      End
      Begin VB.Label Label3 
         Caption         =   "Frequency Step Size (Hz):"
         Height          =   255
         Left            =   360
         TabIndex        =   51
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label lblAmplitude 
         Caption         =   "Amplitude (0 - 10 volts):"
         Height          =   375
         Left            =   360
         TabIndex        =   37
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Label lblScanDuration 
         Caption         =   "Peak Hange Time at Each Frequency (ms):"
         Height          =   495
         Left            =   360
         TabIndex        =   36
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label lblHighFreq 
         Caption         =   "Highest Freq(Hz):"
         Height          =   255
         Left            =   2040
         TabIndex        =   34
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblLowFreq 
         Caption         =   "Lowest-Freq (Hz):"
         Height          =   255
         Left            =   360
         TabIndex        =   33
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Coil"
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   2532
      Begin VB.CheckBox chkLockCoils 
         Caption         =   "Lock coil selection"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   840
         Width           =   1695
      End
      Begin VB.OptionButton optCoil 
         Caption         =   "Transverse"
         Height          =   252
         Index           =   1
         Left            =   1200
         TabIndex        =   2
         Top             =   360
         Width           =   1212
      End
      Begin VB.OptionButton optCoil 
         Caption         =   "Axial"
         Height          =   252
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Value           =   -1  'True
         Width           =   852
      End
   End
End
Attribute VB_Name = "frmAFTuner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Real!()
Dim Imag!()

'These need to be global so that button click event actions can access them
Dim MaxRMS As Double
Dim MinRMS As Double
Dim MaxAmp As Double
Dim MaxValueMinusLabel_RMS As Double
Dim MaxValueMinusLabel_Amp As Double
Dim MinValueMinusLabel_RMS As Double
Dim AmpInterval As Double
Dim RMSInterval As Double
Dim VoltRange As Double
Dim SelectedRampV As Double
Dim SelectedMonitorV As Double

Dim AbortAutoTune As Boolean

'In-form global declaration to hold the string identifier of the AF coil system
Dim CoilString As String

Dim VSelectMode As String

Private Sub ApplyCoilParam(ByVal CoilParam As SaveCoilParam, _
                           Optional ByVal AFCoilSystem As Integer = -128)

    Dim CoilString As String
    Dim ParamString As String
    Dim UserResp As Long

    'Fill in coil system if the user has not
    If AFCoilSystem = -128 Then AFCoilSystem = ActiveCoilSystem
    
    'Get the coil string to use
    If AFCoilSystem = AxialCoilSystem Then CoilString = "Axial"
    If AFCoilSystem = TransverseCoilSystem Then CoilString = "Transverse"
    If CoilParam = resFreq Then ParamString = "resonance frequency"
    If CoilParam = VoltsMax Then ParamString = "max voltages"
    
    'Prompt the user to see if they really want to go through with this change
    UserResp = MsgBox("Are you sure that you want to change the AF " & _
                      CoilString & " coil " & ParamString & "?  The wrong value " & _
                      "could result in coil overheating and damage.", _
                      vbYesNo, _
                      "Warning!")
                
    'Check for a "No" reply
    If UserResp <> vbYes Then
    
        'User does not want to make the change
        Exit Sub
        
    End If
    
    'Apply the Axial frequency
    If AFCoilSystem = AxialCoilSystem Then
    
        If CoilParam = resFreq Then
        
            'Check to see if the desired parameter text-box is empty
            If val(Me.txtFitAxialResFreq) = 0 Then Exit Sub
            
            'Update the display
            Me.txtOldAxialResFreq = Me.txtFitAxialResFreq
            
            'Update the global variable
            modConfig.AfAxialResFreq = val(Me.txtOldAxialResFreq)
        
        ElseIf CoilParam = VoltsMax Then
        
            'Check to see if the desired parameter text-box is empty
            If val(Me.txtFitMaxAxialRamp) <> 0 Then
        
                'Update the display
                Me.txtOldMaxAxialRamp = Me.txtFitMaxAxialRamp
                
                'Update the global variable
                modConfig.AfAxialRampMax = val(Me.txtOldMaxAxialRamp)
                
                'Update label display in frmSettings
                frmSettings.lblAFAxialRampMax = Me.txtOldMaxAxialRamp
                
            End If
        
            'Check to see if the desired parameter text-box is empty
            If val(Me.txtNewMaxAxialMonitor) <> 0 Then
        
                'Update the display
                Me.txtOldMaxAxialMonitor = Me.txtNewMaxAxialMonitor
                
                'Update the global variable
                modConfig.AfAxialMonMax = val(Me.txtOldMaxAxialMonitor)
                
            End If
                
        End If
                
    ElseIf AFCoilSystem = TransverseCoilSystem Then
    
        If CoilParam = resFreq Then
        
            'Check to see if the desired parameter text-box is empty
            If val(Me.txtFitTransverseResFreq) = 0 Then Exit Sub
            
            'Update the display
            Me.txtOldTransverseResFreq = Me.txtFitTransverseResFreq
            
            'Update the global variable
            modConfig.AfTransResFreq = val(Me.txtOldTransverseResFreq)
            
            
        ElseIf CoilParam = VoltsMax Then
        
            'Check to see if the desired parameter text-box is empty
            If val(Me.txtFitMaxTransverseRamp) <> 0 Then
        
                'Update the display
                Me.txtOldMaxTransverseRamp = Me.txtFitMaxTransverseRamp
                
                'Update the global variable
                modConfig.AfTransRampMax = val(Me.txtOldMaxTransverseRamp)
                
                'Update label display in frmSettings
                frmSettings.lblAFTransverseRampMax.Caption = Me.txtOldMaxTransverseRamp
                
            End If
        
            'Check to see if the desired parameter text-box is empty
            If val(Me.txtNewMaxTransverseMonitor) <> 0 Then
        
                'Update the display
                Me.txtOldMaxTransverseMonitor = Me.txtNewMaxTransverseMonitor
                
                'Update the global variable
                modConfig.AfTransMonMax = val(Me.txtOldMaxTransverseMonitor)
                
            End If
                
        End If
                
    End If
        
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

Private Sub cmdAcceptMaxPick_Click()

    'Need to write the accepted values to the correct text boxes
    If ActiveCoilSystem = AxialCoilSystem Then
    
        Me.txtFitMaxAxialRamp = Format(SelectedRampV, "0.0#")
        Me.txtNewMaxAxialMonitor = Format(SelectedMonitorV, "0.0#")
        
    ElseIf ActiveCoilSystem = TransverseCoilSystem Then
    
    
        Me.txtFitMaxTransverseRamp = Format(SelectedRampV, "0.0#")
        Me.txtNewMaxTransverseMonitor = Format(SelectedMonitorV, "0.0#")
    
    End If
    
    'Now hide & clear the max voltage picking picture box, etc.
    Me.picPickMaxVoltage.Visible = False
    Me.lblPickMaxResults.Visible = False
    Me.lblPickMaxVoltages.Visible = False
    Me.cmdAcceptMaxPick.Visible = False
    Me.cmdPickMaxCancel.Visible = False
        
    Me.picPickMaxVoltage.Cls
    
    'Refresh the form
    Me.refresh
    
    'Reset the program form status bar 3rd panel caption
    frmProgram.StatusBar vbNullString, 3
    
End Sub

Private Sub cmdApplyAxialFreq_Click()

    ApplyCoilParam resFreq, AxialCoilSystem

End Sub

Private Sub cmdApplyAxialVolt_Click()

    ApplyCoilParam VoltsMax, AxialCoilSystem

End Sub

Private Sub cmdApplyTransFreq_Click()

    ApplyCoilParam resFreq, TransverseCoilSystem

End Sub

Private Sub cmdApplyTransVolt_Click()

    ApplyCoilParam VoltsMax, TransverseCoilSystem

End Sub

Private Sub cmdAutoTuneAF_Click()
    
    If Me.cmdAutoTuneAF.Caption = "Start Auto-Tune" Then
    
        Me.cmdAutoTuneAF.Caption = "Stop Auto-Tune"
    
        AbortAutoTune = False
    
        SetupAFAutoTune
        
        Me.cmdAutoTuneAF.Caption = "Start Auto-Tune"
        
    Else
        
        AbortAutoTune = True
        
        Me.cmdAutoTuneAF.Caption = "Aborting Auto-Tune..."
                   
    End If
                                                                 
End Sub

Private Sub cmdClose_Click()

    Me.Hide
            
End Sub

Private Sub cmdPickMaxCancel_Click()

    'Now hide & clear the max voltage picking picture box, etc.
    Me.picPickMaxVoltage.Visible = False
    Me.lblPickMaxResults.Visible = False
    Me.lblPickMaxVoltages.Visible = False
    Me.cmdAcceptMaxPick.Visible = False
    Me.cmdPickMaxCancel.Visible = False
        
    Me.picPickMaxVoltage.Cls
    
    'Refresh the form
    Me.refresh
    
    'Reset the program form status bar 3rd panel caption
    frmProgram.StatusBar vbNullString, 3

End Sub

Private Sub cmdSaveAxialFreq_Click()

    If val(Me.txtFitAxialResFreq) > 0 Then
    
        'Apply the new resonance freq
        ApplyCoilParam resFreq, AxialCoilSystem
        
        'Now save the resonance frequency to the INI file
        modConfig.SaveCoilTuningParam resFreq, _
                                      AxialCoilSystem
                                      
        Me.txtOldAxialResFreq = Me.txtFitAxialResFreq
        
    Else
    
        'Tell the user that negative res freq are not allowed
        MsgBox "Negative or Zero AF Coil resonance frequencies are both non-sensical " & _
               "and not allowed.  (So there!)", , _
               "Ooops!"
               
        Me.txtFitAxialResFreq = ""
        
    End If

End Sub

Private Sub cmdSaveAxialVolt_Click()

    If val(Me.txtNewMaxAxialMonitor) > 0 Or _
       val(Me.txtFitMaxAxialRamp) > 0 _
    Then
    
        'Apply the new max voltages
        ApplyCoilParam VoltsMax, AxialCoilSystem
        
        'Now save the max voltages to the INI file
        modConfig.SaveCoilTuningParam VoltsMax, _
                                      AxialCoilSystem
                                      
        If val(Me.txtNewMaxAxialMonitor) > 0 Then Me.txtOldMaxAxialMonitor = Me.txtNewMaxAxialMonitor
        If val(Me.txtFitMaxAxialRamp) > 0 Then Me.txtOldMaxAxialRamp = Me.txtFitMaxAxialRamp
        
    Else
    
        'Tell the user that negative res freq are not allowed
        MsgBox "Negative or Zero AF Coil max voltages are both non-sensical " & _
               "and not allowed.  (So there!)", , _
               "Ooops!"
        
    End If

End Sub

Private Sub cmdSaveTransFreq_Click()

    If val(Me.txtFitTransverseResFreq) > 0 Then
    
        'Apply the new resonance freq
        ApplyCoilParam resFreq, TransverseCoilSystem
        
        'Now save the resonance frequency to the INI file
        modConfig.SaveCoilTuningParam resFreq, _
                                      TransverseCoilSystem
                                      
        Me.txtOldTransverseResFreq = Me.txtFitTransverseResFreq
        
    Else
    
        'Tell the user that negative res freq are not allowed
        MsgBox "Negative or Zero AF Coil resonance frequencies are both non-sensical " & _
               "and not allowed.  (So there!)", , _
               "Ooops!"
               
        Me.txtFitTransverseResFreq = ""
        
    End If
    
End Sub

Private Sub cmdSaveTransVolt_Click()

    If val(Me.txtNewMaxTransverseMonitor) > 0 And _
       val(Me.txtFitMaxTransverseRamp) > 0 _
    Then
    
        'Apply the new resonance freq
        ApplyCoilParam VoltsMax, TransverseCoilSystem
        
        'Now save the resonance frequency to the INI file
        modConfig.SaveCoilTuningParam VoltsMax, _
                                      TransverseCoilSystem
        
        Me.txtOldMaxTransverseMonitor = Me.txtNewMaxTransverseMonitor
        Me.txtOldMaxTransverseRamp = Me.txtFitMaxTransverseRamp
                
    Else
    
        'Tell the user that negative res freq are not allowed
        MsgBox "Negative or Zero AF Coil max voltages are both non-sensical " & _
               "and not allowed.  (So there!)", , _
               "Ooops!"
                 
    End If

End Sub

Private Sub cmdStartAutoClipTest_Click()

    Dim AFData() As Double
    Dim SineFit_Data() As Double
    
    Dim CurTime
    Dim FolderName As String
    Dim CoilString As String
    
    'Nothing's been initialized, do so now
    If SystemBoards Is Nothing Or WaveForms Is Nothing Then
    
        'Raise an Error
        Err.Raise -616, _
                  "frmAFTuner.cmbStartAutoClipTest", _
                  "AF System wave-forms and/or the System Boards collection have not been loaded " & _
                  "properly.  Please check the Paleomag.ini file." & vbNewLine & vbNewLine & _
                  "The code will end now."
                  
        End
        
    End If
    
    'Lock the coils
    CoilsLocked = True
    Me.chkLockCoils.value = Checked
    
    'Set the AF coil string descriptor
    If ActiveCoilSystem = AxialCoilSystem Then CoilString = "Axial"
    If ActiveCoilSystem = TransverseCoilSystem Then CoilString = "Transverse"
    
    'Update the program form 2nd status bar panel
    frmProgram.StatusBar "AF " & CoilString & " Auto-Clip Test Config", 2
    
    'Now change individual settings on each Wave object
    With WaveForms("AFMONITOR")
    
        .PeakVoltage = val(Me.txtMaxClipAmp)
        .SineFreqMin = val(Me.txtClippingSineFreq)
        Set .range = New range
        .range.MaxValue = 10
        .range.MinValue = -10
    
    End With
    
    With WaveForms("AFRAMPUP")
    
        .PeakVoltage = val(Me.txtMaxClipAmp)
        .MinVoltage = val(Me.txtMinClippingAmp)
        .SineFreqMin = WaveForms("AFMONITOR").SineFreqMin
        .Slope = val(Me.txtRampUpSlope)
        Set .range = WaveForms("AFMONITOR").range
        
    End With
    
    With WaveForms("AFRAMPDOWN")
    
        .PeakVoltage = WaveForms("AFRAMPUP").PeakVoltage
        .MinVoltage = WaveForms("AFRAMPUP").MinVoltage
        .SineFreqMin = WaveForms("AFMONITOR").SineFreqMin
        .Slope = val(Me.txtRampDownSlope)
        Set .range = WaveForms("AFMONITOR").range
        
    End With
    
    'Clear the picture results box and post the status of the clipping test
    picDCResponse.Cls
    
    'Change Font size and text display start point
    picDCResponse.FontSize = 10
    picDCResponse.CurrentX = 5000
    picDCResponse.CurrentY = 4000
    
    'Print Clipping Test Ramp status
    picDCResponse.Print CoilString & " coil auto-clipping test:"
    
    'Change text display start point again
    picDCResponse.CurrentX = 6000
    picDCResponse.CurrentY = 4000 + 1.5 * picDCResponse.TextHeight(CoilString)
    picDCResponse.Print "Ramping..."
    
    'Refresh form
    Me.refresh
    
    'Change radio button to an uncalibrated Ramp on the ADWIN AF form
    frmADWIN_AF.optCalRamp(1).value = True
    frmADWIN_AF.optCalRamp(0).value = False
    
    'Update the program form 2nd status bar panel
    frmProgram.StatusBar "AF " & CoilString & " Auto-Clip Test", 2
    
    'Do the Clipping test ramp
    frmADWIN_AF.DoRampADWIN_WithParameterLogging WaveForms("AFMONITOR"), _
                                                 WaveForms("AFRAMPUP"), _
                                                 WaveForms("AFRAMPDOWN"), _
                                                 AFData(), _
                                                 1, _
                                                 0, _
                                                 3
                                
    'Update Ramping Status
    picDCResponse.CurrentX = 6000
    picDCResponse.CurrentY = 4000 + 1.5 * picDCResponse.TextHeight(CoilString)
    picDCResponse.Print "Ramping.... Done"
    
    'Change the text cursor position
    picDCResponse.CurrentX = 6000
    picDCResponse.CurrentY = 4000 + 3 * picDCResponse.TextHeight(CoilString)
    picDCResponse.Print "Analyzing..."
    
    'Refresh form
    Me.refresh
    
    'Do the Sine-Fitting analysis now
    frmADWIN_AF.DoSineFitAnalysis WaveForms("AFMONITOR"), _
                                  AFData, _
                                  SineFit_Data, _
                                  1, _
                                  UBound(AFData, 1) \ Int(val(Me.txtNumSineFits))
                                      
    'Change the text cursor position
    picDCResponse.CurrentX = 6000
    picDCResponse.CurrentY = 4000 + 3 * picDCResponse.TextHeight(CoilString)
    picDCResponse.Print "Analyzing.... Done"
    
    'Refresh form
    Me.refresh
    
    'Pause a half a second (500 ms) for the user to see this
    PauseTill timeGetTime() + 500
                                      
    'Check to see if the user has selected Debug Mode
    If Me.chkDebugMode.value = Checked Then
    
        'Change the text cursor position
        picDCResponse.CurrentX = 6000
        picDCResponse.CurrentY = 4000 + 4.5 * picDCResponse.TextHeight(CoilString)
        picDCResponse.Print "Saving Data..."
        
        'Refresh form
        Me.refresh
        
        'Set the CurTime
        CurTime = Now
        
        'Create the Folder Name
        FolderName = CoilString & " Clip " & Trim(str(val(Me.txtMaxClipAmp))) & "V - " & _
                     Format(CurTime, "MM-DD-YY, HH MM DD") & "/"
        
        'Save the Data
        frmFileSave.MultiRampFileSave_ADWIN AFData, _
                                            WaveForms("AFMONITOR").TimeStep, _
                                            1048000, _
                                            FolderName, _
                                            CurTime, _
                                            SineFit_Data, _
                                            True, _
                                            True, _
                                            WaveForms("AFMONITOR").PtsPerPeriod
                      
        'Change the text cursor position
        picDCResponse.CurrentX = 6000
        picDCResponse.CurrentY = 4000 + 4.5 * picDCResponse.TextHeight(CoilString)
        picDCResponse.Print "Saving Data.... Done"
        
        'Refresh form
        Me.refresh
                      
    End If
    
    'Pause a half a second (500 ms) for the user to see this
    PauseTill timeGetTime() + 500
                           
    'Plot the points now
    PlotAutoClipTestResults SineFit_Data, _
                            Me.picDCResponse, _
                            val(Me.txtMinClippingAmp), _
                            WaveForms("AFRAMPUP").PeakVoltage, _
                            WaveForms("AFRAMPUP").CurrentPoint + 1
    
    'Reset the 2nd status bar panel
    frmProgram.StatusBar vbNullString, 2
    
    'Unlock the coils
    CoilsLocked = False
    Me.chkLockCoils.value = Unchecked
    
End Sub

Private Sub doAFAutoTune(ByRef AFData() As Double, _
                         ByRef UpWave As Wave, _
                         ByRef DownWave As Wave, _
                         ByRef MonitorWave As Wave, _
                         ByVal FreqStepSize As Double, _
                         ByVal PeakHangeTime As Long, _
                         Optional ByVal Verbose As Boolean = False)
                         
    Dim doContinue As Boolean
    Dim SkipLabel As Boolean
    Dim i As Long
    Dim j As Long
    Dim N As Long
    Dim Freq As Double
    Dim ctrXposition As Long
    Dim XInterval As Long
    Dim ErrorCode As Long
    
    Dim AmpInterval As Double
    Dim FreqString As String
    Dim LabelStrArray(2) As String
    
    Dim CoilString As String

    Dim MaxAmps() As Double
    Dim BiggestAmp As Double
    Dim SmallestAmp As Double
    Dim BestFreq As Double
    Dim RoundingPower As Integer
    Dim TempD As Double
    
    Dim CurTime
    Dim FolderName As String
    
    Dim DudArray() As Double
    
    'Set the AF Coil string
    If ActiveCoilSystem = AxialCoilSystem Then CoilString = "Axial"
    If ActiveCoilSystem = TransverseCoilSystem Then CoilString = "Transverse"
    
    'Check for NOCOMM_MODE, exit sub if comm is off
    If NOCOMM_MODE = True Then Exit Sub
    
    If AbortAutoTune = True Then
    
        Me.cmdAutoTuneAF.Caption = "Start Auto-Tune"
        
        Me.picDCResponse.Cls
        
        Exit Sub
        
    End If
    
    'Need an array for the Sine Fit data results for each ramp
    Dim SineFit_Data() As Double

    'Clear old drawing
    picDCResponse.Cls

    'Set BiggestAmp to zero
    BiggestAmp = 0

    'Set SmallestAmp to a relatively ginormous number
    SmallestAmp = 1000000.4161982

    'Count the number of freq steps from Minimum Freq to Max Freq for the scan
    NumSteps = CLng((MonitorWave.SineFreqMax - MonitorWave.SineFreqMin) / FreqStepSize) + 1

    'Size Max Amplitude at ech frequency array so that it has room for NumSteps number of
    'Freq and four columns:
    '           Col - 0: The frequency
    '           Col - 1: The max amplitude at that freq
    '           Col - 2: The X graph coordinate of the leftmost side of
    '                    the bar strip for that freq
    '           Col - 3: The X graph coordinate of the rightmost side of
    '                    the bar strip for that freq
    ReDim MaxAmps(NumSteps, 4)

    With MonitorWave

        'Set Font Size
        picDCResponse.FontSize = 10

        'Draw The Bounds of the DC Response Voltage Display Window
        picDCResponse.Line (1950, 1000)-(1950, 8550) 'Vertical axis
        picDCResponse.Line (1950, 8550)-(14500, 8550) 'Horizontal axis

        'Plot the units for the Y-axis
        picDCResponse.CurrentY = 200
        picDCResponse.CurrentX = 1950 - picDCResponse.TextWidth("Volts") / 2
        picDCResponse.Print "Volts"

        'Plot the label + units for the X-Axis
        picDCResponse.CurrentY = 8700 + CLng(1.5 * picDCResponse.TextWidth("0"))
        picDCResponse.CurrentX = 7750 - picDCResponse.TextWidth("Freq (Hz)")
        picDCResponse.Print "Freq (Hz)"

        'Calculate the amount of width each frequency has in the X-coordinate
        'space for plotting
        XInterval = CLng(12000 / NumSteps)

        SkipLabel = False

        'Lower font size for Freq column labels
        picDCResponse.FontSize = 9

        'Need to now find the rounding factor to use to divide Amp interval into
        'four easy to display numbers
        'NOTE:  If FreqStepSize < 0, the code below will cause an error
        '       by taking the log of a negative number!!
        RoundingPower = Int(Log(FreqStepSize) / Log(10))

        'Change Rounding Power so that it is now the number of places to
        'keep to the right of the decimal point
        If RoundingPower > 0 Then RoundingPower = 0
        RoundingPower = -1 * RoundingPower

        For i = 0 To NumSteps - 1

            'calculate left and right positions
            MaxAmps(i, 2) = 2250 + i * XInterval  'Left position is the first possible
            MaxAmps(i, 3) = 2250 + (i + 1) * XInterval

            'Plot the X axis tick marks for this freq
            picDCResponse.Line (MaxAmps(i, 2), 8550)-(MaxAmps(i, 2), 8750)
            picDCResponse.Line (MaxAmps(i, 3), 8550)-(MaxAmps(i, 3), 8750)

            'Plot the label for this Freq
            'Construct Freq String
            FreqString = Trim(str(Round(.SineFreqMin + _
                                        (.SineFreqMax - .SineFreqMin) * i / (NumSteps - 1), _
                                        RoundingPower)))

            doContinue = False

            If SkipLabel = True Then

                SkipLabel = False

            Else

                 Do

                     'Check to see if the text Width of the Freq label is greater
                     'than the XInterval for each Freq
                     If picDCResponse.TextWidth(FreqString) > 0.8 * XInterval Then

                         'Not enough vertical space, lower the font size and
                         'repeat the label size check
                         picDCResponse.FontSize = picDCResponse.FontSize - 1

                         If picDCResponse.FontSize <= 8.25 Then

                             'Skip every other label
                             SkipLabel = True

                             'Plot this label
                             picDCResponse.CurrentX = CLng(XInterval / 2 _
                                                     - picDCResponse.TextWidth(FreqString) / 2) _
                                                    + MaxAmps(i, 2)
                             picDCResponse.CurrentY = 8700
                             picDCResponse.Print FreqString

                             doContinue = False

                         Else

                            doContinue = True

                        End If

                     Else

                         'There's enough room to plot the Freq label horizontally
                         picDCResponse.CurrentX = CLng(XInterval / 2 _
                                                         - picDCResponse.TextWidth(FreqString) / 2) _
                                                 + MaxAmps(i, 2)

                         picDCResponse.CurrentY = 8700

                         picDCResponse.Print FreqString

                         doContinue = False

                     End If

                Loop Until doContinue = False

            End If

        Next i

        frmAFTuner.refresh

    End With

    'Counter for the Amplitudes at Freq array
    j = 0

    'Start Freq Iteration loop
    For Freq = MonitorWave.SineFreqMin To MonitorWave.SineFreqMax Step FreqStepSize

        'Need to now find the rounding factor to use to divide Amp interval into
        'four easy to display numbers
        'NOTE:  If FreqStepSize < 0, the code below will cause an error
        '       by taking the log of a negative number!!
        RoundingPower = Int(Log(FreqStepSize) / Log(10))

        'Change Rounding Power so that it is now the number of places to
        'keep to the right of the decimal point
        If RoundingPower > 0 Then RoundingPower = 0
        RoundingPower = -1 * RoundingPower

        FreqString = Trim(str(Round(Freq, RoundingPower)))

        'Overwrite last posting with a white box
        picDCResponse.Line (5000, 0)-(14500, 1000), _
                            QBColor(15), _
                            BF

        'UpDate plot window with status
        picDCResponse.CurrentX = 5000
        picDCResponse.CurrentY = 500
        picDCResponse.FontSize = 8
        picDCResponse.Print FreqString & " Hz: Ramping..."
        
        'Update the program status bar
        frmProgram.StatusBar "AF " & CoilString & " Auto-tune @ " & FreqString & " Hz", 2
        
        Me.refresh
        
        'Need to temporarily change MonitorWave.SineFreqMin to current scanning freq
        TempD = MonitorWave.SineFreqMin
        MonitorWave.SineFreqMin = Freq
        
        'Need to calculate the slope to ramp up and ramp down and set that
        'in the UpWave and DownWave objects
        UpWave.Slope = UpWave.PeakVoltage
        DownWave.Slope = UpWave.Slope
        
        'Set the uncalibrated Ramp option in the frmADWIN_AF
        frmADWIN_AF.optCalRamp(1).value = True
        frmADWIN_AF.optCalRamp(0).value = False
        
        If AbortAutoTune = True Then
        
            Me.cmdAutoTuneAF.Caption = "Start Auto-Tune"
            
            Me.picDCResponse.Cls
            
            Exit Sub
            
        End If
        
        
        'Need to Ramp up Now - in clipping mode
        ErrorCode = frmADWIN_AF.DoRampADWIN(MonitorWave, _
                                            UpWave, _
                                            DownWave, _
                                            AFData(), _
                                            1, _
                                            PeakHangeTime, _
                                            3, _
                                            1)
                                
        If ErrorCode <> 0 Then
        
            Exit Sub
            
        End If
        
        'Return MonitorWave.SineFreqMin to it's original value
        MonitorWave.SineFreqMin = TempD
                                    
        'Overwrite last posting with a white box
        picDCResponse.Line (5000, 0)-(14500, 1000), _
                            QBColor(11), _
                            BF

        'UpDate plot window with status
        picDCResponse.CurrentX = 5000
        picDCResponse.CurrentY = 500
        picDCResponse.FontSize = 8
        picDCResponse.Print FreqString & " Hz: Analyzing..."
        
        Me.refresh
                                    
        If AbortAutoTune = True Then
        
            Me.cmdAutoTuneAF.Caption = "Start Auto-Tune"
            
            Me.picDCResponse.Cls
            
            Exit Sub
            
        End If
                
        'Update the program status bar
        frmProgram.StatusBar "Analyzing...", 3

        MaxAmps(j, 0) = MonitorWave.CurrentVoltage
        MaxAmps(j, 1) = Freq
                    
        If MaxAmps(j, 0) > BiggestAmp Then
        
            BiggestAmp = MaxAmps(j, 0)
            BestFreq = MaxAmps(j, 1)
            
        End If
        
        If MaxAmps(j, 0) < SmallestAmp Then SmallestAmp = MaxAmps(j, 0)
                    
        'Increment the counter for the Max Amplitude array
        j = j + 1
            
        'Overwrite last posting with a white box
        picDCResponse.Line (5000, 0)-(14500, 1000), _
                            QBColor(15), _
                            BF

        'UpDate plot window with status
        picDCResponse.CurrentX = 5000
        picDCResponse.CurrentY = 500
        picDCResponse.FontSize = 8
        picDCResponse.Print FreqString & " Hz: Done"
        
        Me.refresh
        
        'Update the program status bar
        frmProgram.StatusBar vbNullString, 3
        
        'Pause 200 milliseconds so the user can see the "Done" update flash by
        PauseTill timeGetTime() + 200
            
    Next Freq

    'Now need to plot the results of this tuning pass to the graph plot
    
    'Update the program status bar
    frmProgram.StatusBar "AF " & CoilString & " Auto-tune", 2
    frmProgram.StatusBar "Plotting Data...", 3

    'Find range of difference between Biggest and Smallest Amplitudes
    AmpInterval = BiggestAmp - SmallestAmp

    'IF AmpInterval is negative, reverse the biggest and smallest amplitudes
    If AmpInterval < 0 Then
    
        TempD = BiggestAmp
        BiggestAmp = SmallestAmp
        SmallestAmp = TempD
        AmpInterval = -1 * AmpInterval
        
    ElseIf AmpInterval = 0 Then
    'If AmpInterval is zero, make it slightly non-zero

        AmpInterval = 0.00001
        
    End If

    'Need to now find the rounding factor to use to divide Amp interval into
    'four easy to display numbers
    'NOTE:  If BiggestAmp < Smallest Amp, the code below will cause an error
    '       by taking the log of a negative number!!
    RoundingPower = Int(Log(AmpInterval / 4) / Log(10))

    'Change Rounding Power so that it is now the number of places to
    'keep to the right of the decimal point
    If RoundingPower > 0 Then RoundingPower = 0
    RoundingPower = -1 * RoundingPower

    'Set plot font back to ten
    picDCResponse.FontSize = 10

    j = 0

    'Need to scale and label the Y-axis
    For i = 8000 To 2000 Step -1500

        picDCResponse.Line (1800, i)-(1950, i)  'Draw Vertical tick mark

        FreqString = Trim(str(Round(SmallestAmp + j * AmpInterval / 4, RoundingPower)))
'        Debug.Print FreqString
        j = j + 1

        'Now run loop to see how to fit the entire freq label in
        'the space available
        doContinue = False

        Do

            If picDCResponse.TextWidth(FreqString) > 1700 Then

                'Cut Label into two pieces at the mid-point
                'and now check if the two pieces will fit
                If picDCResponse.TextWidth(FreqString) > 2400 Then

                    'Lower the Font size and run the loop again
                    picDCResponse.FontSize = picDCResponse.FontSize - 1
                    picDCResponse.FontName = picDCResponse.FontName
                    picDCResponse.FontSize = Int(picDCResponse.FontSize)

                    doContinue = True

                Else

                    'Print out the two lines centered around
                    'the tickmark
                    'First Piece
                    picDCResponse.CurrentX = 500
                    picDCResponse.CurrentY = i - picDCResponse.TextHeight(FreqString)
                    picDCResponse.Print Mid(FreqString, 1, Len(FreqString) \ 2)

                    'Second Piece
                    picDCResponse.CurrentX = 500
                    picDCResponse.CurrentY = i
                    picDCResponse.Print Mid(FreqString, Len(FreqString) \ 2 + 1)

                    doContinue = False

                End If

            Else

                'Freq String for label is small enough to fit in the allotted space
                'Plot the label
                picDCResponse.CurrentX = 1700 - picDCResponse.TextWidth(FreqString)
                picDCResponse.CurrentY = i - picDCResponse.TextHeight(FreqString) / 2

                picDCResponse.Print FreqString

                doContinue = False

            End If

        Loop Until doContinue = False

    Next i

    'Now Draw in the columns for each Freq
    For i = 0 To NumSteps - 1

        picDCResponse.Line _
            (CLng(MaxAmps(i, 2) + 0.1 * XInterval), 8550)-( _
                CLng(MaxAmps(i, 3) - 0.1 * XInterval), _
                8000 - CLng(6000 / AmpInterval * (MaxAmps(i, 0) - SmallestAmp))), _
            QBColor(1), _
            BF

    Next i

    'Based on the coil tuned, change the current fit resonance freq value
    If ActiveCoilSystem = AxialCoilSystem Then

        Me.txtFitAxialResFreq = Trim(str(BestFreq))

    ElseIf ActiveCoilSystem = TransverseCoilSystem Then

        Me.txtFitTransverseResFreq = Trim(str(BestFreq))

    End If
                         
    'Update the program status bar, panel 3
    '(panel 2 will be blanked in the cmdAutoTuneAF_Click subroutine that called this function)
    frmProgram.StatusBar vbNullString, 3
    
    'Unlock the coils
    CoilsLocked = False
    Me.chkLockCoils.value = Unchecked
                         
End Sub

Private Sub DoFFT(ByRef AFData() As Double, _
                  ByRef FFT_Array() As Double, _
                  ByVal StartPoint As Long, _
                  ByVal Endpoint As Long)
                  
    Dim log2N As Double
    
    Dim N As Long
    Dim i As Long
    Dim NP2 As Long
    Dim TempL As Long
    
    'Get the length of the Ramp data array
    N = UBound(AFData, 1)
    
    TempL = N
    N = N - StartPoint
    
    If TempL > Endpoint Then
        
       N = N - (TempL - Endpoint)
            
    End If
    
    'Get the next largest power of 2 after the value N
    log2N = Log(N) / Log(2)
    
    'Now calculate NP2 from the (log2N --> nearest long) ^ 2
    NP2 = 2 ^ (Int(log2N) + 1)
    
    'Redimension the FFT array so that it is NP2 x 1 in size
    ReDim FFT_Array(NP2)
    
    For i = 0 To NP2 - 1
    
        If i < N Then
        'If we're still within the bounds of the AF data array
        'load it's input data value into the corresponding array element
        'of the FFT_Array
        
            FFT_Array(i) = AFData(i + StartPoint, 0)
            
        Else
        'the section of AFData that we want has run out of points,
        'pad the remaining points of the FFT_Array with zeros
        
            FFT_Array(i) = 0
            
        End If
        
    Next i
    
    'Now we have an Array with a power of 2 length that can be loaded into the RVFFT
    'function
    modAF_MCC.RVFFT FFT_Array, NP2
    
    'Now FFT_Array contains the results of the RVFFT algorithm
                  
End Sub

Private Sub Form_Activate()

    If EnableAF = False Then
        
        'AF's not enabled, cannot Tune the AF coils
        'Tell user that calibration is turned off, but
        'can still edit values
        MsgBox "The AF module is currently disabled.  AF coil tuning" & _
               " cannot be performed." & vbNewLine & _
               "However, you can edit the resonance frequency and maximum " & _
               "voltage values by hand.", , _
               "Whoops!"
               
        'Disable all the necessary buttons on the form
        Me.cmdAutoTuneAF.Enabled = False
        Me.cmdStartAutoClipTest.Enabled = False
        
    Else
    
        'Disable all the necessary buttons on the form
        Me.cmdAutoTuneAF.Enabled = True
        Me.cmdStartAutoClipTest.Enabled = True
        
    End If

    'First propagate the locked coils state
    If CoilsLocked = True Then Me.chkLockCoils.value = Checked
    If CoilsLocked = False Then Me.chkLockCoils.value = Unchecked

    'If the window is activated, need to propagate
    'the current active coil settings to the radio buttons
    If ActiveCoilSystem = AxialCoilSystem Then
    
        optCoil(0).value = True
        
    ElseIf ActiveCoilSystem = TransverseCoilSystem Then
    
        optCoil(1).value = True
        
    Else
    
        optCoil(0).value = False
        optCoil(1).value = False
        
        ActiveCoilSystem = NoCoilSystem
        
    End If

End Sub

Private Sub Form_Load()
    
    'Set the Form Width
    Me.Width = 10815
    Me.Height = 9000
    
    '(March 11, 2011 - I Hilburn)
    'Update all the form's icons with the program Ico file
    SetFormIcon Me
    
    'Set the form caption
    Me.Caption = "AF Tuner / Clipping Test"
    
    If EnableAF = False Then
        
        'AF's not enabled, cannot Tune the AF coils
               
        'Disable all the necessary buttons on the form
        Me.cmdAutoTuneAF.Enabled = False
        Me.cmdStartAutoClipTest.Enabled = False
        
    Else
    
        'Disable all the necessary buttons on the form
        Me.cmdAutoTuneAF.Enabled = True
        Me.cmdStartAutoClipTest.Enabled = True
        
    End If
    
    'First propagate the locked coils state
    If CoilsLocked = True Then Me.chkLockCoils.value = Checked
    If CoilsLocked = False Then Me.chkLockCoils.value = Unchecked

    'If the window is activated, need to propagate
    'the current active coil settings to the radio buttons
    If ActiveCoilSystem = AxialCoilSystem Then
    
        optCoil(0).value = True
        
    ElseIf ActiveCoilSystem = TransverseCoilSystem Then
    
        optCoil(1).value = True
        
    Else
    
        optCoil(0).value = False
        optCoil(1).value = False
        
        ActiveCoilSystem = NoCoilSystem
        
    End If
    
    'Set the picture boxes scale height and width
    Me.picDCResponse.ScaleHeight = 10000
    Me.picDCResponse.ScaleWidth = 14500
    
    'Turn Debug mode off!
    Me.chkDebugMode.value = Unchecked
    
    'Make the Select Max Voltage picture box invisible
    Me.picPickMaxVoltage.Visible = False
    Me.cmdPickMaxCancel.Visible = False
    Me.cmdAcceptMaxPick.Visible = False
    Me.lblPickMaxVoltages.Visible = False
    Me.lblPickMaxResults.Visible = False
    
    'Set Visible Frames to AF Auto-Tune frames
    Me.frameAFAutoTune.Visible = True
    Me.frameClippingTest.Visible = True
    Me.frameCoilResFreq.Visible = True
    Me.frameMaxCoilVoltages.Visible = True
    
    'Default fit result values for Clipping and Auto-tune tests
    'to blank values
    Me.txtFitMaxAxialRamp = ""
    Me.txtFitMaxTransverseRamp = ""
    Me.txtNewMaxAxialMonitor = ""
    Me.txtNewMaxTransverseMonitor = ""
    Me.txtFitAxialResFreq = ""
    Me.txtFitTransverseResFreq = ""
    Me.txtOldAxialResFreq = Trim(str(modConfig.AfAxialResFreq))
    Me.txtOldTransverseResFreq = Trim(str(modConfig.AfTransResFreq))
    Me.txtOldMaxAxialRamp = Trim(str(modConfig.AfAxialRampMax))
    Me.txtOldMaxTransverseRamp = Trim(str(modConfig.AfTransRampMax))
    Me.txtOldMaxAxialMonitor = Trim(str(modConfig.AfAxialMonMax))
    Me.txtOldMaxTransverseMonitor = Trim(str(modConfig.AfTransMonMax))
    
    'Clear the Picture Box plot
    picDCResponse.Cls
        
    'Preset Auto-tune run paramters
    Me.txtAmplitude.text = 0.5
    Me.txtHighFreq.text = ""
    Me.txtLowFreq.text = ""
    Me.txtDuration.text = 500
    Me.txtFreqStepSize = ""
    
    'Preset all labels (except status field) to non-bold and black font color
    lblLowFreq.ForeColor = vbBlack
    lblHighFreq.ForeColor = vbBlack
    lblScanDuration.ForeColor = vbBlack
    lblAmplitude.ForeColor = vbBlack
    
    lblLowFreq.FontBold = False
    lblHighFreq.FontBold = False
    lblScanDuration.FontBold = False
    lblAmplitude.FontBold = False
    
    'Plot default Axes on the results Picture plot for the tuning sweep
    picDCResponse.AutoRedraw = True
    
End Sub

Private Sub optCoil_Click(Index As Integer)
    
    If CoilsLocked = True Then Exit Sub
    
    If Index = 0 And _
       optCoil(Index).value = True _
    Then
        
        ActiveCoilSystem = AxialCoilSystem
        CoilString = "Axial"
        
        'Set the Freq for the auto-clip test
        Me.txtClippingSineFreq = Me.txtOldAxialResFreq
                
        'Change the coil relay
        If AFSystem = "2G" Then
        
            frmAF_2G.ConfigureCoil modConfig.AfAxialCoord
            
        ElseIf AFSystem = "ADWIN" Then
        
            frmADWIN_AF.SetAFRelays
            
        End If
                
    ElseIf Index = 1 And _
           optCoil(Index).value = True _
    Then
    
        ActiveCoilSystem = TransverseCoilSystem
        CoilString = "Transverse"
        
        'Set the Freq for the auto-clip test
        Me.txtClippingSineFreq = Me.txtOldTransverseResFreq
                
        'Change the coil relay
        If AFSystem = "2G" Then
        
            frmAF_2G.ConfigureCoil modConfig.AfTransCoord
            
        ElseIf AFSystem = "ADWIN" Then
        
            frmADWIN_AF.SetAFRelays
            
        End If
                
    Else
    
        CoilString = ""
        ActiveCoilSystem = NoCoilSystem
               
        If AFSystem = "ADWIN" Then
        
            frmADWIN_AF.SetAFRelays
            
        End If
                
    End If
    
End Sub

Private Sub picDCResponse_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    'Check to see if picPickMaxVoltage is visible
    If picPickMaxVoltage.Visible <> True Then
    
        'Not in the mode where user needs to click on this picture box
        Exit Sub
        
    ElseIf Button = vbRightButton Or Button = vbLeftButton Then
    
        'Check to see if user has clicked within the bounds of the graph
        If X >= 1950 And X <= 13000 And Y >= 2000 And Y <= 8000 Then
        
            'Check to see if it's the ramp or the monitor voltage we are selecting
            If VSelectMode = "RAMP" Then
            
                'User has clicked within the graph, need to determine what
                'corresponding Ramp & Monitor Voltage value they've clicked on
                SelectedRampV = (X - 1950) / 11500 * VoltRange - val(Me.txtMinClippingAmp)
                Me.lblPickMaxResults = "Ramp Max    = " & Format(SelectedRampV, "0.0#") & " Volts"
                
                'Draw a vertical light green line indicating the ramp max position
                picDCResponse.Line (X, 8000)-(X, 3000), QBColor(10)
                
                'Change prompt
                Me.lblPickMaxVoltages = "Now, please select the max AF " & CoilString & _
                                        " coil Monitor Voltage."
                                        
                'Now change highlight color on the pick max voltage picture box border
                Me.shapePickVoltageBorder.BorderColor = QBColor(13)
                
                'Set the voltage selection mode to select Monitor voltage!
                VSelectMode = "MONITOR"
                
            ElseIf VSelectMode = "MONITOR" Then
            
                SelectedMonitorV = (8000 - Y) / 6000 * AmpInterval - MinAmp
                Me.lblPickMaxResults = "Ramp Max    = " & Format(SelectedRampV, "0.0#") & " Volts" & _
                                       vbCrLf & _
                                       "Monitor Max = " & Format(SelectedMonitorV, "0.0#") & " Volts"
            
                'Now need to display this value in the picPickMaxVolt picture box
                Me.picPickMaxVoltage.SetFocus
                
                'Now change highlight color on the pick max voltage picture box border
                Me.shapePickVoltageBorder.BorderColor = QBColor(10)
                
                'Set the voltage selection mode to select Ramp voltage!
                VSelectMode = "RAMP"
                
                'Put in final text to explain what to do
                Me.lblPickMaxVoltages = "Click 'OK' to accept, or pick a new max Ramp voltage."
                
                'Show the 'OK' button now
                Me.cmdAcceptMaxPick.Visible = True
            
            End If
            
        End If
        
    End If
    
    'Refresh the picture box and form
    Me.picPickMaxVoltage.refresh
    Me.refresh
        
End Sub

Public Sub PlotAutoClipTestResults(ByRef SineFit_Data() As Double, _
                                    ByRef PictureObj As PictureBox, _
                                    ByVal StartVoltage As Double, _
                                    ByVal HighVoltage As Double, _
                                    ByVal PeakPoint As Long)
                                    
    Dim i As Long
    Dim j As Long
    Dim N As Long
    
    Dim RoundingPower As Long
    Dim LabelString As String
    Dim doContinue As Boolean
    Dim SkipLabel As Boolean
    Dim CurX As Long
    Dim CurY As Long
    Dim PrevX As Long
    Dim PrevY As Long
    Dim TempD As Double
                       
    'Update the program form 2nd status bar panel
    frmProgram.StatusBar "Plotting Data...", 3
                       
    'Get number of sine-fits to plot
    N = UBound(SineFit_Data, 1)
                       
    'Draw in Axes and Unit Labels
                                
    'Set Font Size
    PictureObj.FontSize = 10
    
    'Clear Picture Box
    PictureObj.Cls
       
    'Set Picture Object scale height and width properties
    PictureObj.ScaleHeight = 10000
    PictureObj.ScaleWidth = 14500
       
    'Draw The Bounds of the DC Response Voltage Display Window
    PictureObj.Line (1950, 1000)-(1950, 8000) 'Vertical axis - RMS
    PictureObj.Line (13000, 1000)-(13000, 8000) 'Vertical axis - Monitor Voltage
    PictureObj.Line (1950, 8000)-(13000, 8000) 'Horizontal axis
    
    'Plot the units for the Y-axis
    PictureObj.CurrentY = 200
    PictureObj.CurrentX = 1950 - PictureObj.TextWidth("RMS") / 2
    PictureObj.Print "RMS"
    
    'Draw in Lines with corresponding colors for RMS data
    PictureObj.Line (3000, 550)-(4000, 550), QBColor(1)
    PictureObj.DrawStyle = 2
    PictureObj.Line (3000, 750)-(4000, 750), QBColor(1)
    PictureObj.DrawStyle = 0
    
    'Plot the units for the Y-axis
    PictureObj.CurrentY = 200
    PictureObj.CurrentX = 13000 - PictureObj.TextWidth("Monitor") / 2
    PictureObj.Print "Monitor"
    PictureObj.CurrentY = 200 + 0.8 * PictureObj.TextHeight("Monitor")
    PictureObj.CurrentX = 13000 - PictureObj.TextWidth("Voltage") / 2
    PictureObj.Print "Voltage"
            
    'Draw in Lines with corresponding colors for RMS data
    PictureObj.Line (10500, 550)-(11500, 550), QBColor(13)
    PictureObj.DrawStyle = 2
    PictureObj.Line (10500, 750)-(11500, 750), QBColor(13)
    PictureObj.DrawStyle = 0
    
    'Type in Legend
    PictureObj.FontSize = 8
    PictureObj.CurrentY = 300
    PictureObj.CurrentX = 6000
    PictureObj.Print "Solid  = Ramp Up"
    PictureObj.CurrentY = 300 + 1.5 * PictureObj.TextHeight("Ramp Up")
    PictureObj.CurrentX = 6000
    PictureObj.Print "Dashed = Ramp Down"
        
    'Return Font-Size to 10
    PictureObj.FontSize = 10
            
    'Plot the label + units for the X-Axis
    PictureObj.CurrentY = 8700 + CLng(1.5 * PictureObj.TextWidth("0"))
    PictureObj.CurrentX = 7750 - PictureObj.TextWidth("Ramp Output Voltage")
    PictureObj.Print "Ramp Output Voltage"
    
    'Initialize max and min RMS holder variables
    MaxRMS = SineFit_Data(0, 8)
    MinRMS = SineFit_Data(0, 8)
    MaxAmp = SineFit_Data(0, 3)
    
    'Need to find Max and Min RMS now
    For i = 0 To N - 1
    
        If SineFit_Data(i, 9) >= StartVoltage Then
        
            If MaxRMS < SineFit_Data(i, 8) Then MaxRMS = SineFit_Data(i, 8)
            If MinRMS > SineFit_Data(i, 8) And SineFit_Data(i, 8) > 0 Then MinRMS = SineFit_Data(i, 8)
        
            If MaxAmp < SineFit_Data(i, 3) Then MaxAmp = SineFit_Data(i, 3)
            
        End If
    
'        Debug.Print MaxAmp
    
    Next i
        
    'Now can scale and label the y-axis for RMS
    'Find range of difference between Biggest and Smallest Amplitudes
    RMSInterval = MaxRMS - MinRMS
    
    'If RMS Interval < 0, swap min and max RMS
    If RMSInterval < 0 Then
    
        TempD = MaxRMS
        MaxRMS = MinRMS
        MinRMS = TempD
        RMSInterval = -1 * RMSInterval
        
    ElseIf RMSInterval = 0 Then
    
        RMSInterval = 0.0001
        
    End If
        
    'Need to now find the rounding factor to use to divide Amp interval into
    'four easy to display numbers
    'NOTE:  If BiggestAmp < Smallest Amp, the code below will cause an error
    '       by taking the log of a negative number!!
    RoundingPower = Int(Log(RMSInterval / 4) / Log(10))
    
    'Change Rounding Power so that it is now the number of places to
    'keep to the right of the decimal point
    If RoundingPower > 0 Then RoundingPower = 0
    RoundingPower = -1 * RoundingPower
    
    'Need to keep track of how different from the rounded label,
    'the true max value is
    MaxValueMinusLabel_RMS = MaxRMS - Round(MaxRMS, RoundingPower)
    MinValueMinusLabel_RMS = MinRMS - Round(MinRMS, RoundingPower)
    
    'Initialize j to zero
    j = 0
    
    'Need to scale and label the Y-axis
    For i = 8000 To 2000 Step -1500
    
        PictureObj.Line (1800, i)-(1950, i)  'Draw Vertical tick mark
        
        LabelString = Trim(str(Round(MinRMS + j * RMSInterval / 4, RoundingPower)))
'        Debug.Print LabelString
        j = j + 1
        
        'Now run loop to see how to fit the entire freq label in
        'the space available
        doContinue = False
        
        Do
        
            If PictureObj.TextWidth(LabelString) > 1700 Then
            
                'Cut Label into two pieces at the mid-point
                'and now check if the two pieces will fit
                If PictureObj.TextWidth(LabelString) > 800 Then
                
                    'Lower the Font size and run the loop again
                    PictureObj.FontSize = PictureObj.FontSize - 1
                    PictureObj.FontName = PictureObj.FontName
                    PictureObj.FontSize = Int(PictureObj.FontSize)
                    
                    doContinue = True
                    
                Else
                
                    'Print out the two lines centered around
                    'the tickmark
                    'First Piece
                    PictureObj.CurrentX = 500
                    PictureObj.CurrentY = i - PictureObj.TextHeight(LabelString)
                    PictureObj.Print Mid(LabelString, 1, Len(LabelString) \ 2)
                    
                    'Second Piece
                    PictureObj.CurrentX = 500
                    PictureObj.CurrentY = i
                    PictureObj.Print Mid(LabelString, Len(LabelString) \ 2 + 1)
                    
                    doContinue = False
                    
                End If
                
            Else
            
                'Freq String for label is small enough to fit in the allotted space
                'Plot the label
                PictureObj.CurrentX = 1700 - PictureObj.TextWidth(LabelString)
                PictureObj.CurrentY = i - PictureObj.TextHeight(LabelString) / 2
                
                PictureObj.Print LabelString
                
                doContinue = False
                
            End If
        
        Loop Until doContinue = False
        
    Next i
    
    
    'Now can scale and label the y-axis for Monitor Voltage
    'Find range of difference between Biggest and Smallest Amplitudes
    'Debug.Print MaxAmp
    AmpInterval = MaxAmp
    
    If AmpInterval < 0 Then
    
        MaxAmp = -1 * MaxAmp
        AmpInterval = -1 * AmpInterval
        
    ElseIf AmpInterval = 0 Then
    
        AmpInterval = 0.0001
        
    End If
    
    'Need to now find the rounding factor to use to divide Amp interval into
    'four easy to display numbers
    'NOTE:  If BiggestAmp < Smallest Amp, the code below will cause an error
    '       by taking the log of a negative number!!
    RoundingPower = Int(Log(AmpInterval / 4) / Log(10))
    
    'Change Rounding Power so that it is now the number of places to
    'keep to the right of the decimal point
    If RoundingPower > 0 Then RoundingPower = 0
    RoundingPower = -1 * RoundingPower
    
    'Need to keep track of how different from the rounded label,
    'the true max value is
    MaxValueMinusLabel_Amp = MaxAmp - Round(MaxAmp, RoundingPower)
        
    'Initialize j to zero
    j = 0
    
    'Need to scale and label the Y-axis
    For i = 8000 To 2000 Step -1500
    
        PictureObj.Line (13000, i)-(13150, i)  'Draw Vertical tick mark
        
        LabelString = Trim(str(Round(j * AmpInterval / 4, RoundingPower)))
'        Debug.Print LabelString
        j = j + 1
        
        'Now run loop to see how to fit the entire freq label in
        'the space available
        doContinue = False
        
        Do
        
            If PictureObj.TextWidth(LabelString) > 1700 Then
            
                'Cut Label into two pieces at the mid-point
                'and now check if the two pieces will fit
                If PictureObj.TextWidth(LabelString) > 800 Then
                
                    'Lower the Font size and run the loop again
                    PictureObj.FontSize = PictureObj.FontSize - 1
                    PictureObj.FontName = PictureObj.FontName
                    PictureObj.FontSize = Int(PictureObj.FontSize)
                    
                    doContinue = True
                    
                Else
                
                    'Print out the two lines centered around
                    'the tickmark
                    'First Piece
                    PictureObj.CurrentX = 13200
                    PictureObj.CurrentY = i - PictureObj.TextHeight(LabelString)
                    PictureObj.Print Mid(LabelString, 1, Len(LabelString) \ 2)
                    
                    'Second Piece
                    PictureObj.CurrentX = 13200
                    PictureObj.CurrentY = i
                    PictureObj.Print Mid(LabelString, Len(LabelString) \ 2 + 1)
                    
                    doContinue = False
                    
                End If
                
            Else
            
                'Freq String for label is small enough to fit in the allotted space
                'Plot the label
                PictureObj.CurrentX = 13200
                PictureObj.CurrentY = i - PictureObj.TextHeight(LabelString) / 2
                
                PictureObj.Print LabelString
                
                doContinue = False
                
            End If
        
        Loop Until doContinue = False
        
    Next i
    
    'Now Plot X-axis labels and tick marks
    
    'Figure out the range of voltages that were swept through
    VoltRange = HighVoltage - StartVoltage
    
    'if VoltRange is negative, need to swap High & Start voltages
    If VoltRange < 0 Then
    
        TempD = HighVoltage
        HighVoltage = StartVoltage
        StartVoltage = TempD
        VoltRange = -1 * VoltRange
    
    ElseIf VoltRange = 0 Then
    
        VoltRange = 0.0001
        
    End If
        
    'Need to now find the rounding factor to use to divide Amp interval into
    'four easy to display numbers
    'NOTE:  If BiggestAmp < Smallest Amp, the code below will cause an error
    '       by taking the log of a negative number!!
    RoundingPower = Int(Log(VoltRange / 10) / Log(10))
    
    'Change Rounding Power so that it is now the number of places to
    'keep to the right of the decimal point
    If RoundingPower > 0 Then RoundingPower = 0
    RoundingPower = -1 * RoundingPower
    
    'Initialize j to zero
    j = 0
    
    'Initialize font to size 9
    PictureObj.FontSize = 9
    
    For i = 1950 To 13000 Step 1105
    
        PictureObj.Line (i, 8000)-(i, 8200)  'Draw Vertical tick mark
        
        LabelString = Trim(str(Round(StartVoltage + j * VoltRange / 10, RoundingPower)))
'       Debug.Print LabelString
        j = j + 1
        
        doContinue = False
        
        If SkipLabel = True Then
        
            SkipLabel = False
            
        Else
            
            Do
             
                'Check to see if the text Width of the Freq label is greater
                'than the XInterval for each Freq
                If PictureObj.TextWidth(LabelString) > 0.8 * 1225 Then
                 
                    'Not enough vertical space, lower the font size and
                    'repeat the label size check
                    PictureObj.FontSize = PictureObj.FontSize - 1
                     
                    If PictureObj.FontSize <= 8.25 Then
                     
                        'Skip everyother label
                        SkipLabel = True
                         
                        'Plot this label
                        PictureObj.CurrentX = i - CLng(PictureObj.TextWidth(LabelString) / 2)
                        PictureObj.CurrentY = 8300
                        PictureObj.Print LabelString
                         
                        doContinue = False
                         
                    Else
                     
                        doContinue = True
                        
                    End If
                     
                Else
                 
                     'There's enough room to plot the Freq label horizontally
                     PictureObj.CurrentX = i - CLng(PictureObj.TextWidth(LabelString) / 2)
                     PictureObj.CurrentY = 8300
                     
                     PictureObj.Print LabelString
                     
                     doContinue = False
                 
                 End If
            
            Loop Until doContinue = False
    
        End If
    
    Next i
    
    'Set the line draw width larger
    PictureObj.DrawWidth = 1
    
    'Now Plot the RMS VS Ramp Input Voltages
    For i = 1 To N - 1
    
        If SineFit_Data(i - 1, 9) >= StartVoltage Then
        
            'Translate this point in the SineFit_Data into an x-coordinate
            CurX = CLng(1950 + (SineFit_Data(i, 9) - StartVoltage) _
                                 * 11050 / VoltRange)
                                
            'Translate previous point in the SineFit_Data into an x-coordinate
            PrevX = CLng(1950 + (SineFit_Data(i - 1, 9) - StartVoltage) _
                                 * 11050 / VoltRange)
            
            'Translate Cur RMS value into a Y-co0rdinate
            CurY = CLng(8000 - (SineFit_Data(i, 8) - MinRMS + MinValueMinusLabel_RMS) * 6000 / RMSInterval)
            
            'Translate Previous RMS value into a Y-co0rdinate
            PrevY = CLng(8000 - (SineFit_Data(i - 1, 8) - MinRMS + MinValueMinusLabel_RMS) * 6000 / RMSInterval)
            
            'Now check to see if we're in the Ramp Up or Ramp down phase of the data
            If SineFit_Data(i, 0) < PeakPoint Then
            
                PictureObj.DrawStyle = 0
                
            Else
            
                PictureObj.DrawStyle = 2
                
            End If
        
            PictureObj.Line (PrevX, PrevY)-(CurX, CurY), QBColor(1)
        
            'Translate this point in the SineFit_Data into an x-coordinate
            CurX = CLng(1950 + (SineFit_Data(i, 9) - StartVoltage) _
                                 * 11050 / VoltRange)
                                
            'Translate previous point in the SineFit_Data into an x-coordinate
            PrevX = CLng(1950 + (SineFit_Data(i - 1, 9) - StartVoltage) _
                                 * 11050 / VoltRange)
            
            'Translate Cur RMS value into a Y-co0rdinate
            CurY = CLng(8000 - (SineFit_Data(i, 3)) * 6000 / AmpInterval)
            
            'Translate Previous RMS value into a Y-co0rdinate
            PrevY = CLng(8000 - (SineFit_Data(i - 1, 3)) * 6000 / AmpInterval)
            
            PictureObj.Line (PrevX, PrevY)-(CurX, CurY), QBColor(13)
                        
        End If
        
    Next i
    
    'Return Draw Width to 1
    PictureObj.DrawWidth = 1
    
    'Return DrawStyle to 0
    PictureObj.DrawStyle = 0
    
    'Now need to have user select the max voltage
    
    'First clear the pick voltage notification picture box
    Me.lblPickMaxVoltages = "Please Select the AF " & CoilString & " coil Max Ramp voltage " & _
                            "by clicking on the Clip Test graph."
                        
    If ActiveCoilSystem = AxialCoilSystem Then
                        
        Me.lblPickMaxResults = "Ramp Max    = " & Trim(Me.txtFitMaxAxialRamp) & _
                               " Volts" ' & vbNewLine &
                               '"Monitor Max = " & Trim(Me.txtNewMaxAxialMonitor) & " Volts"
    
    ElseIf ActiveCoilSystem = TransverseCoilSystem Then
    
        Me.lblPickMaxResults = "Ramp Max    = " & Trim(Me.txtFitMaxTransverseRamp) & _
                               " Volts" '& vbNewLine &
                               '"Monitor Max = " & Trim(Me.txtNewMaxTransverseMonitor) & " Volts"
                               
    Else
    
        Exit Sub
   
    End If
   
    'Now make the pick max voltage picture box and cancel buttons visible
    'make sure the OK button is NOT visible, though
    Me.picPickMaxVoltage.Top = 1080
    Me.picPickMaxVoltage.Left = 480
    Me.picPickMaxVoltage.Width = 3000
    Me.picPickMaxVoltage.Height = 2000
    Me.picPickMaxVoltage.Visible = True
    Me.lblPickMaxResults.Visible = True
    Me.lblPickMaxVoltages.Visible = True
    Me.cmdPickMaxCancel.Visible = True
    Me.cmdAcceptMaxPick.Visible = False
        
    'Now change highlight color on the pick max voltage picture box border
    Me.shapePickVoltageBorder.BorderColor = QBColor(14)
    
    'Refresh the picture box
    Me.picPickMaxVoltage.refresh
    
    'Refresh the Form
    Me.refresh
    
    'Set the Focus back to the PictureObj picture box
    PictureObj.SetFocus
    
    'Set the VSelect Mode = "RAMP"
    VSelectMode = "RAMP"
    
    'Update the 3rd panel of the program form status bar
    frmProgram.StatusBar "Set Max Voltages...", 3
    
End Sub

Private Sub PlotAutoTuneFFTResults _
    (ByRef AFData() As Double, _
     ByRef UpWave As Wave, _
     ByRef DownWave As Wave, _
     ByRef MonitorWave As Wave, _
     ByVal TuneTime As Double)

    Dim FFT_Array() As Double
    Dim Power_Spectrum() As Double
    Dim FreqBinSize As Double
    Dim TempD As Double
    Dim MaxPower As Double
    Dim BestFreq As Double
    Dim PointPair(4) As Double
    
    Dim i As Long
    Dim j As Long
    Dim N As Long
    Dim Ndiv2 As Long
    Dim MinFreqIndex As Long
    Dim MaxFreqIndex As Long
    
    Dim LabelString As String
    
    Dim doContinue As Boolean
    
    'Do the FFT analysis on the AF input data
    DoFFT AFData, FFT_Array, UpWave.CurrentPoint, DownWave.StartPoint
    
    'Get the length of the FFT_ARRAY
    N = UBound(FFT_Array)
    
    'NDiv2 = N / 2 = N \ 2, which must be a long value, given that the length of the
    'FFT_ARRAY is a power of two
    Ndiv2 = N \ 2
    
    'For an FFT, the frequency bin size is a product of the total number of time-domain
    'data points, the array element of the FFT_Array, and the sampling time of the time-domain
    'data.
    ' delta-FREQ = 1 / (# time pts * TimeStep)
    FreqBinSize = 1 / (UBound(AFData, 1) * MonitorWave.TimeStep)
    
    'NDiv2 = the number of power spectrum elements in the FFT_Array
    'We don't want all of these values, just the range covering the frequencies that we're
    'interested in. Therefore, now we need to find the max and min freq-bin indices that
    'bound the frequency spectrum that we're interested
    MinFreqIndex = Int(MonitorWave.SineFreqMin / FreqBinSize)
    MaxFreqIndex = Int(MonitorWave.SineFreqMax / FreqBinSize) + 1
    
    If MinFreqIndex > MaxFreqIndex Then
    
        'Swap the max and min frequency indices
        TempD = MaxFreqIndex
        MaxFreqIndex = MinFreqIndex
        MinFreqIndex = TempD
        
    End If
    
    'Redimension the Power_Spectrum array to abs(MaxFreqIndex - MinFreqIndex)
    ReDim Power_Spectrum(Abs(MaxFreqIndex - MinFreqIndex) + 1)
    
    'Initialize Max Power to 0 and BestFreq = -1
    MaxPower = 0
    BestFreq = -1
    
    'Now determine the power spectrum for the range of indices that we want from
    'the FFT_Array.
    'Because the second half of the FFT_Array holds the complex valued results:
    'Power_Spectrum(i) = (FFT_Array(i)) ^ 2 + (FFT_Array(i+NDiv2)) ^ 2
    For i = MinFreqIndex To MaxFreqIndex Step 1
    
        'Determine the power spectrum for this frequency bin
        TempD = (FFT_Array(i)) ^ 2 + (FFT_Array(i + Ndiv2)) ^ 2
        
        Power_Spectrum(i - MinFreqIndex) = TempD
                                           
        'Now need to find the maximum power in the Power Spectrum array
        'and the corresponding frequency bin
        '(Note: TempD >= 0, always)
        If MaxPower < TempD Then
        
            MaxPower = TempD
            BestFreq = i * FreqBinSize
        
        End If
        
    Next i
    
    'Now we have the piece of the power_spectrum that we need, we can begin to
    'plot the data in the Picture Box on the AF Tuner form
    
    'Draw in Axes and Unit Labels
                                
    'Set Font Size
    picDCResponse.FontSize = 10
    
    'Clear Picture Box
    picDCResponse.Cls
       
    'Draw The Bounds of the FFT Results Display Window
    picDCResponse.Line (1950, 1000)-(1950, 8000) 'Vertical axis - Freq Power
    picDCResponse.Line (1950, 8000)-(13000, 8000) 'Horizontal axis - Freq
    
    'Plot the units for the Y-axis
    picDCResponse.CurrentY = 200
    picDCResponse.CurrentX = 1950 - picDCResponse.TextWidth("Power") / 2
    picDCResponse.Print "Power"
    
    'Plot the label + units for the X-Axis
    picDCResponse.CurrentY = 8700 + _
                             CLng(1.5 * _
                                  picDCResponse.TextHeight(Trim(str("0"))))
                             
    picDCResponse.CurrentX = 7750 - picDCResponse.TextWidth("Freq (Hz)")
    picDCResponse.Print "Freq (Hz)"
    
    'Set j = 0
    j = 0
    
    'We already have the Maxium power - use to scale the labels on the Y axis
    'Put the maximum tick-mark at 9000 (so have 8000 points for Y-axis range)
    For i = 1000 To 9000 Step 2000
    
        picDCResponse.Line (1800, i)-(1950, i)  'Draw Vertical tick mark
        
        LabelString = Format((MaxPower * j / 4) / MaxPower, "0.##")
        
        'If there's a dot at the end of label string (i.e. the value
        'is an integer) then prune it off
        If Right(LabelString, 1) = "." Then
        
            LabelString = Mid(LabelString, 1, Len(LabelString) - 1)
        
        End If
        
'        Debug.Print LabelString
        j = j + 1
        
        'Freq String for label is small enough to fit in the allotted space
        'Plot the label
        picDCResponse.CurrentX = 1700 - picDCResponse.TextWidth(LabelString)
        picDCResponse.CurrentY = i - picDCResponse.TextHeight(LabelString) / 2
        
        picDCResponse.Print LabelString
        
    Next i
    
    'Set j = 0
    j = 0
    
    'Now need to label and scale the X-axis
    For i = 2500 To 12500 Step 2500
    
        picDCResponse.Line (i, 8000)-(i, 8150) ' Line Tick mark
        
        With MonitorWave
        
            'Generate the LabelString
            LabelString = Format(.SineFreqMin + j / 4 * (.SineFreqMax - .SineFreqMin), _
                                 "0.#")
                                 
            'If there's a dot at the end of label string (i.e. the value
            'is an integer) then prune it off
            If Right(LabelString, 1) = "." Then
            
                LabelString = Mid(LabelString, 1, Len(LabelString) - 1)
            
            End If
                                 
        End With
        
        'Increment j
        j = j + 1
        
        'Move cursor to X & Y position for printing label on graph
        With picDCResponse
        
            .CurrentX = i - .TextWidth(LabelString) / 2
            .CurrentY = 8150 + .TextHeight(LabelString)
            picDCResponse.Print LabelString
            
        End With
        
    Next i
          
    'Now need to display the data, in point pairs
    For i = MinFreqIndex To MaxFreqIndex - 1 Step 1
    
        'Determine the X & Y coordinate of the data point
    
        'Remember the Low Freq value corresponds to X = 2500
        'and the High Freq Value corresponds to X = 12500
        
        'First point of the line (i)
        PointPair(0) = (10000 / Abs(MonitorWave.SineFreqMax - MonitorWave.SineFreqMin)) * _
                       (i * FreqBinSize - MonitorWave.SineFreqMin) + 2500
        PointPair(1) = (8000 * Power_Spectrum(i - MinFreqIndex) / MaxPower) + 1000
        
        'Second point of the line (i + 1)
        PointPair(2) = (10000 / Abs(MonitorWave.SineFreqMax - MonitorWave.SineFreqMin)) * _
                       ((i + 1) * FreqBinSize - MonitorWave.SineFreqMin) + 2500
        PointPair(3) = (8000 * Power_Spectrum(i + 1 - MinFreqIndex) / MaxPower) + 1000
        
        picDCResponse.Line (PointPair(0), PointPair(1))-(PointPair(2), PointPair(3))
        
    Next i

    'Now set the appropriate New Resonance Freq text box = best frequency
    If ActiveCoilSystem = AxialCoilSystem Then
    
        Me.txtFitAxialResFreq = Format(BestFreq, "0.#")
        
    ElseIf ActiveCoilSystem = TransverseCoilSystem Then
    
        Me.txtFitTransverseResFreq = Format(BestFreq, "0.#")
        
    End If

End Sub

Private Sub SetupAFAutoTune()

    Dim AFData() As Double
    Dim SineFit_Data() As Double
    Dim PeakHangeTime As Double

    'Update the program form status bar
    frmProgram.StatusBar "AF Auto-tune Config", 2
    
    'Make sure system Wave and Board objects are setup
    If WaveForms Is Nothing Or SystemBoards Is Nothing Then
    
        'Raise an Error
        Err.Raise -616, _
                  "frmAFTuner.cmbAutoTuneAF", _
                  "AF System wave-forms and/or the System Boards collection have not been loaded " & _
                  "properly.  Please check the Paleomag.ini file." & vbNewLine & vbNewLine & _
                  "The code will end now."
                  
        End
        
    End If
    
    'Lock the Coils
    CoilsLocked = True
    Me.chkLockCoils.value = Checked
    
    'Change sine freq, amplitude, and IO-rate (used to set ADWIN process delay)
    With WaveForms("AFMONITOR")
    
        .SineFreqMin = val(Me.txtLowFreq)
        .SineFreqMax = val(Me.txtHighFreq)
        .PeakVoltage = val(Me.txtAmplitude)
        Set .range = New range
        .range.MaxValue = 10
        .range.MinValue = -10
                
    End With
    
    'Just need to set the WaveForms("AFRAMPUP") peak and duration - used to determine the SlopeUp
    'parameter that's passed into the ADWIN ramp process
    With WaveForms("AFRAMPUP")
    
        .PeakVoltage = val(Me.txtAmplitude)
        Set .range = WaveForms("AFMONITOR").range
        
    End With
    
    'Just need to set the DownWave peak and duration - used to determine the SlopeDown
    'parameter that's passed into the ADWIN ramp process
    With WaveForms("AFRAMPDOWN")
    
        .PeakVoltage = WaveForms("AFRAMPUP").PeakVoltage
        Set .range = WaveForms("AFMONITOR").range
        
    End With
    
    'Set the peak delay time to user inputed value
    PeakHangeTime = val(Me.txtDuration)
    
    'Set the freq step size
    FreqStepSize = val(Me.txtFreqStepSize)
    
    'Do the AF auto tuning
    doAFAutoTune AFData, _
                 WaveForms("AFRAMPUP"), _
                 WaveForms("AFRAMPDOWN"), _
                 WaveForms("AFMONITOR"), _
                 FreqStepSize, _
                 PeakHangeTime, _
                 (Me.chkDebugMode.value = Checked)
                           
    'Reset the 2nd status bar panel
    frmProgram.StatusBar vbNullString, 2
    
    'Unlock the Coils
    CoilsLocked = False
    Me.chkLockCoils.value = Unchecked
     
End Sub

Private Sub txtAmplitude_LostFocus()

    Dim temp As Double
    
    temp = val(txtAmplitude)
    
    If temp < 0 Then txtAmplitude = "0"
    
    If temp > 10 Then txtAmplitude = "10"

End Sub

Private Sub txtClippingSineFreq_lostFocus()

    Dim temp As Double
    
    temp = val(Me.txtClippingSineFreq)
    
    If temp < 0 Then Me.txtClippingSineFreq = "1"
    
    If temp > 5000 Then Me.txtClippingSineFreq = "5000"

End Sub

Private Sub txtDuration_LostFocus()

    Dim temp As Double
    
    temp = val(txtDuration)
    
    If temp < 0 Then txtDuration = "0"
    
    If temp > 15000 Then txtDuration = "15000"

End Sub

Private Sub txtFreqStepSize_LostFocus()

    Dim TempD As Double
    Dim MaxRange As Double
    
    MaxRange = Abs(val(Me.txtHighFreq) - val(Me.txtLowFreq))
    TempD = val(Me.txtFreqStepSize)
    
    If TempD < 0 Then
    
        Me.txtFreqStepSize = Trim(str(MaxRange \ 10))
        
        Exit Sub
        
    End If
    
    If TempD > MaxRange Then
    
        Me.txtFreqStepSize = Trim(str(MaxRange))
    
        Exit Sub
        
    End If

End Sub

Private Sub txtHighFreq_Change()

    If val(Me.txtHighFreq.text) > 1500 Then
    
        Me.txtHighFreq.text = "1500"
        
    End If
    
End Sub

Private Sub txtMaxClipAmp_Change()

    'Trigger the two Slope text-box control change events
    txtRampUpSlope_Change
    txtRampDownSlope_Change

End Sub

Private Sub txtMaxClipAmp_LostFocus()

    txtMinClippingAmp_LostFocus

End Sub

Private Sub txtMinClippingAmp_LostFocus()
    
    Dim tempMin As Double
    Dim tempMax As Double
    
    tempMin = val(txtMinClippingAmp)
    tempMax = val(txtMaxClipAmp)
    
    If tempMin < 0 Then txtMinClippingAmp = "0"

    
    If tempMax < 0 Then txtMaxClipAmp = "0"
        
    If tempMin > 10 Then txtMinClippingAmp = "10"
    
    tempMin = val(txtMinClippingAmp)
    tempMax = val(txtMaxClipAmp)
    
    If tempMin > tempMax Then
    
        tempMax = tempMin + 0.1
        txtMaxClipAmp = Trim(str(tempMax))
        
    End If
    
    tempMax = val(txtMaxClipAmp)
    
    If tempMax > 10 Then txtMaxClipAmp = "10"
    
End Sub

Private Sub txtNumSineFits_LostFocus()

    Dim TempD As Double
    
    TempD = val(Me.txtNumSineFits)
    
    Me.txtNumSineFits = Round(TempD, 0)
    
End Sub

Private Sub txtRampDownSlope_Change()

    Dim TempD As Double
    
    If val(Me.txtRampDownSlope) > 0 Then
    
        'Calculate the Ramp Down duration (in millisecs)
        TempD = val(Me.txtMaxClipAmp) / val(Me.txtRampDownSlope) * 1000

        'Need to Downdate the Ramp Down duration label
        Me.lblRampDownDuration.Caption = PadLeft(Trim(str(CLng(TempD))), 7)

    End If
    
    Me.refresh

End Sub

Private Sub txtRampDownSlope_LostFocus()

    'No negative slopes allowed
    If val(Me.txtRampDownSlope) < 0 Then
    
        Me.txtRampDownSlope = Trim(str(-1 * val(Me.txtRampDownSlope)))
        
        'Trigger the change event
        txtRampDownSlope_Change
        
    End If

End Sub

Private Sub txtRampDownTime_LostFocus()

    Dim temp As Double

    temp = val(txtRampDownTime)
    
    If temp > 10000 Then txtRampDownTime = "10000"
    
    If temp < 0 Then txtRampDownTime = "0"
    
End Sub

Private Sub txtRampUpSlope_Change()

    Dim TempD As Double
    
    If val(Me.txtRampUpSlope) > 0 Then
    
        'Calculate the Ramp up duration (in millisecs)
        TempD = val(Me.txtMaxClipAmp) / val(Me.txtRampUpSlope) * 1000

        'Need to update the Ramp Down duration label
        Me.lblRampUpDuration.Caption = PadLeft(Trim(str(CLng(TempD))), 7)

    End If
    
    Me.refresh

End Sub

Private Sub txtRampUpSlope_LostFocus()

    'No negative slopes allowed
    If val(Me.txtRampUpSlope) < 0 Then
    
        Me.txtRampUpSlope = Trim(str(-1 * val(Me.txtRampUpSlope)))
        
        'Trigger the change event
        txtRampUpSlope_Change
        
    End If

End Sub

Private Sub txtRampUpTime_LostFocus()

    Dim temp As Double

    temp = val(txtRampUpTime)
    
    If temp > 10000 Then txtRampUpTime = "10000"
        
    If temp < 0 Then txtRampUpTime = "0"
    
End Sub

