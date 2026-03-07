VERSION 5.00
Begin VB.Form frmTransverseProbeAutoPosition 
   Caption         =   "Transverse Probe Auto-Position Routine"
   ClientHeight    =   6150
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11685
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6150
   ScaleWidth      =   11685
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H8000000C&
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   1095
      Left            =   7680
      MaskColor       =   &H8000000C&
      TabIndex        =   22
      Top             =   4800
      Width           =   3855
   End
   Begin VB.Frame frameResults 
      Caption         =   "Results"
      Height          =   1455
      Left            =   120
      TabIndex        =   6
      Top             =   4560
      Width           =   7335
      Begin VB.CommandButton cmdResetTurningAngle 
         Caption         =   "Relabel Best Angle as Zero Degrees"
         Height          =   375
         Left            =   4080
         TabIndex        =   21
         Top             =   840
         Width           =   3015
      End
      Begin VB.TextBox txtBestAngle 
         Height          =   285
         Left            =   2400
         TabIndex        =   20
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox txtPriorBestAngle 
         Height          =   285
         Left            =   2400
         TabIndex        =   19
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Best Angle (Current Scan):"
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label lblOldBestAngle 
         Caption         =   "Prior Best Angle:"
         Height          =   255
         Left            =   960
         TabIndex        =   16
         Top             =   480
         Width           =   1335
      End
   End
   Begin VB.Frame frameAFAutoTune 
      Caption         =   "Auto Angle Finder"
      Height          =   4335
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3735
      Begin VB.CommandButton cmdPauseAndResume 
         Caption         =   "Pause"
         Height          =   372
         Left            =   2280
         TabIndex        =   18
         Top             =   3720
         Width           =   1095
      End
      Begin VB.TextBox txtEndScanAngle 
         Height          =   288
         Left            =   2280
         TabIndex        =   9
         Top             =   1320
         Width           =   1092
      End
      Begin VB.TextBox txtStartScanAngle 
         Height          =   288
         Left            =   2280
         TabIndex        =   8
         Top             =   840
         Width           =   1092
      End
      Begin VB.TextBox txtCurrentAngle 
         Height          =   288
         Left            =   2280
         TabIndex        =   7
         Top             =   360
         Width           =   1092
      End
      Begin VB.CommandButton cmdAutoFindTransverseAngle 
         Caption         =   "Start Angle Scan"
         Height          =   372
         Left            =   360
         TabIndex        =   17
         Top             =   3720
         Width           =   1575
      End
      Begin VB.TextBox txtPeakAFHangTime 
         Height          =   288
         Left            =   2520
         TabIndex        =   13
         Top             =   2400
         Width           =   852
      End
      Begin VB.TextBox txtTargetAFMonitorVoltage 
         Height          =   288
         Left            =   2520
         TabIndex        =   15
         Top             =   3120
         Width           =   852
      End
      Begin VB.PictureBox picBluePixel 
         Height          =   12
         Left            =   2280
         Picture         =   "frmTransverseProbeAutoPosition.frx":0000
         ScaleHeight     =   0.027
         ScaleMode       =   0  'User
         ScaleWidth      =   0.027
         TabIndex        =   2
         Top             =   240
         Visible         =   0   'False
         Width           =   12
      End
      Begin VB.TextBox txtAngleStepSize 
         Height          =   288
         Left            =   2520
         TabIndex        =   11
         Top             =   1800
         Width           =   852
      End
      Begin VB.Label Label4 
         Caption         =   "End Angle (Degrees):"
         Height          =   255
         Left            =   360
         TabIndex        =   14
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "Start Angle (Degrees):"
         Height          =   255
         Left            =   360
         TabIndex        =   12
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Current Angle (Degrees):"
         Height          =   255
         Left            =   360
         TabIndex        =   10
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label lblScanDuration 
         Caption         =   "Peak AF Hange Time at Each Angle (ms):"
         Height          =   495
         Left            =   360
         TabIndex        =   5
         Top             =   2280
         Width           =   2055
      End
      Begin VB.Label lblAmplitude 
         Caption         =   "Target AF Monitor Voltage (0 - 5 V):"
         Height          =   495
         Left            =   360
         TabIndex        =   4
         Top             =   3000
         Width           =   1935
      End
      Begin VB.Label Label3 
         Caption         =   "Angle Step Size (Degrees):"
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   1800
         Width           =   1935
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
      Height          =   4335
      Left            =   3960
      MousePointer    =   2  'Cross
      ScaleHeight     =   9726.962
      ScaleMode       =   0  'User
      ScaleWidth      =   16036.42
      TabIndex        =   0
      Top             =   120
      Width           =   7575
   End
End
Attribute VB_Name = "frmTransverseProbeAutoPosition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public prior_best_angle As Double
Public current_best_angle As Double
Private WithEvents angle_finder As ProbeAngleOptimizer
Attribute angle_finder.VB_VarHelpID = -1

Const StartButtonCaption_StartScan As String = "Start Angle Scan"
Const StartButtonCaption_AbortScan As String = "End Angle Scan"
Const PauseButtonCaption_PauseScan As String = "Pause"
Const PauseButtonCaption_ResumeScan As String = "Resume"

Dim light_green As Long
Dim light_red As Long
Dim medium_grey As Long
Dim light_grey As Long

Dim status_msg As String
Dim num_steps As Integer

Dim XInterval As Long
Dim MaxAmps() As Long

Private Sub angle_finder_AngleScanStatusChange(ByVal angle_scan_status As AngleScanStatusEnum)

    Select Case angle_scan_status
    
        Case AngleScanStatusEnum.ABORTED
    
            Me.cmdAutoFindTransverseAngle.Caption = StartButtonCaption_StartScan
            Me.cmdAutoFindTransverseAngle.BackColor = light_green
            Me.cmdPauseAndResume.Caption = PauseButtonCaption_PauseScan
            Me.cmdPauseAndResume.BackColor = light_grey
                        
        Case AngleScanStatusEnum.IDLE
        
            Me.cmdAutoFindTransverseAngle.Caption = StartButtonCaption_StartScan
            Me.cmdAutoFindTransverseAngle.BackColor = light_green
            Me.cmdPauseAndResume.Caption = PauseButtonCaption_PauseScan
            Me.cmdPauseAndResume.BackColor = light_grey
            
        Case AngleScanStatusEnum.PAUSED
        
            Me.cmdPauseAndResume.Caption = PauseButtonCaption_ResumeScan
            Me.cmdPauseAndResume.BackColor = medium_grey
            modFlow.Flow_Pause
                
        Case AngleScanStatusEnum.RUNNING
        
            Me.cmdAutoFindTransverseAngle.Caption = StartButtonCaption_AbortScan
            Me.cmdAutoFindTransverseAngle.BackColor = light_red
            Me.cmdPauseAndResume.Caption = PauseButtonCaption_PauseScan
            Me.cmdPauseAndResume.BackColor = light_grey
            modFlow.Flow_Resume
        
    End Select

End Sub

Private Sub angle_finder_ProgressUpdate(ByVal current_angle As Double)

    UpdateScanProgress (current_angle)

End Sub

Private Function AreInputsValid() As Boolean

    Dim ret_val As Boolean: ret_val = True
    
    Dim concatenated_inputs As String
    
    concatenated_inputs = Trim(Me.txtAngleStepSize.text) & _
                          Trim(Me.txtCurrentAngle.text) & _
                          Trim(Me.txtEndScanAngle.text) & _
                          Trim(Me.txtPeakAFHangTime.text) & _
                          Trim(Me.txtStartScanAngle.text) & _
                          Trim(Me.txtTargetAFMonitorVoltage.text)
                          
    If Not IsNumeric(concatenated_inputs) Then
    
        ret_val = False
        error_str = "All inputs must be numeric values."
    
    End If
                         
    
    AreInputsValid = False
    
    Err.Raise -666, "AreInputsValid", error_str
    
End Function

Private Sub cmdAutoFindTransverseAngle_Click()
    If Me.cmdAutoFindTransverseAngle.Caption = StartButtonCaption_StartScan Then
        
        On Error GoTo Run_Scan_Error
    
        InitScan
        SetupDisplay
        RunScan
        'MockScan
        DisplayScanResults
        
        angle_finder.AngleScanStatus = AngleScanStatusEnum.IDLE
        angle_finder_AngleScanStatusChange AngleScanStatusEnum.IDLE
        
        On Error GoTo 0
    ElseIf Me.cmdAutoFindTransverseAngle.Caption = StartButtonCaption_AbortScan Then
        angle_finder.AngleScanStatus = AngleScanStatusEnum.ABORTED
        angle_finder_AngleScanStatusChange AngleScanStatusEnum.ABORTED
        
    End If
    
    Exit Sub

Run_Scan_Error:

    MsgBox "Error setting up or running the Transverse Probe Auto Angle Scan." & _
           vbCrLf & vbCrLf & _
           Err.Description, vbCritical, "Error!"
           
End Sub

Private Sub cmdClose_Click()
    Me.Hide
    Unload Me
End Sub

Private Sub cmdPauseAndResume_Click()
    If Me.cmdPauseAndResume.Caption = PauseButtonCaption_PauseScan Then
        angle_finder.AngleScanStatus = AngleScanStatusEnum.PAUSED
        angle_finder_AngleScanStatusChange (AngleScanStatusEnum.PAUSED)
    ElseIf Me.cmdPauseAndResume.Caption = PauseButtonCaption_ResumeScan Then
        angle_finder.AngleScanStatus = AngleScanStatusEnum.RUNNING
        angle_finder_AngleScanStatusChange (AngleScanStatusEnum.RUNNING)
    End If
End Sub

Private Sub cmdResetTurningAngle_Click()

    'Move to best angle
   On Error GoTo cmdResetTurningAngle_Click_Error

    frmDCMotors.TurningMotorRotate val(Me.txtBestAngle.text)
    
    'Relabel position as zero
    frmDCMotors.RelabelPos modMotor.MotorTurning, 0

   On Error GoTo 0
   Exit Sub

cmdResetTurningAngle_Click_Error:

    MsgBox "Error " & Err.number & " (" & Err.Description & ") in procedure cmdResetTurningAngle_Click of Form frmTransverseProbeAutoPosition"

End Sub

Private Sub DisplayScanResults()

    frmProgram.StatusBar "Plotting Data...", 3

    'Find range of difference between Biggest and Smallest Amplitudes
    Dim amp_interval As Double
    
    Dim min_angle As AngleVsField_Point
    Dim max_angle As AngleVsField_Point
    
    Set max_angle = angle_finder.GetPoint_WithMaxDCField
    Set min_angle = angle_finder.GetPoint_WithMinDCField
    
    amp_interval = Abs(max_angle.peak_field - min_angle.peak_field)

    If amp_interval = 0 Then
    'If amp_interval is zero, make it slightly non-zero

        amp_interval = 0.00001
        
    End If

    'Need to now find the rounding factor to use to divide Amp interval into
    'four easy to display numbers
    'NOTE:  If BiggestAmp < Smallest Amp, the code below will cause an error
    '       by taking the log of a negative number!!
    RoundingPower = Int(Log(amp_interval / 4) / Log(10))

    'Change Rounding Power so that it is now the number of places to
    'keep to the right of the decimal point
    If RoundingPower > 0 Then RoundingPower = 0
    RoundingPower = -1 * RoundingPower

    'Set plot font back to ten
    picDCResponse.FontSize = 10

    j = 0

    Dim field_str As String

    'Need to scale and label the Y-axis
    For i = 8000 To 2000 Step -1500

        picDCResponse.Line (1800, i)-(1950, i)  'Draw Vertical tick mark

        field_str = Trim(Str(Round(min_angle.peak_field + j * amp_interval / 4, RoundingPower)))
'        Debug.Print field_str
        j = j + 1

        'Now run loop to see how to fit the entire freq label in
        'the space available
        doContinue = False

        Do

            If picDCResponse.TextWidth(field_str) > 1700 Then

                'Cut Label into two pieces at the mid-point
                'and now check if the two pieces will fit
                If picDCResponse.TextWidth(field_str) > 2400 Then

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
                    picDCResponse.CurrentY = i - picDCResponse.TextHeight(field_str)
                    picDCResponse.Print Mid(field_str, 1, Len(field_str) \ 2)

                    'Second Piece
                    picDCResponse.CurrentX = 500
                    picDCResponse.CurrentY = i
                    picDCResponse.Print Mid(field_str, Len(field_str) \ 2 + 1)

                    doContinue = False

                End If

            Else

                'Freq String for label is small enough to fit in the allotted space
                'Plot the label
                picDCResponse.CurrentX = 1700 - picDCResponse.TextWidth(field_str)
                picDCResponse.CurrentY = i - picDCResponse.TextHeight(field_str) / 2

                picDCResponse.Print field_str

                doContinue = False

            End If

        Loop Until doContinue = False

    Next i

    'Now Draw in the columns for each Freq
    For i = 0 To angle_finder.angles_vs_fields.Count - 1

        picDCResponse.Line _
            (CLng(MaxAmps(i, 2) + 0.1 * XInterval), 8550)-( _
                CLng(MaxAmps(i, 3) - 0.1 * XInterval), _
                8000 - CLng(6000 / amp_interval * (angle_finder.angles_vs_fields(i + 1).peak_field - min_angle.peak_field))), _
            QBColor(1), _
            BF

    Next i

    'Update the program status bar, panel 3
    '(panel 2 will be blanked in the cmdAutoTuneAF_Click subroutine that called this function)
    frmProgram.StatusBar vbNullString, 3

End Sub

Private Sub Form_Load()

    Me.Height = 6720
    Me.Width = 11925

    If angle_finder Is Nothing Then
        Set angle_finder = New ProbeAngleOptimizer
        angle_finder.AngleScanStatus = AngleScanStatusEnum.IDLE
        angle_finder_AngleScanStatusChange (AngleScanStatusEnum.IDLE)
    End If

    On Error Resume Next
    Me.txtCurrentAngle.text = Format(frmDCMotors.TurningMotorAngle, "#0.0")
    On Error GoTo 0

    light_green = QBColor(10)
    light_red = QBColor(12)
    medium_grey = QBColor(8)
    light_grey = &HCCCCCC

    Me.cmdAutoFindTransverseAngle.Caption = StartButtonCaption_StartScan
    Me.cmdAutoFindTransverseAngle.BackColor = light_green
    
    Me.cmdPauseAndResume.Caption = PauseButtonCaption_PauseScan
    Me.cmdPauseAndResume.BackColor = light_grey
    
    If modConfig.AfTransResFreq <> 0 Then
    
        Me.txtPeakAFHangTime.text = Format(CLng(1 / modConfig.AfTransResFreq * 100000), "#0.0")
        
    Else
    
        Me.txtPeakAFHangTime.text = "350"
        
    End If
    
End Sub

Private Sub Form_Resize()
    Me.Height = 6720
    Me.Width = 11925
End Sub

Private Sub InitScan()

    If angle_finder Is Nothing Then Set angle_finder = New ProbeAngleOptimizer
    
    angle_finder.AF_PeakHangTime = val(Me.txtPeakAFHangTime.text)
    angle_finder.AF_TargetMonitorVoltage = val(Me.txtTargetAFMonitorVoltage.text)
    angle_finder.AngleScanCalibrationMode = AF
    angle_finder.AngleScanStatus = AngleScanStatusEnum.RUNNING
    angle_finder_AngleScanStatusChange (AngleScanStatusEnum.RUNNING)
    angle_finder.AngleStepSizeInDegrees = val(Me.txtAngleStepSize.text)
    angle_finder.EndAngle = val(Me.txtEndScanAngle.text)
    angle_finder.StartAngle = val(Me.txtStartScanAngle.text)

End Sub

Private Sub MockScan()

    If Not angle_finder.AreValidClassParameters Then
    
        Err.Raise -666, "MockScan", angle_finder.ErrorString
        
    End If
    
    Dim angle As Double
    Dim i As Integer
    Dim num_steps As Integer
    
    num_steps = CInt(Abs((angle_finder.EndAngle - angle_finder.StartAngle) / angle_finder.AngleStepSizeInDegrees)) + 1
    
    i = -num_steps \ 2 + 2
    
    angle_finder.angles_vs_fields.Clear
    
    Dim current_angle As Double
    
    For angle = angle_finder.StartAngle To angle_finder.EndAngle Step angle_finder.AngleStepSizeInDegrees
    
        frmDCMotors.TurningMotorRotate angle, True, True
        
        current_angle = frmDCMotors.TurningMotorAngle
    
        Dim mock_field As Double
        
        angle_finder_ProgressUpdate current_angle
        
        mock_field = Exp(-1 / 2 * (i / 10) ^ 2) * 100
        
        i = i + 1
    
        angle_finder.angles_vs_fields.Add current_angle, mock_field
        
        PauseTill timeGetTime() + 3000
            
    Next angle
    
    Dim dc_peak_field As Double
    dc_peak_field = angle_finder.GetAngleWithMaxDCField
    
    Me.txtPriorBestAngle.text = Me.txtBestAngle.text
    Me.txtBestAngle.text = Format(dc_peak_field, "#0.0")

End Sub

Private Sub RunScan()

    status_msg = "Scanning Angles ...."

    WriteStatusToDisplay status_msg

    Dim best_angle As Double
    best_angle = angle_finder.FindBestTransverseProbeAngle

    Me.txtPriorBestAngle.text = Me.txtBestAngle.text
    Me.txtBestAngle.text = Format(best_angle, "#0.0")
    
End Sub

Public Sub SetupDisplay()

    With angle_finder

        If Not .AreValidClassParameters Then
        
            Exit Sub
            
        End If
        
        num_steps = CInt(Abs((.EndAngle - .StartAngle) / .AngleStepSizeInDegrees)) + 1

        'Set Font Size
        picDCResponse.FontSize = 10
        
        'Clear Picture Box
        picDCResponse.Cls
           
        'Set Picture Object scale height and width properties
        picDCResponse.ScaleHeight = 10000
        picDCResponse.ScaleWidth = 14500

        'Draw The Bounds of the DC Response Peak Field Display Window
        picDCResponse.Line (1950, 1000)-(1950, 8550) 'Vertical axis
        picDCResponse.Line (1950, 8550)-(14500, 8550) 'Horizontal axis

        'Plot the units for the Y-axis
        picDCResponse.CurrentY = 200
        picDCResponse.CurrentX = 1950 - picDCResponse.TextWidth("Field (G)") / 2
        picDCResponse.Print "Field (G)"

        'Plot the label + units for the X-Axis
        picDCResponse.CurrentY = 8700 + CLng(1.5 * picDCResponse.TextWidth("0"))
        picDCResponse.CurrentX = 7750 - picDCResponse.TextWidth("Angle (Deg)")
        picDCResponse.Print "Angle (Deg)"

        'Calculate the amount of width each Angleuency has in the X-coordinate
        'space for plotting
        XInterval = CLng(12000 / num_steps)

        Dim SkipLabel As Boolean: SkipLabel = False

        'Lower font size for Angle column labels
        picDCResponse.FontSize = 9

        'Need to now find the rounding factor to use to divide Amp interval into
        'four easy to display numbers
        'NOTE:  If AngleStepSize < 0, the code below will cause an error
        '       by taking the log of a negative number!!
        RoundingPower = Int(Log(Abs(.AngleStepSizeInDegrees)) / Log(10))

        'Change Rounding Power so that it is now the number of places to
        'keep to the right of the decimal point
        If RoundingPower > 0 Then RoundingPower = 0
        RoundingPower = -1 * RoundingPower

        Dim angle_str As String
        
        ReDim MaxAmps(num_steps, 3)

        For i = 0 To num_steps - 1

            'calculate left and right positions
            MaxAmps(i, 2) = 2250 + i * XInterval  'Left position is the first possible
            MaxAmps(i, 3) = 2250 + (i + 1) * XInterval

            'Plot the X axis tick marks for this Angle
            picDCResponse.Line (MaxAmps(i, 2), 8550)-(MaxAmps(i, 2), 8750)
            picDCResponse.Line (MaxAmps(i, 3), 8550)-(MaxAmps(i, 3), 8750)

            'Plot the label for this Angle
            'Construct Angle String
            angle_str = Trim(Str(Round(.StartAngle + _
                                        (.EndAngle - .StartAngle) * i / (num_steps - 1), _
                                        RoundingPower)))

            Dim doContinue As Boolean: doContinue = False

            If SkipLabel = True Then

                SkipLabel = False

            Else

                 Do

                     'Check to see if the text Width of the Angle label is greater
                     'than the XInterval for each Angle
                     If picDCResponse.TextWidth(angle_str) > 0.8 * XInterval Then

                         'Not enough vertical space, lower the font size and
                         'repeat the label size check
                         picDCResponse.FontSize = picDCResponse.FontSize - 1

                         If picDCResponse.FontSize <= 8.25 Then

                             'Skip every other label
                             SkipLabel = True

                             'Plot this label
                             picDCResponse.CurrentX = CLng(XInterval / 2 _
                                                     - picDCResponse.TextWidth(angle_str) / 2) _
                                                    + MaxAmps(i, 2)
                             picDCResponse.CurrentY = 8700
                             picDCResponse.Print angle_str

                             doContinue = False

                         Else

                            doContinue = True

                        End If

                     Else

                         'There's enough room to plot the Angle label horizontally
                         picDCResponse.CurrentX = CLng(XInterval / 2 _
                                                         - picDCResponse.TextWidth(angle_str) / 2) _
                                                 + MaxAmps(i, 2)

                         picDCResponse.CurrentY = 8700

                         picDCResponse.Print angle_str

                         doContinue = False

                     End If

                Loop Until doContinue = False

            End If

        Next i

    End With

    frmAFTuner.refresh

End Sub

Private Sub UpdateScanProgress(ByVal current_angle As Double)

    Dim angle_str As String
    angle_str = Format(current_angle, "#0.0")
    
    Me.txtCurrentAngle.text = angle_str
    
    status_msg = "Scanning Angles ...." & vbCrLf & _
                 "Current Angle: " & angle_str
                 
    WriteStatusToDisplay status_msg
    
    Me.refresh

End Sub

Private Sub WriteStatusToDisplay(ByVal status_msg As String)

    If Trim(status_msg) = "" Then

        'Overwrite last posting with a white box
        picDCResponse.Line (5000, 0)-(14500, 3000), _
                            vbWhite, _
                            BF
                            
    Else

        'Overwrite last posting with a white box
        picDCResponse.Line (5000, 0)-(14500, 3000), _
                            vbWhite, _
                            BF

        Dim print_lines() As String

        On Error GoTo LabelDisplay_Error

        print_lines() = Strings.Split(status_msg, vbCrLf)


        For i = 0 To UBound(print_lines)
            
            'UpDate plot window with status
            picDCResponse.CurrentX = 5000
            picDCResponse.CurrentY = 500 + CInt(picDCResponse.TextHeight(print_lines(i)) * 1.5) * i
            picDCResponse.FontSize = 8
            picDCResponse.ForeColor = vbBlack
            picDCResponse.Print print_lines(i)

        Next i
        
        On Error GoTo 0
        
        
    End If

    Exit Sub
    
LabelDisplay_Error:

End Sub

