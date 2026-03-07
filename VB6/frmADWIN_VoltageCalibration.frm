VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmADWIN_VoltageCalibration 
   Caption         =   "ADWIN AF Voltage Calibration"
   ClientHeight    =   6240
   ClientLeft      =   15840
   ClientTop       =   4575
   ClientWidth     =   5895
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6240
   ScaleWidth      =   5895
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   375
      Left            =   2160
      TabIndex        =   5
      Top             =   5760
      Width           =   1335
   End
   Begin VB.CheckBox chkVoltCalSuccessful 
      Caption         =   "Check1"
      Height          =   255
      Left            =   2280
      TabIndex        =   4
      Top             =   5880
      Visible         =   0   'False
      Width           =   375
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridVoltageCalibration 
      Height          =   2055
      Left            =   120
      TabIndex        =   3
      Top             =   3600
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   3625
      _Version        =   393216
      SelectionMode   =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.PictureBox picVoltageCalibration 
      BackColor       =   &H8000000E&
      Height          =   3375
      Left            =   120
      ScaleHeight     =   10000
      ScaleMode       =   0  'User
      ScaleWidth      =   14500
      TabIndex        =   2
      Top             =   0
      Width           =   5655
   End
   Begin VB.CommandButton cmdStartCal 
      Caption         =   "Start Calibration"
      Height          =   375
      Left            =   4200
      TabIndex        =   1
      Top             =   5760
      Width           =   1575
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   5760
      Width           =   1335
   End
End
Attribute VB_Name = "frmADWIN_VoltageCalibration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public CoilsToCalibrate As String
Public InitialCoilSystem As String
Dim MonitorWave As Wave
Dim CoilString As String

Private Sub SaveVoltageCalibrations()

    Dim i As Long
    
    If ActiveCoilSystem = AxialCoilSystem Then
    
        modAF_DAQ.SaveGridToArray Me.gridVoltageCalibration, _
                                  modConfig.AFRampAxial, _
                                  1, _
                                  Me.gridVoltageCalibration.Rows - 1, _
                                  1, _
                                  2
                                  
        
    Else
    
        modAF_DAQ.SaveGridToArray Me.gridVoltageCalibration, _
                                  modConfig.AFRampTrans, _
                                  1, _
                                  Me.gridVoltageCalibration.Rows - 1, _
                                  1, _
                                  2
                                          
    End If

End Sub

Private Sub cmdApply_Click()

    'Save the values in the flex-grid to the appropriate global array(s)
    SaveVoltageCalibrations

    'Check to see if all the needed calibrations are done
    If CoilsToCalibrate = "None" Then
    
        'Set the calibration successful check-box to checked
        Me.chkVoltCalSuccessful.Value = Checked
    
        'Reset the active coil to the original coil system set
        'when this form was loaded
        If InitialCoilSystem = AxialCoilSystem Then
        
            frmADWIN_AF.optCoil_Click (0)
            ActiveCoilSystem = AxialCoilSystem
            
        Else
        
            frmADWIN_AF.optCoil_Click (1)
            ActiveCoilSystem = TransverseCoilSystem
            
        End If
        
        'Close this form
        Me.Hide
        
    Else
    
        'Run the next coil calibration
        RunCalibration
        
    End If

End Sub

Private Sub cmdCancel_Click()

    'Set the calibration successful check-box to unchecked
    Me.chkVoltCalSuccessful.Value = Unchecked

    'Hide this form
    Me.Hide


End Sub


Private Sub cmdStartCal_Click()

    'Do the calibration
    RunCalibration

End Sub

Private Sub Form_Load()

    'Set form height and width
    Me.Height = 6750
    Me.Width = 6015

    'Start check-box off as uncalibrated
    Me.chkVoltCalSuccessful.Value = Unchecked

    'Save the current active coil system
    InitialCoilSystem = ActiveCoilSystem

    'Check the AFRampAxial and AFRampTrans count values in modconfig
    'to see which coils need to be calibrated
    If modConfig.AFRampAxialCount = 0 And _
       modConfig.AFRampTransCount = 0 _
    Then
    
        CoilsToCalibrate = "Both"
        
    ElseIf modConfig.AFRampAxialCount = 0 Then
    
        CoilsToCalibrate = "Axial"
        
    ElseIf modConfig.AFRampTransCount = 0 Then
    
        CoilsToCalibrate = "Transverse"
        
    Else
    
        'No coils to calibrate
        CoilsToCalibrate = "None"
        
    End If

    'Load the coilstring, too
    If ActiveCoilSystem = AxialCoilSystem Then
    
        CoilString = "Axial"
        
    Else
    
        CoilString = "Transverse"
        
    End If

    'Clear the picture box
    Me.picVoltageCalibration.Cls
    
    'Clear the flex grid and reload it with headers
    ClearAndReloadGrid Me.gridVoltageCalibration
    
    'Set Caption on Pause button
    Me.cmdStartCal.Caption = "Start Calibration"

    'Now import the Ramp Vs Monitor Voltage calibration values
    ImportCalibrationTable
    
    'If CoilsToCalibrate = None, then
    'click the cmdApply button
    If CoilsToCalibrate = "None" Then
    
        cmdApply_Click

    End If

End Sub

Private Sub ImportCalibrationTable()

    Dim i As Long
    Dim TempL As Long
    
    'Select for the correct coil to display
    If ActiveCoilSystem = AxialCoilSystem Then
    
        'If number of calibration pairs > 0
        'then display them
        If modConfig.AFRampAxialCount > 0 Then
        
            'need to display these values now
            modAF_DAQ.LoadArrayToGrid Me.gridVoltageCalibration, _
                                      modConfig.AFRampAxial, _
                                      Me, _
                                      1, _
                                      0, _
                                      True
        
        End If
    
    Else
    
        'If number of calibration pairs > 0
        'then display them
        If modConfig.AFRampTransCount > 0 Then
        
            'need to display these values now
            modAF_DAQ.LoadArrayToGrid Me.gridVoltageCalibration, _
                                      modConfig.AFRampTrans, _
                                      Me, _
                                      1, _
                                      0, _
                                      True
                                      
        End If
        
    End If
    
End Sub

Private Sub ClearAndReloadGrid(ByRef gridObj As MSHFlexGrid)

    gridObj.ClearStructure
    gridObj.Clear

    'Reload the headers
    With gridObj
    
        .Rows = 2
        .Cols = 3
        .FixedRows = 1
        .FixedCols = 1
        
        .row = 0
        .Col = 1
        .text = "Ramp Peak Volts"
        .ColWidth(1) = Me.TextWidth(.text) * 1.2
        
        .Col = 2
        .text = "Monitor Peak Volts"
        .ColWidth(2) = Me.TextWidth(.text) * 1.2
        
    End With
    
End Sub

Private Sub RunCalibration()

    Dim CurrentRow As Long
    Dim PriorCurrentRow As Long
    Dim RampVolt As Double
    Dim MaxRampVolt As Double
            
    'Check to see what coils need to be calibrated
    If Me.CoilsToCalibrate = "Both" Or _
       Me.CoilsToCalibrate = "Axial" _
    Then
    
        frmADWIN_AF.optCoil_Click (0)
        MaxRampVolt = modConfig.AfAxialRampMax
        
    Else
    
        frmADWIN_AF.optCoil_Click (1)
        MaxRampVolt = modConfig.AfTransRampMax
        
    End If
            
    'Going to now do several clip-tests
    '1) up to 0.2 Ramp volts
    '2) up to 0.5 Ramp volts
    '3) up to 1 Ramp volts
    '4) up to Max Ramp voltage
    
    'Check to see what the Max Ramp voltage is
    If MaxRampVolt <= 0 Then
    
        'Tell user that they need to do a normal
        'clip test first
        'Do a code yellow email + pop-up
        frmSendMail.MailNotification CoilString & " coil clip-test not done yet!", _
                                     "Please do a clip-test using the AF Tuner prior to running " & _
                                     "the ADWIN Voltage calibration." & vbNewLine & vbNewLine & _
                                     "Code execution has been paused.", _
                                     CodeYellow, _
                                     True
               
        'Load & Open the AF Tuner
        Load frmAFTuner
        frmAFTuner.Show
        
        'Click the cancel button
        cmdCancel_Click
        
        DoEvents
        
        Exit Sub
        
    End If
    
    'Start CurrentRow = 0
    CurrentRow = 0
    
    'Set the voltage to use
    RampVolt = 0.2
    
    'Do the 1st clip-test and plot the results into the picture box and flex-grid
    DoRampVSMonClipTest RampVolt, CurrentRow
    
    'Save the prior current row
    PriorCurrentRow = CurrentRow
    
    'Set the voltage to use
    RampVolt = 0.5
    
    'Do the 2nd clip-test and plot the results into the picture box and flex-grid
    DoRampVSMonClipTest RampVolt, CurrentRow
        
    'Reconcile the two clip test sets
    ReconcileMonClipTests Me.gridVoltageCalibration, _
                          CurrentRow, _
                          PriorCurrentRow
        
    'Set the voltage to use
    RampVolt = 1
    
    'Do the 3rd clip-test and plot the results into the picture box and flex-grid
    DoRampVSMonClipTest RampVolt, CurrentRow
    
    'Set the voltage to use
    RampVolt = MaxRampVolt
    
    'Do the 4th clip-test and plot the results into the picture box and flex-grid
    DoRampVSMonClipTest RampVolt, CurrentRow
    
End Sub

Public Sub ReconcileMonClipTest(ByRef gridObj As MSHFlexGrid, _
                                ByVal LastRow, _
                                ByVal PriorLastRow)

    Dim PriorArray() As Double
    Dim CurrentArray()
    
    'This function goes through and substitutes
    'the lowest monitor voltage values for a given ramp
    'value between all the entries in gridObj
    
    'First load the prior run gridobject data rows to an array
    modAF_DAQ.SaveGridToArray gridObj, _
                              PriorArray(), _
                              1, _
                              PriorLastRow, _
                              1, _
                              2
                              
    'Then load the most recently run grid object data rows to an array
    modAF_DAQ.SaveGridToArray gridObj, _
                              CurrentArray(), _
                              PriorLastRow + 1, _
                              LastRow, _
                              1, _
                              2
                              
    'Now need to splice the second array into the first
    'and replace the higher monitor voltage values with corresponding lower voltage values
    modAF_DAQ.ArrayInterpolateAndSplice PriorArray, _
                                        CurrentArray, _
                                        True, _
                                        False
    

End Sub

Private Sub DoRampVSMonClipTest(ByVal RampVolt As Double, _
                                ByRef CurrentRow As Long)

    Dim i As Long
    Dim TempL As Long
    Dim N As Long
    Dim M As Long
    
    Dim AFData() As Double
    Dim SineFit_Data() As Double
    Dim MonitorAmp() As Double
    Dim TempD As Double
    
    'Set font-size on picture box
    Me.picVoltageCalibration.FontSize = 9
    
    'Print Clipping Test Ramp status
    picVoltageCalibration.CurrentX = 4000
    picVoltageCalibration.CurrentY = 4000
    picVoltageCalibration.Print CoilString & " coil voltage calibration:"
    
    'Change text display start point again
    picVoltageCalibration.CurrentX = 6000
    picVoltageCalibration.CurrentY = 4000 + 1.5 * picVoltageCalibration.TextHeight(CoilString)
    picVoltageCalibration.Print "Ramping.... " & Trim(Str(RampVolt)) & " Volts"
    
    'Refresh form
    picVoltageCalibration.refresh
    Me.refresh
    
    'Config the 1st Clip test
    ConfigClipTest RampVolt
    
    'Else run the first clip test
    frmADWIN_AF.DoRampADWIN MonitorWave, _
                            WaveForms("AFRAMPUP"), _
                            WaveForms("AFRAMPDOWN"), _
                            AFData, _
                            1, , _
                            3
        
    'Clear the picture box
    Me.picVoltageCalibration.Cls
        
    'Update Ramping Status
    picVoltageCalibration.CurrentX = 6000
    picVoltageCalibration.CurrentY = 4000 + 1.5 * picVoltageCalibration.TextHeight(CoilString)
    picVoltageCalibration.Print "Ramping.... Done"
    
    'Change the text cursor position
    picVoltageCalibration.CurrentX = 6000
    picVoltageCalibration.CurrentY = 4000 + 3 * picVoltageCalibration.TextHeight(CoilString)
    picVoltageCalibration.Print "Analyzing..."
    
    'Refresh form
    picVoltageCalibration.refresh
    Me.refresh
    
    'Now sine fit the data
    '20 points going up, 20 points coming down
        
    'First get the total number of data rows in AFData
    N = UBound(AFData, 1)
    
    'Divide N by 20 to get the points between fits
    frmADWIN_AF.DoSineFitAnalysis MonitorWave, _
                                  AFData, _
                                  SineFit_Data, _
                                  1, _
                                  CLng(N / 20)
            
    'Plot the SineFit Data using the code in AF Tuner for plotting clip-tests
    frmAFTuner.PlotAutoClipTestResults SineFit_Data, _
                                       Me.picVoltageCalibration, _
                                       0, _
                                       RampVolt, _
                                       WaveForms("AFRAMPUP").CurrentPoint
                                                       
    Me.picVoltageCalibration.refresh
    Me.refresh
                                                       
    'Now take the lower of nearby sine-fit amplitudes
    'and interpolate the matching ramp voltages into the
    'MonAmp array
    GetMonitorMatches SineFit_Data, _
                      MonitorAmp
    
    'Get the size of Monitor Amp
    M = UBound(MonitorAmp, 1)
    
    'Clear & reload the grid with headers
    ClearAndReloadGrid Me.gridVoltageCalibration
        
    'Load the results to the table
    modAF_DAQ.LoadArrayToGrid Me.gridVoltageCalibration, _
                              MonitorAmp, _
                              Me, _
                              1, _
                              0, _
                              True, _
                              "#0.0000"
            
    'Pause for 1 second
    PauseTill timeGetTime() + 1000
    
    'Update Current row
    CurrentRow = CurrentRow + M

End Sub

Private Sub GetMonitorMatches(ByRef SineFit_Data() As Double, _
                              ByRef MonitorAmp() As Double)
                              
    Dim i As Long
    Dim j As Long
    Dim N As Long
    Dim M As Long
    Dim Slope As Double
    Dim MonVolt As Double
                              
    Dim PeakReached As Boolean
                                  
    'Start PeakReached at zero
    PeakReached = False
                                  
    'Get the number of data-rows in the SineFit_Data array
    N = UBound(SineFit_Data, 1)
    
    'Run through the SineFit_Data array and find
    'the point of the highest sine_fit output voltage
    For i = 0 To N - 1
    
        'If this isn't the first point, we can check and
        'see if the peak has been reached
        If i > 0 And PeakReached = False Then
    
            'Check to see if the current point is less than the last point
            If SineFit_Data(i, 9) <= SineFit_Data(i - 1, 9) Then
            
                'Set peak reached flag = True
                PeakReached = True
                
                'Store the position of the last up-going point
                M = i - 1
                
                'exit the for loop
                Exit For
                
            End If
            
        End If
        
    Next i
    
    'Redimension the MonitorAmp array based on M
    ReDim MonitorAmp(M, 2)
    
   
    'Loop through the Sine Fit data rows
    For i = 0 To N - 1
    
        'If this isn't the first point, we can check and
        'see if the peak has been reached
    
        If i < M Then
            
            'Right in the current Ramp output voltage
            'and corresponding sine-fitted monitor input voltage
            MonitorAmp(i, 0) = SineFit_Data(i, 9)
            MonitorAmp(i, 1) = SineFit_Data(i, 3)
            
        Else
            
            'Peak has been reached
            'Now need to compare most next recent
            'elements of MonitorAmp to the new ramp & monitor pair
            'in the SineFit_Data array
            
            'Check to make sure that N - i - 1 >= 0
            If N - i - 1 >= 0 Then
                
                'Is current Sine Fit ramp voltage higher than
                'the next lowest Monitor Amp ramp voltages?
                If MonitorAmp(N - i - 1, 0) < SineFit_Data(i, 9) Then
                
                    'Can interpolate to compare the two monitor voltages
                    Slope = (MonitorAmp(N - i, 0) - MonitorAmp(N - i - 1, 0)) / _
                            (MonitorAmp(N - i, 1) - MonitorAmp(N - i - 1, 1))
                            
                    'Get the matching monitor voltage
                    MonVolt = Slope * (SineFit_Data(i, 9) - MonitorAmp(N - i - 1, 0)) + _
                              MonitorAmp(N - i - 1, 1)
                    
                    'Now compare this monitor voltage to the current SineFit monitor voltage
                    'And keep the lower of the two
                    If MonVolt > SineFit_Data(i, 3) Then
                    
                        MonVolt = SineFit_Data(i, 3)
                    
                        'Save new Monitor voltage / ramp voltage pair over the old value
                        MonitorAmp(N - i, 0) = SineFit_Data(i, 0)
                        MonitorAmp(N - i, 1) = MonVolt
                        
                    Else
                    
                        'Do nothing, need to keep original value
                        
                    End If
                    
                ElseIf MonitorAmp(N - i, 0) = SineFit_Data(i, 9) Then
                
                    'Just flat out compare the Monitor voltages for the two
                    'And keep the lower of the two
                    If MonitorAmp(N - i, 0) > SineFit_Data(i, 3) Then MonitorAmp(N - i, 0) = SineFit_Data(i, 3)
                    
                End If
                              
            End If
                              
        End If
        
    Next i
                              
End Sub
                          

Public Sub ConfigClipTest(ByVal RampPeakVoltage As Double)

    Dim Freq As Double
    
    'Check to see which coil we're using to get the sine-freq
    If ActiveCoilSystem = AxialCoilSystem Then

        Freq = modConfig.AfAxialResFreq
        CoilString = "Axial"
        
    Else
    
        Freq = modConfig.AfTransResFreq
        CoilString = "Transverse"
        
    End If
    
    'Check to see if Freq is non-zero
    If Freq <= 0 Then
    
        'This coil has not yet been tuned
        'Do a code yellow email + pop-up
        frmSendMail.MailNotification "AF " & CoilString & " coil not tuned!", _
                                     "The ADWIN Voltage calibration cannot proceed until " & _
                                     "The " & CoilString & " coil has been calibrated." & _
                                     vbNewLine & vbNewLine & _
                                     "Code execution has been paused.", _
                                     CodeYellow, _
                                     True
                                     
        'Load & Open the AF Tuner
        Load frmAFTuner
        frmAFTuner.Show
        
        'Click the cancel button
        cmdCancel_Click
        
        DoEvents
        
        Exit Sub
        
    End If
    
    'We know tuning has been done, now
    
    With WaveForms("AFRAMPUP")
    
        .PeakVoltage = RampPeakVoltage
        .SineFreqMin = Freq
        .Slope = RampPeakVoltage / (1000 / Freq)
        .TimeStep = 1 / .IORate
        
    End With
    
    With WaveForms("AFRAMPDOWN")
    
        .PeakVoltage = RampPeakVoltage
        .SineFreqMin = Freq
        .Slope = RampPeakVoltage / (1000 / Freq)
        .TimeStep = 1 / .IORate
        
    End With
    
    'Set Monitor Wave to the High-Field AF monitor waveform
    Set MonitorWave = WaveForms("AFHFMONITOR")
    
    With MonitorWave
    
        .PeakVoltage = 1
        .SineFreqMin = Freq
        .TimeStep = 1 / .IORate
        
    End With
    
End Sub

