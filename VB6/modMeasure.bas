Attribute VB_Name = "modMeasure"
' Sample Routines
'
' handle everything from the movement and the selection
' the positioning, the measurement and the demag steps!
Option Explicit  ' enforce variable declaration!
' Time delay after ARC squid box command
Const Measure_ARCDelay = 2.5
' Sample Orientation
Global Const Magnet_SampleOrientationUp As Integer = -1
Global Const Magnet_SampleOrientationDown As Integer = 1
Type Measure_Unfolded
    S As Angular3D     ' Bedding coordinates
    c As Angular3D     ' Core coordinates
    g As Angular3D     ' Geog. coordinates
End Type
Public Type Measure_AvgStats
    ' This type is used to pass information from measurement
    ' of the sample to display on the screen, and output to
    ' a data file
    unfolded   As Measure_Unfolded  ' Unfolded measurement
    SigNoise   As Double            ' Signal/Noise ratio
    SigHolder  As Double            ' Signal/Holder ratio
    SigInduced As Double            ' Signal/Induced ratio
    momentvol  As Double            ' Moment/vol ratio
End Type
Public HolderMeasured As Boolean    ' Has the holder been measured?
                                    ' (frmMagnetometerControl.cmdManHolder)
Public modeFluxCount As Boolean     ' Flux Counting mode.  (Unimplemented)
Public Holder As MeasurementBlock     ' Holder measurements

Public Function Format5Char(d As Double)
    ' This function formats the given number to a 5 character string
    If d >= 100 Then
        Format5Char = Format(d, "000.0")
    ElseIf d >= 10 Then
        Format5Char = " " + Format(d, "00.0")
    End If
End Function

Public Function Measure_Bedding(specimen As Sample, _
    ByVal ginc As Double, ByVal gdec As Double) As Angular3D
    '  Subroutine to make the structural and fold corrections.
    '  This uses the strike of bedding, not the dip direction, given in
    '  a right-handed sense. If fold corrections are also going to be done,
    '  both the remanence direction and a normal vector to the local bedding
    '  planes are rotated such that the fold axis is horizontal. The new
    '  bedding direction is then used to tilt-correct the rotated remanence
    '  direction to the final structurally-corrected orientation, SDEC and
    '  SINC.
    Dim bA As Double, bD As Double, magDec As Double
    With specimen
    magDec = .Parent.magDec
    If Not .FoldRotation Then
        ' Do the simple garden-variety bedding correction.
        bA = DegToRad(.BeddingStrike + magDec + 90)
        bD = DegToRad(.BeddingDip)
        Set Measure_Bedding = Measure_Rotate(ginc, gdec, bA, bD)
    Else
        Dim inc As Double, dec As Double
        Dim firstval As Angular3D, secondval As Angular3D
        ' First, rotate the remanence direction through the amount
        ' necessary to make the fold axis horizontal.
        bA = DegToRad(.FoldAxis + magDec)
        bD = DegToRad(.FoldPlunge)
        Set firstval = Measure_Rotate(ginc, gdec, bA, bD)
        ' Now we must find the new orientation of the bedding planes after
        ' rotating the fold axis up to horizontal.  To do this, rotate the
        ' direction of the normal vector through the same matrix.  The
        ' DEC and INC values calculated in the next two statements should
        ' be a normal vector to the untilted plane, and the DEC and INC
        ' returned from ROTATE should correspond to the fold-corrected
        ' plane direction.
        With specimen
            dec = .BeddingStrike + magDec - 90#
            inc = 90# - .BeddingDip
        End With
        Set secondval = Measure_Rotate(inc, dec, bA, bD)
        ' Now we need to take the new normal vector to the bedding plane,
        ' DEC,INC, and compute the dip direction (BA) and plunge (BD,
        ' both in radians).  We can then re-generate the rotation matrix
        ' and finish the tilt correction process on SDEC and SINC.  Note
        ' that the rotated directions are given with respect to TRUE NORTH,
        ' and only the measurements taken in the field (with the Magnetic
        ' Declination offset on the compass set at ZERO) require the MAGDEC
        ' correction.
        bA = DegToRad((secondval.dec + 180#))
        bD = DegToRad(90# - secondval.inc)
        ' Put the intermediate direction back for the final rotation
        Set Measure_Bedding = Measure_Rotate(firstval.inc, firstval.dec, bA, bD)
    End If
    End With
End Function

Public Function Measure_CalcStats(specimen As Sample, measblock As MeasurementBlocks) As Measure_AvgStats
    Dim vect As Cartesian3D
    Dim workingVol As Double
    Set vect = measblock.VectAvg
    If specimen.Vol > 0 Then workingVol = specimen.Vol Else workingVol = 1
    Measure_CalcStats.momentvol = measblock.Moment / workingVol
    ' Generate signal/noise ratios - Large values imply good data
    Measure_CalcStats.SigNoise = measblock.SigNoise
    Measure_CalcStats.SigHolder = measblock.SigHolder
    Measure_CalcStats.SigInduced = measblock.SigInduced
    '  Calls subroutine to complete all unfolding, to give a
    '  structurally corrected declination and inclination.
    Measure_CalcStats.unfolded = Measure_Unfold(specimen, _
        vect.X, vect.Y, vect.Z, Sqr(vect.X ^ 2 + vect.Y ^ 2))
    ' (February 2010 L Carporzen) Measure the TAF and uncorrect them in sample file
    If Left$(specimen.Parent.measurementSteps.CurrentStep.DemagStepLabel, 5) = "aTAFX" Then Measure_CalcStats.unfolded = Measure_Unfold(specimen, -vect.Z, vect.X, -vect.Y, Sqr(vect.Z ^ 2 + vect.X ^ 2))
    If Left$(specimen.Parent.measurementSteps.CurrentStep.DemagStepLabel, 5) = "aTAFY" Then Measure_CalcStats.unfolded = Measure_Unfold(specimen, vect.X, vect.Z, -vect.Y, Sqr(vect.X ^ 2 + vect.Z ^ 2))
    Set vect = Nothing
End Function

Sub Measure_QueryLoad(sampname As String, SampleOrientation As Integer)
    ' Request the user to manually load the sample with 'SampName'
    ' into the magnetometer and set it to the correct orientation.
    ' We then automatically load frmMeasure and call the 'MeasureSample'
    ' routines
    Dim QueryStr As String
    If (sampname <> SampleNameCurrent) Then
        ' We changed the sample, time to re-measure
        Magnetometer_UnloadSample
        If (sampname <> "Holder") Then
            frmVacuum.ValveConnect True
            QueryStr = "Please load the sample " + sampname
            If (SampleOrientation = Magnet_SampleOrientationUp) Then
                QueryStr = QueryStr + " with the arrow pointing up."
            Else
                QueryStr = QueryStr + " with the arrow pointing down."
            End If
 '           Motor_WaitStop ("UPDOWN")         ' Wait for motor to stop?
            MsgBox QueryStr, vbOKOnly, "Load Sample..."
        Else
            MsgBox "Please remove sample from holder", vbOKOnly, _
                "Remove Sample..."
        End If
        SampleNameCurrent = sampname                 ' ?
        SampleOrientationCurrent = SampleOrientation   ' ?
    ElseIf (SampleOrientation <> SampleOrientationCurrent) Then
        MsgBox "Please turn the sample over.", vbOKOnly, "Flip Sample..."
        SampleOrientationCurrent = SampleOrientation   ' ?
    End If
End Sub

Sub Measure_Read(targetSample As Sample, _
    RMStep As RockmagStep, _
    Optional RockmagMode As Boolean = False)
    ' This function starts the averaging cycles for the up
    ' direction.  It measures two +X, two -X, two +Y, two -y,
    ' and four +Z components
    Dim i As Integer
    Dim j As Integer
    Dim readDats As MeasurementBlocks
    Dim unfolded As Measure_Unfolded
    Set readDats = New MeasurementBlocks
    Dim IsHolder As Boolean
    Dim isUp As Boolean
    Dim doBoth As Boolean
    Dim numAvgSteps As Integer
    Dim curDemag As String
    Dim sdvect As Cartesian3D
    IsHolder = (targetSample.Samplename = "Holder")
    If DEBUG_MODE Then frmDebug.Msg "Reading: " + targetSample.Samplename + " isUp: " + Str(isUp) + " doBoth: " + Str(doBoth)
    SampleNameCurrent = targetSample.Samplename
    SampleStepCurrent = RMStep.DemagStepLabelLong
    Dim vectUMag   As Double                  ' Magnitude of "Up" vector
    Dim vectDMag   As Double                  ' Magnitude of "Down" vector
    Dim fInd As Double
    Dim UpToDn As Double
    Dim ErrorAngle As Double, errorHoriz As Double
    Dim FilePath As String
    Dim filepathbackup As String
    Dim avstats As Measure_AvgStats
    Dim avg As Cartesian3D
    Dim msgret As VbMsgBoxResult
    ' Initialize variables
    If Prog_halted Then Exit Sub
    If IsHolder Then
        ' Do initializations necessary for holder
        Set Holder = Nothing
        Set Holder = New MeasurementBlock
        frmSQUID.ChangeRange "A", "1" ' 1x read mode
        numAvgSteps = SampQueue.maxAvgSteps
    Else
        With targetSample.Parent
            isUp = .doUp
            doBoth = .doBoth
            curDemag = .curDemag
            numAvgSteps = .avgSteps
            If numAvgSteps < 1 Then numAvgSteps = 1
        End With
    End If
    ' Begin
    For i = 1 To numAvgSteps
        '  Do the initial zero measurement here
        readDats.Add Measure_ReadSample(targetSample, IsHolder, isUp)
        For j = 1 To 4
            readDats.Last.SetHolder j, Holder.Sample(j)
        Next j
        readDats.Last.isUp = isUp
        Set avg = readDats.VectAvg
        unfolded = Measure_Unfold(targetSample, avg.X, avg.Y, avg.Z)
        ' (February 2010 L Carporzen) Measure the TAF and uncorrect them in sample file
        If Left$(targetSample.Parent.measurementSteps.CurrentStep.DemagStepLabel, 5) = "aTAFX" Then unfolded = Measure_Unfold(targetSample, -avg.Z, avg.X, -avg.Y)
        If Left$(targetSample.Parent.measurementSteps.CurrentStep.DemagStepLabel, 5) = "aTAFY" Then unfolded = Measure_Unfold(targetSample, avg.X, avg.Z, -avg.Y)
        frmMeasure.ShowStats avg.X, avg.Y, avg.Z, unfolded.S.dec, unfolded.S.inc, _
                             readDats.SigDrift, _
                             readDats.SigHolder, _
                             readDats.SigInduced, _
                             readDats.FischerSD
        If IsHolder Then frmMeasure.AveragePlotEqualArea unfolded.S.dec, unfolded.S.inc, readDats.FischerSD ' (August 2007 L Carporzen) Equal area plot
        Set avg = Nothing
    Next i
    If Prog_halted Then Exit Sub ' (September 2007 L Carporzen) New version of the Halt button
    ' Now we've done the measurements the avgSteps number of times
    If IsHolder Then
        Set Holder = readDats.AverageBlock
    Else
        ' Not a holder measurement
        If (isUp And doBoth) Then
            ' We've measured the up direction, so save it to a temp file and leave
            targetSample.WriteUpMeasurements readDats, curDemag
            If DumpRawDataStats Then targetSample.WriteStatsTable readDats, curDemag
            avstats = Measure_CalcStats(targetSample, readDats)
            Set sdvect = readDats.VectSD
            frmStats.ShowErrors readDats.FischerSD, 0, 0
            frmStats.ShowAvgStats sdvect.X, sdvect.Y, sdvect.Z, _
                avstats.unfolded.c.dec, avstats.unfolded.c.inc, _
                avstats.unfolded.g.dec, avstats.unfolded.g.inc, _
                avstats.unfolded.S.dec, avstats.unfolded.S.inc, _
                avstats.momentvol, avstats.SigNoise, _
                avstats.SigHolder, avstats.SigInduced
            Set sdvect = Nothing
            Exit Sub
        End If
        If doBoth And Not isUp Then
            readDats.Assimilate targetSample.ReadUpMeasurements
            UpToDn = readDats.UpToDown
        End If
        ErrorAngle = readDats.FischerSD
        ' THE HORIZONTAL ERROR ANGLE, EH, IS NEGATIVE IF HOLDER SHOULD BE
        ' ROTATED TO THE LEFT, AND POSITIVE IF IT SHOULD GO TO THE RIGHT
        errorHoriz = readDats.ErrorHorizontal
        frmStats.ShowErrors ErrorAngle, errorHoriz, UpToDn
        Set sdvect = readDats.VectSD
        avstats = Measure_CalcStats(targetSample, readDats)
        frmStats.ShowAvgStats sdvect.X, sdvect.Y, sdvect.Z, _
            avstats.unfolded.c.dec, avstats.unfolded.c.inc, _
            avstats.unfolded.g.dec, avstats.unfolded.g.inc, _
            avstats.unfolded.S.dec, avstats.unfolded.S.inc, _
            avstats.momentvol, avstats.SigNoise, _
            avstats.SigHolder, avstats.SigInduced
        frmMeasure.ImportZijRoutine frmMeasure.lblSampName, _
            avstats.unfolded.c.dec, avstats.unfolded.c.inc, _
            avstats.momentvol, False ' (August 2007 L Carporzen) Zijderveld diagram
        frmMeasure.AveragePlotEqualArea avstats.unfolded.S.dec, avstats.unfolded.S.inc, readDats.FischerSD ' (August 2007 L Carporzen) Equal area plot
        unfolded = avstats.unfolded
        ' Save the measurement if we're not measuring the holder
        targetSample.WriteData curDemag, unfolded.g.dec, _
            unfolded.g.inc, unfolded.S.dec, unfolded.S.inc, _
            unfolded.c.dec, unfolded.c.inc, avstats.momentvol, _
            ErrorAngle, sdvect.X, sdvect.Y, sdvect.Z, readDats.UpToDown
        If RockmagMode Or RMStep.MeasureSusceptibility Then
            targetSample.WriteRockmagData RMStep, readDats.MomentVector.Z, RangeFact * sdvect.Z, readDats.MomentVector.X, RangeFact * sdvect.X, readDats.MomentVector.Y, RangeFact * sdvect.Y, unfolded.c.dec, unfolded.c.inc, avstats.momentvol, ErrorAngle, targetSample.SampleHeight
            ' multiply by rangefact to convert to emu
        End If
        If DumpRawDataStats Then
            targetSample.WriteUpMeasurements readDats, curDemag
            targetSample.WriteStatsTable readDats, curDemag
        End If
        targetSample.BackupSpecFile
    End If
    If NOCOMM_MODE Then DelayTime 5
    Set sdvect = Nothing
    Set readDats = Nothing
    SampleNameCurrent = vbNullString
    SampleStepCurrent = vbNullString
End Sub

Private Function Measure_ReadSample(specimen As Sample, _
    Optional IsHolder As Boolean = False, _
    Optional isUp As Boolean = True, Optional AllowRemeasure As Boolean = True) As MeasurementBlock
    ' This procedure goes forward and measures the sample that
    ' is currently loaded in the magnetometer.  It starts with
    ' the sample in the zero position, and ends with the sample
    ' in the zero position
    Dim msgret As VbMsgBoxResult
    Dim curMeas As Cartesian3D
    Dim X, Y, Z As Double
    Dim unfolded As Measure_Unfolded
    Dim SampleCenterPosition As Long
    Dim j As Integer
    Dim blocks As MeasurementBlocks
    Dim avstats As Measure_AvgStats
    Dim avg As Cartesian3D
    Dim MaxX, MaxY, MaxZ, MinX, MinY, MinZ As Double
    MaxX = -1000000000
    MaxY = -1000000000
    MaxZ = -1000000000
    MinX = 1000000000
    MinY = 1000000000
    MinZ = 1000000000
    Set curMeas = New Cartesian3D
    Set Measure_ReadSample = New MeasurementBlock
    frmMeasure.MomentX.Visible = False ' (October 2007 L Carporzen) Susceptibility versus demagnetization
    frmMeasure.framJumps.Top = 5040
    frmMeasure.framJumps.Left = 5400
    If frmMeasure.lblAvgCycles = 1 Then frmMeasure.InitEqualArea ' (August 2007 L Carporzen) Equal area plot
    frmMeasure.EqualArea.CurrentX = 0
    frmMeasure.EqualArea.CurrentY = 0.92
    frmMeasure.EqualArea.FontBold = True
    If IsHolder Then
    frmMeasure.EqualArea.Print "Holder" & vbCrLf & "measurement"
    frmMeasure.EqualArea.Line (0.8 - 0.01, 0.03 - 0.01)-(0.8 + 0.01, 0.03 + 0.01), 0.01, B
    frmMeasure.EqualArea.Line (0.89 - 0.01, 0.03 - 0.01)-(0.89 + 0.01, 0.03 + 0.01), 0.01, BF
    Else
    frmMeasure.EqualArea.Print "Bedding" & vbCrLf & "coordinates"
    frmMeasure.EqualArea.Circle (0.8, 0.04), 0.01, RGB(255, 0, 0)
    frmMeasure.EqualArea.Line (0.8 - 0.01, 0.015 - 0.01)-(0.8 + 0.01, 0.015 + 0.01), 0.01, B
    frmMeasure.EqualArea.Circle (0.89, 0.04), 0.01, RGB(0, 0, 255)
    frmMeasure.EqualArea.Line (0.89 - 0.01, 0.015 - 0.01)-(0.89 + 0.01, 0.015 + 0.01), 0.01, BF
    End If
    frmMeasure.EqualArea.FontBold = False
    frmMeasure.ShowPlots
    SampleNameCurrent = specimen.Samplename
    Measure_ReadSample.isUp = isUp
    If Not IsHolder Then
        For j = 1 To 4
            Measure_ReadSample.SetHolder j, Holder.Sample(j)
        Next j
    End If
    frmDCMotors.TurningMotorRotate 0 ', False
    'If frmDCMotors.UpDownHeight < SCoilPos + specimen.SampleHeight / 2 Then frmDCMotors.UpDownMove Int(SCoilPos + specimen.SampleHeight / 2), 0 ' (July 2008) Slow down after pickup a sample to don't bump on the sample changer plate
    frmDCMotors.UpDownMove Int(ZeroPos + specimen.SampleHeight / 2), 2
    ' Before the first zero, reset and zero counter
    ' then wait for numbers to settle.
    frmSQUID.CLP "A"
    frmSQUID.ResetCount "A"
    frmProgram.StatusBar "Resetting...", 3
    DelayTime (Measure_ARCDelay * 1)  ' Briefly pause
    ' First zero measurement
    ' latch data from zero position
    frmSQUID.latchVal "A", False
    Set curMeas = frmSQUID.getData(True)
    ' start move into SampleCenterPosition
    SampleCenterPosition = Int(MeasPos + specimen.SampleHeight / 2)
    frmDCMotors.UpDownMove SampleCenterPosition, 0, False
    Measure_ReadSample.SetBaseline 1, curMeas
    frmMeasure.showData curMeas.X, curMeas.Y, curMeas.Z, 0
    ' Communication problem: Rescan the first zero if a 0 appears! (August 2007 L Carporzen)
    'If Not NOCOMM_MODE And (curMeas.x = 0 Or curMeas.y = 0 Or curMeas.z = 0) Then
    '    frmDCMotors.UpDownMove Int(ZeroPos + specimen.SampleHeight / 2), 0
    '    Set Measure_ReadSample = Measure_ReadSample(specimen, isHolder, isUp, True)
    'End If
    ' Lower to sense region and take first measurement
    ' remember to center the sample in the sense region ... SampleBottom is in the INI
    ' file, and the SampleTop value was set when the system picked it up initially.
    ' Note that both positions are measured with the TestAll function homing down, so
    ' the small distance that the turning rod moves up before the limit switch clicks
    ' should not influence the pushbutton position.
    frmDCMotors.UpDownMove SampleCenterPosition, 0
    frmDCMotors.TurningMotorRotate 0 ' (November 2009 L Carporzen)
    DelayTime (Measure_ARCDelay * 1)  ' Briefly pause
    frmSQUID.latchVal "A", True
    Set curMeas = frmSQUID.getData(True)
    Measure_ReadSample.SetSample 1, curMeas
    'DelayTime (Measure_ARCDelay * 1)  ' Briefly pause
    frmDCMotors.TurningMotorRotate 90, False
    ' Adjust to baseline - just for the display
    X = curMeas.X - Measure_ReadSample.Baselines(1).X
    Y = curMeas.Y - Measure_ReadSample.Baselines(1).Y
    Z = curMeas.Z - Measure_ReadSample.Baselines(1).Z
    ' Adjust to direction
    If isUp Then
        ' +X, -Y, +Z direction
        Y = -Y
    Else
        ' +X, +Y, -Z direction
        Z = -Z
    End If
    unfolded = Measure_Unfold(specimen, X, Y, Z)
    ' (February 2010 L Carporzen) Measure the TAF and uncorrect them in sample file
    If Left$(specimen.Parent.measurementSteps.CurrentStep.DemagStepLabel, 5) = "aTAFX" Then unfolded = Measure_Unfold(specimen, -Z, X, -Y)
    If Left$(specimen.Parent.measurementSteps.CurrentStep.DemagStepLabel, 5) = "aTAFY" Then unfolded = Measure_Unfold(specimen, X, Z, -Y)
    If X > MaxX Then MaxX = X
    If Y > MaxY Then MaxY = Y
    If Z > MaxZ Then MaxZ = Z
    If X < MinX Then MinX = X
    If Y < MinY Then MinY = Y
    If Z < MinZ Then MinZ = Z
    frmMeasure.showData X, Y, Z, 1
    frmMeasure.ShowAngDat unfolded.S.dec, unfolded.S.inc, 1
    frmMeasure.PlotEqualArea unfolded.S.dec, unfolded.S.inc ' (August 2007 L Carporzen) Equal area plot
    ' Move to +Y, +X Orientation and take measurement
    MotorTurn_90
    frmSQUID.latchVal "A", True
    Set curMeas = frmSQUID.getData(True)
    frmDCMotors.TurningMotorRotate 180, False
    Measure_ReadSample.SetSample 2, curMeas
    ' Adjust to baseline - just for the display
    X = curMeas.Y - Measure_ReadSample.Baselines(1).Y
    Y = curMeas.X - Measure_ReadSample.Baselines(1).X
    Z = curMeas.Z - Measure_ReadSample.Baselines(1).Z    ' Adjust to direction
    If isUp Then
        ' +Y, +X, +Z direction
    Else
        ' -Y, +X, -Z direction
        Y = -Y
        Z = -Z
    End If
    unfolded = Measure_Unfold(specimen, X, Y, Z)
    ' (February 2010 L Carporzen) Measure the TAF and uncorrect them in sample file
    If Left$(specimen.Parent.measurementSteps.CurrentStep.DemagStepLabel, 5) = "aTAFX" Then unfolded = Measure_Unfold(specimen, -Z, X, -Y)
    If Left$(specimen.Parent.measurementSteps.CurrentStep.DemagStepLabel, 5) = "aTAFY" Then unfolded = Measure_Unfold(specimen, X, Z, -Y)
    If X > MaxX Then MaxX = X
    If Y > MaxY Then MaxY = Y
    If Z > MaxZ Then MaxZ = Z
    If X < MinX Then MinX = X
    If Y < MinY Then MinY = Y
    If Z < MinZ Then MinZ = Z
    frmMeasure.showData X, Y, Z, 2
    frmMeasure.ShowAngDat unfolded.S.dec, unfolded.S.inc, 2
    frmMeasure.PlotEqualArea unfolded.S.dec, unfolded.S.inc ' (August 2007 L Carporzen) Equal area plot
    ' Move to -X, +Y Orientation and measure
    MotorTurn_180
    frmSQUID.latchVal "A", True
    Set curMeas = frmSQUID.getData(True)
    frmDCMotors.TurningMotorRotate 270, False
    Measure_ReadSample.SetSample 3, curMeas
    ' Adjust to baseline - just for the display
    X = curMeas.X - Measure_ReadSample.Baselines(1).X
    Y = curMeas.Y - Measure_ReadSample.Baselines(1).Y
    Z = curMeas.Z - Measure_ReadSample.Baselines(1).Z
    ' Adjust to direction
    If isUp Then
        ' -X, +Y, +Z direction
        X = -X
    Else
        ' -X, -Y, -Z direction
        X = -X
        Y = -Y
        Z = -Z
    End If
    unfolded = Measure_Unfold(specimen, X, Y, Z)
    ' (February 2010 L Carporzen) Measure the TAF and uncorrect them in sample file
    If Left$(specimen.Parent.measurementSteps.CurrentStep.DemagStepLabel, 5) = "aTAFX" Then unfolded = Measure_Unfold(specimen, -Z, X, -Y)
    If Left$(specimen.Parent.measurementSteps.CurrentStep.DemagStepLabel, 5) = "aTAFY" Then unfolded = Measure_Unfold(specimen, X, Z, -Y)
    If X > MaxX Then MaxX = X
    If Y > MaxY Then MaxY = Y
    If Z > MaxZ Then MaxZ = Z
    If X < MinX Then MinX = X
    If Y < MinY Then MinY = Y
    If Z < MinZ Then MinZ = Z
    frmMeasure.showData X, Y, Z, 3
    frmMeasure.ShowAngDat unfolded.S.dec, unfolded.S.inc, 3
    frmMeasure.PlotEqualArea unfolded.S.dec, unfolded.S.inc ' (August 2007 L Carporzen) Equal area plot
    ' Move to -Y, -X Orientation and measure
    MotorTurn_270
    frmSQUID.latchVal "A", True
    Set curMeas = frmSQUID.getData(True)
    frmDCMotors.UpDownMove Int(ZeroPos + specimen.SampleHeight / 2), 0, False
    frmDCMotors.TurningMotorRotate 360, False
    Measure_ReadSample.SetSample 4, curMeas
    ' Adjust to baseline - just for the display
    X = curMeas.Y - Measure_ReadSample.Baselines(1).Y
    Y = curMeas.X - Measure_ReadSample.Baselines(1).X
    Z = curMeas.Z - Measure_ReadSample.Baselines(1).Z
    ' Adjust to direction
    If isUp Then
        ' -Y, -X, +Z direction
        X = -X
        Y = -Y
    Else
        ' +Y, -X, -Z direction
        X = -X
        Z = -Z
    End If
    unfolded = Measure_Unfold(specimen, X, Y, Z)
    ' (February 2010 L Carporzen) Measure the TAF and uncorrect them in sample file
    If Left$(specimen.Parent.measurementSteps.CurrentStep.DemagStepLabel, 5) = "aTAFX" Then unfolded = Measure_Unfold(specimen, -Z, X, -Y)
    If Left$(specimen.Parent.measurementSteps.CurrentStep.DemagStepLabel, 5) = "aTAFY" Then unfolded = Measure_Unfold(specimen, X, Z, -Y)
    If X > MaxX Then MaxX = X
    If Y > MaxY Then MaxY = Y
    If Z > MaxZ Then MaxZ = Z
    If X < MinX Then MinX = X
    If Y < MinY Then MinY = Y
    If Z < MinZ Then MinZ = Z
    frmMeasure.showData X, Y, Z, 4
    frmMeasure.ShowAngDat unfolded.S.dec, unfolded.S.inc, 4
    frmMeasure.PlotEqualArea unfolded.S.dec, unfolded.S.inc ' (August 2007 L Carporzen) Equal area plot
    ' Lift to zero and measure
    ' Rotate the sample back to start direction
    'Motor_MoveMeasdownToZero
    frmDCMotors.UpDownMove Int(ZeroPos + specimen.SampleHeight / 2), 0
    MotorTurn_360   ' remember, on the new systems we keep the tube spinning in the same direction
    frmSQUID.latchVal "A", True
    Set curMeas = frmSQUID.getData(True)
    Measure_ReadSample.SetBaseline 2, curMeas
    frmMeasure.showData curMeas.X, curMeas.Y, curMeas.Z, 5
    ' Communication problem: Rescan automatically the set of measurement if a 0 appears! (August 2007 L Carporzen)
    'If Not NOCOMM_MODE And (curMeas.x = 0 Or curMeas.y = 0 Or curMeas.z = 0) Then Set Measure_ReadSample = Measure_ReadSample(specimen, isHolder, isUp, True)
    Set blocks = New MeasurementBlocks
    blocks.Add Measure_ReadSample
    avstats = Measure_CalcStats(specimen, blocks)
    Set blocks = Nothing
    If Prog_halted Then Exit Function ' (September 2007 L Carporzen) New version of the Halt button
'                               NEW PARAMETERS TO MONITOR THE SQUID NOISEs (August 2007 L Carporzen)
' Look at the homogeneity of each axis: it is only informative for the user, their is no automatic rescanof a bad value.
' The user sees the delta on each axis (in emu) as well as the ratio of this delta by the measured moment
' If the ratio is greater than 1, the boxes corresponding to the "bad" axis light in Orange
' If the ratio is greater than 5, the boxes corresponding to the "bad" axis light in Red
' Notice that X and Y axis share the SQUIDs X and Y.
    If NOCOMM_MODE Then avstats.momentvol = 0.000000001
    If avstats.momentvol = 0 Then avstats.momentvol = 0.000000001
    frmMeasure.lblDeltaX.Caption = Format$(RangeFact * (MaxX - MinX), "0.0000E+")
    frmMeasure.lblRatioX.Caption = Format$(RangeFact * (MaxX - MinX) / (specimen.Vol * avstats.momentvol), "0.00")
    frmMeasure.lblDeltaY.Caption = Format$(RangeFact * (MaxY - MinY), "0.0000E+")
    frmMeasure.lblRatioY.Caption = Format$(RangeFact * (MaxY - MinY) / (specimen.Vol * avstats.momentvol), "0.00")
    frmMeasure.lblDeltaZ.Caption = Format$(RangeFact * (MaxZ - MinZ), "0.0000E+")
    frmMeasure.lblRatioZ.Caption = Format$(RangeFact * (MaxZ - MinZ) / (specimen.Vol * avstats.momentvol), "0.00")
    frmMeasure.lblOrange.Visible = False
    frmMeasure.lblRed.Visible = False
    frmMeasure.lblWarning.Visible = False
    If RangeFact * (MaxX - MinX) / (specimen.Vol * avstats.momentvol) > 0.1 / JumpThreshold Then
        frmMeasure.lblDeltaX.BackColor = ColorOrange
        frmMeasure.lblRatioX.BackColor = ColorOrange
        frmMeasure.lblOrange.Visible = True
        If RangeFact * (MaxX - MinX) / (specimen.Vol * avstats.momentvol) > 0.5 / JumpThreshold Then
            frmMeasure.lblDeltaX.BackColor = ColorRed
            frmMeasure.lblRatioX.BackColor = ColorRed
            frmMeasure.lblRed.Visible = True
            frmMeasure.lblWarning.Visible = True
        End If
    Else
        frmMeasure.lblDeltaX.BackColor = RGB(255, 255, 255)
        frmMeasure.lblRatioX.BackColor = RGB(255, 255, 255)
    End If
    If RangeFact * (MaxY - MinY) / (specimen.Vol * avstats.momentvol) > 0.1 / JumpThreshold Then
        frmMeasure.lblDeltaY.BackColor = ColorOrange
        frmMeasure.lblRatioY.BackColor = ColorOrange
        frmMeasure.lblOrange.Visible = True
        If RangeFact * (MaxY - MinY) / (specimen.Vol * avstats.momentvol) > 0.5 / JumpThreshold Then
            frmMeasure.lblDeltaY.BackColor = ColorRed
            frmMeasure.lblRatioY.BackColor = ColorRed
            frmMeasure.lblRed.Visible = True
            frmMeasure.lblWarning.Visible = True
        End If
    Else
        frmMeasure.lblDeltaY.BackColor = RGB(255, 255, 255)
        frmMeasure.lblRatioY.BackColor = RGB(255, 255, 255)
    End If
    If RangeFact * (MaxZ - MinZ) / (specimen.Vol * avstats.momentvol) > 0.1 / JumpThreshold Then
        frmMeasure.lblDeltaZ.BackColor = ColorOrange
        frmMeasure.lblRatioZ.BackColor = ColorOrange
        frmMeasure.lblOrange.Visible = True
        If RangeFact * (MaxZ - MinZ) / (specimen.Vol * avstats.momentvol) > 0.5 / JumpThreshold Then
            frmMeasure.lblDeltaZ.BackColor = ColorRed
            frmMeasure.lblRatioZ.BackColor = ColorRed
            frmMeasure.lblRed.Visible = True
            frmMeasure.lblWarning.Visible = True
        End If
    Else
        frmMeasure.lblDeltaZ.BackColor = RGB(255, 255, 255)
        frmMeasure.lblRatioZ.BackColor = RGB(255, 255, 255)
    End If
'                               NEW PARAMETERS TO AVOID RECORD OF SQUID JUMPS (August 2007 L Carporzen)
' In the program, three new options are available in the Options menu:
' Above a certain moment, the initial criteria looking at the CSD and the Signal/Noise is apply
' Ian changed that value from 10-7 to 8.10-9 emu because of some unacceptable large CSD on weak samples
' "Critical moment (emu):" (default = 8.10-9 emu)
' If the measured moment is lower than 10-6 emu, the accepted differences between the zero measurements of each of the three SQUID is proportional to the measured moment.
' In case of very low moment, the possible drift of the SQUID may block this criteria. For that reason, you can change the "Jump sensitivity" (default = 1)
' In order to avoid infinite measurement, we put a limit above which the program will accept more easily a measurement.
' However, you can decide to accept only the good measurement which fit all the criteria by increasing the number of try to a much greater value.
' "Number of try:" (default = 5)
' If the zero measurements are too different we need to remeasure
    X = Abs(curMeas.X - Measure_ReadSample.Baselines(1).X)
    Y = Abs(curMeas.Y - Measure_ReadSample.Baselines(1).Y)
    Z = Abs(curMeas.Z - Measure_ReadSample.Baselines(1).Z)
' To avoid repetitive measurements, "Number of try:" (Meascount) is the maximum try per measurement. You can change in the Options menu the "Number of try:" (default = 5)
' You can change in the Options menu the minimum moment ("Critical moment (emu):", default = 8.10-9 emu) where the CSD criteria is apply
    If Meascount >= NbTry Then
' It sends an email when the measurement is accepted just to inform what were the zero differences
        If NbTry > 0 And _
           ((X > JumpThreshold) Or _
           (Y > JumpThreshold) Or _
           (Z > JumpThreshold)) _
        Then
        
            frmSendMail.MailNotification _
                            "Measurement recorded even though... the difference between " & _
                            "the two zero measurements is " & Format$(X, "0.00000") & " on X, " & _
                            Format$(Y, "0.00000") & " on Y and " & Format$(Z, "0.00000") & _
                            " on Z; the moment is " & Format$((specimen.Vol * avstats.momentvol), "0.0000E+") & _
                            " emu and CSD=" & Format$(Measure_ReadSample.FischerSD, "0.0"), _
                            CodeRed
                            
        End If
                            
' First, the CSD original criteria could be enough to rescan
    ElseIf AllowRemeasure And _
           (specimen.Vol * avstats.momentvol) > MomMinForRedo And _
           ((avstats.SigNoise < 1) Or _
            (avstats.SigInduced < 1) Or _
            (Measure_ReadSample.FischerSD > RemeasureCSDThreshold)) _
    Then
        
        frmMeasure.lblRescan.Caption = "CSD = " & Format$(Measure_ReadSample.FischerSD, "0")
        
        Meascount = Meascount + 1
        
        '(March 9, 2011 - I Hilburn)
        'Adding in modification to enable the user to separately set the remeasure Holder
        'setting if the CSD is too high on a high-moment holder measurement
        If IsHolder = True And _
           Meascount > NbHolderTry _
        Then
        
            'Need to skip the code below, this holder should not be remeasured another time
        
        Else
        
            'User setting allows the code to make another try at remeasuring the holder
        
            ' For very strong moment, > StrongMom (default 2.10-2 emu), the SQUID response
            'will not enable this criteria, we look on the CSD original criteria.
        
            If (specimen.Vol * avstats.momentvol) > StrongMom And _
               Meascount > 0 Then _
               
               frmSendMail.MailNotification _
                            "Redoing the measurement because the CSD=" & _
                            Format$(Measure_ReadSample.FischerSD, "0") & _
                            " and the moment is " & Format$((specimen.Vol * avstats.momentvol), "0.0000E+") & _
                            " emu.", _
                            CodeYellow
                            
            End If
                            
            frmDCMotors.UpDownMove Int(ZeroPos + specimen.SampleHeight / 2), 0
            
            DelayTime (Measure_ARCDelay * 1)  ' Briefly pause
            
            Set Measure_ReadSample = Measure_ReadSample(specimen, IsHolder, isUp, True)
            
        End If
        
    ElseIf AllowRemeasure And _
           (specimen.Vol * avstats.momentvol) < StrongMom And _
           (X > JumpThreshold Or Y > JumpThreshold Or Z > JumpThreshold) _
    Then

' The CSD criteria has accepted the measurement
' For moment below StrongMom (default 2.10-2 emu), we rescan if their is a jump > JumpThreshold (default 0.1x10-5 emu)
' For moment > InterMom (default 10-6 emu), the difference between each of the three zero measurements
' needs to be < JumpThreshold (default 0.1x10-5 emu)
        
        If X > JumpThreshold And _
           Meascount > 0 _
        Then
        
            frmSendMail.MailNotification "X=" & Format$(X, "0.00000") & " SQUID jump, " & _
                                         Format$(Meascount, "0") & " redoing the measurement for a moment of " & _
                                         Format$((specimen.Vol * avstats.momentvol), "0.0000E+") & _
                                         " emu, CSD=" & Format$(Measure_ReadSample.FischerSD, "0.0"), _
                                         CodeRed
            
        End If
        
        If Y > JumpThreshold And _
           Meascount > 0 _
        Then
        
            frmSendMail.MailNotification "Y=" & Format$(Y, "0.00000") & " SQUID jump, " & _
                                         Format$(Meascount, "0") & " redoing the measurement for a moment of " & _
                                         Format$((specimen.Vol * avstats.momentvol), "0.0000E+") & _
                                         " emu, CSD=" & Format$(Measure_ReadSample.FischerSD, "0.0"), _
                                         CodeRed
            
        End If
        
        If Z > JumpThreshold And _
           Meascount > 0 _
        Then
        
            frmSendMail.MailNotification "Z=" & Format$(Z, "0.00000") & " SQUID jump, " & _
                                         Format$(Meascount, "0") & " redoing the measurement for a moment of " & _
                                         Format$((specimen.Vol * avstats.momentvol), "0.0000E+") & _
                                         " emu, CSD=" & Format$(Measure_ReadSample.FischerSD, "0.0"), _
                                         CodeRed
            
        End If
        
' Information mails when the measurement will be repeated because of a difference
' between the zero > JumpThreshold (default 0.1) (x10-5 emu)
        
        frmMeasure.lblRescan.Caption = "SQUID jumps"
        
        Meascount = Meascount + 1
        
        frmDCMotors.UpDownMove Int(ZeroPos + specimen.SampleHeight / 2), 0
        
        DelayTime (Measure_ARCDelay * 1)  ' Briefly pause
        
        Set Measure_ReadSample = Measure_ReadSample(specimen, IsHolder, isUp, True)
    
    ElseIf AllowRemeasure And _
           specimen.Vol * avstats.momentvol < IntermMom And _
           specimen.Vol * avstats.momentvol > MomMinForRedo And _
           (X / (specimen.Vol * avstats.momentvol) > JumpSensitivity / RangeFact Or _
            Y / (specimen.Vol * avstats.momentvol) > JumpSensitivity / RangeFact Or _
            Z / (specimen.Vol * avstats.momentvol) > JumpSensitivity / RangeFact) _
    Then
    
' The large jump criteria has accepted the measurement
' For moment < InterMom (default 10-6 emu) and > MomMinForRedo (default 8.10-9 emu),
' the difference between each of the three zero measurements is controled by the measured moment:
' You can change in the Options menu the proportion of the moment ("Jump sensitivity", default = 1)
' which will be use to compare the zero measurements

        frmMeasure.lblRescan.Caption = "Small jumps"
        
        Meascount = Meascount + 1
        
        frmDCMotors.UpDownMove Int(ZeroPos + specimen.SampleHeight / 2), 0
        
        DelayTime (Measure_ARCDelay * 1)  ' Briefly pause
        
        Set Measure_ReadSample = Measure_ReadSample(specimen, IsHolder, isUp, True)
    
    End If ' No jump at all
    
    frmMeasure.lblRescan.Caption = " " ' Reset the rescan label
    
    ' Label the small SQUID jumps in the Measure window:
    
    If specimen.Vol * avstats.momentvol < MomMinForRedo And _
       (X / (specimen.Vol * avstats.momentvol) > JumpSensitivity / RangeFact Or _
        Y / (specimen.Vol * avstats.momentvol) > JumpSensitivity / RangeFact Or _
        Z / (specimen.Vol * avstats.momentvol) > JumpSensitivity / RangeFact) _
    Then
    
        If X / (specimen.Vol * avstats.momentvol) > JumpSensitivity / RangeFact Then
        
            frmMeasure.lblXSQUID.Caption = Format$(X, "0.000000")
            
        Else
        
            frmMeasure.lblXSQUID.Caption = " "
            
        End If
        
        If Y / (specimen.Vol * avstats.momentvol) > JumpSensitivity / RangeFact Then
        
            frmMeasure.lblYSQUID.Caption = Format$(Y, "0.000000")
            
        Else
        
            frmMeasure.lblYSQUID.Caption = " "
            
        End If
        
        If Z / (specimen.Vol * avstats.momentvol) > JumpSensitivity / RangeFact Then
        
            frmMeasure.lblZSQUID.Caption = Format$(Z, "0.000000")
            
        Else
        
            frmMeasure.lblZSQUID.Caption = " "
            
        End If
        
        frmMeasure.lblRescan.Caption = "Small jumps"
        
    Else
    
        frmMeasure.lblXSQUID.Caption = " "
        frmMeasure.lblYSQUID.Caption = " "
        frmMeasure.lblZSQUID.Caption = " "
        
    End If
    
    ' We've finished the measuring cycle
    ' So, now calculate the components, etc.
    
    Meascount = 1
    
    If DEBUG_MODE Then
    
        frmDebug.Msg specimen.Samplename & ": " & _
                     Measure_ReadSample.Average.X & ", " & _
                     Measure_ReadSample.Average.Y & "," & _
                     Measure_ReadSample.Average.Z
        
    End If
    
    ' ADD Range switch code here if necessary. !!
    Set curMeas = Nothing
    
    SampleNameCurrent = vbNullString
    
End Function

Public Function Measure_Rotate(ByVal inc As Double, ByVal dec As Double, _
                       ByVal bA As Double, ByVal bD As Double) _
                       As Angular3D
    ' Subroutine to perform the bedding-style rotations.  The direction
    ' to be rotated should be given in polar coordinates of DEC and INC
    ' (in degrees), while the direction of bedding dip (BA, not the
    ' strike!) and bedding dip (BD) are in radians.  The routine returns
    ' a new DEC and INC corresponding to the tilt-corrected directions.
    Dim X, Y, Z As Double
    Dim SA, CA, CDP, SDP As Double
    Dim xC, yC, zC As Double
    Set Measure_Rotate = New Angular3D
    Z = -Sin(DegToRad(inc))
    X = Cos(DegToRad(inc)) * Cos(DegToRad(dec))
    Y = Cos(DegToRad(inc)) * Sin(DegToRad(dec))
    SA = -Sin(bA)
    CA = Cos(bA)
    CDP = Cos(bD)
    SDP = Sin(bD)
    xC = X * (SA * SA + CA * CA * CDP) + Y * (CA * SA * (1 - CDP)) - Z * SDP * CA
    yC = X * CA * SA * (1 - CDP) + Y * (CA * CA + SA * SA * CDP) + Z * SA * SDP
    zC = X * CA * SDP - Y * SDP * SA + Z * CDP
    ' Corrected incl and decl
    Measure_Rotate.inc = -RadToDeg(Atn(zC / Sqr(xC ^ 2 + yC ^ 2)))
    Measure_Rotate.dec = RadToDeg(Atan2(xC, yC), True)
End Function

' Measure_TreatAndRead
'
' This is the routine for taking care of AF demagnetization,
' susceptibility measurements, etc.
Public Sub Measure_TreatAndRead(targetSample As Sample, Optional ByVal useChanger = False)
    Dim RockmagMode As Boolean
    Dim doMeasure As Boolean
    Dim LabelString As String
    If Prog_halted Then Exit Sub ' (September 2007 L Carporzen) New version of the Halt button
    frmDCMotors.TurningMotorAngleOffset -TrayOffsetAngle  '+ 360 (November 2009 L Carporzen) change to 360 - instead of + because we changed the Sub TrayOffsetAngle
    With targetSample.Parent
    
        RockmagMode = .RockmagMode
        .measurementSteps.CurrentStepIndex = 1
        doMeasure = .measurementSteps.CurrentStep.Measure
        
        If .measurementSteps.Count = 1 And Not .RockmagMode Then doMeasure = True
        
        If RockmagMode Then
                
                targetSample.WriteRockmagInfoLine "Instrument: " & MailFromName
                targetSample.WriteRockmagInfoLine "Time: " & Format(Now, "yyyy-mm-dd hh:mm")
        
        End If
        
        Do While .measurementSteps.CurrentStepIndex > 0
            
            If Prog_halted Then Exit Sub ' (September 2007 L Carporzen) New version of the Halt button
            
            LabelString = targetSample.Samplename & " @ " & .curDemagLong
            
            If .measurementSteps.Count > 1 Then _
                LabelString = LabelString & " [" & Format$(.measurementSteps.CurrentStepIndex, "0") & _
                "/" & Format$(.measurementSteps.Count, "0") & "]"
                
            frmProgram.StatusBar "Measuring samples... (" & LabelString & ")", 1
            
            SampleNameCurrent = targetSample.Samplename
            SampleStepCurrent = .measurementSteps.CurrentStep.DemagStepLabelLong
            
            If .doUp Then
            
                SampleOrientationCurrent = Magnet_SampleOrientationUp
                
            Else
            
                SampleOrientationCurrent = Magnet_SampleOrientationDown
                
            End If
            
            If (.measurementSteps.CurrentStep.MeasureSusceptibility) And (.doUp Or (Not .doBoth)) Then
                
                '(March 10, 2011 - I Hilburn)
                'This If ... then statement replaces the currently commented out code in SampleCommand.Execute
                'and in frmMagnetometerControl.cmdManHolder_Click.
                'This ensures that the susceptibility lag time is only calculated during holder measurements
                'if the user has selected for the susceptibility to be measured and should
                'eliminate the pause associated with all Holder measurements even when the user did not
                'want the susceptibility to be measured.
                If targetSample.Samplename = "Holder" And _
                   COMPortSusceptibility > 0 And _
                   EnableSusceptibility = True _
                Then
                
                    'This is a holder measurement
                    'Need to get the LagTime for the susceptibility measurement
                    frmSusceptibilityMeter.LagTime
                    
                End If
                
                Susceptibility_Measure targetSample, (targetSample.Samplename = "Holder")
            
            End If
            
            .measurementSteps.CurrentStep.PerformStep targetSample
            
            frmMeasure.SetFields .avgSteps, .curDemagLong, .doUp, .doBoth, .filename
            frmMeasure.clearData
            frmMeasure.HideStats
            frmMeasure.clearStats
            
            'If doMeasure Then Measure_Read targetSample, .measurementSteps.CurrentStep, RockmagMode
            '.measurementSteps.AdvanceStep
            If doMeasure Then
            
                Measure_Read targetSample, .measurementSteps.CurrentStep, RockmagMode
                .measurementSteps.AdvanceStep
                
            ElseIf RockmagMode Then
                
                targetSample.WriteRockmagData .measurementSteps.CurrentStep, "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", targetSample.SampleHeight
                .measurementSteps.AdvanceStep
            
            End If
        
        Loop
    
    End With
    
    If Prog_halted Then Exit Sub ' (September 2007 L Carporzen) New version of the Halt button
    frmProgram.StatusBar "Measuring samples...", 1
    frmDCMotors.TurningMotorAngleOffset TrayOffsetAngle '+ 360 (November 2009 L Carporzen) change to 360 + instead of - because we changed the Sub TrayOffsetAngle
    SampleNameCurrent = vbNullString
    SampleStepCurrent = vbNullString
    SampleOrientationCurrent = 0
End Sub

Public Function Measure_Unfold(specimen As Sample, _
    ByVal X As Double, ByVal Y As Double, ByVal Z As Double, _
    Optional HO As Double = 0) As Measure_Unfolded
    '  A subroutine to feed in a direction in sample coordinates, and
    '  to unfold w.r.t. fold axes, bedding orientation, and sample
    '  orientation, to spit out a declination and inclination.  UNFOLD
    '  wants to be fed directions in sample coordinates, in variables
    '  XTEMP, YTEMP, and ZTEMP.
    '  COMPUTE DECL. AND INCL. IN SPECIMEN COORDINATES AS POSITIVE FROM +X
    Dim ret As Measure_Unfolded
    Dim ax As Double, ay As Double, aZ As Double
    Dim DD As Double, DP As Double, SD As Double, CD As Double
    Dim BB As Double, CC As Double, XP As Double
    Dim MT As Double
    Dim magDec As Double
    Set ret.c = New Angular3D
    Set ret.g = New Angular3D
    Set ret.S = New Angular3D
    MT = Sqr(Abs(X ^ 2 + Y ^ 2 + Z ^ 2))
    ret.c.dec = RadToDeg(Atan2(X, Y), True)
    ret.c.inc = RadToDeg(Atn(Z / (HO + 0.00001)))
    '   COORDINATE TRANSFORM FOR SAMPLE ORIENTATION IN FIELD
    '   CORRECT DIP DIRECTION FOR MAGNETIC DECLINATION,
    '   AND USE CALTECH ORIENTATION SYSTEM
    With specimen
        magDec = .Parent.magDec
        DD = .CorePlateStrike + magDec - 90#
        DP = DegToRad((90 - .CorePlateDip))
        SD = Sin(DP)
        CD = Cos(DP)
    End With
    If MT = 0 Then MT = 0.00001
    ax = X / MT
    ay = Y / MT
    aZ = Z / MT
    XP = ax * SD + aZ * CD
    BB = Sqr(XP ^ 2 + ay ^ 2)
    CC = aZ * SD - ax * CD
    If BB = 0 Then BB = 0.00001
    ret.g.inc = RadToDeg(Atn(CC / BB))
    ' COMPUTE DEC = ARCTAN(Y/X)
    ret.g.dec = RadToDeg(Atan2(XP, ay) + DegToRad(DD), True)
    Set ret.S = Measure_Bedding(specimen, ret.g.inc, ret.g.dec)
    Measure_Unfold = ret
End Function

