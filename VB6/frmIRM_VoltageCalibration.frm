VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmIRM_VoltageCalibration 
   Caption         =   "8"
   ClientHeight    =   6930
   ClientLeft      =   8055
   ClientTop       =   5445
   ClientWidth     =   8145
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6930
   ScaleWidth      =   8145
   Begin VB.CommandButton cmdHelp 
      BackColor       =   &H0000FFFF&
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      MaskColor       =   &H0080FFFF&
      Style           =   1  'Graphical
      TabIndex        =   47
      Top             =   2640
      UseMaskColor    =   -1  'True
      Width           =   495
   End
   Begin VB.PictureBox picGetCapacitorVoltage 
      BackColor       =   &H80000013&
      Height          =   2175
      Left            =   2280
      ScaleHeight     =   2115
      ScaleWidth      =   3675
      TabIndex        =   29
      Top             =   1800
      Width           =   3735
      Begin VB.TextBox txtCapacitorVoltage 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         TabIndex        =   34
         Top             =   1080
         Width           =   2175
      End
      Begin VB.CommandButton cmdRedo 
         Caption         =   "ReDo"
         Height          =   375
         Left            =   1200
         TabIndex        =   32
         Top             =   1680
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   2520
         TabIndex        =   31
         Top             =   1680
         Width           =   975
      End
      Begin VB.CommandButton cmdAccept 
         Caption         =   "Accept"
         Height          =   375
         Left            =   120
         TabIndex        =   30
         Top             =   1680
         Width           =   975
      End
      Begin VB.PictureBox picHighlight 
         BackColor       =   &H000080FF&
         BorderStyle     =   0  'None
         Height          =   490
         Left            =   660
         ScaleHeight     =   495
         ScaleWidth      =   2295
         TabIndex        =   35
         Top             =   1040
         Width           =   2300
      End
      Begin VB.Label lblDirections 
         BackColor       =   &H80000003&
         BackStyle       =   0  'Transparent
         Caption         =   "Input the Voltage displayed on the IRM capacitor box:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         TabIndex        =   33
         Top             =   120
         Width           =   3495
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Max Capacitor Volt."
      Height          =   2295
      Left            =   6360
      TabIndex        =   37
      Top             =   120
      Width           =   1695
      Begin VB.TextBox txtIRMTransMaxCapVolts 
         Height          =   285
         Left            =   240
         TabIndex        =   40
         Top             =   1320
         Width           =   1095
      End
      Begin VB.CheckBox chkSameAsAxial 
         Caption         =   "Same as Axial"
         Height          =   495
         Left            =   240
         TabIndex        =   39
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox txtIRMAxialMaxCapVolts 
         Height          =   285
         Left            =   240
         TabIndex        =   38
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label9 
         Caption         =   "Transverse"
         Height          =   255
         Left            =   240
         TabIndex        =   42
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label8 
         Caption         =   "Axial:"
         Height          =   255
         Left            =   240
         TabIndex        =   41
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   375
      Left            =   1800
      TabIndex        =   20
      Top             =   6480
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   3000
      TabIndex        =   19
      Top             =   6480
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   18
      Top             =   6480
      Width           =   975
   End
   Begin VB.CommandButton cmdPause 
      Caption         =   "Pause Cal."
      Height          =   375
      Left            =   6960
      TabIndex        =   17
      Top             =   6480
      Width           =   1095
   End
   Begin VB.CommandButton cmdStartStop 
      Caption         =   "Start Calibration"
      Height          =   375
      Left            =   5280
      TabIndex        =   16
      Top             =   6480
      Width           =   1455
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   375
      Left            =   5160
      TabIndex        =   15
      ToolTipText     =   "Click to add steps starting at From voltage in steps specified by step size, up to or less than the To voltage."
      Top             =   2280
      Width           =   855
   End
   Begin VB.TextBox txtTo 
      Height          =   285
      Left            =   4080
      TabIndex        =   14
      ToolTipText     =   "End calibration voltage (suggested < 400 V)"
      Top             =   2280
      Width           =   735
   End
   Begin VB.TextBox txtFrom 
      Height          =   285
      Left            =   2760
      TabIndex        =   12
      ToolTipText     =   "Starting Voltage (suggested > 5 V)"
      Top             =   2280
      Width           =   735
   End
   Begin VB.TextBox txtStepSize 
      Height          =   285
      Left            =   1080
      TabIndex        =   10
      ToolTipText     =   "Step Size in Capacitor Volts (1 - 450 V) to increment the experiment voltage"
      Top             =   2280
      Width           =   855
   End
   Begin VB.Frame Frame2 
      Caption         =   "DAQ Voltage VS IRM Capacitor Voltage"
      Height          =   1575
      Left            =   1920
      TabIndex        =   4
      Top             =   120
      Width           =   4335
      Begin VB.TextBox txtPulseReturnConversion 
         Height          =   285
         Left            =   1800
         TabIndex        =   26
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox txtReturnError 
         Height          =   285
         Left            =   3120
         TabIndex        =   25
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox txtOutputError 
         Height          =   285
         Left            =   3120
         TabIndex        =   8
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox txtDAQOutputConversion 
         Height          =   285
         Left            =   1800
         TabIndex        =   6
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Pulse Return Volts Conversion factor:"
         Height          =   495
         Left            =   240
         TabIndex        =   28
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label6 
         Caption         =   "Error:"
         Height          =   255
         Left            =   3120
         TabIndex        =   27
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Error:"
         Height          =   255
         Left            =   3120
         TabIndex        =   7
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "DAQ Output Volts Conversion factor:"
         Height          =   495
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "IRM Coil"
      Height          =   1575
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1695
      Begin VB.CheckBox chkLockCoils 
         Caption         =   "Lock coil selection"
         Height          =   375
         Left            =   240
         TabIndex        =   36
         Top             =   1080
         Width           =   1335
      End
      Begin VB.OptionButton optCoil 
         Caption         =   "Transverse"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   3
         Top             =   720
         Width           =   1095
      End
      Begin VB.OptionButton optCoil 
         Caption         =   "Axial"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   1095
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridVoltageCal 
      Height          =   2295
      Left            =   240
      TabIndex        =   0
      Top             =   3480
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   4048
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000010&
      Height          =   3045
      Left            =   120
      ScaleHeight     =   2985
      ScaleWidth      =   7920
      TabIndex        =   21
      Top             =   3360
      Width           =   7980
      Begin VB.CommandButton cmdCalcConvFactor 
         Caption         =   "Calculate Conversion Factors"
         Height          =   375
         Left            =   5400
         TabIndex        =   24
         Top             =   2520
         Width           =   2295
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear"
         Height          =   375
         Left            =   1800
         TabIndex        =   23
         Top             =   2520
         Width           =   1215
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete"
         Height          =   375
         Left            =   240
         TabIndex        =   22
         Top             =   2520
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdAddSingle 
      Caption         =   "Add Single Step"
      Height          =   375
      Left            =   4320
      TabIndex        =   43
      ToolTipText     =   "Add a single voltage step."
      Top             =   2880
      Width           =   1695
   End
   Begin VB.TextBox txtSingleStep 
      Height          =   285
      Left            =   3120
      TabIndex        =   44
      ToolTipText     =   "Voltage (0 - 450 V) to add as a calibration step"
      Top             =   2880
      Width           =   855
   End
   Begin VB.Label Label10 
      Caption         =   "Capacitor Output Voltage Steps to test calibration:"
      Height          =   255
      Left            =   240
      TabIndex        =   46
      Top             =   1920
      Width           =   5895
   End
   Begin VB.Label Label12 
      Caption         =   "Step:"
      Height          =   255
      Left            =   2640
      TabIndex        =   45
      Top             =   2880
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "To:"
      Height          =   255
      Left            =   3720
      TabIndex        =   13
      Top             =   2280
      Width           =   375
   End
   Begin VB.Label Label4 
      Caption         =   "From:"
      Height          =   255
      Left            =   2280
      TabIndex        =   11
      Top             =   2280
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Step Size:"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   2280
      Width           =   855
   End
   Begin VB.Menu mnuIRM 
      Caption         =   "IRM Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuIRMCopy 
         Caption         =   "Copy"
      End
   End
End
Attribute VB_Name = "frmIRM_VoltageCalibration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim CoilString As String
Public CurrentRow As Long

Dim VoltageAccepted As Boolean
Dim VoltageCancelled As Boolean
Dim UnsavedChanges As Boolean

Dim CalStatus As String

'Resize Control
Private Type ControlPositionType
    Left As Single
    Top As Single
    Width As Single
    Height As Single
    FontSize As Single
End Type
Private m_ControlPositions() As ControlPositionType
Private m_FormWid As Single
Private m_FormHgt As Single

Private Sub chkLockCoils_Click()

    If Me.chkLockCoils.Value = Checked Then
    
        CoilsLocked = True
        optCoil(0).Enabled = False
        optCoil(1).Enabled = False
        
    ElseIf Me.chkLockCoils.Value = Unchecked Then
    
        CoilsLocked = False
        optCoil(0).Enabled = True
        optCoil(1).Enabled = True
        
    End If

End Sub

Private Sub chkSameAsAxial_Click()
    
    UnsavedChanges = True
    
    If Me.chkSameAsAxial.Value = Unchecked Then
    
        Me.txtIRMTransMaxCapVolts = modConfig.IRMTransVoltMax
        Me.txtIRMTransMaxCapVolts.Enabled = True
        
    Else
    
        Me.txtIRMTransMaxCapVolts = Me.txtIRMAxialMaxCapVolts
        Me.txtIRMTransMaxCapVolts.Enabled = False
        
    End If

End Sub

Private Sub cmdAccept_Click()

    'Set the voltage accepted flag
    VoltageAccepted = True
    VoltageCancelled = False
    
    'Hide the picture box
    Me.picGetCapacitorVoltage.Visible = False

End Sub

Private Sub cmdAdd_Click()

    Dim i As Long
    Dim Volts As Double
    Dim StepSize As Double
    Dim isLog As Boolean
    Dim FromVolts As Double
    Dim ToVolts As Double
    Dim MaxCoilVolts As Double
    Dim MinCoilVolts As Double
    Dim NumReplicates As Long
    Dim StartRow As Long
    
    'Read out values into local variables for from and to volts
    FromVolts = val(Me.txtFrom)
    ToVolts = val(Me.txtTo)
    
    'Depending on which coil is selected
    'by the user, load that coil's max and min voltages into
    'the two local variables
    If ActiveCoilSystem = AxialCoilSystem Then
    
        MaxCoilVolts = modConfig.PulseVoltMax
        MinCoilVolts = 0
        CoilString = "IRM Axial"
        
    ElseIf ActiveCoilSystem = TransverseCoilSystem Then
    
        MaxCoilVolts = modConfig.PulseVoltMax
        MinCoilVolts = 0
        CoilString = "IRM Transverse"
        
    End If

    'Now validate Max and Min coil voltages
    If MaxCoilVolts <= MinCoilVolts Then
    
        'Quick Message Box to user
        MsgBox "Max " & CoilString & " voltage must be larger than the Min voltage." & _
                vbNewLine & vbNewLine & "Max Voltage = " & Trim(Str(MaxCoilVolts)) & _
                " Volts" & vbNewLine & "Min Voltage = " & Trim(Str(MinCoilVolts)) & _
                " Volts", , _
                "Warning!"
                
        Exit Sub
        
    End If
    
    'Make sure both max and min coil voltages are greater than zero
    'Note: We also never want the Max coil voltage to equal zero.
    If MaxCoilVolts <= 0 Or MinCoilVolts < 0 Then
    
        MsgBox "Max and/or Min " & CoilString & " voltages are less than zero." & _
                vbNewLine & vbNewLine & "Max Voltage = " & Trim(Str(MaxCoilVolts)) & _
                " Volts" & vbNewLine & "Min Voltage = " & Trim(Str(MinCoilVolts)) & _
                " Volts", , _
                "Warning!"
                
        Exit Sub
        
    End If
        
    'Set the StepSize local variable
    StepSize = val(Me.txtStepSize)
   
    If FromVolts > ToVolts And StepSize > 0 Then
    
        MsgBox "From Voltage must be smaller than To voltage with a positive volt step size." & _
                vbNewLine & "From = " & Trim(Str(FromVolts)) & " Volts" & vbNewLine & _
                "To = " & Trim(Str(ToVolts)) & " Volts" & vbNewLine & "Step Size = " & _
                Trim(Str(StepSize))
                
        Exit Sub

    End If
    
    If FromVolts < ToVolts And StepSize < 0 Then
    
        MsgBox "From Voltage must be larger than To voltage with a negative volt step size." & _
                vbNewLine & "From = " & Trim(Str(FromVolts)) & " Volts" & vbNewLine & _
                "To = " & Trim(Str(ToVolts)) & " Volts" & vbNewLine & "Step Size = " & _
                Trim(Str(StepSize))

        Exit Sub

    End If
        
    'Count the number of rows to run
    N = Round((ToVolts * StepSize / Abs(StepSize) - FromVolts * StepSize / Abs(StepSize)) _
                        / StepSize * StepSize / Abs(StepSize), _
                    0)
                    
    If N <= 0 Then Exit Sub
                    
    'Set the StartRow
    StartRow = CurrentRow
                
    If StartRow > 1 Then
    
        With Me.gridVoltageCal
        
        
            .row = StartRow - 1
            .Col = 1
            
            If Abs(val(.text) - FromVolts) < 0.0001 Then
            
                'We're repeating a step at the same voltage
                'as the last voltage of a previously added step set
                'Quick fix - reduce StartRow by 1
                StartRow = StartRow - 1
                
            End If
                        
        End With
        
    End If
        
    Volts = FromVolts
    i = StartRow
        
    Do While Volts <= ToVolts And Volts >= FromVolts
    
        With Me.gridVoltageCal
    
            If i >= .Rows Then
            
                .Rows = i + 1
    
            End If
    
            .row = i
            .Col = 1
            .text = Format(Volts, "#0.0#####")
                       
            .RowExpanded = True
            If .RowHeight(i) = 0 Then .RowHeight(i) = 228
            
            .Col = 0
            .text = Trim(Str(i))
                
            'Increment i
            i = i + 1
            
            'Increment volts
            Volts = Volts + StepSize
                
        End With
        
        DoEvents
        
    Loop
    
    'Resize the first two columns of the grid
    ResizeGrid Me.gridVoltageCal, _
               Me, , , _
               0, _
               1
                   
    If i > 1 Then
        
        Me.gridVoltageCal.TopRow = i - 1
        
    End If
        
    CurrentRow = i
        
End Sub

Private Sub cmdAddSingle_Click()

    Dim N As Long

    With Me.gridVoltageCal
    
        'Determine the number of rows in the calibration grid
        N = .Rows
        
        'Check to see if the top-most row is blank
        If .TextMatrix(N - 1, 1) = "" Or _
           .TextMatrix(N - 1, 1) = vbNullString _
        Then
        
            'Assign the value of the single step to the empty row
            .TextMatrix(N - 1, 1) = Trim(Me.txtSingleStep.text)
            
        Else
        
            'Add a new row
            .Rows = N + 1
            
            'Assign the value of the single step text-box
            'to the new row
            .TextMatrix(N, 1) = Trim(Me.txtSingleStep.text)
            
            'Add a new column number for column #0
            .TextMatrix(N, 0) = N
            
        End If
            
        'Sort the grid
        SortGrid Me.gridVoltageCal, _
                 Me, _
                 1, _
                 .Rows - 1, _
                 1, _
                 .Cols - 1
                 
        'Set the top row to the last row
        .TopRow = .Rows - 1
                   
    End With
    
End Sub

Private Sub cmdApply_Click()

    Dim UserResp As Long

    'Prompt the user and ask them if they really want to do this
    UserResp = MsgBox("Making these changes could break the IRM system!" & vbNewLine & vbNewLine & _
                      "Do you still want to go ahead and make them?", _
                      vbYesNo, _
                      "Warning!")
                      
    'Check for a 'No' answer
    If UserResp = vbNo Then
    
        Exit Sub
        
    End If

    'Save the two conversion factors
    modConfig.PulseMCCVoltConversion = val(Me.txtDAQOutputConversion)
    modConfig.PulseReturnMCCVoltConversion = val(Me.txtPulseReturnConversion)
    
    'Save the max capacitor voltage settings
    modConfig.IRMAxialVoltMax = val(Me.txtIRMAxialMaxCapVolts)
    modConfig.AxialTransMaxCapVoltsSame = (Me.chkSameAsAxial = Checked)
    If modConfig.AxialTransMaxCapVoltsSame = False Then
    
        modConfig.IRMTransVoltMax = val(Me.txtIRMTransMaxCapVolts)
        
    End If
    
    'Set unsaved changes = false
    UnsavedChanges = False

End Sub

Private Sub cmdCalcConvFactor_Click()

    Dim PulseReturnArray() As Double
    Dim CapacitorVoltsArray() As Double
    Dim DAQOutputArray() As Double
    
    Dim TempD As Double
    Dim MaxCapVolts As Double
    
    Dim N As Long
    Dim i As Long
    Dim j As Long
    
    Dim MaxDAQOutReached As Boolean
    Dim MaxCapDispReached As Boolean
    
    Dim Slope As Double
    Dim R2error As Double

    'This button will only work if the calibration is done
    If CalStatus = "DONE" Then
    
        'Resize Calibration Arrays so that they each have only one element
        ReDim PulseReturnArray(1)
        ReDim DAQOutputArray(1)
        ReDim CapacitorVoltsArray(1)
        
        'Start MaxDAQOutReached at false because the maximum output DAQ voltage hasn't been
        'reached yet
        MaxDAQOutReached = False
        
        'Ditto for the maximum Capacitor display voltage boolean status flag
        MaxCapDispReached = False
        
        'For the same reason, start the Max Capacitor display voltage local var
        'out at zero
        MaxCapVolts = 0
        
        With Me.gridVoltageCal
                
            'Get the number of data rows in the calibration grid table
            N = .Rows - 1
        
            'Set j, the index of the three 1D arrays, to zero
            j = 0
        
            'Load the grid rows into the calibration array
            For i = 1 To N
        
                .row = i
                .Col = 4
                If .text = vbNullString Then
                
                    'User did not enter a IRM Capacitor voltage for this row
                    'Go to the next row
                    
                Else
                    
                    'If j + 1 > size of arrays, add a new row to each array
                    If j + 1 > UBound(CapacitorVoltsArray) Then
                    
                        ReDim Preserve PulseReturnArray(j + 1)
                                            
                        'Need to check to see if we've reached the maximum capacitor
                        'display voltage reported
                        If Not MaxCapDispReached Then
                        
                            ReDim Preserve CapacitorVoltsArray(j + 1)
                        
                        End If
                                            
                        'Need to see if the output voltage has saturated at the
                        'max allowed output voltage
                        If Not MaxDAQOutReached Then
                        
                            ReDim Preserve DAQOutputArray(j + 1)
                            
                        End If
                   
                    End If
                    
                    'Need to check to see if we've reached the maximum capacitor
                    'display voltage reported
                    If Not MaxCapDispReached Then
                                        
                        'Store the capacitor display voltage
                        TempD = val(.text)
                        CapacitorVoltsArray(j) = TempD
                    
                        'If the current capacitor voltage is greater than the prior
                        'maximum capacitor voltage, then store it to MaxCapVolts
                        If TempD > MaxCapVolts Then
                        
                            MaxCapVolts = TempD
                            
                        ElseIf CInt(TempD) <= CInt(MaxCapVolts) Then
                        
                            'We've hit a plateau in the MaxCapVolts
                            'which means that the DAQ Output Voltage conversion
                            'factor is too low
                            MaxCapDispReached = True
                            
                        End If
                        
                    End If
                        
                    'Store the return voltage
                    .Col = 3
                    PulseReturnArray(j) = val(.text)
                    
                    'If the DAQ output voltage has already saturated up to 10 volts
                    'then do not record more values, they are meaningless
                    If Not MaxDAQOutReached Then
                        
                        'Store the DAQ output voltage
                        .Col = 2
                        DAQOutputArray(j) = val(.text)
                        
                        'Check to see if the voltage voltage is greater than or equal
                        'to the maximum allowed voltage for the board used to output
                        'the DAQ IRM set voltage
                        If val(.text) >= modConfig.PulseVoltMax _
                        Then MaxDAQOutReached = True
                    
                    End If
                    
                    'Iterate j
                    j = j + 1
                    
                End If
            
            Next i
            
            N = UBound(CapacitorVoltsArray)
            
            'Check to see if the max capacitor display voltage has saturated
            If MaxCapDispReached = True Then
            
                'If there are only two elements in CapacitorVoltsArray, then
                'the capacitor display volts saturated on the first IRM ramp
                'in the calibration table (bad!)
                If N = 2 Then
                
                    'Need to tell the user that the current Pulse return voltage
                    'conversion factor is way too high
                    MsgBox "Current Pulse Return voltage conversion factor is too high." & _
                           vbNewLine & vbNewLine & _
                           "Please lower the current value and try the IRM voltage " & _
                           "calibration again.", , _
                           "Whoops!"
                           
                    Exit Sub
                    
                End If
                
                'Otherwise, there are more than two values
                'Do Linear least squares on the elements 0 ... N - 3 of the
                'Pulse Return Array versus the elements 0 ... N - 3 of the Capacitor
                'display voltage array
                modAF_DAQ.LinearLeastSquares CapacitorVoltsArray(), _
                                         PulseReturnArray(), _
                                         N - 2, _
                                         0, _
                                         TempD, _
                                         Slope, _
                                         R2error
                                         
            Else
            
                'No saturation of the Capacitor display voltage happened
                
                'Now need to do least squares with Capacitor voltages VS return volts
                modAF_DAQ.LinearLeastSquares CapacitorVoltsArray(), _
                                             PulseReturnArray(), _
                                             N, _
                                             0, _
                                             TempD, _
                                             Slope, _
                                             R2error
                                             
            End If
                                             
            'Save Slope & error to Pulse Output & Return voltage calibration values in form
            'The two values should be the same
            Me.txtPulseReturnConversion = Format(Slope, "#0.0####")
            Me.txtReturnError = Format((1 - R2error) * Slope, "#0.0####")
            
            'Now need to get the DAQ output voltage conversion to use
            'For this, need to use data points before the max output voltage is reached
            'which will be the last element in the DAQOutputArray
            N = UBound(DAQOutputArray)
            
            'Check to see if the DAQ output voltage values saturates
            If MaxDAQOutReached = True Then
                
                'if the DAQOutputArray is only one element long, then need to
                'prompt the user to lower the current DAQ Output Conversion factor
                If N = 1 Then
                
                    MsgBox "Current DAQ Output Voltage conversion factor is too high." & _
                           vbNewLine & vbNewLine & _
                           "Please lower the value and try the IRM Voltage calibration again.", , _
                           "Whoops!"
                           
                    Exit Sub
                    
                End If
                
                'Can't use linear least squares.  Because the Pulse return voltage
                'is the controlling factor for what Capacitor display voltage is reached
                'the slope for the DAQ Output voltages vs the Capacitor display voltages
                'is NOT the DAQ output conversion factor.
                'We need to extrapolate to get the
                
                'What is the last, non-saturated value in the DAQOutputArray?
                TempD = DAQOutputArray(N - 1)
                
                'Filter for a zero value in the Pulse volt maximum
                If modConfig.PulseVoltMax = 0 Then
                
                    'Send error message
                    UserResponse = frmDialog.DialogBox( _
                                        "The IRM DAQ Ouput Maximum allowed voltage is currently set " & _
                                        " to 0.  This value must be changed for the IRM to become operational." & _
                                        vbNewLine & vbNewLine & _
                                        "Would you like to go to the Settings window and change the IRM maximum " & _
                                        "DAQ output voltage?  (The default value is 10 Volts.)", _
                                        "Bad IRM Setting!", _
                                        2, _
                                        "Yes", _
                                        "No")
                                        
                    'If user responds yes, need to load frmSettings and set the current
                    'tab to the IRM settings tab (Tab #7)
                    If UserResponse = vbYes Then
                    
                        'Load the settings form
                        Load frmSettings
                        
                        'Set the tabs to tab #7
                        frmSettings.selectTab 7
                        
                        'Show the settings form and send it to the front
                        frmSettings.Show
                        frmSettings.ZOrder 0
                        
                    End If
                        
                    'Exit this subroutine
                    Exit Sub
                    
                End If
                
                'Now divide this value by the maximum DAQ output voltage allowed
                TempD = TempD / modConfig.PulseVoltMax
                
                'Multiply this value by the maximum Capacitor Display Voltage allowed
                'I'm assuming the max cap. voltage for the Axial IRM will always be greater
                'than or equal to the Transverse IRM max capacitor voltage
                TempD = TempD * modConfig.IRMAxialVoltMax
                
                'TempD now stores the Capacitor voltage that corresponds to the
                '(N - 1)th element of the DAQOutputArray,
                Me.txtDAQOutputConversion.text = Trim(Str(DAQOutputArray(N - 1) / TempD))
                
            Else
            
                'The DAQ output voltages did not saturate, can use all the elements
                'of the DAQ output voltage array
                modAF_DAQ.LinearLeastSquares CapacitorVoltsArray(), _
                                             DAQOutputArray(), _
                                             N, _
                                             0, _
                                             TempD, _
                                             Slope, _
                                             R2error
                
            End If
            
            'The Generated Slope = DAQ output conversion factor
            Me.txtDAQOutputConversion.text = Trim(Str(Slope))
                                    
        End With
        
        'Set unsaved changes = false
        UnsavedChanges = False
        
    End If
          
End Sub

Private Sub cmdCancel_Click()

    'Interrupt the IRM charging cycle
    frmIRMARM.IRMInterruptCharge
    
    'Close the picture box
    Me.picGetCapacitorVoltage.Visible = False

    'Set the cancel voltage flag to true
    VoltageCancelled = True
    VoltageAccepted = False

End Sub

Private Sub cmdClear_Click()

    'Set Current Row to 1
    CurrentRow = 1

    'Clear the flex grid and redo column headers
    With Me.gridVoltageCal
    
        'Clear the Flex - grid
        .Clear
        .ClearStructure
                
        'Redo the headers
        .Rows = 2
        .Cols = 5
        .FixedRows = 1
        .FixedCols = 1
        
        .Col = 1
        .row = 0
        .text = "Capacitor Volt."
            
        .Col = 2
        .text = "DAQ Output Volt."
            
        .Col = 3
        .text = "Pulse Return Volt."
                
        .Col = 4
        .text = "Capacitor Display Volt."
                        
        .Col = 0
        .row = 1
        .text = "1"
                
        'Resize grid
        ResizeGrid Me.gridVoltageCal, _
                   Me
                                  
    End With
    
    'Set unsaved changes = false
    UnsavedChanges = False

End Sub

Private Sub cmdClose_Click()

    Dim UserResp As Long
    
    'If user has unsaved changes on this form, need to prompt them
    'before they close the window to save those changes
    If UnsavedChanges = True Then
    
        UserResp = MsgBox("You have made changes to the IRM system without saving them." & _
                          vbNewLine & "Would you like to exit this window anyways?", _
                          vbYesNo, _
                          "Unsaved User Changes!")
                          
        If UserResp = vbNo Then
    
            Exit Sub
            
        End If
        
    End If
    
    UnsavedChanges = False
    
    Me.Hide

End Sub

Private Sub cmdDelete_Click()

    Dim RowStart As Long
    Dim i As Long
    Dim j As Long
    Dim RowEnd As Long

    With Me.gridVoltageCal
    
        'Find bounds on the rows to eliminate
        If .row > .RowSel Then
        
            RowStart = .RowSel
            RowEnd = .row
            
        Else
        
            RowStart = .row
            RowEnd = .RowSel
        
        End If
        
        'Delete selected rows, renumber Col #0, and resize the grid
        DeleteRow Me.gridVoltageCal, _
                  Me, _
                  RowStart, _
                  RowEnd, _
                  True, _
                  True
        
        If .Rows = 1 Then
        
            'Set unsaved changes = false
            UnsavedChanges = False
            
        End If
        
    End With
    
End Sub

Private Sub cmdHelp_Click()

    Dim i As Integer
    Dim UserResp As Long
    i = 1
    
    Do While i < 8
        
        Select Case i
        
            Case 1
                
                 'This will pop-up the frmDialog and tell the user how to use this form.
                 UserResp = frmDialog.DialogBox("How to use the IRM Voltage Calibration tool." & vbNewLine & vbNewLine & _
                              "1) Before using this tool, you must have all of your IRM hardware installed, " & _
                              "and properly connected, including:" & vbNewLine & vbTab & _
                              "a. IRM Capacitor Box" & vbNewLine & vbTab & _
                              "b. Grey Capacitor Box with Relay switches connected to a 2G Demag Box or an ADWIN board." & _
                              vbNewLine & vbTab & _
                              "c. Cables connecting IRM capacitor box control voltage input and read voltage output to the " & _
                              "wiring junction box for the Measurement Computing PCI-DAS6030 board.", _
                              "How to use the IRM Voltage Calibration tool.", _
                              1, _
                              "Continue")
                              
                If UserResp = vbYes Then i = i + 1
                
            Case 2
                      
         UserResp = frmDialog.DialogBox("2) This tool will allow you to calibrate the control voltage sent to the IRM capacitor box, " & _
                      "and the read return voltage from the box that the computer uses to determine the current " & _
                      "voltage charge on the IRM capacitor box.", _
                      "How to use the IRM Voltage Calibration tool.", _
                      3, _
                      "Next", "Back", "Exit")
                      
                If UserResp = vbYes Then i = i + 1
                If UserResp = vbNo Then i = i - 1
                If UserResp = vbCancel Then i = 8
                      
            Case 3
            
        UserResp = frmDialog.DialogBox("3) The operating voltage range for the IRM capacitor box is usually 2 - 450 V." & vbNewLine & _
                      "4) To begin, enter in a voltage step size (best usually 10 - 50 V)." & vbNewLine & _
                      "5) Then enter a from voltage (suggested >= 5 V)." & vbNewLine & _
                      "6) Then enter a to voltage (suggested <= 400 V).", _
                      "How to use the IRM Voltage Calibration tool.", _
                      3, _
                      "Next", "Back", "Exit")
                      
                If UserResp = vbYes Then i = i + 1
                If UserResp = vbNo Then i = i - 1
                If UserResp = vbCancel Then i = 8
                      
            Case 4
                      
         UserResp = frmDialog.DialogBox("7) You can do this multiple times. (i.e. 5 - 55 V in 10 V steps, then 55 - 155 V in 25 V steps, and then 155 - 405 V in 50 V steps)" & vbNewLine & _
                      "8) You can also enter in single voltage steps using the 'Add Single Voltage' button." & vbNewLine & _
                      "9) You really only need about 8 - 10 data points (ranging from 10 - 400 V) to get a good fit for the voltage calibrations.", _
                      "How to use the IRM Voltage Calibration tool.", _
                      3, _
                      "Next", "Back", "Exit")
                      
                If UserResp = vbYes Then i = i + 1
                If UserResp = vbNo Then i = i - 1
                If UserResp = vbCancel Then i = 8
           
            Case 5
         UserResp = frmDialog.DialogBox("10) If this is the 1st time you're doing a voltage calibration, you will need to make a first guess as to what your " & _
                      "voltage calibration constants will be." & vbNewLine & _
                      "If this guess is really off, you may need to run the calibration several times to approach the correct values." & _
                      vbNewLine & vbTab & _
                      "If the calibration is really off, then the maximum Capacitor voltage step that you entered may result in an actual IRM Capacitor Box voltage " & _
                      "that is > 450 V.  However, this is not possible, the code will truncate the voltage at 450 V. This will skew the fit for the calibration." & vbNewLine & vbTab & _
                      "If this happens, you will need to adjust the voltage calibration constants down and retry the calibration." & _
                      vbNewLine & vbTab & _
                      "Similarly, if the calibration constants that you initially guess are too low, then the resulting IRM capacitor values will be too low " & _
                      "and will not give good fit results.", _
                      "How to use the IRM Voltage Calibration tool.", _
                      3, _
                      "Next", "Back", "Exit")
                                  
                If UserResp = vbYes Then i = i + 1
                If UserResp = vbNo Then i = i - 1
                If UserResp = vbCancel Then i = 8
           
           Case 6
                      
         UserResp = frmDialog.DialogBox("11) When you're finished loading your steps, but before you click 'Start Calibration' at the bottom " & _
                      "of the window: " & vbNewLine & vbNewLine & vbTab & "Make sure that you have someone able to watch the IRM capacitor box and see the charge " & _
                      "voltage read out on it.", _
                      "How to use the IRM Voltage Calibration tool.", _
                      3, _
                      "Next", "Back", "Exit")
                      
                If UserResp = vbYes Then i = i + 1
                If UserResp = vbNo Then i = i - 1
                If UserResp = vbCancel Then i = 8
          
          Case 7
         UserResp = frmDialog.DialogBox("12) Click 'Start Calibration' and the instructions on the screen will prompt you on how to " & _
                      "do the rest.  Be prepared to have the person watching the IRM Capacitor Voltage display read out the " & _
                      "maximum voltage that the Box charges to for each IRM charging step.", _
                      "How to use the IRM Voltage Calibration tool.", _
                      2, _
                      "Back", "Finish")
                      
                                      
                If UserResp = vbYes Then i = i - 1
                If UserResp = vbNo Then i = 8
        End Select
    Loop

End Sub

Private Sub cmdOK_Click()

    Dim UserResp As Long

    'Prompt the user and ask them if they really want to do this
    UserResp = MsgBox("Making these changes could break the IRM system!" & vbNewLine & vbNewLine & _
                      "Do you still want to go ahead and make them?", _
                      vbYesNo, _
                      "Warning!")
                      
    'Check for a 'No' answer
    If UserResp = vbNo Then
    
        Exit Sub
        
    End If
    
    modConfig.PulseMCCVoltConversion = val(Me.txtDAQOutputConversion)
    modConfig.PulseReturnMCCVoltConversion = val(Me.txtPulseReturnConversion)
    
    'Save the max capacitor voltage settings
    modConfig.IRMAxialVoltMax = val(Me.txtIRMAxialMaxCapVolts)
    modConfig.AxialTransMaxCapVoltsSame = (Me.chkSameAsAxial = Checked)
    If modConfig.AxialTransMaxCapVoltsSame = False Then
    
        modConfig.IRMTransVoltMax = val(Me.txtIRMTransMaxCapVolts)
        
    End If
    
    'Now, change the relavent fields in the .ini file
    'NOTE:  This is the only the second time in the newly written AF/IRM code where the .ini file
    '       is edited outside of modConfig.
    Config_SaveSetting "IRMPulse", _
                       "PulseMCCVoltConversion", _
                       Trim(Str(modConfig.PulseMCCVoltConversion))
                       
    Config_SaveSetting "IRMPulse", _
                       "PulseReturnMCCVoltConversion", _
                       Trim(Str(modConfig.PulseReturnMCCVoltConversion))
                       
    Config_SaveSetting "IRMAxial", _
                       "IRMAxialVoltMax", _
                       Trim(Str(modConfig.IRMAxialVoltMax))
                       
    Config_SaveSetting "IRMTransverse", _
                       "IRMTransVoltMax", _
                       Trim(Str(modConfig.IRMTransVoltMax))
    
    Config_SaveSetting "IRMPulse", _
                       "AxialTransMaxCapVoltsSame", _
                       Trim(Str(modConfig.AxialTransMaxCapVoltsSame))
                       
                       
    'Set unsaved changes = false
    UnsavedChanges = False
                       
End Sub

Private Sub cmdPause_Click()

    'Check to see which mode the button is in
    If cmdPause.Caption = "Pause Cal." And _
       CalStatus = "RUNNING" _
    Then
    
        'Change the calibration status
        CalStatus = "PAUSED"
        
        'Change the button caption
        cmdPause.Caption = "Resume Cal."
        
    ElseIf cmdPause.Caption = "Resume Cal." And _
           CalStatus = "PAUSED" _
    Then

        'Change the calibration status
        CalStatus = "RUNNING"
        
        'Change the button caption
        cmdPause.Caption = "Pause Cal."
        
    End If

End Sub

Private Sub cmdRedo_Click()

    'Hide the Redo & accept buttons
    cmdRedo.Visible = False
    cmdAccept.Visible = False

    'Do the IRM charge and pulse cycle again
    With Me.gridVoltageCal
    
        'Use current row to access the IRM voltage pulse value
        'from the row that was just used to do the last IRM pulse
        .row = CurrentRow
        .Col = 1
        frmIRMARM.FireIRM val(.text), True
        
        cmdRedo.Visible = True
        cmdAccept.Visible = True

    End With
    
End Sub

Private Sub cmdStartStop_Click()

    Dim i As Long
    Dim N As Long
    
    'Check to see if the user is clicking to end the calibration
    'or start it
    If cmdStartStop.Caption = "Start Calibration" Then
    
        'User has chosen to start the calibration
        CalStatus = "RUNNING"
        
        'Change the Caption
        cmdStartStop.Caption = "End Calibration"
                
    Else
    
        'User has chosen to end the calibration
        CalStatus = "ENDED"
        
        'Change the Caption
        cmdStartStop.Caption = "Start Calibration"
        
        'Exit the sub
        Exit Sub
        
    End If
    
    'Check the flow status of the calibration
    If Do_FlowControl = False Then Exit Sub
    
    'Set Unsaved changes = True
    UnsavedChanges = True
    
    'Lock the Coils
    CoilsLocked = True
    Me.chkLockCoils.Value = Checked
    
    'Get the number of rows of calibration steps to run
    N = Me.gridVoltageCal.Rows - 1
    
    'Do the for loop now
    For i = 1 To N
    
        With Me.gridVoltageCal
        
            'Check the flow status of the calibration
            If Do_FlowControl = False Then Exit Sub
        
            'Get the IRM set voltage
            
            'Set the voltage Accepted or cancelled flag to false
            VoltageAccepted = False
            VoltageCancelled = False
            
            'This statement is needed in case the user deletes rows
            'from the IRM DAQ voltage calibration table during the calibration
            'run
            If i > .Rows - 1 Then Exit For
            
            'Set the Current row
            CurrentRow = i
            
            .row = i
            .Col = 1
            
            'Set the Pulse voltage in the IRM ARM form
            frmIRMARM.txtPulseVolts = Trim(.text)
            
            'Set the Backfield IRM checkbox
            'IRM backfield checkbox should be OFF
            frmIRMARM.chkBackfield.Value = Unchecked
            
            'Set the IRM radio buttons
            If Me.optCoil(0).Value = True Then frmIRMARM.optCoil(0).Value = True
            If Me.optCoil(1).Value = True Then frmIRMARM.optCoil(1).Value = True
            
            'Fire the IRM pulse in calibration mode
            frmIRMARM.FireIRM val(frmIRMARM.txtPulseVolts), True
            
            'Show the accept and redo buttons
            Me.cmdAccept.Visible = True
            Me.cmdRedo.Visible = True
            
            'Wait for the user to do something - either accept or cancel the voltage
            Do

                'Let external events happen
                DoEvents

                'Loop every 20 ms
                PauseTill timeGetTime() + 20

            Loop Until VoltageAccepted = True Or _
                       VoltageCancelled = True
                       
            'If the voltage was accepted, put it into the grid
            If VoltageAccepted = True Then
            
                .Col = 4
                .text = Trim(Str(Me.txtCapacitorVoltage))
                
                If .ColWidth(4) < Me.TextWidth(.text) * 1.2 Then
                
                    .ColWidth(4) = Me.TextWidth(.text) * 1.2
                    
                End If
                
            End If
            
            'Go to the next iteration of the loop
            Me.picGetCapacitorVoltage.Visible = False
                        
        End With
        
        'Check the flow status of the calibration
        If Do_FlowControl = False Then Exit Sub
        
    Next i
    
    CalStatus = "DONE"
    
    'Reset the Start Stop button caption
    Me.cmdStartStop.Caption = "Start Calibration"
    
    'Activate the Calculate conversion factor button
    Me.cmdCalcConvFactor.Enabled = True

    'Unlock the coils
    CoilsLocked = False
    Me.chkLockCoils.Value = Unchecked

End Sub

Private Function Do_FlowControl() As Boolean

    If CalStatus = "ENDED" Or _
       CalStatus = "DONE" _
    Then
               
        Do_FlowControl = False
        
        'If coil selected is an IRM coil, then set IRMChargeInterrupted
        frmIRMARM.IRMInterruptCharge
                
    ElseIf CalStatus = "RUNNING" Then
            
        Do_FlowControl = True
        
    ElseIf CalStatus = "PAUSED" Then
    
        'Go into a do loop until cal-status changes
        'to running or ended
        Do
        
            DoEvents
            
            'Wait 20 ms between loops
            PauseTill timeGetTime() + 20
            
        Loop Until CalStatus = "RUNNING" Or _
                   CalStatus = "ENDED"
                   
        If CalStatus = "ENDED" Then
        
            Do_FlowControl = False
            
        Else
        
            Do_FlowControl = True
            
        End If
                   
    End If
    
End Function

Private Sub Form_Activate()

    'If IRM modules are disabled, need to disable auto-calibration
    'function on this form
    If EnableAxialIRM = False And _
       EnableTransIRM = False _
    Then
    
        'Msgbox the user
        MsgBox "IRM Axial & Transverse modules are currently disabled.  Thus, the " & _
               "DAQ output voltage vs IRM Box voltage calibration cannot be done right now.", , _
               "Whoops!"
               
        'disable all the relevant buttons
        Me.cmdAdd.Enabled = False
        Me.cmdStartStop.Enabled = False
        Me.cmdPause.Enabled = False
        
    Else
    
        'Enable all the relevant buttons
        Me.cmdAdd.Enabled = True
        Me.cmdStartStop.Enabled = True
        Me.cmdPause.Enabled = True
        
    End If

    'Set the coils-locked status
    If CoilsLocked = True Then Me.chkLockCoils.Value = Checked
    If CoilsLocked = False Then Me.chkLockCoils.Value = Unchecked
    
    'Set active coil system
    If ActiveCoilSystem = AxialCoilSystem Then
    
        optCoil(0).Value = True
        
    ElseIf ActiveCoilSystem = TransverseCoilSystem Then
    
        optCoil(1).Value = True
        
    Else
    
        'No coil selected
        optCoil(0).Value = False
        optCoil(1).Value = False
        
        ActiveCoilSystem = NoCoilSystem
        
    End If

    'Update the Max Capacitor volts settings display
    Me.txtIRMAxialMaxCapVolts = Trim(Str(modConfig.IRMAxialVoltMax))
    If modConfig.AxialTransMaxCapVoltsSame = True Then
    
        Me.chkSameAsAxial = Checked
        Me.txtIRMTransMaxCapVolts = Me.txtIRMAxialMaxCapVolts
        Me.txtIRMTransMaxCapVolts.Enabled = False
        
    Else
    
        Me.chkSameAsAxial = Unchecked
        Me.txtIRMTransMaxCapVolts.Enabled = True
        Me.txtIRMTransMaxCapVolts = Trim(Str(modConfig.IRMTransVoltMax))
        
    End If

End Sub

Private Sub Form_Load()

    'Set the form width & height
    Me.Height = 7440
    Me.Width = 8265
    
    '(March 11, 2011 - I Hilburn)
    'Update all the form's icons with the program Ico file
    SetFormIcon Me
    
    Me.Caption = "IRM / DAQ Voltage Calibration"
    
    'If IRM modules are disabled, need to disable auto-calibration
    'function on this form
    If EnableAxialIRM = False And _
       EnableTransIRM = False _
    Then
       
        'disable all the relevant buttons
        Me.cmdAdd.Enabled = False
        Me.cmdStartStop.Enabled = False
        Me.cmdPause.Enabled = False
        
    Else
    
        'Enable all the relevant buttons
        Me.cmdAdd.Enabled = True
        Me.cmdStartStop.Enabled = True
        Me.cmdPause.Enabled = True
        
    End If

    'Set Unsaved changes = False
    UnsavedChanges = False

    'Set the current row = 0
    CurrentRow = 1

    'Set Pulse cancelled = False
    PulseCancelled = True

    'Set isUserChange = True
    isUserChange = True
    
    'Hide the Get capacitor voltage picture box
    Me.picGetCapacitorVoltage.Visible = False
    
    'Set the coils-locked status
    If CoilsLocked = True Then Me.chkLockCoils.Value = Checked
    If CoilsLocked = False Then Me.chkLockCoils.Value = Unchecked
    
    'Set active coil system
    If ActiveCoilSystem = AxialCoilSystem Then
    
        optCoil(0).Value = True
        
    ElseIf ActiveCoilSystem = TransverseCoilSystem Then
    
        optCoil(1).Value = True
        
    Else
    
        'No coil selected
        optCoil(0).Value = False
        optCoil(1).Value = False
        
       ActiveCoilSystem = NoCoilSystem
        
    End If
    
    'Reset the conversion factor displayed
    Me.txtDAQOutputConversion = Trim(Str(modConfig.PulseMCCVoltConversion))
    Me.txtPulseReturnConversion = Trim(Str(modConfig.PulseReturnMCCVoltConversion))
    Me.txtOutputError = vbNullString
    Me.txtReturnError = vbNullString
    
    'Display the Max Capacitor volts settings
    Me.txtIRMAxialMaxCapVolts = Trim(Str(modConfig.IRMAxialVoltMax))
    If modConfig.AxialTransMaxCapVoltsSame = True Then
    
        Me.chkSameAsAxial = Checked
        Me.txtIRMTransMaxCapVolts = Me.txtIRMAxialMaxCapVolts
        Me.txtIRMTransMaxCapVolts.Enabled = False
        
    Else
    
        Me.chkSameAsAxial = Unchecked
        Me.txtIRMTransMaxCapVolts.Enabled = True
        Me.txtIRMTransMaxCapVolts = Trim(Str(modConfig.IRMTransVoltMax))
        
    End If
        
        
    'Gray out the calculate conversion factor button
    Me.cmdCalcConvFactor.Enabled = False
    
    'Clear the calibration grid
    cmdClear_Click
    
    'Set the captions on the start/stop & pause/resume calibration buttons
    Me.cmdStartStop.Caption = "Start Calibration"
    Me.cmdPause.Caption = "Pause Cal."
            
    SaveSizes
End Sub

Private Sub Form_Resize()
If m_FormWid <> 0 Then
    ResizeControls
    End If
End Sub

Private Sub gridVoltageCal_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = vbRightButton Then
    
        'Need to show the copy option
        PopupMenu mnuIRM

    End If

End Sub

Private Sub mnuIRMCopy_Click()

    Dim i, j As Long
    Dim N, M As Long
    Dim CopyText As String
    Dim Delimeter As String
    
    Delimeter = ","
    CopyText = vbNullString

    With Me.gridVoltageCal
    
        N = .Rows
        M = .Cols
        
        For i = 0 To N - 1
        
            For j = 0 To M - 1
            
                If i = 0 And j = 0 Then
                
                    CopyText = .TextMatrix(i, j)
                    
                ElseIf j = 0 Then
                
                    CopyText = CopyText & .TextMatrix(i, j)
                    
                Else
                
                    CopyText = CopyText & vbTab & .TextMatrix(i, j)
                    
                End If
            
            Next j
            
            CopyText = CopyText & vbNewLine
            
        Next i
        
    End With
    
    'Clear and load the saved text into the clipboard
    Clipboard.Clear
    Clipboard.SetText CopyText
            
End Sub

Private Sub optCoil_Click(Index As Integer)

    Dim UserResponse As Long
    Dim CoilString As String

    'If the coil selection is locked, or this
    'is a triggered, non-user click, exit this event
    If CoilsLocked = True Then Exit Sub

    'Get the coilstring
    If ActiveCoilSystem = AxialCoilSystem Then coisltring = "Axial"
    If ActiveCoilSystem = TransverseCoilSystem Then coisltring = "Axial"
    
    'Is this a coil change?  If so, clear the flex grid,
    If (Index = 0 And _
        ActiveCoilSystem <> AxialCoilSystem) Or _
       (Index = 1 And _
        ActiveCoilSystem <> TransverseCoilSystem) Or _
       (Index <> 1 And _
        Index <> 0) _
    Then
    
        'Are their unsaved changes?
        'If yes, prompt the user to ask them if they really want to lose all these changes
        If UnsavedChanges = True Then
            
            'MsgBox the user to let them know that the current values will be lost unless
            'they are saved
            UserResponse = MsgBox("There are unsaved IRM " & CoilString & " coil pulse vs DAQ " & _
                                  "voltage calibration values.  Switching coils will erase " & _
                                  "all the old values." & vbNewLine & vbNewLine & _
                                  "Are you sure you want to do this?", _
                                  vbYesNo, _
                                  "Unsaved Changes!")
                                  
            'User clicked 'No', exit the sub-routine
            If UserResponse = vbNo Then Exit Sub
        
        End If
                
        cmdClear_Click
        
        'Set unsaved changes = false
        UnsavedChanges = False
                
    End If
    
    'Don't set the coils here, just set the active coil system
    If Index = 0 And _
       optCoil(Index).Value = True _
    Then
    
        ActiveCoilSystem = AxialCoilSystem
        frmIRMARM.optCoil(Index).Value = True
        
    ElseIf Index = 1 And _
           optCoil(Index).Value = True _
    Then
    
        ActiveCoilSystem = TransverseCoilSystem
        frmIRMARM.optCoil(Index).Value = True
        
    Else
    
        ActiveCoilSystem = NoCoilSystem
        
    End If
            
End Sub

' Arrange the controls for the new size.
Private Sub ResizeControls()
Dim i As Integer
Dim ctl As Control
Dim x_scale As Single
Dim y_scale As Single

    ' Don't bother if we are minimized.
    If WindowState = vbMinimized Then Exit Sub

    ' Get the form's current scale factors.
    x_scale = ScaleWidth / m_FormWid
    y_scale = ScaleHeight / m_FormHgt

    ' Position the controls.
    i = 1
    For Each ctl In Controls
        With m_ControlPositions(i)
            If TypeOf ctl Is line Then
                ctl.X1 = x_scale * .Left
                ctl.Y1 = y_scale * .Top
                ctl.X2 = ctl.X1 + x_scale * .Width
                ctl.Y2 = ctl.Y1 + y_scale * .Height
            Else
            If i <> 49 And (i <> 50) Then 'Exception is specific to frmIRM_VoltageCalibration
                ctl.Left = x_scale * .Left
                ctl.Top = y_scale * .Top
                ctl.Width = x_scale * .Width
                If Not (TypeOf ctl Is ComboBox) Then
                    ' Cannot change height of ComboBoxes.
                    ctl.Height = y_scale * .Height
                End If
                On Error Resume Next
                ctl.Font.size = y_scale * .FontSize
                On Error GoTo 0
                End If
            End If
        End With
        i = i + 1
    Next ctl
End Sub

' Save the form's and controls' dimensions.
Private Sub SaveSizes()
Dim i As Integer
Dim ctl As Control

    ' Save the controls' positions and sizes.
    ReDim m_ControlPositions(1 To Controls.Count)
    i = 1
    For Each ctl In Controls
        With m_ControlPositions(i)
            If TypeOf ctl Is line Then
                .Left = ctl.X1
                .Top = ctl.Y1
                .Width = ctl.X2 - ctl.X1
                .Height = ctl.Y2 - ctl.Y1
            Else
                If (i <> 49) And (i <> 50) Then 'Exception is specific to frmIRM_VoltageCalibration
                .Left = ctl.Left
                .Top = ctl.Top
                .Width = ctl.Width
                .Height = ctl.Height
                On Error Resume Next
                .FontSize = ctl.Font.size
                On Error GoTo 0
                End If
            End If
        End With
        i = i + 1
    Next ctl

    ' Save the form's size.
    m_FormWid = ScaleWidth
    m_FormHgt = ScaleHeight
End Sub

Private Sub txtIRMAxialMaxCapVolts_Change()

    UnsavedChanges = True
    
    If Me.chkSameAsAxial.Value = Checked Then
    
        Me.txtIRMTransMaxCapVolts = Me.txtIRMAxialMaxCapVolts
        
    End If

End Sub

Private Sub txtIRMTransMaxCapVolts_Change()

    UnsavedChanges = True

End Sub

