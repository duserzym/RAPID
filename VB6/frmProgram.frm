VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Begin VB.MDIForm frmProgram 
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   Caption         =   "Paleomagnetics Magnetometer Control System"
   ClientHeight    =   9030
   ClientLeft      =   120
   ClientTop       =   795
   ClientWidth     =   11685
   Icon            =   "frmProgram.frx":0000
   LinkTopic       =   "MDIForm1"
   NegotiateToolbars=   0   'False
   WindowState     =   2  'Maximized
   Begin ComCtl3.CoolBar CoolBar1 
      Align           =   1  'Align Top
      Height          =   510
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11685
      _ExtentX        =   20611
      _ExtentY        =   900
      BandCount       =   1
      _CBWidth        =   11685
      _CBHeight       =   510
      _Version        =   "6.7.8988"
      MinHeight1      =   450
      Width1          =   1440
      NewRow1         =   0   'False
      Begin VB.CommandButton cmdToggleNoComm 
         Caption         =   "Turn On NOCOMM Mode"
         Height          =   375
         Left            =   7440
         TabIndex        =   8
         Top             =   80
         Width           =   2055
      End
      Begin VB.CommandButton cmdQuitAndExit 
         BackColor       =   &H000000FF&
         Caption         =   "Quit && EXIT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9960
         MaskColor       =   &H0080FFFF&
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   80
         Width           =   1455
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "Stop Webcam"
         Height          =   255
         Left            =   5760
         TabIndex        =   6
         Top             =   120
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CommandButton cmdStart 
         Caption         =   "Start Webcam"
         Height          =   255
         Left            =   5760
         TabIndex        =   5
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton cmdLogout 
         BackColor       =   &H0080FFFF&
         Caption         =   "Log Out"
         Height          =   255
         Left            =   4080
         MaskColor       =   &H0080FFFF&
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton cmdMagnetometerControl 
         Appearance      =   0  'Flat
         Caption         =   "Magnetometer Control"
         Height          =   255
         Left            =   1920
         TabIndex        =   3
         Top             =   120
         Width           =   1935
      End
      Begin VB.CommandButton cmdFileRegistry 
         Appearance      =   0  'Flat
         Caption         =   "File Registry"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   1695
      End
   End
   Begin ComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   8760
      Width           =   11685
      _ExtentX        =   20611
      _ExtentY        =   476
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   4
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   9499
            Text            =   "Initializing..."
            TextSave        =   "Initializing..."
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   4410
            MinWidth        =   4410
            TextSave        =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   4410
            MinWidth        =   4410
            TextSave        =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            AutoSize        =   2
            Object.Width           =   1667
            MinWidth        =   1676
            TextSave        =   "7:41 PM"
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ComctlLib.ImageList Prog_ImageList 
      Left            =   2280
      Top             =   3840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483644
      MaskColor       =   12632256
      _Version        =   327682
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      NegotiatePosition=   1  'Left
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
      End
      Begin VB.Menu mnuFileLogout 
         Caption         =   "&Log out"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      NegotiatePosition=   1  'Left
      Begin VB.Menu mnuViewStatusBar 
         Caption         =   "Status &Bar"
      End
      Begin VB.Menu mnuViewMeasurement 
         Caption         =   "&Measurement Window"
      End
      Begin VB.Menu mnuViewSampleChanger 
         Caption         =   "Sample &Changer Master List"
      End
      Begin VB.Menu mnuViewQueue 
         Caption         =   "Command &Queue"
      End
      Begin VB.Menu mnuViewStepMonitor 
         Caption         =   "Step Monitor"
      End
      Begin VB.Menu mnuViewDebug 
         Caption         =   "&Debug"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSpacer 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuViewSettings 
         Caption         =   "&Settings..."
      End
      Begin VB.Menu mnuViewCalRod 
         Caption         =   "Calibrate Rod / Run U-Channel Sample"
      End
      Begin VB.Menu mnuDiagAFDataFileSettings 
         Caption         =   "Data File Save Settings"
         Index           =   1
      End
      Begin VB.Menu mnuViewOptions 
         Caption         =   "Options..."
      End
   End
   Begin VB.Menu mnuFlow 
      Caption         =   "&Flow"
      NegotiatePosition=   1  'Left
      Begin VB.Menu mnuFlowRunning 
         Caption         =   "&Running"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuFlowPaused 
         Caption         =   "&Paused"
      End
      Begin VB.Menu mnuFlowHalted 
         Caption         =   "&Halted"
      End
      Begin VB.Menu mnuFlowSpacer 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFlowCodeOverride 
         Caption         =   "Code &Override"
      End
   End
   Begin VB.Menu mnuDiagnostics 
      Caption         =   "&Diagnostics"
      NegotiatePosition=   1  'Left
      Begin VB.Menu mnuDiagDCMotors 
         Caption         =   "DC &Motors"
      End
      Begin VB.Menu mnuDiagSQUID 
         Caption         =   "&SQUID"
      End
      Begin VB.Menu mnuDiagVacuum 
         Caption         =   "&Vacuum"
      End
      Begin VB.Menu mnuSpacer2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDiagAF 
         Caption         =   "AF Demagnetizer"
         Begin VB.Menu mnuDiagAFWindow 
            Caption         =   "&AF Demag Window"
         End
         Begin VB.Menu mnuDiagAFTuner 
            Caption         =   "AF Tuner / ClipTest"
         End
         Begin VB.Menu mnuDiagAFCalibration 
            Caption         =   "AF Field Calibration"
         End
         Begin VB.Menu mnuDiagFileSaveSettings 
            Caption         =   "AF Data File Settings"
         End
      End
      Begin VB.Menu mnuDiagIRMARM 
         Caption         =   "IRM/ARM"
         Begin VB.Menu mnuDiagIRMARMWindow 
            Caption         =   "&IRM / ARM Window"
         End
         Begin VB.Menu mnuDiagIRMFieldCal 
            Caption         =   "IRM Field Calibration"
         End
         Begin VB.Menu mnuDiagIRMVoltCal 
            Caption         =   "IRM Voltage Calibration"
         End
      End
      Begin VB.Menu mnuDiagGaussmeter 
         Caption         =   "&908A Gaussmeter"
      End
      Begin VB.Menu mnuDiagDAQComm 
         Caption         =   "DAQ &Comm"
      End
      Begin VB.Menu mnuDiagSusceptibility 
         Caption         =   "Susceptibility &Bridge"
      End
      Begin VB.Menu mnuSpacer9 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDiagVRM 
         Caption         =   "VRM Data Collection"
      End
      Begin VB.Menu mnuSpacer10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDiagTestEmail 
         Caption         =   "Test Email"
      End
   End
   Begin VB.Menu mnuWindows 
      Caption         =   "&Windows"
      NegotiatePosition=   1  'Left
      WindowList      =   -1  'True
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      NegotiatePosition=   1  'Left
      Begin VB.Menu mnuHelpIndex 
         Caption         =   "&Index"
      End
      Begin VB.Menu mnuHelpShow 
         Caption         =   "&Show"
      End
      Begin VB.Menu mnuSpacer3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "About..."
      End
   End
End
Attribute VB_Name = "frmProgram"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hWnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)
Const TitleBase = "Paleomagnetic Magnetometer Controller System"
Private readySignals As Integer
Private initialized As Boolean
Dim continue As Boolean

Private Sub cmdFileRegistry_click()
    frmSampleIndexRegistry.ZOrder
    frmSampleIndexRegistry.Show
    frmSampleIndexRegistry.SetFocus
End Sub

Private Sub cmdLogout_Click()
    Dim f As Form
    ' (February 2010 L Carporzen) Webcam
    If cmdStop.Visible = True Then
        continue = False
        frmWebcam.Hide
        cmdStart.Enabled = True
        cmdStop.Visible = False
        'Make sure to disconnect from capture source!!!
        DoEvents: SendMessage mCapHwnd, Disconnect, 0, 0
        Clipboard.Clear
        Unload frmWebcam
    End If
    If frmVacuum.VacuumConnectOn = True Then
        
        'automatically turn off the air
        If modConfig.DoDegausserCooling = True Then
            frmVacuum.DegausserCooler False
        End If
                
        MsgBox "Before the system can log you out, " & _
               "please remove any samples from the end of the " & _
               "sample-handler tube.  Otherwise, your sample will be " & _
               "dropped when the vacuum switches off."
           
    End If
    If frmLogin.LoginSucceeded = False Then
    ' (February 2010 L Carporzen) Cleaner way to close the program
        For Each f In Forms
            Unload f
        Next
        
        modLogAFParameters.CloseLogFile
        End ' (October 2007 L Carporzen) Avoid the bug when logout whereas the login window is open.
    Else
        Logout
    End If
    frmLogin.RunSQUID
End Sub

Private Sub cmdMagnetometerControl_click()
    frmMagnetometerControl.ZOrder
    frmMagnetometerControl.Show
    frmMagnetometerControl.SetFocus
End Sub

Private Sub cmdQuitAndExit_Click()

    Dim UserResponse As Long

    'Prompt the user and ask if this is what they really want to do
    UserResponse = MsgBox("Are you sure that you want to exit the Paleomag program right now?" & _
                          vbNewLine & vbNewLine & _
                          "Any unsaved current sample run data will be lost.", _
                          vbYesNo, _
                          "Warning!")
                          
    'If user replies 'No', then exit this event handler
    If UserResponse = vbNo Then Exit Sub
    
    'User has confirmed the quit program command
    mnuFileExit_Click

End Sub

Private Sub cmdStart_Click()
    ' (February 2010 L Carporzen) Webcam
    'If FLAG_MagnetUse = False Then
        On Error GoTo oops
            frmWebcam.Show 'picOutput.Visible = True
            cmdStart.Enabled = False
            cmdStop.Visible = True
            'Setup a capture window (You can replace "WebcamCapture" with watever you want)
            mCapHwnd = capCreateCaptureWindow("WebcamCapture", 0, 0, 0, 320, 240, Me.hWnd, 0)
            'Connect to capture device
            DoEvents: SendMessage mCapHwnd, Connect, 0, 0
            'frmWebcam.tmrMain.Enabled = True
            continue = True
            Do While continue
                SendMessage mCapHwnd, GET_FRAME, 0, 0
                SendMessage mCapHwnd, COPY, 0, 0
                frmWebcam.picOutput.Picture = Clipboard.getData
                DoEvents
                If frmWebcam.picOutput.Picture = 0 Then continue = False
            Loop
        On Error GoTo 0
oops:
        continue = False
        frmWebcam.Hide
        cmdStart.Enabled = True
        cmdStop.Visible = False
        'Make sure to disconnect from capture source!!!
        DoEvents: SendMessage mCapHwnd, Disconnect, 0, 0
        Clipboard.Clear
        Unload frmWebcam
    'End If
End Sub

Private Sub cmdStop_Click()
    ' (February 2010 L Carporzen) Webcam
    continue = False
    frmWebcam.Hide
    cmdStart.Enabled = True
    cmdStop.Visible = False
    'Make sure to disconnect from capture source!!!
    DoEvents: SendMessage mCapHwnd, Disconnect, 0, 0
    Clipboard.Clear
    Unload frmWebcam
End Sub

Private Sub cmdToggleNOCOMM_Click()

    Dim UserResponse As Long

    'Allow the user to toggle the NOCOMM MODE on and off as desired
    If Me.cmdToggleNoComm.Caption = "Turn On NOCOMM Mode" Then
    
        'Confirm this with the user
        UserResponse = MsgBox("Are you sure you want to shut-off all communications " & _
                              "with Paleomag & Rockmag devices connected to this computer?" & _
                              vbNewLine & vbNewLine & "If you are currently running samples, this will " & _
                              "sabotage the run, and is NOT(!!!!) recommended.", _
                              vbYesNo, _
                              "Confirmation Required!")
                              
        'If the user answers no, then abort this event handler
        If UserResponse = vbNo Then Exit Sub
    
        modProg.NOCOMM_MODE = True

        Me.cmdToggleNoComm.Caption = "Turn Off NOCOMM Mode"
        
    Else
    
        Me.cmdToggleNoComm.Caption = "Turn On NOCOMM Mode"
        
        modProg.NOCOMM_MODE = False
        
    End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If cmdStop.Visible = True Then
    'Make sure to disconnect from capture source - if it is connected upon termination the program can become unstable
    DoEvents: SendMessage mCapHwnd, Disconnect, 0, 0
    End If
End Sub

Public Sub Logout()
    Prog_halted = True ' (September 2007 L Carporzen) New version of the Halt button
    readySignals = 0
    initialized = False
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
    frmVacuum.DegausserCooler False
    frmLogin.LoginSucceeded = False
    LoginName = vbNullString
    LoginEmail = vbNullString
    Me.Caption = TitleBase & " (" & MailFromName & ")"
    DoEvents            ' Allow form to refresh first
    FLAG_MagnetInit = False     ' Magnetometer uninitialized
    FLAG_MagnetUse = False      ' Magnetometer not in use
    Set SampQueue = New SampleCommands
    Set SampleIndexRegistry = New SampleIndexRegistrations
    Set SampleHolder = SampleIndexRegistry("!Holder").sampleSet("Holder")
    Set SusceptibilityStandard = SampleIndexRegistry("!Holder").sampleSet("SusStd")
    Set MainChanger = New frmChanger
    MainChanger.IsMasterList = True
    Load MainChanger
    frmTip.ZOrder
    frmLogin.ZOrder
    frmTip.Show
    frmLogin.Show
End Sub

Private Sub MDIForm_Activate()

    'Based on the AF System being used, enable / disable the
    'AF Tuner menu
    If AFSystem = "2G" Then
    
        Me.mnuDiagAFTuner.Enabled = False
        
    ElseIf AFSystem = "ADWIN" Then
    
        Me.mnuDiagAFTuner.Enabled = True
        
    End If

End Sub

Private Sub mdiForm_hide()
    If Me.WindowState <> vbMinimized Then
        Config_SaveSetting "Program", "MainWindowLeft", str$(Me.Left)
        Config_SaveSetting "Program", "MainWindowTop", str$(Me.Top)
        Config_SaveSetting "Program", "MainWindowHeight", str$(Me.Height)
        Config_SaveSetting "Program", "MainWindowWidth", str$(Me.Width)
    End If
End Sub

Private Sub mdiForm_Load()

    '(March 11, 2011 - I Hilburn)
    'Update all the form's icons with the program Ico file
    On Error GoTo BadIcoFile:
    
        Set Me.Icon = LoadPicture(Prog_IcoFile)
        
    On Error GoTo 0
    
      
BadIcoFile:

    LoadResStrings Me
    mnuViewMeasurement.Checked = False
    If DEBUG_MODE Then
       mnuViewDebug.Visible = True
    Else
        mnuViewDebug.Visible = False
    End If
    initialized = False
    Me.Caption = TitleBase & " (" & MailFromName & ")"
    
    'Based on the AF System being used, enable / disable the
    'AF Tuner menu
    If AFSystem = "2G" Then
    
        Me.mnuDiagAFTuner.Enabled = False
        
    ElseIf AFSystem = "ADWIN" Then
    
        Me.mnuDiagAFTuner.Enabled = True
        
    End If
    
End Sub

Private Sub mdiform_Unload(Cancel As Integer)
    'Dim howsthat As VbMsgBoxResult
    'howsthat = MsgBox("Power vacuum off?", vbYesNo)
    'If howsthat = vbYes Then frmVacuum.MotorPower False
    ' Disconnect all connections
    
    If NOCOMM_MODE = True Then Exit Sub
    
    frmDCMotors.MotorCommDisconnect
    
    If frmSQUID.MSCommSquid.PortOpen = True Then
    
        frmSQUID.Disconnect ' (September 2007 L Carporzen) Avoid bug when quiting without logout
        
    End If
    
    ' (July 2010 - I Hilburn)
    ' Added in this IF statement to make sure that the ADWIN board power
    ' to the AF / IRM relays is switched off so that the relays don't overheat
    If AFSystem = "ADWIN" And EnableAF = True Then
    
        'Turn off all the relays controlled from the ADWIN board
        WaveForms("AFRAMPUP").BoardUsed.DigitalOut_ADWIN (63)
        
    ElseIf AFSystem = "2G" And EnableAF = True Then
    
        'Disconnect 2G AF box comm
        frmAF_2G.Disconnect
    
    End If
    
    If frmVacuum.VacuumConnectOn = True And _
       frmMagnetometerControl.cmdManHolder.Enabled = False And _
       frmMagnetometerControl.cmdChangerEdit.Enabled = False _
    Then
    
        ' (September 2007 L Carporzen) Avoid disconecting vacuum accidentaly when quiting
        MsgBox "When you log out, the vaccum will switch off." & vbNewLine & vbNewLine & _
               "Please remove any samples left on the quartz glass tube.", , _
               "Warning!"
                   
    End If
    
    frmVacuum.Disconnect
    
    frmSusceptibilityMeter.Disconnect
    
    'Also disconnect the Gaussmeter is needed
    If frm908AGaussmeter.cmdConnectButton.Caption = "Disconnect" Then
        
        frm908AGaussmeter.Disconnect
        
    End If
    
    'Send a zero to the IRM box
    If modConfig.EnableAxialIRM = True Or _
       modConfig.EnableTransIRM = True _
    Then
    
        SystemBoards(IRMVoltageOut.BoardName).AnalogOut IRMVoltageOut, 0
        
    End If
        
End Sub

Private Sub mnuDiagAFCalibration_Click()

    frmCalibrateCoils.InAFMode = True
    Load frmCalibrateCoils
    frmCalibrateCoils.ZOrder
    frmCalibrateCoils.Show
    
End Sub

Private Sub mnuDiagAFDataFileSettings_Click(Index As Integer)

    Load frmFileSave
    frmFileSave.ZOrder
    frmFileSave.Show

End Sub

Private Sub mnuDiagAFTuner_Click()

    Load frmAFTuner
    frmAFTuner.ZOrder
    frmAFTuner.Show

End Sub

Private Sub mnuDiagAFWindow_Click()

    'Link to the correct form
    'based on the contents of the AFsystem object
    If AFSystem = "2G" Then

        frmAF_2G.ZOrder
        frmAF_2G.Show
        
    ElseIf AFSystem = "ADWIN" Then
    
        'Load the AF ADWIN form
        Load frmADWIN_AF
    
'        Debug.Print "3) Active Coil System: " & Trim(Str(ActiveCoilSystem))
   
    
        'Display the form
        frmADWIN_AF.ZOrder
        frmADWIN_AF.Show
        
'        Debug.Print "4) Active Coil System: " & Trim(Str(ActiveCoilSystem))
   
        
    Else
    
        'What the hey?
        'No other type of AF system is supported right now
        'Message box the user, tell them to change the setting
        'and show frmSettings with the correct tab selected
        MsgBox "You've (somehow) set a non-supported AF System " & _
               "in the Settings window.  This must be changed before " & _
               "the AF & Rock magnetics modules can become active.", _
               vbCritical, _
               "Bad AF System Selected!"
               
        'Load Frm Settings
        Load frmSettings
               
        'Now select the AF tab in frmSettings
        frmSettings.selectTab 4
               
        'Now show the Settings form
        frmSettings.Show
        
    End If

End Sub

Private Sub mnuDiagDAQComm_Click()

    Load frmDAQ_Comm
    frmDAQ_Comm.ZOrder
    frmDAQ_Comm.Show

End Sub

Private Sub mnuDiagDCMotors_Click()
    frmDCMotors.ZOrder
    frmDCMotors.Show
End Sub

Private Sub mnuDiagFileSaveSettings_Click()

    Load frmFileSave
    frmFileSave.ZOrder
    frmFileSave.Show

End Sub

Private Sub mnuDiagGaussmeter_Click()

    Load frm908AGaussmeter
    frm908AGaussmeter.ZOrder
    frm908AGaussmeter.Show

End Sub

Private Sub mnuDiagIRMARMWindow_Click()

    Load frmIRMARM
    frmIRMARM.ZOrder
    frmIRMARM.Show

End Sub

Private Sub mnuDiagIRMFieldCal_Click()

    frmCalibrateCoils.InAFMode = False
    Load frmCalibrateCoils
    frmCalibrateCoils.ZOrder
    frmCalibrateCoils.Show

End Sub

Private Sub mnuDiagIRMVoltCal_Click()

    Load frmIRM_VoltageCalibration
    frmIRM_VoltageCalibration.ZOrder
    frmIRM_VoltageCalibration.Show
    
End Sub

Private Sub mnuDiagSQUID_Click()
    frmSQUID.ZOrder
    frmSQUID.Show
End Sub

Private Sub mnuDiagSusceptibility_Click()
    If FileExists(Prog_IcoFile) And LenB(Prog_IcoFile) > 0 Then frmSusceptibilityMeter.Icon = LoadPicture(Prog_IcoFile) ' (October 2007 L Carporzen)
    frmSusceptibilityMeter.ZOrder
    frmSusceptibilityMeter.Show
End Sub

Private Sub mnuDiagTestEmail_Click()
    frmSendMail.ZOrder
    frmSendMail.Show
End Sub

Private Sub mnuDiagVacuum_Click()
    frmVacuum.ZOrder
    frmVacuum.Show
End Sub

Private Sub mnuDiagVRM_Click()
    frmVRM.ZOrder
    frmVRM.Show
End Sub

Private Sub mnuDiagXYTable_Click()
End Sub

Private Sub mnuFileExit_Click()
    Dim f As Form
    ' (February 2010 L Carporzen) Cleaner way to close the program
    
    'Load and show the shutdown msgform
    Load frmShutdownMsg
    frmShutdownMsg.ZOrder 0
    frmShutdownMsg.Show
    
    'Wait 200 ms
    DelayTime 0.2
    
    '(March 10, 2011 - I Hilburn)
    'This IRM voltage zeroing code is to keep the IRM capacitor voltage from
    'being set to 450 volts every time the code quits.
    'If IRM modules are enabled, set IRM voltage to zero
    On Error GoTo BadIRM:
    
        If EnableAxialIRM Or EnableTransIRM Then
        
            'Set IRM Voltage Out to zero
            frmDAQ_Comm.DoDAQIO IRMVoltageOut, 0
    
        End If
        
    On Error GoTo 0
    
BadIRM:

    For Each f In Forms
     
        If f.name <> "frmShutdownMsg" And _
           f.name <> "frmProgram" _
        Then
         
            Unload f
            
        End If
        
        frmShutdownMsg.ZOrder 0
                  
    Next
    
    'Unload frmProgram
    Unload frmProgram
    
    'Hide and unload the shutdown msg form
    frmShutdownMsg.Hide
    Unload frmShutdownMsg
    
    modListenAndLog.CloseLogFile
    
    modLogAFParameters.CloseLogFile
    
    ' Exit Program
    End
End Sub

Private Sub mnuFileLogout_Click()
    Dim f As Form
    ' (February 2010 L Carporzen) Webcam
    If cmdStop.Visible = True Then
        continue = False
        frmWebcam.Hide
        cmdStart.Enabled = True
        cmdStop.Visible = False
        'Make sure to disconnect from capture source!!!
        DoEvents: SendMessage mCapHwnd, Disconnect, 0, 0
        Clipboard.Clear
        Unload frmWebcam
    End If
    If frmLogin.LoginSucceeded = False Then
    ' (February 2010 L Carporzen) Cleaner way to close the program
        For Each f In Forms
            Unload f
        Next
        
        modListenAndLog.CloseLogFile
        
        modLogAFParameters.CloseLogFile
        End
    Else ' (October 2007 L Carporzen) Avoid the bug when logout whereas the login window is open.
        Logout
    End If
End Sub

Private Sub mnuFlowCodeOverride_Click()
    If StatusCodeColorLevel = CodeGrey Then
        SetCodeLevel StatusCodeColorLevelPrior
    Else
        SetCodeLevel CodeGrey
    End If
End Sub

Private Sub mnuFlowHalted_Click()
    Flow_Halt
    updateFlowMenu
End Sub

Private Sub mnuFlowPaused_Click()
    Flow_Pause
    updateFlowMenu
End Sub

Private Sub mnuFlowRunning_Click()
    Flow_Resume
    updateFlowMenu
End Sub

Private Sub mnuHelpAbout_Click()
    frmAbout.ZOrder
    frmAbout.Show
End Sub

Private Sub mnuViewCalRod_Click()

    'Load and show the frmCalRod window
    frmSettings.cmdCalibNewRod_Click

End Sub

Private Sub mnuViewDAQBoardSettings_Click()

'    Load frmDAQBoardSettings
'    frmDAQBoardSettings.Show
'    frmDAQBoardSettings.ZOrder 0

End Sub

Private Sub mnuViewDebug_click()
    frmDebug.ZOrder
    frmDebug.Show
End Sub

Private Sub mnuViewMeasurement_Click()
    If mnuViewMeasurement.Checked Then
        frmMeasure.Hide
        mnuViewMeasurement.Checked = False
    Else
        frmMeasure.ZOrder
        frmMeasure.Show
        mnuViewMeasurement.Checked = True
    End If
End Sub

Private Sub mnuViewOptions_Click()
    frmOptions.ZOrder
    frmOptions.Show
End Sub

Private Sub mnuViewQueue_Click()
    frmSampleQueueMonitor.ZOrder
    frmSampleQueueMonitor.Show
End Sub

Private Sub mnuViewSampleChanger_Click()
    MainChanger.ZOrder
    MainChanger.Show
End Sub

Private Sub mnuViewSettings_Click()
    frmSettings.ZOrder
    frmSettings.Show
End Sub

Private Sub mnuViewStatusBar_Click()
    If mnuViewStatusBar.Checked Then
        frmProgram.sbStatusBar.Visible = False
        mnuViewStatusBar.Checked = False
    Else
        frmProgram.sbStatusBar.Visible = True
        mnuViewStatusBar.Checked = True
    End If
End Sub

Private Sub mnuViewStepMonitor_Click()
    frmStepMonitor.ZOrder
    frmStepMonitor.Show
End Sub

'------------------------------------------------------------------------------------------------------------------------------'
'------------------------------------------------------------------------------------------------------------------------------'
'
'(July 2010, I Hilburn)
'
'Had to comment these two help menu functions out.
'Microsoft's new Internet controls no longer support VB6's web-browser controls.  Therefore, it's impossible to have
'a web-browser based help form.  (Microsoft = Evil @#$%-heads.)
'
'------------------------------------------------------------------------------------------------------------------------------'
'
'Private Sub mnuHelpIndex_Click()
'    frmHelp.loadHelpFile "index.html"
'    frmHelp.ZOrder
'    frmHelp.Show
'End Sub

'Private Sub mnuHelpShow_Click()
'    frmHelp.loadHelpFile "index.html"
'    frmHelp.ZOrder
'    frmHelp.Show
'End Sub
'
'------------------------------------------------------------------------------------------------------------------------------'
'------------------------------------------------------------------------------------------------------------------------------'

Private Sub mnuViewSystemBoardsettings_Click()

    'Insert code for loading and showing the DAQ Board Settings window

End Sub

Public Sub SetProgramCodeLevel(newLevel As String)
    
    Select Case newLevel
        
        Case CodeRed
            
            BackColor = ColorRed
        
        Case CodeOrange
            
            BackColor = ColorOrange
        
        Case CodeYellow
            
            BackColor = ColorYellow
        
        Case CodeGreen
            
            BackColor = ColorGreen
        
        Case CodeBlue
            
            BackColor = ColorBlue
        
        Case CodeGrey
            
            BackColor = ColorGrey
    
    End Select

End Sub

Public Sub SignalReady()
    readySignals = readySignals + 1
    If readySignals >= 2 Then
        frmMagnetometerControl.Show
        DoEvents            ' Allow form to refresh first
        If Not initialized Then
            Me.Caption = TitleBase & " (" & MailFromName & ") - [" & LoginName & "]"
            frmOptions.txtUserName = LoginName ' (February 2010 L Carporzen) Allow to change Name and email while running
            frmOptions.txtUserEmail = LoginEmail
            If Not NOCOMM_MODE Then Magnetometer_Initialize Else FLAG_MagnetInit = True
            frmMagnetometerControl.EnableMagnetCmds
            initialized = True
        End If
        If DEBUG_MODE Then
            mnuViewDebug.Visible = True
        Else
            mnuViewDebug.Visible = False
        End If
        frmSampleIndexRegistry.Show
    End If
End Sub

Public Sub StatBarNew(statstring As String)
    sbStatusBar.Panels(1).text = statstring
    sbStatusBar.Panels(2).text = vbNullString
    sbStatusBar.Panels(3).text = vbNullString
End Sub

Public Sub StatusBar(statstring As String, barpart As Integer)
    If barpart < 1 Or barpart > 3 Then Exit Sub
    sbStatusBar.Panels(barpart).text = statstring
End Sub

Public Sub updateFlowMenu()
    If Prog_halted Then
        mnuFlowRunning.Checked = False
        mnuFlowPaused.Checked = False
        mnuFlowHalted.Checked = True
    Else
        If Prog_paused Then
            mnuFlowRunning.Checked = False
            mnuFlowPaused.Checked = True
            mnuFlowHalted.Checked = False
        Else
            mnuFlowRunning.Checked = True
            mnuFlowPaused.Checked = False
            mnuFlowHalted.Checked = False
        End If
    End If
End Sub

