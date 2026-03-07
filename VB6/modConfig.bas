Attribute VB_Name = "modConfig"
' This module is included to handle all of the initial
' settings for both paleomag and biomag magnetometers.
' DLL declarations
'
'-------------------------------------------------------------------------------------------------------------------------'
'-------------------------------------------------------------------------------------------------------------------------'
'
'   HUGE Modification
'   Done: July 30, 2010
' Author: Isaac Hilburn
'
' Summary: Stripped out all of the global variables and INI read/write calls for the IRM HF & LF systems
'          and replaced them with variables and calls for the IRM Axial and Transverse systems
'          This code no longer supports a HF IRM setup, but the ADWIN AF system setup will
'          allow IRM to be done on the Transverse coil as well as the Axial.
'
'          Also, AF low-field monitor variables have been stripped out and AF high-field monitor variables
'          renamed to indicate that they are just AF monitor variables.  The AF low-field is unnecessary
'          as there is sufficient resolution and accuracy in the ADWIN system at low monitor voltage levels
'          to use the same ammeter donut for high and low fields.
'
'-------------------------------------------------------------------------------------------------------------------------'
'-------------------------------------------------------------------------------------------------------------------------'

' Modification: April 4, 2010
'       Author: Isaac Hilburn
'
'   These function declarations are now called in the new class module
'   CIniFile.cls
'
'Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
'Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

'Sample changer settings
' Holes are found at every slot divisible by HoleSlotNum
' SlotMax is smallest slot number
' SlotMax is largest slot number
Public SlotMin As Integer, SlotMax As Integer, HoleSlotNum As Integer
' Size of movement between slots for the DC Servo motors. On Lowenstam, one rev is
' 4000 steps, small gear is 22 teeth, large is 50 teeth, and 9 slots
' per big wheel. Note that this is a negative number, because going from
' zero to a positive number will DECREASE the slot number under the changer.
' Therefore, on Lowenstam, OneStep = 4000 * 50 / (22 * 9) = -1010.1010101
Public OneStep As Double
'Program settings
Public IniFile As CIniFile
Public DefaultINI As CIniFile
Public OldIniFile As CIniFile
Public Prog_UsageFile     As String
Public Prog_DefaultBackup As String
Public Prog_DefaultPath   As String
Public Prog_LogoFile      As String
Public Prog_IcoFile       As String ' (October 2007 L Carporzen)
Public Prog_TextEditor    As String
Public Prog_HelpURLRoot   As String
Public ADWIN_AFDataLocalDir As String
Public ADWIN_AFDataBackupDir As String
Public TWOG_AFDataLocalDir As String
Public TWOG_AFDataBackupDir As String
Public AFDoDataFileBackup As Boolean
Public DumpRawDataStats   As Boolean

Public LogMessages        As Boolean
Public LogFolderPath      As String
Public LogFileName        As String

' Calibration variables read from file
Public ZCal         As Double
Public XCal         As Double
Public YCal         As Double
Public IRMPos       As Long
Public SCoilPos     As Long
Public ZeroPos      As Long      ' Zero position in motor steps
Public MeasPos      As Long      ' Measuring Position in motor steps
Public AFPos        As Long      ' AF Position in motor steps
Public FloorPos     As Long      ' Distance from UpDown home position (0 motor counts)
                                 ' to the floor or first obstruction beneath the magnetometer
                                 ' in motor counts!
Public MinUpDownPos As Long      ' the lowest position in motor counts that the up/down carriage can
                                 ' reproducably move.
Public SampleBottom As Long      ' distance from load to base of sample;
Public SampleTop    As Long      ' used to measure height of sample
Public SampleHeight As Long
Public SampleHoleAlignmentOffset As Double
Public LiftSpeedSlow            As Long
Public LiftSpeedNormal          As Long
Public LiftSpeedFast            As Long
Public LiftAcceleration
Public CmdHomeToTop             As Integer
Public CmdSamplePickup          As Integer
Public MotorIDTurning           As Integer
Public MotorIDChanger           As Integer
Public MotorIDChangerY           As Integer
Public MotorIDUpDown            As Integer
Public SCurveFactor             As Integer
Public TurningMotorFullRotation As Long
Public TurningMotor1rps         As Long
Public UpDownMotor1cm           As Double
Public TrayOffsetAngle          As Double
Public UpDownTorqueFactor       As Integer
Public UpDownMaxTorque          As Integer
Public PickupTorqueThrottle     As Double
Public ChangerSpeed             As Long
Public TurnerSpeed              As Long
Public RangeFact                As Double      ' Range factor
Public ReadDelay                As Double      ' (March 2008 L Carporzen) Read delay
Public RemeasureCSDThreshold    As Double
' New parameters for the jumps (April-May 2007 L Carporzen)
Public JumpThreshold   As Double  ' The zero measurements of each SQUID need to be below that constant value
Public StrongMom       As Double  ' Maximum moment where the jumps are controlled by comparing the zeros
Public IntermMom       As Double  ' Intermediate moment where the jump test change from proportionnal to the moment (below) to the constant test
Public MomMinForRedo   As Double  ' Critical moment where the CSD test is cancel
Public JumpSensitivity As Double  ' Proportionnality of the SQUID stability and the measured moment
Public NbTry           As Integer ' Number of try before accept an unverified measurement
Public NbHolderTry     As Integer ' Number of tries before accepting a noisy holder measurement
Public Meascount       As Integer ' Count the number of repeated measurements which don't pass the tests

'New AF demag settings (Mar 2010 - I Hilburn)
'Part of implementation of new AF DAQ coding module
'AFSystem = "2G", or "ADWIN" ("MCC" not supported anymore)
Public AFSystem As String

'New AF demag settings (May 2010 - I Hilburn)
'AF DAQ Implementation
'Four Channel Objects to store the AF Ramp & Monitoring channels
Public AFRampChan As Channel            'Channel for Analog ouput of AF Ramp signal
Public AFMonitorChan As Channel         'Chanel for monitoring AF LC circuit current
Public AltAFMonitorChan As Channel      'Channel for monitoring AF LC circuit current from
                                        ' a second Board / Channel for comparison
                                        '(this alternate channel can also be used to "spy"
                                        'on the 2G AF ramp process to help diagnose problems
                                        'with the 2G system / AF coils
Public AFUnits As String                'Units to use for field calibration of AF & IRM Coils
                                        'Also the units for the user definte low-field max value
                                        
Public ADWINBinFolderPath As String     'Path to the Folder containing the ADWIN board ADBasic
                                        'code files, libraries, and drivers
Public ADWINBootFileName As String      'Name of the file in the ADWIN Bin folder that should be
                                        'used to boot and configure the ADWIN board
Public ADWINRampProgFileName As String  'Name of the file in the ADWIN Bin folder that contains
                                        'the ADBasic code to be loaded into the ADWIN board for the AF ramp cycle

'Global variables to store the current active coil system
'(July 29, 2010 - I Hilburn)
'Changed the Name of the AF Coil system variables and constants by replacing
'the string "CoilSystem" with "AFCoilSystem" throughout the entire Paleomag project
'Also added the public variable "PriorAFCoilSystem" to store the prior setting
'during things like IRM's when the AF coils need to be switched off, but
'may still need track of what the prior coil system was (not certain about this)
Public PriorAFCoilSystem As Integer
Public ActiveCoilSystem As Integer
Global Const AxialCoilSystem = -1
Global Const TransverseCoilSystem = 1
Global Const NoCoilSystem = 0

' Now the AF demag calibration values
Public AFDelay          As Integer
Public AFRampRate       As Integer
Public AFWait           As Double
Public TSlope           As Double
Public Toffset          As Double
Public Thot             As Integer
Public Tmax             As Integer
Public Tunits           As String
Public SystemTUnits       As String
Public AfAxialCoord     As String
Public AfTransCoord     As String

'Global boolean var to record if using the XY Table from APS, if false assume chain drive
Public UseXYTableAPS As Boolean
Public SamplesBetweenHolder As Integer

'(August, 2010 - I Hilburn)
' ADWIN AF Demag Ramp speed control values
Public MinRampUpTime_ms As Long
Public MaxRampUpTime_ms As Long
Public AxialRampUpVoltsPerSec As Double
Public TransRampUpVoltsPerSec As Double
Public MinRampDown_NumPeriods As Long
Public MaxRampDown_NumPeriods As Long
Public RampDownNumPeriodsPerVolt As Long
Public HoldAtPeakField_NumPeriods As Long

'(July, 2010, I Hilburn - New Axial & Trans. Coil Settings
'  for recording tuning & clip test results)
Public AfAxialResFreq As Double
Public AfAxialRampMax As Double
Public AfAxialMonMax As Double
Public AfTransResFreq As Double
Public AfTransRampMax As Double
Public AfTransMonMax As Double

'AF Field Calibration Arrays
'(Mar 2010 - changed from 25 elements, fixed, to dynamic arrays with N x 2 dimension)
Public AFAxial()           As Double
Public AFTrans()           As Double
Public AFAxialCount         As Long '(March 2010 - I Hilburn) - Enables variable number
                                    'of AF field calibration values for Axial Coil
Public AFTransCount         As Long '(March 2010 - I Hilburn) - Enables variable number
                                    'of AF field calibration values for Transverse Coil

'Global boolean var to record if the AF coils have been calibrated with the new system
Public AFTransCalDone As Boolean
Public AFAxialCalDone As Boolean

' Axial Coil Calibration Variables
Public AfAxialYpoint    As Double
Public AfAxialXpoint    As Double
Public AfAxialLowSlope  As Double
Public AfAxialHighSlope As Double
Public AfAxialMax       As Double
Public AfAxialMin       As Double

' Transverse Coil Calibration Variables
Public AfTransYpoint    As Double
Public AfTransXpoint    As Double
Public AfTransLowSlope  As Double
Public AfTransHighSlope As Double
Public AfTransMax       As Double
Public AfTransMin       As Double

' Now the IRM Pulse coil variables
Public IRMAxis          As String
Public IRMBackfieldAxis As String
Public PulseAxialVoltMax       As Double
Public PulseTransVoltMax       As Double
Public TrimOnTrue As Boolean

Public AscSetVoltageMinBoostMultiplier As Double
Public AscSetVoltageMaxBoostMultiplier As Double

'(March 2010 - I Hilburn)
'Modified IRM pulse field calibration arrays to make dynamic and N x 2 in dimension
'PulseAxialCount & PulseTransCount = number of field calibration values in .INI file to
'be loaded to the pulse arrays
Public IRMSystem As String
Public PulseVoltMax       As Double
Public PulseAxial()       As Double
Public PulseTrans()       As Double
Public PulseAxialCount    As Long '(June 2010 - I Hilburn) - Enables variable number of
                                    'field calibration values for IRM axial coil pulse
Public PulseTransCount    As Long '(June 2010 - I Hilburn) - Enables variable number of
                                    'field calibration values for IRM transverse coil pulse
                                    
Public IRMTransCalDone    As Boolean 'Status flag indicated whether or not the IRM transverse calibration
                                     'has been done
Public IRMAxialCalDone       As Boolean 'Ditto, but for the IRM axial calibration
Public IRMAxialVoltMax       As Double  'Max IRM Box input voltage to allow for the axial IRM Pulse
Public IRMTransVoltMax       As Double  'ditto, but for the transverse IRM pulse
Public AxialTransMaxCapVoltsSame As Boolean     'Setting to indicate if the Transverse max pulse
                                                'voltage should be slaved (set) to the Axial max
                                                'pulse value.
Public PulseAxialMax         As Double
Public PulseAxialMin         As Double
Public PulseTransMax         As Double
Public PulseTransMin         As Double
Public AscIrmMaxFireAtZeroGaussReadVoltage As Double

' now susceptibility variables
Public SusceptibilityMomentFactorCGS As Double
Public SusceptibilityScaleFactor     As Double
Public SusceptibilitySettings As String

' MCC volt conversion converts a charging volt
' to a MCC output voltage (default: 10 V MCC -> 450 V)
Public PulseMCCVoltConversion       As Double
Public PulseReturnMCCVoltConversion As Double

'System DAQ Boards and Channels Collection Declaration
Public SystemBoards As Boards
Public SystemAssignedChannels As Channels

'Local Copy of the System DAQ Boards and Channels collection for saving temporary
'changes made by the user prior to applying them to the System collections that the
'Paleomag program uses
Public LocalBoards As Boards
Public LocalAssignedChannels As Channels

'Setting for the DAQ Boards settings form - Defaults to true when the code is first started
Public LocalAndSystemDifferent As Boolean

'(March 2010 - I Hilburn)
'Status Flag indicating whether the .ini file board settings
'have been imported into the System Boards global collection
Public ImportBoardsDone As Boolean

'(April 2010 - I Hilburn)
'Status flag indicating that the .ini file contains no board settings
'upon start of the Paleomag program.  This will happen if the .ini file
'has accidentally been erased, or the .ini file has not been updated
'to the new Boards/Channels/WaveForms format.
Public NoINIBoards As Boolean

'(March 2010 - I Hilburn)
'Same thing, but for the import of assigned channels data
Public ChannelsImportDone As Boolean

'(April 2010 - I Hilburn)
'Status flag indicating that the .ini file contains no channel settings
'upon the start of the Paleomag program.
Public NoINIChannels As Boolean

'(March 2010 - I Hilburn)
'System Wave Forms Collection Declaration
Public WaveForms As Waves

'(July 2012 - T Shuma)
'XY Table Positions Collection
Public XYTablePositions(101, 1) As Long
Public HasXYTableBeenHomed As Boolean
'Used for displaying current cup# on measurement screen
Public SampleHandlerCurrentHole As Double

'(April 2010 - I Hilburn)
'Indicator the all the wave forms stored in the .ini file have been imported
Public ImportWavesDone As Boolean

'(April 2010 - I Hilburn)
'Status flag indicating that the .ini file contains no Wave Form settings
'upon the start of the Paleomag program.
Public NoINIWaveForms As Boolean

'(July 2012 - T Shuma)
'Status flag indicating that the .ini file contains no XY Table Positions settings
'upon the start of the Paleomag program.
Public NoINIXYTablePositions As Boolean

' Now the ARM calibration variables
Public ARMMax           As Double
Public ARMVoltGauss     As Double
Public ARMVoltMax       As Double
Public ARMTimeMax       As Double
Public DoVacuumReset    As Boolean
Public DoDegausserCooling As Boolean
Public DropoffVacuumDelay As Double

' (March 2008 - L Carporzen) Put in Settings the IRM/ARM channels (MIT acquisition board does not work on IRMTrim = 3
' Analog channel output
'(March 2010 - I Hilburn) Changed Integer channel/port numbs to Channel objects
Public ARMVoltageOut As Channel
Public IRMVoltageOut  As Channel

' Analog input
'(March 2010 - I Hilburn) Changed Integer chan/port number to Channel object
Public IRMCapacitorVoltageIn  As Channel
Public IRMMonitor As Channel

'Analog MCC Input Channels #'s for Temperature sensors on AF coils
'(March 2010 - L Carporzen)
'(March 2010 - I Hilburn) Changed Integer channel/port numbs to Channel objects
Public AnalogT1 As Channel
Public AnalogT2 As Channel

' DIO line assignments
'(March 2010 - I Hilburn) Changed Integer channel/port numbs to Channel objects
Public ARMSet  As Channel
Public IRMFire  As Channel
Public IRMTrim  As Channel
Public IRMPowerAmpVoltageIn  As Channel
Public MotorToggle As Channel
Public VacuumToggleA As Channel
Public VacuumToggleB As Channel
Public DegausserToggle As Channel
Public AFAxialRelay As Channel
Public AFTransRelay As Channel
Public IRMRelay As Channel

'Old Style DIO Line assignments (channel nums stored as integer)
'(March 2011 - I Hilburn)
Public ARMSetNo As Integer
Public IRMFireNo As Integer
Public IRMTrimNo As Integer
Public MotorToggleNo As Integer
Public VacTogANo As Integer
Public VacTogBNo As Integer
Public DegCoolNo As Integer
Public IRMCapVInNo As Integer
Public ARMVOutNo As Integer
Public IRMVOutNo As Integer

' Now assign the COMM Ports to the proper lines!
Public COMPortSquids    As Integer
Public COMPortAf        As Integer
Public COMPortUpDown    As Integer
Public COMPortTurning   As Integer
Public COMPortChanger   As Integer
Public COMPortChangerY  As Integer
Public COMPortVacuum    As Integer
Public COMPortSusceptibility  As Integer

'VB Send Mail settings
Public MailSMTPHost           As String
Public MailSMTPPort           As Integer
Public MailSMTPPassword    As String
Public MailSMTPUsername    As String
Public MailSMTPAuthenticate   As MailSmtpAuthenticateEnum
Public MailSMTPSendUsing      As MailSMTPSendUsingEnum
Public MailUseSSLEncryption   As Boolean
Public MailFrom               As String
Public MailFromName           As String
Public MailFromPassword       As String
Public MailCCList             As String
Public MailStatusMonitor      As String

Public Enum MailSmtpAuthenticateEnum
    cdoAnonymous = 0
    cdoBasic = 1
    cdoNTLM = 2
End Enum

Public Enum MailSMTPSendUsingEnum
    localSmtp = 1
    remoteSmtpHost = 2
End Enum


'Module Settings
Public EnableAxialIRM         As Boolean
Public EnableTransIRM         As Boolean
Public EnableARM              As Boolean
Public EnableAF               As Boolean
Public EnableAltAFMonitor     As Boolean

'Use Temperature Sensors - boolean status flag
'(Mar, 2010 - L Carporzen)
Public EnableT1               As Boolean
Public EnableT2               As Boolean

'Other Modules user can switch on and off
Public EnableVacuum As Boolean
Public EnableDegausserCooler As Boolean
Public EnableSusceptibility   As Boolean
Public EnableIRMBackfield     As Boolean
Public EnableIRMMonitor       As Boolean    '(July 2010 - I Hilburn - IRM monitor not yet implemented - hardware setup still being
                                            ' developed to do this)
Public EnableAFAnalysis       As Boolean    '(July 2010 - I Hilburn - This turns on a fancy display box that shows the
                                            'characteristics of the AF ramp waveform after each ramp cycle
                                            
Public Enum SaveCoilParam

    resFreq = 0
    VoltsMax = 1

End Enum

'Need a couple of Module-wide variables to deal with tracking errors during the Channel settings import
'from the .INI file
Dim ChanImportError As Boolean
Dim ChanImportErrorCause As String
Dim DisabledModules As String
Dim isAFMonitorMissing As Boolean
Dim isAFRampUpMissing As Boolean
Dim isAFRampDownMissing As Boolean
Dim isAltAFMonitorMissing As Boolean
Dim isIRMMonitorMissing As Boolean

' Private function Add_INIBoard
'
' Created: April 2-3, 2010
'  Author: Isaac Hilburn
'
' Summary: This function appends a new Board at the end of the [Boards] section
'          of the .ini file.  Or, if given permission by the user, overwrites
'          an existing Board in the [Boards] section.
'
'   Input:
'
'   newBoard    -   Referenced Board object, containing all the information to
'                   be saved to the .ini file
'
'   doOverwrite -   Boolean.  Optional argument.  Default value = False.
'                   If set to false, then the code will only add new boards
'                   from System Boards to the .ini that aren't contained in the file already.
'                   If set to True, and a Board with a matching BoardININum, BoardNum, and/or
'                   BoardName exists in the .ini file, then the .Ini version of the Board
'                   will be overwritten by the information in the newBoard object passed
'                   in above
'
'  Output:
'
'       Boolean -   True = add board was successful,
'                   False = Board already exists & overwrite was set to false
'
'
Public Function Add_INIBoard _
    (ByRef NewBoard As Board, _
     Optional ByVal doOverwrite As Boolean = False) As Boolean

    Dim i, j, k As Long
    Dim IniBoardsCount As Long
    Dim TempStr As String
    Dim BoardFound As Boolean
    Dim TempChannels As Channels
    Dim DIOString As String
    Dim BoardName As String
    Dim BoardININum As Long
    
    'Get the current # of boards in the .ini file
    IniBoardsCount = val(Config_GetFromINI("Boards", _
                                           "BoardsCount", _
                                           "-1", _
                                           Prog_INIFile))
                                           
    If IniBoardsCount = -1 Then
    
        'Need to create the Boards section and the IniBoardsCount
        Config_SaveSetting "Boards", _
                           "BoardsCount", _
                           "1"
                           
        'There are no boards loaded in the INI file yet, so:
        BoardFound = False
                           
    Else
    
        'Need to search to see if a board with a matching .ini file value
        'already exists
        TempStr = Config_GetFromINI("Boards", _
                                    "BoardNum" & Format(NewBoard.BoardININum, "0"), _
                                    "-1", _
                                    Prog_INIFile)
                                    
        If TempStr = "-1" Then
        
            BoardFound = False
            
        Else
        
            BoardFound = True
            
        End If
        
    End If
    
    If BoardFound = True And doOverwrite = False Then
    
        'Can't save this board, it already exists and the
        'user hasn't given permission to overwrite the version of the board
        'stored in the .INI file.
        Add_INIBoard = False
        
        Exit Function
        
    End If
        
    
    'Otherwise, go and add it
    With NewBoard
        
        'Copy the BoardName and the BoardININum to local variables
        BoardName = .BoardName
        BoardININum = .BoardININum
        
        Config_SaveSetting "Boards", _
                           "BoardININum" & Format(.BoardININum, "0"), _
                           Trim(str(.BoardININum))
            
        
            
        Config_SaveSetting "Boards", _
                           "BoardNum" & Format(.BoardININum, "0"), _
                           Trim(str(.BoardNum))
                                
        'Store the Board # to a local long variable
        BoardNum = .BoardNum
                                
        Config_SaveSetting "Boards", _
                           "BoardName" & Format(.BoardININum, "0"), _
                           Trim(.BoardName)
                                
        Config_SaveSetting "Boards", _
                           "BoardFunction" & Format(.BoardININum, "0"), _
                           Trim(.BoardFunction)
                                
        Config_SaveSetting "Boards", _
                           "CommProtocol" & Format(.BoardININum, "0"), _
                           Trim(str(.CommProtocol))
                        
        Config_SaveSetting "Boards", _
                           "BoardMode" & Format(.BoardININum, "0"), _
                           Trim(str(.BoardMode))
                        
        Config_SaveSetting "Boards", _
                           "MaxAInRate" & Format(.BoardININum, "0"), _
                           Trim(str(.MaxAInRate))
                        
        Config_SaveSetting "Boards", _
                           "MaxAOutRate" & Format(.BoardININum, "0"), _
                           Trim(str(.MaxAOutRate))
                        
        'Now we have a choice - depending on the comm protocol
        '(i.e. the type / manufacture of the board) we can either
        'save a RangeType value (for Measurement Computing Boards)
        'or RangeMax and RangeMin values (for ADWIN and other board types)
        If .CommProtocol = MCC_UL Then
        
            'Save a RangeType value
            Config_SaveSetting "Boards", _
                               "RangeType" & Format(.BoardININum, "0"), _
                               Trim(str(.range.RangeType))
                            
        Else
        
            'Save RangeMax and RangeMin values
            Config_SaveSetting "Boards", _
                               "RangeMax" & Format(.BoardININum, "0"), _
                               Trim(str(.range.MaxValue))
                            
            Config_SaveSetting "Boards", _
                               "RangeMin" & Format(.BoardININum, "0"), _
                               Trim(str(.range.MinValue))
                            
        End If
        
        'Now need to go through the four channel collections:
        '   1) Analog In
        '   2) Analog Out
        '   3) Digital In
        '   4) Digital Out
        For k = 1 To 4
        
            'As default, set TempChannels = Nothing
            Set TempChannels = Nothing
        
            'Now assign the TempChannels object to the correct
            'channels object on the board, if this object doesn't
            'exist for a given board, then skip to the next
            'iteration of the for loop
            If k = 1 And Not .AInChannels Is Nothing Then
                    
                Set TempChannels = .AInChannels
                
                'Save to INI file
                Config_SaveSetting "Boards", _
                                   "AInChannelsCount" & Format(.BoardININum, "0"), _
                                   TempChannels.Count
                                
            ElseIf k = 2 And Not .AOutChannels Is Nothing Then
            
                Set TempChannels = .AOutChannels
                
                'Save to INI file
                Config_SaveSetting "Boards", _
                                   "AOutChannelsCount" & Format(.BoardININum, "0"), _
                                   TempChannels.Count
                
            ElseIf k = 3 And Not .DInChannels Is Nothing Then
            
                Set TempChannels = .DInChannels
                
                'Write in INI section for whether the Digital I/O is configured
                'this field always comes before the DIChannelsCount field
                If .DIOConfigured = True Then
                    
                    DIOString = "True"
                    
                Else
                
                    DIOString = "False"
                
                End If
                
                Config_SaveSetting "Boards", _
                                   "DIOConfigured" & Format(.BoardININum, "0"), _
                                   DIOString
                
                'Save to INI file
                Config_SaveSetting "Boards", _
                                   "DInChannelsCount" & Format(.BoardININum, "0"), _
                                   TempChannels.Count
                
            ElseIf k = 4 And Not .DOutChannels Is Nothing Then
            
                Set TempChannels = .DOutChannels
                
                'Save to INI file
                Config_SaveSetting "Boards", _
                                   "DOutChannelsCount" & Format(.BoardININum, "0"), _
                                   TempChannels.Count
                
            End If

            'Check to see if TempChannels was successfully set
            'to a Channels object
            If Not TempChannels Is Nothing Then
            
                With TempChannels
            
                    If .Count > 0 Then
                    
                        'If there is more than one channel in this collection
                        'then go through for loop to assign the channel values
                
                        For j = 1 To .Count
                        
                            With .Item(j)
                                
                                'This is where the ChanTypeStr & BoardNum
                                'local variables comes in handy
                                
                                'Check to see if the Board Number attached
                                'to the channel jives with the board
                                'this channel is supposed to belong to
                                If .BoardName <> BoardName Or _
                                   .BoardININum <> BoardININum _
                                Then
                                
                                    'In case of disagreement,
                                    'reset the channel's board num
                                    'to that of the parent board
                                    .BoardName = BoardName
                                    .BoardININum = BoardININum
                                    
                                End If
                                
                                Config_SaveSetting _
                                    "Boards", _
                                    Trim(.ChanType) & "-" & Trim(str(.BoardININum)) & _
                                    "-CH" & Format(j - 1, "0"), _
                                    Trim(.ChanName) & "," & _
                                    Trim(str(.ChanNum))
                                    
                            End With
                            
                        Next j
                        
                    End If
                    
                End With
                
            End If
                          
        Next k
        
    End With

    'Set successful return value
    Add_INIBoard = True

End Function

' Private function Add_INIWaveForm
'
' Created: April 2-3, 2010
'  Author: Isaac Hilburn
'
' Summary: This function appends a new WaveForm at the end of the [WaveForms] section
'          of the .ini file.  Or, if given permission by the user, overwrites
'          an existing Wave Form in the [WaveForms] section.
'
'   Input:
'
'   newWave     -   Referenced Wave object, containing all the information to
'                   be saved to the .ini file
'
'   doOverwrite -   Boolean.  Optional argument.  Default value = False.
'                   If set to false, then the code will only add new Wave objects
'                   from WaveForms to the .ini that aren't contained in the file already.
'                   If set to True, and a Wave with a matching WaveININum exists in the
'                   .ini file, then the .Ini version of the Wave will be overwritten
'                   by the information in the newWave object passed in above
'
'  Output:
'
'       Boolean -   True = add Wave Form was successful,
'                   False = Wave form already exists & overwrite was set to false
'
Private Function Add_IniWaveForm _
    (ByRef NewWave As Wave, _
     Optional ByVal doOverwrite As Boolean = False) As Boolean


    Dim i, j, k As Long
    Dim IniWavesCount As Long
    Dim TempStr As String
    Dim WaveFound As Boolean
        
    'Get the current # of boards in the .ini file
    IniWavesCount = val(Config_GetFromINI("WaveForms", _
                                           "WaveFormCount", _
                                           "-1", _
                                           Prog_INIFile))
                                           
    If IniWavesCount = -1 Then
    
        'Need to create the Boards section and the IniBoardsCount
        Config_SaveSetting "Boards", _
                           "BoardsCount", _
                           "1"
                           
        'There are no Waves loaded in the INI file yet, so:
        WaveFound = False
                           
    Else
    
        'Need to search to see if a Wave with a matching .ini file value
        'already exists
        TempStr = Config_GetFromINI("WaveForms", _
                                    "WaveININum" & Format(NewWave.WaveININum, "0"), _
                                    "-1", _
                                    Prog_INIFile)
                                    
        If TempStr = "-1" Then
        
            WaveFound = False
            
        Else
        
            WaveFound = True
            
        End If
        
    End If
    
    If WaveFound = True And doOverwrite = False Then
    
        'Can't save this Wave, it already exists and the
        'user hasn't given permission to overwrite the version of the Wave
        'stored in the .INI file.
        Add_IniWaveForm = False
        
        Exit Function
        
    End If
    
    
    'Otherwise, go and add / overwrite the WaveForm
    With NewWave
    
        'Save the Wave's .INI # key
        Config_SaveSetting "WaveForms", _
                           "WaveININum" & Format(.WaveININum, "0"), _
                           Trim(str(.WaveININum))
                           
        
        'Save the Wave's Board String - containing the Board Name
        'and the Board's device number (NOT! the board's INI number)
        Config_SaveSetting "WaveForms", _
                           "BoardUsed" & Format(.WaveININum, "0"), _
                           Trim(.BoardUsed.BoardName) & "," & _
                           Trim(str(.BoardUsed.BoardNum))
                           
        'Create INI Channel String from board + channel objects
        TempStr = CreateINIChannelStr(.Chan, _
                                      .BoardUsed)
                           
        'Save the Wave's Channel string - containg the Channel type string
        Config_SaveSetting "WaveForms", _
                           "Chan" & Format(.WaveININum, "0"), _
                           TempStr
                                              
        'Save the WaveName, a string value describing the waves purpose in
        'plain english to use for error messages
        Config_SaveSetting "WaveForms", _
                           "WaveName" & Format(.WaveININum, "0"), _
                           Trim(.WaveName)
                           
                                         
        'Save the start point to begin collecting or outputing data from
        'in the wave
        Config_SaveSetting "WaveForms", _
                           "StartPoint" & Format(.WaveININum, "0"), _
                           Trim(str(.StartPoint))
                           
                                             
        'Save the status of whether memory space - as a global object or
        'a windows memory buffer has been allocated for the wave-form
        If .BufferAlloc = True Then
            TempStr = "True"
        Else
            TempStr = "False"
        End If
        
        Config_SaveSetting "WaveForms", _
                           "MemAlloc" & Format(.WaveININum, "0"), _
                           TempStr
        
        'Save the status flag that indicates if the memory space for the wave-form
        'should be emptied after the wave-form has finished being used by
        'the Paleomag code
        If .DoDeallocate = True Then
            TempStr = "True"
        Else
            TempStr = "False"
        End If
        Config_SaveSetting "WaveForms", _
                           "DoDeallocate" & Format(.WaveININum, "0"), _
                           TempStr
                                                
                                                
        'Save IO type (input or output)
        Config_SaveSetting "WaveForms", _
                           "IO" & Format(.WaveININum, "0"), _
                           Trim(.IO)
                                                
        'Save the IORate at which the wave should be output/input
        Config_SaveSetting "WaveForms", _
                           "IORate" & Format(.WaveININum, "0"), _
                           Trim(str(.IORate))
                           
            
            
        'Check the comm protocol of the board associated with the Wave
        'if the comm protocol is Measurement Computing, then
        'save a RangeType value, else if it's ADWIN or other, save
        'the Max and Min values of the Board's range
        If .BoardUsed.CommProtocol = MCC_UL Then
                        
            'Save the RangeType (Measurement Computing boards, only) for the
            'output/input of the wave
            Config_SaveSetting "WaveForms", _
                               "RangeType" & Format(.WaveININum, "0"), _
                               Trim(str(.range.RangeType))
                                                     
        Else
       
            'This is an ADWIN or OTHER board type - save
            'the Max and Min Range values
                
            Config_SaveSetting "WaveForms", _
                               "RangeMax" & Format(.WaveININum, "0"), _
                               Trim(str(.range.MaxValue))
                               
            Config_SaveSetting "WaveForms", _
                               "RangeMin" & Format(.WaveININum, "0"), _
                               Trim(str(.range.MinValue))
        End If
            
            
        'Save the slope to use in ramping up an output wave
        'this value may be changed later by the code during run-time
        Config_SaveSetting "WaveForms", _
                           "Slope" & Format(.WaveININum, "0"), _
                           Trim(str(.Slope))
                                   
    End With

    'Indicate that the function was successful!
    Add_IniWaveForm = True

End Function

'Private sub AddDisabledModule(String)
'
' Created: August 28, 2010
'  Author: Isaac Hilburn
'
'  Summary: Accepts in a simple string descriptor of the modules or module that
'           needs to be disabled.  The function ensures that the corresponding modules / module
'           are disabled and have been added to the disabled modules string
'
Private Sub AddDisabledModule(ByVal ModName As String)

    Dim TempL As Long
    Dim ModuleStr As String
    Dim ModuleStr2 As String
    
    ModuleStr = vbNullString
    ModuleStr2 = vbNullString
    
    Select Case ModName
    
        Case "AF"
        
            'If current system is "ADWIN" then
            'disable all the AF modules if not already done
            If AFSystem = "ADWIN" Then
            
                EnableAF = False
                EnableAFAnalysis = False
                
            End If
            
            ModuleStr = "All AF Modules"
            
        Case "IRM"
        
            EnableAxialIRM = False
            EnableTransIRM = False
            EnableIRMBackfield = False
            EnableIRMMonitor = False
            
            ModuleStr = "All IRM Modules"
            
        Case "IRMAxial"
        
            EnableAxialIRM = False
            
            ModuleStr = "Axial IRM"
            ModuleStr2 = "All IRM Modules"
            
        Case "IRMTrans"
        
            EnableTransIRM = False
            
            ModuleStr = "Trans IRM"
            ModuleStr2 = "All IRM Modules"
            
        Case "IRMReturn"
        
            EnableTransIRM = False
            EnableAxialIRM = False
            
            ModuleStr = "All IRM Modules"
            
        Case "IRMMonitor"
        
            EnableIRMMonitor = False
            
            ModuleStr = "IRM Monitor"
            ModuleStr2 = "All IRM Modules"
            
        Case "ARM"
        
            EnableARM = False
            
            ModuleStr = "ARM"
            
        Case "AnalogT1"
        
            EnableT1 = False
            
            ModuleStr = "Temp. Sensor #1"
            
        Case "AnalogT2"
        
            EnableT2 = False
            
            ModuleStr = "Temp. Sensor #2"
            
        Case "AltAFMonitor"
        
            EnableAltAFMonitor = False
            
            ModuleStr = "Alternate AF Monitor"
            
        Case "Vacuum"
        
            ModuleStr = "Vacuum System"
            
        Case "Degausser Cooler"
        
            ModuleStr = "Degausser Cooler System"
    End Select
    
    'If the first module string is null, then there's nothing to
    'add to the disabled modules string
    If ModuleStr = vbNullString Then Exit Sub
            
    'Check to see if disabled modules string is null
    If DisabledModules = vbNullString Then
    
        DisabledModules = ModuleStr & ", "
        
    Else
    
        'Need to check to see if the module str is already in disabled modules string
        If InStr(1, DisabledModules, ModuleStr) = 0 Then
        
            'Check to see if the second module str = null
            If ModuleStr2 <> vbNullString Then
            
                'Need to check the second string as well
                If InStr(1, DisabledModules, ModuleStr2) = 0 Then
                
                    'Now can add module string to disabled modules
                    DisabledModules = DisabledModules & _
                                      ModuleStr & ", "
                                      
                End If
                
            Else
            
                'There is no second string
                DisabledModules = DisabledModules & _
                                  ModuleStr & ", "
                                  
            End If
            
        End If
        
    End If
        
End Sub

Public Sub AddToChanColNoKeyRepeat(ByRef ChanColl As Channels, _
                                   ByRef NewChan As Channel, _
                                   ByRef NewKey As String)
    
    Dim N As Long
    Dim M As Long
    Dim i As Long
    Dim TempStr As String
    Dim AddSuccess As Boolean
     
    'Try to add the key using the Add with error checking
    AddSuccess = ChanColl.AddErrorCheck(NewChan, NewKey)
         
    If AddSuccess = False Then
    
        'There was a key repeat
        'Need to access current channel with this key
        'and add to the channel description
        With ChanColl(NewKey)
        
            'Get the size of the channel description array
            'for the channel already in the collection
            N = .ChanDescs.Count
            
            'Get the Size of the channel description array
            'for the new channel
            M = NewChan.ChanDescs.Count
            
            'Loop to add new channel descriptions to the channel already
            'in the collection
            For i = N + 1 To M + N
            
                'Get the new description to add
                TempStr = NewChan.ChanDescs.ChanDesc(i - N)
                
                'Save it to the channel in the collection
                .ChanDescs.AddDesc TempStr
                
            Next i
            
        End With
        
    End If
        
End Sub

'Sub Allocate_DefaultINI()
'
' Created: February 22,2010
'  Author: Isaac Hilburn
'
' Summary: Allocates the global DefaultINI object as a new CINIFile object with the filename pointing to
'          the path "../Paleomag2010-2.4.0/Paleomag 2010/Defaults.INI where the pathway to the Paleomag.vbp project file is:
'          "../Paleomag2010-2.4.0/Paleomag 2010/Paleomag.vbp"
'
' WARNING!!!!
'           This code assumes that no call has been made to alter the original run value of app.path, i.e. it assumes
'           that app.path still points to the directory that contains the Paleomag.vbp file.
'           If app.path has been shifted to some other value, this function will cause an error to occur later on
'           when the code tries to access the Defaults.INI file
Public Sub Allocate_DefaultINI()

    Dim TempStr As String
    
    'Allocate the Default INI file object
    Set DefaultINI = New CIniFile

    'Generate the path-string to the Defaults.INI file
    If Right(App.path, 1) = "\" Then
        
        TempStr = App.path & "Defaults.INI"
        
    Else
    
        TempStr = App.path & "\" & "Defaults.INI"
        
    End If

    'Set the filename
    DefaultINI.filename = TempStr

End Sub

'--------------------------------------------------------------------------------------------------'
'--------------------------------------------------------------------------------------------------'
'
'   Modified: April 4, 2010
'     Author: I Hilburn
'
'    Summary: Function modified to use the CIniFile class library
'                      to implement all the .ini file read/write/delete actions
'                      See CIniFile.cls for more documentation on the properties and
'                      methods in this class.
'
'     Output: String value
'
'    WARNING!!!: Old return value used to be a boolean.
'
'                All functions calling
'                this function have been updated with this change!
'                If they are changed back, this will break the Paleomag code!
'
Public Function Config_AddToINI _
    (sSection As String, _
     sKey As String, _
     sValue As String, _
     sIniFile As String) As String  '(April 2010, I Hilburn) - This used be return a Boolean

    Config_AddToINI = IniFile.EntryWrite(sKey, sValue, sSection)
    
'--------------------------------------------------------------------------------------------------'
'--------------------------------------------------------------------------------------------------'
'
'   Older version of code
'
'    '// Returns True if successful. If section does not
'    '// exist it creates it.
'
'    Dim lRet As Long
'    ' Call DLL
'    lRet = WritePrivateProfileString(sSection, sKey, sValue, sIniFile)
'    Config_AddToINI = (lRet)
'
'--------------------------------------------------------------------------------------------------'
'--------------------------------------------------------------------------------------------------'

End Function

'--------------------------------------------------------------------------------------------------'
'--------------------------------------------------------------------------------------------------'
'
'   Function modified: April 4, 2010
'              Author: I Hilburn
'
'             Summary: Function modified to use the CIniFile class library
'                      to implement all the .ini file read/write/delete actions
'                      See CIniFile.cls for more documentation on the properties and
'                      methods in this class.
'
Public Function Config_GetFromINI(sSection As String, sKey As String, sDefault As String, sIniFile As String)

    Config_GetFromINI = IniFile.EntryRead(sKey, sDefault, sSection)
    
'--------------------------------------------------------------------------------------------------'
'--------------------------------------------------------------------------------------------------'
'
'   Older version of the code to read a string key and associated value from the Paleomag.ini
'   file.  This has been replaced by the code above, using the CIniFile class libary.
'
'   '// VB Web Code Example
'   '// www.vbweb.co.uk
'   '// Functions
'
'    Dim sBuffer As String, lRet As Long
'    ' Fill String with 255 spaces
'    sBuffer = String$(255, 0)
'    ' Call DLL
'    lRet = GetPrivateProfileString(sSection, sKey, vbNullString, sBuffer, Len(sBuffer), sIniFile)
'    If lRet = 0 Then
'        ' DLL failed, save default
'        If LenB(sDefault) <> 0 Then Config_AddToINI sSection, sKey, sDefault, sIniFile
'        Config_GetFromINI = sDefault
'    Else
'        ' DLL successful
'        ' return string
'        Config_GetFromINI = Left(sBuffer, InStr(sBuffer, Chr(0)) - 1)
'    End If
'
'--------------------------------------------------------------------------------------------------'
'--------------------------------------------------------------------------------------------------'
End Function

Public Function Config_GetSetting(sSection As String, sKey As String, sDefault As String) As String
    Config_GetSetting = Config_GetFromINI(sSection, sKey, sDefault, Prog_INIFile)
End Function

Public Sub Config_ReadINISettings()
    ' This procedure reads settings from a file and adjusts
    ' variables in memory accordingly. (c:\paleomag\paleomag.ini)
    Dim i As Integer
    
    'Paleomag Program settings
    Prog_UsageFile = Config_GetFromINI("Program", "UsageFile", "C:\Paleomag\PALEOUSE.DAT", Prog_INIFile)
    Prog_DefaultBackup = Config_GetFromINI("Program", "DefaultBackupDrive", "E", Prog_INIFile)
    Prog_DefaultPath = Config_GetFromINI("Program", "DefaultPath", "C:\USER", Prog_INIFile)
    Prog_HelpURLRoot = Config_GetFromINI("Program", "HelpURLRoot", "file:///C:/Paleomag/Paleomag%202007/Help/", Prog_INIFile) ' (August 2007 L Carporzen)
    Prog_LogoFile = Config_GetFromINI("Program", "LogoFile", "", Prog_INIFile)
    Prog_IcoFile = Config_GetFromINI("Program", "IcoFile", "", Prog_INIFile) ' (October 2007 L Carporzen)
    Prog_TextEditor = Config_GetFromINI("Program", "TextEditor", "notepad.exe", Prog_INIFile)
    DEBUG_MODE = ("True" = Config_GetFromINI("Program", "DebugMode", "False", Prog_INIFile))
    NOCOMM_MODE = ("True" = Config_GetFromINI("Program", "NoCommMode", "False", Prog_INIFile))
    DumpRawDataStats = ("True" = Config_GetFromINI("Program", "DumpRawDataStats", "False", Prog_INIFile))
    LogMessages = ("True" = Config_GetFromINI("Program", "LogMessages", "False", Prog_INIFile))
    LogFolderPath = Trim(Config_GetFromINI("Program", "LogFolderPath", "", Prog_INIFile))
    LogFileName = Trim(Config_GetFromINI("Program", "LogFileName", "", Prog_INIFile))
    
    'Sample Changer Position + DC Motor Settings
    SlotMin = Int(val(Config_GetFromINI("SampleChanger", "SlotMin", "1", Prog_INIFile)))
    SlotMax = Int(val(Config_GetFromINI("SampleChanger", "SlotMax", "200", Prog_INIFile)))
    HoleSlotNum = Int(val(Config_GetFromINI("SampleChanger", "HoleSlotNum", "10", Prog_INIFile)))
    OneStep = val(Config_GetFromINI("SampleChanger", "OneStep", "-1010.1010101", Prog_INIFile))
    ZeroPos = val(Config_GetFromINI("SteppingMotor", "ZeroPos", "-25886", Prog_INIFile))
    MeasPos = val(Config_GetFromINI("SteppingMotor", "MeasPos", "-30607", Prog_INIFile))
    AFPos = val(Config_GetFromINI("SteppingMotor", "AFPos", "-8405", Prog_INIFile))
    IRMPos = val(Config_GetFromINI("SteppingMotor", "IRMPos", "-8405", Prog_INIFile))
    SCoilPos = val(Config_GetFromINI("SteppingMotor", "SCoilPos", "-4202", Prog_INIFile))
    FloorPos = val(Config_GetFromINI("SteppingMotor", "FloorPos", "-43000", Prog_INIFile))
    MinUpDownPos = val(Config_GetFromINI("SteppingMotor", "MinUpDownPos", "-30500", Prog_INIFile))
    SampleBottom = val(Config_GetFromINI("SteppingMotor", "SampleBottom", "-2385", Prog_INIFile))
    SampleTop = val(Config_GetFromINI("SteppingMotor", "SampleTop", "-1979", Prog_INIFile))
    SampleHeight = SampleTop - SampleBottom
    SampleHoleAlignmentOffset = val(Config_GetFromINI("SteppingMotor", "SampleHoleAlignmentOffset", "-0.02", Prog_INIFile))
    LiftSpeedSlow = val(Config_GetFromINI("SteppingMotor", "LiftSpeedSlow", "4000000", Prog_INIFile))
    LiftSpeedNormal = val(Config_GetFromINI("SteppingMotor", "LiftSpeedNormal", "25000000", Prog_INIFile))
    LiftSpeedFast = val(Config_GetFromINI("SteppingMotor", "LiftSpeedFast", "50000000", Prog_INIFile))
    LiftAcceleration = val(Config_GetFromINI("SteppingMotor", "LiftAcceleration", "90000", Prog_INIFile))
    ChangerSpeed = val(Config_GetFromINI("SteppingMotor", "ChangerSpeed", "31000", Prog_INIFile))
    TurnerSpeed = val(Config_GetFromINI("SteppingMotor", "TurnerSpeed", "2000000", Prog_INIFile))
    SCurveFactor = val(Config_GetFromINI("SteppingMotor", "SCurveFactor", "32767", Prog_INIFile))
    TurningMotorFullRotation = val(Config_GetFromINI("SteppingMotor", "TurningMotorFullRotation", "2000", Prog_INIFile))
    TurningMotor1rps = val(Config_GetFromINI("SteppingMotor", "TurningMotor1rps", "16000000", Prog_INIFile))
    UpDownMotor1cm = val(Config_GetFromINI("SteppingMotor", "UpDownMotor1cm", "10", Prog_INIFile))
    UpDownTorqueFactor = Int(val(Config_GetFromINI("SteppingMotor", "UpDownTorqueFactor", "40", Prog_INIFile)))
    UpDownMaxTorque = Int(val(Config_GetFromINI("SteppingMotor", "UpDownMaxTorque", "32000", Prog_INIFile)))
    PickupTorqueThrottle = val(Config_GetFromINI("SteppingMotor", "PickupTorqueThrottle", "0.4", Prog_INIFile))
    TrayOffsetAngle = val(Config_GetFromINI("SteppingMotor", "TrayOffsetAngle", "0", Prog_INIFile))
    CmdHomeToTop = val(Config_GetFromINI("MotorPrograms", "CmdHomeToTop", "206", Prog_INIFile))
    CmdSamplePickup = val(Config_GetFromINI("MotorPrograms", "CmdSamplePickup", "241", Prog_INIFile))
    MotorIDTurning = val(Config_GetFromINI("MotorPrograms", "MotorIDTurning", "16", Prog_INIFile))
    MotorIDChanger = val(Config_GetFromINI("MotorPrograms", "MotorIDChanger", "16", Prog_INIFile))
    MotorIDChangerY = val(Config_GetFromINI("MotorPrograms", "MotorIDChangerY", "16", Prog_INIFile))
    MotorIDUpDown = val(Config_GetFromINI("MotorPrograms", "MotorIDUpDown", "16", Prog_INIFile))
    
                                             
    ' Now Magnetometer Calibration Constants
    ZCal = val(Config_GetFromINI("MagnetometerCalibration", "ZCal", "-2.516", Prog_INIFile))
    XCal = val(Config_GetFromINI("MagnetometerCalibration", "XCal", "-3.410", Prog_INIFile))
    YCal = val(Config_GetFromINI("MagnetometerCalibration", "YCal", "-3.470", Prog_INIFile))
    RangeFact = val(Config_GetFromINI("MagnetometerCalibration", "RangeFact", "0.00001", Prog_INIFile))
    ReadDelay = val(Config_GetFromINI("MagnetometerCalibration", "ReadDelay", "1", Prog_INIFile)) ' (March 2008 L Carporzen) Read delay
    RemeasureCSDThreshold = val(Config_GetFromINI("Magnetometry", "RemeasureCSDThreshold", "8", Prog_INIFile))
    
    ' New selections in the Options menu (April-May 2007 L Carporzen)
    JumpThreshold = val(Config_GetFromINI("Magnetometry", "JumpThreshold", "0.1", Prog_INIFile))
    StrongMom = val(Config_GetFromINI("Magnetometry", "StrongMom", "0.02", Prog_INIFile))
    IntermMom = val(Config_GetFromINI("Magnetometry", "IntermMom", "0.000001", Prog_INIFile))
    MomMinForRedo = val(Config_GetFromINI("Magnetometry", "MomMinForRedo", "0.000000008", Prog_INIFile))
    JumpSensitivity = val(Config_GetFromINI("Magnetometry", "JumpSensitivity", "1", Prog_INIFile))
    NbTry = val(Config_GetFromINI("Magnetometry", "NbTry", "5", Prog_INIFile))
    NbHolderTry = val(Config_GetFromINI("Magnetometry", "NbHolderTry", "0", Prog_INIFile))
    Meascount = 1
    
    ' now the susceptibility factor settings
    SusceptibilityMomentFactorCGS = val(Config_GetFromINI( _
                                            "SusceptibilityCalibration", _
                                            "SusceptibilityMomentFactorCGS", _
                                            "10", _
                                            Prog_INIFile))
                                            
    SusceptibilityScaleFactor = val(Config_GetFromINI( _
                                        "SusceptibilityCalibration", _
                                        "SusceptibilityScaleFactor", _
                                        "1", _
                                        Prog_INIFile))
    
    'AF System Setting
    '(April 2010, Isaac Hilburn)
    'This setting allows the user to toggle between using the 2G and the ADWIN/DAQ AF system setups
    AFSystem = Trim(Config_GetFromINI("AF", _
                                 "AFSystem", _
                                 "2G", _
                                 Prog_INIFile))
                                 
    'ADWIN Bin Folder Path setting
    '(June 2010, Isaac Hilburn)
    'This setting stores the folder for the ADWIN ADBasic code modules, the ADWIN libraries,
    'and drivers.  This path is needed to do anything with the ADWIN board
    ADWINBinFolderPath = Trim(Config_GetFromINI("AF", _
                                                "ADWINBinFolderPath", _
                                                App.path & "\Adwin\", _
                                                Prog_INIFile))
                                                
    'This setting stores the name of the file in the above folder that
    'should be used to boot and configure the ADWIN board
    ADWINBootFileName = Trim(Config_GetFromINI("AF", _
                                               "ADWINBootFile", _
                                               "ADwin9.btl", _
                                               Prog_INIFile))
                                               
    'This setting stores the name of the file in the above folder that
    'contains the compiled ADBasic code for running an AF ramp cycle on the ADwin Board
    ADWINRampProgFileName = Trim(Config_GetFromINI("AF", _
                                                   "ADWINRampProgFile", _
                                                   "sineout.T91", _
                                                   Prog_INIFile))
                                                   
    'Need to actually set the bin folder path and the Boot File name
    'in the ADWIN code module now
    ADWIN.BinFolderPath = ADWINBinFolderPath
    ADWIN.BootFileName = ADWINBootFileName
    ADWIN.CurProcessFile = ADWINRampProgFileName
    
    'Get AF Data File Save paths & backup settings
    ADWIN_AFDataLocalDir = Trim(Config_GetFromINI("AFFileSave", _
                                                  "ADWINDataFileSaveLocalDir", _
                                                  "C:\Documents and Settings\lab\Desktop\ADWIN Ramp Data", _
                                                  Prog_INIFile))
                                                  
    TWOG_AFDataLocalDir = Trim(Config_GetFromINI("AFFileSave", _
                                                  "2GDataFileSaveLocalDir", _
                                                  "C:\Documents and Settings\lab\Desktop\2G Ramp Data", _
                                                  Prog_INIFile))
                                                  
    ADWIN_AFDataBackupDir = Trim(Config_GetFromINI("AFFileSave", _
                                                  "ADWINDataFileSaveBackupDir", _
                                                  "Y:\Paleomagnetics\ADWIN Ramp Data", _
                                                  Prog_INIFile))
                                               
    TWOG_AFDataBackupDir = Trim(Config_GetFromINI("AFFileSave", _
                                                  "2GDataFileSaveBackupDir", _
                                                  "Y:\Paleomagnetics\2G Ramp Data", _
                                                  Prog_INIFile))
                                                  
    AFDoDataFileBackup = (Trim(Config_GetFromINI("AFFileSave", _
                                                 "AFDataFileSaveDoBackup", _
                                                 "False", _
                                                 Prog_INIFile)) = "True")
                                               
    'IRM System Setting
    '(June 2010, Isaac Hilburn)
    'This setting allows the user to toggle between IRM systems using the old (crappy) power amp
    'or the new (shiny) Matsusada power amp.
    'The two types of power amps have different comm-setups (i.e. the Capacitor return
    'voltage in the Matsusada power amp is just the last IRM charging voltage set value
    'sent from the code to the ASC scientific IRM Box
    IRMSystem = Trim(Config_GetFromINI("IRMPulse", _
                                  "IRMSystem", _
                                  "Old", _
                                  Prog_INIFile))
    
    'AF DAQ Settings
    '(August 2010, Isaac Hilburn)
    'These settings allow the user to control how fast the Ramp Up will be for the ADWIN AF system
    'They set bounds on the fastest allowed ramp up time in ms and the slowest allowed ramp up
    MinRampUpTime_ms = val(CLng(Config_GetFromINI("AF", _
                                                  "MinRampUpTime_ms", _
                                                  "500", _
                                                  Prog_INIFile)))
                                                  
    MaxRampUpTime_ms = val(CLng(Config_GetFromINI("AF", _
                                                  "MaxRampUpTime_ms", _
                                                  "1000", _
                                                  Prog_INIFile)))
                                                  
    AxialRampUpVoltsPerSec = val(Config_GetFromINI("AF", _
                                                   "AxialRampUpVoltsPerSec", _
                                                   "4", _
                                                   Prog_INIFile))
                                               
                                               
    TransRampUpVoltsPerSec = val(Config_GetFromINI("AF", _
                                                   "TransRampUpVoltsPerSec", _
                                                   "4", _
                                                   Prog_INIFile))
                                               
    'AF DAQ Settings
    '(August 2010, Isaac Hilburn)
    'These settings allow the user to control how fast the Ramp Down will be for the ADWIN AF system
    'They set bounds on the fastest allowed ramp down in number of periods for the ramp down
    'The ADWIN ADBasic board code allows the user to specify the number of periods that the Ramp down
    'voltage output should last for, and the ramp down ouput will last for precisely that number of periods
    'Additionally, they allow the user to specify a Periods to Voltage relationship to use that will
    'be truncated by the max and min numperiods values
    MinRampDown_NumPeriods = val(CLng(Config_GetFromINI("AF", _
                                                        "MinRampDown_NumPeriods", _
                                                        "500", _
                                                        Prog_INIFile)))
    MaxRampDown_NumPeriods = val(CLng(Config_GetFromINI("AF", _
                                                        "MaxRampDown_NumPeriods", _
                                                        "5000", _
                                                        Prog_INIFile)))
    RampDownNumPeriodsPerVolt = val(CLng(Config_GetFromINI("AF", _
                                                           "RampDownNumPeriodsPerVolt", _
                                                           "1000", _
                                                           Prog_INIFile)))
                            
    HoldAtPeakField_NumPeriods = val(CLng(Config_GetFromINI("AF", _
                                                        "HoldAtPeakField_NumPeriods", _
                                                        "100", _
                                                        Prog_INIFile)))
                                                        
                                               
    'Units for the above values and for displaying values for the AF & IRM field calibration
    'and for saving AF / IRM data to file
    AFUnits = Trim(Config_GetFromINI("AF", _
                                     "AFUnits", _
                                     "G", _
                                     Prog_INIFile))
            
    '2G AF settings
    AFDelay = val(Config_GetFromINI("AF", "AFDelay", "1", Prog_INIFile))
    AFRampRate = val(Config_GetFromINI("AF", "AFRampRate", "3", Prog_INIFile))
    AFWait = Config_GetFromINI("AF", "AFWait", "90", Prog_INIFile)
    AfAxialCoord = Config_GetFromINI("AFAxial", "AFAxialCoord", "Z", Prog_INIFile)
    AfTransCoord = Config_GetFromINI("AFTrans", "AFTransCoord", "Y", Prog_INIFile)
    
    'AF Temperature Sensor settings
    TSlope = val(Config_GetFromINI("AF", "TSlope", "58.86", Prog_INIFile))
    Toffset = val(Config_GetFromINI("AF", "Toffset", "289.6", Prog_INIFile))
    Thot = CInt(Config_GetFromINI("AF", "Thot", "40", Prog_INIFile))
    Tmax = CInt(Config_GetFromINI("AF", "Tmax", "50", Prog_INIFile))
    Tunits = Trim(Config_GetFromINI("AF", "Tunits", "C", Prog_INIFile))
    SystemTUnits = Tunits
    
    'New AF Settings for Axial & Trans Coils
    '(July 2010, I Hilburn)
    AfAxialResFreq = val(Config_GetFromINI("AFAxial", _
                                            "AFAxialResFreq", _
                                            "-1", _
                                            Prog_INIFile))
    AfAxialRampMax = val(Config_GetFromINI("AFAxial", _
                                            "AFAxialRampMax", _
                                            "-1", _
                                            Prog_INIFile))
    AfAxialMonMax = val(Config_GetFromINI("AFAxial", _
                                           "AFAxialMonMax", _
                                           "-1", _
                                           Prog_INIFile))
    AfTransResFreq = val(Config_GetFromINI("AFTrans", _
                                            "AFTransResFreq", _
                                            "-1", _
                                            Prog_INIFile))
    AfTransRampMax = val(Config_GetFromINI("AFTrans", _
                                            "AFTransRampMax", _
                                            "-1", _
                                            Prog_INIFile))
    AfTransMonMax = val(Config_GetFromINI("AFTrans", _
                                           "AFTransMonMax", _
                                           "-1", _
                                           Prog_INIFile))
                                           
    'Read setting - has the AF coils' field calibrations have been done?
    AFAxialCalDone = (Trim(Config_GetFromINI("AFAxial", _
                                             "AFAxialCalDone", _
                                             "False", _
                                             Prog_INIFile)) = "True")
    AFTransCalDone = (Trim(Config_GetFromINI("AFTrans", _
                                             "AFTransCalDone", _
                                             "False", _
                                             Prog_INIFile)) = "True")
    
    ' Now the Af Axial coil calibration numbers
    AfAxialYpoint = val(Config_GetFromINI("AFAxial", "AFAxialYPoint", "979.4", Prog_INIFile))
    AfAxialXpoint = val(Config_GetFromINI("AFAxial", "AFAxialXPoint", "1214", Prog_INIFile))
    AfAxialLowSlope = val(Config_GetFromINI("AFAxial", "AFAxialLowSlope", "0.805975", Prog_INIFile))
    AfAxialHighSlope = val(Config_GetFromINI("AFAxial", "AFAxialHighSlope", "0.791625", Prog_INIFile))
    AfAxialMax = val(Config_GetFromINI("AFAxial", "AFAxialMax", "2900", Prog_INIFile))
    AfAxialMin = val(Config_GetFromINI("AFAxial", "AFAxialMin", "15", Prog_INIFile))
    
    ' Now the Af Transverse coil calibration numbers
    AfTransYpoint = val(Config_GetFromINI("AFTrans", "AFTransYPoint", "139.8", Prog_INIFile))
    AfTransXpoint = val(Config_GetFromINI("AFTrans", "AFTransXPoint", "240", Prog_INIFile))
    AfTransLowSlope = val(Config_GetFromINI("AFTrans", "AFTransLowSlope", "0.660032", Prog_INIFile))
    AfTransHighSlope = val(Config_GetFromINI("AFTrans", "AFTransHighSlope", "0.644856", Prog_INIFile))
    AfTransMax = val(Config_GetFromINI("AFTrans", "AFTransMax", "850", Prog_INIFile))
    AfTransMin = val(Config_GetFromINI("AFTrans", "AFTransMin", "15", Prog_INIFile))
    
    ' Now the IRM Pulse settings (coil non-specific)
    TrimOnTrue = (Trim(Config_GetFromINI("IRMPulse", "TrimOnTrue", "True", Prog_INIFile)) = "True")
    IRMAxis = Config_GetFromINI("IRMPulse", "IRMAxis", "X", Prog_INIFile)
    IRMBackfieldAxis = Config_GetFromINI("IRMPulse", "IRMBackfieldAxis", "Y", Prog_INIFile)
    PulseMCCVoltConversion = val(Config_GetFromINI("IRMPulse", "PulseMCCVoltConversion", "0.022222", Prog_INIFile))
    PulseReturnMCCVoltConversion = val(Config_GetFromINI("IRMPulse", "PulseReturnMCCVoltConversion", "0.022222", Prog_INIFile))
    PulseVoltMax = val(Config_GetFromINI("IRMPulse", "PulseVoltMax", "10", Prog_INIFile))
    AxialTransMaxCapVoltsSame = ("True" = Config_GetFromINI("IRMPulse", "AxialTransMaxCapVoltsSame", "True", Prog_INIFile))
    
    'IRM Axial Coil Pulse settings
    IRMAxialCalDone = (Trim(Config_GetFromINI("IRMAxial", "IRMAxialCalDone", "False", Prog_INIFile)) = "True")
    IRMAxialVoltMax = val(Config_GetFromINI("IRMAxial", "IRMAxialVoltMax", "400", Prog_INIFile))
    PulseAxialMax = val(Config_GetFromINI("IRMAxial", "PulseAxialMax", "13080", Prog_INIFile))
    PulseAxialMin = val(Config_GetFromINI("IRMAxial", "PulseAxialMin", "50", Prog_INIFile))
    AscIrmMaxFireAtZeroGaussReadVoltage = val(Config_GetFromINI("IRMPulse", "AscIrmMaxFireAtZeroGaussReadVoltage", "50", Prog_INIFile))
    AscSetVoltageMaxBoostMultiplier = val(Config_GetFromINI("IRMPulse", "AscSetVoltageMaxBoostMultiplier", "1", Prog_INIFile))
    AscSetVoltageMinBoostMultiplier = val(Config_GetFromINI("IRMPulse", "AscSetVoltageMinBoostMultiplier", "1", Prog_INIFile))
    
    
    'IRM Transverse Coil Pulse settings
    IRMTransCalDone = (Trim(Config_GetFromINI("IRMTrans", "IRMTransCalDone", "False", Prog_INIFile)) = "True")
    IRMTransVoltMax = val(Config_GetFromINI("IRMTrans", "IRMTransVoltMax", "400", Prog_INIFile))
    PulseTransMax = val(Config_GetFromINI("IRMTrans", "PulseTransMax", "13080", Prog_INIFile))
    PulseTransMin = val(Config_GetFromINI("IRMTrans", "PulseTransMin", "50", Prog_INIFile))
        
    ' Now the ARM calibration values as well
    ARMMax = val(Config_GetFromINI("ARM", "ARMMax", "20", Prog_INIFile))
    ARMVoltGauss = val(Config_GetFromINI("ARM", "ARMVoltGauss", "0.1033", Prog_INIFile))
    ARMVoltMax = val(Config_GetFromINI("ARM", "ARMVoltMax", "2.0", Prog_INIFile))
    ARMTimeMax = val(Config_GetFromINI("ARM", "ARMTimeMax", "600", Prog_INIFile))
    

    ' Now the vacuum options
    DoVacuumReset = (Config_GetFromINI("Vacuum", "DoVacuumReset", "False", Prog_INIFile) = "True")
    DropoffVacuumDelay = val(Config_GetFromINI("Vacuum", "DropoffVacuumDelay", "1", Prog_INIFile))
    DoDegausserCooling = (Config_GetFromINI("Vacuum", "DoDegausserCooling", "False", Prog_INIFile) = "True")
    
    ' Now assign the COMM Ports to the proper lines!
    COMPortSquids = Config_GetFromINI("COMPorts", "COMPortSquids", "10", Prog_INIFile)
    COMPortAf = Config_GetFromINI("COMPorts", "COMPortAf", "9", Prog_INIFile)
    COMPortUpDown = Config_GetFromINI("COMPorts", "COMPortUpDown", "4", Prog_INIFile)
    COMPortTurning = Config_GetFromINI("COMPorts", "COMPortTurning", "5", Prog_INIFile)
    COMPortChanger = Config_GetFromINI("COMPorts", "COMPortChanger", "6", Prog_INIFile)
    COMPortChangerY = Config_GetFromINI("COMPorts", "COMPortChangerY", "7", Prog_INIFile)
    COMPortVacuum = Config_GetFromINI("COMPorts", "COMPortVacuum", "3", Prog_INIFile)
    COMPortSusceptibility = Config_GetFromINI("COMPorts", "COMPortSusceptibility", "8", Prog_INIFile)
    SusceptibilitySettings = Config_GetFromINI("COMPorts", "SusceptibilitySettings", "9600,N,8,2", Prog_INIFile)
    
    ' Now settings for mailer application
    MailSMTPHost = Config_GetFromINI("Email", "MailSMTPHost", vbNullString, Prog_INIFile)
    MailSMTPPort = CInt(Config_GetFromINI("Email", "MailSMTPPort", "25", Prog_INIFile))
    MailFrom = Config_GetFromINI("Email", "MailFrom", "paleomag@localhost", Prog_INIFile)
    MailFromName = Config_GetFromINI("Email", "MailFromName", "2G Magnetometer Sample Changer", Prog_INIFile)
    MailFromPassword = Config_GetFromINI("Email", "MailFromPassword", "xxxx", Prog_INIFile)
    MailCCList = Config_GetFromINI("Email", "MailCCList", vbNullString, Prog_INIFile)
    MailStatusMonitor = Config_GetFromINI("Email", "MailStatusMonitor", vbNullString, Prog_INIFile)
    
    ' Settings for SSL Encrypted / Remote SMTP emails
    MailSMTPPassword = Config_GetFromINI("Email", "MailSMTPPassword", vbNullString, Prog_INIFile)
    MailSMTPUsername = Config_GetFromINI("Email", "MailSMTPUsername", vbNullString, Prog_INIFile)
    MailSMTPAuthenticate = CLng(Config_GetFromINI("Email", "MailSMTPAuthenticate", "0", Prog_INIFile))
    MailSMTPSendUsing = CLng(Config_GetFromINI("Email", "MailSMTPSendUsing", "2", Prog_INIFile))
    MailUseSSLEncryption = (Config_GetFromINI("Email", "MailUseSSLEncryption", "False", Prog_INIFile) = "True")
       
    
    'Load XY Table Positions
    UseXYTableAPS = (Trim(Config_GetFromINI("XYTable", _
                                             "UseXYTableAPS", _
                                             "False", _
                                             Prog_INIFile)) = "True")
    XYTablePositions(0, 0) = val(Config_GetFromINI("XYTable", "XYHomeX", vbNullString, Prog_INIFile))
    XYTablePositions(0, 1) = val(Config_GetFromINI("XYTable", "XYHomeY", vbNullString, Prog_INIFile))
    
    Dim temp As String
    For i = 1 To 100
        temp = "XY" + LTrim$(str(i)) + "X"
        XYTablePositions(i, 0) = val(Config_GetFromINI("XYTable", temp, vbNullString, Prog_INIFile))
        temp = "XY" + LTrim$(str(i)) + "Y"
        XYTablePositions(i, 1) = val(Config_GetFromINI("XYTable", temp, vbNullString, Prog_INIFile))
    Next i
'------------------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------------------
'    (March 2010 - Isaac Hilburn)
'
'     Modified AF field calibration arrays (see variable declaration up top)
'     so that the arrays can be dynamically resized to any dimension
'
'     Arrays for the AF calibration are now N x 2 where N = number of calibration points +
'     1 additional element to hold the zero value.  For the 2G calibration, Col 0 =
'     2G 8-bit counts value, Col 1 = matching DC peak field.  For the ADWIN AF calibration,
'     Col 0 = Monitor Peak Voltage for the AF Ramp using the ADwin board, Col 1 = resulting
'     DC peak field for that Monitor peak voltage.
'------------------------------------------------------------------------------------------------------------------------

    'AF Field Calibration Arrays
    AFAxialCount = val(Config_GetFromINI("AFAxial", "AFAxialCount", "0", Prog_INIFile))
    AFTransCount = val(Config_GetFromINI("AFTrans", "AFTransCount", "0", Prog_INIFile))
    
    'Redimension AFAxial and AFTrans arrays (Remember to add one point for the zeros)
    ReDim AFAxial(AFAxialCount + 1, 2)
    ReDim AFTrans(AFTransCount + 1, 2)
    
'(July 31, 2010 - I Hilburn)
' Commented Out
'See global variable declarations at top of modConfig for explanation
'    'Redimension the AFRampAxial and AFRampTrans arrays
'    ReDim AFRampAxial(AFRampAxialCount + 1, 2)
'    ReDim AFRampTrans(AFRampTransCount + 1, 2)
    
    'Set first elements of all four arrays to zero
    AFAxial(0, 0) = 0
    AFAxial(0, 1) = 0
    AFTrans(0, 0) = 0
    AFTrans(0, 1) = 0
    
'(July 31, 2010 - I Hilburn)
' Commented Out
'See global variable declarations at top of modConfig for explanation
'    AFRampAxial(0, 0) = 0
'    AFRampAxial(0, 1) = 0
'    AFRampTrans(0, 0) = 0
'    AFRampTrans(0, 1) = 0
    
    
    If AFAxialCount > 0 Then
    
        'Run through Axial Coil calibration points stored in the INI file
        For i = 1 To AFAxialCount
            
            AFAxial(i, 0) = val(Config_GetFromINI("AFAxial", _
                                                  "AFAxialX" & Format$(i, "0"), _
                                                  "0", _
                                                  Prog_INIFile))
                                                  
            AFAxial(i, 1) = val(Config_GetFromINI("AFAxial", _
                                                  "AFAxialY" & Format$(i, "0"), _
                                                  "0", _
                                                  Prog_INIFile))
        
        Next i
    
    End If
    
    
    If AFTransCount > 0 Then
    
        'Run through Transverse Coil calibration points stored in the INI file
        For i = 1 To AFTransCount
            
            AFTrans(i, 0) = val(Config_GetFromINI("AFTrans", _
                                                  "AFTransX" & Format$(i, "0"), _
                                                  "0", _
                                                  Prog_INIFile))
                                                  
            AFTrans(i, 1) = val(Config_GetFromINI("AFTrans", _
                                                  "AFTransY" & Format$(i, "0"), _
                                                  "0", _
                                                  Prog_INIFile))
        
        Next i
        
    End If
    
'---------------------------------------------------------------------------------------------------------------------------'
'---------------------------------------------------------------------------------------------------------------------------'
'(July 31, 2010 - I Hilburn)
' Commented Out
'See global variable declarations at top of modConfig for explanation
'
'---------------------------------------------------------------------------------------------------------------------------'
'
'    If AFRampAxialCount > 0 Then
'
'        'Run through Axial Coil calibration points stored in the INI file
'        For i = 1 To AFRampAxialCount
'
'            AFRampAxial(i, 0) = val(Config_GetFromINI( _
'                                            "AFAxial", _
'                                            "AFAxialRamp" & Format$(i, "0"), _
'                                            "0", _
'                                            Prog_INIFile))
'
'            AFRampAxial(i, 1) = val(Config_GetFromINI( _
'                                           "AFAxial", _
'                                            "AFAxialMon" & Format$(i, "0"), _
'                                            "0", _
'                                            Prog_INIFile))
'
'        Next i
'
'    End If
'
'
'    If AFRampTransCount > 0 Then
'
'        'Run through Transverse Coil calibration points stored in the INI file
'        For i = 1 To AFRampTransCount
'
'            AFRampTrans(i, 0) = val(Config_GetFromINI( _
'                                            "AFTrans", _
'                                            "AFTransRamp" & Format$(i, "0"), _
'                                            "0", _
'                                            Prog_INIFile))
'
'            AFRampTrans(i, 1) = val(Config_GetFromINI( _
'                                            "AFTrans", _
'                                            "AFTransMon" & Format$(i, "0"), _
'                                            "0", _
'                                            Prog_INIFile))
'
'        Next i
'
'    End If
'
'---------------------------------------------------------------------------------------------------------------------------'
'---------------------------------------------------------------------------------------------------------------------------'
    
    ' IRM Pulse Field Calibration Arrays
    PulseAxialCount = val(Config_GetFromINI("IRMAxial", _
                                         "PulseAxialCount", _
                                         "0", _
                                         Prog_INIFile))
                                         
    PulseTransCount = val(Config_GetFromINI("IRMTrans", _
                                         "PulseTransCount", _
                                         "0", _
                                         Prog_INIFile))
    
    'Redimension IRM calibration arrays = N x 2 dimensions (Remember to add one point for the zeros)
    ReDim PulseAxial(PulseAxialCount + 1, 2)
    ReDim PulseTrans(PulseTransCount + 1, 2)
    
    PulseAxial(0, 0) = 0
    PulseAxial(0, 1) = 0
    PulseTrans(0, 0) = 0
    PulseTrans(0, 1) = 0
    
    If PulseAxialCount > 0 Then
    
        'Run through Axial Coil calibration points stored in the INI file
        For i = 1 To PulseAxialCount
            
            PulseAxial(i, 0) = val(Config_GetFromINI("IRMAxial", _
                                                  "PulseAxialX" & Format$(i, "0"), _
                                                   "0", _
                                                   Prog_INIFile))
                                                   
            PulseAxial(i, 1) = val(Config_GetFromINI("IRMAxial", _
                                                  "PulseAxialY" & Format$(i, "0"), _
                                                  "0", _
                                                  Prog_INIFile))
        
        Next i
        
    End If
    
    If PulseTransCount > 0 Then
    
        'Run through Axial Coil calibration points stored in the INI file
        For i = 1 To PulseTransCount
            
            PulseTrans(i, 0) = val(Config_GetFromINI("IRMTrans", _
                                                  "PulseTransX" & Format$(i, "0"), _
                                                  "0", _
                                                  Prog_INIFile))
                                                  
            PulseTrans(i, 1) = val(Config_GetFromINI("IRMTrans", _
                                                  "PulseTransY" & Format$(i, "0"), _
                                                  "0", _
                                                  Prog_INIFile))
        
        Next i
    
    End If
    
'------------------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------------------
    
    ' now for the rockmag enable/disable modules
    EnableAxialIRM = (Trim(Config_GetFromINI("Modules", "EnableAxialIRM", "True", Prog_INIFile)) = "True")
    EnableTransIRM = (Trim(Config_GetFromINI("Modules", "EnableTransIRM", "False", Prog_INIFile)) = "True")
    EnableIRMBackfield = (Trim(Config_GetFromINI("Modules", "EnableIRMBackfield", "False", Prog_INIFile)) = "True")
    EnableIRMMonitor = (Trim(Config_GetFromINI("Modules", "EnableIRMMonitor", "False", Prog_INIFile)) = "True")
    EnableARM = (Trim(Config_GetFromINI("Modules", "EnableARM", "True", Prog_INIFile)) = "True")
    EnableAF = (Trim(Config_GetFromINI("Modules", "EnableAF", "True", Prog_INIFile)) = "True")
    EnableAltAFMonitor = (Trim(Config_GetFromINI("Modules", "EnableAltAFMonitor", "False", Prog_INIFile)) = "True")
    EnableT1 = (Trim(Config_GetFromINI("Modules", "EnableT1", "False", Prog_INIFile)) = "True")
    EnableT2 = (Trim(Config_GetFromINI("Modules", "EnableT2", "False", Prog_INIFile)) = "True")
    EnableSusceptibility = (Trim(Config_GetFromINI("Modules", "EnableSusceptibility", "True", Prog_INIFile)) = "True")
    EnableAFAnalysis = (Trim(Config_GetFromINI("Modules", "EnableAFAnalysis", "False", Prog_INIFile)) = "True")
    EnableVacuum = (Trim(Config_GetFromINI("Modules", "EnableVacuum", "True", Prog_INIFile)) = "True")
    EnableDegausserCooler = (Trim(Config_GetFromINI("Modules", "EnableDegausserCooler", "True", Prog_INIFile)) = "True")
        
'------------------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------------------
'
'   Major Comm Modification
'   March-April 2010
'   Isaac Hilburn
'
'   Changed DAQ board and DAQ channel/port assignment implementation to use new
'   Board & Channel objects and object collections
'
'   Allows user to hypothetically use any DAQ board with the correct # of Analog Input,
'   Analog Output, & DIO channels to run, read, and fire the necessary values to
'   control the AF, IRM, ARM process and toggle the Vacuum relays and motor on/off.
'
'------------------------------------------------------------------------------------------------------------------------
'   Legacy Code Using the Old port-num assignments from .Ini file:
'
'    '(March 2008 L Carporzen) Put in Settings the IRM/ARM channels (MIT acquisition board does not work on IRMTrim = 3
'
'    'Analog channel output
'    ARMVoltageOut = val(Config_GetFromINI("IRM-ARM", "ARMVoltageOut", "0", Prog_INIFile))
'    IRMVoltageOut = val(Config_GetFromINI("IRM-ARM", "IRMVoltageOut", "1", Prog_INIFile))
'
'    'Analog input
'    IRMCapacitorVoltageIn = val(Config_GetFromINI("IRM-ARM", "IRMCapacitorVoltageIn", "0", Prog_INIFile))
'    AnalogT1 = val(Config_GetFromINI("AF", "AnalogT1", "1", Prog_INIFile))
'    AnalogT2 = val(Config_GetFromINI("AF", "AnalogT2", "2", Prog_INIFile))
'
'    'DIO line assignments
'    ARMSet = val(Config_GetFromINI("IRM-ARM", "ARMSet", "0", Prog_INIFile))
'    IRMFire = val(Config_GetFromINI("IRM-ARM", "IRMFire", "1", Prog_INIFile))
'    IRMTrim = val(Config_GetFromINI("IRM-ARM", "IRMTrim", "3", Prog_INIFile))
'    IRMPowerAmpVoltageIn = val(Config_GetFromINI("IRM-ARM", "IRMPowerAmpVoltageIn", "4", Prog_INIFile))
'    MotorToggle = val(Config_GetFromINI("Vacuum", "MotorToggle", "5", Prog_INIFile))
'    VacuumToggleA = val(Config_GetFromINI("Vacuum", "VacuumToggleA", "6", Prog_INIFile))
'    VacuumToggleB = val(Config_GetFromINI("Vacuum", "VacuumToggleB", "7", Prog_INIFile))
'------------------------------------------------------------------------------------------------------------------------
'
'   New Code
'   March, 2010
'   Isaac Hilburn
'
'   Paleomag.ini structure has been altered to store all the parameters that are needed
'   about the Measurement Computing and ADWIN DAQ Boards that RAPID uses.
'
'   a New [Boards] and [Channels] section has been added to the .ini file
'   with pre-loaded settings for the PCI-DAS6030 board and the ADWIN-light-16 board
'
'   Additionally, the channel assignments have been made for the various non-serial port
'   com settings in the [Channels] section
'
'   Now that the DAQ Boards & PC are conspiring together to generate the AF Ramp waveform,
'   the settings for this ramp signal and Analog IO process need to be stored in the Paleomag.INI
'   file.
'
'   A new Wave object and Waves object collections have been created to handle the storing and
'   passing of the necessary information between forms and modules in the Paleomag program
'   These Wave objects need to be imported here in modConfig from the new [Wave Forms] section
'   of the .INI file.
'-------------------------------------------------------------------------------------------------'

    ImportBoardsDone = False

    'Import DAQ Board(s) settings into Board objects in the SystemBoards collection
    Get_BoardsFromIni
    
    ImportWavesDone = False

    'Import information for loading into the necessary wave objects (load up WaveForms collection)
    'Five Standard Waveforms for systems using a DAQ implemented AF / ARM system
    '   These five waveforms contain all the settings needed to run the
    '   DAQ Af process
    'Plus 1 waveform for systems using the DAQ or the 2G AF system
    Get_WaveFormsFromIni
    
    ImportChannelsDone = False
    
    'Import DAQ Board Channel assignments - assigned to channel objects that
    'replaced the Integer port-number assignments under the old implementation
    Get_ChannelsFromIni
     
'------------------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------------------
        
End Sub

Private Function Config_ReadLine(filenum As Integer) As String
    ' This function reads a line from the file with id
    ' #filenum, and returns a string from the file.  It
    ' ignores lines that start with "'".
    Dim lchar As String
    Line Input #filenum, Config_ReadLine
    lchar = Left$(Config_ReadLine, 1)
    If lchar = Chr$(39) Then
        Config_ReadLine = Config_ReadLine(filenum)
    End If
End Function

Public Sub Config_SaveSetting(sSection As String, sKey As String, sDefault As String)
    '(April 2010, I Hilburn - New version of Config_AddToINI has a string return type
    Dim dummy As String
    dummy = Config_AddToINI(sSection, sKey, sDefault, Prog_INIFile)
End Sub

Public Sub Config_writeAFCommSettingstoINI()
    
    Dim i As Integer
    Dim CurTime
    
    CurTime = Now
    
    'Back-up the old .INI settings file
    FileCopy Prog_INIFile, Prog_INIFile & Format(CurTime, "_MM-DD-YYYY_HH-MM-SS") & ".bak"
    
    'Save the Units
    Config_SaveSetting "AF", "AFUnits", AFUnits
    
    'Save the Alternate AF Monitor Module enables setting
    Config_SaveSetting "Modules", "EnableAltAFMonitor", str$(EnableAltAFMonitor)
    
    'All the AF channel settings are stored in the WaveForms section of the .INI file
    'So just need to write all the wave-forms to the .INI file
    Save_WaveFormsToIni

End Sub

Public Sub Config_writeSettingstoINI()
    Dim i As Integer
    Dim CurTime
    
    CurTime = Now
    
    FileCopy Prog_INIFile, Prog_INIFile & Format(CurTime, "_MM-DD-YYYY_HH-MM-SS") & ".bak"
    
    'Sample Changer Position Settings
    Config_SaveSetting "SampleChanger", "SlotMin", str$(SlotMin)
    Config_SaveSetting "SampleChanger", "SlotMax", str$(SlotMax)
    Config_SaveSetting "SampleChanger", "OneStep", str$(OneStep)
    Config_SaveSetting "SampleChanger", "HoleSlotNum", str$(HoleSlotNum)
    
    'DC Stepping Motor Settings
    'Up/Down Motor
    Config_SaveSetting "SteppingMotor", "ZeroPos", str$(ZeroPos)
    Config_SaveSetting "SteppingMotor", "MeasPos", str$(MeasPos)
    Config_SaveSetting "SteppingMotor", "IRMPos", str$(IRMPos)
    Config_SaveSetting "SteppingMotor", "AFPos", str$(AFPos)
    Config_SaveSetting "SteppingMotor", "SCoilPos", str$(SCoilPos)
    Config_SaveSetting "SteppingMotor", "FloorPos", str$(FloorPos)
    Config_SaveSetting "SteppingMotor", "MinUpDownPos", str$(MinUpDownPos)
    Config_SaveSetting "SteppingMotor", "SampleBottom", str$(SampleBottom)
    Config_SaveSetting "SteppingMotor", "SampleTop", str$(SampleTop)
    Config_SaveSetting "SteppingMotor", "LiftSpeedSlow", str$(LiftSpeedSlow)
    Config_SaveSetting "SteppingMotor", "LiftSpeedNormal", str$(LiftSpeedNormal)
    Config_SaveSetting "SteppingMotor", "LiftSpeedFast", str$(LiftSpeedFast)
    Config_SaveSetting "SteppingMotor", "LiftAcceleration", str$(LiftAcceleration)
    Config_SaveSetting "SteppingMotor", "UpDownMotor1cm", str$(UpDownMotor1cm)
    Config_SaveSetting "SteppingMotor", "UpDownTorqueFactor", str$(UpDownTorqueFactor)
    Config_SaveSetting "SteppingMotor", "UpDownMaxTorque", str$(UpDownMaxTorque)
    Config_SaveSetting "SteppingMotor", "PickupTorqueThrottle", str$(PickupTorqueThrottle)
    
    'Changer Belt Motor
    Config_SaveSetting "SteppingMotor", "SampleHoleAlignmentOffset", str$(SampleHoleAlignmentOffset)
    Config_SaveSetting "SteppingMotor", "ChangerSpeed", str$(ChangerSpeed)
    Config_SaveSetting "SteppingMotor", "SCurveFactor", str$(SCurveFactor)
    Config_SaveSetting "SteppingMotor", "TrayOffsetAngle", str$(TrayOffsetAngle)
    
    'Turning Motor
    Config_SaveSetting "SteppingMotor", "TurnerSpeed", str$(TurnerSpeed)
    Config_SaveSetting "SteppingMotor", "TurningMotorFullRotation", str$(TurningMotorFullRotation)
    Config_SaveSetting "SteppingMotor", "TurningMotor1rps", str$(TurningMotor1rps)
    
    'Magnetometer/SQUID Calibration Settings
    Config_SaveSetting "MagnetometerCalibration", "ZCal", str$(ZCal)
    Config_SaveSetting "MagnetometerCalibration", "XCal", str$(XCal)
    Config_SaveSetting "MagnetometerCalibration", "YCal", str$(YCal)
    Config_SaveSetting "MagnetometerCalibration", "RangeFact", str$(RangeFact)
    Config_SaveSetting "MagnetometerCalibration", "ReadDelay", str$(ReadDelay) ' (March 2008 L Carporzen) Read delay
    
    'SQUID Jump / Remeasure threshold settings
    Config_SaveSetting "Magnetometry", "RemeasureCSDThreshold", str$(RemeasureCSDThreshold)
    
    ' New selections in the Options menu (April-May 2007 L Carporzen)
    Config_SaveSetting "Magnetometry", "JumpThreshold", str$(JumpThreshold)
    Config_SaveSetting "Magnetometry", "StrongMom", str$(StrongMom)
    Config_SaveSetting "Magnetometry", "IntermMom", str$(IntermMom)
    Config_SaveSetting "Magnetometry", "MomMinForRedo", str$(MomMinForRedo)
    Config_SaveSetting "Magnetometry", "JumpSensitivity", str$(JumpSensitivity)
    Config_SaveSetting "Magnetometry", "NbTry", str$(NbTry)
    Config_SaveSetting "Magnetometry", "NbHolderTry", str$(NbHolderTry)
    
    'Susceptibility Settings
    Config_SaveSetting "SusceptibilityCalibration", "SusceptibilityMomentFactorCGS", str$(SusceptibilityMomentFactorCGS)
    Config_SaveSetting "SusceptibilityCalibration", "SusceptibilityScaleFactor", str$(SusceptibilityScaleFactor)

    'AF System Settings
    '(April - May 2010, Isaac Hilburn)
    'This setting allows users to toggle between using the 2G or the ADWIN/DAQ versions
    'of the AF system
    Config_SaveSetting "AF", "AFSystem", Trim(AFSystem)
    Config_SaveSetting "AF", "AFUnits", AFUnits
            
    'Save ADWIN Boot & Program File Names and Directory path info
    '(July 2010, Isaac Hilburn)
    Config_SaveSetting "AF", "ADWINBinFolderPath", Trim(ADWINBinFolderPath)
    Config_SaveSetting "AF", "ADWINBootFile", Trim(ADWINBootFileName)
    Config_SaveSetting "AF", "ADWINRampProgFile", Trim(ADWINRampProgFileName)
            
    'DAQ AF Monitor Data file save path information and settings
    '(July 2010, Isaac Hilburn)
    Config_SaveSetting "AFFileSave", _
                       "ADWINDataFileSaveLocalDir", _
                       ADWIN_AFDataLocalDir
    Config_SaveSetting "AFFileSave", _
                       "ADWINDataFileSaveBackupDir", _
                       ADWIN_AFDataBackupDir
    Config_SaveSetting "AFFileSave", _
                       "2GDataFileSaveLocalDir", _
                       TWOG_AFDataLocalDir
    Config_SaveSetting "AFFileSave", _
                       "2GDataFileSaveBackupDir", _
                       TWOG_AFDataBackupDir
    Config_SaveSetting "AFFileSave", _
                       "AFDataFileSaveDoBackup", _
                       str$(modConfig.AFDoDataFileBackup)
                       
    '2G AF Settings
    Config_SaveSetting "AF", "AFDelay", str$(AFDelay)
    Config_SaveSetting "AF", "AFRampRate", str$(AFRampRate)
    Config_SaveSetting "AF", "AFWait", str$(AFWait)
    Config_SaveSetting "AFAxial", "AFAxialCoord", AfAxialCoord
    Config_SaveSetting "AFTrans", "AFTransCoord", AfTransCoord
    
    'AF Temperature Sensor Settings
    Config_SaveSetting "AF", "TSlope", str$(TSlope)
    Config_SaveSetting "AF", "Toffset", str$(Toffset)
    Config_SaveSetting "AF", "Thot", str$(Thot)
    Config_SaveSetting "AF", "Tmax", str$(Tmax)
    Config_SaveSetting "AF", "Tunits", Trim(Tunits)
    
    'AF ADWIN Ramp Settings
    '(August 2010, Isaac Hilburn)
    Config_SaveSetting "AF", _
                       "MaxRampUpTime_ms", _
                       Trim(str(MaxRampUpTime_ms))
                                                        
    Config_SaveSetting "AF", _
                       "MinRampUpTime_ms", _
                       Trim(str(modConfig.MinRampUpTime_ms))
                       
    Config_SaveSetting "AF", _
                       "AxialRampUpVoltsPerSec", _
                       Trim(str(modConfig.AxialRampUpVoltsPerSec))
                       
    Config_SaveSetting "AF", _
                       "TransRampUpVoltsPerSec", _
                       Trim(str(modConfig.TransRampUpVoltsPerSec))
                       
    Config_SaveSetting "AF", _
                       "MinRampDown_NumPeriods", _
                       Trim(str(MinRampDown_NumPeriods))
                       
    Config_SaveSetting "AF", _
                       "MaxRampDown_NumPeriods", _
                       Trim(str(MaxRampDown_NumPeriods))
                       
    Config_SaveSetting "AF", _
                       "RampDownNumPeriodsPerVolt", _
                       Trim(str(RampDownNumPeriodsPerVolt))
                       
    Config_SaveSetting "AF", _
                       "HoldAtPeakField_NumPeriods", _
                       Trim(str(HoldAtPeakField_NumPeriods))
                       
    
    'New AF Settings for Axial & Trans Coils
    '(July 2010, I Hilburn)
    Config_SaveSetting "AFAxial", "AFAxialResFreq", str$(AfAxialResFreq)
    Config_SaveSetting "AFAxial", "AFAxialRampMax", str$(AfAxialRampMax)
    Config_SaveSetting "AFAxial", "AFAxialMonMax", str$(AfAxialMonMax)
    Config_SaveSetting "AFTrans", "AFTransResFreq", str$(AfTransResFreq)
    Config_SaveSetting "AFTrans", "AFTransRampMax", str$(AfTransRampMax)
    Config_SaveSetting "AFTrans", "AFTransMonMax", str$(AfTransMonMax)
    
    'Save setting indicating whether or not the AF coils are field calibrated
    Config_SaveSetting "AFTrans", "AFTransCalDone", Trim(str$(AFTransCalDone))
    Config_SaveSetting "AFAxial", "AFAxialCalDone", Trim(str$(AFAxialCalDone))
    
    'AF Settings + some calibration
    Config_SaveSetting "AFAxial", "AFAxialYPoint", str$(AfAxialYpoint)
    Config_SaveSetting "AFAxial", "AfAxialXpoint", str$(AfAxialXpoint)
    Config_SaveSetting "AFAxial", "AfAxialHighSlope", str$(AfAxialHighSlope)
    Config_SaveSetting "AFAxial", "AfAxialLowSlope", str$(AfAxialLowSlope)
    Config_SaveSetting "AFAxial", "AfAxialMax", str$(AfAxialMax)
    Config_SaveSetting "AFAxial", "AfAxialMin", str$(AfAxialMin)
    Config_SaveSetting "AFTrans", "AFTransYPoint", str$(AfTransYpoint)
    Config_SaveSetting "AFTrans", "AfTransXpoint", str$(AfTransXpoint)
    Config_SaveSetting "AFTrans", "AfTransHighSlope", str$(AfTransHighSlope)
    Config_SaveSetting "AFTrans", "AfTransLowSlope", str$(AfTransLowSlope)
    Config_SaveSetting "AFTrans", "AfTransMax", str$(AfTransMax)
    Config_SaveSetting "AFTrans", "AfTransMin", str$(AfTransMin)
    
    'IRM settings (non-coil specific)
    Config_SaveSetting "IRMPulse", "IRMSystem", IRMSystem
    Config_SaveSetting "IRMPulse", "TrimOnTrue", str(TrimOnTrue)
    Config_SaveSetting "IRMPulse", "IRMAxis", IRMAxis
    Config_SaveSetting "IRMPulse", "IRMBackfieldAxis", IRMBackfieldAxis
    Config_SaveSetting "IRMPulse", "PulseMCCVoltConversion", str$(PulseMCCVoltConversion)
    Config_SaveSetting "IRMPulse", "PulseVoltMax", str$(PulseVoltMax)
    Config_SaveSetting "IRMPulse", "PulseReturnMCCVoltConversion", str$(PulseReturnMCCVoltConversion)
    Config_SaveSetting "IRMPulse", "AxialTransMaxCapVoltsSame", str$(AxialTransMaxCapVoltsSame)
    Config_SaveSetting "IRMPulse", "AscIrmMaxFireAtZeroGaussReadVoltage", str$(AscIrmMaxFireAtZeroGaussReadVoltage)
    Config_SaveSetting "IRMPulse", "AscSetVoltageMaxBoostMultiplier", Format$(AscSetVoltageMaxBoostMultiplier, "#0.00")
    Config_SaveSetting "IRMPulse", "AscSetVoltageMinBoostMultiplier", Format$(AscSetVoltageMinBoostMultiplier, "#0.00")
                        
    'IRM Axial settings
    Config_SaveSetting "IRMAxial", "IRMAxialCalDone", str$(IRMAxialCalDone)
    Config_SaveSetting "IRMAxial", "IRMAxialVoltMax", str$(IRMAxialVoltMax)
    Config_SaveSetting "IRMAxial", "PulseAxialMax", str$(PulseAxialMax)
    Config_SaveSetting "IRMAxial", "PulseAxialMin", str$(PulseAxialMin)

    'IRM Transverse settings
    Config_SaveSetting "IRMTrans", "IRMTransCalDone", str$(IRMTransCalDone)
    Config_SaveSetting "IRMTrans", "IRMTransVoltMax", str$(IRMTransVoltMax)
    Config_SaveSetting "IRMTrans", "PulseTransMax", str$(PulseTransMax)
    Config_SaveSetting "IRMTrans", "PulseTransMin", str$(PulseTransMin)
       
    'ARM + calibration settings
    Config_SaveSetting "ARM", "ARMMax", str$(ARMMax)
    Config_SaveSetting "ARM", "ARMVoltGauss", str$(ARMVoltGauss)
    Config_SaveSetting "ARM", "ARMVoltMax", str$(ARMVoltMax)
    Config_SaveSetting "ARM", "ARMTimeMax", str$(ARMTimeMax)

'------------------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------------------
'
'   Major Comm Modification
'   March-April 2010
'   Isaac Hilburn
'
'   Changed DAQ board and DAQ channel/port assignment implementation to use new
'   Board & Channel objects and object collections
'
'   Allows user to hypothetically use any DAQ board with the correct # of Analog Input,
'   Analog Output, & DIO channels to run, read, and fire the necessary values to
'   control the AF, IRM, ARM process and toggle the Vacuum relays and motor on/off.
'
'------------------------------------------------------------------------------------------------------------------------
'
'   Old Legacy Code:
'
'    ' (March 2008 L Carporzen) Put in Settings for the IRM/ARM channels
'    '(MIT acquisition board does not work on IRMTrim = 3
'    ' Analog channel output
'    Config_SaveSetting "IRM-ARM", "ARMVoltageOut", Str$(ARMVoltageOut)
'    Config_SaveSetting "IRM-ARM", "IRMVoltageOut", Str$(IRMVoltageOut)
'    ' Analog input
'    Config_SaveSetting "IRM-ARM", "IRMCapacitorVoltageIn", Str$(IRMCapacitorVoltageIn)
'    Config_SaveSetting "AF", "AnalogT1", Str$(AnalogT1)
'    Config_SaveSetting "AF", "AnalogT2", Str$(AnalogT2)
'    ' DIO line assignments
'    Config_SaveSetting "IRM-ARM", "ARMSet", Str$(ARMSet)
'    Config_SaveSetting "IRM-ARM", "IRMFire", Str$(IRMFire)
'    Config_SaveSetting "IRM-ARM", "IRMTrim", Str$(IRMTrim)
'    Config_SaveSetting "IRM-ARM", "IRMPowerAmpVoltageIn", Str$(IRMPowerAmpVoltageIn)
'    Config_SaveSetting "Vacuum", "MotorToggle", Str$(MotorToggle)
'    Config_SaveSetting "Vacuum", "VacuumToggleA", Str$(VacuumToggleA)
'    Config_SaveSetting "Vacuum", "VacuumToggleB", Str$(VacuumToggleB)
'
'------------------------------------------------------------------------------------------------------------------------
'
'   New Code
'   April, 2010
'   Isaac Hilburn
'
'   Paleomag.ini structure has been altered to store all the parameters that are needed
'   about the Measurement Computing and ADWIN DAQ Boards that RAPID uses.
'
'   a New [Boards] and [Channels] section has been added to the .ini file
'   with pre-loaded settings for the PCI-DAS6030 board and the ADWIN-light-16 board
'
'   Additionally, the channel assignments have been made for the various non-serial port
'   com settings in the [Channels] section

    'Save DAQ Board(s) settings into Board objects in the SystemBoards collection
    Save_BoardsToINI
    
    'Save DAQ Board Channel assignments - from channel objects that
    'replaced the Integer port-number assignments under the old implementation
    Save_ChannelsToINI
    
'------------------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------------------

'------------------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------------------
'
'   Major Modification for Saving AF settings
'   for running the AF/ARM Ramp cycle through MCC or ADWIN DAQ boards
'   March-April, 2010
'   Isaac Hilburn
'
'   Now that the DAQ Boards & PC are conspiring together to generate the AF Ramp waveform,
'   the settings for this ramp signal and Analog IO process need to be stored in the Paleomag.INI
'   file.
'
'   A new Wave object and Waves object collections have been created to handle the storing and
'   passing of the necessary information between forms and modules in the Paleomag program
'   These Wave objects need to be imported here in modConfig from the .INI file.
'------------------------------------------------------------------------------------------------------------------------

    'Save information for loading into the necessary wave objects (load up WaveForms collection)
    'Five Standard Waveforms for systems using a DAQ implemented AF / ARM system
    '   These five waveforms contain all the settings needed to run the
    '   DAQ Af process
    'Plus 1 waveform for systems using the DAQ or the 2G AF system
    Save_WaveFormsToIni

'------------------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------------------
    
    'Vacuum System Settings
    Config_SaveSetting "Vacuum", "DoVacuumReset", str$(DoVacuumReset)
    Config_SaveSetting "Vacuum", "DropoffVacuumDelay", str$(DropoffVacuumDelay)
    Config_SaveSetting "Vacuum", "DoDegausserCooling", str$(DoDegausserCooling)
    
    
    'Serial ComPort settings - these setting have nothing to do with DAQ Boards
    'installed in a users Sample Changer system
    Config_SaveSetting "COMPorts", "COMPortSquids", str$(COMPortSquids)
    Config_SaveSetting "COMPorts", "COMPortAf", str$(COMPortAf)
    Config_SaveSetting "COMPorts", "COMPortUpDown", str$(COMPortUpDown)
    Config_SaveSetting "COMPorts", "COMPortTurning", str$(COMPortTurning)
    Config_SaveSetting "COMPorts", "COMPortChanger", str$(COMPortChanger)
    Config_SaveSetting "COMPorts", "COMPortChangerY", str$(COMPortChangerY)
    Config_SaveSetting "COMPorts", "COMPortVacuum", str$(COMPortVacuum)
    Config_SaveSetting "COMPorts", "COMPortSusceptibility", str$(COMPortSusceptibility)
    Config_SaveSetting "COMPorts", "SusceptibilitySettings", SusceptibilitySettings
    
    'Stepping Motor Program settings - including the motor comm ID's
    'and the Motor command ID's for home to top and Sample pickup
    Config_SaveSetting "MotorPrograms", "CmdHomeToTop", str$(CmdHomeToTop)
    Config_SaveSetting "MotorPrograms", "CmdSamplePickup", str$(CmdSamplePickup)
    Config_SaveSetting "MotorPrograms", "MotorIDTurning", str$(MotorIDTurning)
    Config_SaveSetting "MotorPrograms", "MotorIDChanger", str$(MotorIDChanger)
    Config_SaveSetting "MotorPrograms", "MotorIDChangerY", str$(MotorIDChangerY)
    Config_SaveSetting "MotorPrograms", "MotorIDUpDown", str$(MotorIDUpDown)
    
    'Paleomag program ID's
    Config_SaveSetting "Program", "UsageFile", Prog_UsageFile
    Config_SaveSetting "Program", "DefaultPath", Prog_DefaultPath
    Config_SaveSetting "Program", "HelpURLRoot", Prog_HelpURLRoot
    Config_SaveSetting "Program", "NoCommMode", str$(NOCOMM_MODE)
    Config_SaveSetting "Program", "DebugMode", str$(DEBUG_MODE)
    Config_SaveSetting "Program", "DumpRawDataStats", str$(DumpRawDataStats)
    Config_SaveSetting "Program", "LogMessages", str$(LogMessages)
    Config_SaveSetting "Program", "LogFolderPath", Trim(LogFolderPath)
    Config_SaveSetting "Program", "LogFileName", Trim(LogFileName)
    Config_SaveSetting "Program", "LogoFile", Prog_LogoFile
    Config_SaveSetting "Program", "IcoFile", Prog_IcoFile ' (October 2007 L Carporzen)
    Config_SaveSetting "Program", "TextEditor", Prog_TextEditor
    
    'VBSendMail / Email settings
    Config_SaveSetting "Email", "MailSMTPHost", MailSMTPHost
    Config_SaveSetting "Email", "MailSMTPPort", Trim(str(MailSMTPPort))
    Config_SaveSetting "Email", "MailFrom", MailFrom
    Config_SaveSetting "Email", "MailFromName", MailFromName
    Config_SaveSetting "Email", "MailFromPassword", MailFromPassword
    Config_SaveSetting "Email", "MailCCList", MailCCList
    Config_SaveSetting "Email", "MailStatusMonitor", MailStatusMonitor
    Config_SaveSetting "Program", "DefaultBackupDrive", Prog_DefaultBackup
    
    ' Settings for SSL Encrypted / Remote SMTP emails
    Config_SaveSetting "Email", "MailSMTPPassword", Trim(MailSMTPPassword)
    Config_SaveSetting "Email", "MailSMTPUsername", Trim(MailSMTPUsername)
    Config_SaveSetting "Email", "MailSMTPAuthenticate", Trim(str(CLng(MailSMTPAuthenticate)))
    Config_SaveSetting "Email", "MailSMTPSendUsing", Trim(str(CLng(MailSMTPSendUsing)))
    Config_SaveSetting "Email", "MailUseSSLEncryption", Trim(CStr(MailUseSSLEncryption))
    
    
    'XY Table Settings
    Config_SaveSetting "XYTable", _
                       "UseXYTableAPS", _
                       str$(modConfig.UseXYTableAPS)
    Config_SaveSetting "XYTable", "XYHomeX", str$(XYTablePositions(0, 0))
    Config_SaveSetting "XYTable", "XYHomeY", str$(XYTablePositions(0, 1))
    Dim temp As String
        
    For i = 1 To 100
        temp = "XY" + LTrim$(str(i)) + "X"
        Config_SaveSetting "XYTable", temp, str$(XYTablePositions(i, 0))
        temp = "XY" + LTrim$(str(i)) + "Y"
        Config_SaveSetting "XYTable", temp, str$(XYTablePositions(i, 1))
    Next i
    
    
'------------------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------------------
'    (March 2010 - Isaac Hilburn)
'
'     Modified AF field calibration arrays (see variable declaration up top)
'     so that the arrays can be dynamically resized to any dimension
'
'     Arrays are now N x 2 where N = number of calibration points, with 2 columns
'     at each calibration point for the Ramp Voltage / 2G Counts and for the actual
'     measured field (with Hall probe abd gaussmeter) at that voltage / counts value
'------------------------------------------------------------------------------------------------------------------------


    'AF Field Calibration Arrays
    Config_SaveSetting "AFAxial", "AFAxialCount", Trim(str(AFAxialCount))
    Config_SaveSetting "AFTrans", "AFTransCount", Trim(str(AFTransCount))
    
    If AFAxialCount > 0 Then
        
        'Save the Axial Coil calibration points to the INI file
        For i = 1 To AFAxialCount
            
            Config_SaveSetting "AFAxial", _
                               "AFAxialX" & Format$(i, "0"), _
                               Trim(str(AFAxial(i, 0)))
                                                  
            Config_SaveSetting "AFAxial", _
                               "AFAxialY" & Format$(i, "0"), _
                               Trim(str(AFAxial(i, 1)))
        
        Next i
    
    End If
    
    If AFTransCount > 0 Then
        
        'Save the Transverse Coil calibration points to the INI file
        For i = 1 To AFTransCount
            
            Config_SaveSetting "AFTrans", _
                               "AFTransX" & Format$(i, "0"), _
                               Trim(str(AFTrans(i, 0)))
                                                  
            Config_SaveSetting "AFTrans", _
                               "AFTransY" & Format$(i, "0"), _
                               Trim(str(AFTrans(i, 1)))
            
        Next i
    
    End If
    
    ' IRM Pulse Field Calibration Arrays
    Config_SaveSetting "IRMAxial", _
                        "PulseAxialCount", _
                        Trim(str(PulseAxialCount))
                                         
    Config_SaveSetting "IRMTrans", _
                       "PulseTransCount", _
                       Trim(str(PulseTransCount))
                                         
    If PulseAxialCount > 0 Then
        
        'Save the Low-Field IRM pulse calibration points to the INI file
        For i = 1 To PulseAxialCount
            
            Config_SaveSetting "IRMAxial", _
                               "PulseAxialX" & Format$(i, "0"), _
                               Trim(str(PulseAxial(i, 0)))
                                                  
            Config_SaveSetting "IRMAxial", _
                               "PulseAxialY" & Format$(i, "0"), _
                               Trim(str(PulseAxial(i, 1)))
        
        Next i
        
    End If
    
    If PulseTransCount > 0 Then
        
        'Save the High-Field IRM pulse calibration points to the INI file
        For i = 1 To PulseTransCount
            
            Config_SaveSetting "IRMTrans", _
                               "PulseTransX" & Format$(i, "0"), _
                               Trim(str(PulseTrans(i, 0)))
                                                  
            Config_SaveSetting "IRMTrans", _
                               "PulseTransY" & Format$(i, "0"), _
                               Trim(str(PulseTrans(i, 1)))
        
        Next i
        
    End If
    
'------------------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------------------
    
    'Active Module Settings
    Config_SaveSetting "Modules", "EnableAxialIRM", str$(EnableAxialIRM)
    Config_SaveSetting "Modules", "EnableTransIRM", str$(EnableTransIRM)
    Config_SaveSetting "Modules", "EnableIRMBackfield", str$(EnableIRMBackfield)
    Config_SaveSetting "Modules", "EnableIRMMonitor", str$(EnableIRMMonitor)
    Config_SaveSetting "Modules", "EnableARM", str$(EnableARM)
    Config_SaveSetting "Modules", "EnableAF", str$(EnableAF)
    Config_SaveSetting "Modules", "EnableAFAnalysis", str$(EnableAFAnalysis)
    Config_SaveSetting "Modules", "EnableAltAFMonitor", str$(EnableAltAFMonitor)
    Config_SaveSetting "Modules", "EnableT1", str$(EnableT1)
    Config_SaveSetting "Modules", "EnableT2", str$(EnableT2)
    Config_SaveSetting "Modules", "EnableSusceptibility", str$(EnableSusceptibility)
    Config_SaveSetting "Modules", "EnableVacuum", str$(EnableVacuum)
    Config_SaveSetting "Modules", "EnableDegausserCooler", str$(EnableDegausserCooler)
    
End Sub

'Sub Create_BoardsForINI()
'
' Created: December 10, 2010
'  Author: Isaac Hilburn
'
' Summary: Populates the SystemBoards collection during the INI file upgrade process.  This function
'          is only to be called to populate the default DAQ Board settings for a new, upgraded, version
'          2.4 compatible INI file.
'
Public Sub Create_BoardsForINI()
    
    Dim BoardSectionStr As String
       
    'Check to see if the currently active INI file already has a boards section
    If IniFile.SectionExists("Boards") = True Then Exit Sub
    
    'Set the Board Section String
    BoardSectionStr = DefaultINI.SectionRead(True, False, "Boards")
        
    'Write the INI Section string for the [Boards] Section from the default Boards settings
    IniFile.SectionWrite BoardSectionStr, _
                         "Boards"
           
End Sub

Public Function Create_INIChanStr(ByRef ChanObj As Channel) As String

    Create_INIChanStr = ChanObj.ChanType & "-" & _
                        Trim(str(ChanObj.BoardININum)) & "-CH" & _
                        Trim(str(ChanObj.ChanNum))

End Function

'Sub Create_NewINIFile()
'
' Created: December 10,2010
'  Author: Isaac Hilburn
'
' Summary: For paleomag INI file upgrade ONLY.  Creates a new Paleomag.INI file to be written to
'          with upgraded settings for version 2.4 of the Paleomag code.  Will create the Paleomag_v2-4.INI
'          file in the folder that contains the old INI, pre-version 2.4 Paleomag.INI file.
'
Public Sub Create_NewINIFile()

    Dim NewFilePath As String
    Dim SlashLoc As Integer
    Dim TempStr As String
    Dim fso As FileSystemObject
        
    'If there is a slash at the end of the ProgPath, remove it
    If Right(Prog_INIFile, 1) = "\" Then
    
        Prog_INIFile = Mid(Prog_INIFile, 1, Len(Prog_INIFile) - 1)
            
    End If
            
    SlashLoc = InStrRev(Prog_INIFile, "\")
    
    'Generate the new file's path
    NewFilePath = Mid(Prog_INIFile, 1, SlashLoc) & "Paleomag_v2-4.INI"
    
    'Create the new file
    Set fso = New FileSystemObject
    fso.CreateTextFile NewFilePath, True
    
    'Deallocate the fso object
    Set fso = Nothing
    
    'Allocate the old INI file object
    Set OldIniFile = New CIniFile
    
    'Save the oldinifile path to the OldIniFile Object
    OldIniFile.filename = Prog_INIFile
    
    IniFile.filename = NewFilePath
    Prog_INIFile = NewFilePath
                    
End Sub

Public Sub Create_WavesForINI()
    
    Dim WaveSectionStr As String
       
    'Check to see if the currently active INI file already has a Waves section
    If IniFile.SectionExists("Waves") = True Then Exit Sub
    
    'Set the Board Section String
    WaveSectionStr = DefaultINI.SectionRead(True, False, "WaveForms")
        
    'Write the INI Section string for the [WaveForms] Section from the default WaveForms settings
    IniFile.SectionWrite WaveSectionStr, _
                         "WaveForms"
           
End Sub

'Private Function CreateINIChannelStr
'
' Created:  July, 2010
'  Author:  Isaac Hilburn
'
' Summary:  This function takes in a channel object and it's parent board object
'           and uses the channel type and channel name properties of the channel obj
'           to find the matching Channels collection and Channel in the parent board
'           object.  Then, the function counts the position of the channel in
'           the matching channels collection (zero-indexed), and attaches that to the
'           end of the channel INI str.
'
'   Inputs:
'
'   ChanObj     -   Channel object (must not be Nothing) that the user needs to produce
'                   an INI Channel string key for.
'
'   BoardObj    -   Board Object (can be Nothing), that is the parent board of the Channel
'                   object.  If Board is not inputed, or is Nothing, then the .BoardName
'                   property of the Channel object will be used to find the Parent Board
'                   If the parent board cannot be found, then "ERROR" will be returned
'                   as the INI Channel string
'
'   Output:
'
'   INIChannelStr - String containing the channel type, parent board INI Num, and zero-indexed
'                   channel position in the matching type Channels Collections (i.e. Analog input
'                   channels collection for the parent board).
'                   Format: <2-char channel type string>-<BoardINI#>-CH<zero-indexed channel position>
'                           Note: (The "<",">" characters do not appear in the string
'                   i.e.: AO-0-CH0
'
Private Function CreateINIChannelStr(ByRef ChanObj As Channel, _
                                     Optional ByRef BoardObj As Board = Nothing) As String

    Dim i As Long
    Dim N As Long
    Dim TempChannels As Channels
    
    'Init. TempChannels as Nothing
    Set TempChannels = Nothing
    
    'Check to see if BoardObj is nothing
    If BoardObj Is Nothing Then
    
        'User did not input BoardObj, need to use the ChanObj board name to find
        'the parent board in the System Channels collection
        
        'Turn on Error handling
        On Error Resume Next
        
            Set BoardObj = SystemBoards(ChanObj.BoardName)
            
            'Error check
            If Err.number <> 0 Or BoardObj Is Nothing Then
            
                'Couldn't find matching board, return "ERROR"
                CreateINIChannelStr = "ERROR"
                
                Exit Function
                
            End If
            
        'Turn off error handling
        On Error GoTo 0
        
    End If
    
    'Now have both channel and board objects
    'Need to snag the correct Channels collection in the Board
    Select Case ChanObj.ChanType
    
        Case "AI"
        
            Set TempChannels = BoardObj.AInChannels
            
        Case "AO"
        
            Set TempChannels = BoardObj.AOutChannels
            
        Case "DI"
        
            Set TempChannels = BoardObj.DInChannels
            
        Case "DO"
        
            Set TempChannels = BoardObj.DOutChannels
    
    End Select
    
    'Get the number of channels in the matching Channels collection
    'Turn on Error Handling
    On Error Resume Next
    
        N = TempChannels.Count
        
        'Error check
        If Err.number <> 0 Then
        
            'TempChannels is probably Nothing
            'Return "ERROR"
            CreateINIChannelStr = "ERROR"
            
            Exit Function
            
        End If
        
    'Turn off error handling
    On Error GoTo 0
    
    'Check for N <= 0
    If N <= 0 Then
    
        'TempChannels contains no channels
        'Return "ERROR"
        CreateINIChannelStr = "ERROR"
            
        Exit Function
            
    End If
    
    'Else, there is at least one channel in TempChannels
    'Iterate through TempChannels until a matching channel to the
    'ChanObj is found.  Match on Channel Name
    For i = 1 To N
    
        With ChanObj
    
            If .ChanName = TempChannels(i).ChanName Then
        
                'Use the matching channel position to finish
                'the INI Channel str, return the string,
                'and then exit the function
                CreateINIChannelStr = .ChanType & "-" & _
                                      Trim(str(.BoardININum)) & "-CH" & _
                                      Trim(str(i - 1))
                                      
                Exit Function
                
            End If
            
        End With
        
    Next i
    
    'If the code has reached this far, no matching channel was found
    'Return "ERROR"
    CreateINIChannelStr = "ERROR"

End Function

'Private Sub ErrorCheckBoardsImport(string)
'
' Created: August 13, 2010
'  Author: Isaac Hilburn
'
' Summary: After an error has been flagged within the DAQ boards inport process,
'          this function displays the errors to the user and shows which boards were affected
'
' Inputs:  One String - containing the prime reason for flagging the error
Private Sub ErrorCheckBoardsImport(ByVal IsExplicitErrors As Boolean, _
                                  Optional ByVal ErrorCause As String = vbNullString)
    
    Dim N As Long
    Dim M As Long
    Dim i As Long
    Dim UserResp As Long
    Dim TempL As Long
    Dim BoardININum As Long
    Dim LastININum As Long
    Dim ErrorCounter As Long
            
    Dim ErrorArray As Variant
    Dim BoardErrorArray() As String
    Dim BoardINIError() As Long
    Dim TempStr As String
    Dim ErrorMessage As String
    
    'Get the number of waveforms with error handling
    On Error Resume Next
    
        M = SystemBoards.Count
        
        If Err.number <> 0 Then
        
            'Raise Error, ask the user if they want to continue with all of the DAQ
            'modules disabled
            If AFSystem = "2G" Then
            
                ErrorMsg = "ARM, IRM"
                EnableARM = False
                EnableAxialIRM = False
                EnableTransIRM = False
                modConfig.EnableAltAFMonitor = False
                modConfig.EnableIRMBackfield = False
                modConfig.EnableIRMMonitor = False
                
            Else
            
                ErrorMsg = "AF, ARM, IRM"
                EnableAF = False
                EnableAFAnalysis = False
                EnableARM = False
                EnableAxialIRM = False
                EnableTransIRM = False
                modConfig.EnableAltAFMonitor = False
                modConfig.EnableIRMBackfield = False
                modConfig.EnableIRMMonitor = False
                
            End If
            
            UserResp = frmDialog.DialogBox("DAQ Board settings were not loaded successfullly from the .INI " & _
                                           "settings file." & vbNewLine & vbNewLine & _
                                           "Would you like to continue loading the Paleomag code with the " & _
                                           ErrorMsg & " modules disabled?", _
                                           "INI Settings File Error", _
                                           3, _
                                           "Yes", _
                                           "No", _
                                           "Open the .INI file browser")
                                                                                      
            If UserResp = vbNo Then
            
                'Tell the user the code's about to end
                MsgBox "Paleomag code will end now."
                
                
                
                End
                
            End If
            
            If UserResp = vbCancel Then
            
                'Add in code here for the INI file viewer
                
            End If
                    
            'Else, just Exit the subroutine
            Exit Sub
        
        End If
        
        'No boards loaded into the System Boards Collection, but collection exists
        If M <= 0 Then
        
            'Raise Error, ask the user if they want to continue with all of the DAQ
            'modules disabled
            If AFSystem = "2G" Then
            
                ErrorMsg = "ARM, IRM"
                EnableARM = False
                EnableAxialIRM = False
                EnableTransIRM = False
                modConfig.EnableAltAFMonitor = False
                modConfig.EnableIRMBackfield = False
                modConfig.EnableIRMMonitor = False
                
            Else
            
                ErrorMsg = "AF, ARM, IRM"
                EnableAF = False
                EnableAFAnalysis = False
                EnableARM = False
                EnableAxialIRM = False
                EnableTransIRM = False
                modConfig.EnableAltAFMonitor = False
                modConfig.EnableIRMBackfield = False
                modConfig.EnableIRMMonitor = False
                
            End If
            
            UserResp = frmDialog.DialogBox("DAQ Board settings were not loaded successfullly from the .INI " & _
                                           "settings file." & vbNewLine & vbNewLine & _
                                           "Would you like to continue loading the Paleomag code with the " & _
                                           ErrorMsg & " modules disabled?", _
                                           "INI Settings File Error", _
                                           3, _
                                           "Yes", _
                                           "No", _
                                           "Open the .INI file browser")
                                                                                      
            If UserResp = vbNo Then
            
                'Tell the user the code's about to end
                MsgBox "Paleomag code will end now."
                
                
                
                End
                
            End If
            
            If UserResp = vbCancel Then
            
                'Add in code here for the INI file viewer
                
            End If
                    
            'Else, just Exit the subroutine
            Exit Sub
        
        End If
        
    On Error GoTo 0
        
    'If there are no explicit errors, then exit the function
    If IsExplicitErrors = False Or _
       ErrorCause <> vbNullString _
    Then
    
        Exit Sub
        
    End If

    'Clip off additional comma at the end of each error statement
    If Right(ErrorCause, 1) = "," Then

        ErrorCause = Mid(ErrorCause, 1, Len(ErrorCause) - 1)
        
    End If
            
    'Count the number of individual error sub-strings stored in error cause
    ErrorArray = Split(ErrorCause, ",")
    
    N = UBound(ErrorArray) + 1
    
    'Redimension the BoardErrorArray
    ReDim BoardErrorArray(N, 2)
        
    'Initialize last INI num recorded to -1
    LastININum = -1
    
    'Set the dimensions of the BoardINIError array to 1
    ReDim BoardINIError(1)
        
    'Figure out which Wave INI numbers are affected by errors
    For i = 0 To N - 1
    
        'Parse out the Wave object INI num attached to the error
        TempL = InStr(1, ErrorArray(i), ";")
        BoardININum = val(Mid(ErrorArray(i), TempL + 1))
        
        'Parse out the error string
        TempStr = Mid(ErrorArray(i), 1, TempL - 1)
        
        BoardErrorArray(i, 0) = Trim(str(BoardININum))
        BoardErrorArray(i, 1) = TempStr
        
        'Check to see if this is the first pass through the loop
        If LastININum = -1 Then
        
            'Store the new ini number
            BoardINIError(0) = BoardININum
            
            'Set the last INI number to the current number
            LastININum = BoardININum
            
        End If
        
        'Check to see, now, if the wave ini number has changed
        If LastININum <> BoardININum Then
        
            'Need to resize the BoardINIError array
            TempL = UBound(BoardINIError)
                           
            'Redimension preserver
            ReDim Preserve BoardINIError(TempL + 1)
            
            'Store the new wave ini number
            BoardINIError(TempL) = BoardININum
            
            'Set the last INI number to the current number
            LastININum = BoardININum
            
        End If
                    
    Next i
    
    'Need to start constructing the Board Error Message
    ErrorMessage = "One or more errors occurred while loading in the System DAQ Board settings from the " & _
                   ".INI settings file. Some of the DAQ Channels needed for doing AF / Rockmag may not have " & _
                   "loaded correctly." & vbNewLine & vbNewLine
                   
    If UBound(BoardINIError) = 1 Then
    
        ErrorMessage = ErrorMessage & "1 Board "
        
    ElseIf UBound(BoardINIError) > 1 Then
    
        ErrorMessage = ErrorMessage & Trim(str(UBound(BoardINIError))) & " Boards "
        
    End If
    
    ErrorMessage = ErrorMessage & "were affected." & vbNewLine & vbNewLine & _
                   "Please click ""Ok"" to continue or ""Abort"" to end the code now." & _
                   vbNewLine & vbNewLine & _
                   "------------------------" & vbNewLine & _
                   "List of Board Import Errors:" & vbNewLine
                   
    'Need to iterate through the rows and cols of the BoardErrorArray listing the errors by board
    N = UBound(BoardErrorArray, 1)
    
    'Start Error Counter out at Zero
    ErrorCounter = 0

    If N > 0 Then
    
        'Set Last INI Num to -1
        LastININum = -1
                    
        For i = 0 To N - 1
        
            'Check for next Wave objects .INI number
            If LastININum <> val(BoardErrorArray(i, 0)) Then
            
                ErrorMessage = ErrorMessage & vbNewLine & _
                               "Board INI#(" & BoardErrorArray(i, 0) & ")" & vbNewLine
                                              
            End If
            
            'Now increment the error counter
            ErrorCounter = ErrorCounter + 1
            
            ErrorMessage = ErrorMessage & _
                           PadLeft(Trim(str(ErrorCounter)) & ".", 4) & " " & _
                           BoardErrorArray(i, 1) & vbNewLine
                           
        Next i
        
    End If
          
    'Need to send message to the user about the Board import errors
    UserResp = _
        frmDialog.DialogBox(ErrorMessage, _
                            "INI Settings Error(s)", _
                            3, _
                            "OK", _
                            "Abort", _
                            "Open the .INI" & vbNewLine & "file browser")
    
    If UserResp = vbNo Then
    
        'Send message that the code will end now
        MsgBox "The Paleomag code will now end."
    
        'User selected to end the code, so do so
        
        
        End
        
    End If
    
    If UserResp = vbCancel Then
    
        'Add in code to open the INI file browser
        
    End If
                              
End Sub

'Private Sub ErrorCheckBoardsImport(string)
'
' Created: August 13, 2010
'  Author: Isaac Hilburn
'
' Summary: After an error has been flagged within the DAQ boards inport process,
'          this function displays the errors to the user and shows which boards were affected
'
' Inputs:  One String - containing the prime reason for flagging the error
Private Sub ErrorCheckChannelsImport( _
    ByVal IsExplicitErrors As Boolean, _
    Optional ByVal ErrorCause As String = vbNullString)
    
    Dim N As Long
    Dim M As Long
    Dim i As Long
    Dim UserResp As Long
    Dim TempL As Long
    Dim CurChanDesc As String
    Dim LastChanDesc As String
    Dim ErrorCounter As Long
            
    Dim ErrorArray As Variant
    Dim ChanErrorArray() As String
    Dim TempStr As String
    Dim ErrorMessage As String
            
    'If there are no explicit errors, then exit the function
    If IsExplicitErrors = False Or _
       ErrorCause = vbNullString _
    Then Exit Sub

       
    'Clip off additional comma at the end of each error statement
    If Right(ErrorCause, 1) = "," Then

        ErrorCause = Mid(ErrorCause, 1, Len(ErrorCause) - 1)
        
    End If
            
    'Count the number of individual error sub-strings stored in error cause
    ErrorArray = Split(ErrorCause, ",")
    
    N = UBound(ErrorArray) + 1
    
    'Redimension the ChanErrorArray
    ReDim ChanErrorArray(N, 2)
        
    'Figure out which Wave INI numbers are affected by errors
    For i = 0 To N - 1
    
        'Parse out the Wave object INI num attached to the error
        TempL = InStr(1, ErrorArray(i), ";")
        
        If TempL > 0 Then
        
            CurChanDesc = Mid(ErrorArray(i), TempL + 1)
            
            'Parse out the error string
            TempStr = Mid(ErrorArray(i), 1, TempL - 1)
            
        Else
        
            CurChanDesc = vbNullString
            
            'Parse out the error string
            TempStr = ErrorArray(i)
                        
        End If
        
        ChanErrorArray(i, 0) = CurChanDesc
        ChanErrorArray(i, 1) = TempStr
                           
    Next i
    
    'Prune the last ", " off of the disabled modules string
    If Right(DisabledModules, 2) = ", " Then
    
        DisabledModules = Mid(DisabledModules, 1, Len(DisabledModules) - 2)
    
    End If
    
    'Need to start constructing the Channel imports Error Message
    ErrorMessage = "One or more errors occurred while loading in the System DAQ Channel settings from the " & _
                   ".INI settings file." & vbNewLine & vbNewLine & _
                   "The following modules have been disabled:" & vbNewLine & _
                   DisabledModules & vbNewLine & vbNewLine & _
                   "Would you like to continue loading the Paleomag code with these modules turned off?" & _
                   vbNewLine & vbNewLine & "----------------------------" & vbNewLine & _
                   "List of Channel Import Errors:"
                   
    'Need to iterate through the rows and cols of the ChanErrorArray listing the errors by board
    N = UBound(ChanErrorArray, 1)
    
    'Start Error Counter out at Zero
    ErrorCounter = 0

    If N > 0 Then
    
        'Set Last INI Num to -1
        LastChanDesc = "-1"
                    
        For i = 0 To N - 1
        
            'Check for next Wave objects .INI number
            If LastChanDesc <> ChanErrorArray(i, 0) Then
            
                If ChanErrorArray(i, 0) = vbNullString Then
                
                    ErrorMessage = ErrorMessage & vbNewLine & _
                                   "Missing Channels:" & _
                                   vbNewLine
                                
                Else
                
                    ErrorMessage = ErrorMessage & vbNewLine & _
                                   ChanErrorArray(i, 0) & " Channel:" & _
                                   vbNewLine
                                   
                End If
                
                
                                              
            End If
            
            'Now increment the error counter
            ErrorCounter = ErrorCounter + 1
            
            ErrorMessage = ErrorMessage & _
                           PadLeft(Trim(str(ErrorCounter)) & ".", 4) & " " & _
                           ChanErrorArray(i, 1) & vbNewLine
                           
        Next i
        
    End If
          
    'Need to send message to the user about the Board import errors
    UserResp = _
        frmDialog.DialogBox(ErrorMessage, _
                            "INI Settings Error(s)", _
                            3, _
                            "Yes", _
                            "No", _
                            "Open the .INI" & vbNewLine & "file browser")
    
    If UserResp = vbNo Then
    
        'Send message that the code will end now
        MsgBox "The Paleomag code will now end."
    
        'User selected to end the code, so do so
        
        
        End
        
    End If
    
    If UserResp = vbCancel Then
    
        'Add in code to open the INI file browser
        
    End If
                              
End Sub

'Private Sub ErrorCheckWavesImport(string)
'
' Created: August 13, 2010
'  Author: Isaac Hilburn
'
' Summary: After an error has been flagged with the WaveForms input,
'          this function scans the WaveForms system collection and checks
'          to see which wave forms were impacted by the error.
'          The function then notifies the user and asks them if
'          they would like to continue with limited functionality
'          - i.e. no AFs, no IRM Monitor, no Alternate AF monitor
'          - or if they would like to browse their .INI files or end the code
'
' Inputs:  One String - containing the prime reason for flagging the error
Private Sub ErrorCheckWavesImport(ByVal IsExplicitErrors As Boolean, _
                                  Optional ByVal ErrorCause As String = vbNullString)
                                  
    Dim isAFMonitorOK As Boolean
    Dim isAFRampUpOK As Boolean
    Dim isAFRampDownOK As Boolean
    Dim isAltAFMonitorOK As Boolean
    Dim isIRMMonitorOK As Boolean
    
    Dim N As Long
    Dim M As Long
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim Xdim As Long
    Dim Ydim As Long
    Dim UserResp As Long
    Dim TempL As Long
    Dim WaveDesc As String
    Dim LastWaveDesc As String
    Dim ErrorCounter As Long
            
    Dim ErrorArray As Variant
    Dim WaveErrorArray() As String
    Dim DistinctWaveError() As String
    Dim TempStr As String
    Dim DisabledModulesStr As String
    Dim ErrorMessage As String
    
    'Set all the load error status flags to false
    isAFMonitorOK = True
    isAFRampUpOK = True
    isAFRampDownOK = True
    isAltAFMonitorOK = True
    isIRMMonitorOK = True
    
    'Get the number of waveforms with error handling
    On Error Resume Next
    
        M = WaveForms.Count
        
        'error check
        If Err.number <> 0 Then
        
            'Wha-oh, no waveforms collection
            UserResp = frmDialog.DialogBox( _
                                    "The AF ramp & monitor and the IRM monitor waveform " & _
                                    "settings are missing from the .INI file." & vbNewLine & _
                                    vbNewLine & _
                                    "Do you wish to continue with the code, but with the AF modules " & _
                                    "and the IRM monitor module turned off?", _
                                    "Settings Load Error", _
                                    3, _
                                    "Yes", _
                                    "No", _
                                    "Open .INI File Browser")
                                    
            If UserResp = vbYes Then
            
                'User wants to proceed with the code but with the AF and IRM modules turned off
                modConfig.EnableAF = False
                modConfig.EnableAFAnalysis = False
                modConfig.EnableAltAFMonitor = False
                modConfig.EnableIRMMonitor = False
                
                'Exit the subroutine
                Exit Sub
                
            ElseIf UserResp = vbNo Then
            
                'User wants to have the code to end
                MsgBox "The Paleomag code will exit now."
                
                
                
                End
                
            End If
            
        End If
        
    On Error GoTo 0
    
    'Set the isExtraWaves status flag to false
    isExtraWaves = False
    
    
    'Remove the last comma from ErrorCause
    If IsExplicitErrors = True And _
       ErrorCause <> vbNullString _
    Then
    
        If Right(ErrorCause, 1) = "," Then
    
            ErrorCause = Mid(ErrorCause, 1, Len(ErrorCause) - 1)
            
        End If
                
        'Count the number of itemized error substrings stored in error cause
        ErrorArray = Split(ErrorCause, ",")
        
        N = UBound(ErrorArray) + 1
        
        'Redimension the WaveErrorArray
        ReDim WaveErrorArray(N, 2)
            
        'Initialize last INI num recorded to -1
        LastWaveDesc = "-1"
        
        'Set the dimensions of the DistinctWaveError array to 1
        ReDim DistinctWaveError(1)
            
        'Figure out which Wave INI numbers are affected by errors
        For i = 0 To N - 1
        
            'Parse out the Wave object INI num attached to the error
            TempL = InStr(1, ErrorArray(i), ";")
            WaveDesc = Mid(ErrorArray(i), TempL + 1)
            
            'Parse out the error string
            TempStr = Mid(ErrorArray(i), 1, TempL - 1)
            
            WaveErrorArray(i, 0) = WaveDesc
            WaveErrorArray(i, 1) = TempStr
            
            
            'Check to see if this is the first pass through the loop
            If LastWaveDesc = "-1" Then
            
                'Store the new ini number
                DistinctWaveError(0) = WaveDesc
                
                'Set the last INI number to the current number
                LastWaveDesc = WaveDesc
                
            End If
            
            'Check to see, now, if the wave ini number has changed
            If LastWaveDesc <> WaveDesc Then
            
                'Need to resize the DistinctWaveError array
                TempL = UBound(DistinctWaveError)
                               
                'Redimension preserver
                ReDim Preserve DistinctWaveError(TempL + 1)
                
                'Store the new wave ini number
                DistinctWaveError(TempL) = WaveDesc
                
                'Set the last INI number to the current number
                LastWaveDesc = WaveDesc
                
            End If
                        
        Next i
        
                
        For i = 0 To UBound(DistinctWaveError) - 1
            
            'Now, find which wave in WaveForms has a matching .INI number
            'and figure out if that waveform matches one of the needed system
            'waves
            For j = 1 To M
            
                With WaveForms(j)
                
                    If .WaveName = DistinctWaveError(i) Then
                    
                        'This wave object has errors associated with it
                        'Based on the wave name, set the "OK" status flag for that
                        'needed function to false
                        Select Case .WaveName
                        
                            Case "AFRAMPUP"
                            
                                isAFRampUpOK = False
                                
                            Case "AFRAMPDOWN"
                            
                                isAFRampDownOK = False
                                
                            Case "AFMONITOR"
                        
                                isAFMonitorOK = False
                                
                            Case "ALTAFMONITOR"
                            
                                isAltAFMonitorOK = False
                                
                            Case "IRMMONITOR"
                            
                                isIRMMonitorOK = False
                                
                        End Select
                        
                        'Exit the for loop
                        Exit For
                        
                    End If
                            
                End With
                
            Next j
            
        Next i
                            
    End If
                            
    'Now check to see if any wave is just flat out missing, regardless of whether or not
    'an error happened during the load
    '(i.e. - a wave not present at all in the INI file nor included in the INI wave count)
    
    'Set the wave missing status flags all to true
    isAFMonitorMissing = True
    isAFRampUpMissing = True
    isAFRampDownMissing = True
    isAltAFMonitorMissing = True
    isIRMMonitorMissing = True
    isExtraWavesMissing = True
    
    'Loop through the waves in the wave-forms system collection and check off each of the five
    'necessary waves.  Check for extra waves.
    For i = 1 To M
    
        With WaveForms(i)
        
            Select Case .WaveName
            
                Case "AFMONITOR"
                
                    isAFMonitorMissing = False
                    
                Case "AFRAMPUP"
                
                    isAFRampUpMissing = False
                    
                Case "AFRAMPDOWN"
                
                    isAFRampDownMissing = False
                    
                Case "ALTAFMONITOR"
                
                    isAltAFMonitorMissing = False
                    
                Case "IRMMONITOR"
                
                    isIRMMonitorMissing = False
                    
            End Select
            
        End With
        
    Next i
        
    'Start missing wave string out as empty
    MissingWaveStr = vbNullString
    
    'Start disabled modules string out as empty
    DisabledModulesStr = vbNullString
    
    'Set Error Counter at Zero
    ErrorCounter = 0
    
    'Now check through the boolean flags to see what error messages need to be sent
    If isAFMonitorMissing = True Then
        
        'Count this error
        ErrorCounter = ErrorCounter + 1
        
        'Add the error number to the start of the new line in the missing wave string
        MissingWaveStr = PadLeft(Trim(str(ErrorCounter)) & ".", 4) & " " & _
                         "ADWIN AF Monitor settings missing" & vbNewLine
        
        If AFSystem = "ADWIN" Then
        
            EnableAF = False
            EnableAFAnalysis = False
            EnableARM = False
            
        End If
        
    ElseIf isAFMonitorOK = False Then
    
        If AFSystem = "ADWIN" Then
        
            EnableAF = False
            EnableAFAnalysis = False
            EnableARM = False
            
        End If
    
    End If
    
    If isAFRampUpMissing = True Then
    
        'Count this error
        ErrorCounter = ErrorCounter + 1
        
        'Add the error number to the start of the new line in the missing wave string
        MissingWaveStr = MissingWaveStr & _
                         PadLeft(Trim(str(ErrorCounter)) & ".", 4) & " " & _
                         "ADWIN AF Ramp Up settings missing" & vbNewLine
                         
        If AFSystem = "ADWIN" Then
        
            EnableAF = False
            EnableAFAnalysis = False
            EnableARM = False
            
        End If
        
    ElseIf isAFRampUpOK = False Then
    
        If AFSystem = "ADWIN" Then
        
            EnableAF = False
            EnableAFAnalysis = False
            EnableARM = False
            
        End If
        
    End If
    
    If isAFRampDownMissing = True Then
    
        'Count this error
        ErrorCounter = ErrorCounter + 1
        
        'Add the error number to the start of the new line in the missing wave string
        MissingWaveStr = MissingWaveStr & _
                         PadLeft(Trim(str(ErrorCounter)) & ".", 4) & " " & _
                         "ADWIN AF Ramp Down settings missing" & vbNewLine
                         
        If AFSystem = "ADWIN" Then
        
            EnableAF = False
            EnableAFAnalysis = False
            EnableARM = False
            
        End If
        
    ElseIf isAFRampDownOK = False Then
    
        If AFSystem = "ADWIN" Then
        
            EnableAF = False
            EnableAFAnalysis = False
            EnableARM = False
            
        End If
        
    End If
        
    If isAltAFMonitorMissing = True Then
    
        'Count this error
        ErrorCounter = ErrorCounter + 1
        
        'Add the error number to the start of the new line in the missing wave string
        MissingWaveStr = MissingWaveStr & _
                         PadLeft(Trim(str(ErrorCounter)) & ".", 4) & " " & _
                         "Alternate AF Monitor (2G) settings missing" & vbNewLine
                         
        EnableAltAFMonitor = False
        
    ElseIf isAltAFMonitorOK = False Then
    
        EnableAltAFMonitor = False
        
    End If
        
    If isIRMMonitorMissing = True Then
            
        'Count this error
        ErrorCounter = ErrorCounter + 1
        
        'Add the error number to the start of the new line in the missing wave string
        MissingWaveStr = MissingWaveStr & _
                         PadLeft(Trim(str(ErrorCounter)) & ".", 4) & " " & _
                         "IRM Monitor settings missing" & vbNewLine
                         
        EnableIRMMonitor = False
        
    ElseIf isIRMMonitorOK = False Then
    
        EnableIRMMonitor = False
        
    End If
    
    'Now need to create the string to determine which modules have been disabled
    If EnableAF = False Then
    
        'Assume that all the ADWIN AF modules are disabled and the ARM module is disabled as well
        DisabledModulesStr = "AF, ARM, "
        
    End If
    
    If EnableIRMMonitor = False Then
    
        DisabledModulesStr = DisabledModulesStr & "IRM Monitor, "
        
    End If
    
    If EnableAltAFMonitor = False Then
    
        DisabledModulesStr = DisabledModulesStr & "Alternate AF (2G) Monitor"
        
    End If
    
    'Clip off the last ", " from the disabled modules str
    If Right(DisabledModulesStr, 2) = ", " Then
    
        DisabledModulesStr = Mid(DisabledModulesStr, _
                                 1, _
                                 Len(DisabledModulesStr) - 2)
    
    End If
    
    'Does an error message need to be raised?
    If IsExplicitErrors = True Or _
       MissingWaveStr <> vbNullString _
    Then
        
        'Need to create the error message
        ErrorMessage = "One or more errors occurred during the loading of the AF/IRM wave-form settings from the " & _
                       ".INI settings file.  Please check the .INI file and if this problem persists, restore an " & _
                       "older version of the file." & _
                       vbNewLine & vbNewLine & _
                       "The following modules have been disabled:" & vbNewLine & _
                       DisabledModulesStr & vbNewLine & vbNewLine & _
                       "Would you like to continue the code without these modules? Click 'No' to end the code right now." & _
                       vbNewLine & vbNewLine & _
                       "-------------------------------" & vbNewLine & _
                       "List of Wave Form Import Errors:" & vbNewLine
                       
        If MissingWaveStr <> vbNullString Then
        
            ErrorMessage = ErrorMessage & MissingWaveStr & vbNewLine
            
        End If
        
        'Are there any explicit errors to list
        If ErrorCause <> vbNullString Then
        
            'Need to iterate through the rows and cols of the WaveErrorArray listing the errors by board
            N = UBound(WaveErrorArray, 1)
        
            If N > 0 Then
            
                'Set Last INI Num to -1
                LastWaveDesc = -1
                            
                For i = 0 To N - 1
                
                    'Check for next Wave objects .INI number
                    If LastWaveDesc <> val(WaveErrorArray(i, 0)) Then
                    
                        ErrorMessage = ErrorMessage & vbNewLine & _
                                       WaveErrorArray(i, 0) & " Wave Form" & _
                                       vbNewLine
                                                      
                    End If
                    
                    'Now increment the error counter
                    ErrorCounter = ErrorCounter + 1
                    
                    ErrorMessage = ErrorMessage & _
                                   PadLeft(Trim(str(ErrorCounter)) & ".", 4) & " " & _
                                   WaveErrorArray(i, 1) & vbNewLine
                                   
                Next i
                
            End If
                                   
        End If
              
        'Need to send message to the user that there's a missing wave object that's not present
        'in the .INI file
        UserResp = _
            frmDialog.DialogBox(ErrorMessage, _
                                "INI Settings Error(s)", _
                                3, _
                                "Yes", _
                                "No", _
                                "Open the .INI" & vbNewLine & "file browser")
        
        If UserResp = vbNo Then
        
            'Send message that the code will end now
            MsgBox "The Paleomag code will now end."
        
            'User selected to end the code, so do so
            
            
            End
            
        End If
        
        If UserResp = vbCancel Then
        
            'Add in code to open the INI file browser
            
        End If
                                  
    End If
                              
End Sub

' Sub Get_BoardsFromIni()
'
' Created: March 30, 2010
'  Author: Isaac Hilburn
'
' Summary: Reads Paleomag.ini file and parses the [Boards] section of the file using
'          the Config_GetFromIni function.  Creates new board objects in SystemBoards collection
'          using the parameters specified for each board in the .ini file.

Public Sub Get_BoardsFromIni()

    Dim TempBoard As Board
    Dim TempChannels As Channels
    
    Dim NumBoards As Long
    Dim NumChannels As Long
    Dim i As Long
    Dim j As Long
    
    Dim DIO_isConfigured As Boolean
    Dim isBoardLoadError As Boolean
    Dim ErrorMsg As String
    Dim BoardLoadErrorCause As String
    Dim TempBoardLoadError As Boolean
    
    Dim TempStr As String
    Dim BoardNameStr As String
    Dim ChanNameStr As String
    
    'Initialize DIO_isConfigured to False
    DIO_isConfigured = False
    
    'Reset and Initialize system boards collection
    Set SystemBoards = Nothing
    Set SystemBoards = New Boards
    
    Set TempChannels = Nothing
    Set TempBoard = Nothing
    
    Set TempChannels = New Channels
    Set TempBoard = New Board
    
    'First, get the number of Boards in the .ini file
    NumBoards = val(Config_GetFromINI("Boards", "BoardsCount", "-1", Prog_INIFile))
    
    'Set the Board Load Error status flags to defaults (False & vbnullstring)
    isBoardLoadError = False
    
    'Set the error cause string = vbnullstring
    BoardLoadErrorCause = vbNullString
    
    'Check to see if the BoardsCount field exists
    If NumBoards = -1 Or NumBoards = 0 Then
    
        'If NumBoards = -1, the BoardsCount field does not exist
        'This means the entire [Boards] section of the .ini file is missing
        
        'If NumBoards = 0, then the BoardsCount field does exist,
        'But there are no boards loaded.
        
        'Set NoINIBoards status flag = True
        NoINIBoards = True
        
        'Raise Error, ask the user if they want to continue with all of the DAQ
        'modules disabled
        If AFSystem = "2G" Then
        
            ErrorMsg = "ARM, IRM"
            EnableARM = False
            EnableAxialIRM = False
            EnableTransIRM = False
            modConfig.EnableAltAFMonitor = False
            modConfig.EnableIRMBackfield = False
            modConfig.EnableIRMMonitor = False
            
        Else
        
            ErrorMsg = "AF, ARM, IRM"
            EnableAF = False
            EnableAFAnalysis = False
            EnableARM = False
            EnableAxialIRM = False
            EnableTransIRM = False
            modConfig.EnableAltAFMonitor = False
            modConfig.EnableIRMBackfield = False
            modConfig.EnableIRMMonitor = False
            
        End If
        
        UserResp = frmDialog.DialogBox("DAQ Board settings are missing from the .INI settings file " & _
                                       "or the Board settings are corrupted." & vbNewLine & vbNewLine & _
                                       "Would you like to continue loading the Paleomag code with the " & _
                                       ErrorMsg & " modules disabled?", _
                                       "INI Settings File Error", _
                                       3, _
                                       "Yes", _
                                       "No", _
                                       "Open the .INI file browser")
                                       
                                       
        If UserResp = vbNo Then
        
            'Tell the user the code's about to end
            MsgBox "Paleomag code will end now."
            
            
            
            End
            
        End If
        
        If UserResp = vbCancel Then
        
            'Add in code here for the INI file viewer
            
        End If
                
        'Else, just Exit the subroutine
        Exit Sub
    
    End If
    
    'Now iterate through each of the Boards listed in the Boards section
    'and grab the necessary information
    For i = 1 To NumBoards
    
        'Store the Board Name
        BoardNameStr = Trim(Config_GetFromINI( _
                                "Boards", _
                                "BoardName" & Format$(i - 1, "0"), _
                                "ERROR", _
                                Prog_INIFile))
    
        If BoardNameStr = "ERROR" Then
        
            'There is something wrong with this Board's block of settings in the .INI file
            'Set the isBoardLoad error
            isBoardLoadError = True
            
            'Save that the BoardName property for this INI Board num is missing
            BoardLoadErrorCause = BoardLoadErrorCause & "Board Name missing;" & Format$(i - 1, "0") & ","
            
        End If
    
        'If there aren't enough boards in the System Boards collection, add a new one
        If i > SystemBoards.Count Then
            
            'Need to add a new system board
            'Key for board = Board Name
            SystemBoards.Add , BoardNameStr
                
        End If
        
        'With the current, new board in the System Boards collection
        With SystemBoards(BoardNameStr)
        
            'Get the Boards INI number key
            .BoardININum = val(Config_GetFromINI("Boards", _
                                                 "BoardININum" & Format$(i - 1, "0"), _
                                                 Trim(str(i - 1)), _
                                                 Prog_INIFile))
        
            'Get the Board's Name from the .ini file
            .BoardName = BoardNameStr
                                           
            'Get the Board's IO mode - single (= 1) or differential (= 0) mode
            .BoardMode = val(Config_GetFromINI("Boards", _
                                               "BoardMode" & Format$(i - 1, "0"), _
                                               "-1", _
                                               Prog_INIFile))
                                           
            If .BoardMode = -1 Then
            
                isBoardLoadError = True
                BoardLoadErrorCause = BoardLoadErrorCause & "Board Mode missing;" & Format$(i - 1, "0") & ","
                
            End If
                                           
            'Get the Device number assigned to the board by the MS Windows operating system
            .BoardNum = val(Config_GetFromINI("Boards", _
                                              "BoardNum" & Format$(i - 1, "0"), _
                                              "-1", _
                                              Prog_INIFile))
            
            If .BoardNum = -1 Then
            
                isBoardLoadError = True
                BoardLoadErrorCause = BoardLoadErrorCause & "Board Device # missing;" & Format(i - 1, "0") & ","
                
            End If
            
            'Set the Board's Range = a new Range object
            Set .range = New range
            
            'If no RangeType is specified for the board, then
            'the default value of -1 will be caught and the Max & Min
            'Range Values will be extracted manually
            '(The values are normally set by setting the Range Type)
            'The ADWIN board does not support a RangeType
            .range.RangeType = val(Config_GetFromINI( _
                                            "Boards", _
                                            "RangeType" & Format$(i - 1, "0"), _
                                            "-1", _
                                            Prog_INIFile))
            
            'If RangeType has default (no rangetype) value = -1, then
            'Need to read in the Max & Min range values from the INI file
            '(This will be true for the ADWIN-light-16 board)
            If .range.RangeType = -1 Then
            
                .range.MaxValue = val(Config_GetFromINI( _
                                                "Boards", _
                                                "BoardRangeMax" & Format$(i - 1, "0"), _
                                                "10", _
                                                Prog_INIFile))
                                           
                .range.MinValue = val(Config_GetFromINI( _
                                                "Boards", _
                                                "BoardMin" & Format$(i - 1, "0"), _
                                                "-10", _
                                                Prog_INIFile))
                                           
            End If
            
            'This field specifies what type of Board this is
            '1 = MCC_UL, Measurement Computing Board (most likely a PCI-DAS6030)
            '2 = ADWIN, ADWIN-light-16 board
            '3 = OTHER, Some other type off board, currently not supported explicitly by the code
            .CommProtocol = val(Config_GetFromINI( _
                                    "Boards", _
                                    "CommProtocol" & Format$(i - 1, "0"), _
                                    "-1", _
                                    Prog_INIFile))
                                           
            If .CommProtocol = -1 Then
            
                isBoardLoadError = True
                BoardLoadErrorCause = BoardLoadErrorCause & "Comm Protocol missing;" & _
                                                            Format(i - 1, "0") & ","
                
            End If
                                           
            'This field specifies a string letting other portions of the code
            'know what this board can and cannot be used to do
            .BoardFunction = Config_GetFromINI("Boards", _
                                               "BoardFunction" & Format$(i - 1, "0"), _
                                               "ERROR", _
                                               Prog_INIFile)
                                               
            'Don't really care if there's an error loading this field, it's just for error messages
            
            'Maximum Analog Input Rate supported for this board in Hz
            .MaxAInRate = val(Config_GetFromINI("Boards", _
                                                "MaxAInRate" & Format$(i - 1, "0"), _
                                                "50000", _
                                                Prog_INIFile))
                                           
            
            'Maximum Analog Output Rate supported for this board in Hz
            .MaxAOutRate = val(Config_GetFromINI("Boards", _
                                                 "MaxAOutRate" & Format$(i - 1, "0"), _
                                                 "50000", _
                                                 Prog_INIFile))
                                           
            'Now need to load Board channel configuration from the .ini file
                
            'Get Number of Analog Input Channels
            NumChannels = val(Config_GetFromINI("Boards", _
                                                "AInChannelsCount" & Format$(i - 1, "0"), _
                                                "-1", _
                                                Prog_INIFile))
                                           
            'Screen for Import error (Default, NumChannels = -1)
            If NumChannels < 0 Then
            
                'The analog input channel settings for this board are missing!
                isBoardLoadError = True
                BoardLoadErrorCause = BoardLoadErrorCause & _
                                        "Analog Input Channel Count missing;" & Format(i - 1, "0") & ","
            
            ElseIf NumChannels > 0 Then
            
                'Add the necessary number of channels
                'to the Analog Input channels collection on this Board
                For j = 1 To NumChannels
            
                    'Channel Name Str
                    TempStr = Trim(Config_GetFromINI( _
                                            "Boards", _
                                            "AI-" & _
                                            Format(SystemBoards(BoardNameStr).BoardININum, _
                                                   "0") & "-" & _
                                            "CH" & Format(j - 1, "0"), _
                                            "ERROR,-1", _
                                            Prog_INIFile))
                                            
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++'

'       Need to add code to parse channamestr - it currently contains both
'       the channel name and the channel number assignment
                
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++'

                    ChanNameStr = Mid(TempStr, _
                                      1, _
                                      InStr(1, TempStr, ",") - 1)
                
                    'Use Channel name as the key in the channels collection
                    .AInChannels.Add , ChanNameStr
                    
                    With .AInChannels(ChanNameStr)
                    
                        'Set the Channel's parent Board Name
                        .BoardName = BoardNameStr
                    
                        'Set the Channel's parent Board #
                        .BoardININum = SystemBoards(.BoardName).BoardININum
                    
                        'Note: that the format of the key string for the channel
                        '      contains both an element indicating the board that
                        '      the channel is on, and also indicating the channel
                        '      number itself.
                        
                        'Snatch the Channel Name from the INI file
                        .ChanName = ChanNameStr
                                                                                  
                        'Snatch the Channel Number from the INI File
                        .ChanNum = val(Mid(TempStr, _
                                           InStr(1, TempStr, ",") + 1))
                                                      
                        'Set Channel Type
                        .ChanType = "AI"
                        
                        'Set Channel description to an empty string
                        Set .ChanDescs = New ChannelDescs
                                                      
                    End With
                    
                Next j
                                                      
            End If
                        
            'Get Number of Analog Output Channels
            NumChannels = val(Config_GetFromINI("Boards", _
                                           "AOutChannelsCount" & Format$(i - 1, "0"), _
                                           "-1", _
                                           Prog_INIFile))
                                           
            'Screen for Import error (Default, NumChannels = -1)
            If NumChannels < 0 Then
            
                'The analog output channel settings for this board are missing!
                isBoardLoadError = True
                BoardLoadErrorCause = BoardLoadErrorCause & _
                                        "Analog Output Channel Count missing;" & Format(i - 1, "0") & ","
            
            ElseIf NumChannels > 0 Then
            
                'Add the necessary number of channels
                'to the Analog Output channels collection on this Board
                For j = 1 To NumChannels
            
                    'Channel Name Str
                    TempStr = Trim(Config_GetFromINI( _
                                            "Boards", _
                                            "AO-" & _
                                            Format(SystemBoards(BoardNameStr).BoardININum, _
                                                   "0") & "-" & _
                                            "CH" & Format(j - 1, "0"), _
                                            "ERROR,-1", _
                                            Prog_INIFile))
                                                
                    
                                            
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++'

'       Need to add code to parse channamestr - it currently contains both
'       the channel name and the channel number assignment
                
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++'

                    ChanNameStr = Mid(TempStr, _
                                      1, _
                                      InStr(1, TempStr, ",") - 1)


                    'Use Channel name as the key in the channels collection
                    .AOutChannels.Add , ChanNameStr
                    
                    With .AOutChannels(ChanNameStr)
                                                
                        'Set the Channel's parent Board Name
                        .BoardName = BoardNameStr
                    
                        'Set the Channel's parent Board #
                        .BoardININum = SystemBoards(.BoardName).BoardININum
                    
                        'Note: that the format of the key string for the channel
                        '      contains both an element indicating the board that
                        '      the channel is on, and also indicating the channel
                        '      number itself.
                        
                        'Snatch the Channel Name from the INI file
                        .ChanName = ChanNameStr
                                                                                  
                        'Snatch the Channel Number from the INI File
                        .ChanNum = val(Mid(TempStr, _
                                           InStr(1, TempStr, ",") + 1))
                                                      
                        'Set Channel Type
                        .ChanType = "AO"
                        
                        'Set Channel description to an empty string
                        Set .ChanDescs = New ChannelDescs
                                                      
                    End With
                    
                Next j
                                                      
            End If

            'Find whether or not the Digital Input & Output channels are
            'automatically configured for Digital IO mode already
            'The Digital channels on the Measurement Computing boards
            'are the same for input and output, and need to be reconfigured
            'each time they are written to / read from
            'The Digital channels on the ADWIN boards are dedicated to either
            'input or output, and thus do not need to nor can be reconfigured
            .DIOConfigured = (Config_GetFromINI("Boards", _
                                                "DIOConfigured" & Format(i - 1, "0"), _
                                                "False", _
                                                Prog_INIFile) = "True")
                                    
            'If the comm protocol is "MCC", then figure out which
            'Digital Output Port type this is
            If .CommProtocol = MCC_UL Then
            
                .DoutPortType = val(Config_GetFromINI( _
                                        "Boards", _
                                        "DOutPortType" & Format(i - 1, "0"), _
                                        "-1", _
                                        Prog_INIFile))
                                        
                If .DoutPortType = -1 Then
                
                    isBoardLoadError = True
                    BoardLoadErrorCause = "Digital I/O Port Type missing for MCC board;" & _
                                          Format(i - 1, "0") & ","
                                          
                End If
                                        
            End If
                                        
            'Get Number of Digital Input Channels
            NumChannels = val(Config_GetFromINI( _
                                    "Boards", _
                                    "DInChannelsCount" & Format$(i - 1, "0"), _
                                    "-1", _
                                    Prog_INIFile))
                                    
                                           
            'Screen for Import error (Default, NumChannels = -1)
            If NumChannels < 0 Then
            
                'The digital input channel settings for this board are missing!
                isBoardLoadError = True
                BoardLoadErrorCause = BoardLoadErrorCause & _
                                        "Digital Input Channel Count missing;" & Format(i - 1, "0") & ","
            
            ElseIf NumChannels > 0 Then
            
                'Add the necessary number of channels
                'to the Digital Input channels collection on this Board
                For j = 1 To NumChannels
            
                    'Channel Name Str
                    TempStr = Trim(Config_GetFromINI( _
                                            "Boards", _
                                            "DI-" & _
                                            Format(SystemBoards(BoardNameStr).BoardININum, _
                                                   "0") & "-" & _
                                            "CH" & Format(j - 1, "0"), _
                                            "ERROR,-1", _
                                            Prog_INIFile))
                
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++'

'       Need to add code to parse channamestr - it currently contains both
'       the channel name and the channel number assignment
                
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++'

                    ChanNameStr = Mid(TempStr, _
                                      1, _
                                      InStr(1, TempStr, ",") - 1)

                
                    'Use Channel name as the key in the channels collection
                    .DInChannels.Add , ChanNameStr
                        
                    With .DInChannels(ChanNameStr)
                                                
                        'Set the Channel's parent Board Name
                        .BoardName = BoardNameStr
                    
                        'Set the Channel's parent Board #
                        .BoardININum = SystemBoards(.BoardName).BoardININum
                    
                        'Note: that the format of the key string for the channel
                        '      contains both an element indicating the board that
                        '      the channel is on, and also indicating the channel
                        '      number itself.
                        
                        'Snatch the Channel Name from the INI file
                        .ChanName = ChanNameStr
                                                                                  
                        'Snatch the Channel Number from the INI File
                        .ChanNum = val(Mid(TempStr, _
                                           InStr(1, TempStr, ",") + 1))
                                         
                        'Set Channel Type
                        .ChanType = "DI"
                        
                        'Set Channel description to an empty string
                        Set .ChanDescs = New ChannelDescs
                                                      
                    End With
                    
                Next j
                                                      
            End If
            
            'Get Number of Digital Output Channels
            NumChannels = val(Config_GetFromINI("Boards", _
                                           "DOutChannelsCount" & Format$(i - 1, "0"), _
                                           "-1", _
                                           Prog_INIFile))
                                           
            'Screen for Import error (Default, NumChannels = -1)
            If NumChannels < 0 Then
            
                'The digital input channel settings for this board are missing!
                isBoardLoadError = True
                BoardLoadErrorCause = BoardLoadErrorCause & _
                                        "Digital Output Channel Count missing;" & Format(i - 1, "0") & ","
            
            ElseIf NumChannels > 0 Then
            
                'Add the necessary number of channels
                'to the Digital Output channels collection on this Board
                For j = 1 To NumChannels
                    
                    'Channel Name Str
                    TempStr = Trim(Config_GetFromINI( _
                                            "Boards", _
                                            "DO-" & _
                                            Format(SystemBoards(BoardNameStr).BoardININum, _
                                                   "0") & "-" & _
                                            "CH" & Format(j - 1, "0"), _
                                            "ERROR,-1", _
                                            Prog_INIFile))
                
                    ChanNameStr = Mid(TempStr, _
                                      1, _
                                      InStr(1, TempStr, ",") - 1)

                
                    'Use Channel name as the key in the channels collection
                    .DOutChannels.Add , ChanNameStr
                    
                    With .DOutChannels(ChanNameStr)
                                                
                        'Set the Channel's parent Board Name
                        .BoardName = BoardNameStr
                    
                        'Set the Channel's parent Board #
                        .BoardININum = SystemBoards(.BoardName).BoardININum
                    
                        'Note: that the format of the key string for the channel
                        '      contains both an element indicating the board that
                        '      the channel is on, and also indicating the channel
                        '      number itself.
                        
                        'Snatch the Channel Name from the INI file
                        .ChanName = ChanNameStr
                                                                                  
                        'Snatch the Channel Number from the INI File
                        .ChanNum = val(Mid(TempStr, _
                                           InStr(1, TempStr, ",") + 1))
                                                      
                        'Set Channel Type
                        .ChanType = "DO"
                        
                        'Set Channel description to an empty string
                        Set .ChanDescs = New ChannelDescs
                                            
                                                      
                    End With
                    
                Next j
                                                      
            End If
            
        End With
        
    Next i
    
    'Check the error status of the board loading
    ErrorCheckBoardsImport isBoardLoadError, _
                           BoardLoadErrorCause
        
    ImportBoardsDone = True
            
    'Set NOCOMM_MODE = False
    modProg.NOCOMM_MODE = False

End Sub

Public Function Get_ChannelINIStr(ByVal ChanININame As String) As String

    
    Dim TempStr As String
    
    'Get Chan INI identifier
    TempStr = Trim(Config_GetFromINI("Channels", _
                                     ChanININame, _
                                     "ERROR", _
                                     Prog_INIFile))
   
    Get_ChannelINIStr = TempStr
    
End Function

' Sub Get_ChannelsFromIni()
'
' Created: March 30, 2010
'  Author: Isaac Hilburn
'
' Summary: Reads Paleomag.ini file and parses the [Channels] section of the file using
'          the Config_GetFromIni function.  Loads the correct channel from the correct board
'          into each needed global channel object variable (i.e. ARMVoltageOut, AnalogT1, etc.)

Public Sub Get_ChannelsFromIni()
    
    'Initialize all the global channel objects to Nothing
    Set AFRampChan = Nothing
    Set AFMonitorChan = Nothing
    Set AltAFMonitorChan = Nothing
    Set ARMVoltageOut = Nothing
    Set IRMCapacitorVoltageIn = Nothing
    Set IRMMonitor = Nothing
    Set IRMVoltageOut = Nothing
    Set AnalogT1 = Nothing
    Set AnalogT2 = Nothing
    Set AFAxialRelay = Nothing
    Set AFTransRelay = Nothing
    Set IRMRelay = Nothing
    Set ARMSet = Nothing
    Set IRMFire = Nothing
    Set IRMTrim = Nothing
    Set IRMPowerAmpVoltageIn = Nothing
    Set MotorToggle = Nothing
    Set VacuumToggleA = Nothing
    Set VacuumToggleB = Nothing
    Set DegausserToggle = Nothing
    
    'Initialize the System Assigned Channels object
    Set SystemAssignedChannels = New Channels
    
    'Initialize the error status variables to false and vbnullstrings
    ChanImportError = False
    ChanImportErrorCause = vbNullString
    DisabledModules = vbNullString
    
    'Check to see if Get_BoardsFromIni has been called. If not, call it
    If Not ImportBoardsDone And Not NoINIBoards Then
    
        Get_BoardsFromIni
        
    End If
    
    If NoINIBoards Then
    
        'There are no boards in the INI file, therefore the
        'channel settings, if they exist, are meaningless,
        'Set NoINIChannels = True
        NoINIChannels = True
        
        'Exit the subroutine
        Exit Sub
        
    End If
    
    'Check to see if the WaveForms have been loaded
    If ImportWavesDone Then
    
        'Wave import was successfull, can load the Wave Form channels
        
        'Check to see if the Ramp Up and/or Ramp Down wave-form is missing before continuing
        If isAFRampUpMissing = False Then
            
            'For the AF Ramp and three AF Monitor channels - these
            'channel objects can be snatched from the appropriate Wave object (25 cents, cheap!)
            WaveForms("AFRAMPUP").Chan.ChanDescs.Clear
            WaveForms("AFRAMPUP").Chan.ChanDescs.AddDesc "AF Ramp Output"
            
            Set AFRampChan = WaveForms("AFRAMPUP").Chan
                        
        ElseIf isAFRampDownMissing = False Then
            
            WaveForms("AFRAMPDOWN").Chan.ChanDescs.Clear
            WaveForms("AFRAMPDOWN").Chan.ChanDescs.AddDesc "AF Ramp Output"
            Set AFRampChan = WaveForms("AFRAMPDOWN").Chan
            
            WaveForms.Add WaveForms("AFRAMPDOWN"), "AFRAMPUP"
            
        End If
        
        'if the AF module is still enabled
        If isAFRampDownMissing = False Then
            
            'For the AF Ramp and three AF Monitor channels - these
            'channel objects can be snatched from the appropriate Wave object (25 cents, cheap!)
            WaveForms("AFRAMPDOWN").Chan.ChanDescs.Clear
            WaveForms("AFRAMPDOWN").Chan.ChanDescs.AddDesc "AF Ramp Output"
            
            Set AFRampChan = WaveForms("AFRAMPDOWN").Chan
                        
        ElseIf AFRampUpMissing = False Then
            
            WaveForms("AFRAMPUP").Chan.ChanDescs.Clear
            WaveForms("AFRAMPUP").Chan.ChanDescs.AddDesc "AF Ramp Output"
            Set AFRampChan = WaveForms("AFRAMPUP").Chan
            
            WaveForms.Add WaveForms("AFRAMPUP"), "AFRAMPDOWN"
            
        End If
        
        'check to see if both the AFRAMP UP and AFRAMP DOWN waveforms are missing
        If isAFRampDownMissing = True And _
           isAFRampUpMissing = True _
        Then
        
            'Add an item to the error cause, but do not set the error to true
            'this error will only show up if a non-wave related channel load error happens
            ChanImportErrorCause = ChanImportErrorCause & _
                                   "ADWIN AF Ramp Output channel missing,"
                                   
    
            AddDisabledModule "AF"
        
        End If
        
        'If AFRAMPChan is not nothing, then load it into the System Assigned Channels collection
        If Not AFRampChan Is Nothing Then
        
            'Add the AF Ramp Output Channel to the System Assigned Channels collection
            SystemAssignedChannels.Add AFRampChan, AFRampChan.ChanName
            
        End If
            
        'Check to make sure that the AF monitor Wave form is not missing
        If isAFMonitorMissing = False Then
                
            WaveForms("AFMONITOR").Chan.ChanDescs.Clear
            WaveForms("AFMONITOR").Chan.ChanDescs.AddDesc "AF Monitor"
            Set AFMonitorChan = WaveForms("AFMONITOR").Chan
            
            'Add the AF Monitor Channel to the System Assigned Channels collection
            'Add without channel key repeats
            AddToChanColNoKeyRepeat SystemAssignedChannels, _
                                 AFMonitorChan, _
                                 AFMonitorChan.ChanName
                                                                  
        Else
        
            'Wha-oh, AF Monitor waveforms is missing!
            'Add an item to the error cause, but do not set the error to true
            'this error will only show up if a non-wave related channel load error happens
            ChanImportErrorCause = ChanImportErrorCause & _
                                   "ADWIN AF Monitor channel missing,"
            
            AddDisabledModule "AF"
            
        End If
        
        'Check to make sure taht the AF (2G) Alternate Monitor Wave form is not missing
        If isAltAFMonitorMissing = False Then
            
            WaveForms("ALTAFMONITOR").Chan.ChanDescs.Clear
            WaveForms("ALTAFMONITOR").Chan.ChanDescs.AddDesc "Alternate AF Monitor"
            Set AltAFMonitorChan = WaveForms("ALTAFMONITOR").Chan
            
            'Add the Alternate AF Monitor Channel to the System Assigned Channels collection
            'Add without channel key repeats
            AddToChanColNoKeyRepeat SystemAssignedChannels, _
                                 AltAFMonitorChan, _
                                 AltAFMonitorChan.ChanName
                                 
        Else
        
            'Add an item to the error cause, but do not set the error to true
            'this error will only show up if a non-wave related channel load error happens
            ChanImportErrorCause = ChanImportErrorCause & _
                                   "Alternate AF (2G) Monitor channel missing,"
                                   
            AddDisabledModule "AltAFMonitor"
                                   
        End If
    
        If isIRMMonitorMissing = False Then
        
            'Get the IRM Monitor analog input channel
            WaveForms("IRMMONITOR").Chan.ChanDescs.Clear
            WaveForms("IRMMONITOR").Chan.ChanDescs.AddDesc "IRM Monitor"
            Set IRMMonitor = WaveForms("IRMMONITOR").Chan
        
            'Add the IRM Monitor Channel to the System Assigned Channels collection
            'Add without channel key repeats
            AddToChanColNoKeyRepeat SystemAssignedChannels, _
                                    IRMMonitor, _
                                    IRMMonitor.ChanName
            
        Else
        
            'Add an item to the error cause, but do not set the error to true
            'this error will only show up if a non-wave related channel load error happens
            ChanImportErrorCause = ChanImportErrorCause & _
                                   "IRM Monitor channel missing,"
                                   
            AddDisabledModule "IRMMonitor"

        End If
        
    Else
    
        'Trigger a channel import error
        ChanImportError = True
    
        'Add to the Channel Import Error Cause
        ChanImportErrorCause = ChanImportErrorCause & _
                               "AF / IRM wave form settings missing." & _
                               "ADWIN AF / Alternate AF monitor / and IRM monitor channels not loaded,"
                               
        AddDisabledModule "AF"
        AddDisabledModule "AltAFMonitor"
        AddDisabledModule "IRMMonitor"

    End If
    
        
    'Note: For Each channel that we need to load, need to store the ChanStr from the ini
    '      file, then parse the string to link to the correct Board & Channel in the System
    '      Boards Collection
    
    'Get the ARM Voltage Output Channel
    Set ARMVoltageOut = Retrieve_Channel(Get_ChannelINIStr("ARMVoltageOut"), _
                                         "ARM Voltage Output")
    
    'Add the ARM Voltage Output Channel to the System Assigned Channels collection
    'Add without channel key repeats
    If Not ARMVoltageOut Is Nothing Then
        
        AddToChanColNoKeyRepeat SystemAssignedChannels, _
                                ARMVoltageOut, _
                                ARMVoltageOut.ChanName
    
    Else
    
        AddDisabledModule "ARM"
        
    End If
                                         
                                         
    'Get the IRM Voltage Output Channel
    Set IRMVoltageOut = Retrieve_Channel(Get_ChannelINIStr("IRMVoltageOut"), _
                                         "IRM Voltage Output")
    
    'Add the IRM Voltage Output Channel to the System Assigned Channels collection
    'Add without channel key repeats
    If Not IRMVoltageOut Is Nothing Then
        
        AddToChanColNoKeyRepeat SystemAssignedChannels, _
                             IRMVoltageOut, _
                             IRMVoltageOut.ChanName
    
    
    Else
    
        AddDisabledModule "IRM"
        
    End If
                                         
                                         
    'Get the IRM Capacitor Monitor Voltage Input channel
    Set IRMCapacitorVoltageIn = Retrieve_Channel(Get_ChannelINIStr("IRMCapacitorVoltageIn"), _
                                         "IRM Capacitor Return Voltage")
    
    'Add the ARM Voltage Output Channel to the System Assigned Channels collection
    'Add without channel key repeats
    If Not IRMCapacitorVoltageIn Is Nothing Then
        
        AddToChanColNoKeyRepeat SystemAssignedChannels, _
                                IRMCapacitorVoltageIn, _
                                IRMCapacitorVoltageIn.ChanName
    
    Else
    
        AddDisabledModule "IRMReturn"
        
    End If
                                         
    'Get the Analog Temperature #1 sensor input channel
    Set AnalogT1 = Retrieve_Channel(Get_ChannelINIStr("AnalogT1"), _
                                    "Temperature Sensor #1")
    
    'Add the AF Coil Temp. Sensor Input Channel to the System Assigned Channels collection
    'Add without channel key repeats
    If Not AnalogT1 Is Nothing Then
        
        AddToChanColNoKeyRepeat SystemAssignedChannels, _
                                AnalogT1, _
                                AnalogT1.ChanName
    
    Else
    
        AddDisabledModule "AnalogT1"
        
    End If
                                         
                                         
    'Get the Analog Temperature #2 sensor input channel
    Set AnalogT2 = Retrieve_Channel(Get_ChannelINIStr("AnalogT2"), _
                                    "Temperature Sensor #2")
    
    'Add the AF Coil Temp. Sensor Input Channel to the System Assigned Channels collection
    'Add without channel key repeats
    If Not AnalogT2 Is Nothing Then
        
        AddToChanColNoKeyRepeat SystemAssignedChannels, _
                                AnalogT2, _
                                AnalogT2.ChanName
    
    Else
    
        AddDisabledModule "AnalogT2"
        
    End If
    
                                         
    'Get the AF Axial Relay TTL Channel
    Set AFAxialRelay = Retrieve_Channel(Get_ChannelINIStr("AFAxialRelay"), _
                                        "AF Axial Relay")
    
    'Add the AF Axial Relay Channel to the System Assigned Channels collection
    'Add without channel key repeats
    If Not AFAxialRelay Is Nothing Then
        
        AddToChanColNoKeyRepeat SystemAssignedChannels, _
                                AFAxialRelay, _
                                AFAxialRelay.ChanName
    
    Else
    
        AddDisabledModule "AF"
        AddDisabledModule "IRMAxial"
        
    End If
        
                                         
    'Get the AF Trans Relay TTL Channel
    Set AFTransRelay = Retrieve_Channel(Get_ChannelINIStr("AFTransRelay"), _
                                        "AF Transverse Relay")
    
    'Add the AF Transverse Relay Channel to the System Assigned Channels collection
    'Add without channel key repeats
    If Not AFTransRelay Is Nothing Then
        
        AddToChanColNoKeyRepeat SystemAssignedChannels, _
                                AFTransRelay, _
                                AFTransRelay.ChanName
    
    Else
    
        AddDisabledModule "AF"
        AddDisabledModule "IRMTrans"
        
    End If
                                             
                                         
    'Get the IRM Relay TTL Channel
    Set IRMRelay = Retrieve_Channel(Get_ChannelINIStr("IRMRelay"), _
                                    "IRM Relay")
    
    'Add the IRM Relay Channel to the System Assigned Channels collection
    'Add without channel key repeats
    If Not IRMRelay Is Nothing Then
        
        AddToChanColNoKeyRepeat SystemAssignedChannels, _
                                IRMRelay, _
                                IRMRelay.ChanName
    
    Else
    
        AddDisabledModule "IRM"
        
    End If
                                             
                                             
    'Get the ARM Set digital output channel
    Set ARMSet = Retrieve_Channel(Get_ChannelINIStr("ARMSet"), _
                                  "ARM Set")
                                  
    'Add the ARM Set Channel to the System Assigned Channels collection
    'Add without channel key repeats
    If Not ARMSet Is Nothing Then
        
        AddToChanColNoKeyRepeat SystemAssignedChannels, _
                                ARMSet, _
                                ARMSet.ChanName
    
    Else
    
        AddDisabledModule "ARM"
        
    End If
                                             
    
    'Get the IRM Fire digital output channel
    Set IRMFire = Retrieve_Channel(Get_ChannelINIStr("IRMFire"), _
                                   "IRM Fire")
    
    'Add the IRM Fire Channel to the System Assigned Channels collection
    'Add without channel key repeats
    If Not IRMFire Is Nothing Then
        
        AddToChanColNoKeyRepeat SystemAssignedChannels, _
                                IRMFire, _
                                IRMFire.ChanName
    
    Else
    
        AddDisabledModule "IRM"
        
    End If
                                             
    
    'Get the IRM Trim digital output channel
    Set IRMTrim = Retrieve_Channel(Get_ChannelINIStr("IRMTrim"), _
                                   "IRM Trim")
    
    'Add the IRM Trim Channel to the System Assigned Channels collection
    'Add without channel key repeats
    If Not IRMTrim Is Nothing Then
        
        AddToChanColNoKeyRepeat SystemAssignedChannels, _
                                IRMTrim, _
                                IRMTrim.ChanName
    
    Else
    
        AddDisabledModule "IRM"
        
    End If
        
                                         
    'Get the IRM Ready status digital input Channel
    Set IRMPowerAmpVoltageIn = Retrieve_Channel(Get_ChannelINIStr("IRMPowerAmpVoltageIn"), _
                                    "IRM Power Amp Voltage Input")
    
    'Add the IRM Ready Channel to the System Assigned Channels collection
    'Add without channel key repeats
    If Not IRMPowerAmpVoltageIn Is Nothing Then
        
        AddToChanColNoKeyRepeat SystemAssignedChannels, _
                                IRMPowerAmpVoltageIn, _
                                IRMPowerAmpVoltageIn.ChanName
        
    End If
                                             
                                         
    'Get the Motor Toggle digital Output Channel
    Set MotorToggle = Retrieve_Channel(Get_ChannelINIStr("MotorToggle"), _
                                       "Vacuum Motor Toggle")
    
    'Add the Motor Toggle Channel to the System Assigned Channels collection
    'Add without channel key repeats
    If Not MotorToggle Is Nothing Then
        
        AddToChanColNoKeyRepeat SystemAssignedChannels, _
                                MotorToggle, _
                                MotorToggle.ChanName
    
    Else
    
        AddDisabledModule "Vacuum"
        
    End If
        
                                         
    'Get the Vacuum Toggle A digital Output Channel
    Set VacuumToggleA = Retrieve_Channel(Get_ChannelINIStr("VacuumToggleA"), _
                                         "Vacuum Relay Toggle A")
                                            
    'Add the Vacuum Toggle A Channel to the System Assigned Channels collection
    'Add without channel key repeats
    If Not VacuumToggleA Is Nothing Then
        
        AddToChanColNoKeyRepeat SystemAssignedChannels, _
                                VacuumToggleA, _
                                VacuumToggleA.ChanName
    
    Else
    
        AddDisabledModule "Vacuum"
        
    End If
    
    'Get the DegausserToggle digital Output Channel
    Set DegausserToggle = Retrieve_Channel(Get_ChannelINIStr("DegausserToggle"), _
                                         "Degausser Toggle")
                                         
    'Add the Degausser Cooler Channel to the System Assigned Channels collection
    'Add without channel key repeats
    If Not DegausserToggle Is Nothing Then
        
        AddToChanColNoKeyRepeat SystemAssignedChannels, _
                                DegausserToggle, _
                                DegausserToggle.ChanName
    
    Else
    
        AddDisabledModule "DegausserToggle"
        
    End If
        
'----------------------------------------------------------------------------------------------------------------'
'----------------------------------------------------------------------------------------------------------------'
'
'(June 2010, I Hilburn)
'
'These channel assignments have been commented out
'
'With the two current IRM hardware configurations, it is not possible to capture and record the A/D signal
'of the actual IRM pulse, therefore an IRM monitor channel is unnecessary
'
'Vacuum Toggle B is no longer used in the hardware setup, and is therefore obsolete
'----------------------------------------------------------------------------------------------------------------'
'
'    'Get the Vacuum Toggle B digital Output Channel
'    Set VacuumToggleB = Retrieve_Channel(Get_ChannelINIStr("VacuumToggleB"), _
'                                         "Vacuum relay toggle B digital output channel")
'
'    'Add the Vacuum Toggle B Channel to the System Assigned Channels collection
'    'Add without channel key repeats
'    AddToChanColNoKeyRepeat SystemAssignedChannels, _
'                          VacuumToggleB, _
'                          VacuumToggleB.ChanName
'
'----------------------------------------------------------------------------------------------------------------'
'----------------------------------------------------------------------------------------------------------------'

    'Check for errors during the channel import process
    ErrorCheckChannelsImport ChanImportError, _
                             ChanImportErrorCause

    'Set Channel Import Done status flag to True
    ChannelsImportDone = True
    
End Sub

'Public Sub Get_WaveFormsFromIni()
'
' Created: March 31, 2010
' Author: Isaac Hilburn
'
' Summary: Creates the default Wave objects needed for AF/pARM ramp cycle and monitoring,
'          and the monitoring of the IRM capacitor voltage.  The fields in the five waves
'          are also populated using the Config_GetFromIni function that reads in values
'          from the Paleomag.ini file.
'
'          The code also uses the Retrieve_Channel function to get the channel
'          object associated with the Channel string pulled from the Paleomag.ini file
'          for each Wave.  This channel contains the port # & board # the wave will use
'          for it's IO process. The Board # & name in the channel object are redundant.
'          They are also contained in the Wave's .BoardUsed field, but the redundancy
'          is to prevent orphan channels / ambiguity about which board a channel object
'          in and of itself, belongs to.

Public Sub Get_WaveFormsFromIni()

    'Like Get_BoardsFromIni, the number associated with each wave-form is appended onto the
    'string key name in the Paleomag.ini file so that each key is unique to each wave.
    '
    'This allows us to use the pre-existing Config_GetFromIni function that Bob Kopp
    'built.
    
    Dim i As Long
    Dim N As Long
    
    Dim BoardStr As String
    Dim BoardININum As Long
    Dim BoardName As String
    Dim WaveNameStr As String
    Dim ChanStr As String
    Dim ErrorMsg As String
    
    Dim WaveImportError As Boolean
    Dim WaveImportErrorCause As String
    
    'Set import errors equal to false initially
    WaveImportError = False
    WaveImportErrorCause = vbNullString
    
    'Read in the number of Wave objects whose settings are stored in the Paleomag.ini
    'file
    N = val(Config_GetFromINI("WaveForms", _
                              "WaveFormCount", _
                              "-1", _
                              Prog_INIFile))
    
    If N = -1 Or N = 0 Then
    
        'N = -1, the [WaveForms] section of the .ini file doesn't exist
        'N = 0, No waveforms have been saved to the .ini file
        
        'Set NoINIWaveForms = True
        NoINIWaveForms = True
        
        'Check which AF system is being used
        If AFSystem = "2G" Then
        
            ErrorMsg = "IRM Monitor"
            EnableIRMMonitor = False
            EnableAltAFMonitor = False
            
        Else
        
            ErrorMsg = "AF, ARM and the IRM Monitor"
            EnableAF = False
            EnableARM = False
            EnableAFAnalysis = False
            EnableIRMMonitor = False
            EnableAltAFMonitor = False
        
        End If
        
        'Tell the user that the INI file is missing and that a number of modules need to be disabled
        'for the code to continue running
        UserResp = frmDialog.DialogBox("The " & ErrorMsg & " settings block in the [WaveForms] section of the " & _
                                       "INI settings file is missing." & vbNewLine & vbNewLine & _
                                       "Do you wish to continue loading the code with the " & ErrorMsg & " modules " & _
                                       "disabled?", _
                                       "INI Settings File Error", _
                                       3, _
                                       "Yes", _
                                       "No", _
                                       "Open INI File Browser")
                                       
        If UserResp = vbNo Then
        
            'User has asked to end the code
            MsgBox "Paleomag code will end now."
            
            
            
            End
            
        End If
        
        If UserResp = vbCancel Then
        
            'Insert code here to open the INI file viewer
            
        End If
        
        'Else, just exit the subroutine
        Exit Sub
                  
    End If
                  
    'Need to initialize the wave forms collection
    Set WaveForms = New Waves
                  
    'Otherwise, there is at least one waveform to load into
    'the global WaveForms collection
    For i = 0 To N - 1
    
        'Get the WaveName to use as the key for the WaveForms collection
        WaveNameStr = Trim(Config_GetFromINI("WaveForms", _
                                             "WaveName" & Format(i, "0"), _
                                             "ERROR", _
                                             Prog_INIFile))
    
        If WaveNameStr = "ERROR" Then
        
            WaveImportError = True
            WaveImportErrorCause = WaveImportErrorCause & _
                                    "No Wave Name specified;Loop#" & Format(i, 0) & ","
            
        End If
        
        If i + 1 > WaveForms.Count Then
            
            WaveForms.Add , WaveNameStr
        
        End If
            
        With WaveForms(WaveNameStr)
        
            'Get the description of the Wave Form
            .WaveDesc = Trim(Config_GetFromINI("WaveForms", _
                                               "WaveDesc" & Format(i, "0"), _
                                               "ERROR, -1", _
                                               Prog_INIFile))
    
            'Get the string name of the board for this wave form in the
            'BoardName .ini file field
            BoardStr = Trim(Config_GetFromINI("WaveForms", _
                                              "BoardUsed" & Format(i, "0"), _
                                              "ERROR,-1", _
                                              Prog_INIFile))
            
            'Error Checking
            If BoardStr = "ERROR,-1" Then
            
                'Blast it all to tarnation!
                'There's no board specified for this wave object in the INI file!
                'Flag the wave import error status flag
                WaveImportError = True
                WaveImportErrorCause = WaveImportErrorCause & _
                                            "No INI Board Specified for Wave;" & WaveNameStr & ","
                              
            End If
                
            'BoardStr has the following format:
            '<Board Name>,<Board Device Number>
            BoardName = Mid(BoardStr, _
                            1, _
                            InStr(1, BoardStr, ",") - 1)
                            
            BoardININum = val(Mid(BoardStr, _
                                  InStr(1, BoardStr, ",") + 1))
                           
            'Set .BoardUsed property of wave object to nothing
            Set .BoardUsed = Nothing
                           
            'Can use the BoardName to snatch the matching board from
            'the System DAQ Boards collection
            On Error Resume Next
            
                Set .BoardUsed = SystemBoards(BoardName)
                
                If Err.number <> 0 Then
                
                    'Blister me failed liver!
                    'The Board specified for this Wave object in the INI file
                    'does not match an existing board in the Systems board collection.
                    WaveImportError = True
                    WaveImportErrorCause = WaveImportErrorCause & _
                                            "No matching System Board for Wave Board Name = """ & _
                                            BoardName & """;" & WaveNameStr & ","
                              
                End If
                
            On Error GoTo 0
            
            If .BoardUsed Is Nothing Then
            
                'A pox on ye pale-faced commmputer prooogrammers! (Arr!)
                'The Board specified for this Wave object in the INI file still
                'does not match an existing board in the Systems board collection.
                WaveImportError = True
                WaveImportErrorCause = WaveImportErrorCause & _
                                            "No matching System Board for Wave Board Name = """ & _
                                            BoardName & """;" & WaveNameStr & ","
                          
            End If
                                              
            'Get the string identifier of the wave's channel
            ChanStr = Trim(Config_GetFromINI("WaveForms", _
                                             "Chan" & Format(i, "0"), _
                                             "ERROR,-1", _
                                             Prog_INIFile))
            
            'Error Checking
            If ChanStr = "ERROR,-1" Then
            
                'Poo.
                'There's no Channel specified for this wave object in the INI file!
                'NOTE: if the .BoardUsed property wasn't successully loaded,
                '      then this error message will break the code when trying to
                '      access the DAQ Board name from .BoardUsed.
                WaveImportError = True
                WaveImportErrorCause = WaveImportErrorCause & _
                                            "No Wave Channel Specified;" & WaveNameStr & ","
                
                                          
            End If
            
            'Get the Wave's INI # key
            .WaveININum = val(Config_GetFromINI("WaveForms", _
                                                "WaveININum" & Format(i, "0"), _
                                                Trim(str(i)), _
                                                Prog_INIFile))
                                              
            'Get the WaveName
            .WaveName = WaveNameStr
                                      
            'Start .Chan = Nothing
            Set .Chan = Nothing
                                          
            'Turn on error checking
            On Error Resume Next
                
                'Retrieve the channel associated with the wave using
                'the Channel string found above
                Set .Chan = Retrieve_Channel(ChanStr, _
                                             .WaveName & " wave")
                                                        
                'Error check
                If Err.number <> 0 Then
                
                    WaveImportError = True
                    WaveImportErrorCause = WaveImportErrorCause & _
                                            "No matching System Channel for Wave Channel = """ & _
                                            ChanStr & """;" & WaveNameStr & ","
                    
                End If
                
            On Error GoTo 0
                    
            'Error check a second time
            If .Chan Is Nothing Then
            
                WaveImportError = True
                WaveImportErrorCause = WaveImportErrorCause & _
                                            "No matching System Channel for Wave Channel = """ & _
                                            ChanStr & """;" & WaveNameStr & ","
                                             
            End If
                    
                                                 
            'Reset channel description to empty string
            'If .Chan is nothing, this code will crash, need to check for is nothing case
            If Not .Chan Is Nothing Then
                
                Set .Chan.ChanDescs = Nothing
                Set .Chan.ChanDescs = New ChannelDescs
                
            End If
                                         
            'Get the start point to begin collecting or outputing data from
            'in the wave
            .StartPoint = CLng(val(Config_GetFromINI("WaveForms", _
                                                "StartPoint" & Format(i, "0"), _
                                                "0", _
                                                Prog_INIFile)))
                                                 
            'Get the status of whether memory space - as a global object or
            'a windows memory buffer has been allocated for the wave-form
            .BufferAlloc = (Config_GetFromINI("WaveForms", _
                                           "MemAlloc" & Format(i, "0"), _
                                           "False", _
                                           Prog_INIFile) = "True")
                                               
            'Get the status flag that indicates if the memory space for the wave-form
            'should be emptied after the wave-form has finished being used by
            'the Paleomag code
            .DoDeallocate = (Config_GetFromINI("WaveForms", _
                                               "DoDeallocate" & Format(i, "0"), _
                                               "True", _
                                               Prog_INIFile) = "True")
                
            'Record if this is an input or an output wave object
            .IO = Trim(Config_GetFromINI("WaveForms", _
                                         "IO" & Format(i, "0"), _
                                         "ERROR", _
                                         Prog_INIFile))
                                             
            '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
            'This realllllly needs to be error checked, could break
            'the code and some of the hardware
            '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
            If .IO = "ERROR" Then
                
                '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
                'If this error happens, flag the wave import error
                '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
                
                'If it does, then the bug that caused it needs to be traced down
                'and squashed!  (The programmer who wrote in the bug, on the other hand,
                'should be reassured and offered comfy pillows, warm beverages,
                'and a hearty pay-raise.)
                WaveImportError = True
                WaveImportErrorCause = WaveImportErrorCause & _
                                        "Bad IO Mode = """ & .IO & """;" & WaveNameStr & ","
                        
            End If
            
           'Get the IORate at which the wave should be output/input
            .IORate = val(Config_GetFromINI("WaveForms", _
                                            "IORate" & Format(i, "0"), _
                                            "50000", _
                                            Prog_INIFile))
                                                        
            'Now time to error check the IORate and make sure it's smaller
            'than the max allowed rate for the BoardObject associated with this
            'waveform.
            '
            'NOTE: the code below will crash if the .BoardUsed property = Nothing
            '      or is Empty
            If Not .BoardUsed Is Nothing Then
            
                If .IO = "INPUT" And .IORate > .BoardUsed.MaxAInRate Then
                
                    '.IORate exceeds the maximum analog input rate for the Board
                    'associated with this wave object.  Set the IORate = Max Analog
                    'Input rate
                    .IORate = .BoardUsed.MaxAInRate
                    
                ElseIf .IO = "OUTOUT" And .IORate > .BoardUsed.MaxAOutRate Then
                
                    '.IORate exceeds the maximum analog output rate for the Board
                    'associated with this wave object.  Set the IORate = Max Analog
                    'Output rate
                    .IORate = .BoardUsed.MaxAOutRate
                
                End If
                
            End If
            
            'Set the WaveForm's Range = a new Range object
            Set .range = New range
                        
            'Get the RangeType (Measurement Computing boards, only) for the
            'output/input of the wave
            .range.RangeType = val(Config_GetFromINI("WaveForms", _
                                                     "RangeType" & Format(i, "0"), _
                                                     "-1", _
                                                     Prog_INIFile))
                                                     
            'Check to see if there was a RangeType field in the .ini file
            If .range.RangeType = -1 Then
            
                'RangeType is equal to default value, therefore there was
                'no RangeType field for this wave in the .ini file.
                
                'Need to find the RangeMax and RangeMin fields
                '(non-Measurement computing boards only)
                .range.MaxValue = val(Config_GetFromINI("WaveForms", _
                                                        "RangeMax" & Format(i, "0"), _
                                                        "10", _
                                                        Prog_INIFile))
                                                        
                .range.MinValue = val(Config_GetFromINI("WaveForms", _
                                                        "RangeMin" & Format(i, "0"), _
                                                        "-10", _
                                                        Prog_INIFile))
                                                        
            End If
                                                           
        End With
        
    Next i
    
    'Do the error check on the system WaveForms collection
    ErrorCheckWavesImport WaveImportError, WaveImportErrorCause
    
    'Set import status flag to "True"
    ImportWavesDone = True

End Sub

Public Function GetLongUnits(Optional ByVal ShortUnits As String = vbNullString) As String

    'If no value has been entered in for the shortunits, write the value of AFUnits
    'into ShortUnits
    If ShortUnits = vbNullString Then ShortUnits = AFUnits

    If ShortUnits = "G" Then GetLongUnits = "Gauss"
    If ShortUnits = "mT" Then GetLongUnits = "Telsa"
    If ShortUnits = "A/m" Then GetLongUnits = "Amps/m"
    If ShortUnits = "Oe" Then GetLongUnits = "Oersted"

End Function

Public Function PadLeft(ByVal sString As String, _
                        ByVal sLen As Long)
                        
    Dim TempS As String
                        
    'Check for empty string case
    If sLen < 1 Then
    
        PadLeft = ""
        
        Exit Function
        
    End If
    
    'Pad with a sufficient number of white spaces
    TempS = String$(sLen, " ") & sString
    
    'Now cut a string of the desired length from the
    'right-most characters of TempS
    PadLeft = Right(TempS, sLen)
    
End Function

'Private function Retrieve_Channel(String, [String], [String])
'
' Created: March 30, 2010
'  Author: Isaac Hilburn
'
' Summary: Takes a specially formatted channel string, a default value, and a
'          channel description to use for error messages and finds the matching
'          board and channel in the System Boards collection's collections of channels
'          and returns the appropriate channel object to be loaded into a channel
'          object variable.
'
'  Inputs:
'
'     sChan     -   Specially formated string:
'                   "(A,D)(O,I)-{Board #}-CH{Channel number,Format String = "0"}"
'                   Indicating what type of channel we're looking for (A = Analog,
'                   D = Digital; O = Output, I = Input), the device # of the board
'                   on which to search for the channel, and the Channel #
'
'
'  ChanDesc     -  String - Contains the name of the channel in plain English
'                  that is being searched for in the .Ini file.  This string
'                  is used in generating the Error message that is raised if
'                  the Retrieve_Channel function fails to find a Board or matching
'                  channel object.  The specificity in the error message will
'                  allow better diagnosis of the problem, as formatting errors
'                  in the Paleomag.ini file will be the primary cause of an error.
'
' Outputs:
'
'           The function returns a copy of the desired channel object that refers
'           to the correct channel/port & board assignment

Private Function Retrieve_Channel _
    (ByVal sChan As String, _
     Optional ByVal ChanDesc As String = "") As Channel

    Dim ChanNameStr As String
    Dim BoardNameStr As String
    Dim ChanTypeStr As String
    Dim TempStr As String
    
    Dim BoardININum As Long
    
    Dim TempBoard As Board
    Dim TempChan As Channel
    
    'Initialize TempBoard and TempChan to Nothing
    Set TempBoard = Nothing
    Set TempChan = Nothing
    
    'Check sChan to see if it has defaulted to the "ERROR" value
    If sChan = "ERROR" Then
    
        'Toggle the Channel import error on
        ChanImportError = True
        ChanImportErrorCause = ChanImportErrorCause & _
                               "INI Channel Name not found;" & _
                               ChanDesc & ","
                               
        Set Retrieve_Channel = Nothing
        
        Exit Function
    
    End If
    
    'Go directly to the Boards section and get the Channel object's
    'name using the INI Channel string identifier
    TempStr = Trim(Config_GetFromINI("Boards", _
                                     sChan, _
                                     "ERROR,-1", _
                                     Prog_INIFile))
                                    
    If TempStr = "ERROR,-1" Then
    
        'Toggle the Channel import error on
        ChanImportError = True
        ChanImportErrorCause = ChanImportErrorCause & _
                               "No matching Board channel for Channel INI Name = """ & _
                               sChan & """;" & _
                               ChanDesc & ","
                               
        Set Retrieve_Channel = Nothing
        
        Exit Function
                               
    End If
                                    
    'TempStr = "<ChanName>,<ChanNum>"
    ChanNameStr = Mid(TempStr, 1, InStr(1, TempStr, ",") - 1)
    
    'Get Board Number from the channel string
    BoardININum = val(Mid(sChan, 4, 1))
    
    'Now grab that board's BoardName
    BoardNameStr = Trim(Config_GetFromINI("Boards", _
                                          "BoardName" & Trim(str(BoardININum)), _
                                          "ERROR", _
                                          Prog_INIFile))
    
    'Check for error
    If BoardNameStr = "ERROR" Then
    
        'No matching board in Boards collection
        ChanImportError = True
        ChanImportErrorCause = ChanImportErrorCause & _
                               "No parent DAQ Board matching Channel name = """ & _
                               ChanNameStr & """;" & _
                               ChanDesc & ","
        
        Set Retrieve_Channel = Nothing
        
        Exit Function
        
        
    End If
    
    On Error Resume Next
        
        'If BoardName is valid then
        Set TempBoard = SystemBoards(BoardNameStr)
    
        'Error Check
        If Err.number <> 0 Then
    
            'CRAP! No matching board in the system boards collection
            ChanImportError = True
            ChanImportErrorCause = ChanImportErrorCause & _
                                   "No matching System Board for Channel INI Board Name = """ & _
                                   BoardNameStr & """;" & _
                                   ChanDesc
                      
            Set Retrieve_Channel = Nothing
            
            Exit Function
            
        End If
        
    On Error GoTo 0
    
    If TempBoard Is Nothing Then
    
        'CRAP! No matching board in the system boards collection
        ChanImportError = True
        ChanImportErrorCause = ChanImportErrorCause & _
                               "No matching System Board for Channel INI Board Name = """ & _
                               BoardNameStr & """;" & _
                               ChanDesc
                  
        Set Retrieve_Channel = Nothing
        
        Exit Function
                  
    End If
    
    'Parse out the Channel Type
    ChanTypeStr = Mid(sChan, 1, 2)
    
    On Error GoTo BadChannel:
    
        Select Case ChanTypeStr
    
            Case "AO"
        
                Set TempChan = TempBoard.AOutChannels(ChanNameStr)
                
            Case "AI"
            
                Set TempChan = TempBoard.AInChannels(ChanNameStr)
                
            Case "DO"
            
                Set TempChan = TempBoard.DOutChannels(ChanNameStr)
                
            Case "DI"
            
                Set TempChan = TempBoard.DInChannels(ChanNameStr)
                
            Case Else
            
                ChanImportError = True
                ChanImportErrorCause = ChanImportErrorCause & _
                                       "Bad Channel Type = """ & ChanTypeStr & """;" & _
                                       ChanDesc & ","
                
                Set Retrieve_Channel = Nothing
                
                Exit Function
                
        End Select
               
    On Error GoTo 0
                      
    'Error check again for a Nothing value in TempChan
    If TempChan Is Nothing Then
    
        'Arrrrgh! No matching channel object made it into the
        'channels collection of the this type in the System Boards Collection!
        'Possible bad channel name in the INI file, or something wacky
        'going on with the System Boards collection
        ChanImportError = True
        ChanImportErrorCause = ChanImportErrorCause & _
                               "No match in " & ChanTypeStr & " channels collection for System Board = """ & _
                               TempBoard.BoardName & """ and Channel Name = """ & ChanNameStr & """;" & _
                               ChanDesc & ","
        
        Exit Function
        
    End If
    
    'Store the chan description in the TempChan channel object
    Set TempChan.ChanDescs = Nothing
    Set TempChan.ChanDescs = New ChannelDescs
    TempChan.ChanDescs.AddDesc ChanDesc
    
    'Return found channel and deallocate the local object variables
    Set Retrieve_Channel = TempChan
    
    Set TempChan = Nothing
    Set TempBoard = Nothing

    Exit Function
    
BadChannel:
        
        'Arrrrgh! No matching channel object made it into the
        'channels collection of the this type in the System Boards Collection!
        'Possible bad channel name in the INI file, or something wacky
        'going on with the System Boards collection
        ChanImportError = True
        ChanImportErrorCause = ChanImportErrorCause & _
                               "No match in " & ChanTypeStr & " channels collection for System Board = """ & _
                               TempBoard.BoardName & """ and Channel Name = """ & ChanNameStr & """;" & _
                               ChanDesc & ","
        
        'Return a nothing value
        Set Retrieve_Channel = Nothing
    
End Function


'' Commented Out: August 28, 2010
''            By: Isaac Hilburn
''
''        Reason: The function below is obsolete, is never called, and is unsupported by the Boards class library
''                If this function were to be called, it would crash the Paleomag code
''
''Private Function Retrieve_Board()
''    (Optional ByVal BoardNameStr as String = "",
''     Optional ByVal BoardNum as long = -1
''     Optional ByVal BoardDesc As String = "board")
''
'' Created: April 2, 2010
''  Author: Isaac Hilburn
''
''
''
'' Summary: This function takes in either a board name as a string or a board number
''          and searched through the System Boards collection for a matching board
''          object.  The matching board object is then returned. If no matching
''          board object is found, an error is raised.
''
''  Inputs:
''
'' BoardNameStr  -   Optional String, default value = "".  This string should contain
''                   the BoardName that is associated with the desired board object.
''                   The BoardName for the board is stored in the [Boards] section of
''                   the Paleomag.ini file.
''
''  BoardNum     -   Optional Long, default value = -1.  This long number should contain
''                   the Device number associated with the desired board object.  As with
''                   the BoardName, this Board # will be stored in the [Boards] section of
''                   the Paleomag.ini file
''
''  BoardDesc    -   Optional String, default value = "board".  This string contains the plain
''                   English descriptor to use in error messages when referring to the board
''                   that the user was seeking to retrieve.
''
'' Outputs:
''
''   Board object -  Returns the Board object that the user sought to retrieve
''
'' NOTE: This function requires the System Boards collection to have been populated.
''       if the SystemBoards.Count = 0, then this function will call Get_BoardsFromIni
''
'Private Function Retrieve_Board _
'    (Optional ByVal BoardNameStr As String = "", _
'     Optional ByVal BoardNum As Long = -1, _
'     Optional ByVal BoardININum As Long = -1, _
'     Optional ByVal BoardDesc As String = "board") As Board
'
'    Dim TempBoard As Board
'
'    Set TempBoard = Nothing
'
'    'Check for both input fields being empty / set to default values
'    If BoardNameStr = "" And BoardNum = -1 Then
'
'        'This function has been called wrong, raise an error
'        Err.Raise -300, _
'                  "modConfig->Retrieve_Board", _
'                  "Bad inputs given to Retrieve_Board function." & _
'                  vbNewLine & vbNewLine & _
'                  "Board sought = " & BoardDesc & vbNewLine & vbNewLine & _
'                  "The Paleomag Code must end now."
'
'        End
'
'    End If
'
'    'Check to see if the Boards have been imported
'    If Not ImportBoardsDone Then
'
'        'Boards have not been imported, so do so now
'        'Status Flag ImportBoardsDone will be set to True inside this subroutine
'        Get_BoardsFromIni
'
'    End If
'
'    'Now, use the Boards object collection method: .Find_Board, to find the desired board
'
'    Set TempBoard = SystemBoards.Find_Board(BoardININum, _
'                                            BoardNum, _
'                                            BoardNameStr)
'
'    'Return board object and set local board variable to Nothing
'    Set Retrieve_Board = TempBoard
'    Set TempBoard = Nothing
'
'End Function

' Sub Save_BoardsToINI()
'
' Created: April 4, 2010
'  Author: Isaac Hilburn
'
' Summary: Reads SystemBoards collection and overwrites each Board in the .ini
'          file with the Board objects in SystemBoards.  If the number of Boards in
'          the System Boards collection differs from the number of Boards in the
'          .ini file, the [Boards] section of the .ini file is completely deleted
'          before the System Boards are written to the .ini file
'
Public Sub Save_BoardsToINI()

    Dim success As Boolean
    Dim i As Long
    
    'Check to see if the number of Boards in the System Boards collection
    'matches the Board Count in the ini file
    NumBoards = val(Config_GetFromINI("Boards", _
                                      "BoardsCount", _
                                      "-1", _
                                      Prog_INIFile))
                                      
    If NumBoards = -1 Then
    
        'Ini file has not been updated with new format
        'Crap
                    
        'Create new BoardsCount field
        Config_SaveSetting "Boards", _
                           "BoardsCount", _
                           Trim(str(SystemBoards.Count))
        
        'All that is now needed is to save the
        'System Boards collection to the .ini file, which
        'the following code will do.
        
    End If
    
    If NumBoards <> SystemBoards.Count And NumBoards <> -1 Then
    
        'Need to delete the [Boards] section of .ini file
        success = IniFile.SectionDelete("Boards")
          
        If Not success Then
    
            'Raise some huge error
        
        End If
        
        'Then create new BoardsCount field
        Config_SaveSetting "Boards", _
                           "BoardsCount", _
                           Trim(str(SystemBoards.Count))
        
    End If
    
                                      
    For i = 1 To SystemBoards.Count

        Add_INIBoard SystemBoards(i), True
        
    Next i
    
    
End Sub
'Function Get_ChannelINIStr
'
' Created: May 6, 2010
'  Author: Isaac Hilburn
'
' Summary: Simple shell to search the Paleomag.ini file for the
'          specially formatted Channel INI ID string that's associated
'          with a named channel.
'          (i.e. Channel INI name = "ARMVoltageOut", corresponding
'           Channenl INI ID = AO-0-CH0 {Analog Out}-{INI Board # 0}-{Channel 0}
'
'   Input:
'
'   ChanININame -   String containing the special Channel INI name
'                   (i.e. "ARMVoltageOut","IRMCapacitorVoltageIn", etc.)
'
'   Output:
'
'       String -    String containing the Channel INI Id that is specific to
'                   one channel on one board in the [Boards] section of the
'                   Paleomag.ini file

'Sub Save_ChannelsToINI()
'
' Created: April 5, 2010
'  Author: Isaac Hilburn
'Evanston, WY
' Summary: Reads the global channel objects and overwrites the associated
'          fields in the [Channels] section of the .ini file with the needed
'          information
'
Public Sub Save_ChannelsToINI()

    Dim ChanStr As String
        
    'This function may be called before the global channels objects
    'have been setup
    'Need to check each channel to see if it is not Nothing
    'Note: For Each channel that we need to load, need to store the ChanStr from the ini
    '      file, then parse the string to link to the correct Board & Channel in the System
    '      Boards Collection
    
    
    'Save the ARM Voltage Analog Output Channel
    If Not ARMVoltageOut Is Nothing Then
    
        'Setup the Channel string
        With ARMVoltageOut
    
            ChanStr = Trim(.ChanType) & "-" & _
                      Trim(str(.BoardININum)) & _
                      "-CH" & Trim(str(.ChanNum))
        
        End With
        
        Config_SaveSetting "Channels", _
                           "ARMVoltageOut", _
                           ChanStr
                           
    End If
                                         
                                         
    'Save the IRM Capacitor Monitor Voltage Analog Input channel
    If Not IRMCapacitorVoltageIn Is Nothing Then
    
        'Setup the Channel string
        With IRMCapacitorVoltageIn
    
            ChanStr = Trim(.ChanType) & "-" & _
                      Trim(str(.BoardININum)) & _
                      "-CH" & Trim(str(.ChanNum))
        
        End With
        
        Config_SaveSetting "Channels", _
                           "IRMCapacitorVoltageIn", _
                           ChanStr
                           
    End If
    
    'Save the IRM Monitor Voltage Analog Input channel
    If Not IRMMonitor Is Nothing Then
    
        'Setup the Channel string
        With IRMMonitor
    
            ChanStr = Trim(.ChanType) & "-" & _
                      Trim(str(.BoardININum)) & _
                      "-CH" & Trim(str(.ChanNum))
        
        End With
        
        Config_SaveSetting "Channels", _
                           "IRMMonitor", _
                           ChanStr
                           
    End If
                                         
                                         
    'Save the IRM Peak Voltage Analog Output channel
    If Not IRMVoltageOut Is Nothing Then
    
        'Setup the Channel string
        With IRMVoltageOut
    
            ChanStr = Trim(.ChanType) & "-" & _
                      Trim(str(.BoardININum)) & _
                      "-CH" & Trim(str(.ChanNum))
        
        End With
        
        Config_SaveSetting "Channels", _
                           "IRMVoltageOut", _
                           ChanStr
                           
    End If
                                         
                                         
    'Save the AF Coil Temperature Sensor #1 Analog Input channel
    If Not AnalogT1 Is Nothing Then
    
        'Setup the Channel string
        With AnalogT1
    
            ChanStr = Trim(.ChanType) & "-" & _
                      Trim(str(.BoardININum)) & _
                      "-CH" & Trim(str(.ChanNum))
        
        End With
        
        Config_SaveSetting "Channels", _
                           "AnalogT1", _
                           ChanStr
                           
    End If
                                         
                                         
    'Save the AF Coil Temperature Sensor #2 Analog Input channel
    If Not AnalogT2 Is Nothing Then
    
        'Setup the Channel string
        With AnalogT2
    
            ChanStr = Trim(.ChanType) & "-" & _
                      Trim(str(.BoardININum)) & _
                      "-CH" & Trim(str(.ChanNum))
        
        End With
        
        Config_SaveSetting "Channels", _
                           "AnalogT2", _
                           ChanStr
                           
    End If
    
    
    'Save the AF Axial Coil Relay Digital Output channel
    If Not AFAxialRelay Is Nothing Then
    
        'Setup the Channel string
        With AFAxialRelay
    
            ChanStr = Trim(.ChanType) & "-" & _
                      Trim(str(.BoardININum)) & _
                      "-CH" & Trim(str(.ChanNum))
        
        End With
        
        Config_SaveSetting "Channels", _
                           "AFAxialRelay", _
                           ChanStr
                           
    End If
    
    
    'Save the AF Transverse Coil Relay Digital Output channel
    If Not AFTransRelay Is Nothing Then
    
        'Setup the Channel string
        With AFTransRelay
    
            ChanStr = Trim(.ChanType) & "-" & _
                      Trim(str(.BoardININum)) & _
                      "-CH" & Trim(str(.ChanNum))
        
        End With
        
        Config_SaveSetting "Channels", _
                           "AFTransRelay", _
                           ChanStr
                           
    End If
    
    
    'Save the IRM Low-Field Relay Digital Output channel
    If Not IRMRelay Is Nothing Then
    
        'Setup the Channel string
        With IRMRelay
    
            ChanStr = Trim(.ChanType) & "-" & _
                      Trim(str(.BoardININum)) & _
                      "-CH" & Trim(str(.ChanNum))
        
        End With
        
        Config_SaveSetting "Channels", _
                           "IRMRelay", _
                           ChanStr
                           
    End If
                                   
  
    'Save the ARM Set TTL Digital Output channel
    If Not ARMSet Is Nothing Then
    
        'Setup the Channel string
        With ARMSet
    
            ChanStr = Trim(.ChanType) & "-" & _
                      Trim(str(.BoardININum)) & _
                      "-CH" & Trim(str(.ChanNum))
        
        End With
        
        Config_SaveSetting "Channels", _
                           "ARMSet", _
                           ChanStr
                           
    End If
    
    
    'Save the IRM Fire TTL Digital Output channel
    If Not IRMFire Is Nothing Then
    
        'Setup the Channel string
        With IRMFire
    
            ChanStr = Trim(.ChanType) & "-" & _
                      Trim(str(.BoardININum)) & _
                      "-CH" & Trim(str(.ChanNum))
        
        End With
        
        Config_SaveSetting "Channels", _
                           "IRMFire", _
                           ChanStr
                           
    End If
    
    
    'Save the IRM Trim TTL Digital Output channel
    If Not IRMTrim Is Nothing Then
    
        'Setup the Channel string
        With IRMTrim
    
            ChanStr = Trim(.ChanType) & "-" & _
                      Trim(str(.BoardININum)) & _
                      "-CH" & Trim(str(.ChanNum))
        
        End With
        
        Config_SaveSetting "Channels", _
                           "IRMTrim", _
                           ChanStr
                           
    End If
    
    
    'Save the IRM Ready Status Digital Input channel
    If Not IRMPowerAmpVoltageIn Is Nothing Then
    
        'Setup the Channel string
        With IRMPowerAmpVoltageIn
    
            ChanStr = Trim(.ChanType) & "-" & _
                      Trim(str(.BoardININum)) & _
                      "-CH" & Trim(str(.ChanNum))
        
        End With
        
        Config_SaveSetting "Channels", _
                           "IRMPowerAmpVoltageIn", _
                           ChanStr
                           
    End If
    
    
    'Save the Vacuum Motor Toggle Digital Output channel
    If Not MotorToggle Is Nothing Then
    
        'Setup the Channel string
        With MotorToggle
    
            ChanStr = Trim(.ChanType) & "-" & _
                      Trim(str(.BoardININum)) & _
                      "-CH" & Trim(str(.ChanNum))
        
        End With
        
        Config_SaveSetting "Channels", _
                           "MotorToggle", _
                           ChanStr
                           
    End If
    
        
    'Save the VacuumToggle A Digital Output channel
    If Not VacuumToggleA Is Nothing Then
    
        'Setup the Channel string
        With VacuumToggleA
    
            ChanStr = Trim(.ChanType) & "-" & _
                      Trim(str(.BoardININum)) & _
                      "-CH" & Trim(str(.ChanNum))
        
        End With
        
        Config_SaveSetting "Channels", _
                           "VacuumToggleA", _
                           ChanStr
                           
    End If
    
    'Save the Degausser Cooler Digital Output channel
    If Not DegausserToggle Is Nothing Then
    
        'Setup the Channel string
        With DegausserToggle
    
            ChanStr = Trim(.ChanType) & "-" & _
                      Trim(str(.BoardININum)) & _
                      "-CH" & Trim(str(.ChanNum))
        
        End With
        
        Config_SaveSetting "Channels", _
                           "DegausserToggle", _
                           ChanStr
                           
    End If
    
'   (Commented out May 2010, I Hilburn - Vacuum Toggle B is no longer used in the
'    paleomag code)
'
'    'Save the VacuumToggle B Digital Output channel
'    If Not VacuumToggleB Is Nothing Then
'
'        'Setup the Channel string
'        With VacuumToggleB
'
'            ChanStr = Trim(.ChanType) & "-" & _
'                      Trim(Str(.BoardININum)) & _
'                      "-CH" & Trim(Str(.ChanNum))
'
'        End With
'
'        Config_SaveSetting "Channels", _
'                           "VacuumToggleB", _
'                           ChanStr
'
'    End If
    
End Sub

'Public sub-routine Save_WaveFormsToIni()
'
' Created: April 5, 2010
'  Author: Isaac Hilburn
'
' Summary: Takes Wave objects into the global Wave Forms collection and
'          overwrites the [WaveForms] section in the .ini file with
'          the values from the Wave Objects.  If the number of waves in
'          the WaveForms collection and the WaveFormCount field of the .ini
'          disagree, the [WaveForms] section of the .ini file will be deleted
'          before the Wave objects are written to the .ini file
'
Public Sub Save_WaveFormsToIni()

    'Like Get_BoardsFromIni, the number associated with each wave-form is appended onto the
    'string key name in the Paleomag.ini file so that each key is unique to each wave.
    '
    'This allows us to use the pre-existing Config_GetFromIni function that Bob Kopp
    'built.
    
    Dim i As Long
    Dim N As Long
    Dim TempL As Long
    Dim BoardStr As String
    Dim ChanStr As String
    Dim UserResp As Long
    
    'Read in the number of Wave objects whose settings are stored in
    'the Paleomag.ini file
    N = val(Config_GetFromINI("WaveForms", _
                              "WaveFormCount", _
                              "-1", _
                              Prog_INIFile))
    
    'Check to see if the WaveForm object is accessible
    On Error Resume Next
    
        TempL = WaveForms.Count
        
        If Err.number <> 0 Then
        
            'Wha-oh, the WaveForms collection hasn't been loaded yet
            'Ask the user for input
            UserResp = MsgBox("There currently are no WaveForms loaded in the Paleomag program." & _
                              vbNewLine & vbNewLine & _
                              "Do you wish to delete all the wave forms stored in the .ini file?" & _
                              vbNewLine & _
                              "(Number of wave forms in the .ini file = " & Trim(str(N)), _
                              vbYesNo, _
                              "Delete .Ini File Wave Forms?")
            
            If UserResp = 6 Then
                'User has selected "yes"
                'Delete the [WaveForms] section of the .ini file
                
                IniFile.SectionClear "WaveForms", True
                
                'Save new # of Waves to INI file
                Config_SaveSetting "WaveForms", _
                                   "WaveFormCount", _
                                   "0"
                                    
                'Set the No INI wave forms flag
                NoINIWaveForms = True
            
            End If
            
            'Exit the sub
            Exit Sub
    
        End If
        
    On Error GoTo 0
    
    'Second error check for the WaveForms being empty
    If WaveForms Is Nothing Then
    
        'Wha-oh, the WaveForms collection hasn't been loaded yet
        'Ask the user for input
        UserResp = MsgBox("There currently are no WaveForms loaded in the Paleomag program." & _
                          vbNewLine & vbNewLine & _
                          "Do you wish to delete all the wave forms stored in the .ini file?" & _
                          vbNewLine & _
                          "(Number of wave forms in the .ini file = " & Trim(str(N)), _
                          vbYesNo, _
                          "Delete .Ini File Wave Forms?")
        
        If UserResp = 6 Then
            'User has selected "yes"
            'Delete the [WaveForms] section of the .ini file
            
            IniFile.SectionClear "WaveForms", True
            
            'Save new # of Waves to INI file
            Config_SaveSetting "WaveForms", _
                               "WaveFormCount", _
                               "0"
        
            'Set the No INI wave forms flag
            NoINIWaveForms = True
        
        End If
            
        'Exit the sub
        Exit Sub
        
    'Now check for WaveForms not containing any Wave objects
    ElseIf WaveForms.Count = 0 Then
    
        'Save condition as above
        'Ask the user for input
        UserResp = MsgBox("There currently are no WaveForms loaded in the Paleomag program." & _
                          vbNewLine & vbNewLine & _
                          "Do you wish to delete all the wave forms stored in the .ini file?" & _
                          vbNewLine & _
                          "(Number of wave forms in the .ini file = " & Trim(str(N)), _
                          vbYesNo, _
                          "Delete .Ini File Wave Forms?")
        
        If UserResp = 6 Then
            'User has selected "yes"
            'Delete the [WaveForms] section of the .ini file
            IniFile.SectionClear "WaveForms", True
            
            'Set the Number of waveforms in the INI file to 0
            Config_SaveSetting "WaveForms", _
                               "WaveFormCount", _
                               "0"
                                
            'Set the No INI wave forms flag
            NoINIWaveForms = True
        
        End If
            
        'Exit the sub
        Exit Sub
        
    ElseIf N <> WaveForms.Count Then
    
        'Overwrite the INI file with the new set of wave-forms in
        'the System's collection
    
        'Delete the [WaveForms] section of the .ini file
        IniFile.SectionClear "WaveForms", True
    
        Config_SaveSetting "WaveForms", _
                            "WaveFormCount", _
                            Trim(str(WaveForms.Count))
         
        'Reset N
        N = WaveForms.Count
    
    ElseIf N = WaveForms.Count Then
    
        'Still need to overwrite the INI File
        'Can't assume that the program wave-forms
        'are the same as the .INI wave-forms
        
        'Delete the [WaveForms] section of the .ini file
        IniFile.SectionClear "WaveForms", True
    
        Config_SaveSetting "WaveForms", _
                            "WaveFormCount", _
                            Trim(str(WaveForms.Count))
    
    End If
                         
    'there is at least one waveform to load into
    'the global WaveForms collection
    
    'Now go and sequentially add each Wave in WaveForms to the .ini file
    For i = 1 To N
    
        Add_IniWaveForm WaveForms(i), True
        
    Next i

    'Set no INI wave forms flag to false
    NoINIWaveForms = False

End Sub

Public Sub SaveCoilTuningParam(ByVal CoilParam As SaveCoilParam, _
                               Optional ByVal AFCoilSystem As Integer = -128)
                               
    'Check for no coil system input
    If AFCoilSystem = -128 Then AFCoilSystem = ActiveCoilSystem
                              
    'Now select which coil to change the INI settings for
    If AFCoilSystem = AxialCoilSystem Then
    
        'Coil parameter belongs to the Axial coil
    
        If CoilParam = resFreq Then
        
            'Save the Axial Res Freq
            Config_SaveSetting "AFAxial", _
                               "AFAxialResFreq", _
                                Trim(str(AfAxialResFreq))

        ElseIf CoilParam = VoltsMax Then
        
            'Save Axial max monitor and ramp voltages
            Config_SaveSetting "AFAxial", _
                               "AFAxialMonMax", _
                               Trim(str(AfAxialMonMax))
            Config_SaveSetting "AFAxial", _
                               "AFAxialRampMax", _
                               Trim(str(AfAxialRampMax))
                               
        End If
        
    ElseIf AFCoilSystem = TransverseCoilSystem Then
        
        'Coil parameter belongs to the Transverse coil
        
        If CoilParam = resFreq Then
        
            'Save Trans Res Freq
            Config_SaveSetting "AFTrans", _
                               "AFTransResFreq", _
                               Trim(str(AfTransResFreq))
                               
        ElseIf CoilParam = VoltsMax Then
        
            'Save Trans max monitor and ramp voltages
            Config_SaveSetting "AFTrans", _
                               "AFTransMonMax", _
                               Trim(str(AfTransMonMax))
            Config_SaveSetting "AFTrans", _
                               "AFTransRampMax", _
                               Trim(str(AfTransRampMax))
                               
        End If
        
    End If
                               
End Sub

Public Function String_SizeInBytes(ByVal sString As String) As Long

    'The size in bytes of a string in Visual Basic 6 is equal to:
    '   Character length of the String * 2
    '   + 4 bytes for the string pre-fix
    '   + 2 bytes for the string terminator
    
    String_SizeInBytes = 2 * Len(sString) + 6
    
End Function

