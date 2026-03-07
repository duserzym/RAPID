Attribute VB_Name = "modConfig"
' This module is included to handle all of the initial
' settings for both paleomag and biomag magnetometers.
' DLL declarations
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long



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
Public Prog_UsageFile     As String
Public Prog_DefaultBackup As String
Public Prog_DefaultPath   As String
Public Prog_LogoFile      As String
Public Prog_IcoFile       As String ' (October 2007 L Carporzen)
Public Prog_TextEditor    As String
Public Prog_HelpURLRoot   As String
Public DumpRawDataStats   As Boolean
Public LogMessages        As Boolean
' Calibration variables read from file
Public ZCal         As Double
Public XCal         As Double
Public YCal         As Double
Public IRMPos       As Long
Public IRMHiPos     As Long
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
Public CmdHomeToTop             As Integer
Public CmdSamplePickup          As Integer
Public MotorIDTurning           As Integer
Public MotorIDChanger           As Integer
Public MotorIDUpDown            As Integer
Public SCurveFactor             As Integer
Public TurningMotorFullRotation As Long
Public TurningMotor1rps         As Long
Public UpDownMotor1cm           As Double
Public TrayOffsetAngle          As Double
Public UpDownTorqueFactor       As Integer
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
Public Meascount       As Integer ' Count the number of repeated measurements which don't pass the tests

'New AF demag settings (Mar 2010 - I Hilburn)
'Part of implementation of new AF DAQ coding module
Public AFSystem As AFSystemObj

' Now the Af demag calibration values
Public AFDelay          As Integer
Public AFRampRate       As Integer
Public AFWait           As Double
Public TSlope           As Double
Public Toffset          As Double
Public Thot             As Integer
Public Tmax             As Integer
Public AfAxialCoord     As String
Public AfTransCoord     As String

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
Public IRMHFAxis          As String
Public IRMLFAxis          As String
Public IRMLFBackfieldAxis As String
Public PulseVoltMax       As Double

'(March 2010 - I Hilburn)
'Modified IRM pulse field calibration arrays to make dynamic and N x 2 in dimension
'PulseLFCount & PulseHFCount = number of field calibration values in .INI file to
'be loaded to the pulse arrays
Public PulseLF()       As Double
Public PulseHF()       As Double
Public PulseLFMax         As Double
Public PulseLFMin         As Double
Public PulseHFMax         As Double
Public PulseHFMin         As Double

' now susceptibility variables
Public SusceptibilityMomentFactorCGS As Double
Public SusceptibilityScaleFactor     As Double
Public SusceptibilitySettings As String

' MCC volt conversion converts a charging volt
' to a MCC output voltage (default: 10 V MCC -> 450 V)
Public PulseMCCVoltConversion       As Double
Public PulseReturnMCCVoltConversion As Double

'System DAQ Boards Collection Declaration
Public SystemBoards As Boards

'Boolean Status Flag to indicate whether the .ini file board settings
'have been imported into the System Boards global collection
'This will be reset to false when .ini file is updated with new
'board settings that aren't contained in the System Boards global collection
Public ImportBoardsDone As Boolean

'System Wave Forms Collection Declaration
Public WaveForms As Waves

' Now the ARM calibration variables
Public ARMMax           As Double
Public ARMVoltGauss     As Double
Public ARMVoltMax       As Double
Public ARMTimeMax       As Double
Public DoVacuumReset    As Boolean

' (March 2008 - L Carporzen) Put in Settings the IRM/ARM channels (MIT acquisition board does not work on IRMTrim = 3
' Analog channel output
'(March 2010 - I Hilburn) Changed Integer channel/port numbs to Channel objects
Public ARMVoltageOut As Channel
Public IRMVoltageOut  As Channel

' Analog input
'(March 2010 - I Hilburn) Changed Integer chan/port number to Channel object
Public IRMCapacitorVoltageIn  As Channel

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
Public IRMReady  As Channel
Public MotorToggle As Channel
Public VacuumToggleA As Channel
Public VacuumToggleB As Channel

' Now assign the COMM Ports to the proper lines!
Public COMPortSquids    As Integer
Public COMPortAf        As Integer
Public COMPortUpDown    As Integer
Public COMPortTurning   As Integer
Public COMPortChanger   As Integer
Public COMPortVacuum    As Integer
Public COMPortSusceptibility  As Integer

'VB Send Mail settings
Public MailSMTPHost           As String
Public MailFrom               As String
Public MailFromName           As String
Public MailCCList             As String
Public MailStatusMonitor      As String

'AF Field Calibration Arrays
'(Mar 2010 - changed from 25 elements, fixed, to dynamic arrays with N x 2 dimensions)
Public AFAxial()           As Double
Public AFTrans()           As Double
Public AFAxialCount         As Long '(March 2010 - I Hilburn) - Enables variable number of AF field calibration values for Axial Coil
Public AFTransCount         As Long '(March 2010 - I Hilburn) - Enables variable number of AF field calibration values for Transverse Coil

'IRM Field Calibration Arrays

'Module Settings
Public EnableIRM              As Boolean
Public EnableIRMHi            As Boolean
Public EnableARM              As Boolean
Public EnableAF               As Boolean

'Use Temperature Sensors - boolean status flag
'(Mar, 2010 - L Carporzen)
Public EnableT1               As Boolean
Public EnableT2               As Boolean

Public EnableSusceptibility   As Boolean
Public EnableIRMBackfield     As Boolean
Public EnableIRMReturn        As Boolean

Public Enum AFSystemObj

    TwoG = 0
    MCC = 1
    ADWIN = 2

End Enum

Public Sub Config_ReadINISettings()
    ' This procedure reads settings from a file and adjusts
    ' variables in memory accordingly. (c:\paleomag\paleomag.ini)
    Dim i As Integer
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
    SlotMin = Int(val(Config_GetFromINI("SampleChanger", "SlotMin", "1", Prog_INIFile)))
    SlotMax = Int(val(Config_GetFromINI("SampleChanger", "SlotMax", "200", Prog_INIFile)))
    HoleSlotNum = Int(val(Config_GetFromINI("SampleChanger", "HoleSlotNum", "10", Prog_INIFile)))
    OneStep = val(Config_GetFromINI("SampleChanger", "OneStep", "-1010.1010101", Prog_INIFile))
    ZeroPos = val(Config_GetFromINI("SteppingMotor", "ZeroPos", "-25886", Prog_INIFile))
    MeasPos = val(Config_GetFromINI("SteppingMotor", "MeasPos", "-30607", Prog_INIFile))
    AFPos = val(Config_GetFromINI("SteppingMotor", "AFPos", "-8405", Prog_INIFile))
    IRMPos = val(Config_GetFromINI("SteppingMotor", "IRMPos", "-8405", Prog_INIFile))
    IRMHiPos = val(Config_GetFromINI("SteppingMotor", "IRMHiPos", "-8405", Prog_INIFile))
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
    ChangerSpeed = val(Config_GetFromINI("SteppingMotor", "ChangerSpeed", "40000000", Prog_INIFile))
    TurnerSpeed = val(Config_GetFromINI("SteppingMotor", "TurnerSpeed", "2000000", Prog_INIFile))
    SCurveFactor = val(Config_GetFromINI("SteppingMotor", "SCurveFactor", "32767", Prog_INIFile))
    TurningMotorFullRotation = val(Config_GetFromINI("SteppingMotor", "TurningMotorFullRotation", "2000", Prog_INIFile))
    TurningMotor1rps = val(Config_GetFromINI("SteppingMotor", "TurningMotor1rps", "16000000", Prog_INIFile))
    UpDownMotor1cm = val(Config_GetFromINI("SteppingMotor", "UpDownMotor1cm", "10", Prog_INIFile))
    UpDownTorqueFactor = Int(val(Config_GetFromINI("SteppingMotor", "UpDownTorqueFactor", "40", Prog_INIFile)))
    PickupTorqueThrottle = val(Config_GetFromINI("SteppingMotor", "PickupTorqueThrottle", "0.4", Prog_INIFile))
    TrayOffsetAngle = val(Config_GetFromINI("SteppingMotor", "TrayOffsetAngle", "0", Prog_INIFile))
    CmdHomeToTop = val(Config_GetFromINI("MotorPrograms", "CmdHomeToTop", "206", Prog_INIFile))
    CmdSamplePickup = val(Config_GetFromINI("MotorPrograms", "CmdSamplePickup", "241", Prog_INIFile))
    MotorIDTurning = val(Config_GetFromINI("MotorPrograms", "MotorIDTurning", "16", Prog_INIFile))
    MotorIDChanger = val(Config_GetFromINI("MotorPrograms", "MotorIDChanger", "16", Prog_INIFile))
    MotorIDUpDown = val(Config_GetFromINI("MotorPrograms", "MotorIDUpDown", "16", Prog_INIFile))
    ' Now Magnetometer Calibration Constants
    ZCal = val(Config_GetFromINI("MagnetometerCalibration", "ZCal", "-2.516", Prog_INIFile))
    XCal = val(Config_GetFromINI("MagnetometerCalibration", "XCal", "-3.410", Prog_INIFile))
    YCal = val(Config_GetFromINI("MagnetometerCalibration", "YCal", "-3.470", Prog_INIFile))
    RangeFact = val(Config_GetFromINI("MagnetometerCalibration", "RangeFact", "0.00001", Prog_INIFile))
    ReadDelay = val(Config_GetFromINI("MagnetometerCalibration", "ReadDelay", "1", Prog_INIFile)) ' (March 2008 L Carporzen) Read delay
    RemeasureCSDThreshold = val(Config_GetFromINI("Magnetometery", "RemeasureCSDThreshold", "8", Prog_INIFile))
    ' New selections in the Options menu (April-May 2007 L Carporzen)
    JumpThreshold = val(Config_GetFromINI("Magnetometery", "JumpThreshold", "0.1", Prog_INIFile))
    StrongMom = val(Config_GetFromINI("Magnetometery", "StrongMom", "0.02", Prog_INIFile))
    IntermMom = val(Config_GetFromINI("Magnetometery", "IntermMom", "0.000001", Prog_INIFile))
    MomMinForRedo = val(Config_GetFromINI("Magnetometery", "MomMinForRedo", "0.000000008", Prog_INIFile))
    JumpSensitivity = val(Config_GetFromINI("Magnetometery", "JumpSensitivity", "1", Prog_INIFile))
    NbTry = val(Config_GetFromINI("Magnetometery", "NbTry", "5", Prog_INIFile))
    Meascount = 1
    ' now the susceptibility factor
    SusceptibilityMomentFactorCGS = val(Config_GetFromINI("SusceptibilityCalibration", "SusceptibilityMomentFactorCGS", "10", Prog_INIFile))
    SusceptibilityScaleFactor = val(Config_GetFromINI("SusceptibilityCalibration", "SusceptibilityScaleFactor", "1", Prog_INIFile))
    ' Now assign AF axes
    AFDelay = val(Config_GetFromINI("AF", "AFDelay", "1", Prog_INIFile))
    AFRampRate = val(Config_GetFromINI("AF", "AFRampRate", "3", Prog_INIFile))
    AFWait = Config_GetFromINI("AF", "AFWait", "90", Prog_INIFile)
    TSlope = Config_GetFromINI("AF", "TSlope", "58.86", Prog_INIFile)
    Toffset = Config_GetFromINI("AF", "Toffset", "289.6", Prog_INIFile)
    Thot = Config_GetFromINI("AF", "Thot", "40", Prog_INIFile)
    Tmax = Config_GetFromINI("AF", "Tmax", "50", Prog_INIFile)
    AfAxialCoord = Config_GetFromINI("AFAxial", "AFAxialCoord", "X", Prog_INIFile)
    AfTransCoord = Config_GetFromINI("AFTrans", "AFTransCoord", "Y", Prog_INIFile)
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
    ' Now the IRM Pulse coil calibration numbers (CIT Lowenstam only)
    IRMHFAxis = Config_GetFromINI("IRMPulse", "IRMHFAxis", "Z", Prog_INIFile)
    IRMLFAxis = Config_GetFromINI("IRMPulse", "IRMLFAxis", "X", Prog_INIFile)
    IRMLFBackfieldAxis = Config_GetFromINI("IRMPulse", "IRMLFBackfieldAxis", "Y", Prog_INIFile)
    PulseHFMax = val(Config_GetFromINI("IRMPulseHF", "PulseHFMax", "13080", Prog_INIFile))
    PulseHFMin = val(Config_GetFromINI("IRMPulseHF", "PulseHFMin", "50", Prog_INIFile))
    PulseLFMax = val(Config_GetFromINI("IRMPulse", "PulseLFMax", "13080", Prog_INIFile))
    PulseLFMin = val(Config_GetFromINI("IRMPulse", "PulseLFMin", "50", Prog_INIFile))
    PulseMCCVoltConversion = val(Config_GetFromINI("IRMPulse", "PulseMCCVoltConverstion", "0.022222", Prog_INIFile))
    PulseVoltMax = val(Config_GetFromINI("IRMPulse", "PulseVoltMax", "10", Prog_INIFile))
    PulseReturnMCCVoltConversion = val(Config_GetFromINI("IRMPulse", "PulseReturnMCCVoltConverstion", "0.022222", Prog_INIFile))
    ' Now the ARM calibration values as well
    ARMMax = val(Config_GetFromINI("ARM", "ARMMax", "20", Prog_INIFile))
    ARMVoltGauss = val(Config_GetFromINI("ARM", "ARMVoltGauss", "0.1033", Prog_INIFile))
    ARMVoltMax = val(Config_GetFromINI("ARM", "ARMVoltMax", "2.0", Prog_INIFile))
    ARMTimeMax = val(Config_GetFromINI("ARM", "ARMTimeMax", "600", Prog_INIFile))
    
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
'    IRMReady = val(Config_GetFromINI("IRM-ARM", "IRMReady", "4", Prog_INIFile))
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

    'Import DAQ Board(s) settings into Board objects in the SystemBoards collection
    Import_Boards
    
    'Import DAQ Board Channel assignments - assigned to channel objects that
    'replaced the Integer port-number assignments under the old implementation
    Import_Channels
    
'------------------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------------------

'------------------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------------------
'
'   Major Modification for importing AF settings
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

    'Import information for loading into the necessary wave objects (load up WaveForms collection)
    Import_WaveForms

'------------------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------------------

    ' Now the vacuum options
    DoVacuumReset = (Config_GetFromINI("Vacuum", "DoVacuumReset", "False", Prog_INIFile) = "True")
    ' Now assign the COMM Ports to the proper lines!
    COMPortSquids = Config_GetFromINI("COMPorts", "COMPortSquids", "10", Prog_INIFile)
    COMPortAf = Config_GetFromINI("COMPorts", "COMPortAf", "9", Prog_INIFile)
    COMPortUpDown = Config_GetFromINI("COMPorts", "COMPortUpDown", "4", Prog_INIFile)
    COMPortTurning = Config_GetFromINI("COMPorts", "COMPortTurning", "5", Prog_INIFile)
    COMPortChanger = Config_GetFromINI("COMPorts", "COMPortChanger", "6", Prog_INIFile)
    COMPortVacuum = Config_GetFromINI("COMPorts", "COMPortVacuum", "3", Prog_INIFile)
    COMPortSusceptibility = Config_GetFromINI("COMPorts", "COMPortSusceptibility", "8", Prog_INIFile)
    SusceptibilitySettings = Config_GetFromINI("COMPorts", "SusceptibilitySettings", "9600,N,8,2", Prog_INIFile)
    ' Now settings for mailer application
    MailSMTPHost = Config_GetFromINI("Email", "MailSMTPHost", vbNullString, Prog_INIFile)
    MailFrom = Config_GetFromINI("Email", "MailFrom", "paleomag@localhost", Prog_INIFile)
    MailFromName = Config_GetFromINI("Email", "MailFromName", "2G Magnetometer Sample Changer", Prog_INIFile)
    MailCCList = Config_GetFromINI("Email", "MailCCList", vbNullString, Prog_INIFile)
    MailStatusMonitor = Config_GetFromINI("Email", "MailStatusMonitor", vbNullString, Prog_INIFile)

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
    AFAxialCount = Config_GetFromINI("AFAxial", "AFAxialCount", "0", Prog_INIFile)
    AFTransCount = Config_GetFromINI("AFTrans", "AFTransCount", "0", Prog_INIFile)
    
    'Redimension AFAxial and AFTrans arrays (Remember to add one point for the zeros)
    ReDim AFAxial(AFAxialCount + 1, 2)
    ReDim AFTrans(AFTransCount + 1, 2)
    
    AFAxial(0, 0) = 0
    AFAxial(0, 1) = 0
    AFTrans(0, 0) = 0
    AFTrans(0, 1) = 0
    
    'Run through Axial Coil calibration points stored in the INI file
    For i = 1 To AFAxialCount
        
        AFAxial(i, 0) = val(Config_GetFromINI("AFAxial", "AFAxialX" & Format$(i, "0"), "0", Prog_INIFile))
        AFAxial(i, 1) = val(Config_GetFromINI("AFAxial", "AFAxialY" & Format$(i, "0"), "0", Prog_INIFile))
    
    Next i
    
    'Run through Axial Coil calibration points stored in the INI file
    For i = 1 To AFTransCount
        
        AFTrans(i, 0) = val(Config_GetFromINI("AFTrans", "AFTransX" & Format$(i, "0"), "0", Prog_INIFile))
        AFTrans(i, 1) = val(Config_GetFromINI("AFTrans", "AFTransY" & Format$(i, "0"), "0", Prog_INIFile))
    
    Next i
    
        
    
    ' IRM Pulse Field Calibration Arrays
    PulseLFCount = Config_GetFromINI("IRMPulse", "PulseLFCount", "0", Prog_INIFile)
    PulseHFCount = Config_GetFromINI("IRMPulse", "PulseHFCount", "0", Prog_INIFile)
    
    'Redimension AFAxial and AFTrans arrays (Remember to add one point for the zeros)
    ReDim PulseLF(PulseLFCount + 1, 2)
    ReDim PulseHF(PulseHFCount + 1, 2)
    
    PulseLF(0, 0) = 0
    PulseLF(0, 1) = 0
    PulseHF(0, 0) = 0
    PulseHF(0, 1) = 0
    
    If PulseLFCount > 0 Then
    
        'Run through Axial Coil calibration points stored in the INI file
        For i = 1 To PulseLFCount
            
            PulseLF(i, 0) = val(Config_GetFromINI("IRMPulse", "PulseLFX" & Format$(i, "0"), "0", Prog_INIFile))
            PulseLF(i, 1) = val(Config_GetFromINI("IRMPulse", "PulseLFY" & Format$(i, "0"), "0", Prog_INIFile))
        
        Next i
        
    End If
    
    If PulseHFCount > 0 Then
    
        'Run through Axial Coil calibration points stored in the INI file
        For i = 1 To PulseHFCount
            
            PulseHF(i, 0) = val(Config_GetFromINI("IRMPulse", "PulseHFX" & Format$(i, "0"), "0", Prog_INIFile))
            PulseHF(i, 1) = val(Config_GetFromINI("IRMPulse", "PulseHFY" & Format$(i, "0"), "0", Prog_INIFile))
        
        Next i
    
    End If
    
'------------------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------------------
    
    ' now for the rockmag enable/disable modules
    EnableIRM = (Config_GetFromINI("Modules", "EnableIRM", "True", Prog_INIFile) = "True")
    EnableIRMHi = (Config_GetFromINI("Modules", "EnableIRMHi", "False", Prog_INIFile) = "True")
    EnableIRMBackfield = (Config_GetFromINI("Modules", "EnableIRMBackfield", "False", Prog_INIFile) = "True")
    EnableIRMReturn = (Config_GetFromINI("Modules", "EnableIRMReturn", "False", Prog_INIFile) = "True")
    EnableARM = (Config_GetFromINI("Modules", "EnableARM", "True", Prog_INIFile) = "True")
    EnableAF = (Config_GetFromINI("Modules", "EnableAF", "True", Prog_INIFile) = "True")
    EnableT1 = (Config_GetFromINI("Modules", "EnableT1", "False", Prog_INIFile) = "True")
    EnableT2 = (Config_GetFromINI("Modules", "EnableT2", "False", Prog_INIFile) = "True")
    EnableSusceptibility = (Config_GetFromINI("Modules", "EnableSusceptibility", "True", Prog_INIFile) = "True")
End Sub

Public Sub Config_writeSettingstoINI()
    Dim i As Integer
    FileCopy Prog_INIFile, Prog_INIFile + ".bak"
    Config_SaveSetting "SampleChanger", "SlotMin", Str$(SlotMin)
    Config_SaveSetting "SampleChanger", "SlotMax", Str$(SlotMax)
    Config_SaveSetting "SampleChanger", "OneStep", Str$(OneStep)
    Config_SaveSetting "SampleChanger", "HoleSlotNum", Str$(HoleSlotNum)
    Config_SaveSetting "SteppingMotor", "ZeroPos", Str$(ZeroPos)
    Config_SaveSetting "SteppingMotor", "MeasPos", Str$(MeasPos)
    Config_SaveSetting "SteppingMotor", "IRMPos", Str$(IRMPos)
    Config_SaveSetting "SteppingMotor", "IRMHiPos", Str$(IRMHiPos)
    Config_SaveSetting "SteppingMotor", "AFPos", Str$(AFPos)
    Config_SaveSetting "SteppingMotor", "SCoilPos", Str$(SCoilPos)
    Config_SaveSetting "SteppingMotor", "FloorPos", Str$(FloorPos)
    Config_SaveSetting "SteppingMotor", "MinUpDownPos", Str$(MinUpDownPos)
    Config_SaveSetting "SteppingMotor", "SampleBottom", Str$(SampleBottom)
    Config_SaveSetting "SteppingMotor", "SampleTop", Str$(SampleTop)
    Config_SaveSetting "SteppingMotor", "SampleHoleAlignmentOffset", Str$(SampleHoleAlignmentOffset)
    Config_SaveSetting "SteppingMotor", "LiftSpeedSlow", Str$(LiftSpeedSlow)
    Config_SaveSetting "SteppingMotor", "LiftSpeedNormal", Str$(LiftSpeedNormal)
    Config_SaveSetting "SteppingMotor", "LiftSpeedFast", Str$(LiftSpeedFast)
    Config_SaveSetting "SteppingMotor", "TurnerSpeed", Str$(TurnerSpeed)
    Config_SaveSetting "SteppingMotor", "ChangerSpeed", Str$(ChangerSpeed)
    Config_SaveSetting "SteppingMotor", "SCurveFactor", Str$(SCurveFactor)
    Config_SaveSetting "SteppingMotor", "TurningMotorFullRotation", Str$(TurningMotorFullRotation)
    Config_SaveSetting "SteppingMotor", "TurningMotor1rps", Str$(TurningMotor1rps)
    Config_SaveSetting "SteppingMotor", "UpDownMotor1cm", Str$(UpDownMotor1cm)
    Config_SaveSetting "SteppingMotor", "UpDownTorqueFactor", Str$(UpDownTorqueFactor)
    Config_SaveSetting "SteppingMotor", "PickupTorqueThrottle", Str$(PickupTorqueThrottle)
    Config_SaveSetting "SteppingMotor", "TrayOffsetAngle", Str$(TrayOffsetAngle)
    Config_SaveSetting "MagnetometerCalibration", "ZCal", Str$(ZCal)
    Config_SaveSetting "MagnetometerCalibration", "XCal", Str$(XCal)
    Config_SaveSetting "MagnetometerCalibration", "YCal", Str$(YCal)
    Config_SaveSetting "MagnetometerCalibration", "RangeFact", Str$(RangeFact)
    Config_SaveSetting "MagnetometerCalibration", "ReadDelay", Str$(ReadDelay) ' (March 2008 L Carporzen) Read delay
    Config_SaveSetting "Magnetometery", "RemeasureCSDThreshold", Str$(RemeasureCSDThreshold)
    ' New selections in the Options menu (April-May 2007 L Carporzen)
    Config_SaveSetting "Magnetometery", "JumpThreshold", Str$(JumpThreshold)
    Config_SaveSetting "Magnetometery", "StrongMom", Str$(StrongMom)
    Config_SaveSetting "Magnetometery", "IntermMom", Str$(IntermMom)
    Config_SaveSetting "Magnetometery", "MomMinForRedo", Str$(MomMinForRedo)
    Config_SaveSetting "Magnetometery", "JumpSensitivity", Str$(JumpSensitivity)
    Config_SaveSetting "Magnetometery", "NbTry", Str$(NbTry)
    Config_SaveSetting "SusceptibilityCalibration", "SusceptibilityMomentFactorCGS", Str$(SusceptibilityMomentFactorCGS)
    Config_SaveSetting "SusceptibilityCalibration", "SusceptibilityScaleFactor", Str$(SusceptibilityScaleFactor)

    Config_SaveSetting "AF", "AFDelay", Str$(AFDelay)
    Config_SaveSetting "AF", "AFRampRate", Str$(AFRampRate)
    Config_SaveSetting "AF", "AFWait", Str$(AFWait)
    Config_SaveSetting "AF", "TSlope", Str$(TSlope)
    Config_SaveSetting "AF", "Toffset", Str$(Toffset)
    Config_SaveSetting "AF", "Thot", Str$(Thot)
    Config_SaveSetting "AF", "Tmax", Str$(Tmax)
    Config_SaveSetting "AFAxial", "AFAxialCoord", AfAxialCoord
    Config_SaveSetting "AFAxial", "AFAxialYPoint", Str$(AfAxialYpoint)
    Config_SaveSetting "AFAxial", "AfAxialXpoint", Str$(AfAxialXpoint)
    Config_SaveSetting "AFAxial", "AfAxialHighSlope", Str$(AfAxialHighSlope)
    Config_SaveSetting "AFAxial", "AfAxialLowSlope", Str$(AfAxialLowSlope)
    Config_SaveSetting "AFAxial", "AfAxialMax", Str$(AfAxialMax)
    Config_SaveSetting "AFAxial", "AfAxialMin", Str$(AfAxialMin)
    Config_SaveSetting "AFTrans", "AFTransCoord", AfTransCoord
    Config_SaveSetting "AFTrans", "AFTransYPoint", Str$(AfTransYpoint)
    Config_SaveSetting "AFTrans", "AfTransXpoint", Str$(AfTransXpoint)
    Config_SaveSetting "AFTrans", "AfTransHighSlope", Str$(AfTransHighSlope)
    Config_SaveSetting "AFTrans", "AfTransLowSlope", Str$(AfTransLowSlope)
    Config_SaveSetting "AFTrans", "AfTransMax", Str$(AfTransMax)
    Config_SaveSetting "AFTrans", "AfTransMin", Str$(AfTransMin)
    Config_SaveSetting "IRMPulse", "IRMHFAxis", IRMHFAxis
    Config_SaveSetting "IRMPulse", "IRMLFAxis", IRMLFAxis
    Config_SaveSetting "IRMPulse", "IRMLFBackfieldAxis", IRMLFBackfieldAxis
    Config_SaveSetting "IRMPulseHF", "PulseHFMax", Str$(PulseHFMax)
    Config_SaveSetting "IRMPulseHF", "PulseHFMin", Str$(PulseHFMin)
    Config_SaveSetting "IRMPulse", "PulseLFMax", Str$(PulseLFMax)
    Config_SaveSetting "IRMPulse", "PulseLFMin", Str$(PulseLFMin)
    Config_SaveSetting "IRMPulse", "PulseMCCVoltConverstion", Str$(PulseMCCVoltConversion)
    Config_SaveSetting "IRMPulse", "PulseVoltMax", Str$(PulseVoltMax)
    Config_SaveSetting "IRMPulse", "PulseReturnMCCVoltConverstion", Str$(PulseReturnMCCVoltConversion)
    Config_SaveSetting "ARM", "ARMMax", Str$(ARMMax)
    Config_SaveSetting "ARM", "ARMVoltGauss", Str$(ARMVoltGauss)
    Config_SaveSetting "ARM", "ARMVoltMax", Str$(ARMVoltMax)
    Config_SaveSetting "ARM", "ARMTimeMax", Str$(ARMTimeMax)
    ' (March 2008 L Carporzen) Put in Settings the IRM/ARM channels (MIT acquisition board does not work on IRMTrim = 3
    ' Analog channel output
    Config_SaveSetting "IRM-ARM", "ARMVoltageOut", Str$(ARMVoltageOut)
    Config_SaveSetting "IRM-ARM", "IRMVoltageOut", Str$(IRMVoltageOut)
    ' Analog input
    Config_SaveSetting "IRM-ARM", "IRMCapacitorVoltageIn", Str$(IRMCapacitorVoltageIn)
    Config_SaveSetting "AF", "AnalogT1", Str$(AnalogT1)
    Config_SaveSetting "AF", "AnalogT2", Str$(AnalogT2)
    ' DIO line assignments
    Config_SaveSetting "IRM-ARM", "ARMSet", Str$(ARMSet)
    Config_SaveSetting "IRM-ARM", "IRMFire", Str$(IRMFire)
    Config_SaveSetting "IRM-ARM", "IRMTrim", Str$(IRMTrim)
    Config_SaveSetting "IRM-ARM", "IRMReady", Str$(IRMReady)
    Config_SaveSetting "Vacuum", "MotorToggle", Str$(MotorToggle)
    Config_SaveSetting "Vacuum", "VacuumToggleA", Str$(VacuumToggleA)
    Config_SaveSetting "Vacuum", "VacuumToggleB", Str$(VacuumToggleB)
    Config_SaveSetting "Vacuum", "DoVacuumReset", Str$(DoVacuumReset)
    Config_SaveSetting "COMPorts", "COMPortSquids", Str$(COMPortSquids)
    Config_SaveSetting "COMPorts", "COMPortAf", Str$(COMPortAf)
    Config_SaveSetting "COMPorts", "COMPortUpDown", Str$(COMPortUpDown)
    Config_SaveSetting "COMPorts", "COMPortTurning", Str$(COMPortTurning)
    Config_SaveSetting "COMPorts", "COMPortChanger", Str$(COMPortChanger)
    Config_SaveSetting "COMPorts", "COMPortVacuum", Str$(COMPortVacuum)
    Config_SaveSetting "COMPorts", "COMPortSusceptibility", Str$(COMPortSusceptibility)
    Config_SaveSetting "COMPorts", "SusceptibilitySettings", SusceptibilitySettings
    Config_SaveSetting "MotorPrograms", "CmdHomeToTop", Str$(CmdHomeToTop)
    Config_SaveSetting "MotorPrograms", "CmdSamplePickup", Str$(CmdSamplePickup)
    Config_SaveSetting "MotorPrograms", "MotorIDTurning", Str$(MotorIDTurning)
    Config_SaveSetting "MotorPrograms", "MotorIDChanger", Str$(MotorIDChanger)
    Config_SaveSetting "MotorPrograms", "MotorIDUpDown", Str$(MotorIDUpDown)
    Config_SaveSetting "Program", "UsageFile", Prog_UsageFile
    Config_SaveSetting "Program", "DefaultPath", Prog_DefaultPath
    Config_SaveSetting "Program", "HelpURLRoot", Prog_HelpURLRoot
    Config_SaveSetting "Program", "NoCommMode", Str$(NOCOMM_MODE)
    Config_SaveSetting "Program", "DebugMode", Str$(DEBUG_MODE)
    Config_SaveSetting "Program", "DumpRawDataStats", Str$(DumpRawDataStats)
    Config_SaveSetting "Program", "LogMessages", Str$(LogMessages)
    Config_SaveSetting "Program", "LogoFile", Prog_LogoFile
    Config_SaveSetting "Program", "IcoFile", Prog_IcoFile ' (October 2007 L Carporzen)
    Config_SaveSetting "Program", "TextEditor", Prog_TextEditor
    Config_SaveSetting "Email", "MailSMTPHost", MailSMTPHost
    Config_SaveSetting "Email", "MailFrom", MailFrom
    Config_SaveSetting "Email", "MailFromName", MailFromName
    Config_SaveSetting "Email", "MailCCList", MailCCList
    Config_SaveSetting "Email", "MailStatusMonitor", MailStatusMonitor
    Config_SaveSetting "Program", "DefaultBackupDrive", Prog_DefaultBackup
    For i = 1 To 25
        If AFAxialX(i) > 0 Then Config_SaveSetting "AFAxial", "AFAxialX" & Format$(i, "0"), Str$(AFAxialX(i))
        If AFAxialY(i) > 0 Then Config_SaveSetting "AFAxial", "AFAxialY" & Format$(i, "0"), Str$(AFAxialY(i))
        If AFTransX(i) > 0 Then Config_SaveSetting "AFTrans", "AFTransX" & Format$(i, "0"), Str$(AFTransX(i))
        If AFTransY(i) > 0 Then Config_SaveSetting "AFTrans", "AFTransY" & Format$(i, "0"), Str$(AFTransY(i))
    Next i
    For i = 1 To 20
        If PulseLFX(i) > 0 Then Config_SaveSetting "IRMPulse", "PulseLFX" & Format$(i, "0"), Str$(PulseLFX(i))
        If PulseLFY(i) > 0 Then Config_SaveSetting "IRMPulse", "PulseLFY" & Format$(i, "0"), Str$(PulseLFY(i))
        If PulseHFX(i) > 0 Then Config_SaveSetting "IRMPulseHF", "PulseHFX" & Format$(i, "0"), Str$(PulseHFX(i))
        If PulseHFY(i) > 0 Then Config_SaveSetting "IRMPulseHF", "PulseHFY" & Format$(i, "0"), Str$(PulseHFY(i))
    Next i
    Config_SaveSetting "Modules", "EnableIRM", Str$(EnableIRM)
    Config_SaveSetting "Modules", "EnableIRMHi", Str$(EnableIRMHi)
    Config_SaveSetting "Modules", "EnableIRMBackfield", Str$(EnableIRMBackfield)
    Config_SaveSetting "Modules", "EnableIRMReturn", Str$(EnableIRMReturn)
    Config_SaveSetting "Modules", "EnableARM", Str$(EnableARM)
    Config_SaveSetting "Modules", "EnableAF", Str$(EnableAF)
    Config_SaveSetting "Modules", "EnableT1", Str$(EnableT1)
    Config_SaveSetting "Modules", "EnableT2", Str$(EnableT2)
    Config_SaveSetting "Modules", "EnableSusceptibility", Str$(EnableSusceptibility)
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

Public Function Config_GetSetting(sSection As String, sKey As String, sDefault As String) As String
    Config_GetSetting = Config_GetFromINI(sSection, sKey, sDefault, Prog_INIFile)
End Function

Public Sub Config_SaveSetting(sSection As String, sKey As String, sDefault As String)
    Dim dummy As Boolean
    dummy = Config_AddToINI(sSection, sKey, sDefault, Prog_INIFile)
End Sub

'// VB Web Code Example
'// www.vbweb.co.uk
'// Functions
Public Function Config_GetFromINI(sSection As String, sKey As String, sDefault As String, sIniFile As String)
    Dim sBuffer As String, lRet As Long
    ' Fill String with 255 spaces
    sBuffer = String$(255, 0)
    ' Call DLL
    lRet = GetPrivateProfileString(sSection, sKey, vbNullString, sBuffer, Len(sBuffer), sIniFile)
    If lRet = 0 Then
        ' DLL failed, save default
        If LenB(sDefault) <> 0 Then Config_AddToINI sSection, sKey, sDefault, sIniFile
        Config_GetFromINI = sDefault
    Else
        ' DLL successful
        ' return string
        Config_GetFromINI = Left(sBuffer, InStr(sBuffer, Chr(0)) - 1)
    End If
End Function

'// Returns True if successful. If section does not
'// exist it creates it.
Public Function Config_AddToINI(sSection As String, sKey As String, sValue As String, sIniFile As String) As Boolean
    Dim lRet As Long
    ' Call DLL
    lRet = WritePrivateProfileString(sSection, sKey, sValue, sIniFile)
    Config_AddToINI = (lRet)
End Function

' Sub Import_Boards()
'
' Created: March 30, 2010
'  Author: Isaac Hilburn
'
' Summary: Reads Paleomag.ini file and parses the [Boards] section of the file using
'          the Config_GetFromIni function.  Creates new board objects in SystemBoards collection
'          using the parameters specified for each board in the .ini file.

Public Sub Import_Boards()

    Dim TempBoard As Board
    Dim TempChannels As Channels
    Dim temprange As Range
    Dim NumBoards As Long
    Dim NumChannels As Long
    Dim DIO_isConfigured As Boolean
    Dim i As Long
    Dim j As Long
    
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
    NumBoards = val(Config_GetFromINI("Boards", "BoardsCount", "2", Prog_INIFile))
    
    'Now iterate through each of the Boards listed in the Boards section
    'and grab the necessary information
    For i = 1 To NumBoards
    
        'If there aren't enough boards in the System Boards collection, add a new one
        If i > SystemBoards.count Then
            
            'Need to add a new system board
            'Key for board = string number, Note: using long number will also work as a key
            SystemBoards.add , Trim(Str(i))
                
        End If
        
        'With the current, new board in the System Boards collection
        With SystemBoards(i)
        
            'Get the Board's Name from the .ini file
            .BoardName = Config_GetFromINI("Boards", _
                                           "BoardName" & Format$(i - 1, "0"), _
                                           "PCI-DAS6030", _
                                           Prog_INIFile)
                                           
            'Get the Board's IO mode - single (= 1) or differential (= 0) mode
            .BoardMode = val(Config_GetFromINI("Boards", _
                                           "BoardMode" & Format$(i - 1, "0"), _
                                           "1", _
                                           Prog_INIFile))
                                           
            'Get the Device number assigned to the board by the MS Windows operating system
            .BoardNum = val(Config_GetFromINI("Boards", _
                                           "BoardNum" & Format$(i - 1, "0"), _
                                           "0", _
                                           Prog_INIFile))
            
            'If no RangeType is specified for the board, then
            'the default value of -1 will be caught and the Max & Min
            'Range Values will be extracted manually
            '(The values are normally set by setting the Range Type)
            'The ADWIN board does not support a RangeType
            .BoardRange.RangeType = val(Config_GetFromINI("Boards", _
                                           "BoardRangeType" & Format$(i - 1, "0"), _
                                           "-1", _
                                           Prog_INIFile))
            
            'If RangeType has default (no rangetype) value = -1, then
            'Need to read in the Max & Min range values from the INI file
            '(This will be true for the ADWIN-light-16 board)
            If .BoardRange.RangeType = -1 Then
            
                .BoardRange.MaxValue = val(Config_GetFromINI("Boards", _
                                           "BoardRangeMax" & Format$(i - 1, "0"), _
                                           "10", _
                                           Prog_INIFile))
                                           
                .BoardRange.MinValue = val(Config_GetFromINI("Boards", _
                                           "BoardMin" & Format$(i - 1, "0"), _
                                           "-10", _
                                           Prog_INIFile))
                                           
            End If
            
            'This field specifies what type of Board this is
            '1 = MCC_UL, Measurement Computing Board (most likely a PCI-DAS6030)
            '2 = ADWIN, ADWIN-light-16 board
            '3 = OTHER, Some other type off board, currently not supported explicitly by the code
            .CommProtocol = val(Config_GetFromINI("Boards", _
                                           "CommProtocol" & Format$(i - 1, "0"), _
                                           Trim(Str(ADWIN)), _
                                           Prog_INIFile))
                                           
            'This field specifies a string letting other portions of the code
            'know what this board can and cannot be used to do
            .BoardFunction = Config_GetFromINI("Boards", _
                                           "BoardFunction" & Format$(i - 1, "0"), _
                                           "ARM Output," & _
                                           "ARM Relays," & _
                                           "IRM Output," & _
                                           "IRM Monitor," & _
                                           "IRM Relays," & _
                                           "AF Control," & _
                                           "AF Monitor," & _
                                           "AF Relays," & _
                                           "Vacuum", _
                                           Prog_INIFile)
            
            'Maximum Analog Input Rate supported for this board in Hz
            .MaxAInRate = val(Config_GetFromINI("Boards", _
                                           "MaxAInRate" & Format$(i - 1, "0"), _
                                           "100000", _
                                           Prog_INIFile))
                                           
            
            'Maximum Analog Output Rate supported for this board in Hz
            .MaxAInRate = val(Config_GetFromINI("Boards", _
                                           "MaxAOutRate" & Format$(i - 1, "0"), _
                                           "100000", _
                                           Prog_INIFile))
                                           
            'Now need to load Board channel configuration from the .ini file
                
            'Get Number of Analog Input Channels
            NumChannels = val(Config_GetFromINI("Boards", _
                                           "AInChannelsCount" & Format$(i - 1, "0"), _
                                           "0", _
                                           Prog_INIFile))
                                           
            'Screen for Import error (Default, NumChannels = 0)
            If NumChannels <> 0 Then
            
                'Add the necessary number of channels
                'to the Analog Input channels collection on this Board
                For j = 1 To NumChannels
            
                    If j > .AInChannels.count Then
                    
                        .AInChannels.add , Trim(Str(j))
                    
                    End If
                    
                    With .AInChannels(j)
                    
                        
                        'Note: that the format of the key string for the channel
                        '      contains both an element indicating the board that
                        '      the channel is on, and also indicating the channel
                        '      number itself.
                        
                        'Snatch the Channel Name from the INI file
                        .ChanName = Config_GetFromINI("Boards", _
                                                      "AI-" & Format(i - 1, "0") & "-" & _
                                                        "CH" & Format(j - 1, "00"), _
                                                      "AI" & Format(j - 1, "00"), _
                                                      Prog_INIFile)
                                                                                  
                        'Snatch the Channel Number from the INI File
                        .ChanNum = val(Config_GetFromINI("Boards", _
                                                      "AI-" & Format(i - 1, "0") & "-" & _
                                                        "CH" & Format(j - 1, "00"), _
                                                      Trim(Str(j - 1)), _
                                                      Prog_INIFile))
                                                      
                    End With
                    
                Next j
                                                      
            End If
                        
            'Get Number of Analog Output Channels
            NumChannels = val(Config_GetFromINI("Boards", _
                                           "AOutChannelsCount" & Format$(i - 1, "0"), _
                                           "0", _
                                           Prog_INIFile))
                                           
            'Screen for Import error (Default, NumChannels = 0)
            If NumChannels <> 0 Then
            
                'Add the necessary number of channels
                'to the Analog Output channels collection on this Board
                For j = 1 To NumChannels
            
                    If j > .AOutChannels.count Then
                    
                        .AOutChannels.add , Trim(Str(j))
                    
                    With .AOutChannels(j)
                                                
                        'Note: that the format of the key string for the channel
                        '      contains both an element indicating the board that
                        '      the channel is on, and also indicating the channel
                        '      number itself.
                        
                        'Snatch the Channel Name from the INI file
                        .ChanName = Config_GetFromINI("Boards", _
                                                      "AO-" & Format(i - 1, "0") & "-" & _
                                                        "CH" & Format(j - 1, "00"), _
                                                      "AO" & Format(j - 1, "00"), _
                                                      Prog_INIFile)
                                                                                  
                        'Snatch the Channel Number from the INI File
                        .ChanNum = val(Config_GetFromINI("Boards", _
                                                      "AO-" & Format(i - 1, "0") & "-" & _
                                                        "CH" & Format(j - 1, "00"), _
                                                      Trim(Str(j - 1)), _
                                                      Prog_INIFile))
                                                      
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
            DIO_isConfigured = (Config_GetFromINI("Boards", _
                                                      "DIOConfigured" & Format(i - 1, "0"), _
                                                      "False", _
                                                      Prog_INIFile) = "True")

            'Get Number of Digital Input Channels
            NumChannels = val(Config_GetFromINI("Boards", _
                                           "DInChannelsCount" & Format$(i - 1, "0"), _
                                           "0", _
                                           Prog_INIFile))
                                           
            'Screen for Import error (Default, NumChannels = 0)
            If NumChannels <> 0 Then
            
                'Add the necessary number of channels
                'to the Digital Input channels collection on this Board
                For j = 1 To NumChannels
            
                    If j > .DInChannels.count Then
                    
                        .DInChannels.add , Trim(Str(j))
                    
                    With .DInChannels(j)
                                                
                        'Note: that the format of the key string for the channel
                        '      contains both an element indicating the board that
                        '      the channel is on, and also indicating the channel
                        '      number itself.
                        
                        'Snatch the Channel Name from the INI file
                        .ChanName = Config_GetFromINI("Boards", _
                                                      "DI-" & Format(i - 1, "0") & "-" & _
                                                        "CH" & Format(j - 1, "00"), _
                                                      "DI" & Format(j - 1, "00"), _
                                                      Prog_INIFile)
                                                                                  
                        'Snatch the Channel Number from the INI File
                        .ChanNum = val(Config_GetFromINI("Boards", _
                                                      "DI-" & Format(i - 1, "0") & "-" & _
                                                        "CH" & Format(j - 1, "00"), _
                                                      Trim(Str(j - 1)), _
                                                      Prog_INIFile))
                                                      
                        'Now set whether or not the channel is configured for input
                        .DIOConfigured = DIO_isConfigured
                        
                        'If the DIO is configured, then need to set the mode to Input
                        If .DIOConfigured Then
                        
                            .DIOMode = IOINPUT
                            
                        End If
                                                      
                    End With
                    
                Next j
                                                      
            End If
            
            'Get Number of Digital Output Channels
            NumChannels = val(Config_GetFromINI("Boards", _
                                           "DOutChannelsCount" & Format$(i - 1, "0"), _
                                           "0", _
                                           Prog_INIFile))
                                           
            'Screen for Import error (Default, NumChannels = 0)
            If NumChannels <> 0 Then
            
                'Add the necessary number of channels
                'to the Digital Output channels collection on this Board
                For j = 1 To NumChannels
            
                    If j > .DOutChannels.count Then
                    
                        .DOutChannels.add , Trim(Str(j))
                    
                    With .DOutChannels(j)
                                                
                        'Note: that the format of the key string for the channel
                        '      contains both an element indicating the board that
                        '      the channel is on, and also indicating the channel
                        '      number itself.
                        
                        'Snatch the Channel Name from the INI file
                        .ChanName = Config_GetFromINI("Boards", _
                                                      "DO-" & Format(i - 1, "0") & "-" & _
                                                        "CH" & Format(j - 1, "00"), _
                                                      "DO" & Format(j - 1, "00"), _
                                                      Prog_INIFile)
                                                                                  
                        'Snatch the Channel Number from the INI File
                        .ChanNum = val(Config_GetFromINI("Boards", _
                                                      "DO-" & Format(i - 1, "0") & "-" & _
                                                        "CH" & Format(j - 1, "00"), _
                                                      Trim(Str(j - 1)), _
                                                      Prog_INIFile))
                                                      
                        'Now set whether or not the channel is configured for input
                        .DIOConfigured = DIO_isConfigured
                        
                        'If the DIO is configured, then need to set the mode to Input
                        If .DIOConfigured Then
                        
                            .DIOMode = IOOUTPUT
                            
                        End If
                                                      
                    End With
                    
                Next j
                                                      
            End If
            
        End With
        
    Next i

End Sub

' Sub Import_Channels()
'
' Created: March 30, 2010
'  Author: Isaac Hilburn
'
' Summary: Reads Paleomag.ini file and parses the [Channels] section of the file using
'          the Config_GetFromIni function.  Loads the correct channel from the correct board
'          into each needed global channel object variable (i.e. ARMVoltageOut, AnalogT1, etc.)

Public Sub Import_Channels()
    
    Dim ChanStr As String
    Dim BoardNum As Long
    
    'Check to see if Import_Boards has been called. If not, call it
    If Not ImportBoardsDone Then
    
        Import_Boards
        
    End If
    
    'For Each channel that we need to parse out the correct channel / board combination from
    'the now populated board
    
    ' (March 2008 - L Carporzen) Put in Settings the IRM/ARM channels (MIT acquisition board does not work on IRMTrim = 3
' Analog channel output
'(March 2010 - I Hilburn) Changed Integer channel/port numbs to Channel objects
 ARMVoltageOut As Channel
 IRMVoltageOut  As Channel

' Analog input
'(March 2010 - I Hilburn) Changed Integer chan/port number to Channel object
 IRMCapacitorVoltageIn  As Channel

'Analog MCC Input Channels #'s for Temperature sensors on AF coils
'(March 2010 - L Carporzen)
'(March 2010 - I Hilburn) Changed Integer channel/port numbs to Channel objects
 AnalogT1 As Channel
 AnalogT2 As Channel

' DIO line assignments
'(March 2010 - I Hilburn) Changed Integer channel/port numbs to Channel objects
 ARMSet  As Channel
 IRMFire  As Channel
 IRMTrim  As Channel
 IRMReady  As Channel
 MotorToggle As Channel
 VacuumToggleA As Channel
 VacuumToggleB As Channel
    
End Sub
