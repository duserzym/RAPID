Attribute VB_Name = "modProg"
' PALEOMAGNETIC MAGNETOMETER CONTROL SYSTEM
' by Robert E. Kopp / rkopp@caltech.edu
' Copyright (C) 2010 by the California Institute of Technology ' (June 2007 BP Weiss changed 2006 to 2007)
' Licensed under the GNU General Public License
' -------------------------------------------------
' This module stores settings and procedures that handle program-specific
' details.  These, for the most part are general data and functions that
' cannot be categorized into the other modules
' ------ Global Variables ------
' (February 2010 L Carporzen) Webcam parameters
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function capCreateCaptureWindow Lib "avicap32.dll" Alias "capCreateCaptureWindowA" (ByVal lpszWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hwndParent As Long, ByVal nID As Long) As Long
Public mCapHwnd As Long
Public Const Connect As Long = 1034
Public Const Disconnect As Long = 1035
Public Const GET_FRAME As Long = 1084
Public Const COPY As Long = 1054
Public lastHelpFile As String
Public Prog_INIFile As String
Public Prog_1stTime As Boolean  'Record if this is the first time the Paleomag code has been run
Public INIConversionDone As Boolean
Public DEBUG_MODE           As Boolean
Public NOCOMM_MODE          As Boolean
Public LoginName           As String      ' Name of current user
Public LoginEmail          As String
Public currentPosInitialized As Boolean
Public CoilsLocked        As Boolean    '(July 30, 2010 - I Hilburn) Added to allow program wide locking &
                                        'unlocking of the Axial / Transverse coil selection
Public SampleHolder       As Sample
Public SusceptibilityStandard As Sample
Public MainChanger         As frmChanger
Public fChangerSampOrder   As frmChangerSampOrder
Public MeasurementsSinceHolder As Long
Public SampleIndexRegistry As SampleIndexRegistrations
Public SampQueue As SampleCommands
Global FLAG_MagnetInit     As Boolean     ' Whether we've finished initializing the magnetometer
Global FLAG_MagnetUse      As Boolean     ' Whether we've finished using the magnetometer
Global Const DEMAGLEN As Integer = 6     ' Maximum length of "demag" string
Global Const MAX_DOUBLE As Double = 1.79769313486231E+307

' ------ Types            ------
Type SamFileDat
    ' This data type stores all the data fields
    ' that describe the current sample begin measured
    locality As String
    siteLat  As Double
    sitelon  As Double
    magDec   As Double
    NumSamples As Integer
End Type
Type AF_Status
    Status As Integer
    Delay As Integer
    Coil As String
    Amplitude As Integer
End Type
Type AF
    axis As String
    Max As Double
    ypoint As Double
    xpoint As Double
    loslope As Double
    hislope As Double
End Type
Public Type QueueEntry
    CmdId   As String
    hole    As Double
    fileid  As Integer
End Type
Public Type Measure_Specimen
    ' Read in from the first records of a specimen file:
    '   Core, bedding, and fold orientation data
    Name            As String  ' The unique name of this sample
    holeNum         As Long    ' The hole that the sample is in
    sampleid        As Integer
    fileid          As Integer
    CorePlateStrike As Double
    CorePlateDip    As Double
    BeddingStrike   As Double
    BeddingDip      As Double
    Vol             As Double  ' Sample volume
    FoldRotation    As Boolean
    FoldAxis        As Double
    FoldPlunge      As Double
End Type
Type SampleIdentifier
    filename As String
    Samplename As String
End Type

Sub AppendLog(aline As String)
    ' This sub opens the usage log and appends to it the
    ' string "aline".  It's simple.
    Dim f As Integer
    f = FreeFile
    If Not FileExists(Prog_UsageFile) Then
    
        On Error GoTo BadPaleoUsageFilePath:
    
            Open Prog_UsageFile For Output As #f
            
        On Error GoTo 0
        
    Else
        Open Prog_UsageFile For Append As #f
    End If
    Print #f, aline
    Close #f
    
    Exit Sub
    
BadPaleoUsageFilePath:
    
    'The default path for the paleo-usage file doesn't work on this computer
    'Need to start a common dialog to guide the user to select a new path
    With frmSplash.dialogGetINIFile
    
        'Setup the file dialog
        .DefaultExt = ".DAT"
        .filter = "Data File (.DAT)|*.DAT"
        .flags = cdlOFNCreatePrompt
        .DialogTitle = "Choose / Create Paleo-usage Log File:"
        .InitDir = App.path
                             
        'Open up the file dialog
        .ShowOpen
        
        Prog_UsageFile = .filename
    
    End With
    
    Open Prog_UsageFile For Append As #f
    Print #f, aline
    Close #f
    
    modConfig.Config_SaveSetting "Program", "UsageFile", Prog_UsageFile
    
End Sub

'Sub ArrayCopy
'
' Created: Feb. 24, 2011
'  Author: I. Hilburn
'
' Summary:  Takes in references to two 1D or 2D arrays of variant data type
'           and the name of the function to use to change
'           the data type of array 1 into the desired type for
'           array2.
'           Sub routine then copies elements (0 to N - 1) x (0 to M - 1)
'           while using CallByName to cast the data type of each element based upon the
'           function name given.
'
'           This subroutine will overwrite the values of N and M if they exceed
'           the bounds of Array1 or Array2 and automatically tests if the arrays are 1D or 2D
'           Additionally, if one array is 2D and the other is 1D, then only the first column
'           of the 2D array will be read from / written to
'
'  Inputs:
'
'   Array1  -   Referenced 1D or 2D array object with any data type that can fit within a Variant
'               Data will be read from this array element by element
'
'   Array2  -   Referenced 1D or 2D array object with any data type that can fit within a Variant
'               Data will be written to this array element by element
'
'   N       -   Number of rows in Array1 to be copied to Array2
'
'   M       -   Number of cols in Array1 to be copied to Array2, defaults to 2
'
Public Sub ArrayCopy(ByRef Array1 As Variant, _
                     ByRef Array2 As Variant, _
                     ByVal N As Long, _
                     Optional ByVal M As Long = 2)

    Dim i, j As Long
    Dim temp As Long
    
    'Check to see if Array1 and Array2 are arrays
    If Not IsArray(Array1) Or _
       Not IsArray(Array2) Then Exit Sub
    
    'Check to make sure that N is within the bounds of Array1
    temp = UBound(Array1, 1) + 1
    If N > temp Then N = temp
    
    'Now check to make sure the N & M are within the bounds of Array2
    temp = UBound(Array2, 1) + 1
    If N > temp Then N = temp
    
    'Check M vs the 2nd dimension of array 1 and array 2
    'Turn on error checking, if this is not a 2D array, this
    'Ubound call will cause an error
    On Error GoTo Array1Not2D:
        
        temp = UBound(Array1, 2) + 1
        If M > temp Then M = temp
    
    On Error GoTo 0
    
    'Turn on error checking, if this is not a 2D array, this
    'Ubound call will cause an error
    On Error GoTo Array2Not2D:
        
        temp = UBound(Array2, 2) + 1
        If M > temp Then M = temp
    
    On Error GoTo 0
    
    
    For i = 0 To N - 1
    
        For j = 0 To M - 1
        
            Array2(i, j) = Array1(i, j)
                                      
        Next j
        
    Next i
    
    'Task done, exit the function
    Exit Sub
    
Array1Not2D:

    'Treat Array1 as 1D array
    
    'Is Array2 a 1D array as well?
    On Error GoTo BothNot2D:
    
        temp = UBound(Array2, 2)
        
    On Error GoTo 0
    
    'Array2 is 2D
    'Copy Array1 into the first column of Array2
    For i = 0 To N - 1
    
        Array2(i, 0) = Array1(i)
                                  
    Next i
    
    'The copying is done, leave the subroutine
    Exit Sub
    
Array2Not2D:

    'Treat Array2 as 1D array
    
    'Is array1 a 1D array as well?
    On Error GoTo BothNot2D:
    
        temp = UBound(Array1, 2)
        
    On Error GoTo 0
        
    'Array1 is 2D
    'Copy the first column of Array1 into Array2
    For i = 0 To N - 1
    
        Array2(i) = Array1(i, 0)
                                  
    Next i

    'Task is done, exit the subroutine
    Exit Sub
    
BothNot2D:

    'Both arrays are 1D
    'Array1 is 2D
    'Copy the first column of Array1 into Array2
    For i = 0 To N - 1
    
        Array2(i) = Array1(i)
                                  
    Next i

End Sub

    
Private Function CheckINIFileFormat() As Boolean

    'Default CheckINIFileFormat = True
    CheckINIFileFormat = True

    'Need to check that every [Section] that's in the new INI file format exists in the current INI file
    With IniFile
    
        If Not .SectionExists("Program") Then CheckINIFileFormat = False
        If Not .SectionExists("SampleChanger") Then CheckINIFileFormat = False
        If Not .SectionExists("SteppingMotor") Then CheckINIFileFormat = False
        If Not .SectionExists("MotorPrograms") Then CheckINIFileFormat = False
        If Not .SectionExists("Boards") Then CheckINIFileFormat = False
        If Not .SectionExists("WaveForms") Then CheckINIFileFormat = False
        If Not .SectionExists("Channels") Then CheckINIFileFormat = False
        If Not .SectionExists("MagnetometerCalibration") Then CheckINIFileFormat = False
        If Not .SectionExists("AF") Then CheckINIFileFormat = False
        If Not .SectionExists("AFAxial") Then CheckINIFileFormat = False
        If Not .SectionExists("AFTrans") Then CheckINIFileFormat = False
        If Not .SectionExists("IRMPulse") Then CheckINIFileFormat = False
        If Not .SectionExists("IRMAxial") Then CheckINIFileFormat = False
        If Not .SectionExists("IRMTrans") Then CheckINIFileFormat = False
        If Not .SectionExists("ARM") Then CheckINIFileFormat = False
        If Not .SectionExists("Vacuum") Then CheckINIFileFormat = False
        If Not .SectionExists("COMPorts") Then CheckINIFileFormat = False
        If Not .SectionExists("Email") Then CheckINIFileFormat = False
        If Not .SectionExists("SusceptibilityCalibration") Then CheckINIFileFormat = False
        If Not .SectionExists("Magnetometry") Then CheckINIFileFormat = False
        If Not .SectionExists("RockmagRoutineDefaults") Then CheckINIFileFormat = False
    
    End With
    
End Function

Sub DelayTime(PauseTime As Double)
    ' This procedure pauses the program for some time allowing other
    ' to continue.  PauseTime is in seconds.
    ' CHANGELOG: 8-30-99  Added check for Timer reset at midnight
    Dim Start, Finish, TotalTime, CurTimer
    CurTimer = Timer
    Start = CurTimer   ' Set start time.
    Do While CurTimer < Start + PauseTime
        DoEvents    ' Yield to other processes.
        CurTimer = Timer
        If CurTimer < Start Then Start = Start - 86400
    Loop
    Finish = Timer  ' Set end time.
    TotalTime = Finish - Start  ' Calculate total time.
End Sub

Public Function FileExists(p As String) As Boolean
    ' This function determines whether a file exists.
    ' It returns the corresponding boolean value
    FileExists = False
    On Error GoTo fin:
    If LenB(dir$(p, vbNormal + vbDirectory)) <> 0 Then
        FileExists = True
    End If
    On Error GoTo 0
fin:
End Function

Function FormatNumber(ByVal val As Double) As String
    ' Now select the proper format for printing out this range
    ' information based on the TESTIT variable
    Dim frmt As String
    Dim testit As Double
    testit = val
    If (testit >= 1000000) Or (testit <= -100000) Then
        frmt = "00000000"
    ElseIf (testit >= 100000) Or (testit <= -10000) Then
        frmt = "000000.0"
    ElseIf (testit >= 10000) Or (testit <= -1000) Then
        frmt = "00000.00"
    ElseIf (testit >= 1000) Or (testit <= -100) Then
        frmt = "0000.000"
    ElseIf (testit >= 100) Or (testit <= -10) Then
        frmt = "000.0000"
    ElseIf (testit >= 10) Or (testit <= -1) Then
        frmt = "00.00000"
    Else
        frmt = "0.000000"
    End If
    FormatNumber = Format$(val, frmt)
End Function

Sub LoadResStrings(frm As Form)
    On Error Resume Next
    Dim ctl As Control
    Dim obj As Object
    Dim fnt As Object
    Dim sCtlType As String
    Dim nVal As Integer
    'set the form's caption
    frm.Caption = LoadResString(CInt(frm.Tag))
    'set the font
    Set fnt = frm.Font
    fnt.Name = LoadResString(20)
    fnt.size = CInt(LoadResString(21))
    'set the controls' captions using the caption
    'property for menu items and the Tag property
    'for all other controls
    For Each ctl In frm.Controls
        Set ctl.Font = fnt
        sCtlType = TypeName(ctl)
        If sCtlType = "Label" Then
            ctl.Caption = LoadResString(CInt(ctl.Tag))
        ElseIf sCtlType = "Menu" Then
            ctl.Caption = LoadResString(CInt(ctl.Caption))
        ElseIf sCtlType = "TabStrip" Then
            For Each obj In ctl.Tabs
                obj.Caption = LoadResString(CInt(obj.Tag))
                obj.ToolTipText = LoadResString(CInt(obj.ToolTipText))
            Next
        ElseIf sCtlType = "Toolbar" Then
            For Each obj In ctl.Buttons
                obj.ToolTipText = LoadResString(CInt(obj.ToolTipText))
            Next
        ElseIf sCtlType = "ListView" Then
            For Each obj In ctl.ColumnHeaders
                obj.text = LoadResString(CInt(obj.Tag))
            Next
        Else
            nVal = 0
            nVal = val(ctl.Tag)
            If nVal > 0 Then ctl.Caption = LoadResString(nVal)
            nVal = 0
            nVal = val(ctl.ToolTipText)
            If nVal > 0 Then ctl.ToolTipText = LoadResString(nVal)
        End If
    Next
End Sub

Sub Main()

    Dim INISearchPath As String
    Dim TempL As Long
    Dim fso As FileSystemObject

    Set fso = New FileSystemObject

    'Set time interval for timegettime() function to 1 ms
    timeBeginPeriod 1

    'Check to see if this is the first time the v2.4.0 code has been run
    Prog_1stTime = (Trim(GetSetting(App.EXEName, "Settings", "FirstTime", "True")) = "True")

    'If this is the first time, search for the INI file attached to the pre-version 2.4
    'instance of the Paleomag program
    If Prog_1stTime Then
    
        Prog_INIFile = GetSetting("Paleomag.exe", "Settings", "INIFile", "-1")
        
    Else
    
        'Not the first time, retrieve INI file setting for v2.4.0 of the Paleomag code
        Prog_INIFile = GetSetting(App.EXEName, "Settings", "INIFile", "-1")
        
    End If
        
    'Check to see if the program successully got the Ini file
    If Prog_INIFile = "-1" Or _
       fso.FileExists(Prog_INIFile) = False _
    Then
    
        'Need to Prompt the user for the ini file path
    
OpenINIFile:
    
        'Setup INI File dialog browser
        With frmSplash.dialogGetINIFile
        
            .DefaultExt = ".ini"
            .filter = "Settings File (*.ini)|*.ini"
            .flags = cdlOFNFileMustExist
            .DialogTitle = "Open the Paleomag.ini settings file."
            
            INISearchPath = App.path
            TempL = InStrRev(INISearchPath, "\")
            
            INISearchPath = Mid(INISearchPath, 1, TempL - 1)
            
            .InitDir = INISearchPath
            
            .ShowOpen
            
            Prog_INIFile = .filename
            
        End With
                
    End If
        
    SaveSetting App.EXEName, "Settings", "INIFile", Prog_INIFile
    
    'Create the Ini File object
    Set modConfig.IniFile = New CIniFile
    IniFile.filename = Prog_INIFile
    
    '--------------------------------------------------------------------------------------------------------------------'
    '--------------------------------------------------------------------------------------------------------------------'
    '
    '   Modification: June, 2010
    '         Author: Isaac Hilburn
    '
    '        Summary: Need to add a bit of code in to see if we need to convert
    '                 an old-style (Pre-Version 2.4.0) formated .INI file into
    '                 a new-style (Post-Version 2.4.0) formated .INI file.
    '--------------------------------------------------------------------------------------------------------------------'
    
        'Local = System for the DAQ Boards Settings form
        '(the user hasn't had a chance to change the DAQ Board setup, yet)
        'This variable needs to be set before attempting a .INI file format conversion
        'as frmSystemBoardsettings will be called in the conversion process
        modConfig.LocalAndSystemDifferent = False
               
        'Check to see if the user has a pre-version 2.4.0 formated .INI file
        If CheckINIFileFormat = False Then
        
            'Prompt user to see if they want to select a new INI file
            Dim userResp2 As Long
            userResp2 = MsgBox("Your INI file appears to be in the pre-v2.4.0 format." & vbNewLine & _
                               "File: " & Prog_INIFile & vbNewLine & vbNewLine & _
                               "Would you like to open a different INI file?", vbYesNo, _
                               "Open New INI File?")
        
            If userResp2 = vbYes Then GoTo OpenINIFile:
                
        
            'Set INIConversionDone flag = False
            INIConversionDone = False
        
            'Load and show the INI Converter form
            Load frmINIConverter
            frmINIConverter.Left = (Screen.Width - frmINIConverter.Width) / 2
            frmINIConverter.Top = (Screen.Height - frmINIConverter.Height) / 2
                        
            
            frmINIConverter.ZOrder 0
            
            frmINIConverter.Show
            
            frmINIConverter.ZOrder 0
            
            Do While INIConversionDone = False
            
                PauseTill timeGetTime + 50
                
            Loop
            
            'Save the new INI file location to the application settings
            SaveSetting App.EXEName, "Settings", "INIFile", Prog_INIFile
            
        End If
                
    '--------------------------------------------------------------------------------------------------------------------'
    '--------------------------------------------------------------------------------------------------------------------'
                
    Config_ReadINISettings
    
    'We've now gone through and loaded the settings from the upgraded v2.4.0 compatible
    'INI file.  If no errors have happened yet, we can safely set the Application settings
    'to show that this is no longer the first time running v2.4.0 of the Paleomag code
    SaveSetting App.EXEName, "Settings", "FirstTime", False
        
    '(August 5, 2010 - I Hilburn) Set active coil system to Axial coil to save on problems
    'in the IRM initialization.  If this is changed to NoCoilSystem, the code will break
    'the first time the IRM is discharged at the start of the code warm-up.
    ActiveCoilSystem = AxialCoilSystem
    
    'Make sure the coil selection is unlocked
    CoilsLocked = False
            
    'Set the isSystemCoilChange flag = False
    modAF_DAQ.isSystemCoilChange = False
        
    FLAG_MagnetInit = False     ' Magnetometer uninitialized
    FLAG_MagnetUse = False      ' Magnetometer not in use
    SetCodeLevel CodeGrey
    Load frmProgram
    frmProgram.ZOrder
    frmProgram.Show
    ' Show splash screen while everything is loading
    Load frmSplash
    On Error GoTo oops
    
    'Load the Program Logo File
    If FileExists(Prog_LogoFile) And LenB(Prog_LogoFile) > 0 Then _
        frmSplash.imgLogo.Picture = LoadPicture(Prog_LogoFile)
       
           

'(July 2010 - I Hilburn) Commented out and replaced with the code line above
'    If FileExists(Prog_IcoFile) And LenB(Prog_IcoFile) > 0 Then frmProgram.Icon = LoadPicture(Prog_IcoFile) ' (October 2007 L Carporzen)
oops:
    On Error GoTo 0
    Load frmProgram
    
    '(March 25, 2011 - I Hilburn)
    'Added in error handling to get rid of the "Invalid Picture" error
    On Error GoTo BadIcoFile:
    
         'If the ICO file exists, then load it to the frmProgram Image List
         If FileExists(Prog_IcoFile) Then
         
             frmProgram.Prog_ImageList.ListImages.Add , "Icon", LoadPicture(Prog_IcoFile)
        
         End If
         
         frmProgram.Icon = frmProgram.Prog_ImageList.ListImages("Icon").Picture
         
BadIcoFile:

    On Error GoTo 0
    
    frmSplash.ZOrder
    frmSplash.Show
    frmSplash.refresh
    ' Load all forms into memory
    frmSplash.SplashStatus "Loading tip..."
    frmSplash.progress 1 / 9
    Load frmTip
    Load frmDebug
    frmSplash.SplashStatus "Loading login..."
    frmSplash.progress 2 / 9
    Load frmLogin
    frmLogin.Hide
    frmSplash.SplashStatus "Loading main program..."
    frmSplash.progress 3 / 9
    Load frmMagnetometerControl
    frmMagnetometerControl.Hide
    frmSplash.SplashStatus "Loading DC motors controller..."
    frmSplash.progress 4 / 9
    Load frmDCMotors
    frmSplash.SplashStatus "Loading sample changer management..."
    frmSplash.progress 5 / 9
    Set SampQueue = New SampleCommands
    Set SampleIndexRegistry = New SampleIndexRegistrations

''=================================================================================================
'    '(March 10, 2011 - I Hilburn)
'    'This statement is now handled within SampleIndexRegistrations.MakeSampleHolder
'    'which in turn is called when Initializing an instance of the SampleIndexRegistrations
'    'class (as has been done above)
''-------------------------------------------------------------------------------------------------
'    Set SampleHolder = SampleIndexRegistry("!Holder").sampleSet("Holder")
''=================================================================================================

'    Set SusceptibilityStandard = SampleIndexRegistry("!Holder").sampleSet("SusStd")
    Set MainChanger = New frmChanger
    MainChanger.IsMasterList = True
    Load MainChanger
    frmSplash.SplashStatus "Loading SQUID controller..."
    frmSplash.progress 6 / 9
    Load frmSQUID
    If EnableAxialIRM Or EnableARM Then frmSplash.SplashStatus "Loading IRM and ARM controller..."
    Load frmIRMARM
    frmSplash.SplashStatus "Loading vacuum controller..."
    frmSplash.progress 7 / 9
    Load frmVacuum
    frmSplash.SplashStatus "Loading sendmail..."
    frmSplash.progress 8 / 9
    Load frmSendMail
    frmSplash.SplashStatus "Ready."
    frmSplash.progress 1
    frmTip.ZOrder
    frmLogin.ZOrder
    frmTip.Show
    frmLogin.Show
    Unload frmSplash
    SetCodeLevel CodeBlue
End Sub

Public Sub Prompt_NOCOMM()

    Dim UserResponse As Long
    
    'Ask the user if they would like to switch on NOCOMM mode
    UserResponse = MsgBox("Would you like to switch on NOCOMM mode?" & _
                          vbNewLine & vbNewLine & "This will prevent the Paleomag program from trying to connect to " & _
                          "any peripheral devices and can be turned off by clicking the 'Turn Off NOCOMM mode' button in the " & _
                          "Main program window.", _
                          vbYesNo, _
                          "Whoops!")
                          
    'If the user answers yes, then change the caption in the NOCOMM mode toggle button and turn on NOCOMM Mode
    If UserResponse = vbYes Then
    
        frmProgram.cmdToggleNoComm.Caption = "Turn Off NOCOMM Mode"
        
        NOCOMM_MODE = True
        
    End If
    
End Sub

Public Sub SetFormIcon(ByRef FormObj As Form, _
                       Optional ByVal IcoFilePath As String = vbNullString)
                       
    'Check for default IcoFilePath value
    If IcoFilePath = vbNullString Then IcoFilePath = Prog_IcoFile
    
    'Turn on Error handling
    On Error GoTo BadIcoFile:
    
        Set FormObj.Icon = frmProgram.Prog_ImageList.ListImages("Icon").Picture
        
    On Error GoTo 0
    
BadIcoFile:
    
End Sub

