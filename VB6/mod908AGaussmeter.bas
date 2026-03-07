Attribute VB_Name = "mod908AGaussmeter"
Option Explicit

Public handle As Long
Public connected As Boolean
Public PollTime As Long
Public data As gm_store

Public DataArray() As gm_store

Public sampleindex As Integer
Public samplerate As Integer
Dim initstatus As Boolean

Public Comport As Integer
Public ConnectStatus As Integer

Public unitsrange(4, 4) As String
Public unitsrangefmt(4, 4) As String

Public unitsrangesclar(4, 4) As Double

Public baseunits(4)
Public modestr(5)

Public Mode As Integer

Dim dots As Integer

'*******************************************************************************************
'*                                                                                         *
'*                              Gaussmeter structures                                      *
'*                                                                                         *
'*******************************************************************************************

Public Type gm_time
    sec As Byte
    min As Byte
    hour As Byte
    day As Byte
    month As Byte
    year As Byte
End Type

Public Type gm_store
    time As gm_time
    range As Byte
    Mode As Byte
    Units As Byte
    value As Single
End Type

'*******************************************************************************************
'*                                                                                         *
'*                              Gaussmeter DLL prototypes                                  *
'*                                                                                         *
'*******************************************************************************************

' start stop functions
Public Declare Function gm0_newgm Lib "gm0.dll" (ByVal port As Long, ByVal Mode As Long) As Long
Public Declare Function gm0_startconnect Lib "gm0.dll" (ByVal handle As Long) As Long
Public Declare Function gm0_killgm Lib "gm0.dll" (ByVal handle As Long) As Long
Public Declare Function gm0_getconnect Lib "gm0.dll" (ByVal handle As Long) As Boolean

'set functions
Public Declare Function gm0_setrange Lib "gm0.dll" (ByVal handle As Long, ByVal range As Byte) As Long
Public Declare Function gm0_setunits Lib "gm0.dll" (ByVal handle As Long, ByVal Units As Byte) As Long
Public Declare Function gm0_setmode Lib "gm0.dll" (ByVal handle As Long, ByVal Mode As Byte) As Long

Public Declare Function gm0_isnewdata Lib "gm0.dll" (ByVal handle As Long) As Long

'get functions
Public Declare Function gm0_getrange Lib "gm0.dll" (ByVal handle As Long) As Long
Public Declare Function gm0_getunits Lib "gm0.dll" (ByVal handle As Long) As Long
Public Declare Function gm0_getmode Lib "gm0.dll" (ByVal handle As Long) As Long
Public Declare Function gm0_getvalue Lib "gm0.dll" (ByVal handle As Long) As Double

' null and peak detect
Public Declare Function gm0_donull Lib "gm0.dll" (ByVal handle As Long) As Long
Public Declare Function gm0_doaz Lib "gm0.dll" (ByVal handle As Long) As Long
Public Declare Function gm0_resetnull Lib "gm0.dll" (ByVal handle As Long) As Long
Public Declare Function gm0_resetpeak Lib "gm0.dll" (ByVal handle As Long) As Long

'time functions
Public Declare Function gm0_sendtime Lib "gm0.dll" (ByVal handle As Long, ByVal enable As Boolean) As Long
Public Declare Function gm0_settime2 Lib "gm0.dll" (ByVal handle As Long, time As gm_time) As Long
Public Declare Function gm0_gettime Lib "gm0.dll" (ByVal handle As Long) As gm_time
Public Declare Function gm0_getstore Lib "gm0.dll" (ByVal handle As Long, ByVal pos As Long) As gm_store

'call back functions
'Public Declare Function gm0_setcallback2 Lib "gm0.dll" (ByVal handle As Long, ByVal fn As Long) As Long
'Public Declare Function gm0_setconnectcallback Lib "gm0.dll" (ByVal handle As Long, ByVal fn As Long) As Long

'disable datamode functions
Public Declare Function gm0_startcmd Lib "gm0.dll" (ByVal handle As Long) As Long
Public Declare Function gm0_endcmd Lib "gm0.dll" (ByVal handle As Long) As Long

Function cleanup()

'    EndTimer

    If NOCOMM_MODE = True Then
    
        connected = False
        ConnectStatus = 0
        initstatus = False
        
        Exit Function
        
    End If

    'Add error handling
    On Error GoTo NoDLL:
    
        If (handle >= 0) Then
            gm0_killgm (handle)
            handle = -1
        End If
  
    On Error GoTo 0
  
    connected = False
    ConnectStatus = 0
    initstatus = False
    
NoDLL:

    'Do nothing
    
End Function

'*******************************************************************************************
'*                                                                                         *
'*                              Gaussmeter callbacks                                       *
'*                                                                                         *
'*******************************************************************************************

Sub connectcallback()

    ' DONT call ANY user UI code from here or excel WILL bomb

    'Check for NOCOMM_MODE
    If NOCOMM_MODE = True Then Exit Sub

    connected = True
    initstatus = False
    ConnectStatus = 2
    frm908AGaussmeter.connected
    
End Sub

Sub datacallback(ByVal gm_handle As Long)

    Dim temprange As Integer
    
    ' save the data localy
    
    'Check for NOCOMM
    If NOCOMM_MODE = True Then Exit Sub
    
    data.value = gm0_getvalue(gm_handle)
    data.Mode = gm0_getmode(gm_handle)
    data.Units = gm0_getunits(gm_handle)
    temprange = gm0_getrange(gm_handle)
    
    ' hide autorange flag from data.range
    If (temprange > 3) Then
        temprange = temprange - 4
    End If
    
    data.range = temprange
    
End Sub

Public Sub doaz()
        
  'Check for NOCOMM_MODE
  If NOCOMM_MODE = True Then Exit Sub
        
  If (handle < 0) Then Exit Sub
  MsgBox ("Shield probe and press enter")
  gm0_doaz handle
  MsgBox ("Auto zero finished")
   
End Sub

'*******************************************************************************************
'*                                                                                         *
'*                              Timer functions                                            *
'*                                                                                         *
'*******************************************************************************************

'Public Sub sampletosheetproc()
'    frm908AGaussmeter.writedatatosheet data, False
'    Application.OnTime Now + TimeValue("00:00:" + Str(mod908AGaussmeter.samplerate)), "sampletosheetproc", Schedule:=True
'End Sub

'Sub TimerProc()
'
'    If (connected = True) Then
'
'        frm908AGaussmeter.newdata
'
'        If (initstatus = False) Then
'            frm908AGaussmeter.enablebuttons
'            initstatus = True
'        End If
'    End If
'
'   Application.OnTime Now + TimeValue("00:00:01"), "TimerProc"
'
'End Sub


'Sub EndTimer()
'    ' end ALL timers and resume on error to ensure all are killed
'
'    On Error Resume Next
'
'    Application.OnTime Now + TimeValue("00:00:01"), "lookbusy", Schedule:=False
'
'    On Error Resume Next
'
'   Application.OnTime Now + TimeValue("00:00:01"), "TimerProc", Schedule:=False
'
'    On Error Resume Next
'
'   Application.OnTime Now + TimeValue("00:00:01"), "sampletosheetproc", Schedule:=False
'
'    On Error Resume Next
'
'   Application.OnTime Now + TimeValue("00:00:" + Str(mod908AGaussmeter.samplerate)), "sampletosheetproc", Schedule:=False
'
'
'
'On Error GoTo 0
'End Sub


'*******************************************************************************************
'*                                                                                         *
'*                              init and kill                                              *
'*                                                                                         *
'*******************************************************************************************


'Sub startconnect(mode)
'
'    dots = 7
'    Application.OnTime Now + TimeValue("00:00:01"), "lookbusy"
'
'
'
'End Sub

Sub doconnect(Mode)

    Dim doContinue As Boolean
    Dim StartTime
    Dim ElapsedTime

    'Check for NOCOMM mode
    If NOCOMM_MODE = True Then Exit Sub
    
    'Add in error checking

    On Error GoTo NoDLL:
        mod908AGaussmeter.handle = mod908AGaussmeter.gm0_newgm(Comport, Mode)
        mod908AGaussmeter.gm0_startconnect handle
        
        doContinue = False
    
        'Set the start time
        StartTime = timeGetTime()
    
        Do While doContinue = False
    
            If mod908AGaussmeter.gm0_getconnect(handle) Then
    
                doContinue = True
                
                DoEvents
    
            End If
            
            ElapsedTime = timeGetTime - StartTime
            
            'If trying to connect for more than 30 seconds
            'abort the connect process
            If ElapsedTime > 30000 Then
            
                'set connect status = -1
                'to indicate error
                
                'If partially connected, need to
                'kill the gaussmeter connection
                On Error Resume Next
                
                    mod908AGaussmeter.gm0_killgm (handle)
                    
                On Error GoTo 0
                
                mod908AGaussmeter.ConnectStatus = -1
                
                Exit Sub
                
            End If
                
        Loop
        
        mod908AGaussmeter.connectcallback
        
        'Wait for new data
        mod908AGaussmeter.waitfordata mod908AGaussmeter.handle
        
    '    'Check to see if PollTime has ever been assigned a value
    '    If IsEmpty(PollTime) Or IsNull(PollTime) Or isnothing(PollTime) Then
    '
    '        PollTime = 500
    '
    '    End If
    '
    '    If PollTime = 0 Then PollTime = 500
    '
    '    'Wait PollTime (ms)
    '    PauseTill timeGetTime() + PollTime
        
        mod908AGaussmeter.datacallback mod908AGaussmeter.handle
           
        ConnectStatus = 1
    
    On Error GoTo 0
    
NoDLL:

    'Do Nothing
    
End Sub

Public Sub donull()

  'Check for NOCOMM_MODE
  If NOCOMM_MODE = True Then Exit Sub
  
    If (handle < 0) Then Exit Sub
    MsgBox ("Sheild probe and press enter")
    gm0_donull handle
    MsgBox ("Null finished")

End Sub

Public Sub DoSilentNull()

  'Check for NOCOMM_MODE
  If NOCOMM_MODE = True Then Exit Sub
  
    If (handle < 0) Then Exit Sub
    gm0_donull handle
    

End Sub

Sub endcmdseq()

    If (handle < 0) Then Exit Sub
    gm0_endcmd (handle)
    
End Sub

Public Function getgmtime() As gm_time

    'Check for NOCOMM_MODE
  If NOCOMM_MODE = True Then Exit Function
  
    Dim thetime As gm_time
    If (handle < 0) Then Exit Function
    thetime = gm0_gettime(handle)
    getgmtime = thetime

End Function

Sub init()

    initstatus = False
    cleanup
    
    unitsrange(0, 0) = "T"
    unitsrange(0, 1) = "mT"
    unitsrange(0, 2) = "mT"
    unitsrange(0, 3) = "mT"
    
    unitsrangesclar(0, 0) = 1
    unitsrangesclar(0, 1) = 1000
    unitsrangesclar(0, 2) = 1000
    unitsrangesclar(0, 3) = 1000
    
    unitsrangefmt(0, 0) = " 0.000;-0.000;0.000"
    unitsrangefmt(0, 1) = " 000.0;-000.0;0.0"
    unitsrangefmt(0, 2) = " 00.00;-00.00;0.00"
    unitsrangefmt(0, 3) = " 0.000;-0.000;0.000"
         
    unitsrange(1, 0) = "kG"
    unitsrange(1, 1) = "kG"
    unitsrange(1, 2) = "G"
    unitsrange(1, 3) = "G"
    
    unitsrangefmt(1, 0) = " 00.00;-00.00;0.00"
    unitsrangefmt(1, 1) = " 0.000;-0.000;0.000"
    unitsrangefmt(1, 2) = " 000.0;-000.0;0.0"
    unitsrangefmt(1, 3) = " 00.00;-00.00;0.00"
    
    unitsrangesclar(1, 0) = 0.001
    unitsrangesclar(1, 1) = 0.001
    unitsrangesclar(1, 2) = 1
    unitsrangesclar(1, 3) = 1
 
    unitsrange(2, 0) = "kA/m"
    unitsrange(2, 1) = "kA/m"
    unitsrange(2, 2) = "kA/m"
    unitsrange(2, 3) = "kA/m"
    
    unitsrangefmt(2, 0) = " 0000;-0000;0"
    unitsrangefmt(2, 1) = " 000.0;-000.0;0.0"
    unitsrangefmt(2, 2) = " 00.00;-00.00;0.00"
    unitsrangefmt(2, 3) = " 0.000;-0.000;0.000"
      
    unitsrangesclar(2, 0) = 0.001
    unitsrangesclar(2, 1) = 0.001
    unitsrangesclar(2, 2) = 0.001
    unitsrangesclar(2, 3) = 0.001
         
    unitsrange(3, 0) = "kOe"
    unitsrange(3, 1) = "kOe"
    unitsrange(3, 2) = "Oe"
    unitsrange(3, 3) = "Oe"
    
    unitsrangefmt(3, 0) = " 00.00;-00.00;0.00"
    unitsrangefmt(3, 1) = " 000.0;-000.0;0.0"
    unitsrangefmt(3, 2) = " 0.000;-0.000;0.000"
    unitsrangefmt(3, 3) = " 00.00;-00.00;0.00"
  
    unitsrangesclar(3, 0) = 0.001
    unitsrangesclar(3, 1) = 0.001
    unitsrangesclar(3, 2) = 1
    unitsrangesclar(3, 3) = 1
  
    baseunits(0) = "T"
    baseunits(1) = "G"
    baseunits(2) = "A/m"
    baseunits(3) = "Oe"
            
    modestr(0) = "DC"
    modestr(1) = "DC Pk"
    modestr(2) = "AC"
    modestr(3) = "AC Mx"
    modestr(4) = "AX Pk"
    
End Sub

'*******************************************************************************************
'*                                                                                         *
'*                              Helper functions                                           *
'*                                                                                         *
'*******************************************************************************************

Public Function makeactualvalue(lData As gm_store) As Double
    ' take a value as per gm0 output and convert into correct value based on range and units
    Dim tempvalue As Double
    makeactualvalue = lData.value * unitsrangesclar(lData.Units, lData.range)
    
End Function

'*******************************************************************************************
'*                                                                                         *
'*                              Gaussmeter command wrappers                                *
'*                                                                                         *
'*******************************************************************************************
Public Sub resetpeak()

    'Check for NOCOMM_MODE
    If NOCOMM_MODE = True Then Exit Sub
    
    If (handle < 0) Then Exit Sub
    gm0_resetpeak handle
    
End Sub

Public Sub run_gaussmeter()
    frm908AGaussmeter.Show False
End Sub

Public Sub setmode(Mode As Integer)

    'Check for NOCOMM_MODE
    If NOCOMM_MODE = True Then Exit Sub
  
    If (handle < 0) Then Exit Sub
    gm0_setmode handle, Mode

End Sub

Public Sub setrange(range As Integer)

    'Check for NOCOMM_MODE
    If NOCOMM_MODE = True Then Exit Sub
  
    If (handle < 0) Then Exit Sub
    gm0_setrange handle, range

End Sub

'Function getregister(reg As Integer) As gm_store
'
'    Dim temp As gm_store
'    If (handle < 0) Then Exit Function
'
'    If (reg >= 0 And reg < 100) Then
'        temp = gm0_getstore(handle, reg)
'        getregister = temp
'    Else
'        Beep
'    End If
'
'End Function

Public Sub setsystime()

    Dim thetime As gm_time

    'Check for NOCOMM_MODE
    If NOCOMM_MODE = True Then Exit Sub
      
    If (handle < 0) Then Exit Sub
    
    thetime.year = year(Now) - 2000
    thetime.month = month(Now)
    thetime.day = day(Now)
    thetime.hour = hour(Now)
    thetime.min = Minute(Now)
    thetime.sec = Second(Now)
    gm0_settime2 handle, thetime
    
    MsgBox ("GM05 time set to system time")

End Sub

Public Sub SetUnits(Units As Integer)

    'Check for NOCOMM_MODE
    If NOCOMM_MODE = True Then Exit Sub
  
    If (handle < 0) Then Exit Sub
    gm0_setunits handle, Units

End Sub

Sub startcmdseq()

    If (handle < 0) Then Exit Sub
        On Error Resume Next
        gm0_startcmd (handle)
        
End Sub

'Sub lookbusy()
'    Dim dotstr(6)
'
'     DoEvents
'
'    dotstr(1) = ""
'    dotstr(2) = "."
'    dotstr(3) = ".."
'    dotstr(4) = "..."
'    dotstr(5) = "...."
'    dotstr(6) = "....."
'
'    dots = dots + 1
'
'    If (dots >= 7) Then
'        dots = 1
'        switchcomspeed
'    End If
'
'    DoEvents
'    frm908AGaussmeter.Display.Value = "Connecting" + dotstr(dots)
'    DoEvents
'    If (connected = False) Then
'        Application.OnTime Now + TimeValue("00:00:01"), "lookbusy"
'    End If
'
'End Sub

Sub switchcomspeed()

        If (mod908AGaussmeter.handle >= 0) Then
            gm0_killgm (mod908AGaussmeter.handle)
            mod908AGaussmeter.handle = -1
        End If
        
        If (mod908AGaussmeter.Comport > 0) Then
            Mode = Mode + 1
            If Mode = 2 Then Mode = 0
        End If
        
        doconnect Mode
        
End Sub

Sub waitfordata(ByVal gm_handle As Long)

    Dim doContinue As Boolean
    
    doContinue = False
    
    If connected Then
    
        'Wait for new data to be available from the gaussmeter
        Do While Not doContinue
        
            If (mod908AGaussmeter.gm0_isnewdata(gm_handle)) Then
            
                doContinue = True
                
            End If
            
            DoEvents
            
            'Wait 1 ms
            PauseTill timeGetTime() + 1
            
        Loop
        
    End If

End Sub

Public Sub StartLogDataFile(ByRef file_name As String)

    'Set Path for log file
    Dim folder_path As String
    folder_path = modConfig.ADWIN_AFDataLocalDir & "\Gaussmeter Data Logs\"
                        
    Dim fso As FileSystemObject
    Set fso = New FileSystemObject
    
    If Not fso.FolderExists(folder_path) Then
    
        fso.CreateFolder (folder_path)
        
    End If
    
    If Strings.Right(file_name, 4) <> ".csv" Then
        file_name = file_name & ".csv"
    End If
    
    If fso.FileExists(file_name) Then
    
        file_name = Strings.Replace(file_name, ".csv", "_duplicatefilename.csv")
        
    End If
    
    Dim fs As TextStream
    Set fs = fso.CreateTextFile(folder_path & file_name, True)
    
    Dim header_line As String
    
    header_line = "908A Gaussmeter Data Log"
    fs.WriteLine (header_line)
    
    header_line = "Started:," & Format(Now, "yyyy/mm/dd, hh:nn:ss")
    fs.WriteLine (header_line)
    fs.WriteBlankLines (1)
    
    header_line = "Data,Range,Mode,Probe,Date,Time"
    fs.WriteLine (header_line)
    
    fs.Close
    
    Set fso = Nothing

End Sub

Public Sub LogDataToFile(ByVal gm_data As String, _
                          ByVal gm_range As String, _
                          ByVal gm_mode As String, _
                          ByVal gm_probe_name As String, _
                          ByVal date_time As Date, _
                          ByVal file_name As String)
                               
                               
    'Set Path for log file
    Dim folder_path As String
    folder_path = modConfig.ADWIN_AFDataLocalDir & "\Gaussmeter Data Logs\"
                        
    Dim fso As FileSystemObject
    Set fso = New FileSystemObject
    
    If Not fso.FolderExists(folder_path) Then
    
        fso.CreateFolder (folder_path)
        
    End If
    
    Dim fs As TextStream
    Set fs = fso.OpenTextFile(folder_path & file_name, ForAppending, True)
    
    Dim line_str As String
    Dim date_str As String
    
    date_str = Format(date_time, "yyyy/mm/dd, hh:nn:ss")
    
    line_str = gm_data & "," & _
               gm_range & "," & _
               gm_mode & "," & _
               gm_probe_name & "," & _
               date_str
               
    fs.WriteLine (line_str)
    
    fs.Close

    Set fso = Nothing
                               
End Sub

