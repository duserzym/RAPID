VERSION 5.00
Begin VB.Form frm908AGaussmeter 
   Caption         =   "Gaussmeter Control"
   ClientHeight    =   6015
   ClientLeft      =   14595
   ClientTop       =   5730
   ClientWidth     =   6990
   Icon            =   "frm908AGaussmeter.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6015
   ScaleWidth      =   6990
   Begin VB.CommandButton cmdGetTime 
      Caption         =   "Get Gaussmeter Time"
      Height          =   372
      Left            =   2040
      TabIndex        =   27
      Top             =   5040
      Width           =   1812
   End
   Begin VB.CommandButton cmdSetTime 
      Caption         =   "Set Gaussmeter Time"
      Height          =   372
      Left            =   120
      TabIndex        =   26
      Top             =   5040
      Width           =   1812
   End
   Begin VB.Frame Frame4 
      Caption         =   "Trigger"
      Height          =   2412
      Left            =   3960
      TabIndex        =   20
      Top             =   3480
      Width           =   2892
      Begin VB.CommandButton cmdSampleOnAlarm 
         Caption         =   "Sample On Alarm"
         Height          =   372
         Left            =   360
         TabIndex        =   28
         Top             =   1320
         Width           =   2292
      End
      Begin VB.TextBox txtAlarmLow 
         Height          =   288
         Left            =   1560
         TabIndex        =   24
         Top             =   840
         Width           =   972
      End
      Begin VB.TextBox txtAlarmHigh 
         Height          =   288
         Left            =   1560
         TabIndex        =   23
         Top             =   360
         Width           =   972
      End
      Begin VB.Label Label6 
         Caption         =   "Note: Values must be in the units selected to the left."
         Height          =   372
         Left            =   120
         TabIndex        =   25
         Top             =   1920
         Width           =   2532
      End
      Begin VB.Label Label5 
         Caption         =   "Alarm Level  -"
         Height          =   252
         Left            =   240
         TabIndex        =   22
         Top             =   840
         Width           =   1212
      End
      Begin VB.Label Label4 
         Caption         =   "Alarm Level  +"
         Height          =   252
         Left            =   240
         TabIndex        =   21
         Top             =   360
         Width           =   1212
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Measure"
      Height          =   3252
      Left            =   3960
      TabIndex        =   12
      Top             =   120
      Width           =   2892
      Begin VB.CommandButton cmdClearData 
         Caption         =   "Clear Data"
         Height          =   372
         Left            =   360
         TabIndex        =   29
         Top             =   2640
         Width           =   2172
      End
      Begin VB.CommandButton cmdSampleNow 
         Caption         =   "Sample Once Now"
         Height          =   372
         Left            =   360
         TabIndex        =   19
         Top             =   2160
         Width           =   2172
      End
      Begin VB.TextBox txtNumSamples 
         Height          =   288
         Left            =   1800
         TabIndex        =   17
         Top             =   840
         Width           =   732
      End
      Begin VB.CommandButton cmdStartSampling 
         Caption         =   "Start Sampling"
         Height          =   372
         Left            =   360
         TabIndex        =   15
         Top             =   1680
         Width           =   2172
      End
      Begin VB.TextBox txtRate 
         Height          =   288
         Left            =   1800
         TabIndex        =   14
         Top             =   360
         Width           =   732
      End
      Begin VB.Label lblSamplingDuration 
         Caption         =   "Sampling Duration:"
         Height          =   252
         Left            =   240
         TabIndex        =   18
         Top             =   1320
         Width           =   2292
      End
      Begin VB.Label Label3 
         Caption         =   "Number Samples:"
         Height          =   372
         Left            =   240
         TabIndex        =   16
         Top             =   840
         Width           =   1452
      End
      Begin VB.Label Label2 
         Caption         =   "Sampling rate (ms):"
         Height          =   372
         Left            =   240
         TabIndex        =   13
         Top             =   360
         Width           =   1452
      End
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H80000010&
      Caption         =   "Close"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   120
      TabIndex        =   11
      Top             =   5520
      Width           =   3732
   End
   Begin VB.Frame Frame2 
      Caption         =   "Gaussmeter reading"
      Height          =   3612
      Left            =   120
      TabIndex        =   6
      Top             =   1320
      Width           =   3732
      Begin VB.CommandButton cmdResetPeak 
         Caption         =   "Reset Peak"
         Height          =   372
         Left            =   1920
         TabIndex        =   48
         Top             =   2640
         Width           =   1692
      End
      Begin VB.Frame Frame7 
         BorderStyle     =   0  'None
         Caption         =   "Frame7"
         Height          =   1452
         Left            =   120
         TabIndex        =   41
         Top             =   960
         Width           =   1212
         Begin VB.OptionButton optRange 
            Caption         =   "Range 0"
            Height          =   252
            Index           =   0
            Left            =   120
            TabIndex        =   45
            Top             =   120
            Width           =   972
         End
         Begin VB.OptionButton optRange 
            Caption         =   "Range 1"
            Height          =   252
            Index           =   1
            Left            =   120
            TabIndex        =   44
            Top             =   480
            Width           =   972
         End
         Begin VB.OptionButton optRange 
            Caption         =   "Range 2"
            Height          =   252
            Index           =   2
            Left            =   120
            TabIndex        =   43
            Top             =   840
            Width           =   972
         End
         Begin VB.OptionButton optRange 
            Caption         =   "Range 3"
            Height          =   252
            Index           =   3
            Left            =   120
            TabIndex        =   42
            Top             =   1200
            Width           =   972
         End
      End
      Begin VB.Frame Frame6 
         BorderStyle     =   0  'None
         Caption         =   "Frame6"
         Height          =   1212
         Left            =   2400
         TabIndex        =   36
         Top             =   960
         Width           =   1212
         Begin VB.OptionButton optUnits 
            Caption         =   "Tesla"
            Height          =   252
            Index           =   0
            Left            =   120
            TabIndex        =   40
            Top             =   120
            Width           =   852
         End
         Begin VB.OptionButton optUnits 
            Caption         =   "Gauss"
            Height          =   252
            Index           =   1
            Left            =   120
            TabIndex        =   39
            Top             =   360
            Width           =   852
         End
         Begin VB.OptionButton optUnits 
            Caption         =   "KA/m"
            Height          =   252
            Index           =   2
            Left            =   120
            TabIndex        =   38
            Top             =   600
            Width           =   852
         End
         Begin VB.OptionButton optUnits 
            Caption         =   "Oersted"
            Height          =   252
            Index           =   3
            Left            =   120
            TabIndex        =   37
            Top             =   840
            Width           =   852
         End
      End
      Begin VB.Frame Frame5 
         BorderStyle     =   0  'None
         Caption         =   "Frame5"
         Height          =   1572
         Left            =   1200
         TabIndex        =   30
         Top             =   960
         Width           =   1212
         Begin VB.OptionButton optFunction 
            Caption         =   "DC"
            Height          =   312
            Index           =   0
            Left            =   120
            TabIndex        =   35
            Top             =   120
            Width           =   972
         End
         Begin VB.OptionButton optFunction 
            Caption         =   "DC PK"
            Height          =   312
            Index           =   1
            Left            =   120
            TabIndex        =   34
            Top             =   360
            Width           =   972
         End
         Begin VB.OptionButton optFunction 
            Caption         =   "AC"
            Height          =   312
            Index           =   2
            Left            =   120
            TabIndex        =   33
            Top             =   600
            Width           =   972
         End
         Begin VB.OptionButton optFunction 
            Caption         =   "AC MAX"
            Height          =   312
            Index           =   3
            Left            =   120
            TabIndex        =   32
            Top             =   840
            Width           =   972
         End
         Begin VB.OptionButton optFunction 
            Caption         =   "AC MAX PK"
            Height          =   432
            Index           =   4
            Left            =   120
            TabIndex        =   31
            Top             =   1080
            Width           =   1332
         End
      End
      Begin VB.CommandButton cmdAutoZero 
         Caption         =   "Auto Zero"
         Height          =   372
         Left            =   1920
         TabIndex        =   10
         Top             =   3120
         Width           =   1692
      End
      Begin VB.CommandButton cmdNull 
         Caption         =   "Auto Null"
         Height          =   372
         Left            =   120
         TabIndex        =   9
         Top             =   3120
         Width           =   1692
      End
      Begin VB.CommandButton cmdAutoRange 
         Caption         =   "Auto Range"
         Height          =   372
         Left            =   120
         TabIndex        =   8
         Top             =   2640
         Width           =   1692
      End
      Begin VB.TextBox txtDisplay 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   22.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   612
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Width           =   3252
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Connection"
      Height          =   1092
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3732
      Begin VB.TextBox txtPollTime 
         Height          =   288
         Left            =   2760
         TabIndex        =   47
         Top             =   600
         Width           =   732
      End
      Begin VB.CommandButton cmdConnectButton 
         Caption         =   "Connect"
         Height          =   372
         Left            =   1080
         TabIndex        =   5
         Top             =   600
         Width           =   1332
      End
      Begin VB.TextBox txtComPort 
         Height          =   288
         Left            =   2040
         TabIndex        =   4
         Top             =   204
         Width           =   372
      End
      Begin VB.OptionButton optCommMode 
         Caption         =   "USB"
         Height          =   192
         Index           =   1
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   972
      End
      Begin VB.OptionButton optCommMode 
         Caption         =   "RS232"
         Height          =   192
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   852
      End
      Begin VB.Label Label7 
         Caption         =   "Poll Time (ms):"
         Height          =   252
         Left            =   2520
         TabIndex        =   46
         Top             =   360
         Width           =   1092
      End
      Begin VB.Label Label1 
         Caption         =   "COM Port:"
         Height          =   372
         Left            =   1200
         TabIndex        =   3
         Top             =   240
         Width           =   732
      End
   End
End
Attribute VB_Name = "frm908AGaussmeter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim OldData As gm_store
Dim ActualValue As Double

Dim SampleActive As Boolean
Dim SampleOnAlarm As Boolean
Dim UseBaseUnits As Boolean
Dim optIgnore As Boolean
Dim ButtonsEnabled As Boolean

Dim AlarmHigh As Double
Dim AlarmLow As Double

Dim DataArray() As Double

Private Sub cmdAutoRange_click()
    
    On Error GoTo GaussmeterError:
        
        mod908AGaussmeter.setrange (4)
    
        mod908AGaussmeter.waitfordata mod908AGaussmeter.handle
        mod908AGaussmeter.datacallback mod908AGaussmeter.handle
      
        Me.newdata

    On Error GoTo 0
    
GaussmeterError:

End Sub

Private Sub cmdAutoZero_Click()
    
    txtDisplay.text = "AZ in progress"
    mod908AGaussmeter.doaz

End Sub

Public Sub cmdClearData_Click()

    ReDim mod908AGaussmeter.DataArray(1)

    
'    Debug.Print UBound(mod908AGaussmeter.DataArray)
    cmdClearData.Enabled = False

End Sub

Private Sub cmdClose_Click()
    
    On Error GoTo GaussmeterError:
    
        mod908AGaussmeter.cleanup
        Me.Hide
        
    On Error GoTo 0
    
GaussmeterError:

End Sub

Private Sub cmdConnectButton_Click()
   
    If (mod908AGaussmeter.ConnectStatus = 0) Then
        
        Connect
        
    Else
        
        Disconnect
        
    End If
   
End Sub

'*******************************************************************************************
'*                                                                                         *
'*                              Time control                                               *
'*                                                                                         *
'*******************************************************************************************

Private Sub cmdGetTime_Click()
    
    Dim thetime As gm_time
    thetime = mod908AGaussmeter.getgmtime
    MsgBox ("The Gaussmeter Time is " & Trim(str(thetime.day)) & "/" & _
            Trim(str(thetime.month)) & "/" & Trim(str(thetime.year)) & " " & _
            Trim(str(thetime.hour)) & ":" & Trim(str(thetime.min)) & ":" & _
            Trim(str(thetime.sec)))

End Sub

'*******************************************************************************************
'*                                                                                         *
'*                              Null and Auto-Zero                                         *
'*                                                                                         *
'*******************************************************************************************

Private Sub cmdNull_Click()

    On Error GoTo GaussmeterError:

        txtDisplay.text = "NULL in progress"
        mod908AGaussmeter.donull
        
    On Error GoTo 0
    
GaussmeterError:
        
End Sub

Public Sub cmdResetPeak_Click()

    On Error GoTo GaussmeterError:

        'Call gm0 shell function
        mod908AGaussmeter.resetpeak
        
        'Wait for new data
        mod908AGaussmeter.waitfordata mod908AGaussmeter.handle
        
        'Get new data
        mod908AGaussmeter.datacallback mod908AGaussmeter.handle
        
        Me.newdata
    
    On Error GoTo 0
    
GaussmeterError:
    
End Sub

Public Sub cmdSampleNow_Click()

    Dim ArraySize As Long

    On Error GoTo GaussmeterError:
    
        ArraySize = UBound(mod908AGaussmeter.DataArray)
    
        If Err.number <> 0 Then
        
            ReDim mod908AGaussmeter.DataArray(1)
            cmdClearData.Enabled = True
            
            ArraySize = 1
            
        Else
    
            ArraySize = ArraySize + 1
            ReDim Preserve mod908AGaussmeter.DataArray(ArraySize)
            cmdClearData.Enabled = True
            
        End If
           
        mod908AGaussmeter.waitfordata mod908AGaussmeter.handle
        mod908AGaussmeter.datacallback mod908AGaussmeter.handle
    
        Me.newdata
        
        mod908AGaussmeter.DataArray(ArraySize - 1) = mod908AGaussmeter.data
        
    On Error GoTo 0
    
    Exit Sub
    
GaussmeterError:

End Sub

'*******************************************************************************************
'*                                                                                         *
'*                              Sample to sheet and by alarm                               *
'*                                                                                         *
'*******************************************************************************************

Private Sub cmdSampleOnAlarm_Click()
    If (SampleOnAlarm = False) Then
        SampleOnAlarm = True
        Me.cmdStartSampling.Enabled = False
        Me.txtAlarmHigh.Enabled = False
        Me.txtAlarmLow.Enabled = False
        AlarmLow = val(Me.txtAlarmLow.text)
        AlarmHigh = val(Me.txtAlarmHigh.text)
        Me.cmdSampleOnAlarm.Caption = "Turn Alarm Off"
    Else
        SampleOnAlarm = False
        Me.cmdStartSampling.Enabled = True
        Me.txtAlarmHigh.Enabled = True
        Me.txtAlarmLow.Enabled = True
        Me.cmdSampleOnAlarm.Caption = "Sample on alarm"
    End If
    
End Sub

Private Sub cmdSetTime_Click()
    mod908AGaussmeter.setsystime
End Sub

Private Sub cmdStartSampling_Click()

    Dim i As Long
    Dim TimeInterval As Double
    Dim StartTime As Double
    Dim EndTime As Double
    Dim ArraySize As Long
    Dim N As Long
    Dim StartPoint As Long
    Dim doStop
    
    N = val(Me.txtNumSamples)
    TimeInterval = val(Me.txtRate)
        
    On Error GoTo GaussmeterError:
        
        StartPoint = UBound(mod908AGaussmeter.DataArray)
    
        If Err.number <> 0 Then
        
            ReDim mod908AGaussmeter.DataArray(1)
            cmdClearData.Enabled = True
            
            StartPoint = 0
            
        End If
        
        i = StartPoint
        
        StartTime = timeGetTime()
        EndTime = StartTime + N * TimeInterval
        
        doStop = False
        
        Do
        
            mod908AGaussmeter.datacallback mod908AGaussmeter.handle
      
            Me.newdata
        
            If SampleOnAlarm = True _
                And (mod908AGaussmeter.data.value < AlarmLow Or _
                     mod908AGaussmeter.data.value > AlarmHigh) Then
            
                'Do Nothing, we're outside the alarm zone
                
            Else
            
                'We're in the alarm zone, record the data
                mod908AGaussmeter.DataArray(i) = mod908AGaussmeter.data
                i = i + 1
                
            End If
                
            PauseTill timeGetTime() + TimeInterval
        
            If i >= StartPoint + N Or EndTime >= timeGetTime() Then
            
                doStop = True
                
            End If
            
            DoEvents
        
        Loop Until doStop = True
    
    On Error GoTo 0
    
GaussmeterError:
    
End Sub

Public Function Connect() As Long

    Dim ComNum As Integer

    On Error GoTo GaussmeterError:
    
        If mod908AGaussmeter.ConnectStatus = 1 Then
        
            'Gaussmeter is already connected
            Connect = 1
            cmdConnectButton.Caption = "Disconnect"
            
            Exit Function
            
        End If
    
        If (optCommMode(0) = True) Then
            ComNum = -1
        Else
            ComNum = val(txtComPort.text)
        End If
        
        mod908AGaussmeter.Comport = ComNum
        cmdConnectButton.Caption = "Disconnect"
        txtDisplay.text = "Connecting"
        mod908AGaussmeter.PollTime = CLng(val(Me.txtPollTime))
        mod908AGaussmeter.doconnect 1
        
        Me.newdata
    
        Connect = mod908AGaussmeter.ConnectStatus
        If Connect = 0 Then
         cmdConnectButton.Caption = "Connect"
        Else
        cmdConnectButton.Caption = "Disconnect"
        End If

    On Error GoTo 0
    
GaussmeterError:

End Function

Sub connected()
    
    ' DONT DO UI STUFF FROM HERE
    ' A simple caption seems ok but more will crash excel
    ' its a callback from a thread so the hwnds will be NULL!!!!!!
    cmdConnectButton.Caption = "Disconnect"
    
End Sub

Public Sub ConvertLastData(ByRef gmReading As String, ByVal Units As String)

    Dim N As Long
    Dim lData As gm_store

    On Error GoTo GaussmeterError:
        
        N = UBound(mod908AGaussmeter.DataArray)
        
        If Err.number <> 0 Then
        
            gmReading = "ERR," & Trim(str(Err.number))
               
            Exit Sub
            
        End If
            
        lData = mod908AGaussmeter.DataArray(N - 1)
        
        If Units <> unitsrange(lData.Units, lData.range) Then
        
            If "k" & Units = unitsrange(lData.Units, lData.range) _
                Or Units = "m" & unitsrange(lData.Units, lData.range) Then
            
                lData.value = lData.value * 1000
                
            End If
            
        End If
        
        gmReading = Format(makeactualvalue(lData), _
                    unitsrangefmt(lData.Units, lData.range))

    On Error GoTo 0
    
GaussmeterError:

End Sub

Public Function CurrentRange() As Long

    If Me.optRange(0).value = True Then
    
        CurrentRange = 0
        
    ElseIf Me.optRange(1).value = True Then
    
        CurrentRange = 1
        
    ElseIf Me.optRange(2).value = True Then
    
        CurrentRange = 2
    
    ElseIf Me.optRange(3).value = True Then
    
        CurrentRange = 3
        
    End If
    
End Function

'*******************************************************************************************
'*                                                                                         *
'*                              Enable/disable User Interface                              *
'*                                                                                         *
'*******************************************************************************************

Sub disablebuttons()

ButtonsEnabled = False

Me.cmdStartSampling.Enabled = False
Me.cmdSampleOnAlarm.Enabled = False
Me.cmdNull.Enabled = False
Me.cmdAutoRange.Enabled = False
Me.cmdAutoZero.Enabled = False
Me.cmdGetTime.Enabled = False
Me.cmdSetTime.Enabled = False

Me.optRange(0).Enabled = False
Me.optRange(1).Enabled = False
Me.optRange(2).Enabled = False
Me.optRange(3).Enabled = False

Me.optFunction(0).Enabled = False
Me.optFunction(1).Enabled = False
Me.optFunction(2).Enabled = False
Me.optFunction(3).Enabled = False
Me.optFunction(4).Enabled = False

Me.optUnits(0).Enabled = False
Me.optUnits(1).Enabled = False
Me.optUnits(2).Enabled = False
Me.optUnits(3).Enabled = False

Me.cmdSampleNow.Enabled = False

End Sub

Public Function Disconnect() As Long

    On Error GoTo GaussmeterError:
    
        If mod908AGaussmeter.ConnectStatus = 0 Then
        
            'Gaussmeter is already disconnected
            Disconnect = 0
            
            Exit Function
            
        End If
    
        mod908AGaussmeter.init
        mod908AGaussmeter.cleanup
        cmdConnectButton.Caption = "Connect"
        txtDisplay.text = "Disconnected"
        disablebuttons
    
        Disconnect = mod908AGaussmeter.ConnectStatus

    On Error GoTo 0
    
GaussmeterError:

End Function

Public Function DoSilentNull() As Boolean

    On Error GoTo GaussmeterError:

        txtDisplay.text = "NULL in progress"
        
        'Pause 1 second
        PauseTill timeGetTime() + 1000
        mod908AGaussmeter.DoSilentNull
        
        'Pause 3 seconds
        PauseTill timeGetTime() + 3000
        
        Me.newdata
        
    On Error GoTo 0
    
    DoSilentNull = True
    Exit Function
    
GaussmeterError:

    DoSilentNull = False
        
End Function

Sub enablebuttons()

Me.cmdStartSampling.Enabled = True
Me.cmdSampleOnAlarm.Enabled = True
Me.cmdNull.Enabled = True
Me.cmdAutoRange.Enabled = True
Me.cmdAutoZero.Enabled = True
Me.cmdGetTime.Enabled = True
Me.cmdSetTime.Enabled = True

Me.optRange(0).Enabled = True
Me.optRange(1).Enabled = True
Me.optRange(2).Enabled = True
Me.optRange(3).Enabled = True

Me.optFunction(0).Enabled = True
Me.optFunction(1).Enabled = True
Me.optFunction(2).Enabled = True
Me.optFunction(3).Enabled = True
Me.optFunction(4).Enabled = True

Me.optUnits(0).Enabled = True
Me.optUnits(1).Enabled = True
Me.optUnits(2).Enabled = True
Me.optUnits(3).Enabled = True

Me.cmdSampleNow.Enabled = True

Me.refresh

ButtonsEnabled = True

End Sub

'*******************************************************************************************
'*                                                                                         *
'*                              Init and destructor                                        *
'*                                                                                         *
'*******************************************************************************************
Private Sub Form_Load()

    LoadForm
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    mod908AGaussmeter.cleanup
    
    
End Sub

'---------------------------------------------------------------------------------------------------------'
'---------------------------------------------------------------------------------------------------------'
'
'   Code for sampling and saving data from the analog output of the Hirst 908A Gaussmeter using an
'   analog input channel on the PCI-DAS6030 board.
'
'   Created: October, 2010
'    Author: Isaac Hilburn
'
'   This code is not fully integrated into the settings / object hierarchy scheme
'   for the code.  The PCI-DAS6030 board is hard-coded here as the board to use.
'   If this board is phased out in the future, then the Gauss meter data-sampling and
'   data recording code procedures below will not work.  They will cause a collections
'   bad key error in the Boards class Public Property Get Item(ByVal String).
'
'---------------------------------------------------------------------------------------------------------'
'---------------------------------------------------------------------------------------------------------'

Public Sub InitializeDCFieldRecord(ByRef InWave As Wave, _
                                   ByVal RampTimeEst_msec As Long)

    With InWave
    
        Set .BoardUsed = SystemBoards("PCI-DAS6030")
        Set .Chan = .BoardUsed.AInChannels.Item(3)
        .IOOptions = BACKGROUND
        .DoDeallocate = True
        .BufferAlloc = False
        .IORate = 50000
        .NumPoints = (.IORate * RampTimeEst_msec) \ 1000    'Convert msec to sec
        
        Set .range = New range
        .range.RangeType = BIP10VOLTS
        .StartPoint = 0
        .TimeStep = 1 / .IORate
        .WaveName = "Gaussmeter Analog Input"
                
    End With

End Sub

Public Sub LoadForm()

    'Set Form dimensions
    Me.Height = 6525
    Me.Width = 7110
    
    '(March 11, 2011 - I Hilburn)
    'Update all the form's icons with the program Ico file
    SetFormIcon Me
    
    Me.optCommMode(1).value = False
    Me.optCommMode(1).value = True
    mod908AGaussmeter.init
    mod908AGaussmeter.cleanup
    SampleActive = False
    cmdConnectButton.Caption = "Connect"
    txtDisplay.text = "Disconnected"
    txtPollTime.text = "500"
    disablebuttons
    OldData.Mode = 10
    OldData.range = 10
    OldData.Units = 10
    UseBaseUnits = False
    mod908AGaussmeter.sampleindex = 1
    
    ReDim mod908AGaussmeter.DataArray(1)

End Sub

'*******************************************************************************************
'*                                                                                         *
'*                              Data display                                               *
'*                                                                                         *
'*******************************************************************************************

Sub newdata()

    Dim ActualValue As Double
    Dim value As Double
    
    On Error GoTo GaussmeterError:
        
        If Not ButtonsEnabled Then
        
            enablebuttons
            
        End If
    
        value = mod908AGaussmeter.makeactualvalue(mod908AGaussmeter.data)
    '    Debug.Print "Function = " & mod908AGaussmeter.Data.Mode
        txtDisplay.text = Format(value, _
                                 mod908AGaussmeter.unitsrangefmt( _
                                    mod908AGaussmeter.data.Units, _
                                    mod908AGaussmeter.data.range)) & " " & _
                                 mod908AGaussmeter.unitsrange( _
                                    mod908AGaussmeter.data.Units, _
                                    mod908AGaussmeter.data.range) & " " & _
                                    mod908AGaussmeter.modestr(mod908AGaussmeter.data.Mode)
    
        If (mod908AGaussmeter.data.Mode <> OldData.Mode) Then
            newmode (mod908AGaussmeter.data.Mode)
        End If
            
        If (mod908AGaussmeter.data.range <> OldData.range) Then
            newrange (mod908AGaussmeter.data.range)
        End If
        
        If (mod908AGaussmeter.data.Units <> OldData.Units) Then
           NewUnits (mod908AGaussmeter.data.Units)
        End If
            
        OldData = mod908AGaussmeter.data
    
    On Error GoTo 0
    
GaussmeterError:
    
End Sub

Sub newmode(Mode As Integer)

    Dim i As Integer

    For i = 0 To optFunction.Count - 1

        optFunction.Item(i) = False
        
    Next i
    
    optIgnore = True
    optFunction.Item(Mode) = True
    optIgnore = False

End Sub

Sub newrange(Index As Integer)

    Dim i As Integer

    For i = 0 To optRange.Count - 1

        optRange.Item(i) = False
        
    Next i
    
    optIgnore = True
    optRange.Item(Index) = True
    optIgnore = False

End Sub

Sub NewUnits(Units As Integer)

    Dim i As Long
    
    For i = 0 To optUnits.Count - 1
    
        optUnits.Item(i) = False
        
    Next i
    
    optIgnore = True
    optUnits.Item(Units) = True
    optIgnore = False

End Sub

Private Sub optCommMode_Click(Index As Integer)

    If Index = 0 Then
    
        txtComPort.text = "1"
        txtComPort.Enabled = True
        
    Else
    
        txtComPort.text = "-1"
        txtComPort.Enabled = False
        
    End If

End Sub

'*******************************************************************************************
'*                                                                                         *
'*                              Function Mode control                                      *
'*                                                                                         *
'*******************************************************************************************
Private Sub optFunction_Click(Index As Integer)

    Dim doContinue As Boolean

    On Error GoTo GaussmeterError:
    
        If Not optIgnore Then
        
            mod908AGaussmeter.setmode (Index)
            
            mod908AGaussmeter.waitfordata mod908AGaussmeter.handle
            mod908AGaussmeter.datacallback mod908AGaussmeter.handle
      
            doContinue = False
            
            Do While Not doContinue
      
                If Not mod908AGaussmeter.data.Mode = Index Then
                
                    mod908AGaussmeter.waitfordata mod908AGaussmeter.handle
                    mod908AGaussmeter.datacallback mod908AGaussmeter.handle
                    
                Else
                
                    doContinue = True
                    
                End If
                
                DoEvents
                
            Loop
            
            Me.newdata
      
        End If

    On Error GoTo 0
    
GaussmeterError:

End Sub

'*******************************************************************************************
'*                                                                                         *
'*                              Range control                                              *
'*                                                                                         *
'*******************************************************************************************

Private Sub optRange_Click(Index As Integer)

    Dim doContinue As Boolean

    On Error GoTo GaussmeterError:
    
        If Not optIgnore Then
    
            mod908AGaussmeter.setrange (Index)
    
            mod908AGaussmeter.waitfordata mod908AGaussmeter.handle
            mod908AGaussmeter.datacallback mod908AGaussmeter.handle
      
            doContinue = False
            
            Do While Not doContinue
      
                If Not mod908AGaussmeter.data.range = Index Then
                
                    mod908AGaussmeter.waitfordata mod908AGaussmeter.handle
                    mod908AGaussmeter.datacallback mod908AGaussmeter.handle
                    
                Else
                
                    doContinue = True
                    
                End If
                
                DoEvents
                
            Loop
        
            Me.newdata
    
        End If
        
    On Error GoTo 0

GaussmeterError:

End Sub

'*******************************************************************************************
'*                                                                                         *
'*                              Units control                                              *
'*                                                                                         *
'*******************************************************************************************

Private Sub optUnits_Click(Index As Integer)

    Dim doContinue As Boolean

    On Error GoTo GaussmeterError:
    
        If Not optIgnore Then
        
            mod908AGaussmeter.SetUnits (Index)
            
            mod908AGaussmeter.waitfordata mod908AGaussmeter.handle
            mod908AGaussmeter.datacallback mod908AGaussmeter.handle
      
            doContinue = False
            
            Do While Not doContinue
      
                If Not mod908AGaussmeter.data.Units = Index Then
                
                    mod908AGaussmeter.waitfordata mod908AGaussmeter.handle
                    mod908AGaussmeter.datacallback mod908AGaussmeter.handle
                    
                Else
                
                    doContinue = True
                    
                End If
                
                DoEvents
                
            Loop
      
            Me.newdata
            
        End If

    On Error GoTo 0

GaussmeterError:

End Sub

Public Sub SaveDCFieldRecord(ByRef DCFieldRecord() As Double, _
                             ByRef InWave As Wave, _
                             Optional ByVal NumPtsPerFile As Long = 1048000, _
                             Optional ByRef SaveMaxEnvelope As Boolean = True, _
                             Optional ByVal PtsWindowMaxEnvelope As Long = -1, _
                             Optional ByVal isAFRamp As Boolean = True)
                             
    Dim i, j, k As Long
    Dim N, M As Long
    Dim EndPt As Long
    Dim CurTime
    Dim MainLocalFolderPath As String
    Dim FolderName As String
    Dim filename As String
    Dim EnvFileName As String
    Dim PeakField As String
    Dim DCRecordType As String
    Dim DCRecordDesc As String
    Dim ErrorMsg As String
        
    Dim fso As FileSystemObject
    Dim DataStream As TextStream
    Dim EnvStream As TextStream
    
    Dim EnvMax As Double
    
    'Get the current time
    CurTime = Now
    
    'Get the total number of points
    N = UBound(DCFieldRecord)
    
    'Make sure N >=1
    If N < 1 Then
    
        'Msg Box to say there is no data in the DC Field Record
        MsgBox "DC Field Record is empty! No data to save.", , _
               "Whoops!"
               
        Exit Sub
        
    End If
    
    'Determine which main data folder to save to
    If AFSystem = "2G" Then
    
        MainLocalFolderPath = modConfig.TWOG_AFDataLocalDir
        
    Else
    
        MainLocalFolderPath = modConfig.ADWIN_AFDataLocalDir
        
    End If
    
    'Make sure main local folder path ends with a "\"
    If Right(MainLocalFolderPath, 1) <> "\" Then
    
        MainLocalFolderPath = MainLocalFolderPath & "\"
        
    End If
    
    'Generate the File Header string information
    '1st, is this an AF or an IRM DC field record
    If isAFRamp = True Then
                             
        'Get the ramp information from the AF Ramp Up and AF Monitor Wave objects
        InWave.SineFreqMin = WaveForms("AFRAMPUP").SineFreqMin
        
        'Determine which coil was ramped on (which coil is active)
        If AFSystem = "2G" Then
        
            If frmAF_2G.optActiveAxial.value = True Then
            
                'Set the DC Record Type
                DCRecordType = "AF 2G Axial"
                                
            Else
            
                DCRecordType = "AF 2G Axial"
                                
            End If
            
            PeakField = frmAF_2G.txtAmplitude
            
            DCRecordDesc = DCRecordType & " Ramp to " & _
                           PeakField & " " & _
                           modConfig.AFUnits & " @ " & _
                           Trim(str(InWave.SineFreqMin)) & " Hz"
                               
        'AF system is ADWIN
        Else
                             
            If frmADWIN_AF.optCoil(0).value = True Then
            
                'Set the DC Record Type
                DCRecordType = "AF ADWIN Axial"
                
            Else
            
                DCRecordType = "AF ADWIN Transverse"
                                
            End If
            
            PeakField = frmADWIN_AF.txtPeakField
                                        
            DCRecordDesc = DCRecordType & " Ramp to " & _
                           PeakField & " " & _
                           modConfig.AFUnits & " @ " & _
                           Trim(str(InWave.SineFreqMin)) & " Hz"
                                          
        End If
        
        'Check to see if the window for the max envelope has been set
        If PtsWindowMaxEnvelope = -1 Then
        
            PtsWindowMaxEnvelope = CInt(InWave.IORate / InWave.SineFreqMin) + 1
            
        End If
        
    Else
                
        'Get the axis
        If frmIRMARM.optCoil(0).value = True Then
        
            DCRecordType = "IRM Axial"
            
        Else
        
            DCRecordType = "IRM Transverse"
            
        End If
        
        PeakField = frmIRMARM.txtPulseField
        
        DCRecordDesc = DCRecordType & " Pulse to " & _
                       PeakField & " " & _
                       modConfig.AFUnits & " @ " & _
                       Trim(str(InWave.SineFreqMin)) & " Hz"
                       
        'Check to see if the window for the max envelope has been set
        If PtsWindowMaxEnvelope = -1 Then
        
            'Default to 100 pts - the IRM should be entirely positive or entirely
            'negative, therefore, this is just finding the largest point in a 100 point
            'window for easier viewing of the over curve
            PtsWindowMaxEnvelope = 100
            
        End If
                       
    End If
    
    'Pop-up the shutdown msg form with the "DC Record File Saving..." message on it
    Load frmShutdownMsg
    frmShutdownMsg.FontSize = 30
    frmShutdownMsg.lblShutdownMsg.Caption = "DC Record File Saving..."
    frmShutdownMsg.Show
    frmShutdownMsg.ZOrder 0
    
    'Now check to see if the main data folder is still there
    Set fso = New FileSystemObject
       
    'turn on error handling
    On Error GoTo DCRecordSaveError:
        
        ErrorMsg = "checking for main Data Folder:" & vbNewLine & _
                   MainLocalFolderPath
        
        If fso.FolderExists(MainLocalFolderPath) = False Then
        
            ErrorMsg = "creating main 2G Data Folder:" & vbNewLine & _
                       MainLocalFolderPath
        
            fso.CreateFolder MainLocalFolderPath
                
        End If
            
        ErrorMsg = "checking for 'DC Field Records\' sub-folder:" & vbNewLine & _
                   MainLocalFolderPath & "DC Field Records\"
                
        'Now check for the DC Field Records sub-folder
        If fso.FolderExists(MainLocalFolderPath & "DC Field Records\") = False Then
        
            ErrorMsg = "creating 'DC Field Records\' sub-folder:" & vbNewLine & _
                       MainLocalFolderPath & "DC Field Records\"
        
            'Need to create the DC Field Records folder
            fso.CreateFolder MainLocalFolderPath & "DC Field Records\"
        
        End If
        
        'Now need to generate the name for the new DC field record to be saved
        FolderName = DCRecordType & " @ " & Trim(PeakField) & " " & modConfig.AFUnits & _
                     Format(CurTime, "_MM-DD-YYYY_HH-MM-SS") & "\"
                     
        'Now need to create the folder
        ErrorMsg = "creating data folder:" & vbNewLine & _
                   MainLocalFolderPath & "DC Field Records\" & FolderName
                   
        If fso.FolderExists(MainLocalFolderPath & "DC Field Records\" & FolderName) = False Then
        
            fso.CreateFolder (MainLocalFolderPath & "DC Field Records\" & FolderName)
            
        End If
                   
        'Now need to generate the file-name for the first raw data file
        If N <= 1048000 Then
        
            EndPt = N
            
        Else
        
            EndPt = 1048000
            
        End If
        
        filename = "DCField_pts0-" & Trim(str(EndPt)) & _
                   Format(CurTime, "_MM-DD-YYYY_HH-MM-SS") & ".csv"
               
        'Set the current folder to the newly created data folder
        ErrorMsg = "Creating 1st DC Field data file:" & _
                   MainLocalFolderPath & "DC Field Records\" & FolderName & filename
        
        'Now need to create the first raw data file
        Set DataStream = fso.CreateTextFile(MainLocalFolderPath & _
                                     "DC Field Records\" & _
                                     FolderName & _
                                     filename, True)
                                     
        'Now need to write the header
        DataStream.WriteLine DCRecordDesc
        DataStream.WriteLine Format(CurTime, "long date") & "," & Format(CurTime, "long time")
        DataStream.WriteBlankLines (1)
        DataStream.WriteLine "DC Field Raw Data"
        DataStream.WriteLine "Start Pt: 0"
        DataStream.WriteLine "End Pt: " & Trim(str(EndPt))
        DataStream.WriteBlankLines (1)
        DataStream.WriteLine "Pt #, Time, DC Field(" & modConfig.AFUnits & ")"
        
        
        'Do we need to also setup a DC Field Envelope file?
        If SaveMaxEnvelope = True Then
        
            EnvFileName = "DCEnvelope" & _
                          Format(CurTime, "_MM-DD-YYYY_HH-MM-SS") & ".csv"
                          
            'Set the current folder to the newly created data folder
            ErrorMsg = "Creating DC field envelope file:" & _
                       MainLocalFolderPath & "DC Field Records\" & FolderName & filename
            
            'Now need to create the first raw data file
            Set EnvStream = fso.CreateTextFile(MainLocalFolderPath & _
                                         "DC Field Records\" & _
                                         FolderName & _
                                         EnvFileName, True)
        
            'Now need to write the header
            EnvStream.WriteLine DCRecordDesc
            EnvStream.WriteLine Format(CurTime, "long date") & "," & Format(CurTime, "long time")
            EnvStream.WriteBlankLines (1)
            EnvStream.WriteLine "DC Field Envelope Data"
            EnvStream.WriteBlankLines (1)
            EnvStream.WriteLine "Pt #, Time, DC Env. Field(" & modConfig.AFUnits & ")"
        
        End If
        
        'Set j counter & EnvMax to zero
        j = 0
        EnvMax = 0
        
        'No need to loop through and write to file the data points
        For i = 0 To N
            
            ErrorMsg = "writing data point #" & Trim(str(i)) & " to file:" & vbNewLine & _
                        MainLocalFolderPath & "DC Field Records\" & FolderName & filename
                                
            'Write the current line of data
            DataStream.WriteLine Trim(str(i)) & "," & _
                                 Format(i * InWave.TimeStep, "#0.000000") & "," & _
                                 Format(DCFieldRecord(i), "0.00000")
                                 
            'Record the maximum absolute value
            If EnvMax < Abs(DCFieldRecord(i)) Then
            
                EnvMax = Abs(DCFieldRecord(i))
                
            End If
                                 
            'Update j counter
            j = j + 1
            
            'Check to see if j > envelope window
            If j > PtsWindowMaxEnvelope Then
            
                'Reset j = 0
                j = 0
                
                'Save the Envelope max
                ErrorMsg = "writing envelope max data point #" & Trim(str(i)) & " to file:" & vbNewLine & _
                            MainLocalFolderPath & "DC Field Records\" & FolderName & EnvFileName
                                    
                'Write the current line of data
                EnvStream.WriteLine Trim(str(i)) & "," & _
                                     Format(i * InWave.TimeStep, "#0.000000") & "," & _
                                     Format(EnvMax, "0.00000")
                                 
                'Reset Envelope max
                EnvMax = 0
                
            End If
            
            'Check to see if we need to create a new data file to hold the next set of pts
            If i >= EndPt And i < N Then
            
                'Update end pt
                EndPt = EndPt + NumPtsPerFile
            
                'Check to see if the new end-pt is greater than the total number of pts
                If EndPt > N Then EndPt = N
                
                'Create new file
                filename = "DCField_pts" & Trim(str(i + 1)) & "-" & Trim(str(EndPt)) & _
                           Format(CurTime, "_MM-DD-YYYY_HH-MM-SS") & ".csv"
               
                'Set the current folder to the newly created data folder
                ErrorMsg = "Creating next DC Field data file:" & _
                           MainLocalFolderPath & "DC Field Records\" & FolderName & filename
                
                'Now need to create the first raw data file
                Set DataStream = fso.CreateTextFile(MainLocalFolderPath & _
                                             "DC Field Records\" & _
                                             FolderName & _
                                             filename, True)
                                     
                'Now need to write the header
                DataStream.WriteLine DCRecordDesc
                DataStream.WriteLine Format(CurTime, "long date") & "," & Format(CurTime, "long time")
                DataStream.WriteBlankLines (1)
                DataStream.WriteLine "DC Field Raw Data"
                DataStream.WriteLine "Start Pt: " & Trim(str(i + 1))
                DataStream.WriteLine "End Pt: " & Trim(str(EndPt))
                DataStream.WriteBlankLines (1)
                DataStream.WriteLine "Pt #, Time, DC Field(" & modConfig.AFUnits & ")"
                
            End If
            
        Next i
        
    'Turn off error handling
    On Error GoTo 0
    
    'Deallocate the file-system & text stream objects
    DataStream.Close
    EnvStream.Close
    Set DataStream = Nothing
    Set EnvStream = Nothing
    Set fso = Nothing
    
    'Hide the shutdown msg form
    frmShutdownMsg.Hide
    
    Exit Sub
    
DCRecordSaveError:
        
    MsgBox "An error occurred during the DC Field data file save process while " & _
           ErrorMsg & vbNewLine & vbNewLine & _
           "Err #: " & Err.number & vbNewLine & _
           "Err Msg: " & Err.Description, , _
           "File Save Error"
                   
End Sub

Public Sub SaveGaussmeterData(ByVal FilePath As String)

    Dim fso As Scripting.FileSystemObject
    Dim SaveFile As File
    Dim SaveStream As TextStream
    Dim i, N As Long
    Dim lData As gm_store
    Dim lTime As gm_time
    
    On Error GoTo GaussmeterError:
    
        N = UBound(mod908AGaussmeter.DataArray)
    
        If Err.number <> 0 Then
        
            ReDim mod908AGaussmeter.DataArray(1)
            cmdClearData.Enabled = True
            
            N = 1
            
            Exit Sub
            
        End If
    
        Set fso = Nothing
        Set SaveFile = Nothing
        Set SaveStream = Nothing
        
        Set fso = New FileSystemObject
        
        fso.CreateTextFile FilePath, True
        
        If Err.number <> 0 Then
        
            'Bad File name
            Err.Raise Err.number, _
                    "Bad file path for Gaussmeter save file." & vbNewLine & vbNewLine & _
                    "Attempted file path = " & FilePath, , _
                    "Create File Error"
                    
            Exit Sub
            
        End If
            
        Set SaveFile = fso.GetFile(FilePath)
        Set SaveStream = SaveFile.OpenAsTextStream(ForWriting)
                
        SaveStream.WriteLine "Date,Time,Function,Units,Value"
        
        For i = 0 To N - 1
        
            lData = mod908AGaussmeter.DataArray(i)
            lTime = lData.time
                    
                    
            SaveStream.WriteLine Format(Trim(str(lTime.month)) & "/" & _
                                        Trim(str(lTime.day)) & "/20" & _
                                        Trim(str(lTime.year)), _
                                             "long date") & "," & _
                                 Format(Trim(str(lTime.hour)) & ":" & _
                                        Trim(str(lTime.min)) & ":" & _
                                        Trim(str(lTime.sec)), _
                                        "long time") & "," & _
                                 Trim(str(modestr(lData.Mode))) & "," & _
                                 Trim(str(unitsrange(lData.Units, _
                                                     lData.range))) & "," & _
                                 Format(makeactualvalue(lData), _
                                        unitsrangefmt(lData.Units, _
                                                      lData.range))
                                    
        Next i
        
    On Error GoTo 0
        
    Exit Sub
    
GaussmeterError:
                                
End Sub

'---------------------------------------------------------------------------------------------------------'
'---------------------------------------------------------------------------------------------------------'
'
'   Code for controlling the 908A gaussmeter via USB or RS-232 from the Computer
'
'---------------------------------------------------------------------------------------------------------'
'---------------------------------------------------------------------------------------------------------'

Public Sub SetUnits(ByVal UnitsStr As String)

    Dim i As Long
    Dim j As Long
    Dim gmUnits As Byte
    Dim gmRange As Byte
    
    Select Case UnitsStr
    
        Case "G"
        
            gmUnits = 1
            gmRange = 2
        
        Case "kG"
        
            gmUnits = 1
            gmRange = 0
        
        Case "T"
        
            gmUnits = 0
            gmRange = 0
        
        Case "mT"
        
            gmUnits = 0
            gmRange = 2
        
        Case "kA/m"
        
            gmUnits = 2
            gmRange = 1
        
        Case "Oe"
        
            gmUnits = 3
            gmRange = 2
            
        Case "kOe"
        
            gmUnits = 3
            gmRange = 0
            
        Case Else
        
            gmUnits = 1
            gmRange = 2
        
    End Select
    
    'Now use form controls to trigger
    'click events and set the units and range to the desired value
    Me.optUnits(gmUnits).value = True
    Me.optRange(gmRange).value = True

End Sub

Public Function StartDCFieldRecord(ByRef InWave As Wave) As Boolean
                                   
    Dim ReturnStatus As Boolean
                                   
    'Start a background process on the board using the channel specified in the
    'InWave wave object
    With InWave
    
        ReturnStatus = _
            .ManageBackgroundProcess(AIFUNCTION, _
                                     DataArray, _
                                     .WaveName, _
                                     False, _
                                     False, _
                                     1)
                                     
        StartDCFieldRecord = ReturnStatus

    End With
                                   
End Function

Public Function StopDCFieldRecord(ByRef InWave As Wave, _
                                  Optional ByVal SaveData As Boolean = True) As Boolean

    Dim ReturnStatus As Boolean
                                   
    'Stop a background process on the board that's running on
    'the channel specified in the InWave wave object
    With InWave
    
        ReturnStatus = _
            .ManageBackgroundProcess(AIFUNCTION, _
                                     DataArray, _
                                     .WaveName, _
                                     True, _
                                     True, _
                                     1)
                                     
        StopDCFieldRecord = ReturnStatus

    End With
    
    If SaveData = True Then
    
        SaveDCFieldRecord DataArray, InWave
        
    End If

End Function

