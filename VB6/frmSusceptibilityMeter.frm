VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "mscomm32.ocx"
Begin VB.Form frmSusceptibilityMeter 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Susceptibility Meter"
   ClientHeight    =   3210
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   5040
   Icon            =   "frmSusceptibilityMeter.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   5040
   Begin VB.TextBox StatusText 
      Height          =   285
      Left            =   840
      TabIndex        =   10
      Top             =   840
      Width           =   3615
   End
   Begin VB.CommandButton cmdLagTime 
      Caption         =   "Calibrate lag time"
      Height          =   375
      Left            =   2520
      TabIndex        =   8
      Top             =   1320
      Width           =   2055
   End
   Begin VB.CommandButton cmdMeasure 
      Caption         =   "Measure"
      Height          =   375
      Left            =   840
      TabIndex        =   7
      Top             =   1920
      Width           =   1335
   End
   Begin VB.CommandButton cmdZero 
      Caption         =   "Zero"
      Height          =   375
      Left            =   840
      TabIndex        =   6
      Top             =   1320
      Width           =   1335
   End
   Begin VB.TextBox InputText 
      Height          =   288
      Left            =   2520
      TabIndex        =   3
      Top             =   360
      Width           =   1932
   End
   Begin VB.TextBox OutputText 
      Height          =   285
      Left            =   840
      TabIndex        =   2
      Top             =   360
      Width           =   852
   End
   Begin VB.CommandButton ConnectButton 
      Caption         =   "Connect"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   372
      Left            =   3480
      TabIndex        =   0
      Top             =   2520
      Width           =   1212
   End
   Begin MSCommLib.MSComm MSCommSusceptibility 
      Left            =   2280
      Top             =   2520
      _ExtentX        =   794
      _ExtentY        =   794
      _Version        =   393216
      CommPort        =   8
      DTREnable       =   -1  'True
      BaudRate        =   1200
      DataBits        =   7
   End
   Begin VB.Label Label 
      Caption         =   "Status:"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "Output:"
      Height          =   252
      Left            =   120
      TabIndex        =   5
      Top             =   360
      Width           =   612
   End
   Begin VB.Label Label2 
      Caption         =   "Input:"
      Height          =   252
      Left            =   1920
      TabIndex        =   4
      Top             =   360
      Width           =   492
   End
End
Attribute VB_Name = "frmSusceptibilityMeter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit ' enforce variable declaration!
Public TimeForMeasurement As Double

Private Sub cmdClose_Click()
    Me.Hide
End Sub

Private Sub cmdLagTime_Click()
    LagTime
End Sub

Private Sub cmdMeasure_Click()
    Measure
End Sub

Private Sub cmdZero_Click()
    Zero
End Sub

'
'
' start up the Susceptibility driver
'

Public Sub Connect()


    If Not EnableSusceptibility Then Exit Sub
    
    If MSCommSusceptibility.PortOpen = False And Not NOCOMM_MODE Then
        On Error GoTo ErrorHandler  ' Enable error-handling routine.
                        
        Me.StatusText.text = "Connecting: COM-" & Trim(CStr(COMPortSusceptibility)) & _
                             ", " & SusceptibilitySettings
                             
        MSCommSusceptibility.CommPort = COMPortSusceptibility
        MSCommSusceptibility.Settings = SusceptibilitySettings
        MSCommSusceptibility.SThreshold = 1
        MSCommSusceptibility.RThreshold = 0
        MSCommSusceptibility.inputlen = 1
        MSCommSusceptibility.PortOpen = True
        On Error GoTo 0 ' Turn off error trapping.
        If MSCommSusceptibility.PortOpen = True Then
            ConnectButton.Caption = "Disconnect"
            '
            ' disable the other connection buttons here until com is free
            '
        End If
    End If
    
    Me.StatusText.text = vbNullString
    
Exit Sub        ' Exit to avoid handler.
ErrorHandler:   ' Error-handling routine.

    Me.StatusText.text = "Connection Error: " & Trim(CStr(Err.number))

    Select Case Err.number  ' Evaluate error number.
        Case 8002
            MsgBox "Invalid Port Number"
        Case 8005
            MsgBox "Port already open" + Chr(13) + "(Already is use?)"
        Case 8010
            MsgBox "The hardware is not available (locked by another device)"
        Case 8012
            MsgBox "The device is not open"
        Case 8013
            MsgBox "The device is already open"
        Case Else
            MsgBox "Unknown error trying to Connect Comm Port"
    End Select

    
End Sub

Private Sub ConnectButton_Click()
    If MSCommSusceptibility.PortOpen = False Then
        Connect
    Else
        Disconnect
    End If
End Sub

Public Sub Disconnect()
    
    If Not EnableSusceptibility Then Exit Sub

    If MSCommSusceptibility.PortOpen = True Then
        MSCommSusceptibility.PortOpen = False
        ConnectButton.Caption = "Connect"
    End If
End Sub

Private Sub Form_Load()

    '(March 11, 2011 - I Hilburn)
    'Update all the form's icons with the program Ico file
    SetFormIcon Me

End Sub

Private Sub Form_Unload(Cancel As Integer)
    If MSCommSusceptibility.PortOpen = True Then
        MSCommSusceptibility.PortOpen = False
    End If
End Sub

Private Function GetResponse() As String
    Dim Delay As Double
    Dim inputchar As String
    
    'Check for susceptibility modules enabled status
    If EnableSusceptibility = False Then Exit Function

    If COMPortSusceptibility < 1 Then Exit Function

    Delay = Timer   ' Set delaystart time.
    inputchar = vbNullString
    
    Me.StatusText.text = "Getting Response:"
       
    
    Do While Right$(inputchar, 1) <> vbCr And Not NOCOMM_MODE
        DoEvents
        
        PauseTill_NoEvents timeGetTime() + 100
        
        If MSCommSusceptibility.InBufferCount > 0 Then
            inputchar = inputchar + MSCommSusceptibility.Input
        End If
        If Timer < Delay Then Delay = Delay - 86400
        
        Me.StatusText.text = "Getting Response: " & Format$(Timer - Delay, "#0.0")
        
        If Timer - Delay > TimeForMeasurement + 1 Then
            Exit Do
            MsgBox "Timeout sending command to Susceptibility"
        End If
    Loop
    InputText = inputchar
    GetResponse = inputchar
    If DEBUG_MODE Then frmDebug.Msg "COM " & Str$(MSCommSusceptibility.CommPort) & "in: " & inputchar
    
    Me.StatusText.text = vbNullString

End Function

Public Function LagTime() As Double
    Dim waitingTime As Double
    
    If EnableSusceptibility = False Then Exit Function
    
    Me.InputText.text = "Calculating Lag Time ..."
        
    TimeForMeasurement = 30
    
    waitingTime = Timer
    Measure (False)
    If Timer < waitingTime Then waitingTime = waitingTime - 86400
    TimeForMeasurement = Timer - waitingTime
    LagTime = TimeForMeasurement
        
End Function

Public Function Measure(Optional ByVal checkCalibration = True) As Double
    Dim a As String
    
    If EnableSusceptibility = False Then Exit Function
    
    If checkCalibration And TimeForMeasurement <= 0 Then
        LagTime
        
        Me.StatusText.text = "LagTime: " & Trim(CStr(TimeForMeasurement)) & " secs"
        PauseTill_NoEvents timeGetTime() + 500
    Else
        Me.StatusText.text = "Measuring ..."
    End If
    SendCommand ("M")
    a = GetResponse
    If a = vbNullString Then Measure = -1 Else Measure = val(a) * SusceptibilityScaleFactor
    InputText = Measure
    Me.StatusText.text = vbNullString
End Function

Private Sub SendCommand(outstring As String)

    If Not EnableSusceptibility Then Exit Sub
    
    If Not MSCommSusceptibility.PortOpen Then Connect
    
    If MSCommSusceptibility.PortOpen = True Then
        MSCommSusceptibility.RTSEnable = False
        MSCommSusceptibility.OutBufferCount = 0
        MSCommSusceptibility.InBufferCount = 0
        MSCommSusceptibility.Output = outstring + vbCrLf
        OutputText = outstring
        If DEBUG_MODE Then frmDebug.Msg "COM " & Str$(MSCommSusceptibility.CommPort) & "out: " & outstring

    Else
        If Not NOCOMM_MODE Then MsgBox "Susceptibility Comm Port Not Open"
    End If
End Sub

Public Sub Zero()
    
    'Prevent zeroing if susceptibility module is not enabled
    If EnableSusceptibility = False Then Exit Sub

    SendCommand ("Z")
    GetResponse
End Sub

