VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmVacuum 
   Caption         =   "Vacuum / Cooling"
   ClientHeight    =   3780
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4560
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3780
   ScaleWidth      =   4560
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   495
      Left            =   240
      TabIndex        =   15
      Top             =   2640
      Width           =   4095
      Begin VB.OptionButton DegausserCoolerOff 
         Caption         =   "Off"
         Height          =   372
         Left            =   2520
         TabIndex        =   17
         Top             =   0
         Width           =   732
      End
      Begin VB.OptionButton DegausserCoolerOn 
         Caption         =   "On"
         Height          =   372
         Left            =   1680
         TabIndex        =   16
         Top             =   0
         Width           =   732
      End
      Begin VB.Label Label1 
         Caption         =   "Degausser Cooler:"
         Height          =   255
         Left            =   0
         TabIndex        =   18
         Top             =   120
         Width           =   1455
      End
   End
   Begin MSCommLib.MSComm MSCommVacuum 
      Left            =   3960
      Top             =   840
      _ExtentX        =   794
      _ExtentY        =   794
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.TextBox OutputText 
      Height          =   285
      Left            =   840
      TabIndex        =   0
      Top             =   240
      Width           =   852
   End
   Begin VB.TextBox InputText 
      Height          =   288
      Left            =   2520
      TabIndex        =   1
      Top             =   240
      Width           =   1932
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   4095
      Begin VB.OptionButton VacuumConnectOn 
         Caption         =   "On"
         Height          =   372
         Left            =   1680
         TabIndex        =   3
         Top             =   0
         Width           =   732
      End
      Begin VB.OptionButton VacuumConnectOff 
         Caption         =   "Off"
         Height          =   372
         Left            =   2520
         TabIndex        =   4
         Top             =   0
         Width           =   732
      End
      Begin VB.Label Label4 
         Caption         =   "Vacuum Connect:"
         Height          =   255
         Left            =   0
         TabIndex        =   5
         Top             =   120
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   1320
      Width           =   4095
      Begin VB.OptionButton VacuumMotorOn 
         Caption         =   "On"
         Height          =   372
         Left            =   1680
         TabIndex        =   7
         Top             =   0
         Width           =   732
      End
      Begin VB.OptionButton VacuumMotorOff 
         Caption         =   "Off"
         Height          =   372
         Left            =   2520
         TabIndex        =   8
         Top             =   0
         Width           =   732
      End
      Begin VB.Label Label5 
         Caption         =   "Vacuum Motor:"
         Height          =   252
         Left            =   0
         TabIndex        =   9
         Top             =   120
         Width           =   1452
      End
   End
   Begin VB.CommandButton ConnectButton 
      Caption         =   "Connect"
      Height          =   375
      Left            =   480
      TabIndex        =   10
      Top             =   1800
      Width           =   1335
   End
   Begin VB.CommandButton ResetButton 
      Caption         =   "Reset"
      Height          =   372
      Left            =   2640
      TabIndex        =   11
      Top             =   1800
      Width           =   1212
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   372
      Left            =   3120
      TabIndex        =   12
      Top             =   3240
      Width           =   1212
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   4200
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Label Label3 
      Caption         =   "Output:"
      Height          =   252
      Left            =   120
      TabIndex        =   13
      Top             =   240
      Width           =   612
   End
   Begin VB.Label Label2 
      Caption         =   "Input:"
      Height          =   252
      Left            =   1920
      TabIndex        =   14
      Top             =   240
      Width           =   492
   End
End
Attribute VB_Name = "frmVacuum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Vacuum Controlling system for Chris Baumgarter's Vacuum Box
' modeled after the modAF_2G by JLK, Oct. 23, 2003
' rewritten by REK, 28 Oct.,
' updated 28 Mar 2004, and for new vacuum commands Sept. 2007
Option Explicit ' enforce variable declaration!
Dim ValveConnected As Boolean
Dim MotorPowered As Boolean
'Dim i As Integer
'Const COMPort = COMPortVacuum
'Const Settings = "9600,n,8,1"

Private Sub cmdClose_Click()
    Me.Hide
End Sub

' start up the Vacuum driver
Public Sub Connect()
    
    'Check to see if the Vacuum module is enabled
    If EnableVacuum = False Then Exit Sub
    
    If MSCommVacuum.PortOpen = False And Not NOCOMM_MODE And COMPortVacuum > 0 Then
        On Error GoTo ErrorHandler  ' Enable error-handling routine.
        MSCommVacuum.CommPort = COMPortVacuum
        MSCommVacuum.Settings = "9600,n,8,1" ' Settings
        MSCommVacuum.SThreshold = 1
        MSCommVacuum.RThreshold = 0
        MSCommVacuum.inputlen = 1
        MSCommVacuum.PortOpen = True
        On Error GoTo 0 ' Turn off error trapping.
        If MSCommVacuum.PortOpen = True Then
            ConnectButton.Caption = "Disconnect"
            ' disable the other connection buttons here until com is free
        End If
    End If
Exit Sub        ' Exit to avoid handler.
ErrorHandler:   ' Error-handling routine.
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
    
    'Prompt the user if they want to turn on NOCOMM_MODE
    Prompt_NOCOMM
    
End Sub

Private Sub ConnectButton_Click()
    If MSCommVacuum.PortOpen = False Then
        Connect
    Else
        Disconnect
    End If
End Sub

Public Sub DegausserCooler(Optional ByVal switch As Boolean = False)
    
    'If NOCOMM_MODE, then exit sub
    If NOCOMM_MODE = True Then Exit Sub
    
    'If DegausserCooler not enabled, exit sub
    If EnableDegausserCooler = False Then Exit Sub
        
    If switch Then
        frmDAQ_Comm.DoDAQIO DegausserToggle, , True
    Else
        frmDAQ_Comm.DoDAQIO DegausserToggle, , False
    End If
End Sub

Private Sub DegausserCoolerOff_Click()
   DegausserCooler False
End Sub

Private Sub DegausserCoolerOn_Click()
   DegausserCooler True
End Sub

Public Sub Disconnect()
    If MSCommVacuum.PortOpen = True Then
        MSCommVacuum.PortOpen = False
        ConnectButton.Caption = "Connect"
    End If
End Sub

Private Sub Form_Load()
    
    '(March 11, 2011 - I Hilburn)
    'Update all the form's icons with the program Ico file
    SetFormIcon Me
    
    If Not NOCOMM_MODE And COMPortVacuum > 0 And EnableVacuum Then
        If MSCommVacuum.PortOpen = False Then Connect
        Reset
        ValveConnect False
        MotorPower False
        DegausserCooler False
        DelayTime (0.2)
        ValveConnect False
        MotorPower False
        DegausserCooler False
    End If
    
End Sub

Private Sub Form_Resize()
    Me.Height = 4515
    Me.Width = 4725
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If MSCommVacuum.PortOpen = True Then
        MSCommVacuum.PortOpen = False
    End If
End Sub

Private Sub GetResponse()
    Dim Delay As Double
    Dim inputchar As String
    Delay = Timer   ' Set delaystart time.
    inputchar = vbNullString
    Do While Right$(inputchar, 1) <> vbCr And Not NOCOMM_MODE And COMPortVacuum > 0
        DoEvents
        If MSCommVacuum.InBufferCount > 0 Then
            inputchar = inputchar + MSCommVacuum.Input
        End If
        If Timer < Delay Then Delay = Delay - 86400
        If Timer - Delay > 0.3 Then
            inputchar = Chr(vbKeyReturn)
            Exit Do
            'MsgBox "Timeout sending command to vacuum"
        End If
    Loop
    InputText = inputchar
    If DEBUG_MODE Then frmDebug.Msg "COM " & Str$(MSCommVacuum.CommPort) & " in: " & inputchar
End Sub

Public Sub MotorPower(switch As Boolean)
    
    'If NOCOMM_MODE, then exit sub
    If NOCOMM_MODE = True Then Exit Sub
    
    'If vacuum module not enabled, exit sub
    If EnableVacuum = False Then Exit Sub
        
    If switch Then
        SendCommand ("E")
        SendCommand ("10MFF")
        GetResponse
        VacuumMotorOn = True
        MotorPowered = True
        ' now the digital lines
        '(July 2010 - I Hilburn) Replaces old frmMCC with frmDAQ_Comm using the Channel object variables
        frmDAQ_Comm.DoDAQIO MotorToggle, , True
 '       txtMotorToggle = "1"
    Else
        SendCommand ("D")
        SendCommand ("10M00")
        GetResponse
        VacuumMotorOff = True
        MotorPowered = False
        ' now the digital lines
        '(July 2010 - I Hilburn) Replaces old frmMCC with frmDAQ_Comm using the Channel object variables
        frmDAQ_Comm.DoDAQIO MotorToggle, , False
'        txtMotorToggle = "0"
    End If
End Sub

Public Sub Reset()

    'Check for nocomm and vacuum module enable state
    If NOCOMM_MODE = True Or _
       EnableVacuum = False Or _
       modConfig.DoVacuumReset = False Then Exit Sub

   SendCommand ("10R00")
   DelayTime 0.2
   SendCommand ("10TFF")
   GetResponse
End Sub

Private Sub ResetButton_Click()
    Reset
    'frmSendMail.MailNotification "AF too hot", "Testing", CodeYellow (For testing SendMail)
End Sub

Private Sub SendCommand(outstring As String)
        
    'Check to see if the Vacuum module is enabled
    If EnableVacuum = False Then Exit Sub
    
    'check for NOCOMM_MODE
    If NOCOMM_MODE = True Then Exit Sub
    
    If MSCommVacuum.PortOpen = True Then
        MSCommVacuum.RTSEnable = True
        MSCommVacuum.OutBufferCount = 0
        MSCommVacuum.InBufferCount = 0
        MSCommVacuum.Output = vbCr
        DelayTime 0.1
        MSCommVacuum.Output = outstring
        DelayTime 0.1
        MSCommVacuum.Output = vbCr
        DelayTime 0.1
        OutputText = outstring
        If DEBUG_MODE Then frmDebug.Msg "COM " & Str$(MSCommVacuum.CommPort) & " out: " & outstring
    Else
        If Not NOCOMM_MODE And COMPortVacuum > 0 Then MsgBox "Vacuum Comm Port Not Open"
    End If
End Sub

Public Function VacuumActive() As Boolean
    VacuumActive = ValveConnected And MotorPowered
End Function

Private Sub VacuumConnectOff_Click()
   ValveConnect False
End Sub

Private Sub VacuumConnectOn_Click()
   ValveConnect True
End Sub

Private Sub VacuumMotorOff_Click()
   MotorPower False
End Sub

Private Sub VacuumMotorOn_Click()
   MotorPower True
End Sub

Public Sub ValveConnect(switch As Boolean)
    ' the O, C, E, and D comands were added by JLK Sept. 2007 to enable
    ' Chris Baumgarter's new vacuum boxes to work. The old boxes should ignore
    ' the extra commands. May '08 added new commands to toggle the MC Digital IO lines
        
    'Check for NOCOMM_MODE
    If NOCOMM_MODE = True Then Exit Sub
        
    'Check for whether the vacuum module is enabled
    If EnableVacuum = False Then Exit Sub
        
    If switch Then
        SendCommand ("O")
        SendCommand ("10VFF")
        GetResponse
        VacuumConnectOn = True
        ValveConnected = True
        ' now the digital lines
        '(July 2010 - I Hilburn) Replaces old frmMCC with frmDAQ_Comm using the Channel object variables
        frmDAQ_Comm.DoDAQIO VacuumToggleA, , True
 '       txtVacuumToggleA = "1"
    Else
        SendCommand ("C")
        SendCommand ("10V00")
        GetResponse
        VacuumConnectOff = True
        ValveConnected = False
        ' now the digital lines
        '(July 2010 - I Hilburn) Replaces old frmMCC with frmDAQ_Comm using the Channel object variables
        frmDAQ_Comm.DoDAQIO VacuumToggleA, , False
 '       txtVacuumToggleA = "0"
    End If
End Sub

