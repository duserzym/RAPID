VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "mscomm32.ocx"
Begin VB.Form frmSQUID 
   Caption         =   "SQUID"
   ClientHeight    =   3705
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5325
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3705
   ScaleWidth      =   5325
   Begin MSCommLib.MSComm MSCommSquid 
      Left            =   4800
      Top             =   840
      _ExtentX        =   794
      _ExtentY        =   794
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.CommandButton ConnectButton 
      Caption         =   "Connect"
      Height          =   375
      Left            =   1800
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   372
      Left            =   3840
      TabIndex        =   1
      Top             =   240
      Width           =   1092
   End
   Begin VB.TextBox OutputText 
      Height          =   285
      Left            =   1200
      TabIndex        =   2
      Top             =   960
      Width           =   852
   End
   Begin VB.TextBox InputText 
      Height          =   288
      Left            =   2880
      TabIndex        =   3
      Top             =   960
      Width           =   852
   End
   Begin VB.OptionButton AxisOptionX 
      Caption         =   "X"
      Height          =   252
      Left            =   360
      TabIndex        =   4
      Top             =   1440
      Width           =   372
   End
   Begin VB.OptionButton AxisOptionY 
      Caption         =   "Y"
      Height          =   252
      Left            =   360
      TabIndex        =   5
      Top             =   1680
      Width           =   372
   End
   Begin VB.OptionButton AxisOptionZ 
      Caption         =   "Z"
      Height          =   252
      Left            =   360
      TabIndex        =   6
      Top             =   1920
      Width           =   372
   End
   Begin VB.OptionButton AxisOptionA 
      Caption         =   "A"
      Height          =   252
      Left            =   360
      TabIndex        =   7
      Top             =   2160
      Width           =   372
   End
   Begin VB.CommandButton ConfCmdButton 
      Caption         =   "CR1"
      Height          =   252
      Index           =   0
      Left            =   960
      TabIndex        =   8
      Top             =   1440
      Width           =   612
   End
   Begin VB.CommandButton ConfCmdButton 
      Caption         =   "CLC"
      Height          =   252
      Index           =   1
      Left            =   960
      TabIndex        =   9
      Top             =   1680
      Width           =   612
   End
   Begin VB.CommandButton ConfCmdButton 
      Caption         =   "CSE"
      Height          =   252
      Index           =   2
      Left            =   960
      TabIndex        =   10
      Top             =   1920
      Width           =   612
   End
   Begin VB.CommandButton ConfCmdButton 
      Caption         =   "CF1"
      Height          =   252
      Index           =   3
      Left            =   960
      TabIndex        =   11
      Top             =   2160
      Width           =   612
   End
   Begin VB.CommandButton ConfCmdButton 
      Caption         =   "CLP"
      Height          =   252
      Index           =   4
      Left            =   960
      TabIndex        =   12
      Top             =   2400
      Width           =   612
   End
   Begin VB.CommandButton ReadCountButton 
      Caption         =   "Read Count"
      Height          =   372
      Left            =   1800
      TabIndex        =   13
      Top             =   1440
      Width           =   1332
   End
   Begin VB.CommandButton ReadDataButton 
      Caption         =   "Read Data"
      Height          =   372
      Left            =   1800
      TabIndex        =   14
      Top             =   1920
      Width           =   1332
   End
   Begin VB.CommandButton ReadRangeButton 
      Caption         =   "Read Range"
      Height          =   372
      Left            =   1800
      TabIndex        =   15
      Top             =   2400
      Width           =   1332
   End
   Begin VB.CommandButton ResetCountButton 
      Caption         =   "Reset Count"
      Height          =   372
      Left            =   3600
      TabIndex        =   16
      Top             =   1440
      Width           =   1092
   End
   Begin VB.Frame ChangeRangeFrame 
      Caption         =   "Change Range"
      Height          =   972
      Left            =   3360
      TabIndex        =   17
      Top             =   1920
      Width           =   1692
      Begin VB.CommandButton ChangeRangeF 
         Caption         =   "F"
         Height          =   252
         Index           =   0
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   252
      End
      Begin VB.CommandButton ChangeRangeT 
         Caption         =   "T"
         Height          =   252
         Index           =   2
         Left            =   600
         TabIndex        =   19
         Top             =   240
         Width           =   252
      End
      Begin VB.CommandButton ChangeRangeE 
         Caption         =   "E"
         Default         =   -1  'True
         Height          =   252
         Index           =   4
         Left            =   1080
         TabIndex        =   20
         Top             =   240
         Width           =   252
      End
      Begin VB.CommandButton ChangeRange1 
         Caption         =   "1"
         Height          =   252
         Index           =   1
         Left            =   120
         TabIndex        =   21
         Top             =   600
         Width           =   252
      End
      Begin VB.CommandButton ChangeRangeH 
         Caption         =   "H"
         Height          =   252
         Index           =   3
         Left            =   600
         TabIndex        =   22
         Top             =   600
         Width           =   252
      End
   End
   Begin VB.CommandButton ConfigureButton 
      Caption         =   "Configure SQUID"
      Height          =   495
      Left            =   240
      TabIndex        =   23
      Top             =   2760
      Width           =   1332
   End
   Begin VB.CommandButton ReadButton 
      Caption         =   "Read"
      Height          =   372
      Left            =   3720
      TabIndex        =   24
      Top             =   3120
      Width           =   732
   End
   Begin VB.Label Label3 
      Caption         =   "Output:"
      Height          =   252
      Left            =   480
      TabIndex        =   25
      Top             =   960
      Width           =   612
   End
   Begin VB.Label Label2 
      Caption         =   "Input:"
      Height          =   252
      Left            =   2280
      TabIndex        =   26
      Top             =   960
      Width           =   492
   End
End
Attribute VB_Name = "frmSQUID"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' (March 2008 L Carporzen) Const ReadDelay = 1
Dim ActiveAxis As String

Private Sub AxisOptionA_Click()
    ActiveAxis = "A"
End Sub

Private Sub AxisOptionX_Click()
    ActiveAxis = "X"
End Sub

Private Sub AxisOptionY_Click()
    ActiveAxis = "Y"
End Sub

Private Sub AxisOptionZ_Click()
    ActiveAxis = "Z"
End Sub

Public Function Calibrate(axis As String, val As Double) _
    As Double
    ' This function takes a string representing the axis
    ' being measured, a value measured from the axis,
    ' and returns a calibrated value, using constants
    ' previously read from a file.
    Select Case axis
        Case "X":
            Calibrate = val * XCal
        Case "Y":
            Calibrate = val * YCal
        Case "Z":
            Calibrate = val * ZCal
        Case Else:
            MsgBox ("Error occured in frmSQUID.Calibrate:" & _
                vbCr & "Invalid axis argument given to the function.")
    End Select
End Function

Public Sub ChangeRange(axis As String, ChangeRangeSelected As String)
    Select Case ChangeRangeSelected
        Case "F":
            ' Set the system up for the extended range flux
            ' counting stuff. First, enable (turn ON) the
            ' fast-slew
            SendCommand (axis + "CSE")
            SendCommand (axis + "CR1") ' All control rate 1
        Case "1", "T", "H", "E":
            SendCommand (axis + "CR" + ChangeRangeSelected)
        Case Else:
            ' This should never happen
            MsgBox "Error occurred in ChangeRangeButton: " & _
                "invalid range specifed: " + ChangeRangeSelected, vbOKOnly, "ERROR!"
    End Select
End Sub

Private Sub ChangeRange1_Click(Index As Integer)
    ChangeRange ActiveAxis, "1"
End Sub

Private Sub ChangeRangeE_Click(Index As Integer)
    ChangeRange ActiveAxis, "E"
End Sub

Private Sub ChangeRangeF_Click(Index As Integer)
    ChangeRange ActiveAxis, "F"
End Sub

Private Sub ChangeRangeH_Click(Index As Integer)
    ChangeRange ActiveAxis, "H"
End Sub

Private Sub ChangeRangeT_Click(Index As Integer)
    ChangeRange ActiveAxis, "T"
End Sub

Public Sub CLP(axis As String)
    SendCommand (axis + "CLP")
End Sub

Private Sub cmdClose_Click()
    Me.Hide
End Sub

Private Sub ConfCmdButton_Click(Index As Integer)
    SendCommand (ActiveAxis + ConfCmdButton(Index).Caption)
End Sub

Public Sub Configure(axis As String)
    Dim delayed As Double
    delayed = 0.05
    SendCommand (axis + "CR1")
    DelayTime delayed
    SendCommand (axis + "CLC")
    DelayTime delayed
    SendCommand (axis + "CSE")
    DelayTime delayed
    SendCommand (axis + "CF1")
    DelayTime delayed
    SendCommand (axis + "CLP")
End Sub

Private Sub ConfigureButton_Click()
    Configure (ActiveAxis)
End Sub

Private Sub Connect()
        If MSCommSquid.PortOpen = False And Not NOCOMM_MODE Then
            On Error GoTo ErrorHandler  ' Enable error-handling routine.
            MSCommSquid.CommPort = COMPortSquids
            MSCommSquid.Settings = "1200,N,8,1"
            MSCommSquid.SThreshold = 1
            MSCommSquid.RThreshold = 0
            MSCommSquid.inputlen = 1
            MSCommSquid.PortOpen = True
            On Error GoTo 0 ' Turn off error trapping.
            If MSCommSquid.PortOpen = True Then
                ConnectButton.Caption = "Disconnect"
                ' disable the other connection buttons here until com is free
            Else
                MSCommSquid.PortOpen = False
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
    
    'Prompt the user to see if they would like to turn on NOCOMM mode
    Prompt_NOCOMM
    
End Sub

Private Sub ConnectButton_Click()
    '    SetPorts
    If MSCommSquid.PortOpen = False Then
        Connect
    Else
        Disconnect
    End If
End Sub

Public Sub Disconnect()
    If MSCommSquid.PortOpen = True Then
        MSCommSquid.PortOpen = False
    End If
    ConnectButton.Caption = "Connect"
End Sub

Private Sub Form_Load()

    '(March 11, 2011 - I Hilburn)
    'Update all the form's icons with the program Ico file
    SetFormIcon Me

    AxisOptionA_Click
    If MSCommSquid.PortOpen = False Then Connect
    
End Sub

Private Sub Form_Resize()
    Me.Height = 4110
    Me.Width = 5445
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If MSCommSquid.PortOpen = True Then
        MSCommSquid.PortOpen = False
    End If
End Sub

Public Function getData(Optional ByVal AlreadyLatched As Boolean = False) As Cartesian3D
    ' This function returns the data gathered from the three axes
    ' in the magnetometer, reading the squid boxes.
    Dim X As Double, Y As Double, Z As Double
    Set getData = New Cartesian3D
    If Not AlreadyLatched Then latchVal "A", True   ' latch values with delay for a short time
    ' Gather data from squid boxes
    X = getVal("X", True)
    Y = getVal("Y", True)
    Z = getVal("Z", True)
    ' Calibrate the data just received
    With getData
            .X = Calibrate("X", X)
            .Y = Calibrate("Y", Y)
            .Z = Calibrate("Z", Z)
    End With
End Function

Private Function GetResponse() As String
    Dim Delay As Double
    Dim inputchar As String
    If MSCommSquid.PortOpen = True Then
        Delay = Timer   ' Set delaystart time.
        inputchar = vbNullString
        Do While Right$(inputchar, 1) <> vbCr
            DoEvents
            If MSCommSquid.InBufferCount > 0 Then
                inputchar = inputchar + MSCommSquid.Input
            End If
            If Timer < Delay Then Delay = Delay - 86400
            If Timer - Delay > 1 Then
                'MsgBox "Timeout reading from SQUID"
                inputchar = vbCr
                Exit Do
            End If
        Loop
        InputText = Left$(inputchar, Len(inputchar) - 1)
    Else
        If Not NOCOMM_MODE Then MsgBox "SQUID Comm Port Not Open"
    End If
    GetResponse = inputchar
End Function

Public Function getVal(ByVal dir As String, Optional ByVal AlreadyLatched As Boolean = False) As Double
    ' This function automatically swtiches the line accessed by
    ' COM 2 to the 2G Squid boxes, and reads in the value of the
    ' axis described by 'dir'.  If this is the first zero
    ' measurement, then isFirstZero should be true
    Dim rangeStr As String
    Dim rangeval As Double
    Dim Count, data As Double
    ' Look in Paleomag.GETVAL:
    If Not AlreadyLatched Then latchVal dir, False
    Count = SendCount(dir)
    data = SendData(dir)          ' Read data
    ' Check to make sure we're on the right scale ...
    modeFluxCount = True  ' !!! Flux counting mode not implemented
    If Not modeFluxCount Then
        ' Ask for range on Squid boxes
        ' Read range
        rangeStr = Mid(frmSQUID.ReadRange(dir), 2, 1)  ' Response like "R1"
        Select Case rangeStr
            Case "1"
                rangeval = 1
            Case "T"
                rangeval = 10
            Case "H"
                rangeval = 100
            Case "E"
                rangeval = 1000
            Case Else
              MsgBox "Error occurred in Measure_getVal:" & _
                    vbCr & "Invalid range read from 2G Squid boxes: " + rangeStr
        End Select
    Else
        ' In flux counting mode, don't need to ask for range
        rangeval = 1
    End If
    getVal = -val(data) - val(Count) * rangeval
End Function

Public Sub LatchCount(axis As String)
    SendCommand (axis + "LC")
    DelayTime 0.1
End Sub

Public Sub LatchData(axis As String)
    SendCommand (axis + "LD")
    DelayTime 0.12
End Sub

Public Sub latchVal(Optional ByVal dir As String = "A", Optional ByVal withDelay As Boolean = False)
    'If Prog_halted Then Exit Sub ' (September 2007 L Carporzen) New version of the Halt button
    If withDelay Then
        frmProgram.StatusBar "Settling...", 3
        DelayTime ReadDelay
        frmProgram.StatusBar vbNullString, 3
    End If
    Select Case dir
        Case "A"
            LatchCount "A"
            LatchData "A"
        Case "X"
            LatchCount "X"
            LatchData "X"
        Case "Y"
            LatchCount "Y"
            LatchData "Y"
        Case "Z"
            LatchCount "Z"
            LatchData "Z"
    End Select
End Sub

Private Sub ReadButton_Click()
    GetResponse
End Sub

Public Function ReadCount(axis As String) As Double
    SendCommand (axis & "LC")
    DelayTime 0.12
    SendCommand (axis & "SC")
    ReadCount = val(GetResponse)
End Function

Private Sub ReadCountButton_Click()
    ReadCount (ActiveAxis)
End Sub

Public Function ReadData(axis As String) As Double
    SendCommand (axis + "LD")
    DelayTime 0.12
    SendCommand (axis + "SD")
    ReadData = val(GetResponse)
End Function

Private Sub ReadDataButton_Click()
    ReadData (ActiveAxis)
End Sub

Public Function ReadRange(axis As String) As String
    ' if Axis = "A", this is a clear error!
    If axis = "A" Then MsgBox "Error occurred in ReadRange:" & _
                    vbCr & "2G Squid boxes cannot talk at once!"
    SendCommand (axis + "SSR")
    ReadRange = GetResponse
End Function

Private Sub ReadRangeButton_Click()
    ReadRange (ActiveAxis)
End Sub

Public Sub ResetCount(axis As String)
    SendCommand (axis + "RC")
End Sub

Private Sub ResetCountButton_Click()
    ResetCount (ActiveAxis)
End Sub

Public Sub SelectAxis(axis As String)
    Select Case axis
        Case "X":
            AxisOptionX_Click
        Case "Y":
            AxisOptionY_Click
        Case "Z":
            AxisOptionZ_Click
        Case "A":
            AxisOptionA_Click
    End Select
End Sub

Private Sub SendCommand(outstring As String)
    Dim i As Integer
    If MSCommSquid.PortOpen = True Then
        MSCommSquid.RTSEnable = True
        MSCommSquid.OutBufferCount = 0
        If Left$(outstring, 1) = "D" Then
            ' Write the string to the port slowly
            MSCommSquid.Output = Chr$(13)
            DelayTime (0.15)
            For i = 1 To Len(outstring)
                MSCommSquid.Output = Mid$(outstring, i, 1) + Chr(13)
                DelayTime (0.15)
            Next i
        Else
            MSCommSquid.Output = Chr$(13) + outstring + Chr(13)
        End If
        OutputText = outstring
    Else
        If Not NOCOMM_MODE Then MsgBox "SQUID Comm Port Not Open"
    End If
End Sub

Public Function SendCount(axis As String)
    SendCommand (axis & "SC")
    SendCount = val(GetResponse)
End Function

Public Function SendData(axis As String)
    SendCommand (axis + "SD")
    SendData = val(GetResponse)
End Function

Public Function SquidConnected() As Boolean
    SquidConnected = MSCommSquid.PortOpen
End Function

