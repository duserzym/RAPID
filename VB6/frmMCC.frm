VERSION 5.00
Begin VB.Form frmMCC 
   Caption         =   "MCC Interface"
   ClientHeight    =   3855
   ClientLeft      =   135
   ClientTop       =   420
   ClientWidth     =   2985
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3855
   ScaleWidth      =   2985
   Begin VB.CommandButton cmdDigitalOutput 
      Caption         =   "Digital Output"
      Height          =   372
      Left            =   1560
      TabIndex        =   10
      Top             =   2400
      Width           =   1212
   End
   Begin VB.CommandButton cmdDigitalInput 
      Caption         =   "Digital Input"
      Height          =   372
      Left            =   1560
      TabIndex        =   9
      Top             =   1920
      Width           =   1212
   End
   Begin VB.CommandButton cmdAnalogOutput 
      Caption         =   "Analog Output"
      Height          =   372
      Left            =   120
      TabIndex        =   8
      Top             =   2400
      Width           =   1212
   End
   Begin VB.CommandButton cmdAnalogInput 
      Caption         =   "Analog Input"
      Height          =   372
      Left            =   120
      TabIndex        =   7
      Top             =   1920
      Width           =   1212
   End
   Begin VB.TextBox txtEng 
      Height          =   285
      Left            =   1560
      TabIndex        =   6
      Top             =   1200
      Width           =   735
   End
   Begin VB.TextBox txtRaw 
      Height          =   285
      Left            =   1560
      TabIndex        =   5
      Top             =   720
      Width           =   735
   End
   Begin VB.ComboBox cmbChan 
      Height          =   315
      Left            =   1560
      TabIndex        =   4
      Text            =   "Combo1"
      Top             =   240
      Width           =   735
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   372
      Left            =   840
      TabIndex        =   0
      Top             =   2880
      Width           =   1212
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Outputs read English value."
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   3480
      Width           =   2895
   End
   Begin VB.Label Label3 
      Caption         =   "English Value:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Raw Value:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Channel/Pin:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin VB.Menu mnuLines 
      Caption         =   "Lines"
      NegotiatePosition=   3  'Right
      Visible         =   0   'False
      Begin VB.Menu mnuDigiCheck 
         Caption         =   "DigiCheck"
      End
      Begin VB.Menu mnuClearBit 
         Caption         =   "Clear Bit"
      End
      Begin VB.Menu mnuSetBit 
         Caption         =   "Set Bit"
      End
   End
End
Attribute VB_Name = "frmMCC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' MCC Controller
'
' This is the driver for the MCC I/O Interface
'

Dim DIOConfig(7) As Integer

Option Explicit

Const BoardNum% = 0              ' Board number

Private Sub cmdAnalogInput_Click()
    AnalogInput (val(cmbChan))
End Sub

Private Sub cmdAnalogOutput_Click()
    AnalogOutput val(cmbChan), val(txtEng)
End Sub

Private Sub cmdDigitalInput_Click()
    DigitalInput (val(cmbChan))
End Sub

Private Sub cmdDigitalOutput_Click()
    DigitalOutput val(cmbChan), val(txtEng)
End Sub

Private Sub form_resize()
    Me.Height = 4260
    Me.Width = 3105
End Sub

Private Sub cmdClose_Click()
    Me.Hide
End Sub

Private Sub Form_Load()
    Dim i As Integer

    'Load Icon
    If FileExists(Prog_IcoFile) And _
       LenB(Prog_IcoFile) > 0 _
    Then
    
        Me.Icon = LoadPicture(Prog_IcoFile)
        
    End If

    cmbChan.Clear
    For i = 0 To 7
        cmbChan.AddItem Str$(i)
    Next i
    cmbChan.ListIndex = 0     ' Flux counting as default

End Sub


Public Function AnalogInput(ByVal Chan As Long, Optional ByVal gain As Long = BIP10VOLTS) As Double
    If NOCOMM_MODE Then Exit Function

    Dim ULStat As Long
    Dim DataValue As Integer
    Dim engUnits As Single

    txtRaw = vbNullString
    txtEng = vbNullString

   ' Collect the data with cbAIn%()

   '  Parameters:
   '    BoardNum%    :the number used by CB.CFG to describe this board
   '    Chan       :the input channel number
   '    Gain       :the gain for the board.
   '    DataValue%  :the name for the value collected
    
   ULStat = cbAIn(BoardNum%, Chan, gain, DataValue)
   If ULStat = 30 Then MsgBox "Change the Gain argument to one supported by this board.", 0, "Unsupported Gain"
   If ULStat <> 0 Then
        txtEng = "ERROR"
        Exit Function
    End If
   
    cmbChan = Str$(Chan)
   txtRaw = Str$(DataValue)
   
   ULStat = cbToEngUnits(BoardNum%, gain, DataValue, engUnits)
    If ULStat <> 0 Then
        txtEng = "ERROR"
        Exit Function
    End If

    txtEng = Str$(engUnits)
        
    AnalogInput = engUnits

End Function

Public Sub AnalogOutput(ByVal Chan As Long, ByVal engUnits As Double, Optional ByVal Range As Long = BIP10VOLTS)

    If NOCOMM_MODE Then Exit Sub


    Dim ULStat As Long
    Dim DataValue As Integer

     ' send the digital output value to D/A Chan with cbAOut%()

   ' Parameters:
   '   BoardNum    :the number used by CB.CFG to describe this board
   '   Chan%       :the D/A output channel
   '   Range%      :ignored if board does not have programmable rage
   '   DataValue%  :the value to send to Chan%
   
   
   ULStat = cbFromEngUnits(BoardNum%, Range, engUnits, DataValue)
   If ULStat <> 0 Then
        txtEng = "ERROR"
        Exit Sub
   End If
         
   ULStat = cbAOut(BoardNum%, Chan, Range, DataValue)
   If ULStat <> 0 Then
        txtEng = "ERROR"
        Exit Sub
   End If

    cmbChan = Str$(Chan)
    txtRaw = Str$(DataValue)
    txtEng = Str$(engUnits)

End Sub


Public Function DigitalInput(ByVal BitNum As Long) As Long


    Dim ULStat As Long
    Dim DataValue As Integer

    If NOCOMM_MODE Then Exit Function

    If Not DIOConfig(BitNum) = DigitalIn Then
        ULStat = cbDConfigBit(BoardNum%, AUXPORT, BitNum, DigitalIn)
        If ULStat <> 0 Then
                txtEng = "ERROR"
                Exit Function
        End If
        DIOConfig(BitNum) = DigitalIn
    End If

    txtRaw = vbNullString
    txtEng = vbNullString
    
   ' read digital input and display
     
    
   ULStat = cbDBitIn(BoardNum%, AUXPORT, BitNum, DataValue)
   If ULStat <> 0 Then
    txtEng = "ERROR"
    Exit Function
   End If

    DigitalInput = DataValue
    
    cmbChan = Str$(BitNum)
    txtRaw = Str$(DataValue)
    txtEng = Str$(DataValue)
    
End Function

Public Sub DigitalOutput(ByVal BitNum As Long, ByVal DataValue As Long)
   
   Dim ULStat As Long
   
   If NOCOMM_MODE Then Exit Sub
   
   txtRaw = txtEng
   
    If Not DIOConfig(BitNum) = DigitalOut Then
        ULStat = cbDConfigBit(BoardNum%, AUXPORT, BitNum, DigitalOut)
        If ULStat <> 0 Then
                txtEng = "ERROR"
                Exit Sub
        End If
        DIOConfig(BitNum) = DigitalOut
    End If
   
   ' write the value
  
   ULStat = cbDBitOut(BoardNum, AUXPORT, BitNum, DataValue)
   
   If ULStat <> 0 Then
      txtEng = "ERROR"
   End If
    cmbChan = Str$(BitNum)
    txtRaw = Str$(DataValue)
    txtEng = Str$(DataValue)

End Sub
