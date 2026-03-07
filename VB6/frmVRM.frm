VERSION 5.00
Begin VB.Form frmVRM 
   Caption         =   "VRM Decay Test"
   ClientHeight    =   2985
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmVRM.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2985
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtData 
      Height          =   375
      Index           =   2
      Left            =   1440
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   1920
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtData 
      Height          =   375
      Index           =   1
      Left            =   2640
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   1920
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtData 
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   1920
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtTime 
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   1440
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdStartStop 
      Caption         =   "Start"
      Height          =   375
      Left            =   3480
      TabIndex        =   6
      Top             =   2520
      Width           =   975
   End
   Begin VB.TextBox txtFileName 
      Height          =   285
      Left            =   1200
      TabIndex        =   5
      Top             =   840
      Width           =   2775
   End
   Begin VB.ComboBox cmbStepScale 
      Height          =   315
      Left            =   3240
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   240
      Width           =   1095
   End
   Begin VB.TextBox txtStepSpacing 
      Height          =   285
      Left            =   1440
      TabIndex        =   0
      Text            =   "10"
      Top             =   240
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "Write to file:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "seconds"
      Height          =   255
      Left            =   2280
      TabIndex        =   2
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Measure every"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmVRM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CurrentlyRunning As Boolean
Dim LogarithmicStepSpacing As Boolean
Dim StepSpacing As Double
Dim WriteFile As String

Private Sub cmdStartStop_Click()
    If CurrentlyRunning Then
        CurrentlyRunning = False
        cmdStartStop.Caption = "Start"
    Else
        CurrentlyRunning = True
        LogarithmicStepSpacing = cmbStepScale.ListIndex
        StepSpacing = val(txtStepSpacing)
        WriteFile = txtFileName
        cmdStartStop.Caption = "Stop"
        RunVRMtest
    End If
End Sub

Private Sub Form_Load()
    
    '(March 11, 2011 - I Hilburn)
    'Update all the form's icons with the program Ico file
    SetFormIcon Me
    
    cmbStepScale.Clear
    cmbStepScale.AddItem "Linear"
    cmbStepScale.AddItem "Log"
    cmbStepScale.ListIndex = 0
    CurrentlyRunning = False
    LogarithmicStepSpacing = False
    StepSpacing = 1
    
End Sub

Private Sub RunVRMtest()
    Dim cumulativeTime As Double
    Dim lasttime As Double
    Dim waitingTime As Double
    Dim StartTime As Double
    Dim currentTime As Double
    Dim days As Long
    Dim CurrentData As Cartesian3D
    Dim i As Long
    WriteVRMDataHeaders WriteFile
    StartTime = Timer
    lasttime = Timer
    cumulativeTime = 0
    days = 0
    waitingTime = StepSpacing
    txtTime = ""
    txtTime.Visible = True
    For i = 0 To 2
        txtData(i) = ""
        txtData(i).Visible = True
    Next i
    Do While CurrentlyRunning
        DelayTime waitingTime
        currentTime = Timer
        If currentTime < lasttime Then
            days = days + 1
            ' (May 2007 L Carporzen) The previous version add one day at each measurement after midnight.
            If currentTime + (days - 1) * 86400 < lasttime Then
                currentTime = currentTime + days * 86400
            Else
                days = days - 1
                currentTime = currentTime + days * 86400
            End If
        End If
        lasttime = currentTime
        cumulativeTime = currentTime - StartTime
        Set CurrentData = New Cartesian3D
        CurrentData.X = 0
        CurrentData.Y = 0
        CurrentData.Z = 0
        Set CurrentData = frmSQUID.getData
        txtTime = cumulativeTime
        txtData(0) = Str$(CurrentData.X)
        txtData(1) = Str$(CurrentData.Y)
        txtData(2) = Str$(CurrentData.Z)
        WriteVRMData WriteFile, cumulativeTime, CurrentData
        Set CurrentData = Nothing
        If LogarithmicStepSpacing Then
            waitingTime = waitingTime * StepSpacing
        Else
            waitingTime = StepSpacing
        End If
    Loop
    txtTime.Visible = False
    For i = 0 To 2
        txtData(i).Visible = False
    Next i
End Sub

Private Sub WriteVRMData(filename As String, meastime As Double, data As Cartesian3D)
    Dim filenum As Integer
    filenum = FreeFile
    On Error GoTo oops
    Open filename For Append As #filenum
    With data
        Print #filenum, meastime; ","; .X; ","; .Y; ","; .Z
    End With
    Close #filenum
    GoTo stillworking
oops:
    CurrentlyRunning = False
    MsgBox "Unable to write to " & filename & "! Stopping VRM run."
stillworking:
End Sub

Private Sub WriteVRMDataHeaders(filename As String)
    Dim filenum As Integer
    filenum = FreeFile
    On Error GoTo oops
    Open filename For Append As #filenum
        Print #filenum, "Time (s), x, y, z"
    Close #filenum
    GoTo stillworking
oops:
    CurrentlyRunning = False
    MsgBox "Unable to write to " & filename & "! Stopping VRM run."
stillworking:
End Sub

