VERSION 5.00
Begin VB.Form frmStats 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Sample Statistics"
   ClientHeight    =   6060
   ClientLeft      =   5670
   ClientTop       =   315
   ClientWidth     =   5655
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6060
   ScaleWidth      =   5655
   Visible         =   0   'False
   Begin VB.CommandButton cmdHide 
      Caption         =   "&Hide"
      Height          =   345
      Left            =   3480
      TabIndex        =   21
      Top             =   960
      Width           =   858
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   345
      Left            =   4444
      TabIndex        =   20
      Top             =   960
      Width           =   858
   End
   Begin VB.Frame framErrors 
      Height          =   1100
      Left            =   242
      TabIndex        =   13
      Top             =   4750
      Width           =   5093
      Begin VB.Label lblErrAngle 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   253
         Left            =   242
         TabIndex        =   19
         Top             =   560
         Width           =   1342
      End
      Begin VB.Label lblHErrAngle 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   253
         Left            =   1815
         TabIndex        =   18
         Top             =   560
         Width           =   1342
      End
      Begin VB.Label lblMUDRatio 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   253
         Left            =   3388
         TabIndex        =   17
         Top             =   560
         Width           =   1342
      End
      Begin VB.Label Label26 
         Caption         =   "Circular Std. Dev."
         Height          =   253
         Left            =   242
         TabIndex        =   16
         Top             =   300
         Width           =   1320
      End
      Begin VB.Label Label27 
         Caption         =   "Horiz. Err. Angle"
         Height          =   253
         Left            =   1815
         TabIndex        =   15
         Top             =   300
         Width           =   1320
      End
      Begin VB.Label Label28 
         Caption         =   "Up/Down Ratio"
         Height          =   253
         Left            =   3388
         TabIndex        =   14
         Top             =   300
         Width           =   1320
      End
   End
   Begin VB.Frame framStats 
      Height          =   3300
      Left            =   242
      TabIndex        =   0
      Top             =   1440
      Width           =   5082
      Begin VB.Label Label25 
         Alignment       =   2  'Center
         Caption         =   "Bedding"
         Height          =   253
         Left            =   3872
         TabIndex        =   50
         Top             =   2228
         Width           =   858
      End
      Begin VB.Label Label24 
         Alignment       =   2  'Center
         Caption         =   "Geographic"
         Height          =   253
         Left            =   2904
         TabIndex        =   49
         Top             =   2228
         Width           =   858
      End
      Begin VB.Label Label15 
         Caption         =   "Average Inc:"
         Height          =   253
         Left            =   242
         TabIndex        =   47
         Top             =   2833
         Width           =   1221
      End
      Begin VB.Label Label14 
         Caption         =   "Average Dec:"
         Height          =   253
         Left            =   242
         TabIndex        =   46
         Top             =   2470
         Width           =   1221
      End
      Begin VB.Label lblCInc 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   253
         Left            =   1936
         TabIndex        =   45
         Top             =   2833
         Width           =   847
      End
      Begin VB.Label lblGInc 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   253
         Left            =   2893
         TabIndex        =   44
         Top             =   2833
         Width           =   858
      End
      Begin VB.Label lblBInc 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   253
         Left            =   3861
         TabIndex        =   43
         Top             =   2833
         Width           =   858
      End
      Begin VB.Label lblCDec 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   253
         Left            =   1936
         TabIndex        =   42
         Top             =   2470
         Width           =   847
      End
      Begin VB.Label lblGDec 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   253
         Left            =   2893
         TabIndex        =   41
         Top             =   2470
         Width           =   858
      End
      Begin VB.Label lblBDec 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   253
         Left            =   3861
         TabIndex        =   40
         Top             =   2470
         Width           =   858
      End
      Begin VB.Label Label3 
         Caption         =   "Moment/Vol Ratio"
         Height          =   253
         Left            =   3388
         TabIndex        =   39
         Top             =   1573
         Width           =   1320
      End
      Begin VB.Label lblMomVol 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   253
         Left            =   3388
         TabIndex        =   38
         Top             =   1815
         Width           =   1342
      End
      Begin VB.Label lblHolderZ 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   253
         Left            =   2167
         TabIndex        =   37
         Top             =   1815
         Width           =   858
      End
      Begin VB.Label lblHolderY 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   253
         Left            =   1199
         TabIndex        =   36
         Top             =   1815
         Width           =   858
      End
      Begin VB.Label lblHolderX 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   253
         Left            =   242
         TabIndex        =   35
         Top             =   1815
         Width           =   847
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         Caption         =   "Holder X"
         Height          =   253
         Left            =   242
         TabIndex        =   34
         Top             =   1573
         Width           =   847
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         Caption         =   "Holder Y"
         Height          =   253
         Left            =   1210
         TabIndex        =   33
         Top             =   1573
         Width           =   858
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         Caption         =   "Holder Z"
         Height          =   253
         Left            =   2167
         TabIndex        =   32
         Top             =   1573
         Width           =   858
      End
      Begin VB.Label Label17 
         Caption         =   "Std. Dev. X"
         Height          =   253
         Left            =   242
         TabIndex        =   12
         Top             =   363
         Width           =   858
      End
      Begin VB.Label Label18 
         Caption         =   "Std. Dev. Y"
         Height          =   253
         Left            =   242
         TabIndex        =   11
         Top             =   740
         Width           =   858
      End
      Begin VB.Label Label19 
         Caption         =   "Std. Dev. Z"
         Height          =   253
         Left            =   242
         TabIndex        =   10
         Top             =   1100
         Width           =   858
      End
      Begin VB.Label lblsdX 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   253
         Left            =   1331
         TabIndex        =   9
         Top             =   363
         Width           =   858
      End
      Begin VB.Label lblsdY 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   253
         Left            =   1331
         TabIndex        =   8
         Top             =   740
         Width           =   858
      End
      Begin VB.Label lblsdZ 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   253
         Left            =   1331
         TabIndex        =   7
         Top             =   1100
         Width           =   858
      End
      Begin VB.Label lblSigInd 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   253
         Left            =   3872
         TabIndex        =   6
         Top             =   1100
         Width           =   858
      End
      Begin VB.Label lblSigNoise 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   253
         Left            =   3872
         TabIndex        =   5
         Top             =   363
         Width           =   858
      End
      Begin VB.Label lblSigHold 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   253
         Left            =   3861
         TabIndex        =   4
         Top             =   740
         Width           =   858
      End
      Begin VB.Label Label20 
         Caption         =   "Signal/Noise:"
         Height          =   252
         Left            =   2544
         TabIndex        =   3
         Top             =   360
         Width           =   1308
      End
      Begin VB.Label Label5 
         Caption         =   "Signal/Holder:"
         Height          =   252
         Left            =   2532
         TabIndex        =   2
         Top             =   744
         Width           =   1320
      End
      Begin VB.Label Label6 
         Caption         =   "Signal/Induced:"
         Height          =   252
         Left            =   2532
         TabIndex        =   1
         Top             =   1104
         Width           =   1320
      End
      Begin VB.Label Label23 
         Alignment       =   2  'Center
         Caption         =   "Core"
         Height          =   253
         Left            =   1936
         TabIndex        =   48
         Top             =   2228
         Width           =   858
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Sample:"
      Height          =   253
      Index           =   2
      Left            =   242
      TabIndex        =   31
      Top             =   242
      Width           =   847
   End
   Begin VB.Label Label1 
      Caption         =   "Avg. Cycles"
      Height          =   252
      Index           =   1
      Left            =   2172
      TabIndex        =   30
      Top             =   240
      Width           =   960
   End
   Begin VB.Label Label8 
      Caption         =   "Demag:"
      Height          =   253
      Left            =   3718
      TabIndex        =   29
      Top             =   242
      Width           =   737
   End
   Begin VB.Label Label7 
      Caption         =   "File Path:"
      Height          =   253
      Left            =   242
      TabIndex        =   28
      Top             =   605
      Width           =   847
   End
   Begin VB.Label lblSampName 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Height          =   253
      Left            =   1078
      TabIndex        =   27
      Top             =   242
      Width           =   979
   End
   Begin VB.Label lblAvgCycles 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Height          =   253
      Left            =   3124
      TabIndex        =   26
      Top             =   242
      Width           =   495
   End
   Begin VB.Label lblDemag 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Height          =   253
      Left            =   4444
      TabIndex        =   25
      Top             =   242
      Width           =   858
   End
   Begin VB.Label lblDataFileName 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1080
      TabIndex        =   24
      Top             =   600
      Width           =   4305
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblDirs 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Height          =   253
      Left            =   1089
      TabIndex        =   23
      Top             =   968
      Width           =   495
   End
   Begin VB.Label Label4 
      Caption         =   "Directions:"
      Height          =   253
      Left            =   242
      TabIndex        =   22
      Top             =   968
      Width           =   858
   End
End
Attribute VB_Name = "frmStats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdHide_Click()
    Me.Hide
    frmMeasure.cmdStats.Enabled = True
End Sub

Private Sub cmdPrint_Click()
    On Error GoTo ErrorHandler
    Dim numlines As Integer
    cmdPrint.Enabled = False
    numlines = 19
    If Print_LinesLeft < numlines Then
        Print_PageBreak
    End If
    Print_Line
    Print_Line "Statistics:"
    Print_Line "Sample - " & lblSampName.Caption
    Print_Line
    Print_Line "sdX: " & lblsdX.Caption & vbTab & _
                  "sdy: " & lblsdY.Caption & vbTab & _
                  "sdz: " & lblsdZ.Caption
    Print_Line
    Print_Line "Signal / Noise:  " & vbTab & lblSigNoise.Caption
    Print_Line "Signal / Holder: " & vbTab & lblSigHold.Caption
    Print_Line "Signal / Induced:" & vbTab & lblSigInd.Caption
    Print_Line "Moment / Vol:    " & vbTab & lblMomVol.Caption
    Print_Line
    Print_Line "Holder x: " & lblHolderX.Caption & vbTab & _
                  "Holder y: " & lblHolderY.Caption & vbTab & _
                  "Holder z: " & lblHolderZ.Caption
    Print_Line
    Print_Line vbTab & vbTab & "Core" & vbTab & vbTab & _
                                  "Geog." & vbTab & vbTab & _
                                  "Bedding"
    Print_Line "Dec:" & vbTab & vbTab & _
                  lblCDec.Caption & vbTab & _
                  lblGDec.Caption & vbTab & _
                  lblBDec.Caption
    Print_Line "Inc:" & vbTab & vbTab & _
                  lblCInc.Caption & vbTab & _
                  lblGInc.Caption & vbTab & _
                  lblBInc.Caption
    Print_Line
    Exit Sub
ErrorHandler:
    MsgBox ("There was a problem printing to the printer.")
End Sub

Private Sub Form_Load()
    
    '(March 11, 2011 - I Hilburn)
    'Update all the form's icons with the program Ico file
    SetFormIcon Me
    
    HideErrors
    
End Sub

Public Sub HideErrors()
    ' Hide the error frame and shrink the size of the form
    framErrors.Visible = False
    Me.Height = 5226
    lblErrAngle.Caption = vbNullString
    lblHErrAngle.Caption = vbNullString
    lblMUDRatio.Caption = vbNullString
End Sub

Public Sub ShowAvgStats(sdx As Double, sdy As Double, sdz As Double, _
    crdec As Double, crinc As Double, gdec As Double, ginc As Double, _
    bdec As Double, binc As Double, momentvol As Double, _
    SigNoise As Double, SigHolder As Double, Optional SigInduced As Double)
    ' This procedure displays statistical information for the
    ' entire set of data gathered from the magnetometer.  (after
    ' all 'n' averaging cycles have been completed)
    lblHolderX.Caption = Format$(Holder.Average.X, "0.0000E+")
    lblHolderY.Caption = Format$(Holder.Average.Y, "0.0000E+")
    lblHolderZ.Caption = Format$(Holder.Average.Z, "0.0000E+") ' was define by Holder.Average.X before Karin Louzada report the display bug
    lblsdX.Caption = FormatNumber(sdx)
    lblsdY.Caption = FormatNumber(sdy)
    lblsdZ.Caption = FormatNumber(sdz)
    lblMomVol.Caption = Format$(momentvol, "0.0000E+")
    lblCDec.Caption = FormatNumber(crdec)
    lblCInc.Caption = FormatNumber(crinc)
    lblGDec.Caption = FormatNumber(gdec)
    lblGInc.Caption = FormatNumber(ginc)
    lblBDec.Caption = FormatNumber(bdec)
    lblBInc.Caption = FormatNumber(binc)
    lblSigNoise.Caption = FormatNumber(SigNoise)
    lblSigHold.Caption = FormatNumber(SigHolder)
'    If isPaleomag Then
        lblSigInd.Caption = FormatNumber(SigInduced)
'    End If
' (September 2007 L Carporzen) Don't show the Stats window if it has been hide
    If frmMeasure.cmdStats.Enabled = False Then
        Me.ZOrder
        Me.Show
    End If
End Sub

Public Sub ShowErrors(errangle As Double, herrangle As Double, _
    momentupdown As Double)
    ' This procedure is called if we are doing both up and down
    ' measurements.  It displays error fields specific for this
    ' kind of measurement
    Me.Height = 6336
    framErrors.Visible = True
    lblErrAngle.Caption = FormatNumber(errangle)
    lblHErrAngle.Caption = FormatNumber(herrangle)
    lblMUDRatio.Caption = FormatNumber(momentupdown)
End Sub

