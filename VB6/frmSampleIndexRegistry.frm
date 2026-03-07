VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmSampleIndexRegistry 
   Caption         =   "Sample Index Registry"
   ClientHeight    =   7530
   ClientLeft      =   1305
   ClientTop       =   2145
   ClientWidth     =   8010
   Icon            =   "frmSampleIndexRegistry.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7530
   ScaleWidth      =   8010
   Begin VB.Frame framFiles 
      Caption         =   "Load SAM File"
      Height          =   3855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7935
      Begin VB.CommandButton LoadSamplesButton 
         Caption         =   "Move to Load Position"
         Height          =   375
         Left            =   4320
         TabIndex        =   45
         Top             =   2760
         Width           =   2055
      End
      Begin MSComDlg.CommonDialog dlgCommonDialog 
         Left            =   0
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.TextBox txtDir 
         Height          =   285
         Left            =   1560
         TabIndex        =   1
         Top             =   240
         Width           =   4935
      End
      Begin VB.CommandButton cmdBrowseDir 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   275
         Left            =   6600
         TabIndex        =   2
         Top             =   252
         Width           =   375
      End
      Begin VB.ComboBox cmbSampCode 
         Height          =   288
         Left            =   1560
         TabIndex        =   3
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox txtSampDesc 
         Enabled         =   0   'False
         Height          =   492
         Left            =   3240
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   720
         Width           =   3735
      End
      Begin VB.CheckBox chkBck 
         Height          =   255
         Left            =   1548
         TabIndex        =   5
         Top             =   1320
         Value           =   1  'Checked
         Width           =   180
      End
      Begin VB.DriveListBox lstDrvBck 
         Height          =   288
         Left            =   1872
         TabIndex        =   6
         Top             =   1320
         Width           =   1265
      End
      Begin VB.TextBox txtBck 
         Height          =   285
         Left            =   3252
         TabIndex        =   7
         Top             =   1320
         Width           =   3255
      End
      Begin VB.OptionButton optSAMSetDemag 
         Caption         =   "NRM"
         Height          =   192
         Index           =   0
         Left            =   2040
         TabIndex        =   8
         Top             =   1800
         Value           =   -1  'True
         Width           =   732
      End
      Begin VB.OptionButton optSAMSetDemag 
         Caption         =   "AF"
         Height          =   192
         Index           =   1
         Left            =   2760
         TabIndex        =   9
         Top             =   1800
         Width           =   612
      End
      Begin VB.OptionButton optSAMSetDemag 
         Caption         =   "TT"
         Height          =   192
         Index           =   2
         Left            =   3360
         TabIndex        =   10
         Top             =   1800
         Width           =   612
      End
      Begin VB.OptionButton optSAMSetDemag 
         Caption         =   "MW"
         Height          =   192
         Index           =   3
         Left            =   2040
         TabIndex        =   11
         Top             =   2040
         Width           =   732
      End
      Begin VB.OptionButton optSAMSetDemag 
         Caption         =   "Rockmag"
         Height          =   192
         Index           =   5
         Left            =   2760
         TabIndex        =   12
         Top             =   2040
         Width           =   1092
      End
      Begin VB.OptionButton optSAMSetDemag 
         Caption         =   "Other:"
         Height          =   192
         Index           =   4
         Left            =   3960
         TabIndex        =   13
         Top             =   2040
         Width           =   732
      End
      Begin VB.TextBox txtSAMSetDemag 
         Enabled         =   0   'False
         Height          =   288
         Left            =   4680
         TabIndex        =   14
         Top             =   2004
         Width           =   732
      End
      Begin VB.TextBox txtSAMSetDemagLevel 
         Height          =   288
         Left            =   5760
         TabIndex        =   15
         Text            =   "0"
         Top             =   2004
         Width           =   492
      End
      Begin VB.CommandButton cmdRMLevel 
         Caption         =   "Set Levels"
         Height          =   252
         Left            =   5760
         TabIndex        =   16
         Top             =   2040
         Visible         =   0   'False
         Width           =   1092
      End
      Begin VB.CheckBox chkSAMdoUp 
         Caption         =   "Up"
         Height          =   252
         Left            =   2040
         TabIndex        =   17
         Top             =   2400
         Value           =   1  'Checked
         Width           =   612
      End
      Begin VB.CheckBox chkSAMdoDown 
         Caption         =   "Down"
         Height          =   252
         Left            =   2640
         TabIndex        =   18
         Top             =   2400
         Value           =   1  'Checked
         Width           =   732
      End
      Begin VB.CheckBox chkSAMalreadyDoneUp 
         Caption         =   "Up already measured"
         Height          =   252
         Left            =   2040
         TabIndex        =   19
         Top             =   2640
         Width           =   1932
      End
      Begin VB.TextBox txtSAMaveragesteps 
         Height          =   300
         Left            =   6840
         TabIndex        =   20
         Text            =   "1"
         Top             =   2400
         Width           =   732
      End
      Begin VB.CheckBox chkMeasureSusceptibility 
         Caption         =   "Measure susceptibility"
         Height          =   252
         Left            =   120
         TabIndex        =   21
         Top             =   3360
         Value           =   1  'Checked
         Width           =   1932
      End
      Begin VB.CommandButton buttonAddToSamRegistry 
         Caption         =   "Add to registry"
         Height          =   372
         Left            =   4320
         TabIndex        =   22
         Top             =   3240
         Width           =   1572
      End
      Begin VB.CommandButton buttonDataAnalysis 
         Caption         =   "Open SAM file"
         Height          =   372
         Left            =   6120
         TabIndex        =   23
         Top             =   3240
         Width           =   1572
      End
      Begin VB.Label SampleTableTypeLabel 
         Caption         =   "Chain Drive"
         Height          =   255
         Left            =   2040
         TabIndex        =   44
         Top             =   3000
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "Sample Table Type:"
         Height          =   255
         Left            =   120
         TabIndex        =   43
         Top             =   3000
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Data Directory:"
         Height          =   252
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   1212
      End
      Begin VB.Label Label2 
         Caption         =   "Sample Code:"
         Height          =   252
         Left            =   120
         TabIndex        =   25
         Top             =   720
         Width           =   1092
      End
      Begin VB.Label Label3 
         Caption         =   "Backup Data:"
         Height          =   252
         Left            =   132
         TabIndex        =   26
         Top             =   1320
         Width           =   1092
      End
      Begin VB.Label Label15 
         Caption         =   "Type of demagnetization step:"
         Height          =   372
         Left            =   120
         TabIndex        =   27
         Top             =   1800
         Width           =   1812
      End
      Begin VB.Label Label14 
         Caption         =   "Level:"
         Height          =   252
         Left            =   5760
         TabIndex        =   28
         Top             =   1800
         Width           =   492
      End
      Begin VB.Label Label16 
         Caption         =   "Directions to Measure:"
         Height          =   252
         Left            =   120
         TabIndex        =   29
         Top             =   2400
         Width           =   1692
      End
      Begin VB.Label Label17 
         Caption         =   "Measurement blocks per cycle:"
         Height          =   255
         Left            =   4440
         TabIndex        =   30
         Top             =   2445
         Width           =   2415
      End
   End
   Begin VB.Frame framInfo 
      Caption         =   "File Info"
      Height          =   972
      Left            =   0
      TabIndex        =   31
      Top             =   3960
      Width           =   7935
      Begin VB.Label lblLoc 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Index           =   0
         Left            =   360
         TabIndex        =   32
         Top             =   480
         Width           =   1020
      End
      Begin VB.Label lblLoc 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Index           =   1
         Left            =   1680
         TabIndex        =   33
         Top             =   480
         Width           =   1020
      End
      Begin VB.Label lblLoc 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Index           =   2
         Left            =   3120
         TabIndex        =   34
         Top             =   480
         Width           =   1020
      End
      Begin VB.Label lblSampNum 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   5520
         TabIndex        =   35
         Top             =   480
         Width           =   1188
      End
      Begin VB.Label Label10 
         Caption         =   "Latitude"
         Height          =   252
         Left            =   360
         TabIndex        =   36
         Top             =   240
         Width           =   852
      End
      Begin VB.Label Label11 
         Caption         =   "Longitude"
         Height          =   252
         Left            =   1680
         TabIndex        =   37
         Top             =   240
         Width           =   852
      End
      Begin VB.Label Label12 
         Caption         =   "Mag. Dec"
         Height          =   252
         Left            =   3120
         TabIndex        =   38
         Top             =   240
         Width           =   972
      End
      Begin VB.Label Label6 
         Caption         =   "# of Samples:"
         Height          =   252
         Left            =   5520
         TabIndex        =   39
         Top             =   240
         Width           =   1092
      End
   End
   Begin VB.Frame framSAMRegistry 
      Caption         =   "SAM file registry"
      Height          =   2412
      Left            =   0
      TabIndex        =   40
      Top             =   5040
      Width           =   7935
      Begin ComctlLib.ListView lvwSAMRegistry 
         Height          =   1215
         Left            =   240
         TabIndex        =   41
         Top             =   360
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   2143
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   327682
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.CommandButton buttonClearAll 
         Caption         =   "Clear registry"
         Height          =   372
         Left            =   240
         TabIndex        =   42
         Top             =   1800
         Width           =   1572
      End
   End
End
Attribute VB_Name = "frmSampleIndexRegistry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private MeasSusc As Boolean
Private SampleCode As String
Private DataFileDrv As String
Private DataFileDir As String
Private DataFileName As String
Private fileReadyToLoad As Boolean
Private initialized As Boolean
Private Warning As Boolean
Private UseXYTable As Boolean

Public workingSamIndex As SampleIndexRegistration

Private Sub buttonAddToSamRegistry_Click()
    Dim tempBckDir As String
    Dim filedoup As Boolean
    Dim filedoboth As Boolean
    Dim i As Integer
    Dim curDemag As String
    Dim target As SampleIndexRegistration
    If frmMagnetometerControl.cmdManHolder.Enabled = False And frmMagnetometerControl.cmdChangerEdit.Enabled = False Then Exit Sub ' (September 2007 L Carporzen) Avoid changing the registry before the end of the measurement (the current step labels in the sample file could be changed)
    updateBackupDir
    updateMeasurementSteps
    UpdateDoUpDoBoth
    If Not CheckSetFields Then Exit Sub
    If fileReadyToLoad Then
        Set target = SampleIndexRegistry.AddSampleIndex(workingSamIndex)
        target.loadInfo
    End If
    refreshSAMRegistryDisplay
    frmPlots.RefreshSamples
    If Not LenB(frmMagnetometerControl.cmbManSample.text) = 0 Then
        frmMagnetometerControl.RefreshManSampleList ' (September 2007 L Carporzen) Refresh the sample list
        frmMagnetometerControl.cmbManSample.text = "" ' (September 2007 L Carporzen) Empty the sample name in the Manual Data Collection window
    End If
    If Warning = False Then
    
        '(July 2011 - I Hilburn)
        'Warning message statements revised to include switching on the Coil thermal sensors
    
        Dim rockmagMsg As String
        rockmagMsg = vbNullString

        
        
        If (optSAMSetDemag(1).value = True Or optSAMSetDemag(5).value = True) And _
           frmRockmagRoutine.chksusceptibility.value = Checked _
        Then
         'automatically turn on the air
            If modConfig.DoDegausserCooling = True Then
                frmVacuum.DegausserCooler True
                rockmagMsg = "Please: " & vbNewLine & vbNewLine & _
                         " - Verify the air is on" & vbNewLine & _
                         " - Make sure the susceptibility meter is well positioned."
            Else
            rockmagMsg = "Please: " & vbNewLine & vbNewLine & _
                         " - Turn the air on" & vbNewLine & _
                         " - Make sure the susceptibility meter is well positioned."
                        ' (October 2007 L Carporzen)
            End If
                        
            Warning = True
            
        ElseIf optSAMSetDemag(1).value = True Or optSAMSetDemag(5).value = True Then
         'automatically turn on the air
            If modConfig.DoDegausserCooling = True Then
                frmVacuum.DegausserCooler True
                rockmagMsg = "Please: " & vbNewLine & vbNewLine & _
                         " - Verify the air is on" & vbNewLine & _
                         " - Make sure the susceptibility meter is well positioned."
            Else
            rockmagMsg = "Please: " & vbNewLine & vbNewLine & _
                         " - Turn the air on" & vbNewLine & _
                         " - Make sure the susceptibility meter is well positioned."
                        ' (October 2007 L Carporzen)
            End If
            Warning = True
        ElseIf chkMeasureSusceptibility.value = Checked Then
            rockmagMsg = "Please: " & vbNewLine & vbNewLine & _
                         " - Make sure the susceptibility meter is well positioned." & vbNewLine  '(May 2007 L Carporzen)
            Warning = True
        End If
        
        If (optSAMSetDemag(1).value = True Or optSAMSetDemag(5).value = True) And _
           (EnableT1 = True Or EnableT2 = True) _
        Then
            rockmagMsg = rockmagMsg & " - Switch the power on for the Rockmag coil thermal sensors"
        End If
           
        If Len(rockmagMsg) > 0 Then
            MsgBox rockmagMsg
        End If
        
    End If
    If Prog_halted Then ' (September 2007 L Carporzen) New version of the Halt button
        HolderMeasured = False
        Flow_Resume
        frmMeasure.updateFlowStatus
    End If
End Sub

Private Sub buttonClearAll_Click()
    If frmMagnetometerControl.cmdManHolder.Enabled = False And frmMagnetometerControl.cmdChangerEdit.Enabled = False Then Exit Sub ' (September 2007 L Carporzen) Avoid changing the registry before the end of the measurement (the current step labels in the sample file could be changed)
    SampQueue.Clear
    MainChanger.Clear
    SampleIndexRegistry.Clear
    Set workingSamIndex = Nothing
    Set workingSamIndex = New SampleIndexRegistration
    refreshSAMRegistryDisplay
    refreshFields
    chkBck.value = Checked
    chkMeasureSusceptibility.value = Checked
    chkSAMdoUp.value = Checked
    chkSAMdoDown.value = Checked
    Warning = False
End Sub

Private Sub buttonDataAnalysis_Click()
    If fileReadyToLoad Then DataAnalysis_SAMFile DataFileName, txtDir
End Sub

Private Function CheckCurrentDemagString() As Boolean
    Dim demaglev As Long
    Dim demagstrlen As Integer
    Dim currentDemagString As String
    Dim i As Integer
    CheckCurrentDemagString = True
    ' Demag string can be no longer than DEMAGLEN chars
    For i = 0 To 5
        ' Determine which option button is selected
        If optSAMSetDemag(i).value = True Then Exit For
    Next i
    If i = 4 Then
        currentDemagString = txtSAMSetDemag
        If Len(currentDemagString) >= DEMAGLEN Then
            ' The string is too long, error message
            MsgBox ("The demag string can be no longer than " & _
                DEMAGLEN - 1 & " characters.")
                txtSAMSetDemag.SetFocus
                txtSAMSetDemag.SelStart = 0
                txtSAMSetDemag.SelLength = demagstrlen
                CheckCurrentDemagString = False
            Exit Function
        End If
    ElseIf i = 5 Then
        currentDemagString = "RkMg"
    Else
        currentDemagString = optSAMSetDemag(i).Caption
    End If
    demagstrlen = Len(currentDemagString)
    demaglev = val(txtSAMSetDemagLevel.text)
    If demaglev < 0 Or demaglev > ((10 ^ (DEMAGLEN - demagstrlen)) - 1) Then
        ' The demag level is out of bounds.
        MsgBox ("The demag level must be between 0 and " & _
            ((10 ^ (DEMAGLEN - demagstrlen)) - 1) & ".")
        txtSAMSetDemagLevel.SetFocus
        CheckCurrentDemagString = False
        Exit Function
    End If
End Function

'-----------------------------------------------------------------------------
'   CheckSetFields
'
'   Description:        This function determines whether all the fields have
'                       valid values.  If not, an error message is given.
'                       Otherwise, the function returns true.
'   Revision History:
'      Albert Hsiao     2/9/99       updated comments
'      Albert Hsiao     2/24/99      limited demag strings to DEMAGLEN chars
'
Private Function CheckSetFields() As Boolean
    Dim i As Integer
    Dim tmpint As Integer
    Dim tmpstr As String
    Dim demagstrlen As Integer, demaglev As Long
    CheckSetFields = False
    tmpstr = txtSAMaveragesteps.text
    tmpint = val(tmpstr)
    If tmpint <= 0 Then
        MsgBox ("Value must be greater than 0.")
        txtSAMaveragesteps.SelStart = 0
        txtSAMaveragesteps.SelLength = Len(tmpstr)
        txtSAMaveragesteps.SetFocus
        Exit Function
    End If
'    If isBiomag And tmpint < 3 Then
'        MsgBox ("At least 3 steps must be done on biomag.")
'        txtAvgSteps.SelStart = 0
'        txtAvgSteps.SelLength = Len(tmpstr)
'        txtAvgSteps.SetFocus
'        Exit Function
'    End If
    If Not chkSAMdoUp.value = Checked And _
        Not chkSAMdoDown.value = Checked Then
        ' Neither direction boxes are checked, at least one
        ' must be checked to do anything.
        MsgBox ("Please select a measurement direction.")
        chkSAMdoUp.SetFocus
        Exit Function
    End If
    CheckSetFields = True
End Function

Private Sub chkBck_Click()
    ' Enable or disable backup of files (by default this is on)
    If chkBck.value = 1 Then
        lstDrvBck.Enabled = True
        txtBck.Enabled = True
        'cmdBrowseBck.Enabled = True
    Else
        lstDrvBck.Enabled = False
        txtBck.Enabled = False
        'cmdBrowseBck.Enabled = False
    End If
End Sub

Private Sub chkSAMalreadyDoneUp_Click()
    If chkSAMalreadyDoneUp.value Then chkSAMdoUp.value = False
End Sub

Private Sub chkSAMdoUp_Click()
    If chkSAMdoUp.value Then chkSAMalreadyDoneUp.value = False ' (October 2007 L Carporzen)
End Sub

Sub ClearCmbSampCode()
    ' This subroutine clears the combo box that lists the samples
    ' that it finds in the current directory specified by txtDir
    Do While cmbSampCode.ListCount > 0
        cmbSampCode.RemoveItem (0)
    Loop
End Sub

Private Sub cmbSampCode_Change()
    ' This procedure reads the sample header file when a new one is selected
    Dim filedir  As String         ' String that holds the root path
    Dim filename As String         ' String that holds the filename + path
    Dim filenum
    Dim LineTxt  As String
    Dim ind      As Integer        ' Index of first "\" in directory name
    Dim i        As Integer
    Dim selectedSamFileId As Integer
'    On Error GoTo ErrorHandler                        ' Turn on error handling
    If LenB(cmbSampCode.text) = 0 Or LenB(txtDir.text) = 0 Then
        ' Exit if no sample specified
        txtSampDesc.text = vbNullString
        fileReadyToLoad = False
        buttonAddToSamRegistry.Enabled = False
        Exit Sub
    End If
    If cmbSampCode.ListIndex = -1 Then
        ' Exit if sample specified does not exist
        ind = 0
        For i = 1 To cmbSampCode.ListCount
            ' Look to see if something in the list matches; and if we
            ' find something, then change the value of 'ind'
            If StrComp(cmbSampCode.List(i - 1), cmbSampCode.text, _
                vbTextCompare) = 0 Then
                ind = i
                Exit For
            End If
        Next i
        If ind = 0 Then
            ' We did not find a match in the list
            txtSampDesc.text = vbNullString
            fileReadyToLoad = False
            buttonAddToSamRegistry.Enabled = False
            Exit Sub
        End If
    End If
    filename = txtDir.text
    If Not Mid(filename, Len(filename), Len(filename)) = "\" Then
        filename = filename & "\"
    End If
    filedir = filename
    filename = filename & cmbSampCode.text & "\" & cmbSampCode.text & ".sam"
    filenum = FreeFile
    If Not FileExists(filename) Then
        ' Exit if file doesn't exist
        txtSampDesc.text = vbNullString
        fileReadyToLoad = False
        buttonAddToSamRegistry.Enabled = False
    End If
    SampleCode = cmbSampCode.text
    ind = InStr(filedir, "\")                         ' Find the first "\"
    DataFileDrv = Mid(txtDir.text, 1, ind - 1)
    DataFileDir = Mid(txtDir.text, ind + 1, Len(txtDir.text))
    DataFileName = filename
    With workingSamIndex
        .SampleCode = SampleCode
        .filedir = filedir
        .filename = filename
        .loadInfo
    End With
    ReadSamInfo
    ' Update the form
    fileReadyToLoad = True
    buttonAddToSamRegistry.Enabled = True
    ' We've changed the data directories, update the backup
    Call lstDrvBck_Change
    Exit Sub
ErrorHandler:
    fileReadyToLoad = False
    buttonAddToSamRegistry.Enabled = False
    Select Case Err.number  ' Evaluate error number.
        Case 55 ' "File already open" error.
            Close #1    ' Close open file.
        Case Else
            MsgBox ("Error " & Err.number & _
            " found in cmbSampCode_Change() " & Err.Description)        ' Handle other situations here...
    End Select
    Exit Sub
End Sub

Private Sub cmbSampCode_Click()
    ' We may have changed the contents of the combobox
    Call cmbSampCode_Change
End Sub

Private Sub cmdBrowseDir_Click()
    Dim sfilen  As String           ' Path + filename of selected file
    Dim dirname As String           ' Path of selected file
    Dim fname   As String           ' Filename of selected file
    Dim ind     As Variant          ' index of first character of filename
    ' Initialize the dialog box for Sample File Open
    dlgCommonDialog.filter = "Sample description file (*.sam)|*.sam|All files (*.*)|*.*"
    dlgCommonDialog.flags = cdlOFNHideReadOnly Or cdlOFNFileMustExist
    dlgCommonDialog.DialogTitle = "Open Sample description file..."
    If LenB(txtDir.text) <> 0 Then
        dlgCommonDialog.InitDir = txtDir.text
    Else
        If FileExists(Prog_DefaultPath) Then
            dlgCommonDialog.InitDir = Prog_DefaultPath
        Else
            dlgCommonDialog.InitDir = "\"
        End If
    End If
    dlgCommonDialog.ShowOpen
    ' ----- Start parsing the filename -----
    ' Parse the file name
    fname = dlgCommonDialog.FileTitle
    ind = InStr(UCase$(fname), ".SAM")              ' Find the first character of extension
    If (ind > 1) Then
        fname = Mid$(fname, 1, ind - 1)
    End If
    ' Parse the directory name
    sfilen = dlgCommonDialog.filename
    If LenB(sfilen) = 0 Then Exit Sub          ' If we don't have a filename
                                            ' then don't processs it.
    ind = InStr(UCase$(sfilen), UCase$(fname))        ' Get first char of samplename
    dirname = Mid$(sfilen, 1, ind - 1)       ' Take sample name out of path
    ' Update the form
    txtDir.text = dirname
    cmbSampCode.text = fname
ErrHandler:
    'User pressed the Cancel button
    Exit Sub
End Sub

Private Sub cmdRMLevel_Click()
'   You should be able to modify rock mag sequences on the fly if you're careful about it.
'   This modification is counter-productive.
'    If frmMagnetometerControl.cmdManHolder.Enabled = False And frmMagnetometerControl.cmdChangerEdit.Enabled = False Then Exit Sub ' (September 2007 L Carporzen) Avoid changing the registry before the end of the measurement (the current step labels in the sample file could be changed)
    If optSAMSetDemag(1).value = True Then
        ' if this is AF
        frmRockmagRoutine.ZOrder
        frmRockmagRoutine.Show
    ElseIf optSAMSetDemag(5).value = True Then
        ' if this is Rock Mag
        frmRockmagRoutine.ZOrder
        frmRockmagRoutine.Show
    End If
    updateMeasurementSteps
End Sub

Private Sub Form_Hide(Cancel As Integer)
    'Close all sub forms
    If Me.WindowState <> vbMinimized Then
        Config_SaveSetting "Program", "SampleIndexRegistryWindowLeft", str(Me.Left)
        Config_SaveSetting "Program", "SampleIndexRegistryWindowTop", str(Me.Top)
    End If
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim drvfound As Boolean
    Dim colX As ColumnHeader
    On Error GoTo ErrorHandler
    
    '(March 11, 2011 - I Hilburn)
    'Update all the form's icons with the program Ico file
    SetFormIcon Me
    
    Set workingSamIndex = New SampleIndexRegistration
    Set colX = lvwSAMRegistry.ColumnHeaders.Add(1)
    colX.text = "Sample set"
    colX.Width = Me.TextWidth(colX.text & " ")
    Set colX = lvwSAMRegistry.ColumnHeaders.Add(2)
    colX.text = "Step"
    colX.Width = Me.TextWidth(colX.text & " ")
    Set colX = lvwSAMRegistry.ColumnHeaders.Add(3)
    colX.text = "Do up?"
    colX.Width = Me.TextWidth(colX.text)
    Set colX = lvwSAMRegistry.ColumnHeaders.Add(4)
    colX.text = "Do both?"
    colX.Width = Me.TextWidth(colX.text)
    Set colX = lvwSAMRegistry.ColumnHeaders.Add(5)
    colX.text = "Blocks"
    colX.Width = Me.TextWidth(colX.text)
    Set colX = lvwSAMRegistry.ColumnHeaders.Add(6)
    colX.text = "Path"
    colX.Width = Me.TextWidth(colX.text & "WWWWWWWWWWW")
    Set colX = Nothing
    ' Initialize private variables
    initialized = False
    HolderMeasured = False
    LoadResStrings Me
    ' Load settings from previous run of this program
    Me.Left = val(Config_GetSetting("Program", "MainWindowLeft", "1000"))
    Me.Top = val(Config_GetSetting("Program", "MainWindowTop", "1000"))
    ' Set initial window conditions
    refreshSAMRegistryDisplay
    buttonAddToSamRegistry.Enabled = False
    fileReadyToLoad = False
    ' Set the default values for controls
    For i = 0 To lstDrvBck.ListCount
        If UCase(Left(lstDrvBck.List(i), 1)) = UCase(Prog_DefaultBackup) Then
            lstDrvBck.ListIndex = i
            drvfound = True
        End If
    Next i
    cmdRMLevel.Visible = False
    txtSAMSetDemagLevel.Visible = True
    chkBck.value = Checked
    chkMeasureSusceptibility.value = Checked
    chkSAMdoUp.value = Checked
    chkSAMdoDown.value = Checked
FormLoadChkDrive:
    If Not drvfound Then chkBck.value = Unchecked
    Exit Sub
ErrorHandler:
    Select Case Err.number
        Case 68
            ' Default backup drive was not accessible.
            GoTo FormLoadChkDrive
        Case Else
            MsgBox "Unknown error " & Err.number & " occurred in " & _
                "frmMagnetometerControl.Form_Load()", vbCritical, "Unhandled error!"
            End
    End Select
End Sub

Private Sub Form_Resize()
    Warning = False
    If Me.WindowState = vbNormal Then
        Me.Height = 8040
        Me.Width = 8130
    End If
End Sub

Private Sub form_show()
    Warning = False
    Me.Left = val(Config_GetSetting("Program", "SampleIndexRegistryWindowLeft", "0"))
    Me.Top = val(Config_GetSetting("Program", "SampleIndexRegistryWindowLeft", "0"))
    'UseXYTable = Config_GetSetting("XYTable", "UseXYTableAPS", "False")
    If UseXYTableAPS Then
    SampleTableTypeLabel.Caption = "XY Table APS"
    LoadSamplesButton.Visible = True
    Else
    SampleTableTypeLabel.Caption = "Chain Drive"
    LoadSamplesButton.Visible = False
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer
    If Me.WindowState <> vbMinimized Then
        Config_SaveSetting "Program", "MainWindowLeft", str(Me.Left)
        Config_SaveSetting "Program", "MainWindowTop", str(Me.Top)
    End If
    Set workingSamIndex = Nothing
End Sub

Private Sub LoadSamplesButton_Click()
' Move to the corner
    frmDCMotors.MoveToCorner pauseOveride:=True
End Sub

Private Sub lstDrvBck_Change()
    Dim DriveSelected As String            ' Current drive selected
    Dim ind           As Integer           ' index of first ":"
    Dim BackFileDrv As String
    Dim BackFileDir As String
    Dim shareEnds As Long
    ' Drive is selected, change the directory of the backup files
    If LenB(SampleCode) <> 0 Then
        ind = InStr(lstDrvBck.Drive, ":")
        BackFileDrv = Mid(lstDrvBck.Drive, 1, ind)
        If Left$(DataFileDir, 1) = "\" Then
            shareEnds = 1 + InStr(Mid$(DataFileDir, 2), "\")
            shareEnds = shareEnds + InStr(Mid$(DataFileDir, shareEnds + 1), "\")
            BackFileDir = Mid$(DataFileDir, shareEnds + 1)
        Else
            BackFileDir = DataFileDir
        End If
        txtBck.text = BackFileDrv & "\" & BackFileDir
        workingSamIndex.BackupFileDir = txtBck
    End If
End Sub

Private Sub lvwSAMRegistry_click()
    Dim i As Integer, selectedindex As Integer
    If lvwSAMRegistry.ListItems.Count = 0 Then Exit Sub
    If lvwSAMRegistry.SelectedItem.Index > 0 Then
        Set workingSamIndex = Nothing
        Set workingSamIndex = SampleIndexRegistry.Item(lvwSAMRegistry.SelectedItem.Index)
        refreshFields
    End If
End Sub

Private Sub lvwSAMRegistry_mousedown(Button As Integer, _
      Shift As Integer, X As Single, Y As Single)
    Dim i As Integer, selectedindex As Integer
    If Button = vbRightButton Then
    End If
End Sub

Private Sub optSAMSetDemag_Click(Index As Integer)
    ' If the "Other" Option is selected then enable the text box
    Set workingSamIndex.measurementSteps = Nothing
    Set workingSamIndex.measurementSteps = New RockmagSteps
    If optSAMSetDemag(1).value = True Then
        txtSAMSetDemagLevel.Visible = False
        cmdRMLevel.Visible = True
        chkMeasureSusceptibility.Enabled = False
        Load frmRockmagRoutine
    ElseIf optSAMSetDemag(5).value = True Then
        txtSAMSetDemagLevel.Visible = False
        cmdRMLevel.Visible = True
        chkMeasureSusceptibility.Enabled = True
        Load frmRockmagRoutine
    Else
        chkMeasureSusceptibility.Enabled = True
        cmdRMLevel.Visible = False
        txtSAMSetDemagLevel.Visible = True
        workingSamIndex.measurementSteps.Add optSAMSetDemag(Index).Caption, val(txtSAMSetDemagLevel)
        Unload frmRockmagRoutine
        Unload frmRockmagRoutine
    End If
    If optSAMSetDemag(4).value = True Then
        txtSAMSetDemag.Enabled = True
    Else
        txtSAMSetDemag.Enabled = False
    End If
End Sub

Private Sub ReadSamInfo()
    ' This procedure updates the 'Info' frame with data from the
    ' sample header file.
    txtSampDesc.text = workingSamIndex.locality
    lblLoc(0).Caption = workingSamIndex.siteLat
    lblLoc(1).Caption = workingSamIndex.siteLong
    lblLoc(2).Caption = workingSamIndex.magDec
    lblSampNum.Caption = workingSamIndex.sampleSet.Count
End Sub

Private Sub refreshFields()
    With workingSamIndex
        txtDir.text = .filedir
        txtBck.text = .BackupFileDir
        cmbSampCode.text = .SampleCode
        ReadSamInfo
        If LenB(.BackupFileDir) <> 0 Then
            chkBck.value = 1
        Else
            chkBck.value = 0
        End If
        cmbSampCode.text = .SampleCode
        If .doBoth Then
            chkSAMdoDown.value = 1
            If .doUp Then
                chkSAMdoUp.value = 1
                chkSAMalreadyDoneUp.value = 0
            Else
                chkSAMdoUp.value = 0
                chkSAMalreadyDoneUp.value = 1
            End If
        Else
            If .doUp Then
                chkSAMdoUp.value = 1
                chkSAMdoDown.value = 0
                chkSAMalreadyDoneUp.value = 0
            Else
                chkSAMdoUp.value = 0
                chkSAMdoDown.value = 1
                chkSAMalreadyDoneUp.value = 0
            End If
        End If
        If .avgSteps > 0 Then
            txtSAMaveragesteps.text = str$(.avgSteps)
        Else
            txtSAMaveragesteps.text = "1"
        End If
        txtSAMSetDemag.text = vbNullString
        If .measurementSteps.Count > 1 Then
            optSAMSetDemag(5) = True
            Set frmRockmagRoutine.rmStepList = Nothing
            Set frmRockmagRoutine.rmStepList = .measurementSteps
        Else
            Select Case Left$(.curDemag, 2)
                Case "TT":
                    optSAMSetDemag(2) = True
                    txtSAMSetDemagLevel = Mid$(.curDemag, 3)
                Case "NR":
                    optSAMSetDemag(0) = True
                Case "MW":
                    optSAMSetDemag(3) = True
                    txtSAMSetDemagLevel = Mid$(.curDemag, 3)
                Case "AF":
                    optSAMSetDemag(1) = True
                    'frmRockmagRoutine.SetActiveFile targetid
                    'frmRockmagRoutine.InitializeFromRegistry
                Case vbNullString:
                    optSAMSetDemag(0) = True
                    txtSAMSetDemagLevel = vbNullString
                Case Else
                    optSAMSetDemag(4) = True
                    txtSAMSetDemag.text = .curDemag
            End Select
        End If
    End With
End Sub

Public Sub refreshSAMRegistryDisplay()
    Dim i As Integer
    Dim curItem As ListItem
    lvwSAMRegistry.ListItems.Clear
    
    If modConfig.UseXYTableAPS Then
    SampleTableTypeLabel.Caption = "XY Table"
    Else
    SampleTableTypeLabel.Caption = "Chain Drive"
    End If
    
    With SampleIndexRegistry
        If .Count = 0 Then Exit Sub
        For i = 1 To .Count
            With .Item(i)
                'watch for duplicates
                On Error GoTo oops
                Set curItem = lvwSAMRegistry.ListItems.Add
                curItem.text = .SampleCode
                curItem.SubItems(1) = .curDemag
                If .doUp Then curItem.SubItems(2) = "Y" Else curItem.SubItems(2) = "N"
                If .doBoth Then curItem.SubItems(3) = "Y" Else curItem.SubItems(3) = "N"
                curItem.SubItems(4) = .avgSteps
                curItem.SubItems(5) = .filename
            End With
        Next i
    End With
    On Error GoTo 0
    Exit Sub
oops:
    Select Case Err.number
        Case 35602 'key not unique
            MsgBox lvwSAMRegistry.ListItems.Item(1).key
        Case Else
            MsgBox Err.number & ": " & Err.Description
    End Select
End Sub

Private Sub txtDir_Change()
    Dim fname As String                 ' Full path name of sample file
    Dim ind As Integer                  ' Index of first character of extension
    Dim gotFile As Boolean              ' Does such a sample exist?
    Dim nextdir As String               ' Next directory to check for samples
    Dim MyPath, MyName As String
    Dim DirList As String
    On Error GoTo ErrorHandler
    Call ClearCmbSampCode               ' Clear the old combo box when we
                                        ' change directories
    If LenB(txtDir.text) = 0 Then Exit Sub   ' Exit if we don't have a directory
    MyPath = txtDir.text
    ' Make sure we have a valid directory first
    If LenB(dir$(MyPath & "\", vbDirectory)) = 0 Then GoTo OutSub
    ' Grab all directory names
    MyName = dir$(MyPath & "\", vbDirectory)     ' Retrieve the first entry in the dir
    Do While LenB(MyName) >= 1                 ' Start loop
        ' Ignore the current directory and the encompassing directory.
        If MyName <> "." And MyName <> ".." Then
            GetAttr (MyPath & MyName & "\" & MyName & ".sam")
            ' Add the item if we find a good sample header
            cmbSampCode.AddItem (MyName)
        End If
BadFile:
        MyName = dir$                      ' Get next entry.
    Loop
OutSub:
    If cmbSampCode.ListCount > 0 Then
        cmbSampCode.text = cmbSampCode.List(0)
        cmbSampCode.SetFocus
    End If
    Call cmbSampCode_Change
    Exit Sub
ErrorHandler:
    Select Case Err.number
        Case 5                          ' Accessing file as directory
            Resume BadFile
        Case 53                     ' File not found
            Resume BadFile
        Case 76
            Resume BadFile          ' File not found
        Case Else
            MsgBox ("Error " & Err.number & _
                " occurred in frmMagnetometerControl!txtDir_Change()." & vbCr & _
                Err.Description)
    End Select
End Sub

Private Sub updateBackupDir()
    If chkBck.value = 0 Then
        workingSamIndex.BackupFileDir = vbNullString
    Else
        workingSamIndex.BackupFileDir = txtBck
    End If
End Sub

Private Sub UpdateDoUpDoBoth()
    Dim filedoboth As Boolean, filedoup As Boolean
    If ((chkSAMalreadyDoneUp.value Or chkSAMdoUp.value) And chkSAMdoDown.value) Then
        filedoboth = True
    Else
        filedoboth = False
    End If
    filedoup = (chkSAMdoUp.value = Checked)
    workingSamIndex.doUp = filedoup
    workingSamIndex.doBoth = filedoboth
End Sub

Private Sub updateMeasurementSteps()
    Dim i As Integer

    MeasSusc = False
    
    Set workingSamIndex.measurementSteps = Nothing
    If optSAMSetDemag(1).value = True Then
        workingSamIndex.RockmagMode = True ' False (March 2008 L Carporzen) Always write the RMG file
        Set workingSamIndex.measurementSteps = frmRockmagRoutine.rmStepList
        UpdateMeasureSusceptibility
    ElseIf optSAMSetDemag(5).value = True Then
        workingSamIndex.RockmagMode = True
        Set workingSamIndex.measurementSteps = frmRockmagRoutine.rmStepList
        
        'Check to see if any of the RockMag steps need their susceptibility to be measured
        For i = 1 To workingSamIndex.measurementSteps.Count
        
            With workingSamIndex.measurementSteps(i)
            
                If (.MeasureSusceptibility = True Or _
                    MeasSusc = True) And EnableSusceptibility _
                Then
                
                    MeasSusc = True
                    
                    Exit For
                    
                ElseIf Not EnableSusceptibility Then
                
                    MeasSusc = False
                    .MeasureSusceptibility = False
                            
                End If
                
            End With
            
        Next i
        
        If chkSAMdoUp And workingSamIndex.measurementSteps.Count > 1 Then chkSAMdoDown = False
    ElseIf optSAMSetDemag(4).value = True Then
        Set workingSamIndex.measurementSteps = New RockmagSteps
        workingSamIndex.measurementSteps.Add txtSAMSetDemag, _
                                             val(txtSAMSetDemagLevel), _
                                             MeasureSusceptibility:= _
                                                (chkMeasureSusceptibility = Checked) And _
                                                EnableSusceptibility
        workingSamIndex.RockmagMode = True ' False (March 2008 L Carporzen) Always write the RMG file
    Else
        Set workingSamIndex.measurementSteps = New RockmagSteps
        For i = 0 To 5
            ' Determine which option button is selected
            If optSAMSetDemag(i).value = True Then Exit For
        Next i
        workingSamIndex.RockmagMode = True ' False (March 2008 L Carporzen) Always write the RMG file
        workingSamIndex.measurementSteps.Add optSAMSetDemag(i).Caption, _
                                             val(txtSAMSetDemagLevel), _
                                             MeasureSusceptibility:= _
                                                (chkMeasureSusceptibility.value = Checked) And _
                                                EnableSusceptibility
    End If
    workingSamIndex.avgSteps = val(txtSAMaveragesteps)
    
    'Do final check to see if the susceptibility needs to be measured
    If chkMeasureSusceptibility.value = Checked And EnableSusceptibility Then
    
        MeasSusc = True
        
    End If
    
    'Now, overwrite the Sample Holder Sample Index in the global SampleIndexRegistrations object
    SampleIndexRegistry.MakeSampleHolder MeasSusc
    
End Sub

Private Sub UpdateMeasureSusceptibility()
    Dim RMStep As RockmagStep
    For Each RMStep In workingSamIndex.measurementSteps
        RMStep.MeasureSusceptibility = (chkMeasureSusceptibility.value = Checked) And EnableSusceptibility
    Next RMStep
    
    If chkMeasureSusceptibility.value = Checked And EnableSusceptibility Then
    
        MeasSusc = True
        
    End If
    
End Sub

