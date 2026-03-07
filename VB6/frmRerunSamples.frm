VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmRerunSamples 
   Caption         =   "Rerun Samples"
   ClientHeight    =   5610
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6885
   Icon            =   "frmRerunSamples.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5610
   ScaleWidth      =   6885
   Begin VB.CommandButton btnClose 
      Caption         =   "Close"
      CausesValidation=   0   'False
      Height          =   372
      Left            =   120
      TabIndex        =   12
      Top             =   5040
      Width           =   1212
   End
   Begin VB.CheckBox chkSAMalreadyDoneUp 
      Caption         =   "Use up already measured"
      Height          =   252
      Left            =   2520
      TabIndex        =   10
      Top             =   1440
      Width           =   2295
   End
   Begin VB.CheckBox chkSAMdoUp 
      Caption         =   "Up"
      Height          =   252
      Left            =   2520
      TabIndex        =   9
      Top             =   1200
      Value           =   1  'Checked
      Width           =   612
   End
   Begin VB.CheckBox chkSAMdoDown 
      Caption         =   "Down"
      Height          =   252
      Left            =   3120
      TabIndex        =   8
      Top             =   1200
      Value           =   1  'Checked
      Width           =   732
   End
   Begin VB.CommandButton cmdRescan 
      Caption         =   "Rescan"
      Height          =   372
      Left            =   3600
      TabIndex        =   7
      Top             =   600
      Width           =   1332
   End
   Begin VB.TextBox txtErrorAngle 
      Height          =   285
      Left            =   2400
      TabIndex        =   6
      Text            =   "15"
      Top             =   600
      Width           =   735
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Rerun Samples"
      CausesValidation=   0   'False
      Height          =   372
      Left            =   5400
      TabIndex        =   4
      Top             =   5040
      Width           =   1212
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Enabled         =   0   'False
      Height          =   372
      Left            =   2280
      TabIndex        =   3
      Top             =   5040
      Width           =   1332
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   372
      Left            =   3840
      TabIndex        =   2
      Top             =   5040
      Width           =   1332
   End
   Begin ComctlLib.ListView lvwSamples 
      Height          =   2655
      Left            =   120
      TabIndex        =   1
      Top             =   1920
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   4683
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label Label16 
      Caption         =   "Directions to Measure:"
      Height          =   255
      Left            =   600
      TabIndex        =   11
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Maximum error angle:"
      Height          =   255
      Left            =   600
      TabIndex        =   5
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "The following samples had large error angles. Click ""Ok"" to rerun them."
      Height          =   255
      Left            =   600
      TabIndex        =   0
      Top             =   120
      Width           =   5415
   End
End
Attribute VB_Name = "frmRerunSamples"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rerunSamples As Collection
Dim thresholdAngle As Double

Public Sub addSample(hole As Integer)
    If MainChanger.isValidSampleIn(hole) Then rerunSamples.Add hole
    refreshListDisplay
End Sub

Private Sub btnClose_Click()
    Me.Hide
End Sub

Private Sub chkSAMalreadyDoneUp_Click()
    If chkSAMalreadyDoneUp.value Then chkSAMdoUp.value = False
End Sub

Public Sub Clear()
    Set rerunSamples = Nothing
    Set rerunSamples = New Collection
    refreshListDisplay
End Sub

Private Sub cmdClear_Click()
    Clear
End Sub

Private Sub cmdDelete_Click()
    Dim targetItem As ListItem
    If lvwSamples.SelectedItem.index > 0 Then
        For Each targetItem In lvwSamples.ListItems
            If targetItem.Selected Then
                rerunSamples.Remove targetItem.index
            End If
        Next targetItem
    Else
        cmdDelete.Enabled = False
    End If
    refreshListDisplay
End Sub

Private Sub cmdOK_Click()
    Me.Hide
    If rerunSamples.Count > 0 Then
        ProcessSamplesToQueue
        SampQueue.Execute
    End If
End Sub

Private Sub cmdRescan_Click()
    thresholdAngle = val(txtErrorAngle)
    scanForLargeErrorAngles
End Sub

Private Sub Form_Activate()
    refreshListDisplay
    setCheckBoxesFromSampleIndexRegistryForm
End Sub

Private Sub Form_Load()
    
    Dim colX As ColumnHeader ' Declare variable.
    
    '(March 11, 2011 - I Hilburn)
    'Update all the form's icons with the program Ico file
    SetFormIcon Me
    
    Set rerunSamples = New Collection
    Set colX = lvwSamples.ColumnHeaders.Add(1)
    colX.text = "Hole"
    colX.Width = Me.TextWidth(colX.text & " ")
    Set colX = lvwSamples.ColumnHeaders.Add(2)
    colX.text = "Sample"
    colX.Width = Me.TextWidth(colX.text & " ")
    Set colX = lvwSamples.ColumnHeaders.Add(3)
    colX.text = "Circular Std. Dev."
    colX.Width = Me.TextWidth(colX.text & " ")
    Set colX = lvwSamples.ColumnHeaders.Add(4)
    colX.text = "Moment"
    colX.Width = Me.TextWidth(colX.text & " ")
    Set colX = lvwSamples.ColumnHeaders.Add(5)
    colX.text = "Up/Down"
    colX.Width = Me.TextWidth(colX.text & " ")
    Set colX = lvwSamples.ColumnHeaders.Add(6)
    colX.text = "File"
    colX.Width = Me.TextWidth(colX.text & "XXXX ")
    thresholdAngle = RemeasureCSDThreshold
    
End Sub

Private Sub Form_Resize()
    Me.Width = 7005
    cmdDelete.Top = Me.Height - 935
    cmdClear.Top = Me.Height - 935
    cmdOk.Top = Me.Height - 935
    lvwSamples.Height = Me.Height - 3350
End Sub

Private Sub form_show()
    scanForLargeErrorAngles
    
End Sub

Private Sub setCheckBoxesFromSampleIndexRegistryForm()

    Me.chkSAMdoUp.value = frmSampleIndexRegistry.chkSAMdoUp.value
    Me.chkSAMdoDown.value = frmSampleIndexRegistry.chkSAMdoDown.value
    Me.chkSAMalreadyDoneUp.value = frmSampleIndexRegistry.chkSAMalreadyDoneUp.value

End Sub

Private Sub lvwsamples_click()
    On Error GoTo fin
    If LenB(lvwSamples.SelectedItem.text) > 0 Then
        cmdDelete.Enabled = True
    Else
        cmdDelete.Enabled = False
    End If
fin:
End Sub

Public Sub ProcessSamplesToQueue()
    Dim i As Integer
    Dim newChangerList As frmChanger
    Dim hole As Integer
    Dim workingFilename As String
    Set newChangerList = New frmChanger
    Load newChangerList
    With rerunSamples
    If .Count = 0 Then Exit Sub
    For i = 1 To .Count
        hole = .Item(i)
        If MainChanger.isValidSampleIn(hole) Then
            workingFilename = MainChanger.ChangerFileName(hole)
            
            newChangerList.IncorporateSample hole, _
                                             MainChanger.ChangerFileName(hole), _
                                             MainChanger.ChangerSampleName(hole)
                                             
            UpdateDoUpDoBoth (workingFilename)
            
        End If
    Next i
    End With
    
    If UseXYTableAPS Then
    
        newChangerList.ProcessSamplesToQueue MainChanger.optOrder(0).value, _
                                         True, _
                                         MainChanger.optReturn(0).value, _
                                         MainChanger.optRepeatHolder(0).value
                                         
        MainChanger.optLoadReturn(0).value = True
    
    Else
    
        newChangerList.ProcessSamplesToQueue MainChanger.optOrder(0).value, _
                                         MainChanger.optLoadReturn(0).value, _
                                         MainChanger.optReturn(0).value, _
                                         MainChanger.optRepeatHolder(0).value
    
    End If
    
    
    
    Set newChangerList = Nothing
    
    Clear
    
End Sub

Public Sub refreshListDisplay()
    Dim i As Integer
    Dim curItem As ListItem
    txtErrorAngle = Format$(thresholdAngle, "##.0")
    lvwSamples.ListItems.Clear
    With rerunSamples
        If .Count = 0 Then Exit Sub
        For i = 1 To .Count
            Set curItem = lvwSamples.ListItems.Add(i)
            curItem.text = .Item(i)
            With MainChanger.ChangerSample(.Item(i))
                curItem.SubItems(1) = .Samplename
                curItem.SubItems(2) = Format$(.ErrorAngle, "##.00")
                curItem.SubItems(3) = Format$(.Moment, "0.000E-")
                If .Parent.doBoth Then curItem.SubItems(4) = Format$(.UpDownRatio, "0.00") Else curItem.SubItems(4) = vbNullString
                curItem.SubItems(5) = .IndexFile
            End With
        Next i
    End With
End Sub

Public Function scanForLargeErrorAngles(Optional ByVal minangle As Double = 0) As Integer
    Dim i As Integer
    Clear
    If minangle = 0 Then minangle = thresholdAngle
    For i = SlotMin To SlotMax
        If MainChanger.isValidSampleIn(i) Then
            With MainChanger.ChangerSample(i)
                If .ErrorAngle > minangle Then addSample i
            End With
        End If
    Next i
    refreshListDisplay
    scanForLargeErrorAngles = rerunSamples.Count
End Function

Private Sub UpdateDoUpDoBoth(filename As String)
    Dim filedoboth As Boolean, filedoup As Boolean
    If ((chkSAMalreadyDoneUp.value Or chkSAMdoUp.value) And chkSAMdoDown.value) Then
        filedoboth = True
    Else
        filedoboth = False
    End If
    filedoup = (chkSAMdoUp.value = Checked)
    SampleIndexRegistry(filename).doUp = filedoup
    SampleIndexRegistry(filename).doBoth = filedoboth
End Sub

