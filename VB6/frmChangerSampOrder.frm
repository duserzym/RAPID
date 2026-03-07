VERSION 5.00
Begin VB.Form frmChangerSampOrder 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sample Settings"
   ClientHeight    =   3210
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9315
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   13.375
   ScaleMode       =   4  'Character
   ScaleWidth      =   77.625
   ShowInTaskbar   =   0   'False
   Begin VB.Frame FrameMeasHolder 
      Caption         =   "Measure Holder"
      Height          =   615
      Left            =   5640
      TabIndex        =   22
      Top             =   2040
      Width           =   3495
      Begin VB.TextBox txtBetweenHolder 
         Height          =   285
         Left            =   2040
         TabIndex        =   24
         Text            =   "10"
         Top             =   240
         Width           =   615
      End
      Begin VB.CheckBox chkMeasHolder 
         Caption         =   "Measure holder every"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.Label Label3 
         Caption         =   "samples"
         Height          =   255
         Left            =   2760
         TabIndex        =   25
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "AF Holder"
      Height          =   735
      Left            =   6360
      TabIndex        =   20
      Top             =   1200
      Width           =   2772
      Begin VB.CheckBox chkAFHolder 
         Caption         =   "AF Holder before measuring"
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   324
      Left            =   6720
      TabIndex        =   19
      Top             =   2760
      Width           =   732
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Send list to command queue"
      Enabled         =   0   'False
      Height          =   324
      Left            =   3720
      TabIndex        =   18
      Top             =   2760
      Width           =   2652
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "&View new sample list"
      Default         =   -1  'True
      Height          =   324
      Left            =   1680
      TabIndex        =   17
      Top             =   2760
      Width           =   1812
   End
   Begin VB.Frame FrameRepHolder 
      Caption         =   "Multiple holder measurements"
      Height          =   975
      Left            =   6360
      TabIndex        =   14
      Top             =   120
      Width           =   2772
      Begin VB.OptionButton optRepeatHolder 
         Caption         =   "Skip            (strong samples)"
         Height          =   612
         Index           =   1
         Left            =   1440
         TabIndex        =   16
         Top             =   240
         Width           =   1092
      End
      Begin VB.OptionButton optRepeatHolder 
         Caption         =   "Repeat           (weak samples)"
         Height          =   612
         Index           =   0
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Value           =   -1  'True
         Width           =   1332
      End
   End
   Begin VB.Frame FrameFinalPos 
      Caption         =   "Final position"
      Height          =   975
      Left            =   3360
      TabIndex        =   11
      Top             =   960
      Width           =   2772
      Begin VB.OptionButton optReturn 
         Caption         =   "Leave at end"
         Height          =   375
         Index           =   1
         Left            =   1440
         TabIndex        =   13
         Top             =   360
         Width           =   1212
      End
      Begin VB.OptionButton optReturn 
         Caption         =   "Return to start"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Value           =   -1  'True
         Width           =   1212
      End
   End
   Begin VB.Frame FrameReloadPos 
      Caption         =   "Reload position"
      Height          =   732
      Left            =   3360
      TabIndex        =   8
      Top             =   120
      Width           =   2772
      Begin VB.OptionButton optLoadReturn 
         Caption         =   "Leave at end"
         Height          =   372
         Index           =   1
         Left            =   1440
         TabIndex        =   10
         Top             =   240
         Width           =   1212
      End
      Begin VB.OptionButton optLoadReturn 
         Caption         =   "Return to start"
         Height          =   372
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Value           =   -1  'True
         Width           =   1332
      End
   End
   Begin VB.ComboBox lstSAMFile 
      Height          =   288
      Left            =   1200
      TabIndex        =   6
      Text            =   "Combo1"
      Top             =   720
      Width           =   1452
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add to list"
      Height          =   324
      Left            =   840
      TabIndex        =   5
      Top             =   2040
      Width           =   1212
   End
   Begin VB.TextBox txtPos 
      Height          =   300
      Left            =   2160
      TabIndex        =   0
      Text            =   "199"
      Top             =   228
      Width           =   492
   End
   Begin VB.Frame Frame1LoadOrder 
      Caption         =   "Load Order"
      Height          =   735
      Left            =   240
      TabIndex        =   3
      Top             =   1200
      Width           =   2652
      Begin VB.OptionButton optOrder 
         Caption         =   "&Descending"
         Height          =   255
         Index           =   1
         Left            =   1320
         TabIndex        =   2
         Top             =   300
         Width           =   1215
      End
      Begin VB.OptionButton optOrder 
         Caption         =   "&Ascending"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   300
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.Label Label2 
      Caption         =   "From file:"
      Height          =   252
      Left            =   240
      TabIndex        =   7
      Top             =   720
      Width           =   1692
   End
   Begin VB.Label Label1 
      Caption         =   "Position of first sample:"
      Height          =   252
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   1692
   End
End
Attribute VB_Name = "frmChangerSampOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const AllSamples = "[All]"
Const ChangerIsFullString = "Full"
Public newChangerList As frmChanger
Public AdditionalHolderMeasurements As Boolean

Private Sub cmdAdd_Click()
    Dim firstpos As Double
    
    If txtPos.text = ChangerIsFullString Then
    
        If UseXYTableAPS Then
        
            'Tell user that the xy stage tray is full
            'And prompt if they'd like the stage sample manifest to be cleared
            Dim f As frmDialog
            
            Set f = New frmDialog
            
            Dim user_resp As VbMsgBoxResult
            
            user_resp = f.DialogBox("The XY tray is fully loaded with samples.  There is no space for additional samples.", _
                                    "XY tray full!", _
                                    2, _
                                    "OK", _
                                    "Clear XY tray")
                                    
            If user_resp = vbNo Then
                
                On Error Resume Next
                newChangerList.Clear
                On Error GoTo 0
                
                txtPos.text = "1"
                
            End If
            
            Exit Sub
            
        Else
        
            CorrectFirstPos
            
        End If
            
    End If
            
    
    cmdOK.Enabled = True ' (September 2007 L Carporzen) Avoid running the sample changer without filling the charger table
    firstpos = val(txtPos.text)
    If Changer_ValidStart(firstpos) Then
        ' We have a valid first position
        ' Modify fields on frmChanger
        Dim next_pos As Integer
        
        next_pos = OrderSamplesFromRegistry(selectedFileId, val(txtPos))
        
        If UseXYTableAPS And next_pos > SlotMax Then
        
            txtPos.text = ChangerIsFullString
        
        Else
        
            txtPos.text = Str(next_pos)
            
        End If
    Else
        CorrectFirstPos
    End If
End Sub

Private Sub cmdCancel_Click()
    frmMagnetometerControl.cmdChangerOK.Enabled = False
    Unload Me
End Sub

Public Sub cmdOK_Click()
    Dim cursamp As SampleIdentifier
    Dim i As Integer
    Dim allgood As Boolean
    Dim howsthat As VbMsgBoxResult
    
    If UseXYTableAPS Then
        If Not ((frmDCMotors.CheckInternalStatus(MotorChanger, 5) = 0) And (frmDCMotors.CheckInternalStatus(MotorChangerY, 6) = 0)) Then
        MotorXYTable_CenterReset
        End If
        If chkMeasHolder.Value = Checked Then
            SamplesBetweenHolder = Int(val(txtBetweenHolder.text))
        Else
            SamplesBetweenHolder = 1000 'Never will be used
        End If
    End If
    
    frmMagnetometerControl.cmdChangerOK.Enabled = True ' (September 2007 L Carporzen) New version of the Halt button
    allgood = False
    Me.MousePointer = 11
    For i = SlotMin To SlotMax
        cursamp.Samplename = newChangerList.ChangerSampleName(i)
        cursamp.filename = newChangerList.ChangerFileName(i)
'        If i > 195 Then MsgBox cursamp.samplename
        If SampleIndexRegistry.IsValidSample(cursamp.filename, cursamp.Samplename) And _
           cursamp.Samplename <> MainChanger.ChangerSampleName(i) _
        Then
        
            If MainChanger.isValidSampleIn(i) And Changer_ValidSlot(i) Then
                Me.MousePointer = 0
                howsthat = MsgBox("Replace sample " & MainChanger.ChangerSampleName(i) & " in hole " & _
                    Str$(i) & vbCr & "with sample " & newChangerList.ChangerSampleName(i) & "?", vbYesNoCancel)
                Me.MousePointer = 11
                If howsthat = vbYes Then allgood = True
                If howsthat = vbNo Then
                    newChangerList.emptyHole i
                End If
                If howsthat = vbCancel Then Exit Sub
            Else
                allgood = True
            End If
            If allgood Then MainChanger.IncorporateSample i, cursamp.filename, cursamp.Samplename
        End If
    Next i
    MainChanger.RefreshControls
    MainChanger.optOrder(0) = newChangerList.optOrder(0)
    MainChanger.optLoadReturn(0) = newChangerList.optLoadReturn(0)
    MainChanger.optReturn(0) = newChangerList.optReturn(0)
    MainChanger.optRepeatHolder(0) = newChangerList.optRepeatHolder(0)
    newChangerList.ProcessSamplesToQueue optOrder(0).Value, _
                                         optLoadReturn(0).Value, _
                                         optReturn(0).Value, _
                                         optRepeatHolder(0).Value
                                         
    Me.MousePointer = 0
    If chkAFHolder.Value = Checked Then
            SampleHolder.Parent.measurementSteps(1).StepType = "AFmax"
            SampleHolder.Parent.measurementSteps(1).Level = AfTransMax
    Else
            SampleHolder.Parent.measurementSteps(1).StepType = "NRM"
            SampleHolder.Parent.measurementSteps(1).Level = 0
    End If
    Unload frmChanger
    Unload Me
End Sub

Private Sub cmdPreview_Click()
    cmdOK.Enabled = True ' (September 2007 L Carporzen)  Avoid running the sample changer without filling the charger table
    newChangerList.IsMasterList = False
    newChangerList.RefreshControls
    newChangerList.ZOrder
    newChangerList.Show
End Sub

Private Sub CorrectFirstPos()
    ' This procedure asks the user to give a valid position for
    ' the first sample of the series
    
    If UseXYTableAPS Then
    
        MsgBox "The position of the first sample must be " & vbCr & _
               "between " & SlotMin & " and " & SlotMax & ", and " & _
               "cannot equal to 46 (the XY stage hole).", vbOK, "Invalid Position"
    Else
    
        MsgBox "The position of the first sample must be " & vbCr & _
               "between " & SlotMin & " and " & SlotMax - 1 & ", and " & _
               "cannot be one of the holes.", vbOK, "Invalid Position"
    
    End If
    
    
    txtPos.SelStart = 0
    txtPos.SelLength = Len(txtPos.text)
    txtPos.SetFocus
End Sub

Private Sub fillLstSAMFile()
    Dim i As SampleIndexRegistration
    Do While lstSAMFile.ListCount > 0
        lstSAMFile.RemoveItem 0
    Loop
    lstSAMFile.List(0) = AllSamples
    If SampleIndexRegistry Is Nothing Then Exit Sub
    For Each i In SampleIndexRegistry
        lstSAMFile.AddItem i.filename
    Next i
    lstSAMFile.ListIndex = 0
End Sub

Private Sub Form_Load()

    '(March 11, 2011 - I Hilburn)
    'Update all the form's icons with the program Ico file
    SetFormIcon Me

    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
    ' Initialize defaults.
    If UseXYTableAPS Then
        txtPos.text = Str(SlotMin)
        optOrder(1).Value = False
    Else
        txtPos.text = Str(SlotMax - 1)
        optOrder(1).Value = True
    End If
    cmdOK.Default = True
    
    fillLstSAMFile
    Set newChangerList = Nothing
    Set newChangerList = New frmChanger
    newChangerList.IsMasterList = False
    newChangerList.Caption = "Hole Sample List - Preview Addition"
    newChangerList.RefreshControls
    
    Set newChangerList = MainChanger
    
    If UseXYTableAPS Then
        FrameMeasHolder.Visible = True
    Else
        FrameMeasHolder.Visible = False
    End If
    
    
    
End Sub

Private Sub Form_Resize()
    Me.Height = 3885
    Me.Width = 9405
End Sub

Private Function isAscend() As Boolean
    isAscend = optOrder(0).Value
End Function

Public Function OrderSamplesFromRegistry(fileid As String, firstpos As Integer) As Integer
    ' This procedure orders samples in the order described by
    ' isAscend using firstpos as the position of the first sample
    ' if isAscend is true, then we do it in ascending order,
    '           otherwise, we do so in descending order
    ' This returns the next hole in line.
    ' Use fileid = AllSamples to load samples from all files.
    Dim i As Integer
    Dim slotpos As Integer
    Dim fileItem As SampleIndexRegistration
    newChangerList.setOrder (isAscend)
    slotpos = firstpos
    
    If fileid = AllSamples Then
        For Each fileItem In SampleIndexRegistry
        fileid = fileItem.filename
        With fileItem.sampleSet
        
        If UseXYTableAPS Then
    
            If firstpos + .Count > SlotMax Then
            
                'Prompt user that there isn't enough space and some
                'samples will be left out of the measurement queue
                Dim user_resp As VbMsgBoxResult
                
                user_resp = MsgBox("Insufficient space on the XY stage to fit the remaining samples from " & _
                                    fileItem.SampleCode & "." & _
                                   vbCrLf & vbCrLf & "The last " & Trim(Str(.Count + firstpos - SlotMax - 1)) & _
                                   " sample(s) will be left out of the current automated sample run." & _
                                   vbCrLf & vbCrLf & "Click 'Yes' if this is okay, or click 'No' if you'd like to exclude all " & _
                                   "of the samples in " & fileItem.SampleCode & " from the current automated " & _
                                   "sample run.", vbYesNo, "XY Stage full!")
                                   
                If user_resp <> vbYes Then
                
                    'exit the larger for loop
                    Exit For
                
                End If
                                   
            
            End If
        
        End If
        
        
        For i = 1 To .Count
            If slotpos >= SlotMax And _
               Not UseXYTableAPS Then
                ' The slot must always be within range
                slotpos = slotpos Mod SlotMax
                
            ElseIf UseXYTableAPS And _
                   slotpos > SlotMax Then
                   
                   'Exit the for loop
                   Exit For
            End If
            While slotpos < SlotMin
                ' Increment slot until it is within range
                slotpos = slotpos + SlotMax
            Wend
            ' The (slotpos)'th slot has this sample
            If LenB(.Item(i).Samplename) > 0 Then
                newChangerList.IncorporateSample slotpos, fileid, .Item(i).Samplename
            End If
            ' Move to the next slot for the next sample
            If isAscend Then
                slotpos = slotpos + 1                   ' Increment
            Else
                slotpos = slotpos - 1                   ' Decrement
            End If
            slotpos = SkipHoles(slotpos, isAscend)
        Next i
        End With
        Next fileItem
    ElseIf SampleIndexRegistry.IsValidFile(fileid) Then
        With SampleIndexRegistry(fileid).sampleSet
        For i = 1 To .Count
            If slotpos >= SlotMax And _
               Not UseXYTableAPS Then
                ' The slot must always be within range
                slotpos = slotpos Mod SlotMax
                
            ElseIf UseXYTableAPS And _
                   slotpos > SlotMax Then
                   
                'Exit the for loop
                Exit For
            End If
            While slotpos < SlotMin
                ' Increment slot until it is within range
                slotpos = slotpos + SlotMax
            Wend
            ' The (slotpos)'th slot has this sample
            If LenB(.Item(i).Samplename) > 0 Then
                newChangerList.IncorporateSample slotpos, fileid, .Item(i).Samplename
            End If
            ' Move to the next slot for the next sample
            If isAscend Then
                slotpos = slotpos + 1                   ' Increment
            Else
                slotpos = slotpos - 1                   ' Decrement
            End If
            slotpos = SkipHoles(slotpos, isAscend)
        Next i
        End With
    End If
    
    OrderSamplesFromRegistry = slotpos
End Function

Private Function selectedFileId() As String
    With lstSAMFile
        selectedFileId = .List(.ListIndex)
    End With
End Function

Public Function SkipHoles(slotpos As Integer, asc As Boolean) As Integer
    ' This function returns slotpos if the slot is valid.
    ' Otherwise, it returns the number of the next valid slot,
    ' incrementing if asc is true, decrementing if not
    
    If modConfig.UseXYTableAPS Then
    
        If slotpos = HoleSlotNum Then
            If asc Then
                slotpos = slotpos + 1               ' Increment
            Else
                slotpos = slotpos - 1               ' Decrement
            End If
        End If
        
    ElseIf slotpos Mod HoleSlotNum = 0 Then
        
        If asc Then
            slotpos = slotpos + 1               ' Increment
        Else
            slotpos = slotpos - 1               ' Decrement
        End If
        
    End If
    SkipHoles = slotpos
End Function

