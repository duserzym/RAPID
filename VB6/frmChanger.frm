VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmChanger 
   AutoRedraw      =   -1  'True
   Caption         =   "Hole Sample List"
   ClientHeight    =   6150
   ClientLeft      =   135
   ClientTop       =   420
   ClientWidth     =   10305
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6084.89
   ScaleMode       =   0  'User
   ScaleWidth      =   10752.43
   Visible         =   0   'False
   Begin VB.Frame FrameOrder 
      Caption         =   "Sample order"
      Height          =   732
      Left            =   7320
      TabIndex        =   14
      Top             =   1800
      Width           =   2772
      Begin VB.OptionButton optOrder 
         Caption         =   "Descending"
         Height          =   372
         Index           =   1
         Left            =   1440
         TabIndex        =   16
         Top             =   240
         Width           =   1212
      End
      Begin VB.OptionButton optOrder 
         Caption         =   "Ascending"
         Height          =   372
         Index           =   0
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Value           =   -1  'True
         Width           =   1332
      End
   End
   Begin VB.Frame FrameReloadPos 
      Caption         =   "Reload position"
      Height          =   732
      Left            =   7320
      TabIndex        =   11
      Top             =   2760
      Width           =   2772
      Begin VB.OptionButton optLoadReturn 
         Caption         =   "Return to start"
         Height          =   372
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Value           =   -1  'True
         Width           =   1332
      End
      Begin VB.OptionButton optLoadReturn 
         Caption         =   "Leave at end"
         Height          =   372
         Index           =   1
         Left            =   1440
         TabIndex        =   12
         Top             =   240
         Width           =   1212
      End
   End
   Begin VB.Frame FrameFinalPos 
      Caption         =   "Final position"
      Height          =   975
      Left            =   7320
      TabIndex        =   8
      Top             =   3720
      Width           =   2772
      Begin VB.OptionButton optReturn 
         Caption         =   "Return to start"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Value           =   -1  'True
         Width           =   1212
      End
      Begin VB.OptionButton optReturn 
         Caption         =   "Leave at end"
         Height          =   375
         Index           =   1
         Left            =   1440
         TabIndex        =   9
         Top             =   360
         Width           =   1212
      End
   End
   Begin VB.Frame FrameRepHolder 
      Caption         =   "Multiple holder measurements"
      Height          =   972
      Left            =   7320
      TabIndex        =   5
      Top             =   4920
      Width           =   2772
      Begin VB.OptionButton optRepeatHolder 
         Caption         =   "Repeat           (weak samples)"
         Height          =   612
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Value           =   -1  'True
         Width           =   1332
      End
      Begin VB.OptionButton optRepeatHolder 
         Caption         =   "Skip            (strong samples)"
         Height          =   612
         Index           =   1
         Left            =   1440
         TabIndex        =   6
         Top             =   240
         Width           =   1092
      End
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   315
      Left            =   7440
      TabIndex        =   3
      Top             =   1200
      Width           =   972
   End
   Begin VB.CommandButton cmdSeq 
      Caption         =   "&Sequential Ordering"
      Height          =   495
      Left            =   8880
      TabIndex        =   2
      Top             =   480
      Width           =   972
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
      Height          =   315
      Left            =   7440
      TabIndex        =   1
      Top             =   840
      Width           =   972
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   315
      Left            =   7440
      TabIndex        =   0
      Top             =   480
      Width           =   972
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridSamples 
      Height          =   5772
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   6852
      _ExtentX        =   12091
      _ExtentY        =   10186
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Menu mnuSample 
      Caption         =   "Sample"
      NegotiatePosition=   3  'Right
      Visible         =   0   'False
      Begin VB.Menu mnuSampleDelete 
         Caption         =   "&Delete"
      End
      Begin VB.Menu mnuInsertSample 
         Caption         =   "&Insert Sample"
      End
      Begin VB.Menu mnuSampleInfo 
         Caption         =   "Sample Info"
      End
      Begin VB.Menu mnuDeleteNext 
         Caption         =   "Delete next 9 samples"
      End
      Begin VB.Menu mnuSampleDeleteAndShift 
         Caption         =   "Delete without &Gap"
      End
   End
End
Attribute VB_Name = "frmChanger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Sample changer form.  This is for handling assignments
' of samples to holes.
Option Explicit
Private Chg_isAscend As Boolean    ' Order of changer measurement
Public newSampQueue As SampleCommands
Private ChangerSamples() As SampleIdentifier
Private WhichFilesLoaded() As Integer
Private gridrows As Integer
Private gridcols As Integer
Public IsMasterList As Boolean

Public Function ChangerFileName(ByVal slot As Integer) As String
    If slot < SlotMin Or slot > SlotMax Then Exit Function
    ChangerFileName = ChangerSamples(slot).filename
End Function

Public Function ChangerSample(slot As Integer) As Sample
    On Error GoTo oops
    With ChangerSamples(slot)
        Set ChangerSample = SampleIndexRegistry(.filename).sampleSet(.Samplename)
    End With
    On Error GoTo 0
    Exit Function
oops:
    Set ChangerSample = New Sample
End Function

Public Function ChangerSampleName(ByVal slot As Integer) As String
    If slot < SlotMin Or slot > SlotMax Then Exit Function
    ChangerSampleName = ChangerSamples(slot).Samplename
End Function

Public Sub Clear()
' This procedure clears all of the samples from the
    ' screen, resetting the sample form.
    Dim i As Integer
    For i = SlotMin To SlotMax
        With ChangerSamples(i)
            If IsMasterList Then
                If SampleIndexRegistry.IsValidSample(.filename, .Samplename) Then
                    SampleIndexRegistry(.filename).sampleSet(.Samplename).sampleHole = 0
                End If
            End If
            .filename = vbNullString
            .Samplename = vbNullString
        End With
    Next i
    fillSamplesGrid
    If IsMasterList Then
        SampQueue.Clear
    End If
End Sub

Private Sub cmdClear_Click()
    Clear
End Sub

Private Sub cmdOK_Click()
    ' The settings are acceptable to the user, so continue.
    'SetSequenceVariables
    'If IsMasterList Then
     '   ProcessSamplesToQueue
    'End If
    Me.Hide
End Sub

Public Sub cmdSeq_Click()
    ' This procedure requests a little information about the
    ' load order, and automatically assigns samples to sample
    ' changer holes.
    Set fChangerSampOrder = New frmChangerSampOrder
    Load fChangerSampOrder
    fChangerSampOrder.ZOrder
    fChangerSampOrder.Show
End Sub

Private Sub countFileCalls()
    Dim i As Integer
    ReDim WhichFilesLoaded(SampleIndexRegistry.Count) As Integer
    For i = 1 To SlotMax
        With ChangerSamples(i)
        If LenB(.filename) > 0 Then WhichFilesLoaded(SampleIndexRegistry.Index(.filename)) = WhichFilesLoaded(SampleIndexRegistry.Index(.filename)) + 1
        End With
    Next i
End Sub

Public Sub emptyHole(holeid As Integer)
    If Changer_ValidStart(holeid) And Not Changer_isHole(holeid) Then
        With ChangerSamples(holeid)
            If IsMasterList Then
                If SampleIndexRegistry.IsValidSample(.filename, .Samplename) Then
                    SampleIndexRegistry(.filename).sampleSet(.Samplename).sampleHole = 0
                End If
            End If
            .filename = vbNullString
            .Samplename = vbNullString
        End With
        fillSamplesGrid
    End If
End Sub

Public Sub fillSamplesGrid()
    Dim i As Integer
    Dim j As Integer
    Dim doublegridcols As Double
    gridSamples.FixedCols = 0
    gridSamples.FixedRows = 1
    
    'Added code to fix the y-axis resize issues (resizing height)
    Dim corrected_height
    
    'Store calculated height to local var
    'Under certain use-cases, this value will be negative.
    corrected_height = Me.Height - 1000
        
    'Set minimum height threshold of 10 px
    If corrected_height < 10 Then corrected_height = 10
        
    'Set the corrected height to the Sample Grid control
    gridSamples.Height = corrected_height
    
    If cmdOK.Left < FrameOrder.Left Then
        gridSamples.Width = cmdOK.Left - (2 * gridSamples.Left)
    Else
        'Get the width to use for the Sample Grid from the width of the Order frame
        Dim corrected_width As Double
        corrected_width = FrameOrder.Left - (2 * gridSamples.Left)
        
        'Correct for case where the resulting Sample Grid width is less than zero
        'But, set minimum possible width to 10
        If corrected_width < 10 Then corrected_width = 10
                                                              
        'Set Sample grid width to the corrected width
        gridSamples.Width = corrected_width
    End If
    gridrows = HoleSlotNum * Int(((gridSamples.Height) / (gridSamples.RowHeight(0)) - 2) / HoleSlotNum)
    If gridrows < HoleSlotNum Then gridrows = HoleSlotNum
    doublegridcols = ((SlotMax - SlotMin) + 1) / gridrows
    If doublegridcols <> Int(doublegridcols) Then doublegridcols = doublegridcols + 1
    gridcols = 2 * Int(doublegridcols)
    gridSamples.Rows = gridrows + 1
    gridSamples.Cols = gridcols
    For j = 0 To gridcols - 1
        If j Mod 2 = 0 Then
            gridSamples.ColWidth(j) = Me.TextWidth(Str(SlotMax) & "W")
        Else
            gridSamples.ColWidth(j) = Me.TextWidth("WWWWWWW")
        End If
        For i = 0 To gridrows
            gridSamples.TextMatrix(i, j) = gridContent(i, j)
        Next i
    Next j
End Sub

'-----------------------------------------------------------------------------
'  FirstSlot
'
'  Description:       This function returns the slot of the first sample.
'
'  Revision History:
'      Albert Hsiao   2.24.99     initial revision
'      Albert Hsiao   4.12.99     fixed bugs!
'      Bob Kopp       11.24.03    rewritten for registry
'
Public Function FirstSlot() As Integer
    Dim i As Integer, j As Integer, k As Integer
    FirstSlot = -1
    If SampleIndexRegistry.Count = 0 Then Exit Function
    For i = 1 To SampleIndexRegistry.Count
        With SampleIndexRegistry(i).sampleSet
        If .Count = 0 Then GoTo carryforth
        ' Start looking for the lowest numbered sample
        If Chg_isAscend Then
            For k = 1 To .Count
            
                If UseXYTableAPS Then
                
                    For j = 1 To SlotMax
                        If Changer_isHole(j) Then j = j + 1
                        If ChangerSamples(j).Samplename = .Item(k).Samplename And ChangerSamples(j).filename = SampleIndexRegistry(i).filename Then
                            FirstSlot = j
                            k = .Count
                            Exit For
                        End If
                    Next j
                
                Else
                
                    For j = 1 To (SlotMax - 1)
                        If Changer_isHole(j) Then j = j + 1
                        If ChangerSamples(j).Samplename = .Item(k).Samplename And ChangerSamples(j).filename = SampleIndexRegistry(i).filename Then
                            FirstSlot = j
                            k = .Count
                            Exit For
                        End If
                    Next j
                    
                End If
            
                
            Next k
        Else
            For k = 1 To .Count
            
                If UseXYTableAPS Then
                
                    For j = SlotMax To SlotMin Step -1
                        If Changer_isHole(j) Then j = j - 1
                        If ChangerSamples(j).Samplename = .Item(k).Samplename And ChangerSamples(j).filename = SampleIndexRegistry(i).filename Then
                            FirstSlot = j
                            k = .Count
                            Exit For
                        End If
                    Next j
                        
                Else
                
                    For j = (SlotMax - 1) To SlotMin Step -1
                        If Changer_isHole(j) Then j = j - 1
                        If ChangerSamples(j).Samplename = .Item(k).Samplename And ChangerSamples(j).filename = SampleIndexRegistry(i).filename Then
                            FirstSlot = j
                            k = .Count
                            Exit For
                        End If
                    Next j
                
                End If
            
                
            Next k
        End If
carryforth:
        If FirstSlot <> -1 Then Exit For
        End With
    Next i
End Function

Private Sub Form_Load()
    
    ReDim ChangerSamples(SlotMax) As SampleIdentifier
    IsMasterList = False
    Me.Width = 0.8 * frmProgram.Width
    Me.Height = 0.8 * frmProgram.Height
    Me.Left = 0.03 * frmProgram.Width
    Me.Top = 0.03 * frmProgram.Height
    ' set some field properties
    'Dim i As Integer
    'For i = SLOTMIN To (SLOTMAX - 1)
        'If (i Mod HOLESLOTNUM <> 0) Then
            'txtNum(i).Enabled = True
            'txtNum(i).Font = "Arial"
            'txtNum(i).FontSize = 7
        'End If
    'Next i
    optOrder(1).Value = True
    optLoadReturn(0).Value = True
    optReturn(0).Value = True
    RefreshControls
    Clear             ' clear data on the form
    
    '(March 11, 2011 - I Hilburn)
    'Update all the form's icons with the program Ico file
    SetFormIcon Me
    
End Sub

Private Sub Form_Resize()
    cmdSeq.Left = Me.Width - cmdSeq.Width * (-cmdSeq.Visible) - 200
    cmdOK.Left = cmdSeq.Left - cmdOK.Width - 400
    cmdPrint.Left = cmdSeq.Left - cmdPrint.Width - 400
    cmdClear.Left = cmdSeq.Left - cmdClear.Width - 400
    FrameOrder.Left = Me.Width - (FrameOrder.Width) * (-FrameOrder.Visible)
    FrameReloadPos.Left = Me.Width - (FrameReloadPos.Width) * (-FrameReloadPos.Visible)
    FrameFinalPos.Left = Me.Width - (FrameFinalPos.Width) * (-FrameFinalPos.Visible)
    FrameRepHolder.Left = Me.Width - (FrameRepHolder.Width) * (-FrameRepHolder.Visible)
    fillSamplesGrid
End Sub

Private Sub form_show()
    Form_Resize
End Sub

Public Function GetCurrentChangerPos()
    ' Ask user to specify which sample is currently under the holder
    ' A wrong answer could be disastrous.
    Dim newpos As Double
    Dim rangeStr As String

    If currentPosInitialized Then Exit Function
    newpos = -1
    rangeStr = Str(SlotMin) & "-" & Str(SlotMax)
    While Not (newpos >= SlotMin And newpos <= SlotMax)
    If UseXYTableAPS Then
    newpos = val(InputBox("Which XY cup is current under the quartz glass holder?" & vbCr & "CAREFUL - A wrong answer " & _
        "here could BREAK THE SYSTEM!", "Important!", frmDCMotors.ChangerHole))
    Else
    newpos = val(InputBox("Which sample slot is now under " & _
        "the quartz glass holder " & rangeStr & "?" & vbCr & "CAREFUL - A wrong answer " & _
        "here could BREAK THE SYSTEM!", "Important!", frmDCMotors.ChangerHole))
        End If
    Wend
    frmDCMotors.SetChangerHole newpos
    currentPosInitialized = True
End Function

Private Function gridContent(row As Integer, Column As Integer) As String
    Dim holeLabel As Boolean
    Dim holeid As Integer
    holeid = gridHole(row, Column)
    If (holeid = -1) Then
        If Column Mod 2 = 0 Then
            gridContent = "Hole"
        Else
            gridContent = "Sample"
        End If
    Else
        If holeid = -2 Then
            holeLabel = True
            holeid = gridHole(row, Column + 1)
        End If
        If holeid = -3 Then Exit Function
        If holeLabel Then
            gridContent = Str(holeid)
        Else
            If Changer_isHole(holeid) Then
                gridContent = " ---------- "
            Else
                gridContent = ChangerSamples(holeid).Samplename
            End If
        End If
    End If
End Function

Private Function gridHole(row As Integer, Column As Integer) As Integer
    Dim holeLabel As Boolean
    Dim holeid As Integer
    If (row = 0) Then
        gridHole = -1
    Else
        If Column Mod 2 = 0 Then holeLabel = True
        holeid = (Column \ 2) * gridrows
        holeid = holeid + SlotMin + (row - 1)
        If holeid <= SlotMax And holeid >= SlotMin Then
            If holeLabel Then
                gridHole = -2
            Else
                gridHole = holeid
            End If
        Else
            gridHole = -3
        End If
    End If
End Function

Private Sub gridsamples_mousedown(Button As Integer, _
      Shift As Integer, X As Single, Y As Single)
    Dim holeid As Integer
    'If Button <> vbRightButton Then Exit Sub
    holeid = gridHole(gridSamples.row, gridSamples.Col)
    If Changer_ValidStart(holeid) And Not Changer_isHole(holeid) Then
        If SampleIndexRegistry.IsValidSample(ChangerSamples(holeid).filename, ChangerSamples(holeid).Samplename) And (Not Changer_isHole(holeid)) Then
            mnuSampleInfo.Visible = True
            mnuSampleDelete.Visible = True
            mnuSampleDeleteAndShift.Visible = True
            PopupMenu mnuSample
        Else
            mnuSampleInfo.Visible = False
            mnuSampleDelete.Visible = False
            mnuSampleDeleteAndShift.Visible = False
            PopupMenu mnuSample
        End If
    End If
End Sub

Private Function holeGridPosition(holeNum As Integer, row As Boolean) As Integer
    If row Then
        holeGridPosition = (holeNum - SlotMin) Mod (gridrows) + gridSamples.FixedRows
    Else
        holeGridPosition = ((holeNum - SlotMin) \ gridcols) * 2 + 1
    End If
End Function

Public Sub IncorporateSample(ByVal addHole As Integer, ByVal filename As String, ByVal Samplename As String)
    If LenB(filename) = 0 And LenB(Samplename) = 0 Then Exit Sub
    If Changer_ValidSlot(addHole) Then
        With ChangerSamples(addHole)
            If isValidSampleIn(addHole) Then _
                SampleIndexRegistry(.filename).sampleSet(.Samplename).sampleHole = 0
                
            If SampleIndexRegistry.IsValidSample(filename, Samplename) Then
                .filename = filename
                .Samplename = Samplename
                If IsMasterList Then
                    SampleIndexRegistry(.filename).sampleSet(.Samplename).sampleHole = addHole
                End If
            Else
                .filename = vbNullString
                .Samplename = vbNullString
            End If
            
        End With
    End If
End Sub

Public Function isValidSampleIn(slot As Integer)
    isValidSampleIn = False
    With ChangerSamples(slot)
        If SampleIndexRegistry.IsValidSample(.filename, .Samplename) Then isValidSampleIn = True
    End With
End Function

Private Sub mnuDeleteNext_Click()
    Dim holeid As Integer
    Dim i As Integer
    holeid = gridHole(gridSamples.row, gridSamples.Col)
    If Not Changer_ValidStart(holeid) Then Exit Sub
    For i = 0 To 9
        emptyHole (holeid + i) ' (September 2007 L Carporzen) Delete the next 9 samples
    Next i
    
    Me.refresh
    
End Sub

Private Sub mnuInsertSample_click()
    Dim selectionDialog As frmSampleSelect
    Set selectionDialog = New frmSampleSelect
    Dim holeid As Integer
    Dim addFile As String, addSample As String
    holeid = gridHole(gridSamples.row, gridSamples.Col)
    If Not Changer_ValidStart(holeid) Then Exit Sub
    selectionDialog.Show
    addFile = selectionDialog.filename
    addSample = selectionDialog.Samplename
    If SampleIndexRegistry.IsValidSample(addFile, addSample) Then IncorporateSample holeid, addFile, addSample
    Set selectionDialog = Nothing
    fillSamplesGrid
End Sub

Private Sub mnuSampleDelete_Click()
    Dim holeid As Integer
    holeid = gridHole(gridSamples.row, gridSamples.Col)
    If Not Changer_ValidStart(holeid) Then Exit Sub
    If SampleIndexRegistry.IsValidSample(ChangerSamples(holeid).filename, ChangerSamples(holeid).Samplename) And (Not Changer_isHole(holeid)) Then
        emptyHole holeid
    End If
End Sub

Private Sub mnuSampleDeleteAndShift_click()
    Dim holeid As Integer
    holeid = gridHole(gridSamples.row, gridSamples.Col)
    If Not Changer_ValidStart(holeid) Then Exit Sub
    If SampleIndexRegistry.IsValidSample(ChangerSamples(holeid).filename, ChangerSamples(holeid).Samplename) And (Not Changer_isHole(holeid)) Then
        emptyHole holeid
        SetSequenceVariables
        If Chg_isAscend Then
            ShiftSamples holeid, SlotMax, -1
        Else
            ShiftSamples SlotMin, holeid, 1
        End If
    End If
    
    Me.refresh
End Sub

Private Sub mnuSampleInfo_Click()
    Dim holeid As Integer
    holeid = gridHole(gridSamples.row, gridSamples.Col)
    If Not Changer_ValidStart(holeid) Then Exit Sub
    'MsgBox "Sample ID: " & Str$(GetSampleId(holeid)) & vbCr & _
        "Sample Name: " & samplesList(holeid) & vbCr & _
        "Sample File: " & SamFileRegistry(samplesFiles(holeid)).filename, vbOKOnly, "Sample Info - Hole " & Str$(holeid)
End Sub

Private Function numChangerSamples() As Integer
    Dim i As Integer
    numChangerSamples = 0
    If UseXYTableAPS Then
    
        For i = SlotMin To SlotMax
            If Changer_isHole(i) Then i = i + 1
            If LenB(ChangerSamples(i).filename) <> 0 Then
                numChangerSamples = numChangerSamples + 1
            End If
        Next i
            
    Else
    
        For i = SlotMin To (SlotMax - 1)
            If Changer_isHole(i) Then i = i + 1
            If LenB(ChangerSamples(i).filename) <> 0 Then
                numChangerSamples = numChangerSamples + 1
            End If
        Next i
        
    End If
End Function

Public Function ProcessOrderSlot(ByVal processnum As Integer, _
    isAscend As Boolean) As Integer
    ' This function should return the n'th slot to process when
    ' using the sample changer. Return -1 if we've already "passed" the
    ' last sample.  This returns slots in sequential order, starting from
    ' the lowest numbered slot with a sample in it.  It skips slots that
    ' do not contain samples, but includes "red" slots (for holder
    ' measurements).
    ' It needs upgrading to handle samples across the line.
    Dim i As Integer, j As Integer
    ' curNum        - current item we're looking for (holes and samples)
    ' curSlot       - current slot we're looking at
    ' samplespassed - actual number of samples we've skipped past
    Dim curSlot As Integer, curNum As Integer, samplespassed As Integer
    ProcessOrderSlot = -1
    If processnum = 1 Then Exit Function
    ' Find the slot of the correct "item" to process
    curNum = 2
    curSlot = FirstSlot
    samplespassed = 0
    If curSlot = -1 Then Exit Function
    While samplespassed < numChangerSamples And ProcessOrderSlot = -1
        If curSlot < SlotMin Then curSlot = SlotMax
        If curSlot > SlotMax Then curSlot = SlotMin
        If Changer_isHole(curSlot) Then
            If Changer_isHole(curSlot) Then
                If curNum = processnum Then ProcessOrderSlot = curSlot
                curNum = curNum + 1
            End If
        ElseIf LenB(ChangerSamples(curSlot).Samplename) <> 0 Then
            If curNum = processnum Then ProcessOrderSlot = curSlot
            curNum = curNum + 1
            samplespassed = samplespassed + 1
        End If
        If isAscend Then
            curSlot = curSlot + 1
        Else
            curSlot = curSlot - 1
        End If
    Wend
End Function

Public Sub ProcessSamplesToQueue(chgAscend As Boolean, chgLoadReturn As Boolean, chgDoReturn As Boolean, chgRepHolder As Boolean)
    ' This procedure begins processing samples according to the list
    ' of samples in mainchanger.  That form holds a record of the order
    ' and positions of samples that need to be processed.
    '
    ' Revision History:
    '
    '   7-15-99   Geoff Matters     Removed 'Pos' variable, lines marked DELETE
    '   8-30-99   Geoff Matters     Added 'repHolder' arg
    '  11-12-03   Bob Kopp          rewrote to send samples to sample queue
    '  12-13-03   Bob Kopp          removed isAscend
    Dim i As Integer, sampSlot As Integer
    Dim testSAMFile As SampleIndexRegistration
    Dim NumSamplesAdded As Integer
    NumSamplesAdded = 0
    
    If Not currentPosInitialized Then GetCurrentChangerPos
    Set newSampQueue = New SampleCommands
    frmProgram.StatBarNew "Processing samples to queue..."
    countFileCalls
    If SampleIndexRegistry.Count = 0 Then Exit Sub
    For i = 1 To SampleIndexRegistry.Count
        If WhichFilesLoaded(i) > 0 Then
         With SampleIndexRegistry(i)
            If .measurementSteps.Count > 1 Then .doBoth = False
            If .doBoth And .doUp Then
                ' Initialize the ".up" file if we are about to do an "up" measurement
                newSampQueue.Add "InitUp", 0, .filename
            End If
         End With
        End If
    Next i
    ' Measure the holder first
    newSampQueue.Add "Holder"
    i = 1
    Do While sampSlot <> -1                   ' while there is a sample to process
        If Prog_halted Then Exit Sub ' (September 2007 L Carporzen) New version of the Halt button
        ' Next sample...
        i = i + 1
        sampSlot = ProcessOrderSlot(i, chgAscend)
        If sampSlot = -1 Then Exit Do
        If Changer_isHole(sampSlot) Then
            ' Only perform repeat holder measurements if requested
            ' (For strong samples, holder does not need to be remeasured)
            If chgRepHolder And newSampQueue(newSampQueue.Count).commandType <> "Holder" Then
                newSampQueue.Add "Holder", sampSlot
            End If
        Else
            If isValidSampleIn(sampSlot) Then
                newSampQueue.Add "Meas", sampSlot
                NumSamplesAdded = NumSamplesAdded + 1
                If (NumSamplesAdded Mod SamplesBetweenHolder) = 0 Then
                      If chgRepHolder And newSampQueue(newSampQueue.Count).commandType <> "Holder" Then
                            newSampQueue.Add "Holder", HoleSlotNum
                      End If
                End If
            End If
            ' Process slot "sampSlot"
        End If
    Loop
    If chgLoadReturn Then
    
        If UseXYTableAPS Then
            'Goto the changer load position
            newSampQueue.Add "Goto", -1
        
        Else
        
            ' Move to the hole with the first sample!
             newSampQueue.Add "Goto", (ProcessOrderSlot(2, chgAscend))
             
        End If
    
    
    End If
    newSampQueue.Preprocess
    If chgDoReturn Then
        If UseXYTableAPS Then
        
            'Goto the changer load position
            newSampQueue.Add "Goto", -1
        
        Else
        
            ' Move to the hole with the first sample!
             newSampQueue.Add "Goto", (ProcessOrderSlot(2, chgAscend))
             
        End If
    End If
    SampQueue.Assimilate newSampQueue
    frmProgram.StatBarNew vbNullString
End Sub

Public Sub RefreshControls()
    If IsMasterList Then
        Me.Caption = "Hole Sample List - Master List"
        cmdSeq.Visible = True
        FrameOrder.Visible = False
        FrameReloadPos.Visible = False
        FrameFinalPos.Visible = False
        FrameRepHolder.Visible = False
        mnuInsertSample.Visible = False
        mnuSampleDelete.Visible = False
        mnuSampleDeleteAndShift.Visible = False
    Else
        Me.Caption = "Hole Sample List - New Sample Set"
        cmdSeq.Visible = False
        FrameOrder.Visible = True
        FrameReloadPos.Visible = False
        FrameFinalPos.Visible = False
        FrameRepHolder.Visible = False
        mnuSampleDelete.Visible = True
        mnuInsertSample.Visible = True
        mnuSampleDeleteAndShift.Visible = True
    End If
    
    Me.refresh
    
End Sub

Public Function setOrder(ascending As Boolean)
    Chg_isAscend = ascending
    If ascending Then
        optOrder(0).Value = True
        optOrder(1).Value = False
    Else
        optOrder(1).Value = True
        optOrder(0).Value = False
    End If
    'SetSequenceVariables
End Function

Private Sub SetSequenceVariables()
    Chg_isAscend = optOrder(0).Value
    'Chg_doReturn = optReturn(0).value
    'Chg_loadReturn = optLoadReturn(0).value
    'Chg_repHolder = optRepeatHolder(0).value
End Sub

Public Sub ShiftSamples(ShiftStart As Integer, ShiftEnd As Integer, ShiftSize As Integer)
    Dim i As Integer
    If Not (Changer_ValidStart(ShiftStart) And Changer_ValidStart(ShiftEnd)) Then Exit Sub
    If ShiftStart > ShiftEnd Then Exit Sub
    If ShiftSize > 0 Then
        For i = ShiftEnd To ShiftStart Step -1
            i = SkipHoles(i, False)
            If SkipHoles(i - ShiftSize, False) < ShiftStart Then Exit For
            With ChangerSamples((SkipHoles(i - ShiftSize, False)))
                IncorporateSample i, .filename, .Samplename
            End With
        Next i
    Else
        For i = ShiftStart To ShiftEnd Step 1
            i = SkipHoles(i, True)
            If SkipHoles(i - ShiftSize, True) > ShiftEnd Then Exit For
            With ChangerSamples((SkipHoles(i - ShiftSize, True)))
                IncorporateSample i, .filename, .Samplename
            End With
        Next i
    End If
    RefreshControls
End Sub

Public Function SkipHoles(slotpos As Integer, asc As Boolean) As Integer
    ' This function returns slotpos if the slot is valid.
    ' Otherwise, it returns the number of the next valid slot,
    ' incrementing if asc is true, decrementing if not
    If Changer_isHole(slotpos) Then
        ' We need a valid slot that is not a hole
        If asc Then
            slotpos = slotpos + 1               ' Increment
        Else
            slotpos = slotpos - 1               ' Decrement
        End If
    End If
    SkipHoles = slotpos
End Function

