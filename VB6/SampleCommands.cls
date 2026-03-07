VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SampleCommands"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"SampleCommand"
Attribute VB_Ext_KEY = "Member0" ,"SampleCommand"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable to hold collection
Private mCol As Collection
Private keyCounter As Long
Private WhichFilesLoaded() As Integer
Private Executing As Boolean
Private maxAvgStepsOfFiles

Public Property Get Count() As Long
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    Count = mCol.Count
End Property

Public Property Get Item(vntIndexKey As Variant) As SampleCommand
Attribute Item.VB_UserMemId = 0
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
On Error GoTo oops
  Set Item = mCol(vntIndexKey)
  On Error GoTo 0
  Exit Sub
oops:
    Set Item = New SampleCommand
End Property

Public Property Get maxAvgSteps() As Integer
    countFileCalls
    If maxAvgStepsOfFiles < 1 Then
        maxAvgSteps = 1
    Else
        maxAvgSteps = maxAvgStepsOfFiles
    End If
End Property

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Set NewEnum = mCol.[_NewEnum]
End Property

Public Function Add(ByVal commandType As String, _
                    Optional ByVal hole As Long = 0, _
                    Optional ByVal fileid As String = vbNullString, _
                    Optional ByVal Sample As String = vbNullString, _
                    Optional ByVal BeforeStep As Long, _
                    Optional ByVal AfterStep As Long) As SampleCommand
                    
    'create a new object
    Dim objNewMember As SampleCommand
    Set objNewMember = New SampleCommand
    'set the properties passed into the method
    objNewMember.key = "C" & Str$(keyCounter)
    objNewMember.commandType = commandType
    objNewMember.fileid = fileid
    objNewMember.Sample = Sample
    objNewMember.hole = hole
On Error GoTo carryforth:
    If LenB(MainChanger.ChangerFileName(hole)) > 0 Then
        objNewMember.fileid = MainChanger.ChangerFileName(hole)
    End If
    If LenB(MainChanger.ChangerSampleName(hole)) > 0 Then
        objNewMember.Sample = MainChanger.ChangerSampleName(hole)
    End If
On Error GoTo 0
carryforth:
    If BeforeStep > 0 Then
        mCol.Add objNewMember, "C" & Str$(keyCounter), Before:=BeforeStep
    ElseIf AfterStep > 0 Then
        mCol.Add objNewMember, "C" & Str$(keyCounter), After:=AfterStep
    Else
        mCol.Add objNewMember, "C" & Str$(keyCounter)
    End If
    keyCounter = keyCounter + 1
    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing
End Function

Public Sub Assimilate(assimilateeCommands As SampleCommands)
    Dim i As Integer
    For i = 1 To assimilateeCommands.Count
        With assimilateeCommands(i)
            Add .commandType, .hole, .fileid, .Sample
        End With
    Next i
End Sub

Private Sub Class_Initialize()
    'creates the collection when this class is created
    Set mCol = New Collection
    keyCounter = 0
End Sub

Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Set mCol = Nothing
End Sub

Public Sub Clear()
    Set mCol = Nothing
    Set mCol = New Collection
    keyCounter = 0
End Sub

Private Sub countFileCalls()
    Dim i As Integer
    Dim targetItem As ListItem
    Dim targetFile As Integer
    ReDim WhichFilesLoaded(SampleIndexRegistry.Count)
    For i = 0 To SampleIndexRegistry.Count
        WhichFilesLoaded(i) = 0
    Next i
    For i = 1 To Me.Count
        If SampleIndexRegistry.IsValidFile(Item(i).fileid) Then
            targetFile = SampleIndexRegistry.Index(Item(i).fileid)
            If WhichFilesLoaded(targetFile) = 0 Then
                If maxAvgStepsOfFiles < SampleIndexRegistry(targetFile).avgSteps Then maxAvgStepsOfFiles = SampleIndexRegistry(targetFile).avgSteps
            End If
            WhichFilesLoaded(targetFile) = WhichFilesLoaded(targetFile) + 1
        End If
    Next i
End Sub

Public Sub Execute()
    Static curpos As Long
    Dim activeItem As SampleCommand
    If Prog_halted Then Exit Sub ' (September 2007 L Carporzen) New version of the Halt button
    If Executing Then Exit Sub
    Executing = True
    MeasurementsSinceHolder = 1000
    If Not Changer_ValidStart(curpos) Then frmChanger.GetCurrentChangerPos
    curpos = frmDCMotors.ChangerHole
    Load frmMeasure              ' Load the measurement form
    Load frmStats                ' Load the stats form
    frmProgram.StatBarNew "Measuring samples..."
    frmVacuum.MotorPower True
    DelayTime 0.5
                
    Do While Count > 0
        If Prog_halted Then Exit Sub ' (September 2007 L Carporzen) New version of the Halt button
        Flow_WaitForUnpaused
        Set activeItem = Item(1)
        activeItem.Execute
        Remove 1
    Loop
    Set activeItem = Nothing
    Executing = False
    frmProgram.StatBarNew vbNullString
    If UseXYTableAPS Then
        frmDCMotors.MoveToCorner
    Else
        Changer_NearestHole
    End If
     
    MsgBox "Measurements are done."
    
    
    'Check to see if rockmag was done
    If frmSampleIndexRegistry.optSAMSetDemag(1).Value = True Or _
       frmSampleIndexRegistry.optSAMSetDemag(5).Value = True _
    Then
        'This is a rockmag run
        'automatically turn off the air
        If modConfig.DoDegausserCooling = True Then
            frmVacuum.DegausserCooler False
            MsgBox "Please:" & vbNewLine & vbNewLine & _
                   " - Verify the air is off." & vbNewLine & _
                   " - Switch off the power to the Rockmag coil temperature sensors."
        Else
            MsgBox "Please:" & vbNewLine & vbNewLine & _
                   " - Turn off the air" & vbNewLine & _
                   " - Switch off the power to the Rockmag coil temperature sensors."
        End If
    End If
        
    SetCodeLevel CodeBlue, True
    frmVacuum.MotorPower False
    
    If frmRerunSamples.scanForLargeErrorAngles > 0 Then
        frmRerunSamples.ZOrder
        frmRerunSamples.Show
    End If
End Sub

Public Function fileInUse(filename As String) As Boolean
    fileInUse = False
    If SampleIndexRegistry.IsValidFile(filename) Then
        countFileCalls
        If WhichFilesLoaded(SampleIndexRegistry.Index(filename)) > 0 Then fileInUse = True
    End If
End Function

Public Sub Preprocess()
    Dim i As Integer
    Dim samplesflipped() As Boolean
    Dim NumSamplesAdded As Integer: NumSamplesAdded = 0
    
    ReDim samplesflipped(SampleIndexRegistry.Count)
    
    For i = 1 To Count
        With Item(i)
            On Error GoTo carryforth
                If SampleIndexRegistry.IsValidFile(.fileid) Then
                    Select Case .commandType
                        Case "InitUp":
                            If SampleIndexRegistry(.fileid).doBoth Then
                            
                                If UseXYTableAPS Then
                                
                                    Add "Flip", -1, .fileid
                                    
                                Else
                            
                                    Add "Flip", 0, .fileid
                                    
                                End If
                                
                                Add "Holder", modChanger.Find_NearestChangerHole
                                
                            End If
                            samplesflipped(SampleIndexRegistry.Index(.fileid)) = True
                            
                        Case "Meas":
                            If samplesflipped(SampleIndexRegistry.Index(.fileid)) And _
                                SampleIndexRegistry(.fileid).doBoth Then
                        
                                Add "Meas", .hole, .fileid, .Sample
                                
                                NumSamplesAdded = NumSamplesAdded + 1
                                If (NumSamplesAdded Mod SamplesBetweenHolder) = 0 And _
                                   SamplesBetweenHolder < SlotMax And _
                                   SamplesBetweenHolder > 0 Then
                                    
                                    On Error GoTo MainChanger_AccessError
                                    
                                    If Not MainChanger Is Nothing Then
                                        If MainChanger.optRepeatHolder(0).Value Then
                                            Add "Holder", modChanger.Find_NearestChangerHole
                                        End If
                                    End If
                                    
                                    On Error GoTo 0
                                    
MainChanger_AccessError:
                                    
                                End If
                            End If
                            
                        Case "Flip"
                            samplesflipped(SampleIndexRegistry.Index(.fileid)) = False
                        
                    End Select
                End If
            On Error GoTo 0
carryforth:
        End With
    Next i
    countFileCalls
    For i = 1 To SampleIndexRegistry.Count
        If WhichFilesLoaded(i) > 0 Then
        
            If UseXYTableAPS Then
                'Goto XY Load Position
                Add "Fin", -1, Item(i).fileid
            Else
                'Goto hole 0
                Add "Fin", 0, Item(i).fileid
            End If
        End If
    Next i
End Sub

Public Sub Remove(vntIndexKey As Variant)
    'used when removing an element from the collection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)
    mCol.Remove vntIndexKey
End Sub

