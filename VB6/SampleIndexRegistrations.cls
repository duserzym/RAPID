VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SampleIndexRegistrations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"SampleIndexRegistration"
Attribute VB_Ext_KEY = "Member0" ,"SampleIndexRegistration"
'locsal variable to hold collection
Private mCol As Collection
Private SampleHolderIndex As SampleIndexRegistration

Const SampleHolderIndexTag = "!Holder"

Public Property Get Count() As Long
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    Count = mCol.Count
End Property

Public Property Get Item(vntIndexKey As Variant) As SampleIndexRegistration
Attribute Item.VB_UserMemId = 0
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Dim blankItem As SampleIndexRegistration
  On Error GoTo oops
  If vntIndexKey = SampleHolderIndexTag Then
    Set Item = SampleHolderIndex
  Else
    Set Item = mCol(vntIndexKey)
  End If
  On Error GoTo 0
  Exit Sub
oops:
  Select Case Err.number  ' Evaluate error number.
    Case 5         ' invalid index key
        Set blankItem = New SampleIndexRegistration
        Set Item = blankItem
        Set blankItem = Nothing
    Case Else
        MsgBox "unknown error in sampleindexregistrations"
  End Select
  
End Property

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Set NewEnum = mCol.[_NewEnum]
End Property

Public Function Add(filedir As String, _
                    filename As String, _
                    BackupFileDir As String, _
                    SampleCode As String, _
                    Optional avgSteps As Integer = 1, _
                    Optional doUp As Boolean = True, _
                    Optional doBoth As Boolean = True, _
                    Optional measurementSteps As RockmagSteps, _
                    Optional ByVal RockmagMode As Boolean = False) As SampleIndexRegistration
                    
    'create a new object
    Dim objNewMember As SampleIndexRegistration
    Set objNewMember = New SampleIndexRegistration


    'set the properties passed into the method
    objNewMember.filedir = filedir
    objNewMember.filename = filename
    objNewMember.BackupFileDir = BackupFileDir
    objNewMember.SampleCode = SampleCode
    objNewMember.avgSteps = avgSteps
    objNewMember.doUp = doUp
    objNewMember.doBoth = doBoth
    objNewMember.RockmagMode = RockmagMode
    Set objNewMember.measurementSteps = measurementSteps
    
    objNewMember.loadInfo
    
    mCol.Add objNewMember, filename

    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Function AddSampleIndex(sampleindex As SampleIndexRegistration) As SampleIndexRegistration
    
    On Error GoTo oops
    
    With sampleindex
        If LenB(Item(.filename).filename) > 0 Then Remove .filename
        Set AddSampleIndex = Add(.filedir, .filename, .BackupFileDir, .SampleCode _
            , .avgSteps, .doUp, .doBoth, .measurementSteps, .RockmagMode)
    End With
    'return the object created
    
    
    On Error GoTo 0
    Exit Function
oops:
    Select Case Err.number
    Case Else
        MsgBox "Error in SampleIndexRegistrations" & vbCrLf & Err.number & ": " & Err.Description
    End Select
    Set objNewMember = Nothing

End Function

Private Sub Class_Initialize()
    'creates the collection when this class is created
    Set mCol = New Collection
    
    MakeSampleHolder
        
End Sub

Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Set mCol = Nothing
    Set SampleHolderIndex = Nothing
End Sub

Public Sub Clear()
    Set mCol = Nothing
    Set mCol = New Collection
End Sub

Public Function Index(filename As String) As Integer
    On Error GoTo fin
    Dim i As Integer
    Dim okay As Boolean
    okay = False
    For i = 1 To Count
        If Item(i).filename = filename Then
            okay = True
            Exit For
        End If
    Next i
    If Not okay Then Index = 0 Else Index = i
fin:
    On Error GoTo 0
End Function

Public Function IsValidFile(filename As String)
    IsValidFile = False
    On Error GoTo fin
    IsValidFile = (LenB(Item(filename).filename) > 0)
fin:
    On Error GoTo 0
End Function

Public Function IsValidSample(filename As String, Samplename As String)
    IsValidSample = False
    On Error GoTo fin
    IsValidSample = Item(filename).sampleSet.IsValidSample(Samplename)
fin:
    On Error GoTo 0
End Function

'Public Sub MakeSampleHolder
'
' Original By: R Kopp
'  Orig. Date: ??, ca. 2002
'
' Modified By: I Hilburn
'        Date: March 11, 2011
'
' Original State:
'   1) This was a private function
'   2) SampleHolderIndex was not deallocated before allocation
'   3) SampleHolder global variable was not set inside of this subroutine
'      but was set instead in a code line in modProg following the initialization
'      of the SampleIndexRegistry object (whose constructor called this subroutine)
'
' Current State:
'
'    Inputs:    MeasSusceptibility
'               Optional Boolean (Default = False)
'               Allows the calling code to set whether or not the Holder will have it's
'               susceptibility measured every time.  The default at code start is that
'               MeasSusceptibility = False.  This can be overwritten later in the code
'               (in frmSampleIndexRegistry.updateMeasurementSteps)
'
'   Outputs:    None
'
'   Effects:    Allocates the SampleHolderIndex Sample Registry object within the
'               main system collection of Sample Index registries.
'               This subroutine can be used now to overwrite both the old Sample Holder Index object
'               And to also overwrite the SampleHolder sample object
'
'               At the end of this code, the SampleHolder object is overwritten with the
'               New SampleHolder object that was added to the SampleHolderIndex
'
'   Potential Bugs:
'               If the the SampleHolder object does not exist or is not accessible,
'               the error is handled and the SampleHolder object will not be overwritten.
'               Thus, there may be a mismatch between the SampleHolderIndex and the SampleHolder object
'               Leading to confusion and additional errors downstream in the Paleomag code.
Public Sub MakeSampleHolder(Optional ByVal MeasSusceptibility As Boolean = False)
    
    'Deallocate SampleHolderIndex, first
    '(in case there's a prior instance we're overwriting)
    Set SampleHolderIndex = Nothing
    
    'Reallocate as a new Sample Index Registry class object
    Set SampleHolderIndex = New SampleIndexRegistration
    
    'Setup SampleHolder Index
    With SampleHolderIndex
        .filename = SampleHolderIndexTag
        .avgSteps = 1
        .doUp = True
        .doBoth = False
        Set .measurementSteps = New RockmagSteps
        Set .sampleSet = New Samples
        
        '(March 10, 2011 - I Hilburn)
        'UGH! This one line has been the source of the wait time on the holder and
        'the holder always having it's susceptibility measured.
        'Instead of Defaulting measure holder susc. to True
        'I'm defaulting it to "False", though allowing that value to be passed in
        'as an argument to the MakeSampleHolder command, and then
        'will set it to true later in the code.
        .measurementSteps.Add "NRM", MeasureSusceptibility:=MeasSusceptibility
        .sampleSet.IndexFile = SampleHolderIndexTag
        .sampleSet.Add "Holder"
    End With
    
    'Reset the Sample stored in SampleHolder global object
    'With error checking in case a change is made to the code
    'that eliminates the SampleHolder global variable
    On Error GoTo NoSampleHolder:
        
        Set SampleHolder = Nothing
        Set SampleHolder = SampleHolderIndex.sampleSet("Holder")
        
    On Error GoTo 0
        
NoSampleHolder:
        
End Sub

Public Sub Remove(vntIndexKey As Variant)
    'used when removing an element from the collection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)


    mCol.Remove vntIndexKey
End Sub

Public Function SampleByIndex(vindex As Integer) As Sample
    Dim sampCount As Integer
    Dim i As Integer
    
    sampCount = SampleCount
    On Error GoTo oops
    If vindex < 1 Or vindex > sampCount Then GoTo oops
    For i = 1 To Count
        With Item(i).sampleSet
        If vindex <= .Count Then Exit For Else vindex = vindex - .Count
        End With
    Next i
    Set SampleByIndex = Item(i).sampleSet.Item(vindex)
    On Error GoTo 0
    Exit Function
oops:
    Set SampleByIndex = New Sample
End Function

Public Function SampleCount() As Integer
    Dim i As Integer
    Dim j As Integer

    SampleCount = 0
    If Count = 0 Then Exit Function
    For i = 1 To Count
        If Item(i).sampleSet.Count > 0 Then
            With Item(i).sampleSet
            For j = 1 To .Count
                SampleCount = SampleCount + 1
            Next j
            End With
        End If
    Next i
    
End Function

Public Function SampleFileByIndex(vindex As Integer) As String
    Dim sampCount As Integer
    Dim i As Integer
    
    sampCount = SampleCount
    SampleFileByIndex = vbNullString
    If vindex < 1 Or vindex > sampCount Then Exit Function
    For i = 1 To Count
        With Item(i).sampleSet
        If vindex <= .Count Then Exit For Else vindex = vindex - .Count
        End With
    Next i
    SampleFileByIndex = Item(i).filename
End Function

