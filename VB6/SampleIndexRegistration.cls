VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SampleIndexRegistration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"SampleIndexRegistrations"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvarFileDir As String 'local copy
Private mvarFilename As String 'local copy
Private mvarBackupFileDir As String 'local copy
Private mvarSampleCode As String 'local copy
Private mvarAvgSteps As Integer 'local copy
Private mvardoUp As Boolean 'local copy
Private mvardoBoth As Boolean 'local copy
Private mvarlocality As String 'local copy
Private mvarsiteLat As Double 'local copy
Private mvarsiteLong As Double 'local copy
Private mvarmagDec As Double 'local copy
Private mvarmeasurementSteps As RockmagSteps 'local copy
Private mvarsampleSet As Samples 'local copy
Public RockmagMode As Boolean

Public Property Let avgSteps(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.AvgSteps = 5
    mvarAvgSteps = vData
End Property

Public Property Get avgSteps() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.AvgSteps
    avgSteps = mvarAvgSteps
End Property

Public Property Let BackupFileDir(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.BackupFileDir = 5
    mvarBackupFileDir = vData
End Property

Public Property Get BackupFileDir() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.BackupFileDir
    BackupFileDir = mvarBackupFileDir
End Property

Public Property Get curDemag() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.curDemag
    Dim demaglev As Long
    Dim demagstrlen As Integer
    Dim i As Integer
    
    curDemag = vbNullString
    
    If measurementSteps.Count = 0 Then
        Exit Sub
    Else
        curDemag = measurementSteps.CurrentStep.DemagStepLabel
    End If
    
End Property

Public Property Get curDemagLong() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.curDemag
    Dim demaglev As Long
    Dim demagstrlen As Integer
    Dim i As Integer
    
    curDemagLong = vbNullString
    
    If measurementSteps.Count = 0 Then
        Exit Property
    Else
        curDemagLong = measurementSteps.CurrentStep.DemagStepLabelLong
    End If
    
End Property

Public Property Get CurrentStepFilePathPrefix() As String
    CurrentStepFilePathPrefix = filedir & "\" & SampleCode _
        & "\" & curDemag
End Property

Public Property Let doBoth(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.doBoth = 5
    mvardoBoth = vData
End Property

Public Property Get doBoth() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.doBoth
    doBoth = mvardoBoth
End Property

Public Property Let doUp(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.doUp = 5
    mvardoUp = vData
End Property

Public Property Get doUp() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.doUp
    doUp = mvardoUp
End Property

Public Property Let filedir(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FileDir = 5
    mvarFileDir = vData
End Property

Public Property Get filedir() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FileDir
    If LenB(mvarFileDir) > 0 Then
        filedir = mvarFileDir
    Else
        filedir = "."
    End If
End Property

Public Property Let filename(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FileName = 5
    mvarFilename = vData
End Property

Public Property Get filename() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FileName
    filename = mvarFilename
End Property

Public Property Let locality(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.locality = 5
    mvarlocality = vData
End Property

Public Property Get locality() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.locality
    locality = mvarlocality
End Property

Public Property Let magDec(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.magDec = 5
    mvarmagDec = vData
End Property

Public Property Get magDec() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.magDec
    magDec = mvarmagDec
End Property

Public Property Set measurementSteps(ByVal vData As RockmagSteps)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.measurementSteps = Form1
    Set mvarmeasurementSteps = vData
End Property

Public Property Get measurementSteps() As RockmagSteps
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.measurementSteps
    Set measurementSteps = mvarmeasurementSteps
End Property

Public Property Let SampleCode(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SampleCode = 5
    mvarSampleCode = vData
End Property

Public Property Get SampleCode() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SampleCode
    SampleCode = mvarSampleCode
End Property

Public Property Set sampleSet(ByVal vData As Samples)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.sampleSet = Form1
    Set mvarsampleSet = vData
End Property

Public Property Get sampleSet() As Samples
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sampleSet
    Set sampleSet = mvarsampleSet
End Property

Public Property Get SampleSetFilePathPrefix() As String
    SampleSetFilePathPrefix = filedir & "\" & SampleCode _
        & "\" & SampleCode
End Property

Public Property Let siteLat(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.siteLat = 5
    mvarsiteLat = vData
End Property

Public Property Get siteLat() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.siteLat
    siteLat = mvarsiteLat
End Property

Public Property Let siteLong(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.siteLong = 5
    mvarsiteLong = vData
End Property

Public Property Get siteLong() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.siteLong
    siteLong = mvarsiteLong
End Property

Public Property Get UpFilePath() As String
    UpFilePath = CurrentStepFilePathPrefix & ".UP"
End Property

Private Sub Class_Initialize()
    Set mvarsampleSet = New Samples
    Set mvarmeasurementSteps = New RockmagSteps
End Sub

Private Sub Class_Terminate()
    Set mvarsampleSet = Nothing
    Set mvarmeasurementSteps = Nothing
End Sub

Public Sub loadInfo()
    ' This subroutine loads the descriptive information
    ' and the samples from the index SAM file.
    
    Dim LineText As String
    Dim filenum As Integer                ' File number for access
    
    mvarsampleSet.Clear
    mvarsampleSet.IndexFile = filename
    
    If Not LenB(dir$(filename)) > 0 Then Exit Sub
    
    filenum = FreeFile
    Open filename For Input Access Read As #filenum ' Open file.
    Line Input #filenum, LineText                   ' Grab sample description
    mvarlocality = LineText                         ' Update SAM info in memory
    
    Line Input #filenum, LineText                   ' Grab a line from file
    mvarsiteLat = val(Mid(LineText, 1, 6))         ' Get the lat string
    mvarsiteLong = val(Mid$(LineText, 7, 11))     ' Get the lon string
    mvarmagDec = val(Mid$(LineText, 13, 17))      ' Get the mag dec string
    Do While Not EOF(filenum)                       ' Loop until end of file.
        Input #filenum, LineText                    ' Grab a line from the file
        If LenB(LineText) > 0 Then mvarsampleSet.Add LineText
    Loop
    Close #filenum                                  ' Close file.

End Sub

