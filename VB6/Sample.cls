VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Sample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private mvarSamplename As String 'local copy
Private mvarsampleHole As Long 'local copy
Private mvarIndexFile As String 'local copy
Private mvarCorePlateStrike As Double
Private mvarCorePlateDip    As Double
Private mvarBeddingStrike   As Double
Private mvarBeddingDip      As Double
Private mvarSampleHeight       As Long
Private mvarSusceptibility  As Double
Private mvarVol             As Double  ' Sample volume
Private mvarComment         As String
Public UpDownRatio     As Double
Public Moment           As Double
Public SignalToNoise   As Double
Public SignalToHolder  As Double
Private mvarFoldRotation    As Boolean
Private mvarFoldAxis        As Double
Private mvarFoldPlunge      As Double
Private mvarErrorAngle      As Double
Private alreadyReadSpec As Boolean
Private Const Hold = "Holder"

Public Property Get BackFilePath() As String
    With Parent
        If LenB(.BackupFileDir) > 0 Then
        '--------------------------------------------------------------------------------------------
        '   Quick Fix to double backslash filepath error
        '   9/10/2008
        '   Isaac Hilburn
        '   added replace(badstring,"\\","\",4) around each string that combined the
        '   .BackupFileDir with the .SampleCode, or with the .SampleCode and the .SampleName
        '   The last condition of the replace statement will avoid removing legitimate "\\"
        '   at the start of the backup file path, i.e. "\\Biopoesis\..." or "f:\\user\joe\..."
        '--------------------------------------------------------------------------------------------
            BackFilePath = Mid$(.BackupFileDir & "\" & .SampleCode & "\" & Samplename, 1, 3) & _
                            Replace(.BackupFileDir & "\" & .SampleCode & "\" & Samplename, "\\", "\", 4)
        Else
            BackFilePath = vbNullString
        End If
    End With
End Property

Public Property Let BeddingDip(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sampleHole = 5
    mvarBeddingDip = vData
End Property

Public Property Get BeddingDip() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sampleHole
    ReadSpec
    BeddingDip = mvarBeddingDip
End Property

Public Property Let BeddingStrike(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sampleHole = 5
    mvarBeddingStrike = vData
End Property

Public Property Get BeddingStrike() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sampleHole
    ReadSpec
    BeddingStrike = mvarBeddingStrike
End Property

Public Property Get Comment() As String
    Comment = mvarComment
End Property

Public Property Let CorePlateDip(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sampleHole = 5
    mvarCorePlateDip = vData
End Property

Public Property Get CorePlateDip() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sampleHole
    ReadSpec
    CorePlateDip = mvarCorePlateDip
End Property

Public Property Let CorePlateStrike(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sampleHole = 5
    mvarCorePlateStrike = vData
End Property

Public Property Get CorePlateStrike() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sampleHole
    ReadSpec
    CorePlateStrike = mvarCorePlateStrike
End Property

Public Property Let ErrorAngle(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sampleHole = 5
    mvarErrorAngle = vData
End Property

Public Property Get ErrorAngle() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sampleHole
    ErrorAngle = mvarErrorAngle
End Property

Public Property Let FoldAxis(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sampleHole = 5
    mvarFoldAxis = vData
End Property

Public Property Get FoldAxis() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sampleHole
    ReadSpec
    FoldAxis = mvarFoldAxis
End Property

Public Property Let FoldPlunge(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sampleHole = 5
    mvarFoldPlunge = vData
End Property

Public Property Get FoldPlunge() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sampleHole
    ReadSpec
    FoldPlunge = mvarFoldPlunge
End Property

Public Property Let FoldRotation(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sampleHole = 5
    mvarFoldRotation = vData
End Property

Public Property Get FoldRotation() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sampleHole
    ReadSpec
    FoldRotation = mvarFoldRotation
End Property

Public Property Let IndexFile(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sampleName = 5
    mvarIndexFile = vData
End Property

Public Property Get IndexFile() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sampleName
    IndexFile = mvarIndexFile
End Property

Public Property Get Parent() As SampleIndexRegistration
    If LenB(IndexFile) > 0 Then
        Set Parent = SampleIndexRegistry(IndexFile)
    Else
        Set Parent = New SampleIndexRegistration
    End If
End Property

Public Property Let SampleHeight(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sampleHole = 5
    mvarSampleHeight = vData
End Property

Public Property Get SampleHeight() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sampleHole
    SampleHeight = mvarSampleHeight
End Property

Public Property Let sampleHole(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sampleHole = 5
    mvarsampleHole = vData
End Property

Public Property Get sampleHole() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sampleHole
    sampleHole = mvarsampleHole
End Property

Public Property Let Samplename(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sampleName = 5
    mvarSamplename = vData
End Property

Public Property Get Samplename() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sampleName
    Samplename = mvarSamplename
End Property

Public Property Get SpecFilePath() As String
    With Parent
    '--------------------------------------------------------------------------------------------
    '   Quick Fix to double backslash filepath error
    '   9/10/2008
    '   Isaac Hilburn
    '   added replace(badstring,"\\","\",4) around each string that combined the
    '   .BackupFileDir with the .SampleCode, or with the .SampleCode and the .SampleName
    '   The last condition of the replace statement will avoid removing legitimate "\\"
    '   at the start of the backup file path, i.e. "\\Biopoesis\..." or "f:\\user\joe\..."
    '--------------------------------------------------------------------------------------------
        SpecFilePath = Mid$(.filedir & "\" & .SampleCode & "\" & Samplename, 1, 3) & _
                        Replace(.filedir & "\" & .SampleCode & "\" & Samplename, "\\", "\", 4)
    End With
End Property

Public Property Let Susceptibility(ByVal vData As Double)
    mvarSusceptibility = vData
End Property

Public Property Get Susceptibility() As Double
    Susceptibility = mvarSusceptibility
End Property

Public Property Let Vol(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sampleHole = 5
    mvarVol = vData
End Property

Public Property Get Vol() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sampleHole
    ReadSpec
    Vol = mvarVol
End Property

Public Sub BackupRockmagData()
    Dim TempFileStr As String
    If LenB(BackFilePath) = 0 Then Exit Sub
    If LenB(dir$(SpecFilePath)) = 0 Then Exit Sub
    On Error GoTo oops
    With Parent
        '--------------------------------------------------------------------------------------------
        '   Quick Fix to double backslash filepath error
        '   9/10/2008
        '   Isaac Hilburn
        '
        '   Created local variable, TempFileStr to store the corrected file string with
        '   erroneous "\\" replaced with "\"
        '   TempFileStr = mid$(badstring,1,3) & replace(badstring,"\\","\",4)
        '   The last condition of the replace statement will avoid removing legitimate "\\"
        '   at the start of the backup file path, i.e. "\\Biopoesis\..." or "f:\\user\joe\..."
        '   The mid$ function at the front adds back the first three characters of the filepath
        '--------------------------------------------------------------------------------------------
        TempFileStr = Mid(.BackupFileDir & "\" & .SampleCode, 1, 3) & _
                        Replace(.BackupFileDir & "\" & .SampleCode, "\\", "\", 4)
                        
        'Repeat this replace twice to deal with clumps of "\" up to 4 long (i.e. "\\\\", "\\\")
        TempFileStr = Mid(.BackupFileDir & "\" & .SampleCode, 1, 3) & Replace(.BackupFileDir & "\" & .SampleCode, "\\", "\", 4)
        TempFileStr = Mid(.BackupFileDir & "\" & .SampleCode, 1, 3) & Replace(.BackupFileDir & "\" & .SampleCode, "\\", "\", 4)
                        
        '--------------------------------------------------
        '   Code correction
        '   9/10/2008
        '   Isaac Hilburn
        '
        '   Changed check to see if backup directories exist
        '   from: LenB(Dir$(.BackupFileDir, vbDirectory)) = 0
        '   to: Not FileExists(.BackupFileDir)
        '   Original file check condition didn't seem to be
        '   working properly
        '--------------------------------------------------
        If Not FileExists(.BackupFileDir) Then
            MkDir .BackupFileDir
            If Not FileExists(.BackupFileDir) Then GoTo oops
        End If
        If Not FileExists(TempFileStr) Then
            MkDir TempFileStr
            If Not FileExists(TempFileStr) Then GoTo oops
        End If
    End With
    FileCopy SpecFilePath & ".rmg", BackFilePath & ".rmg"
    On Error GoTo 0
    Exit Sub
oops:
        frmSendMail.MailNotification "Error writing rockmag backup file", "Error writing backup file " & Parent.BackupFileDir & "\" & Parent.SampleCode, "Yellow"
End Sub

Public Sub BackupSpecFile()
    Dim TempFileStr As String
    If LenB(BackFilePath) = 0 Then Exit Sub
    If LenB(dir$(SpecFilePath)) = 0 Then Exit Sub
    On Error GoTo oops
    With Parent
        '--------------------------------------------------------------------------------------------
        '   Quick Fix to double backslash filepath error
        '   9/10/2008
        '   Isaac Hilburn
        '   Created local variable, TempFileStr to store the corrected file string with
        '   erroneous "\\" replaced with "\"
        '   TempFileStr = mid$(badstring,1,3) & replace(badstring,"\\","\",4)
        '   The last condition of the replace statement will avoid removing legitimate "\\"
        '   at the start of the backup file path, i.e. "\\Biopoesis\..." or "f:\\user\joe\..."
        '   The mid$ function at the front adds back the first three characters of the filepath
        '--------------------------------------------------------------------------------------------
        TempFileStr = Mid(.BackupFileDir & "\" & .SampleCode, 1, 3) & _
                    Replace(.BackupFileDir & "\" & .SampleCode, "\\", "\", 4)
        
        'Repeat this replace twice to deal with clumps of "\" up to 4 long (i.e. "\\\\", "\\\")
        TempFileStr = Mid(.BackupFileDir & "\" & .SampleCode, 1, 3) & Replace(.BackupFileDir & "\" & .SampleCode, "\\", "\", 4)
        TempFileStr = Mid(.BackupFileDir & "\" & .SampleCode, 1, 3) & Replace(.BackupFileDir & "\" & .SampleCode, "\\", "\", 4)
        
        If Not FileExists(.BackupFileDir) Then
            MkDir .BackupFileDir
            If Not FileExists(.BackupFileDir) Then GoTo oops
        End If
        If Not FileExists(TempFileStr) Then
            MkDir TempFileStr
            If Not FileExists(TempFileStr) Then GoTo oops
        End If
    FileCopy SpecFilePath, BackFilePath
    On Error GoTo 0
    Exit Sub
oops:
        frmSendMail.MailNotification "Error writing backup file", "Error writing backup file " & .BackupFileDir & "\" & .SampleCode, CodeYellow
    End With
End Sub

Public Sub ReadSpec()
    '' This function reads the file for the current specimen
    ' and returns the data in a "Specimen" type
    ' set some default values in case we get an error
    mvarCorePlateStrike = 0
    mvarCorePlateDip = 0
    mvarBeddingStrike = 0
    mvarBeddingDip = 0
    mvarVol = 1
    mvarFoldRotation = False
    mvarFoldAxis = 0
    mvarFoldPlunge = 0
    On Error GoTo ReadError
    Dim filenum As Integer
    Dim S As String, SS As String
    If alreadyReadSpec Then Exit Sub
    filenum = FreeFile
    Open SpecFilePath For Input As #filenum
    Line Input #filenum, mvarComment
    Line Input #filenum, SS$
    Close #filenum
    mvarCorePlateStrike = val(Mid$(SS$, 9, 5))
    mvarCorePlateDip = val(Mid$(SS$, 15, 5))
    mvarBeddingStrike = val(Mid$(SS$, 21, 5))
    mvarBeddingDip = val(Mid$(SS$, 27, 5))
    mvarVol = val(Mid$(SS$, 33, 5))
    If mvarVol = 0 Then mvarVol = 1
    If Len(SS$) > 42 Then
        mvarFoldRotation = True
        mvarFoldAxis = val(Mid$(SS$, 39, 5))
        mvarFoldPlunge = val(Mid$(SS$, 45, 5))
    Else
        mvarFoldRotation = False
        mvarFoldAxis = 0
        mvarFoldPlunge = 0
    End If
On Error GoTo 0
Exit Sub
ReadError:
    If Samplename <> Hold Then MsgBox ("Error in ReadSpec: " & vbCr & _
        specpath & " could not be opened.")
End Sub

Public Function ReadUpMeasurements() As MeasurementBlocks
    Dim filenum As Integer
    Dim UpFilePath As String
    Dim workingBlock As MeasurementBlock
    Dim workingvector As Cartesian3D
    Dim whole_file As String
    Dim lines As Variant
    Dim the_array As Variant
    Dim num_rows As Long
    Dim r As Long
    Dim readSampname As String
    Dim readTotalBlocks As Long
    Dim readSampname2 As String
    Dim readMeastype As String
    Dim readBlocknum As Long
    Dim readMeasnum As Long
    Dim CurrentSampleFirstLine As Long
    Dim CurrentSampleLastLine As Long
    Dim TotalBlocks As Long
    Dim currentBlock As Long
    CurrentSampleFirstLine = -1
    CurrentSampleLastLine = -1
    TotalBlocks = -1
    currentBlock = 0
    Set ReadUpMeasurements = New MeasurementBlocks
    sampData = vbNullString
    filenum = FreeFile
    UpFilePath = Parent.UpFilePath
    Open UpFilePath For Input As #filenum
    whole_file = Input$(LOF(filenum), #filenum)
    Close #filenum
    lines = Split(whole_file, vbCrLf)
    num_rows = UBound(lines)
    ReDim the_array(num_rows)
    For r = 0 To num_rows
        the_array(r) = Split(lines(r), "|")
    Next r
    For r = 0 To UBound(the_array) - 1
        readSampname = the_array(r)(0)
        If readSampname = Samplename Then
            readTotalBlocks = val(the_array(r)(2))
            If r - 1 + 10 * readTotalBlocks < UBound(the_array) Then
                readSampname2 = the_array(r - 1 + 10 * readTotalBlocks)(0)
                If (readSampname2 = Samplename) And (the_array(r)(1) = "U") Then
                    CurrentSampleFirstLine = r
                    CurrentSampleLastLine = r - 1 + 10 * readTotalBlocks
                End If
            End If
        End If
    Next r
    If CurrentSampleLastLine > CurrentSampleFirstLine Then
        TotalBlocks = val(the_array(CurrentSampleFirstLine)(2))
        If TotalBlocks = 0 Then Exit Function
        For r = CurrentSampleFirstLine To CurrentSampleLastLine
            If val(the_array(r)(4)) > currentBlock Then
                If currentBlock > 0 Then
                    ReadUpMeasurements.Add workingBlock
                    Set workingBlock = Nothing
                End If
                currentBlock = val(the_array(r)(4))
                Set workingBlock = New MeasurementBlock
                If the_array(r)(1) = "D" Then
                    workingBlock.isUp = False
                Else
                    workingBlock.isUp = True
                End If
            End If
            With workingBlock
                readMeastype = the_array(r)(3)
                readBlocknum = val(the_array(r)(4))
                readMeasnum = val(the_array(r)(5))
                Set workingvector = New Cartesian3D
                workingvector.ReadString (the_array(r)(6))
                If readMeastype = "Z" Then
                    .SetBaseline readMeasnum, workingvector
                ElseIf readMeastype = "H" Then
                    .SetHolder readMeasnum, workingvector
                ElseIf readMeastype = "S" Then
                    .SetSample readMeasnum, workingvector
                End If
                Set workingvector = Nothing
            End With
        Next r
        ReadUpMeasurements.Add workingBlock
        Set workingBlock = Nothing
    Else
        MsgBox "Sample not found in '.up' file!", vbCritical, "Critical Error!"
        Exit Function
    End If
End Function

'-----------------------------------------------------------------------------
'  WriteData
'
'  Description:       This function dumps the data given to a specified file.
'
'  Revision History:
'      Albert Hsiao        2/19/99       Formatted output.
'
Public Sub WriteData(demag As String, _
    gdec As Double, ginc As Double, sdec As Double, sinc As Double, _
    crdec As Double, crinc As Double, mmnt As Double, _
    errangle As Double, ByVal sdx As Double, ByVal sdy As Double, _
    ByVal sdz As Double, Optional ByVal UpToDn As Double = 0)
    Dim filenum As Integer
    Dim FilePath As String
    FilePath = SpecFilePath
    filenum = FreeFile
    If Not FileExists(FilePath) Then
        ' Create the file new if it doesn't exist
        Open FilePath For Output As #filenum
        Print #filenum, "original spec file not found"
        Print #filenum, "        ";
        Print #filenum, WriteNum(mvarCorePlateStrike); " ";
        Print #filenum, WriteNum(mvarCorePlateDip); " ";
        Print #filenum, WriteNum(mvarBeddingStrike); " ";
        Print #filenum, WriteNum(mvarBeddingDip); " ";
        Print #filenum, WriteNum(mvol)
    Else
        Open FilePath For Append As #filenum
    End If
    ErrorAngle = errangle
    Moment = mmnt
    UpDownRatio = UpToDn
    If ErrorAngle > 999 Then ErrorAngle = 999.9
    Print #filenum, demag$; " ";
    Print #filenum, WriteNum(gdec); " ";
    Print #filenum, WriteInc(ginc); " ";
    Print #filenum, WriteNum(sdec); " ";
    Print #filenum, WriteInc(sinc); " ";
    Print #filenum, Format$(Moment, "0.00E+00"); " ";
    Print #filenum, Format$(ErrorAngle, "000.0"); " ";
    Print #filenum, WriteNum(crdec); " ";
    Print #filenum, WriteInc(crinc); " ";
    Print #filenum, FormatNumber(sdx); " ";
    Print #filenum, FormatNumber(sdy); " ";
    Print #filenum, FormatNumber(sdz); " ";
    Print #filenum, Left$(LCase$(MailFromName), 8); " ";
    Print #filenum, Format(Now, "yyyy-mm-dd hh:mm:ss"); " " ' (August 2007 L Carporzen) Time added for VRM experiments
    Close #filenum
End Sub

Function WriteInc(num As Double) As String
'writes inc in a [sp]xx.x or -xx.x format
    If num > 99999 Then num = 99999  'Tame really bad ones.
    If num < -9999 Then num = -9999
    If Round(num, 1) >= 1000 Then  'Display oversize as integers
         WriteInc = Format(num, "00000")
    ElseIf Round(num, 1) <= -100 Then
         WriteInc = Format(num, "0000")
    ElseIf Round(num, 1) >= 100 Then
         WriteInc = Format(num, "000.0") 'Here for bad inc >90...
    ElseIf num >= 0 Then
         WriteInc = " " + Format(num, "00.0") 'Here for good + inc
    Else
         WriteInc = "-" + Format(-num, "00.0")   'or here for -inc.
    End If
End Function

Public Sub WriteLogFile(message As String)
    Dim filenum As Integer
    Dim fileid As Integer
    Dim LogFilePath As String
    filenum = FreeFile
    LogFilePath = Parent.SampleSetFilePathPrefix + ".log"
    If Not FileExists(LogFilePath) Then
        ' Create the file new if it doesn't exist
        Open LogFilePath For Output As #filenum
    Else
        Open LogFilePath For Append As #filenum
    End If
    message = Replace(message, vbCr, " ")
    message = Replace(message, vbLf, " ")
    If DEBUG_MODE Then frmDebug.Msg "Writing " + LogFilePath
    Print #filenum, Format(Now, "yyyy-mm-dd hh:mm:ss");
    Print #filenum, ": " & Samplename & ": ";
    Print #filenum, message
    Close #filenum
End Sub

Function WriteNum(ByVal num As Double) As String
    'If num >= 100000 Or num <= -10000 Then
        'WriteNum = "99999"
    'ElseIf num >= 10000 Or num <= -1000 Then
        'WriteNum = Format$(num, "####0")
    'ElseIf num >= 1000 Or num <= -100 Then
        'WriteNum = " " + Format$(num, "###0")
    'ElseIf num >= 100 Or num <= -10 Then
        'WriteNum = Format$(num, "##0.0")
    'ElseIf num >= 10 Or num < 0 Then
        'WriteNum = " " + Format$(num, "#0.0")
    'Else
        'WriteNum = "  " + Format$(num, "0.0")
    'End If
    'new code by Bogue 17 April 2006.  9.96 etc. bug
    'WriteNum stuffs any num into 5 spaces, with leading zeroes
    If num > 99999 Then num = 99999  'Tame really bad ones.
    If num < -9999 Then num = -9999
    If Round(num, 1) >= 1000 Then  'Display oversize as integers
         WriteNum = Format(num, "00000")
    ElseIf Round(num, 1) <= -100 Then
         WriteNum = Format(num, "0000")
    ElseIf num >= 0 Then
         WriteNum = Format(num, "000.0")  'Usually here +, <999..
    Else
         WriteNum = Format(num, "00.0")   'or here for -,>-99.
    End If
End Function

Public Sub WriteRockmagData(RMStep As RockmagStep, Optional magZ As Double = 0, Optional StdDevZ As Double = 0, Optional magX As Double = 0, Optional StdDevX As Double = 0, Optional magY As Double = 0, Optional StdDevY As Double = 0, Optional crdec As Double, Optional crinc As Double, Optional mmnt As Double, Optional errangle As Double, Optional SampleHeight As Double)
    Dim FilePath As String
    Dim filepathbackup As String
    Dim filenum As Integer
    FilePath = SpecFilePath & ".rmg"
    filenum = FreeFile
    WriteRockmagHeaders
    Open FilePath For Append As #filenum
    With RMStep
        Print #filenum, .StepType; ",";
        Print #filenum, .Level; ",";
        Print #filenum, .BiasField; ",";
        Print #filenum, .SpinSpeed; ",";
        Print #filenum, .HoldTime; ",";
        Print #filenum, magZ; ",";
        Print #filenum, StdDevZ; ",";
        Print #filenum, magZ / Vol; ",";
        Print #filenum, Susceptibility; ",";
        Print #filenum, magX; ",";
        Print #filenum, StdDevX; ",";
        Print #filenum, magY; ",";
        Print #filenum, StdDevY; ",";
        Print #filenum, .Remarks; ","; ' (November 2007 L Carporzen) Remarks column in RMG
        Print #filenum, crdec; ",";
        Print #filenum, crinc; ",";
        Print #filenum, mmnt; ",";
        Print #filenum, errangle; ",";
        Print #filenum, SampleHeight / UpDownMotor1cm; ",";
        Print #filenum, Now; ' (November 2007 L Carporzen) Date and time column in RMG
    End With
    Print #filenum, " "
    Close #filenum
    BackupRockmagData
End Sub

Private Sub WriteRockmagHeaders()
    Dim FilePath As String
    Dim filepathbackup As String
    Dim filenum As Integer
    FilePath = SpecFilePath & ".rmg"
    filenum = FreeFile
    If Not FileExists(FilePath) Then
        ' Create the file new if it doesn't exist
        Open FilePath For Output As #filenum
        Print #filenum, Samplename; ",";
        Print #filenum, Comment; ",";
        Print #filenum, "Vol: " & Vol; ",";
        Print #filenum, " "
        Print #filenum, " ,";
        Print #filenum, "Level"; ",";
        Print #filenum, "Bias Field (G)"; ",";
        Print #filenum, "Spin Speed (rps)"; ",";
        Print #filenum, "Hold Time (s)"; ",";
        Print #filenum, "Mz (emu)"; ",";
        Print #filenum, "Std. Dev. Z"; ",";
        Print #filenum, "Mz/Vol"; ",";
        Print #filenum, "Moment Susceptibility (emu/Oe)"; ",";
        Print #filenum, "Mx (emu)"; ",";
        Print #filenum, "Std. Dev. X"; ",";
        Print #filenum, "My (emu)"; ",";
        Print #filenum, "Std. Dev. Y"; ",";
        Print #filenum, "Remarks"; ","; ' (November 2007 L Carporzen) Remarks column in RMG
        Print #filenum, "Core Dec"; ",";
        Print #filenum, "Core Inc"; ",";
        Print #filenum, "M (emu)"; ",";
        Print #filenum, "CSD"; ",";
        Print #filenum, "Sample Height (cm)"; ",";
        Print #filenum, "Date/Time"; " " ' (November 2007 L Carporzen) Date and time column in RMG
        Close #filenum
    End If
End Sub

Public Sub WriteRockmagInfoLine(Description As String, Optional ByVal Level As Double = 0)
    Dim FilePath As String
    Dim filepathbackup As String
    Dim filenum As Integer
    FilePath = SpecFilePath & ".rmg"
    filenum = FreeFile
    WriteRockmagHeaders
    Open FilePath For Append As #filenum
        Print #filenum, Description; ",";
        Print #filenum, Level; ",";
        Print #filenum, " 0 ,";
        Print #filenum, " 0 ,";
        Print #filenum, " 0 ,";
        Print #filenum, " 0 ,";
        Print #filenum, Holder.Average.Z * RangeFact; ",";
        Print #filenum, " 0 ,";
        Print #filenum, SampleHolder.Susceptibility * SusceptibilityMomentFactorCGS; ",";
        Print #filenum, " 0 ,";
        Print #filenum, Holder.Average.X * RangeFact; ",";
        Print #filenum, " 0 ,";
        Print #filenum, Holder.Average.Y * RangeFact; ",";
        Print #filenum, " Holder ,";
        Print #filenum, " 0 ,";
        Print #filenum, " 0 ,";
        Print #filenum, Sqr(Abs(Holder.Average.X ^ 2 + Holder.Average.Y ^ 2 + Holder.Average.Z ^ 2)) * RangeFact; ",";
        Print #filenum, " 0 ,";
        Print #filenum, SampleHolder.SampleHeight; ",";
        Print #filenum, Now;
    Print #filenum, " "
    Close #filenum
End Sub

Public Sub WriteStatsTable(MData As MeasurementBlocks, _
    demag As String)
    Dim filenum As Integer
    Dim fileid As Integer
    Dim StatFilePath As String
    Dim key As String
    Dim outstring As String
    Dim i As Integer, cnt As Integer, j As Integer
    filenum = FreeFile
    StatFilePath = Parent.CurrentStepFilePathPrefix + ".stat"
    If Not FileExists(StatFilePath) Then
        ' Create the file new if it doesn't exist
        Open StatFilePath For Output As #filenum
        Print #filenum, "Sample" & "     Sig/Drift" & "     Sig/Holder" & "     Sig/Ind" & "     Sig/Noise" & "     CSD" & "     HorizErrAng" & "     Up/Down"
    Else
        Open StatFilePath For Append As #filenum
    End If
    If DEBUG_MODE Then frmDebug.Msg "Writing " + StatFilePath
    Print #filenum, Samplename;
    Print #filenum, "     "; WriteNum(MData.SigDrift);
    Print #filenum, "     "; WriteNum(MData.SigHolder);
    Print #filenum, "     "; WriteNum(MData.SigInduced);
    Print #filenum, "     "; WriteNum(MData.SigNoise);
    Print #filenum, "     "; WriteNum(MData.FischerSD);
    Print #filenum, "     "; WriteNum(MData.ErrorHorizontal);
    Print #filenum, "     "; Format$(MData.UpToDown, "0.000")
    Close #filenum
End Sub

'-----------------------------------------------------------------------------
'  WriteUpMeasurements
'
'  Description:       This function dumps the "up" data given to a temp file.
'
Public Sub WriteUpMeasurements(MData As MeasurementBlocks, _
    demag As String)
    Dim filenum As Integer
    Dim UpFilePath As String
    Dim i As Integer, cnt As Integer, j As Integer
    Dim initialString As String
    filenum = FreeFile
    UpFilePath = Parent.UpFilePath
    cnt = MData.Count
    If cnt = 0 Then Exit Sub
    If DEBUG_MODE Then frmDebug.Msg "Writing " + UpFilePath
    filenum = FreeFile
    If Not FileExists(UpFilePath) Then
        ' Create the file new if it doesn't exist
        Open UpFilePath For Output As #filenum
        Print #filenum, "Sample" & "|Direction" & "|Blocks" & "|MsmtType" & "|Block" & "|MsmtNum" & "|X,Y,Z"
    Else
        Open UpFilePath For Append As #filenum
    End If
    For i = 1 To cnt
        With MData.Item(i)
            If .isUp Then
                initialString = Samplename & "|U|" & Format$(cnt, "0") & "|"
            Else
                initialString = Samplename & "|D|" & Format$(cnt, "0") & "|"
            End If
            For j = 1 To 2
                Print #filenum, initialString;
                Print #filenum, "Z|" & Format$(i, "0") & "|" & Format$(j, "0") & "|";
                Print #filenum, .Baselines(j).WriteString;
                Print #filenum, "|" & Format(Now, "yyyy-mm-dd hh:mm:ss")
            Next j
            For j = 1 To 4
                Print #filenum, initialString;
                Print #filenum, "S|" & Format$(i, "0") & "|" & Format$(j, "0") & "|";
                Print #filenum, .Sample(j).WriteString;
                Print #filenum, "|" & Format(Now, "yyyy-mm-dd hh:mm:ss")
            Next j
            For j = 1 To 4
                Print #filenum, initialString;
                Print #filenum, "H|" & Format$(i, "0") & "|" & Format$(j, "0") & "|";
                Print #filenum, .Holder(j).WriteString;
                Print #filenum, "|" & Format(Now, "yyyy-mm-dd hh:mm:ss")
            Next j
        End With
    Next i
    Close #filenum
End Sub

