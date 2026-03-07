VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmFileSave 
   Caption         =   " "
   ClientHeight    =   6945
   ClientLeft      =   1560
   ClientTop       =   1395
   ClientWidth     =   6375
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6945
   ScaleWidth      =   6375
   Begin VB.Frame frameADWINAFBootDir 
      Caption         =   "ADWIN AF File Settings"
      Height          =   2775
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   6135
      Begin VB.TextBox txtADWINRampProgFile 
         Height          =   615
         Left            =   2520
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   22
         Top             =   1920
         Width           =   2775
      End
      Begin VB.CommandButton cmdADWINRampProgFile 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   5400
         TabIndex        =   21
         Top             =   1920
         Width           =   492
      End
      Begin VB.TextBox txtADWINBootFile 
         Height          =   495
         Left            =   2520
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   19
         Top             =   1200
         Width           =   2775
      End
      Begin VB.CommandButton cmdADWINBootFile 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   5400
         TabIndex        =   18
         Top             =   1200
         Width           =   492
      End
      Begin VB.TextBox txtADWINAFDir 
         Height          =   612
         Left            =   1800
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   15
         Top             =   360
         Width           =   3495
      End
      Begin VB.CommandButton cmdADWINAFDir 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   5400
         TabIndex        =   14
         Top             =   360
         Width           =   492
      End
      Begin VB.Label Label5 
         Caption         =   "ADWIN AF Ramp ADBasic Program File Name:"
         Height          =   735
         Left            =   240
         TabIndex        =   20
         Top             =   1920
         Width           =   2295
      End
      Begin VB.Label Label4 
         Caption         =   "ADWIN Boot (.btl) File Name:"
         Height          =   495
         Left            =   240
         TabIndex        =   17
         Top             =   1200
         Width           =   2175
      End
      Begin VB.Label Label3 
         Caption         =   "ADWIN AF Boot / Prog. Folder Path:"
         Height          =   495
         Left            =   240
         TabIndex        =   16
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame frameAFDataFileSettings 
      Caption         =   "AF Monitor Data File Settings"
      Height          =   2415
      Left            =   120
      TabIndex        =   5
      Top             =   3000
      Width           =   6135
      Begin VB.TextBox txtRampDataLocalFolder 
         Height          =   612
         Left            =   1800
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Top             =   360
         Width           =   3492
      End
      Begin VB.CommandButton cmdLocalFolderBrowse 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   5400
         TabIndex        =   9
         Top             =   360
         Width           =   492
      End
      Begin VB.TextBox txtRampDataBackupFolder 
         Height          =   612
         Left            =   1800
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Top             =   1200
         Width           =   3492
      End
      Begin VB.CommandButton cmbBackupFolderBrowse 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   5400
         TabIndex        =   7
         Top             =   1200
         Width           =   492
      End
      Begin VB.CheckBox chkBackupRampData 
         Caption         =   "Backup AF Data"
         Height          =   372
         Left            =   240
         TabIndex        =   6
         Top             =   1920
         Width           =   1932
      End
      Begin VB.Label Label1 
         Caption         =   "Local Ramp Data Main Folder Path:"
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Backup Ramp Data Main Folder Path:"
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   1200
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   372
      Left            =   3960
      TabIndex        =   4
      Top             =   5520
      Width           =   1092
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   372
      Left            =   2280
      TabIndex        =   3
      Top             =   5520
      Width           =   972
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Default         =   -1  'True
      Height          =   372
      Left            =   120
      TabIndex        =   2
      Top             =   5520
      Width           =   1212
   End
   Begin ComctlLib.ProgressBar progressFileSave 
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   6480
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin MSComDlg.CommonDialog dialogSaveFile 
      Left            =   5520
      Top             =   5520
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
   End
   Begin VB.Label lblFileSaveProgress 
      Caption         =   "File Save Progress:     % complete"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   6120
      Width           =   3735
   End
End
Attribute VB_Name = "frmFileSave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const SettingsOnlyHeight = 6495
Const SettingsOnlyWidth = 6495
Const ProgressHeight = 1440
Const ProgressWidth = 5415
Const ProgBarTop = 600
Const ProglblTop = 240
Const SettingsOnlyBarTop = 6600
Const SettingsOnlylblTop = 6840

Private Sub cmbBackupFolderBrowse_Click()

    Dim FolderPath As String
    Dim StartFolder As String
    Dim fso As FileSystemObject

    'Use File dialog object to allow the user to browse for and (if necessary)
    'create a new directory.  Allow user to create a new directory if they
    'want to.
    
    'First see if the backup file folder that's currently in
    'the local data folder path text box actually exists
    'if it doesn't, set the start folder to the C drive
    Set fso = New FileSystemObject
    If Not fso.FolderExists(Me.txtRampDataBackupFolder) Then
    
        StartFolder = "C:\"
    
    Else
    
        StartFolder = Me.txtRampDataBackupFolder
        
    End If

    'Call public shell function in modSaveFile that'll setup
    'all of the API properties that we need and don't really want to know about
    FolderPath = modFileSave.OpenDir(StartFolder, _
                        "Select AF Ramp Backup Data Folder", _
                        Me)

    'Now write the folderpath to the local data folder text box control
    Me.txtRampDataBackupFolder = FolderPath

    Set fso = Nothing

End Sub

Private Sub cmdADWINAFDir_Click()

    Dim FolderPath As String
    Dim StartFolder As String
    Dim fso As FileSystemObject

    'Use File dialog object to allow the user to browse for and (if necessary)
    'create a new directory.  Allow user to create a new directory if they
    'want to.
    
    'First see if the ADWIN Boot/.ini/Program file folder that's currently in
    'the folder path text box actually exists
    'if it doesn't, set the start folder to app.path (we know that folder must exist)
    Set fso = New FileSystemObject
    If Not fso.FolderExists(Me.txtADWINAFDir) Then
            
        StartFolder = App.path
        
    Else
    
        StartFolder = Me.txtADWINAFDir
        
    End If

    'Call public shell function in modSaveFile that'll setup
    'all of the API properties that we need and don't really want to know about
    FolderPath = modFileSave.OpenDir(StartFolder, _
                                     "Select ADWIN Bin and ADBasic Files Directory", _
                                     Me)

    'Make sure that a "\" is at the end of the path
    If Right(FolderPath, 1) <> "\" Then

        FolderPath = FolderPath & "\"
        
    End If
    
    'Now crawl through this directory and make sure the necessary files are present

    'Now write the folderpath to the ADWIN AF folder text box control
    Me.txtADWINAFDir = FolderPath
    
    'Deallocate file system object
    Set fso = Nothing
    
End Sub

Private Sub cmdADWINBootFile_Click()

    Dim fso As FileSystemObject
    Dim BootFilePath As String
    Dim TempL As Long

    Set fso = New FileSystemObject

    'Need to browse to the ADWIN boot file
    'And Open a dialog for the user to select the correct file to use
    
    'First check to see if the user has entered a valid folder for the ADWIN Bin Folder path
    If Not fso.FolderExists(Me.txtADWINAFDir) Then
    
        'No valid folder was entered
        'Send the user a message
        MsgBox "Please entered a valid ADWIN AF Boot & Program File Directory path before " & _
               "trying to set the ADWIN Boot File name.", , _
               "Ooops!"
               
        Exit Sub
        
    End If
    
    'Otherwise, the folder is valid and can be used as the start folder for the file
    'dialog
    With Me.dialogSaveFile
    
        .DefaultExt = ".btl"
        .DialogTitle = "Select the ADWIN Board Boot File"
        .FILTER = "Boot File (*.btl)|*.btl"
        .InitDir = Me.txtADWINAFDir
        .flags = cdlOFNFileMustExist
        .ShowOpen
        
        'Get the resulting file path
        BootFilePath = .filename
        
    End With
    
    'Now Parse that file path to get just the name of the file at the end
    TempL = InStrRev(BootFilePath, "\")
    Me.txtADWINBootFile = Trim(Mid(BootFilePath, TempL + 1))
    
End Sub

Private Sub cmdADWINRampProgFile_Click()

    Dim fso As New FileSystemObject
    Dim RampProgFilePath As String
    Dim TempL As Long

    'Need to browse to the ADWIN boot file
    'And Open a dialog for the user to select the correct file to use
    
    'First check to see if the user has entered a valid folder for the ADWIN Bin Folder path
    If Not fso.FolderExists(txtADWINAFDir) Then
    
        'No valid folder was entered
        'Send the user a message
        MsgBox "Please entered a valid ADWIN AF Boot & Program File Directory path before " & _
               "trying to set the ADWIN Ramp Program File name.", , _
               "Ooops!"
               
        Exit Sub
        
    End If
    
    'Otherwise, the folder is valid and can be used as the start folder for the file
    'dialog
    With Me.dialogSaveFile
    
        .DialogTitle = "Select the ADWIN Ramp Program File"
        .FILTER = "T91 Program File (*.T91)|*.T91|T93 Program File (*.T93)|*.T93"
        .InitDir = Me.txtADWINAFDir
        .flags = cdlOFNFileMustExist
        .ShowOpen
        
        'Get the resulting file path
        RampProgFilePath = .filename
        
    End With
    
    'Now Parse that file path to get just the name of the file at the end
    TempL = InStrRev(RampProgFilePath, "\")
    Me.txtADWINRampProgFile = Trim(Mid(RampProgFilePath, TempL + 1))

End Sub

Private Sub cmdApply_Click()

    Dim UserResp As Long
    Dim MsgStr As String

    'Need to save the file settings to the .INI file
    
    'Depending on which AF system is active, compose a custom insert into the
    'confirmation of change message
    If AFSystem = "2G" Then
    
        MsgStr = "AF Monitor File Save settings."
        
    Else
    
        MsgStr = "AF Monitor File Save and ADWIN AF Boot & Program file settings. " & _
                 "This could unintentionally disable your AF / IRM system " & _
                 "if the settings are incorrect."
    
    End If
    
    'Ask user if they want to proceed with this change
    UserResp = MsgBox("You are about to make major changes to the " & MsgStr & _
                      vbNewLine & vbNewLine & "Are you sure you want to do this?", _
                      vbYesNo, _
                      "Warning!!")
                      
    If UserResp <> vbYes Then
    
        'User doesn't want to go through with the changes
        Exit Sub
        
    End If
    
    'User does want to go on! (Yay!)
    ExportSettings
    
    'Now hide this form
    Me.Hide
        
End Sub

Private Sub cmdClose_Click()

    Unload Me
    Me.Hide

End Sub

Private Sub cmdLocalFolderBrowse_Click()

    Dim FolderPath As String
    Dim StartFolder As String
    Dim fso As FileSystemObject
    
    'Use File dialog object to allow the user to browse for and (if necessary)
    'create a new directory.  Allow user to create a new directory if they
    'want to.
    
    'First see if the local file folder that's currently in
    'the local data folder path text box actually exists
    'if it doesn't, set the start folder to the C drive
    Set fso = New FileSystemObject
    If Not fso.FolderExists(Me.txtRampDataLocalFolder) Then
    
        StartFolder = "C:\"
    
    Else
    
        StartFolder = Me.txtRampDataLocalFolder
        
    End If
    
    'Call public shell function in modSaveFile that'll setup
    'all of the API properties that we need and don't really want to know about
    FolderPath = modFileSave.OpenDir(StartFolder, _
                        "Select AF Ramp Data Folder", _
                        Me)

    'Now write the folderpath to the local data folder text box control
    Me.txtRampDataLocalFolder = FolderPath

''-----------------------------------------------------------------------------------
'   '(Mar 2010 - I Hilburn)
'   'Commented out code -
'   'This code below uses the BrowseForFolder function in modFileSave
'   'to run a Directory browser dialog by making direct calls to the
'   'appropriate MS Windows .dll files.
'   'This code DOES NOT allow the user to create a new directory
'   'And thus is not as good as the code above
'   'Plus, directory browser blow major monkey chunks
'
'    FolderPath = modFileSave.BrowseForFolder(frmFileSave, _
'                                            "Select Main Ramp Data Folder", _
'                                            StartFolder)
'
'    If Len(FolderPath) = 0 Then
'
'        'User Cancelled and did not select a folder
'        'exit the sub-routine
'
'        Exit Sub
'
'    End If
'
'    'Add a "\" to the end of the folder path if there isn't a terminal "\" already
'    If Right(FolderPath, 1) <> "\" Then
'
'        FolderPath = FolderPath & "\"
'
'    End If
'
'    'user didn't cancel, save folder path to local folder path text box
'    Me.txtRampDataLocalFolder = FolderPath
''--------------------------------------------------------------------------------------
    
    Set fso = Nothing
    
End Sub

Private Sub cmdOK_Click()

    'Depending on which AF system is active, compose a custom insert into the
    'confirmation of change message
    If AFSystem = "2G" Then
    
        MsgStr = "AF Monitor File Save settings."
        
    Else
    
        MsgStr = "AF Monitor File Save and ADWIN AF Boot & Program file settings. " & _
                 "This could unintentionally disable your AF / IRM system " & _
                 "if the settings are incorrect."
    
    End If
    
    'Ask user if they want to proceed with this change
    UserResp = MsgBox("You are about to make major changes to the " & MsgStr & _
                      vbNewLine & vbNewLine & "Are you sure you want to do this?", _
                      vbYesNo, _
                      "Warning!!")
                      
    If UserResp <> vbYes Then
    
        'User doesn't want to go through with the changes
        Exit Sub
        
    End If
    
    'Apply the settings to save them in the System Global variables
    ExportSettings
    
    'Now save the settings to the INI file
    modConfig.Config_writeSettingstoINI

    'Now read the settings from the INI file
    modConfig.Config_ReadINISettings
    
    'Now Hide this form
    Me.Hide

End Sub

Private Sub ExportSettings()

    'Save all the settings on this form
    If AFSystem = "2G" Then
    
        'Save the local & backup dir info to the system global variables
        modConfig.TWOG_AFDataLocalDir = Trim(Me.txtRampDataLocalFolder)
        modConfig.TWOG_AFDataBackupDir = Trim(Me.txtRampDataBackupFolder)
        
    Else
    'User has set the ADWIN system to be used
    
        'Save the local & backup dir info to the system global variables
        modConfig.ADWIN_AFDataLocalDir = Trim(Me.txtRampDataLocalFolder)
        modConfig.ADWIN_AFDataBackupDir = Trim(Me.txtRampDataBackupFolder)
    
        'Save the ADWIN Boot & Program directory and file settings as well
        modConfig.ADWINBinFolderPath = Trim(Me.txtADWINAFDir)
        modConfig.ADWINBootFileName = Trim(Me.txtADWINBootFile)
        modConfig.ADWINRampProgFileName = Trim(Me.txtADWINRampProgFile)
        
        'Save the ADWIN path info to the ADWIN module globals
        ADWIN.BinFolderPath = ADWINBinFolderPath
        ADWIN.BootFileName = ADWINBootFileName
        ADWIN.CurProcessFile = ADWINRampProgFileName
        
    End If
    
    'Save the Do Backup status
    If Me.chkBackupRampData.value = Checked Then
    
        modConfig.AFDoDataFileBackup = True
        
    Else
    
        modConfig.AFDoDataFileBackup = False
        
    End If

End Sub

Private Sub Form_Activate()
    
    LoadFileSaveForm
    
End Sub

Private Sub Form_Load()

    LoadFileSaveForm
    
End Sub

Public Function LoadAFCalibrationTable _
    (ByRef CalTable As MSHFlexGrid, _
     Optional ByVal Units As String = "G") As Boolean

    Dim i, j As Long
    Dim SearchResult As Long
    Dim NumCols As Long
    Dim NumReplicates As Long
    Dim NumRows As Long
        
    Dim fso As FileSystemObject
    Dim FileStream As TextStream
    Dim FilePath As String
    Dim CalibrationFolder As String
        
    Dim TempString As String
    Dim TempString2 As String
    Dim RowStrArray() As String
    
    Dim ErrorMessage As String
    
    Set fso = New FileSystemObject
    
    'Check to make sure the local data folder exists
    If fso.FolderExists(Me.txtRampDataLocalFolder) = False Then
    
        'Create the folder
        On Error Resume Next
            
            fso.CreateFolder (Me.txtRampDataLocalFolder)
            
            If Err.number <> 0 Then
            
                SetCodeLevel CodeRed
            
                ErrorMessage = "Unable to create local data folder: " & vbNewLine & _
                               vbTab & Me.txtRampDataLocalFolder & vbNewLine & _
                               vbNewLine & "Please check / change folder path in the " & _
                               "AF File Save Settings window." & vbNewLine & vbNewLine & _
                               "Code execution has been paused."
            
                'Raise Err - unable to create folder
                frmSendMail.MailNotification _
                            "Data Folder Error", _
                            ErrorMessage, _
                            CodeRed, _
                            True
                
                MsgBox ErrorMessage
                
                frmProgram.SetProgramCodeLevel StatusCodeColorLevelPrior
                
                'Return a false to indicate that loading the table has failed
                LoadAFCalibrationTable = False
                
                Exit Function
                
            End If
            
        On Error GoTo 0
                          
    End If
    
    'Check to see if the calibration folder already exists
    CalibrationFolder = "Calibration Values/"
    
    If fso.FolderExists(Me.txtRampDataLocalFolder & CalibrationFolder) = False Then
    
        'Clear out the Calibration folder name
        CalibrationFolder = ""
        
    End If
    
    'Prompt the user to load the calibration data file, starting with
    'the calibration folder name
    Me.dialogSaveFile.InitDir = Trim(Me.txtRampDataLocalFolder & CalibrationFolder)
    Me.dialogSaveFile.DefaultExt = "*.csv"
    Me.dialogSaveFile.DialogTitle = "Open AF Calibration .CSV file"
    Me.dialogSaveFile.FILTER = "(*.csv)|*.csv|commas-separated values file"
    Me.dialogSaveFile.ShowOpen
    
    FilePath = Me.dialogSaveFile.filename
    
    If FilePath = "" Then
    
        'If no path returned by dialog, exit the function with
        'a failed load value returned
        LoadAFCalibrationTable = False
        
        Exit Function
        
    End If
    
    'Show only the progress label and progress bar
    Me.Height = ProgressHeight
    Me.frameADWINAFBootDir.Visible = False
    Me.lblFileSaveProgress.Top = ProglblTop
    Me.progressFileSave.Top = ProgBarTop
    
    'Setup the progress bar
    Me.progressFileSave.Max = 32767
    Me.progressFileSave.min = 0
    Me.progressFileSave.value = 0
    
    'Update the Cation on the progress label
    Me.lblFileSaveProgress.Caption = "File Load Progress:"
    
    'Refresh the form
    Me.Show
    Me.refresh
    
    Set FileStream = fso.OpenTextFile(FilePath, ForReading)
    
    'Skip by the first four lines of the file
    For i = 1 To 4
    
        FileStream.SkipLine
        
        'Update the file read progres
        
    Next i
    
    'Get the first line of data
    TempString = FileStream.ReadLine
    
    'Set j = 1, start of TempString
    j = 1
    
    'Set i = comma count = 0
    i = 0
    
    'Figure out how many commas (data elements) are in this first line (row) of data
    Do
    
        'Search the 1st line of data for commas
        SearchResult = InStr(j, TempString, ",")
        
        If SearchResult > 0 Then
        
            'Update j with position of new comma + 1
            j = SearchResult + 1
            
            'Increment i - one more comma has been found
            i = i + 1
            
        End If
        
    'If no new comma is found, end the do loop
    Loop Until SearchResult = 0
    
    'The # of columns = # of commas + 2 - 2 (one extra column for the line number,
    'minus two columns used in the text file for the average monitor voltage
    'and it's standard deviation)
    NumCols = i
    
    'Size the Row-string array = NumCols+2 - 1 (one column skipped for the line number,
    'plus two additional data entries in the row-string array for the average voltage
    'and it's standard deviation
    ReDim RowStrArray(NumCols + 1)
        
    'The # of replicate measurements = (NumCols - 4) \ 2
    NumReplicates = (NumCols - 4) \ 2
    
    'Now need to resize the MSH grid object with: 2 rows x NumCols
    CalTable.ClearStructure
    CalTable.Rows = 2
    CalTable.Cols = NumCols
    
    'Now load headers for the first row
    With CalTable
    
        'Write in the Column Headers
        .row = 0
        .Col = 1
        .text = "Monitor Voltage"
        .RowSizingMode = flexRowSizeIndividual
        .RowHeight(0) = 456
           
        .Col = 2
        .text = "Field (" & Units & ")"
            
        .Col = 3
        .text = "StDev (" & Units & ")"
    
        For i = 2 To 2 + NumReplicates * 2 - 1 Step 2
    
            .Col = i + 2
            .text = "Mon. Field #" & Trim(str(i \ 2)) & " (" & Units & ")"
    
            .Col = i + 3
            .text = "Max Volt. #" & Trim(str(i \ 2)) & " (V)"
            
        Next i
        
        'Set i = row counter = 2
        i = 2
        
        Do
            
            'First line of data has already been read in, don't want to skip it
            If i > 2 Then
                
                'Read in the next line of data
                TempString = FileStream.ReadLine
                
            End If
            
            'Split the line string up into an array whose elements
            'contain each data column element of the new row in the calibration table
            RowStrArray = Split(TempString, ",")
            
            'Check to see if the Calibration table needs to be resized
            If .Rows < i Then .Rows = i
            
            'Write in the number of the current data row into
            'the 0th column
            .row = i - 1
            .Col = 0
            .text = Trim(str(i - 1))
            
            'Run through the first three column entries in the RowStrArray
            'and write them to the appropriate column of this row
            'of the calibration table
            For j = 0 To 2
        
                'Set column
                .Col = j + 1
                
                'Write in text - convert file text to val then back to string
                'to get rid of unnecessary zeroes padded on the left-hand side
                'of the data
                .text = Trim(str(val(RowStrArray(j))))
                
            Next j
            
            'Skipping elements 3 & 4, go through the rest of the RowStr Array
            'and write it to columns 3 through NumCols - 1 of the calibration table
            For j = 5 To NumCols
            
                'Set column
                .Col = j - 1
                
                .text = Trim(str(val(RowStrArray(j))))
            
            Next j
            
            'Iterate i
            i = i + 1
            
            DoEvents
            
        'Continue going through file and creating new rows until
        'we've run all the way to the end of the file
        Loop Until FileStream.AtEndOfStream = True
    
    End With
    
    'Hide the filesave form
    Me.Hide
    
    'Reload the file save form so that the bottom is hidden and the top is shown
    Load frmFileSave
    
    LoadAFCalibrationTable = True
        
End Function

Private Sub LoadFileSaveForm()

    
    'Make Window size such that the file save progress
    'label and progress bar are not shown
    Me.Height = SettingsOnlyHeight
    Me.Width = SettingsOnlyWidth
    Me.Top = 500
    Me.Left = 500
    
    '(March 11, 2011 - I Hilburn)
    'Update all the form's icons with the program Ico file
    SetFormIcon Me
    
    'Move the file save progress bar and label
    'to the undisplayed bottom of the form
    Me.progressFileSave.Top = SettingsOnlyBarTop
    Me.lblFileSaveProgress.Top = SettingsOnlylblTop
        
    'Set the visibility status of all the frames to True
    Me.frameADWINAFBootDir.Visible = True
    Me.frameAFDataFileSettings.Visible = True
       
    'Check to see which AF system is in use
    If AFSystem = "2G" Then
        
        'Set the default values for the various fields
        Me.txtRampDataLocalFolder = modConfig.TWOG_AFDataLocalDir
        Me.txtRampDataBackupFolder = modConfig.TWOG_AFDataBackupDir
        
    Else
    
        'Set the default values for the various fields
        Me.txtRampDataLocalFolder = modConfig.ADWIN_AFDataLocalDir
        Me.txtRampDataBackupFolder = modConfig.ADWIN_AFDataBackupDir
    
    End If
        
    'Checked see if do backup is checked
    If modConfig.AFDoDataFileBackup = True Then
    
        Me.chkBackupRampData.value = Checked
        
    Else
    
        Me.chkBackupRampData.value = Unchecked
        
    End If
    
    'Load the values into the ADWIN file and folder path settings
    Me.txtADWINAFDir = modConfig.ADWINBinFolderPath
    Me.txtADWINBootFile = modConfig.ADWINBootFileName
    Me.txtADWINRampProgFile = modConfig.ADWINRampProgFileName
    
    'Depending on the AF System, enable the ADWIN Boot Dir controls and frame
    Me.frameADWINAFBootDir.Enabled = (AFSystem = "ADWIN")
    Me.txtADWINAFDir.Enabled = (AFSystem = "ADWIN")
    Me.cmdADWINAFDir.Enabled = (AFSystem = "ADWIN")
    Me.txtADWINBootFile.Enabled = (AFSystem = "ADWIN")
    Me.cmdADWINBootFile.Enabled = (AFSystem = "ADWIN")
    Me.txtADWINRampProgFile.Enabled = (AFSystem = "ADWIN")
    Me.cmdADWINRampProgFile.Enabled = (AFSystem = "ADWIN")
        
    'Enable all of the AF Data file save display controls
    Me.txtRampDataBackupFolder.Enabled = True
    Me.txtRampDataLocalFolder.Enabled = True
    Me.chkBackupRampData.Enabled = True
    Me.cmdLocalFolderBrowse.Enabled = True
    Me.cmbBackupFolderBrowse.Enabled = True

End Sub

Public Function MultiRampFileSave_ADWIN _
    (ByRef AFData() As Double, _
     ByVal TimeStep As Double, _
     ByVal PtsPerFile As Long, _
     ByVal FolderName As String, _
     ByVal CurTime, _
     ByRef SineFit_Data() As Double, _
     Optional WriteToBackup As Boolean = False, _
     Optional SaveMaxAmp As Boolean = False, _
     Optional PtsWindowForFindingMax As Long = 0) As Boolean
                                
    Dim i, ii, j, k, N, N_out, N_in, DataPoints As Long
    Dim NumBackupFiles As Long
    Dim NumCols As Long
    
    Dim TempTime As Double
    Dim TempV As Double
    Dim MaxAmpIn As Double
    Dim MaxAmpOut As Double
    Dim MaxField As Double
    Dim TempIn As Double
    Dim TempOut As Double
    Dim TempField As Double
    
    Dim loopDone As Boolean
    
    Dim fso As FileSystemObject
    Dim DataStream As TextStream
    Dim MaxAmpStream As TextStream
    Dim SineFitStream As TextStream
    Dim MaxAmpFileName As String
    Dim SineFitFileName As String
    Dim DataFileName As String
    
    Dim PercentDone As String
    Dim TempS As String
    
    Dim DataFolder As Folder
    Dim ErrorMessage As String
    
    'Show only the progress label and progress bar
    'Me.Height = ProgressHeight
    'Me.Width = ProgressWidth
    'Me.frameADWINAFBootDir.Visible = False
    'Me.lblFileSaveProgress.Top = ProglblTop
    'Me.progressFileSave.Top = ProgBarTop
    
    'Set Percent done to 0%
    PercentDone = "  0.0%"
    
    'Make File Save Progress bar visible
    'Me.progressFileSave.min = 0
    'Me.progressFileSave.Max = 32767
    'Me.progressFileSave.value = 0
    'Me.progressFileSave.Visible = True
    'Me.lblFileSaveProgress.Caption = "File Save % Complete: " & PercentDone
    'Me.lblFileSaveProgress.Visible = True
    
    'Update program form status bar
    frmProgram.StatusBar "Saving... " & PercentDone, 3
    
    'Show file save window
    'Me.Show
    'Me.refresh
    
    Set fso = New FileSystemObject
    If Not fso.FolderExists(modConfig.ADWIN_AFDataLocalDir) Then
    
        On Error GoTo BadLocalMainFolder:
        
            fso.CreateFolder (modConfig.ADWIN_AFDataLocalDir)
            
        On Error GoTo 0
    
    End If
    
    If Not fso.FolderExists(modConfig.ADWIN_AFDataLocalDir & FolderName) Then
    
        On Error GoTo BadLocalFolderName:
        
            fso.CreateFolder (modConfig.ADWIN_AFDataLocalDir & FolderName)
            
        On Error GoTo 0
    
    End If
    
    If ActiveCoilSystem = AxialCoilSystem Then
    
        CoilString = "Axial"
        
    ElseIf ActiveCoilSystem = TransverseCoilSystem Then
    
        CoilString = "Transverse"
        
    End If
    
    'Get the number of rows = # of data points to save
    DataPoints = UBound(AFData, 1)
    
    'Get the number of columns to determine if the field level
    'needs to be saved as well
    NumCols = UBound(AFData, 2)
    
    On Error Resume Next
        
        N = UBound(SineFit_Data, 1)
    
        If Err.number = 0 Then
            
            On Error GoTo 0
        
            If N > 1 Or SineFit_Data(0, 0) <> -1 Then
            
                SineFitFileName = "SineFits_" & Format(CurTime, "MM-DD-YY_HH-MM-SS") & ".csv"
                
                On Error GoTo SineFitFileCreateError:
            
                    Set SineFitStream = fso.CreateTextFile(modConfig.ADWIN_AFDataLocalDir & FolderName & SineFitFileName, True)
                    
                On Error GoTo 0
                
                SineFitStream.WriteLine AFSystem & " AF Ramp Sine Fits on " & CoilString & " coil"
                SineFitStream.WriteLine Format(CurTime, "long date") & ", " & Format(CurTime, "long time")
                SineFitStream.WriteBlankLines (1)
                TempS = "Pt. #,Time (ms),"
                
                'If NumCols = 3, then need to insert the Calc. Field header in here
                If NumCols = 3 Then
                
                    'Add in the units as well from the global AF Units string variable
                    TempS = TempS & "Calc. Field (" & modConfig.AFUnits & "),"
                    
                End If
                
                'Now add the rest of the headers
                TempS = TempS & "Mon. Amp, Fit Amp, Actual Freq, Fit Freq," & _
                                "Y-offset,Phase,RMS,Output Amp"
                
                'Write the header linestring to file
                SineFitStream.WriteLine TempS
                
                For i = 0 To N - 1
                
                    'Add in the data point # and time fields
                    TempS = Format(SineFit_Data(i, 0), "0") & "," & _
                            Trim(str(SineFit_Data(i, 1) * 1000)) & ","
                          
                    'If NumCols = 3, then add in the Calculated Field value here
                    If NumCols = 3 Then
                    
                        TempS = TempS & Trim(str(SineFit_Data(i, 10))) & ","
                        
                    End If
                    
                    'Now add the rest of the data fields
                    TempS = TempS & Trim(str(SineFit_Data(i, 2))) & "," & _
                                    Trim(str(SineFit_Data(i, 3))) & "," & _
                                    Trim(str(SineFit_Data(i, 4))) & "," & _
                                    Trim(str(SineFit_Data(i, 5))) & "," & _
                                    Trim(str(SineFit_Data(i, 6))) & "," & _
                                    Trim(str(SineFit_Data(i, 7))) & "," & _
                                    Trim(str(SineFit_Data(i, 8))) & "," & _
                                    Trim(str(SineFit_Data(i, 9)))
                                    
                    'Write the data line-string to file
                    SineFitStream.WriteLine TempS
                                            
                    'Need to update progress
                    'Me.progressFileSave.value = CInt(32767 * i / (N + DataPoints - 1))
                    PercentDone = Format(100 * i / (N + DataPoints - 1), "#0.0")
                    
                    
                    'Pad percent done with up to two whitespaces
                    PercentDone = PadLeft(PercentDone, 5) & "%"
                    
                    'Update file save window
                    'Me.lblFileSaveProgress.Caption = "File Save % Complete: " & PercentDone
                    'Me.refresh
                    
                    'Update frmProgram status bar
                    frmProgram.StatusBar "Saving... " & PercentDone, 3
                    
                    DoEvents
                
                Next i
                
                SineFitStream.Close
                
            End If
        
        End If
        
    On Error GoTo 0
        
    If SaveMaxAmp = True Then
    
        MaxAmpFileName = "AFRamplitude_" & Format(CurTime, "MM-DD-YY_HH-MM-SS") & ".csv"
    
        On Error GoTo MaxAmpFileCreateError:
    
            Set MaxAmpStream = fso.CreateTextFile(modConfig.ADWIN_AFDataLocalDir & FolderName & MaxAmpFileName, True)
            
        On Error GoTo 0
        
        MaxAmpStream.WriteLine AFSystem & " AF Ramp Amplitudes on " & CoilString & " coil"
        MaxAmpStream.WriteLine Format(CurTime, "long date") & "," & Format(CurTime, "long time")
        MaxAmpStream.WriteBlankLines (1)
        MaxAmpStream.WriteLine "Sliding Window =" & "," & Trim(str(PtsWindowForFindingMax))
        
        'Add headers for the data point # and time columns
        TempS = "Point #,Time (ms),"
        
        'If NumCols = 3, need to add header for Calculated Field column
        If NumCols = 3 Then
        
            'Add in units of the field as well from global AF units var.
            TempS = TempS & "Calc. Field (" & modConfig.AFUnits & "),"
            
        End If
        
        'Add in the two voltage fields
        TempS = TempS & "Monitor V, Ramp V"
    
        'Now write the header line-string to file
        MaxAmpStream.WriteLine TempS
    
    Else
    
        Set MaxAmpStream = Nothing

    End If
    
    'Create first data file
    DataFileName = "AFramp_pts0-" & Trim(str(PtsPerFile)) & "_" & Format(CurTime, "MM-DD-YY_HH-MM-SS") & ".csv"
    
    On Error GoTo DataFileCreateError:
    
        Set DataStream = fso.CreateTextFile(modConfig.ADWIN_AFDataLocalDir & FolderName & DataFileName, True)
        
    On Error GoTo 0
           
    DataStream.WriteLine AFSystem & " AF Ramp on " & CoilString & " coil"
    DataStream.WriteLine Format(CurTime, "long date") & ", " & Format(CurTime, "long time")
    DataStream.WriteBlankLines (1)
    DataStream.WriteLine "From = ,0"
    DataStream.WriteLine "To = ," & Trim(str(PtsPerFile))
        
    'Add headers for the data point # and time columns
    TempS = "Point #,Time (ms),"
    
    'If NumCols = 3, need to add header for Calculated Field column
    If NumCols = 3 Then
    
        'Add in units of the field as well from global AF units var.
        TempS = TempS & "Calc. Field (" & modConfig.AFUnits & "),"
        
    End If
    
    'Add in the two voltage fields
    TempS = TempS & "Monitor V, Ramp V"
    
    'Write header line-string to file
    DataStream.WriteLine TempS
    
    j = PtsWindowForFindingMax
    
    'Initialize k = # pts per file + 1
    k = PtsPerFile + 1
    
    MaxAmpIn = 0
    MaxAmpOut = 0
    
    'Initialize ii at zero, ii is the counter for the Ramp out-voltage points
    ii = 0
    
    For i = 0 To DataPoints - 1
    
        DoEvents
    
        If i Mod 5000 = 0 Then
        
            'Need to update progress
            'Me.progressFileSave.value = CInt(32767 * (i + N) / (N - 1 + DataPoints))
            PercentDone = Format(100 * (i + N) / (N - 1 + DataPoints), "#0.0")
            
            
            'Pad percent done with up to two whitespaces
            PercentDone = PadLeft(PercentDone, 5) & "%"
            
            'Update file save window
            'Me.lblFileSaveProgress.Caption = "File Save % Complete: " & PercentDone
            'Me.refresh
            
            'Update frmProgram status bar
            frmProgram.StatusBar "Saving... " & PercentDone, 3
            
    
        End If
    
        'Store the current Monitor input voltage and time
        TempTime = i * TimeStep
        TempIn = AFData(i, 0)
        TempOut = AFData(i, 1)
        
        'If NumCols = 3, also store the Field value to temp local var.
        If NumCols = 3 Then
            
            TempField = AFData(i, 2)
            
        End If
            
        'Initialize loop done to false
        loopDone = False
        
        'Search for and save max amplitudes for input and output and field points
        If SaveMaxAmp = True Then
        
            If Abs(TempIn) > MaxAmpIn Then MaxAmpIn = Abs(TempIn)
            If Abs(TempOut) > MaxAmpOut Then MaxAmpOut = Abs(TempOut)
                    
            If NumCols = 3 And _
               Abs(TempField) > MaxField _
            Then MaxField = Abs(TempField)
                    
        End If
        
        'Add in the data pt. # and time
        TempS = Trim(str(i)) & "," & _
                Trim(str(TempTime)) & ","
                                 
        'If NumCols = 3, then add in the Calculated field data
        If NumCols = 3 Then
        
            TempS = TempS & Trim(str(TempField)) & ","

        End If
        
        'Now add in the two voltage data values
        TempS = TempS & Trim(str(TempIn)) & "," & _
                        Trim(str(TempOut))
            
        'Write the data line-string to file
        DataStream.WriteLine TempS
                                 
        'Decrement the period and file remaining points counters
        j = j - 1
        k = k - 1
        
        If j = 0 And SaveMaxAmp = True Then
        
            'Add in the data point # and time
            TempS = Trim(str(i)) & "," & _
                    Trim(str(i * TimeStep)) & ","
                    
            'If NumCols = 3, then add in the Calculated max field value
            If NumCols = 3 Then
            
                TempS = TempS & Trim(str(MaxField)) & ","
                    
            End If
                    
            'Add in the two Max voltage values
            TempS = TempS & Trim(str(MaxAmpIn)) & "," & _
                            Trim(str(MaxAmpOut))
            
            'Write data line-string to file
            MaxAmpStream.WriteLine TempS
            
            j = PtsWindowForFindingMax
            
            MaxAmpIn = 0
            MaxAmpOut = 0
            MaxField = 0
            
        End If
        
        If k = 0 And Not i = DataPoints - 1 Then
        
            'Need to close the current text stream, create a new text file,
            'and open it up for writing
            DataStream.Close
            
            'Create next data file
            'Adjust name to refltect the current point we're on and adjust the
            'point to for the final file
            If i + PtsPerFile > DataPoints - 1 Then
                
                DataFileName = "AFramp_pts" & Trim(str(i + 1)) & "-" & Trim(str(DataPoints - 1)) & _
                                "_" & Format(CurTime, "MM-DD-YY_HH-MM-SS") & ".csv"
                                
            Else
            
                DataFileName = "AFramp_pts" & Trim(str(i + 1)) & "-" & Trim(str(i + PtsPerFile)) & _
                                "_" & Format(CurTime, "MM-DD-YY_HH-MM-SS") & ".csv"
                                
            End If
            
            On Error GoTo DataFileCreateError:
            
                Set DataStream = fso.CreateTextFile(modConfig.ADWIN_AFDataLocalDir & FolderName & DataFileName, True)
                
            On Error GoTo 0
            
            DataStream.WriteLine AFSystem & " AF Ramp on " & CoilString & " coil"
            DataStream.WriteLine Format(CurTime, "long date") & ", " & Format(CurTime, "long time")
            DataStream.WriteBlankLines (1)
            DataStream.WriteLine "From = ," & Trim(str(i + 1))
            DataStream.WriteLine "To = ," & Trim(str(i + PtsPerFile))
            
            'Add headers for the data point # and time columns
            TempS = "Point #,Time (ms),"
            
            'If NumCols = 3, need to add header for Calculated Field column
            If NumCols = 3 Then
            
                'Add in units of the field as well from global AF units var.
                TempS = TempS & "Calc. Field (" & modConfig.AFUnits & "),"
                
            End If
            
            'Add in the two voltage fields
            TempS = TempS & "Monitor V, Ramp V"
            
            'Write header line-string to file
            DataStream.WriteLine TempS
            
            k = PtsPerFile
            
        End If
        
    Next i
    
    'Close the final file string
    DataStream.Close
                
    If WriteToBackup And _
       Me.chkBackupRampData.value = Checked _
    Then
    
        'Now need to copy the ramp files to the remote backup path
                
        'Change Value on Progress bar to 0
        'Me.progressFileSave.value = 0
        
        'Determine the number of files to backup
        NumBackupFiles = CInt(N / PtsPerFile) + 1
        
        If SaveMaxAmp = True Then NumBackupFiles = NumBackupFiles + 1
        
        'Change Caption in lblFileSaveprog
        'Me.lblFileSaveProgress.Caption = "Creating Backup Folder..."
        'Me.refresh
        
        frmProgram.StatusBar "Creating Backup Folder...", 3
        
        'See if there is a main backup ramp data folder yet,
        'if not, create it
        If Not fso.FolderExists(modConfig.ADWIN_AFDataBackupDir) Then
        
            On Error GoTo BadBackupMainFolder:
            
                fso.CreateFolder modConfig.ADWIN_AFDataBackupDir
                
            On Error GoTo 0
            
        End If
        
        'Create the folder for this AF ramp's worth of data files
        If Not fso.FolderExists(modConfig.ADWIN_AFDataBackupDir & FolderName) Then
        
            On Error GoTo CreateBackupFolderError:
            
                fso.CreateFolder modConfig.ADWIN_AFDataBackupDir & FolderName
        
            On Error GoTo 0
            
        End If
        
        'Change Caption on lblFileSaveProg
        'Me.lblFileSaveProgress.Caption = "Backing up Files: " & _
                                            Trim(str(NumBackupFiles)) & " Files remaining..."
        'Me.refresh
                
        PercentDone = "  0.0%"
                
        'Change caption on frmProgram status bar
        frmProgram.StatusBar "Backing up... " & PercentDone, 3
        
                
        'Navigate the local AF Ramp data folder, and copy the files in the
        'folder to the backup folder one by one
        Set DataFolder = fso.GetFolder(modConfig.ADWIN_AFDataLocalDir & FolderName)
                
               
        N = DataFolder.Files.Count
                
        For i = 1 To N
        
            DoEvents
        
            With DataFolder.Files.Item(i)
            
                On Error GoTo BadBackupFileCopy:
            
                    .COPY modConfig.ADWIN_AFDataBackupDir & FolderName & .name, True
                
                On Error GoTo 0
                
                Me.lblFileSaveProgress.Caption = "Backing up Files: " & _
                                                    Trim(str(NumBackupFiles - i)) & _
                                                    " Files remaining..."
                
                'Update file save form bar
                'Me.progressFileSave.value = CInt(32767 * i / N)
                'Me.refresh
                
                'Calculate percent done
                PercentDone = Format(CInt(i / NumBackupFiles), "#0.0")
                
                'Pad Percentdone
                PercentDone = PadLeft(PercentDone, 5) & "%"
                
                'Update program for status bar text
                frmProgram.StatusBar "Backing up... " & PercentDone, 3
                
            End With
            
        Next i
        
    End If
                
    'Pause a half-second (500 ms)
    PauseTill timeGetTime() + 500
    
    'Clear the third panel of the program status bar
    frmProgram.StatusBar vbNullString, 3
                
   'Uncheck the Monitor & Save AF Ramp box
    frmADWIN_AF.chkVerbose.value = Unchecked

    MultiRampFileSave = True
    
    'Hide the filesave form
    Me.Hide
    
    'Reload the file save form so that the bottom is hidden and the top is shown
    Load frmFileSave

    Exit Function
    
BadLocalMainFolder:

    ErrorMessage = "Could not find/access AF Ramp main data folder. Code Execution paused." & vbNewLine & _
                    "Current path = " & modConfig.ADWIN_AFDataLocalDir & vbNewLine & _
                    "Error Code = " & Trim(str(Err.number))
    
    SetCodeLevel CodeRed
    
    frmSendMail.MailNotification "Folder Path Error", _
                                 ErrorMessage, _
                                 CodeRed, _
                                 True
    
    MsgBox ErrorMessage
    
    frmProgram.SetProgramCodeLevel modStatusCode.StatusCodeColorLevelPrior
    
    
    Resume Next
    
    MultiRampFileSave = False
    
    'Resize window so top is visible, but bottom is hidden
    LoadFileSaveForm
    
    Me.refresh
    
    Exit Function

BadLocalFolderName:

    'This is an internal code error that should never happen
    'send MsgBox if this occurs
    ErrorMessage = "Could not create data folder for this AF Ramp. Code Execution paused.  " & vbNewLine & _
                    "Current Folder Name: " & modConfig.ADWIN_AFDataLocalDir & FolderName & vbNewLine & _
                    "Error Code = " & Trim(str(Err.number))
    
    SetCodeLevel CodeRed
    
    frmSendMail.MailNotification "Data Folder Error", _
                                 ErrorMessage, _
                                 CodeRed, _
                                 True
    
    MsgBox ErrorMessage
    
    frmProgram.SetProgramCodeLevel modStatusCode.StatusCodeColorLevelPrior
    
    
    Resume Next
    
    MultiRampFileSave = False
    
    'Resize window so top is visible, but bottom is hidden
    LoadFileSaveForm
    
    Me.refresh
    
    Exit Function
    
SineFitFileCreateError:

    ErrorMessage = "Could not create Sine Fit log data file. Code Execution paused.  " & vbNewLine & _
                    "File Path = " & modConfig.ADWIN_AFDataLocalDir & FolderName & MaxAmpFileName & vbNewLine & _
                    "Error Code = " & Trim(str(Err.number))
                    
    SetCodeLevel CodeRed
    
    frmSendMail.MailNotification "Data File Error", _
                                 ErrorMessage, _
                                 CodeRed, _
                                 True
    
    MsgBox ErrorMessage
    
    frmProgram.SetProgramCodeLevel modStatusCode.StatusCodeColorLevelPrior

    
    MultiRampFileSave = False
    
    'Resize window so top is visible, but bottom is hidden
    LoadFileSaveForm
    
    Me.refresh
    
    Exit Function
    
MaxAmpFileCreateError:

    ErrorMessage = "Could not create AF Ramp amplitudes data file. Code Execution paused.  " & vbNewLine & _
                    "File Path = " & modConfig.ADWIN_AFDataLocalDir & FolderName & MaxAmpFileName & vbNewLine & _
                    "Error Code = " & Trim(str(Err.number))
                    
    SetCodeLevel CodeRed
    
    frmSendMail.MailNotification "Data File Error", _
                                 ErrorMessage, _
                                 CodeRed, _
                                 True
    
    MsgBox ErrorMessage
    
    frmProgram.SetProgramCodeLevel modStatusCode.StatusCodeColorLevelPrior
        
    Resume Next
    
    MultiRampFileSave = False
    
    'Resize window so top is visible, but bottom is hidden
    LoadFileSaveForm
    
    Me.refresh
    
    Exit Function
    
DataFileCreateError:

    ErrorMessage = "Could not create AF ramp Data File. Code Execution paused.  " & vbNewLine & _
                    "File Path = " & modConfig.ADWIN_AFDataLocalDir & FolderName & DataFileName & vbNewLine & _
                    "Error Code = " & Trim(str(Err.number))
                    
    SetCodeLevel CodeRed
    
    frmSendMail.MailNotification "Data File Error", _
                                 ErrorMessage, _
                                 CodeRed, _
                                 True
    
    MsgBox ErrorMessage
    
    frmProgram.SetProgramCodeLevel modStatusCode.StatusCodeColorLevelPrior
    
    
    Resume Next
    
    'Resize window so top is visible, but bottom is hidden
    LoadFileSaveForm
    
    Me.refresh
    
    MultiRampFileSave = False
    
    Exit Function

BadBackupMainFolder:

    ErrorMessage = "Could not access/create Main Backup AF Ramp Data Folder. Code Execution will continue." & vbNewLine & _
                    "Backup Folder Path = " & modConfig.ADWIN_AFDataBackupDir & vbNewLine & _
                    "Error Code = " & Trim(str(Err.number)) & vbNewLine & vbNewLine & _
                    "Code Execution will continue."

    frmSendMail.MailNotification "Backup File Error", _
                                 ErrorMessage, _
                                 CodeYellow

    MultiRampFileSave = True
    
    'Resize window so top is visible, but bottom is hidden
    LoadFileSaveForm
    
    Me.refresh
    
    Exit Function

CreateBackupFolderError:

    'This error should only happen if there are permissions issues with creating folders
    'on the backup drive
    ErrorMessage = "Could not create backup folder for this AF Ramp's set of data files. Code Execution will continue.  " & vbNewLine & _
                    "Backup Folder Name: " & LocalFolder & FolderName & vbNewLine & _
                    "Error Code = " & Trim(str(Err.number)) & vbNewLine & vbNewLine & _
                    "Code Execution will continue."
    
    frmSendMail.MailNotification "Backup File Error", _
                                 ErrorMessage, _
                                 CodeYellow
    
    MultiRampFileSave = True
    
    'Resize window so top is visible, but bottom is hidden
    LoadFileSaveForm
    
    Me.refresh
    
    Exit Function

BadBackupFileCopy:

    'This error should also only happen if there are permissions issues on the backup drive
    ErrorMessage = "Could not backup file: " & DataFolder.Files(i).name & vbNewLine & _
                    "To backup folder: " & modConfig.ADWIN_AFDataBackupDir & FolderName & vbNewLine & _
                    "Error Code = " & Trim(str(Err.number)) & vbNewLine & vbNewLine & _
                    "Code Execution will continue."

    frmSendMail.MailNotification "Backup File Error", _
                                 ErrorMessage, _
                                 CodeYellow
            
    MultiRampFileSave = True

    'Resize window so top is visible, but bottom is hidden
    LoadFileSaveForm
    
    Me.refresh

End Function

Public Function MultiRampFileSave_MCC _
    (ByRef DataArray() As Double, _
     ByVal TimeStep As Double, _
     ByVal PtsPerFile As Long, _
     ByVal FolderName As String, _
     ByVal CurTime, _
     ByRef SineFitArray() As Double, _
     Optional WriteToBackup As Boolean = False, _
     Optional SaveMaxAmp As Boolean = False, _
     Optional PtsWindowForFindingMax As Long = 0) As Boolean
                                
    Dim i, ii, j, k, N, SinePoints As Long
    Dim NumBackupFiles As Long
    Dim NumChannels As Long
    Dim temp As Double
    Dim MaxAmp() As Double
    Dim ChannelsString As String
    
    Dim fso As FileSystemObject
    Dim DataStream As TextStream
    Dim MaxAmpStream As TextStream
    Dim SineFitStream As TextStream
    Dim MaxAmpFileName As String
    Dim SineFitFileName As String
    Dim DataFileName As String

    Dim DataFolder As Folder
    Dim ErrorMessage As String
        
    'Show only the progress label and progress bar
    Me.Height = ProgressHeight
    Me.Width = ProgressWidth
    Me.frameADWINAFBootDir.Visible = False
    Me.lblFileSaveProgress.Top = ProglblTop
    Me.progressFileSave.Top = ProgBarTop
    
    'Make File Save Progress bar visible
    Me.progressFileSave.min = 0
    Me.progressFileSave.Max = 32767
    Me.progressFileSave.value = 0
    Me.progressFileSave.Visible = True
    Me.lblFileSaveProgress.Caption = "File Save % Complete:   0%"
    Me.lblFileSaveProgress.Visible = True
    
    'Show file save window
    Me.Show
    
    'Find number of channels from the BaselineAvg array
    NumChannels = UBound(BaselineAvg)
        
    'Redimension MaxAmp() array so that it's the same size as the BaselineAvg array
    ReDim MaxAmp(NumChannels)
    
    Set fso = New FileSystemObject
    If Not fso.FolderExists(modConfig.TWOG_AFDataLocalDir) Then
    
        On Error GoTo BadLocalMainFolder:
        
            fso.CreateFolder modConfig.TWOG_AFDataLocalDir
            
        On Error GoTo 0
    
    End If
    
    If Not fso.FolderExists(modConfig.TWOG_AFDataLocalDir & FolderName) Then
    
        On Error GoTo BadLocalFolderName:
        
            fso.CreateFolder (modConfig.TWOG_AFDataLocalDir & FolderName)
            
        On Error GoTo 0
    
    End If
    
    On Error Resume Next
    
        SinePoints = UBound(SineFitArray)
    
        If Err.number <> 0 Then
        
            ReDim SineFitArray(1)
            
            SineFitArray(0) = "EMPTY"
            
        End If
        
    On Error GoTo 0
    
    N = UBound(DataArray)
    
    If ActiveCoilSystem = AxialCoilSystem Then
    
        CoilString = "Axial"
        
    ElseIf ActiveCoilSystem = TransverseCoilSystem Then
    
        CoilString = "Transverse"
        
    End If
    
    If SineFitArray(0) <> "EMPTY" Then
    
        SineFitFileName = "SineFits_" & Format(CurTime, "MM-DD-YY_HH-MM-SS") & ".csv"
        
        On Error GoTo SineFitFileCreateError:
    
            Set SineFitStream = fso.CreateTextFile(modConfig.TWOG_AFDataLocalDir & FolderName & SineFitFileName, True)
            
        On Error GoTo 0
        
        SineFitStream.WriteLine AFSystem & " AF Ramp Sine Fits on " & CoilString & " coil"
        SineFitStream.WriteLine Format(CurTime, "long date") & ", " & Format(CurTime, "long time")
        SineFitStream.WriteBlankLines (1)
        SineFitStream.WriteLine "From Pt. #," & _
                                "Time (ms)," & _
                                "Offset," & _
                                "Amplitude," & _
                                "Signal Freq," & _
                                "Fit Freq," & _
                                "Phase," & _
                                "RMS," & _
                                "IORate," & _
                                "Ramp Voltage," & _
                                "Ramp Counts"
        
        For i = 0 To SinePoints - 1
        
            SineFitStream.WriteLine Trim(str(SineFitArray(0))) & "," & _
                                    Trim(str(SineFitArray(1))) & "," & _
                                    Trim(str(SineFitArray(2))) & "," & _
                                    Trim(str(SineFitArray(3))) & "," & _
                                    Trim(str(SineFitArray(4))) & "," & _
                                    Trim(str(SineFitArray(5))) & "," & _
                                    Trim(str(SineFitArray(6))) & "," & _
                                    Trim(str(SineFitArray(7))) & "," & _
                                    Trim(str(SineFitArray(8))) & "," & _
                                    Trim(str(SineFitArray(9)))

            'Need to update progress
            Me.progressFileSave.value = CInt(32767 * i / (N + SinePoints - 1))
            PercComplete = Trim(str(CInt(100 * i / (N + SinePoints - 1))))
            Do While Len(PercComplete) < 4
            
                PercComplete = " " & PercComplete
                
            Loop
            
            Me.lblFileSaveProgress.Caption = "File Save % Complete:" & PercComplete & "%"
    
            Me.refresh
            
            DoEvents
        
        Next i
        
        SineFitStream.Close
        
    End If
    
    
    If SaveMaxAmp = True Then
    
        MaxAmpFileName = "AFRamplitude_" & Format(CurTime, "MM-DD-YY_HH-MM-SS") & ".csv"
    
        On Error GoTo MaxAmpFileCreateError:
    
            Set MaxAmpStream = fso.CreateTextFile(modConfig.TWOG_AFDataLocalDir & FolderName & MaxAmpFileName, True)
            
        On Error GoTo 0
        
        MaxAmpStream.WriteLine AFSystem & " AF Ramp Amplitudes on " & CoilString & " coil"
        MaxAmpStream.WriteLine Format(CurTime, "long date") & "," & Format(CurTime, "long time")
        MaxAmpStream.WriteBlankLines (1)
        MaxAmpStream.WriteLine "Sliding Window =" & "," & Trim(str(PtsWindowForFindingMax))
        
        ChannelsString = ""
        
        For j = 0 To NumChannels - 1
            
            If BaselineAvg Is Nothing Then
            
                ChannelsString = ChannelsString & "Ch " & Trim(str(j)) & ","
            
            Else
            
                ChannelsString = ChannelsString & "Ch " & Trim(str(j)) & " Raw,"
                ChannelsString = ChannelsString & "Ch " & Trim(str(j)) & " Cor,"
                
            End If
            
        Next j
        
        MaxAmpStream.WriteLine "Point #,Time (ms)," & ChannelsString
    
    Else
    
        Set MaxAmpStream = Nothing

    End If
    
    'Create first data file
    DataFileName = "AFramp_pts0-" & Trim(str(PtsPerFile)) & "_" & Format(CurTime, "MM-DD-YY_HH-MM-SS") & ".csv"
    
    On Error GoTo DataFileCreateError:
    
        Set DataStream = fso.CreateTextFile(modConfig.TWOG_AFDataLocalDir & FolderName & DataFileName, True)
        
    On Error GoTo 0
    
    
    
    DataStream.WriteLine AFSystem & " AF Ramp on " & CoilString & " coil"
    DataStream.WriteLine Format(CurTime, "long date") & ", " & Format(CurTime, "long time")
    DataStream.WriteBlankLines (1)
    DataStream.WriteLine "From = ,0"
    DataStream.WriteLine "To = ," & Trim(str(PtsPerFile))
    
    ChannelsString = ""
    
    For j = 0 To NumChannels - 1
        
        If BaselineAvg Is Nothing Then
        
            ChannelsString = ChannelsString & "Ch " & Trim(str(j)) & ","
        
        Else
        
            ChannelsString = ChannelsString & "Ch " & Trim(str(j)) & " Raw,"
            ChannelsString = ChannelsString & "Ch " & Trim(str(j)) & " Cor,"
            
        End If
        
    Next j
    
    DataStream.WriteLine "Point #,Time (ms)," & ChannelsString
    
    j = PtsWindowForFindingMax
                    
    N = UBound(DataArray)
    
    'Initialize k = # pts per file + 1
    k = PtsPerFile + 1
    
    'Initialize max amp to the absolute value first point of the monitor array
    For ii = 0 To NumChannels - 1
    
        MaxAmp(ii) = 0
        
    Next ii
    
    For i = 0 To N - 1 Step NumChannels
    
        DoEvents
    
        If i Mod 5000 = 0 Then
        
            'Need to update progress
            Me.progressFileSave.value = CInt(32767 * (i + SinePoints) / (N - 1 + SinePoints))
            PercComplete = Trim(str(CInt(100 * (i + SinePoints) / (N - 1 + SinePoints))))
            Do While Len(PercComplete) < 4
            
                PercComplete = " " & PercComplete
                
            Loop
            
            Me.lblFileSaveProgress.Caption = "File Save % Complete:" & PercComplete & "%"
    
            Me.refresh
    
        End If
    
        ChannelsString = ""
    
        For ii = 0 To NumChannels - 1
            
            temp = DataArray(i + ii)
            
            If BaselineAvg Is Nothing Then
            
                ChannelsString = ChannelsString & Trim(str(temp)) & ","
            
            Else
            
                ChannelsString = ChannelsString & Trim(str(temp)) & "," & _
                                    Trim(str(temp - BaselineAvg(ii))) & ","
                
            End If
          
            If MaxAmp(ii) < Abs(temp) And SaveMaxAmp = True Then MaxAmp(ii) = Abs(temp)
          
        Next ii
        
        DataStream.WriteLine Trim(str(i)) & "," & _
                                 Trim(str(i * TimeStep)) & "," & _
                                 ChannelsString
            
        
        j = j - 1
        k = k - 1
        
        If j = 0 And SaveMaxAmp = True Then
        
            ChannelsString = ""
    
            For ii = 0 To NumChannels - 1
            
                If BaselineAvg Is Nothing Then
                
                    ChannelsString = ChannelsString & Trim(str(MaxAmp(ii))) & ","
              
                Else
              
                    ChannelsString = ChannelsString & Trim(str(MaxAmp(ii))) & "," & _
                                      Trim(str(MaxAmp(ii) - BaselineAvg(ii))) & ","
                  
                End If
            
                MaxAmp(ii) = -100
          
            Next ii
            
            
            MaxAmpStream.WriteLine Trim(str(i)) & "," & _
                                       Trim(str(i * TimeStep)) & "," & _
                                       ChannelsString
            
            j = PtsWindowForFindingMax
            
        End If
        
        If k = 0 And Not i = N - 1 Then
        
            'Need to close the current text stream, create a new text file,
            'and open it up for writing
            DataStream.Close
            
            'Create next data file
            'Adjust name to refltect the current point we're on and adjust the
            'point to for the final file
            If i + PtsPerFile > N - 1 Then
                
                DataFileName = "AFramp_pts" & Trim(str(i + 1)) & "-" & Trim(str(N - 1)) & _
                                "_" & Format(CurTime, "MM-DD-YY_HH-MM-SS") & ".csv"
                                
            Else
            
                DataFileName = "AFramp_pts" & Trim(str(i + 1)) & "-" & Trim(str(i + PtsPerFile)) & _
                                "_" & Format(CurTime, "MM-DD-YY_HH-MM-SS") & ".csv"
                                
            End If
            
            On Error GoTo DataFileCreateError:
            
                Set DataStream = fso.CreateTextFile(modConfig.TWOG_AFDataLocalDir & FolderName & DataFileName, True)
                
            On Error GoTo 0
            
            DataStream.WriteLine AFSystem & " AF Ramp on " & CoilString & " coil"
            DataStream.WriteLine Format(CurTime, "long date") & ", " & Format(CurTime, "long time")
            DataStream.WriteBlankLines (1)
            DataStream.WriteLine "From = ," & Trim(str(i + 1))
            DataStream.WriteLine "To = ," & Trim(str(i + PtsPerFile))
            
            ChannelsString = ""
    
            For ii = 0 To NumChannels - 1
                
                If BaselineAvg Is Nothing Then
                
                    ChannelsString = ChannelsString & "Ch " & Trim(str(ii)) & ","
                
                Else
                
                    ChannelsString = ChannelsString & "Ch " & Trim(str(ii)) & " Raw,"
                    ChannelsString = ChannelsString & "Ch " & Trim(str(ii)) & " Cor,"
                    
                End If
                
            Next ii
            
            DataStream.WriteLine "Point #,Time (ms)," & ChannelsString
            
            k = PtsPerFile
            
        End If
        
    Next i
    
    'Close the final file string
    DataStream.Close
                
    If WriteToBackup And _
       Me.chkBackupRampData.value = Checked _
    Then
    
        'Now need to copy the ramp files to the remote backup path
                
        'Change Value on Progress bar to 0
        Me.progressFileSave.value = 0
        
        'Determine the number of files to backup
        NumBackupFiles = CInt(N / PtsPerFile) + 1
        
        If SaveMaxAmp = True Then NumBackupFiles = NumBackupFiles + 1
        
        'Change Caption in lblFileSaveprog
        Me.lblFileSaveProgress.Caption = "Creating Backup Folder..."
        Me.refresh
        
        'See if there is a main backup ramp data folder yet,
        'if not, create it
        If Not fso.FolderExists(modConfig.TWOG_AFDataBackupDir) Then
        
            On Error GoTo BadBackupMainFolder:
            
                fso.CreateFolder modConfig.TWOG_AFDataBackupDir
                
            On Error GoTo 0
            
        End If
        
        'Create the folder for this AF ramp's worth of data files
        If Not fso.FolderExists(modConfig.TWOG_AFDataBackupDir & FolderName) Then
        
            On Error GoTo CreateBackupFolderError:
            
                fso.CreateFolder modConfig.TWOG_AFDataBackupDir & FolderName
        
            On Error GoTo 0
            
        End If
        
        'Change Caption on lblFileSaveProg
        Me.lblFileSaveProgress.Caption = "Backing up Files: " & _
                                            Trim(str(NumBackupFiles)) & " Files remaining..."
        Me.refresh
                
        'Navigate the local AF Ramp data folder, and copy the files in the
        'folder to the backup folder one by one
        Set DataFolder = fso.GetFolder(modConfig.TWOG_AFDataLocalDir & FolderName)
                
               
        N = DataFolder.Files.Count
                
        For i = 1 To N
        
            With DataFolder.Files(i)
            
                On Error GoTo BadBackupFileCopy:
            
                    .COPY modConfig.TWOG_AFDataBackupDir & FolderName & .name, True
                
                On Error GoTo 0
                
                Me.lblFileSaveProgress.Caption = "Backing up Files: " & _
                                                    Trim(str(NumBackupFiles - i)) & _
                                                    " Files remaining..."
        
                Me.progressFileSave.value = CInt(32767 * i / N)
                Me.refresh
                
            End With
            
        Next i
        
    End If
                
    
    'Uncheck the Monitor & Save AF Ramp box
    frmADWIN_AF.chkVerbose.value = Unchecked

    'Clear the third panel of the program status bar
    frmProgram.StatusBar vbNullString, 3

    MultiRampFileSave = True

    'Hide the filesave form
    Me.Hide
    
    'Reload the file save form so that the bottom is hidden and the top is shown
    Load frmFileSave

    Exit Function
    
BadLocalMainFolder:

    ErrorMessage = "Could not find/access AF Ramp main data folder. Code Execution paused." & vbNewLine & _
                    "Current path = " & modConfig.TWOG_AFDataLocalDir & vbNewLine & _
                    "Error Code = " & Trim(str(Err.number))
    
    SetCodeLevel CodeRed
    
    frmSendMail.MailNotification "Folder Path Error", _
                                 ErrorMessage, _
                                 CodeRed, _
                                 True
    
    MsgBox ErrorMessage
    
    frmProgram.SetProgramCodeLevel modStatusCode.StatusCodeColorLevelPrior
    
    Resume Next
    
    MultiRampFileSave = False
       
    'Resize window so top is visible, but bottom is hidden
    LoadFileSaveForm
    
    Me.refresh
    
    Exit Function

BadLocalFolderName:

    'This is an internal code error that should never happen
    'send MsgBox if this occurs
    ErrorMessage = "Could not create data folder for this AF Ramp. Code Execution paused.  " & vbNewLine & _
                    "Current Folder Name: " & modConfig.TWOG_AFDataLocalDir & FolderName & vbNewLine & _
                    "Error Code = " & Trim(str(Err.number))
    
    SetCodeLevel CodeRed
    
    frmSendMail.MailNotification "Data Folder Error", _
                                 ErrorMessage, _
                                 CodeRed, _
                                 True
    
    MsgBox ErrorMessage
    
    frmProgram.SetProgramCodeLevel modStatusCode.StatusCodeColorLevelPrior
    
    Resume Next
    
    MultiRampFileSave = False
       
    'Resize window so top is visible, but bottom is hidden
    LoadFileSaveForm
    
    Me.refresh
    
    Exit Function
    
SineFitFileCreateError:

    ErrorMessage = "Could not create Sine Fit log data file. Code Execution paused.  " & vbNewLine & _
                    "File Path = " & modConfig.TWOG_AFDataLocalDir & FolderName & MaxAmpFileName & vbNewLine & _
                    "Error Code = " & Trim(str(Err.number))
                    
    SetCodeLevel CodeRed
    
    frmSendMail.MailNotification "Data File Error", _
                                 ErrorMessage, _
                                 CodeRed, _
                                 True
    
    MsgBox ErrorMessage
    
    frmProgram.SetProgramCodeLevel modStatusCode.StatusCodeColorLevelPrior
    
    Resume Next
    
    MultiRampFileSave = False
        
    'Resize window so top is visible, but bottom is hidden
    LoadFileSaveForm
    
    Me.refresh
    
    Exit Function
    
MaxAmpFileCreateError:

    ErrorMessage = "Could not create AF Ramp amplitudes data file. Code Execution paused.  " & vbNewLine & _
                    "File Path = " & modConfig.TWOG_AFDataLocalDir & FolderName & MaxAmpFileName & vbNewLine & _
                    "Error Code = " & Trim(str(Err.number))
                    
    SetCodeLevel CodeRed
    
    frmSendMail.MailNotification "Data File Error", _
                                 ErrorMessage, _
                                 CodeRed, _
                                 True
    
    MsgBox ErrorMessage
    
    frmProgram.SetProgramCodeLevel modStatusCode.StatusCodeColorLevelPrior
    
    
    Resume Next
    
    MultiRampFileSave = False
    
    'Resize window so top is visible, but bottom is hidden
    LoadFileSaveForm
    
    Me.refresh
    
    Exit Function
    
DataFileCreateError:

    ErrorMessage = "Could not create AF ramp Data File. Code Execution paused.  " & vbNewLine & _
                    "File Path = " & modConfig.TWOG_AFDataLocalDir & FolderName & DataFileName & vbNewLine & _
                    "Error Code = " & Trim(str(Err.number))
                    
    SetCodeLevel CodeRed
    
    frmSendMail.MailNotification "Data File Error", _
                                 ErrorMessage, _
                                 CodeRed, _
                                 True
    
    MsgBox ErrorMessage
    
    frmProgram.SetProgramCodeLevel modStatusCode.StatusCodeColorLevelPrior
    
    
    Resume Next
    
    'Resize window so top is visible, but bottom is hidden
    LoadFileSaveForm
    
    Me.refresh
        
    MultiRampFileSave = False
    
    Exit Function

BadBackupMainFolder:

    ErrorMessage = "Could not access/create Main Backup AF Ramp Data Folder. Code Execution will continue." & vbNewLine & _
                    "Backup Folder Path = " & modConfig.TWOG_AFDataBackupDir & vbNewLine & _
                    "Error Code = " & Trim(str(Err.number)) & vbNewLine & vbNewLine & _
                    "Code Execution will continue."

    frmSendMail.MailNotification "Backup File Error", _
                                  ErrorMessage, _
                                  CodeYellow

    MultiRampFileSave = True
    
    'Resize window so top is visible, but bottom is hidden
    LoadFileSaveForm
    
    Me.refresh
    
    Exit Function

CreateBackupFolderError:

    'This error should only happen if there are permissions issues with creating folders
    'on the backup drive
    ErrorMessage = "Could not create backup folder for this AF Ramp's set of data files. Code Execution will continue.  " & vbNewLine & _
                    "Backup Folder Name: " & modConfig.TWOG_AFDataLocalDir & FolderName & vbNewLine & _
                    "Error Code = " & Trim(str(Err.number)) & vbNewLine & vbNewLine & _
                    "Code Execution will continue."
    
    frmSendMail.MailNotification "Backup File Error", _
                                 ErrorMessage, _
                                 CodeYellow
                                     
    MultiRampFileSave = True
    
    'Resize window so top is visible, but bottom is hidden
    LoadFileSaveForm
    
    Me.refresh
    
    Exit Function

BadBackupFileCopy:

    'This error should also only happen if there are permissions issues on the backup drive
    ErrorMessage = "Could not backup file: " & DataFolder.Files(i).name & vbNewLine & _
                    "To backup folder: " & modConfig.TWOG_AFDataBackupDir & FolderName & vbNewLine & _
                    "Error Code = " & Trim(str(Err.number)) & vbNewLine & vbNewLine & _
                    "Code Execution will continue."

    frmSendMail.MailNotification "Backup File Error", _
                                 ErrorMessage, _
                                 CodeYellow
            
    MultiRampFileSave = True
    
    'Resize window so top is visible, but bottom is hidden
    LoadFileSaveForm
    
    Me.refresh
    
End Function

Public Function SaveAFCalibrationTable _
    (ByRef CalTable As MSHFlexGrid, _
     ByVal CurTime As Variant, _
     Optional ByVal Units As String = "G") As Boolean

    Dim i, j As Long
    Dim NumReplicates As Long
    Dim NumRows As Long
    Dim SumVolts As Double
    Dim SumVarVolts As Double
    Dim AvgVolts As Double
    Dim StdDevVolts As Double
        
    Dim fso As FileSystemObject
    Dim FileStream As TextStream
    Dim filename As String
    Dim CalibrationFolder As String
    
    Dim TempString As String
    Dim TempString2 As String
    Dim ErrorMessage As String
    
    'Allocate file system object
    Set fso = New FileSystemObject
        
    'Store the default calibration folder name
    CalibrationFolder = "Calibration Values/"
    
    'Show only the progress label and progress bar
    Me.Height = ProgressHeight
    Me.frameADWINAFBootDir.Visible = False
    Me.lblFileSaveProgress.Top = ProglblTop
    Me.progressFileSave.Top = ProgBarTop
    Me.lblFileSaveProgress.Caption = "File Save Progress:"
    
    'Setup the progress bar
    Me.progressFileSave.Max = 32767
    Me.progressFileSave.min = 0
    Me.progressFileSave.value = 0
    
    'Refresh the form
    Me.Show
    Me.refresh
    
    'Check to see if the local data folder already exists
    If fso.FolderExists(Me.txtRampDataLocalFolder) = False Then
    
        'Create the folder
        On Error Resume Next
            
            fso.CreateFolder (Me.txtRampDataLocalFolder)
            
            If Err.number <> 0 Then
            
                SetCodeLevel CodeRed
            
                ErrorMessage = "Unable to create local data folder: " & vbNewLine & _
                               vbTab & Me.txtRampDataLocalFolder & vbNewLine & _
                               vbNewLine & "Please check / change folder path in the " & _
                               "AF File Save Settings window." & vbNewLine & vbNewLine & _
                               "Code execution has been paused."
            
                'Raise Err - unable to create folder
                frmSendMail.MailNotification _
                          "Data Folder Error", _
                          ErrorMessage, _
                          CodeRed, _
                          True
                                
                MsgBox ErrorMessage
                
                frmProgram.SetProgramCodeLevel StatusCodeColorLevelPrior
                
                'Return a false to indicate that loading the table has failed
                SaveAFCalibrationTable = False
                
                Exit Function
                
            End If
            
        On Error GoTo 0
                          
    End If
    
    'Check to see if the calibration folder already exists
    If fso.FolderExists(Me.txtRampDataLocalFolder & CalibrationFolder) = False Then
    
        fso.CreateFolder (Me.txtRampDataLocalFolder & CalibrationFolder)
        
    End If
    
    'Set the CoilString to the active coil system
    If ActiveCoilSystem = AxialCoilSystem Then
    
        CoilString = "Axial"
        
    ElseIf ActiveCoilSystem = TransverseCoilSystem Then
    
        CoilString = "Transverse"
        
    End If
    
    'Create the .csv file name from the active AF coil string, and the time
    filename = CoilString & "_Cal_" & Format(CurTime, "MM-DD-YY_HH-MM-SS") & ".csv"
    
    'Create the .csv calibration data file
    Set FileStream = fso.CreateTextFile(Me.txtRampDataLocalFolder & _
                                        CalibrationFolder & _
                                        filename)
    
    'Write in the calibration meta-data (time, coil + column headers, etc.
    FileStream.WriteLine "AF " & CoilString & " coil Peak Field Calibration"
    
    'Determine the number of replicate measurements being done
    NumReplicates = (frmCalibrateCoils.gridCalibration.Cols - 4) \ 2
    
    'Save the file write progress
    Me.progressFileSave = CInt(32767 * (1 / (NumReplicates + 4)))
    
    FileStream.WriteLine Format(CurTime, "Long Date") & "," & Format(CurTime, "Long Time")
    
    'Save the file write progress
    Me.progressFileSave = CInt(32767 * (2 / (NumReplicates + 4)))
    
    FileStream.WriteBlankLines (1)
    
    'Save the file write progress
    Me.progressFileSave = CInt(32767 * (3 / (NumReplicates + 4)))
    
    'Store the start of the column headers file text line
    TempString = "Target Mon. Volt. (V),Field (" & Units & ")," & _
                 "Std Dev (" & Units & "),Actual Mon. Volt. (V),Std Dev (V)"
    
    'Determine the number of replicates
    NumReplicates = (CalTable.Cols - 4) \ 2
    
    'Generate the column headers for all the replicates
    For i = 1 To NumReplicates
    
        TempString = TempString & _
                     ",Field #" & Trim(str(i)) & " (" & Units & ") " & _
                     ",Volts #" & Trim(str(i)) & " (V)"
                     
    Next i
    
    'Write the headers text line to file
    FileStream.WriteLine TempString
    
    'Save the file write progress
    Me.progressFileSave = CInt(32767 * (4 / (NumReplicates + 4)))
    
    'Set the number of data rows to write to file
    NumRows = CalTable.Rows - 1
    
    'Loop through the data rows in the calibration table
    'and write the data in each row + average voltage & it's std. dev.
    'to file
    For i = 1 To NumRows
    
        'Need to store each column value to TempString = file line string
        'that will be written to the file once all the table data column values
        'for this row have been stored to it
        
        With CalTable
        
            'Store string Target monitor volt for this series of replicate ramps
            .row = i
            .Col = 1
            TempString = Trim(str(val(.text))) & ","
            
            'Add string of Resulting average Field value
            .row = i
            .Col = 2
            TempString = TempString & Trim(str(val(.text))) & ","
            
            'Add string of Standard deviation of the average field value
            .row = i
            .Col = 3
            TempString = TempString & Trim(str(val(.text)))
            
            'Empty this string - it's going to store the text values for the replicates
            TempString2 = ""
            
            'These values are used to figure out the average monitor voltage
            'and the standard deviation of that average
            SumVolts = 0
            SumVarVolts = 0
            
            'Loop through the replicates
            For j = 4 To 4 + NumReplicates * 2 - 1 Step 2
            
                'Add string of Field replicate measurement
                .row = i
                .Col = j
                TempString2 = TempString2 & "," & Trim(str(val(.text)))
            
                'Add string of Monitor voltage replicate value
                .row = i
                .Col = j + 1
                TempString2 = TempString2 & "," & Trim(str(val(.text)))
                
                'Sum all the replicate voltages for this row together
                SumVolts = SumVolts + val(.text)
                
            Next j
            
            'Calculate the average actual monitor voltage for this row
            AvgVolts = SumVolts / NumReplicates
            
            'Now need to do that again for the standard deviation
            For j = 4 To 4 + NumReplicates * 2 - 1 Step 2
            
                .row = i
                .Col = j + 1
                
                'Sum the variances from the average monitor voltage for this row
                SumVarVolts = SumVarVolts + (val(.text) - AvgVolts) ^ 2
                
            Next j
            
            'Change the sum of variances into a standard deviation value
            StdDevVolts = Sqr(SumVarVolts / NumReplicates)
        
            'Add the average actual monitor voltage and it's standard deviation
            'to the line string
            TempString = TempString & _
                         "," & Trim(str(AvgVolts)) & _
                         "," & Trim(str(StdDevVolts))
            
            'Concatenate the replicate column data string with the
            'string containing the target voltage, averages, and std dev's
            TempString = TempString & TempString2
            
            'Write the full file line string for this row to the .csv file
            FileStream.WriteLine TempString
            
            'Save the file write progress
            Me.progressFileSave = CInt(32767 * ((i + 4) / (NumReplicates + 4 + NumRows)))
            
        End With
        
    Next i
    
    'Save successful, return "True"
    SaveAFCalibrationTable = True
    
    'Hide the filesave form
    Me.Hide
    
    'Reload the file save form so that the bottom is hidden and the top is shown
    Load frmFileSave
        
End Function

Public Function SaveIRMCalibrationTable _
    (ByRef CalTable As MSHFlexGrid, _
     ByVal CurTime As Variant, _
     Optional ByVal Units As String = "G") As Boolean

    Dim i, j As Long
    Dim NumReplicates As Long
    Dim NumRows As Long
    Dim SumVolts As Double
    Dim SumVarVolts As Double
    Dim AvgVolts As Double
    Dim StdDevVolts As Double
        
    Dim fso As FileSystemObject
    Dim FileStream As TextStream
    Dim filename As String
    Dim CalibrationFolder As String
    
    Dim TempString As String
    Dim TempString2 As String
    Dim ErrorMessage As String
    
    'Allocate file system object
    Set fso = New FileSystemObject
        
    'Store the default calibration folder name
    CalibrationFolder = "Calibration Values/"
    
    'Show only the progress label and progress bar
    Me.Height = ProgressHeight
    Me.frameADWINAFBootDir.Visible = False
    Me.lblFileSaveProgress.Top = ProglblTop
    Me.progressFileSave.Top = ProgBarTop
    Me.lblFileSaveProgress.Caption = "File Save Progress:"
    
    'Setup the progress bar
    Me.progressFileSave.Max = 32767
    Me.progressFileSave.min = 0
    Me.progressFileSave.value = 0
    
    'Refresh the form
    Me.Show
    Me.refresh
    
    'Check to see if the local data folder already exists
    If fso.FolderExists(Me.txtRampDataLocalFolder) = False Then
    
        'Create the folder
        On Error Resume Next
            
            fso.CreateFolder (Me.txtRampDataLocalFolder)
            
            If Err.number <> 0 Then
            
                SetCodeLevel CodeRed
            
                ErrorMessage = "Unable to create local data folder: " & vbNewLine & _
                               vbTab & Me.txtRampDataLocalFolder & vbNewLine & _
                               vbNewLine & "Please check / change folder path in the " & _
                               "AF File Save Settings window." & vbNewLine & vbNewLine & _
                               "Code execution has been paused."
            
                'Raise Err - unable to create folder
                frmSendMail.MailNotification _
                          "Data Folder Error", _
                          ErrorMessage, _
                          CodeRed, _
                          True
                                
                MsgBox ErrorMessage
                
                frmProgram.SetProgramCodeLevel StatusCodeColorLevelPrior
                
                'Return a false to indicate that loading the table has failed
                SaveIRMCalibrationTable = False
                
                Exit Function
                
            End If
            
        On Error GoTo 0
                          
    End If
    
    'Check to see if the calibration folder already exists
    If fso.FolderExists(Me.txtRampDataLocalFolder & CalibrationFolder) = False Then
    
        fso.CreateFolder (Me.txtRampDataLocalFolder & CalibrationFolder)
        
    End If
    
    'Set the CoilString to the active coil system
    If ActiveCoilSystem = AxialCoilSystem Then
    
        CoilString = "IRM Axial"
        
    ElseIf ActiveCoilSystem = TransverseCoilSystem Then
    
        CoilString = "IRM Transverse"
        
    End If
    
    'Create the .csv file name from the active AF coil string, and the time
    filename = CoilString & "_Cal_" & Format(CurTime, "MM-DD-YY_HH-MM-SS") & ".csv"
    
    'Create the .csv calibration data file
    Set FileStream = fso.CreateTextFile(Me.txtRampDataLocalFolder & _
                                        CalibrationFolder & _
                                        filename)
    
    'Write in the calibration meta-data (time, coil + column headers, etc.
    FileStream.WriteLine "IRM " & CoilString & " coil Peak Field Calibration"
    
    'Save the file write progress
    Me.progressFileSave = CInt(32767 * (1 / (NumReplicates + 4)))
    
    FileStream.WriteLine Format(CurTime, "Long Date") & "," & Format(CurTime, "Long Time")
    
    'Save the file write progress
    Me.progressFileSave = CInt(32767 * (2 / (NumReplicates + 4)))
    
    FileStream.WriteBlankLines (1)
    
    'Save the file write progress
    Me.progressFileSave = CInt(32767 * (3 / (NumReplicates + 4)))
    
    'Store the start of the column headers file text line
    TempString = "Target Pulse Volt. (V),Field (" & Units & ")," & _
                 "Std Dev (" & Units & "),Actual Pulse Volt. (V),Std Dev (V)"
    
    'Determine the number of replicates
    NumReplicates = (CalTable.Cols - 4) \ 2
    
    'Generate the column headers for all the replicates
    For i = 1 To NumReplicates
    
        TempString = TempString & _
                     ",Field #" & Trim(str(i)) & " (" & Units & ") " & _
                     ",Volts #" & Trim(str(i)) & " (V)"
                     
    Next i
    
    'Write the headers text line to file
    FileStream.WriteLine TempString
    
    'Save the file write progress
    Me.progressFileSave = CInt(32767 * (4 / (NumReplicates + 4)))
    
    'Set the number of data rows to write to file
    NumRows = CalTable.Rows - 1
    
    'Loop through the data rows in the calibration table
    'and write the data in each row + average voltage & it's std. dev.
    'to file
    For i = 1 To NumRows
    
        'Need to store each column value to TempString = file line string
        'that will be written to the file once all the table data column values
        'for this row have been stored to it
        
        With CalTable
        
            'Store string Target monitor volt for this series of replicate ramps
            .row = i
            .Col = 1
            TempString = Trim(str(val(.text))) & ","
            
            'Add string of Resulting average Field value
            .row = i
            .Col = 2
            TempString = TempString & Trim(str(val(.text))) & ","
            
            'Add string of Standard deviation of the average field value
            .row = i
            .Col = 3
            TempString = TempString & Trim(str(val(.text)))
            
            'Empty this string - it's going to store the text values for the replicates
            TempString2 = ""
            
            'These values are used to figure out the average monitor voltage
            'and the standard deviation of that average
            SumVolts = 0
            SumVarVolts = 0
            
            'Loop through the replicates
            For j = 4 To 4 + NumReplicates * 2 - 1 Step 2
            
                'Add string of Field replicate measurement
                .row = i
                .Col = j
                TempString2 = TempString2 & "," & Trim(str(val(.text)))
            
                'Add string of Monitor voltage replicate value
                .row = i
                .Col = j + 1
                TempString2 = TempString2 & "," & Trim(str(val(.text)))
                
                'Sum all the replicate voltages for this row together
                SumVolts = SumVolts + val(.text)
                
            Next j
            
            'Calculate the average actual monitor voltage for this row
            If NumReplicates > 0 Then
                AvgVolts = SumVolts / NumReplicates
            Else
                AvgVolts = 0
            End If
            
            'Now need to do that again for the standard deviation
            For j = 4 To 4 + NumReplicates * 2 - 1 Step 2
            
                .row = i
                .Col = j + 1
                
                'Sum the variances from the average monitor voltage for this row
                SumVarVolts = SumVarVolts + (val(.text) - AvgVolts) ^ 2
                
            Next j
            
            'Change the sum of variances into a standard deviation value
            If NumReplicates > 0 Then
                StdDevVolts = Sqr(SumVarVolts / NumReplicates)
            Else
                StdDevVolts = 0
            End If
        
            'Add the average actual monitor voltage and it's standard deviation
            'to the line string
            TempString = TempString & _
                         "," & Trim(str(AvgVolts)) & _
                         "," & Trim(str(StdDevVolts))
            
            'Concatenate the replicate column data string with the
            'string containing the target voltage, averages, and std dev's
            TempString = TempString & TempString2
            
            'Write the full file line string for this row to the .csv file
            FileStream.WriteLine TempString
            
            'Save the file write progress
            Me.progressFileSave = CInt(32767 * ((i + 4) / (NumRows + 4)))
            
        End With
        
    Next i
    
    'Hide the filesave form
    Me.Hide
    
    'Reload the file save form so that the bottom is hidden and the top is shown
    Load frmFileSave
    
    'Save successful, return "True"
    SaveIRMCalibrationTable = True
        
End Function

