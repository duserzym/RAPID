VERSION 5.00
Begin VB.Form frmADWIN_AF_CommSettings 
   Caption         =   "AF DAQ Comm Settings"
   ClientHeight    =   7350
   ClientLeft      =   13770
   ClientTop       =   1395
   ClientWidth     =   8160
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7350
   ScaleWidth      =   8160
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   372
      Left            =   6120
      TabIndex        =   2
      Top             =   6600
      Width           =   1332
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   372
      Left            =   2760
      TabIndex        =   1
      Top             =   5520
      Width           =   1572
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   372
      Left            =   360
      TabIndex        =   0
      Top             =   5280
      Width           =   1452
   End
End
Attribute VB_Name = "frmADWIN_AF_CommSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PriorListIndex As Long



Private Sub cmbAFHFMonitorBoard_GotFocus()

    PriorListIndex = cmbAFHFMonitorBoard.ListIndex

End Sub

Private Sub cmbAFLFMonitorBoard_GotFocus()

    PriorListIndex = cmbAFLFMonitorBoard.ListIndex

End Sub

Private Sub cmbAFRampBoard_GotFocus()

    PriorListIndex = cmbAFHFMonitorBoard.ListIndex

End Sub

Private Sub cmbAltAFMonitorBoard_GotFocus()

    PriorListIndex = cmbAltAFMonitorBoard.ListIndex

End Sub

Private Sub cmbUnits_Click()

    Dim Response As Long
    
    If PriorListIndex <> cmbUnits.ListIndex Then
        
        'Check with user to confirm units change
        Response = MsgBox("Incorrect units/values can break the system!" & vbCrLf & _
                          "Are you sure you want to change units?", _
                          vbYesNo, _
                          "Warning!")
                          
        If Response = vbYes And _
           modConfig.AFUnits <> cmbUnits.List(cmbUnits.ListIndex) _
        Then
    
            'Update the global variable
            modConfig.AFUnits = cmbUnits.List(cmbUnits.ListIndex)
            
            'Convert all of the AF field values in the program
            'except that for this form
            frmSettings.ConvertFieldValues Me
            
        ElseIf AFUnits <> cmbUnits.List(cmbUnits.ListIndex) Then
        
            'Need to revert cmbUnits back to the prior units value
            If AFUnits = "G" Then cmbUnits.ListIndex = 0
            If AFUnits = "mT" Then cmbUnits.ListIndex = 1
            
        End If
        
        PriorListIndex = cmbUnits.ListIndex
        
    End If
        
End Sub

Private Sub cmbUnits_GotFocus()

    PriorListIndex = cmbUnits.ListIndex

End Sub

Private Sub cmdApply_Click()

    Dim UserResp As Long
    
    'First confirm with the user that these changes should be made
    UserResp = MsgBox("Incorrect values can break the system!" & vbCrLf & _
                      "Are you sure you want to make changes?", _
                      vbYesNo, _
                      "Warning!")
                      
    If UserResp = vbYes Then

        'Need to write these settings to the global variables
        
        'Text settings first
        modConfig.AFAxialLowFieldMaxValue = val(Me.txtAxialLowFieldMax)
        modConfig.AFTransLowFieldMaxValue = val(Me.txtTransLowFieldMax)
        modConfig.AFUnits = Me.cmbUnits.List(cmbUnits.ListIndex)
        
        'Alt AF monitor checkbox setting
        modConfig.EnableAltAFMonitor = (Me.checkAltAFMonitorEnabled.Value = Checked)
        
        'Now need to get the channel settings loaded into the correct
        'wave and channel objects
        
        'Change the Error handling so that we capture errors into a labelled
        'section at the end of the subroutine
        On Error GoTo BadWave:
            
            'Analog Output
            
            'Save the AF Ramp Analog output channel settings
            Set modConfig.AFRampChan = frmSettings.SaveDAQSetting(Me.cmbAFRampBoard, _
                                                                  Me.cmbAFRampChan, _
                                                                  "AO", _
                                                                  "AF Ramp Analog Output")
                                                                  
            'Store this setting to the AFRAMPUP & AFRAMPDOWN wave objects
            Set modConfig.WaveForms("AFRAMPUP").Chan = modConfig.AFRampChan
            Set modConfig.WaveForms("AFRAMPDOWN").Chan = modConfig.AFRampChan
                                                                  
                                                                  
            'Analog Input
            
            'Save the AF Low-Field Monitor Analog Input channel settings
            Set modConfig.AFLFMonitorChan = _
                            frmSettings.SaveDAQSetting(Me.cmbAFLFMonitorBoard, _
                                                       Me.cmbAFLFMonitorChan, _
                                                       "AI", _
                                                       "AF Low-Field Monitor Analog Input")
            
            'Store this setting to the AFLFMONITOR wave object
            Set modConfig.WaveForms("AFLFMONITOR").Chan = modConfig.AFLFMonitorChan
            
            
            'Save the AF High-Field Monitor Analog Input channel settings
            Set modConfig.AFMonitorChan = _
                            frmSettings.SaveDAQSetting(Me.cmbAFHFMonitorBoard, _
                                                       Me.cmbAFMonitorChan, _
                                                       "AI", _
                                                       "AF High-Field Monitor Analog Input")
                                                       
            'Store this setting to the AFHFMONITOR wave object
            Set modConfig.WaveForms("AFMONITOR").Chan = modConfig.AFMonitorChan
                                                       
                                                       
            'Save the Alternate AF Monitor Analog Input channel settings
            Set modConfig.AltAFMonitorChan = _
                            frmSettings.SaveDAQSetting(Me.cmbAltAFMonitorBoard, _
                                                       Me.cmbAltAFMonitorChan, _
                                                       "AI", _
                                                       "Alternate AF Monitor Analog Input")
                                                       
            'Store this setting to the ALTAFMONITOR wave object
            Set modConfig.WaveForms("ALTAFMONITOR").Chan = modConfig.AltAFMonitorChan
        
        'Return error flow to normal
        On Error Resume Next
    
    End If
    
    'Run the cancel button code
    cmdCancel_Click
    
    'Exit the subroutine before the error handling section at the end
    Exit Sub
    
BadWave:
    
    Resume Next
    
    'Raise an error
    Err.Raise Err.number, _
              "frmSettings->frmADWIN_AF_CommSettings.cmdApply_Click", _
              "Bad Wave name used to access Wave object in the System Waves " & _
              "Collection.  Please check the your Paleomag.ini file format."
    
End Sub

Private Sub cmdCancel_Click()

    Me.Hide
    Unload Me
        
End Sub

Private Sub cmdOK_Click()

    Dim Response As Long
    
    'First confirm with the user that these changes should be made
    Response = MsgBox("Incorrect values can break the system!" & vbCrLf & _
                      "Are you sure you want to make changes?", _
                      vbYesNo, _
                      "Warning!")
                      
    If Response = vbYes Then

        'Need to write these settings to the global variables
        
        'Text settings first
        modConfig.AFAxialLowFieldMaxValue = val(Me.txtAxialLowFieldMax)
        modConfig.AFTransLowFieldMaxValue = val(Me.txtTransLowFieldMax)
        modConfig.AFUnits = Me.cmbUnits.List(cmbUnits.ListIndex)
        
        'Alt AF monitor checkbox setting
        modConfig.EnableAltAFMonitor = (Me.checkAltAFMonitorEnabled.Value = Checked)
        
        'Now need to get the channel settings loaded into the correct
        'wave and channel objects
        
        'Change the Error handling so that we capture errors into a labelled
        'section at the end of the subroutine
        On Error GoTo BadWave:
            
            'Analog Output
            
            'Save the AF Ramp Analog output channel settings
            Set modConfig.AFRampChan = frmSettings.SaveDAQSetting(Me.cmbAFRampBoard, _
                                                                  Me.cmbAFRampChan, _
                                                                  "AO", _
                                                                  "AF Ramp Analog Output")
                                                                  
            'Store this setting to the AFRAMPUP & AFRAMPDOWN wave objects
            Set modConfig.WaveForms("AFRAMPUP").Chan = modConfig.AFRampChan
            Set modConfig.WaveForms("AFRAMPDOWN").Chan = modConfig.AFRampChan
                                                                  
                                                                  
            'Analog Input
            
            'Save the AF Low-Field Monitor Analog Input channel settings
            Set modConfig.AFLFMonitorChan = _
                            frmSettings.SaveDAQSetting(Me.cmbAFLFMonitorBoard, _
                                                       Me.cmbAFLFMonitorChan, _
                                                       "AI", _
                                                       "AF Low-Field Monitor Analog Input")
            
            'Store this setting to the AFLFMONITOR wave object
            Set modConfig.WaveForms("AFLFMONITOR").Chan = modConfig.AFLFMonitorChan
            
            
            'Save the AF High-Field Monitor Analog Input channel settings
            Set modConfig.AFMonitorChan = _
                            frmSettings.SaveDAQSetting(Me.cmbAFHFMonitorBoard, _
                                                       Me.cmbAFMonitorChan, _
                                                       "AI", _
                                                       "AF High-Field Monitor Analog Input")
                                                       
            'Store this setting to the AFHFMONITOR wave object
            Set modConfig.WaveForms("AFMONITOR").Chan = modConfig.AFMonitorChan
                                                       
                                                       
            'Save the Alternate AF Monitor Analog Input channel settings
            Set modConfig.AltAFMonitorChan = _
                            frmSettings.SaveDAQSetting(Me.cmbAltAFMonitorBoard, _
                                                       Me.cmbAltAFMonitorChan, _
                                                       "AI", _
                                                       "Alternate AF Monitor Analog Input")
                                                       
            'Store this setting to the ALTAFMONITOR wave object
            Set modConfig.WaveForms("ALTAFMONITOR").Chan = modConfig.AltAFMonitorChan
        
        'Return error flow to normal
        On Error GoTo 0
    
    End If
    
    'Now, change the necessary AF comm fields in the .ini file
    'while changing as few of the other .ini fields as possible
    modConfig.Config_writeAFCommSettingstoINI
    
    'Run the cancel button code
    cmdCancel_Click
    
    'Exit the subroutine before the error handling section at the end
    Exit Sub
    
BadWave:
    
    Resume Next
    
    'Raise an error
    Err.Raise Err.number, _
              "frmADWIN_AF_CommSettings.cmdApply_Click", _
              "Bad Wave name used to access Wave object in the System Waves " & _
              "Collection.  Please check the your Paleomag.ini file format."
    
End Sub

Private Sub Form_Load()

    'Import values for text fields from the global variables
    Me.txtAxialLowFieldMax = Trim(Str(modConfig.AFAxialLowFieldMaxValue))
    Me.txtTransLowFieldMax = Trim(Str(modConfig.AFTransLowFieldMaxValue))
    
    'Import value for Alt AF monitoring checkbox
    If modConfig.EnableAltAFMonitor = True Then
    
        Me.checkAltAFMonitorEnabled.Value = Checked
        
    Else
        
        Me.checkAltAFMonitorEnabled.Value = Unchecked
    
    End If

    'Check which AF system is being used
    'if it's ADWIN, enabled all the controls
    'if it's 2G disable all controls except
    'the Alternate AF Monitor controls
    Me.frameAFHFMonitor.Enabled = (AFSystem = "ADWIN")
    Me.frameAFLFMonitor.Enabled = (AFSystem = "ADWIN")
    Me.frameAFRamp.Enabled = (AFSystem = "ADWIN")
    Me.frameLowFieldMaxValue.Enabled = (AFSystem = "ADWIN")
    Me.cmbAFHFMonitorBoard.Enabled = (AFSystem = "ADWIN")
    Me.cmbAFMonitorChan.Enabled = (AFSystem = "ADWIN")
    Me.cmbAFLFMonitorBoard.Enabled = (AFSystem = "ADWIN")
    Me.cmbAFLFMonitorChan.Enabled = (AFSystem = "ADWIN")
    Me.cmbAFRampBoard.Enabled = (AFSystem = "ADWIN")
    Me.cmbAFRampChan.Enabled = (AFSystem = "ADWIN")
    Me.cmbUnits.Enabled = (AFSystem = "ADWIN")
    Me.txtAxialLowFieldMax.Enabled = (AFSystem = "ADWIN")
    Me.txtTransLowFieldMax.Enabled = (AFSystem = "ADWIN")
    
    'Unlock all the channel & board combo-boxes
    Me.cmbAFHFMonitorBoard.locked = False
    Me.cmbAFMonitorChan.locked = False
    Me.cmbAFLFMonitorBoard.locked = False
    Me.cmbAFLFMonitorChan.locked = False
    Me.cmbAFRampBoard.locked = False
    Me.cmbAFRampChan.locked = False
    Me.cmbUnits.locked = False
    Me.cmbAltAFMonitorBoard.locked = False
    Me.cmbAltAFMonitorChan.locked = False
                              
    'Make sure the Alterate AF Monitor frame and combo boxes are enabled
    Me.cmbAltAFMonitorBoard.Enabled = True
    Me.cmbAltAFMonitorChan.Enabled = True
    Me.frameAltAFMonitor.Enabled = True
    Me.checkAltAFMonitorEnabled.Enabled = True
    
    'Now need to clear and then load the board and channel comboboxes
    ClearComboBoxes
    LoadBoardChanComboBoxes
    
    'Now import settings from Wave Object to select which boards / channels
    'are selected
    
    'Analog Input for Alternate AF Monitor
    frmSettings.SetBoardAndChanComboBoxes Me.cmbAltAFMonitorBoard, _
                                          Me.cmbAltAFMonitorChan, _
                                          modConfig.AltAFMonitorChan, _
                                          "Alternate AF Monitor"
                                        
    'Analog Input for AF Low-Field Monitor
    frmSettings.SetBoardAndChanComboBoxes Me.cmbAFLFMonitorBoard, _
                                          Me.cmbAFLFMonitorChan, _
                                          modConfig.AFLFMonitorChan, _
                                          "AF Low-Field Monitor"
                                          
    'Analog Input for AF High-Field Monitor
    frmSettings.SetBoardAndChanComboBoxes Me.cmbAFHFMonitorBoard, _
                                          Me.cmbAFMonitorChan, _
                                          modConfig.AFMonitorChan, _
                                          "AF High-Field Monitor"
                                          
    'Analog Output for AF Ramp
    frmSettings.SetBoardAndChanComboBoxes Me.cmbAFRampBoard, _
                                          Me.cmbAFRampChan, _
                                          modConfig.AFRampChan, _
                                          "AF Ramp"
    
    'Now load units into the cmbUnits control
    cmbUnits.Clear
    cmbUnits.AddItem "G", 0
    
    'Select Units from Global variables
    If modConfig.AFUnits = "G" Then
    
        'this should deactivate the cmbUnits click event from setting the list index
        PriorListIndex = 0
        cmbUnits.ListIndex = 0
        
    End If
        
    Me.Height = 6390
    Me.Width = 5895
        
End Sub
Private Sub ClearComboBoxes()

    'Clear the combo boxes
    Me.cmbAFHFMonitorBoard.Clear
    Me.cmbAFMonitorChan.Clear
    Me.cmbAFLFMonitorBoard.Clear
    Me.cmbAFLFMonitorChan.Clear
    Me.cmbAFRampBoard.Clear
    Me.cmbAFRampChan.Clear
    Me.cmbAltAFMonitorBoard.Clear
    Me.cmbAltAFMonitorChan.Clear

End Sub
Private Sub LoadBoardChanComboBoxes()

    Dim i As Long
    Dim N As Long

    'Get the number of boards in the System Boards collection
    N = SystemBoards.Count

    'Need to load the board combo boxes from the System DAQ Boards collection now
    'For each Board in the System Boards collection, add on new item into the combo-boxes
    For i = 1 To N
    
        Me.cmbAFHFMonitorBoard.AddItem SystemBoards(i).BoardName
        Me.cmbAFLFMonitorBoard.AddItem SystemBoards(i).BoardName
        Me.cmbAFRampBoard.AddItem SystemBoards(i).BoardName
        Me.cmbAltAFMonitorBoard.AddItem SystemBoards(i).BoardName
        
    Next
        
    'Set the active list-index for all the board combo boxes
    Me.cmbAFHFMonitorBoard.ListIndex = 0
    Me.cmbAFLFMonitorBoard.ListIndex = 0
    Me.cmbAFRampBoard.ListIndex = 0
    Me.cmbAltAFMonitorBoard.ListIndex = 0
        
    'Now call the board combo-box change event for every board
    'These functions will load the possible channels into the
    'matching channel combo-boxes
    PriorListIndex = -1
    cmbAFHFMonitorBoard_Click
    PriorListIndex = -1
    cmbAFLFMonitorBoard_Click
    PriorListIndex = -1
    cmbAFRampBoard_Click
    PriorListIndex = -1
    cmbAltAFMonitorBoard_Click
    
End Sub
Private Sub cmbAFRampBoard_Click()
    
    If PriorListIndex <> cmbAFRampBoard.ListIndex And _
       PriorListIndex <> -1 _
    Then
        
        frmSettings.LoadChannelComboBox cmbAFRampChan, _
                                        AFRampChan, _
                                        cmbAFRampBoard
                                        
    ElseIf PriorListIndex = -1 Then
        
        frmSettings.LoadChannelComboBox cmbAFRampChan, _
                                        AFRampChan
        
    End If
    
    PriorListIndex = cmbAFRampBoard.ListIndex
    
End Sub
Private Sub cmbAFLFMonitorBoard_Click()
        
    If PriorListIndex <> cmbAFLFMonitorBoard.ListIndex And _
       PriorListIndex <> -1 _
    Then
        
        frmSettings.LoadChannelComboBox cmbAFLFMonitorChan, _
                                        AFLFMonitorChan, _
                                        cmbAFLFMonitorBoard
                                        
    ElseIf PriorListIndex = -1 Then
        
        frmSettings.LoadChannelComboBox cmbAFLFMonitorChan, _
                                        AFLFMonitorChan
        
    End If
    
    PriorListIndex = cmbAFLFMonitorBoard.ListIndex
    
End Sub
Private Sub cmbAFHFMonitorBoard_Click()
        
    If PriorListIndex <> cmbAFHFMonitorBoard.ListIndex And _
       PriorListIndex <> -1 _
    Then
        
        frmSettings.LoadChannelComboBox cmbAFMonitorChan, _
                                        AFMonitorChan, _
                                        cmbAFHFMonitorBoard
                                        
    ElseIf PriorListIndex = -1 Then
        
        frmSettings.LoadChannelComboBox cmbAFMonitorChan, _
                                        AFMonitorChan
        
    End If
    
    PriorListIndex = cmbAFHFMonitorBoard.ListIndex
    
End Sub

Private Sub cmbAltAFMonitorBoard_Click()
    
    If PriorListIndex <> cmbAltAFMonitorBoard.ListIndex And _
       PriorListIndex <> -1 _
    Then
        
        frmSettings.LoadChannelComboBox cmbAltAFMonitorChan, _
                                        AltAFMonitorChan, _
                                        cmbAltAFMonitorBoard
                                        
    ElseIf PriorListIndex = -1 Then
        
        frmSettings.LoadChannelComboBox cmbAltAFMonitorChan, _
                                        AltAFMonitorChan
        
    End If
    
    PriorListIndex = cmbAltAFMonitorBoard.ListIndex
    
End Sub

Private Sub txtAxialLowFieldMax_LostFocus()

    If val(Me.txtAxialLowFieldMax) < 0 Then
    
        Me.txtAxialLowFieldMax = "0"
        
    End If
    
    If val(Me.txtAxialLowFieldMax) > modConfig.AfAxialMax Then
    
        Me.txtAxialLowFieldMax = Trim(Str(modConfig.AfAxialMax))
        
    End If
          
End Sub

Private Sub txtTransLowFieldMax_LostFocus()

    If val(Me.txtTransLowFieldMax) < 0 Then
    
        Me.txtTransLowFieldMax = "0"
        
    End If
    
    If val(Me.txtTransLowFieldMax) > modConfig.AfTransMax Then
    
        Me.txtTransLowFieldMax = Trim(Str(modConfig.AfTransMax))
        
    End If
          
End Sub
