VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AdwinAfPauseConstants"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Const m_pause_between_boot_and_init_in_ms As Long = 300
Private Const m_pause_between_init_and_run_in_ms As Long = 300
Private Const m_pause_between_rampend_and_read_ramp_outputs_in_ms As Long = 300
Private Const m_pause_after_clear_all_processes_in_ms As Long = 300

Public Property Get MsecsBetweenBootAndInit() As Long

    MsecsBetweenBootAndInit = m_pause_between_boot_and_init_in_ms

End Property

Public Property Get MsecsBetweenInitAndRun() As Long

    MsecsBetweenInitAndRun = m_pause_between_init_and_run_in_ms

End Property

Public Property Get MsecsBetweenRampEndAndReadRampOutputs() As Long

    MsecsBetweenRampEndAndReadRampOutputs = m_pause_between_rampend_and_read_ramp_outputs_in_ms

End Property

Public Property Get MsecsAfterClearAllProcesses() As Long

    MsecsAfterClearAllProcesses = m_pause_after_clear_all_processes_in_ms

End Property




