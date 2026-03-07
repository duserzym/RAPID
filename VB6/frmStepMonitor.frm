VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmStepMonitor 
   Caption         =   "Step Monitor"
   ClientHeight    =   4260
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7665
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4260
   ScaleWidth      =   7665
   Begin VB.Timer Timer1 
      Interval        =   10000
      Left            =   960
      Top             =   3960
   End
   Begin ComctlLib.ListView lvwRmgSteps 
      Height          =   3012
      Left            =   360
      TabIndex        =   1
      Top             =   360
      Width           =   6612
      _ExtentX        =   11668
      _ExtentY        =   5318
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
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   372
      Left            =   5880
      TabIndex        =   0
      Top             =   3720
      Width           =   1212
   End
End
Attribute VB_Name = "frmStepMonitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' This module manages the sample queue.
Option Explicit

Private Sub cmdClose_Click()
    Me.Hide
End Sub

Private Sub Form_Activate()
    refreshQueueDisplay
End Sub

Private Sub Form_Load()
    
    Dim colX As ColumnHeader ' Declare variable.
    
    '(March 11, 2011 - I Hilburn)
    'Update all the form's icons with the program Ico file
    SetFormIcon Me
    
    Set colX = lvwRmgSteps.ColumnHeaders.Add(1)
    colX.text = "#"
    colX.Width = Me.TextWidth(colX.text & " ")
    Set colX = lvwRmgSteps.ColumnHeaders.Add(2)
    colX.text = "Step Type"
    colX.Width = Me.TextWidth(colX.text & " ")
    Set colX = lvwRmgSteps.ColumnHeaders.Add(3)
    colX.text = "Level"
    colX.Width = Me.TextWidth(colX.text & " ")
    Set colX = lvwRmgSteps.ColumnHeaders.Add(4)
    colX.text = "Bias Field"
    colX.Width = Me.TextWidth(colX.text & " ")
    Set colX = lvwRmgSteps.ColumnHeaders.Add(5)
    colX.text = "Spin Speed"
    colX.Width = Me.TextWidth(colX.text & " ")
    Set colX = lvwRmgSteps.ColumnHeaders.Add(6)
    colX.text = "Measure"
    colX.Width = Me.TextWidth(colX.text & " ")
    Set colX = lvwRmgSteps.ColumnHeaders.Add(7)
    colX.text = "Measure Suscep."
    colX.Width = Me.TextWidth(colX.text & " ")
    Set colX = lvwRmgSteps.ColumnHeaders.Add(8) ' (November 2007 L Carporzen) Remarks column in RMG
    colX.text = "Remarks"
    colX.Width = Me.TextWidth(colX.text & " ")
    
End Sub

Private Sub Form_Resize()
    Me.Width = 7785
    cmdClose.Top = Me.Height - 935
    lvwRmgSteps.Height = Me.Height - 1650
End Sub

Private Sub lvwRmgSteps_click()
    On Error GoTo fin
fin:
End Sub

Public Sub refreshQueueDisplay()
    Dim i As Integer
    Dim curItem As ListItem
    lvwRmgSteps.ListItems.Clear
    If SampQueue.Count = 0 Then Exit Sub
    If SampQueue.Item(1).commandType <> "Meas" Then Exit Sub
    With SampleIndexRegistry(SampQueue.Item(1).fileid).measurementSteps
        If .Count = 0 Then Exit Sub
        For i = .CurrentStepIndex To .Count
            Set curItem = lvwRmgSteps.ListItems.Add(i + 1 - .CurrentStepIndex, .CurrentStep.key)
            With .CurrentStep
                curItem.text = .key
                curItem.SubItems(1) = .StepType
                curItem.SubItems(2) = .Level
                curItem.SubItems(3) = .BiasField
                curItem.SubItems(4) = .SpinSpeed
                curItem.SubItems(5) = .Measure
                curItem.SubItems(6) = .MeasureSusceptibility
                curItem.SubItems(7) = .Remarks ' (November 2007 L Carporzen) Remarks column in RMG
            End With
        Next i
    End With
End Sub

Private Sub Timer1_Timer()
    refreshQueueDisplay
End Sub

