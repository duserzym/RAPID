VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmSampleQueueMonitor 
   Caption         =   "Sample Queue Monitor"
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
   Begin ComctlLib.ListView lvwSampQueue 
      Height          =   3012
      Left            =   360
      TabIndex        =   3
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
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   372
      Left            =   4320
      TabIndex        =   2
      Top             =   3720
      Width           =   1332
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Enabled         =   0   'False
      Height          =   372
      Left            =   2760
      TabIndex        =   1
      Top             =   3720
      Width           =   1332
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
Attribute VB_Name = "frmSampleQueueMonitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' This module manages the sample queue.

Option Explicit

Public Sub Clear()
    SampQueue.Clear
    refreshQueueDisplay
End Sub

Private Sub cmdClear_Click()
    Clear
End Sub

Private Sub cmdClose_Click()
    Me.Hide
End Sub

Private Sub cmdDelete_Click()
    Dim targetItem As ListItem
    
    If lvwSampQueue.SelectedItem.Index > 0 Then
        For Each targetItem In lvwSampQueue.ListItems
            If targetItem.Selected Then
                SampQueue.Remove targetItem.key
            End If
        Next targetItem
    Else
        cmdDelete.Enabled = False
    End If
End Sub

Private Sub Form_Activate()
    refreshQueueDisplay
End Sub

Private Sub Form_Load()

    Dim colX As ColumnHeader ' Declare variable.
    
    '(March 11, 2011 - I Hilburn)
    'Update all the form's icons with the program Ico file
    SetFormIcon Me

    If DEBUG_MODE Then frmDebug.Msg "Loading a sample queue"
    
    Set colX = lvwSampQueue.ColumnHeaders.Add(1)
    colX.text = "#"
    colX.Width = Me.TextWidth(colX.text & " ")

    Set colX = lvwSampQueue.ColumnHeaders.Add(2)
    colX.text = "Command"
    colX.Width = Me.TextWidth(colX.text & " ")
    
    Set colX = lvwSampQueue.ColumnHeaders.Add(3)
    colX.text = "Hole"
    colX.Width = Me.TextWidth(colX.text & " ")
    
    Set colX = lvwSampQueue.ColumnHeaders.Add(4)
    colX.text = "File"
    colX.Width = Me.TextWidth(colX.text & " ")
    
    Set colX = lvwSampQueue.ColumnHeaders.Add(5)
    colX.text = "Sample"
    colX.Width = Me.TextWidth(colX.text & " ")
    
End Sub

Private Sub Form_Resize()
    Me.Width = 7785
    cmdDelete.Top = Me.Height - 935
    cmdClear.Top = Me.Height - 935
    cmdClose.Top = Me.Height - 935
    lvwSampQueue.Height = Me.Height - 1650
End Sub

Private Sub lvwsampqueue_click()
    On Error GoTo fin
    If LenB(lvwSampQueue.SelectedItem.text) > 0 Then
        cmdDelete.Enabled = True
    Else
        cmdDelete.Enabled = False
    End If
fin:
End Sub

Public Sub refreshQueueDisplay()
    Dim i As Integer
    Dim curItem As ListItem
    
    lvwSampQueue.ListItems.Clear
    With SampQueue
        If .Count = 0 Then Exit Sub
        For i = 1 To .Count
            With .Item(i)
                Set curItem = lvwSampQueue.ListItems.Add(i, .key)
                curItem.text = .key
                curItem.SubItems(1) = .commandType
                curItem.SubItems(2) = .hole
                curItem.SubItems(3) = .fileid
                curItem.SubItems(4) = .Sample
            End With
        Next i
    End With

End Sub

Private Sub Timer1_Timer()
    refreshQueueDisplay
End Sub

