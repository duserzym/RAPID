VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmSampleSelect 
   Caption         =   "Select a sample"
   ClientHeight    =   4140
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7140
   Icon            =   "frmSampleSelect.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4140
   ScaleWidth      =   7140
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      CausesValidation=   0   'False
      Height          =   372
      Left            =   5400
      TabIndex        =   0
      Top             =   3480
      Width           =   1212
   End
   Begin ComctlLib.ListView lvwSamples 
      Height          =   2655
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   4683
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
End
Attribute VB_Name = "frmSampleSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mvarFilename As String
Private mvarSamplename As String
Private SelectedFile As Boolean

Private Sub cmdOK_Click()
    mvarFilename = lvwSamples.SelectedItem.SubItems(2)
    mvarSamplename = lvwSamples.SelectedItem.text
    SelectedFile = True
    Me.Hide
End Sub

Public Function filename() As String
    WaitForUserSelection
    filename = mvarFilename
End Function

Private Sub Form_Activate()
    refreshListDisplay
End Sub

Private Sub Form_Load()

    Dim colX As ColumnHeader ' Declare variable.
    
    '(March 11, 2011 - I Hilburn)
    'Update all the form's icons with the program Ico file
    SetFormIcon Me
    
    Set colX = lvwSamples.ColumnHeaders.Add(1)
    colX.text = "Sample"
    colX.Width = Me.TextWidth(colX.text & "XX ")
    Set colX = lvwSamples.ColumnHeaders.Add(2)
    colX.text = "Sample set"
    colX.Width = Me.TextWidth(colX.text & " XX")
    Set colX = lvwSamples.ColumnHeaders.Add(3)
    colX.text = "Sample Index File"
    colX.Width = 0
    SelectedFile = False
    
End Sub

Private Sub Form_Resize()
    Me.Width = 7260
    cmdOk.Top = Me.Height - 1050
    lvwSamples.Height = Me.Height - 1900
End Sub

Public Sub refreshListDisplay()
    Dim i As Integer
    Dim SampleCount As Integer
    Dim curItem As ListItem
    lvwSamples.ListItems.Clear
    SampleCount = SampleIndexRegistry.SampleCount
    If SampleCount = 0 Then Exit Sub
    For i = 1 To SampleCount
        Set curItem = lvwSamples.ListItems.Add
        With SampleIndexRegistry.SampleByIndex(i)
            curItem.text = .Samplename
            curItem.SubItems(1) = .Parent.SampleCode
            curItem.SubItems(2) = .IndexFile
        End With
    Next i
End Sub

Public Function Samplename() As String
    WaitForUserSelection
    Samplename = mvarSamplename
End Function

Private Sub WaitForUserSelection()
    Do While Not SelectedFile
        DelayTime 0.05
    Loop
End Sub

