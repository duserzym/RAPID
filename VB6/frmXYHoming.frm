VERSION 5.00
Begin VB.Form frmXYHoming 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1635
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9045
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1635
   ScaleWidth      =   9045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frameBorder 
      Height          =   1575
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9015
      Begin VB.Label lblXYHoming 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   480
         TabIndex        =   1
         Top             =   600
         Width           =   45
      End
   End
End
Attribute VB_Name = "frmXYHoming"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

    Dim ExitMsg As String
    Dim i As Double

    'Need to resize the shutdown window based upon the size of the computer screen
    Me.Width = Screen.Width / 2
    
    Me.Left = (Screen.Width - Me.Width) / 2
    
    'Set the shutdown message
    ExitMsg = "Please wait.  Homing XY Table"
        
    'Now need to pick a font size so that the "Homing XY Table" message fits the
    'width of the form
    For i = 8 To 200 Step 0.5
    
        Me.FontSize = i
    
        If Me.TextWidth(ExitMsg) > 0.8 * Me.ScaleWidth Then
        
            Me.FontSize = Me.FontSize - 0.5
            Exit For
            
        End If
        
    Next i
    
    Me.Height = 2 * Me.TextHeight(ExitMsg)
    Me.Top = (Screen.Height - Me.Height) / 2
    
    'Size and position the border frame
    Me.frameBorder.Width = Me.ScaleWidth - 50
    Me.frameBorder.Height = Me.ScaleHeight - 50
    Me.frameBorder.Top = 0
    Me.frameBorder.Left = 25
    
    Me.lblXYHoming.ForeColor = QBColor(0)
    Me.lblXYHoming.FontSize = Me.FontSize
    Me.lblXYHoming.Caption = ExitMsg
    Me.lblXYHoming.Top = Me.ScaleHeight / 4
    Me.lblXYHoming.Left = 0.05 * Me.ScaleWidth
    Me.lblXYHoming.Visible = True
    
    'Place the cursor at the starting point for writing the text
    
    Me.refresh

End Sub

