VERSION 5.00
Begin VB.Form frmWebcam 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Webcam"
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4095
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   4095
   Begin VB.PictureBox picOutput 
      Height          =   3135
      Left            =   0
      ScaleHeight     =   3075
      ScaleWidth      =   4035
      TabIndex        =   0
      Top             =   0
      Width           =   4095
   End
End
Attribute VB_Name = "frmWebcam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()

    '(March 11, 2011 - I Hilburn)
    'Update all the form's icons with the program Ico file
    SetFormIcon Me

End Sub

Private Sub picOutput_Click()
    ' (February 2010 L Carporzen) Webcam
    frmProgram.cmdStart.Enabled = True
    frmProgram.cmdStop.Visible = False
    'Make sure to disconnect from capture source!!!
    DoEvents: SendMessage mCapHwnd, Disconnect, 0, 0
    Clipboard.Clear
    Unload Me
    Me.Hide
End Sub

