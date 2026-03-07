VERSION 5.00
Begin VB.Form frmTip 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reminder"
   ClientHeight    =   3465
   ClientLeft      =   2355
   ClientTop       =   2385
   ClientWidth     =   4830
   Icon            =   "frmTip.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   4830
   Visible         =   0   'False
   WhatsThisHelp   =   -1  'True
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   2715
      Left            =   120
      Picture         =   "frmTip.frx":0ECA
      ScaleHeight     =   2655
      ScaleWidth      =   4515
      TabIndex        =   1
      Top             =   120
      Width           =   4575
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Make sure power is on to the following devices:"
         Height          =   435
         Left            =   540
         TabIndex        =   3
         Top             =   180
         Width           =   3495
      End
      Begin VB.Label lblTipText 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Tip Here"
         Height          =   1755
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   4215
      End
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   3240
      TabIndex        =   0
      Top             =   3000
      Width           =   1455
   End
End
Attribute VB_Name = "frmTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
    Me.Hide
    ' Show the Main form if we've cleared the tip
    frmProgram.SignalReady
    Unload Me
End Sub

Private Sub Form_Load()
    
    Dim itemNum As Integer
    
    '(March 11, 2011 - I Hilburn)
    'Update all the form's icons with the program Ico file
    SetFormIcon Me
    
    ' Center the Tip box above the Login form
    Left = (Screen.Width / 2) - (Width / 2)
    Top = (Screen.Height / 2) - Height + 300
    lblTipText.Caption = _
        "1. All three 2G SQUID boxes" & vbCrLf & _
        "2. Motor Driver Box"
    itemNum = 3
        If EnableSusceptibility Then
        lblTipText.Caption = lblTipText.Caption & vbCrLf & _
        Format$(itemNum, "0. ") & "Bartington susceptibility bridge (set to CGS and " & _
        Format$(SusceptibilityScaleFactor, "0.0") & ")"
        itemNum = itemNum + 1
    End If
    If EnableAF Then
        lblTipText.Caption = lblTipText.Caption & vbCrLf & _
        Format$(itemNum, "0. ") & "AF units and cooling air, with the AF degausser on computer control (if you are going to do AF or rockmag)"
        itemNum = itemNum + 1
    End If
    If EnableAxialIRM Then
        lblTipText.Caption = lblTipText.Caption & vbCrLf & _
        Format$(itemNum, "0. ") & "IRM pulse box (if you are going to do rockmag)"
        itemNum = itemNum + 1
    End If
    If EnableARM Then
        lblTipText.Caption = lblTipText.Caption & vbCrLf & _
        Format$(itemNum, "0. ") & "ARM bias box (if you are going to do rockmag)"
        itemNum = itemNum + 1
    End If
End Sub

