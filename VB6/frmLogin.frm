VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Paleomag Machine Login"
   ClientHeight    =   1770
   ClientLeft      =   2850
   ClientTop       =   3390
   ClientWidth     =   4830
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1041.309
   ScaleMode       =   0  'User
   ScaleWidth      =   4549.193
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtUserName 
      Height          =   345
      Left            =   1440
      TabIndex        =   1
      Top             =   120
      Width           =   3285
   End
   Begin VB.TextBox EmailAddressText 
      Height          =   345
      Left            =   1440
      TabIndex        =   2
      Top             =   600
      Width           =   3285
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   1440
      TabIndex        =   3
      Top             =   1128
      Width           =   1425
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   3240
      TabIndex        =   4
      Top             =   1128
      Width           =   1425
   End
   Begin VB.Label lblLabels 
      Caption         =   "&User Name:"
      Height          =   270
      Index           =   0
      Left            =   225
      TabIndex        =   0
      Top             =   135
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Caption         =   "Email"
      Height          =   276
      Index           =   2
      Left            =   240
      TabIndex        =   5
      Top             =   600
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public LoginSucceeded As Boolean

Private Sub cmdCancel_Click()
    Dim f As Form
    'set the global var to false
    'to denote a failed login
    LoginSucceeded = False
    Me.Hide
    'Unload Me
    ' (February 2010 L Carporzen) Cleaner way to close the program
    For Each f In Forms
     Unload f
     Next
    End
End Sub

Private Sub cmdOK_Click()
    Prog_halted = False ' (September 2007 L Carporzen) New version of the Halt button
    ' Record the login name
    LoginSucceeded = True
    LoginName = txtUserName.text
    LoginEmail = EmailAddressText.text
    Dim login_date As Date: login_date = Now
    
    
    
    AppendLog (LoginName & " " & Format$(login_date, "yyyy/m/dd hh:mm:ss"))
    Config_SaveSetting "Program", "LastLogin", LoginName
    Config_SaveSetting "Program", "LastEmail", LoginEmail
    frmTip.cmdOK.Enabled = True
    ' Hide this form
    Me.Hide
    ' Show the Main form if we've cleared the tip
    frmProgram.SignalReady
    Unload Me
End Sub

Private Sub Form_Load()
        
    Left = Screen.Width / 2 - Width / 2
    Top = Screen.Height / 2 + 340
    txtUserName.text = Config_GetSetting("Program", "LastLogin", vbNullString)
    EmailAddressText.text = Config_GetSetting("Program", "LastEmail", vbNullString)
    
    '(March 11, 2011 - I Hilburn)
    'Update all the form's icons with the program Ico file
    SetFormIcon Me
    
End Sub
