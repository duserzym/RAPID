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
Dim CurrentlyRunning As Boolean
Dim WriteFile As String
Dim iMsg, iConf, Flds
Dim schema As String
Dim SendEmailGmail As Boolean
Public LoginSucceeded As Boolean

Private Sub cmdCancel_Click()
    Dim f As Form
    If CurrentlyRunning Then ' (December 2008 L Carporzen) Monopole survey
    CurrentlyRunning = False
    On Error GoTo alive
    Set iMsg = CreateObject("CDO.Message")
    Set iConf = CreateObject("CDO.Configuration")
    Set Flds = iConf.Fields
    schema = "http://schemas.microsoft.com/cdo/configuration/"
    Flds.Item(schema & "sendusing") = 2
    Flds.Item(schema & "smtpserver") = "smtp.gmail.com"
    Flds.Item(schema & "smtpserverport") = 465
    Flds.Item(schema & "smtpauthenticate") = 1
    Flds.Item(schema & "sendusername") = "khramov.Ifz@gmail.com"
    Flds.Item(schema & "sendpassword") = "magnetometer107"
    Flds.Item(schema & "smtpusessl") = 1
    Flds.Update
    With iMsg
    .To = "khramov.ifz@gmail.com"
    .From = "RAPID <khramov.ifz@gmail.com>"
    .Subject = MailFromName
    .HTMLBody = "SQUID log"
    .Sender = MailFromName
    .Organization = MailFromName
    .ReplyTo = "khramov.ifz@gmail.com"
    Set .Configuration = iConf
    .AddAttachment (WriteFile)
    SendEmailGmail = .Send
    End With
    Set iMsg = Nothing
    Set iConf = Nothing
    Set Flds = Nothing
    Kill (WriteFile)
alive:
    End If
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
    If CurrentlyRunning Then ' (December 2008 L Carporzen) Monopole survey
    CurrentlyRunning = False
    On Error GoTo alive
    Set iMsg = CreateObject("CDO.Message")
    Set iConf = CreateObject("CDO.Configuration")
    Set Flds = iConf.Fields
    schema = "http://schemas.microsoft.com/cdo/configuration/"
    Flds.Item(schema & "sendusing") = 2
    Flds.Item(schema & "smtpserver") = "smtp.gmail.com"
    Flds.Item(schema & "smtpserverport") = 465
    Flds.Item(schema & "smtpauthenticate") = 1
    Flds.Item(schema & "sendusername") = "khramov.Ifz@gmail.com"
    Flds.Item(schema & "sendpassword") = "magnetometer107"
    Flds.Item(schema & "smtpusessl") = 1
    Flds.Update
    With iMsg
    .To = "khramov.ifz@gmail.com"
    .From = "RAPID <khramov.ifz@gmail.com>"
    .Subject = MailFromName
    .HTMLBody = "SQUID log"
    .Sender = MailFromName
    .Organization = MailFromName
    .ReplyTo = "khramov.ifz@gmail.com"
    Set .Configuration = iConf
    .AddAttachment (WriteFile)
    SendEmailGmail = .Send
    End With
    Set iMsg = Nothing
    Set iConf = Nothing
    Set Flds = Nothing
    Kill (WriteFile)
alive:
    End If
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

Public Sub RunSQUID()
    Dim CurrentData As Cartesian3D
    Dim waitingTime As Double
    Dim StartTime As String
    CurrentlyRunning = True ' (December 2008 L Carporzen) Monopole survey
    waitingTime = 1
    StartTime = CStr(Timer) 'Format(Now, "yyyy-mm-dd hh:mm:ss"))
    WriteFile = "C:\Paleomag\Paleomag 2010\SQUID" '& startTime
    Do While CurrentlyRunning
        DelayTime waitingTime
        Set CurrentData = New Cartesian3D
        CurrentData.X = 0
        CurrentData.Y = 0
        CurrentData.Z = 0
        Set CurrentData = frmSQUID.getData
        WriteSQUIDData WriteFile, Format(Now, "yyyy-mm-dd hh:mm:ss"), CurrentData
        Set CurrentData = Nothing
        If CStr(Timer) > (StartTime + 86400) Or CStr(Timer) < StartTime Then CurrentlyRunning = False ' send the email everyday or the next day
    Loop
    If CStr(Timer) > (StartTime + 86400) Or CStr(Timer) < StartTime Then ' (December 2008 L Carporzen) Monopole survey
    CurrentlyRunning = True
    On Error GoTo alive
    Set iMsg = CreateObject("CDO.Message")
    Set iConf = CreateObject("CDO.Configuration")
    Set Flds = iConf.Fields
    schema = "http://schemas.microsoft.com/cdo/configuration/"
    Flds.Item(schema & "sendusing") = 2
    Flds.Item(schema & "smtpserver") = "smtp.gmail.com"
    Flds.Item(schema & "smtpserverport") = 465
    Flds.Item(schema & "smtpauthenticate") = 1
    Flds.Item(schema & "sendusername") = "khramov.Ifz@gmail.com"
    Flds.Item(schema & "sendpassword") = "magnetometer107"
    Flds.Item(schema & "smtpusessl") = 1
    Flds.Update
    With iMsg
    .To = "khramov.ifz@gmail.com"
    .From = "RAPID <khramov.ifz@gmail.com>"
    .Subject = MailFromName
    .HTMLBody = "SQUID log partial"
    .Sender = MailFromName
    .Organization = MailFromName
    .ReplyTo = "khramov.ifz@gmail.com"
    Set .Configuration = iConf
    .AddAttachment (WriteFile)
    SendEmailGmail = .Send
    End With
    Set iMsg = Nothing
    Set iConf = Nothing
    Set Flds = Nothing
    Kill (WriteFile)
alive:
    RunSQUID
    End If
End Sub

Private Sub WriteSQUIDData(filename As String, meastime As String, data As Cartesian3D)
    Dim filenum As Integer
    filenum = FreeFile
    On Error GoTo oops
    If CurrentlyRunning Then ' (December 2008 L Carporzen) Monopole survey
    Open filename For Append As #filenum
    With data
        Print #filenum, meastime; ","; .X; ","; .Y; ","; .Z
    End With
    Close #filenum
    End If
    GoTo stillworking
oops:
    CurrentlyRunning = False
stillworking:
End Sub

