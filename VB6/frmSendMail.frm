VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmSendMail 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SendMail client"
   ClientHeight    =   7605
   ClientLeft      =   1755
   ClientTop       =   1710
   ClientWidth     =   12810
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7605
   ScaleWidth      =   12810
   Begin VB.CheckBox chUseSSLEncryption 
      Caption         =   "Use SSL Encryption"
      Height          =   495
      Left            =   3720
      TabIndex        =   44
      ToolTipText     =   "Use Login Authorization When Connecting to a Host"
      Top             =   480
      Width           =   1875
   End
   Begin VB.Frame frameEmailOptions 
      Caption         =   "Options"
      Height          =   2535
      Left            =   6360
      TabIndex        =   36
      Top             =   120
      Width           =   6135
      Begin VB.OptionButton optEncodeType 
         Caption         =   "UUEncode"
         Height          =   195
         Index           =   1
         Left            =   600
         TabIndex        =   42
         ToolTipText     =   "Use UU Encoding for Attachments."
         Top             =   1320
         Width           =   1095
      End
      Begin VB.OptionButton optEncodeType 
         Caption         =   "MIME"
         Height          =   195
         Index           =   0
         Left            =   600
         TabIndex        =   41
         ToolTipText     =   "Use MIME encoding for Mail & Attachments."
         Top             =   1080
         Value           =   -1  'True
         Width           =   915
      End
      Begin VB.ComboBox cboPriority 
         Height          =   315
         Left            =   2040
         TabIndex        =   40
         Text            =   "cboPriority"
         ToolTipText     =   "Sets the Prioirty of the Mail Message"
         Top             =   1620
         Width           =   1410
      End
      Begin VB.CheckBox ckReceipt 
         Caption         =   "Request Check for Receipt of Email"
         Height          =   255
         Left            =   615
         TabIndex        =   39
         ToolTipText     =   "Request a Return Receipt"
         Top             =   2040
         Width           =   2955
      End
      Begin VB.CheckBox chPlainText 
         Caption         =   "Send as Plain Text (Default)"
         Height          =   195
         Left            =   600
         TabIndex        =   38
         ToolTipText     =   "Mail Body is HTML / Plain Text"
         Top             =   720
         Width           =   2355
      End
      Begin VB.CheckBox ckHtml 
         Caption         =   "Send as HTML"
         Height          =   195
         Left            =   600
         TabIndex        =   37
         ToolTipText     =   "Mail Body is HTML / Plain Text"
         Top             =   360
         Width           =   1395
      End
      Begin VB.Label Label5 
         Caption         =   "Message Priority:"
         Height          =   255
         Left            =   600
         TabIndex        =   43
         Top             =   1680
         Width           =   1215
      End
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1920
      PasswordChar    =   "*"
      TabIndex        =   35
      Top             =   1800
      Width           =   2250
   End
   Begin VB.TextBox txtUserName 
      Height          =   285
      Left            =   1920
      TabIndex        =   32
      Top             =   1440
      Width           =   2250
   End
   Begin VB.CheckBox ckLogin 
      Caption         =   "Requires Login"
      Height          =   255
      Left            =   1920
      TabIndex        =   31
      ToolTipText     =   "Use Login Authorization When Connecting to a Host"
      Top             =   960
      Width           =   1515
   End
   Begin VB.TextBox txtSMTPPort 
      Height          =   285
      Left            =   1920
      TabIndex        =   30
      Top             =   480
      Width           =   1080
   End
   Begin MSComDlg.CommonDialog cmDialog 
      Left            =   8880
      Top             =   5880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtBcc 
      Height          =   285
      Left            =   1800
      TabIndex        =   27
      Top             =   5520
      Width           =   4200
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   315
      Left            =   3480
      TabIndex        =   26
      Top             =   2640
      Width           =   1275
   End
   Begin VB.ListBox lstStatus 
      BackColor       =   &H8000000F&
      Height          =   1230
      Left            =   1860
      TabIndex        =   24
      Top             =   6240
      Width           =   5400
   End
   Begin VB.TextBox txtCcName 
      Height          =   285
      Left            =   1800
      TabIndex        =   5
      Top             =   4800
      Width           =   4200
   End
   Begin VB.TextBox txtCc 
      Height          =   285
      Left            =   1800
      TabIndex        =   6
      Top             =   5160
      Width           =   4200
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Browse..."
      Height          =   315
      Left            =   7560
      TabIndex        =   10
      Top             =   5880
      Width           =   1275
   End
   Begin VB.TextBox txtAttach 
      Height          =   285
      Left            =   7560
      TabIndex        =   9
      Top             =   5520
      Width           =   4200
   End
   Begin VB.TextBox txtMsg 
      Height          =   1620
      Left            =   7560
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Top             =   3720
      Width           =   4920
   End
   Begin VB.TextBox txtSubject 
      Height          =   285
      Left            =   7560
      TabIndex        =   7
      Top             =   3360
      Width           =   4920
   End
   Begin VB.TextBox txtFrom 
      Height          =   285
      Left            =   1800
      TabIndex        =   2
      Top             =   3660
      Width           =   4200
   End
   Begin VB.TextBox txtFromName 
      Height          =   285
      Left            =   1800
      TabIndex        =   1
      Top             =   3300
      Width           =   4200
   End
   Begin VB.TextBox txtTo 
      Height          =   285
      Left            =   1800
      TabIndex        =   4
      Top             =   4440
      Width           =   4200
   End
   Begin VB.TextBox txtToName 
      Height          =   285
      Left            =   1800
      TabIndex        =   3
      Top             =   4020
      Width           =   4200
   End
   Begin VB.TextBox txtServer 
      Height          =   285
      Left            =   1920
      TabIndex        =   0
      Top             =   75
      Width           =   4200
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset"
      Height          =   315
      Left            =   1920
      TabIndex        =   12
      Top             =   2640
      Width           =   1275
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Height          =   315
      Left            =   360
      TabIndex        =   11
      Top             =   2640
      Width           =   1275
   End
   Begin VB.Label Label4 
      Caption         =   "Password:"
      Height          =   255
      Left            =   360
      TabIndex        =   34
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Username:"
      Height          =   255
      Left            =   360
      TabIndex        =   33
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label lblSMTPServerPort 
      Caption         =   "SMTP Port:"
      Height          =   255
      Left            =   360
      TabIndex        =   29
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label lblBcc 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bcc: Email"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   600
      TabIndex        =   28
      Top             =   5520
      Width           =   915
   End
   Begin VB.Label lblStatus 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Status"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   600
      TabIndex        =   25
      Top             =   6240
      Width           =   555
   End
   Begin VB.Label lblProgress 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Progress  "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   360
      TabIndex        =   23
      Top             =   6840
      Width           =   870
   End
   Begin VB.Label lblCcName 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cc: Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   720
      TabIndex        =   22
      Top             =   4800
      Width           =   840
   End
   Begin VB.Label lblCC 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cc: Email"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   720
      TabIndex        =   21
      Top             =   5160
      Width           =   810
   End
   Begin VB.Label lblAttach 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Attachment"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   6360
      TabIndex        =   20
      Top             =   5520
      Width           =   975
   End
   Begin VB.Label lblMsg 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Message"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   6360
      TabIndex        =   19
      Top             =   3720
      Width           =   765
   End
   Begin VB.Label lblSubject 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Subject"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   6360
      TabIndex        =   18
      Top             =   3360
      Width           =   660
   End
   Begin VB.Label lblFrom 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sender Email"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   480
      TabIndex        =   17
      Top             =   3720
      Width           =   1125
   End
   Begin VB.Label lblFromName 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sender Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   360
      TabIndex        =   16
      Top             =   3360
      Width           =   1155
   End
   Begin VB.Label lblTo 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Recipient Email"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   15
      Top             =   4440
      Width           =   1335
   End
   Begin VB.Label lblToName 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Recipient Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   14
      Top             =   4080
      Width           =   1365
   End
   Begin VB.Label lblServer 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SMTP Server"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   360
      TabIndex        =   13
      Top             =   105
      Width           =   1140
   End
End
Attribute VB_Name = "frmSendMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Based on vbSendMail example client

Option Explicit
Option Compare Text

' *****************************************************************************
' Required declaration of the vbSendMail component (withevents is optional)
' You also need a reference to the vbSendMail component in the Project References
' *****************************************************************************
Private WithEvents poSendMail As vbSendMail.clsSendMail
Attribute poSendMail.VB_VarHelpID = -1

' misc local vars
Dim bAuthLogin As Boolean
Dim bPopLogin As Boolean
Dim bHtml As Boolean
Dim MyEncodeType As ENCODE_METHOD
Dim etPriority As MAIL_PRIORITY
Dim bReceipt As Boolean

Private Sub AlignControlsLeft(StandardizeWidth As Boolean, base As Object, ParamArray cnts())

    ' This is a modified version of a function in the SmartForm class,
    ' which is available on http://www.freevbcode.com
    On Error Resume Next

    Dim i As Integer
    For i = 0 To UBound(cnts)
        cnts(i).Left = base.Left
        If StandardizeWidth Then cnts(i).Width = base.Width
    Next

End Sub

Public Sub AlignControlsTop(StandardizeHeight As Boolean, base As Object, ParamArray cnts())

    ' This is a modified version of a function in the SmartForm class,
    ' which is available on http://www.freevbcode.com

    On Error Resume Next
    Dim i As Integer
    For i = 0 To UBound(cnts)
        cnts(i).Top = base.Top
        If StandardizeHeight Then cnts(i).Height = base.Height
    Next

End Sub

Private Sub cboPriority_Click()

    Select Case cboPriority.ListIndex

        Case 0: etPriority = NORMAL_PRIORITY
        Case 1: etPriority = HIGH_PRIORITY
        Case 2: etPriority = LOW_PRIORITY

    End Select

End Sub

Private Sub cboPriority_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode

        Case 38, 40

        Case Else: KeyCode = 0

    End Select

End Sub

Private Sub cboPriority_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub CenterControlHorizontal(child As Object)

    child.Left = (Me.ScaleWidth - child.Width) / 2

End Sub

Public Sub CenterControlRelativeVertical(ctl As Object, RelativeTo As Object)

    ' This is a modified version of a function in the SmartForm class,
    ' which is available on http://www.freevbcode.com

    On Error Resume Next
    ctl.Top = RelativeTo.Top + ((RelativeTo.Height - ctl.Height) / 2)

End Sub

Public Sub CenterControlsHorizontal(space As Single, AlignTop As Boolean, ParamArray cnts())

    ' This is a modified version of a function in the SmartForm class,
    ' which is available on http://www.freevbcode.com

    Dim sngTotalSpace As Single
    Dim i As Integer
    Dim sngBaseTop As Single
    Dim sngParentWidth As Single

    sngParentWidth = Me.ScaleWidth

    For i = 0 To UBound(cnts)
        sngTotalSpace = sngTotalSpace + cnts(i).Width
    Next

    sngTotalSpace = sngTotalSpace + (space * (UBound(cnts)))

    cnts(0).Left = (sngParentWidth - sngTotalSpace) / 2
    sngBaseTop = cnts(0).Top

    For i = 1 To UBound(cnts)
        cnts(i).Left = cnts(i - 1).Left + cnts(i - 1).Width + space
        If AlignTop Then cnts(i).Top = sngBaseTop
    Next

End Sub



Private Sub ckHtml_Click()

    If ckHtml.Value = vbChecked Then bHtml = True Else bHtml = False

End Sub

Private Sub ckLogin_Click()

    If ckLogin.Value = vbChecked Then
        bAuthLogin = True
        Me.txtUserName.Enabled = True
        Me.txtPassword.Enabled = True
    Else
        bAuthLogin = False
        Me.txtUserName.Enabled = False
        Me.txtPassword.Enabled = False
    End If

End Sub

Private Sub ckReceipt_Click()

    If ckReceipt.Value = vbChecked Then bReceipt = True Else bReceipt = False

End Sub

Public Sub ClearTextBoxesOnForm()

    ' Snippet Taken From http://www.freevbcode.com

    Dim ctl As Control
    For Each ctl In Me.Controls
        If TypeOf ctl Is TextBox Then
            ctl.text = vbNullString
        End If
    Next

End Sub

Private Sub cmdBrowse_Click()

    Dim sFilenames() As String
    Dim i As Integer

    On Local Error GoTo Err_Cancel

    With cmDialog
        .filename = vbNullString
        .CancelError = True
        .filter = "All Files (*.*)|*.*|HTML Files (*.htm;*.html;*.shtml)|*.htm;*.html;*.shtml|Images (*.bmp;*.jpg;*.gif)|*.bmp;*.jpg;*.gif"
        .FilterIndex = 1
        .DialogTitle = "Select File Attachment(s)"
        .MaxFileSize = &H7FFF
        .flags = &H4 Or &H800 Or &H40000 Or &H200 Or &H80000
        .ShowOpen
        ' get the selected name(s)
        sFilenames = Split(.filename, vbNullChar)
    End With

    If UBound(sFilenames) = 0 Then
        If txtAttach.text = vbNullString Then
            txtAttach.text = sFilenames(0)
        Else
            txtAttach.text = txtAttach.text & ";" & sFilenames(0)
        End If
    ElseIf UBound(sFilenames) > 0 Then
        If Right$(sFilenames(0), 1) <> "\" Then sFilenames(0) = sFilenames(0) & "\"
        For i = 1 To UBound(sFilenames)
            If txtAttach.text = vbNullString Then
                txtAttach.text = sFilenames(0) & sFilenames(i)
            Else
                txtAttach.text = txtAttach.text & ";" & sFilenames(0) & sFilenames(i)
            End If
        Next
    Else
        Exit Sub
    End If

Err_Cancel:

End Sub

Private Sub cmdExit_Click()

    Me.Hide

End Sub

Private Sub cmdReset_Click()

    ClearTextBoxesOnForm
    lstStatus.Clear
    lblProgress = vbNullString
    RetrieveSavedValues

End Sub

Private Sub cmdSend_Click()

    ' *****************************************************************************
    ' This is where all of the Components Properties are set / Methods called
    ' *****************************************************************************

    cmdSend.Enabled = False
    lstStatus.Clear
    Screen.MousePointer = vbHourglass

    '---Send Mail Bug Fix - wrong delimeter error------------------------------------------------------------
    '
    '   November 10, 2009
    '   Isaac Hilburn
    '
    '   Need to replace all comma delimiters in the emails entered by the user with semi-colons, otherwise
    '   the sendmail will not work, the email will not be sent, and no error will be raised
    '--------------------------------------------------------------------------------------------------------

    txtFrom.text = Replace(txtFrom.text, ",", ";")
    txtTo.text = Replace(txtTo.text, ",", ";")
    txtCc.text = Replace(txtCc.text, ",", ";")
    txtBcc.text = Replace(txtBcc.text, ",", ";")

    '--------------------------------------------------------------------------------------------------------
    '--------------------------------------------------------------------------------------------------------



    With poSendMail

        ' **************************************************************************
        ' Optional properties for sending email, but these should be set first
        ' if you are going to use them
        ' **************************************************************************

        .SMTPHostValidation = VALIDATE_HOST_NONE         ' Optional, default = VALIDATE_HOST_DNS
        .EmailAddressValidation = VALIDATE_SYNTAX   ' Optional, default = VALIDATE_SYNTAX
        .Delimiter = ";"                            ' Optional, default = ";" (semicolon)

        ' **************************************************************************
        ' Basic properties for sending email
        ' **************************************************************************
        .SMTPHost = txtServer.text                  ' Required the fist time, optional thereafter
        .From = txtFrom.text                        ' Required the fist time, optional thereafter
        .FromDisplayName = txtFromName.text         ' Optional, saved after first use
        .Recipient = txtTo.text                     ' Required, separate multiple entries with delimiter character
        .RecipientDisplayName = txtToName.text      ' Optional, separate multiple entries with delimiter character
        .CcRecipient = txtCc                        ' Optional, separate multiple entries with delimiter character
        .CcDisplayName = txtCcName                  ' Optional, separate multiple entries with delimiter character
        .BccRecipient = txtBcc                      ' Optional, separate multiple entries with delimiter character
        .ReplyToAddress = txtFrom.text              ' Optional, used when different than 'From' address
        .Subject = txtSubject.text                  ' Optional
        .message = txtMsg.text                      ' Optional
        .Attachment = Trim(txtAttach.text)          ' Optional, separate multiple entries with delimiter character

        ' **************************************************************************
        ' Additional Optional properties, use as required by your application / environment
        ' **************************************************************************
        .AsHTML = bHtml                             ' Optional, default = FALSE, send mail as html or plain text
        .ContentBase = vbNullString                           ' Optional, default = Null String, reference base for embedded links
        .EncodeType = MyEncodeType                  ' Optional, default = MIME_ENCODE
        .Priority = etPriority                      ' Optional, default = PRIORITY_NORMAL
        .Receipt = bReceipt                         ' Optional, default = FALSE
        .UseAuthentication = bAuthLogin             ' Optional, default = FALSE
        .UsePopAuthentication = bPopLogin           ' Optional, default = FALSE
        .UserName = txtUserName                     ' Optional, default = Null String
        .Password = txtPassword                     ' Optional, default = Null String, value is NOT saved

        .MaxRecipients = 100                        ' Optional, default = 100, recipient count before error is raised

        ' **************************************************************************
        ' Advanced Properties, change only if you have a good reason to do so.
        ' **************************************************************************
        ' .ConnectTimeout = 10                      ' Optional, default = 10
        ' .ConnectRetry = 5                         ' Optional, default = 5
        ' .MessageTimeout = 60                      ' Optional, default = 60
        ' .PersistentSettings = True                ' Optional, default = TRUE
        If Me.txtSMTPPort.text = "" Then Me.txtSMTPPort.text = CStr(modConfig.MailSMTPPort)
        .SMTPPort = CLng(Me.txtSMTPPort.text)                            ' Optional, default = 25


        'Now check the SSL settings
        If Me.chUseSSLEncryption.Value = Checked Then

            SendSSLEncryptedEmailToSMTPServer

        Else

            ' **************************************************************************
            ' OK, all of the properties are set, send the email...
            ' **************************************************************************
            ' .Connect                                  ' Optional, use when sending bulk mail
            .Send                                       ' Required
            ' .Disconnect                               ' Optional, use when sending bulk mail
            txtServer.text = .SMTPHost                  ' Optional, re-populate the Host in case
            ' MX look up was used to find a host
        End If

    End With
    Screen.MousePointer = vbDefault
    cmdSend.Enabled = True

End Sub

Private Sub Form_Load()

    ' *****************************************************************************
    ' Required to activate the vbSendMail component.
    ' *****************************************************************************
    Set poSendMail = New clsSendMail

    '(March 11, 2011 - I Hilburn)
    'Update all the form's icons with the program Ico file
    SetFormIcon Me

    With Me
        .Move (Screen.Width - .Width) / 2, (Screen.Height - .Height) / 2
        .lblProgress = vbNullString
    End With

    cboPriority.Clear
    cboPriority.AddItem "Normal"
    cboPriority.AddItem "High"
    cboPriority.AddItem "Low"
    cboPriority.ListIndex = 0

    If modConfig.MailUseSSLEncryption Then Me.chUseSSLEncryption.Value = Checked

End Sub

Private Sub Form_Unload(Cancel As Integer)

    ' *****************************************************************************
    ' Unload the component before quiting.
    ' *****************************************************************************

    Set poSendMail = Nothing

End Sub

Public Sub MailNotification(Subject As String, Body As String, Optional ByVal CodeLevel = CodeGreen, Optional ByVal VerboseStatusUpdate = False)
    Dim fileid As String
    Dim sampleid As String
    Dim MsgHead As String
    Dim MsgFoot As String
    Dim LogLine As String

    If DEBUG_MODE Then frmDebug.Msg Subject
    sampleid = vbNullString
    fileid = vbNullString
    LogLine = vbNullString
    MsgHead = LoginName & ":" & vbCrLf & vbCrLf
    If SampQueue.Count > 0 Then
        fileid = SampQueue(0).fileid
        If Changer_ValidSlot(SampQueue(1).hole) Then
            sampleid = MainChanger.ChangerSampleName(SampQueue(1).hole)
        End If
    End If

    MsgFoot = vbCrLf

    If LenB(SampleNameCurrent) > 0 Then
        MsgFoot = MsgFoot & vbCrLf & "Sample: " & SampleNameCurrent
    End If
    If LenB(SampleStepCurrent) > 0 Then
        MsgFoot = MsgFoot & vbCrLf & "Step: " & SampleStepCurrent
        LogLine = LogLine & SampleStepCurrent
    End If
    If SampleOrientationCurrent = Magnet_SampleOrientationUp Then
        MsgFoot = MsgFoot & vbCrLf & "Orientation: Up"
        LogLine = LogLine & " (Up)"
    ElseIf SampleOrientationCurrent = Magnet_SampleOrientationDown Then
        MsgFoot = MsgFoot & vbCrLf & "Orientation: Down"
        LogLine = LogLine & " (Dn)"
    End If

    LogLine = LogLine & ": " & Body

    MsgFoot = MsgFoot & vbCrLf & "Time: " & Date & " " & time & vbCrLf & "Code: " & CodeLevel & vbCrLf

    If LenB(sampleid) > 0 And LogMessages Then
        MainChanger.ChangerSample(SampQueue(1).hole).WriteLogFile (LogLine)
    End If

    If (LenB(LoginEmail) <> 0 And LenB(MailSMTPHost) <> 0) Then
        bAuthLogin = True
        bPopLogin = False
        bHtml = False
        MyEncodeType = MIME_ENCODE
        etPriority = NORMAL_PRIORITY
        bReceipt = False

        txtServer.text = MailSMTPHost
        txtFrom.text = MailFrom
        txtFromName.text = MailFromName
        txtUserName.text = MailFrom
        txtPassword.text = MailFromPassword

        If VerboseStatusUpdate Then
            txtToName.text = "Status Monitor Update"
            txtTo.text = MailStatusMonitor
        Else
            txtTo.text = LoginEmail

            txtToName.text = LoginName
            If LenB(MailStatusMonitor) > 0 Then
                If LenB(MailCCList) > 0 Then
                    txtCc.text = MailCCList & ";" & MailStatusMonitor
                Else
                    txtCc.text = MailStatusMonitor
                End If
            Else
                txtCc.text = MailCCList
            End If
        End If

        txtSubject.text = Subject
        txtMsg.text = MsgHead + Body + MsgFoot
        cmdSend_Click
    End If
End Sub

Private Sub optEncodeType_Click(Index As Integer)

    If optEncodeType(0).Value = True Then
        MyEncodeType = MIME_ENCODE
        cboPriority.Enabled = True
        ckHtml.Enabled = True
        ckReceipt.Enabled = True
        ckLogin.Enabled = True
    Else
        MyEncodeType = UU_ENCODE
        ckHtml.Value = vbUnchecked
        ckReceipt.Value = vbUnchecked
        ckLogin.Value = vbUnchecked
        cboPriority.Enabled = False
        ckHtml.Enabled = False
        ckReceipt.Enabled = False
        ckLogin.Enabled = False
    End If

End Sub

' *****************************************************************************
' The following four Subs capture the Events fired by the vbSendMail component
' *****************************************************************************

Private Sub poSendMail_Progress(lPercentCompete As Long)

    ' vbSendMail 'Progress Event'
    lblProgress = lPercentCompete & "% complete"

End Sub

Private Sub poSendMail_SendFailed(Explanation As String)

    ' vbSendMail 'SendFailed Event
    'MsgBox ("Your attempt to send mail failed for the following reason(s): " & vbCrLf & Explanation)
    lblProgress = vbNullString
    lstStatus.AddItem Explanation
    lstStatus.ListIndex = lstStatus.ListCount - 1
    lstStatus.ListIndex = -1

    Screen.MousePointer = vbDefault
    cmdSend.Enabled = True

End Sub

Private Sub poSendMail_SendSuccesful()

    ' vbSendMail 'SendSuccesful Event'
    'MsgBox "Send Successful!"
    lblProgress = vbNullString

End Sub

Private Sub poSendMail_Status(Status As String)

    ' vbSendMail 'Status Event'
    lstStatus.AddItem Status
    lstStatus.ListIndex = lstStatus.ListCount - 1
    lstStatus.ListIndex = -1

End Sub

Private Sub RetrieveSavedValues()

    ' *****************************************************************************
    ' Retrieve saved values by reading the components 'Persistent' properties
    ' *****************************************************************************
    poSendMail.PersistentSettings = True
    txtServer.text = poSendMail.SMTPHost
    txtFrom.text = poSendMail.From
    txtFromName.text = poSendMail.FromDisplayName
    txtUserName = poSendMail.UserName
    txtPassword = poSendMail.Password
    optEncodeType(poSendMail.EncodeType).Value = True
    If poSendMail.UseAuthentication Then ckLogin = vbChecked Else ckLogin = vbUnchecked

End Sub

Private Sub SendSSLEncryptedEmailToSMTPServer()

    Dim objMsg As Object
    Dim objConfig As Object
    Dim varFields As Variant

    On Error GoTo SendSSLEncryptedEmailToSMTPServer_Error

    poSendMail_Status "Provision SSL Protocol over SMTP"

    Set objMsg = CreateObject("CDO.Message")
    Set objConfig = CreateObject("CDO.Configuration")

    objConfig.Load -1
    Set varFields = objConfig.Fields

    Dim use_ssl As Boolean
    use_ssl = Me.chUseSSLEncryption.Value = Checked

    Dim smtp_authenticate As Long

    smtp_authenticate = CLng(modConfig.MailSMTPAuthenticate)

    If use_ssl And Me.ckLogin.Value = Checked Then

        smtp_authenticate = 1

    End If

    If smtp_authenticate = 1 Then

        If Me.txtPassword.text = "" Then Me.txtPassword.text = modConfig.MailSMTPPassword
        If Me.txtUserName.text = "" Then Me.txtUserName.text = modConfig.MailSMTPUsername

    End If

    With varFields

        .Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = use_ssl
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = smtp_authenticate
        .Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = Trim(Me.txtUserName.text)
        .Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = Me.txtPassword.text
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = poSendMail.SMTPHost
        .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = CLng(modConfig.MailSMTPSendUsing)
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = poSendMail.SMTPPort
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = poSendMail.ConnectTimeout
        .Update

    End With

    With objMsg

        Set .Configuration = objConfig

        poSendMail_Status "Configure SMTP message"

        Dim to_address As String
        Dim cc_address As String
        Dim bcc_address As String
        Dim from_address As String

        If Len(Trim(poSendMail.RecipientDisplayName)) > 0 Then

            to_address = """" & poSendMail.RecipientDisplayName & """" & _
                         " <" & poSendMail.Recipient & ">"
        Else

            to_address = Replace(poSendMail.Recipient, ";", ",")

        End If

        If Len(Trim(poSendMail.CcDisplayName)) > 0 Then

            cc_address = """" & poSendMail.CcDisplayName & """" & _
                         " <" & poSendMail.CcRecipient & ">"
        Else

            cc_address = Replace(poSendMail.CcRecipient, ";", ",")

        End If

        bcc_address = Replace(poSendMail.BccRecipient, ";", ",")

        If Len(Trim(poSendMail.FromDisplayName)) > 0 Then

            from_address = """" & poSendMail.FromDisplayName & """" & _
                           " <" & poSendMail.From & ">"
        Else

            from_address = Replace(poSendMail.From, ";", ",")

        End If

        .To = to_address
        .CC = cc_address
        .BCC = bcc_address
        .From = from_address

        If poSendMail.EncodeType = MIME_ENCODE Then

            .MimeFormatted = True

        Else

            .MimeFormatted = False

        End If

        If poSendMail.AsHTML Then

            .HTMLBody = poSendMail.message

        Else

            .TextBody = poSendMail.message

        End If

        .Subject = poSendMail.Subject

        If poSendMail.Receipt Then

            .MDNRequested = True

        End If


        poSendMail_Status "Begin Send to SMTP Host: " & poSendMail.SMTPHost
        .Send



    End With

    Set objMsg = Nothing
    Set objConfig = Nothing
    Set varFields = Nothing

    poSendMail_SendSuccesful

    On Error GoTo 0
    Exit Sub

SendSSLEncryptedEmailToSMTPServer_Error:

    poSendMail_SendFailed "Error sending message to remote SMTP server."
    poSendMail_SendFailed Err.number & " - " & Err.Description

    Set objMsg = Nothing
    Set objConfig = Nothing
    Set varFields = Nothing

End Sub

Public Sub StatusMonitorUpdate(CodeLevel As String)
    MailNotification "Status Code " & CodeLevel, MailFromName & " is now operating at code " & CodeLevel & vbCrLf, CodeLevel, True
End Sub

