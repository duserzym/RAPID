VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmOptions 
   Caption         =   "Options dialog"
   ClientHeight    =   6135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6960
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   6960
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frameOptions 
      BorderStyle     =   0  'None
      Caption         =   "Email settings"
      Height          =   4815
      Index           =   3
      Left            =   240
      TabIndex        =   27
      Top             =   480
      Width           =   6492
      Begin VB.CheckBox chkUseSslEncryption 
         Caption         =   "Use SSL Encryption"
         Height          =   255
         Left            =   3600
         TabIndex        =   91
         ToolTipText     =   "Use Login Authorization When Connecting to a Host"
         Top             =   600
         Width           =   2235
      End
      Begin VB.CommandButton cmdTestEmailSettings 
         Caption         =   "Test Email Settings"
         Height          =   375
         Left            =   2040
         TabIndex        =   90
         Top             =   4440
         Width           =   2535
      End
      Begin VB.TextBox txtMailPassword 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   2040
         PasswordChar    =   "*"
         TabIndex        =   89
         Top             =   1800
         Width           =   2250
      End
      Begin VB.TextBox txtMailUsername 
         Height          =   285
         Left            =   2040
         TabIndex        =   86
         Top             =   1440
         Width           =   2250
      End
      Begin VB.CheckBox ckLogin 
         Caption         =   "Requires Login"
         Height          =   255
         Left            =   2040
         TabIndex        =   85
         ToolTipText     =   "Use Login Authorization When Connecting to a Host"
         Top             =   1080
         Width           =   1515
      End
      Begin VB.TextBox txtSmtpPort 
         Height          =   285
         Left            =   2040
         TabIndex        =   84
         Top             =   600
         Width           =   1080
      End
      Begin VB.TextBox txtMailStatusMonitor 
         Height          =   288
         Left            =   2040
         TabIndex        =   39
         Top             =   3960
         Width           =   4212
      End
      Begin VB.TextBox txtMailCCList 
         Height          =   288
         Left            =   2040
         TabIndex        =   34
         Top             =   3480
         Width           =   4212
      End
      Begin VB.TextBox txtMailSMTPHost 
         Height          =   288
         Left            =   2040
         TabIndex        =   30
         Top             =   120
         Width           =   4212
      End
      Begin VB.TextBox txtMailFromAddress 
         Height          =   288
         Left            =   2040
         TabIndex        =   28
         Top             =   2880
         Width           =   4212
      End
      Begin VB.TextBox txtMailFromName 
         Height          =   288
         Left            =   2040
         TabIndex        =   29
         Top             =   2400
         Width           =   4212
      End
      Begin VB.Label Label37 
         Caption         =   "Password:"
         Height          =   255
         Left            =   240
         TabIndex        =   88
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label36 
         Caption         =   "Username:"
         Height          =   255
         Left            =   240
         TabIndex        =   87
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label lblSMTPServerPort 
         Caption         =   "SMTP Port:"
         Height          =   255
         Left            =   240
         TabIndex        =   83
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label14 
         Caption         =   "Status Monitor Email:"
         Height          =   375
         Left            =   240
         TabIndex        =   40
         Top             =   3960
         Width           =   1575
      End
      Begin VB.Label Label19 
         Caption         =   "CC List:"
         Height          =   375
         Left            =   240
         TabIndex        =   35
         Top             =   3480
         Width           =   1575
      End
      Begin VB.Label Label18 
         Caption         =   "From address:"
         Height          =   375
         Left            =   240
         TabIndex        =   33
         Top             =   2880
         Width           =   1575
      End
      Begin VB.Label Label17 
         Caption         =   "From name:"
         Height          =   255
         Left            =   240
         TabIndex        =   32
         Top             =   2400
         Width           =   1455
      End
      Begin VB.Label Label16 
         Caption         =   "SMTP Host:"
         Height          =   255
         Left            =   240
         TabIndex        =   31
         Top             =   120
         Width           =   1575
      End
   End
   Begin VB.Frame frameOptions 
      BorderStyle     =   0  'None
      Caption         =   "General program options"
      Height          =   3372
      Index           =   0
      Left            =   240
      TabIndex        =   4
      Top             =   720
      Width           =   6492
      Begin VB.TextBox txtINIFile 
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   82
         Top             =   120
         Width           =   5415
      End
      Begin MSComDlg.CommonDialog dialogFileBrowser 
         Left            =   240
         Top             =   2760
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton cmdHelpURLBrowse 
         Caption         =   ". . ."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5760
         TabIndex        =   69
         Top             =   2040
         Width           =   615
      End
      Begin VB.CommandButton cmdDefaultPathBrowse 
         Caption         =   ". . ."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5760
         TabIndex        =   68
         Top             =   1560
         Width           =   615
      End
      Begin VB.CommandButton cmdUsageFileBrowse 
         Caption         =   ". . ."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5760
         TabIndex        =   67
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox txtUserEmail 
         Height          =   285
         Left            =   4680
         TabIndex        =   66
         Top             =   3000
         Width           =   1455
      End
      Begin VB.TextBox txtUserName 
         Height          =   285
         Left            =   4680
         TabIndex        =   64
         Top             =   2640
         Width           =   1455
      End
      Begin VB.DriveListBox driveDefaultBackupDrive 
         Height          =   315
         Left            =   2040
         TabIndex        =   7
         Top             =   480
         Width           =   2012
      End
      Begin VB.TextBox txtUsageFile 
         Height          =   285
         Left            =   2040
         TabIndex        =   6
         Top             =   960
         Width           =   3495
      End
      Begin VB.TextBox txtDefaultPath 
         Height          =   285
         Left            =   2040
         TabIndex        =   5
         Top             =   1560
         Width           =   3495
      End
      Begin VB.TextBox txtHelpURLRoot 
         Height          =   285
         Left            =   2040
         TabIndex        =   44
         Top             =   2040
         Width           =   3495
      End
      Begin VB.TextBox txtRemeasureCSDThreshold 
         Height          =   285
         Left            =   2340
         TabIndex        =   46
         Top             =   2640
         Width           =   855
      End
      Begin VB.Label Label29 
         Caption         =   "Ini File:"
         Height          =   255
         Left            =   240
         TabIndex        =   81
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label26 
         Caption         =   "Email:"
         Height          =   255
         Left            =   4200
         TabIndex        =   65
         Top             =   3000
         Width           =   1215
      End
      Begin VB.Label Label24 
         Caption         =   "User Name:"
         Height          =   255
         Left            =   3720
         TabIndex        =   63
         Top             =   2640
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Default Backup Drive:"
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   480
         Width           =   2240
      End
      Begin VB.Label Label2 
         Caption         =   "Usage file:"
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   960
         Width           =   2250
      End
      Begin VB.Label Label3 
         Caption         =   "Default Paleomag Data Folder path:"
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   1440
         Width           =   1635
      End
      Begin VB.Label Label13 
         Caption         =   "Help URL Root:"
         Height          =   375
         Left            =   240
         TabIndex        =   45
         Top             =   2040
         Width           =   2235
      End
      Begin VB.Label Label20 
         Caption         =   "Remeasure CSD Threshold:"
         Height          =   375
         Left            =   120
         TabIndex        =   47
         Top             =   2640
         Width           =   2235
      End
   End
   Begin VB.Frame frameOptions 
      BorderStyle     =   0  'None
      Caption         =   "Advanced settings"
      Height          =   3372
      Index           =   4
      Left            =   240
      TabIndex        =   36
      Top             =   720
      Width           =   6492
      Begin VB.CheckBox chkLogMessages 
         Caption         =   "Log messages"
         Height          =   252
         Left            =   480
         TabIndex        =   48
         Top             =   1560
         Width           =   3252
      End
      Begin VB.CheckBox chkDumpRawDataStats 
         Caption         =   "Dump raw data and stats"
         Height          =   252
         Left            =   480
         TabIndex        =   41
         Top             =   1200
         Width           =   3252
      End
      Begin VB.CheckBox checkNOCOMM_MODE 
         Caption         =   "Disable communications"
         Height          =   252
         Left            =   480
         TabIndex        =   38
         Top             =   480
         Width           =   3252
      End
      Begin VB.CheckBox checkDEBUG_MODE 
         Caption         =   "Debug mode"
         Height          =   252
         Left            =   480
         TabIndex        =   37
         Top             =   840
         Width           =   3252
      End
   End
   Begin VB.Frame frameOptions 
      BorderStyle     =   0  'None
      Caption         =   "General program options"
      Height          =   3372
      Index           =   1
      Left            =   240
      TabIndex        =   70
      Top             =   720
      Width           =   6492
      Begin VB.TextBox txtLogoFile 
         Height          =   285
         Left            =   2040
         TabIndex        =   74
         Top             =   1440
         Width           =   3495
      End
      Begin VB.TextBox txtIconFile 
         Height          =   285
         Left            =   2040
         TabIndex        =   73
         Top             =   960
         Width           =   3495
      End
      Begin VB.CommandButton cmdIconFileBrowser 
         Caption         =   ". . ."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5760
         TabIndex        =   72
         Top             =   960
         Width           =   615
      End
      Begin VB.CommandButton cmdLogoFileBrowse 
         Caption         =   ". . ."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5760
         TabIndex        =   71
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label Label33 
         Caption         =   "Startup Logo File:"
         Height          =   375
         Left            =   240
         TabIndex        =   76
         Top             =   1440
         Width           =   2240
      End
      Begin VB.Label Label30 
         Caption         =   "Form Icon File:"
         Height          =   372
         Left            =   240
         TabIndex        =   75
         Top             =   960
         Width           =   2244
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Save to .INI File"
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      Top             =   5640
      Width           =   1935
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3360
      TabIndex        =   2
      Top             =   5640
      Width           =   1095
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Save to Current Session, only"
      Height          =   375
      Left            =   4560
      TabIndex        =   1
      Top             =   5640
      Width           =   2295
   End
   Begin VB.Frame frameOptions 
      BorderStyle     =   0  'None
      Caption         =   "Communication ports"
      Height          =   3372
      Index           =   2
      Left            =   240
      TabIndex        =   14
      Top             =   600
      Width           =   6492
      Begin VB.TextBox txtCOMPortChangerY 
         Height          =   288
         Left            =   2040
         TabIndex        =   79
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox txtCOMPortSusceptibility 
         Height          =   288
         Left            =   2040
         TabIndex        =   42
         Top             =   2400
         Width           =   612
      End
      Begin VB.TextBox txtCOMPortAF 
         Height          =   288
         Left            =   5400
         TabIndex        =   23
         Top             =   1440
         Width           =   612
      End
      Begin VB.TextBox txtComPortVacuum 
         Height          =   288
         Left            =   5400
         TabIndex        =   22
         Top             =   960
         Width           =   612
      End
      Begin VB.TextBox txtCOMPortSquids 
         Height          =   288
         Left            =   5400
         TabIndex        =   21
         Top             =   480
         Width           =   612
      End
      Begin VB.TextBox txtCOMPortUpDown 
         Height          =   288
         Left            =   2040
         TabIndex        =   20
         Top             =   1920
         Width           =   612
      End
      Begin VB.TextBox txtCOMPortTurning 
         Height          =   288
         Left            =   2040
         TabIndex        =   19
         Top             =   1440
         Width           =   612
      End
      Begin VB.TextBox txtCOMPortChanger 
         Height          =   288
         Left            =   2040
         TabIndex        =   18
         Top             =   480
         Width           =   612
      End
      Begin VB.Label Label28 
         Caption         =   "Y Motor:"
         Height          =   255
         Left            =   240
         TabIndex        =   80
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label15 
         Caption         =   "Susceptibility:"
         Height          =   375
         Left            =   240
         TabIndex        =   43
         Top             =   2400
         Width           =   1575
      End
      Begin VB.Label Label12 
         Caption         =   "SQUIDs:"
         Height          =   255
         Left            =   3480
         TabIndex        =   26
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label11 
         Caption         =   "Vacuum:"
         Height          =   252
         Left            =   3480
         TabIndex        =   25
         Top             =   960
         Width           =   1452
      End
      Begin VB.Label Label10 
         Caption         =   "AF:"
         Height          =   255
         Left            =   3480
         TabIndex        =   24
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "X motor:"
         Height          =   252
         Left            =   240
         TabIndex        =   17
         Top             =   480
         Width           =   1572
      End
      Begin VB.Label Label5 
         Caption         =   "Turning motor:"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Up/Down motor:"
         Height          =   375
         Left            =   240
         TabIndex        =   15
         Top             =   1920
         Width           =   1575
      End
   End
   Begin VB.Frame frameOptions 
      BorderStyle     =   0  'None
      Caption         =   "Jump settings"
      Height          =   3855
      Index           =   5
      Left            =   240
      TabIndex        =   55
      Top             =   480
      Width           =   6492
      Begin VB.TextBox txtNbHolderTry 
         Height          =   285
         Left            =   5880
         TabIndex        =   78
         Top             =   3240
         Width           =   512
      End
      Begin VB.TextBox txtJumpThreshold 
         Height          =   285
         Left            =   5865
         TabIndex        =   59
         Top             =   795
         Width           =   512
      End
      Begin VB.TextBox txtJumpSensitivity 
         Height          =   285
         Left            =   4485
         TabIndex        =   51
         Top             =   1680
         Width           =   512
      End
      Begin VB.TextBox txtStrongMom 
         Height          =   285
         Left            =   3120
         TabIndex        =   60
         Top             =   1320
         Width           =   512
      End
      Begin VB.TextBox txtIntermMom 
         Height          =   285
         Left            =   4320
         TabIndex        =   61
         Top             =   1320
         Width           =   512
      End
      Begin VB.TextBox txtMomMinForRedo 
         Height          =   285
         Left            =   5880
         TabIndex        =   52
         Top             =   2205
         Width           =   512
      End
      Begin VB.TextBox txtNbTry 
         Height          =   285
         Left            =   5880
         TabIndex        =   54
         Top             =   2640
         Width           =   512
      End
      Begin VB.Label Label27 
         Caption         =   $"frmOptions.frx":0000
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   77
         Top             =   3000
         Width           =   5640
      End
      Begin VB.Label Label25 
         Caption         =   "Jumps during zero measurements:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   62
         Top             =   480
         Width           =   3015
      End
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         Caption         =   "SQUID jump protection parameters"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   56
         Top             =   120
         Width           =   6495
      End
      Begin VB.Label Label32 
         Caption         =   "Jump threshold (x10-5 emu) (maximum difference between the zero measurements of each SQUID):"
         Height          =   375
         Left            =   480
         TabIndex        =   57
         Top             =   795
         Width           =   5640
      End
      Begin VB.Label Label22 
         Caption         =   "Jump threshold test applies between              emu and             emu."
         Height          =   255
         Left            =   480
         TabIndex        =   49
         Top             =   1365
         Width           =   4920
      End
      Begin VB.Label Label34 
         Caption         =   "For weaker magnetizations, the jump threshold becomes             x moment"
         Height          =   375
         Left            =   480
         TabIndex        =   58
         Top             =   1755
         Width           =   5520
      End
      Begin VB.Label Label21 
         Caption         =   "Minimum moment for applying the CSD and zero jump criteria (emu):"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   50
         Top             =   2205
         Width           =   5760
      End
      Begin VB.Label Label23 
         Caption         =   "# of tries before accepting last measurement (0 = no mail):"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   53
         Top             =   2640
         Width           =   5880
      End
   End
   Begin ComctlLib.TabStrip tbsOptions 
      Height          =   5415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   9551
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   6
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&General"
            Key             =   "General"
            Object.Tag             =   ""
            Object.ToolTipText     =   "General program options"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Image &Files"
            Key             =   "Image Files"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Comm ports"
            Key             =   "CommPorts"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Communication ports"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Email"
            Key             =   "Email"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Email settings for notifications"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Advanced"
            Key             =   "Advanced"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Advanced settings"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Jumps"
            Key             =   "Jumps"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Jump settings"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Label Label7 
      Caption         =   "Default path:"
      Height          =   372
      Left            =   3600
      TabIndex        =   11
      Top             =   2520
      Width           =   1572
   End
   Begin VB.Label Label9 
      Caption         =   "Default Backup Drive:"
      Height          =   252
      Left            =   3600
      TabIndex        =   13
      Top             =   1560
      Width           =   1572
   End
   Begin VB.Label Label8 
      Caption         =   "Usage file:"
      Height          =   252
      Left            =   3600
      TabIndex        =   12
      Top             =   2040
      Width           =   1452
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdApply_Click()
    LoginName = txtUserName ' (February 2010 L Carporzen) Allow to change Name and email while running
    LoginEmail = txtUserEmail
    Prog_DefaultBackup = Left(driveDefaultBackupDrive.Drive, 1)
    
    
'-------------------------------------------------------------------------------------------------------'
    '(September 2010, I Hilburn)
    'Added in code to disconnect and reconnect the SQUID com port so that the new com assignment
    'can be made (prior setup - if the user made changes to the comm port setting, the code would have
    'to be restarted before the new setting came into effect).
    
    If Int(val(Me.txtCOMPortSquids)) <> COMPortSquids Then
    
        COMPortSquids = Int(val(txtCOMPortSquids))
        
        'Disconnect the SQUID comm port
        frmSQUID.Disconnect
        
        'Reassign and reconnect the SQUID comm port
        Load frmSQUID
    
    End If
    
'-------------------------------------------------------------------------------------------------------'
    
    'AF comm port is reopened and closed during each comm session with
    'the 2G box, so don't need to disconnect / reconnect here
    COMPortAf = Int(val(txtCOMPortAF))
    
    
'-------------------------------------------------------------------------------------------------------'
    '(September 2010, I Hilburn)
    'Added in code to disconnect and reconnect the motor com ports so that the new com assignments
    'can be made (prior setup - if the user made changes to the comm port settings, the code would have
    'to be restarted before the new settings came into effect).
    
    If COMPortUpDown <> Int(val(txtCOMPortUpDown)) Then
        
        COMPortUpDown = Int(val(txtCOMPortUpDown))
        
        'Disconnect the Up-Down motor com port
        frmDCMotors.MotorCommDisconnect MotorUpDown
        
        'Reassign and reconnect the up-down com port
        frmDCMotors.MotorCommConnect MotorUpDown
        
    End If
    
    If COMPortTurning <> Int(val(txtCOMPortTurning)) Then
        
        COMPortTurning = Int(val(txtCOMPortTurning))
        
        'Disconnect the Up-Down motor com port
        frmDCMotors.MotorCommDisconnect MotorUpDown
        
        'Reassign and reconnect the up-down com port
        frmDCMotors.MotorCommConnect MotorUpDown
        
    End If
    
    If COMPortChanger <> Int(val(txtCOMPortChanger)) Then
        
        COMPortChanger = Int(val(txtCOMPortChanger))
        
        'Disconnect the Up-Down motor com port
        frmDCMotors.MotorCommDisconnect MotorUpDown
        
        'Reassign and reconnect the up-down com port
        frmDCMotors.MotorCommConnect MotorUpDown
        
    End If
    
    If COMPortChangerY <> Int(val(txtCOMPortChangerY)) Then
        
        COMPortChangerY = Int(val(txtCOMPortChangerY))
        
        'Disconnect the Up-Down motor com port
        frmDCMotors.MotorCommDisconnect MotorUpDown
        
        'Reassign and reconnect the up-down com port
        frmDCMotors.MotorCommConnect MotorUpDown
        
    End If
    
'-------------------------------------------------------------------------------------------------------'
    
'-------------------------------------------------------------------------------------------------------'
    '(September 2010, I Hilburn)
    'Added in code to disconnect and reconnect the Vacuum com port so that the new com assignment
    'can be made (prior setup - if the user made changes to the comm port setting, the code would have
    'to be restarted before the new setting came into effect).
    
    If COMPortVacuum <> Int(val(txtComPortVacuum)) Then
    
        COMPortVacuum = Int(val(txtComPortVacuum))
    
        'Disconnect the Vacuum comm port
        frmVacuum.Disconnect
        
        'Reassign and reconnect the Vacuum comm port
        frmVacuum.Connect
    
    End If
        
'-------------------------------------------------------------------------------------------------------'

'-------------------------------------------------------------------------------------------------------'
    '(September 2010, I Hilburn)
    'Added in code to disconnect and reconnect the Susceptibility com port so that the new com assignment
    'can be made (prior setup - if the user made changes to the comm port setting, the code would have
    'to be restarted before the new setting came into effect).
    
    If COMPortSusceptibility <> Int(val(txtCOMPortSusceptibility)) Then
    
        COMPortSusceptibility = Int(val(txtCOMPortSusceptibility))
        
        'Disconnect the Susceptibility comm port
        frmSusceptibilityMeter.Disconnect
        
        'Reassign and reconnect the Susceptibility comm port
        frmSusceptibilityMeter.Connect

    End If

'-------------------------------------------------------------------------------------------------------'

    MailSMTPHost = txtMailSMTPHost
    MailSMTPPort = CLng(Me.txtSmtpPort)
    MailFrom = txtMailFromAddress
    MailFromName = txtMailFromName
    MailFromPassword = Me.txtMailPassword
    MailSMTPPassword = Me.txtMailPassword
    MailSMTPUsername = Me.txtMailUsername
    modConfig.MailSMTPAuthenticate = IIf(Me.ckLogin.Value = Checked, cdoBasic, cdoAnonymous)
    modConfig.MailUseSSLEncryption = Me.chkUseSslEncryption.Value = Checked
    
    MailCCList = txtMailCCList
    MailStatusMonitor = txtMailStatusMonitor
    Prog_UsageFile = txtUsageFile
    Prog_DefaultPath = txtDefaultPath
    Prog_HelpURLRoot = txtHelpURLRoot
    Prog_IcoFile = txtIconFile
    Prog_LogoFile = txtLogoFile
    
    NOCOMM_MODE = (checkNOCOMM_MODE.Value = 1)
    DEBUG_MODE = (checkDEBUG_MODE.Value = 1)
    DumpRawDataStats = (chkDumpRawDataStats.Value = 1)
    LogMessages = (chkLogMessages.Value = 1)
    RemeasureCSDThreshold = val(txtRemeasureCSDThreshold)
    ' New parameters for the jumps (April-May 2007 L Carporzen)
    JumpThreshold = val(txtJumpThreshold)
    StrongMom = val(txtStrongMom)
    IntermMom = val(txtIntermMom)
    MomMinForRedo = val(txtMomMinForRedo)
    JumpSensitivity = val(txtJumpSensitivity)
    NbTry = val(txtNbTry)
    NbHolderTry = val(Me.txtNbHolderTry.text)  '(Mar 2011 - I Hilburn)
                                               'Added in to resolve annoying issue of holder being
                                               'remeasured over and over and over again due to
                                               'both it's moment and CSD being too high.
                                               'This setting does NOT apply to SQUID jumps, for those
                                               'the NbTry setting is used, even for Holder measurements
    If DEBUG_MODE Then
        frmProgram.mnuViewDebug.Visible = True
    Else
        frmProgram.mnuViewDebug.Visible = False
    End If
    importSettings
    Me.refresh
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDefaultPathBrowse_Click()

    Dim FolderPath As String
    Dim StartFolder As String
    Dim fso As FileSystemObject

    'Use File dialog object to allow the user to browse for and (if necessary)
    'create a new directory.  Allow user to create a new directory if they
    'want to.
    
    'First see if the default file folder that's currently in
    'the folder path text box actually exists
    'if it doesn't, set the start folder to app.path (we know that folder must exist)
    Set fso = New FileSystemObject
    If Not fso.FolderExists(Prog_DefaultPath) Then
            
        StartFolder = App.path
        
    Else
    
        StartFolder = Prog_DefaultPath
        
    End If

    'Call public shell function in modSaveFile that'll setup
    'all of the API properties that we need and don't really want to know about
    FolderPath = modFileSave.OpenDir(StartFolder, _
                                     "Select the Default Paleomag Data Folder ", _
                                     Me)

    'Make sure that a "\" is at the end of the path
    If Right(FolderPath, 1) <> "\" Then

        FolderPath = FolderPath & "\"
        
    End If

    'Now check to make sure that this folder path exists
    If fso.FolderExists(FolderPath) Then
    
        'Set the text-box display to this folder path
        Me.txtDefaultPath = FolderPath
        
    End If

End Sub

Private Sub cmdHelpURLBrowse_Click()

    Dim InitialPath As String
    Dim NewPath As String
    Dim TempL As Long
    Dim UserResp As Long
    
    Dim fso As FileSystemObject
    
    'Search for "file:///" at the start of the Help Root File path
    TempL = InStr(1, Prog_HelpURLRoot, "file:///")
    
    If TempL > 0 Then
    
        InitialPath = Mid(Prog_HelpURLRoot, TempL + 1)
        
    Else
    
        'Check to see if the root help path is an http:// path
        TempL = InStr(1, Prog_HelpURLRoot, "http://")
        
        If TempL > 0 Then
        
            'Tell user that help root path is an online path
            'and that this browser can only search for
            'local files
            'Ask the user if they still want to replace
            'this online path with a local path
            UserResp = MsgBox("The current help root file has an online address. This browser " & _
                              vbNewLine & "can only find local root help files and cannot browse the internet." & _
                              vbNewLine & vbNewLine & "Would you like to replace the online root help file with " & _
                              " a local file?", _
                              vbYesNo, _
                              "Warning!")
                              
            If UserResp = vbNo Then
            
                Exit Sub
                
            End If
            
            InitialPath = vbNullString
            
        Else
    
            InitialPath = Prog_HelpURLRoot
            
        End If
        
    End If
        
    'Check to see if the resulting initial help path exists
    If Not fso.FolderExists(InitialPath) Then
    
        'Get the folder one up from the application path
        InitialPath = Mid(App.path, _
                          InStrRev(App.path, _
                                   "\", _
                                   InStrRev(App.path, _
                                            "\") - 1))
        
    End If
                
    'Open a browser dialog
    NewPath = modFileSave.OpenDir(InitialPath, _
                                  "Browse for Root Help File.", _
                                  Me)
                                  
    'Check to see if the new path exists
    If fso.FolderExists(NewPath) Then
    
        'Replace the old path in the form textbox
        Me.txtHelpURLRoot = "file:///" & NewPath
    
    End If

End Sub

Private Sub cmdIconFileBrowser_Click()

    Dim FilePath As String
    
    FilePath = GenericFileOpenBrowser(Me.dialogFileBrowser, _
                                      Prog_IcoFile, _
                                      "Browse for Icon image file...", _
                                      "Icons (*.ico)|*.ico", _
                                      cdlOFNFileMustExist)
                                      
                                          
    'Check to see if the returned file exists
    If FileExists(FilePath) Then
    
        Me.txtIconFile = FilePath
        
    End If

End Sub

Private Sub cmdLogoFileBrowse_Click()

    Dim FilePath As String
    
    FilePath = GenericFileOpenBrowser( _
                    Me.dialogFileBrowser, _
                    Prog_LogoFile, _
                    "Browse for Logo image file...", _
                    "Images (*.bmp;*.ico;*.wmf;*.emf;*.gif;*.jpeg)|*.bmp;*.ico;*.wmf;*.emf;*.gif;*.jpeg", _
                    cdlOFNFileMustExist)
                                                                                
    'Check to see if the returned file exists
    If FileExists(FilePath) Then
    
        Me.txtLogoFile = FilePath
        
    End If

End Sub

Private Sub cmdOK_Click()
    cmdApply_Click
    Config_writeSettingstoINI
    Unload Me
End Sub

Private Sub cmdTestEmailSettings_Click()

    With frmSendMail
    
        .txtServer.text = Me.txtMailSMTPHost.text
        .txtSmtpPort.text = Me.txtSmtpPort.text
        .chUseSSLEncryption.Value = Me.chkUseSslEncryption.Value
        .ckLogin.Value = Me.ckLogin.Value
        .txtUserName.text = Me.txtMailUsername.text
        .txtPassword.text = Me.txtMailPassword.text
        .txtTo.text = Me.txtUserEmail.text
        .txtToName.text = Me.txtUserName.text
        .txtCc.text = Me.txtMailCCList
        .txtBcc.text = Me.txtMailStatusMonitor
        .txtFrom = Me.txtMailFromAddress
        .txtFromName = Me.txtMailFromName
        .txtMsg = "Testing Email settings for Paleomag program." & vbNewLine & _
                  vbNewLine & "SMTP Server:" & .txtServer.text & _
                  vbNewLine & "SMTP Port:" & .txtSmtpPort.text & _
                  vbNewLine & "Use SSL Encryption: " & CStr(IIf(.chUseSSLEncryption.Value = Checked, "Yes", "No"))
                  
        .txtSubject = "Test Paleomag Email Settings"
        .chPlainText.Value = Checked
        .ckHtml.Value = Unchecked
        .lstStatus.Clear
        
    End With
    
    frmSendMail.ZOrder
    frmSendMail.Show

End Sub

Private Sub cmdUsageFileBrowse_Click()

   
    Dim FilePath As String
    
    FilePath = GenericFileOpenBrowser(Me.dialogFileBrowser, _
                                      Prog_UsageFile, _
                                      "Browse for Paleomag Usage File...", _
                                      "(*.dat)|*.dat", _
                                      cdlOFNCreatePrompt)
                                      
    'Check to see if the returned file exists
    If FileExists(FilePath) Then
    
        'Write in the new file path in the Usage File textbox
        Me.txtUsageFile = FilePath

    End If
        
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    'handle ctrl+tab to move to the next tab
    If Shift = vbCtrlMask And KeyCode = vbKeyTab Then
        i = tbsOptions.SelectedItem.Index
        If i = tbsOptions.Tabs.Count Then
            'last tab so we need to wrap to tab 1
            Set tbsOptions.SelectedItem = tbsOptions.Tabs(1)
        Else
            'increment the tab
            Set tbsOptions.SelectedItem = tbsOptions.Tabs(i + 1)
        End If
    End If
End Sub

Private Sub Form_Load()

    'Set Form width and height
    Me.Height = 6705
    Me.Width = 7200

    '(March 11, 2011 - I Hilburn)
    'Update all the form's icons with the program Ico file
    SetFormIcon Me

    'center the form
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
    importSettings
    selectTab 1
    
End Sub

Private Sub Form_Resize()
    Me.Height = 6705
    Me.Width = 7200
End Sub

Private Sub importSettings()
    Dim i As Integer
    On Error GoTo carryforth
    For i = 0 To driveDefaultBackupDrive.ListCount
        If UCase(Left(driveDefaultBackupDrive.List(i), 1)) = UCase(Prog_DefaultBackup) Then
            driveDefaultBackupDrive.ListIndex = i
        End If
    Next i
carryforth:
    On Error GoTo 0
    txtCOMPortSquids = COMPortSquids
    txtCOMPortAF = COMPortAf
    txtCOMPortUpDown = COMPortUpDown
    txtCOMPortTurning = COMPortTurning
    txtCOMPortChanger = COMPortChanger
    txtCOMPortChangerY = COMPortChangerY
    txtComPortVacuum = COMPortVacuum
    txtCOMPortSusceptibility = COMPortSusceptibility
    txtMailSMTPHost = MailSMTPHost
    txtSmtpPort = MailSMTPPort
    txtMailFromAddress = MailFrom
    Me.txtMailPassword = MailSMTPPassword
    Me.txtMailUsername = MailSMTPUsername
    If modConfig.MailUseSSLEncryption Then
        Me.chkUseSslEncryption.Value = Checked
    Else
        Me.chkUseSslEncryption.Value = Unchecked
    End If
    If modConfig.MailSMTPAuthenticate = cdoBasic Then
    
        Me.ckLogin.Value = Checked
    
    Else
    
        Me.ckLogin.Value = Unchecked
        
    End If
        
    txtMailFromName = MailFromName
    txtMailCCList = MailCCList
    txtMailStatusMonitor = MailStatusMonitor
    txtINIFile = Prog_INIFile
    txtUsageFile = Prog_UsageFile
    txtDefaultPath = Prog_DefaultPath
    txtHelpURLRoot = Prog_HelpURLRoot
    txtIconFile = Prog_IcoFile
    txtLogoFile = Prog_LogoFile
    txtRemeasureCSDThreshold = str$(RemeasureCSDThreshold)
    ' New parameters for the jumps (April-May 2007 L Carporzen)
    txtJumpThreshold = str$(JumpThreshold)
    txtStrongMom = Format$(StrongMom, "0E+00")
    txtIntermMom = Format$(IntermMom, "0E+00")
    txtMomMinForRedo = Format$(MomMinForRedo, "0E+00")
    txtJumpSensitivity = str$(JumpSensitivity)
    txtNbTry = str$(NbTry)
    txtNbHolderTry = str$(NbHolderTry)
    checkNOCOMM_MODE.Value = Abs(NOCOMM_MODE)
    checkDEBUG_MODE.Value = Abs(DEBUG_MODE)
    chkDumpRawDataStats.Value = Abs(DumpRawDataStats)
    chkLogMessages.Value = Abs(LogMessages)
End Sub

Private Sub selectTab(tabtoselect As Integer)
    Dim i As Integer
    'show and enable the selected tab's controls
    'and hide and disable all others
    For i = 0 To tbsOptions.Tabs.Count - 1
        If i = tabtoselect - 1 Then
            'frameOptions(i).Left = 210
            frameOptions(i).Visible = True
            frameOptions(i).Enabled = True
            frameOptions(i).ZOrder 0
        Else
            'frameOptions(i).Left = -20000
            frameOptions(i).Visible = False
            frameOptions(i).Enabled = False
        End If
    Next
End Sub

Private Sub tbsOptions_Click()
    selectTab tbsOptions.SelectedItem.Index
End Sub

